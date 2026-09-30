const { basename } = require('node:path')
const JSZip = require('jszip')
const WordExtractor = require('word-extractor')
const { MAX_FILE_BYTES, validateDocxBytes } = require('./docx-files.cjs')
const { findOfficeConverter, convertOfficeBytes } = require('./office-converter.cjs')

const OLE_COMPOUND_FILE_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
const MINIMUM_OLE_FILE_BYTES = 512
const ZIP_ENTRY_DATE = new Date('1980-01-01T00:00:00.000Z')

function toBytes(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data)
  return Buffer.from(value || [])
}

function validateLegacyDocBytes(data) {
  const bytes = toBytes(data)
  if (bytes.byteLength < MINIMUM_OLE_FILE_BYTES) {
    throw new Error('This legacy .doc file is empty or incomplete.')
  }
  if (bytes.byteLength > MAX_FILE_BYTES) {
    throw new Error('This legacy .doc file exceeds the 256 MB safe-open limit.')
  }
  if (!bytes.subarray(0, OLE_COMPOUND_FILE_SIGNATURE.length).equals(OLE_COMPOUND_FILE_SIGNATURE)) {
    throw new Error('This is not a valid legacy .doc file. Choose a Word 97–2003 document, or save it as .docx and try again.')
  }
  return bytes
}

function normalizeLegacyText(value) {
  return String(value ?? '')
    .replace(/\r\n?/g, '\n')
    .replace(/\u0000/g, '')
}

function isXmlCharacter(codePoint) {
  return codePoint === 0x9
    || codePoint === 0xa
    || codePoint === 0xd
    || (codePoint >= 0x20 && codePoint <= 0xd7ff)
    || (codePoint >= 0xe000 && codePoint <= 0xfffd)
    || (codePoint >= 0x10000 && codePoint <= 0x10ffff)
}

function xmlText(value) {
  let safe = ''
  for (const character of String(value ?? '')) {
    const codePoint = character.codePointAt(0)
    if (isXmlCharacter(codePoint)) safe += character
  }
  return safe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function textRunsXml(line) {
  const pieces = line.split('\t')
  const runs = []
  for (let index = 0; index < pieces.length; index += 1) {
    if (index > 0) runs.push('<w:r><w:tab/></w:r>')
    if (pieces[index]) runs.push(`<w:r><w:t xml:space="preserve">${xmlText(pieces[index])}</w:t></w:r>`)
  }
  return runs.join('')
}

function textParagraphsXml(value, style = null) {
  const pages = normalizeLegacyText(value).split('\f')
  const paragraphs = []
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
    if (pageIndex > 0) paragraphs.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    const lines = pages[pageIndex].split('\n')
    for (const line of lines) {
      const paragraphProperties = style ? `<w:pPr><w:pStyle w:val="${xmlText(style)}"/></w:pPr>` : ''
      paragraphs.push(`<w:p>${paragraphProperties}${textRunsXml(line)}</w:p>`)
    }
  }
  return paragraphs.join('')
}

function safeExtractedPart(document, methodName, options, warnings) {
  if (!document || typeof document[methodName] !== 'function') return ''
  try {
    return normalizeLegacyText(document[methodName](options))
  } catch {
    warnings.push(`Simple Docs could not recover ${methodName.replace(/^get/, '').toLowerCase()} text from this file.`)
    return ''
  }
}

function readLegacyDocumentParts(document) {
  if (!document || typeof document.getBody !== 'function') {
    throw new Error('The legacy Word extractor returned an unreadable document.')
  }

  let body
  try {
    body = normalizeLegacyText(document.getBody({ filterUnicode: false }))
  } catch (error) {
    throw new Error('Simple Docs could not recover the document body from this legacy .doc file.', { cause: error })
  }

  const warnings = []
  const parts = {
    body,
    headers: safeExtractedPart(document, 'getHeaders', { filterUnicode: false, includeFooters: false }, warnings),
    footers: safeExtractedPart(document, 'getFooters', { filterUnicode: false }, warnings),
    footnotes: safeExtractedPart(document, 'getFootnotes', { filterUnicode: false }, warnings),
    endnotes: safeExtractedPart(document, 'getEndnotes', { filterUnicode: false }, warnings),
    annotations: safeExtractedPart(document, 'getAnnotations', { filterUnicode: false }, warnings),
    textboxes: safeExtractedPart(
      document,
      'getTextboxes',
      { filterUnicode: false, includeHeadersAndFooters: false, includeBody: true },
      warnings,
    ),
  }
  return { parts, warnings }
}

function visibleSupplementaryParts(parts) {
  const candidates = [
    ['Imported headers', parts.headers],
    ['Imported footers', parts.footers],
    ['Imported footnotes', parts.footnotes],
    ['Imported endnotes', parts.endnotes],
    ['Imported comments', parts.annotations],
    ['Imported text boxes', parts.textboxes],
  ]
  const seen = new Set([normalizeLegacyText(parts.body).trim()])
  return candidates.filter(([, value]) => {
    const normalized = normalizeLegacyText(value).trim()
    if (!normalized || seen.has(normalized)) return false
    seen.add(normalized)
    return true
  })
}

function documentXml(parts) {
  const bodyText = normalizeLegacyText(parts.body)
  const supplementary = visibleSupplementaryParts(parts)
  const content = []

  if (bodyText) content.push(textParagraphsXml(bodyText))
  else content.push('<w:p/>')

  if (supplementary.length) {
    if (bodyText) content.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    for (const [heading, value] of supplementary) {
      content.push(textParagraphsXml(heading, 'Heading1'))
      content.push(textParagraphsXml(value))
    }
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${content.join('\n    ')}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`
}

function packageParts(parts, title) {
  const safeTitle = xmlText(basename(String(title || 'Imported legacy document')).replace(/\.doc$/i, ''))
  return {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
    'word/document.xml': documentXml(parts),
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
    'word/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="320" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr>
  </w:style>
</w:styles>`,
    'docProps/core.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${safeTitle}</dc:title>
  <dc:creator>Simple Docs</dc:creator>
  <cp:lastModifiedBy>Simple Docs</cp:lastModifiedBy>
</cp:coreProperties>`,
    'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Simple Docs</Application>
</Properties>`,
  }
}

async function buildEditableDocx(parts, options = {}) {
  const zip = new JSZip()
  for (const [name, value] of Object.entries(packageParts(parts || {}, options.title))) {
    zip.file(name, value, { date: ZIP_ENTRY_DATE })
  }
  const bytes = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    platform: 'DOS',
  })
  return validateDocxBytes(bytes)
}

function actionableImportError(error) {
  const message = error instanceof Error ? error.message : String(error || '')
  if (/password|encrypted|encryption/i.test(message)) {
    return new Error('Password-protected legacy Word documents are not supported. Remove the password in Word, then try the .doc file again.', { cause: error })
  }
  return new Error('Simple Docs could not import this legacy .doc file. Open it in Word or LibreOffice, save it as .docx, and try again.', { cause: error })
}

async function convertLegacyDocToDocx(data, options = {}) {
  const inputBytes = toBytes(data)
  if (inputBytes[0] === 0x50 && inputBytes[1] === 0x4b) {
    const docxBytes = validateDocxBytes(inputBytes)
    return {
      data: docxBytes,
      warnings: ['This .doc filename contained a modern DOCX document. Simple Docs opened its original editable content and will save it with the correct .docx extension.'],
    }
  }

  const bytes = validateLegacyDocBytes(inputBytes)
  let conversionWarning = null
  if (!options.Extractor) {
    try {
      const executable = await findOfficeConverter()
      if (executable || options.convertOffice) {
        const converted = await (options.convertOffice || convertOfficeBytes)({ bytes, inputExtension: 'doc', outputExtension: 'docx', filter: 'Office Open XML Text' }, { executable })
        return { data: validateDocxBytes(converted), conversionMethod: 'layout', warnings: [] }
      }
    } catch (error) {
      conversionWarning = error instanceof Error ? error.message : 'The layout converter could not open this file.'
    }
  }
  const Extractor = options.Extractor || WordExtractor
  let extracted
  try {
    const extractor = new Extractor()
    extracted = await extractor.extract(bytes)
  } catch (error) {
    throw actionableImportError(error)
  }

  let result
  try {
    result = readLegacyDocumentParts(extracted)
  } catch (error) {
    throw actionableImportError(error)
  }
  const converted = await buildEditableDocx(result.parts, { title: options.title })
  const supplementaryPartCount = visibleSupplementaryParts(result.parts).length
  return {
    data: converted,
    conversionMethod: 'text',
    warnings: [
      'Text-only import: original formatting, images, tables, and page layout were not preserved. Install LibreOffice or open a DOCX copy to keep the document structure.',
      ...(conversionWarning ? [conversionWarning] : []),
      ...(supplementaryPartCount ? ['Headers, footers, notes, comments, or text boxes were appended as editable sections.'] : []),
      ...result.warnings,
    ],
  }
}

module.exports = {
  MINIMUM_OLE_FILE_BYTES,
  OLE_COMPOUND_FILE_SIGNATURE,
  buildEditableDocx,
  convertLegacyDocToDocx,
  normalizeLegacyText,
  readLegacyDocumentParts,
  validateLegacyDocBytes,
  visibleSupplementaryParts,
}
