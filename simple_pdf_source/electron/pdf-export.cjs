'use strict'

const JSZip = require('jszip')

const TEXT_EXPORT_FORMATS = new Set(['docx', 'txt', 'md', 'html'])
const IMAGE_EXPORT_FORMATS = new Set(['png', 'jpeg', 'webp'])

function safeExportBaseName(value) {
  return String(value || 'Untitled')
    // Callers normally pass the already extensionless PDF name. Only remove an
    // actual PDF suffix here so meaningful dotted suffixes such as
    // "Quarterly.v2" are not mistaken for a file extension a second time.
    .replace(/\.pdf$/i, '')
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim() || 'Untitled'
}

function stripInvalidXmlCharacters(value) {
  // XML 1.0 permits tabs and line breaks, but not the remaining C0 controls or
  // the two BMP noncharacters. PDF text streams occasionally expose those
  // bytes; removing them here keeps DOCX and HTML outputs parseable.
  return String(value ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
}

function escapeXml(value) {
  return stripInvalidXmlCharacters(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function normalizeTextPages(input) {
  if (!Array.isArray(input) || input.length < 1 || input.length > 10_000) {
    throw new Error('Choose between 1 and 10,000 pages to export.')
  }
  let totalLength = 0
  return input.map((page, index) => {
    const pageNumber = Number(page?.pageNumber)
    const text = stripInvalidXmlCharacters(page?.text).replace(/\r\n?/g, '\n')
    if (!Number.isInteger(pageNumber) || pageNumber < 1 || pageNumber > 1_000_000) {
      throw new Error(`Page ${index + 1} has an invalid page number.`)
    }
    totalLength += text.length
    if (totalLength > 50_000_000) throw new Error('The extracted text is too large to export safely in one file.')
    return { pageNumber, text }
  })
}

function plainTextDocument(pages, title) {
  const heading = String(title || '').trim()
  const body = pages.map((page) => `Page ${page.pageNumber}\n${page.text.trimEnd()}`).join('\n\n\f\n\n')
  return `${heading ? `${heading}\n${'='.repeat(Math.min(heading.length, 72))}\n\n` : ''}${body}\n`
}

function markdownDocument(pages, title) {
  const heading = String(title || '').trim()
  const body = pages.map((page) => `## Page ${page.pageNumber}\n\n${page.text.trimEnd()}`).join('\n\n---\n\n')
  return `${heading ? `# ${heading.replace(/^#+\s*/gm, '')}\n\n` : ''}${body}\n`
}

function htmlDocument(pages, title) {
  const documentTitle = escapeXml(String(title || 'Exported PDF'))
  const sections = pages.map((page) => {
    const lines = page.text.split('\n').map((line) => `<div>${escapeXml(line) || '&nbsp;'}</div>`).join('')
    return `<section class="page"><header>Page ${page.pageNumber}</header><main>${lines}</main></section>`
  }).join('\n')
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${documentTitle}</title><style>
body{margin:0;padding:32px;background:#eee;color:#18181b;font:15px/1.55 "Segoe UI",Arial,sans-serif}
.page{box-sizing:border-box;max-width:850px;min-height:1050px;margin:0 auto 24px;padding:64px 72px;background:#fff;box-shadow:0 2px 14px #0002;page-break-after:always}
.page:last-child{page-break-after:auto}.page>header{margin-bottom:30px;padding-bottom:10px;border-bottom:1px solid #ddd;color:#666;font-size:12px;text-transform:uppercase;letter-spacing:.08em}
.page main div{min-height:1.55em;white-space:pre-wrap;overflow-wrap:anywhere}@media print{body{padding:0;background:#fff}.page{max-width:none;min-height:0;margin:0;padding:18mm;box-shadow:none}}
</style></head><body>${sections}</body></html>`
}

function paragraphXml(text, extraRunProperties = '') {
  if (!text) return '<w:p/>'
  return `<w:p><w:r>${extraRunProperties}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
}

async function docxDocument(pages, title) {
  const zip = new JSZip()
  const body = []
  if (title) {
    body.push(paragraphXml(title, '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'))
    body.push('<w:p/>')
  }
  pages.forEach((page, pageIndex) => {
    if (pageIndex) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    body.push(paragraphXml(`Page ${page.pageNumber}`, '<w:rPr><w:b/><w:color w:val="666666"/></w:rPr>'))
    for (const line of page.text.split('\n')) body.push(paragraphXml(line))
  })
  body.push('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr>')

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`)
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join('')}</w:body></w:document>`)
  zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/></w:rPr></w:style></w:styles>`)
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${escapeXml(title || 'Exported PDF')}</dc:title><dc:creator>simple</dc:creator></cp:coreProperties>`)
  zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>simple</Application></Properties>`)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } })
}

async function buildTextExport(format, rawPages, title) {
  if (!TEXT_EXPORT_FORMATS.has(format)) throw new Error('This text export format is not supported.')
  const pages = normalizeTextPages(rawPages)
  if (format === 'docx') return docxDocument(pages, String(title || ''))
  const value = format === 'html'
    ? htmlDocument(pages, title)
    : format === 'md' ? markdownDocument(pages, title) : plainTextDocument(pages, title)
  return Buffer.from(value, 'utf8')
}

function imageExportFileName(baseName, pageNumber, pageCount, format) {
  if (!IMAGE_EXPORT_FORMATS.has(format)) throw new Error('This image export format is not supported.')
  const width = Math.max(3, String(Math.max(1, pageCount)).length)
  const extension = format === 'jpeg' ? 'jpg' : format
  return `${safeExportBaseName(baseName)} - page ${String(pageNumber).padStart(width, '0')}.${extension}`
}

module.exports = {
  IMAGE_EXPORT_FORMATS,
  TEXT_EXPORT_FORMATS,
  buildTextExport,
  imageExportFileName,
  normalizeTextPages,
  safeExportBaseName,
}
