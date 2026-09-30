const test = require('node:test')
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const {
  OLE_COMPOUND_FILE_SIGNATURE,
  buildEditableDocx,
  convertLegacyDocToDocx,
  validateLegacyDocBytes,
} = require('../electron/legacy-doc.cjs')
const { SUPPORTED_EXTENSIONS, isSupportedDocumentPath } = require('../electron/document-files.cjs')
const { validateDocxBytes, validateDocxPackage } = require('../electron/docx-files.cjs')

function fakeLegacyDocBytes() {
  const bytes = Buffer.alloc(512)
  OLE_COMPOUND_FILE_SIGNATURE.copy(bytes)
  return bytes
}

function fakeExtractorFor(documentOrError) {
  return class FakeWordExtractor {
    async extract(bytes) {
      assert.equal(bytes.subarray(0, 8).equals(OLE_COMPOUND_FILE_SIGNATURE), true)
      if (documentOrError instanceof Error) throw documentOrError
      return documentOrError
    }
  }
}

test('document routing recognizes legacy and modern Word extensions case-insensitively', () => {
  assert.deepEqual([...SUPPORTED_EXTENSIONS], ['.docx', '.doc'])
  assert.equal(isSupportedDocumentPath('C:\\Documents\\Legacy.DOC'), true)
  assert.equal(isSupportedDocumentPath('/tmp/report.docx'), true)
  assert.equal(isSupportedDocumentPath('/tmp/report.pdf'), false)
  assert.equal(isSupportedDocumentPath(null), false)
})

test('legacy validation requires a complete OLE compound-file signature', () => {
  assert.equal(validateLegacyDocBytes(fakeLegacyDocBytes()).length, 512)
  assert.throws(() => validateLegacyDocBytes(Buffer.alloc(511)), /empty or incomplete/i)
  assert.throws(() => validateLegacyDocBytes(Buffer.alloc(512)), /not a valid legacy \.doc/i)
  const partialSignature = Buffer.alloc(512)
  OLE_COMPOUND_FILE_SIGNATURE.subarray(0, 4).copy(partialSignature)
  assert.throws(() => validateLegacyDocBytes(partialSignature), /not a valid legacy \.doc/i)
})

test('editable DOCX generation preserves Unicode, tabs, blank lines, and manual page breaks', async () => {
  const docx = await buildEditableDocx({
    body: 'Résumé\t東京\n\nSecond paragraph\fSecond page',
    footnotes: 'A recovered footnote',
  }, { title: 'Legacy & notes.doc' })

  assert.equal(validateDocxBytes(docx), docx)
  assert.equal(validateDocxPackage(docx).mainDocumentPart, 'word/document.xml')
  const zip = await JSZip.loadAsync(docx)
  const documentXml = await zip.file('word/document.xml').async('string')
  const coreXml = await zip.file('docProps/core.xml').async('string')
  assert.match(documentXml, /Résumé/)
  assert.match(documentXml, /東京/)
  assert.match(documentXml, /<w:tab\/>/)
  assert.match(documentXml, /w:type="page"/)
  assert.match(documentXml, /Imported footnotes/)
  assert.match(coreXml, /Legacy &amp; notes/)

  const { runImport } = await import('@forevka/wordcanvas/import')
  const imported = runImport(new Uint8Array(docx), undefined, { collectMediaBytes: true })
  const importedJson = JSON.stringify(imported.doc)
  assert.match(importedJson, /Résumé/)
  assert.match(importedJson, /東京/)
  assert.match(importedJson, /Second page/)
  assert.match(importedJson, /A recovered footnote/)
})

test('legacy conversion uses the injected extractor and keeps auxiliary text editable', async () => {
  const calls = []
  const extracted = {
    getBody(options) {
      calls.push(['body', options])
      return 'Main body\tvalue'
    },
    getHeaders(options) {
      calls.push(['headers', options])
      return 'Header text'
    },
    getFooters() { return 'Footer text' },
    getFootnotes() { return '' },
    getEndnotes() { return '' },
    getAnnotations() { return 'Reviewer note' },
    getTextboxes(options) {
      calls.push(['textboxes', options])
      return 'Text box text'
    },
  }
  const result = await convertLegacyDocToDocx(fakeLegacyDocBytes(), {
    title: 'old-contract.doc',
    Extractor: fakeExtractorFor(extracted),
  })

  validateDocxPackage(result.data)
  assert.match(result.warnings.join(' '), /formatting, images, tables/i)
  assert.match(result.warnings.join(' '), /appended as editable sections/i)
  assert.deepEqual(calls[0], ['body', { filterUnicode: false }])
  assert.deepEqual(calls.find(([name]) => name === 'headers')[1], { filterUnicode: false, includeFooters: false })
  assert.deepEqual(calls.find(([name]) => name === 'textboxes')[1], {
    filterUnicode: false,
    includeHeadersAndFooters: false,
    includeBody: true,
  })

  const zip = await JSZip.loadAsync(result.data)
  const xml = await zip.file('word/document.xml').async('string')
  for (const expected of ['Main body', 'Header text', 'Footer text', 'Reviewer note', 'Text box text']) {
    assert.match(xml, new RegExp(expected))
  }
})

test('a modern DOCX mislabeled as .doc keeps its original editable package', async () => {
  const original = await buildEditableDocx({ body: 'Full fidelity mislabeled content' }, { title: 'mislabeled' })
  const result = await convertLegacyDocToDocx(original, {
    title: 'mislabeled.doc',
    Extractor: class ExtractorMustNotRun {
      extract() { throw new Error('the legacy extractor should not run') }
    },
  })
  assert.equal(result.data, original)
  assert.match(result.warnings.join(' '), /modern DOCX.*original editable content.*correct \.docx extension/i)
})

test('extractor failures become actionable errors without hiding password guidance', async () => {
  await assert.rejects(
    convertLegacyDocToDocx(fakeLegacyDocBytes(), { Extractor: fakeExtractorFor(new Error('invalid FIB')) }),
    /Open it in Word or LibreOffice, save it as \.docx/i,
  )
  await assert.rejects(
    convertLegacyDocToDocx(fakeLegacyDocBytes(), { Extractor: fakeExtractorFor(new Error('Encrypted package')) }),
    /Password-protected legacy Word documents are not supported.*Remove the password/i,
  )
})
