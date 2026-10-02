const test = require('node:test')
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { EXPORT_FORMATS, MAX_EXPORT_BYTES, ODT_MIME_TYPE, availableExportFormats, exportFormat, validateExportBytes, validateExportContent } = require('../electron/export-files.cjs')

test('export format allowlist covers the document conversion menu', () => {
  assert.deepEqual(Object.keys(EXPORT_FORMATS), ['docx', 'pdf', 'html', 'md', 'txt', 'odt', 'rtf', 'doc'])
  assert.deepEqual(exportFormat('HTML'), {
    id: 'html', extension: '.html', label: 'Web page', mimeType: 'text/html;charset=utf-8',
  })
  assert.equal(exportFormat('odt').mimeType, ODT_MIME_TYPE)
  assert.equal(exportFormat('rtf').extension, '.rtf')
  assert.throws(() => exportFormat('exe'), /not supported/)
  assert.throws(() => exportFormat('../pdf'), /not supported/)
  assert.throws(() => exportFormat('constructor'), /not supported/)
})

test('.doc export is offered only when a local Office engine exists (DOC-SIE-24)', () => {
  const without = availableExportFormats({ officeEngine: false }).map((format) => format.id)
  assert.deepEqual(without, ['docx', 'pdf', 'html', 'md', 'txt', 'odt', 'rtf'])
  const withEngine = availableExportFormats({ officeEngine: true })
  assert.deepEqual(withEngine.map((format) => format.id), ['docx', 'pdf', 'html', 'md', 'txt', 'odt', 'rtf', 'doc'])
  assert.deepEqual(withEngine.at(-1), { id: 'doc', label: 'Word 97–2003 document', extension: '.doc' })
})

test('export byte validation accepts typed arrays without widening their view', () => {
  const source = new Uint8Array([0, 1, 2, 3])
  const bytes = validateExportBytes(source.subarray(1, 3))
  assert.deepEqual([...bytes], [1, 2])
  assert.throws(() => validateExportBytes(new Uint8Array()), /empty/)
  assert.throws(() => validateExportBytes('text'), /invalid/)
  assert.equal(MAX_EXPORT_BYTES, 512 * 1024 * 1024)
})

test('exported bytes must really be the chosen format before they reach the disk', async () => {
  assert.doesNotThrow(() => validateExportContent('rtf', Buffer.from('{\\rtf1\\ansi {\\b x\\}} \\{ }\n')))
  assert.throws(() => validateExportContent('rtf', Buffer.from('{\\rtf1 {unbalanced}')), /not valid|incomplete/)
  assert.throws(() => validateExportContent('rtf', Buffer.from('plain text')), /not valid/)
  assert.throws(() => validateExportContent('pdf', Buffer.from('not a pdf')), /PDF is not valid/)
  assert.doesNotThrow(() => validateExportContent('pdf', Buffer.from('%PDF-1.7\n')))
  assert.throws(() => validateExportContent('docx', Buffer.from('text')), /Word document is not valid/)

  const odt = new JSZip()
  odt.file('mimetype', ODT_MIME_TYPE, { compression: 'STORE' })
  odt.file('content.xml', '<x/>')
  const valid = await odt.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  assert.equal(validateExportContent('odt', valid), valid)
  const deflated = new JSZip()
  deflated.file('mimetype', ODT_MIME_TYPE.repeat(4), { compression: 'DEFLATE' })
  const deflatedBytes = await deflated.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  assert.throws(() => validateExportContent('odt', deflatedBytes), /not valid/)
  assert.throws(() => validateExportContent('odt', Buffer.from('PK')), /not valid/)
  const misordered = new JSZip()
  misordered.file('content.xml', '<x/>')
  misordered.file('mimetype', ODT_MIME_TYPE, { compression: 'STORE' })
  const misorderedBytes = await misordered.generateAsync({ type: 'nodebuffer' })
  assert.throws(() => validateExportContent('odt', misorderedBytes), /not valid/)
})
