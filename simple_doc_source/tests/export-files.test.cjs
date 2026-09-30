const test = require('node:test')
const assert = require('node:assert/strict')
const { EXPORT_FORMATS, MAX_EXPORT_BYTES, exportFormat, validateExportBytes } = require('../electron/export-files.cjs')

test('export format allowlist covers the document conversion menu', () => {
  assert.deepEqual(Object.keys(EXPORT_FORMATS), ['docx', 'pdf', 'html', 'md', 'txt'])
  assert.deepEqual(exportFormat('HTML'), {
    id: 'html', extension: '.html', label: 'Web page', mimeType: 'text/html;charset=utf-8',
  })
  assert.throws(() => exportFormat('exe'), /not supported/)
  assert.throws(() => exportFormat('../pdf'), /not supported/)
})

test('export byte validation accepts typed arrays without widening their view', () => {
  const source = new Uint8Array([0, 1, 2, 3])
  const bytes = validateExportBytes(source.subarray(1, 3))
  assert.deepEqual([...bytes], [1, 2])
  assert.throws(() => validateExportBytes(new Uint8Array()), /empty/)
  assert.throws(() => validateExportBytes('text'), /invalid/)
  assert.equal(MAX_EXPORT_BYTES, 512 * 1024 * 1024)
})
