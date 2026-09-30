'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { saveFormat, unchangedSourceBytes, assertSourceUnchanged } = require('../electron/workbook-save.cjs')
const { serializeWorkbook } = require('../electron/workbooks.cjs')

async function main() {
  for (const sourceFormat of ['xls', 'xlsx', 'csv', 'tsv', 'ods']) assert.equal(saveFormat({}, { sourceFormat }), sourceFormat)
  assert.equal(saveFormat({ sourceUnmodified: true }, { sourceFormat: 'xlsm' }), 'xlsm')
  assert.throws(() => saveFormat({}, { sourceFormat: 'xlsm' }), /original file has not been changed/)
  assert.equal(saveFormat({ format: 'xlsx' }, { sourceFormat: 'xls' }), 'xlsx')
  const bytes = Buffer.from('untouched source bytes')
  const copied = await unchangedSourceBytes({ sourceSnapshot: { bytes } }, true)
  assert.ok(copied.equals(bytes))
  assert.notEqual(copied, bytes)
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-calc-save-qa-'))
  try {
    const filePath = path.join(directory, 'foreign.xls')
    await fs.writeFile(filePath, bytes)
    const stat = await fs.stat(filePath)
    const record = { path: filePath, sourceSnapshot: { path: filePath, size: stat.size, modified: stat.mtimeMs } }
    assert.ok((await unchangedSourceBytes(record, true)).equals(bytes))
    await assertSourceUnchanged(record)
    await fs.writeFile(filePath, 'changed by another application')
    await assert.rejects(() => assertSourceUnchanged(record), /changed outside Simple/)
    await assert.rejects(() => unchangedSourceBytes(record), /changed outside Simple/)
    assert.equal(await fs.readFile(filePath, 'utf8'), 'changed by another application')
  } finally {
    const safeRoot = path.resolve(os.tmpdir()) + path.sep
    if (!path.resolve(directory).startsWith(safeRoot)) throw new Error('Unsafe QA cleanup path')
    await fs.rm(directory, { recursive: true, force: true })
  }
  const tooLarge = { version: 1, name: 'Limits', activeSheetId: 's', sheets: [{ id: 's', name: 'Sheet1', cells: { IW1: { value: 9 } }, rowCount: 1, colCount: 257, merges: [], rowHeights: {}, colWidths: {} }] }
  await assert.rejects(() => serializeWorkbook(tooLarge, 'xls'), /65,536 rows and 256 columns/)
  console.log('Save policy QA passed: original-format defaults, explicit conversion, exact byte copies, external-file conflicts, and XLS limits.')
}

main().catch((error) => { console.error(error); process.exitCode = 1 })
