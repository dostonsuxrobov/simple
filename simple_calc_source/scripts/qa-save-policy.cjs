'use strict'

const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const ExcelJS = require('exceljs')
const XLSX = require('xlsx')
const { saveFormat, unchangedSourceBytes, assertSourceUnchanged, saveDocument, siblingPath, saveFilters, hasOriginalBytes } = require('../electron/workbook-save.cjs')
const { serializeWorkbook, importWorkbookPath, workbookPayloadFromBytes, delimitedDialectFor } = require('../electron/workbooks.cjs')

const sha256 = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex')

function dependencies(overrides = {}) {
  const calls = []
  return {
    calls,
    officeEngine: false,
    serialize: async (workbook, format, options) => {
      calls.push({ format, options })
      return serializeWorkbook(workbook, format, options)
    },
    writeFile: (target, bytes) => fs.writeFile(target, bytes),
    dialectFor: delimitedDialectFor,
    backupDirectory: null,
    chooseSavePath: async () => { throw new Error('No save dialog was expected.') },
    ...overrides,
  }
}

/** A document record the way main.cjs registers an opened file. */
async function openRecord(filePath) {
  const imported = await importWorkbookPath(filePath, { officeEngine: false })
  const record = {
    path: filePath,
    sourceFormat: imported.payload.sourceFormat,
    originalName: imported.payload.name,
    sourceSnapshot: { path: filePath, size: imported.stat.size, modified: imported.stat.mtimeMs },
    mergeBase: imported.mergeBase,
    dialect: imported.payload.workbook.metadata.dialect || null,
  }
  return { record, payload: imported.payload }
}

async function legacyPolicies() {
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
    await cleanup(directory)
  }
  const tooLarge = { version: 1, name: 'Limits', activeSheetId: 's', sheets: [{ id: 's', name: 'Sheet1', cells: { IW1: { value: 9 } }, rowCount: 1, colCount: 257, merges: [], rowHeights: {}, colWidths: {} }] }
  await assert.rejects(() => serializeWorkbook(tooLarge, 'xls', { officeEngine: false, valuesOnly: true }), /65,536 rows and 256 columns/)
}

async function newWorkbook(directory) {
  // CALC-SIE-3 / calc-file-io-objects-12: a brand-new, still-empty workbook can be saved.
  const record = { path: null, sourceFormat: 'xlsx', originalName: 'Untitled.xlsx' }
  assert.equal(hasOriginalBytes(record), false)
  const target = path.join(directory, 'Budget.xlsx')
  const workbook = { version: 1, name: 'Untitled', activeSheetId: 's1', metadata: {}, sheets: [{ id: 's1', name: 'Sheet1', rowCount: 100, colCount: 26, merges: [], cells: {} }] }
  const result = await saveDocument(record, { workbook, saveAs: true, format: 'xlsx', suggestedName: 'Untitled.xlsx', sourceUnmodified: true }, dependencies({
    chooseSavePath: async ({ format }) => { assert.equal(format, 'xlsx'); return target },
  }))
  assert.equal(result.path, target)
  const saved = new ExcelJS.Workbook()
  await saved.xlsx.load(await fs.readFile(target))
  assert.equal(saved.worksheets.length, 1)
  assert.equal(record.path, target)
}

async function legacyWithoutEngine(directory) {
  // CALC-SIE-2: an edited .xls saves its edits to an .xlsx next to the original, which stays
  // byte-identical; the document is rebound to the new file.
  const legacy = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(legacy, XLSX.utils.aoa_to_sheet([['Item', 'Price'], ['Book', 12.5]]), 'Prices')
  const xlsPath = path.join(directory, 'Prices.xls')
  const original = XLSX.write(legacy, { type: 'buffer', bookType: 'biff8' })
  await fs.writeFile(xlsPath, original)
  const first = await openRecord(xlsPath)
  assert.equal(first.payload.requiresSaveAs, true)
  first.payload.workbook.sheets[0].cells.B2 = { value: 15 }
  const deps = dependencies()
  const result = await saveDocument(first.record, { workbook: first.payload.workbook, saveAs: false, format: 'xls', suggestedName: 'Prices.xls', sourceUnmodified: false }, deps)
  assert.equal(result.path, path.join(directory, 'Prices.xlsx'))
  assert.equal(result.format, 'xlsx')
  assert.equal(result.redirected, true)
  assert.equal(result.originalPath, xlsPath)
  assert.match(result.message, /next to "Prices\.xls"/)
  assert.equal(sha256(await fs.readFile(xlsPath)), sha256(original), 'the original .xls is not changed')
  const reopened = await workbookPayloadFromBytes('Prices.xlsx', await fs.readFile(result.path))
  assert.equal(reopened.workbook.sheets[0].cells.B2.value, 15)
  assert.equal(first.record.path, result.path)
  assert.equal(first.record.sourceFormat, 'xlsx')
  // The next Ctrl+S writes the rebound .xlsx in place.
  first.payload.workbook.sheets[0].cells.B2 = { value: 16 }
  const again = await saveDocument(first.record, { workbook: first.payload.workbook, saveAs: false, format: 'xlsx', sourceUnmodified: false }, dependencies())
  assert.equal(again.path, result.path)
  assert.equal(again.redirected, undefined)
  // A second document from the same .xls never overwrites the existing copy.
  const second = await openRecord(xlsPath)
  const next = await saveDocument(second.record, { workbook: second.payload.workbook, saveAs: false, format: 'xls', sourceUnmodified: false }, dependencies())
  assert.equal(next.path, path.join(directory, 'Prices (edited).xlsx'))
  // Save As offers XLSX; a name typed with .xls still produces an .xlsx beside it.
  const third = await openRecord(xlsPath)
  let offered = null
  const typed = await saveDocument(third.record, { workbook: third.payload.workbook, saveAs: true, format: 'xls', sourceUnmodified: false }, dependencies({
    chooseSavePath: async ({ format }) => { offered = format; return path.join(directory, 'Typed.xls') },
  }))
  assert.equal(offered, 'xlsx')
  assert.equal(typed.path, path.join(directory, 'Typed.xlsx'))
  assert.equal(typed.redirected, true)
  await assert.rejects(fs.stat(path.join(directory, 'Typed.xls')), /ENOENT/)
  // An unchanged .xls can still be copied as .xls without the engine (no conversion needed).
  const fourth = await openRecord(xlsPath)
  const copy = await saveDocument(fourth.record, { workbook: fourth.payload.workbook, saveAs: true, format: 'xls', sourceUnmodified: true }, dependencies({
    chooseSavePath: async ({ format }) => { assert.equal(format, 'xls'); return path.join(directory, 'Copy.xls') },
  }))
  assert.equal(sha256(await fs.readFile(copy.path)), sha256(original))
  assert.equal(await siblingPath(path.join(directory, 'Prices.xlsx'), 'xlsx'), path.join(directory, 'Prices (edited 2).xlsx'), 'the original itself is never chosen')
  // Save As filters follow the engine.
  const extensions = (filters) => filters.flatMap((filter) => filter.extensions)
  assert.deepEqual(extensions(saveFilters('xlsx', { officeEngine: false })), ['xlsx', 'csv', 'tsv'])
  assert.ok(extensions(saveFilters('xlsx', { officeEngine: true })).includes('xls'))
}

async function mergeBase(directory) {
  // calc-file-io-objects-1: every XLSX save of a document merges over the package it was
  // imported from, never over the file the previous save wrote.
  const anchor = { from: { row: 5, col: 1, rowOffsetEmu: 0, colOffsetEmu: 0 }, to: { row: 15, col: 6, rowOffsetEmu: 0, colOffsetEmu: 0 } }
  const sheet = (id, name) => ({ id, name, rowCount: 3, colCount: 2, merges: [], colWidths: {}, rowHeights: {}, cells: { A1: { value: 'x' }, B1: { value: 1 }, B2: { value: 2 }, B3: { value: 3 } } })
  const second = sheet('sheet-2', 'Sheet2')
  second.charts = [{ id: 'cA', type: 'column', title: 'Chart A', anchor, series: [{ id: 's', valuesRef: 'Sheet2!B1:B3' }] }]
  const sourcePath = path.join(directory, 'charts.xlsx')
  await fs.writeFile(sourcePath, await serializeWorkbook({ version: 1, name: 'charts', activeSheetId: 'sheet-1', sheets: [sheet('sheet-1', 'Sheet1'), second], metadata: {} }, 'xlsx'))
  const { record, payload } = await openRecord(sourcePath)
  const imported = Buffer.from(record.mergeBase)
  payload.workbook.sheets[0].charts = [{ id: 'cN', type: 'line', title: 'Chart N', anchor, series: [{ id: 's', valuesRef: 'Sheet1!B1:B3' }] }]
  const deps = dependencies()
  await saveDocument(record, { workbook: payload.workbook, saveAs: false, format: 'xlsx', sourceUnmodified: false }, deps)
  payload.workbook.sheets[0].cells.C1 = { value: 'edit' }
  await saveDocument(record, { workbook: payload.workbook, saveAs: false, format: 'xlsx', sourceUnmodified: false }, deps)
  assert.equal(deps.calls.length, 2)
  for (const call of deps.calls) assert.ok(call.options.baseBytes.equals(imported), 'the imported package is the base of every save')
  const reopened = await workbookPayloadFromBytes('charts.xlsx', await fs.readFile(sourcePath))
  assert.equal(reopened.workbook.sheets[1].charts[0].title, 'Chart A', 'Sheet2 keeps its own chart after two saves')
  assert.equal(reopened.workbook.sheets[0].charts[0].title, 'Chart N')
  // Dropped (path-less) bytes and an unchanged copy still work.
  assert.equal(hasOriginalBytes({ sourceSnapshot: { bytes: Buffer.from('x') } }), true)
}

async function delimitedInPlace(directory) {
  // CALC-SIE-18: Ctrl+S on a semicolon / decimal-comma cp1252 CSV writes it back the same way.
  const csvPath = path.join(directory, 'Konto.csv')
  const source = Buffer.from('Name;Betrag;Datum\r\nM\xfcller;1.234,56;15.03.2024\r\nSch\xf6n;-7,50;01.02.2024\r\n', 'latin1')
  await fs.writeFile(csvPath, source)
  const { record, payload } = await openRecord(csvPath)
  payload.workbook.sheets[0].cells.A3 = { value: 'Sch\xf6ner' }
  delete payload.workbook.metadata.dialect // the record's copy is used when the renderer drops it
  await saveDocument(record, { workbook: payload.workbook, saveAs: false, format: 'csv', sourceUnmodified: false }, dependencies())
  assert.equal((await fs.readFile(csvPath)).toString('latin1'), 'Name;Betrag;Datum\r\nM\xfcller;1.234,56;15.03.2024\r\nSch\xf6ner;-7,50;01.02.2024\r\n')
}

async function cleanup(directory) {
  const safeRoot = path.resolve(os.tmpdir()) + path.sep
  if (!path.resolve(directory).startsWith(safeRoot)) throw new Error('Unsafe QA cleanup path')
  await fs.rm(directory, { recursive: true, force: true })
}

async function main() {
  await legacyPolicies()
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-calc-save-qa-'))
  try {
    await newWorkbook(directory)
    await legacyWithoutEngine(directory)
    await mergeBase(directory)
    await delimitedInPlace(directory)
  } finally {
    await cleanup(directory)
  }
  console.log('Save policy QA passed: original-format defaults, explicit conversion, exact byte copies, external-file conflicts, XLS limits, new-workbook saves, .xls/.ods saved as .xlsx without the engine, the import package as merge base, and dialect-preserving CSV saves.')
}

main().catch((error) => { console.error(error); process.exitCode = 1 })
