'use strict'

// The main-process save, open and export handlers (electron/main.cjs) driven like the renderer
// drives them, with a stand-in `electron` module: no window, no dialog, no Electron process.
// Runs in the "no document engine" configuration (SIMPLE_FORCE_NO_OFFICE=1).
const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const Module = require('node:module')
const os = require('node:os')
const path = require('node:path')
const XLSX = require('xlsx')

const sandbox = require('node:fs').mkdtempSync(path.join(os.tmpdir(), 'simple-calc-main-qa-'))
// The app cleans stale print folders in the temp directory on start; keep that inside the sandbox.
process.env.TEMP = sandbox
process.env.TMP = sandbox
process.env.SIMPLE_FORCE_NO_OFFICE = '1'

const handlers = new Map()
let nextSavePath = null
let lastSaveOptions = null
const fakeElectron = {
  app: {
    setName() {}, requestSingleInstanceLock: () => true, on() {}, quit() {}, isPackaged: false,
    getPath: () => path.join(sandbox, 'user-data'), getVersion: () => '0.0.0',
    whenReady: () => ({ then: (callback) => { setImmediate(callback); return Promise.resolve() } }),
  },
  BrowserWindow: class {
    constructor() { this.webContents = { on() {}, once() {}, setWindowOpenHandler() {}, send() {}, id: 1 } }
    removeMenu() {} loadURL() {} loadFile() {} once() {} on() {}
    static fromWebContents() { return null }
    static getAllWindows() { return [] }
    static getFocusedWindow() { return null }
  },
  dialog: {
    showSaveDialog: async (_window, options) => { lastSaveOptions = options; return nextSavePath ? { canceled: false, filePath: nextSavePath } : { canceled: true } },
    showOpenDialog: async () => ({ canceled: true, filePaths: [] }),
  },
  ipcMain: { handle: (channel, handler) => handlers.set(channel, handler), on: (channel, handler) => handlers.set(channel, handler) },
  shell: { showItemInFolder() {}, openExternal: async () => {} },
  clipboard: {},
}
const originalLoad = Module._load
Module._load = function load(request, parent, isMain) {
  return request === 'electron' ? fakeElectron : originalLoad.call(this, request, parent, isMain)
}

const frame = { url: 'file:///app/dist/index.html' }
const event = { sender: { id: 7, mainFrame: frame }, senderFrame: frame }
const call = (channel, ...args) => handlers.get(channel)(event, ...args)
const sha256 = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex')

async function main() {
  require('../electron/main.cjs')
  await new Promise((resolve) => setTimeout(resolve, 50))
  const directory = path.join(sandbox, 'files')
  await fs.mkdir(directory)
  assert.deepEqual(await call('workbook:capabilities'), { officeEngine: false })

  // CALC-SIE-2: Ctrl+S on an edited .xls writes Prices.xlsx beside it; the .xls is untouched.
  const legacy = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(legacy, XLSX.utils.aoa_to_sheet([['Item', 'Price'], ['Book', 12.5]]), 'Prices')
  const xlsPath = path.join(directory, 'Prices.xls')
  const xlsBytes = XLSX.write(legacy, { type: 'buffer', bookType: 'biff8' })
  await fs.writeFile(xlsPath, xlsBytes)
  const opened = await call('workbook:open-path', xlsPath)
  assert.equal(opened.requiresSaveAs, true)
  opened.workbook.sheets[0].cells.B2 = { value: 15 }
  const saved = await call('workbook:save', { documentId: opened.documentId, workbook: opened.workbook, saveAs: false, format: 'xls', suggestedName: opened.name, sourceUnmodified: false })
  assert.equal(saved.path, path.join(directory, 'Prices.xlsx'))
  assert.equal(saved.format, 'xlsx')
  assert.equal(saved.redirected, true)
  assert.match(saved.message, /original was not changed/)
  assert.equal(sha256(await fs.readFile(xlsPath)), sha256(xlsBytes))
  // Save As offers only formats this machine can write.
  nextSavePath = path.join(directory, 'Copy.xlsx')
  await call('workbook:save', { documentId: opened.documentId, workbook: opened.workbook, saveAs: true, format: 'xlsx', suggestedName: 'Prices.xlsx', sourceUnmodified: false })
  assert.deepEqual(lastSaveOptions.filters.flatMap((filter) => filter.extensions), ['xlsx', 'csv', 'tsv'])

  // CALC-SIE-3: a brand-new, untouched workbook can be saved.
  const created = await call('workbook:create')
  nextSavePath = path.join(directory, 'New.xlsx')
  const blank = { version: 1, name: 'Untitled', activeSheetId: 's', metadata: {}, sheets: [{ id: 's', name: 'Sheet1', rowCount: 100, colCount: 26, merges: [], cells: {} }] }
  assert.equal((await call('workbook:save', { documentId: created.documentId, workbook: blank, saveAs: true, format: 'xlsx', suggestedName: 'Untitled.xlsx', sourceUnmodified: true })).path, nextSavePath)

  // CALC-SIE-18: a dropped semicolon CSV keeps its dialect when saved as CSV.
  const dropped = await call('workbook:open-bytes', { name: 'werte.csv', data: Buffer.from('Name;Wert\r\nA;1,5\r\n') })
  dropped.workbook.sheets[0].cells.B2 = { value: 2.5 }
  nextSavePath = path.join(directory, 'werte.csv')
  await call('workbook:save', { documentId: dropped.documentId, workbook: dropped.workbook, saveAs: false, format: 'csv', suggestedName: 'werte.csv', sourceUnmodified: false })
  assert.equal(await fs.readFile(nextSavePath, 'utf8'), 'Name;Wert\r\nA;2,5\r\n')

  // CALC-SIE-14: XLS export needs consent before any dialog; ODS export works with basic formatting.
  lastSaveOptions = null
  await assert.rejects(call('workbook:export', { documentId: dropped.documentId, workbook: opened.workbook, format: 'xls', suggestedName: 'Prices.xls', sourceUnmodified: false }), /values only/)
  assert.equal(lastSaveOptions, null, 'no file name is asked for an export that cannot run')
  assert.equal((await call('workbook:export-check', { workbook: opened.workbook, format: 'xls' })).confirmationRequired, true)
  nextSavePath = path.join(directory, 'Export.xls')
  assert.equal((await call('workbook:export', { documentId: dropped.documentId, workbook: opened.workbook, format: 'xls', suggestedName: 'Prices.xls', sourceUnmodified: false, acceptLoss: true })).path, nextSavePath)
  nextSavePath = path.join(directory, 'Export.ods')
  assert.equal((await call('workbook:export', { documentId: dropped.documentId, workbook: opened.workbook, format: 'ods', suggestedName: 'Prices.ods', sourceUnmodified: false })).path, nextSavePath)

  // CALC-SIE-20: a file without an extension opens by content.
  const plain = path.join(directory, 'export')
  await fs.writeFile(plain, 'a,b\r\n1,2\r\n')
  assert.equal((await call('workbook:open-path', plain)).sourceFormat, 'csv')

  // Review F9: the renderer can ask first whether an XLS export needs consent (values only).
  const exportCheck = await call('workbook:export-check', { workbook: opened.workbook, format: 'xls' })
  assert.equal(exportCheck.confirmationRequired, true)
  assert.ok(Array.isArray(exportCheck.losses))
  // Review F13: the print dialog's printer list is answered.
  assert.equal(typeof handlers.get('workbook:list-printers'), 'function', 'the printer list handler is registered')
  // Review F6: a link never opens a network share in a new window.
  for (const share of ['\\\\server\\share\\x.xlsx', '//server/share/x.xlsx', '\\\\?\\UNC\\server\\share\\x.xlsx']) {
    await assert.rejects(call('workbook:new-window', share), /network locations/)
  }
  console.log('Main save QA passed: .xls saved as a sibling .xlsx without the engine, engine-aware Save As filters, new-workbook saves, dialect-preserving CSV saves, XLS/ODS exports without the engine, and extensionless opens.')
}

main()
  .catch((error) => { console.error(error); process.exitCode = 1 })
  .finally(async () => {
    Module._load = originalLoad
    await fs.rm(sandbox, { recursive: true, force: true }).catch(() => {})
  })
