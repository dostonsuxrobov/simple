'use strict'

// Save / Open dialogs, decision prompts and preferences of the shared layer
// (io-dialogs.cjs, design §3.4, §3.5, §6.4), and the io:* IPC handlers and
// preload bridge behind window.simpleIO (io-ipc.cjs, io-bridge.cjs, §8). The
// `electron` module is replaced by a small stub that records every dialog and
// shell call, so no window or native dialog ever opens. The preload bridge is
// run in a separate context, wired to the stub's ipcMain like a real page.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const vm = require('node:vm')
const Module = require('node:module')
const { EventEmitter } = require('node:events')
const { pathToFileURL } = require('node:url')

const SHARED = path.resolve(__dirname, '..', 'shared')
const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-dialogs-'))
const APP_PATH = path.join(ROOT, 'app')
fs.mkdirSync(path.join(APP_PATH, 'dist'), { recursive: true })

// ---------------------------------------------------------------------------
// Electron stub
// ---------------------------------------------------------------------------

function makeIpcMain() {
  const ipcMain = new EventEmitter()
  ipcMain.handlers = new Map()
  ipcMain.handle = (channel, handler) => {
    if (ipcMain.handlers.has(channel)) throw new Error(`Attempted to register a second handler for '${channel}'`)
    ipcMain.handlers.set(channel, handler)
  }
  ipcMain.removeHandler = (channel) => { ipcMain.handlers.delete(channel) }
  return ipcMain
}

function createElectronStub() {
  const record = { save: [], open: [], box: [], shown: [], openPath: [], openExternal: [] }
  const queue = { save: [], open: [], box: [] }
  const windows = new Set()
  let nextId = 1
  const take = async (list, options) => {
    if (!list.length) return undefined
    const next = list.shift()
    return typeof next === 'function' ? next(options) : next
  }
  class FakeWebContents extends EventEmitter {
    constructor(url) {
      super()
      this.id = nextId++
      this.destroyed = false
      this.mainFrame = { url }
      this.sent = []
      this.renderer = null
    }
    send(channel, ...args) {
      if (this.destroyed) throw new Error('Object has been destroyed')
      this.sent.push({ channel, args })
      if (this.renderer) this.renderer(channel, ...args)
    }
    isDestroyed() { return this.destroyed }
    destroy() {
      if (this.destroyed) return
      this.destroyed = true
      this.emit('destroyed')
    }
  }
  class BrowserWindow extends EventEmitter {
    constructor(options = {}) {
      super()
      this.webContents = new FakeWebContents(options.url || pathToFileURL(path.join(APP_PATH, 'dist', 'index.html')).href)
      this.destroyed = false
      this.focusCount = 0
      windows.add(this)
    }
    static fromWebContents(contents) {
      for (const win of windows) if (win.webContents === contents) return win
      return null
    }
    static getAllWindows() { return [...windows] }
    isDestroyed() { return this.destroyed }
    destroy() {
      if (this.destroyed) return
      this.destroyed = true
      windows.delete(this)
      this.webContents.destroy()
      this.emit('closed')
    }
    focus() { this.focusCount += 1 }
    show() {}
    isMinimized() { return false }
    restore() {}
  }
  const app = new EventEmitter()
  Object.assign(app, {
    isPackaged: false,
    paths: {},
    getPath(name) {
      if (!this.paths[name]) throw new Error(`no ${name} path in the stub`)
      return this.paths[name]
    },
    getAppPath() { return APP_PATH },
    getVersion() { return '9.9.9' },
  })
  const dialog = {
    async showSaveDialog(win, options) {
      const opts = options || win
      record.save.push(opts)
      const answer = await take(queue.save, opts)
      return answer ? { canceled: false, filePath: answer } : { canceled: true, filePath: '' }
    },
    async showOpenDialog(win, options) {
      const opts = options || win
      record.open.push(opts)
      const answer = await take(queue.open, opts)
      return Array.isArray(answer) && answer.length ? { canceled: false, filePaths: answer } : { canceled: true, filePaths: [] }
    },
    async showMessageBox(win, options) {
      const opts = options || win
      record.box.push(opts)
      let answer = await take(queue.box, opts)
      if (answer === undefined) answer = opts.cancelId
      const index = typeof answer === 'number' ? answer : opts.buttons.indexOf(answer)
      if (index < 0) throw new Error(`No button "${answer}" in ${opts.buttons.join(', ')}`)
      return { response: index, checkboxChecked: false }
    },
  }
  const shell = {
    showItemInFolder(target) { record.shown.push(target) },
    openPath(target) { record.openPath.push(target); return Promise.resolve('') },
    openExternal(target) { record.openExternal.push(target); return Promise.resolve() },
  }
  const clipboard = {
    content: {},
    readText() { return this.content.text || '' },
    readHTML() { return this.content.html || '' },
    readRTF() { return this.content.rtf || '' },
    readImage() {
      const png = this.content.png
      return { isEmpty: () => !png, toPNG: () => png || Buffer.alloc(0) }
    },
    readBuffer(format) { return format === 'FileNameW' && this.content.file ? Buffer.from(`${this.content.file}\u0000`, 'utf16le') : Buffer.alloc(0) },
  }
  return { app, BrowserWindow, dialog, ipcMain: makeIpcMain(), shell, clipboard, record, queue, windows }
}

const electron = createElectronStub()
const originalResolve = Module._resolveFilename
Module._resolveFilename = function resolveElectronStub(request, ...rest) {
  if (request === 'electron') return 'electron-test-stub'
  return originalResolve.call(this, request, ...rest)
}
require.cache['electron-test-stub'] = { id: 'electron-test-stub', filename: 'electron-test-stub', loaded: true, exports: electron }

const core = require(path.join(SHARED, 'electron', 'io-core.cjs'))
const dialogs = require(path.join(SHARED, 'electron', 'io-dialogs.cjs'))
const ioIpc = require(path.join(SHARED, 'electron', 'io-ipc.cjs'))
const guard = require(path.join(SHARED, 'electron', 'document-guard.cjs'))
const stores = require(path.join(SHARED, 'electron', 'stores.cjs'))

core.configureIo({ journalDir: path.join(ROOT, 'journal'), logger: { warn() {}, info() {} } })
stores.configureStores({ recoveryDir: path.join(ROOT, 'recovery'), versionsDir: path.join(ROOT, 'versions'), logger: { warn() {} } })
dialogs.configureDialogs({ logger: { warn() {} } })
ioIpc.configureIoIpc({ logger: { warn() {} } })
guard.configureDocumentGuard({ logger: { warn() {} } })

test.after(() => {
  Module._resolveFilename = originalResolve
  fs.rmSync(ROOT, { recursive: true, force: true })
})

let caseCount = 0
function scenario(name, module = 'calc') {
  caseCount += 1
  const folder = path.join(ROOT, `${String(caseCount).padStart(2, '0')}-${name}`)
  const layout = {
    folder,
    userData: path.join(folder, 'userData'),
    documents: path.join(folder, 'Documents'),
    desktop: path.join(folder, 'Desktop'),
    work: path.join(folder, 'Work'),
  }
  for (const key of ['documents', 'desktop', 'work']) fs.mkdirSync(layout[key], { recursive: true })
  dialogs.configureDialogs({ module, userData: layout.userData, documentsDir: layout.documents, desktopDir: layout.desktop })
  electron.queue.save.length = 0
  electron.queue.open.length = 0
  electron.queue.box.length = 0
  return layout
}

function lastSave() {
  return electron.record.save[electron.record.save.length - 1]
}

function lastBox() {
  return electron.record.box[electron.record.box.length - 1]
}

function readPrefs(layout) {
  return JSON.parse(fs.readFileSync(path.join(layout.userData, 'io-prefs.json'), 'utf8'))
}

/** Runs the real preload bridge in its own context, wired to an ipcMain like a page of `win`. */
function loadBridge(win, ipcMain, { argv = [], env = {} } = {}) {
  const source = fs.readFileSync(path.join(SHARED, 'preload', 'io-bridge.cjs'), 'utf8')
  const listeners = new Map()
  const event = () => ({ sender: win.webContents, senderFrame: win.webContents.mainFrame })
  const ipcRenderer = {
    on(channel, listener) {
      if (!listeners.has(channel)) listeners.set(channel, [])
      listeners.get(channel).push(listener)
    },
    send(channel, ...args) { ipcMain.emit(channel, event(), ...structuredClone(args)) },
    invoke(channel, ...args) {
      const handler = ipcMain.handlers.get(channel)
      if (!handler) return Promise.reject(new Error(`No handler registered for '${channel}'`))
      return Promise.resolve().then(() => handler(event(), ...structuredClone(args))).then((value) => structuredClone(value))
    },
  }
  win.webContents.renderer = (channel, ...args) => {
    for (const listener of listeners.get(channel) || []) listener({}, ...structuredClone(args))
  }
  let exposed = null
  const preloadElectron = {
    contextBridge: { exposeInMainWorld: (name, api) => { exposed = { name, api } } },
    ipcRenderer,
    webUtils: { getPathForFile: (file) => (file && file.path) || '' },
  }
  vm.runInNewContext(source, {
    require: (id) => {
      if (id !== 'electron') throw new Error(`the bridge may only require electron, not ${id}`)
      return preloadElectron
    },
    process: { argv, env },
  })
  assert.equal(exposed.name, 'simpleIO')
  return exposed.api
}

// ---------------------------------------------------------------------------
// Save dialog: folders, names and extension rules
// ---------------------------------------------------------------------------

test('Save As starts in the document\'s own folder with its name, and remembers the folder', async () => {
  const layout = scenario('source-folder')
  const source = path.join(layout.work, 'Budget.xlsx')
  fs.writeFileSync(source, 'x')
  electron.queue.save.push(path.join(layout.work, 'Budget 2026.xlsx'))
  const chosen = await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Budget.xlsx', format: 'xlsx', formats: ['xlsx', 'csv'] }, { sourcePath: source })
  assert.deepEqual(chosen, { path: path.join(layout.work, 'Budget 2026.xlsx'), format: 'xlsx' })
  assert.equal(lastSave().defaultPath, source)
  assert.deepEqual(lastSave().filters, [
    { name: 'Excel workbook (*.xlsx)', extensions: ['xlsx'] },
    { name: 'Comma-separated values (*.csv)', extensions: ['csv'] },
  ])
  assert.equal(readPrefs(layout).lastSaveFolder, layout.work)

  // The registry decides the formats when the request names none: Calc saves xlsx, csv and tsv in place.
  electron.queue.save.push(null)
  assert.equal(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Budget.xlsx', format: 'xlsx' }, { sourcePath: source }), null)
  assert.deepEqual(lastSave().filters.map((filter) => filter.extensions[0]), ['xlsx', 'csv', 'tsv'])
})

test('an untitled document starts in the remembered save folder, else in Documents', async () => {
  const layout = scenario('untitled')
  electron.queue.save.push(null)
  await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Untitled spreadsheet', format: 'xlsx' })
  assert.equal(lastSave().defaultPath, path.join(layout.documents, 'Untitled spreadsheet.xlsx'))
  assert.equal(await dialogs.setPref('lastSaveFolder', layout.work), true)
  electron.queue.save.push(null)
  await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Untitled spreadsheet', format: 'xlsx' })
  assert.equal(lastSave().defaultPath, path.join(layout.work, 'Untitled spreadsheet.xlsx'))
  // A source folder that is gone falls back the same way.
  electron.queue.save.push(null)
  await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Old.xlsx', format: 'xlsx' }, { sourcePath: path.join(layout.folder, 'gone', 'Old.xlsx') })
  assert.equal(lastSave().defaultPath, path.join(layout.work, 'Old.xlsx'))
})

test('Export starts in the remembered export folder, then the document\'s folder', async () => {
  const layout = scenario('export-folder')
  const source = path.join(layout.work, 'Budget.xlsx')
  fs.writeFileSync(source, 'x')
  electron.queue.save.push((options) => options.defaultPath)
  const first = await dialogs.chooseSavePath({ purpose: 'export', name: 'Budget.xlsx', format: 'pdf' }, { sourcePath: source })
  assert.deepEqual(first, { path: path.join(layout.work, 'Budget.pdf'), format: 'pdf' })
  const exportsFolder = path.join(layout.folder, 'Exports')
  fs.mkdirSync(exportsFolder)
  electron.queue.save.push(path.join(exportsFolder, 'Budget.pdf'))
  await dialogs.chooseSavePath({ purpose: 'export', name: 'Budget.xlsx', format: 'pdf' }, { sourcePath: source })
  assert.equal(readPrefs(layout).lastExportFolder, exportsFolder)
  electron.queue.save.push(null)
  await dialogs.chooseSavePath({ purpose: 'export', name: 'Budget.xlsx', format: 'csv' }, { sourcePath: source })
  assert.equal(lastSave().defaultPath, path.join(exportsFolder, 'Budget.csv'))
  assert.equal(readPrefs(layout).lastSaveFolder, undefined, 'exports never change the save folder')
})

test('extension rules: unknown endings are kept, a writable one switches the format, a known other one is replaced', () => {
  const docx = { format: 'docx', formats: ['docx'] }
  const images = { format: 'png', formats: ['png', 'jpeg', 'webp'] }
  const sheets = { format: 'xlsx', formats: ['xlsx', 'csv', 'tsv'] }
  const rule = (name, rules) => {
    const result = dialogs.applyExtensionRules(path.join(ROOT, name), rules)
    return { name: path.basename(result.path), format: result.format }
  }
  assert.deepEqual(rule('Q3 plan v2.1', docx), { name: 'Q3 plan v2.1.docx', format: 'docx' })
  assert.deepEqual(rule('Q3 plan', docx), { name: 'Q3 plan.docx', format: 'docx' })
  assert.deepEqual(rule('Budget.2024.xlsx', sheets), { name: 'Budget.2024.xlsx', format: 'xlsx' })
  assert.deepEqual(rule('photo.jpg', images), { name: 'photo.jpg', format: 'jpeg' })
  assert.deepEqual(rule('photo.JPEG', images), { name: 'photo.JPEG', format: 'jpeg' })
  assert.deepEqual(rule('photo.jpg.png', images), { name: 'photo.jpg', format: 'jpeg' }, 'never photo.jpg.png')
  assert.deepEqual(rule('report.pdf', docx), { name: 'report.docx', format: 'docx' }, 'a known extension is replaced, not appended to')
  assert.deepEqual(rule('report.pdf.docx', docx), { name: 'report.docx', format: 'docx' })
  assert.deepEqual(rule('data.csv', sheets), { name: 'data.csv', format: 'csv' })
  assert.deepEqual(rule('notes.txt', sheets), { name: 'notes.xlsx', format: 'xlsx' })
  assert.equal(dialogs.safeStem('Q3 plan v2.1'), 'Q3 plan v2.1')
  assert.equal(dialogs.safeStem('photo.jpg'), 'photo')
  assert.equal(dialogs.nameForFormat('photo.jpeg', 'jpeg'), 'photo.jpeg')
  assert.equal(dialogs.nameForFormat('photo.png', 'jpeg'), 'photo.jpg')
  assert.equal(dialogs.nameForFormat('Q3 plan v2.1', 'docx'), 'Q3 plan v2.1.docx')
})

test('the Save dialog applies the rules: photo.jpg in a PNG Save As becomes a JPEG', async () => {
  const layout = scenario('rules-dialog', 'image')
  electron.queue.save.push(path.join(layout.work, 'photo.jpg'))
  assert.deepEqual(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'photo.png', format: 'png' }), { path: path.join(layout.work, 'photo.jpg'), format: 'jpeg' })
  electron.queue.save.push(path.join(layout.work, 'photo.jpg.png'))
  assert.deepEqual(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'photo.png', format: 'png' }), { path: path.join(layout.work, 'photo.jpg'), format: 'jpeg' })
  dialogs.configureDialogs({ module: 'docs' })
  electron.queue.save.push(path.join(layout.work, 'Q3 plan v2.1'))
  assert.deepEqual(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Q3 plan v2.1', format: 'docx' }), { path: path.join(layout.work, 'Q3 plan v2.1.docx'), format: 'docx' })
  assert.equal(electron.record.box.filter((box) => box.message.includes('already exists')).length, 0, 'no confirm when nothing is replaced')
})

test('a name the rules changed asks before replacing a file, with Cancel as the default', async () => {
  const layout = scenario('collision', 'docs')
  const existing = path.join(layout.work, 'Q3 plan v2.1.docx')
  fs.writeFileSync(existing, 'someone else\'s file')
  electron.queue.save.push(path.join(layout.work, 'Q3 plan v2.1'), path.join(layout.work, 'Q3 plan v3.docx'))
  electron.queue.box.push('Cancel')
  const second = await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Q3 plan v2.1', format: 'docx' })
  const box = electron.record.box.find((item) => item.message === '"Q3 plan v2.1.docx" already exists. Do you want to replace it?')
  assert.ok(box, 'the replace confirm was shown')
  assert.deepEqual(box.buttons, ['Replace', 'Cancel'])
  assert.equal(box.defaultId, 1)
  assert.equal(box.cancelId, 1)
  assert.equal(lastSave().defaultPath, existing, 'Cancel goes back to the dialog with the final name')
  assert.deepEqual(second, { path: path.join(layout.work, 'Q3 plan v3.docx'), format: 'docx' })
  electron.queue.save.push(path.join(layout.work, 'Q3 plan v2.1'))
  electron.queue.box.push('Replace')
  assert.deepEqual(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Q3 plan v2.1', format: 'docx' }), { path: existing, format: 'docx' })
})

test('a file that can\'t be saved where it is gets Save As with the same name in a folder that works', async () => {
  const layout = scenario('fallback')
  const failing = path.join(layout.work, 'Budget.xlsx')
  fs.writeFileSync(failing, 'x')
  const suggest = (reason, sourcePath = failing) => {
    electron.queue.save.push(null)
    return dialogs.chooseSavePath({ purpose: 'fallback', reason, name: 'Budget.xlsx', format: 'xlsx' }, { sourcePath }).then(() => lastSave().defaultPath)
  }
  assert.equal(await suggest('READ_ONLY'), path.join(layout.documents, 'Budget.xlsx'))
  for (const code of ['LOCKED', 'NO_PERMISSION', 'READ_ONLY_VOLUME', 'FILE_UNAVAILABLE', 'NAME_TOO_LONG']) {
    assert.equal(await suggest(code), path.join(layout.documents, 'Budget.xlsx'), code)
  }
  assert.equal(await suggest('FOLDER_MISSING', path.join(layout.folder, 'unplugged', 'Budget.xlsx')), path.join(layout.documents, 'Budget.xlsx'))
  // The remembered save folder comes first, unless it is the failing folder.
  const other = path.join(layout.folder, 'Other')
  fs.mkdirSync(other)
  await dialogs.setPref('lastSaveFolder', other)
  assert.equal(await suggest('LOCKED'), path.join(other, 'Budget.xlsx'))
  await dialogs.setPref('lastSaveFolder', layout.work)
  assert.equal(await suggest('LOCKED'), path.join(layout.documents, 'Budget.xlsx'))
  // A Documents folder that is not there falls through to the Desktop.
  dialogs.configureDialogs({ documentsDir: path.join(layout.folder, 'NoDocuments') })
  assert.equal(await suggest('READ_ONLY'), path.join(layout.desktop, 'Budget.xlsx'))
  dialogs.configureDialogs({ documentsDir: layout.documents })
  // A conflict keeps both versions: "<name> (edited)" beside the original.
  assert.equal(await suggest('CHANGED_ON_DISK'), path.join(layout.work, 'Budget (edited).xlsx'))
  fs.writeFileSync(path.join(layout.work, 'Budget (edited).xlsx'), 'x')
  assert.equal(await suggest('SOURCE_MISSING'), path.join(layout.work, 'Budget (edited 2).xlsx'))
  // A failure that is about the format, not the folder, starts in the document's folder.
  assert.equal(await suggest('SERIALIZE_FAILED'), failing)
})

test('sibling files are named "<stem>.<ext>", then "<stem> (edited).<ext>"', async () => {
  const layout = scenario('sibling')
  const original = path.join(layout.work, 'Q3 Budget.xls')
  fs.writeFileSync(original, 'legacy')
  assert.equal(await dialogs.siblingPathFor(original, 'xlsx'), path.join(layout.work, 'Q3 Budget.xlsx'))
  fs.writeFileSync(path.join(layout.work, 'Q3 Budget.xlsx'), 'unrelated')
  assert.equal(await dialogs.siblingPathFor(original, 'xlsx'), path.join(layout.work, 'Q3 Budget (edited).xlsx'))
  assert.equal(await dialogs.siblingPathFor(original, 'xlsx', { reuse: path.join(layout.work, 'Q3 Budget.xlsx') }), path.join(layout.work, 'Q3 Budget.xlsx'))
  fs.writeFileSync(path.join(layout.work, 'Q3 Budget (edited).xlsx'), 'unrelated')
  assert.equal(await dialogs.siblingPathFor(original, 'xlsx'), path.join(layout.work, 'Q3 Budget (edited 2).xlsx'))
  const signed = path.join(layout.work, 'Contract.pdf')
  assert.equal(await dialogs.siblingPathFor(signed, 'pdf', { edited: true }), path.join(layout.work, 'Contract (edited).pdf'))
  assert.equal(await dialogs.siblingPathFor(signed, 'pdf'), path.join(layout.work, 'Contract (edited).pdf'), 'never the original itself')
})

// ---------------------------------------------------------------------------
// Prompts, Open dialog, preferences, QA script
// ---------------------------------------------------------------------------

test('prompts use the catalog: Save is the default, Esc is Cancel, and failures end with the details', async () => {
  scenario('prompts')
  electron.queue.box.push('Save')
  assert.equal(await dialogs.showPrompt('prompts.unsaved', { name: 'Budget.xlsx' }), 'save')
  const unsaved = lastBox()
  assert.deepEqual(unsaved.buttons, ['Save', 'Don\'t Save', 'Cancel'])
  assert.equal(unsaved.defaultId, 0)
  assert.equal(unsaved.cancelId, 2)
  assert.equal(unsaved.noLink, true)
  assert.equal(unsaved.title, 'Simple Spreadsheets')
  assert.equal(unsaved.message, 'Do you want to save the changes to "Budget.xlsx"?')
  assert.equal(await dialogs.showPrompt('prompts.unsaved', { name: 'Budget.xlsx' }), 'cancel', 'Esc answers Cancel')
  assert.equal(await dialogs.showPrompt('prompts.unsaved-untitled', {}), 'cancel')
  assert.equal(lastBox().message, 'Do you want to save this new document?')
  assert.deepEqual(lastBox().buttons, ['Save…', 'Don\'t Save', 'Cancel'])

  electron.queue.box.push('Try Again')
  assert.equal(await dialogs.showPrompt('saveFailed.LOCKED', { name: 'Budget.xlsx', technical: 'EBUSY rename after 8 attempts' }), 'retry')
  assert.equal(lastBox().type, 'warning')
  assert.match(lastBox().detail, /\n\nDetails: EBUSY rename after 8 attempts$/)
  const controller = new AbortController()
  electron.queue.box.push((options) => new Promise((resolve) => options.signal.addEventListener('abort', () => resolve(options.cancelId))))
  const pending = dialogs.showPrompt('prompts.not-responding', {}, { signal: controller.signal })
  controller.abort()
  assert.equal(await pending, 'wait', 'a prompt closed by its signal answers with its cancel button')
  await assert.rejects(() => dialogs.showPrompt('status.saving'), TypeError)
})

test('the Open dialog lists what the workspace opens first and All files last, and remembers the folder', async () => {
  const layout = scenario('open')
  const a = path.join(layout.work, 'a.xlsx')
  const b = path.join(layout.work, 'b.csv')
  electron.queue.open.push([a, b])
  assert.deepEqual(await dialogs.chooseOpenPaths({ multi: true, purpose: 'open' }), [a, b])
  const options = electron.record.open[electron.record.open.length - 1]
  assert.equal(options.filters[0].name, 'All spreadsheets')
  assert.ok(options.filters[0].extensions.includes('xlsx') && options.filters[0].extensions.includes('csv'))
  assert.deepEqual(options.filters[options.filters.length - 1], { name: 'All files', extensions: ['*'] })
  assert.deepEqual(options.properties, ['openFile', 'multiSelections'])
  assert.equal(options.defaultPath, layout.documents)
  assert.equal(readPrefs(layout).lastOpenFolder, layout.work)
  electron.queue.open.push([a, b])
  assert.deepEqual(await dialogs.chooseOpenPaths({ multi: false }), [a], 'a single-file Open returns one path')
  assert.deepEqual(electron.record.open[electron.record.open.length - 1].properties, ['openFile'])
  assert.equal(electron.record.open[electron.record.open.length - 1].defaultPath, layout.work)
  assert.deepEqual(await dialogs.chooseOpenPaths({ multi: true }), [], 'canceled')
})

test('preferences are validated and written through safe-write', async () => {
  const layout = scenario('prefs')
  assert.equal(await dialogs.getPref('exportFormat'), undefined)
  assert.equal(await dialogs.setPref('exportFormat', 'pdf'), true)
  assert.equal(await dialogs.setPref('openAfterExport', true), true)
  assert.equal(await dialogs.getPref('exportFormat'), 'pdf')
  assert.deepEqual(readPrefs(layout), { exportFormat: 'pdf', openAfterExport: true })
  await dialogs.setPref('openAfterExport', undefined)
  assert.deepEqual(readPrefs(layout), { exportFormat: 'pdf' })
  await assert.rejects(() => dialogs.setPref('__proto__', 1), TypeError)
  await assert.rejects(() => dialogs.setPref('bad key', 1), TypeError)
  await assert.rejects(() => dialogs.setPref('lastSaveFolder', 'relative\\folder'), TypeError)
  await assert.rejects(() => dialogs.setPref('big', 'x'.repeat(70 * 1024)), TypeError)
  assert.deepEqual(fs.readdirSync(layout.userData).filter((name) => name.startsWith('~simple-')), [], 'no temp files are left')
})

test('SIMPLE_QA_DIALOGS answers dialogs in unpackaged runs and is ignored when packaged', async (t) => {
  const layout = scenario('qa')
  const script = path.join(layout.folder, 'qa.json')
  const log = path.join(layout.folder, 'qa-log.jsonl')
  const opened = path.join(layout.work, 'opened.xlsx')
  fs.writeFileSync(script, JSON.stringify({
    save: ['default', { name: 'Other.xlsx' }, 'cancel'],
    open: [[opened]],
    prompts: { 'prompts.unsaved': ['dont-save'], '*': ['cancel'] },
    log,
  }))
  const previous = process.env.SIMPLE_QA_DIALOGS
  process.env.SIMPLE_QA_DIALOGS = script
  t.after(() => {
    if (previous === undefined) delete process.env.SIMPLE_QA_DIALOGS
    else process.env.SIMPLE_QA_DIALOGS = previous
    electron.app.isPackaged = false
  })
  const dialogsBefore = electron.record.save.length + electron.record.open.length + electron.record.box.length
  assert.deepEqual(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Budget.xlsx', format: 'xlsx' }), { path: path.join(layout.documents, 'Budget.xlsx'), format: 'xlsx' })
  assert.deepEqual(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Budget.xlsx', format: 'xlsx' }), { path: path.join(layout.documents, 'Other.xlsx'), format: 'xlsx' })
  assert.equal(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Budget.xlsx', format: 'xlsx' }), null)
  assert.equal(await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Budget.xlsx', format: 'xlsx' }), null, 'an exhausted list cancels')
  assert.deepEqual(await dialogs.chooseOpenPaths({ multi: true }), [opened])
  assert.equal(await dialogs.showPrompt('prompts.unsaved', { name: 'x' }), 'dont-save')
  assert.equal(await dialogs.showPrompt('prompts.crashed', {}), 'close', 'an answer that is not a button becomes the cancel button')
  assert.equal(electron.record.save.length + electron.record.open.length + electron.record.box.length, dialogsBefore, 'no real dialog was shown')
  const lines = fs.readFileSync(log, 'utf8').trim().split('\n').map((line) => JSON.parse(line))
  assert.deepEqual(lines.map((line) => line.type), ['save', 'save', 'save', 'save', 'open', 'prompt', 'prompt'])
  assert.equal(lines[5].key, 'prompts.unsaved')

  electron.app.isPackaged = true
  electron.queue.save.push(null)
  await dialogs.chooseSavePath({ purpose: 'save-as', name: 'Budget.xlsx', format: 'xlsx' })
  assert.equal(electron.record.save.length + electron.record.open.length + electron.record.box.length, dialogsBefore + 1, 'a packaged app always shows the real dialog')
})

// ---------------------------------------------------------------------------
// IPC handlers and the preload bridge
// ---------------------------------------------------------------------------

function register(module, extra = {}) {
  const ipcMain = makeIpcMain()
  const launches = []
  const registration = ioIpc.registerSharedIo({ ipcMain, module, launch: (command, args) => launches.push({ command, args }), ...extra })
  return { ipcMain, launches, registration }
}

function eventFrom(win, overrides = {}) {
  return { sender: win.webContents, senderFrame: win.webContents.mainFrame, ...overrides }
}

test('the bridge exposes exactly the shared contract and knows its workspace', async (t) => {
  scenario('bridge', 'video')
  const { ipcMain, registration } = register('video')
  t.after(() => registration.dispose())
  const win = new electron.BrowserWindow()
  const io = loadBridge(win, ipcMain, { argv: ['--simple-io-module=video'] })
  assert.deepEqual(Object.keys(io).sort(), [
    'capabilities', 'chooseOpenPaths', 'chooseSavePath', 'clipboard', 'module', 'onCapabilitiesChanged', 'onRequest', 'openInSimple',
    'officeEngine', 'pathForFile', 'prefs', 'prompt', 'recovery', 'shell', 'version', 'versions',
  ].sort())
  assert.equal(io.version, 1)
  assert.equal(io.module, 'video')
  assert.deepEqual(Object.keys(io.shell).sort(), ['openPath', 'showItem'])
  assert.deepEqual(Object.keys(io.officeEngine), ['status'], 'there is no install flow')
  assert.equal(loadBridge(new electron.BrowserWindow(), ipcMain, { env: { SIMPLE_IO_MODULE: 'calc' } }).module, 'calc')
  assert.equal(io.pathForFile({ path: 'C:\\Drops\\a.mp4' }), 'C:\\Drops\\a.mp4')
  assert.equal(io.pathForFile({}), null)
  const caps = await io.capabilities()
  assert.equal(caps.module, 'video')
  assert.equal(caps.version, 1)
  assert.equal(caps.officeEngine.available, false, 'a workspace without the engine reports none')
  assert.deepEqual(await io.recovery.list(), [], 'Video keeps no recovery copies')
  assert.equal(typeof io.onRequest('not-a-request', () => true), 'function')
  electron.clipboard.content = { text: 'hello', html: '<b>hello</b>', png: Buffer.from([1, 2, 3]), file: 'C:\\Copied\\report.docx' }
  const pasted = await io.clipboard.read()
  assert.equal(pasted.text, 'hello')
  assert.equal(pasted.html, '<b>hello</b>')
  assert.deepEqual([...pasted.png], [1, 2, 3])
  assert.deepEqual(pasted.files, process.platform === 'win32' ? [path.resolve('C:\\Copied\\report.docx')] : [])
  electron.clipboard.content = {}
})

test('every io handler refuses requests that do not come from the app\'s own page', async (t) => {
  scenario('trust')
  const { ipcMain, registration } = register('calc')
  t.after(() => registration.dispose())
  assert.deepEqual([...ipcMain.handlers.keys()].sort(), [...ioIpc.CHANNELS].sort())
  const win = new electron.BrowserWindow()
  const ok = await ipcMain.handlers.get('io:prefs-get')(eventFrom(win), 'exportFormat')
  assert.equal(ok, undefined)
  const untrusted = [
    eventFrom(win, { senderFrame: { url: win.webContents.mainFrame.url } }),
    eventFrom(win, { senderFrame: null }),
    eventFrom(new electron.BrowserWindow({ url: 'https://example.invalid/index.html' })),
    eventFrom(new electron.BrowserWindow({ url: pathToFileURL(path.join(ROOT, 'elsewhere', 'print.html')).href })),
  ]
  for (const event of untrusted) {
    for (const channel of ioIpc.CHANNELS) {
      assert.throws(() => ipcMain.handlers.get(channel)(event), /Untrusted request/, channel)
    }
  }
  assert.equal(ioIpc.isTrustedSender(eventFrom(new electron.BrowserWindow({ url: 'http://localhost:5173/' })), { devServerUrl: 'http://localhost:5173' }), true)
  assert.equal(ioIpc.isTrustedSender(eventFrom(new electron.BrowserWindow({ url: 'http://localhost:9999/' })), { devServerUrl: 'http://localhost:5173' }), false)
  assert.throws(() => ioIpc.registerSharedIo({ ipcMain, module: 'calc' }), /already/)
  assert.throws(() => ioIpc.registerSharedIo({ ipcMain: makeIpcMain(), module: 'word' }), TypeError)
})

test('the prompt handler shows only catalog decision prompts', async (t) => {
  scenario('ipc-prompt')
  const { ipcMain, registration } = register('calc')
  t.after(() => registration.dispose())
  const win = new electron.BrowserWindow()
  const io = loadBridge(win, ipcMain)
  electron.queue.box.push('Save As…')
  assert.equal(await io.prompt('saveFailed.READ_ONLY', { name: 'Budget.xlsx', technical: 'read-only attribute', evil: { nested: true } }), 'save-as')
  assert.equal(lastBox().message, '"Budget.xlsx" is read-only.')
  await assert.rejects(() => io.prompt('status.saving', {}), /Unknown prompt/)
  await assert.rejects(() => io.prompt('../../package', {}), /Unknown prompt/)

  // Answers that let a save skip a check are recorded for this window and file only.
  const file = path.join(ROOT, 'decided', 'Budget.xlsx')
  const other = new electron.BrowserWindow()
  assert.equal(dialogs.hasDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file), false)
  electron.queue.box.push('Cancel')
  assert.equal(await io.prompt('saveFailed.CHANGED_ON_DISK', { name: 'Budget.xlsx', path: file }), 'cancel')
  assert.equal(dialogs.hasDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file), false, 'Cancel is not Replace')
  electron.queue.box.push('Replace')
  assert.equal(await io.prompt('saveFailed.CHANGED_ON_DISK', { name: 'Budget.xlsx', path: file }), 'replace')
  assert.equal(dialogs.hasDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file), true)
  assert.equal(dialogs.hasDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', path.join(ROOT, 'decided', 'Other.xlsx')), false)
  assert.equal(dialogs.hasDecision(other.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file), false, 'another window never inherits it')
  dialogs.clearDecisions(win.webContents, file)
  assert.equal(dialogs.hasDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file), false)
})

test('choosing a save path through IPC grants it to that window only', async (t) => {
  const layout = scenario('grant')
  const { ipcMain, registration } = register('calc', { guard })
  t.after(() => registration.dispose())
  const win = new electron.BrowserWindow()
  const other = new electron.BrowserWindow()
  const io = loadBridge(win, ipcMain)
  const source = path.join(layout.work, 'Budget.xlsx')
  fs.writeFileSync(source, 'x')
  const doc = guard.registerDocument({ webContents: win.webContents, path: source, stamp: { size: 1, mtimeMs: 1 } })
  electron.queue.save.push((options) => path.join(path.dirname(options.defaultPath), 'Budget copy.xlsx'))
  const chosen = await io.chooseSavePath({ purpose: 'save-copy', docId: doc.docId, name: 'Budget.xlsx', format: 'xlsx', formats: ['xlsx', 'csv'] })
  assert.deepEqual(chosen, { path: path.join(layout.work, 'Budget copy.xlsx'), format: 'xlsx' })
  assert.equal(lastSave().defaultPath, source, 'the document\'s folder comes from main\'s registry')
  assert.equal(dialogs.isGranted(win.webContents, chosen.path), true)
  assert.equal(dialogs.isGranted(other.webContents, chosen.path), false)
  // Another window's document id does not reveal its folder.
  electron.queue.save.push(null)
  await loadBridge(other, ipcMain).chooseSavePath({ purpose: 'save-as', docId: doc.docId, name: 'Mine.xlsx', format: 'xlsx', formats: ['xlsx'] })
  assert.equal(lastSave().defaultPath, path.join(layout.work, 'Mine.xlsx'), 'it starts in the remembered folder instead')
  guard.forgetDocument(doc.docId)
})

test('open in Simple uses the unified app\'s own routing and never the system\'s default app', async (t) => {
  const layout = scenario('open-in-simple', 'pdf')
  const sheet = path.join(layout.work, 'Budget.xlsx')
  const odd = path.join(layout.work, 'archive.xyz')
  fs.writeFileSync(sheet, 'PK')
  fs.writeFileSync(odd, 'x')
  const { ipcMain, launches, registration } = register('pdf', { unified: true })
  t.after(() => registration.dispose())
  const win = new electron.BrowserWindow()
  const io = loadBridge(win, ipcMain)
  const shownBefore = electron.record.shown.length

  const launched = await io.openInSimple(sheet)
  assert.equal(launched.ok, true)
  assert.equal(launched.action, 'launched')
  assert.equal(launched.mode, 'calc')
  assert.equal(launched.appName, 'Simple Spreadsheets')
  assert.equal(launched.shownInFolder, false)
  assert.deepEqual(launches.at(-1), { command: process.execPath, args: [APP_PATH, sheet] }, 'unpackaged: electron with the app folder and the file')
  assert.equal((await io.shell.openPath(sheet)).action, 'launched', 'shell.openPath also opens inside Simple')
  assert.equal(launches.length, 2)

  const odder = await io.openInSimple(odd)
  assert.equal(odder.action, 'shown-in-folder')
  assert.equal(odder.shownInFolder, true)
  assert.equal(odder.reason, 'unsupported')
  assert.deepEqual(electron.record.shown.slice(shownBefore), [odd])
  assert.equal(launches.length, 2, 'files no workspace opens are never launched')

  const missing = await io.openInSimple(path.join(layout.work, 'gone.xlsx'))
  assert.equal(missing.ok, false)
  assert.equal(missing.code, 'NOT_FOUND')
  await assert.rejects(() => io.openInSimple('relative\\Budget.xlsx'), /absolute/)
  await assert.rejects(() => io.openInSimple('\\\\.\\PhysicalDrive0'), process.platform === 'win32' ? /not supported/ : /absolute/)
  const folder = await io.openInSimple(layout.work)
  assert.equal(folder.ok, false)

  electron.app.isPackaged = true
  try {
    assert.deepEqual(ioIpc.selfLaunchCommand([sheet]), { command: process.execPath, args: [sheet] }, 'packaged: only the file')
  } finally {
    electron.app.isPackaged = false
  }
  const portable = path.join(layout.folder, 'simple.exe')
  fs.writeFileSync(portable, 'MZ')
  const previous = process.env.PORTABLE_EXECUTABLE_FILE
  process.env.PORTABLE_EXECUTABLE_FILE = portable
  try {
    assert.deepEqual(ioIpc.selfLaunchCommand([sheet]), { command: portable, args: [sheet] }, 'the portable EXE relaunches itself')
    process.env.PORTABLE_EXECUTABLE_FILE = path.join(layout.folder, 'not-there.exe')
    assert.equal(ioIpc.selfLaunchCommand([sheet]).command, process.execPath)
  } finally {
    if (previous === undefined) delete process.env.PORTABLE_EXECUTABLE_FILE
    else process.env.PORTABLE_EXECUTABLE_FILE = previous
  }
  assert.throws(() => ioIpc.selfLaunchCommand(['--inspect=9229']), /absolute/, 'flags are never passed')
  assert.deepEqual(electron.record.openPath, [], 'shell.openPath was never called')
  assert.deepEqual(electron.record.openExternal, [])
})

test('a standalone build opens its own files in a new window and shows other files in their folder', async (t) => {
  const layout = scenario('standalone', 'calc')
  const sheet = path.join(layout.work, 'Budget.xlsx')
  const pdf = path.join(layout.work, 'Scan.pdf')
  fs.writeFileSync(sheet, 'PK')
  fs.writeFileSync(pdf, '%PDF-1.7')
  const opened = []
  const { ipcMain, launches, registration } = register('calc', { unified: false, guard, openInWindow: (filePath, info) => { opened.push({ filePath, mode: info.mode }) } })
  t.after(() => registration.dispose())
  const win = new electron.BrowserWindow()
  const io = loadBridge(win, ipcMain)
  assert.equal((await io.openInSimple(sheet)).action, 'opened')
  assert.deepEqual(opened, [{ filePath: sheet, mode: 'calc' }])
  const other = await io.openInSimple(pdf)
  assert.equal(other.action, 'shown-in-folder')
  assert.equal(other.reason, 'other-workspace')
  assert.equal(other.mode, 'pdf')
  assert.equal(launches.length, 0, 'a standalone build never starts another program')

  // A file already open in a window of this process is focused instead.
  const holder = new electron.BrowserWindow()
  const doc = guard.registerDocument({ webContents: holder.webContents, path: sheet, stamp: { size: 2, mtimeMs: 1 } })
  const focused = await io.openInSimple(sheet)
  assert.equal(focused.action, 'focused')
  assert.equal(holder.focusCount, 1)
  assert.equal(opened.length, 1)
  guard.forgetDocument(doc.docId)
  assert.deepEqual(electron.record.openPath, [])
})

test('recovery through the bridge: main supplies the file and its stamp, and a window cannot touch another window\'s entry', async (t) => {
  const layout = scenario('ipc-recovery')
  const opened = []
  const { ipcMain, registration } = register('calc', { guard, stores, openVersion: (version, info) => { opened.push({ version, sender: info.sender.id }) } })
  t.after(() => registration.dispose())
  const first = new electron.BrowserWindow()
  const second = new electron.BrowserWindow()
  const io1 = loadBridge(first, ipcMain)
  const io2 = loadBridge(second, ipcMain)
  const file = path.join(layout.work, 'Budget.csv')
  const helper = path.join(layout.work, 'lookup.csv')
  fs.writeFileSync(file, 'a,b\n1,2\n')
  fs.writeFileSync(helper, 'x\n')
  const { docId, stamp } = await guard.openDocument(file, { webContents: first.webContents })

  const written = await io1.recovery.write({
    docId,
    revision: 4,
    title: 'Budget.csv',
    sourcePath: 'C:\Somewhere\else.csv',
    meta: { sheet: 2 },
    parts: [
      { name: 'workbook.json', data: '{"cells":{"B2":3}}', compress: true },
      { name: 'base', ref: 'source' },
    ],
  })
  assert.equal(written.ok, true, JSON.stringify(written))
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'recovery', docId, 'manifest.json'), 'utf8'))
  assert.equal(manifest.sourcePath, file, 'the page cannot choose which file its edits belong to')
  assert.deepEqual(manifest.sourceStamp, stamp, 'the stamp is the one main took when it read the file')
  assert.deepEqual(manifest.parts.find((part) => part.name === 'base').ref, { path: file, stamp })
  assert.equal(manifest.module, 'calc')
  // A page may refer only to its own file: any other path would make main read and hash that file.
  await assert.rejects(io1.recovery.write({ docId, revision: 5, title: 'Budget.csv', parts: [{ name: 'lookup', ref: { path: helper } }] }), /own file/)

  const blocked = await io2.recovery.write({ docId, revision: 9, title: 'Budget.csv', parts: [{ name: 'workbook.json', data: '{}' }] })
  assert.equal(blocked.code, 'LOCKED')
  assert.equal(blocked.reason, 'open-in-another-window')
  assert.equal(blocked.message, '"Budget.csv" is open in another Simple window. Close it there first, or save your changes with a different name.')
  assert.equal((await io2.recovery.discard(docId)).code, 'LOCKED')
  assert.deepEqual(await io2.recovery.list(), [], 'a live entry is not offered')

  // The first window crashes: its entry is offered, and the window that restores it takes the document over.
  await stores.recoveryStore().markOrphaned([docId])
  guard.forgetDocument(docId)
  const listed = await io2.recovery.list()
  assert.deepEqual(listed.map((entry) => [entry.id, entry.title, entry.orphaned, entry.sourceChanged, entry.restorable]), [[docId, 'Budget.csv', true, false, true]])
  assert.equal(listed[0].folder, undefined, 'internal folders stay in main')
  const restored = await io2.recovery.read(docId)
  assert.equal(restored.ok, true)
  assert.deepEqual(restored.meta, { sheet: 2 })
  assert.deepEqual(restored.parts.map((part) => part.name), ['workbook.json', 'base'])
  assert.equal(restored.parts[0].data, '{"cells":{"B2":3}}')
  assert.equal(restored.parts[1].ref.matches, true)
  // Reading changes nothing: a restore that fails in the page leaves no document bound and the entry on offer.
  assert.equal(guard.getDocument(docId), null, 'nothing is adopted before the page restored the content')
  assert.equal(guard.findDocumentByPath(file), null)
  assert.deepEqual((await io2.recovery.list()).map((entry) => entry.id), [docId], 'the entry is still offered')
  assert.equal((await io1.recovery.adopt(docId)).code, 'NOT_FOUND', 'only the window that read the entry can take it over')
  const taken = await io2.recovery.adopt(docId)
  assert.deepEqual(taken, { ok: true, docId, untitled: false, path: file })
  const adopted = guard.getDocument(docId)
  assert.equal(adopted.webContentsId, second.webContents.id)
  assert.equal(adopted.path, file)
  assert.deepEqual(adopted.stamp, stamp, 'saving the restored document checks the file against the snapshot\'s stamp')
  assert.deepEqual(await io2.recovery.list(), [], 'a restored entry is not offered again')
  assert.deepEqual(await io2.recovery.discard(docId, 4), { ok: true, removed: true })

  // Earlier versions open through the workspace's hook, never as the stored file itself.
  await stores.versionsStore().backup(file)
  const versions = await io2.versions.list(file)
  assert.equal(versions.length, 1)
  assert.deepEqual(await io2.versions.open(versions[0].id), { ok: true, name: 'Budget.csv', createdAt: versions[0].createdAt })
  assert.equal(opened.length, 1)
  assert.equal(opened[0].sender, second.webContents.id)
  assert.equal(fs.readFileSync(opened[0].version.path, 'utf8'), 'a,b\n1,2\n')
  assert.equal((await io2.versions.open(`${'f'.repeat(40)}/../../escape`)).code, 'NOT_FOUND')
  guard.forgetDocument(docId)
})

test('registering the Documents workspace moves its old recovery copies into the journal and prunes the stores', async (t) => {
  scenario('maintenance', 'docs')
  const legacyIndex = path.join(ROOT, 'recoveries.json')
  const legacyFile = path.join(ROOT, 'recovery', 'legacy-1.docx')
  fs.mkdirSync(path.dirname(legacyFile), { recursive: true })
  fs.writeFileSync(legacyFile, 'PK\u0003\u0004 legacy')
  fs.writeFileSync(legacyIndex, JSON.stringify([{ id: 'legacy-1', title: 'Letter', sourcePath: null, updatedAt: Date.now() }]))
  const { registration } = register('docs', { guard, stores })
  t.after(() => registration.dispose())
  await registration.maintenance
  assert.equal(fs.existsSync(legacyIndex), false)
  assert.equal(fs.existsSync(legacyFile), false)
  const entry = (await stores.recoveryStore().list()).find((item) => item.id === 'legacy-1')
  assert.equal(entry.title, 'Letter.docx')
  await stores.recoveryStore().discard('legacy-1')
  const { registration: video } = register('video', { maintenance: true })
  await video.maintenance
  video.dispose()
})
