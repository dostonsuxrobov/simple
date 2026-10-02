'use strict'

// The window and app guards of the shared Save layer (document-guard.cjs,
// design §3.7, §10.1): closing, quitting, crashed and hung pages, and Windows
// shutdown. `electron` is a stub: windows emit the same events Electron does,
// message boxes answer from a queue, and each window runs the real preload
// bridge (simple/shared/preload/io-bridge.cjs) with a fake page session that
// answers close-query, save-now, discard and recovery-flush.

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
const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-window-guard-'))
const APP_PATH = path.join(ROOT, 'app')

// ---------------------------------------------------------------------------
// Electron stub
// ---------------------------------------------------------------------------

const boxes = []
const answers = []
const windows = new Set()
let nextContentsId = 1

class FakeWebContents extends EventEmitter {
  constructor() {
    super()
    this.id = nextContentsId++
    this.destroyed = false
    this.crashed = false
    this.reloads = 0
    this.forcedCrashes = 0
    this.mainFrame = { url: pathToFileURL(path.join(APP_PATH, 'dist', 'index.html')).href }
    this.renderer = null
  }
  send(channel, ...args) {
    if (this.destroyed) throw new Error('Object has been destroyed')
    if (this.renderer && !this.crashed) this.renderer(channel, ...args)
  }
  isDestroyed() { return this.destroyed }
  isCrashed() { return this.crashed }
  reload() {
    this.reloads += 1
    this.crashed = false
  }
  forcefullyCrashRenderer() {
    this.forcedCrashes += 1
    this.crashed = true
  }
  destroy() {
    if (this.destroyed) return
    this.destroyed = true
    this.emit('destroyed')
  }
}

class BrowserWindow extends EventEmitter {
  constructor() {
    super()
    this.webContents = new FakeWebContents()
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
  /** Like Electron: 'close' can be prevented; otherwise the window is destroyed. Returns the event. */
  close() {
    const event = { defaultPrevented: false, preventDefault() { this.defaultPrevented = true } }
    if (this.destroyed) return event
    this.emit('close', event)
    if (!event.defaultPrevented) this.destroy()
    return event
  }
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

function makeApp() {
  const app = new EventEmitter()
  app.quitCount = 0
  app.isPackaged = false
  app.getAppPath = () => APP_PATH
  app.getPath = () => path.join(ROOT, 'userData')
  app.quit = () => {
    const event = { defaultPrevented: false, preventDefault() { this.defaultPrevented = true } }
    app.emit('before-quit', event)
    if (event.defaultPrevented) return
    for (const win of [...windows]) {
      // As in Electron, a window that refuses to close stops the quit.
      if (!win.close().defaultPrevented) continue
      return
    }
    app.quitCount += 1
    app.emit('quit')
  }
  return app
}

const ipcMain = Object.assign(new EventEmitter(), { handle() {}, removeHandler() {} })
const electronStub = {
  app: makeApp(),
  BrowserWindow,
  ipcMain,
  dialog: {
    async showMessageBox(win, options) {
      const opts = options || win
      boxes.push(opts)
      let answer = answers.length ? answers.shift() : undefined
      if (typeof answer === 'function') answer = await answer(opts)
      if (answer === undefined) answer = opts.cancelId
      const index = typeof answer === 'number' ? answer : opts.buttons.indexOf(answer)
      if (index < 0) throw new Error(`No button "${answer}" in ${opts.buttons.join(', ')}`)
      return { response: index, checkboxChecked: false }
    },
    async showSaveDialog() { return { canceled: true, filePath: '' } },
    async showOpenDialog() { return { canceled: true, filePaths: [] } },
  },
  shell: { showItemInFolder() {}, openPath() { throw new Error('documents never go to another app') } },
}
const originalResolve = Module._resolveFilename
Module._resolveFilename = function resolveElectronStub(request, ...rest) {
  if (request === 'electron') return 'electron-test-stub'
  return originalResolve.call(this, request, ...rest)
}
require.cache['electron-test-stub'] = { id: 'electron-test-stub', filename: 'electron-test-stub', loaded: true, exports: electronStub }

const core = require(path.join(SHARED, 'electron', 'io-core.cjs'))
const stores = require(path.join(SHARED, 'electron', 'stores.cjs'))
const dialogs = require(path.join(SHARED, 'electron', 'io-dialogs.cjs'))
const guard = require(path.join(SHARED, 'electron', 'document-guard.cjs'))

const quiet = { warn() {}, info() {} }
core.configureIo({ journalDir: path.join(ROOT, 'journal'), logger: quiet })
stores.configureStores({ recoveryDir: path.join(ROOT, 'recovery'), versionsDir: path.join(ROOT, 'versions'), logger: quiet })
dialogs.configureDialogs({ module: 'calc', userData: path.join(ROOT, 'userData'), logger: quiet })
guard.configureDocumentGuard({ module: 'calc', logger: quiet })

test.after(() => {
  Module._resolveFilename = originalResolve
  fs.rmSync(ROOT, { recursive: true, force: true })
})
test.afterEach(() => {
  answers.length = 0
  for (const win of [...windows]) win.destroy()
})

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const TIMING = { queryTimeoutMs: 150, pollMs: 10, discardTimeoutMs: 300, flushTimeoutMs: 120 }

/** Runs the real preload bridge for a window, wired to the stub's ipcMain. */
function loadBridge(win) {
  const source = fs.readFileSync(path.join(SHARED, 'preload', 'io-bridge.cjs'), 'utf8')
  const listeners = new Map()
  const ipcRenderer = {
    on(channel, listener) {
      if (!listeners.has(channel)) listeners.set(channel, [])
      listeners.get(channel).push(listener)
    },
    send(channel, ...args) { ipcMain.emit(channel, { sender: win.webContents, senderFrame: win.webContents.mainFrame }, ...structuredClone(args)) },
    invoke() { return Promise.reject(new Error('not used here')) },
  }
  win.webContents.renderer = (channel, ...args) => {
    for (const listener of listeners.get(channel) || []) listener({}, ...structuredClone(args))
  }
  let api = null
  vm.runInNewContext(source, {
    require: () => ({ contextBridge: { exposeInMainWorld: (_name, value) => { api = value } }, ipcRenderer, webUtils: { getPathForFile: () => '' } }),
    process: { argv: ['--simple-io-module=calc'], env: {} },
  })
  return api
}

/**
 * A guarded window whose page has a fake session. `page.status` answers
 * close-query; `page.onSave` answers save-now (default: saves and becomes clean).
 */
function openWindow({ session = true, status = { dirty: false, saving: false, title: 'Budget.xlsx', kind: 'spreadsheet' } } = {}) {
  const win = new BrowserWindow()
  const api = guard.installWindowGuard(win, { module: 'calc', ...TIMING })
  const page = { status: { ...status }, calls: [], onSave: null, onQuery: null }
  if (session) {
    const io = loadBridge(win)
    io.onRequest('close-query', (payload) => {
      page.calls.push(['close-query', payload])
      return page.onQuery ? page.onQuery(payload) : page.status
    })
    io.onRequest('save-now', async (payload) => {
      page.calls.push(['save-now', payload])
      if (page.onSave) return page.onSave(payload)
      page.status = { ...page.status, dirty: false, lossy: null }
      return true
    })
    io.onRequest('discard', () => {
      page.calls.push(['discard'])
      return true
    })
    io.onRequest('recovery-flush', (payload) => {
      page.calls.push(['recovery-flush', payload])
      return { ok: true }
    })
  } else {
    loadBridge(win)
  }
  return { win, api, page }
}

async function until(check, label, timeoutMs = 4000) {
  const started = Date.now()
  while (!check()) {
    if (Date.now() - started > timeoutMs) throw new Error(`Timed out waiting for ${label}`)
    await new Promise((resolve) => setTimeout(resolve, 5))
  }
}

function callsOf(page, type) {
  return page.calls.filter((call) => call[0] === type)
}

// ---------------------------------------------------------------------------
// Closing
// ---------------------------------------------------------------------------

test('a dirty window asks Save, Don\'t Save or Cancel, with Save as the default and Esc as Cancel', async () => {
  const { win, api, page } = openWindow({ status: { dirty: true, saving: false, title: 'Budget.xlsx' } })
  assert.equal(electronStub.app.listenerCount('before-quit'), 1, 'the first guarded window guards quitting too')
  openWindow()
  assert.equal(electronStub.app.listenerCount('before-quit'), 1, 'once per process')
  const event = win.close()
  assert.equal(event.defaultPrevented, true, 'closing waits for the decision')
  const outcome = await api.requestClose()
  assert.equal(outcome, 'canceled', 'no answer is Esc, which is Cancel')
  const box = boxes.at(-1)
  assert.deepEqual(box.buttons, ['Save', 'Don\'t Save', 'Cancel'])
  assert.equal(box.defaultId, 0)
  assert.equal(box.cancelId, 2)
  assert.equal(box.noLink, true)
  assert.equal(box.message, 'Do you want to save the changes to "Budget.xlsx"?')
  assert.equal(box.detail, 'Your changes will be lost if you don\'t save them.')
  assert.equal(win.isDestroyed(), false)
  assert.equal(callsOf(page, 'save-now').length + callsOf(page, 'discard').length, 0, 'Cancel does nothing')
})

test('Save closes the window only after the page has saved, and a failed save keeps it open', async () => {
  const { win, page } = openWindow({ status: { dirty: true, saving: false, title: 'Budget.xlsx' } })
  let finishSave
  page.onSave = () => new Promise((resolve) => { finishSave = resolve })
  answers.push('Save')
  win.close()
  await until(() => callsOf(page, 'save-now').length === 1, 'save-now')
  assert.deepEqual(callsOf(page, 'save-now')[0][1], { reason: 'close' })
  await new Promise((resolve) => setTimeout(resolve, 50))
  assert.equal(win.isDestroyed(), false, 'still open while the save runs')
  page.status = { ...page.status, dirty: false }
  finishSave(true)
  await until(() => win.isDestroyed(), 'the window to close')

  const second = openWindow({ status: { dirty: true, saving: false, title: 'Report.xlsx' } })
  second.page.onSave = () => false
  answers.push('Save')
  assert.equal(await second.api.requestClose(), 'canceled')
  assert.equal(second.win.isDestroyed(), false, 'a save that did not happen never closes the window')
})

test('Don\'t Save drops the page\'s recovery copy and closes', async () => {
  const { win, page } = openWindow({ status: { dirty: true, saving: false, title: 'Budget.xlsx' } })
  answers.push('Don\'t Save')
  win.close()
  await until(() => win.isDestroyed(), 'the window to close')
  assert.equal(callsOf(page, 'discard').length, 1)
  assert.equal(callsOf(page, 'save-now').length, 0)
})

test('a clean window, or a page without a document session, closes without asking', async () => {
  const before = boxes.length
  const clean = openWindow({ status: { dirty: false, saving: false, title: 'Budget.xlsx' } })
  clean.win.close()
  await until(() => clean.win.isDestroyed(), 'the clean window to close')
  const bare = openWindow({ session: false })
  bare.win.close()
  await until(() => bare.win.isDestroyed(), 'the window without a session to close')
  assert.equal(boxes.length, before, 'no prompt was shown')

  const approved = openWindow()
  const unload = { prevented: false, preventDefault() { this.prevented = true } }
  approved.win.webContents.emit('will-prevent-unload', unload)
  assert.equal(unload.prevented, false, 'a page may still keep itself open before closing was approved')
  assert.equal(await approved.api.requestClose(), 'approved')
  approved.win.webContents.emit('will-prevent-unload', unload)
  assert.equal(unload.prevented, true, 'once approved, a beforeunload handler cannot undo the close')
})

test('an untitled document asks whether to save the new spreadsheet', async () => {
  const { api } = openWindow({ status: { dirty: true, saving: false, title: 'Untitled spreadsheet', kind: 'spreadsheet', docId: 'new-sheet-1' } })
  answers.push('Cancel')
  assert.equal(await api.requestClose(), 'canceled')
  assert.equal(boxes.at(-1).message, 'Do you want to save this new spreadsheet?')
  assert.deepEqual(boxes.at(-1).buttons, ['Save…', 'Don\'t Save', 'Cancel'])
})

test('a window is never closed while a save is running', async () => {
  const { win, page } = openWindow({ status: { dirty: true, saving: true, title: 'Budget.xlsx' } })
  let queries = 0
  page.onQuery = () => {
    queries += 1
    return queries < 3 ? { dirty: true, saving: true, title: 'Budget.xlsx' } : { dirty: false, saving: false, title: 'Budget.xlsx' }
  }
  const before = boxes.length
  win.close()
  await until(() => win.isDestroyed(), 'the window to close after the save')
  assert.equal(queries, 3)
  assert.equal(boxes.length, before, 'no prompt while saving')

  // A save main is still running for the window holds the close too.
  const second = openWindow({ status: { dirty: false, saving: false, title: 'Ledger.csv' } })
  const folder = path.join(ROOT, 'save-in-flight')
  fs.mkdirSync(folder, { recursive: true })
  const file = path.join(folder, 'Ledger.csv')
  fs.writeFileSync(file, 'a\n')
  const { docId } = await guard.openDocument(file, { webContents: second.win.webContents })
  let release
  const gate = new Promise((resolve) => { release = resolve })
  const saving = guard.performSave({ sender: second.win.webContents }, { docId, mode: 'save', revision: 1 }, { serialize: async () => { await gate; return Buffer.from('b\n') } })
  second.win.close()
  await new Promise((resolve) => setTimeout(resolve, 60))
  assert.equal(callsOf(second.page, 'close-query').length, 0, 'the page is asked only after the save finished')
  assert.equal(second.win.isDestroyed(), false)
  release()
  assert.equal((await saving).ok, true)
  await until(() => second.win.isDestroyed(), 'the window to close after main\'s save')
  assert.equal(fs.readFileSync(file, 'utf8'), 'b\n')
})

test('a lossy document offers a full copy before closing', async () => {
  const lossy = { format: 'csv', formatLabel: 'Comma-separated values (.csv)', lost: ['formulas', 'formatting'], fullFormat: 'xlsx', fullLabel: 'Excel workbook (.xlsx)' }
  const { win, page } = openWindow({ status: { dirty: false, saving: false, title: 'Budget.csv', lossy } })
  answers.push('Save Full Copy…')
  win.close()
  await until(() => win.isDestroyed(), 'the window to close')
  const box = boxes.find((item) => item.message.startsWith('"Budget.csv" was saved as'))
  assert.equal(box.message, '"Budget.csv" was saved as Comma-separated values (.csv), which doesn\'t keep formulas and formatting.')
  assert.equal(box.detail, 'Save a full copy as Excel workbook (.xlsx) before closing?')
  assert.deepEqual(box.buttons, ['Save Full Copy…', 'Close Anyway', 'Cancel'])
  assert.equal(box.defaultId, 0)
  assert.deepEqual(callsOf(page, 'save-now')[0][1], { reason: 'close', mode: 'save-copy', format: 'xlsx', fullCopy: true })
})

// ---------------------------------------------------------------------------
// Pages that do not answer, crash or hang
// ---------------------------------------------------------------------------

test('a page that does not answer gets the not-responding prompt, and Close Window keeps its recovery copy', async () => {
  const { win, page } = openWindow()
  page.onQuery = () => new Promise(() => {})
  const doc = guard.registerDocument({ webContents: win.webContents, path: null, name: 'Untitled spreadsheet' })
  assert.equal((await stores.recoveryStore().write({ docId: doc.docId, revision: 3, title: 'Untitled spreadsheet', parts: [{ name: 'workbook.json', data: '{}' }] })).ok, true)
  assert.deepEqual((await stores.recoveryStore().list()).map((entry) => entry.id), [], 'a live document is not offered')
  answers.push('Wait', 'Close Window')
  win.close()
  await until(() => win.isDestroyed(), 'the window to close')
  const prompts = boxes.filter((box) => box.message === 'Simple Spreadsheets isn\'t responding.')
  assert.equal(prompts.length, 2, 'Wait asked the page again')
  assert.deepEqual(prompts[0].buttons, ['Wait', 'Reopen and Recover', 'Close Window'])
  assert.equal(prompts[0].defaultId, 0)
  assert.equal(prompts[0].cancelId, 0)
  assert.match(prompts[0].detail, /Changes up to .+ will be offered the next time you open Simple Spreadsheets\./)
  await until(() => callsOf(page, 'close-query').length === 2, 'two close queries')
  const offered = await stores.recoveryStore().list()
  assert.deepEqual(offered.map((entry) => entry.id), [doc.docId], 'the closed window\'s work is offered for recovery')
  await stores.recoveryStore().discard(doc.docId)
})

test('a crashed page leaves a closable window, keeps its recovery copy and offers to reopen', async () => {
  const { win, page } = openWindow({ status: { dirty: true, saving: false, title: 'Budget.xlsx' } })
  const doc = guard.registerDocument({ webContents: win.webContents, path: null, name: 'Budget.xlsx' })
  await stores.recoveryStore().write({ docId: doc.docId, revision: 1, title: 'Budget.xlsx', parts: [{ name: 'workbook.json', data: '{}' }] })
  let reopen
  answers.push(() => new Promise((resolve) => { reopen = resolve }))
  win.webContents.crashed = true
  win.webContents.emit('render-process-gone', {}, { reason: 'crashed', exitCode: -1 })
  await until(() => boxes.at(-1) && boxes.at(-1).message === 'Simple Spreadsheets stopped unexpectedly.', 'the crash prompt')
  assert.deepEqual(boxes.at(-1).buttons, ['Reopen and Recover', 'Close'])
  assert.equal(boxes.at(-1).defaultId, 0)
  assert.equal(guard.getDocument(doc.docId), null, 'the crashed page\'s documents are released')
  assert.deepEqual((await stores.recoveryStore().list()).map((entry) => entry.id), [doc.docId])
  reopen('Reopen and Recover')
  await until(() => win.webContents.reloads === 1, 'the reload')
  assert.equal(win.isDestroyed(), false)
  assert.equal(callsOf(page, 'close-query').length, 0)

  const second = openWindow({ status: { dirty: true, saving: false, title: 'Other.xlsx' } })
  answers.push(() => new Promise(() => {}))
  second.win.webContents.crashed = true
  second.win.webContents.emit('render-process-gone', {}, { reason: 'oom', exitCode: -1 })
  await until(() => boxes.at(-1).message === 'Simple Spreadsheets stopped unexpectedly.', 'the second crash prompt')
  const event = second.win.close()
  assert.equal(event.defaultPrevented, false, 'a crashed window can always be closed')
  assert.equal(second.win.isDestroyed(), true)
  await stores.recoveryStore().discard(doc.docId)
})

test('a hung page gets the not-responding prompt, which closes itself when the page recovers', async () => {
  const { win } = openWindow()
  const start = boxes.length
  answers.push((options) => new Promise((resolve) => options.signal.addEventListener('abort', () => resolve(options.cancelId))))
  win.emit('unresponsive')
  await until(() => boxes.at(-1) && boxes.at(-1).message === 'Simple Spreadsheets isn\'t responding.' && boxes.at(-1).signal, 'the hang prompt')
  win.emit('unresponsive')
  win.emit('responsive')
  await new Promise((resolve) => setTimeout(resolve, 20))
  assert.equal(boxes.slice(start).filter((box) => box.message === 'Simple Spreadsheets isn\'t responding.').length, 1, 'one prompt at a time')
  assert.equal(win.webContents.forcedCrashes, 0)
  assert.equal(win.isDestroyed(), false)

  const before = boxes.length
  answers.push('Reopen and Recover')
  win.emit('unresponsive')
  await until(() => win.webContents.reloads === 1, 'the forced reload')
  assert.equal(win.webContents.forcedCrashes, 1)
  win.webContents.emit('render-process-gone', {}, { reason: 'killed', exitCode: -1 })
  await new Promise((resolve) => setTimeout(resolve, 20))
  assert.equal(boxes.length, before + 1, 'the expected crash shows no second prompt')
  assert.equal(win.isDestroyed(), false)
})

test('reopening a hung page while it was being closed keeps the window open', async () => {
  const { win, api, page } = openWindow({ status: { dirty: true, saving: false, title: 'Budget.xlsx' } })
  page.onQuery = () => new Promise(() => {})
  answers.push('Reopen and Recover')
  const flow = api.requestClose()
  await until(() => callsOf(page, 'close-query').length === 1, 'the close query')
  win.emit('unresponsive')
  await until(() => win.webContents.forcedCrashes === 1 && win.webContents.reloads === 1, 'the forced reload')
  win.webContents.emit('render-process-gone', {}, { reason: 'killed', exitCode: -1 })
  assert.equal(await flow, 'canceled', 'the reload, not the pending close, decides')
  assert.equal(win.isDestroyed(), false)
})

test('at Windows shutdown each page gets two seconds to keep its edits, and shutdown is never blocked', async () => {
  const { win, page } = openWindow({ status: { dirty: true, saving: false, title: 'Budget.xlsx' } })
  const event = { reasons: ['shutdown'], defaultPrevented: false, preventDefault() { this.defaultPrevented = true } }
  win.emit('query-session-end', event)
  assert.equal(event.defaultPrevented, false)
  await until(() => callsOf(page, 'recovery-flush').length === 1, 'the recovery flush')
  assert.deepEqual(callsOf(page, 'recovery-flush')[0][1], { reason: 'session-end', budgetMs: TIMING.flushTimeoutMs })
})

// ---------------------------------------------------------------------------
// Quitting
// ---------------------------------------------------------------------------

test('quitting waits for pending writes before it asks anything', async () => {
  const app = makeApp()
  guard.installAppGuard(app)
  const { win, page } = openWindow({ status: { dirty: false, saving: false, title: 'Budget.xlsx' } })
  let finishWrite
  core.trackPendingWrite(new Promise((resolve) => { finishWrite = resolve }), 'Budget.xlsx')
  app.quit()
  assert.equal(app.quitCount, 0, 'quit is held')
  await new Promise((resolve) => setTimeout(resolve, 80))
  assert.equal(callsOf(page, 'close-query').length, 0, 'no window is asked while a write is pending')
  assert.equal(app.quitCount, 0)
  finishWrite()
  await until(() => app.quitCount === 1, 'the quit')
  assert.equal(win.isDestroyed(), true)
})

test('quitting asks each window in turn, and Cancel stops the quit', async () => {
  const app = makeApp()
  const api = guard.installAppGuard(app)
  assert.equal(api.isQuitApproved(), false, 'a new app starts unapproved')
  const first = openWindow({ status: { dirty: true, saving: false, title: 'First.xlsx' } })
  const second = openWindow({ status: { dirty: true, saving: false, title: 'Second.xlsx' } })
  answers.push('Don\'t Save', 'Cancel')
  app.quit()
  await until(() => boxes.filter((box) => box.message.includes('"Second.xlsx"')).length === 1, 'the second window\'s prompt')
  await until(() => api.quitFlow() === null, 'the quit flow to end')
  assert.equal(app.quitCount, 0, 'Cancel stopped the quit')
  assert.equal(first.win.isDestroyed(), true, 'a window closes right after its decision')
  assert.equal(second.win.isDestroyed(), false)
  assert.ok(first.win.focusCount >= 1 && second.win.focusCount >= 1, 'each window is brought forward before it asks')
  assert.equal(callsOf(first.page, 'discard').length, 1)
  assert.deepEqual(callsOf(first.page, 'close-query')[0][1], { quitting: true })

  answers.push('Save')
  app.quit()
  await until(() => app.quitCount === 1, 'the second quit')
  assert.equal(callsOf(second.page, 'save-now').length, 1)
  assert.deepEqual(callsOf(second.page, 'save-now')[0][1], { reason: 'quit' })
  assert.equal(second.win.isDestroyed(), true)
  assert.equal(api.isQuitApproved(), true)
})

test('a window that opens while quitting asks is asked too, never closed unasked', async () => {
  const app = makeApp()
  const api = guard.installAppGuard(app)
  openWindow({ status: { dirty: true, saving: false, title: 'First.xlsx' } })
  openWindow({ status: { dirty: true, saving: false, title: 'Second.xlsx' } })
  let third = null
  // While the first window's prompt is showing, a double-clicked file opens a new window with edits.
  answers.push(() => {
    third = openWindow({ status: { dirty: true, saving: false, title: 'Third.xlsx' } })
    return 'Don\'t Save'
  }, 'Don\'t Save', 'Cancel')
  app.quit()
  await until(() => boxes.some((box) => box.message.includes('"Third.xlsx"')), 'the new window\'s prompt')
  await until(() => api.quitFlow() === null, 'the quit flow to end')
  assert.equal(app.quitCount, 0, 'Cancel in the new window stopped the quit')
  assert.equal(third.win.isDestroyed(), false, 'the new window and its edits are still there')
  assert.equal(callsOf(third.page, 'discard').length, 0)
  assert.equal(api.isQuitApproved(), false)
  // Closing it on its own still asks.
  answers.push('Cancel')
  assert.equal(third.win.close().defaultPrevented, true)
})

test('a quit that something else stops switches the guards back on', async () => {
  guard.configureDocumentGuard({ quitConfirmMs: 100 })
  const app = makeApp()
  const api = guard.installAppGuard(app)
  // Another part of the app refuses to quit after the guard approved it.
  app.on('before-quit', (event) => { if (api.isQuitApproved()) event.preventDefault() })
  const { win, page } = openWindow({ status: { dirty: false, saving: false, title: 'Budget.xlsx' } })
  app.quit()
  await until(() => api.isQuitApproved(), 'the approval')
  assert.equal(app.quitCount, 0)
  assert.equal(win.isDestroyed(), true, 'the clean window closed with its decision')
  await until(() => !api.isQuitApproved(), 'the guards to switch back on', 2000)
  const next = openWindow({ status: { dirty: true, saving: false, title: 'Later.xlsx' } })
  answers.push('Cancel')
  assert.equal(next.win.close().defaultPrevented, true, 'closing asks about unsaved changes again')
  await until(() => boxes.at(-1).message === 'Do you want to save the changes to "Later.xlsx"?', 'the prompt')
  assert.equal(callsOf(page, 'close-query').length, 1)
  guard.configureDocumentGuard({ quitConfirmMs: 10_000 })
})
