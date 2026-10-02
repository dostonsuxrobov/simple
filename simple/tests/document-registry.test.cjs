'use strict'

// The document registry and performSave() of the shared Save layer
// (document-guard.cjs, design §3.1, §3.3, §3.4, §3.8, §9 step 3). Real files
// in a temp folder are opened, changed by "another program", deleted, locked
// by another window and saved; `electron` is a stub that records dialogs and
// answers them from a queue. Nothing outside the temp folder is touched.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const Module = require('node:module')
const { EventEmitter } = require('node:events')
const { pathToFileURL } = require('node:url')
const { findServiceNames } = require('../scripts/local-only-guard.cjs')

const SHARED = path.resolve(__dirname, '..', 'shared')
const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-document-registry-'))
const APP_PATH = path.join(ROOT, 'app')

// ---------------------------------------------------------------------------
// Electron stub
// ---------------------------------------------------------------------------

const record = { save: [], box: [], openPath: [] }
const queue = { save: [], box: [] }
const windows = new Set()
let nextContentsId = 1
const take = async (list, options) => {
  if (!list.length) return undefined
  const next = list.shift()
  return typeof next === 'function' ? next(options) : next
}
class FakeWebContents extends EventEmitter {
  constructor() {
    super()
    this.id = nextContentsId++
    this.destroyed = false
    this.mainFrame = { url: pathToFileURL(path.join(APP_PATH, 'dist', 'index.html')).href }
  }
  send() {}
  isDestroyed() { return this.destroyed }
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
const electronStub = {
  app: Object.assign(new EventEmitter(), { isPackaged: false, getPath: () => path.join(ROOT, 'unused'), getAppPath: () => APP_PATH, getVersion: () => '9.9.9' }),
  BrowserWindow,
  dialog: {
    async showSaveDialog(win, options) {
      const opts = options || win
      record.save.push(opts)
      const answer = await take(queue.save, opts)
      return answer ? { canceled: false, filePath: answer } : { canceled: true, filePath: '' }
    },
    async showOpenDialog() { return { canceled: true, filePaths: [] } },
    async showMessageBox(win, options) {
      const opts = options || win
      record.box.push(opts)
      let answer = await take(queue.box, opts)
      if (answer === undefined) answer = opts.cancelId
      const index = typeof answer === 'number' ? answer : opts.buttons.indexOf(answer)
      if (index < 0) throw new Error(`No button "${answer}" in ${opts.buttons.join(', ')}`)
      return { response: index, checkboxChecked: false }
    },
  },
  ipcMain: Object.assign(new EventEmitter(), { handle() {}, removeHandler() {} }),
  shell: { showItemInFolder() {}, openPath: (target) => { record.openPath.push(target); return Promise.resolve('') } },
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
dialogs.configureDialogs({ module: 'calc', userData: path.join(ROOT, 'userData'), documentsDir: path.join(ROOT, 'Documents'), desktopDir: path.join(ROOT, 'Desktop'), logger: quiet })
guard.configureDocumentGuard({ module: 'calc', logger: quiet })
fs.mkdirSync(path.join(ROOT, 'Documents'), { recursive: true })

test.after(() => {
  Module._resolveFilename = originalResolve
  fs.rmSync(ROOT, { recursive: true, force: true })
})

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

let caseCount = 0
function scenario(name) {
  caseCount += 1
  const folder = path.join(ROOT, `${String(caseCount).padStart(2, '0')}-${name}`)
  fs.mkdirSync(folder, { recursive: true })
  queue.save.length = 0
  queue.box.length = 0
  return folder
}

function eventOf(win) {
  return { sender: win.webContents }
}

/** A minimal, valid OOXML workbook (stored ZIP entries). */
function storedZip(files) {
  const locals = []
  const centrals = []
  let offset = 0
  for (const [name, content] of files) {
    const data = Buffer.from(content)
    const nameBytes = Buffer.from(name)
    const crc = zlib.crc32(data)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(data.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(nameBytes.length, 26)
    const central = Buffer.alloc(46)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(20, 4)
    central.writeUInt16LE(20, 6)
    central.writeUInt32LE(crc, 16)
    central.writeUInt32LE(data.length, 20)
    central.writeUInt32LE(data.length, 24)
    central.writeUInt16LE(nameBytes.length, 28)
    central.writeUInt32LE(offset, 42)
    locals.push(local, nameBytes, data)
    centrals.push(central, nameBytes)
    offset += 30 + nameBytes.length + data.length
  }
  const directory = Buffer.concat(centrals)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(files.length, 8)
  end.writeUInt16LE(files.length, 10)
  end.writeUInt32LE(directory.length, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...locals, directory, end])
}

function workbook(marker) {
  return storedZip([
    ['[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>'],
    ['_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', `<?xml version="1.0"?><workbook><!-- ${marker} --></workbook>`],
  ])
}

const CFB = Buffer.concat([Buffer.from('d0cf11e0a1b11ae1', 'hex'), Buffer.alloc(1024)])

/** Saves text, whatever the format (CSV is the format these cases write). */
function textHooks(text, extra = {}) {
  let calls = 0
  const hooks = {
    serialize: () => {
      calls += 1
      return Buffer.from(typeof text === 'function' ? text() : text)
    },
    ...extra,
  }
  return { hooks, calls: () => calls }
}

async function openIn(win, filePath, options = {}) {
  const opened = await guard.openDocument(filePath, { webContents: win.webContents, ...options })
  assert.equal(opened.ok, true, JSON.stringify(opened))
  return opened
}

// ---------------------------------------------------------------------------
// Registry
// ---------------------------------------------------------------------------

test('open records main\'s stamp, and a second window opening the same file focuses the first', async () => {
  const folder = scenario('open')
  const file = path.join(folder, 'Budget.csv')
  fs.writeFileSync(file, 'a,b\n1,2\n')
  const first = new BrowserWindow()
  const opened = await openIn(first, file)
  assert.equal(opened.focused, false)
  assert.deepEqual(opened.stamp, await core.stampFile(file))
  assert.ok(opened.bytes.equals(fs.readFileSync(file)))
  assert.equal(opened.format, 'csv')
  assert.deepEqual(guard.getDocument(opened.docId).stamp, opened.stamp)

  const second = new BrowserWindow()
  const again = await guard.openDocument(process.platform === 'win32' ? file.toUpperCase() : file, { webContents: second.webContents })
  assert.equal(again.focused, true)
  assert.equal(again.docId, opened.docId)
  assert.equal(first.focusCount, 1, 'the window that has the file comes to the front')
  assert.deepEqual(guard.documentsForWindow(second), [])

  const reread = await openIn(first, file)
  assert.equal(reread.docId, opened.docId, 'a window re-reading its own file keeps the document')
  const missing = await guard.openDocument(path.join(folder, 'gone.csv'), { webContents: second.webContents })
  assert.equal(missing.ok, false)
  assert.equal(missing.code, 'NOT_FOUND')
  assert.equal(missing.context, 'open')
  assert.deepEqual(guard.documentsForWindow(second), [], 'a failed open binds nothing')

  first.destroy()
  assert.equal(guard.getDocument(opened.docId), null, 'closing the window forgets its documents')
  assert.equal(guard.findDocumentByPath(file), null)
  second.destroy()
})

test('a file opened in a window replaces what it showed, unless the window keeps several documents', async () => {
  const folder = scenario('replace')
  const [a, b, c] = ['A.csv', 'B.csv', 'C.csv'].map((name) => path.join(folder, name))
  for (const file of [a, b, c]) fs.writeFileSync(file, 'x\n')
  const win = new BrowserWindow()
  await openIn(win, a)
  await openIn(win, b)
  assert.equal(guard.findDocumentByPath(a), null, 'the earlier file is no longer held by this window')
  assert.deepEqual(guard.documentsForWindow(win).map((doc) => doc.path), [b])
  await openIn(win, c, { replace: false })
  assert.deepEqual(guard.documentsForWindow(win).map((doc) => doc.path).sort(), [b, c].sort())
  const other = new BrowserWindow()
  assert.equal((await openIn(other, a)).focused, false, 'the released file opens normally elsewhere')
  win.destroy()
  other.destroy()
})

// ---------------------------------------------------------------------------
// performSave against the file on disk
// ---------------------------------------------------------------------------

test('a save writes in place, also after a sync program re-stamped the file', async () => {
  const folder = scenario('in-place')
  const file = path.join(folder, 'Budget.csv')
  fs.writeFileSync(file, 'a,b\n1,2\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  const { hooks } = textHooks('a,b\n3,4\n')
  const saved = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 4 }, hooks)
  assert.equal(saved.ok, true, JSON.stringify(saved))
  assert.equal(saved.path, file)
  assert.equal(saved.revision, 4, 'the revision is echoed so only that revision is marked saved')
  assert.equal(saved.rebound, false)
  assert.equal(saved.folderChanged, false)
  assert.equal(fs.readFileSync(file, 'utf8'), 'a,b\n3,4\n')
  assert.deepEqual(guard.getDocument(docId).stamp, saved.stamp)

  const later = new Date(Date.now() + 120_000)
  fs.utimesSync(file, later, later)
  const again = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 5 }, textHooks('a,b\n5,6\n').hooks)
  assert.equal(again.ok, true, 'equal content with a new time is not a conflict')
  assert.equal(fs.readFileSync(file, 'utf8'), 'a,b\n5,6\n')
  win.destroy()
})

test('a file changed by another program is never overwritten without asking; Replace overwrites it', async () => {
  const folder = scenario('changed')
  const file = path.join(folder, 'Budget.csv')
  fs.writeFileSync(file, 'a,b\n1,2\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  fs.writeFileSync(file, 'a,b\n9,9\nchanged elsewhere\n')
  const refused = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 2 }, textHooks('a,b\nmine\n').hooks)
  assert.equal(refused.ok, false)
  assert.equal(refused.code, 'CHANGED_ON_DISK')
  assert.equal(refused.name, 'Budget.csv')
  assert.equal(refused.docId, docId)
  assert.equal(fs.readFileSync(file, 'utf8'), 'a,b\n9,9\nchanged elsewhere\n', 'the other program\'s version is untouched')
  // force alone is the page's word: it counts only after main showed the prompt and the user chose Replace.
  const unasked = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 2, force: true }, textHooks('a,b\nmine\n').hooks)
  assert.equal(unasked.code, 'CHANGED_ON_DISK')
  assert.equal(fs.readFileSync(file, 'utf8'), 'a,b\n9,9\nchanged elsewhere\n')
  dialogs.recordDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file)
  const replaced = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 2, force: true }, textHooks('a,b\nmine\n').hooks)
  assert.equal(replaced.ok, true)
  assert.equal(fs.readFileSync(file, 'utf8'), 'a,b\nmine\n')
  assert.equal(dialogs.hasDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file), false, 'a decision is used up by the save it allowed')
  win.destroy()
})

test('Replace replaces only the version the user was shown; a newer change asks again, also on Try Again', async () => {
  const folder = scenario('replace-shown')
  const file = path.join(folder, 'Budget.csv')
  fs.writeFileSync(file, 'a,b\n1,2\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  fs.writeFileSync(file, 'theirs, version 1\n')
  const first = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 2 }, textHooks('mine\n').hooks)
  assert.equal(first.code, 'CHANGED_ON_DISK')
  dialogs.recordDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file)
  // The other program saves again (as it closes) before the user's Replace or Try Again arrives.
  fs.writeFileSync(file, 'theirs, version 2 with more rows\n')
  const again = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 2, force: true }, textHooks('mine\n').hooks)
  assert.equal(again.code, 'CHANGED_ON_DISK', 'a change after the prompt is never overwritten on the old answer')
  assert.equal(fs.readFileSync(file, 'utf8'), 'theirs, version 2 with more rows\n')
  dialogs.recordDecision(win.webContents, 'saveFailed.CHANGED_ON_DISK', 'replace', file)
  const replaced = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 2, force: true }, textHooks('mine\n').hooks)
  assert.equal(replaced.ok, true, JSON.stringify(replaced))
  assert.equal(fs.readFileSync(file, 'utf8'), 'mine\n')
  win.destroy()
})

test('a moved or deleted file reports SOURCE_MISSING; Save Here Again recreates it', async () => {
  const folder = scenario('missing')
  const file = path.join(folder, 'Budget.csv')
  fs.writeFileSync(file, 'a,b\n1,2\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  fs.rmSync(file)
  const refused = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1 }, textHooks('a,b\nmine\n').hooks)
  assert.equal(refused.code, 'SOURCE_MISSING')
  assert.equal(fs.existsSync(file), false)
  const unasked = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1, recreate: true }, textHooks('a,b\nmine\n').hooks)
  assert.equal(unasked.code, 'SOURCE_MISSING', 'recreate counts only after the user chose Save Here Again')
  dialogs.recordDecision(win.webContents, 'saveFailed.SOURCE_MISSING', 'recreate', file)
  const recreated = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1, recreate: true }, textHooks('a,b\nmine\n').hooks)
  assert.equal(recreated.ok, true)
  assert.equal(recreated.strategy, 'new')
  assert.equal(fs.readFileSync(file, 'utf8'), 'a,b\nmine\n')
  win.destroy()
})

test('bookkeeping after a save can never turn it into a failure', async () => {
  const folder = scenario('after-save')
  const file = path.join(folder, 'Budget.csv')
  fs.writeFileSync(file, 'a\n1\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  const seen = []
  const saved = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1 }, textHooks('a\n2\n', {
    afterSave: (result) => {
      seen.push(result.path)
      throw new Error('the recent files list is locked')
    },
  }).hooks)
  assert.equal(saved.ok, true)
  assert.deepEqual(seen, [file])
  assert.equal(fs.readFileSync(file, 'utf8'), 'a\n2\n')
  win.destroy()
})

test('Save As rebinds the document; Save a Copy and Export do not', async () => {
  const folder = scenario('targets')
  const win = new BrowserWindow()
  // The workspace knows a plain table loses nothing in CSV, so no lossy prompt here.
  const hooks = textHooks('a,b\n1,2\n', { formats: ['csv', 'tsv'], lossy: () => null }).hooks
  queue.save.push(path.join(folder, 'New.csv'))
  const first = await guard.performSave(eventOf(win), { docId: 'untitled-doc-1', mode: 'save', revision: 1 }, hooks)
  assert.equal(first.ok, true, JSON.stringify(first))
  assert.equal(record.save.at(-1).defaultPath, path.join(ROOT, 'Documents', 'Untitled spreadsheet.csv'), 'an untitled Save asks where, starting in Documents')
  assert.equal(first.rebound, true)
  assert.equal(first.folderChanged, true)
  assert.equal(guard.getDocument('untitled-doc-1').path, path.join(folder, 'New.csv'))

  queue.save.push(path.join(folder, 'Renamed.csv'))
  const renamed = await guard.performSave(eventOf(win), { docId: 'untitled-doc-1', mode: 'save-as', revision: 2 }, hooks)
  assert.equal(renamed.ok, true)
  assert.equal(record.save.at(-1).defaultPath, path.join(folder, 'New.csv'), 'Save As starts at the document\'s own file')
  assert.equal(guard.findDocumentByPath(path.join(folder, 'New.csv')), null)
  assert.equal(guard.findDocumentByPath(path.join(folder, 'Renamed.csv')).docId, 'untitled-doc-1')

  queue.save.push(path.join(folder, 'Copy.csv'))
  const copy = await guard.performSave(eventOf(win), { docId: 'untitled-doc-1', mode: 'save-copy', revision: 3 }, hooks)
  assert.equal(copy.ok, true)
  assert.equal(copy.rebound, false)
  assert.equal(guard.getDocument('untitled-doc-1').path, path.join(folder, 'Renamed.csv'), 'a copy leaves the document where it was')

  const chosen = await dialogs.chooseSavePath({ purpose: 'export', name: 'Renamed.csv', format: 'tsv', formats: ['tsv', 'csv'] }, { sender: win.webContents })
  assert.equal(chosen, null, 'no answer queued: canceled')
  queue.save.push(path.join(folder, 'Data.tsv'))
  const target = await dialogs.chooseSavePath({ purpose: 'export', name: 'Renamed.csv', format: 'tsv', formats: ['tsv', 'csv'] }, { sender: win.webContents })
  const exported = await guard.performSave(eventOf(win), { docId: 'untitled-doc-1', mode: 'export', path: target.path, format: target.format }, hooks)
  assert.equal(exported.ok, true)
  assert.equal(exported.format, 'tsv')
  assert.equal(guard.getDocument('untitled-doc-1').path, path.join(folder, 'Renamed.csv'))

  const sneaky = await guard.performSave(eventOf(win), { docId: 'untitled-doc-1', mode: 'export', path: path.join(folder, 'Never chosen.csv'), format: 'csv' }, hooks)
  assert.equal(sneaky.ok, false, 'a path the user did not choose in this window is refused')
  assert.equal(fs.existsSync(path.join(folder, 'Never chosen.csv')), false)

  // An export over the document's own file refreshes its stamp, so the next Save is not a false conflict.
  const own = path.join(folder, 'Renamed.csv')
  const overOwn = await guard.performSave(eventOf(win), { docId: 'untitled-doc-1', mode: 'export', path: own, format: 'csv' }, textHooks('exported\n').hooks)
  assert.equal(overOwn.ok, true)
  assert.equal(overOwn.ownFileChanged, true, 'the page learns that its file no longer holds the document')
  const next = await guard.performSave(eventOf(win), { docId: 'untitled-doc-1', mode: 'save', revision: 4 }, textHooks('a,b\n4,4\n').hooks)
  assert.equal(next.ok, true, JSON.stringify(next))
  win.destroy()
})

test('after an export over the document\'s own file, an untouched Save writes the document, not nothing', async () => {
  const folder = scenario('export-over-own')
  const file = path.join(folder, 'data.csv')
  fs.writeFileSync(file, 'a,b\n1,2\n3,4\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  queue.save.push(file)
  const target = await dialogs.chooseSavePath({ purpose: 'export', name: 'data.csv', format: 'csv', formats: ['csv'] }, { sender: win.webContents })
  const exported = await guard.performSave(eventOf(win), { docId, mode: 'export', path: target.path, format: 'csv' }, textHooks('a;b\n1;2\n').hooks)
  assert.equal(exported.ok, true, JSON.stringify(exported))
  assert.equal(exported.ownFileChanged, true)
  const model = textHooks('a,b\n1,2\n3,4\n', { formats: ['csv'] })
  const saved = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 0, pristine: true }, model.hooks)
  assert.equal(saved.ok, true, JSON.stringify(saved))
  assert.notEqual(saved.strategy, 'unchanged', 'the pristine shortcut is refused: the file holds the export')
  assert.equal(model.calls(), 1, 'the document is serialized, not the export\'s bytes reused')
  assert.equal(fs.readFileSync(file, 'utf8'), 'a,b\n1,2\n3,4\n')
  const after = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 0, pristine: true }, model.hooks)
  assert.equal(after.strategy, 'unchanged', 'once the file holds the document again, the shortcut is back')
  win.destroy()
})

test('a restored recovery copy is always written: the pristine shortcut and the opened bytes are refused until a real save', async () => {
  const folder = scenario('adopt-pristine')
  const file = path.join(folder, 'Notes.csv')
  fs.writeFileSync(file, 'ORIGINAL\n')
  const stamp = await core.stampFile(file)
  const win = new BrowserWindow()
  guard.adoptDocument(win.webContents, { docId: 'restored-doc-1', path: file, stamp, format: 'csv', name: 'Notes.csv' })
  const model = textHooks('RECOVERED EDITS\n', { formats: ['csv'] })
  const saved = await guard.performSave(eventOf(win), { docId: 'restored-doc-1', mode: 'save', revision: 3, pristine: true }, model.hooks)
  assert.equal(saved.ok, true, JSON.stringify(saved))
  assert.notEqual(saved.strategy, 'unchanged')
  assert.equal(fs.readFileSync(file, 'utf8'), 'RECOVERED EDITS\n', 'the recovered work reaches the file')
  const again = await guard.performSave(eventOf(win), { docId: 'restored-doc-1', mode: 'save', revision: 3, pristine: true }, model.hooks)
  assert.equal(again.strategy, 'unchanged', 'after a real write the shortcut is allowed again')
  win.destroy()
})

test('a restored copy replaces the window\'s document for the same file, so the banner restore can be saved', async () => {
  const folder = scenario('banner-restore')
  const file = path.join(folder, 'Budget.csv')
  fs.writeFileSync(file, 'a,b\n1,2\n')
  const stamp = await core.stampFile(file)
  const win = new BrowserWindow()
  const { docId: openedId } = await openIn(win, file)
  assert.equal(guard.canAdoptDocument(win.webContents, { docId: 'entry-d', path: file }), true)
  guard.adoptDocument(win.webContents, { docId: 'entry-d', path: file, stamp, format: 'csv', name: 'Budget.csv' })
  assert.equal(guard.getDocument(openedId), null, 'the document the window showed for that file is released')
  const saved = await guard.performSave(eventOf(win), { docId: 'entry-d', mode: 'save', revision: 2 }, textHooks('a,b\nrestored\n').hooks)
  assert.equal(saved.ok, true, JSON.stringify(saved))
  assert.equal(fs.readFileSync(file, 'utf8'), 'a,b\nrestored\n')
  const other = new BrowserWindow()
  assert.equal(guard.canAdoptDocument(other.webContents, { docId: 'entry-e', path: file }), false, 'another window still cannot take a file this window shows')
  assert.throws(() => guard.adoptDocument(other.webContents, { docId: 'entry-e', path: file, stamp, format: 'csv' }), /another window/)
  win.destroy()
  other.destroy()
})

test('a document bound to a file always has a stamp from main; without one, saving asks first', async () => {
  const folder = scenario('stampless')
  const file = path.join(folder, 'Data.csv')
  fs.writeFileSync(file, 'x\n1\n')
  const win = new BrowserWindow()
  // An integrator registers a file it read itself: main stamps it.
  guard.registerDocument({ webContents: win.webContents, docId: 'own-read-1', path: file, format: 'csv' })
  const saved = await guard.performSave(eventOf(win), { docId: 'own-read-1', mode: 'save', revision: 1 }, textHooks('x\n2\n').hooks)
  assert.equal(saved.ok, true, JSON.stringify(saved))
  fs.writeFileSync(file, 'x\n2\nanother program\n')
  const conflict = await guard.performSave(eventOf(win), { docId: 'own-read-1', mode: 'save', revision: 2 }, textHooks('x\n3\n').hooks)
  assert.equal(conflict.code, 'CHANGED_ON_DISK', 'conflict detection works for a registered path')

  // bindDocument with a path and no stamp stamps in main too.
  const other = path.join(folder, 'Other.csv')
  fs.writeFileSync(other, 'o\n1\n')
  guard.bindDocument('own-read-1', { path: other, stamp: null })
  await new Promise((resolve) => setTimeout(resolve, 50))
  assert.ok(guard.getDocument('own-read-1').stamp, 'main stamped the bound file')
  fs.writeFileSync(other, 'o\n1\nchanged by another program\n')
  const bound = await guard.performSave(eventOf(win), { docId: 'own-read-1', mode: 'save', revision: 3 }, textHooks('o\n9\n').hooks)
  assert.equal(bound.code, 'CHANGED_ON_DISK')

  // A file that could not be stamped (it did not exist yet) is never overwritten unasked.
  const late = path.join(folder, 'Late.csv')
  guard.registerDocument({ webContents: win.webContents, docId: 'late-1', path: late, format: 'csv' })
  await new Promise((resolve) => setTimeout(resolve, 20))
  fs.writeFileSync(late, 'written by another program\n')
  const refused = await guard.performSave(eventOf(win), { docId: 'late-1', mode: 'save', revision: 1 }, textHooks('mine\n').hooks)
  assert.equal(refused.code, 'CHANGED_ON_DISK')
  assert.equal(fs.readFileSync(late, 'utf8'), 'written by another program\n')

  // A recovery entry without a snapshot-time stamp comes back untitled, with Save As starting at its file.
  const adopted = guard.adoptDocument(win.webContents, { docId: 'no-stamp-entry', path: file, stamp: null, format: 'csv', name: 'Data.csv' })
  assert.equal(adopted.untitled, true)
  assert.equal(adopted.suggestedPath, file)
  win.destroy()
})

test('flags that skip a check need the matching answer from a prompt main showed this window', async () => {
  const folder = scenario('decisions')
  const file = path.join(folder, 'Plain.csv')
  fs.writeFileSync(file, 'a\n1\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  // confirmedLossy from the page alone does not skip main's lossy prompt.
  const lossyHooks = textHooks('a\n2\n', { formats: ['xlsx', 'csv'], fullFormat: 'xlsx', lossy: () => ({ keeps: 'the values', lost: ['formulas'] }) }).hooks
  queue.box.push('Cancel')
  const asked = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1, confirmedLossy: true }, lossyHooks)
  assert.equal(asked.code, 'CANCELED')
  assert.match(record.box.at(-1).message, /^Save "Plain\.csv" as /, 'main asked itself')
  dialogs.recordDecision(win.webContents, 'prompts.lossy-save', 'save-lossy')
  const boxes = record.box.length
  const confirmed = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1, confirmedLossy: true }, lossyHooks)
  assert.equal(confirmed.ok, true, JSON.stringify(confirmed))
  assert.equal(record.box.length, boxes, 'after the page showed the prompt itself, main does not ask again')
  win.destroy()
})

test('an untouched document writes its opened bytes, and an untouched Save writes nothing', async () => {
  const folder = scenario('pristine')
  const file = path.join(folder, 'Budget.csv')
  const original = Buffer.from('x;y\r\n1;2\r\n')
  fs.writeFileSync(file, original)
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  const serializer = textHooks('re-serialized\n', { formats: ['csv'] })
  const before = fs.statSync(file).mtimeMs
  const unchanged = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 0, pristine: true }, serializer.hooks)
  assert.equal(unchanged.ok, true)
  assert.equal(unchanged.strategy, 'unchanged')
  assert.equal(fs.statSync(file).mtimeMs, before)
  queue.save.push(path.join(folder, 'Copy of Budget.csv'))
  const copy = await guard.performSave(eventOf(win), { docId, mode: 'save-copy', revision: 0, pristine: true }, serializer.hooks)
  assert.equal(copy.ok, true)
  assert.ok(fs.readFileSync(path.join(folder, 'Copy of Budget.csv')).equals(original), 'the copy has the opened bytes exactly')
  assert.equal(serializer.calls(), 0, 'an untouched file is never re-serialized')
  fs.writeFileSync(file, 'changed elsewhere')
  const conflict = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 0, pristine: true }, serializer.hooks)
  assert.equal(conflict.code, 'CHANGED_ON_DISK')
  win.destroy()
})

test('a serializer failure keeps the original, and a missing engine offers a modern copy next to it', async () => {
  const folder = scenario('serialize')
  const file = path.join(folder, 'Budget.csv')
  fs.writeFileSync(file, 'a\n1\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  const failed = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1 }, { serialize: () => { throw new Error('cell A1 has a cycle') } })
  assert.equal(failed.ok, false)
  assert.equal(failed.code, 'SERIALIZE_FAILED')
  assert.equal(failed.formatLabel, 'Comma-separated values (.csv)')
  assert.equal(failed.message, 'Simple couldn\'t create the Comma-separated values (.csv) file.')
  assert.match(failed.technical, /cell A1 has a cycle/)
  assert.equal(fs.readFileSync(file, 'utf8'), 'a\n1\n')
  const invalid = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1 }, { serialize: () => Buffer.from([0xff, 0xfe, 0x00, 0xd8, 0x00]) })
  assert.equal(invalid.ok, false)
  assert.ok(['VALIDATION_FAILED', 'VERIFY_FAILED'].includes(invalid.code), invalid.code)
  assert.equal(fs.readFileSync(file, 'utf8'), 'a\n1\n', 'bytes that fail the format check never replace the file')

  const legacy = path.join(folder, 'Old.xls')
  fs.writeFileSync(legacy, CFB)
  const { docId: legacyId } = await openIn(win, legacy)
  const needsEngine = Object.assign(new Error('Simple can\'t save Excel 97–2003 workbook files on this PC.'), {
    name: 'OfficeEngineError', code: 'NEEDS_OFFICE_ENGINE', formatLabel: 'Excel 97–2003 workbook', altFormat: 'xlsx', altLabel: 'Excel workbook', altExt: '.xlsx',
  })
  const engineGone = await guard.performSave(eventOf(win), { docId: legacyId, mode: 'save', revision: 1 }, {
    engine: true,
    serialize: (format) => { if (format === 'xls') throw needsEngine; return workbook('modern') },
  })
  assert.equal(engineGone.ok, false)
  assert.equal(engineGone.code, 'NEEDS_OFFICE_ENGINE')
  assert.equal(engineGone.altFormat, 'xlsx')
  assert.equal(engineGone.altExt, '.xlsx')
  const modern = await guard.performSave(eventOf(win), { docId: legacyId, mode: 'save', revision: 1, format: 'xlsx', fallbackReason: 'NEEDS_OFFICE_ENGINE' }, {
    serialize: () => workbook('modern'),
  })
  assert.equal(modern.ok, true, JSON.stringify(modern))
  assert.equal(modern.path, path.join(folder, 'Old.xlsx'))
  assert.equal(modern.sibling, true)
  assert.equal(modern.originalName, 'Old.xls')
  assert.equal(modern.originalFormatLabel, 'Excel 97–2003 workbook (.xls)')
  assert.ok(fs.readFileSync(legacy).equals(CFB), 'the original stays unchanged')
  assert.equal(guard.getDocument(legacyId).path, path.join(folder, 'Old.xlsx'))
  win.destroy()
})

test('a format Simple can\'t update goes to a sibling file after one confirmation', async () => {
  const folder = scenario('sibling')
  const legacy = path.join(folder, 'Q3.xls')
  fs.writeFileSync(legacy, CFB)
  const win = new BrowserWindow()
  const { docId } = await openIn(win, legacy)
  const hooks = { serialize: (format) => (format === 'xlsx' ? workbook('Q3 edited') : Buffer.from('never')) }
  queue.box.push('Save as "Q3.xlsx"')
  const saved = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1 }, hooks)
  const prompt = record.box.at(-1)
  assert.equal(prompt.message, 'Save your changes as "Q3.xlsx"?')
  assert.deepEqual(prompt.buttons, ['Save as "Q3.xlsx"', 'Choose Location…', 'Cancel'])
  assert.equal(prompt.defaultId, 0)
  assert.equal(prompt.cancelId, 2)
  assert.match(prompt.detail, /"Q3\.xls" stays unchanged/)
  assert.equal(saved.ok, true, JSON.stringify(saved))
  assert.equal(saved.path, path.join(folder, 'Q3.xlsx'))
  assert.equal(saved.sibling, true)
  assert.ok(fs.readFileSync(legacy).equals(CFB))
  const boxes = record.box.length
  const second = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 2 }, { serialize: () => workbook('Q3 again') })
  assert.equal(second.ok, true)
  assert.equal(second.path, path.join(folder, 'Q3.xlsx'), 'later saves go straight to the sibling')
  assert.equal(record.box.length, boxes, 'without asking again')

  const other = path.join(folder, 'R.xls')
  fs.writeFileSync(other, CFB)
  fs.writeFileSync(path.join(folder, 'R.xlsx'), 'somebody else\'s workbook')
  const { docId: otherId } = await openIn(win, other)
  queue.box.push('Cancel')
  const canceled = await guard.performSave(eventOf(win), { docId: otherId, mode: 'save', revision: 1 }, hooks)
  assert.equal(canceled.code, 'CANCELED')
  assert.equal(fs.existsSync(path.join(folder, 'R (edited).xlsx')), false)
  queue.box.push('Save as "R (edited).xlsx"')
  const edited = await guard.performSave(eventOf(win), { docId: otherId, mode: 'save', revision: 1 }, hooks)
  assert.equal(edited.ok, true, JSON.stringify(edited))
  assert.equal(edited.path, path.join(folder, 'R (edited).xlsx'))
  assert.equal(fs.readFileSync(path.join(folder, 'R.xlsx'), 'utf8'), 'somebody else\'s workbook', 'an unrelated file with the sibling name is never replaced')
  win.destroy()
})

test('the first overwrite in a session keeps the original in the versions store', async () => {
  const folder = scenario('versions')
  const file = path.join(folder, 'Ledger.csv')
  fs.writeFileSync(file, 'original,ledger\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  const first = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1 }, textHooks('first,save\n').hooks)
  assert.equal(first.ok, true)
  assert.equal(first.versionBackup.ok, true)
  assert.ok(first.backupPath && fs.readFileSync(first.backupPath, 'utf8') === 'original,ledger\n')
  const second = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 2 }, textHooks('second,save\n').hooks)
  assert.equal(second.ok, true)
  assert.deepEqual(second.versionBackup, { ok: true, skipped: 'already-backed-up' })
  const versions = await stores.versionsStore().list(file)
  assert.equal(versions.length, 1)
  assert.equal(fs.readFileSync(versions[0].path, 'utf8'), 'original,ledger\n')
  win.destroy()
})

test('a file another window is editing is never written over, and documents belong to their window', async () => {
  const folder = scenario('one-window')
  const file = path.join(folder, 'Shared.csv')
  fs.writeFileSync(file, 'theirs\n')
  const holder = new BrowserWindow()
  const { docId: heldId } = await openIn(holder, file)
  const writer = new BrowserWindow()
  queue.save.push(file)
  const refused = await guard.performSave(eventOf(writer), { docId: 'writer-doc', mode: 'save-as', revision: 1 }, textHooks('mine\n', { formats: ['csv'] }).hooks)
  assert.equal(refused.ok, false)
  assert.equal(refused.code, 'LOCKED')
  assert.equal(refused.reason, 'open-in-another-window')
  assert.match(refused.message, /open in another Simple window/)
  assert.equal(fs.readFileSync(file, 'utf8'), 'theirs\n')

  const stolen = await guard.performSave(eventOf(writer), { docId: heldId, mode: 'save', revision: 1 }, textHooks('mine\n').hooks)
  assert.equal(stolen.ok, false, 'a window cannot save another window\'s document')
  assert.equal(fs.readFileSync(file, 'utf8'), 'theirs\n')
  assert.throws(() => guard.ensureDocument(writer.webContents, heldId), /another window/)
  holder.destroy()
  writer.destroy()
})

test('saving into a lossy format asks once, marks the document, and a full-format save clears it', async () => {
  const folder = scenario('lossy')
  const file = path.join(folder, 'Budget.xlsx')
  fs.writeFileSync(file, workbook('original'))
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  const hooks = { serialize: (format) => (format === 'xlsx' ? workbook('saved') : Buffer.from('a,b\n1,2\n')) }
  queue.save.push(path.join(folder, 'Budget.csv'))
  queue.box.push('Save as Comma-separated values (.csv)')
  const lossy = await guard.performSave(eventOf(win), { docId, mode: 'save-as', revision: 3 }, hooks)
  const prompt = record.box.at(-1)
  assert.equal(prompt.message, 'Save "Budget.xlsx" as Comma-separated values (.csv)?')
  assert.deepEqual(prompt.buttons, ['Save as Comma-separated values (.csv)', 'Save as Excel workbook (.xlsx)', 'Cancel'])
  assert.equal(prompt.defaultId, 1, 'the full format is the default')
  assert.match(prompt.detail, /These will be lost: formulas, formatting, other sheets, charts and images\./)
  assert.equal(lossy.ok, true, JSON.stringify(lossy))
  assert.equal(lossy.lossy.format, 'csv')
  assert.equal(lossy.lossy.fullFormat, 'xlsx')
  assert.equal(lossy.lossy.fullLabel, 'Excel workbook (.xlsx)')
  assert.ok(lossy.lossy.lost.includes('formulas'))
  assert.ok(fs.readFileSync(file).equals(workbook('original')), 'the workbook itself is untouched')

  const boxes = record.box.length
  const again = await guard.performSave(eventOf(win), { docId, mode: 'save', revision: 4 }, hooks)
  assert.equal(again.ok, true)
  assert.equal(record.box.length, boxes, 'asked once per format and session')
  assert.equal(again.lossy && again.lossy.format, 'csv', 'the document stays marked until a full-format save')

  queue.save.push(path.join(folder, 'Budget full.xlsx'))
  const full = await guard.performSave(eventOf(win), { docId, mode: 'save-as', revision: 5 }, hooks)
  assert.equal(full.ok, true)
  assert.equal(full.lossy, null)

  // Choosing the full format in the prompt switches the dialog to it.
  const other = path.join(folder, 'Plan.xlsx')
  fs.writeFileSync(other, workbook('plan'))
  const { docId: planId } = await openIn(win, other)
  queue.save.push(path.join(folder, 'Plan.csv'), path.join(folder, 'Plan kept.xlsx'))
  queue.box.push('Save as Excel workbook (.xlsx)')
  const kept = await guard.performSave(eventOf(win), { docId: planId, mode: 'save-as', revision: 1 }, hooks)
  assert.equal(kept.ok, true)
  assert.equal(kept.path, path.join(folder, 'Plan kept.xlsx'))
  assert.equal(kept.lossy, null)
  assert.equal(fs.existsSync(path.join(folder, 'Plan.csv')), false)
  win.destroy()
})

test('a window counts as saving while performSave runs for it', async () => {
  const folder = scenario('saving-flag')
  const file = path.join(folder, 'Slow.csv')
  fs.writeFileSync(file, 'a\n')
  const win = new BrowserWindow()
  const { docId } = await openIn(win, file)
  let release
  const gate = new Promise((resolve) => { release = resolve })
  const pending = guard.performSave(eventOf(win), { docId, mode: 'save', revision: 1 }, { serialize: async () => { await gate; return Buffer.from('b\n') } })
  assert.equal(guard.isSaving(win), true)
  release()
  assert.equal((await pending).ok, true)
  assert.equal(guard.isSaving(win), false)
  assert.deepEqual(record.openPath, [], 'nothing was handed to another program')
  win.destroy()
})

// ---------------------------------------------------------------------------
// Static rules for the new shared files
// ---------------------------------------------------------------------------

test('the new shared files need only built-ins, electron and siblings vendored to the same workspaces', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(SHARED, 'manifest.json'), 'utf8'))
  const files = ['electron/document-guard.cjs', 'electron/io-dialogs.cjs', 'electron/stores.cjs', 'electron/io-ipc.cjs', 'preload/io-bridge.cjs']
  const siblingsOf = {}
  for (const file of files) {
    const source = fs.readFileSync(path.join(SHARED, file), 'utf8')
    const requires = [...source.matchAll(/require\((['"])([^'"]+)\1\)/g)].map((match) => match[2])
    for (const id of requires) assert.ok(id.startsWith('node:') || id === 'electron' || id.startsWith('./'), `${file} requires ${id}`)
    siblingsOf[file] = requires.filter((id) => id.startsWith('./')).map((id) => `electron/${id.slice(2)}`)
    assert.deepEqual(findServiceNames(source), [], `${file} names no cloud service`)
    assert.doesNotMatch(source, /\bcloud\b/i, `${file} names no cloud service`)
    assert.doesNotMatch(source, /https?:\/\//i, `${file} has no remote URL`)
    assert.doesNotMatch(source, /\.openPath\(|openExternal\(|require\(['"]node:(?:http|https|net|tls|dgram)['"]\)|\bfetch\(/, `${file} never hands files to other apps or the network`)
  }
  assert.deepEqual(siblingsOf['preload/io-bridge.cjs'], [], 'the sandboxed preload block requires nothing but electron')
  assert.deepEqual(siblingsOf['electron/io-ipc.cjs'].sort(), ['electron/formats.cjs', 'electron/io-core.cjs', 'electron/io-dialogs.cjs'],
    'modules only some workspaces receive are passed in, never required')
  const closure = (file, seen = new Set()) => {
    if (seen.has(file) || !file.endsWith('.cjs')) return seen
    seen.add(file)
    const source = fs.readFileSync(path.join(SHARED, file), 'utf8')
    for (const match of source.matchAll(/require\((['"])\.\/([^'"]+)\1\)/g)) {
      const sibling = `electron/${match[2]}`
      if (sibling.endsWith('.json')) seen.add(sibling)
      else closure(sibling, seen)
    }
    return seen
  }
  for (const [name, workspace] of Object.entries(manifest.workspaces)) {
    for (const file of workspace.electron) {
      if (!files.includes(file)) continue
      for (const needed of closure(file)) assert.ok(workspace.electron.includes(needed), `${name} vendors ${file} but not ${needed}`)
    }
    if (workspace.electron.includes('electron/document-guard.cjs')) {
      assert.ok(workspace.wiring.some((rule) => rule.call === 'installWindowGuard('), `${name} must wire installWindowGuard (it also guards quitting)`)
    }
    assert.ok(workspace.wiring.some((rule) => rule.call === 'registerSharedIo('), `${name} must wire registerSharedIo`)
    assert.equal(workspace.preload, true, `${name} receives the preload bridge`)
  }
  assert.equal(manifest.preloadBlock.source, 'preload/io-bridge.cjs')
})
