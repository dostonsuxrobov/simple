// Vendored from simple/shared/electron/io-ipc.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// registerSharedIo(): the io:* IPC handlers behind window.simpleIO (the
// preload bridge in simple/shared/preload/io-bridge.cjs). Every handler first
// checks its sender: the main frame of one of the app's own pages (a file:
// page inside the app, or the Vite dev server in development).
//
// io:open-in-simple hands a file to the Simple workspace that opens it: the
// workspace's own new window when it is this one, otherwise the unified app's
// own executable (its single-instance routing passes the path on). Files that
// no Simple workspace opens are only shown in File Explorer. Documents are
// never handed to the operating system's default app.
//
// Modules that only some workspaces receive (document-guard, stores,
// office-engine) are passed in by the workspace, so this file only requires
// files every workspace has. Only Node built-ins and electron are required.

const fs = require('node:fs')
const path = require('node:path')
const { spawn } = require('node:child_process')
const { fileURLToPath } = require('node:url')
const core = require('./io-core.cjs')
const formats = require('./formats.cjs')
const dialogs = require('./io-dialogs.cjs')

/** Workspaces that can register the shared IPC. */
const MODULES = Object.freeze(['pdf', 'calc', 'docs', 'image', 'video', 'launcher'])
/** Version of the window.simpleIO contract. */
const BRIDGE_VERSION = 1
/** Every channel registerSharedIo handles (renderer → main, invoke). */
const CHANNELS = Object.freeze([
  'io:capabilities',
  'io:engine-status',
  'io:choose-save-path',
  'io:choose-open-paths',
  'io:prompt',
  'io:recovery-write',
  'io:recovery-list',
  'io:recovery-read',
  'io:recovery-adopt',
  'io:recovery-discard',
  'io:versions-list',
  'io:versions-open',
  'io:clipboard-read',
  'io:shell-show-item',
  'io:open-in-simple',
  'io:prefs-get',
  'io:prefs-set',
])
const MAX_CLIPBOARD_BYTES = 64 * 1024 * 1024
const MAX_VAR_CHARS = 2000
const PROMPT_KEY_PATTERN = /^(saveFailed|prompts)\.[A-Za-z0-9_-]+$/
/**
 * Prompt answers that let a later save request skip a safety check. Main
 * records them when it shows the prompt (io:prompt) and performSave honours
 * the matching request flag only after such an answer.
 */
const DECISION_ANSWERS = Object.freeze({
  'saveFailed.CHANGED_ON_DISK': 'replace',
  'saveFailed.SOURCE_MISSING': 'recreate',
  'prompts.lossy-save': 'save-lossy',
  'prompts.sibling-save': 'save-sibling',
})
const UNIFIED_PACKAGE_NAME = 'simple-unified'

let logger = {
  warn: (...args) => console.warn('[simple-io]', ...args),
}

const registrations = new WeakMap()

function electron() {
  try {
    const value = require('electron')
    return value && typeof value === 'object' ? value : null
  } catch {
    return null
  }
}

/** Office engine status for a workspace that never uses the engine. */
function noEngineStatus() {
  return {
    available: false,
    source: null,
    path: null,
    version: null,
    verified: false,
    reason: 'not-installed',
    detail: 'This workspace does not use an office engine.',
    checkedAt: Date.now(),
  }
}

// ---------------------------------------------------------------------------
// Trust
// ---------------------------------------------------------------------------

function defaultFileRoot() {
  const value = electron()
  const app = value && value.app
  if (!app || typeof app.getAppPath !== 'function') return null
  try {
    return path.resolve(app.getAppPath())
  } catch {
    return null
  }
}

function isInsideFolder(child, parent) {
  const relative = path.relative(parent, child)
  if (!relative) return false
  if (process.platform === 'win32' ? relative.toLowerCase().startsWith('..') : relative.startsWith('..')) return false
  return !path.isAbsolute(relative)
}

/**
 * Whether an IPC event comes from the app's own page: the main frame, loaded
 * from a file: URL inside the app folder (or from the Vite dev server in
 * development). Print and conversion windows, iframes and remote pages fail.
 *
 * @param {{sender?: object, senderFrame?: object}} event
 * @param {((event: object) => boolean)|{fileRoot?: string|null, devServerUrl?: string|null}} [trust]
 *   a custom check, or overrides: fileRoot null allows any file: page; devServerUrl defaults to VITE_DEV_SERVER_URL
 * @returns {boolean}
 */
function isTrustedSender(event, trust = {}) {
  if (typeof trust === 'function') {
    try {
      return Boolean(trust(event))
    } catch {
      return false
    }
  }
  const frame = event && event.senderFrame
  const sender = event && event.sender
  if (!frame || !sender || frame !== sender.mainFrame) return false
  let url
  try {
    url = new URL(String(frame.url))
  } catch {
    return false
  }
  const devServerUrl = trust && Object.prototype.hasOwnProperty.call(trust, 'devServerUrl') ? trust.devServerUrl : process.env.VITE_DEV_SERVER_URL
  if (devServerUrl) {
    try {
      const allowed = new URL(devServerUrl)
      if ((url.protocol === 'http:' || url.protocol === 'https:') && url.origin === allowed.origin) return true
    } catch {}
  }
  if (url.protocol !== 'file:') return false
  const root = trust && Object.prototype.hasOwnProperty.call(trust, 'fileRoot') ? trust.fileRoot : defaultFileRoot()
  if (!root) return true
  let filePath
  try {
    filePath = fileURLToPath(url)
  } catch {
    return false
  }
  return isInsideFolder(path.resolve(filePath), path.resolve(root))
}

/**
 * Throws unless the event comes from the app's own page (see isTrustedSender).
 * @param {object} event
 * @param {Parameters<typeof isTrustedSender>[1]} [trust]
 */
function assertTrustedSender(event, trust) {
  if (!isTrustedSender(event, trust)) throw new Error('Untrusted request.')
}

// ---------------------------------------------------------------------------
// Input checks
// ---------------------------------------------------------------------------

function checkFilePathArgument(value) {
  if (typeof value !== 'string' || !value || value.length > 32_000 || value.includes('\u0000')) {
    throw new TypeError('A file path is required.')
  }
  if (!path.isAbsolute(value)) throw new TypeError('The file path must be absolute.')
  // Device and raw namespaces (\\?\, \\.\) are never document paths.
  if (/^[\\/]{2}[?.][\\/]/.test(value)) throw new TypeError('This kind of path is not supported.')
  return path.resolve(value)
}

function cleanVars(vars) {
  const result = {}
  if (!vars || typeof vars !== 'object') return result
  for (const [key, value] of Object.entries(vars)) {
    if (!/^[A-Za-z][A-Za-z0-9]{0,39}$/.test(key)) continue
    if (typeof value === 'string') result[key] = value.slice(0, MAX_VAR_CHARS)
    else if (typeof value === 'number' && Number.isFinite(value)) result[key] = String(value)
  }
  return result
}

function toUint8(buffer) {
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength)
}

/** The catalog's wording for a document another Simple window is showing. */
function shownElsewhere(name) {
  return core.formatTemplate(core.catalog.notices.OPEN_IN_ANOTHER_WINDOW, { name: typeof name === 'string' && name ? name : undefined })
}

function publicEntry(entry) {
  if (!entry) return null
  const { folder: _folder, ...rest } = entry
  return rest
}

// ---------------------------------------------------------------------------
// Opening files inside Simple
// ---------------------------------------------------------------------------

/**
 * Whether this process is the unified Simple app (one EXE for every
 * workspace) rather than a standalone workspace build.
 * @returns {boolean}
 */
function isUnifiedApp() {
  const value = electron()
  const app = value && value.app
  if (!app || typeof app.getAppPath !== 'function') return false
  try {
    const manifest = JSON.parse(fs.readFileSync(path.join(app.getAppPath(), 'package.json'), 'utf8').replace(/^\uFEFF/, ''))
    return Boolean(manifest && manifest.name === UNIFIED_PACKAGE_NAME)
  } catch {
    return false
  }
}

function portableExecutable() {
  const value = process.env.PORTABLE_EXECUTABLE_FILE
  if (!value || !path.isAbsolute(value) || !/\.exe$/i.test(value)) return null
  try {
    return fs.statSync(value).isFile() ? path.resolve(value) : null
  } catch {
    return null
  }
}

/**
 * The command that starts this app again with some file paths, as the
 * unified app's own launcher does: the portable executable when Simple runs
 * from the portable EXE, else process.execPath (plus the app folder when the
 * app is not packaged). Only absolute file paths are ever passed, never flags.
 * @param {string[]} filePaths absolute paths
 * @returns {{command: string, args: string[]}}
 */
function selfLaunchCommand(filePaths) {
  const value = electron()
  const app = value && value.app
  const files = filePaths.map((filePath) => checkFilePathArgument(filePath))
  const portable = portableExecutable()
  if (portable) return { command: portable, args: files }
  const packaged = Boolean(app && app.isPackaged)
  return { command: process.execPath, args: packaged || !app ? files : [path.resolve(app.getAppPath()), ...files] }
}

function launchDetached(command, args) {
  const child = spawn(command, args, { detached: true, stdio: 'ignore', windowsHide: true })
  child.on('error', (error) => logger.warn(`Could not start Simple for ${args[args.length - 1]}: ${error.message}`))
  child.unref()
}

/**
 * Opens a file inside Simple. In order:
 * 1. a window of this process already shows it: focus that window;
 * 2. this workspace opens it and `openInWindow` was given: open it there;
 * 3. the unified app: start the app's own executable with the path, so its
 *    single-instance routing hands the file to the workspace that opens it;
 * 4. otherwise (no Simple workspace opens it, or a standalone build cannot
 *    reach the other workspace): show it in File Explorer.
 * Never hands the file to the operating system's default app.
 *
 * @param {string} filePath absolute path
 * @param {object} [options]
 * @param {string} [options.module] this workspace
 * @param {object} [options.guard] the document-guard module, when this workspace has one
 * @param {(filePath: string, info: {mode: string, route: object, sender?: object}) => (boolean|void|Promise<boolean|void>)} [options.openInWindow]
 * @param {boolean} [options.unified] override the unified-app detection
 * @param {(command: string, args: string[]) => void} [options.launch] starts the app (tests replace it)
 * @param {object} [options.sender] the calling webContents
 * @returns {Promise<{ok: true, action: 'focused'|'opened'|'launched'|'shown-in-folder', shownInFolder: boolean,
 *   mode: string|null, appName: string|null, path: string, reason?: 'unsupported'|'other-workspace'}
 *   | {ok: false, code: string, message: string, path: string}>}
 */
async function openInSimple(filePath, options = {}) {
  const target = checkFilePathArgument(filePath)
  let stat
  try {
    stat = await fs.promises.stat(target)
  } catch (error) {
    return { ...core.toIoResult(error, { path: target, stage: 'read', context: 'open' }), path: target }
  }
  if (!stat.isFile()) {
    return { ok: false, code: 'UNSUPPORTED', context: 'open', message: core.formatTemplate(core.catalog.notices.NOT_A_FILE, { name: path.basename(target) }), path: target }
  }
  // The same routing the unified app's bootstrap applies to its command line
  // (extension, plus content for .txt/.html/.xml), so a relaunch always reaches a
  // workspace. The content sample is read without blocking this process.
  const route = await formats.routeForPathAsync(target)
  const mode = route.mode || null
  const appName = mode ? dialogs.appNameFor(mode) : null
  const guard = options.guard
  if (guard && typeof guard.findDocumentByPath === 'function') {
    const open = guard.findDocumentByPath(target)
    if (open && typeof guard.focusDocument === 'function' && guard.focusDocument(open.docId)) {
      const here = options.module || mode
      return { ok: true, action: 'focused', shownInFolder: false, mode: here, appName: dialogs.appNameFor(here), path: target }
    }
  }
  if (mode && mode === options.module && typeof options.openInWindow === 'function') {
    const opened = await options.openInWindow(target, { mode, route, sender: options.sender })
    if (opened !== false) return { ok: true, action: 'opened', shownInFolder: false, mode, appName, path: target }
  }
  const unified = typeof options.unified === 'boolean' ? options.unified : isUnifiedApp()
  if (mode && unified) {
    const { command, args } = selfLaunchCommand([target])
    ;(typeof options.launch === 'function' ? options.launch : launchDetached)(command, args)
    return { ok: true, action: 'launched', shownInFolder: false, mode, appName, path: target }
  }
  const value = electron()
  if (value && value.shell && typeof value.shell.showItemInFolder === 'function') value.shell.showItemInFolder(target)
  return { ok: true, action: 'shown-in-folder', shownInFolder: true, mode, appName, path: target, reason: mode ? 'other-workspace' : 'unsupported' }
}

// ---------------------------------------------------------------------------
// Clipboard
// ---------------------------------------------------------------------------

/**
 * Reads the system clipboard for a paste: a copied file (Windows FileNameW,
 * first file only), an image as PNG bytes, HTML, RTF and text. Each part is
 * left out above 64 MB.
 * @returns {{files: string[], png?: Uint8Array, html?: string, rtf?: string, text?: string}}
 */
function readClipboard() {
  const value = electron()
  const clipboard = value && value.clipboard
  const result = { files: [] }
  if (!clipboard) return result
  const textPart = (read) => {
    try {
      const text = read()
      return typeof text === 'string' && text && Buffer.byteLength(text) <= MAX_CLIPBOARD_BYTES ? text : undefined
    } catch {
      return undefined
    }
  }
  const text = textPart(() => clipboard.readText())
  if (text !== undefined) result.text = text
  const html = textPart(() => clipboard.readHTML())
  if (html !== undefined) result.html = html
  const rtf = textPart(() => clipboard.readRTF())
  if (rtf !== undefined) result.rtf = rtf
  try {
    const image = clipboard.readImage()
    if (image && typeof image.isEmpty === 'function' && !image.isEmpty()) {
      const png = image.toPNG()
      if (png && png.length && png.length <= MAX_CLIPBOARD_BYTES) result.png = toUint8(png)
    }
  } catch {}
  if (process.platform === 'win32' && typeof clipboard.readBuffer === 'function') {
    try {
      const raw = clipboard.readBuffer('FileNameW')
      if (raw && raw.length) {
        const name = raw.toString('utf16le').split('\u0000')[0]
        if (name && path.isAbsolute(name)) result.files.push(path.resolve(name))
      }
    } catch {}
  }
  return result
}

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

/**
 * Registers the io:* handlers behind window.simpleIO for one workspace.
 * Call it once in app.whenReady(), before windows open.
 *
 *   registerSharedIo({ ipcMain, module: 'calc', guard, stores, officeEngine,
 *     openInWindow: (filePath) => createWindow(filePath) })
 *
 * @param {object} options
 * @param {object} [options.ipcMain] default: electron's
 * @param {'pdf'|'calc'|'docs'|'image'|'video'|'launcher'} options.module
 * @param {Parameters<typeof isTrustedSender>[1]} [options.trust] sender check overrides
 * @param {object} [options.guard] the workspace's vendored simple-io/document-guard.cjs module (not in Video)
 * @param {object} [options.stores] the workspace's vendored simple-io/stores.cjs module (not in Video)
 * @param {object} [options.officeEngine] the workspace's vendored simple-io/office-engine.cjs module (PDF, Calc, Docs)
 * @param {(filePath: string, info: object) => (boolean|void|Promise<boolean|void>)} [options.openInWindow]
 *   opens a file of this workspace in a new window (open-in-Simple, standalone builds)
 * @param {(version: {path: string, name: string, createdAt: string, size: number, sourcePath: string|null},
 *   info: {sender: object}) => any} [options.openVersion] opens a stored version as an untitled copy
 * @param {boolean} [options.unified] override the unified-app detection
 * @param {(command: string, args: string[]) => void} [options.launch] replaces process start-up (tests)
 * @param {boolean} [options.maintenance=true] with `stores`: migrate the Documents workspace's old recovery
 *   layout and prune both stores in the background
 * @returns {{channels: readonly string[], maintenance: Promise<void>, dispose: () => void,
 *   openInSimple: (filePath: string) => ReturnType<typeof openInSimple>}}
 */
function registerSharedIo(options = {}) {
  const value = electron()
  const ipcMain = options.ipcMain || (value && value.ipcMain)
  if (!ipcMain || typeof ipcMain.handle !== 'function') throw new TypeError('registerSharedIo needs ipcMain.')
  const module = String(options.module || '')
  if (!MODULES.includes(module)) throw new TypeError(`registerSharedIo needs a workspace module (${MODULES.join(', ')}).`)
  if (registrations.has(ipcMain)) throw new Error('registerSharedIo was already called in this process.')
  const trust = options.trust || {}
  const guard = options.guard || null
  const storesModule = options.stores || null
  const engine = options.officeEngine || null
  dialogs.configureDialogs({ module })
  if (guard && typeof guard.configureDocumentGuard === 'function') guard.configureDocumentGuard({ module })
  // Lets the sandboxed preload of windows created from now on know its workspace.
  process.env.SIMPLE_IO_MODULE = module

  const recovery = () => (storesModule && typeof storesModule.recoveryStore === 'function' ? storesModule.recoveryStore() : null)
  const versions = () => (storesModule && typeof storesModule.versionsStore === 'function' ? storesModule.versionsStore() : null)
  const engineStatus = async () => {
    if (!engine || typeof engine.getOfficeEngineStatus !== 'function') return noEngineStatus()
    try {
      return { ...(await engine.getOfficeEngineStatus()) }
    } catch (error) {
      return { ...noEngineStatus(), detail: `The office engine check failed: ${error && error.message}` }
    }
  }
  const engineAvailable = async () => Boolean((await engineStatus()).available)
  const ownDocument = (sender, docId) => {
    if (!guard || typeof docId !== 'string' || !docId) return null
    const doc = guard.getDocument(docId)
    return doc && doc.webContentsId === sender.id ? doc : null
  }
  const openOptions = (sender) => ({
    module,
    guard,
    openInWindow: options.openInWindow,
    unified: options.unified,
    launch: options.launch,
    sender,
  })

  const channels = []
  const handle = (channel, handler) => {
    ipcMain.handle(channel, (event, ...args) => {
      assertTrustedSender(event, trust)
      return handler(event, ...args)
    })
    channels.push(channel)
  }

  handle('io:capabilities', async () => ({
    version: BRIDGE_VERSION,
    module,
    appName: dialogs.appNameFor(module),
    platform: process.platform,
    unified: typeof options.unified === 'boolean' ? options.unified : isUnifiedApp(),
    officeEngine: await engineStatus(),
  }))

  handle('io:engine-status', () => engineStatus())

  handle('io:choose-save-path', async (event, request) => {
    const req = request && typeof request === 'object' ? request : {}
    const doc = ownDocument(event.sender, req.docId)
    return dialogs.chooseSavePath({
      purpose: req.purpose,
      name: typeof req.name === 'string' ? req.name.slice(0, 260) : (doc && doc.name) || undefined,
      format: typeof req.format === 'string' ? req.format : (doc && doc.format) || undefined,
      formats: Array.isArray(req.formats) ? req.formats.filter((id) => typeof id === 'string') : undefined,
      reason: typeof req.reason === 'string' ? req.reason : undefined,
    }, {
      sender: event.sender,
      module,
      sourcePath: doc ? (doc.path || doc.suggestedPath) : null,
      engine: await engineAvailable(),
    })
  })

  handle('io:choose-open-paths', async (event, request) => {
    const req = request && typeof request === 'object' ? request : {}
    return dialogs.chooseOpenPaths({ multi: req.multi === true, purpose: req.purpose }, { sender: event.sender, module, engine: await engineAvailable() })
  })

  handle('io:prompt', async (event, key, vars) => {
    if (typeof key !== 'string' || !PROMPT_KEY_PATTERN.test(key)) throw new TypeError('Unknown prompt.')
    const values = cleanVars(vars)
    const answer = await dialogs.showPrompt(key, values, { sender: event.sender, module })
    if (Object.prototype.hasOwnProperty.call(DECISION_ANSWERS, key) && answer === DECISION_ANSWERS[key]) {
      dialogs.recordDecision(event.sender, key, answer, typeof values.path === 'string' && path.isAbsolute(values.path) ? path.resolve(values.path) : undefined)
    }
    return answer
  })

  handle('io:recovery-write', async (event, snapshot) => {
    const store = recovery()
    if (!store) return { ok: false, code: 'UNSUPPORTED', technical: 'this workspace keeps no recovery copies' }
    if (!snapshot || typeof snapshot !== 'object' || !Array.isArray(snapshot.parts)) throw new TypeError('A recovery snapshot is required.')
    const docId = String(snapshot.docId || '')
    let doc = null
    if (guard) {
      try {
        doc = guard.ensureDocument(event.sender, docId, { name: typeof snapshot.title === 'string' ? snapshot.title : undefined, format: typeof snapshot.format === 'string' ? snapshot.format : undefined })
      } catch (error) {
        return { ok: false, code: 'LOCKED', reason: 'open-in-another-window', message: shownElsewhere(snapshot.title), technical: error.message, docId }
      }
    }
    const parts = []
    for (const part of snapshot.parts) {
      if (!part || typeof part !== 'object') continue
      if (part.ref !== undefined) {
        // A page may refer only to its document's own file ('source'), never to
        // any other path: main would read and hash whatever file it named.
        if (part.ref !== 'source') throw new TypeError('A recovery part can only refer to the document\'s own file ("source").')
        if (!doc || !doc.path) continue
        // The stamp is the one main took when it read the file, never one from the page.
        parts.push({ name: part.name, ref: { path: doc.path, stamp: doc.stamp }, required: part.required === true })
        continue
      }
      parts.push({ name: part.name, data: part.data, compress: part.compress === true })
    }
    return store.write({
      docId,
      revision: snapshot.revision,
      parts,
      module,
      title: typeof snapshot.title === 'string' ? snapshot.title : (doc && doc.name) || null,
      kind: typeof snapshot.kind === 'string' ? snapshot.kind : null,
      format: typeof snapshot.format === 'string' ? snapshot.format : (doc && doc.format) || null,
      sourcePath: doc ? doc.path : null,
      sourceStamp: doc ? doc.stamp : null,
      suggestedPath: doc ? doc.suggestedPath : null,
      extra: snapshot.meta !== undefined ? snapshot.meta : snapshot.extra,
    })
  })

  handle('io:recovery-list', async () => {
    const store = recovery()
    if (!store) return []
    return (await store.list()).map(publicEntry)
  })

  // Reading an entry changes nothing: the window takes it over (io:recovery-adopt)
  // only after its page rebuilt the document, so a failed restore leaves no
  // document bound to the entry's file and the entry stays on offer.
  const readEntries = new Map()
  const adoptionFor = (entry) => ({
    docId: entry.docId,
    path: entry.sourcePath,
    stamp: entry.sourceStamp,
    format: entry.format,
    name: entry.title,
    suggestedPath: entry.suggestedPath,
  })

  handle('io:recovery-read', async (event, id) => {
    const store = recovery()
    if (!store) return { ok: false, code: 'UNSUPPORTED', technical: 'this workspace keeps no recovery copies' }
    const result = await store.read(String(id || ''))
    if (!result.ok) return result
    const entry = result.entry
    if (guard && typeof guard.canAdoptDocument === 'function' && !guard.canAdoptDocument(event.sender, adoptionFor(entry))) {
      return { ok: false, code: 'LOCKED', reason: 'open-in-another-window', message: shownElsewhere(entry.title), technical: 'another window shows this document or its file', docId: entry.docId }
    }
    const senderId = event.sender.id
    if (!readEntries.has(senderId)) {
      readEntries.set(senderId, new Map())
      if (typeof event.sender.once === 'function') event.sender.once('destroyed', () => readEntries.delete(senderId))
    }
    readEntries.get(senderId).set(entry.docId, adoptionFor(entry))
    const parts = result.parts.map((part) => (part.ref
      ? { name: part.name, ref: { path: part.ref.path, stamp: part.ref.stamp, matches: part.ref.matches }, required: part.required === true }
      : { name: part.name, data: typeof part.data === 'string' ? part.data : toUint8(part.data) }))
    return {
      ok: true,
      docId: entry.docId,
      entry: publicEntry(entry),
      revision: result.revision,
      generation: result.generation,
      fellBack: result.fellBack,
      parts,
      meta: result.extra,
    }
  })

  handle('io:recovery-adopt', async (event, id) => {
    const store = recovery()
    if (!store) return { ok: false, code: 'UNSUPPORTED', technical: 'this workspace keeps no recovery copies' }
    const docId = String(id || '')
    const pending = readEntries.get(event.sender.id)
    const adoption = pending && pending.get(docId)
    if (!adoption) return { ok: false, code: 'NOT_FOUND', technical: 'this window has not read that recovery entry' }
    let doc = null
    if (guard) {
      try {
        // The restoring window takes the entry over: same id, the original path,
        // and the stamp from snapshot time, so a changed file is never overwritten silently.
        doc = guard.adoptDocument(event.sender, adoption)
      } catch (error) {
        return { ok: false, code: 'LOCKED', reason: 'open-in-another-window', message: shownElsewhere(adoption.name), technical: error.message, docId }
      }
    }
    pending.delete(docId)
    await store.claim(docId).catch(() => false)
    return { ok: true, docId, untitled: doc ? doc.untitled : !adoption.path, path: doc ? doc.path : (adoption.path || null) }
  })

  handle('io:recovery-discard', async (event, id, upToRevision) => {
    const store = recovery()
    if (!store) return { ok: false, code: 'UNSUPPORTED', technical: 'this workspace keeps no recovery copies' }
    const docId = String(id || '')
    if (guard) {
      const doc = guard.getDocument(docId)
      if (doc && doc.webContentsId !== event.sender.id) {
        return { ok: false, code: 'LOCKED', reason: 'open-in-another-window', message: shownElsewhere(doc.name), docId }
      }
    }
    return store.discard(docId, upToRevision === undefined || upToRevision === null ? undefined : upToRevision)
  })

  handle('io:versions-list', async (_event, filePath) => {
    const store = versions()
    if (!store) return []
    return store.list(checkFilePathArgument(filePath))
  })

  handle('io:versions-open', async (event, id) => {
    const store = versions()
    if (!store) return { ok: false, code: 'UNSUPPORTED', technical: 'this workspace keeps no earlier versions' }
    const version = await store.resolve(String(id || ''))
    if (!version) return { ok: false, code: 'NOT_FOUND', message: core.formatTemplate(core.catalog.versions.unavailable, {}) }
    if (typeof options.openVersion !== 'function') return { ok: false, code: 'UNSUPPORTED', message: core.formatTemplate(core.catalog.versions.unavailable, {}) }
    await options.openVersion(version, { sender: event.sender })
    return { ok: true, name: version.name, createdAt: version.createdAt }
  })

  handle('io:clipboard-read', () => readClipboard())

  handle('io:shell-show-item', (_event, filePath) => {
    const target = checkFilePathArgument(filePath)
    const shell = value && value.shell
    if (shell && typeof shell.showItemInFolder === 'function') shell.showItemInFolder(target)
    return { ok: true }
  })

  handle('io:open-in-simple', (event, filePath) => openInSimple(filePath, openOptions(event.sender)))

  handle('io:prefs-get', (_event, key) => dialogs.getPref(key))

  handle('io:prefs-set', (_event, key, prefValue) => dialogs.setPref(key, prefValue))

  // Office engine changes (installed or removed outside Simple) reach open pages.
  let lastAvailable = null
  let focusListener = null
  const app = value && value.app
  if (engine && app && typeof app.on === 'function') {
    focusListener = () => {
      engineStatus().then((status) => {
        const available = Boolean(status.available)
        if (lastAvailable !== null && available !== lastAvailable) {
          const windows = value.BrowserWindow && typeof value.BrowserWindow.getAllWindows === 'function' ? value.BrowserWindow.getAllWindows() : []
          for (const win of windows) {
            try {
              if (!win.isDestroyed()) win.webContents.send('io:capabilities-changed')
            } catch {}
          }
        }
        lastAvailable = available
      }).catch(() => {})
    }
    app.on('browser-window-focus', focusListener)
    focusListener()
  }

  // Startup upkeep of the stores, in the background: the Documents workspace's old
  // recovery layout moves into the journal, then old offered entries, damaged
  // entries and expired versions are pruned. Nothing that was never offered is removed.
  const maintenance = storesModule && options.maintenance !== false
    ? (async () => {
      const store = recovery()
      if (store && module === 'docs') await store.migrateDocsLegacy()
      if (store) await store.prune()
      const versionStore = versions()
      if (versionStore) await versionStore.prune()
    })().catch((error) => logger.warn(`Store upkeep failed: ${error && error.message}`))
    : Promise.resolve()

  const registration = {
    channels: Object.freeze([...channels]),
    /** Resolves when the startup upkeep of the stores has finished. */
    maintenance,
    openInSimple: (filePath) => openInSimple(filePath, openOptions(null)),
    dispose() {
      for (const channel of channels) {
        try { ipcMain.removeHandler(channel) } catch {}
      }
      if (focusListener && app && typeof app.removeListener === 'function') app.removeListener('browser-window-focus', focusListener)
      registrations.delete(ipcMain)
    },
  }
  registrations.set(ipcMain, registration)
  return registration
}

/**
 * The `webPreferences.additionalArguments` that tell the preload bridge its
 * workspace (window.simpleIO.module).
 * @param {string} module
 * @returns {string[]}
 */
function bridgeArguments(module) {
  if (!MODULES.includes(module)) throw new TypeError(`Unknown workspace "${module}".`)
  return [`--simple-io-module=${module}`]
}

/**
 * Configures where io-ipc logs problems.
 * @param {{logger?: {warn: Function}}} [options]
 */
function configureIoIpc(options = {}) {
  if (options.logger && typeof options.logger.warn === 'function') logger = { warn: options.logger.warn }
}

module.exports = {
  BRIDGE_VERSION,
  CHANNELS,
  MODULES,
  assertTrustedSender,
  bridgeArguments,
  configureIoIpc,
  isTrustedSender,
  isUnifiedApp,
  openInSimple,
  readClipboard,
  registerSharedIo,
  selfLaunchCommand,
}
