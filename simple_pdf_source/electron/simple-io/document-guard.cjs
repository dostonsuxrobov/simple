// Vendored from simple/shared/electron/document-guard.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// The main-process half of "an edit is never silently lost":
// - the document registry: which window shows which file, the stamp (size,
//   mtime, SHA-256) main took when it read the file, and one window per file;
// - performSave(): the one Save / Save As / Save a Copy / Export pipeline.
//   It resolves the target (dialogs, sibling files, lossy formats), serializes,
//   copies the original to the versions store, writes through safeWriteFile and
//   rebinds the document. It always resolves to an IoResult and never throws
//   for expected failures;
// - the renderer RPC used by the guards (close-query, save-now, discard,
//   recovery-flush) through the preload bridge;
// - installWindowGuard(): Save / Don't Save / Cancel when a window closes,
//   never closing during a save, and crashed, hung and shutting-down windows;
// - installAppGuard(): quitting waits for pending writes and asks each window
//   in turn.
// Only Node built-ins, electron (loaded lazily) and sibling files are required.

const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')
const core = require('./io-core.cjs')
const formats = require('./formats.cjs')
const dialogs = require('./io-dialogs.cjs')
const stores = require('./stores.cjs')
const { safeWriteFile } = require('./safe-write.cjs')

const { IoError } = core
const DOC_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9_-]{0,79}$/
const SAVE_MODES = Object.freeze(['save', 'save-as', 'save-copy', 'export'])
/** Requests main sends to a window's page through the preload bridge. */
const REQUEST_TYPES = Object.freeze(['close-query', 'save-now', 'discard', 'recovery-flush'])
const MAX_CLOSE_ROUNDS = 25
const SAVE_WAIT_ROUND_MS = 5 * 60 * 1000

const settings = {
  module: null,
  keepBytesLimit: 32 * 1024 * 1024,
  queryTimeoutMs: 5000,
  discardTimeoutMs: 5000,
  flushTimeoutMs: 2000,
  pollMs: 250,
  pendingWritesTimeoutMs: 30_000,
  // An approved quit that has not reached app 'quit' by then was stopped elsewhere.
  quitConfirmMs: 10_000,
}

let logger = {
  warn: (...args) => console.warn('[simple-io]', ...args),
}

/**
 * Configures the guard for this process (one workspace per process).
 * @param {object} [options]
 * @param {string} [options.module] 'pdf' | 'calc' | 'docs' | 'image' | 'video'
 * @param {number} [options.keepBytesLimit] files up to this size keep their opened bytes for the pristine shortcut (32 MB)
 * @param {number} [options.queryTimeoutMs] how long a page may take to answer close-query (5 s)
 * @param {number} [options.discardTimeoutMs] how long Don't Save waits for the page (5 s)
 * @param {number} [options.flushTimeoutMs] recovery flush budget at Windows shutdown (2 s)
 * @param {number} [options.pollMs] re-check interval while a page reports a save in progress
 * @param {number} [options.pendingWritesTimeoutMs] how long quitting waits for writes (30 s)
 * @param {number} [options.quitConfirmMs] after an approved quit that never happened, guards switch back on (10 s)
 * @param {{warn: Function}} [options.logger]
 */
function configureDocumentGuard(options = {}) {
  if (options.module !== undefined) settings.module = options.module ? String(options.module) : null
  for (const key of ['keepBytesLimit', 'queryTimeoutMs', 'discardTimeoutMs', 'flushTimeoutMs', 'pollMs', 'pendingWritesTimeoutMs', 'quitConfirmMs']) {
    if (options[key] !== undefined && Number.isFinite(Number(options[key])) && Number(options[key]) >= 0) settings[key] = Number(options[key])
  }
  if (options.logger && typeof options.logger.warn === 'function') logger = { warn: options.logger.warn }
}

function electron() {
  try {
    const value = require('electron')
    return value && typeof value === 'object' ? value : null
  } catch {
    return null
  }
}

function isDestroyed(target) {
  try {
    return Boolean(target && typeof target.isDestroyed === 'function' && target.isDestroyed())
  } catch {
    return true
  }
}

function windowOf(webContents) {
  const value = electron()
  if (!webContents || !value || !value.BrowserWindow || typeof value.BrowserWindow.fromWebContents !== 'function') return null
  try {
    const win = value.BrowserWindow.fromWebContents(webContents)
    return win && !isDestroyed(win) ? win : null
  } catch {
    return null
  }
}

function contentsOf(target) {
  if (!target) return null
  if (target.webContents && typeof target.webContents.send === 'function') return target.webContents
  if (typeof target.send === 'function' && typeof target.id === 'number') return target
  return null
}

function formatLabel(formatId) {
  const entry = formatId ? formats.formatById(formatId) : null
  if (!entry) return null
  return `${entry.label} (${entry.extensions[0]})`
}

function formatOfPath(filePath) {
  const entry = formats.formatForExtension(path.extname(String(filePath || '')).toLowerCase())
  return entry ? entry.id : null
}

// ---------------------------------------------------------------------------
// Document registry
// ---------------------------------------------------------------------------

const documents = new Map()
const watchedContents = new WeakSet()
const activeSaves = new Map()

/**
 * The registry key of a file: its real path, lower-cased on Windows, so
 * "C:\\Work\\Budget.xlsx" and "c:\\work\\BUDGET.XLSX" (or a short 8.3 name) match.
 * @param {string} filePath
 * @returns {string}
 */
function docKey(filePath) {
  return stores.fileKey(filePath)
}

function checkDocId(value) {
  const id = String(value ?? '')
  if (!DOC_ID_PATTERN.test(id)) throw new TypeError(`"${id.slice(0, 40)}" is not a valid document id.`)
  return id
}

function publicDoc(doc) {
  if (!doc) return null
  return {
    docId: doc.docId,
    path: doc.path,
    name: doc.name,
    format: doc.format,
    sourceFormat: doc.sourceFormat,
    stamp: doc.stamp ? { ...doc.stamp } : null,
    untitled: !doc.path,
    suggestedPath: doc.suggestedPath || null,
    lossy: doc.lossy ? { ...doc.lossy } : null,
    webContentsId: doc.webContentsId,
  }
}

function watchContents(webContents) {
  if (watchedContents.has(webContents) || typeof webContents.once !== 'function') return
  watchedContents.add(webContents)
  const id = webContents.id
  webContents.once('destroyed', () => {
    for (const doc of [...documents.values()]) if (doc.webContentsId === id) documents.delete(doc.docId)
    activeSaves.delete(id)
    dialogs.revokeGrants(id)
    failRequestsFor(id, 'gone')
  })
}

function keepBytes(doc, bytes, sha256) {
  if (bytes && bytes.length <= settings.keepBytesLimit) {
    doc.originalBytes = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    doc.originalSha256 = sha256 || core.hashBytes(doc.originalBytes)
  } else {
    doc.originalBytes = null
    doc.originalSha256 = null
  }
}

/**
 * Adds (or updates) a document shown by a window. Open handlers call this
 * with the bytes and stamp main read (see openDocument, which does both).
 * The stamp always comes from main, never from the renderer. A path given
 * without a stamp is stamped here (in the background; saving waits for it),
 * so a document bound to a file always detects changes by other programs.
 *
 * @param {object} input
 * @param {object} input.webContents the window's webContents (or `window`, or `sender`)
 * @param {string} [input.docId] keep an id (recovery restore); default a new UUID
 * @param {string|null} [input.path] the file, or null for an untitled document
 * @param {{size: number, mtimeMs: number, sha256: string}|null} [input.stamp]
 * @param {Buffer|Uint8Array} [input.bytes] the opened bytes (kept up to 32 MB for the pristine shortcut)
 * @param {string} [input.format] registry format id (default: from the extension)
 * @param {string} [input.name] display name (default: the file name)
 * @param {string|null} [input.suggestedPath] where Save As should start for an untitled document
 * @param {boolean} [input.mustWrite] the window shows content that is not in the file (a restored recovery
 *   copy): saving always writes, even when the page reports the document unchanged
 * @returns {ReturnType<typeof publicDoc>}
 */
function registerDocument(input = {}) {
  const webContents = contentsOf(input.webContents || input.sender || input.window)
  if (!webContents) throw new TypeError('registerDocument needs the window or webContents that shows the document.')
  const docId = input.docId ? checkDocId(input.docId) : crypto.randomUUID()
  const existing = documents.get(docId)
  if (existing && existing.webContentsId !== webContents.id && !isDestroyed(existing.webContents)) {
    throw new Error('This document is shown in another window.')
  }
  const filePath = input.path ? path.resolve(String(input.path)) : null
  const doc = existing || { docId, lossyConfirmed: new Set(), lossy: null, createdAt: Date.now() }
  doc.webContentsId = webContents.id
  doc.webContents = webContents
  doc.path = filePath
  doc.key = filePath ? docKey(filePath) : null
  doc.name = input.name ? String(input.name) : (filePath ? path.basename(filePath) : (existing && existing.name) || null)
  doc.format = input.format || (filePath ? formatOfPath(filePath) : null) || (existing && existing.format) || null
  doc.sourceFormat = doc.format
  doc.stamp = stores.cleanStamp(input.stamp)
  doc.suggestedPath = input.suggestedPath ? path.resolve(String(input.suggestedPath)) : (existing && existing.suggestedPath) || null
  doc.mustWrite = input.mustWrite === true
  doc.conflictStamp = null
  if (input.bytes) keepBytes(doc, input.bytes, doc.stamp && doc.stamp.sha256)
  else if (!existing) keepBytes(doc, null)
  documents.set(docId, doc)
  watchContents(webContents)
  ensureStamp(doc)
  return publicDoc(doc)
}

/**
 * Stamps a document's file in main when it is bound to a path without a
 * stamp (an integrator registered a file it read itself). Saving waits for
 * this; a file that could not be stamped counts as changed by another program.
 */
function ensureStamp(doc) {
  if (!doc.path || doc.stamp) {
    doc.stampPending = null
    return
  }
  const filePath = doc.path
  const pending = core.stampFile(filePath).then((stamp) => {
    if (doc.stampPending === pending && doc.path === filePath && !doc.stamp) doc.stamp = stores.cleanStamp(stamp)
  }, () => {}).then(() => {
    if (doc.stampPending === pending) doc.stampPending = null
  })
  doc.stampPending = pending
}

/**
 * Reads a file for a window, stamps it and registers it. If another window
 * already shows the file, that window is focused instead and nothing is read
 * (one window per file). Opening a window's own file again re-reads it.
 *
 * @param {string} filePath
 * @param {object} options
 * @param {object} options.webContents the window's webContents (or `window`, or `sender`)
 * @param {string} [options.format] registry format id (default: from the extension)
 * @param {string} [options.name]
 * @param {number} [options.maxBytes] larger files fail with TOO_LARGE
 * @param {boolean} [options.focusExisting=true]
 * @param {boolean} [options.replace=true] the file replaces what the window showed, so the window's
 *   other documents are released; pass false for a window that keeps several documents open
 * @returns {Promise<{ok: true, focused: true, docId: string, path: string, name: string}
 *   | {ok: true, focused: false, docId: string, path: string, name: string, format: string|null,
 *      stamp: {size: number, mtimeMs: number, sha256: string}, bytes: Buffer}
 *   | {ok: false, code: string, message: string, context: 'open'}>}
 */
async function openDocument(filePath, options = {}) {
  const webContents = contentsOf(options.webContents || options.sender || options.window)
  if (!webContents) throw new TypeError('openDocument needs the window or webContents that will show the document.')
  const resolved = path.resolve(String(filePath))
  const existing = findDoc(resolved)
  if (existing && options.focusExisting !== false && existing.webContentsId !== webContents.id && !isDestroyed(existing.webContents)) {
    focusDocument(existing.docId)
    return { ok: true, focused: true, docId: existing.docId, path: existing.path, name: existing.name }
  }
  let read
  try {
    read = await core.readFileStamped(resolved, { maxBytes: options.maxBytes })
  } catch (error) {
    return core.toIoResult(error, { path: resolved, stage: 'read', context: 'open' })
  }
  const reuse = existing && existing.webContentsId === webContents.id ? existing.docId : options.docId
  const doc = registerDocument({ webContents, docId: reuse, path: resolved, stamp: read.stamp, bytes: read.bytes, format: options.format, name: options.name })
  if (options.replace !== false) {
    // A file opened in this window replaces what it showed; one window per file
    // must not keep pointing at the earlier file.
    for (const other of docsForContents(webContents.id)) if (other.docId !== doc.docId) documents.delete(other.docId)
  }
  return { ok: true, focused: false, docId: doc.docId, path: resolved, name: doc.name, format: doc.format, stamp: read.stamp, bytes: read.bytes }
}

function findDoc(filePath) {
  if (!filePath) return null
  const key = docKey(filePath)
  for (const doc of documents.values()) if (doc.key === key && !isDestroyed(doc.webContents)) return doc
  return null
}

/** A document of another window that shows `filePath` (one window per file). */
function findHolder(filePath, webContentsId) {
  if (!filePath) return null
  const key = docKey(filePath)
  for (const doc of documents.values()) {
    if (doc.key === key && doc.webContentsId !== webContentsId && !isDestroyed(doc.webContents)) return doc
  }
  return null
}

/**
 * The open document that shows a file, if any window of this process has it.
 * @param {string} filePath
 * @returns {ReturnType<typeof publicDoc>|null}
 */
function findDocumentByPath(filePath) {
  return publicDoc(findDoc(filePath))
}

/**
 * @param {string} docId
 * @returns {ReturnType<typeof publicDoc>|null}
 */
function getDocument(docId) {
  return publicDoc(documents.get(String(docId)))
}

/**
 * Documents shown by a window (or webContents), oldest first.
 * @param {object} target BrowserWindow or webContents
 * @returns {Array<ReturnType<typeof publicDoc>>}
 */
function documentsForWindow(target) {
  const webContents = contentsOf(target)
  if (!webContents) return []
  return [...documents.values()].filter((doc) => doc.webContentsId === webContents.id).map(publicDoc)
}

function docsForContents(webContentsId) {
  return [...documents.values()].filter((doc) => doc.webContentsId === webContentsId)
}

/**
 * Brings the window that shows a document to the front.
 * @param {string} docId
 * @returns {boolean} false when the window is gone
 */
function focusDocument(docId) {
  const doc = documents.get(String(docId))
  const win = doc ? windowOf(doc.webContents) : null
  if (!win) return false
  try {
    if (typeof win.isMinimized === 'function' && win.isMinimized()) win.restore()
    if (typeof win.show === 'function') win.show()
    win.focus()
  } catch {}
  return true
}

/**
 * Removes a document from the registry (its window replaced or closed it).
 * @param {string} docId
 * @returns {boolean}
 */
function forgetDocument(docId) {
  return documents.delete(String(docId))
}

/**
 * Points a document at a file after the workspace wrote it outside
 * performSave (rare), or refreshes its stamp. The stamp must come from main;
 * a path bound without one is stamped here.
 * @param {string} docId
 * @param {{path?: string|null, stamp?: object|null, format?: string, name?: string}} patch
 * @returns {ReturnType<typeof publicDoc>|null}
 */
function bindDocument(docId, patch = {}) {
  const doc = documents.get(String(docId))
  if (!doc) return null
  if (Object.prototype.hasOwnProperty.call(patch, 'path')) {
    doc.path = patch.path ? path.resolve(String(patch.path)) : null
    doc.key = doc.path ? docKey(doc.path) : null
    if (!patch.name) doc.name = doc.path ? path.basename(doc.path) : doc.name
  }
  if (Object.prototype.hasOwnProperty.call(patch, 'stamp')) {
    doc.stamp = stores.cleanStamp(patch.stamp)
    keepBytes(doc, null)
  } else if (Object.prototype.hasOwnProperty.call(patch, 'path')) {
    // A new path never keeps the previous file's stamp.
    doc.stamp = null
    keepBytes(doc, null)
  }
  if (patch.format) doc.format = String(patch.format)
  if (patch.name) doc.name = String(patch.name)
  doc.conflictStamp = null
  // A path without a stamp is stamped in main (see registerDocument).
  ensureStamp(doc)
  return publicDoc(doc)
}

/**
 * The document a request from a window refers to, creating an untitled one
 * for an id the window made up itself. Used by the shared IPC handlers.
 * @param {object} webContents the calling webContents
 * @param {string} docId
 * @param {{name?: string, format?: string}} [init] for a new untitled document
 * @returns {ReturnType<typeof publicDoc>}
 * @throws {Error} when the document belongs to another window
 */
function ensureDocument(webContents, docId, init = {}) {
  const contents = contentsOf(webContents)
  if (!contents) throw new TypeError('ensureDocument needs the calling webContents.')
  const id = checkDocId(docId)
  const existing = documents.get(id)
  if (existing) {
    if (existing.webContentsId !== contents.id && !isDestroyed(existing.webContents)) throw new Error('This document is shown in another window.')
    if (existing.webContentsId !== contents.id) {
      existing.webContentsId = contents.id
      existing.webContents = contents
      watchContents(contents)
    }
    return publicDoc(existing)
  }
  return registerDocument({ webContents: contents, docId: id, path: null, name: init.name, format: init.format })
}

/**
 * Binds a restored recovery entry to the window that restored it: the
 * document keeps its original path and the stamp captured when the snapshot
 * was taken, so saving asks before replacing a file that changed since.
 * - The restored content is not in the file, so saving always writes it
 *   (the pristine shortcut is refused until a save really wrote the file).
 * - An entry without a snapshot-time stamp can't prove the file is unchanged:
 *   it becomes untitled, with Save As starting at its old path.
 * - The window's other documents for the same file are released (the restored
 *   copy replaces what the window showed), as when a file is opened.
 * @param {object} webContents
 * @param {{docId: string, path?: string|null, stamp?: object|null, format?: string, name?: string, suggestedPath?: string|null}} entry
 * @returns {ReturnType<typeof publicDoc>}
 * @throws {Error} when another window shows the entry's document or file
 */
function adoptDocument(webContents, entry) {
  const contents = contentsOf(webContents)
  if (!contents) throw new TypeError('adoptDocument needs the calling webContents.')
  const sourcePath = entry.path ? path.resolve(String(entry.path)) : null
  const stamp = sourcePath ? stores.cleanStamp(entry.stamp) : null
  if (sourcePath) {
    const holder = findHolder(sourcePath, contents.id)
    if (holder && holder.docId !== entry.docId) throw new Error('This file is shown in another window.')
    const key = docKey(sourcePath)
    for (const other of docsForContents(contents.id)) if (other.docId !== entry.docId && other.key === key) documents.delete(other.docId)
  }
  const bound = Boolean(sourcePath && stamp)
  return registerDocument({
    webContents: contents,
    docId: entry.docId,
    path: bound ? sourcePath : null,
    stamp: bound ? stamp : null,
    format: entry.format || undefined,
    name: entry.name || undefined,
    suggestedPath: bound ? (entry.suggestedPath || null) : (sourcePath || entry.suggestedPath || null),
    mustWrite: true,
  })
}

/**
 * Whether `adoptDocument` would succeed for this window: neither the entry's
 * document nor its file is shown in another window.
 * @param {object} webContents
 * @param {{docId: string, path?: string|null}} entry
 * @returns {boolean}
 */
function canAdoptDocument(webContents, entry) {
  const contents = contentsOf(webContents)
  if (!contents) return false
  const existing = documents.get(String(entry.docId || ''))
  if (existing && existing.webContentsId !== contents.id && !isDestroyed(existing.webContents)) return false
  const holder = entry.path ? findHolder(path.resolve(String(entry.path)), contents.id) : null
  return !holder || holder.docId === entry.docId
}

/**
 * Whether performSave is running for a window.
 * @param {object} target BrowserWindow or webContents
 * @returns {boolean}
 */
function isSaving(target) {
  const webContents = contentsOf(target)
  const set = webContents ? activeSaves.get(webContents.id) : null
  return Boolean(set && set.size)
}

function trackSave(webContentsId, promise) {
  if (!activeSaves.has(webContentsId)) activeSaves.set(webContentsId, new Set())
  const set = activeSaves.get(webContentsId)
  set.add(promise)
  const release = () => {
    set.delete(promise)
    if (!set.size && activeSaves.get(webContentsId) === set) activeSaves.delete(webContentsId)
  }
  promise.then(release, release)
  return promise
}

// ---------------------------------------------------------------------------
// performSave
// ---------------------------------------------------------------------------

function normalizeSaveRequest(req) {
  const raw = req && typeof req === 'object' ? req : {}
  const mode = SAVE_MODES.includes(raw.mode) ? raw.mode : 'save'
  const revision = Number.isSafeInteger(raw.revision) ? raw.revision : null
  let bytes = raw.bytes !== undefined ? raw.bytes : raw.data
  if (bytes !== undefined && bytes !== null && !(Buffer.isBuffer(bytes) || ArrayBuffer.isView(bytes) || bytes instanceof ArrayBuffer || typeof bytes === 'string')) bytes = null
  const reason = typeof raw.reason === 'string' ? raw.reason : raw.fallbackReason
  return {
    raw,
    mode,
    revision,
    docId: typeof raw.docId === 'string' && raw.docId ? raw.docId : null,
    format: typeof raw.format === 'string' && formats.formatById(raw.format) ? raw.format : null,
    path: typeof raw.path === 'string' && raw.path ? raw.path : null,
    pristine: raw.pristine === true,
    force: raw.force === true,
    recreate: raw.recreate === true,
    reason: typeof reason === 'string' && core.IO_CODES[reason] ? reason : null,
    confirmedLossy: raw.confirmedLossy === true,
    confirmedSibling: raw.confirmedSibling === true,
    name: typeof raw.name === 'string' ? raw.name.slice(0, 260) : null,
    bytes: bytes === undefined ? null : bytes,
  }
}

function documentForRequest(sender, request) {
  if (request.docId) {
    const existing = documents.get(request.docId)
    if (existing) {
      if (existing.webContentsId !== sender.id) throw new Error('This document is shown in another window.')
      return existing
    }
    registerDocument({ webContents: sender, docId: request.docId, path: null, name: request.name, format: request.format })
    return documents.get(request.docId)
  }
  const own = docsForContents(sender.id)
  if (own.length === 1) return own[0]
  const created = registerDocument({ webContents: sender, path: null, name: request.name, format: request.format })
  return documents.get(created.docId)
}

function saveFailure(error, info) {
  const result = core.toIoResult(error, { path: info.path, stage: info.stage || 'replace' })
  return {
    ...result,
    docId: info.docId,
    mode: info.mode,
    revision: info.revision,
    ...(info.format ? { format: info.format, formatLabel: result.formatLabel || formatLabel(info.format) } : {}),
  }
}

/**
 * The target is the file of a document another Simple window is editing (one
 * window per file). Reported as LOCKED, whose prompt fits: close it there and
 * try again, or save with the same name in another folder.
 */
function alreadyOpenResult(target, info) {
  const name = path.basename(target)
  return {
    ok: false,
    code: 'LOCKED',
    reason: 'open-in-another-window',
    message: core.formatTemplate(core.catalog.notices.OPEN_IN_ANOTHER_WINDOW, { name }),
    technical: 'the file is open in another Simple window',
    name,
    path: target,
    folder: path.dirname(target),
    docId: info.docId,
    mode: info.mode,
    revision: info.revision,
    ...(info.format ? { format: info.format, formatLabel: formatLabel(info.format) } : {}),
  }
}

function asBuffer(data) {
  if (Buffer.isBuffer(data)) return data
  if (ArrayBuffer.isView(data)) return Buffer.from(data.buffer, data.byteOffset, data.byteLength)
  if (data instanceof ArrayBuffer) return Buffer.from(data)
  if (typeof data === 'string') return Buffer.from(data, 'utf8')
  return null
}

function isLossyFormat(formatId) {
  const entry = formatId ? formats.formatById(formatId) : null
  return Boolean(entry && entry.lossy)
}

function canceledResult(info) {
  return saveFailure(new IoError('CANCELED', { path: info.path || undefined, fileName: info.name || undefined }), info)
}

function saveTargetFor(doc, module, engine, hooks) {
  if (typeof hooks.saveTarget === 'function') {
    const custom = hooks.saveTarget(publicDoc(doc))
    if (custom === 'in-place') return { target: 'in-place' }
    if (custom && typeof custom === 'object' && custom.sibling && formats.formatById(custom.sibling)) {
      return { target: 'sibling', formatId: custom.sibling, edited: Boolean(custom.edited) }
    }
    if (custom === 'save-as') return { target: 'save-as' }
  }
  if (!doc.format) return { target: 'save-as' }
  const policy = formats.savePolicy(doc.format, module, { engine })
  if (!policy) return { target: 'save-as' }
  if (policy.target === 'in-place') return { target: 'in-place' }
  return { target: 'sibling', formatId: policy.formatId, edited: false }
}

async function lossInfoFor(doc, format, request, hooks, mode) {
  if (mode === 'export') return null
  const entry = formats.formatById(format)
  let info = null
  if (typeof hooks.lossy === 'function') {
    const custom = await hooks.lossy(format, request.raw, publicDoc(doc))
    if (custom === null || custom === false) return null
    if (custom && typeof custom === 'object') {
      info = {
        keeps: String(custom.keeps || (entry && entry.lossy && entry.lossy.keeps) || ''),
        lost: (Array.isArray(custom.lost) ? custom.lost : (entry && entry.lossy && entry.lossy.loses) || []).map(String),
      }
    }
  }
  if (!info && entry && entry.lossy && format !== doc.sourceFormat) {
    info = { keeps: entry.lossy.keeps || '', lost: (entry.lossy.loses || []).map(String) }
  }
  return info && info.lost.length ? info : null
}

function fullFormatFor(writable, hooks, doc) {
  if (hooks.fullFormat && formats.formatById(hooks.fullFormat)) return hooks.fullFormat
  const full = writable.find((id) => {
    const entry = formats.formatById(id)
    return entry && !entry.lossy
  })
  return full || doc.format
}

function joinList(items) {
  const list = items.filter(Boolean)
  if (list.length <= 1) return list.join('')
  return `${list.slice(0, -1).join(', ')} and ${list[list.length - 1]}`
}

function shortLoss(lost) {
  const list = lost.slice(0, 2)
  const text = joinList(list)
  return text ? text.charAt(0).toUpperCase() + text.slice(1) : ''
}

async function pristineBytes(doc) {
  if (doc.mustWrite || !doc.path || !doc.stamp || !doc.stamp.sha256) return null
  if (doc.originalBytes && doc.originalSha256 === doc.stamp.sha256) return doc.originalBytes
  try {
    const read = await core.readFileStamped(doc.path)
    return read.stamp.sha256 === doc.stamp.sha256 ? read.bytes : null
  } catch {
    return null
  }
}

function normalizePayload(output, format) {
  if (output === null || output === undefined) {
    throw new IoError('SERIALIZE_FAILED', { formatLabel: formatLabel(format) || undefined, technical: 'the serializer returned nothing' })
  }
  if (typeof output === 'function' || Buffer.isBuffer(output) || ArrayBuffer.isView(output) || output instanceof ArrayBuffer || typeof output === 'string') {
    return { data: output, warnings: [] }
  }
  if (typeof output === 'object') {
    const data = output.bytes ?? output.data ?? output.producer
    if (data === undefined || data === null) {
      throw new IoError('SERIALIZE_FAILED', { formatLabel: formatLabel(format) || undefined, technical: 'the serializer returned no bytes' })
    }
    return {
      data,
      warnings: Array.isArray(output.warnings) ? output.warnings.map(String) : [],
      lossy: output.lossy && typeof output.lossy === 'object' ? output.lossy : null,
      expectedSize: Number.isFinite(output.expectedSize) ? output.expectedSize : undefined,
      validatorOptions: output.validatorOptions && typeof output.validatorOptions === 'object' ? output.validatorOptions : undefined,
    }
  }
  throw new IoError('SERIALIZE_FAILED', { formatLabel: formatLabel(format) || undefined, technical: `the serializer returned ${typeof output}` })
}

async function serializeFor(hooks, format, request, doc, target, mode) {
  try {
    if (typeof hooks.serialize === 'function') {
      return normalizePayload(await hooks.serialize(format, request.raw, publicDoc(doc), { target, mode }), format)
    }
    if (request.bytes !== null && request.bytes !== undefined) return normalizePayload(request.bytes, format)
    throw new IoError('SERIALIZE_FAILED', { formatLabel: formatLabel(format) || undefined, technical: 'no serializer and no bytes were given' })
  } catch (error) {
    if (core.isIoError(error)) throw error
    if (error && typeof error.code === 'string' && core.IO_CODES[error.code] && error.code !== 'UNKNOWN') throw error
    throw new IoError('SERIALIZE_FAILED', {
      path: target,
      formatLabel: formatLabel(format) || undefined,
      reason: error && typeof error.userMessage === 'string' ? error.userMessage : undefined,
      technical: core.technicalDetail(error),
      cause: error,
    })
  }
}

async function safeHook(hook, ...args) {
  if (typeof hook !== 'function') return
  try {
    await hook(...args)
  } catch (error) {
    // Bookkeeping after a completed save (recents, preferences) never turns it into a failure.
    logger.warn(`A step after saving failed: ${error && error.message}`)
  }
}

/**
 * Saves, saves as, saves a copy or exports a document, and always resolves to
 * an IoResult. Workspace save and export handlers are thin wrappers:
 *
 *   ipcMain.handle('workbook:save', (event, req) => performSave(event, req, {
 *     serialize: (format, req, doc) => serializeWorkbook(req.workbook, format),
 *     validate: (format, bytes) => format === 'xls' ? verifyXlsReopen(bytes) : null,
 *   }))
 *
 * Steps: resolve the target (bound file, sibling file after one confirmation,
 * or a Save dialog; the lossy prompt once per format and session) → serialize,
 * or reuse the opened bytes for an untouched document → copy the original to
 * the versions store before its first overwrite this session → safeWriteFile
 * with the stamp main took at open → rebind the document (Save, Save As) or
 * refresh its stamp (export over its own file) → bookkeeping hooks.
 *
 * @param {{sender: object}} event the IPC event of the calling window
 * @param {object} req from the renderer
 * @param {string} [req.docId] the document (from the open handler); a new id makes an untitled document
 * @param {'save'|'save-as'|'save-copy'|'export'} [req.mode='save']
 * @param {number} [req.revision] echoed in the result, so the renderer clears dirty for that revision only
 * @param {boolean} [req.pristine] the document is unchanged since it was opened or last saved (ignored for a
 *   restored recovery copy and after an export replaced the document's own file, until a save writes it)
 * @param {string} [req.format] format id to write (export, Save As, or the alternative after NEEDS_OFFICE_ENGINE)
 * @param {string} [req.path] a path the user chose with simpleIO.chooseSavePath in this window
 * @param {boolean} [req.force] the user chose Replace after CHANGED_ON_DISK; honoured only when main showed this
 *   window that prompt for the target (simpleIO.prompt) and the user chose Replace. For the document's own file
 *   it replaces only the version the user was shown: a newer change is reported again
 * @param {boolean} [req.recreate] the user chose Save Here Again after SOURCE_MISSING (same rule as force)
 * @param {boolean} [req.confirmedLossy] the page showed prompts.lossy-save through simpleIO.prompt and the user
 *   chose save-lossy; otherwise main asks itself
 * @param {boolean} [req.confirmedSibling] the same for prompts.sibling-save and save-sibling
 * @param {string} [req.reason] with mode 'save-as': the failure code that led here (opens the fallback folder)
 * @param {Uint8Array} [req.bytes] the serialized document, for workspaces that serialize in the renderer
 * @param {object} [hooks]
 * @param {(format: string, req: object, doc: object, info: {target: string, mode: string}) => any} [hooks.serialize]
 *   bytes, a producer, or {bytes|producer, warnings?, lossy?, expectedSize?, validatorOptions?}; may be async
 * @param {(format: string, bytes: Buffer|null, info: object) => any} [hooks.validate] deep check on the written bytes
 * @param {string[]} [hooks.formats] Save As formats, current first (default: registry in-place formats)
 * @param {string} [hooks.fullFormat] the full-fidelity format offered by the lossy prompt (default: first non-lossy)
 * @param {(format: string, req: object, doc: object) => any} [hooks.lossy] null/false: nothing is lost;
 *   {keeps, lost}: ask first; undefined: registry default (asked when converting into a lossy format)
 * @param {(doc: object) => 'in-place'|'save-as'|{sibling: string, edited?: boolean}|undefined} [hooks.saveTarget]
 *   override where Save writes (signed PDFs: {sibling: 'pdf', edited: true})
 * @param {(result: object, doc: object) => any} [hooks.afterSave] recents and other bookkeeping; errors are logged only
 * @param {boolean} [hooks.engine] the optional office engine is available
 * @param {string} [hooks.module] workspace (default: configureDocumentGuard)
 * @param {object} [hooks.validatorOptions] options for the structural validator (text encoding, …)
 * @param {(progress: object) => void} [hooks.onProgress]
 * @returns {Promise<IoResult>}
 */
function performSave(event, req, hooks = {}) {
  const sender = event && event.sender
  if (!sender || typeof sender.id !== 'number') throw new TypeError('performSave needs the IPC event of the calling window.')
  const request = normalizeSaveRequest(req)
  let doc
  try {
    doc = documentForRequest(sender, request)
  } catch (error) {
    return Promise.resolve(saveFailure(new IoError('UNKNOWN', { technical: error.message, cause: error }), { mode: request.mode, revision: request.revision }))
  }
  const run = saveDocument(sender, doc, request, hooks || {})
  trackSave(sender.id, run)
  return run
}

async function saveDocument(sender, doc, request, hooks) {
  const module = hooks.module || settings.module
  const engine = Boolean(hooks.engine)
  const window = windowOf(sender)
  const mode = request.mode
  const info = { docId: doc.docId, mode, revision: request.revision, path: doc.path, name: doc.name, format: null, stage: 'preflight' }
  // Flags that skip a safety check count only when main showed this window the
  // matching prompt and the user chose that answer (never on the page's word alone).
  const decided = (key, answer, filePath) => dialogs.hasDecision(sender, key, answer, filePath)
  if (doc.stampPending) await doc.stampPending
  const writableFor = (purpose) => (purpose !== 'export' && Array.isArray(hooks.formats) && hooks.formats.length
    ? hooks.formats.filter((id) => formats.formatById(id))
    : dialogs.writableFormats(module, purpose, { engine }))
  const untitledName = () => doc.name || core.formatTemplate(core.catalog.recovery.untitled, { kind: core.catalog.kinds[module] || 'document' })

  async function chooseTarget(purpose, formatId, rebind) {
    const choice = await dialogs.chooseSavePath(
      { purpose, name: untitledName(), format: formatId || undefined, formats: writableFor(mode === 'export' ? 'export' : purpose), reason: request.reason || undefined },
      { window, sender, module, sourcePath: doc.path || doc.suggestedPath, engine },
    )
    if (!choice) return { canceled: true }
    return { target: choice.path, format: choice.format, rebind, fromDialog: true }
  }

  async function resolveTarget() {
    const rebindsHere = mode === 'save' || mode === 'save-as'
    if (request.path) {
      const chosen = path.resolve(request.path)
      const own = Boolean(doc.path) && dialogs.samePath(chosen, doc.path)
      if (!own && !dialogs.isGranted(sender, chosen)) {
        throw new IoError('UNKNOWN', { path: chosen, technical: 'the target was not chosen in a Save dialog of this window' })
      }
      return { target: chosen, format: request.format || formatOfPath(chosen) || doc.format, rebind: rebindsHere, fromDialog: !own }
    }
    if (mode === 'save' && doc.path) {
      if (request.format && request.format !== doc.format) {
        // The user picked the offered alternative ("Save as .xlsx") after NEEDS_OFFICE_ENGINE:
        // the edits go next to the original, which stays unchanged.
        const sibling = await dialogs.siblingPathFor(doc.path, request.format, {})
        return { target: sibling, format: request.format, rebind: true, sibling: { formatLabel: formatLabel(doc.format) } }
      }
      const policy = saveTargetFor(doc, module, engine, hooks)
      if (policy.target === 'in-place') return { target: doc.path, format: doc.format, rebind: false }
      if (policy.target === 'sibling') {
        const sibling = await dialogs.siblingPathFor(doc.path, policy.formatId, { edited: policy.edited })
        const answer = request.confirmedSibling && decided('prompts.sibling-save', 'save-sibling')
          ? 'save-sibling'
          : await dialogs.showPrompt('prompts.sibling-save', {
            name: doc.name || path.basename(doc.path),
            siblingName: path.basename(sibling),
            formatLabel: formatLabel(doc.format) || path.extname(doc.path),
          }, { window, module })
        if (answer === 'save-sibling') {
          return { target: sibling, format: policy.formatId, rebind: true, sibling: { formatLabel: policy.edited ? null : formatLabel(doc.format) } }
        }
        if (answer === 'choose-location') return chooseTarget('save-as', policy.formatId, true)
        return { canceled: true }
      }
      return chooseTarget('save-as', doc.format, true)
    }
    // With a failure code this is the fallback Save As: for a file that can't be written where
    // it is, the same name in a folder that works; for a conflict, "<name> (edited)" beside it.
    const purpose = request.reason ? 'fallback' : (mode === 'save' ? 'save-as' : mode)
    return chooseTarget(purpose, request.format || doc.format, rebindsHere)
  }

  try {
    let resolved = await resolveTarget()
    if (resolved.canceled) return canceledResult(info)
    let { target, format } = resolved
    info.path = target
    info.format = format
    if (!format || !formats.formatById(format)) {
      throw new IoError('INVALID_NAME', { path: target, reason: 'unknown-format', technical: `no format is known for "${path.extname(target) || path.basename(target)}"` })
    }

    // One window per file: never write over a file another window is editing.
    const holder = findHolder(target, sender.id)
    if (holder && holder.docId !== doc.docId) return alreadyOpenResult(target, info)

    // Lossy formats: ask once per format and session before the first save into one.
    const fullFormat = fullFormatFor(writableFor('save-as'), hooks, doc)
    let loss = await lossInfoFor(doc, format, request, hooks, mode)
    const lossConfirmed = request.confirmedLossy && decided('prompts.lossy-save', 'save-lossy')
    if (loss && !lossConfirmed && !doc.lossyConfirmed.has(format)) {
      const answer = await dialogs.showPrompt('prompts.lossy-save', {
        name: doc.name || path.basename(target),
        formatLabel: formatLabel(format),
        fullLabel: formatLabel(fullFormat) || 'the full format',
        keeps: loss.keeps || 'part of the content',
        lostList: joinList(loss.lost),
      }, { window, module })
      if (answer === 'cancel') return canceledResult(info)
      if (answer === 'save-full' && fullFormat && fullFormat !== format) {
        resolved = await chooseTarget(mode === 'save-copy' ? 'save-copy' : 'save-as', fullFormat, mode !== 'save-copy')
        if (resolved.canceled) return canceledResult(info)
        target = resolved.target
        format = resolved.format
        info.path = target
        info.format = format
        const again = findHolder(target, sender.id)
        if (again && again.docId !== doc.docId) return alreadyOpenResult(target, info)
        loss = await lossInfoFor(doc, format, request, hooks, mode)
      } else {
        doc.lossyConfirmed.add(format)
      }
    }

    const own = Boolean(doc.path) && dialogs.samePath(target, doc.path)
    const sameFormat = format === doc.format
    // Replace after CHANGED_ON_DISK and Save Here Again after SOURCE_MISSING.
    const forced = request.force && decided('saveFailed.CHANGED_ON_DISK', 'replace', target)
    const recreating = request.recreate && decided('saveFailed.SOURCE_MISSING', 'recreate', target)
    // A restored recovery copy, or a file an export replaced, does not hold the document: always write.
    const pristine = request.pristine && !doc.mustWrite

    // An untouched document saved to its own file in its own format needs no write.
    if (pristine && own && sameFormat && mode === 'save' && doc.stamp && !forced && !recreating) {
      const comparison = await core.compareStamp(target, doc.stamp)
      if (comparison.state === 'same' || comparison.state === 'restamped') {
        if (comparison.state === 'restamped') doc.stamp = { ...comparison.stamp }
        return {
          ok: true, docId: doc.docId, mode, revision: request.revision, path: target, name: path.basename(target), format,
          formatLabel: formatLabel(format), stamp: { ...doc.stamp }, strategy: 'unchanged', attempts: 0, ms: 0,
          warnings: [], notes: [], lossy: doc.lossy ? { ...doc.lossy } : null, folderChanged: false, rebound: false, sibling: false, versionBackup: null,
        }
      }
      if (comparison.state === 'changed') throw new IoError('CHANGED_ON_DISK', { path: target, technical: 'the file changed after it was opened' })
      throw new IoError('SOURCE_MISSING', { path: target, technical: 'the file is no longer at this path' })
    }

    // Bytes: the opened original for an untouched document in its own format, else serialize.
    info.stage = 'serialize'
    let payload = null
    if (pristine && sameFormat && doc.path) {
      const original = await pristineBytes(doc)
      if (original) payload = { data: original, warnings: [], pristine: true }
    }
    if (!payload) payload = await serializeFor(hooks, format, request, doc, target, mode)

    // A stale file is never backed up or replaced without asking.
    info.stage = 'preflight'
    let expectedStamp = own && doc.stamp ? doc.stamp : undefined
    let force = forced
    if (forced && own && doc.conflictStamp) {
      // Replace means "replace the version I was shown": if another program
      // writes the file again meanwhile, the user is asked again.
      expectedStamp = doc.conflictStamp
      force = false
    }
    const recreate = recreating || (own && Boolean(resolved.fromDialog))
    let versionBackup = null
    if (await core.pathExists(target).catch(() => false)) {
      if (own && !expectedStamp && !force) {
        // A document bound to its file without a stamp can't prove the file is unchanged.
        throw new IoError('CHANGED_ON_DISK', { path: target, reason: 'no-stamp', technical: 'the file was never stamped when it was opened, so a change by another program cannot be ruled out' })
      }
      if (expectedStamp && !force) {
        const comparison = await core.compareStamp(target, expectedStamp)
        if (comparison.state === 'changed') {
          throw new IoError('CHANGED_ON_DISK', { path: target, technical: 'the content differs from the version that was opened' })
        }
      }
      versionBackup = await stores.versionsStore().backupOnce(target).catch((error) => ({ ok: false, reason: error && error.message }))
    }

    info.stage = 'replace'
    const validate = typeof hooks.validate === 'function'
      ? (bytes, details) => hooks.validate(format, bytes, { ...details, doc: publicDoc(doc), request: request.raw })
      : undefined
    const written = await safeWriteFile(target, payload.data, {
      expectedStamp,
      force,
      recreate,
      format: payload.pristine ? undefined : format,
      validatorOptions: payload.validatorOptions || hooks.validatorOptions,
      validate: payload.pristine ? undefined : validate,
      expectedSize: payload.expectedSize,
      onProgress: hooks.onProgress,
    })

    // Bookkeeping. Nothing below may turn the completed save into a failure.
    const previousFolder = doc.path ? path.dirname(doc.path) : null
    const originalPath = doc.path
    const rebound = Boolean(resolved.rebind) && !own
    const updatesDocument = rebound || (own && mode !== 'export' && mode !== 'save-copy')
    const outputLoss = loss || (payload.lossy && Array.isArray(payload.lossy.lost) && payload.lossy.lost.length
      ? { keeps: '', lost: payload.lossy.lost.map(String) }
      : null)
    const describeLoss = (lost) => ({
      format,
      formatLabel: formatLabel(format),
      lost,
      lostShort: shortLoss(lost),
      fullFormat: fullFormat && fullFormat !== format ? fullFormat : null,
      fullLabel: fullFormat && fullFormat !== format ? formatLabel(fullFormat) : null,
    })
    if (rebound) {
      doc.path = written.path
      doc.key = docKey(written.path)
      doc.name = path.basename(written.path)
      doc.format = format
      doc.sourceFormat = format
    }
    let ownFileChanged = false
    if (updatesDocument) {
      doc.stamp = { ...written.stamp }
      doc.stampPending = null
      doc.conflictStamp = null
      doc.mustWrite = false
      keepBytes(doc, asBuffer(payload.data), written.stamp.sha256)
      // A lossy save marks the document until it is saved in a full format again.
      if (outputLoss) doc.lossy = describeLoss(outputLoss.lost)
      else if (!isLossyFormat(format)) doc.lossy = null
    } else if (own) {
      // An export or copy written over the document's own file: the file changed, the document did not.
      // The file no longer holds the document, so the next Save must write it (never the pristine shortcut).
      doc.stamp = { ...written.stamp }
      doc.conflictStamp = null
      doc.mustWrite = true
      keepBytes(doc, null)
      ownFileChanged = true
    }
    dialogs.clearDecisions(sender, written.path)
    let lossResult = null
    if (updatesDocument) lossResult = doc.lossy ? { ...doc.lossy } : null
    else if (mode === 'save-copy' && outputLoss) lossResult = describeLoss(outputLoss.lost)
    const sibling = Boolean(resolved.sibling) && Boolean(originalPath)
    const result = {
      ok: true,
      docId: doc.docId,
      mode,
      revision: request.revision,
      path: written.path,
      name: path.basename(written.path),
      format,
      formatLabel: formatLabel(format),
      stamp: { ...written.stamp },
      strategy: written.strategy,
      attempts: written.attempts,
      ms: written.ms,
      warnings: payload.warnings || [],
      notes: written.notes || [],
      lossy: lossResult,
      folderChanged: !previousFolder || !dialogs.samePath(previousFolder, path.dirname(written.path)),
      rebound,
      sibling,
      versionBackup: versionBackup ? summarizeBackup(versionBackup) : null,
    }
    if (ownFileChanged) result.ownFileChanged = true
    if (sibling) {
      result.originalPath = originalPath
      result.originalName = path.basename(originalPath)
      if (resolved.sibling.formatLabel) result.originalFormatLabel = resolved.sibling.formatLabel
    }
    if (versionBackup && versionBackup.ok && versionBackup.path) result.backupPath = versionBackup.path
    else if (versionBackup && !versionBackup.ok) result.backupFailed = versionBackup.reason || 'unknown'
    await safeHook(hooks.afterSave, { ...result }, publicDoc(doc))
    return result
  } catch (error) {
    const failure = saveFailure(error, info)
    if (failure.code === 'CHANGED_ON_DISK' && doc.path && info.path && dialogs.samePath(info.path, doc.path)) {
      // The version the user is about to be shown: Replace replaces exactly this one.
      doc.conflictStamp = await core.stampFile(info.path).then(stores.cleanStamp, () => null)
    }
    return failure
  }
}

function summarizeBackup(backup) {
  if (backup.ok && backup.id) return { ok: true, id: backup.id }
  if (backup.ok) return { ok: true, skipped: backup.skipped || null }
  return { ok: false, reason: backup.reason || 'unknown' }
}

// ---------------------------------------------------------------------------
// Renderer RPC (main → page through the preload bridge)
// ---------------------------------------------------------------------------

const pendingRequests = new Map()
let responseListener = null

function settleRequest(id, outcome) {
  const entry = pendingRequests.get(id)
  if (!entry) return
  pendingRequests.delete(id)
  if (entry.timer) clearTimeout(entry.timer)
  entry.resolve(outcome)
}

function failRequestsFor(webContentsId, reason) {
  for (const [id, entry] of [...pendingRequests]) if (entry.webContentsId === webContentsId) settleRequest(id, { ok: false, reason })
}

/**
 * Handles an `io:response` message from a page. Only the page a request was
 * sent to can answer it.
 * @param {{sender: object}} event
 * @param {{id: string, ok: boolean, value?: unknown, unhandled?: boolean, error?: string}} message
 */
function handleRendererResponse(event, message) {
  if (!message || typeof message !== 'object' || typeof message.id !== 'string') return
  const entry = pendingRequests.get(message.id)
  const senderId = event && event.sender && event.sender.id
  if (!entry || senderId !== entry.webContentsId) return
  if (message.ok) settleRequest(message.id, { ok: true, value: message.value === undefined ? null : message.value })
  else settleRequest(message.id, { ok: false, reason: message.unhandled ? 'unhandled' : 'error', error: typeof message.error === 'string' ? message.error.slice(0, 500) : null })
}

function ensureResponseListener(ipcMain) {
  if (responseListener) return true
  const value = ipcMain || (electron() && electron().ipcMain)
  if (!value || typeof value.on !== 'function') return false
  const listener = (event, message) => handleRendererResponse(event, message)
  value.on('io:response', listener)
  responseListener = { ipcMain: value, listener }
  return true
}

/**
 * Sends a request to a window's page and waits for its answer. Pages answer
 * through `simpleIO.onRequest(type, handler)`; a page without a handler
 * answers at once with reason 'unhandled'.
 *
 * @param {object} target BrowserWindow or webContents
 * @param {'close-query'|'save-now'|'discard'|'recovery-flush'} type
 * @param {unknown} [payload]
 * @param {{timeoutMs?: number}} [options] 0 or absent waits until the page answers or goes away
 * @returns {Promise<{ok: true, value: unknown} | {ok: false, reason: 'timeout'|'gone'|'unhandled'|'error', error?: string|null}>}
 */
function requestRenderer(target, type, payload, options = {}) {
  if (!REQUEST_TYPES.includes(type)) throw new TypeError(`Unknown renderer request "${type}".`)
  const webContents = contentsOf(target)
  if (!webContents || isDestroyed(webContents) || (typeof webContents.isCrashed === 'function' && webContents.isCrashed())) {
    return Promise.resolve({ ok: false, reason: 'gone' })
  }
  ensureResponseListener()
  const id = crypto.randomUUID()
  return new Promise((resolve) => {
    const entry = { resolve, webContentsId: webContents.id, timer: null }
    pendingRequests.set(id, entry)
    const timeoutMs = Number(options.timeoutMs)
    if (Number.isFinite(timeoutMs) && timeoutMs > 0) entry.timer = setTimeout(() => settleRequest(id, { ok: false, reason: 'timeout' }), timeoutMs)
    try {
      webContents.send('io:request', { id, type, payload: payload === undefined ? null : payload })
    } catch {
      settleRequest(id, { ok: false, reason: 'gone' })
    }
  })
}

// ---------------------------------------------------------------------------
// Window guard
// ---------------------------------------------------------------------------

const guardedWindows = new Map()
const appGuard = { installed: false, app: null, quitApproved: false, flow: null, api: null, listener: null }

/**
 * Normalizes a page's close-query answer:
 * {dirty, saving, title?, kind?, docId?, untitled?, lossy?: {formatLabel, fullFormat?, fullLabel?, lost?|lostShort?}}.
 * A document is untitled when the page says so, or when its registry entry has no file.
 */
function normalizeStatus(value, state) {
  const status = typeof value === 'boolean' ? { dirty: value } : (value && typeof value === 'object' ? value : {})
  const docs = docsForContents(state.webContentsId)
  const named = typeof status.docId === 'string' ? documents.get(status.docId) : null
  const doc = named && named.webContentsId === state.webContentsId ? named : (docs.length === 1 ? docs[0] : null)
  let untitled = status.untitled === true
  if (status.untitled === undefined && typeof status.docId === 'string') untitled = !named || !named.path
  else if (status.untitled === undefined && doc) untitled = !doc.path
  const lossy = status.lossy && typeof status.lossy === 'object' ? status.lossy : null
  const lost = lossy && Array.isArray(lossy.lost) ? lossy.lost.map(String) : []
  return {
    dirty: status.dirty === true,
    saving: status.saving === true,
    untitled,
    title: typeof status.title === 'string' && status.title ? status.title.slice(0, 260) : (doc ? doc.name : null),
    kind: typeof status.kind === 'string' && status.kind ? status.kind.slice(0, 40) : null,
    lossy: lossy ? {
      formatLabel: String(lossy.formatLabel || ''),
      fullFormat: typeof lossy.fullFormat === 'string' && formats.formatById(lossy.fullFormat) ? lossy.fullFormat : null,
      fullLabel: String(lossy.fullLabel || ''),
      lostShort: String(lossy.lostShort || shortLoss(lost)),
    } : null,
  }
}

function lastAutosaveText(state) {
  const ids = docsForContents(state.webContentsId).map((doc) => doc.docId)
  let at = null
  try { at = stores.recoveryStore().lastWriteFor(ids) } catch {}
  if (!at) return undefined
  try {
    return new Date(at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
  } catch {
    return new Date(at).toISOString().slice(11, 16)
  }
}

function promptVars(state, extra = {}) {
  return { appName: dialogs.appNameFor(state.module), lastAutosave: lastAutosaveText(state), ...extra }
}

function orphanWindowDocuments(state) {
  const ids = docsForContents(state.webContentsId).map((doc) => doc.docId)
  for (const id of ids) documents.delete(id)
  if (!ids.length) return
  try {
    stores.recoveryStore().markOrphaned(ids).catch((error) => logger.warn(`Could not keep the recovery copies of a closed window: ${error && error.message}`))
  } catch (error) {
    logger.warn(`Could not keep the recovery copies of a closed window: ${error && error.message}`)
  }
}

function isGone(state) {
  return isDestroyed(state.win) || isDestroyed(state.webContents)
}

async function waitForWindowSaves(state) {
  for (;;) {
    const set = activeSaves.get(state.webContentsId)
    if (!set || !set.size) return
    let timer
    const pause = new Promise((resolve) => { timer = setTimeout(resolve, SAVE_WAIT_ROUND_MS) })
    await Promise.race([Promise.allSettled([...set]), pause])
    clearTimeout(timer)
  }
}

function reopenAndRecover(state) {
  if (isGone(state)) return
  orphanWindowDocuments(state)
  state.expectingCrash = true
  try { state.webContents.forcefullyCrashRenderer() } catch {}
  try { state.webContents.reload() } catch {}
}

function closeWindowNow(state) {
  if (isGone(state)) return
  orphanWindowDocuments(state)
  state.approved = true
  try { state.win.destroy() } catch {}
}

/**
 * Shows the "not responding" prompt once (concurrent callers share it) and
 * carries out the choice. The prompt closes by itself when the page recovers.
 * @returns {Promise<'wait'|'reopen'|'close-window'>}
 */
function handleNotResponding(state) {
  if (state.notResponding) return state.notResponding.promise
  const controller = new AbortController()
  const promise = dialogs.showPrompt('prompts.not-responding', promptVars(state), { window: state.win, module: state.module, signal: controller.signal })
    .catch((error) => {
      logger.warn(`Could not show the not-responding prompt: ${error && error.message}`)
      return 'wait'
    })
    .then((answer) => {
      state.notResponding = null
      const choice = controller.signal.aborted ? 'wait' : answer
      if (choice === 'reopen') reopenAndRecover(state)
      else if (choice === 'close-window') closeWindowNow(state)
      return choice
    })
  state.notResponding = { controller, promise }
  return promise
}

async function closeFlow(state, options = {}) {
  let lossyHandled = false
  // A page that crashed meanwhile is handled by its own prompt (Reopen and
  // Recover); quitting simply goes on, since its recovery copy is kept.
  const settledByCrash = () => (isGone(state) ? 'approved' : state.crashed ? (options.quitting ? 'approved' : 'canceled') : null)
  for (let round = 0; round < MAX_CLOSE_ROUNDS; round += 1) {
    if (settledByCrash()) return settledByCrash()
    await waitForWindowSaves(state)
    const reply = await requestRenderer(state.webContents, 'close-query', { quitting: Boolean(options.quitting) }, { timeoutMs: state.config.queryTimeoutMs })
    if (settledByCrash()) return settledByCrash()
    if (!reply.ok) {
      // No document session in the page: nothing can be lost.
      if (reply.reason === 'unhandled') return 'approved'
      // The page went away without the window: it crashed or is being reopened, and
      // its own prompt (or the reload) decides; quitting goes on.
      if (reply.reason === 'gone') return isGone(state) || options.quitting ? 'approved' : 'canceled'
      const answer = await handleNotResponding(state)
      if (answer === 'wait') continue
      return isGone(state) ? 'approved' : 'canceled'
    }
    const status = normalizeStatus(reply.value, state)
    if (status.saving) {
      // Never close during a save: wait for it, then ask again.
      await core.sleep(state.config.pollMs)
      continue
    }
    if (status.dirty) {
      const key = status.untitled ? 'prompts.unsaved-untitled' : 'prompts.unsaved'
      const answer = await dialogs.showPrompt(key, promptVars(state, {
        name: status.title || undefined,
        kind: status.kind || core.catalog.kinds[state.module] || 'document',
      }), { window: state.win, module: state.module })
      if (answer === 'dont-save') {
        await requestRenderer(state.webContents, 'discard', { reason: options.quitting ? 'quit' : 'close' }, { timeoutMs: state.config.discardTimeoutMs })
        return 'approved'
      }
      if (answer !== 'save') return 'canceled'
      const saved = await requestRenderer(state.webContents, 'save-now', { reason: options.quitting ? 'quit' : 'close' })
      if (settledByCrash()) return settledByCrash()
      if (saved.ok && saved.value === true) continue
      return 'canceled'
    }
    if (status.lossy && !lossyHandled) {
      const answer = await dialogs.showPrompt('prompts.lossy-close', promptVars(state, {
        name: status.title || undefined,
        formatLabel: status.lossy.formatLabel || undefined,
        fullLabel: status.lossy.fullLabel || undefined,
        lostShort: status.lossy.lostShort ? status.lossy.lostShort.toLowerCase() : undefined,
      }), { window: state.win, module: state.module })
      if (answer === 'close-anyway') return 'approved'
      if (answer !== 'save-full-copy') return 'canceled'
      const saved = await requestRenderer(state.webContents, 'save-now', {
        reason: options.quitting ? 'quit' : 'close',
        mode: 'save-copy',
        format: status.lossy.fullFormat || undefined,
        fullCopy: true,
      })
      if (settledByCrash()) return settledByCrash()
      if (!(saved.ok && saved.value === true)) return 'canceled'
      lossyHandled = true
      continue
    }
    return 'approved'
  }
  return 'canceled'
}

function startCloseFlow(state, options = {}) {
  if (!state.flow) {
    state.flow = closeFlow(state, options)
      .catch((error) => {
        logger.warn(`The close check failed: ${error && error.message}`)
        return 'canceled'
      })
      .then((outcome) => {
        state.flow = null
        if (outcome === 'approved' && !isDestroyed(state.win)) {
          state.approved = true
          try { state.win.close() } catch {}
        }
        return outcome
      })
  }
  return state.flow
}

/**
 * Guards one window against losing edits:
 * - Closing asks the page (close-query) and, when it has unsaved changes, shows
 *   Save / Don't Save / Cancel with Save as the default and Esc as Cancel.
 *   Save asks the page to save (save-now) and closes only when that succeeded.
 *   A window is never closed while a save is running.
 * - A page that does not answer within 5 s gets the "not responding" prompt:
 *   Wait, Reopen and Recover, or Close Window (its recovery copy is kept).
 * - A crashed page leaves a closable window and the "stopped unexpectedly"
 *   prompt; its recovery copies are offered to the reloaded page.
 * - At Windows shutdown or sign-out each page gets 2 s to write its recovery
 *   copy; shutdown is never blocked.
 * The first guarded window also installs the app guard (installAppGuard) on
 * Electron's app, unless `appGuard: false`.
 *
 * @param {object} win BrowserWindow
 * @param {object} [options]
 * @param {string} [options.module] workspace (default: configureDocumentGuard)
 * @param {boolean} [options.appGuard=true] also guard quitting (installAppGuard), once per process
 * @param {number} [options.queryTimeoutMs] default 5000
 * @param {number} [options.discardTimeoutMs] default 5000
 * @param {number} [options.flushTimeoutMs] default 2000
 * @param {number} [options.pollMs] default 250
 * @param {object} [options.ipcMain] for the io:response listener (default: electron's)
 * @returns {{requestClose: () => Promise<'approved'|'canceled'>, isApproved: () => boolean, dispose: () => void}}
 */
function installWindowGuard(win, options = {}) {
  if (!win || typeof win.on !== 'function' || !win.webContents) throw new TypeError('installWindowGuard needs a BrowserWindow.')
  if (guardedWindows.has(win)) return guardedWindows.get(win).api
  const webContents = win.webContents
  const state = {
    win,
    webContents,
    webContentsId: webContents.id,
    module: options.module || settings.module,
    config: {
      queryTimeoutMs: options.queryTimeoutMs ?? settings.queryTimeoutMs,
      discardTimeoutMs: options.discardTimeoutMs ?? settings.discardTimeoutMs,
      flushTimeoutMs: options.flushTimeoutMs ?? settings.flushTimeoutMs,
      pollMs: options.pollMs ?? settings.pollMs,
    },
    approved: false,
    crashed: false,
    expectingCrash: false,
    crashPrompt: false,
    flow: null,
    notResponding: null,
    api: null,
  }
  ensureResponseListener(options.ipcMain)
  watchContents(webContents)
  // Quitting must be as safe as closing, so the first guarded window also guards the app.
  if (!appGuard.installed && options.appGuard !== false) {
    const value = electron()
    if (value && value.app && typeof value.app.on === 'function') installAppGuard(value.app)
  }

  const onClose = (event) => {
    // Only this window's own decision lets it close: a window that opened
    // while quitting was being decided is still asked.
    if (state.approved || state.crashed) return
    event.preventDefault()
    startCloseFlow(state)
  }
  const onUnresponsive = () => {
    if (isGone(state) || state.crashed || state.approved) return
    handleNotResponding(state)
  }
  const onResponsive = () => {
    if (state.notResponding) state.notResponding.controller.abort()
  }
  const onSessionEnd = () => {
    // Never block a shutdown or sign-out; give the page 2 s to keep its edits.
    requestRenderer(webContents, 'recovery-flush', { reason: 'session-end', budgetMs: state.config.flushTimeoutMs }, { timeoutMs: state.config.flushTimeoutMs })
  }
  const onGone = (_event, details) => {
    if (details && details.reason === 'clean-exit') return
    failRequestsFor(state.webContentsId, 'gone')
    if (state.notResponding) state.notResponding.controller.abort()
    orphanWindowDocuments(state)
    if (state.expectingCrash) {
      state.expectingCrash = false
      return
    }
    if (appGuard.quitApproved || isDestroyed(win)) return
    state.crashed = true
    state.approved = true
    if (state.crashPrompt) return
    state.crashPrompt = true
    dialogs.showPrompt('prompts.crashed', promptVars(state), { window: win, module: state.module })
      .catch(() => 'close')
      .then((answer) => {
        state.crashPrompt = false
        if (isDestroyed(win)) return
        if (answer === 'reopen') {
          state.crashed = false
          state.approved = false
          try { webContents.reload() } catch {}
        } else {
          try { win.destroy() } catch {}
        }
      })
  }
  const onPreventUnload = (event) => {
    // A page's beforeunload handler must not undo an approved close.
    if (state.approved) event.preventDefault()
  }

  win.on('close', onClose)
  win.on('unresponsive', onUnresponsive)
  win.on('responsive', onResponsive)
  win.on('query-session-end', onSessionEnd)
  webContents.on('render-process-gone', onGone)
  webContents.on('will-prevent-unload', onPreventUnload)
  const dispose = () => {
    win.removeListener('close', onClose)
    win.removeListener('unresponsive', onUnresponsive)
    win.removeListener('responsive', onResponsive)
    win.removeListener('query-session-end', onSessionEnd)
    if (!isDestroyed(webContents)) {
      webContents.removeListener('render-process-gone', onGone)
      webContents.removeListener('will-prevent-unload', onPreventUnload)
    }
    guardedWindows.delete(win)
  }
  win.once('closed', () => {
    guardedWindows.delete(win)
    failRequestsFor(state.webContentsId, 'gone')
  })
  state.api = Object.freeze({
    requestClose: () => startCloseFlow(state),
    isApproved: () => state.approved,
    dispose,
  })
  guardedWindows.set(win, state)
  return state.api
}

// ---------------------------------------------------------------------------
// App guard
// ---------------------------------------------------------------------------

function nextWindowToAsk() {
  for (const state of guardedWindows.values()) if (!isGone(state) && !state.approved) return state
  return null
}

async function runQuitFlow(application, timeoutMs) {
  await core.waitForPendingWrites({ timeoutMs })
  // Windows that open while a prompt is showing are asked too: the loop ends
  // only when no guarded window is left that has not been decided.
  for (let asked = 0; asked < 10_000; asked += 1) {
    const state = nextWindowToAsk()
    if (!state) break
    try {
      if (typeof state.win.isMinimized === 'function' && state.win.isMinimized()) state.win.restore()
      if (typeof state.win.show === 'function') state.win.show()
      state.win.focus()
    } catch {}
    const outcome = await startCloseFlow(state, { quitting: true })
    if (outcome !== 'approved' && !isGone(state)) return false
    if (!isGone(state) && !state.approved) state.approved = true
  }
  await core.waitForPendingWrites({ timeoutMs })
  if (nextWindowToAsk()) return runQuitFlow(application, timeoutMs)
  appGuard.quitApproved = true
  // If something else stops this quit, the guards must not stay switched off.
  let quitting = false
  const onQuit = () => { quitting = true }
  if (typeof application.once === 'function') application.once('quit', onQuit)
  const timer = setTimeout(() => {
    if (typeof application.removeListener === 'function') application.removeListener('quit', onQuit)
    if (!quitting && appGuard.app === application) {
      appGuard.quitApproved = false
      logger.warn('Quitting was stopped by another part of the app; closing windows asks about unsaved changes again.')
    }
  }, settings.quitConfirmMs)
  if (typeof timer.unref === 'function') timer.unref()
  application.quit()
  return true
}

/**
 * Makes quitting safe: before-quit is held until every pending write has
 * settled (at most 30 s), then each guarded window is asked in turn (focused
 * first) exactly as when it closes. A window is closed right after its
 * decision; Cancel in any window stops the quit and keeps the remaining
 * windows open.
 *
 * @param {object} [app] Electron's app (default: electron.app)
 * @param {{pendingWritesTimeoutMs?: number}} [options]
 * @returns {{isQuitApproved: () => boolean, quitFlow: () => Promise<boolean>|null}}
 */
function installAppGuard(app, options = {}) {
  const application = app || (electron() && electron().app)
  if (!application || typeof application.on !== 'function') throw new TypeError('installAppGuard needs Electron\'s app.')
  if (appGuard.installed && appGuard.app === application) return appGuard.api
  if (appGuard.installed && appGuard.app && appGuard.listener) appGuard.app.removeListener('before-quit', appGuard.listener)
  const timeoutMs = options.pendingWritesTimeoutMs ?? settings.pendingWritesTimeoutMs
  const listener = (event) => {
    if (appGuard.quitApproved) return
    event.preventDefault()
    if (appGuard.flow) return
    appGuard.flow = runQuitFlow(application, timeoutMs)
      .catch((error) => {
        logger.warn(`Could not finish quitting safely: ${error && error.message}`)
        return false
      })
      .finally(() => { appGuard.flow = null })
  }
  application.on('before-quit', listener)
  appGuard.installed = true
  appGuard.app = application
  appGuard.listener = listener
  appGuard.quitApproved = false
  appGuard.api = Object.freeze({
    isQuitApproved: () => appGuard.quitApproved,
    quitFlow: () => appGuard.flow,
  })
  return appGuard.api
}

/**
 * @typedef {object} IoResult
 * @property {boolean} ok
 * @property {string} [code] failure code from io-core IO_CODES (a file another Simple window is editing is
 *   LOCKED with reason 'open-in-another-window')
 * @property {string} [message] plain user-facing text
 * @property {string} [reason] machine-readable detail of a failure
 * @property {string} [technical] "<errno> <syscall> after N attempts" and similar, for "Details:"
 * @property {string} docId
 * @property {'save'|'save-as'|'save-copy'|'export'} mode
 * @property {number|null} revision echoed from the request
 * @property {string} [path]
 * @property {string} [name]
 * @property {string} [format]
 * @property {string} [formatLabel] e.g. "Excel workbook (.xlsx)"
 * @property {{size: number, mtimeMs: number, sha256: string}} [stamp]
 * @property {'rename'|'swap'|'in-place'|'new'|'unchanged'} [strategy] 'unchanged': an untouched Save needed no write
 * @property {string[]} [warnings] things the serializer could not include (user-facing)
 * @property {{format: string, formatLabel: string, lost: string[], lostShort: string,
 *   fullFormat: string|null, fullLabel: string|null}|null} [lossy] for Save/Save As: the document's lossy
 *   state after the save (kept until a full-format save); for Save a Copy: what the copy left out
 * @property {boolean} [folderChanged] the file went to another folder than the document's (or it had none)
 * @property {boolean} [rebound] the document now points at the new file
 * @property {boolean} [ownFileChanged] an export or copy replaced the document's own file, which no longer holds
 *   the document: the page marks the document changed so the next Save writes it
 * @property {boolean} [sibling] the edits went into a new file next to the original, which is unchanged
 * @property {string} [originalPath] with sibling: the unchanged original
 * @property {string} [originalName]
 * @property {string} [originalFormatLabel] with sibling: the format Simple could not update
 * @property {string} [backupPath] the versions-store copy taken before the first overwrite
 * @property {string} [backupFailed] why that copy could not be taken (the save still succeeded)
 * @property {{ok: boolean, id?: string, skipped?: string|null, reason?: string}|null} [versionBackup]
 */

module.exports = {
  REQUEST_TYPES,
  adoptDocument,
  bindDocument,
  canAdoptDocument,
  configureDocumentGuard,
  docKey,
  documentsForWindow,
  ensureDocument,
  findDocumentByPath,
  focusDocument,
  forgetDocument,
  getDocument,
  handleRendererResponse,
  installAppGuard,
  installWindowGuard,
  isSaving,
  openDocument,
  performSave,
  registerDocument,
  requestRenderer,
}
