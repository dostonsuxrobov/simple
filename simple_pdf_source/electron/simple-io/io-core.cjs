// Vendored from simple/shared/electron/io-core.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// Core of the shared Save / Import / Export layer:
// - IoError, the codes the UI understands, and the mapping from Node errors
// - file stamps (size, mtime, SHA-256) used to detect changes by other programs
// - folder probes (exists, really writable, free space)
// - the pending-write registry that quitting waits on
// - the temp journal and sweep(), which finish or undo a save that a crash
//   interrupted. sweep() only touches paths the journal recorded; it never
//   deletes files by name pattern.
// Only Node built-ins (and electron, when running inside it) are required.

const crypto = require('node:crypto')
const { constants: fsConstants } = require('node:fs')
const fsp = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const catalog = require('./io-catalog.json')

/** Every result code a shared I/O call can report. Save and open share one list. */
const IO_CODES = Object.freeze({
  CANCELED: 'CANCELED',
  BLOCKED: 'BLOCKED',
  LOCKED: 'LOCKED',
  READ_ONLY: 'READ_ONLY',
  NO_PERMISSION: 'NO_PERMISSION',
  READ_ONLY_VOLUME: 'READ_ONLY_VOLUME',
  FOLDER_MISSING: 'FOLDER_MISSING',
  FILE_UNAVAILABLE: 'FILE_UNAVAILABLE',
  DISK_FULL: 'DISK_FULL',
  NAME_TOO_LONG: 'NAME_TOO_LONG',
  INVALID_NAME: 'INVALID_NAME',
  CHANGED_ON_DISK: 'CHANGED_ON_DISK',
  SOURCE_MISSING: 'SOURCE_MISSING',
  VERIFY_FAILED: 'VERIFY_FAILED',
  VALIDATION_FAILED: 'VALIDATION_FAILED',
  SERIALIZE_FAILED: 'SERIALIZE_FAILED',
  NEEDS_OFFICE_ENGINE: 'NEEDS_OFFICE_ENGINE',
  RESTORE_NEEDED: 'RESTORE_NEEDED',
  NOT_FOUND: 'NOT_FOUND',
  DAMAGED: 'DAMAGED',
  ENCRYPTED: 'ENCRYPTED',
  TOO_LARGE: 'TOO_LARGE',
  UNSUPPORTED: 'UNSUPPORTED',
  UNKNOWN: 'UNKNOWN',
})

/**
 * Codes whose prompt offers Save As with the same name in the first writable
 * fallback folder (never the folder that failed, so a full drive is not offered again).
 */
const FALLBACK_CODES = Object.freeze(new Set([
  'LOCKED', 'READ_ONLY', 'NO_PERMISSION', 'READ_ONLY_VOLUME', 'FOLDER_MISSING', 'FILE_UNAVAILABLE', 'NAME_TOO_LONG', 'DISK_FULL',
]))

/** Codes whose Save As suggests "<stem> (edited).<ext>" in the same folder. */
const SAME_FOLDER_CODES = Object.freeze(new Set(['CHANGED_ON_DISK', 'SOURCE_MISSING']))

/** Errors that usually clear by themselves when another program lets go of a file. */
const TRANSIENT_ERRNOS = Object.freeze(new Set(['EPERM', 'EBUSY', 'EACCES']))

/** Errors meaning the file or its storage cannot be reached right now (offline drive, network, placeholder files). */
const UNAVAILABLE_ERRNOS = Object.freeze(new Set([
  'EIO', 'ETIMEDOUT', 'ECANCELED', 'UNKNOWN', 'EAGAIN', 'ENXIO', 'ENODEV', 'EHOSTDOWN', 'EHOSTUNREACH',
  'ENETDOWN', 'ENETUNREACH', 'ENETRESET', 'ECONNRESET', 'ECONNABORTED', 'ECONNREFUSED', 'ENOTCONN', 'ESTALE', 'EPIPE',
]))

/** Errors while reading back the temp file that mean the temp itself was taken away or held, not the folder. */
const VERIFY_STAGE_ERRNOS = Object.freeze(new Set(['ENOENT', 'ENOTDIR', 'EPERM', 'EACCES', 'EBUSY']))

/** Retry delays for replacing a file another program holds: about 3.55 s in total. */
const RETRY_DELAYS = Object.freeze([50, 100, 200, 400, 800, 1000, 1000])

/** Prefix of every temporary file Simple creates next to a user's file. */
const TEMP_PREFIX = '~simple-'

const MAX_NAME_UNITS = 255
const MAX_PATH_UNITS = 32_000
const HASH_CHUNK_BYTES = 1024 * 1024
const JOURNAL_FILE_PATTERN = /^journal-(\d+)-([0-9a-f]+)\.json$/
const JOURNAL_STALE_MS = 30 * 60 * 1000
const DAMAGED_JOURNAL_KEEP_MS = 30 * 24 * 60 * 60 * 1000
const RESERVED_WINDOWS_NAMES = /^(con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³])$/i
const INVALID_NAME_CHARACTERS = /[<>:"\/\\|?*\u0000-\u001f]/

let configuredJournalDir = null
let logger = {
  warn: (...args) => console.warn('[simple-io]', ...args),
  info: () => {},
}

// ---------------------------------------------------------------------------
// Catalog and messages
// ---------------------------------------------------------------------------

/**
 * Replaces {placeholders} in a catalog string. Missing values fall back to the
 * catalog's neutral wording ("this file"), never to an empty gap.
 * @param {string} template
 * @param {Record<string, unknown>} [vars]
 * @returns {string}
 */
function formatTemplate(template, vars = {}) {
  if (typeof template !== 'string') return ''
  return template.replace(/\{(\w+)\}/g, (match, key) => {
    const value = vars[key]
    if (value !== undefined && value !== null && value !== '') return String(value)
    const fallback = catalog.placeholders?.[key]
    return fallback !== undefined ? fallback : match
  })
}

/**
 * Looks up a catalog value by dotted key, such as "saveFailed.LOCKED" or "status.saving".
 * @param {string} key
 * @returns {any} the raw catalog value, or undefined
 */
function catalogEntry(key) {
  let node = catalog
  for (const part of String(key).split('.')) {
    if (!node || typeof node !== 'object' || !Object.prototype.hasOwnProperty.call(node, part)) return undefined
    node = node[part]
  }
  return node
}

/**
 * Formats a catalog prompt or message entry with variables.
 * @param {string} key dotted catalog key, for example "saveFailed.LOCKED" or "prompts.unsaved"
 * @param {Record<string, unknown>} [vars]
 * @returns {{message: string, detail: string, buttons: Array<{id: string, label: string}>, defaultId: string|null, cancelId: string|null}|null}
 */
function formatCatalogPrompt(key, vars = {}) {
  const entry = catalogEntry(key)
  if (!entry || typeof entry !== 'object') return null
  return {
    message: formatTemplate(entry.message, vars),
    detail: formatTemplate(entry.detail || '', vars),
    buttons: (entry.buttons || []).map((button) => ({ id: button.id, label: formatTemplate(button.label, vars) })),
    defaultId: entry.defaultId || null,
    cancelId: entry.cancelId || null,
  }
}

function errorVars(details) {
  const filePath = details.path || null
  return {
    name: details.fileName || (filePath ? path.basename(filePath) : undefined),
    folder: details.folder || (filePath ? path.dirname(filePath) : undefined),
    drive: details.drive,
    needed: typeof details.needed === 'number' ? formatBytes(details.needed) : details.needed,
    backupPath: details.backupPath || details.asidePath,
    technical: details.technical,
    reason: details.reason,
    formatLabel: details.formatLabel,
    altLabel: details.altLabel,
    altExt: details.altExt,
    limit: typeof details.limit === 'number' ? formatBytes(details.limit) : details.limit,
    ext: details.ext,
  }
}

function messageFor(code, details, context) {
  const tables = context === 'open' ? ['openFailed', 'saveFailed'] : ['saveFailed', 'openFailed']
  for (const table of tables) {
    const entry = catalog[table]?.[code]
    if (entry?.message) return formatTemplate(entry.message, errorVars(details))
  }
  if (typeof catalog.notices?.[code] === 'string') return formatTemplate(catalog.notices[code], errorVars(details))
  return formatTemplate(catalog[context === 'open' ? 'openFailed' : 'saveFailed'].UNKNOWN.message, errorVars(details))
}

// ---------------------------------------------------------------------------
// IoError
// ---------------------------------------------------------------------------

/** Details an IoError carries (and toIoError copies from other coded errors). */
const DETAIL_KEYS = Object.freeze([
  'path', 'fileName', 'folder', 'nearestFolder', 'needed', 'drive', 'reason', 'technical', 'backupPath', 'asidePath',
  'attempts', 'errno', 'syscall', 'formatLabel', 'altLabel', 'altExt', 'altFormat', 'limit', 'ext',
])

/**
 * The one error type of the shared I/O layer. `message` is plain, user-facing
 * text from io-catalog.json; `technical` carries the errno, syscall and attempts.
 */
class IoError extends Error {
  /**
   * @param {string} code one of IO_CODES
   * @param {object} [details]
   * @param {'save'|'open'} [details.context='save'] which catalog table words the message
   * @param {string} [details.path] the file the operation was about
   * @param {string} [details.fileName] display name (defaults to the base name of `path`)
   * @param {string} [details.folder] folder the operation was about
   * @param {string|null} [details.nearestFolder] closest folder that still exists (FOLDER_MISSING)
   * @param {number} [details.needed] bytes that must be freed (DISK_FULL)
   * @param {string} [details.drive] drive label such as "drive C:"
   * @param {string} [details.reason] short machine-readable or plain reason
   * @param {string} [details.technical] "<errno> <syscall> after N attempts"
   * @param {string} [details.backupPath] where the untouched original is kept (RESTORE_NEEDED)
   * @param {string} [details.asidePath] moved-aside original during a swap (RESTORE_NEEDED)
   * @param {number} [details.attempts]
   * @param {unknown} [details.cause]
   */
  constructor(code, details = {}) {
    const known = Object.prototype.hasOwnProperty.call(IO_CODES, code) ? code : 'UNKNOWN'
    const context = details.context === 'open' ? 'open' : 'save'
    super(messageFor(known, details, context), details.cause !== undefined ? { cause: details.cause } : undefined)
    this.name = 'IoError'
    this.code = known
    this.context = context
    for (const key of DETAIL_KEYS) {
      if (details[key] !== undefined) this[key] = details[key]
    }
    if (this.fileName === undefined && this.path) this.fileName = path.basename(this.path)
  }
}

/**
 * @param {unknown} value
 * @returns {value is IoError}
 */
function isIoError(value) {
  return value instanceof IoError || Boolean(value && value.name === 'IoError' && typeof value.code === 'string' && IO_CODES[value.code])
}

/**
 * Builds the "<errno> <syscall> after N attempts" line shown after "Details:".
 * @param {unknown} error
 * @param {number} [attempts]
 * @returns {string}
 */
function technicalDetail(error, attempts) {
  const parts = []
  const errno = error && typeof error.code === 'string' ? error.code : null
  if (errno) parts.push(errno)
  if (error && typeof error.syscall === 'string') parts.push(error.syscall)
  if (!parts.length && error) parts.push(error.name && error.name !== 'Error' ? error.name : 'Error')
  let text = parts.join(' ')
  if (attempts && attempts > 1) text += ` after ${attempts} attempts`
  const message = error && typeof error.message === 'string' ? error.message.trim() : ''
  if (message && !errno) text += `: ${message.length > 300 ? `${message.slice(0, 297)}…` : message}`
  return text
}

/**
 * Converts anything thrown by fs, a validator or a serializer into an IoError.
 * The stage decides what an ambiguous errno means: EPERM while creating the
 * temp file is a folder permission problem, while EPERM when replacing the
 * target (after the retry budget) means another program holds it. While the
 * temp file is read back ('verify'), a temp that vanished or stayed locked
 * (a virus scanner quarantined or held it) is VERIFY_FAILED: the folder is fine.
 *
 * @param {unknown} error
 * @param {object} [info]
 * @param {string} [info.path] file the operation was about
 * @param {'preflight'|'temp'|'verify'|'replace'|'in-place'|'read'|'probe'} [info.stage]
 * @param {number} [info.attempts]
 * @param {'save'|'open'} [info.context] defaults to 'open' for stage 'read', else 'save'
 * @returns {IoError}
 */
function toIoError(error, info = {}) {
  if (isIoError(error)) {
    if (!error.path && info.path) {
      error.path = info.path
      if (!error.fileName) error.fileName = path.basename(info.path)
    }
    return error
  }
  const stage = info.stage || 'replace'
  const context = info.context || (stage === 'read' ? 'open' : 'save')
  const errno = error && typeof error.code === 'string' ? error.code : null
  const attempts = info.attempts || (error && error.attempts) || undefined
  if (errno && errno !== 'UNKNOWN' && Object.prototype.hasOwnProperty.call(IO_CODES, errno)) {
    // A coded error from another shared module (office engine, a serializer):
    // keep its code and the details its prompt needs.
    const details = { context, path: info.path, attempts, cause: error }
    for (const key of DETAIL_KEYS) {
      if (error[key] !== undefined && details[key] === undefined) details[key] = error[key]
    }
    if (details.reason === undefined && error.message) details.reason = error.message
    if (details.technical === undefined) details.technical = `${errno}${error.message ? `: ${error.message}` : ''}`
    return new IoError(errno, details)
  }
  let code = 'UNKNOWN'
  if (error && (error.name === 'AbortError' || errno === 'ABORT_ERR')) code = 'CANCELED'
  else if (stage === 'verify' && context === 'save' && VERIFY_STAGE_ERRNOS.has(errno)) code = 'VERIFY_FAILED'
  else if (errno === 'ENOSPC' || errno === 'EDQUOT') code = 'DISK_FULL'
  else if (errno === 'ENAMETOOLONG') code = 'NAME_TOO_LONG'
  else if (errno === 'EROFS') code = 'READ_ONLY_VOLUME'
  else if (errno === 'EISDIR') code = context === 'open' ? 'UNSUPPORTED' : 'INVALID_NAME'
  else if (errno === 'ENOENT' || errno === 'ENOTDIR') {
    if (context === 'open') code = 'NOT_FOUND'
    else code = stage === 'replace' || stage === 'in-place' ? 'FILE_UNAVAILABLE' : 'FOLDER_MISSING'
  } else if (errno === 'EBUSY') code = context === 'open' ? 'LOCKED' : (stage === 'temp' || stage === 'probe' ? 'FILE_UNAVAILABLE' : 'LOCKED')
  else if (errno === 'EPERM' || errno === 'EACCES') {
    if (context === 'open' || stage === 'temp' || stage === 'preflight' || stage === 'probe') code = 'NO_PERMISSION'
    else code = 'LOCKED'
  } else if (errno && UNAVAILABLE_ERRNOS.has(errno)) code = 'FILE_UNAVAILABLE'
  else if (errno === 'ERR_FS_FILE_TOO_LARGE' || errno === 'ERR_STRING_TOO_LONG' || (error instanceof RangeError && /allocation|too large|invalid array length/i.test(error.message))) {
    code = context === 'open' ? 'TOO_LARGE' : 'UNKNOWN'
  }
  const filePath = info.path
  return new IoError(code, {
    context,
    path: filePath,
    folder: filePath ? path.dirname(filePath) : undefined,
    drive: filePath && (code === 'DISK_FULL' || code === 'READ_ONLY_VOLUME') ? describeDrive(filePath) : undefined,
    technical: technicalDetail(error, attempts),
    attempts,
    errno: errno || undefined,
    syscall: error && typeof error.syscall === 'string' ? error.syscall : undefined,
    cause: error,
  })
}

/**
 * Serializable failure result for IPC (`IoResult` with ok:false). Expected
 * failures are returned, never thrown across IPC.
 * @param {unknown} error
 * @param {object} [info] passed to toIoError when `error` is not an IoError yet
 * @returns {{ok: false, code: string, message: string, name?: string, path?: string, folder?: string,
 *   nearestFolder?: string, needed?: string, neededBytes?: number, drive?: string, reason?: string,
 *   technical?: string, backupPath?: string, asidePath?: string}}
 */
function toIoResult(error, info) {
  const ioError = toIoError(error, info)
  const result = { ok: false, code: ioError.code, message: ioError.message }
  if (ioError.context === 'open') result.context = 'open'
  if (ioError.fileName) result.name = ioError.fileName
  for (const key of ['path', 'folder', 'nearestFolder', 'drive', 'reason', 'technical', 'backupPath', 'asidePath', 'formatLabel', 'altLabel', 'altExt', 'altFormat', 'ext']) {
    if (ioError[key] !== undefined && ioError[key] !== null) result[key] = ioError[key]
  }
  if (typeof ioError.limit === 'number') result.limit = formatBytes(ioError.limit)
  if (typeof ioError.needed === 'number') {
    result.needed = formatBytes(ioError.needed)
    result.neededBytes = ioError.needed
  }
  return result
}

/**
 * The catalog prompt for a failure: message, detail ending with "Details: …",
 * buttons, default and cancel ids. Used by the native decision dialogs.
 * @param {unknown} error an IoError, an IoResult with ok:false, or anything thrown
 * @returns {{code: string, message: string, detail: string, buttons: Array<{id: string, label: string}>, defaultId: string|null, cancelId: string|null}}
 */
function describeIoError(error) {
  const source = error && error.ok === false && typeof error.code === 'string'
    ? { ...error, fileName: error.name, needed: error.neededBytes }
    : toIoError(error)
  const code = Object.prototype.hasOwnProperty.call(IO_CODES, source.code) ? source.code : 'UNKNOWN'
  const context = source.context === 'open' ? 'openFailed' : 'saveFailed'
  const key = catalog[context][code] ? `${context}.${code}` : `${context}.UNKNOWN`
  const vars = errorVars(source)
  const prompt = formatCatalogPrompt(key, vars)
  const technical = source.technical ? formatTemplate(catalog.detailsLine, { technical: source.technical }) : ''
  return {
    code,
    message: prompt.message,
    detail: [prompt.detail, technical].filter(Boolean).join('\n\n'),
    buttons: prompt.buttons,
    defaultId: prompt.defaultId,
    cancelId: prompt.cancelId,
  }
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

/**
 * @param {number} milliseconds
 * @returns {Promise<void>}
 */
function sleep(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, Math.max(0, milliseconds)))
}

/**
 * Human-readable size for messages ("1.4 MB").
 * @param {number} bytes
 * @returns {string}
 */
function formatBytes(bytes) {
  const value = Number(bytes)
  if (!Number.isFinite(value) || value < 0) return ''
  if (value < 1024) return `${Math.round(value)} bytes`
  const units = ['KB', 'MB', 'GB', 'TB']
  let scaled = value / 1024
  let unit = 0
  while (scaled >= 1024 && unit < units.length - 1) {
    scaled /= 1024
    unit += 1
  }
  const digits = scaled < 10 ? 1 : 0
  return `${scaled.toFixed(digits).replace(/\.0$/, '')} ${units[unit]}`
}

/**
 * A drive label for messages: "drive C:", or the network share "\\server\share".
 * @param {string} filePath
 * @returns {string}
 */
function describeDrive(filePath) {
  const root = path.parse(path.resolve(String(filePath || '.'))).root
  if (/^[a-z]:[\\/]?$/i.test(root)) return `drive ${root.slice(0, 2).toUpperCase()}`
  if (/^[\\/]{2}[^\\/]+[\\/][^\\/]+/.test(root)) return root.replace(/[\\/]+$/, '').replace(/\//g, '\\')
  return 'this drive'
}

/**
 * Free bytes the save needs before it starts: the new file plus 5 % and 1 MiB,
 * because the temp file and the original coexist until the replace.
 * @param {number} size
 * @returns {number}
 */
function requiredFreeBytes(size) {
  return Math.ceil(Math.max(0, size) * 1.05) + 1024 * 1024
}

/**
 * Checks one file-name component against what Windows (and Simple's own
 * temp naming) can store.
 * @param {string} name
 * @returns {{ok: true} | {ok: false, code: 'INVALID_NAME'|'NAME_TOO_LONG', reason: string}}
 */
function checkFileName(name) {
  const value = String(name ?? '')
  if (!value || value === '.' || value === '..') return { ok: false, code: 'INVALID_NAME', reason: 'empty' }
  if (value.length > MAX_NAME_UNITS) return { ok: false, code: 'NAME_TOO_LONG', reason: 'name-too-long' }
  if (INVALID_NAME_CHARACTERS.test(value)) return { ok: false, code: 'INVALID_NAME', reason: 'characters' }
  if (/[. ]$/.test(value)) return { ok: false, code: 'INVALID_NAME', reason: 'trailing-dot-or-space' }
  const stem = value.split('.')[0].trimEnd()
  if (RESERVED_WINDOWS_NAMES.test(stem)) return { ok: false, code: 'INVALID_NAME', reason: 'reserved' }
  return { ok: true }
}

/**
 * Checks a whole target path (length, name) before anything is written.
 * @param {string} filePath absolute path
 * @returns {{ok: true} | {ok: false, code: 'INVALID_NAME'|'NAME_TOO_LONG', reason: string}}
 */
function checkTargetPath(filePath) {
  if (String(filePath).length > MAX_PATH_UNITS) return { ok: false, code: 'NAME_TOO_LONG', reason: 'path-too-long' }
  return checkFileName(path.basename(filePath))
}

/**
 * A fresh "~simple-<8 hex>.<extension>" name: at most 20 characters, so very
 * long target names never push the temp path over a limit.
 * @param {string} [extension='tmp']
 * @returns {string}
 */
function tempFileName(extension = 'tmp') {
  return `${TEMP_PREFIX}${crypto.randomBytes(4).toString('hex')}.${extension}`
}

/**
 * Runs `operation` and retries it while it fails with a transient error.
 * @template T
 * @param {(attempt: number) => Promise<T>} operation
 * @param {object} [options]
 * @param {readonly number[]} [options.delays=RETRY_DELAYS] delay before each retry
 * @param {(error: any) => boolean} [options.isTransient] default: EPERM, EBUSY, EACCES
 * @param {(error: any, attempt: number) => void} [options.onRetry]
 * @returns {Promise<{value: T, attempts: number}>} rejects with the last error; `error.attempts` is set
 */
async function retryTransient(operation, options = {}) {
  const delays = options.delays || RETRY_DELAYS
  const isTransient = options.isTransient || ((error) => TRANSIENT_ERRNOS.has(error && error.code))
  for (let attempt = 1; ; attempt += 1) {
    try {
      return { value: await operation(attempt), attempts: attempt }
    } catch (error) {
      if (!isTransient(error) || attempt > delays.length) {
        if (error && typeof error === 'object') {
          try { error.attempts = attempt } catch {}
        }
        throw error
      }
      if (options.onRetry) options.onRetry(error, attempt)
      await sleep(delays[attempt - 1])
    }
  }
}

/**
 * Milliseconds to stall each write before its replace step, for acceptance
 * tests (SIMPLE_QA_WRITE_DELAY). Always 0 in a packaged app.
 * @returns {number}
 */
function qaWriteDelayMs() {
  const raw = process.env.SIMPLE_QA_WRITE_DELAY
  if (!raw) return 0
  const app = electronApp()
  if (app && app.isPackaged) return 0
  const value = Number(raw)
  return Number.isFinite(value) && value > 0 ? Math.min(value, 120_000) : 0
}

function electronApp() {
  if (!process.versions.electron) return null
  try {
    const electron = require('electron')
    return electron && electron.app && typeof electron.app.getPath === 'function' ? electron.app : null
  } catch {
    return null
  }
}

// ---------------------------------------------------------------------------
// Hashes and stamps
// ---------------------------------------------------------------------------

/**
 * @param {Buffer|Uint8Array|string} bytes
 * @returns {string} lowercase hex SHA-256
 */
function hashBytes(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex')
}

/**
 * Streams a file through SHA-256 without loading it whole.
 * @param {string} filePath
 * @param {object} [options]
 * @param {object} [options.fs] fs/promises-compatible implementation (tests inject failures)
 * @param {AbortSignal} [options.signal]
 * @returns {Promise<{sha256: string, size: number}>}
 */
async function hashFile(filePath, options = {}) {
  const fsImpl = options.fs || fsp
  const handle = await fsImpl.open(filePath, 'r')
  try {
    const hash = crypto.createHash('sha256')
    const buffer = Buffer.allocUnsafe(HASH_CHUNK_BYTES)
    let size = 0
    for (;;) {
      if (options.signal?.aborted) throw new IoError('CANCELED', { path: filePath })
      const { bytesRead } = await handle.read(buffer, 0, buffer.length, size)
      if (!bytesRead) break
      hash.update(buffer.subarray(0, bytesRead))
      size += bytesRead
    }
    return { sha256: hash.digest('hex'), size }
  } finally {
    await handle.close().catch(() => {})
  }
}

/**
 * The identity of a file's content at a moment: size, mtime and (optionally) SHA-256.
 * @param {string} filePath
 * @param {{hash?: boolean, fs?: object}} [options] hash defaults to true
 * @returns {Promise<{size: number, mtimeMs: number, sha256?: string}>}
 */
async function stampFile(filePath, options = {}) {
  const fsImpl = options.fs || fsp
  const stat = await fsImpl.stat(filePath)
  if (options.hash === false) return { size: stat.size, mtimeMs: stat.mtimeMs }
  const { sha256, size } = await hashFile(filePath, { fs: fsImpl })
  return { size, mtimeMs: stat.mtimeMs, sha256 }
}

/**
 * A stamp for bytes that were just read from or written to `stat`'s file.
 * @param {Buffer|Uint8Array} bytes
 * @param {{mtimeMs: number}} stat
 * @returns {{size: number, mtimeMs: number, sha256: string}}
 */
function stampFromBytes(bytes, stat) {
  return { size: bytes.length, mtimeMs: stat.mtimeMs, sha256: hashBytes(bytes) }
}

/**
 * Compares a file on disk with a stamp captured earlier.
 * - 'same': size and mtime are unchanged (fast path, no hashing)
 * - 'restamped': size or mtime changed but the content hash is equal (a sync
 *   program, backup tool or virus scanner touched it); `stamp` is the fresh one
 * - 'changed': the content differs (or no hash was recorded to prove otherwise)
 * - 'missing': the file is gone
 * @param {string} filePath
 * @param {{size: number, mtimeMs: number, sha256?: string}} expected
 * @param {{fs?: object}} [options]
 * @returns {Promise<{state: 'same'|'restamped'|'changed'|'missing', stamp: object|null}>}
 */
async function compareStamp(filePath, expected, options = {}) {
  const fsImpl = options.fs || fsp
  let stat
  try {
    stat = await fsImpl.stat(filePath)
  } catch (error) {
    if (error.code === 'ENOENT' || error.code === 'ENOTDIR') return { state: 'missing', stamp: null }
    throw error
  }
  if (stat.size === expected.size && stat.mtimeMs === expected.mtimeMs) {
    return { state: 'same', stamp: { size: stat.size, mtimeMs: stat.mtimeMs, ...(expected.sha256 ? { sha256: expected.sha256 } : {}) } }
  }
  if (!expected.sha256) return { state: 'changed', stamp: { size: stat.size, mtimeMs: stat.mtimeMs } }
  const { sha256, size } = await hashFile(filePath, { fs: fsImpl })
  const stamp = { size, mtimeMs: stat.mtimeMs, sha256 }
  return { state: sha256 === expected.sha256 ? 'restamped' : 'changed', stamp }
}

/**
 * Reads a whole file for opening and stamps it in the same pass. Briefly
 * retries while another program is writing it. Failures are IoErrors worded
 * for opening (NOT_FOUND only when the folder still exists).
 * @param {string} filePath
 * @param {object} [options]
 * @param {number} [options.maxBytes] larger files fail with TOO_LARGE
 * @param {readonly number[]} [options.retryDelays]
 * @param {object} [options.fs]
 * @returns {Promise<{bytes: Buffer, stamp: {size: number, mtimeMs: number, sha256: string}}>}
 */
async function readFileStamped(filePath, options = {}) {
  const fsImpl = options.fs || fsp
  const target = path.resolve(String(filePath))
  try {
    const { value } = await retryTransient(async () => {
      const handle = await fsImpl.open(target, 'r')
      try {
        const stat = await handle.stat()
        if (options.maxBytes !== undefined && stat.size > options.maxBytes) {
          throw new IoError('TOO_LARGE', { context: 'open', path: target, limit: options.maxBytes })
        }
        const bytes = await handle.readFile()
        return { bytes, stamp: stampFromBytes(bytes, stat) }
      } finally {
        await handle.close().catch(() => {})
      }
    }, { delays: options.retryDelays || [50, 100, 200, 400, 800], isTransient: (error) => error && error.code === 'EBUSY' })
    return value
  } catch (error) {
    if (!isIoError(error) && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) {
      const folder = path.dirname(target)
      if (!(await isDirectory(folder, fsImpl))) {
        throw new IoError('FOLDER_MISSING', {
          context: 'open', path: target, folder, nearestFolder: await nearestExistingFolder(folder, { fs: fsImpl }), technical: technicalDetail(error), cause: error,
        })
      }
    }
    throw toIoError(error, { path: target, stage: 'read', context: 'open' })
  }
}

// ---------------------------------------------------------------------------
// Folder probes
// ---------------------------------------------------------------------------

async function isDirectory(folder, fsImpl = fsp) {
  try {
    return (await fsImpl.stat(folder)).isDirectory()
  } catch {
    return false
  }
}

/**
 * The closest folder above (or at) `folder` that still exists, for the
 * FOLDER_MISSING prompt.
 * @param {string} folder
 * @param {{fs?: object}} [options]
 * @returns {Promise<string|null>}
 */
async function nearestExistingFolder(folder, options = {}) {
  const fsImpl = options.fs || fsp
  let current = path.resolve(String(folder))
  for (;;) {
    if (await isDirectory(current, fsImpl)) return current
    const parent = path.dirname(current)
    if (parent === current) return null
    current = parent
  }
}

/**
 * Free and total bytes of the volume holding `folder`, or null when the
 * platform or file system cannot say.
 * @param {string} folder
 * @param {{fs?: object}} [options]
 * @returns {Promise<{free: number, total: number}|null>}
 */
async function freeSpace(folder, options = {}) {
  const fsImpl = options.fs || fsp
  if (typeof fsImpl.statfs !== 'function') return null
  try {
    const info = await fsImpl.statfs(folder)
    const free = Number(info.bavail) * Number(info.bsize)
    const total = Number(info.blocks) * Number(info.bsize)
    return Number.isFinite(free) ? { free, total } : null
  } catch {
    return null
  }
}

/**
 * Whether Simple can really create files in a folder. fs.access(W_OK) ignores
 * Windows ACLs, so this creates and deletes a 0-byte "~simple-probe-<hex>.tmp".
 * @param {string} folder
 * @param {{fs?: object}} [options]
 * @returns {Promise<{folder: string, exists: boolean, writable: boolean, code?: string, nearestFolder?: string|null, technical?: string}>}
 */
async function probeFolder(folder, options = {}) {
  const fsImpl = options.fs || fsp
  const resolved = path.resolve(String(folder))
  if (!(await isDirectory(resolved, fsImpl))) {
    return { folder: resolved, exists: false, writable: false, code: 'FOLDER_MISSING', nearestFolder: await nearestExistingFolder(resolved, { fs: fsImpl }) }
  }
  const probe = path.join(resolved, `${TEMP_PREFIX}probe-${crypto.randomBytes(4).toString('hex')}.tmp`)
  try {
    const handle = await fsImpl.open(probe, 'wx')
    await handle.close()
  } catch (error) {
    const ioError = toIoError(error, { path: probe, stage: 'probe' })
    return { folder: resolved, exists: true, writable: false, code: ioError.code, technical: ioError.technical }
  }
  await removeFile(probe, fsImpl)
  return { folder: resolved, exists: true, writable: true }
}

/**
 * Deletes one file, retrying briefly while a scanner holds it.
 * @param {string} filePath
 * @param {object} [fsImpl] fs/promises-compatible implementation
 * @returns {Promise<boolean>} true when the file is gone (or never existed)
 */
async function removeFile(filePath, fsImpl = fsp) {
  try {
    await retryTransient(() => fsImpl.unlink(filePath), { delays: [50, 100, 200] })
    return true
  } catch (error) {
    if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) return true
    return false
  }
}

/**
 * Whether anything (file, folder or link) exists at a path.
 * @param {string} filePath
 * @param {object} [fsImpl] fs/promises-compatible implementation
 * @returns {Promise<boolean>} rejects for errors other than "not found"
 */
async function pathExists(filePath, fsImpl = fsp) {
  try {
    await fsImpl.lstat(filePath)
    return true
  } catch (error) {
    if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) return false
    throw error
  }
}

// ---------------------------------------------------------------------------
// Pending writes (quit waits for these)
// ---------------------------------------------------------------------------

const pendingWrites = new Map()

/**
 * Registers a write so quitting can wait for it. Returns the same promise.
 * @template T
 * @param {Promise<T>} promise
 * @param {string} [label] usually the target path
 * @returns {Promise<T>}
 */
function trackPendingWrite(promise, label = '') {
  const token = Symbol(label)
  pendingWrites.set(token, { promise, label: String(label), startedAt: Date.now() })
  const release = () => { pendingWrites.delete(token) }
  promise.then(release, release)
  return promise
}

/** @returns {number} writes that have not settled yet */
function pendingWriteCount() {
  return pendingWrites.size
}

/** @returns {Array<{label: string, startedAt: number}>} */
function listPendingWrites() {
  return [...pendingWrites.values()].map(({ label, startedAt }) => ({ label, startedAt }))
}

/**
 * Waits until every registered write has settled, including writes that start
 * while waiting, or until the timeout.
 * @param {{timeoutMs?: number}} [options] default 30 s
 * @returns {Promise<{settled: boolean, remaining: number}>}
 */
async function waitForPendingWrites(options = {}) {
  const deadline = Date.now() + (options.timeoutMs ?? 30_000)
  while (pendingWrites.size) {
    const left = deadline - Date.now()
    if (left <= 0) return { settled: false, remaining: pendingWrites.size }
    let timer
    const timeout = new Promise((resolve) => { timer = setTimeout(resolve, left) })
    await Promise.race([Promise.allSettled([...pendingWrites.values()].map((entry) => entry.promise)), timeout])
    clearTimeout(timer)
  }
  return { settled: true, remaining: 0 }
}

// ---------------------------------------------------------------------------
// Temp journal
// ---------------------------------------------------------------------------
// Each process keeps its own journal file, journal-<pid>-<token>.json, in the
// journal folder (default <userData>/io-journal). It records every temp file,
// moved-aside original and in-place backup before it exists, so sweep() at the
// next start can finish or undo the save and remove only recorded leftovers.

const processToken = crypto.randomBytes(4).toString('hex')
const journalState = { entries: new Map(), queue: Promise.resolve(), file: null, lastError: null }

/**
 * Configures where the journal lives and where problems are logged.
 * @param {object} [options]
 * @param {string|null} [options.journalDir] folder for journal files; null restores the default
 * @param {{warn: Function, info?: Function}} [options.logger]
 */
function configureIo(options = {}) {
  if (Object.prototype.hasOwnProperty.call(options, 'journalDir')) {
    configuredJournalDir = options.journalDir ? path.resolve(options.journalDir) : null
    journalState.file = null
  }
  if (options.logger) logger = { warn: options.logger.warn || (() => {}), info: options.logger.info || (() => {}) }
}

/**
 * The journal folder: configureIo(), else SIMPLE_IO_JOURNAL_DIR, else
 * <userData>/io-journal inside Electron, else <tmp>/simple-io-journal.
 * @returns {string}
 */
function journalDir() {
  if (configuredJournalDir) return configuredJournalDir
  if (process.env.SIMPLE_IO_JOURNAL_DIR) return path.resolve(process.env.SIMPLE_IO_JOURNAL_DIR)
  const app = electronApp()
  if (app) {
    try { return path.join(app.getPath('userData'), 'io-journal') } catch {}
  }
  return path.join(os.tmpdir(), 'simple-io-journal')
}

function ownJournalFile() {
  const dir = journalDir()
  if (!journalState.file || path.dirname(journalState.file) !== dir) {
    journalState.file = path.join(dir, `journal-${process.pid}-${processToken}.json`)
  }
  return journalState.file
}

/** Writes a small internal file atomically (temp, fsync, rename). Never used for user files. */
async function writeInternalFile(filePath, text) {
  await fsp.mkdir(path.dirname(filePath), { recursive: true })
  const temporary = `${filePath}.${crypto.randomBytes(3).toString('hex')}.tmp`
  const handle = await fsp.open(temporary, 'wx')
  try {
    await handle.writeFile(text, 'utf8')
    await handle.sync()
  } finally {
    await handle.close()
  }
  try {
    await retryTransient(() => fsp.rename(temporary, filePath), { delays: [20, 50, 100, 200, 400] })
  } catch (error) {
    await fsp.unlink(temporary).catch(() => {})
    throw error
  }
}

function persistJournal() {
  const file = ownJournalFile()
  const entries = [...journalState.entries.values()]
  if (!entries.length) return fsp.unlink(file).catch((error) => { if (error.code !== 'ENOENT') throw error })
  const payload = { schema: 1, pid: process.pid, token: processToken, updatedAt: new Date().toISOString(), entries }
  return writeInternalFile(file, `${JSON.stringify(payload, null, 2)}\n`)
}

function enqueueJournal(mutate) {
  const run = journalState.queue.then(async () => {
    mutate()
    await persistJournal()
    journalState.lastError = null
    return true
  })
  journalState.queue = run.catch(() => {})
  return run.catch((error) => {
    journalState.lastError = error || new Error('unknown journal error')
    logger.warn(`Could not update the save journal in ${journalDir()}: ${error && error.message}`)
    return false
  })
}

function sameTarget(a, b) {
  const left = path.resolve(String(a || ''))
  const right = path.resolve(String(b || ''))
  return process.platform === 'win32' ? left.toLowerCase() === right.toLowerCase() : left === right
}

/**
 * The journal of the current process. Every method resolves to true when the
 * entry reached disk; a journal failure is logged and never fails a save.
 */
const journal = Object.freeze({
  /**
   * @param {object} entry {target, temp, phase, sha256?, size?, aside?, backup?, keepBackup?}
   * @returns {Promise<string>} the entry id
   */
  async add(entry) {
    const id = crypto.randomBytes(6).toString('hex')
    const now = new Date().toISOString()
    await enqueueJournal(() => journalState.entries.set(id, { ...entry, id, startedAt: now, updatedAt: now }))
    return id
  },
  /**
   * @param {string} id
   * @param {object} patch fields to merge
   * @returns {Promise<boolean>}
   */
  update(id, patch) {
    return enqueueJournal(() => {
      const current = journalState.entries.get(id)
      if (current) journalState.entries.set(id, { ...current, ...patch, updatedAt: new Date().toISOString() })
    })
  },
  /**
   * @param {string} id
   * @returns {Promise<boolean>}
   */
  remove(id) {
    return enqueueJournal(() => journalState.entries.delete(id))
  },
  /**
   * Marks the kept entries of earlier, failed saves of `target` (RESTORE_NEEDED)
   * as superseded after a later save of the same file succeeded: the file now
   * holds newer work, so sweep() must never put the older original back over
   * it, and only removes those entries' leftover files.
   * @param {string} target
   * @param {string} [exceptId] the entry of the save that just succeeded
   * @returns {Promise<boolean>}
   */
  supersede(target, exceptId) {
    const affected = [...journalState.entries.values()].filter((entry) => entry.id !== exceptId && entry.restoreNeeded && !entry.superseded && sameTarget(entry.target, target))
    if (!affected.length) return Promise.resolve(true)
    return enqueueJournal(() => {
      const now = new Date().toISOString()
      for (const entry of affected) {
        const current = journalState.entries.get(entry.id)
        if (current) journalState.entries.set(entry.id, { ...current, superseded: true, updatedAt: now })
      }
    })
  },
  /** @returns {object[]} copies of the open entries of this process */
  entries() {
    return [...journalState.entries.values()].map((entry) => ({ ...entry }))
  },
  /** @returns {Error|null} why the last journal update failed, or null when it succeeded */
  lastError() {
    return journalState.lastError || null
  },
  /** @returns {string} this process's journal file */
  file() {
    return ownJournalFile()
  },
})

/**
 * Whether this process holds Electron's single-instance lock and `dir` lies in
 * its userData folder, so no other live process can write journals there.
 */
function holdsJournalFolderLock(dir) {
  const app = electronApp()
  if (!app || typeof app.hasSingleInstanceLock !== 'function') return false
  try {
    if (!app.hasSingleInstanceLock()) return false
    const userData = path.resolve(app.getPath('userData'))
    const relative = path.relative(userData, path.resolve(dir))
    return relative === '' || (!relative.startsWith('..') && !path.isAbsolute(relative))
  } catch {
    return false
  }
}

function processIsAlive(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false
  if (pid === process.pid) return true
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return error && error.code === 'EPERM'
  }
}

async function uniqueSiblingPath(target, label) {
  const folder = path.dirname(target)
  const extension = path.extname(target)
  const stem = path.basename(target, extension)
  for (let index = 1; index < 1000; index += 1) {
    const suffix = index === 1 ? ` (${label})` : ` (${label} ${index})`
    const candidate = path.join(folder, `${stem}${suffix}${extension}`)
    if (!(await pathExists(candidate))) return candidate
  }
  return path.join(folder, `${stem} (${label} ${Date.now()})${extension}`)
}

async function sameHash(filePath, sha256) {
  if (!sha256) return false
  try {
    return (await hashFile(filePath)).sha256 === sha256
  } catch {
    return false
  }
}

async function copyInto(sourcePath, targetPath) {
  // Writes the backup's bytes into the target in place, keeping its identity.
  const source = await fsp.open(sourcePath, 'r')
  let target
  try {
    target = await fsp.open(targetPath, 'r+').catch(async (error) => {
      if (error.code === 'ENOENT') return fsp.open(targetPath, 'wx')
      throw error
    })
    const buffer = Buffer.allocUnsafe(HASH_CHUNK_BYTES)
    let position = 0
    for (;;) {
      const { bytesRead } = await source.read(buffer, 0, buffer.length, position)
      if (!bytesRead) break
      await target.write(buffer, 0, bytesRead, position)
      position += bytesRead
    }
    await target.truncate(position)
    await target.sync()
  } finally {
    await source.close().catch(() => {})
    if (target) await target.close().catch(() => {})
  }
}

const MAX_INTERRUPTED_COMPARE_BYTES = 64 * 1024 * 1024

/**
 * Whether `target` is exactly what an interrupted in-place write leaves: the
 * verified new content (`temp`, proven by `sha256`) up to some offset and the
 * original (`backup`) after it, or the original up to some offset and the new
 * content after it (the original was being written back). Anything else,
 * such as newer work saved later, answers false. Files over 64 MB are not
 * compared and answer false.
 */
async function isInterruptedWrite(target, temp, sha256, backup) {
  if (!temp || !sha256 || !backup) return false
  let parts
  try {
    const sizes = await Promise.all([target, temp, backup].map(async (item) => (await fsp.stat(item)).size))
    if (sizes.some((size) => size > MAX_INTERRUPTED_COMPARE_BYTES)) return false
    parts = await Promise.all([target, temp, backup].map((item) => fsp.readFile(item)))
  } catch {
    return false
  }
  const [current, written, original] = parts
  if (hashBytes(written) !== sha256) return false
  return isPrefixMix(current, written, original) || isPrefixMix(current, original, written)
}

/** current === first[0:k] + second[k:] for some k, with the length a sequential write leaves. */
function isPrefixMix(current, first, second) {
  let k = 0
  const limit = Math.min(current.length, first.length)
  while (k < limit && current[k] === first[k]) k += 1
  if (k >= second.length) return current.length === k
  if (current.length !== second.length) return false
  return current.subarray(k).equals(second.subarray(k))
}

/**
 * Brings one interrupted save to a safe end. Returns the outcome; 'retry-later'
 * keeps the entry for the next sweep.
 */
async function recoverEntry(entry) {
  const { target, temp, aside, backup } = entry
  const removeAll = async (...paths) => {
    let ok = true
    for (const item of paths) if (item) ok = (await removeFile(item)) && ok
    return ok
  }
  try {
    if (entry.phase === 'writing' || entry.phase === 'verified') {
      return (await removeAll(temp)) ? { outcome: 'removed-temp' } : { outcome: 'retry-later' }
    }
    if (entry.phase === 'cleanup') {
      return (await removeAll(temp, aside)) ? { outcome: 'cleaned' } : { outcome: 'retry-later' }
    }
    if (entry.superseded && (entry.phase === 'swapping' || entry.phase === 'in-place')) {
      // A later save of the same file succeeded, so the file holds newer work
      // than this failed save's original: only the leftovers go.
      const leftovers = [temp, aside, entry.keepBackup ? null : backup]
      return (await removeAll(...leftovers)) ? { outcome: 'superseded' } : { outcome: 'retry-later' }
    }
    if (entry.phase === 'swapping') {
      const targetExists = await pathExists(target)
      const asideExists = aside ? await pathExists(aside) : false
      const tempVerified = temp ? await sameHash(temp, entry.sha256) : false
      if (!targetExists && entry.restoreNeeded && tempVerified && asideExists) {
        // The user was told the original is safe at the aside path: the save is
        // finished as the user asked, and that original is kept under a visible
        // name instead of being deleted.
        await retryTransient(() => fsp.rename(temp, target))
        const keptOriginal = await uniqueSiblingPath(target, 'original')
        await retryTransient(() => fsp.rename(aside, keptOriginal))
        return { outcome: 'completed', detail: keptOriginal }
      }
      if (!targetExists) {
        if (tempVerified) {
          await retryTransient(() => fsp.rename(temp, target))
          await removeAll(aside)
          return { outcome: 'completed' }
        }
        if (asideExists) {
          await retryTransient(() => fsp.rename(aside, target))
          await removeAll(temp)
          return { outcome: 'rolled-back' }
        }
        if (temp && (await pathExists(temp))) {
          const kept = await uniqueSiblingPath(target, 'recovered')
          await fsp.rename(temp, kept)
          return { outcome: 'kept-copy', detail: kept }
        }
        return { outcome: 'nothing-to-restore' }
      }
      if (await sameHash(target, entry.sha256)) {
        await removeAll(aside, temp)
        return { outcome: 'completed' }
      }
      if (asideExists) {
        // The target holds something else (it was recreated after the crash):
        // keep every version under a visible name instead of deleting any.
        const keptOriginal = await uniqueSiblingPath(target, 'original')
        await fsp.rename(aside, keptOriginal)
        let keptEdit = null
        if (tempVerified) {
          keptEdit = await uniqueSiblingPath(target, 'unsaved changes')
          await fsp.rename(temp, keptEdit)
        } else {
          await removeAll(temp)
        }
        return { outcome: 'kept-both', detail: [keptOriginal, keptEdit].filter(Boolean).join(' | ') }
      }
      await removeAll(temp)
      return { outcome: 'cleaned' }
    }
    if (entry.phase === 'backing-up') {
      // The crash came while the original was being copied: the target was not
      // touched yet, and the backup may be incomplete.
      return (await removeAll(temp, backup)) ? { outcome: 'removed-temp' } : { outcome: 'retry-later' }
    }
    if (entry.phase === 'in-place') {
      const targetExists = await pathExists(target)
      const backupExists = backup ? await pathExists(backup) : false
      if (targetExists && (await sameHash(target, entry.sha256))) {
        if (backupExists && !entry.keepBackup) await removeAll(backup)
        await removeAll(temp)
        return { outcome: 'completed' }
      }
      if (targetExists && (await sameHash(target, entry.originalSha256))) {
        if (backupExists && !entry.keepBackup) await removeAll(backup)
        await removeAll(temp)
        return { outcome: 'untouched' }
      }
      if (backupExists && targetExists && (await sameHash(backup, entry.originalSha256)) && (await isInterruptedWrite(target, temp, entry.sha256, backup))) {
        // Byte for byte the file is the new content up to some point and the
        // original after it (or the other way round while the original was
        // being written back): this save's own interrupted write. Put the original back.
        await copyInto(backup, target)
        if (!(await sameHash(target, entry.originalSha256))) return { outcome: 'failed', detail: `the restored file does not match the backup at ${backup}` }
        if (!entry.keepBackup) await removeAll(backup)
        await removeAll(temp)
        return { outcome: 'rolled-back' }
      }
      if (backupExists && targetExists && (await sameHash(backup, entry.originalSha256))) {
        // The file holds neither the interrupted write nor the original: newer
        // work saved by Simple or another program, or a write that can't be
        // proven to be this one. Its content is unknown, so it is never
        // overwritten; the original is kept beside it under a visible name.
        const keptOriginal = await uniqueSiblingPath(target, 'original')
        await fsp.copyFile(backup, keptOriginal, fsConstants.COPYFILE_EXCL)
        if (!(await sameHash(keptOriginal, entry.originalSha256))) {
          await removeAll(keptOriginal)
          return { outcome: 'failed', detail: `the copy of the original does not match the backup at ${backup}` }
        }
        if (!entry.keepBackup) await removeAll(backup)
        await removeAll(temp)
        return { outcome: 'kept-original', detail: keptOriginal }
      }
      if (backupExists && (await sameHash(backup, entry.originalSha256))) {
        // The file is gone: the original goes back where it was.
        await copyInto(backup, target)
        if (!(await sameHash(target, entry.originalSha256))) return { outcome: 'failed', detail: `the restored file does not match the backup at ${backup}` }
        if (!entry.keepBackup) await removeAll(backup)
        await removeAll(temp)
        return { outcome: 'rolled-back' }
      }
      // Nothing proven complete to restore from: keep every file where it is.
      return { outcome: 'failed', detail: backupExists ? `the backup at ${backup} does not match the original` : 'no backup was found' }
    }
    return { outcome: 'ignored' }
  } catch (error) {
    if (error && TRANSIENT_ERRNOS.has(error.code)) return { outcome: 'retry-later', detail: technicalDetail(error) }
    return { outcome: 'failed', detail: technicalDetail(error) }
  }
}

/**
 * Replays journals left by processes that ended without finishing their saves
 * (crash, kill, power loss): deletes recorded temp files, completes a swap
 * whose verified temp is still there, moves a moved-aside original back, or
 * restores an interrupted in-place write from a backup whose hash proves it
 * complete. Call it once from app.whenReady(), after the single-instance lock
 * is held, and await it before opening documents. Only paths recorded in a
 * journal are touched, and nothing that cannot be proven redundant is deleted:
 * a file whose content is unknown is never overwritten (the original is kept
 * beside it instead).
 *
 * While this process holds Electron's single-instance lock and the journal
 * folder belongs to its userData folder, no other live process can own a
 * journal there, so every other journal is replayed even when its process id
 * now belongs to an unrelated program (Windows reuses ids quickly).
 * @param {object} [options]
 * @param {string} [options.journalDir] defaults to journalDir()
 * @param {boolean} [options.exclusive] the caller guarantees no other live process writes journals in this
 *   folder (default: detected from Electron's single-instance lock, as described above)
 * @param {number} [options.staleMs] without exclusivity, journals of live processes older than this are
 *   swept too (PID reuse)
 * @returns {Promise<{journals: number, actions: Array<{target: string, phase: string, outcome: string, detail?: string}>}>}
 */
async function sweep(options = {}) {
  const dir = options.journalDir ? path.resolve(options.journalDir) : journalDir()
  const staleMs = options.staleMs ?? JOURNAL_STALE_MS
  const own = path.join(journalDir(), `journal-${process.pid}-${processToken}.json`)
  const exclusive = typeof options.exclusive === 'boolean' ? options.exclusive : holdsJournalFolderLock(dir)
  const report = { journals: 0, actions: [] }
  let names
  try {
    names = await fsp.readdir(dir)
  } catch (error) {
    if (error.code === 'ENOENT') return report
    logger.warn(`Could not read the save journal folder ${dir}: ${error.message}`)
    return report
  }
  for (const name of names) {
    const filePath = path.join(dir, name)
    if (name.startsWith('damaged-')) {
      try {
        if (Date.now() - (await fsp.stat(filePath)).mtimeMs > DAMAGED_JOURNAL_KEEP_MS) await removeFile(filePath)
      } catch {}
      continue
    }
    const leftover = /^journal-(\d+)-([0-9a-f]+)\.json\.[0-9a-f]+\.tmp$/.exec(name)
    if (leftover) {
      const leftoverPid = Number(leftover[1])
      const foreign = leftoverPid !== process.pid || leftover[2] !== processToken
      if ((exclusive && foreign) || !processIsAlive(leftoverPid) || (leftoverPid === process.pid && leftover[2] !== processToken)) await removeFile(filePath)
      continue
    }
    const match = JOURNAL_FILE_PATTERN.exec(name)
    if (!match || path.resolve(filePath) === path.resolve(own)) continue
    const pid = Number(match[1])
    let stat
    try { stat = await fsp.stat(filePath) } catch { continue }
    const sameProcessOldToken = pid === process.pid && match[2] !== processToken
    if (!exclusive && !sameProcessOldToken && processIsAlive(pid) && Date.now() - stat.mtimeMs < staleMs) continue
    let data
    try {
      data = JSON.parse((await fsp.readFile(filePath, 'utf8')).replace(/^\uFEFF/, ''))
      if (!data || !Array.isArray(data.entries)) throw new Error('no entries')
    } catch (error) {
      const damaged = path.join(dir, `damaged-${Date.now()}-${name}`)
      await fsp.rename(filePath, damaged).catch(() => {})
      logger.warn(`Moved an unreadable save journal aside: ${damaged} (${error.message})`)
      continue
    }
    report.journals += 1
    for (const entry of data.entries) {
      if (!entry || typeof entry.target !== 'string') continue
      const result = await recoverEntry(entry)
      report.actions.push({ target: entry.target, phase: entry.phase, ...result })
      if (result.outcome === 'retry-later') {
        const { id: _id, startedAt: _startedAt, updatedAt: _updatedAt, ...rest } = entry
        await journal.add({ ...rest, carriedFrom: name })
      } else if (result.outcome !== 'removed-temp' && result.outcome !== 'cleaned') {
        logger.warn(`Finished an interrupted save of ${entry.target}: ${result.outcome}${result.detail ? ` (${result.detail})` : ''}`)
      }
    }
    await removeFile(filePath)
  }
  return report
}

module.exports = {
  FALLBACK_CODES,
  IO_CODES,
  IoError,
  RETRY_DELAYS,
  SAME_FOLDER_CODES,
  TEMP_PREFIX,
  TRANSIENT_ERRNOS,
  UNAVAILABLE_ERRNOS,
  catalog,
  catalogEntry,
  checkFileName,
  checkTargetPath,
  compareStamp,
  configureIo,
  describeDrive,
  describeIoError,
  formatBytes,
  formatCatalogPrompt,
  formatTemplate,
  freeSpace,
  hashBytes,
  hashFile,
  isIoError,
  journal,
  journalDir,
  listPendingWrites,
  nearestExistingFolder,
  pathExists,
  pendingWriteCount,
  probeFolder,
  qaWriteDelayMs,
  readFileStamped,
  removeFile,
  requiredFreeBytes,
  retryTransient,
  sleep,
  stampFile,
  stampFromBytes,
  sweep,
  technicalDetail,
  tempFileName,
  toIoError,
  toIoResult,
  trackPendingWrite,
  waitForPendingWrites,
}
