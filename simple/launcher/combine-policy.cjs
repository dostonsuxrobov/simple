'use strict'

// Which files Combine accepts, and how a file that only a local office engine
// could convert is reported. Node built-ins only: the launcher main process
// loads this unbundled, and the Combine worker bundle includes it too.

const path = require('node:path')

/** Extensions Combine accepts, without the dot, in file-picker order. */
const COMBINE_EXTENSIONS = Object.freeze(['pdf', 'docx', 'doc', 'xlsx', 'xls', 'ods', 'csv', 'png', 'jpg', 'jpeg'])

/** Older binary Office formats that only a local office engine can turn into PDF pages. */
const LEGACY_FORMATS = Object.freeze({
  '.doc': Object.freeze({ label: 'older Word documents (.doc)', altExt: '.docx', workspace: 'docs' }),
  '.xls': Object.freeze({ label: 'older Excel workbooks (.xls)', altExt: '.xlsx', workspace: 'calc' }),
})

/** Error with a stable `code` that survives the worker boundary and IPC as plain fields. */
class CombineError extends Error {
  /**
   * @param {string} code e.g. NEEDS_OFFICE_ENGINE, UNSUPPORTED, TOO_LARGE, LOCKED
   * @param {string} message plain, user-facing text
   * @param {object} [details] extra plain fields (altExt, workspace, path, technical)
   */
  constructor(code, message, details = {}) {
    super(message)
    this.name = 'CombineError'
    this.code = code
    for (const [key, value] of Object.entries(details)) {
      if (value !== undefined && key !== 'message' && key !== 'name' && key !== 'code') this[key] = value
    }
  }
}

/**
 * Lower-case extension with the dot, e.g. ".docx".
 * @param {string} filePath
 * @returns {string}
 */
function combineExtension(filePath) {
  return path.extname(String(filePath || '')).toLowerCase()
}

/**
 * True when Combine accepts this file name.
 * @param {string} filePath
 * @returns {boolean}
 */
function isCombinePath(filePath) {
  return COMBINE_EXTENSIONS.includes(combineExtension(filePath).slice(1))
}

/**
 * True for .doc and .xls, which need a local office engine to become PDF pages.
 * @param {string} filePath
 * @returns {boolean}
 */
function isLegacyOfficePath(filePath) {
  return Object.prototype.hasOwnProperty.call(LEGACY_FORMATS, combineExtension(filePath))
}

/**
 * The coded error for an older Office file when no local engine is available.
 * It names what Simple can do natively instead: open the file in Simple and
 * save it in the modern format, then combine that copy.
 * @param {string} filePath the file (or name) that cannot be combined
 * @returns {CombineError} code NEEDS_OFFICE_ENGINE
 */
function legacyFormatError(filePath) {
  const extension = combineExtension(filePath)
  const format = LEGACY_FORMATS[extension] || LEGACY_FORMATS['.doc']
  return new CombineError(
    'NEEDS_OFFICE_ENGINE',
    `Simple can't combine ${format.label} on this PC. Open it in Simple and save it as ${format.altExt} first, then add the ${format.altExt} file.`,
    { altExt: format.altExt, workspace: format.workspace, path: typeof filePath === 'string' ? filePath : undefined },
  )
}

/**
 * Plain wording for a file-system error while reading a source file.
 * @param {NodeJS.ErrnoException} error
 * @returns {CombineError|null} null for an error that is not about the file
 */
function readError(error) {
  if (!error || typeof error.code !== 'string') return null
  const technical = `${error.code} ${error.syscall || ''}`.trim()
  if (error.code === 'ENOENT' || error.code === 'ENOTDIR') return new CombineError('NOT_FOUND', 'This file was moved, renamed or deleted. Add it again from its new place.', { technical })
  if (['EPERM', 'EACCES', 'EBUSY'].includes(error.code)) return new CombineError('LOCKED', "Simple couldn't read this file. Another program may be using it, or you may not have permission to open it.", { technical })
  if (['EIO', 'ETIMEDOUT', 'ECANCELED', 'UNKNOWN'].includes(error.code)) return new CombineError('FILE_UNAVAILABLE', "This file isn't available right now. Make sure its drive or folder can be opened, then try again.", { technical })
  return null
}

/**
 * The error for one source file that could not be read: its name, then the
 * plain reason (never Node's technical text, which stays in `technical`).
 * @param {string} filePath
 * @param {NodeJS.ErrnoException} error
 * @returns {CombineError}
 */
function sourceFileError(filePath, error) {
  const mapped = readError(error) || new CombineError('UNREADABLE', "Simple couldn't read this file. Try again, or add a copy of it.", {
    technical: error && error.code ? `${error.code} ${error.syscall || ''}`.trim() : undefined,
  })
  return new CombineError(mapped.code, `${path.basename(String(filePath || ''))}: ${mapped.message}`, { technical: mapped.technical })
}

/**
 * Plain fields of an error, safe to post to a worker, a window or over IPC.
 * @param {unknown} error
 * @returns {{message: string, code: string}}
 */
function serializeError(error) {
  const message = error && typeof error.message === 'string' && error.message ? error.message : 'Simple could not combine these files. Your files are unchanged.'
  const code = error && typeof error.code === 'string' && /^[A-Z][A-Z0-9_]*$/.test(error.code) ? error.code : 'UNKNOWN'
  return { message, code }
}

module.exports = {
  COMBINE_EXTENSIONS,
  CombineError,
  LEGACY_FORMATS,
  combineExtension,
  isCombinePath,
  isLegacyOfficePath,
  legacyFormatError,
  readError,
  serializeError,
  sourceFileError,
}
