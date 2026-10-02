'use strict'

// Checks the launcher's main process runs on the files picked for Combine,
// before anything is read: each must be a local file Combine accepts, and the
// combined PDF must never replace one of them. A file that can't be reached
// is reported in plain words with its name (moved, locked, or on a drive that
// isn't available), never with Node's technical text. Node built-ins only;
// the office engine probe is injectable for tests.

const fsp = require('node:fs/promises')
const path = require('node:path')
const { CombineError, isCombinePath, isLegacyOfficePath, legacyFormatError, sourceFileError } = require('./combine-policy.cjs')

const MAX_FILES = 100
const MAX_FILE_BYTES = 256 * 1024 * 1024

/**
 * Checks files for Combine. Older .doc and .xls files are set aside with a
 * clear reason when this PC has no local office engine, so the rest can still
 * be added.
 * @param {string[]} paths
 * @param {object} [options]
 * @param {() => Promise<{available: boolean}>} [options.officeEngineStatus] default: the shared engine probe
 * @param {{stat: typeof fsp.stat}} [options.fs] replaces node:fs/promises (tests)
 * @returns {Promise<{entries: Array<{path: string, name: string, size: number, pages: string}>, skipped: Array<{path: string, name: string, code: string, message: string}>}>}
 * @throws {CombineError} INVALID, UNSUPPORTED, TOO_LARGE, or a file error (NOT_FOUND, LOCKED, FILE_UNAVAILABLE, UNREADABLE)
 */
async function describeCombinePaths(paths, options = {}) {
  if (!Array.isArray(paths) || paths.length > MAX_FILES) throw new CombineError('INVALID', 'Choose up to 100 files at a time.')
  const files = options.fs || fsp
  const engineStatus = options.officeEngineStatus || (() => require('../shared/electron/office-engine.cjs').getOfficeEngineStatus())
  const engine = paths.some(isLegacyOfficePath) ? await engineStatus() : null
  const entries = []
  const skipped = []
  for (const filePath of paths) {
    if (typeof filePath !== 'string' || !path.isAbsolute(filePath) || !isCombinePath(filePath)) {
      throw new CombineError('UNSUPPORTED', 'Combine accepts PDFs, Word documents, Excel, ODS and CSV spreadsheets, PNGs, and JPEGs.')
    }
    let stat
    try {
      stat = await files.stat(filePath)
    } catch (error) {
      throw sourceFileError(filePath, error)
    }
    if (!stat.isFile() || !stat.size || stat.size > MAX_FILE_BYTES) throw new CombineError('TOO_LARGE', `${path.basename(filePath)} must be a nonempty file smaller than 256 MB.`)
    if (engine && !engine.available && isLegacyOfficePath(filePath)) {
      const reason = legacyFormatError(filePath)
      skipped.push({ path: filePath, name: path.basename(filePath), code: reason.code, message: reason.message })
      continue
    }
    entries.push({ path: filePath, name: path.basename(filePath), size: stat.size, pages: '' })
  }
  return { entries, skipped }
}

/**
 * Refuses a save target that is one of the source files (also through a
 * link or another spelling of the same path).
 * @param {Array<{path: string}>} entries
 * @param {string} target the chosen .pdf path
 * @param {{fs?: {realpath: typeof fsp.realpath}}} [options] fs replaces node:fs/promises (tests)
 * @throws {CombineError} SOURCE_TARGET, or a file error for a source that can no longer be reached
 */
async function assertTargetIsNotASource(entries, target, options = {}) {
  const files = options.fs || fsp
  const realTarget = await files.realpath(target).catch(() => path.resolve(target))
  for (const entry of entries) {
    let realSource
    try {
      realSource = await files.realpath(entry.path)
    } catch (error) {
      throw sourceFileError(entry.path, error)
    }
    if (realSource.toLowerCase() === realTarget.toLowerCase()) throw new CombineError('SOURCE_TARGET', 'Choose a new filename so the combined PDF keeps your source file intact.')
  }
}

module.exports = { assertTargetIsNotASource, describeCombinePaths }
