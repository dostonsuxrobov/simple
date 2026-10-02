'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const { Worker } = require('node:worker_threads')
const { safeWriteFile } = require('../shared/electron/safe-write.cjs')
const { CombineError, serializeError } = require('./combine-policy.cjs')

/** What to do next after each shared save failure, worded for a combined PDF. */
const NEXT_STEP = Object.freeze({
  LOCKED: 'Close it there and save again, or choose a different name.',
  READ_ONLY: 'Save the combined PDF with a different name or in another folder.',
  NO_PERMISSION: 'Choose another folder.',
  READ_ONLY_VOLUME: 'Choose another drive.',
  FOLDER_MISSING: 'Choose another folder.',
  FILE_UNAVAILABLE: 'Try again in a moment, or choose another folder.',
  DISK_FULL: 'Free up some space, or choose another drive.',
  NAME_TOO_LONG: 'Choose a shorter name or a folder closer to the top of the drive.',
  INVALID_NAME: 'Choose another name.',
  VERIFY_FAILED: 'Try again, or save in another folder. Your source files are unchanged.',
  VALIDATION_FAILED: 'Try again. Your source files are unchanged.',
})

/**
 * Saves the combined PDF through the shared safe-write path: a verified temp
 * file in the same folder, then rename, two-step swap or verified in-place
 * write, retried while another program (a viewer, a sync program, a virus
 * scanner) briefly holds the target. A read-only target is never overwritten,
 * and on failure the previous file is left exactly as it was.
 * @param {string} target absolute .pdf path
 * @param {Uint8Array} bytes
 * @returns {Promise<{path: string, strategy: string}>}
 * @throws {CombineError} with the shared code (LOCKED, READ_ONLY, NO_PERMISSION, DISK_FULL, …) and a plain message
 */
async function atomicWrite(target, bytes) {
  try {
    const result = await safeWriteFile(target, bytes, { format: 'pdf' })
    return { path: result.path, strategy: result.strategy }
  } catch (error) {
    if (error instanceof TypeError) throw error
    const code = typeof error?.code === 'string' && /^[A-Z][A-Z0-9_]*$/.test(error.code) ? error.code : 'UNKNOWN'
    const kept = error?.backupPath || error?.asidePath
    const next = code === 'RESTORE_NEEDED' && kept ? `The previous file is safe at "${kept}".` : (NEXT_STEP[code] || 'Your source files are unchanged.')
    throw new CombineError(code, `${error?.message || `Simple couldn't save "${path.basename(target)}".`} ${next}`, { technical: error?.technical })
  }
}

/**
 * Where MuPDF is: the copy the PDF workspace ships (modules/pdf/vendor/mupdf,
 * placed by scripts/sync-build.cjs), or, in an unpackaged checkout without a
 * build, the PDF workspace's own package.
 * @returns {string|null} absolute path of mupdf.js
 */
function mupdfPath() {
  const candidates = [
    path.join(__dirname, '..', 'modules', 'pdf', 'vendor', 'mupdf', 'dist', 'mupdf.js'),
    path.join(__dirname, '..', '..', 'simple_pdf_source', 'node_modules', 'mupdf', 'dist', 'mupdf.js'),
  ]
  return candidates.find((candidate) => fs.existsSync(candidate)) || null
}

let mupdfModule = null
function loadMupdf() {
  const file = mupdfPath()
  if (!file) return Promise.reject(new CombineError('ENCRYPTED', "Simple can't read this protected PDF here."))
  mupdfModule ??= import(pathToFileURL(file).href).catch((error) => { mupdfModule = null; throw error })
  return mupdfModule
}

/**
 * Removes encryption from a PDF that opens without a password (a
 * permissions or "owner" password only), as the PDF workspace does when it
 * opens one: MuPDF writes an unencrypted copy in memory. The file on disk is
 * never touched.
 * @param {Uint8Array} bytes the encrypted PDF
 * @returns {Promise<Uint8Array>} the same pages without encryption
 * @throws {CombineError} ENCRYPTED when the PDF needs a password to open, or MuPDF can't read it
 */
async function unlockPdfBytes(bytes) {
  const mupdf = await loadMupdf()
  let document = null
  try {
    document = mupdf.Document.openDocument(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), 'application/pdf')
  } catch (error) {
    throw new CombineError('ENCRYPTED', "Simple can't read this protected PDF. Open it in Simple, save a copy, then add the copy.", { technical: String(error?.message || error) })
  }
  try {
    if (document.needsPassword() && !document.authenticatePassword('')) {
      throw new CombineError('ENCRYPTED', 'This PDF needs a password to open. Open it in Simple with its password, save a copy, then add the copy.')
    }
    return new Uint8Array(document.saveToBuffer('encrypt=none').asUint8Array())
  } finally {
    document.destroy?.()
  }
}

/** The bundled worker (unpacked from the asar in a packaged app). */
function defaultWorkerPath() {
  return path.join(__dirname, '..', 'modules', 'shared', 'combine-worker.cjs').replace(/app\.asar([\\/])/, 'app.asar.unpacked$1')
}

function workerError(value) {
  if (typeof value === 'string') return new CombineError('UNKNOWN', value)
  const { message, code } = serializeError(value)
  return new CombineError(code, message)
}

/**
 * Combines files in a worker thread. The worker asks for each of Simple's own
 * page layouts to be printed, and for each PDF protected by a permissions
 * password to be unlocked; those requests are answered with
 * `options.printHtml` and `options.unlockPdf` in this (main) process.
 * @param {Array<{path: string, pages?: string}>} entries
 * @param {(progress: {index: number, total: number, name: string}) => void} onProgress
 * @param {object} [options]
 * @param {(html: string, options: object) => Promise<Uint8Array>} [options.printHtml] the hardened
 *   html-to-pdf printer; without it, files that need Simple's own layout fail with a clear message
 * @param {(bytes: Uint8Array) => Promise<Uint8Array>} [options.unlockPdf] removes a permissions-only
 *   encryption (default: unlockPdfBytes); a PDF that needs a password is refused either way
 * @param {string} [options.title] the combined PDF's document title (default "Combined document")
 * @param {string} [options.workerPath] worker script (default: the bundled modules/shared/combine-worker.cjs)
 * @param {number} [options.timeoutMs=600000]
 * @returns {Promise<{bytes: Uint8Array, pageCount: number, builtIn?: string[]}>}
 * @throws {CombineError} with the worker's code and message
 */
function runCombine(entries, onProgress, options = {}) {
  const worker = new Worker(options.workerPath || defaultWorkerPath(), { workerData: { entries, title: options.title } })
  const unlockPdf = typeof options.unlockPdf === 'function' ? options.unlockPdf : unlockPdfBytes
  return new Promise((resolve, reject) => {
    let settled = false
    const finish = (error, result) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      void worker.terminate()
      if (error) reject(error)
      else resolve(result)
    }
    const timer = setTimeout(() => finish(new CombineError('TIMEOUT', 'Combining took too long. Try fewer files. Your source files are unchanged.')), options.timeoutMs || 10 * 60_000)
    const print = async ({ id, html, options: printOptions }) => {
      try {
        if (typeof options.printHtml !== 'function') throw new CombineError('PRINT_UNAVAILABLE', "Simple can't make PDF pages from this file here.")
        const bytes = await options.printHtml(html, printOptions)
        if (!settled) worker.postMessage({ printed: { id, bytes } })
      } catch (error) {
        if (!settled) worker.postMessage({ printed: { id, error: serializeError(error) } })
      }
    }
    const unlock = async ({ id, bytes }) => {
      try {
        const unlocked = await unlockPdf(bytes)
        if (!settled) worker.postMessage({ unlocked: { id, bytes: unlocked } })
      } catch (error) {
        if (!settled) worker.postMessage({ unlocked: { id, error: serializeError(error) } })
      }
    }
    worker.on('message', (message) => {
      if (message.progress) onProgress?.(message.progress)
      else if (message.print) void print(message.print)
      else if (message.unlock) void unlock(message.unlock)
      else if (message.error) finish(workerError(message.error))
      else if (message.result) finish(null, message.result)
    })
    worker.once('error', (error) => finish(error))
    worker.once('exit', (code) => { if (!settled) finish(new CombineError('UNKNOWN', `Combining stopped unexpectedly (${code}). Try again.`)) })
  })
}

module.exports = { atomicWrite, defaultWorkerPath, mupdfPath, runCombine, unlockPdfBytes }
