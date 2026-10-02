'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const { codedError } = require('./pdf-problems.cjs')

// pdf-lib cannot edit encrypted files, so unlock a copy with mupdf: authenticate,
// then rewrite without encryption. The original file is never touched.
async function loadMupdf() {
  const bundled = path.join(__dirname, '../vendor/mupdf/dist/mupdf.js')
  return import(pathToFileURL(fs.existsSync(bundled) ? bundled : require.resolve('mupdf')).href)
}

const ENCRYPT = Buffer.from('/Encrypt')
const PDF_WHITESPACE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20])
const PDF_DELIMITERS = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25])

function asBuffer(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  return Buffer.from(value || [])
}

function isDigit(byte) {
  return byte >= 0x30 && byte <= 0x39
}

/**
 * True when a trailer (or cross-reference stream dictionary) names an
 * encryption dictionary: `/Encrypt 12 0 R` or `/Encrypt << … >>`.
 *
 * The whole file is searched. Linearized ("Fast Web View") files keep the
 * full trailer in the first-page section near the start of the file while the
 * final trailer holds only /Size, so a scan of the file's tail misses them.
 * Trailers and xref stream dictionaries are never compressed, so the entry is
 * always visible in the raw bytes.
 */
function hasEncryptionDictionary(value) {
  const bytes = asBuffer(value)
  let index = bytes.indexOf(ENCRYPT)
  while (index >= 0) {
    let cursor = index + ENCRYPT.length
    // `/EncryptMetadata` and other longer names are different keys.
    if (cursor >= bytes.length || PDF_WHITESPACE.has(bytes[cursor]) || PDF_DELIMITERS.has(bytes[cursor])) {
      while (cursor < bytes.length && PDF_WHITESPACE.has(bytes[cursor])) cursor += 1
      if (bytes[cursor] === 0x3c && bytes[cursor + 1] === 0x3c) return true
      let digits = 0
      while (isDigit(bytes[cursor])) { cursor += 1; digits += 1 }
      if (digits) {
        while (PDF_WHITESPACE.has(bytes[cursor])) cursor += 1
        let generation = 0
        while (isDigit(bytes[cursor])) { cursor += 1; generation += 1 }
        while (PDF_WHITESPACE.has(bytes[cursor])) cursor += 1
        if (generation && bytes[cursor] === 0x52) return true
      }
    }
    index = bytes.indexOf(ENCRYPT, index + ENCRYPT.length)
  }
  return false
}

/**
 * @returns {Promise<{ status: 'none' | 'needs-password' | 'wrong-password' | 'unlocked', data?: Uint8Array }>}
 */
async function unlockPdf(data, password = '') {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
  if (!hasEncryptionDictionary(bytes)) return { status: 'none' }
  const mupdf = await loadMupdf()
  const doc = mupdf.Document.openDocument(bytes, 'application/pdf')
  try {
    // Owner-password-only files open without a prompt; they are still rewritten unencrypted.
    if (doc.needsPassword() && !doc.authenticatePassword(String(password || ''))) {
      return { status: password ? 'wrong-password' : 'needs-password' }
    }
    const out = doc.saveToBuffer('encrypt=none').asUint8Array()
    return { status: 'unlocked', data: new Uint8Array(out) }
  } finally {
    doc.destroy?.()
  }
}

/** 'none', 'owner' (permissions password only) or 'password' (needs a password to open). */
async function encryptionKind(data) {
  const bytes = asBuffer(data)
  if (!hasEncryptionDictionary(bytes)) return 'none'
  const mupdf = await loadMupdf()
  const doc = mupdf.Document.openDocument(bytes, 'application/pdf')
  try {
    return doc.needsPassword() ? 'password' : 'owner'
  } finally {
    doc.destroy?.()
  }
}

function isEncryptedPdfError(error) {
  return error?.name === 'EncryptedPDFError' || /PDFDocument\.load` is encrypted|is encrypted/i.test(String(error?.message || ''))
}

/**
 * Replace pdf-lib's "Input document to `PDFDocument.load` is encrypted…" with
 * PASSWORD_REQUIRED or OWNER_LOCKED and a message a user can act on.
 */
async function encryptedDocumentError(data, name = '') {
  let kind = 'password'
  try { kind = await encryptionKind(data) } catch { /* Report the stricter case. */ }
  const subject = name ? `“${name}”` : 'This PDF'
  if (kind === 'owner') {
    return codedError('OWNER_LOCKED', `${subject} has editing restrictions set by its author. Open it again so Simple can make an editable copy; the protected original is not changed.`)
  }
  return codedError('PASSWORD_REQUIRED', `${subject} is password-protected. Open it and enter its password before editing, printing or adding its pages.`)
}

/**
 * Bytes pdf-lib can read: unencrypted input as is, owner-password-only input
 * decrypted to a copy. Files that need a password raise PASSWORD_REQUIRED.
 * Use only where the result can never overwrite the protected original
 * (print output, pages copied into another document).
 */
async function readableCopy(data, name = '') {
  if (!hasEncryptionDictionary(data)) return data
  const result = await unlockPdf(data, '')
  if (result.status === 'unlocked' && result.data) return result.data
  if (result.status === 'none') return data
  throw await encryptedDocumentError(data, name)
}

module.exports = {
  encryptedDocumentError,
  encryptionKind,
  hasEncryptionDictionary,
  isEncryptedPdfError,
  readableCopy,
  unlockPdf,
}
