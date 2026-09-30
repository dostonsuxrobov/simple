'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { pathToFileURL } = require('node:url')

// pdf-lib cannot edit encrypted files, so unlock a copy with mupdf: authenticate,
// then rewrite without encryption. The original file is never touched.
async function loadMupdf() {
  const bundled = path.join(__dirname, '../vendor/mupdf/dist/mupdf.js')
  return import(pathToFileURL(fs.existsSync(bundled) ? bundled : require.resolve('mupdf')).href)
}

function mightBeEncrypted(bytes) {
  const tail = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).subarray(Math.max(0, bytes.byteLength - 262_144))
  return tail.includes('/Encrypt')
}

/**
 * @returns {Promise<{ status: 'none' | 'needs-password' | 'wrong-password' | 'unlocked', data?: Uint8Array }>}
 */
async function unlockPdf(data, password = '') {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
  if (!mightBeEncrypted(bytes)) return { status: 'none' }
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

module.exports = { unlockPdf }
