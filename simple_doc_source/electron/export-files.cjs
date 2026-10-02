'use strict'

const ODT_MIME_TYPE = 'application/vnd.oasis.opendocument.text'

const EXPORT_FORMATS = Object.freeze({
  docx: Object.freeze({ extension: '.docx', label: 'Word document', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
  pdf: Object.freeze({ extension: '.pdf', label: 'PDF document', mimeType: 'application/pdf' }),
  html: Object.freeze({ extension: '.html', label: 'Web page', mimeType: 'text/html;charset=utf-8' }),
  md: Object.freeze({ extension: '.md', label: 'Markdown document', mimeType: 'text/markdown;charset=utf-8' }),
  txt: Object.freeze({ extension: '.txt', label: 'Plain text', mimeType: 'text/plain;charset=utf-8' }),
  odt: Object.freeze({ extension: '.odt', label: 'OpenDocument text', mimeType: ODT_MIME_TYPE }),
  rtf: Object.freeze({ extension: '.rtf', label: 'Rich Text document', mimeType: 'application/rtf' }),
  // Written from the DOCX bytes by a local Office engine, only when one is installed.
  doc: Object.freeze({ extension: '.doc', label: 'Word 97–2003 document', mimeType: 'application/msword', requiresOfficeEngine: true }),
})

const MAX_EXPORT_BYTES = 512 * 1024 * 1024

function exportFormat(format) {
  const normalized = String(format || '').toLowerCase()
  const definition = Object.hasOwn(EXPORT_FORMATS, normalized) ? EXPORT_FORMATS[normalized] : null
  if (!definition) throw new Error('That export format is not supported.')
  return { id: normalized, ...definition }
}

/** The formats Export As can offer right now. */
function availableExportFormats({ officeEngine = false } = {}) {
  return Object.entries(EXPORT_FORMATS)
    .filter(([, definition]) => !definition.requiresOfficeEngine || officeEngine)
    .map(([id, definition]) => ({ id, label: definition.label, extension: definition.extension }))
}

function validateExportBytes(value) {
  let bytes
  if (Buffer.isBuffer(value)) bytes = value
  else if (ArrayBuffer.isView(value)) bytes = Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  else if (value instanceof ArrayBuffer) bytes = Buffer.from(value)
  else if (value && value.type === 'Buffer' && Array.isArray(value.data)) bytes = Buffer.from(value.data)
  else throw new Error('The exported file data is invalid.')
  if (bytes.byteLength === 0) throw new Error('The exported file is empty.')
  if (bytes.byteLength > MAX_EXPORT_BYTES) throw new Error('The exported file is too large to save safely.')
  return bytes
}

/** The first stored entry of a ZIP package, used for the ODF mimetype check. */
function firstZipEntry(bytes) {
  if (bytes.length < 30 || bytes.readUInt32LE(0) !== 0x04034b50) return null
  const method = bytes.readUInt16LE(8)
  const compressedSize = bytes.readUInt32LE(18)
  const nameLength = bytes.readUInt16LE(26)
  const extraLength = bytes.readUInt16LE(28)
  const name = bytes.subarray(30, 30 + nameLength).toString('utf8')
  const start = 30 + nameLength + extraLength
  return { name, method, data: bytes.subarray(start, Math.min(bytes.length, start + compressedSize)) }
}

/**
 * Check that bytes really are the chosen format before they replace a file:
 * a broken or mislabeled export never reaches the user's folder.
 */
function validateExportContent(format, value) {
  const bytes = validateExportBytes(value)
  const id = exportFormat(format).id
  if (id === 'pdf' && bytes.subarray(0, 5).toString('latin1') !== '%PDF-') throw new Error('The exported PDF is not valid.')
  if (id === 'rtf') {
    const text = bytes.toString('latin1')
    if (!/^\s*\{\\rtf1/.test(text) || !/\}\s*$/.test(text)) throw new Error('The exported Rich Text document is not valid.')
    let depth = 0
    for (let index = 0; index < text.length; index += 1) {
      const character = text[index]
      if (character === '\\') { index += 1; continue }
      if (character === '{') depth += 1
      else if (character === '}' && --depth < 0) break
    }
    if (depth !== 0) throw new Error('The exported Rich Text document is incomplete.')
  }
  if (id === 'odt') {
    const first = firstZipEntry(bytes)
    if (!first || first.name !== 'mimetype' || first.method !== 0 || first.data.toString('ascii') !== ODT_MIME_TYPE) {
      throw new Error('The exported OpenDocument file is not valid.')
    }
  }
  if (id === 'docx' || id === 'doc') {
    // Both arrive as a DOCX package; .doc is converted afterwards.
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('The exported Word document is not valid.')
  }
  return bytes
}

module.exports = { EXPORT_FORMATS, MAX_EXPORT_BYTES, ODT_MIME_TYPE, availableExportFormats, exportFormat, validateExportBytes, validateExportContent }
