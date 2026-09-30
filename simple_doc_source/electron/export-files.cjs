'use strict'

const EXPORT_FORMATS = Object.freeze({
  docx: Object.freeze({ extension: '.docx', label: 'Word document', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
  pdf: Object.freeze({ extension: '.pdf', label: 'PDF document', mimeType: 'application/pdf' }),
  html: Object.freeze({ extension: '.html', label: 'Web page', mimeType: 'text/html;charset=utf-8' }),
  md: Object.freeze({ extension: '.md', label: 'Markdown document', mimeType: 'text/markdown;charset=utf-8' }),
  txt: Object.freeze({ extension: '.txt', label: 'Plain text', mimeType: 'text/plain;charset=utf-8' }),
})

const MAX_EXPORT_BYTES = 512 * 1024 * 1024

function exportFormat(format) {
  const normalized = String(format || '').toLowerCase()
  const definition = EXPORT_FORMATS[normalized]
  if (!definition) throw new Error('That export format is not supported.')
  return { id: normalized, ...definition }
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

module.exports = { EXPORT_FORMATS, MAX_EXPORT_BYTES, exportFormat, validateExportBytes }
