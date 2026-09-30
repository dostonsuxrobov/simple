const path = require('node:path')
const legacyDoc = require('./legacy-doc.cjs')

const SUPPORTED_EXTENSIONS = new Set(['.docx', '.doc'])

function isSupportedDocumentPath(filePath) {
  return typeof filePath === 'string' && SUPPORTED_EXTENSIONS.has(path.extname(filePath).toLowerCase())
}

module.exports = {
  SUPPORTED_EXTENSIONS,
  isSupportedDocumentPath,
  ...legacyDoc,
}
