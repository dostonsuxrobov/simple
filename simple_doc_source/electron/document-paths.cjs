'use strict'

const path = require('node:path')

// Only real document extensions are removed from a name. A plain title such as
// "J. Smith CV" or "Budget v1.2" keeps every character after its dots.
const KNOWN_DOCUMENT_EXTENSION = /\.(?:docx|docm|dotx|dotm|doc|dot|pdf|html?|mht|mhtml|md|markdown|txt|rtf|odt)$/i
const UNSAFE_NAME_CHARACTERS = /[<>:"/\\|?*\u0000-\u001f]/g
const RESERVED_WINDOWS_NAME = /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])$/i
const FALLBACK_STEM = 'Untitled document'

function baseName(value) {
  return String(value ?? '').split(/[\\/]/).pop() || ''
}

function safeStem(value) {
  const stem = baseName(value)
    .replace(KNOWN_DOCUMENT_EXTENSION, '')
    .replace(UNSAFE_NAME_CHARACTERS, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    // Windows silently drops trailing dots and spaces from file names.
    .replace(/[. ]+$/, '')
  if (!stem || RESERVED_WINDOWS_NAME.test(stem)) return stem ? `${stem} document` : FALLBACK_STEM
  return Array.from(stem).slice(0, 180).join('').trim() || FALLBACK_STEM
}

function isAbsoluteDocumentPath(value) {
  return typeof value === 'string' && value.length > 0 && value.length < 32_768 && path.isAbsolute(value) && !value.includes('\u0000')
}

/**
 * The default for a Save as or Export dialog: the document's own folder when it
 * is known (like Word), otherwise a bare file name so Windows chooses.
 */
function dialogDefaultPath({ name, extension, folderOf } = {}) {
  const fileName = `${safeStem(name)}${extension || ''}`
  const reference = [folderOf].flat().find(isAbsoluteDocumentPath)
  return reference ? path.join(path.dirname(path.resolve(reference)), fileName) : fileName
}

/** Candidate names beside a source document: name.docx, then "name (edited).docx". */
function* siblingCandidates(sourcePath, extension = '.docx') {
  const directory = path.dirname(path.resolve(sourcePath))
  const stem = safeStem(path.basename(sourcePath))
  yield path.join(directory, `${stem}${extension}`)
  yield path.join(directory, `${stem} (edited)${extension}`)
  for (let index = 2; index < 1000; index += 1) yield path.join(directory, `${stem} (edited ${index})${extension}`)
}

function sameFilePath(left, right) {
  if (!isAbsoluteDocumentPath(left) || !isAbsoluteDocumentPath(right)) return false
  const normalize = (value) => path.resolve(value).replace(/[\\/]+$/, '')
  return process.platform === 'win32' || /^[a-z]:[\\/]/i.test(left)
    ? normalize(left).toLowerCase() === normalize(right).toLowerCase()
    : normalize(left) === normalize(right)
}

module.exports = {
  KNOWN_DOCUMENT_EXTENSION,
  dialogDefaultPath,
  isAbsoluteDocumentPath,
  safeStem,
  sameFilePath,
  siblingCandidates,
}
