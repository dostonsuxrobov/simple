'use strict'

const path = require('node:path')

// Ownership is intentionally exclusive. Some upstream apps can technically
// import the same format; the unified app chooses the editor that best matches
// the extension so a double-click always has one deterministic destination.
const EXTENSIONS_BY_MODE = Object.freeze({
  docs: Object.freeze(['.docx']),
  pdf: Object.freeze(['.pdf', '.txt', '.md', '.doc']),
  image: Object.freeze(['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg', '.avif']),
  video: Object.freeze(['.mp4', '.m4v', '.webm', '.ogv', '.mov', '.mkv']),
  calc: Object.freeze([
    '.xlsx', '.xlsm', '.xlsb', '.xls', '.xltx', '.xltm', '.xlt', '.xlam', '.xla',
    '.xml', '.ods', '.fods', '.csv', '.tsv', '.tab', '.numbers', '.slk', '.sylk',
    '.dif', '.dbf', '.prn', '.wk1', '.wk2', '.wk3', '.wk4', '.wks', '.wq1',
    '.wq2', '.wb1', '.wb2', '.wb3', '.123', '.qpw', '.html', '.htm',
  ]),
})

const MODE_BY_EXTENSION = new Map()
for (const [mode, extensions] of Object.entries(EXTENSIONS_BY_MODE)) {
  for (const extension of extensions) {
    if (MODE_BY_EXTENSION.has(extension)) throw new Error(`Duplicate extension owner: ${extension}`)
    MODE_BY_EXTENSION.set(extension, mode)
  }
}

const SUPPORTED_EXTENSIONS = Object.freeze([...MODE_BY_EXTENSION.keys()])
const MODES = Object.freeze(Object.keys(EXTENSIONS_BY_MODE))

function normalizedExtension(filePath) {
  return path.extname(String(filePath || '')).toLowerCase()
}

function modeForPath(filePath) {
  return MODE_BY_EXTENSION.get(normalizedExtension(filePath)) || null
}

function supportedPaths(argv) {
  const seen = new Set()
  const paths = []
  for (const argument of Array.isArray(argv) ? argv : []) {
    if (typeof argument !== 'string' || argument.startsWith('-') || !modeForPath(argument)) continue
    const key = process.platform === 'win32' ? argument.toLowerCase() : argument
    if (seen.has(key)) continue
    seen.add(key)
    paths.push(argument)
  }
  return paths
}

function groupPathsByMode(paths) {
  const groups = new Map()
  for (const filePath of paths) {
    const mode = modeForPath(filePath)
    if (!mode) continue
    if (!groups.has(mode)) groups.set(mode, [])
    groups.get(mode).push(filePath)
  }
  return groups
}

function explicitMode(argv) {
  for (const argument of Array.isArray(argv) ? argv : []) {
    const match = /^--simple-mode=([a-z]+)$/i.exec(String(argument))
    const requested = match?.[1].toLowerCase()
    if (requested && MODES.includes(requested)) return requested
  }
  return null
}

module.exports = {
  EXTENSIONS_BY_MODE,
  MODES,
  MODE_BY_EXTENSION,
  SUPPORTED_EXTENSIONS,
  explicitMode,
  groupPathsByMode,
  modeForPath,
  normalizedExtension,
  supportedPaths,
}
