'use strict'

const fs = require('node:fs')
const path = require('node:path')
const formats = require('../shared/electron/formats.cjs')
const { OPEN_LIST_SWITCH } = require('./open-list.cjs')

// Every table here is derived from the format registry
// (simple/shared/electron/formats.json); edit formats there, never here.
//
// Ownership is intentionally exclusive. Some upstream apps can technically
// import the same format; the unified app chooses the editor that best matches
// the extension so a double-click always has one deterministic destination.
// Content routing (.txt, .html/.htm, .xml) only chooses between owners the
// registry declares, and only when that workspace can open the format now:
// for example a consistently delimited .txt opens in Spreadsheets, any other
// .txt keeps its PDF route.
const EXTENSIONS_BY_MODE = formats.extensionsByMode

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

/**
 * The route a path takes without reading it: null when its extension is
 * not routed, or when its content could choose between owners.
 * @param {string} filePath
 * @returns {{mode: string|null, formatId: string|null, extension: string, by: 'extension'|null}|null}
 */
function routeWithoutReading(filePath) {
  const extension = normalizedExtension(filePath)
  if (!MODE_BY_EXTENSION.has(extension)) return { mode: null, formatId: null, extension, by: null }
  // Content only chooses between owners that open the format now; with one
  // possible owner (.html and .xml today) the answer needs no read.
  if (candidateModesForPath(filePath).length > 1) return null
  return { mode: MODE_BY_EXTENSION.get(extension), formatId: formats.formatForExtension(extension)?.id || null, extension, by: 'extension' }
}

/**
 * Full routing decision for one path (see formats.routeForPath). Reads the
 * file only when its content can change the workspace (a .txt today).
 * @param {string} filePath
 * @param {{cwd?: string}} [options] cwd resolves relative second-instance paths.
 * @returns {{mode: string|null, formatId: string|null, extension: string, by: 'extension'|'content'|null, kind?: string}}
 */
function routeForPath(filePath, options = {}) {
  return routeWithoutReading(filePath) || formats.routeForPath(filePath, { cwd: options.cwd })
}

/**
 * routeForPath() without blocking the calling process (the launcher's main
 * thread): a content sample is read asynchronously.
 * @param {string} filePath
 * @param {{cwd?: string}} [options]
 * @returns {Promise<ReturnType<typeof routeForPath>>}
 */
async function routeForPathAsync(filePath, options = {}) {
  return routeWithoutReading(filePath) || formats.routeForPathAsync(filePath, { cwd: options.cwd })
}

/**
 * The workspace that opens a path, or null when Simple does not route it.
 * @param {string} filePath
 * @param {{cwd?: string}} [options]
 * @returns {string|null}
 */
function modeForPath(filePath, options) {
  return routeForPath(filePath, options).mode
}

/**
 * The arguments that are paths Simple routes, by extension only (no file is
 * read), without repeats and in order.
 * @param {string[]} argv
 * @returns {string[]}
 */
function supportedPaths(argv) {
  const seen = new Set()
  const paths = []
  for (const argument of Array.isArray(argv) ? argv : []) {
    if (typeof argument !== 'string' || argument.startsWith('-') || !MODE_BY_EXTENSION.has(normalizedExtension(argument))) continue
    const key = process.platform === 'win32' ? argument.toLowerCase() : argument
    if (seen.has(key)) continue
    seen.add(key)
    paths.push(argument)
  }
  return paths
}

/**
 * Groups paths by the workspace that opens them, routing each path once.
 * @param {string[]} paths
 * @param {{cwd?: string}} [options]
 * @returns {Map<string, string[]>}
 */
function groupPathsByMode(paths, options = {}) {
  const groups = new Map()
  for (const filePath of paths) {
    const mode = modeForPath(filePath, options)
    if (!mode) continue
    if (!groups.has(mode)) groups.set(mode, [])
    groups.get(mode).push(filePath)
  }
  return groups
}

/** How many content samples groupPathsByModeAsync reads at a time. */
const ASYNC_ROUTE_BATCH = 8

/**
 * groupPathsByMode() without blocking the calling process: content samples
 * are read asynchronously, a few at a time. Group order follows the paths.
 * @param {string[]} paths
 * @param {{cwd?: string}} [options]
 * @returns {Promise<Map<string, string[]>>}
 */
async function groupPathsByModeAsync(paths, options = {}) {
  const modes = []
  for (let start = 0; start < paths.length; start += ASYNC_ROUTE_BATCH) {
    const batch = paths.slice(start, start + ASYNC_ROUTE_BATCH)
    modes.push(...(await Promise.all(batch.map(async (filePath) => (await routeForPathAsync(filePath, options)).mode))))
  }
  const groups = new Map()
  paths.forEach((filePath, index) => {
    const mode = modes[index]
    if (!mode) return
    if (!groups.has(mode)) groups.set(mode, [])
    groups.get(mode).push(filePath)
  })
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

/**
 * Every workspace a path's extension can be routed to (its extension route
 * plus the owners its content can choose). Reads nothing from disk.
 * @param {string} filePath
 * @returns {string[]}
 */
function candidateModesForPath(filePath) {
  return formats.candidateModesForExtension(normalizedExtension(filePath))
}

function pathKey(filePath, cwd) {
  const resolved = path.resolve(cwd || process.cwd(), String(filePath))
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved
}

/**
 * Routes a command line once. A mode forwarded with --simple-mode (by the
 * instance that already routed these paths) wins for every path that mode
 * can open, and those paths are never read; other paths are routed by
 * extension, and by content when their content can choose between owners
 * (see formats.candidateModesForExtension), relative to `cwd`.
 * @param {string[]} argv
 * @param {{cwd?: string}} [options]
 * @returns {{mode: string, groups: Map<string, string[]>, forwarded: string|null}}
 *   mode: the workspace this process runs ('launcher' without paths or a mode)
 */
function routeCommandLine(argv, options = {}) {
  const forwarded = explicitMode(argv)
  const groups = new Map()
  for (const filePath of supportedPaths(argv)) {
    const owner = forwarded && candidateModesForPath(filePath).includes(forwarded)
      ? forwarded
      : modeForPath(filePath, { cwd: options.cwd })
    if (!owner) continue
    if (!groups.has(owner)) groups.set(owner, [])
    groups.get(owner).push(filePath)
  }
  const first = groups.keys().next().value || null
  const mode = forwarded && groups.has(forwarded) ? forwarded : (first || forwarded || 'launcher')
  return { mode, groups, forwarded }
}

/**
 * Removes from a command line every routed path that is not one of
 * `ownPaths`, so a workspace never sees a file routed to another one.
 * Flags and other arguments are kept.
 * @param {string[]} argv
 * @param {string[]} ownPaths
 * @returns {string[]} a new array
 */
function keepOwnPaths(argv, ownPaths) {
  const own = new Set(ownPaths.map((filePath) => pathKey(filePath)))
  return argv.filter((argument) => {
    if (typeof argument !== 'string' || argument.startsWith('-') || !MODE_BY_EXTENSION.has(normalizedExtension(argument))) return true
    return own.has(pathKey(argument))
  })
}

/**
 * Filters a second instance's command line in place for the running
 * workspace `mode`, without reading any file. When the new instance sent its
 * routing decision (`additionalData.simpleRouting`), only the paths it routed
 * here stay; otherwise only paths whose extension can never reach this
 * workspace are removed. Kept paths are made absolute against the new
 * instance's working folder. Routed paths missing from the command line
 * (the new instance read them from a --simple-open-list file) are appended,
 * and the list switch itself is removed.
 * @param {string[]} argv mutated
 * @param {{mode: string, workingDirectory?: string, additionalData?: unknown}} context
 * @returns {string[]} the same array
 */
function filterSecondInstanceArgv(argv, context) {
  if (!Array.isArray(argv)) return argv
  const routing = context.additionalData && typeof context.additionalData === 'object' ? context.additionalData.simpleRouting : null
  const decided = routing && routing.mode === context.mode && Array.isArray(routing.paths)
    ? new Set(routing.paths.filter((item) => typeof item === 'string').map((item) => pathKey(item)))
    : null
  const cwd = typeof context.workingDirectory === 'string' && context.workingDirectory ? context.workingDirectory : undefined
  const present = new Set()
  for (let index = argv.length - 1; index >= 0; index -= 1) {
    const argument = argv[index]
    // The new instance already read (and deleted) its list file; the paths
    // it routed here arrive in additionalData.
    if (typeof argument === 'string' && argument.startsWith(OPEN_LIST_SWITCH)) { argv.splice(index, 1); continue }
    if (typeof argument !== 'string' || argument.startsWith('-') || !MODE_BY_EXTENSION.has(normalizedExtension(argument))) continue
    const keep = decided ? decided.has(pathKey(argument, cwd)) : candidateModesForPath(argument).includes(context.mode)
    if (!keep) argv.splice(index, 1)
    else {
      if (cwd) argv[index] = path.resolve(cwd, argument)
      present.add(pathKey(argument, cwd))
    }
  }
  // Paths the new instance routed here that are not on its raw command line
  // (they came from a list file) are added, so the workspace opens them all.
  if (decided) {
    for (const item of routing.paths) {
      if (typeof item !== 'string' || !path.isAbsolute(item)) continue
      const key = pathKey(item)
      if (present.has(key)) continue
      present.add(key)
      argv.push(item)
    }
  }
  return argv
}

/**
 * Rewrites only the `build.fileAssociations` array of a package.json so it
 * equals the registry, keeping every other byte (formatting, line endings).
 * @param {string} packageJsonPath
 * @returns {boolean} true when the file changed.
 */
function writeFileAssociations(packageJsonPath) {
  const source = fs.readFileSync(packageJsonPath, 'utf8')
  const eol = source.includes('\r\n') ? '\r\n' : '\n'
  const key = /^([ \t]*)"fileAssociations"\s*:\s*\[/m.exec(source)
  if (!key) throw new Error(`${packageJsonPath} has no build.fileAssociations array.`)
  const open = key.index + key[0].length - 1
  let depth = 0
  let close = -1
  let quoted = false
  for (let index = open; index < source.length; index += 1) {
    const char = source[index]
    if (quoted) {
      if (char === '\\') index += 1
      else if (char === '"') quoted = false
    } else if (char === '"') quoted = true
    else if (char === '[') depth += 1
    else if (char === ']' && --depth === 0) { close = index; break }
  }
  if (close === -1) throw new Error(`${packageJsonPath} has an unterminated fileAssociations array.`)
  const indent = key[1]
  const rendered = JSON.stringify(formats.fileAssociations(), null, 2)
    .split('\n')
    .map((line, index) => (index === 0 ? line : `${indent}${line}`))
    .join(eol)
  const next = `${source.slice(0, open)}${rendered}${source.slice(close + 1)}`
  JSON.parse(next)
  if (next === source) return false
  fs.writeFileSync(packageJsonPath, next)
  return true
}

module.exports = {
  EXTENSIONS_BY_MODE,
  MODES,
  MODE_BY_EXTENSION,
  SUPPORTED_EXTENSIONS,
  candidateModesForPath,
  explicitMode,
  filterSecondInstanceArgv,
  groupPathsByMode,
  groupPathsByModeAsync,
  keepOwnPaths,
  modeForPath,
  normalizedExtension,
  routeCommandLine,
  routeForPath,
  routeForPathAsync,
  supportedPaths,
  writeFileAssociations,
}

// `node electron/routing.cjs --write-associations` regenerates the
// fileAssociations in simple/package.json after formats.json changes.
if (require.main === module) {
  if (process.argv.includes('--write-associations')) {
    const target = path.join(__dirname, '..', 'package.json')
    console.log(writeFileAssociations(target) ? `Updated fileAssociations in ${target}.` : 'fileAssociations are current.')
  } else {
    console.log(JSON.stringify(EXTENSIONS_BY_MODE, null, 2))
  }
}
