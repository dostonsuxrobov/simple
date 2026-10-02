'use strict'

// Long file lists between Simple's own processes.
//
// The portable EXE is a small wrapper that forwards its command line to the
// real app through a fixed 8,192-character buffer; a longer command line
// reaches the app as "too_long" or cut in the middle of a path. So when a
// list of files would make the command line long, the starting process
// writes the paths to a list file in Simple's own profile folder and passes
// only --simple-open-list=<file>. The new process reads the list, deletes it
// and routes the paths as if they had been on its command line.
//
// Only list files in <profile>/open-lists named open-<32 hex>.json are ever
// read or deleted, so a crafted command line can't make Simple remove
// another file. Node built-ins only.

const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')

const OPEN_LIST_SWITCH = '--simple-open-list='
const OPEN_LIST_FOLDER = 'open-lists'
const OPEN_LIST_NAME = /^open-[0-9a-f]{32}\.json$/
/** Above this many command-line characters (program, quotes and spaces included) a list file is used. */
const MAX_COMMAND_LINE_CHARS = 6000
const MAX_LIST_BYTES = 8 * 1024 * 1024
const MAX_LIST_PATHS = 10_000
/** List files a process never picked up (it failed to start) are removed after a day. */
const STALE_LIST_MS = 24 * 60 * 60 * 1000

/**
 * The folder that holds list files for a profile.
 * @param {string} stateRoot the profile root (SIMPLE_USER_DATA_DIR)
 * @returns {string}
 */
function openListFolder(stateRoot) {
  return path.join(path.resolve(stateRoot), OPEN_LIST_FOLDER)
}

/**
 * Characters Windows needs for a command line: each argument quoted and
 * separated by a space.
 * @param {string} command
 * @param {string[]} args
 * @returns {number}
 */
function commandLineLength(command, args) {
  return [command, ...args].reduce((total, argument) => total + String(argument).length + 3, 0)
}

/**
 * True when a command line is too long for the portable wrapper.
 * @param {string} command
 * @param {string[]} args
 * @param {number} [limit]
 * @returns {boolean}
 */
function needsOpenList(command, args, limit = MAX_COMMAND_LINE_CHARS) {
  return commandLineLength(command, args) > limit
}

function checkListPaths(paths) {
  if (!Array.isArray(paths) || paths.some((item) => typeof item !== 'string' || !path.isAbsolute(item))) {
    throw new TypeError('An open list holds absolute file paths only.')
  }
  if (paths.length > MAX_LIST_PATHS) throw new RangeError(`Open up to ${MAX_LIST_PATHS.toLocaleString('en-US')} files at a time.`)
}

function newListFile(folder) {
  return path.join(folder, `open-${crypto.randomBytes(16).toString('hex')}.json`)
}

/**
 * Writes file paths to a new list file.
 * @param {string} stateRoot the profile root
 * @param {string[]} paths absolute file paths
 * @returns {Promise<string>} the list file
 */
async function writeOpenList(stateRoot, paths) {
  checkListPaths(paths)
  const folder = openListFolder(stateRoot)
  await fs.promises.mkdir(folder, { recursive: true })
  const file = newListFile(folder)
  await fs.promises.writeFile(file, JSON.stringify({ version: 1, paths }), { flag: 'wx' })
  return file
}

/**
 * writeOpenList() in the same tick. launchDetached uses it so the new
 * process is started before the caller goes on: the bootstrap loads its
 * workspace right after, and a workspace that is already running quits this
 * process while it loads, before any awaited write could finish.
 * @param {string} stateRoot the profile root
 * @param {string[]} paths absolute file paths
 * @returns {string} the list file
 */
function writeOpenListSync(stateRoot, paths) {
  checkListPaths(paths)
  const folder = openListFolder(stateRoot)
  fs.mkdirSync(folder, { recursive: true })
  const file = newListFile(folder)
  fs.writeFileSync(file, JSON.stringify({ version: 1, paths }), { flag: 'wx' })
  return file
}

/**
 * Removes a list file synchronously (a process that is about to quit). Never throws.
 * @param {string} file
 */
function removeOpenListSync(file) {
  try { fs.unlinkSync(file) } catch {}
}

/**
 * Removes a list file this process wrote (for example when the process that
 * should have read it could not start). Never throws.
 * @param {string} file
 */
async function removeOpenList(file) {
  try { await fs.promises.unlink(file) } catch {}
}

function sameFolder(a, b) {
  return process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b
}

/** Reads and deletes one list file; [] when it is not a valid list of this profile. */
function takeOpenList(value, folder) {
  const file = path.resolve(String(value || ''))
  if (!sameFolder(path.dirname(file), folder) || !OPEN_LIST_NAME.test(path.basename(file))) return []
  let parsed = null
  try {
    const stat = fs.statSync(file)
    if (stat.isFile() && stat.size <= MAX_LIST_BYTES) parsed = JSON.parse(fs.readFileSync(file, 'utf8'))
  } catch {}
  try { fs.unlinkSync(file) } catch {}
  const paths = parsed && parsed.version === 1 && Array.isArray(parsed.paths) ? parsed.paths : []
  return paths.filter((item) => typeof item === 'string' && path.isAbsolute(item)).slice(0, MAX_LIST_PATHS)
}

/** Deletes list files older than a day. Never throws. */
function sweepOpenLists(folder, now = Date.now()) {
  let names = []
  try { names = fs.readdirSync(folder) } catch { return }
  for (const name of names) {
    if (!OPEN_LIST_NAME.test(name)) continue
    const file = path.join(folder, name)
    try {
      if (now - fs.statSync(file).mtimeMs > STALE_LIST_MS) fs.unlinkSync(file)
    } catch {}
  }
}

/**
 * Replaces every --simple-open-list=<file> argument with the paths in that
 * file, deleting the file. Lists outside `<stateRoot>/open-lists` are
 * dropped unread. Other arguments are kept in place. Every call (every
 * start of Simple) also removes list files older than a day that no process
 * picked up, so the paths of a failed hand-off never stay in the profile.
 * @param {string[]} argv
 * @param {string} stateRoot the profile root
 * @returns {string[]} a new array
 */
function expandOpenLists(argv, stateRoot) {
  if (!Array.isArray(argv)) return argv
  const folder = openListFolder(stateRoot)
  if (!argv.some((argument) => typeof argument === 'string' && argument.startsWith(OPEN_LIST_SWITCH))) {
    sweepOpenLists(folder)
    return [...argv]
  }
  const expanded = []
  for (const argument of argv) {
    if (typeof argument === 'string' && argument.startsWith(OPEN_LIST_SWITCH)) expanded.push(...takeOpenList(argument.slice(OPEN_LIST_SWITCH.length), folder))
    else expanded.push(argument)
  }
  sweepOpenLists(folder)
  return expanded
}

/**
 * True for a --simple-open-list=<file> argument.
 * @param {unknown} argument
 * @returns {boolean}
 */
function isOpenListArgument(argument) {
  return typeof argument === 'string' && argument.startsWith(OPEN_LIST_SWITCH)
}

module.exports = {
  MAX_COMMAND_LINE_CHARS,
  MAX_LIST_PATHS,
  OPEN_LIST_SWITCH,
  commandLineLength,
  expandOpenLists,
  isOpenListArgument,
  needsOpenList,
  openListFolder,
  removeOpenList,
  removeOpenListSync,
  sweepOpenLists,
  writeOpenList,
  writeOpenListSync,
}
