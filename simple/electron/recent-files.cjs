'use strict'

// The files Simple was asked to open, newest first, for the start screen.
// One small JSON file in the profile root, written whenever a command line
// with files is routed (every way of opening a file from Windows or from the
// start screen goes through the bootstrap). Never throws: a list that cannot
// be read or written only means an emptier start screen.

const fs = require('node:fs')
const path = require('node:path')

const FILE = 'recent-files.json'
const LIMIT = 30

function read(stateRoot) {
  try {
    const parsed = JSON.parse(fs.readFileSync(path.join(stateRoot, FILE), 'utf8'))
    return Array.isArray(parsed) ? parsed.filter((entry) => entry && typeof entry.path === 'string' && path.isAbsolute(entry.path)) : []
  } catch {
    return []
  }
}

const key = (filePath) => (process.platform === 'win32' ? filePath.toLowerCase() : filePath)

/**
 * Puts these files at the top of the list.
 * @param {string} stateRoot the profile root (every Simple process shares it)
 * @param {Array<{path: string, mode: string}>} files absolute paths with their workspace
 * @param {number} [now]
 */
function recordRecent(stateRoot, files, now = Date.now()) {
  if (!Array.isArray(files) || !files.length) return
  try {
    const added = files.filter((entry) => entry && typeof entry.path === 'string' && path.isAbsolute(entry.path))
      .map((entry) => ({ path: entry.path, mode: String(entry.mode || ''), opened: now }))
    const seen = new Set(added.map((entry) => key(entry.path)))
    const next = [...added, ...read(stateRoot).filter((entry) => !seen.has(key(entry.path)))].slice(0, LIMIT)
    fs.mkdirSync(stateRoot, { recursive: true })
    const target = path.join(stateRoot, FILE)
    const temporary = `${target}.${process.pid}.tmp`
    fs.writeFileSync(temporary, JSON.stringify(next))
    fs.renameSync(temporary, target)
  } catch {
    // The start screen is a convenience; opening the file never depends on it.
  }
}

/**
 * The list for the start screen: files that still exist, newest first.
 * @param {string} stateRoot
 * @param {number} [limit]
 * @returns {Array<{path: string, name: string, folder: string, mode: string, opened: number}>}
 */
function listRecent(stateRoot, limit = 12) {
  const result = []
  for (const entry of read(stateRoot)) {
    if (result.length >= limit) break
    try { if (!fs.statSync(entry.path).isFile()) continue } catch { continue }
    result.push({ path: entry.path, name: path.basename(entry.path), folder: path.dirname(entry.path), mode: String(entry.mode || ''), opened: Number(entry.opened) || 0 })
  }
  return result
}

/**
 * The same list without ever holding up the process: every file is looked at in parallel,
 * and one that does not answer in time (a share that went offline) is left out for now.
 * The shared process also runs the document windows, so this must not block.
 * @param {string} stateRoot
 * @param {number} [limit]
 * @param {number} [timeoutMs] per file
 */
async function listRecentAsync(stateRoot, limit = 12, timeoutMs = 600) {
  const entries = read(stateRoot)
  const present = await Promise.all(entries.map((entry) => Promise.race([
    fs.promises.stat(entry.path).then((stat) => stat.isFile(), () => false),
    new Promise((resolve) => { const timer = setTimeout(() => resolve(false), timeoutMs); if (typeof timer.unref === 'function') timer.unref() }),
  ])))
  return entries.filter((_, index) => present[index]).slice(0, limit)
    .map((entry) => ({ path: entry.path, name: path.basename(entry.path), folder: path.dirname(entry.path), mode: String(entry.mode || ''), opened: Number(entry.opened) || 0 }))
}

/** Forgets one file (the start screen's remove button). */
function forgetRecent(stateRoot, filePath) {
  try {
    const next = read(stateRoot).filter((entry) => key(entry.path) !== key(String(filePath)))
    const target = path.join(stateRoot, FILE)
    const temporary = `${target}.${process.pid}.tmp`
    fs.writeFileSync(temporary, JSON.stringify(next))
    fs.renameSync(temporary, target)
  } catch {
    // Nothing to forget.
  }
}

module.exports = { LIMIT, forgetRecent, listRecent, listRecentAsync, recordRecent }
