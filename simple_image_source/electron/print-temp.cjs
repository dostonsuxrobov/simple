'use strict'

const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')

const PRINT_DIRECTORY_PREFIX = 'simple-image-print-'
const OWNER_FILE = '.simple-print-owner.json'
const LEGACY_STALE_AGE_MS = 24 * 60 * 60 * 1000

function isDirectPrintDirectory(directory, root = os.tmpdir()) {
  const resolvedRoot = path.resolve(root)
  const resolvedDirectory = path.resolve(directory)
  return path.dirname(resolvedDirectory) === resolvedRoot && path.basename(resolvedDirectory).startsWith(PRINT_DIRECTORY_PREFIX)
}

function processIsAlive(pid) {
  if (!Number.isSafeInteger(pid) || pid <= 0) return false
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return error?.code === 'EPERM'
  }
}

async function createOwnedPrintDirectory(root = os.tmpdir(), owner = {}) {
  const resolvedRoot = path.resolve(root)
  const directory = await fs.mkdtemp(path.join(resolvedRoot, PRINT_DIRECTORY_PREFIX))
  if (!isDirectPrintDirectory(directory, resolvedRoot)) {
    throw new Error('The print directory was created outside the temporary folder.')
  }
  try {
    await fs.writeFile(path.join(directory, OWNER_FILE), JSON.stringify({
      pid: Number.isSafeInteger(owner.pid) && owner.pid > 0 ? owner.pid : process.pid,
      startedAt: Number.isFinite(owner.startedAt) ? owner.startedAt : Date.now(),
    }), { encoding: 'utf8', flag: 'wx' })
    return directory
  } catch (error) {
    await fs.rm(directory, { recursive: true, force: true }).catch(() => {})
    throw error
  }
}

async function removePrintDirectory(directory, root = os.tmpdir()) {
  if (!isDirectPrintDirectory(directory, root)) return false
  await fs.rm(path.resolve(directory), { recursive: true, force: true })
  return true
}

async function cleanupStalePrintDirectories(options = {}) {
  const root = path.resolve(options.root || os.tmpdir())
  const now = Number.isFinite(options.now) ? options.now : Date.now()
  const isAlive = typeof options.isProcessAlive === 'function' ? options.isProcessAlive : processIsAlive
  let entries = []
  try { entries = await fs.readdir(root, { withFileTypes: true }) } catch { return { removed: [], preserved: [] } }
  const removed = []
  const preserved = []

  for (const entry of entries) {
    if (!entry.isDirectory() || !entry.name.startsWith(PRINT_DIRECTORY_PREFIX)) continue
    const directory = path.join(root, entry.name)
    if (!isDirectPrintDirectory(directory, root)) continue
    let owner = null
    try {
      const parsed = JSON.parse(await fs.readFile(path.join(directory, OWNER_FILE), 'utf8'))
      if (Number.isSafeInteger(parsed?.pid) && parsed.pid > 0 && Number.isFinite(parsed?.startedAt)) owner = parsed
    } catch {
      owner = null
    }

    if (owner && isAlive(owner.pid)) {
      preserved.push(directory)
      continue
    }
    if (!owner) {
      let stat
      try { stat = await fs.stat(directory) } catch { continue }
      if (now - stat.mtimeMs < LEGACY_STALE_AGE_MS) {
        preserved.push(directory)
        continue
      }
    }
    if (await removePrintDirectory(directory, root).catch(() => false)) removed.push(directory)
  }
  return { removed, preserved }
}

module.exports = {
  LEGACY_STALE_AGE_MS,
  OWNER_FILE,
  PRINT_DIRECTORY_PREFIX,
  cleanupStalePrintDirectories,
  createOwnedPrintDirectory,
  isDirectPrintDirectory,
  processIsAlive,
  removePrintDirectory,
}
