'use strict'

const path = require('node:path')
// Electron's fs treats app.asar as a folder (one made-up time for every build,
// and rm would recurse into it); original-fs sees the real file. Plain Node
// (tests) has only node:fs.
const fs = (() => {
  try { return require('original-fs').promises } catch { return require('node:fs/promises') }
})()

// The portable launcher (build/portable-fast.nsi) unpacks each build once into
// %LOCALAPPDATA%\simple\cache\<build id> (about 360 MB) and reuses it, but never
// removed the folders of earlier builds. A running Simple removes the folders of
// builds older than its own. Windows refuses to rename a folder while a runtime
// in it is running (Electron holds app.asar and its .pak files open), so each
// folder is first renamed aside and only deleted when that succeeds. A removed
// build that is started again is unpacked again by its launcher.

const TRASH = '.simple-trash-'

/** When this build was made: the launcher keeps the packaged files' times. */
async function buildTime(directory) {
  try {
    return (await fs.stat(path.join(directory, 'resources', 'app.asar'))).mtimeMs
  } catch {
    return null
  }
}

/**
 * Removes earlier builds next to the runtime at `execPath`. Never throws.
 * Only acts inside a ...\simple\cache\<build id>\ folder.
 * @returns {Promise<{removed: string[], kept: string[]}>} folder names
 */
async function pruneRuntimeCache({ execPath }) {
  const result = { removed: [], kept: [] }
  try {
    const ownDirectory = path.dirname(execPath)
    const root = path.dirname(ownDirectory)
    if (path.basename(root).toLowerCase() !== 'cache' || path.basename(path.dirname(root)).toLowerCase() !== 'simple') return result
    const own = await buildTime(ownDirectory)
    if (own === null) return result
    const ownName = path.basename(ownDirectory)
    for (const entry of await fs.readdir(root, { withFileTypes: true })) {
      if (!entry.isDirectory() || entry.name === ownName) continue
      const directory = path.join(root, entry.name)
      if (entry.name.includes(TRASH)) {
        // Left by a removal that was interrupted; nothing runs from it.
        await fs.rm(directory, { recursive: true, force: true }).catch(() => {})
        continue
      }
      // The same or a newer build stays; a folder without a runtime is incomplete.
      const built = await buildTime(directory)
      if (built !== null && built >= own) {
        result.kept.push(entry.name)
        continue
      }
      const aside = `${directory}${TRASH}${process.pid}`
      try {
        await fs.rename(directory, aside)
      } catch {
        result.kept.push(entry.name)
        continue
      }
      await fs.rm(aside, { recursive: true, force: true, maxRetries: 2 }).catch(() => {})
      result.removed.push(entry.name)
    }
  } catch {
    // Cleanup is best effort; the app never depends on it.
  }
  return result
}

module.exports = { pruneRuntimeCache }
