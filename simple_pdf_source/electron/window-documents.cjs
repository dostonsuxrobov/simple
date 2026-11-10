'use strict'

const path = require('node:path')

// Which window shows which PDF file, and what that file looked like on disk
// (size and modified time) when the window last read or wrote it. The main
// process uses this to focus the window that already shows a file instead of
// opening it twice, and to ask before a save replaces a file that another
// program (or another window) changed in the meantime.
// No Electron here: windows are known only by their webContents id.

/** Windows file names ignore case, so "C:\A\x.PDF" and "c:\a\x.pdf" are one file. */
function defaultFileKey(filePath) {
  return path.resolve(String(filePath)).toLowerCase()
}

/** True when the file on disk is not the one described by `stamp`. */
function changedOnDisk(stamp, stat) {
  if (!stamp || !stat) return false
  return stamp.size !== stat.size || stamp.mtimeMs !== stat.mtimeMs
}

/**
 * @param {object} [options]
 * @param {(filePath: string) => string} [options.fileKey] one key per file on disk
 * @param {() => number} [options.now]
 * @param {number} [options.pendingMs] how long a window that was just created
 *   for a file counts as showing it before its page has confirmed the open
 */
function createWindowDocuments(options = {}) {
  const fileKey = options.fileKey || defaultFileKey
  const now = options.now || Date.now
  const pendingMs = options.pendingMs ?? 10_000
  /** window id -> { path, key, pendingUntil } */
  const current = new Map()
  /** window id -> Map(file key -> { size, mtimeMs }) */
  const stamps = new Map()

  function keyOf(filePath) {
    try {
      return filePath ? fileKey(filePath) : null
    } catch {
      return null
    }
  }

  return {
    /**
     * A window was created to open this file. It counts as showing the file
     * for a short while, so opening the same file twice in a row (a double
     * double-click) does not create two windows. The page confirms with
     * setCurrent(); if the open fails, the claim simply runs out.
     */
    expect(windowId, filePath) {
      const key = keyOf(filePath)
      if (key) current.set(windowId, { path: String(filePath), key, pendingUntil: now() + pendingMs })
    },

    /** The window now shows this file (null: no file, such as the home screen or an unsaved copy). */
    setCurrent(windowId, filePath) {
      const key = keyOf(filePath)
      if (!key) {
        current.delete(windowId)
        stamps.delete(windowId)
        return
      }
      current.set(windowId, { path: String(filePath), key, pendingUntil: null })
      // Stamps of files this window showed earlier are no longer needed.
      const known = stamps.get(windowId)
      const stamp = known?.get(key)
      if (stamp) stamps.set(windowId, new Map([[key, stamp]]))
      else stamps.delete(windowId)
    },

    /** Remember what the file looked like when this window read or wrote it. */
    noteStamp(windowId, filePath, stat) {
      const key = keyOf(filePath)
      if (!key) return
      if (!stat) {
        // Unknown state: better no check at the next save than a false alarm.
        stamps.get(windowId)?.delete(key)
        return
      }
      if (!stamps.has(windowId)) stamps.set(windowId, new Map())
      stamps.get(windowId).set(key, { size: stat.size, mtimeMs: stat.mtimeMs })
    },

    /** @returns {{size: number, mtimeMs: number} | null} */
    stampFor(windowId, filePath) {
      const key = keyOf(filePath)
      return (key && stamps.get(windowId)?.get(key)) || null
    },

    /** @returns the id of a window that shows this file, or null */
    findWindow(filePath, exceptWindowId = null) {
      const key = keyOf(filePath)
      if (!key) return null
      for (const [windowId, entry] of current) {
        if (windowId === exceptWindowId || entry.key !== key) continue
        if (entry.pendingUntil !== null && entry.pendingUntil < now()) continue
        return windowId
      }
      return null
    },

    /** The open this window was created for failed: it no longer counts as showing that file. */
    releasePending(windowId) {
      if (current.get(windowId)?.pendingUntil != null) current.delete(windowId)
    },

    /** The window closed. */
    forget(windowId) {
      current.delete(windowId)
      stamps.delete(windowId)
    },
  }
}

module.exports = { changedOnDisk, createWindowDocuments, defaultFileKey }
