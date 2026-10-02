'use strict'

/**
 * Runs inside the unpackaged app's main process when SIMPLE_QA_DRIVER points
 * here (simple/electron/bootstrap.cjs ignores it in a packaged app). It reads
 * the step from SIMPLE_QA_DRIVER_CONFIG, runs it against the app's windows
 * with the workspace adapter, writes a JSON result and, unless the step keeps
 * the app running, exits.
 *
 * While it runs, documents handed out of a window are recorded instead of
 * opened: calls to Electron's shell.openPath / openExternal / openItem, and
 * requests on the shared io:open-in-simple channel (which would otherwise
 * start another copy of Simple). The result lists both, so a scenario can
 * prove "open after export" stays inside Simple.
 */

const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

function hashFile(filePath) {
  try { return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex') } catch { return null }
}

function tempLeftovers(folder) {
  try { return fs.readdirSync(folder).filter((name) => /^~simple-.*\.tmp$/i.test(name)) } catch { return [] }
}

/** Replaces the io:open-in-simple handler, now or when the workspace registers it, with a recorder. */
function recordHandOffs(ipcMain, record) {
  const recorder = (_event, filePath) => {
    record.openInSimple.push(String(filePath))
    return { ok: true, action: 'launched', shownInFolder: false, mode: null, appName: null, path: String(filePath) }
  }
  const handle = ipcMain.handle.bind(ipcMain)
  ipcMain.handle = (channel, listener) => handle(channel, channel === 'io:open-in-simple' ? recorder : listener)
  try {
    ipcMain.removeHandler('io:open-in-simple')
    handle('io:open-in-simple', recorder)
  } catch {}
}

function recordShell(shell, record) {
  for (const name of ['openPath', 'openExternal', 'openItem']) {
    try {
      shell[name] = (...args) => {
        record.shell.push({ call: name, args: args.map(String) })
        return Promise.resolve('')
      }
    } catch {}
  }
}

/**
 * Helpers handed to scenarios and adapters.
 * @param {typeof import('electron')} electron
 */
function makeHelpers(electron) {
  const { BrowserWindow } = electron
  const helpers = {
    sleep,
    hashFile,
    tempLeftovers,
    /** Polls until `check()` returns a truthy value. */
    async waitFor(check, timeoutMs = 30_000, label = 'condition') {
      const started = Date.now()
      for (;;) {
        const value = await check()
        if (value) return value
        if (Date.now() - started > timeoutMs) throw new Error(`Timed out waiting for ${label}.`)
        await sleep(150)
      }
    },
    /** Open app windows with a loaded page. */
    windows() {
      return BrowserWindow.getAllWindows().filter((win) => !win.isDestroyed() && win.webContents && !win.webContents.isLoading() && win.webContents.getURL())
    },
    /** The first app window, once its page has loaded. */
    async mainWindow(timeoutMs = 60_000) {
      return helpers.waitFor(() => helpers.windows()[0], timeoutMs, 'a window')
    },
    /** Runs JavaScript in the page (as a user gesture) and returns its value. */
    exec(win, code) {
      return win.webContents.executeJavaScript(code, true)
    },
    /** Clicks the first element matching a CSS selector; false when there is none. */
    click(win, selector) {
      return helpers.exec(win, `(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; el.click(); return true })()`)
    },
    /** Sends a key press such as "Control+S" or "Escape" to the page. */
    async key(win, accelerator) {
      const parts = accelerator.split('+')
      const keyCode = parts.pop()
      const modifiers = parts.map((part) => part.toLowerCase())
      win.webContents.focus()
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode, modifiers })
      if (keyCode.length === 1 && !modifiers.some((modifier) => modifier === 'control' || modifier === 'alt')) {
        win.webContents.sendInputEvent({ type: 'char', keyCode, modifiers })
      }
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode, modifiers })
      await sleep(50)
    },
    /** Types text into the focused element. */
    async type(win, text) {
      for (const char of text) {
        win.webContents.sendInputEvent({ type: 'char', keyCode: char })
        await sleep(5)
      }
    },
    /** Waits until a file's hash differs from `before` (null: until it exists) and no save temp is left beside it. */
    async fileWritten(filePath, before, timeoutMs = 60_000) {
      return helpers.waitFor(() => {
        const now = hashFile(filePath)
        return now && now !== before && !tempLeftovers(path.dirname(filePath)).length ? now : null
      }, timeoutMs, `${path.basename(filePath)} to be written`)
    },
  }
  return helpers
}

/** Steps the acceptance scripts can ask for. Each returns plain data for the result file. */
const STEPS = {
  /** Proves the harness itself: a window opens and its page answers. */
  async smoke(h) {
    const win = await h.mainWindow()
    const page = await h.exec(win, '({ title: document.title, simpleIO: typeof window.simpleIO })')
    return { windows: h.windows().length, ...page }
  },

  /** Edits the opened document and saves with Ctrl+S; waits for `watch` to be written. */
  async editSave(h, config, adapter) {
    const win = await h.mainWindow()
    await adapter.waitReady(win, h, config)
    await adapter.edit(win, h, config.marker, config)
    await h.key(win, 'Control+S')
    const hash = await h.fileWritten(config.watch, config.watchBefore || null, config.writeTimeoutMs)
    const status = typeof adapter.statusText === 'function' ? await adapter.statusText(win, h) : null
    return { hash, status }
  },

  /** Edits, waits for the recovery copy, then reports and stays running (the runner kills the app). */
  async editAndWait(h, config, adapter) {
    const win = await h.mainWindow()
    await adapter.waitReady(win, h, config)
    await adapter.edit(win, h, config.marker, config)
    await h.sleep(config.waitMs || 4000)
    return { edited: true }
  },

  /** Restores the entry offered on the welcome card, then saves with Ctrl+S. */
  async restoreSave(h, config, adapter) {
    const win = await h.mainWindow()
    const entries = await h.exec(win, 'window.simpleIO ? window.simpleIO.recovery.list() : []')
    await adapter.restoreFromCard(win, h, config)
    const status = typeof adapter.statusText === 'function' ? await adapter.statusText(win, h) : null
    await h.key(win, 'Control+S')
    const hash = await h.fileWritten(config.watch, config.watchBefore || null, config.writeTimeoutMs)
    return { offered: Array.isArray(entries) ? entries.length : 0, statusAfterRestore: status, hash }
  },

  /** Edits, starts a save and closes the window at once; the app must exit by itself after the write. */
  async closeWhileSaving(h, config, adapter) {
    const win = await h.mainWindow()
    await adapter.waitReady(win, h, config)
    await adapter.edit(win, h, config.marker, config)
    await h.key(win, 'Control+S')
    await h.sleep(100)
    win.close()
    return { closing: true }
  },

  /** Drops a file onto the window, then edits and saves it in place without a dialog. */
  async dropSave(h, config, adapter) {
    const win = await h.mainWindow()
    const path = await adapter.drop(win, h, config.dropFile, config)
    await adapter.waitReady(win, h, config)
    await adapter.edit(win, h, config.marker, config)
    await h.key(win, 'Control+S')
    const hash = await h.fileWritten(config.watch, config.watchBefore || null, config.writeTimeoutMs)
    return { droppedPath: path || null, hash }
  },

  /**
   * Opens the file and reads the notice the workspace shows when it opened a
   * file in a simpler way than the optional office engine would (no-office-matrix.cjs).
   */
  async openNotice(h, config, adapter) {
    const win = await h.mainWindow()
    await adapter.waitReady(win, h, config)
    if (typeof adapter.noticeText !== 'function') throw new Error('the adapter must implement noticeText(win, h) for notice rows')
    return { notice: await adapter.noticeText(win, h, config) }
  },

  /** Runs Export As with "Open the file after exporting" on; waits for the exported file. */
  async exportOpen(h, config, adapter) {
    const win = await h.mainWindow()
    await adapter.waitReady(win, h, config)
    await adapter.exportAs(win, h, { format: config.format, openAfter: true }, config)
    const hash = await h.fileWritten(config.watch, null, config.writeTimeoutMs)
    await h.sleep(500)
    return { hash }
  },
}

/**
 * Entry point called by bootstrap.cjs after app ready.
 * @param {{app: object, mode: string}} context
 */
module.exports = async function runDriver({ app }) {
  const electron = require('electron')
  const record = { openInSimple: [], shell: [] }
  // Synchronously, before the workspace's own ready handlers register their IPC.
  recordHandOffs(electron.ipcMain, record)
  recordShell(electron.shell, record)
  const config = JSON.parse(fs.readFileSync(process.env.SIMPLE_QA_DRIVER_CONFIG, 'utf8'))
  const result = { step: config.scenario }
  const write = () => fs.writeFileSync(config.resultPath, JSON.stringify({ ...result, handOffs: record }))
  try {
    const step = STEPS[config.scenario]
    if (!step) throw new Error(`Unknown step "${config.scenario}".`)
    const adapter = config.adapterPath ? require(config.adapterPath) : null
    result.value = await step(makeHelpers(electron), config, adapter)
    result.ok = true
  } catch (error) {
    result.ok = false
    result.error = String((error && error.stack) || error)
  }
  write()
  if (config.keepRunning) {
    // The app keeps running (closing or being killed is part of the step); later hand-offs are still recorded.
    const timer = setInterval(write, 250)
    app.once('will-quit', () => { clearInterval(timer); write() })
    return
  }
  app.exit(result.ok ? 0 : 1)
}
