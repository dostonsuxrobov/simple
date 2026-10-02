'use strict'

// Loads electron/main.cjs with a stubbed `electron` module so the real
// ipcMain handlers run in plain Node. Every test file gets its own process
// under `node --test`, so the stub is installed once per file.
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const Module = require('node:module')

let harness = null

function createHarness() {
  // Keep the app's temp-folder housekeeping inside a private folder.
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-pdf-harness-'))
  process.env.TEMP = temp
  process.env.TMP = temp
  process.once('exit', () => {
    try { fs.rmSync(temp, { recursive: true, force: true, maxRetries: 3 }) } catch { /* best effort */ }
  })

  const handlers = new Map()
  const saveDialogs = []
  const saveAnswers = []
  let resolveReady
  const ready = new Promise((resolve) => { resolveReady = resolve })

  class FakeWebContents {
    constructor() { this.id = 1 }
    on() {}
    once() {}
    send() {}
    startDrag() {}
    async getPrintersAsync() { return [] }
  }
  class FakeBrowserWindow {
    constructor() { this.webContents = new FakeWebContents() }
    static fromWebContents() { return null }
    static getAllWindows() { return [] }
    static getFocusedWindow() { return null }
    removeMenu() {}
    loadURL() {}
    loadFile() {}
    once() {}
    on() {}
    show() {}
    close() {}
    destroy() {}
    isDestroyed() { return false }
  }
  const electron = {
    app: {
      isPackaged: true,
      requestSingleInstanceLock: () => true,
      on() {},
      whenReady: () => ready,
      getVersion: () => 'test',
      quit() {},
    },
    BrowserWindow: FakeBrowserWindow,
    dialog: {
      async showSaveDialog(options) {
        saveDialogs.push(options)
        const answer = saveAnswers.shift()
        return answer ? { canceled: false, filePath: answer } : { canceled: true }
      },
      async showOpenDialog() { return { canceled: true, filePaths: [] } },
    },
    ipcMain: {
      handle: (channel, handler) => handlers.set(channel, handler),
      on() {},
    },
    nativeImage: { createFromPath: () => ({}) },
    shell: { showItemInFolder() {}, async openExternal() {} },
  }

  const originalLoad = Module._load
  Module._load = function load(request, parent, isMain) {
    if (request === 'electron') return electron
    return originalLoad.call(this, request, parent, isMain)
  }
  require('../../electron/main.cjs')
  resolveReady()

  async function invoke(channel, ...args) {
    await ready
    await new Promise((resolve) => setImmediate(resolve))
    const handler = handlers.get(channel)
    if (!handler) throw new Error(`No IPC handler for ${channel}`)
    try {
      return await handler({ sender: new FakeWebContents() }, ...args)
    } catch (error) {
      // What the renderer receives: Electron forwards only String(error).
      throw new Error(`Error invoking remote method '${channel}': ${error}`)
    }
  }

  function tempFile(name, bytes) {
    const target = path.join(temp, name)
    if (bytes !== undefined) fs.writeFileSync(target, bytes)
    return target
  }

  return { invoke, handlers, saveDialogs, saveAnswers, temp, tempFile }
}

function loadMain() {
  harness ||= createHarness()
  return harness
}

module.exports = { loadMain }
