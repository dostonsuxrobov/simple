'use strict'

// Electron entry point for tests/office-import.test.cjs. It loads
// electron/main.cjs with a few Electron APIs wrapped so that no window is ever
// shown and the app UI is never loaded, then runs a job file:
//   { outDir, convert: [{ id, file, name? }], ipc: [{ id, channel, args }] }
// Every conversion uses the real hidden print windows. Results (and the main
// process's worst event-loop stall) are written to <outDir>/results.json.
const fs = require('node:fs')
const path = require('node:path')
const Module = require('node:module')
const electron = require('electron')

const job = JSON.parse(fs.readFileSync(process.env.SIMPLE_IMPORT_JOB, 'utf8'))
const handlers = new Map()
const appWindows = []
const openDialogs = []
const openAnswers = [...(job.openAnswers || [])]

class HiddenBrowserWindow extends electron.BrowserWindow {
  constructor(options = {}) {
    super({ ...options, show: false })
    this.isAppWindow = Boolean(options.webPreferences?.preload)
    if (this.isAppWindow) appWindows.push(this)
  }
  show() {}
  showInactive() {}
  focus() {}
  loadFile(file, options) { return this.isAppWindow ? Promise.resolve() : super.loadFile(file, options) }
  loadURL(url, options) { return this.isAppWindow ? Promise.resolve() : super.loadURL(url, options) }
}

const app = new Proxy(electron.app, {
  get(target, property) {
    if (property === 'requestSingleInstanceLock') return () => true
    if (property === 'quit') return () => {}
    if (property === 'on') {
      return (event, listener) => (['window-all-closed', 'second-instance', 'before-quit', 'activate'].includes(event) ? app : target.on(event, listener))
    }
    const value = Reflect.get(target, property)
    return typeof value === 'function' ? value.bind(target) : value
  },
})
const ipcMain = { handle: (channel, handler) => handlers.set(channel, handler), on() {}, removeHandler() {} }
const dialog = {
  async showOpenDialog(...args) {
    openDialogs.push(args.at(-1))
    const answer = openAnswers.shift()
    return answer ? { canceled: false, filePaths: answer } : { canceled: true, filePaths: [] }
  },
  async showSaveDialog() { return { canceled: true } },
  async showMessageBox() { return { response: 0 } },
}
const wrapped = new Proxy(electron, {
  get(target, property) {
    if (property === 'BrowserWindow') return HiddenBrowserWindow
    if (property === 'app') return app
    if (property === 'ipcMain') return ipcMain
    if (property === 'dialog') return dialog
    return Reflect.get(target, property)
  },
})
const originalLoad = Module._load
Module._load = function load(request, parent, isMain) {
  if (request === 'electron') return wrapped
  return originalLoad.call(this, request, parent, isMain)
}

electron.app.disableHardwareAcceleration()
electron.app.on('window-all-closed', () => {})

function describeError(error) {
  return { code: error?.code || null, name: error?.name || null, message: String(error?.message || error), details: error?.details ?? null }
}

function summarize(value, id) {
  if (!value || typeof value !== 'object') return value
  const result = {}
  for (const [key, item] of Object.entries(value)) {
    if (item instanceof Uint8Array || Buffer.isBuffer(item)) {
      const target = path.join(job.outDir, `${id}.pdf`)
      fs.writeFileSync(target, item)
      result[key] = { file: target, length: item.length }
    } else result[key] = item
  }
  return result
}

async function main() {
  const results = { convert: {}, ipc: {}, openDialogs, maxStallMs: 0 }
  let last = Date.now()
  let itemStall = 0
  const watch = setInterval(() => {
    const now = Date.now()
    const stall = now - last - 25
    results.maxStallMs = Math.max(results.maxStallMs, stall)
    itemStall = Math.max(itemStall, stall)
    last = now
  }, 25)
  try {
    require(path.join(__dirname, '..', '..', 'electron', 'main.cjs'))
    await electron.app.whenReady()
    for (let attempt = 0; attempt < 200 && (!handlers.size || !appWindows.length); attempt += 1) await new Promise((resolve) => setTimeout(resolve, 10))
    const { convertToPdf } = require(path.join(__dirname, '..', '..', 'electron', 'office-import.cjs'))
    for (const item of job.convert || []) {
      const input = fs.readFileSync(item.file)
      await new Promise((resolve) => setTimeout(resolve, 60))
      itemStall = 0
      const started = Date.now()
      try {
        const output = await convertToPdf(input, { name: item.name || path.basename(item.file), sourcePath: item.file })
        const ms = Date.now() - started
        const target = path.join(job.outDir, `${item.id}.pdf`)
        fs.writeFileSync(target, output.data)
        results.convert[item.id] = { ok: true, file: target, kind: output.kind, warnings: output.warnings, ms, stallMs: itemStall }
      } catch (error) {
        results.convert[item.id] = { ok: false, error: describeError(error), ms: Date.now() - started, stallMs: itemStall }
      }
    }
    const sender = appWindows[0].webContents
    // { $file } becomes the file's bytes and { $inputs } a list of dropped files.
    const resolve = (value) => {
      if (!value || typeof value !== 'object') return value
      if (value.$file) return new Uint8Array(fs.readFileSync(value.$file))
      if (value.$inputs) return value.$inputs.map((input) => ({ name: input.name, data: new Uint8Array(fs.readFileSync(input.file)) }))
      if (Array.isArray(value)) return value.map(resolve)
      return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, resolve(item)]))
    }
    for (const call of job.ipc || []) {
      const started = Date.now()
      const args = (call.args || []).map(resolve)
      try {
        const handler = handlers.get(call.channel)
        if (!handler) throw new Error(`No IPC handler for ${call.channel}`)
        const value = await handler({ sender }, ...args)
        results.ipc[call.id] = { ok: true, value: summarize(value, call.id), ms: Date.now() - started }
      } catch (error) {
        // What the renderer receives: Electron forwards only String(error).
        results.ipc[call.id] = { ok: false, error: describeError(error), rendererText: String(error), ms: Date.now() - started }
      }
    }
  } catch (error) {
    results.fatal = describeError(error)
  } finally {
    clearInterval(watch)
    fs.writeFileSync(path.join(job.outDir, 'results.json'), JSON.stringify(results, null, 2))
    electron.app.exit(0)
  }
}

main()
