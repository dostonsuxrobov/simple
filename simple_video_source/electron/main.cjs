'use strict'

const { app, BrowserWindow, dialog, ipcMain, session, shell } = require('electron')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const { SUPPORTED_EXTENSIONS, isSupportedVideoPath, supportedPaths } = require('./routing.cjs')
const { mergeRecent, sanitizeRecents } = require('./recent-files.cjs')

const MAX_VIDEO_BYTES = 1024 * 1024 * 1024 * 1024
const RECENT_FILE = 'video-recents.json'
let ready = false
const pendingMacPaths = []
let recentMutations = Promise.resolve()

function callingWindow(event) {
  const window = BrowserWindow.fromWebContents(event.sender)
  return window && !window.isDestroyed() ? window : null
}

function recentStatePath() {
  return path.join(app.getPath('userData'), RECENT_FILE)
}

async function readRecents() {
  try {
    return sanitizeRecents(JSON.parse(await fs.readFile(recentStatePath(), 'utf8')))
  } catch {
    return []
  }
}

async function writeRecents(recents) {
  const target = recentStatePath()
  const token = crypto.randomUUID()
  const temporary = `${target}.${token}.tmp`
  const backup = `${target}.${token}.bak`
  let backedUp = false
  await fs.mkdir(path.dirname(target), { recursive: true })
  await fs.writeFile(temporary, JSON.stringify(sanitizeRecents(recents), null, 2), 'utf8')
  try {
    try {
      await fs.rename(target, backup)
      backedUp = true
    } catch (error) {
      if (error.code !== 'ENOENT') throw error
    }
    await fs.rename(temporary, target)
    if (backedUp) await fs.rm(backup, { force: true })
  } catch (error) {
    await fs.rm(temporary, { force: true }).catch(() => {})
    if (backedUp) await fs.rename(backup, target).catch(() => {})
    throw error
  }
}

function mutateRecents(updater) {
  const mutation = recentMutations.then(async () => {
    const next = sanitizeRecents(await updater(await readRecents()))
    await writeRecents(next)
    return next
  })
  recentMutations = mutation.then(() => undefined, () => undefined)
  return mutation
}

async function addRecent(filePath) {
  await mutateRecents((current) => mergeRecent(current, filePath))
}

async function videoPayload(filePath) {
  if (typeof filePath !== 'string' || !isSupportedVideoPath(filePath)) {
    throw new Error('Simple Video does not support this file type.')
  }
  const resolved = path.resolve(filePath)
  const stat = await fs.stat(resolved)
  if (!stat.isFile()) throw new Error('The selected item is not a video file.')
  if (stat.size <= 0) throw new Error('This video file is empty.')
  if (stat.size > MAX_VIDEO_BYTES) throw new Error('This video file is too large to open safely.')
  await addRecent(resolved)
  return {
    path: resolved,
    name: path.basename(resolved),
    url: pathToFileURL(resolved).href,
    size: stat.size,
    modifiedAt: stat.mtimeMs,
    extension: path.extname(resolved).slice(1).toUpperCase(),
  }
}

function createWindow(openPath = null) {
  const window = new BrowserWindow({
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    width: 1280,
    height: 800,
    minWidth: 760,
    minHeight: 500,
    show: false,
    frame: false,
    backgroundColor: '#111111',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: !app.isPackaged,
      webviewTag: false,
    },
  })

  window.removeMenu()
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  window.webContents.on('will-navigate', (event, destination) => {
    const current = window.webContents.getURL()
    if (current && destination !== current) event.preventDefault()
  })
  window.once('ready-to-show', () => window.show())
  window.on('maximize', () => window.webContents.send('window:maximized', true))
  window.on('unmaximize', () => window.webContents.send('window:maximized', false))
  window.on('enter-full-screen', () => window.webContents.send('window:fullscreen', true))
  window.on('leave-full-screen', () => window.webContents.send('window:fullscreen', false))

  if (process.env.VITE_DEV_SERVER_URL) {
    window.loadURL(process.env.VITE_DEV_SERVER_URL)
  } else {
    window.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))
  }
  window.webContents.once('did-finish-load', () => {
    if (openPath) window.webContents.send('file:open-external', openPath)
  })
  return window
}

function showOpenDialogFor(event, options) {
  const owner = callingWindow(event)
  return owner ? dialog.showOpenDialog(owner, options) : dialog.showOpenDialog(options)
}

function registerIpc() {
  ipcMain.handle('file:open-dialog', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Open a video',
      properties: ['openFile'],
      filters: [
        { name: 'Video files', extensions: SUPPORTED_EXTENSIONS.map((extension) => extension.slice(1)) },
        { name: 'MP4 video', extensions: ['mp4', 'm4v'] },
        { name: 'WebM video', extensions: ['webm'] },
        { name: 'Ogg video', extensions: ['ogv'] },
      ],
    })
    if (result.canceled || !result.filePaths[0]) return null
    return videoPayload(result.filePaths[0])
  })

  ipcMain.handle('file:open-path', (_event, filePath) => videoPayload(filePath))
  ipcMain.handle('file:open-in-new-window', async (_event, filePath) => {
    if (filePath) await videoPayload(filePath)
    createWindow(filePath || null)
    return true
  })
  ipcMain.handle('app:new-window', () => {
    createWindow()
    return true
  })

  ipcMain.handle('recent:list', async () => {
    await recentMutations
    return readRecents()
  })
  ipcMain.handle('recent:remove', async (_event, filePath) => {
    const resolved = path.resolve(String(filePath || ''))
    return mutateRecents((current) => current.filter((item) => item.path.toLowerCase() !== resolved.toLowerCase()))
  })
  ipcMain.handle('recent:clear', () => mutateRecents(() => []))
  ipcMain.handle('shell:show-item', (_event, filePath) => {
    if (typeof filePath !== 'string' || !isSupportedVideoPath(filePath)) return false
    shell.showItemInFolder(path.resolve(filePath))
    return true
  })

  ipcMain.on('window:set-title', (event, title) => {
    const window = callingWindow(event)
    if (window) window.setTitle(String(title || 'simple video').slice(0, 256))
  })
  ipcMain.on('window:minimize', (event) => callingWindow(event)?.minimize())
  ipcMain.on('window:toggle-maximize', (event) => {
    const window = callingWindow(event)
    if (!window) return
    if (window.isMaximized()) window.unmaximize()
    else window.maximize()
  })
  ipcMain.on('window:toggle-fullscreen', (event) => {
    const window = callingWindow(event)
    if (window) window.setFullScreen(!window.isFullScreen())
  })
  ipcMain.on('window:exit-fullscreen', (event) => callingWindow(event)?.setFullScreen(false))
  ipcMain.on('window:close', (event) => callingWindow(event)?.close())
}

function focusExistingWindow() {
  const window = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0]
  if (!window) return
  if (window.isMinimized()) window.restore()
  window.focus()
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (_event, argv, workingDirectory) => {
    const incoming = supportedPaths(argv, workingDirectory)
    if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
    else focusExistingWindow()
  })

  app.on('open-file', (event, filePath) => {
    event.preventDefault()
    if (!isSupportedVideoPath(filePath)) return
    if (ready) createWindow(path.resolve(filePath))
    else pendingMacPaths.push(path.resolve(filePath))
  })

  app.whenReady().then(() => {
    ready = true
    app.setAppUserModelId('com.simple.video')
    session.defaultSession.setPermissionRequestHandler((_webContents, _permission, callback) => callback(false))
    registerIpc()

    const incoming = [...pendingMacPaths, ...supportedPaths(process.argv)]
    pendingMacPaths.length = 0
    if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
    else createWindow()

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
  })
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})
