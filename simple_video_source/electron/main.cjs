'use strict'

const { app, BrowserWindow, clipboard, dialog, ipcMain, nativeImage, session, shell } = require('electron')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const { SUPPORTED_EXTENSIONS, isSupportedVideoPath, supportedPaths } = require('./routing.cjs')
const { mergeRecent, sanitizeRecents } = require('./recent-files.cjs')
const {
  FRAME_FORMATS,
  assertCopyTarget,
  assertFrameTarget,
  suggestedCopyName,
  suggestedFrameName,
  validateFrameBytes,
} = require('./video-export.cjs')
const {
  createVideoPrintDocument,
  ensurePrinterAvailable,
  pngDimensions,
  runSilentPrintJob,
} = require('./video-print.cjs')

const MAX_VIDEO_BYTES = 1024 * 1024 * 1024 * 1024
const RECENT_FILE = 'video-recents.json'
let ready = false
const pendingMacPaths = []
let recentMutations = Promise.resolve()
const printSessions = new Map()
const PRINT_SESSION_MAX_AGE = 15 * 60 * 1000

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

async function checkedVideo(filePath) {
  if (typeof filePath !== 'string' || !isSupportedVideoPath(filePath)) {
    throw new Error('Simple Video does not support this file type.')
  }
  const resolved = path.resolve(filePath)
  const stat = await fs.stat(resolved)
  if (!stat.isFile()) throw new Error('The selected item is not a video file.')
  if (stat.size <= 0) throw new Error('This video file is empty.')
  if (stat.size > MAX_VIDEO_BYTES) throw new Error('This video file is too large to open safely.')
  return { resolved, stat }
}

async function videoPayload(filePath) {
  const { resolved, stat } = await checkedVideo(filePath)
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
  const ownerId = window.webContents.id
  window.webContents.once('destroyed', () => clearPrintSessionsForOwner(ownerId))
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

async function showSaveDialogFor(event, options) {
  const qaDirectory = process.env.SIMPLE_VIDEO_QA_EXPORT_DIRECTORY
  if (!app.isPackaged && qaDirectory && options?.defaultPath) {
    const directory = path.resolve(qaDirectory)
    await fs.mkdir(directory, { recursive: true })
    return { canceled: false, filePath: path.join(directory, path.basename(options.defaultPath)) }
  }
  const owner = callingWindow(event)
  return owner ? dialog.showSaveDialog(owner, options) : dialog.showSaveDialog(options)
}

async function replaceFileAtomically(targetPath, writeTemporary) {
  const token = crypto.randomUUID()
  const temporary = path.join(path.dirname(targetPath), `.${path.basename(targetPath)}.${token}.tmp`)
  const backup = path.join(path.dirname(targetPath), `.${path.basename(targetPath)}.${token}.bak`)
  let backedUp = false
  let committed = false

  await writeTemporary(temporary)
  try {
    try {
      await fs.rename(targetPath, backup)
      backedUp = true
    } catch (error) {
      if (error.code !== 'ENOENT') throw error
    }
    await fs.rename(temporary, targetPath)
    committed = true
    if (backedUp) await fs.rm(backup, { force: true }).catch(() => {})
  } catch (error) {
    await fs.rm(temporary, { force: true }).catch(() => {})
    if (backedUp && !committed) await fs.rename(backup, targetPath).catch(() => {})
    throw error
  }
}

function clearPrintSessionsForOwner(ownerId) {
  for (const [sessionId, value] of printSessions) {
    if (value.ownerId === ownerId) printSessions.delete(sessionId)
  }
}

function expirePrintSessions() {
  const cutoff = Date.now() - PRINT_SESSION_MAX_AGE
  for (const [sessionId, value] of printSessions) {
    if (value.createdAt < cutoff) printSessions.delete(sessionId)
  }
}

function printSessionFor(event, sessionId) {
  expirePrintSessions()
  if (typeof sessionId !== 'string' || sessionId.length > 100) throw new Error('The print session is invalid.')
  const value = printSessions.get(sessionId)
  if (!value || value.ownerId !== event.sender.id) throw new Error('This captured frame is no longer available. Open Print again.')
  value.createdAt = Date.now()
  return value
}

function publicPrintPreview(printDocument) {
  return {
    html: printDocument.html,
    title: printDocument.title,
    page: printDocument.page,
    frame: printDocument.frame,
    placement: printDocument.placement,
    options: printDocument.options,
  }
}

async function printVideoFrame(printDocument, owner) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-video-print-'))
  const htmlPath = path.join(directory, 'video-frame.html')
  let printWindow = null
  try {
    await fs.writeFile(htmlPath, printDocument.html, 'utf8')
    printWindow = new BrowserWindow({
      ...(owner && !owner.isDestroyed() ? { parent: owner } : {}),
      show: false,
      title: `Printing — ${printDocument.title}`,
      backgroundColor: '#ffffff',
      webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true, devTools: false },
    })
    printWindow.removeMenu()
    await printWindow.loadFile(htmlPath)
    const decodedFrame = await printWindow.webContents.executeJavaScript(`(() => {
      const image = document.querySelector('.frame-slot img')
      if (!image) throw new Error('The printable video frame is missing.')
      return image.decode().then(() => new Promise((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve({
          width: image.naturalWidth,
          height: image.naturalHeight,
        })))
      }))
    })()`)
    if (decodedFrame?.width !== printDocument.frame.width || decodedFrame?.height !== printDocument.frame.height) {
      throw new Error('The printable video frame did not decode at the expected size.')
    }

    const qaDirectory = process.env.SIMPLE_VIDEO_QA_PRINT_DIRECTORY
    const qaDelay = !app.isPackaged ? Math.max(0, Math.min(2_000, Number(process.env.SIMPLE_VIDEO_QA_PRINT_DELAY_MS) || 0)) : 0
    if (qaDelay) await new Promise((resolve) => setTimeout(resolve, qaDelay))
    if (!app.isPackaged && qaDirectory) {
      const targetDirectory = path.resolve(qaDirectory)
      await fs.mkdir(targetDirectory, { recursive: true })
      const pdf = await printWindow.webContents.printToPDF({
        printBackground: true,
        landscape: printDocument.options.orientation === 'landscape',
        preferCSSPageSize: true,
        generateTaggedPDF: true,
      })
      const targetPath = path.join(targetDirectory, 'video-frame-print.pdf')
      await fs.writeFile(targetPath, pdf)
      return { printed: true, canceled: false, path: targetPath }
    }

    await ensurePrinterAvailable(printWindow.webContents)
    const outcome = await new Promise((resolve, reject) => {
      let settled = false
      const finish = (callback) => {
        if (settled) return
        settled = true
        callback()
      }
      printWindow.once('closed', () => finish(() => reject(new Error('The print window closed before the job was started.'))))
      runSilentPrintJob(printWindow.webContents, printDocument)
        .then((result) => finish(() => resolve(result)))
        .catch((error) => finish(() => reject(error)))
    })
    return outcome
  } finally {
    if (printWindow && !printWindow.isDestroyed()) printWindow.destroy()
    await fs.rm(directory, { recursive: true, force: true }).catch(() => {})
  }
}

async function cleanupStalePrintDirectories() {
  let entries = []
  try { entries = await fs.readdir(os.tmpdir(), { withFileTypes: true }) } catch { return }
  const cutoff = Date.now() - 24 * 60 * 60 * 1_000
  await Promise.all(entries
    .filter((entry) => entry.isDirectory() && entry.name.startsWith('simple-video-print-'))
    .map(async (entry) => {
      const target = path.join(os.tmpdir(), entry.name)
      try {
        const stat = await fs.stat(target)
        if (stat.mtimeMs < cutoff) await fs.rm(target, { recursive: true, force: true })
      } catch {}
    }))
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

  ipcMain.handle('export:copy-original', async (event, filePath) => {
    const { resolved } = await checkedVideo(filePath)
    const extension = path.extname(resolved)
    const result = await showSaveDialogFor(event, {
      title: 'Export As — Original video copy',
      defaultPath: path.join(path.dirname(resolved), suggestedCopyName(resolved)),
      filters: [{ name: `${extension.slice(1).toUpperCase()} video`, extensions: [extension.slice(1)] }],
    })
    if (result.canceled || !result.filePath) return { canceled: true }
    const targetPath = path.resolve(result.filePath)
    assertCopyTarget(resolved, targetPath)
    await replaceFileAtomically(targetPath, (temporary) => fs.copyFile(resolved, temporary))
    return { canceled: false, path: targetPath, name: path.basename(targetPath) }
  })

  ipcMain.handle('export:frame', async (event, payload) => {
    if (!payload || typeof payload !== 'object') throw new Error('The video frame export request is invalid.')
    const { resolved } = await checkedVideo(payload.sourcePath)
    const format = String(payload.format || '').toLowerCase()
    const definition = FRAME_FORMATS[format]
    if (!definition) throw new Error('Choose PNG or JPEG for a video frame.')
    const bytes = validateFrameBytes(payload.bytes, format)
    const result = await showSaveDialogFor(event, {
      title: `Export As — ${definition.label}`,
      defaultPath: path.join(path.dirname(resolved), suggestedFrameName(resolved, format, Number(payload.seconds))),
      filters: [{ name: definition.label, extensions: format === 'jpeg' ? ['jpg', 'jpeg'] : ['png'] }],
    })
    if (result.canceled || !result.filePath) return { canceled: true }
    const targetPath = path.resolve(result.filePath)
    assertFrameTarget(targetPath, format)
    await replaceFileAtomically(targetPath, (temporary) => fs.writeFile(temporary, bytes))
    return { canceled: false, path: targetPath, name: path.basename(targetPath) }
  })

  ipcMain.handle('clipboard:write-frame', (event, value) => {
    if (!callingWindow(event)) throw new Error('This clipboard request did not come from the video workspace.')
    const frame = pngDimensions(value)
    const image = nativeImage.createFromBuffer(frame.bytes)
    if (image.isEmpty()) throw new Error('The copied video frame could not be decoded.')
    clipboard.writeImage(image)
    return { width: frame.width, height: frame.height }
  })

  ipcMain.handle('print:start', async (event, payload) => {
    if (!payload || typeof payload !== 'object') throw new Error('The video print request is invalid.')
    const { resolved } = await checkedVideo(payload.sourcePath)
    const frame = pngDimensions(payload.bytes)
    clearPrintSessionsForOwner(event.sender.id)
    const sessionId = crypto.randomUUID()
    const seconds = Math.max(0, Number(payload.seconds) || 0)
    printSessions.set(sessionId, {
      ownerId: event.sender.id,
      createdAt: Date.now(),
      bytes: frame.bytes,
      sourceName: path.basename(resolved),
      seconds,
    })
    return { sessionId, width: frame.width, height: frame.height, seconds, sourceName: path.basename(resolved) }
  })

  ipcMain.handle('print:preview', (event, payload) => {
    if (!payload || typeof payload !== 'object') throw new Error('The print preview request is invalid.')
    const captured = printSessionFor(event, payload.sessionId)
    return publicPrintPreview(createVideoPrintDocument({ ...captured, options: payload.options }))
  })

  ipcMain.handle('print:run', async (event, payload) => {
    if (!payload || typeof payload !== 'object') throw new Error('The print request is invalid.')
    const captured = printSessionFor(event, payload.sessionId)
    const printDocument = createVideoPrintDocument({ ...captured, options: payload.options })
    return printVideoFrame(printDocument, callingWindow(event))
  })

  ipcMain.handle('print:end', (event, sessionId) => {
    const captured = typeof sessionId === 'string' ? printSessions.get(sessionId) : null
    if (captured?.ownerId === event.sender.id) printSessions.delete(sessionId)
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

  app.whenReady().then(async () => {
    ready = true
    app.setAppUserModelId('com.simple.video')
    session.defaultSession.setPermissionRequestHandler((_webContents, _permission, callback) => callback(false))
    await cleanupStalePrintDirectories()
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
