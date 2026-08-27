const { app, BrowserWindow, dialog, ipcMain } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const { pathToFileURL } = require('node:url')
const {
  EDITABLE_EXTENSIONS,
  MAX_IMAGE_BYTES,
  MIME_BY_EXTENSION,
  SUPPORTED_EXTENSIONS,
  atomicWrite,
  ensureOutputExtension,
  extensionOf,
  isSupportedExtension,
  outputExtension,
  safeStem,
  toBytes,
  validateImageBytes,
  validateImageDimensions,
} = require('./image-files.cjs')
const { ensurePdfExtension, imageToPdfBytes } = require('./pdf-export.cjs')

const closeApprovedWindows = new WeakSet()
const productionRendererUrl = pathToFileURL(path.join(__dirname, '..', 'dist', 'index.html')).href

function developmentRendererUrl() {
  if (app.isPackaged || !process.env.VITE_DEV_SERVER_URL) return null
  const parsed = new URL(process.env.VITE_DEV_SERVER_URL)
  if (parsed.protocol !== 'http:' || !['127.0.0.1', 'localhost', '[::1]'].includes(parsed.hostname)) {
    throw new Error('The image development server must use local HTTP.')
  }
  return parsed.href
}

function isTrustedRenderer(event) {
  const senderUrl = event.senderFrame?.url || event.sender.getURL()
  if (senderUrl === productionRendererUrl) return true
  const developmentUrl = developmentRendererUrl()
  if (!developmentUrl) return false
  try { return new URL(senderUrl).origin === new URL(developmentUrl).origin } catch { return false }
}

function callingWindow(event) {
  if (!isTrustedRenderer(event)) return null
  const owner = BrowserWindow.fromWebContents(event.sender)
  return owner && !owner.isDestroyed() ? owner : null
}

function requireCallingWindow(event) {
  const owner = callingWindow(event)
  if (!owner) throw new Error('This request did not come from the image workspace.')
  return owner
}

function showOpenDialogFor(event, options) {
  return dialog.showOpenDialog(requireCallingWindow(event), options)
}

function showSaveDialogFor(event, options) {
  return dialog.showSaveDialog(requireCallingWindow(event), options)
}

function fileFilters() {
  return [
    { name: 'All supported images', extensions: SUPPORTED_EXTENSIONS.map((extension) => extension.slice(1)) },
    { name: 'PNG image', extensions: ['png'] },
    { name: 'JPEG image', extensions: ['jpg', 'jpeg'] },
    { name: 'WebP image', extensions: ['webp'] },
    { name: 'GIF image', extensions: ['gif'] },
    { name: 'Bitmap image', extensions: ['bmp'] },
    { name: 'SVG image', extensions: ['svg'] },
    { name: 'AVIF image', extensions: ['avif'] },
  ]
}

async function filePayload(filePath) {
  if (typeof filePath !== 'string' || !filePath || filePath.length > 32_768) throw new Error('The image path is invalid.')
  const resolved = path.resolve(filePath)
  const extension = extensionOf(resolved)
  if (!SUPPORTED_EXTENSIONS.includes(extension)) throw new Error('This image format is not supported.')
  const stat = await fs.stat(resolved)
  if (!stat.isFile()) throw new Error('Choose an image file.')
  if (!stat.size || stat.size > MAX_IMAGE_BYTES) throw new Error('This image is empty or larger than the 512 MB safety limit.')
  const bytes = validateImageBytes(await fs.readFile(resolved), extension)
  validateImageDimensions(bytes, extension)
  return {
    data: new Uint8Array(bytes),
    name: path.basename(resolved),
    path: resolved,
    size: bytes.byteLength,
    format: extension.slice(1),
    mime: MIME_BY_EXTENSION[extension],
    directSave: EDITABLE_EXTENSIONS.includes(extension),
  }
}

async function bytesPayload(input) {
  const name = path.basename(String(input?.name || 'Dropped image'))
  const extension = extensionOf(name)
  const bytes = validateImageBytes(input?.data, extension)
  validateImageDimensions(bytes, extension)
  return {
    data: new Uint8Array(bytes),
    name,
    path: null,
    size: bytes.byteLength,
    format: extension.slice(1),
    mime: MIME_BY_EXTENSION[extension],
    directSave: false,
  }
}

async function printImage(input, owner) {
  const png = validateImageBytes(input?.data, '.png')
  validateImageDimensions(png, '.png')
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-print-'))
  const imagePath = path.join(directory, `${safeStem(input?.name)}.png`)
  const htmlPath = path.join(directory, 'print.html')
  let printWindow = null
  const cleanup = async () => {
    if (printWindow && !printWindow.isDestroyed()) printWindow.destroy()
    await fs.rm(directory, { recursive: true, force: true }).catch(() => {})
  }
  try {
    await fs.writeFile(imagePath, png)
    const imageUrl = pathToFileURL(imagePath).href.replace(/&/g, '&amp;').replace(/"/g, '&quot;')
    const html = `<!doctype html><html><head><meta charset="utf-8"><style>
      @page { margin: 12mm; }
      html, body { width: 100%; height: 100%; margin: 0; }
      body { display: grid; place-items: center; }
      img { display: block; max-width: 100%; max-height: 100%; object-fit: contain; }
    </style></head><body><img src="${imageUrl}" alt=""></body></html>`
    await fs.writeFile(htmlPath, html, 'utf8')
    printWindow = new BrowserWindow({
      ...(owner ? { parent: owner } : {}),
      width: 800,
      height: 700,
      show: false,
      title: `Print — ${safeStem(input?.name)}`,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: true,
        javascript: false,
        devTools: false,
      },
    })
    printWindow.removeMenu()
    printWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
    printWindow.webContents.on('will-navigate', (event, url) => {
      if (url !== pathToFileURL(htmlPath).href) event.preventDefault()
    })
    await printWindow.loadFile(htmlPath)
    return await new Promise((resolve, reject) => {
      printWindow.webContents.print({ silent: false, printBackground: true }, async (success, reason) => {
        await cleanup()
        if (success) resolve(true)
        else if (/cancel/i.test(String(reason))) resolve(false)
        else reject(new Error(reason || 'Windows could not start printing.'))
      })
    })
  } catch (error) {
    await cleanup()
    throw error
  }
}

async function cleanupStalePrintDirectories() {
  let entries = []
  try { entries = await fs.readdir(os.tmpdir(), { withFileTypes: true }) } catch { return }
  await Promise.all(entries
    .filter((entry) => entry.isDirectory() && entry.name.startsWith('simple-image-print-'))
    .map((entry) => fs.rm(path.join(os.tmpdir(), entry.name), { recursive: true, force: true }).catch(() => {})))
}

function createWindow(openPath = null) {
  const developmentUrl = developmentRendererUrl()
  const rendererUrl = developmentUrl || productionRendererUrl
  const window = new BrowserWindow({
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    width: 1480,
    height: 940,
    minWidth: 860,
    minHeight: 600,
    show: false,
    frame: false,
    backgroundColor: '#e9e9e9',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: !app.isPackaged,
    },
  })
  window.removeMenu()
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  window.webContents.on('will-navigate', (event, url) => {
    if (url !== rendererUrl) event.preventDefault()
  })
  if (developmentUrl) window.loadURL(developmentUrl)
  else window.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))
  window.once('ready-to-show', () => window.show())
  window.webContents.once('did-finish-load', () => {
    if (openPath) window.webContents.send('file:open-external', openPath)
  })
  window.on('maximize', () => window.webContents.send('window:maximized', true))
  window.on('unmaximize', () => window.webContents.send('window:maximized', false))
  window.on('close', (event) => {
    if (closeApprovedWindows.has(window)) return
    event.preventDefault()
    window.webContents.send('window:close-requested')
  })
  return window
}

function supportedPaths(argv) {
  return [...new Set(argv.filter((argument) => typeof argument === 'string' && isSupportedExtension(argument)))]
}

function registerIpc() {
  ipcMain.handle('file:open-dialog', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Open an image',
      properties: ['openFile'],
      filters: fileFilters(),
    })
    if (result.canceled || !result.filePaths[0]) return null
    return filePayload(result.filePaths[0])
  })
  ipcMain.handle('file:open-path', (event, filePath) => {
    requireCallingWindow(event)
    return filePayload(filePath)
  })
  ipcMain.handle('file:open-bytes', (event, input) => {
    requireCallingWindow(event)
    return bytesPayload(input)
  })
  ipcMain.handle('file:open-in-new-window', async (event, requestedPath = null) => {
    requireCallingWindow(event)
    let selectedPath = requestedPath
    if (!selectedPath) {
      const result = await showOpenDialogFor(event, { title: 'Open an image', properties: ['openFile'], filters: fileFilters() })
      if (result.canceled || !result.filePaths[0]) return false
      selectedPath = result.filePaths[0]
    }
    if (!isSupportedExtension(selectedPath)) throw new Error('This image format is not supported.')
    createWindow(selectedPath)
    return true
  })
  ipcMain.handle('file:save', async (event, input) => {
    requireCallingWindow(event)
    const format = String(input?.format || '').toLowerCase()
    const extension = outputExtension(format)
    const bytes = validateImageBytes(input?.data, extension)
    validateImageDimensions(bytes, extension)
    let targetPath = input?.forceDialog ? null : input?.path
    if (targetPath) {
      const targetExtension = extensionOf(targetPath)
      const compatible = targetExtension === extension || (extension === '.jpg' && targetExtension === '.jpeg')
      if (!compatible) targetPath = null
    }
    if (!targetPath) {
      const label = format === 'jpeg' || format === 'jpg' ? 'JPEG' : format.toUpperCase()
      const result = await showSaveDialogFor(event, {
        title: input?.forceDialog ? 'Save image as' : 'Save image',
        defaultPath: `${safeStem(input?.name)}${extension}`,
        filters: [{ name: `${label} image`, extensions: extension === '.jpg' ? ['jpg', 'jpeg'] : [extension.slice(1)] }],
      })
      if (result.canceled || !result.filePath) return null
      targetPath = ensureOutputExtension(result.filePath, format)
    }
    await atomicWrite(targetPath, bytes)
    return {
      path: targetPath,
      name: path.basename(targetPath),
      size: bytes.byteLength,
      format: extensionOf(targetPath).slice(1),
    }
  })
  ipcMain.handle('image:convert-to-pdf', async (event, input) => {
    requireCallingWindow(event)
    const png = validateImageBytes(input?.data, '.png')
    validateImageDimensions(png, '.png')
    const result = await showSaveDialogFor(event, {
      title: 'Convert image to PDF',
      defaultPath: `${safeStem(input?.name)}.pdf`,
      filters: [{ name: 'PDF document', extensions: ['pdf'] }],
    })
    if (result.canceled || !result.filePath) return null
    const targetPath = ensurePdfExtension(result.filePath)
    const pdf = await imageToPdfBytes(png, safeStem(input?.name))
    await atomicWrite(targetPath, pdf)
    return {
      path: targetPath,
      name: path.basename(targetPath),
      size: pdf.byteLength,
    }
  })
  ipcMain.handle('image:print', (event, input) => printImage(input, requireCallingWindow(event)))
  ipcMain.handle('app:new-window', (event) => { requireCallingWindow(event); createWindow(); return true })
  ipcMain.handle('app:get-version', (event) => { requireCallingWindow(event); return app.getVersion() })
  ipcMain.on('window:set-title', (event, title) => {
    const window = callingWindow(event)
    if (window) window.setTitle(`${String(title || 'Untitled image')} — Simple Image`)
  })
  ipcMain.on('window:minimize', (event) => callingWindow(event)?.minimize())
  ipcMain.on('window:toggle-maximize', (event) => {
    const window = callingWindow(event)
    if (!window) return
    if (window.isMaximized()) window.unmaximize()
    else window.maximize()
  })
  ipcMain.on('window:confirm-close', (event) => {
    const window = callingWindow(event)
    if (!window) return
    closeApprovedWindows.add(window)
    window.close()
  })
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (_event, argv) => {
    const incoming = supportedPaths(argv)
    if (incoming.length) {
      for (const filePath of incoming) createWindow(filePath)
      return
    }
    const window = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0]
    if (window) {
      if (window.isMinimized()) window.restore()
      window.focus()
    }
  })
  app.whenReady().then(async () => {
    registerIpc()
    await cleanupStalePrintDirectories()
    const incoming = supportedPaths(process.argv)
    if (incoming.length) for (const filePath of incoming) createWindow(filePath)
    else createWindow()
    app.on('activate', () => {
      if (!BrowserWindow.getAllWindows().length) createWindow()
    })
  })
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', (event) => {
  const windowsNeedingApproval = BrowserWindow.getAllWindows().filter((window) => !closeApprovedWindows.has(window))
  if (!windowsNeedingApproval.length) return
  event.preventDefault()
  for (const window of windowsNeedingApproval) window.webContents.send('window:close-requested')
})
