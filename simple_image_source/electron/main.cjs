const { app, BrowserWindow, clipboard, dialog, ipcMain, nativeImage } = require('electron')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const {
  MAX_IMAGE_BYTES,
  SUPPORTED_EXTENSIONS,
  detectAnimation,
  ensureOutputExtension,
  extensionOf,
  formatOfExtension,
  inspectImageBytes,
  isCompatibleOutputExtension,
  isSupportedExtension,
  nameWithFormat,
  outputExtension,
  safeStem,
  toBytes,
  validateImageBytes,
  validateImageDimensions,
} = require('./image-files.cjs')
const { extractMetadata, fidelityWarnings, metadataSize, spliceMetadata } = require('./image-metadata.cjs')
const { ensurePdfExtension, imageToPdfBytes } = require('./pdf-export.cjs')
const { buildPrintHtml, computePrintLayout, electronPrintOptions } = require('./print-layout.cjs')
const { cleanupStalePrintDirectories, createOwnedPrintDirectory, removePrintDirectory } = require('./print-temp.cjs')
const { ensurePrinterInstalled, submitPrintJob } = require('./default-printer.cjs')
const { preparePrintableImage } = require('./print-image.cjs')
const { registerSharedIo, bridgeArguments } = require('./simple-io/io-ipc.cjs')
const { sweep } = require('./simple-io/io-core.cjs')
const { safeWriteFile } = require('./simple-io/safe-write.cjs')
const guard = require('./simple-io/document-guard.cjs')
const stores = require('./simple-io/stores.cjs')

const IO_MODULE = 'image'

// Every saved or exported image goes through the shared verified write: temp file in
// the same folder, flush, read-back check, then replace with retries while another
// program (a viewer, a backup tool, an antivirus scan) holds the file.
async function atomicWrite(targetPath, value) {
  await safeWriteFile(targetPath, toBytes(value))
}

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
    { name: 'PNG image', extensions: ['png', 'apng'] },
    { name: 'JPEG image', extensions: ['jpg', 'jpeg', 'jfif', 'jpe', 'jif'] },
    { name: 'WebP image', extensions: ['webp'] },
    { name: 'GIF image', extensions: ['gif'] },
    { name: 'Bitmap image', extensions: ['bmp'] },
    { name: 'SVG image', extensions: ['svg', 'svgz'] },
    { name: 'AVIF image', extensions: ['avif'] },
    { name: 'Icon', extensions: ['ico'] },
    { name: 'Photoshop document', extensions: ['psd'] },
  ]
}

// #region opened sources (metadata carried into edited saves)
// Each open remembers what an edited save needs from the original file (its EXIF/XMP/ICC/... blocks, bit
// depth, animation, folder) under a random token that the payload returns as `sourceToken`. The renderer
// sends the token back with file:save. Only a few recent sources per window are kept, and they are released
// when the window's renderer goes away.
const SOURCES_PER_WINDOW = 4
const MAX_REMEMBERED_METADATA_BYTES = 32 * 1024 * 1024
const openedSources = new Map()

function digestOf(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex')
}

function rememberSource(sender, entry) {
  if (!sender || typeof sender.id !== 'number') return null
  let sources = openedSources.get(sender.id)
  if (!sources) {
    sources = new Map()
    openedSources.set(sender.id, sources)
    const id = sender.id
    sender.once?.('destroyed', () => openedSources.delete(id))
  }
  const token = crypto.randomUUID()
  sources.set(token, entry)
  while (sources.size > SOURCES_PER_WINDOW) sources.delete(sources.keys().next().value)
  return token
}

function rememberedSource(sender, token) {
  if (!sender || typeof token !== 'string') return null
  return openedSources.get(sender.id)?.get(token) || null
}

function metadataFormat(format) {
  return ['jpeg', 'png', 'webp', 'avif', 'psd'].includes(format)
}

function openPayload(inspected, original, location, sender) {
  const metadata = metadataFormat(inspected.format) ? extractMetadata(original) : null
  const orientation = metadata?.orientation ?? null
  // JPEG frame headers use sensor axes; report the upright size the renderer will decode.
  const swap = inspected.format === 'jpeg' && orientation >= 5 && orientation <= 8
  const sourceToken = rememberSource(sender, {
    path: location.path,
    name: location.name,
    format: inspected.format,
    animated: inspected.animated,
    digest: digestOf(original),
    metadata: metadata && metadataSize(metadata) <= MAX_REMEMBERED_METADATA_BYTES ? metadata : null,
  })
  return {
    data: new Uint8Array(inspected.bytes),
    name: location.name,
    path: location.path,
    size: original.byteLength,
    format: inspected.payloadFormat,
    mime: inspected.mime,
    directSave: Boolean(location.path) && inspected.canSaveInPlace,
    /** Extension of the file name without the dot ('' when there is none). */
    extension: inspected.extension.slice(1),
    /** The content is a different format than the extension says (e.g. a WebP named photo.jpg). */
    formatMismatch: inspected.mismatch,
    animated: inspected.animated,
    /** Frames in the source animation (1 for a still image, null when unknown). */
    frameCount: inspected.frameCount,
    /** Size the renderer should decode/draw at (SVG: the raster size; JPEG: after EXIF orientation). */
    width: swap ? inspected.height : inspected.width,
    height: swap ? inspected.width : inspected.height,
    /** SVG only: the drawing's own size in CSS pixels before rasterizing. */
    intrinsicSize: inspected.intrinsic,
    bitDepth: metadata?.bitDepth ?? 8,
    colorModel: metadata?.colorModel ?? 'rgb',
    orientation,
    /** User-facing notes about how this file opened (format mismatch, animation, SVG raster size). */
    notices: inspected.notices,
    /** File name Save As should propose when the file cannot be saved in place, else null. */
    suggestedName: inspected.suggestedName,
    sourceToken,
  }
}

async function filePayload(filePath, sender = null) {
  if (typeof filePath !== 'string' || !filePath || filePath.length > 32_768) throw new Error('The image path is invalid.')
  const resolved = path.resolve(filePath)
  const stat = await fs.stat(resolved)
  if (!stat.isFile()) throw new Error('Choose an image file.')
  if (!stat.size || stat.size > MAX_IMAGE_BYTES) throw new Error('This image is empty or larger than the 512 MB safety limit.')
  const original = await fs.readFile(resolved)
  // The content decides the format: a WebP saved as .jpg opens as WebP (and is not overwritten in place).
  const inspected = inspectImageBytes(original, resolved)
  return openPayload(inspected, original, { name: path.basename(resolved), path: resolved }, sender)
}

async function bytesPayload(input, sender = null) {
  const name = path.basename(String(input?.name || 'Dropped image'))
  const original = toBytes(input?.data)
  const inspected = inspectImageBytes(original, name)
  return openPayload(inspected, original, { name, path: null }, sender)
}

async function readExistingImage(filePath) {
  try {
    const stat = await fs.stat(filePath)
    if (!stat.isFile() || !stat.size || stat.size > MAX_IMAGE_BYTES) return null
    return await fs.readFile(filePath)
  } catch {
    return null
  }
}

async function existingDirectory(directory) {
  try { return (await fs.stat(directory)).isDirectory() } catch { return false }
}
// #endregion opened sources

// #region clipboard image (paste)
function clipboardHasImage() {
  let formats = []
  try { formats = clipboard.availableFormats() } catch { return false }
  return Array.isArray(formats) && formats.some((format) => typeof format === 'string' && format.toLowerCase().startsWith('image/'))
}

/** The clipboard image as PNG bytes checked like any opened image (signature, 50 MP / 20,000 px), or null. */
function readClipboardPng() {
  const image = clipboard.readImage()
  if (!image || image.isEmpty()) return null
  const png = image.toPNG()
  if (!png || !png.length) return null
  validateImageBytes(png, '.png')
  validateImageDimensions(png, '.png')
  // A copy: a small Buffer can be a slice of a shared pool, and IPC would carry the whole pool.
  return new Uint8Array(png)
}
// #endregion clipboard image

async function printImage(input, owner) {
  const { bytes, extension, dimensions } = preparePrintableImage(input)
  const layout = computePrintLayout(dimensions.width, dimensions.height, input?.settings)
  const directory = await createOwnedPrintDirectory()
  const imagePath = path.join(directory, `${safeStem(input?.name)}${extension}`)
  const htmlPath = path.join(directory, 'print.html')
  let printWindow = null
  const cleanup = async () => {
    if (printWindow && !printWindow.isDestroyed()) printWindow.destroy()
    await removePrintDirectory(directory).catch(() => {})
  }
  try {
    await fs.writeFile(imagePath, bytes)
    const html = buildPrintHtml(pathToFileURL(imagePath).href, dimensions.width, dimensions.height, layout.settings, safeStem(input?.name))
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
        // The page contains no scripts, but JavaScript must remain enabled so the
        // main process can await img.decode() before handing the page to Chromium.
        javascript: true,
        devTools: false,
      },
    })
    closeApprovedWindows.add(printWindow)
    printWindow.removeMenu()
    printWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
    printWindow.webContents.on('will-navigate', (event, url) => {
      if (url !== pathToFileURL(htmlPath).href) event.preventDefault()
    })
    await printWindow.loadFile(htmlPath)
    const decoded = await printWindow.webContents.executeJavaScript(`(() => {
      const image = document.querySelector('img')
      if (!image) throw new Error('The printable image is missing.')
      return image.decode().then(() => ({ width: image.naturalWidth, height: image.naturalHeight }))
    })()`)
    if (decoded?.width !== dimensions.width || decoded?.height !== dimensions.height) {
      throw new Error('The printable image did not decode at the expected size.')
    }
    await ensurePrinterInstalled(printWindow.webContents)
    return await submitPrintJob(printWindow.webContents, electronPrintOptions(layout.settings))
  } finally {
    await cleanup()
  }
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
      additionalArguments: bridgeArguments(IO_MODULE),
    },
  })
  // Save / Don't Save / Cancel on Alt+F4, taskbar close, quit and Windows sign-out,
  // never closing during a save, and crash/hang handling.
  guard.installWindowGuard(window)
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
    return filePayload(result.filePaths[0], event.sender)
  })
  ipcMain.handle('file:open-path', (event, filePath) => {
    requireCallingWindow(event)
    return filePayload(filePath, event.sender)
  })
  ipcMain.handle('file:list-siblings', async (event, filePath) => {
    requireCallingWindow(event)
    if (typeof filePath !== 'string' || !filePath || filePath.length > 32_768) return []
    const directory = path.dirname(path.resolve(filePath))
    let entries
    try { entries = await fs.readdir(directory, { withFileTypes: true }) } catch { return [] }
    return entries
      .filter((entry) => entry.isFile() && !entry.name.startsWith('.') && isSupportedExtension(entry.name))
      .map((entry) => path.join(directory, entry.name))
      .sort((a, b) => path.basename(a).localeCompare(path.basename(b), undefined, { numeric: true, sensitivity: 'base' }))
  })
  ipcMain.handle('file:open-bytes', (event, input) => {
    requireCallingWindow(event)
    return bytesPayload(input, event.sender)
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
    const contentFormat = formatOfExtension(extension)
    let bytes = validateImageBytes(input?.data, extension)
    validateImageDimensions(bytes, extension)
    const source = rememberedSource(event.sender, input?.sourceToken)
    let targetPath = input?.forceDialog ? null : input?.path
    if (targetPath && (typeof targetPath !== 'string' || targetPath.length > 32_768 || !isCompatibleOutputExtension(extensionOf(targetPath), format))) {
      targetPath = null
    }
    // The file being replaced in place: its bytes tell whether this save is unchanged, whether it would
    // flatten an animation, and (without a sourceToken) where the metadata to keep comes from.
    let replaced = targetPath && contentFormat !== 'psd' ? await readExistingImage(targetPath) : null
    let stillOfAnimation = Boolean(source?.animated) && !detectAnimation(bytes, contentFormat).animated
    if (replaced && contentFormat !== 'psd' && detectAnimation(replaced).animated && !detectAnimation(bytes, contentFormat).animated) {
      // Never overwrite an animated PNG/WebP/GIF with one frame: ask where the still copy goes instead.
      stillOfAnimation = true
      targetPath = null
      replaced = null
    }
    if (!targetPath) {
      const label = contentFormat === 'jpeg' ? 'JPEG' : contentFormat === 'psd' ? 'Photoshop' : format.toUpperCase()
      const exporting = input?.purpose === 'export'
      const fromPath = (typeof input?.path === 'string' && input.path) || source?.path
        || (typeof input?.sourcePath === 'string' && input.sourcePath) || null
      const baseName = input?.name || (fromPath && path.basename(fromPath))
      const frameSuffix = stillOfAnimation && !/\(frame 1\)$/i.test(safeStem(baseName)) ? ' (frame 1)' : ''
      const fileName = nameWithFormat(baseName, contentFormat, frameSuffix)
      const folder = fromPath && await existingDirectory(path.dirname(fromPath)) ? path.dirname(fromPath) : null
      const result = await showSaveDialogFor(event, {
        title: exporting ? `Export image as ${label}` : stillOfAnimation ? 'Save a still copy' : input?.forceDialog ? 'Save image as' : 'Save image',
        defaultPath: folder ? path.join(folder, fileName) : fileName,
        filters: [{
          name: contentFormat === 'psd' ? 'Photoshop document' : `${label} image`,
          extensions: extension === '.jpg' ? ['jpg', 'jpeg'] : [extension.slice(1)],
        }],
      })
      if (result.canceled || !result.filePath) return null
      targetPath = ensureOutputExtension(result.filePath, format)
    }

    const warnings = []
    let metadata = { kept: [], dropped: [] }
    if (contentFormat !== 'psd') {
      let original = source?.metadata || null
      if (!original && replaced) original = extractMetadata(replaced)
      if (!original && !source && typeof input?.sourcePath === 'string' && input.sourcePath.length <= 32_768) {
        const sourceBytes = await readExistingImage(input.sourcePath)
        if (sourceBytes) original = extractMetadata(sourceBytes)
      }
      // An unchanged save writes the original bytes, which already carry everything.
      const unchanged = (replaced && replaced.equals(bytes)) || (source?.digest && source.digest === digestOf(bytes))
      if (original?.format && !unchanged) {
        warnings.push(...fidelityWarnings(original, contentFormat))
        if (input?.preserveMetadata !== false) {
          const spliced = spliceMetadata(bytes, original, { colorProfile: input?.colorProfile })
          metadata = { kept: spliced.kept, dropped: spliced.dropped }
          if (spliced.bytes !== bytes) {
            try {
              validateImageBytes(spliced.bytes, extension)
              validateImageDimensions(spliced.bytes, extension)
              bytes = spliced.bytes
            } catch {
              metadata = { kept: [], dropped: ['All metadata (the result failed validation)'] }
            }
          }
        }
      }
    }
    await atomicWrite(targetPath, bytes)
    return {
      path: targetPath,
      name: path.basename(targetPath),
      size: bytes.byteLength,
      format: extensionOf(targetPath).slice(1),
      /** User-facing fidelity notes, e.g. a 16-bit source saved with 8 bits per channel. */
      warnings,
      /** Metadata blocks copied from the original (kept) and those that could not be (dropped). */
      metadata,
      /** True when the source was animated and this save wrote a single-frame still. */
      stillOfAnimation,
    }
  })
  ipcMain.handle('image:convert-to-pdf', async (event, input) => {
    requireCallingWindow(event)
    const png = validateImageBytes(input?.data, '.png')
    validateImageDimensions(png, '.png')
    const result = await showSaveDialogFor(event, {
      title: 'Export image as PDF',
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
  ipcMain.handle('clipboard:write-png', (event, input) => {
    requireCallingWindow(event)
    const bytes = validateImageBytes(input, '.png')
    const dimensions = validateImageDimensions(bytes, '.png')
    const image = nativeImage.createFromBuffer(bytes)
    if (image.isEmpty()) throw new Error('The copied image could not be decoded.')
    clipboard.writeImage(image)
    return dimensions
  })
  // Paste: the system clipboard image as validated PNG bytes (Windows offers bitmaps, which Chromium's
  // own clipboard API cannot always read), or null when the clipboard holds no image.
  ipcMain.handle('clipboard:read-image', (event) => {
    requireCallingWindow(event)
    return readClipboardPng()
  })
  ipcMain.handle('clipboard:has-image', (event) => {
    requireCallingWindow(event)
    return clipboardHasImage()
  })
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
    // The window guard asks the page again; after the in-app question it reports no changes.
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
    registerSharedIo({
      ipcMain,
      module: IO_MODULE,
      guard,
      stores,
      openInWindow: (filePath) => { createWindow(filePath); return true },
    })
    // Finish or undo any save a crash interrupted before images open.
    await sweep().catch((error) => console.error('[simple-io] sweep failed', error))
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

// Quitting is guarded by the shared app guard (installed with the first window guard):
// it waits for pending writes and asks each window about unsaved work.
