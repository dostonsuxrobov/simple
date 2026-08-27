const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const crypto = require('node:crypto')
const { pathToFileURL } = require('node:url')

const DOCX_EXTENSION = '.docx'
const MAX_FILE_BYTES = 256 * 1024 * 1024
const MAX_ZIP_ENTRIES = 20_000
const MAX_ZIP_ENTRY_BYTES = 512 * 1024 * 1024
const MAX_ZIP_TOTAL_BYTES = 1024 * 1024 * 1024
const closeApprovedWindows = new WeakSet()

function callingWindow(event) {
  const window = BrowserWindow.fromWebContents(event.sender)
  return window && !window.isDestroyed() ? window : null
}

function toBytes(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data)
  return Buffer.from(value)
}

function safeStem(value) {
  return path.basename(String(value || 'Untitled document'), path.extname(String(value || '')))
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim() || 'Untitled document'
}

function ensureExtension(filePath, extension) {
  return filePath.toLowerCase().endsWith(extension) ? filePath : `${filePath}${extension}`
}

function validateDocxPackage(bytes) {
  const minimumEocdOffset = Math.max(0, bytes.length - 65_557)
  let eocdOffset = -1
  for (let offset = bytes.length - 22; offset >= minimumEocdOffset; offset -= 1) {
    if (bytes.readUInt32LE(offset) === 0x06054b50) {
      eocdOffset = offset
      break
    }
  }
  if (eocdOffset < 0) throw new Error('This DOCX archive is incomplete.')

  const entryCount = bytes.readUInt16LE(eocdOffset + 10)
  const centralSize = bytes.readUInt32LE(eocdOffset + 12)
  const centralOffset = bytes.readUInt32LE(eocdOffset + 16)
  if (entryCount === 0xffff || centralSize === 0xffffffff || centralOffset === 0xffffffff) {
    throw new Error('This DOCX uses an unsupported ZIP64 layout.')
  }
  if (!entryCount || entryCount > MAX_ZIP_ENTRIES || centralOffset + centralSize > bytes.length) {
    throw new Error('This DOCX archive has an unsafe directory structure.')
  }

  let cursor = centralOffset
  let totalUncompressed = 0
  let hasContentTypes = false
  let hasMainDocument = false
  for (let index = 0; index < entryCount; index += 1) {
    if (cursor + 46 > bytes.length || bytes.readUInt32LE(cursor) !== 0x02014b50) {
      throw new Error('This DOCX archive has a damaged file directory.')
    }
    const flags = bytes.readUInt16LE(cursor + 8)
    const compressedSize = bytes.readUInt32LE(cursor + 20)
    const uncompressedSize = bytes.readUInt32LE(cursor + 24)
    const nameLength = bytes.readUInt16LE(cursor + 28)
    const extraLength = bytes.readUInt16LE(cursor + 30)
    const commentLength = bytes.readUInt16LE(cursor + 32)
    const nextCursor = cursor + 46 + nameLength + extraLength + commentLength
    if (nextCursor > bytes.length || compressedSize === 0xffffffff || uncompressedSize === 0xffffffff) {
      throw new Error('This DOCX archive contains an unsupported entry.')
    }
    if ((flags & 0x1) !== 0) throw new Error('Password-protected DOCX files are not supported.')
    if (uncompressedSize > MAX_ZIP_ENTRY_BYTES) throw new Error('This DOCX contains an entry that is too large.')
    totalUncompressed += uncompressedSize
    if (totalUncompressed > MAX_ZIP_TOTAL_BYTES) throw new Error('This DOCX expands beyond the safe size limit.')

    const entryName = bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8').replace(/\\/g, '/')
    if (entryName.startsWith('/') || entryName.split('/').includes('..')) {
      throw new Error('This DOCX contains an unsafe file path.')
    }
    if (entryName === '[Content_Types].xml') hasContentTypes = true
    if (entryName === 'word/document.xml') hasMainDocument = true
    cursor = nextCursor
  }
  if (cursor > centralOffset + centralSize || !hasContentTypes || !hasMainDocument) {
    throw new Error('This archive does not contain a complete Word document.')
  }
}

function validateDocxBytes(data) {
  const bytes = toBytes(data)
  if (bytes.byteLength < 4 || bytes.byteLength > MAX_FILE_BYTES) throw new Error('The DOCX file is empty or too large.')
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('This is not a valid DOCX file.')
  validateDocxPackage(bytes)
  return bytes
}

async function atomicWrite(targetPath, data) {
  const directory = path.dirname(targetPath)
  const extension = path.extname(targetPath) || '.tmp'
  const token = crypto.randomUUID()
  const temporary = path.join(directory, `.${path.basename(targetPath, extension)}-${token}${extension}.tmp`)
  const backup = path.join(directory, `.${path.basename(targetPath)}-${token}.bak`)
  let backedUp = false
  await fs.writeFile(temporary, toBytes(data))
  try {
    try {
      await fs.rename(targetPath, backup)
      backedUp = true
    } catch (error) {
      if (error.code !== 'ENOENT') throw error
    }
    await fs.rename(temporary, targetPath)
    if (backedUp) await fs.rm(backup, { force: true })
  } catch (error) {
    await fs.rm(temporary, { force: true }).catch(() => {})
    if (backedUp) await fs.rename(backup, targetPath).catch(() => {})
    throw error
  }
}

function statePath(fileName) {
  return path.join(app.getPath('userData'), fileName)
}

async function readJson(fileName, fallback) {
  try { return JSON.parse(await fs.readFile(statePath(fileName), 'utf8')) } catch { return fallback }
}

async function writeJson(fileName, value) {
  await fs.mkdir(app.getPath('userData'), { recursive: true })
  await atomicWrite(statePath(fileName), Buffer.from(JSON.stringify(value, null, 2), 'utf8'))
}

async function addRecent(filePath) {
  const recents = await readJson('recent-files.json', [])
  const normalized = path.resolve(filePath)
  const next = [{ path: normalized, name: path.basename(normalized), openedAt: Date.now() }, ...recents.filter((item) => item.path?.toLowerCase() !== normalized.toLowerCase())].slice(0, 12)
  await writeJson('recent-files.json', next)
}

async function loadDocument(filePath) {
  const resolved = path.resolve(filePath)
  if (path.extname(resolved).toLowerCase() !== DOCX_EXTENSION) throw new Error('Simple Docs opens .docx files.')
  const stat = await fs.stat(resolved)
  if (!stat.isFile() || stat.size > MAX_FILE_BYTES) throw new Error('The DOCX file is too large.')
  const data = validateDocxBytes(await fs.readFile(resolved))
  await addRecent(resolved)
  return { data: new Uint8Array(data), name: path.basename(resolved), path: resolved, size: data.byteLength }
}

function showOpenDialogFor(event, options) {
  const owner = callingWindow(event)
  return owner ? dialog.showOpenDialog(owner, options) : dialog.showOpenDialog(options)
}

function showSaveDialogFor(event, options) {
  const owner = callingWindow(event)
  return owner ? dialog.showSaveDialog(owner, options) : dialog.showSaveDialog(options)
}

async function printPdf(data, name, owner) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-docs-print-'))
  const pdfPath = path.join(directory, `${safeStem(name)}.pdf`)
  let printWindow = null
  const cleanup = () => fs.rm(directory, { recursive: true, force: true }).catch(() => {})
  try {
    await fs.writeFile(pdfPath, toBytes(data))
    printWindow = new BrowserWindow({
      ...(owner ? { parent: owner } : {}),
      width: 1040,
      height: 820,
      minWidth: 720,
      minHeight: 540,
      show: false,
      title: `Print preview — ${safeStem(name)}`,
      backgroundColor: '#f2f2f2',
      webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true, plugins: true, devTools: false },
    })
    printWindow.removeMenu()
    printWindow.webContents.setWindowOpenHandler(({ url }) => {
      if (/^(https?:|mailto:)/i.test(url)) void shell.openExternal(url)
      return { action: 'deny' }
    })
    printWindow.webContents.on('will-navigate', (event, url) => {
      if (url !== pathToFileURL(pdfPath).href) event.preventDefault()
    })
    printWindow.on('closed', cleanup)
    await printWindow.loadURL(pathToFileURL(pdfPath).href)
    printWindow.show()
    return true
  } catch (error) {
    if (printWindow && !printWindow.isDestroyed()) printWindow.destroy()
    await cleanup()
    throw error
  }
}

async function cleanupStalePrintDirectories() {
  const temporaryRoot = os.tmpdir()
  let entries = []
  try { entries = await fs.readdir(temporaryRoot, { withFileTypes: true }) } catch { return }
  await Promise.all(entries
    .filter((entry) => entry.isDirectory() && entry.name.startsWith('simple-docs-print-'))
    .map((entry) => fs.rm(path.join(temporaryRoot, entry.name), { recursive: true, force: true }).catch(() => {})))
}

function createWindow(openPath = null) {
  const window = new BrowserWindow({
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    width: 1480,
    height: 940,
    minWidth: 880,
    minHeight: 620,
    show: false,
    frame: false,
    backgroundColor: '#f2f2f2',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: !app.isPackaged,
    },
  })
  window.removeMenu()
  if (!app.isPackaged) {
    window.webContents.on('did-fail-load', (_event, code, description, url) => {
      console.error(`[load:${code}] ${description} ${url}`)
    })
  }
  window.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?:|mailto:)/i.test(url)) void shell.openExternal(url)
    return { action: 'deny' }
  })
  window.webContents.on('will-navigate', (event, url) => {
    const current = window.webContents.getURL()
    if (url !== current) {
      event.preventDefault()
      if (/^(https?:|mailto:)/i.test(url)) void shell.openExternal(url)
    }
  })
  if (process.env.VITE_DEV_SERVER_URL) window.loadURL(process.env.VITE_DEV_SERVER_URL)
  else window.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))
  window.once('ready-to-show', () => window.show())
  window.webContents.once('did-finish-load', () => { if (openPath) window.webContents.send('file:open-external', openPath) })
  window.on('maximize', () => window.webContents.send('window:maximized', true))
  window.on('unmaximize', () => window.webContents.send('window:maximized', false))
  window.on('enter-full-screen', () => window.webContents.send('window:fullscreen', true))
  window.on('leave-full-screen', () => window.webContents.send('window:fullscreen', false))
  window.on('close', (event) => {
    if (closeApprovedWindows.has(window)) return
    event.preventDefault()
    window.webContents.send('window:close-requested')
  })
  return window
}

function registerIpc() {
  ipcMain.handle('file:open-dialog', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Open a Word document',
      properties: ['openFile'],
      filters: [{ name: 'Word documents', extensions: ['docx'] }],
    })
    if (result.canceled || !result.filePaths[0]) return null
    return loadDocument(result.filePaths[0])
  })
  ipcMain.handle('file:open-path', (_event, filePath) => loadDocument(filePath))
  ipcMain.handle('file:open-in-new-window', async (event, requestedPath = null) => {
    let selectedPath = requestedPath
    if (!selectedPath) {
      const result = await showOpenDialogFor(event, {
        title: 'Open a Word document', properties: ['openFile'], filters: [{ name: 'Word documents', extensions: ['docx'] }],
      })
      if (result.canceled || !result.filePaths[0]) return false
      selectedPath = result.filePaths[0]
    }
    if (path.extname(selectedPath).toLowerCase() !== DOCX_EXTENSION) throw new Error('Simple Docs opens .docx files.')
    await fs.access(selectedPath)
    createWindow(path.resolve(selectedPath))
    return true
  })
  ipcMain.handle('app:new-window', () => { createWindow(); return true })
  ipcMain.handle('file:save-docx', async (event, input) => {
    const data = validateDocxBytes(input.data)
    let target = input.forceDialog ? null : input.path
    if (!target) {
      const result = await showSaveDialogFor(event, {
        title: input.forceDialog ? 'Save document as' : 'Save document',
        defaultPath: ensureExtension(safeStem(input.name), '.docx'),
        filters: [{ name: 'Word document', extensions: ['docx'] }],
      })
      if (result.canceled || !result.filePath) return null
      target = ensureExtension(result.filePath, '.docx')
    }
    target = ensureExtension(path.resolve(target), '.docx')
    await atomicWrite(target, data)
    await addRecent(target)
    return { path: target, name: path.basename(target) }
  })
  ipcMain.handle('file:save-pdf', async (event, input) => {
    const result = await showSaveDialogFor(event, {
      title: 'Export PDF', defaultPath: `${safeStem(input.name)}.pdf`, filters: [{ name: 'PDF document', extensions: ['pdf'] }],
    })
    if (result.canceled || !result.filePath) return null
    const target = ensureExtension(result.filePath, '.pdf')
    await atomicWrite(target, input.data)
    return { path: target, name: path.basename(target) }
  })
  ipcMain.handle('file:print-pdf', (event, input) => printPdf(input.data, input.name, callingWindow(event)))
  ipcMain.handle('recent:list', async () => {
    const recents = await readJson('recent-files.json', [])
    const valid = []
    for (const item of recents) {
      try { await fs.access(item.path); valid.push(item) } catch {}
    }
    if (valid.length !== recents.length) await writeJson('recent-files.json', valid)
    return valid
  })
  ipcMain.handle('recent:remove', async (_event, filePath) => {
    const recents = await readJson('recent-files.json', [])
    await writeJson('recent-files.json', recents.filter((item) => item.path?.toLowerCase() !== String(filePath).toLowerCase()))
  })
  ipcMain.handle('recovery:save', async (_event, input) => {
    const safeId = String(input.id).replace(/[^a-zA-Z0-9-]/g, '')
    if (!safeId) throw new Error('Invalid recovery id.')
    const directory = statePath('recovery')
    await fs.mkdir(directory, { recursive: true })
    const target = path.join(directory, `${safeId}.docx`)
    await atomicWrite(target, validateDocxBytes(input.data))
    const recoveries = await readJson('recoveries.json', [])
    const next = [{ id: safeId, title: safeStem(input.title), sourcePath: input.sourcePath || null, updatedAt: Date.now() }, ...recoveries.filter((item) => item.id !== safeId)].slice(0, 10)
    await writeJson('recoveries.json', next)
  })
  ipcMain.handle('recovery:list', async () => {
    const recoveries = await readJson('recoveries.json', [])
    const valid = []
    for (const item of recoveries) {
      try { await fs.access(statePath(path.join('recovery', `${item.id}.docx`))); valid.push(item) } catch {}
    }
    if (valid.length !== recoveries.length) await writeJson('recoveries.json', valid)
    return valid
  })
  ipcMain.handle('recovery:load', async (_event, id) => {
    const safeId = String(id).replace(/[^a-zA-Z0-9-]/g, '')
    const recoveries = await readJson('recoveries.json', [])
    const metadata = recoveries.find((item) => item.id === safeId)
    if (!metadata) throw new Error('Recovery copy not found.')
    const data = validateDocxBytes(await fs.readFile(statePath(path.join('recovery', `${safeId}.docx`))))
    return { data: new Uint8Array(data), name: `${metadata.title}.docx`, path: metadata.sourcePath || null, size: data.byteLength }
  })
  ipcMain.handle('recovery:clear', async (_event, id) => {
    const safeId = String(id).replace(/[^a-zA-Z0-9-]/g, '')
    await fs.rm(statePath(path.join('recovery', `${safeId}.docx`)), { force: true }).catch(() => {})
    const recoveries = await readJson('recoveries.json', [])
    await writeJson('recoveries.json', recoveries.filter((item) => item.id !== safeId))
  })
  ipcMain.handle('shell:open-external', async (_event, url) => {
    const parsed = new URL(url)
    if (!['http:', 'https:', 'mailto:'].includes(parsed.protocol)) throw new Error('This link type is not allowed.')
    await shell.openExternal(parsed.toString())
  })
  ipcMain.on('window:set-title', (event, title) => callingWindow(event)?.setTitle(String(title).slice(0, 260)))
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
  ipcMain.on('window:confirm-close', (event) => {
    const window = callingWindow(event)
    if (!window) return
    closeApprovedWindows.add(window)
    window.close()
  })
}

function incomingDocxPaths(argv) {
  return [...new Set(argv.filter((argument) => path.extname(argument).toLowerCase() === DOCX_EXTENSION).map((argument) => path.resolve(argument)))]
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) app.quit()
else {
  app.on('second-instance', (_event, argv) => {
    const incoming = incomingDocxPaths(argv)
    if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
    else {
      const window = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0]
      if (window) { if (window.isMinimized()) window.restore(); window.focus() }
      else createWindow()
    }
  })
  app.whenReady().then(() => {
    app.setAppUserModelId('com.simple.docs')
    registerIpc()
    void cleanupStalePrintDirectories().finally(() => {
      const incoming = incomingDocxPaths(process.argv)
      if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
      else createWindow()
    })
  })
}

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
