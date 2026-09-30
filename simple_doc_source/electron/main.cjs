const { app, BrowserWindow, dialog, ipcMain, shell, protocol, net } = require('electron')
const fs = require('node:fs/promises')
const { constants: fsConstants } = require('node:fs')
const path = require('node:path')
const os = require('node:os')
const crypto = require('node:crypto')
const { pathToFileURL } = require('node:url')
const { MAX_FILE_BYTES, validateDocxBytes } = require('./docx-files.cjs')
const { validateLegacyDocBytes } = require('./legacy-doc.cjs')
const { SUPPORTED_EXTENSIONS, convertLegacyDocToDocx, isSupportedDocumentPath } = require('./document-files.cjs')
const { needsOfficeLayout } = require('./document-layout.cjs')
const { exportFormat, validateExportBytes } = require('./export-files.cjs')
const { isStalePrintDirectory, loadPdfForPrinting, printWebContentsSilently, PRINT_DIRECTORY_PREFIX } = require('./print-host.cjs')
const { composePrintPdf, nativePrintOptions, resolvePrinter, validatePdfBytes } = require('./print-layout.cjs')
const { installedDocumentFonts } = require('./document-fonts.cjs')
const { findOfficeConverter, convertOfficeBytes } = require('./office-converter.cjs')
const JSZip = require('jszip')

protocol.registerSchemesAsPrivileged([{ scheme: 'simple-font', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true } }])
let documentFontsPromise
function documentFonts() { return documentFontsPromise ||= installedDocumentFonts() }

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

async function loadDocumentBytes(data, name, sourcePath = null) {
  const fileName = path.basename(String(name || ''))
  const extension = path.extname(fileName).toLowerCase()
  if (!SUPPORTED_EXTENSIONS.has(extension)) throw new Error('Simple Docs opens .docx and legacy .doc files.')
  const bytes = toBytes(data)
  if (!bytes.byteLength || bytes.byteLength > MAX_FILE_BYTES) throw new Error('The Word document is empty or too large.')

  if (extension === '.doc') {
    if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
      const validated = validateDocxBytes(bytes)
      return {
        data: new Uint8Array(validated),
        name: `${safeStem(fileName)}.docx`,
        path: null,
        sourcePath,
        size: validated.byteLength,
        convertedFrom: 'docx-renamed',
        requiresSaveAs: true,
      }
    }
    // Prepare the reading/printing view while the editable conversion runs.
    // The bounded engine cache makes the later preview request reuse this work.
    if (await findOfficeConverter()) void convertOfficeBytes({ bytes, inputExtension: 'doc', outputExtension: 'pdf', filter: 'writer_pdf_Export' }).catch(() => {})
    const converted = await convertLegacyDocToDocx(bytes, { title: fileName })
    return {
      data: new Uint8Array(converted.data),
      name: converted.conversionMethod === 'layout' ? fileName : `${safeStem(fileName)}.docx`,
      path: converted.conversionMethod === 'layout' ? sourcePath : null,
      format: converted.conversionMethod === 'layout' ? 'doc' : 'docx',
      sourceData: converted.conversionMethod === 'layout' ? new Uint8Array(bytes) : undefined,
      sourceHash: sourcePath ? crypto.createHash('sha256').update(bytes).digest('hex') : undefined,
      sourcePath,
      size: converted.data.byteLength,
      convertedFrom: 'doc',
      conversionMethod: converted.conversionMethod,
      ...(converted.conversionMethod === 'layout' ? { originalLayout: { data: new Uint8Array(bytes), extension: 'doc' } } : {}),
      requiresSaveAs: converted.conversionMethod !== 'layout' || !sourcePath,
      conversionWarnings: converted.warnings,
    }
  }

  const validated = validateDocxBytes(bytes)
  let originalLayout
  // Anchored objects and floating tables need a full Office page-layout engine.
  // Preserve their original rendered view rather than pretending canvas reflow
  // has retained their positions.
  if (await findOfficeConverter()) {
    try {
      const zip = await JSZip.loadAsync(validated)
      const parts = Object.keys(zip.files).filter((name) => /^word\/(document|header\d+|footer\d+)\.xml$/.test(name))
      for (const name of parts) {
        const xml = await zip.file(name)?.async('string')
        if (xml && needsOfficeLayout(xml, name)) {
          originalLayout = { data: new Uint8Array(validated), extension: 'docx' }
          break
        }
      }
    } catch { /* The validated document can still open in the editor. */ }
  }
  return {
    data: new Uint8Array(validated),
    name: fileName,
    path: sourcePath,
    sourcePath,
    size: validated.byteLength,
    requiresSaveAs: !sourcePath,
    originalLayout,
    sourceHash: sourcePath ? crypto.createHash('sha256').update(bytes).digest('hex') : undefined,
  }
}

async function loadDocument(filePath) {
  const resolved = path.resolve(filePath)
  if (!isSupportedDocumentPath(resolved)) throw new Error('Simple Docs opens .docx and legacy .doc files.')
  const stat = await fs.stat(resolved)
  if (!stat.isFile() || !stat.size || stat.size > MAX_FILE_BYTES) throw new Error('The Word document is empty or too large.')
  const payload = await loadDocumentBytes(await fs.readFile(resolved), path.basename(resolved), resolved)
  await addRecent(resolved)
  return payload
}

function showOpenDialogFor(event, options) {
  const owner = callingWindow(event)
  return owner ? dialog.showOpenDialog(owner, options) : dialog.showOpenDialog(options)
}

function showSaveDialogFor(event, options) {
  const owner = callingWindow(event)
  return owner ? dialog.showSaveDialog(owner, options) : dialog.showSaveDialog(options)
}

function printerCapabilityHints(options = {}) {
  const flattened = Object.entries(options)
    .map(([key, value]) => `${key}=${String(value)}`)
    .join(' ')
    .toLowerCase()
  return {
    supportsDuplex: /duplex|two[-_ ]?sided|sides[-_]supported/.test(flattened),
    supportsColor: /(^|\W)colou?r/.test(flattened) || !/monochrome|black[-_ ]?and[-_ ]?white/.test(flattened),
  }
}

function printerSummary(printer) {
  return {
    name: printer.name,
    displayName: printer.displayName || printer.name,
    ...printerCapabilityHints(printer.options),
  }
}

async function printPdf(input, owner, contents) {
  let printers
  try {
    printers = await contents.getPrintersAsync()
  } catch (error) {
    return { success: false, failureReason: `Simple could not read the Windows printers: ${error instanceof Error ? error.message : String(error)}` }
  }
  let printer
  try {
    printer = resolvePrinter(printers, input?.deviceName)
  } catch (error) {
    return { success: false, failureReason: error instanceof Error ? error.message : String(error) }
  }
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), PRINT_DIRECTORY_PREFIX))
  const pdfPath = path.join(directory, `${safeStem(input?.name)}.pdf`)
  let printWindow = null
  const cleanup = () => fs.rm(directory, { recursive: true, force: true }).catch(() => {})
  try {
    await fs.writeFile(pdfPath, validatePdfBytes(input?.data))
    printWindow = new BrowserWindow({
      ...(owner ? { parent: owner } : {}),
      show: false,
      title: `Printing — ${safeStem(input?.name)}`,
      backgroundColor: '#ffffff',
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
    await loadPdfForPrinting(printWindow, pathToFileURL(pdfPath).href)
    const options = nativePrintOptions({ ...input, deviceName: printer.name })
    const printed = await printWebContentsSilently(printWindow.webContents, options)
    if (printWindow && !printWindow.isDestroyed()) printWindow.destroy()
    await cleanup()
    return {
      ...printed,
      printerName: printer.name,
      printerLabel: printer.displayName || printer.name,
    }
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
    .filter((entry) => entry.isDirectory() && entry.name.startsWith(PRINT_DIRECTORY_PREFIX))
    .map(async (entry) => {
      const target = path.join(temporaryRoot, entry.name)
      try {
        const stats = await fs.stat(target)
        if (isStalePrintDirectory(entry.name, stats)) await fs.rm(target, { recursive: true, force: true })
      } catch {}
    }))
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
      plugins: true,
      devTools: !app.isPackaged,
    },
  })
  window.removeMenu()
  window.webContents.on('before-input-event', (event, input) => {
    if (input.type !== 'keyDown' || !(input.control || input.meta) || input.alt) return
    const key = String(input.key).toLowerCase()
    const action = key === 'p' ? 'print' : key === 's' ? input.shift ? 'save-as' : 'save' : key === 'e' && input.shift ? 'export' : null
    if (!action) return
    event.preventDefault()
    window.webContents.send('document:shortcut', action)
  })
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
  ipcMain.handle('document:fonts', async () => (await documentFonts()).fonts)
  ipcMain.handle('document:original-pdf', async (_event, input) => {
    const extension = input?.extension
    if (!['doc', 'docx'].includes(extension)) throw new Error('Unsupported original document format.')
    const bytes = toBytes(input?.data)
    if (!bytes.length || bytes.length > MAX_FILE_BYTES) throw new Error('The original document is empty or too large.')
    return new Uint8Array(validatePdfBytes(await convertOfficeBytes({ bytes, inputExtension: extension, outputExtension: 'pdf', filter: 'writer_pdf_Export' })))
  })
  ipcMain.handle('file:open-dialog', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Open a Word document',
      properties: ['openFile'],
      filters: [{ name: 'Word documents', extensions: ['docx', 'doc'] }],
    })
    if (result.canceled || !result.filePaths[0]) return null
    return loadDocument(result.filePaths[0])
  })
  ipcMain.handle('file:open-path', (_event, filePath) => loadDocument(filePath))
  ipcMain.handle('file:open-bytes', (_event, input) => loadDocumentBytes(input?.data, input?.name))
  ipcMain.handle('file:open-in-new-window', async (event, requestedPath = null) => {
    let selectedPath = requestedPath
    if (!selectedPath) {
      const result = await showOpenDialogFor(event, {
        title: 'Open a Word document', properties: ['openFile'], filters: [{ name: 'Word documents', extensions: ['docx', 'doc'] }],
      })
      if (result.canceled || !result.filePaths[0]) return false
      selectedPath = result.filePaths[0]
    }
    if (!isSupportedDocumentPath(selectedPath)) throw new Error('Simple Docs opens .docx and legacy .doc files.')
    await fs.access(selectedPath)
    createWindow(path.resolve(selectedPath))
    return true
  })
  ipcMain.handle('app:new-window', () => { createWindow(); return true })
  ipcMain.handle('file:save-docx', async (event, input) => {
    const format = input.format === 'doc' && !input.forceDialog ? 'doc' : 'docx'
    const extension = `.${format}`
    let data = validateDocxBytes(input.data)
    let target = input.forceDialog ? null : input.path
    if (!target) {
      const result = await showSaveDialogFor(event, {
        title: input.forceDialog ? 'Save document as' : 'Save document',
        defaultPath: ensureExtension(safeStem(input.name), extension),
        filters: [{ name: format === 'doc' ? 'Word 97–2003 document' : 'Word document', extensions: [format] }],
      })
      if (result.canceled || !result.filePath) return null
      target = ensureExtension(result.filePath, extension)
    }
    target = ensureExtension(path.resolve(target), extension)
    if (format === 'doc') data = input.sourceData
      ? validateLegacyDocBytes(input.sourceData)
      : validateLegacyDocBytes(await convertOfficeBytes({ bytes: data, inputExtension: 'docx', outputExtension: 'doc', filter: 'MS Word 97' }))
    if (!input.forceDialog && input.path && input.expectedHash) {
      const existing = await fs.readFile(target).catch(() => null)
      if (!existing || crypto.createHash('sha256').update(existing).digest('hex') !== input.expectedHash) throw new Error('This file changed outside Simple Docs. Use Save as to keep both versions.')
    }
    if (!input.forceDialog && input.path && input.protectOriginal) {
      const backupPath = path.join(path.dirname(target), `${path.basename(target, path.extname(target))}.before-simple-edit${path.extname(target)}`)
      try { await fs.copyFile(target, backupPath, fsConstants.COPYFILE_EXCL) } catch (error) { if (!['ENOENT', 'EEXIST'].includes(error.code)) throw error }
    }
    await atomicWrite(target, data)
    await addRecent(target)
    return { path: target, name: path.basename(target), format, sourceHash: crypto.createHash('sha256').update(data).digest('hex'), ...(format === 'doc' ? { sourceData: new Uint8Array(data) } : {}) }
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
  ipcMain.handle('file:save-export', async (event, input) => {
    const definition = exportFormat(input?.format)
    let data = validateExportBytes(input?.data)
    if (definition.id === 'docx') data = validateDocxBytes(data)
    const result = await showSaveDialogFor(event, {
      title: `Export as ${definition.label}`,
      defaultPath: `${safeStem(input?.name)}${definition.extension}`,
      filters: [{ name: definition.label, extensions: [definition.extension.slice(1)] }],
    })
    if (result.canceled || !result.filePath) return null
    const target = ensureExtension(path.resolve(result.filePath), definition.extension)
    await atomicWrite(target, data)
    return { path: target, name: path.basename(target) }
  })
  ipcMain.handle('file:compose-print-pdf', (_event, input) => composePrintPdf(input))
  ipcMain.handle('print:list-printers', async (event) => (await event.sender.getPrintersAsync()).map(printerSummary))
  ipcMain.handle('file:print-pdf', (event, input) => printPdf(input, callingWindow(event), event.sender))
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

function incomingDocumentPaths(argv) {
  return [...new Set(argv.filter((argument) => isSupportedDocumentPath(argument)).map((argument) => path.resolve(argument)))]
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) app.quit()
else {
  app.on('second-instance', (_event, argv) => {
    const incoming = incomingDocumentPaths(argv)
    if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
    else {
      const window = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0]
      if (window) { if (window.isMinimized()) window.restore(); window.focus() }
      else createWindow()
    }
  })
  app.whenReady().then(() => {
    app.setAppUserModelId('com.simple.docs')
    protocol.handle('simple-font', async (request) => {
      const url = new URL(request.url)
      const filePath = url.hostname === 'installed' && /^\/\d+$/.test(url.pathname)
        ? (await documentFonts()).files.get(url.pathname.slice(1)) : null
      if (!filePath) return new Response('Font not found', { status: 404 })
      const response = await net.fetch(pathToFileURL(filePath).href)
      return new Response(response.body, { headers: { 'Content-Type': 'font/ttf', 'Access-Control-Allow-Origin': '*', 'Cache-Control': 'public, max-age=86400' } })
    })
    registerIpc()
    void cleanupStalePrintDirectories().finally(() => {
      const incoming = incomingDocumentPaths(process.argv)
      if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
      else createWindow()
    })
  })
}

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
