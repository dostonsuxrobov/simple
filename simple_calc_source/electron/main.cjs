const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron')
const { registerClipboardHandlers } = require('./clipboard-bridge.cjs')
const { registerProtectionHandlers } = require('./protection.cjs')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const {
  SUPPORTED_EXTENSIONS,
  workbookPayloadFromPath,
  workbookPayloadFromBytes,
  serializeWorkbook,
} = require('./workbooks.cjs')
const { createSpreadsheetPrintDocument } = require('./spreadsheet-print.cjs')
const { createSpreadsheetExport, exportFilter, normalizeExportFormat } = require('./spreadsheet-export.cjs')
const { directPrintOptions, ensurePrinterInstalled, printFailureMessage } = require('./default-printer.cjs')
const { EDITABLE_SAVE_FORMATS, saveFormat, unchangedSourceBytes, assertSourceUnchanged } = require('./workbook-save.cjs')

app.setName('simple_calc')

const MAX_OPEN_BYTES = 512 * 1024 * 1024
const RICH_OOXML_FORMATS = new Set(['xlsx', 'xlsm', 'xltx', 'xltm', 'xlam'])
const closeApprovedWindows = new WeakSet()
const documentsByWebContents = new Map()
let isQuitting = false

function callingWindow(event) {
  const browserWindow = BrowserWindow.fromWebContents(event.sender)
  return browserWindow && !browserWindow.isDestroyed() ? browserWindow : null
}

function assertTrustedSender(event) {
  const frame = event.senderFrame
  if (!frame || frame !== event.sender.mainFrame) throw new Error('Untrusted request.')
  const url = new URL(frame.url)
  if (process.env.VITE_DEV_SERVER_URL) {
    const allowed = new URL(process.env.VITE_DEV_SERVER_URL)
    if (url.origin !== allowed.origin) throw new Error('Untrusted request origin.')
  } else if (url.protocol !== 'file:') {
    throw new Error('Untrusted request origin.')
  }
}

function extensionOf(filePath) {
  return path.extname(String(filePath || '')).toLowerCase()
}

function assertSupportedPath(filePath) {
  if (typeof filePath !== 'string' || !filePath.trim()) throw new Error('A workbook path is required.')
  if (!SUPPORTED_EXTENSIONS.has(extensionOf(filePath))) throw new Error('This spreadsheet format is not supported.')
  return path.resolve(filePath)
}

function workbookFilters() {
  return [
    { name: 'Spreadsheet files', extensions: ['xlsx', 'xlsm', 'xlsb', 'xls', 'xltx', 'xltm', 'xlt', 'xlam', 'xla', 'xml', 'ods', 'fods', 'numbers', 'csv', 'tsv', 'tab', 'txt', 'slk', 'sylk', 'dif', 'dbf', 'prn', 'wk1', 'wk2', 'wk3', 'wk4', 'wks', 'wq1', 'wq2', 'wb1', 'wb2', 'wb3', '123', 'qpw', 'html', 'htm'] },
    { name: 'Modern Excel workbooks', extensions: ['xlsx', 'xlsm', 'xlsb'] },
    { name: 'OpenDocument spreadsheets', extensions: ['ods', 'fods'] },
    { name: 'Delimited text', extensions: ['csv', 'tsv', 'txt'] },
    { name: 'All files', extensions: ['*'] },
  ]
}

function saveFilters(preferred = 'xlsx') {
  const filters = [
    { name: 'Excel workbook', extensions: ['xlsx'] },
    { name: 'Excel 97–2003 workbook', extensions: ['xls'] },
    { name: 'OpenDocument spreadsheet', extensions: ['ods'] },
    { name: 'Comma-separated values', extensions: ['csv'] },
    { name: 'Tab-separated values', extensions: ['tsv'] },
  ]
  if (!EDITABLE_SAVE_FORMATS.has(preferred)) filters.push({ name: `${preferred.toUpperCase()} original file`, extensions: [preferred] })
  return filters.sort((left, right) => Number(right.extensions.includes(preferred)) - Number(left.extensions.includes(preferred)))
}

function formatForPath(filePath, fallback = 'xlsx') {
  const ext = extensionOf(filePath).slice(1)
  return EDITABLE_SAVE_FORMATS.has(ext) || ext === fallback ? ext : fallback
}

function safeSuggestedName(value, format = 'xlsx') {
  const base = path.basename(String(value || 'Untitled'), path.extname(String(value || '')))
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim() || 'Untitled'
  return `${base}.${format}`
}

function documentStore(event) {
  const id = event.sender.id
  let store = documentsByWebContents.get(id)
  if (!store) {
    store = new Map()
    documentsByWebContents.set(id, store)
  }
  return store
}

function rememberDocument(event, payload, sourcePath = null, sourceSnapshot = null) {
  const documentId = crypto.randomUUID()
  documentStore(event).set(documentId, {
    path: sourcePath,
    sourceFormat: payload.sourceFormat,
    originalName: payload.name,
    sourceSnapshot,
  })
  return { ...payload, documentId, path: sourcePath }
}

async function readWorkbookPath(event, requestedPath) {
  const filePath = assertSupportedPath(requestedPath)
  const info = await fs.stat(filePath)
  if (!info.isFile()) throw new Error('The selected item is not a file.')
  if (info.size > MAX_OPEN_BYTES) throw new Error('This workbook is larger than the 512 MB safety limit.')
  const payload = await workbookPayloadFromPath(filePath)
  return rememberDocument(event, payload, filePath, {
    path: filePath,
    size: info.size,
    modified: info.mtimeMs,
  })
}

async function sourcePackageBytes(record) {
  if (!record || !RICH_OOXML_FORMATS.has(record.sourceFormat)) return null
  return unchangedSourceBytes(record)
}

async function atomicWrite(targetPath, data) {
  const bytes = data && data.data && !Buffer.isBuffer(data) ? data.data : data
  const directory = path.dirname(targetPath)
  const temporaryPath = path.join(directory, `.${path.basename(targetPath)}.${crypto.randomUUID()}.tmp`)
  try {
    await fs.writeFile(temporaryPath, bytes)
    await fs.rename(temporaryPath, targetPath).catch(async (error) => {
      if (!['EEXIST', 'EPERM'].includes(error.code)) throw error
      const backupPath = `${temporaryPath}.previous`
      await fs.rename(targetPath, backupPath)
      try {
        await fs.rename(temporaryPath, targetPath)
        await fs.rm(backupPath, { force: true })
      } catch (replaceError) {
        await fs.rename(backupPath, targetPath).catch(() => {})
        throw replaceError
      }
    })
  } finally {
    await fs.rm(temporaryPath, { force: true }).catch(() => {})
  }
}

async function renderPrintDocumentToPdf(printDocument, directory) {
  const htmlPath = path.join(directory, 'spreadsheet.html')
  let generatorWindow = null
  try {
    await fs.writeFile(htmlPath, printDocument.html, 'utf8')
    generatorWindow = new BrowserWindow({
      show: false,
      backgroundColor: '#ffffff',
      webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true, devTools: false },
    })
    generatorWindow.removeMenu()
    await generatorWindow.loadFile(htmlPath)
    return await generatorWindow.webContents.printToPDF({
      printBackground: true,
      landscape: printDocument.options.orientation === 'landscape',
      preferCSSPageSize: true,
      generateTaggedPDF: true,
      generateDocumentOutline: true,
    })
  } finally {
    if (generatorWindow && !generatorWindow.isDestroyed()) generatorWindow.destroy()
  }
}

function printPreviewResult(printDocument) {
  return {
    html: printDocument.html,
    title: printDocument.title,
    sheets: printDocument.sheetCount,
    cells: printDocument.printedCells,
    pages: printDocument.pageCount,
    pageBreaks: printDocument.pageBreaks,
    minimumScale: printDocument.minimumScale,
    oversizedDimensions: printDocument.oversizedDimensions,
    paper: printDocument.paper,
    options: printDocument.options,
    warnings: printDocument.warnings || [],
  }
}

async function printSpreadsheet(input, owner) {
  const printDocument = createSpreadsheetPrintDocument(input)
  if (printDocument.mixedPaperSizes) throw new Error('This workbook uses different paper sizes. Print each sheet separately so the printer keeps its saved page layout.')
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-calc-print-'))
  const htmlPath = path.join(directory, 'spreadsheet.html')
  let printWindow = null
  const cleanup = () => fs.rm(directory, { recursive: true, force: true }).catch(() => {})
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
    await ensurePrinterInstalled(printWindow.webContents)
    const outcome = await new Promise((resolve, reject) => {
      printWindow.webContents.print(directPrintOptions({
        printBackground: true,
        color: true,
        landscape: printDocument.options.orientation === 'landscape',
        margins: { marginType: 'none' },
        pageSize: printDocument.options.paperSize === 'a4' ? 'A4' : printDocument.options.paperSize === 'legal' ? 'Legal' : 'Letter',
        scaleFactor: 100,
        pagesPerSheet: 1,
        collate: true,
      }), (success, failureReason) => {
        if (success) {
          resolve({ printed: true, canceled: false })
          return
        }
        reject(new Error(printFailureMessage(failureReason)))
      })
    })
    return {
      ...outcome,
      sheets: printDocument.sheetCount,
      cells: printDocument.printedCells,
      pages: printDocument.pageCount,
    }
  } finally {
    if (printWindow && !printWindow.isDestroyed()) printWindow.destroy()
    await cleanup()
  }
}

async function cleanupStalePrintDirectories() {
  let entries = []
  try { entries = await fs.readdir(os.tmpdir(), { withFileTypes: true }) } catch { return }
  const cutoff = Date.now() - 24 * 60 * 60 * 1_000
  await Promise.all(entries
    .filter((entry) => entry.isDirectory() && entry.name.startsWith('simple-calc-print-'))
    .map(async (entry) => {
      const target = path.join(os.tmpdir(), entry.name)
      try {
        const stat = await fs.stat(target)
        if (stat.mtimeMs < cutoff) await fs.rm(target, { recursive: true, force: true })
      } catch {}
    }))
}

function createWindow(openPath = null) {
  const browserWindow = new BrowserWindow({
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    width: 1440,
    height: 900,
    minWidth: 900,
    minHeight: 620,
    show: false,
    frame: false,
    backgroundColor: '#f7f7f5',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: !app.isPackaged,
    },
  })

  browserWindow.removeMenu()
  browserWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  browserWindow.webContents.on('will-navigate', (event, targetUrl) => {
    const current = new URL(browserWindow.webContents.getURL())
    const target = new URL(targetUrl)
    if (target.origin !== current.origin || target.protocol !== current.protocol) event.preventDefault()
  })
  if (process.env.VITE_DEV_SERVER_URL) browserWindow.loadURL(process.env.VITE_DEV_SERVER_URL)
  else browserWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))

  browserWindow.once('ready-to-show', () => browserWindow.show())
  browserWindow.webContents.once('did-finish-load', () => {
    if (openPath) browserWindow.webContents.send('file:open-external', openPath)
  })
  browserWindow.on('maximize', () => browserWindow.webContents.send('window:maximized', true))
  browserWindow.on('unmaximize', () => browserWindow.webContents.send('window:maximized', false))
  browserWindow.on('close', (event) => {
    if (isQuitting || closeApprovedWindows.has(browserWindow)) return
    event.preventDefault()
    browserWindow.webContents.send('window:close-requested')
  })
  browserWindow.webContents.on('destroyed', () => documentsByWebContents.delete(browserWindow.webContents.id))
  return browserWindow
}

function supportedPaths(argv) {
  return [...new Set(argv.filter((argument) => typeof argument === 'string' && SUPPORTED_EXTENSIONS.has(extensionOf(argument))))]
}

function registerIpc() {
  // Rich clipboard (HTML plus Excel's XML Spreadsheet flavour, which carries formulas).
  registerClipboardHandlers(ipcMain, { assertTrustedSender })
  registerProtectionHandlers(ipcMain, { assertTrustedSender })
  ipcMain.handle('workbook:create', (event) => {
    assertTrustedSender(event)
    const documentId = crypto.randomUUID()
    documentStore(event).set(documentId, { path: null, sourceFormat: 'xlsx', originalName: 'Untitled.xlsx' })
    return { documentId }
  })

  ipcMain.handle('workbook:new-window', async (event, requestedPath = null) => {
    assertTrustedSender(event)
    let selectedPath = requestedPath
    if (!selectedPath) {
      const result = await dialog.showOpenDialog(callingWindow(event), {
        title: 'Open a spreadsheet', properties: ['openFile'], filters: workbookFilters(),
      })
      if (result.canceled || !result.filePaths[0]) return false
      selectedPath = result.filePaths[0]
    }
    createWindow(assertSupportedPath(selectedPath))
    return true
  })

  ipcMain.handle('workbook:open-dialog', async (event) => {
    assertTrustedSender(event)
    const result = await dialog.showOpenDialog(callingWindow(event), {
      title: 'Open a spreadsheet', properties: ['openFile'], filters: workbookFilters(),
    })
    if (result.canceled || !result.filePaths[0]) return null
    return readWorkbookPath(event, result.filePaths[0])
  })

  ipcMain.handle('workbook:open-path', async (event, requestedPath) => {
    assertTrustedSender(event)
    return readWorkbookPath(event, requestedPath)
  })

  ipcMain.handle('workbook:open-bytes', async (event, input) => {
    assertTrustedSender(event)
    if (!input || typeof input.name !== 'string' || !input.data) throw new Error('Invalid dropped workbook.')
    const bytes = Buffer.from(input.data)
    if (bytes.byteLength > MAX_OPEN_BYTES) throw new Error('This workbook is larger than the 512 MB safety limit.')
    if (!SUPPORTED_EXTENSIONS.has(extensionOf(input.name))) throw new Error('This spreadsheet format is not supported.')
    const payload = await workbookPayloadFromBytes(input.name, bytes)
    return rememberDocument(event, payload, null, { bytes: Buffer.from(bytes), size: bytes.length })
  })

  ipcMain.handle('workbook:save', async (event, input) => {
    assertTrustedSender(event)
    if (!input || typeof input.documentId !== 'string' || !input.workbook) throw new Error('Invalid save request.')
    const store = documentStore(event)
    const record = store.get(input.documentId)
    if (!record) throw new Error('This workbook is no longer attached to this window.')

    if (record.saving) throw new Error('This workbook is already saving. Please wait for it to finish.')
    record.saving = true
    try {
      let requestedFormat = saveFormat(input, record)
      const canOverwrite = !input.saveAs && record.path && extensionOf(record.path) === `.${requestedFormat}`
      let targetPath = canOverwrite ? record.path : null
      if (!targetPath) {
        const result = await dialog.showSaveDialog(callingWindow(event), {
          title: input.saveAs ? 'Save spreadsheet as' : 'Save spreadsheet',
          defaultPath: record.path ? path.join(path.dirname(record.path), safeSuggestedName(input.suggestedName || record.originalName, requestedFormat)) : safeSuggestedName(input.suggestedName || record.originalName, requestedFormat),
          filters: saveFilters(requestedFormat),
        })
        if (result.canceled || !result.filePath) return null
        requestedFormat = formatForPath(result.filePath, requestedFormat)
        targetPath = result.filePath.toLowerCase().endsWith(`.${requestedFormat}`)
          ? result.filePath
          : `${result.filePath}.${requestedFormat}`
      }

      const originalCopy = requestedFormat === record.sourceFormat && input.sourceUnmodified === true
      const baseBytes = originalCopy ? await unchangedSourceBytes(record, true) : requestedFormat === 'xlsx'
        ? await sourcePackageBytes(record).catch((error) => { if (input.saveAs) return null; throw error }) : null
      const bytes = originalCopy
        ? baseBytes
        : await serializeWorkbook(input.workbook, requestedFormat, {
            baseBytes,
            sourceFormat: record.sourceFormat,
          })
      const overwritesSource = record.path && path.resolve(targetPath).toLowerCase() === path.resolve(record.path).toLowerCase()
      if (overwritesSource) await assertSourceUnchanged(record)
      // Legacy edits pass through a real Office converter. Keep the original binary
      // once per open document so unmodelled legacy features remain recoverable.
      if (overwritesSource && ['xls', 'ods'].includes(record.sourceFormat) && !originalCopy && !record.backupPath) {
        const originalBytes = await unchangedSourceBytes(record, true)
        const backupDirectory = path.join(app.getPath('userData'), 'workbook-backups')
        await fs.mkdir(backupDirectory, { recursive: true })
        const backupPath = path.join(backupDirectory, `${path.basename(record.path, path.extname(record.path))}-${crypto.randomUUID()}.${record.sourceFormat}`)
        await fs.writeFile(backupPath, originalBytes, { flag: 'wx' })
        record.backupPath = backupPath
        await assertSourceUnchanged(record)
      }
      await atomicWrite(path.resolve(targetPath), bytes)
      record.path = path.resolve(targetPath)
      record.sourceFormat = requestedFormat
      record.originalName = path.basename(targetPath)
      const info = await fs.stat(record.path)
      record.sourceSnapshot = { path: record.path, size: info.size, modified: info.mtimeMs }
      return { path: record.path, name: record.originalName, format: requestedFormat, backupPath: record.backupPath }
    } finally {
      record.saving = false
    }
  })

  ipcMain.handle('workbook:export', async (event, input) => {
    assertTrustedSender(event)
    if (!input || typeof input.documentId !== 'string' || !input.workbook) throw new Error('Invalid export request.')
    const record = documentStore(event).get(input.documentId)
    if (!record) throw new Error('This workbook is no longer attached to this window.')
    const requestedFormat = normalizeExportFormat(input.format)
    const result = await dialog.showSaveDialog(callingWindow(event), {
      title: `Export spreadsheet as ${requestedFormat.toUpperCase()}`,
      defaultPath: safeSuggestedName(input.suggestedName || record.originalName, requestedFormat),
      filters: exportFilter(requestedFormat),
    })
    if (result.canceled || !result.filePath) return null
    const targetPath = path.resolve(result.filePath.toLowerCase().endsWith(`.${requestedFormat}`)
      ? result.filePath
      : `${result.filePath}.${requestedFormat}`)
    const baseBytes = requestedFormat === 'xlsx' ? await sourcePackageBytes(record) : null
    const generated = requestedFormat === 'xlsx' && record.sourceFormat === 'xlsx' && input.sourceUnmodified === true && baseBytes
      ? { format: requestedFormat, printDocument: null, bytes: baseBytes }
      : await createSpreadsheetExport(input, requestedFormat, {
          baseBytes,
          sourceFormat: record.sourceFormat,
        })
    let bytes = generated.bytes
    let temporaryDirectory = null
    try {
      if (requestedFormat === 'pdf') {
        temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-calc-export-'))
        bytes = await renderPrintDocumentToPdf(generated.printDocument, temporaryDirectory)
      }
      if (!bytes) throw new Error(`Could not create the ${requestedFormat.toUpperCase()} export.`)
      await atomicWrite(targetPath, bytes)
    } finally {
      if (temporaryDirectory) await fs.rm(temporaryDirectory, { recursive: true, force: true }).catch(() => {})
    }
    return {
      path: targetPath,
      name: path.basename(targetPath),
      format: requestedFormat,
      sheets: generated.printDocument?.sheetCount,
      cells: generated.printDocument?.printedCells,
    }
  })

  ipcMain.handle('workbook:print-preview', async (event, input) => {
    assertTrustedSender(event)
    if (!input || typeof input.documentId !== 'string' || !input.workbook) throw new Error('Invalid print request.')
    if (!documentStore(event).has(input.documentId)) throw new Error('This workbook is no longer attached to this window.')
    return printPreviewResult(createSpreadsheetPrintDocument(input))
  })

  ipcMain.handle('workbook:print', async (event, input) => {
    assertTrustedSender(event)
    if (!input || typeof input.documentId !== 'string' || !input.workbook) throw new Error('Invalid print request.')
    if (!documentStore(event).has(input.documentId)) throw new Error('This workbook is no longer attached to this window.')
    return printSpreadsheet(input, callingWindow(event))
  })

  ipcMain.handle('shell:show-item', (event, filePath) => {
    assertTrustedSender(event)
    if (typeof filePath !== 'string') throw new Error('Invalid path.')
    shell.showItemInFolder(path.resolve(filePath))
  })
  ipcMain.handle('shell:open-external', async (event, target) => {
    assertTrustedSender(event)
    if (typeof target !== 'string' || target.length > 8_192) throw new Error('Invalid link.')
    const url = new URL(target)
    if (!['http:', 'https:', 'mailto:'].includes(url.protocol)) throw new Error('Only web and email links can be opened.')
    await shell.openExternal(url.toString())
  })
  ipcMain.handle('app:get-version', (event) => { assertTrustedSender(event); return app.getVersion() })
  ipcMain.on('window:minimize', (event) => { assertTrustedSender(event); callingWindow(event)?.minimize() })
  ipcMain.on('window:toggle-maximize', (event) => {
    assertTrustedSender(event)
    const browserWindow = callingWindow(event)
    if (!browserWindow) return
    if (browserWindow.isMaximized()) browserWindow.unmaximize()
    else browserWindow.maximize()
  })
  ipcMain.on('window:confirm-close', (event) => {
    assertTrustedSender(event)
    const browserWindow = callingWindow(event)
    if (!browserWindow) return
    closeApprovedWindows.add(browserWindow)
    browserWindow.close()
  })
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) app.quit()
else {
  app.on('second-instance', (_event, argv) => {
    const incoming = supportedPaths(argv)
    if (incoming.length) {
      incoming.forEach((filePath) => createWindow(filePath))
      return
    }
    const browserWindow = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0]
    if (browserWindow) {
      if (browserWindow.isMinimized()) browserWindow.restore()
      browserWindow.focus()
    }
  })

  app.whenReady().then(async () => {
    await cleanupStalePrintDirectories()
    registerIpc()
    const incoming = supportedPaths(process.argv)
    if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
    else createWindow()
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow() })
  })
}

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
app.on('before-quit', () => { isQuitting = true })
