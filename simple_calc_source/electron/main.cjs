const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron')
const { registerClipboardHandlers } = require('./clipboard-bridge.cjs')
const { registerProtectionHandlers } = require('./protection.cjs')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const {
  SUPPORTED_EXTENSIONS,
  importWorkbookPath,
  importWorkbookBytes,
  serializeWorkbook,
  delimitedDialectFor,
  supportedOrExtensionless,
} = require('./workbooks.cjs')
const { createSpreadsheetPrintDocument } = require('./spreadsheet-print.cjs')
const { createSpreadsheetExport, exportFilter, normalizeExportFormat, exportLosses } = require('./spreadsheet-export.cjs')
const { ensurePrinterInstalled, printFailureMessage, printOptionsForDocument, registerPrinterHandlers } = require('./default-printer.cjs')
const { saveFilters, hasOriginalBytes, unchangedSourceBytes, assertSourceUnchanged, saveDocument } = require('./workbook-save.cjs')
const { officeEngineAvailable } = require('./office-converter.cjs')
const { registerSharedIo, bridgeArguments } = require('./simple-io/io-ipc.cjs')
const { sweep } = require('./simple-io/io-core.cjs')
const { safeWriteFile } = require('./simple-io/safe-write.cjs')
const guard = require('./simple-io/document-guard.cjs')
const stores = require('./simple-io/stores.cjs')
const officeEngine = require('./simple-io/office-engine.cjs')

const IO_MODULE = 'calc'

app.setName('simple_calc')

const MAX_OPEN_BYTES = 512 * 1024 * 1024
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
  // Files without an extension are opened by content (CSV exports often have none).
  if (!supportedOrExtensionless(filePath)) throw new Error('This spreadsheet format is not supported.')
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

/**
 * A window shows one document at a time. When it opens or creates another one, the source
 * bytes held by the previous records (merge base, dropped-file bytes) are released so memory
 * does not grow with every file opened; the records themselves stay valid.
 */
function releaseReplacedDocuments(event) {
  for (const record of documentStore(event).values()) {
    if (record.saving) continue
    record.mergeBase = null
    if (record.sourceSnapshot && Buffer.isBuffer(record.sourceSnapshot.bytes) && record.path) delete record.sourceSnapshot.bytes
  }
}

/**
 * Register an opened document. `mergeBase` is the OOXML package the model was imported from;
 * it stays the base of every XLSX save of this document (see workbook-save.cjs), so imported
 * charts, pictures and worksheets are always copied from their own original parts.
 */
function rememberDocument(event, payload, sourcePath = null, sourceSnapshot = null, extras = {}) {
  const documentId = crypto.randomUUID()
  const metadata = payload.workbook && payload.workbook.metadata
  releaseReplacedDocuments(event)
  documentStore(event).set(documentId, {
    path: sourcePath,
    sourceFormat: payload.sourceFormat,
    originalName: payload.name,
    sourceSnapshot,
    mergeBase: Buffer.isBuffer(extras.mergeBase) ? extras.mergeBase : null,
    dialect: metadata && metadata.dialect && typeof metadata.dialect === 'object' ? metadata.dialect : null,
  })
  return { ...payload, documentId, path: sourcePath }
}

async function readWorkbookPath(event, requestedPath) {
  const filePath = assertSupportedPath(requestedPath)
  const info = await fs.stat(filePath)
  if (!info.isFile()) throw new Error('The selected item is not a file.')
  if (info.size > MAX_OPEN_BYTES) throw new Error('This workbook is larger than the 512 MB safety limit.')
  const imported = await importWorkbookPath(filePath)
  return rememberDocument(event, imported.payload, filePath, {
    path: filePath,
    size: imported.stat.size,
    modified: imported.stat.mtimeMs,
  }, { mergeBase: imported.mergeBase })
}

// Every user-visible write goes through the shared verified write: temp file in the
// same folder, flush, read-back check, then replace with retries while another
// program (Excel, a backup tool, an antivirus scan) holds the file.
async function atomicWrite(targetPath, data) {
  const bytes = data && data.data && !Buffer.isBuffer(data) ? data.data : data
  await safeWriteFile(targetPath, bytes)
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
    // The printer, copies and collation chosen in the print dialog (printDocument.printJob);
    // without a choice the job goes to the Windows default printer.
    await ensurePrinterInstalled(printWindow.webContents, printDocument.printJob)
    const outcome = await new Promise((resolve, reject) => {
      printWindow.webContents.print(printOptionsForDocument(printDocument), (success, failureReason) => {
        if (success) {
          resolve({ printed: true, canceled: false })
          return
        }
        reject(new Error(printFailureMessage(failureReason, printDocument.printJob)))
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
      additionalArguments: bridgeArguments(IO_MODULE),
    },
  })
  // Save / Don't Save / Cancel on every close path, never closing during a save,
  // and crash/hang/sign-out handling.
  guard.installWindowGuard(browserWindow)

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
  // The print dialog's printer list (workbook:list-printers).
  registerPrinterHandlers(ipcMain, { assertTrustedSender })
  ipcMain.handle('workbook:create', (event) => {
    assertTrustedSender(event)
    const documentId = crypto.randomUUID()
    releaseReplacedDocuments(event)
    documentStore(event).set(documentId, { path: null, sourceFormat: 'xlsx', originalName: 'Untitled.xlsx' })
    return { documentId }
  })

  ipcMain.handle('workbook:new-window', async (event, requestedPath = null) => {
    assertTrustedSender(event)
    let selectedPath = requestedPath
    // A path sent by the page comes from a link in a workbook: network shares are never opened
    // that way (reading one connects to another machine and can send it the Windows sign-in).
    if (typeof selectedPath === 'string' && /^[\\/]{2}/.test(selectedPath.trim()) && !/^[\\/]{2}[?.][\\/][a-z]:[\\/]/i.test(selectedPath.trim())) {
      throw new Error('Links to network locations are not opened. Use File > Open to open that file.')
    }
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
    if (!supportedOrExtensionless(input.name)) throw new Error('This spreadsheet format is not supported.')
    const imported = await importWorkbookBytes(input.name, bytes)
    // The dropped bytes are both the original (for an unchanged copy) and the merge base.
    return rememberDocument(event, imported.payload, null, { bytes: imported.mergeBase || Buffer.from(bytes), size: bytes.length }, { mergeBase: imported.mergeBase })
  })

  // What this machine can write: without the local document engine, XLS/ODS saves of edited
  // files go to an .xlsx next to the original, and XLS/ODS exports use the basic writers.
  ipcMain.handle('workbook:capabilities', async (event) => {
    assertTrustedSender(event)
    return { officeEngine: await officeEngineAvailable() }
  })

  // Features a basic XLS/ODS export would lose (empty when nothing is lost or the engine is present).
  ipcMain.handle('workbook:export-check', async (event, input) => {
    assertTrustedSender(event)
    if (!input || !input.workbook) throw new Error('Invalid export request.')
    const format = normalizeExportFormat(input.format)
    const officeEngine = await officeEngineAvailable()
    return { format, officeEngine, losses: exportLosses(input.workbook, format, { officeEngine }), confirmationRequired: format === 'xls' && !officeEngine }
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
      const officeEngine = await officeEngineAvailable()
      return await saveDocument(record, input, {
        officeEngine,
        serialize: serializeWorkbook,
        writeFile: atomicWrite,
        dialectFor: delimitedDialectFor,
        backupDirectory: path.join(app.getPath('userData'), 'workbook-backups'),
        chooseSavePath: async ({ format, suggestedName, directory, saveAs }) => {
          const result = await dialog.showSaveDialog(callingWindow(event), {
            title: saveAs ? 'Save spreadsheet as' : 'Save spreadsheet',
            defaultPath: directory ? path.join(directory, safeSuggestedName(suggestedName, format)) : safeSuggestedName(suggestedName, format),
            filters: saveFilters(format, { officeEngine }),
          })
          return result.canceled || !result.filePath ? null : result.filePath
        },
      })
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
    const officeEngine = ['xls', 'ods'].includes(requestedFormat) ? await officeEngineAvailable() : true
    const acceptLoss = input.acceptLoss === true || input.valuesOnly === true
    // Ask before the file name is chosen: a values-only XLS needs the user's consent.
    if (requestedFormat === 'xls' && !officeEngine && !acceptLoss) {
      const losses = exportLosses(input.workbook, 'xls', { officeEngine })
      throw new Error(`Excel 97-2003 export without the document engine keeps values only${losses.length ? ` and would lose: ${losses.join('; ')}` : ''}. Export as XLSX to keep everything.`)
    }
    const result = await dialog.showSaveDialog(callingWindow(event), {
      title: `Export spreadsheet as ${requestedFormat.toUpperCase()}`,
      defaultPath: record.path ? path.join(path.dirname(record.path), safeSuggestedName(input.suggestedName || record.originalName, requestedFormat)) : safeSuggestedName(input.suggestedName || record.originalName, requestedFormat),
      filters: exportFilter(requestedFormat),
    })
    if (result.canceled || !result.filePath) return null
    const targetPath = path.resolve(result.filePath.toLowerCase().endsWith(`.${requestedFormat}`)
      ? result.filePath
      : `${result.filePath}.${requestedFormat}`)
    const overwritesSource = Boolean(record.path) && targetPath.toLowerCase() === path.resolve(record.path).toLowerCase()
    // Exporting over the document's own file must not discard changes made outside Simple
    // (a file that no longer exists has nothing to lose).
    if (overwritesSource) await assertSourceUnchanged(record).catch((error) => { if (error && error.code !== 'ENOENT') throw error })
    // An unedited XLSX is exported as the current file; otherwise the model is written over the
    // package it was imported from (never over a file a previous save produced).
    const unchangedBytes = requestedFormat === 'xlsx' && record.sourceFormat === 'xlsx' && input.sourceUnmodified === true && hasOriginalBytes(record)
      ? await unchangedSourceBytes(record).catch(() => null)
      : null
    const notes = []
    const generated = unchangedBytes
      ? { format: requestedFormat, printDocument: null, bytes: unchangedBytes }
      : await createSpreadsheetExport(input, requestedFormat, {
          baseBytes: requestedFormat === 'xlsx' ? record.mergeBase : null,
          sourceFormat: record.sourceFormat,
          officeEngine,
          acceptLoss,
          warnings: notes,
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
    if (overwritesSource) {
      // The document's own file was replaced: refresh the snapshot so the next save does not
      // report a change made outside Simple.
      const info = await fs.stat(targetPath)
      record.sourceSnapshot = { path: record.path, size: info.size, modified: info.mtimeMs }
    }
    return {
      path: targetPath,
      name: path.basename(targetPath),
      format: requestedFormat,
      sheets: generated.printDocument?.sheetCount,
      cells: generated.printDocument?.printedCells,
      ...(notes.length ? { notes: [...new Set(notes)] } : {}),
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
    if (typeof target !== 'string' || !target.trim() || target.length > 8_192) throw new Error('This link is empty or too long to open.')
    // Places in the workbook ("#Sheet2!A1") are followed by the window itself; anything that is
    // not a web or email address is refused with a message instead of a URL TypeError.
    if (target.trim().startsWith('#')) throw new Error('This link points to a place in the workbook, not a web page.')
    let url
    try {
      url = new URL(target.trim())
    } catch {
      throw new Error(`This link isn’t a valid web address: ${target.trim().slice(0, 200)}`)
    }
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
    // The window guard asks the page again; after Save or Discard it reports no changes.
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
    registerSharedIo({
      ipcMain,
      module: IO_MODULE,
      guard,
      stores,
      officeEngine,
      openInWindow: (filePath) => { createWindow(filePath); return true },
    })
    // Finish or undo any save a crash interrupted before workbooks open.
    await sweep().catch((error) => console.error('[simple-io] sweep failed', error))
    const incoming = supportedPaths(process.argv)
    if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
    else createWindow()
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow() })
  })
}

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
app.on('before-quit', () => { isQuitting = true })
