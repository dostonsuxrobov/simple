const { app, BrowserWindow, dialog, ipcMain, shell, protocol, net } = require('electron')
const fs = require('node:fs/promises')
const { constants: fsConstants } = require('node:fs')
const path = require('node:path')
const os = require('node:os')
const crypto = require('node:crypto')
const { pathToFileURL } = require('node:url')
const {
  IMPORT_FORMATS, MAX_FILE_BYTES, MAX_IMPORT_BYTES, OPEN_DIALOG_FILTERS, UNSUPPORTED_FILE_MESSAGE,
  documentOpenKind, isOpenableDocumentPath, validateDocxBytes, wordPackageInfo,
} = require('./docx-files.cjs')
const { LEGACY_CONTENT_LABELS, sniffLegacyDocContent, validateLegacyDocBytes } = require('./legacy-doc.cjs')
const { convertLegacyDocToDocx } = require('./document-files.cjs')
const { needsOfficeLayout } = require('./document-layout.cjs')
const { availableExportFormats, exportFormat, validateExportContent } = require('./export-files.cjs')
const { isStalePrintDirectory, loadPdfForPrinting, printWebContentsSilently, PRINT_DIRECTORY_PREFIX } = require('./print-host.cjs')
const { composePrintPdf, nativePrintOptions, resolvePrinter, validatePdfBytes } = require('./print-layout.cjs')
const { createFontCatalog } = require('./document-fonts.cjs')
const { findOfficeConverter, convertOfficeBytes } = require('./office-converter.cjs')
const { dialogDefaultPath, isAbsoluteDocumentPath, safeStem, sameFilePath } = require('./document-paths.cjs')
const { proposedSiblingPath, writeBesideSource } = require('./sibling-save.cjs')
const { documentShortcutAction } = require('./shortcuts.cjs')
const { repairDocxPackage } = require('./docx-repair.cjs')
const { configureSpellingSession, createUserDictionary, createWindowsSpellHost, registerSpellingIpc } = require('./spelling.cjs')
const { registerSharedIo, bridgeArguments } = require('./simple-io/io-ipc.cjs')
const { sweep } = require('./simple-io/io-core.cjs')
const { safeWriteFile } = require('./simple-io/safe-write.cjs')
const guard = require('./simple-io/document-guard.cjs')
const stores = require('./simple-io/stores.cjs')
const officeEngine = require('./simple-io/office-engine.cjs')
const JSZip = require('jszip')

const IO_MODULE = 'docs'

protocol.registerSchemesAsPrivileged([{ scheme: 'simple-font', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true } }])
// Installed fonts, enumerated locally: the curated set every window loads, any other
// installed family a window asks for (recent and document fonts), and the font list.
const fontCatalog = createFontCatalog()


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

// The document each window is showing, as main itself opened or saved it.
// Dialog defaults and the beside-the-original save never trust a renderer path
// for where an original lives.
const documentContexts = new WeakMap()

function rememberDocument(sender, context) {
  if (!sender) return
  const sourcePath = isAbsoluteDocumentPath(context?.sourcePath) ? path.resolve(context.sourcePath) : null
  const boundPath = isAbsoluteDocumentPath(context?.path) ? path.resolve(context.path) : null
  documentContexts.set(sender, { sourcePath, path: boundPath, kind: context?.kind || null })
}

function documentContext(sender) {
  return (sender && documentContexts.get(sender)) || { sourcePath: null, path: null, kind: null }
}

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex')
}

/** name.docx beside the source, then "name (edited).docx", never an existing file. */
function proposedSiblingDocx(sourcePath) {
  return proposedSiblingPath(sourcePath, '.docx')
}

function ensureExtension(filePath, extension) {
  return filePath.toLowerCase().endsWith(extension) ? filePath : `${filePath}${extension}`
}

// Every user-visible write goes through the shared verified write: temp file in the
// same folder, flush, read-back check, then replace with retries while another
// program holds the file. Failures arrive as coded errors with plain messages.
async function atomicWrite(targetPath, data) {
  await safeWriteFile(targetPath, toBytes(data))
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

/**
 * Reads a file for the editor. Its format comes from the name, corrected by the bytes
 * (an RTF or a DOCX saved under another name opens as what it is). Word packages are
 * validated here; RTF, OpenDocument, web pages, Markdown and plain text go to the
 * renderer's native importers as they are (`importFormat`). Only a plain .docx is bound
 * to its path for in-place saving: everything else saves as a separate .docx.
 */
async function loadDocumentBytes(data, name, sourcePath = null) {
  const fileName = path.basename(String(name || ''))
  const bytes = toBytes(data)
  if (!bytes.byteLength) throw new Error('This file is empty.')
  if (bytes.byteLength > MAX_FILE_BYTES) throw new Error('This file is too large to open as a document.')
  const { kind, renamed } = documentOpenKind(fileName, bytes)

  if (kind === 'doc') {
    const payload = await loadLegacyDocBytes(bytes, fileName, sourcePath)
    if (!renamed) return payload
    // A Word 97-2003 file under another name is never overwritten with .doc bytes.
    return {
      ...payload,
      name: `${safeStem(fileName)}.docx`,
      path: null,
      format: 'docx',
      sourceData: undefined,
      requiresSaveAs: true,
      conversionWarnings: payload.conversionWarnings?.length ? payload.conversionWarnings : ['This Word 97–2003 document has another file extension. Saving creates a .docx file; the original stays unchanged.'],
    }
  }
  if (IMPORT_FORMATS.has(kind)) {
    if (bytes.byteLength > MAX_IMPORT_BYTES) throw new Error('This file is too large to open as a document.')
    return {
      data: new Uint8Array(bytes),
      name: fileName,
      path: null,
      format: 'docx',
      sourcePath,
      sourceHash: sourcePath ? sha256(bytes) : undefined,
      size: bytes.byteLength,
      importFormat: kind,
      convertedFrom: kind,
      requiresSaveAs: true,
    }
  }
  return loadWordPackageBytes(bytes, fileName, sourcePath, kind, renamed)
}

/** Word 97-2003 (.doc), including RTF, web pages and DOCX packages saved with a .doc name. */
async function loadLegacyDocBytes(bytes, fileName, sourcePath) {
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
  // RTF, HTML, web archives and Word 2003 XML saved with a .doc name open
  // natively, as in Word. They are never overwritten with DOCX bytes.
  const content = sniffLegacyDocContent(bytes)
  if (content && Object.hasOwn(LEGACY_CONTENT_LABELS, content)) {
    const converted = await convertLegacyDocToDocx(bytes, { title: fileName })
    return {
      data: new Uint8Array(converted.data),
      name: `${safeStem(fileName)}.docx`,
      path: null,
      format: 'docx',
      sourcePath,
      sourceHash: sourcePath ? sha256(bytes) : undefined,
      size: converted.data.byteLength,
      convertedFrom: content,
      conversionMethod: content,
      requiresSaveAs: true,
      conversionWarnings: converted.warnings,
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
    sourceHash: sourcePath ? sha256(bytes) : undefined,
    sourcePath,
    size: converted.data.byteLength,
    convertedFrom: 'doc',
    conversionMethod: converted.conversionMethod,
    ...(converted.conversionMethod === 'layout' ? { originalLayout: { data: new Uint8Array(bytes), extension: 'doc' } } : {}),
    requiresSaveAs: converted.conversionMethod !== 'layout' || !sourcePath,
    conversionWarnings: converted.warnings,
  }
}

/**
 * Anchored objects and floating tables need a full Office page-layout engine. When one
 * is installed, the original rendered view is kept rather than pretending canvas
 * reflow has retained their positions.
 */
async function officeLayoutFor(validated) {
  if (!await findOfficeConverter()) return undefined
  try {
    const zip = await JSZip.loadAsync(validated)
    const parts = Object.keys(zip.files).filter((name) => /^word\/(document|header\d+|footer\d+)\.xml$/.test(name))
    for (const name of parts) {
      const xml = await zip.file(name)?.async('string')
      if (xml && needsOfficeLayout(xml, name)) return { data: new Uint8Array(validated), extension: 'docx' }
    }
  } catch { /* The validated document can still open in the editor. */ }
  return undefined
}

/**
 * .docx, .docm, .dotx and .dotm packages (also under another name). A template opens
 * as a new untitled document; a macro-enabled or misnamed package keeps its original
 * and saves as a separate .docx. Macros are never run or kept.
 */
async function loadWordPackageBytes(bytes, fileName, sourcePath, kind, renamed) {
  const validated = validateDocxBytes(bytes)
  const { macros } = wordPackageInfo(validated)
  const originalLayout = await officeLayoutFor(validated)
  const data = new Uint8Array(validated)
  if (kind === 'dotx' || kind === 'dotm') {
    return { data, name: fileName, path: null, sourcePath: null, size: validated.byteLength, requiresSaveAs: true, template: true, macros: macros || kind === 'dotm', originalLayout }
  }
  if (kind === 'docm' || renamed) {
    return {
      data,
      name: renamed ? `${safeStem(fileName)}.docx` : fileName,
      path: null,
      format: 'docx',
      sourcePath,
      sourceHash: sourcePath ? sha256(bytes) : undefined,
      size: validated.byteLength,
      convertedFrom: renamed ? 'docx-renamed' : 'docm',
      macros: macros || kind === 'docm',
      requiresSaveAs: true,
      originalLayout,
    }
  }
  return {
    data,
    name: fileName,
    path: sourcePath,
    sourcePath,
    size: validated.byteLength,
    requiresSaveAs: !sourcePath,
    ...(macros ? { macros: true } : {}),
    originalLayout,
    sourceHash: sourcePath ? sha256(bytes) : undefined,
  }
}

async function loadDocument(filePath, sender = null) {
  const resolved = path.resolve(filePath)
  const stat = await fs.stat(resolved)
  if (!stat.isFile()) throw new Error(UNSUPPORTED_FILE_MESSAGE)
  if (!stat.size) throw new Error('This file is empty.')
  if (stat.size > MAX_FILE_BYTES) throw new Error('This file is too large to open as a document.')
  const payload = await loadDocumentBytes(await fs.readFile(resolved), path.basename(resolved), resolved)
  // A template is only a starting point: dialogs and saves never point back at it.
  rememberDocument(sender, {
    sourcePath: payload.template ? null : resolved,
    path: payload.path,
    kind: payload.template ? 'template' : payload.convertedFrom === 'doc' ? `doc-${payload.conversionMethod}` : payload.convertedFrom || payload.format || 'docx',
  })
  await addRecent(resolved)
  return payload
}

/**
 * Throws unless a path is a file Simple Docs can open: a known extension, or (chosen
 * through "All files") content it recognizes.
 */
async function assertOpenableFile(filePath) {
  const stat = await fs.stat(filePath)
  if (!stat.isFile()) throw new Error(UNSUPPORTED_FILE_MESSAGE)
  if (isOpenableDocumentPath(filePath)) return
  if (!stat.size) throw new Error('This file is empty.')
  if (stat.size > MAX_FILE_BYTES) throw new Error('This file is too large to open as a document.')
  documentOpenKind(path.basename(filePath), await fs.readFile(filePath))
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
      additionalArguments: bridgeArguments(IO_MODULE),
    },
  })
  // Save / Don't Save / Cancel on every close path (title bar, Alt+F4, taskbar, quit,
  // Windows sign-out), never closing during a save, and crash/hang handling.
  guard.installWindowGuard(window)
  window.removeMenu()
  window.webContents.on('before-input-event', (event, input) => {
    // Shortcuts follow the physical key on Cyrillic, Greek, Hebrew, Arabic and
    // other non-Latin layouts, where `key` is the layout's own character.
    const action = documentShortcutAction(input)
    if (!action) return
    event.preventDefault()
    if (action === 'new') createWindow()
    else window.webContents.send('document:shortcut', action)
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
  return window
}

// Spell checking stays on this machine: the Windows Spell Checking API through a hidden
// helper, Chromium's Hunspell downloader pointed at a local folder, and the user
// dictionary (Add to dictionary / Ignore all) in userData.
const spellHost = createWindowsSpellHost()

function setupSpelling() {
  const { session } = require('electron')
  configureSpellingSession(session.defaultSession, statePath('offline-dictionaries'))
  registerSpellingIpc(ipcMain, {
    host: spellHost,
    dictionary: createUserDictionary({
      read: () => readJson('spelling-dictionary.json', null),
      write: (value) => writeJson('spelling-dictionary.json', value),
    }),
    session: session.defaultSession,
    preferredLanguages: () => (typeof app.getPreferredSystemLanguages === 'function' ? app.getPreferredSystemLanguages() : [app.getLocale()]),
    broadcast: (channel, value) => {
      for (const window of BrowserWindow.getAllWindows()) if (!window.isDestroyed()) window.webContents.send(channel, value)
    },
  })
  app.on('will-quit', () => spellHost.dispose())
}

function registerIpc() {
  ipcMain.handle('document:fonts', async (_event, options) => {
    const families = Array.isArray(options?.families) ? options.families.filter((family) => typeof family === 'string').slice(0, 64) : []
    return fontCatalog.documentFonts(families)
  })
  ipcMain.handle('document:font-families', async () => fontCatalog.familyList())
  ipcMain.handle('document:original-pdf', async (_event, input) => {
    const extension = input?.extension
    if (!['doc', 'docx'].includes(extension)) throw new Error('Unsupported original document format.')
    const bytes = toBytes(input?.data)
    if (!bytes.length || bytes.length > MAX_FILE_BYTES) throw new Error('The original document is empty or too large.')
    return new Uint8Array(validatePdfBytes(await convertOfficeBytes({ bytes, inputExtension: extension, outputExtension: 'pdf', filter: 'writer_pdf_Export' })))
  })
  ipcMain.handle('file:open-dialog', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Open a document',
      properties: ['openFile'],
      filters: OPEN_DIALOG_FILTERS,
    })
    if (result.canceled || !result.filePaths[0]) return null
    return loadDocument(result.filePaths[0], event.sender)
  })
  ipcMain.handle('file:open-path', (event, filePath) => loadDocument(filePath, event.sender))
  ipcMain.handle('file:open-bytes', async (event, input) => {
    const payload = await loadDocumentBytes(input?.data, input?.name)
    rememberDocument(event.sender, {})
    return payload
  })
  ipcMain.handle('file:open-in-new-window', async (event, requestedPath = null) => {
    let selectedPath = requestedPath
    if (!selectedPath) {
      const result = await showOpenDialogFor(event, {
        title: 'Open a document', properties: ['openFile'], filters: OPEN_DIALOG_FILTERS,
      })
      if (result.canceled || !result.filePaths[0]) return false
      selectedPath = result.filePaths[0]
    }
    if (typeof selectedPath !== 'string' || !path.isAbsolute(selectedPath)) throw new Error(UNSUPPORTED_FILE_MESSAGE)
    await assertOpenableFile(selectedPath)
    createWindow(path.resolve(selectedPath))
    return true
  })
  ipcMain.handle('app:new-window', () => { createWindow(); return true })
  ipcMain.handle('file:save-docx', async (event, input) => {
    const format = input.format === 'doc' && !input.forceDialog ? 'doc' : 'docx'
    const extension = `.${format}`
    let data = validateDocxBytes(input.data)
    let warnings = []
    if (format === 'doc' && !input.sourceData && !await findOfficeConverter()) {
      throw new Error('Edits to this Word 97–2003 (.doc) file cannot be written in the .doc format on this computer because no local Office engine is installed. Save a .docx copy beside the original instead; the original .doc stays unchanged.')
    }
    let target = input.forceDialog ? null : input.path
    if (!target) {
      const context = documentContext(event.sender)
      const source = context.sourcePath
      // A converted original (.doc, RTF, HTML) gets a free name.docx beside it.
      const defaultPath = source && path.extname(source).toLowerCase() !== extension && extension === '.docx'
        ? await proposedSiblingDocx(source).catch(() => dialogDefaultPath({ name: input.name, extension, folderOf: [source] }))
        : dialogDefaultPath({ name: input.name, extension, folderOf: [input.path, context.path, source] })
      const result = await showSaveDialogFor(event, {
        title: input.forceDialog ? 'Save document as' : 'Save document',
        defaultPath,
        filters: [{ name: format === 'doc' ? 'Word 97–2003 document' : 'Word document', extensions: [format] }],
      })
      if (result.canceled || !result.filePath) return null
      target = ensureExtension(result.filePath, extension)
    }
    target = ensureExtension(path.resolve(target), extension)
    if (format === 'doc') data = input.sourceData
      ? validateLegacyDocBytes(input.sourceData)
      : validateLegacyDocBytes(await convertOfficeBytes({ bytes: data, inputExtension: 'docx', outputExtension: 'doc', filter: 'MS Word 97' }))
    else if (!input.expectedHash || sha256(data) !== input.expectedHash) {
      // Never alter bytes identical to the file on disk; repair edited packages.
      const repaired = await repairDocxPackage(data)
      data = repaired.bytes
      warnings = repaired.warnings
    }
    if (!input.forceDialog && input.path && input.expectedHash) {
      const existing = await fs.readFile(target).catch(() => null)
      if (!existing || sha256(existing) !== input.expectedHash) throw new Error('This file changed outside Simple Docs. Use Save as to keep both versions.')
    }
    if (!input.forceDialog && input.path && input.protectOriginal) {
      const backupPath = path.join(path.dirname(target), `${path.basename(target, path.extname(target))}.before-simple-edit${path.extname(target)}`)
      try { await fs.copyFile(target, backupPath, fsConstants.COPYFILE_EXCL) } catch (error) { if (!['ENOENT', 'EEXIST'].includes(error.code)) throw error }
    }
    await atomicWrite(target, data)
    rememberDocument(event.sender, { sourcePath: target, path: target, kind: format })
    await addRecent(target)
    return { path: target, name: path.basename(target), format, sourceHash: sha256(data), ...(format === 'doc' ? { sourceData: new Uint8Array(data) } : {}), ...(warnings.length ? { warnings } : {}) }
  })
  // Legacy .doc (and RTF/HTML saved as .doc) without a local Office engine:
  // the renderer asks the user, then edits go to name.docx (or
  // "name (edited).docx") beside the original, which is never written.
  ipcMain.handle('file:sibling-docx-plan', async (event) => {
    const context = documentContext(event.sender)
    if (!context.sourcePath) return null
    const proposed = await proposedSiblingDocx(context.sourcePath)
    return { sourcePath: context.sourcePath, path: proposed, name: path.basename(proposed), officeEngine: Boolean(await findOfficeConverter()) }
  })
  ipcMain.handle('file:save-beside-source', async (event, input) => {
    const context = documentContext(event.sender)
    const sourcePath = context.sourcePath
    if (!sourcePath) throw new Error('Simple Docs does not know where the original document is. Use Save as to choose a location.')
    if (input?.sourcePath && !sameFilePath(input.sourcePath, sourcePath)) throw new Error('This window’s original document has changed. Use Save as to choose a location.')
    const repaired = await repairDocxPackage(validateDocxBytes(input?.data))
    const target = await writeBesideSource(sourcePath, repaired.bytes, atomicWrite)
    rememberDocument(event.sender, { sourcePath: target, path: target, kind: 'docx' })
    await addRecent(target)
    return {
      path: target,
      name: path.basename(target),
      format: 'docx',
      sourceHash: sha256(repaired.bytes),
      originalPath: sourcePath,
      ...(repaired.warnings.length ? { warnings: repaired.warnings } : {}),
    }
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
  ipcMain.handle('export:formats', async () => availableExportFormats({ officeEngine: Boolean(await findOfficeConverter()) }))
  ipcMain.handle('file:save-export', async (event, input) => {
    const definition = exportFormat(input?.format)
    let data = validateExportContent(definition.id, input?.data)
    let warnings = []
    if (definition.id === 'doc' && !await findOfficeConverter()) {
      throw new Error('Word 97–2003 (.doc) export needs a local Office engine, and none is installed on this computer. Export a Word document (.docx) instead.')
    }
    if (definition.id === 'docx' || definition.id === 'doc') {
      const repaired = await repairDocxPackage(validateDocxBytes(data))
      data = repaired.bytes
      warnings = repaired.warnings
    }
    const context = documentContext(event.sender)
    const result = await showSaveDialogFor(event, {
      title: `Export as ${definition.label}`,
      defaultPath: dialogDefaultPath({ name: input?.name, extension: definition.extension, folderOf: [input?.path, context.path, context.sourcePath] }),
      filters: [{ name: definition.label, extensions: [definition.extension.slice(1)] }],
    })
    if (result.canceled || !result.filePath) return null
    const target = ensureExtension(path.resolve(result.filePath), definition.extension)
    if (definition.id === 'doc') data = validateLegacyDocBytes(await convertOfficeBytes({ bytes: data, inputExtension: 'docx', outputExtension: 'doc', filter: 'MS Word 97' }))
    await atomicWrite(target, data)
    return { path: target, name: path.basename(target), ...(warnings.length ? { warnings } : {}) }
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
  ipcMain.handle('recovery:load', async (event, id) => {
    const safeId = String(id).replace(/[^a-zA-Z0-9-]/g, '')
    const recoveries = await readJson('recoveries.json', [])
    const metadata = recoveries.find((item) => item.id === safeId)
    if (!metadata) throw new Error('Recovery copy not found.')
    const data = validateDocxBytes(await fs.readFile(statePath(path.join('recovery', `${safeId}.docx`))))
    rememberDocument(event.sender, { sourcePath: metadata.sourcePath || null, path: null, kind: 'recovery' })
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
    // The window guard asks the page again; after Save or Discard it reports no changes.
    window.close()
  })
}

function incomingDocumentPaths(argv) {
  return [...new Set(argv.filter((argument) => isOpenableDocumentPath(argument)).map((argument) => path.resolve(argument)))]
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
        ? await fontCatalog.fileForToken(url.pathname.slice(1)) : null
      if (!filePath) return new Response('Font not found', { status: 404 })
      const response = await net.fetch(pathToFileURL(filePath).href)
      return new Response(response.body, { headers: { 'Content-Type': /\.otf$/i.test(filePath) ? 'font/otf' : 'font/ttf', 'Access-Control-Allow-Origin': '*', 'Cache-Control': 'public, max-age=86400' } })
    })
    registerIpc()
    setupSpelling()
    registerSharedIo({
      ipcMain,
      module: IO_MODULE,
      guard,
      stores,
      officeEngine,
      openInWindow: (filePath) => { createWindow(filePath); return true },
    })
    // Finish or undo any save a crash interrupted before documents open.
    void sweep().catch((error) => console.error('[simple-io] sweep failed', error))
      .then(() => cleanupStalePrintDirectories()).finally(() => {
      const incoming = incomingDocumentPaths(process.argv)
      if (incoming.length) incoming.forEach((filePath) => createWindow(filePath))
      else createWindow()
    })
  })
}

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
