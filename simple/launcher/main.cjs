'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron')
const { EXTENSIONS_BY_MODE, MODES, supportedPaths } = require('../electron/routing.cjs')
// Opens files in their workspaces: one new Simple process per workspace,
// however many files (a long list travels in a list file, see
// electron/open-list.cjs), with the routing decision forwarded.
const { launchValidatedPaths } = require('./open-paths.cjs')
const { launchDetached, portableExecutablePath } = require('../electron/launch.cjs')
const formats = require('../shared/electron/formats.cjs')
const { registerAssociations, unregisterAssociations } = require('./associations.cjs')
const { atomicWrite, runCombine, unlockPdfBytes } = require('./combine-host.cjs')
const { CombineError, serializeError } = require('./combine-policy.cjs')
// Checks on the picked files and the save target, with plain file errors.
const { assertTargetIsNotASource, describeCombinePaths } = require('./combine-paths.cjs')
const { printHtmlToPdf } = require('../shared/electron/html-to-pdf.cjs')
const { sweep } = require('../shared/electron/io-core.cjs')
const { getOfficeEngineStatus } = require('../shared/electron/office-engine.cjs')

let mainWindow = null
let combineBusy = false

/** A failure the Combine window shows as text; never a raw IPC error. */
function combineFailure(error) {
  const { code, message } = serializeError(error)
  return { canceled: false, ok: false, code, message }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1040,
    height: 720,
    minWidth: 820,
    minHeight: 600,
    show: false,
    frame: false,
    title: 'simple',
    backgroundColor: '#f3f3f1',
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: !app.isPackaged,
    },
  })
  mainWindow.removeMenu()
  mainWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  mainWindow.webContents.on('will-navigate', (event, url) => {
    if (url !== mainWindow.webContents.getURL()) event.preventDefault()
  })
  mainWindow.loadFile(path.join(__dirname, 'index.html'))
  mainWindow.once('ready-to-show', () => mainWindow.show())
}

async function buildInfo() {
  try {
    const manifest = JSON.parse(await fs.readFile(path.join(__dirname, '..', 'modules', 'manifest.json'), 'utf8'))
    return { version: app.getVersion(), packaged: app.isPackaged, manifest, extensions: EXTENSIONS_BY_MODE }
  } catch {
    return { version: app.getVersion(), packaged: app.isPackaged, manifest: null, extensions: EXTENSIONS_BY_MODE }
  }
}

function registerIpc() {
  ipcMain.handle('launcher:info', buildInfo)
  ipcMain.handle('launcher:combine-add', async (event, paths) => {
    try {
      if (combineBusy) throw new CombineError('BUSY', 'Wait for the current PDF to finish.')
      if (paths !== undefined) return await describeCombinePaths(paths)
      // The registry lists what Combine opens on this PC; .doc and .xls only with a local office engine.
      const engine = Boolean((await getOfficeEngineStatus())?.available)
      const result = await dialog.showOpenDialog(BrowserWindow.fromWebContents(event.sender), {
        title: 'Add files to combine', properties: ['openFile', 'multiSelections'],
        filters: formats.dialogFilters('combine', { engine, allFiles: false }),
      })
      return result.canceled ? { entries: [], skipped: [] } : await describeCombinePaths(result.filePaths)
    } catch (error) {
      return { entries: [], skipped: [], ...combineFailure(error) }
    }
  })
  ipcMain.handle('launcher:combine-save', async (event, entries) => {
    if (combineBusy) return combineFailure(new CombineError('BUSY', 'Wait for the current PDF to finish.'))
    combineBusy = true
    try {
      if (!Array.isArray(entries) || entries.length < 2) throw new CombineError('INVALID', 'Add at least two files.')
      const { skipped } = await describeCombinePaths(entries.map((entry) => entry?.path))
      if (skipped.length) throw new CombineError(skipped[0].code, `${skipped[0].name}: ${skipped[0].message}`)
      const result = await dialog.showSaveDialog(BrowserWindow.fromWebContents(event.sender), {
        title: 'Save combined PDF', defaultPath: path.join(path.dirname(entries[0].path), 'Combined.pdf'),
        filters: [{ name: 'PDF document', extensions: ['pdf'] }],
      })
      if (result.canceled || !result.filePath) return { canceled: true }
      const target = path.extname(result.filePath) ? result.filePath : `${result.filePath}.pdf`
      if (path.extname(target).toLowerCase() !== '.pdf') throw new CombineError('INVALID_NAME', 'Choose a .pdf filename for the combined document.')
      await assertTargetIsNotASource(entries, target)
      const combined = await runCombine(entries, (progress) => {
        if (!event.sender.isDestroyed()) event.sender.send('launcher:combine-progress', progress)
      }, {
        printHtml: (html, options) => printHtmlToPdf(html, options),
        unlockPdf: (bytes) => unlockPdfBytes(bytes),
        title: path.basename(target, path.extname(target)),
      })
      await atomicWrite(target, combined.bytes)
      // The result opens in Simple's own PDF workspace, never in another app.
      // The PDF is saved either way; a failure to open it is reported, not thrown.
      let openFailure = null
      try { await launchDetached([target]) } catch (error) { openFailure = error?.message || "Simple couldn't open it." }
      return {
        canceled: false, ok: true, name: path.basename(target), pageCount: combined.pageCount,
        builtIn: Array.isArray(combined.builtIn) ? combined.builtIn : [],
        opened: !openFailure, ...(openFailure ? { openFailure } : {}),
      }
    } catch (error) {
      return combineFailure(error)
    } finally { combineBusy = false }
  })
  ipcMain.handle('launcher:open', async (event) => {
    const owner = BrowserWindow.fromWebContents(event.sender)
    const result = await dialog.showOpenDialog(owner, {
      title: 'Open with simple',
      properties: ['openFile', 'multiSelections'],
      // Every routed extension, grouped by workspace, from the format registry.
      filters: formats.dialogFilters('launcher', { allFiles: false }),
    })
    if (result.canceled) return { opened: 0, unsupported: 0, failed: 0 }
    return launchValidatedPaths(result.filePaths)
  })
  ipcMain.handle('launcher:launch-paths', (_event, paths) => {
    if (!Array.isArray(paths) || paths.some((item) => typeof item !== 'string')) throw new Error('Invalid file list.')
    return launchValidatedPaths(paths)
  })
  ipcMain.handle('launcher:launch-mode', async (_event, mode) => {
    if (!MODES.includes(mode)) throw new Error('Unknown workspace.')
    await launchDetached([`--simple-mode=${mode}`])
    return true
  })
  ipcMain.handle('launcher:register-file-types', async () => {
    if (!app.isPackaged) throw new Error('Build the portable app before registering file types.')
    return registerAssociations(path.resolve(portableExecutablePath()))
  })
  ipcMain.handle('launcher:unregister-file-types', unregisterAssociations)
  ipcMain.handle('launcher:open-default-apps', () => shell.openExternal('ms-settings:defaultapps'))
  ipcMain.on('window:minimize', (event) => BrowserWindow.fromWebContents(event.sender)?.minimize())
  ipcMain.on('window:close', (event) => BrowserWindow.fromWebContents(event.sender)?.close())
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (_event, argv, workingDirectory) => {
    const incoming = supportedPaths(argv)
    if (incoming.length) {
      launchValidatedPaths(incoming, { cwd: typeof workingDirectory === 'string' && workingDirectory ? workingDirectory : undefined })
        .catch((error) => process.stderr.write(`Could not open ${incoming.length} file(s): ${error && error.message}\n`))
    }
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.focus()
    }
  })
  app.whenReady().then(() => {
    app.setAppUserModelId('com.simple.unified')
    // Finish or undo a combined-PDF save that a crash interrupted. Only paths
    // recorded in the launcher's own save journal are ever touched.
    void sweep().catch(() => {})
    registerIpc()
    createWindow()
  })
}

app.on('window-all-closed', () => app.quit())
