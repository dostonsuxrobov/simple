'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron')
const { EXTENSIONS_BY_MODE, groupPathsByMode, MODES, supportedPaths } = require('../electron/routing.cjs')
const { launchDetached, portableExecutablePath } = require('../electron/launch.cjs')
const { registerAssociations, unregisterAssociations } = require('./associations.cjs')
const { atomicWrite, runCombine } = require('./combine-host.cjs')

let mainWindow = null
let combineBusy = false
const COMBINE_EXTENSIONS = ['pdf', 'docx', 'doc', 'xls', 'xlsx', 'ods', 'png', 'jpg', 'jpeg']

async function describeCombinePaths(paths) {
  if (!Array.isArray(paths) || paths.length > 100) throw new Error('Choose up to 100 files at a time.')
  return Promise.all(paths.map(async (filePath) => {
    if (typeof filePath !== 'string' || !path.isAbsolute(filePath) || !COMBINE_EXTENSIONS.includes(path.extname(filePath).slice(1).toLowerCase())) throw new Error('Combine accepts PDFs, Word documents, Excel and ODS spreadsheets, PNGs, and JPEGs.')
    const stat = await fs.stat(filePath)
    if (!stat.isFile() || !stat.size || stat.size > 256 * 1024 * 1024) throw new Error(`${path.basename(filePath)} must be a nonempty file smaller than 256 MB.`)
    return { path: filePath, name: path.basename(filePath), size: stat.size, pages: '' }
  }))
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

function launchValidatedPaths(paths) {
  const valid = supportedPaths(paths)
  const groups = groupPathsByMode(valid)
  for (const pathsForMode of groups.values()) launchDetached(pathsForMode)
  return { opened: valid.length, unsupported: Math.max(0, paths.length - valid.length) }
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
    if (combineBusy) throw new Error('Wait for the current PDF to finish.')
    if (paths !== undefined) return describeCombinePaths(paths)
    const result = await dialog.showOpenDialog(BrowserWindow.fromWebContents(event.sender), {
      title: 'Add files to combine', properties: ['openFile', 'multiSelections'],
      filters: [{ name: 'PDF, Word, spreadsheets, and photos', extensions: COMBINE_EXTENSIONS }],
    })
    return result.canceled ? [] : describeCombinePaths(result.filePaths)
  })
  ipcMain.handle('launcher:combine-save', async (event, entries) => {
    if (combineBusy) throw new Error('Wait for the current PDF to finish.')
    if (!Array.isArray(entries) || entries.length < 2) throw new Error('Add at least two files.')
    await describeCombinePaths(entries.map((entry) => entry?.path))
    combineBusy = true
    try {
      const result = await dialog.showSaveDialog(BrowserWindow.fromWebContents(event.sender), {
        title: 'Save combined PDF', defaultPath: path.join(path.dirname(entries[0].path), 'Combined.pdf'),
        filters: [{ name: 'PDF document', extensions: ['pdf'] }],
      })
      if (result.canceled || !result.filePath) return { canceled: true }
      const target = path.extname(result.filePath) ? result.filePath : `${result.filePath}.pdf`
      if (path.extname(target).toLowerCase() !== '.pdf') throw new Error('Choose a .pdf filename for the combined document.')
      const realTarget = await fs.realpath(target).catch(() => path.resolve(target))
      for (const entry of entries) {
        const realSource = await fs.realpath(entry.path)
        if (realSource.toLowerCase() === realTarget.toLowerCase()) throw new Error('Choose a new filename so the combined PDF keeps your source file intact.')
      }
      const combined = await runCombine(entries, (progress) => {
        if (!event.sender.isDestroyed()) event.sender.send('launcher:combine-progress', progress)
      })
      await atomicWrite(target, combined.bytes)
      launchDetached([target])
      return { canceled: false, name: path.basename(target), pageCount: combined.pageCount }
    } finally { combineBusy = false }
  })
  ipcMain.handle('launcher:open', async (event) => {
    const owner = BrowserWindow.fromWebContents(event.sender)
    const result = await dialog.showOpenDialog(owner, {
      title: 'Open with simple',
      properties: ['openFile', 'multiSelections'],
      filters: [
        { name: 'All supported files', extensions: Object.values(EXTENSIONS_BY_MODE).flat().map((extension) => extension.slice(1)) },
        { name: 'Documents', extensions: EXTENSIONS_BY_MODE.docs.map((extension) => extension.slice(1)) },
        { name: 'Spreadsheets', extensions: EXTENSIONS_BY_MODE.calc.map((extension) => extension.slice(1)) },
        { name: 'PDF and text', extensions: EXTENSIONS_BY_MODE.pdf.map((extension) => extension.slice(1)) },
        { name: 'Images', extensions: EXTENSIONS_BY_MODE.image.map((extension) => extension.slice(1)) },
        { name: 'Videos', extensions: EXTENSIONS_BY_MODE.video.map((extension) => extension.slice(1)) },
      ],
    })
    if (result.canceled) return { opened: 0, unsupported: 0 }
    return launchValidatedPaths(result.filePaths)
  })
  ipcMain.handle('launcher:launch-paths', (_event, paths) => {
    if (!Array.isArray(paths) || paths.some((item) => typeof item !== 'string')) throw new Error('Invalid file list.')
    return launchValidatedPaths(paths)
  })
  ipcMain.handle('launcher:launch-mode', (_event, mode) => {
    if (!MODES.includes(mode)) throw new Error('Unknown workspace.')
    launchDetached([`--simple-mode=${mode}`])
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
  app.on('second-instance', (_event, argv) => {
    const incoming = supportedPaths(argv)
    if (incoming.length) launchValidatedPaths(incoming)
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.focus()
    }
  })
  app.whenReady().then(() => {
    app.setAppUserModelId('com.simple.unified')
    registerIpc()
    createWindow()
  })
}

app.on('window-all-closed', () => app.quit())
