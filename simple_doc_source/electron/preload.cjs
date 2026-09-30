const { contextBridge, ipcRenderer, webUtils } = require('electron')

const pendingExternalOpenPaths = []
let externalOpenCallback = null

ipcRenderer.on('file:open-external', (_event, filePath) => {
  if (externalOpenCallback) externalOpenCallback(filePath)
  else pendingExternalOpenPaths.push(filePath)
})

contextBridge.exposeInMainWorld('simpleDocs', {
  getDocumentFonts: () => ipcRenderer.invoke('document:fonts'),
  getOriginalLayoutPdf: (input) => ipcRenderer.invoke('document:original-pdf', input),
  onDocumentShortcut: (callback) => {
    const listener = (_event, action) => callback(action)
    ipcRenderer.on('document:shortcut', listener)
    return () => ipcRenderer.removeListener('document:shortcut', listener)
  },
  openFile: () => ipcRenderer.invoke('file:open-dialog'),
  openPath: (filePath) => ipcRenderer.invoke('file:open-path', filePath),
  openBytes: (input) => ipcRenderer.invoke('file:open-bytes', input),
  openInNewWindow: (filePath) => ipcRenderer.invoke('file:open-in-new-window', filePath),
  newWindow: () => ipcRenderer.invoke('app:new-window'),
  saveDocx: (input) => ipcRenderer.invoke('file:save-docx', input),
  savePdf: (input) => ipcRenderer.invoke('file:save-pdf', input),
  saveExport: (input) => ipcRenderer.invoke('file:save-export', input),
  composePrintPdf: (input) => ipcRenderer.invoke('file:compose-print-pdf', input),
  listPrinters: () => ipcRenderer.invoke('print:list-printers'),
  printPdf: (input) => ipcRenderer.invoke('file:print-pdf', input),
  getRecents: () => ipcRenderer.invoke('recent:list'),
  removeRecent: (filePath) => ipcRenderer.invoke('recent:remove', filePath),
  saveRecovery: (input) => ipcRenderer.invoke('recovery:save', input),
  getRecoveries: () => ipcRenderer.invoke('recovery:list'),
  loadRecovery: (id) => ipcRenderer.invoke('recovery:load', id),
  clearRecovery: (id) => ipcRenderer.invoke('recovery:clear', id),
  pathForFile: (file) => {
    try { return webUtils.getPathForFile(file) || null } catch { return null }
  },
  setTitle: (title) => ipcRenderer.send('window:set-title', title),
  minimize: () => ipcRenderer.send('window:minimize'),
  toggleMaximize: () => ipcRenderer.send('window:toggle-maximize'),
  toggleFullscreen: () => ipcRenderer.send('window:toggle-fullscreen'),
  confirmClose: () => ipcRenderer.send('window:confirm-close'),
  openExternal: (url) => ipcRenderer.invoke('shell:open-external', url),
  onMaximized: (callback) => {
    const listener = (_event, maximized) => callback(maximized)
    ipcRenderer.on('window:maximized', listener)
    return () => ipcRenderer.removeListener('window:maximized', listener)
  },
  onFullscreen: (callback) => {
    const listener = (_event, fullscreen) => callback(fullscreen)
    ipcRenderer.on('window:fullscreen', listener)
    return () => ipcRenderer.removeListener('window:fullscreen', listener)
  },
  onOpenExternal: (callback) => {
    externalOpenCallback = callback
    while (pendingExternalOpenPaths.length) callback(pendingExternalOpenPaths.shift())
    return () => { if (externalOpenCallback === callback) externalOpenCallback = null }
  },
  onCloseRequested: (callback) => {
    const listener = () => callback()
    ipcRenderer.on('window:close-requested', listener)
    return () => ipcRenderer.removeListener('window:close-requested', listener)
  },
})
