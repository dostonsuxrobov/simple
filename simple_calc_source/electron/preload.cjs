const { contextBridge, ipcRenderer } = require('electron')

const pendingExternalOpenPaths = []
let externalOpenCallback = null

ipcRenderer.on('file:open-external', (_event, filePath) => {
  if (externalOpenCallback) externalOpenCallback(filePath)
  else pendingExternalOpenPaths.push(filePath)
})

contextBridge.exposeInMainWorld('simpleCalc', {
  createWorkbook: () => ipcRenderer.invoke('workbook:create'),
  openWorkbook: () => ipcRenderer.invoke('workbook:open-dialog'),
  openPath: (filePath) => ipcRenderer.invoke('workbook:open-path', filePath),
  openBytes: (name, data) => ipcRenderer.invoke('workbook:open-bytes', { name, data }),
  openInNewWindow: (filePath) => ipcRenderer.invoke('workbook:new-window', filePath),
  saveWorkbook: (input) => ipcRenderer.invoke('workbook:save', input),
  exportWorkbook: (input) => ipcRenderer.invoke('workbook:export', input),
  renderPrintPreview: (input) => ipcRenderer.invoke('workbook:print-preview', input),
  printWorkbook: (input) => ipcRenderer.invoke('workbook:print', input),
  showItem: (filePath) => ipcRenderer.invoke('shell:show-item', filePath),
  openExternal: (target) => ipcRenderer.invoke('shell:open-external', target),
  getVersion: () => ipcRenderer.invoke('app:get-version'),
  protection: {
    hash: (password) => ipcRenderer.invoke('protection:hash', password),
    verify: (protection, password) => ipcRenderer.invoke('protection:verify', { protection, password }),
  },
  clipboard: {
    read: () => ipcRenderer.invoke('clipboard:read-rich'),
    write: (payload) => ipcRenderer.invoke('clipboard:write-rich', payload),
  },
  minimize: () => ipcRenderer.send('window:minimize'),
  toggleMaximize: () => ipcRenderer.send('window:toggle-maximize'),
  close: () => ipcRenderer.send('window:confirm-close'),
  onMaximized: (callback) => {
    const listener = (_event, maximized) => callback(maximized)
    ipcRenderer.on('window:maximized', listener)
    return () => ipcRenderer.removeListener('window:maximized', listener)
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
