const { contextBridge, ipcRenderer, webUtils } = require('electron')

const pendingExternalOpenPaths = []
let externalOpenCallback = null

ipcRenderer.on('file:open-external', (_event, filePath) => {
  if (externalOpenCallback) externalOpenCallback(filePath)
  else pendingExternalOpenPaths.push(filePath)
})

contextBridge.exposeInMainWorld('simpleImage', {
  openFile: () => ipcRenderer.invoke('file:open-dialog'),
  openPath: (filePath) => ipcRenderer.invoke('file:open-path', filePath),
  openBytes: (name, data) => ipcRenderer.invoke('file:open-bytes', { name, data }),
  openInNewWindow: (filePath) => ipcRenderer.invoke('file:open-in-new-window', filePath),
  newWindow: () => ipcRenderer.invoke('app:new-window'),
  saveImage: (input) => ipcRenderer.invoke('file:save', input),
  convertToPdf: (input) => ipcRenderer.invoke('image:convert-to-pdf', input),
  printImage: (input) => ipcRenderer.invoke('image:print', input),
  getVersion: () => ipcRenderer.invoke('app:get-version'),
  pathForFile: (file) => {
    try { return webUtils.getPathForFile(file) || null } catch { return null }
  },
  setTitle: (title) => ipcRenderer.send('window:set-title', title),
  minimize: () => ipcRenderer.send('window:minimize'),
  toggleMaximize: () => ipcRenderer.send('window:toggle-maximize'),
  confirmClose: () => ipcRenderer.send('window:confirm-close'),
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
