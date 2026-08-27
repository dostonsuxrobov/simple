'use strict'

const { contextBridge, ipcRenderer, webUtils } = require('electron')

const queuedOpenPaths = []
let openPathListener = null

ipcRenderer.on('file:open-external', (_event, filePath) => {
  if (openPathListener) openPathListener(filePath)
  else queuedOpenPaths.push(filePath)
})

contextBridge.exposeInMainWorld('simpleVideo', {
  openFile: () => ipcRenderer.invoke('file:open-dialog'),
  openPath: (filePath) => ipcRenderer.invoke('file:open-path', filePath),
  openInNewWindow: (filePath) => ipcRenderer.invoke('file:open-in-new-window', filePath),
  newWindow: () => ipcRenderer.invoke('app:new-window'),
  listRecents: () => ipcRenderer.invoke('recent:list'),
  removeRecent: (filePath) => ipcRenderer.invoke('recent:remove', filePath),
  clearRecents: () => ipcRenderer.invoke('recent:clear'),
  revealFile: (filePath) => ipcRenderer.invoke('shell:show-item', filePath),
  pathForFile: (file) => {
    try {
      return webUtils.getPathForFile(file) || null
    } catch {
      return null
    }
  },
  setTitle: (title) => ipcRenderer.send('window:set-title', title),
  minimize: () => ipcRenderer.send('window:minimize'),
  toggleMaximize: () => ipcRenderer.send('window:toggle-maximize'),
  toggleFullscreen: () => ipcRenderer.send('window:toggle-fullscreen'),
  exitFullscreen: () => ipcRenderer.send('window:exit-fullscreen'),
  close: () => ipcRenderer.send('window:close'),
  onOpenExternal: (callback) => {
    openPathListener = callback
    while (queuedOpenPaths.length) callback(queuedOpenPaths.shift())
    return () => {
      if (openPathListener === callback) openPathListener = null
    }
  },
  onMaximized: (callback) => {
    const listener = (_event, value) => callback(Boolean(value))
    ipcRenderer.on('window:maximized', listener)
    return () => ipcRenderer.removeListener('window:maximized', listener)
  },
  onFullscreen: (callback) => {
    const listener = (_event, value) => callback(Boolean(value))
    ipcRenderer.on('window:fullscreen', listener)
    return () => ipcRenderer.removeListener('window:fullscreen', listener)
  },
})
