'use strict'

const { contextBridge, ipcRenderer, webUtils } = require('electron')

contextBridge.exposeInMainWorld('simpleLauncher', {
  info: () => ipcRenderer.invoke('launcher:info'),
  open: () => ipcRenderer.invoke('launcher:open'),
  launchMode: (mode) => ipcRenderer.invoke('launcher:launch-mode', mode),
  launchPaths: (paths) => ipcRenderer.invoke('launcher:launch-paths', paths),
  registerFileTypes: () => ipcRenderer.invoke('launcher:register-file-types'),
  unregisterFileTypes: () => ipcRenderer.invoke('launcher:unregister-file-types'),
  openDefaultApps: () => ipcRenderer.invoke('launcher:open-default-apps'),
  minimize: () => ipcRenderer.send('window:minimize'),
  close: () => ipcRenderer.send('window:close'),
  pathForFile: (file) => {
    try { return webUtils.getPathForFile(file) || null } catch { return null }
  },
})
