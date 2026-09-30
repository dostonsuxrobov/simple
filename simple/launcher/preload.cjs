'use strict'

const { contextBridge, ipcRenderer, webUtils } = require('electron')

contextBridge.exposeInMainWorld('simpleLauncher', {
  info: () => ipcRenderer.invoke('launcher:info'),
  open: () => ipcRenderer.invoke('launcher:open'),
  combineAdd: (paths) => ipcRenderer.invoke('launcher:combine-add', paths),
  combineSave: (entries) => ipcRenderer.invoke('launcher:combine-save', entries),
  onCombineProgress: (callback) => {
    const listener = (_event, progress) => callback(progress)
    ipcRenderer.on('launcher:combine-progress', listener)
    return () => ipcRenderer.removeListener('launcher:combine-progress', listener)
  },
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
