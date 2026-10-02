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
  getCapabilities: () => ipcRenderer.invoke('workbook:capabilities'),
  checkExport: (input) => ipcRenderer.invoke('workbook:export-check', input),
  renderPrintPreview: (input) => ipcRenderer.invoke('workbook:print-preview', input),
  printWorkbook: (input) => ipcRenderer.invoke('workbook:print', input),
  // Installed printers for the print dialog ([{ name, displayName, description }]).
  listPrinters: () => ipcRenderer.invoke('workbook:list-printers'),
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

// <simple-io-bridge v1>
// Vendored from simple/shared/preload/io-bridge.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
// Exposes window.simpleIO, the shared Save / Open / Export bridge (design §8.1).
// This block is copied verbatim between the simple-io-bridge markers of every
// workspace preload, which runs sandboxed and cannot require other files, so it
// is self-contained. The leading semicolon keeps it safe after a preceding
// expression without one; the function scope keeps its names private.
;(() => {
  'use strict'
  const { contextBridge, ipcRenderer, webUtils } = require('electron')
  const REQUEST_TYPES = ['close-query', 'save-now', 'discard', 'recovery-flush']
  const moduleArgument = (Array.isArray(process.argv) ? process.argv : []).find((value) => typeof value === 'string' && value.startsWith('--simple-io-module='))
  const environment = process.env && typeof process.env === 'object' ? process.env : {}
  const moduleName = moduleArgument ? moduleArgument.slice('--simple-io-module='.length) : (environment.SIMPLE_IO_MODULE || null)
  const requestHandlers = new Map()
  const capabilityListeners = new Set()

  // Main asks the page (close-query, save-now, discard, recovery-flush) and
  // waits for the answer. A page without a handler answers at once.
  ipcRenderer.on('io:request', (_event, message) => {
    if (!message || typeof message.id !== 'string') return
    const { id, type, payload } = message
    const handler = requestHandlers.get(type)
    if (!handler) {
      ipcRenderer.send('io:response', { id, ok: false, unhandled: true })
      return
    }
    Promise.resolve()
      .then(() => handler(payload))
      .then(
        (value) => ipcRenderer.send('io:response', { id, ok: true, value: value === undefined ? null : value }),
        (error) => ipcRenderer.send('io:response', { id, ok: false, error: String((error && error.message) || error || 'failed') }),
      )
  })

  ipcRenderer.on('io:capabilities-changed', () => {
    for (const listener of [...capabilityListeners]) {
      try { listener() } catch {}
    }
  })

  const invoke = (channel, ...args) => ipcRenderer.invoke(channel, ...args)

  contextBridge.exposeInMainWorld('simpleIO', {
    version: 1,
    module: moduleName,
    capabilities: () => invoke('io:capabilities'),
    onCapabilitiesChanged: (callback) => {
      if (typeof callback !== 'function') return () => {}
      const listener = () => callback()
      capabilityListeners.add(listener)
      return () => { capabilityListeners.delete(listener) }
    },
    pathForFile: (file) => {
      try { return webUtils.getPathForFile(file) || null } catch { return null }
    },
    chooseSavePath: (request) => invoke('io:choose-save-path', request),
    chooseOpenPaths: (request) => invoke('io:choose-open-paths', request),
    prompt: (key, vars) => invoke('io:prompt', key, vars || {}),
    recovery: {
      write: (snapshot) => invoke('io:recovery-write', snapshot),
      list: () => invoke('io:recovery-list'),
      read: (id) => invoke('io:recovery-read', id),
      adopt: (id) => invoke('io:recovery-adopt', id),
      discard: (id, upToRevision) => invoke('io:recovery-discard', id, upToRevision),
    },
    versions: {
      list: (filePath) => invoke('io:versions-list', filePath),
      open: (id) => invoke('io:versions-open', id),
    },
    clipboard: {
      read: () => invoke('io:clipboard-read'),
    },
    shell: {
      showItem: (filePath) => invoke('io:shell-show-item', filePath),
      // Opens the file inside Simple, never in another app.
      openPath: (filePath) => invoke('io:open-in-simple', filePath),
    },
    prefs: {
      get: (key) => invoke('io:prefs-get', key),
      set: (key, value) => invoke('io:prefs-set', key, value),
    },
    officeEngine: {
      status: () => invoke('io:engine-status'),
    },
    openInSimple: (filePath) => invoke('io:open-in-simple', filePath),
    onRequest: (type, handler) => {
      if (!REQUEST_TYPES.includes(type) || typeof handler !== 'function') return () => {}
      requestHandlers.set(type, handler)
      return () => {
        if (requestHandlers.get(type) === handler) requestHandlers.delete(type)
      }
    },
  })
})()
// </simple-io-bridge>
