const { contextBridge, ipcRenderer, webFrame, webUtils } = require('electron')

const pendingExternalOpenPaths = []
let externalOpenCallback = null

// Spell checking: the Windows Spell Checking API through the main process (offline).
// Chromium's own checker (webFrame) answers only when it has a local Hunspell dictionary,
// which a nonsense word proves; on Windows-supported languages it always says "correct".
function chromiumHasDictionary() {
  try { return webFrame.isWordMisspelled('qzxwvkjhbt') } catch { return false }
}

async function checkWords(words, language) {
  const list = Array.isArray(words) ? words.map((word) => String(word)) : []
  const result = await ipcRenderer.invoke('spell:check', { words: list, language })
  if (result && result.source) return result
  if (!chromiumHasDictionary()) return { source: null, language: null, misspelled: [] }
  return { source: 'chromium', language: null, misspelled: list.map((word) => { try { return webFrame.isWordMisspelled(word) } catch { return false } }) }
}

async function wordSuggestions(word, language) {
  const result = await ipcRenderer.invoke('spell:suggest', { word: String(word), language })
  if (result && result.source) return result.suggestions
  if (!chromiumHasDictionary()) return []
  try { return webFrame.getWordSuggestions(String(word)) } catch { return [] }
}

ipcRenderer.on('file:open-external', (_event, filePath) => {
  if (externalOpenCallback) externalOpenCallback(filePath)
  else pendingExternalOpenPaths.push(filePath)
})

contextBridge.exposeInMainWorld('simpleDocs', {
  // The curated fonts plus the installed families a window asks for (recent and document fonts).
  getDocumentFonts: (options) => ipcRenderer.invoke('document:fonts', { families: Array.isArray(options?.families) ? options.families.map(String) : [] }),
  // Every installed font family Simple can show and embed (enumerated locally, never downloaded).
  getFontFamilies: () => ipcRenderer.invoke('document:font-families'),
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
  // Formats Export As can offer now; .doc only when a local Office engine exists.
  getExportFormats: () => ipcRenderer.invoke('export:formats'),
  // Legacy/converted originals: where a .docx beside the original would go,
  // and the confirmed save that leaves the original untouched.
  planSaveBesideSource: () => ipcRenderer.invoke('file:sibling-docx-plan'),
  saveBesideSource: (input) => ipcRenderer.invoke('file:save-beside-source', input),
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
  spell: {
    isWordMisspelled: async (word, language) => (await checkWords([word], language)).misspelled[0] === true,
    getWordSuggestions: (word, language) => wordSuggestions(word, language),
    checkWords: (words, language) => checkWords(words, language),
    getLanguages: () => ipcRenderer.invoke('spell:languages'),
    getUserDictionary: () => ipcRenderer.invoke('spell:dictionary'),
    addWord: (word) => ipcRenderer.invoke('spell:add-word', word),
    removeWord: (word) => ipcRenderer.invoke('spell:remove-word', word),
    ignoreWord: (word) => ipcRenderer.invoke('spell:ignore-word', word),
    unignoreWord: (word) => ipcRenderer.invoke('spell:unignore-word', word),
    onDictionaryChanged: (callback) => {
      const listener = (_event, dictionary) => callback(dictionary)
      ipcRenderer.on('spell:dictionary-changed', listener)
      return () => ipcRenderer.removeListener('spell:dictionary-changed', listener)
    },
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
