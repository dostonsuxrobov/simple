const { contextBridge, ipcRenderer } = require('electron')

const PDF_SIGNATURE_FIELD = Buffer.from('/Type /Sig')
const pendingExternalOpenPaths = []
let externalOpenCallback = null

// Capture shell-open events as soon as the preload starts. The main process can
// now dispatch immediately after navigation without racing React's effect.
ipcRenderer.on('file:open-external', (_event, filePath) => {
  if (externalOpenCallback) externalOpenCallback(filePath)
  else pendingExternalOpenPaths.push(filePath)
})

function openBytes(name, data) {
  // A dropped PDF is already in the renderer. Sending the complete file to the
  // main process only to send the same bytes straight back made large drops pay
  // for two extra IPC serializations. Conversions still go through the trusted
  // main-process handler.
  if (/\.pdf$/i.test(String(name || ''))) {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
    const bufferView = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    return Promise.resolve({
      data: bytes,
      name,
      path: null,
      sourcePath: null,
      converted: false,
      signatureDetected: bufferView.includes(PDF_SIGNATURE_FIELD),
    })
  }
  return ipcRenderer.invoke('file:open-bytes', { name, data })
}

contextBridge.exposeInMainWorld('simple', {
  openFile: () => ipcRenderer.invoke('file:open-dialog'),
  openInNewWindow: (filePath) => ipcRenderer.invoke('file:open-in-new-window', filePath),
  openPath: (filePath) => ipcRenderer.invoke('file:open-path', filePath),
  openBytes,
  mutatePdf: (data, operation) => ipcRenderer.invoke('pdf:mutate', data, operation),
  flattenOverlays: (data, overlays, formValues, documentEdits) => ipcRenderer.invoke('pdf:flatten-overlays', data, overlays, formValues, documentEdits),
  insertFiles: (data, insertIndex) => ipcRenderer.invoke('pdf:insert-files', data, insertIndex),
  insertDroppedFiles: (data, insertIndex, files) => ipcRenderer.invoke('pdf:insert-dropped-files', data, insertIndex, files),
  pickImage: () => ipcRenderer.invoke('file:pick-image'),
  exportPages: (data, indices, suggestedName) => ipcRenderer.invoke('pdf:export-pages', data, indices, suggestedName),
  startPageDrag: (data, indices, suggestedName) => ipcRenderer.invoke('pdf:start-page-drag', data, indices, suggestedName),
  printPdf: (data, documentName) => ipcRenderer.invoke('pdf:print', data, documentName),
  listPrinters: () => ipcRenderer.invoke('print:list-printers'),
  printPdfDirect: (data, documentName, options) => ipcRenderer.invoke('pdf:print-direct', data, documentName, options),
  savePdf: (input) => ipcRenderer.invoke('pdf:save', input),
  showItem: (filePath) => ipcRenderer.invoke('shell:show-item', filePath),
  openExternal: (url) => ipcRenderer.invoke('shell:open-external', url),
  getVersion: () => ipcRenderer.invoke('app:get-version'),
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
    return () => {
      if (externalOpenCallback === callback) externalOpenCallback = null
    }
  },
  onCloseRequested: (callback) => {
    const listener = () => callback()
    ipcRenderer.on('window:close-requested', listener)
    return () => ipcRenderer.removeListener('window:close-requested', listener)
  },
})
