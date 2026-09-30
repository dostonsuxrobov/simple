'use strict'

/**
 * Optional clipboard bridge for the renderer (exposed by the preload as
 * `window.simpleCalc.clipboard`). The web Clipboard API already reads/writes text/html in
 * this Electron build; the bridge adds Excel's "XML Spreadsheet" flavour, which carries
 * formulas and exact values and can only be read from the main process.
 *
 * Wiring (main.cjs):   registerClipboardHandlers(ipcMain, { assertTrustedSender })
 * Wiring (preload.cjs, inside the simpleCalc object):
 *   clipboard: {
 *     read: () => ipcRenderer.invoke('clipboard:read-rich'),
 *     write: (payload) => ipcRenderer.invoke('clipboard:write-rich', payload),
 *   },
 */
const { clipboard } = require('electron')

const MAX_CLIPBOARD_BYTES = 64 * 1024 * 1024
const EXCEL_XML_FORMAT = 'XML Spreadsheet'

function decodeBuffer(buffer) {
  if (!buffer || !buffer.length || buffer.length > MAX_CLIPBOARD_BYTES) return undefined
  let text
  if (buffer[0] === 0xff && buffer[1] === 0xfe) text = buffer.subarray(2).toString('utf16le')
  else if (buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) text = buffer.subarray(3).toString('utf8')
  else text = buffer.toString('utf8')
  // Windows clipboard blocks are often NUL-terminated.
  const end = text.indexOf('\u0000')
  if (end !== -1) text = text.slice(0, end)
  return text.trim() ? text : undefined
}

function readExcelXml() {
  if (process.platform !== 'win32') return undefined
  try {
    return decodeBuffer(clipboard.readBuffer(EXCEL_XML_FORMAT))
  } catch {
    return undefined
  }
}

function registerClipboardHandlers(ipcMain, options = {}) {
  const assertTrustedSender = typeof options.assertTrustedSender === 'function' ? options.assertTrustedSender : () => {}

  ipcMain.handle('clipboard:read-rich', (event) => {
    assertTrustedSender(event)
    const html = clipboard.readHTML()
    const payload = { text: clipboard.readText() }
    if (html) payload.html = html
    const excelXml = readExcelXml()
    if (excelXml) payload.excelXml = excelXml
    // A copied picture (a screenshot, an image from a browser) pastes as a floating picture.
    // Only when there is no text: spreadsheet copies also carry a bitmap of the range.
    if (!payload.text && !html) {
      const image = clipboard.readImage()
      if (image && !image.isEmpty()) payload.image = image.toDataURL()
    }
    return payload
  })

  ipcMain.handle('clipboard:write-rich', (event, payload) => {
    assertTrustedSender(event)
    if (!payload || typeof payload.text !== 'string') return false
    const html = typeof payload.html === 'string' ? payload.html : undefined
    if (payload.text.length + (html ? html.length : 0) > MAX_CLIPBOARD_BYTES) return false
    clipboard.write(html ? { text: payload.text, html } : { text: payload.text })
    return true
  })
}

module.exports = { registerClipboardHandlers }
