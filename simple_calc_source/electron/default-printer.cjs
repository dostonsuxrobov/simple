'use strict'

const NO_PRINTERS_MESSAGE = 'No printers are installed. Add a printer in Windows Settings and try again.'
const NO_DEFAULT_PRINTER_MESSAGE = 'Windows could not find an available default printer. Choose a default printer in Windows Settings and try again.'
const MAX_PRINT_COPIES = 999
const MAX_LISTED_PRINTERS = 200

/** The explicitly chosen printer queue of a print job, or '' for the Windows default. */
function jobDeviceName(job) {
  const name = job && typeof job === 'object' && typeof job.deviceName === 'string' ? job.deviceName.trim() : ''
  return name && name.length <= 256 && !/[\u0000-\u001f\u007f]/.test(name) ? name : ''
}

/**
 * The installed printers for the print dialog: name (the queue the OS knows), displayName
 * and description. PrinterInfo no longer marks the Windows default, so the dialog offers
 * "Default printer" itself and sends no deviceName for it.
 */
async function listPrinters(webContents) {
  if (!webContents || typeof webContents.getPrintersAsync !== 'function') return []
  const printers = await webContents.getPrintersAsync()
  const seen = new Set()
  const result = []
  for (const printer of Array.isArray(printers) ? printers : []) {
    const name = printer && typeof printer.name === 'string' ? printer.name.trim() : ''
    if (!name || name.length > 256 || seen.has(name)) continue
    seen.add(name)
    const displayName = typeof printer.displayName === 'string' && printer.displayName.trim() ? printer.displayName.trim().slice(0, 256) : name
    const description = typeof printer.description === 'string' ? printer.description.trim().slice(0, 256) : ''
    result.push({ name, displayName, description })
    if (result.length >= MAX_LISTED_PRINTERS) break
  }
  return result.sort((first, second) => first.displayName.localeCompare(second.displayName))
}

async function ensurePrinterInstalled(webContents, job = null) {
  // PrinterInfo no longer identifies the Windows default. Enumeration is only a
  // useful early check for the unambiguous zero-printer case; Chromium chooses
  // the system default when a silent print omits deviceName.
  if (!webContents || typeof webContents.getPrintersAsync !== 'function') return
  let printers
  try {
    printers = await webContents.getPrintersAsync()
  } catch {
    // Let the actual print request report spooler/driver failures. Enumeration
    // can fail independently even when Chromium can still queue the document.
    return
  }
  if (Array.isArray(printers) && printers.length === 0) throw new Error(NO_PRINTERS_MESSAGE)
  // A chosen printer that was removed since the dialog listed it must not silently
  // become another queue.
  const deviceName = jobDeviceName(job)
  if (deviceName && Array.isArray(printers) && !printers.some((printer) => printer && printer.name === deviceName)) {
    throw new Error(`The printer "${deviceName}" is no longer available. Choose another printer and try again.`)
  }
}

/**
 * Options for a silent webContents.print. Without a job (or without job.deviceName) the
 * document goes to the Windows default printer; a job adds the chosen printer, copies and
 * collation. Layout settings (scale, margins, page range) are already in the document.
 */
function directPrintOptions(options = {}, job = null) {
  const result = { ...options, silent: true }
  // Supplying any deviceName changes Electron from "Windows default" to an
  // explicitly named queue. Only a validated job choice may set it.
  delete result.deviceName
  if (job && typeof job === 'object') {
    const deviceName = jobDeviceName(job)
    if (deviceName) result.deviceName = deviceName
    const copies = Math.trunc(Number(job.copies))
    if (Number.isFinite(copies) && copies >= 1) result.copies = Math.min(MAX_PRINT_COPIES, copies)
    if (typeof job.collate === 'boolean') result.collate = job.collate
  }
  return result
}

/** webContents.print options for a document from createSpreadsheetPrintDocument. */
function printOptionsForDocument(printDocument) {
  const options = (printDocument && printDocument.options) || {}
  return directPrintOptions({
    printBackground: true,
    color: true,
    landscape: options.orientation === 'landscape',
    margins: { marginType: 'none' },
    pageSize: options.paperSize === 'a4' ? 'A4' : options.paperSize === 'legal' ? 'Legal' : 'Letter',
    scaleFactor: 100,
    pagesPerSheet: 1,
    collate: true,
  }, printDocument && printDocument.printJob)
}

function printFailureMessage(reason, job = null) {
  const detail = String(reason || '').trim()
  const deviceName = jobDeviceName(job)
  if (deviceName) {
    if (!detail || /invalid printer settings|no printer|printer.*not (?:found|available)/i.test(detail)) {
      return `Windows could not use the printer "${deviceName}". Check that it is installed and online, or choose another printer.`
    }
    if (/offline|unavailable|spool|driver/i.test(detail)) {
      return `The printer "${deviceName}" is unavailable. Check that it is online and that the Windows Print Spooler is running, then try again. (${detail})`
    }
    return `Windows could not start the print job on "${deviceName}": ${detail}`
  }
  if (!detail || /invalid printer settings|default printer|no printer|printer.*not (?:found|available)/i.test(detail)) {
    return NO_DEFAULT_PRINTER_MESSAGE
  }
  if (/offline|unavailable|spool|driver/i.test(detail)) {
    return `The default printer is unavailable. Check that it is online and that the Windows Print Spooler is running, then try again. (${detail})`
  }
  return `Windows could not start the print job on the default printer: ${detail}`
}

/**
 * Printer IPC for the main process: `workbook:list-printers` answers the print dialog's
 * printer list (preload: simpleCalc.listPrinters). Call once from registerIpc with the
 * same sender check the other handlers use.
 */
function registerPrinterHandlers(ipcMain, { assertTrustedSender } = {}) {
  ipcMain.handle('workbook:list-printers', async (event) => {
    if (typeof assertTrustedSender === 'function') assertTrustedSender(event)
    return listPrinters(event && event.sender)
  })
}

module.exports = {
  MAX_PRINT_COPIES,
  NO_DEFAULT_PRINTER_MESSAGE,
  NO_PRINTERS_MESSAGE,
  directPrintOptions,
  ensurePrinterInstalled,
  listPrinters,
  printFailureMessage,
  printOptionsForDocument,
  registerPrinterHandlers,
}
