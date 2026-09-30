'use strict'

const NO_PRINTERS_MESSAGE = 'No printers are installed. Add a printer in Windows Settings and try again.'
const NO_DEFAULT_PRINTER_MESSAGE = 'Windows could not find an available default printer. Choose a default printer in Windows Settings and try again.'

async function ensurePrinterInstalled(webContents) {
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
}

function directPrintOptions(options = {}) {
  const result = { ...options, silent: true }
  // Supplying any deviceName changes Electron from "Windows default" to an
  // explicitly named queue. Always omit it for this direct-default workflow.
  delete result.deviceName
  return result
}

function printFailureMessage(reason) {
  const detail = String(reason || '').trim()
  if (!detail || /invalid printer settings|default printer|no printer|printer.*not (?:found|available)/i.test(detail)) {
    return NO_DEFAULT_PRINTER_MESSAGE
  }
  if (/offline|unavailable|spool|driver/i.test(detail)) {
    return `The default printer is unavailable. Check that it is online and that the Windows Print Spooler is running, then try again. (${detail})`
  }
  return `Windows could not start the print job on the default printer: ${detail}`
}

module.exports = {
  NO_DEFAULT_PRINTER_MESSAGE,
  NO_PRINTERS_MESSAGE,
  directPrintOptions,
  ensurePrinterInstalled,
  printFailureMessage,
}
