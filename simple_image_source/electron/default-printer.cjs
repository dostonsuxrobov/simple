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

function submitPrintJob(webContents, options, { timeoutMs = 120_000 } = {}) {
  return new Promise((resolve, reject) => {
    let settled = false
    let timer
    const finish = (error) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      webContents?.removeListener?.('destroyed', onClosed)
      webContents?.removeListener?.('render-process-gone', onClosed)
      if (error) reject(error)
      else resolve(true)
    }
    const onClosed = () => finish(new Error('The print renderer stopped before Windows confirmed the job. Check the print queue before trying again.'))
    if (!webContents || webContents.isDestroyed?.()) { onClosed(); return }
    webContents.once?.('destroyed', onClosed)
    webContents.once?.('render-process-gone', onClosed)
    timer = setTimeout(() => finish(new Error('Windows did not confirm the print job within two minutes. Check the print queue before trying again to avoid a duplicate copy.')), Math.max(1, timeoutMs))
    try {
      webContents.print(options, (success, reason) => finish(success ? null : new Error(printFailureMessage(reason))))
    } catch (error) { finish(error) }
  })
}

module.exports = {
  NO_DEFAULT_PRINTER_MESSAGE,
  NO_PRINTERS_MESSAGE,
  directPrintOptions,
  ensurePrinterInstalled,
  printFailureMessage,
  submitPrintJob,
}
