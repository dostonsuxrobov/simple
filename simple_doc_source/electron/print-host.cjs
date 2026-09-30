const PDF_HOST_READY_TIMEOUT_MS = 20_000
const PRINT_JOB_TIMEOUT_MS = 120_000
const PRINT_DIRECTORY_PREFIX = 'simple-docs-print-'
const STALE_PRINT_DIRECTORY_AGE_MS = 24 * 60 * 60 * 1_000

function isStalePrintDirectory(name, stats, now = Date.now()) {
  if (!String(name).startsWith(PRINT_DIRECTORY_PREFIX) || !stats?.isDirectory?.()) return false
  const modifiedAt = Number(stats.mtimeMs)
  return Number.isFinite(modifiedAt) && modifiedAt <= now - STALE_PRINT_DIRECTORY_AGE_MS
}

function loadPdfForPrinting(printWindow, url, options = {}) {
  const timeoutMs = Number.isFinite(options.timeoutMs) && options.timeoutMs > 0
    ? options.timeoutMs
    : PDF_HOST_READY_TIMEOUT_MS
  if (!printWindow || printWindow.isDestroyed?.() || !printWindow.webContents || printWindow.webContents.isDestroyed?.()) {
    return Promise.reject(new Error('The print preview window is unavailable.'))
  }

  const contents = printWindow.webContents
  return new Promise((resolve, reject) => {
    let settled = false
    let navigationComplete = false
    let pdfParsed = false
    let timeout = null

    const cleanup = () => {
      if (timeout !== null) clearTimeout(timeout)
      contents.removeListener('page-title-updated', onPageTitleUpdated)
      contents.removeListener('render-process-gone', onRenderProcessGone)
      printWindow.removeListener('closed', onClosed)
    }
    const finish = (callback, value) => {
      if (settled) return
      settled = true
      cleanup()
      callback(value)
    }
    const maybeResolve = () => {
      if (navigationComplete && pdfParsed) finish(resolve)
    }
    const onPageTitleUpdated = (_event, _title, explicitSet) => {
      if (explicitSet === true) {
        pdfParsed = true
        maybeResolve()
      }
    }
    const onRenderProcessGone = (_event, details) => {
      finish(reject, new Error(`The PDF print renderer stopped before it was ready${details?.reason ? ` (${details.reason})` : ''}.`))
    }
    const onClosed = () => finish(reject, new Error('The PDF print window closed before it was ready.'))

    contents.on('page-title-updated', onPageTitleUpdated)
    contents.on('render-process-gone', onRenderProcessGone)
    printWindow.on('closed', onClosed)
    timeout = setTimeout(() => {
      finish(reject, new Error('The PDF print preview did not finish loading.'))
    }, timeoutMs)

    let navigation
    try {
      navigation = printWindow.loadURL(url)
    } catch (error) {
      finish(reject, error)
      return
    }
    Promise.resolve(navigation).then(() => {
      navigationComplete = true
      maybeResolve()
    }, (error) => finish(reject, error))
  })
}

function printWebContentsSilently(webContents, options = {}, timing = {}) {
  if (!webContents || webContents.isDestroyed?.()) {
    return Promise.reject(new Error('The print renderer is unavailable.'))
  }
  if (options.silent !== true) {
    return Promise.reject(new Error('Direct printing requires silent mode.'))
  }
  return new Promise((resolve) => {
    let settled = false
    let timeout
    const finish = (result) => {
      if (settled) return
      settled = true
      clearTimeout(timeout)
      webContents.removeListener?.('render-process-gone', onRendererGone)
      webContents.removeListener?.('destroyed', onDestroyed)
      resolve(result)
    }
    const onRendererGone = () => finish({ success: false, failureReason: 'The print renderer stopped. Check the Windows print queue before trying again.' })
    const onDestroyed = () => finish({ success: false, failureReason: 'The print window closed before Windows confirmed the job. Check the print queue before trying again.' })
    webContents.on?.('render-process-gone', onRendererGone)
    webContents.on?.('destroyed', onDestroyed)
    timeout = setTimeout(() => finish({
      success: false,
      failureReason: 'Windows did not confirm the print job in time. Check the print queue before trying again to avoid duplicate copies.',
    }), Number.isFinite(timing.timeoutMs) && timing.timeoutMs > 0 ? timing.timeoutMs : PRINT_JOB_TIMEOUT_MS)
    try {
      webContents.print(options, (success, failureReason) => {
        finish({ success: Boolean(success), failureReason: success ? '' : String(failureReason || 'The printer did not accept the job.') })
      })
    } catch (error) {
      finish({ success: false, failureReason: error instanceof Error ? error.message : String(error) })
    }
  })
}

module.exports = {
  isStalePrintDirectory,
  loadPdfForPrinting,
  printWebContentsSilently,
  PDF_HOST_READY_TIMEOUT_MS,
  PRINT_JOB_TIMEOUT_MS,
  PRINT_DIRECTORY_PREFIX,
  STALE_PRINT_DIRECTORY_AGE_MS,
}
