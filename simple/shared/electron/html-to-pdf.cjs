// Vendored from simple/shared/electron/html-to-pdf.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// Prints HTML to PDF with Chromium in a hidden window that cannot run scripts
// or reach the network. The window uses a private, in-memory session whose
// request filter lets through only data: URLs and local files inside the job
// folder (plus folders the caller names); http(s), ws, ftp, blob, remote file
// shares and every other scheme are cancelled. A job that does not finish
// within 30 seconds is stopped. Must run in the Electron main process; the
// module itself loads without Electron so its pure helpers can be tested.

const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { fileURLToPath, pathToFileURL } = require('node:url')

const PARTITION = 'simple-html-to-pdf'
const JOB_PREFIX = 'simple-html-pdf-'
const DEFAULT_TIMEOUT_MS = 30_000
const MAX_TIMEOUT_MS = 120_000
const MAX_HTML_BYTES = 64 * 1024 * 1024
const MAX_CONCURRENT_JOBS = 2
const SETTLE_MS = 200
const PAGE_SIZES = Object.freeze(['A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'Legal', 'Letter', 'Tabloid', 'Ledger'])
const CONTENT_SECURITY_POLICY = "default-src 'none'; img-src data: file:; style-src 'unsafe-inline' data: file:; font-src data: file:; media-src data: file:; base-uri 'none'; form-action 'none'; frame-src 'none'; object-src 'none'"

/** Allowed folders per print window, keyed by webContents id. */
const allowedRootsByContents = new Map()
let hardenedSession = null
let activeJobs = 0
const waitingJobs = []

/**
 * A coded print error. Codes: PRINT_TIMEOUT, PRINT_FAILED, PRINT_TOO_LARGE
 * and INVALID_OPTIONS. `message` is plain user-facing text; `technical` has the detail.
 */
class HtmlToPdfError extends Error {
  /**
   * @param {string} code
   * @param {string} message
   * @param {{technical?: string}} [details]
   */
  constructor(code, message, details = {}) {
    super(message)
    this.name = 'HtmlToPdfError'
    this.code = code
    if (details.technical) this.technical = details.technical
  }
}

function insideFolder(child, folder) {
  const relative = path.relative(folder, child)
  return relative === '' || (!!relative && !relative.startsWith('..') && !path.isAbsolute(relative))
}

function comparablePath(value) {
  const resolved = path.resolve(value)
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved
}

/**
 * Request policy of the print session: data: URLs, and file: URLs on this PC
 * (no host, so never a network share) that resolve inside one of `roots`.
 * @param {string} url
 * @param {string[]} roots absolute folders
 * @returns {boolean}
 */
function isAllowedRequest(url, roots) {
  let parsed
  try { parsed = new URL(url) } catch { return false }
  if (parsed.protocol === 'data:') return true
  if (parsed.protocol !== 'file:' || parsed.host !== '') return false
  let filePath
  try { filePath = fileURLToPath(parsed) } catch { return false }
  if (/^[\\/]{2}/.test(filePath)) return false
  const target = comparablePath(filePath)
  return (roots || []).some((root) => typeof root === 'string' && root && insideFolder(target, comparablePath(root)))
}

/**
 * Escapes text for HTML element content and attribute values.
 * @param {unknown} value
 * @returns {string}
 */
function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

/**
 * Turns caller HTML into the document that is printed: a fragment is wrapped
 * in a minimal page, and every document gets the restrictive
 * Content-Security-Policy and a UTF-8 charset as the first head elements.
 * @param {string} html a complete document or a body fragment
 * @param {{title?: string}} [options]
 * @returns {string}
 */
function prepareHtmlDocument(html, options = {}) {
  const head = `<meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${CONTENT_SECURITY_POLICY}">`
  const text = String(html ?? '')
  if (!/<html[\s>]/i.test(text) && !/^\s*<!doctype/i.test(text)) {
    return `<!doctype html><html><head>${head}<title>${escapeHtml(options.title || 'Document')}</title></head><body>${text}</body></html>`
  }
  if (/<head[\s>]/i.test(text)) return text.replace(/<head(\s[^>]*)?>/i, (match) => `${match}${head}`)
  return text.replace(/<html(\s[^>]*)?>/i, (match) => `${match}<head>${head}</head>`)
}

function finiteBetween(value, minimum, maximum, label) {
  const number = Number(value)
  if (!Number.isFinite(number) || number < minimum || number > maximum) {
    throw new HtmlToPdfError('INVALID_OPTIONS', "Simple couldn't make the PDF with these page settings.", { technical: `${label} must be between ${minimum} and ${maximum}.` })
  }
  return number
}

/**
 * Validates print options and maps them to Electron's printToPDF options.
 * @param {object} [options] see printHtmlToPdf
 * @returns {object} options for webContents.printToPDF
 */
function printOptionsFor(options = {}) {
  const result = {
    printBackground: options.printBackground !== false,
    landscape: Boolean(options.landscape),
    preferCSSPageSize: options.preferCSSPageSize !== undefined ? Boolean(options.preferCSSPageSize) : !options.pageSize,
    generateTaggedPDF: options.generateTaggedPDF !== false,
    generateDocumentOutline: options.generateDocumentOutline !== undefined ? Boolean(options.generateDocumentOutline) : options.generateTaggedPDF !== false,
  }
  if (options.pageSize !== undefined) {
    if (typeof options.pageSize === 'string') {
      if (!PAGE_SIZES.includes(options.pageSize)) throw new HtmlToPdfError('INVALID_OPTIONS', "Simple couldn't make the PDF with these page settings.", { technical: `Unknown page size ${options.pageSize}.` })
      result.pageSize = options.pageSize
    } else if (options.pageSize && typeof options.pageSize === 'object') {
      result.pageSize = {
        width: finiteBetween(options.pageSize.width, 0.5, 200, 'pageSize.width'),
        height: finiteBetween(options.pageSize.height, 0.5, 200, 'pageSize.height'),
      }
    } else throw new HtmlToPdfError('INVALID_OPTIONS', "Simple couldn't make the PDF with these page settings.", { technical: 'pageSize must be a name or {width, height} in inches.' })
  }
  if (options.margins !== undefined) {
    const margins = options.margins || {}
    result.margins = {}
    for (const side of ['top', 'right', 'bottom', 'left']) result.margins[side] = finiteBetween(margins[side] ?? 0, 0, 20, `margins.${side}`)
  }
  if (options.scale !== undefined) result.scale = finiteBetween(options.scale, 0.1, 2, 'scale')
  if (options.headerTemplate || options.footerTemplate) {
    result.displayHeaderFooter = true
    // An empty template must be explicit; otherwise Chromium prints its own date and title.
    result.headerTemplate = String(options.headerTemplate || '<span></span>')
    result.footerTemplate = String(options.footerTemplate || '<span></span>')
  }
  return result
}

function acquireSlot() {
  if (activeJobs < MAX_CONCURRENT_JOBS) {
    activeJobs += 1
    return Promise.resolve()
  }
  return new Promise((resolve) => waitingJobs.push(resolve))
}

function releaseSlot() {
  const next = waitingJobs.shift()
  if (next) next()
  else activeJobs -= 1
}

function printSession(electron) {
  if (hardenedSession) return hardenedSession
  const ses = electron.session.fromPartition(PARTITION, { cache: false })
  ses.webRequest.onBeforeRequest((details, callback) => {
    const roots = allowedRootsByContents.get(details.webContentsId)
    callback({ cancel: !(roots && isAllowedRequest(details.url, roots)) })
  })
  ses.setPermissionRequestHandler((_contents, _permission, callback) => callback(false))
  ses.setPermissionCheckHandler(() => false)
  ses.on('will-download', (event) => event.preventDefault())
  // The spell checker would fetch dictionaries; printing never needs it.
  if (typeof ses.setSpellCheckerEnabled === 'function') ses.setSpellCheckerEnabled(false)
  hardenedSession = ses
  return ses
}

function lockDown(contents, allowedUrl) {
  contents.setWindowOpenHandler(() => ({ action: 'deny' }))
  contents.on('will-navigate', (event, url) => { if (url !== allowedUrl) event.preventDefault() })
  contents.on('will-redirect', (event) => event.preventDefault())
  contents.on('will-attach-webview', (event) => event.preventDefault())
  contents.setAudioMuted(true)
}

async function removeJobDirectory(directory) {
  // Only the uniquely allocated job folder under the temp folder is removed.
  const resolved = path.resolve(directory)
  const relative = path.relative(path.resolve(os.tmpdir()), resolved)
  if (relative && !relative.startsWith('..') && !path.isAbsolute(relative) && path.basename(resolved).startsWith(JOB_PREFIX)) {
    await fs.rm(resolved, { recursive: true, force: true }).catch(() => {})
  }
}

/**
 * Prints HTML to a PDF.
 *
 * @param {string} html a complete HTML document or a body fragment; it should
 *   reference only data: URLs or files inside `fileRoots`
 * @param {object} [options]
 * @param {string} [options.title] title for a wrapped fragment
 * @param {string|{width: number, height: number}} [options.pageSize] a paper name
 *   (A0–A6, Letter, Legal, Tabloid, Ledger) or a size in inches; without it the
 *   document's CSS @page size is used
 * @param {boolean} [options.landscape=false]
 * @param {{top?: number, right?: number, bottom?: number, left?: number}} [options.margins] in inches
 * @param {number} [options.scale=1] 0.1–2
 * @param {boolean} [options.preferCSSPageSize] default: true when no pageSize is given
 * @param {boolean} [options.printBackground=true]
 * @param {string} [options.headerTemplate] Chromium header template (enables header and footer)
 * @param {string} [options.footerTemplate] Chromium footer template
 * @param {boolean} [options.generateTaggedPDF=true]
 * @param {boolean} [options.generateDocumentOutline] default: same as generateTaggedPDF
 * @param {string[]} [options.fileRoots] extra local folders the page may load files from
 * @param {number} [options.timeoutMs=30000] stop a job that takes longer (at most 120 s)
 * @returns {Promise<Buffer>} PDF bytes
 * @throws {HtmlToPdfError} PRINT_TIMEOUT, PRINT_FAILED, PRINT_TOO_LARGE or INVALID_OPTIONS
 */
async function printHtmlToPdf(html, options = {}) {
  if (typeof html !== 'string') throw new TypeError('printHtmlToPdf needs an HTML string.')
  const document = prepareHtmlDocument(html, options)
  if (Buffer.byteLength(document, 'utf8') > MAX_HTML_BYTES) {
    throw new HtmlToPdfError('PRINT_TOO_LARGE', 'This content is too large to turn into a PDF in one piece.', { technical: `${Buffer.byteLength(document, 'utf8')} bytes of HTML.` })
  }
  const printOptions = printOptionsFor(options)
  const timeoutMs = Math.round(finiteBetween(options.timeoutMs ?? DEFAULT_TIMEOUT_MS, 1, MAX_TIMEOUT_MS, 'timeoutMs'))
  const fileRoots = (Array.isArray(options.fileRoots) ? options.fileRoots : [])
    .filter((root) => typeof root === 'string' && path.isAbsolute(root) && !/^[\\/]{2}/.test(root))
  const electron = require('electron')
  await electron.app.whenReady()
  await acquireSlot()
  let directory = null
  let window = null
  let contentsId = null
  let timer = null
  try {
    directory = await fs.mkdtemp(path.join(os.tmpdir(), JOB_PREFIX))
    const htmlPath = path.join(directory, 'document.html')
    await fs.writeFile(htmlPath, document, 'utf8')
    window = new electron.BrowserWindow({
      show: false,
      width: 900,
      height: 1200,
      skipTaskbar: true,
      focusable: false,
      paintWhenInitiallyHidden: true,
      backgroundColor: '#ffffff',
      webPreferences: {
        session: printSession(electron),
        sandbox: true,
        contextIsolation: true,
        nodeIntegration: false,
        nodeIntegrationInSubFrames: false,
        javascript: false,
        webSecurity: true,
        allowRunningInsecureContent: false,
        webgl: false,
        plugins: false,
        experimentalFeatures: false,
        spellcheck: false,
        devTools: false,
        disableDialogs: true,
        navigateOnDragDrop: false,
        autoplayPolicy: 'document-user-activation-required',
        backgroundThrottling: false,
      },
    })
    contentsId = window.webContents.id
    allowedRootsByContents.set(contentsId, [directory, ...fileRoots])
    lockDown(window.webContents, pathToFileURL(htmlPath).href)
    const job = (async () => {
      await window.webContents.loadFile(htmlPath)
      await new Promise((resolve) => setTimeout(resolve, SETTLE_MS))
      return window.webContents.printToPDF(printOptions)
    })()
    job.catch(() => {})
    const timeout = new Promise((_resolve, reject) => {
      timer = setTimeout(() => reject(new HtmlToPdfError('PRINT_TIMEOUT', 'Making the PDF took too long, so Simple stopped. Your files are unchanged.', { technical: `Stopped after ${timeoutMs} ms.` })), timeoutMs)
    })
    const data = await Promise.race([job, timeout])
    const bytes = Buffer.from(data)
    if (bytes.length < 8 || bytes.subarray(0, 5).toString('latin1') !== '%PDF-') {
      throw new HtmlToPdfError('PRINT_FAILED', "Simple couldn't make the PDF. Your files are unchanged.", { technical: 'Chromium returned no PDF data.' })
    }
    return bytes
  } catch (error) {
    if (error instanceof HtmlToPdfError) throw error
    throw new HtmlToPdfError('PRINT_FAILED', "Simple couldn't make the PDF. Your files are unchanged.", { technical: error?.message || String(error) })
  } finally {
    clearTimeout(timer)
    if (contentsId !== null) allowedRootsByContents.delete(contentsId)
    if (window && !window.isDestroyed()) window.destroy()
    if (directory) await removeJobDirectory(directory)
    releaseSlot()
  }
}

module.exports = {
  CONTENT_SECURITY_POLICY,
  DEFAULT_TIMEOUT_MS,
  HtmlToPdfError,
  PAGE_SIZES,
  PARTITION,
  escapeHtml,
  isAllowedRequest,
  prepareHtmlDocument,
  printHtmlToPdf,
  printOptionsFor,
}
