const { app, BrowserWindow, dialog, ipcMain, nativeImage, shell } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const crypto = require('node:crypto')
const { pathToFileURL } = require('node:url')
const { replacePdfOutlines } = require('./pdf-outlines.cjs')
const { removePageImageDraws } = require('./pdf-content-edits.cjs')
const { removeNativeText } = require('./text-removal.cjs')
const {
  addGlyphlessFont, addOcrTextLayer, glyphlessFontKey, removeInvisibleText, validateOcrLayerOperation, writeInvisibleRun,
} = require('./ocr-text-layer.cjs')
const { registerSharedIo, bridgeArguments } = require('./simple-io/io-ipc.cjs')
const { sweep } = require('./simple-io/io-core.cjs')
const { safeWriteFile } = require('./simple-io/safe-write.cjs')
const guard = require('./simple-io/document-guard.cjs')
const stores = require('./simple-io/stores.cjs')
const officeEngine = require('./simple-io/office-engine.cjs')

const IO_MODULE = 'pdf'
/** Every file type the PDF workspace opens; mirrors IMPORT_FORMATS in office-import.cjs (checked by tests/supported-extensions.test.cjs and simple/scripts/verify.cjs). */
const SUPPORTED_EXTENSIONS = new Set([
  '.pdf', '.docx', '.docm', '.dotx', '.dotm', '.doc', '.odt', '.rtf',
  '.xlsx', '.xlsm', '.xltx', '.xltm', '.xlsb', '.xls', '.ods', '.csv', '.tsv',
  '.md', '.markdown', '.mdown', '.mkd', '.html', '.htm', '.xhtml', '.txt', '.text', '.log',
  '.png', '.jpg', '.jpeg', '.jpe', '.jfif', '.gif', '.webp', '.bmp', '.dib', '.tif', '.tiff', '.svg', '.avif', '.ico',
  '.pptx', '.ppt', '.odp',
])
const {
  IMAGE_EXPORT_FORMATS,
  TEXT_EXPORT_FORMATS,
  buildTextExportFiles,
  imageExportFileName,
  safeExportBaseName,
} = require('./pdf-export.cjs')
const { loadPdfViewerForPrint, nativePrintOptions, preparePrintPdf, printWebContentsSilently, resolvePrinter } = require('./pdf-print.cjs')
const { jpegOrientation, orientationPlacement } = require('./image-to-pdf.cjs')
const {
  convertToPdf,
  importDialogFilters,
  isImportableName,
  officeEngineAvailable,
  pickedImageData,
  sniffType,
  unsupportedFormatError,
} = require('./office-import.cjs')
const { embedPdfFont } = require('./font-embedding.cjs')
const { compactUnreachable, copyPagesRemapped, detachRemovedPages, pageLeaves } = require('./pdf-compact.cjs')
const { detectSignatures, documentSignatureStatus, fileSignatureStatus } = require('./pdf-signatures.cjs')
const { encryptedDocumentError, hasEncryptionDictionary, isEncryptedPdfError, readableCopy } = require('./pdf-unlock.cjs')
const { codedError, problem, unsavedChangesError } = require('./pdf-problems.cjs')
const { FontLibrary, loadFaceFile } = require('./font-fallback.cjs')
const { applyFormValues, verifyFormValues } = require('./form-fill.cjs')
const { createTextBackgroundCache } = require('./text-background-cache.cjs')

// Text that cannot fit its box is reduced down to this size before it is
// allowed to run past the box (it is never dropped).
const MIN_FITTED_TEXT_SIZE = 6
// Share of a one-line text box above the baseline (pdf.js's default ascent).
const NATIVE_RUN_ASCENT = 0.8

let isQuitting = false
let fontkitModule
let pdfLibModule
let dragExportDirectory = null
const imageExportSessions = new Map()

// Conversion/editing dependencies are intentionally loaded only when a user
// invokes those features. Fontkit is expensive to initialise and is not
// needed at all for the common PDF viewing path; office-import.cjs loads its
// converters (mammoth, SheetJS, markdown-it) the same way.
function getFontkit() {
  fontkitModule ||= require('@pdf-lib/fontkit')
  return fontkitModule
}

function getPdfLib() {
  pdfLibModule ||= require('pdf-lib')
  return pdfLibModule
}

function toBytes(value) {
  if (Buffer.isBuffer(value)) return value
  // Buffer views over typed-array storage avoid another full-document copy on
  // every mutate, save, print, insert, and export IPC request.
  if (ArrayBuffer.isView(value)) {
    return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  }
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data)
  return Buffer.from(value)
}

function serializableBytes(value) {
  const bytes = value instanceof Uint8Array ? value : new Uint8Array(value)
  return bytes
}

function safeBaseName(filePath) {
  return path.basename(filePath, path.extname(filePath))
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim() || 'Untitled'
}

async function getPdfFont(pdfDoc, requestedFamily = 'Segoe UI', cache = new Map(), style = {}) {
  const { StandardFonts } = getPdfLib()
  const windowsDir = process.env.WINDIR || 'C:\\Windows'
  const family = String(requestedFamily || 'Segoe UI').toLowerCase()
  const bold = Number(style.fontWeight) >= 600 || /bold|black|semibold|demi/.test(family)
  const italic = style.fontStyle === 'italic' || /italic|oblique/.test(family)
  const definition = windowsFontDefinition(family, bold, italic, StandardFonts)
  const embeddedKey = style.fontKey ? `embedded:${style.fontKey}:${style.fontData?.byteLength || 0}` : ''
  if (embeddedKey && style.fontData?.byteLength) {
    try {
      const fontBytes = toBytes(style.fontData)
      const decodedFont = getFontkit().create(fontBytes)
      const missing = Array.from(String(style.text || ''))
        .filter((character) => !/\s/.test(character) && !decodedFont.hasGlyphForCodePoint(character.codePointAt(0)))
      // Keep the document's own face when the matched Windows family cannot
      // draw the missing characters either (e.g. CJK or emoji typed into Latin
      // text); they come from fallback fonts per character instead.
      if (missing.length && await windowsFamilyCovers(definition, windowsDir, missing)) {
        throw new Error('Decoded PDF subset does not contain every replacement glyph.')
      }
      if (cache.has(embeddedKey)) return cache.get(embeddedKey)
      pdfDoc.registerFontkit(getFontkit())
      const embedded = await embedPdfFont(pdfDoc, fontBytes, style.text || '')
      // Fontkit defers subset serialization until PDF.save(). Exercise it here
      // while we can still fall back from a malformed browser-decoded font.
      cache.set(embeddedKey, embedded)
      return embedded
    } catch {
      // Some PDFs expose a browser-only decoded face. Use a matched Windows
      // family below when fontkit cannot re-embed those bytes.
    }
  }

  if (cache.has(definition.key)) return cache.get(definition.key)
  const candidates = definition.files.map((fileName) => path.join(windowsDir, 'Fonts', fileName))

  for (const candidate of candidates) {
    try {
      const fontBytes = await fs.readFile(candidate)
      pdfDoc.registerFontkit(getFontkit())
      const font = await embedPdfFont(pdfDoc, fontBytes, style.text || '')
      cache.set(definition.key, font)
      return font
    } catch {
      // Continue to the built-in fallback.
    }
  }

  const font = await pdfDoc.embedFont(definition.fallback)
  cache.set(definition.key, font)
  return font
}

async function windowsFamilyCovers(definition, windowsDir, characters) {
  for (const fileName of definition.files) {
    const face = await loadFaceFile(path.join(windowsDir, 'Fonts', fileName))
    if (face) return characters.some((character) => face.font.hasGlyphForCodePoint(character.codePointAt(0)))
  }
  return true
}

function windowsFontDefinition(family, bold, italic, StandardFonts) {
  const variant = (regular, boldFile, italicFile, boldItalicFile) => (
    bold && italic ? boldItalicFile : bold ? boldFile : italic ? italicFile : regular
  )
  const standardVariant = (regular, boldFont, italicFont, boldItalicFont) => (
    bold && italic ? boldItalicFont : bold ? boldFont : italic ? italicFont : regular
  )
  const timesFallback = standardVariant(StandardFonts.TimesRoman, StandardFonts.TimesRomanBold, StandardFonts.TimesRomanItalic, StandardFonts.TimesRomanBoldItalic)
  const sansFallback = standardVariant(StandardFonts.Helvetica, StandardFonts.HelveticaBold, StandardFonts.HelveticaOblique, StandardFonts.HelveticaBoldOblique)
  const monoFallback = standardVariant(StandardFonts.Courier, StandardFonts.CourierBold, StandardFonts.CourierOblique, StandardFonts.CourierBoldOblique)
  let definition
  if (family.includes('calibri')) {
    definition = { key: 'calibri', files: [variant('calibri.ttf', 'calibrib.ttf', 'calibrii.ttf', 'calibriz.ttf')], fallback: sansFallback }
  } else if (family.includes('cambria')) {
    definition = { key: 'cambria', files: [variant('cambria.ttc', 'cambriab.ttf', 'cambriai.ttf', 'cambriaz.ttf')], fallback: timesFallback }
  } else if (family.includes('garamond') || family.includes('crimson')) {
    definition = { key: 'garamond', files: [variant('GARA.TTF', 'GARABD.TTF', 'GARAIT.TTF', 'GARAIT.TTF')], fallback: timesFallback }
  } else if (family.includes('palatino')) {
    definition = { key: 'palatino linotype', files: [variant('pala.ttf', 'palab.ttf', 'palai.ttf', 'palabi.ttf')], fallback: timesFallback }
  } else if (family.includes('georgia') || family.includes('playfair')) {
    definition = { key: 'georgia', files: [variant('georgia.ttf', 'georgiab.ttf', 'georgiai.ttf', 'georgiaz.ttf')], fallback: timesFallback }
  } else if (family.includes('times') || family.includes('minion') || (family.includes('serif') && !family.includes('sans-serif'))) {
    definition = { key: 'times new roman', files: [variant('times.ttf', 'timesbd.ttf', 'timesi.ttf', 'timesbi.ttf')], fallback: timesFallback }
  } else if (family.includes('courier') || family.includes('mono')) {
    definition = { key: 'courier new', files: [variant('cour.ttf', 'courbd.ttf', 'couri.ttf', 'courbi.ttf')], fallback: monoFallback }
  } else if (family.includes('verdana')) {
    definition = { key: 'verdana', files: [variant('verdana.ttf', 'verdanab.ttf', 'verdanai.ttf', 'verdanaz.ttf')], fallback: sansFallback }
  } else if (family.includes('tahoma')) {
    definition = { key: 'tahoma', files: [variant('tahoma.ttf', 'tahomabd.ttf', 'tahoma.ttf', 'tahomabd.ttf')], fallback: sansFallback }
  } else if (family.includes('trebuchet')) {
    definition = { key: 'trebuchet ms', files: [variant('trebuc.ttf', 'trebucbd.ttf', 'trebucit.ttf', 'trebucbi.ttf')], fallback: sansFallback }
  } else if (family.includes('arial') || family.includes('helvetica') || family.includes('sans-serif')) {
    definition = { key: 'arial', files: [variant('arial.ttf', 'arialbd.ttf', 'ariali.ttf', 'arialbi.ttf')], fallback: sansFallback }
  } else {
    definition = { key: 'segoe ui', files: [variant('segoeui.ttf', 'segoeuib.ttf', 'segoeuii.ttf', 'segoeuiz.ttf'), 'arial.ttf'], fallback: sansFallback }
  }

  definition.key = `${definition.key}:${bold ? 'bold' : 'regular'}:${italic ? 'italic' : 'normal'}`
  return definition
}

/**
 * Any file Simple can open as PDF bytes (see office-import.cjs): PDFs pass
 * through, everything else is converted without an office suite.
 * @returns {Promise<{data: Buffer|Uint8Array, warnings: string[], kind: string, converted: boolean}>}
 */
function importAsPdf(bytes, name, sourcePath = null) {
  return convertToPdf(bytes, sourcePath ? { name, sourcePath } : { name })
}

/** True for a path Simple opens: a known extension, or a PDF without one. */
async function openablePath(filePath) {
  if (isImportableName(filePath)) return true
  let handle
  try {
    handle = await fs.open(filePath, 'r')
    const { buffer, bytesRead } = await handle.read(Buffer.alloc(1024), 0, 1024, 0)
    return sniffType(buffer.subarray(0, bytesRead)) === 'pdf'
  } catch {
    return false
  } finally {
    await handle?.close().catch(() => {})
  }
}

async function filePayload(filePath) {
  if (!(await openablePath(filePath))) throw unsupportedFormatError(filePath)
  const original = await fs.readFile(filePath)
  // A PDF keeps the fs Buffer: constructing a Uint8Array from it used to
  // duplicate the complete file before Electron's unavoidable IPC copy.
  const result = await importAsPdf(original, path.basename(filePath), filePath)
  return {
    data: serializableBytes(result.data),
    name: result.converted ? `${safeBaseName(filePath)}.pdf` : path.basename(filePath),
    path: result.converted ? null : filePath,
    sourcePath: filePath,
    converted: result.converted,
    ...(result.warnings.length ? { warnings: result.warnings } : {}),
    ...documentProtection(result.converted ? '' : '.pdf', original),
  }
}

/**
 * `signatureDetected` makes the renderer save a copy instead of overwriting a
 * signed original. `encrypted` is decided from the whole file: linearized
 * files keep /Encrypt near the start, which a scan of the tail misses.
 */
function documentProtection(ext, bytes) {
  if (ext !== '.pdf') return { signatureDetected: false, encrypted: false }
  return {
    signatureDetected: detectSignatures(bytes).signed,
    encrypted: hasEncryptionDictionary(bytes),
  }
}

async function loadPdf(data, name = '') {
  const { PDFDocument } = getPdfLib()
  const bytes = toBytes(data)
  try {
    return await PDFDocument.load(bytes, {
      updateMetadata: false,
      throwOnInvalidObject: false,
    })
  } catch (error) {
    // pdf-lib's "Input document to `PDFDocument.load` is encrypted…" becomes
    // PASSWORD_REQUIRED or OWNER_LOCKED with a message a user can act on.
    if (isEncryptedPdfError(error)) throw await encryptedDocumentError(bytes, name)
    throw error
  }
}

/** Pages copied from another file: owner-locked sources are decrypted to a copy. */
async function loadSourcePdf(bytes, name) {
  return loadPdf(await readableCopy(bytes, name), name)
}

function pageAt(pdfDoc, index) {
  const pages = pdfDoc.getPages()
  return Number.isInteger(index) && index >= 0 && index < pages.length ? pages[index] : undefined
}

async function reorderDocument(pdfDoc, order) {
  const pageCount = pdfDoc.getPageCount()
  if (!Array.isArray(order) || order.length !== pageCount) throw new Error('Invalid page order.')
  const unique = new Set(order)
  if (unique.size !== pageCount || Math.min(...order) !== 0 || Math.max(...order) !== pageCount - 1) {
    throw new Error('Invalid page order.')
  }
  // Move the existing page objects in the original catalog. Rebuilding a PDF
  // from copied pages can discard outlines, attachments, metadata, page labels,
  // named destinations, viewer preferences, and form relationships.
  const currentOrder = Array.from({ length: pageCount }, (_, index) => index)
  const pages = pdfDoc.getPages().slice()
  for (let targetIndex = 0; targetIndex < order.length; targetIndex += 1) {
    const sourceIndex = currentOrder.indexOf(order[targetIndex])
    if (sourceIndex === targetIndex) continue
    const [page] = pages.splice(sourceIndex, 1)
    const [identity] = currentOrder.splice(sourceIndex, 1)
    pdfDoc.removePage(sourceIndex)
    pdfDoc.insertPage(targetIndex, page)
    pages.splice(targetIndex, 0, page)
    currentOrder.splice(targetIndex, 0, identity)
  }
  pdfDoc.setProducer('simple')
  return pdfDoc
}

async function applyMutation(data, operation) {
  let pdfDoc = await loadPdf(data)
  switch (operation.type) {
    case 'delete': {
      const indices = [...new Set(operation.indices || [])]
        .filter((index) => Number.isInteger(index) && index >= 0 && index < pdfDoc.getPageCount())
        .sort((a, b) => b - a)
      if (!indices.length) return pdfDoc.save()
      if (indices.length >= pdfDoc.getPageCount()) throw new Error('A PDF must keep at least one page.')
      const leaves = pageLeaves(pdfDoc)
      const removedRefs = indices.map((index) => leaves[index]?.ref).filter(Boolean)
      indices.forEach((index) => pdfDoc.removePage(index))
      // pdf-lib's removePage() leaves its page list cache stale.
      pdfDoc.pageCache?.invalidate?.()
      // Otherwise the page, its text and images stay recoverable in the saved
      // file, and bookmarks, links and fields keep pointing at it.
      detachRemovedPages(pdfDoc, removedRefs)
      break
    }
    case 'rotate': {
      const { degrees } = getPdfLib()
      const amount = Number(operation.degrees) || 90
      for (const index of operation.indices || []) {
        const page = pdfDoc.getPage(index)
        const current = page.getRotation().angle || 0
        page.setRotation(degrees(((current + amount) % 360 + 360) % 360))
      }
      break
    }
    case 'reorder': {
      pdfDoc = await reorderDocument(pdfDoc, operation.order)
      break
    }
    case 'duplicate': {
      const indices = [...new Set(operation.indices || [])]
        .filter((index) => Number.isInteger(index) && index >= 0 && index < pdfDoc.getPageCount())
        .sort((a, b) => a - b)
      let inserted = 0
      for (const originalIndex of indices) {
        const sourceIndex = originalIndex + inserted
        const [copy] = await pdfDoc.copyPages(pdfDoc, [sourceIndex])
        pdfDoc.insertPage(sourceIndex + 1, copy)
        inserted += 1
      }
      break
    }
    case 'crop': {
      const page = pdfDoc.getPage(operation.pageIndex)
      const { x, y, width, height } = operation.rect
      if (![x, y, width, height].every(Number.isFinite) || width < 8 || height < 8) {
        throw new Error('The crop area is too small.')
      }
      page.setCropBox(x, y, width, height)
      break
    }
    case 'blank': {
      const width = Number(operation.width) || 595.28
      const height = Number(operation.height) || 841.89
      pdfDoc.insertPage(Math.max(0, Math.min(operation.index, pdfDoc.getPageCount())), [width, height])
      break
    }
    case 'ocr-text-layer': {
      // Recognised text arrives from the renderer and is validated here. With
      // replaceExisting, invisible text recognised earlier (by Simple or by
      // another tool) is removed first; pages listed without words only lose it.
      const op = validateOcrLayerOperation(operation, pdfDoc.getPageCount())
      let replaced = false
      if (op.replaceExisting) {
        const bytes = toBytes(data)
        const cleaned = await removeInvisibleText(bytes, op.pages.map((page) => page.pageIndex))
        if (cleaned !== bytes) {
          pdfDoc = await loadPdf(cleaned)
          replaced = true
        }
      }
      const pages = op.pages.filter((page) => page.lines.length)
      const fontRef = pages.length ? addGlyphlessFont(pdfDoc) : null
      for (const page of pages) addOcrTextLayer(pdfDoc, page.pageIndex, page.lines, { fontRef, meta: op.meta })
      // The earlier layers' content streams and fonts are no longer referenced;
      // drop them so the replaced text does not linger in the file.
      if (replaced) compactUnreachable(pdfDoc)
      break
    }
    default:
      throw new Error(`Unknown PDF operation: ${operation.type}`)
  }
  pdfDoc.setProducer('simple')
  return pdfDoc.save({ useObjectStreams: true })
}

function applyDocumentEdits(pdfDoc, documentEdits = {}) {
  const { degrees } = getPdfLib()
  for (const [rawIndex, rawAmount] of Object.entries(documentEdits.pageRotations || {})) {
    const index = Number(rawIndex)
    const amount = Number(rawAmount) || 0
    if (!Number.isInteger(index) || index < 0 || index >= pdfDoc.getPageCount() || amount % 360 === 0) continue
    const page = pdfDoc.getPage(index)
    const current = page.getRotation().angle || 0
    page.setRotation(degrees(((current + amount) % 360 + 360) % 360))
  }
  if (Object.prototype.hasOwnProperty.call(documentEdits, 'bookmarks')) {
    replacePdfOutlines(pdfDoc, documentEdits.bookmarks, getPdfLib())
  }
}

function withoutNamePrefix(error, name) {
  const message = String(error?.message || error || 'This file could not be added.')
  const prefix = `${path.basename(String(name))}: `
  return message.startsWith(prefix) ? message.slice(prefix.length) : message
}

/**
 * Inserts the pages of every source at `insertIndex`. Each source converts
 * on its own: one that cannot be read is skipped and reported in `skipped`
 * (never silently, and never aborting the others).
 * @param {{name: string, sourcePath?: string, read: () => Promise<Buffer>}[]} sources
 * @returns {Promise<{data: Uint8Array, added: number, skipped: {name: string, code: string|null, reason: string}[], warnings: string[]}>}
 */
async function insertSources(baseData, insertIndex, sources) {
  const baseDoc = await loadPdf(baseData)
  const originalCount = baseDoc.getPageCount()
  const start = Math.max(0, Math.min(Number(insertIndex) || 0, originalCount))
  let targetIndex = start
  const skipped = []
  const warnings = []
  let firstError = null
  for (const source of sources) {
    try {
      const converted = await importAsPdf(await source.read(), source.name, source.sourcePath)
      const sourceDoc = await loadSourcePdf(converted.data, source.name)
      if (sourceDoc.getForm().getFields().length) {
        throw new Error(`“${source.name}” contains interactive form fields. Flatten or print it to a static PDF before adding its pages.`)
      }
      // Links between the inserted pages point at their inserted copies, not
      // at hidden duplicates of the source pages.
      const { pages } = copyPagesRemapped(baseDoc, sourceDoc, sourceDoc.getPageIndices())
      for (const page of pages) baseDoc.insertPage(targetIndex++, page)
      for (const warning of converted.warnings) warnings.push(`${source.name}: ${warning}`)
    } catch (error) {
      firstError ||= error
      skipped.push({ name: source.name, code: typeof error?.code === 'string' ? error.code : null, reason: withoutNamePrefix(error, source.name) })
    }
  }
  if (targetIndex === start) {
    if (!skipped.length) throw new Error('Drop a supported PDF or document file between the pages.')
    // One file: its own error, unchanged. Several: one message naming each.
    if (skipped.length === 1) throw firstError
    throw codedError('NOTHING_ADDED', `No pages were added. ${skipped.map((item) => `${item.name}: ${item.reason}`).join(' ')}`, { skipped })
  }
  return {
    data: await baseDoc.save({ useObjectStreams: true }),
    added: baseDoc.getPageCount() - originalCount,
    skipped,
    warnings,
  }
}

function insertDocuments(baseData, insertIndex, paths) {
  return insertSources(baseData, insertIndex, paths.map((filePath) => ({
    name: path.basename(filePath),
    sourcePath: filePath,
    read: () => fs.readFile(filePath),
  })))
}

function insertDocumentPayloads(baseData, insertIndex, inputs) {
  return insertSources(baseData, insertIndex, (Array.isArray(inputs) ? inputs : []).map((input) => ({
    name: String(input?.name || 'document.pdf'),
    read: async () => toBytes(input?.data ?? []),
  })))
}

async function exportedPages(data, indices, suggestedName) {
  const { PDFDocument } = getPdfLib()
  const sourceDoc = await loadPdf(data)
  const resultDoc = await PDFDocument.create()
  const validIndices = [...new Set(indices)]
    .filter((index) => Number.isInteger(index) && index >= 0 && index < sourceDoc.getPageCount())
    .sort((a, b) => a - b)
  if (!validIndices.length) throw new Error('Select at least one page to export.')
  // pdf-lib's copyPages() also copied every page that a link, bookmark-like
  // destination or annotation of a selected page referred to, as hidden
  // content. Copy only the selected pages and cut references to the others.
  const { pages, excludedRefs } = copyPagesRemapped(resultDoc, sourceDoc, validIndices)
  pages.forEach((page) => resultDoc.addPage(page))
  detachRemovedPages(resultDoc, excludedRefs)
  resultDoc.setTitle(safeBaseName(suggestedName || 'Exported pages'))
  resultDoc.setProducer('simple')
  return resultDoc.save({ useObjectStreams: true })
}

function parseColor(color, fallback = [1, 0.88, 0.22]) {
  if (!Array.isArray(color) || color.length !== 3) return fallback
  return color.map((channel) => Math.max(0, Math.min(1, Number(channel) || 0)))
}

function validRect(rect) {
  if (!rect || typeof rect !== 'object') return null
  const value = {
    x: Number(rect.x),
    y: Number(rect.y),
    width: Number(rect.width),
    height: Number(rect.height),
  }
  return Object.values(value).every(Number.isFinite) && value.width > 0 && value.height > 0 ? value : null
}

function textWidthAtSize(font, text, size, letterSpacing = 0, scaleX = 1) {
  try {
    const source = String(text || '')
    return (font.widthOfTextAtSize(source, size)
      + Math.max(0, source.length - 1) * letterSpacing) * scaleX
  } catch {
    return Number.POSITIVE_INFINITY
  }
}

function relativeWidthError(measured, expected) {
  if (!Number.isFinite(measured) || !Number.isFinite(expected) || expected <= 0) return Number.POSITIVE_INFINITY
  return Math.abs(measured - expected) / expected
}

function textTokens(text) {
  return String(text || '').match(/\S+|\s+/g) || []
}

function whitespaceLength(text) {
  return Array.from(String(text || '')).filter((character) => /\s/.test(character)).length
}

function separatedTextWidth(font, text, size, letterSpacing, scaleX, spaceAdvance) {
  return textTokens(text).reduce((width, token) => (
    /\s/.test(token)
      ? width + Array.from(token).length * spaceAdvance
      : width + textWidthAtSize(font, token, size, letterSpacing, scaleX)
  ), 0)
}

function hasReliableWindowsFontMatch(requestedFamily) {
  const family = String(requestedFamily || '').toLowerCase()
  if (family.includes('eb garamond')) return false
  return /times new roman|calibri|cambria|arial|helvetica|courier new|georgia|garamond|palatino|verdana|tahoma|trebuchet|segoe ui/.test(family)
}

const pageFontKeys = new WeakMap()

/** One /Font resource entry per font and page, however many lines use it. */
function pageFontKey(page, font) {
  let keys = pageFontKeys.get(page.node)
  if (!keys) pageFontKeys.set(page.node, keys = new Map())
  if (!keys.has(font.ref.tag)) keys.set(font.ref.tag, page.node.newFontDictionary(font.name, font.ref))
  return keys.get(font.ref.tag)
}

/**
 * Draw one line. `font` is a pdf-lib font or a FontRuns: consecutive runs
 * share one text object, so the text position advances exactly as it would
 * for a single string.
 */
function drawStyledTextLine(page, font, text, options) {
  const {
    beginText, endText, popGraphicsState, pushGraphicsState,
    rotateAndSkewTextRadiansAndTranslate, setCharacterSpacing,
    setCharacterSqueeze, setFillingRgbColor, setFontAndSize, showText,
  } = getPdfLib()
  const runs = (typeof font.runs === 'function' ? font.runs(text) : [{ font, text }]).filter((run) => run.text)
  if (!runs.length) return
  const operators = [
    pushGraphicsState(),
    beginText(),
    setFillingRgbColor(options.color[0], options.color[1], options.color[2]),
  ]
  if (Math.abs(options.scaleX - 1) > 0.001) operators.push(setCharacterSqueeze(options.scaleX * 100))
  if (Math.abs(options.letterSpacing) > 0.001) operators.push(setCharacterSpacing(options.letterSpacing))
  operators.push(rotateAndSkewTextRadiansAndTranslate(options.angle, 0, 0, options.x, options.y))
  for (const run of runs) {
    operators.push(setFontAndSize(pageFontKey(page, run.font), options.size), showText(run.font.encodeText(run.text)))
  }
  operators.push(endText(), popGraphicsState())
  page.pushOperators(...operators)
}

function drawStyledSeparatedLine(page, font, text, options, spaceAdvance) {
  let x = options.x
  let y = options.y
  const cosine = Math.cos(options.angle)
  const sine = Math.sin(options.angle)
  for (const token of textTokens(text)) {
    const advance = /\s/.test(token)
      ? Array.from(token).length * spaceAdvance
      : textWidthAtSize(font, token, options.size, options.letterSpacing, options.scaleX)
    if (!/\s/.test(token)) drawStyledTextLine(page, font, token, { ...options, x, y })
    x += advance * cosine
    y += advance * sine
  }
}

function dataUrlBytes(dataUrl) {
  const match = /^data:(image\/(?:png|jpe?g));base64,([A-Za-z0-9+/=\r\n]+)$/i.exec(String(dataUrl || ''))
  if (!match) throw new Error('The edited image data is not valid.')
  // An exact copy at offset 0: small decoded Buffers share Node's pool, and
  // pdf-lib's JPEG reader ignores byteOffset ('SOI not found in JPEG').
  return { mime: match[1].toLowerCase(), bytes: new Uint8Array(Buffer.from(match[2], 'base64')) }
}

function shortText(value, limit = 40) {
  const text = String(value || '').replace(/\s+/g, ' ').trim()
  return text.length > limit ? `${text.slice(0, limit - 1)}…` : text
}

function overlayDescription(overlay) {
  if (overlay.type === 'text') return overlay.originalText ? 'Edited text' : 'Added text'
  if (overlay.type === 'object') return overlay.kind === 'image' ? 'An image' : 'An object'
  if (overlay.type === 'ink') return 'A drawing'
  return 'A mark-up'
}

/** Operators pushed by a failed overlay are taken back out. */
function operatorMark(page) {
  const stream = page.contentStream
  return { stream, length: stream?.operators?.length ?? 0 }
}

function rollbackOperators(page, mark) {
  const stream = page.contentStream
  if (!stream?.operators) return
  stream.operators.length = stream === mark.stream ? Math.min(mark.length, stream.operators.length) : 0
}

/**
 * Place a point given in a text box's reading frame: `u` along the baseline
 * from the box's start edge, `v` down from its top edge. The frame is centred
 * on the PDF rect and turned by `angle`, as the editor rotates the box about
 * its centre; for 0/90/180/270 this equals the box corners used before.
 */
function textFramePoint(rect, boxWidth, boxHeight, angle, u, v) {
  const cosine = Math.cos(angle)
  const sine = Math.sin(angle)
  const along = u - boxWidth / 2
  const down = v - boxHeight / 2
  return {
    x: rect.x + rect.width / 2 + along * cosine + down * sine,
    y: rect.y + rect.height / 2 + along * sine - down * cosine,
  }
}

/**
 * The scan part of a text overlay (an edit of scanned, OCR'd text; see
 * src/lib/ocr/scanEdit.ts), checked: which recognised text it replaces
 * (`removal`), what it draws (the changed words, or the whole line) and the
 * retouch patch that hides the old printed words. Returns null for ordinary
 * text, { invalid: true } for a scan edit that cannot be written safely.
 */
function scanEditPlan(overlay) {
  const scan = overlay?.scan
  if (!scan || typeof scan !== 'object') return null
  const point = (value) => (value && Number.isFinite(Number(value.x)) && Number.isFinite(Number(value.y)) ? { x: Number(value.x), y: Number(value.y) } : null)
  const runOf = (value) => {
    if (!value || typeof value !== 'object') return null
    const origin = point(value.origin)
    const dir = point(value.dir)
    const length = Number(value.length)
    const fontSize = Number(value.fontSize)
    if (!origin || !dir || !Number.isFinite(length) || length < 0 || length > 1e5 || !(fontSize > 0 && fontSize <= 500)) return null
    const norm = Math.hypot(dir.x, dir.y)
    if (!(Math.abs(norm - 1) < 0.05)) return null
    return { origin, dir: { x: dir.x / norm, y: dir.y / norm }, length, fontSize }
  }
  const patchOf = (value) => (value && typeof value.dataUrl === 'string' && /^data:image\/png;base64,/i.test(value.dataUrl) && validRect(value.rect)
    ? { dataUrl: value.dataUrl, rect: validRect(value.rect) }
    : null)
  const replace = scan.replace && typeof scan.replace === 'object' ? scan.replace : null
  const removal = runOf(replace ? replace.run : scan.run)
  const source = replace || overlay
  const drawing = {
    text: String(source.text ?? ''),
    rect: validRect(source.rect),
    baselineOffset: Number(source.baselineOffset),
    originalText: String(source.originalText ?? ''),
    originalRect: validRect(source.originalRect),
    patch: patchOf(replace ? replace.patch : scan.patch),
  }
  if (!removal || !drawing.rect) return { invalid: true }
  return {
    mode: scan.mode === 'recognized-text' ? 'recognized-text' : 'appearance',
    removal,
    drawing,
    paper: parseColor(scan.paper, [1, 1, 1]),
    // Prepared without a patch: there was no printed ink to hide. Otherwise
    // (preparation failed or never finished) the paper colour covers the words.
    cover: scan.status === 'ready' ? null : drawing.originalRect || validRect(overlay.inkRect) || validRect(overlay.originalRect),
  }
}

/**
 * Boxes along a run of recognised text for removeNativeText(), which removes
 * the glyphs whose centres fall in a box. Short pieces of the run's band keep
 * a slanted line's box from reaching into the lines above and below it.
 */
function scanRemovalTargets(overlay, run) {
  const up = { x: -run.dir.y, y: run.dir.x }
  const pieces = Math.max(1, Math.ceil(run.length / Math.max(0.5, run.fontSize * 0.8)))
  const angle = Math.atan2(run.dir.y, run.dir.x)
  const targets = []
  for (let index = 0; index < pieces; index += 1) {
    const u0 = (run.length * index) / pieces
    const u1 = (run.length * (index + 1)) / pieces
    const corners = [[u0, -0.25], [u1, -0.25], [u0, 0.85], [u1, 0.85]].map(([u, v]) => ({
      x: run.origin.x + run.dir.x * u + up.x * v * run.fontSize,
      y: run.origin.y + run.dir.y * u + up.y * v * run.fontSize,
    }))
    const xs = corners.map((corner) => corner.x)
    const ys = corners.map((corner) => corner.y)
    const x = Math.min(...xs)
    const y = Math.min(...ys)
    // No original text: a piece over a gap between words may find nothing.
    targets.push({ type: 'text', cover: true, pageIndex: overlay.pageIndex, originalText: '', angle: Math.abs(angle) < 0.01 ? 0 : angle, originalRect: { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y } })
  }
  return targets
}

/**
 * Flatten edits into the PDF and report what could not be written as asked.
 * @returns {Promise<{ data: Uint8Array, warnings: object[], failures: object[], signatureDetected: boolean }>}
 */
async function flattenOverlaysDetailed(data, overlays = [], formValues = {}, documentEdits = {}) {
  const { layoutText, resolveTextFit } = await import('./text-layout.mjs')
  const {
    BlendMode, LineCapStyle, concatTransformationMatrix, popGraphicsState, pushGraphicsState, rgb,
  } = getPdfLib()
  const warnings = []
  const failures = []
  const list = (Array.isArray(overlays) ? overlays : []).filter((overlay) => overlay && typeof overlay === 'object')
  const skipped = new Set()

  // Edits of scanned text: their old printed words are hidden by a retouch
  // patch and only their recognised (invisible) text is removed from the
  // content, run by run. A malformed one is left out, never drawn over the
  // words it was meant to replace.
  const scanPlans = new Map()
  for (const overlay of list) {
    if (overlay.type !== 'text') continue
    const plan = scanEditPlan(overlay)
    if (!plan) continue
    if (plan.invalid) {
      skipped.add(overlay)
      failures.push(problem('OVERLAY_NOT_SAVED',
        `Edited scanned text on page ${Number(overlay.pageIndex) + 1} could not be saved: its position on the page is missing. Edit the line again.`,
        { overlayId: overlay.id, pageIndex: overlay.pageIndex }))
      continue
    }
    scanPlans.set(overlay, plan)
  }
  const removals = [
    ...list.filter((overlay) => !scanPlans.has(overlay) && !skipped.has(overlay)),
    ...[...scanPlans].flatMap(([overlay, plan]) => scanRemovalTargets(overlay, plan.removal)),
  ]

  // An edit whose original glyphs cannot be found is reported and left out;
  // drawing its replacement over the old text would garble both.
  const removed = await removeNativeText(toBytes(data), removals, {
    onUnlocated: (edit) => {
      skipped.add(edit)
      failures.push(problem('TEXT_SOURCE_NOT_FOUND',
        `The original text “${shortText(edit.originalText)}” on page ${Number(edit.pageIndex) + 1} could not be located, so this edit was not saved. Reopen the text selection and edit it again.`,
        { overlayId: edit.id, pageIndex: edit.pageIndex }))
    },
  })
  const pdfDoc = await loadPdf(removed)
  const signatureDetected = documentSignatureStatus(pdfDoc).signed
  applyDocumentEdits(pdfDoc, documentEdits)
  const fontCache = new Map()
  const fonts = new FontLibrary(pdfDoc)

  // Form values: only fields whose value really changes are touched.
  const formReport = await applyFormValues(pdfDoc, formValues, { fonts })
  warnings.push(...formReport.warnings)
  failures.push(...formReport.failures)

  // Native image edits target the image invocation itself. Remove that `Do`
  // operator before drawing its replacement so moving/deleting an image does
  // not leave the old pixels beneath a white patch. Unsupported inline/nested
  // images retain the existing visual-cover fallback below.
  const removedObjectOverlayIds = new Set()
  for (const overlay of list) {
    if (overlay.type !== 'object' || overlay.kind !== 'image' || overlay.cover === false) continue
    const sourceRect = validRect(overlay.originalRect)
    const page = pageAt(pdfDoc, overlay.pageIndex)
    if (!sourceRect || !page) continue
    try {
      if (removePageImageDraws(page, [sourceRect]) > 0) removedObjectOverlayIds.add(overlay.id)
    } catch {
      // The white cover below still hides the original image.
    }
  }

  // Retouch patches go first, under every edit, so a patch can never cover
  // the new text of another edit. Each is an RGBA image whose alpha covers
  // only the old glyphs: the scan image itself is not changed. Without a
  // patch (it could not be made) the paper colour hides the old words.
  const patchRanges = new Map()
  for (const [overlay, plan] of scanPlans) {
    if (plan.mode !== 'appearance' || skipped.has(overlay)) continue
    const page = pageAt(pdfDoc, overlay.pageIndex)
    if (!page) continue
    const mark = operatorMark(page)
    try {
      if (plan.drawing.patch) {
        const source = dataUrlBytes(plan.drawing.patch.dataUrl)
        const image = await pdfDoc.embedPng(source.bytes)
        page.drawImage(image, { ...plan.drawing.patch.rect })
      } else if (plan.cover) {
        const [r, g, b] = plan.paper
        page.drawRectangle({ x: plan.cover.x - 0.5, y: plan.cover.y - 0.5, width: plan.cover.width + 1, height: plan.cover.height + 1, color: rgb(r, g, b), borderWidth: 0 })
      }
      const stream = page.contentStream
      patchRanges.set(overlay, { stream, start: stream === mark.stream ? mark.length : 0, end: stream?.operators?.length ?? 0 })
    } catch (error) {
      rollbackOperators(page, mark)
      skipped.add(overlay)
      failures.push(problem('OVERLAY_NOT_SAVED',
        `Edited scanned text on page ${overlay.pageIndex + 1} could not be saved: ${error instanceof Error ? error.message : String(error)}`,
        { overlayId: overlay.id, pageIndex: overlay.pageIndex }))
    }
  }
  const failedPatches = []

  let glyphlessFont = null
  for (const overlay of list) {
    if (skipped.has(overlay)) continue
    const page = pageAt(pdfDoc, overlay.pageIndex)
    if (!page) {
      failures.push(problem('OVERLAY_PAGE_MISSING',
        `${overlayDescription(overlay)} belongs to page ${Number(overlay.pageIndex) + 1}, which is no longer in the document, so it was not saved.`,
        { overlayId: overlay.id, pageIndex: overlay.pageIndex }))
      continue
    }
    const mark = operatorMark(page)
    const plan = scanPlans.get(overlay)
    try {
      if (plan?.mode === 'recognized-text') {
        // Only the recognised text changes: written invisibly over the
        // scanned words it describes; the page looks exactly as before.
        const text = plan.drawing.text.replace(/\s+/g, ' ').trim()
        if (text) {
          glyphlessFont ||= addGlyphlessFont(pdfDoc)
          const run = plan.removal
          writeInvisibleRun(page, glyphlessFontKey(page, glyphlessFont), {
            text, origin: run.origin, dir: run.dir, width: Math.max(run.length, run.fontSize * 0.5), fontSize: run.fontSize,
            meta: { engine: 'Simple (corrected)', language: 'eng' },
          })
        }
      } else if (plan) {
        await drawOverlay(page, {
          ...overlay,
          text: plan.drawing.text,
          rect: plan.drawing.rect,
          baselineOffset: Number.isFinite(plan.drawing.baselineOffset) ? plan.drawing.baselineOffset : undefined,
          originalText: plan.drawing.originalText,
          originalRect: plan.drawing.originalRect || validRect(overlay.originalRect),
          // The matched Windows font, scaled as the editor showed it.
          preserveSourceMetrics: false,
          fontKey: undefined,
          fontData: undefined,
        })
      } else {
        await drawOverlay(page, overlay)
      }
    } catch (error) {
      rollbackOperators(page, mark)
      // Never leave the old words retouched away without their replacement.
      if (patchRanges.has(overlay)) failedPatches.push(patchRanges.get(overlay))
      failures.push(problem('OVERLAY_NOT_SAVED',
        `${overlayDescription(overlay)} on page ${overlay.pageIndex + 1} could not be saved: ${error instanceof Error ? error.message : String(error)}`,
        { overlayId: overlay.id, pageIndex: overlay.pageIndex }))
    }
  }
  for (const range of failedPatches.sort((a, b) => b.start - a.start)) {
    if (range.stream?.operators) range.stream.operators.splice(range.start, Math.max(0, range.end - range.start))
  }

  async function drawOverlay(page, overlay) {
    if (overlay.type === 'ink') {
      const points = Array.isArray(overlay.points)
        ? overlay.points.filter((point) => point && Number.isFinite(point.x) && Number.isFinite(point.y))
        : []
      const [r, g, b] = parseColor(overlay.color, [0.86, 0.15, 0.15])
      for (let index = 1; index < points.length; index += 1) {
        page.drawLine({
          start: points[index - 1],
          end: points[index],
          thickness: Math.max(0.5, Math.min(12, Number(overlay.thickness) || 1.8)),
          color: rgb(r, g, b),
          opacity: Number.isFinite(overlay.opacity) ? overlay.opacity : 0.95,
          lineCap: LineCapStyle.Round,
        })
      }
      return
    }

    const rect = overlay.rect || {}
    const values = [rect.x, rect.y, rect.width, rect.height]
    if (!values.every(Number.isFinite)) return

    if (overlay.type === 'highlight') {
      const [r, g, b] = parseColor(overlay.color)
      page.drawRectangle({
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
        color: rgb(r, g, b),
        opacity: Number.isFinite(overlay.opacity) ? overlay.opacity : 0.34,
        blendMode: BlendMode.Multiply,
        borderWidth: 0,
      })
    }

    if (overlay.type === 'markup') {
      const [r, g, b] = parseColor(overlay.color, [0.86, 0.15, 0.15])
      const thickness = Math.max(0.5, Math.min(12, Number(overlay.thickness) || 1.35))
      const opacity = Number.isFinite(overlay.opacity) ? overlay.opacity : 0.95
      if (overlay.style === 'rectangle') {
        page.drawRectangle({
          x: rect.x,
          y: rect.y,
          width: rect.width,
          height: rect.height,
          borderColor: rgb(r, g, b),
          borderWidth: thickness,
          borderOpacity: opacity,
          opacity: 0,
        })
      } else {
        // The rect is stored in unrotated PDF space; the line must still sit
        // under (or through) the text as it read on the rotated screen.
        const displayRotation = (((Number(overlay.displayRotation) || 0) % 360) + 360) % 360
        if (displayRotation === 90 || displayRotation === 270) {
          const inset = overlay.style === 'strikeout' ? rect.width * 0.48 : Math.max(0.8, rect.width * 0.04)
          const x = displayRotation === 90 ? rect.x + rect.width - inset : rect.x + inset
          page.drawLine({
            start: { x, y: rect.y },
            end: { x, y: rect.y + rect.height },
            thickness,
            color: rgb(r, g, b),
            opacity,
            lineCap: LineCapStyle.Round,
          })
        } else {
          const inset = overlay.style === 'strikeout' ? rect.height * 0.48 : Math.max(0.8, rect.height * 0.04)
          const y = displayRotation === 180 ? rect.y + rect.height - inset : rect.y + inset
          page.drawLine({
            start: { x: rect.x, y },
            end: { x: rect.x + rect.width, y },
            thickness,
            color: rgb(r, g, b),
            opacity,
            lineCap: LineCapStyle.Round,
          })
        }
      }
    }

    if (overlay.type === 'text') {
      const family = String(overlay.fontFamily || 'Segoe UI').toLowerCase()
      const style = {
        bold: Number(overlay.fontWeight) >= 600 || /bold|black|semibold|demi/.test(family),
        italic: overlay.fontStyle === 'italic' || /italic|oblique/.test(family),
      }
      const text = String(overlay.text || '')
      // Measured strings include the source text, so it must be covered too.
      const coverageText = `${text}\n${overlay.originalText || ''}`
      let font = await fonts.runsFor(await getPdfFont(pdfDoc, overlay.fontFamily || 'Segoe UI', fontCache, overlay), coverageText, style)
      const size = Math.max(4, Math.min(96, Number(overlay.fontSize) || Math.max(8, rect.height * 0.72)))
      const [r, g, b] = parseColor(overlay.color, [0.04, 0.04, 0.05])
      let scaleX = Math.max(0.25, Math.min(4, Number(overlay.scaleX) || 1))
      const letterSpacing = Math.max(-4, Math.min(20, Number(overlay.letterSpacing) || 0))
      const angle = Number.isFinite(overlay.angle) ? Number(overlay.angle) : 0
      const displayRotation = (((Number(overlay.displayRotation) || 0) % 360) + 360) % 360
      const sourceRect = validRect(overlay.originalRect)
      // The box is laid out in its reading frame (reading width × stacked
      // height) and then mapped into the unrotated PDF rect. A native run
      // stores the axis-aligned box of its rotated glyphs, so a vertical
      // label reads along the rect's height; a box the user rotated in the
      // inspector keeps its own width and turns about its centre. The
      // renderer may send `rectAngle` (the angle the rect was captured at).
      const rectAngle = Number.isFinite(Number(overlay.rectAngle)) ? Number(overlay.rectAngle) : sourceRect ? angle : 0
      const sideways = Math.abs(Math.sin(rectAngle + displayRotation * Math.PI / 180)) > 0.7
      const boxWidth = sideways ? rect.height : rect.width
      const boxHeight = sideways ? rect.width : rect.height
      const textAngle = angle + displayRotation * Math.PI / 180
      const preserveSourceMetrics = overlay.preserveSourceMetrics !== false && Math.abs(angle) < 0.01
      if (preserveSourceMetrics && sourceRect && overlay.originalText && overlay.fontData?.byteLength && hasReliableWindowsFontMatch(overlay.fontFamily)) {
        // PDF.js sometimes exposes a decoded browser font whose glyph outlines
        // are correct but whose synthetic space advance is not. Napoleon's
        // embedded Times face reports a 1593/2048-em space instead of 512/2048,
        // making re-embedded phrases visibly collide and compress. Compare the
        // candidate against the exact selected-run width and use the matched
        // Windows family only when it is materially closer.
        const embeddedError = relativeWidthError(
          textWidthAtSize(font, overlay.originalText, size, letterSpacing, scaleX),
          sourceRect.width,
        )
        if (embeddedError > 0.025) {
          const matchedFont = await fonts.runsFor(await getPdfFont(pdfDoc, overlay.fontFamily || 'Segoe UI', fontCache, {
            ...overlay,
            fontKey: undefined,
            fontData: undefined,
          }), coverageText, style)
          const matchedError = relativeWidthError(
            textWidthAtSize(matchedFont, overlay.originalText, size, letterSpacing, scaleX),
            sourceRect.width,
          )
          if (matchedError + 0.01 < embeddedError) font = matchedFont
        }
      }
      let sourceSpaceAdvance = null
      if (preserveSourceMetrics && sourceRect && overlay.originalText) {
        const sourceSpaceCount = whitespaceLength(overlay.originalText)
        if (sourceSpaceCount > 0) {
          const sourceWordWidth = textTokens(overlay.originalText)
            .filter((token) => !/\s/.test(token))
            .reduce((width, token) => width + textWidthAtSize(font, token, size, letterSpacing, scaleX), 0)
          const exactSourceSpace = Number(overlay.sourceSpaceWidth)
          if (Number.isFinite(exactSourceSpace) && exactSourceSpace > 0 && exactSourceSpace <= size * 1.25) {
            sourceSpaceAdvance = exactSourceSpace
            const targetWordWidth = sourceRect.width - exactSourceSpace * sourceSpaceCount
            const correction = sourceWordWidth > 0 ? targetWordWidth / sourceWordWidth : 1
            if (Number.isFinite(correction) && correction >= 0.5 && correction <= 2) {
              scaleX = Math.max(0.25, Math.min(4, scaleX * correction))
            }
          }
          const measuredSpace = (sourceRect.width - sourceWordWidth) / sourceSpaceCount
          // Never pass source spaces through a browser-decoded subset font.
          // Several real PDFs map a space to CID 0 with a huge advance when
          // that decoded face is re-embedded. Explicit word positions preserve
          // the exact glyph outlines and recover the original run's spacing.
          if (sourceSpaceAdvance === null && Number.isFinite(measuredSpace) && measuredSpace >= size * 0.12 * scaleX && measuredSpace <= size * 1.25) {
            sourceSpaceAdvance = measuredSpace
          } else if (sourceSpaceAdvance === null) {
            const nominalSpace = size * 0.25 * scaleX
            const targetWordWidth = sourceRect.width - nominalSpace * sourceSpaceCount
            const correction = sourceWordWidth > 0 ? targetWordWidth / sourceWordWidth : 1
            if (Number.isFinite(correction) && correction >= 0.5 && correction <= 2) {
              scaleX = Math.max(0.25, Math.min(4, scaleX * correction))
            }
            sourceSpaceAdvance = nominalSpace
          }
        } else {
          const sourceWidth = textWidthAtSize(font, overlay.originalText, size, letterSpacing, scaleX)
          const metricCorrection = sourceWidth > 0 ? sourceRect.width / sourceWidth : 1
          if (Number.isFinite(metricCorrection) && metricCorrection >= 0.5 && metricCorrection <= 2) {
            scaleX = Math.max(0.25, Math.min(4, scaleX * metricCorrection))
          }
        }
      }
      const fitMode = resolveTextFit(overlay)
      // Lay the text out at `factor` × its size. Size, spacing and line
      // height scale together; the first native baseline stays where it was.
      const layoutAt = (factor) => {
        const fontSize = size * factor
        const spacing = letterSpacing * factor
        const spaceAdvance = sourceSpaceAdvance === null ? null : sourceSpaceAdvance * factor
        const layout = layoutText(text, boxWidth, (line) => (
          spaceAdvance === null
            ? textWidthAtSize(font, line, fontSize, spacing, scaleX)
            : separatedTextWidth(font, line, fontSize, spacing, scaleX, spaceAdvance)
        ), fitMode)
        const lineHeight = Math.max(fontSize * 0.8, (Number(overlay.lineHeight) || size * 1.18) * factor)
        // Baseline distance measured downward from the displayed top of the box.
        // A rotated native run comes without a measured baseline; its box is
        // one line tall, with the baseline at the font's ascent.
        const baselineV = Number.isFinite(overlay.baselineOffset)
          ? boxHeight - Number(overlay.baselineOffset)
          : sourceRect && Math.abs(angle) >= 0.01
            ? Math.min(fontSize, boxHeight * NATIVE_RUN_ASCENT)
            : Math.min(fontSize, boxHeight)
        const lastContentLine = layout.lines.findLastIndex((line) => line.trim())
        // A PDF.js selection box can end above its source baseline (e.g. fonts
        // with unusual ascent metrics). Preserving that one native baseline is
        // valid; it must not make Save, Print, and every export format fail.
        const preservesNativeBaseline = sourceRect && Number.isFinite(overlay.baselineOffset) && lastContentLine === 0
        const fits = preservesNativeBaseline || lastContentLine < 0 || baselineV + lastContentLine * lineHeight <= boxHeight + 0.5
        return { fontSize, spacing, spaceAdvance, layout, lineHeight, baselineV, fits }
      }
      let chosen = layoutAt(1)
      if (!chosen.fits) {
        // Text that does not fit is made smaller, never cut off. Below the
        // minimum size it is drawn past the box and reported.
        const minimum = Math.min(1, MIN_FITTED_TEXT_SIZE / size)
        const smallest = layoutAt(minimum)
        if (!smallest.fits) {
          chosen = smallest
          warnings.push(problem('TEXT_OVERFLOW',
            `Text on page ${overlay.pageIndex + 1} (“${shortText(text)}”) is too long for its box even at ${Math.round(smallest.fontSize * 10) / 10} pt, so it runs past the box. Enlarge the box to keep it inside.`,
            { overlayId: overlay.id, pageIndex: overlay.pageIndex, fontSize: smallest.fontSize, originalFontSize: size }))
        } else {
          let low = minimum
          let high = 1
          for (let step = 0; step < 14; step += 1) {
            const middle = (low + high) / 2
            if (layoutAt(middle).fits) low = middle
            else high = middle
          }
          chosen = layoutAt(low)
          warnings.push(problem('TEXT_SHRUNK',
            `Text on page ${overlay.pageIndex + 1} (“${shortText(text)}”) did not fit its box and was reduced from ${Math.round(size * 10) / 10} pt to ${Math.round(chosen.fontSize * 10) / 10} pt.`,
            { overlayId: overlay.id, pageIndex: overlay.pageIndex, fontSize: chosen.fontSize, originalFontSize: size }))
        }
      }
      const lineScaleX = scaleX * chosen.layout.fitScale
      const lineSpaceAdvance = chosen.spaceAdvance === null ? null : chosen.spaceAdvance * chosen.layout.fitScale
      let baselineV = chosen.baselineV
      for (const line of chosen.layout.lines) {
        const lineWidth = lineSpaceAdvance === null
          ? textWidthAtSize(font, line, chosen.fontSize, chosen.spacing, lineScaleX)
          : separatedTextWidth(font, line, chosen.fontSize, chosen.spacing, lineScaleX, lineSpaceAdvance)
        const align = overlay.align === 'center' || overlay.align === 'right' ? overlay.align : 'left'
        const lineU = align === 'center'
          ? Math.max(0, (boxWidth - lineWidth) / 2)
          : align === 'right'
            ? Math.max(0, boxWidth - lineWidth)
            : 0
        const origin = textFramePoint(rect, boxWidth, boxHeight, textAngle, lineU, baselineV)
        const drawOptions = {
          x: origin.x,
          y: origin.y,
          size: chosen.fontSize,
          color: [r, g, b],
          scaleX: lineScaleX,
          letterSpacing: chosen.spacing,
          angle: textAngle,
        }
        if (lineSpaceAdvance === null) drawStyledTextLine(page, font, line, drawOptions)
        else drawStyledSeparatedLine(page, font, line, drawOptions, lineSpaceAdvance)
        baselineV += chosen.lineHeight
      }
      const missing = [...new Set(Array.from(text))].filter((character) => font.missing.has(character))
      if (missing.length) {
        failures.push(problem('MISSING_GLYPHS',
          `Text on page ${overlay.pageIndex + 1} contains ${missing.length === 1 ? 'a character' : 'characters'} (${missing.join(' ')}) that no installed font can draw; ${missing.length === 1 ? 'it was' : 'they were'} replaced by “�”.`,
          {
            overlayId: overlay.id,
            pageIndex: overlay.pageIndex,
            chars: missing,
            blockingMessage: `Text on page ${overlay.pageIndex + 1} contains ${missing.length === 1 ? 'a character' : 'characters'} (${missing.join(' ')}) that no installed font can draw. Remove ${missing.length === 1 ? 'it' : 'them'} and save again.`,
          }))
      }
    }

    if (overlay.type === 'object') {
      const coverRect = validRect(overlay.originalRect)
      if (overlay.cover !== false && coverRect && !removedObjectOverlayIds.has(overlay.id)) {
        page.drawRectangle({
          x: coverRect.x - 0.75,
          y: coverRect.y - 0.75,
          width: coverRect.width + 1.5,
          height: coverRect.height + 1.5,
          color: rgb(1, 1, 1),
          opacity: 1,
          borderWidth: 0,
        })
      }
      if (overlay.dataUrl) {
        const source = dataUrlBytes(overlay.dataUrl)
        const image = source.mime === 'image/png'
          ? await pdfDoc.embedPng(source.bytes)
          : await pdfDoc.embedJpg(source.bytes)
        const opacity = Math.max(0.05, Math.min(1, Number(overlay.opacity) || 1))
        // Chromium shows (and sizes the box for) a JPEG turned by its EXIF
        // orientation; pdf-lib embeds the stored pixels, so turn them here.
        const orientation = source.mime === 'image/png' ? 1 : jpegOrientation(source.bytes)
        if (orientation === 1) {
          page.drawImage(image, { x: rect.x, y: rect.y, width: rect.width, height: rect.height, opacity })
        } else {
          const swapped = orientation >= 5
          const rawWidth = swapped ? rect.height : rect.width
          const rawHeight = swapped ? rect.width : rect.height
          const placement = orientationPlacement(orientation, rawWidth, rawHeight)
          page.pushOperators(
            pushGraphicsState(),
            concatTransformationMatrix(1, 0, 0, 1, rect.x, rect.y),
            concatTransformationMatrix(...placement.matrix),
          )
          page.drawImage(image, { x: 0, y: 0, width: rawWidth, height: rawHeight, opacity })
          page.pushOperators(popGraphicsState())
        }
      }
    }
  }

  pdfDoc.setProducer('simple')
  // Replaced content streams, removed images and old field appearances are
  // no longer referenced; drop them so the edited-away content is not kept.
  await pdfDoc.flush()
  compactUnreachable(pdfDoc)
  const output = await pdfDoc.save({ useObjectStreams: true, updateFieldAppearances: false })
  if (formReport.written.length) failures.push(...await verifyFormValues(output, formReport.written))
  return { data: output, warnings, failures, signatureDetected }
}

/**
 * Legacy contract (bytes only). Callers that cannot show a problem report
 * get an error instead of a file that silently lacks an edit: any failure,
 * or a warning that changed what the user typed, aborts with a clear message.
 */
async function flattenOverlays(data, overlays = [], formValues = {}, documentEdits = {}) {
  const result = await flattenOverlaysDetailed(data, overlays, formValues, documentEdits)
  const blocking = [...result.failures, ...result.warnings.filter((warning) => warning.dataLoss)]
  if (blocking.length) throw unsavedChangesError(blocking)
  return result.data
}

// Every user-visible write goes through the shared verified write: temp file in the
// same folder, flush, read-back check, then replace with retries while another
// program (another PDF reader, a scanner, a backup tool) holds the file.
async function atomicWrite(targetPath, data) {
  await safeWriteFile(targetPath, toBytes(data))
}

function printerCapabilityHints(options = {}) {
  const flattened = Object.entries(options)
    .map(([key, value]) => `${key}=${String(value)}`)
    .join(' ')
    .toLowerCase()
  return {
    supportsDuplex: /duplex|two[-_ ]?sided|sides[-_]supported/.test(flattened),
    supportsColor: /(^|\W)colou?r/.test(flattened) || !/monochrome|black[-_ ]?and[-_ ]?white/.test(flattened),
  }
}

async function printPdfDirect(data, documentName, options = {}) {
  const printDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-print-'))
  const printPath = path.join(printDirectory, `${safeBaseName(documentName || 'Document')}.pdf`)
  let printWindow = null
  try {
    const printBytes = await preparePrintPdf(toBytes(data), options)
    await fs.writeFile(printPath, toBytes(printBytes))
    printWindow = new BrowserWindow({
      width: 800,
      height: 600,
      show: false,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: true,
        plugins: true,
        devTools: false,
      },
    })
    await loadPdfViewerForPrint(printWindow, pathToFileURL(printPath).href)
    return await printWebContentsSilently(printWindow.webContents, nativePrintOptions(options))
  } catch (error) {
    return { success: false, failureReason: error instanceof Error ? error.message : String(error) }
  } finally {
    if (printWindow && !printWindow.isDestroyed()) printWindow.destroy()
    await fs.rm(printDirectory, { recursive: true, force: true }).catch(() => {})
  }
}

async function cleanupStalePrintDirectories() {
  let entries = []
  try { entries = await fs.readdir(os.tmpdir(), { withFileTypes: true }) } catch { return }
  const cutoff = Date.now() - 24 * 60 * 60 * 1_000
  await Promise.all(entries
    .filter((entry) => entry.isDirectory() && entry.name.startsWith('simple-print-'))
    .map(async (entry) => {
      const target = path.join(os.tmpdir(), entry.name)
      try {
        const stat = await fs.stat(target)
        if (stat.mtimeMs < cutoff) await fs.rm(target, { recursive: true, force: true })
      } catch {}
    }))
}

function createWindow(openPath = null) {
  const browserWindow = new BrowserWindow({
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    width: 1440,
    height: 920,
    minWidth: 900,
    minHeight: 620,
    show: false,
    frame: false,
    backgroundColor: '#f4f4f5',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: !app.isPackaged,
      additionalArguments: bridgeArguments(IO_MODULE),
    },
  })
  // Save / Don't Save / Cancel on every close path, never closing during a save,
  // and crash/hang/sign-out handling.
  guard.installWindowGuard(browserWindow)

  browserWindow.removeMenu()
  if (process.env.VITE_DEV_SERVER_URL) {
    browserWindow.loadURL(process.env.VITE_DEV_SERVER_URL)
  } else {
    browserWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))
  }

  browserWindow.once('ready-to-show', () => browserWindow.show())
  browserWindow.webContents.on('will-navigate', (event) => event.preventDefault())
  browserWindow.webContents.once('did-finish-load', () => {
    if (openPath) {
      // The preload buffers this event until React subscribes, so launching the
      // app from a PDF no longer needs an arbitrary half-second safety delay.
      browserWindow.webContents.send('file:open-external', openPath)
    }
  })
  browserWindow.on('maximize', () => browserWindow.webContents.send('window:maximized', true))
  browserWindow.on('unmaximize', () => browserWindow.webContents.send('window:maximized', false))
  const rendererId = browserWindow.webContents.id
  browserWindow.on('closed', () => {
    for (const [id, session] of imageExportSessions) {
      if (session.ownerId !== rendererId) continue
      imageExportSessions.delete(id)
      if (session.directory) void fs.rm(session.targetPath, { recursive: true, force: true }).catch(() => {})
    }
  })
  return browserWindow
}

function callingWindow(event) {
  const browserWindow = BrowserWindow.fromWebContents(event.sender)
  return browserWindow && !browserWindow.isDestroyed() ? browserWindow : null
}

function showOpenDialogFor(event, options) {
  const browserWindow = callingWindow(event)
  return browserWindow ? dialog.showOpenDialog(browserWindow, options) : dialog.showOpenDialog(options)
}

function showSaveDialogFor(event, options) {
  const browserWindow = callingWindow(event)
  return browserWindow ? dialog.showSaveDialog(browserWindow, options) : dialog.showSaveDialog(options)
}

function withRequiredExtension(filePath, extension) {
  return filePath.toLowerCase().endsWith(extension) ? filePath : `${filePath}${extension}`
}

async function uniqueExportDirectory(parentDirectory, requestedName) {
  const baseName = safeExportBaseName(requestedName)
  for (let suffix = 0; suffix < 1_000; suffix += 1) {
    const candidate = path.join(parentDirectory, suffix ? `${baseName} (${suffix + 1})` : baseName)
    try {
      await fs.mkdir(candidate)
      return candidate
    } catch (error) {
      if (error?.code !== 'EEXIST') throw error
    }
  }
  throw new Error('Could not create a unique export folder.')
}

async function exportTextDocument(event, input = {}) {
  const format = String(input.format || '').toLowerCase()
  if (!TEXT_EXPORT_FORMATS.has(format)) throw new Error('This document export format is not supported.')
  const baseName = safeExportBaseName(input.baseName || input.title || 'Exported PDF')
  const metadata = {
    docx: { extension: '.docx', filter: { name: 'Word document', extensions: ['docx'] } },
    txt: { extension: '.txt', filter: { name: 'Plain text', extensions: ['txt'] } },
    md: { extension: '.md', filter: { name: 'Markdown document', extensions: ['md'] } },
    html: { extension: '.html', filter: { name: 'Web page', extensions: ['html'] } },
  }[format]
  const result = await showSaveDialogFor(event, {
    title: `Export as ${format.toUpperCase()}`,
    defaultPath: `${baseName}${metadata.extension}`,
    filters: [metadata.filter],
  })
  if (result.canceled || !result.filePath) return null
  const targetPath = withRequiredExtension(result.filePath, metadata.extension)
  // Markdown keeps its pictures as files in a folder beside it, named after it.
  let imageFolder
  if (format === 'md') {
    const stem = safeExportBaseName(path.basename(targetPath, path.extname(targetPath)))
    for (let suffix = 1; suffix < 1_000 && !imageFolder; suffix += 1) {
      const candidate = suffix === 1 ? `${stem} images` : `${stem} images (${suffix})`
      try { await fs.access(path.join(path.dirname(targetPath), candidate)) } catch { imageFolder = candidate }
    }
  }
  const output = await buildTextExportFiles(format, input.pages, input.title || baseName, { imageFolder })
  if (output.assets.length && imageFolder) {
    const folder = path.join(path.dirname(targetPath), imageFolder)
    await fs.mkdir(folder, { recursive: true })
    for (const asset of output.assets) await atomicWrite(path.join(folder, asset.name), asset.data)
  }
  await atomicWrite(targetPath, output.data)
  return targetPath
}

async function beginImageExport(event, input = {}) {
  const format = String(input.format || '').toLowerCase()
  if (!IMAGE_EXPORT_FORMATS.has(format)) throw new Error('This image export format is not supported.')
  const pageCount = Math.trunc(Number(input.pageCount))
  if (!Number.isInteger(pageCount) || pageCount < 1 || pageCount > 10_000) {
    throw new Error('Choose between 1 and 10,000 pages to export.')
  }
  const activeForWindow = [...imageExportSessions.values()].filter((session) => session.ownerId === event.sender.id).length
  if (activeForWindow >= 4 || imageExportSessions.size >= 32) throw new Error('Finish or cancel another image export before starting a new one.')
  const baseName = safeExportBaseName(input.baseName || 'Exported PDF')
  const extension = format === 'jpeg' ? '.jpg' : `.${format}`
  let targetPath
  let directory = false
  if (pageCount === 1) {
    const result = await showSaveDialogFor(event, {
      title: `Export page as ${format.toUpperCase()}`,
      defaultPath: `${baseName}${extension}`,
      filters: [{ name: format === 'jpeg' ? 'JPEG image' : `${format.toUpperCase()} image`, extensions: [extension.slice(1)] }],
    })
    if (result.canceled || !result.filePath) return null
    targetPath = withRequiredExtension(result.filePath, extension)
  } else {
    const result = await showOpenDialogFor(event, {
      title: `Choose where to export ${pageCount} ${format.toUpperCase()} images`,
      buttonLabel: 'Choose folder',
      properties: ['openDirectory', 'createDirectory'],
    })
    if (result.canceled || !result.filePaths[0]) return null
    targetPath = await uniqueExportDirectory(result.filePaths[0], `${baseName} - ${format.toUpperCase()} pages`)
    directory = true
  }
  const id = crypto.randomUUID()
  imageExportSessions.set(id, {
    ownerId: event.sender.id,
    format,
    pageCount,
    baseName,
    targetPath,
    directory,
    writtenPages: new Set(),
  })
  return { id, targetPath }
}

function imageExportSession(event, id) {
  const session = imageExportSessions.get(String(id || ''))
  if (!session || session.ownerId !== event.sender.id) throw new Error('This image export session is no longer available.')
  return session
}

async function writeImageExportPage(event, input = {}) {
  const session = imageExportSession(event, input.id)
  const pageNumber = Math.trunc(Number(input.pageNumber))
  if (!Number.isInteger(pageNumber) || pageNumber < 1 || pageNumber > 1_000_000 || session.writtenPages.has(pageNumber) || session.writtenPages.size >= session.pageCount) {
    throw new Error('The exported page number is invalid or duplicated.')
  }
  const data = toBytes(input.data)
  if (!data.length || data.length > 100 * 1024 * 1024) throw new Error('The rendered page image is empty or too large.')
  const validSignature = session.format === 'png'
    ? data.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
    : session.format === 'jpeg'
      ? data[0] === 0xff && data[1] === 0xd8
      : data.subarray(0, 4).toString('ascii') === 'RIFF' && data.subarray(8, 12).toString('ascii') === 'WEBP'
  if (!validSignature) throw new Error(`The rendered page is not a valid ${session.format.toUpperCase()} image.`)
  const targetPath = session.directory
    ? path.join(session.targetPath, imageExportFileName(session.baseName, pageNumber, session.pageCount, session.format))
    : session.targetPath
  await atomicWrite(targetPath, data)
  session.writtenPages.add(pageNumber)
  return targetPath
}

async function finishImageExport(event, id) {
  const session = imageExportSession(event, id)
  if (session.writtenPages.size !== session.pageCount) {
    throw new Error(`Only ${session.writtenPages.size} of ${session.pageCount} page images were written.`)
  }
  imageExportSessions.delete(String(id))
  return session.targetPath
}

async function cancelImageExport(event, id) {
  const session = imageExportSessions.get(String(id || ''))
  if (!session || session.ownerId !== event.sender.id) return false
  imageExportSessions.delete(String(id))
  if (session.directory) await fs.rm(session.targetPath, { recursive: true, force: true }).catch(() => {})
  return true
}

function supportedPaths(argv) {
  return [...new Set(argv.filter((argument) => isImportableName(argument)))]
}

/** File-picker filters for documents; office-engine formats only when one is installed. */
async function documentDialogFilters() {
  return importDialogFilters({ engine: await officeEngineAvailable() })
}

async function hasIntactSignature(filePath) {
  try {
    return (await fileSignatureStatus(filePath)).intact
  } catch {
    return false
  }
}

let textBackgroundCacheInstance = null

/** Text-removal previews, cached per document revision and page. */
function textBackgroundCache() {
  textBackgroundCacheInstance ||= createTextBackgroundCache({
    parse: (bytes) => loadPdf(bytes),
    extractPage: async (source, pageIndex) => {
      const single = await getPdfLib().PDFDocument.create()
      const { pages, excludedRefs } = copyPagesRemapped(single, source, [pageIndex])
      single.addPage(pages[0])
      detachRemovedPages(single, excludedRefs)
      return single.save()
    },
    removeText: (pageBytes, edits) => removeNativeText(pageBytes, edits),
  })
  return textBackgroundCacheInstance
}

function registerIpc() {
  ipcMain.handle('file:open-dialog', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Open a document',
      properties: ['openFile'],
      filters: await documentDialogFilters(),
    })
    if (result.canceled || !result.filePaths[0]) return null
    return filePayload(result.filePaths[0])
  })

  ipcMain.handle('file:open-in-new-window', async (event, requestedPath = null) => {
    let selectedPath = requestedPath
    if (!selectedPath) {
      const result = await showOpenDialogFor(event, {
        title: 'Open a document',
        properties: ['openFile'],
        filters: await documentDialogFilters(),
      })
      if (result.canceled || !result.filePaths[0]) return false
      selectedPath = result.filePaths[0]
    }
    if (!(await openablePath(selectedPath))) throw unsupportedFormatError(selectedPath)
    createWindow(selectedPath)
    return true
  })

  ipcMain.handle('file:open-path', async (_event, filePath) => filePayload(filePath))

  ipcMain.handle('file:open-bytes', async (_event, input) => {
    const buffer = toBytes(input.data)
    const name = String(input?.name || 'Document')
    const result = await importAsPdf(buffer, name)
    return {
      data: serializableBytes(result.data),
      name: result.converted ? `${safeBaseName(name)}.pdf` : name,
      path: null,
      sourcePath: null,
      converted: result.converted,
      ...(result.warnings.length ? { warnings: result.warnings } : {}),
      ...documentProtection(result.converted ? '' : '.pdf', buffer),
    }
  })

  ipcMain.handle('pdf:unlock', async (_event, data, password) => {
    const result = await require('./pdf-unlock.cjs').unlockPdf(data, typeof password === 'string' ? password : '')
    return result.data ? { status: result.status, data: serializableBytes(result.data) } : result
  })

  ipcMain.handle('pdf:mutate', async (_event, data, operation) => serializableBytes(await applyMutation(data, operation)))
  ipcMain.handle('pdf:text-background', async (_event, data, pageIndex, edits) => (
    serializableBytes(await textBackgroundCache().render(data, pageIndex, edits))
  ))
  // `report: true` (as a fifth argument, or inside documentEdits so the
  // current preload passes it through) returns { ok, data, warnings,
  // failures, signatureDetected } instead of throwing on the first problem.
  ipcMain.handle('pdf:flatten-overlays', async (_event, data, overlays, formValues, documentEdits, options) => {
    const report = options?.report === true || documentEdits?.report === true
    if (!report) return serializableBytes(await flattenOverlays(data, overlays, formValues, documentEdits))
    const result = await flattenOverlaysDetailed(data, overlays, formValues, documentEdits)
    return {
      ok: result.failures.length === 0,
      data: serializableBytes(result.data),
      warnings: result.warnings,
      failures: result.failures,
      signatureDetected: result.signatureDetected,
    }
  })

  // Both insert handlers return { data, added, skipped, warnings }: files that
  // could not be added are listed in `skipped` with a plain reason.
  ipcMain.handle('pdf:insert-files', async (event, data, insertIndex) => {
    const result = await showOpenDialogFor(event, {
      title: 'Add pages',
      properties: ['openFile', 'multiSelections'],
      filters: await documentDialogFilters(),
    })
    if (result.canceled || !result.filePaths.length) return null
    const inserted = await insertDocuments(data, insertIndex, result.filePaths)
    return { ...inserted, data: serializableBytes(inserted.data) }
  })

  ipcMain.handle('pdf:insert-dropped-files', async (_event, data, insertIndex, inputs) => {
    const inserted = await insertDocumentPayloads(data, insertIndex, inputs)
    return { ...inserted, data: serializableBytes(inserted.data) }
  })

  // Any picture format becomes PNG or JPEG here (WebP, GIF, BMP, AVIF, ICO and
  // SVG through Chromium, TIFF's first page directly), so the editor and the
  // flattener only ever see those two.
  ipcMain.handle('file:pick-image', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Choose an image',
      properties: ['openFile'],
      filters: importDialogFilters({ imagesOnly: true }),
    })
    if (result.canceled || !result.filePaths[0]) return null
    const filePath = result.filePaths[0]
    const image = await pickedImageData(await fs.readFile(filePath), path.basename(filePath))
    return {
      dataUrl: `data:${image.mime};base64,${image.data.toString('base64')}`,
      name: path.basename(filePath),
    }
  })

  ipcMain.handle('pdf:export-as-pdf', async (event, input = {}) => {
    const suggestedName = `${safeExportBaseName(input.suggestedName || 'Exported PDF')}.pdf`
    const result = await showSaveDialogFor(event, {
      title: 'Export as PDF',
      defaultPath: suggestedName,
      filters: [{ name: 'PDF file', extensions: ['pdf'] }],
    })
    if (result.canceled || !result.filePath) return null
    const targetPath = withRequiredExtension(result.filePath, '.pdf')
    const output = input.fullDocument
      ? toBytes(input.data)
      : await exportedPages(input.data, input.indices, suggestedName)
    await atomicWrite(targetPath, output)
    return targetPath
  })

  ipcMain.handle('pdf:export-text-document', (event, input) => exportTextDocument(event, input))
  ipcMain.handle('pdf:begin-image-export', (event, input) => beginImageExport(event, input))
  ipcMain.handle('pdf:write-image-export-page', (event, input) => writeImageExportPage(event, input))
  ipcMain.handle('pdf:finish-image-export', (event, id) => finishImageExport(event, id))
  ipcMain.handle('pdf:cancel-image-export', (event, id) => cancelImageExport(event, id))

  ipcMain.handle('pdf:export-pages', async (event, data, indices, suggestedName) => {
    const result = await showSaveDialogFor(event, {
      title: 'Export pages',
      defaultPath: suggestedName || 'Exported pages.pdf',
      filters: [{ name: 'PDF file', extensions: ['pdf'] }],
    })
    if (result.canceled || !result.filePath) return null
    await atomicWrite(result.filePath, await exportedPages(data, indices, suggestedName))
    return result.filePath
  })

  ipcMain.handle('pdf:start-page-drag', async (event, data, indices, suggestedName) => {
    if (!dragExportDirectory) dragExportDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-page-drag-'))
    const dragDirectory = await fs.mkdtemp(path.join(dragExportDirectory, 'drag-'))
    const fileName = `${safeBaseName(suggestedName || 'Exported page')}.pdf`
    const targetPath = path.join(dragDirectory, fileName)
    await atomicWrite(targetPath, await exportedPages(data, indices, fileName))
    const iconPath = path.join(__dirname, '..', 'build', 'icon.ico')
    const icon = nativeImage.createFromPath(iconPath)
    event.sender.startDrag({ file: targetPath, icon })
    return targetPath
  })

  ipcMain.handle('print:list-printers', async (event) => {
    const printers = await event.sender.getPrintersAsync()
    return printers.map((printer) => ({
      name: printer.name,
      displayName: printer.displayName || printer.name,
      ...printerCapabilityHints(printer.options),
    }))
  })

  ipcMain.handle('pdf:print-direct', async (event, data, documentName, options = {}) => {
    try {
      const printer = resolvePrinter(await event.sender.getPrintersAsync(), options.deviceName)
      return printPdfDirect(data, documentName, { ...options, deviceName: printer.name })
    } catch (error) {
      return { success: false, failureReason: error instanceof Error ? error.message : String(error) }
    }
  })

  ipcMain.handle('pdf:save', async (event, input) => {
    let targetPath = input.forceDialog ? null : input.path
    let defaultPath = input.name || 'Untitled.pdf'
    let keptSignedOriginal = false
    if (targetPath && await hasIntactSignature(targetPath)) {
      // Never overwrite a validly signed original without asking: a full
      // rewrite destroys the signed revision. Offer a copy beside it.
      defaultPath = path.join(path.dirname(targetPath), `${path.basename(targetPath, path.extname(targetPath))} (edited).pdf`)
      targetPath = null
      keptSignedOriginal = true
    }
    if (!targetPath) {
      const result = await showSaveDialogFor(event, {
        title: input.forceDialog || keptSignedOriginal ? 'Save PDF as' : 'Save PDF',
        defaultPath,
        filters: [{ name: 'PDF file', extensions: ['pdf'] }],
      })
      if (result.canceled || !result.filePath) return null
      targetPath = result.filePath.toLowerCase().endsWith('.pdf') ? result.filePath : `${result.filePath}.pdf`
    }
    await atomicWrite(targetPath, input.data)
    return { path: targetPath, name: path.basename(targetPath), ...(keptSignedOriginal ? { keptSignedOriginal } : {}) }
  })

  ipcMain.handle('shell:show-item', (_event, filePath) => shell.showItemInFolder(filePath))
  ipcMain.handle('shell:open-external', async (_event, url) => {
    const parsed = new URL(url)
    if (!['http:', 'https:', 'mailto:'].includes(parsed.protocol)) throw new Error('This link type is not allowed.')
    await shell.openExternal(parsed.toString())
  })
  ipcMain.handle('app:get-version', () => app.getVersion())
  ipcMain.on('window:minimize', (event) => callingWindow(event)?.minimize())
  ipcMain.on('window:toggle-maximize', (event) => {
    const browserWindow = callingWindow(event)
    if (!browserWindow) return
    if (browserWindow.isMaximized()) browserWindow.unmaximize()
    else browserWindow.maximize()
  })
  ipcMain.on('window:confirm-close', (event) => {
    const browserWindow = callingWindow(event)
    if (!browserWindow) return
    // The window guard asks the page again; after Save or Discard it reports no changes.
    browserWindow.close()
  })
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (_event, argv) => {
    const incoming = supportedPaths(argv)
    if (incoming.length) {
      for (const filePath of incoming) createWindow(filePath)
      return
    }
    const browserWindow = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0]
    if (browserWindow) {
      if (browserWindow.isMinimized()) browserWindow.restore()
      browserWindow.focus()
    }
  })

  app.whenReady().then(async () => {
    void cleanupStalePrintDirectories()
    registerIpc()
    registerSharedIo({
      ipcMain,
      module: IO_MODULE,
      guard,
      stores,
      officeEngine,
      openInWindow: (filePath) => { createWindow(filePath); return true },
    })
    // Finish or undo any save a crash interrupted before documents open.
    await sweep().catch((error) => console.error('[simple-io] sweep failed', error))
    const incoming = supportedPaths(process.argv)
    if (incoming.length) {
      for (const filePath of incoming) createWindow(filePath)
    } else {
      createWindow()
    }
    if (process.platform === 'darwin') {
      app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow()
      })
    }
  })
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', () => {
  isQuitting = true
  if (dragExportDirectory) fs.rm(dragExportDirectory, { recursive: true, force: true }).catch(() => {})
})
