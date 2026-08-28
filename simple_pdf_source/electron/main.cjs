const { app, BrowserWindow, dialog, ipcMain, nativeImage, shell } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const crypto = require('node:crypto')
const { pathToFileURL } = require('node:url')
const { replacePdfOutlines } = require('./pdf-outlines.cjs')
const { removePageImageDraws } = require('./pdf-content-edits.cjs')

const SUPPORTED_EXTENSIONS = new Set([
  '.pdf', '.png', '.jpg', '.jpeg', '.txt', '.md', '.docx', '.doc',
])
const PDF_SIGNATURE_FIELD = Buffer.from('/Type /Sig')

const closeApprovedWindows = new WeakSet()
let isQuitting = false
let mammothModule
let WordExtractorModule
let fontkitModule
let pdfLibModule
let dragExportDirectory = null

// Conversion/editing dependencies are intentionally loaded only when a user
// invokes those features. In particular, mammoth and fontkit are expensive to
// initialise and are not needed at all for the common PDF viewing path.
function getMammoth() {
  mammothModule ||= require('mammoth')
  return mammothModule
}

function getWordExtractor() {
  WordExtractorModule ||= require('word-extractor')
  return WordExtractorModule
}

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
  const embeddedKey = style.fontKey ? `embedded:${style.fontKey}:${style.fontData?.byteLength || 0}` : ''
  if (embeddedKey && style.fontData?.byteLength) {
    try {
      const fontBytes = toBytes(style.fontData)
      const decodedFont = getFontkit().create(fontBytes)
      const missingRequiredGlyph = Array.from(String(style.text || ''))
        .some((character) => !/\s/.test(character) && !decodedFont.hasGlyphForCodePoint(character.codePointAt(0)))
      if (missingRequiredGlyph) throw new Error('Decoded PDF subset does not contain every replacement glyph.')
      if (cache.has(embeddedKey)) return cache.get(embeddedKey)
      pdfDoc.registerFontkit(getFontkit())
      const embedded = await pdfDoc.embedFont(fontBytes, { subset: true })
      cache.set(embeddedKey, embedded)
      return embedded
    } catch {
      // Some PDFs expose a browser-only decoded face. Use a matched Windows
      // family below when fontkit cannot re-embed those bytes.
    }
  }

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

  if (cache.has(definition.key)) return cache.get(definition.key)
  const candidates = definition.files.map((fileName) => path.join(windowsDir, 'Fonts', fileName))

  for (const candidate of candidates) {
    try {
      const fontBytes = await fs.readFile(candidate)
      pdfDoc.registerFontkit(getFontkit())
      const font = await pdfDoc.embedFont(fontBytes, { subset: true })
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

function splitLongToken(token, font, fontSize, maxWidth) {
  const chunks = []
  let chunk = ''
  for (const character of token) {
    if (font.widthOfTextAtSize(chunk + character, fontSize) > maxWidth && chunk) {
      chunks.push(chunk)
      chunk = character
    } else {
      chunk += character
    }
  }
  if (chunk) chunks.push(chunk)
  return chunks
}

function wrapParagraph(paragraph, font, fontSize, maxWidth) {
  if (!paragraph) return ['']
  const sourceWords = paragraph.trim().split(/\s+/)
  const words = sourceWords.flatMap((word) => (
    font.widthOfTextAtSize(word, fontSize) > maxWidth
      ? splitLongToken(word, font, fontSize, maxWidth)
      : [word]
  ))
  const lines = []
  let line = ''
  for (const word of words) {
    const candidate = line ? `${line} ${word}` : word
    if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth) {
      line = candidate
    } else {
      if (line) lines.push(line)
      line = word
    }
  }
  if (line) lines.push(line)
  return lines.length ? lines : ['']
}

async function textToPdfBytes(text, title = 'Converted document') {
  const { PDFDocument, rgb } = getPdfLib()
  const pdfDoc = await PDFDocument.create()
  const font = await getPdfFont(pdfDoc)
  const pageWidth = 595.28
  const pageHeight = 841.89
  const margin = 54
  const fontSize = 11
  const lineHeight = 16
  const maxWidth = pageWidth - margin * 2
  const paragraphs = String(text || '').replace(/\r\n?/g, '\n').split('\n')
  const lines = paragraphs.flatMap((paragraph) => wrapParagraph(paragraph, font, fontSize, maxWidth))
  let page = pdfDoc.addPage([pageWidth, pageHeight])
  let y = pageHeight - margin

  for (const line of lines) {
    if (y < margin + lineHeight) {
      page = pdfDoc.addPage([pageWidth, pageHeight])
      y = pageHeight - margin
    }
    try {
      page.drawText(line, { x: margin, y, font, size: fontSize, color: rgb(0.07, 0.07, 0.08) })
    } catch {
      const compatibleLine = line.replace(/[^\x20-\x7E]/g, '?')
      page.drawText(compatibleLine, { x: margin, y, font, size: fontSize, color: rgb(0.07, 0.07, 0.08) })
    }
    y -= lineHeight
  }

  pdfDoc.setTitle(title)
  pdfDoc.setCreator('simple')
  pdfDoc.setProducer('simple')
  return pdfDoc.save()
}

async function imageToPdfBytes(buffer, extension, title = 'Converted image') {
  const { PDFDocument } = getPdfLib()
  const pdfDoc = await PDFDocument.create()
  const lower = extension.toLowerCase()
  const image = lower === '.png'
    ? await pdfDoc.embedPng(buffer)
    : await pdfDoc.embedJpg(buffer)
  const dimensions = image.scale(1)
  const maxWidth = 841.89
  const maxHeight = 841.89
  const scale = Math.min(1, maxWidth / dimensions.width, maxHeight / dimensions.height)
  const width = dimensions.width * scale
  const height = dimensions.height * scale
  const page = pdfDoc.addPage([width, height])
  page.drawImage(image, { x: 0, y: 0, width, height })
  pdfDoc.setTitle(title)
  pdfDoc.setCreator('simple')
  return pdfDoc.save()
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

async function htmlToPdfBytes(bodyHtml, title = 'Converted document') {
  const conversionDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-convert-'))
  const htmlPath = path.join(conversionDirectory, 'document.html')
  const html = `<!doctype html>
<html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>
@page { size: A4; margin: 18mm 17mm 19mm; }
html { color: #111113; background: #fff; font-family: "Segoe UI", Arial, sans-serif; }
body { margin: 0; font-size: 10.5pt; line-height: 1.48; overflow-wrap: anywhere; }
h1 { font-size: 22pt; line-height: 1.18; margin: 0 0 12pt; }
h2 { font-size: 17pt; line-height: 1.22; margin: 17pt 0 8pt; }
h3 { font-size: 13.5pt; line-height: 1.25; margin: 14pt 0 6pt; }
p { margin: 0 0 8pt; }
ul, ol { margin: 0 0 9pt; padding-left: 22pt; }
li { margin: 0 0 3pt; }
table { width: 100%; margin: 8pt 0 12pt; border-collapse: collapse; break-inside: avoid; }
th, td { padding: 5pt 6pt; border: 0.6pt solid #c9c9ce; vertical-align: top; }
th { background: #f4f4f5; font-weight: 650; }
img { max-width: 100%; height: auto; break-inside: avoid; }
blockquote { margin: 9pt 0; padding: 2pt 0 2pt 12pt; border-left: 2pt solid #d4d4d8; color: #52525b; }
a { color: #18181b; text-decoration: underline; }
pre { padding: 8pt; border: 0.6pt solid #e4e4e7; background: #fafafa; white-space: pre-wrap; }
</style></head><body>${bodyHtml}</body></html>`
  let conversionWindow = null
  try {
    await fs.writeFile(htmlPath, html, 'utf8')
    conversionWindow = new BrowserWindow({
      show: false,
      width: 900,
      height: 1100,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: true,
        javascript: false,
        devTools: false,
      },
    })
    await conversionWindow.loadFile(htmlPath)
    await new Promise((resolve) => setTimeout(resolve, 180))
    return new Uint8Array(await conversionWindow.webContents.printToPDF({
      printBackground: true,
      pageSize: 'A4',
      preferCSSPageSize: true,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
    }))
  } finally {
    if (conversionWindow && !conversionWindow.isDestroyed()) conversionWindow.destroy()
    await fs.rm(conversionDirectory, { recursive: true, force: true }).catch(() => {})
  }
}

async function convertInputToPdf(buffer, extension, name) {
  const ext = extension.toLowerCase()
  if (ext === '.pdf') return buffer
  if (['.png', '.jpg', '.jpeg'].includes(ext)) {
    return imageToPdfBytes(buffer, ext, safeBaseName(name))
  }
  if (['.txt', '.md'].includes(ext)) {
    return textToPdfBytes(buffer.toString('utf8'), safeBaseName(name))
  }
  if (ext === '.docx') {
    const mammoth = getMammoth()
    const result = await mammoth.convertToHtml({ buffer }, {
      convertImage: mammoth.images.imgElement(async (image) => ({
        src: `data:${image.contentType};base64,${await image.read('base64')}`,
      })),
    })
    return htmlToPdfBytes(result.value, safeBaseName(name))
  }
  if (ext === '.doc') {
    const WordExtractor = getWordExtractor()
    const extractor = new WordExtractor()
    const extracted = await extractor.extract(buffer)
    return textToPdfBytes(extracted.getBody(), safeBaseName(name))
  }
  throw new Error(`Unsupported file type: ${ext || 'unknown'}`)
}

async function filePayload(filePath) {
  const ext = path.extname(filePath).toLowerCase()
  if (!SUPPORTED_EXTENSIONS.has(ext)) throw new Error('This file type is not supported.')
  const original = await fs.readFile(filePath)
  const converted = ext !== '.pdf'
  // Keep the fs Buffer for PDFs. Constructing a Uint8Array from it used to
  // duplicate the complete file before Electron performed its unavoidable IPC
  // serialization copy.
  const pdfBytes = converted ? await convertInputToPdf(original, ext, filePath) : original
  const signatureDetected = ext === '.pdf' && original.includes(PDF_SIGNATURE_FIELD)
  return {
    data: serializableBytes(pdfBytes),
    name: converted ? `${safeBaseName(filePath)}.pdf` : path.basename(filePath),
    path: converted ? null : filePath,
    sourcePath: filePath,
    converted,
    signatureDetected,
  }
}

async function loadPdf(data) {
  const { PDFDocument } = getPdfLib()
  return PDFDocument.load(toBytes(data), {
    updateMetadata: false,
    throwOnInvalidObject: false,
  })
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
      indices.forEach((index) => pdfDoc.removePage(index))
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

async function insertDocuments(baseData, insertIndex, paths) {
  const { PDFDocument } = getPdfLib()
  const baseDoc = await loadPdf(baseData)
  const originalCount = baseDoc.getPageCount()
  let targetIndex = Math.max(0, Math.min(Number(insertIndex) || 0, baseDoc.getPageCount()))
  for (const filePath of paths) {
    const ext = path.extname(filePath).toLowerCase()
    const sourceBuffer = await fs.readFile(filePath)
    const sourceBytes = await convertInputToPdf(sourceBuffer, ext, filePath)
    const sourceDoc = await PDFDocument.load(sourceBytes)
    const pages = await baseDoc.copyPages(sourceDoc, sourceDoc.getPageIndices())
    for (const page of pages) baseDoc.insertPage(targetIndex++, page)
  }
  return {
    data: await baseDoc.save({ useObjectStreams: true }),
    added: baseDoc.getPageCount() - originalCount,
  }
}

async function insertDocumentPayloads(baseData, insertIndex, inputs) {
  const { PDFDocument } = getPdfLib()
  const baseDoc = await loadPdf(baseData)
  const originalCount = baseDoc.getPageCount()
  let targetIndex = Math.max(0, Math.min(Number(insertIndex) || 0, originalCount))
  for (const input of Array.isArray(inputs) ? inputs : []) {
    const name = String(input?.name || 'document.pdf')
    const ext = path.extname(name).toLowerCase()
    if (!SUPPORTED_EXTENSIONS.has(ext)) continue
    const sourceBytes = await convertInputToPdf(toBytes(input.data), ext, name)
    const sourceDoc = await PDFDocument.load(sourceBytes, { updateMetadata: false, throwOnInvalidObject: false })
    const pages = await baseDoc.copyPages(sourceDoc, sourceDoc.getPageIndices())
    for (const page of pages) baseDoc.insertPage(targetIndex++, page)
  }
  if (targetIndex === Math.max(0, Math.min(Number(insertIndex) || 0, originalCount))) {
    throw new Error('Drop a supported PDF or document file between the pages.')
  }
  return {
    data: await baseDoc.save({ useObjectStreams: true }),
    added: baseDoc.getPageCount() - originalCount,
  }
}

async function exportedPages(data, indices, suggestedName) {
  const { PDFDocument } = getPdfLib()
  const sourceDoc = await loadPdf(data)
  const resultDoc = await PDFDocument.create()
  const validIndices = [...new Set(indices)]
    .filter((index) => Number.isInteger(index) && index >= 0 && index < sourceDoc.getPageCount())
    .sort((a, b) => a - b)
  if (!validIndices.length) throw new Error('Select at least one page to export.')
  const pages = await resultDoc.copyPages(sourceDoc, validIndices)
  pages.forEach((page) => resultDoc.addPage(page))
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

function drawStyledTextLine(page, font, text, options) {
  const {
    beginText, endText, popGraphicsState, pushGraphicsState,
    rotateAndSkewTextRadiansAndTranslate, setCharacterSpacing,
    setCharacterSqueeze, setFillingRgbColor, setFontAndSize, showText,
  } = getPdfLib()
  const fontKey = page.node.newFontDictionary(font.name, font.ref)
  const operators = [
    pushGraphicsState(),
    beginText(),
    setFillingRgbColor(options.color[0], options.color[1], options.color[2]),
    setFontAndSize(fontKey, options.size),
  ]
  if (Math.abs(options.scaleX - 1) > 0.001) operators.push(setCharacterSqueeze(options.scaleX * 100))
  if (Math.abs(options.letterSpacing) > 0.001) operators.push(setCharacterSpacing(options.letterSpacing))
  operators.push(
    rotateAndSkewTextRadiansAndTranslate(options.angle, 0, 0, options.x, options.y),
    showText(font.encodeText(text)),
    endText(),
    popGraphicsState(),
  )
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
  return { mime: match[1].toLowerCase(), bytes: Buffer.from(match[2], 'base64') }
}

function wrapText(text, font, size, maxWidth) {
  const output = []
  for (const paragraph of String(text || '').replace(/\r\n?/g, '\n').split('\n')) {
    if (!paragraph) {
      output.push('')
      continue
    }
    const words = paragraph.split(/\s+/)
    let line = ''
    for (const word of words) {
      const candidate = line ? `${line} ${word}` : word
      if (!line || font.widthOfTextAtSize(candidate, size) <= maxWidth) {
        line = candidate
        continue
      }
      output.push(line)
      line = word
    }
    if (line) output.push(line)
  }
  return output
}

async function flattenOverlays(data, overlays = [], formValues = {}, documentEdits = {}) {
  const { BlendMode, LineCapStyle, rgb } = getPdfLib()
  const pdfDoc = await loadPdf(data)
  applyDocumentEdits(pdfDoc, documentEdits)
  const fontCache = new Map()
  const formFont = await getPdfFont(pdfDoc, 'Segoe UI', fontCache)

  try {
    const form = pdfDoc.getForm()
    for (const [name, value] of Object.entries(formValues || {})) {
      try {
        const field = form.getField(name)
        if (typeof field.setText === 'function') field.setText(String(value))
        else if (typeof field.check === 'function' && typeof field.uncheck === 'function') {
          if (value) field.check()
          else field.uncheck()
        } else if (typeof field.select === 'function') field.select(String(value))
      } catch {
        // A malformed or duplicated field should not prevent the rest of the save.
      }
    }
    if (Object.keys(formValues || {}).length) form.updateFieldAppearances(formFont)
  } catch {
    // Documents without a valid AcroForm simply skip this step.
  }

  // Native image edits target the image invocation itself. Remove that `Do`
  // operator before drawing its replacement so moving/deleting an image does
  // not leave the old pixels beneath a white patch. Unsupported inline/nested
  // images retain the existing visual-cover fallback below.
  const removedObjectOverlayIds = new Set()
  for (const overlay of overlays) {
    if (overlay?.type !== 'object' || overlay.kind !== 'image' || overlay.cover === false) continue
    const sourceRect = validRect(overlay.originalRect)
    const page = pdfDoc.getPage(overlay.pageIndex)
    if (!sourceRect || !page) continue
    if (removePageImageDraws(page, [sourceRect]) > 0) removedObjectOverlayIds.add(overlay.id)
  }

  for (const overlay of overlays) {
    const page = pdfDoc.getPage(overlay.pageIndex)
    if (!page) continue

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
      continue
    }

    const rect = overlay.rect || {}
    const values = [rect.x, rect.y, rect.width, rect.height]
    if (!values.every(Number.isFinite)) continue

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
      let font = await getPdfFont(pdfDoc, overlay.fontFamily || 'Segoe UI', fontCache, overlay)
      const size = Math.max(4, Math.min(96, Number(overlay.fontSize) || Math.max(8, rect.height * 0.72)))
      const [r, g, b] = parseColor(overlay.color, [0.04, 0.04, 0.05])
      let scaleX = Math.max(0.25, Math.min(4, Number(overlay.scaleX) || 1))
      const letterSpacing = Math.max(-4, Math.min(24, Number(overlay.letterSpacing) || 0))
      const angle = Number.isFinite(overlay.angle) ? Number(overlay.angle) : 0
      // Text typed on a rotated page reads along the rotated axes. Lay lines
      // out in the displayed box (reading width × stacked height) and map the
      // result back into the unrotated rect below.
      const displayRotation = (((Number(overlay.displayRotation) || 0) % 360) + 360) % 360
      const sideways = displayRotation === 90 || displayRotation === 270
      const boxWidth = sideways ? rect.height : rect.width
      const boxHeight = sideways ? rect.width : rect.height
      const sourceRect = validRect(overlay.originalRect)
      if (sourceRect && overlay.originalText && overlay.fontData?.byteLength && hasReliableWindowsFontMatch(overlay.fontFamily)) {
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
          const matchedFont = await getPdfFont(pdfDoc, overlay.fontFamily || 'Segoe UI', fontCache, {
            ...overlay,
            fontKey: undefined,
            fontData: undefined,
          })
          const matchedError = relativeWidthError(
            textWidthAtSize(matchedFont, overlay.originalText, size, letterSpacing, scaleX),
            sourceRect.width,
          )
          if (matchedError + 0.01 < embeddedError) font = matchedFont
        }
      }
      let sourceSpaceAdvance = null
      if (sourceRect && overlay.originalText) {
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
      if (overlay.cover !== false) {
        const coverRect = sourceRect || rect
        const [coverR, coverG, coverB] = parseColor(overlay.backgroundColor, [1, 1, 1])
        // The renderer supplies the exact selected glyph range. Cover only that
        // range plus a hairline for antialiasing; broad font-relative padding
        // visibly erased neighboring words and lines on dense book pages.
        // Italic glyphs can overhang a PDF.js text item's nominal rectangle by
        // about one point. Use separate horizontal/vertical bleed so deletion
        // removes those edge pixels without broadly erasing adjacent lines.
        const horizontalBleed = Math.max(0.75, Math.min(1.5, size * 0.065))
        const verticalBleed = Math.max(0.25, Math.min(0.75, size * 0.03))
        page.drawRectangle({
          x: coverRect.x - horizontalBleed,
          y: coverRect.y - verticalBleed,
          width: coverRect.width + horizontalBleed * 2,
          height: coverRect.height + verticalBleed * 2,
          color: rgb(coverR, coverG, coverB),
          opacity: 1,
          borderWidth: 0,
        })
      }
      const nativeSingleLine = typeof overlay.originalText === 'string'
        && !overlay.originalText.includes('\n')
        && !String(overlay.text || '').includes('\n')
      if (nativeSingleLine && overlay.text) {
        const replacementWidth = sourceSpaceAdvance === null
          ? textWidthAtSize(font, overlay.text, size, letterSpacing, scaleX)
          : separatedTextWidth(font, overlay.text, size, letterSpacing, scaleX, sourceSpaceAdvance)
        if (replacementWidth > boxWidth && replacementWidth > 0) {
          const fit = boxWidth / replacementWidth
          scaleX = Math.max(0.25, scaleX * fit)
          if (sourceSpaceAdvance !== null) sourceSpaceAdvance *= fit
        }
      }
      const lines = nativeSingleLine
        ? [String(overlay.text || '')]
        : wrapText(overlay.text, font, size, Math.max(4, boxWidth / scaleX))
      const lineHeight = Math.max(size * 0.8, Number(overlay.lineHeight) || size * 1.18)
      // Baseline distance measured downward from the displayed top of the box.
      let baselineV = Number.isFinite(overlay.baselineOffset)
        ? boxHeight - Number(overlay.baselineOffset)
        : Math.min(size, boxHeight)
      for (const line of lines) {
        if (baselineV > boxHeight + lineHeight) break
        const lineWidth = sourceSpaceAdvance === null
          ? textWidthAtSize(font, line, size, letterSpacing, scaleX)
          : separatedTextWidth(font, line, size, letterSpacing, scaleX, sourceSpaceAdvance)
        const align = overlay.align === 'center' || overlay.align === 'right' ? overlay.align : 'left'
        const lineU = align === 'center'
          ? Math.max(0, (boxWidth - lineWidth) / 2)
          : align === 'right'
            ? Math.max(0, boxWidth - lineWidth)
            : 0
        const textX = displayRotation === 90
          ? rect.x + baselineV
          : displayRotation === 180
            ? rect.x + rect.width - lineU
            : displayRotation === 270
              ? rect.x + rect.width - baselineV
              : rect.x + lineU
        const textY = displayRotation === 90
          ? rect.y + lineU
          : displayRotation === 180
            ? rect.y + baselineV
            : displayRotation === 270
              ? rect.y + rect.height - lineU
              : rect.y + boxHeight - baselineV
        const drawOptions = {
          x: textX,
          y: textY,
          size,
          color: [r, g, b],
          scaleX,
          letterSpacing,
          angle: angle + displayRotation * Math.PI / 180,
        }
        if (sourceSpaceAdvance === null) drawStyledTextLine(page, font, line, drawOptions)
        else drawStyledSeparatedLine(page, font, line, drawOptions, sourceSpaceAdvance)
        baselineV += lineHeight
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
        page.drawImage(image, {
          x: rect.x,
          y: rect.y,
          width: rect.width,
          height: rect.height,
          opacity: Math.max(0.05, Math.min(1, Number(overlay.opacity) || 1)),
        })
      }
    }
  }
  pdfDoc.setProducer('simple')
  return pdfDoc.save({ useObjectStreams: true })
}

async function atomicWrite(targetPath, data) {
  const directory = path.dirname(targetPath)
  const extension = path.extname(targetPath) || '.pdf'
  const tempPath = path.join(directory, `.${path.basename(targetPath, extension)}-${crypto.randomUUID()}${extension}.tmp`)
  try {
    await fs.writeFile(tempPath, toBytes(data))
    await fs.rename(tempPath, targetPath)
  } catch (error) {
    await fs.rm(tempPath, { force: true }).catch(() => {})
    throw error
  }
}

async function printPdfBytes(data, documentName, ownerWindow) {
  const printDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-print-'))
  const printPath = path.join(printDirectory, `${safeBaseName(documentName || 'Document')}.pdf`)
  let printWindow = null
  const cleanup = async () => {
    await fs.rm(printDirectory, { recursive: true, force: true }).catch(() => {})
  }

  try {
    await fs.writeFile(printPath, toBytes(data))
    printWindow = new BrowserWindow({
      ...(ownerWindow && !ownerWindow.isDestroyed() ? { parent: ownerWindow } : {}),
      width: 960,
      height: 780,
      minWidth: 680,
      minHeight: 520,
      show: false,
      title: `Print preview — ${documentName || 'PDF'}`,
      backgroundColor: '#f4f4f5',
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: true,
        plugins: true,
        devTools: false,
      },
    })
    printWindow.removeMenu()
    printWindow.on('closed', cleanup)
    await printWindow.loadURL(pathToFileURL(printPath).href)
    printWindow.show()
    // Keep the preview open so Chromium's PDF viewer owns the complete print
    // lifecycle. Its toolbar opens the native Windows print dialog and avoids
    // Electron's unreliable callback when that dialog is cancelled.
    return true
  } catch (error) {
    if (printWindow && !printWindow.isDestroyed()) printWindow.close()
    await cleanup()
    throw error
  }
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
    const indices = Array.isArray(options.pageIndices) && options.pageIndices.length ? options.pageIndices : null
    const printBytes = indices ? await exportedPages(data, indices, documentName || 'Document') : data
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
    await printWindow.loadURL(pathToFileURL(printPath).href)
    // Chromium's PDF viewer keeps parsing after did-finish-load; printing
    // immediately produces blank pages, so give it a moment to settle.
    await new Promise((resolve) => setTimeout(resolve, 400))
    return await new Promise((resolve) => {
      printWindow.webContents.print({
        silent: true,
        ...(options.deviceName ? { deviceName: String(options.deviceName) } : {}),
        copies: Math.max(1, Math.trunc(Number(options.copies) || 1)),
        ...(typeof options.landscape === 'boolean' ? { landscape: options.landscape } : {}),
        color: options.color !== false,
        ...(options.duplexMode ? { duplexMode: options.duplexMode } : {}),
        collate: options.collate !== false,
        margins: { marginType: 'default' },
      }, (success, failureReason) => resolve({ success, failureReason: failureReason || '' }))
    })
  } catch (error) {
    return { success: false, failureReason: error instanceof Error ? error.message : String(error) }
  } finally {
    if (printWindow && !printWindow.isDestroyed()) printWindow.destroy()
    await fs.rm(printDirectory, { recursive: true, force: true }).catch(() => {})
  }
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
    },
  })

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
  browserWindow.on('close', (event) => {
    if (isQuitting || closeApprovedWindows.has(browserWindow)) return
    event.preventDefault()
    browserWindow.webContents.send('window:close-requested')
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

function supportedPaths(argv) {
  return [...new Set(argv.filter((argument) => SUPPORTED_EXTENSIONS.has(path.extname(argument).toLowerCase())))]
}

function registerIpc() {
  ipcMain.handle('file:open-dialog', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Open a document',
      properties: ['openFile'],
      filters: [
        { name: 'Documents', extensions: ['pdf', 'png', 'jpg', 'jpeg', 'docx', 'doc', 'txt', 'md'] },
        { name: 'PDF files', extensions: ['pdf'] },
        { name: 'Images', extensions: ['png', 'jpg', 'jpeg'] },
        { name: 'Word files', extensions: ['docx', 'doc'] },
        { name: 'Text files', extensions: ['txt', 'md'] },
      ],
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
        filters: [
          { name: 'Documents', extensions: ['pdf', 'png', 'jpg', 'jpeg', 'docx', 'doc', 'txt', 'md'] },
          { name: 'PDF files', extensions: ['pdf'] },
          { name: 'Images', extensions: ['png', 'jpg', 'jpeg'] },
          { name: 'Word files', extensions: ['docx', 'doc'] },
          { name: 'Text files', extensions: ['txt', 'md'] },
        ],
      })
      if (result.canceled || !result.filePaths[0]) return false
      selectedPath = result.filePaths[0]
    }
    if (!SUPPORTED_EXTENSIONS.has(path.extname(selectedPath).toLowerCase())) {
      throw new Error('This file type is not supported.')
    }
    createWindow(selectedPath)
    return true
  })

  ipcMain.handle('file:open-path', async (_event, filePath) => filePayload(filePath))

  ipcMain.handle('file:open-bytes', async (_event, input) => {
    const buffer = toBytes(input.data)
    const ext = path.extname(input.name).toLowerCase()
    const converted = ext !== '.pdf'
    const pdfBytes = await convertInputToPdf(buffer, ext, input.name)
    return {
      data: serializableBytes(pdfBytes),
      name: converted ? `${safeBaseName(input.name)}.pdf` : input.name,
      path: null,
      sourcePath: null,
      converted,
      signatureDetected: ext === '.pdf' && buffer.includes(PDF_SIGNATURE_FIELD),
    }
  })

  ipcMain.handle('pdf:mutate', async (_event, data, operation) => serializableBytes(await applyMutation(data, operation)))
  ipcMain.handle('pdf:flatten-overlays', async (_event, data, overlays, formValues, documentEdits) => serializableBytes(await flattenOverlays(data, overlays, formValues, documentEdits)))

  ipcMain.handle('pdf:insert-files', async (event, data, insertIndex) => {
    const result = await showOpenDialogFor(event, {
      title: 'Add pages',
      properties: ['openFile', 'multiSelections'],
      filters: [{ name: 'Documents', extensions: ['pdf', 'png', 'jpg', 'jpeg', 'docx', 'txt', 'md'] }],
    })
    if (result.canceled || !result.filePaths.length) return null
    const inserted = await insertDocuments(data, insertIndex, result.filePaths)
    return { data: serializableBytes(inserted.data), added: inserted.added }
  })

  ipcMain.handle('pdf:insert-dropped-files', async (_event, data, insertIndex, inputs) => {
    const inserted = await insertDocumentPayloads(data, insertIndex, inputs)
    return { data: serializableBytes(inserted.data), added: inserted.added }
  })

  ipcMain.handle('file:pick-image', async (event) => {
    const result = await showOpenDialogFor(event, {
      title: 'Choose an image',
      properties: ['openFile'],
      filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg'] }],
    })
    if (result.canceled || !result.filePaths[0]) return null
    const filePath = result.filePaths[0]
    const extension = path.extname(filePath).toLowerCase()
    const mime = extension === '.png' ? 'image/png' : 'image/jpeg'
    const imageBytes = await fs.readFile(filePath)
    return {
      dataUrl: `data:${mime};base64,${imageBytes.toString('base64')}`,
      name: path.basename(filePath),
    }
  })

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

  ipcMain.handle('pdf:print', async (event, data, documentName) => printPdfBytes(data, documentName, callingWindow(event)))

  ipcMain.handle('print:list-printers', async (event) => {
    const printers = await event.sender.getPrintersAsync()
    return printers.map((printer) => ({
      name: printer.name,
      displayName: printer.displayName || printer.name,
      isDefault: Boolean(printer.isDefault),
      ...printerCapabilityHints(printer.options),
    }))
  })

  ipcMain.handle('pdf:print-direct', async (_event, data, documentName, options) => printPdfDirect(data, documentName, options))

  ipcMain.handle('pdf:save', async (event, input) => {
    let targetPath = input.forceDialog ? null : input.path
    if (!targetPath) {
      const result = await showSaveDialogFor(event, {
        title: input.forceDialog ? 'Save PDF as' : 'Save PDF',
        defaultPath: input.name || 'Untitled.pdf',
        filters: [{ name: 'PDF file', extensions: ['pdf'] }],
      })
      if (result.canceled || !result.filePath) return null
      targetPath = result.filePath.toLowerCase().endsWith('.pdf') ? result.filePath : `${result.filePath}.pdf`
    }
    await atomicWrite(targetPath, input.data)
    return { path: targetPath, name: path.basename(targetPath) }
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
    closeApprovedWindows.add(browserWindow)
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

  app.whenReady().then(() => {
    registerIpc()
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
