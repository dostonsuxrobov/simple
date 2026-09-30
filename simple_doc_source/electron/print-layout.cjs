const { PDFDocument, StandardFonts, rgb } = require('pdf-lib')

const MAX_PDF_BYTES = 512 * 1024 * 1024
const POINTS_PER_INCH = 72
const MICRONS_PER_INCH = 25_400
const PAPER_SIZES = Object.freeze({
  Document: Object.freeze({ label: 'Document size', width: 612, height: 792 }),
  Letter: Object.freeze({ label: 'Letter', width: 612, height: 792 }),
  A4: Object.freeze({ label: 'A4', width: 595.28, height: 841.89 }),
  Legal: Object.freeze({ label: 'Legal', width: 612, height: 1008 }),
  A5: Object.freeze({ label: 'A5', width: 419.53, height: 595.28 }),
  Tabloid: Object.freeze({ label: 'Tabloid', width: 792, height: 1224 }),
})
const MARGIN_PRESETS = Object.freeze({
  none: Object.freeze({ top: 0, right: 0, bottom: 0, left: 0 }),
  narrow: Object.freeze({ top: 0.35, right: 0.35, bottom: 0.35, left: 0.35 }),
  normal: Object.freeze({ top: 0.75, right: 0.75, bottom: 0.75, left: 0.75 }),
  wide: Object.freeze({ top: 1, right: 1, bottom: 1, left: 1 }),
})

function finiteNumber(value, fallback) {
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : fallback
}

function clamp(value, minimum, maximum) {
  return Math.max(minimum, Math.min(maximum, value))
}

function cleanTitle(value) {
  const cleaned = String(value || 'Untitled document')
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  return Array.from(cleaned).slice(0, 180).join('') || 'Untitled document'
}

function validatePdfBytes(value) {
  let bytes
  if (Buffer.isBuffer(value)) bytes = value
  else if (ArrayBuffer.isView(value)) bytes = Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  else if (value instanceof ArrayBuffer) bytes = Buffer.from(value)
  else if (value && value.type === 'Buffer' && Array.isArray(value.data)) bytes = Buffer.from(value.data)
  else throw new Error('Print data must be a PDF byte array.')
  if (bytes.byteLength < 5 || bytes.byteLength > MAX_PDF_BYTES) throw new Error('The print PDF has an invalid size.')
  if (bytes.subarray(0, 5).toString('ascii') !== '%PDF-') throw new Error('The print data is not a PDF document.')
  return bytes
}

function normalizeMargins(input, paperWidth, paperHeight) {
  const preset = String(input?.preset || 'normal')
  const source = preset === 'custom'
    ? input?.custom
    : MARGIN_PRESETS[preset] || MARGIN_PRESETS.normal
  const maxHorizontal = Math.min(3, Math.max(0, paperWidth / 72 / 2 - 0.25))
  const maxVertical = Math.min(3, Math.max(0, paperHeight / 72 / 2 - 0.25))
  return {
    preset: preset === 'custom' || MARGIN_PRESETS[preset] ? preset : 'normal',
    top: clamp(finiteNumber(source?.top, 0.75), 0, maxVertical),
    right: clamp(finiteNumber(source?.right, 0.75), 0, maxHorizontal),
    bottom: clamp(finiteNumber(source?.bottom, 0.75), 0, maxVertical),
    left: clamp(finiteNumber(source?.left, 0.75), 0, maxHorizontal),
  }
}

function normalizePrintSettings(input = {}, sourceSize = undefined) {
  const paper = Object.hasOwn(PAPER_SIZES, input.paper) ? input.paper : 'Document'
  const sourceWidth = Math.max(1, finiteNumber(sourceSize?.width, PAPER_SIZES.Letter.width))
  const sourceHeight = Math.max(1, finiteNumber(sourceSize?.height, PAPER_SIZES.Letter.height))
  const requestedOrientation = input.orientation === 'landscape' ? 'landscape' : 'portrait'
  const orientation = paper === 'Document'
    ? (sourceWidth > sourceHeight ? 'landscape' : 'portrait')
    : requestedOrientation
  const base = PAPER_SIZES[paper]
  const paperWidth = paper === 'Document' ? sourceWidth : orientation === 'landscape' ? base.height : base.width
  const paperHeight = paper === 'Document' ? sourceHeight : orientation === 'landscape' ? base.width : base.height
  const scaling = ['fit', 'actual', 'custom'].includes(input.scaling) ? input.scaling : 'fit'
  return {
    paper,
    orientation,
    paperWidth,
    paperHeight,
    margins: normalizeMargins(input.margins, paperWidth, paperHeight),
    scaling,
    scalePercent: clamp(Math.round(finiteNumber(input.scalePercent, 100)), 25, 200),
    pages: input.pages === 'custom' ? 'custom' : 'all',
    pageRange: String(input.pageRange || '').trim().slice(0, 240),
    printTitle: Boolean(input.printTitle),
    printPageNumbers: Boolean(input.printPageNumbers),
    centerContent: input.centerContent !== false,
  }
}

function parsePageRange(value, totalPages) {
  if (!Number.isSafeInteger(totalPages) || totalPages < 1) throw new Error('The PDF has no printable pages.')
  const source = String(value || '').trim()
  if (!source) throw new Error('Enter a page range, such as 1-3, 5.')
  const selected = new Set()
  for (const token of source.split(',')) {
    const part = token.trim()
    const match = part.match(/^(\d+)(?:\s*-\s*(\d+))?$/)
    if (!match) throw new Error(`“${part || token}” is not a valid page or range.`)
    const start = Number(match[1])
    const end = Number(match[2] || match[1])
    if (start < 1 || end < start) throw new Error(`“${part}” is not a valid ascending page range.`)
    if (end > totalPages) throw new Error(`Page ${end} is outside this ${totalPages}-page document.`)
    for (let page = start; page <= end; page += 1) selected.add(page - 1)
  }
  if (!selected.size) throw new Error('Select at least one page.')
  return [...selected].sort((left, right) => left - right)
}

function fitScale(sourceWidth, sourceHeight, availableWidth, availableHeight, settings) {
  if (settings.scaling === 'actual') return 1
  if (settings.scaling === 'custom') return settings.scalePercent / 100
  return Math.min(availableWidth / sourceWidth, availableHeight / sourceHeight)
}

function fontSupportsText(font, text) {
  try {
    font.encodeText(text)
    return true
  } catch {
    return false
  }
}

function safeTextForFont(value, font) {
  let result = ''
  for (const symbol of Array.from(String(value || ''))) {
    if (fontSupportsText(font, symbol)) {
      result += symbol
      continue
    }
    const decomposed = symbol.normalize('NFKD').replace(/\p{M}/gu, '')
    let fallback = ''
    for (const candidate of Array.from(decomposed)) {
      if (fontSupportsText(font, candidate)) fallback += candidate
    }
    result += fallback || '?'
  }
  return result
}

function trimToWidth(text, font, size, maxWidth) {
  if (font.widthOfTextAtSize(text, size) <= maxWidth) return text
  const symbols = Array.from(text)
  const suffix = '...'
  let low = 0
  let high = symbols.length
  while (low < high) {
    const midpoint = Math.ceil((low + high) / 2)
    if (font.widthOfTextAtSize(`${symbols.slice(0, midpoint).join('')}${suffix}`, size) <= maxWidth) low = midpoint
    else high = midpoint - 1
  }
  return `${symbols.slice(0, low).join('')}${suffix}`
}

function pointsToMicrons(value, label) {
  const points = Number(value)
  if (!Number.isFinite(points) || points <= 0 || points > POINTS_PER_INCH * 200) {
    throw new Error(`${label} must be a valid composed PDF page dimension.`)
  }
  return Math.max(353, Math.round(points / POINTS_PER_INCH * MICRONS_PER_INCH))
}

function nativePrintOptions(input = {}) {
  const paper = Object.hasOwn(PAPER_SIZES, input.paper) ? input.paper : null
  if (!paper) throw new Error('The composed print job has an invalid paper size.')
  const options = {
    silent: true,
    printBackground: true,
    margins: { marginType: 'none' },
    scaleFactor: 100,
    landscape: input.orientation === 'landscape',
    copies: Math.min(999, Math.max(1, Math.trunc(Number(input.copies) || 1))),
    color: input.color !== false,
    collate: input.collate !== false,
  }
  if (input.deviceName) options.deviceName = String(input.deviceName)
  if (['simplex', 'shortEdge', 'longEdge'].includes(input.duplexMode)) options.duplexMode = input.duplexMode
  if (paper !== 'Document') {
    options.pageSize = paper
    return options
  }
  if (input.mixedPaperSizes) {
    // Electron exposes one native page size per job. Let the selected printer use its
    // default paper for mixed-size PDFs instead of claiming page one represents them all.
    options.usePrinterDefaultPageSize = true
    return options
  }
  const width = pointsToMicrons(input.paperWidth, 'Paper width')
  const height = pointsToMicrons(input.paperHeight, 'Paper height')
  options.pageSize = { width: Math.min(width, height), height: Math.max(width, height) }
  return options
}

function resolvePrinter(printers, requestedDeviceName = '') {
  const available = Array.isArray(printers)
    ? printers.filter((printer) => printer && typeof printer.name === 'string' && printer.name.trim())
    : []
  if (!available.length) throw new Error('No printer is available. Add or enable a printer in Windows, then try again.')
  const requested = String(requestedDeviceName || '')
  if (requested) {
    const selected = available.find((printer) => printer.name === requested)
    if (!selected) throw new Error('The selected printer is no longer available. Choose another printer and try again.')
    return selected
  }
  // Electron 36+ no longer exposes PrinterInfo.isDefault. Omitting deviceName
  // from a silent job delegates the choice to the Windows default printer.
  return { name: '', displayName: 'Default Windows printer', supportsDuplex: true, supportsColor: true }
}

async function composePrintPdf(input) {
  const bytes = validatePdfBytes(input?.data)
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const sourcePageCount = source.getPageCount()
  if (sourcePageCount < 1) throw new Error('The PDF has no printable pages.')
  const documentFirstPage = source.getPage(0)
  let settings = normalizePrintSettings(input?.settings, { width: documentFirstPage.getWidth(), height: documentFirstPage.getHeight() })
  const selected = settings.pages === 'custom'
    ? parsePageRange(settings.pageRange, sourcePageCount)
    : Array.from({ length: sourcePageCount }, (_value, index) => index)
  const firstSelectedPage = source.getPage(selected[0])
  settings = normalizePrintSettings(input?.settings, { width: firstSelectedPage.getWidth(), height: firstSelectedPage.getHeight() })
  const output = await PDFDocument.create()
  const font = await output.embedFont(StandardFonts.Helvetica)
  const title = cleanTitle(input?.name)
  const printableTitle = safeTextForFont(title, font)
  const headerBand = settings.printTitle ? 18 : 0
  const footerBand = settings.printPageNumbers ? 18 : 0

  const embeddedPages = await output.embedPdf(source, selected)
  const outputSizes = embeddedPages.map((embedded) => settings.paper === 'Document'
    ? { width: embedded.width, height: embedded.height }
    : { width: settings.paperWidth, height: settings.paperHeight })
  const firstOutputSize = outputSizes[0]
  const mixedPaperSizes = outputSizes.some((size) => (
    Math.abs(size.width - firstOutputSize.width) > 0.01
    || Math.abs(size.height - firstOutputSize.height) > 0.01
  ))
  const placements = []
  for (let index = 0; index < embeddedPages.length; index += 1) {
    const embedded = embeddedPages[index]
    const pageSize = outputSizes[index]
    const pageMargins = settings.paper === 'Document'
      ? normalizeMargins(input?.settings?.margins, pageSize.width, pageSize.height)
      : settings.margins
    const margin = {
      top: pageMargins.top * POINTS_PER_INCH,
      right: pageMargins.right * POINTS_PER_INCH,
      bottom: pageMargins.bottom * POINTS_PER_INCH,
      left: pageMargins.left * POINTS_PER_INCH,
    }
    const availableWidth = Math.max(18, pageSize.width - margin.left - margin.right)
    const availableHeight = Math.max(18, pageSize.height - margin.top - margin.bottom - headerBand - footerBand)
    const page = output.addPage([pageSize.width, pageSize.height])
    const scale = fitScale(embedded.width, embedded.height, availableWidth, availableHeight, settings)
    const drawnWidth = embedded.width * scale
    const drawnHeight = embedded.height * scale
    const contentBottom = margin.bottom + footerBand
    const x = settings.centerContent ? margin.left + (availableWidth - drawnWidth) / 2 : margin.left
    const y = settings.centerContent ? contentBottom + (availableHeight - drawnHeight) / 2 : contentBottom + availableHeight - drawnHeight
    page.drawPage(embedded, { x, y, width: drawnWidth, height: drawnHeight })

    if (settings.printTitle) {
      const size = 8
      const value = trimToWidth(printableTitle, font, size, availableWidth)
      page.drawText(value, {
        x: margin.left,
        y: pageSize.height - margin.top - size,
        size,
        font,
        color: rgb(0.28, 0.28, 0.28),
      })
    }
    if (settings.printPageNumbers) {
      const size = 8
      const value = `Page ${index + 1} of ${embeddedPages.length}`
      const width = font.widthOfTextAtSize(value, size)
      page.drawText(value, {
        x: margin.left + Math.max(0, (availableWidth - width) / 2),
        y: margin.bottom,
        size,
        font,
        color: rgb(0.28, 0.28, 0.28),
      })
    }
    placements.push({
      sourcePage: selected[index] + 1,
      scale: Number(scale.toFixed(4)),
      clipped: drawnWidth > availableWidth + 0.01 || drawnHeight > availableHeight + 0.01,
    })
  }
  output.setTitle(title)
  output.setCreator('Simple Docs')
  output.setProducer('Simple Docs print layout')
  const result = await output.save({ useObjectStreams: true, addDefaultPage: false })
  return {
    data: new Uint8Array(result),
    sourcePageCount,
    outputPageCount: selected.length,
    paperWidth: firstOutputSize.width,
    paperHeight: firstOutputSize.height,
    mixedPaperSizes,
    placements,
    settings,
  }
}

module.exports = {
  MAX_PDF_BYTES,
  MARGIN_PRESETS,
  PAPER_SIZES,
  cleanTitle,
  composePrintPdf,
  nativePrintOptions,
  resolvePrinter,
  normalizePrintSettings,
  parsePageRange,
  safeTextForFont,
  validatePdfBytes,
}
