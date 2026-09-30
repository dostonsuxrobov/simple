'use strict'

const { directPrintOptions } = require('./default-printer.cjs')

const PRINT_PAPERS = Object.freeze({
  letter: Object.freeze({ label: 'Letter', electronName: 'Letter', widthMm: 215.9, heightMm: 279.4 }),
  a4: Object.freeze({ label: 'A4', electronName: 'A4', widthMm: 210, heightMm: 297 }),
  legal: Object.freeze({ label: 'Legal', electronName: 'Legal', widthMm: 215.9, heightMm: 355.6 }),
})

const DEFAULT_PRINT_SETTINGS = Object.freeze({
  paper: 'letter',
  orientation: 'portrait',
  marginMm: 12.7,
  scaleMode: 'fit',
  scalePercent: 100,
  position: 'center',
  background: '#ffffff',
  grayscale: false,
  copies: 1,
})

const SCALE_MODES = new Set(['fit', 'fill', 'actual', 'custom'])
const POSITIONS = new Set(['center', 'top-left', 'top-right', 'bottom-left', 'bottom-right'])

function finiteNumber(value, fallback) {
  const number = Number(value)
  return Number.isFinite(number) ? number : fallback
}

function clamp(value, minimum, maximum) {
  return Math.max(minimum, Math.min(maximum, value))
}

function normalizePrintSettings(input = {}) {
  const paper = Object.hasOwn(PRINT_PAPERS, input.paper) ? input.paper : DEFAULT_PRINT_SETTINGS.paper
  const orientation = input.orientation === 'landscape' ? 'landscape' : 'portrait'
  const paperDefinition = PRINT_PAPERS[paper]
  const shortestEdge = Math.min(paperDefinition.widthMm, paperDefinition.heightMm)
  return {
    paper,
    orientation,
    marginMm: clamp(finiteNumber(input.marginMm, DEFAULT_PRINT_SETTINGS.marginMm), 0, Math.min(50, shortestEdge / 2 - 1)),
    scaleMode: SCALE_MODES.has(input.scaleMode) ? input.scaleMode : DEFAULT_PRINT_SETTINGS.scaleMode,
    scalePercent: clamp(finiteNumber(input.scalePercent, DEFAULT_PRINT_SETTINGS.scalePercent), 10, 400),
    position: POSITIONS.has(input.position) ? input.position : DEFAULT_PRINT_SETTINGS.position,
    background: /^#[0-9a-f]{6}$/i.test(String(input.background || '')) ? String(input.background).toLowerCase() : DEFAULT_PRINT_SETTINGS.background,
    grayscale: input.grayscale === true,
    copies: Math.trunc(clamp(finiteNumber(input.copies, DEFAULT_PRINT_SETTINGS.copies), 1, 99)),
  }
}

function orientedPaper(settingsInput = {}) {
  const settings = normalizePrintSettings(settingsInput)
  const definition = PRINT_PAPERS[settings.paper]
  return settings.orientation === 'landscape'
    ? { ...definition, widthMm: definition.heightMm, heightMm: definition.widthMm }
    : { ...definition }
}

function alignedOffset(available, occupied, position, leading, trailing) {
  if (position === leading) return 0
  if (position === trailing) return available - occupied
  return (available - occupied) / 2
}

function computePrintLayout(imageWidth, imageHeight, settingsInput = {}) {
  const widthPx = finiteNumber(imageWidth, 0)
  const heightPx = finiteNumber(imageHeight, 0)
  if (widthPx <= 0 || heightPx <= 0) throw new Error('The image dimensions are invalid for printing.')

  const settings = normalizePrintSettings(settingsInput)
  const paper = orientedPaper(settings)
  const printableWidthMm = Math.max(0.1, paper.widthMm - settings.marginMm * 2)
  const printableHeightMm = Math.max(0.1, paper.heightMm - settings.marginMm * 2)
  let scaleMmPerPixel

  if (settings.scaleMode === 'fit') {
    scaleMmPerPixel = Math.min(printableWidthMm / widthPx, printableHeightMm / heightPx)
  } else if (settings.scaleMode === 'fill') {
    scaleMmPerPixel = Math.max(printableWidthMm / widthPx, printableHeightMm / heightPx)
  } else {
    const actualScale = 25.4 / 96
    scaleMmPerPixel = actualScale * (settings.scaleMode === 'custom' ? settings.scalePercent / 100 : 1)
  }

  const imageWidthMm = widthPx * scaleMmPerPixel
  const imageHeightMm = heightPx * scaleMmPerPixel
  const xMm = alignedOffset(printableWidthMm, imageWidthMm, settings.position, 'top-left', 'top-right')
  const yMm = alignedOffset(printableHeightMm, imageHeightMm, settings.position, 'top-left', 'bottom-left')
  const resolvedXMm = settings.position === 'bottom-right'
    ? printableWidthMm - imageWidthMm
    : settings.position === 'bottom-left'
      ? 0
      : xMm
  const resolvedYMm = settings.position === 'top-right'
    ? 0
    : settings.position === 'bottom-right'
      ? printableHeightMm - imageHeightMm
      : yMm

  return {
    settings,
    paper,
    printable: {
      xMm: settings.marginMm,
      yMm: settings.marginMm,
      widthMm: printableWidthMm,
      heightMm: printableHeightMm,
    },
    image: {
      xMm: resolvedXMm,
      yMm: resolvedYMm,
      widthMm: imageWidthMm,
      heightMm: imageHeightMm,
    },
    clipped: imageWidthMm > printableWidthMm + 0.001 || imageHeightMm > printableHeightMm + 0.001,
    effectiveDpi: 25.4 / scaleMmPerPixel,
  }
}

function electronPrintOptions(settingsInput = {}) {
  const settings = normalizePrintSettings(settingsInput)
  const paper = orientedPaper(settings)
  return directPrintOptions({
    printBackground: true,
    color: !settings.grayscale,
    landscape: settings.orientation === 'landscape',
    scaleFactor: 100,
    copies: settings.copies,
    collate: true,
    margins: { marginType: 'none' },
    pageSize: paper.electronName,
  })
}

function cssMillimetres(value) {
  return `${Number(value.toFixed(4))}mm`
}

function buildPrintHtml(imageSource, imageWidth, imageHeight, settingsInput = {}, title = 'Image') {
  const source = String(imageSource || '')
  if (!/^data:image\/png;base64,[a-z0-9+/=]+$/i.test(source) && !/^file:\/{2,3}[^"<>]+$/i.test(source)) {
    throw new Error('The printable image data is invalid.')
  }
  const layout = computePrintLayout(imageWidth, imageHeight, settingsInput)
  const grayscale = layout.settings.grayscale ? 'grayscale(1)' : 'none'
  const safeTitle = String(title || 'Image').replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  })[character])
  const safeSource = source.replace(/&/g, '&amp;').replace(/"/g, '&quot;')
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>${safeTitle}</title><style>
  @page { size: ${layout.paper.electronName} ${layout.settings.orientation}; margin: 0; }
  * { box-sizing: border-box; }
  html, body { width: ${cssMillimetres(layout.paper.widthMm)}; height: ${cssMillimetres(layout.paper.heightMm)}; margin: 0; padding: 0; overflow: hidden; background: #fff; }
  .page { position: relative; width: 100%; height: 100%; overflow: hidden; background: #fff; break-after: avoid; page-break-after: avoid; }
  .printable { position: absolute; left: ${cssMillimetres(layout.printable.xMm)}; top: ${cssMillimetres(layout.printable.yMm)}; width: ${cssMillimetres(layout.printable.widthMm)}; height: ${cssMillimetres(layout.printable.heightMm)}; overflow: hidden; }
  .image-frame { position: absolute; left: ${cssMillimetres(layout.image.xMm)}; top: ${cssMillimetres(layout.image.yMm)}; width: ${cssMillimetres(layout.image.widthMm)}; height: ${cssMillimetres(layout.image.heightMm)}; overflow: hidden; background: ${layout.settings.background}; filter: ${grayscale}; }
  img { display: block; width: 100%; height: 100%; object-fit: fill; }
</style></head><body><main class="page"><div class="printable"><div class="image-frame"><img src="${safeSource}" width="${Math.trunc(imageWidth)}" height="${Math.trunc(imageHeight)}" alt=""></div></div></main></body></html>`
}

module.exports = {
  PRINT_PAPERS,
  DEFAULT_PRINT_SETTINGS,
  normalizePrintSettings,
  orientedPaper,
  computePrintLayout,
  electronPrintOptions,
  buildPrintHtml,
}
