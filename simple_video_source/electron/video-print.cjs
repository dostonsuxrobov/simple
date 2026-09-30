'use strict'

const path = require('node:path')
const { validateFrameBytes } = require('./video-export.cjs')

const PAPER_SIZES = Object.freeze({
  letter: Object.freeze({ css: 'Letter', label: 'Letter', width: 8.5, height: 11 }),
  a4: Object.freeze({ css: 'A4', label: 'A4', width: 210 / 25.4, height: 297 / 25.4 }),
  legal: Object.freeze({ css: 'Legal', label: 'Legal', width: 8.5, height: 14 }),
})

const MARGINS = Object.freeze({
  normal: Object.freeze({ label: 'Normal', inches: 0.5 }),
  narrow: Object.freeze({ label: 'Narrow', inches: 0.25 }),
  wide: Object.freeze({ label: 'Wide', inches: 0.75 }),
})

const MAX_FRAME_DIMENSION = 16_384
const MAX_FRAME_PIXELS = 100_000_000

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function pngDimensions(bytes) {
  const validated = validateFrameBytes(bytes, 'png')
  if (validated.length < 24 || validated.subarray(12, 16).toString('ascii') !== 'IHDR') {
    throw new Error('The captured print frame does not contain a valid PNG header.')
  }
  const width = validated.readUInt32BE(16)
  const height = validated.readUInt32BE(20)
  if (!width || !height || width > MAX_FRAME_DIMENSION || height > MAX_FRAME_DIMENSION || width * height > MAX_FRAME_PIXELS) {
    throw new Error('The captured print frame dimensions are too large or invalid.')
  }
  return { bytes: validated, width, height }
}

function safePrintOptions(value) {
  const input = value && typeof value === 'object' ? value : {}
  const customScale = Math.max(10, Math.min(400, Number(input.customScale) || 100))
  return {
    paperSize: Object.hasOwn(PAPER_SIZES, input.paperSize) ? input.paperSize : 'letter',
    orientation: input.orientation === 'landscape' ? 'landscape' : 'portrait',
    margins: Object.hasOwn(MARGINS, input.margins) ? input.margins : 'normal',
    scaleMode: ['fit', 'fill', 'actual', 'custom'].includes(input.scaleMode) ? input.scaleMode : 'fit',
    customScale,
    colorMode: input.colorMode === 'grayscale' ? 'grayscale' : 'color',
    metadata: input.metadata === true,
  }
}

function formatTimestamp(seconds) {
  const total = Math.max(0, Math.floor(Number.isFinite(seconds) ? seconds : 0))
  const hours = Math.floor(total / 3600)
  const minutes = Math.floor((total % 3600) / 60)
  const remainder = total % 60
  return hours
    ? `${hours}:${String(minutes).padStart(2, '0')}:${String(remainder).padStart(2, '0')}`
    : `${minutes}:${String(remainder).padStart(2, '0')}`
}

function safeSourceName(value) {
  const base = path.basename(String(value || 'Video frame')).replace(/[\u0000-\u001f\u007f]/g, ' ').trim()
  return (base || 'Video frame').slice(0, 240)
}

function createVideoPrintDocument(input) {
  if (!input || typeof input !== 'object') throw new Error('The video print request is invalid.')
  const frame = pngDimensions(input.bytes)
  const options = safePrintOptions(input.options)
  const paper = PAPER_SIZES[options.paperSize]
  const margin = MARGINS[options.margins]
  const widthInches = options.orientation === 'landscape' ? paper.height : paper.width
  const heightInches = options.orientation === 'landscape' ? paper.width : paper.height
  const pageWidth = widthInches * 96
  const pageHeight = heightInches * 96
  const contentWidth = Math.max(1, pageWidth - margin.inches * 192)
  const contentHeight = Math.max(1, pageHeight - margin.inches * 192)
  const metadataHeight = options.metadata ? 46 : 0
  const metadataGap = options.metadata ? 10 : 0
  const frameAreaHeight = Math.max(1, contentHeight - metadataHeight - metadataGap)
  const fitScale = Math.min(contentWidth / frame.width, frameAreaHeight / frame.height)
  const fillScale = Math.max(contentWidth / frame.width, frameAreaHeight / frame.height)
  const scale = options.scaleMode === 'fill'
    ? fillScale
    : options.scaleMode === 'actual'
      ? 1
      : options.scaleMode === 'custom'
        ? options.customScale / 100
        : fitScale
  const renderedWidth = frame.width * scale
  const renderedHeight = frame.height * scale
  const offsetX = (contentWidth - renderedWidth) / 2
  const offsetY = (frameAreaHeight - renderedHeight) / 2
  const cropped = renderedWidth > contentWidth + 0.01 || renderedHeight > frameAreaHeight + 0.01
  const sourceName = safeSourceName(input.sourceName)
  const timestamp = formatTimestamp(Number(input.seconds))
  const dataUrl = `data:image/png;base64,${frame.bytes.toString('base64')}`
  const filter = options.colorMode === 'grayscale' ? 'filter:grayscale(1);' : ''
  const metadata = options.metadata
    ? `<footer class="frame-metadata"><strong>${escapeHtml(sourceName)}</strong><span>Frame at ${timestamp} · ${frame.width} × ${frame.height}px</span></footer>`
    : ''
  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(sourceName)} — frame ${timestamp}</title>
<style>
@page { size: ${paper.css} ${options.orientation}; margin: 0; }
* { box-sizing: border-box; }
:root { --preview-zoom:1; }
html, body { margin: 0; min-height: 100%; padding: 0; font-family: Arial, sans-serif; }
.print-page { width:${pageWidth.toFixed(2)}px; height:${pageHeight.toFixed(2)}px; padding:${(margin.inches * 96).toFixed(2)}px; overflow:hidden; background:#fff; break-after:page; page-break-after:always; }
.page-content { width:${contentWidth.toFixed(2)}px; height:${contentHeight.toFixed(2)}px; overflow:hidden; }
.frame-slot { position:relative; width:${contentWidth.toFixed(2)}px; height:${frameAreaHeight.toFixed(2)}px; overflow:hidden; background:#fff; }
.frame-slot img { position:absolute; left:${offsetX.toFixed(3)}px; top:${offsetY.toFixed(3)}px; width:${renderedWidth.toFixed(3)}px; height:${renderedHeight.toFixed(3)}px; display:block; ${filter} }
.frame-metadata { height:${metadataHeight}px; display:flex; align-items:center; justify-content:space-between; gap:18px; margin-top:${metadataGap}px; padding-top:8px; border-top:1px solid #c9c9c9; color:#222; font-size:9px; line-height:1.2; }
.frame-metadata strong, .frame-metadata span { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.frame-metadata strong { max-width:65%; font-weight:650; }
.frame-metadata span { color:#555; }
@media screen {
  html { background:#d8dbd9; }
  body { width:max-content; min-width:100%; display:flex; justify-content:center; padding:24px; background:#d8dbd9; zoom:var(--preview-zoom); }
  .print-page { flex:none; box-shadow:0 3px 18px rgba(0,0,0,.24); }
}
@media print {
  html, body { background:#fff; }
  body { display:block; }
  .print-page { margin:0; box-shadow:none; }
}
</style></head><body data-paper-size="${options.paperSize}" data-orientation="${options.orientation}" data-margins="${options.margins}" data-scale-mode="${options.scaleMode}" data-color-mode="${options.colorMode}" data-metadata="${options.metadata}"><section class="print-page" aria-label="Printed video frame"><main class="page-content"><div class="frame-slot"><img src="${dataUrl}" alt="Video frame at ${timestamp}"></div>${metadata}</main></section></body></html>`

  return {
    html,
    title: `${sourceName} — frame ${timestamp}`,
    options,
    page: { label: paper.label, widthInches, heightInches, marginInches: margin.inches },
    frame: { width: frame.width, height: frame.height, timestamp },
    placement: {
      scale,
      renderedWidth,
      renderedHeight,
      availableWidth: contentWidth,
      availableHeight: frameAreaHeight,
      cropped,
    },
  }
}

function assertPrintersInstalled(printers) {
  if (!Array.isArray(printers) || printers.length === 0) {
    throw new Error('No printers are installed. Add a printer in Windows Settings, set it as the default, and try again.')
  }
  return printers.length
}

async function ensurePrinterAvailable(webContents) {
  if (!webContents || typeof webContents.getPrintersAsync !== 'function') {
    throw new Error('Simple cannot read the printers available on this computer.')
  }
  let printers
  try {
    printers = await webContents.getPrintersAsync()
  } catch {
    throw new Error('Simple could not read the installed printers. Check Windows printer settings and try again.')
  }
  return assertPrintersInstalled(printers)
}

function nativePrintOptions(printDocument) {
  const paperSize = printDocument.options.paperSize === 'a4' ? 'A4' : printDocument.options.paperSize === 'legal' ? 'Legal' : 'Letter'
  return {
    silent: true,
    printBackground: true,
    color: printDocument.options.colorMode === 'color',
    landscape: printDocument.options.orientation === 'landscape',
    margins: { marginType: 'none' },
    pageSize: paperSize,
    scaleFactor: 100,
    pagesPerSheet: 1,
    collate: true,
  }
}

function runSilentPrintJob(webContents, printDocument, { timeoutMs = 120_000 } = {}) {
  if (!webContents || typeof webContents.print !== 'function') {
    return Promise.reject(new Error('Direct printing is unavailable in this window.'))
  }
  const options = nativePrintOptions(printDocument)
  return new Promise((resolve, reject) => {
    let settled = false
    let timer
    const finish = (error) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      webContents.removeListener?.('destroyed', onStopped)
      webContents.removeListener?.('render-process-gone', onStopped)
      if (error) reject(error)
      else resolve({ printed: true, canceled: false })
    }
    const onStopped = () => finish(new Error('The print renderer stopped before Windows confirmed the job. Check the print queue before trying again.'))
    if (webContents.isDestroyed?.()) { onStopped(); return }
    webContents.once?.('destroyed', onStopped)
    webContents.once?.('render-process-gone', onStopped)
    timer = setTimeout(() => finish(new Error('Windows did not confirm the print job within two minutes. Check the print queue before trying again to avoid a duplicate copy.')), Math.max(1, timeoutMs))
    try { webContents.print(options, (success, failureReason) => {
      if (success) {
        finish()
        return
      }
      const reason = String(failureReason || '').trim()
      finish(new Error(reason
        ? `The default printer did not accept the print job: ${reason}`
        : 'The default printer did not accept the print job. Make sure a default printer is configured and online, then try again.'))
    }) } catch (error) { finish(error) }
  })
}

module.exports = {
  MARGINS,
  PAPER_SIZES,
  assertPrintersInstalled,
  createVideoPrintDocument,
  ensurePrinterAvailable,
  formatTimestamp,
  nativePrintOptions,
  pngDimensions,
  runSilentPrintJob,
  safePrintOptions,
}
