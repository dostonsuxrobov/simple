import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { ExportTextPage } from '../types'

type ExportImageFormat = 'png' | 'jpeg' | 'webp'

interface TextContentItem {
  str: string
  hasEOL?: boolean
}

function isTextItem(value: unknown): value is TextContentItem {
  return Boolean(value && typeof value === 'object' && typeof (value as TextContentItem).str === 'string')
}

export function textContentToPlainText(items: unknown[]) {
  let output = ''
  for (const rawItem of items) {
    if (!isTextItem(rawItem) || !rawItem.str) continue
    const value = rawItem.str.replace(/\u0000/g, '')
    const previous = output.at(-1) || ''
    if (output && !/\s/.test(previous) && !/^\s|^[,.;:!?%)\]}]/.test(value)) output += ' '
    output += value
    if (rawItem.hasEOL) output = `${output.trimEnd()}\n`
  }
  return output
    .split('\n')
    .map((line) => line.trimEnd())
    .join('\n')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim()
}

export async function extractPdfText(pdf: PDFDocumentProxy, pageIndices: number[]): Promise<ExportTextPage[]> {
  const pages: ExportTextPage[] = []
  for (const pageIndex of pageIndices) {
    const page = await pdf.getPage(pageIndex + 1)
    try {
      const content = await page.getTextContent()
      pages.push({ pageNumber: pageIndex + 1, text: textContentToPlainText(content.items) })
    } finally {
      page.cleanup()
    }
  }
  return pages
}

function canvasBlob(canvas: HTMLCanvasElement, mimeType: string, quality?: number) {
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error('The rendered page could not be encoded.')), mimeType, quality)
  })
}

export async function renderPdfPageImage(
  pdf: PDFDocumentProxy,
  pageIndex: number,
  format: ExportImageFormat,
  requestedScale: number,
  jpegQuality: number,
) {
  const page = await pdf.getPage(pageIndex + 1)
  let canvas: HTMLCanvasElement | null = null
  try {
    const unitViewport = page.getViewport({ scale: 1, rotation: page.rotate || 0 })
    const pixelLimitScale = Math.sqrt(64_000_000 / Math.max(1, unitViewport.width * unitViewport.height))
    const dimensionLimitScale = 8_192 / Math.max(1, unitViewport.width, unitViewport.height)
    // Apply safety caps last: a minimum of 0.1 *after* the cap still creates
    // enormous canvases for large-format PDFs and makes every image codec fail.
    const scale = Math.min(Math.max(0.01, Number(requestedScale) || 1), pixelLimitScale, dimensionLimitScale)
    const viewport = page.getViewport({ scale, rotation: page.rotate || 0 })
    canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.floor(viewport.width))
    canvas.height = Math.max(1, Math.floor(viewport.height))
    const context = canvas.getContext('2d', { alpha: format !== 'jpeg' })
    if (!context) throw new Error('The page image canvas could not be created.')
    if (format === 'jpeg') {
      context.fillStyle = '#ffffff'
      context.fillRect(0, 0, canvas.width, canvas.height)
    }
    await page.render({ canvasContext: context, viewport,
      transform: [canvas.width / viewport.width, 0, 0, canvas.height / viewport.height, 0, 0],
    }).promise
    const mimeType = format === 'jpeg' ? 'image/jpeg' : format === 'webp' ? 'image/webp' : 'image/png'
    const blob = await canvasBlob(canvas, mimeType, format === 'jpeg' || format === 'webp' ? jpegQuality : undefined)
    return new Uint8Array(await blob.arrayBuffer())
  } finally {
    if (canvas) {
      canvas.width = 0
      canvas.height = 0
    }
    page.cleanup()
  }
}
