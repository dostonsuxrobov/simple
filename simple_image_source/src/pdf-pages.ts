// PDF pages as pictures, for the image window: a PDF dropped on the Advanced editor
// becomes layers, and a PDF met while stepping through a folder (or dropped on the
// viewer) is shown as its first page. Loaded only when a PDF is actually met, so the
// viewer does not carry the PDF engine around for ordinary images.
import * as pdfjs from 'pdfjs-dist'
import PdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url'

pdfjs.GlobalWorkerOptions.workerSrc = PdfWorker
// A real worker, made here: left to itself pdf.js wraps the worker in a blob: URL for a page
// loaded from a file, which this window's content policy refuses, and then quietly does all
// its parsing and drawing on the window's own thread (a large PDF froze the window).
try { pdfjs.GlobalWorkerOptions.workerPort = new Worker(PdfWorker, { type: 'module' }) } catch { /* pdf.js falls back to the window's thread */ }

// Character maps and the standard fonts, for PDFs that do not carry their own (much Chinese,
// Japanese and Korean text; the 14 standard fonts). The build copies them beside the page
// (vite.config.ts); without them such text is simply not drawn.
const PDFJS_DATA = new URL('./pdfjs/', document.baseURI).href

/** Pixels per PDF point: 150 DPI, enough to read and to edit as an image. */
const SCALE = 150 / 72
/** The same limits the image editor has for any picture. */
const MAX_SIDE = 20_000
const MAX_PIXELS = 50_000_000

export interface RenderedPdfPage {
  /** 1-based page number in the PDF. */
  readonly page: number
  readonly width: number
  readonly height: number
  /** Straight-alpha RGBA pixels, on white paper (a PDF page has no transparency of its own). */
  readonly data: Uint8ClampedArray
}

export interface RenderedPdf {
  readonly pageCount: number
  readonly pages: readonly RenderedPdfPage[]
}

/**
 * Renders the first `maxPages` pages of a PDF, one at a time. With `onPage` each page is
 * handed over as soon as it is drawn and not kept (ten large pages are never in memory
 * together); returning false from it stops the run.
 * @throws Error with a plain message for a locked or unreadable PDF
 */
export async function renderPdfPages(bytes: Uint8Array, maxPages: number, onPage?: (page: RenderedPdfPage, pageCount: number) => boolean | void): Promise<RenderedPdf> {
  let pdf: Awaited<ReturnType<typeof pdfjs.getDocument>['promise']>
  // Dynamic evaluation stays off for untrusted files, as in the PDF window.
  const task = pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false, cMapUrl: `${PDFJS_DATA}cmaps/`, cMapPacked: true, standardFontDataUrl: `${PDFJS_DATA}standard_fonts/` })
  try {
    pdf = await task.promise
  } catch (error) {
    await task.destroy().catch(() => {})
    const name = error instanceof Error ? error.name : ''
    if (name === 'PasswordException') throw new Error('This PDF is locked with a password. Open it in the PDF window instead.')
    throw new Error('This PDF could not be read.')
  }
  try {
    const pages: RenderedPdfPage[] = []
    const count = Math.min(pdf.numPages, Math.max(1, maxPages))
    for (let number = 1; number <= count; number += 1) {
      const page = await pdf.getPage(number)
      const base = page.getViewport({ scale: 1 })
      // A huge page is drawn smaller rather than refused.
      const fit = Math.min(SCALE, MAX_SIDE / base.width, MAX_SIDE / base.height, Math.sqrt(MAX_PIXELS / (base.width * base.height)))
      const viewport = page.getViewport({ scale: fit })
      const width = Math.max(1, Math.floor(viewport.width))
      const height = Math.max(1, Math.floor(viewport.height))
      const canvas = document.createElement('canvas')
      canvas.width = width
      canvas.height = height
      const context = canvas.getContext('2d', { willReadFrequently: true })
      if (!context) throw new Error('This PDF page could not be drawn.')
      try {
        context.fillStyle = '#ffffff'
        context.fillRect(0, 0, width, height)
        await page.render({ canvasContext: context, viewport }).promise
        const rendered: RenderedPdfPage = { page: number, width, height, data: context.getImageData(0, 0, width, height).data }
        if (onPage) { if (onPage(rendered, pdf.numPages) === false) break }
        else pages.push(rendered)
      } finally {
        canvas.width = 1
        canvas.height = 1
        page.cleanup()
      }
    }
    return { pageCount: pdf.numPages, pages }
  } finally {
    await pdf.destroy()
  }
}

/** PNG bytes of one rendered page (what the viewer opens). */
export async function pageToPng(page: RenderedPdfPage): Promise<Uint8Array> {
  const canvas = new OffscreenCanvas(page.width, page.height)
  try {
    const context = canvas.getContext('2d')
    if (!context) throw new Error('This PDF page could not be drawn.')
    context.putImageData(new ImageData(new Uint8ClampedArray(page.data), page.width, page.height), 0, 0)
    return new Uint8Array(await (await canvas.convertToBlob({ type: 'image/png' })).arrayBuffer())
  } finally {
    canvas.width = 1
    canvas.height = 1
  }
}
