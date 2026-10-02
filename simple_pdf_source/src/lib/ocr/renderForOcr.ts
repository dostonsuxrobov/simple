import type { PDFPageProxy } from 'pdfjs-dist'
import { boundedCanvasSize } from '../../../electron/canvas-size.mjs'
import { pdfjs } from '../pdf'
import { OcrError, ocrAbortError, type OcrRotation, type PdfPoint } from './types'

export const OCR_DEFAULT_DPI = 300
/** Large-format pages lower their resolution to stay under this many pixels. */
export const OCR_MAX_PIXELS = 36_000_000

type RenderParameters = Parameters<PDFPageProxy['render']>[0]

export interface OcrRenderOptions {
  /** The user's pending view rotation for this page (added to /Rotate). */
  pageRotation?: number
  /** Resolution of the page's scan image, when known. */
  nativeDpi?: number
  /** Explicit resolution (e.g. a quick orientation probe); overrides nativeDpi. */
  dpi?: number
  maxPixels?: number
  /**
   * 'display' (default while the window is visible) yields to the UI between
   * chunks via requestAnimationFrame, which stalls in a hidden or minimised
   * window; 'print' (default when hidden) does not depend on frames.
   */
  intent?: 'display' | 'print'
  /** With print intent, pass the document's display configuration to keep on-screen layer visibility. */
  optionalContentConfigPromise?: RenderParameters['optionalContentConfigPromise']
  /** Release the page's decoded images afterwards (default true): batches would otherwise hold every scan. */
  cleanup?: boolean
  signal?: AbortSignal
}

export interface OcrRender {
  /** RGBA pixels (ImageData); transfer `rgba.buffer` to the prep worker. */
  rgba: Uint8ClampedArray
  width: number
  height: number
  dpi: number
  /** /Rotate plus the pending view rotation. */
  rotation: OcrRotation
  intent: 'display' | 'print'
  /** OCR pixel (continuous, top-left origin) -> unrotated PDF user space. */
  toPdf(x: number, y: number): PdfPoint
  /** The same mapping as an affine matrix [a, b, c, d, e, f]: (x, y) -> (a x + c y + e, b x + d y + f). */
  pdfFromPixel: [number, number, number, number, number, number]
}

/** Scans between 280 and 400 DPI are recognised at their own resolution, everything else at 300. */
export function chooseOcrDpi(nativeDpi?: number): number {
  return typeof nativeDpi === 'number' && Number.isFinite(nativeDpi) && nativeDpi >= 280 && nativeDpi <= 400
    ? nativeDpi
    : OCR_DEFAULT_DPI
}

export function normalizeRotation(degrees: number): OcrRotation {
  const quarter = Math.round((Number.isFinite(degrees) ? degrees : 0) / 90)
  return ((((quarter % 4) + 4) % 4) * 90) as OcrRotation
}

/**
 * Render a page for recognition the way the user sees it (including a pending
 * view rotation) but without annotations, form widgets or Simple's overlays,
 * which live outside the page content.
 */
export async function renderForOcr(page: PDFPageProxy, options: OcrRenderOptions = {}): Promise<OcrRender> {
  if (options.signal?.aborted) throw ocrAbortError()
  const rotation = normalizeRotation((page.rotate || 0) + (options.pageRotation ?? 0))
  const requestedDpi = options.dpi ?? chooseOcrDpi(options.nativeDpi)
  const base = page.getViewport({ scale: 1, rotation })
  const size = boundedCanvasSize(base.width, base.height, requestedDpi / 72, options.maxPixels ?? OCR_MAX_PIXELS)
  const scale = size.width / base.width
  const viewport = page.getViewport({ scale, rotation })
  const { width, height } = size
  const intent = options.intent ?? (typeof document !== 'undefined' && document.visibilityState === 'hidden' ? 'print' : 'display')

  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  // Software canvas: one read-back of up to 36 MP without a GPU copy or a GPU-sized allocation.
  const context = canvas.getContext('2d', { alpha: false, willReadFrequently: true })
  if (!context) throw new OcrError('render-failed', 'This page could not be prepared for text recognition.', 'No 2D canvas context')
  const task = page.render({
    canvasContext: context,
    viewport,
    intent,
    annotationMode: pdfjs.AnnotationMode.DISABLE,
    background: '#ffffff',
    transform: [width / viewport.width, 0, 0, height / viewport.height, 0, 0],
    ...(options.optionalContentConfigPromise ? { optionalContentConfigPromise: options.optionalContentConfigPromise } : {}),
  })
  const onAbort = () => task.cancel()
  options.signal?.addEventListener('abort', onAbort, { once: true })
  let rgba: Uint8ClampedArray
  try {
    await task.promise
    if (options.signal?.aborted) throw ocrAbortError()
    rgba = context.getImageData(0, 0, width, height).data
  } catch (error) {
    if (options.signal?.aborted) throw ocrAbortError()
    throw new OcrError('render-failed', 'This page could not be prepared for text recognition.', String(error))
  } finally {
    options.signal?.removeEventListener('abort', onAbort)
    canvas.width = 0
    canvas.height = 0
    if (options.cleanup !== false) page.cleanup()
  }

  const sx = viewport.width / width
  const sy = viewport.height / height
  const toPdf = (x: number, y: number): PdfPoint => {
    const [px, py] = viewport.convertToPdfPoint(x * sx, y * sy)
    return { x: px, y: py }
  }
  const o = toPdf(0, 0)
  const ex = toPdf(1, 0)
  const ey = toPdf(0, 1)
  return {
    rgba,
    width,
    height,
    dpi: 72 * scale,
    rotation,
    intent,
    toPdf,
    pdfFromPixel: [ex.x - o.x, ex.y - o.y, ey.x - o.x, ey.y - o.y, o.x, o.y],
  }
}
