import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist'
import { textPaintResolver } from '../../../electron/text-appearance.mjs'
import { getPageTextContent, pdfjs } from '../pdf'
import type { PdfRect } from '../../types'
import { ocrAbortError } from './types'

// Which pages are scans, and which already carry text: drives "Recognize
// text" (which pages need it), the Edit-mode offer, the inspector, the search
// hint and the export check. Text painted invisibly (render mode 3 or 7, an
// OCR layer) is told apart from visible text, so a page recognised earlier
// (by Simple or another tool) is not mistaken for a born-digital page.

export type PageKind = 'native' | 'scan' | 'searchable-scan' | 'mixed' | 'vector-only' | 'blank'

export interface PageScanState {
  kind: PageKind
  /** Non-whitespace characters of visibly painted text. */
  visibleChars: number
  /** Non-whitespace characters of invisible text (an OCR layer). */
  invisibleChars: number
  /** Share of the page (its crop box) covered by images, 0..1. */
  imageCoverage: number
  /** Resolution of the largest image covering at least 30% of the page. */
  nativeDpi?: number
  /** Path construction and painting operators (outlined text has thousands). */
  vectorOps: number
  /** A scanned page without text: recognising it is offered. */
  needsOcr: boolean
  /** The page already has recognised (invisible) text. */
  hasOcrText: boolean
  /** That text uses GlyphLessFont, as Simple's (and Tesseract-based tools') layers do. */
  ocrBySimple: boolean
  /** Bounds of the visible text runs (PDF space): recognised words over them are duplicates. */
  visibleTextRects: PdfRect[]
  /** Bounds of the images, clipped to the page (PDF space). */
  imageRects: PdfRect[]
}

/** Tunable thresholds of the rules in pageKindFromStats. */
export const PAGE_CLASSIFIER = Object.freeze({
  /** Fewer characters than this count as "no text". */
  minChars: 20,
  /** Image coverage from which a page is a scan (or a mixed page). */
  scanCoverage: 0.25,
  /** An image must cover this much of the page to give the page's scan resolution. */
  dpiCoverage: 0.3,
  /** Outlined text: at least this many path operators and almost no text. */
  outlinedTextOps: 200,
  /** Fewer path operators than this on a page without text or images: blank. */
  blankVectorOps: 5,
  /** Images are rasterised on a grid this many cells wide and high to measure coverage. */
  grid: 64,
  /** Images covering this much of a scanned page are the scan itself. */
  pageCoveringImage: 0.85,
  /** Rects kept per page (very dense pages keep the first ones). */
  maxRects: 5000,
})

const GLYPHLESS_FONT = /(^|\+)GlyphLessFont$/i

export interface PageStats {
  visibleChars: number
  invisibleChars: number
  imageCoverage: number
  hasImages: boolean
  vectorOps: number
}

/** The page kind for these counts (design rules; order matters). */
export function pageKindFromStats(stats: PageStats): PageKind {
  const { minChars, scanCoverage, outlinedTextOps, blankVectorOps } = PAGE_CLASSIFIER
  const visible = stats.visibleChars
  const invisible = stats.invisibleChars
  if (!stats.hasImages && visible === 0 && invisible === 0 && stats.vectorOps < blankVectorOps) return 'blank'
  if (visible < minChars && invisible < minChars && stats.imageCoverage >= scanCoverage) return 'scan'
  if (invisible >= minChars && visible < minChars) return 'searchable-scan'
  if (visible >= minChars && stats.imageCoverage >= scanCoverage) return 'mixed'
  if (visible < minChars && stats.imageCoverage < scanCoverage && stats.vectorOps >= outlinedTextOps) return 'vector-only'
  return 'native'
}

export function scanStateFromStats(stats: PageStats, extra: Partial<Pick<PageScanState, 'nativeDpi' | 'ocrBySimple' | 'visibleTextRects' | 'imageRects'>> = {}): PageScanState {
  const kind = pageKindFromStats(stats)
  const hasOcrText = stats.invisibleChars >= PAGE_CLASSIFIER.minChars
  return {
    kind,
    visibleChars: stats.visibleChars,
    invisibleChars: stats.invisibleChars,
    imageCoverage: stats.imageCoverage,
    ...(extra.nativeDpi ? { nativeDpi: extra.nativeDpi } : {}),
    vectorOps: stats.vectorOps,
    needsOcr: kind === 'scan',
    hasOcrText,
    ocrBySimple: hasOcrText && Boolean(extra.ocrBySimple),
    visibleTextRects: extra.visibleTextRects ?? [],
    imageRects: extra.imageRects ?? [],
  }
}

type Matrix = [number, number, number, number, number, number]
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]

function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ]
}

function asMatrix(value: unknown): Matrix | null {
  if (!value || typeof value !== 'object' || typeof (value as ArrayLike<number>).length !== 'number') return null
  const source = value as ArrayLike<number>
  if (source.length < 6) return null
  const matrix = Array.from({ length: 6 }, (_, index) => Number(source[index])) as Matrix
  return matrix.every(Number.isFinite) ? matrix : null
}

/** The page's visible area as a rect (pdf.js page.view: the crop box). */
export function pageViewRect(view: ArrayLike<number>): PdfRect {
  const x0 = Math.min(Number(view[0]), Number(view[2]))
  const y0 = Math.min(Number(view[1]), Number(view[3]))
  return { x: x0, y: y0, width: Math.abs(Number(view[2]) - Number(view[0])), height: Math.abs(Number(view[3]) - Number(view[1])) }
}

function boundsOf(points: Array<[number, number]>): PdfRect {
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const [x, y] of points) {
    if (x < minX) minX = x
    if (y < minY) minY = y
    if (x > maxX) maxX = x
    if (y > maxY) maxY = y
  }
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY }
}

/** Where the unit square lands under `matrix` (an image's placement). */
export function unitSquareBounds(matrix: Matrix): PdfRect {
  const [a, b, c, d, e, f] = matrix
  return boundsOf([[e, f], [a + e, b + f], [c + e, d + f], [a + c + e, b + d + f]])
}

export function intersectRects(a: PdfRect, b: PdfRect): PdfRect | null {
  const x0 = Math.max(a.x, b.x)
  const y0 = Math.max(a.y, b.y)
  const x1 = Math.min(a.x + a.width, b.x + b.width)
  const y1 = Math.min(a.y + a.height, b.y + b.height)
  return x1 > x0 && y1 > y0 ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null
}

/** Share of `view` covered by the union of `rects`, measured on a grid. */
export function imageCoverage(rects: PdfRect[], view: PdfRect, grid: number = PAGE_CLASSIFIER.grid): number {
  if (!rects.length || view.width <= 0 || view.height <= 0) return 0
  const cells = new Uint8Array(grid * grid)
  const cellWidth = view.width / grid
  const cellHeight = view.height / grid
  let covered = 0
  for (const rect of rects) {
    // Cells whose centre lies inside the rect.
    const first = Math.max(0, Math.ceil((rect.x - view.x) / cellWidth - 0.5))
    const last = Math.min(grid - 1, Math.floor((rect.x + rect.width - view.x) / cellWidth - 0.5))
    const bottom = Math.max(0, Math.ceil((rect.y - view.y) / cellHeight - 0.5))
    const top = Math.min(grid - 1, Math.floor((rect.y + rect.height - view.y) / cellHeight - 0.5))
    for (let row = bottom; row <= top; row += 1) {
      for (let column = first; column <= last; column += 1) {
        const index = row * grid + column
        if (cells[index]) continue
        cells[index] = 1
        covered += 1
      }
    }
    if (covered === cells.length) break
  }
  return covered / cells.length
}

interface TextItemLike {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  fontName?: string
}

/** The bounds of a pdf.js text item in PDF space, from its baseline origin, direction, advance and font size. */
export function textItemRect(item: TextItemLike, style?: { ascent?: number; descent?: number }): PdfRect | null {
  const transform = item.transform
  if (!transform || transform.length < 6) return null
  const [a, b, c, d, e, f] = transform.map(Number)
  const along = Math.hypot(a, b)
  const across = Math.hypot(c, d)
  const size = Number(item.height) > 0 ? Number(item.height) : across
  const width = Math.max(0, Number(item.width) || 0)
  if (![a, b, c, d, e, f, size, width].every(Number.isFinite) || size <= 0) return null
  const dir: [number, number] = along > 0 ? [a / along, b / along] : [1, 0]
  const up: [number, number] = across > 0 ? [c / across, d / across] : [-dir[1], dir[0]]
  const descent = Number.isFinite(style?.descent) && style!.descent! < 0 ? style!.descent! : -0.2
  const ascent = Number.isFinite(style?.ascent) && style!.ascent! > 0 ? style!.ascent! : 0.8
  const low = descent * size
  const high = ascent * size
  const corner = (u: number, v: number): [number, number] => [e + dir[0] * u + up[0] * v, f + dir[1] * u + up[1] * v]
  const rect = boundsOf([corner(0, low), corner(width, low), corner(0, high), corner(width, high)])
  return rect.width > 0 || rect.height > 0 ? rect : null
}

export interface PageGeometry {
  /** Each painted image: bounds and, when known, the resolution it is drawn at. */
  images: Array<{ rect: PdfRect; dpi?: number }>
  vectorOps: number
}

interface OperatorListLike {
  fnArray: ArrayLike<number>
  argsArray: ArrayLike<unknown>
}

function imageDpi(matrix: Matrix, width: unknown, height: unknown): number | undefined {
  const pixelsWide = Number(width)
  const pixelsHigh = Number(height)
  const spanWide = Math.hypot(matrix[0], matrix[1])
  const spanHigh = Math.hypot(matrix[2], matrix[3])
  if (!(pixelsWide > 0 && pixelsHigh > 0 && spanWide > 0 && spanHigh > 0)) return undefined
  const dpi = Math.min(pixelsWide / (spanWide / 72), pixelsHigh / (spanHigh / 72))
  return Number.isFinite(dpi) ? Math.round(dpi) : undefined
}

/** Images (with their drawn resolution) and path operators of an operator list. */
export function pageGeometry(list: OperatorListLike, OPS: Record<string, number>): PageGeometry {
  const images: PageGeometry['images'] = []
  const stack: Matrix[] = []
  let matrix: Matrix = [...IDENTITY]
  let vectorOps = 0
  const pathOps = new Set([
    OPS.constructPath, OPS.fill, OPS.eoFill, OPS.stroke, OPS.closeStroke, OPS.fillStroke,
    OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke,
  ].filter((op) => op !== undefined))
  const addImage = (placement: Matrix, width?: unknown, height?: unknown) => {
    if (images.length >= PAGE_CLASSIFIER.maxRects) return
    const rect = unitSquareBounds(placement)
    if (!(rect.width > 0 && rect.height > 0)) return
    images.push({ rect, dpi: imageDpi(placement, width, height) })
  }
  const count = list.fnArray.length
  for (let index = 0; index < count; index += 1) {
    const op = list.fnArray[index]
    const args = (list.argsArray[index] || []) as unknown[]
    if (pathOps.has(op)) {
      vectorOps += 1
      continue
    }
    switch (op) {
      case OPS.save:
        stack.push([...matrix] as Matrix)
        break
      case OPS.restore:
        matrix = stack.pop() || [...IDENTITY]
        break
      case OPS.transform: {
        const next = asMatrix(args)
        if (next) matrix = multiply(matrix, next)
        break
      }
      case OPS.paintFormXObjectBegin: {
        stack.push([...matrix] as Matrix)
        const next = asMatrix(args[0])
        if (next) matrix = multiply(matrix, next)
        break
      }
      case OPS.paintFormXObjectEnd:
        matrix = stack.pop() || [...IDENTITY]
        break
      case OPS.paintImageXObject:
        addImage(matrix, args[1], args[2])
        break
      case OPS.paintInlineImageXObject:
      case OPS.paintImageMaskXObject: {
        const image = args[0] as { width?: unknown; height?: unknown } | undefined
        addImage(matrix, image?.width, image?.height)
        break
      }
      case OPS.paintSolidColorImageMask:
        addImage(matrix)
        break
      case OPS.paintImageXObjectRepeat: {
        const scaleX = Number(args[1])
        const scaleY = Number(args[2])
        const positions = args[3] as ArrayLike<number> | undefined
        if (!Number.isFinite(scaleX) || !Number.isFinite(scaleY) || !positions || typeof positions.length !== 'number') break
        for (let offset = 0; offset + 1 < positions.length; offset += 2) {
          addImage(multiply(matrix, [scaleX, 0, 0, scaleY, Number(positions[offset]), Number(positions[offset + 1])]))
        }
        break
      }
      case OPS.paintInlineImageXObjectGroup:
      case OPS.paintImageMaskXObjectGroup: {
        const entries = (op === OPS.paintInlineImageXObjectGroup ? args[1] : args[0]) as Array<{ transform?: unknown }> | undefined
        if (!Array.isArray(entries)) break
        for (const entry of entries) {
          const placement = asMatrix(entry?.transform)
          if (placement) addImage(multiply(matrix, placement))
        }
        break
      }
      default:
        break
    }
  }
  return { images, vectorOps }
}

/** Visible and invisible text of a page, in text-content order. */
export function textStats(
  items: readonly unknown[],
  styles: Record<string, { ascent?: number; descent?: number }> | undefined,
  paint: (fontName: string, text: string) => { invisible: boolean },
  fontName: (id: string) => string,
) {
  let visibleChars = 0
  let invisibleChars = 0
  let glyphless = false
  const visibleTextRects: PdfRect[] = []
  for (const raw of items) {
    const item = raw as TextItemLike
    if (typeof item?.str !== 'string') continue
    // The resolver consumes each item in content order, so call it for every one.
    const { invisible } = paint(String(item.fontName || ''), item.str)
    const characters = item.str.replace(/\s/gu, '').length
    if (!characters) continue
    if (invisible) {
      invisibleChars += characters
      if (!glyphless && GLYPHLESS_FONT.test(fontName(String(item.fontName || '')))) glyphless = true
      continue
    }
    visibleChars += characters
    if (visibleTextRects.length < PAGE_CLASSIFIER.maxRects) {
      const rect = textItemRect(item, styles?.[String(item.fontName || '')])
      if (rect) visibleTextRects.push(rect)
    }
  }
  return { visibleChars, invisibleChars, glyphless, visibleTextRects }
}

function loadedFontName(page: PDFPageProxy, id: string) {
  try {
    if (!page.commonObjs.has(id)) return ''
    return String((page.commonObjs.get(id) as { name?: unknown } | null)?.name || '')
  } catch {
    return ''
  }
}

async function computeScanState(page: PDFPageProxy): Promise<PageScanState> {
  const content = await getPageTextContent(page)
  // The default parameters share pdf.js' cached list with rendering and image detection.
  const list = await page.getOperatorList()
  let paint: (fontName: string, text: string) => { invisible: boolean } = () => ({ invisible: false })
  try {
    paint = textPaintResolver(list, pdfjs.OPS as unknown as Record<string, number>)
  } catch {
    // An unusual operator list: every text item then counts as visible.
  }
  const styles = content.styles as Record<string, { ascent?: number; descent?: number }> | undefined
  const text = textStats(content.items, styles, paint, (id) => loadedFontName(page, id))
  const view = pageViewRect(page.view)
  const geometry = pageGeometry(list, pdfjs.OPS as unknown as Record<string, number>)
  const imageRects = geometry.images
    .map((image) => intersectRects(image.rect, view))
    .filter((rect): rect is PdfRect => Boolean(rect))
  const viewArea = view.width * view.height
  let nativeDpi: number | undefined
  let largest = 0
  for (const image of geometry.images) {
    if (!image.dpi) continue
    const clipped = intersectRects(image.rect, view)
    const area = clipped ? clipped.width * clipped.height : 0
    if (viewArea > 0 && area / viewArea >= PAGE_CLASSIFIER.dpiCoverage && area > largest) {
      largest = area
      nativeDpi = image.dpi
    }
  }
  return scanStateFromStats({
    visibleChars: text.visibleChars,
    invisibleChars: text.invisibleChars,
    imageCoverage: imageCoverage(imageRects, view),
    hasImages: geometry.images.length > 0,
    vectorOps: geometry.vectorOps,
  }, { nativeDpi, ocrBySimple: text.glyphless, visibleTextRects: text.visibleTextRects, imageRects })
}

const pageStates = new WeakMap<PDFPageProxy, Promise<PageScanState>>()

/** Classify one page. Cached for the page's lifetime (every document change makes new pages). */
export function classifyPage(page: PDFPageProxy): Promise<PageScanState> {
  let request = pageStates.get(page)
  if (!request) {
    request = computeScanState(page)
    pageStates.set(page, request)
    request.catch(() => { if (pageStates.get(page) === request) pageStates.delete(page) })
  }
  return request
}

/** The classification of this page when it has already been computed (no work is started). */
export function knownPageScanState(page: PDFPageProxy): Promise<PageScanState> | undefined {
  return pageStates.get(page)
}

/**
 * Whether pdf.js already holds decoded images (or other objects) for this
 * page. Those belong to whoever loaded them (the viewer, image selection) and
 * must not be released by OCR: pdf.js closes the bitmaps on cleanup.
 */
export function holdsDecodedObjects(page: PDFPageProxy) {
  try {
    for (const entry of page.objs as unknown as Iterable<unknown>) {
      if (entry) return true
    }
  } catch {
    return true
  }
  return false
}

export interface ClassifyPagesOptions {
  signal?: AbortSignal
  /** Called after each page with the states known so far. */
  onProgress?: (done: number, total: number, states: ReadonlyMap<number, PageScanState>) => void
  /**
   * Whether a page's decoded images may be released after it is classified.
   * Pages on screen must keep theirs (image selection reuses them); every
   * other page is released, or classifying a long scan would hold every scan.
   */
  canRelease?: (pageIndex: number) => boolean
}

/** Classify several pages one after the other (a failed page counts as native). */
export async function classifyPages(pdf: PDFDocumentProxy, pageIndices: readonly number[], options: ClassifyPagesOptions = {}) {
  const states = new Map<number, PageScanState>()
  let done = 0
  for (const pageIndex of pageIndices) {
    if (options.signal?.aborted) throw ocrAbortError()
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= pdf.numPages) continue
    const page = await pdf.getPage(pageIndex + 1)
    // Release only what classifying loaded, never images someone else decoded.
    const owned = !knownPageScanState(page) && !holdsDecodedObjects(page)
    try {
      states.set(pageIndex, await classifyPage(page))
    } catch {
      states.set(pageIndex, scanStateFromStats({ visibleChars: PAGE_CLASSIFIER.minChars, invisibleChars: 0, imageCoverage: 0, hasImages: false, vectorOps: 0 }))
    }
    if (owned && (options.canRelease?.(pageIndex) ?? true)) page.cleanup()
    done += 1
    options.onProgress?.(done, pageIndices.length, states)
  }
  return states
}

/**
 * Whether a page is a scan without text, reading as little as possible: a
 * page with plenty of text (visible or recognised) cannot be one, so only the
 * others are fully classified.
 */
export async function pageNeedsOcr(page: PDFPageProxy, options: { release?: boolean } = {}): Promise<boolean> {
  const known = knownPageScanState(page)
  if (known) return (await known).needsOcr
  const content = await getPageTextContent(page)
  let characters = 0
  for (const item of content.items) {
    if ('str' in item) characters += item.str.replace(/\s/gu, '').length
    if (characters >= PAGE_CLASSIFIER.minChars * 2) return false
  }
  const owned = !holdsDecodedObjects(page)
  const state = await classifyPage(page)
  if (options.release && owned) page.cleanup()
  return state.needsOcr
}

/** Images on a scan page that cover the page: the scan itself, not something to select. */
export function isPageCoveringImage(rect: PdfRect, view: ArrayLike<number>) {
  const page = pageViewRect(view)
  const clipped = intersectRects(rect, page)
  const area = page.width * page.height
  return Boolean(clipped && area > 0 && (clipped.width * clipped.height) / area >= PAGE_CLASSIFIER.pageCoveringImage)
}
