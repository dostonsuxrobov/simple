import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist'
import { composePatch } from '../../../electron/ocr-retouch.mjs'
import { planScanReplacement } from '../../../electron/scan-replacement.mjs'
import type { ScanStyleEstimate } from '../../../electron/scan-style.mjs'
import { pdfjs } from '../pdf'
import type {
  PageTextEdit,
  PdfRect,
  ScanEditInfo,
  ScanPatch,
  ScanReplacement,
  ScanTextRun,
  ScanTextStyle,
  ScanWord,
  TextOverlay,
} from '../../types'
import type { OcrPageResult, PdfPoint } from './types'

// Edit scanned text (design 4.8, WP4). A click on recognised (invisible) text
// opens the whole line as an edit. The line is rendered again from the page
// at the scan's resolution and handed to the preparation worker, which
// measures the printed font (electron/scan-style.mjs) and retouches the old
// glyphs out of the scan (electron/ocr-retouch.mjs). The patch is an RGBA PNG
// whose alpha covers only those glyphs, so the scan image itself never
// changes. On commit only the words that changed are replaced
// (electron/scan-replacement.mjs): the others keep their scanned pixels and
// their recognised text. Retouching is not redaction: the original pixels
// stay in the file under the patch.

/** Main thread -> preparation worker. `rgba` is transferred. */
export interface ScanRetouchRequest {
  type: 'retouch-line'
  id: number
  rgba: ArrayBuffer
  width: number
  height: number
  dpi: number
  /** Baseline origin and unit direction in region pixels (y down). */
  baseline: { x: number; y: number; dx: number; dy: number }
  /** Line extent along the baseline, pixels. */
  length: number
  /** Em size of the OCR text, pixels. */
  fontSize: number
  text: string
  /** Word boxes from recognition, in region pixels. */
  targets?: Array<{ quad: [PdfPoint, PdfPoint, PdfPoint, PdfPoint] }>
  seed: string
}

export interface ScanRetouchWord {
  text?: string
  labels: number[]
  box: { x0: number; y0: number; x1: number; y1: number; u0: number; u1: number } | null
}

export type ScanRetouchResponse =
  | {
      type: 'retouched'
      id: number
      style: ScanStyleEstimate
      words: ScanRetouchWord[]
      segmented: boolean
      inkBox: { x0: number; y0: number; x1: number; y1: number } | null
      /** The whole line's patch, PNG encoded in the worker (or RGBA when it could not encode). */
      patch: { x: number; y: number; width: number; height: number; dataUrl?: string; rgba?: ArrayBuffer } | null
      /** What later patches for some of the words are composed from. */
      retouch: { width: number; height: number; bilevel: boolean; paperColor: number[]; filled: ArrayBuffer; labels: ArrayBuffer; blocked: ArrayBuffer }
    }
  | { type: 'retouch-error'; id: number; message: string }

/** The line a scan edit starts from (PDF space). */
export interface ScanLineGeometry {
  pageIndex: number
  text: string
  origin: PdfPoint
  /** Unit reading direction. */
  dir: PdfPoint
  /** Extent along dir, points. */
  length: number
  /** Size of the OCR text, points (Tesseract's row height for Simple's layers). */
  fontSize: number
  /** Exact word boxes when the page was recognised in this window. */
  words?: Array<{ text: string; quad: [PdfPoint, PdfPoint, PdfPoint, PdfPoint] }>
  /** Baseline distance to the neighbouring line of the paragraph, points. */
  lineSpacing?: number
}

/** Result of preparing a scan edit (what App merges into the edit). */
export interface ScanPreparation {
  key: string
  status: 'ready' | 'failed'
  error?: string
  patch?: ScanPatch
  words?: ScanWord[]
  inkRect?: PdfRect
  style?: ScanTextStyle
  lineSpacing?: number
}

interface RegionMapping {
  /** Page view origin (x0) and top (y1) in PDF space. */
  viewX: number
  viewTop: number
  /** Region offset in pixels at `scale` pixels per point. */
  px0: number
  py0: number
  scale: number
}

interface RetouchCacheEntry {
  key: string
  retouch: { width: number; height: number; bilevel: boolean; paperColor: number[]; filled: Uint8ClampedArray; labels: Int16Array; blocked: Uint8Array }
  mapping: RegionMapping
  dpi: number
  /** Mask labels per word, aligned with ScanPreparation.words. */
  wordLabels: number[][]
}

const RETOUCH_CACHE_LIMIT = 8
const RESULT_LIMIT = 32
const WORKER_IDLE_MS = 60_000
const MAX_REGION_PIXELS = 6_000_000
/** Extra room past the end of a line a minimal replacement may use, in em. */
const LINE_END_ROOM = 1
/** Ascent and descent of the invisible OCR font (GlyphLessFont, patched metrics). */
const OCR_ASCENT = 0.8
const OCR_DESCENT = 0.22

const retouchCache = new Map<string, RetouchCacheEntry>()
const preparations = new Map<string, Promise<ScanPreparation>>()
const results = new Map<string, ScanPreparation>()
const documentIds = new WeakMap<PDFDocumentProxy, number>()
let nextDocumentId = 1

function remember<T>(map: Map<string, T>, key: string, value: T, limit: number) {
  map.delete(key)
  map.set(key, value)
  while (map.size > limit) {
    const oldest = map.keys().next().value
    if (oldest === undefined) break
    map.delete(oldest)
  }
}

/** A line of a page, identified by its baseline: stable across renders and sessions. */
export function scanLineKey(pageIndex: number, origin: PdfPoint, dir: PdfPoint) {
  const degrees = Math.round((Math.atan2(dir.y, dir.x) * 180) / Math.PI)
  return `p${pageIndex}:${origin.x.toFixed(1)},${origin.y.toFixed(1)}:${degrees}`
}

/** Preparation key: the line in this document (a page change makes a new document proxy). */
export function scanPreparationKey(pdf: PDFDocumentProxy, lineKey: string) {
  let id = documentIds.get(pdf)
  if (!id) documentIds.set(pdf, id = nextDocumentId++)
  return `d${id}|${lineKey}`
}

/** The seed of a patch's grain: the line, so preview and save draw the same pixels. */
function grainSeed(key: string) {
  return key.replace(/^d\d+\|/, '')
}

// ---------------------------------------------------------------------------
// Geometry

export function runQuad(run: ScanTextRun, start = 0, length = run.length): [PdfPoint, PdfPoint, PdfPoint, PdfPoint] {
  const up = { x: -run.dir.y, y: run.dir.x }
  const at = (u: number, v: number) => ({ x: run.origin.x + run.dir.x * u + up.x * v, y: run.origin.y + run.dir.y * u + up.y * v })
  const ascent = OCR_ASCENT * run.fontSize
  const descent = OCR_DESCENT * run.fontSize
  return [at(start, ascent), at(start + length, ascent), at(start + length, -descent), at(start, -descent)]
}

export function quadBounds(points: PdfPoint[]): PdfRect {
  const xs = points.map((point) => point.x)
  const ys = points.map((point) => point.y)
  const x = Math.min(...xs)
  const y = Math.min(...ys)
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y }
}

function unionRects(rects: Array<PdfRect | null | undefined>): PdfRect | null {
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -Infinity
  let y1 = -Infinity
  for (const rect of rects) {
    if (!rect) continue
    x0 = Math.min(x0, rect.x)
    y0 = Math.min(y0, rect.y)
    x1 = Math.max(x1, rect.x + rect.width)
    y1 = Math.max(y1, rect.y + rect.height)
  }
  return Number.isFinite(x0) ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null
}

/** The text angle of a line: its baseline angle, or exactly 0 for lines within 0.5 degrees of level. */
export function scanTextAngle(dir: PdfPoint) {
  const angle = Math.atan2(dir.y, dir.x)
  return Math.abs(angle) < (0.5 * Math.PI) / 180 ? 0 : angle
}

/**
 * The text box of a run in the editor's and the saver's reading-frame
 * convention: centred on the run, `width` along it and `ascent + descent`
 * across it (axes swapped for sideways text), with the baseline `descent`
 * above the box's bottom edge. Laying text out in it puts its baseline on the
 * run's baseline, starting `start` points along it.
 */
export function scanTextFrame(origin: PdfPoint, dir: PdfPoint, start: number, width: number, ascent: number, descent: number) {
  const up = { x: -dir.y, y: dir.x }
  const height = ascent + descent
  const centre = {
    x: origin.x + dir.x * (start + width / 2) + up.x * ((ascent - descent) / 2),
    y: origin.y + dir.y * (start + width / 2) + up.y * ((ascent - descent) / 2),
  }
  const sideways = Math.abs(Math.sin(scanTextAngle(dir))) > 0.7
  const boxWidth = sideways ? height : width
  const boxHeight = sideways ? width : height
  return {
    rect: { x: centre.x - boxWidth / 2, y: centre.y - boxHeight / 2, width: boxWidth, height: boxHeight },
    baselineOffset: descent,
  }
}

// ---------------------------------------------------------------------------
// Fonts (main thread)

let measureContext: CanvasRenderingContext2D | null = null
function context2d() {
  if (!measureContext) measureContext = document.createElement('canvas').getContext('2d')
  return measureContext
}

export function scanFontCss(edit: Pick<PageTextEdit, 'fontStyle' | 'fontWeight' | 'fontFamily'>, size: number) {
  return `${edit.fontStyle || 'normal'} ${edit.fontWeight || 400} ${size}px ${edit.fontFamily}`
}

/** Width of text in points at `size` in the edit's font (1 CSS px = 1 pt here). */
export function measureScanText(edit: Pick<PageTextEdit, 'fontStyle' | 'fontWeight' | 'fontFamily'>, size: number, text: string) {
  const context = context2d()
  if (!context) return Array.from(text).length * size * 0.5
  context.font = scanFontCss(edit, size)
  return context.measureText(text).width
}

/** Ascent and descent of the line box the editor gives one line of this font (the CSS line box). */
function lineBox(edit: Pick<PageTextEdit, 'fontStyle' | 'fontWeight' | 'fontFamily'>, size: number, lineHeight: number) {
  const context = context2d()
  let ascent = 0.905 * size
  let descent = 0.212 * size
  if (context) {
    context.font = scanFontCss(edit, size)
    const metrics = context.measureText('Hg')
    if (Number.isFinite(metrics.fontBoundingBoxAscent) && metrics.fontBoundingBoxAscent > 0) ascent = metrics.fontBoundingBoxAscent
    if (Number.isFinite(metrics.fontBoundingBoxDescent) && metrics.fontBoundingBoxDescent >= 0) descent = metrics.fontBoundingBoxDescent
  }
  // The line's half-leading sits above and below the font's own box.
  const leading = (lineHeight - ascent - descent) / 2
  return { ascent: ascent + leading, descent: lineHeight - (ascent + leading) }
}

/**
 * Where the ink of text starts after the pen position (its left side
 * bearing) and how wide the ink is, in points at `size` in the edit's font.
 * Scanned word boxes are ink boxes, so this is what lines them up.
 */
function inkMetrics(edit: Pick<PageTextEdit, 'fontStyle' | 'fontWeight' | 'fontFamily'>, size: number, text: string) {
  const context = context2d()
  if (!context || !text.trim()) return { left: 0, width: measureScanText(edit, size, text) }
  context.font = scanFontCss(edit, size)
  const metrics = context.measureText(text)
  const left = -metrics.actualBoundingBoxLeft
  const width = metrics.actualBoundingBoxLeft + metrics.actualBoundingBoxRight
  return {
    left: Number.isFinite(left) ? left : 0,
    width: Number.isFinite(width) && width > 0 ? width : metrics.width,
  }
}

/**
 * Horizontal scale that makes the matched font as wide as the scanned words:
 * their total ink width over the font's (clamped).
 */
function widthScale(edit: Pick<PageTextEdit, 'fontStyle' | 'fontWeight' | 'fontFamily' | 'fontSize'>, words: ScanWord[] | undefined, lineText: string, lineLength: number) {
  let scanned = 0
  let natural = 0
  if (words?.length) {
    for (const word of words) {
      scanned += Math.max(0, word.end - word.start)
      natural += inkMetrics(edit, edit.fontSize, word.text).width
    }
  } else {
    scanned = lineLength
    natural = inkMetrics(edit, edit.fontSize, lineText.trim()).width
  }
  const scale = natural > 0 ? scanned / natural : 1
  return Math.max(0.75, Math.min(1.33, Number.isFinite(scale) ? scale : 1))
}

/**
 * The edit's whole-line text box for its current font, starting a side
 * bearing early so the first glyph's ink starts where the scanned ink did.
 */
function lineFrame(edit: PageTextEdit, scan: ScanEditInfo) {
  const lineHeight = Math.max(edit.fontSize * 0.8, edit.lineHeight || edit.fontSize * 1.18)
  const box = lineBox(edit, edit.fontSize, lineHeight)
  const bearing = inkMetrics(edit, edit.fontSize, scan.lineText.trim().slice(0, 1)).left * (edit.scaleX || 1)
  return scanTextFrame(scan.baseline.origin, scan.baseline.dir, -bearing, scan.run.length + bearing, box.ascent, box.descent)
}

// ---------------------------------------------------------------------------
// Building an edit

/**
 * A scan edit for a line, before preparation: the whole line, a provisional
 * style (from an earlier preparation of the same line when there is one) and
 * the paper colour that hides the old words until the patch is ready.
 */
export function createScanEdit(input: {
  pageIndex: number
  key: string
  line: ScanLineGeometry
  words?: ScanWord[]
  color: [number, number, number]
  paper: [number, number, number]
  caretOffset: number
}): PageTextEdit {
  const { line } = input
  const run: ScanTextRun = { origin: { ...line.origin }, dir: { ...line.dir }, length: line.length, fontSize: line.fontSize }
  const lineRect = quadBounds(runQuad(run))
  const prepared = results.get(input.key)
  const scan: ScanEditInfo = {
    key: input.key,
    mode: 'appearance',
    status: 'pending',
    lineRect,
    baseline: { origin: { ...line.origin }, dir: { ...line.dir } },
    lineText: line.text,
    run,
    words: input.words,
    paper: input.paper,
  }
  const text = line.text
  let edit: PageTextEdit = {
    pageIndex: input.pageIndex,
    rect: lineRect,
    originalRect: lineRect,
    inkRect: lineRect,
    originalText: text,
    text,
    // Never the OCR layer's own font: GlyphLessFont draws every character as a box.
    fontSize: Math.max(4, Math.min(96, Math.round(line.fontSize * 0.9 * 10) / 10)),
    fontFamily: 'Arial',
    fontWeight: 400,
    fontStyle: 'normal',
    textFit: 'fit',
    preserveSourceMetrics: false,
    lineHeight: undefined,
    letterSpacing: 0,
    scaleX: 1,
    angle: scanTextAngle(line.dir),
    direction: 'ltr',
    // Native text is authored in the PDF's own axes (as textEditFromSpan does).
    displayRotation: 0,
    align: 'left',
    color: input.color,
    backgroundColor: undefined,
    cover: true,
    scan,
    modified: false,
    caretOffset: input.caretOffset,
    selectionStart: input.caretOffset,
    selectionEnd: input.caretOffset,
  }
  edit = { ...edit, lineHeight: (line.lineSpacing && line.lineSpacing > edit.fontSize * 0.9 && line.lineSpacing < edit.fontSize * 2.2) ? line.lineSpacing : edit.fontSize * 1.18 }
  edit = withScanFrame(edit)
  return prepared ? applyScanPreparation(edit, prepared) : edit
}

/** Recompute the whole-line text box (and the width scale) for the edit's current font. */
export function withScanFrame(edit: PageTextEdit, options: { rescale?: boolean } = {}): PageTextEdit {
  const scan = edit.scan
  if (!scan) return edit
  const scaleX = options.rescale === false ? edit.scaleX : widthScale(edit, scan.words, scan.lineText, scan.run.length)
  const frame = lineFrame({ ...edit, scaleX }, scan)
  return { ...edit, rect: frame.rect, baselineOffset: frame.baselineOffset, scaleX }
}

/**
 * Merge a preparation into an edit (or a committed scan overlay) of that
 * line. The patch, words and ink box always apply; the matched style only
 * while the user has not restyled the text.
 */
export function applyScanPreparation<T extends PageTextEdit | TextOverlay>(edit: T, prepared: ScanPreparation): T {
  const scan = edit.scan
  if (!scan || scan.key !== prepared.key) return edit
  const nextScan: ScanEditInfo = {
    ...scan,
    status: prepared.status,
    patch: prepared.patch ?? scan.patch,
    words: prepared.words ?? scan.words,
    style: prepared.style ?? scan.style,
    paper: prepared.style?.background ?? scan.paper,
  }
  let next = { ...edit, scan: nextScan, inkRect: prepared.inkRect ?? edit.inkRect } as T
  const style = prepared.style
  if (style && !scan.userStyled) {
    const lineHeight = prepared.lineSpacing && prepared.lineSpacing > style.fontSize * 0.9 && prepared.lineSpacing < style.fontSize * 2.2
      ? prepared.lineSpacing
      : style.fontSize * 1.18
    next = {
      ...next,
      fontFamily: style.fontFamily,
      fontWeight: style.fontWeight,
      fontStyle: style.italic ? 'italic' : 'normal',
      fontSize: style.fontSize,
      lineHeight,
      color: style.color,
    }
  }
  if ('modified' in next) return withScanFrame(next as PageTextEdit, { rescale: !scan.userStyled }) as T
  // A committed overlay keeps the frame it was committed with, re-derived for the matched font.
  const asEdit = withScanFrame({ ...(next as TextOverlay), originalText: next.originalText ?? '', modified: true } as PageTextEdit, { rescale: !scan.userStyled })
  return { ...next, rect: asEdit.rect, baselineOffset: asEdit.baselineOffset, scaleX: asEdit.scaleX } as T
}

/** The edit that reopens a committed scan overlay: the whole line, as edited. */
export function scanEditFromOverlay(overlay: TextOverlay): PageTextEdit {
  const scan = overlay.scan as ScanEditInfo
  return {
    overlayId: overlay.id,
    pageIndex: overlay.pageIndex,
    rect: overlay.rect,
    originalRect: overlay.originalRect,
    originalText: overlay.originalText ?? scan.lineText,
    inkRect: overlay.inkRect,
    text: overlay.text,
    fontSize: overlay.fontSize,
    fontFamily: overlay.fontFamily,
    fontWeight: overlay.fontWeight,
    fontStyle: overlay.fontStyle,
    textFit: overlay.textFit,
    preserveSourceMetrics: overlay.preserveSourceMetrics,
    lineHeight: overlay.lineHeight,
    letterSpacing: overlay.letterSpacing,
    scaleX: overlay.scaleX,
    angle: overlay.angle,
    direction: overlay.direction,
    baselineOffset: overlay.baselineOffset,
    sourceSpaceWidth: overlay.sourceSpaceWidth,
    displayRotation: overlay.displayRotation,
    align: overlay.align || 'left',
    color: overlay.color,
    backgroundColor: overlay.backgroundColor,
    cover: overlay.cover,
    // The line reopens whole; the next commit decides again what changed.
    scan: { ...scan, replace: undefined },
    modified: false,
    caretOffset: overlay.text.length,
    selectionStart: overlay.text.length,
    selectionEnd: overlay.text.length,
  }
}

/** Whether a committed scan overlay belongs to the line whose baseline starts at `origin`. */
export function isSameScanLine(scan: ScanEditInfo, pageIndex: number, overlayPage: number, origin: PdfPoint, dir: PdfPoint) {
  if (pageIndex !== overlayPage) return false
  const tolerance = Math.max(0.6, scan.run.fontSize * 0.3)
  const dot = scan.baseline.dir.x * dir.x + scan.baseline.dir.y * dir.y
  return dot > 0.995
    && Math.hypot(scan.baseline.origin.x - origin.x, scan.baseline.origin.y - origin.y) <= tolerance
}

// ---------------------------------------------------------------------------
// OCR results (exact word boxes for pages recognised in this window)

/** The recognised line of `result` that starts at this baseline, with its paragraph's line spacing. */
export function findOcrLine(result: OcrPageResult | undefined, origin: PdfPoint, dir: PdfPoint, fontSize: number) {
  if (!result) return null
  const tolerance = Math.max(1, fontSize * 0.35)
  for (const paragraph of result.paragraphs) {
    for (const [index, line] of paragraph.lines.entries()) {
      const first = line.words[0]
      if (!first) continue
      const start = first.origin
      const dot = line.baseline.dir.x * dir.x + line.baseline.dir.y * dir.y
      if (dot < 0.995 || Math.hypot(start.x - origin.x, start.y - origin.y) > tolerance) continue
      const neighbour = paragraph.lines[index + 1] ?? paragraph.lines[index - 1]
      const spacing = neighbour?.words[0]
        ? Math.abs((neighbour.words[0].origin.x - start.x) * -dir.y + (neighbour.words[0].origin.y - start.y) * dir.x)
        : undefined
      return { line, lineSpacing: spacing }
    }
  }
  return null
}

/** Word boxes of a recognised line as scan words, when they match the line's text. */
export function scanWordsFromOcr(words: Array<{ text: string; quad: [PdfPoint, PdfPoint, PdfPoint, PdfPoint] }>, text: string, origin: PdfPoint, dir: PdfPoint): ScanWord[] | undefined {
  const tokens = text.trim().split(/\s+/u).filter(Boolean)
  if (tokens.length !== words.length) return undefined
  const along = (point: PdfPoint) => (point.x - origin.x) * dir.x + (point.y - origin.y) * dir.y
  const result: ScanWord[] = []
  for (const [index, word] of words.entries()) {
    if (word.text.normalize('NFC') !== tokens[index].normalize('NFC')) return undefined
    const positions = word.quad.map(along)
    result.push({ text: tokens[index], rect: quadBounds(word.quad), start: Math.min(...positions), end: Math.max(...positions) })
  }
  return result
}

// ---------------------------------------------------------------------------
// Preparation (region render on the main thread, the rest in the worker)

let worker: Worker | null = null
let workerIdle: ReturnType<typeof setTimeout> | null = null
let sequence = 0
const pending = new Map<number, { resolve: (response: ScanRetouchResponse) => void; reject: (error: unknown) => void }>()

function failPending(error: unknown) {
  const waiting = [...pending.values()]
  pending.clear()
  for (const request of waiting) request.reject(error)
}

function workerInstance() {
  if (worker) return worker
  const instance = new Worker(new URL('./ocrPrep.worker.ts', import.meta.url), { type: 'module', name: 'simple-scan-edit' })
  instance.addEventListener('message', (event: MessageEvent<ScanRetouchResponse>) => {
    const request = pending.get(event.data?.id)
    if (!request) return
    pending.delete(event.data.id)
    request.resolve(event.data)
    if (!pending.size) scheduleIdle()
  })
  instance.addEventListener('error', (event) => {
    if (worker === instance) worker = null
    instance.terminate()
    failPending(new Error(event.message || 'The scanned text could not be prepared.'))
  })
  worker = instance
  return instance
}

function scheduleIdle() {
  if (workerIdle) clearTimeout(workerIdle)
  workerIdle = setTimeout(() => {
    workerIdle = null
    if (!pending.size) {
      worker?.terminate()
      worker = null
    }
  }, WORKER_IDLE_MS)
}

function retouchInWorker(request: Omit<ScanRetouchRequest, 'id' | 'type'>) {
  const id = ++sequence
  return new Promise<ScanRetouchResponse>((resolve, reject) => {
    pending.set(id, { resolve, reject })
    if (workerIdle) { clearTimeout(workerIdle); workerIdle = null }
    try {
      workerInstance().postMessage({ type: 'retouch-line', id, ...request } satisfies ScanRetouchRequest, [request.rgba])
    } catch (error) {
      pending.delete(id)
      reject(error)
    }
  })
}

/** Start the worker ahead of the first edit of scanned text (its modules load while the user reads). */
export function warmScanEditWorker() {
  if (worker) return
  workerInstance()
  if (!pending.size) scheduleIdle()
}

/** Stop the worker (it restarts on demand). */
export function disposeScanEditWorker() {
  if (workerIdle) { clearTimeout(workerIdle); workerIdle = null }
  worker?.terminate()
  worker = null
  failPending(new DOMException('Stopped.', 'AbortError'))
}

/** Scan resolution to work at: the scan's own between 300 and 600 DPI, else 300. */
export function scanEditDpi(nativeDpi?: number) {
  return Math.max(300, Math.min(600, Number.isFinite(nativeDpi) && nativeDpi ? nativeDpi : 300))
}

async function renderRegion(page: PDFPageProxy, region: PdfRect, dpi: number) {
  const view = page.view
  const viewX = Math.min(view[0], view[2])
  const viewTop = Math.max(view[1], view[3])
  let scale = dpi / 72
  const pixels = (region.width * scale) * (region.height * scale)
  if (pixels > MAX_REGION_PIXELS) scale *= Math.sqrt(MAX_REGION_PIXELS / pixels)
  const px0 = Math.floor((region.x - viewX) * scale)
  const py0 = Math.floor((viewTop - (region.y + region.height)) * scale)
  const px1 = Math.ceil((region.x + region.width - viewX) * scale)
  const py1 = Math.ceil((viewTop - region.y) * scale)
  const width = Math.max(1, px1 - px0)
  const height = Math.max(1, py1 - py0)
  // Unrotated PDF space, so the patch can be drawn into the page as it is.
  const viewport = page.getViewport({ scale, rotation: 0, offsetX: -px0, offsetY: -py0 })
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext('2d', { alpha: false, willReadFrequently: true })
  if (!context) throw new Error('The scanned text could not be prepared.')
  try {
    await page.render({
      canvasContext: context,
      viewport,
      annotationMode: pdfjs.AnnotationMode.DISABLE,
      background: '#ffffff',
      intent: document.visibilityState === 'hidden' ? 'print' : 'display',
    }).promise
    const rgba = context.getImageData(0, 0, width, height).data
    return { rgba, width, height, mapping: { viewX, viewTop, px0, py0, scale } satisfies RegionMapping }
  } finally {
    canvas.width = 0
    canvas.height = 0
  }
}

const toPdfX = (mapping: RegionMapping, x: number) => mapping.viewX + (mapping.px0 + x) / mapping.scale
const toPdfY = (mapping: RegionMapping, y: number) => mapping.viewTop - (mapping.py0 + y) / mapping.scale

function boxToPdf(mapping: RegionMapping, box: { x0: number; y0: number; x1: number; y1: number }): PdfRect {
  const x = toPdfX(mapping, box.x0)
  const y = toPdfY(mapping, box.y1)
  return { x, y, width: (box.x1 - box.x0) / mapping.scale, height: (box.y1 - box.y0) / mapping.scale }
}

function pngDataUrl(rgba: Uint8ClampedArray, width: number, height: number) {
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext('2d')
  if (!context) throw new Error('The patch could not be encoded.')
  context.putImageData(new ImageData(new Uint8ClampedArray(rgba), width, height), 0, 0)
  const url = canvas.toDataURL('image/png')
  canvas.width = 0
  canvas.height = 0
  return url
}

/**
 * Prepare a scan edit: render the line's region at the scan's resolution,
 * then measure the font and retouch the old glyphs in the worker. Memoized
 * per key; the result is also kept for later edits of the same line.
 */
export function prepareScanEdit(page: PDFPageProxy, key: string, line: ScanLineGeometry, options: { nativeDpi?: number } = {}): Promise<ScanPreparation> {
  const existing = preparations.get(key)
  if (existing) return existing
  const request = runPreparation(page, key, line, options).catch((error): ScanPreparation => ({
    key, status: 'failed', error: error instanceof Error ? error.message : String(error),
  })).then((result) => {
    remember(results, key, result, RESULT_LIMIT)
    // A failed preparation may be retried by the next edit of the line.
    if (result.status === 'failed') preparations.delete(key)
    return result
  })
  remember(preparations, key, request, RESULT_LIMIT)
  return request
}

async function runPreparation(page: PDFPageProxy, key: string, line: ScanLineGeometry, options: { nativeDpi?: number }): Promise<ScanPreparation> {
  const run: ScanTextRun = { origin: line.origin, dir: line.dir, length: line.length, fontSize: line.fontSize }
  const lineHeight = Math.max(line.fontSize * 1.1, line.lineSpacing ?? line.fontSize * 1.2)
  // The line grown by a third of a line height above and below and a quarter em at each end.
  const up = { x: -line.dir.y, y: line.dir.x }
  const corner = (u: number, v: number) => ({ x: line.origin.x + line.dir.x * u + up.x * v, y: line.origin.y + line.dir.y * u + up.y * v })
  const grow = 0.35 * lineHeight
  const ends = 0.25 * line.fontSize
  const view = page.view
  const page_ = { x0: Math.min(view[0], view[2]), y0: Math.min(view[1], view[3]), x1: Math.max(view[0], view[2]), y1: Math.max(view[1], view[3]) }
  const wanted = quadBounds([
    corner(-ends, OCR_ASCENT * line.fontSize + grow), corner(line.length + ends, OCR_ASCENT * line.fontSize + grow),
    corner(-ends, -OCR_DESCENT * line.fontSize - grow), corner(line.length + ends, -OCR_DESCENT * line.fontSize - grow),
  ])
  const x0 = Math.max(page_.x0, wanted.x)
  const y0 = Math.max(page_.y0, wanted.y)
  const x1 = Math.min(page_.x1, wanted.x + wanted.width)
  const y1 = Math.min(page_.y1, wanted.y + wanted.height)
  if (!(x1 > x0 && y1 > y0)) throw new Error('The scanned line lies outside the page.')
  const dpi = scanEditDpi(options.nativeDpi)
  const render = await renderRegion(page, { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }, dpi)
  const { mapping } = render
  const toPixels = (point: PdfPoint) => ({ x: (point.x - mapping.viewX) * mapping.scale - mapping.px0, y: (mapping.viewTop - point.y) * mapping.scale - mapping.py0 })
  const base = toPixels(line.origin)
  const response = await retouchInWorker({
    rgba: render.rgba.buffer as ArrayBuffer,
    width: render.width,
    height: render.height,
    dpi: mapping.scale * 72,
    baseline: { x: base.x, y: base.y, dx: line.dir.x, dy: -line.dir.y },
    length: line.length * mapping.scale,
    fontSize: line.fontSize * mapping.scale,
    text: line.text,
    targets: line.words?.map((word) => ({ quad: word.quad.map(toPixels) as [PdfPoint, PdfPoint, PdfPoint, PdfPoint] })),
    seed: grainSeed(key),
  })
  if (response.type !== 'retouched') throw new Error(response.message || 'The scanned text could not be prepared.')
  const tokens = line.text.trim().split(/\s+/u).filter(Boolean)
  // Words: recognition's (exact) or the segmented ones matched to the text.
  let words: ScanWord[] | undefined
  const wordLabels: number[][] = []
  const fromWords = response.words.filter((word) => word.box)
  if (line.words?.length && response.words.length === tokens.length && fromWords.length === tokens.length) {
    words = response.words.map((word, index) => {
      wordLabels.push(word.labels)
      return { text: tokens[index], rect: boxToPdf(mapping, word.box!), start: word.box!.u0 / mapping.scale, end: word.box!.u1 / mapping.scale }
    })
  } else if (response.segmented && response.words.length === tokens.length && fromWords.length === tokens.length) {
    words = response.words.map((word, index) => {
      wordLabels.push(word.labels)
      return { text: tokens[index], rect: boxToPdf(mapping, word.box!), start: word.box!.u0 / mapping.scale, end: word.box!.u1 / mapping.scale }
    })
  }
  const retouch = response.retouch
  remember(retouchCache, key, {
    key,
    retouch: {
      width: retouch.width,
      height: retouch.height,
      bilevel: retouch.bilevel,
      paperColor: retouch.paperColor,
      filled: new Uint8ClampedArray(retouch.filled),
      labels: new Int16Array(retouch.labels),
      blocked: new Uint8Array(retouch.blocked),
    },
    mapping,
    dpi: mapping.scale * 72,
    wordLabels,
  }, RETOUCH_CACHE_LIMIT)
  let patch: ScanPatch | undefined
  if (response.patch) {
    const dataUrl = response.patch.dataUrl
      ?? (response.patch.rgba ? pngDataUrl(new Uint8ClampedArray(response.patch.rgba), response.patch.width, response.patch.height) : '')
    if (dataUrl) {
      patch = {
        rect: boxToPdf(mapping, { x0: response.patch.x, y0: response.patch.y, x1: response.patch.x + response.patch.width, y1: response.patch.y + response.patch.height }),
        dataUrl,
        dpi: Math.round(mapping.scale * 72),
      }
    }
  }
  const style: ScanTextStyle = {
    ...response.style,
    color: response.style.color.map((channel) => Math.max(0, Math.min(1, channel))) as [number, number, number],
    background: response.style.background.map((channel) => Math.max(0, Math.min(1, channel))) as [number, number, number],
  }
  return {
    key,
    status: 'ready',
    patch,
    words,
    inkRect: response.inkBox ? boxToPdf(mapping, response.inkBox) : undefined,
    style,
    lineSpacing: line.lineSpacing,
  }
}

/** The finished preparation of a line, when there is one. */
export function scanPreparationResult(key: string) {
  return results.get(key)
}

/**
 * Wait (at most `timeout` ms) for the preparations of scan overlays that are
 * still pending, and merge them: a save never writes an edit of scanned text
 * without its patch when the patch is about to arrive.
 */
export async function settleScanOverlays<T extends { type: string }>(overlays: T[], timeout = 5_000): Promise<T[]> {
  const waiting = overlays.filter((overlay) => overlay.type === 'text' && (overlay as unknown as TextOverlay).scan?.status === 'pending') as unknown as TextOverlay[]
  if (!waiting.length) return overlays
  const keys = [...new Set(waiting.map((overlay) => overlay.scan!.key))]
  const timer = new Promise<void>((resolve) => setTimeout(resolve, timeout))
  await Promise.race([Promise.all(keys.map((key) => preparations.get(key) ?? Promise.resolve(undefined))), timer])
  return overlays.map((overlay) => {
    const scan = (overlay as unknown as TextOverlay).scan
    const prepared = scan?.status === 'pending' ? results.get(scan.key) : undefined
    if (!prepared) return overlay
    const merged = applyScanPreparation(overlay as unknown as TextOverlay, prepared)
    const finalized = finalizeScanEdit({ ...merged, originalText: merged.originalText ?? '', modified: true } as PageTextEdit)
    return (finalized.kind === 'edit' ? { ...merged, ...pickOverlayFields(finalized.edit) } : merged) as unknown as T
  })
}

function pickOverlayFields(edit: PageTextEdit): Partial<TextOverlay> {
  return { rect: edit.rect, baselineOffset: edit.baselineOffset, scan: edit.scan }
}

// ---------------------------------------------------------------------------
// Commit: replace only the words that changed

function narrowedPatch(key: string, labels: Set<number>): ScanPatch | undefined {
  const entry = retouchCache.get(key)
  if (!entry || !labels.size) return undefined
  const patch = composePatch(entry.retouch, labels)
  if (!patch) return undefined
  const { mapping } = entry
  return {
    rect: boxToPdf(mapping, { x0: patch.x, y0: patch.y, x1: patch.x + patch.width, y1: patch.y + patch.height }),
    dataUrl: pngDataUrl(patch.rgba, patch.width, patch.height),
    dpi: Math.round(entry.dpi),
  }
}

export type FinalizedScanEdit = { kind: 'none' } | { kind: 'edit'; edit: PageTextEdit }

/** Whether the edit still sits in its line's own text box, laid out as the line was. */
function keepsLineLayout(edit: PageTextEdit, scan: ScanEditInfo) {
  const frame = lineFrame(edit, scan).rect
  const close = (a: number, b: number) => Math.abs(a - b) <= 0.5
  return close(edit.rect.x, frame.x) && close(edit.rect.y, frame.y) && close(edit.rect.width, frame.width) && close(edit.rect.height, frame.height)
    && (edit.align || 'left') === 'left'
    && Math.abs((edit.angle || 0) - scanTextAngle(scan.baseline.dir)) < 1e-6
    && edit.textFit !== 'wrap'
}

/**
 * Decide, before an edit of scanned text is committed, what it replaces
 * (design 4.8.5): nothing (the text and its look are unchanged), some words
 * (only their glyphs are retouched and only their recognised text is
 * removed; the patch is rebuilt for just those words) or the whole line (a
 * longer change, a restyled or moved box). Synchronous: the patch for a few
 * words is composed from the cached retouch of the line.
 */
export function finalizeScanEdit(edit: PageTextEdit): FinalizedScanEdit {
  const scan = edit.scan
  if (!scan) return { kind: 'edit', edit }
  const base: ScanEditInfo = { ...scan, replace: undefined }
  const recognizedOnly = scan.mode === 'recognized-text'
  const scaleX = edit.scaleX || 1
  const spacing = edit.letterSpacing || 0
  const measure = recognizedOnly
    ? () => 0
    : (text: string) => measureScanText(edit, edit.fontSize, text) * scaleX + Math.max(0, Array.from(text).length - 1) * spacing
  const words = scan.status === 'ready' ? scan.words : undefined
  const lineEnd = words?.length ? words[words.length - 1].end + LINE_END_ROOM * edit.fontSize : scan.run.length + LINE_END_ROOM * edit.fontSize
  const plan = planScanReplacement({
    words: words ?? [],
    along: words?.map((word) => ({ start: word.start, end: word.end })),
    originalText: scan.lineText,
    newText: edit.text,
    measure,
    spaceWidth: recognizedOnly ? 0 : measure(' '),
    lineEnd,
  })
  // A box the user moved, resized, turned, re-aligned or set to wrap is
  // re-typeset where it now is, and a restyled line in its new style: the
  // whole line, also when its words did not change.
  const relaidOut = !recognizedOnly && (Boolean(scan.userStyled) || !keepsLineLayout(edit, scan))
  if (plan.kind === 'none') return relaidOut ? { kind: 'edit', edit: { ...edit, scan: base } } : { kind: 'none' }
  if (plan.kind === 'line' || !words || relaidOut) return { kind: 'edit', edit: { ...edit, scan: base } }
  let { first, last } = plan
  let text = plan.text
  let originalText = plan.originalText
  if (last < first) {
    // A pure insertion (only possible for invisible text): attach it to a neighbouring word.
    if (first > 0) { first -= 1; text = `${words[first].text} ${text}`; originalText = words[first].text }
    else { last = first; text = `${text} ${words[first].text}`; originalText = words[first].text }
  }
  const changed = words.slice(first, last + 1)
  const start = changed[0].start
  const end = changed[changed.length - 1].end
  const run: ScanTextRun = {
    origin: { x: scan.run.origin.x + scan.run.dir.x * start, y: scan.run.origin.y + scan.run.dir.y * start },
    dir: scan.run.dir,
    length: Math.max(0, end - start),
    fontSize: scan.run.fontSize,
  }
  const lineHeight = Math.max(edit.fontSize * 0.8, edit.lineHeight || edit.fontSize * 1.18)
  const box = lineBox(edit, edit.fontSize, lineHeight)
  const room = recognizedOnly ? end - start : Math.max(end - start, plan.room)
  // The new word's ink starts where the old word's did.
  const bearing = recognizedOnly ? 0 : inkMetrics(edit, edit.fontSize, text.trim().slice(0, 1)).left * scaleX
  const frame = scanTextFrame(scan.baseline.origin, scan.baseline.dir, start - bearing, room + bearing, box.ascent, box.descent)
  let patch: ScanPatch | undefined
  if (!recognizedOnly) {
    const entry = retouchCache.get(scan.key)
    const labels = new Set<number>(entry ? changed.flatMap((_, offset) => entry.wordLabels[first + offset] ?? []) : [])
    patch = narrowedPatch(scan.key, labels)
    // Without the line's retouch at hand the words cannot be patched alone: replace the line.
    if (!patch) return { kind: 'edit', edit: { ...edit, scan: base } }
  }
  const replace: ScanReplacement = {
    first,
    last,
    originalText,
    text,
    rect: frame.rect,
    originalRect: unionRects(changed.map((word) => word.rect)) ?? quadBounds(runQuad(run)),
    baselineOffset: frame.baselineOffset,
    run,
    patch,
  }
  return { kind: 'edit', edit: { ...edit, scan: { ...base, replace } } }
}

/** What a scan overlay draws (the changed words, or the whole line). */
export function scanOverlayDrawing(overlay: TextOverlay) {
  const scan = overlay.scan
  const replace = scan?.replace
  return {
    text: replace ? replace.text : overlay.text,
    rect: replace ? replace.rect : overlay.rect,
    baselineOffset: replace ? replace.baselineOffset : overlay.baselineOffset,
    patch: replace ? replace.patch : scan?.patch,
    originalText: replace ? replace.originalText : overlay.originalText,
  }
}
