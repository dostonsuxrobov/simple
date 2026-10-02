import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, PointerEvent as ReactPointerEvent } from 'react'
import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist'
import { Check, Crop, ExternalLink, GripHorizontal, LoaderCircle, X } from 'lucide-react'
import type {
  ActiveSearchMatch,
  DetectedPageObject,
  DisplayRotation,
  PageObjectEdit,
  PageTextEdit,
  PdfFormValue,
  PdfOverlay,
  PdfRect,
  ScanPatch,
  ToolMode,
  TextOverlay,
} from '../types'
import { layoutText, resolveTextFit } from '../../electron/text-layout.mjs'
import { textPaintResolver } from '../../electron/text-appearance.mjs'
import { boundedCanvasSize } from '../../electron/canvas-size.mjs'
import { getPageTextContent, pdfjs, pdfRectToViewport, viewportRectToPdf } from '../lib/pdf'
import { detectPageObjects } from '../lib/pageObjects'
import { classifyPage, isPageCoveringImage, type PageScanState } from '../lib/ocr/pageClassifier'
import { ocrResultForPage } from '../lib/ocr/ocrCache'
import {
  createScanEdit,
  findOcrLine,
  isSameScanLine,
  prepareScanEdit,
  quadBounds,
  runQuad,
  scanEditFromOverlay,
  scanLineKey,
  scanOverlayDrawing,
  scanPreparationKey,
  scanPreparationResult,
  scanWordsFromOcr,
  warmScanEditWorker,
  type ScanLineGeometry,
  type ScanPreparation,
} from '../lib/ocr/scanEdit'
import { resizeTextEditRect } from '../lib/editClipboard'
import { textBackground } from '../lib/textBackground'
import { findTextLayerSearchRects } from '../lib/search'
import {
  caretOffsetAtPoint,
  selectionClientRectsWithin,
  selectedTextSliceWithinSpan,
} from '../lib/textSelection'
import { clamp, cx } from '../lib/utils'

interface ViewerRect {
  left: number
  top: number
  width: number
  height: number
}

interface SearchMatchGeometry {
  key: string
  rects: ViewerRect[]
}

type ResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'

interface TransformGesture {
  target: 'text' | 'object'
  mode: 'move' | 'resize'
  handle?: ResizeHandle
  startPoint: { x: number; y: number }
  startRect: ViewerRect
}

interface InlineTextVisualStyle {
  pageIndex: number
  sourceText: string
  sourceRect: PdfRect
  fontWeight: string
  fontStyle: string
  fontStretch: string
  lineHeight: number
  letterSpacing: number
  scaleX: number
  angle: number
  direction: 'ltr' | 'rtl'
  backgroundColor: string
}

interface SampledCanvasColors {
  text: [number, number, number]
  background: string
}

interface PdfCommonFont {
  name?: string
  loadedName?: string
  fallbackName?: string
  bold?: boolean
  black?: boolean
  italic?: boolean
  data?: Uint8Array
  widths?: ArrayLike<number> | Record<number, number>
  defaultWidth?: number
  toUnicode?: { _map?: ArrayLike<string | undefined> | Record<number, string | undefined> }
  systemFontInfo?: { css?: string; fontFamily?: string }
}

/** After recognising a page from the Edit-mode offer: open the text the user clicked. */
export interface PendingEditAt {
  pageIndex: number
  /** The clicked point in unrotated PDF space. */
  x: number
  y: number
  token: number
}

interface PdfPageProps {
  pdf: PDFDocumentProxy
  pageIndex: number
  zoom: number
  rotation: number
  tool: ToolMode
  overlays: PdfOverlay[]
  formValues: Record<string, PdfFormValue>
  textEdit: PageTextEdit | null
  objectEdit: PageObjectEdit | null
  activeSearchMatch: ActiveSearchMatch | null
  selectingObjectRegion: boolean
  onPageReady?: (size: { width: number; height: number }) => void
  onRequestTextEdit: (edit: PageTextEdit) => void
  onTextEditChange: (edit: PageTextEdit) => void
  onCommitTextEdit: () => void
  onCancelTextEdit: () => void
  onRequestObjectEdit: (edit: PageObjectEdit) => void
  onObjectEditChange: (edit: PageObjectEdit) => void
  onCommitObjectEdit: () => void
  onCancelObjectEdit: () => void
  onObjectRegionSelected: () => void
  onHighlight: (rects: PdfRect[]) => void
  onTextMarkup: (style: 'underline' | 'strikeout', rects: PdfRect[], displayRotation: DisplayRotation) => void
  onInk: (points: Array<{ x: number; y: number }>) => void
  onRectangle: (rect: PdfRect) => void
  onCrop: (rect: PdfRect) => void
  onPlaceSignature: (point: { x: number; y: number }, displayRotation: DisplayRotation) => void
  onNavigate: (pageIndex: number) => void
  // A method signature on purpose: list boxes send string[] values, and the
  // viewer in between still declares its pass-through as string | boolean.
  onFormChange(name: string, value: PdfFormValue): void
  /** A click in Edit mode landed on a scanned page that has no text yet. */
  onRequestOcrOffer?: (pageIndex: number, point: { x: number; y: number }, client: { x: number; y: number }) => void
  pendingEditAt?: PendingEditAt | null
  onPendingEditAtHandled?: (token: number) => void
  /** An edit of scanned text finished preparing (patch, words, matched style). */
  onScanEditPrepared?: (key: string, result: ScanPreparation) => void
}

const RESIZE_HANDLES: ResizeHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']

let textMeasurementCanvas: HTMLCanvasElement | null = null

function previewTextLayout(edit: PageTextEdit | TextOverlay, width: number, sourceScale?: number) {
  textMeasurementCanvas ||= document.createElement('canvas')
  const context = textMeasurementCanvas.getContext('2d')
  const size = edit.fontSize
  const spacing = edit.letterSpacing || 0
  if (context) context.font = `${edit.fontStyle || 'normal'} ${edit.fontWeight || 400} ${size}px ${edit.fontFamily}`
  const measure = (text: string) => (context?.measureText(text).width ?? Array.from(text).length * size * 0.5)
    + Math.max(0, Array.from(text).length - 1) * spacing
  let scaleX = sourceScale || edit.scaleX || 1
  if (edit.preserveSourceMetrics !== false && edit.originalText && edit.originalRect && Math.abs(edit.angle || 0) < 0.01) {
    const measured = measure(edit.originalText) * scaleX
    const correction = measured > 0 ? edit.originalRect.width / measured : 1
    if (correction >= 0.5 && correction <= 2) scaleX *= correction
  }
  const layout = layoutText(edit.text, width, (text: string) => measure(text) * scaleX, resolveTextFit(edit))
  return { ...layout, scaleX: scaleX * layout.fitScale, lineHeight: Math.max(size * 0.8, edit.lineHeight || size * 1.18) }
}

function colorCss(color: [number, number, number]) {
  return `rgb(${color.map((channel) => Math.round(clamp(channel, 0, 1) * 255)).join(', ')})`
}

function cssColorToPdf(color: string): [number, number, number] {
  const channels = color.match(/[\d.]+/g)?.slice(0, 3).map(Number)
  if (!channels || channels.length < 3 || channels.some((value) => !Number.isFinite(value))) return [1, 1, 1]
  return channels.map((value) => clamp(value / 255, 0, 1)) as [number, number, number]
}

function sameRect(a?: PdfRect, b?: PdfRect) {
  if (!a || !b) return false
  return Math.abs(a.x - b.x) < 0.01
    && Math.abs(a.y - b.y) < 0.01
    && Math.abs(a.width - b.width) < 0.01
    && Math.abs(a.height - b.height) < 0.01
}

function checkboxValueIsChecked(value: unknown) {
  if (typeof value === 'boolean') return value
  if (value === null || value === undefined) return false
  const normalized = String(value).trim().toLocaleLowerCase()
  return Boolean(normalized) && !['off', 'false', '0', 'no'].includes(normalized)
}

// Form field flags (PDF 32000-1, 12.7.3.1 and 12.7.4).
const FIELD_FLAG_PASSWORD = 0x2000
const FIELD_FLAG_EDITABLE_COMBO = 0x40000
const FIELD_FLAG_DO_NOT_SPELL_CHECK = 0x400000
const TEXT_ALIGNMENTS = ['left', 'center', 'right'] as const

interface PdfChoiceOption {
  exportValue: string
  displayValue: string
}

/** pdf.js widget annotation data this page reads (see pdf.js WidgetAnnotation). */
interface PdfWidgetAnnotation {
  id: string
  fieldName: string
  fieldType?: string
  fieldValue?: unknown
  fieldFlags?: number
  alternativeText?: string
  readOnly?: boolean
  hidden?: boolean
  multiLine?: boolean
  comb?: boolean
  maxLen?: number
  textAlignment?: number | null
  defaultAppearanceData?: { fontSize?: number }
  checkBox?: boolean
  radioButton?: boolean
  exportValue?: string
  buttonValue?: string | null
  combo?: boolean
  multiSelect?: boolean
  options?: PdfChoiceOption[]
}

/** A text field's value: what was typed, else the document's. */
function formTextValue(value: PdfFormValue | undefined, documentValue: unknown) {
  if (typeof value === 'string') return value
  return typeof documentValue === 'string' ? documentValue : ''
}

/** Whether this check box widget shows as checked. */
function checkBoxIsOn(annotation: PdfWidgetAnnotation, value: PdfFormValue | undefined) {
  if (typeof value === 'boolean') return value
  const current = value === undefined ? annotation.fieldValue : value
  // pdf.js reports each widget's own on-state; a field whose widgets have
  // different on-states is checked only on the matching one.
  return annotation.exportValue ? current === annotation.exportValue : checkboxValueIsChecked(current)
}

/** Selected export values of a choice field: what was chosen, else the document's. */
function choiceSelection(value: PdfFormValue | undefined, documentValue: unknown): string[] {
  const source = value === undefined ? documentValue : value
  if (Array.isArray(source)) return source.filter((item): item is string => typeof item === 'string')
  return typeof source === 'string' && source ? [source] : []
}

/** The editor's font size in CSS pixels: the field's own size, or auto. */
function formFieldFontSize(annotation: PdfWidgetAnnotation, boxHeight: number, zoom: number) {
  const declared = Number(annotation.defaultAppearanceData?.fontSize)
  if (declared > 0) return declared * zoom
  // Auto size: one line fills a single-line field (up to 12 pt); multi-line
  // and list fields use a small fixed size, as PDF viewers do.
  if (annotation.multiLine || (annotation.fieldType === 'Ch' && !annotation.combo)) return 10 * zoom
  return clamp((boxHeight / zoom - 2) / 1.2, 6, 12) * zoom
}

/** The text run at (or on the line next to) a point, and a point inside it. */
function textSpanNear(layer: HTMLElement, clientX: number, clientY: number) {
  let best: { span: HTMLSpanElement; rect: DOMRect; distance: number } | null = null
  for (const span of layer.querySelectorAll<HTMLSpanElement>('[data-text-item="true"]')) {
    if (span.dataset.textWhitespace === 'true' || span.dataset.editCovered === 'true') continue
    const rect = span.getBoundingClientRect()
    if (rect.width < 0.5 || rect.height < 1) continue
    const dx = clientX < rect.left ? rect.left - clientX : clientX > rect.right ? clientX - rect.right : 0
    const dy = clientY < rect.top ? rect.top - clientY : clientY > rect.bottom ? clientY - rect.bottom : 0
    // Only the clicked line counts (half a line of slack), not text further away.
    if (dy > rect.height * 0.5 || dx > rect.height * 3) continue
    const distance = Math.hypot(dx, dy)
    if (!best || distance < best.distance) best = { span, rect, distance }
  }
  if (!best) return null
  return {
    span: best.span,
    x: clamp(clientX, best.rect.left + 0.5, Math.max(best.rect.left + 0.5, best.rect.right - 0.5)),
    y: clamp(clientY, best.rect.top + 0.5, Math.max(best.rect.top + 0.5, best.rect.bottom - 0.5)),
  }
}

function textViewportSignature(viewport: { width: number; height: number; rotation: number; scale: number }) {
  return `${viewport.width}x${viewport.height}x${viewport.rotation}x${viewport.scale}`
}

function liveSelectionIntersectsTextLayer(layer: HTMLElement) {
  const viewer = layer.closest<HTMLElement>('.viewer')
  if (viewer?.dataset.pointerSelectingText === 'true') return true
  const selection = window.getSelection()
  if (!selection || selection.isCollapsed || selection.rangeCount < 1) return false
  if (layer.contains(selection.anchorNode) || layer.contains(selection.focusNode)) return true
  try {
    return selection.getRangeAt(0).intersectsNode(layer)
  } catch {
    return false
  }
}

function quoteCssFamily(family: string) {
  const source = family.trim()
  if (source.includes(',')) return source
  const clean = source.replace(/^['"]|['"]$/g, '')
  return clean && !/^(?:serif|sans-serif|monospace)$/i.test(clean)
    ? `"${clean.replace(/["\\]/g, '')}"`
    : clean
}

/** Convert PDF/PostScript face names into the family Windows and the saver use. */
function canonicalPdfFontFamily(sourceName: string | undefined, fallback = '') {
  let name = String(sourceName || fallback || '')
    .replace(/^['"]|['"]$/g, '')
    .replace(/^[A-Z]{6}\+/, '')
    .replace(/^\*+/, '')
    .replace(/[-,]\d{3,}$/, '')
  // A comma in a PostScript name normally separates the face, not a fallback.
  name = name.replace(/,(?:regular|roman|book|medium|light|semibold|demibold|bold|black|italic|oblique).*$/i, '')
  name = name.replace(/[-_](?:regular|roman|book|medium|light|semibold|demibold|bold|black|italic|oblique|bolditalic|boldoblique)(?:mt|ps)?$/i, '')
  name = name.replace(/(?:PSMT|PS|MT)$/i, '')
  name = name
    .replace(/([a-z\d])([A-Z])/g, '$1 $2')
    .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
    .replace(/[-_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

  const lower = name.toLowerCase().replace(/\s+/g, '')
  if (lower === 'timesnewroman') return 'Times New Roman'
  if (lower === 'arial') return 'Arial'
  if (lower === 'couriernew') return 'Courier New'
  if (lower === 'segoeui') return 'Segoe UI'
  if (lower === 'trebuchetms') return 'Trebuchet MS'
  if (lower === 'palatinolinotype') return 'Palatino Linotype'
  if (lower === 'ebgaramond') return 'EB Garamond'
  return name
}

function pdfFontTraits(font: PdfCommonFont | undefined, sourceName: string) {
  const face = sourceName.toLowerCase().replace(/[\s_-]/g, '')
  let weight = font?.black ? 900 : font?.bold ? 700 : 400
  if (/thin|hairline/.test(face)) weight = 100
  else if (/extralight|ultralight/.test(face)) weight = 200
  else if (/light/.test(face)) weight = 300
  else if (/medium/.test(face)) weight = 500
  else if (/semibold|demibold|demi/.test(face)) weight = 600
  else if (/extrabold|ultrabold/.test(face)) weight = 800
  else if (/black|heavy/.test(face)) weight = 900
  else if (/bold/.test(face)) weight = 700
  const italic = Boolean(font?.italic) || /italic|oblique|slanted/.test(face)
  return { weight, style: italic ? 'italic' as const : 'normal' as const }
}

function pdfFontCssFamily(font: PdfCommonFont | undefined, sourceName: string, fallback = '') {
  const canonical = canonicalPdfFontFamily(sourceName, fallback)
  const generic = /serif/i.test(fallback) && !/sans-serif/i.test(fallback)
    ? 'serif'
    : /mono/i.test(fallback)
      ? 'monospace'
      : 'sans-serif'
  // loadedName is the @font-face PDF.js used for the canvas. Keeping it first
  // makes the direct editor use the same outlines; the canonical family later
  // in the list lets the native saver match an installed Windows font when the
  // browser-decoded font has corrupt metrics (common for subset spaces).
  const candidates = [
    font?.loadedName,
    font?.systemFontInfo?.css,
    font?.systemFontInfo?.fontFamily,
    canonical,
    generic,
  ].filter((value): value is string => Boolean(value?.trim()))
  return [...new Set(candidates.map(quoteCssFamily))].join(', ')
}

function sourcePdfSpaceWidth(font: PdfCommonFont | undefined, text: string, itemWidth: number, fontSize: number, scaleX: number) {
  const unicodeMap = font?.toUnicode?._map
  const widths = font?.widths
  if (!unicodeMap || !widths || !text.includes(' ') || itemWidth <= 0 || fontSize <= 0) return undefined
  const mappedWidths = new Map<string, number>()
  for (const rawKey of Object.keys(unicodeMap)) {
    const cid = Number(rawKey)
    const unicode = unicodeMap[cid]
    const width = Number(widths[cid] ?? font?.defaultWidth)
    if (unicode && Number.isFinite(width) && width > 0 && !mappedWidths.has(unicode)) mappedWidths.set(unicode, width)
  }
  const spaceUnits = mappedWidths.get(' ')
  if (!spaceUnits) return undefined
  const mappings = [...mappedWidths.entries()].sort((a, b) => b[0].length - a[0].length)
  let sourceUnits = 0
  for (let offset = 0; offset < text.length;) {
    const match = mappings.find(([unicode]) => text.startsWith(unicode, offset))
    if (!match) return undefined
    sourceUnits += match[1]
    offset += match[0].length
  }
  const rawWidth = sourceUnits / 1000 * fontSize * scaleX
  if (!Number.isFinite(rawWidth) || rawWidth <= 0) return undefined
  return spaceUnits / 1000 * fontSize * scaleX * (itemWidth / rawWidth)
}

export const PdfPage = memo(function PdfPage({
  pdf, pageIndex, zoom, rotation, tool, overlays, formValues, textEdit, objectEdit, activeSearchMatch, selectingObjectRegion,
  onPageReady, onRequestTextEdit, onTextEditChange, onCommitTextEdit, onCancelTextEdit,
  onRequestObjectEdit, onObjectEditChange, onCommitObjectEdit, onCancelObjectEdit,
  onObjectRegionSelected, onHighlight, onCrop, onPlaceSignature, onNavigate, onFormChange,
  onTextMarkup, onInk, onRectangle, onRequestOcrOffer, pendingEditAt, onPendingEditAtHandled, onScanEditPrepared,
}: PdfPageProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const surfaceRef = useRef<HTMLDivElement>(null)
  const textLayerRef = useRef<HTMLDivElement>(null)
  const searchMatchLayerRef = useRef<HTMLDivElement>(null)
  const textEditorRef = useRef<HTMLTextAreaElement>(null)
  const transformRef = useRef<TransformGesture | null>(null)
  const textVisualStyleRef = useRef<InlineTextVisualStyle | null>(null)
  const fontDataRef = useRef(new Map<string, Uint8Array>())
  const [pageState, setPageState] = useState<{ page: PDFPageProxy; owner: PDFDocumentProxy } | null>(null)
  const page = pageState?.page || null
  const [loading, setLoading] = useState(true)
  const [renderViewport, setRenderViewport] = useState<ReturnType<PDFPageProxy['getViewport']> | null>(null)
  const [canvasReadyViewport, setCanvasReadyViewport] = useState<ReturnType<PDFPageProxy['getViewport']> | null>(null)
  const [renderError, setRenderError] = useState('')
  const [annotations, setAnnotations] = useState<any[]>([])
  const [imageCandidates, setImageCandidates] = useState<DetectedPageObject[]>([])
  // What the page holds (scan, recognised scan, native text…), for Edit mode.
  const [scanState, setScanState] = useState<{ page: PDFPageProxy; state: PageScanState | null } | null>(null)
  // The page the text layer's spans were built from (they survive a document swap until rebuilt).
  const textLayerPageRef = useRef<PDFPageProxy | null>(null)
  // Where the current press started, so only a plain click (not a drag) offers recognition.
  const offerPressRef = useRef<{ x: number; y: number } | null>(null)
  const [cropDraft, setCropDraft] = useState<ViewerRect | null>(null)
  const [cropOrigin, setCropOrigin] = useState<{ x: number; y: number } | null>(null)
  const [regionDraft, setRegionDraft] = useState<ViewerRect | null>(null)
  const [regionOrigin, setRegionOrigin] = useState<{ x: number; y: number } | null>(null)
  const [shapeDraft, setShapeDraft] = useState<ViewerRect | null>(null)
  const [shapeOrigin, setShapeOrigin] = useState<{ x: number; y: number } | null>(null)
  const [inkDraft, setInkDraft] = useState<Array<{ x: number; y: number }>>([])
  const [textLayerVersion, setTextLayerVersion] = useState(0)
  const [textLayerRefreshVersion, setTextLayerRefreshVersion] = useState(0)
  const [searchMatchGeometry, setSearchMatchGeometry] = useState<SearchMatchGeometry>({ key: '', rects: [] })
  const canvasReadyViewportRef = useRef(canvasReadyViewport)
  canvasReadyViewportRef.current = canvasReadyViewport
  const activeSearchKey = activeSearchMatch?.pageIndex === pageIndex
    ? `${activeSearchMatch.query}\u0000${activeSearchMatch.occurrenceIndex}`
    : ''

  useEffect(() => {
    let active = true
    setRenderError('')
    pdf.getPage(pageIndex + 1).then((nextPage) => {
      if (active) setPageState({ page: nextPage, owner: pdf })
    }).catch((error) => {
      if (active) setRenderError(error instanceof Error ? error.message : 'Page could not be opened.')
    })
    return () => { active = false }
  }, [pdf, pageIndex])

  const viewport = useMemo(() => {
    if (!page) return null
    if (pageState?.owner !== pdf && canvasReadyViewportRef.current) return canvasReadyViewportRef.current
    const angle = (((page.rotate || 0) + rotation) % 360 + 360) % 360
    return page.getViewport({ scale: zoom, rotation: angle })
  }, [page, pageState?.owner, pdf, zoom, rotation])
  const pageVisualReady = Boolean(viewport && pageState?.owner === pdf && canvasReadyViewport === viewport)
  const displayedViewport = canvasReadyViewport && !pageVisualReady ? canvasReadyViewport : viewport
  const displayRotation = ((((viewport?.rotation || 0) % 360) + 360) % 360) as DisplayRotation
  // Only the source selection matters. Typing, dragging, and resizing the new
  // text must not repeatedly parse a book or remove a different source region.
  // Scanned text is hidden by its retouch patch instead: removing its
  // invisible OCR glyphs would not change the picture (they go on save).
  const removedTextKey = JSON.stringify([
    ...overlays.filter((edit): edit is TextOverlay => edit.type === 'text' && edit.pageIndex === pageIndex),
    ...(textEdit?.pageIndex === pageIndex ? [textEdit] : []),
  ].filter(edit => edit.cover && edit.originalRect && !edit.scan).map(edit => ({
    type: 'text', cover: true, originalRect: edit.originalRect, originalText: edit.originalText,
  })))
  const [backgroundPage, setBackgroundPage] = useState<{ owner: PDFDocumentProxy; key: string; page: PDFPageProxy; disposed: boolean } | null>(null)
  const paintingPage = backgroundPage?.owner === pdf && backgroundPage.key === removedTextKey && !backgroundPage.disposed ? backgroundPage.page : page

  useEffect(() => {
    if (removedTextKey === '[]') { setBackgroundPage(null); return }
    let cancelled = false
    let document: PDFDocumentProxy | null = null
    let entry: NonNullable<typeof backgroundPage> | null = null
    textBackground(pdf, pageIndex, JSON.parse(removedTextKey)).then(async result => {
      document = result
      if (cancelled) { await result.destroy(); return }
      const nextPage = await result.getPage(1)
      if (!cancelled) {
        entry = { owner: pdf, key: removedTextKey, page: nextPage, disposed: false }
        setBackgroundPage(entry)
      }
    }).catch(error => {
      if (!cancelled) setRenderError(error instanceof Error ? error.message : 'The text preview could not be prepared.')
    })
    return () => {
      cancelled = true
      if (entry) entry.disposed = true
      if (document) void document.destroy().catch(() => {})
    }
  }, [pdf, pageIndex, removedTextKey])

  useEffect(() => {
    if (!viewport) return
    setCropDraft(null)
    setRegionDraft(null)
    setShapeDraft(null)
    setInkDraft([])
  }, [viewport])

  useEffect(() => {
    if (!displayedViewport || !pageVisualReady) return
    onPageReady?.({ width: displayedViewport.width, height: displayedViewport.height })
  }, [displayedViewport, pageVisualReady, onPageReady])

  // Ctrl+wheel zoom arrives in bursts. While a canvas rendered at the same
  // rotation already exists, the layout effect below CSS-scales it for instant
  // feedback and the expensive pdf.js re-render waits for the gesture to settle.
  useEffect(() => {
    if (!viewport) {
      setRenderViewport(null)
      return
    }
    const ready = canvasReadyViewportRef.current
    if (!ready || ready.rotation !== viewport.rotation) {
      setRenderViewport(viewport)
      return
    }
    const timer = window.setTimeout(() => setRenderViewport(viewport), 140)
    return () => window.clearTimeout(timer)
  }, [viewport])

  useLayoutEffect(() => {
    const canvas = canvasRef.current
    if (!viewport || !canvas) return
    if (canvasReadyViewport && canvasReadyViewport !== viewport) return
    canvas.style.width = `${viewport.width}px`
    canvas.style.height = `${viewport.height}px`
    // Spans built for another viewport must never stay visible floating over
    // the rescaled canvas; the rebuild restores visibility when it lands.
    const layer = textLayerRef.current
    if (layer?.dataset.viewportSignature && layer.dataset.viewportSignature !== textViewportSignature(viewport)) {
      layer.style.visibility = 'hidden'
    }
  }, [viewport, canvasReadyViewport])

  useEffect(() => {
    if (!page || !renderViewport || !canvasRef.current) return
    if (removedTextKey !== '[]' && paintingPage === page) return
    let cancelled = false
    const canvas = canvasRef.current
    const stagingCanvas = document.createElement('canvas')
    // Chromium's grayscale canvas text is visibly softer than native PDF
    // viewers when it is rendered at exactly one backing pixel per CSS pixel.
    // A modest supersampling floor restores fine serifs on common 100%-scaled
    // Windows displays, while the area cap keeps oversized architectural pages
    // bounded. ContinuousPdfViewer still mounts only nearby canvases.
    const desiredDpr = Math.min(Math.max(window.devicePixelRatio || 1, 1.5), 2.5)
    const backing = boundedCanvasSize(renderViewport.width, renderViewport.height, desiredDpr)
    stagingCanvas.width = backing.width
    stagingCanvas.height = backing.height
    const context = stagingCanvas.getContext('2d', { alpha: false })
    if (!context) return
    // Render offscreen and copy only when complete. The visible canvas keeps
    // the previous page pixels throughout a document-proxy swap, so a page
    // operation never flashes white or falls back to a loading placeholder.
    context.fillStyle = '#fff'
    context.fillRect(0, 0, stagingCanvas.width, stagingCanvas.height)
    setLoading(true)
    setRenderError('')
    const renderTask = (paintingPage || page).render({
      canvasContext: context,
      viewport: renderViewport,
      transform: [stagingCanvas.width / renderViewport.width, 0, 0, stagingCanvas.height / renderViewport.height, 0, 0],
    })
    renderTask.promise.then(() => {
      if (!cancelled) {
        canvas.width = stagingCanvas.width
        canvas.height = stagingCanvas.height
        canvas.style.width = `${renderViewport.width}px`
        canvas.style.height = `${renderViewport.height}px`
        const visibleContext = canvas.getContext('2d', { alpha: false })
        visibleContext?.drawImage(stagingCanvas, 0, 0)
        canvas.dataset.textRemovals = removedTextKey
        setLoading(false)
        setCanvasReadyViewport(renderViewport)
      }
    }).catch((error) => {
      if (!cancelled && error?.name !== 'RenderingCancelledException') {
        setRenderError(error instanceof Error ? error.message : 'This page could not be rendered.')
        setLoading(false)
      }
    }).finally(() => {
      // Release the staging bitmap after the renderer has finished or canceled.
      // The visible canvas owns the completed copy and stays crisp while zooming.
      stagingCanvas.width = stagingCanvas.height = 1
    })
    return () => {
      cancelled = true
      renderTask.cancel()
    }
  }, [page, paintingPage, renderViewport, removedTextKey])

  useEffect(() => {
    if (!page || tool !== 'edit') {
      setImageCandidates([])
      return
    }
    let active = true
    detectPageObjects(page, pageIndex).then((items) => {
      if (active) setImageCandidates(items)
    }).catch(() => {
      if (active) setImageCandidates([])
    })
    return () => { active = false }
  }, [page, pageIndex, tool])

  useEffect(() => {
    if (!page || tool !== 'edit') return
    let active = true
    classifyPage(page).then((state) => {
      if (active) setScanState({ page, state })
    }).catch(() => {
      // Unknown: behave as on any page with text.
      if (active) setScanState({ page, state: null })
    })
    return () => { active = false }
  }, [page, tool])
  const scanKnown = scanState?.page === page
  const pageScan = scanKnown ? scanState.state : null
  // Editing recognised text is likely here: load the preparation worker now.
  useEffect(() => {
    if (tool === 'edit' && pageScan?.hasOcrText) warmScanEditWorker()
  }, [tool, pageScan])

  useEffect(() => {
    if (!page || !viewport || canvasReadyViewport !== viewport || !textLayerRef.current) return
    const currentPage = page
    const currentViewport = viewport
    let cancelled = false
    const layer = textLayerRef.current
    const viewportSignature = textViewportSignature(currentViewport)
    // The deferral below protects a live Range against node REMOVAL only.
    // Populating an empty layer that sits between the range's endpoints (a
    // page remounting inside a multi-page selection) merely inserts nodes and
    // leaves both boundary points untouched, so build it right away.
    if (layer.childElementCount > 0 && liveSelectionIntersectsTextLayer(layer)) {
      // The canvas underneath has already re-rendered. While the rebuild is
      // deferred, never leave spans built for a different viewport visible
      // floating over the new orientation.
      if (layer.dataset.viewportSignature && layer.dataset.viewportSignature !== viewportSignature) {
        layer.style.visibility = 'hidden'
      }
      let retryFrame = 0
      const retryWhenSafe = () => {
        window.cancelAnimationFrame(retryFrame)
        retryFrame = window.requestAnimationFrame(() => {
          if (!liveSelectionIntersectsTextLayer(layer)) {
            setTextLayerRefreshVersion((version) => version + 1)
          }
        })
      }
      document.addEventListener('selectionchange', retryWhenSafe)
      document.addEventListener('pointerup', retryWhenSafe, true)
      document.addEventListener('pointercancel', retryWhenSafe, true)
      return () => {
        window.cancelAnimationFrame(retryFrame)
        document.removeEventListener('selectionchange', retryWhenSafe)
        document.removeEventListener('pointerup', retryWhenSafe, true)
        document.removeEventListener('pointercancel', retryWhenSafe, true)
      }
    }
    layer.replaceChildren()
    textLayerPageRef.current = null

    async function renderText() {
      const content = await getPageTextContent(currentPage)
      let originalTextPaint: (font: string, text: string) => { color?: number[]; invisible: boolean } = () => ({ invisible: false })
      try {
        originalTextPaint = textPaintResolver(await currentPage.getOperatorList(), pdfjs.OPS as unknown as Record<string, number>)
      } catch {
        // A damaged/unsupported operator stream can still use canvas sampling.
      }
      if (cancelled) return
      try {
        await document.fonts.ready
      } catch {
        // FontFaceSet is best-effort; canvas fallback measurement below remains
        // valid when an embedded font cannot be exposed to the DOM.
      }
      if (cancelled) return
      const styles = content.styles as Record<string, { fontFamily?: string; ascent?: number; descent?: number; vertical?: boolean }>
      const spans: Array<{ span: HTMLSpanElement; width: number; angle: number; measured: number; flow: number; read: number; fontHeight: number }> = []
      const fragment = document.createDocumentFragment()
      const measurementContext = document.createElement('canvas').getContext('2d')
      const rotationRadians = (((currentViewport.rotation % 360) + 360) % 360) * (Math.PI / 180)
      const flowAxis = { x: -Math.sin(rotationRadians), y: Math.cos(rotationRadians) }
      const readAxis = { x: Math.cos(rotationRadians), y: Math.sin(rotationRadians) }

      // PDF.js encodes most line ends as empty items flagged hasEOL. They get
      // no span, so carry the marker onto the run that precedes them; copy
      // uses it (with the line ids below) to keep line breaks and spaces.
      let previousSpan: HTMLSpanElement | null = null
      for (const [itemIndex, rawItem] of content.items.entries()) {
        if (!('str' in rawItem)) continue
        if (!rawItem.str) {
          if ((rawItem as { hasEOL?: boolean }).hasEOL && previousSpan) previousSpan.dataset.textEol = 'true'
          continue
        }
        const item = rawItem as typeof rawItem & { transform: number[]; width: number; height: number; fontName: string; dir?: string; hasEOL?: boolean }
        const tx = pdfjs.Util.transform(currentViewport.transform, item.transform)
        const style = styles[item.fontName] || {}
        let fontObject: PdfCommonFont | undefined
        try {
          fontObject = currentPage.commonObjs.get(item.fontName)
        } catch {
          // Canvas rendering normally resolves this object first; the text
          // content family below remains the safe PDF.js fallback.
        }
        let angle = Math.atan2(tx[1], tx[0])
        let pdfAngle = Math.atan2(item.transform[1], item.transform[0])
        if (style.vertical) {
          angle += Math.PI / 2
          pdfAngle += Math.PI / 2
        }
        const fontHeight = Math.hypot(tx[2], tx[3])
        const pdfFontSize = Math.hypot(item.transform[2], item.transform[3])
        const pdfScaleX = pdfFontSize > 0
          ? Math.hypot(item.transform[0], item.transform[1]) / pdfFontSize
          : 1
        const sourceFontName = fontObject?.name || fontObject?.fallbackName || item.fontName
        const fontTraits = pdfFontTraits(fontObject, sourceFontName)
        const fontFamily = pdfFontCssFamily(fontObject, sourceFontName, style.fontFamily)
        const fontAscent = style.ascent
          ? style.ascent * fontHeight
          : style.descent
            ? (1 + style.descent) * fontHeight
            : fontHeight
        const span = document.createElement('span')
        span.textContent = item.str
        span.dir = item.dir || 'ltr'
        span.dataset.textItem = 'true'
        span.dataset.textId = `text-${pageIndex}-${itemIndex}`
        span.dataset.textItemIndex = String(itemIndex)
        span.dataset.fontSize = String(pdfFontSize || fontHeight / zoom)
        span.dataset.fontFamily = fontFamily
        span.dataset.fontWeight = String(fontTraits.weight)
        span.dataset.fontStyle = fontTraits.style
        span.dataset.pdfFontName = item.fontName
        span.dataset.pdfFontSourceName = sourceFontName
        const paint = originalTextPaint(item.fontName, item.str)
        if (paint.color) span.dataset.pdfTextColor = JSON.stringify(paint.color)
        // Text that paints nothing (render mode 3/7) is an OCR layer over a
        // scan, Simple's or another tool's: editing it edits the scanned words.
        if (paint.invisible) {
          span.dataset.pdfTextInvisible = 'true'
          if (/(^|\+)GlyphLessFont$/i.test(sourceFontName)) span.dataset.pdfTextOcr = 'simple'
        }
        span.dataset.pdfOriginX = String(item.transform[4])
        span.dataset.pdfTextWidth = String(style.vertical ? item.height : item.width)
        span.dataset.pdfTextScaleX = String(pdfScaleX)
        span.dataset.pdfTextAngle = String(pdfAngle)
        span.dataset.pdfBaselineY = String(item.transform[5])
        const pdfSpaceWidth = sourcePdfSpaceWidth(fontObject, item.str, item.width, pdfFontSize, pdfScaleX)
        if (Number.isFinite(pdfSpaceWidth)) span.dataset.pdfSourceSpaceWidth = String(pdfSpaceWidth)
        if (fontObject?.data?.byteLength) fontDataRef.current.set(item.fontName, fontObject.data)
        span.dataset.textAngle = String(angle)
        span.dataset.textBaselineX = String(tx[4])
        span.dataset.textBaselineY = String(tx[5])
        // PDF text origins are baselines. Rotated runs need their ascent moved
        // along that rotated axis; otherwise their transparent hit boxes land
        // above/below neighbouring visual lines.
        const isUnrotated = Math.abs(angle) < 1e-7
        const left = isUnrotated ? tx[4] : tx[4] + fontAscent * Math.sin(angle)
        const top = isUnrotated ? tx[5] - fontAscent : tx[5] - fontAscent * Math.cos(angle)
        span.style.left = `${left}px`
        span.style.top = `${top}px`
        span.style.fontSize = `${fontHeight}px`
        span.style.fontFamily = span.dataset.fontFamily
        span.style.fontWeight = span.dataset.fontWeight
        span.style.fontStyle = span.dataset.fontStyle
        span.style.transformOrigin = '0 0'
        span.style.userSelect = 'text'
        span.style.webkitUserSelect = 'text'
        span.setAttribute('role', 'presentation')
        if (!/[^\s\u0000]/u.test(item.str)) {
          // Keep whitespace in DOM/copy order, but never let an invisible,
          // overlapping space become the pointer's selection endpoint.
          span.dataset.textWhitespace = 'true'
          span.style.pointerEvents = 'none'
        }
        let measured = 0
        if (measurementContext) {
          try {
            measurementContext.font = `${fontTraits.style} ${fontTraits.weight} ${fontHeight}px ${fontFamily}`
            measured = measurementContext.measureText(item.str).width
          } catch {
            // Fall back to the untransformed DOM width after insertion.
          }
        }
        const advanceWidth = (style.vertical ? item.height : item.width) * currentViewport.scale
        span.dataset.textAdvance = String(advanceWidth)
        if (item.hasEOL) span.dataset.textEol = 'true'
        previousSpan = span
        spans.push({
          span,
          width: advanceWidth,
          angle,
          measured,
          flow: tx[4] * flowAxis.x + tx[5] * flowAxis.y,
          read: tx[4] * readAxis.x + tx[5] * readAxis.y,
          fontHeight,
        })
      }

      // Native ::selection paints the DOM interval between anchor and focus.
      // Append spans in visual reading order — baselines clustered into lines
      // along the page's flow axis, items sorted along the reading axis — so a
      // drag never leaps across lines when the content stream is out of order.
      // This also puts copied text into reading order.
      const lines: Array<{ flow: number; fontHeight: number; items: typeof spans }> = []
      for (const item of [...spans].sort((a, b) => a.flow - b.flow)) {
        const line = lines.at(-1)
        if (line && Math.abs(item.flow - line.flow) <= Math.max(1, Math.max(item.fontHeight, line.fontHeight) * 0.4)) {
          line.items.push(item)
        } else if (line && line.fontHeight < item.fontHeight * 0.85 && item.flow - line.flow <= item.fontHeight * 0.8) {
          // A superscript is raised (and shrunk) beyond the baseline rule
          // above, yet still overlaps the ascent of the taller run following
          // it. Merge them and let the taller run take over as the reference.
          line.items.push(item)
          line.flow = item.flow
          line.fontHeight = item.fontHeight
        } else {
          lines.push({ flow: item.flow, fontHeight: item.fontHeight, items: [item] })
        }
      }
      for (const [lineIndex, line] of lines.entries()) {
        for (const item of line.items) item.span.dataset.textLine = String(lineIndex)
      }
      // Side-by-side columns share flow clusters, so emitting whole clusters
      // would interleave the columns row by row. Split each cluster at reading
      // gaps of an em or more and pool the runs into column bands; every flush
      // emits complete bands left to right, keeping each column's lines
      // contiguous. Single-column pages form one band and come out unchanged.
      const bands: Array<{ start: number; end: number; lastFlow: number; items: typeof spans }> = []
      const overlappingBands = (segment: { start: number; end: number }) =>
        bands.filter((band) => segment.start <= band.end && band.start <= segment.end)
      const flushBands = () => {
        for (const band of [...bands].sort((a, b) => a.start - b.start)) {
          for (const entry of band.items) fragment.appendChild(entry.span)
        }
        bands.length = 0
      }
      for (const line of lines) {
        const segments: Array<{ start: number; end: number; items: typeof spans }> = []
        for (const item of [...line.items].sort((a, b) => a.read - b.read)) {
          const segment = segments.at(-1)
          const previous = segment?.items.at(-1)
          if (segment && previous && item.read - (previous.read + previous.width) <= Math.max(previous.fontHeight, item.fontHeight)) {
            segment.items.push(item)
            segment.end = Math.max(segment.end, item.read + item.width)
          } else {
            segments.push({ start: item.read, end: item.read + item.width, items: [item] })
          }
        }
        // A clear vertical break ends the current column layout; so does a run
        // bridging several bands (a full-width heading between column regions)
        // or two runs of one line landing in a single band (columns starting
        // beneath full-width text). Emit the finished bands before continuing.
        if (bands.length && bands.every((band) => line.flow - band.lastFlow > Math.max(1, line.fontHeight) * 3)) flushBands()
        const claimed = new Set<(typeof bands)[number]>()
        for (const segment of segments) {
          const overlapping = overlappingBands(segment)
          if (overlapping.length > 1 || (overlapping.length === 1 && claimed.has(overlapping[0]))) {
            flushBands()
            break
          }
          if (overlapping.length === 1) claimed.add(overlapping[0])
        }
        for (const segment of segments) {
          const band = overlappingBands(segment)[0]
          if (band) {
            band.start = Math.min(band.start, segment.start)
            band.end = Math.max(band.end, segment.end)
            band.lastFlow = line.flow
            band.items.push(...segment.items)
          } else {
            bands.push({ start: segment.start, end: segment.end, lastFlow: line.flow, items: segment.items })
          }
        }
      }
      flushBands()

      layer.appendChild(fragment)
      layer.dataset.viewportSignature = viewportSignature
      textLayerPageRef.current = currentPage
      layer.style.visibility = ''
      // Read every *untransformed* width before writing transforms. Measuring a
      // rotated bounding box uses the font height as its width; at 90 degrees
      // that previously produced scale factors above 20x and selection regions
      // thousands of pixels tall.
      const measuredWidths = spans.map(({ span, measured }) => measured > 0 ? measured : span.getBoundingClientRect().width)
      for (const [{ span, width, angle }, measured] of spans.map((item, index) => [item, measuredWidths[index]] as const)) {
        const scaleX = measured > 0 ? width / measured : 1
        span.dataset.textScaleX = String(scaleX)
        span.style.transform = `rotate(${angle}rad) scaleX(${scaleX})`
      }
      if (!cancelled) setTextLayerVersion((version) => version + 1)
    }

    renderText().catch(() => {})
    return () => {
      cancelled = true
      // A Range whose endpoints are removed is repainted by Chromium at a
      // seemingly unrelated line. Leave the old layer connected until the
      // pointer/selection is released; the effect above then refreshes it.
      if (!liveSelectionIntersectsTextLayer(layer)) {
        layer.replaceChildren()
        textLayerPageRef.current = null
      }
    }
  }, [page, viewport, canvasReadyViewport, zoom, pageIndex, textLayerRefreshVersion])

  useEffect(() => {
    let cancelled = false
    const match = activeSearchMatch
    setSearchMatchGeometry({ key: '', rects: [] })
    if (!activeSearchKey || !match || match.pageIndex !== pageIndex) return
    const frame = window.requestAnimationFrame(() => {
      const textLayer = textLayerRef.current
      const surface = surfaceRef.current
      if (cancelled || !textLayer || !surface) return
      const rects = findTextLayerSearchRects(textLayer, surface, match.query, match.occurrenceIndex)
      if (!cancelled) setSearchMatchGeometry({ key: activeSearchKey, rects })
    })
    return () => {
      cancelled = true
      window.cancelAnimationFrame(frame)
    }
  }, [activeSearchKey, activeSearchMatch, pageIndex, textLayerVersion, viewport])

  // Bring a hit into view once per activation. Zoom, rotation, refits and
  // text-layer rebuilds recompute the geometry too; those only repaint the
  // highlight, so reading elsewhere is never pulled back to an old hit.
  const scrolledSearchMatchRef = useRef<ActiveSearchMatch | null>(null)
  useEffect(() => {
    if (!activeSearchKey || searchMatchGeometry.key !== activeSearchKey || !searchMatchGeometry.rects.length) return
    if (!activeSearchMatch || scrolledSearchMatchRef.current === activeSearchMatch) return
    const match = activeSearchMatch
    const frame = window.requestAnimationFrame(() => {
      const rect = searchMatchLayerRef.current?.querySelector<HTMLElement>('[data-search-match-rect="true"]')
      if (!rect) return
      scrolledSearchMatchRef.current = match
      rect.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'auto' })
    })
    return () => window.cancelAnimationFrame(frame)
  }, [activeSearchKey, activeSearchMatch, searchMatchGeometry])

  useEffect(() => {
    if (!page || !viewport || canvasReadyViewport !== viewport) return
    let active = true
    page.getAnnotations({ intent: 'display' }).then((items) => {
      if (active) setAnnotations(items)
    }).catch(() => {
      if (active) setAnnotations([])
    })
    return () => { active = false }
  }, [page, viewport, canvasReadyViewport])

  const activeText = textEdit?.pageIndex === pageIndex ? textEdit : null
  const activeObject = objectEdit?.pageIndex === pageIndex ? objectEdit : null
  const activeTextBox = activeText && viewport ? pdfRectToViewport(viewport, activeText.rect) : null
  const activeObjectBox = activeObject && viewport ? pdfRectToViewport(viewport, activeObject.rect) : null
  const activeTextDisplayDelta = activeText
    ? (((displayRotation - (activeText.displayRotation ?? 0)) % 360) + 360) % 360
    : 0
  // Scanned text keeps the saver's frame convention (electron/main.cjs): a
  // box whose own text angle is near 90 degrees holds its reading width in
  // its PDF height, which matters on pages scanned (and shown) sideways.
  const scanFrameSideways = (edit: { angle?: number; originalRect?: PdfRect; displayRotation?: number }) => {
    const rectAngle = edit.originalRect ? (edit.angle || 0) : 0
    const frameSideways = Math.abs(Math.sin(rectAngle + (edit.displayRotation || 0) * Math.PI / 180)) > 0.7
    return frameSideways !== (displayRotation === 90 || displayRotation === 270)
  }
  const activeTextSideways = activeText?.scan
    ? scanFrameSideways(activeText)
    : activeTextDisplayDelta === 90 || activeTextDisplayDelta === 270
  const activeTextScreenAngle = activeTextDisplayDelta * Math.PI / 180 - (activeText?.angle || 0)
  const activeTextLayout = useMemo(() => activeText && activeTextBox
    ? previewTextLayout(activeText, (activeTextSideways ? activeTextBox.height : activeTextBox.width) / zoom)
    : null, [activeText, activeTextBox?.width, activeTextBox?.height, activeTextSideways, zoom, textLayerVersion])
  const activeTextRequiredHeight = activeTextLayout && activeText
    ? Math.max(activeText.fontSize, activeTextLayout.lines.length * activeTextLayout.lineHeight)
    : 0
  const activeTextOverflow = Boolean(activeTextLayout && activeTextBox
    && activeTextRequiredHeight > (activeTextSideways ? activeTextBox.width : activeTextBox.height) / zoom + 1)

  useLayoutEffect(() => {
    if (!activeText || !activeTextBox || !viewport || !activeText.modified || !activeTextOverflow
      || Math.abs(activeTextScreenAngle) > 0.01) return
    if (resolveTextFit(activeText) === 'fit' && !activeText.text.includes('\n') && activeText.preserveSourceMetrics !== false) return
    const height = Math.min(activeTextRequiredHeight * zoom, viewport.height - activeTextBox.top)
    if (height <= activeTextBox.height + 1) return
    const rect = viewportRectToPdf(viewport, {
      left: activeTextBox.left, top: activeTextBox.top,
      right: activeTextBox.left + activeTextBox.width, bottom: activeTextBox.top + height,
    })
    onTextEditChange({
      ...activeText, rect,
      baselineOffset: Number.isFinite(activeText.baselineOffset)
        ? Number(activeText.baselineOffset) + activeText.rect.y - rect.y
        : undefined,
    })
  }, [activeText, activeTextBox?.width, activeTextBox?.height, activeTextOverflow, activeTextRequiredHeight,
    activeTextScreenAngle, viewport, zoom, onTextEditChange])
  const activeTextKey = activeText
    ? `${activeText.overlayId || 'native'}:${activeText.originalText}:${activeText.originalRect?.x ?? activeText.rect.x}:${activeText.originalRect?.y ?? activeText.rect.y}`
    : ''

  useEffect(() => {
    if (!activeText || !textEditorRef.current) return
    const editor = textEditorRef.current
    const frame = window.requestAnimationFrame(() => {
      editor.focus()
      const fallback = clamp(activeText.caretOffset ?? activeText.text.length, 0, activeText.text.length)
      const start = clamp(activeText.selectionStart ?? fallback, 0, activeText.text.length)
      const end = clamp(activeText.selectionEnd ?? start, start, activeText.text.length)
      editor.setSelectionRange(start, end)
    })
    return () => window.cancelAnimationFrame(frame)
    // Focus only when a different text block opens; typing must not reset the caret.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTextKey])

  const pageOverlays = useMemo(
    () => overlays.filter((overlay) => overlay.pageIndex === pageIndex),
    [overlays, pageIndex],
  )
  // On a scanned page the picture covering the page is the scan itself: a
  // click there offers recognition (or edits recognised text), it never picks
  // up the whole page. Until the page is classified, such a picture is not
  // offered either. "Select artwork area" still edits any part of it.
  const scanCovered = !scanKnown || pageScan?.kind === 'scan' || pageScan?.kind === 'searchable-scan'
  const selectableImageCandidates = useMemo(() => imageCandidates.filter((candidate) => (
    !pageOverlays.some((overlay) => overlay.type === 'object'
      && overlay.cover
      && overlay.originalRect
      && Math.abs(overlay.originalRect.x - candidate.rect.x) <= 1
      && Math.abs(overlay.originalRect.y - candidate.rect.y) <= 1
      && Math.abs(overlay.originalRect.width - candidate.rect.width) <= 1
      && Math.abs(overlay.originalRect.height - candidate.rect.height) <= 1)
    && !(scanCovered && page && isPageCoveringImage(candidate.rect, page.view))
  )), [imageCandidates, pageOverlays, scanCovered, page])

  useEffect(() => {
    const layer = textLayerRef.current
    const surface = surfaceRef.current
    if (!layer || !surface || !viewport) return
    const frame = window.requestAnimationFrame(() => {
      const pageBounds = surface.getBoundingClientRect()
      const covered = pageOverlays.flatMap((overlay) => (
        overlay.type === 'text' && overlay.cover && overlay.originalRect
          ? [pdfRectToViewport(viewport, overlay.originalRect)]
          : []
      ))
      for (const span of layer.querySelectorAll<HTMLSpanElement>('[data-text-item="true"]')) {
        const bounds = span.getBoundingClientRect()
        const local = {
          left: bounds.left - pageBounds.left,
          top: bounds.top - pageBounds.top,
          width: bounds.width,
          height: bounds.height,
        }
        const hidden = covered.some((rect) => {
          const overlapWidth = Math.max(0, Math.min(local.left + local.width, rect.left + rect.width) - Math.max(local.left, rect.left))
          const overlapHeight = Math.max(0, Math.min(local.top + local.height, rect.top + rect.height) - Math.max(local.top, rect.top))
          const smallerArea = Math.max(1, Math.min(local.width * local.height, rect.width * rect.height))
          return overlapWidth * overlapHeight / smallerArea >= 0.72
        })
        if (hidden) {
          span.dataset.editCovered = 'true'
          span.style.pointerEvents = 'none'
        } else {
          delete span.dataset.editCovered
          span.style.pointerEvents = span.dataset.textWhitespace === 'true' ? 'none' : ''
        }
      }
    })
    return () => window.cancelAnimationFrame(frame)
  }, [pageOverlays, textLayerVersion, viewport])

  // After "Recognize text" from the Edit-mode offer, open the line that was
  // clicked as soon as the page's new text layer is in place.
  useEffect(() => {
    const target = pendingEditAt
    if (!target || target.pageIndex !== pageIndex || !page || !viewport || pageState?.owner !== pdf || canvasReadyViewport !== viewport) return
    const layer = textLayerRef.current
    const surface = surfaceRef.current
    if (!layer || !surface || textLayerPageRef.current !== page || layer.dataset.viewportSignature !== textViewportSignature(viewport)) return
    const frame = window.requestAnimationFrame(() => {
      if (tool === 'edit') {
        const [x, y] = viewport.convertToViewportPoint(target.x, target.y)
        const bounds = surface.getBoundingClientRect()
        const hit = textSpanNear(layer, bounds.left + x, bounds.top + y)
        if (hit) requestTextEditForSpan(hit.span, hit.x, hit.y)
      }
      onPendingEditAtHandled?.(target.token)
    })
    return () => window.cancelAnimationFrame(frame)
    // requestTextEditForSpan reads the render it runs in; the text layer version marks a rebuilt layer.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pendingEditAt, pageIndex, page, viewport, pageState, pdf, canvasReadyViewport, textLayerVersion, tool])

  const localPoint = useCallback((clientX: number, clientY: number) => {
    const bounds = surfaceRef.current?.getBoundingClientRect()
    if (!bounds || !viewport) return null
    return {
      x: clamp(clientX - bounds.left, 0, viewport.width),
      y: clamp(clientY - bounds.top, 0, viewport.height),
    }
  }, [viewport])

  const captureRect = useCallback((rect: PdfRect) => {
    if (!viewport || !canvasRef.current) return undefined
    const box = pdfRectToViewport(viewport, rect)
    const source = canvasRef.current
    const scaleX = source.width / viewport.width
    const scaleY = source.height / viewport.height
    const sx = clamp(Math.floor(box.left * scaleX), 0, source.width - 1)
    const sy = clamp(Math.floor(box.top * scaleY), 0, source.height - 1)
    const sw = clamp(Math.ceil(box.width * scaleX), 1, source.width - sx)
    const sh = clamp(Math.ceil(box.height * scaleY), 1, source.height - sy)
    // The on-screen canvas is rotated by the display rotation. Undo it here so
    // the stored pixels are in unrotated PDF orientation and the saver's
    // axis-aligned draw stays correct.
    const rotation = ((viewport.rotation % 360) + 360) % 360
    const output = document.createElement('canvas')
    output.width = rotation % 180 ? sh : sw
    output.height = rotation % 180 ? sw : sh
    const context = output.getContext('2d')
    if (!context) return undefined
    if (rotation) {
      context.translate(output.width / 2, output.height / 2)
      context.rotate(-rotation * Math.PI / 180)
      context.drawImage(source, sx, sy, sw, sh, -sw / 2, -sh / 2, sw, sh)
    } else {
      context.drawImage(source, sx, sy, sw, sh, 0, 0, sw, sh)
    }
    return output.toDataURL('image/png')
  }, [viewport])

  function sampleCanvasColors(bounds: DOMRect): SampledCanvasColors {
    const fallback: SampledCanvasColors = { text: [0.04, 0.04, 0.05], background: 'rgb(255, 255, 255)' }
    const canvas = canvasRef.current
    const surface = surfaceRef.current
    if (!canvas || !surface || !viewport) return fallback
    const pageBounds = surface.getBoundingClientRect()
    const scaleX = canvas.width / viewport.width
    const scaleY = canvas.height / viewport.height
    const left = clamp(Math.floor((bounds.left - pageBounds.left - 1) * scaleX), 0, canvas.width - 1)
    const top = clamp(Math.floor((bounds.top - pageBounds.top - 1) * scaleY), 0, canvas.height - 1)
    const right = clamp(Math.ceil((bounds.right - pageBounds.left + 1) * scaleX), left + 1, canvas.width)
    const bottom = clamp(Math.ceil((bounds.bottom - pageBounds.top + 1) * scaleY), top + 1, canvas.height)
    const context = canvas.getContext('2d', { willReadFrequently: true })
    if (!context) return fallback

    let pixels: Uint8ClampedArray
    try {
      pixels = context.getImageData(left, top, right - left, bottom - top).data
    } catch {
      return fallback
    }
    if (pixels.length < 4) return fallback

    const buckets = new Map<number, { count: number; r: number; g: number; b: number }>()
    for (let index = 0; index < pixels.length; index += 4) {
      if (pixels[index + 3] < 192) continue
      const r = pixels[index]
      const g = pixels[index + 1]
      const b = pixels[index + 2]
      const key = (Math.round(r / 16) << 16) | (Math.round(g / 16) << 8) | Math.round(b / 16)
      const bucket = buckets.get(key) || { count: 0, r: 0, g: 0, b: 0 }
      bucket.count += 1
      bucket.r += r
      bucket.g += g
      bucket.b += b
      buckets.set(key, bucket)
    }
    const backgroundBucket = Array.from(buckets.values()).sort((a, b) => b.count - a.count)[0]
    if (!backgroundBucket) return fallback
    const background = [
      backgroundBucket.r / backgroundBucket.count,
      backgroundBucket.g / backgroundBucket.count,
      backgroundBucket.b / backgroundBucket.count,
    ]
    let maxDistance = 0
    for (let index = 0; index < pixels.length; index += 4) {
      if (pixels[index + 3] < 192) continue
      maxDistance = Math.max(maxDistance, Math.hypot(
        pixels[index] - background[0],
        pixels[index + 1] - background[1],
        pixels[index + 2] - background[2],
      ))
    }
    if (maxDistance < 18) {
      return {
        ...fallback,
        background: `rgb(${background.map((value) => Math.round(value)).join(', ')})`,
      }
    }

    const threshold = Math.max(20, maxDistance * 0.72)
    const candidates: Array<[number, number, number]> = []
    for (let index = 0; index < pixels.length; index += 4) {
      if (pixels[index + 3] < 192) continue
      const candidate: [number, number, number] = [pixels[index], pixels[index + 1], pixels[index + 2]]
      const distance = Math.hypot(
        candidate[0] - background[0],
        candidate[1] - background[1],
        candidate[2] - background[2],
      )
      if (distance >= threshold) candidates.push(candidate)
    }
    const channelMedian = (channel: 0 | 1 | 2) => {
      const values = candidates.map((candidate) => candidate[channel]).sort((a, b) => a - b)
      return values[Math.floor(values.length / 2)] ?? 10
    }
    return {
      text: [channelMedian(0) / 255, channelMedian(1) / 255, channelMedian(2) / 255],
      background: `rgb(${background.map((value) => Math.round(value)).join(', ')})`,
    }
  }

  function textEditFromSpan(span: HTMLSpanElement, clientX: number, clientY: number): PageTextEdit | null {
    if (!viewport || !surfaceRef.current) return null
    const pageBounds = surfaceRef.current.getBoundingClientRect()
    const selectionSlice = selectedTextSliceWithinSpan(window.getSelection(), span)
    const sourceItemText = span.textContent || ''
    const sourceBounds = span.getBoundingClientRect()
    const bounds = sourceBounds
    const local = {
      left: clamp(bounds.left - pageBounds.left, 0, viewport.width),
      top: clamp(bounds.top - pageBounds.top, 0, viewport.height),
      right: clamp(bounds.right - pageBounds.left, 0, viewport.width),
      bottom: clamp(bounds.bottom - pageBounds.top, 0, viewport.height),
    }
    const sourceLocal = {
      left: clamp(sourceBounds.left - pageBounds.left, 0, viewport.width),
      top: clamp(sourceBounds.top - pageBounds.top, 0, viewport.height),
      right: clamp(sourceBounds.right - pageBounds.left, 0, viewport.width),
      bottom: clamp(sourceBounds.bottom - pageBounds.top, 0, viewport.height),
    }
    const text = sourceItemText
    const rect = viewportRectToPdf(viewport, local)
    const sourceItemRect = viewportRectToPdf(viewport, sourceLocal)
    const computed = window.getComputedStyle(span)
    const family = span.dataset.fontFamily || computed.fontFamily || 'sans-serif'
    const lineHeight = Number.parseFloat(computed.lineHeight)
    const fontHeight = Number.parseFloat(computed.fontSize)
    const letterSpacing = Number.parseFloat(computed.letterSpacing)
    const colors = sampleCanvasColors(bounds)
    if (span.dataset.pdfTextColor) {
      try {
        const color = JSON.parse(span.dataset.pdfTextColor)
        if (Array.isArray(color) && color.length === 3 && color.every(n => Number.isFinite(n) && n >= 0 && n <= 1)) colors.text = color as [number, number, number]
      } catch { /* Keep the sampled fallback for malformed metadata. */ }
    }
    const caret = caretOffsetAtPoint(span, clientX, clientY)
    const pdfBaselineY = Number(span.dataset.pdfBaselineY)
    const pdfTextAngle = Number(span.dataset.pdfTextAngle) || 0
    let inkRect: PdfRect | undefined
    if (Number.isFinite(pdfBaselineY) && Math.abs(pdfTextAngle) < 0.01) {
      const measureCanvas = document.createElement('canvas')
      const context = measureCanvas.getContext('2d')
      if (context) {
        context.font = `${computed.fontStyle} ${computed.fontWeight} ${computed.fontSize} ${computed.fontFamily}`
        const metrics = context.measureText(sourceItemText)
        const ascent = metrics.actualBoundingBoxAscent / zoom
        const descent = metrics.actualBoundingBoxDescent / zoom
        if (Number.isFinite(ascent) && Number.isFinite(descent) && ascent >= 0 && descent >= 0) {
          const y = Math.min(rect.y, pdfBaselineY - descent)
          const top = Math.max(rect.y + rect.height, pdfBaselineY + ascent)
          inkRect = { ...rect, y, height: top - y }
        }
      }
      measureCanvas.width = measureCanvas.height = 1
    }
    const baselineOffset = Number.isFinite(pdfBaselineY) && Math.abs(pdfTextAngle) < 0.01
      ? pdfBaselineY - rect.y
      : undefined
    textVisualStyleRef.current = {
      pageIndex,
      sourceText: text,
      sourceRect: rect,
      fontWeight: span.dataset.fontWeight || computed.fontWeight,
      fontStyle: span.dataset.fontStyle || computed.fontStyle,
      fontStretch: computed.fontStretch,
      lineHeight: (Number.isFinite(lineHeight) ? lineHeight : fontHeight) / zoom,
      letterSpacing: (Number.isFinite(letterSpacing) ? letterSpacing : 0) / zoom,
      scaleX: clamp(Number(span.dataset.textScaleX) || 1, 0.25, 4),
      angle: Number(span.dataset.textAngle) || 0,
      direction: computed.direction === 'rtl' ? 'rtl' : 'ltr',
      backgroundColor: colors.background,
    }
    return {
      pageIndex,
      rect,
      originalRect: rect,
      originalText: text,
      inkRect,
      text,
      fontSize: Number(span.dataset.fontSize) || 11,
      fontFamily: family,
      fontWeight: Number.parseInt(span.dataset.fontWeight || computed.fontWeight, 10) || 400,
      fontStyle: (span.dataset.fontStyle || computed.fontStyle) === 'italic' ? 'italic' : 'normal',
      textFit: 'fit',
      preserveSourceMetrics: true,
      lineHeight: (Number.isFinite(lineHeight) ? lineHeight : fontHeight) / zoom,
      letterSpacing: (Number.isFinite(letterSpacing) ? letterSpacing : 0) / zoom,
      scaleX: clamp(Number(span.dataset.pdfTextScaleX) || 1, 0.25, 4),
      angle: pdfTextAngle,
      direction: computed.direction === 'rtl' ? 'rtl' : 'ltr',
      fontKey: span.dataset.pdfFontName,
      fontData: span.dataset.pdfFontName ? fontDataRef.current.get(span.dataset.pdfFontName) : undefined,
      baselineOffset: Number.isFinite(baselineOffset) ? baselineOffset : undefined,
      sourceSpaceWidth: Number.isFinite(Number(span.dataset.pdfSourceSpaceWidth))
        ? Number(span.dataset.pdfSourceSpaceWidth)
        : undefined,
      // Native text is authored in the PDF's own axes. Page viewing rotation
      // belongs to the viewport, and must not be baked into it a second time.
      displayRotation: 0,
      sourceItemText,
      sourceItemRect,
      sourceSelectionStart: 0,
      sourceSelectionEnd: sourceItemText.length,
      align: computed.direction === 'rtl' ? 'right' : 'left',
      color: colors.text,
      backgroundColor: undefined,
      cover: true,
      modified: false,
      caretOffset: caret,
      selectionStart: selectionSlice?.start ?? caret,
      selectionEnd: selectionSlice?.end ?? caret,
    }
  }

  /**
   * The scanned line an invisible (OCR) text run belongs to. pdf.js splits
   * slanted or widely spaced lines into several runs: neighbouring invisible
   * runs on the same baseline (within 0.35 em, same angle) and less than 1.5
   * em apart are joined, in reading order.
   */
  function scanLineFromSpan(span: HTMLSpanElement): Omit<ScanLineGeometry, 'pageIndex'> | null {
    const layer = textLayerRef.current
    if (!layer) return null
    const read = (item: HTMLSpanElement) => ({
      span: item,
      x: Number(item.dataset.pdfOriginX),
      y: Number(item.dataset.pdfBaselineY),
      angle: Number(item.dataset.pdfTextAngle) || 0,
      width: Number(item.dataset.pdfTextWidth),
      size: Number(item.dataset.fontSize),
      text: item.textContent || '',
    })
    const clicked = read(span)
    if (![clicked.x, clicked.y, clicked.width, clicked.size].every(Number.isFinite) || clicked.size <= 0) return null
    const dir = { x: Math.cos(clicked.angle), y: Math.sin(clicked.angle) }
    const up = { x: -dir.y, y: dir.x }
    const along = (item: { x: number; y: number }) => (item.x - clicked.x) * dir.x + (item.y - clicked.y) * dir.y
    const across = (item: { x: number; y: number }) => (item.x - clicked.x) * up.x + (item.y - clicked.y) * up.y
    const runs = [...layer.querySelectorAll<HTMLSpanElement>('[data-pdf-text-invisible="true"]')]
      .map(read)
      .filter((item) => [item.x, item.y, item.width, item.size].every(Number.isFinite)
        && Math.abs(item.angle - clicked.angle) <= Math.PI / 180
        && Math.abs(across(item)) <= 0.35 * clicked.size)
      .sort((a, b) => along(a) - along(b))
    let first = runs.findIndex((item) => item.span === span)
    if (first < 0) return null
    let last = first
    const gap = (a: (typeof runs)[number], b: (typeof runs)[number]) => along(b) - (along(a) + a.width)
    while (first > 0 && gap(runs[first - 1], runs[first]) < 1.5 * clicked.size) first -= 1
    while (last + 1 < runs.length && gap(runs[last], runs[last + 1]) < 1.5 * clicked.size) last += 1
    const line = runs.slice(first, last + 1)
    let text = ''
    line.forEach((item, index) => {
      if (index && !/\s$/u.test(text) && !/^\s/u.test(item.text) && gap(line[index - 1], item) > 0.1 * clicked.size) text += ' '
      text += item.text
    })
    text = text.replace(/\s+/gu, ' ').trim()
    if (!text) return null
    const start = line[0]
    const end = line[line.length - 1]
    const sizes = line.map((item) => item.size).sort((a, b) => a - b)
    return {
      text,
      origin: { x: start.x, y: start.y },
      dir,
      length: Math.max(1, along(end) + end.width - along(start)),
      fontSize: sizes[sizes.length >> 1],
    }
  }

  /** Paper and ink colours around a box of the page as rendered. */
  function sampleBoxColors(box: { left: number; top: number; width: number; height: number }) {
    const bounds = surfaceRef.current?.getBoundingClientRect()
    if (!bounds) return { text: [0.04, 0.04, 0.05] as [number, number, number], background: 'rgb(255, 255, 255)' }
    return sampleCanvasColors(new DOMRect(bounds.left + box.left, bounds.top + box.top, box.width, box.height))
  }

  /**
   * Edit the scanned line under an invisible text run (design 4.8.2): the
   * whole line opens in a matched font over a retouch patch that hides its
   * printed words. A line edited before reopens as edited.
   */
  function requestScanEditForSpan(span: HTMLSpanElement, clientX: number, clientY: number) {
    if (!viewport || !page) return
    const line = scanLineFromSpan(span)
    if (!line) return
    const committed = pageOverlays.find((overlay): overlay is TextOverlay => overlay.type === 'text'
      && Boolean(overlay.scan) && isSameScanLine(overlay.scan!, pageIndex, overlay.pageIndex, line.origin, line.dir))
    textVisualStyleRef.current = null
    if (committed) {
      onRequestTextEdit(scanEditFromOverlay(committed))
      return
    }
    // Exact word boxes when this page was recognised in this window.
    const recognised = findOcrLine(ocrResultForPage(pdf, pageIndex), line.origin, line.dir, line.fontSize)
    const recognisedWords = recognised?.line.words.map((word) => ({ text: word.text, quad: word.quad }))
    const words = recognisedWords ? scanWordsFromOcr(recognisedWords, line.text, line.origin, line.dir) : undefined
    const geometry: ScanLineGeometry = {
      ...line,
      pageIndex,
      words: words ? recognisedWords : undefined,
      lineSpacing: recognised?.lineSpacing,
    }
    const key = scanPreparationKey(pdf, scanLineKey(pageIndex, line.origin, line.dir))
    const run = { origin: line.origin, dir: line.dir, length: line.length, fontSize: line.fontSize }
    const colors = sampleBoxColors(pdfRectToViewport(viewport, quadBounds(runQuad(run))))
    // The caret: along the baseline to the word clicked, then within it.
    let caretOffset = line.text.length
    const local = localPoint(clientX, clientY)
    const knownWords = words ?? scanPreparationResult(key)?.words
    if (local) {
      const [x, y] = viewport.convertToPdfPoint(local.x, local.y)
      const u = (x - line.origin.x) * line.dir.x + (y - line.origin.y) * line.dir.y
      const tokens = line.text.split(' ')
      if (knownWords?.length === tokens.length) {
        let offset = 0
        let best = { distance: Infinity, offset: line.text.length }
        knownWords.forEach((word, index) => {
          const width = Math.max(1e-6, word.end - word.start)
          const share = Math.max(0, Math.min(1, (u - word.start) / width))
          const distance = u < word.start ? word.start - u : u > word.end ? u - word.end : 0
          if (distance < best.distance) best = { distance, offset: offset + Math.round(share * tokens[index].length) }
          offset += tokens[index].length + 1
        })
        caretOffset = best.offset
      } else {
        caretOffset = Math.round(Math.max(0, Math.min(1, u / line.length)) * line.text.length)
      }
    }
    const edit = createScanEdit({
      pageIndex,
      key,
      line: geometry,
      words,
      color: colors.text,
      paper: cssColorToPdf(colors.background),
      caretOffset,
    })
    onRequestTextEdit(edit)
    if (edit.scan?.status !== 'ready') {
      void prepareScanEdit(page, key, geometry, { nativeDpi: pageScan?.nativeDpi }).then((result) => onScanEditPrepared?.(key, result))
    }
  }

  /** Start editing the text run at this point (Edit mode). */
  function requestTextEditForSpan(span: HTMLSpanElement, clientX: number, clientY: number) {
    // Invisible text is recognised text over a scan: edit the scanned words,
    // never the invisible glyphs alone (they would leave the printed words
    // visible under a box-glyph font).
    if (span.dataset.pdfTextInvisible === 'true') {
      requestScanEditForSpan(span, clientX, clientY)
      return
    }
    const edit = textEditFromSpan(span, clientX, clientY)
    if (edit) onRequestTextEdit(edit)
  }

  function handleClick(event: React.MouseEvent) {
    const press = offerPressRef.current
    offerPressRef.current = null
    if ((event.target as HTMLElement).closest('[data-edit-ui="true"]')) return
    if (tool === 'edit') {
      const span = (event.target as HTMLElement).closest<HTMLSpanElement>('[data-text-item="true"]')
      if (span) {
        event.preventDefault()
        event.stopPropagation()
        requestTextEditForSpan(span, event.clientX, event.clientY)
        return
      }
      // A plain click (not the end of a drag) on a scan that has no text yet.
      if (pageScan?.kind === 'scan' && onRequestOcrOffer && viewport && !selectingObjectRegion && press
        && Math.hypot(event.clientX - press.x, event.clientY - press.y) < 5
        && !(event.target as HTMLElement).closest('button, input, textarea, select, a')) {
        const point = localPoint(event.clientX, event.clientY)
        if (point) {
          const [x, y] = viewport.convertToPdfPoint(point.x, point.y)
          onRequestOcrOffer(pageIndex, { x, y }, { x: event.clientX, y: event.clientY })
        }
      }
      return
    }
    // Markup tools operate on a real DOM Range in handleMouseUp. A click with
    // no range intentionally does nothing; a PDF.js text item may represent a
    // complete line and must never be mistaken for a focused selection.
    if (tool === 'highlight' || tool === 'underline' || tool === 'strikeout') return
    if (tool === 'sign') {
      if (!viewport || (event.target as HTMLElement).closest('button, input, textarea, select')) return
      const point = localPoint(event.clientX, event.clientY)
      if (!point) return
      const [x, y] = viewport.convertToPdfPoint(point.x, point.y)
      onPlaceSignature({ x, y }, displayRotation)
      return
    }
    if (tool !== 'addText' || !viewport || !surfaceRef.current) return
    if ((event.target as HTMLElement).closest('button, input, textarea, select')) return
    const point = localPoint(event.clientX, event.clientY)
    if (!point) return
    const left = clamp(point.x, 0, Math.max(0, viewport.width - 80))
    const top = clamp(point.y, 0, Math.max(0, viewport.height - 28))
    const width = Math.min(260, viewport.width - left)
    const height = Math.min(44, viewport.height - top)
    const local = {
      left,
      top,
      right: left + Math.max(1, width),
      bottom: top + Math.max(1, height),
    }
    const rect = viewportRectToPdf(viewport, local)
    textVisualStyleRef.current = null
    onRequestTextEdit({
      pageIndex,
      rect,
      originalText: '',
      text: '',
      fontSize: 12,
      fontFamily: 'Segoe UI',
      displayRotation,
      align: 'left',
      color: [0.04, 0.04, 0.05],
      cover: false,
      modified: true,
      caretOffset: 0,
      selectionStart: 0,
      selectionEnd: 0,
    })
  }

  function commitTextMarkupSelection() {
    if (!['highlight', 'underline', 'strikeout'].includes(tool) || !viewport || !surfaceRef.current || !textLayerRef.current) return
    window.setTimeout(() => {
      const selection = window.getSelection()
      const surface = surfaceRef.current
      const textLayer = textLayerRef.current
      if (!selection || selection.isCollapsed || !surface || !textLayer) return
      const bounds = surface.getBoundingClientRect()
      const rects = selectionClientRectsWithin(selection, textLayer)
        .map((rect) => ({
          left: clamp(rect.left - bounds.left, 0, viewport.width),
          top: clamp(rect.top - bounds.top, 0, viewport.height),
          right: clamp(rect.right - bounds.left, 0, viewport.width),
          bottom: clamp(rect.bottom - bounds.top, 0, viewport.height),
        }))
        .filter((rect) => rect.right - rect.left > 2 && rect.bottom - rect.top > 2)
        .map((rect) => viewportRectToPdf(viewport, rect))
      if (rects.length) {
        if (tool === 'highlight') onHighlight(rects)
        else onTextMarkup(tool as 'underline' | 'strikeout', rects, displayRotation)
        selection.removeAllRanges()
      }
    }, 0)
  }

  function handlePointerDown(event: ReactPointerEvent<HTMLDivElement>) {
    offerPressRef.current = null
    if (!viewport) return
    const target = event.target as HTMLElement
    if (target.closest('[data-edit-ui="true"],button,input,textarea,select')) return
    if (tool === 'edit') {
      if (!selectingObjectRegion && event.button === 0) offerPressRef.current = { x: event.clientX, y: event.clientY }
      if (selectingObjectRegion) {
        const point = localPoint(event.clientX, event.clientY)
        if (!point) return
        event.currentTarget.setPointerCapture(event.pointerId)
        setRegionOrigin(point)
        setRegionDraft({ left: point.x, top: point.y, width: 0, height: 0 })
        return
      }
      // handleClick -> beginTextEdit owns a text-to-text transition. Committing
      // here too can materialize the previous native edit twice.
      if (target.closest('[data-text-item="true"]')) return
      if (activeText) onCommitTextEdit()
      if (activeObject) onCommitObjectEdit()
      return
    }
    if (tool === 'draw') {
      const point = localPoint(event.clientX, event.clientY)
      if (!point) return
      event.currentTarget.setPointerCapture(event.pointerId)
      setInkDraft([point])
      return
    }
    if (tool === 'rectangle') {
      const point = localPoint(event.clientX, event.clientY)
      if (!point) return
      event.currentTarget.setPointerCapture(event.pointerId)
      setShapeOrigin(point)
      setShapeDraft({ left: point.x, top: point.y, width: 0, height: 0 })
      return
    }
    if (tool !== 'crop') return
    const point = localPoint(event.clientX, event.clientY)
    if (!point) return
    event.currentTarget.setPointerCapture(event.pointerId)
    setCropOrigin(point)
    setCropDraft({ left: point.x, top: point.y, width: 0, height: 0 })
  }

  function handlePointerMove(event: ReactPointerEvent<HTMLDivElement>) {
    if (tool === 'edit' && regionOrigin) {
      const point = localPoint(event.clientX, event.clientY)
      if (!point) return
      setRegionDraft({
        left: Math.min(regionOrigin.x, point.x),
        top: Math.min(regionOrigin.y, point.y),
        width: Math.abs(point.x - regionOrigin.x),
        height: Math.abs(point.y - regionOrigin.y),
      })
      return
    }
    if (tool === 'draw' && inkDraft.length) {
      const point = localPoint(event.clientX, event.clientY)
      if (!point) return
      const previous = inkDraft.at(-1)
      if (!previous || Math.hypot(point.x - previous.x, point.y - previous.y) >= 1.5) {
        setInkDraft((current) => [...current, point])
      }
      return
    }
    if (tool === 'rectangle' && shapeOrigin) {
      const point = localPoint(event.clientX, event.clientY)
      if (!point) return
      setShapeDraft({
        left: Math.min(shapeOrigin.x, point.x),
        top: Math.min(shapeOrigin.y, point.y),
        width: Math.abs(point.x - shapeOrigin.x),
        height: Math.abs(point.y - shapeOrigin.y),
      })
      return
    }
    if (tool !== 'crop' || !cropOrigin) return
    const point = localPoint(event.clientX, event.clientY)
    if (!point) return
    setCropDraft({
      left: Math.min(cropOrigin.x, point.x),
      top: Math.min(cropOrigin.y, point.y),
      width: Math.abs(point.x - cropOrigin.x),
      height: Math.abs(point.y - cropOrigin.y),
    })
  }

  function handlePointerUp(event: ReactPointerEvent<HTMLDivElement>) {
    if (tool === 'highlight' || tool === 'underline' || tool === 'strikeout') {
      commitTextMarkupSelection()
      return
    }
    if (tool === 'edit' && regionOrigin) {
      if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
      const draft = regionDraft
      setRegionOrigin(null)
      setRegionDraft(null)
      onObjectRegionSelected()
      if (!draft || draft.width < 16 || draft.height < 16 || !viewport) return
      const rect = viewportRectToPdf(viewport, {
        left: draft.left,
        top: draft.top,
        right: draft.left + draft.width,
        bottom: draft.top + draft.height,
      })
      onRequestObjectEdit({
        pageIndex,
        kind: 'artwork',
        rect,
        originalRect: rect,
        dataUrl: captureRect(rect),
        opacity: 1,
        cover: true,
        displayRotation,
        label: 'Artwork region',
        modified: false,
      })
      return
    }
    if (tool === 'draw' && inkDraft.length) {
      if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
      const points = inkDraft
      setInkDraft([])
      if (points.length < 2 || !viewport) return
      onInk(points.map((point) => {
        const [x, y] = viewport.convertToPdfPoint(point.x, point.y)
        return { x, y }
      }))
      return
    }
    if (tool === 'rectangle' && shapeOrigin) {
      if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
      const draft = shapeDraft
      setShapeOrigin(null)
      setShapeDraft(null)
      if (!draft || draft.width < 4 || draft.height < 4 || !viewport) return
      onRectangle(viewportRectToPdf(viewport, {
        left: draft.left,
        top: draft.top,
        right: draft.left + draft.width,
        bottom: draft.top + draft.height,
      }))
      return
    }
    if (tool !== 'crop' || !cropOrigin) return
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
    setCropOrigin(null)
    setCropDraft((draft) => draft && draft.width >= 16 && draft.height >= 16 ? draft : null)
  }

  function startTransform(event: ReactPointerEvent<HTMLElement>, target: 'text' | 'object', mode: 'move' | 'resize', handle?: ResizeHandle) {
    if (!viewport) return
    const editRect = target === 'text' ? activeText?.rect : activeObject?.rect
    const point = localPoint(event.clientX, event.clientY)
    if (!editRect || !point) return
    event.preventDefault()
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)
    transformRef.current = {
      target,
      mode,
      handle,
      startPoint: point,
      startRect: pdfRectToViewport(viewport, editRect),
    }
  }

  function updateTransform(event: ReactPointerEvent<HTMLElement>) {
    const gesture = transformRef.current
    const point = localPoint(event.clientX, event.clientY)
    if (!gesture || !point || !viewport) return
    event.preventDefault()
    event.stopPropagation()
    const dx = point.x - gesture.startPoint.x
    const dy = point.y - gesture.startPoint.y
    let left = gesture.startRect.left
    let top = gesture.startRect.top
    let right = left + gesture.startRect.width
    let bottom = top + gesture.startRect.height
    if (gesture.mode === 'move') {
      const width = right - left
      const height = bottom - top
      left = clamp(left + dx, 0, viewport.width - width)
      top = clamp(top + dy, 0, viewport.height - height)
      right = left + width
      bottom = top + height
    } else {
      const handle = gesture.handle || 'se'
      if (handle.includes('w')) left = clamp(left + dx, 0, right - 24)
      if (handle.includes('e')) right = clamp(right + dx, left + 24, viewport.width)
      if (handle.includes('n')) top = clamp(top + dy, 0, bottom - 18)
      if (handle.includes('s')) bottom = clamp(bottom + dy, top + 18, viewport.height)
    }
    const rect = viewportRectToPdf(viewport, { left, top, right, bottom })
    // A native text run keeps its baseline as an offset from the box bottom;
    // resizing must keep the text where the editor shows it (top-anchored),
    // or the saved baseline would follow the dragged bottom edge.
    if (gesture.target === 'text' && activeText) onTextEditChange({ ...resizeTextEditRect(activeText, rect), modified: true })
    if (gesture.target === 'object' && activeObject) onObjectEditChange({ ...activeObject, rect, modified: true })
  }

  function endTransform(event: ReactPointerEvent<HTMLElement>) {
    if (!transformRef.current) return
    event.preventDefault()
    event.stopPropagation()
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
    transformRef.current = null
  }

  function transformHandles(target: 'text' | 'object') {
    return RESIZE_HANDLES.map((handle) => (
      <span
        key={handle}
        role="presentation"
        className={`selection-handle handle-${handle}`}
        onPointerDown={(event) => startTransform(event, target, 'resize', handle)}
        onPointerMove={updateTransform}
        onPointerUp={endTransform}
      />
    ))
  }

  async function followAnnotation(annotation: any) {
    if (annotation.url) {
      await window.simple.openExternal(annotation.url)
      return
    }
    if (!annotation.dest) return
    try {
      const destination = typeof annotation.dest === 'string' ? await pdf.getDestination(annotation.dest) : annotation.dest
      if (!destination?.[0]) return
      const target = typeof destination[0] === 'object' ? await pdf.getPageIndex(destination[0]) : Number(destination[0])
      if (Number.isInteger(target)) onNavigate(target)
    } catch {
      // Ignore malformed destinations.
    }
  }

  /**
   * An editable control for a form widget: text (multi-line fields keep their
   * line breaks), check box, radio button, dropdown (editable or not) and
   * list box. Values are keyed by the field name pdf.js reports; the main
   * process writes them and reports anything it cannot.
   */
  function renderFormField(annotation: PdfWidgetAnnotation, box: ViewerRect) {
    const name = annotation.fieldName
    const value = formValues[name]
    const readOnly = Boolean(annotation.readOnly)
    const flags = Number(annotation.fieldFlags) || 0
    const label = annotation.alternativeText || name
    const common = {
      'data-form-field': name,
      'aria-label': label,
      title: annotation.alternativeText || undefined,
    }
    const fontSize = formFieldFontSize(annotation, box.height, zoom)
    if (annotation.fieldType === 'Tx') {
      const text = formTextValue(value, annotation.fieldValue).replace(/\r\n?/g, '\n')
      const maxLength = Number(annotation.maxLen) > 0 ? Number(annotation.maxLen) : undefined
      const style: CSSProperties = {
        ...box,
        fontSize,
        textAlign: TEXT_ALIGNMENTS[Number(annotation.textAlignment) || 0] ?? 'left',
      }
      const spellCheck = !(flags & FIELD_FLAG_DO_NOT_SPELL_CHECK)
      if (annotation.multiLine) {
        return (
          <textarea
            key={annotation.id}
            {...common}
            className="pdf-form-field is-multiline"
            style={{ ...style, resize: 'none', lineHeight: 1.15, overflow: 'auto' }}
            value={text}
            maxLength={maxLength}
            readOnly={readOnly}
            spellCheck={spellCheck}
            onChange={(event) => onFormChange(name, event.target.value)}
          />
        )
      }
      return (
        <input
          key={annotation.id}
          {...common}
          type={flags & FIELD_FLAG_PASSWORD ? 'password' : 'text'}
          className="pdf-form-field"
          style={style}
          value={text}
          maxLength={maxLength}
          readOnly={readOnly}
          spellCheck={spellCheck}
          onChange={(event) => onFormChange(name, event.target.value)}
        />
      )
    }
    if (annotation.fieldType === 'Btn' && annotation.checkBox) {
      return (
        <input
          key={annotation.id}
          {...common}
          type="checkbox"
          className="pdf-form-checkbox"
          style={box}
          checked={checkBoxIsOn(annotation, value)}
          disabled={readOnly}
          // This widget's own on-state, so a field whose widgets differ
          // checks only the one that was clicked.
          onChange={(event) => onFormChange(name, event.target.checked ? (annotation.exportValue || true) : false)}
        />
      )
    }
    if (annotation.fieldType === 'Btn' && annotation.radioButton && annotation.buttonValue) {
      const buttonValue = annotation.buttonValue
      const selected = value === undefined ? annotation.fieldValue : value
      return (
        <input
          key={annotation.id}
          {...common}
          type="radio"
          name={`pdf-field:${name}`}
          className="pdf-form-checkbox"
          style={box}
          checked={selected === buttonValue}
          disabled={readOnly}
          onChange={(event) => { if (event.target.checked) onFormChange(name, buttonValue) }}
        />
      )
    }
    if (annotation.fieldType === 'Ch') {
      const options = Array.isArray(annotation.options) ? annotation.options : []
      const selection = choiceSelection(value, annotation.fieldValue)
      const style: CSSProperties = { ...box, fontSize, fontFamily: 'inherit', padding: '0 2px' }
      const optionElements = options.map((option, index) => (
        <option key={`${index}:${option.exportValue}`} value={option.exportValue}>{option.displayValue || option.exportValue}</option>
      ))
      if (annotation.combo) {
        const current = selection[0] ?? ''
        if (flags & FIELD_FLAG_EDITABLE_COMBO) {
          const listId = `pdf-choices-${annotation.id}`
          return (
            <span key={annotation.id} style={{ display: 'contents' }}>
              <input
                {...common}
                className="pdf-form-field"
                style={{ ...style, padding: '1px 4px' }}
                list={listId}
                value={current}
                readOnly={readOnly}
                onChange={(event) => onFormChange(name, event.target.value)}
              />
              <datalist id={listId}>{optionElements}</datalist>
            </span>
          )
        }
        return (
          <select
            key={annotation.id}
            {...common}
            className="pdf-form-field"
            style={style}
            value={current}
            disabled={readOnly}
            onChange={(event) => onFormChange(name, event.target.value)}
          >
            {!options.some((option) => option.exportValue === current) && <option value={current}>{current}</option>}
            {optionElements}
          </select>
        )
      }
      return (
        <select
          key={annotation.id}
          {...common}
          className="pdf-form-field is-list"
          style={style}
          multiple={Boolean(annotation.multiSelect)}
          size={Math.max(2, options.length)}
          value={annotation.multiSelect ? selection : (selection[0] ?? '')}
          disabled={readOnly}
          onChange={(event) => onFormChange(name, annotation.multiSelect
            ? Array.from(event.target.selectedOptions, (option) => option.value)
            : event.target.value)}
        >
          {optionElements}
        </select>
      )
    }
    return null
  }

  if (!viewport) {
    return <div className="page-loading"><LoaderCircle className="spin" size={20} /><span>Preparing page {pageIndex + 1}…</span></div>
  }

  // Object pixels are stored in unrotated PDF orientation (captureRect and the
  // signature/image placement normalize them), so on a rotated page the raw
  // <img> must be turned by the current display rotation to match the canvas
  // and the axis-aligned draw the saver performs.
  const objectImageStyle = (box: { width: number; height: number }): CSSProperties | undefined => {
    if (!displayRotation) return undefined
    if (displayRotation === 180) return { transform: 'rotate(180deg)' }
    return {
      position: 'absolute',
      left: '50%',
      top: '50%',
      width: box.height,
      height: box.width,
      transform: `translate(-50%, -50%) rotate(${displayRotation}deg)`,
    }
  }
  /**
   * A retouch patch (pixels in unrotated PDF orientation, like object
   * images), or, while it is not ready or could not be made, the paper
   * colour over the old words.
   */
  const renderScanPatch = (key: string, patch: ScanPatch | undefined, cover: PdfRect | undefined, paper: [number, number, number] | undefined, zIndex: number) => {
    if (patch) {
      const box = pdfRectToViewport(viewport, patch.rect)
      return (
        <span key={key} className="scan-patch" data-scan-patch="true" aria-hidden="true" style={{ position: 'absolute', ...box, zIndex, overflow: 'hidden', pointerEvents: 'none' }}>
          <img src={patch.dataUrl} alt="" draggable={false} style={{ display: 'block', width: '100%', height: '100%', ...objectImageStyle(box) }} />
        </span>
      )
    }
    if (!cover) return null
    return <span key={key} aria-hidden="true" style={{ position: 'absolute', ...pdfRectToViewport(viewport, cover), zIndex, pointerEvents: 'none', background: colorCss(paper ?? [1, 1, 1]) }} />
  }
  const activeScan = activeText?.scan
  const rememberedTextVisual = textVisualStyleRef.current
  const activeTextVisual = activeText
    && rememberedTextVisual?.pageIndex === pageIndex
    && rememberedTextVisual.sourceText === activeText.originalText
    && sameRect(rememberedTextVisual.sourceRect, activeText.originalRect || activeText.rect)
    ? rememberedTextVisual
    : null
  const editorScaleX = activeTextLayout?.scaleX || 1

  return (
    <div className="page-stage" style={{ minHeight: 'auto' }}>
      <div
        ref={surfaceRef}
        className={cx('page-surface', `tool-${tool}`, selectingObjectRegion && 'is-selecting-object', !pageVisualReady && 'is-page-transitioning')}
        style={{ width: displayedViewport?.width || viewport.width, height: displayedViewport?.height || viewport.height }}
        aria-busy={!pageVisualReady}
        onClick={handleClick}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={handlePointerUp}
      >
        <canvas ref={canvasRef} className="page-canvas" />
        <div
          ref={textLayerRef}
          className="text-layer"
          aria-label={`Selectable text on page ${pageIndex + 1}`}
          style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
        />
        {activeSearchMatch?.pageIndex === pageIndex
          && searchMatchGeometry.key === activeSearchKey
          && searchMatchGeometry.rects.length > 0 && (
          <div
            ref={searchMatchLayerRef}
            className="search-match-layer"
            aria-hidden="true"
            data-search-highlight="true"
            data-search-active="true"
            data-page-index={pageIndex}
            data-occurrence-index={activeSearchMatch.occurrenceIndex}
            data-search-query={activeSearchMatch.query}
          >
            {searchMatchGeometry.rects.map((rect, index) => (
              <span
                key={`${index}:${rect.left}:${rect.top}`}
                className="search-match-rect"
                data-search-match-rect="true"
                style={rect}
              />
            ))}
          </div>
        )}
        <div className="annotation-layer">
          {annotations.map((annotation) => {
            if (!annotation.rect) return null
            const converted = viewport.convertToViewportRectangle(annotation.rect)
            const box = {
              left: Math.min(converted[0], converted[2]),
              top: Math.min(converted[1], converted[3]),
              width: Math.abs(converted[2] - converted[0]),
              height: Math.abs(converted[3] - converted[1]),
            }
            if (annotation.subtype === 'Link') {
              return <button key={annotation.id} type="button" className="pdf-link" style={box} title={annotation.url || 'Go to linked page'} onClick={(event) => { event.stopPropagation(); followAnnotation(annotation) }}><ExternalLink size={10} /></button>
            }
            if (annotation.subtype === 'Widget' && annotation.fieldName && !annotation.hidden) {
              return renderFormField(annotation as PdfWidgetAnnotation, box)
            }
            return null
          })}
        </div>

        <div className="overlay-layer">
          {/* Retouch patches first, so no patch can cover another edit's new text. */}
          {pageOverlays.map((overlay) => (overlay.type === 'text' && overlay.scan?.mode === 'appearance' && activeText?.overlayId !== overlay.id
            // Ready without a patch: there was no printed ink to hide.
            ? renderScanPatch(`scan-patch-${overlay.id}`, scanOverlayDrawing(overlay).patch, overlay.scan.status === 'ready' ? undefined : overlay.inkRect || overlay.originalRect, overlay.scan.paper, 0)
            : null))}
          {pageOverlays.map((overlay) => {
            if (overlay.type === 'ink') {
              const points = overlay.points.map((point) => viewport.convertToViewportPoint(point.x, point.y).join(',')).join(' ')
              return (
                <svg key={overlay.id} className="ink-overlay" viewBox={`0 0 ${viewport.width} ${viewport.height}`} aria-hidden="true">
                  <polyline points={points} fill="none" stroke={colorCss(overlay.color)} strokeWidth={overlay.thickness * zoom} strokeOpacity={overlay.opacity} strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              )
            }
            const box = pdfRectToViewport(viewport, overlay.rect)
            if (overlay.type === 'highlight') {
              return <div key={overlay.id} className="highlight-overlay" style={box} title="Saved highlight" />
            }
            if (overlay.type === 'markup') {
              if (overlay.style === 'rectangle') {
                return <div key={overlay.id} className="rectangle-overlay" style={{ ...box, borderColor: colorCss(overlay.color), borderWidth: overlay.thickness * zoom, opacity: overlay.opacity }} title="Saved rectangle" />
              }
              // Keep the bar under (or through) the marked text as it was seen
              // when created, even after the page is rotated on screen.
              const relativeRotation = ((((overlay.displayRotation || 0) - displayRotation) % 360) + 360) % 360
              const barThickness = Math.max(1, overlay.thickness * zoom)
              const barStyle: CSSProperties = relativeRotation % 180
                ? {
                    background: colorCss(overlay.color),
                    width: barThickness,
                    height: '100%',
                    top: 0,
                    ...(overlay.style === 'strikeout'
                      ? { left: '50%', transform: 'translateX(-50%)' }
                      : relativeRotation === 90
                        ? { left: 'auto', right: 1 }
                        : { left: 1 }),
                  }
                : {
                    background: colorCss(overlay.color),
                    height: barThickness,
                    ...(relativeRotation === 180 && overlay.style === 'underline' ? { top: 1, bottom: 'auto' } : {}),
                  }
              return (
                <div key={overlay.id} className={`text-markup-overlay markup-${overlay.style}`} style={{ ...box, opacity: overlay.opacity }} title={overlay.style === 'underline' ? 'Saved underline' : 'Saved strikeout'}>
                  <span style={barStyle} />
                </div>
              )
            }
            if (overlay.type === 'object') {
              if (activeObject?.overlayId === overlay.id) return null
              const originalBox = overlay.cover && overlay.originalRect ? pdfRectToViewport(viewport, overlay.originalRect) : null
              return (
                <div key={overlay.id} className="committed-object-group">
                  {originalBox && <span className="object-original-cover" style={originalBox} />}
                  {overlay.dataUrl && (
                    <button
                      type="button"
                      data-edit-ui="true"
                      className={cx('object-overlay', tool === 'edit' && 'is-editable')}
                      style={{ ...box, opacity: overlay.opacity }}
                      onClick={(event) => {
                        event.stopPropagation()
                        if (tool !== 'edit') return
                        onRequestObjectEdit({
                          overlayId: overlay.id,
                          pageIndex,
                          kind: overlay.kind,
                          rect: overlay.rect,
                          originalRect: overlay.originalRect,
                          dataUrl: overlay.dataUrl,
                          opacity: overlay.opacity,
                          cover: overlay.cover,
                          displayRotation: overlay.displayRotation,
                          label: overlay.kind === 'image' ? 'Image' : 'Artwork region',
                          modified: false,
                        })
                      }}
                    ><img src={overlay.dataUrl} alt="Edited PDF object" draggable={false} style={objectImageStyle(box)} /></button>
                  )}
                </div>
              )
            }
            if (activeText?.overlayId === overlay.id) return null
            // An edit of scanned text draws only the words it replaces (the
            // rest of the line keeps its scanned pixels); a correction of the
            // recognised text alone draws nothing.
            const scanDrawing = overlay.scan ? scanOverlayDrawing(overlay) : null
            const shown: TextOverlay = scanDrawing
              ? { ...overlay, text: scanDrawing.text, rect: scanDrawing.rect, baselineOffset: scanDrawing.baselineOffset, originalText: scanDrawing.originalText }
              : overlay
            const textBox = scanDrawing ? pdfRectToViewport(viewport, scanDrawing.rect) : box
            const invisibleText = overlay.scan?.mode === 'recognized-text'
            const originalBox = overlay.cover && overlay.originalRect ? pdfRectToViewport(viewport, overlay.inkRect || overlay.originalRect) : null
            const sourceDisplayRotation = overlay.displayRotation ?? 0
            const displayDelta = (((displayRotation - sourceDisplayRotation) % 360) + 360) % 360
            const sidewaysText = overlay.scan ? scanFrameSideways(overlay) : displayDelta === 90 || displayDelta === 270
            const screenAngle = displayDelta * Math.PI / 180 - (overlay.angle || 0)
            const preview = previewTextLayout(shown, (sidewaysText ? textBox.height : textBox.width) / zoom)
            const textScaleX = preview.scaleX
            const reopen = (event: React.MouseEvent) => {
              event.stopPropagation()
              if (tool !== 'edit') return
              textVisualStyleRef.current = null
              if (overlay.scan) {
                onRequestTextEdit(scanEditFromOverlay(overlay))
                return
              }
              onRequestTextEdit({
                overlayId: overlay.id,
                pageIndex,
                rect: overlay.rect,
                originalRect: overlay.originalRect,
                originalText: overlay.originalText ?? overlay.text,
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
                fontKey: overlay.fontKey,
                fontData: overlay.fontData,
                baselineOffset: overlay.baselineOffset,
                sourceSpaceWidth: overlay.sourceSpaceWidth,
                displayRotation: overlay.displayRotation,
                align: overlay.align || 'left',
                color: overlay.color,
                backgroundColor: overlay.backgroundColor,
                cover: overlay.cover,
                modified: false,
                caretOffset: overlay.text.length,
                selectionStart: overlay.text.length,
                selectionEnd: overlay.text.length,
              })
            }
            return (
              <div key={overlay.id} className="committed-text-group">
                {overlay.scan && tool === 'edit' && (scanDrawing?.rect !== overlay.rect || invisibleText) && (
                  // The whole edited line stays one click target, also where
                  // only some of its words were replaced (or none visibly).
                  <button
                    type="button"
                    data-edit-ui="true"
                    data-scan-line="true"
                    className="text-overlay is-editable"
                    title="Edit scanned text"
                    aria-label="Edit scanned text"
                    style={{ ...box, backgroundColor: 'transparent', color: 'transparent' }}
                    onClick={reopen}
                  />
                )}
                <button
                  type="button"
                  data-edit-ui="true"
                  className={cx('text-overlay', overlay.cover && 'is-replacement', tool === 'edit' && 'is-editable')}
                  aria-hidden={invisibleText || undefined}
                  tabIndex={invisibleText ? -1 : undefined}
                  style={{
                    left: textBox.left + textBox.width / 2,
                    top: textBox.top + textBox.height / 2,
                    width: (sidewaysText ? textBox.height : textBox.width) / textScaleX,
                    height: sidewaysText ? textBox.width : textBox.height,
                    fontSize: overlay.fontSize * zoom,
                    fontFamily: overlay.fontFamily,
                    fontWeight: overlay.fontWeight,
                    fontStyle: overlay.fontStyle,
                    lineHeight: `${preview.lineHeight * zoom}px`,
                    letterSpacing: overlay.letterSpacing ? `${overlay.letterSpacing * zoom}px` : undefined,
                    direction: overlay.direction,
                    textAlign: overlay.align || 'left',
                    color: invisibleText ? 'transparent' : colorCss(overlay.color),
                    backgroundColor: originalBox ? 'transparent' : overlay.backgroundColor ? colorCss(overlay.backgroundColor) : undefined,
                    pointerEvents: invisibleText ? 'none' : undefined,
                    transform: `translate(-50%, -50%) rotate(${screenAngle}rad) scaleX(${textScaleX})`,
                    transformOrigin: 'center',
                    whiteSpace: 'pre',
                    // A block <button> centres its content vertically; the
                    // editor and the saved PDF start the text at the top of
                    // the box. A flex column keeps it there.
                    display: 'flex',
                    flexDirection: 'column',
                    justifyContent: 'flex-start',
                    alignItems: 'stretch',
                  }}
                  onClick={reopen}
                ><span className="text-overlay-lines" style={{ display: 'block', whiteSpace: 'pre' }}>{preview.lines.join('\n')}</span></button>
              </div>
            )
          })}
        </div>

        {tool === 'edit' && !selectingObjectRegion && selectableImageCandidates.map((candidate) => {
          if (activeObject?.candidateId === candidate.id) return null
          const box = pdfRectToViewport(viewport, candidate.rect)
          return (
            <button
              key={candidate.id}
              type="button"
              data-edit-ui="true"
              className="detected-object"
              style={box}
              title="Select image"
              aria-label="Select image"
              onClick={(event) => {
                event.stopPropagation()
                onRequestObjectEdit({
                  candidateId: candidate.id,
                  pageIndex,
                  kind: candidate.kind,
                  rect: candidate.rect,
                  originalRect: candidate.rect,
                  dataUrl: candidate.dataUrl || captureRect(candidate.rect),
                  opacity: 1,
                  cover: true,
                  displayRotation,
                  label: candidate.label,
                  modified: false,
                })
              }}
            />
          )
        })}

        {activeScan?.mode === 'appearance' && renderScanPatch('active-scan-patch', activeScan.patch, undefined, activeScan.paper, 7)}
        {activeText && activeTextBox && (
          <div
            className="inline-text-frame"
            data-edit-ui="true"
            data-scan-edit={activeScan ? activeScan.status : undefined}
            style={{
              ...activeTextBox,
              borderWidth: 0,
              backgroundColor: activeScan
                // Scanned text: the patch hides the old words; until it is
                // ready (or if it fails) the paper colour does. Correcting
                // only the recognised text keeps the page as it is, so the
                // typed text gets a light backing to be readable.
                ? activeScan.mode === 'recognized-text'
                  ? 'rgba(255, 255, 255, 0.88)'
                  : activeScan.patch || activeScan.status === 'ready' ? 'transparent' : colorCss(activeScan.paper ?? [1, 1, 1])
                : activeText.cover && activeText.originalRect
                  ? 'transparent'
                  : activeTextVisual?.backgroundColor || (activeText.backgroundColor ? colorCss(activeText.backgroundColor) : undefined),
            }}
          >
            <textarea
              ref={textEditorRef}
              className="inline-pdf-text-editor"
              value={activeText.text}
              wrap={resolveTextFit(activeText) === 'fit' ? 'off' : 'soft'}
              spellCheck
              aria-label="Edit text directly on the PDF"
              style={{
                fontSize: activeText.fontSize * zoom,
                fontFamily: activeText.fontFamily,
                color: colorCss(activeText.color),
                textAlign: activeText.align,
                fontWeight: activeText.fontWeight,
                fontStyle: activeText.fontStyle,
                fontStretch: activeTextVisual?.fontStretch,
                lineHeight: `${(activeTextLayout?.lineHeight || activeText.fontSize * 1.18) * zoom}px`,
                letterSpacing: `${(activeText.letterSpacing ?? 0) * zoom}px`,
                direction: activeText.direction,
                position: 'absolute',
                left: '50%',
                top: '50%',
                width: `${(activeTextSideways ? activeTextBox.height : activeTextBox.width) / editorScaleX}px`,
                height: `${activeTextSideways ? activeTextBox.width : activeTextBox.height}px`,
                transform: `translate(-50%, -50%) rotate(${activeTextScreenAngle}rad) scaleX(${editorScaleX})`,
                transformOrigin: 'center',
                backgroundColor: activeText.cover && activeText.originalRect || activeTextVisual || activeText.backgroundColor ? 'transparent' : undefined,
                whiteSpace: resolveTextFit(activeText) === 'fit' ? 'pre' : 'pre-wrap',
              }}
              onChange={(event) => {
                const target = event.currentTarget
                let rect = activeText.rect
                // A native PDF text run has a fixed baseline. Browser textarea
                // wrapping must not turn a slightly longer replacement into a
                // taller box: growing downward in viewport coordinates moves
                // rect.y (and therefore the saved PDF baseline) down a line.
                const staysSingleLine = resolveTextFit(activeText) === 'fit'
                  && !target.value.includes('\n')
                if (!staysSingleLine && Math.abs(activeTextScreenAngle) < 0.01 && target.scrollHeight > target.clientHeight + 1) {
                  const box = pdfRectToViewport(viewport, activeText.rect)
                  const desiredHeight = clamp(target.scrollHeight + 3, box.height, viewport.height - box.top)
                  rect = viewportRectToPdf(viewport, {
                    left: box.left,
                    top: box.top,
                    right: box.left + box.width,
                    bottom: box.top + desiredHeight,
                  })
                }
                onTextEditChange({
                  ...activeText,
                  rect,
                  baselineOffset: Number.isFinite(activeText.baselineOffset)
                    ? Number(activeText.baselineOffset) + activeText.rect.y - rect.y
                    : undefined,
                  text: target.value,
                  modified: true,
                  caretOffset: target.selectionStart,
                  selectionStart: target.selectionStart,
                  selectionEnd: target.selectionEnd,
                })
              }}
              onSelect={(event) => {
                const target = event.currentTarget
                if (activeText.selectionStart !== target.selectionStart || activeText.selectionEnd !== target.selectionEnd) {
                  onTextEditChange({
                    ...activeText,
                    caretOffset: target.selectionStart,
                    selectionStart: target.selectionStart,
                    selectionEnd: target.selectionEnd,
                  })
                }
              }}
              onKeyDown={(event) => {
                // Escape leaves the box and keeps what was typed (as Acrobat
                // does); the inspector's Cancel is the way to discard it.
                // (While an IME composes, Escape only cancels the composition.)
                if (event.key === 'Escape' && !event.nativeEvent.isComposing) { event.preventDefault(); onCommitTextEdit() }
                if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) { event.preventDefault(); onCommitTextEdit() }
              }}
            />
            {(activeTextOverflow || (activeTextLayout?.fitScale ?? 1) < 0.995) && (
              <span className={cx('text-fit-feedback', activeTextOverflow && 'has-overflow')} role="status">
                {activeTextOverflow
                  ? 'Text exceeds the box — enlarge or move it'
                  : `Fitted to ${Math.round((activeTextLayout?.fitScale || 1) * 100)}% width`}
              </span>
            )}
            <span
              role="presentation"
              className="selection-move-handle"
              title="Move text box"
              onPointerDown={(event) => startTransform(event, 'text', 'move')}
              onPointerMove={updateTransform}
              onPointerUp={endTransform}
            ><GripHorizontal size={13} /></span>
            {transformHandles('text')}
          </div>
        )}

        {activeObject?.cover && activeObject.originalRect && (!activeObject.candidateId || activeObject.modified) && (
          <span className="object-original-cover active-original-cover" style={pdfRectToViewport(viewport, activeObject.originalRect)} />
        )}
        {activeObject && activeObjectBox && (
          <div className={cx('object-edit-frame', activeObject.candidateId && !activeObject.modified && 'is-native-selection')} data-edit-ui="true" style={activeObjectBox}>
            {activeObject.dataUrl && (!activeObject.candidateId || activeObject.modified) && <img src={activeObject.dataUrl} alt={activeObject.label} draggable={false} style={{ opacity: activeObject.opacity, ...objectImageStyle(activeObjectBox) }} />}
            <span
              role="presentation"
              className="selection-move-handle"
              title={`Move ${activeObject.label.toLowerCase()}`}
              onPointerDown={(event) => startTransform(event, 'object', 'move')}
              onPointerMove={updateTransform}
              onPointerUp={endTransform}
            ><GripHorizontal size={13} /></span>
            {transformHandles('object')}
          </div>
        )}

        {regionDraft && <div className="object-region-selection" style={regionDraft}><span>Select artwork</span></div>}

        {shapeDraft && <div className="rectangle-draft" style={shapeDraft} />}
        {inkDraft.length > 1 && (
          <svg className="ink-draft" viewBox={`0 0 ${viewport.width} ${viewport.height}`} aria-hidden="true">
            <polyline points={inkDraft.map((point) => `${point.x},${point.y}`).join(' ')} fill="none" stroke="#dc2626" strokeWidth={2 * zoom} strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        )}

        {cropDraft && (
          <div className="crop-mask" aria-label="Crop selection">
            <div className="crop-selection" style={cropDraft}>
              <span className="crop-size">{Math.round(cropDraft.width / zoom)} × {Math.round(cropDraft.height / zoom)} pt</span>
              <div className="crop-actions">
                <button type="button" title="Cancel crop" onClick={(event) => { event.stopPropagation(); setCropDraft(null) }}><X size={14} /></button>
                <button
                  type="button"
                  className="apply-crop"
                  title="Apply crop"
                  onClick={(event) => {
                    event.stopPropagation()
                    const local = { left: cropDraft.left, top: cropDraft.top, right: cropDraft.left + cropDraft.width, bottom: cropDraft.top + cropDraft.height }
                    onCrop(viewportRectToPdf(viewport, local))
                    setCropDraft(null)
                  }}
                ><Check size={14} /> Apply</button>
              </div>
            </div>
          </div>
        )}

        {loading && !canvasReadyViewport && <div className="page-rendering"><LoaderCircle className="spin" size={18} /></div>}
        {renderError && <div className="page-error"><Crop size={18} /><strong>Page unavailable</strong><span>{renderError}</span></div>}
      </div>
    </div>
  )
})
