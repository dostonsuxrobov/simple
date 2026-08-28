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
  PdfOverlay,
  PdfRect,
  ToolMode,
} from '../types'
import { getPageTextContent, pdfjs, pdfRectToViewport, viewportRectToPdf } from '../lib/pdf'
import { detectPageObjects } from '../lib/pageObjects'
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

interface PdfPageProps {
  pdf: PDFDocumentProxy
  pageIndex: number
  zoom: number
  rotation: number
  tool: ToolMode
  overlays: PdfOverlay[]
  formValues: Record<string, string | boolean>
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
  onFormChange: (name: string, value: string | boolean) => void
}

const RESIZE_HANDLES: ResizeHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']

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
  onTextMarkup, onInk, onRectangle,
}: PdfPageProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const surfaceRef = useRef<HTMLDivElement>(null)
  const textLayerRef = useRef<HTMLDivElement>(null)
  const searchMatchLayerRef = useRef<HTMLDivElement>(null)
  const textEditorRef = useRef<HTMLTextAreaElement>(null)
  const transformRef = useRef<TransformGesture | null>(null)
  const textVisualStyleRef = useRef<InlineTextVisualStyle | null>(null)
  const fontDataRef = useRef(new Map<string, Uint8Array>())
  const [page, setPage] = useState<PDFPageProxy | null>(null)
  const [loading, setLoading] = useState(true)
  const [renderViewport, setRenderViewport] = useState<ReturnType<PDFPageProxy['getViewport']> | null>(null)
  const [canvasReadyViewport, setCanvasReadyViewport] = useState<ReturnType<PDFPageProxy['getViewport']> | null>(null)
  const [renderError, setRenderError] = useState('')
  const [annotations, setAnnotations] = useState<any[]>([])
  const [imageCandidates, setImageCandidates] = useState<DetectedPageObject[]>([])
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
    setPage(null)
    setLoading(true)
    setRenderError('')
    setCanvasReadyViewport(null)
    setAnnotations([])
    setImageCandidates([])
    pdf.getPage(pageIndex + 1).then((nextPage) => {
      if (active) setPage(nextPage)
    }).catch((error) => {
      if (active) setRenderError(error instanceof Error ? error.message : 'Page could not be opened.')
    })
    return () => { active = false }
  }, [pdf, pageIndex])

  const viewport = useMemo(() => {
    if (!page) return null
    const angle = (((page.rotate || 0) + rotation) % 360 + 360) % 360
    return page.getViewport({ scale: zoom, rotation: angle })
  }, [page, zoom, rotation])
  const displayRotation = ((((viewport?.rotation || 0) % 360) + 360) % 360) as DisplayRotation

  useEffect(() => {
    if (!viewport) return
    setCropDraft(null)
    setRegionDraft(null)
    setShapeDraft(null)
    setInkDraft([])
  }, [viewport])

  useEffect(() => {
    if (!viewport) return
    onPageReady?.({ width: viewport.width, height: viewport.height })
  }, [viewport, onPageReady])

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
    canvas.style.width = `${viewport.width}px`
    canvas.style.height = `${viewport.height}px`
    // Spans built for another viewport must never stay visible floating over
    // the rescaled canvas; the rebuild restores visibility when it lands.
    const layer = textLayerRef.current
    if (layer?.dataset.viewportSignature && layer.dataset.viewportSignature !== textViewportSignature(viewport)) {
      layer.style.visibility = 'hidden'
    }
  }, [viewport])

  useEffect(() => {
    if (!page || !renderViewport || !canvasRef.current) return
    let cancelled = false
    const canvas = canvasRef.current
    // Chromium's grayscale canvas text is visibly softer than native PDF
    // viewers when it is rendered at exactly one backing pixel per CSS pixel.
    // A modest supersampling floor restores fine serifs on common 100%-scaled
    // Windows displays, while the area cap keeps oversized architectural pages
    // bounded. ContinuousPdfViewer still mounts only nearby canvases.
    const desiredDpr = Math.min(Math.max(window.devicePixelRatio || 1, 1.5), 2.5)
    const maxCanvasPixels = 24_000_000
    const pixelArea = renderViewport.width * renderViewport.height * desiredDpr * desiredDpr
    const dpr = pixelArea > maxCanvasPixels
      ? Math.max(1, desiredDpr * Math.sqrt(maxCanvasPixels / pixelArea))
      : desiredDpr
    canvas.width = Math.max(1, Math.ceil(renderViewport.width * dpr))
    canvas.height = Math.max(1, Math.ceil(renderViewport.height * dpr))
    const context = canvas.getContext('2d', { alpha: false })
    if (!context) return
    // Resizing an opaque canvas clears it to black. Paint the page background
    // right away so rotate/zoom/mount never flash before pdf.js renders.
    context.fillStyle = '#fff'
    context.fillRect(0, 0, canvas.width, canvas.height)
    setLoading(true)
    setRenderError('')
    const renderTask = page.render({
      canvasContext: context,
      viewport: renderViewport,
      transform: dpr === 1 ? undefined : [canvas.width / renderViewport.width, 0, 0, canvas.height / renderViewport.height, 0, 0],
    })
    renderTask.promise.then(() => {
      if (!cancelled) {
        setLoading(false)
        setCanvasReadyViewport(renderViewport)
      }
    }).catch((error) => {
      if (!cancelled && error?.name !== 'RenderingCancelledException') {
        setRenderError(error instanceof Error ? error.message : 'This page could not be rendered.')
        setLoading(false)
      }
    })
    return () => {
      cancelled = true
      renderTask.cancel()
    }
  }, [page, renderViewport])

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
    if (!page || !viewport || canvasReadyViewport !== viewport || !textLayerRef.current) return
    const currentPage = page
    const currentViewport = viewport
    let cancelled = false
    const layer = textLayerRef.current
    const viewportSignature = textViewportSignature(currentViewport)
    if (liveSelectionIntersectsTextLayer(layer)) {
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

    async function renderText() {
      const content = await getPageTextContent(currentPage)
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

      for (const [itemIndex, rawItem] of content.items.entries()) {
        if (!('str' in rawItem) || !rawItem.str) continue
        const item = rawItem as typeof rawItem & { transform: number[]; width: number; height: number; fontName: string; dir?: string }
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
        span.dataset.fontSize = String(pdfFontSize || fontHeight / zoom)
        span.dataset.fontFamily = fontFamily
        span.dataset.fontWeight = String(fontTraits.weight)
        span.dataset.fontStyle = fontTraits.style
        span.dataset.pdfFontName = item.fontName
        span.dataset.pdfFontSourceName = sourceFontName
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
        spans.push({
          span,
          width: (style.vertical ? item.height : item.width) * currentViewport.scale,
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
        if (line && Math.abs(item.flow - line.flow) <= Math.max(1, Math.max(item.fontHeight, line.fontHeight) * 0.4)) line.items.push(item)
        else lines.push({ flow: item.flow, fontHeight: item.fontHeight, items: [item] })
      }
      for (const line of lines) {
        for (const entry of [...line.items].sort((a, b) => a.read - b.read)) fragment.appendChild(entry.span)
      }

      layer.appendChild(fragment)
      layer.dataset.viewportSignature = viewportSignature
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
      if (!liveSelectionIntersectsTextLayer(layer)) layer.replaceChildren()
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

  useEffect(() => {
    if (!activeSearchKey || searchMatchGeometry.key !== activeSearchKey || !searchMatchGeometry.rects.length) return
    const frame = window.requestAnimationFrame(() => {
      searchMatchLayerRef.current?.querySelector<HTMLElement>('[data-search-match-rect="true"]')
        ?.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'auto' })
    })
    return () => window.cancelAnimationFrame(frame)
  }, [activeSearchKey, searchMatchGeometry])

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
  const selectableImageCandidates = useMemo(() => imageCandidates.filter((candidate) => (
    !pageOverlays.some((overlay) => overlay.type === 'object'
      && overlay.cover
      && overlay.originalRect
      && Math.abs(overlay.originalRect.x - candidate.rect.x) <= 1
      && Math.abs(overlay.originalRect.y - candidate.rect.y) <= 1
      && Math.abs(overlay.originalRect.width - candidate.rect.width) <= 1
      && Math.abs(overlay.originalRect.height - candidate.rect.height) <= 1)
  )), [imageCandidates, pageOverlays])

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
    const caret = caretOffsetAtPoint(span, clientX, clientY)
    const pdfBaselineY = Number(span.dataset.pdfBaselineY)
    const pdfTextAngle = Number(span.dataset.pdfTextAngle) || 0
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
      text,
      fontSize: Number(span.dataset.fontSize) || 11,
      fontFamily: family,
      fontWeight: Number.parseInt(span.dataset.fontWeight || computed.fontWeight, 10) || 400,
      fontStyle: (span.dataset.fontStyle || computed.fontStyle) === 'italic' ? 'italic' : 'normal',
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
      sourceItemText,
      sourceItemRect,
      sourceSelectionStart: 0,
      sourceSelectionEnd: sourceItemText.length,
      align: computed.direction === 'rtl' ? 'right' : 'left',
      color: colors.text,
      backgroundColor: cssColorToPdf(colors.background),
      cover: true,
      modified: false,
      caretOffset: caret,
      selectionStart: selectionSlice?.start ?? caret,
      selectionEnd: selectionSlice?.end ?? caret,
    }
  }

  function handleClick(event: React.MouseEvent) {
    if ((event.target as HTMLElement).closest('[data-edit-ui="true"]')) return
    if (tool === 'edit') {
      const span = (event.target as HTMLElement).closest<HTMLSpanElement>('[data-text-item="true"]')
      if (span) {
        event.preventDefault()
        event.stopPropagation()
        const edit = textEditFromSpan(span, event.clientX, event.clientY)
        if (edit) onRequestTextEdit(edit)
      }
      return
    }
    // Markup tools operate on a real DOM Range in handleMouseUp. A click with
    // no range intentionally does nothing; a PDF.js text item may represent a
    // complete line and must never be mistaken for a focused selection.
    if (tool === 'highlight' || tool === 'underline' || tool === 'strikeout') return
    if (tool === 'sign') {
      if (!viewport || (event.target as HTMLElement).closest('button, input, textarea')) return
      const point = localPoint(event.clientX, event.clientY)
      if (!point) return
      const [x, y] = viewport.convertToPdfPoint(point.x, point.y)
      onPlaceSignature({ x, y }, displayRotation)
      return
    }
    if (tool !== 'addText' || !viewport || !surfaceRef.current) return
    if ((event.target as HTMLElement).closest('button, input, textarea')) return
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
    if (!viewport) return
    const target = event.target as HTMLElement
    if (target.closest('[data-edit-ui="true"],button,input,textarea')) return
    if (tool === 'edit') {
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
    if (gesture.target === 'text' && activeText) onTextEditChange({ ...activeText, rect, modified: true })
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

  if (!viewport) {
    return <div className="page-loading"><LoaderCircle className="spin" size={20} /><span>Preparing page {pageIndex + 1}…</span></div>
  }

  const activeTextBox = activeText ? pdfRectToViewport(viewport, activeText.rect) : null
  const activeObjectBox = activeObject ? pdfRectToViewport(viewport, activeObject.rect) : null
  const rememberedTextVisual = textVisualStyleRef.current
  const activeTextVisual = activeText
    && rememberedTextVisual?.pageIndex === pageIndex
    && rememberedTextVisual.sourceText === activeText.originalText
    && sameRect(rememberedTextVisual.sourceRect, activeText.originalRect || activeText.rect)
    ? rememberedTextVisual
    : null
  const editorScaleX = activeTextVisual && Math.abs(activeTextVisual.angle) < 0.01
    ? activeTextVisual.scaleX
    : activeText?.scaleX || 1

  return (
    <div className="page-stage" style={{ minHeight: 'auto' }}>
      <div
        ref={surfaceRef}
        className={cx('page-surface', `tool-${tool}`, selectingObjectRegion && 'is-selecting-object')}
        style={{ width: viewport.width, height: viewport.height }}
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
            if (annotation.subtype === 'Widget' && annotation.fieldType === 'Tx' && annotation.fieldName) {
              return (
                <input
                  key={annotation.id}
                  className="pdf-form-field"
                  style={box}
                  value={String(formValues[annotation.fieldName] ?? annotation.fieldValue ?? '')}
                  onChange={(event) => onFormChange(annotation.fieldName, event.target.value)}
                />
              )
            }
            if (annotation.subtype === 'Widget' && annotation.fieldType === 'Btn' && annotation.checkBox && annotation.fieldName) {
              return (
                <input
                  key={annotation.id}
                  type="checkbox"
                  className="pdf-form-checkbox"
                  style={box}
                  checked={Boolean(formValues[annotation.fieldName] ?? annotation.fieldValue)}
                  onChange={(event) => onFormChange(annotation.fieldName, event.target.checked)}
                />
              )
            }
            return null
          })}
        </div>

        <div className="overlay-layer">
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
                    ><img src={overlay.dataUrl} alt="Edited PDF object" draggable={false} /></button>
                  )}
                </div>
              )
            }
            if (activeText?.overlayId === overlay.id) return null
            const originalBox = overlay.cover && overlay.originalRect ? pdfRectToViewport(viewport, overlay.originalRect) : null
            return (
              <div key={overlay.id} className="committed-text-group">
                {originalBox && <span className="text-original-cover" style={{ ...originalBox, backgroundColor: overlay.backgroundColor ? colorCss(overlay.backgroundColor) : undefined }} />}
                <button
                  type="button"
                  data-edit-ui="true"
                  className={cx('text-overlay', overlay.cover && 'is-replacement', tool === 'edit' && 'is-editable')}
                  style={{
                    ...box,
                    fontSize: overlay.fontSize * zoom,
                    fontFamily: overlay.fontFamily,
                    fontWeight: overlay.fontWeight,
                    fontStyle: overlay.fontStyle,
                    lineHeight: overlay.lineHeight ? `${overlay.lineHeight * zoom}px` : undefined,
                    letterSpacing: overlay.letterSpacing ? `${overlay.letterSpacing * zoom}px` : undefined,
                    direction: overlay.direction,
                    textAlign: overlay.align || 'left',
                    color: colorCss(overlay.color),
                    backgroundColor: overlay.backgroundColor ? colorCss(overlay.backgroundColor) : undefined,
                    transform: overlay.scaleX && Math.abs(overlay.scaleX - 1) > 0.01 ? `scaleX(${overlay.scaleX})` : undefined,
                    transformOrigin: '0 0',
                  }}
                  onClick={(event) => {
                    event.stopPropagation()
                    if (tool !== 'edit') return
                    textVisualStyleRef.current = null
                    onRequestTextEdit({
                      overlayId: overlay.id,
                      pageIndex,
                      rect: overlay.rect,
                      originalRect: overlay.originalRect,
                      originalText: overlay.originalText ?? overlay.text,
                      text: overlay.text,
                      fontSize: overlay.fontSize,
                      fontFamily: overlay.fontFamily,
                      fontWeight: overlay.fontWeight,
                      fontStyle: overlay.fontStyle,
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
                  }}
                >{overlay.text}</button>
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

        {activeText?.cover && activeText.originalRect && !sameRect(activeText.rect, activeText.originalRect) && (
          <span className="text-original-cover active-original-cover" style={{
            ...pdfRectToViewport(viewport, activeText.originalRect),
            backgroundColor: activeText.backgroundColor ? colorCss(activeText.backgroundColor) : undefined,
          }} />
        )}
        {activeText && activeTextBox && (
          <div
            className="inline-text-frame"
            data-edit-ui="true"
            style={{
              ...activeTextBox,
              borderWidth: 0,
              backgroundColor: activeTextVisual?.backgroundColor || (activeText.backgroundColor ? colorCss(activeText.backgroundColor) : undefined),
            }}
          >
            <textarea
              ref={textEditorRef}
              className="inline-pdf-text-editor"
              value={activeText.text}
              wrap={activeText.originalText && !activeText.originalText.includes('\n') ? 'off' : 'soft'}
              spellCheck
              aria-label="Edit text directly on the PDF"
              style={{
                fontSize: activeText.fontSize * zoom,
                fontFamily: activeText.fontFamily,
                color: colorCss(activeText.color),
                textAlign: activeText.align,
                fontWeight: activeTextVisual?.fontWeight || activeText.fontWeight,
                fontStyle: activeTextVisual?.fontStyle || activeText.fontStyle,
                fontStretch: activeTextVisual?.fontStretch,
                lineHeight: `${(activeTextVisual?.lineHeight || activeText.lineHeight || activeText.fontSize * 1.18) * zoom}px`,
                letterSpacing: `${(activeTextVisual?.letterSpacing ?? activeText.letterSpacing ?? 0) * zoom}px`,
                direction: activeTextVisual?.direction || activeText.direction,
                width: editorScaleX === 1 ? '100%' : `${100 / editorScaleX}%`,
                transform: editorScaleX === 1 ? undefined : `scaleX(${editorScaleX})`,
                transformOrigin: '0 0',
                backgroundColor: activeTextVisual || activeText.backgroundColor ? 'transparent' : undefined,
                whiteSpace: activeText.originalText && !activeText.originalText.includes('\n') ? 'pre' : 'pre-wrap',
              }}
              onChange={(event) => {
                const target = event.currentTarget
                let rect = activeText.rect
                // A native PDF text run has a fixed baseline. Browser textarea
                // wrapping must not turn a slightly longer replacement into a
                // taller box: growing downward in viewport coordinates moves
                // rect.y (and therefore the saved PDF baseline) down a line.
                const staysSingleLine = Boolean(activeText.originalText)
                  && !activeText.originalText.includes('\n')
                  && !target.value.includes('\n')
                if (!staysSingleLine && target.scrollHeight > target.clientHeight + 1) {
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
                if (event.key === 'Escape') { event.preventDefault(); onCancelTextEdit() }
                if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) { event.preventDefault(); onCommitTextEdit() }
              }}
            />
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
            {activeObject.dataUrl && (!activeObject.candidateId || activeObject.modified) && <img src={activeObject.dataUrl} alt={activeObject.label} draggable={false} style={{ opacity: activeObject.opacity }} />}
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

        {loading && <div className="page-rendering"><LoaderCircle className="spin" size={18} /></div>}
        {renderError && <div className="page-error"><Crop size={18} /><strong>Page unavailable</strong><span>{renderError}</span></div>}
      </div>
    </div>
  )
})
