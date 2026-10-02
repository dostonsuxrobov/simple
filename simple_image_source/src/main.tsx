import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import {
  Brush,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ClipboardPaste,
  Copy,
  Crop,
  Eraser,
  FileDown,
  FileImage,
  FolderOpen,
  ImagePlus,
  Layers,
  Maximize2,
  Minus,
  MousePointer2,
  PenLine,
  Pipette,
  Printer,
  Redo2,
  RotateCcw,
  RotateCw,
  Save,
  Scaling,
  SlidersHorizontal,
  TriangleAlert,
  Undo2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'
import { INSPECTOR_STORAGE_KEY, inspectorOpenFromStored, storedInspectorOpen } from './inspector-preference.js'
import { PrintDialog } from './PrintDialog'
import type { ImagePrintSettings } from '../electron/print-layout.mjs'
import type { ExportFormat, OpenImage, RevisionClock, SaveFormat } from './shared/session.ts'
import './styles.css'
import './simple/simple.css'
import { pixelsHaveAlpha, useAdvancedMode } from './advanced/useAdvancedMode.tsx'
import type { LookId, PixelBuffer, QuickAdjust, ResampleMethod } from './imaging/types.ts'
import { cropBuffer, hasTransparency } from './imaging/buffer.ts'
import { putBuffer, readCanvasPixels } from './shared/canvas.ts'
import { getImagingClient } from './shared/workerClient.ts'
import { cssFont, drawShape, drawText } from './shared/vector.ts'
import type { AspectPreset, CropHandle } from './simple/cropMath.ts'
import {
  constrainDrag,
  constrainDragInside,
  cropFrame,
  defaultCropRect,
  defaultPortrait,
  fitAspect,
  inscribedCrop,
  insideRotated,
  levelAngle,
  nudgeRect,
  presetCanSwap,
  presetRatio,
  rectFromPoints,
  rectFromPointsInside,
  toIntegerRect,
} from './simple/cropMath.ts'
import type { GeometryOp, HistoryEntry, HistoryState, PatchRecorder } from './simple/simpleHistory.ts'
import { clipRect, createPatchRecorder, geometryEntry, pushEntry, redoEntry, snapshotEntry, undoEntry } from './simple/simpleHistory.ts'
import { createCanvasSurface, cropCanvas, encodeCanvas, imageContext, readRegion, regionHasTransparency, replaceCanvasPixels, rgbHex, rotatedRegionCanvas, sampleColor, transformCanvas } from './simple/ops.ts'
import type { EncodeHints } from './simple/encodeHints.ts'
import { NO_ENCODE_HINTS, encodeHintsFor, saveQuality } from './simple/encodeHints.ts'
import { MIN_ZOOM, clampZoom, isPanEvent, usePointerReadout, useViewportNavigation } from './simple/useViewportNavigation.ts'
import { CropOptions, LevelLineLayer } from './simple/CropOptions.tsx'
import { ResizeDialog } from './simple/ResizeDialog.tsx'
import { AdjustPanel } from './simple/AdjustPanel.tsx'
import type { AdjustTab } from './simple/AdjustPanel.tsx'
import type { PendingAdjust } from './simple/useAdjustPreview.ts'
import { NEUTRAL_ADJUST, hasLook, isNeutralAdjust, isNeutralQuick, quickSpec, useAdjustPreview } from './simple/useAdjustPreview.ts'
import { MarkupLayer, MarkupOptions, useMarkupSession } from './simple/Markup.tsx'
import { createImageItem, markupBounds, nextMarkupId, paintableItems, textStyleOf, toShapeSpec, toTextSpec } from './simple/markupModel.ts'
import { autoEnhance } from './imaging/autoEnhance.ts'
import { getSimpleIO } from './simple-io/io-client.ts'

type Tool = 'view' | 'crop' | 'brush' | 'eraser' | 'eyedropper' | 'markup'
type PaintTool = 'brush' | 'eraser'
type SimplePanel = 'details' | 'adjust'

interface Dimensions { width: number; height: number }
interface Point { x: number; y: number }
interface CropRect { x: number; y: number; width: number; height: number }
interface CropDrag { handle: CropHandle | 'new'; start: Point; initial: CropRect; pointerId: number }
interface ConfirmState { title: string; detail: string; action: () => void | Promise<void>; cropPending?: boolean }
interface ChoiceOption { value: string; label: string; primary?: boolean }
interface ChoiceState { title: string; detail: string; tone: 'crop' | 'warning'; options: ChoiceOption[]; resolve: (value: string) => void }
interface PrintSnapshot { data: Uint8Array; url: string; name: string; width: number; height: number }
interface ImageContextMenu { x: number; y: number }
interface WorkingState { label: string; progress: number | null }
/** Fields of the open payload that src/types.d.ts does not declare yet (electron/main.cjs openPayload). */
interface OpenPayloadExtras {
  animated?: boolean
  frameCount?: number | null
  notices?: string[]
  suggestedName?: string | null
  sourceToken?: string
}
/** Per-document facts, keyed by documentIdRef so a load elsewhere (WP1's PSD open) can never leave stale data. */
interface DocumentMeta {
  documentId: number
  /** Nothing has changed the pixels since the file was opened (Undo back here restores it). */
  pristine: boolean
  sourceToken: string | null
  animated: boolean
  hints: EncodeHints
}
/** A brush or eraser stroke while the pointer is down. */
interface StrokeState {
  pointerId: number
  erasing: boolean
  color: string
  size: number
  hadAlpha: boolean
  last: Point
  /** Open history segment (null after Ctrl+S mid-stroke: the next draw opens a new step). */
  recorder: PatchRecorder | null
  before: HistoryState
  revision: number
  element: HTMLCanvasElement
}

const SUPPORTED = ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'svg', 'avif', 'psd']
const EXPORT_OPTIONS: ReadonlyArray<{ format: ExportFormat; label: string; detail: string }> = [
  { format: 'png', label: 'PNG image', detail: 'Lossless and transparency-safe' },
  { format: 'jpeg', label: 'JPEG image', detail: 'Compact for photos; transparent areas become white' },
  { format: 'webp', label: 'WebP image', detail: 'Smaller modern image with transparency' },
  { format: 'pdf', label: 'PDF document', detail: 'One page sized to the image aspect ratio' },
]
const MAX_CANVAS_DIMENSION = 20_000
const MAX_CANVAS_PIXELS = 50_000_000
const MAX_BRUSH_SIZE = 120
const PASTED_NAME = 'Pasted image.png'

/**
 * The letter of a shortcut by physical key when Ctrl/Cmd is held, so Ctrl+S, Ctrl+Z and the
 * rest work on Cyrillic, Greek, Hebrew, Arabic and other non-Latin layouts too.
 */
function shortcutKey(event: KeyboardEvent) {
  const key = event.key.toLowerCase()
  if ((event.ctrlKey || event.metaKey) && /^Key[A-Z]$/.test(event.code) && !/^[a-z]$/.test(key)) return event.code.slice(3).toLowerCase()
  return key
}

function clamp(value: number, minimum: number, maximum: number) {
  return Math.max(minimum, Math.min(maximum, value))
}

function formatBytes(bytes: number) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 ** 2).toFixed(bytes < 10 * 1024 ** 2 ? 1 : 0)} MB`
}

function gcd(a: number, b: number): number {
  return b ? gcd(b, a % b) : a
}

function aspectRatio(width: number, height: number) {
  const divisor = gcd(width, height)
  const left = width / divisor
  const right = height / divisor
  return left <= 100 && right <= 100 ? `${left}:${right}` : (width / height).toFixed(2)
}

function defaultSaveFormat(format: string): SaveFormat {
  if (format === 'jpg' || format === 'jpeg') return 'jpeg'
  if (format === 'webp') return 'webp'
  return 'png'
}

function outputMime(format: SaveFormat) {
  if (format === 'psd') return 'image/vnd.adobe.photoshop'
  return format === 'jpeg' ? 'image/jpeg' : `image/${format}`
}

function canvasToBlob(canvas: HTMLCanvasElement, format: SaveFormat, quality = 0.92): Promise<Blob> {
  // A canvas cannot encode PSD (toBlob would silently fall back to PNG); layered output comes from Advanced.
  if (format === 'psd') return Promise.reject(new Error('Photoshop documents are saved from the Advanced editor.'))
  return encodeCanvas(canvas, format, quality)
}

function canvasHasTransparency(canvas: HTMLCanvasElement) {
  return regionHasTransparency(canvas)
}

/** Brush sizes [ and ] step through: every px to 10, then 5 px to 50, 10 px to 100, then 120. */
const BRUSH_LADDER: readonly number[] = [
  ...Array.from({ length: 10 }, (_, index) => index + 1),
  15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90, 100, MAX_BRUSH_SIZE,
]

function stepBrushSize(size: number, direction: 1 | -1) {
  if (direction > 0) return BRUSH_LADDER.find((value) => value > size) ?? MAX_BRUSH_SIZE
  for (let index = BRUSH_LADDER.length - 1; index >= 0; index -= 1) if (BRUSH_LADDER[index] < size) return BRUSH_LADDER[index]
  return 1
}

function messageOf(error: unknown, fallback: string) {
  return error instanceof Error && error.message ? error.message : fallback
}

function App() {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const contextMenuRef = useRef<HTMLDivElement>(null)
  const viewportRef = useRef<HTMLDivElement>(null)
  const stackRef = useRef<HTMLDivElement>(null)
  const adjustOverlayRef = useRef<HTMLCanvasElement>(null)
  const brushRingRef = useRef<HTMLDivElement>(null)
  const undoRef = useRef<HistoryEntry[]>([])
  const redoRef = useRef<HistoryEntry[]>([])
  const paintingRef = useRef(false)
  const lastPointRef = useRef<Point | null>(null)
  const strokeRef = useRef<StrokeState | null>(null)
  const lastStrokeEndRef = useRef<{ documentId: number; point: Point } | null>(null)
  const cropDragRef = useRef<CropDrag | null>(null)
  const cropAdjustedRef = useRef(false)
  /** The user already answered the in-app close question; the window guard must not ask again. */
  const closeDecidedRef = useRef(false)
  const toastTimerRef = useRef<number | null>(null)
  const imageRef = useRef<OpenImage | null>(null)
  const dirtyRef = useRef(false)
  const documentIdRef = useRef(0)
  const loadRequestRef = useRef(0)
  const originalImageRef = useRef<{ blob: Blob; format: SaveFormat; printable: boolean } | null>(null)
  const nextRevisionRef = useRef(0)
  const contentRevisionRef = useRef(0)
  const savedRevisionRef = useRef(0)
  const savingRef = useRef(false)
  const workingRef = useRef(false)
  const docMetaRef = useRef<DocumentMeta>({ documentId: 0, pristine: true, sourceToken: null, animated: false, hints: NO_ENCODE_HINTS })
  const pastedImagesRef = useRef(new Map<string, HTMLImageElement>())
  const saveGroupRef = useRef<HTMLDivElement>(null)
  const exportGroupRef = useRef<HTMLDivElement>(null)
  const printButtonRef = useRef<HTMLButtonElement>(null)
  const printFocusReturnRef = useRef<HTMLElement | null>(null)
  const printPendingFocusRef = useRef<HTMLElement | null>(null)
  const printDialogOpenRef = useRef(false)
  /** Fit mode as of the latest request (state lags: a resize notification must not re-fit a zoom the user just chose). */
  const fitModeRef = useRef(true)
  // Late-bound Simple hooks for WP1's settleSimple / adoptFlattened (declared inside its region).
  const simpleSettleRef = useRef<() => Promise<void>>(async () => {})
  const simpleAdoptRef = useRef<(flattened: HTMLCanvasElement, beforeRevision: number, pushUndoStep: boolean) => void>(() => {})

  const [image, setImage] = useState<OpenImage | null>(null)
  const [album, setAlbum] = useState<{ paths: string[]; index: number }>({ paths: [], index: -1 })
  const [dimensions, setDimensions] = useState<Dimensions>({ width: 0, height: 0 })
  const [dirty, setDirtyState] = useState(false)
  const [saving, setSaving] = useState(false)
  const [exporting, setExporting] = useState<ExportFormat | null>(null)
  const [printing, setPrinting] = useState(false)
  const [printSnapshot, setPrintSnapshot] = useState<PrintSnapshot | null>(null)
  const [tool, setTool] = useState<Tool>('view')
  const [paintTool, setPaintTool] = useState<PaintTool>('brush')
  const [brushColor, setBrushColor] = useState('#111111')
  const [brushSize, setBrushSize] = useState(12)
  const [zoom, setZoom] = useState(1)
  const [fitMode, setFitMode] = useState(true)
  const [cropRect, setCropRect] = useState<CropRect | null>(null)
  const [cropAspect, setCropAspect] = useState<AspectPreset>('free')
  const [cropPortrait, setCropPortrait] = useState(false)
  const [straighten, setStraighten] = useState(0)
  const [leveling, setLeveling] = useState(false)
  const [cropAdjusted, setCropAdjustedState] = useState(false)
  const [historyVersion, setHistoryVersion] = useState(0)
  const [isMaximized, setIsMaximized] = useState(false)
  const [draggingFile, setDraggingFile] = useState(false)
  useEffect(() => {
    const clear = () => setDraggingFile(false)
    const onLeave = (event: DragEvent) => { if (!event.relatedTarget) clear() }
    window.addEventListener('dragleave', onLeave)
    window.addEventListener('drop', clear)
    window.addEventListener('dragend', clear)
    window.addEventListener('blur', clear)
    return () => {
      window.removeEventListener('dragleave', onLeave)
      window.removeEventListener('drop', clear)
      window.removeEventListener('dragend', clear)
      window.removeEventListener('blur', clear)
    }
  }, [])
  const [toast, setToast] = useState<{ message: string; tone: 'normal' | 'error' } | null>(null)
  const [confirmState, setConfirmState] = useState<ConfirmState | null>(null)
  const [choiceState, setChoiceState] = useState<ChoiceState | null>(null)
  const [saveMenu, setSaveMenu] = useState(false)
  const [exportMenu, setExportMenu] = useState(false)
  const [contextMenu, setContextMenu] = useState<ImageContextMenu | null>(null)
  const [simplePanel, setSimplePanel] = useState<SimplePanel>('details')
  const [adjustTab, setAdjustTab] = useState<AdjustTab>('light')
  const [pendingAdjust, setPendingAdjust] = useState<PendingAdjust>(NEUTRAL_ADJUST)
  const [comparing, setComparing] = useState(false)
  const [resizeOpen, setResizeOpen] = useState(false)
  const [working, setWorkingState] = useState<WorkingState | null>(null)
  const [inspectorOpen, setInspectorOpen] = useState(() => {
    try { return inspectorOpenFromStored(window.localStorage.getItem(INSPECTOR_STORAGE_KEY)) } catch { return true }
  })

  imageRef.current = image
  dirtyRef.current = dirty
  printDialogOpenRef.current = Boolean(printSnapshot)
  const markup = useMarkupSession()

  const reflectDirty = useCallback(() => {
    const value = contentRevisionRef.current !== savedRevisionRef.current
    dirtyRef.current = value
    setDirtyState(value)
  }, [])

  /** The current document's facts (fresh defaults when another path loaded a new document). */
  const documentMeta = useCallback((): DocumentMeta => {
    if (docMetaRef.current.documentId !== documentIdRef.current) {
      docMetaRef.current = { documentId: documentIdRef.current, pristine: true, sourceToken: null, animated: false, hints: NO_ENCODE_HINTS }
    }
    return docMetaRef.current
  }, [])

  const setPristine = useCallback((value: boolean) => {
    documentMeta().pristine = value
  }, [documentMeta])

  /** IMAGE-SIE-1: Save may write the original bytes only when nothing changed the pixels since open. */
  const isPristine = useCallback(() => documentMeta().pristine && contentRevisionRef.current === 0, [documentMeta])

  const markEdited = useCallback(() => {
    nextRevisionRef.current += 1
    contentRevisionRef.current = nextRevisionRef.current
    setPristine(false)
    reflectDirty()
  }, [reflectDirty, setPristine])

  const setCropAdjusted = useCallback((value: boolean) => {
    cropAdjustedRef.current = value
    setCropAdjustedState(value)
  }, [])

  const setWorking = useCallback((value: WorkingState | null) => {
    workingRef.current = Boolean(value)
    setWorkingState(value)
  }, [])

  const notify = useCallback((message: string, tone: 'normal' | 'error' = 'normal') => {
    if (toastTimerRef.current) window.clearTimeout(toastTimerRef.current)
    setToast({ message, tone })
    toastTimerRef.current = window.setTimeout(() => setToast(null), tone === 'error' ? 5200 : 3200)
  }, [])

  // While straightening, the document box is the turned image's frame.
  const displaySize = useMemo<Dimensions>(
    () => (tool === 'crop' && straighten ? cropFrame(dimensions, straighten) : dimensions),
    [dimensions, straighten, tool],
  )

  const fitImage = useCallback((width = displaySize.width, height = displaySize.height) => {
    const viewport = viewportRef.current
    if (!viewport || !width || !height) return
    const availableWidth = Math.max(80, viewport.clientWidth - 88)
    const availableHeight = Math.max(80, viewport.clientHeight - 88)
    fitModeRef.current = true
    setZoom(clampZoom(Math.max(MIN_ZOOM, Math.min(availableWidth / width, availableHeight / height, 1)), { width, height }))
    setFitMode(true)
  }, [displaySize.height, displaySize.width])

  const applyZoom = useCallback((next: number) => {
    fitModeRef.current = false
    setZoom(next)
    setFitMode(false)
  }, [])

  /** Re-fits after the canvas size changed, unless the user picked a zoom meanwhile. */
  const refitSoon = useCallback((width: number, height: number) => {
    if (!fitModeRef.current) return
    requestAnimationFrame(() => { if (fitModeRef.current) fitImage(width, height) })
  }, [fitImage])

  const navigation = useViewportNavigation({
    viewportRef,
    stackRef,
    enabled: Boolean(image),
    zoom,
    contentSize: displaySize,
    applyZoom,
    fit: () => fitImage(),
    fitMode,
    panWithPrimary: tool === 'view',
  })

  const readout = usePointerReadout(canvasRef, Boolean(image) && !(tool === 'crop' && straighten !== 0))

  const context2d = useCallback(() => {
    const canvas = canvasRef.current
    if (!canvas) throw new Error('The image canvas is unavailable.')
    return imageContext(canvas)
  }, [])

  /** The document state an edit starts from (undo restores it exactly). */
  const currentState = useCallback((): HistoryState => ({
    revision: contentRevisionRef.current,
    hasAlpha: imageRef.current?.hasAlpha ?? false,
    pristine: documentMeta().pristine,
  }), [documentMeta])

  const setHasAlpha = useCallback((hasAlpha: boolean) => {
    if (imageRef.current) imageRef.current = { ...imageRef.current, hasAlpha }
    setImage((current) => current && current.hasAlpha !== hasAlpha ? { ...current, hasAlpha } : current)
  }, [])

  const recordEntry = useCallback((entry: HistoryEntry) => {
    pushEntry({ undo: undoRef.current, redo: redoRef.current }, entry)
    setHistoryVersion((version) => version + 1)
  }, [])

  /** After an edit that changed the canvas size: dimensions, crop box and Fit follow. */
  const afterResize = useCallback((canvas: HTMLCanvasElement) => {
    setDimensions({ width: canvas.width, height: canvas.height })
    refitSoon(canvas.width, canvas.height)
  }, [refitSoon])

  // #region advanced-integration (WP1)
  // Host side of Advanced mode (src/advanced/useAdvancedMode.tsx). Outside this region WP1 only touches
  // the call sites marked "advanced-integration": open (PSD), proceed prompts, Save, Print, Export,
  // Copy, keyboard delegation and the toolbar/menus. settleSimple and adoptFlattened are the Simple
  // hooks Advanced relies on; Simple-mode work (WP8) extends them in place.
  const confirmOpenRef = useRef(false)
  confirmOpenRef.current = Boolean(confirmState)
  // Save/Print/Export are declared further down; the editor reaches them through this late-bound ref.
  const hostActionsRef = useRef<{ save: (forceDialog: boolean) => Promise<boolean>; print: () => void; openExportMenu: () => void }>({
    save: async () => false,
    print: () => {},
    openExportMenu: () => {},
  })
  const revisions = useMemo<RevisionClock>(() => ({
    next: () => {
      nextRevisionRef.current += 1
      return nextRevisionRef.current
    },
    current: () => contentRevisionRef.current,
    set: (revision) => {
      contentRevisionRef.current = revision
      reflectDirty()
    },
    saved: () => savedRevisionRef.current,
  }), [reflectDirty])
  const updateImage = useCallback((patch: Partial<OpenImage>) => {
    setImage((current) => current ? { ...current, ...patch } : current)
  }, [])
  /** Ends Simple sessions before Advanced reads the canvas (WP8: stroke, markup, adjustments, crop box). */
  const settleSimple = useCallback(async () => {
    await simpleSettleRef.current()
  }, [])
  /** Puts flattened Advanced output into the Simple canvas; with pushUndoStep, one undo restores the pre-Advanced pixels and revision (WP8 history). */
  const adoptFlattened = useCallback((flattened: HTMLCanvasElement, beforeRevision: number, pushUndoStep: boolean) => {
    simpleAdoptRef.current(flattened, beforeRevision, pushUndoStep)
  }, [])
  const advanced = useAdvancedMode({
    canvasRef,
    getImage: () => imageRef.current,
    updateImage,
    revisions,
    notify,
    isSuspended: () => printDialogOpenRef.current || confirmOpenRef.current,
    save: (forceDialog) => hostActionsRef.current.save(forceDialog),
    openExportMenu: () => hostActionsRef.current.openExportMenu(),
    print: () => hostActionsRef.current.print(),
    settleSimple,
    adoptFlattened,
  })
  /** A PSD opens straight into Advanced. It is parsed before any session state changes, so a failure keeps the current image. */
  const openLayeredPayload = useCallback(async (payload: ImagePayload, request: number) => {
    const opened = await advanced.importPsdPayload(payload)
    if (!opened || request !== loadRequestRef.current) return
    const { document, issues } = opened
    originalImageRef.current = { blob: new Blob([payload.data as BlobPart], { type: payload.mime }), format: 'psd', printable: false }
    documentIdRef.current += 1
    nextRevisionRef.current = 0
    contentRevisionRef.current = 0
    savedRevisionRef.current = 0
    setImage({
      name: payload.name,
      path: payload.directSave ? payload.path : null,
      size: payload.size,
      sourceFormat: payload.format,
      saveFormat: 'psd',
      mime: payload.mime,
      hasAlpha: pixelsHaveAlpha(document.composite),
      webpLossless: false,
      psdIssues: issues,
    })
    setAlbum({ paths: payload.path ? [payload.path] : [], index: payload.path ? 0 : -1 })
    if (payload.path) {
      const currentPath = payload.path
      const loadedDocument = documentIdRef.current
      void window.simpleImage.listSiblings(currentPath).then((paths) => {
        if (loadedDocument !== documentIdRef.current || !paths.length) return
        const index = paths.findIndex((entry) => entry.toLowerCase() === currentPath.toLowerCase())
        setAlbum(index >= 0 ? { paths, index } : { paths: [currentPath, ...paths], index: 0 })
      }).catch(() => {})
    }
    setDimensions({ width: document.width, height: document.height })
    dirtyRef.current = false
    setDirtyState(false)
    setTool('view')
    setCropRect(null)
    undoRef.current = []
    redoRef.current = []
    setHistoryVersion((version) => version + 1)
    window.simpleImage.setTitle(payload.name)
    advanced.enterWithDocument(document, payload.name)
    notify(`${payload.name} opened in Advanced mode.`)
  }, [advanced.enterWithDocument, advanced.importPsdPayload, notify])
  // #endregion advanced-integration

  // ---------------------------------------------------------------------------------------------
  // History (WP8: byte-budgeted, tile patches, exact inverse geometry)
  // ---------------------------------------------------------------------------------------------

  const applyHistoryState = useCallback((state: HistoryState) => {
    contentRevisionRef.current = state.revision
    setPristine(state.pristine)
    setHasAlpha(state.hasAlpha)
    reflectDirty()
  }, [reflectDirty, setHasAlpha, setPristine])

  const resetCropSession = useCallback(() => {
    cropDragRef.current = null
    setCropRect(null)
    setStraighten(0)
    setLeveling(false)
    setCropAdjusted(false)
  }, [setCropAdjusted])

  const stepHistory = useCallback((direction: 'undo' | 'redo') => {
    if (strokeRef.current || workingRef.current) return
    const canvas = canvasRef.current
    if (!canvas) return
    const before = { width: canvas.width, height: canvas.height }
    let entry: HistoryEntry | null
    try {
      const stacks = { undo: undoRef.current, redo: redoRef.current }
      entry = direction === 'undo' ? undoEntry(stacks, createCanvasSurface(canvas)) : redoEntry(stacks, createCanvasSurface(canvas))
    } catch (error) {
      notify(messageOf(error, 'That step could not be restored.'), 'error')
      return
    }
    if (!entry) return
    applyHistoryState(direction === 'undo' ? entry.before : entry.after)
    lastStrokeEndRef.current = null
    if (tool === 'crop') {
      resetCropSession()
      setTool('view')
    }
    if (before.width !== canvas.width || before.height !== canvas.height || entry.kind !== 'patch') {
      setDimensions({ width: canvas.width, height: canvas.height })
      refitSoon(canvas.width, canvas.height)
    }
    setHistoryVersion((version) => version + 1)
  }, [applyHistoryState, notify, refitSoon, resetCropSession, tool])

  const undoCanvas = useCallback(() => stepHistory('undo'), [stepHistory])
  const redoCanvas = useCallback(() => stepHistory('redo'), [stepHistory])

  // ---------------------------------------------------------------------------------------------
  // Brush and eraser strokes (IMAGE-SIE-1: every draw is part of a recorded, revisioned step)
  // ---------------------------------------------------------------------------------------------

  /** Closes the open history step of the stroke (the pointer may stay down: the next draw opens a new step). */
  const commitStrokeSegment = useCallback(() => {
    const stroke = strokeRef.current
    if (!stroke?.recorder) return
    const recorder = stroke.recorder
    stroke.recorder = null
    let hasAlpha = stroke.hadAlpha
    if (stroke.erasing) hasAlpha = true
    else if (stroke.hadAlpha) {
      const canvas = canvasRef.current
      hasAlpha = canvas ? canvasHasTransparency(canvas) : true
    }
    stroke.hadAlpha = hasAlpha
    setHasAlpha(hasAlpha)
    const entry = recorder.finish(stroke.erasing ? 'Eraser' : 'Brush', stroke.before, { revision: contentRevisionRef.current, hasAlpha, pristine: false })
    if (entry) recordEntry(entry)
  }, [recordEntry, setHasAlpha])

  const endStroke = useCallback(() => {
    const stroke = strokeRef.current
    if (!stroke) return
    commitStrokeSegment()
    strokeRef.current = null
    paintingRef.current = false
    lastPointRef.current = null
    lastStrokeEndRef.current = { documentId: documentIdRef.current, point: stroke.last }
    try { if (stroke.element.hasPointerCapture(stroke.pointerId)) stroke.element.releasePointerCapture(stroke.pointerId) } catch { /* pointer already gone */ }
  }, [commitStrokeSegment])

  const drawStrokeSegment = useCallback((stroke: StrokeState, from: Point, to: Point, width: number) => {
    const canvas = canvasRef.current
    if (!canvas) return
    if (!stroke.recorder || contentRevisionRef.current !== stroke.revision) {
      if (stroke.recorder) commitStrokeSegment()
      stroke.recorder = createPatchRecorder(createCanvasSurface(canvas))
      stroke.before = currentState()
      markEdited()
      stroke.revision = contentRevisionRef.current
    }
    const pad = width / 2 + 2
    const left = Math.floor(Math.min(from.x, to.x) - pad)
    const top = Math.floor(Math.min(from.y, to.y) - pad)
    stroke.recorder.touch({ x: left, y: top, width: Math.ceil(Math.max(from.x, to.x) + pad) - left, height: Math.ceil(Math.max(from.y, to.y) + pad) - top })
    const context = imageContext(canvas)
    context.save()
    context.globalCompositeOperation = stroke.erasing ? 'destination-out' : 'source-over'
    context.strokeStyle = stroke.color
    context.fillStyle = stroke.color
    context.lineWidth = width
    context.lineCap = 'round'
    context.lineJoin = 'round'
    context.beginPath()
    context.moveTo(from.x, from.y)
    context.lineTo(to.x, to.y)
    context.stroke()
    if (from.x === to.x && from.y === to.y) {
      context.beginPath()
      context.arc(to.x, to.y, width / 2, 0, Math.PI * 2)
      context.fill()
    }
    context.restore()
  }, [commitStrokeSegment, currentState, markEdited])

  // ---------------------------------------------------------------------------------------------
  // Geometry: rotate, flip, crop, straighten, resize
  // ---------------------------------------------------------------------------------------------

  /** Exact quarter turns and flips: zero-byte history steps (undo applies the inverse). */
  const applyGeometryNow = useCallback((op: GeometryOp, label: string) => {
    const canvas = canvasRef.current
    if (!canvas || !imageRef.current) return
    const before = currentState()
    transformCanvas(canvas, op)
    markEdited()
    recordEntry(geometryEntry(label, op, before, { ...currentState(), pristine: false }))
    afterResize(canvas)
  }, [afterResize, currentState, markEdited, recordEntry])

  const finishWork = useCallback(<T,>(label: string, work: (progress: (fraction: number) => void) => Promise<T>): Promise<T | null> => {
    if (workingRef.current) return Promise.resolve(null)
    const documentId = documentIdRef.current
    setWorking({ label, progress: 0 })
    let last = 0
    const progress = (fraction: number) => {
      const now = performance.now()
      if (now - last < 80 && fraction < 1) return
      last = now
      setWorkingState({ label, progress: clamp(fraction, 0, 1) })
    }
    return work(progress)
      .then((value) => (documentId === documentIdRef.current ? value : null))
      .catch((error: unknown) => {
        if ((error as Error)?.name !== 'AbortError') notify(messageOf(error, `${label} failed.`), 'error')
        return null
      })
      .finally(() => setWorking(null))
  }, [notify, setWorking])

  /** Whole-image pixel step: the before image is stored, the new pixels replace the canvas. */
  const replaceWithResult = useCallback((label: string, before: PixelBuffer, beforeState: HistoryState, result: PixelBuffer, hasAlpha: boolean) => {
    const canvas = canvasRef.current
    if (!canvas) return
    const sizeChanged = result.width !== canvas.width || result.height !== canvas.height
    if (sizeChanged) replaceCanvasPixels(canvas, result)
    else putBuffer(canvas, result, 0, 0)
    markEdited()
    setHasAlpha(hasAlpha)
    recordEntry(snapshotEntry(label, before, beforeState, { revision: contentRevisionRef.current, hasAlpha, pristine: false }))
    if (sizeChanged) afterResize(canvas)
  }, [afterResize, markEdited, recordEntry, setHasAlpha])

  const cropFrameSize = useCallback((angle = straighten): Dimensions => cropFrame(dimensions, angle), [dimensions, straighten])

  /** The crop box a straighten angle / aspect starts with: auto-crop inside the turned image. */
  const autoCropBox = useCallback((angle: number, ratio: number | null): CropRect => {
    if (!angle) {
      if (ratio) return fitAspect(defaultCropRect(dimensions), ratio, dimensions)
      return { x: 0, y: 0, width: dimensions.width, height: dimensions.height }
    }
    return inscribedCrop(dimensions, angle, ratio)
  }, [dimensions])

  const activeRatio = useCallback((aspect = cropAspect, portrait = cropPortrait) => presetRatio(aspect, dimensions, portrait), [cropAspect, cropPortrait, dimensions])

  const cancelCrop = useCallback(() => {
    resetCropSession()
    setTool((current) => (current === 'crop' ? 'view' : current))
  }, [resetCropSession])

  /** Applies the crop box (and the straighten angle) as one undo step. Resolves true when applied. */
  const applyCrop = useCallback(async (): Promise<boolean> => {
    const canvas = canvasRef.current
    const box = cropRect
    if (!canvas || !box || tool !== 'crop') return false
    const angle = straighten
    if (!angle) {
      const rect = toIntegerRect(box, { width: canvas.width, height: canvas.height })
      if (rect.x === 0 && rect.y === 0 && rect.width === canvas.width && rect.height === canvas.height) {
        cancelCrop()
        return true
      }
      const beforeState = currentState()
      const before = readRegion(canvas, { x: 0, y: 0, width: canvas.width, height: canvas.height })
      cropCanvas(canvas, rect)
      markEdited()
      const hasAlpha = beforeState.hasAlpha ? canvasHasTransparency(canvas) : false
      setHasAlpha(hasAlpha)
      recordEntry(snapshotEntry('Crop', before, beforeState, { revision: contentRevisionRef.current, hasAlpha, pristine: false }))
      resetCropSession()
      setTool('view')
      afterResize(canvas)
      return true
    }
    const applied = await finishWork('Straightening', async (progress) => {
      const beforeState = currentState()
      const before = readCanvasPixels(canvas)
      const turned = await getImagingClient().run('rotate', {
        src: { width: before.width, height: before.height, data: new Uint8ClampedArray(before.data) },
        degrees: angle,
        fit: 'expand',
        interpolation: 'bicubic',
        background: { r: 0, g: 0, b: 0, a: 0 },
      }, { onProgress: progress })
      const rect = toIntegerRect(box, { width: turned.width, height: turned.height })
      const result = cropBuffer(turned, rect)
      replaceWithResult('Straighten', before, beforeState, result, hasTransparency(result))
      return true
    })
    if (!applied) return false
    resetCropSession()
    setTool('view')
    return true
  }, [afterResize, cancelCrop, cropRect, currentState, finishWork, markEdited, recordEntry, replaceWithResult, resetCropSession, setHasAlpha, straighten, tool])

  const beginCropNow = useCallback(() => {
    if (!imageRef.current || !dimensions.width || !dimensions.height) return
    setTool('crop')
    setStraighten(0)
    setLeveling(false)
    setCropAspect('free')
    setCropPortrait(defaultPortrait(dimensions))
    setCropAdjusted(false)
    setCropRect(defaultCropRect(dimensions))
  }, [dimensions, setCropAdjusted])

  const chooseAspect = useCallback((preset: AspectPreset) => {
    const portrait = preset === cropAspect ? cropPortrait : defaultPortrait(dimensions)
    setCropAspect(preset)
    setCropPortrait(portrait)
    const ratio = presetRatio(preset, dimensions, portrait)
    if (!ratio) return
    const frame = cropFrameSize()
    setCropRect((current) => (straighten ? inscribedCrop(dimensions, straighten, ratio) : fitAspect(current ?? defaultCropRect(frame), ratio, frame)))
    setCropAdjusted(true)
  }, [cropAspect, cropFrameSize, cropPortrait, dimensions, setCropAdjusted, straighten])

  const swapAspect = useCallback(() => {
    if (!presetCanSwap(cropAspect)) return
    const portrait = !cropPortrait
    setCropPortrait(portrait)
    const ratio = presetRatio(cropAspect, dimensions, portrait)
    if (!ratio) return
    const frame = cropFrameSize()
    setCropRect((current) => (straighten ? inscribedCrop(dimensions, straighten, ratio) : fitAspect(current ?? defaultCropRect(frame), ratio, frame)))
    setCropAdjusted(true)
  }, [cropAspect, cropFrameSize, cropPortrait, dimensions, setCropAdjusted, straighten])

  const changeStraighten = useCallback((degrees: number) => {
    setStraighten(degrees)
    setCropRect(autoCropBox(degrees, activeRatio()))
    setCropAdjusted(degrees !== 0 || cropAdjustedRef.current)
  }, [activeRatio, autoCropBox, setCropAdjusted])

  const nudgeCrop = useCallback((dx: number, dy: number) => {
    if (!cropRect) return
    const frame = cropFrameSize()
    const next = straighten
      ? constrainDragInside(cropRect, 'move', dx, dy, { ratio: null, fromCenter: false, bounds: frame }, (rect) => insideRotated(rect, dimensions, straighten))
      : nudgeRect(cropRect, dx, dy, frame)
    setCropRect(next)
    setCropAdjusted(true)
  }, [cropFrameSize, cropRect, dimensions, setCropAdjusted, straighten])

  const applyResize = useCallback(async (width: number, height: number, method: ResampleMethod) => {
    const canvas = canvasRef.current
    if (!canvas) return
    const done = await finishWork('Resizing', async (progress) => {
      const beforeState = currentState()
      const before = readCanvasPixels(canvas)
      const result = await getImagingClient().run('resample', {
        src: { width: before.width, height: before.height, data: new Uint8ClampedArray(before.data) },
        width,
        height,
        method,
      }, { onProgress: progress })
      replaceWithResult('Resize', before, beforeState, result, beforeState.hasAlpha ? hasTransparency(result) : false)
      return true
    })
    if (done) {
      setResizeOpen(false)
      notify(`Resized to ${width.toLocaleString()} × ${height.toLocaleString()} px.`)
    }
  }, [currentState, finishWork, notify, replaceWithResult])

  // ---------------------------------------------------------------------------------------------
  // Adjust and Looks
  // ---------------------------------------------------------------------------------------------

  const closeAdjust = useCallback(() => {
    setSimplePanel('details')
    setPendingAdjust(NEUTRAL_ADJUST)
    setComparing(false)
  }, [])

  /** Applies the pending adjustments at full resolution in the worker: one undo step, alpha untouched. */
  const commitAdjust = useCallback(async (pending: PendingAdjust): Promise<boolean> => {
    const canvas = canvasRef.current
    if (!canvas || isNeutralAdjust(pending)) {
      closeAdjust()
      return true
    }
    const done = await finishWork('Applying adjustments', async (progress) => {
      const beforeState = currentState()
      const before = readCanvasPixels(canvas)
      const look = hasLook(pending)
      const quick = !isNeutralQuick(pending.quick)
      let result: PixelBuffer = { width: before.width, height: before.height, data: new Uint8ClampedArray(before.data) }
      const client = getImagingClient()
      if (quick) {
        result = await client.run('adjust', { src: result, specs: [quickSpec(pending.quick)], mask: null, opacity: 1 }, { onProgress: (fraction) => progress(look ? fraction / 2 : fraction) })
      }
      if (look) {
        result = await client.run('look', { src: result, look: pending.look, intensity: pending.intensity }, { onProgress: (fraction) => progress(quick ? 0.5 + fraction / 2 : fraction) })
      }
      replaceWithResult(look && !quick ? 'Look' : 'Adjust', before, beforeState, result, beforeState.hasAlpha)
      return true
    })
    if (done) closeAdjust()
    return Boolean(done)
  }, [closeAdjust, currentState, finishWork, replaceWithResult])

  // ---------------------------------------------------------------------------------------------
  // Markup baking
  // ---------------------------------------------------------------------------------------------

  const releasePastedImages = useCallback(() => {
    for (const url of pastedImagesRef.current.keys()) URL.revokeObjectURL(url)
    pastedImagesRef.current.clear()
  }, [])

  /** Draws the markup into the image (shared/vector.ts, like Advanced layers): one undo step of tile patches. */
  const bakeMarkup = useCallback(async (): Promise<boolean> => {
    if (markup.current().editingId) markup.finishEditing()
    const items = paintableItems(markup.current().history.present)
    const canvas = canvasRef.current
    if (!items.length || !canvas) {
      markup.reset()
      releasePastedImages()
      return true
    }
    await Promise.all(items.map((item) => item.kind === 'text'
      ? Promise.resolve(document.fonts?.load(cssFont(textStyleOf(item)), item.text)).catch(() => undefined)
      : Promise.resolve(undefined)))
    const beforeState = currentState()
    const bounds = markupBounds(items)
    const area = bounds ? clipRect({ x: Math.floor(bounds.x), y: Math.floor(bounds.y), width: Math.ceil(bounds.width) + 2, height: Math.ceil(bounds.height) + 2 }, canvas.width, canvas.height) : null
    if (!area) {
      markup.reset()
      releasePastedImages()
      return true
    }
    const recorder = createPatchRecorder(createCanvasSurface(canvas))
    recorder.touch(area)
    const context = imageContext(canvas)
    for (const item of items) {
      if (item.kind === 'text') drawText(context, toTextSpec(item))
      else if (item.kind === 'image') {
        const element = pastedImagesRef.current.get(item.src)
        if (!element) continue
        context.save()
        context.imageSmoothingEnabled = true
        context.imageSmoothingQuality = 'high'
        context.drawImage(element, item.x, item.y, item.width, item.height)
        context.restore()
      } else drawShape(context, toShapeSpec(item))
    }
    markEdited()
    const hasAlpha = beforeState.hasAlpha ? canvasHasTransparency(canvas) : false
    setHasAlpha(hasAlpha)
    const entry = recorder.finish('Markup', beforeState, { revision: contentRevisionRef.current, hasAlpha, pristine: false })
    if (entry) recordEntry(entry)
    markup.reset()
    releasePastedImages()
    return true
  }, [currentState, markEdited, markup, recordEntry, releasePastedImages, setHasAlpha])

  // ---------------------------------------------------------------------------------------------
  // Pending sessions (IMAGE-SIE-1/2: nothing in progress is ever silently dropped)
  // ---------------------------------------------------------------------------------------------

  /**
   * Flushes every in-progress edit into the image before Save, Export, Print, Copy, a prompt, a tool switch
   * or Advanced: the brush stroke (kept alive for Save/Export so painting may continue as a new step),
   * the markup (baked) and the Adjust panel (applied). The crop box is handled by the callers.
   */
  const commitPendingEdits = useCallback(async (options: { keepStrokeAlive?: boolean } = {}): Promise<boolean> => {
    if (strokeRef.current) {
      if (options.keepStrokeAlive) commitStrokeSegment()
      else endStroke()
    }
    let complete = true
    if (tool === 'markup' || markup.hasItems) complete = (await bakeMarkup()) && complete
    if (simplePanel === 'adjust') complete = (await commitAdjust(pendingAdjust)) && complete
    return complete
  }, [bakeMarkup, commitAdjust, commitStrokeSegment, endStroke, markup.hasItems, pendingAdjust, simplePanel, tool])
  const commitPendingRef = useRef(commitPendingEdits)
  commitPendingRef.current = commitPendingEdits

  /** True while something unapplied would be lost by closing (markup items, pending adjustments). */
  const hasPendingSession = useCallback(() => markup.hasItems || (simplePanel === 'adjust' && !isNeutralAdjust(pendingAdjust)), [markup.hasItems, pendingAdjust, simplePanel])
  const pendingSession = hasPendingSession()
  const pendingSessionRef = useRef(pendingSession)
  pendingSessionRef.current = pendingSession

  simpleSettleRef.current = async () => {
    if (!(await commitPendingRef.current())) throw new Error('Finish or cancel the open edit first.')
    cancelCrop()
    setTool((current) => (current === 'markup' || current === 'eyedropper' ? 'view' : current))
  }

  simpleAdoptRef.current = (flattened, beforeRevision, pushUndoStep) => {
    const canvas = canvasRef.current
    if (!canvas) throw new Error('The image canvas is unavailable.')
    const beforeState: HistoryState = { ...currentState(), revision: beforeRevision }
    const before = pushUndoStep ? readRegion(canvas, { x: 0, y: 0, width: canvas.width, height: canvas.height }) : null
    canvas.width = flattened.width
    canvas.height = flattened.height
    const context = imageContext(canvas)
    context.clearRect(0, 0, canvas.width, canvas.height)
    context.drawImage(flattened, 0, 0)
    setPristine(false)
    if (before) {
      const hasAlpha = canvasHasTransparency(canvas)
      pushEntry({ undo: undoRef.current, redo: redoRef.current }, snapshotEntry('Advanced edits', before, beforeState, { revision: contentRevisionRef.current, hasAlpha, pristine: false }))
    }
    lastStrokeEndRef.current = null
    setDimensions({ width: canvas.width, height: canvas.height })
    setCropRect(null)
    setHistoryVersion((version) => version + 1)
    reflectDirty()
    refitSoon(canvas.width, canvas.height)
  }

  // ---------------------------------------------------------------------------------------------
  // Open
  // ---------------------------------------------------------------------------------------------

  const resetSimpleSessions = useCallback(() => {
    strokeRef.current = null
    paintingRef.current = false
    lastPointRef.current = null
    lastStrokeEndRef.current = null
    markup.reset()
    releasePastedImages()
    resetCropSession()
    setSimplePanel('details')
    setPendingAdjust(NEUTRAL_ADJUST)
    setComparing(false)
    setResizeOpen(false)
  }, [markup, releasePastedImages, resetCropSession])

  const loadPayloadNow = useCallback(async (payload: ImagePayload, request: number) => {
    if (request !== loadRequestRef.current) return
    // advanced-integration (WP1): layered Photoshop documents open in Advanced mode.
    if (payload.format === 'psd') return openLayeredPayload(payload, request)
    const extras = payload as ImagePayload & OpenPayloadExtras
    const blob = new Blob([payload.data as BlobPart], { type: payload.mime })
    const url = URL.createObjectURL(blob)
    try {
      const decoded = new Image()
      decoded.decoding = 'async'
      decoded.src = url
      await decoded.decode()
      if (request !== loadRequestRef.current) return
      if (!decoded.naturalWidth || !decoded.naturalHeight) throw new Error('The image has no drawable pixels.')
      if (
        decoded.naturalWidth > MAX_CANVAS_DIMENSION
        || decoded.naturalHeight > MAX_CANVAS_DIMENSION
        || decoded.naturalWidth * decoded.naturalHeight > MAX_CANVAS_PIXELS
      ) {
        throw new Error('This image is too large to edit safely. The limit is 50 megapixels and 20,000 pixels per side.')
      }
      const canvas = canvasRef.current
      if (!canvas) throw new Error('The image canvas is unavailable.')
      resetSimpleSessions()
      canvas.width = decoded.naturalWidth
      canvas.height = decoded.naturalHeight
      const context = context2d()
      context.clearRect(0, 0, canvas.width, canvas.height)
      context.drawImage(decoded, 0, 0)
      // JPEG cannot contain transparency. Avoid a full pixel readback and scan
      // just to rediscover this on every large photograph.
      const hasAlpha = /^(jpe?g)$/i.test(payload.format) ? false : canvasHasTransparency(canvas)
      const animated = Boolean(extras.animated)
      // IMAGE-SIE-3: an animated source is never flattened in place; Save writes a still PNG copy.
      const saveFormat = animated ? 'png' : defaultSaveFormat(payload.format)
      const hints = encodeHintsFor(payload.format, payload.data)
      originalImageRef.current = {
        blob,
        format: defaultSaveFormat(payload.format),
        printable: /^(png|jpe?g)$/i.test(payload.format),
      }
      documentIdRef.current += 1
      docMetaRef.current = {
        documentId: documentIdRef.current,
        pristine: true,
        sourceToken: typeof extras.sourceToken === 'string' ? extras.sourceToken : null,
        animated,
        hints,
      }
      nextRevisionRef.current = 0
      contentRevisionRef.current = 0
      savedRevisionRef.current = 0
      const nextImage: OpenImage = {
        name: payload.name,
        path: payload.directSave && !animated ? payload.path : null,
        size: payload.size,
        sourceFormat: payload.format,
        saveFormat,
        mime: payload.mime,
        hasAlpha,
        webpLossless: hints.webpLossless,
        psdIssues: [],
      }
      imageRef.current = nextImage
      setImage(nextImage)
      setAlbum({ paths: payload.path ? [payload.path] : [], index: payload.path ? 0 : -1 })
      if (payload.path) {
        const currentPath = payload.path
        const loadedDocument = documentIdRef.current
        void window.simpleImage.listSiblings(currentPath).then((paths) => {
          if (loadedDocument !== documentIdRef.current || !paths.length) return
          const index = paths.findIndex((entry) => entry.toLowerCase() === currentPath.toLowerCase())
          setAlbum(index >= 0 ? { paths, index } : { paths: [currentPath, ...paths], index: 0 })
        }).catch(() => {})
      }
      setDimensions({ width: canvas.width, height: canvas.height })
      dirtyRef.current = false
      setDirtyState(false)
      setTool('view')
      setCropRect(null)
      undoRef.current = []
      redoRef.current = []
      setHistoryVersion((version) => version + 1)
      window.simpleImage.setTitle(payload.name)
      // advanced-integration (WP1): the new image replaces a document that was open in Advanced.
      advanced.discard()
      fitModeRef.current = true
      requestAnimationFrame(() => { if (fitModeRef.current) fitImage(canvas.width, canvas.height) })
      const notices = Array.isArray(extras.notices) ? extras.notices.filter((notice) => typeof notice === 'string' && notice) : []
      if (animated && !notices.some((notice) => /animated/i.test(notice))) notices.unshift('Animated image: only the first frame is editable. Save creates a still copy.')
      if (notices.length) notify(notices.join(' '))
      else if (payload.format === 'gif') notify('GIF opened as an editable first frame.')
      else notify(`${payload.format.toUpperCase()} opened locally.`)
    } finally {
      URL.revokeObjectURL(url)
    }
  }, [advanced.discard, context2d, fitImage, notify, openLayeredPayload, resetSimpleSessions])

  const requestProceed = useCallback((title: string, detail: string, action: () => void | Promise<void>) => {
    if (printDialogOpenRef.current) return
    const proceed = () => {
      const cropPending = cropAdjustedRef.current && Boolean(imageRef.current)
      // An edit that could not be applied (pendingSessionRef) still counts: never close over it silently.
      if (!dirtyRef.current && !cropPending && !pendingSessionRef.current) {
        void action()
        return
      }
      setConfirmState({ title, detail: cropPending ? `${detail} The crop you set up has not been applied yet.` : detail, action, cropPending })
    }
    // WP8: commit the stroke, markup and adjustments first so the prompt sees every edit.
    void commitPendingRef.current().catch(() => false).then(() => {
      // advanced-integration (WP1): commit open Advanced sessions (text, transform) first so the prompt
      // sees every edit; if that fails, ask rather than risk discarding work.
      if (advanced.mode === 'advanced') void advanced.settle().then(proceed, () => setConfirmState({ title, detail, action }))
      else proceed()
    })
  }, [advanced.mode, advanced.settle])

  const loadPayload = useCallback((payload: ImagePayload, request: number) => {
    if (request !== loadRequestRef.current) return
    requestProceed('Save changes before opening another image?', 'Your current edits have not been saved.', async () => {
      try { await loadPayloadNow(payload, request) } catch (error) { if (request === loadRequestRef.current) notify(error instanceof Error ? error.message : 'The image could not be opened.', 'error') }
    })
  }, [loadPayloadNow, notify, requestProceed])

  const openPath = useCallback(async (filePath: string) => {
    if (printDialogOpenRef.current) return
    const request = ++loadRequestRef.current
    try { loadPayload(await window.simpleImage.openPath(filePath), request) }
    catch (error) { if (request === loadRequestRef.current) notify(error instanceof Error ? error.message : 'The image could not be opened.', 'error') }
  }, [loadPayload, notify])

  // The album cursor moves even when a file fails to open, so Next/Previous step past an
  // unreadable sibling instead of retrying it forever. A successful open re-syncs it.
  const albumCursorRef = useRef(-1)
  useEffect(() => { albumCursorRef.current = album.index }, [album.index, album.paths])
  const stepAlbum = useCallback((delta: number) => {
    if (album.paths.length < 2 || album.index < 0) return
    const from = albumCursorRef.current >= 0 && albumCursorRef.current < album.paths.length ? albumCursorRef.current : album.index
    const next = (from + delta + album.paths.length) % album.paths.length
    albumCursorRef.current = next
    void openPath(album.paths[next])
  }, [album, openPath])

  const openDialog = useCallback(async () => {
    if (printDialogOpenRef.current) return
    const request = ++loadRequestRef.current
    try {
      const payload = await window.simpleImage.openFile()
      if (payload) loadPayload(payload, request)
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be opened.', 'error')
    }
  }, [loadPayload, notify])

  const openDroppedFile = useCallback(async (file: File) => {
    if (printDialogOpenRef.current) return
    const request = ++loadRequestRef.current
    try {
      const extension = file.name.split('.').pop()?.toLowerCase() || ''
      if (!SUPPORTED.includes(extension)) throw new Error(`.${extension || '?'} images are not supported.`)
      const filePath = window.simpleImage.pathForFile(file)
      const payload = filePath
        ? await window.simpleImage.openPath(filePath)
        : await window.simpleImage.openBytes(file.name, await file.arrayBuffer())
      loadPayload(payload, request)
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The dropped image could not be opened.', 'error')
    }
  }, [loadPayload, notify])

  /** A small modal with explicit choices; resolves with the chosen value or 'cancel'. */
  const askChoice = useCallback((title: string, detail: string, tone: ChoiceState['tone'], options: ChoiceOption[]) => new Promise<string>((resolve) => {
    setChoiceState({ title, detail, tone, options, resolve: (value) => { setChoiceState(null); resolve(value) } })
  }), [])

  // ---------------------------------------------------------------------------------------------
  // Save, Print, Export, Copy
  // ---------------------------------------------------------------------------------------------

  const saveImage = useCallback(async (forceDialog = false, formatOverride?: SaveFormat): Promise<boolean> => {
    const canvas = canvasRef.current
    const current = imageRef.current
    // Defect 13: a ref guards re-entrancy (a held Ctrl+S can fire twice before React re-renders).
    if (!canvas || !current || savingRef.current || exporting || printing || workingRef.current) return false
    savingRef.current = true
    setSaving(true)
    setSaveMenu(false)
    try {
      // IMAGE-SIE-1: the stroke so far, markup and adjustments are part of what gets saved.
      if (!(await commitPendingRef.current({ keepStrokeAlive: true }))) return false
      // IMAGE-SIE-2: an adjusted crop box is not silently ignored.
      if (advanced.mode !== 'advanced' && cropAdjustedRef.current) {
        const choice = await askChoice(
          'Apply the crop before saving?',
          'You set up a crop that has not been applied yet.',
          'crop',
          [{ value: 'cancel', label: 'Cancel' }, { value: 'without', label: 'Save without crop' }, { value: 'apply', label: 'Apply crop and save', primary: true }],
        )
        if (choice === 'cancel') return false
        if (choice === 'apply' && !(await applyCrop())) return false
      }
      // advanced-integration (WP1): commit open Advanced sessions first; a layered document saves as
      // PSD (Save As when the file is not a .psd), and pixels come from the flattened Advanced output.
      if (advanced.mode === 'advanced') await advanced.settle()
      const documentId = documentIdRef.current
      let saveFormat = formatOverride ?? advanced.effectiveSaveFormat()
      let showDialog = forceDialog || !current.path || saveFormat !== current.saveFormat
      // IMAGE-SIE-14: JPEG cannot keep transparency; say so before writing white.
      if (saveFormat === 'jpeg' && advanced.mode !== 'advanced' && imageRef.current?.hasAlpha) {
        const choice = await askChoice(
          'JPEG can’t keep transparency',
          'Transparent and erased areas will be saved as white. Save a PNG copy to keep them.',
          'warning',
          [{ value: 'cancel', label: 'Cancel' }, { value: 'jpeg', label: 'Save anyway' }, { value: 'png', label: 'Save as PNG…', primary: true }],
        )
        if (choice === 'cancel') return false
        if (choice === 'png') {
          saveFormat = 'png'
          showDialog = true
        }
      }
      const revision = contentRevisionRef.current
      const flatCopy = advanced.wouldLoseLayers(saveFormat)
      const original = originalImageRef.current
      const meta = documentMeta()
      // An unchanged Save/Save As must not recompress a JPEG or discard its metadata. Undo back to the
      // opened pixels also restores the exact source (IMAGE-SIE-1: an explicit flag, not "revision 0").
      const unchanged = isPristine() && original?.format === saveFormat
        && original.blob.type === outputMime(saveFormat)
      if (saveFormat === 'psd' && !showDialog && !unchanged) {
        const choice = await advanced.confirmPsdOverwrite(current)
        if (choice === 'cancel') return false
        if (choice === 'copy') showDialog = true
      }
      const quality = saveQuality(saveFormat, meta.hints)
      const data = unchanged && original
        ? new Uint8Array(await original.blob.arrayBuffer())
        : saveFormat === 'psd'
          ? await advanced.encodePsd()
          : new Uint8Array(await (await advanced.withOutputCanvas((output) => canvasToBlob(output, saveFormat, quality))).arrayBuffer())
      const input: SaveImageInput & { sourceToken?: string } = {
        data,
        path: showDialog ? current.path : current.path,
        name: current.name,
        format: saveFormat,
        forceDialog: showDialog,
      }
      if (meta.sourceToken) input.sourceToken = meta.sourceToken
      const result = await window.simpleImage.saveImage(input) as (SavedImage & { warnings?: string[]; stillOfAnimation?: boolean }) | null
      if (!result) return false
      if (documentIdRef.current === documentId) {
        if (flatCopy) {
          // advanced-integration (WP1): the file holds a flattened copy and the layers exist only in this
          // session, so the document keeps its path and stays Modified (closing still asks).
          notify(`Saved a flattened copy as ${result.name}. The layers are not in that file; save as Photoshop to keep them.`)
          return false
        }
        const keepsAnimation = meta.animated && saveFormat === 'png' && !result.stillOfAnimation && unchanged
        setImage((value) => value ? {
          ...value,
          name: result.name,
          path: result.path,
          size: result.size,
          sourceFormat: result.format,
          saveFormat,
          mime: outputMime(saveFormat),
          webpLossless: saveFormat === 'webp' ? value.webpLossless : false,
          ...(saveFormat === 'psd' ? { psdIssues: [] } : {}),
        } : value)
        if (meta.animated && !keepsAnimation) meta.animated = false
        savedRevisionRef.current = revision
        reflectDirty()
        const fullySaved = contentRevisionRef.current === revision
        const warning = Array.isArray(result.warnings) && result.warnings.length ? ` ${result.warnings.join(' ')}` : ''
        const still = result.stillOfAnimation ? ' as a still copy of the first frame' : ''
        notify(fullySaved
          ? `Saved ${result.name}${still}${warning ? `.${warning}` : ''}`
          : `Saved ${result.name}${still}; newer edits remain unsaved.${warning}`)
        return fullySaved
      } else {
        notify(`Saved ${result.name}`)
        return false
      }
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be saved.', 'error')
      return false
    } finally {
      savingRef.current = false
      setSaving(false)
    }
  }, [advanced.confirmPsdOverwrite, advanced.effectiveSaveFormat, advanced.encodePsd, advanced.mode, advanced.settle, advanced.withOutputCanvas, advanced.wouldLoseLayers, applyCrop, askChoice, documentMeta, exporting, isPristine, notify, printing, reflectDirty])

  const openPrintDialog = useCallback(async (invoker?: HTMLElement | null) => {
    const canvas = canvasRef.current
    if (!canvas || !image || printing || exporting || saving || workingRef.current) return
    const activeElement = document.activeElement instanceof HTMLElement ? document.activeElement : null
    printFocusReturnRef.current = invoker || activeElement || printButtonRef.current
    setPrinting(true)
    setSaveMenu(false)
    setExportMenu(false)
    try {
      if (!(await commitPendingRef.current())) return
      // advanced-integration (WP1): commit open Advanced sessions; Advanced prints its flattened output.
      if (advanced.mode === 'advanced') await advanced.settle()
      const documentId = documentIdRef.current
      const revision = contentRevisionRef.current
      const { width, height } = advanced.outputSize() ?? canvas
      const original = originalImageRef.current
      const blob = isPristine() && original?.printable
        ? original.blob
        : await advanced.withOutputCanvas((output) => canvasToBlob(output, 'png'))
      const data = new Uint8Array(await blob.arrayBuffer())
      if (documentIdRef.current !== documentId || contentRevisionRef.current !== revision) {
        notify('The image changed while preparing print. Open Print again to use the latest image.')
        return
      }
      setPrintSnapshot({
        data,
        url: URL.createObjectURL(blob),
        name: image.name,
        width,
        height,
      })
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The print preview could not be prepared.', 'error')
    } finally {
      setPrinting(false)
    }
  }, [advanced.mode, advanced.outputSize, advanced.settle, advanced.withOutputCanvas, exporting, image, isPristine, notify, printing, saving])

  const finishPrintDialog = useCallback(() => {
    setPrintSnapshot(null)
    printPendingFocusRef.current = printFocusReturnRef.current
    printFocusReturnRef.current = null
  }, [])

  const closePrintDialog = useCallback(() => {
    if (!printing) finishPrintDialog()
  }, [finishPrintDialog, printing])

  const runPrint = useCallback(async (settings: ImagePrintSettings) => {
    if (!printSnapshot || printing) return
    setPrinting(true)
    try {
      const printed = await window.simpleImage.printImage({
        data: printSnapshot.data,
        name: printSnapshot.name,
        width: printSnapshot.width,
        height: printSnapshot.height,
        settings,
      })
      if (printed) {
        finishPrintDialog()
        notify('The image was sent directly to your default printer.')
      } else {
        notify('The image could not be sent to the printer.', 'error')
      }
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be printed.', 'error')
    } finally {
      setPrinting(false)
    }
  }, [finishPrintDialog, notify, printSnapshot, printing])

  const exportImage = useCallback(async (format: ExportFormat) => {
    const canvas = canvasRef.current
    const current = imageRef.current
    if (!canvas || !current || exporting || savingRef.current || printing || workingRef.current) return
    setExporting(format)
    setSaveMenu(false)
    setExportMenu(false)
    try {
      if (!(await commitPendingRef.current({ keepStrokeAlive: true }))) return
      // advanced-integration (WP1): commit open Advanced sessions; Advanced exports its flattened output,
      // and PSD (offered only in Advanced) exports the layers.
      if (advanced.mode === 'advanced') await advanced.settle()
      if (format === 'pdf') {
        const blob = await advanced.withOutputCanvas((output) => canvasToBlob(output, 'png'))
        const result = await window.simpleImage.convertToPdf({
          data: new Uint8Array(await blob.arrayBuffer()),
          name: current.name,
        })
        if (result) notify(`Exported ${result.name}`)
        else notify('PDF export was canceled.')
        return
      }
      const data = format === 'psd'
        ? await advanced.encodePsd()
        : new Uint8Array(await (await advanced.withOutputCanvas((output) => canvasToBlob(output, format))).arrayBuffer())
      const result = await window.simpleImage.saveImage({
        data,
        path: null,
        name: current.name,
        format,
        forceDialog: true,
        purpose: 'export',
      })
      if (result) notify(`Exported ${result.name}`)
      else notify(`${format === 'jpeg' ? 'JPEG' : format.toUpperCase()} export was canceled.`)
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be exported.', 'error')
    } finally {
      setExporting(null)
    }
  }, [advanced.encodePsd, advanced.mode, advanced.settle, advanced.withOutputCanvas, exporting, notify, printing])

  const copyImageToClipboard = useCallback(async () => {
    const canvas = canvasRef.current
    if (!canvas || !image || saving || exporting || printing || workingRef.current) return false
    let copyCanvas: HTMLCanvasElement | null = null
    try {
      let copiedSelection = false
      if (tool === 'crop' && cropRect) {
        if (straighten) {
          const frame = cropFrame({ width: canvas.width, height: canvas.height }, straighten)
          copyCanvas = rotatedRegionCanvas(canvas, frame, toIntegerRect(cropRect, frame), straighten)
        } else {
          const rect = toIntegerRect(cropRect, { width: canvas.width, height: canvas.height })
          copyCanvas = document.createElement('canvas')
          copyCanvas.width = rect.width
          copyCanvas.height = rect.height
          const context = copyCanvas.getContext('2d')
          if (!context) throw new Error('The selected image area could not be prepared.')
          context.drawImage(canvas, rect.x, rect.y, rect.width, rect.height, 0, 0, rect.width, rect.height)
        }
        copiedSelection = true
      } else {
        await commitPendingRef.current()
      }
      // advanced-integration (WP1): the whole image comes from the flattened output in Advanced.
      const blob = copiedSelection && copyCanvas
        ? await canvasToBlob(copyCanvas, 'png')
        : await advanced.withOutputCanvas((output) => canvasToBlob(output, 'png'))
      await window.simpleImage.copyPng(new Uint8Array(await blob.arrayBuffer()))
      notify(copiedSelection ? 'Selected area copied to the clipboard.' : 'Image copied to the clipboard.')
      return true
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be copied.', 'error')
      return false
    } finally {
      if (copyCanvas) {
        copyCanvas.width = 1
        copyCanvas.height = 1
      }
    }
  }, [advanced.withOutputCanvas, cropRect, exporting, image, notify, printing, saving, straighten, tool])

  // advanced-integration (WP1): late-bound host actions for the Advanced editor (see the region above).
  hostActionsRef.current = {
    save: (forceDialog) => saveImage(forceDialog),
    print: () => { void openPrintDialog() },
    openExportMenu: () => {
      setSaveMenu(false)
      setExportMenu(true)
    },
  }

  // ---------------------------------------------------------------------------------------------
  // Tools
  // ---------------------------------------------------------------------------------------------

  const rotate = useCallback(async (clockwise: boolean) => {
    if (!imageRef.current || workingRef.current) return
    await commitPendingRef.current()
    applyGeometryNow(clockwise ? 'rotate-cw' : 'rotate-ccw', clockwise ? 'Rotate right' : 'Rotate left')
    if (tool === 'crop') {
      // The box belongs to the old orientation: start over on the turned image.
      setCropRect(null)
      setStraighten(0)
      setCropAdjusted(false)
      requestAnimationFrame(() => {
        const canvas = canvasRef.current
        if (canvas) setCropRect(defaultCropRect({ width: canvas.width, height: canvas.height }))
      })
    }
  }, [applyGeometryNow, setCropAdjusted, tool])

  const flip = useCallback(async (horizontal: boolean) => {
    if (!imageRef.current || workingRef.current) return
    await commitPendingRef.current()
    applyGeometryNow(horizontal ? 'flip-h' : 'flip-v', horizontal ? 'Flip horizontal' : 'Flip vertical')
    if (tool === 'crop' && cropRect) {
      // Mirror the box with the image; a mirrored straighten angle turns the other way.
      const frame = cropFrameSize()
      setCropRect(horizontal
        ? { ...cropRect, x: frame.width - cropRect.x - cropRect.width }
        : { ...cropRect, y: frame.height - cropRect.y - cropRect.height })
      if (straighten) setStraighten(-straighten)
    }
  }, [applyGeometryNow, cropFrameSize, cropRect, straighten, tool])

  const leaveTool = useCallback(async (next: Tool) => {
    if (strokeRef.current) endStroke()
    if (tool === 'markup' && next !== 'markup') await bakeMarkup()
    if (tool === 'crop' && next !== 'crop') cancelCrop()
    if (simplePanel === 'adjust' && next !== 'view') await commitAdjust(pendingAdjust)
  }, [bakeMarkup, cancelCrop, commitAdjust, endStroke, pendingAdjust, simplePanel, tool])

  const chooseTool = useCallback(async (next: Tool) => {
    if (!imageRef.current || workingRef.current) return
    if (next === tool && next !== 'crop') return
    await leaveTool(next)
    if (next === 'brush' || next === 'eraser') setPaintTool(next)
    if (next === 'crop') beginCropNow()
    else setTool(next)
  }, [beginCropNow, leaveTool, tool])

  const beginCrop = useCallback(() => { void chooseTool('crop') }, [chooseTool])

  const openAdjust = useCallback(async () => {
    if (!imageRef.current || workingRef.current) return
    if (simplePanel === 'adjust') {
      await commitAdjust(pendingAdjust)
      return
    }
    await leaveTool('view')
    setTool('view')
    setPendingAdjust(NEUTRAL_ADJUST)
    setAdjustTab('light')
    setComparing(false)
    setSimplePanel('adjust')
  }, [commitAdjust, leaveTool, pendingAdjust, simplePanel])

  const openResize = useCallback(async () => {
    if (!imageRef.current || workingRef.current) return
    await commitPendingRef.current()
    if (tool === 'crop') cancelCrop()
    setContextMenu(null)
    setResizeOpen(true)
  }, [cancelCrop, tool])

  const pickColorAt = useCallback((point: Point) => {
    const canvas = canvasRef.current
    if (!canvas) return false
    const color = sampleColor(canvas, point.x, point.y, 1)
    if (!color) return false
    if (color.a === 0) {
      notify('That pixel is transparent; pick a visible colour.')
      return false
    }
    setBrushColor(rgbHex(color))
    return true
  }, [notify])

  const pickFromScreen = useCallback(async () => {
    const Dropper = (window as Window & { EyeDropper?: new () => { open(): Promise<{ sRGBHex: string }> } }).EyeDropper
    if (!Dropper) return
    try {
      const result = await new Dropper().open()
      if (/^#[0-9a-f]{6}$/i.test(result.sRGBHex)) setBrushColor(result.sRGBHex.toLowerCase())
    } catch {
      // Escape cancels the screen picker.
    }
  }, [])

  const copyHex = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(brushColor.toUpperCase())
      notify(`Copied ${brushColor.toUpperCase()}`)
    } catch {
      notify(`Colour ${brushColor.toUpperCase()}`)
    }
  }, [brushColor, notify])

  // ---------------------------------------------------------------------------------------------
  // Paste (Ctrl+V): a new document when nothing is open, a movable picture in Markup otherwise
  // ---------------------------------------------------------------------------------------------

  const pasteFromClipboard = useCallback(async () => {
    if (printDialogOpenRef.current || workingRef.current || savingRef.current) return
    const reader = window.simpleImage.readClipboardImage
    if (!reader) {
      notify('Paste is not available in this version.', 'error')
      return
    }
    let png: Uint8Array | null = null
    try { png = await reader() } catch (error) {
      notify(messageOf(error, 'The clipboard image could not be read.'), 'error')
      return
    }
    if (!png || !png.byteLength) {
      notify('The clipboard has no image to paste.')
      return
    }
    if (!imageRef.current) {
      const request = ++loadRequestRef.current
      const payload: ImagePayload = { data: png, name: PASTED_NAME, path: null, size: png.byteLength, format: 'png', mime: 'image/png', directSave: false }
      try {
        await loadPayloadNow(payload, request)
        if (request !== loadRequestRef.current) return
        // A pasted image exists only in this window until it is saved.
        markEdited()
        notify('Pasted as a new image. Save to keep it.')
      } catch (error) {
        notify(messageOf(error, 'The pasted image could not be opened.'), 'error')
      }
      return
    }
    if (advanced.mode === 'advanced') return
    const url = URL.createObjectURL(new Blob([png as BlobPart], { type: 'image/png' }))
    try {
      const element = new Image()
      element.src = url
      await element.decode()
      await leaveTool('markup')
      pastedImagesRef.current.set(url, element)
      setTool('markup')
      markup.addItem(createImageItem(nextMarkupId(), url, element.naturalWidth, element.naturalHeight, dimensions))
      notify('Pasted picture: drag to place it, then choose Done.')
    } catch (error) {
      URL.revokeObjectURL(url)
      notify(messageOf(error, 'The pasted image could not be read.'), 'error')
    }
  }, [advanced.mode, dimensions, leaveTool, loadPayloadNow, markEdited, markup, notify])

  // ---------------------------------------------------------------------------------------------
  // Canvas pointer input: brush, eraser, eyedropper
  // ---------------------------------------------------------------------------------------------

  /** Pointer position in image pixels; strokes pass clampToCanvas=false so leaving the image does not paint a stripe along its edge (the canvas clips the paint). */
  const canvasPoint = useCallback((clientX: number, clientY: number, clampToCanvas = true): Point => {
    const canvas = canvasRef.current
    if (!canvas) return { x: 0, y: 0 }
    const bounds = canvas.getBoundingClientRect()
    return {
      x: clampToCanvas ? clamp((clientX - bounds.left) * canvas.width / bounds.width, 0, canvas.width) : (clientX - bounds.left) * canvas.width / bounds.width,
      y: clampToCanvas ? clamp((clientY - bounds.top) * canvas.height / bounds.height, 0, canvas.height) : (clientY - bounds.top) * canvas.height / bounds.height,
    }
  }, [])

  const pressureWidth = useCallback((event: PointerEvent, size: number) => {
    // Pen pressure scales the size (10%..100%); a mouse paints at full size.
    if (event.pointerType !== 'pen' || !(event.pressure > 0)) return size
    return Math.max(0.5, size * (0.1 + 0.9 * clamp(event.pressure, 0, 1)))
  }, [])

  const updateBrushRing = useCallback((clientX: number, clientY: number, visible: boolean) => {
    const ring = brushRingRef.current
    const stack = stackRef.current
    if (!ring || !stack) return
    if (!visible) {
      ring.style.display = 'none'
      return
    }
    const rect = stack.getBoundingClientRect()
    ring.style.display = 'block'
    ring.style.transform = `translate(${clientX - rect.left}px, ${clientY - rect.top}px)`
  }, [])

  const onCanvasPointerDown = useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (event.button !== 0 || !event.isPrimary || workingRef.current || isPanEvent(event.nativeEvent)) return
    const point = canvasPoint(event.clientX, event.clientY)
    if (tool === 'eyedropper' || ((tool === 'brush' || tool === 'eraser') && event.altKey)) {
      event.preventDefault()
      // A picked colour is for painting: the pipette hands over to the brush.
      if (pickColorAt(point) && tool === 'eyedropper') {
        setPaintTool('brush')
        setTool('brush')
      }
      return
    }
    if (tool !== 'brush' && tool !== 'eraser') return
    if (strokeRef.current) endStroke()
    event.currentTarget.setPointerCapture(event.pointerId)
    const current = imageRef.current
    const stroke: StrokeState = {
      pointerId: event.pointerId,
      erasing: tool === 'eraser',
      color: brushColor,
      size: brushSize,
      hadAlpha: current?.hasAlpha ?? false,
      last: point,
      recorder: null,
      before: currentState(),
      revision: -1,
      element: event.currentTarget,
    }
    strokeRef.current = stroke
    paintingRef.current = true
    const width = pressureWidth(event.nativeEvent, brushSize)
    const previous = lastStrokeEndRef.current
    // Shift+click continues with a straight segment from the previous stroke's end.
    if (event.shiftKey && previous && previous.documentId === documentIdRef.current) drawStrokeSegment(stroke, previous.point, point, width)
    else drawStrokeSegment(stroke, point, point, width)
    stroke.last = point
    lastPointRef.current = point
  }, [brushColor, brushSize, canvasPoint, currentState, drawStrokeSegment, endStroke, paintTool, pickColorAt, pressureWidth, tool])

  const onCanvasPointerMove = useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    readout.track(event.clientX, event.clientY)
    const painting = tool === 'brush' || tool === 'eraser'
    updateBrushRing(event.clientX, event.clientY, painting && brushSize * zoom >= 4 && !navigation.spaceHeld)
    const stroke = strokeRef.current
    if (!stroke || stroke.pointerId !== event.pointerId) return
    const native = event.nativeEvent
    const samples = typeof native.getCoalescedEvents === 'function' ? native.getCoalescedEvents() : []
    for (const sample of samples.length ? samples : [native]) {
      const point = canvasPoint(sample.clientX, sample.clientY, false)
      drawStrokeSegment(stroke, stroke.last, point, pressureWidth(sample, stroke.size))
      stroke.last = point
    }
    lastPointRef.current = stroke.last
  }, [brushSize, canvasPoint, drawStrokeSegment, navigation.spaceHeld, pressureWidth, readout, tool, updateBrushRing, zoom])

  const finishPainting = useCallback(() => {
    endStroke()
  }, [endStroke])

  // ---------------------------------------------------------------------------------------------
  // Crop box dragging (ratio lock, Shift/Alt, inside the turned image while straightening)
  // ---------------------------------------------------------------------------------------------

  const cropPoint = useCallback((event: React.PointerEvent, element: Element): Point => {
    const frame = cropFrameSize()
    const bounds = element.getBoundingClientRect()
    return {
      x: clamp((event.clientX - bounds.left) * frame.width / Math.max(1, bounds.width), 0, frame.width),
      y: clamp((event.clientY - bounds.top) * frame.height / Math.max(1, bounds.height), 0, frame.height),
    }
  }, [cropFrameSize])

  const startCropDrag = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0 || !event.isPrimary || isPanEvent(event.nativeEvent)) return
    if (!cropRect) return
    event.currentTarget.setPointerCapture(event.pointerId)
    const element = event.target as HTMLElement
    const handle = (element.dataset.handle as CropHandle | undefined) || (element.closest('.crop-box') ? 'move' : 'new')
    const start = cropPoint(event, event.currentTarget)
    cropDragRef.current = { handle, start, initial: cropRect, pointerId: event.pointerId }
  }, [cropPoint, cropRect])

  const moveCropDrag = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const drag = cropDragRef.current
    if (!drag || drag.pointerId !== event.pointerId) return
    const point = cropPoint(event, event.currentTarget)
    const frame = cropFrameSize()
    const ratio = activeRatio()
    const locked = ratio ?? (event.shiftKey ? (drag.handle === 'new' ? 1 : drag.initial.width / Math.max(1e-9, drag.initial.height)) : null)
    const options = { ratio: locked, fromCenter: event.altKey, bounds: frame }
    const inside = straighten ? (rect: CropRect) => insideRotated(rect, dimensions, straighten) : null
    let next: CropRect | null
    if (drag.handle === 'new') {
      if (Math.hypot(point.x - drag.start.x, point.y - drag.start.y) < 2) return
      next = inside ? rectFromPointsInside(drag.start, point, options, inside) : rectFromPoints(drag.start, point, options)
    } else {
      const dx = point.x - drag.start.x
      const dy = point.y - drag.start.y
      next = inside ? constrainDragInside(drag.initial, drag.handle, dx, dy, options, inside) : constrainDrag(drag.initial, drag.handle, dx, dy, options)
    }
    if (!next) return
    setCropRect(next)
    if (!cropAdjustedRef.current) setCropAdjusted(true)
  }, [activeRatio, cropFrameSize, cropPoint, dimensions, setCropAdjusted, straighten])

  const finishCropDrag = useCallback(() => { cropDragRef.current = null }, [])

  const handleDrop = useCallback((event: React.DragEvent) => {
    event.preventDefault()
    setDraggingFile(false)
    if (printDialogOpenRef.current) {
      event.stopPropagation()
      return
    }
    const file = event.dataTransfer.files[0]
    if (file) void openDroppedFile(file)
  }, [openDroppedFile])

  const confirmDiscard = useCallback(async () => {
    const action = confirmState?.action
    setConfirmState(null)
    if (action) await action()
  }, [confirmState])

  const confirmSave = useCallback(async () => {
    const state = confirmState
    if (!state) return
    // IMAGE-SIE-2: the prompt's Save applies the pending crop first.
    if (state.cropPending && cropAdjustedRef.current && !(await applyCrop())) return
    if (await saveImage(false)) {
      setConfirmState(null)
      await state.action()
    }
  }, [applyCrop, confirmState, saveImage])

  const toggleInspector = useCallback(() => setInspectorOpen((value) => !value), [])

  // Undo / Redo: the markup session first while Markup is open, then the image history.
  const undo = useCallback(() => {
    if (tool === 'markup' && markup.undo()) return
    undoCanvas()
  }, [markup, tool, undoCanvas])

  const redo = useCallback(() => {
    if (tool === 'markup' && markup.redo()) return
    redoCanvas()
  }, [markup, redoCanvas, tool])

  useEffect(() => window.simpleImage.onMaximized(setIsMaximized), [])
  useEffect(() => window.simpleImage.onOpenExternal(openPath), [openPath])
  // Answers for the shared window guard (Alt+F4, taskbar, quit, Windows sign-out, crash/hang).
  // Pending strokes, markup, adjustments and Advanced sessions are committed first so the
  // answer covers every edit; an unapplied crop counts as unsaved work.
  const ioStateRef = useRef({ name: '', untitled: true, saving: false, saveImage })
  ioStateRef.current = { name: image?.name ?? '', untitled: !image?.path, saving, saveImage }
  useEffect(() => {
    const io = getSimpleIO()
    if (!io) return undefined
    const settle = async () => {
      await commitPendingRef.current().catch(() => false)
      if (advanced.mode === 'advanced') await advanced.settle().catch(() => undefined)
    }
    const offs = [
      io.onRequest('close-query', async () => {
        if (closeDecidedRef.current) return { dirty: false }
        await settle()
        const cropPending = cropAdjustedRef.current && Boolean(imageRef.current)
        const state = ioStateRef.current
        return { dirty: dirtyRef.current || cropPending || Boolean(pendingSessionRef.current), saving: state.saving, title: state.name || undefined, kind: 'image', untitled: state.untitled }
      }),
      io.onRequest('save-now', async () => {
        await settle()
        return ioStateRef.current.saveImage(false)
      }),
      io.onRequest('discard', () => { closeDecidedRef.current = true; return true }),
    ]
    return () => offs.forEach((off) => off())
  }, [advanced.mode, advanced.settle])

  useEffect(() => window.simpleImage.onCloseRequested(() => {
    requestProceed('Save changes before closing?', 'Your edits will be lost if you close without saving.', () => { closeDecidedRef.current = true; window.simpleImage.confirmClose() })
  }), [requestProceed])

  // IMAGE-SIE-2: unsaved work shows in the window title (and so in the taskbar) as "• name".
  const titleName = image?.name ?? null
  const titleDirty = dirty || cropAdjusted || pendingSession
  useEffect(() => {
    if (!titleName) return
    window.simpleImage.setTitle(`${titleDirty ? '• ' : ''}${titleName}`)
  }, [titleDirty, titleName])

  useEffect(() => {
    try { window.localStorage.setItem(INSPECTOR_STORAGE_KEY, storedInspectorOpen(inspectorOpen)) } catch { /* Preference persistence is optional. */ }
  }, [inspectorOpen])

  useEffect(() => () => {
    if (printSnapshot) URL.revokeObjectURL(printSnapshot.url)
  }, [printSnapshot])

  useEffect(() => {
    const backgroundSurfaces = document.querySelectorAll<HTMLElement>('[data-print-background]')
    backgroundSurfaces.forEach((surface) => { surface.inert = Boolean(printSnapshot) })
    return () => backgroundSurfaces.forEach((surface) => { surface.inert = false })
  }, [printSnapshot])

  useEffect(() => {
    if (printSnapshot || printing || !printPendingFocusRef.current) return
    const preferred = printPendingFocusRef.current
    printPendingFocusRef.current = null
    window.requestAnimationFrame(() => {
      const fallback = printButtonRef.current
      const target = preferred.isConnected && !preferred.hasAttribute('disabled') ? preferred : fallback
      if (target?.isConnected && !target.hasAttribute('disabled')) target.focus()
    })
  }, [printSnapshot, printing])

  useEffect(() => {
    const closeMenus = (event: PointerEvent) => {
      const target = event.target as Node
      if (!saveGroupRef.current?.contains(target)) setSaveMenu(false)
      if (!exportGroupRef.current?.contains(target)) setExportMenu(false)
    }
    document.addEventListener('pointerdown', closeMenus)
    return () => document.removeEventListener('pointerdown', closeMenus)
  }, [])

  useEffect(() => {
    if (!contextMenu) return
    const closeOutside = (event: PointerEvent) => {
      if (!(event.target instanceof Element) || !event.target.closest('.image-context-menu')) setContextMenu(null)
    }
    const close = () => setContextMenu(null)
    const viewport = viewportRef.current
    document.addEventListener('pointerdown', closeOutside)
    window.addEventListener('blur', close)
    window.addEventListener('resize', close)
    viewport?.addEventListener('scroll', close, { passive: true })
    window.requestAnimationFrame(() => contextMenuRef.current?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus())
    return () => {
      document.removeEventListener('pointerdown', closeOutside)
      window.removeEventListener('blur', close)
      window.removeEventListener('resize', close)
      viewport?.removeEventListener('scroll', close)
    }
  }, [contextMenu])

  useEffect(() => {
    if (!fitMode || !image) return
    const observer = new ResizeObserver(() => { if (fitModeRef.current) fitImage() })
    if (viewportRef.current) observer.observe(viewportRef.current)
    return () => observer.disconnect()
  }, [fitImage, fitMode, image])

  // Straightening changes the document box (the turned frame): keep it fitted.
  useEffect(() => {
    if (fitModeRef.current && image) fitImage()
  }, [displaySize.height, displaySize.width])

  // The Adjust preview (proxy at the displayed device size, overlay over the image).
  const adjustActive = simplePanel === 'adjust' && Boolean(image) && advanced.mode !== 'advanced'
  const adjustPreview = useAdjustPreview({
    active: adjustActive,
    canvasRef,
    overlayRef: adjustOverlayRef,
    pending: pendingAdjust,
    displayScale: zoom * navigation.dpr,
    sourceKey: `${documentIdRef.current}:${historyVersion}:${dimensions.width}x${dimensions.height}`,
  })
  const previewVisible = adjustActive && adjustPreview.showing && !comparing

  const runAuto = useCallback(() => {
    const proxy = adjustPreview.proxy
    if (!proxy) return
    try {
      const result = autoEnhance(proxy)
      const quick: QuickAdjust = result.quick
      setPendingAdjust((current) => ({ ...current, quick }))
      if (result.strength < 0.02) notify('This photo already looks balanced; Auto changed very little.')
    } catch (error) {
      notify(messageOf(error, 'Auto could not analyse this image.'), 'error')
    }
  }, [adjustPreview.proxy, notify])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (printSnapshot) {
        event.preventDefault()
        if (event.key === 'Escape' && !printing) closePrintDialog()
        if ((event.ctrlKey || event.metaKey) && shortcutKey(event) === 'p' && !printing) {
          const printButton = document.querySelector<HTMLButtonElement>('.image-print-dialog .print-submit')
          printButton?.click()
        }
        return
      }
      if (choiceState) {
        if (event.key === 'Escape') {
          event.preventDefault()
          choiceState.resolve('cancel')
        } else if (event.key !== 'Tab' && event.key !== 'Enter' && event.key !== ' ') event.preventDefault()
        return
      }
      // The resize dialog handles its own keys (Escape, Enter submits).
      if (resizeOpen) return
      if (confirmState) {
        if (event.key === 'Escape' && !saving) setConfirmState(null)
        if (event.key === 'Tab' || event.key === 'Enter' || event.key === ' ') return
        event.preventDefault()
        return
      }
      if (saving || exporting || working) {
        if (event.key === 'Tab') return
        event.preventDefault()
        return
      }
      if (contextMenu) {
        const menu = contextMenuRef.current
        const buttons = [...menu?.querySelectorAll<HTMLButtonElement>('button:not(:disabled)') ?? []]
        const current = Math.max(0, buttons.indexOf(document.activeElement as HTMLButtonElement))
        const key = shortcutKey(event)
        if ((event.ctrlKey || event.metaKey) && key === 'c') {
          event.preventDefault()
          setContextMenu(null)
          void copyImageToClipboard()
        } else if ((event.ctrlKey || event.metaKey) && key === 'z' && !event.shiftKey) {
          event.preventDefault()
          setContextMenu(null)
          undo()
        } else if ((event.ctrlKey || event.metaKey) && (key === 'y' || (key === 'z' && event.shiftKey))) {
          event.preventDefault()
          setContextMenu(null)
          redo()
        } else if (key === 'escape') {
          event.preventDefault()
          setContextMenu(null)
        } else if (key === 'arrowdown' || key === 'arrowup' || key === 'home' || key === 'end' || key === 'tab') {
          event.preventDefault()
          const next = key === 'home'
            ? 0
            : key === 'end'
              ? buttons.length - 1
              : (current + (key === 'arrowup' || (key === 'tab' && event.shiftKey) ? -1 : 1) + buttons.length) % buttons.length
          buttons[next]?.focus()
        } else if (!(event.target instanceof Element) || !event.target.closest('.image-context-menu') || (event.ctrlKey || event.metaKey)) {
          event.preventDefault()
        }
        return
      }
      // advanced-integration (WP1): while Advanced (or one of its dialogs) is open, the editor owns every
      // key except the host file shortcuts Ctrl+O, Ctrl+S, Ctrl+Shift+S and Ctrl+P, which continue below.
      if (advanced.mode === 'advanced' || advanced.modalOpen) {
        if (event.key === 'Escape' && (saveMenu || exportMenu)) {
          setSaveMenu(false)
          setExportMenu(false)
          return
        }
        if (advanced.handleHostKey(event)) return
      }
      const control = event.ctrlKey || event.metaKey
      const key = shortcutKey(event)
      const target = event.target instanceof HTMLElement ? event.target : null
      const textInput = Boolean(target?.closest('input:not([type="range"]):not([type="color"]):not([type="checkbox"]), textarea, select, [contenteditable="true"]'))
      const nativeTextSelected = Boolean(window.getSelection()?.toString().trim())
      const typing = textInput || event.isComposing
      const onFormControl = Boolean(target?.closest('input, select, textarea'))
      // IMAGE-SIE-1: never act in the middle of a stroke. Undo/Redo wait for the stroke to end; other
      // shortcuts (Save, Export, Copy) see the stroke so far as one finished step.
      if (strokeRef.current && control) {
        if (key === 'z' || key === 'y') {
          event.preventDefault()
          return
        }
        commitStrokeSegment()
      }
      // Ctrl+A outside a text field must not select the panels' text (Ctrl+C would then copy UI text instead of the image).
      if (control && key === 'a' && !textInput) event.preventDefault()
      if (control && key === 'c' && image && !textInput && !nativeTextSelected) { event.preventDefault(); void copyImageToClipboard() }
      if (control && key === 'v' && !textInput && !event.shiftKey && advanced.mode !== 'advanced') { event.preventDefault(); void pasteFromClipboard() }
      if (!control && !event.altKey && !event.shiftKey && tool === 'view' && !onFormControl && image && album.paths.length > 1
        && (event.key === 'ArrowLeft' || event.key === 'ArrowRight')) {
        const viewport = viewportRef.current
        const scrollable = viewport ? viewport.scrollWidth > viewport.clientWidth + 1 : false
        if (!scrollable) { event.preventDefault(); stepAlbum(event.key === 'ArrowRight' ? 1 : -1) }
      }
      if (!control && !event.altKey && tool === 'crop' && cropRect && !onFormControl && event.key.startsWith('Arrow')) {
        event.preventDefault()
        const step = event.shiftKey ? 10 : 1
        nudgeCrop(event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0, event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0)
      }
      if (control && key === 'o') { event.preventDefault(); void openDialog() }
      if (control && key === 's') { event.preventDefault(); void saveImage(event.shiftKey) }
      if (control && event.shiftKey && key === 'e' && image) {
        event.preventDefault()
        setSaveMenu(false)
        setExportMenu((value) => !value)
      }
      if (control && key === 'p') { event.preventDefault(); void openPrintDialog() }
      if (control && !event.shiftKey && key === 'z' && !textInput) { event.preventDefault(); undo() }
      if (control && (key === 'y' || (event.shiftKey && key === 'z')) && !textInput) { event.preventDefault(); redo() }
      if (control && (event.key === '+' || event.key === '=') && image) { event.preventDefault(); navigation.step(1) }
      if (control && (event.key === '-' || event.key === '_') && image) { event.preventDefault(); navigation.step(-1) }
      if (control && event.key === '0' && image) { event.preventDefault(); fitImage() }
      if (control && event.key === '1' && image) { event.preventDefault(); navigation.actualPixels() }
      // Alt+F4 must keep closing the window; only plain F4 toggles the details panel.
      if (event.key === 'F4' && !event.altKey && image && !event.repeat) { event.preventDefault(); toggleInspector() }
      if (!control && !event.altKey && image && !typing) {
        if (navigation.handleSpaceDown(event)) return
        if ((event.code === 'BracketLeft' || event.code === 'BracketRight') && !event.shiftKey) {
          event.preventDefault()
          setBrushSize((size) => stepBrushSize(size, event.code === 'BracketRight' ? 1 : -1))
        }
        if (key === 'i' && !event.shiftKey && !event.repeat && advanced.mode !== 'advanced') {
          event.preventDefault()
          void chooseTool(tool === 'eyedropper' ? paintTool : 'eyedropper')
        }
        if (event.code === 'Backslash' && simplePanel === 'adjust' && !event.repeat) {
          event.preventDefault()
          setComparing((value) => !value)
        }
        if ((event.key === 'Delete' || event.key === 'Backspace') && tool === 'markup') {
          if (markup.deleteSelected()) event.preventDefault()
        }
      }
      if (event.key === 'Escape') {
        setSaveMenu(false)
        setExportMenu(false)
        if (!typing) {
          if (leveling) setLeveling(false)
          else if (tool === 'crop') cancelCrop()
          else if (tool === 'markup') markup.select(null)
          else if (tool === 'eyedropper') setTool(paintTool)
        }
      }
      // Defect 3: Enter applies the crop ONCE. preventDefault stops the focused toolbar button (e.g. Crop)
      // from also activating; buttons in the crop strip and dialogs keep Enter for themselves.
      if (event.key === 'Enter' && !control && tool === 'crop' && cropRect && !typing) {
        const button = target?.closest('button')
        if (!(button && button.closest('.context-strip, .modal, .image-context-menu'))) {
          event.preventDefault()
          if (!event.repeat) void applyCrop()
        }
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [advanced.handleHostKey, advanced.modalOpen, advanced.mode, album, applyCrop, cancelCrop, choiceState, chooseTool, closePrintDialog, commitStrokeSegment, confirmState, contextMenu, copyImageToClipboard, cropRect, exportMenu, exporting, fitImage, image, leveling, markup, navigation, nudgeCrop, openDialog, openPrintDialog, paintTool, pasteFromClipboard, printSnapshot, printing, redo, resizeOpen, saveImage, saveMenu, saving, simplePanel, stepAlbum, toggleInspector, tool, undo, working])

  const cropStyle = useMemo(() => cropRect && displaySize.width && displaySize.height ? {
    left: `${cropRect.x / displaySize.width * 100}%`,
    top: `${cropRect.y / displaySize.height * 100}%`,
    width: `${cropRect.width / displaySize.width * 100}%`,
    height: `${cropRect.height / displaySize.height * 100}%`,
  } : undefined, [cropRect, displaySize.height, displaySize.width])

  const megapixels = dimensions.width && dimensions.height ? dimensions.width * dimensions.height / 1_000_000 : 0
  const markupOpen = tool === 'markup' && Boolean(image)
  const canUndo = undoRef.current.length > 0 || (markupOpen && markup.canUndo)
  const canRedo = redoRef.current.length > 0 || (markupOpen && markup.canRedo)
  void historyVersion
  const busy = Boolean(working)
  const toolsDisabled = !image || busy
  const straightening = tool === 'crop' && straighten !== 0
  const showRing = (tool === 'brush' || tool === 'eraser') && brushSize * zoom >= 4
  const stateLabel = saving ? 'Saving…' : working ? `${working.label}…` : cropAdjusted ? 'Crop not applied' : dirty || pendingSession ? 'Modified' : 'Saved'
  const zoomLabel = `${navigation.devicePercent >= 10 ? Math.round(navigation.devicePercent) : navigation.devicePercent.toFixed(1)}%`
  const eyeDropperSupported = typeof window !== 'undefined' && 'EyeDropper' in window
  const canvasStyle: React.CSSProperties = straightening
    ? {
        position: 'absolute',
        left: (displaySize.width - dimensions.width) / 2 * zoom,
        top: (displaySize.height - dimensions.height) / 2 * zoom,
        width: dimensions.width * zoom,
        height: dimensions.height * zoom,
        transform: `rotate(${straighten}deg)`,
      }
    : { width: dimensions.width * zoom, height: dimensions.height * zoom }
  if (previewVisible) canvasStyle.opacity = 0

  const blockBackgroundAction = (event: React.SyntheticEvent) => {
    if (!printSnapshot) return
    const target = event.target
    if (target instanceof Element && target.closest('.image-print-overlay')) return
    event.preventDefault()
    event.stopPropagation()
  }

  return (
    <main
      className="app-shell"
      onClickCapture={blockBackgroundAction}
      onPointerDownCapture={blockBackgroundAction}
      onInputCapture={blockBackgroundAction}
      onChangeCapture={blockBackgroundAction}
      onSubmitCapture={blockBackgroundAction}
      onDragEnter={(event) => {
        event.preventDefault()
        if (printSnapshot) {
          event.dataTransfer.dropEffect = 'none'
          setDraggingFile(false)
        } else setDraggingFile(true)
      }}
      onDragOver={(event) => { event.preventDefault(); if (printSnapshot) event.dataTransfer.dropEffect = 'none' }}
      onDrop={handleDrop}
    >
      <header className="titlebar" data-print-background aria-hidden={printSnapshot ? true : undefined}>
        <div className="title-identity drag-region">
          <img className="brand-mark" src="./brand-icon.png" alt="" />
          <span className="app-name">simple</span>
          <span className="title-divider" />
          <span className="document-title">{image?.name || 'Image'}</span>
          {image && <span className={`save-state ${dirty || cropAdjusted || pendingSession ? 'is-dirty' : ''}`}>{stateLabel}</span>}
        </div>
        <div className="title-actions">
          <button title="Open image (Ctrl+O)" onClick={() => void openDialog()}><FolderOpen /><span>Open</span></button>
          <div ref={saveGroupRef} className="save-group">
            <button className="primary-button" disabled={!image || saving || Boolean(exporting) || printing || busy} title="Save (Ctrl+S)" onClick={() => void saveImage(false)}><Save /><span>{saving ? 'Saving…' : 'Save'}</span></button>
            <button
              className="primary-button split-button"
              disabled={!image || saving || Boolean(exporting) || printing || busy}
              title="Save options"
              aria-haspopup="menu"
              aria-expanded={saveMenu}
              onClick={() => { setExportMenu(false); setSaveMenu((value) => !value) }}
            ><ChevronDown /></button>
            {saveMenu && (
              <div className="save-menu" role="menu" aria-label="Save options">
                <button role="menuitem" onClick={() => void saveImage(true, advanced.flatSaveFormat())}><Save /><span><strong>Save as…</strong><small>{advanced.mode === 'advanced' ? 'One flattened image in the current format' : 'Move or rename the working image'}</small></span></button>
                {advanced.psdAvailable && (
                  <button role="menuitem" onClick={() => void saveImage(true, 'psd')}><Layers /><span><strong>Save as Photoshop (.psd)…</strong><small>Keeps layers, masks and blend modes</small></span></button>
                )}
              </div>
            )}
          </div>
          <div ref={exportGroupRef} className="export-group">
            <button
              className="export-button"
              disabled={!image || Boolean(exporting) || printing || saving || busy}
              title={advanced.mode === 'advanced' ? 'Export As (Alt+Shift+Ctrl+W)' : 'Export As (Ctrl+Shift+E)'}
              aria-haspopup="menu"
              aria-expanded={exportMenu}
              onClick={() => { setSaveMenu(false); setExportMenu((value) => !value) }}
            >
              <FileDown />
              <span>{exporting ? `Exporting ${exporting === 'jpeg' ? 'JPEG' : exporting.toUpperCase()}…` : 'Export As'}</span>
              <ChevronDown className="button-chevron" />
            </button>
            {exportMenu && image && (
              <div className="export-menu" role="menu" aria-label="Export image as">
                <div className="menu-heading"><strong>Export a copy</strong><small>The open image and unsaved edits stay in place.</small></div>
                {[...EXPORT_OPTIONS, ...advanced.extraExportOptions].map((option) => (
                  <button
                    key={option.format}
                    role="menuitem"
                    data-export-format={option.format}
                    onClick={() => void exportImage(option.format)}
                  >
                    <FileDown />
                    <span><strong>{option.label}</strong><small>{option.detail}</small></span>
                  </button>
                ))}
                <p className="export-note">SVG input is rasterized at its displayed pixel dimensions. Vector paths are not preserved, so SVG export is not offered.</p>
              </div>
            )}
          </div>
          <button ref={printButtonRef} disabled={!image || printing || Boolean(exporting) || saving || busy} title="Print (Ctrl+P)" onClick={() => void openPrintDialog(printButtonRef.current)}><Printer /><span>{printing ? 'Preparing…' : 'Print'}</span></button>
        </div>
        <div className="window-actions">
          <button aria-label="Minimize" onClick={() => window.simpleImage.minimize()}><Minus /></button>
          <button aria-label={isMaximized ? 'Restore' : 'Maximize'} onClick={() => window.simpleImage.toggleMaximize()}><Maximize2 /></button>
          <button className="close-button" aria-label="Close" onClick={() => requestProceed('Save changes before closing?', 'Your edits will be lost if you close without saving.', () => { closeDecidedRef.current = true; window.simpleImage.confirmClose() })}><X /></button>
        </div>
      </header>

      <section className="toolbar" data-print-background aria-label="Image tools" aria-hidden={printSnapshot ? true : undefined} hidden={advanced.mode === 'advanced'}>
        <div className="toolbar-row">
          <div className="tool-group">
            <button className={tool === 'view' ? 'active' : ''} disabled={toolsDisabled} onClick={() => void chooseTool('view')} title="View"><MousePointer2 /><span>View</span></button>
            <button className={tool === 'crop' ? 'active' : ''} disabled={toolsDisabled} onClick={beginCrop} title="Crop"><Crop /><span>Crop</span></button>
            <button disabled={toolsDisabled} onClick={() => void rotate(false)} title="Rotate left"><RotateCcw /></button>
            <button disabled={toolsDisabled} onClick={() => void rotate(true)} title="Rotate right"><RotateCw /></button>
          </div>
          <div className="toolbar-rule" />
          <div className="tool-group">
            <button className={simplePanel === 'adjust' ? 'active' : ''} disabled={toolsDisabled} onClick={() => void openAdjust()} title="Adjust light and colour, or apply a look"><SlidersHorizontal /><span>Adjust</span></button>
            <button className={tool === 'markup' ? 'active' : ''} disabled={toolsDisabled} onClick={() => void chooseTool('markup')} title="Markup: text, arrows and shapes"><PenLine /><span>Markup</span></button>
          </div>
          <div className="toolbar-rule" />
          <div className="tool-group">
            <button className={tool === 'brush' ? 'active' : ''} disabled={toolsDisabled} onClick={() => void chooseTool('brush')} title="Brush"><Brush /><span>Brush</span></button>
            <button className={tool === 'eraser' ? 'active' : ''} disabled={toolsDisabled} onClick={() => void chooseTool('eraser')} title="Eraser"><Eraser /><span>Eraser</span></button>
            <label className="color-control" title="Brush color"><input type="color" value={brushColor} onChange={(event) => setBrushColor(event.target.value)} /><span style={{ background: brushColor }} /></label>
            <button className={`icon-tool ${tool === 'eyedropper' ? 'active' : ''}`} disabled={toolsDisabled} onClick={() => void chooseTool(tool === 'eyedropper' ? paintTool : 'eyedropper')} title="Pick a colour from the image (I; Alt+click while painting)" aria-label="Pick a colour from the image"><Pipette /></button>
            <label className="size-control"><span>Size</span><input type="range" min="1" max={MAX_BRUSH_SIZE} value={brushSize} onChange={(event) => setBrushSize(Number(event.target.value))} /><output>{brushSize}px</output></label>
          </div>
          <div className="toolbar-spacer" />
          <div className="tool-group">
            <button disabled={!canUndo || busy} onClick={undo} title="Undo (Ctrl+Z)"><Undo2 /></button>
            <button disabled={!canRedo || busy} onClick={redo} title="Redo (Ctrl+Y)"><Redo2 /></button>
            <button
              className={image && inspectorOpen ? 'active' : ''}
              disabled={!image}
              aria-controls="image-inspector"
              aria-expanded={Boolean(image && inspectorOpen)}
              onClick={toggleInspector}
              title={`${inspectorOpen ? 'Hide' : 'Show'} image details (F4)`}
            ><FileImage /><span>Details</span></button>
            {/* advanced-integration (WP1): the one entry point into the Photoshop-style editor */}
            <button
              className={`advanced-button ${advanced.busy ? 'is-busy' : ''}`}
              title="Advanced editor: layers, selections, adjustment layers and Photoshop shortcuts"
              disabled={!image || saving || Boolean(exporting) || printing || advanced.busy || busy}
              onClick={() => void advanced.enter()}
            ><Layers /><span>{advanced.busy ? 'Opening…' : 'Advanced'}</span></button>
          </div>
        </div>
        {tool === 'crop' && cropRect && image && (
          <CropOptions
            aspect={cropAspect}
            canSwap={presetCanSwap(cropAspect)}
            straighten={straighten}
            leveling={leveling}
            size={{ width: cropRect.width, height: cropRect.height }}
            busy={busy}
            onAspect={chooseAspect}
            onSwap={swapAspect}
            onStraighten={changeStraighten}
            onLevel={() => setLeveling((value) => !value)}
            onFlip={(horizontal) => void flip(horizontal)}
            onResize={() => void openResize()}
            onCancel={cancelCrop}
            onApply={() => void applyCrop()}
          />
        )}
        {markupOpen && <MarkupOptions session={markup} busy={busy} onDone={() => void chooseTool('view')} />}
      </section>

      <section className="content" data-print-background aria-hidden={printSnapshot ? true : undefined} hidden={advanced.mode === 'advanced'}>
        <div
          ref={viewportRef}
          className={`viewport ${image ? '' : 'is-empty'} ${navigation.spaceHeld ? 'is-pan-ready' : ''} ${navigation.panning ? 'is-panning' : ''} ${tool === 'view' && image ? 'is-view-tool' : ''}`}
        >
          {image && album.paths.length > 1 && (
            <>
              <button className="album-nav album-prev" title="Previous image (←)" aria-label="Previous image" onClick={() => stepAlbum(-1)}><ChevronLeft /></button>
              <button className="album-nav album-next" title="Next image (→)" aria-label="Next image" onClick={() => stepAlbum(1)}><ChevronRight /></button>
              <div className="album-counter">{album.index + 1} / {album.paths.length}</div>
            </>
          )}
          {!image && (
            <div className="welcome">
              <img src="./brand-icon.png" alt="" />
              <h1>See every detail.</h1>
              <p>Open, inspect, crop, rotate, adjust, mark up, print, or export a copy as PNG, JPEG, WebP, or PDF without uploading anything.</p>
              <div className="welcome-actions">
                <button className="welcome-primary" onClick={() => void openDialog()}><ImagePlus /> Open an image</button>
                <button onClick={() => void window.simpleImage.newWindow()}><FileImage /> New window</button>
              </div>
              <div className="format-list">PNG · JPEG · WebP · GIF · BMP · SVG · AVIF</div>
              <small>Or drop an image anywhere in this window, or paste one with Ctrl+V</small>
            </div>
          )}
          <div
            className="canvas-stage"
            hidden={!image}
            style={{ minWidth: displaySize.width * zoom + 80, minHeight: displaySize.height * zoom + 80 }}
            onContextMenu={(event) => {
              event.preventDefault()
              if (!image || saving || exporting || printing || printSnapshot || busy) return
              setSaveMenu(false)
              setExportMenu(false)
              setContextMenu({
                x: Math.max(8, Math.min(event.clientX, window.innerWidth - 224)),
                y: Math.max(8, Math.min(event.clientY, window.innerHeight - 220)),
              })
            }}
          >
            <div
              ref={stackRef}
              className={`canvas-stack ${straightening ? 'is-straightening' : ''}`}
              style={{ width: displaySize.width * zoom, height: displaySize.height * zoom }}
            >
              <canvas
                ref={canvasRef}
                className={`image-canvas tool-${tool} ${navigation.pixelated ? 'is-pixelated' : ''} ${showRing ? 'has-ring' : ''}`}
                style={canvasStyle}
                onPointerDown={onCanvasPointerDown}
                onPointerMove={onCanvasPointerMove}
                onPointerUp={finishPainting}
                onPointerCancel={finishPainting}
                onPointerLeave={() => { updateBrushRing(0, 0, false); readout.clear() }}
              />
              {adjustActive && <canvas ref={adjustOverlayRef} className={`adjust-preview ${navigation.pixelated ? 'is-pixelated' : ''}`} hidden={!previewVisible} aria-hidden="true" />}
              {showRing && <div ref={brushRingRef} className="brush-ring" style={{ width: brushSize * zoom, height: brushSize * zoom, marginLeft: -brushSize * zoom / 2, marginTop: -brushSize * zoom / 2 }} aria-hidden="true" />}
              {markupOpen && <MarkupLayer width={dimensions.width} height={dimensions.height} zoom={zoom} session={markup} disabled={busy} />}
              {tool === 'crop' && cropRect && leveling && (
                <LevelLineLayer
                  frame={displaySize}
                  onLine={(from, to) => { setLeveling(false); changeStraighten(levelAngle(from, to, straighten)) }}
                  onCancel={() => setLeveling(false)}
                />
              )}
              {tool === 'crop' && cropRect && !leveling && (
                <div className="crop-layer" onPointerDown={startCropDrag} onPointerMove={moveCropDrag} onPointerUp={finishCropDrag} onPointerCancel={finishCropDrag}>
                  <div className="crop-box" style={cropStyle}>
                    <span className="crop-grid vertical one" /><span className="crop-grid vertical two" />
                    <span className="crop-grid horizontal one" /><span className="crop-grid horizontal two" />
                    {['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map((handle) => <i key={handle} data-handle={handle} className={`crop-handle ${handle}`} />)}
                    <b>{Math.round(cropRect.width)} × {Math.round(cropRect.height)}</b>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        {image && (
          <aside id="image-inspector" className={`inspector ${simplePanel === 'adjust' ? 'is-adjusting' : ''}`} aria-label={simplePanel === 'adjust' ? 'Adjust' : 'Image details'} hidden={!inspectorOpen && simplePanel !== 'adjust'}>
            {simplePanel === 'adjust' ? (
              <AdjustPanel
                pending={pendingAdjust}
                tab={adjustTab}
                thumbnails={adjustPreview.thumbnails}
                busy={busy}
                progress={working?.progress ?? null}
                comparing={comparing}
                autoAvailable={Boolean(adjustPreview.proxy)}
                onTab={setAdjustTab}
                onQuick={(quick) => setPendingAdjust((current) => ({ ...current, quick }))}
                onLook={(look: LookId) => setPendingAdjust((current) => ({ ...current, look, intensity: look === current.look ? current.intensity : 100 }))}
                onIntensity={(intensity) => setPendingAdjust((current) => ({ ...current, intensity }))}
                onAuto={runAuto}
                onReset={() => setPendingAdjust(NEUTRAL_ADJUST)}
                onCompare={setComparing}
                onCancel={closeAdjust}
                onDone={() => void commitAdjust(pendingAdjust)}
              />
            ) : (
              <>
                <div className="inspector-heading">
                  <FileImage />
                  <div><strong>Image details</strong><span>Read from this asset</span></div>
                  <button className="inspector-close" type="button" aria-label="Close image details" title="Close image details (F4)" onClick={() => setInspectorOpen(false)}><X /></button>
                </div>
                <dl>
                  <div><dt>Format</dt><dd>{image.sourceFormat.toUpperCase()}</dd></div>
                  <div><dt>Dimensions</dt><dd className="dimension-value">{dimensions.width.toLocaleString()} × {dimensions.height.toLocaleString()} px<button type="button" className="inline-action" disabled={busy} onClick={() => void openResize()} title="Resize the image">Resize…</button></dd></div>
                  <div><dt>Resolution</dt><dd>{megapixels < 0.1 ? megapixels.toFixed(3) : megapixels.toFixed(1)} MP</dd></div>
                  <div><dt>Aspect ratio</dt><dd>{aspectRatio(dimensions.width, dimensions.height)}</dd></div>
                  <div><dt>Color</dt><dd>{image.hasAlpha ? 'RGBA · transparency' : 'RGB · opaque'}</dd></div>
                  <div><dt>File size</dt><dd>{formatBytes(image.size)}</dd></div>
                </dl>
                <div className="privacy-note"><span />Processed locally on this computer</div>
                <div className="inspector-section">
                  <strong>Painting</strong>
                  <div className="paint-preview"><span style={{ width: Math.min(64, brushSize), height: Math.min(64, brushSize), background: tool === 'eraser' ? 'transparent' : brushColor }} /></div>
                  <div className="paint-values"><span>{tool === 'eraser' ? 'Eraser' : 'Brush'}</span><span>{brushSize} px</span></div>
                  <div className="paint-color-row">
                    <button type="button" className="hex-value" onClick={() => void copyHex()} title="Copy the colour value"><span style={{ background: brushColor }} />{brushColor.toUpperCase()}</button>
                    {eyeDropperSupported && <button type="button" className="inline-action" onClick={() => void pickFromScreen()} title="Pick a colour anywhere on the screen"><Pipette />Pick from screen</button>}
                  </div>
                </div>
              </>
            )}
          </aside>
        )}
      </section>

      <footer className="statusbar" data-print-background aria-hidden={printSnapshot ? true : undefined} hidden={advanced.mode === 'advanced'}>
        <div>
          {image ? <><strong>{image.sourceFormat.toUpperCase()}</strong><span>{dimensions.width.toLocaleString()} × {dimensions.height.toLocaleString()} px</span><span>{formatBytes(image.size)}</span></> : <span>Local image workspace</span>}
          {image && readout.readout && (
            <span className="pixel-readout">
              x {readout.readout.x} y {readout.readout.y}
              {readout.readout.color && <><i style={{ background: rgbHex(readout.readout.color) }} />{rgbHex(readout.readout.color).toUpperCase()}{readout.readout.color.a < 255 ? ` · ${Math.round(readout.readout.color.a / 2.55)}%` : ''}</>}
            </span>
          )}
          {working && <span className="working-state">{working.label}… {working.progress !== null ? `${Math.round(working.progress * 100)}%` : ''}</span>}
        </div>
        <div className="zoom-controls">
          <button disabled={!image} onClick={() => navigation.step(-1)} aria-label="Zoom out" title="Zoom out (Ctrl+-)"><ZoomOut /></button>
          <button disabled={!image} className="zoom-value" onClick={() => image && fitImage()} title="Fit (Ctrl+0)">{zoomLabel}</button>
          <button disabled={!image} onClick={() => navigation.step(1)} aria-label="Zoom in" title="Zoom in (Ctrl++)"><ZoomIn /></button>
          <button disabled={!image} onClick={() => fitImage()}>Fit</button>
          <button disabled={!image} onClick={() => navigation.actualPixels()} title="Actual pixels (Ctrl+1)">1:1</button>
        </div>
      </footer>

      {/* advanced-integration (WP1): the Advanced editor replaces the toolbar, content and status bar;
          it renders after them so the Simple canvas stays the first canvas in the document. */}
      {advanced.view}

      {draggingFile && (
        <div className="drop-overlay" onDragLeave={(event) => { if (event.currentTarget === event.target) setDraggingFile(false) }}>
          <div><ImagePlus /><strong>Drop to open</strong><span>Your current image stays safe until you confirm.</span></div>
        </div>
      )}
      {contextMenu && image && (
        <div ref={contextMenuRef} className="image-context-menu" role="menu" aria-label="Image selection menu" style={{ left: contextMenu.x, top: contextMenu.y }} onContextMenu={(event) => event.preventDefault()}>
          <button role="menuitem" onClick={() => { setContextMenu(null); void copyImageToClipboard() }}><Copy /><span>{tool === 'crop' && cropRect ? 'Copy selected area' : 'Copy image'}</span><kbd>Ctrl C</kbd></button>
          {!(tool === 'crop' && cropRect) && <button role="menuitem" onClick={() => { setContextMenu(null); void pasteFromClipboard() }}><ClipboardPaste /><span>Paste</span><kbd>Ctrl V</kbd></button>}
          <div role="separator" />
          <button role="menuitem" disabled={!canUndo} onClick={() => { setContextMenu(null); undo() }}><Undo2 /><span>Undo</span><kbd>Ctrl Z</kbd></button>
          <button role="menuitem" disabled={!canRedo} onClick={() => { setContextMenu(null); redo() }}><Redo2 /><span>Redo</span><kbd>Ctrl Y</kbd></button>
          {tool === 'crop' && cropRect
            ? <><div role="separator" /><button role="menuitem" onClick={() => { setContextMenu(null); void applyCrop() }}><Crop /><span>Apply crop</span><kbd>Enter</kbd></button></>
            : <><div role="separator" /><button role="menuitem" onClick={() => { setContextMenu(null); void openResize() }}><Scaling /><span>Resize…</span><kbd /></button></>}
        </div>
      )}
      {toast && <div className="toast" data-tone={toast.tone} role={toast.tone === 'error' ? 'alert' : 'status'}>{toast.message}</div>}
      {printSnapshot && (
        <PrintDialog
          imageUrl={printSnapshot.url}
          imageName={printSnapshot.name}
          imageWidth={printSnapshot.width}
          imageHeight={printSnapshot.height}
          busy={printing}
          onPrint={(settings) => void runPrint(settings)}
          onClose={closePrintDialog}
        />
      )}
      {advanced.dialog}
      {resizeOpen && image && (
        <ResizeDialog
          width={dimensions.width}
          height={dimensions.height}
          busy={busy}
          progress={working?.progress ?? null}
          onCancel={() => setResizeOpen(false)}
          onApply={(width, height, method) => void applyResize(width, height, method)}
        />
      )}
      {confirmState && (
        <div className="modal-backdrop">
          <div className="modal" role="dialog" aria-modal="true" aria-labelledby="confirm-title">
            <div className="modal-icon">{confirmState.cropPending ? <Crop /> : <Save />}</div>
            <h2 id="confirm-title">{confirmState.title}</h2>
            <p>{confirmState.detail}</p>
            <div className="modal-actions">
              <button disabled={saving} onClick={() => setConfirmState(null)}>Cancel</button>
              <button disabled={saving} onClick={() => void confirmDiscard()}>Don't save</button>
              <button className="modal-primary" autoFocus disabled={saving} onClick={() => void confirmSave()}>{saving ? 'Saving…' : confirmState.cropPending ? 'Apply crop and save' : 'Save'}</button>
            </div>
          </div>
        </div>
      )}
      {choiceState && (
        <div className="modal-backdrop choice-backdrop">
          <div className="modal choice-modal" role="dialog" aria-modal="true" aria-labelledby="choice-title" aria-describedby="choice-detail">
            <div className="modal-icon">{choiceState.tone === 'crop' ? <Crop /> : <TriangleAlert />}</div>
            <h2 id="choice-title">{choiceState.title}</h2>
            <p id="choice-detail">{choiceState.detail}</p>
            <div className="modal-actions">
              {choiceState.options.map((option) => (
                <button
                  key={option.value}
                  className={option.primary ? 'modal-primary' : ''}
                  data-choice={option.value}
                  autoFocus={option.primary}
                  onClick={() => choiceState.resolve(option.value)}
                >{option.label}</button>
              ))}
            </div>
          </div>
        </div>
      )}
      {image && <span className="sr-only" aria-live="polite">{working ? `${working.label}` : ''}</span>}
    </main>
  )
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><App /></React.StrictMode>)
