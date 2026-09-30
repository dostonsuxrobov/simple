import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import {
  Brush,
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Copy,
  Crop,
  Eraser,
  FileDown,
  FileImage,
  FolderOpen,
  ImagePlus,
  Maximize2,
  Minus,
  MousePointer2,
  Printer,
  Redo2,
  RotateCcw,
  RotateCw,
  Save,
  Undo2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'
import { INSPECTOR_STORAGE_KEY, inspectorOpenFromStored, storedInspectorOpen } from './inspector-preference.js'
import { PrintDialog } from './PrintDialog'
import type { ImagePrintSettings } from '../electron/print-layout.mjs'
import './styles.css'

type Tool = 'view' | 'crop' | 'brush' | 'eraser'
type SaveFormat = 'png' | 'jpeg' | 'webp'
type ExportFormat = SaveFormat | 'pdf'

interface OpenImage {
  name: string
  path: string | null
  size: number
  sourceFormat: string
  saveFormat: SaveFormat
  mime: string
  hasAlpha: boolean
}

interface Dimensions { width: number; height: number }
interface Point { x: number; y: number }
interface CropRect { x: number; y: number; width: number; height: number }
interface Snapshot { width: number; height: number; pixels: ImageData; revision: number; hasAlpha: boolean }
interface CropDrag { handle: string; start: Point; initial: CropRect }
interface ConfirmState { title: string; detail: string; action: () => void | Promise<void> }
interface PrintSnapshot { data: Uint8Array; url: string; name: string; width: number; height: number }
interface ImageContextMenu { x: number; y: number }

const SUPPORTED = ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'svg', 'avif']
const EXPORT_OPTIONS: ReadonlyArray<{ format: ExportFormat; label: string; detail: string }> = [
  { format: 'png', label: 'PNG image', detail: 'Lossless and transparency-safe' },
  { format: 'jpeg', label: 'JPEG image', detail: 'Compact for photos; transparent areas become white' },
  { format: 'webp', label: 'WebP image', detail: 'Smaller modern image with transparency' },
  { format: 'pdf', label: 'PDF document', detail: 'One page sized to the image aspect ratio' },
]
const MAX_ZOOM = 8
const MIN_ZOOM = 0.05
const MAX_CANVAS_DIMENSION = 20_000
const MAX_CANVAS_PIXELS = 50_000_000

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
  return format === 'jpeg' ? 'image/jpeg' : `image/${format}`
}

function canvasToBlob(canvas: HTMLCanvasElement, format: SaveFormat): Promise<Blob> {
  let exportCanvas = canvas
  if (format === 'jpeg') {
    exportCanvas = document.createElement('canvas')
    exportCanvas.width = canvas.width
    exportCanvas.height = canvas.height
    const context = exportCanvas.getContext('2d', { alpha: false })
    if (!context) return Promise.reject(new Error('The image canvas is unavailable.'))
    context.fillStyle = '#ffffff'
    context.fillRect(0, 0, exportCanvas.width, exportCanvas.height)
    context.drawImage(canvas, 0, 0)
  }
  return new Promise<Blob>((resolve, reject) => {
    exportCanvas.toBlob(
      (blob) => blob ? resolve(blob) : reject(new Error('The image could not be encoded.')),
      outputMime(format),
      format === 'png' ? undefined : 0.92,
    )
  }).finally(() => {
    if (exportCanvas !== canvas) { exportCanvas.width = 1; exportCanvas.height = 1 }
  })
}

function canvasHasTransparency(canvas: HTMLCanvasElement) {
  const context = canvas.getContext('2d', { willReadFrequently: true })
  if (!context) return false
  const rowsPerTile = Math.max(1, Math.min(256, Math.floor(16_000_000 / Math.max(4, canvas.width * 4))))
  for (let top = 0; top < canvas.height; top += rowsPerTile) {
    const height = Math.min(rowsPerTile, canvas.height - top)
    const pixels = context.getImageData(0, top, canvas.width, height).data
    for (let index = 3; index < pixels.length; index += 4) {
      if (pixels[index] < 255) return true
    }
  }
  return false
}

function App() {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const contextMenuRef = useRef<HTMLDivElement>(null)
  const viewportRef = useRef<HTMLDivElement>(null)
  const undoRef = useRef<Snapshot[]>([])
  const redoRef = useRef<Snapshot[]>([])
  const paintingRef = useRef(false)
  const lastPointRef = useRef<Point | null>(null)
  const cropDragRef = useRef<CropDrag | null>(null)
  const toastTimerRef = useRef<number | null>(null)
  const imageRef = useRef<OpenImage | null>(null)
  const dirtyRef = useRef(false)
  const documentIdRef = useRef(0)
  const loadRequestRef = useRef(0)
  const originalImageRef = useRef<{ blob: Blob; format: SaveFormat; printable: boolean } | null>(null)
  const nextRevisionRef = useRef(0)
  const contentRevisionRef = useRef(0)
  const savedRevisionRef = useRef(0)
  const saveGroupRef = useRef<HTMLDivElement>(null)
  const exportGroupRef = useRef<HTMLDivElement>(null)
  const printButtonRef = useRef<HTMLButtonElement>(null)
  const printFocusReturnRef = useRef<HTMLElement | null>(null)
  const printPendingFocusRef = useRef<HTMLElement | null>(null)
  const printDialogOpenRef = useRef(false)

  const [image, setImage] = useState<OpenImage | null>(null)
  const [album, setAlbum] = useState<{ paths: string[]; index: number }>({ paths: [], index: -1 })
  const [dimensions, setDimensions] = useState<Dimensions>({ width: 0, height: 0 })
  const [dirty, setDirtyState] = useState(false)
  const [saving, setSaving] = useState(false)
  const [exporting, setExporting] = useState<ExportFormat | null>(null)
  const [printing, setPrinting] = useState(false)
  const [printSnapshot, setPrintSnapshot] = useState<PrintSnapshot | null>(null)
  const [tool, setTool] = useState<Tool>('view')
  const [brushColor, setBrushColor] = useState('#111111')
  const [brushSize, setBrushSize] = useState(12)
  const [zoom, setZoom] = useState(1)
  const [fitMode, setFitMode] = useState(true)
  const [cropRect, setCropRect] = useState<CropRect | null>(null)
  const [historyVersion, setHistoryVersion] = useState(0)
  const [isMaximized, setIsMaximized] = useState(false)
  const [draggingFile, setDraggingFile] = useState(false)
  const [toast, setToast] = useState<{ message: string; tone: 'normal' | 'error' } | null>(null)
  const [confirmState, setConfirmState] = useState<ConfirmState | null>(null)
  const [saveMenu, setSaveMenu] = useState(false)
  const [exportMenu, setExportMenu] = useState(false)
  const [contextMenu, setContextMenu] = useState<ImageContextMenu | null>(null)
  const [inspectorOpen, setInspectorOpen] = useState(() => {
    try { return inspectorOpenFromStored(window.localStorage.getItem(INSPECTOR_STORAGE_KEY)) } catch { return true }
  })

  imageRef.current = image
  dirtyRef.current = dirty
  printDialogOpenRef.current = Boolean(printSnapshot)

  const reflectDirty = useCallback(() => {
    const value = contentRevisionRef.current !== savedRevisionRef.current
    dirtyRef.current = value
    setDirtyState(value)
  }, [])

  const markEdited = useCallback(() => {
    nextRevisionRef.current += 1
    contentRevisionRef.current = nextRevisionRef.current
    reflectDirty()
  }, [reflectDirty])

  const notify = useCallback((message: string, tone: 'normal' | 'error' = 'normal') => {
    if (toastTimerRef.current) window.clearTimeout(toastTimerRef.current)
    setToast({ message, tone })
    toastTimerRef.current = window.setTimeout(() => setToast(null), 3200)
  }, [])

  const fitImage = useCallback((width = dimensions.width, height = dimensions.height) => {
    const viewport = viewportRef.current
    if (!viewport || !width || !height) return
    const availableWidth = Math.max(80, viewport.clientWidth - 88)
    const availableHeight = Math.max(80, viewport.clientHeight - 88)
    setZoom(clamp(Math.min(availableWidth / width, availableHeight / height, 1), MIN_ZOOM, MAX_ZOOM))
    setFitMode(true)
  }, [dimensions.height, dimensions.width])

  const changeZoom = useCallback((next: number) => {
    setZoom(clamp(next, MIN_ZOOM, MAX_ZOOM))
    setFitMode(false)
  }, [])

  const context2d = useCallback(() => {
    const context = canvasRef.current?.getContext('2d', { willReadFrequently: true })
    if (!context) throw new Error('The image canvas is unavailable.')
    return context
  }, [])

  const snapshot = useCallback((): Snapshot => {
    const canvas = canvasRef.current
    if (!canvas) throw new Error('There is no image to edit.')
    return {
      width: canvas.width,
      height: canvas.height,
      pixels: context2d().getImageData(0, 0, canvas.width, canvas.height),
      revision: contentRevisionRef.current,
      hasAlpha: imageRef.current?.hasAlpha ?? false,
    }
  }, [context2d])

  const historyLimit = useCallback(() => {
    const pixels = Math.max(1, dimensions.width * dimensions.height)
    return clamp(Math.floor((160 * 1024 * 1024) / (pixels * 4)), 1, 16)
  }, [dimensions.height, dimensions.width])

  const pushUndo = useCallback(() => {
    undoRef.current.push(snapshot())
    const limit = historyLimit()
    if (undoRef.current.length > limit) undoRef.current.splice(0, undoRef.current.length - limit)
    redoRef.current = []
    setHistoryVersion((version) => version + 1)
  }, [historyLimit, snapshot])

  const restoreSnapshot = useCallback((entry: Snapshot) => {
    const canvas = canvasRef.current
    if (!canvas) return
    canvas.width = entry.width
    canvas.height = entry.height
    const context = context2d()
    context.putImageData(entry.pixels, 0, 0)
    contentRevisionRef.current = entry.revision
    setDimensions({ width: entry.width, height: entry.height })
    setImage((current) => current ? { ...current, hasAlpha: entry.hasAlpha } : current)
    setCropRect(null)
    reflectDirty()
    if (fitMode) requestAnimationFrame(() => fitImage(entry.width, entry.height))
  }, [context2d, fitImage, fitMode, reflectDirty])

  const undo = useCallback(() => {
    const previous = undoRef.current.pop()
    if (!previous) return
    redoRef.current.push(snapshot())
    restoreSnapshot(previous)
    setHistoryVersion((version) => version + 1)
  }, [restoreSnapshot, snapshot])

  const redo = useCallback(() => {
    const next = redoRef.current.pop()
    if (!next) return
    undoRef.current.push(snapshot())
    restoreSnapshot(next)
    setHistoryVersion((version) => version + 1)
  }, [restoreSnapshot, snapshot])

  const updateAlphaMetadata = useCallback(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const hasAlpha = canvasHasTransparency(canvas)
    setImage((current) => current ? { ...current, hasAlpha } : current)
  }, [])

  const loadPayloadNow = useCallback(async (payload: ImagePayload, request: number) => {
    if (request !== loadRequestRef.current) return
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
      canvas.width = decoded.naturalWidth
      canvas.height = decoded.naturalHeight
      const context = context2d()
      context.clearRect(0, 0, canvas.width, canvas.height)
      context.drawImage(decoded, 0, 0)
      // JPEG cannot contain transparency. Avoid a full pixel readback and scan
      // just to rediscover this on every large photograph.
      const hasAlpha = /^(jpe?g)$/i.test(payload.format) ? false : canvasHasTransparency(canvas)
      const saveFormat = defaultSaveFormat(payload.format)
      originalImageRef.current = {
        blob,
        format: saveFormat,
        printable: /^(png|jpe?g)$/i.test(payload.format),
      }
      documentIdRef.current += 1
      nextRevisionRef.current = 0
      contentRevisionRef.current = 0
      savedRevisionRef.current = 0
      const nextImage: OpenImage = {
        name: payload.name,
        path: payload.directSave ? payload.path : null,
        size: payload.size,
        sourceFormat: payload.format,
        saveFormat,
        mime: payload.mime,
        hasAlpha,
      }
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
      requestAnimationFrame(() => fitImage(canvas.width, canvas.height))
      if (payload.format === 'gif') notify('GIF opened as an editable first frame.')
      else notify(`${payload.format.toUpperCase()} opened locally.`)
    } finally {
      URL.revokeObjectURL(url)
    }
  }, [context2d, fitImage, notify])

  const requestProceed = useCallback((title: string, detail: string, action: () => void | Promise<void>) => {
    if (printDialogOpenRef.current) return
    if (!dirtyRef.current) {
      void action()
      return
    }
    setConfirmState({ title, detail, action })
  }, [])

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

  const stepAlbum = useCallback((delta: number) => {
    if (album.paths.length < 2 || album.index < 0) return
    const next = (album.index + delta + album.paths.length) % album.paths.length
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

  const saveImage = useCallback(async (forceDialog = false): Promise<boolean> => {
    const canvas = canvasRef.current
    const current = imageRef.current
    if (!canvas || !current || saving || exporting || printing) return false
    const documentId = documentIdRef.current
    const revision = contentRevisionRef.current
    const saveFormat = current.saveFormat
    setSaving(true)
    setSaveMenu(false)
    try {
      const original = originalImageRef.current
      // An unchanged Save/Save As must not recompress a JPEG or discard its
      // metadata. Undo back to revision zero also restores the exact source.
      const blob = revision === 0 && original?.format === saveFormat
        && original.blob.type === outputMime(saveFormat)
        ? original.blob : await canvasToBlob(canvas, saveFormat)
      const result = await window.simpleImage.saveImage({
        data: new Uint8Array(await blob.arrayBuffer()),
        path: current.path,
        name: current.name,
        format: saveFormat,
        forceDialog: forceDialog || !current.path,
      })
      if (!result) return false
      if (documentIdRef.current === documentId) {
        setImage((value) => value ? {
          ...value,
          name: result.name,
          path: result.path,
          size: result.size,
          sourceFormat: result.format,
          mime: outputMime(saveFormat),
        } : value)
        window.simpleImage.setTitle(result.name)
        savedRevisionRef.current = revision
        reflectDirty()
        const fullySaved = contentRevisionRef.current === revision
        notify(fullySaved
          ? `Saved ${result.name}`
          : `Saved ${result.name}; newer edits remain unsaved.`)
        return fullySaved
      } else {
        notify(`Saved ${result.name}`)
        return false
      }
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be saved.', 'error')
      return false
    } finally {
      setSaving(false)
    }
  }, [exporting, notify, printing, reflectDirty, saving])

  const openPrintDialog = useCallback(async (invoker?: HTMLElement | null) => {
    const canvas = canvasRef.current
    if (!canvas || !image || printing || exporting || saving) return
    const activeElement = document.activeElement instanceof HTMLElement ? document.activeElement : null
    printFocusReturnRef.current = invoker || activeElement || printButtonRef.current
    setPrinting(true)
    setSaveMenu(false)
    setExportMenu(false)
    const documentId = documentIdRef.current
    const revision = contentRevisionRef.current
    const width = canvas.width
    const height = canvas.height
    try {
      const original = originalImageRef.current
      const blob = revision === 0 && original?.printable ? original.blob : await canvasToBlob(canvas, 'png')
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
  }, [exporting, image, notify, printing, saving])

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
    if (!canvas || !current || exporting || saving || printing) return
    setExporting(format)
    setSaveMenu(false)
    setExportMenu(false)
    try {
      if (format === 'pdf') {
        const blob = await canvasToBlob(canvas, 'png')
        const result = await window.simpleImage.convertToPdf({
          data: new Uint8Array(await blob.arrayBuffer()),
          name: current.name,
        })
        if (result) notify(`Exported ${result.name}`)
        else notify('PDF export was canceled.')
        return
      }
      const blob = await canvasToBlob(canvas, format)
      const result = await window.simpleImage.saveImage({
        data: new Uint8Array(await blob.arrayBuffer()),
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
  }, [exporting, notify, printing, saving])

  const copyImageToClipboard = useCallback(async () => {
    const canvas = canvasRef.current
    if (!canvas || !image || saving || exporting || printing) return false
    let copyCanvas: HTMLCanvasElement | null = null
    try {
      let source = canvas
      let copiedSelection = false
      if (tool === 'crop' && cropRect) {
        const x = clamp(Math.round(cropRect.x), 0, Math.max(0, canvas.width - 1))
        const y = clamp(Math.round(cropRect.y), 0, Math.max(0, canvas.height - 1))
        const width = Math.min(Math.max(1, Math.round(cropRect.width)), canvas.width - x)
        const height = Math.min(Math.max(1, Math.round(cropRect.height)), canvas.height - y)
        copyCanvas = document.createElement('canvas')
        copyCanvas.width = width
        copyCanvas.height = height
        const context = copyCanvas.getContext('2d')
        if (!context) throw new Error('The selected image area could not be prepared.')
        context.drawImage(canvas, x, y, width, height, 0, 0, width, height)
        source = copyCanvas
        copiedSelection = true
      }
      const blob = await canvasToBlob(source, 'png')
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
  }, [cropRect, exporting, image, notify, printing, saving, tool])

  const rotate = useCallback((clockwise: boolean) => {
    const canvas = canvasRef.current
    if (!canvas || !image) return
    pushUndo()
    const copy = document.createElement('canvas')
    copy.width = canvas.width
    copy.height = canvas.height
    copy.getContext('2d')?.drawImage(canvas, 0, 0)
    canvas.width = copy.height
    canvas.height = copy.width
    const context = context2d()
    context.translate(canvas.width / 2, canvas.height / 2)
    context.rotate((clockwise ? 1 : -1) * Math.PI / 2)
    context.drawImage(copy, -copy.width / 2, -copy.height / 2)
    context.setTransform(1, 0, 0, 1, 0, 0)
    copy.width = 1
    copy.height = 1
    setDimensions({ width: canvas.width, height: canvas.height })
    setCropRect(null)
    markEdited()
    if (fitMode) requestAnimationFrame(() => fitImage(canvas.width, canvas.height))
  }, [context2d, fitImage, fitMode, image, markEdited, pushUndo])

  const beginCrop = useCallback(() => {
    if (!image) return
    setTool('crop')
    setCropRect((current) => current || {
      x: Math.round(dimensions.width * 0.1),
      y: Math.round(dimensions.height * 0.1),
      width: Math.max(1, Math.round(dimensions.width * 0.8)),
      height: Math.max(1, Math.round(dimensions.height * 0.8)),
    })
  }, [dimensions.height, dimensions.width, image])

  const applyCrop = useCallback(() => {
    const canvas = canvasRef.current
    if (!canvas || !cropRect) return
    const crop = {
      x: clamp(Math.round(cropRect.x), 0, canvas.width - 1),
      y: clamp(Math.round(cropRect.y), 0, canvas.height - 1),
      width: clamp(Math.round(cropRect.width), 1, canvas.width),
      height: clamp(Math.round(cropRect.height), 1, canvas.height),
    }
    crop.width = Math.min(crop.width, canvas.width - crop.x)
    crop.height = Math.min(crop.height, canvas.height - crop.y)
    pushUndo()
    const copy = document.createElement('canvas')
    copy.width = crop.width
    copy.height = crop.height
    copy.getContext('2d')?.drawImage(canvas, crop.x, crop.y, crop.width, crop.height, 0, 0, crop.width, crop.height)
    canvas.width = crop.width
    canvas.height = crop.height
    context2d().drawImage(copy, 0, 0)
    copy.width = 1
    copy.height = 1
    setDimensions({ width: crop.width, height: crop.height })
    setCropRect(null)
    setTool('view')
    markEdited()
    updateAlphaMetadata()
    if (fitMode) requestAnimationFrame(() => fitImage(crop.width, crop.height))
  }, [context2d, cropRect, fitImage, fitMode, markEdited, pushUndo, updateAlphaMetadata])

  const canvasPoint = useCallback((event: React.PointerEvent): Point => {
    const canvas = canvasRef.current
    if (!canvas) return { x: 0, y: 0 }
    const bounds = canvas.getBoundingClientRect()
    return {
      x: clamp((event.clientX - bounds.left) * canvas.width / bounds.width, 0, canvas.width),
      y: clamp((event.clientY - bounds.top) * canvas.height / bounds.height, 0, canvas.height),
    }
  }, [])

  const drawSegment = useCallback((from: Point, to: Point) => {
    const context = context2d()
    context.save()
    context.globalCompositeOperation = tool === 'eraser' ? 'destination-out' : 'source-over'
    context.strokeStyle = brushColor
    context.fillStyle = brushColor
    context.lineWidth = brushSize
    context.lineCap = 'round'
    context.lineJoin = 'round'
    context.beginPath()
    context.moveTo(from.x, from.y)
    context.lineTo(to.x, to.y)
    context.stroke()
    if (from.x === to.x && from.y === to.y) {
      context.beginPath()
      context.arc(to.x, to.y, brushSize / 2, 0, Math.PI * 2)
      context.fill()
    }
    context.restore()
  }, [brushColor, brushSize, context2d, tool])

  const onCanvasPointerDown = useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (event.button !== 0 || !event.isPrimary) return
    if (tool !== 'brush' && tool !== 'eraser') return
    event.currentTarget.setPointerCapture(event.pointerId)
    pushUndo()
    const point = canvasPoint(event)
    paintingRef.current = true
    lastPointRef.current = point
    drawSegment(point, point)
    markEdited()
  }, [canvasPoint, drawSegment, markEdited, pushUndo, tool])

  const onCanvasPointerMove = useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (!paintingRef.current || (tool !== 'brush' && tool !== 'eraser')) return
    const point = canvasPoint(event)
    const previous = lastPointRef.current || point
    drawSegment(previous, point)
    lastPointRef.current = point
  }, [canvasPoint, drawSegment, tool])

  const finishPainting = useCallback(() => {
    if (!paintingRef.current) return
    paintingRef.current = false
    lastPointRef.current = null
    updateAlphaMetadata()
  }, [updateAlphaMetadata])

  const cropPoint = useCallback((event: React.PointerEvent): Point => {
    const canvas = canvasRef.current
    if (!canvas) return { x: 0, y: 0 }
    const bounds = canvas.getBoundingClientRect()
    return {
      x: clamp((event.clientX - bounds.left) * canvas.width / bounds.width, 0, canvas.width),
      y: clamp((event.clientY - bounds.top) * canvas.height / bounds.height, 0, canvas.height),
    }
  }, [])

  const startCropDrag = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0 || !event.isPrimary) return
    if (!cropRect) return
    event.currentTarget.setPointerCapture(event.pointerId)
    const element = event.target as HTMLElement
    const handle = element.dataset.handle || (element.closest('.crop-box') ? 'move' : 'new')
    const start = cropPoint(event)
    cropDragRef.current = { handle, start, initial: handle === 'new' ? { x: start.x, y: start.y, width: 1, height: 1 } : cropRect }
    if (handle === 'new') setCropRect({ x: start.x, y: start.y, width: 1, height: 1 })
  }, [cropPoint, cropRect])

  const moveCropDrag = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const drag = cropDragRef.current
    if (!drag) return
    const point = cropPoint(event)
    const dx = point.x - drag.start.x
    const dy = point.y - drag.start.y
    let { x, y, width, height } = drag.initial
    if (drag.handle === 'new') {
      x = Math.min(drag.start.x, point.x)
      y = Math.min(drag.start.y, point.y)
      width = Math.abs(point.x - drag.start.x)
      height = Math.abs(point.y - drag.start.y)
    } else if (drag.handle === 'move') {
      x = clamp(x + dx, 0, dimensions.width - width)
      y = clamp(y + dy, 0, dimensions.height - height)
    } else {
      if (drag.handle.includes('w')) { x += dx; width -= dx }
      if (drag.handle.includes('e')) width += dx
      if (drag.handle.includes('n')) { y += dy; height -= dy }
      if (drag.handle.includes('s')) height += dy
      if (width < 8) { if (drag.handle.includes('w')) x -= 8 - width; width = 8 }
      if (height < 8) { if (drag.handle.includes('n')) y -= 8 - height; height = 8 }
      x = clamp(x, 0, dimensions.width - 1)
      y = clamp(y, 0, dimensions.height - 1)
      width = clamp(width, 1, dimensions.width - x)
      height = clamp(height, 1, dimensions.height - y)
    }
    setCropRect({ x, y, width, height })
  }, [cropPoint, dimensions.height, dimensions.width])

  const finishCropDrag = useCallback(() => { cropDragRef.current = null }, [])

  const chooseTool = useCallback((next: Tool) => {
    setTool(next)
    if (next !== 'crop') setCropRect(null)
    else beginCrop()
  }, [beginCrop])

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
    const action = confirmState?.action
    if (!action) return
    if (await saveImage(false)) {
      setConfirmState(null)
      await action()
    }
  }, [confirmState, saveImage])

  const toggleInspector = useCallback(() => setInspectorOpen((value) => !value), [])

  useEffect(() => window.simpleImage.onMaximized(setIsMaximized), [])
  useEffect(() => window.simpleImage.onOpenExternal(openPath), [openPath])
  useEffect(() => window.simpleImage.onCloseRequested(() => {
    requestProceed('Save changes before closing?', 'Your edits will be lost if you close without saving.', () => window.simpleImage.confirmClose())
  }), [requestProceed])

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
    const observer = new ResizeObserver(() => fitImage())
    if (viewportRef.current) observer.observe(viewportRef.current)
    return () => observer.disconnect()
  }, [fitImage, fitMode, image])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (printSnapshot) {
        event.preventDefault()
        if (event.key === 'Escape' && !printing) closePrintDialog()
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'p' && !printing) {
          const printButton = document.querySelector<HTMLButtonElement>('.image-print-dialog .print-submit')
          printButton?.click()
        }
        return
      }
      if (confirmState) {
        if (event.key === 'Escape' && !saving) setConfirmState(null)
        if (event.key === 'Tab' || event.key === 'Enter' || event.key === ' ') return
        event.preventDefault()
        return
      }
      if (saving || exporting) {
        if (event.key === 'Tab') return
        event.preventDefault()
        return
      }
      if (contextMenu) {
        const menu = contextMenuRef.current
        const buttons = [...menu?.querySelectorAll<HTMLButtonElement>('button:not(:disabled)') ?? []]
        const current = Math.max(0, buttons.indexOf(document.activeElement as HTMLButtonElement))
        const key = event.key.toLowerCase()
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
      const control = event.ctrlKey || event.metaKey
      const target = event.target instanceof HTMLElement ? event.target : null
      const textInput = Boolean(target?.closest('input, textarea, [contenteditable="true"]'))
      const nativeTextSelected = Boolean(window.getSelection()?.toString().trim())
      if (control && event.key.toLowerCase() === 'c' && image && !textInput && !nativeTextSelected) { event.preventDefault(); void copyImageToClipboard() }
      if (!control && !event.altKey && !event.shiftKey && tool === 'view' && !textInput && image && album.paths.length > 1
        && (event.key === 'ArrowLeft' || event.key === 'ArrowRight')) {
        const viewport = viewportRef.current
        const scrollable = viewport ? viewport.scrollWidth > viewport.clientWidth + 1 : false
        if (!scrollable) { event.preventDefault(); stepAlbum(event.key === 'ArrowRight' ? 1 : -1) }
      }
      if (control && event.key.toLowerCase() === 'o') { event.preventDefault(); void openDialog() }
      if (control && event.key.toLowerCase() === 's') { event.preventDefault(); void saveImage(event.shiftKey) }
      if (control && event.shiftKey && event.key.toLowerCase() === 'e' && image) {
        event.preventDefault()
        setSaveMenu(false)
        setExportMenu((value) => !value)
      }
      if (control && event.key.toLowerCase() === 'p') { event.preventDefault(); void openPrintDialog() }
      if (control && !event.shiftKey && event.key.toLowerCase() === 'z') { event.preventDefault(); undo() }
      if (control && (event.key.toLowerCase() === 'y' || (event.shiftKey && event.key.toLowerCase() === 'z'))) { event.preventDefault(); redo() }
      if (control && (event.key === '+' || event.key === '=')) { event.preventDefault(); changeZoom(zoom * 1.2) }
      if (control && event.key === '-') { event.preventDefault(); changeZoom(zoom / 1.2) }
      if (control && event.key === '0') { event.preventDefault(); fitImage() }
      if (event.key === 'F4' && image && !event.repeat) { event.preventDefault(); toggleInspector() }
      if (event.key === 'Escape') {
        setSaveMenu(false)
        setExportMenu(false)
        if (tool === 'crop') { setCropRect(null); setTool('view') }
      }
      if (event.key === 'Enter' && tool === 'crop' && cropRect) applyCrop()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [album, stepAlbum, applyCrop, changeZoom, closePrintDialog, confirmState, contextMenu, copyImageToClipboard, cropRect, exporting, fitImage, image, openDialog, openPrintDialog, printSnapshot, printing, redo, saveImage, saving, toggleInspector, tool, undo, zoom])

  const cropStyle = useMemo(() => cropRect && dimensions.width && dimensions.height ? {
    left: `${cropRect.x / dimensions.width * 100}%`,
    top: `${cropRect.y / dimensions.height * 100}%`,
    width: `${cropRect.width / dimensions.width * 100}%`,
    height: `${cropRect.height / dimensions.height * 100}%`,
  } : undefined, [cropRect, dimensions.height, dimensions.width])

  const megapixels = dimensions.width && dimensions.height ? dimensions.width * dimensions.height / 1_000_000 : 0
  const canUndo = undoRef.current.length > 0
  const canRedo = redoRef.current.length > 0
  void historyVersion

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
          {image && <span className={`save-state ${dirty ? 'is-dirty' : ''}`}>{saving ? 'Saving…' : dirty ? 'Modified' : 'Saved'}</span>}
        </div>
        <div className="title-actions">
          <button title="Open image (Ctrl+O)" onClick={() => void openDialog()}><FolderOpen /><span>Open</span></button>
          <div ref={saveGroupRef} className="save-group">
            <button className="primary-button" disabled={!image || saving || Boolean(exporting) || printing} title="Save (Ctrl+S)" onClick={() => void saveImage(false)}><Save /><span>{saving ? 'Saving…' : 'Save'}</span></button>
            <button
              className="primary-button split-button"
              disabled={!image || saving || Boolean(exporting) || printing}
              title="Save options"
              aria-haspopup="menu"
              aria-expanded={saveMenu}
              onClick={() => { setExportMenu(false); setSaveMenu((value) => !value) }}
            ><ChevronDown /></button>
            {saveMenu && (
              <div className="save-menu" role="menu" aria-label="Save options">
                <button role="menuitem" onClick={() => void saveImage(true)}><Save /><span><strong>Save as…</strong><small>Move or rename the working image</small></span></button>
              </div>
            )}
          </div>
          <div ref={exportGroupRef} className="export-group">
            <button
              className="export-button"
              disabled={!image || Boolean(exporting) || printing || saving}
              title="Export As (Ctrl+Shift+E)"
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
                {EXPORT_OPTIONS.map((option) => (
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
          <button ref={printButtonRef} disabled={!image || printing || Boolean(exporting) || saving} title="Print (Ctrl+P)" onClick={() => void openPrintDialog(printButtonRef.current)}><Printer /><span>{printing ? 'Preparing…' : 'Print'}</span></button>
        </div>
        <div className="window-actions">
          <button aria-label="Minimize" onClick={() => window.simpleImage.minimize()}><Minus /></button>
          <button aria-label={isMaximized ? 'Restore' : 'Maximize'} onClick={() => window.simpleImage.toggleMaximize()}><Maximize2 /></button>
          <button className="close-button" aria-label="Close" onClick={() => requestProceed('Save changes before closing?', 'Your edits will be lost if you close without saving.', () => window.simpleImage.confirmClose())}><X /></button>
        </div>
      </header>

      <section className="toolbar" data-print-background aria-label="Image tools" aria-hidden={printSnapshot ? true : undefined}>
        <div className="tool-group">
          <button className={tool === 'view' ? 'active' : ''} disabled={!image} onClick={() => chooseTool('view')} title="View"><MousePointer2 /><span>View</span></button>
          <button className={tool === 'crop' ? 'active' : ''} disabled={!image} onClick={beginCrop} title="Crop"><Crop /><span>Crop</span></button>
          <button disabled={!image} onClick={() => rotate(false)} title="Rotate left"><RotateCcw /></button>
          <button disabled={!image} onClick={() => rotate(true)} title="Rotate right"><RotateCw /></button>
        </div>
        <div className="toolbar-rule" />
        <div className="tool-group">
          <button className={tool === 'brush' ? 'active' : ''} disabled={!image} onClick={() => chooseTool('brush')} title="Brush"><Brush /><span>Brush</span></button>
          <button className={tool === 'eraser' ? 'active' : ''} disabled={!image} onClick={() => chooseTool('eraser')} title="Eraser"><Eraser /><span>Eraser</span></button>
          <label className="color-control" title="Brush color"><input type="color" value={brushColor} onChange={(event) => setBrushColor(event.target.value)} /><span style={{ background: brushColor }} /></label>
          <label className="size-control"><span>Size</span><input type="range" min="1" max="120" value={brushSize} onChange={(event) => setBrushSize(Number(event.target.value))} /><output>{brushSize}px</output></label>
        </div>
        <div className="toolbar-spacer" />
        <div className="tool-group">
          <button disabled={!canUndo} onClick={undo} title="Undo (Ctrl+Z)"><Undo2 /></button>
          <button disabled={!canRedo} onClick={redo} title="Redo (Ctrl+Y)"><Redo2 /></button>
          <button
            className={image && inspectorOpen ? 'active' : ''}
            disabled={!image}
            aria-controls="image-inspector"
            aria-expanded={Boolean(image && inspectorOpen)}
            onClick={toggleInspector}
            title={`${inspectorOpen ? 'Hide' : 'Show'} image details (F4)`}
          ><FileImage /><span>Details</span></button>
        </div>
        {tool === 'crop' && cropRect && (
          <div className="crop-actions">
            <span>{Math.round(cropRect.width)} × {Math.round(cropRect.height)} px</span>
            <button onClick={() => { setCropRect(null); setTool('view') }}>Cancel</button>
            <button className="apply-button" onClick={applyCrop}><Check /> Apply crop</button>
          </div>
        )}
      </section>

      <section className="content" data-print-background aria-hidden={printSnapshot ? true : undefined}>
        <div
          ref={viewportRef}
          className={`viewport ${image ? '' : 'is-empty'}`}
          onWheel={(event) => {
            if (!event.ctrlKey || !image) return
            event.preventDefault()
            changeZoom(zoom * (event.deltaY < 0 ? 1.12 : 1 / 1.12))
          }}
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
              <p>Open, inspect, crop, rotate, paint, print, or export a copy as PNG, JPEG, WebP, or PDF without uploading anything.</p>
              <div className="welcome-actions">
                <button className="welcome-primary" onClick={() => void openDialog()}><ImagePlus /> Open an image</button>
                <button onClick={() => void window.simpleImage.newWindow()}><FileImage /> New window</button>
              </div>
              <div className="format-list">PNG · JPEG · WebP · GIF · BMP · SVG · AVIF</div>
              <small>Or drop an image anywhere in this window</small>
            </div>
          )}
          <div
            className="canvas-stage"
            hidden={!image}
            style={{ minWidth: dimensions.width * zoom + 80, minHeight: dimensions.height * zoom + 80 }}
            onContextMenu={(event) => {
              event.preventDefault()
              if (!image || saving || exporting || printing || printSnapshot) return
              setSaveMenu(false)
              setExportMenu(false)
              setContextMenu({
                x: Math.max(8, Math.min(event.clientX, window.innerWidth - 224)),
                y: Math.max(8, Math.min(event.clientY, window.innerHeight - 190)),
              })
            }}
          >
            <div className="canvas-stack" style={{ width: dimensions.width * zoom, height: dimensions.height * zoom }}>
              <canvas
                ref={canvasRef}
                className={`image-canvas tool-${tool}`}
                style={{ width: dimensions.width * zoom, height: dimensions.height * zoom }}
                onPointerDown={onCanvasPointerDown}
                onPointerMove={onCanvasPointerMove}
                onPointerUp={finishPainting}
                onPointerCancel={finishPainting}
              />
              {tool === 'crop' && cropRect && (
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
          <aside id="image-inspector" className="inspector" aria-label="Image details" hidden={!inspectorOpen}>
            <div className="inspector-heading">
              <FileImage />
              <div><strong>Image details</strong><span>Read from this asset</span></div>
              <button className="inspector-close" type="button" aria-label="Close image details" title="Close image details (F4)" onClick={() => setInspectorOpen(false)}><X /></button>
            </div>
            <dl>
              <div><dt>Format</dt><dd>{image.sourceFormat.toUpperCase()}</dd></div>
              <div><dt>Dimensions</dt><dd>{dimensions.width.toLocaleString()} × {dimensions.height.toLocaleString()} px</dd></div>
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
            </div>
          </aside>
        )}
      </section>

      <footer className="statusbar" data-print-background aria-hidden={printSnapshot ? true : undefined}>
        <div>{image ? <><strong>{image.sourceFormat.toUpperCase()}</strong><span>{dimensions.width.toLocaleString()} × {dimensions.height.toLocaleString()} px</span><span>{formatBytes(image.size)}</span></> : <span>Local image workspace</span>}</div>
        <div className="zoom-controls">
          <button disabled={!image} onClick={() => changeZoom(zoom / 1.2)} aria-label="Zoom out"><ZoomOut /></button>
          <button disabled={!image} className="zoom-value" onClick={() => image && fitImage()}>{Math.round(zoom * 100)}%</button>
          <button disabled={!image} onClick={() => changeZoom(zoom * 1.2)} aria-label="Zoom in"><ZoomIn /></button>
          <button disabled={!image} onClick={() => fitImage()}>Fit</button>
          <button disabled={!image} onClick={() => changeZoom(1)}>1:1</button>
        </div>
      </footer>

      {draggingFile && (
        <div className="drop-overlay" onDragLeave={(event) => { if (event.currentTarget === event.target) setDraggingFile(false) }}>
          <div><ImagePlus /><strong>Drop to open</strong><span>Your current image stays safe until you confirm.</span></div>
        </div>
      )}
      {contextMenu && image && (
        <div ref={contextMenuRef} className="image-context-menu" role="menu" aria-label="Image selection menu" style={{ left: contextMenu.x, top: contextMenu.y }} onContextMenu={(event) => event.preventDefault()}>
          <button role="menuitem" onClick={() => { setContextMenu(null); void copyImageToClipboard() }}><Copy /><span>{tool === 'crop' && cropRect ? 'Copy selected area' : 'Copy image'}</span><kbd>Ctrl C</kbd></button>
          <div role="separator" />
          <button role="menuitem" disabled={!canUndo} onClick={() => { setContextMenu(null); undo() }}><Undo2 /><span>Undo</span><kbd>Ctrl Z</kbd></button>
          <button role="menuitem" disabled={!canRedo} onClick={() => { setContextMenu(null); redo() }}><Redo2 /><span>Redo</span><kbd>Ctrl Y</kbd></button>
          {tool === 'crop' && cropRect && <><div role="separator" /><button role="menuitem" onClick={() => { setContextMenu(null); applyCrop() }}><Crop /><span>Apply crop</span><kbd>Enter</kbd></button></>}
        </div>
      )}
      {toast && <div className="toast" data-tone={toast.tone}>{toast.message}</div>}
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
      {confirmState && (
        <div className="modal-backdrop">
          <div className="modal" role="dialog" aria-modal="true" aria-labelledby="confirm-title">
            <div className="modal-icon"><Save /></div>
            <h2 id="confirm-title">{confirmState.title}</h2>
            <p>{confirmState.detail}</p>
            <div className="modal-actions">
              <button disabled={saving} onClick={() => setConfirmState(null)}>Cancel</button>
              <button disabled={saving} onClick={() => void confirmDiscard()}>Don't save</button>
              <button className="modal-primary" disabled={saving} onClick={() => void confirmSave()}>{saving ? 'Saving…' : 'Save'}</button>
            </div>
          </div>
        </div>
      )}
    </main>
  )
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><App /></React.StrictMode>)
