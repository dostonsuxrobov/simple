import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import {
  Brush,
  Check,
  ChevronDown,
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
import './styles.css'

type Tool = 'view' | 'crop' | 'brush' | 'eraser'
type SaveFormat = 'png' | 'jpeg' | 'webp'

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
interface Snapshot { width: number; height: number; pixels: ImageData; revision: number }
interface CropDrag { handle: string; start: Point; initial: CropRect }
interface ConfirmState { title: string; detail: string; action: () => void | Promise<void> }

const SUPPORTED = ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'svg', 'avif']
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
  return new Promise((resolve, reject) => {
    exportCanvas.toBlob(
      (blob) => blob ? resolve(blob) : reject(new Error('The image could not be encoded.')),
      outputMime(format),
      format === 'png' ? undefined : 0.92,
    )
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
  const nextRevisionRef = useRef(0)
  const contentRevisionRef = useRef(0)
  const savedRevisionRef = useRef(0)

  const [image, setImage] = useState<OpenImage | null>(null)
  const [dimensions, setDimensions] = useState<Dimensions>({ width: 0, height: 0 })
  const [dirty, setDirtyState] = useState(false)
  const [saving, setSaving] = useState(false)
  const [convertingPdf, setConvertingPdf] = useState(false)
  const [printing, setPrinting] = useState(false)
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

  imageRef.current = image
  dirtyRef.current = dirty

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
    setCropRect(null)
    reflectDirty()
  }, [context2d, reflectDirty])

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

  const loadPayloadNow = useCallback(async (payload: ImagePayload) => {
    const blob = new Blob([payload.data as BlobPart], { type: payload.mime })
    const url = URL.createObjectURL(blob)
    try {
      const decoded = new Image()
      decoded.decoding = 'async'
      decoded.src = url
      await decoded.decode()
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
      const hasAlpha = canvasHasTransparency(canvas)
      const saveFormat = defaultSaveFormat(payload.format)
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
    if (!dirtyRef.current) {
      void action()
      return
    }
    setConfirmState({ title, detail, action })
  }, [])

  const loadPayload = useCallback((payload: ImagePayload) => {
    requestProceed('Save changes before opening another image?', 'Your current edits have not been saved.', async () => {
      try { await loadPayloadNow(payload) } catch (error) { notify(error instanceof Error ? error.message : 'The image could not be opened.', 'error') }
    })
  }, [loadPayloadNow, notify, requestProceed])

  const openPath = useCallback(async (filePath: string) => {
    try { loadPayload(await window.simpleImage.openPath(filePath)) }
    catch (error) { notify(error instanceof Error ? error.message : 'The image could not be opened.', 'error') }
  }, [loadPayload, notify])

  const openDialog = useCallback(async () => {
    try {
      const payload = await window.simpleImage.openFile()
      if (payload) loadPayload(payload)
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be opened.', 'error')
    }
  }, [loadPayload, notify])

  const openDroppedFile = useCallback(async (file: File) => {
    try {
      const extension = file.name.split('.').pop()?.toLowerCase() || ''
      if (!SUPPORTED.includes(extension)) throw new Error(`.${extension || '?'} images are not supported.`)
      const filePath = window.simpleImage.pathForFile(file)
      const payload = filePath
        ? await window.simpleImage.openPath(filePath)
        : await window.simpleImage.openBytes(file.name, await file.arrayBuffer())
      loadPayload(payload)
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The dropped image could not be opened.', 'error')
    }
  }, [loadPayload, notify])

  const saveImage = useCallback(async (forceDialog = false): Promise<boolean> => {
    const canvas = canvasRef.current
    const current = imageRef.current
    if (!canvas || !current || saving || convertingPdf || printing) return false
    const documentId = documentIdRef.current
    const revision = contentRevisionRef.current
    const saveFormat = current.saveFormat
    setSaving(true)
    setSaveMenu(false)
    try {
      const blob = await canvasToBlob(canvas, saveFormat)
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
  }, [convertingPdf, notify, printing, reflectDirty, saving])

  const printImage = useCallback(async () => {
    const canvas = canvasRef.current
    if (!canvas || !image || printing || convertingPdf || saving) return
    setPrinting(true)
    try {
      const blob = await canvasToBlob(canvas, 'png')
      const printed = await window.simpleImage.printImage({ data: new Uint8Array(await blob.arrayBuffer()), name: image.name })
      if (!printed) notify('Printing was canceled.')
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be printed.', 'error')
    } finally {
      setPrinting(false)
    }
  }, [convertingPdf, image, notify, printing, saving])

  const convertToPdf = useCallback(async () => {
    const canvas = canvasRef.current
    const current = imageRef.current
    if (!canvas || !current || convertingPdf || saving || printing) return
    setConvertingPdf(true)
    setSaveMenu(false)
    try {
      const blob = await canvasToBlob(canvas, 'png')
      const result = await window.simpleImage.convertToPdf({
        data: new Uint8Array(await blob.arrayBuffer()),
        name: current.name,
      })
      if (result) notify(`Created ${result.name}`)
      else notify('PDF conversion was canceled.')
    } catch (error) {
      notify(error instanceof Error ? error.message : 'The image could not be converted to PDF.', 'error')
    } finally {
      setConvertingPdf(false)
    }
  }, [convertingPdf, notify, printing, saving])

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
    setDimensions({ width: canvas.width, height: canvas.height })
    setCropRect(null)
    markEdited()
    updateAlphaMetadata()
    if (fitMode) requestAnimationFrame(() => fitImage(canvas.width, canvas.height))
  }, [context2d, fitImage, fitMode, image, markEdited, pushUndo, updateAlphaMetadata])

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

  useEffect(() => window.simpleImage.onMaximized(setIsMaximized), [])
  useEffect(() => window.simpleImage.onOpenExternal(openPath), [openPath])
  useEffect(() => window.simpleImage.onCloseRequested(() => {
    requestProceed('Save changes before closing?', 'Your edits will be lost if you close without saving.', () => window.simpleImage.confirmClose())
  }), [requestProceed])

  useEffect(() => {
    if (!fitMode || !image) return
    const observer = new ResizeObserver(() => fitImage())
    if (viewportRef.current) observer.observe(viewportRef.current)
    return () => observer.disconnect()
  }, [fitImage, fitMode, image])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (confirmState) {
        if (event.key === 'Escape' && !saving) setConfirmState(null)
        if (event.key === 'Tab' || event.key === 'Enter' || event.key === ' ') return
        event.preventDefault()
        return
      }
      if (saving) {
        if (event.key === 'Tab') return
        event.preventDefault()
        return
      }
      const control = event.ctrlKey || event.metaKey
      if (control && event.key.toLowerCase() === 'o') { event.preventDefault(); void openDialog() }
      if (control && event.key.toLowerCase() === 's') { event.preventDefault(); void saveImage(event.shiftKey) }
      if (control && event.key.toLowerCase() === 'p') { event.preventDefault(); void printImage() }
      if (control && !event.shiftKey && event.key.toLowerCase() === 'z') { event.preventDefault(); undo() }
      if (control && (event.key.toLowerCase() === 'y' || (event.shiftKey && event.key.toLowerCase() === 'z'))) { event.preventDefault(); redo() }
      if (control && (event.key === '+' || event.key === '=')) { event.preventDefault(); changeZoom(zoom * 1.2) }
      if (control && event.key === '-') { event.preventDefault(); changeZoom(zoom / 1.2) }
      if (control && event.key === '0') { event.preventDefault(); fitImage() }
      if (event.key === 'Escape' && tool === 'crop') { setCropRect(null); setTool('view') }
      if (event.key === 'Enter' && tool === 'crop' && cropRect) applyCrop()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [applyCrop, changeZoom, confirmState, cropRect, fitImage, openDialog, printImage, redo, saveImage, saving, tool, undo, zoom])

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

  return (
    <main
      className="app-shell"
      onDragEnter={(event) => { event.preventDefault(); setDraggingFile(true) }}
      onDragOver={(event) => event.preventDefault()}
      onDrop={handleDrop}
    >
      <header className="titlebar">
        <div className="title-identity drag-region">
          <img className="brand-mark" src="./brand-icon.png" alt="" />
          <span className="app-name">simple</span>
          <span className="title-divider" />
          <span className="document-title">{image?.name || 'Image'}</span>
          {image && <span className={`save-state ${dirty ? 'is-dirty' : ''}`}>{saving ? 'Saving…' : dirty ? 'Modified' : 'Saved'}</span>}
        </div>
        <div className="title-actions">
          <button title="Open image (Ctrl+O)" onClick={() => void openDialog()}><FolderOpen /><span>Open</span></button>
          <div className="save-group">
            <button className="primary-button" disabled={!image || saving || convertingPdf || printing} title="Save (Ctrl+S)" onClick={() => void saveImage(false)}><Save /><span>{saving ? 'Saving…' : 'Save'}</span></button>
            <button className="primary-button split-button" disabled={!image || saving || convertingPdf || printing} title="Save options" onClick={() => setSaveMenu((value) => !value)}><ChevronDown /></button>
            {saveMenu && (
              <div className="save-menu">
                <button onClick={() => void saveImage(true)}><Save /><span><strong>Save as…</strong><small>Choose another file</small></span></button>
                <div className="menu-rule" />
                {(['png', 'jpeg', 'webp'] as SaveFormat[]).map((format) => (
                  <button key={format} onClick={() => { setImage((value) => value ? { ...value, saveFormat: format } : value); setSaveMenu(false) }}>
                    {image?.saveFormat === format ? <Check /> : <span className="check-placeholder" />}
                    <span><strong>{format === 'jpeg' ? 'JPEG' : format.toUpperCase()}</strong><small>{format === 'png' ? 'Lossless, supports transparency' : format === 'jpeg' ? 'Smaller photos, no transparency' : 'Efficient with transparency'}</small></span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <button disabled={!image || convertingPdf || printing || saving} title="Convert image to PDF" onClick={() => void convertToPdf()}><FileDown /><span>{convertingPdf ? 'Converting…' : 'Convert to PDF'}</span></button>
          <button disabled={!image || printing || convertingPdf || saving} title="Print (Ctrl+P)" onClick={() => void printImage()}><Printer /><span>{printing ? 'Printing…' : 'Print'}</span></button>
        </div>
        <div className="window-actions">
          <button aria-label="Minimize" onClick={() => window.simpleImage.minimize()}><Minus /></button>
          <button aria-label={isMaximized ? 'Restore' : 'Maximize'} onClick={() => window.simpleImage.toggleMaximize()}><Maximize2 /></button>
          <button className="close-button" aria-label="Close" onClick={() => requestProceed('Save changes before closing?', 'Your edits will be lost if you close without saving.', () => window.simpleImage.confirmClose())}><X /></button>
        </div>
      </header>

      <section className="toolbar" aria-label="Image tools">
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
        </div>
        {tool === 'crop' && cropRect && (
          <div className="crop-actions">
            <span>{Math.round(cropRect.width)} × {Math.round(cropRect.height)} px</span>
            <button onClick={() => { setCropRect(null); setTool('view') }}>Cancel</button>
            <button className="apply-button" onClick={applyCrop}><Check /> Apply crop</button>
          </div>
        )}
      </section>

      <section className="content">
        <div
          ref={viewportRef}
          className={`viewport ${image ? '' : 'is-empty'}`}
          onWheel={(event) => {
            if (!event.ctrlKey || !image) return
            event.preventDefault()
            changeZoom(zoom * (event.deltaY < 0 ? 1.12 : 1 / 1.12))
          }}
        >
          {!image && (
            <div className="welcome">
              <img src="./brand-icon.png" alt="" />
              <h1>See every detail.</h1>
              <p>Open, inspect, crop, rotate, paint, convert to PDF, print, and save images without uploading them anywhere.</p>
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
          <aside className="inspector">
            <div className="inspector-heading"><FileImage /><div><strong>Image details</strong><span>Read from this asset</span></div></div>
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

      <footer className="statusbar">
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
      {toast && <div className="toast" data-tone={toast.tone}>{toast.message}</div>}
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
