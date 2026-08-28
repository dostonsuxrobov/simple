import { useCallback, useEffect, useRef, useState } from 'react'
import type { DragEvent, PointerEvent as ReactPointerEvent } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { AlertTriangle, FileWarning, FolderOpen, RotateCcw } from 'lucide-react'
import { EditInspector } from './components/EditInspector'
import { ContinuousPdfViewer } from './components/ContinuousPdfViewer'
import { PrintDialog, type PrintDialogSubmission } from './components/PrintDialog'
import { SignaturePanel, type SignatureImage } from './components/SignaturePanel'
import { Sidebar } from './components/Sidebar'
import { StatusBar } from './components/StatusBar'
import { TitleBar } from './components/TitleBar'
import { Toolbar } from './components/Toolbar'
import { BusyOverlay, Button, Toast } from './components/ui'
import { Welcome } from './components/Welcome'
import { getPdfOutlineBookmarks, loadPdf, makeId, normalizeBytes, viewportRectToPdf } from './lib/pdf'
import {
  copyEditSelection as copyPageEditSelection,
  pasteEditToPage,
  type EditClipboardPayload,
} from './lib/editClipboard'
import { pageIndicesForTransfer } from './lib/pageTransfer'
import { clamp, errorMessage, isTypingTarget, withoutExtension } from './lib/utils'
import type {
  ActiveSearchMatch,
  Bookmark,
  DisplayRotation,
  DocumentPayload,
  OpenDocument,
  PageObjectEdit,
  PageTextEdit,
  PdfOverlay,
  PdfRect,
  RecentFile,
  ToolMode,
} from './types'

interface Snapshot {
  bytes: Uint8Array
  overlays: PdfOverlay[]
  formValues: Record<string, string | boolean>
  bookmarks: Bookmark[]
  bookmarksDirty: boolean
  pageRotations: Record<number, number>
  pageIndex: number
  dirty: boolean
}

const RECENT_KEY = 'folio:recent-files:v1'
const MAX_HISTORY_BYTES = 150 * 1024 * 1024

function fittedZoom(viewer: HTMLElement, viewport: { width: number; height: number }) {
  const availableWidth = Math.max(120, viewer.clientWidth - 104)
  const availableHeight = Math.max(120, viewer.clientHeight - 52)
  return clamp(Math.min(availableWidth / viewport.width, availableHeight / viewport.height), 0.35, 4)
}

function boundHistoryStack(next: Snapshot[]) {
  const uniqueDocuments = new Set<Uint8Array>()
  let total = 0
  for (const item of next) {
    if (uniqueDocuments.has(item.bytes)) continue
    uniqueDocuments.add(item.bytes)
    total += item.bytes.byteLength
  }
  while (next.length > 1 && total > MAX_HISTORY_BYTES) {
    const removed = next.shift()
    if (removed && !next.some((item) => item.bytes === removed.bytes)) total -= removed.bytes.byteLength
  }
  return next.slice(-20)
}

function viewerHasActiveTextSelection(viewer: HTMLElement) {
  if (viewer.dataset.pointerSelectingText === 'true') return true
  const selection = window.getSelection()
  if (!selection || selection.isCollapsed || selection.rangeCount < 1) return false
  const anchor = selection.anchorNode
  const focus = selection.focusNode
  return Boolean((anchor && viewer.contains(anchor)) || (focus && viewer.contains(focus)))
}

function bookmarkStorageKey(filePath: string | null, fingerprint: string) {
  const identity = filePath ? `path:${filePath.toLowerCase()}` : `fingerprint:${fingerprint}`
  return `folio:bookmarks:${encodeURIComponent(identity)}`
}

function readRecentFiles(): RecentFile[] {
  try {
    const value = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]')
    return Array.isArray(value) ? value.slice(0, 8) : []
  } catch {
    return []
  }
}

function remapAfterDelete(index: number, deleted: number[]): number | null {
  if (deleted.includes(index)) return null
  return index - deleted.filter((deletedIndex) => deletedIndex < index).length
}

function cloneOverlay(overlay: PdfOverlay): PdfOverlay {
  if (overlay.type === 'ink') {
    return { ...overlay, points: overlay.points.map((point) => ({ ...point })) }
  }
  return {
    ...overlay,
    rect: { ...overlay.rect },
    ...('originalRect' in overlay && overlay.originalRect ? { originalRect: { ...overlay.originalRect } } : {}),
  } as PdfOverlay
}

function upsertTextEdit(current: PdfOverlay[], edit: PageTextEdit): PdfOverlay[] {
  if (!edit.modified) return current
  const sourceStart = Number(edit.sourceSelectionStart)
  const sourceEnd = Number(edit.sourceSelectionEnd)
  const sourceItemText = edit.sourceItemText
  const sourceItemRect = edit.sourceItemRect
  const canReflowSourceItem = !edit.overlayId
    && edit.cover
    && typeof sourceItemText === 'string'
    && sourceItemRect
    && Number.isInteger(sourceStart)
    && Number.isInteger(sourceEnd)
    && sourceStart >= 0
    && sourceEnd >= sourceStart
    && sourceEnd <= sourceItemText.length
    && sourceItemText.slice(sourceStart, sourceEnd) === edit.originalText
  const selectedSourceRect = edit.originalRect || edit.rect
  const deltaX = edit.rect.x - selectedSourceRect.x
  const deltaY = edit.rect.y - selectedSourceRect.y
  const committedRect = canReflowSourceItem
    ? {
        x: sourceItemRect.x + deltaX,
        y: sourceItemRect.y + deltaY,
        width: Math.max(1, sourceItemRect.width + edit.rect.width - selectedSourceRect.width),
        height: Math.max(1, sourceItemRect.height + edit.rect.height - selectedSourceRect.height),
      }
    : edit.rect
  const committedBaselineOffset = canReflowSourceItem && Number.isFinite(edit.baselineOffset)
    ? selectedSourceRect.y + Number(edit.baselineOffset) - sourceItemRect.y
    : edit.baselineOffset
  const overlay: PdfOverlay = {
    id: edit.overlayId || makeId('text'),
    type: 'text',
    pageIndex: edit.pageIndex,
    rect: { ...committedRect },
    originalRect: canReflowSourceItem
      ? { ...sourceItemRect }
      : edit.originalRect ? { ...edit.originalRect } : undefined,
    originalText: canReflowSourceItem ? sourceItemText : edit.originalText,
    text: canReflowSourceItem
      ? `${sourceItemText.slice(0, sourceStart)}${edit.text}${sourceItemText.slice(sourceEnd)}`
      : edit.text,
    fontSize: edit.fontSize,
    fontFamily: edit.fontFamily,
    fontWeight: edit.fontWeight,
    fontStyle: edit.fontStyle,
    lineHeight: edit.lineHeight,
    letterSpacing: edit.letterSpacing,
    scaleX: edit.scaleX,
    angle: edit.angle,
    direction: edit.direction,
    fontKey: edit.fontKey,
    fontData: edit.fontData,
    baselineOffset: committedBaselineOffset,
    sourceSpaceWidth: edit.sourceSpaceWidth,
    displayRotation: edit.displayRotation,
    align: edit.align,
    color: edit.color,
    backgroundColor: edit.backgroundColor,
    cover: edit.cover,
  }
  if (edit.overlayId) return current.map((item) => item.id === edit.overlayId ? overlay : item)
  return [...current, overlay]
}

function upsertObjectEdit(current: PdfOverlay[], edit: PageObjectEdit): PdfOverlay[] {
  if (!edit.modified) return current
  const overlay: PdfOverlay = {
    id: edit.overlayId || makeId('object'),
    type: 'object',
    pageIndex: edit.pageIndex,
    kind: edit.kind,
    rect: { ...edit.rect },
    originalRect: edit.originalRect ? { ...edit.originalRect } : undefined,
    dataUrl: edit.dataUrl,
    opacity: edit.opacity,
    cover: edit.cover,
    displayRotation: edit.displayRotation,
  }
  if (edit.overlayId) return current.map((item) => item.id === edit.overlayId ? overlay : item)
  return [...current, overlay]
}

async function imageDimensions(dataUrl: string) {
  return new Promise<{ width: number; height: number }>((resolve, reject) => {
    const image = new Image()
    image.onload = () => resolve({ width: image.naturalWidth || image.width, height: image.naturalHeight || image.height })
    image.onerror = () => reject(new Error('The image could not be decoded.'))
    image.src = dataUrl
  })
}

async function rotateImageData(dataUrl: string, direction: 'left' | 'right') {
  const source = new Image()
  await new Promise<void>((resolve, reject) => {
    source.onload = () => resolve()
    source.onerror = () => reject(new Error('The image could not be rotated.'))
    source.src = dataUrl
  })
  const canvas = document.createElement('canvas')
  canvas.width = source.naturalHeight || source.height
  canvas.height = source.naturalWidth || source.width
  const context = canvas.getContext('2d')
  if (!context) throw new Error('The image could not be rotated.')
  context.translate(canvas.width / 2, canvas.height / 2)
  context.rotate(direction === 'right' ? Math.PI / 2 : -Math.PI / 2)
  context.drawImage(source, -source.naturalWidth / 2, -source.naturalHeight / 2)
  return canvas.toDataURL('image/png')
}

/**
 * Rotate an upright (as-displayed) image by -displayRotation so the stored
 * pixels sit in unrotated PDF orientation — the same convention captureRect
 * uses — keeping the saver's axis-aligned draw correct on rotated pages.
 */
async function normalizeImageOrientation(dataUrl: string, displayRotation: number) {
  const rotation = ((displayRotation % 360) + 360) % 360
  if (!rotation) return dataUrl
  const source = new Image()
  await new Promise<void>((resolve, reject) => {
    source.onload = () => resolve()
    source.onerror = () => reject(new Error('The image could not be rotated.'))
    source.src = dataUrl
  })
  const width = source.naturalWidth || source.width
  const height = source.naturalHeight || source.height
  const canvas = document.createElement('canvas')
  canvas.width = rotation % 180 ? height : width
  canvas.height = rotation % 180 ? width : height
  const context = canvas.getContext('2d')
  if (!context) throw new Error('The image could not be rotated.')
  context.translate(canvas.width / 2, canvas.height / 2)
  context.rotate(-rotation * Math.PI / 180)
  context.drawImage(source, -width / 2, -height / 2, width, height)
  return canvas.toDataURL('image/png')
}

export default function App() {
  const [documentFile, setDocumentFile] = useState<OpenDocument | null>(null)
  const [bytes, setBytes] = useState<Uint8Array | null>(null)
  const [pdf, setPdf] = useState<PDFDocumentProxy | null>(null)
  const [documentLoading, setDocumentLoading] = useState(false)
  const [documentError, setDocumentError] = useState('')
  const [passwordProtected, setPasswordProtected] = useState(false)
  const [dirty, setDirty] = useState(false)
  const [busy, setBusy] = useState('')
  const [toast, setToast] = useState<{ message: string; action?: string; onAction?: () => void } | null>(null)
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [searchRequestId, setSearchRequestId] = useState(0)
  const [activeSearchMatch, setActiveSearchMatch] = useState<ActiveSearchMatch | null>(null)
  const [pageIndex, setPageIndex] = useState(0)
  const [selectedPages, setSelectedPages] = useState<Set<number>>(new Set([0]))
  const [zoom, setZoom] = useState(1)
  const [zoomMode, setZoomMode] = useState<'fit' | 'custom'>('fit')
  const [tool, setTool] = useState<ToolMode>('select')
  const [overlays, setOverlays] = useState<PdfOverlay[]>([])
  const [formValues, setFormValues] = useState<Record<string, string | boolean>>({})
  const [bookmarks, setBookmarks] = useState<Bookmark[]>([])
  const [bookmarksDirty, setBookmarksDirty] = useState(false)
  const [pageRotations, setPageRotations] = useState<Record<number, number>>({})
  const [bookmarkKey, setBookmarkKey] = useState('')
  const [textEdit, setTextEdit] = useState<PageTextEdit | null>(null)
  const [objectEdit, setObjectEdit] = useState<PageObjectEdit | null>(null)
  const [selectingObjectRegion, setSelectingObjectRegion] = useState(false)
  const [pendingSignature, setPendingSignature] = useState<SignatureImage | null>(null)
  const [undoStack, setUndoStack] = useState<Snapshot[]>([])
  const [redoStack, setRedoStack] = useState<Snapshot[]>([])
  const [recentFiles, setRecentFiles] = useState<RecentFile[]>(readRecentFiles)
  const [isDragging, setIsDragging] = useState(false)
  const [isPanning, setIsPanning] = useState(false)
  const [printDialogOpen, setPrintDialogOpen] = useState(false)
  const [immersive, setImmersive] = useState(false)
  const dragDepth = useRef(0)
  const viewerRef = useRef<HTMLElement>(null)
  const panRef = useRef<{ x: number; y: number; scrollLeft: number; scrollTop: number } | null>(null)
  const editClipboardRef = useRef<EditClipboardPayload | null>(null)
  const [hasEditClipboard, setHasEditClipboard] = useState(false)
  const fitRequestGenerationRef = useRef(0)
  const lastSelectedPage = useRef(0)
  const undoStackRef = useRef<Snapshot[]>([])
  const redoStackRef = useRef<Snapshot[]>([])
  const dragExportBytesRef = useRef<Uint8Array | null>(null)
  const liveStateRef = useRef({ bytes, overlays, formValues, bookmarks, bookmarksDirty, pageRotations, pageIndex, dirty })

  useEffect(() => {
    liveStateRef.current = { bytes, overlays, formValues, bookmarks, bookmarksDirty, pageRotations, pageIndex, dirty }
    undoStackRef.current = undoStack
    redoStackRef.current = redoStack
  }, [bytes, overlays, formValues, bookmarks, bookmarksDirty, pageRotations, pageIndex, dirty, undoStack, redoStack])

  useEffect(() => {
    dragExportBytesRef.current = null
  }, [bytes, overlays, formValues, bookmarks, bookmarksDirty, pageRotations, textEdit, objectEdit])

  const showToast = useCallback((message: string, action?: string, onAction?: () => void) => {
    setToast({ message, action, onAction })
  }, [])

  useEffect(() => {
    if (!toast) return
    const timeout = window.setTimeout(() => setToast(null), 4200)
    return () => window.clearTimeout(timeout)
  }, [toast])

  useEffect(() => {
    if (!bytes) {
      setPdf(null)
      return
    }
    let cancelled = false
    let loadedPdf: PDFDocumentProxy | null = null
    setDocumentLoading(true)
    setDocumentError('')
    setPdf(null)
    loadPdf(bytes, (updatePassword) => {
      setPasswordProtected(true)
      const password = window.prompt('This PDF is password protected. Enter its password to read it:')
      updatePassword(password || '')
    }).then(async (nextPdf) => {
      loadedPdf = nextPdf
      if (cancelled) {
        nextPdf.destroy()
        return
      }
      const nextPageIndex = clamp(pageIndex, 0, Math.max(0, nextPdf.numPages - 1))
      // Establish the final fit scale before mounting the first page. Otherwise
      // the initial 100% text layer can briefly appear and then be replaced by
      // the fitted layer, moving the glyphs underneath an early pointer drag.
      if (zoomMode === 'fit' && viewerRef.current) {
        const currentPage = await nextPdf.getPage(nextPageIndex + 1)
        if (cancelled) return
        const pageRotation = (((currentPage.rotate || 0) + (pageRotations[nextPageIndex] || 0)) % 360 + 360) % 360
        const viewport = currentPage.getViewport({ scale: 1, rotation: pageRotation })
        const nextZoom = fittedZoom(viewerRef.current, viewport)
        setZoom((current) => Math.abs(current - nextZoom) < 0.002 ? current : nextZoom)
      }
      setPdf(nextPdf)
      setPageIndex(nextPageIndex)
      setSelectedPages((current) => {
        const next = new Set([...current].filter((index) => index < nextPdf.numPages))
        return next.size ? next : new Set([0])
      })
    }).catch((error) => {
      if (!cancelled) setDocumentError(errorMessage(error))
    }).finally(() => {
      if (!cancelled) setDocumentLoading(false)
    })
    return () => {
      cancelled = true
      loadedPdf?.destroy()
    }
  }, [bytes])

  useEffect(() => {
    if (!pdf || bookmarkKey) return
    let cancelled = false
    const fingerprint = pdf.fingerprints?.[0] || `${documentFile?.name || 'document'}:${bytes?.byteLength || 0}`
    const key = bookmarkStorageKey(documentFile?.path || null, fingerprint)
    getPdfOutlineBookmarks(pdf).then((outline) => {
      if (cancelled) return
      setBookmarkKey(key)
      const nativeBookmarks: Bookmark[] = outline.flatMap((item) => item.pageIndex === null ? [] : [{
        id: item.id,
        pageIndex: item.pageIndex,
        label: item.title,
        depth: item.depth,
        source: 'document' as const,
        expanded: item.expanded,
        bold: item.bold,
        italic: item.italic,
        color: item.color,
        url: item.url,
      }])
      if (nativeBookmarks.length) {
        setBookmarks(nativeBookmarks)
        setBookmarksDirty(false)
        return
      }
      try {
        const stored = JSON.parse(localStorage.getItem(key) || '[]')
        const legacy = Array.isArray(stored)
          ? stored.filter((item) => Number.isInteger(item.pageIndex) && item.pageIndex < pdf.numPages)
            .map((item) => ({ ...item, source: 'simple' as const }))
          : []
        if (legacy.length) localStorage.removeItem(key)
        setBookmarks(legacy)
        // Migrate legacy app-only bookmarks into the PDF on the next save.
        setBookmarksDirty(Boolean(legacy.length))
      } catch {
        setBookmarks([])
        setBookmarksDirty(false)
      }
    }).catch(() => {
      if (!cancelled) {
        setBookmarkKey(key)
        setBookmarks([])
      }
    })
    return () => { cancelled = true }
  }, [pdf, bookmarkKey, documentFile, bytes])

  function rememberFile(file: OpenDocument) {
    if (!file.path) return
    setRecentFiles((current) => {
      const next = [
        { path: file.path!, name: file.name, openedAt: Date.now() },
        ...current.filter((item) => item.path.toLocaleLowerCase() !== file.path!.toLocaleLowerCase()),
      ].slice(0, 8)
      localStorage.setItem(RECENT_KEY, JSON.stringify(next))
      return next
    })
  }

  function currentSnapshot(): Snapshot | null {
    const live = liveStateRef.current
    if (!live.bytes) return null
    return {
      // PDF byte arrays are immutable in Simple. Keep the reference so a small
      // annotation edit does not clone an entire large document for Undo.
      bytes: live.bytes,
      overlays: live.overlays.map(cloneOverlay),
      formValues: { ...live.formValues },
      bookmarks: live.bookmarks.map((bookmark) => ({ ...bookmark })),
      bookmarksDirty: live.bookmarksDirty,
      pageRotations: { ...live.pageRotations },
      pageIndex: live.pageIndex,
      dirty: live.dirty,
    }
  }

  function addUndoSnapshot(snapshot: Snapshot) {
    setUndoStack((current) => {
      const bounded = boundHistoryStack([...current, snapshot])
      undoStackRef.current = bounded
      return bounded
    })
    redoStackRef.current = []
    setRedoStack([])
  }

  function restoreSnapshot(snapshot: Snapshot) {
    setBytes(snapshot.bytes)
    setOverlays(snapshot.overlays.map(cloneOverlay))
    setFormValues({ ...snapshot.formValues })
    setBookmarks(snapshot.bookmarks.map((bookmark) => ({ ...bookmark })))
    setBookmarksDirty(snapshot.bookmarksDirty)
    setPageRotations({ ...snapshot.pageRotations })
    setPageIndex(snapshot.pageIndex)
    setSelectedPages(new Set([snapshot.pageIndex]))
    setDirty(snapshot.dirty)
    setTextEdit(null)
    setObjectEdit(null)
    setSelectingObjectRegion(false)
  }

  function undo() {
    const previous = undoStackRef.current.at(-1)
    const current = currentSnapshot()
    if (!previous || !current) return
    const nextUndo = undoStackRef.current.slice(0, -1)
    const nextRedo = boundHistoryStack([...redoStackRef.current, current])
    undoStackRef.current = nextUndo
    redoStackRef.current = nextRedo
    setUndoStack(nextUndo)
    setRedoStack(nextRedo)
    restoreSnapshot(previous)
    showToast('Undid the last change')
  }

  function redo() {
    const next = redoStackRef.current.at(-1)
    const current = currentSnapshot()
    if (!next || !current) return
    const nextRedo = redoStackRef.current.slice(0, -1)
    const nextUndo = [...undoStackRef.current, current].slice(-20)
    redoStackRef.current = nextRedo
    undoStackRef.current = nextUndo
    setRedoStack(nextRedo)
    setUndoStack(nextUndo)
    restoreSnapshot(next)
    showToast('Redid the last change')
  }

  async function openPayload(payload: DocumentPayload) {
    const file: OpenDocument = {
      bytes: normalizeBytes(payload.data),
      name: payload.name,
      path: payload.path,
      sourcePath: payload.sourcePath,
      converted: payload.converted,
      signatureDetected: payload.signatureDetected,
    }
    setDocumentFile(file)
    setBytes(file.bytes)
    setPasswordProtected(false)
    setDirty(file.converted)
    setPageIndex(0)
    setActiveSearchMatch(null)
    setZoomMode('fit')
    setSelectedPages(new Set([0]))
    setOverlays([])
    setFormValues({})
    setBookmarks([])
    setBookmarksDirty(false)
    setPageRotations({})
    setBookmarkKey('')
    setUndoStack([])
    setRedoStack([])
    setTextEdit(null)
    setObjectEdit(null)
    setSelectingObjectRegion(false)
    setPendingSignature(null)
    setTool('select')
    rememberFile(file)
    if (file.converted) showToast('Converted to PDF — save to choose where to keep it')
    else if (file.signatureDetected) showToast('This PDF is digitally signed. Editing can invalidate its signature.')
  }

  function canReplaceCurrent() {
    return !(dirty || textEdit?.modified || objectEdit?.modified) || window.confirm('This document has unsaved changes. Discard them and open another file?')
  }

  async function openFile() {
    setBusy(documentFile ? 'Opening in a new window…' : 'Opening document…')
    try {
      if (documentFile) {
        await window.simple.openInNewWindow()
        return
      }
      const payload = await window.simple.openFile()
      if (payload) await openPayload(payload)
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  async function openPath(filePath: string) {
    setBusy(documentFile ? 'Opening in a new window…' : 'Opening document…')
    try {
      if (documentFile) {
        await window.simple.openInNewWindow(filePath)
        return
      }
      await openPayload(await window.simple.openPath(filePath))
    } catch (error) {
      setRecentFiles((current) => {
        const next = current.filter((item) => item.path !== filePath)
        localStorage.setItem(RECENT_KEY, JSON.stringify(next))
        return next
      })
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  const openPathRef = useRef(openPath)
  openPathRef.current = openPath

  useEffect(() => window.simple.onOpenExternal((filePath) => { openPathRef.current(filePath) }), [])

  async function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault()
    dragDepth.current = 0
    setIsDragging(false)
    if (!canReplaceCurrent()) return
    const files = Array.from(event.dataTransfer.files)
    const file = files[0]
    if (!file) return
    setBusy('Opening document…')
    try {
      const payload = await window.simple.openBytes(file.name, await file.arrayBuffer())
      await openPayload(payload)
      if (files.length > 1) showToast(`Opened ${file.name} — ${files.length - 1} more ${files.length === 2 ? 'file' : 'files'} ignored`)
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  async function preparedBytes() {
    if (!bytes) throw new Error('No PDF is open.')
    let outputOverlays = overlays
    if (textEdit?.modified) outputOverlays = upsertTextEdit(outputOverlays, textEdit)
    if (objectEdit?.modified) outputOverlays = upsertObjectEdit(outputOverlays, objectEdit)
    const hasRotations = Object.values(pageRotations).some((amount) => amount % 360 !== 0)
    if (!outputOverlays.length && !Object.keys(formValues).length && !hasRotations && !bookmarksDirty) return bytes
    return normalizeBytes(await window.simple.flattenOverlays(bytes, outputOverlays, formValues, {
      ...(hasRotations ? { pageRotations } : {}),
      ...(bookmarksDirty ? { bookmarks } : {}),
    }))
  }

  async function save(forceDialog = false) {
    if (!documentFile || !bytes) return
    if (passwordProtected && dirty) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    if (documentFile.signatureDetected && dirty && !window.confirm('Saving edits can invalidate this PDF’s digital signature. Continue with Save As?')) return
    setBusy('Saving PDF…')
    try {
      let retainedOverlays = overlays
      if (textEdit?.modified) retainedOverlays = upsertTextEdit(retainedOverlays, textEdit)
      if (objectEdit?.modified) retainedOverlays = upsertObjectEdit(retainedOverlays, objectEdit)
      const outputBytes = await preparedBytes()
      const result = await window.simple.savePdf({
        data: outputBytes,
        path: documentFile.signatureDetected ? null : documentFile.path,
        name: documentFile.name,
        forceDialog: forceDialog || documentFile.signatureDetected,
      })
      if (!result) return
      // Keep the immutable editing base plus the logical overlays in memory.
      // Every later save is regenerated from that base, so a long edit session
      // never feeds already-flattened covers/replacements back into itself.
      const nextFile = { ...documentFile, bytes, path: result.path, name: result.name, converted: false, signatureDetected: false }
      const nextBookmarkKey = bookmarkStorageKey(result.path, pdf?.fingerprints?.[0] || `${result.name}:${outputBytes.byteLength}`)
      if (bookmarkKey && bookmarkKey !== nextBookmarkKey) localStorage.removeItem(bookmarkKey)
      setBookmarkKey(nextBookmarkKey)
      setDocumentFile(nextFile)
      setOverlays(retainedOverlays)
      setTextEdit(null)
      setObjectEdit(null)
      setSelectingObjectRegion(false)
      setDirty(false)
      setUndoStack([])
      setRedoStack([])
      rememberFile(nextFile)
      showToast('Saved successfully', 'Show file', () => window.simple.showItem(result.path))
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  async function mutate(operation: Record<string, unknown>, label: string, after?: () => void) {
    if (!bytes) return false
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return false
    }
    const snapshot = currentSnapshot()
    setBusy(label)
    try {
      const result = normalizeBytes(await window.simple.mutatePdf(bytes, operation))
      if (snapshot) addUndoSnapshot(snapshot)
      setBytes(result)
      setDirty(true)
      after?.()
      return true
    } catch (error) {
      showToast(errorMessage(error))
      return false
    } finally {
      setBusy('')
    }
  }

  function targetPages() {
    return selectedPages.size ? [...selectedPages].sort((a, b) => a - b) : [pageIndex]
  }

  function rotatePages(degrees = 90) {
    const indices = targetPages()
    // A live native selection freezes text-layer rebuilds and page pruning;
    // rotated canvases must never sit beneath the old orientation's spans.
    window.getSelection()?.removeAllRanges()
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setPageRotations((current) => {
      const next = { ...current }
      for (const index of indices) {
        const amount = (((next[index] || 0) + degrees) % 360 + 360) % 360
        if (amount) next[index] = amount
        else delete next[index]
      }
      return next
    })
    setDirty(true)
    showToast(`${indices.length === 1 ? 'Page' : `${indices.length} pages`} rotated`)
  }

  async function duplicatePages() {
    if (!pdf) return
    const indices = targetPages()
    let sourceOverlays = overlays
    if (textEdit?.modified) sourceOverlays = upsertTextEdit(sourceOverlays, textEdit)
    if (objectEdit?.modified) sourceOverlays = upsertObjectEdit(sourceOverlays, objectEdit)
    if (await mutate({ type: 'duplicate', indices }, `Duplicating ${indices.length === 1 ? 'page' : 'pages'}…`, () => {
      const before = (index: number) => indices.filter((selected) => selected < index).length
      const through = (index: number) => indices.filter((selected) => selected <= index).length
      const remapped = sourceOverlays.map((overlay) => ({ ...cloneOverlay(overlay), pageIndex: overlay.pageIndex + before(overlay.pageIndex) }))
      const copies = sourceOverlays
        .filter((overlay) => indices.includes(overlay.pageIndex))
        .map((overlay) => ({ ...cloneOverlay(overlay), id: makeId(overlay.type), pageIndex: overlay.pageIndex + through(overlay.pageIndex) }))
      setOverlays([...remapped, ...copies])
      setPageRotations((current) => {
        const next: Record<number, number> = {}
        for (const [rawIndex, amount] of Object.entries(current)) {
          const index = Number(rawIndex)
          next[index + before(index)] = amount
          if (indices.includes(index)) next[index + through(index)] = amount
        }
        return next
      })
      setBookmarks((current) => current.map((bookmark) => ({ ...bookmark, pageIndex: bookmark.pageIndex + before(bookmark.pageIndex) })))
      const duplicatedIndices = indices.map((index) => index + through(index))
      setPageIndex(duplicatedIndices[0])
      setSelectedPages(new Set(duplicatedIndices))
      setTextEdit(null)
      setObjectEdit(null)
      showToast(`${indices.length === 1 ? 'Page duplicated' : `${indices.length} pages duplicated`}`)
    })) return
  }

  async function insertBlankPage() {
    if (!pdf) return
    const currentPage = await pdf.getPage(pageIndex + 1)
    const baseViewport = currentPage.getViewport({ scale: 1 })
    const insertIndex = pageIndex + 1
    await mutate({ type: 'blank', index: insertIndex, width: baseViewport.width, height: baseViewport.height }, 'Adding blank page…', () => {
      setOverlays((current) => current.map((overlay) => ({ ...overlay, pageIndex: overlay.pageIndex >= insertIndex ? overlay.pageIndex + 1 : overlay.pageIndex })))
      setPageRotations((current) => Object.fromEntries(Object.entries(current).map(([rawIndex, amount]) => {
        const index = Number(rawIndex)
        return [index >= insertIndex ? index + 1 : index, amount]
      })))
      setBookmarks((current) => current.map((bookmark) => ({ ...bookmark, pageIndex: bookmark.pageIndex >= insertIndex ? bookmark.pageIndex + 1 : bookmark.pageIndex })))
      setPageIndex(insertIndex)
      setSelectedPages(new Set([insertIndex]))
      showToast('Blank page added')
    })
  }

  async function deletePages() {
    if (!pdf) return
    const indices = targetPages()
    if (indices.length >= pdf.numPages) {
      showToast('A PDF must keep at least one page.')
      return
    }
    if (!window.confirm(`Delete ${indices.length === 1 ? `page ${indices[0] + 1}` : `${indices.length} selected pages`}? You can undo this action.`)) return
    await mutate({ type: 'delete', indices }, `Deleting ${indices.length === 1 ? 'page' : 'pages'}…`, () => {
      setOverlays((current) => current.flatMap((overlay) => {
        const mapped = remapAfterDelete(overlay.pageIndex, indices)
        return mapped === null ? [] : [{ ...overlay, pageIndex: mapped }]
      }))
      setBookmarks((current) => current.flatMap((bookmark) => {
        const mapped = remapAfterDelete(bookmark.pageIndex, indices)
        return mapped === null ? [] : [{ ...bookmark, pageIndex: mapped }]
      }))
      setPageRotations((current) => Object.fromEntries(Object.entries(current).flatMap(([rawIndex, amount]) => {
        const mapped = remapAfterDelete(Number(rawIndex), indices)
        return mapped === null ? [] : [[mapped, amount]]
      })))
      const nextPage = clamp(pageIndex - indices.filter((index) => index < pageIndex).length, 0, pdf.numPages - indices.length - 1)
      setPageIndex(nextPage)
      setSelectedPages(new Set([nextPage]))
      showToast(`${indices.length === 1 ? 'Page' : `${indices.length} pages`} deleted`, 'Undo', undo)
    })
  }

  async function reorderPage(from: number, to: number) {
    if (!pdf || from === to) return
    const order = Array.from({ length: pdf.numPages }, (_, index) => index)
    const [moved] = order.splice(from, 1)
    order.splice(to, 0, moved)
    await mutate({ type: 'reorder', order }, 'Reordering pages…', () => {
      const mapIndex = (oldIndex: number) => order.indexOf(oldIndex)
      setOverlays((current) => current.map((overlay) => ({ ...overlay, pageIndex: mapIndex(overlay.pageIndex) })))
      setPageRotations((current) => Object.fromEntries(Object.entries(current).map(([rawIndex, amount]) => [mapIndex(Number(rawIndex)), amount])))
      setBookmarks((current) => current.map((bookmark) => ({ ...bookmark, pageIndex: mapIndex(bookmark.pageIndex) })))
      const nextIndex = mapIndex(from)
      setPageIndex(nextIndex)
      setSelectedPages(new Set([nextIndex]))
      showToast('Page moved')
    })
  }

  async function addPages() {
    if (!bytes || !pdf) return
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    const insertIndex = pageIndex + 1
    let sourceOverlays = overlays
    if (textEdit?.modified) sourceOverlays = upsertTextEdit(sourceOverlays, textEdit)
    if (objectEdit?.modified) sourceOverlays = upsertObjectEdit(sourceOverlays, objectEdit)
    const snapshot = currentSnapshot()
    if (snapshot && (textEdit?.modified || objectEdit?.modified)) {
      snapshot.overlays = sourceOverlays.map(cloneOverlay)
      snapshot.dirty = true
    }
    setBusy('Adding pages…')
    try {
      const result = await window.simple.insertFiles(bytes, insertIndex)
      if (!result) return
      if (snapshot) addUndoSnapshot(snapshot)
      setBytes(normalizeBytes(result.data))
      setOverlays(sourceOverlays.map((overlay) => ({ ...overlay, pageIndex: overlay.pageIndex >= insertIndex ? overlay.pageIndex + result.added : overlay.pageIndex })))
      setPageRotations((current) => Object.fromEntries(Object.entries(current).map(([rawIndex, amount]) => {
        const index = Number(rawIndex)
        return [index >= insertIndex ? index + result.added : index, amount]
      })))
      setBookmarks((current) => current.map((bookmark) => ({ ...bookmark, pageIndex: bookmark.pageIndex >= insertIndex ? bookmark.pageIndex + result.added : bookmark.pageIndex })))
      setPageIndex(insertIndex)
      const insertedPages = Array.from({ length: result.added }, (_, offset) => insertIndex + offset)
      setSelectedPages(new Set(insertedPages))
      lastSelectedPage.current = insertedPages.at(-1) ?? insertIndex
      setTextEdit(null)
      setObjectEdit(null)
      setDirty(true)
      showToast(`${result.added} ${result.added === 1 ? 'page' : 'pages'} added`)
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  async function importPagesAt(files: File[], insertIndex: number) {
    if (!bytes || !pdf || !files.length) return
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    let sourceOverlays = overlays
    if (textEdit?.modified) sourceOverlays = upsertTextEdit(sourceOverlays, textEdit)
    if (objectEdit?.modified) sourceOverlays = upsertObjectEdit(sourceOverlays, objectEdit)
    const snapshot = currentSnapshot()
    if (snapshot && (textEdit?.modified || objectEdit?.modified)) {
      snapshot.overlays = sourceOverlays.map(cloneOverlay)
      snapshot.dirty = true
    }
    setBusy('Adding dropped pages…')
    try {
      const inputs = await Promise.all(files.map(async (file) => ({ name: file.name, data: await file.arrayBuffer() })))
      const result = await window.simple.insertDroppedFiles(bytes, insertIndex, inputs)
      if (snapshot) addUndoSnapshot(snapshot)
      setBytes(normalizeBytes(result.data))
      setOverlays(sourceOverlays.map((overlay) => ({
        ...overlay,
        pageIndex: overlay.pageIndex >= insertIndex ? overlay.pageIndex + result.added : overlay.pageIndex,
      })))
      setPageRotations((current) => Object.fromEntries(Object.entries(current).map(([rawIndex, amount]) => {
        const index = Number(rawIndex)
        return [index >= insertIndex ? index + result.added : index, amount]
      })))
      setBookmarks((current) => current.map((bookmark) => ({
        ...bookmark,
        pageIndex: bookmark.pageIndex >= insertIndex ? bookmark.pageIndex + result.added : bookmark.pageIndex,
      })))
      setPageIndex(insertIndex)
      const insertedPages = Array.from({ length: result.added }, (_, offset) => insertIndex + offset)
      setSelectedPages(new Set(insertedPages))
      lastSelectedPage.current = insertedPages.at(-1) ?? insertIndex
      setTextEdit(null)
      setObjectEdit(null)
      setDirty(true)
      showToast(`${result.added} ${result.added === 1 ? 'page' : 'pages'} added where dropped`)
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  function importDroppedPages(files: File[], insertIndex: number) {
    if (!files.length) {
      showToast('Dropped files must be PDF, image, Word, or text documents.')
      return
    }
    void importPagesAt(files, insertIndex)
  }

  async function startPageDrag(index: number, event: DragEvent<HTMLElement>) {
    if (!bytes || !documentFile) return
    event.preventDefault()
    try {
      const indices = pageIndicesForTransfer(selectedPages, index)
      if (!selectedPages.has(index)) {
        setPageIndex(index)
        setSelectedPages(new Set([index]))
        lastSelectedPage.current = index
      }
      const output = dragExportBytesRef.current ?? await preparedBytes()
      dragExportBytesRef.current = output
      const suggested = `${withoutExtension(documentFile.name)} - ${indices.length === 1 ? `page ${indices[0] + 1}` : `${indices.length} pages`}.pdf`
      await window.simple.startPageDrag(output, indices, suggested)
    } catch (error) {
      showToast(errorMessage(error))
    }
  }

  async function exportPages() {
    if (!bytes || !documentFile) return
    const indices = targetPages()
    setBusy('Preparing export…')
    try {
      const output = await preparedBytes()
      const suggested = `${withoutExtension(documentFile.name)} — ${indices.length === 1 ? `page ${indices[0] + 1}` : `${indices.length} pages`}.pdf`
      const exportPath = await window.simple.exportPages(output, indices, suggested)
      if (exportPath) showToast('Pages exported', 'Show file', () => window.simple.showItem(exportPath))
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  async function printDocument() {
    if (!bytes || !documentFile) return
    setBusy('Preparing print…')
    try {
      const output = await preparedBytes()
      const printed = await window.simple.printPdf(output, documentFile.name)
      if (printed) showToast('Print preview opened')
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  async function runPrintJob(job: PrintDialogSubmission) {
    if (!bytes || !documentFile || !pdf) return
    setPrintDialogOpen(false)
    setBusy('Printing…')
    try {
      const output = await preparedBytes()
      let landscape: boolean
      if (job.orientation === 'auto') {
        const probeIndex = job.pageIndices?.[0] ?? 0
        const probePage = await pdf.getPage(probeIndex + 1)
        const rotation = (((probePage.rotate || 0) + (pageRotations[probeIndex] || 0)) % 360 + 360) % 360
        const viewport = probePage.getViewport({ scale: 1, rotation })
        landscape = viewport.width > viewport.height
      } else {
        landscape = job.orientation === 'landscape'
      }
      const result = await window.simple.printPdfDirect(output, documentFile.name, {
        deviceName: job.deviceName,
        copies: job.copies,
        ...(job.pageIndices ? { pageIndices: job.pageIndices } : {}),
        landscape,
        color: job.color,
        ...(job.duplexMode ? { duplexMode: job.duplexMode } : {}),
        collate: job.collate,
      })
      if (result.success) showToast(`Sent to ${job.printerLabel}`)
      else showToast(result.failureReason ? `Print failed — ${result.failureReason}` : 'Print failed', 'Open preview', () => { void printDocument() })
    } catch (error) {
      showToast(errorMessage(error), 'Open preview', () => { void printDocument() })
    } finally {
      setBusy('')
    }
  }

  async function cropPage(rect: PdfRect, targetPageIndex = pageIndex) {
    if (await mutate({ type: 'crop', pageIndex: targetPageIndex, rect }, 'Cropping page…')) {
      setTool('select')
      showToast('Crop applied — save to keep it')
    }
  }

  function addHighlights(rects: PdfRect[], targetPageIndex = pageIndex) {
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    const additions: PdfOverlay[] = rects.map((rect) => ({
      id: makeId('highlight'),
      type: 'highlight',
      pageIndex: targetPageIndex,
      rect,
      color: [1, 0.86, 0.16],
      opacity: 0.34,
    }))
    setOverlays((current) => [...current, ...additions])
    setDirty(true)
    showToast(`${additions.length === 1 ? 'Highlight' : 'Highlights'} added`)
  }

  function addTextMarkup(style: 'underline' | 'strikeout', rects: PdfRect[], targetPageIndex = pageIndex, displayRotation: DisplayRotation = 0) {
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    const additions: PdfOverlay[] = rects.map((rect) => ({
      id: makeId(style),
      type: 'markup',
      pageIndex: targetPageIndex,
      style,
      rect,
      color: [0.86, 0.15, 0.15],
      opacity: 0.95,
      thickness: 1.35,
      displayRotation,
    }))
    setOverlays((current) => [...current, ...additions])
    setDirty(true)
    showToast(`${additions.length === 1 ? 'Markup' : 'Markups'} added`)
  }

  function addInk(points: Array<{ x: number; y: number }>, targetPageIndex = pageIndex) {
    if (points.length < 2) return
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setOverlays((current) => [...current, {
      id: makeId('ink'),
      type: 'ink',
      pageIndex: targetPageIndex,
      points,
      color: [0.86, 0.15, 0.15],
      opacity: 0.95,
      thickness: 1.8,
    }])
    setDirty(true)
    showToast('Drawing added')
  }

  function addRectangle(rect: PdfRect, targetPageIndex = pageIndex) {
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setOverlays((current) => [...current, {
      id: makeId('rectangle'),
      type: 'markup',
      pageIndex: targetPageIndex,
      style: 'rectangle',
      rect,
      color: [0.86, 0.15, 0.15],
      opacity: 0.95,
      thickness: 1.5,
    }])
    setDirty(true)
    showToast('Rectangle added')
  }

  function commitTextEditValue(edit: PageTextEdit, announce = true) {
    if (!edit.modified) {
      setTextEdit(null)
      return
    }
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setOverlays((current) => {
      if (!edit.text && !edit.cover) {
        return edit.overlayId ? current.filter((overlay) => overlay.id !== edit.overlayId) : current
      }
      return upsertTextEdit(current, edit)
    })
    setTextEdit(null)
    setDirty(true)
    if (announce) showToast(edit.text ? (edit.cover ? 'Text updated' : 'Text added') : 'Text removed')
  }

  function commitTextEdit() {
    if (textEdit) commitTextEditValue(textEdit)
  }

  function commitObjectEditValue(edit: PageObjectEdit, announce = true) {
    if (!edit.modified) {
      setObjectEdit(null)
      return
    }
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setOverlays((current) => upsertObjectEdit(current, edit))
    setObjectEdit(null)
    setDirty(true)
    if (announce) showToast(edit.dataUrl ? `${edit.label} updated` : `${edit.label} removed`)
  }

  function commitObjectEdit() {
    if (objectEdit) commitObjectEditValue(objectEdit)
  }

  function beginTextEdit(edit: PageTextEdit) {
    if (textEdit && textEdit !== edit) commitTextEditValue(textEdit, false)
    if (objectEdit) commitObjectEditValue(objectEdit, false)
    setObjectEdit(null)
    setSelectingObjectRegion(false)
    setPageIndex(edit.pageIndex)
    setSelectedPages(new Set([edit.pageIndex]))
    setTextEdit(edit)
    setTool('edit')
  }

  function beginObjectEdit(edit: PageObjectEdit) {
    if (textEdit) commitTextEditValue(textEdit, false)
    if (objectEdit && objectEdit !== edit) commitObjectEditValue(objectEdit, false)
    setTextEdit(null)
    setSelectingObjectRegion(false)
    setPageIndex(edit.pageIndex)
    setSelectedPages(new Set([edit.pageIndex]))
    setObjectEdit(edit)
    setTool('edit')
  }

  function cancelEditSelection() {
    setTextEdit(null)
    setObjectEdit(null)
    setSelectingObjectRegion(false)
  }

  function deleteEditSelection() {
    if (textEdit) {
      commitTextEditValue({ ...textEdit, text: '', modified: true })
      return
    }
    if (!objectEdit) return
    if (!objectEdit.cover && !objectEdit.originalRect) {
      if (objectEdit.overlayId) {
        const snapshot = currentSnapshot()
        if (snapshot) addUndoSnapshot(snapshot)
        setOverlays((current) => current.filter((overlay) => overlay.id !== objectEdit.overlayId))
        setDirty(true)
      }
      setObjectEdit(null)
      showToast('Image removed')
      return
    }
    commitObjectEditValue({ ...objectEdit, dataUrl: undefined, modified: true })
  }

  async function addImageObject() {
    if (!pdf) return
    try {
      const picked = await window.simple.pickImage()
      if (!picked) return
      const dimensions = await imageDimensions(picked.dataUrl)
      const page = await pdf.getPage(pageIndex + 1)
      const pageViewport = page.getViewport({ scale: 1 })
      // The rect lives in unrotated PDF space; size and orient the image
      // against the page edges the user sees so it reads correctly as
      // displayed, and store its pixels in unrotated PDF orientation.
      const displayRotation = (((((page.rotate || 0) + (pageRotations[pageIndex] || 0)) % 360) + 360) % 360) as DisplayRotation
      const sideways = displayRotation === 90 || displayRotation === 270
      const maxWidth = (sideways ? pageViewport.height : pageViewport.width) * 0.48
      const maxHeight = (sideways ? pageViewport.width : pageViewport.height) * 0.48
      const scale = Math.min(1, maxWidth / dimensions.width, maxHeight / dimensions.height)
      const displayedWidth = Math.max(36, dimensions.width * scale)
      const displayedHeight = Math.max(36, dimensions.height * scale)
      const width = sideways ? displayedHeight : displayedWidth
      const height = sideways ? displayedWidth : displayedHeight
      const left = (pageViewport.width - width) / 2
      const top = (pageViewport.height - height) / 2
      const rect = viewportRectToPdf(pageViewport, { left, top, right: left + width, bottom: top + height })
      beginObjectEdit({
        pageIndex,
        kind: 'image',
        rect,
        dataUrl: await normalizeImageOrientation(picked.dataUrl, displayRotation),
        opacity: 1,
        cover: false,
        displayRotation,
        label: picked.name,
        modified: true,
      })
    } catch (error) {
      showToast(errorMessage(error))
    }
  }

  async function placeSignature(targetPageIndex: number, point: { x: number; y: number }, displayRotation: DisplayRotation) {
    const signature = pendingSignature
    if (!pdf || !signature) return
    try {
      const page = await pdf.getPage(targetPageIndex + 1)
      const view = page.view
      const pageBounds: PdfRect = {
        x: Math.min(view[0], view[2]),
        y: Math.min(view[1], view[3]),
        width: Math.abs(view[2] - view[0]),
        height: Math.abs(view[3] - view[1]),
      }
      // The rect lives in unrotated PDF space; size against the page edge the
      // user sees as its width so the signature reads correctly as displayed.
      const sideways = displayRotation === 90 || displayRotation === 270
      const displayedPageWidth = sideways ? pageBounds.height : pageBounds.width
      const scale = Math.min(1, displayedPageWidth * 0.4 / signature.width)
      const width = sideways ? signature.height * scale : signature.width * scale
      const height = sideways ? signature.width * scale : signature.height * scale
      beginObjectEdit({
        pageIndex: targetPageIndex,
        kind: 'image',
        rect: {
          x: clamp(point.x - width / 2, pageBounds.x, pageBounds.x + Math.max(0, pageBounds.width - width)),
          y: clamp(point.y - height / 2, pageBounds.y, pageBounds.y + Math.max(0, pageBounds.height - height)),
          width,
          height,
        },
        // The signature pixels are upright as displayed; store them rotated
        // into unrotated PDF orientation so the saver's axis-aligned draw
        // matches the axis-swapped rect above on rotated pages.
        dataUrl: await normalizeImageOrientation(signature.dataUrl, displayRotation),
        opacity: 1,
        cover: false,
        displayRotation,
        label: 'Signature',
        modified: true,
      })
      // Cleared only now, batched with beginObjectEdit's tool switch: clearing
      // before the awaits would mount the signature panel for a frame on every
      // placement (tool still 'sign' with no pending signature).
      setPendingSignature(null)
    } catch (error) {
      showToast(errorMessage(error))
    }
  }

  async function replaceObjectImage() {
    if (!objectEdit) return
    try {
      const picked = await window.simple.pickImage()
      if (!picked) return
      setObjectEdit({
        ...objectEdit,
        // The picked pixels are upright as displayed; store them in unrotated
        // PDF orientation like every other object image.
        dataUrl: await normalizeImageOrientation(picked.dataUrl, objectEdit.displayRotation || 0),
        kind: 'image',
        label: picked.name,
        modified: true,
      })
    } catch (error) {
      showToast(errorMessage(error))
    }
  }

  async function rotateObject(direction: 'left' | 'right') {
    if (!objectEdit?.dataUrl) return
    try {
      const rotated = await rotateImageData(objectEdit.dataUrl, direction)
      const centerX = objectEdit.rect.x + objectEdit.rect.width / 2
      const centerY = objectEdit.rect.y + objectEdit.rect.height / 2
      const width = objectEdit.rect.height
      const height = objectEdit.rect.width
      setObjectEdit({
        ...objectEdit,
        dataUrl: rotated,
        rect: { x: centerX - width / 2, y: centerY - height / 2, width, height },
        modified: true,
      })
    } catch (error) {
      showToast(errorMessage(error))
    }
  }

  function duplicateObject() {
    if (!objectEdit?.dataUrl) return
    beginObjectEdit({
      pageIndex: objectEdit.pageIndex,
      kind: objectEdit.kind,
      rect: { ...objectEdit.rect, x: objectEdit.rect.x + 12, y: objectEdit.rect.y - 12 },
      dataUrl: objectEdit.dataUrl,
      opacity: objectEdit.opacity,
      cover: false,
      label: `${objectEdit.label} copy`,
      modified: true,
    })
  }

  function copyCurrentEdit(removeAfterCopy = false) {
    const payload = copyPageEditSelection(textEdit, objectEdit)
    if (!payload) return false
    editClipboardRef.current = payload
    setHasEditClipboard(true)
    const plainText = payload.kind === 'text' ? payload.edit.text : payload.edit.label
    void navigator.clipboard?.writeText(plainText).catch(() => {})
    if (removeAfterCopy) deleteEditSelection()
    else showToast(`${payload.kind === 'text' ? 'Text box' : payload.edit.label} copied`)
    return true
  }

  async function pasteCurrentEditClipboard() {
    if (!pdf || !editClipboardRef.current) return false
    try {
      const targetPage = await pdf.getPage(pageIndex + 1)
      const view = targetPage.view
      const pageBounds: PdfRect = {
        x: Math.min(view[0], view[2]),
        y: Math.min(view[1], view[3]),
        width: Math.abs(view[2] - view[0]),
        height: Math.abs(view[3] - view[1]),
      }
      const pasted = pasteEditToPage(editClipboardRef.current, pageIndex, pageBounds)
      if (pasted.kind === 'text') beginTextEdit(pasted.edit)
      else beginObjectEdit(pasted.edit)
      showToast(`${pasted.kind === 'text' ? 'Text box' : pasted.edit.label} pasted`)
      return true
    } catch (error) {
      showToast(errorMessage(error))
      return false
    }
  }

  function duplicateText() {
    if (!textEdit) return
    editClipboardRef.current = copyPageEditSelection(textEdit, null)
    setHasEditClipboard(true)
    void pasteCurrentEditClipboard()
  }

  function toggleBookmark() {
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    const existing = bookmarks.find((bookmark) => bookmark.pageIndex === pageIndex)
    if (existing) {
      setBookmarks((current) => current.filter((bookmark) => bookmark.id !== existing.id))
      showToast('Bookmark removed')
    } else {
      setBookmarks((current) => [...current, { id: makeId('bookmark'), pageIndex, label: `Page ${pageIndex + 1}`, depth: 0, source: 'simple' }])
      showToast('Bookmark saved')
    }
    setBookmarksDirty(true)
    setDirty(true)
  }

  function deleteBookmark(id: string) {
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setBookmarks((current) => current.filter((bookmark) => bookmark.id !== id))
    setBookmarksDirty(true)
    setDirty(true)
  }

  function renameBookmark(id: string, label: string) {
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setBookmarks((current) => current.map((bookmark) => bookmark.id === id ? { ...bookmark, label } : bookmark))
    setBookmarksDirty(true)
    setDirty(true)
  }

  function goToPage(index: number) {
    if (!pdf) return
    if (textEdit) commitTextEditValue(textEdit, false)
    if (objectEdit) commitObjectEditValue(objectEdit, false)
    setSelectingObjectRegion(false)
    const next = clamp(index, 0, pdf.numPages - 1)
    setPageIndex(next)
    setSelectedPages(new Set([next]))
    lastSelectedPage.current = next
  }

  function selectPage(index: number, event: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean }) {
    if (textEdit) commitTextEditValue(textEdit, false)
    if (objectEdit) commitObjectEditValue(objectEdit, false)
    setSelectingObjectRegion(false)
    setPageIndex(index)
    if (event.shiftKey) {
      const from = Math.min(lastSelectedPage.current, index)
      const to = Math.max(lastSelectedPage.current, index)
      setSelectedPages(new Set(Array.from({ length: to - from + 1 }, (_, offset) => from + offset)))
    } else if (event.ctrlKey || event.metaKey) {
      setSelectedPages((current) => {
        const next = new Set(current)
        if (next.has(index) && next.size > 1) next.delete(index)
        else next.add(index)
        return next
      })
      lastSelectedPage.current = index
    } else {
      setSelectedPages(new Set([index]))
      lastSelectedPage.current = index
    }
  }

  const updateCurrentPageFromScroll = useCallback((index: number) => {
    setPageIndex(index)
    setSelectedPages((current) => {
      if (current.size > 1) return current
      lastSelectedPage.current = index
      return new Set([index])
    })
  }, [])

  const fitView = useCallback(async () => {
    if (!pdf || !viewerRef.current) return
    const requestGeneration = ++fitRequestGenerationRef.current
    try {
      const currentPage = await pdf.getPage(pageIndex + 1)
      if (requestGeneration !== fitRequestGenerationRef.current) return
      const pageRotation = (((currentPage.rotate || 0) + (pageRotations[pageIndex] || 0)) % 360 + 360) % 360
      const viewport = currentPage.getViewport({ scale: 1, rotation: pageRotation })
      const viewer = viewerRef.current
      // A fitted zoom rebuilds every text span. Keep the currently selected
      // glyph nodes fixed until the pointer/native selection is released.
      if (viewerHasActiveTextSelection(viewer)) return
      const next = fittedZoom(viewer, viewport)
      setZoom((current) => Math.abs(current - next) < 0.002 ? current : next)
    } catch (error) {
      if (requestGeneration !== fitRequestGenerationRef.current) return
      showToast(errorMessage(error))
    }
  }, [pdf, pageIndex, pageRotations, showToast])

  useEffect(() => {
    if (!pdf || zoomMode !== 'fit') return
    const viewer = viewerRef.current
    if (!viewer) return
    let frame = 0
    const applyFit = () => {
      window.cancelAnimationFrame(frame)
      frame = window.requestAnimationFrame(() => { void fitView() })
    }
    applyFit()
    const observer = new ResizeObserver(applyFit)
    observer.observe(viewer)
    const applyFitWhenSelectionClears = () => {
      if (!viewerHasActiveTextSelection(viewer)) applyFit()
    }
    document.addEventListener('selectionchange', applyFitWhenSelectionClears)
    document.addEventListener('pointerup', applyFitWhenSelectionClears, true)
    document.addEventListener('pointercancel', applyFitWhenSelectionClears, true)
    return () => {
      fitRequestGenerationRef.current += 1
      window.cancelAnimationFrame(frame)
      observer.disconnect()
      document.removeEventListener('selectionchange', applyFitWhenSelectionClears)
      document.removeEventListener('pointerup', applyFitWhenSelectionClears, true)
      document.removeEventListener('pointercancel', applyFitWhenSelectionClears, true)
    }
  }, [pdf, zoomMode, fitView, pageRotations, sidebarOpen, tool, immersive])

  function changeZoom(next: number) {
    setZoomMode('custom')
    setZoom(next)
  }

  async function toggleImmersive() {
    try {
      if (document.fullscreenElement) {
        await document.exitFullscreen()
        setImmersive(false)
        return
      }
      await document.documentElement.requestFullscreen()
      setImmersive(true)
    } catch (error) {
      showToast(errorMessage(error))
    }
  }

  useEffect(() => {
    const syncFullscreen = () => setImmersive(Boolean(document.fullscreenElement))
    document.addEventListener('fullscreenchange', syncFullscreen)
    return () => document.removeEventListener('fullscreenchange', syncFullscreen)
  }, [])

  function beginPan(event: ReactPointerEvent<HTMLElement>) {
    if (tool !== 'hand' || event.button !== 0) return
    if ((event.target as HTMLElement).closest('button,input,textarea,select,a')) return
    const viewer = event.currentTarget
    panRef.current = {
      x: event.clientX,
      y: event.clientY,
      scrollLeft: viewer.scrollLeft,
      scrollTop: viewer.scrollTop,
    }
    viewer.setPointerCapture(event.pointerId)
    setIsPanning(true)
    event.preventDefault()
  }

  function updatePan(event: ReactPointerEvent<HTMLElement>) {
    const start = panRef.current
    if (!start) return
    const viewer = event.currentTarget
    viewer.scrollLeft = start.scrollLeft - (event.clientX - start.x)
    viewer.scrollTop = start.scrollTop - (event.clientY - start.y)
    event.preventDefault()
  }

  function endPan(event: ReactPointerEvent<HTMLElement>) {
    if (!panRef.current) return
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
    panRef.current = null
    setIsPanning(false)
  }

  function changeTool(nextTool: ToolMode) {
    if (nextTool !== tool) {
      if (textEdit) commitTextEditValue(textEdit, false)
      if (objectEdit) commitObjectEditValue(objectEdit, false)
      setSelectingObjectRegion(false)
    }
    // Reactivating the sign tool reopens the panel to pick another signature.
    setPendingSignature(null)
    setTool(nextTool)
  }

  function handleKeyDown(event: KeyboardEvent) {
    if (event.defaultPrevented) return
    if (printDialogOpen) {
      // The dialog owns Escape and Enter itself; ignore every other shortcut
      // so typing a range or tabbing between fields cannot switch tools.
      if (event.key === 'Escape') { event.preventDefault(); setPrintDialogOpen(false) }
      return
    }
    if (tool === 'sign' && pdf && !pendingSignature) {
      // The signature dialog owns Escape/Enter/Tab while focus sits inside it;
      // ignore every other shortcut so a stray tool key cannot close the
      // dialog and silently discard an unsaved drawing.
      if (event.key === 'Escape') { event.preventDefault(); setTool('select') }
      return
    }
    const typing = isTypingTarget(event.target)
    const control = event.ctrlKey || event.metaKey
    const unmodified = !control && !event.altKey
    const key = event.key.toLocaleLowerCase()
    const directTextEditor = event.target instanceof HTMLTextAreaElement
      && event.target.getAttribute('aria-label') === 'Edit text directly on the PDF'
    const nativeTextRange = directTextEditor && event.target.selectionStart !== event.target.selectionEnd
    if (control && (key === 'c' || key === 'x') && (textEdit || objectEdit) && !nativeTextRange) {
      event.preventDefault()
      copyCurrentEdit(key === 'x')
      return
    }
    if (control && key === 'v' && !typing && editClipboardRef.current && (tool === 'edit' || tool === 'addText')) {
      event.preventDefault()
      void pasteCurrentEditClipboard()
      return
    }
    if (control && event.key.toLocaleLowerCase() === 'f' && pdf) {
      event.preventDefault()
      if (immersive) {
        if (document.fullscreenElement) void document.exitFullscreen()
        setImmersive(false)
      }
      setSidebarOpen(true)
      setSearchRequestId((request) => request + 1)
      return
    }
    if (control && event.key.toLocaleLowerCase() === 'o') { event.preventDefault(); openFile(); return }
    if (control && event.key.toLocaleLowerCase() === 's') { event.preventDefault(); save(event.shiftKey); return }
    if (control && event.key.toLocaleLowerCase() === 'p') { event.preventDefault(); if (pdf && bytes && documentFile) setPrintDialogOpen(true); return }
    if (control && event.key.toLocaleLowerCase() === 'z') { event.preventDefault(); undo(); return }
    if (control && event.key.toLocaleLowerCase() === 'y') { event.preventDefault(); redo(); return }
    if (event.key === 'F4') {
      event.preventDefault()
      if (sidebarOpen) setActiveSearchMatch(null)
      setSidebarOpen((open) => !open)
      return
    }
    if (event.key === 'Escape') {
      if (immersive) {
        event.preventDefault()
        if (document.fullscreenElement) void document.exitFullscreen()
        setImmersive(false)
        return
      }
      if (tool === 'sign') {
        setPendingSignature(null)
        setTool('select')
        return
      }
      if (textEdit || objectEdit || selectingObjectRegion) cancelEditSelection()
      else window.getSelection()?.removeAllRanges()
      return
    }
    if (!pdf || typing) return
    if (event.key === 'PageUp') { event.preventDefault(); goToPage(pageIndex - 1) }
    else if (event.key === 'PageDown') { event.preventDefault(); goToPage(pageIndex + 1) }
    else if (control && event.key === 'Home') { event.preventDefault(); goToPage(0) }
    else if (control && event.key === 'End') { event.preventDefault(); goToPage(pdf.numPages - 1) }
    else if (control && (event.key === '+' || event.key === '=')) { event.preventDefault(); setZoomMode('custom'); setZoom((value) => clamp(value + 0.15, 0.35, 4)) }
    else if (control && event.key === '-') { event.preventDefault(); setZoomMode('custom'); setZoom((value) => clamp(value - 0.15, 0.35, 4)) }
    else if (control && event.key === '0') { event.preventDefault(); setZoomMode('fit'); void fitView() }
    else if (unmodified && event.key.toLocaleLowerCase() === 'h') changeTool('hand')
    else if (unmodified && event.key.toLocaleLowerCase() === 'v') changeTool('select')
    else if (unmodified && event.key.toLocaleLowerCase() === 'e') changeTool('edit')
    else if (unmodified && event.key.toLocaleLowerCase() === 't') changeTool('addText')
    else if (unmodified && event.key.toLocaleLowerCase() === 'c') changeTool('crop')
    else if (event.key === 'Delete' && (textEdit || objectEdit)) deleteEditSelection()
    else if (event.key === 'Delete' && (document.activeElement as HTMLElement | null)?.closest('.sidebar')) deletePages()
  }

  const handleKeyDownRef = useRef(handleKeyDown)
  handleKeyDownRef.current = handleKeyDown

  useEffect(() => {
    const listener = (event: KeyboardEvent) => handleKeyDownRef.current(event)
    window.addEventListener('keydown', listener)
    return () => window.removeEventListener('keydown', listener)
  }, [])

  function closeWindow() {
    if ((dirty || textEdit?.modified || objectEdit?.modified) && !window.confirm('Close simple and discard unsaved changes?')) return
    window.simple.close()
  }

  const closeWindowRef = useRef(closeWindow)
  closeWindowRef.current = closeWindow

  useEffect(() => window.simple.onCloseRequested(() => closeWindowRef.current()), [])

  function removeRecent(filePath: string) {
    setRecentFiles((current) => {
      const next = current.filter((item) => item.path !== filePath)
      localStorage.setItem(RECENT_KEY, JSON.stringify(next))
      return next
    })
  }

  const titleBar = <TitleBar fileName={documentFile?.name} dirty={Boolean(dirty || textEdit?.modified || objectEdit?.modified)} onClose={closeWindow} />

  if (!documentFile || !bytes) {
    return (
      <div className="app-root">
        {titleBar}
        <Welcome
          recentFiles={recentFiles}
          isDragging={isDragging}
          onOpen={openFile}
          onOpenRecent={(file) => openPath(file.path)}
          onRemoveRecent={removeRecent}
          onDrop={handleDrop}
          onDragEnter={(event) => { event.preventDefault(); dragDepth.current += 1; setIsDragging(true) }}
          onDragLeave={(event) => { event.preventDefault(); dragDepth.current -= 1; if (dragDepth.current <= 0) setIsDragging(false) }}
        />
        {busy && <BusyOverlay label={busy} />}
        {toast && <Toast {...toast} onClose={() => setToast(null)} />}
      </div>
    )
  }

  return (
    <div className={`app-root document-app${immersive ? ' is-immersive' : ''}`}>
      {titleBar}
      <Toolbar
        sidebarOpen={sidebarOpen}
        pageIndex={pageIndex}
        pageCount={pdf?.numPages || 0}
        zoom={zoom}
        tool={tool}
        bookmarked={bookmarks.some((bookmark) => bookmark.pageIndex === pageIndex)}
        canUndo={Boolean(undoStack.length)}
        canRedo={Boolean(redoStack.length)}
        onToggleSidebar={() => {
          if (sidebarOpen) setActiveSearchMatch(null)
          setSidebarOpen((open) => !open)
        }}
        onOpen={openFile}
        onSave={() => save(false)}
        onUndo={undo}
        onRedo={redo}
        onTool={changeTool}
        onPage={goToPage}
        onZoom={changeZoom}
        onPrint={() => { if (pdf) setPrintDialogOpen(true) }}
        onBookmark={toggleBookmark}
        onImmersive={() => { void toggleImmersive() }}
      />

      <div className="document-body">
        {sidebarOpen && !pdf && <aside className="sidebar" aria-hidden="true" />}
        {sidebarOpen && pdf && (
          <Sidebar
            pdf={pdf}
            pageIndex={pageIndex}
            selectedPages={selectedPages}
            pageRotations={pageRotations}
            bookmarks={bookmarks}
            searchRequestId={searchRequestId}
            onActiveSearchMatch={setActiveSearchMatch}
            onClose={() => { setActiveSearchMatch(null); setSidebarOpen(false) }}
            onPage={goToPage}
            onSelectPage={selectPage}
            onReorder={reorderPage}
            onInsertBlank={insertBlankPage}
            onAddPages={addPages}
            onDuplicatePages={duplicatePages}
            onRotateLeft={() => rotatePages(-90)}
            onRotateRight={() => rotatePages(90)}
            onDeletePages={deletePages}
            onPageDragStart={startPageDrag}
            onImportPagesAt={importDroppedPages}
            onExportPages={exportPages}
            onDeleteBookmark={deleteBookmark}
            onRenameBookmark={renameBookmark}
          />
        )}
        <main
          ref={viewerRef}
          className={`viewer${isPanning ? ' is-panning' : ''}`}
          onPointerDown={beginPan}
          onPointerMove={updatePan}
          onPointerUp={endPan}
          onPointerCancel={endPan}
          onWheel={(event) => {
            if (!event.ctrlKey) return
            event.preventDefault()
            setZoomMode('custom')
            setZoom((value) => clamp(value + (event.deltaY < 0 ? 0.1 : -0.1), 0.35, 4))
          }}
        >
          {documentLoading && !pdf && (
            <div className="document-loading">
              <div className="document-loading-mark"><RotateCcw className="spin" size={20} /></div>
              <strong>Opening {documentFile.name}</strong>
              <span>Preparing pages and selectable text…</span>
            </div>
          )}
          {documentError && (
            <div className="document-error">
              <span className="error-icon"><FileWarning size={20} /></span>
              <h2>We couldn’t open this PDF</h2>
              <p>{documentError}</p>
              <div>
                <Button icon={FolderOpen} variant="primary" onClick={openFile}>Open another file</Button>
                <Button variant="secondary" onClick={() => { setDocumentFile(null); setBytes(null) }}>Back home</Button>
              </div>
            </div>
          )}
          {pdf && (
            <ContinuousPdfViewer
              pdf={pdf}
              viewerRef={viewerRef}
              currentPage={pageIndex}
              selectedPages={selectedPages}
              zoom={zoom}
              rotations={pageRotations}
              tool={tool}
              overlays={overlays}
              formValues={formValues}
              textEdit={textEdit}
              objectEdit={objectEdit}
              activeSearchMatch={activeSearchMatch}
              selectingObjectRegion={selectingObjectRegion}
              onCurrentPage={updateCurrentPageFromScroll}
              onSelectPage={selectPage}
              onPageDragStart={startPageDrag}
              onImportPagesAt={importDroppedPages}
              onRequestTextEdit={beginTextEdit}
              onTextEditChange={setTextEdit}
              onCommitTextEdit={commitTextEdit}
              onCancelTextEdit={() => setTextEdit(null)}
              onRequestObjectEdit={beginObjectEdit}
              onObjectEditChange={setObjectEdit}
              onCommitObjectEdit={commitObjectEdit}
              onCancelObjectEdit={() => setObjectEdit(null)}
              onObjectRegionSelected={() => setSelectingObjectRegion(false)}
              onHighlight={(index, rects) => addHighlights(rects, index)}
              onTextMarkup={(index, style, rects, displayRotation) => addTextMarkup(style, rects, index, displayRotation)}
              onInk={(index, points) => addInk(points, index)}
              onRectangle={(index, rect) => addRectangle(rect, index)}
              onCrop={(index, rect) => cropPage(rect, index)}
              onPlaceSignature={(index, point, displayRotation) => { void placeSignature(index, point, displayRotation) }}
              onNavigate={goToPage}
              onFormChange={(name, value) => { setFormValues((current) => ({ ...current, [name]: value })); setDirty(true) }}
            />
          )}
          {passwordProtected && pdf && (
            <div className="readonly-badge"><AlertTriangle size={13} /> Protected PDF · reading only</div>
          )}
        </main>
        {(tool === 'edit' || tool === 'addText') && pdf && (
          <EditInspector
            textEdit={textEdit}
            objectEdit={objectEdit}
            selectingObjectRegion={selectingObjectRegion}
            onTextChange={setTextEdit}
            onObjectChange={setObjectEdit}
            onCommitText={commitTextEdit}
            onCommitObject={commitObjectEdit}
            onCancelSelection={cancelEditSelection}
            onDeleteSelection={deleteEditSelection}
            onCopySelection={() => { copyCurrentEdit(false) }}
            onPasteSelection={() => { void pasteCurrentEditClipboard() }}
            canPasteSelection={hasEditClipboard}
            onDuplicateText={duplicateText}
            onAddText={() => { cancelEditSelection(); setTool('addText') }}
            onAddImage={addImageObject}
            onReplaceImage={replaceObjectImage}
            onRotateObject={rotateObject}
            onDuplicateObject={duplicateObject}
            onSelectObjectRegion={() => {
              if (textEdit) commitTextEditValue(textEdit, false)
              if (objectEdit) commitObjectEditValue(objectEdit, false)
              setTextEdit(null)
              setObjectEdit(null)
              setTool('edit')
              setSelectingObjectRegion((current) => !current)
            }}
            onClose={() => {
              if (textEdit) commitTextEditValue(textEdit, false)
              if (objectEdit) commitObjectEditValue(objectEdit, false)
              setSelectingObjectRegion(false)
              setTool('select')
            }}
          />
        )}
      </div>

      {pdf && <StatusBar tool={tool} pageIndex={pageIndex} pageCount={pdf.numPages} selectionCount={selectedPages.size} />}
      {tool === 'sign' && pdf && !pendingSignature && (
        <SignaturePanel
          onUse={setPendingSignature}
          onClose={() => setTool('select')}
        />
      )}
      {printDialogOpen && pdf && (
        <PrintDialog
          documentName={documentFile.name}
          pageCount={pdf.numPages}
          currentPage={pageIndex}
          selectedPages={[...selectedPages].sort((a, b) => a - b)}
          onPrint={(submission) => { void runPrintJob(submission) }}
          onPreview={() => { setPrintDialogOpen(false); void printDocument() }}
          onClose={() => setPrintDialogOpen(false)}
        />
      )}
      {busy && <BusyOverlay label={busy} />}
      {toast && <Toast {...toast} onClose={() => setToast(null)} />}
    </div>
  )
}
