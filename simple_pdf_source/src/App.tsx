import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { DragEvent, MouseEvent as ReactMouseEvent, PointerEvent as ReactPointerEvent } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  AlertTriangle,
  ClipboardPaste,
  Copy,
  FileWarning,
  FolderOpen,
  Lock,
  Redo2,
  RotateCcw,
  Scissors,
  Signature,
  Trash2,
  Undo2,
} from 'lucide-react'
import { EditInspector } from './components/EditInspector'
import { ContinuousPdfViewer } from './components/ContinuousPdfViewer'
import { ExportDialog, type PdfExportSubmission } from './components/ExportDialog'
import { OcrDialog, OcrExportPrompt, OcrOffer, readAutoOcr, writeAutoOcr, type OcrDialogRequest } from './components/OcrDialog'
import type { PendingEditAt } from './components/PdfPage'
import { PrintDialog, type PrintDialogSubmission } from './components/PrintDialog'
import { SignaturePanel, type SignatureImage } from './components/SignaturePanel'
import { Sidebar } from './components/Sidebar'
import { StatusBar } from './components/StatusBar'
import { TitleBar } from './components/TitleBar'
import { Toolbar } from './components/Toolbar'
import { BusyOverlay, Button, Toast } from './components/ui'
import { Welcome } from './components/Welcome'
import { getPdfOutlineBookmarks, loadPdf, makeId, normalizeBytes, pdfRectToViewport, viewportRectToPdf } from './lib/pdf'
import {
  chooseEditPasteSource,
  copyEditSelection as copyPageEditSelection,
  editClipboardPlainText,
  pasteEditToPage,
  pastedTextBoxSize,
  type EditClipboardPayload,
} from './lib/editClipboard'
import { pageIndicesForTransfer } from './lib/pageTransfer'
import { convertImageToPng, isConvertibleImage } from './lib/imageTransfer'
import { extractPdfText, renderPdfPageImage } from './lib/pdfExport'
import {
  bookmarkSubtreeSize,
  bookmarksAfterPageDelete,
  bookmarksFromOutline,
  remapBookmarkPages,
  toolbarBookmarkFor,
  withoutBookmark,
} from './lib/bookmarks'
import { blockingProblems, mayBeEncrypted, problemText, signedCopyName, summarizeProblems } from './lib/openSave'
import { classifyPage, pageNeedsOcr, type PageScanState } from './lib/ocr/pageClassifier'
import { abortOcrRun, disposeOcrRuntime, runOcr, summarizeOcrRun, warmUpOcr, type OcrRunOutcome, type OcrRunProgress, type OcrScope } from './lib/ocr/ocrRunner'
import { bindOcrPage } from './lib/ocr/ocrCache'
import {
  applyScanPreparation,
  disposeScanEditWorker,
  finalizeScanEdit,
  scanPreparationResult,
  settleScanOverlays,
  type ScanPreparation,
} from './lib/ocr/scanEdit'
import { clamp, errorMessage, isTypingTarget, withoutExtension } from './lib/utils'
import { getSimpleIO } from './simple-io/io-client'
import type {
  ActiveSearchMatch,
  Bookmark,
  DisplayRotation,
  DocumentPayload,
  FlattenReport,
  OpenDocument,
  PageObjectEdit,
  PageTextEdit,
  PdfFormValue,
  PdfOverlay,
  PdfRect,
  PdfSaveProblem,
  RecentFile,
  ScanEditInfo,
  TextOverlay,
  ToolMode,
} from './types'

interface Snapshot {
  bytes: Uint8Array
  overlays: PdfOverlay[]
  formValues: Record<string, PdfFormValue>
  bookmarks: Bookmark[]
  bookmarksDirty: boolean
  pageRotations: Record<number, number>
  pageIndex: number
  /**
   * Identifies this exact document state. The document is clean when its
   * revision is the one last saved, so undoing back to it after a save shows
   * no unsaved changes, and undoing past it does.
   */
  revision: number
}

/** The edit selected on the page, before it is folded into the document. */
interface PendingEditState {
  textEdit: PageTextEdit | null
  objectEdit: PageObjectEdit | null
}

/** What a save would write, and what it could not write as asked. */
interface PreparedOutput {
  data: Uint8Array
  warnings: PdfSaveProblem[]
  failures: PdfSaveProblem[]
  signatureDetected: boolean
}

interface SaveReportState {
  /** 'blocked': nothing written yet; 'partial': written without these changes; 'notes': written, with these notes. */
  kind: 'blocked' | 'partial' | 'notes'
  problems: PdfSaveProblem[]
  onSaveAnyway?: () => void
}

interface ToastState {
  message: string
  action?: string
  onAction?: () => void
  /** Stays until dismissed (used for anything the user must read). */
  sticky?: boolean
}

// pdf.js PasswordResponses.INCORRECT_PASSWORD
const PDFJS_INCORRECT_PASSWORD = 2
const PASSWORD_CANCELLED = 'Opening the protected PDF was cancelled.'

interface PdfContextMenuState {
  x: number
  y: number
}

const RECENT_KEY = 'folio:recent-files:v1'
const MAX_HISTORY_BYTES = 150 * 1024 * 1024

function fittedZoom(viewer: HTMLElement, viewport: { width: number; height: number }, mode: 'fit' | 'width' = 'fit') {
  const availableWidth = Math.max(120, viewer.clientWidth - 104)
  const availableHeight = Math.max(120, viewer.clientHeight - 52)
  return clamp(mode === 'width' ? availableWidth / viewport.width : Math.min(availableWidth / viewport.width, availableHeight / viewport.height), 0.35, 4)
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

/** A scan edit's own copy of its geometry (patch images are immutable strings and stay shared). */
function cloneScanInfo(scan: ScanEditInfo): ScanEditInfo {
  return {
    ...scan,
    lineRect: { ...scan.lineRect },
    run: { ...scan.run, origin: { ...scan.run.origin }, dir: { ...scan.run.dir } },
    words: scan.words?.map((word) => ({ ...word, rect: { ...word.rect } })),
    patch: scan.patch ? { ...scan.patch, rect: { ...scan.patch.rect } } : undefined,
    replace: scan.replace
      ? { ...scan.replace, rect: { ...scan.replace.rect }, originalRect: { ...scan.replace.originalRect }, patch: scan.replace.patch ? { ...scan.replace.patch, rect: { ...scan.replace.patch.rect } } : undefined }
      : undefined,
  }
}

function cloneOverlay(overlay: PdfOverlay): PdfOverlay {
  if (overlay.type === 'ink') {
    return { ...overlay, points: overlay.points.map((point) => ({ ...point })) }
  }
  return {
    ...overlay,
    rect: { ...overlay.rect },
    ...('originalRect' in overlay && overlay.originalRect ? { originalRect: { ...overlay.originalRect } } : {}),
    ...(overlay.type === 'text' && overlay.inkRect ? { inkRect: { ...overlay.inkRect } } : {}),
    ...(overlay.type === 'text' && overlay.scan ? { scan: cloneScanInfo(overlay.scan) } : {}),
  } as PdfOverlay
}

function upsertTextEdit(current: PdfOverlay[], edit: PageTextEdit): PdfOverlay[] {
  if (!edit.modified) return current
  if (edit.scan) {
    // Scanned text: decide what actually changed (only those words are
    // replaced). Unchanged text adds nothing, and reverting an edited line
    // to what the page says removes its edit.
    const finalized = finalizeScanEdit(edit)
    if (finalized.kind === 'none') return edit.overlayId ? current.filter((item) => item.id !== edit.overlayId) : current
    edit = finalized.edit
  }
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
    inkRect: edit.inkRect ? { ...edit.inkRect } : undefined,
    text: canReflowSourceItem
      ? `${sourceItemText.slice(0, sourceStart)}${edit.text}${sourceItemText.slice(sourceEnd)}`
      : edit.text,
    fontSize: edit.fontSize,
    fontFamily: edit.fontFamily,
    fontWeight: edit.fontWeight,
    fontStyle: edit.fontStyle,
    lineHeight: edit.lineHeight,
    textFit: edit.textFit,
    preserveSourceMetrics: edit.preserveSourceMetrics,
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
    ...(edit.scan ? { scan: cloneScanInfo(edit.scan) } : {}),
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

/** True when committing this text edit would leave nothing on the page (an empty box that replaced nothing). */
function isEmptyAddedText(edit: PageTextEdit) {
  return !edit.text && !edit.cover
}

/** The overlays with the selected edits folded in, exactly as committing them does. */
function overlaysWithPendingEdits(current: PdfOverlay[], textEdit: PageTextEdit | null, objectEdit: PageObjectEdit | null) {
  let next = current
  if (textEdit?.modified) {
    next = isEmptyAddedText(textEdit)
      ? (textEdit.overlayId ? next.filter((overlay) => overlay.id !== textEdit.overlayId) : next)
      : upsertTextEdit(next, textEdit)
  }
  if (objectEdit?.modified) next = upsertObjectEdit(next, objectEdit)
  return next
}

function blobToDataUrl(blob: Blob) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result || ''))
    reader.onerror = () => reject(reader.error || new Error('The clipboard image could not be read.'))
    reader.readAsDataURL(blob)
  })
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

async function flipImageData(dataUrl: string, direction: 'horizontal' | 'vertical') {
  const source = new Image()
  await new Promise<void>((resolve, reject) => {
    source.onload = () => resolve()
    source.onerror = () => reject(new Error('The image could not be flipped.'))
    source.src = dataUrl
  })
  const canvas = document.createElement('canvas')
  canvas.width = source.naturalWidth
  canvas.height = source.naturalHeight
  const context = canvas.getContext('2d')
  if (!context) throw new Error('The image could not be flipped.')
  context.translate(direction === 'horizontal' ? canvas.width : 0, direction === 'vertical' ? canvas.height : 0)
  context.scale(direction === 'horizontal' ? -1 : 1, direction === 'vertical' ? -1 : 1)
  context.drawImage(source, 0, 0)
  return canvas.toDataURL('image/png')
}

function reorientPastedRect(
  rect: PdfRect,
  pageBounds: PdfRect,
  sourceRotation: DisplayRotation,
  targetRotation: DisplayRotation,
) {
  if (Math.abs(targetRotation - sourceRotation) % 180 === 0) return rect
  const centerX = rect.x + rect.width / 2
  const centerY = rect.y + rect.height / 2
  const width = Math.min(rect.height, pageBounds.width)
  const height = Math.min(rect.width, pageBounds.height)
  return {
    x: clamp(centerX - width / 2, pageBounds.x, pageBounds.x + pageBounds.width - width),
    y: clamp(centerY - height / 2, pageBounds.y, pageBounds.y + pageBounds.height - height),
    width,
    height,
  }
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
  const [documentUpdating, setDocumentUpdating] = useState(false)
  const [documentError, setDocumentError] = useState('')
  const [passwordProtected, setPasswordProtected] = useState(false)
  const [passwordRequest, setPasswordRequest] = useState<{ name: string; wrong: boolean; resolve: (password: string | null) => void } | null>(null)
  const passwordRequestRef = useRef<{ resolve: (password: string | null) => void } | null>(null)
  const [passwordDraft, setPasswordDraft] = useState('')
  const [dirty, setDirty] = useState(false)
  const [busy, setBusy] = useState('')
  const [toast, setToast] = useState<ToastState | null>(null)
  const [saveReport, setSaveReport] = useState<SaveReportState | null>(null)
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [searchRequestId, setSearchRequestId] = useState(0)
  const [activeSearchMatch, setActiveSearchMatch] = useState<ActiveSearchMatch | null>(null)
  const [pageIndex, setPageIndex] = useState(0)
  const [selectedPages, setSelectedPages] = useState<Set<number>>(new Set([0]))
  const [zoom, setZoom] = useState(1)
  const [zoomMode, setZoomMode] = useState<'fit' | 'width' | 'custom'>('fit')
  const [tool, setTool] = useState<ToolMode>('select')
  const [overlays, setOverlays] = useState<PdfOverlay[]>([])
  const [formValues, setFormValues] = useState<Record<string, PdfFormValue>>({})
  const [hasInteractiveForms, setHasInteractiveForms] = useState<boolean | null>(null)
  const [bookmarks, setBookmarks] = useState<Bookmark[]>([])
  const [bookmarksDirty, setBookmarksDirty] = useState(false)
  // When the document's outline cannot be read, editing the (empty) list and
  // saving would replace the outline the file really has.
  const [bookmarksReadOnly, setBookmarksReadOnly] = useState(false)
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
  const [printPreviewPdf, setPrintPreviewPdf] = useState<PDFDocumentProxy | null>(null)
  const [printPreviewBytes, setPrintPreviewBytes] = useState<Uint8Array | null>(null)
  const [exportDialogOpen, setExportDialogOpen] = useState(false)
  const [immersive, setImmersive] = useState(false)
  const [contextMenu, setContextMenu] = useState<PdfContextMenuState | null>(null)
  const dragDepth = useRef(0)
  const viewerRef = useRef<HTMLElement>(null)
  const contextMenuRef = useRef<HTMLDivElement>(null)
  const contextMenuFocusReturnRef = useRef<HTMLElement | null>(null)
  const panRef = useRef<{ x: number; y: number; scrollLeft: number; scrollTop: number } | null>(null)
  const editClipboardRef = useRef<EditClipboardPayload | null>(null)
  const [hasEditClipboard, setHasEditClipboard] = useState(false)
  // Plain text written to the system clipboard with the in-app payload; a
  // paste compares it with the clipboard to see whether something newer won.
  const editClipboardTextRef = useRef<string | null>(null)
  // What the current page holds (scan, recognised scan, born-digital text…).
  // Cheap: it reuses pdf.js' text and operator lists of the rendered page.
  const [pageScanState, setPageScanState] = useState<PageScanState | null>(null)
  useEffect(() => {
    let active = true
    setPageScanState(null)
    if (pdf) {
      pdf.getPage(pageIndex + 1).then(classifyPage).then((state) => {
        if (active) setPageScanState(state)
      }).catch(() => { /* Keep normal tools when the page cannot be classified. */ })
    }
    return () => { active = false }
  }, [pdf, pageIndex])
  // Recognize text (OCR): the dialog (and its progress), the Edit-mode offer,
  // the export check, and what to do once the recognised text is in place.
  const [ocrDialog, setOcrDialog] = useState<{ scope: OcrScope } | null>(null)
  const [ocrProgress, setOcrProgress] = useState<OcrRunProgress | null>(null)
  const [ocrStopping, setOcrStopping] = useState(false)
  const [ocrOffer, setOcrOffer] = useState<{ pageIndex: number; point: { x: number; y: number }; client: { x: number; y: number } } | null>(null)
  const [ocrExportPrompt, setOcrExportPrompt] = useState<{ job: PdfExportSubmission; pages: number[] } | null>(null)
  const [pendingEditAt, setPendingEditAt] = useState<PendingEditAt | null>(null)
  const ocrAbortRef = useRef<AbortController | null>(null)
  // Bumped when another document opens: a run still finishing must not apply its text.
  const ocrRunTokenRef = useRef(0)
  const ocrLanguageRef = useRef('eng')
  // Cache keys of freshly recognised pages, bound to the next document proxy.
  const pendingOcrBindingRef = useRef<Map<number, string> | null>(null)
  // A text export waiting for its pages to be recognised.
  const pendingExportRef = useRef<PdfExportSubmission | null>(null)
  // Pages (per document proxy) where the offer was declined or nothing was found: not offered again.
  const ocrQuietPagesRef = useRef(new WeakMap<PDFDocumentProxy, Set<number>>())
  const fitRequestGenerationRef = useRef(0)
  const editSelectionGenerationRef = useRef(0)
  const lastSelectedPage = useRef(0)
  const undoStackRef = useRef<Snapshot[]>([])
  const redoStackRef = useRef<Snapshot[]>([])
  const dragExportBytesRef = useRef<Uint8Array | null>(null)
  const documentOperationRef = useRef(false)
  const awaitingDocumentSwapRef = useRef(false)
  const pendingDocumentSwapRef = useRef<{ commit: () => void; rollback: () => void } | null>(null)
  const skipPdfLoadForBytesRef = useRef<Uint8Array | null>(null)
  const liveStateRef = useRef({ bytes, overlays, formValues, bookmarks, bookmarksDirty, pageRotations, pageIndex, dirty })
  // Document revisions: every change gets a new number (see Snapshot.revision).
  const revisionRef = useRef(0)
  const revisionCounterRef = useRef(0)
  // -1: this content has never been saved (for example a converted Word file).
  const savedRevisionRef = useRef(0)
  // Ctrl+S pressed while a page operation was running: save once it settles.
  const pendingSaveRef = useRef<{ forceDialog: boolean } | null>(null)
  // Steps inside the selected edit (inspector changes, nudges, image tools),
  // so Ctrl+Z undoes the last of them instead of the whole edit.
  const pendingEditHistoryRef = useRef<{ past: PendingEditState[]; future: PendingEditState[]; at: number }>({ past: [], future: [], at: 0 })
  // The form field being typed into, so a word is one undo step, not one per letter.
  const formEditRef = useRef<{ name: string; at: number } | null>(null)

  useEffect(() => {
    liveStateRef.current = { bytes, overlays, formValues, bookmarks, bookmarksDirty, pageRotations, pageIndex, dirty }
    undoStackRef.current = undoStack
    redoStackRef.current = redoStack
  }, [bytes, overlays, formValues, bookmarks, bookmarksDirty, pageRotations, pageIndex, dirty, undoStack, redoStack])

  useEffect(() => {
    dragExportBytesRef.current = null
  }, [bytes, overlays, formValues, bookmarks, bookmarksDirty, pageRotations, textEdit, objectEdit])

  // A scan edit brought back without its preparation (an undo step taken
  // before the patch arrived) gets the finished one.
  useEffect(() => {
    const key = textEdit?.scan?.status === 'pending' ? textEdit.scan.key : ''
    const result = key ? scanPreparationResult(key) : undefined
    if (result) setTextEdit((current) => current?.scan?.key === key ? applyScanPreparation(current, result) : current)
  }, [textEdit])

  const showToast = useCallback((message: string, action?: string, onAction?: () => void, options: { sticky?: boolean } = {}) => {
    setToast({ message, action, onAction, sticky: options.sticky })
  }, [])

  useEffect(() => {
    if (!toast || toast.sticky) return
    const timeout = window.setTimeout(() => setToast(null), 4200)
    return () => window.clearTimeout(timeout)
  }, [toast])

  useEffect(() => {
    if (!bytes) {
      setPdf(null)
      return
    }
    if (skipPdfLoadForBytesRef.current === bytes) {
      skipPdfLoadForBytesRef.current = null
      setDocumentLoading(false)
      return
    }
    let cancelled = false
    let candidatePdf: PDFDocumentProxy | null = null
    let adopted = false
    let passwordCancelled = false
    setDocumentLoading(true)
    setDocumentError('')
    loadPdf(bytes, (updatePassword, reason) => {
      // Encrypted files are decrypted to a copy when they are opened, so this
      // runs only if that detection missed one. Answering '' again would make
      // pdf.js ask forever: ask the user instead. The file then opens for
      // reading only, because it cannot be rewritten while it is encrypted.
      const respond = updatePassword as (password: string | Error) => void
      void requestPassword(documentFile?.name || 'This PDF', reason === PDFJS_INCORRECT_PASSWORD).then((password) => {
        if (cancelled || password === null) {
          passwordCancelled = true
          respond(new Error(PASSWORD_CANCELLED))
          return
        }
        setPasswordProtected(true)
        respond(password)
      })
    }).then(async (nextPdf) => {
      candidatePdf = nextPdf
      if (cancelled) {
        void nextPdf.destroy()
        candidatePdf = null
        return
      }
      const nextPageIndex = clamp(pageIndex, 0, Math.max(0, nextPdf.numPages - 1))
      const pendingSwap = pendingDocumentSwapRef.current
      // Establish the final fit scale before mounting the first page. Otherwise
      // the initial 100% text layer can briefly appear and then be replaced by
      // the fitted layer, moving the glyphs underneath an early pointer drag.
      if (!pendingSwap && zoomMode !== 'custom' && viewerRef.current) {
        const currentPage = await nextPdf.getPage(nextPageIndex + 1)
        if (cancelled) return
        const pageRotation = (((currentPage.rotate || 0) + (pageRotations[nextPageIndex] || 0)) % 360 + 360) % 360
        const viewport = currentPage.getViewport({ scale: 1, rotation: pageRotation })
        const nextZoom = fittedZoom(viewerRef.current, viewport, zoomMode)
        setZoom((current) => Math.abs(current - nextZoom) < 0.002 ? current : nextZoom)
      }
      pendingDocumentSwapRef.current = null
      pendingSwap?.commit()
      adopted = true
      setPdf(nextPdf)
      if (!pendingSwap) {
        setPageIndex(nextPageIndex)
        setSelectedPages((current) => {
          const next = new Set([...current].filter((index) => index < nextPdf.numPages))
          return next.size ? next : new Set([0])
        })
      }
    }).catch((error) => {
      if (!cancelled) {
        if (candidatePdf && !adopted) {
          void candidatePdf.destroy()
          candidatePdf = null
        }
        const pendingSwap = pendingDocumentSwapRef.current
        pendingDocumentSwapRef.current = null
        if (pendingSwap) {
          pendingSwap.rollback()
          showToast(`The page change was not applied — ${errorMessage(error)}`)
        } else if (passwordCancelled) {
          setDocumentFile(null)
          setBytes(null)
          showToast(PASSWORD_CANCELLED)
        } else {
          setDocumentError(errorMessage(error))
        }
      }
    }).finally(() => {
      if (!cancelled) {
        setDocumentLoading(false)
        setDocumentUpdating(false)
        if (awaitingDocumentSwapRef.current) {
          awaitingDocumentSwapRef.current = false
          documentOperationRef.current = false
        }
      }
    })
    return () => {
      cancelled = true
      if (candidatePdf && !adopted) void candidatePdf.destroy()
    }
  }, [bytes])

  // A parsed document remains valid throughout the render that displays it.
  // React runs this cleanup only after a replacement (or the empty state) has
  // committed, so canvases never render against an already-destroyed proxy.
  useEffect(() => () => {
    if (pdf) void pdf.destroy()
  }, [pdf])

  useEffect(() => () => {
    if (printPreviewPdf) void printPreviewPdf.destroy()
  }, [printPreviewPdf])

  // Recognised text has just become part of the document: index its pages
  // under the new proxy, so editing them can reuse the exact word boxes.
  useEffect(() => {
    const binding = pendingOcrBindingRef.current
    if (!pdf || !binding) return
    pendingOcrBindingRef.current = null
    for (const [index, contentKey] of binding) bindOcrPage(pdf, index, contentKey)
  }, [pdf])

  useEffect(() => {
    let cancelled = false
    if (!pdf) {
      setHasInteractiveForms(null)
      return
    }
    setHasInteractiveForms(null)
    pdf.getFieldObjects().then((fields) => {
      if (!cancelled) setHasInteractiveForms(Boolean(fields && Object.keys(fields).length))
    }).catch(() => {
      if (!cancelled) setHasInteractiveForms(null)
    })
    return () => { cancelled = true }
  }, [pdf])

  useEffect(() => {
    if (!pdf || bookmarkKey) return
    let cancelled = false
    const fingerprint = pdf.fingerprints?.[0] || `${documentFile?.name || 'document'}:${bytes?.byteLength || 0}`
    const key = bookmarkStorageKey(documentFile?.path || null, fingerprint)
    getPdfOutlineBookmarks(pdf).then((outline) => {
      if (cancelled) return
      setBookmarkKey(key)
      setBookmarksReadOnly(false)
      // Every outline item is kept, including headings and web links that do
      // not go to a page: saving reuses the document's own items, so an edit
      // to one bookmark never drops, re-parents or re-targets the others.
      const nativeBookmarks = bookmarksFromOutline(outline)
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
        setBookmarksDirty(false)
        setBookmarksReadOnly(true)
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

  function nextRevision() {
    revisionCounterRef.current += 1
    return revisionCounterRef.current
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
      revision: revisionRef.current,
    }
  }

  function capturePageMutation() {
    const committedPendingEdit = Boolean(textEdit?.modified || objectEdit?.modified)
    const nextOverlays = overlaysWithPendingEdits(overlays, textEdit, objectEdit)
    const snapshot = currentSnapshot()
    if (snapshot && committedPendingEdit) {
      // The state before the page change, with the selected edit applied: a
      // state that was never saved, so restoring it shows unsaved changes.
      snapshot.overlays = nextOverlays.map(cloneOverlay)
      snapshot.revision = nextRevision()
    }
    return { overlays: nextOverlays, snapshot }
  }

  function finishPageMutation(nextOverlays: PdfOverlay[]) {
    setOverlays(nextOverlays)
    setTextEdit(null)
    setObjectEdit(null)
    setSelectingObjectRegion(false)
  }

  function rollbackDocumentBytes(previousBytes: Uint8Array) {
    skipPdfLoadForBytesRef.current = previousBytes
    setBytes(previousBytes)
  }

  /**
   * Record the state before a change. Every document change goes through
   * here, so this is also where the document gets its new revision.
   */
  function addUndoSnapshot(snapshot: Snapshot) {
    setUndoStack((current) => {
      const bounded = boundHistoryStack([...current, snapshot])
      undoStackRef.current = bounded
      return bounded
    })
    redoStackRef.current = []
    setRedoStack([])
    revisionRef.current = nextRevision()
    formEditRef.current = null
  }

  function applySnapshotState(snapshot: Snapshot) {
    setOverlays(snapshot.overlays.map(cloneOverlay))
    setFormValues({ ...snapshot.formValues })
    setBookmarks(snapshot.bookmarks.map((bookmark) => ({ ...bookmark })))
    setBookmarksDirty(snapshot.bookmarksDirty)
    setPageRotations({ ...snapshot.pageRotations })
    setPageIndex(snapshot.pageIndex)
    setSelectedPages(new Set([snapshot.pageIndex]))
    revisionRef.current = snapshot.revision
    setDirty(snapshot.revision !== savedRevisionRef.current)
    formEditRef.current = null
    resetPendingEditHistory()
    setTextEdit(null)
    setObjectEdit(null)
    setSelectingObjectRegion(false)
  }

  function resetPendingEditHistory() {
    pendingEditHistoryRef.current = { past: [], future: [], at: 0 }
  }

  /**
   * Remember the selected edit before a discrete change to it (an inspector
   * control, an arrow-key nudge, an image tool) so Ctrl+Z can take back just
   * that change. Rapid changes (typing a number, dragging a slider) count as one.
   */
  function recordPendingEditStep() {
    const history = pendingEditHistoryRef.current
    const now = performance.now()
    if (!history.past.length || now - history.at > 600) {
      history.past.push({ textEdit, objectEdit })
      if (history.past.length > 60) history.past.shift()
    }
    history.at = now
    history.future = []
  }

  function changeTextEditStep(next: PageTextEdit) {
    recordPendingEditStep()
    setTextEdit(next)
  }

  function changeObjectEditStep(next: PageObjectEdit) {
    recordPendingEditStep()
    setObjectEdit(next)
  }

  function restoreSnapshot(snapshot: Snapshot, onCommit: () => void) {
    const currentBytes = liveStateRef.current.bytes
    if (snapshot.bytes === currentBytes) {
      applySnapshotState(snapshot)
      onCommit()
      return
    }
    if (!currentBytes) return
    documentOperationRef.current = true
    awaitingDocumentSwapRef.current = true
    setDocumentUpdating(true)
    pendingDocumentSwapRef.current = {
      commit: () => {
        applySnapshotState(snapshot)
        onCommit()
      },
      rollback: () => rollbackDocumentBytes(currentBytes),
    }
    setBytes(snapshot.bytes)
  }

  function undo() {
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    if (textEdit?.modified || objectEdit?.modified) {
      const history = pendingEditHistoryRef.current
      const step = history.past.pop()
      if (step) {
        // Take back only the last change made to the selection.
        history.future.push({ textEdit, objectEdit })
        history.at = 0
        setTextEdit(step.textEdit)
        setObjectEdit(step.objectEdit)
        return
      }
      // Nothing smaller left to undo: apply the edit as one history step and
      // undo that step, so Redo brings the whole edit back.
      const before = currentSnapshot()
      if (!before) return
      const applied: Snapshot = {
        ...before,
        overlays: overlaysWithPendingEdits(before.overlays, textEdit, objectEdit).map(cloneOverlay),
        revision: nextRevision(),
      }
      redoStackRef.current = [applied]
      setRedoStack([applied])
      resetPendingEditHistory()
      setTextEdit(null)
      setObjectEdit(null)
      setSelectingObjectRegion(false)
      showToast('Undid the edit', 'Redo', () => historyActionsRef.current.redo())
      return
    }
    const previous = undoStackRef.current.at(-1)
    const current = currentSnapshot()
    if (!previous || !current) return
    const nextUndo = undoStackRef.current.slice(0, -1)
    const nextRedo = boundHistoryStack([...redoStackRef.current, current])
    restoreSnapshot(previous, () => {
      undoStackRef.current = nextUndo
      redoStackRef.current = nextRedo
      setUndoStack(nextUndo)
      setRedoStack(nextRedo)
      showToast('Undid the last change')
    })
  }

  function redo() {
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    const editFuture = pendingEditHistoryRef.current.future
    if ((textEdit || objectEdit) && editFuture.length) {
      const step = editFuture.pop()!
      pendingEditHistoryRef.current.past.push({ textEdit, objectEdit })
      pendingEditHistoryRef.current.at = 0
      setTextEdit(step.textEdit)
      setObjectEdit(step.objectEdit)
      return
    }
    if (textEdit?.modified || objectEdit?.modified) {
      showToast('Finish the current edit before redoing document changes.')
      return
    }
    const next = redoStackRef.current.at(-1)
    const current = currentSnapshot()
    if (!next || !current) return
    const nextRedo = redoStackRef.current.slice(0, -1)
    const nextUndo = boundHistoryStack([...undoStackRef.current, current])
    restoreSnapshot(next, () => {
      redoStackRef.current = nextRedo
      undoStackRef.current = nextUndo
      setRedoStack(nextRedo)
      setUndoStack(nextUndo)
      showToast('Redid the last change')
    })
  }

  // Toast actions run later: they must call the current undo and redo, not
  // the ones of the render that showed the toast.
  const historyActionsRef = useRef({ undo, redo })
  historyActionsRef.current = { undo, redo }

  /**
   * Ask for a document's password. The dialog is rendered on the home screen
   * as well as over a document, so an open never waits on a prompt nobody
   * can see. A newer request replaces (cancels) an unanswered one.
   */
  function requestPassword(name: string, wrong: boolean) {
    passwordRequestRef.current?.resolve(null)
    return new Promise<string | null>((resolve) => {
      const request = {
        name,
        wrong,
        resolve: (password: string | null) => {
          if (passwordRequestRef.current === request) passwordRequestRef.current = null
          setPasswordRequest((current) => current === request ? null : current)
          resolve(password)
        },
      }
      passwordRequestRef.current = request
      setPasswordDraft('')
      setPasswordRequest(request)
    })
  }

  async function openPayload(payload: DocumentPayload) {
    let fileBytes = normalizeBytes(payload.data)
    let unlocked = false
    // Files read by the main process arrive with its whole-file answer.
    const encrypted = typeof payload.encrypted === 'boolean' ? payload.encrypted : mayBeEncrypted(fileBytes)
    if (encrypted) {
      let attempt = await window.simple.unlockPdf(fileBytes, '')
      let wrong = false
      while (attempt.status === 'needs-password' || attempt.status === 'wrong-password') {
        const password = await requestPassword(payload.name, wrong)
        if (password === null) return
        attempt = await window.simple.unlockPdf(fileBytes, password)
        wrong = attempt.status === 'wrong-password'
      }
      if (attempt.status === 'unlocked' && attempt.data) {
        fileBytes = normalizeBytes(attempt.data)
        unlocked = true
      }
    }
    const file: OpenDocument = {
      bytes: fileBytes,
      name: payload.name,
      // An unlocked copy must never silently overwrite the protected original.
      path: unlocked ? null : payload.path,
      sourcePath: payload.sourcePath ?? payload.path,
      converted: payload.converted,
      signatureDetected: payload.signatureDetected,
      unlocked,
    }
    revisionRef.current = 0
    revisionCounterRef.current = 0
    // A converted file has never been saved as a PDF; an unlocked copy has
    // no changes yet (Save still asks where to write it, path is null).
    savedRevisionRef.current = file.converted ? -1 : 0
    pendingSaveRef.current = null
    // A recognition still running belongs to the previous document.
    ocrRunTokenRef.current += 1
    abortOcrRun(ocrAbortRef.current)
    ocrAbortRef.current = null
    void disposeOcrRuntime().catch(() => {})
    disposeScanEditWorker()
    pendingOcrBindingRef.current = null
    pendingExportRef.current = null
    setOcrDialog(null)
    setOcrProgress(null)
    setOcrStopping(false)
    setOcrOffer(null)
    setOcrExportPrompt(null)
    setPendingEditAt(null)
    resetPendingEditHistory()
    setSaveReport(null)
    setBookmarksReadOnly(false)
    setDocumentFile(file)
    setExportDialogOpen(false)
    closePrintDialog()
    pendingDocumentSwapRef.current = null
    awaitingDocumentSwapRef.current = false
    documentOperationRef.current = false
    setDocumentUpdating(false)
    setBytes(file.bytes)
    setPasswordProtected(false)
    setDirty(file.converted)
    setPageIndex(0)
    setActiveSearchMatch(null)
    setZoomMode('fit')
    setSelectedPages(new Set([0]))
    setOverlays([])
    setFormValues({})
    setHasInteractiveForms(null)
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
    // Signed and unlocked documents get a persistent one-line notice instead.
    if (file.converted) showToast('Converted to PDF — save to choose where to keep it')
  }

  /** Save / Don't Save / Cancel before another document replaces this one (Save is the default). */
  async function canReplaceCurrent() {
    if (!(dirty || textEdit?.modified || objectEdit?.modified)) return true
    const io = getSimpleIO()
    if (!io) return window.confirm('This document has unsaved changes. Discard them and open another file?')
    const answer = await io.prompt(documentFile?.path ? 'prompts.unsaved' : 'prompts.unsaved-untitled', { name: documentFile?.name ?? 'this document', kind: 'PDF' })
    if (answer === 'dont-save') return true
    if (answer !== 'save') return false
    await save(false)
    return revisionRef.current === savedRevisionRef.current
  }

  async function openFile() {
    // Opening is blocked while a password prompt waits for its answer.
    if (passwordRequestRef.current) return
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
    if (passwordRequestRef.current && !documentFile) {
      showToast('Answer or cancel the password prompt first.')
      return
    }
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
    if (passwordRequestRef.current) return
    if (!(await canReplaceCurrent())) return
    const files = Array.from(event.dataTransfer.files)
    const file = files[0]
    if (!file) return
    setBusy('Opening document…')
    try {
      // A file dropped from disk opens by its path, exactly like Open: Save
      // writes back to it and it joins Recent files. Files without one
      // (dragged out of a browser or an archive) open from their bytes.
      const filePath = window.simple.getPathForFile?.(file) || ''
      const payload = filePath
        ? await window.simple.openPath(filePath)
        : await window.simple.openBytes(file.name, await file.arrayBuffer())
      await openPayload(payload)
      if (files.length > 1) showToast(`Opened ${file.name} — ${files.length - 1} more ${files.length === 2 ? 'file' : 'files'} ignored`)
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }

  async function preparedBytes() {
    const output = await prepareOutput()
    const blocking = blockingProblems(output)
    // Print, export and page drags never hand out a file that silently lacks an edit.
    if (blocking.length) throw new Error(`Some changes cannot be written: ${summarizeProblems(blocking)}`)
    return output.data
  }

  /**
   * The document as it would be written: the unchanged base plus every edit,
   * including the one selected on the page. The main process reports what it
   * could not write exactly as asked (form fields, text that does not fit…)
   * instead of dropping it.
   */
  async function prepareOutput(state = {
    overlays: overlaysWithPendingEdits(overlays, textEdit, objectEdit),
    formValues,
    pageRotations,
    bookmarks,
    bookmarksDirty,
  }): Promise<PreparedOutput> {
    if (!bytes) throw new Error('No PDF is open.')
    // An edit of scanned text committed before its patch was ready waits for
    // it here (briefly): the file never gets new text over the old words.
    state = { ...state, overlays: await settleScanOverlays(state.overlays) }
    const hasRotations = Object.values(state.pageRotations).some((amount) => amount % 360 !== 0)
    if (!state.overlays.length && !Object.keys(state.formValues).length && !hasRotations && !state.bookmarksDirty) {
      return { data: bytes, warnings: [], failures: [], signatureDetected: false }
    }
    const report: FlattenReport = await window.simple.flattenOverlays(bytes, state.overlays, state.formValues, {
      ...(hasRotations ? { pageRotations: state.pageRotations } : {}),
      ...(state.bookmarksDirty ? { bookmarks: state.bookmarks } : {}),
      report: true,
    })
    return {
      data: normalizeBytes(report.data),
      warnings: Array.isArray(report.warnings) ? report.warnings : [],
      failures: Array.isArray(report.failures) ? report.failures : [],
      signatureDetected: report.signatureDetected === true,
    }
  }

  /** Ctrl+S during a page operation saves as soon as it settles; it is never ignored. */
  function queueSave(forceDialog: boolean) {
    pendingSaveRef.current = { forceDialog: Boolean(pendingSaveRef.current?.forceDialog || forceDialog) }
    showToast('Will save when the current operation finishes')
  }

  async function save(forceDialog = false, acceptProblems = false) {
    if (!documentFile || !bytes) return
    if (passwordRequestRef.current) return
    if (documentOperationRef.current || documentUpdating) {
      queueSave(forceDialog)
      return
    }
    if (passwordProtected && (dirty || textEdit?.modified || objectEdit?.modified)) {
      showToast('This protected PDF is open for reading only, so changes cannot be saved.')
      return
    }
    setSaveReport(null)
    // What is being saved: every edit, with the one selected on the page
    // applied first as an undoable step (exactly as clicking away applies it).
    const state = {
      overlays: overlaysWithPendingEdits(overlays, textEdit, objectEdit),
      formValues,
      pageRotations,
      bookmarks,
      bookmarksDirty,
    }
    commitPendingEdits(false)
    // Typing after the save starts a new undo step.
    formEditRef.current = null
    const savedRevision = revisionRef.current
    documentOperationRef.current = true
    setBusy('Saving PDF…')
    try {
      const output = await prepareOutput(state)
      const blocking = blockingProblems(output)
      if (blocking.length && !acceptProblems) {
        // Nothing is written yet; the user decides (the edits stay in the document either way).
        setSaveReport({
          kind: 'blocked',
          problems: blocking,
          onSaveAnyway: () => {
            setSaveReport(null)
            // The current save(), not this render's: the edit selected then is committed now.
            void saveRef.current(forceDialog, true)
          },
        })
        return
      }
      // A signed original is never overwritten: Save writes a copy beside it.
      const signed = documentFile.signatureDetected || output.signatureDetected
      const result = await window.simple.savePdf({
        data: output.data,
        path: signed ? null : documentFile.path,
        name: signed ? signedCopyName(documentFile) : documentFile.name,
        forceDialog: forceDialog || signed,
      })
      if (!result) return
      // Keep the immutable editing base plus the logical overlays in memory.
      // Every later save is regenerated from that base, so a long edit session
      // never feeds already-flattened covers/replacements back into itself.
      const nextFile: OpenDocument = { ...documentFile, bytes, path: result.path, name: result.name, converted: false, signatureDetected: false, unlocked: false }
      const nextBookmarkKey = bookmarkStorageKey(result.path, pdf?.fingerprints?.[0] || `${result.name}:${output.data.byteLength}`)
      if (bookmarkKey && bookmarkKey !== nextBookmarkKey) localStorage.removeItem(bookmarkKey)
      setBookmarkKey(nextBookmarkKey)
      setDocumentFile(nextFile)
      // Undo and Redo survive a save. The document is clean only when every
      // edit was written and nothing changed while the file was being written.
      if (!blocking.length) {
        savedRevisionRef.current = savedRevision
        if (revisionRef.current === savedRevision) setDirty(false)
      }
      rememberFile(nextFile)
      const notes = output.warnings.filter((warning) => !warning.dataLoss)
      const showFile = () => window.simple.showItem(result.path)
      if (blocking.length) {
        showToast(`Saved without ${blocking.length === 1 ? '1 change' : `${blocking.length} changes`}; they are still unsaved here`, 'Details', () => setSaveReport({ kind: 'partial', problems: blocking }), { sticky: true })
      } else if (notes.length) {
        showToast(`Saved with ${notes.length === 1 ? 'a note' : `${notes.length} notes`}: ${summarizeProblems(notes)}`, 'Details', () => setSaveReport({ kind: 'notes', problems: notes }), { sticky: true })
      } else if (signed || result.keptSignedOriginal) {
        showToast('Saved as a copy; the signed original was not changed', 'Show file', showFile)
      } else {
        showToast('Saved successfully', 'Show file', showFile)
      }
    } catch (error) {
      showToast(`Not saved: ${errorMessage(error)}`, undefined, undefined, { sticky: true })
    } finally {
      documentOperationRef.current = false
      setBusy('')
    }
  }

  const saveRef = useRef(save)
  saveRef.current = save

  useEffect(() => {
    const pending = pendingSaveRef.current
    if (!pending || busy || documentUpdating || documentLoading || documentOperationRef.current) return
    pendingSaveRef.current = null
    if (!pending.forceDialog && !dirty && !textEdit?.modified && !objectEdit?.modified) return
    void save(pending.forceDialog)
    // Runs when an operation settles (a recognition that found nothing to
    // add ends without one); save() reads the latest state itself.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [busy, documentUpdating, documentLoading, ocrProgress])

  async function mutate(
    operation: Record<string, unknown>,
    label: string,
    after?: () => void,
    capturedSnapshot?: Snapshot | null,
  ) {
    if (!bytes) return false
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return false
    }
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return false
    }
    documentOperationRef.current = true
    setDocumentUpdating(true)
    const snapshot = capturedSnapshot === undefined ? currentSnapshot() : capturedSnapshot
    let awaitingSwap = false
    setBusy(label)
    try {
      const result = normalizeBytes(await window.simple.mutatePdf(bytes, operation))
      pendingDocumentSwapRef.current = {
        commit: () => {
          if (snapshot) addUndoSnapshot(snapshot)
          setDirty(true)
          after?.()
        },
        rollback: () => rollbackDocumentBytes(bytes),
      }
      awaitingDocumentSwapRef.current = true
      awaitingSwap = true
      setBytes(result)
      return true
    } catch (error) {
      showToast(errorMessage(error))
      return false
    } finally {
      if (!awaitingSwap) {
        documentOperationRef.current = false
        setDocumentUpdating(false)
      }
      setBusy('')
    }
  }

  function targetPages() {
    return selectedPages.size ? [...selectedPages].sort((a, b) => a - b) : [pageIndex]
  }

  /** Pages on screen keep their decoded images (image selection reuses them); others are released after reading. */
  function canReleasePage(index: number) {
    return !viewerRef.current?.querySelector(`.continuous-page-slot[data-page-index="${index}"] .page-surface`)
  }

  function openOcrDialog(scope: OcrScope = 'needed') {
    if (!pdf || !bytes) return
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    setOcrOffer(null)
    setOcrDialog({ scope })
  }

  function quietOcrPages(source: PDFDocumentProxy, indices: number[]) {
    let pages = ocrQuietPagesRef.current.get(source)
    if (!pages) ocrQuietPagesRef.current.set(source, pages = new Set())
    for (const index of indices) pages.add(index)
  }

  /**
   * Recognize text: read the pages, then write their text into the document
   * as an invisible layer (one undoable change, like every page operation).
   * The selected edit is applied with it and stays an edit. Stop keeps the
   * pages finished so far. `editAt` reopens the clicked line afterwards (the
   * Edit-mode offer); `exportJob` continues a text export.
   */
  async function recognizeText(request: OcrDialogRequest & { editAt?: { pageIndex: number; x: number; y: number }; exportJob?: PdfExportSubmission }) {
    if (!pdf || !bytes || !documentFile) return
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    const sourcePdf = pdf
    const sourceBytes = bytes
    const token = ++ocrRunTokenRef.current
    const controller = new AbortController()
    ocrAbortRef.current = controller
    ocrLanguageRef.current = request.language
    const captured = capturePageMutation()
    documentOperationRef.current = true
    setOcrOffer(null)
    setOcrStopping(false)
    setOcrDialog((current) => current ?? { scope: request.scope })
    setOcrProgress({ phase: 'checking', done: 0, total: request.pages.length, fraction: 0 })
    let outcome: OcrRunOutcome | null = null
    let failure: unknown = null
    let lastProgress = { at: 0, phase: '', done: -1 }
    try {
      outcome = await runOcr(sourcePdf, {
        pages: request.pages,
        explicit: request.explicit,
        language: request.language,
        replaceExisting: request.replaceExisting,
        pageRotations: { ...liveStateRef.current.pageRotations },
        signal: controller.signal,
        canReleasePage,
        onProgress: (progress) => {
          if (ocrRunTokenRef.current !== token) return
          // At most ten updates a second: rendering the app is not free while pages are read.
          const now = performance.now()
          if (now - lastProgress.at < 100 && progress.phase === lastProgress.phase && progress.done === lastProgress.done) return
          lastProgress = { at: now, phase: progress.phase, done: progress.done }
          setOcrProgress(progress)
        },
      })
    } catch (error) {
      failure = error
    } finally {
      if (ocrAbortRef.current === controller) ocrAbortRef.current = null
    }
    // Another document was opened meanwhile; it has reset the guard and the dialog.
    if (ocrRunTokenRef.current !== token) return
    documentOperationRef.current = false
    setOcrProgress(null)
    setOcrStopping(false)
    setOcrDialog(null)
    if (!outcome) {
      showToast(`Text was not recognized — ${errorMessage(failure)}`, undefined, undefined, { sticky: true })
      return
    }
    if (liveStateRef.current.bytes !== sourceBytes) return
    const operation = outcome.operation
    if (!operation) {
      // Pages without any text are not offered again for this document.
      quietOcrPages(sourcePdf, outcome.results.filter((result) => result.wordCount === 0).map((result) => result.pageIndex))
      if (outcome.fatal) showToast(outcome.fatal.message, undefined, undefined, { sticky: true })
      else showToast(summarizeOcrRun(outcome, false).message)
      const job = request.exportJob
      if (job && !outcome.fatal) window.setTimeout(() => { void runExportAsRef.current(job, { skipScanCheck: true }) }, 0)
      return
    }
    const summary = summarizeOcrRun(outcome, true)
    const message = outcome.fatal ? `${summary.message} · ${outcome.fatal.message}` : summary.message
    const contentKeys = outcome.contentKeys
    // A live selection freezes text-layer rebuilds; the new text must replace the old spans.
    window.getSelection()?.removeAllRanges()
    await mutate(operation as unknown as Record<string, unknown>, 'Adding recognized text…', () => {
      finishPageMutation(captured.overlays)
      pendingOcrBindingRef.current = contentKeys
      if (summary.sideways.length) showToast(message, 'Rotate', () => rotatePagesUpright(summary.sideways))
      else showToast(message, 'Undo', () => historyActionsRef.current.undo())
      if (request.editAt) setPendingEditAt({ ...request.editAt, token: Date.now() })
      if (request.exportJob) pendingExportRef.current = request.exportJob
    }, captured.snapshot)
  }

  function stopOcr() {
    const controller = ocrAbortRef.current
    if (!controller || controller.signal.aborted) return
    setOcrStopping(true)
    controller.abort()
  }

  /** A click in Edit mode on a scanned page without text: offer to recognise it (or just do it). */
  function requestOcrOffer(index: number, point: { x: number; y: number }, client: { x: number; y: number }) {
    if (!pdf || passwordProtected || documentOperationRef.current) return
    if (ocrQuietPagesRef.current.get(pdf)?.has(index)) return
    setPageIndex(index)
    setSelectedPages(new Set([index]))
    lastSelectedPage.current = index
    if (readAutoOcr()) {
      void recognizeText({ scope: 'current', pages: [index], explicit: true, language: ocrLanguageRef.current, replaceExisting: false, editAt: { pageIndex: index, ...point } })
      return
    }
    setOcrOffer({ pageIndex: index, point, client })
    // Most offers are accepted: start the engine while the offer is read.
    void warmUpOcr(ocrLanguageRef.current)
  }

  function acceptOcrOffer(automatically: boolean) {
    const offer = ocrOffer
    if (!offer) return
    if (automatically) writeAutoOcr(true)
    void recognizeText({ scope: 'current', pages: [offer.pageIndex], explicit: true, language: ocrLanguageRef.current, replaceExisting: false, editAt: { pageIndex: offer.pageIndex, ...offer.point } })
  }

  function declineOcrOffer() {
    if (ocrOffer && pdf) quietOcrPages(pdf, [ocrOffer.pageIndex])
    setOcrOffer(null)
  }

  /** Turn pages that were recognised sideways upright on screen (the toast's Rotate). */
  function rotatePagesUpright(items: Array<{ pageIndex: number; degrees: number }>) {
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    window.getSelection()?.removeAllRanges()
    const captured = capturePageMutation()
    if (captured.snapshot) addUndoSnapshot(captured.snapshot)
    finishPageMutation(captured.overlays)
    setPageRotations((current) => {
      const next = { ...current }
      for (const { pageIndex: index, degrees } of items) {
        const amount = (((next[index] || 0) + degrees) % 360 + 360) % 360
        if (amount) next[index] = amount
        else delete next[index]
      }
      return next
    })
    setDirty(true)
    showToast(`${items.length === 1 ? 'Page' : `${items.length} pages`} rotated`, 'Undo', () => historyActionsRef.current.undo())
  }

  /** Pages among these that are scans without any text (they would export empty). */
  async function scannedPagesIn(source: PDFDocumentProxy, indices: number[]) {
    const scanned: number[] = []
    for (const index of indices) {
      const page = await source.getPage(index + 1)
      if (await pageNeedsOcr(page, { release: canReleasePage(index) })) scanned.push(index)
    }
    return scanned
  }

  async function resolveHasInteractiveForms() {
    if (hasInteractiveForms !== null) return hasInteractiveForms
    if (!pdf) return false
    try {
      const fields = await pdf.getFieldObjects()
      const detected = Boolean(fields && Object.keys(fields).length)
      setHasInteractiveForms(detected)
      return detected
    } catch {
      setHasInteractiveForms(null)
      return null
    }
  }

  async function confirmFormSensitiveOperation(action: string) {
    const interactiveForms = await resolveHasInteractiveForms()
    if (interactiveForms === null) {
      showToast(`Could not verify this PDF's form structure, so ${action.toLocaleLowerCase()} is blocked to protect the document.`)
      return false
    }
    if (!interactiveForms) return true
    showToast(`${action} is blocked for fillable PDFs to protect form fields. Create a static copy from Print preview first.`)
    return false
  }

  function rotatePages(degrees = 90) {
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    const indices = targetPages()
    // A live native selection freezes text-layer rebuilds and page pruning;
    // rotated canvases must never sit beneath the old orientation's spans.
    window.getSelection()?.removeAllRanges()
    const captured = capturePageMutation()
    if (captured.snapshot) addUndoSnapshot(captured.snapshot)
    finishPageMutation(captured.overlays)
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
    if (!(await confirmFormSensitiveOperation('Duplicating pages'))) return
    const indices = targetPages()
    const captured = capturePageMutation()
    const sourceOverlays = captured.overlays
    if (await mutate({ type: 'duplicate', indices }, `Duplicating ${indices.length === 1 ? 'page' : 'pages'}…`, () => {
      const before = (index: number) => indices.filter((selected) => selected < index).length
      const through = (index: number) => indices.filter((selected) => selected <= index).length
      const remapped = sourceOverlays.map((overlay) => ({ ...cloneOverlay(overlay), pageIndex: overlay.pageIndex + before(overlay.pageIndex) }))
      const copies = sourceOverlays
        .filter((overlay) => indices.includes(overlay.pageIndex))
        .map((overlay) => ({ ...cloneOverlay(overlay), id: makeId(overlay.type), pageIndex: overlay.pageIndex + through(overlay.pageIndex) }))
      finishPageMutation([...remapped, ...copies])
      setPageRotations((current) => {
        const next: Record<number, number> = {}
        for (const [rawIndex, amount] of Object.entries(current)) {
          const index = Number(rawIndex)
          next[index + before(index)] = amount
          if (indices.includes(index)) next[index + through(index)] = amount
        }
        return next
      })
      setBookmarks((current) => remapBookmarkPages(current, (index) => index + before(index)))
      const duplicatedIndices = indices.map((index) => index + through(index))
      setPageIndex(duplicatedIndices[0])
      setSelectedPages(new Set(duplicatedIndices))
      showToast(`${indices.length === 1 ? 'Page duplicated' : `${indices.length} pages duplicated`}`)
    }, captured.snapshot)) return
  }

  async function insertBlankPage() {
    if (!pdf) return
    if (!(await confirmFormSensitiveOperation('Adding a blank page'))) return
    const currentPage = await pdf.getPage(pageIndex + 1)
    const baseViewport = currentPage.getViewport({ scale: 1 })
    const insertIndex = pageIndex + 1
    const captured = capturePageMutation()
    await mutate({ type: 'blank', index: insertIndex, width: baseViewport.width, height: baseViewport.height }, 'Adding blank page…', () => {
      finishPageMutation(captured.overlays.map((overlay) => ({ ...overlay, pageIndex: overlay.pageIndex >= insertIndex ? overlay.pageIndex + 1 : overlay.pageIndex })))
      setPageRotations((current) => Object.fromEntries(Object.entries(current).map(([rawIndex, amount]) => {
        const index = Number(rawIndex)
        return [index >= insertIndex ? index + 1 : index, amount]
      })))
      setBookmarks((current) => remapBookmarkPages(current, (index) => index >= insertIndex ? index + 1 : index))
      setPageIndex(insertIndex)
      setSelectedPages(new Set([insertIndex]))
      showToast('Blank page added')
    }, captured.snapshot)
  }

  async function deletePages() {
    if (!pdf) return
    if (!(await confirmFormSensitiveOperation('Deleting pages'))) return
    const indices = targetPages()
    if (indices.length >= pdf.numPages) {
      showToast('A PDF must keep at least one page.')
      return
    }
    if (!window.confirm(`Delete ${indices.length === 1 ? `page ${indices[0] + 1}` : `${indices.length} selected pages`}? You can undo this action.`)) return
    const captured = capturePageMutation()
    await mutate({ type: 'delete', indices }, `Deleting ${indices.length === 1 ? 'page' : 'pages'}…`, () => {
      finishPageMutation(captured.overlays.flatMap((overlay) => {
        const mapped = remapAfterDelete(overlay.pageIndex, indices)
        return mapped === null ? [] : [{ ...overlay, pageIndex: mapped }]
      }))
      // Mirrors how the main process pruned the document's outline.
      setBookmarks((current) => bookmarksAfterPageDelete(current, indices))
      setPageRotations((current) => Object.fromEntries(Object.entries(current).flatMap(([rawIndex, amount]) => {
        const mapped = remapAfterDelete(Number(rawIndex), indices)
        return mapped === null ? [] : [[mapped, amount]]
      })))
      const nextPage = clamp(pageIndex - indices.filter((index) => index < pageIndex).length, 0, pdf.numPages - indices.length - 1)
      setPageIndex(nextPage)
      setSelectedPages(new Set([nextPage]))
      showToast(`${indices.length === 1 ? 'Page' : `${indices.length} pages`} deleted`, 'Undo', () => historyActionsRef.current.undo())
    }, captured.snapshot)
  }

  async function reorderPages(indices: number[], insertIndex: number) {
    if (!pdf || !indices.length) return
    const moving = [...new Set(indices)].filter((index) => index >= 0 && index < pdf.numPages).sort((a, b) => a - b)
    const remaining = Array.from({ length: pdf.numPages }, (_, index) => index).filter((index) => !moving.includes(index))
    const position = Math.max(0, Math.min(remaining.length, insertIndex - moving.filter((index) => index < insertIndex).length))
    const order = [...remaining.slice(0, position), ...moving, ...remaining.slice(position)]
    if (order.every((value, index) => value === index)) return
    if (!(await confirmFormSensitiveOperation('Reordering pages'))) return
    const captured = capturePageMutation()
    await mutate({ type: 'reorder', order }, 'Reordering pages…', () => {
      const mapIndex = (oldIndex: number) => order.indexOf(oldIndex)
      finishPageMutation(captured.overlays.map((overlay) => ({ ...overlay, pageIndex: mapIndex(overlay.pageIndex) })))
      setPageRotations((current) => Object.fromEntries(Object.entries(current).map(([rawIndex, amount]) => [mapIndex(Number(rawIndex)), amount])))
      setBookmarks((current) => remapBookmarkPages(current, mapIndex))
      const moved = moving.map(mapIndex).sort((a, b) => a - b)
      setPageIndex(moved[0])
      setSelectedPages(new Set(moved))
      lastSelectedPage.current = moved.at(-1) ?? moved[0]
      showToast(moved.length === 1 ? `Page moved to position ${moved[0] + 1}` : `${moved.length} pages moved to position ${moved[0] + 1}`, 'Undo', () => historyActionsRef.current.undo())
    }, captured.snapshot)
  }

  async function addPages() {
    if (!bytes || !pdf) return
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    if (!(await confirmFormSensitiveOperation('Adding pages'))) return
    const insertIndex = pageIndex + 1
    const { overlays: sourceOverlays, snapshot } = capturePageMutation()
    documentOperationRef.current = true
    setDocumentUpdating(true)
    let awaitingSwap = false
    setBusy('Adding pages…')
    try {
      const result = await window.simple.insertFiles(bytes, insertIndex)
      if (!result) return
      pendingDocumentSwapRef.current = {
        commit: () => {
          if (snapshot) addUndoSnapshot(snapshot)
          setOverlays(sourceOverlays.map((overlay) => ({ ...overlay, pageIndex: overlay.pageIndex >= insertIndex ? overlay.pageIndex + result.added : overlay.pageIndex })))
          setPageRotations((current) => Object.fromEntries(Object.entries(current).map(([rawIndex, amount]) => {
            const index = Number(rawIndex)
            return [index >= insertIndex ? index + result.added : index, amount]
          })))
          setBookmarks((current) => remapBookmarkPages(current, (index) => index >= insertIndex ? index + result.added : index))
          setPageIndex(insertIndex)
          const insertedPages = Array.from({ length: result.added }, (_, offset) => insertIndex + offset)
          setSelectedPages(new Set(insertedPages))
          lastSelectedPage.current = insertedPages.at(-1) ?? insertIndex
          setTextEdit(null)
          setObjectEdit(null)
          setDirty(true)
          showToast(`${result.added} ${result.added === 1 ? 'page' : 'pages'} added`)
        },
        rollback: () => rollbackDocumentBytes(bytes),
      }
      awaitingDocumentSwapRef.current = true
      awaitingSwap = true
      setBytes(normalizeBytes(result.data))
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      if (!awaitingSwap) {
        documentOperationRef.current = false
        setDocumentUpdating(false)
      }
      setBusy('')
    }
  }

  async function importPagesAt(files: File[], insertIndex: number) {
    if (!bytes || !pdf || !files.length) return
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    if (!(await confirmFormSensitiveOperation('Adding pages'))) return
    const { overlays: sourceOverlays, snapshot } = capturePageMutation()
    documentOperationRef.current = true
    setDocumentUpdating(true)
    let awaitingSwap = false
    setBusy('Adding dropped pages…')
    try {
      const inputs = await Promise.all(files.map(async (file) => isConvertibleImage(file)
        ? convertImageToPng(file)
        : { name: file.name, data: await file.arrayBuffer() }))
      const result = await window.simple.insertDroppedFiles(bytes, insertIndex, inputs)
      pendingDocumentSwapRef.current = {
        commit: () => {
          if (snapshot) addUndoSnapshot(snapshot)
          setOverlays(sourceOverlays.map((overlay) => ({
            ...overlay,
            pageIndex: overlay.pageIndex >= insertIndex ? overlay.pageIndex + result.added : overlay.pageIndex,
          })))
          setPageRotations((current) => Object.fromEntries(Object.entries(current).map(([rawIndex, amount]) => {
            const index = Number(rawIndex)
            return [index >= insertIndex ? index + result.added : index, amount]
          })))
          setBookmarks((current) => remapBookmarkPages(current, (index) => index >= insertIndex ? index + result.added : index))
          setPageIndex(insertIndex)
          const insertedPages = Array.from({ length: result.added }, (_, offset) => insertIndex + offset)
          setSelectedPages(new Set(insertedPages))
          lastSelectedPage.current = insertedPages.at(-1) ?? insertIndex
          setTextEdit(null)
          setObjectEdit(null)
          setDirty(true)
          showToast(`${result.added} ${result.added === 1 ? 'page' : 'pages'} added where dropped`)
        },
        rollback: () => rollbackDocumentBytes(bytes),
      }
      awaitingDocumentSwapRef.current = true
      awaitingSwap = true
      setBytes(normalizeBytes(result.data))
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      if (!awaitingSwap) {
        documentOperationRef.current = false
        setDocumentUpdating(false)
      }
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
    if (!(await confirmFormSensitiveOperation('Exporting pages'))) return
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
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    if (!(await confirmFormSensitiveOperation('Exporting pages'))) return
    const indices = targetPages()
    documentOperationRef.current = true
    setBusy('Preparing export…')
    try {
      const output = await preparedBytes()
      const suggested = `${withoutExtension(documentFile.name)} — ${indices.length === 1 ? `page ${indices[0] + 1}` : `${indices.length} pages`}.pdf`
      const exportPath = await window.simple.exportPages(output, indices, suggested)
      if (exportPath) showToast('Pages exported', 'Show file', () => window.simple.showItem(exportPath))
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      documentOperationRef.current = false
      setBusy('')
    }
  }

  async function runExportAs(job: PdfExportSubmission, options: { skipScanCheck?: boolean } = {}) {
    if (!bytes || !documentFile || !pdf) return
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    const textExport = job.format === 'docx' || job.format === 'txt' || job.format === 'md' || job.format === 'html'
    if (textExport && !options.skipScanCheck && !passwordProtected) {
      // Scanned pages would export empty: offer to recognise their text first.
      setExportDialogOpen(false)
      documentOperationRef.current = true
      setBusy('Checking for scanned pages…')
      let scanned: number[] = []
      try {
        scanned = await scannedPagesIn(pdf, job.pageIndices)
      } catch {
        scanned = []
      } finally {
        documentOperationRef.current = false
        setBusy('')
      }
      if (scanned.length) {
        setOcrExportPrompt({ job, pages: scanned })
        return
      }
    }
    const fullDocument = job.pageIndices.length === pdf.numPages
      && job.pageIndices.every((index, position) => index === position)
    if (job.format === 'pdf' && !fullDocument && !(await confirmFormSensitiveOperation('Exporting a page subset'))) return
    const pendingEdits = Boolean(dirty || textEdit?.modified || objectEdit?.modified)
    if (
      job.format === 'pdf'
      && documentFile.signatureDetected
      && (pendingEdits || !fullDocument)
      && !window.confirm('Exporting this PDF can invalidate its digital signature. Continue with an unsigned exported copy?')
    ) return

    documentOperationRef.current = true
    setExportDialogOpen(false)
    setBusy('Preparing export…')
    let imageSessionId = ''
    let exportPdf: PDFDocumentProxy | null = null
    let ownsExportPdf = false
    try {
      const output = await preparedBytes()
      const baseName = withoutExtension(documentFile.name)
      const scopeSuffix = fullDocument
        ? ''
        : job.pageIndices.length === 1
          ? ` - page ${job.pageIndices[0] + 1}`
          : ' - selected pages'
      let exportedPath: string | null = null

      if (job.format === 'pdf') {
        exportedPath = await window.simple.exportAsPdf({
          data: output,
          indices: job.pageIndices,
          fullDocument,
          suggestedName: `${baseName}${scopeSuffix}.pdf`,
        })
      } else {
        if (output === bytes) {
          exportPdf = pdf
        } else {
          setBusy('Preparing converted pages…')
          exportPdf = await loadPdf(output)
          ownsExportPdf = true
        }

        if (job.format === 'png' || job.format === 'jpeg' || job.format === 'webp') {
          const session = await window.simple.beginImageExport({
            format: job.format,
            pageCount: job.pageIndices.length,
            baseName: `${baseName}${scopeSuffix}`,
          })
          if (!session) return
          imageSessionId = session.id
          for (let position = 0; position < job.pageIndices.length; position += 1) {
            const targetPageIndex = job.pageIndices[position]
            setBusy(`Rendering page ${position + 1} of ${job.pageIndices.length}…`)
            const data = await renderPdfPageImage(exportPdf, targetPageIndex, job.format, job.imageScale, job.jpegQuality)
            await window.simple.writeImageExportPage({ id: session.id, pageNumber: targetPageIndex + 1, data })
          }
          exportedPath = await window.simple.finishImageExport(session.id)
          imageSessionId = ''
        } else {
          setBusy('Extracting document text…')
          const pages = await extractPdfText(exportPdf, job.pageIndices)
          exportedPath = await window.simple.exportTextDocument({
            format: job.format,
            pages,
            title: baseName,
            baseName: `${baseName}${scopeSuffix}`,
          })
        }
      }

      if (exportedPath) showToast('Export completed', 'Show file', () => window.simple.showItem(exportedPath as string))
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      if (imageSessionId) await window.simple.cancelImageExport(imageSessionId).catch(() => false)
      if (ownsExportPdf && exportPdf) await exportPdf.destroy().catch(() => {})
      documentOperationRef.current = false
      setBusy('')
    }
  }

  const runExportAsRef = useRef(runExportAs)
  runExportAsRef.current = runExportAs

  // A text export waiting for its scanned pages: continue once the recognised text is in place.
  useEffect(() => {
    const job = pendingExportRef.current
    if (!job || busy || documentUpdating || documentLoading || documentOperationRef.current) return
    pendingExportRef.current = null
    void runExportAsRef.current(job, { skipScanCheck: true })
  }, [busy, documentUpdating, documentLoading, pdf])

  function closePrintDialog() {
    setPrintDialogOpen(false)
    setPrintPreviewBytes(null)
    setPrintPreviewPdf(null)
  }

  async function openPrintDialog() {
    if (!bytes || !documentFile || !pdf) return
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    documentOperationRef.current = true
    setBusy('Preparing print preview...')
    try {
      const output = await preparedBytes()
      const preview = await loadPdf(output)
      setPrintPreviewBytes(output)
      setPrintPreviewPdf(preview)
      setPrintDialogOpen(true)
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      documentOperationRef.current = false
      setBusy('')
    }
  }

  async function runPrintJob(job: PrintDialogSubmission) {
    if (!bytes || !documentFile || !pdf) return
    if (documentOperationRef.current) {
      showToast('Please wait for the current document operation to finish.')
      return
    }
    documentOperationRef.current = true
    setBusy('Printing…')
    try {
      const output = printPreviewBytes || await preparedBytes()
      const result = await window.simple.printPdfDirect(output, documentFile.name, {
        deviceName: job.deviceName,
        copies: job.copies,
        ...(job.pageIndices ? { pageIndices: job.pageIndices } : {}),
        landscape: job.landscape,
        color: job.color,
        ...(job.duplexMode ? { duplexMode: job.duplexMode } : {}),
        collate: job.collate,
        paperSize: job.paperSize,
        marginMode: job.marginMode,
        scaleMode: job.scaleMode,
        customScale: job.customScale,
      })
      if (result.success) {
        closePrintDialog()
        showToast(`Sent to ${job.printerLabel}`)
      } else {
        showToast(result.failureReason ? `Print failed — ${result.failureReason}` : 'The printer did not accept the job')
      }
    } catch (error) {
      showToast(errorMessage(error))
    } finally {
      documentOperationRef.current = false
      setBusy('')
    }
  }

  async function cropPage(rect: PdfRect, targetPageIndex = pageIndex) {
    const captured = capturePageMutation()
    await mutate(
      { type: 'crop', pageIndex: targetPageIndex, rect },
      'Cropping page…',
      () => {
        finishPageMutation(captured.overlays)
        setTool('select')
        showToast('Crop applied — save to keep it')
      },
      captured.snapshot,
    )
  }

  function addHighlights(rects: PdfRect[], targetPageIndex = pageIndex) {
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
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
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
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
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
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
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
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

  // Typing into one text field is one undoable step until another field (or
  // anything else) changes or the typing pauses; each choice is its own step.
  function changeFormValue(name: string, value: PdfFormValue) {
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    const now = performance.now()
    const last = formEditRef.current
    if (typeof value === 'string' && last?.name === name && now - last.at < 1500) {
      // Still the same step, but a different document state.
      revisionRef.current = nextRevision()
    } else {
      const snapshot = currentSnapshot()
      if (snapshot) addUndoSnapshot(snapshot)
    }
    formEditRef.current = typeof value === 'string' ? { name, at: now } : null
    setFormValues((current) => ({ ...current, [name]: value }))
    setDirty(true)
  }

  function commitTextEditValue(edit: PageTextEdit, announce = true) {
    resetPendingEditHistory()
    if (passwordProtected) {
      setTextEdit(null)
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
    // An unchanged selection, or an added box left empty, changes nothing.
    if (!edit.modified || (isEmptyAddedText(edit) && !edit.overlayId)) {
      setTextEdit(null)
      return
    }
    // Scanned text typed back to what the page says: nothing to add (an
    // earlier edit of the line is taken back, as an undoable step).
    const unchangedScan = Boolean(edit.scan) && finalizeScanEdit(edit).kind === 'none'
    if (unchangedScan && !edit.overlayId) {
      setTextEdit(null)
      return
    }
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setOverlays((current) => overlaysWithPendingEdits(current, edit, null))
    setTextEdit(null)
    setDirty(true)
    if (announce) showToast(unchangedScan ? 'Scanned text restored' : edit.text ? (edit.cover ? 'Text updated' : 'Text added') : 'Text removed')
  }

  function commitTextEdit() {
    if (textEdit) commitTextEditValue(textEdit)
  }

  /**
   * An edit of scanned text finished preparing: its retouch patch, word
   * boxes and matched style go into the edit on the page and into an edit of
   * that line committed meanwhile (not an undo step: it completes the edit).
   */
  function mergeScanPreparation(key: string, result: ScanPreparation) {
    setTextEdit((current) => current?.scan?.key === key ? applyScanPreparation(current, result) : current)
    setOverlays((current) => {
      if (!current.some((overlay) => overlay.type === 'text' && overlay.scan?.key === key && overlay.scan.status === 'pending')) return current
      return current.map((overlay) => {
        if (overlay.type !== 'text' || overlay.scan?.key !== key || overlay.scan.status !== 'pending') return overlay
        const merged = applyScanPreparation(overlay, result)
        const finalized = finalizeScanEdit({ ...merged, originalText: merged.originalText ?? '', modified: true })
        return finalized.kind === 'edit'
          ? { ...merged, rect: finalized.edit.rect, baselineOffset: finalized.edit.baselineOffset, scan: finalized.edit.scan } as TextOverlay
          : merged
      })
    })
  }

  function commitObjectEditValue(edit: PageObjectEdit, announce = true) {
    resetPendingEditHistory()
    if (passwordProtected) {
      setObjectEdit(null)
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
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

  /**
   * Apply whatever is selected on the page (typed text, a moved image, a
   * placed signature) as one undoable step, the way clicking elsewhere does.
   * Escape, Save and page changes use this, so an edit is never thrown away.
   */
  function commitPendingEdits(announce = false) {
    editSelectionGenerationRef.current += 1
    if (textEdit) commitTextEditValue(textEdit, announce)
    if (objectEdit) commitObjectEditValue(objectEdit, announce)
    setSelectingObjectRegion(false)
  }

  function beginTextEdit(edit: PageTextEdit) {
    editSelectionGenerationRef.current += 1
    if (textEdit && textEdit !== edit) commitTextEditValue(textEdit, false)
    if (objectEdit) commitObjectEditValue(objectEdit, false)
    resetPendingEditHistory()
    setObjectEdit(null)
    setSelectingObjectRegion(false)
    setPageIndex(edit.pageIndex)
    setSelectedPages(new Set([edit.pageIndex]))
    setTextEdit(edit)
    setTool('edit')
  }

  function beginObjectEdit(edit: PageObjectEdit) {
    editSelectionGenerationRef.current += 1
    if (textEdit) commitTextEditValue(textEdit, false)
    if (objectEdit && objectEdit !== edit) commitObjectEditValue(objectEdit, false)
    resetPendingEditHistory()
    setTextEdit(null)
    setSelectingObjectRegion(false)
    setPageIndex(edit.pageIndex)
    setSelectedPages(new Set([edit.pageIndex]))
    setObjectEdit(edit)
    setTool('edit')
  }

  /**
   * The inspector's Cancel, the one explicit way to discard the selected
   * edit. Even that can be taken back: the discarded edit waits on Redo.
   */
  function cancelEditSelection() {
    const before = textEdit?.modified || objectEdit?.modified ? currentSnapshot() : null
    if (before) {
      const applied: Snapshot = {
        ...before,
        overlays: overlaysWithPendingEdits(before.overlays, textEdit, objectEdit).map(cloneOverlay),
        revision: nextRevision(),
      }
      redoStackRef.current = [applied]
      setRedoStack([applied])
    }
    editSelectionGenerationRef.current += 1
    resetPendingEditHistory()
    setTextEdit(null)
    setObjectEdit(null)
    setSelectingObjectRegion(false)
    if (before) showToast('Edit discarded', 'Undo', () => historyActionsRef.current.redo())
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
      changeObjectEditStep({
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
      changeObjectEditStep({
        ...objectEdit,
        dataUrl: rotated,
        rect: { x: centerX - width / 2, y: centerY - height / 2, width, height },
        modified: true,
      })
    } catch (error) {
      showToast(errorMessage(error))
    }
  }

  async function flipObject(direction: 'horizontal' | 'vertical') {
    if (!objectEdit?.dataUrl) return
    try {
      const sideways = objectEdit.displayRotation === 90 || objectEdit.displayRotation === 270
      const storedDirection = sideways
        ? (direction === 'horizontal' ? 'vertical' : 'horizontal')
        : direction
      changeObjectEditStep({
        ...objectEdit,
        dataUrl: await flipImageData(objectEdit.dataUrl, storedDirection),
        modified: true,
      })
    } catch (error) {
      showToast(errorMessage(error))
    }
  }

  function duplicateObject() {
    if (!objectEdit?.dataUrl) return
    // Duplicate pastes a private copy and leaves the clipboard alone.
    const payload = copyPageEditSelection(null, objectEdit)
    if (payload) void pasteCurrentEditClipboard(payload)
  }

  function copyCurrentEdit(removeAfterCopy = false) {
    const payload = copyPageEditSelection(textEdit, objectEdit)
    if (!payload) return false
    editClipboardRef.current = payload
    setHasEditClipboard(true)
    const plainText = editClipboardPlainText(payload)
    editClipboardTextRef.current = plainText
    void navigator.clipboard?.writeText(plainText).catch(() => {})
    if (removeAfterCopy) deleteEditSelection()
    else showToast(`${payload.kind === 'text' ? 'Text box' : payload.edit.label} copied`)
    return true
  }

  async function pasteCurrentEditClipboard(source: EditClipboardPayload | null = editClipboardRef.current) {
    if (!pdf || !source) return false
    try {
      const targetPage = await pdf.getPage(pageIndex + 1)
      const view = targetPage.view
      const pageBounds: PdfRect = {
        x: Math.min(view[0], view[2]),
        y: Math.min(view[1], view[3]),
        width: Math.abs(view[2] - view[0]),
        height: Math.abs(view[3] - view[1]),
      }
      const payload = source
      const sourceRotation = (((Number(payload.edit.displayRotation) || 0) % 360 + 360) % 360) as DisplayRotation
      const targetRotation = (((Number(targetPage.rotate || 0) + (pageRotations[pageIndex] || 0)) % 360 + 360) % 360) as DisplayRotation
      const pasted = pasteEditToPage(payload, pageIndex, pageBounds, 0)
      const orientedRect = reorientPastedRect(pasted.edit.rect, pageBounds, sourceRotation, targetRotation)
      const targetViewport = targetPage.getViewport({ scale: 1, rotation: targetRotation })
      const displayedRect = pdfRectToViewport(targetViewport, orientedRect)
      const left = clamp(displayedRect.left + 12, 0, Math.max(0, targetViewport.width - displayedRect.width))
      const top = clamp(displayedRect.top + 12, 0, Math.max(0, targetViewport.height - displayedRect.height))
      const rect = viewportRectToPdf(targetViewport, {
        left,
        top,
        right: left + displayedRect.width,
        bottom: top + displayedRect.height,
      })
      if (pasted.kind === 'text') {
        beginTextEdit({ ...pasted.edit, rect, displayRotation: targetRotation })
      } else {
        const rotationDelta = (((targetRotation - sourceRotation) % 360 + 360) % 360) as DisplayRotation
        beginObjectEdit({
          ...pasted.edit,
          rect,
          displayRotation: targetRotation,
          dataUrl: pasted.edit.dataUrl
            ? await normalizeImageOrientation(pasted.edit.dataUrl, rotationDelta)
            : undefined,
        })
      }
      showToast(`${pasted.kind === 'text' ? 'Text box' : pasted.edit.label} pasted`)
      return true
    } catch (error) {
      showToast(errorMessage(error))
      return false
    }
  }

  function duplicateText() {
    if (!textEdit) return
    const payload = copyPageEditSelection(textEdit, null)
    if (payload) void pasteCurrentEditClipboard(payload)
  }

  // What other apps last put on the system clipboard. An unreadable clipboard
  // reports systemReadable false so the in-app payload still pastes.
  async function readSystemClipboard() {
    let systemReadable = false
    let systemText = ''
    let imageDataUrl: string | null = null
    try {
      systemText = await navigator.clipboard.readText()
      systemReadable = true
    } catch {
      // Clipboard access denied or unavailable.
    }
    if (!systemText.trim()) {
      try {
        const items = await navigator.clipboard.read()
        systemReadable = true
        for (const item of items) {
          const type = item.types.find((candidate) => candidate.startsWith('image/'))
          if (!type) continue
          imageDataUrl = await blobToDataUrl(await item.getType(type))
          break
        }
      } catch {
        // No image, or the format cannot be read.
      }
    }
    return { systemReadable, systemText, imageDataUrl }
  }

  // Centre of the visible part of a page, as fractions of its displayed size,
  // so pasted content lands where the user is looking.
  function visiblePageCenter(index: number) {
    const viewer = viewerRef.current
    const surface = viewer?.querySelector<HTMLElement>(`.continuous-page-slot[data-page-index="${index}"] .page-surface`)
    if (!viewer || !surface) return { x: 0.5, y: 0.5 }
    const view = viewer.getBoundingClientRect()
    const page = surface.getBoundingClientRect()
    const left = Math.max(view.left, page.left)
    const right = Math.min(view.right, page.right)
    const top = Math.max(view.top, page.top)
    const bottom = Math.min(view.bottom, page.bottom)
    if (page.width < 1 || page.height < 1 || right <= left || bottom <= top) return { x: 0.5, y: 0.5 }
    return { x: ((left + right) / 2 - page.left) / page.width, y: ((top + bottom) / 2 - page.top) / page.height }
  }

  async function pasteSystemText(rawText: string) {
    if (!pdf) return false
    const text = rawText.replace(/\r\n?/g, '\n').replace(/\s+$/u, '')
    if (!text) return false
    try {
      const page = await pdf.getPage(pageIndex + 1)
      const displayRotation = ((((page.rotate || 0) + (pageRotations[pageIndex] || 0)) % 360 + 360) % 360) as DisplayRotation
      const viewport = page.getViewport({ scale: 1, rotation: displayRotation })
      const fontSize = 12
      const box = pastedTextBoxSize(text, fontSize, viewport.width * 0.8, viewport.height * 0.9)
      const center = visiblePageCenter(pageIndex)
      const left = clamp(center.x * viewport.width - box.width / 2, 0, Math.max(0, viewport.width - box.width))
      const top = clamp(center.y * viewport.height - box.height / 2, 0, Math.max(0, viewport.height - box.height))
      beginTextEdit({
        pageIndex,
        rect: viewportRectToPdf(viewport, { left, top, right: left + box.width, bottom: top + box.height }),
        originalText: '',
        text,
        fontSize,
        fontFamily: 'Segoe UI',
        textFit: box.wraps ? 'wrap' : undefined,
        displayRotation,
        align: 'left',
        color: [0.04, 0.04, 0.05],
        cover: false,
        modified: true,
        caretOffset: text.length,
        selectionStart: text.length,
        selectionEnd: text.length,
      })
      showToast('Text pasted')
      return true
    } catch (error) {
      showToast(errorMessage(error))
      return false
    }
  }

  async function pasteSystemImage(dataUrl: string) {
    if (!pdf) return false
    try {
      const dimensions = await imageDimensions(dataUrl)
      const page = await pdf.getPage(pageIndex + 1)
      const displayRotation = ((((page.rotate || 0) + (pageRotations[pageIndex] || 0)) % 360 + 360) % 360) as DisplayRotation
      const viewport = page.getViewport({ scale: 1, rotation: displayRotation })
      const scale = Math.min(1, viewport.width * 0.48 / dimensions.width, viewport.height * 0.48 / dimensions.height)
      const width = Math.max(36, dimensions.width * scale)
      const height = Math.max(36, dimensions.height * scale)
      const center = visiblePageCenter(pageIndex)
      const left = clamp(center.x * viewport.width - width / 2, 0, Math.max(0, viewport.width - width))
      const top = clamp(center.y * viewport.height - height / 2, 0, Math.max(0, viewport.height - height))
      beginObjectEdit({
        pageIndex,
        kind: 'image',
        rect: viewportRectToPdf(viewport, { left, top, right: left + width, bottom: top + height }),
        dataUrl: await normalizeImageOrientation(dataUrl, displayRotation),
        opacity: 1,
        cover: false,
        displayRotation,
        label: 'Pasted image',
        modified: true,
      })
      showToast('Image pasted')
      return true
    } catch (error) {
      showToast(errorMessage(error))
      return false
    }
  }

  // Ctrl+V and the Paste commands: the in-app text box or image only while the
  // system clipboard still holds what Simple copied, otherwise the newer
  // system text or image.
  async function pasteFromClipboard() {
    if (!pdf || passwordProtected) return false
    const internal = editClipboardRef.current
    const system = await readSystemClipboard()
    const source = chooseEditPasteSource({
      internalText: internal ? (editClipboardTextRef.current ?? editClipboardPlainText(internal)) : null,
      systemReadable: system.systemReadable,
      systemText: system.systemText,
      systemHasImage: Boolean(system.imageDataUrl),
    })
    if (source === 'internal') return pasteCurrentEditClipboard(internal)
    if (source === 'image' && system.imageDataUrl) return pasteSystemImage(system.imageDataUrl)
    if (source === 'text') return pasteSystemText(system.systemText)
    return false
  }

  const hasContextEditSelection = Boolean(textEdit || objectEdit)
  const canContextUndo = Boolean(textEdit?.modified || objectEdit?.modified || undoStack.length)
  const canContextRedo = Boolean((redoStack.length && !textEdit?.modified && !objectEdit?.modified)
    || ((textEdit || objectEdit) && pendingEditHistoryRef.current.future.length))
  const canContextPaste = Boolean(hasEditClipboard && !passwordProtected)
  const hasContextCommands = canContextUndo || canContextRedo || hasContextEditSelection || canContextPaste

  function closeContextMenu(restoreFocus = false) {
    setContextMenu(null)
    const returnTarget = contextMenuFocusReturnRef.current
    contextMenuFocusReturnRef.current = null
    if (restoreFocus && returnTarget?.isConnected) {
      window.requestAnimationFrame(() => returnTarget.focus({ preventScroll: true }))
    }
  }

  function runContextMenuAction(action: () => unknown | Promise<unknown>, restoreFocus = true) {
    closeContextMenu(restoreFocus)
    void action()
  }

  function handleDocumentContextMenu(event: ReactMouseEvent<HTMLDivElement>) {
    const target = event.target
    if (!(target instanceof Element) || !target.closest('.viewer, .edit-inspector')) {
      closeContextMenu(false)
      return
    }

    // Chromium and Electron already provide the correct caret-aware commands
    // for real fields and DOM text selections. In particular, routing a
    // textarea selection through copyCurrentEdit would copy the complete PDF
    // text box instead of the characters the user selected.
    if (target.closest('input, textarea, select, [contenteditable="true"]')) {
      closeContextMenu(false)
      return
    }
    const nativeSelection = window.getSelection()
    if (nativeSelection && !nativeSelection.isCollapsed && nativeSelection.toString()) {
      closeContextMenu(false)
      return
    }

    // Suppress the generic workspace menu even when no semantic command is
    // currently available. Showing enabled Cut/Delete items against a blank
    // page can otherwise act on stale sidebar focus.
    event.preventDefault()
    event.stopPropagation()
    if (
      !hasContextCommands
      || documentUpdating
      || documentOperationRef.current
      || printDialogOpen
      || exportDialogOpen
      || busy
    ) {
      closeContextMenu(false)
      return
    }

    contextMenuFocusReturnRef.current = document.activeElement instanceof HTMLElement
      ? document.activeElement
      : null
    setContextMenu({ x: event.clientX, y: event.clientY })
  }

  useLayoutEffect(() => {
    if (!contextMenu || !contextMenuRef.current) return
    const menu = contextMenuRef.current
    const bounds = menu.getBoundingClientRect()
    menu.style.left = `${Math.max(4, Math.min(contextMenu.x, window.innerWidth - bounds.width - 4))}px`
    menu.style.top = `${Math.max(4, Math.min(contextMenu.y, window.innerHeight - bounds.height - 4))}px`
    menu.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus({ preventScroll: true })
  }, [contextMenu])

  useEffect(() => {
    if (!contextMenu) return
    const closeOutside = (event: globalThis.PointerEvent) => {
      if (!contextMenuRef.current?.contains(event.target as Node)) closeContextMenu(false)
    }
    const closeWithoutFocus = () => closeContextMenu(false)
    const viewer = viewerRef.current
    document.addEventListener('pointerdown', closeOutside, true)
    window.addEventListener('blur', closeWithoutFocus)
    window.addEventListener('resize', closeWithoutFocus)
    viewer?.addEventListener('scroll', closeWithoutFocus, { passive: true })
    return () => {
      document.removeEventListener('pointerdown', closeOutside, true)
      window.removeEventListener('blur', closeWithoutFocus)
      window.removeEventListener('resize', closeWithoutFocus)
      viewer?.removeEventListener('scroll', closeWithoutFocus)
    }
  }, [contextMenu])

  function bookmarksEditable() {
    if (passwordProtected) {
      showToast('Password-protected PDFs are read-only in this version.')
      return false
    }
    if (bookmarksReadOnly) {
      // Saving an edited copy of an outline that could not be read would
      // replace the bookmarks the file really has.
      showToast('This PDF’s bookmarks could not be read, so they cannot be changed.')
      return false
    }
    return true
  }

  // The toolbar button adds or removes its own "Page N" bookmark; a chapter
  // of the document's outline on the same page is never removed by it.
  function toggleBookmark() {
    if (!bookmarksEditable()) return
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    const existing = toolbarBookmarkFor(bookmarks, pageIndex)
    if (existing) {
      setBookmarks((current) => current.filter((bookmark) => bookmark.id !== existing.id))
      showToast('Bookmark removed', 'Undo', () => historyActionsRef.current.undo())
    } else {
      setBookmarks((current) => [...current, { id: makeId('bookmark'), pageIndex, label: `Page ${pageIndex + 1}`, depth: 0, source: 'simple' }])
      showToast('Bookmark added — save to keep it in the PDF')
    }
    setBookmarksDirty(true)
    setDirty(true)
  }

  // Removing a bookmark removes the bookmarks nested under it too.
  function deleteBookmark(id: string) {
    if (!bookmarksEditable()) return
    const removed = bookmarkSubtreeSize(bookmarks, id)
    if (!removed) return
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setBookmarks((current) => withoutBookmark(current, id))
    setBookmarksDirty(true)
    setDirty(true)
    showToast(removed > 1 ? `Bookmark and ${removed - 1} nested ${removed === 2 ? 'bookmark' : 'bookmarks'} removed` : 'Bookmark removed', 'Undo', () => historyActionsRef.current.undo())
  }

  function renameBookmark(id: string, label: string) {
    const title = label.trim()
    const current = bookmarks.find((bookmark) => bookmark.id === id)
    if (!title || !current || current.label === title) return
    if (!bookmarksEditable()) return
    const snapshot = currentSnapshot()
    if (snapshot) addUndoSnapshot(snapshot)
    setBookmarks((list) => list.map((bookmark) => bookmark.id === id ? { ...bookmark, label: title } : bookmark))
    setBookmarksDirty(true)
    setDirty(true)
  }

  // Bookmarks that do not go to a page: a web link opens in the browser
  // (http, https and mailto only); a heading only groups the ones below it.
  function openBookmark(bookmark: Bookmark) {
    if (typeof bookmark.pageIndex === 'number') {
      goToPage(bookmark.pageIndex)
      return
    }
    if (bookmark.url) {
      window.simple.openExternal(bookmark.url).catch((error) => showToast(errorMessage(error)))
    }
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

  // PageDown/PageUp in continuous view move one screen at a time, as Acrobat
  // does, and only step to the neighbouring page (shown from its top) at a
  // page boundary, so keyboard reading never skips part of a tall page.
  function scrollByScreen(direction: 1 | -1) {
    if (!pdf) return
    const viewer = viewerRef.current
    const slotFor = (index: number) => viewer?.querySelector<HTMLElement>(`.continuous-page-slot[data-page-index="${index}"]`)
    const slot = slotFor(pageIndex)
    if (!viewer || !slot) {
      goToPage(pageIndex + direction)
      return
    }
    const view = viewer.getBoundingClientRect()
    const page = slot.getBoundingClientRect()
    const step = Math.max(40, viewer.clientHeight - 40)
    if (direction > 0) {
      const remaining = page.bottom - view.bottom
      if (remaining > 2) viewer.scrollTop += Math.min(step, remaining)
      else if (pageIndex < pdf.numPages - 1) goToPage(pageIndex + 1)
      return
    }
    const hidden = view.top - page.top
    if (hidden > 2) {
      viewer.scrollTop -= Math.min(step, hidden)
      return
    }
    if (pageIndex <= 0) {
      viewer.scrollTop = 0
      return
    }
    const previous = slotFor(pageIndex - 1)?.getBoundingClientRect()
    // A previous page taller than the window is entered from its bottom.
    const delta = previous ? view.bottom - previous.bottom : 0
    if (previous && previous.height > viewer.clientHeight && delta > 2) viewer.scrollTop -= Math.min(step, delta)
    else goToPage(pageIndex - 1)
  }

  function selectAllDocumentText() {
    const pages = viewerRef.current?.querySelector('.continuous-pages')
    const selection = window.getSelection()
    if (!pages || !selection) return
    selection.selectAllChildren(pages)
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

  const fitView = useCallback(async (mode: 'fit' | 'width' = zoomMode === 'width' ? 'width' : 'fit') => {
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
      const next = fittedZoom(viewer, viewport, mode)
      setZoom((current) => Math.abs(current - next) < 0.002 ? current : next)
    } catch (error) {
      if (requestGeneration !== fitRequestGenerationRef.current) return
      showToast(errorMessage(error))
    }
  }, [pdf, pageIndex, pageRotations, zoomMode, showToast])

  useEffect(() => {
    if (!pdf || zoomMode === 'custom') return
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
    fitRequestGenerationRef.current += 1
    setZoomMode('custom')
    setZoom(next)
  }

  function changeFit(mode: 'fit' | 'width' | 'actual' | 'custom') {
    if (mode === 'actual') { changeZoom(1); return }
    if (mode === 'custom') return
    setZoomMode(mode)
    void fitView(mode)
  }

  // Work in displayed page coordinates so the arrows stay intuitive on rotated
  // and cropped pages. Functional updates preserve repeated keystrokes.
  async function nudgeEdit(key: string, distance: number) {
    const selected = textEdit || objectEdit
    if (!pdf || !selected) return
    const generation = editSelectionGenerationRef.current
    recordPendingEditStep()
    try {
      const page = await pdf.getPage(selected.pageIndex + 1)
      const rotation = (((page.rotate || 0) + (pageRotations[selected.pageIndex] || 0)) % 360 + 360) % 360
      const viewport = page.getViewport({ scale: 1, rotation })
      const dx = key === 'ArrowRight' ? distance : key === 'ArrowLeft' ? -distance : 0
      const dy = key === 'ArrowDown' ? distance : key === 'ArrowUp' ? -distance : 0
      const move = <T extends PageTextEdit | PageObjectEdit>(current: T | null): T | null => {
        if (generation !== editSelectionGenerationRef.current || !current || current.pageIndex !== selected.pageIndex || current.overlayId !== selected.overlayId || current.originalRect !== selected.originalRect) return current
        const box = pdfRectToViewport(viewport, current.rect)
        const left = clamp(box.left + dx, 0, Math.max(0, viewport.width - box.width))
        const top = clamp(box.top + dy, 0, Math.max(0, viewport.height - box.height))
        if (Math.abs(left - box.left) < 0.001 && Math.abs(top - box.top) < 0.001) return current
        return { ...current, rect: viewportRectToPdf(viewport, { left, top, right: left + box.width, bottom: top + box.height }), modified: true }
      }
      if (textEdit) setTextEdit(move)
      else setObjectEdit(move)
    } catch (error) { showToast(errorMessage(error)) }
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
    if (passwordProtected && nextTool !== 'hand' && nextTool !== 'select') {
      showToast('Password-protected PDFs are read-only in this version.')
      return
    }
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
    // Keys typed while an IME composes text belong to the composition.
    if (event.defaultPrevented || event.isComposing) return
    // The password prompt owns the keyboard until it is answered (pasting
    // and selecting inside its field keep working).
    if (passwordRequestRef.current) return
    // Ctrl+S is never swallowed: an open dialog closes first, and a save
    // pressed during a page operation runs as soon as the operation ends.
    if ((event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLocaleLowerCase() === 's') {
      event.preventDefault()
      if (!documentFile) return
      if (printDialogOpen) closePrintDialog()
      if (exportDialogOpen) setExportDialogOpen(false)
      if (ocrDialog && !ocrAbortRef.current) setOcrDialog(null)
      setOcrExportPrompt(null)
      setOcrOffer(null)
      void save(event.shiftKey)
      return
    }
    // While text is being recognised, Escape is Stop (wherever the focus is).
    if (ocrAbortRef.current) {
      event.preventDefault()
      if (event.key === 'Escape') stopOcr()
      return
    }
    if (documentUpdating || documentOperationRef.current) {
      event.preventDefault()
      return
    }
    if (saveReport) {
      if (event.key === 'Escape') { event.preventDefault(); setSaveReport(null) }
      return
    }
    if (printDialogOpen) {
      // The dialog owns Escape and Enter itself; ignore every other shortcut
      // so typing a range or tabbing between fields cannot switch tools.
      if (event.key === 'Escape') { event.preventDefault(); closePrintDialog() }
      else if (event.ctrlKey || event.metaKey) event.preventDefault()
      return
    }
    if (exportDialogOpen) {
      if (event.key === 'Escape') { event.preventDefault(); setExportDialogOpen(false) }
      else if (event.ctrlKey || event.metaKey) event.preventDefault()
      return
    }
    if (ocrDialog || ocrExportPrompt) {
      if (event.key === 'Escape') { event.preventDefault(); setOcrDialog(null); setOcrExportPrompt(null) }
      else if (event.ctrlKey || event.metaKey) event.preventDefault()
      return
    }
    if (ocrOffer && event.key === 'Escape') {
      event.preventDefault()
      setOcrOffer(null)
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
    if (typing && control && (key === 'z' || key === 'y')) return
    if (control && (key === 'c' || key === 'x') && (textEdit || objectEdit) && !typing) {
      event.preventDefault()
      copyCurrentEdit(key === 'x')
      return
    }
    if (control && key === 'v' && !typing && (tool === 'edit' || tool === 'addText')) {
      event.preventDefault()
      void pasteFromClipboard()
      return
    }
    if (control && key === 'a' && !typing && !event.shiftKey && !event.altKey && pdf) {
      // Select the whole document's text, not just the virtualised pages
      // that happen to be in the DOM; copy fills the rest (ContinuousPdfViewer).
      event.preventDefault()
      selectAllDocumentText()
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
    if (control && event.shiftKey && key === 'e') { event.preventDefault(); if (pdf && bytes && documentFile) setExportDialogOpen(true); return }
    if (control && event.key.toLocaleLowerCase() === 'p') { event.preventDefault(); if (pdf && bytes && documentFile) void openPrintDialog(); return }
    if (control && key === 'z') { event.preventDefault(); if (event.shiftKey) redo(); else undo(); return }
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
      const target = event.target instanceof HTMLElement ? event.target : null
      if (target && typing && target.closest('.edit-inspector')) {
        // Leaving an inspector field keeps its value and the selection.
        event.preventDefault()
        target.blur()
        return
      }
      if (textEdit || objectEdit) {
        // Escape keeps what was typed, moved or placed, as in Acrobat; only
        // the inspector's Cancel discards an edit.
        event.preventDefault()
        commitPendingEdits(true)
        return
      }
      if (selectingObjectRegion) setSelectingObjectRegion(false)
      else window.getSelection()?.removeAllRanges()
      return
    }
    if (!pdf || typing) return
    if (unmodified && event.key.startsWith('Arrow') && (textEdit || objectEdit)) {
      event.preventDefault(); void nudgeEdit(event.key, event.shiftKey ? 10 : 1)
    }
    else if (control && key === 'd' && (textEdit || objectEdit)) {
      event.preventDefault(); if (textEdit) duplicateText(); else duplicateObject()
    }
    else if (event.key === 'PageUp') { event.preventDefault(); scrollByScreen(-1) }
    else if (event.key === 'PageDown') { event.preventDefault(); scrollByScreen(1) }
    else if (control && event.key === 'Home') { event.preventDefault(); goToPage(0) }
    else if (control && event.key === 'End') { event.preventDefault(); goToPage(pdf.numPages - 1) }
    else if (control && (event.key === '+' || event.key === '=')) { event.preventDefault(); setZoomMode('custom'); setZoom((value) => clamp(value + 0.15, 0.35, 4)) }
    else if (control && event.key === '-') { event.preventDefault(); setZoomMode('custom'); setZoom((value) => clamp(value - 0.15, 0.35, 4)) }
    else if (control && event.key === '0') { event.preventDefault(); changeFit('fit') }
    else if (control && event.key === '1') { event.preventDefault(); changeZoom(1) }
    else if (control && event.key === '2') { event.preventDefault(); changeFit('width') }
    else if (unmodified && event.key.toLocaleLowerCase() === 'h') changeTool('hand')
    else if (unmodified && event.key.toLocaleLowerCase() === 'v') changeTool('select')
    else if (unmodified && event.key.toLocaleLowerCase() === 'e') changeTool('edit')
    else if (unmodified && event.key.toLocaleLowerCase() === 't') changeTool('addText')
    else if (unmodified && event.key.toLocaleLowerCase() === 'c') changeTool('crop')
    else if (event.key === 'Delete' && (textEdit || objectEdit)) deleteEditSelection()
    // Pages are deleted only from the Pages panel; Delete on a focused
    // bookmark or search result must never offer to delete pages.
    else if (event.key === 'Delete' && (document.activeElement as HTMLElement | null)?.closest('.thumbnails')) deletePages()
  }

  const handleKeyDownRef = useRef(handleKeyDown)
  handleKeyDownRef.current = handleKeyDown

  useEffect(() => {
    const listener = (event: KeyboardEvent) => handleKeyDownRef.current(event)
    window.addEventListener('keydown', listener)
    return () => window.removeEventListener('keydown', listener)
  }, [])

  // The shared window guard asks about unsaved work (Save / Don't Save / Cancel) on
  // every close path, so the title-bar button simply closes.
  function closeWindow() {
    window.simple.close()
  }

  // Answers for the window guard. Refs keep the handlers registered once while
  // always reading the latest document state.
  const ioStateRef = useRef({ dirty: false, saving: false, name: '', untitled: true, save: (_forceDialog?: boolean) => Promise.resolve() as Promise<unknown> })
  ioStateRef.current = {
    dirty: Boolean(dirty || textEdit?.modified || objectEdit?.modified),
    saving: busy === 'Saving PDF…',
    name: documentFile?.name ?? '',
    untitled: !documentFile?.path,
    save,
  }
  useEffect(() => {
    const io = getSimpleIO()
    if (!io) return undefined
    const offs = [
      io.onRequest('close-query', () => {
        const state = ioStateRef.current
        return { dirty: state.dirty, saving: state.saving, title: state.name || undefined, kind: 'PDF', untitled: state.untitled }
      }),
      io.onRequest('save-now', async () => {
        await ioStateRef.current.save(false)
        // save() records the saved revision synchronously; React state may not have re-rendered yet.
        return revisionRef.current === savedRevisionRef.current
      }),
      io.onRequest('discard', () => true),
    ]
    return () => offs.forEach((off) => off())
  }, [])

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

  // After Save reports a problem: show the form field or page it concerns.
  function revealProblem(problems: PdfSaveProblem[]) {
    const first = problems[0]
    if (!first) return
    if (first.field) {
      const control = document.querySelector<HTMLElement>(`[data-form-field="${CSS.escape(first.field)}"]`)
      if (control) {
        control.scrollIntoView({ block: 'center', inline: 'nearest' })
        control.focus({ preventScroll: true })
        return
      }
    }
    if (typeof first.pageIndex === 'number' && pdf) goToPage(first.pageIndex)
  }

  const titleBar = <TitleBar fileName={documentFile?.name} dirty={Boolean(dirty || textEdit?.modified || objectEdit?.modified)} onClose={closeWindow} />
  const pendingEditModified = Boolean(textEdit?.modified || objectEdit?.modified)
  const canUndoNow = Boolean(undoStack.length || pendingEditModified)
  const canRedoNow = Boolean((redoStack.length && !pendingEditModified) || ((textEdit || objectEdit) && pendingEditHistoryRef.current.future.length))

  // Rendered on the home screen and over a document alike: an open that needs
  // a password must always be able to ask for it.
  const passwordDialog = passwordRequest && (
    <div className="export-dialog-overlay" role="presentation">
      <form
        className="export-dialog"
        style={{ width: 400 }}
        role="dialog"
        aria-modal="true"
        aria-label="Password required"
        onKeyDown={(event) => {
          if (event.key !== 'Escape') return
          event.preventDefault()
          event.stopPropagation()
          passwordRequest.resolve(null)
        }}
        onSubmit={(event) => {
          event.preventDefault()
          passwordRequest.resolve(passwordDraft)
        }}
      >
        <div className="export-dialog-header"><strong>Password required</strong></div>
        <div className="export-dialog-body">
          <p style={{ margin: '0 0 10px', fontSize: 12 }}>“{passwordRequest.name}” is password protected. Enter the password to open it.</p>
          <input
            type="password"
            autoFocus
            aria-label="Password"
            value={passwordDraft}
            onChange={(event) => setPasswordDraft(event.target.value)}
            style={{ width: '100%', padding: '7px 9px', border: '1px solid var(--border-strong)', borderRadius: 6 }}
          />
          {passwordRequest.wrong && <p role="alert" style={{ margin: '8px 0 0', color: '#b42318', fontSize: 11 }}>That password is incorrect.</p>}
        </div>
        <div className="export-dialog-footer">
          <Button variant="secondary" type="button" onClick={() => passwordRequest.resolve(null)}>Cancel</Button>
          <Button variant="primary" type="submit">Open</Button>
        </div>
      </form>
    </div>
  )

  const saveReportBlocked = saveReport?.kind === 'blocked'
  const saveReportTitle = saveReportBlocked
    ? 'Not saved: some changes cannot be written as they are'
    : saveReport?.kind === 'partial' ? 'Saved, but without these changes' : 'Saved, with these notes'
  const saveReportDialog = saveReport && (
    <div className="export-dialog-overlay" role="presentation">
      <div className="export-dialog" style={{ width: 480 }} role="dialog" aria-modal="true" aria-label={saveReportTitle}>
        <div className="export-dialog-header">
          <strong>{saveReportTitle}</strong>
        </div>
        <div className="export-dialog-body">
          <ul className="save-report-list" style={{ margin: 0, paddingLeft: 18, fontSize: 12, lineHeight: 1.45, userSelect: 'text' }}>
            {[...new Set(saveReport.problems.map(problemText))].map((message) => <li key={message}>{message}</li>)}
          </ul>
          {saveReportBlocked && (
            <p style={{ margin: 0, color: 'var(--muted)', fontSize: 11 }}>Nothing was written and your changes are all still here. Fix them and save again, or save now without these changes.</p>
          )}
        </div>
        <div className="export-dialog-footer">
          {saveReportBlocked ? (
            <>
              <Button variant="secondary" autoFocus onClick={() => { const problems = saveReport.problems; setSaveReport(null); revealProblem(problems) }}>Keep editing</Button>
              <Button variant="primary" onClick={saveReport.onSaveAnyway}>Save anyway</Button>
            </>
          ) : (
            <Button variant="primary" autoFocus onClick={() => setSaveReport(null)}>OK</Button>
          )}
        </div>
      </div>
    </div>
  )

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
        {busy && !passwordRequest && <BusyOverlay label={busy} />}
        {passwordDialog}
        {toast && <Toast {...toast} onClose={() => setToast(null)} />}
      </div>
    )
  }

  const documentNotice = passwordProtected
    ? 'Protected PDF · reading only'
    : documentFile.signatureDetected
      ? 'Signed PDF · Save keeps the signed original and saves your changes as a copy'
      : documentFile.unlocked
        ? 'Protected PDF · Save writes an unprotected copy; the original is not changed'
        : ''

  return (
    <div
      className={`app-root document-app${immersive ? ' is-immersive' : ''}`}
      aria-busy={documentUpdating}
      onContextMenu={handleDocumentContextMenu}
    >
      {titleBar}
      <Toolbar
        sidebarOpen={sidebarOpen}
        pageIndex={pageIndex}
        pageCount={pdf?.numPages || 0}
        zoom={zoom}
        zoomMode={zoomMode}
        tool={tool}
        bookmarked={Boolean(toolbarBookmarkFor(bookmarks, pageIndex))}
        canUndo={canUndoNow}
        canRedo={canRedoNow}
        onToggleSidebar={() => {
          if (sidebarOpen) setActiveSearchMatch(null)
          setSidebarOpen((open) => !open)
        }}
        onOpen={openFile}
        onSave={() => save(false)}
        onExport={() => setExportDialogOpen(true)}
        onUndo={undo}
        onRedo={redo}
        onTool={changeTool}
        onPage={goToPage}
        onZoom={changeZoom}
        onFit={changeFit}
        onPrint={() => { if (pdf) void openPrintDialog() }}
        onRecognizeText={() => openOcrDialog('needed')}
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
            onReorder={reorderPages}
            onInsertBlank={insertBlankPage}
            onAddPages={addPages}
            onDuplicatePages={duplicatePages}
            onRotateLeft={() => rotatePages(-90)}
            onRotateRight={() => rotatePages(90)}
            onDeletePages={deletePages}
            onPageDragStart={startPageDrag}
            onImportPagesAt={importDroppedPages}
            onExportPages={exportPages}
            bookmarksEditable={!bookmarksReadOnly && !passwordProtected}
            onOpenBookmark={openBookmark}
            onDeleteBookmark={deleteBookmark}
            onRenameBookmark={renameBookmark}
            onRecognizeText={passwordProtected ? undefined : () => openOcrDialog('needed')}
          />
        )}
        <main
          ref={viewerRef}
          className={`viewer${isPanning ? ' is-panning' : ''}`}
          tabIndex={-1}
          style={{ outline: 'none' }}
          onPointerDownCapture={(event) => {
            // A secondary click is a menu gesture, never the beginning of an
            // ink, crop, rectangle, artwork-region, or transform gesture.
            if (event.button === 2) event.stopPropagation()
          }}
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
          {documentLoading && pdf && (
            <div className="document-refreshing" role="status" aria-live="polite">
              <RotateCcw className="spin" size={14} />
              <span>Updating pages…</span>
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
              // The viewer passes these through to PdfPage untouched; list
              // boxes hold string[] values (PdfFormValue), which its prop type
              // does not spell out yet.
              formValues={formValues as Record<string, string | boolean>}
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
              // Nothing on the page discards an edit (Escape keeps it); only
              // the inspector's Cancel does.
              onCancelTextEdit={commitTextEdit}
              onRequestObjectEdit={beginObjectEdit}
              onObjectEditChange={setObjectEdit}
              onCommitObjectEdit={commitObjectEdit}
              onCancelObjectEdit={commitObjectEdit}
              onObjectRegionSelected={() => setSelectingObjectRegion(false)}
              onHighlight={(index, rects) => addHighlights(rects, index)}
              onTextMarkup={(index, style, rects, displayRotation) => addTextMarkup(style, rects, index, displayRotation)}
              onInk={(index, points) => addInk(points, index)}
              onRectangle={(index, rect) => addRectangle(rect, index)}
              onCrop={(index, rect) => cropPage(rect, index)}
              onPlaceSignature={(index, point, displayRotation) => { void placeSignature(index, point, displayRotation) }}
              onNavigate={goToPage}
              onFormChange={changeFormValue}
              onRequestOcrOffer={requestOcrOffer}
              pendingEditAt={pendingEditAt}
              onPendingEditAtHandled={(token) => setPendingEditAt((current) => current?.token === token ? null : current)}
              onScanEditPrepared={mergeScanPreparation}
            />
          )}
          {documentNotice && pdf && (
            <div
              className="readonly-badge document-notice"
              role="status"
              // Stay clear of the edit inspector's buttons along the right edge.
              style={tool === 'edit' || tool === 'addText' ? { right: 17 + 278 } : undefined}
            >
              {documentFile.signatureDetected && !passwordProtected ? <Signature size={13} /> : passwordProtected ? <AlertTriangle size={13} /> : <Lock size={13} />}
              {documentNotice}
            </div>
          )}
        </main>
        {(tool === 'edit' || tool === 'addText') && pdf && (
          <EditInspector
            textEdit={textEdit}
            pageScanState={pageScanState}
            onRecognizeText={() => openOcrDialog('needed')}
            objectEdit={objectEdit}
            selectingObjectRegion={selectingObjectRegion}
            onTextChange={changeTextEditStep}
            onObjectChange={changeObjectEditStep}
            onCommitText={commitTextEdit}
            onCommitObject={commitObjectEdit}
            onCancelSelection={cancelEditSelection}
            onDeleteSelection={deleteEditSelection}
            onCopySelection={() => { copyCurrentEdit(false) }}
            onPasteSelection={() => { void pasteFromClipboard() }}
            canPasteSelection={hasEditClipboard}
            onDuplicateText={duplicateText}
            onAddText={() => { cancelEditSelection(); setTool('addText') }}
            onAddImage={addImageObject}
            onReplaceImage={replaceObjectImage}
            onRotateObject={rotateObject}
            onFlipObject={flipObject}
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
      {exportDialogOpen && pdf && (
        <ExportDialog
          documentName={documentFile.name}
          pageCount={pdf.numPages}
          currentPage={pageIndex}
          selectedPages={[...selectedPages].sort((a, b) => a - b)}
          onExport={(submission) => { void runExportAs(submission) }}
          onClose={() => setExportDialogOpen(false)}
        />
      )}
      {ocrDialog && pdf && (
        <OcrDialog
          pdf={pdf}
          documentName={documentFile.name}
          currentPage={pageIndex}
          selectedPages={[...selectedPages].sort((a, b) => a - b)}
          initialScope={ocrDialog.scope}
          signatureDetected={documentFile.signatureDetected}
          canReleasePage={canReleasePage}
          progress={ocrProgress}
          stopping={ocrStopping}
          onRecognize={(request) => { void recognizeText(request) }}
          onStop={stopOcr}
          onClose={() => { if (!ocrAbortRef.current) setOcrDialog(null) }}
        />
      )}
      {ocrOffer && !ocrDialog && (
        <OcrOffer
          x={ocrOffer.client.x}
          y={ocrOffer.client.y}
          onRecognize={acceptOcrOffer}
          onDecline={declineOcrOffer}
          onDismiss={() => setOcrOffer(null)}
        />
      )}
      {ocrExportPrompt && (
        <OcrExportPrompt
          pageCount={ocrExportPrompt.pages.length}
          onRecognize={() => {
            const prompt = ocrExportPrompt
            setOcrExportPrompt(null)
            void recognizeText({ scope: 'selected', pages: prompt.pages, explicit: true, language: ocrLanguageRef.current, replaceExisting: false, exportJob: prompt.job })
          }}
          onExportAnyway={() => {
            const prompt = ocrExportPrompt
            setOcrExportPrompt(null)
            void runExportAs(prompt.job, { skipScanCheck: true })
          }}
          onCancel={() => setOcrExportPrompt(null)}
        />
      )}
      {printDialogOpen && printPreviewPdf && (
        <PrintDialog
          documentName={documentFile.name}
          pdf={printPreviewPdf}
          pageCount={printPreviewPdf.numPages}
          currentPage={pageIndex}
          selectedPages={[...selectedPages].sort((a, b) => a - b)}
          onPrint={(submission) => { void runPrintJob(submission) }}
          onClose={closePrintDialog}
        />
      )}
      {contextMenu && (
        <div
          ref={contextMenuRef}
          className="pdf-context-menu"
          role="menu"
          aria-label={hasContextEditSelection ? 'Selected PDF content' : 'PDF commands'}
          style={{ left: contextMenu.x, top: contextMenu.y }}
          onContextMenu={(event) => event.preventDefault()}
          onKeyDown={(event) => {
            const buttons = [...event.currentTarget.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')]
            const current = buttons.indexOf(document.activeElement as HTMLButtonElement)
            const focusAt = (index: number) => buttons[index]?.focus({ preventScroll: true })
            if (event.key === 'ArrowDown' || (event.key === 'Tab' && !event.shiftKey)) {
              event.preventDefault()
              focusAt((current + 1 + buttons.length) % buttons.length)
            } else if (event.key === 'ArrowUp' || (event.key === 'Tab' && event.shiftKey)) {
              event.preventDefault()
              focusAt((current - 1 + buttons.length) % buttons.length)
            } else if (event.key === 'Home') {
              event.preventDefault()
              focusAt(0)
            } else if (event.key === 'End') {
              event.preventDefault()
              focusAt(buttons.length - 1)
            } else if (event.key === 'Escape') {
              event.preventDefault()
              event.stopPropagation()
              closeContextMenu(true)
            }
          }}
        >
          {canContextUndo && (
            <button type="button" role="menuitem" data-context-action="undo" onClick={() => runContextMenuAction(undo)}>
              <Undo2 size={15} aria-hidden="true" /><span>Undo</span><kbd>Ctrl+Z</kbd>
            </button>
          )}
          {canContextRedo && (
            <button type="button" role="menuitem" data-context-action="redo" onClick={() => runContextMenuAction(redo)}>
              <Redo2 size={15} aria-hidden="true" /><span>Redo</span><kbd>Ctrl+Y</kbd>
            </button>
          )}
          {(canContextUndo || canContextRedo) && (hasContextEditSelection || canContextPaste) && <div className="pdf-context-separator" role="separator" />}
          {hasContextEditSelection && (
            <button type="button" role="menuitem" data-context-action="cut" onClick={() => runContextMenuAction(() => { copyCurrentEdit(true) })}>
              <Scissors size={15} aria-hidden="true" /><span>Cut</span><kbd>Ctrl+X</kbd>
            </button>
          )}
          {hasContextEditSelection && (
            <button type="button" role="menuitem" data-context-action="copy" onClick={() => runContextMenuAction(() => { copyCurrentEdit(false) })}>
              <Copy size={15} aria-hidden="true" /><span>Copy</span><kbd>Ctrl+C</kbd>
            </button>
          )}
          {canContextPaste && (
            <button type="button" role="menuitem" data-context-action="paste" onClick={() => runContextMenuAction(pasteFromClipboard, false)}>
              <ClipboardPaste size={15} aria-hidden="true" /><span>Paste</span><kbd>Ctrl+V</kbd>
            </button>
          )}
          {hasContextEditSelection && (
            <button type="button" role="menuitem" className="is-destructive" data-context-action="delete" onClick={() => runContextMenuAction(deleteEditSelection)}>
              <Trash2 size={15} aria-hidden="true" /><span>Delete</span><kbd>Del</kbd>
            </button>
          )}
        </div>
      )}
      {documentUpdating && <div className="document-update-guard" aria-hidden="true" />}
      {busy && !passwordRequest && <BusyOverlay label={busy} />}
      {saveReportDialog}
      {passwordDialog}
      {toast && <Toast {...toast} onClose={() => setToast(null)} />}
    </div>
  )
}
