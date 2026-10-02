import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { DragEvent as ReactDragEvent, RefObject } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { GripVertical } from 'lucide-react'
import type { ActiveSearchMatch, DisplayRotation, PageObjectEdit, PageTextEdit, PdfOverlay, PdfRect, ToolMode } from '../types'
import { isImportableTransferFile, pageDropEdge, pageDropInsertIndex, type PageDropEdge } from '../lib/pageTransfer'
import {
  collectSelectedPdfText,
  joinSelectedPageTexts,
  plainTextFromTextContentItems,
  textCaretAtPoint,
  wordTextSliceAtPoint,
} from '../lib/textSelection'
import type { PdfTextContentItemLike, TextCaretPoint } from '../lib/textSelection'
import { getPageTextContent } from '../lib/pdf'
import { isTypingTarget } from '../lib/utils'
import { PdfPage, type PendingEditAt } from './PdfPage'
import type { ScanPreparation } from '../lib/ocr/scanEdit'

interface PageSize {
  width: number
  height: number
}

interface ContinuousPdfViewerProps {
  pdf: PDFDocumentProxy
  viewerRef: RefObject<HTMLElement>
  currentPage: number
  selectedPages: Set<number>
  zoom: number
  rotations: Record<number, number>
  tool: ToolMode
  overlays: PdfOverlay[]
  formValues: Record<string, string | boolean>
  textEdit: PageTextEdit | null
  objectEdit: PageObjectEdit | null
  activeSearchMatch: ActiveSearchMatch | null
  selectingObjectRegion: boolean
  onCurrentPage: (index: number) => void
  onSelectPage: (index: number, event: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean }) => void
  onPageDragStart: (index: number, event: ReactDragEvent<HTMLButtonElement>) => void
  onImportPagesAt: (files: File[], insertIndex: number) => void
  onRequestTextEdit: (edit: PageTextEdit) => void
  onTextEditChange: (edit: PageTextEdit) => void
  onCommitTextEdit: () => void
  onCancelTextEdit: () => void
  onRequestObjectEdit: (edit: PageObjectEdit) => void
  onObjectEditChange: (edit: PageObjectEdit) => void
  onCommitObjectEdit: () => void
  onCancelObjectEdit: () => void
  onObjectRegionSelected: () => void
  onHighlight: (pageIndex: number, rects: PdfRect[]) => void
  onTextMarkup: (pageIndex: number, style: 'underline' | 'strikeout', rects: PdfRect[], displayRotation: DisplayRotation) => void
  onInk: (pageIndex: number, points: Array<{ x: number; y: number }>) => void
  onRectangle: (pageIndex: number, rect: PdfRect) => void
  onCrop: (pageIndex: number, rect: PdfRect) => void
  onPlaceSignature: (pageIndex: number, point: { x: number; y: number }, displayRotation: DisplayRotation) => void
  onNavigate: (pageIndex: number) => void
  onFormChange: (name: string, value: string | boolean) => void
  /** A click in Edit mode landed on a scanned page without text. */
  onRequestOcrOffer?: (pageIndex: number, point: { x: number; y: number }, client: { x: number; y: number }) => void
  /** Open the text at this point once the page's (new) text is ready. */
  pendingEditAt?: PendingEditAt | null
  onPendingEditAtHandled?: (token: number) => void
  /** An edit of scanned text finished preparing (patch, words, matched style). */
  onScanEditPrepared?: (key: string, result: ScanPreparation) => void
}

const PAGE_VERTICAL_CHROME = 40
const POINTER_TEXT_TOOLS: ToolMode[] = ['select', 'edit', 'highlight', 'underline', 'strikeout']
// An idle selection keeps the pages between its ends mounted (unmounting
// nodes inside a live Range makes Chromium repaint it elsewhere), but only up
// to this span; longer selections keep just their end pages, and copy fills
// the pages in between from getTextContent().
const MAX_PROTECTED_SELECTION_PAGES = 12
// A page change requested by a click inside the viewer (editing, signing, the
// page handle) never scrolls a page that is already on screen.
const QUIET_PAGE_CHANGE_MS = 1500
const NAVIGATION_SUPPRESS_MS = 700

interface PageScopedHandlers {
  onPageReady: (size: PageSize) => void
  onHighlight: (rects: PdfRect[]) => void
  onTextMarkup: (style: 'underline' | 'strikeout', rects: PdfRect[], displayRotation: DisplayRotation) => void
  onInk: (points: Array<{ x: number; y: number }>) => void
  onRectangle: (rect: PdfRect) => void
  onCrop: (rect: PdfRect) => void
  onPlaceSignature: (point: { x: number; y: number }, displayRotation: DisplayRotation) => void
}

function selectionNodeIsPdfText(root: HTMLElement, node: Node | null) {
  if (!node || !root.contains(node)) return false
  const element = node instanceof Element ? node : node.parentElement
  return Boolean(element?.closest('.text-layer'))
}

function slotPageIndex(node: Node | null) {
  const element = node instanceof Element ? node : node?.parentElement
  const slot = element?.closest<HTMLElement>('.continuous-page-slot')
  const index = Number(slot?.dataset.pageIndex)
  return slot && Number.isInteger(index) ? index : undefined
}

/** Pages an idle native selection needs mounted to stay intact. */
function selectionProtectedPages(root: HTMLElement) {
  const selection = window.getSelection()
  if (!selection || selection.isCollapsed || selection.rangeCount < 1) return new Set<number>()
  const ends = [selection.anchorNode, selection.focusNode]
    .filter((node): node is Node => Boolean(node && root.contains(node)))
    .map(slotPageIndex)
    .filter((index): index is number => index !== undefined)
  if (!ends.length) return new Set<number>()
  const first = Math.min(...ends)
  const last = Math.max(...ends)
  if (last - first + 1 > MAX_PROTECTED_SELECTION_PAGES) return new Set(ends)
  return new Set(Array.from({ length: last - first + 1 }, (_, offset) => first + offset))
}

function slotIntersectsViewport(root: HTMLElement, slot: HTMLElement | undefined) {
  if (!slot) return false
  const bounds = root.getBoundingClientRect()
  const rect = slot.getBoundingClientRect()
  return rect.bottom > bounds.top + 1 && rect.top < bounds.bottom - 1 && rect.height > 0
}

/**
 * The page the user is looking at: the first page at the very top of the
 * document, the last page at the very bottom, otherwise the page occupying
 * the most viewport height (ties go to the upper page).
 */
function mostVisiblePage(root: HTMLElement, slots: HTMLElement[], pageCount: number) {
  const maxScroll = root.scrollHeight - root.clientHeight
  if (maxScroll > 1 && pageCount > 0) {
    if (root.scrollTop <= 1) return 0
    if (root.scrollTop >= maxScroll - 1) return pageCount - 1
  }
  const bounds = root.getBoundingClientRect()
  let best: number | undefined
  let bestVisible = 0
  const ordered = slots
    .map((slot) => ({ slot, index: Number(slot.dataset.pageIndex) }))
    .filter(({ index }) => Number.isInteger(index))
    .sort((a, b) => a.index - b.index)
  for (const { slot, index } of ordered) {
    const rect = slot.getBoundingClientRect()
    const visible = Math.min(rect.bottom, bounds.bottom) - Math.max(rect.top, bounds.top)
    if (visible > bestVisible + 0.5) {
      best = index
      bestVisible = visible
    }
  }
  return best
}

export function ContinuousPdfViewer(props: ContinuousPdfViewerProps) {
  const {
    pdf, viewerRef, currentPage, selectedPages, zoom, rotations, tool, overlays, formValues,
    textEdit, objectEdit, activeSearchMatch, selectingObjectRegion, onCurrentPage,
    onSelectPage, onPageDragStart, onImportPagesAt, pendingEditAt,
  } = props
  const [mountedPages, setMountedPages] = useState<Set<number>>(() => new Set([currentPage]))
  // Page sizes are cached at scale 1 (unzoomed CSS pixels) so a zoom change
  // rescales every placeholder without discarding measurements.
  const [pageSizes, setPageSizes] = useState<Record<number, PageSize>>({})
  const [basePageSizes, setBasePageSizes] = useState<Record<number, PageSize>>({})
  const [estimatedSize, setEstimatedSize] = useState<PageSize>({ width: 612, height: 792 })
  const [dropTarget, setDropTarget] = useState<{ pageIndex: number; edge: PageDropEdge } | null>(null)
  const slotsRef = useRef(new Map<number, HTMLElement>())
  const reportedFromScrollRef = useRef(new Map<number, number>())
  const currentPageRef = useRef(currentPage)
  const pinnedPagesRef = useRef<Array<number | undefined>>([])
  const propsRef = useRef(props)
  const pageHandlersRef = useRef(new Map<number, PageScopedHandlers>())
  const previousRotationsRef = useRef<Record<number, number>>(rotations)
  const quietPageChangesRef = useRef(new Map<number, number>())
  const visiblePagesRef = useRef(new Set<number>())
  const renderIntersectingPagesRef = useRef(new Set<number>())
  // True only while a pointer drag is selecting text; an idle selection left
  // behind never freezes page tracking or pruning.
  const pointerSelectionActiveRef = useRef(false)
  const suppressObserverUntilRef = useRef(0)
  const suppressionTimerRef = useRef(0)
  const evaluateCurrentPageRef = useRef<((onlyWhenCurrentHidden?: boolean) => void) | null>(null)
  const copyGenerationRef = useRef(0)
  const basePageSizesRef = useRef(basePageSizes)
  const renderedPagesRef = useRef(new Set<number>())
  const layoutAnchorRef = useRef<{ pageIndex: number; offset: number } | null>(null)
  const layoutSignatureRef = useRef('')
  const layoutPdfRef = useRef(pdf)
  const stopLayoutAnchorRef = useRef<(() => void) | null>(null)

  currentPageRef.current = currentPage
  pinnedPagesRef.current = [textEdit?.pageIndex, objectEdit?.pageIndex, activeSearchMatch?.pageIndex]
  basePageSizesRef.current = basePageSizes
  propsRef.current = props
  renderedPagesRef.current = new Set([currentPage, textEdit?.pageIndex, objectEdit?.pageIndex, activeSearchMatch?.pageIndex, ...mountedPages]
    .filter((index): index is number => Number.isInteger(index)))

  // Anchor the viewport to the current page before a zoom- or rotation-driven
  // relayout; the layout effect below restores the same fractional offset.
  const layoutSignature = `${zoom}|${Object.entries(rotations).map(([index, amount]) => `${index}:${amount}`).join(',')}`
  if (layoutPdfRef.current !== pdf) {
    const root = viewerRef.current
    const slot = slotsRef.current.get(currentPage)
    if (root && slot) {
      const rootBounds = root.getBoundingClientRect()
      const slotBounds = slot.getBoundingClientRect()
      layoutAnchorRef.current = {
        pageIndex: currentPage,
        offset: slotBounds.height > 0 ? (rootBounds.top - slotBounds.top) / slotBounds.height : 0,
      }
    }
    layoutPdfRef.current = pdf
    layoutSignatureRef.current = layoutSignature
  }
  if (layoutSignatureRef.current !== layoutSignature) {
    const root = viewerRef.current
    const slot = slotsRef.current.get(currentPage)
    if (layoutSignatureRef.current && root && slot) {
      const rootBounds = root.getBoundingClientRect()
      const slotBounds = slot.getBoundingClientRect()
      layoutAnchorRef.current = {
        pageIndex: currentPage,
        offset: slotBounds.height > 0 ? (rootBounds.top - slotBounds.top) / slotBounds.height : 0,
      }
    }
    layoutSignatureRef.current = layoutSignature
  }

  // Programmatic scrolls briefly pause page tracking. When the pause ends,
  // catch the indicator up if the current page has left the screen meanwhile.
  const suppressPageTracking = (milliseconds: number) => {
    suppressObserverUntilRef.current = performance.now() + milliseconds
    window.clearTimeout(suppressionTimerRef.current)
    suppressionTimerRef.current = window.setTimeout(() => evaluateCurrentPageRef.current?.(true), milliseconds + 20)
  }
  const suppressPageTrackingRef = useRef(suppressPageTracking)
  suppressPageTrackingRef.current = suppressPageTracking
  useEffect(() => () => window.clearTimeout(suppressionTimerRef.current), [])

  useLayoutEffect(() => {
    const anchor = layoutAnchorRef.current
    if (!anchor) return
    layoutAnchorRef.current = null
    const root = viewerRef.current
    if (!root || !slotsRef.current.get(anchor.pageIndex)) return
    suppressPageTrackingRef.current(NAVIGATION_SUPPRESS_MS)
    const applyAnchor = () => {
      const slot = slotsRef.current.get(anchor.pageIndex)
      if (!slot) return
      const rootBounds = root.getBoundingClientRect()
      const slotBounds = slot.getBoundingClientRect()
      const delta = slotBounds.top - rootBounds.top + anchor.offset * slotBounds.height
      if (Math.abs(delta) >= 1) root.scrollTop += delta
    }
    applyAnchor()
    // Rendered pages adopt the new size a commit or two later (and the
    // geometry read here can still predate this commit), which used to leave
    // the view several pages away from where the user was reading. Hold the
    // anchor for a few frames until the layout settles, unless the user
    // scrolls or navigates meanwhile.
    let frame = 0
    let active = true
    const until = performance.now() + 600
    const stop = () => {
      if (!active) return
      active = false
      window.cancelAnimationFrame(frame)
      root.removeEventListener('wheel', stop)
      root.removeEventListener('pointerdown', stop, true)
      window.removeEventListener('keydown', stop, true)
      if (stopLayoutAnchorRef.current === stop) stopLayoutAnchorRef.current = null
    }
    const tick = () => {
      if (!active) return
      if (performance.now() > until) {
        stop()
        return
      }
      applyAnchor()
      frame = window.requestAnimationFrame(tick)
    }
    stopLayoutAnchorRef.current?.()
    stopLayoutAnchorRef.current = stop
    root.addEventListener('wheel', stop, { passive: true })
    root.addEventListener('pointerdown', stop, true)
    window.addEventListener('keydown', stop, true)
    frame = window.requestAnimationFrame(tick)
    return stop
  }, [pdf, zoom, rotations, viewerRef])

  useEffect(() => {
    let cancelled = false
    pdf.getPage(1).then((page) => {
      if (cancelled) return
      const angle = (((page.rotate || 0) + (rotations[0] || 0)) % 360 + 360) % 360
      const viewport = page.getViewport({ scale: 1, rotation: angle })
      setEstimatedSize({ width: viewport.width, height: viewport.height })
    }).catch(() => {})
    return () => { cancelled = true }
  }, [pdf, rotations[0]])

  // Fetch every page's intrinsic dimensions once so mixed-size documents get
  // accurate placeholders. This reads page metadata only — no rendering.
  useEffect(() => {
    let cancelled = false
    setBasePageSizes({})
    ;(async () => {
      const sizes: Record<number, PageSize> = {}
      for (let index = 0; index < pdf.numPages; index += 1) {
        const page = await pdf.getPage(index + 1)
        if (cancelled) return
        const viewport = page.getViewport({ scale: 1 })
        sizes[index] = { width: viewport.width, height: viewport.height }
      }
      if (!cancelled) setBasePageSizes(sizes)
    })().catch(() => {})
    return () => { cancelled = true }
  }, [pdf])

  useEffect(() => {
    const previous = previousRotationsRef.current
    const changed = new Set([...Object.keys(previous), ...Object.keys(rotations)]
      .map(Number)
      .filter((index) => (previous[index] || 0) !== (rotations[index] || 0)))
    previousRotationsRef.current = rotations
    if (!changed.size) return
    setPageSizes((current) => {
      let mutated = false
      const next = { ...current }
      for (const index of changed) {
        const size = next[index]
        if (!size) continue
        const base = basePageSizesRef.current[index]
        if (base) {
          const expected = (rotations[index] || 0) % 180 ? { width: base.height, height: base.width } : base
          if (Math.abs(size.width - expected.width) >= 0.5 || Math.abs(size.height - expected.height) >= 0.5) {
            next[index] = { ...expected }
            mutated = true
          }
        } else if (!renderedPagesRef.current.has(index)) {
          // A mounted page re-reports its own rotated size; only stale entries
          // for unmounted pages need the orientation swap here.
          const delta = ((((rotations[index] || 0) - (previous[index] || 0)) % 360) + 360) % 360
          if (delta % 180) {
            next[index] = { width: size.height, height: size.width }
            mutated = true
          }
        }
      }
      return mutated ? next : current
    })
  }, [rotations])

  useEffect(() => {
    setMountedPages(new Set([currentPage]))
    setPageSizes({})
    visiblePagesRef.current.clear()
    renderIntersectingPagesRef.current.clear()
    reportedFromScrollRef.current.clear()
    quietPageChangesRef.current.clear()
    pointerSelectionActiveRef.current = false
  }, [pdf])

  useEffect(() => {
    const root = viewerRef.current
    if (!root) return
    let selectingWithPointer = false
    let settleFrame = 0

    // Drop pages that are off screen and no longer hold an end (or, for short
    // selections, the middle) of the native selection.
    const prunePages = () => {
      if (pointerSelectionActiveRef.current) return
      const protectedPages = selectionProtectedPages(root)
      setMountedPages((current) => {
        const next = new Set([...current].filter((index) => (
          renderIntersectingPagesRef.current.has(index)
          || index === currentPageRef.current
          || pinnedPagesRef.current.includes(index)
          || protectedPages.has(index)
        )))
        if (!next.size) next.add(currentPageRef.current)
        return next.size === current.size && [...next].every((index) => current.has(index)) ? current : next
      })
    }

    const syncSelectionState = () => {
      const dragging = selectingWithPointer || root.dataset.pointerSelectingText === 'true'
      const wasDragging = pointerSelectionActiveRef.current
      pointerSelectionActiveRef.current = dragging
      if (dragging) return
      prunePages()
      // Page tracking paused for the drag; catch up with where it ended.
      if (wasDragging) evaluateCurrentPageRef.current?.()
    }
    const scheduleSync = () => {
      window.cancelAnimationFrame(settleFrame)
      settleFrame = window.requestAnimationFrame(syncSelectionState)
    }

    const handlePointerDown = (event: PointerEvent) => {
      const target = event.target
      selectingWithPointer = target instanceof Element
        && root.contains(target)
        && Boolean(target.closest('.text-layer'))
      if (selectingWithPointer) pointerSelectionActiveRef.current = true
    }
    const handlePointerFinished = () => {
      selectingWithPointer = false
      scheduleSync()
    }

    document.addEventListener('selectionchange', scheduleSync)
    document.addEventListener('pointerdown', handlePointerDown, true)
    document.addEventListener('pointerup', handlePointerFinished, true)
    document.addEventListener('pointercancel', handlePointerFinished, true)
    return () => {
      window.cancelAnimationFrame(settleFrame)
      document.removeEventListener('selectionchange', scheduleSync)
      document.removeEventListener('pointerdown', handlePointerDown, true)
      document.removeEventListener('pointerup', handlePointerFinished, true)
      document.removeEventListener('pointercancel', handlePointerFinished, true)
    }
  }, [pdf, viewerRef])

  // Clicking anywhere in the document gives it keyboard focus (as Acrobat
  // does), even where a tool cancels the pointerdown default, so page keys
  // and tool shortcuts work again after using the search, page or zoom boxes.
  useEffect(() => {
    const root = viewerRef.current
    if (!root) return
    const focusDocument = (event: PointerEvent) => {
      const active = document.activeElement
      if (active === root || (active instanceof Node && root.contains(active))) return
      const target = event.target
      if (target instanceof Element
        && target.closest('input, textarea, select, button, a[href], [contenteditable="true"]')) return
      root.focus({ preventScroll: true })
    }
    root.addEventListener('pointerdown', focusDocument, true)
    return () => root.removeEventListener('pointerdown', focusDocument, true)
  }, [viewerRef])

  // Copy builds the text itself: text-layer spans are absolutely positioned,
  // so the native serialisation glues lines and words together, and pages
  // outside the virtualised window are not in the DOM at all.
  useEffect(() => {
    const root = viewerRef.current
    if (!root) return
    const handleCopy = (event: ClipboardEvent) => {
      if (event.defaultPrevented || isTypingTarget(document.activeElement)) return
      const selection = window.getSelection()
      if (!selection || selection.isCollapsed || selection.rangeCount < 1) return
      const pages = collectSelectedPdfText(root, selection.getRangeAt(0))
      if (!pages) return
      const missing = pages.some((page) => page.text === null)
      const available = joinSelectedPageTexts(pages.map((page) => page.text))
      if (!available && !missing) return
      event.preventDefault()
      event.clipboardData?.setData('text/plain', available)
      const generation = ++copyGenerationRef.current
      if (!missing) return
      void Promise.all(pages.map(async (page) => {
        if (page.text !== null) return page.text
        const content = await getPageTextContent(await pdf.getPage(page.pageIndex + 1))
        return plainTextFromTextContentItems(content.items as PdfTextContentItemLike[])
      })).then((texts) => {
        if (generation !== copyGenerationRef.current) return
        return navigator.clipboard?.writeText(joinSelectedPageTexts(texts))
      }).catch(() => {})
    }
    document.addEventListener('copy', handleCopy)
    return () => document.removeEventListener('copy', handleCopy)
  }, [pdf, viewerRef])

  useEffect(() => {
    const root = viewerRef.current
    if (!root || !POINTER_TEXT_TOOLS.includes(tool)) return

    type SelectionPoint = { node: Text; offset: number }
    type SelectionMode = 'character' | 'word' | 'line'

    type SelectionGesture = {
      pointerId: number
      mode: SelectionMode
      anchorStart: SelectionPoint
      anchorEnd: SelectionPoint
      focusNode: Text
      focusOffset: number
      focusSpan: HTMLSpanElement
      startX: number
      startY: number
      clientX: number
      clientY: number
    }

    let gesture: SelectionGesture | null = null
    let lastPress: { time: number; x: number; y: number; button: number; clickCount: number } | null = null
    let autoScrollFrame = 0

    const comparePoints = (a: SelectionPoint, b: SelectionPoint) => {
      if (a.node === b.node) return a.offset - b.offset
      return a.node.compareDocumentPosition(b.node) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1
    }

    const wordRangeAtPoint = (caret: TextCaretPoint, clientX: number, clientY: number): { start: SelectionPoint; end: SelectionPoint } | null => {
      const slice = wordTextSliceAtPoint(caret.span, clientX, clientY)
      return slice
        ? { start: { node: caret.node, offset: slice.start }, end: { node: caret.node, offset: slice.end } }
        : null
    }

    // All spans on the clicked span's visual line, matched with the same
    // baseline clustering rule the text layer used to order them: flow-axis
    // distance within 40% of the font height, at a comparable angle. Spans in
    // a side-by-side column share that flow band, so the candidates are then
    // split at reading-axis gaps of an em or more and only the run around the
    // clicked span survives — triple-click never selects the other column.
    const lineRangeForSpan = (span: HTMLSpanElement): { start: SelectionPoint; end: SelectionPoint } | null => {
      const layer = span.closest<HTMLElement>('.text-layer')
      const anchorAngle = Number(span.dataset.textAngle) || 0
      const flowX = -Math.sin(anchorAngle)
      const flowY = Math.cos(anchorAngle)
      const readX = Math.cos(anchorAngle)
      const readY = Math.sin(anchorAngle)
      const flowOf = (item: HTMLSpanElement) => Number(item.dataset.textBaselineX) * flowX + Number(item.dataset.textBaselineY) * flowY
      const readOf = (item: HTMLSpanElement) => Number(item.dataset.textBaselineX) * readX + Number(item.dataset.textBaselineY) * readY
      const advanceOf = (item: HTMLSpanElement) => Number(item.dataset.textAdvance) || 0
      const heightOf = (item: HTMLSpanElement) => Number.parseFloat(item.style.fontSize) || 0
      const anchorFlow = flowOf(span)
      if (!layer || !Number.isFinite(anchorFlow)) return null
      const candidates: HTMLSpanElement[] = []
      for (const item of layer.querySelectorAll<HTMLSpanElement>('[data-text-item="true"]')) {
        if (!(item.firstChild instanceof Text)) continue
        if (item !== span) {
          const angle = Number(item.dataset.textAngle) || 0
          const turn = Math.atan2(Math.sin(angle - anchorAngle), Math.cos(angle - anchorAngle))
          const flow = flowOf(item)
          if (Math.abs(turn) > 0.26 || !Number.isFinite(flow)
            || Math.abs(flow - anchorFlow) > Math.max(1, Math.max(heightOf(span), heightOf(item)) * 0.4)) continue
        }
        candidates.push(item)
      }
      candidates.sort((a, b) => readOf(a) - readOf(b))
      let segment: HTMLSpanElement[] = []
      for (const [index, item] of candidates.entries()) {
        const previous = candidates[index - 1]
        if (previous && readOf(item) - (readOf(previous) + advanceOf(previous)) > Math.max(heightOf(previous), heightOf(item))) {
          if (segment.includes(span)) break
          segment = []
        }
        segment.push(item)
      }
      const first = segment[0]?.firstChild
      const last = segment.at(-1)?.firstChild
      return first instanceof Text && last instanceof Text
        ? { start: { node: first, offset: 0 }, end: { node: last, offset: last.length } }
        : null
    }

    const applySelection = (
      selection: Selection,
      mode: SelectionMode,
      anchorStart: SelectionPoint,
      anchorEnd: SelectionPoint,
      caret: TextCaretPoint,
      clientX: number,
      clientY: number,
    ) => {
      let focusStart: SelectionPoint = { node: caret.node, offset: caret.offset }
      let focusEnd = focusStart
      if (mode === 'word') {
        const word = wordRangeAtPoint(caret, clientX, clientY)
        if (word) ({ start: focusStart, end: focusEnd } = word)
      } else if (mode === 'line') {
        const line = lineRangeForSpan(caret.span)
        if (line) ({ start: focusStart, end: focusEnd } = line)
      }
      const forward = comparePoints(anchorStart, focusStart) <= 0
      const base = forward ? anchorStart : anchorEnd
      const extent = forward ? focusEnd : focusStart
      selection.setBaseAndExtent(base.node, base.offset, extent.node, extent.offset)
    }

    const setFocusFromPoint = (clientX: number, clientY: number) => {
      if (!gesture) return false
      const caret = textCaretAtPoint(root, clientX, clientY, gesture.focusSpan)
      if (!caret) return false
      if (caret.node === gesture.focusNode && caret.offset === gesture.focusOffset) return true
      const selection = window.getSelection()
      if (!selection || !gesture.anchorStart.node.isConnected || !gesture.anchorEnd.node.isConnected || !caret.node.isConnected) return false
      try {
        applySelection(selection, gesture.mode, gesture.anchorStart, gesture.anchorEnd, caret, clientX, clientY)
        gesture.focusNode = caret.node
        gesture.focusOffset = caret.offset
        gesture.focusSpan = caret.span
        return true
      } catch {
        return false
      }
    }

    const autoScrollAmount = (clientY: number) => {
      const bounds = root.getBoundingClientRect()
      const edge = Math.min(44, Math.max(24, bounds.height * 0.08))
      if (clientY < bounds.top + edge) {
        const pressure = Math.min(1, (bounds.top + edge - clientY) / edge)
        return -Math.max(2, Math.round(22 * pressure))
      }
      if (clientY > bounds.bottom - edge) {
        const pressure = Math.min(1, (clientY - (bounds.bottom - edge)) / edge)
        return Math.max(2, Math.round(22 * pressure))
      }
      return 0
    }

    const stopAutoScroll = () => {
      window.cancelAnimationFrame(autoScrollFrame)
      autoScrollFrame = 0
    }

    const runAutoScroll = () => {
      if (!gesture) {
        stopAutoScroll()
        return
      }
      const amount = autoScrollAmount(gesture.clientY)
      if (!amount) {
        stopAutoScroll()
        return
      }
      const before = root.scrollTop
      root.scrollTop += amount
      if (root.scrollTop !== before) setFocusFromPoint(gesture.clientX, gesture.clientY)
      autoScrollFrame = window.requestAnimationFrame(runAutoScroll)
    }

    const scheduleAutoScroll = () => {
      if (!gesture || !autoScrollAmount(gesture.clientY)) {
        stopAutoScroll()
        return
      }
      if (!autoScrollFrame) autoScrollFrame = window.requestAnimationFrame(runAutoScroll)
    }

    const handlePointerDown = (event: PointerEvent) => {
      if (event.button !== 0 || !event.isPrimary || event.pointerType === 'touch') return
      const target = event.target
      if (!(target instanceof Element) || !root.contains(target)
        || !target.closest('[data-text-item="true"], .text-layer')) return
      // event.detail is unreliable once pointerdown is prevented, so count
      // multi-clicks manually: same button, close in time and position.
      const clickCount = lastPress
        && lastPress.button === event.button
        && event.timeStamp - lastPress.time <= 500
        && Math.hypot(event.clientX - lastPress.x, event.clientY - lastPress.y) <= 4
        ? Math.min(3, lastPress.clickCount + 1)
        : 1
      lastPress = { time: event.timeStamp, x: event.clientX, y: event.clientY, button: event.button, clickCount }
      const selection = window.getSelection()
      if (!selection) return
      const caret = textCaretAtPoint(root, event.clientX, event.clientY)
      if (!caret) {
        // Never yield to Chromium's native line-inference drag while a text
        // tool is active; a missed press behaves like a click on whitespace.
        event.preventDefault()
        // Also clears a Select All, whose ends sit on the page container.
        if (!event.shiftKey && !selection.isCollapsed
          && (selectionNodeIsPdfText(root, selection.anchorNode) || selectionNodeIsPdfText(root, selection.focusNode)
            || Boolean(selection.anchorNode && root.contains(selection.anchorNode)))) {
          selection.removeAllRanges()
        }
        return
      }

      let mode: SelectionMode = 'character'
      let anchorStart: SelectionPoint = { node: caret.node, offset: caret.offset }
      let anchorEnd = anchorStart
      if (clickCount === 2) {
        const word = wordRangeAtPoint(caret, event.clientX, event.clientY)
        if (word) {
          mode = 'word'
          anchorStart = word.start
          anchorEnd = word.end
        }
      } else if (clickCount >= 3) {
        const line = lineRangeForSpan(caret.span)
        if (line) {
          mode = 'line'
          anchorStart = line.start
          anchorEnd = line.end
        }
      } else if (event.shiftKey && selection.anchorNode instanceof Text
        && selectionNodeIsPdfText(root, selection.anchorNode) && selection.anchorNode.isConnected) {
        anchorStart = {
          node: selection.anchorNode,
          offset: Math.max(0, Math.min(selection.anchorNode.length, selection.anchorOffset)),
        }
        anchorEnd = anchorStart
      }

      event.preventDefault()
      try {
        applySelection(selection, mode, anchorStart, anchorEnd, caret, event.clientX, event.clientY)
      } catch {
        return
      }
      pointerSelectionActiveRef.current = true
      gesture = {
        pointerId: event.pointerId,
        mode,
        anchorStart,
        anchorEnd,
        focusNode: caret.node,
        focusOffset: caret.offset,
        focusSpan: caret.span,
        startX: event.clientX,
        startY: event.clientY,
        clientX: event.clientX,
        clientY: event.clientY,
      }
      root.dataset.pointerSelectingText = 'true'
    }

    const handlePointerMove = (event: PointerEvent) => {
      if (!gesture || event.pointerId !== gesture.pointerId) return
      event.preventDefault()
      gesture.clientX = event.clientX
      gesture.clientY = event.clientY
      if (Math.hypot(event.clientX - gesture.startX, event.clientY - gesture.startY) >= 1) {
        setFocusFromPoint(event.clientX, event.clientY)
      }
      scheduleAutoScroll()
    }

    const finishPointerSelection = (event: PointerEvent, updateFocus: boolean) => {
      if (!gesture || event.pointerId !== gesture.pointerId) return
      if (updateFocus) setFocusFromPoint(event.clientX, event.clientY)
      gesture = null
      delete root.dataset.pointerSelectingText
      stopAutoScroll()
    }

    const handlePointerUp = (event: PointerEvent) => finishPointerSelection(event, true)
    const handlePointerCancel = (event: PointerEvent) => finishPointerSelection(event, false)

    root.addEventListener('pointerdown', handlePointerDown, true)
    document.addEventListener('pointermove', handlePointerMove, true)
    document.addEventListener('pointerup', handlePointerUp, true)
    document.addEventListener('pointercancel', handlePointerCancel, true)
    return () => {
      root.removeEventListener('pointerdown', handlePointerDown, true)
      document.removeEventListener('pointermove', handlePointerMove, true)
      document.removeEventListener('pointerup', handlePointerUp, true)
      document.removeEventListener('pointercancel', handlePointerCancel, true)
      gesture = null
      delete root.dataset.pointerSelectingText
      stopAutoScroll()
    }
  }, [pdf, tool, viewerRef])

  useEffect(() => {
    const root = viewerRef.current
    if (!root || typeof IntersectionObserver === 'undefined') {
      setMountedPages(new Set(Array.from({ length: Math.min(pdf.numPages, 5) }, (_, index) => index)))
      return
    }

    const renderObserver = new IntersectionObserver((entries) => {
      let protectedPages: Set<number> | null = null
      setMountedPages((current) => {
        const next = new Set(current)
        for (const entry of entries) {
          const index = Number((entry.target as HTMLElement).dataset.pageIndex)
          if (!Number.isInteger(index)) continue
          if (entry.isIntersecting) {
            renderIntersectingPagesRef.current.add(index)
            next.add(index)
          } else {
            renderIntersectingPagesRef.current.delete(index)
            // Removing nodes between a native Selection's anchor and focus
            // makes Chromium repaint the grey selection at a different range,
            // so a live drag keeps every traversed page and an idle selection
            // keeps its own page span (bounded). Everything else is released.
            if (pointerSelectionActiveRef.current
              || index === currentPageRef.current
              || pinnedPagesRef.current.includes(index)) continue
            protectedPages ??= selectionProtectedPages(root)
            if (!protectedPages.has(index)) next.delete(index)
          }
        }
        return next.size === current.size && [...next].every((index) => current.has(index)) ? current : next
      })
    }, { root, rootMargin: '1100px 0px', threshold: 0 })

    const evaluateCurrentPage = (onlyWhenCurrentHidden = false) => {
      if (pointerSelectionActiveRef.current) return
      if (performance.now() < suppressObserverUntilRef.current) return
      if (onlyWhenCurrentHidden && slotIntersectsViewport(root, slotsRef.current.get(currentPageRef.current))) return
      const visible = [...visiblePagesRef.current]
        .map((index) => slotsRef.current.get(index))
        .filter((slot): slot is HTMLElement => Boolean(slot))
      const index = mostVisiblePage(root, visible.length ? visible : [...slotsRef.current.values()], pdf.numPages)
      if (index === undefined || !Number.isInteger(index) || index === currentPageRef.current) return
      const now = performance.now()
      for (const [reportedIndex, reportedAt] of reportedFromScrollRef.current) {
        if (now - reportedAt > 1500) reportedFromScrollRef.current.delete(reportedIndex)
      }
      reportedFromScrollRef.current.set(index, now)
      onCurrentPage(index)
    }
    evaluateCurrentPageRef.current = evaluateCurrentPage

    const visibleObserver = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        const index = Number((entry.target as HTMLElement).dataset.pageIndex)
        if (!Number.isInteger(index)) continue
        if (entry.isIntersecting) visiblePagesRef.current.add(index)
        else visiblePagesRef.current.delete(index)
      }
      evaluateCurrentPage()
    }, { root, threshold: 0 })

    // Scrolling within the same set of visible pages changes which one
    // dominates without any intersection change, so track scroll as well.
    let scrollFrame = 0
    const handleScroll = () => {
      if (scrollFrame) return
      scrollFrame = window.requestAnimationFrame(() => {
        scrollFrame = 0
        evaluateCurrentPage()
      })
    }
    root.addEventListener('scroll', handleScroll, { passive: true })

    for (const slot of slotsRef.current.values()) {
      renderObserver.observe(slot)
      visibleObserver.observe(slot)
    }
    return () => {
      window.cancelAnimationFrame(scrollFrame)
      root.removeEventListener('scroll', handleScroll)
      renderObserver.disconnect()
      visibleObserver.disconnect()
      visiblePagesRef.current.clear()
      if (evaluateCurrentPageRef.current === evaluateCurrentPage) evaluateCurrentPageRef.current = null
    }
  }, [pdf, viewerRef, onCurrentPage])

  useEffect(() => {
    setMountedPages((current) => {
      if (current.has(currentPage)) return current
      const next = new Set(current)
      next.add(currentPage)
      return next
    })
    const now = performance.now()
    const reportedAt = reportedFromScrollRef.current.get(currentPage)
    if (reportedAt !== undefined) {
      reportedFromScrollRef.current.delete(currentPage)
      if (now - reportedAt <= 1500) return
    }
    const root = viewerRef.current
    const quietAt = quietPageChangesRef.current.get(currentPage)
    quietPageChangesRef.current.clear()
    // Editing, signing or selecting on a page that is already on screen only
    // makes it current; scrolling would move the clicked spot away.
    if (quietAt !== undefined && now - quietAt <= QUIET_PAGE_CHANGE_MS
      && root && slotIntersectsViewport(root, slotsRef.current.get(currentPage))) return
    stopLayoutAnchorRef.current?.()
    suppressPageTrackingRef.current(NAVIGATION_SUPPRESS_MS)
    // Navigation shows the page from its top, as Acrobat does; centring hid
    // the heading of every page taller than the window.
    const alignPage = () => {
      if (pointerSelectionActiveRef.current) return
      const viewer = viewerRef.current
      const slot = slotsRef.current.get(currentPage)
      if (!viewer || !slot) return
      if (currentPage === 0) {
        viewer.scrollTop = 0
        return
      }
      const offset = slot.getBoundingClientRect().top - viewer.getBoundingClientRect().top
      if (Math.abs(offset) >= 1) viewer.scrollTop += offset
    }
    alignPage()
    // A search hit centres itself once its geometry is ready (PdfPage); a
    // second, late page alignment would push hits near page edges off screen.
    if (propsRef.current.activeSearchMatch?.pageIndex === currentPage) return
    // Re-align once placeholders above have settled, unless the user has
    // already moved on (wheel, keys such as a quick second PageDown, clicks).
    const settleTimer = window.setTimeout(alignPage, 220)
    const cancelSettle = () => window.clearTimeout(settleTimer)
    root?.addEventListener('wheel', cancelSettle, { passive: true })
    root?.addEventListener('pointerdown', cancelSettle, true)
    window.addEventListener('keydown', cancelSettle, true)
    return () => {
      cancelSettle()
      root?.removeEventListener('wheel', cancelSettle)
      root?.removeEventListener('pointerdown', cancelSettle, true)
      window.removeEventListener('keydown', cancelSettle, true)
    }
  }, [currentPage])

  const pageIndices = useMemo(() => Array.from({ length: pdf.numPages }, (_, index) => index), [pdf])

  // PdfPage is memoized. Route every callback through refs with a stable
  // identity so mountedPages/currentPage/dropTarget churn (and parent
  // re-renders) only re-render the slots whose data actually changed.
  const markQuietPageChange = (pageIndex: number) => {
    if (pageIndex !== currentPageRef.current) quietPageChangesRef.current.set(pageIndex, performance.now())
  }
  const markQuietPageChangeRef = useRef(markQuietPageChange)
  markQuietPageChangeRef.current = markQuietPageChange

  const sharedHandlers = useMemo(() => ({
    onRequestTextEdit: (edit: PageTextEdit) => {
      markQuietPageChangeRef.current(edit.pageIndex)
      propsRef.current.onRequestTextEdit(edit)
    },
    onTextEditChange: (edit: PageTextEdit) => propsRef.current.onTextEditChange(edit),
    onCommitTextEdit: () => propsRef.current.onCommitTextEdit(),
    onCancelTextEdit: () => propsRef.current.onCancelTextEdit(),
    onRequestObjectEdit: (edit: PageObjectEdit) => {
      markQuietPageChangeRef.current(edit.pageIndex)
      propsRef.current.onRequestObjectEdit(edit)
    },
    onObjectEditChange: (edit: PageObjectEdit) => propsRef.current.onObjectEditChange(edit),
    onCommitObjectEdit: () => propsRef.current.onCommitObjectEdit(),
    onCancelObjectEdit: () => propsRef.current.onCancelObjectEdit(),
    onObjectRegionSelected: () => propsRef.current.onObjectRegionSelected(),
    onNavigate: (pageIndex: number) => propsRef.current.onNavigate(pageIndex),
    onFormChange: (name: string, value: string | boolean) => propsRef.current.onFormChange(name, value),
    onRequestOcrOffer: (pageIndex: number, point: { x: number; y: number }, client: { x: number; y: number }) => {
      markQuietPageChangeRef.current(pageIndex)
      propsRef.current.onRequestOcrOffer?.(pageIndex, point, client)
    },
    onPendingEditAtHandled: (token: number) => propsRef.current.onPendingEditAtHandled?.(token),
    onScanEditPrepared: (key: string, result: ScanPreparation) => propsRef.current.onScanEditPrepared?.(key, result),
  }), [])

  const pageHandlersFor = (index: number): PageScopedHandlers => {
    let handlers = pageHandlersRef.current.get(index)
    if (!handlers) {
      handlers = {
        onPageReady: (nextSize: PageSize) => {
          const scale = propsRef.current.zoom
          setPageSizes((current) => {
            const scaled = { width: nextSize.width / scale, height: nextSize.height / scale }
            const previous = current[index]
            if (previous && Math.abs(previous.width - scaled.width) < 0.5 && Math.abs(previous.height - scaled.height) < 0.5) return current
            return { ...current, [index]: scaled }
          })
        },
        onHighlight: (rects) => propsRef.current.onHighlight(index, rects),
        onTextMarkup: (style, rects, displayRotation) => propsRef.current.onTextMarkup(index, style, rects, displayRotation),
        onInk: (points) => propsRef.current.onInk(index, points),
        onRectangle: (rect) => propsRef.current.onRectangle(index, rect),
        onCrop: (rect) => propsRef.current.onCrop(index, rect),
        onPlaceSignature: (point, displayRotation) => {
          markQuietPageChangeRef.current(index)
          propsRef.current.onPlaceSignature(index, point, displayRotation)
        },
      }
      pageHandlersRef.current.set(index, handlers)
    }
    return handlers
  }

  const unzoomedSize = (index: number): PageSize => {
    const cached = pageSizes[index]
    if (cached) return cached
    const base = basePageSizes[index]
    if (!base) return estimatedSize
    return (rotations[index] || 0) % 180 ? { width: base.height, height: base.width } : base
  }

  return (
    <div className="continuous-pages" aria-label={`${pdf.numPages}-page continuous document`}>
      {pageIndices.map((index) => {
        const scaleOneSize = unzoomedSize(index)
        const size = { width: scaleOneSize.width * zoom, height: scaleOneSize.height * zoom }
        const pageHandlers = pageHandlersFor(index)
        const forceMounted = index === currentPage
          || index === textEdit?.pageIndex
          || index === objectEdit?.pageIndex
          || index === activeSearchMatch?.pageIndex
        const mounted = forceMounted || mountedPages.has(index)
        const selected = selectedPages.has(index)
        const dropEdge = dropTarget?.pageIndex === index ? dropTarget.edge : null
        return (
          <section
            key={index}
            ref={(element) => {
              if (element) slotsRef.current.set(index, element)
              else slotsRef.current.delete(index)
            }}
            className={`continuous-page-slot${index === currentPage ? ' is-current' : ''}${selected ? ' is-selected' : ''}`}
            data-page-index={index}
            data-drop-edge={dropEdge || undefined}
            aria-label={`Page ${index + 1}`}
            aria-selected={selected}
            style={{ minHeight: size.height + PAGE_VERTICAL_CHROME }}
            onDragOver={(event) => {
              if (!Array.from(event.dataTransfer.types).includes('Files')) return
              event.preventDefault()
              event.stopPropagation()
              event.dataTransfer.dropEffect = 'copy'
              const bounds = event.currentTarget.getBoundingClientRect()
              const edge = pageDropEdge(event.clientY, bounds.top, bounds.height)
              setDropTarget((current) => current?.pageIndex === index && current.edge === edge ? current : { pageIndex: index, edge })
            }}
            onDragLeave={(event) => {
              if (event.currentTarget.contains(event.relatedTarget as Node | null)) return
              setDropTarget((current) => current?.pageIndex === index ? null : current)
            }}
            onDrop={(event) => {
              if (!Array.from(event.dataTransfer.types).includes('Files')) return
              event.preventDefault()
              event.stopPropagation()
              const bounds = event.currentTarget.getBoundingClientRect()
              const edge = pageDropEdge(event.clientY, bounds.top, bounds.height)
              const files = Array.from(event.dataTransfer.files).filter(isImportableTransferFile)
              setDropTarget(null)
              if (event.dataTransfer.files.length) onImportPagesAt(files, pageDropInsertIndex(index, edge))
            }}
          >
            <button
              type="button"
              className="page-transfer-handle"
              draggable
              aria-pressed={selected}
              aria-label={`${selected ? 'Selected ' : ''}page ${index + 1}. Click to select; drag to export${selectedPages.size > 1 && selected ? ` ${selectedPages.size} selected pages` : ''}.`}
              title={selectedPages.size > 1 && selected ? `Drag ${selectedPages.size} selected pages as one PDF` : `Drag page ${index + 1} as a PDF`}
              onPointerDown={(event) => event.stopPropagation()}
              onClick={(event) => {
                event.stopPropagation()
                markQuietPageChange(index)
                onSelectPage(index, event)
              }}
              onDragStart={(event) => {
                event.stopPropagation()
                event.dataTransfer.effectAllowed = 'copy'
                onPageDragStart(index, event)
              }}
            >
              <GripVertical size={12} />
              <span>Page {index + 1}</span>
              {selectedPages.size > 1 && selected && <strong>{selectedPages.size}</strong>}
            </button>
            {dropEdge && <div className={`page-drop-indicator is-${dropEdge}`}>Drop to insert {dropEdge} page {index + 1}</div>}
            {mounted ? (
              <PdfPage
                pdf={pdf}
                pageIndex={index}
                zoom={zoom}
                rotation={rotations[index] || 0}
                tool={tool}
                overlays={overlays}
                formValues={formValues}
                textEdit={textEdit}
                objectEdit={objectEdit}
                activeSearchMatch={activeSearchMatch}
                selectingObjectRegion={selectingObjectRegion && currentPage === index}
                onPageReady={pageHandlers.onPageReady}
                onRequestTextEdit={sharedHandlers.onRequestTextEdit}
                onTextEditChange={sharedHandlers.onTextEditChange}
                onCommitTextEdit={sharedHandlers.onCommitTextEdit}
                onCancelTextEdit={sharedHandlers.onCancelTextEdit}
                onRequestObjectEdit={sharedHandlers.onRequestObjectEdit}
                onObjectEditChange={sharedHandlers.onObjectEditChange}
                onCommitObjectEdit={sharedHandlers.onCommitObjectEdit}
                onCancelObjectEdit={sharedHandlers.onCancelObjectEdit}
                onObjectRegionSelected={sharedHandlers.onObjectRegionSelected}
                onHighlight={pageHandlers.onHighlight}
                onTextMarkup={pageHandlers.onTextMarkup}
                onInk={pageHandlers.onInk}
                onRectangle={pageHandlers.onRectangle}
                onCrop={pageHandlers.onCrop}
                onPlaceSignature={pageHandlers.onPlaceSignature}
                onNavigate={sharedHandlers.onNavigate}
                onFormChange={sharedHandlers.onFormChange}
                onRequestOcrOffer={sharedHandlers.onRequestOcrOffer}
                pendingEditAt={pendingEditAt?.pageIndex === index ? pendingEditAt : null}
                onPendingEditAtHandled={sharedHandlers.onPendingEditAtHandled}
                onScanEditPrepared={sharedHandlers.onScanEditPrepared}
              />
            ) : (
              <div className="continuous-page-placeholder" style={{ width: size.width, height: size.height }} aria-hidden="true" />
            )}
          </section>
        )
      })}
    </div>
  )
}
