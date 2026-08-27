import { useEffect, useMemo, useRef, useState } from 'react'
import type { DragEvent as ReactDragEvent, RefObject } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { GripVertical } from 'lucide-react'
import type { ActiveSearchMatch, PageObjectEdit, PageTextEdit, PdfOverlay, PdfRect, ToolMode } from '../types'
import { isPdfTransferFile, pageDropEdge, pageDropInsertIndex, type PageDropEdge } from '../lib/pageTransfer'
import { textCaretAtPoint } from '../lib/textSelection'
import { PdfPage } from './PdfPage'

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
  onTextMarkup: (pageIndex: number, style: 'underline' | 'strikeout', rects: PdfRect[]) => void
  onInk: (pageIndex: number, points: Array<{ x: number; y: number }>) => void
  onRectangle: (pageIndex: number, rect: PdfRect) => void
  onCrop: (pageIndex: number, rect: PdfRect) => void
  onNavigate: (pageIndex: number) => void
  onFormChange: (name: string, value: string | boolean) => void
}

const PAGE_VERTICAL_CHROME = 40
const POINTER_TEXT_TOOLS: ToolMode[] = ['select', 'edit', 'highlight', 'underline', 'strikeout']

function selectionNodeIsPdfText(root: HTMLElement, node: Node | null) {
  if (!node || !root.contains(node)) return false
  const element = node instanceof Element ? node : node.parentElement
  return Boolean(element?.closest('.text-layer'))
}

function closestCenteredPage(
  entries: Map<number, IntersectionObserverEntry>,
  root: HTMLElement,
) {
  const rootCenter = root.getBoundingClientRect().top + root.clientHeight / 2
  return [...entries.entries()]
    .map(([index, entry]) => {
      const bounds = entry.target.getBoundingClientRect()
      return { index, distance: Math.abs(bounds.top + bounds.height / 2 - rootCenter) }
    })
    .sort((a, b) => a.distance - b.distance)[0]?.index
}

export function ContinuousPdfViewer(props: ContinuousPdfViewerProps) {
  const {
    pdf, viewerRef, currentPage, selectedPages, zoom, rotations, tool, overlays, formValues,
    textEdit, objectEdit, activeSearchMatch, selectingObjectRegion, onCurrentPage,
    onSelectPage, onPageDragStart, onImportPagesAt,
    onRequestTextEdit, onTextEditChange, onCommitTextEdit, onCancelTextEdit,
    onRequestObjectEdit, onObjectEditChange, onCommitObjectEdit, onCancelObjectEdit,
    onObjectRegionSelected, onHighlight, onTextMarkup, onInk, onRectangle,
    onCrop, onNavigate, onFormChange,
  } = props
  const [mountedPages, setMountedPages] = useState<Set<number>>(() => new Set([currentPage]))
  const [pageSizes, setPageSizes] = useState<Record<number, PageSize>>({})
  const [estimatedSize, setEstimatedSize] = useState<PageSize>({ width: 612 * zoom, height: 792 * zoom })
  const [dropTarget, setDropTarget] = useState<{ pageIndex: number; edge: PageDropEdge } | null>(null)
  const slotsRef = useRef(new Map<number, HTMLElement>())
  const reportedFromScrollRef = useRef(new Map<number, number>())
  const currentPageRef = useRef(currentPage)
  const activeEditPagesRef = useRef<Array<number | undefined>>([])
  const previousRotationsRef = useRef<Record<number, number>>(rotations)
  const centeredEntriesRef = useRef(new Map<number, IntersectionObserverEntry>())
  const renderIntersectingPagesRef = useRef(new Set<number>())
  const nativeTextSelectionActiveRef = useRef(false)
  const suppressObserverUntilRef = useRef(0)

  currentPageRef.current = currentPage
  activeEditPagesRef.current = [textEdit?.pageIndex, objectEdit?.pageIndex]

  useEffect(() => {
    let cancelled = false
    pdf.getPage(1).then((page) => {
      if (cancelled) return
      const angle = (((page.rotate || 0) + (rotations[0] || 0)) % 360 + 360) % 360
      const viewport = page.getViewport({ scale: zoom, rotation: angle })
      setEstimatedSize({ width: viewport.width, height: viewport.height })
    }).catch(() => {})
    return () => { cancelled = true }
  }, [pdf, zoom, rotations[0]])

  useEffect(() => {
    setPageSizes({})
  }, [zoom])

  useEffect(() => {
    const previous = previousRotationsRef.current
    const changed = new Set([...Object.keys(previous), ...Object.keys(rotations)]
      .map(Number)
      .filter((index) => (previous[index] || 0) !== (rotations[index] || 0)))
    previousRotationsRef.current = rotations
    if (!changed.size) return
    setPageSizes((current) => Object.fromEntries(Object.entries(current).filter(([rawIndex]) => !changed.has(Number(rawIndex)))))
  }, [rotations])

  useEffect(() => {
    setMountedPages(new Set([currentPage]))
    setPageSizes({})
    centeredEntriesRef.current.clear()
    renderIntersectingPagesRef.current.clear()
    reportedFromScrollRef.current.clear()
    nativeTextSelectionActiveRef.current = false
  }, [pdf])

  useEffect(() => {
    const root = viewerRef.current
    if (!root) return
    let selectingWithPointer = false
    let settleFrame = 0

    const currentSelectionBelongsToViewer = () => {
      const selection = window.getSelection()
      return Boolean(selection && !selection.isCollapsed && selection.rangeCount
        && (selectionNodeIsPdfText(root, selection.anchorNode) || selectionNodeIsPdfText(root, selection.focusNode)))
    }

    const finishSelection = () => {
      if (!nativeTextSelectionActiveRef.current) return
      nativeTextSelectionActiveRef.current = false
      const centeredIndex = closestCenteredPage(centeredEntriesRef.current, root)
      const pageToKeep = centeredIndex ?? currentPageRef.current
      setMountedPages((current) => {
        const next = new Set([...current].filter((index) => (
          renderIntersectingPagesRef.current.has(index)
          || index === pageToKeep
          || activeEditPagesRef.current.includes(index)
        )))
        if (!next.size) next.add(pageToKeep)
        return next.size === current.size && [...next].every((index) => current.has(index)) ? current : next
      })
      if (centeredIndex !== undefined && centeredIndex !== currentPageRef.current) {
        reportedFromScrollRef.current.set(centeredIndex, performance.now())
        onCurrentPage(centeredIndex)
      }
    }

    const syncSelectionState = () => {
      if (selectingWithPointer || currentSelectionBelongsToViewer()) {
        nativeTextSelectionActiveRef.current = true
      } else {
        finishSelection()
      }
    }

    const handlePointerDown = (event: PointerEvent) => {
      const target = event.target
      selectingWithPointer = target instanceof Element
        && root.contains(target)
        && Boolean(target.closest('.text-layer'))
      if (selectingWithPointer) nativeTextSelectionActiveRef.current = true
    }
    const handlePointerFinished = () => {
      selectingWithPointer = false
      window.cancelAnimationFrame(settleFrame)
      settleFrame = window.requestAnimationFrame(syncSelectionState)
    }

    document.addEventListener('selectionchange', syncSelectionState)
    document.addEventListener('pointerdown', handlePointerDown, true)
    document.addEventListener('pointerup', handlePointerFinished, true)
    document.addEventListener('pointercancel', handlePointerFinished, true)
    return () => {
      window.cancelAnimationFrame(settleFrame)
      document.removeEventListener('selectionchange', syncSelectionState)
      document.removeEventListener('pointerdown', handlePointerDown, true)
      document.removeEventListener('pointerup', handlePointerFinished, true)
      document.removeEventListener('pointercancel', handlePointerFinished, true)
    }
  }, [pdf, viewerRef, onCurrentPage])

  useEffect(() => {
    const root = viewerRef.current
    if (!root || !POINTER_TEXT_TOOLS.includes(tool)) return

    type SelectionGesture = {
      pointerId: number
      anchorNode: Text
      anchorOffset: number
      focusNode: Text
      focusOffset: number
      startX: number
      startY: number
      clientX: number
      clientY: number
    }

    let gesture: SelectionGesture | null = null
    let autoScrollFrame = 0

    const setFocusFromPoint = (clientX: number, clientY: number) => {
      if (!gesture) return false
      const caret = textCaretAtPoint(root, clientX, clientY)
      if (!caret) return false
      if (caret.node === gesture.focusNode && caret.offset === gesture.focusOffset) return true
      const selection = window.getSelection()
      if (!selection || !gesture.anchorNode.isConnected || !caret.node.isConnected) return false
      try {
        selection.setBaseAndExtent(gesture.anchorNode, gesture.anchorOffset, caret.node, caret.offset)
        gesture.focusNode = caret.node
        gesture.focusOffset = caret.offset
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
      if (event.button !== 0 || !event.isPrimary || event.pointerType === 'touch' || event.detail > 1) return
      const target = event.target
      if (!(target instanceof Element) || !root.contains(target) || !target.closest('[data-text-item="true"]')) return
      const caret = textCaretAtPoint(root, event.clientX, event.clientY)
      if (!caret) return
      const selection = window.getSelection()
      if (!selection) return

      let anchorNode = caret.node
      let anchorOffset = caret.offset
      if (event.shiftKey && selection.anchorNode instanceof Text
        && selectionNodeIsPdfText(root, selection.anchorNode) && selection.anchorNode.isConnected) {
        anchorNode = selection.anchorNode
        anchorOffset = Math.max(0, Math.min(anchorNode.length, selection.anchorOffset))
      }

      event.preventDefault()
      try {
        selection.setBaseAndExtent(anchorNode, anchorOffset, caret.node, caret.offset)
      } catch {
        return
      }
      nativeTextSelectionActiveRef.current = true
      gesture = {
        pointerId: event.pointerId,
        anchorNode,
        anchorOffset,
        focusNode: caret.node,
        focusOffset: caret.offset,
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
            // Removing any node between a native Selection's anchor and focus
            // makes Chromium repaint the grey selection at a different range
            // (or collapse it completely). Keep every traversed page mounted
            // until the user clears that selection, then prune in one pass.
            if (!nativeTextSelectionActiveRef.current
              && index !== currentPageRef.current
              && !activeEditPagesRef.current.includes(index)) next.delete(index)
          }
        }
        return next.size === current.size && [...next].every((index) => current.has(index)) ? current : next
      })
    }, { root, rootMargin: '1100px 0px', threshold: 0 })

    const currentObserver = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        const index = Number((entry.target as HTMLElement).dataset.pageIndex)
        if (!Number.isInteger(index)) continue
        if (entry.isIntersecting) centeredEntriesRef.current.set(index, entry)
        else centeredEntriesRef.current.delete(index)
      }
      const index = closestCenteredPage(centeredEntriesRef.current, root)
      if (index === undefined) return
      if (nativeTextSelectionActiveRef.current) return
      if (performance.now() < suppressObserverUntilRef.current) return
      if (!Number.isInteger(index) || index === currentPageRef.current) return
      const now = performance.now()
      for (const [reportedIndex, reportedAt] of reportedFromScrollRef.current) {
        if (now - reportedAt > 1500) reportedFromScrollRef.current.delete(reportedIndex)
      }
      reportedFromScrollRef.current.set(index, now)
      onCurrentPage(index)
    }, { root, rootMargin: '-42% 0px -42% 0px', threshold: 0 })

    for (const slot of slotsRef.current.values()) {
      renderObserver.observe(slot)
      currentObserver.observe(slot)
    }
    return () => {
      renderObserver.disconnect()
      currentObserver.disconnect()
      centeredEntriesRef.current.clear()
    }
  }, [pdf, viewerRef, onCurrentPage])

  useEffect(() => {
    setMountedPages((current) => {
      if (current.has(currentPage)) return current
      const next = new Set(current)
      next.add(currentPage)
      return next
    })
    const reportedAt = reportedFromScrollRef.current.get(currentPage)
    if (reportedAt !== undefined) {
      reportedFromScrollRef.current.delete(currentPage)
      if (performance.now() - reportedAt <= 1500) return
    }
    suppressObserverUntilRef.current = performance.now() + 700
    const centerPage = () => {
      if (nativeTextSelectionActiveRef.current) return
      slotsRef.current.get(currentPage)?.scrollIntoView({ block: 'center', behavior: 'auto' })
    }
    centerPage()
    const settleTimer = window.setTimeout(centerPage, 220)
    return () => window.clearTimeout(settleTimer)
  }, [currentPage, pdf])

  const pageIndices = useMemo(() => Array.from({ length: pdf.numPages }, (_, index) => index), [pdf])

  return (
    <div className="continuous-pages" aria-label={`${pdf.numPages}-page continuous document`}>
      {pageIndices.map((index) => {
        const size = pageSizes[index] || estimatedSize
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
              const files = Array.from(event.dataTransfer.files).filter(isPdfTransferFile)
              setDropTarget(null)
              if (files.length) onImportPagesAt(files, pageDropInsertIndex(index, edge))
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
                onPageReady={(nextSize) => setPageSizes((current) => {
                  const previous = current[index]
                  if (previous && Math.abs(previous.width - nextSize.width) < 0.5 && Math.abs(previous.height - nextSize.height) < 0.5) return current
                  return { ...current, [index]: nextSize }
                })}
                onRequestTextEdit={onRequestTextEdit}
                onTextEditChange={onTextEditChange}
                onCommitTextEdit={onCommitTextEdit}
                onCancelTextEdit={onCancelTextEdit}
                onRequestObjectEdit={onRequestObjectEdit}
                onObjectEditChange={onObjectEditChange}
                onCommitObjectEdit={onCommitObjectEdit}
                onCancelObjectEdit={onCancelObjectEdit}
                onObjectRegionSelected={onObjectRegionSelected}
                onHighlight={(rects) => onHighlight(index, rects)}
                onTextMarkup={(style, rects) => onTextMarkup(index, style, rects)}
                onInk={(points) => onInk(index, points)}
                onRectangle={(rect) => onRectangle(index, rect)}
                onCrop={(rect) => onCrop(index, rect)}
                onNavigate={onNavigate}
                onFormChange={onFormChange}
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
