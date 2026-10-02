import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { DragEvent as ReactDragEvent, KeyboardEvent as ReactKeyboardEvent, MouseEvent } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  Bookmark as BookmarkIcon,
  CopyPlus,
  Download,
  ExternalLink,
  File,
  FilePlus2,
  GripVertical,
  MoreHorizontal,
  PanelLeftClose,
  Pencil,
  RotateCcw,
  RotateCw,
  ScanText,
  Search,
  Trash2,
  X,
} from 'lucide-react'
import type { ActiveSearchMatch, Bookmark, SearchResult } from '../types'
import { bookmarkTargetLabel } from '../lib/bookmarks'
import { getPageTextContent } from '../lib/pdf'
import { isImportableTransferFile, pageIndicesForTransfer } from '../lib/pageTransfer'
import { normalizeSearchValue } from '../lib/search'
import { createDocumentSearch } from '../lib/documentSearch'
import { cx, errorMessage } from '../lib/utils'
import { EmptyState, IconButton } from './ui'

type SidebarTab = 'pages' | 'bookmarks' | 'search'

const THUMBNAIL_ROW_HEIGHT = 222
const THUMBNAIL_OVERSCAN = 2
const INTERNAL_PAGE_DRAG_TYPE = 'application/x-simple-pdf-page'

// Pages the search index read without any text: scans that may need recognising.
const pagesWithoutText = new WeakMap<PDFDocumentProxy, Set<number>>()

const searchDocument = createDocumentSearch<PDFDocumentProxy>(async (pdf, pageIndex, isCancelled) => {
  if (isCancelled()) return null
  const page = await pdf.getPage(pageIndex + 1)
  if (isCancelled()) return null
  const content = await getPageTextContent(page)
  if (isCancelled()) return null
  const text = content.items.map((item) => ('str' in item ? item.str : '')).join(' ')
  let empty = pagesWithoutText.get(pdf)
  if (!empty) pagesWithoutText.set(pdf, empty = new Set())
  if (text.trim()) empty.delete(pageIndex)
  else empty.add(pageIndex)
  return text
})

interface SearchOccurrence {
  resultIndex: number
  pageIndex: number
  occurrenceIndex: number
}

function flattenSearchResults(results: SearchResult[]): SearchOccurrence[] {
  return results.flatMap((result, resultIndex) => (
    Array.from({ length: result.count }, (_, occurrenceIndex) => ({
      resultIndex,
      pageIndex: result.pageIndex,
      occurrenceIndex,
    }))
  ))
}

function findLastOccurrenceAtOrBefore(occurrences: SearchOccurrence[], pageIndex: number) {
  for (let index = occurrences.length - 1; index >= 0; index -= 1) {
    if (occurrences[index].pageIndex <= pageIndex) return index
  }
  return -1
}

interface PageThumbnailProps {
  pdf: PDFDocumentProxy
  index: number
  selected: boolean
  current: boolean
  bookmarked: boolean
  rotation: number
  canDelete: boolean
  exportPageCount: number
  onClick: (event: MouseEvent) => void
  onDragStart: (event: ReactDragEvent<HTMLDivElement>) => void
  onExportDragStart: (event: ReactDragEvent<HTMLButtonElement>) => void
  onDragEnd: () => void
  onDrop: (event: ReactDragEvent<HTMLDivElement>, insertIndex: number) => void
  onRotateLeft: () => void
  onRotateRight: () => void
  onDelete: () => void
}

function PageThumbnail({
  pdf, index, selected, current, bookmarked, rotation, canDelete, exportPageCount, onClick,
  onDragStart, onExportDragStart, onDragEnd, onDrop, onRotateLeft, onRotateRight, onDelete,
}: PageThumbnailProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const paperRef = useRef<HTMLDivElement>(null)
  const [error, setError] = useState(false)
  const [dropEdge, setDropEdge] = useState<'before' | 'after' | null>(null)

  useEffect(() => {
    let cancelled = false
    let renderTask: { cancel: () => void; promise: Promise<void> } | null = null
    let renderTimer = 0
    setError(false)
    if (canvasRef.current) {
      canvasRef.current.width = 1
      canvasRef.current.height = 1
    }
    async function render() {
      try {
        const page = await pdf.getPage(index + 1)
        if (cancelled || !canvasRef.current) return
        const pageRotation = (((page.rotate || 0) + rotation) % 360 + 360) % 360
        const base = page.getViewport({ scale: 1, rotation: pageRotation })
        const scale = Math.min(168 / base.width, 168 / base.height)
        const viewport = page.getViewport({ scale, rotation: pageRotation })
        const dpr = Math.min(window.devicePixelRatio || 1, 2)
        const canvas = canvasRef.current
        canvas.width = Math.floor(viewport.width * dpr)
        canvas.height = Math.floor(viewport.height * dpr)
        canvas.style.width = `${viewport.width}px`
        canvas.style.height = `${viewport.height}px`
        if (paperRef.current) {
          paperRef.current.style.width = `${viewport.width}px`
          paperRef.current.style.height = `${viewport.height}px`
        }
        const context = canvas.getContext('2d', { alpha: false })
        if (!context) return
        renderTask = page.render({
          canvasContext: context,
          viewport,
          transform: dpr === 1 ? undefined : [dpr, 0, 0, dpr, 0, 0],
        })
        await renderTask.promise
      } catch (renderError) {
        if (!cancelled && (renderError as Error)?.name !== 'RenderingCancelledException') setError(true)
      }
    }
    // Give the full-size page first access to pdf.js' worker during document open.
    renderTimer = window.setTimeout(render, 300)
    return () => {
      cancelled = true
      window.clearTimeout(renderTimer)
      renderTask?.cancel()
    }
  }, [pdf, index, rotation])

  return (
    <div
      className={cx('thumbnail-item', selected && 'is-selected', current && 'is-current')}
      role="option"
      aria-selected={selected}
      tabIndex={current ? 0 : -1}
      draggable
      onDragStart={onDragStart}
      onDragEnd={() => { setDropEdge(null); onDragEnd() }}
      onDragOver={(event) => {
        event.preventDefault()
        const externalFiles = Array.from(event.dataTransfer.types).includes('Files')
        event.dataTransfer.dropEffect = externalFiles ? 'copy' : 'move'
        const bounds = event.currentTarget.getBoundingClientRect()
        setDropEdge(event.clientY < bounds.top + bounds.height / 2 ? 'before' : 'after')
      }}
      onDragLeave={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setDropEdge(null)
      }}
      onDrop={(event) => {
        event.preventDefault()
        event.stopPropagation()
        const bounds = event.currentTarget.getBoundingClientRect()
        const insertIndex = event.clientY < bounds.top + bounds.height / 2 ? index : index + 1
        setDropEdge(null)
        onDrop(event, insertIndex)
      }}
      onClick={onClick}
      data-drop-edge={dropEdge || undefined}
    >
      <div className="thumbnail-paper" ref={paperRef}>
        {error ? <span className="thumbnail-error">Preview unavailable</span> : <canvas ref={canvasRef} />}
        {bookmarked && <span className="thumbnail-bookmark" title="Bookmarked"><BookmarkIcon size={11} fill="currentColor" /></span>}
        {selected && (
          <div className="thumbnail-actions" role="group" aria-label={`Actions for page ${index + 1}`}>
            <button
              type="button"
              title="Rotate page left"
              aria-label={`Rotate page ${index + 1} left`}
              onPointerDown={(event) => event.stopPropagation()}
              onClick={(event) => { event.stopPropagation(); onRotateLeft() }}
            ><RotateCcw size={13} strokeWidth={1.9} /></button>
            <button
              type="button"
              title="Rotate page right"
              aria-label={`Rotate page ${index + 1} right`}
              onPointerDown={(event) => event.stopPropagation()}
              onClick={(event) => { event.stopPropagation(); onRotateRight() }}
            ><RotateCw size={13} strokeWidth={1.9} /></button>
            <button
              type="button"
              className="thumbnail-delete"
              title={canDelete ? 'Delete page' : 'A PDF must keep one page'}
              aria-label={`Delete page ${index + 1}`}
              disabled={!canDelete}
              onPointerDown={(event) => event.stopPropagation()}
              onClick={(event) => { event.stopPropagation(); onDelete() }}
            ><Trash2 size={13} strokeWidth={1.9} /></button>
          </div>
        )}
      </div>
      <div className="thumbnail-meta">
        <button
          type="button"
          className="thumbnail-export-handle"
          draggable
          title={exportPageCount > 1 ? `Drag to export ${exportPageCount} selected pages as one PDF file` : `Drag to export page ${index + 1} as a PDF file`}
          aria-label={exportPageCount > 1 ? `Drag to export ${exportPageCount} selected pages as one PDF file` : `Drag to export page ${index + 1} as a PDF file`}
          onPointerDown={(event) => event.stopPropagation()}
          onClick={(event) => event.stopPropagation()}
          onDragStart={(event) => {
            event.stopPropagation()
            onExportDragStart(event)
          }}
        ><GripVertical size={13} /></button>
        <span>{index + 1}</span>
        {current && <small>Current</small>}
      </div>
    </div>
  )
}

interface SidebarProps {
  pdf: PDFDocumentProxy
  pageIndex: number
  selectedPages: Set<number>
  pageRotations: Record<number, number>
  bookmarks: Bookmark[]
  searchRequestId: number
  onActiveSearchMatch: (match: ActiveSearchMatch | null) => void
  onClose: () => void
  onPage: (index: number) => void
  onSelectPage: (index: number, event: MouseEvent) => void
  onReorder: (indices: number[], insertIndex: number) => void
  onInsertBlank: () => void
  onAddPages: () => void
  onDuplicatePages: () => void
  onRotateLeft: () => void
  onRotateRight: () => void
  onDeletePages: () => void
  onPageDragStart?: (index: number, event: ReactDragEvent<HTMLButtonElement>) => void
  onImportPagesAt?: (files: File[], insertIndex: number) => void
  onExportPages: () => void
  /** False when the outline could not be read (or the PDF is read-only): no rename or delete. */
  bookmarksEditable: boolean
  /** Go to the bookmark's page, or open its web link. */
  onOpenBookmark: (bookmark: Bookmark) => void
  onDeleteBookmark: (id: string) => void
  onRenameBookmark: (id: string, label: string) => void
  /** Opens Recognize text for the pages that need it. */
  onRecognizeText?: () => void
}

export function Sidebar({
  pdf, pageIndex, selectedPages, pageRotations, bookmarks, searchRequestId, onClose, onPage, onSelectPage,
  onReorder, onInsertBlank, onAddPages, onDuplicatePages, onRotateLeft, onRotateRight,
  onDeletePages, onPageDragStart, onImportPagesAt, onExportPages,
  bookmarksEditable, onOpenBookmark, onDeleteBookmark, onRenameBookmark, onActiveSearchMatch, onRecognizeText,
}: SidebarProps) {
  const [tab, setTab] = useState<SidebarTab>('pages')
  // Bookmark being renamed in place (Electron has no window.prompt).
  const [renamingBookmarkId, setRenamingBookmarkId] = useState<string | null>(null)
  const [renameDraft, setRenameDraft] = useState('')
  const renamingBookmarkRef = useRef<string | null>(null)
  const [draggedPage, setDraggedPage] = useState<number | null>(null)
  const draggedIndicesRef = useRef<number[]>([])
  const autoScrollRef = useRef({ velocity: 0, lastOver: 0, frame: 0 })
  const [query, setQuery] = useState('')
  const [searching, setSearching] = useState(false)
  const [searchError, setSearchError] = useState('')
  const [results, setResults] = useState<SearchResult[]>([])
  const [lastSearchedQuery, setLastSearchedQuery] = useState('')
  const [activeMatchIndex, setActiveMatchIndex] = useState(-1)
  const [textlessPages, setTextlessPages] = useState(0)
  const [pageMenuOpen, setPageMenuOpen] = useState(false)
  const currentRef = useRef<HTMLDivElement>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const lastNonSearchTab = useRef<Exclude<SidebarTab, 'search'>>('pages')
  const searchRunRef = useRef(0)
  const [thumbnailWindow, setThumbnailWindow] = useState(() => ({
    start: 0,
    end: Math.min(pdf.numPages, 8),
  }))
  // The row being dragged must stay mounted while the list auto-scrolls away from it.
  const visibleThumbnailRows = useMemo(() => {
    const rows = Array.from({ length: Math.max(0, thumbnailWindow.end - thumbnailWindow.start) }, (_, offset) => thumbnailWindow.start + offset)
    if (draggedPage !== null && draggedPage < pdf.numPages && !rows.includes(draggedPage)) rows.push(draggedPage)
    return rows
  }, [thumbnailWindow, draggedPage, pdf.numPages])
  const searchOccurrences = useMemo(() => flattenSearchResults(results), [results])
  const activeResultIndex = searchOccurrences[activeMatchIndex]?.resultIndex ?? -1

  const stopAutoScroll = useCallback(() => {
    const state = autoScrollRef.current
    state.velocity = 0
    if (state.frame) { cancelAnimationFrame(state.frame); state.frame = 0 }
  }, [])

  const autoScrollWhileDragging = useCallback((event: ReactDragEvent<HTMLElement>) => {
    const element = currentRef.current
    if (!element) return
    const state = autoScrollRef.current
    const bounds = element.getBoundingClientRect()
    const zone = Math.min(90, bounds.height / 3)
    const fromTop = event.clientY - bounds.top
    const fromBottom = bounds.bottom - event.clientY
    // Speed ramps up the closer the pointer gets to the edge.
    state.velocity = fromTop < zone ? -Math.ceil((1 - Math.max(0, fromTop) / zone) * 28)
      : fromBottom < zone ? Math.ceil((1 - Math.max(0, fromBottom) / zone) * 28) : 0
    state.lastOver = performance.now()
    if (state.velocity && !state.frame) {
      const tick = () => {
        state.frame = 0
        if (!state.velocity || performance.now() - state.lastOver > 250) { state.velocity = 0; return }
        element.scrollTop += state.velocity
        state.frame = requestAnimationFrame(tick)
      }
      state.frame = requestAnimationFrame(tick)
    }
  }, [])

  useEffect(() => stopAutoScroll, [stopAutoScroll])

  const dropOnPages = useCallback((event: ReactDragEvent<HTMLElement>, insertIndex: number) => {
    stopAutoScroll()
    const internal = Array.from(event.dataTransfer.types).includes(INTERNAL_PAGE_DRAG_TYPE)
    const dragged = draggedIndicesRef.current
    draggedIndicesRef.current = []
    setDraggedPage(null)
    if (internal && dragged.length) onReorder(dragged, insertIndex)
    else if (event.dataTransfer.files.length) onImportPagesAt?.(Array.from(event.dataTransfer.files).filter(isImportableTransferFile), insertIndex)
  }, [onImportPagesAt, onReorder, stopAutoScroll])

  const updateThumbnailWindow = useCallback((element: HTMLDivElement) => {
    const visibleStart = Math.floor(element.scrollTop / THUMBNAIL_ROW_HEIGHT)
    const visibleEnd = Math.ceil((element.scrollTop + element.clientHeight) / THUMBNAIL_ROW_HEIGHT)
    const next = {
      start: Math.min(pdf.numPages, Math.max(0, visibleStart - THUMBNAIL_OVERSCAN)),
      end: Math.min(pdf.numPages, visibleEnd + THUMBNAIL_OVERSCAN),
    }
    setThumbnailWindow((current) => current.start === next.start && current.end === next.end ? current : next)
  }, [pdf.numPages])

  useEffect(() => {
    if (tab !== 'pages' || !currentRef.current) return
    const element = currentRef.current
    const targetTop = pageIndex * THUMBNAIL_ROW_HEIGHT
    const targetBottom = targetTop + THUMBNAIL_ROW_HEIGHT
    if (targetTop < element.scrollTop) element.scrollTop = targetTop
    else if (targetBottom > element.scrollTop + element.clientHeight) element.scrollTop = targetBottom - element.clientHeight
    updateThumbnailWindow(element)
  }, [pageIndex, tab, updateThumbnailWindow])

  useEffect(() => {
    if (tab !== 'pages' || !currentRef.current) return
    const element = currentRef.current
    updateThumbnailWindow(element)
    if (typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => updateThumbnailWindow(element))
    observer.observe(element)
    return () => observer.disconnect()
  }, [tab, updateThumbnailWindow])

  useEffect(() => {
    searchRunRef.current += 1
    setQuery('')
    setResults([])
    setLastSearchedQuery('')
    setActiveMatchIndex(-1)
    setTextlessPages(0)
    onActiveSearchMatch(null)
    setSearchError('')
    setSearching(false)
    return () => { searchRunRef.current += 1 }
  }, [pdf, onActiveSearchMatch])

  useEffect(() => {
    if (searchRequestId <= 0) return
    if (tab !== 'search') {
      lastNonSearchTab.current = tab
      setTab('search')
    }
    const focusSearch = () => {
      searchInputRef.current?.focus()
      searchInputRef.current?.select()
    }
    const frame = window.requestAnimationFrame(focusSearch)
    const timer = window.setTimeout(focusSearch, 0)
    return () => {
      window.cancelAnimationFrame(frame)
      window.clearTimeout(timer)
    }
  }, [searchRequestId])

  useEffect(() => {
    if (tab === 'search') searchInputRef.current?.focus()
  }, [tab])

  useEffect(() => {
    if (tab !== 'search' || activeResultIndex < 0) return
    document.querySelector<HTMLElement>(`[data-search-result="${activeResultIndex}"]`)?.scrollIntoView({ block: 'nearest' })
  }, [tab, activeResultIndex])

  function activateOccurrence(
    nextIndex: number,
    occurrences = searchOccurrences,
    needle = lastSearchedQuery,
  ) {
    const occurrence = occurrences[nextIndex]
    if (!occurrence || !needle) return
    setActiveMatchIndex(nextIndex)
    onActiveSearchMatch({
      pageIndex: occurrence.pageIndex,
      occurrenceIndex: occurrence.occurrenceIndex,
      query: needle,
    })
    onPage(occurrence.pageIndex)
  }

  function navigateResults(direction: 1 | -1) {
    if (!searchOccurrences.length) return
    let nextIndex: number
    if (activeMatchIndex >= 0) {
      nextIndex = (activeMatchIndex + direction + searchOccurrences.length) % searchOccurrences.length
    } else if (direction === 1) {
      const afterCurrent = searchOccurrences.findIndex((result) => result.pageIndex >= pageIndex)
      nextIndex = afterCurrent >= 0 ? afterCurrent : 0
    } else {
      const beforeCurrent = findLastOccurrenceAtOrBefore(searchOccurrences, pageIndex)
      nextIndex = beforeCurrent >= 0 ? beforeCurrent : searchOccurrences.length - 1
    }
    activateOccurrence(nextIndex)
  }

  async function runSearch(direction: 1 | -1 = 1) {
    const runId = ++searchRunRef.current
    const needle = normalizeSearchValue(query)
    if (!needle) {
      setResults([])
      setLastSearchedQuery('')
      setActiveMatchIndex(-1)
      onActiveSearchMatch(null)
      setSearching(false)
      setSearchError('')
      return
    }
    setSearching(true)
    setSearchError('')
    try {
      const matches = await searchDocument(pdf, needle, { isCancelled: () => runId !== searchRunRef.current })
      if (!matches || runId !== searchRunRef.current) return
      setResults(matches)
      setLastSearchedQuery(needle)
      // Every page has been read by now.
      setTextlessPages(pagesWithoutText.get(pdf)?.size ?? 0)
      if (matches.length) {
        const occurrences = flattenSearchResults(matches)
        let nextIndex: number
        if (direction === 1) {
          const afterCurrent = occurrences.findIndex((result) => result.pageIndex >= pageIndex)
          nextIndex = afterCurrent >= 0 ? afterCurrent : 0
        } else {
          const beforeCurrent = findLastOccurrenceAtOrBefore(occurrences, pageIndex)
          nextIndex = beforeCurrent >= 0 ? beforeCurrent : occurrences.length - 1
        }
        activateOccurrence(nextIndex, occurrences, needle)
      } else {
        setActiveMatchIndex(-1)
        onActiveSearchMatch(null)
      }
    } catch (error) {
      if (runId === searchRunRef.current) setSearchError(errorMessage(error))
    } finally {
      if (runId === searchRunRef.current) setSearching(false)
    }
  }

  function openSearch() {
    if (tab !== 'search') lastNonSearchTab.current = tab
    setTab('search')
  }

  function dismissSearch() {
    searchRunRef.current += 1
    setSearching(false)
    onActiveSearchMatch(null)
    setTab(lastNonSearchTab.current)
    window.requestAnimationFrame(() => {
      document.querySelector<HTMLButtonElement>('.sidebar-tabs button[aria-selected="true"]')?.focus()
    })
  }

  function handleSearchKeyDown(event: ReactKeyboardEvent<HTMLInputElement>) {
    if (event.key === 'Escape') {
      // Fullscreen Escape remains owned by App so immersive mode always exits.
      if (document.fullscreenElement) return
      event.preventDefault()
      event.stopPropagation()
      dismissSearch()
      return
    }
    if (event.key !== 'Enter') return
    event.preventDefault()
    const direction: 1 | -1 = event.shiftKey ? -1 : 1
    if (!searching && results.length && normalizeSearchValue(query) === lastSearchedQuery) navigateResults(direction)
    else void runSearch(direction)
  }

  function startRename(bookmark: Bookmark) {
    if (!bookmarksEditable) return
    renamingBookmarkRef.current = bookmark.id
    setRenameDraft(bookmark.label)
    setRenamingBookmarkId(bookmark.id)
  }

  // Enter or leaving the field keeps the new name; Escape keeps the old one.
  // Runs once per rename even though removing the field can also blur it.
  function finishRename(keep: boolean) {
    const id = renamingBookmarkRef.current
    if (!id) return
    renamingBookmarkRef.current = null
    setRenamingBookmarkId(null)
    if (keep) onRenameBookmark(id, renameDraft)
    // Return focus to the row, unless something else has taken it since.
    window.setTimeout(() => {
      const active = document.activeElement
      if (active && active !== document.body) return
      document.querySelector<HTMLElement>(`[data-bookmark-id="${CSS.escape(id)}"] .bookmark-main`)?.focus({ preventScroll: true })
    }, 0)
  }

  return (
    <aside className="sidebar">
      <div className="sidebar-header">
        <div className="sidebar-tabs" role="tablist">
          <button type="button" role="tab" aria-selected={tab === 'pages'} onClick={() => { onActiveSearchMatch(null); setTab('pages') }}>Pages</button>
          <button type="button" role="tab" aria-selected={tab === 'bookmarks'} onClick={() => { onActiveSearchMatch(null); setTab('bookmarks') }}>Bookmarks</button>
          <button type="button" className="tab-icon" role="tab" aria-label="Search" aria-selected={tab === 'search'} onClick={openSearch}><Search size={15} /></button>
        </div>
        <IconButton icon={PanelLeftClose} label="Close sidebar" compact onClick={onClose} />
      </div>

      {tab === 'pages' && (
        <>
          <div className="sidebar-summary">
            <span>{pdf.numPages} {pdf.numPages === 1 ? 'page' : 'pages'}</span>
            {selectedPages.size > 1 && <strong>{selectedPages.size} selected</strong>}
          </div>
          <div
            className="thumbnails"
            ref={currentRef}
            role="listbox"
            aria-label="PDF pages"
            aria-multiselectable="true"
            onScroll={(event) => updateThumbnailWindow(event.currentTarget)}
            onDragOver={(event) => {
              event.preventDefault()
              event.dataTransfer.dropEffect = Array.from(event.dataTransfer.types).includes('Files') ? 'copy' : 'move'
              autoScrollWhileDragging(event)
            }}
            onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) stopAutoScroll() }}
            onDrop={(event) => { event.preventDefault(); dropOnPages(event, pdf.numPages) }}
          >
            <div className="thumbnail-virtual-spacer" style={{ height: pdf.numPages * THUMBNAIL_ROW_HEIGHT }}>
              {visibleThumbnailRows.map((index) => (
                <div key={index} className="thumbnail-virtual-row" style={{ transform: `translateY(${index * THUMBNAIL_ROW_HEIGHT}px)` }}>
                  <PageThumbnail
                    pdf={pdf}
                    index={index}
                    selected={selectedPages.has(index)}
                    current={pageIndex === index}
                    bookmarked={bookmarks.some((bookmark) => bookmark.pageIndex === index)}
                    rotation={pageRotations[index] || 0}
                    canDelete={selectedPages.size < pdf.numPages}
                    exportPageCount={selectedPages.has(index) ? selectedPages.size : 1}
                    onClick={(event) => onSelectPage(index, event)}
                    onRotateLeft={onRotateLeft}
                    onRotateRight={onRotateRight}
                    onDelete={onDeletePages}
                    onDragStart={(event) => {
                      const indices = pageIndicesForTransfer(selectedPages, index)
                      draggedIndicesRef.current = indices
                      event.dataTransfer.effectAllowed = 'move'
                      event.dataTransfer.setData(INTERNAL_PAGE_DRAG_TYPE, indices.join(','))
                      // Re-rendering inside dragstart can make Chromium cancel the drag.
                      window.setTimeout(() => setDraggedPage(index), 0)
                    }}
                    onExportDragStart={(event) => {
                      event.dataTransfer.effectAllowed = 'copy'
                      onPageDragStart?.(index, event)
                    }}
                    onDragEnd={() => { stopAutoScroll(); draggedIndicesRef.current = []; setDraggedPage(null) }}
                    onDrop={dropOnPages}
                  />
                </div>
              ))}
            </div>
          </div>
          <div className="sidebar-footer">
            <div className="page-action-menu">
              <IconButton icon={MoreHorizontal} label="More page actions" active={pageMenuOpen} onClick={() => setPageMenuOpen((open) => !open)} />
              {pageMenuOpen && (
                <div className="page-actions-popover" role="menu">
                  <button type="button" role="menuitem" aria-label="Export selected pages" onClick={() => { setPageMenuOpen(false); onExportPages() }}><Download size={15} /><span><strong>Export selected</strong><small>Save the selection as one PDF</small></span></button>
                  <button type="button" role="menuitem" aria-label="Add pages from files" onClick={() => { setPageMenuOpen(false); onAddPages() }}><FilePlus2 size={15} /><span><strong>Add pages from files</strong><small>Insert after the current page</small></span></button>
                  <button type="button" role="menuitem" onClick={() => { setPageMenuOpen(false); onDuplicatePages() }}><CopyPlus size={15} /><span><strong>Duplicate selected</strong><small>Create a copy after each page</small></span></button>
                  <button type="button" role="menuitem" onClick={() => { setPageMenuOpen(false); onInsertBlank() }}><File size={15} /><span><strong>Insert blank page</strong><small>Add after the current page</small></span></button>
                </div>
              )}
            </div>
          </div>
        </>
      )}

      {tab === 'bookmarks' && (
        <div className="sidebar-pane">
          {bookmarks.length ? (
            <div className="bookmark-list">
              {bookmarks
                .map((bookmark) => {
                  const renaming = renamingBookmarkId === bookmark.id
                  const titleStyle = { fontWeight: bookmark.bold ? 700 : undefined, fontStyle: bookmark.italic ? 'italic' as const : undefined }
                  return (
                    <div
                      key={bookmark.id}
                      data-bookmark-id={bookmark.id}
                      className={cx('bookmark-row', bookmark.pageIndex === pageIndex && 'is-current')}
                      style={{ paddingLeft: Math.min(6, bookmark.depth || 0) * 13 }}
                      onKeyDown={(event) => {
                        if (event.key !== 'F2' || renaming || !bookmarksEditable) return
                        event.preventDefault()
                        startRename(bookmark)
                      }}
                    >
                      {renaming ? (
                        <form
                          className="bookmark-main"
                          style={{ gridColumn: '1 / -1', cursor: 'default' }}
                          onSubmit={(event) => { event.preventDefault(); finishRename(true) }}
                        >
                          <BookmarkIcon size={15} fill="currentColor" />
                          <input
                            autoFocus
                            aria-label="Bookmark name"
                            value={renameDraft}
                            maxLength={500}
                            onFocus={(event) => event.currentTarget.select()}
                            onChange={(event) => setRenameDraft(event.target.value)}
                            onBlur={() => finishRename(true)}
                            onKeyDown={(event) => {
                              // Ctrl+S keeps the new name and then saves (the app's
                              // handler sees the renamed list, rendered in between).
                              if ((event.ctrlKey || event.metaKey) && event.key.toLocaleLowerCase() === 's') {
                                finishRename(true)
                                return
                              }
                              // Typing here must not reach the document shortcuts (Delete, Ctrl+Z…).
                              event.stopPropagation()
                              if (event.key === 'Escape') { event.preventDefault(); finishRename(false) }
                              if (event.key === 'Enter') { event.preventDefault(); finishRename(true) }
                            }}
                            style={{ minWidth: 0, flex: 1, padding: '3px 6px', border: '1px solid var(--border-strong)', borderRadius: 5, font: 'inherit', fontSize: 11.5, ...titleStyle }}
                          />
                        </form>
                      ) : (
                        <button
                          type="button"
                          className="bookmark-main"
                          title={bookmark.url || undefined}
                          // The second click of a double-click renames; it must
                          // not open a web link a second time.
                          onClick={(event) => { if (event.detail <= 1) onOpenBookmark(bookmark) }}
                          onDoubleClick={() => startRename(bookmark)}
                        >
                          {typeof bookmark.pageIndex !== 'number' && bookmark.url
                            ? <ExternalLink size={15} />
                            : <BookmarkIcon size={15} fill={typeof bookmark.pageIndex === 'number' ? 'currentColor' : 'none'} />}
                          <span>
                            <strong style={titleStyle}>{bookmark.label}</strong>
                            <small>{bookmarkTargetLabel(bookmark)}</small>
                          </span>
                        </button>
                      )}
                      {!renaming && bookmarksEditable && (
                        <>
                          <button
                            type="button"
                            className="bookmark-more"
                            title="Rename bookmark (F2)"
                            aria-label={`Rename bookmark ${bookmark.label}`}
                            onClick={() => startRename(bookmark)}
                          ><Pencil size={13} /></button>
                          <button
                            type="button"
                            className="bookmark-delete"
                            title="Remove bookmark"
                            aria-label={`Remove bookmark ${bookmark.label}`}
                            onClick={() => onDeleteBookmark(bookmark.id)}
                          ><X size={14} /></button>
                        </>
                      )}
                    </div>
                  )
                })}
            </div>
          ) : (
            <EmptyState icon={BookmarkIcon} title="No bookmarks yet" detail="Use the bookmark button in the toolbar to save a page." />
          )}
        </div>
      )}

      {tab === 'search' && (
        <div className="sidebar-pane search-pane">
          <form className="search-box" onSubmit={(event) => { event.preventDefault(); void runSearch() }}>
            <Search size={15} />
            <input
              ref={searchInputRef}
              autoFocus
              value={query}
              onChange={(event) => {
                searchRunRef.current += 1
                setQuery(event.target.value)
                setResults([])
                setLastSearchedQuery('')
                setActiveMatchIndex(-1)
                onActiveSearchMatch(null)
                setSearchError('')
                setSearching(false)
              }}
              onKeyDown={handleSearchKeyDown}
              placeholder="Find in document"
              aria-label="Find in document"
            />
            {query && <button type="button" aria-label="Clear search" onClick={() => { searchRunRef.current += 1; setQuery(''); setResults([]); setLastSearchedQuery(''); setActiveMatchIndex(-1); onActiveSearchMatch(null); setSearchError(''); setSearching(false); searchInputRef.current?.focus() }}><X size={13} /></button>}
          </form>
          <div className="search-status">
            {searching
              ? 'Searching…'
              : results.length
                ? `${searchOccurrences.length} matches · ${activeMatchIndex + 1} of ${searchOccurrences.length}`
                : query ? 'Press Enter to search' : 'Search every page'}
          </div>
          {searchError && <div className="inline-error">{searchError}</div>}
          {!searching && lastSearchedQuery && textlessPages > 0 && onRecognizeText && (
            <div className="search-ocr-hint" role="note">
              <span>{textlessPages === 1 ? '1 page has no searchable text — it may be scanned.' : `${textlessPages} pages have no searchable text — they may be scanned.`}</span>
              <button type="button" onClick={onRecognizeText}><ScanText size={13} aria-hidden="true" />Recognize text</button>
            </div>
          )}
          <div className="search-results">
            {results.map((result, index) => (
              <button
                key={result.pageIndex}
                type="button"
                className={cx(index === activeResultIndex && 'is-current')}
                data-search-result={index}
                aria-current={index === activeResultIndex ? 'page' : undefined}
                onClick={() => {
                  const firstOccurrence = searchOccurrences.findIndex((occurrence) => occurrence.resultIndex === index)
                  if (firstOccurrence >= 0) activateOccurrence(firstOccurrence)
                }}
              >
                <span>Page {result.pageIndex + 1}</span>
                <p>{result.excerpt}</p>
                <small>{result.count} {result.count === 1 ? 'match' : 'matches'}</small>
              </button>
            ))}
            {!searching && query && !results.length && !searchError && <EmptyState icon={Search} title="No results yet" detail="Press Enter to search the document." />}
          </div>
        </div>
      )}
    </aside>
  )
}
