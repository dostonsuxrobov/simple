import { useEffect, useRef, useState } from 'react'
import {
  Bookmark,
  ChevronLeft,
  ChevronRight,
  Crop,
  Download,
  Eye,
  FolderOpen,
  Hand,
  Highlighter,
  MoreHorizontal,
  MousePointer2,
  PanelLeftClose,
  PanelLeftOpen,
  Pencil,
  PenLine,
  Printer,
  Redo2,
  Save,
  ScanText,
  Signature,
  Square,
  Strikethrough,
  TextCursorInput,
  Type,
  Underline,
  Undo2,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'
import type { ToolMode } from '../types'
import { clamp } from '../lib/utils'
import { Button, IconButton, Separator } from './ui'

interface ToolbarProps {
  sidebarOpen: boolean
  pageIndex: number
  pageCount: number
  zoom: number
  zoomMode: 'fit' | 'width' | 'custom'
  tool: ToolMode
  bookmarked: boolean
  canUndo: boolean
  canRedo: boolean
  /** A PDF is loaded: document actions are available. */
  documentReady: boolean
  onToggleSidebar: () => void
  onOpen: () => void
  onSave: () => void
  onSaveAs: () => void
  onExport: () => void
  onUndo: () => void
  onRedo: () => void
  onTool: (tool: ToolMode) => void
  onPage: (index: number) => void
  onZoom: (zoom: number) => void
  onFit: (mode: 'fit' | 'width' | 'actual' | 'custom') => void
  onPrint: () => void
  /** Opens Recognize text (OCR). */
  onRecognizeText: () => void
  onBookmark: () => void
  onImmersive: () => void
}

/** Tools for reading; every other tool changes the document. */
const READING_TOOLS: ToolMode[] = ['hand', 'select']

/**
 * Reading shows only what reading needs: open, save, page, zoom, print. The
 * editing and markup tools (and undo/redo) appear when Tools is on, or as soon
 * as one of them is chosen from the keyboard; less common actions live under
 * More. The tool buttons stay in the page while hidden, so shortcuts and
 * automation that press them keep working.
 */
export function Toolbar(props: ToolbarProps) {
  const {
    sidebarOpen, pageIndex, pageCount, zoom, zoomMode, tool, bookmarked, canUndo, canRedo, documentReady,
    onToggleSidebar, onOpen, onSave, onSaveAs, onExport, onUndo, onRedo, onTool,
    onPage, onZoom, onFit, onPrint, onRecognizeText, onBookmark, onImmersive,
  } = props
  const [pageValue, setPageValue] = useState(String(pageIndex + 1))
  const [zoomValue, setZoomValue] = useState(String(Math.round(zoom * 100)))
  const [toolsOpen, setToolsOpen] = useState(false)
  const [moreOpen, setMoreOpen] = useState(false)
  const moreRef = useRef<HTMLDivElement>(null)
  // Enter and Escape decide what happens to the page field themselves; the blur they cause
  // must not commit a second time (that applied a pending text edit twice).
  const skipBlurCommit = useRef(false)
  const showTools = toolsOpen || !READING_TOOLS.includes(tool)

  useEffect(() => setPageValue(String(pageIndex + 1)), [pageIndex])
  useEffect(() => setZoomValue(String(Math.round(zoom * 100))), [zoom])

  useEffect(() => {
    if (!moreOpen) return
    const close = (event: MouseEvent) => { if (!moreRef.current?.contains(event.target as Node)) setMoreOpen(false) }
    // Closing the menu is all this Escape does (it must not also leave the active tool).
    const escape = (event: KeyboardEvent) => { if (event.key === 'Escape') { event.preventDefault(); setMoreOpen(false) } }
    // pointerdown, in capture: the Hand tool cancels the mousedown that would follow.
    window.addEventListener('pointerdown', close, true)
    window.addEventListener('keydown', escape, true)
    return () => {
      window.removeEventListener('pointerdown', close, true)
      window.removeEventListener('keydown', escape, true)
    }
  }, [moreOpen])

  function commitZoom() {
    const percent = Number(zoomValue.replace('%', '').trim())
    if (!Number.isFinite(percent) || percent <= 0) {
      setZoomValue(String(Math.round(zoom * 100)))
      return
    }
    const next = clamp(percent / 100, 0.35, 4)
    // Merely tabbing through the field should keep automatic fitting active.
    if (Math.round(next * 100) !== Math.round(zoom * 100)) onZoom(next)
    setZoomValue(String(Math.round(next * 100)))
  }

  function commitPage() {
    const typed = Number.parseInt(pageValue, 10)
    // Empty or 0 is not a page: keep the current one instead of jumping to page 1.
    if (!Number.isFinite(typed) || typed < 1) { setPageValue(String(pageIndex + 1)); return }
    const next = clamp(typed - 1, 0, Math.max(0, pageCount - 1))
    onPage(next)
    setPageValue(String(next + 1))
  }

  function toggleTools() {
    if (showTools) {
      // Back to reading: the tools go away and so does the active one.
      setToolsOpen(false)
      if (!READING_TOOLS.includes(tool)) onTool('select')
    } else {
      setToolsOpen(true)
    }
  }

  const more = (run: () => void) => () => { setMoreOpen(false); run() }

  return (
    <div className={`toolbar${showTools ? ' has-tools' : ' is-reading'}`} role="toolbar" aria-label="PDF tools">
      <div className="toolbar-group toolbar-leading">
        <IconButton icon={sidebarOpen ? PanelLeftClose : PanelLeftOpen} label={sidebarOpen ? 'Close sidebar (F4)' : 'Open sidebar (F4)'} active={sidebarOpen} onClick={onToggleSidebar} />
        <IconButton icon={FolderOpen} label="Open (Ctrl+O)" onClick={onOpen} />
        <Button icon={Save} variant="primary" className="save-button" title="Save (Ctrl+S)" onClick={onSave}>Save</Button>
        {/* Undo stays in reach while there is something to undo, also after the tools are hidden. */}
        {(showTools || canUndo || canRedo) && (
          <>
            <Separator />
            <IconButton icon={Undo2} label="Undo (Ctrl+Z)" disabled={!canUndo} onClick={onUndo} />
            <IconButton icon={Redo2} label="Redo (Ctrl+Y)" disabled={!canRedo} onClick={onRedo} />
          </>
        )}
      </div>

      <div className="toolbar-group tool-modes" aria-label="Editing modes" hidden={!showTools}>
        <IconButton icon={Hand} label="Hand tool (H)" active={tool === 'hand'} onClick={() => onTool('hand')} />
        <IconButton icon={MousePointer2} label="Select text (V)" active={tool === 'select'} onClick={() => onTool('select')} />
        <Button icon={TextCursorInput} variant={tool === 'edit' ? 'primary' : 'ghost'} className="edit-mode-button" aria-label="Edit text (E)" title="Edit text and images (E)" aria-pressed={tool === 'edit'} onClick={() => onTool('edit')}>Edit PDF</Button>
        <IconButton icon={Type} label="Add text (T)" active={tool === 'addText'} onClick={() => onTool('addText')} />
        <IconButton icon={Highlighter} label="Highlight text" active={tool === 'highlight'} onClick={() => onTool('highlight')} />
        <IconButton icon={Underline} label="Underline text" active={tool === 'underline'} onClick={() => onTool('underline')} />
        <IconButton icon={Strikethrough} label="Strikeout text" active={tool === 'strikeout'} onClick={() => onTool('strikeout')} />
        <IconButton icon={Pencil} label="Draw freehand" active={tool === 'draw'} onClick={() => onTool('draw')} />
        <IconButton icon={Square} label="Draw rectangle" active={tool === 'rectangle'} onClick={() => onTool('rectangle')} />
        <IconButton icon={Crop} label="Crop page (C)" active={tool === 'crop'} onClick={() => onTool('crop')} />
        <IconButton icon={Signature} label="Sign document" active={tool === 'sign'} onClick={() => onTool('sign')} />
      </div>

      <div className="toolbar-group page-controls">
        <IconButton icon={ChevronLeft} label="Previous page (Page Up)" disabled={pageIndex <= 0} onClick={() => onPage(pageIndex - 1)} />
        <div className="page-field" title="Jump to page">
          <input
            aria-label="Current page"
            inputMode="numeric"
            value={pageValue}
            onChange={(event) => setPageValue(event.target.value.replace(/\D/g, ''))}
            onBlur={() => { if (skipBlurCommit.current) { skipBlurCommit.current = false; return } commitPage() }}
            onKeyDown={(event) => {
              // Enter goes to the page and hands the keyboard back (Page Up / Page Down work again).
              if (event.key === 'Enter') { commitPage(); skipBlurCommit.current = true; event.currentTarget.blur() }
              // Escape only gives the field back its page: no jump, and the window's own Escape does not run.
              if (event.key === 'Escape') { event.preventDefault(); skipBlurCommit.current = true; setPageValue(String(pageIndex + 1)); event.currentTarget.blur() }
            }}
          />
          <span>/ {pageCount}</span>
        </div>
        <IconButton icon={ChevronRight} label="Next page (Page Down)" disabled={pageIndex >= pageCount - 1} onClick={() => onPage(pageIndex + 1)} />
        <Separator />
        <IconButton icon={ZoomOut} label="Zoom out" disabled={zoom <= 0.35} onClick={() => onZoom(clamp(zoom - 0.15, 0.35, 4))} />
        <label className="zoom-value" title="Enter a zoom percentage (35–400%)">
          <input aria-label="Zoom percentage" inputMode="decimal" value={zoomValue} onChange={(event) => setZoomValue(event.target.value)} onFocus={(event) => event.currentTarget.select()} onBlur={commitZoom} onKeyDown={(event) => {
            if (event.key === 'Enter') { event.preventDefault(); commitZoom(); event.currentTarget.blur() }
            if (event.key === 'Escape') { event.preventDefault(); setZoomValue(String(Math.round(zoom * 100))) }
          }} /><span>%</span>
        </label>
        <IconButton icon={ZoomIn} label="Zoom in" disabled={zoom >= 4} onClick={() => onZoom(clamp(zoom + 0.15, 0.35, 4))} />
        <select className="page-fit-select" aria-label="Page fit" value={zoomMode === 'custom' && zoom === 1 ? 'actual' : zoomMode} onChange={(event) => onFit(event.target.value as 'fit' | 'width' | 'actual' | 'custom')}>
          <option value="fit">Fit page</option>
          <option value="width">Fit width</option>
          <option value="actual">Actual size</option>
          {zoomMode === 'custom' && zoom !== 1 && <option value="custom">Custom zoom</option>}
        </select>
      </div>

      <div className="toolbar-group toolbar-trailing">
        <Button icon={PenLine} variant={showTools ? 'primary' : 'ghost'} className="tools-toggle" aria-label={showTools ? 'Hide tools' : 'Show tools'} title="Edit, mark up and sign" aria-pressed={showTools} onClick={toggleTools}>Tools</Button>
        <IconButton icon={Printer} label="Print (Ctrl+P)" disabled={!documentReady} onClick={onPrint} />
        <div className="toolbar-more" ref={moreRef}>
          <IconButton icon={MoreHorizontal} label="More" active={moreOpen} aria-haspopup="menu" aria-expanded={moreOpen} onClick={() => setMoreOpen((open) => !open)} />
          <div className="toolbar-more-menu" role="menu" hidden={!moreOpen}>
            <button type="button" role="menuitem" aria-label="Save As" disabled={!documentReady} onClick={more(onSaveAs)}><Save size={15} /><span>Save As…</span><kbd>Ctrl+Shift+S</kbd></button>
            <button type="button" role="menuitem" className="export-as-button" aria-label="Export As" title="Export As (Ctrl+Shift+E)" disabled={!documentReady} onClick={more(onExport)}><Download size={15} /><span>Export As…</span><kbd>Ctrl+Shift+E</kbd></button>
            <button type="button" role="menuitem" aria-label="Recognize text (OCR)" disabled={!documentReady} onClick={more(onRecognizeText)}><ScanText size={15} /><span>Recognize text (OCR)</span></button>
            <button type="button" role="menuitem" aria-label={bookmarked ? 'Remove bookmark' : 'Bookmark this page'} disabled={!documentReady} onClick={more(onBookmark)}><Bookmark size={15} /><span>{bookmarked ? 'Remove bookmark' : 'Bookmark this page'}</span></button>
            <button type="button" role="menuitem" aria-label="Immersive reading (Esc to exit)" onClick={more(onImmersive)}><Eye size={15} /><span>Immersive reading</span><kbd>Esc exits</kbd></button>
          </div>
        </div>
      </div>
    </div>
  )
}
