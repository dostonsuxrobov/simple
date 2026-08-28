import { useEffect, useState } from 'react'
import {
  Bookmark,
  ChevronLeft,
  ChevronRight,
  Crop,
  Eye,
  FolderOpen,
  Hand,
  Highlighter,
  MousePointer2,
  PanelLeftClose,
  PanelLeftOpen,
  Pencil,
  Printer,
  Redo2,
  Save,
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
  tool: ToolMode
  bookmarked: boolean
  canUndo: boolean
  canRedo: boolean
  onToggleSidebar: () => void
  onOpen: () => void
  onSave: () => void
  onUndo: () => void
  onRedo: () => void
  onTool: (tool: ToolMode) => void
  onPage: (index: number) => void
  onZoom: (zoom: number) => void
  onPrint: () => void
  onBookmark: () => void
  onImmersive: () => void
}

export function Toolbar(props: ToolbarProps) {
  const {
    sidebarOpen, pageIndex, pageCount, zoom, tool, bookmarked, canUndo, canRedo,
    onToggleSidebar, onOpen, onSave, onUndo, onRedo, onTool,
    onPage, onZoom, onPrint, onBookmark, onImmersive,
  } = props
  const [pageValue, setPageValue] = useState(String(pageIndex + 1))

  useEffect(() => setPageValue(String(pageIndex + 1)), [pageIndex])

  function commitPage() {
    const next = clamp((Number.parseInt(pageValue, 10) || 1) - 1, 0, Math.max(0, pageCount - 1))
    onPage(next)
    setPageValue(String(next + 1))
  }

  return (
    <div className="toolbar" role="toolbar" aria-label="PDF tools">
      <div className="toolbar-group toolbar-leading">
        <IconButton icon={sidebarOpen ? PanelLeftClose : PanelLeftOpen} label={sidebarOpen ? 'Close sidebar (F4)' : 'Open sidebar (F4)'} active={sidebarOpen} onClick={onToggleSidebar} />
        <IconButton icon={FolderOpen} label="Open (Ctrl+O)" onClick={onOpen} />
        <Button icon={Save} variant="primary" className="save-button" onClick={onSave}>Save</Button>
        <Separator />
        <IconButton icon={Undo2} label="Undo (Ctrl+Z)" disabled={!canUndo} onClick={onUndo} />
        <IconButton icon={Redo2} label="Redo (Ctrl+Y)" disabled={!canRedo} onClick={onRedo} />
      </div>

      <div className="toolbar-group tool-modes" aria-label="Editing modes">
        <IconButton icon={Hand} label="Hand tool (H)" active={tool === 'hand'} onClick={() => onTool('hand')} />
        <IconButton icon={MousePointer2} label="Select text (V)" active={tool === 'select'} onClick={() => onTool('select')} />
        <IconButton icon={TextCursorInput} label="Edit text (E)" active={tool === 'edit'} onClick={() => onTool('edit')} />
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
            onBlur={commitPage}
            onKeyDown={(event) => event.key === 'Enter' && commitPage()}
          />
          <span>/ {pageCount}</span>
        </div>
        <IconButton icon={ChevronRight} label="Next page (Page Down)" disabled={pageIndex >= pageCount - 1} onClick={() => onPage(pageIndex + 1)} />
        <Separator />
        <IconButton icon={ZoomOut} label="Zoom out" disabled={zoom <= 0.35} onClick={() => onZoom(clamp(zoom - 0.15, 0.35, 4))} />
        <span className="zoom-value" aria-label={`Zoom ${Math.round(zoom * 100)} percent`}>{Math.round(zoom * 100)}%</span>
        <IconButton icon={ZoomIn} label="Zoom in" disabled={zoom >= 4} onClick={() => onZoom(clamp(zoom + 0.15, 0.35, 4))} />
      </div>

      <div className="toolbar-group toolbar-trailing">
        <IconButton icon={Eye} label="Immersive reading (Esc to exit)" onClick={onImmersive} />
        <IconButton icon={Bookmark} label={bookmarked ? 'Remove bookmark' : 'Bookmark this page'} active={bookmarked} onClick={onBookmark} />
        <IconButton icon={Printer} label="Print (Ctrl+P)" onClick={onPrint} />
      </div>
    </div>
  )
}
