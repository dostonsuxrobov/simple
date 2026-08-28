import { Crop, Hand, Highlighter, MousePointer2, Pencil, Signature, Square, Strikethrough, TextCursorInput, Type, Underline } from 'lucide-react'
import type { ToolMode } from '../types'

const toolDetails: Record<ToolMode, { icon: typeof Hand; label: string; hint: string }> = {
  hand: { icon: Hand, label: 'Hand', hint: 'Scroll to move around the page' },
  select: { icon: MousePointer2, label: 'Select', hint: 'Drag across text to select and copy' },
  edit: { icon: TextCursorInput, label: 'Edit content', hint: 'Click text, images, or artwork to edit directly' },
  addText: { icon: Type, label: 'Add text', hint: 'Click anywhere on the page to add text' },
  highlight: { icon: Highlighter, label: 'Highlight', hint: 'Drag across text, or click a text line' },
  underline: { icon: Underline, label: 'Underline', hint: 'Select text to underline it' },
  strikeout: { icon: Strikethrough, label: 'Strikeout', hint: 'Select text to strike it out' },
  draw: { icon: Pencil, label: 'Draw', hint: 'Drag directly on the page to add ink' },
  rectangle: { icon: Square, label: 'Rectangle', hint: 'Drag a rectangle directly on the page' },
  crop: { icon: Crop, label: 'Crop', hint: 'Drag a rectangle, then choose Apply' },
  sign: { icon: Signature, label: 'Sign', hint: 'Click a saved signature, then click the page to place it' },
}

export function StatusBar({ tool, pageIndex, pageCount, selectionCount }: { tool: ToolMode; pageIndex: number; pageCount: number; selectionCount: number }) {
  const details = toolDetails[tool]
  const Icon = details.icon
  return (
    <footer className="statusbar">
      <div className="status-tool"><Icon size={13} /><strong>{details.label}</strong><span>{details.hint}</span></div>
      <div className="status-document">
        {selectionCount > 1 && <span>{selectionCount} pages selected</span>}
        <span>Page {pageIndex + 1} of {pageCount}</span>
      </div>
    </footer>
  )
}
