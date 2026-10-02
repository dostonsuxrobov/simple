import { useEffect, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import { Download, FileText, FileType2, Globe2, Image, Images, X } from 'lucide-react'
import type { PdfExportFormat } from '../types'
import { parsePageRange } from './PrintDialog'
import { Button, IconButton } from './ui'

type ExportRangeMode = 'all' | 'current' | 'selected' | 'custom'

export interface PdfExportSubmission {
  format: PdfExportFormat
  pageIndices: number[]
  imageScale: number
  jpegQuality: number
}

interface ExportDialogProps {
  documentName: string
  pageCount: number
  currentPage: number
  selectedPages: number[]
  onExport: (submission: PdfExportSubmission) => void
  onClose: () => void
}

const FORMATS: Array<{
  format: PdfExportFormat
  label: string
  detail: string
  icon: typeof FileText
}> = [
  { format: 'pdf', label: 'PDF', detail: 'A new PDF containing the chosen pages', icon: FileText },
  { format: 'png', label: 'PNG images', detail: 'Lossless page images for design or archiving', icon: Images },
  { format: 'jpeg', label: 'JPEG images', detail: 'Smaller page images for sharing', icon: Image },
  { format: 'webp', label: 'WebP images', detail: 'Compact modern images with transparency', icon: Images },
  { format: 'docx', label: 'Word document', detail: 'Editable extracted text; complex layout may change', icon: FileType2 },
  { format: 'txt', label: 'Plain text', detail: 'Text only, organized by PDF page', icon: FileText },
  { format: 'md', label: 'Markdown', detail: 'Portable text with page headings', icon: FileText },
  { format: 'html', label: 'Web page', detail: 'A printable HTML document with page sections', icon: Globe2 },
]

export function ExportDialog({ documentName, pageCount, currentPage, selectedPages, onExport, onClose }: ExportDialogProps) {
  const [format, setFormat] = useState<PdfExportFormat>('png')
  const [rangeMode, setRangeMode] = useState<ExportRangeMode>('all')
  const [customRange, setCustomRange] = useState('')
  const [imageResolution, setImageResolution] = useState('2')
  const [jpegQuality, setJpegQuality] = useState('0.9')
  const firstFormatRef = useRef<HTMLButtonElement>(null)
  const customIndices = parsePageRange(customRange, pageCount)
  const customInvalid = rangeMode === 'custom' && customRange.trim() !== '' && !customIndices
  const hasSelection = selectedPages.length > 1

  useEffect(() => firstFormatRef.current?.focus(), [])

  function resolvedIndices() {
    if (rangeMode === 'current') return [currentPage]
    if (rangeMode === 'selected') return selectedPages
    if (rangeMode === 'custom') return customIndices || []
    return Array.from({ length: pageCount }, (_, index) => index)
  }

  const pageIndices = resolvedIndices()
  const canExport = pageIndices.length > 0 && !customInvalid

  function submit() {
    if (!canExport) return
    onExport({
      format,
      pageIndices,
      imageScale: Number(imageResolution),
      jpegQuality: Number(jpegQuality),
    })
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }
    if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement) && !(event.target instanceof HTMLSelectElement)) {
      event.preventDefault()
      event.stopPropagation()
      submit()
      return
    }
    if (event.key === 'Tab') {
      const focusable = [...event.currentTarget.querySelectorAll<HTMLElement>('button, input, select')]
        .filter((element) => !element.hasAttribute('disabled'))
      if (!focusable.length) return
      const first = focusable[0]
      const last = focusable[focusable.length - 1]
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
    }
  }

  const imageFormat = format === 'png' || format === 'jpeg' || format === 'webp'

  return (
    <div className="export-dialog-overlay" onPointerDown={(event) => event.target === event.currentTarget && onClose()}>
      <div className="export-dialog" role="dialog" aria-modal="true" aria-label="Export As" onKeyDown={handleKeyDown}>
        <header className="export-dialog-header">
          <Download size={16} strokeWidth={1.8} aria-hidden="true" />
          <strong>Export As</strong>
          <span className="export-dialog-document">{documentName}</span>
          <IconButton icon={X} label="Cancel export" compact onClick={onClose} />
        </header>

        <div className="export-dialog-body">
          <section>
            <h3>Format</h3>
            <div className="export-format-grid" role="radiogroup" aria-label="Export format">
              {FORMATS.map((option, index) => {
                const Icon = option.icon
                return (
                  <button
                    key={option.format}
                    ref={index === 0 ? firstFormatRef : undefined}
                    type="button"
                    role="radio"
                    aria-checked={format === option.format}
                    className={format === option.format ? 'is-selected' : ''}
                    onClick={() => setFormat(option.format)}
                  >
                    <span className="export-format-icon"><Icon size={18} strokeWidth={1.7} /></span>
                    <span><strong>{option.label}</strong><small>{option.detail}</small></span>
                  </button>
                )
              })}
            </div>
          </section>

          <section className="export-settings-grid">
            <fieldset className="export-field">
              <legend>Pages</legend>
              <label><input type="radio" name="export-range" checked={rangeMode === 'all'} onChange={() => setRangeMode('all')} />All {pageCount === 1 ? 'page' : `${pageCount} pages`}</label>
              <label><input type="radio" name="export-range" checked={rangeMode === 'current'} onChange={() => setRangeMode('current')} />Current page ({currentPage + 1})</label>
              {hasSelection && <label><input type="radio" name="export-range" checked={rangeMode === 'selected'} onChange={() => setRangeMode('selected')} />Selected pages ({selectedPages.length})</label>}
              <label className="export-custom-range">
                <input type="radio" name="export-range" checked={rangeMode === 'custom'} onChange={() => setRangeMode('custom')} />
                Custom
                <input aria-label="Custom export page range" placeholder="1-3, 5" value={customRange} onFocus={() => setRangeMode('custom')} onChange={(event) => setCustomRange(event.target.value)} />
              </label>
              {customInvalid && <small className="export-error">Use page numbers between 1 and {pageCount}.</small>}
            </fieldset>

            <div className="export-field export-image-options" aria-disabled={!imageFormat}>
              <span>Image quality</span>
              <label>
                Resolution
                <select value={imageResolution} disabled={!imageFormat} onChange={(event) => setImageResolution(event.target.value)}>
                  <option value="1.333333">Screen · 96 DPI</option>
                  <option value="2">Sharp · 144 DPI</option>
                  <option value="3">High · 216 DPI</option>
                </select>
              </label>
              <label>
                Lossy compression
                <select value={jpegQuality} disabled={format !== 'jpeg' && format !== 'webp'} onChange={(event) => setJpegQuality(event.target.value)}>
                  <option value="0.8">Smaller file</option>
                  <option value="0.9">Balanced</option>
                  <option value="0.96">Best quality</option>
                </select>
              </label>
              <small>{imageFormat && pageIndices.length > 1 ? 'Each page becomes a numbered image in a new folder.' : imageFormat ? 'The chosen page becomes one image file.' : 'Image controls apply to PNG, JPEG, and WebP exports.'}</small>
            </div>
          </section>

          {(format === 'docx' || format === 'txt' || format === 'md' || format === 'html') && (
            <p className="export-layout-note">Text exports recover readable content. Scans need OCR, which is offered before the export starts. Complex columns, forms, or precise page positioning may not carry over.</p>
          )}
        </div>

        <footer className="export-dialog-footer">
          <span>{pageIndices.length} {pageIndices.length === 1 ? 'page' : 'pages'} selected</span>
          <Button variant="secondary" onClick={onClose}>Cancel</Button>
          <Button icon={Download} variant="primary" disabled={!canExport} onClick={submit}>Export</Button>
        </footer>
      </div>
    </div>
  )
}
