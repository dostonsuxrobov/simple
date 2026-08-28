import { useEffect, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import { Eye, Printer, X } from 'lucide-react'
import type { PrintDuplexMode, PrinterSummary } from '../types'
import { Button, IconButton } from './ui'

type PageRangeMode = 'all' | 'current' | 'selected' | 'custom'
type PrintOrientation = 'auto' | 'portrait' | 'landscape'

export interface PrintDialogSubmission {
  deviceName: string
  printerLabel: string
  copies: number
  /** Zero-based, sorted, deduplicated; null means every page. */
  pageIndices: number[] | null
  orientation: PrintOrientation
  color: boolean
  duplexMode?: PrintDuplexMode
  collate: boolean
}

interface PrintDialogProps {
  documentName: string
  pageCount: number
  currentPage: number
  selectedPages: number[]
  onPrint: (submission: PrintDialogSubmission) => void
  onPreview: () => void
  onClose: () => void
}

interface StoredPrinterSettings {
  copies?: number
  orientation?: PrintOrientation
  color?: boolean
  duplexMode?: PrintDuplexMode
  collate?: boolean
}

const PRINT_SETTINGS_KEY = 'folio:print-settings:v1'

function readStoredSettings(): Record<string, StoredPrinterSettings> {
  try {
    const value = JSON.parse(localStorage.getItem(PRINT_SETTINGS_KEY) || '{}')
    return value && typeof value === 'object' && !Array.isArray(value) ? value : {}
  } catch {
    return {}
  }
}

function storeSettings(deviceName: string, settings: StoredPrinterSettings) {
  try {
    localStorage.setItem(PRINT_SETTINGS_KEY, JSON.stringify({ ...readStoredSettings(), [deviceName]: settings }))
  } catch {
    // Persisting print preferences is best-effort only.
  }
}

export function parsePageRange(input: string, pageCount: number): number[] | null {
  const trimmed = input.trim()
  if (!trimmed) return null
  const indices = new Set<number>()
  for (const part of trimmed.split(',')) {
    const token = part.trim()
    if (!token) return null
    const match = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(token)
    if (!match) return null
    const start = Number(match[1])
    const end = match[2] ? Number(match[2]) : start
    if (start < 1 || end < start || end > pageCount) return null
    for (let page = start; page <= end; page += 1) indices.add(page - 1)
  }
  return indices.size ? [...indices].sort((a, b) => a - b) : null
}

export function PrintDialog({ documentName, pageCount, currentPage, selectedPages, onPrint, onPreview, onClose }: PrintDialogProps) {
  const [printers, setPrinters] = useState<PrinterSummary[] | null>(null)
  const [deviceName, setDeviceName] = useState('')
  const [copies, setCopies] = useState('1')
  const [rangeMode, setRangeMode] = useState<PageRangeMode>('all')
  const [customRange, setCustomRange] = useState('')
  const [orientation, setOrientation] = useState<PrintOrientation>('auto')
  const [color, setColor] = useState(true)
  const [duplexMode, setDuplexMode] = useState<PrintDuplexMode>('simplex')
  const [collate, setCollate] = useState(true)
  const printerSelectRef = useRef<HTMLSelectElement>(null)
  const hasSelection = selectedPages.length > 1

  useEffect(() => {
    let cancelled = false
    window.simple.listPrinters().then((list) => {
      if (cancelled) return
      setPrinters(list)
      const preferred = list.find((printer) => printer.isDefault) || list[0]
      if (preferred) setDeviceName(preferred.name)
    }).catch(() => {
      if (!cancelled) setPrinters([])
    })
    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (!deviceName) return
    const stored = readStoredSettings()[deviceName]
    if (!stored) return
    if (Number.isInteger(stored.copies) && (stored.copies as number) >= 1) setCopies(String(stored.copies))
    if (stored.orientation === 'auto' || stored.orientation === 'portrait' || stored.orientation === 'landscape') setOrientation(stored.orientation)
    if (typeof stored.color === 'boolean') setColor(stored.color)
    if (stored.duplexMode === 'simplex' || stored.duplexMode === 'longEdge' || stored.duplexMode === 'shortEdge') setDuplexMode(stored.duplexMode)
    if (typeof stored.collate === 'boolean') setCollate(stored.collate)
  }, [deviceName])

  useEffect(() => {
    if (printers?.length) printerSelectRef.current?.focus()
  }, [printers])

  const activePrinter = printers?.find((printer) => printer.name === deviceName) || null
  const customIndices = parsePageRange(customRange, pageCount)
  const customInvalid = rangeMode === 'custom' && customRange.trim() !== '' && !customIndices
  const copiesValue = Math.max(1, Math.trunc(Number(copies) || 0))
  const canPrint = Boolean(activePrinter) && copiesValue >= 1 && (rangeMode !== 'custom' || Boolean(customIndices))

  function resolvedIndices(): number[] | null {
    if (rangeMode === 'current') return [currentPage]
    if (rangeMode === 'selected') return selectedPages
    if (rangeMode === 'custom') return customIndices
    return null
  }

  function submit() {
    if (!activePrinter || !canPrint) return
    const submission: PrintDialogSubmission = {
      deviceName: activePrinter.name,
      printerLabel: activePrinter.displayName,
      copies: copiesValue,
      pageIndices: resolvedIndices(),
      orientation,
      color,
      ...(activePrinter.supportsDuplex ? { duplexMode } : {}),
      collate,
    }
    storeSettings(activePrinter.name, { copies: copiesValue, orientation, color, duplexMode, collate })
    onPrint(submission)
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }
    if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement)) {
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

  return (
    <div className="print-dialog-overlay" onPointerDown={(event) => event.target === event.currentTarget && onClose()}>
      <div className="print-dialog" role="dialog" aria-modal="true" aria-label="Print" onKeyDown={handleKeyDown}>
        <header className="print-dialog-header">
          <Printer size={15} strokeWidth={1.8} aria-hidden="true" />
          <strong>Print</strong>
          <span className="print-dialog-document">{documentName}</span>
          <IconButton icon={X} label="Cancel printing" compact onClick={onClose} />
        </header>

        <div className="print-dialog-body">
          <label className="print-field">
            <span>Printer</span>
            <select
              ref={printerSelectRef}
              value={deviceName}
              disabled={!printers?.length}
              onChange={(event) => setDeviceName(event.target.value)}
            >
              {printers?.map((printer) => (
                <option key={printer.name} value={printer.name}>
                  {printer.displayName}{printer.isDefault ? ' (default)' : ''}
                </option>
              ))}
            </select>
            {printers === null && <small className="print-note">Looking for printers…</small>}
            {printers?.length === 0 && <small className="print-note">No printers found — use the preview to print from Windows.</small>}
          </label>

          <div className="print-field-row">
            <label className="print-field">
              <span>Copies</span>
              <input
                type="number"
                min={1}
                max={999}
                value={copies}
                onChange={(event) => setCopies(event.target.value)}
                onBlur={() => setCopies(String(copiesValue))}
              />
            </label>
            <label className="print-check">
              <input type="checkbox" checked={collate} disabled={copiesValue < 2} onChange={(event) => setCollate(event.target.checked)} />
              Collate copies
            </label>
          </div>

          <fieldset className="print-field">
            <span>Pages</span>
            <label className="print-range-option">
              <input type="radio" name="print-range" checked={rangeMode === 'all'} onChange={() => setRangeMode('all')} />
              All {pageCount === 1 ? 'pages' : `${pageCount} pages`}
            </label>
            <label className="print-range-option">
              <input type="radio" name="print-range" checked={rangeMode === 'current'} onChange={() => setRangeMode('current')} />
              Current page ({currentPage + 1})
            </label>
            {hasSelection && (
              <label className="print-range-option">
                <input type="radio" name="print-range" checked={rangeMode === 'selected'} onChange={() => setRangeMode('selected')} />
                Selected pages ({selectedPages.length})
              </label>
            )}
            <label className="print-range-option">
              <input type="radio" name="print-range" checked={rangeMode === 'custom'} onChange={() => setRangeMode('custom')} />
              Custom
              <input
                type="text"
                aria-label="Custom page range"
                placeholder="e.g. 1-3, 5"
                value={customRange}
                onFocus={() => setRangeMode('custom')}
                onChange={(event) => setCustomRange(event.target.value)}
              />
            </label>
            {customInvalid && <small className="print-range-error">Use page numbers between 1 and {pageCount}, like 1-3, 5.</small>}
          </fieldset>

          <div className="print-field-row">
            <label className="print-field">
              <span>Orientation</span>
              <select value={orientation} onChange={(event) => setOrientation(event.target.value as PrintOrientation)}>
                <option value="auto">Auto</option>
                <option value="portrait">Portrait</option>
                <option value="landscape">Landscape</option>
              </select>
            </label>
            <label className="print-field">
              <span>Color</span>
              <select value={color ? 'color' : 'bw'} onChange={(event) => setColor(event.target.value === 'color')}>
                <option value="color">Color</option>
                <option value="bw">Black &amp; white</option>
              </select>
            </label>
          </div>

          {activePrinter?.supportsDuplex && (
            <label className="print-field">
              <span>Two-sided</span>
              <select value={duplexMode} onChange={(event) => setDuplexMode(event.target.value as PrintDuplexMode)}>
                <option value="simplex">One-sided</option>
                <option value="longEdge">Two-sided, flip on long edge</option>
                <option value="shortEdge">Two-sided, flip on short edge</option>
              </select>
            </label>
          )}
        </div>

        <footer className="print-dialog-footer">
          <Button icon={Eye} variant="ghost" onClick={onPreview}>Open preview</Button>
          <Button variant="secondary" onClick={onClose}>Cancel</Button>
          <Button icon={Printer} variant="primary" disabled={!canPrint} onClick={submit}>Print</Button>
        </footer>
      </div>
    </div>
  )
}
