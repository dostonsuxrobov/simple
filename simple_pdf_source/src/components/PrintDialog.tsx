import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { PDFDocumentProxy, PDFPageProxy, RenderTask } from 'pdfjs-dist'
import { ChevronLeft, ChevronRight, Printer, X } from 'lucide-react'
import { calculatePdfPrintPlacement } from '../lib/pdfPrint'
import type {
  PrintDuplexMode,
  PrintMarginMode,
  PrintPaperSize,
  PrintScaleMode,
  PrinterSummary,
} from '../types'
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
  landscape: boolean
  color: boolean
  duplexMode?: PrintDuplexMode
  collate: boolean
  paperSize: PrintPaperSize
  marginMode: PrintMarginMode
  scaleMode: PrintScaleMode
  /** Decimal scale, where 1 is 100%. */
  customScale: number
}

interface PrintDialogProps {
  documentName: string
  pdf: PDFDocumentProxy
  pageCount: number
  currentPage: number
  selectedPages: number[]
  onPrint: (submission: PrintDialogSubmission) => void
  onClose: () => void
}

interface StoredPrinterSettings {
  copies?: number
  orientation?: PrintOrientation
  color?: boolean
  duplexMode?: PrintDuplexMode
  collate?: boolean
  paperSize?: PrintPaperSize
  marginMode?: PrintMarginMode
  scaleMode?: PrintScaleMode
  customScale?: number
}

interface PreviewBitmap {
  dataUrl: string
  sourceWidth: number
  sourceHeight: number
  rotation: number
  pageIndex: number
}

const PRINT_SETTINGS_KEY = 'folio:print-settings:v2'

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
    localStorage.setItem(PRINT_SETTINGS_KEY, JSON.stringify({ ...readStoredSettings(), [deviceName || 'preview']: settings }))
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

function printableIndices(indices: number[] | null, pageCount: number) {
  return indices || Array.from({ length: pageCount }, (_, index) => index)
}

export function PrintDialog({ documentName, pdf, pageCount, currentPage, selectedPages, onPrint, onClose }: PrintDialogProps) {
  const [printers, setPrinters] = useState<PrinterSummary[] | null>(null)
  const [deviceName, setDeviceName] = useState('')
  const [copies, setCopies] = useState('1')
  const [rangeMode, setRangeMode] = useState<PageRangeMode>('all')
  const [customRange, setCustomRange] = useState('')
  const [orientation, setOrientation] = useState<PrintOrientation>('auto')
  const [color, setColor] = useState(true)
  const [duplexMode, setDuplexMode] = useState<PrintDuplexMode>('simplex')
  const [collate, setCollate] = useState(true)
  const [paperSize, setPaperSize] = useState<PrintPaperSize>('Letter')
  const [marginMode, setMarginMode] = useState<PrintMarginMode>('normal')
  const [scaleMode, setScaleMode] = useState<PrintScaleMode>('fit')
  const [customScale, setCustomScale] = useState('100')
  const [previewPosition, setPreviewPosition] = useState(0)
  const [autoLandscape, setAutoLandscape] = useState(false)
  const [autoOrientationReady, setAutoOrientationReady] = useState(false)
  const [previewBitmap, setPreviewBitmap] = useState<PreviewBitmap | null>(null)
  const [previewError, setPreviewError] = useState('')
  const [previewStageSize, setPreviewStageSize] = useState({ width: 0, height: 0 })
  const returnFocusRef = useRef<HTMLElement | null>(document.activeElement instanceof HTMLElement ? document.activeElement : null)
  const printerSelectRef = useRef<HTMLSelectElement>(null)
  const paperSelectRef = useRef<HTMLSelectElement>(null)
  const previewStageRef = useRef<HTMLDivElement>(null)
  const hasSelection = selectedPages.length > 0

  useEffect(() => {
    let cancelled = false
    window.simple.listPrinters().then((list) => {
      if (cancelled) return
      setPrinters(list)
      if (list.length) setDeviceName('')
    }).catch(() => {
      if (!cancelled) setPrinters([])
    })
    return () => { cancelled = true }
  }, [])

  useEffect(() => {
    const stored = readStoredSettings()[deviceName || 'preview']
    if (!stored) return
    if (Number.isInteger(stored.copies) && (stored.copies as number) >= 1) setCopies(String(Math.min(999, stored.copies as number)))
    if (stored.orientation === 'auto' || stored.orientation === 'portrait' || stored.orientation === 'landscape') setOrientation(stored.orientation)
    if (typeof stored.color === 'boolean') setColor(stored.color)
    if (stored.duplexMode === 'simplex' || stored.duplexMode === 'longEdge' || stored.duplexMode === 'shortEdge') setDuplexMode(stored.duplexMode)
    if (typeof stored.collate === 'boolean') setCollate(stored.collate)
    if (stored.paperSize === 'Letter' || stored.paperSize === 'A4' || stored.paperSize === 'Legal') setPaperSize(stored.paperSize)
    if (stored.marginMode === 'none' || stored.marginMode === 'minimum' || stored.marginMode === 'normal') setMarginMode(stored.marginMode)
    if (stored.scaleMode === 'fit' || stored.scaleMode === 'actual' || stored.scaleMode === 'shrink' || stored.scaleMode === 'custom') setScaleMode(stored.scaleMode)
    if (Number.isFinite(stored.customScale)) setCustomScale(String(Math.round(Math.min(4, Math.max(0.25, stored.customScale as number)) * 100)))
  }, [deviceName])

  useEffect(() => {
    if (document.activeElement instanceof HTMLElement && document.activeElement.closest('.print-layout-dialog')) return
    if (printers?.length) printerSelectRef.current?.focus()
    else if (printers) paperSelectRef.current?.focus()
  }, [printers])

  useEffect(() => {
    paperSelectRef.current?.focus()
    return () => returnFocusRef.current?.focus()
  }, [])

  useEffect(() => {
    const stage = previewStageRef.current
    if (!stage) return
    const measure = () => setPreviewStageSize({
      width: Math.max(0, stage.clientWidth - 44),
      height: Math.max(0, stage.clientHeight - 44),
    })
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(stage)
    return () => observer.disconnect()
  }, [])

  const activePrinter = !printers?.length
    ? null
    : deviceName
      ? printers.find((printer) => printer.name === deviceName) || null
      : { name: '', displayName: 'Default Windows printer', supportsDuplex: true, supportsColor: true }
  const customIndices = parsePageRange(customRange, pageCount)
  const customInvalid = rangeMode === 'custom' && !customIndices
  const copiesValue = Math.min(999, Math.max(1, Math.trunc(Number(copies) || 0)))
  const customScaleValue = Math.min(400, Math.max(25, Math.trunc(Number(customScale) || 100)))

  const resolvedPageIndices = useMemo(() => {
    if (rangeMode === 'current') return [currentPage]
    if (rangeMode === 'selected') return selectedPages
    if (rangeMode === 'custom') return customIndices
    return null
  }, [currentPage, customIndices, rangeMode, selectedPages])
  const previewIndices = useMemo(() => printableIndices(resolvedPageIndices, pageCount), [pageCount, resolvedPageIndices])
  const orientationAnchorIndex = previewIndices[0] ?? 0
  const previewPageIndex = previewIndices[Math.min(previewPosition, Math.max(0, previewIndices.length - 1))] ?? 0

  useEffect(() => { setPreviewPosition(0) }, [rangeMode, customRange, selectedPages])

  useEffect(() => {
    let cancelled = false
    setAutoLandscape(false)
    setAutoOrientationReady(false)
    pdf.getPage(orientationAnchorIndex + 1).then((page) => {
      const sourceWidth = Math.abs(page.view[2] - page.view[0])
      const sourceHeight = Math.abs(page.view[3] - page.view[1])
      const quarterTurn = Math.abs(page.rotate || 0) % 180 === 90
      if (!cancelled) {
        setAutoLandscape(quarterTurn ? sourceHeight > sourceWidth : sourceWidth > sourceHeight)
        setAutoOrientationReady(true)
      }
      page.cleanup()
    }).catch(() => { if (!cancelled) setAutoOrientationReady(true) })
    return () => { cancelled = true }
  }, [orientationAnchorIndex, pdf])

  useEffect(() => {
    let cancelled = false
    let renderTask: RenderTask | null = null
    let renderedPage: PDFPageProxy | null = null
    let renderCanvas: HTMLCanvasElement | null = null
    setPreviewBitmap(null)
    setPreviewError('')
    pdf.getPage(previewPageIndex + 1).then(async (page) => {
      renderedPage = page
      if (cancelled) { page.cleanup(); return }
      const sourceWidth = Math.abs(page.view[2] - page.view[0])
      const sourceHeight = Math.abs(page.view[3] - page.view[1])
      const baseViewport = page.getViewport({ scale: 1, rotation: page.rotate })
      const renderScale = Math.min(2, 1400 / Math.max(1, baseViewport.width), 1800 / Math.max(1, baseViewport.height))
      const viewport = page.getViewport({ scale: renderScale, rotation: page.rotate })
      const canvas = document.createElement('canvas')
      renderCanvas = canvas
      canvas.width = Math.max(1, Math.round(viewport.width))
      canvas.height = Math.max(1, Math.round(viewport.height))
      const context = canvas.getContext('2d', { alpha: false })
      if (!context) throw new Error('The print preview could not create a page canvas.')
      context.fillStyle = '#fff'
      context.fillRect(0, 0, canvas.width, canvas.height)
      renderTask = page.render({ canvasContext: context, viewport })
      await renderTask.promise
      if (!cancelled) {
        setPreviewBitmap({ dataUrl: canvas.toDataURL('image/jpeg', 0.9), sourceWidth, sourceHeight, rotation: page.rotate || 0, pageIndex: previewPageIndex })
      }
      page.cleanup()
      canvas.width = 1
      canvas.height = 1
      renderedPage = null
      renderCanvas = null
    }).catch((error) => {
      if (!cancelled && error?.name !== 'RenderingCancelledException') setPreviewError(error instanceof Error ? error.message : String(error))
    })
    return () => {
      cancelled = true
      renderTask?.cancel()
      renderedPage?.cleanup()
      if (renderCanvas) {
        renderCanvas.width = 1
        renderCanvas.height = 1
      }
    }
  }, [pdf, previewPageIndex])

  // Orientation is a job-level printer setting. Auto follows the first page in
  // the chosen range and stays stable while the user browses later preview pages.
  const landscape = orientation === 'landscape' || (orientation === 'auto' && autoLandscape)
  const layout = { paperSize, landscape, marginMode, scaleMode, customScale: customScaleValue / 100 } as const
  const activePreviewBitmap = previewBitmap?.pageIndex === previewPageIndex ? previewBitmap : null
  const placement = calculatePdfPrintPlacement(
    activePreviewBitmap?.sourceWidth || 612,
    activePreviewBitmap?.sourceHeight || 792,
    activePreviewBitmap?.rotation || 0,
    layout,
  )
  const paperRatio = placement.paperWidth / placement.paperHeight
  const previewPaperWidth = previewStageSize.width > 0 && previewStageSize.height > 0
    ? Math.max(1, Math.min(510, previewStageSize.width, previewStageSize.height * paperRatio))
    : undefined
  const canPreview = !customInvalid && previewIndices.length > 0 && (orientation !== 'auto' || autoOrientationReady)
    && Boolean(activePreviewBitmap) && !previewError
  const canPrint = Boolean(activePrinter) && copiesValue >= 1 && canPreview

  function submission(): PrintDialogSubmission {
    return {
      deviceName: activePrinter?.name || '',
      printerLabel: activePrinter?.displayName || 'selected printer',
      copies: copiesValue,
      pageIndices: resolvedPageIndices,
      orientation,
      landscape,
      color,
      ...(activePrinter?.supportsDuplex ? { duplexMode } : {}),
      collate,
      paperSize,
      marginMode,
      scaleMode,
      customScale: customScaleValue / 100,
    }
  }

  function persist() {
    storeSettings(deviceName || 'preview', {
      copies: copiesValue,
      orientation,
      color,
      duplexMode,
      collate,
      paperSize,
      marginMode,
      scaleMode,
      customScale: customScaleValue / 100,
    })
  }

  function submit() {
    if (!canPrint) return
    persist()
    onPrint(submission())
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
      const focusable = [...event.currentTarget.querySelectorAll<HTMLElement>('button, input, select, [tabindex="0"]')]
        .filter((element) => !element.hasAttribute('disabled') && element.getClientRects().length > 0)
      if (!focusable.length) return
      const first = focusable[0]
      const last = focusable[focusable.length - 1]
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
    }
  }

  return (
    <div className="print-dialog-overlay" onPointerDown={(event) => event.target === event.currentTarget && onClose()}>
      <div className="print-dialog print-layout-dialog" role="dialog" aria-modal="true" aria-label="Print" onKeyDown={handleKeyDown}>
        <header className="print-dialog-header">
          <Printer size={15} strokeWidth={1.8} aria-hidden="true" />
          <strong>Print</strong>
          <span className="print-dialog-document">{documentName}</span>
          <IconButton icon={X} label="Cancel printing" compact onClick={onClose} />
        </header>

        <div className="print-layout-body">
          <section className="print-options-pane" aria-label="Print options">
            <label className="print-field">
              <span>Printer</span>
              <select ref={printerSelectRef} value={deviceName} disabled={!printers?.length} onChange={(event) => setDeviceName(event.target.value)}>
                {Boolean(printers?.length) && <option value="">Default Windows printer</option>}
                {printers?.map((printer) => (
                  <option key={printer.name} value={printer.name}>{printer.displayName}</option>
                ))}
              </select>
              {printers === null && <small className="print-note">Looking for printers...</small>}
              {printers?.length === 0 && <small className="print-note">No printer is available. Add or enable one in Windows, then try again.</small>}
              {Boolean(printers?.length) && <small className="print-note">Print sends directly to this printer without opening another dialog.</small>}
            </label>

            <div className="print-field-row">
              <label className="print-field">
                <span>Copies</span>
                <input type="number" min={1} max={999} value={copies} onChange={(event) => setCopies(event.target.value)} onBlur={() => setCopies(String(copiesValue))} />
              </label>
              <label className="print-check"><input type="checkbox" checked={collate} disabled={copiesValue < 2} onChange={(event) => setCollate(event.target.checked)} />Collate</label>
            </div>

            <fieldset className="print-field">
              <span>Pages</span>
              <label className="print-range-option"><input type="radio" name="print-range" checked={rangeMode === 'all'} onChange={() => setRangeMode('all')} />All {pageCount} {pageCount === 1 ? 'page' : 'pages'}</label>
              <label className="print-range-option"><input type="radio" name="print-range" checked={rangeMode === 'current'} onChange={() => setRangeMode('current')} />Current page ({currentPage + 1})</label>
              {hasSelection && <label className="print-range-option"><input type="radio" name="print-range" checked={rangeMode === 'selected'} onChange={() => setRangeMode('selected')} />Selected pages ({selectedPages.length})</label>}
              <label className="print-range-option">
                <input type="radio" name="print-range" checked={rangeMode === 'custom'} onChange={() => setRangeMode('custom')} />Custom
                <input type="text" aria-label="Custom page range" placeholder="1-3, 5" value={customRange} onFocus={() => setRangeMode('custom')} onChange={(event) => setCustomRange(event.target.value)} />
              </label>
              {customInvalid && <small className="print-range-error">Use page numbers between 1 and {pageCount}, such as 1-3, 5.</small>}
            </fieldset>

            <div className="print-field-row">
              <label className="print-field">
                <span>Paper</span>
                <select ref={paperSelectRef} value={paperSize} onChange={(event) => setPaperSize(event.target.value as PrintPaperSize)}>
                  <option value="Letter">Letter (8.5 x 11 in)</option><option value="A4">A4 (210 x 297 mm)</option><option value="Legal">Legal (8.5 x 14 in)</option>
                </select>
              </label>
              <label className="print-field">
                <span>Orientation</span>
                <select value={orientation} onChange={(event) => setOrientation(event.target.value as PrintOrientation)}>
                  <option value="auto">Auto</option><option value="portrait">Portrait</option><option value="landscape">Landscape</option>
                </select>
              </label>
            </div>

            <div className="print-field-row">
              <label className="print-field">
                <span>Margins</span>
                <select value={marginMode} onChange={(event) => setMarginMode(event.target.value as PrintMarginMode)}>
                  <option value="normal">Normal (0.5 in)</option><option value="minimum">Minimum (0.25 in)</option><option value="none">None</option>
                </select>
              </label>
              <label className="print-field">
                <span>Color</span>
                <select value={color ? 'color' : 'bw'} onChange={(event) => setColor(event.target.value === 'color')}>
                  <option value="color">Color</option><option value="bw">Black and white</option>
                </select>
              </label>
            </div>

            <label className="print-field">
              <span>Page sizing</span>
              <select value={scaleMode} onChange={(event) => setScaleMode(event.target.value as PrintScaleMode)}>
                <option value="fit">Fit to printable area</option><option value="actual">Actual size</option><option value="shrink">Shrink oversized pages</option><option value="custom">Custom scale</option>
              </select>
            </label>
            {scaleMode === 'custom' && (
              <label className="print-field print-scale-field">
                <span>Custom scale</span>
                <div><input type="number" min={25} max={400} value={customScale} onChange={(event) => setCustomScale(event.target.value)} onBlur={() => setCustomScale(String(customScaleValue))} /><b>%</b></div>
              </label>
            )}

            {activePrinter?.supportsDuplex && (
              <label className="print-field">
                <span>Two-sided</span>
                <select value={duplexMode} onChange={(event) => setDuplexMode(event.target.value as PrintDuplexMode)}>
                  <option value="simplex">One-sided</option><option value="longEdge">Flip on long edge</option><option value="shortEdge">Flip on short edge</option>
                </select>
              </label>
            )}
          </section>

          <section className="print-preview-pane" aria-label="Live print preview">
            <div className="print-preview-toolbar">
              <div><strong>Print preview</strong><span>{paperSize} {landscape ? 'landscape' : 'portrait'} - {Math.round(placement.scale * 100)}%</span></div>
              <div className="print-preview-navigation">
                <IconButton icon={ChevronLeft} label="Previous print page" compact disabled={previewPosition <= 0} onClick={() => setPreviewPosition((value) => Math.max(0, value - 1))} />
                <span>{previewPosition + 1} / {previewIndices.length}</span>
                <IconButton icon={ChevronRight} label="Next print page" compact disabled={previewPosition >= previewIndices.length - 1} onClick={() => setPreviewPosition((value) => Math.min(previewIndices.length - 1, value + 1))} />
              </div>
            </div>
            <div ref={previewStageRef} className="print-preview-stage">
              <div key={`${paperSize}-${landscape}-${marginMode}-${scaleMode}-${customScaleValue}-${color}-${previewPageIndex}-${Math.round(previewPaperWidth || 0)}`} className="print-paper" style={{
                width: previewPaperWidth,
                height: previewPaperWidth ? previewPaperWidth / paperRatio : undefined,
                aspectRatio: `${placement.paperWidth} / ${placement.paperHeight}`,
              }}>
                <div className="print-margin-guide" style={{ inset: `${placement.margin / placement.paperHeight * 100}% ${placement.margin / placement.paperWidth * 100}%` }} />
                {activePreviewBitmap && (
                  <img
                    className="print-page-image"
                    src={activePreviewBitmap.dataUrl}
                    alt={`Preview of PDF page ${previewPageIndex + 1}`}
                    style={{
                      left: `${placement.x / placement.paperWidth * 100}%`,
                      top: `${placement.y / placement.paperHeight * 100}%`,
                      width: `${placement.contentWidth / placement.paperWidth * 100}%`,
                      height: `${placement.contentHeight / placement.paperHeight * 100}%`,
                      filter: color ? 'none' : 'grayscale(1)',
                    }}
                  />
                )}
                {!activePreviewBitmap && !previewError && <div className="print-preview-loading">Rendering page {previewPageIndex + 1}...</div>}
                {previewError && <div className="print-preview-error">{previewError}</div>}
              </div>
            </div>
            <div className="print-preview-summary" aria-live="polite">
              <span>PDF page {previewPageIndex + 1}</span>
              <span>{placement.cropped ? 'Content outside the paper will be cropped' : 'Content fits inside the selected margins'}</span>
              <small>Printer hardware can impose an additional non-printable edge. Simple sends this layout directly with the selected copies, color, collate, and duplex settings.</small>
            </div>
          </section>
        </div>

        <footer className="print-dialog-footer">
          <Button variant="secondary" onClick={onClose}>Cancel</Button>
          <Button icon={Printer} variant="primary" disabled={!canPrint} onClick={submit}>Print</Button>
        </footer>
      </div>
    </div>
  )
}
