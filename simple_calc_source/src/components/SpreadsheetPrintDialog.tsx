import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import { ChevronRight, CircleAlert, FileDown, FileSpreadsheet, LoaderCircle, Printer, X } from 'lucide-react'
import type { PrintMargins, PrintScaling, SpreadsheetPrintOptions, SpreadsheetPrintPreview, SpreadsheetPrintResult } from '../spreadsheet-types'
import {
  HEADER_FOOTER_FIELDS,
  HEADER_FOOTER_SECTIONS,
  columnLabel,
  headerFooterCode,
  headerFooterTexts,
  normalizeTitleColumns,
  normalizeTitleRows,
  rangeLabel,
} from '../lib/print-page-setup'
import type { HeaderFooterSection, SpreadsheetPageSetupPatch, SpreadsheetPageSetupState } from '../lib/print-page-setup'
import './print-dialog.css'

/** One previewed page: its sheet and the rows/columns it holds after repeated titles (0-based, inclusive). */
export interface SpreadsheetPreviewPage {
  pageNumber: number
  sheetId: string
  firstRow: number
  lastRow: number
  firstColumn: number
  lastColumn: number
}

/** The page model of a preview document (electron/spreadsheet-print.cjs section attributes). */
export function previewPages(html: string): SpreadsheetPreviewPage[] {
  const pages: SpreadsheetPreviewPage[] = []
  for (const match of html.matchAll(/<section class="print-page print-sheet"[^>]*>/g)) {
    const tag = match[0]
    const attribute = (name: string) => new RegExp(`${name}="([^"]*)"`).exec(tag)?.[1] ?? ''
    const rows = /^(\d+):(\d+)$/.exec(attribute('data-body-rows'))
    const columns = /^(\d+):(\d+)$/.exec(attribute('data-body-columns'))
    if (!rows || !columns) continue
    pages.push({
      pageNumber: Number(attribute('data-page-number')),
      sheetId: attribute('data-sheet-id').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'),
      firstRow: Number(rows[1]) - 1,
      lastRow: Number(rows[2]) - 1,
      firstColumn: Number(columns[1]) - 1,
      lastColumn: Number(columns[2]) - 1,
    })
  }
  return pages
}

/** A printer queue as main lists it (preload simpleCalc.listPrinters). */
export interface SpreadsheetPrinterInfo {
  name: string
  displayName: string
  description?: string
}

/** Scaling choices beyond the shared PrintScaling type (electron/spreadsheet-print.cjs). */
export type PrintScalingChoice = PrintScaling | 'fit-height' | 'custom' | 'fit-pages'
export type PrintMarginChoice = PrintMargins | 'custom'

export interface PrintCustomMargins {
  top: number
  right: number
  bottom: number
  left: number
}

/**
 * The options this dialog sends: SpreadsheetPrintOptions plus custom scale, fit-to-pages,
 * custom margins, a page range and the native job (printer, copies, collation). The print
 * engine validates all of them; the object is passed through onRenderPreview/onPrint as is.
 */
export type SpreadsheetPrintJobOptions = Omit<SpreadsheetPrintOptions, 'scaling' | 'margins'> & {
  scaling: PrintScalingChoice
  margins: PrintMarginChoice
  scalePercent?: number
  fitWidth?: number
  fitHeight?: number
  customMargins?: PrintCustomMargins
  pageRange?: { from: number; to?: number }
  printer?: { deviceName?: string; copies: number; collate: boolean }
}

interface SpreadsheetPrintDialogProps {
  sheetCount: number
  selectionLabel: string
  defaultGridlines: boolean
  compatibilityWarning: boolean
  onClose: () => void
  onRenderPreview: (options: SpreadsheetPrintOptions) => Promise<SpreadsheetPrintPreview>
  onPrint: (options: SpreadsheetPrintOptions) => Promise<SpreadsheetPrintResult>
  /**
   * Installed printers (defaults to window.simpleCalc.listPrinters). The printer, copies and
   * collation controls appear only when this resolves, i.e. when main honours them.
   */
  listPrinters?: () => Promise<SpreadsheetPrinterInfo[]>
  /** Adds "Save as PDF" at the top of the printer list; resolves true once the file is saved. */
  onSavePdf?: (options: SpreadsheetPrintOptions) => Promise<boolean>
  /** The active sheet's page setup; together with onPageSetupChange it enables the Page setup section. */
  pageSetup?: SpreadsheetPageSetupState
  /** Apply a page-setup change to the active sheet (saved with the workbook, undoable). */
  onPageSetupChange?: (patch: SpreadsheetPageSetupPatch) => void | Promise<void>
  /**
   * The previewed pages after every preview, e.g. to draw Excel's dashed page breaks on the
   * grid once the dialog closes (a page range reports only the pages in it).
   */
  onPreviewPages?: (pages: SpreadsheetPreviewPage[]) => void
}

const DEFAULT_OPTIONS: SpreadsheetPrintJobOptions = {
  useSavedLayout: true,
  scope: 'active-sheet',
  orientation: 'portrait',
  scaling: 'fit-width',
  paperSize: 'letter',
  margins: 'normal',
  gridlines: true,
  headings: false,
}

const DEFAULT_CUSTOM_MARGINS: PrintCustomMargins = { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 }
const PDF_DESTINATION = 'pdf'
const FOCUSABLE = 'button:not([disabled]), select:not([disabled]), input:not([disabled]), textarea:not([disabled]), summary, [tabindex]:not([tabindex="-1"])'

type PageDrafts = Record<'from' | 'to' | 'copies', string>

const INITIAL_PAGE_DRAFTS: PageDrafts = { from: '', to: '', copies: '1' }

/** A whole number in range from typed text; '' means "automatic"/unset; null is invalid. */
function wholeNumber(text: string, min: number, max: number): number | '' | null {
  const value = text.trim()
  if (!value) return ''
  if (!/^\d+$/.test(value)) return null
  const number = Number(value)
  return number >= min && number <= max ? number : null
}

/** Inches from typed text (a decimal comma is accepted). */
function inches(text: string): number | null {
  const value = text.trim().replace(',', '.')
  if (!/^\d+(?:\.\d+)?$|^\.\d+$/.test(value)) return null
  const number = Number(value)
  return number >= 0 && number <= 10 ? number : null
}

type LayoutDrafts = Record<'scalePercent' | 'fitWidth' | 'fitHeight' | 'top' | 'right' | 'bottom' | 'left', string>

const MARGIN_SIDES = ['top', 'bottom', 'left', 'right'] as const

/** What is wrong with the typed custom layout numbers that are in use, or ''. */
function layoutProblem(drafts: LayoutDrafts, scaling: PrintScalingChoice, margins: PrintMarginChoice): string {
  if (margins === 'custom' && MARGIN_SIDES.some((side) => inches(drafts[side]) === null)) return 'Enter each margin in inches, from 0 to 10.'
  if (scaling === 'custom' && typeof wholeNumber(drafts.scalePercent, 10, 400) !== 'number') return 'Enter a scale from 10% to 400%.'
  if (scaling === 'fit-pages' && (wholeNumber(drafts.fitWidth, 0, 100) === null || wholeNumber(drafts.fitHeight, 0, 100) === null)) return 'Enter the number of pages, or leave it blank for automatic.'
  return ''
}

interface CustomLayoutFieldsProps {
  scaling: PrintScalingChoice
  margins: PrintMarginChoice
  customMargins?: PrintCustomMargins
  /** Functional update of the dialog's options (pass the options state setter). */
  onChange: (update: (current: SpreadsheetPrintJobOptions) => SpreadsheetPrintJobOptions) => void
  /** Reports what is wrong with the typed numbers ('' when they are all valid). */
  onProblemChange: (problem: string) => void
}

/**
 * Custom margins (inches), Custom scale (%) and Fit to pages (wide by tall, blank =
 * automatic), shown for the matching Margins/Scaling choice. Valid numbers apply as they
 * are typed; invalid text is flagged and never sent.
 */
export function CustomLayoutFields({ scaling, margins, customMargins, onChange, onProblemChange }: CustomLayoutFieldsProps) {
  const [drafts, setDrafts] = useState<LayoutDrafts>(() => ({
    scalePercent: '100',
    fitWidth: '1',
    fitHeight: '',
    top: String(customMargins?.top ?? 0.5),
    right: String(customMargins?.right ?? 0.5),
    bottom: String(customMargins?.bottom ?? 0.5),
    left: String(customMargins?.left ?? 0.5),
  }))
  const problem = layoutProblem(drafts, scaling, margins)
  useEffect(() => { onProblemChange(problem) }, [onProblemChange, problem])

  const edit = (key: keyof LayoutDrafts, value: string) => {
    setDrafts((current) => ({ ...current, [key]: value }))
    if (key === 'scalePercent') {
      const number = wholeNumber(value, 10, 400)
      if (typeof number === 'number') onChange((current) => ({ ...current, scalePercent: number }))
    } else if (key === 'fitWidth' || key === 'fitHeight') {
      const number = wholeNumber(value, 0, 100)
      if (number !== null) onChange((current) => ({ ...current, [key]: number === '' ? 0 : number }))
    } else {
      const number = inches(value)
      if (number !== null) onChange((current) => ({ ...current, customMargins: { ...(current.customMargins ?? DEFAULT_CUSTOM_MARGINS), [key]: number } }))
    }
  }

  return (
    <>
      {margins === 'custom' && <div className="print-margin-fields" role="group" aria-label="Custom margins in inches">
        {MARGIN_SIDES.map((side) => {
          const label = side[0].toUpperCase() + side.slice(1)
          return (
            <label key={side}>
              <span>{label} (in)</span>
              <input type="number" inputMode="decimal" min={0} max={10} step={0.05} aria-label={`${label} margin in inches`} aria-invalid={inches(drafts[side]) === null || undefined} value={drafts[side]} onChange={(event) => edit(side, event.target.value)} />
            </label>
          )
        })}
      </div>}
      {scaling === 'custom' && <label className="print-span-field">
        <span>Scale (% of normal size)</span>
        <input type="number" inputMode="numeric" min={10} max={400} step={5} aria-label="Custom scale percent" aria-invalid={typeof wholeNumber(drafts.scalePercent, 10, 400) !== 'number' || undefined} value={drafts.scalePercent} onChange={(event) => edit('scalePercent', event.target.value)} />
      </label>}
      {scaling === 'fit-pages' && <div className="print-fit-fields" role="group" aria-label="Fit to pages">
        <label>
          <span>Pages wide</span>
          <input type="number" inputMode="numeric" min={0} max={100} aria-label="Pages wide" placeholder="Auto" aria-invalid={wholeNumber(drafts.fitWidth, 0, 100) === null || undefined} value={drafts.fitWidth} onChange={(event) => edit('fitWidth', event.target.value)} />
        </label>
        <span className="print-fit-by" aria-hidden="true">×</span>
        <label>
          <span>Pages tall</span>
          <input type="number" inputMode="numeric" min={0} max={100} aria-label="Pages tall" placeholder="Auto" aria-invalid={wholeNumber(drafts.fitHeight, 0, 100) === null || undefined} value={drafts.fitHeight} onChange={(event) => edit('fitHeight', event.target.value)} />
        </label>
      </div>}
    </>
  )
}

/** The page-shaping part of the options (what the preview and the PDF depend on). */
export function layoutOptionsOf(options: SpreadsheetPrintJobOptions): SpreadsheetPrintJobOptions {
  const result: SpreadsheetPrintJobOptions = {
    useSavedLayout: options.useSavedLayout,
    scope: options.scope,
    orientation: options.orientation,
    scaling: options.scaling,
    paperSize: options.paperSize,
    margins: options.margins,
    gridlines: options.gridlines,
    headings: options.headings,
  }
  if (options.scaling === 'custom') result.scalePercent = options.scalePercent ?? 100
  if (options.scaling === 'fit-pages') { result.fitWidth = options.fitWidth ?? 1; result.fitHeight = options.fitHeight ?? 0 }
  if (options.margins === 'custom') result.customMargins = options.customMargins ?? DEFAULT_CUSTOM_MARGINS
  if (options.pageRange) result.pageRange = options.pageRange
  return result
}

function defaultListPrinters(): Promise<SpreadsheetPrinterInfo[]> {
  const bridge = (typeof window !== 'undefined' ? window.simpleCalc : undefined) as unknown as { listPrinters?: () => Promise<SpreadsheetPrinterInfo[]> } | undefined
  if (!bridge?.listPrinters) return Promise.reject(new Error('Printer selection is not available.'))
  return bridge.listPrinters()
}

export function SpreadsheetPrintDialog({
  sheetCount,
  selectionLabel,
  defaultGridlines,
  compatibilityWarning,
  onClose,
  onRenderPreview,
  onPrint,
  listPrinters = defaultListPrinters,
  onSavePdf,
  pageSetup,
  onPageSetupChange,
  onPreviewPages,
}: SpreadsheetPrintDialogProps) {
  const [options, setOptions] = useState<SpreadsheetPrintJobOptions>({ ...DEFAULT_OPTIONS, gridlines: defaultGridlines })
  const [drafts, setDrafts] = useState<PageDrafts>(INITIAL_PAGE_DRAFTS)
  const [layoutError, setLayoutError] = useState('')
  const [printers, setPrinters] = useState<SpreadsheetPrinterInfo[] | null>(null)
  const [destination, setDestination] = useState('')
  const [collate, setCollate] = useState(true)
  const [preview, setPreview] = useState<SpreadsheetPrintPreview | null>(null)
  const [previewFor, setPreviewFor] = useState('')
  const [previewing, setPreviewing] = useState(true)
  const [printing, setPrinting] = useState(false)
  const [error, setError] = useState('')
  const [previewZoom, setPreviewZoom] = useState(0.75)
  const dialogRef = useRef<HTMLElement>(null)
  const scopeRef = useRef<HTMLSelectElement>(null)
  const previousFocusRef = useRef<HTMLElement | null>(null)
  const busyRef = useRef(false)
  const requestRef = useRef(0)
  const pageSetupEditable = Boolean(pageSetup && onPageSetupChange)

  // Only settings that change the pages are part of the preview key; the printer and copies are not.
  const layoutOptions = useMemo(() => layoutOptionsOf(options), [options])
  const pageSetupKey = useMemo(() => JSON.stringify(pageSetup ? { ...pageSetup, selection: undefined } : null), [pageSetup])
  const optionsKey = useMemo(() => `${JSON.stringify(layoutOptions)}|${pageSetupKey}`, [layoutOptions, pageSetupKey])
  const previewReady = Boolean(preview && previewFor === optionsKey && !previewing)
  const savingPdf = destination === PDF_DESTINATION && Boolean(onSavePdf)
  const copies = wholeNumber(drafts.copies, 1, 999)

  busyRef.current = printing

  useEffect(() => {
    previousFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const frame = window.requestAnimationFrame(() => scopeRef.current?.focus())
    return () => {
      window.cancelAnimationFrame(frame)
      const previous = previousFocusRef.current
      if (previous?.isConnected) previous.focus()
    }
  }, [])

  useEffect(() => {
    let active = true
    let request: Promise<SpreadsheetPrinterInfo[]>
    try {
      request = listPrinters()
    } catch (reason) {
      request = Promise.reject(reason)
    }
    request.then((list) => {
      if (active) setPrinters(Array.isArray(list) ? list.filter((printer) => printer && typeof printer.name === 'string' && printer.name) : [])
    }).catch(() => {
      // Without the printer list the job goes to the Windows default printer, as before.
      if (active) setPrinters(null)
    })
    return () => { active = false }
  }, [listPrinters])

  useEffect(() => {
    const requestId = ++requestRef.current
    let active = true
    setPreviewing(true)
    setError('')
    const timer = window.setTimeout(() => {
      void onRenderPreview(layoutOptions as unknown as SpreadsheetPrintOptions).then((result) => {
        if (!active || requestRef.current !== requestId) return
        setPreview(result)
        setPreviewFor(optionsKey)
        setPreviewing(false)
      }).catch((reason) => {
        if (!active || requestRef.current !== requestId) return
        setError(reason instanceof Error ? reason.message : String(reason))
        setPreviewing(false)
      })
    }, 120)
    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [onRenderPreview, layoutOptions, optionsKey])

  const previewPagesRef = useRef(onPreviewPages)
  previewPagesRef.current = onPreviewPages
  useEffect(() => {
    if (preview && previewPagesRef.current) previewPagesRef.current(previewPages(preview.html))
  }, [preview])

  const previewHtml = useMemo(() => preview?.html.replace('--preview-zoom:1;', `--preview-zoom:${previewZoom};`) || '', [preview, previewZoom])
  const totalPages = useMemo(() => {
    const match = preview ? /data-total-pages="(\d+)"/.exec(preview.html) : null
    return match ? Number(match[1]) : preview?.pages ?? 0
  }, [preview])
  const firstPage = useMemo(() => {
    const match = preview ? /data-first-page="(\d+)"/.exec(preview.html) : null
    return match ? Number(match[1]) : 1
  }, [preview])

  /** The page range applies as soon as it is valid (blank = from the first / to the last page). */
  const editPages = (key: 'from' | 'to', value: string) => {
    const next = { ...drafts, [key]: value }
    setDrafts(next)
    const from = wholeNumber(next.from, 1, 1_000_000)
    const to = wholeNumber(next.to, 1, 1_000_000)
    if (from === null || to === null) return
    const start = from === '' ? 1 : from
    if (to !== '' && to < start) return
    setOptions((current) => ({ ...current, pageRange: from === '' && to === '' ? undefined : to === '' ? { from: start } : { from: start, to } }))
  }

  const pagesInvalid = (() => {
    const from = wholeNumber(drafts.from, 1, 1_000_000)
    const to = wholeNumber(drafts.to, 1, 1_000_000)
    return from === null || to === null || (typeof from === 'number' && typeof to === 'number' && to < from)
  })()
  const copiesInvalid = Boolean(printers) && !savingPdf && typeof copies !== 'number'
  const inputProblem = pagesInvalid ? 'Enter a page range such as 2 to 5.' : copiesInvalid ? 'Enter 1 to 999 copies.' : !options.useSavedLayout ? layoutError : ''
  const inputsValid = !inputProblem

  const focusDialogEdge = (last: boolean) => {
    if (!dialogRef.current) return
    const controls = Array.from(dialogRef.current.querySelectorAll<HTMLElement>(FOCUSABLE))
      .filter((controlElement) => controlElement.getClientRects().length > 0)
    controls[last ? controls.length - 1 : 0]?.focus()
  }

  const bindPreviewKeyboard = (frame: HTMLIFrameElement) => {
    const frameDocument = frame.contentDocument
    if (!frameDocument) return
    frameDocument.onkeydown = (event) => {
      const control = event.ctrlKey || event.metaKey
      const key = event.key.toLocaleLowerCase()
      if (event.key === 'Tab') {
        event.preventDefault()
        focusDialogEdge(event.shiftKey)
      } else if (event.key === 'Escape') {
        event.preventDefault()
        if (!busyRef.current) onClose()
      } else if (control && key === 'p') {
        event.preventDefault()
        if (previewReady && !busyRef.current) void submit()
      } else if (control && ['o', 's', 'n', 'e'].includes(key)) {
        event.preventDefault()
      }
    }
  }

  const submit = async () => {
    if (!previewReady || printing || !inputsValid) return
    setPrinting(true)
    setError('')
    try {
      if (savingPdf && onSavePdf) {
        if (await onSavePdf(layoutOptions as unknown as SpreadsheetPrintOptions)) {
          onClose()
          return
        }
        setPrinting(false)
        return
      }
      const job: SpreadsheetPrintJobOptions = { ...layoutOptions }
      if (printers) {
        job.printer = {
          ...(destination.startsWith('printer:') ? { deviceName: destination.slice('printer:'.length) } : {}),
          copies: typeof copies === 'number' ? copies : 1,
          collate,
        }
      }
      const result = await onPrint(job as unknown as SpreadsheetPrintOptions)
      if (result.printed) {
        onClose()
        return
      }
      setPrinting(false)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setPrinting(false)
    }
  }

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLElement>) => {
    const control = event.ctrlKey || event.metaKey
    if (event.key === 'Escape' && !busyRef.current) {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }
    if (control && event.key.toLocaleLowerCase() === 'p') {
      event.preventDefault()
      event.stopPropagation()
      if (previewReady && !printing) void submit()
      return
    }
    if (control && ['o', 's', 'n', 'e'].includes(event.key.toLocaleLowerCase())) {
      event.preventDefault()
      event.stopPropagation()
      return
    }
    if (event.key !== 'Tab' || !dialogRef.current) return
    const controls = Array.from(dialogRef.current.querySelectorAll<HTMLElement>(FOCUSABLE))
      .filter((controlElement) => controlElement.getClientRects().length > 0)
    if (!controls.length) return
    const first = controls[0]
    const last = controls[controls.length - 1]
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault()
      last.focus()
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault()
      first.focus()
    }
  }

  const showDestination = Boolean(printers) || Boolean(onSavePdf)
  const pageSummary = preview
    ? totalPages > preview.pages
      ? `${preview.pages === 1 ? `Page ${firstPage}` : `Pages ${firstPage}–${firstPage + preview.pages - 1}`} of ${totalPages}`
      : `${preview.pages} ${preview.pages === 1 ? 'page' : 'pages'}`
    : ''
  const note = savingPdf
    ? 'Saves these pages as a PDF file on this computer.'
    : destination.startsWith('printer:')
      ? `Print sends these pages to ${printers?.find((printer) => `printer:${printer.name}` === destination)?.displayName || 'the chosen printer'}.`
      : 'Print sends these pages directly to your default Windows printer.'

  return (
    <div className="prompt-overlay print-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget && !printing) onClose() }}>
      <section ref={dialogRef} className="prompt-card print-card print-layout-dialog" role="dialog" aria-modal="true" aria-labelledby="print-dialog-title" onKeyDown={handleKeyDown}>
        <header className="print-card-header">
          <span className="print-card-icon"><Printer size={17} aria-hidden="true" /></span>
          <div>
            <h2 id="print-dialog-title">Print spreadsheet</h2>
            <p>Adjust the page on the left and verify the exact paper layout on the right.</p>
          </div>
          <button type="button" className="print-close" aria-label="Close print dialog" disabled={printing} onClick={onClose}><X size={15} /></button>
        </header>

        <div className="print-dialog-columns">
          <aside className="print-controls-pane" aria-label="Print options">
            <div className="print-controls-scroll">
              {showDestination && <fieldset className="print-control-group" disabled={printing}>
                <legend>Printer</legend>
                <label>
                  <span>Print to</span>
                  <select aria-label="Printer" value={destination} onChange={(event) => setDestination(event.target.value)}>
                    {onSavePdf && <option value={PDF_DESTINATION}>Save as PDF</option>}
                    <option value="">Default printer</option>
                    {printers?.map((printer) => <option key={printer.name} value={`printer:${printer.name}`} title={printer.description || undefined}>{printer.displayName}</option>)}
                  </select>
                </label>
                {printers && !savingPdf && <div className="print-inline-fields">
                  <label>
                    <span>Copies</span>
                    <input type="number" inputMode="numeric" min={1} max={999} aria-label="Copies" aria-invalid={copiesInvalid || undefined} value={drafts.copies} onChange={(event) => setDrafts((current) => ({ ...current, copies: event.target.value }))} />
                  </label>
                  {typeof copies === 'number' && copies > 1 && <label className="print-check">
                    <input type="checkbox" checked={collate} onChange={(event) => setCollate(event.target.checked)} />
                    <span>Collate</span>
                  </label>}
                </div>}
              </fieldset>}

              <fieldset className="print-control-group" disabled={printing}>
                <legend>Content</legend>
                <label>
                  <span>Print</span>
                  <select ref={scopeRef} aria-label="Print scope" value={options.scope} onChange={(event) => setOptions((current) => ({ ...current, scope: event.target.value as SpreadsheetPrintOptions['scope'] }))}>
                    <option value="active-sheet">Active sheet</option>
                    <option value="selection">Selected cells ({selectionLabel})</option>
                    <option value="workbook">Entire workbook ({sheetCount} {sheetCount === 1 ? 'sheet' : 'sheets'})</option>
                  </select>
                </label>
                <div className="print-range-field">
                  <span id="print-pages-label">Pages</span>
                  <div className="print-range-inputs" role="group" aria-labelledby="print-pages-label">
                    <input type="number" inputMode="numeric" min={1} aria-label="First page" placeholder="1" aria-invalid={pagesInvalid || undefined} value={drafts.from} onChange={(event) => editPages('from', event.target.value)} />
                    <span>to</span>
                    <input type="number" inputMode="numeric" min={1} aria-label="Last page" placeholder={totalPages ? String(totalPages) : 'last'} aria-invalid={pagesInvalid || undefined} value={drafts.to} onChange={(event) => editPages('to', event.target.value)} />
                  </div>
                </div>
              </fieldset>

              <fieldset className="print-options print-detail-options" disabled={printing}>
                <legend>Saved layout</legend>
                <label><input type="checkbox" aria-label="Use saved page layout" checked={options.useSavedLayout === true} onChange={(event) => setOptions((current) => ({ ...current, useSavedLayout: event.target.checked }))} /><span><strong>Use saved page layout</strong><small>Keep the file’s paper, margins and print scale. Turn off to adjust below.</small></span></label>
              </fieldset>

              {!options.useSavedLayout && <fieldset className="print-control-group print-page-controls" disabled={printing}>
                <legend>Page setup</legend>
                <label>
                  <span>Paper</span>
                  <select aria-label="Print paper size" value={options.paperSize} onChange={(event) => setOptions((current) => ({ ...current, paperSize: event.target.value as SpreadsheetPrintOptions['paperSize'] }))}>
                    <option value="letter">Letter — 8.5 × 11 in</option>
                    <option value="a4">A4 — 210 × 297 mm</option>
                    <option value="legal">Legal — 8.5 × 14 in</option>
                  </select>
                </label>
                <label>
                  <span>Orientation</span>
                  <select aria-label="Print orientation" value={options.orientation} onChange={(event) => setOptions((current) => ({ ...current, orientation: event.target.value as SpreadsheetPrintOptions['orientation'] }))}>
                    <option value="portrait">Portrait</option>
                    <option value="landscape">Landscape</option>
                  </select>
                </label>
                <label>
                  <span>Margins</span>
                  <select aria-label="Print margins" value={options.margins} onChange={(event) => {
                    const margins = event.target.value as PrintMarginChoice
                    setOptions((current) => ({ ...current, margins, ...(margins === 'custom' && !current.customMargins ? { customMargins: DEFAULT_CUSTOM_MARGINS } : {}) }))
                  }}>
                    <option value="normal">Normal — 0.5 in</option>
                    <option value="narrow">Narrow — 0.25 in</option>
                    <option value="wide">Wide — 0.75 in</option>
                    <option value="custom">Custom…</option>
                  </select>
                </label>
                <label>
                  <span>Scaling</span>
                  <select aria-label="Print scaling" value={options.scaling} onChange={(event) => setOptions((current) => ({ ...current, scaling: event.target.value as PrintScalingChoice }))}>
                    <option value="fit-width">Fit all columns on one page</option>
                    <option value="fit-height">Fit all rows on one page</option>
                    <option value="fit-sheet">Fit sheet on one page</option>
                    <option value="actual">Actual size</option>
                    <option value="custom">Custom scale…</option>
                    <option value="fit-pages">Fit to pages…</option>
                  </select>
                </label>
                <CustomLayoutFields scaling={options.scaling} margins={options.margins} customMargins={options.customMargins} onChange={setOptions} onProblemChange={setLayoutError} />
              </fieldset>}

              {!options.useSavedLayout && <fieldset className="print-options print-detail-options" disabled={printing}>
                <legend>Sheet details</legend>
                <label><input type="checkbox" checked={options.gridlines} onChange={(event) => setOptions((current) => ({ ...current, gridlines: event.target.checked }))} /><span><strong>Gridlines</strong><small>Show cell boundaries.</small></span></label>
                <label><input type="checkbox" checked={options.headings} onChange={(event) => setOptions((current) => ({ ...current, headings: event.target.checked }))} /><span><strong>Row and column headings</strong><small>Include A, B, C and 1, 2, 3.</small></span></label>
              </fieldset>}

              {pageSetupEditable && pageSetup && onPageSetupChange && (
                <SheetPageSetupSection pageSetup={pageSetup} disabled={printing} scope={options.scope} onChange={onPageSetupChange} />
              )}

              {compatibilityWarning && !preview?.warnings?.length && <div className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>Floating charts, drawings, or images may not appear. Cell values and formatting will print.</span></div>}
              {preview?.warnings?.map((warning) => <div key={warning} className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>{warning}</span></div>)}
              {preview && preview.minimumScale < 0.35 && <div className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>This layout scales content to {Math.round(preview.minimumScale * 100)}%. Consider landscape, narrower margins, or actual size for easier reading.</span></div>}
              {preview && preview.oversizedDimensions > 0 && <div className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>{preview.oversizedDimensions} oversized {preview.oversizedDimensions === 1 ? 'row or column is' : 'rows or columns are'} wider or taller than one printable page and will be clipped at this scale. Choose Fit width or Fit sheet to show all content.</span></div>}
              {inputProblem && <div className="print-error" role="alert">{inputProblem}</div>}
              {error && <div className="print-error" role="alert">{error}</div>}
            </div>

            <div className="print-controls-actions">
              <span className="print-native-note">{note}</span>
              <div>
                <button type="button" className="secondary-action" disabled={printing} onClick={onClose}>Cancel</button>
                <button type="button" className="primary-action" disabled={!previewReady || printing || !inputsValid} onClick={() => { void submit() }}>
                  {savingPdf ? <><FileDown size={14} />{printing ? 'Saving…' : 'Save PDF…'}</> : <><Printer size={14} />{printing ? 'Sending…' : 'Print now'}</>}
                </button>
              </div>
            </div>
          </aside>

          <section className="print-preview-pane" aria-label="Live print preview" aria-busy={previewing}>
            <div className="print-preview-toolbar">
              <div className="print-preview-summary">
                <FileSpreadsheet size={15} aria-hidden="true" />
                <span>{preview ? `${pageSummary} · ${preview.paper.label} · ${(preview.options?.orientation || options.orientation) === 'portrait' ? 'Portrait' : 'Landscape'}` : 'Preparing pages…'}</span>
                {preview && <small>{preview.pageBreaks} {preview.pageBreaks === 1 ? 'page break' : 'page breaks'} · {preview.cells.toLocaleString()} cells · {Math.round(preview.minimumScale * 100)}% print scale</small>}
              </div>
              <label className="print-preview-zoom">
                <span>Preview</span>
                <select aria-label="Preview zoom" disabled={printing} value={previewZoom} onChange={(event) => setPreviewZoom(Number(event.target.value))}>
                  <option value="0.5">50%</option>
                  <option value="0.75">75%</option>
                  <option value="1">100%</option>
                </select>
              </label>
            </div>
            <div className="print-preview-surface">
              {preview && <iframe key={`${previewFor}-${previewZoom}`} className="print-preview-frame" title="Spreadsheet print pages" tabIndex={-1} sandbox="allow-same-origin" srcDoc={previewHtml} onLoad={(event) => bindPreviewKeyboard(event.currentTarget)} />}
              {!preview && !error && <div className="print-preview-empty"><LoaderCircle className="print-preview-spinner" size={24} aria-hidden="true" /><span>Building the paper preview…</span></div>}
              {previewing && preview && <div className="print-preview-updating"><LoaderCircle className="print-preview-spinner" size={16} aria-hidden="true" />Updating preview…</div>}
              {!preview && error && <div className="print-preview-empty is-error"><CircleAlert size={24} aria-hidden="true" /><span>The preview could not be generated. Adjust the print range and try again.</span></div>}
            </div>
          </section>
        </div>
      </section>
    </div>
  )
}

interface SheetPageSetupSectionProps {
  pageSetup: SpreadsheetPageSetupState
  disabled: boolean
  scope: SpreadsheetPrintOptions['scope']
  onChange: (patch: SpreadsheetPageSetupPatch) => void | Promise<void>
}

/** Collapsed by default: print area, print titles, header/footer and manual page breaks. */
function SheetPageSetupSection({ pageSetup, disabled, scope, onChange }: SheetPageSetupSectionProps) {
  const [open, setOpen] = useState(false)
  const [rowsDraft, setRowsDraft] = useState(pageSetup.printTitlesRow)
  const [columnsDraft, setColumnsDraft] = useState(pageSetup.printTitlesColumn)
  const [header, setHeader] = useState(() => headerFooterTexts(pageSetup.oddHeader))
  const [footer, setFooter] = useState(() => headerFooterTexts(pageSetup.oddFooter))
  const [problem, setProblem] = useState('')
  const lastFieldRef = useRef<{ part: 'header' | 'footer'; section: HeaderFooterSection; element: HTMLTextAreaElement } | null>(null)

  // Re-read a setting when the sheet's value changes (after a commit, an undo, or a sheet switch).
  useEffect(() => { setRowsDraft(pageSetup.printTitlesRow) }, [pageSetup.printTitlesRow])
  useEffect(() => { setColumnsDraft(pageSetup.printTitlesColumn) }, [pageSetup.printTitlesColumn])
  useEffect(() => { setHeader(headerFooterTexts(pageSetup.oddHeader)) }, [pageSetup.oddHeader])
  useEffect(() => { setFooter(headerFooterTexts(pageSetup.oddFooter)) }, [pageSetup.oddFooter])

  const apply = (patch: SpreadsheetPageSetupPatch) => {
    setProblem('')
    void Promise.resolve(onChange(patch)).catch((reason) => setProblem(reason instanceof Error ? reason.message : String(reason)))
  }

  const commitTitles = (kind: 'rows' | 'columns', text: string) => {
    const normalized = kind === 'rows' ? normalizeTitleRows(text) : normalizeTitleColumns(text)
    if (normalized === null) {
      setProblem(kind === 'rows' ? 'Rows to repeat must look like 1 or 1:2.' : 'Columns to repeat must look like A or A:B.')
      return
    }
    const current = kind === 'rows' ? pageSetup.printTitlesRow : pageSetup.printTitlesColumn
    if (kind === 'rows') setRowsDraft(normalized)
    else setColumnsDraft(normalized)
    if (normalized === current) { setProblem(''); return }
    apply(kind === 'rows' ? { printTitlesRow: normalized || null } : { printTitlesColumn: normalized || null })
  }

  const commitHeaderFooter = (part: 'header' | 'footer', texts: Record<HeaderFooterSection, string>) => {
    const original = part === 'header' ? pageSetup.oddHeader : pageSetup.oddFooter
    const code = headerFooterCode(original, texts)
    if (code === original) return
    apply(part === 'header' ? { oddHeader: code } : { oddFooter: code })
  }

  const insertField = (token: string) => {
    const target = lastFieldRef.current
    const part = target?.part ?? 'header'
    const section = target?.section ?? 'center'
    const texts = { ...(part === 'header' ? header : footer) }
    const element = target?.element
    const start = element && element.isConnected ? element.selectionStart ?? texts[section].length : texts[section].length
    const end = element && element.isConnected ? element.selectionEnd ?? start : start
    texts[section] = texts[section].slice(0, start) + token + texts[section].slice(end)
    if (part === 'header') setHeader(texts)
    else setFooter(texts)
    commitHeaderFooter(part, texts)
    if (element && element.isConnected) window.requestAnimationFrame(() => {
      element.focus()
      element.setSelectionRange(start + token.length, start + token.length)
    })
  }

  const selection = pageSetup.selection
  const selectionText = rangeLabel(selection)
  const breakRow = selection.top + 1
  const hasBreak = pageSetup.rowBreaks.includes(breakRow)
  const frozenRows = pageSetup.frozenRows > 0 ? `1:${pageSetup.frozenRows}` : ''
  const frozenColumns = pageSetup.frozenColumns > 0 ? `A:${columnLabel(pageSetup.frozenColumns - 1)}` : ''

  const field = (part: 'header' | 'footer', section: HeaderFooterSection) => {
    const texts = part === 'header' ? header : footer
    const set = part === 'header' ? setHeader : setFooter
    const name = `${part === 'header' ? 'Header' : 'Footer'} ${section}`
    return (
      <textarea
        key={section}
        rows={1}
        aria-label={name}
        placeholder={section[0].toUpperCase() + section.slice(1)}
        value={texts[section]}
        onFocus={(event) => { lastFieldRef.current = { part, section, element: event.currentTarget } }}
        onChange={(event) => set({ ...texts, [section]: event.target.value })}
        onBlur={() => commitHeaderFooter(part, texts)}
      />
    )
  }

  return (
    <details className="print-advanced" open={open} onToggle={(event) => setOpen(event.currentTarget.open)}>
      <summary><ChevronRight size={13} aria-hidden="true" />Print area, titles and headers</summary>
      <fieldset className="print-advanced-body" disabled={disabled}>
        <div className="print-setup-row">
          <span className="print-setup-label">Print area of {pageSetup.sheetName}</span>
          <span className="print-setup-value">{pageSetup.printArea || 'Whole sheet (used cells)'}</span>
          <div className="print-setup-actions">
            <button type="button" className="print-small-button" disabled={pageSetup.printArea === selectionText} onClick={() => apply({ printArea: selectionText })}>Set to {selectionText}</button>
            {pageSetup.printArea && <button type="button" className="print-small-button" onClick={() => apply({ printArea: null })}>Clear</button>}
          </div>
          {scope === 'selection' && pageSetup.printArea && <small className="print-setup-note">Printing Selected cells ignores the print area.</small>}
        </div>

        <div className="print-setup-pair">
          <label>
            <span className="print-setup-label">Rows to repeat at top</span>
            <input type="text" aria-label="Rows to repeat at top" placeholder="e.g. 1:1" value={rowsDraft} onChange={(event) => setRowsDraft(event.target.value)} onBlur={(event) => commitTitles('rows', event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); commitTitles('rows', event.currentTarget.value) } }} />
          </label>
          <label>
            <span className="print-setup-label">Columns to repeat at left</span>
            <input type="text" aria-label="Columns to repeat at left" placeholder="e.g. A:A" value={columnsDraft} onChange={(event) => setColumnsDraft(event.target.value)} onBlur={(event) => commitTitles('columns', event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); commitTitles('columns', event.currentTarget.value) } }} />
          </label>
          {(frozenRows || frozenColumns) && <div className="print-setup-actions">
            {frozenRows && frozenRows !== pageSetup.printTitlesRow && <button type="button" className="print-small-button" onClick={() => apply({ printTitlesRow: frozenRows })}>Repeat frozen rows</button>}
            {frozenColumns && frozenColumns !== pageSetup.printTitlesColumn && <button type="button" className="print-small-button" onClick={() => apply({ printTitlesColumn: frozenColumns })}>Repeat frozen columns</button>}
          </div>}
        </div>

        <div className="print-setup-row">
          <span className="print-setup-label">Header</span>
          <div className="print-hf-grid">{HEADER_FOOTER_SECTIONS.map((section) => field('header', section))}</div>
          <span className="print-setup-label">Footer</span>
          <div className="print-hf-grid">{HEADER_FOOTER_SECTIONS.map((section) => field('footer', section))}</div>
          <div className="print-setup-actions print-chip-row" role="group" aria-label="Insert a field into the header or footer">
            {HEADER_FOOTER_FIELDS.map((item) => (
              <button key={item.token} type="button" className="print-chip" title={`Insert ${item.token}`} onMouseDown={(event) => event.preventDefault()} onClick={() => insertField(item.token)}>{item.label}</button>
            ))}
          </div>
          {(pageSetup.differentFirst || pageSetup.differentOddEven) && <small className="print-setup-note">The file’s own first-page and even-page headers are kept.</small>}
        </div>

        <div className="print-setup-row">
          <span className="print-setup-label">Manual page breaks</span>
          <span className="print-setup-value">{pageSetup.rowBreaks.length ? `Above row ${pageSetup.rowBreaks.join(', ')}` : 'None'}</span>
          <div className="print-setup-actions">
            {hasBreak
              ? <button type="button" className="print-small-button" onClick={() => apply({ rowBreaks: pageSetup.rowBreaks.filter((row) => row !== breakRow) })}>Remove break above row {breakRow}</button>
              : <button type="button" className="print-small-button" disabled={breakRow < 2} onClick={() => apply({ rowBreaks: [...pageSetup.rowBreaks, breakRow] })}>Insert break above row {breakRow}</button>}
            {pageSetup.rowBreaks.length > 0 && <button type="button" className="print-small-button" onClick={() => apply({ rowBreaks: [] })}>Remove all</button>}
          </div>
        </div>
        {problem && <div className="print-error" role="alert">{problem}</div>}
      </fieldset>
    </details>
  )
}

