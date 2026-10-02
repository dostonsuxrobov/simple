import { useEffect, useRef, useState } from 'react'
import { CircleAlert, Download, FileSpreadsheet, X } from 'lucide-react'
import type { SpreadsheetExportFormat, SpreadsheetPrintOptions } from '../spreadsheet-types'
import { CustomLayoutFields, layoutOptionsOf } from './SpreadsheetPrintDialog'
import type { PrintMarginChoice, PrintScalingChoice, SpreadsheetPrintJobOptions } from './SpreadsheetPrintDialog'
import './print-dialog.css'

interface SpreadsheetExportDialogProps {
  sheetCount: number
  selectionLabel: string
  defaultGridlines: boolean
  compatibilityWarning: boolean
  onClose: () => void
  onExport: (format: SpreadsheetExportFormat, options: SpreadsheetPrintOptions) => Promise<boolean>
}

const FORMATS: Array<{
  value: SpreadsheetExportFormat
  title: string
  extension: string
  description: string
}> = [
  { value: 'pdf', title: 'PDF document', extension: '.pdf', description: 'Print-ready pages with cell formatting.' },
  { value: 'xlsx', title: 'Excel workbook', extension: '.xlsx', description: 'All sheets, formulas, and common formatting.' },
  { value: 'xls', title: 'Excel 97–2003', extension: '.xls', description: 'A copy in the original Excel 97–2003 format.' },
  { value: 'html', title: 'Web page', extension: '.html', description: 'A standalone formatted page for any browser.' },
  { value: 'ods', title: 'OpenDocument', extension: '.ods', description: 'All sheets for LibreOffice and compatible apps.' },
  { value: 'csv', title: 'Comma-separated', extension: '.csv', description: 'Values from the active sheet only.' },
  { value: 'tsv', title: 'Tab-separated', extension: '.tsv', description: 'Values from the active sheet only.' },
]

export function SpreadsheetExportDialog({
  sheetCount,
  selectionLabel,
  defaultGridlines,
  compatibilityWarning,
  onClose,
  onExport,
}: SpreadsheetExportDialogProps) {
  const [format, setFormat] = useState<SpreadsheetExportFormat>('pdf')
  const [options, setOptions] = useState<SpreadsheetPrintJobOptions>({
    useSavedLayout: true,
    scope: 'workbook',
    orientation: 'portrait',
    scaling: 'fit-width',
    paperSize: 'letter',
    margins: 'normal',
    gridlines: defaultGridlines,
    headings: false,
  })
  const [exporting, setExporting] = useState(false)
  const [error, setError] = useState('')
  const [layoutError, setLayoutError] = useState('')
  const firstChoiceRef = useRef<HTMLButtonElement>(null)
  const pageFormat = format === 'pdf' || format === 'html'
  const activeSheetOnly = format === 'csv' || format === 'tsv'
  // Custom scale and margin fields only matter for page formats with a manual layout.
  const inputProblem = pageFormat && !options.useSavedLayout ? layoutError : ''

  useEffect(() => {
    firstChoiceRef.current?.focus()
    const handleKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !exporting) {
        event.preventDefault()
        onClose()
      }
    }
    window.addEventListener('keydown', handleKey)
    return () => window.removeEventListener('keydown', handleKey)
  }, [exporting, onClose])

  const submit = async () => {
    if (inputProblem) return
    setExporting(true)
    setError('')
    try {
      // The print engine validates the extended layout choices (custom scale, margins, fit to pages).
      if (await onExport(format, layoutOptionsOf(options) as unknown as SpreadsheetPrintOptions)) onClose()
      else setExporting(false)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setExporting(false)
    }
  }

  return (
    <div className="prompt-overlay export-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget && !exporting) onClose() }}>
      <section
        className="prompt-card export-card"
        role="dialog"
        aria-modal="true"
        aria-labelledby="export-dialog-title"
        onKeyDown={(event) => {
          if (event.key !== 'Tab') return
          const focusable = [...event.currentTarget.querySelectorAll<HTMLElement>('button:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])')]
            .filter((element) => element.getClientRects().length > 0)
          if (!focusable.length) return
          const first = focusable[0]
          const last = focusable[focusable.length - 1]
          if (event.shiftKey && document.activeElement === first) {
            event.preventDefault()
            last.focus()
          } else if (!event.shiftKey && document.activeElement === last) {
            event.preventDefault()
            first.focus()
          }
        }}
      >
        <header className="print-card-header">
          <span className="print-card-icon"><Download size={17} aria-hidden="true" /></span>
          <div>
            <h2 id="export-dialog-title">Export As</h2>
            <p>Create a new copy. Your open workbook and unsaved changes stay exactly as they are.</p>
          </div>
          <button type="button" className="print-close" aria-label="Close export options" disabled={exporting} onClick={onClose}><X size={15} /></button>
        </header>

        <div className="export-format-grid" role="radiogroup" aria-label="Export format">
          {FORMATS.map((item, index) => (
            <button
              key={item.value}
              ref={index === 0 ? firstChoiceRef : undefined}
              type="button"
              role="radio"
              aria-checked={format === item.value}
              data-export-format={item.value}
              className={`export-format-choice${format === item.value ? ' is-selected' : ''}`}
              disabled={exporting}
              onClick={() => { setFormat(item.value); setError('') }}
            >
              <span className="export-format-icon"><FileSpreadsheet size={17} aria-hidden="true" /></span>
              <span className="export-format-copy"><strong>{item.title}</strong><small>{item.description}</small></span>
              <span className="export-format-extension">{item.extension}</span>
            </button>
          ))}
        </div>

        {pageFormat && (
          <div className="export-page-options">
            <fieldset className="print-options" disabled={exporting}>
              <legend>Saved layout</legend>
              <label><input type="checkbox" aria-label="Use saved export layout" checked={options.useSavedLayout === true} onChange={(event) => setOptions(current => ({ ...current, useSavedLayout: event.target.checked }))} /><span><strong>Use saved page layout</strong><small>Keep the file’s paper, margins and print scale.</small></span></label>
            </fieldset>
            <div className="print-settings-grid">
              <label>
                <span>Include</span>
                <select aria-label="Export scope" value={options.scope} onChange={(event) => setOptions((current) => ({ ...current, scope: event.target.value as SpreadsheetPrintOptions['scope'] }))}>
                  <option value="workbook">Entire workbook ({sheetCount} {sheetCount === 1 ? 'sheet' : 'sheets'})</option>
                  <option value="active-sheet">Active sheet</option>
                  <option value="selection">Selected cells ({selectionLabel})</option>
                </select>
              </label>
              {!options.useSavedLayout && <><label>
                <span>Orientation</span>
                <select aria-label="Export orientation" disabled={options.useSavedLayout} value={options.orientation} onChange={(event) => setOptions((current) => ({ ...current, orientation: event.target.value as SpreadsheetPrintOptions['orientation'] }))}>
                  <option value="portrait">Portrait</option>
                  <option value="landscape">Landscape</option>
                </select>
              </label>
              <label>
                <span>Scaling</span>
                <select aria-label="Export scaling" disabled={options.useSavedLayout} value={options.scaling} onChange={(event) => setOptions((current) => ({ ...current, scaling: event.target.value as PrintScalingChoice }))}>
                  <option value="fit-width">Fit all columns on one page</option>
                  <option value="fit-height">Fit all rows on one page</option>
                  <option value="fit-sheet">Fit sheet on one page</option>
                  <option value="actual">Actual size</option>
                  <option value="custom">Custom scale…</option>
                  <option value="fit-pages">Fit to pages…</option>
                </select>
              </label>
              <label>
                <span>Paper</span>
                <select aria-label="Export paper size" disabled={options.useSavedLayout} value={options.paperSize} onChange={(event) => setOptions((current) => ({ ...current, paperSize: event.target.value as SpreadsheetPrintOptions['paperSize'] }))}>
                  <option value="letter">Letter</option>
                  <option value="a4">A4</option>
                  <option value="legal">Legal</option>
                </select>
              </label>
              <label>
                <span>Margins</span>
                <select aria-label="Export margins" disabled={options.useSavedLayout} value={options.margins} onChange={(event) => setOptions((current) => ({ ...current, margins: event.target.value as PrintMarginChoice }))}>
                  <option value="normal">Normal</option>
                  <option value="narrow">Narrow</option>
                  <option value="wide">Wide</option>
                  <option value="custom">Custom…</option>
                </select>
              </label>
              <CustomLayoutFields scaling={options.scaling} margins={options.margins} customMargins={options.customMargins} onChange={setOptions} onProblemChange={setLayoutError} /></>}
            </div>
            {!options.useSavedLayout && <fieldset className="print-options" disabled={exporting}>
              <legend>Sheet details</legend>
              <label><input type="checkbox" checked={options.gridlines} onChange={(event) => setOptions((current) => ({ ...current, gridlines: event.target.checked }))} /><span><strong>Gridlines</strong><small>Show cell boundaries.</small></span></label>
              <label><input type="checkbox" checked={options.headings} onChange={(event) => setOptions((current) => ({ ...current, headings: event.target.checked }))} /><span><strong>Row and column headings</strong><small>Include A, B, C and 1, 2, 3.</small></span></label>
            </fieldset>}
          </div>
        )}

        {activeSheetOnly && sheetCount > 1 && <div className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>This format exports only the active sheet. Other sheets, formatting, merges, formulas, charts, and images are not included.</span></div>}
        {format === 'ods' && <div className="export-note">OpenDocument keeps all worksheets, values, formulas, and common styles. Excel-only features may be simplified.</div>}
        {pageFormat && compatibilityWarning && <div className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>Floating charts, drawings, or images may not appear. Cell values and formatting are included.</span></div>}
        {inputProblem && <div className="print-error" role="alert">{inputProblem}</div>}
        {error && <div className="print-error" role="alert">{error}</div>}

        <div className="prompt-actions print-actions">
          <span className="prompt-hint">The export is a separate file and does not change the open workbook.</span>
          <button type="button" className="secondary-action" disabled={exporting} onClick={onClose}>Cancel</button>
          <button type="button" className="primary-action" disabled={exporting || Boolean(inputProblem)} onClick={() => { void submit() }}><Download size={14} />{exporting ? 'Exporting…' : `Export ${format.toUpperCase()}`}</button>
        </div>
      </section>
    </div>
  )
}
