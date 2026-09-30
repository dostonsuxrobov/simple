import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import { CircleAlert, FileSpreadsheet, LoaderCircle, Printer, X } from 'lucide-react'
import type { SpreadsheetPrintOptions, SpreadsheetPrintPreview, SpreadsheetPrintResult } from '../spreadsheet-types'

interface SpreadsheetPrintDialogProps {
  sheetCount: number
  selectionLabel: string
  defaultGridlines: boolean
  compatibilityWarning: boolean
  onClose: () => void
  onRenderPreview: (options: SpreadsheetPrintOptions) => Promise<SpreadsheetPrintPreview>
  onPrint: (options: SpreadsheetPrintOptions) => Promise<SpreadsheetPrintResult>
}

const DEFAULT_OPTIONS: SpreadsheetPrintOptions = {
  useSavedLayout: true,
  scope: 'active-sheet',
  orientation: 'portrait',
  scaling: 'fit-width',
  paperSize: 'letter',
  margins: 'normal',
  gridlines: true,
  headings: false,
}

const FOCUSABLE = 'button:not([disabled]), select:not([disabled]), input:not([disabled]), [tabindex]:not([tabindex="-1"])'

export function SpreadsheetPrintDialog({
  sheetCount,
  selectionLabel,
  defaultGridlines,
  compatibilityWarning,
  onClose,
  onRenderPreview,
  onPrint,
}: SpreadsheetPrintDialogProps) {
  const [options, setOptions] = useState<SpreadsheetPrintOptions>({ ...DEFAULT_OPTIONS, gridlines: defaultGridlines })
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
  const optionsKey = useMemo(() => JSON.stringify(options), [options])
  const previewReady = Boolean(preview && previewFor === optionsKey && !previewing)

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
    const requestId = ++requestRef.current
    let active = true
    setPreviewing(true)
    setError('')
    const timer = window.setTimeout(() => {
      void onRenderPreview(options).then((result) => {
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
  }, [onRenderPreview, options, optionsKey])

  const previewHtml = useMemo(() => preview?.html.replace('--preview-zoom:1;', `--preview-zoom:${previewZoom};`) || '', [preview, previewZoom])

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
    if (!previewReady || printing) return
    setPrinting(true)
    setError('')
    try {
      const result = await onPrint(options)
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
                  <select aria-label="Print margins" value={options.margins} onChange={(event) => setOptions((current) => ({ ...current, margins: event.target.value as SpreadsheetPrintOptions['margins'] }))}>
                    <option value="normal">Normal — 0.5 in</option>
                    <option value="narrow">Narrow — 0.25 in</option>
                    <option value="wide">Wide — 0.75 in</option>
                  </select>
                </label>
                <label>
                  <span>Scaling</span>
                  <select aria-label="Print scaling" value={options.scaling} onChange={(event) => setOptions((current) => ({ ...current, scaling: event.target.value as SpreadsheetPrintOptions['scaling'] }))}>
                    <option value="fit-width">Fit all columns on one page</option>
                    <option value="fit-sheet">Fit sheet on one page</option>
                    <option value="actual">Actual size</option>
                  </select>
                </label>
              </fieldset>}

              {!options.useSavedLayout && <fieldset className="print-options print-detail-options" disabled={printing}>
                <legend>Sheet details</legend>
                <label><input type="checkbox" checked={options.gridlines} onChange={(event) => setOptions((current) => ({ ...current, gridlines: event.target.checked }))} /><span><strong>Gridlines</strong><small>Show cell boundaries.</small></span></label>
                <label><input type="checkbox" checked={options.headings} onChange={(event) => setOptions((current) => ({ ...current, headings: event.target.checked }))} /><span><strong>Row and column headings</strong><small>Include A, B, C and 1, 2, 3.</small></span></label>
              </fieldset>}

              {compatibilityWarning && !preview?.warnings?.length && <div className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>Floating charts, drawings, or images may not appear. Cell values and formatting will print.</span></div>}
              {preview?.warnings?.map((warning) => <div key={warning} className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>{warning}</span></div>)}
              {preview && preview.minimumScale < 0.35 && <div className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>This layout scales content to {Math.round(preview.minimumScale * 100)}%. Consider landscape, narrower margins, or actual size for easier reading.</span></div>}
              {preview && preview.oversizedDimensions > 0 && <div className="print-warning"><CircleAlert size={15} aria-hidden="true" /><span>{preview.oversizedDimensions} oversized {preview.oversizedDimensions === 1 ? 'row or column is' : 'rows or columns are'} wider or taller than one printable page and will be clipped at Actual size. Choose Fit width or Fit sheet to show all content.</span></div>}
              {error && <div className="print-error" role="alert">{error}</div>}
            </div>

            <div className="print-controls-actions">
              <span className="print-native-note">Print sends these pages directly to your default Windows printer.</span>
              <div>
                <button type="button" className="secondary-action" disabled={printing} onClick={onClose}>Cancel</button>
                <button type="button" className="primary-action" disabled={!previewReady || printing} onClick={() => { void submit() }}><Printer size={14} />{printing ? 'Sending…' : 'Print now'}</button>
              </div>
            </div>
          </aside>

          <section className="print-preview-pane" aria-label="Live print preview" aria-busy={previewing}>
            <div className="print-preview-toolbar">
              <div className="print-preview-summary">
                <FileSpreadsheet size={15} aria-hidden="true" />
                <span>{preview ? `${preview.pages} ${preview.pages === 1 ? 'page' : 'pages'} · ${preview.paper.label} · ${(preview.options?.orientation || options.orientation) === 'portrait' ? 'Portrait' : 'Landscape'}` : 'Preparing pages…'}</span>
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
