import { useEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent } from 'react'
import { AlertTriangle, Printer, X } from 'lucide-react'
import {
  DEFAULT_PRINT_SETTINGS,
  PRINT_PAPERS,
  computePrintLayout,
  normalizePrintSettings,
  type ImagePrintSettings,
  type PrintPosition,
  type PrintScaleMode,
} from '../electron/print-layout.mjs'

interface PrintDialogProps {
  imageUrl: string
  imageName: string
  imageWidth: number
  imageHeight: number
  busy: boolean
  onPrint: (settings: ImagePrintSettings) => void
  onClose: () => void
}

const SETTINGS_KEY = 'simple-image:print-settings:v2'

function readSettings(): ImagePrintSettings {
  try {
    const stored = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}')
    return normalizePrintSettings(stored && typeof stored === 'object' ? stored : {})
  } catch {
    return { ...DEFAULT_PRINT_SETTINGS }
  }
}

function millimetres(value: number) {
  return `${value.toFixed(value < 10 ? 1 : 0)} mm`
}


interface DraftNumberProps {
  value: number
  min: number
  max: number
  step?: number
  setting: string
  onCommit: (value: number) => void
}

/** Typing is free (Backspace to empty, then a new number); the value is clamped only on blur or Enter. */
function DraftNumber({ value, min, max, step, setting, onCommit }: DraftNumberProps) {
  const [draft, setDraft] = useState<string | null>(null)
  const commit = () => {
    if (draft === null) return
    const parsed = Number(draft)
    onCommit(Number.isFinite(parsed) && draft.trim() !== '' ? Math.min(max, Math.max(min, parsed)) : value)
    setDraft(null)
  }
  return (
    <input
      data-print-setting={setting}
      type="number"
      min={min}
      max={max}
      step={step}
      value={draft ?? value}
      onChange={(event) => {
        setDraft(event.target.value)
        const parsed = Number(event.target.value)
        // In-range values update the preview right away; out-of-range ones wait for blur/Enter.
        if (event.target.value.trim() !== '' && Number.isFinite(parsed) && parsed >= min && parsed <= max) onCommit(parsed)
      }}
      onBlur={commit}
      onKeyDown={(event) => { if (event.key === 'Enter') commit() }}
    />
  )
}
export function PrintDialog({ imageUrl, imageName, imageWidth, imageHeight, busy, onPrint, onClose }: PrintDialogProps) {
  const [settings, setSettings] = useState<ImagePrintSettings>(readSettings)
  const [customMargin, setCustomMargin] = useState(![0, 6.35, 12.7, 25.4].includes(settings.marginMm))
  const [submitted, setSubmitted] = useState(false)
  const firstControlRef = useRef<HTMLSelectElement>(null)
  const dialogRef = useRef<HTMLDivElement>(null)
  const layout = useMemo(
    () => computePrintLayout(imageWidth, imageHeight, settings),
    [imageHeight, imageWidth, settings],
  )

  useEffect(() => {
    firstControlRef.current?.focus()
  }, [])

  useEffect(() => {
    if (!busy) setSubmitted(false)
  }, [busy])

  const locked = busy || submitted

  function update(patch: Partial<ImagePrintSettings>) {
    if (locked) return
    setSettings((current) => normalizePrintSettings({ ...current, ...patch }))
  }

  function submit() {
    if (locked) return
    try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)) } catch { /* Preferences are optional. */ }
    setSubmitted(true)
    onPrint(settings)
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    event.stopPropagation()
    const control = event.ctrlKey || event.metaKey
    if (event.key === 'Escape') {
      event.preventDefault()
      if (!locked) onClose()
      return
    }
    if (control && event.key.toLowerCase() === 'p') {
      event.preventDefault()
      submit()
      return
    }
    if (event.key !== 'Tab') return
    const focusable = [...(dialogRef.current?.querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled), select:not(:disabled)') || [])]
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
  }

  const printableLeft = layout.printable.xMm / layout.paper.widthMm * 100
  const printableTop = layout.printable.yMm / layout.paper.heightMm * 100
  const printableWidth = layout.printable.widthMm / layout.paper.widthMm * 100
  const printableHeight = layout.printable.heightMm / layout.paper.heightMm * 100
  const imageLeft = layout.image.xMm / layout.printable.widthMm * 100
  const imageTop = layout.image.yMm / layout.printable.heightMm * 100
  const imagePreviewWidth = layout.image.widthMm / layout.printable.widthMm * 100
  const imagePreviewHeight = layout.image.heightMm / layout.printable.heightMm * 100
  const scaleValue = settings.scaleMode === 'custom' ? settings.scalePercent : Math.round(96 / layout.effectiveDpi * 100)
  const marginPreset = customMargin ? 'custom' : String(settings.marginMm)

  return (
    <div
      className="image-print-overlay"
      onPointerDown={(event) => event.target === event.currentTarget && !locked && onClose()}
      onDragEnter={(event) => { event.preventDefault(); event.stopPropagation(); event.dataTransfer.dropEffect = 'none' }}
      onDragOver={(event) => { event.preventDefault(); event.stopPropagation(); event.dataTransfer.dropEffect = 'none' }}
      onDrop={(event) => { event.preventDefault(); event.stopPropagation() }}
    >
      <div
        ref={dialogRef}
        className="image-print-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="image-print-title"
        aria-describedby="image-print-description"
        onKeyDown={handleKeyDown}
      >
        <header className="image-print-header">
          <div>
            <span className="image-print-kicker"><Printer aria-hidden="true" /> Print image</span>
            <h2 id="image-print-title">Set up your print</h2>
            <p id="image-print-description">Choose the paper and placement on the left. The right side shows the resulting sheet.</p>
          </div>
          <button type="button" disabled={locked} aria-label="Close print setup" onClick={onClose}><X /></button>
        </header>

        <div className="image-print-body">
          <section className="image-print-controls" aria-label="Print options">
            <fieldset className="image-print-control-lock" disabled={locked} aria-label="Image print settings">
            <div className="print-control-section">
              <h3>Paper</h3>
              <div className="print-control-grid">
                <label>
                  <span>Size</span>
                  <select ref={firstControlRef} data-print-setting="paper" value={settings.paper} onChange={(event) => update({ paper: event.target.value as ImagePrintSettings['paper'] })}>
                    {Object.entries(PRINT_PAPERS).map(([key, paper]) => <option key={key} value={key}>{paper.label} · {paper.widthMm} × {paper.heightMm} mm</option>)}
                  </select>
                </label>
                <label>
                  <span>Orientation</span>
                  <select data-print-setting="orientation" value={settings.orientation} onChange={(event) => update({ orientation: event.target.value as ImagePrintSettings['orientation'] })}>
                    <option value="portrait">Portrait</option>
                    <option value="landscape">Landscape</option>
                  </select>
                </label>
              </div>
              <label>
                <span>Margins</span>
                <select
                  data-print-setting="margin-preset"
                  value={marginPreset}
                  onChange={(event) => {
                    if (event.target.value === 'custom') setCustomMargin(true)
                    else { setCustomMargin(false); update({ marginMm: Number(event.target.value) }) }
                  }}
                >
                  <option value="0">None</option>
                  <option value="6.35">Narrow · 6.35 mm</option>
                  <option value="12.7">Normal · 12.7 mm</option>
                  <option value="25.4">Wide · 25.4 mm</option>
                  <option value="custom">Custom</option>
                </select>
              </label>
              {customMargin && (
                <label>
                  <span>Margin on every side</span>
                  <div className="print-number-unit"><DraftNumber setting="margin" min={0} max={50} step={0.5} value={settings.marginMm} onCommit={(marginMm) => update({ marginMm })} /><span>mm</span></div>
                </label>
              )}
            </div>

            <div className="print-control-section">
              <h3>Image placement</h3>
              <label>
                <span>Scaling</span>
                <select data-print-setting="scale-mode" value={settings.scaleMode} onChange={(event) => update({ scaleMode: event.target.value as PrintScaleMode })}>
                  <option value="fit">Fit inside printable area</option>
                  <option value="fill">Fill printable area · crop edges</option>
                  <option value="actual">Actual pixels · 96 px/in</option>
                  <option value="custom">Custom actual-size scale</option>
                </select>
              </label>
              {settings.scaleMode === 'custom' && (
                <label>
                  <span>Custom scale</span>
                  <div className="print-number-unit"><DraftNumber setting="scale-percent" min={10} max={400} step={5} value={settings.scalePercent} onCommit={(scalePercent) => update({ scalePercent })} /><span>%</span></div>
                </label>
              )}
              <label>
                <span>Position</span>
                <select data-print-setting="position" value={settings.position} onChange={(event) => update({ position: event.target.value as PrintPosition })}>
                  <option value="center">Center</option>
                  <option value="top-left">Top left</option>
                  <option value="top-right">Top right</option>
                  <option value="bottom-left">Bottom left</option>
                  <option value="bottom-right">Bottom right</option>
                </select>
              </label>
            </div>

            <div className="print-control-section">
              <h3>Appearance</h3>
              <label>
                <span>Behind transparent pixels</span>
                <div className="print-background-control">
                  <button type="button" className={settings.background === '#ffffff' ? 'active' : ''} aria-label="White background" title="White" onClick={() => update({ background: '#ffffff' })} style={{ '--swatch': '#ffffff' } as CSSProperties} />
                  <button type="button" className={settings.background === '#000000' ? 'active' : ''} aria-label="Black background" title="Black" onClick={() => update({ background: '#000000' })} style={{ '--swatch': '#000000' } as CSSProperties} />
                  <input data-print-setting="background" type="color" aria-label="Custom print background" value={settings.background} onChange={(event) => update({ background: event.target.value })} />
                  <code>{settings.background.toUpperCase()}</code>
                </div>
              </label>
              <label className="print-check-control"><input data-print-setting="grayscale" type="checkbox" checked={settings.grayscale} onChange={(event) => update({ grayscale: event.target.checked })} /><span>Print image and transparency background in grayscale</span></label>
            </div>

            <div className="print-control-section print-last-section">
              <label>
                <span>Copies</span>
                <DraftNumber setting="copies" min={1} max={99} value={settings.copies} onCommit={(copies) => update({ copies })} />
              </label>
              <small>Simple sends this layout directly to your default Windows printer.</small>
            </div>
            </fieldset>
          </section>

          <section className="image-print-preview" aria-label="Live print preview">
            <div className="print-preview-heading">
              <div><strong>Print preview</strong><span>{layout.paper.label} · {layout.settings.orientation}</span></div>
              <span>{millimetres(layout.image.widthMm)} × {millimetres(layout.image.heightMm)}</span>
            </div>
            <div className="print-preview-stage">
              <div className="print-preview-sheet" data-paper={settings.paper} data-orientation={settings.orientation} style={{ aspectRatio: `${layout.paper.widthMm} / ${layout.paper.heightMm}` }}>
                <div className="print-preview-printable" style={{ left: `${printableLeft}%`, top: `${printableTop}%`, width: `${printableWidth}%`, height: `${printableHeight}%` }}>
                  <div className="print-preview-image-frame" style={{ left: `${imageLeft}%`, top: `${imageTop}%`, width: `${imagePreviewWidth}%`, height: `${imagePreviewHeight}%`, background: settings.background, filter: settings.grayscale ? 'grayscale(1)' : 'none' }}>
                    <img src={imageUrl} alt="Image positioned on the selected paper" />
                  </div>
                </div>
              </div>
            </div>
            <div className="print-preview-summary" aria-live="polite">
              <span>Printable area <strong>{millimetres(layout.printable.widthMm)} × {millimetres(layout.printable.heightMm)}</strong></span>
              <span>Effective resolution <strong>{Math.round(layout.effectiveDpi)} dpi</strong></span>
              <span>Scale <strong>{scaleValue}%</strong></span>
            </div>
            {layout.clipped && <p className="print-preview-warning"><AlertTriangle /> Part of the image is outside the printable area and will be cropped.</p>}
            <p className="print-preview-note">The sheet, margins, crop, position, background and grayscale treatment are passed to the print renderer exactly as shown. A printer may still impose its own unprintable edge.</p>
          </section>
        </div>

        <footer className="image-print-footer">
          <span>{imageName} · {imageWidth.toLocaleString()} × {imageHeight.toLocaleString()} px</span>
          <div>
            <button type="button" disabled={locked} onClick={onClose}>Cancel</button>
            <button type="button" className="print-submit" disabled={locked} onClick={submit}><Printer />{locked ? 'Sending to printer…' : 'Print now'}</button>
          </div>
        </footer>
      </div>
    </div>
  )
}
