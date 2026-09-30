import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import { CircleAlert, FileVideo2, LoaderCircle, Printer, X } from 'lucide-react'

interface PrintDialogProps {
  session: VideoPrintSession
  onClose: () => void
}

const DEFAULT_OPTIONS: VideoPrintOptions = {
  paperSize: 'letter',
  orientation: 'landscape',
  margins: 'normal',
  scaleMode: 'fit',
  customScale: 100,
  colorMode: 'color',
  metadata: true,
}

const FOCUSABLE = 'button:not([disabled]), select:not([disabled]), input:not([disabled]), [tabindex]:not([tabindex="-1"])'

export function PrintDialog({ session, onClose }: PrintDialogProps) {
  const [options, setOptions] = useState<VideoPrintOptions>(DEFAULT_OPTIONS)
  const [preview, setPreview] = useState<VideoPrintPreview | null>(null)
  const [previewFor, setPreviewFor] = useState('')
  const [previewing, setPreviewing] = useState(true)
  const [printing, setPrinting] = useState(false)
  const [error, setError] = useState('')
  const [previewZoom, setPreviewZoom] = useState(0.75)
  const dialogRef = useRef<HTMLElement>(null)
  const firstControlRef = useRef<HTMLSelectElement>(null)
  const previousFocusRef = useRef<HTMLElement | null>(null)
  const requestRef = useRef(0)
  const optionsKey = useMemo(() => JSON.stringify(options), [options])
  const previewReady = Boolean(preview && previewFor === optionsKey && !previewing)

  useEffect(() => {
    previousFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const frame = requestAnimationFrame(() => firstControlRef.current?.focus())
    return () => {
      cancelAnimationFrame(frame)
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
      void window.simpleVideo.renderPrintPreview({ sessionId: session.sessionId, options }).then((result) => {
        if (!active || requestRef.current !== requestId) return
        setPreview(result)
        setPreviewFor(optionsKey)
        setPreviewing(false)
      }).catch((reason) => {
        if (!active || requestRef.current !== requestId) return
        setError(reason instanceof Error ? reason.message : String(reason))
        setPreviewing(false)
      })
    }, 100)
    return () => {
      active = false
      clearTimeout(timer)
    }
  }, [options, optionsKey, session.sessionId])

  const previewHtml = useMemo(() => preview?.html.replace('--preview-zoom:1;', `--preview-zoom:${previewZoom};`) || '', [preview, previewZoom])

  const submit = async () => {
    if (!previewReady || printing) return
    setPrinting(true)
    setError('')
    try {
      const result = await window.simpleVideo.printFrame({ sessionId: session.sessionId, options })
      if (result.printed) {
        onClose()
        return
      }
      setPrinting(false)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message.replace(/^Error invoking remote method '[^']+':\s*/i, '') : String(reason))
      setPrinting(false)
    }
  }

  useEffect(() => {
    const guardPrintShortcuts = (event: KeyboardEvent) => {
      const key = event.key.toLocaleLowerCase()
      const control = event.ctrlKey || event.metaKey
      if (key === 'escape') {
        event.preventDefault()
        event.stopPropagation()
        if (!printing) onClose()
      } else if (control && key === 'p') {
        event.preventDefault()
        event.stopPropagation()
        if (previewReady && !printing) void submit()
      }
    }
    window.addEventListener('keydown', guardPrintShortcuts, true)
    return () => window.removeEventListener('keydown', guardPrintShortcuts, true)
  }, [onClose, previewReady, printing])

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLElement>) => {
    const control = event.ctrlKey || event.metaKey
    const key = event.key.toLocaleLowerCase()
    if (key === 'escape' && !printing) {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }
    if (control && key === 'p') {
      event.preventDefault()
      event.stopPropagation()
      if (previewReady && !printing) void submit()
      return
    }
    if (control && ['o', 'e', 'n', 's'].includes(key)) {
      event.preventDefault()
      event.stopPropagation()
      return
    }
    if (event.key !== 'Tab' || !dialogRef.current) return
    const focusable = [...dialogRef.current.querySelectorAll<HTMLElement>(FOCUSABLE)]
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

  const bindPreviewKeyboard = (frame: HTMLIFrameElement) => {
    if (!frame.contentWindow) return
    frame.contentWindow.onkeydown = (event) => {
      const control = event.ctrlKey || event.metaKey
      const key = event.key.toLocaleLowerCase()
      if (key === 'escape' && !printing) {
        event.preventDefault()
        onClose()
      } else if (control && key === 'p') {
        event.preventDefault()
        if (previewReady && !printing) void submit()
      } else if (event.key === 'Tab' && dialogRef.current) {
        event.preventDefault()
        const focusable = [...dialogRef.current.querySelectorAll<HTMLElement>(FOCUSABLE)]
          .filter((element) => element.getClientRects().length > 0)
        const target = event.shiftKey ? focusable.at(-1) : focusable[0]
        target?.focus()
      }
    }
  }

  const scalePercent = preview ? Math.round(preview.placement.scale * 100) : null

  return (
    <div className="modal-backdrop video-print-backdrop" onMouseDown={(event) => event.currentTarget === event.target && !printing && onClose()}>
      <section ref={dialogRef} className="video-print-dialog" role="dialog" aria-modal="true" aria-labelledby="video-print-title" aria-describedby="video-print-description" onKeyDown={handleKeyDown}>
        <header className="video-print-heading">
          <span className="modal-mark"><Printer /></span>
          <div><h2 id="video-print-title">Print current frame</h2><p id="video-print-description">Choose the page setup on the left and check the exact printed page on the right.</p></div>
          <button type="button" aria-label="Close print dialog" disabled={printing} onClick={onClose}><X /></button>
        </header>

        <div className="video-print-columns">
          <aside className="video-print-controls" aria-label="Print options">
            <div className="video-print-controls-scroll">
              <div className="video-print-source"><FileVideo2 /><span><strong>{session.sourceName}</strong><small>Frame at {formatFrameTime(session.seconds)} · {session.width} × {session.height}px</small></span></div>

              <fieldset className="video-print-group" disabled={printing}>
                <legend>Page setup</legend>
                <label><span>Paper</span><select ref={firstControlRef} aria-label="Print paper size" value={options.paperSize} onChange={(event) => setOptions((current) => ({ ...current, paperSize: event.target.value as VideoPrintOptions['paperSize'] }))}><option value="letter">Letter — 8.5 × 11 in</option><option value="a4">A4 — 210 × 297 mm</option><option value="legal">Legal — 8.5 × 14 in</option></select></label>
                <label><span>Orientation</span><select aria-label="Print orientation" value={options.orientation} onChange={(event) => setOptions((current) => ({ ...current, orientation: event.target.value as VideoPrintOptions['orientation'] }))}><option value="landscape">Landscape</option><option value="portrait">Portrait</option></select></label>
                <label><span>Margins</span><select aria-label="Print margins" value={options.margins} onChange={(event) => setOptions((current) => ({ ...current, margins: event.target.value as VideoPrintOptions['margins'] }))}><option value="normal">Normal — 0.5 in</option><option value="narrow">Narrow — 0.25 in</option><option value="wide">Wide — 0.75 in</option></select></label>
              </fieldset>

              <fieldset className="video-print-group" disabled={printing}>
                <legend>Frame placement</legend>
                <label><span>Size</span><select aria-label="Print scaling" value={options.scaleMode} onChange={(event) => setOptions((current) => ({ ...current, scaleMode: event.target.value as VideoPrintOptions['scaleMode'] }))}><option value="fit">Fit — show the whole frame</option><option value="fill">Fill — crop to the paper</option><option value="actual">Actual pixels at 96 PPI</option><option value="custom">Custom scale</option></select></label>
                <div className={`video-custom-scale${options.scaleMode === 'custom' ? '' : ' is-disabled'}`}>
                  <label htmlFor="video-print-custom-scale">Custom scale</label>
                  <input id="video-print-custom-scale" aria-label="Custom print scale" type="range" min="10" max="400" step="5" disabled={options.scaleMode !== 'custom'} value={options.customScale} onChange={(event) => setOptions((current) => ({ ...current, customScale: Number(event.target.value) }))} />
                  <label className="video-scale-number"><input aria-label="Custom scale percentage" type="number" min="10" max="400" step="5" disabled={options.scaleMode !== 'custom'} value={options.customScale} onChange={(event) => setOptions((current) => ({ ...current, customScale: Math.max(10, Math.min(400, Number(event.target.value) || 10)) }))} /><span>%</span></label>
                </div>
              </fieldset>

              <fieldset className="video-print-group video-output-options" disabled={printing}>
                <legend>Output</legend>
                <label><span>Color</span><select aria-label="Print color mode" value={options.colorMode} onChange={(event) => setOptions((current) => ({ ...current, colorMode: event.target.value as VideoPrintOptions['colorMode'] }))}><option value="color">Color</option><option value="grayscale">Black and white</option></select></label>
                <label className="video-print-check"><input type="checkbox" checked={options.metadata} onChange={(event) => setOptions((current) => ({ ...current, metadata: event.target.checked }))} /><span><strong>Frame details</strong><small>Print filename, time, and resolution below the image.</small></span></label>
              </fieldset>

              {preview?.placement.cropped && <p className="video-print-warning"><CircleAlert /> Parts of the frame are outside the printable area. Choose Fit to show the complete image.</p>}
              {error && <p className="video-print-error" role="alert">{error}</p>}
            </div>
            <footer className="video-print-actions"><span>Prints directly to your default printer.</span><div><button type="button" className="video-print-cancel" disabled={printing} onClick={onClose}>Cancel</button><button type="button" className="video-print-action" disabled={!previewReady || printing} onClick={() => void submit()}><Printer />{printing ? 'Printing…' : 'Print'}</button></div></footer>
          </aside>

          <section className="video-print-preview" aria-label="Live print preview" aria-busy={previewing}>
            <div className="video-print-preview-toolbar">
              <div><strong>{preview ? `${preview.page.label} · ${options.orientation === 'landscape' ? 'Landscape' : 'Portrait'}` : 'Preparing page…'}</strong>{preview && <small>{scalePercent}% frame scale{preview.placement.cropped ? ' · cropped' : ''}</small>}</div>
              <label><span>Preview</span><select aria-label="Preview zoom" disabled={printing} value={previewZoom} onChange={(event) => setPreviewZoom(Number(event.target.value))}><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option></select></label>
            </div>
            <div className="video-print-preview-surface">
              {preview && <iframe key={`${previewFor}-${previewZoom}`} className="video-print-preview-frame" title="Printed video frame page" tabIndex={-1} sandbox="allow-same-origin" srcDoc={previewHtml} onLoad={(event) => bindPreviewKeyboard(event.currentTarget)} />}
              {!preview && !error && <div className="video-print-preview-empty"><LoaderCircle /><span>Capturing the paper layout…</span></div>}
              {previewing && preview && <div className="video-print-preview-updating"><LoaderCircle />Updating preview…</div>}
              {!preview && error && <div className="video-print-preview-empty is-error"><CircleAlert /><span>The preview could not be generated.</span></div>}
            </div>
          </section>
        </div>
      </section>
    </div>
  )
}

function formatFrameTime(seconds: number) {
  const total = Math.max(0, Math.floor(Number.isFinite(seconds) ? seconds : 0))
  const hours = Math.floor(total / 3600)
  const minutes = Math.floor((total % 3600) / 60)
  const remainder = total % 60
  return hours ? `${hours}:${String(minutes).padStart(2, '0')}:${String(remainder).padStart(2, '0')}` : `${minutes}:${String(remainder).padStart(2, '0')}`
}
