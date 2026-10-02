// src/simple/ResizeDialog.tsx (WP8)
// Resize (design 4.3): width and height in pixels or percent, linked by default; presets 50%, 25% and long
// edges 3840 / 1920 / 1280 / 640; method "Automatic (best)" (area pre-filter, Lanczos-3 when shrinking,
// Catmull-Rom when enlarging, all premultiplied so transparent edges never darken). Shows the new
// megapixels, warns above 200%, and refuses sizes over 20,000 px per side or 50 MP.
import { useEffect, useMemo, useRef, useState } from 'react'
import { Link2, Link2Off, Scaling } from 'lucide-react'
import type { ResampleMethod } from '../imaging/types.ts'

export const RESIZE_MAX_DIMENSION = 20_000
export const RESIZE_MAX_PIXELS = 50_000_000

const METHODS: readonly { readonly value: ResampleMethod; readonly label: string }[] = [
  { value: 'auto', label: 'Automatic (best)' },
  { value: 'bicubic', label: 'Bicubic (smooth)' },
  { value: 'bilinear', label: 'Bilinear (fast)' },
  { value: 'nearest', label: 'Nearest neighbour (pixel art)' },
]

export interface ResizeDialogProps {
  readonly width: number
  readonly height: number
  readonly busy: boolean
  readonly progress: number | null
  readonly onCancel: () => void
  readonly onApply: (width: number, height: number, method: ResampleMethod) => void
}

type Unit = 'px' | '%'

function format(value: number): string {
  if (!Number.isFinite(value)) return ''
  return String(Math.round(value * 100) / 100)
}

/** Target size for a long-edge preset (never enlarges). */
export function longEdgeSize(width: number, height: number, edge: number): { width: number; height: number } {
  const scale = Math.min(1, edge / Math.max(width, height))
  return { width: Math.max(1, Math.round(width * scale)), height: Math.max(1, Math.round(height * scale)) }
}

export function validateResize(width: number, height: number): string | null {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) return 'Enter a width and height of at least 1 pixel.'
  if (width > RESIZE_MAX_DIMENSION || height > RESIZE_MAX_DIMENSION) return 'The limit is 20,000 pixels per side.'
  if (width * height > RESIZE_MAX_PIXELS) return 'The limit is 50 megapixels.'
  return null
}

export function ResizeDialog(props: ResizeDialogProps) {
  const { width: sourceWidth, height: sourceHeight, busy, progress } = props
  const [unit, setUnit] = useState<Unit>('px')
  const [linked, setLinked] = useState(true)
  const [widthText, setWidthText] = useState(String(sourceWidth))
  const [heightText, setHeightText] = useState(String(sourceHeight))
  const [method, setMethod] = useState<ResampleMethod>('auto')
  const widthRef = useRef<HTMLInputElement>(null)
  const restoreFocusRef = useRef<HTMLElement | null>(null)

  useEffect(() => {
    restoreFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
    widthRef.current?.focus()
    widthRef.current?.select()
    return () => {
      const target = restoreFocusRef.current
      if (target?.isConnected) window.requestAnimationFrame(() => target.focus())
    }
  }, [])

  const toPixels = (text: string, source: number) => {
    const value = Number(text)
    if (!Number.isFinite(value)) return NaN
    return unit === 'px' ? Math.round(value) : Math.round((source * value) / 100)
  }
  const target = useMemo(() => ({ width: toPixels(widthText, sourceWidth), height: toPixels(heightText, sourceHeight) }), [widthText, heightText, unit, sourceWidth, sourceHeight])
  const error = validateResize(target.width, target.height)
  const megapixels = !error ? (target.width * target.height) / 1_000_000 : 0
  const enlarging = !error && (target.width > sourceWidth * 2 || target.height > sourceHeight * 2)
  const unchanged = !error && target.width === sourceWidth && target.height === sourceHeight

  const setBoth = (width: number, height: number) => {
    if (unit === 'px') {
      setWidthText(String(width))
      setHeightText(String(height))
    } else {
      setWidthText(format((width / sourceWidth) * 100))
      setHeightText(format((height / sourceHeight) * 100))
    }
  }
  const changeWidth = (text: string) => {
    setWidthText(text)
    if (!linked) return
    const value = Number(text)
    if (!Number.isFinite(value) || value <= 0) return
    if (unit === '%') setHeightText(text)
    else setHeightText(String(Math.max(1, Math.round((value * sourceHeight) / sourceWidth))))
  }
  const changeHeight = (text: string) => {
    setHeightText(text)
    if (!linked) return
    const value = Number(text)
    if (!Number.isFinite(value) || value <= 0) return
    if (unit === '%') setWidthText(text)
    else setWidthText(String(Math.max(1, Math.round((value * sourceWidth) / sourceHeight))))
  }
  const changeUnit = (next: Unit) => {
    if (next === unit) return
    const width = Number.isFinite(target.width) ? target.width : sourceWidth
    const height = Number.isFinite(target.height) ? target.height : sourceHeight
    setUnit(next)
    if (next === 'px') {
      setWidthText(String(width))
      setHeightText(String(height))
    } else {
      setWidthText(format((width / sourceWidth) * 100))
      setHeightText(format((height / sourceHeight) * 100))
    }
  }
  const presets = [
    { label: '50%', size: { width: Math.max(1, Math.round(sourceWidth / 2)), height: Math.max(1, Math.round(sourceHeight / 2)) } },
    { label: '25%', size: { width: Math.max(1, Math.round(sourceWidth / 4)), height: Math.max(1, Math.round(sourceHeight / 4)) } },
    // Long-edge presets only when they shrink the image (an equal size would be a no-op).
    ...[3840, 1920, 1280, 640].filter((edge) => edge < Math.max(sourceWidth, sourceHeight)).map((edge) => ({ label: `${edge} px`, size: longEdgeSize(sourceWidth, sourceHeight, edge) })),
  ].filter((preset) => preset.size.width !== sourceWidth || preset.size.height !== sourceHeight)

  return (
    <div className="modal-backdrop">
      <form
        className="modal resize-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="resize-title"
        onSubmit={(event) => {
          event.preventDefault()
          if (busy || error || unchanged) return
          props.onApply(target.width, target.height, method)
        }}
        onKeyDown={(event) => {
          if (event.key === 'Escape' && !busy) {
            event.preventDefault()
            props.onCancel()
          }
        }}
      >
        <div className="modal-icon"><Scaling /></div>
        <h2 id="resize-title">Resize image</h2>
        <p>Now {sourceWidth.toLocaleString()} × {sourceHeight.toLocaleString()} px · {((sourceWidth * sourceHeight) / 1_000_000).toFixed(1)} MP</p>
        <div className="resize-presets" role="group" aria-label="Presets">
          {presets.map((preset) => (
            <button
              key={preset.label}
              type="button"
              disabled={busy}
              className={target.width === preset.size.width && target.height === preset.size.height ? 'active' : ''}
              onClick={() => { setLinked(true); setBoth(preset.size.width, preset.size.height) }}
              title={`${preset.size.width} × ${preset.size.height} px`}
            >{preset.label}</button>
          ))}
        </div>
        <div className="resize-fields">
          <label><span>Width</span><input ref={widthRef} type="number" min={unit === 'px' ? 1 : 0.01} step="any" inputMode="decimal" value={widthText} disabled={busy} onChange={(event) => changeWidth(event.target.value)} data-resize="width" /></label>
          <button
            type="button"
            className={`resize-link ${linked ? 'active' : ''}`}
            aria-pressed={linked}
            title={linked ? 'Width and height are linked' : 'Link width and height'}
            aria-label={linked ? 'Unlink width and height' : 'Link width and height'}
            disabled={busy}
            onClick={() => setLinked((value) => !value)}
          >{linked ? <Link2 /> : <Link2Off />}</button>
          <label><span>Height</span><input type="number" min={unit === 'px' ? 1 : 0.01} step="any" inputMode="decimal" value={heightText} disabled={busy} onChange={(event) => changeHeight(event.target.value)} data-resize="height" /></label>
          <label className="resize-unit"><span>Unit</span>
            <select value={unit} disabled={busy} onChange={(event) => changeUnit(event.target.value as Unit)}>
              <option value="px">Pixels</option>
              <option value="%">Percent</option>
            </select>
          </label>
        </div>
        <label className="resize-method"><span>Method</span>
          <select value={method} disabled={busy} onChange={(event) => setMethod(event.target.value as ResampleMethod)}>
            {METHODS.map((entry) => <option key={entry.value} value={entry.value}>{entry.label}</option>)}
          </select>
        </label>
        <div className="resize-summary" aria-live="polite">
          {error
            ? <span className="resize-error">{error}</span>
            : <span>New size: <strong>{target.width.toLocaleString()} × {target.height.toLocaleString()} px</strong> · {megapixels < 0.1 ? megapixels.toFixed(3) : megapixels.toFixed(1)} MP</span>}
          {enlarging && <span className="resize-warning">Enlarging more than 200% can look soft.</span>}
          {busy && <span className="resize-progress">Resizing… {Math.round((progress ?? 0) * 100)}%</span>}
        </div>
        <div className="modal-actions">
          <button type="button" disabled={busy} onClick={props.onCancel}>Cancel</button>
          <button type="submit" className="modal-primary" disabled={busy || Boolean(error) || unchanged}>{busy ? 'Resizing…' : 'Resize'}</button>
        </div>
      </form>
    </div>
  )
}
