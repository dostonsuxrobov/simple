// src/advanced/panels/CurvesEditor.tsx (WP6)
// Photoshop's Curves editor (design 5.10, 5.14): a 256 x 256 graph over the histogram with a channel menu
// (RGB, Red, Green, Blue). The curve is the natural cubic spline the adjustment applies (curveFunction).
//   - Click empty space to add a point (at most 16) and drag it; points keep their order (a point cannot
//     pass its neighbours). Drag a point off the graph to delete it (two points always stay).
//   - The selected point moves with the arrow keys (Shift = 10), Delete removes it, + and - select the
//     next / previous point; Input and Output show and edit its position.
//   - In RGB the per-channel curves show faintly in their colours.
import { useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, PointerEvent as ReactPointerEvent } from 'react'
import type { AdjustmentSpec, CurvePoint, Histogram } from '../../imaging/types.ts'
import { curveFunction } from '../../imaging/adjustments.ts'
import type { ChangePhase } from './AdjustmentControls.tsx'
import { NumberInput, SelectField, clampNumber } from './AdjustmentControls.tsx'
import { HistogramView } from './LevelsEditor.tsx'

type CurveKey = 'rgb' | 'red' | 'green' | 'blue'

const CHANNELS: readonly (readonly [CurveKey, string])[] = [['rgb', 'RGB'], ['red', 'Red'], ['green', 'Green'], ['blue', 'Blue']]
const CURVE_COLORS: Readonly<Record<CurveKey, string>> = Object.freeze({ rgb: '#151515', red: '#d33a3a', green: '#2f9a3f', blue: '#3460d6' })

export const MAX_CURVE_POINTS = 16
/** Graph margin in value units around 0..255 (handles at the edges stay grabbable). */
const PAD = 6
const VIEW = 255 + PAD * 2
/** Pointer distance (value units) that grabs a point. */
const HIT = 9
/** How far outside the graph a dragged point is deleted. */
const REMOVE_DISTANCE = 22

function sortPoints(points: readonly CurvePoint[]): CurvePoint[] {
  return [...points].sort((a, b) => a.x - b.x)
}

/** SVG path of a curve sampled at every input value. */
function curvePath(points: readonly CurvePoint[]): string {
  const curve = curveFunction(points)
  let path = ''
  for (let x = 0; x <= 255; x += 1) path += `${x === 0 ? 'M' : 'L'}${x} ${(255 - curve(x)).toFixed(2)}`
  return path
}

function isIdentity(points: readonly CurvePoint[]): boolean {
  return points.length === 2 && points[0].x === 0 && points[0].y === 0 && points[1].x === 255 && points[1].y === 255
}

export interface CurvesEditorProps {
  readonly spec: Extract<AdjustmentSpec, { type: 'curves' }>
  readonly onChange: (next: AdjustmentSpec, phase: ChangePhase) => void
  readonly histogram?: Histogram | null
  readonly disabled?: boolean
  readonly onInteractive?: (active: boolean) => void
}

export function CurvesEditor({ spec, onChange, histogram, disabled, onInteractive }: CurvesEditorProps) {
  const [key, setKey] = useState<CurveKey>('rgb')
  const [selected, setSelected] = useState<number | null>(null)
  const [removing, setRemoving] = useState(false)
  const svgRef = useRef<SVGSVGElement>(null)
  const specRef = useRef(spec)
  specRef.current = spec
  const points = useMemo(() => sortPoints(spec[key]), [key, spec])
  const selectedIndex = selected !== null && selected < points.length ? selected : null
  const point = selectedIndex !== null ? points[selectedIndex] : null

  const emit = (next: readonly CurvePoint[], phase: ChangePhase) => {
    const updated = { ...specRef.current, [key]: [...next] }
    specRef.current = updated
    onChange(updated, phase)
  }

  const toValue = (clientX: number, clientY: number) => {
    const rect = svgRef.current?.getBoundingClientRect()
    if (!rect || rect.width <= 0 || rect.height <= 0) return { x: 0, y: 0 }
    return {
      x: ((clientX - rect.left) / rect.width) * VIEW - PAD,
      y: 255 - (((clientY - rect.top) / rect.height) * VIEW - PAD),
    }
  }

  const onPointerDown = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (disabled || event.button !== 0) return
    event.preventDefault()
    const svg = event.currentTarget
    svg.focus({ preventScroll: true })
    const start = toValue(event.clientX, event.clientY)
    let list = sortPoints(specRef.current[key])
    let index = -1
    let best = HIT
    list.forEach((candidate, i) => {
      const distance = Math.max(Math.abs(candidate.x - start.x), Math.abs(candidate.y - start.y))
      if (distance <= best) {
        best = distance
        index = i
      }
    })
    if (index < 0) {
      if (start.x < 0 || start.x > 255 || start.y < 0 || start.y > 255) return
      const x = Math.round(clampNumber(start.x, 0, 255))
      const same = list.findIndex((candidate) => candidate.x === x)
      if (same >= 0) {
        index = same
      } else {
        if (list.length >= MAX_CURVE_POINTS) return
        const added = { x, y: Math.round(clampNumber(start.y, 0, 255)) }
        list = sortPoints([...list, added])
        index = list.indexOf(added)
        emit(list, 'live')
      }
    }
    setSelected(index)
    try {
      svg.setPointerCapture(event.pointerId)
    } catch {
      // ignore
    }
    onInteractive?.(true)
    const minX = index > 0 ? list[index - 1].x + 1 : 0
    const maxX = index < list.length - 1 ? list[index + 1].x - 1 : 255
    let current = list
    let gone = false
    const move = (moveEvent: PointerEvent) => {
      const p = toValue(moveEvent.clientX, moveEvent.clientY)
      const outside = p.x < -REMOVE_DISTANCE || p.x > 255 + REMOVE_DISTANCE || p.y < -REMOVE_DISTANCE || p.y > 255 + REMOVE_DISTANCE
      gone = outside && list.length > 2
      setRemoving(gone)
      if (gone) {
        current = list.filter((_, i) => i !== index)
      } else {
        const moved = { x: Math.round(clampNumber(p.x, minX, maxX)), y: Math.round(clampNumber(p.y, 0, 255)) }
        current = list.map((candidate, i) => (i === index ? moved : candidate))
      }
      emit(current, 'live')
    }
    const end = () => {
      svg.removeEventListener('pointermove', move)
      svg.removeEventListener('pointerup', end)
      svg.removeEventListener('pointercancel', end)
      setRemoving(false)
      onInteractive?.(false)
      if (gone) setSelected(null)
      emit(current, 'commit')
    }
    svg.addEventListener('pointermove', move)
    svg.addEventListener('pointerup', end)
    svg.addEventListener('pointercancel', end)
  }

  const moveSelected = (dx: number, dy: number) => {
    if (selectedIndex === null) return
    const list = points
    const minX = selectedIndex > 0 ? list[selectedIndex - 1].x + 1 : 0
    const maxX = selectedIndex < list.length - 1 ? list[selectedIndex + 1].x - 1 : 255
    const target = list[selectedIndex]
    const moved = { x: clampNumber(target.x + dx, minX, maxX), y: clampNumber(target.y + dy, 0, 255) }
    emit(list.map((candidate, i) => (i === selectedIndex ? moved : candidate)), 'commit')
  }

  const onKeyDown = (event: ReactKeyboardEvent<SVGSVGElement>) => {
    const big = event.shiftKey ? 10 : 1
    if (event.key === 'ArrowLeft') moveSelected(-big, 0)
    else if (event.key === 'ArrowRight') moveSelected(big, 0)
    else if (event.key === 'ArrowUp') moveSelected(0, big)
    else if (event.key === 'ArrowDown') moveSelected(0, -big)
    else if ((event.key === 'Delete' || event.key === 'Backspace') && selectedIndex !== null && points.length > 2) {
      emit(points.filter((_, i) => i !== selectedIndex), 'commit')
      setSelected(null)
    } else if (event.key === '+' || event.key === '=') {
      setSelected(selectedIndex === null ? 0 : (selectedIndex + 1) % points.length)
    } else if (event.key === '-') {
      setSelected(selectedIndex === null ? points.length - 1 : (selectedIndex - 1 + points.length) % points.length)
    } else {
      return
    }
    event.preventDefault()
    event.stopPropagation()
  }

  const setPoint = (patch: Partial<CurvePoint>) => {
    if (selectedIndex === null) return
    const minX = selectedIndex > 0 ? points[selectedIndex - 1].x + 1 : 0
    const maxX = selectedIndex < points.length - 1 ? points[selectedIndex + 1].x - 1 : 255
    const target = points[selectedIndex]
    const next = { x: clampNumber(patch.x ?? target.x, minX, maxX), y: clampNumber(patch.y ?? target.y, 0, 255) }
    emit(points.map((candidate, i) => (i === selectedIndex ? next : candidate)), 'commit')
  }

  const overlays = key === 'rgb'
    ? (['red', 'green', 'blue'] as const).filter((channel) => !isIdentity(sortPoints(spec[channel])))
    : []
  return (
    <div className="ae-curves">
      <div className="ae-curves-head">
        <SelectField label="Channel" inline value={key} options={CHANNELS} disabled={disabled} onChange={(next) => { setKey(next); setSelected(null) }} />
        <button
          type="button"
          className="ae-text-button"
          title="Reset this channel to a straight line"
          disabled={disabled || isIdentity(points)}
          onClick={() => { setSelected(null); emit([{ x: 0, y: 0 }, { x: 255, y: 255 }], 'commit') }}
        >Reset</button>
      </div>
      <div className="ae-curves-graph">
        <HistogramView histogram={histogram ?? null} channel={key === 'rgb' ? 'luma' : key} height={256} className="is-backdrop" />
        <svg
          ref={svgRef}
          viewBox={`${-PAD} ${-PAD} ${VIEW} ${VIEW}`}
          role="group"
          aria-label={`${key === 'rgb' ? 'RGB' : key} curve. Click to add a point, drag a point off the graph to remove it.`}
          tabIndex={disabled ? -1 : 0}
          className={removing ? 'is-removing' : undefined}
          onPointerDown={onPointerDown}
          onKeyDown={onKeyDown}
        >
          <rect x={0} y={0} width={255} height={255} className="ae-curves-frame" />
          {[64, 128, 191].map((v) => (
            <g key={v} className="ae-curves-grid">
              <line x1={v} y1={0} x2={v} y2={255} />
              <line x1={0} y1={v} x2={255} y2={v} />
            </g>
          ))}
          <line x1={0} y1={255} x2={255} y2={0} className="ae-curves-diagonal" />
          {overlays.map((channel) => <path key={channel} d={curvePath(spec[channel])} className="ae-curves-overlay" stroke={CURVE_COLORS[channel]} />)}
          <path d={curvePath(points)} className="ae-curves-line" stroke={CURVE_COLORS[key]} />
          {points.map((candidate, i) => (
            <rect
              key={`${i}-${candidate.x}`}
              x={candidate.x - 3.5}
              y={255 - candidate.y - 3.5}
              width={7}
              height={7}
              className={`ae-curves-point ${i === selectedIndex ? 'is-selected' : ''}`}
            />
          ))}
        </svg>
      </div>
      <div className="ae-curves-fields">
        <label className="ae-mini-field">
          <span>Input</span>
          <NumberInput label="Input value of the selected point" value={point ? point.x : 0} min={0} max={255} disabled={disabled || !point} onChange={(x) => setPoint({ x })} />
        </label>
        <label className="ae-mini-field">
          <span>Output</span>
          <NumberInput label="Output value of the selected point" value={point ? point.y : 0} min={0} max={255} disabled={disabled || !point} onChange={(y) => setPoint({ y })} />
        </label>
        <span className="ae-curves-count">{points.length} / {MAX_CURVE_POINTS}</span>
      </div>
    </div>
  )
}
