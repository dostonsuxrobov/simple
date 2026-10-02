// src/simple/CropOptions.tsx (WP8)
// The contextual strip shown while Crop & rotate is open (design 3.1 / 4.2): aspect chips with an
// orientation swap, Straighten (-45..45 degrees, or "level by drawing a line"), Flip, Resize, the box size,
// Cancel and Apply crop. It renders after the main toolbar groups, so the first colour input and the first
// range input in the document stay the brush colour and size (qa contract). `.crop-actions` and
// `.apply-button` keep their names for the existing smokes.
import { useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent } from 'react'
import { Check, FlipHorizontal2, FlipVertical2, RotateCcwSquare, Ruler, Scaling } from 'lucide-react'
import type { Point } from '../imaging/types.ts'
import type { AspectPreset } from './cropMath.ts'
import { ASPECT_PRESETS, STRAIGHTEN_LIMIT, STRAIGHTEN_STEP, clampStraighten } from './cropMath.ts'
import { isPanEvent } from './useViewportNavigation.ts'

export interface CropOptionsProps {
  readonly aspect: AspectPreset
  readonly canSwap: boolean
  readonly straighten: number
  readonly leveling: boolean
  readonly size: { readonly width: number; readonly height: number }
  readonly busy: boolean
  readonly onAspect: (preset: AspectPreset) => void
  readonly onSwap: () => void
  readonly onStraighten: (degrees: number) => void
  readonly onLevel: () => void
  readonly onFlip: (horizontal: boolean) => void
  readonly onResize: () => void
  readonly onCancel: () => void
  readonly onApply: () => void
}

export function CropOptions(props: CropOptionsProps) {
  const { aspect, canSwap, straighten, leveling, size, busy } = props
  return (
    <div className="context-strip crop-strip" role="toolbar" aria-label="Crop and rotate">
      <div className="strip-group aspect-group" role="group" aria-label="Aspect ratio">
        {ASPECT_PRESETS.map((preset) => (
          <button
            key={preset.id}
            type="button"
            className={aspect === preset.id ? 'active' : ''}
            aria-pressed={aspect === preset.id}
            disabled={busy}
            data-aspect={preset.id}
            onClick={() => props.onAspect(preset.id)}
          >{preset.label}</button>
        ))}
        <button type="button" className="icon-only" title="Swap the box orientation" aria-label="Swap the box orientation" disabled={busy || !canSwap} onClick={props.onSwap}><RotateCcwSquare /></button>
      </div>
      <div className="strip-rule" />
      <div className="strip-group straighten-group">
        <label className="straighten-control" title="Straighten (drag, or use the arrow keys for 0.1 degree steps)">
          <span>Straighten</span>
          <input
            type="range"
            min={-STRAIGHTEN_LIMIT}
            max={STRAIGHTEN_LIMIT}
            step={STRAIGHTEN_STEP}
            value={straighten}
            disabled={busy}
            aria-label="Straighten angle"
            onChange={(event) => props.onStraighten(clampStraighten(Number(event.target.value)))}
            onDoubleClick={() => props.onStraighten(0)}
          />
          <output>{straighten > 0 ? '+' : ''}{straighten.toFixed(1)}°</output>
        </label>
        <button
          type="button"
          className={`icon-only ${leveling ? 'active' : ''}`}
          title="Level: draw a line along something that should be horizontal or vertical"
          aria-label="Level by drawing a line"
          aria-pressed={leveling}
          disabled={busy}
          onClick={props.onLevel}
        ><Ruler /></button>
      </div>
      <div className="strip-rule" />
      <div className="strip-group">
        <button type="button" className="icon-only" title="Flip horizontal" aria-label="Flip horizontal" disabled={busy} onClick={() => props.onFlip(true)}><FlipHorizontal2 /></button>
        <button type="button" className="icon-only" title="Flip vertical" aria-label="Flip vertical" disabled={busy} onClick={() => props.onFlip(false)}><FlipVertical2 /></button>
        <button type="button" title="Change the image size" disabled={busy} onClick={props.onResize}><Scaling /><span>Resize…</span></button>
      </div>
      <div className="crop-actions">
        <span>{Math.round(size.width)} × {Math.round(size.height)} px</span>
        <button type="button" disabled={busy} onClick={props.onCancel}>Cancel</button>
        <button type="button" className="apply-button" disabled={busy} onClick={props.onApply}><Check /> Apply crop</button>
      </div>
    </div>
  )
}

/**
 * "Level by drawing a line": drag along a horizon or a wall; on release the straighten angle makes that
 * line level. Coordinates are in the crop frame (the turned preview), `frame` px wide and high.
 */
export function LevelLineLayer(props: { readonly frame: { readonly width: number; readonly height: number }; readonly onLine: (from: Point, to: Point) => void; readonly onCancel: () => void }) {
  const { frame } = props
  const [line, setLine] = useState<{ from: Point; to: Point } | null>(null)
  const startRef = useRef<{ id: number; from: Point; clientX: number; clientY: number } | null>(null)
  const toFrame = (event: ReactPointerEvent<SVGSVGElement>): Point => {
    const rect = event.currentTarget.getBoundingClientRect()
    return {
      x: ((event.clientX - rect.left) * frame.width) / Math.max(1, rect.width),
      y: ((event.clientY - rect.top) * frame.height) / Math.max(1, rect.height),
    }
  }
  return (
    <svg
      className="level-line-layer"
      viewBox={`0 0 ${frame.width} ${frame.height}`}
      preserveAspectRatio="none"
      onPointerDown={(event) => {
        if (event.button !== 0 || !event.isPrimary || isPanEvent(event.nativeEvent)) return
        event.currentTarget.setPointerCapture(event.pointerId)
        const from = toFrame(event)
        startRef.current = { id: event.pointerId, from, clientX: event.clientX, clientY: event.clientY }
        setLine({ from, to: from })
      }}
      onPointerMove={(event) => {
        const start = startRef.current
        if (!start || start.id !== event.pointerId) return
        setLine({ from: start.from, to: toFrame(event) })
      }}
      onPointerUp={(event) => {
        const start = startRef.current
        if (!start || start.id !== event.pointerId) return
        startRef.current = null
        const to = toFrame(event)
        setLine(null)
        // Ignore clicks: a level line needs some length on screen.
        if (Math.hypot(event.clientX - start.clientX, event.clientY - start.clientY) < 12) return
        props.onLine(start.from, to)
      }}
      onPointerCancel={() => { startRef.current = null; setLine(null) }}
      onContextMenu={(event) => { event.preventDefault(); props.onCancel() }}
    >
      {line && (
        <>
          <line x1={line.from.x} y1={line.from.y} x2={line.to.x} y2={line.to.y} className="level-line-shadow" vectorEffect="non-scaling-stroke" />
          <line x1={line.from.x} y1={line.from.y} x2={line.to.x} y2={line.to.y} className="level-line" vectorEffect="non-scaling-stroke" />
        </>
      )}
    </svg>
  )
}
