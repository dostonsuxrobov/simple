import { memo } from 'react'
import type { ChartGeometry } from '../lib/charts'

export interface TraceArrow {
  /** Precedent range (or the traced cell, for dependents), 0-based. */
  from: { top: number; left: number; bottom: number; right: number }
  to: { row: number; col: number }
  /** A precedent or dependent on another sheet: drawn as a dashed arrow to/from a sheet icon. */
  external?: string
  /** Red when an error value flows along the arrow (Excel colours error paths red). */
  error?: boolean
  /** An off-sheet dependent: the arrow leaves the cell towards the sheet icon. */
  outgoing?: boolean
}

interface TraceArrowLayerProps {
  arrows: TraceArrow[]
  geometry: ChartGeometry
}

/** Excel's trace precedent/dependent arrows, drawn over the grid in canvas coordinates. */
export const TraceArrowLayer = memo(function TraceArrowLayer({ arrows, geometry }: TraceArrowLayerProps) {
  if (!arrows.length) return null
  let width = 0
  let height = 0
  const shapes = arrows.map((arrow, index) => {
    const first = geometry.cellRect(arrow.from.top, arrow.from.left)
    const last = geometry.cellRect(arrow.from.bottom, arrow.from.right)
    const box = { left: first.left, top: first.top, right: last.left + last.width, bottom: last.top + last.height }
    const target = geometry.cellRect(arrow.to.row, arrow.to.col)
    const end = { x: target.left + Math.min(12, target.width / 3), y: target.top + target.height / 2 }
    const range = arrow.from.bottom > arrow.from.top || arrow.from.right > arrow.from.left
    const start = arrow.external
      ? { x: Math.max(10, end.x - 90), y: Math.max(10, end.y - 46) }
      : { x: range ? box.left + 10 : first.left + Math.min(12, first.width / 3), y: range ? box.top + 10 : first.top + first.height / 2 }
    width = Math.max(width, box.right, target.left + target.width, start.x + 20) + 4
    height = Math.max(height, box.bottom, target.top + target.height, start.y + 20) + 4
    const color = arrow.error ? '#c0392b' : '#1f5fbf'
    return (
      <g key={index} className="trace-arrow" stroke={color} fill={color}>
        {range && !arrow.external && <rect x={box.left + 1} y={box.top + 1} width={Math.max(0, box.right - box.left - 2)} height={Math.max(0, box.bottom - box.top - 2)} fill="none" strokeWidth={1.5} />}
        {arrow.external
          ? (
            <g>
              <rect x={start.x - 9} y={start.y - 7} width={18} height={14} rx={2} fill="#fff" strokeWidth={1.2} />
              <path d={`M ${start.x - 5} ${start.y - 2} h 10 M ${start.x - 5} ${start.y + 2} h 10 M ${start.x - 1} ${start.y - 7} v 14`} strokeWidth={0.8} fill="none" />
              <title>{arrow.external}</title>
            </g>
          )
          : <circle cx={start.x} cy={start.y} r={2.6} stroke="none" />}
        {arrow.outgoing
          ? <line x1={end.x} y1={end.y} x2={start.x + 9} y2={start.y + 7} strokeWidth={1.6} strokeDasharray="5 3" markerEnd={`url(#trace-head-${arrow.error ? 'error' : 'normal'})`} />
          : <line x1={start.x} y1={start.y} x2={end.x} y2={end.y} strokeWidth={1.6} strokeDasharray={arrow.external ? '5 3' : undefined} markerEnd={`url(#trace-head-${arrow.error ? 'error' : 'normal'})`} />}
      </g>
    )
  })
  return (
    <svg className="trace-arrow-layer" width={width} height={height} aria-hidden="true">
      <defs>
        {(['normal', 'error'] as const).map((kind) => (
          <marker key={kind} id={`trace-head-${kind}`} viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
            <path d="M 0 0 L 10 5 L 0 10 z" fill={kind === 'error' ? '#c0392b' : '#1f5fbf'} />
          </marker>
        ))}
      </defs>
      {shapes}
    </svg>
  )
})
