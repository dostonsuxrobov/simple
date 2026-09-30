import { memo, useEffect, useMemo, useRef, useState, type CSSProperties, type KeyboardEvent as ReactKeyboardEvent, type PointerEvent as ReactPointerEvent } from 'react'
import type { SheetChart } from '../spreadsheet-types'
import { anchorToRect, chartDisplayTitle, rectToAnchor, snapRectToCells, type ChartGeometry, type ChartRect, type ResolvedChartData } from '../lib/charts'
import { renderChartSvg, type ChartRenderOptions } from '../lib/chart-render'
import './charts.css'

export interface ChartLayerProps {
  /** Charts of the sheet being shown (sheet.charts). */
  charts: SheetChart[] | undefined
  /**
   * Grid geometry in the layer's coordinate space (the scrolled sheet canvas) at the current zoom.
   * May be a new object every render; it is read through a ref.
   */
  geometry: ChartGeometry
  /** Any value that changes when column widths, row heights or hidden rows/columns change. */
  geometryVersion?: unknown
  zoom: number
  selectedChartId: string | null
  /** Resolve a chart's data from calculated values (see resolveChartData). May change identity freely. */
  resolveData: (chart: SheetChart) => ResolvedChartData
  /** Bump when calculated values change; data is re-resolved only then (undefined = every render, de-duplicated). */
  dataVersion?: unknown
  onSelect: (id: string | null) => void
  /** Move/resize commits (anchor only; `modified` is left untouched). */
  onChange: (chart: SheetChart) => void
  onDelete: (id: string) => void
  /** Double-click / Enter: open the chart editor. */
  onEdit: (id: string) => void
  onContextMenu?: (id: string, position: { clientX: number; clientY: number }) => void
  readOnly?: boolean
  renderOptions?: ChartRenderOptions
  /** Visible canvas rectangle; charts outside it are not mounted (the selected chart always is). */
  viewport?: ChartRect
  /** Deselect when the user presses outside charts and outside `.chart-editor` / [data-chart-keep-selection]. Default true. */
  deselectOnOutsidePointer?: boolean
  className?: string
  style?: CSSProperties
}

type Handle = 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw'
const HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
const MIN_SIZE = 24

interface FrameCallbacks {
  select(id: string): void
  commit(chart: SheetChart, rect: ChartRect, snap: boolean): void
  remove(id: string): void
  edit(id: string): void
  deselect(): void
  contextMenu(id: string, clientX: number, clientY: number): void
}

interface FrameProps {
  chart: SheetChart
  data: ResolvedChartData
  left: number
  top: number
  width: number
  height: number
  zoom: number
  selected: boolean
  readOnly: boolean
  renderOptions?: ChartRenderOptions
  callbacks: FrameCallbacks
  z: number
}

interface Interaction {
  pointerId: number
  mode: 'move' | Handle
  startX: number
  startY: number
  origin: ChartRect
  active: boolean
}

function resizeRect(origin: ChartRect, handle: Handle, dx: number, dy: number, keepAspect: boolean, minSize: number): ChartRect {
  let { left, top, width, height } = origin
  if (handle.includes('e')) width = Math.max(minSize, origin.width + dx)
  if (handle.includes('s')) height = Math.max(minSize, origin.height + dy)
  if (handle.includes('w')) { width = Math.max(minSize, origin.width - dx); left = origin.left + origin.width - width }
  if (handle.includes('n')) { height = Math.max(minSize, origin.height - dy); top = origin.top + origin.height - height }
  if (keepAspect && handle.length === 2) {
    const ratio = origin.width / Math.max(1, origin.height)
    if (width / Math.max(1, height) > ratio) width = height * ratio
    else height = width / ratio
    if (handle.includes('w')) left = origin.left + origin.width - width
    if (handle.includes('n')) top = origin.top + origin.height - height
  }
  return { left: Math.max(0, left), top: Math.max(0, top), width, height }
}

const ChartFrame = memo(function ChartFrame({ chart, data, left, top, width, height, zoom, selected, readOnly, renderOptions, callbacks, z }: FrameProps) {
  const frameRef = useRef<HTMLDivElement | null>(null)
  const interaction = useRef<Interaction | null>(null)
  const [draft, setDraft] = useState<ChartRect | null>(null)
  const rect = draft || { left, top, width, height }
  // The SVG is drawn in unzoomed pixels and scaled by CSS, so zooming never re-renders it.
  const svgWidth = Math.max(20, Math.round(rect.width / Math.max(zoom, 0.05)))
  const svgHeight = Math.max(20, Math.round(rect.height / Math.max(zoom, 0.05)))
  const svg = useMemo(() => {
    try {
      return renderChartSvg(chart, data, svgWidth, svgHeight, renderOptions)
    } catch (error) {
      console.error('Chart render failed', error)
      return ''
    }
  }, [chart, data, svgWidth, svgHeight, renderOptions])

  useEffect(() => {
    if (selected && document.activeElement !== frameRef.current && !frameRef.current?.contains(document.activeElement)) {
      frameRef.current?.focus({ preventScroll: true })
    }
  }, [selected])

  const begin = (event: ReactPointerEvent<HTMLElement>, mode: Interaction['mode']) => {
    if (event.button !== 0) return
    event.stopPropagation()
    if (!selected) callbacks.select(chart.id)
    if (readOnly) return
    event.preventDefault()
    try { frameRef.current?.setPointerCapture(event.pointerId) } catch { /* pointer already released */ }
    interaction.current = { pointerId: event.pointerId, mode, startX: event.clientX, startY: event.clientY, origin: { left, top, width, height }, active: mode !== 'move' }
  }

  const onPointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const current = interaction.current
    if (!current || current.pointerId !== event.pointerId) return
    const dx = event.clientX - current.startX
    const dy = event.clientY - current.startY
    if (!current.active && Math.hypot(dx, dy) < 3) return
    current.active = true
    if (current.mode === 'move') setDraft({ ...current.origin, left: Math.max(0, current.origin.left + dx), top: Math.max(0, current.origin.top + dy) })
    else setDraft(resizeRect(current.origin, current.mode, dx, dy, event.shiftKey, MIN_SIZE * zoom))
  }

  const finish = (event: ReactPointerEvent<HTMLDivElement>, cancel = false) => {
    const current = interaction.current
    if (!current || current.pointerId !== event.pointerId) return
    interaction.current = null
    try { if (frameRef.current?.hasPointerCapture(event.pointerId)) frameRef.current.releasePointerCapture(event.pointerId) } catch { /* already released */ }
    const next = draft
    setDraft(null)
    if (!cancel && current.active && next) callbacks.commit(chart, next, event.altKey)
  }

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.target !== frameRef.current) return
    if (event.key === 'Delete' || event.key === 'Backspace') {
      if (readOnly) return
      event.preventDefault(); event.stopPropagation()
      callbacks.remove(chart.id)
    } else if (event.key === 'Enter' || event.key === 'F2') {
      event.preventDefault(); event.stopPropagation()
      callbacks.edit(chart.id)
    } else if (event.key === 'Escape') {
      event.preventDefault(); event.stopPropagation()
      if (interaction.current) { interaction.current = null; setDraft(null) } else callbacks.deselect()
    } else if (!readOnly && event.key.startsWith('Arrow')) {
      event.preventDefault(); event.stopPropagation()
      const step = (event.shiftKey ? 10 : 1) * zoom
      const dx = event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0
      const dy = event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0
      if (event.ctrlKey || event.metaKey) callbacks.commit(chart, { left, top, width: Math.max(MIN_SIZE * zoom, width + dx), height: Math.max(MIN_SIZE * zoom, height + dy) }, false)
      else callbacks.commit(chart, { left: Math.max(0, left + dx), top: Math.max(0, top + dy), width, height }, false)
    }
  }

  const title = chartDisplayTitle(chart, data) || 'Chart'
  return (
    <div
      ref={frameRef}
      className={`chart-frame${selected ? ' is-selected' : ''}${draft ? ' is-dragging' : ''}${readOnly ? ' is-readonly' : ''}`}
      style={{ left: rect.left, top: rect.top, width: rect.width, height: rect.height, zIndex: selected ? 1000 + z : z + 1 }}
      data-chart-id={chart.id}
      role="group"
      aria-label={`${title} (chart)`}
      tabIndex={selected ? 0 : -1}
      onPointerDown={(event) => begin(event, 'move')}
      onPointerMove={onPointerMove}
      onPointerUp={(event) => finish(event)}
      onPointerCancel={(event) => finish(event, true)}
      onLostPointerCapture={(event) => { if (interaction.current?.pointerId === event.pointerId) finish(event) }}
      onDoubleClick={(event) => { event.stopPropagation(); callbacks.edit(chart.id) }}
      onContextMenu={(event) => { event.preventDefault(); event.stopPropagation(); if (!selected) callbacks.select(chart.id); callbacks.contextMenu(chart.id, event.clientX, event.clientY) }}
      onKeyDown={onKeyDown}
    >
      <div className="chart-frame-svg" dangerouslySetInnerHTML={{ __html: svg }} />
      {selected && !readOnly && HANDLES.map((handle) => (
        <div key={handle} className={`chart-handle chart-handle-${handle}`} data-handle={handle} onPointerDown={(event) => begin(event, handle)} aria-hidden="true" />
      ))}
    </div>
  )
})

function dataSignature(data: ResolvedChartData) {
  return JSON.stringify(data)
}

/**
 * Absolutely positioned chart frames over the grid. The layer itself ignores
 * pointer events so the grid underneath stays interactive; only chart frames
 * capture them. Place it inside the scrolled sheet canvas.
 */
export function ChartLayer(props: ChartLayerProps) {
  const { charts, zoom, selectedChartId, readOnly = false, renderOptions, viewport, className, style } = props
  const latest = useRef(props)
  latest.current = props
  const dataCache = useRef(new WeakMap<SheetChart, { version: unknown; data: ResolvedChartData; signature: string }>())

  const callbacks = useMemo<FrameCallbacks>(() => ({
    select: (id) => latest.current.onSelect(id),
    deselect: () => latest.current.onSelect(null),
    remove: (id) => latest.current.onDelete(id),
    edit: (id) => latest.current.onEdit(id),
    contextMenu: (id, clientX, clientY) => latest.current.onContextMenu?.(id, { clientX, clientY }),
    commit: (chart, rect, snap) => {
      const { geometry, zoom: currentZoom, onChange } = latest.current
      const target = snap ? snapRectToCells(rect, geometry) : rect
      const anchor = rectToAnchor(target, geometry, currentZoom, chart.anchor?.editAs)
      const before = chart.anchor
      if (JSON.stringify(before) === JSON.stringify(anchor)) return
      onChange({ ...chart, anchor })
    },
  }), [])

  // Deselect on outside presses (capture phase so grid handlers still run normally).
  useEffect(() => {
    if (!selectedChartId || props.deselectOnOutsidePointer === false) return
    const listener = (event: PointerEvent) => {
      const target = event.target as Element | null
      if (!target || !(target instanceof Element)) return
      if (target.closest('.chart-frame, .chart-editor, [data-chart-keep-selection]')) return
      latest.current.onSelect(null)
    }
    document.addEventListener('pointerdown', listener, true)
    return () => document.removeEventListener('pointerdown', listener, true)
  }, [selectedChartId, props.deselectOnOutsidePointer])

  const list = Array.isArray(charts) ? charts : []
  const frames = list.map((chart, index) => {
    if (!chart?.anchor || !chart.id) return null
    let rect: ChartRect
    try { rect = anchorToRect(chart.anchor, props.geometry, zoom) } catch { return null }
    const selected = chart.id === selectedChartId
    if (viewport && !selected) {
      const margin = 64
      if (rect.left > viewport.left + viewport.width + margin || rect.top > viewport.top + viewport.height + margin ||
        rect.left + rect.width < viewport.left - margin || rect.top + rect.height < viewport.top - margin) return null
    }
    let cached = dataCache.current.get(chart)
    const version = props.dataVersion
    if (!cached || version === undefined || cached.version !== version) {
      let data: ResolvedChartData
      try { data = props.resolveData(chart) } catch { data = { categories: [], series: [] } }
      const signature = dataSignature(data)
      if (!cached || cached.signature !== signature) cached = { version, data, signature }
      else cached = { ...cached, version }
      dataCache.current.set(chart, cached)
    }
    return (
      <ChartFrame
        key={chart.id}
        chart={chart}
        data={cached.data}
        left={Math.round(rect.left * 10) / 10}
        top={Math.round(rect.top * 10) / 10}
        width={Math.round(rect.width * 10) / 10}
        height={Math.round(rect.height * 10) / 10}
        zoom={zoom}
        selected={selected}
        readOnly={readOnly}
        renderOptions={renderOptions}
        callbacks={callbacks}
        z={index}
      />
    )
  })

  if (!list.length) return null
  return (
    <div className={`chart-layer${className ? ` ${className}` : ''}`} style={style}>
      {frames}
    </div>
  )
}

export default ChartLayer
