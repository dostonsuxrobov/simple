import { memo, useEffect, useMemo, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type PointerEvent as ReactPointerEvent } from 'react'
import type { SheetImage } from '../spreadsheet-types'
import { anchorToRect, rectToAnchor, snapRectToCells, type ChartGeometry, type ChartRect } from '../lib/charts'
import './charts.css'

type Handle = 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw'
const HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
const MIN_SIZE = 12

export interface ImageLayerProps {
  images: SheetImage[] | undefined
  geometry: ChartGeometry
  zoom: number
  selectedImageId: string | null
  viewport?: ChartRect
  onSelect: (id: string | null) => void
  onChange: (image: SheetImage) => void
  onDelete: (id: string) => void
  onContextMenu?: (id: string, position: { clientX: number; clientY: number }) => void
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

interface FrameProps {
  image: SheetImage
  rect: ChartRect
  zoom: number
  selected: boolean
  z: number
  onSelect: (id: string) => void
  onCommit: (image: SheetImage, rect: ChartRect, snap: boolean) => void
  onRemove: (id: string) => void
  onDeselect: () => void
  onContextMenu: (id: string, clientX: number, clientY: number) => void
}

const ImageFrame = memo(function ImageFrame({ image, rect: base, zoom, selected, z, onSelect, onCommit, onRemove, onDeselect, onContextMenu }: FrameProps) {
  const frameRef = useRef<HTMLDivElement | null>(null)
  const interaction = useRef<Interaction | null>(null)
  const [draft, setDraft] = useState<ChartRect | null>(null)
  const rect = draft || base

  useEffect(() => {
    if (selected && document.activeElement !== frameRef.current) frameRef.current?.focus({ preventScroll: true })
  }, [selected])

  const begin = (event: ReactPointerEvent<HTMLElement>, mode: Interaction['mode']) => {
    if (event.button !== 0) return
    event.stopPropagation()
    event.preventDefault()
    if (!selected) onSelect(image.id)
    try { frameRef.current?.setPointerCapture(event.pointerId) } catch { /* released */ }
    interaction.current = { pointerId: event.pointerId, mode, startX: event.clientX, startY: event.clientY, origin: base, active: mode !== 'move' }
  }
  const onPointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const current = interaction.current
    if (!current || current.pointerId !== event.pointerId) return
    const dx = event.clientX - current.startX
    const dy = event.clientY - current.startY
    if (!current.active && Math.hypot(dx, dy) < 3) return
    current.active = true
    if (current.mode === 'move') setDraft({ ...current.origin, left: Math.max(0, current.origin.left + dx), top: Math.max(0, current.origin.top + dy) })
    // Pictures keep their proportions from a corner unless Shift is held (Excel's lock aspect ratio).
    else setDraft(resizeRect(current.origin, current.mode, dx, dy, !event.shiftKey, MIN_SIZE * zoom))
  }
  const finish = (event: ReactPointerEvent<HTMLDivElement>, cancel = false) => {
    const current = interaction.current
    if (!current || current.pointerId !== event.pointerId) return
    interaction.current = null
    try { if (frameRef.current?.hasPointerCapture(event.pointerId)) frameRef.current.releasePointerCapture(event.pointerId) } catch { /* released */ }
    const next = draft
    setDraft(null)
    if (!cancel && current.active && next) onCommit(image, next, event.altKey)
  }
  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.target !== frameRef.current) return
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault(); event.stopPropagation()
      onRemove(image.id)
    } else if (event.key === 'Escape') {
      event.preventDefault(); event.stopPropagation()
      if (interaction.current) { interaction.current = null; setDraft(null) } else onDeselect()
    } else if (event.key.startsWith('Arrow')) {
      event.preventDefault(); event.stopPropagation()
      const step = (event.shiftKey ? 10 : 1) * zoom
      const dx = event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0
      const dy = event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0
      if (event.ctrlKey || event.metaKey) onCommit(image, { ...base, width: Math.max(MIN_SIZE * zoom, base.width + dx), height: Math.max(MIN_SIZE * zoom, base.height + dy) }, false)
      else onCommit(image, { ...base, left: Math.max(0, base.left + dx), top: Math.max(0, base.top + dy) }, false)
    }
  }
  const renderable = /^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml|avif)/i.test(image.src)
  return (
    <div
      ref={frameRef}
      className={`chart-frame image-frame${selected ? ' is-selected' : ''}${draft ? ' is-dragging' : ''}`}
      style={{ left: rect.left, top: rect.top, width: rect.width, height: rect.height, zIndex: selected ? 1000 + z : z + 1 }}
      data-image-id={image.id}
      role="img"
      aria-label={image.altText || image.name || 'Picture'}
      tabIndex={selected ? 0 : -1}
      onPointerDown={(event) => begin(event, 'move')}
      onPointerMove={onPointerMove}
      onPointerUp={(event) => finish(event)}
      onPointerCancel={(event) => finish(event, true)}
      onLostPointerCapture={(event) => { if (interaction.current?.pointerId === event.pointerId) finish(event) }}
      onContextMenu={(event) => { event.preventDefault(); event.stopPropagation(); if (!selected) onSelect(image.id); onContextMenu(image.id, event.clientX, event.clientY) }}
      onKeyDown={onKeyDown}
    >
      {renderable
        ? <img className="image-frame-img" src={image.src} alt={image.altText || ''} draggable={false} />
        : <div className="image-frame-placeholder">{image.name || 'Picture'}<small>Preview not available</small></div>}
      {selected && HANDLES.map((handle) => (
        <div key={handle} className={`chart-handle chart-handle-${handle}`} data-handle={handle} onPointerDown={(event) => begin(event, handle)} aria-hidden="true" />
      ))}
    </div>
  )
})

/** Floating pictures over the grid (drawn in canvas coordinates, like charts). */
export function ImageLayer({ images, geometry, zoom, selectedImageId, viewport, onSelect, onChange, onDelete, onContextMenu }: ImageLayerProps) {
  const latest = useRef({ geometry, zoom, onChange, onSelect, onDelete, onContextMenu })
  latest.current = { geometry, zoom, onChange, onSelect, onDelete, onContextMenu }
  const callbacks = useMemo(() => ({
    select: (id: string) => latest.current.onSelect(id),
    deselect: () => latest.current.onSelect(null),
    remove: (id: string) => latest.current.onDelete(id),
    contextMenu: (id: string, clientX: number, clientY: number) => latest.current.onContextMenu?.(id, { clientX, clientY }),
    commit: (image: SheetImage, rect: ChartRect, snap: boolean) => {
      const { geometry: currentGeometry, zoom: currentZoom } = latest.current
      const target = snap ? snapRectToCells(rect, currentGeometry) : rect
      const anchor = rectToAnchor(target, currentGeometry, currentZoom, image.anchor?.editAs)
      if (JSON.stringify(anchor) === JSON.stringify(image.anchor)) return
      latest.current.onChange({ ...image, anchor })
    },
  }), [])

  useEffect(() => {
    if (!selectedImageId) return
    const listener = (event: PointerEvent) => {
      const target = event.target as Element | null
      if (target instanceof Element && target.closest('.image-frame, [data-image-keep-selection]')) return
      latest.current.onSelect(null)
    }
    document.addEventListener('pointerdown', listener, true)
    return () => document.removeEventListener('pointerdown', listener, true)
  }, [selectedImageId])

  const list = Array.isArray(images) ? images : []
  if (!list.length) return null
  return (
    <div className="chart-layer image-layer">
      {list.map((image, index) => {
        if (!image?.anchor || !image.id) return null
        let rect: ChartRect
        try { rect = anchorToRect(image.anchor, geometry, zoom) } catch { return null }
        const selected = image.id === selectedImageId
        if (viewport && !selected) {
          const margin = 64
          if (rect.left > viewport.left + viewport.width + margin || rect.top > viewport.top + viewport.height + margin ||
            rect.left + rect.width < viewport.left - margin || rect.top + rect.height < viewport.top - margin) return null
        }
        return (
          <ImageFrame
            key={image.id}
            image={image}
            rect={{ left: Math.round(rect.left * 10) / 10, top: Math.round(rect.top * 10) / 10, width: Math.round(rect.width * 10) / 10, height: Math.round(rect.height * 10) / 10 }}
            zoom={zoom}
            selected={selected}
            z={index}
            onSelect={callbacks.select}
            onCommit={callbacks.commit}
            onRemove={callbacks.remove}
            onDeselect={callbacks.deselect}
            onContextMenu={callbacks.contextMenu}
          />
        )
      })}
    </div>
  )
}
