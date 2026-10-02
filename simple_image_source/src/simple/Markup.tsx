// src/simple/Markup.tsx (WP8)
// Markup in Simple mode (design 4.6): text, arrows, rectangles, ellipses, lines and pasted pictures that stay
// editable until Done. An SVG overlay with viewBox = the image size draws them crisply at any zoom (lines
// and arrows use the exact outlines the baker fills, so preview and result match), with selection
// handles; drag moves or resizes, Shift constrains to 45 degrees and squares, Delete removes, double-click
// edits text in a positioned <textarea> (IME-safe). While Markup is open, Undo/Redo use the session stack.
// Baking (main.tsx) draws the items with src/shared/vector.ts, exactly like Advanced shape and text layers.
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, MouseEvent as ReactMouseEvent, PointerEvent as ReactPointerEvent } from 'react'
import { ArrowUpRight, Bold, Check, Circle, Minus, Square, Trash2, Type } from 'lucide-react'
import type { Point } from '../imaging/types.ts'
import { FONT_FAMILIES, lineOutlines } from '../shared/vector.ts'
import { isPanEvent } from './useViewportNavigation.ts'
import type { BoxHandle, MarkupCreateKind, MarkupHandle, MarkupHistory, MarkupItem, MarkupStyle, TextItem } from './markupModel.ts'
import {
  DEFAULT_MARKUP_STYLE,
  MARKUP_CREATE_KINDS,
  MAX_FONT_SIZE,
  MAX_STROKE_WIDTH,
  MIN_FONT_SIZE,
  MIN_STROKE_WIDTH,
  commitMarkup,
  createShapeItem,
  createTextItem,
  emptyMarkupHistory,
  handlesOf,
  hitHandle,
  hitTest,
  isLineKind,
  isMeaningfulShape,
  itemBounds,
  moveItem,
  nextMarkupId,
  normalizedBox,
  redoMarkup,
  removeItem,
  replaceItem,
  resizeItem,
  restyleItem,
  textBox,
  toShapeSpec,
  undoMarkup,
} from './markupModel.ts'

// ---------------------------------------------------------------------------------------------
// Session state
// ---------------------------------------------------------------------------------------------

export interface MarkupSessionState {
  readonly history: MarkupHistory
  /** Items while a drag or text edit is in progress (not yet a history step), or null. */
  readonly live: readonly MarkupItem[] | null
  readonly selectedId: string | null
  readonly editingId: string | null
  readonly createKind: MarkupCreateKind
  readonly style: MarkupStyle
}

export interface MarkupSession {
  readonly state: MarkupSessionState
  readonly items: readonly MarkupItem[]
  readonly selected: MarkupItem | null
  readonly canUndo: boolean
  readonly canRedo: boolean
  /** Items exist (Done/Save would bake something). */
  readonly hasItems: boolean
  setCreateKind(kind: MarkupCreateKind): void
  setStyle(patch: Partial<MarkupStyle>): void
  select(id: string | null): void
  startEditing(id: string): void
  finishEditing(): void
  setLive(items: readonly MarkupItem[] | null): void
  commit(items: readonly MarkupItem[], selectedId?: string | null): void
  addItem(item: MarkupItem, edit?: boolean): void
  deleteSelected(): boolean
  undo(): boolean
  redo(): boolean
  /** Ends the session (after baking or discarding); keeps the kind and style for next time. */
  reset(): void
  /** Latest state, for callbacks that run after awaits. */
  current(): MarkupSessionState
}

const RESTYLE_COALESCE_MS = 1200

function initialSession(style: MarkupStyle = DEFAULT_MARKUP_STYLE, createKind: MarkupCreateKind = 'arrow'): MarkupSessionState {
  return { history: emptyMarkupHistory(), live: null, selectedId: null, editingId: null, createKind, style }
}

export function useMarkupSession(): MarkupSession {
  const [state, setState] = useState<MarkupSessionState>(() => initialSession())
  const stateRef = useRef(state)
  const restyleRef = useRef<{ id: string; at: number } | null>(null)

  // Eager updates: event handlers read session.current() right after changing it (pointer moves are
  // batched by React), so the ref is the source of truth and React state follows it.
  const update = useCallback((next: (current: MarkupSessionState) => MarkupSessionState) => {
    const value = next(stateRef.current)
    if (value === stateRef.current) return
    stateRef.current = value
    setState(value)
  }, [])

  const items = state.live ?? state.history.present
  const selected = items.find((item) => item.id === state.selectedId) ?? null

  const commit = useCallback((next: readonly MarkupItem[], selectedId?: string | null) => {
    restyleRef.current = null
    update((current) => ({
      ...current,
      history: commitMarkup(current.history, next),
      live: null,
      selectedId: selectedId === undefined ? (next.some((item) => item.id === current.selectedId) ? current.selectedId : null) : selectedId,
    }))
  }, [update])

  const finishEditing = useCallback(() => {
    update((current) => {
      const id = current.editingId
      if (!id) return current
      const working = current.live ?? current.history.present
      const item = working.find((entry) => entry.id === id)
      const keep = item && item.kind === 'text' && item.text.trim().length > 0
      const next = keep ? working : removeItem(working, id)
      const unchanged = next === current.history.present || (next.length === current.history.present.length && next.every((entry, index) => entry === current.history.present[index]))
      return {
        ...current,
        history: unchanged ? current.history : commitMarkup(current.history, next),
        live: null,
        editingId: null,
        selectedId: keep ? id : null,
      }
    })
  }, [update])

  return useMemo<MarkupSession>(() => ({
    state,
    items,
    selected,
    canUndo: state.history.past.length > 0,
    canRedo: state.history.future.length > 0,
    hasItems: items.length > 0,
    setCreateKind: (kind) => update((current) => ({ ...current, createKind: kind })),
    setStyle: (patch) => {
      const current = stateRef.current
      const style = { ...current.style, ...patch }
      const target = (current.live ?? current.history.present).find((item) => item.id === current.selectedId)
      if (!target || current.editingId === target.id && current.live) {
        update((value) => ({ ...value, style, live: value.live && target ? replaceItem(value.live, restyleItem(target, patch)) : value.live }))
        return
      }
      const restyled = restyleItem(target, patch)
      const now = Date.now()
      const coalesce = restyleRef.current && restyleRef.current.id === target.id && now - restyleRef.current.at < RESTYLE_COALESCE_MS
      restyleRef.current = { id: target.id, at: now }
      update((value) => {
        const present = replaceItem(value.history.present, restyled)
        const history = coalesce ? { ...value.history, present, future: [] } : commitMarkup(value.history, present)
        return { ...value, style, history, live: null }
      })
    },
    select: (id) => update((current) => (current.selectedId === id ? current : { ...current, selectedId: id })),
    startEditing: (id) => update((current) => ({ ...current, editingId: id, selectedId: id, live: current.live ?? current.history.present })),
    finishEditing,
    setLive: (next) => update((current) => ({ ...current, live: next })),
    commit,
    addItem: (item, edit = false) => {
      restyleRef.current = null
      update((current) => {
        const next = [...current.history.present, item]
        if (edit) return { ...current, live: next, selectedId: item.id, editingId: item.id }
        return { ...current, history: commitMarkup(current.history, next), live: null, selectedId: item.id, editingId: null }
      })
    },
    deleteSelected: () => {
      const current = stateRef.current
      if (!current.selectedId || current.editingId) return false
      const next = removeItem(current.history.present, current.selectedId)
      if (next.length === current.history.present.length) return false
      commit(next, null)
      return true
    },
    undo: () => {
      const current = stateRef.current
      if (current.editingId || !current.history.past.length) return false
      restyleRef.current = null
      update((value) => ({ ...value, history: undoMarkup(value.history), live: null }))
      return true
    },
    redo: () => {
      const current = stateRef.current
      if (current.editingId || !current.history.future.length) return false
      restyleRef.current = null
      update((value) => ({ ...value, history: redoMarkup(value.history), live: null }))
      return true
    },
    reset: () => {
      restyleRef.current = null
      update((current) => initialSession(current.style, current.createKind))
    },
    current: () => stateRef.current,
  }), [commit, finishEditing, items, selected, state, update])
}

// ---------------------------------------------------------------------------------------------
// Overlay
// ---------------------------------------------------------------------------------------------

type Drag =
  | { readonly mode: 'create'; readonly pointerId: number; readonly id: string; readonly kind: Exclude<MarkupCreateKind, 'text'>; readonly start: Point }
  | { readonly mode: 'text'; readonly pointerId: number; readonly start: Point }
  | { readonly mode: 'move'; readonly pointerId: number; readonly original: MarkupItem; readonly start: Point; moved: boolean }
  | { readonly mode: 'resize'; readonly pointerId: number; readonly original: MarkupItem; readonly handle: MarkupHandle }

const HANDLE_CSS_PX = 9
const HIT_CSS_PX = 6

function outlinePath(outlines: readonly (readonly Point[])[]): string {
  return outlines.map((outline) => `M${outline.map((p) => `${p.x.toFixed(2)} ${p.y.toFixed(2)}`).join('L')}Z`).join('')
}

const HANDLE_CURSORS: Readonly<Record<BoxHandle, string>> = {
  nw: 'nwse-resize', se: 'nwse-resize', ne: 'nesw-resize', sw: 'nesw-resize', n: 'ns-resize', s: 'ns-resize', e: 'ew-resize', w: 'ew-resize',
}

function MarkupGraphic({ item }: { readonly item: MarkupItem }) {
  if (item.kind === 'text') {
    const box = textBox(item)
    return (
      <text fontFamily={`"${item.fontFamily}", sans-serif`} fontSize={item.fontSize} fontWeight={item.bold ? 700 : 400} fill={item.color} style={{ whiteSpace: 'pre' }}>
        {box.lines.map((line, index) => <tspan key={index} x={line.x} y={line.baseline}>{line.text}</tspan>)}
      </text>
    )
  }
  if (item.kind === 'image') {
    return <image href={item.src} x={item.x} y={item.y} width={item.width} height={item.height} preserveAspectRatio="none" />
  }
  if (isLineKind(item.kind)) return <path d={outlinePath(lineOutlines(toShapeSpec(item)))} fill={item.color} fillRule="nonzero" />
  const box = normalizedBox(item)
  const common = { fill: item.fill ?? 'none', stroke: item.color, strokeWidth: item.width }
  if (item.kind === 'ellipse') return <ellipse cx={box.x + box.width / 2} cy={box.y + box.height / 2} rx={box.width / 2} ry={box.height / 2} {...common} />
  return <rect x={box.x} y={box.y} width={box.width} height={box.height} strokeLinejoin="miter" {...common} />
}

export interface MarkupLayerProps {
  /** Image size (the overlay's viewBox). */
  readonly width: number
  readonly height: number
  /** CSS px per image px. */
  readonly zoom: number
  readonly session: MarkupSession
  readonly disabled: boolean
}

export function MarkupLayer(props: MarkupLayerProps) {
  const { width, height, zoom, session, disabled } = props
  const { items, selected, state } = session
  const dragRef = useRef<Drag | null>(null)
  const scale = Math.max(1e-6, zoom)

  const toImage = (event: ReactPointerEvent<SVGSVGElement> | PointerEvent, element: Element): Point => {
    const rect = element.getBoundingClientRect()
    return {
      x: ((event.clientX - rect.left) * width) / Math.max(1, rect.width),
      y: ((event.clientY - rect.top) * height) / Math.max(1, rect.height),
    }
  }

  const onPointerDown = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (disabled || event.button !== 0 || !event.isPrimary || isPanEvent(event.nativeEvent)) return
    const svg = event.currentTarget
    const point = toImage(event, svg)
    const tolerance = HIT_CSS_PX / scale
    const current = session.current()
    if (current.editingId) {
      // A click outside the text being edited only ends the edit.
      event.preventDefault()
      session.finishEditing()
      return
    }
    svg.setPointerCapture(event.pointerId)
    if (selected) {
      const handle = hitHandle(selected, point, (HANDLE_CSS_PX * 0.75 + 2) / scale)
      if (handle) {
        dragRef.current = { mode: 'resize', pointerId: event.pointerId, original: selected, handle }
        return
      }
    }
    const hit = hitTest(items, point, tolerance)
    if (hit) {
      session.select(hit.id)
      dragRef.current = { mode: 'move', pointerId: event.pointerId, original: hit, start: point, moved: false }
      return
    }
    session.select(null)
    if (current.createKind === 'text') {
      dragRef.current = { mode: 'text', pointerId: event.pointerId, start: point }
      return
    }
    dragRef.current = { mode: 'create', pointerId: event.pointerId, id: nextMarkupId(), kind: current.createKind, start: point }
  }

  const onPointerMove = (event: ReactPointerEvent<SVGSVGElement>) => {
    const drag = dragRef.current
    if (!drag || drag.pointerId !== event.pointerId) return
    const point = toImage(event, event.currentTarget)
    const present = session.current().history.present
    if (drag.mode === 'create') {
      const shape = createShapeItem(drag.id, drag.kind, drag.start, point, session.current().style, event.shiftKey)
      session.setLive([...present, shape])
    } else if (drag.mode === 'move') {
      const dx = point.x - drag.start.x
      const dy = point.y - drag.start.y
      if (!drag.moved && Math.hypot(dx, dy) * scale < 2) return
      drag.moved = true
      session.setLive(replaceItem(present, moveItem(drag.original, dx, dy)))
    } else if (drag.mode === 'resize') {
      session.setLive(replaceItem(present, resizeItem(drag.original, drag.handle, point, event.shiftKey)))
    }
  }

  const onPointerUp = (event: ReactPointerEvent<SVGSVGElement>) => {
    const drag = dragRef.current
    if (!drag || drag.pointerId !== event.pointerId) return
    dragRef.current = null
    const current = session.current()
    if (drag.mode === 'text') {
      const item = createTextItem(nextMarkupId(), drag.start, current.style)
      session.addItem(item, true)
      return
    }
    if (drag.mode === 'create') {
      const shape = current.live?.find((item) => item.id === drag.id)
      if (shape && shape.kind !== 'text' && shape.kind !== 'image' && isMeaningfulShape(shape)) session.commit(current.live ?? [], shape.id)
      else session.setLive(null)
      return
    }
    if (drag.mode === 'move' && !drag.moved) {
      session.setLive(null)
      return
    }
    if (current.live) session.commit(current.live, drag.original.id)
  }

  const onPointerCancel = () => {
    dragRef.current = null
    session.setLive(null)
  }

  const onDoubleClick = (event: ReactMouseEvent<SVGSVGElement>) => {
    if (disabled) return
    const rect = event.currentTarget.getBoundingClientRect()
    const point = { x: ((event.clientX - rect.left) * width) / Math.max(1, rect.width), y: ((event.clientY - rect.top) * height) / Math.max(1, rect.height) }
    const hit = hitTest(items, point, HIT_CSS_PX / scale)
    if (hit?.kind === 'text') session.startEditing(hit.id)
  }

  const editing = state.editingId ? items.find((item) => item.id === state.editingId) ?? null : null
  const handleSize = HANDLE_CSS_PX / scale
  const selectionBox = selected && !editing ? (selected.kind === 'text' ? textBox(selected).rect : itemBounds(selected)) : null

  return (
    <>
      <svg
        className={`markup-layer kind-${state.createKind}`}
        viewBox={`0 0 ${width} ${height}`}
        preserveAspectRatio="none"
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerCancel}
        onDoubleClick={onDoubleClick}
        aria-label="Markup"
      >
        {items.map((item) => (item.id === state.editingId ? null : <MarkupGraphic key={item.id} item={item} />))}
        {selectionBox && (
          <rect className="markup-selection" x={selectionBox.x} y={selectionBox.y} width={selectionBox.width} height={selectionBox.height} vectorEffect="non-scaling-stroke" />
        )}
        {selected && !editing && handlesOf(selected).map(({ handle, point }) => (
          <rect
            key={handle}
            className="markup-handle"
            x={point.x - handleSize / 2}
            y={point.y - handleSize / 2}
            width={handleSize}
            height={handleSize}
            vectorEffect="non-scaling-stroke"
            style={{ cursor: handle === 'start' || handle === 'end' ? 'move' : HANDLE_CURSORS[handle as BoxHandle] }}
          />
        ))}
      </svg>
      {editing && editing.kind === 'text' && <TextEditor item={editing} zoom={scale} session={session} />}
    </>
  )
}

function TextEditor({ item, zoom, session }: { readonly item: TextItem; readonly zoom: number; readonly session: MarkupSession }) {
  const ref = useRef<HTMLTextAreaElement>(null)
  const box = textBox({ ...item, text: item.text || ' ' })
  useEffect(() => {
    const element = ref.current
    if (!element) return
    element.focus()
    element.setSelectionRange(element.value.length, element.value.length)
  }, [item.id])
  const style: CSSProperties = {
    left: item.x * zoom,
    top: item.y * zoom,
    width: Math.max(box.rect.width + item.fontSize * 0.6, item.fontSize * 2) * zoom + 6,
    height: box.rect.height * zoom + 4,
    color: item.color,
    font: `${item.bold ? 700 : 400} ${item.fontSize * zoom}px "${item.fontFamily}", sans-serif`,
    lineHeight: 1.2,
  }
  const onKeyDown = (event: ReactKeyboardEvent<HTMLTextAreaElement>) => {
    if (event.nativeEvent.isComposing) return
    if (event.key === 'Escape' || (event.key === 'Enter' && (event.ctrlKey || event.metaKey))) {
      event.preventDefault()
      event.stopPropagation()
      session.finishEditing()
    }
  }
  return (
    <textarea
      ref={ref}
      className="markup-text-editor"
      style={style}
      value={item.text}
      spellCheck={false}
      aria-label="Markup text"
      placeholder="Type here"
      onChange={(event) => {
        const working = session.current().live ?? session.current().history.present
        session.setLive(replaceItem(working, { ...item, text: event.target.value }))
      }}
      onKeyDown={onKeyDown}
      onBlur={() => session.finishEditing()}
    />
  )
}

// ---------------------------------------------------------------------------------------------
// Options strip
// ---------------------------------------------------------------------------------------------

const KIND_META: Readonly<Record<MarkupCreateKind, { readonly label: string; readonly icon: typeof Type }>> = {
  text: { label: 'Text', icon: Type },
  arrow: { label: 'Arrow', icon: ArrowUpRight },
  rectangle: { label: 'Rectangle', icon: Square },
  ellipse: { label: 'Ellipse', icon: Circle },
  line: { label: 'Line', icon: Minus },
}

export interface MarkupOptionsProps {
  readonly session: MarkupSession
  readonly busy: boolean
  readonly onDone: () => void
}

export function MarkupOptions({ session, busy, onDone }: MarkupOptionsProps) {
  const { state, selected } = session
  const style = state.style
  const kind = selected?.kind ?? state.createKind
  const showsText = kind === 'text'
  const showsStroke = kind !== 'text' && kind !== 'image'
  const showsFill = kind === 'rectangle' || kind === 'ellipse'
  const color = selected && selected.kind !== 'image' ? selected.color : style.color
  const width = selected && selected.kind !== 'image' && selected.kind !== 'text' ? selected.width : style.width
  const fill = selected && (selected.kind === 'rectangle' || selected.kind === 'ellipse') ? selected.fill : style.fill
  const fontSize = selected?.kind === 'text' ? selected.fontSize : style.fontSize
  const fontFamily = selected?.kind === 'text' ? selected.fontFamily : style.fontFamily
  const bold = selected?.kind === 'text' ? selected.bold : style.bold
  const families = FONT_FAMILIES.includes(fontFamily) ? FONT_FAMILIES : [fontFamily, ...FONT_FAMILIES]
  return (
    <div className="context-strip markup-strip" role="toolbar" aria-label="Markup">
      <div className="strip-group" role="group" aria-label="Add">
        {MARKUP_CREATE_KINDS.map((entry) => {
          const Icon = KIND_META[entry].icon
          return (
            <button
              key={entry}
              type="button"
              className={state.createKind === entry ? 'active' : ''}
              aria-pressed={state.createKind === entry}
              title={`${KIND_META[entry].label}: drag on the image${entry === 'text' ? ' (or click)' : ''}`}
              data-markup-kind={entry}
              disabled={busy}
              onClick={() => { session.setCreateKind(entry); session.select(null) }}
            ><Icon /><span>{KIND_META[entry].label}</span></button>
          )
        })}
      </div>
      <div className="strip-rule" />
      <div className="strip-group markup-style">
        {kind !== 'image' && (
          <label className="color-control" title="Colour"><input type="color" value={color} disabled={busy} onChange={(event) => session.setStyle({ color: event.target.value })} aria-label="Markup colour" /><span style={{ background: color }} /></label>
        )}
        {showsStroke && (
          <label className="size-control markup-width" title="Line width"><span>Width</span><input type="range" min={MIN_STROKE_WIDTH} max={MAX_STROKE_WIDTH} value={width} disabled={busy} onChange={(event) => session.setStyle({ width: Number(event.target.value) })} /><output>{width}px</output></label>
        )}
        {showsFill && (
          <>
            <label className="markup-check"><input type="checkbox" checked={Boolean(fill)} disabled={busy} onChange={(event) => session.setStyle({ fill: event.target.checked ? (fill ?? '#ffffff') : null })} /><span>Fill</span></label>
            {fill && <label className="color-control" title="Fill colour"><input type="color" value={fill} disabled={busy} onChange={(event) => session.setStyle({ fill: event.target.value })} aria-label="Fill colour" /><span style={{ background: fill }} /></label>}
          </>
        )}
        {showsText && (
          <>
            <select className="markup-font" value={fontFamily} disabled={busy} aria-label="Font" onChange={(event) => session.setStyle({ fontFamily: event.target.value })}>
              {families.map((family) => <option key={family} value={family}>{family}</option>)}
            </select>
            <label className="markup-number" title="Font size (px)"><input type="number" min={MIN_FONT_SIZE} max={MAX_FONT_SIZE} value={fontSize} disabled={busy} onChange={(event) => { const value = Number(event.target.value); if (Number.isFinite(value) && value > 0) session.setStyle({ fontSize: value }) }} aria-label="Font size" /><span>px</span></label>
            <button type="button" className={`icon-only ${bold ? 'active' : ''}`} aria-pressed={bold} title="Bold" aria-label="Bold" disabled={busy} onClick={() => session.setStyle({ bold: !bold })}><Bold /></button>
          </>
        )}
      </div>
      <div className="strip-spacer" />
      <div className="strip-group">
        <button type="button" className="icon-only" title="Delete the selected item (Delete)" aria-label="Delete the selected item" disabled={busy || !selected} onClick={() => session.deleteSelected()}><Trash2 /></button>
        <button type="button" className="strip-primary" disabled={busy} onClick={onDone} title="Draw the markup into the image"><Check /><span>Done</span></button>
      </div>
    </div>
  )
}
