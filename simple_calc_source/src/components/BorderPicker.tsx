import { useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import { ChevronDown, PencilLine } from 'lucide-react'
import type { CellBorderSide, SpreadsheetColor } from '../spreadsheet-types'
import { DEFAULT_BORDER_SIDE, DEFAULT_THEME_COLORS, resolveColorHex } from '../lib/cell-styles'
import type { BorderPreset } from '../lib/cell-styles'
import { ColorPicker, FormatPopover, moveFocusByArrow } from './ColorPicker'
import type { AnchorRect } from './ColorPicker'
import './format-ui.css'

export { applyBorderPreset, borderPresetChange } from '../lib/cell-styles'
export type { BorderPreset, CellPosition, FormatBorderChange } from '../lib/cell-styles'

// ---------------------------------------------------------------------------------------
// Line styles
// ---------------------------------------------------------------------------------------

export interface LineStyleSpec {
  label: string
  width: number
  dash?: string
  double?: boolean
}

export const LINE_STYLE_SPECS: Record<string, LineStyleSpec> = {
  hair: { label: 'Hair', width: 1, dash: '1 1' },
  dotted: { label: 'Dotted', width: 1, dash: '2 2' },
  dashDotDot: { label: 'Dash dot dot', width: 1, dash: '6 2 2 2 2 2' },
  dashDot: { label: 'Dash dot', width: 1, dash: '6 2 2 2' },
  dashed: { label: 'Dashed', width: 1, dash: '4 2' },
  thin: { label: 'Thin', width: 1 },
  mediumDashDotDot: { label: 'Medium dash dot dot', width: 2, dash: '8 3 2 3 2 3' },
  slantDashDot: { label: 'Slanted dash dot', width: 2, dash: '10 2 4 2' },
  mediumDashDot: { label: 'Medium dash dot', width: 2, dash: '8 3 3 3' },
  mediumDashed: { label: 'Medium dashed', width: 2, dash: '8 3' },
  medium: { label: 'Medium', width: 2 },
  thick: { label: 'Thick', width: 3 },
  double: { label: 'Double', width: 1, double: true },
}

/** Excel's Format Cells line-style gallery order: two columns, top to bottom. */
export const LINE_STYLE_GALLERY: readonly (string | null)[] = [
  null, 'hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin',
  'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double',
]

/** Compact list for the toolbar dropdown (most common first). */
const LINE_STYLE_MENU = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'dashDot', 'dashDotDot', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot']

/** Draws one border side as SVG line(s) between two points. Use inside an <svg>. */
export function BorderLine({ side, x1, y1, x2, y2, themeColors = DEFAULT_THEME_COLORS, fallbackColor = '#000000' }: {
  side: CellBorderSide
  x1: number; y1: number; x2: number; y2: number
  themeColors?: readonly string[]
  fallbackColor?: string
}) {
  const spec = LINE_STYLE_SPECS[String(side.style)] || LINE_STYLE_SPECS.thin
  const color = resolveColorHex(side.color, themeColors) || fallbackColor
  if (spec.double) {
    const dx = x2 - x1
    const dy = y2 - y1
    const length = Math.hypot(dx, dy) || 1
    const ox = (-dy / length) * 1.5
    const oy = (dx / length) * 1.5
    return (
      <g stroke={color} strokeWidth={1} shapeRendering="crispEdges">
        <line x1={x1 + ox} y1={y1 + oy} x2={x2 + ox} y2={y2 + oy} />
        <line x1={x1 - ox} y1={y1 - oy} x2={x2 - ox} y2={y2 - oy} />
      </g>
    )
  }
  return <line x1={x1} y1={y1} x2={x2} y2={y2} stroke={color} strokeWidth={spec.width} strokeDasharray={spec.dash} shapeRendering={x1 === x2 || y1 === y2 ? 'crispEdges' : undefined} />
}

/** A horizontal sample of a line style (galleries, dropdown buttons). */
export function LineStyleSample({ style, color = '#000000', width = 56 }: { style: string | null; color?: string; width?: number }) {
  if (!style) return <span className="fmt-line-none">None</span>
  return (
    <svg className="fmt-line-sample" width={width} height={9} viewBox={`0 0 ${width} 9`} aria-hidden="true">
      <BorderLine side={{ style, color }} x1={2} y1={4.5} x2={width - 2} y2={4.5} />
    </svg>
  )
}

// ---------------------------------------------------------------------------------------
// Preset icons
// ---------------------------------------------------------------------------------------

type Segment = 'top' | 'bottom' | 'left' | 'right' | 'h' | 'v'

export const BORDER_PRESET_ITEMS: readonly { id: BorderPreset; label: string; segments: Segment[]; needs?: 'row' | 'column' | 'any' }[] = [
  { id: 'all', label: 'All borders', segments: ['top', 'bottom', 'left', 'right', 'h', 'v'] },
  { id: 'inner', label: 'Inner borders', segments: ['h', 'v'], needs: 'any' },
  { id: 'horizontal', label: 'Horizontal borders', segments: ['h'], needs: 'row' },
  { id: 'vertical', label: 'Vertical borders', segments: ['v'], needs: 'column' },
  { id: 'outer', label: 'Outer borders', segments: ['top', 'bottom', 'left', 'right'] },
  { id: 'left', label: 'Left border', segments: ['left'] },
  { id: 'top', label: 'Top border', segments: ['top'] },
  { id: 'right', label: 'Right border', segments: ['right'] },
  { id: 'bottom', label: 'Bottom border', segments: ['bottom'] },
  { id: 'clear', label: 'Clear borders', segments: [] },
]

const SEGMENT_LINES: Record<Segment, [number, number, number, number]> = {
  top: [2.5, 2.5, 15.5, 2.5],
  bottom: [2.5, 15.5, 15.5, 15.5],
  left: [2.5, 2.5, 2.5, 15.5],
  right: [15.5, 2.5, 15.5, 15.5],
  h: [2.5, 9, 15.5, 9],
  v: [9, 2.5, 9, 15.5],
}

export function BorderPresetIcon({ segments, size = 18 }: { segments: readonly Segment[]; size?: number }) {
  const all: Segment[] = ['top', 'bottom', 'left', 'right', 'h', 'v']
  return (
    <svg width={size} height={size} viewBox="0 0 18 18" aria-hidden="true" className="fmt-border-icon">
      {all.filter((segment) => !segments.includes(segment)).map((segment) => {
        const [x1, y1, x2, y2] = SEGMENT_LINES[segment]
        return <line key={segment} x1={x1} y1={y1} x2={x2} y2={y2} stroke="currentColor" strokeOpacity={0.38} strokeWidth={1} strokeDasharray="1 1.6" />
      })}
      {segments.map((segment) => {
        const [x1, y1, x2, y2] = SEGMENT_LINES[segment]
        return <line key={segment} x1={x1} y1={y1} x2={x2} y2={y2} stroke="currentColor" strokeWidth={1.7} strokeLinecap="square" />
      })}
    </svg>
  )
}

// ---------------------------------------------------------------------------------------
// Border picker (Google Sheets-style dropdown)
// ---------------------------------------------------------------------------------------

export interface BorderPickerProps {
  /** Popover anchor; omit (or null) to render inline. */
  anchor?: AnchorRect | null
  ignoreElement?: HTMLElement | null
  /** Popover only: Escape, outside press, and after a preset is applied. */
  onClose?: () => void
  themeColors?: readonly string[]
  /** Current line style + colour (controlled). Uncontrolled pickers remember the last choice. */
  side?: CellBorderSide
  onSideChange?: (side: CellBorderSide) => void
  /** Apply a preset to the selection — see `applyBorderPreset(preset, side, position, style)`. */
  onApply: (preset: BorderPreset, side: CellBorderSide) => void
  /** Enables inner / horizontal / vertical presets. Default true. */
  multiRow?: boolean
  multiColumn?: boolean
  /** Shows "More borders…" (open Format Cells → Border). */
  onMoreBorders?: () => void
}

let rememberedSide: CellBorderSide = DEFAULT_BORDER_SIDE

export function BorderPicker({ anchor, ignoreElement, onClose, themeColors = DEFAULT_THEME_COLORS, side, onSideChange, onApply, multiRow = true, multiColumn = true, onMoreBorders }: BorderPickerProps) {
  const [localSide, setLocalSide] = useState<CellBorderSide>(rememberedSide)
  const current = side ?? localSide
  const [menu, setMenu] = useState<null | { kind: 'color' | 'style'; anchor: AnchorRect }>(null)
  const gridRef = useRef<HTMLDivElement>(null)
  const colorButtonRef = useRef<HTMLButtonElement>(null)
  const styleButtonRef = useRef<HTMLButtonElement>(null)
  const colorHex = resolveColorHex(current.color, themeColors) || '#000000'

  const updateSide = (next: CellBorderSide) => {
    rememberedSide = next
    setLocalSide(next)
    onSideChange?.(next)
  }

  const apply = (preset: BorderPreset) => {
    onApply(preset, current)
    if (anchor) onClose?.()
  }

  const openMenu = (kind: 'color' | 'style', element: HTMLButtonElement | null) => {
    if (menu?.kind === kind) { setMenu(null); return }
    const rect = element?.getBoundingClientRect()
    if (rect) setMenu({ kind, anchor: { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom } })
  }

  const handleGridKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const items = Array.from(gridRef.current?.querySelectorAll<HTMLElement>('button:not([disabled])') || [])
    const moved = moveFocusByArrow(event, items)
    if (moved) items.forEach((item) => item.setAttribute('tabindex', item === moved ? '0' : '-1'))
  }

  const available = (needs?: 'row' | 'column' | 'any') => (
    needs === 'row' ? multiRow : needs === 'column' ? multiColumn : needs === 'any' ? multiRow || multiColumn : true
  )

  const body = (
    <div className="fmt-border-picker">
      <div className="fmt-border-main">
        <div ref={gridRef} className="fmt-border-grid" role="group" aria-label="Border presets" onKeyDown={handleGridKey}>
          {BORDER_PRESET_ITEMS.map((item, index) => (
            <button
              key={item.id}
              type="button"
              className="fmt-border-preset"
              tabIndex={index === 0 ? 0 : -1}
              disabled={!available(item.needs)}
              title={item.label}
              aria-label={item.label}
              onClick={() => apply(item.id)}
            >
              <BorderPresetIcon segments={item.segments} />
            </button>
          ))}
        </div>
        <div className="fmt-border-side-tools">
          <button
            ref={colorButtonRef}
            type="button"
            className="fmt-border-tool"
            aria-haspopup="dialog"
            aria-expanded={menu?.kind === 'color'}
            aria-label="Border color"
            title="Border color"
            onClick={() => openMenu('color', colorButtonRef.current)}
          >
            <span className="fmt-border-tool-icon">
              <PencilLine size={14} aria-hidden="true" />
              <span className="fmt-border-tool-bar" style={{ background: colorHex }} />
            </span>
            <ChevronDown size={11} aria-hidden="true" />
          </button>
          <button
            ref={styleButtonRef}
            type="button"
            className="fmt-border-tool"
            aria-haspopup="listbox"
            aria-expanded={menu?.kind === 'style'}
            aria-label={`Border style: ${LINE_STYLE_SPECS[String(current.style)]?.label || 'Thin'}`}
            title="Border style"
            onClick={() => openMenu('style', styleButtonRef.current)}
          >
            <LineStyleSample style={String(current.style || 'thin')} width={24} />
            <ChevronDown size={11} aria-hidden="true" />
          </button>
        </div>
      </div>
      {onMoreBorders && (
        <button type="button" className="fmt-border-more" onClick={() => { onMoreBorders(); onClose?.() }}>More borders…</button>
      )}
      {menu?.kind === 'color' && (
        <ColorPicker
          anchor={menu.anchor}
          ignoreElement={colorButtonRef.current}
          value={current.color ?? null}
          themeColors={themeColors}
          mode="text"
          label="Border color"
          onChange={(color) => updateSide({ ...current, color: color ?? { argb: 'FF000000' } })}
          onClose={() => setMenu(null)}
        />
      )}
      {menu?.kind === 'style' && (
        <FormatPopover anchor={menu.anchor} ignoreElement={styleButtonRef.current} label="Border style" onClose={() => setMenu(null)} initialFocus='[aria-checked="true"]'>
          <LineStyleMenu
            value={String(current.style || 'thin')}
            color={colorHex}
            onSelect={(style) => { updateSide({ ...current, style }); setMenu(null) }}
          />
        </FormatPopover>
      )}
    </div>
  )

  if (!anchor) return body
  return (
    <FormatPopover anchor={anchor} ignoreElement={ignoreElement} onClose={() => onClose?.()} label="Borders">
      {body}
    </FormatPopover>
  )
}

function LineStyleMenu({ value, color, onSelect }: { value: string; color: string; onSelect: (style: string) => void }) {
  const listRef = useRef<HTMLDivElement>(null)
  const handleKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const items = Array.from(listRef.current?.querySelectorAll<HTMLElement>('[role="menuitemradio"]') || [])
    moveFocusByArrow(event, items)
  }
  return (
    <div ref={listRef} className="fmt-line-menu" role="menu" aria-label="Line style" onKeyDown={handleKey}>
      {LINE_STYLE_MENU.map((style) => (
        <button
          key={style}
          type="button"
          role="menuitemradio"
          aria-checked={style === value}
          className={`fmt-line-menu-item${style === value ? ' is-selected' : ''}`}
          title={LINE_STYLE_SPECS[style].label}
          onClick={() => onSelect(style)}
        >
          <LineStyleSample style={style} color={color} width={96} />
        </button>
      ))}
    </div>
  )
}

/** Resolves a picker colour choice to a border colour (Automatic draws black). */
export function borderColorOrAutomatic(color: SpreadsheetColor | null): SpreadsheetColor {
  return color ?? { argb: 'FF000000' }
}
