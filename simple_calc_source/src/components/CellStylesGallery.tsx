import { useRef } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent } from 'react'
import { CELL_STYLES, CELL_STYLE_GROUPS, DEFAULT_THEME_COLORS, stylePreviewCss } from '../lib/cell-styles'
import type { CellStylePreset } from '../lib/cell-styles'
import { FormatPopover, moveFocusByArrow } from './ColorPicker'
import type { AnchorRect } from './ColorPicker'
import './format-ui.css'

export { CELL_STYLES, applyCellStylePreset, applyCellStylePresetToCell } from '../lib/cell-styles'
export type { CellStylePreset } from '../lib/cell-styles'

export interface CellStylesGalleryProps {
  /** Popover anchor; omit (or null) to render inline. */
  anchor?: AnchorRect | null
  ignoreElement?: HTMLElement | null
  themeColors?: readonly string[]
  /** Apply with `applyCellStylePresetToCell(cell, preset)` for each selected cell. */
  onSelect: (preset: CellStylePreset) => void
  /** Live preview while hovering / focusing a tile; null when the pointer leaves. Optional. */
  onPreview?: (preset: CellStylePreset | null) => void
  /** Popover only: Escape, outside press, and after a style is chosen. */
  onClose?: () => void
  /** Highlights the style currently applied, if known. */
  currentStyleId?: string
}

function tileCss(preset: CellStylePreset, themeColors: readonly string[]): CSSProperties {
  const css = stylePreviewCss(preset.style, themeColors)
  const size = preset.style.font?.size
  if (size) css.fontSize = `${Math.min(15, Math.max(11, size * 0.82))}px`
  if (preset.group === 'number') css.fontStyle = 'normal'
  return css
}

export function CellStylesGallery({ anchor, ignoreElement, themeColors = DEFAULT_THEME_COLORS, onSelect, onPreview, onClose, currentStyleId }: CellStylesGalleryProps) {
  const rootRef = useRef<HTMLDivElement>(null)
  const select = (preset: CellStylePreset) => {
    onPreview?.(null)
    onSelect(preset)
    if (anchor) onClose?.()
  }
  const handleKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const items = Array.from(rootRef.current?.querySelectorAll<HTMLElement>('[data-fmt-style-tile]') || [])
    const moved = moveFocusByArrow(event, items)
    if (moved) items.forEach((item) => item.setAttribute('tabindex', item === moved ? '0' : '-1'))
  }
  const firstFocus = CELL_STYLES.findIndex((preset) => preset.id === currentStyleId)
  let index = -1

  const body = (
    <div ref={rootRef} className="fmt-styles-gallery" onKeyDown={handleKey} onPointerLeave={() => onPreview?.(null)}>
      {CELL_STYLE_GROUPS.map((group) => (
        <section key={group.id} className="fmt-styles-group" aria-label={group.label}>
          <h3>{group.label}</h3>
          <div className="fmt-styles-grid">
            {CELL_STYLES.filter((preset) => preset.group === group.id).map((preset) => {
              index += 1
              const focusable = index === (firstFocus >= 0 ? firstFocus : 0)
              return (
                <button
                  key={preset.id}
                  type="button"
                  data-fmt-style-tile=""
                  tabIndex={focusable ? 0 : -1}
                  className={`fmt-style-tile${preset.id === currentStyleId ? ' is-selected' : ''}${preset.group === 'number' ? ' is-number' : ''}`}
                  style={tileCss(preset, themeColors)}
                  title={preset.name}
                  onClick={() => select(preset)}
                  onPointerEnter={() => onPreview?.(preset)}
                  onFocus={() => onPreview?.(preset)}
                >
                  <span>{preset.name}</span>
                </button>
              )
            })}
          </div>
        </section>
      ))}
    </div>
  )

  if (!anchor) return body
  return (
    <FormatPopover anchor={anchor} ignoreElement={ignoreElement} onClose={() => { onPreview?.(null); onClose?.() }} label="Cell styles" initialFocus='[data-fmt-style-tile][tabindex="0"]'>
      {body}
    </FormatPopover>
  )
}
