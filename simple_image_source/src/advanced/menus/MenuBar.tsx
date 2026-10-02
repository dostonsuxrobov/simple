// src/advanced/menus/MenuBar.tsx (WP6)
// The compact menu row of the Advanced editor (design 3.2, 5.14): Edit, Image, Layer, Select, Filter, View.
//   - MENUS is plain data (command ids, submenus, separators); labels, shortcuts and the enabled state come
//     from the command registry when a menu opens, so a menu never offers what cannot run right now.
//   - Behaves like a desktop menu bar: click opens, hovering another title switches while one is open,
//     submenus open on hover or the Right arrow; Up / Down / Home / End move, Enter or Space runs,
//     Left / Right walk the bar, Escape closes and returns focus to the canvas. Keys used here never
//     reach the editor's shortcuts (handled and stopped at the menu).
//   - ARIA: menubar > menuitem[aria-haspopup] > menu > menuitem / menuitemcheckbox.
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react'
import { Check, ChevronRight } from 'lucide-react'
import type { AdjustmentType, FilterType } from '../../imaging/types.ts'
import type { EditorCommandId } from '../commands.ts'

export type MenuItem =
  | { readonly kind: 'command'; readonly command: EditorCommandId; readonly label?: string }
  | { readonly kind: 'check'; readonly command: EditorCommandId; readonly label?: string }
  | { readonly kind: 'submenu'; readonly label: string; readonly items: readonly MenuItem[] }
  | { readonly kind: 'separator' }

export interface Menu {
  readonly id: string
  readonly label: string
  readonly items: readonly MenuItem[]
}

const cmd = (command: EditorCommandId, label?: string): MenuItem => ({ kind: 'command', command, label })
const check = (command: EditorCommandId, label?: string): MenuItem => ({ kind: 'check', command, label })
const sub = (label: string, items: readonly MenuItem[]): MenuItem => ({ kind: 'submenu', label, items })
const SEP: MenuItem = { kind: 'separator' }

const ADJUSTMENTS: readonly (AdjustmentType | '-')[] = [
  'brightness-contrast', 'levels', 'curves', 'exposure', '-',
  'vibrance', 'hue-saturation', 'color-balance', 'black-white', 'photo-filter', '-',
  'invert', 'posterize', 'threshold', 'gradient-map',
]

const adjustmentItems = (prefix: 'adjust' | 'adjustment-layer'): MenuItem[] => ADJUSTMENTS.map((type) => (type === '-'
  ? SEP
  : cmd(`${prefix}.${type}` as EditorCommandId)))

const filter = (type: FilterType): MenuItem => cmd(`filter.${type}` as EditorCommandId)

/** The editor's menus (Photoshop order and names). */
export const MENUS: readonly Menu[] = Object.freeze([
  {
    id: 'edit',
    label: 'Edit',
    items: [
      cmd('edit.undo'), cmd('edit.redo'), cmd('edit.toggle-last'), SEP,
      cmd('edit.cut'), cmd('edit.copy'), cmd('edit.copy-merged'), cmd('edit.paste'), cmd('edit.paste-in-place'), cmd('edit.clear'), SEP,
      cmd('edit.fill-foreground'), cmd('edit.fill-background'), cmd('edit.fill-foreground-preserve'), SEP,
      cmd('edit.free-transform'),
    ],
  },
  {
    id: 'image',
    label: 'Image',
    items: [
      sub('Adjustments', [...adjustmentItems('adjust'), SEP, cmd('image.desaturate')]),
      SEP,
      cmd('image.auto-tone'), cmd('image.auto-contrast'), cmd('image.auto-color'), SEP,
      cmd('image.size'), cmd('image.canvas-size'),
      sub('Image Rotation', [
        cmd('image.rotate-180'), cmd('image.rotate-cw'), cmd('image.rotate-ccw'), cmd('image.rotate-arbitrary'), SEP,
        cmd('image.flip-horizontal'), cmd('image.flip-vertical'),
      ]),
      cmd('image.crop-to-selection'), cmd('image.trim'),
    ],
  },
  {
    id: 'layer',
    label: 'Layer',
    items: [
      cmd('layer.new'), cmd('layer.via-copy'), cmd('layer.via-cut'),
      sub('New Adjustment Layer', adjustmentItems('adjustment-layer')),
      cmd('layer.duplicate'), cmd('layer.delete'), SEP,
      sub('Layer Mask', [cmd('layer.add-mask'), cmd('layer.delete-mask'), cmd('layer.apply-mask'), cmd('layer.toggle-mask')]),
      cmd('layer.toggle-clipping'), cmd('layer.rasterize'), cmd('layer.from-background'), SEP,
      sub('Arrange', [cmd('layer.to-front'), cmd('layer.raise'), cmd('layer.lower'), cmd('layer.to-back')]),
      SEP,
      cmd('layer.merge-down'), cmd('layer.merge-visible'), cmd('layer.stamp-visible'), cmd('layer.flatten'),
    ],
  },
  {
    id: 'select',
    label: 'Select',
    items: [
      cmd('select.all'), cmd('select.deselect'), cmd('select.reselect'), cmd('select.inverse'), SEP,
      sub('Modify', [cmd('select.feather'), cmd('select.expand'), cmd('select.contract')]),
      cmd('select.load-layer-alpha'),
    ],
  },
  {
    id: 'filter',
    label: 'Filter',
    items: [
      cmd('filter.repeat'), SEP,
      sub('Blur', [filter('gaussian-blur'), filter('motion-blur')]),
      sub('Sharpen', [filter('sharpen'), cmd('filter.sharpen-more'), filter('unsharp-mask')]),
      sub('Noise', [filter('add-noise'), filter('median'), filter('reduce-noise')]),
      sub('Pixelate', [filter('pixelate')]),
      sub('Stylize', [filter('emboss'), filter('find-edges')]),
    ],
  },
  {
    id: 'view',
    label: 'View',
    items: [
      cmd('view.zoom-in'), cmd('view.zoom-out'), cmd('view.fit'), cmd('view.actual-pixels'), SEP,
      check('view.panel-layers'), check('view.panel-properties'), check('view.panel-history'), check('view.panel-color'), SEP,
      cmd('view.toggle-panels'),
    ],
  },
] satisfies Menu[])

/** Every command reachable from the menus (tests and the smoke check that each one is registered). */
export function menuCommands(menus: readonly Menu[] = MENUS): EditorCommandId[] {
  const out: EditorCommandId[] = []
  const walk = (items: readonly MenuItem[]) => {
    for (const item of items) {
      if (item.kind === 'command' || item.kind === 'check') out.push(item.command)
      else if (item.kind === 'submenu') walk(item.items)
    }
  }
  for (const menu of menus) walk(menu.items)
  return out
}

export interface MenuBarProps {
  readonly menus?: readonly Menu[]
  readonly run: (command: EditorCommandId) => void
  readonly isEnabled: (command: EditorCommandId) => boolean
  readonly label: (command: EditorCommandId) => string
  readonly shortcut: (command: EditorCommandId) => string
  readonly isChecked?: (command: EditorCommandId) => boolean
  /** Host modal or editor busy: the bar is inert. */
  readonly disabled?: boolean
  /** Called when a menu closes without running a command (focus goes back to the canvas). */
  readonly onDismiss?: () => void
}

type Path = readonly number[]

function itemsAt(menu: Menu, path: Path): readonly MenuItem[] {
  let items = menu.items
  for (const index of path) {
    const item = items[index]
    if (!item || item.kind !== 'submenu') return []
    items = item.items
  }
  return items
}

function focusable(items: readonly MenuItem[], enabled: (item: MenuItem) => boolean): number[] {
  const out: number[] = []
  items.forEach((item, index) => {
    if (item.kind !== 'separator' && enabled(item)) out.push(index)
  })
  return out
}

function hasEnabledItem(items: readonly MenuItem[], isEnabled: (command: EditorCommandId) => boolean): boolean {
  return items.some((item) => (item.kind === 'command' || item.kind === 'check')
    ? isEnabled(item.command)
    : item.kind === 'submenu' ? hasEnabledItem(item.items, isEnabled) : false)
}

export function MenuBar({ menus = MENUS, run, isEnabled, label, shortcut, isChecked, disabled, onDismiss }: MenuBarProps) {
  const [open, setOpen] = useState<number | null>(null)
  /** Open submenu chain (indices into the item lists) and the focused item per level. */
  const [subPath, setSubPath] = useState<number[]>([])
  const [focus, setFocus] = useState<number[]>([])
  const barRef = useRef<HTMLDivElement>(null)
  const titleRefs = useRef<(HTMLButtonElement | null)[]>([])
  const itemRefs = useRef(new Map<string, HTMLButtonElement>())

  const close = useCallback((dismissed: boolean) => {
    setOpen(null)
    setSubPath([])
    setFocus([])
    if (dismissed) onDismiss?.()
  }, [onDismiss])

  // Close on outside pointer, window blur and when the bar becomes inert.
  useEffect(() => {
    if (open === null) return
    const onPointer = (event: PointerEvent) => {
      if (barRef.current && event.target instanceof Node && barRef.current.contains(event.target)) return
      close(false)
    }
    const onBlur = () => close(false)
    document.addEventListener('pointerdown', onPointer, true)
    window.addEventListener('blur', onBlur)
    return () => {
      document.removeEventListener('pointerdown', onPointer, true)
      window.removeEventListener('blur', onBlur)
    }
  }, [close, open])

  useEffect(() => {
    if (disabled && open !== null) close(false)
  }, [close, disabled, open])

  // Keyboard focus follows the focused item.
  useLayoutEffect(() => {
    if (open === null) return
    const level = focus.length - 1
    if (level < 0) return
    const key = [...subPath.slice(0, level), focus[level]].join('.')
    itemRefs.current.get(`${open}:${key}`)?.focus({ preventScroll: true })
  }, [focus, open, subPath])

  const enabledItem = useCallback((item: MenuItem) => (item.kind === 'command' || item.kind === 'check')
    ? isEnabled(item.command)
    : item.kind === 'submenu' ? hasEnabledItem(item.items, isEnabled) : false, [isEnabled])

  const openMenu = (index: number, focusFirst: boolean) => {
    setOpen(index)
    setSubPath([])
    const menu = menus[index]
    const first = menu ? focusable(menu.items, enabledItem)[0] : undefined
    setFocus(focusFirst && first !== undefined ? [first] : [])
    if (!focusFirst) titleRefs.current[index]?.focus({ preventScroll: true })
  }

  const activate = (item: MenuItem, level: number, index: number) => {
    if (item.kind === 'separator') return
    if (item.kind === 'submenu') {
      if (!enabledItem(item)) return
      const nextPath = [...subPath.slice(0, level), index]
      setSubPath(nextPath)
      const first = focusable(item.items, enabledItem)[0]
      setFocus([...focus.slice(0, level), index, ...(first !== undefined ? [first] : [])])
      return
    }
    if (!isEnabled(item.command)) return
    close(false)
    run(item.command)
  }

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (open === null) {
      if ((event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') && event.target instanceof HTMLElement) {
        const index = titleRefs.current.indexOf(event.target as HTMLButtonElement)
        if (index >= 0) {
          event.preventDefault()
          event.stopPropagation()
          openMenu(index, true)
        }
      } else if ((event.key === 'ArrowRight' || event.key === 'ArrowLeft') && event.target instanceof HTMLElement) {
        const index = titleRefs.current.indexOf(event.target as HTMLButtonElement)
        if (index >= 0) {
          event.preventDefault()
          event.stopPropagation()
          const next = (index + (event.key === 'ArrowRight' ? 1 : -1) + menus.length) % menus.length
          titleRefs.current[next]?.focus()
        }
      }
      return
    }
    const menu = menus[open]
    if (!menu) return
    const level = Math.max(0, focus.length - 1)
    const items = itemsAt(menu, subPath.slice(0, level))
    const indices = focusable(items, enabledItem)
    const current = focus[level]
    const move = (to: number | undefined) => {
      if (to === undefined) return
      setFocus([...focus.slice(0, level), to])
      setSubPath(subPath.slice(0, level))
    }
    let handled = true
    switch (event.key) {
      case 'ArrowDown': {
        const at = indices.indexOf(current)
        move(indices[(at + 1 + indices.length) % indices.length])
        break
      }
      case 'ArrowUp': {
        const at = indices.indexOf(current)
        move(indices[at < 0 ? indices.length - 1 : (at - 1 + indices.length) % indices.length])
        break
      }
      case 'Home':
        move(indices[0])
        break
      case 'End':
        move(indices[indices.length - 1])
        break
      case 'ArrowRight': {
        const item = current !== undefined ? items[current] : undefined
        if (item && item.kind === 'submenu') activate(item, level, current)
        else openMenu((open + 1) % menus.length, true)
        break
      }
      case 'ArrowLeft':
        if (level > 0) {
          setFocus(focus.slice(0, level))
          setSubPath(subPath.slice(0, level - 1))
        } else {
          openMenu((open - 1 + menus.length) % menus.length, true)
        }
        break
      case 'Enter':
      case ' ': {
        const item = current !== undefined ? items[current] : undefined
        if (item) activate(item, level, current)
        break
      }
      case 'Escape':
        if (level > 0) {
          setFocus(focus.slice(0, level))
          setSubPath(subPath.slice(0, level - 1))
        } else {
          close(true)
        }
        break
      case 'Tab':
        close(true)
        break
      default:
        handled = false
    }
    if (handled) {
      event.preventDefault()
    }
    // Keys pressed inside an open menu never become editor shortcuts.
    event.stopPropagation()
  }

  const renderItems = (menuIndex: number, items: readonly MenuItem[], level: number, prefix: number[]): ReactNode => (
    <div className="ae-menu" role="menu" aria-label={level === 0 ? menus[menuIndex]?.label : undefined}>
      {items.map((item, index) => {
        if (item.kind === 'separator') return <div key={`sep-${index}`} className="ae-menu-separator" role="separator" />
        const key = [...prefix, index].join('.')
        const enabled = enabledItem(item)
        const refKey = `${menuIndex}:${key}`
        const setRef = (element: HTMLButtonElement | null) => {
          if (element) itemRefs.current.set(refKey, element)
          else itemRefs.current.delete(refKey)
        }
        if (item.kind === 'submenu') {
          const expanded = subPath[level] === index
          return (
            <div key={key} className="ae-submenu">
              <button
                ref={setRef}
                type="button"
                role="menuitem"
                aria-haspopup="menu"
                aria-expanded={expanded}
                aria-disabled={!enabled || undefined}
                className={`ae-menu-item ${expanded ? 'is-open' : ''}`}
                tabIndex={-1}
                onPointerEnter={() => enabled && activate(item, level, index)}
                onClick={() => activate(item, level, index)}
              >
                <span className="ae-menu-check" />
                <span className="ae-menu-label">{item.label}</span>
                <ChevronRight className="ae-menu-arrow" aria-hidden="true" />
              </button>
              {expanded && renderItems(menuIndex, item.items, level + 1, [...prefix, index])}
            </div>
          )
        }
        const text = item.label ?? label(item.command)
        const keys = shortcut(item.command)
        const checked = item.kind === 'check' ? Boolean(isChecked?.(item.command)) : undefined
        return (
          <button
            key={key}
            ref={setRef}
            type="button"
            role={item.kind === 'check' ? 'menuitemcheckbox' : 'menuitem'}
            aria-checked={checked}
            aria-disabled={!enabled || undefined}
            aria-keyshortcuts={keys || undefined}
            className="ae-menu-item"
            tabIndex={-1}
            onPointerEnter={() => {
              setSubPath(subPath.slice(0, level))
              if (enabled) setFocus([...focus.slice(0, level), index])
            }}
            onClick={() => activate(item, level, index)}
          >
            <span className="ae-menu-check">{checked ? <Check aria-hidden="true" /> : null}</span>
            <span className="ae-menu-label">{text}</span>
            {keys && <span className="ae-menu-keys">{keys}</span>}
          </button>
        )
      })}
    </div>
  )

  return (
    <div ref={barRef} className="ae-menubar" role="menubar" aria-label="Advanced editor menus" onKeyDown={onKeyDown}>
      {menus.map((menu, index) => (
        <div key={menu.id} className="ae-menubar-entry">
          <button
            ref={(element) => { titleRefs.current[index] = element }}
            type="button"
            role="menuitem"
            aria-haspopup="menu"
            aria-expanded={open === index}
            className={`ae-menubar-title ${open === index ? 'is-open' : ''}`}
            disabled={disabled}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => (open === index ? close(true) : openMenu(index, false))}
            onPointerEnter={() => {
              if (open !== null && open !== index) openMenu(index, false)
            }}
          >{menu.label}</button>
          {open === index && renderItems(index, menu.items, 0, [])}
        </div>
      ))}
    </div>
  )
}
