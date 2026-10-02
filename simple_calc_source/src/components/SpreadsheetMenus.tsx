import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react'
import { ChevronRight, Search } from 'lucide-react'
import './app-polish.css'

export interface SpreadsheetMenuItem {
  id: string
  label: string
  action?: () => void
  checked?: boolean
  disabled?: boolean
  shortcut?: string
  icon?: ReactNode
  children?: SpreadsheetMenuItem[]
  separatorBefore?: boolean
}

export interface SpreadsheetMenuDefinition {
  id: string
  label: string
  items: SpreadsheetMenuItem[]
}

interface SpreadsheetMenusProps {
  menus: SpreadsheetMenuDefinition[]
  /**
   * Commands that live outside the menu bar (toolbar buttons, context menus): found by the
   * command search ("Menus", Alt+/) but not shown as menus.
   */
  extraCommands?: SpreadsheetMenuDefinition[]
  /** Changing this number opens the command search (Help > Search the menus). */
  searchRequest?: number
}

/** Up/Down/Home/End between the enabled items of one menu level; Right/Left open and close submenus. */
function moveWithinMenu(event: ReactKeyboardEvent<HTMLElement>, onLeaveRoot?: (direction: 1 | -1) => void) {
  const target = event.target instanceof HTMLElement ? event.target : null
  const entry = target?.closest('.sheet-menu-entry')
  const level = entry?.parentElement
  if (!entry || !level) return false
  const buttons = [...level.children]
    .map((child) => child.querySelector<HTMLButtonElement>(':scope > button'))
    .filter((button): button is HTMLButtonElement => Boolean(button && !button.disabled))
  const index = buttons.indexOf(target as HTMLButtonElement)
  if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
    if (!buttons.length) return true
    const delta = event.key === 'ArrowDown' ? 1 : -1
    buttons[(index + delta + buttons.length) % buttons.length]?.focus()
    return true
  }
  if (event.key === 'Home' || event.key === 'End') {
    buttons[event.key === 'Home' ? 0 : buttons.length - 1]?.focus()
    return true
  }
  if (event.key === 'ArrowRight') {
    const child = entry.querySelector<HTMLButtonElement>(':scope > .sheet-submenu > .sheet-menu-entry > button:not(:disabled)')
    if (child) { child.focus(); return true }
    onLeaveRoot?.(1)
    return true
  }
  if (event.key === 'ArrowLeft') {
    const parentEntry = level.closest('.sheet-submenu')?.parentElement
    const parentButton = parentEntry?.querySelector<HTMLButtonElement>(':scope > button')
    if (parentButton) { parentButton.focus(); return true }
    onLeaveRoot?.(-1)
    return true
  }
  return false
}

function MenuItems({ items, close }: { items: SpreadsheetMenuItem[]; close: () => void }) {
  return items.map((item) => (
    <div key={item.id} className={`sheet-menu-entry${item.children?.length ? ' has-children' : ''}${item.separatorBefore ? ' has-separator' : ''}`}>
      <button
        type="button"
        role="menuitem"
        aria-haspopup={item.children?.length ? 'menu' : undefined}
        aria-checked={item.checked === undefined ? undefined : item.checked}
        disabled={item.disabled}
        data-menu-action={item.id}
        onClick={() => {
          if (item.children?.length || !item.action) return
          item.action()
          close()
        }}
      >
        <span className="sheet-menu-check" aria-hidden="true">{item.checked ? '✓' : ''}</span>
        <span className="sheet-menu-icon" aria-hidden="true">{item.icon}</span>
        <span className="sheet-menu-label">{item.label}</span>
        {item.shortcut && <span className="sheet-menu-shortcut">{item.shortcut}</span>}
        {item.children?.length ? <ChevronRight className="sheet-menu-chevron" size={13} aria-hidden="true" /> : null}
      </button>
      {item.children?.length ? (
        <div className="sheet-submenu" role="menu" aria-label={item.label}>
          <MenuItems items={item.children} close={close} />
        </div>
      ) : null}
    </div>
  ))
}

interface SpreadsheetContextMenuProps {
  x: number
  y: number
  label: string
  items: SpreadsheetMenuItem[]
  onClose: (restoreGridFocus?: boolean) => void
}

export function SpreadsheetContextMenu({ x, y, label, items, onClose }: SpreadsheetContextMenuProps) {
  const rootRef = useRef<HTMLDivElement>(null)
  const [position, setPosition] = useState({ left: x, top: y })

  useLayoutEffect(() => {
    const rect = rootRef.current?.getBoundingClientRect()
    if (!rect) return
    setPosition({
      left: Math.max(4, Math.min(x, window.innerWidth - rect.width - 4)),
      top: Math.max(4, Math.min(y, window.innerHeight - rect.height - 4)),
    })
  }, [items, x, y])

  useEffect(() => {
    const closeOutside = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) onClose(false)
    }
    const closeWithEscape = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return
      event.preventDefault()
      onClose(true)
    }
    const closeForViewportChange = () => onClose(false)
    window.addEventListener('pointerdown', closeOutside)
    window.addEventListener('keydown', closeWithEscape)
    window.addEventListener('blur', closeForViewportChange)
    window.addEventListener('resize', closeForViewportChange)
    window.addEventListener('scroll', closeForViewportChange, true)
    return () => {
      window.removeEventListener('pointerdown', closeOutside)
      window.removeEventListener('keydown', closeWithEscape)
      window.removeEventListener('blur', closeForViewportChange)
      window.removeEventListener('resize', closeForViewportChange)
      window.removeEventListener('scroll', closeForViewportChange, true)
    }
  }, [onClose])

  useEffect(() => {
    rootRef.current?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus()
  }, [])

  const moveFocus = (delta: number) => {
    const buttons = [...(rootRef.current?.querySelectorAll<HTMLButtonElement>('button:not(:disabled)') || [])]
    if (!buttons.length) return
    const index = buttons.indexOf(document.activeElement as HTMLButtonElement)
    buttons[(index + delta + buttons.length) % buttons.length].focus()
  }

  return (
    <div
      ref={rootRef}
      className="sheet-context-menu"
      role="menu"
      aria-label={label}
      style={{ left: position.left, top: position.top }}
      onContextMenu={(event) => event.preventDefault()}
      onKeyDown={(event) => {
        if (event.key === 'ArrowDown') { event.preventDefault(); moveFocus(1) }
        else if (event.key === 'ArrowUp') { event.preventDefault(); moveFocus(-1) }
      }}
    >
      <MenuItems items={items} close={() => onClose(true)} />
    </div>
  )
}

function flattenItems(menus: SpreadsheetMenuDefinition[]) {
  const output: Array<SpreadsheetMenuItem & { path: string }> = []
  const visit = (items: SpreadsheetMenuItem[], path: string) => {
    items.forEach((item) => {
      const nextPath = `${path} › ${item.label}`
      if (item.children?.length) visit(item.children, nextPath)
      else if (item.action && !item.disabled) output.push({ ...item, path: nextPath })
    })
  }
  menus.forEach((menu) => visit(menu.items, menu.label))
  return output
}

/** Commands whose path or shortcut matches every word of the query, best matches first. */
function searchCommands(commands: Array<SpreadsheetMenuItem & { path: string }>, query: string) {
  const words = query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean)
  if (!words.length) return commands.slice(0, 18)
  const scored: Array<{ item: SpreadsheetMenuItem & { path: string }; score: number }> = []
  for (const item of commands) {
    const path = item.path.toLocaleLowerCase()
    const label = item.label.toLocaleLowerCase()
    const shortcut = (item.shortcut || '').toLocaleLowerCase()
    if (!words.every((word) => path.includes(word) || shortcut.includes(word))) continue
    // The command's own name beats its menu path: "copy" finds Edit › Copy before Copy sheet.
    const score = (label.startsWith(words[0]) ? 0 : label.includes(words[0]) ? 1 : 2) * 1000 + label.length
    scored.push({ item, score })
  }
  return scored.sort((a, b) => a.score - b.score).slice(0, 18).map((entry) => entry.item)
}

export function SpreadsheetMenus({ menus, extraCommands, searchRequest }: SpreadsheetMenusProps) {
  const [open, setOpen] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [activeIndex, setActiveIndex] = useState(0)
  const rootRef = useRef<HTMLDivElement>(null)
  const searchRef = useRef<HTMLInputElement>(null)
  const resultsRef = useRef<HTMLDivElement>(null)
  // Where the keyboard was before the search opened, so Esc can return it there.
  const returnFocusRef = useRef<HTMLElement | null>(null)
  const commands = useMemo(() => {
    const fromMenus = flattenItems(menus)
    // A toolbar or context command that is also in a menu is listed once, under its menu.
    const known = new Set(fromMenus.map((item) => `${item.label.toLocaleLowerCase()}|${item.shortcut || ''}`))
    const extra = flattenItems(extraCommands || []).filter((item) => !known.has(`${item.label.toLocaleLowerCase()}|${item.shortcut || ''}`))
    return [...fromMenus, ...extra]
  }, [extraCommands, menus])
  const matches = useMemo(() => searchCommands(commands, query), [commands, query])
  const activeMatch = matches[Math.min(activeIndex, matches.length - 1)]

  useEffect(() => {
    const closeOutside = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(null)
    }
    const closeWithEscape = (event: KeyboardEvent) => {
      if (event.key !== 'Escape' || !open) return
      event.preventDefault()
      const menuId = open
      setOpen(null)
      // The keyboard goes back to the sheet (or wherever Alt+/ opened the search from).
      window.requestAnimationFrame(() => {
        const previous = menuId === 'command-search' ? returnFocusRef.current : null
        if (previous?.isConnected) previous.focus({ preventScroll: true })
        else document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })
      })
    }
    window.addEventListener('pointerdown', closeOutside)
    window.addEventListener('keydown', closeWithEscape)
    return () => {
      window.removeEventListener('pointerdown', closeOutside)
      window.removeEventListener('keydown', closeWithEscape)
    }
  }, [open])

  // Alt+/ opens the command search from anywhere, as in Google Sheets.
  useEffect(() => {
    const openSearch = (event: KeyboardEvent) => {
      if (!event.altKey || event.ctrlKey || event.metaKey || (event.key !== '/' && event.code !== 'Slash')) return
      if (document.querySelector('[aria-modal="true"]')) return
      event.preventDefault()
      returnFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
      setOpen('command-search')
    }
    window.addEventListener('keydown', openSearch)
    return () => window.removeEventListener('keydown', openSearch)
  }, [])

  useEffect(() => {
    if (!searchRequest) return
    returnFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
    setOpen('command-search')
  }, [searchRequest])

  useEffect(() => {
    if (open !== 'command-search') return
    setQuery('')
    setActiveIndex(0)
    window.requestAnimationFrame(() => searchRef.current?.focus())
  }, [open])

  useEffect(() => { setActiveIndex(0) }, [query])
  useEffect(() => {
    resultsRef.current?.querySelector<HTMLElement>('.is-active')?.scrollIntoView({ block: 'nearest' })
  }, [activeIndex])

  const close = () => setOpen(null)
  const runCommand = (item: SpreadsheetMenuItem | undefined) => {
    if (!item?.action) return
    close()
    item.action()
  }
  const switchMenu = (menuId: string, direction: 1 | -1) => {
    const index = menus.findIndex((menu) => menu.id === menuId)
    const next = menus[(index + direction + menus.length) % menus.length]
    if (!next) return
    setOpen(next.id)
    window.requestAnimationFrame(() => rootRef.current?.querySelector<HTMLButtonElement>('.sheet-menu-panel .sheet-menu-entry > button:not(:disabled)')?.focus())
  }
  return (
    <div className="spreadsheet-menu-bar" ref={rootRef}>
      <div className="sheet-menu-root">
        <button
          type="button"
          className={`menu-search-trigger${open === 'command-search' ? ' is-open' : ''}`}
          aria-haspopup="menu"
          aria-expanded={open === 'command-search'}
          title="Search the menus (Alt+/)"
          onClick={() => {
            returnFocusRef.current = null
            setOpen((current) => current === 'command-search' ? null : 'command-search')
          }}
        >
          <Search size={13} aria-hidden="true" />
          <span>Menus</span>
        </button>
        {open === 'command-search' && (
          <div className="command-search-panel" role="dialog" aria-label="Command search">
            <div className="command-search-input">
              <Search size={13} aria-hidden="true" />
              <input
                ref={searchRef}
                role="combobox"
                aria-label="Search menus"
                aria-expanded="true"
                aria-controls="command-search-results"
                aria-activedescendant={activeMatch ? `command-search-${activeMatch.id}` : undefined}
                placeholder="Search commands"
                value={query}
                onChange={(event) => setQuery(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                    event.preventDefault()
                    if (!matches.length) return
                    const delta = event.key === 'ArrowDown' ? 1 : -1
                    setActiveIndex((index) => (Math.min(index, matches.length - 1) + delta + matches.length) % matches.length)
                  } else if (event.key === 'Enter') {
                    event.preventDefault()
                    runCommand(activeMatch)
                  } else if (event.key === 'Tab') {
                    event.preventDefault()
                  }
                }}
              />
            </div>
            <div className="command-search-results" id="command-search-results" role="listbox" aria-label="Commands" ref={resultsRef}>
              {matches.map((item, index) => (
                <button
                  key={item.id}
                  id={`command-search-${item.id}`}
                  type="button"
                  role="option"
                  tabIndex={-1}
                  aria-selected={item === activeMatch}
                  className={item === activeMatch ? 'is-active' : undefined}
                  data-menu-action={`search-${item.id}`}
                  onMouseEnter={() => setActiveIndex(index)}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => runCommand(item)}
                >
                  <span>{item.path}</span>{item.shortcut && <small>{item.shortcut}</small>}
                </button>
              ))}
              {!matches.length && <p>No matching commands</p>}
            </div>
          </div>
        )}
      </div>
      {menus.map((menu) => (
        <div className="sheet-menu-root" key={menu.id}>
          <button
            type="button"
            className={`sheet-menu-trigger${open === menu.id ? ' is-open' : ''}`}
            aria-haspopup="menu"
            aria-expanded={open === menu.id}
            data-menu-trigger={menu.id}
            onPointerEnter={() => { if (open && open !== 'command-search') setOpen(menu.id) }}
            onClick={() => setOpen((current) => current === menu.id ? null : menu.id)}
            onKeyDown={(event) => {
              if (event.key === 'ArrowDown' || ((event.key === 'Enter' || event.key === ' ') && open !== menu.id)) {
                event.preventDefault()
                setOpen(menu.id)
                window.requestAnimationFrame(() => rootRef.current?.querySelector<HTMLButtonElement>('.sheet-menu-panel .sheet-menu-entry > button:not(:disabled)')?.focus())
              } else if (event.key === 'ArrowRight' || event.key === 'ArrowLeft') {
                event.preventDefault()
                const index = menus.findIndex((item) => item.id === menu.id)
                const next = menus[(index + (event.key === 'ArrowRight' ? 1 : -1) + menus.length) % menus.length]
                rootRef.current?.querySelector<HTMLButtonElement>(`[data-menu-trigger="${next.id}"]`)?.focus()
                if (open) setOpen(next.id)
              }
            }}
          >{menu.label}</button>
          {open === menu.id && (
            <div
              className="sheet-menu-panel"
              role="menu"
              aria-label={`${menu.label} menu`}
              onKeyDown={(event) => {
                if (moveWithinMenu(event, (direction) => switchMenu(menu.id, direction))) event.preventDefault()
              }}
            >
              <MenuItems items={menu.items} close={close} />
            </div>
          )}
        </div>
      ))}
    </div>
  )
}
