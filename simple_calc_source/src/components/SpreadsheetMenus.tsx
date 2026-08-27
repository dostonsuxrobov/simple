import { useEffect, useMemo, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { ChevronRight, Search } from 'lucide-react'

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

export function SpreadsheetMenus({ menus }: SpreadsheetMenusProps) {
  const [open, setOpen] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const rootRef = useRef<HTMLDivElement>(null)
  const searchRef = useRef<HTMLInputElement>(null)
  const commands = useMemo(() => flattenItems(menus), [menus])
  const matches = useMemo(() => {
    const needle = query.trim().toLocaleLowerCase()
    return (needle ? commands.filter((item) => item.path.toLocaleLowerCase().includes(needle)) : commands).slice(0, 18)
  }, [commands, query])

  useEffect(() => {
    const closeOutside = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(null)
    }
    const closeWithEscape = (event: KeyboardEvent) => {
      if (event.key !== 'Escape' || !open) return
      event.preventDefault()
      setOpen(null)
    }
    window.addEventListener('pointerdown', closeOutside)
    window.addEventListener('keydown', closeWithEscape)
    return () => {
      window.removeEventListener('pointerdown', closeOutside)
      window.removeEventListener('keydown', closeWithEscape)
    }
  }, [open])

  useEffect(() => {
    if (open !== 'command-search') return
    setQuery('')
    window.requestAnimationFrame(() => searchRef.current?.focus())
  }, [open])

  const close = () => setOpen(null)
  return (
    <div className="spreadsheet-menu-bar" ref={rootRef}>
      <div className="sheet-menu-root">
        <button
          type="button"
          className={`menu-search-trigger${open === 'command-search' ? ' is-open' : ''}`}
          aria-haspopup="menu"
          aria-expanded={open === 'command-search'}
          onClick={() => setOpen((current) => current === 'command-search' ? null : 'command-search')}
        >
          <Search size={13} aria-hidden="true" />
          <span>Menus</span>
        </button>
        {open === 'command-search' && (
          <div className="command-search-panel" role="menu" aria-label="Search menus">
            <div className="command-search-input"><Search size={13} aria-hidden="true" /><input ref={searchRef} aria-label="Search menus" placeholder="Search commands" value={query} onChange={(event) => setQuery(event.target.value)} /></div>
            <div className="command-search-results">
              {matches.map((item) => (
                <button key={item.id} type="button" role="menuitem" data-menu-action={`search-${item.id}`} onClick={() => { item.action?.(); close() }}>
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
          >{menu.label}</button>
          {open === menu.id && (
            <div className="sheet-menu-panel" role="menu" aria-label={`${menu.label} menu`}>
              <MenuItems items={menu.items} close={close} />
            </div>
          )}
        </div>
      ))}
    </div>
  )
}
