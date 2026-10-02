import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { Search } from 'lucide-react'
import './app-polish.css'

export interface PickListPopoverProps {
  /** Viewport rectangle of the cell the list belongs to. */
  anchor: { left: number; top: number; bottom: number; width: number }
  items: string[]
  /** Accessible name, e.g. "Pick from the entries in column B". */
  label: string
  onPick: (value: string) => void
  onClose: () => void
}

/**
 * Alt+Down: Excel's "Pick From Drop-down List" (the column's entries, or the cell's validation
 * list), with type-to-filter as in Excel's searchable dropdowns.
 */
export function PickListPopover({ anchor, items, label, onPick, onClose }: PickListPopoverProps) {
  const [query, setQuery] = useState('')
  const [active, setActive] = useState(0)
  const [position, setPosition] = useState({ left: anchor.left, top: anchor.bottom + 2 })
  const rootRef = useRef<HTMLDivElement>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const id = useId()
  const matches = useMemo(() => {
    const needle = query.trim().toLocaleLowerCase()
    if (!needle) return items
    // Entries that start with the typed text come first, then those that contain it.
    const starts = items.filter((item) => item.toLocaleLowerCase().startsWith(needle))
    const contains = items.filter((item) => !item.toLocaleLowerCase().startsWith(needle) && item.toLocaleLowerCase().includes(needle))
    return [...starts, ...contains]
  }, [items, query])

  useEffect(() => { inputRef.current?.focus() }, [])
  useEffect(() => { setActive(0) }, [query])
  useLayoutEffect(() => {
    const rect = rootRef.current?.getBoundingClientRect()
    if (!rect) return
    const below = window.innerHeight - anchor.bottom - 6
    const top = below < rect.height && anchor.top - rect.height - 2 > 4 ? anchor.top - rect.height - 2 : anchor.bottom + 2
    const left = Math.max(4, Math.min(anchor.left, window.innerWidth - rect.width - 4))
    setPosition((current) => (current.left === left && current.top === top ? current : { left, top }))
  }, [anchor, matches.length])
  useEffect(() => {
    listRef.current?.querySelector<HTMLElement>('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' })
  }, [active])

  return (
    <div className="pick-list-layer" onMouseDown={(event) => { if (event.target === event.currentTarget) { event.preventDefault(); onClose() } }}>
      <div
        ref={rootRef}
        className="pick-list"
        role="dialog"
        aria-modal="true"
        aria-label={label}
        style={{ left: position.left, top: position.top, minWidth: Math.max(160, Math.min(320, anchor.width)) }}
        onKeyDown={(event) => {
          if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onClose(); return }
          if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
            event.preventDefault()
            event.stopPropagation()
            if (!matches.length) return
            const delta = event.key === 'ArrowDown' ? 1 : -1
            setActive((index) => (index + delta + matches.length) % matches.length)
            return
          }
          if (event.key === 'Enter' || event.key === 'Tab') {
            event.preventDefault()
            event.stopPropagation()
            if (matches[active] !== undefined) onPick(matches[active])
            else onClose()
          }
        }}
      >
        <div className="pick-list-search">
          <Search size={12} aria-hidden="true" />
          <input
            ref={inputRef}
            aria-label="Filter the list"
            aria-controls={`${id}-list`}
            aria-activedescendant={matches.length ? `${id}-item-${active}` : undefined}
            placeholder="Type to filter"
            spellCheck={false}
            value={query}
            onChange={(event) => setQuery(event.target.value)}
          />
        </div>
        <div ref={listRef} id={`${id}-list`} className="pick-list-items" role="listbox" aria-label={label}>
          {matches.map((item, index) => (
            <div
              key={`${item}-${index}`}
              id={`${id}-item-${index}`}
              role="option"
              aria-selected={index === active}
              className={`pick-list-item${index === active ? ' is-active' : ''}`}
              title={item.length > 40 ? item : undefined}
              onMouseEnter={() => setActive(index)}
              onMouseDown={(event) => { event.preventDefault(); onPick(item) }}
            >{item}</div>
          ))}
          {!matches.length && <div className="pick-list-empty">{items.length ? 'No matching entries' : 'No entries in this column yet'}</div>}
        </div>
      </div>
    </div>
  )
}
