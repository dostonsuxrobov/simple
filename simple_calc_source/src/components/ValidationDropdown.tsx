import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { Check, Search } from 'lucide-react'
import { chipTextColor, dropdownOptionColor, splitMultipleSelection } from '../lib/validation'
import type { DropdownPresentation } from '../lib/validation'
import './validation-dropdown.css'

export interface ValidationDropdownProps {
  /** Viewport rectangle of the cell the list belongs to. */
  anchor: { left: number; top: number; bottom: number; width: number }
  options: string[]
  /** The cell's value as displayed (the current pick, or "A, B" for several). */
  current: string
  presentation: DropdownPresentation
  /** Accessible name, e.g. "Choose a value for B2". */
  label: string
  /** One option picked (single choice), or '' to clear the cell. */
  onPick: (value: string) => void
  /** The set of picks when the rule allows several, in the order they were chosen. */
  onPickMany: (values: string[]) => void
  onClose: () => void
}

const fold = (text: string) => text.trim().toLocaleLowerCase()

/**
 * A cell's list dropdown: Excel's searchable list (2024) with Sheets' option colours and, when
 * the rule allows it, multiple picks. Type to filter; Up/Down/Home/End/PageUp/PageDown move;
 * Enter picks (or toggles, for several); Esc cancels. With several picks, clicking outside or
 * Tab keeps the choices, as Sheets applies them.
 */
export function ValidationDropdown({ anchor, options, current, presentation, label, onPick, onPickMany, onClose }: ValidationDropdownProps) {
  const multiple = presentation.multiple
  const [query, setQuery] = useState('')
  const [picks, setPicks] = useState<string[]>(() => (multiple ? splitMultipleSelection(current) : []))
  const pickedKeys = useMemo(() => new Set((multiple ? picks : [current]).map(fold).filter(Boolean)), [current, multiple, picks])
  const matches = useMemo(() => {
    const needle = fold(query)
    if (!needle) return options
    const starts = options.filter((item) => fold(item).startsWith(needle))
    const contains = options.filter((item) => !fold(item).startsWith(needle) && fold(item).includes(needle))
    return [...starts, ...contains]
  }, [options, query])
  const [active, setActive] = useState(() => Math.max(0, options.findIndex((option) => pickedKeys.has(fold(option)))))
  const [position, setPosition] = useState({ left: anchor.left, top: anchor.bottom + 2 })
  const rootRef = useRef<HTMLDivElement>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const picksRef = useRef(picks)
  picksRef.current = picks
  const changedRef = useRef(false)
  const id = useId()

  useEffect(() => { inputRef.current?.focus() }, [])
  // Typing moves to the best match; opening keeps the current value highlighted.
  const lastQueryRef = useRef(query)
  useEffect(() => {
    if (lastQueryRef.current === query) return
    lastQueryRef.current = query
    setActive(0)
  }, [query])
  useLayoutEffect(() => {
    const rect = rootRef.current?.getBoundingClientRect()
    if (!rect) return
    const below = window.innerHeight - anchor.bottom - 6
    const top = below < rect.height && anchor.top - rect.height - 2 > 4 ? anchor.top - rect.height - 2 : anchor.bottom + 2
    const left = Math.max(4, Math.min(anchor.left, window.innerWidth - rect.width - 4))
    setPosition((value) => (value.left === left && value.top === top ? value : { left, top }))
  }, [anchor, matches.length])
  useEffect(() => {
    listRef.current?.querySelector<HTMLElement>('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' })
  }, [active])

  const toggle = (option: string) => {
    changedRef.current = true
    setPicks((list) => (list.some((item) => fold(item) === fold(option)) ? list.filter((item) => fold(item) !== fold(option)) : [...list, option]))
  }
  const choose = (option: string) => (multiple ? toggle(option) : onPick(option))
  /** Leaves the list: several picks are kept unless the list was cancelled with Esc. */
  const finish = (keep: boolean) => {
    if (multiple && keep && changedRef.current) onPickMany(picksRef.current)
    else onClose()
  }
  const hasValue = current.trim() !== '' || (multiple && picks.length > 0)

  return (
    <div className="validation-dropdown-layer" onMouseDown={(event) => { if (event.target === event.currentTarget) { event.preventDefault(); finish(true) } }}>
      <div
        ref={rootRef}
        className={`validation-dropdown${multiple ? ' is-multiple' : ''}`}
        role="dialog"
        aria-modal="true"
        aria-label={label}
        data-validation-dropdown
        style={{ left: position.left, top: position.top, minWidth: Math.max(168, Math.min(320, anchor.width)) }}
        onKeyDown={(event) => {
          const count = matches.length
          const move = (index: number) => { event.preventDefault(); event.stopPropagation(); if (count) setActive(Math.max(0, Math.min(count - 1, index))) }
          if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onClose(); return }
          if (event.key === 'ArrowDown') { move(count ? (active + 1) % count : 0); return }
          if (event.key === 'ArrowUp') { move(count ? (active - 1 + count) % count : 0); return }
          if (event.key === 'PageDown') { move(active + 8); return }
          if (event.key === 'PageUp') { move(active - 8); return }
          if ((event.key === 'Home' || event.key === 'End') && (event.ctrlKey || !query)) { move(event.key === 'Home' ? 0 : count - 1); return }
          if (event.key === 'Enter') {
            event.preventDefault()
            event.stopPropagation()
            if (multiple && (event.ctrlKey || event.metaKey)) { finish(true); return }
            if (matches[active] !== undefined) choose(matches[active])
            else if (!multiple) onClose()
            return
          }
          if (event.key === 'Tab') { event.preventDefault(); event.stopPropagation(); finish(true) }
        }}
      >
        <div className="validation-dropdown-search">
          <Search size={12} aria-hidden="true" />
          <input
            ref={inputRef}
            aria-label="Search the list"
            aria-controls={`${id}-list`}
            aria-activedescendant={matches.length ? `${id}-item-${active}` : undefined}
            placeholder="Search"
            spellCheck={false}
            value={query}
            onChange={(event) => setQuery(event.target.value)}
          />
        </div>
        <div ref={listRef} id={`${id}-list`} className="validation-dropdown-items" role="listbox" aria-label={label} aria-multiselectable={multiple || undefined}>
          {matches.map((option, index) => {
            const color = dropdownOptionColor(presentation, option)
            const picked = pickedKeys.has(fold(option))
            return (
              <div
                key={`${option}-${index}`}
                id={`${id}-item-${index}`}
                role="option"
                aria-selected={index === active}
                aria-checked={multiple ? picked : undefined}
                data-option={option}
                className={`validation-dropdown-item${index === active ? ' is-active' : ''}${picked ? ' is-picked' : ''}`}
                title={option.length > 40 ? option : undefined}
                onMouseEnter={() => setActive(index)}
                onMouseDown={(event) => { event.preventDefault(); choose(option) }}
              >
                <span className={`validation-dropdown-mark${multiple ? ' is-box' : ''}`} aria-hidden="true">{picked && <Check size={11} strokeWidth={2.6} />}</span>
                {color || presentation.style === 'chip'
                  ? <span className="validation-chip" style={color ? { background: color, color: chipTextColor(color) } : undefined}>{option}</span>
                  : <span className="validation-dropdown-text">{option}</span>}
              </div>
            )
          })}
          {!matches.length && <div className="validation-dropdown-empty">{options.length ? 'No matching options' : 'This list has no options'}</div>}
        </div>
        {(hasValue || multiple) && (
          <div className="validation-dropdown-footer">
            {hasValue && (
              <button
                type="button"
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => {
                  if (!multiple) { onPick(''); return }
                  changedRef.current = true
                  setPicks([])
                }}
              >Clear</button>
            )}
            {multiple && (
              <button type="button" className="is-primary" onMouseDown={(event) => event.preventDefault()} onClick={() => finish(true)}>Done</button>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
