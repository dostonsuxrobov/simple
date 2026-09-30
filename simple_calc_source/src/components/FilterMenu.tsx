import { useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react'
import { ArrowDownAZ, ArrowUpZA, Check, ChevronRight, FilterX, Palette, Search, X } from 'lucide-react'
import type { SheetFilterCriteria } from '../spreadsheet-types'
import { FILTER_OPERATORS, filterOperatorInfo } from '../lib/filter'
import type { DistinctValue, DistinctValuesResult, FilterOperatorGroup } from '../lib/filter'
import { Swatch, colorLabel, trapTab, useInitialFocus } from './DataToolsDialogFrame'
import './data-tools.css'

export interface FilterMenuAnchor { left: number; top: number; right: number; bottom: number }

export interface FilterMenuProps {
  /** Header text of the column, e.g. "Region". */
  columnLabel: string
  /** The column's current criteria (null / undefined when unfiltered). */
  criteria?: SheetFilterCriteria | null
  /** From distinctColumnValues(state, column, host). */
  values: DistinctValuesResult
  /** Viewport rect of the header filter button. */
  anchor: FilterMenuAnchor
  /** Marks the active sort on this column, if any. */
  sortDirection?: 'asc' | 'desc' | null
  onApply: (criteria: SheetFilterCriteria | null) => void
  onSort: (descending: boolean) => void
  onSortByColor?: (target: { sortOn: 'fillColor' | 'fontColor'; color: string | null }) => void
  onClose: () => void
}

type Mode = 'values' | 'condition' | 'color'

const ROW_HEIGHT = 24
const LIST_HEIGHT = 216
const OVERSCAN = 6
const BLANKS_KEY = '\u0000blanks'

const GROUP_LABELS: Record<FilterOperatorGroup, string> = { general: 'General', text: 'Text', number: 'Number', date: 'Date' }

function initialMode(criteria: SheetFilterCriteria | null | undefined): Mode {
  if (criteria?.condition?.operator) return 'condition'
  if (criteria?.fillColor || criteria?.fontColor) return 'color'
  return 'values'
}

function sortLabels(kind: DistinctValuesResult['kind']): [string, string] {
  if (kind === 'number') return ['Sort smallest to largest', 'Sort largest to smallest']
  if (kind === 'date') return ['Sort oldest to newest', 'Sort newest to oldest']
  return ['Sort A to Z', 'Sort Z to A']
}

function inputPlaceholder(operator: string, second = false): string {
  const info = filterOperatorInfo(operator)
  if (!info) return 'Value'
  if (operator === 'top' || operator === 'bottom') return 'Items (e.g. 10)'
  if (operator === 'topPercent' || operator === 'bottomPercent') return 'Percent (e.g. 10)'
  if (info.input === 'date') return second ? 'End date' : info.inputs === 2 ? 'Start date' : 'm/d/yyyy'
  if (info.input === 'formula') return '=B2>100'
  if (info.inputs === 2) return second ? 'Maximum' : 'Minimum'
  return 'Value'
}

function Disclosure({ title, open, onToggle, active, children, icon }: { title: string; open: boolean; onToggle: () => void; active?: boolean; children: ReactNode; icon?: ReactNode }) {
  return (
    <div className={`dt-disclosure${open ? ' is-open' : ''}`}>
      <button type="button" className="dt-disclosure-toggle" aria-expanded={open} onClick={onToggle}>
        <ChevronRight className="dt-disclosure-chevron" size={13} aria-hidden="true" />
        {icon ? <span className="dt-disclosure-icon" aria-hidden="true">{icon}</span> : null}
        <span>{title}</span>
        {active ? <span className="dt-active-dot" aria-label="active" /> : null}
      </button>
      {open ? <div className="dt-disclosure-body">{children}</div> : null}
    </div>
  )
}

function OperatorSelect({ value, onChange, combinableOnly, label, groups }: { value: string; onChange: (value: string) => void; combinableOnly?: boolean; label: string; groups: FilterOperatorGroup[] }) {
  return (
    <select className="dt-select" aria-label={label} value={value} onChange={(event) => onChange(event.target.value)}>
      <option value="">{combinableOnly ? '(no second condition)' : 'None'}</option>
      {groups.map((group) => {
        const operators = FILTER_OPERATORS.filter((info) => info.group === group && (!combinableOnly || info.combinable))
        if (!operators.length) return null
        return (
          <optgroup key={group} label={GROUP_LABELS[group]}>
            {operators.map((info) => <option key={info.id} value={info.id}>{info.label}</option>)}
          </optgroup>
        )
      })}
    </select>
  )
}

/**
 * Excel / Sheets style column filter dropdown: sort, filter by colour, by condition (two
 * conditions with And / Or) and by values (searchable, virtualised checklist with counts).
 */
export function FilterMenu({ columnLabel, criteria, values, anchor, sortDirection, onApply, onSort, onSortByColor, onClose }: FilterMenuProps) {
  const rootRef = useRef<HTMLDivElement>(null)
  const searchRef = useRef<HTMLInputElement>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const [position, setPosition] = useState<{ left: number; top: number; maxHeight: number }>({ left: anchor.left, top: anchor.bottom + 2, maxHeight: 600 })
  const [mode, setMode] = useState<Mode>(() => initialMode(criteria))
  const [open, setOpen] = useState<Record<string, boolean>>(() => ({
    condition: initialMode(criteria) === 'condition',
    values: initialMode(criteria) !== 'condition',
    color: initialMode(criteria) === 'color',
    sortColor: false,
  }))
  // Checklist state: keys the user unchecked (all checked by default).
  const [excluded, setExcluded] = useState<Set<string>>(() => {
    const set = new Set<string>()
    for (const item of values.items) if (!item.checked) set.add(item.key)
    if (values.blanks && !values.blanks.checked) set.add(BLANKS_KEY)
    return set
  })
  const [search, setSearch] = useState('')
  // While searching, results start fully checked (Excel); unchecks apply to the search only.
  const [searchExcluded, setSearchExcluded] = useState<Set<string>>(() => new Set())
  const [addToSelection, setAddToSelection] = useState(false)
  const [scrollTop, setScrollTop] = useState(0)
  const [operator, setOperator] = useState(criteria?.condition?.operator || '')
  const [value1, setValue1] = useState(criteria?.condition?.value === undefined ? '' : String(criteria.condition.value))
  const [value2, setValue2] = useState(criteria?.condition?.value2 === undefined ? '' : String(criteria.condition.value2))
  const [operator2, setOperator2] = useState(criteria?.condition?.operator2 || '')
  const [join, setJoin] = useState<'and' | 'or'>(criteria?.condition?.join === 'or' ? 'or' : 'and')
  const [colorChoice, setColorChoice] = useState<{ kind: 'fill' | 'font'; color: string | null } | null>(() => (
    criteria?.fillColor ? { kind: 'fill', color: criteria.fillColor === 'none' ? null : criteria.fillColor }
      : criteria?.fontColor ? { kind: 'font', color: criteria.fontColor === 'none' || criteria.fontColor === 'auto' ? null : criteria.fontColor } : null
  ))

  useInitialFocus(rootRef, searchRef)

  useLayoutEffect(() => {
    const rect = rootRef.current?.getBoundingClientRect()
    const width = rect?.width || 280
    const height = rect?.height || 480
    const margin = 6
    const below = window.innerHeight - anchor.bottom - margin
    const above = anchor.top - margin
    const placeAbove = below < Math.min(height, 420) && above > below
    const maxHeight = Math.max(240, (placeAbove ? above : below) - 4)
    const top = placeAbove ? Math.max(margin, anchor.top - Math.min(height, maxHeight) - 2) : anchor.bottom + 2
    const left = Math.max(margin, Math.min(anchor.left, window.innerWidth - width - margin))
    setPosition({ left, top, maxHeight })
  }, [anchor.bottom, anchor.left, anchor.top, open])

  const query = search.trim().toLocaleLowerCase()
  const visibleItems = useMemo(() => (query ? values.items.filter((item) => item.text.toLocaleLowerCase().includes(query)) : values.items), [query, values.items])
  const showBlanksRow = Boolean(values.blanks) && (!query || '(blanks)'.includes(query))
  const rowCount = visibleItems.length + (showBlanksRow ? 1 : 0)
  const isChecked = (key: string) => (query ? !searchExcluded.has(key) : !excluded.has(key))
  const checkedVisible = visibleItems.reduce((count, item) => count + (isChecked(item.key) ? 1 : 0), 0) + (showBlanksRow && isChecked(BLANKS_KEY) ? 1 : 0)
  const allState = checkedVisible === 0 ? 'none' : checkedVisible === rowCount ? 'all' : 'some'

  const groups: FilterOperatorGroup[] = values.kind === 'number' ? ['number', 'general', 'text', 'date']
    : values.kind === 'date' ? ['date', 'general', 'number', 'text'] : ['text', 'general', 'number', 'date']
  const operatorInfo = filterOperatorInfo(operator)
  const secondInfo = filterOperatorInfo(operator2)
  const suggestions = useMemo(() => values.items.slice(0, 300).map((item) => item.text), [values.items])
  const datalistId = useRef(`dt-filter-values-${Math.random().toString(36).slice(2, 8)}`).current

  const hasFillColors = values.fillColors.length > 1
  const hasFontColors = values.fontColors.length > 1
  const hasColors = hasFillColors || hasFontColors
  const [ascendingLabel, descendingLabel] = sortLabels(values.kind)
  const hasCriteria = Boolean(criteria && (criteria.values || criteria.blanks === false || criteria.condition?.operator || criteria.fillColor || criteria.fontColor))

  const setChecked = (keys: string[], checked: boolean) => {
    setMode('values')
    ;(query ? setSearchExcluded : setExcluded)((current) => {
      const next = new Set(current)
      for (const key of keys) {
        if (checked) next.delete(key)
        else next.add(key)
      }
      return next
    })
  }

  const buildCriteria = (): SheetFilterCriteria | null | 'invalid' => {
    if (mode === 'condition') {
      if (!operator) return null
      const info = filterOperatorInfo(operator)
      if (!info) return null
      if (info.inputs >= 1 && !value1.trim() && info.input !== 'text') return 'invalid'
      if (info.inputs === 2 && !value2.trim()) return 'invalid'
      const condition: NonNullable<SheetFilterCriteria['condition']> = { operator }
      if (info.inputs >= 1) condition.value = value1
      if (info.inputs === 2) condition.value2 = value2
      else if (info.combinable && operator2) {
        const second = filterOperatorInfo(operator2)
        condition.operator2 = operator2
        if (second && second.inputs >= 1) condition.value2 = value2
        condition.join = join
      }
      return { condition }
    }
    if (mode === 'color') {
      if (!colorChoice) return null
      return colorChoice.kind === 'fill' ? { fillColor: colorChoice.color || 'none' } : { fontColor: colorChoice.color || 'none' }
    }
    // Values. While searching only the checked matches stay visible, unless "Add current
    // selection to filter" keeps the earlier selection too (Excel).
    let selected: string[]
    let blanks: boolean
    if (query) {
      const matched = new Set(visibleItems.filter((item) => !searchExcluded.has(item.key)).map((item) => item.key))
      const searchBlanks = showBlanksRow && !searchExcluded.has(BLANKS_KEY)
      selected = values.items.filter((item) => matched.has(item.key) || (addToSelection && !excluded.has(item.key))).map((item) => item.text)
      blanks = Boolean(values.blanks) && (searchBlanks || (addToSelection && !excluded.has(BLANKS_KEY)))
    } else {
      selected = values.items.filter((item) => !excluded.has(item.key)).map((item) => item.text)
      blanks = Boolean(values.blanks) && !excluded.has(BLANKS_KEY)
    }
    if (selected.length === values.items.length && (!values.blanks || blanks)) return null
    return { values: selected, blanks }
  }

  const pending = buildCriteria()
  const nothingSelected = mode === 'values' && pending !== null && pending !== 'invalid' && !pending.values?.length && !pending.blanks

  const apply = () => {
    const result = buildCriteria()
    if (result === 'invalid' || nothingSelected) return
    onApply(result)
  }

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }
    if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement) && !(event.target instanceof HTMLSelectElement) && !(event.target instanceof HTMLInputElement && event.target.type === 'checkbox')) {
      event.preventDefault()
      event.stopPropagation()
      apply()
      return
    }
    if (event.key === 'ArrowDown' && event.target === searchRef.current) {
      event.preventDefault()
      listRef.current?.querySelector<HTMLInputElement>('input[type="checkbox"]')?.focus()
      return
    }
    if (trapTab(event, rootRef.current)) event.stopPropagation()
  }

  const firstRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN)
  const lastRow = Math.min(rowCount, Math.ceil((scrollTop + LIST_HEIGHT) / ROW_HEIGHT) + OVERSCAN)
  const renderRow = (index: number) => {
    const isBlanks = index === visibleItems.length
    const item: DistinctValue | null = isBlanks ? null : visibleItems[index]
    const key = item ? item.key : BLANKS_KEY
    const checked = isChecked(key)
    const text = item ? (item.text === '' ? ' ' : item.text) : '(Blanks)'
    const count = item ? item.count : values.blanks?.count || 0
    return (
      <label key={key} className={`dt-values-row${item ? '' : ' is-blanks'}`} style={{ top: index * ROW_HEIGHT }} title={item?.text}>
        <input
          type="checkbox"
          checked={checked}
          onChange={(event) => setChecked([key], event.target.checked)}
          onKeyDown={(event) => {
            if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return
            event.preventDefault()
            const rows = Array.from(listRef.current?.querySelectorAll<HTMLInputElement>('input[type="checkbox"]') || [])
            const position = rows.indexOf(event.currentTarget)
            const next = rows[position + (event.key === 'ArrowDown' ? 1 : -1)]
            if (next) next.focus()
            else if (event.key === 'ArrowUp') searchRef.current?.focus()
            else if (listRef.current) listRef.current.scrollTop += ROW_HEIGHT
          }}
        />
        <span className={`dt-values-text${item?.kind === 'number' || item?.kind === 'date' ? ' is-number' : ''}`}>{text}</span>
        <span className="dt-count">{count.toLocaleString()}</span>
      </label>
    )
  }

  const rows: ReactNode[] = []
  for (let index = firstRow; index < lastRow; index += 1) rows.push(renderRow(index))

  const colorRows = (kind: 'fill' | 'font', onPick: (color: string | null) => void, selected?: string | null | undefined) => (
    (kind === 'fill' ? values.fillColors : values.fontColors).map((entry) => {
      const isSelected = selected !== undefined && selected === entry.color
      return (
        <button key={`${kind}-${entry.color}`} type="button" className={`dt-color-row${isSelected ? ' is-selected' : ''}`} aria-pressed={selected === undefined ? undefined : isSelected} onClick={() => onPick(entry.color)}>
          {kind === 'fill' ? <Swatch color={entry.color} /> : <span className="dt-font-swatch" style={entry.color ? { color: entry.color } : undefined}>A</span>}
          <span>{colorLabel(entry.color, kind)}</span>
          <span className="dt-count">{entry.count.toLocaleString()}</span>
        </button>
      )
    })
  )

  return (
    <div className="dt-popover-layer" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose() }}>
      <div
        ref={rootRef}
        className="dt-filter-menu"
        role="dialog"
        aria-modal="true"
        aria-label={`Filter ${columnLabel}`}
        style={{ left: position.left, top: position.top, maxHeight: position.maxHeight }}
        onKeyDown={handleKeyDown}
      >
        <div className="dt-filter-scroll">
          <div className="dt-menu-group">
            <button type="button" className="dt-menu-item" onClick={() => onSort(false)}>
              <ArrowDownAZ size={14} aria-hidden="true" /><span>{ascendingLabel}</span>{sortDirection === 'asc' ? <Check size={13} className="dt-menu-check" aria-label="current" /> : null}
            </button>
            <button type="button" className="dt-menu-item" onClick={() => onSort(true)}>
              <ArrowUpZA size={14} aria-hidden="true" /><span>{descendingLabel}</span>{sortDirection === 'desc' ? <Check size={13} className="dt-menu-check" aria-label="current" /> : null}
            </button>
            {onSortByColor && hasColors ? (
              <Disclosure title="Sort by color" icon={<Palette size={13} />} open={open.sortColor} onToggle={() => setOpen((current) => ({ ...current, sortColor: !current.sortColor }))}>
                {hasFillColors ? <div className="dt-color-caption">Cell color</div> : null}
                {hasFillColors ? colorRows('fill', (color) => onSortByColor({ sortOn: 'fillColor', color })) : null}
                {hasFontColors ? <div className="dt-color-caption">Font color</div> : null}
                {hasFontColors ? colorRows('font', (color) => onSortByColor({ sortOn: 'fontColor', color })) : null}
              </Disclosure>
            ) : null}
          </div>

          <div className="dt-menu-group">
            <button type="button" className="dt-menu-item" disabled={!hasCriteria} onClick={() => onApply(null)}>
              <FilterX size={14} aria-hidden="true" /><span>Clear filter from “{columnLabel}”</span>
            </button>
          </div>

          <div className="dt-menu-group dt-menu-sections">
            {hasColors ? (
              <Disclosure title="Filter by color" active={mode === 'color' && Boolean(colorChoice)} open={open.color} onToggle={() => setOpen((current) => ({ ...current, color: !current.color }))}>
                {hasFillColors ? <div className="dt-color-caption">Filter by cell color</div> : null}
                {hasFillColors ? colorRows('fill', (color) => { setMode('color'); setColorChoice({ kind: 'fill', color }) }, colorChoice?.kind === 'fill' && mode === 'color' ? colorChoice.color : undefined) : null}
                {hasFontColors ? <div className="dt-color-caption">Filter by font color</div> : null}
                {hasFontColors ? colorRows('font', (color) => { setMode('color'); setColorChoice({ kind: 'font', color }) }, colorChoice?.kind === 'font' && mode === 'color' ? colorChoice.color : undefined) : null}
              </Disclosure>
            ) : null}

            <Disclosure title="Filter by condition" active={mode === 'condition' && Boolean(operator)} open={open.condition} onToggle={() => setOpen((current) => ({ ...current, condition: !current.condition }))}>
              <div className="dt-condition">
                <OperatorSelect label="Condition" value={operator} groups={groups} onChange={(next) => { setMode('condition'); setOperator(next); if (!filterOperatorInfo(next)?.combinable) setOperator2('') }} />
                {operatorInfo && operatorInfo.inputs >= 1 ? (
                  <input className="dt-input" aria-label="Condition value" list={operatorInfo.input === 'formula' ? undefined : datalistId} placeholder={inputPlaceholder(operator)} value={value1} onChange={(event) => { setMode('condition'); setValue1(event.target.value) }} />
                ) : null}
                {operatorInfo && operatorInfo.inputs === 2 ? (
                  <>
                    <span className="dt-and-label">and</span>
                    <input className="dt-input" aria-label="Second value" list={datalistId} placeholder={inputPlaceholder(operator, true)} value={value2} onChange={(event) => { setMode('condition'); setValue2(event.target.value) }} />
                  </>
                ) : null}
                {operatorInfo?.combinable ? (
                  <>
                    <div className="dt-join" role="radiogroup" aria-label="Combine conditions">
                      <label className="dt-check"><input type="radio" name="dt-join" checked={join === 'and'} onChange={() => { setMode('condition'); setJoin('and') }} />And</label>
                      <label className="dt-check"><input type="radio" name="dt-join" checked={join === 'or'} onChange={() => { setMode('condition'); setJoin('or') }} />Or</label>
                    </div>
                    <OperatorSelect label="Second condition" combinableOnly value={operator2} groups={groups} onChange={(next) => { setMode('condition'); setOperator2(next) }} />
                    {secondInfo && secondInfo.inputs >= 1 ? (
                      <input className="dt-input" aria-label="Second condition value" list={datalistId} placeholder={inputPlaceholder(operator2)} value={value2} onChange={(event) => { setMode('condition'); setValue2(event.target.value) }} />
                    ) : null}
                  </>
                ) : null}
                {operatorInfo?.input === 'text' || secondInfo?.input === 'text' ? <p className="dt-hint">Use ? for any single character and * for any run of characters.</p> : null}
                <datalist id={datalistId}>{suggestions.map((text) => <option key={text} value={text} />)}</datalist>
              </div>
            </Disclosure>

            <Disclosure title="Filter by values" active={mode === 'values' && pending !== null && pending !== 'invalid'} open={open.values} onToggle={() => setOpen((current) => ({ ...current, values: !current.values }))}>
              <div className="dt-values">
                <div className="dt-search">
                  <Search size={13} aria-hidden="true" />
                  <input ref={searchRef} type="text" aria-label="Search values" placeholder="Search" value={search} onChange={(event) => { setSearch(event.target.value); setSearchExcluded(new Set()); setMode('values'); setScrollTop(0); if (listRef.current) listRef.current.scrollTop = 0 }} />
                  {search ? <button type="button" className="dt-search-clear" aria-label="Clear search" onClick={() => { setSearch(''); searchRef.current?.focus() }}><X size={12} /></button> : null}
                </div>
                <div className="dt-values-toolbar">
                  <label className="dt-check dt-select-all">
                    <input
                      type="checkbox"
                      checked={allState === 'all'}
                      ref={(element) => { if (element) element.indeterminate = allState === 'some' }}
                      onChange={(event) => setChecked([...visibleItems.map((item) => item.key), ...(showBlanksRow ? [BLANKS_KEY] : [])], event.target.checked)}
                    />
                    {query ? '(Select all search results)' : '(Select all)'}
                  </label>
                  <span className="dt-count">{checkedVisible.toLocaleString()} / {rowCount.toLocaleString()}</span>
                </div>
                {query ? (
                  <label className="dt-check dt-add-selection"><input type="checkbox" checked={addToSelection} onChange={(event) => setAddToSelection(event.target.checked)} />Add current selection to filter</label>
                ) : null}
                <div ref={listRef} className="dt-values-list" role="group" aria-label={`Values in ${columnLabel}`} style={{ height: Math.min(LIST_HEIGHT, Math.max(rowCount, 1) * ROW_HEIGHT + 2) }} onScroll={(event) => setScrollTop(event.currentTarget.scrollTop)}>
                  <div className="dt-values-spacer" style={{ height: rowCount * ROW_HEIGHT }}>{rows}</div>
                  {rowCount === 0 ? <div className="dt-values-empty">{query ? 'No matches' : 'No values'}</div> : null}
                </div>
              </div>
            </Disclosure>
          </div>
        </div>
        <div className="dt-filter-footer">
          {nothingSelected ? <span className="dt-footer-note">Select at least one value</span> : <span className="dt-footer-note">{values.rows.toLocaleString()} rows</span>}
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" disabled={pending === 'invalid' || nothingSelected} onClick={apply}>OK</button>
        </div>
      </div>
    </div>
  )
}
