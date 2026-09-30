import { useMemo, useRef, useState } from 'react'
import { ArrowDown, ArrowDownUp, ArrowUp, Copy, Plus, Settings2, Trash2 } from 'lucide-react'
import { BUILTIN_CUSTOM_LISTS } from '../lib/sort'
import type { SortLevel, SortOn, SortOrientation } from '../lib/sort'
import { DataToolsDialog, Swatch, colorLabel } from './DataToolsDialogFrame'
import './data-tools.css'

export type SortKeyKind = 'text' | 'number' | 'date' | 'mixed'

export interface SortDialogResult {
  levels: SortLevel[]
  hasHeader: boolean
  orientation: SortOrientation
  caseSensitive: boolean
}

export interface SortDialogProps {
  /** e.g. "A1:F200". */
  rangeLabel: string
  initialHasHeader: boolean
  initialLevels?: SortLevel[]
  initialOrientation?: SortOrientation
  initialCaseSensitive?: boolean
  /** Names for the "Sort by" list: sortKeyLabels(bounds, orientation, hasHeader, host). */
  labelsFor: (orientation: SortOrientation, hasHeader: boolean) => string[]
  /** Colours for colour sorts: sortKeyColors(bounds, orientation, key, sortOn, hasHeader, host). */
  colorsFor?: (orientation: SortOrientation, key: number, sortOn: Exclude<SortOn, 'values'>, hasHeader: boolean) => Array<string | null>
  /** Data type of a key, to word the order list like Excel (A to Z / Smallest to Largest / Oldest to Newest). */
  kindFor?: (orientation: SortOrientation, key: number, hasHeader: boolean) => SortKeyKind
  customLists?: Array<{ id: string; label: string; items: string[] }>
  onSort: (result: SortDialogResult) => void
  onClose: () => void
}

const MAX_LEVELS = 64

interface LevelDraft extends SortLevel {
  id: number
  /** Selected custom list id, 'custom' for a typed list. */
  listId?: string
  listText?: string
}

let nextLevelId = 1

function draftFrom(level: SortLevel, lists: Array<{ id: string; items: string[] }>): LevelDraft {
  const draft: LevelDraft = { ...level, id: nextLevelId++ }
  if (level.customList?.length) {
    const match = lists.find((list) => list.items.join('\u0000').toLocaleLowerCase() === level.customList!.join('\u0000').toLocaleLowerCase())
    draft.listId = match ? match.id : 'custom'
    draft.listText = level.customList.join(', ')
  }
  return draft
}

function orderLabels(kind: SortKeyKind | undefined): [string, string] {
  if (kind === 'number') return ['Smallest to Largest', 'Largest to Smallest']
  if (kind === 'date') return ['Oldest to Newest', 'Newest to Oldest']
  return ['A to Z', 'Z to A']
}

/** Excel's Sort dialog: levels (add / delete / copy / reorder), header toggle, sort on values or colours, custom lists, options. */
export function SortDialog({
  rangeLabel,
  initialHasHeader,
  initialLevels,
  initialOrientation = 'rows',
  initialCaseSensitive = false,
  labelsFor,
  colorsFor,
  kindFor,
  customLists = BUILTIN_CUSTOM_LISTS,
  onSort,
  onClose,
}: SortDialogProps) {
  const [hasHeader, setHasHeader] = useState(initialHasHeader)
  const [orientation, setOrientation] = useState<SortOrientation>(initialOrientation)
  const [caseSensitive, setCaseSensitive] = useState(initialCaseSensitive)
  const [levels, setLevels] = useState<LevelDraft[]>(() => (initialLevels?.length ? initialLevels : [{ key: 0 }]).map((level) => draftFrom(level, customLists)))
  const [selected, setSelected] = useState(0)
  const [showOptions, setShowOptions] = useState(initialCaseSensitive || initialOrientation === 'columns')
  const [error, setError] = useState('')
  const firstSelectRef = useRef<HTMLSelectElement>(null)

  const labels = useMemo(() => labelsFor(orientation, hasHeader), [hasHeader, labelsFor, orientation])

  const update = (index: number, patch: Partial<LevelDraft>) => {
    setError('')
    setLevels((current) => current.map((level, position) => position === index ? { ...level, ...patch } : level))
  }

  const addLevel = () => {
    if (levels.length >= MAX_LEVELS) return
    const used = new Set(levels.map((level) => level.key))
    const key = labels.findIndex((_, index) => !used.has(index))
    setLevels((current) => [...current, { id: nextLevelId++, key: key < 0 ? 0 : key }])
    setSelected(levels.length)
  }
  const deleteLevel = () => {
    if (!levels.length) return
    setLevels((current) => current.filter((_, index) => index !== selected))
    setSelected((current) => Math.max(0, Math.min(current, levels.length - 2)))
  }
  const copyLevel = () => {
    if (!levels[selected] || levels.length >= MAX_LEVELS) return
    setLevels((current) => [...current.slice(0, selected + 1), { ...current[selected], id: nextLevelId++ }, ...current.slice(selected + 1)])
    setSelected(selected + 1)
  }
  const moveLevel = (delta: -1 | 1) => {
    const target = selected + delta
    if (target < 0 || target >= levels.length) return
    setLevels((current) => {
      const next = [...current]
      ;[next[selected], next[target]] = [next[target], next[selected]]
      return next
    })
    setSelected(target)
  }

  const submit = () => {
    if (!levels.length) { setError('Add at least one sort level.'); return }
    const result: SortLevel[] = []
    for (const [index, level] of levels.entries()) {
      if (level.key < 0 || level.key >= labels.length) { setError(`Level ${index + 1}: choose a ${orientation === 'rows' ? 'column' : 'row'}.`); return }
      const sortOn = level.sortOn || 'values'
      const clean: SortLevel = { key: level.key }
      if (sortOn !== 'values') {
        clean.sortOn = sortOn
        clean.color = level.color ?? null
        clean.position = level.position || 'top'
      } else {
        if (level.descending) clean.descending = true
        if (level.listId) {
          const items = level.listId === 'custom'
            ? (level.listText || '').split(/[,\n]/).map((item) => item.trim()).filter(Boolean)
            : customLists.find((list) => list.id === level.listId)?.items || []
          if (!items.length) { setError(`Level ${index + 1}: enter the custom list items, separated by commas.`); return }
          clean.customList = items
        }
      }
      result.push(clean)
    }
    const seen = new Set<string>()
    for (const level of result) {
      const signature = `${level.key}:${level.sortOn || 'values'}:${level.color ?? ''}`
      if (seen.has(signature)) {
        setError(`"${labels[level.key]}" is being sorted by ${level.sortOn === 'fillColor' ? 'the same cell color' : level.sortOn === 'fontColor' ? 'the same font color' : 'values'} more than once. Delete the duplicate level.`)
        return
      }
      seen.add(signature)
    }
    onSort({ levels: result, hasHeader, orientation, caseSensitive })
  }

  const keyNoun = orientation === 'rows' ? 'Column' : 'Row'

  return (
    <DataToolsDialog
      title="Sort"
      subtitle={<>Range <strong>{rangeLabel}</strong>{hasHeader ? ` · first ${orientation === 'rows' ? 'row' : 'column'} is a header` : ''}</>}
      icon={<ArrowDownUp size={17} />}
      width={680}
      className="dt-sort-dialog"
      onClose={onClose}
      onConfirm={submit}
      initialFocus={firstSelectRef}
      footer={(
        <>
          <span className="dt-footer-start dt-footer-note">{levels.length} {levels.length === 1 ? 'level' : 'levels'} · blanks always sort last</span>
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" onClick={submit}>Sort</button>
        </>
      )}
    >
      <div className="dt-toolbar">
        <button type="button" className="dt-button is-compact" disabled={levels.length >= MAX_LEVELS} onClick={addLevel}><Plus size={13} />Add level</button>
        <button type="button" className="dt-button is-compact" disabled={!levels.length} onClick={deleteLevel}><Trash2 size={13} />Delete level</button>
        <button type="button" className="dt-button is-compact" disabled={!levels.length || levels.length >= MAX_LEVELS} onClick={copyLevel}><Copy size={13} />Copy level</button>
        <span className="dt-toolbar-divider" />
        <button type="button" className="dt-icon-button" aria-label="Move level up" title="Move up" disabled={selected <= 0} onClick={() => moveLevel(-1)}><ArrowUp size={14} /></button>
        <button type="button" className="dt-icon-button" aria-label="Move level down" title="Move down" disabled={selected >= levels.length - 1} onClick={() => moveLevel(1)}><ArrowDown size={14} /></button>
        <span className="dt-toolbar-spacer" />
        <button type="button" className={`dt-button is-compact${showOptions ? ' is-pressed' : ''}`} aria-expanded={showOptions} onClick={() => setShowOptions((value) => !value)}><Settings2 size={13} />Options</button>
        <label className="dt-check"><input type="checkbox" checked={hasHeader} onChange={(event) => setHasHeader(event.target.checked)} />My data has headers</label>
      </div>

      {showOptions ? (
        <div className="dt-options-row">
          <label className="dt-check"><input type="checkbox" checked={caseSensitive} onChange={(event) => setCaseSensitive(event.target.checked)} />Case sensitive</label>
          <div className="dt-segmented" role="radiogroup" aria-label="Orientation">
            <button type="button" role="radio" aria-checked={orientation === 'rows'} className={orientation === 'rows' ? 'is-active' : ''} onClick={() => { setOrientation('rows'); setLevels((current) => current.map((level) => ({ ...level, key: 0 }))) }}>Sort top to bottom</button>
            <button type="button" role="radio" aria-checked={orientation === 'columns'} className={orientation === 'columns' ? 'is-active' : ''} onClick={() => { setOrientation('columns'); setLevels((current) => current.map((level) => ({ ...level, key: 0 }))) }}>Sort left to right</button>
          </div>
        </div>
      ) : null}

      <div className="dt-levels" role="grid" aria-label="Sort levels">
        <div className="dt-levels-head" role="row">
          <span role="columnheader" />
          <span role="columnheader">{keyNoun}</span>
          <span role="columnheader">Sort on</span>
          <span role="columnheader">Order</span>
        </div>
        {levels.map((level, index) => {
          const sortOn = level.sortOn || 'values'
          const kind = kindFor?.(orientation, level.key, hasHeader)
          const [ascending, descending] = orderLabels(kind)
          const available = sortOn !== 'values' && colorsFor ? colorsFor(orientation, level.key, sortOn, hasHeader) : []
          const colors = level.color !== undefined && !available.includes(level.color) ? [...available, level.color] : available
          const orderValue = level.listId ? `list:${level.listId}` : level.descending ? 'desc' : 'asc'
          return (
            <div
              key={level.id}
              role="row"
              aria-selected={index === selected}
              className={`dt-level${index === selected ? ' is-selected' : ''}`}
              onMouseDown={() => setSelected(index)}
              onFocus={() => setSelected(index)}
            >
              <span className="dt-level-label" role="rowheader">{index === 0 ? 'Sort by' : 'Then by'}</span>
              <select ref={index === 0 ? firstSelectRef : undefined} className="dt-select" aria-label={`${index === 0 ? 'Sort by' : 'Then by'} ${keyNoun.toLocaleLowerCase()}`} value={level.key} onChange={(event) => update(index, { key: Number(event.target.value), color: undefined })}>
                {labels.map((label, key) => <option key={key} value={key}>{label}</option>)}
              </select>
              <select className="dt-select" aria-label="Sort on" value={sortOn} onChange={(event) => {
                const next = event.target.value as SortOn
                const available = next !== 'values' && colorsFor ? colorsFor(orientation, level.key, next, hasHeader) : []
                update(index, { sortOn: next, color: next === 'values' ? undefined : available[0] ?? null, position: 'top', listId: undefined })
              }}>
                <option value="values">Cell values</option>
                <option value="fillColor">Cell color</option>
                <option value="fontColor">Font color</option>
              </select>
              {sortOn === 'values' ? (
                <div className="dt-order-cell">
                  <select className="dt-select" aria-label="Order" value={orderValue} onChange={(event) => {
                    const value = event.target.value
                    if (value === 'asc' || value === 'desc') update(index, { descending: value === 'desc', listId: undefined })
                    else update(index, { listId: value.slice(5), descending: false })
                  }}>
                    <option value="asc">{ascending}</option>
                    <option value="desc">{descending}</option>
                    <optgroup label="Custom list">
                      {customLists.map((list) => <option key={list.id} value={`list:${list.id}`}>{list.label}</option>)}
                      <option value="list:custom">Custom…</option>
                    </optgroup>
                  </select>
                  {level.listId ? (
                    <div className="dt-order-extra">
                      {level.listId === 'custom' ? (
                        <input className="dt-input" aria-label="Custom list items" placeholder="High, Medium, Low" value={level.listText || ''} onChange={(event) => update(index, { listText: event.target.value })} />
                      ) : null}
                      <label className="dt-check"><input type="checkbox" checked={Boolean(level.descending)} onChange={(event) => update(index, { descending: event.target.checked })} />Reverse</label>
                    </div>
                  ) : null}
                </div>
              ) : (
                <div className="dt-order-cell dt-color-order">
                  <div className="dt-swatch-picker" role="radiogroup" aria-label={sortOn === 'fillColor' ? 'Cell color' : 'Font color'}>
                    {(colors.length ? colors : [level.color ?? null]).map((color) => (
                      <button
                        key={String(color)}
                        type="button"
                        role="radio"
                        aria-checked={(level.color ?? null) === color}
                        title={colorLabel(color, sortOn === 'fillColor' ? 'fill' : 'font')}
                        className={`dt-swatch-button${(level.color ?? null) === color ? ' is-selected' : ''}`}
                        onClick={() => update(index, { color })}
                      >
                        {sortOn === 'fillColor' ? <Swatch color={color} /> : <span className="dt-font-swatch" style={color ? { color } : undefined}>A</span>}
                      </button>
                    ))}
                  </div>
                  <select className="dt-select" aria-label="Color position" value={level.position || 'top'} onChange={(event) => update(index, { position: event.target.value as 'top' | 'bottom' })}>
                    <option value="top">On {orientation === 'rows' ? 'top' : 'left'}</option>
                    <option value="bottom">On {orientation === 'rows' ? 'bottom' : 'right'}</option>
                  </select>
                </div>
              )}
            </div>
          )
        })}
        {!levels.length ? <div className="dt-empty-levels">No sort levels. Choose <strong>Add level</strong> to sort.</div> : null}
      </div>
      {error ? <div className="dt-error" role="alert">{error}</div> : null}
    </DataToolsDialog>
  )
}
