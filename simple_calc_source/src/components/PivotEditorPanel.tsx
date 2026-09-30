import { useEffect, useMemo, useState } from 'react'
import { ArrowDown, ArrowUp, RefreshCw, Rows3, Table, X } from 'lucide-react'
import { PIVOT_DATE_GROUPS, PIVOT_SUMMARIES, pivotFieldIsDate, pivotFieldKeys, valueFieldLabel } from '../lib/pivot'
import type { PivotSource } from '../lib/pivot'
import type { PivotAxisField, PivotDateGroup, PivotFilterField, PivotShowAs, PivotSummarize, PivotTableModel, PivotValueField } from '../spreadsheet-types'
import { DataToolsDialog } from './DataToolsDialogFrame'
import './charts.css'
import './table-tools.css'

type Area = 'rows' | 'columns' | 'values' | 'filters'

const SHOW_AS: Array<{ id: PivotShowAs; label: string }> = [
  { id: 'normal', label: 'Default' },
  { id: 'percentOfGrandTotal', label: '% of grand total' },
  { id: 'percentOfRowTotal', label: '% of row total' },
  { id: 'percentOfColumnTotal', label: '% of column total' },
  { id: 'runningTotal', label: 'Running total' },
]

const AREA_LABEL: Record<Area, string> = { filters: 'Filters', columns: 'Columns', rows: 'Rows', values: 'Values' }

export interface PivotEditorPanelProps {
  pivot: PivotTableModel
  source: PivotSource | null
  sourceError?: string | null
  onChange: (pivot: PivotTableModel) => void
  onRefresh: () => void
  onDelete: () => void
  onClose: () => void
}

function moveItem<T>(list: T[], index: number, delta: number) {
  const target = index + delta
  if (target < 0 || target >= list.length) return list
  const next = list.slice()
  const [item] = next.splice(index, 1)
  next.splice(target, 0, item)
  return next
}

/** Google Sheets-style pivot editor: pick fields into Rows, Columns, Values and Filters. */
export function PivotEditorPanel({ pivot, source, sourceError, onChange, onRefresh, onDelete, onClose }: PivotEditorPanelProps) {
  const [sourceText, setSourceText] = useState(pivot.source)
  const [openFilter, setOpenFilter] = useState<string | null>(null)
  const [filterSearch, setFilterSearch] = useState('')
  useEffect(() => { setSourceText(pivot.source) }, [pivot.source])
  const headers = source?.headers || []
  const used = useMemo(() => new Set([...pivot.rows, ...pivot.columns, ...pivot.values, ...pivot.filters].map((item) => item.field.toLocaleLowerCase())), [pivot])
  const isNumeric = (field: string) => {
    const index = headers.indexOf(field)
    if (index < 0 || !source) return false
    let numbers = 0
    let others = 0
    for (const record of source.records.slice(0, 200)) {
      const value = record[index]?.value
      if (value === null || value === undefined || value === '') continue
      if (typeof value === 'number' && !record[index]?.isDate) numbers += 1
      else others += 1
    }
    return numbers > 0 && numbers >= others
  }
  const update = (patch: Partial<PivotTableModel>) => onChange({ ...pivot, ...patch })
  const addField = (area: Area, field: string) => {
    if (area === 'values') update({ values: [...pivot.values, { field, summarize: isNumeric(field) ? 'sum' : 'count' }] })
    else if (area === 'filters') update({ filters: [...pivot.filters, { field }] })
    else {
      const entry: PivotAxisField = { field, ...(source && pivotFieldIsDate(source, field) ? { dateGroup: 'yearMonth' as PivotDateGroup } : {}) }
      update({ [area]: [...pivot[area].filter((item) => item.field !== field), entry] } as Partial<PivotTableModel>)
    }
  }
  const toggleField = (field: string, checked: boolean) => {
    if (checked) addField(isNumeric(field) ? 'values' : 'rows', field)
    else {
      const keep = <T extends { field: string }>(list: T[]) => list.filter((item) => item.field.toLocaleLowerCase() !== field.toLocaleLowerCase())
      update({ rows: keep(pivot.rows), columns: keep(pivot.columns), values: keep(pivot.values), filters: keep(pivot.filters) })
    }
  }
  const moveToArea = (from: Area, index: number, to: Area) => {
    const item = pivot[from][index]
    if (!item) return
    const removed = { [from]: pivot[from].filter((_, position) => position !== index) } as Partial<PivotTableModel>
    const nextModel = { ...pivot, ...removed }
    const field = item.field
    if (to === 'values') onChange({ ...nextModel, values: [...nextModel.values, { field, summarize: isNumeric(field) ? 'sum' : 'count' }] })
    else if (to === 'filters') onChange({ ...nextModel, filters: [...nextModel.filters, { field }] })
    else onChange({ ...nextModel, [to]: [...nextModel[to], { field }] } as PivotTableModel)
  }

  const addPicker = (area: Area) => (
    <select
      className="pivot-add"
      aria-label={`Add field to ${AREA_LABEL[area]}`}
      value=""
      onChange={(event) => { if (event.target.value) addField(area, event.target.value) }}
    >
      <option value="">Add…</option>
      {headers.map((header) => <option key={header} value={header}>{header}</option>)}
    </select>
  )

  const areaSelect = (area: Area, index: number) => (
    <select className="pivot-move" aria-label="Move to area" value={area} onChange={(event) => moveToArea(area, index, event.target.value as Area)}>
      {(['rows', 'columns', 'values', 'filters'] as Area[]).map((option) => <option key={option} value={option}>{AREA_LABEL[option]}</option>)}
    </select>
  )

  const axisCard = (area: 'rows' | 'columns', item: PivotAxisField, index: number) => {
    const list = pivot[area]
    const set = (patch: Partial<PivotAxisField>) => update({ [area]: list.map((entry, position) => (position === index ? { ...entry, ...patch } : entry)) } as Partial<PivotTableModel>)
    const dates = source ? pivotFieldIsDate(source, item.field) : false
    return (
      <li key={`${area}-${index}-${item.field}`} className="pivot-card" data-pivot-card={`${area}-${item.field}`}>
        <div className="pivot-card-head">
          <strong title={item.field}>{item.field}</strong>
          <button type="button" className="chart-icon-button" aria-label="Move up" title="Move up" disabled={index === 0} onClick={() => update({ [area]: moveItem(list, index, -1) } as Partial<PivotTableModel>)}><ArrowUp size={13} /></button>
          <button type="button" className="chart-icon-button" aria-label="Move down" title="Move down" disabled={index === list.length - 1} onClick={() => update({ [area]: moveItem(list, index, 1) } as Partial<PivotTableModel>)}><ArrowDown size={13} /></button>
          <button type="button" className="chart-icon-button" aria-label={`Remove ${item.field}`} title="Remove" onClick={() => update({ [area]: list.filter((_, position) => position !== index) } as Partial<PivotTableModel>)}><X size={13} /></button>
        </div>
        <div className="pivot-card-row">
          <label className="chart-field chart-field-compact">
            <span>Order</span>
            <select value={item.order || 'asc'} onChange={(event) => set({ order: event.target.value as 'asc' | 'desc' })}>
              <option value="asc">Ascending</option>
              <option value="desc">Descending</option>
            </select>
          </label>
          {dates && (
            <label className="chart-field chart-field-compact">
              <span>Group by</span>
              <select value={item.dateGroup || ''} onChange={(event) => set({ dateGroup: (event.target.value || undefined) as PivotDateGroup | undefined })}>
                <option value="">Dates</option>
                {PIVOT_DATE_GROUPS.map((group) => <option key={group.id} value={group.id}>{group.label}</option>)}
              </select>
            </label>
          )}
        </div>
        <div className="pivot-card-row">
          {index < list.length - 1 && (
            <label className="chart-toggle"><input type="checkbox" checked={item.showTotals !== false} onChange={(event) => set({ showTotals: event.target.checked })} /><span>Show totals</span></label>
          )}
          <span className="pivot-spacer" />
          {areaSelect(area, index)}
        </div>
      </li>
    )
  }

  const valueCard = (item: PivotValueField, index: number) => {
    const set = (patch: Partial<PivotValueField>) => update({ values: pivot.values.map((entry, position) => (position === index ? { ...entry, ...patch } : entry)) })
    return (
      <li key={`values-${index}-${item.field}`} className="pivot-card" data-pivot-card={`values-${item.field}`}>
        <div className="pivot-card-head">
          <strong title={valueFieldLabel(item)}>{valueFieldLabel(item)}</strong>
          <button type="button" className="chart-icon-button" aria-label="Move up" title="Move up" disabled={index === 0} onClick={() => update({ values: moveItem(pivot.values, index, -1) })}><ArrowUp size={13} /></button>
          <button type="button" className="chart-icon-button" aria-label="Move down" title="Move down" disabled={index === pivot.values.length - 1} onClick={() => update({ values: moveItem(pivot.values, index, 1) })}><ArrowDown size={13} /></button>
          <button type="button" className="chart-icon-button" aria-label={`Remove ${item.field}`} title="Remove" onClick={() => update({ values: pivot.values.filter((_, position) => position !== index) })}><X size={13} /></button>
        </div>
        <div className="pivot-card-row">
          <label className="chart-field chart-field-compact">
            <span>Summarize by</span>
            <select aria-label={`Summarize ${item.field} by`} value={item.summarize} onChange={(event) => set({ summarize: event.target.value as PivotSummarize, label: undefined })}>
              {PIVOT_SUMMARIES.map((summary) => <option key={summary.id} value={summary.id}>{summary.label}</option>)}
            </select>
          </label>
          <label className="chart-field chart-field-compact">
            <span>Show as</span>
            <select aria-label={`Show ${item.field} as`} value={item.showAs || 'normal'} onChange={(event) => set({ showAs: event.target.value as PivotShowAs })}>
              {SHOW_AS.map((option) => <option key={option.id} value={option.id}>{option.label}</option>)}
            </select>
          </label>
        </div>
        <div className="pivot-card-row"><span className="pivot-spacer" />{areaSelect('values', index)}</div>
      </li>
    )
  }

  const filterCard = (item: PivotFilterField, index: number) => {
    const keys = source ? pivotFieldKeys(source, item.field) : []
    const excluded = new Set(item.exclude || [])
    const setExcluded = (next: Set<string>) => update({ filters: pivot.filters.map((entry, position) => (position === index ? { ...entry, exclude: [...next] } : entry)) })
    const open = openFilter === `${index}`
    const visible = keys.filter((key) => !filterSearch || key.label.toLocaleLowerCase().includes(filterSearch.toLocaleLowerCase()))
    return (
      <li key={`filters-${index}-${item.field}`} className="pivot-card" data-pivot-card={`filters-${item.field}`}>
        <div className="pivot-card-head">
          <strong title={item.field}>{item.field}</strong>
          <span className="pivot-filter-status">{excluded.size ? `${keys.length - excluded.size} of ${keys.length}` : 'All'}</span>
          <button type="button" className="chart-icon-button" aria-label={`Remove ${item.field}`} title="Remove" onClick={() => update({ filters: pivot.filters.filter((_, position) => position !== index) })}><X size={13} /></button>
        </div>
        <div className="pivot-card-row">
          <button type="button" className="pivot-link" aria-expanded={open} onClick={() => { setOpenFilter(open ? null : `${index}`); setFilterSearch('') }}>{open ? 'Hide values' : 'Choose values…'}</button>
          <span className="pivot-spacer" />
          {areaSelect('filters', index)}
        </div>
        {open && (
          <div className="pivot-filter-list">
            <input className="pivot-filter-search" placeholder="Search" aria-label={`Search ${item.field} values`} value={filterSearch} onChange={(event) => setFilterSearch(event.target.value)} />
            <div className="pivot-filter-actions">
              <button type="button" className="pivot-link" onClick={() => setExcluded(new Set())}>Select all</button>
              <button type="button" className="pivot-link" onClick={() => setExcluded(new Set(keys.map((key) => key.id)))}>Clear</button>
            </div>
            <ul>
              {visible.map((key) => (
                <li key={key.id}>
                  <label className="chart-toggle">
                    <input type="checkbox" checked={!excluded.has(key.id)} onChange={(event) => { const next = new Set(excluded); if (event.target.checked) next.delete(key.id); else next.add(key.id); setExcluded(next) }} />
                    <span>{key.label}</span>
                    <small>{key.count}</small>
                  </label>
                </li>
              ))}
            </ul>
          </div>
        )}
      </li>
    )
  }

  const commitSource = () => {
    const text = sourceText.trim().replace(/^=/, '')
    if (text && text !== pivot.source) update({ source: text })
  }

  return (
    <aside className="chart-editor pivot-editor" aria-label="Pivot table editor" data-chart-keep-selection>
      <header className="chart-editor-header">
        <span className="chart-editor-icon"><Table size={16} aria-hidden="true" /></span>
        <h2>{pivot.name}</h2>
        <button type="button" className="chart-icon-button" onClick={onClose} aria-label="Close pivot table editor" title="Close"><X size={16} /></button>
      </header>
      <div className="chart-editor-body">
        <section className="chart-editor-section">
          <label className="chart-field">
            <span>Data range</span>
            <input
              value={sourceText}
              spellCheck={false}
              aria-label="Pivot data range"
              className={sourceError ? 'has-error' : ''}
              onChange={(event) => setSourceText(event.target.value)}
              onBlur={commitSource}
              onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); commitSource() } }}
            />
          </label>
          {sourceError && <p className="chart-error" role="alert">{sourceError}</p>}
          <div className="pivot-field-list" role="group" aria-label="Fields">
            {headers.map((header) => (
              <label key={header} className="chart-toggle">
                <input type="checkbox" checked={used.has(header.toLocaleLowerCase())} onChange={(event) => toggleField(header, event.target.checked)} />
                <span>{header}</span>
              </label>
            ))}
          </div>
        </section>
        {(['rows', 'columns', 'values', 'filters'] as Area[]).map((area) => (
          <section key={area} className="chart-editor-section" data-pivot-area={area}>
            <div className="pivot-area-head">
              <h3>{AREA_LABEL[area]}</h3>
              {addPicker(area)}
            </div>
            <ul className="pivot-cards">
              {area === 'values' ? pivot.values.map(valueCard)
                : area === 'filters' ? pivot.filters.map(filterCard)
                  : pivot[area].map((item, index) => axisCard(area, item, index))}
            </ul>
          </section>
        ))}
        <section className="chart-editor-section">
          <h3>Totals</h3>
          <label className="chart-toggle"><input type="checkbox" checked={pivot.showRowGrandTotal !== false} onChange={(event) => update({ showRowGrandTotal: event.target.checked })} /><span>Grand total row</span></label>
          <label className="chart-toggle"><input type="checkbox" checked={pivot.showColumnGrandTotal !== false} onChange={(event) => update({ showColumnGrandTotal: event.target.checked })} /><span>Grand total column</span></label>
        </section>
      </div>
      <footer className="chart-editor-footer">
        <button type="button" className="chart-primary-button" onClick={onRefresh}><RefreshCw size={13} aria-hidden="true" /> Refresh</button>
        <button type="button" className="chart-danger-button" onClick={onDelete}><Rows3 size={13} aria-hidden="true" /> Delete pivot table</button>
      </footer>
    </aside>
  )
}

export interface CreatePivotDialogProps {
  initialRange: string
  initialLocation: string
  onCreate: (range: string, destination: 'new' | string) => string | null
  onClose: () => void
}

/** Choose the data and where the pivot table goes (a new sheet by default, as in Excel). */
export function CreatePivotDialog({ initialRange, initialLocation, onCreate, onClose }: CreatePivotDialogProps) {
  const [range, setRange] = useState(initialRange)
  const [mode, setMode] = useState<'new' | 'existing'>('new')
  const [location, setLocation] = useState(initialLocation)
  const [error, setError] = useState<string | null>(null)
  const create = () => setError(onCreate(range.trim().replace(/^=/, ''), mode === 'new' ? 'new' : location.trim().replace(/^=/, '')))
  return (
    <DataToolsDialog
      title="Create pivot table"
      subtitle="Summarize a range by any of its columns."
      icon={<Table size={16} />}
      width={420}
      className="create-pivot-dialog"
      onClose={onClose}
      onConfirm={create}
      footer={(
        <>
          <span className="dt-footer-start" />
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" onClick={create}>Create</button>
        </>
      )}
    >
      <label className="dt-field">
        <span>Data range</span>
        <input className="dt-input dt-mono" aria-label="Pivot data range" value={range} spellCheck={false} data-autofocus onChange={(event) => { setRange(event.target.value); setError(null) }} />
      </label>
      <fieldset className="dt-fieldset" role="radiogroup" aria-label="Insert to">
        <label className="dt-check"><input type="radio" name="pivot-destination" checked={mode === 'new'} onChange={() => setMode('new')} />New sheet</label>
        <label className="dt-check"><input type="radio" name="pivot-destination" checked={mode === 'existing'} onChange={() => setMode('existing')} />Existing sheet</label>
      </fieldset>
      {mode === 'existing' && (
        <label className="dt-field">
          <span>Location</span>
          <input className="dt-input dt-mono" aria-label="Pivot location" value={location} spellCheck={false} onChange={(event) => { setLocation(event.target.value); setError(null) }} />
        </label>
      )}
      {error && <p className="dt-error" role="alert">{error}</p>}
    </DataToolsDialog>
  )
}
