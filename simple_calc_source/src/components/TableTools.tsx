import { memo, useEffect, useMemo, useState } from 'react'
import { Table2, X } from 'lucide-react'
import { TABLE_STYLE_GALLERY, tableCellPaint } from '../lib/table-styles'
import { TOTAL_FUNCTIONS } from '../lib/tables'
import type { TotalFunctionId } from '../lib/tables'
import type { SheetTable } from '../spreadsheet-types'
import { DataToolsDialog } from './DataToolsDialogFrame'
import './charts.css'
import './table-tools.css'

const PREVIEW_COLUMNS = 5
const PREVIEW_ROWS = 6

function strokeOf(css: string | undefined) {
  const match = /^(\d+)px\s+\w+\s+(#[0-9a-f]{3,8})$/i.exec(css || '')
  return match ? { width: Number(match[1]) >= 2 ? 1.6 : 0.8, color: match[2] } : null
}

/** A miniature of a built-in table style, painted with the same rules as the grid. */
export const TableStylePreview = memo(function TableStylePreview({ styleName, theme, showRowStripes = true, showColumnStripes = false, showFirstColumn = false, showLastColumn = false }: {
  styleName: string
  theme?: readonly string[]
  showRowStripes?: boolean
  showColumnStripes?: boolean
  showFirstColumn?: boolean
  showLastColumn?: boolean
}) {
  const cellWidth = 12
  const cellHeight = 7
  const table: SheetTable = {
    id: 'preview',
    name: 'Preview',
    ref: 'A1:E6',
    headerRow: true,
    totalsRow: false,
    columns: [],
    style: { theme: styleName, showRowStripes, showColumnStripes, showFirstColumn, showLastColumn },
  }
  const entry = { table, top: 0, bottom: PREVIEW_ROWS - 1, left: 0, right: PREVIEW_COLUMNS - 1 }
  const parts: JSX.Element[] = []
  for (let row = 0; row < PREVIEW_ROWS; row += 1) {
    for (let col = 0; col < PREVIEW_COLUMNS; col += 1) {
      const paint = tableCellPaint(entry, row, col, theme) || {}
      const x = col * cellWidth
      const y = row * cellHeight
      parts.push(<rect key={`f${row}-${col}`} x={x} y={y} width={cellWidth} height={cellHeight} fill={paint.fill || '#ffffff'} />)
      const ink = paint.color || '#3a3a3a'
      parts.push(<rect key={`t${row}-${col}`} x={x + 3} y={y + 3} width={cellWidth - 6} height={1.2} fill={ink} opacity={paint.bold ? 0.9 : 0.35} />)
      for (const [side, css] of [['top', paint.borderTop], ['bottom', paint.borderBottom], ['left', paint.borderLeft], ['right', paint.borderRight]] as const) {
        const stroke = strokeOf(css)
        if (!stroke) continue
        const [x1, y1, x2, y2] = side === 'top' ? [x, y, x + cellWidth, y] : side === 'bottom' ? [x, y + cellHeight, x + cellWidth, y + cellHeight] : side === 'left' ? [x, y, x, y + cellHeight] : [x + cellWidth, y, x + cellWidth, y + cellHeight]
        parts.push(<line key={`${side}${row}-${col}`} x1={x1} y1={y1} x2={x2} y2={y2} stroke={stroke.color} strokeWidth={stroke.width} />)
      }
    }
  }
  return (
    <svg className="table-style-preview" viewBox={`0 0 ${PREVIEW_COLUMNS * cellWidth} ${PREVIEW_ROWS * cellHeight}`} aria-hidden="true" preserveAspectRatio="none">
      {parts}
    </svg>
  )
})

function styleLabel(name: string) {
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name)
  return match ? `${match[1]} ${match[2]}` : name
}

export function TableStyleGallery({ value, theme, onPick, options }: {
  value?: string
  theme?: readonly string[]
  onPick: (style: string) => void
  options?: Pick<NonNullable<SheetTable['style']>, 'showRowStripes' | 'showColumnStripes' | 'showFirstColumn' | 'showLastColumn'>
}) {
  const groups = useMemo(() => (['Light', 'Medium', 'Dark'] as const).map((group) => ({
    group,
    styles: TABLE_STYLE_GALLERY.filter((name) => name.startsWith(`TableStyle${group}`)),
  })), [])
  return (
    <div className="table-style-gallery">
      {groups.map(({ group, styles }) => (
        <section key={group} className="table-style-group">
          <h4>{group}</h4>
          <div className="table-style-grid" role="listbox" aria-label={`${group} table styles`}>
            {styles.map((name) => (
              <button
                key={name}
                type="button"
                role="option"
                aria-selected={value === name}
                className={`table-style-option${value === name ? ' is-active' : ''}`}
                title={`Table style ${styleLabel(name)}`}
                data-table-style={name}
                onClick={() => onPick(name)}
              >
                <TableStylePreview styleName={name} theme={theme} {...options} />
              </button>
            ))}
          </div>
        </section>
      ))}
    </div>
  )
}

export interface CreateTableDialogProps {
  initialRange: string
  initialHasHeaders: boolean
  initialStyle?: string
  theme?: readonly string[]
  /** Creates the table; returns an error message to keep the dialog open. */
  onCreate: (range: string, hasHeaders: boolean, style: string) => string | null
  onClose: () => void
}

export function CreateTableDialog({ initialRange, initialHasHeaders, initialStyle = 'TableStyleMedium2', theme, onCreate, onClose }: CreateTableDialogProps) {
  const [range, setRange] = useState(initialRange)
  const [hasHeaders, setHasHeaders] = useState(initialHasHeaders)
  const [style, setStyle] = useState(initialStyle)
  const [error, setError] = useState<string | null>(null)
  const create = () => {
    const problem = onCreate(range.trim().replace(/^=/, ''), hasHeaders, style)
    if (problem) setError(problem)
  }
  return (
    <DataToolsDialog
      title="Create table"
      subtitle="Turn a range into a structured table with its own name, filters and styles."
      icon={<Table2 size={16} />}
      width={560}
      className="create-table-dialog"
      onClose={onClose}
      onConfirm={create}
      footer={(
        <>
          <span className="dt-footer-start" />
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" onClick={create}>Create table</button>
        </>
      )}
    >
      <label className="dt-field">
        <span>Where is the data for your table?</span>
        <input
          className={`dt-input dt-mono${error ? ' is-invalid' : ''}`}
          value={range}
          spellCheck={false}
          aria-label="Table range"
          data-autofocus
          onChange={(event) => { setRange(event.target.value); setError(null) }}
        />
      </label>
      <label className="dt-check"><input type="checkbox" checked={hasHeaders} onChange={(event) => setHasHeaders(event.target.checked)} />My table has headers</label>
      {error && <p className="dt-error" role="alert">{error}</p>}
      <div className="create-table-styles">
        <span className="create-table-label">Style</span>
        <TableStyleGallery value={style} theme={theme} onPick={setStyle} />
      </div>
    </DataToolsDialog>
  )
}

export type TableOption = 'headerRow' | 'totalsRow' | 'showRowStripes' | 'showColumnStripes' | 'showFirstColumn' | 'showLastColumn' | 'showFilterButton'

export interface TableDesignPanelProps {
  table: SheetTable
  theme?: readonly string[]
  onRename: (name: string) => string | null
  onResize: (range: string) => string | null
  onOption: (option: TableOption, value: boolean) => void
  onStyle: (style: string) => void
  onTotalsFunction: (index: number, fn: TotalFunctionId) => void
  onConvert: () => void
  onRemoveDuplicates: () => void
  onClose: () => void
}

function Toggle({ label, checked, onChange, disabled }: { label: string; checked: boolean; onChange: (value: boolean) => void; disabled?: boolean }) {
  return (
    <label className={`chart-toggle${disabled ? ' is-disabled' : ''}`}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(event) => onChange(event.target.checked)} />
      <span>{label}</span>
    </label>
  )
}

/** Excel's Table Design tab as a docked panel beside the grid. */
export function TableDesignPanel({ table, theme, onRename, onResize, onOption, onStyle, onTotalsFunction, onConvert, onRemoveDuplicates, onClose }: TableDesignPanelProps) {
  const [name, setName] = useState(table.name)
  const [range, setRange] = useState(table.ref)
  const [nameError, setNameError] = useState<string | null>(null)
  const [rangeError, setRangeError] = useState<string | null>(null)
  useEffect(() => { setName(table.name); setNameError(null) }, [table.name])
  useEffect(() => { setRange(table.ref); setRangeError(null) }, [table.ref])
  const style = table.style || {}
  const commitName = () => {
    if (name.trim() === table.name) return
    const problem = onRename(name)
    setNameError(problem)
    if (problem) setName(table.name)
  }
  const commitRange = () => {
    if (range.trim().toUpperCase() === table.ref.toUpperCase()) return
    const problem = onResize(range.trim().replace(/^=/, ''))
    setRangeError(problem)
  }
  const headerRow = table.headerRow !== false
  return (
    <aside className="chart-editor table-design-panel" aria-label="Table design" data-chart-keep-selection>
      <header className="chart-editor-header">
        <span className="chart-editor-icon"><Table2 size={16} aria-hidden="true" /></span>
        <h2>Table design</h2>
        <button type="button" className="chart-icon-button" onClick={onClose} aria-label="Close table design" title="Close"><X size={16} /></button>
      </header>
      <div className="chart-editor-body">
        <section className="chart-editor-section">
          <h3>Properties</h3>
          <label className="chart-field">
            <span className="chart-label">Table name</span>
            <input
              value={name}
              spellCheck={false}
              aria-label="Table name"
              className={nameError ? 'has-error' : ''}
              onChange={(event) => { setName(event.target.value); setNameError(null) }}
              onBlur={commitName}
              onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); commitName() } if (event.key === 'Escape') { setName(table.name); setNameError(null) } }}
            />
          </label>
          {nameError && <p className="chart-error" role="alert">{nameError}</p>}
          <label className="chart-field">
            <span className="chart-label">Range</span>
            <input
              value={range}
              spellCheck={false}
              aria-label="Table range"
              className={rangeError ? 'has-error' : ''}
              onChange={(event) => { setRange(event.target.value); setRangeError(null) }}
              onBlur={commitRange}
              onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); commitRange() } if (event.key === 'Escape') { setRange(table.ref); setRangeError(null) } }}
            />
          </label>
          {rangeError && <p className="chart-error" role="alert">{rangeError}</p>}
        </section>
        <section className="chart-editor-section">
          <h3>Style options</h3>
          <div className="table-design-toggles">
            <Toggle label="Header row" checked={headerRow} onChange={(value) => onOption('headerRow', value)} />
            <Toggle label="First column" checked={Boolean(style.showFirstColumn)} onChange={(value) => onOption('showFirstColumn', value)} />
            <Toggle label="Total row" checked={Boolean(table.totalsRow)} onChange={(value) => onOption('totalsRow', value)} />
            <Toggle label="Last column" checked={Boolean(style.showLastColumn)} onChange={(value) => onOption('showLastColumn', value)} />
            <Toggle label="Banded rows" checked={style.showRowStripes !== false} onChange={(value) => onOption('showRowStripes', value)} />
            <Toggle label="Banded columns" checked={Boolean(style.showColumnStripes)} onChange={(value) => onOption('showColumnStripes', value)} />
            <Toggle label="Filter button" checked={headerRow && table.showFilterButton !== false} disabled={!headerRow} onChange={(value) => onOption('showFilterButton', value)} />
          </div>
        </section>
        {table.totalsRow && (
          <section className="chart-editor-section">
            <h3>Total row</h3>
            <div className="table-design-totals">
              {table.columns.map((column, index) => (
                <label key={`${index}-${column.name}`} className="chart-field chart-field-compact">
                  <span className="chart-label" title={column.name}>{column.name}</span>
                  <select
                    aria-label={`Total for ${column.name}`}
                    value={column.totalsRowFunction && column.totalsRowFunction !== 'custom' ? column.totalsRowFunction : 'none'}
                    onChange={(event) => onTotalsFunction(index, event.target.value as TotalFunctionId)}
                  >
                    {TOTAL_FUNCTIONS.map((fn) => <option key={fn.id} value={fn.id}>{fn.label}</option>)}
                  </select>
                </label>
              ))}
            </div>
          </section>
        )}
        <section className="chart-editor-section">
          <h3>Table styles</h3>
          <TableStyleGallery value={style.theme || 'TableStyleMedium2'} theme={theme} onPick={onStyle} options={style} />
        </section>
      </div>
      <footer className="chart-editor-footer">
        <button type="button" className="chart-primary-button" onClick={onRemoveDuplicates}>Remove duplicates…</button>
        <button type="button" className="chart-danger-button" onClick={onConvert}>Convert to range</button>
      </footer>
    </aside>
  )
}
