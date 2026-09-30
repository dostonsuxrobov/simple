import { memo, useEffect, useId, useMemo, useState, type ReactNode } from 'react'
import { ChartColumnBig, SquareDashedMousePointer, Trash2, X } from 'lucide-react'
import type { ChartAxis, ChartDataLabels, ChartLegendPosition, ChartSeries, ChartSeriesType, SheetChart } from '../spreadsheet-types'
import {
  CHART_PALETTES,
  CHART_TYPE_OPTIONS,
  applyChartDataRange,
  changeChartType,
  inferChartDataRange,
  chartPalette,
  chartTypeKey,
  normalizeHex,
  paletteColor,
  parseChartRef,
  updateChart,
  type ChartTypeOption,
  type ChartWorkbookAccessor,
  type ResolvedChartData,
} from '../lib/charts'
import { renderChartThumbnail } from '../lib/chart-render'
import './charts.css'

export interface ChartEditorPanelProps {
  chart: SheetChart
  sheetNames: string[]
  /** Sheet the chart lives on; unqualified ranges ("A1:C9") refer to it. */
  activeSheetName?: string
  /** Calculated-value accessor, used to rebuild series when the range or layout changes. */
  accessor?: ChartWorkbookAccessor | null
  /** Resolved data of the chart (used for series names/colours in the list). */
  data?: ResolvedChartData | null
  /** Workbook theme palette (see workbookChartPalette). */
  themePalette?: string[]
  onChange: (chart: SheetChart) => void
  onClose: () => void
  onDelete: () => void
  /** Optional interactive range picker: call `apply` with the chosen range text. */
  onPickRange?: (apply: (rangeText: string) => void) => void
  initialTab?: 'setup' | 'customize'
  className?: string
}

const SAMPLE: ResolvedChartData = {
  categories: ['A', 'B', 'C', 'D'],
  series: [
    { name: 'One', values: [4, 6, 5, 8], x: [1, 2, 3, 4], color: '#476B57' },
    { name: 'Two', values: [3, 4, 6, 5], x: [1.5, 2.5, 3.2, 4.4], color: '#8FB3A0' },
    { name: 'Three', values: [2, 3, 2, 4], x: [0.8, 2.2, 3.6, 3.9], color: '#C9A227' },
  ],
}

const THUMB_PALETTE = ['#476B57', '#8FB3A0', '#C9A227']

function thumbnailChart(option: ChartTypeOption): SheetChart {
  const series: ChartSeries[] = SAMPLE.series.map((item, index) => ({
    id: `thumb-${index}`,
    name: item.name,
    marker: option.type === 'scatter' ? 'circle' : 'none',
    markerSize: 3,
    lineWidth: 1.5,
    showLine: option.type === 'scatter' ? false : undefined,
    type: option.type === 'combo' ? (index === SAMPLE.series.length - 1 ? 'line' : 'column') : undefined,
  }))
  return {
    id: `thumb-${option.key}`,
    type: option.type,
    grouping: option.grouping,
    anchor: { from: { row: 0, col: 0 }, to: { row: 1, col: 1 } },
    series: option.type === 'pie' || option.type === 'doughnut' ? series.slice(0, 1) : series,
    legend: 'none',
    holeSize: 55,
    gapWidth: 60,
    axes: { x: { visible: false }, y: { visible: false, gridlines: false } },
    style: { palette: THUMB_PALETTE, border: null },
  }
}

const TypeThumbnail = memo(function TypeThumbnail({ option }: { option: ChartTypeOption }) {
  const svg = useMemo(() => {
    const chart = thumbnailChart(option)
    const data: ResolvedChartData = option.type === 'pie' || option.type === 'doughnut'
      ? { categories: SAMPLE.categories, series: [{ ...SAMPLE.series[0], pointColors: ['#476B57', '#8FB3A0', '#C9A227', '#5C7FA3'] }] }
      : SAMPLE
    return renderChartThumbnail(chart, data, 58, 38)
  }, [option])
  return <span className="chart-type-thumb" aria-hidden="true" dangerouslySetInnerHTML={{ __html: svg }} />
})

function Section({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section className="chart-editor-section">
      <h3>{title}</h3>
      {children}
    </section>
  )
}

function Toggle({ label, checked, onChange, disabled }: { label: string; checked: boolean; onChange: (value: boolean) => void; disabled?: boolean }) {
  return (
    <label className={`chart-toggle${disabled ? ' is-disabled' : ''}`}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(event) => onChange(event.target.checked)} />
      <span>{label}</span>
    </label>
  )
}

function Segmented<T extends string>({ value, options, onChange, label }: { value: T; options: Array<{ value: T; label: string }>; onChange: (value: T) => void; label: string }) {
  return (
    <div className="chart-segmented" role="radiogroup" aria-label={label}>
      {options.map((option) => (
        <button key={option.value} type="button" role="radio" aria-checked={value === option.value} className={value === option.value ? 'is-active' : ''} onClick={() => onChange(option.value)}>
          {option.label}
        </button>
      ))}
    </div>
  )
}

function NumberField({ label, value, onCommit, placeholder }: { label: string; value: number | undefined; onCommit: (value: number | undefined) => void; placeholder?: string }) {
  const [text, setText] = useState(value === undefined ? '' : String(value))
  useEffect(() => { setText(value === undefined ? '' : String(value)) }, [value])
  const commit = () => {
    const trimmed = text.trim()
    if (!trimmed) { if (value !== undefined) onCommit(undefined); return }
    const number = Number(trimmed)
    if (Number.isFinite(number)) { if (number !== value) onCommit(number) } else setText(value === undefined ? '' : String(value))
  }
  return (
    <label className="chart-field chart-field-compact">
      <span>{label}</span>
      <input type="text" inputMode="decimal" value={text} placeholder={placeholder || 'Auto'} onChange={(event) => setText(event.target.value)} onBlur={commit} onKeyDown={(event) => { if (event.key === 'Enter') commit() }} />
    </label>
  )
}

function TextField({ label, value, onCommit, placeholder }: { label: string; value: string | undefined; onCommit: (value: string) => void; placeholder?: string }) {
  const [text, setText] = useState(value ?? '')
  useEffect(() => { setText(value ?? '') }, [value])
  const commit = () => { if (text !== (value ?? '')) onCommit(text) }
  return (
    <label className="chart-field">
      <span>{label}</span>
      <input type="text" value={text} placeholder={placeholder} onChange={(event) => setText(event.target.value)} onBlur={commit} onKeyDown={(event) => { if (event.key === 'Enter') commit() }} />
    </label>
  )
}

const LEGEND_OPTIONS: Array<{ value: ChartLegendPosition; label: string }> = [
  { value: 'none', label: 'None' },
  { value: 'bottom', label: 'Bottom' },
  { value: 'top', label: 'Top' },
  { value: 'right', label: 'Right' },
  { value: 'left', label: 'Left' },
]

const SERIES_TYPES: Array<{ value: ChartSeriesType; label: string }> = [
  { value: 'column', label: 'Column' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
]

/** Google Sheets-style side panel for inserting and editing a chart. */
export function ChartEditorPanel({ chart, sheetNames, activeSheetName, accessor, data, themePalette, onChange, onClose, onDelete, onPickRange, initialTab = 'setup', className }: ChartEditorPanelProps) {
  const [tab, setTab] = useState<'setup' | 'customize'>(initialTab)
  // Imported charts have no typed range; offer the block their references cover.
  const inferred = useMemo(() => (chart.dataRange ? null : inferChartDataRange(chart)), [chart])
  const [rangeText, setRangeText] = useState(chart.dataRange || inferred?.range || '')
  const [rangeError, setRangeError] = useState('')
  const id = useId()
  useEffect(() => { setRangeText(chart.dataRange || inferred?.range || ''); setRangeError('') }, [chart.id, chart.dataRange, inferred?.range])

  const unsupported = chart.type === 'unsupported'
  const pie = chart.type === 'pie' || chart.type === 'doughnut'
  const cartesian = !pie && chart.type !== 'radar' && !unsupported
  const stackable = chart.type === 'column' || chart.type === 'bar' || chart.type === 'line' || chart.type === 'area'
  const hasLines = chart.type === 'line' || chart.type === 'scatter' || chart.type === 'radar' || (chart.type === 'combo' && chart.series.some((item) => item.type === 'line'))
  const palette = chartPalette(chart, themePalette)
  const sheetName = activeSheetName || sheetNames[0] || 'Sheet1'

  const emit = (patch: Partial<SheetChart>) => onChange(updateChart(chart, patch))
  const setAxis = (key: 'x' | 'y' | 'y2', patch: Partial<ChartAxis>) => {
    const axes = { ...(chart.axes || {}) }
    const next = { ...(axes[key] || {}), ...patch }
    for (const field of Object.keys(next) as Array<keyof ChartAxis>) if (next[field] === undefined) delete next[field]
    axes[key] = next
    emit({ axes })
  }
  const setSeries = (index: number, patch: Partial<ChartSeries>) => {
    const series = chart.series.map((item, position) => {
      if (position !== index) return item
      const next = { ...item, ...patch }
      for (const field of Object.keys(next) as Array<keyof ChartSeries>) if (next[field] === undefined) delete next[field]
      return next
    })
    emit({ series })
  }
  const setAllSeries = (patch: Partial<ChartSeries>) => emit({ series: chart.series.map((item) => ({ ...item, ...patch })) })

  const applyRange = (text: string, layout?: Parameters<typeof applyChartDataRange>[4]) => {
    const trimmed = text.trim()
    const areas = parseChartRef(trimmed, sheetName)
    if (!areas || areas.length !== 1) { setRangeError('Enter a range such as A1:C12 or Sheet1!A1:C12.'); return }
    const target = areas[0].sheet || sheetName
    if (sheetNames.length && !sheetNames.some((name) => name.toLocaleLowerCase() === target.toLocaleLowerCase())) { setRangeError(`There is no sheet named “${target}”.`); return }
    const base = !chart.seriesIn && inferred ? { seriesIn: inferred.seriesIn, firstRowHeaders: inferred.firstRowHeaders, firstColumnLabels: inferred.firstColumnLabels } : {}
    const next = applyChartDataRange(chart, trimmed, accessor || null, sheetName, { ...base, ...layout })
    if (!next) { setRangeError('That range cannot be charted.'); return }
    if (!next.series.length) { setRangeError('The range has no data columns after headers and labels.'); return }
    setRangeError('')
    onChange(next)
  }

  const selectType = (option: ChartTypeOption) => {
    if (chartTypeKey(chart) === option.key) return
    onChange(changeChartType(chart, option.type, option.grouping))
  }

  const labels = chart.series[0]?.dataLabels
  const labelsOn = chart.series.some((item) => item.dataLabels && (item.dataLabels.showValue || item.dataLabels.showPercent || item.dataLabels.showCategory || item.dataLabels.showSeriesName))
  const setLabels = (patch: Partial<ChartDataLabels> | null) => {
    if (!patch) { setAllSeries({ dataLabels: undefined }); return }
    setAllSeries({ dataLabels: { ...(labels || {}), ...patch } })
  }
  const smoothOn = chart.series.some((item) => item.smooth)
  const markersOn = chart.series.some((item) => item.marker && item.marker !== 'none')
  const hasSecondary = chart.type === 'combo' && chart.series.some((item) => item.secondaryAxis)
  const paletteId = CHART_PALETTES.find((entry) => entry.colors.join() === (chart.style?.palette || []).join())?.id || (chart.style?.palette?.length ? 'custom' : 'theme')

  return (
    <aside className={`chart-editor${className ? ` ${className}` : ''}`} aria-label="Chart editor" data-chart-keep-selection>
      <header className="chart-editor-header">
        <span className="chart-editor-icon"><ChartColumnBig size={16} aria-hidden="true" /></span>
        <h2>Chart editor</h2>
        <button type="button" className="chart-icon-button" onClick={onClose} aria-label="Close chart editor" title="Close"><X size={16} /></button>
      </header>
      <div className="chart-editor-tabs" role="tablist" aria-label="Chart editor sections">
        <button type="button" role="tab" id={`${id}-setup`} aria-selected={tab === 'setup'} className={tab === 'setup' ? 'is-active' : ''} onClick={() => setTab('setup')}>Setup</button>
        <button type="button" role="tab" id={`${id}-customize`} aria-selected={tab === 'customize'} className={tab === 'customize' ? 'is-active' : ''} onClick={() => setTab('customize')}>Customize</button>
      </div>
      <div className="chart-editor-body" role="tabpanel" aria-labelledby={`${id}-${tab}`}>
        {unsupported && (
          <div className="chart-editor-note">
            This {chart.unsupportedKind ? `${chart.unsupportedKind} ` : ''}chart can’t be edited here yet. It stays in the workbook exactly as it is when you save as .xlsx; you can still move, resize or delete it.
          </div>
        )}
        {!unsupported && tab === 'setup' && (
          <>
            <Section title="Chart type">
              <div className="chart-type-grid" role="listbox" aria-label="Chart type">
                {CHART_TYPE_OPTIONS.map((option) => {
                  const active = chartTypeKey(chart) === option.key
                  return (
                    <button key={option.key} type="button" role="option" aria-selected={active} className={`chart-type-option${active ? ' is-active' : ''}`} onClick={() => selectType(option)} title={option.label}>
                      <TypeThumbnail option={option} />
                      <span>{option.label}</span>
                    </button>
                  )
                })}
              </div>
            </Section>
            <Section title="Data range">
              <div className="chart-range-row">
                <input
                  type="text"
                  className={rangeError ? 'has-error' : ''}
                  value={rangeText}
                  spellCheck={false}
                  aria-label="Data range"
                  aria-invalid={Boolean(rangeError)}
                  placeholder="A1:C12"
                  onChange={(event) => setRangeText(event.target.value)}
                  onBlur={() => { if (rangeText.trim() && rangeText.trim() !== (chart.dataRange || inferred?.range || '')) applyRange(rangeText) }}
                  onKeyDown={(event) => { if (event.key === 'Enter') applyRange(rangeText) }}
                />
                {onPickRange && (
                  <button type="button" className="chart-icon-button" title="Select range on the sheet" aria-label="Select range on the sheet" onClick={() => onPickRange((text) => { setRangeText(text); applyRange(text) })}>
                    <SquareDashedMousePointer size={16} />
                  </button>
                )}
              </div>
              {rangeError && <p className="chart-error" role="alert">{rangeError}</p>}
              {chart.dataRange && chart.type !== 'scatter' && (
                <>
                  <div className="chart-field-row">
                    <span className="chart-label">Series in</span>
                    <Segmented label="Series in" value={chart.seriesIn || 'columns'} options={[{ value: 'columns', label: 'Columns' }, { value: 'rows', label: 'Rows' }]} onChange={(value) => applyRange(chart.dataRange || rangeText, { seriesIn: value })} />
                  </div>
                  <Toggle label="Use first row as headers" checked={chart.firstRowHeaders === true} onChange={(value) => applyRange(chart.dataRange || rangeText, { firstRowHeaders: value })} />
                  <Toggle label="Use first column as labels" checked={chart.firstColumnLabels === true} onChange={(value) => applyRange(chart.dataRange || rangeText, { firstColumnLabels: value })} />
                </>
              )}
            </Section>
            <Section title={pie ? 'Values' : 'Series'}>
              {!chart.series.length && <p className="chart-muted">No series yet. Enter a data range above.</p>}
              <ul className="chart-series-list">
                {chart.series.map((series, index) => {
                  const color = normalizeHex(series.color) || data?.series[index]?.color || paletteColor(index, palette)
                  const name = data?.series[index]?.name || series.name || `Series ${index + 1}`
                  return (
                    <li key={series.id || index} className="chart-series-item">
                      {!pie && (
                        <label className="chart-swatch" style={{ background: color }} title="Series colour">
                          <input type="color" value={color.toLowerCase()} aria-label={`Colour of ${name}`} onChange={(event) => setSeries(index, { color: event.target.value.toUpperCase() })} />
                        </label>
                      )}
                      <span className="chart-series-name" title={series.valuesRef || ''}>{name}</span>
                      {chart.type === 'combo' && (
                        <>
                          <select aria-label={`Chart type of ${name}`} value={series.type || 'column'} onChange={(event) => setSeries(index, { type: event.target.value as ChartSeriesType, marker: event.target.value === 'line' ? series.marker || 'none' : series.marker })}>
                            {SERIES_TYPES.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                          </select>
                          <label className="chart-mini-toggle" title="Plot on the right axis">
                            <input type="checkbox" checked={series.secondaryAxis === true} onChange={(event) => setSeries(index, { secondaryAxis: event.target.checked || undefined })} />
                            <span>Right</span>
                          </label>
                        </>
                      )}
                    </li>
                  )
                })}
              </ul>
            </Section>
          </>
        )}
        {!unsupported && tab === 'customize' && (
          <>
            <Section title="Titles">
              <TextField label="Chart title" value={chart.title} placeholder={chart.series.length === 1 ? (data?.series[0]?.name || 'Automatic') : 'None'} onCommit={(value) => emit({ title: value, autoTitleDeleted: value === '' ? true : undefined })} />
              {cartesian && (
                <>
                  <TextField label={chart.type === 'bar' ? 'Vertical axis title' : 'Horizontal axis title'} value={chart.axes?.x?.title} onCommit={(value) => setAxis('x', { title: value || undefined })} />
                  <TextField label={chart.type === 'bar' ? 'Horizontal axis title' : 'Vertical axis title'} value={chart.axes?.y?.title} onCommit={(value) => setAxis('y', { title: value || undefined })} />
                  {hasSecondary && <TextField label="Right axis title" value={chart.axes?.y2?.title} onCommit={(value) => setAxis('y2', { title: value || undefined })} />}
                </>
              )}
            </Section>
            <Section title="Legend">
              <label className="chart-field">
                <span>Position</span>
                <select value={chart.legend || 'none'} onChange={(event) => emit({ legend: event.target.value as ChartLegendPosition })}>
                  {LEGEND_OPTIONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                </select>
              </label>
            </Section>
            <Section title="Series">
              {stackable && (
                <div className="chart-field-row">
                  <span className="chart-label">Stacking</span>
                  <Segmented label="Stacking" value={chart.grouping || 'clustered'} options={[{ value: 'clustered', label: 'None' }, { value: 'stacked', label: 'Stacked' }, { value: 'percentStacked', label: '100%' }]} onChange={(value) => emit({ grouping: value })} />
                </div>
              )}
              <Toggle label="Data labels" checked={labelsOn} onChange={(value) => setLabels(value ? (pie ? { showPercent: true } : { showValue: true }) : null)} />
              {labelsOn && pie && (
                <div className="chart-indent">
                  <Toggle label="Show percentage" checked={labels?.showPercent === true} onChange={(value) => setLabels({ showPercent: value })} />
                  <Toggle label="Show value" checked={labels?.showValue === true} onChange={(value) => setLabels({ showValue: value })} />
                  <Toggle label="Show category" checked={labels?.showCategory === true} onChange={(value) => setLabels({ showCategory: value })} />
                </div>
              )}
              {hasLines && <Toggle label="Smooth lines" checked={smoothOn} onChange={(value) => emit({ series: chart.series.map((item) => (chart.type !== 'combo' || item.type === 'line' ? { ...item, smooth: value || undefined } : item)) })} />}
              {(chart.type === 'line' || chart.type === 'scatter' || chart.type === 'radar' || chart.type === 'combo') && (
                <Toggle label="Markers" checked={markersOn} onChange={(value) => emit({ series: chart.series.map((item) => (chart.type !== 'combo' || item.type === 'line' ? { ...item, marker: value ? 'circle' : 'none' } : item)) })} />
              )}
              {chart.type === 'scatter' && <Toggle label="Connect points with lines" checked={chart.series.some((item) => item.showLine)} onChange={(value) => setAllSeries({ showLine: value })} />}
              {chart.type === 'doughnut' && (
                <label className="chart-field">
                  <span>Hole size <output>{chart.holeSize ?? 75}%</output></span>
                  <input type="range" min={10} max={90} step={5} value={chart.holeSize ?? 75} onChange={(event) => emit({ holeSize: Number(event.target.value) })} />
                </label>
              )}
              {(chart.type === 'column' || chart.type === 'bar' || (chart.type === 'combo' && chart.series.some((item) => (item.type || 'column') === 'column'))) && (
                <label className="chart-field">
                  <span>Gap width <output>{chart.gapWidth ?? (chart.grouping && chart.grouping !== 'clustered' ? 150 : 219)}%</output></span>
                  <input type="range" min={0} max={500} step={10} value={chart.gapWidth ?? (chart.grouping && chart.grouping !== 'clustered' ? 150 : 219)} onChange={(event) => emit({ gapWidth: Number(event.target.value) })} />
                </label>
              )}
            </Section>
            {cartesian && (
              <Section title="Axes and gridlines">
                <Toggle label="Major gridlines" checked={chart.axes?.y?.gridlines !== false} onChange={(value) => setAxis('y', { gridlines: value })} />
                <Toggle label={chart.type === 'bar' ? 'Horizontal category gridlines' : 'Vertical gridlines'} checked={chart.axes?.x?.gridlines === true} onChange={(value) => setAxis('x', { gridlines: value || undefined })} />
                <div className="chart-field-grid">
                  <NumberField label="Min" value={chart.axes?.y?.min} onCommit={(value) => setAxis('y', { min: value })} />
                  <NumberField label="Max" value={chart.axes?.y?.max} onCommit={(value) => setAxis('y', { max: value })} />
                </div>
                <label className="chart-field">
                  <span>Number format</span>
                  <select value={chart.axes?.y?.numFmt || ''} onChange={(event) => setAxis('y', { numFmt: event.target.value || undefined })}>
                    <option value="">From source data</option>
                    <option value="General">General</option>
                    <option value="#,##0">1,235</option>
                    <option value="#,##0.00">1,234.56</option>
                    <option value="0%">12%</option>
                    <option value="0.0%">12.3%</option>
                    <option value="$#,##0">$1,235</option>
                    <option value="$#,##0.00">$1,234.56</option>
                  </select>
                </label>
              </Section>
            )}
            <Section title="Colours">
              <div className="chart-palette-list" role="radiogroup" aria-label="Colour palette">
                <button type="button" role="radio" aria-checked={paletteId === 'theme'} className={`chart-palette${paletteId === 'theme' ? ' is-active' : ''}`} onClick={() => emit({ style: { ...(chart.style || {}), palette: undefined }, series: chart.series.map(({ color: _color, ...item }) => item) })} title="Workbook theme">
                  {(themePalette || CHART_PALETTES[0].colors).slice(0, 6).map((color) => <i key={color} style={{ background: color }} />)}
                  <span>Theme</span>
                </button>
                {CHART_PALETTES.map((entry) => (
                  <button key={entry.id} type="button" role="radio" aria-checked={paletteId === entry.id} className={`chart-palette${paletteId === entry.id ? ' is-active' : ''}`} onClick={() => emit({ style: { ...(chart.style || {}), palette: entry.colors }, series: chart.series.map(({ color: _color, pointColors: _points, ...item }) => item) })} title={entry.label}>
                    {entry.colors.map((color) => <i key={color} style={{ background: color }} />)}
                    <span>{entry.label}</span>
                  </button>
                ))}
              </div>
              <Toggle label="Transparent background" checked={chart.style?.background === 'transparent'} onChange={(value) => emit({ style: { ...(chart.style || {}), background: value ? 'transparent' : undefined } })} />
              <Toggle label="Chart border" checked={chart.style?.border !== null} onChange={(value) => emit({ style: { ...(chart.style || {}), border: value ? undefined : null } })} />
            </Section>
          </>
        )}
      </div>
      <footer className="chart-editor-footer">
        <button type="button" className="chart-danger-button" onClick={onDelete}><Trash2 size={14} aria-hidden="true" /> Delete chart</button>
        <button type="button" className="chart-primary-button" onClick={onClose}>Done</button>
      </footer>
    </aside>
  )
}

export default ChartEditorPanel
