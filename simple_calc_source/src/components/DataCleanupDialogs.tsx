import { useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { MouseEvent as ReactMouseEvent } from 'react'
import { CircleAlert, CircleCheck, CopyX, Split } from 'lucide-react'
import { guessDelimiters, previewSplit } from '../lib/data-cleanup'
import type { SplitColumnFormat, SplitOptions } from '../lib/data-cleanup'
import { parseAddress } from '../lib/data-tools-core'
import { DataToolsDialog } from './DataToolsDialogFrame'
import './data-tools.css'

// ---------------------------------------------------------------------------------------------
// Remove duplicates
// ---------------------------------------------------------------------------------------------

export interface RemoveDuplicatesDialogProps {
  /** e.g. "A1:D500". */
  rangeLabel: string
  initialHasHeader: boolean
  /** Column names for the checklist (header text or "Column A"): sortKeyLabels(bounds, 'rows', hasHeader, host). */
  labelsFor: (hasHeader: boolean) => string[]
  /**
   * Run removeDuplicates and apply it. Return the counts to show Excel's summary, an error to
   * show inline, or nothing to close immediately.
   */
  onApply: (options: { columns: number[]; hasHeader: boolean }) => { removed: number; remaining: number } | { error: string } | void
  onClose: () => void
}

/** Excel's Remove Duplicates: pick the compared columns, toggle the header, then see what was removed. */
export function RemoveDuplicatesDialog({ rangeLabel, initialHasHeader, labelsFor, onApply, onClose }: RemoveDuplicatesDialogProps) {
  const [hasHeader, setHasHeader] = useState(initialHasHeader)
  const labels = useMemo(() => labelsFor(hasHeader), [hasHeader, labelsFor])
  const [unchecked, setUnchecked] = useState<Set<number>>(() => new Set())
  const [error, setError] = useState('')
  const [result, setResult] = useState<{ removed: number; remaining: number } | null>(null)
  const okRef = useRef<HTMLButtonElement>(null)
  const selected = labels.map((_, index) => index).filter((index) => !unchecked.has(index))

  const submit = () => {
    if (result) { onClose(); return }
    if (!selected.length) { setError('Select at least one column.'); return }
    const outcome = onApply({ columns: selected, hasHeader })
    if (!outcome) { onClose(); return }
    if ('error' in outcome) { setError(outcome.error); return }
    setResult(outcome)
    window.requestAnimationFrame(() => okRef.current?.focus())
  }

  return (
    <DataToolsDialog
      title="Remove duplicates"
      subtitle={<>In <strong>{rangeLabel}</strong></>}
      icon={<CopyX size={17} />}
      width={440}
      onClose={onClose}
      onConfirm={submit}
      footer={result ? (
        <button ref={okRef} type="button" className="dt-button is-primary" onClick={onClose}>OK</button>
      ) : (
        <>
          <span className="dt-footer-start dt-footer-note">{selected.length} of {labels.length} columns</span>
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" disabled={!selected.length} onClick={submit}>Remove duplicates</button>
        </>
      )}
    >
      {result ? (
        <div className={`dt-result${result.removed ? '' : ' is-neutral'}`} role="status">
          {result.removed ? <CircleCheck size={22} aria-hidden="true" /> : <CircleAlert size={22} aria-hidden="true" />}
          <div>
            <strong>{result.removed ? `${result.removed.toLocaleString()} duplicate ${result.removed === 1 ? 'row' : 'rows'} removed` : 'No duplicate values found'}</strong>
            <span>{result.remaining.toLocaleString()} unique {result.remaining === 1 ? 'row remains' : 'rows remain'}.{result.removed ? ' Use Undo to bring them back.' : ''}</span>
          </div>
        </div>
      ) : (
        <>
          <p className="dt-lead">Rows whose selected columns all match an earlier row are deleted and the rest move up. Matching ignores case and compares values as displayed.</p>
          <div className="dt-toolbar">
            <button type="button" className="dt-button is-compact" onClick={() => { setError(''); setUnchecked(new Set()) }}>Select all</button>
            <button type="button" className="dt-button is-compact" onClick={() => setUnchecked(new Set(labels.map((_, index) => index)))}>Unselect all</button>
            <span className="dt-toolbar-spacer" />
            <label className="dt-check"><input type="checkbox" checked={hasHeader} onChange={(event) => setHasHeader(event.target.checked)} />My data has headers</label>
          </div>
          <div className="dt-column-list" role="group" aria-label="Columns to compare">
            {labels.map((label, index) => (
              <label key={index} className="dt-column-row">
                <input type="checkbox" checked={!unchecked.has(index)} onChange={(event) => {
                  setError('')
                  setUnchecked((current) => {
                    const next = new Set(current)
                    if (event.target.checked) next.delete(index)
                    else next.add(index)
                    return next
                  })
                }} />
                <span>{label}</span>
              </label>
            ))}
          </div>
          {error ? <div className="dt-error" role="alert"><CircleAlert size={13} aria-hidden="true" />{error}</div> : null}
        </>
      )}
    </DataToolsDialog>
  )
}

// ---------------------------------------------------------------------------------------------
// Text to columns
// ---------------------------------------------------------------------------------------------

export interface TextToColumnsDialogProps {
  /** e.g. "A2:A120". */
  sourceLabel: string
  /** Source text, one entry per row (splitSourceLines(bounds, host)); the preview uses the first 10. */
  sampleLines: string[]
  /** Initial destination, usually the source's first cell ("A2"). */
  defaultDestination: string
  /**
   * Run splitTextToColumns with these options. Return { confirm } (e.g. "There's already data
   * here. Replace it?") to ask first — the dialog calls again with confirmOverwrite = true — or
   * { error } to show it; return nothing when applied.
   */
  onApply: (options: SplitOptions, confirmOverwrite: boolean) => { error?: string; confirm?: string } | void
  onClose: () => void
}

const FORMAT_LABELS: Record<SplitColumnFormat, string> = {
  general: 'General',
  text: 'Text',
  'date-mdy': 'Date MDY',
  'date-dmy': 'Date DMY',
  'date-ymd': 'Date YMD',
  skip: 'Skip',
}

/** Excel's Convert Text to Columns wizard on one screen: delimiters or fixed width, per-column formats, live preview. */
export function TextToColumnsDialog({ sourceLabel, sampleLines, defaultDestination, onApply, onClose }: TextToColumnsDialogProps) {
  const guessed = useMemo(() => guessDelimiters(sampleLines), [sampleLines])
  const [mode, setMode] = useState<SplitOptions['mode']>('delimited')
  const [delimiters, setDelimiters] = useState({ tab: false, semicolon: false, comma: false, space: false, ...guessed, otherOn: false, other: '' })
  const [consecutive, setConsecutive] = useState(Boolean(guessed.space))
  const [qualifier, setQualifier] = useState<'"' | "'" | ''>('"')
  const [breaks, setBreaks] = useState<number[]>([])
  const [breaksText, setBreaksText] = useState('')
  const [formats, setFormats] = useState<SplitColumnFormat[]>([])
  const [selectedColumn, setSelectedColumn] = useState(0)
  const [destination, setDestination] = useState(defaultDestination)
  const [error, setError] = useState('')
  const [confirm, setConfirm] = useState('')
  const rulerRef = useRef<HTMLDivElement>(null)
  const measureRef = useRef<HTMLSpanElement>(null)
  const [charWidth, setCharWidth] = useState(7)

  useLayoutEffect(() => {
    const width = measureRef.current?.getBoundingClientRect().width
    if (width) setCharWidth(width / 10)
  }, [mode])

  const options: SplitOptions = useMemo(() => ({
    mode,
    delimiters: { tab: delimiters.tab, semicolon: delimiters.semicolon, comma: delimiters.comma, space: delimiters.space, other: delimiters.otherOn ? delimiters.other : '' },
    treatConsecutiveAsOne: consecutive,
    textQualifier: qualifier,
    breaks,
    columnFormats: formats,
  }), [breaks, consecutive, delimiters, formats, mode, qualifier])

  const preview = useMemo(() => previewSplit(sampleLines, options, 10), [options, sampleLines])
  const columnCount = Math.max(1, preview.columns)
  const formatOf = (index: number): SplitColumnFormat => formats[index] || 'general'
  const setFormat = (index: number, format: SplitColumnFormat) => {
    setConfirm('')
    setFormats((current) => {
      const next = [...current]
      while (next.length <= index) next.push('general')
      next[index] = format
      return next
    })
  }
  const changed = () => { setError(''); setConfirm('') }

  const toggleBreak = (position: number) => {
    if (position <= 0) return
    changed()
    setBreaks((current) => {
      const next = current.includes(position) ? current.filter((value) => value !== position) : [...current, position].sort((a, b) => a - b)
      setBreaksText(next.join(', '))
      return next
    })
  }

  const rulerClick = (event: ReactMouseEvent<HTMLDivElement>) => {
    const rect = rulerRef.current?.getBoundingClientRect()
    if (!rect) return
    const x = event.clientX - rect.left - 8 + (rulerRef.current?.scrollLeft || 0)
    toggleBreak(Math.round(x / charWidth))
  }

  const longest = Math.min(200, sampleLines.slice(0, 10).reduce((max, line) => Math.max(max, line.length), 0))

  const submit = () => {
    const target = parseAddress(destination.replace(/^=/, ''))
    if (!target) { setError('Enter a destination cell such as B2.'); return }
    if (mode === 'delimited' && !delimiters.tab && !delimiters.semicolon && !delimiters.comma && !delimiters.space && !(delimiters.otherOn && delimiters.other)) {
      setError('Choose at least one delimiter.')
      return
    }
    const outcome = onApply({ ...options, destination: target }, Boolean(confirm))
    if (!outcome) return
    if (outcome.error) { setError(outcome.error); setConfirm(''); return }
    if (outcome.confirm) setConfirm(outcome.confirm)
  }

  return (
    <DataToolsDialog
      title="Split text to columns"
      subtitle={<>Source <strong>{sourceLabel}</strong> · {sampleLines.length.toLocaleString()} {sampleLines.length === 1 ? 'row' : 'rows'}</>}
      icon={<Split size={17} />}
      width={720}
      className="dt-split-dialog"
      onClose={onClose}
      onConfirm={submit}
      footer={(
        <>
          <label className="dt-inline-field dt-footer-start">
            <span>Destination</span>
            <input className="dt-input is-narrow" aria-label="Destination cell" value={destination} onChange={(event) => { changed(); setDestination(event.target.value) }} />
          </label>
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className={`dt-button is-primary${confirm ? ' is-warning' : ''}`} onClick={submit}>{confirm ? 'Replace' : 'Split'}</button>
        </>
      )}
    >
      <div className="dt-segmented" role="radiogroup" aria-label="Split type">
        <button type="button" role="radio" aria-checked={mode === 'delimited'} className={mode === 'delimited' ? 'is-active' : ''} onClick={() => { changed(); setMode('delimited') }}>Delimited</button>
        <button type="button" role="radio" aria-checked={mode === 'fixed'} className={mode === 'fixed' ? 'is-active' : ''} onClick={() => { changed(); setMode('fixed') }}>Fixed width</button>
      </div>

      {mode === 'delimited' ? (
        <div className="dt-split-options">
          <fieldset className="dt-fieldset">
            <legend>Delimiters</legend>
            {([['tab', 'Tab'], ['semicolon', 'Semicolon'], ['comma', 'Comma'], ['space', 'Space']] as const).map(([key, label]) => (
              <label key={key} className="dt-check"><input type="checkbox" checked={delimiters[key]} onChange={(event) => { changed(); setDelimiters((current) => ({ ...current, [key]: event.target.checked })) }} />{label}</label>
            ))}
            <label className="dt-check">
              <input type="checkbox" checked={delimiters.otherOn} onChange={(event) => { changed(); setDelimiters((current) => ({ ...current, otherOn: event.target.checked })) }} />Other
              <input className="dt-input is-char" aria-label="Other delimiter" maxLength={1} value={delimiters.other} onChange={(event) => { changed(); setDelimiters((current) => ({ ...current, other: event.target.value, otherOn: Boolean(event.target.value) })) }} />
            </label>
          </fieldset>
          <fieldset className="dt-fieldset">
            <legend>Parsing</legend>
            <label className="dt-check"><input type="checkbox" checked={consecutive} onChange={(event) => { changed(); setConsecutive(event.target.checked) }} />Treat consecutive delimiters as one</label>
            <label className="dt-inline-field">
              <span>Text qualifier</span>
              <select className="dt-select is-narrow" value={qualifier} onChange={(event) => { changed(); setQualifier(event.target.value as '"' | "'" | '') }}>
                <option value={'"'}>"</option>
                <option value="'">'</option>
                <option value="">{'{none}'}</option>
              </select>
            </label>
          </fieldset>
        </div>
      ) : (
        <div className="dt-fixed">
          <label className="dt-inline-field">
            <span>Break after characters</span>
            <input className="dt-input" placeholder="e.g. 5, 12" value={breaksText} onChange={(event) => {
              changed()
              setBreaksText(event.target.value)
              setBreaks([...new Set(event.target.value.split(/[,\s]+/).map(Number).filter((value) => Number.isInteger(value) && value > 0))].sort((a, b) => a - b))
            }} />
          </label>
          <p className="dt-hint">Click the preview to add or remove a break line.</p>
          <div ref={rulerRef} className="dt-ruler" onClick={rulerClick} role="presentation">
            <span ref={measureRef} className="dt-ruler-measure" aria-hidden="true">0000000000</span>
            <div className="dt-ruler-scale" style={{ width: (longest + 2) * charWidth }}>
              {Array.from({ length: Math.floor((longest + 1) / 10) + 1 }, (_, index) => (
                <span key={index} style={{ left: index * 10 * charWidth }}>{index * 10}</span>
              ))}
            </div>
            {sampleLines.slice(0, 10).map((line, index) => <div key={index} className="dt-ruler-line">{line || ' '}</div>)}
            {breaks.map((position) => <span key={position} className="dt-ruler-break" style={{ left: 8 + position * charWidth }} />)}
          </div>
        </div>
      )}

      <div className="dt-format-row">
        <span className="dt-format-title">Column {selectedColumn + 1} format</span>
        <div className="dt-segmented is-small" role="radiogroup" aria-label={`Column ${selectedColumn + 1} data format`}>
          {(['general', 'text', 'date-mdy', 'skip'] as const).map((format) => {
            const active = format === 'date-mdy' ? formatOf(selectedColumn).startsWith('date') : formatOf(selectedColumn) === format
            return (
              <button key={format} type="button" role="radio" aria-checked={active} className={active ? 'is-active' : ''} onClick={() => setFormat(selectedColumn, format === 'date-mdy' && formatOf(selectedColumn).startsWith('date') ? formatOf(selectedColumn) : format)}>
                {format === 'date-mdy' ? 'Date' : format === 'skip' ? 'Do not import' : FORMAT_LABELS[format]}
              </button>
            )
          })}
        </div>
        {formatOf(selectedColumn).startsWith('date') ? (
          <select className="dt-select is-narrow" aria-label="Date order" value={formatOf(selectedColumn)} onChange={(event) => setFormat(selectedColumn, event.target.value as SplitColumnFormat)}>
            <option value="date-mdy">MDY</option>
            <option value="date-dmy">DMY</option>
            <option value="date-ymd">YMD</option>
          </select>
        ) : null}
      </div>

      <div className="dt-preview" role="table" aria-label="Preview of the split data">
        <table>
          <thead>
            <tr>
              {Array.from({ length: columnCount }, (_, index) => (
                <th key={index} className={`${index === selectedColumn ? 'is-selected' : ''}${formatOf(index) === 'skip' ? ' is-skipped' : ''}`}>
                  <button type="button" onClick={() => setSelectedColumn(index)} aria-pressed={index === selectedColumn}>{FORMAT_LABELS[formatOf(index)]}</button>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {preview.rows.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {Array.from({ length: columnCount }, (_, index) => (
                  <td key={index} className={`${index === selectedColumn ? 'is-selected' : ''}${formatOf(index) === 'skip' ? ' is-skipped' : ''}`} onClick={() => setSelectedColumn(index)}>{row[index] ?? ''}</td>
                ))}
              </tr>
            ))}
            {!preview.rows.length ? <tr><td className="dt-preview-empty">No text in the selection</td></tr> : null}
          </tbody>
        </table>
      </div>
      {confirm ? <div className="dt-banner is-warning" role="alert"><CircleAlert size={14} aria-hidden="true" /><span>{confirm}</span></div> : null}
      {error ? <div className="dt-error" role="alert"><CircleAlert size={13} aria-hidden="true" />{error}</div> : null}
    </DataToolsDialog>
  )
}
