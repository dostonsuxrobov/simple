import { useMemo, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { CircleAlert, Info, ListChecks, OctagonX, ShieldCheck, TriangleAlert } from 'lucide-react'
import {
  DEFAULT_VALIDATION_DIALOG_STATE,
  DEFAULT_VALIDATION_ERROR,
  VALIDATION_OPERATORS,
  VALIDATION_TYPES,
  describeValidation,
  dialogStateToValidation,
  requirementText,
  validationToDialogState,
} from '../lib/validation'
import type { DataValidationModel, ValidationDialogState, ValidationErrorStyle, ValidationOperator } from '../lib/validation'
import { DataToolsDialog } from './DataToolsDialogFrame'
import './data-tools.css'

export interface DataValidationDialogProps {
  /** e.g. "B2:B200". */
  rangeLabel: string
  /** The selection's current rule (ExcelJS model), or null. */
  initial: unknown | null
  /** The selection holds several different rules: start from blank settings (Excel asks to erase). */
  mixed?: boolean
  /** How many other ranges share the initial rule; > 0 shows "Apply these changes to all other cells with the same settings". */
  sameSettingsCount?: number
  date1904?: boolean
  /** null removes validation from the range. */
  onApply: (validation: DataValidationModel | null, options: { applyToSameSettings: boolean }) => void
  onClose: () => void
}

type Tab = 'settings' | 'input' | 'error'

const STYLE_META: Record<ValidationErrorStyle, { label: string; icon: ReactNode; hint: string }> = {
  stop: { label: 'Stop', icon: <OctagonX size={22} />, hint: 'Invalid entries are rejected.' },
  warning: { label: 'Warning', icon: <TriangleAlert size={22} />, hint: 'People are asked whether to keep an invalid entry.' },
  information: { label: 'Information', icon: <Info size={22} />, hint: 'People are told, but the entry is kept.' },
}

function boundLabels(state: ValidationDialogState): [string, string] {
  const two = state.operator === 'between' || state.operator === 'notBetween'
  if (state.type === 'date') {
    if (two) return ['Start date', 'End date']
    return [state.operator === 'lessThan' || state.operator === 'lessThanOrEqual' ? 'End date' : state.operator === 'greaterThan' || state.operator === 'greaterThanOrEqual' ? 'Start date' : 'Date', '']
  }
  if (state.type === 'time') {
    if (two) return ['Start time', 'End time']
    return [state.operator === 'lessThan' || state.operator === 'lessThanOrEqual' ? 'End time' : state.operator === 'greaterThan' || state.operator === 'greaterThanOrEqual' ? 'Start time' : 'Time', '']
  }
  if (two) return ['Minimum', 'Maximum']
  if (state.operator === 'greaterThan' || state.operator === 'greaterThanOrEqual') return ['Minimum', '']
  if (state.operator === 'lessThan' || state.operator === 'lessThanOrEqual') return ['Maximum', '']
  return [state.type === 'textLength' ? 'Length' : 'Value', '']
}

function placeholderFor(state: ValidationDialogState, second: boolean): string {
  if (state.type === 'date') return second ? '12/31/2025' : '1/1/2025'
  if (state.type === 'time') return second ? '5:00 PM' : '9:00 AM'
  if (state.type === 'decimal') return second ? '100.5' : '0.5'
  return second ? '100' : '1'
}

/** Excel's Data Validation dialog: Settings, Input Message and Error Alert tabs. */
export function DataValidationDialog({ rangeLabel, initial, mixed = false, sameSettingsCount = 0, date1904 = false, onApply, onClose }: DataValidationDialogProps) {
  const [state, setState] = useState<ValidationDialogState>(() => (mixed ? { ...DEFAULT_VALIDATION_DIALOG_STATE } : validationToDialogState(initial, { date1904 })))
  const [tab, setTab] = useState<Tab>('settings')
  const [applyToSame, setApplyToSame] = useState(false)
  const [error, setError] = useState<{ text: string; field?: 'value1' | 'value2' } | null>(null)
  const allowRef = useRef<HTMLSelectElement>(null)
  const tabIds = useRef(`dt-dv-${Math.random().toString(36).slice(2, 8)}`).current

  const set = <K extends keyof ValidationDialogState>(key: K, value: ValidationDialogState[K]) => {
    setError(null)
    setState((current) => ({ ...current, [key]: value }))
  }

  const conversion = useMemo(() => dialogStateToValidation(state, { date1904 }), [date1904, state])
  const summary = conversion.ok ? describeValidation(conversion.validation) : ''
  const hasOperator = !['any', 'list', 'custom'].includes(state.type)
  const twoInputs = hasOperator && (state.operator === 'between' || state.operator === 'notBetween')
  const [label1, label2] = boundLabels(state)

  const submit = () => {
    const result = dialogStateToValidation(state, { date1904 })
    if (!result.ok) {
      setTab('settings')
      setError({ text: result.error, field: result.field })
      return
    }
    const validation = result.validation
    const empty = validation.type === 'any' && !validation.prompt && !validation.promptTitle
    onApply(empty ? null : validation, { applyToSameSettings: applyToSame })
  }

  const clearAll = () => {
    setError(null)
    setState({ ...DEFAULT_VALIDATION_DIALOG_STATE })
    setTab('settings')
    allowRef.current?.focus()
  }

  const styleMeta = STYLE_META[state.errorStyle]

  return (
    <DataToolsDialog
      title="Data validation"
      subtitle={<>Rules for <strong>{rangeLabel}</strong></>}
      icon={<ShieldCheck size={17} />}
      width={560}
      className="dt-validation-dialog"
      onClose={onClose}
      onConfirm={submit}
      initialFocus={allowRef}
      footer={(
        <>
          <button type="button" className="dt-button dt-footer-start" onClick={clearAll}>Clear all</button>
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" onClick={submit}>OK</button>
        </>
      )}
    >
      {mixed ? (
        <div className="dt-banner is-warning"><CircleAlert size={14} aria-hidden="true" /><span>The selection contains more than one type of validation. Choosing OK replaces them all with these settings.</span></div>
      ) : null}

      <div className="dt-tabs" role="tablist" aria-label="Validation settings">
        {([['settings', 'Settings'], ['input', 'Input message'], ['error', 'Error alert']] as const).map(([id, label]) => (
          <button
            key={id}
            id={`${tabIds}-${id}`}
            type="button"
            role="tab"
            aria-selected={tab === id}
            aria-controls={`${tabIds}-${id}-panel`}
            className={`dt-tab${tab === id ? ' is-active' : ''}`}
            onClick={() => setTab(id)}
            onKeyDown={(event) => {
              const order: Tab[] = ['settings', 'input', 'error']
              if (event.key !== 'ArrowRight' && event.key !== 'ArrowLeft') return
              event.preventDefault()
              const next = order[(order.indexOf(id) + (event.key === 'ArrowRight' ? 1 : 2)) % 3]
              setTab(next)
              document.getElementById(`${tabIds}-${next}`)?.focus()
            }}
          >
            {label}
          </button>
        ))}
      </div>

      <div id={`${tabIds}-settings-panel`} role="tabpanel" aria-labelledby={`${tabIds}-settings`} hidden={tab !== 'settings'} className="dt-tab-panel">
        <div className="dt-grid-2">
          <label className="dt-field">
            <span>Allow</span>
            <select ref={allowRef} className="dt-select" value={state.type} onChange={(event) => {
              const type = event.target.value as ValidationDialogState['type']
              setError(null)
              setState((current) => ({ ...current, type, value1: current.type === type ? current.value1 : '', value2: current.type === type ? current.value2 : '' }))
            }}>
              {VALIDATION_TYPES.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
            </select>
          </label>
          {hasOperator ? (
            <label className="dt-field">
              <span>Data</span>
              <select className="dt-select" value={state.operator} onChange={(event) => set('operator', event.target.value as ValidationOperator)}>
                {VALIDATION_OPERATORS.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
              </select>
            </label>
          ) : <span />}
        </div>

        <div className="dt-checks-row">
          <label className={`dt-check${state.type === 'any' ? ' is-disabled' : ''}`}><input type="checkbox" disabled={state.type === 'any'} checked={state.allowBlank} onChange={(event) => set('allowBlank', event.target.checked)} />Ignore blank</label>
          {state.type === 'list' ? <label className="dt-check"><input type="checkbox" checked={state.inCellDropdown} onChange={(event) => set('inCellDropdown', event.target.checked)} />In-cell dropdown</label> : null}
        </div>

        {hasOperator ? (
          <div className={twoInputs ? 'dt-grid-2' : 'dt-grid-1'}>
            <label className="dt-field">
              <span>{label1}</span>
              <input className={`dt-input${error?.field === 'value1' ? ' is-invalid' : ''}`} aria-invalid={error?.field === 'value1'} placeholder={placeholderFor(state, false)} value={state.value1} onChange={(event) => set('value1', event.target.value)} />
            </label>
            {twoInputs ? (
              <label className="dt-field">
                <span>{label2}</span>
                <input className={`dt-input${error?.field === 'value2' ? ' is-invalid' : ''}`} aria-invalid={error?.field === 'value2'} placeholder={placeholderFor(state, true)} value={state.value2} onChange={(event) => set('value2', event.target.value)} />
              </label>
            ) : null}
          </div>
        ) : null}
        {hasOperator ? <p className="dt-hint">Type a value, or a formula / reference starting with “=”, such as =$B$1.</p> : null}

        {state.type === 'list' ? (
          <label className="dt-field">
            <span>Source</span>
            <input className={`dt-input${error?.field === 'value1' ? ' is-invalid' : ''}`} aria-invalid={error?.field === 'value1'} placeholder="Yes, No, Maybe   or   =$A$1:$A$10" value={state.value1} onChange={(event) => set('value1', event.target.value)} />
            <small className="dt-hint">Separate items with commas, or refer to cells or a named range with “=”.</small>
          </label>
        ) : null}

        {state.type === 'custom' ? (
          <label className="dt-field">
            <span>Formula</span>
            <input className={`dt-input dt-mono${error?.field === 'value1' ? ' is-invalid' : ''}`} aria-invalid={error?.field === 'value1'} placeholder="=AND(ISNUMBER(A1), A1>0)" value={state.value1} onChange={(event) => set('value1', event.target.value)} />
            <small className="dt-hint">Write the formula for the first cell of {rangeLabel}; it must return TRUE for valid entries.</small>
          </label>
        ) : null}

        {error ? <div className="dt-error" role="alert"><CircleAlert size={13} aria-hidden="true" />{error.text}</div> : null}

        <div className="dt-summary" aria-live="polite">
          <ListChecks size={14} aria-hidden="true" />
          <span>{state.type === 'any' ? 'Any value is allowed.' : summary ? <>Rule: <strong>{summary}</strong>{state.allowBlank ? ' · blanks allowed' : ' · blanks rejected'}</> : 'Complete the fields to preview the rule.'}</span>
        </div>

        {sameSettingsCount > 0 ? (
          <label className="dt-check dt-apply-same"><input type="checkbox" checked={applyToSame} onChange={(event) => setApplyToSame(event.target.checked)} />Apply these changes to all other cells with the same settings ({sameSettingsCount.toLocaleString()} {sameSettingsCount === 1 ? 'range' : 'ranges'})</label>
        ) : null}
      </div>

      <div id={`${tabIds}-input-panel`} role="tabpanel" aria-labelledby={`${tabIds}-input`} hidden={tab !== 'input'} className="dt-tab-panel">
        <label className="dt-check"><input type="checkbox" checked={state.showInputMessage} onChange={(event) => set('showInputMessage', event.target.checked)} />Show input message when a cell is selected</label>
        <p className="dt-hint">When a cell is selected, show this message:</p>
        <label className="dt-field">
          <span>Title</span>
          <input className="dt-input" maxLength={32} disabled={!state.showInputMessage} value={state.promptTitle} onChange={(event) => set('promptTitle', event.target.value)} />
        </label>
        <label className="dt-field">
          <span>Input message <em className="dt-counter">{state.prompt.length}/255</em></span>
          <textarea className="dt-textarea" maxLength={255} rows={4} disabled={!state.showInputMessage} value={state.prompt} onChange={(event) => set('prompt', event.target.value)} />
        </label>
        {state.showInputMessage && (state.promptTitle || state.prompt) ? (
          <div className="dt-preview-note" aria-label="Input message preview">
            {state.promptTitle ? <strong>{state.promptTitle}</strong> : null}
            {state.prompt ? <span>{state.prompt}</span> : null}
          </div>
        ) : null}
      </div>

      <div id={`${tabIds}-error-panel`} role="tabpanel" aria-labelledby={`${tabIds}-error`} hidden={tab !== 'error'} className="dt-tab-panel">
        <label className="dt-check"><input type="checkbox" checked={state.showErrorMessage} onChange={(event) => set('showErrorMessage', event.target.checked)} />Show error alert after invalid data is entered</label>
        <div className="dt-error-grid">
          <label className="dt-field">
            <span>Style</span>
            <select className="dt-select" disabled={!state.showErrorMessage} value={state.errorStyle} onChange={(event) => set('errorStyle', event.target.value as ValidationErrorStyle)}>
              {(Object.keys(STYLE_META) as ValidationErrorStyle[]).map((style) => <option key={style} value={style}>{STYLE_META[style].label}</option>)}
            </select>
            <small className="dt-hint">{styleMeta.hint}</small>
          </label>
          <div className="dt-error-fields">
            <label className="dt-field">
              <span>Title</span>
              <input className="dt-input" maxLength={32} disabled={!state.showErrorMessage} value={state.errorTitle} onChange={(event) => set('errorTitle', event.target.value)} />
            </label>
            <label className="dt-field">
              <span>Error message <em className="dt-counter">{state.error.length}/255</em></span>
              <textarea className="dt-textarea" maxLength={255} rows={3} disabled={!state.showErrorMessage} value={state.error} onChange={(event) => set('error', event.target.value)} />
            </label>
          </div>
        </div>
        {state.showErrorMessage ? (
          <div className={`dt-alert-preview is-${state.errorStyle}`} aria-label="Error alert preview">
            <span className="dt-alert-icon" aria-hidden="true">{styleMeta.icon}</span>
            <div>
              <strong>{state.errorTitle || (state.errorStyle === 'stop' ? 'Invalid entry' : state.errorStyle === 'warning' ? 'Check this entry' : 'Please note')}</strong>
              <span>{state.error || (conversion.ok ? requirementText(conversion.validation) : '') || DEFAULT_VALIDATION_ERROR}</span>
              <em>{state.errorStyle === 'stop' ? 'Retry · Cancel' : state.errorStyle === 'warning' ? 'Continue? Yes · No · Cancel' : 'OK · Cancel'}</em>
            </div>
          </div>
        ) : (
          <p className="dt-hint">With the alert off, invalid entries are accepted silently; use Circle invalid data to find them later.</p>
        )}
      </div>
    </DataToolsDialog>
  )
}
