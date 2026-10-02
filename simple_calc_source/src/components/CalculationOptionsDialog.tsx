import { useEffect, useRef, useState } from 'react'
import type { CalculationOptions } from '../lib/calc-engine'
import './app-polish.css'

export interface CalculationOptionsDialogProps {
  options: CalculationOptions
  onApply: (options: CalculationOptions) => void
  onClose: () => void
}

/**
 * Excel's File > Options > Formulas (Sheets' Settings > Calculation): automatic or manual
 * calculation, and iterative calculation for circular references. Saved with the workbook.
 */
export function CalculationOptionsDialog({ options, onApply, onClose }: CalculationOptionsDialogProps) {
  const [mode, setMode] = useState(options.mode)
  const [iterate, setIterate] = useState(options.iterate)
  const [iterations, setIterations] = useState(String(options.maxIterations))
  const [maxChange, setMaxChange] = useState(String(options.maxChange))
  const [error, setError] = useState('')
  const firstRef = useRef<HTMLInputElement>(null)
  useEffect(() => { firstRef.current?.focus() }, [])

  const submit = () => {
    const count = Number(iterations.trim())
    const change = Number(maxChange.trim().replace(',', '.'))
    if (iterate && (!Number.isInteger(count) || count < 1 || count > 32_767)) { setError('Maximum iterations must be a whole number from 1 to 32,767.'); return }
    if (iterate && (!Number.isFinite(change) || change < 0)) { setError('Maximum change must be a number of 0 or more.'); return }
    onApply({
      mode,
      iterate,
      maxIterations: Number.isInteger(count) && count >= 1 ? Math.min(32_767, count) : options.maxIterations,
      maxChange: Number.isFinite(change) && change >= 0 ? change : options.maxChange,
    })
  }

  return (
    <div className="prompt-overlay" onPointerDown={(event) => { if (event.target === event.currentTarget) onClose() }}>
      <form
        className="prompt-card calc-options-card"
        role="dialog"
        aria-modal="true"
        aria-label="Calculation options"
        onSubmit={(event) => { event.preventDefault(); submit() }}
        onKeyDown={(event) => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onClose() } }}
      >
        <h2>Calculation options</h2>
        <fieldset>
          <legend>Workbook calculation</legend>
          <label className="calc-choice">
            <input ref={mode === 'automatic' ? firstRef : undefined} type="radio" name="calc-mode" checked={mode === 'automatic'} onChange={() => setMode('automatic')} />
            <span>Automatic<small>Formulas update as soon as the cells they use change.</small></span>
          </label>
          <label className="calc-choice">
            <input ref={mode === 'manual' ? firstRef : undefined} type="radio" name="calc-mode" checked={mode === 'manual'} onChange={() => setMode('manual')} />
            <span>Manual<small>Formulas update when you press F9 (Shift+F9 for the active sheet).</small></span>
          </label>
        </fieldset>
        <fieldset>
          <legend>Circular references</legend>
          <label className="calc-choice">
            <input type="checkbox" checked={iterate} onChange={(event) => { setIterate(event.target.checked); setError('') }} />
            <span>Enable iterative calculation<small>Formulas that refer to their own cell are repeated until the result settles.</small></span>
          </label>
          <div className="calc-fields">
            <label>
              <span>Maximum iterations</span>
              <input inputMode="numeric" aria-label="Maximum iterations" disabled={!iterate} value={iterations} onChange={(event) => { setIterations(event.target.value); setError('') }} />
            </label>
            <label>
              <span>Maximum change</span>
              <input inputMode="decimal" aria-label="Maximum change" disabled={!iterate} value={maxChange} onChange={(event) => { setMaxChange(event.target.value); setError('') }} />
            </label>
          </div>
        </fieldset>
        {error && <p className="calc-options-error" role="alert">{error}</p>}
        <div className="prompt-actions">
          <button type="button" className="secondary-action" onClick={onClose}>Cancel</button>
          <button type="submit" className="primary-action">OK</button>
        </div>
      </form>
    </div>
  )
}
