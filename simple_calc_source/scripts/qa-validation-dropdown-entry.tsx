// Browser entry for scripts/qa-validation-dropdown.cjs: the real in-cell list dropdown with the
// presentations App.tsx passes (Excel's single choice, Sheets' coloured chips with several
// picks). Calls are recorded on window.harness.
import { useState } from 'react'
import { createRoot } from 'react-dom/client'
import '../src/styles.css'
import { ValidationDropdown } from '../src/components/ValidationDropdown'
import { dropdownPresentation } from '../src/lib/validation'

interface Harness {
  picks: string[]
  many: string[][]
  closed: number
  mode: string
}

declare global {
  interface Window {
    harness: Harness
  }
}

const mode = new URLSearchParams(window.location.search).get('mode') || 'single'
window.harness = { picks: [], many: [], closed: 0, mode }

const multiple = dropdownPresentation({
  type: 'list',
  formulae: ['"Open,In progress,Done"'],
  simpleDropdown: { style: 'chip', multiple: true, colors: { Done: '#b7e1cd', 'In progress': '#fce8b2' } },
})
const single = dropdownPresentation({ type: 'list', formulae: ['"Low,Medium,High"'] })

function Harness() {
  const [open, setOpen] = useState(true)
  if (!open) return <div className="closed">closed</div>
  const done = () => setOpen(false)
  return mode === 'multiple' ? (
    <ValidationDropdown
      anchor={{ left: 40, top: 40, bottom: 60, width: 120 }}
      options={['Open', 'In progress', 'Done']}
      current="Open"
      presentation={multiple}
      label="Choose values for B2"
      onPick={(value) => { window.harness.picks.push(value); done() }}
      onPickMany={(values) => { window.harness.many.push(values); done() }}
      onClose={() => { window.harness.closed += 1; done() }}
    />
  ) : (
    <ValidationDropdown
      anchor={{ left: 40, top: 40, bottom: 60, width: 120 }}
      options={['Low', 'Medium', 'High']}
      current="Medium"
      presentation={single}
      label="Choose a value for B2"
      onPick={(value) => { window.harness.picks.push(value); done() }}
      onPickMany={(values) => { window.harness.many.push(values); done() }}
      onClose={() => { window.harness.closed += 1; done() }}
    />
  )
}

createRoot(document.getElementById('root')!).render(<Harness />)
