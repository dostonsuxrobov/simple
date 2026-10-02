import { Fragment, useEffect, useMemo, useRef, useState } from 'react'
import { Search, X } from 'lucide-react'
import { SHORTCUT_GROUPS, filterShortcuts } from '../lib/shortcuts'
import './app-polish.css'

function KeyCombo({ keys }: { keys: string }) {
  const alternatives = keys.split(' or ')
  return (
    <span className="shortcut-keys">
      {alternatives.map((alternative, index) => (
        <Fragment key={alternative}>
          {index > 0 && <span>or</span>}
          {/* "Ctrl+-" and "Ctrl+Shift+=" keep their last key even when it is a symbol. */}
          {alternative.split(/\+(?=.)/).map((key, keyIndex) => <kbd key={`${key}-${keyIndex}`}>{key}</kbd>)}
        </Fragment>
      ))}
    </span>
  )
}

/** Help > Keyboard shortcuts (Ctrl+/): every shortcut the app handles, searchable. */
export function KeyboardShortcutsDialog({ onClose }: { onClose: () => void }) {
  const [query, setQuery] = useState('')
  const searchRef = useRef<HTMLInputElement>(null)
  const groups = useMemo(() => filterShortcuts(SHORTCUT_GROUPS, query), [query])
  useEffect(() => { searchRef.current?.focus() }, [])
  return (
    <div className="prompt-overlay" onPointerDown={(event) => { if (event.target === event.currentTarget) onClose() }}>
      <section
        className="prompt-card shortcuts-card"
        role="dialog"
        aria-modal="true"
        aria-labelledby="shortcuts-title"
        onKeyDown={(event) => {
          if (event.key === 'Escape' || ((event.ctrlKey || event.metaKey) && event.key === '/')) {
            event.preventDefault()
            event.stopPropagation()
            onClose()
          }
        }}
      >
        <header>
          <h2 id="shortcuts-title">Keyboard shortcuts</h2>
          <button type="button" className="shortcuts-close" aria-label="Close keyboard shortcuts" onClick={onClose}><X size={15} /></button>
        </header>
        <label className="shortcuts-search command-search-input">
          <Search size={13} aria-hidden="true" />
          <input ref={searchRef} type="search" aria-label="Search shortcuts" placeholder="Search shortcuts" value={query} onChange={(event) => setQuery(event.target.value)} />
        </label>
        <div className="shortcuts-groups">
          {groups.map((group) => (
            <section key={group.title} className="shortcuts-group" aria-label={group.title}>
              <h3>{group.title}</h3>
              {group.items.map((item) => (
                <div key={`${group.title}-${item.keys}-${item.label}`} className="shortcut-row" data-shortcut={item.keys}>
                  <span>{item.label}</span>
                  <KeyCombo keys={item.keys} />
                </div>
              ))}
            </section>
          ))}
          {!groups.length && <p className="shortcuts-empty">No shortcuts match “{query}”.</p>}
        </div>
      </section>
    </div>
  )
}
