import { forwardRef, memo, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { CSSProperties } from 'react'
import { FunctionSquare, Hash, Table2 } from 'lucide-react'
import './formula-assist.css'

export interface AssistFunction {
  name: string
  description: string
  signature: string
  args: Array<{ name: string; description?: string; optional?: boolean; repeating?: boolean }>
  category?: string
}

export interface AssistItem {
  kind: 'function' | 'name' | 'table'
  label: string
  detail: string
  insertText: string
}

export interface FormulaAssistState {
  items: AssistItem[]
  activeIndex: number
  call: { fn: AssistFunction; argumentIndex: number } | null
}

interface FormulaAssistProps {
  anchor: { left: number; top: number; bottom: number; width: number } | null
  state: FormulaAssistState
  onPick: (item: AssistItem) => void
  onHover: (index: number) => void
}

function signatureParts(fn: AssistFunction, activeIndex: number) {
  // Map a repeating last argument (number1, [number2], ...) onto any later index.
  const args = fn.args
  let highlighted = activeIndex
  if (args.length && activeIndex >= args.length) {
    const repeatingIndex = args.findIndex((arg) => arg.repeating)
    highlighted = repeatingIndex >= 0 ? Math.min(args.length - 1, repeatingIndex + ((activeIndex - repeatingIndex) > 0 ? 1 : 0)) : -1
  }
  return { args, highlighted }
}

export const FormulaAssist = memo(function FormulaAssist({ anchor, state, onPick, onHover }: FormulaAssistProps) {
  const listRef = useRef<HTMLDivElement>(null)
  const [position, setPosition] = useState<{ left: number; top: number; above: boolean } | null>(null)

  useLayoutEffect(() => {
    if (!anchor) { setPosition(null); return }
    const height = (state.items.length ? Math.min(8, state.items.length) * 30 + 40 : 0) + (state.call ? 70 : 0)
    const above = anchor.bottom + height + 8 > window.innerHeight && anchor.top - height - 8 > 0
    setPosition({
      left: Math.max(4, Math.min(anchor.left, window.innerWidth - 380)),
      top: above ? anchor.top - 4 : anchor.bottom + 4,
      above,
    })
  }, [anchor, state.call, state.items.length])

  useEffect(() => {
    listRef.current?.querySelector<HTMLElement>('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' })
  }, [state.activeIndex])

  if (!anchor || !position || (!state.items.length && !state.call)) return null
  const active = state.items[state.activeIndex]
  const call = state.call
  const signature = call ? signatureParts(call.fn, call.argumentIndex) : null
  const activeArgument = call && signature && signature.highlighted >= 0 ? signature.args[signature.highlighted] : null

  return (
    <div
      className={`fa-popover${position.above ? ' is-above' : ''}`}
      style={{ left: position.left, top: position.top }}
      onMouseDown={(event) => event.preventDefault()}
      role="presentation"
    >
      {call && !state.items.length && signature && (
        <div className="fa-signature" role="tooltip">
          <div className="fa-signature-line">
            <strong>{call.fn.name}</strong>(
            {signature.args.map((arg, index) => (
              <span key={`${arg.name}-${index}`}>
                {index > 0 && ', '}
                <span className={index === signature.highlighted ? 'fa-arg is-active' : 'fa-arg'}>
                  {arg.optional ? `[${arg.name}]` : arg.name}{arg.repeating ? ', …' : ''}
                </span>
              </span>
            ))}
            )
          </div>
          {activeArgument?.description
            ? <p className="fa-signature-help"><span>{activeArgument.name}</span> {activeArgument.description}</p>
            : call.fn.description && <p className="fa-signature-help">{call.fn.description}</p>}
        </div>
      )}
      {state.items.length > 0 && (
        <>
          <div className="fa-list" role="listbox" aria-label="Formula suggestions" ref={listRef}>
            {state.items.map((item, index) => (
              <div
                key={`${item.kind}-${item.label}`}
                role="option"
                aria-selected={index === state.activeIndex}
                className={`fa-item${index === state.activeIndex ? ' is-active' : ''}`}
                onMouseEnter={() => onHover(index)}
                onMouseDown={(event) => { event.preventDefault(); onPick(item) }}
              >
                <span className="fa-icon" aria-hidden="true">
                  {item.kind === 'function' ? <FunctionSquare size={13} /> : item.kind === 'table' ? <Table2 size={13} /> : <Hash size={13} />}
                </span>
                <span className="fa-label">{item.label}</span>
              </div>
            ))}
          </div>
          {active && (
            <div className="fa-detail">
              <p>{active.detail}</p>
              <span className="fa-hint">Tab to insert · ↑↓ to choose</span>
            </div>
          )}
        </>
      )}
    </div>
  )
})

/** Syntax-coloured mirror of a formula, drawn over a transparent-text editor. */
export const FormulaHighlight = memo(forwardRef<HTMLDivElement, {
  parts: Array<{ text: string; color?: string }>
  className?: string
  style?: CSSProperties
}>(function FormulaHighlight({ parts, className, style }, ref) {
  return (
    <div ref={ref} className={`fa-highlight${className ? ` ${className}` : ''}`} style={style} aria-hidden="true">
      {parts.map((part, index) => (
        <span key={index} style={part.color ? { color: part.color } : undefined}>{part.text}</span>
      ))}
      {'​'}
    </div>
  )
}))
