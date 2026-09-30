import { Fragment, useCallback, useLayoutEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { MoreHorizontal } from 'lucide-react'
import { FormatPopover } from './ColorPicker'
import './toolbar.css'
import type { AnchorRect } from './ColorPicker'

export interface CommandGroup {
  id: string
  content: ReactNode
}

interface CommandBarProps {
  leading: ReactNode
  groups: CommandGroup[]
  trailing?: ReactNode
}

const MORE_BUTTON_WIDTH = 40
const SEPARATOR_WIDTH = 17

/**
 * The formatting toolbar. Groups that do not fit move, in order, into a "More" popover so
 * every command stays reachable at any window width (like Google Sheets' toolbar).
 */
export function CommandBar({ leading, groups, trailing }: CommandBarProps) {
  const barRef = useRef<HTMLDivElement>(null)
  const leadingRef = useRef<HTMLDivElement>(null)
  const trailingRef = useRef<HTMLDivElement>(null)
  const widths = useRef(new Map<string, number>())
  const [visibleCount, setVisibleCount] = useState(groups.length)
  const [moreAnchor, setMoreAnchor] = useState<AnchorRect | null>(null)
  const moreRef = useRef<HTMLButtonElement>(null)

  const recompute = useCallback(() => {
    const bar = barRef.current
    if (!bar) return
    for (const element of bar.querySelectorAll<HTMLElement>('[data-command-group]')) {
      const id = element.dataset.commandGroup!
      if (element.offsetWidth) widths.current.set(id, element.offsetWidth)
    }
    const available = bar.clientWidth - (leadingRef.current?.offsetWidth || 0) - (trailingRef.current?.offsetWidth || 0) - 24
    let used = 0
    let count = 0
    for (const group of groups) {
      const width = (widths.current.get(group.id) ?? 120) + SEPARATOR_WIDTH
      const remaining = groups.length - count - 1
      if (used + width + (remaining > 0 ? MORE_BUTTON_WIDTH : 0) > available) break
      used += width
      count += 1
    }
    setVisibleCount((current) => (current === count ? current : count))
  }, [groups])

  useLayoutEffect(() => {
    recompute()
  })

  useLayoutEffect(() => {
    const bar = barRef.current
    if (!bar) return
    const observer = new ResizeObserver(() => recompute())
    observer.observe(bar)
    return () => observer.disconnect()
  }, [recompute])

  const visible = groups.slice(0, visibleCount)
  const hidden = groups.slice(visibleCount)
  return (
    <div className="command-bar" ref={barRef} role="toolbar" aria-label="Formatting">
      <div className="command-group file-group" ref={leadingRef}>{leading}</div>
      {visible.map((group) => (
        <Fragment key={group.id}>
          <span className="command-separator" />
          <div className="command-group" data-command-group={group.id}>{group.content}</div>
        </Fragment>
      ))}
      {hidden.length > 0 && (
        <>
          <span className="command-separator" />
          <button
            ref={moreRef}
            type="button"
            className={`tool-button command-more${moreAnchor ? ' is-active' : ''}`}
            aria-label="More tools"
            title="More tools"
            aria-haspopup="true"
            aria-expanded={Boolean(moreAnchor)}
            onClick={() => {
              if (moreAnchor) { setMoreAnchor(null); return }
              const rect = moreRef.current?.getBoundingClientRect()
              if (rect) setMoreAnchor({ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom })
            }}
          ><MoreHorizontal size={16} /></button>
        </>
      )}
      <div className="command-spacer" />
      {trailing && <div className="command-trailing" ref={trailingRef}>{trailing}</div>}
      {moreAnchor && hidden.length > 0 && (
        <FormatPopover anchor={moreAnchor} onClose={() => setMoreAnchor(null)} label="More tools" ignoreElement={moreRef.current} className="command-more-popover" align="end">
          <div className="command-more-panel">
            {hidden.map((group) => <div key={group.id} className="command-group">{group.content}</div>)}
          </div>
        </FormatPopover>
      )}
    </div>
  )
}
