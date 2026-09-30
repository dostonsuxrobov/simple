import { useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { ChevronDown } from 'lucide-react'
import { FormatPopover } from './ColorPicker'
import type { AnchorRect } from './ColorPicker'

export interface ToolbarMenuItem {
  id: string
  label: string
  icon?: ReactNode
  checked?: boolean
  shortcut?: string
  disabled?: boolean
  separatorBefore?: boolean
  action: () => void
}

interface ToolbarMenuButtonProps {
  label: string
  icon: ReactNode
  items?: ToolbarMenuItem[]
  /** Custom popover content instead of an item list. Receives a close callback. */
  render?: (close: () => void, anchor: AnchorRect) => ReactNode
  active?: boolean
  disabled?: boolean
  showChevron?: boolean
  className?: string
  /** Primary click on the icon part; the chevron still opens the menu. */
  onPrimary?: () => void
  text?: string
}

/** A compact toolbar button that opens an anchored menu (or custom panel). */
export function ToolbarMenuButton({ label, icon, items, render, active, disabled, showChevron = true, className, onPrimary, text }: ToolbarMenuButtonProps) {
  const [anchor, setAnchor] = useState<AnchorRect | null>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  const open = () => {
    if (anchor) { setAnchor(null); return }
    const rect = rootRef.current?.getBoundingClientRect()
    if (rect) setAnchor({ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom })
  }
  const close = () => setAnchor(null)
  return (
    <div ref={rootRef} className={`toolbar-menu-button${onPrimary ? ' is-split' : ''}${active ? ' is-active' : ''}${className ? ` ${className}` : ''}`}>
      <button
        type="button"
        className="toolbar-menu-main"
        aria-label={label}
        title={label}
        disabled={disabled}
        aria-haspopup={onPrimary ? undefined : 'menu'}
        aria-expanded={onPrimary ? undefined : Boolean(anchor)}
        onClick={onPrimary ? onPrimary : open}
      >
        {icon}
        {text && <span className="toolbar-menu-text">{text}</span>}
        {!onPrimary && showChevron && <ChevronDown size={11} className="toolbar-menu-chevron" aria-hidden="true" />}
      </button>
      {onPrimary && (
        <button type="button" className="toolbar-menu-chevron-button" aria-label={`${label} options`} disabled={disabled} aria-haspopup="menu" aria-expanded={Boolean(anchor)} onClick={open}>
          <ChevronDown size={11} aria-hidden="true" />
        </button>
      )}
      {anchor && (
        <FormatPopover anchor={anchor} onClose={close} label={label} ignoreElement={rootRef.current} className="toolbar-menu-popover">
          {render ? render(close, anchor) : (
            <div className="toolbar-menu-list" role="menu" aria-label={label}>
              {(items || []).map((item) => (
                <button
                  type="button"
                  role="menuitemcheckbox"
                  aria-checked={item.checked ?? false}
                  key={item.id}
                  className={`toolbar-menu-item${item.separatorBefore ? ' has-separator' : ''}${item.checked ? ' is-checked' : ''}`}
                  disabled={item.disabled}
                  data-toolbar-action={item.id}
                  onClick={() => { item.action(); close() }}
                >
                  <span className={`toolbar-menu-item-icon${item.icon ? '' : ' is-empty'}`} aria-hidden="true">{item.icon}</span>
                  <span className="toolbar-menu-item-label">{item.label}</span>
                  {item.shortcut && <span className="toolbar-menu-item-shortcut">{item.shortcut}</span>}
                </button>
              ))}
            </div>
          )}
        </FormatPopover>
      )}
    </div>
  )
}
