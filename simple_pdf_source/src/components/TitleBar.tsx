import { useEffect, useState } from 'react'
import { Minus, Square, Copy, X } from 'lucide-react'
import { IconButton } from './ui'

interface TitleBarProps {
  fileName?: string
  dirty?: boolean
  onClose: () => void
}

export function TitleBar({ fileName, dirty, onClose }: TitleBarProps) {
  const [maximized, setMaximized] = useState(false)

  useEffect(() => window.simple.onMaximized(setMaximized), [])

  return (
    <header className="titlebar">
      <div className="titlebar-drag">
        <span className="app-title">simple</span>
        {fileName && (
          <>
            <span className="title-divider">/</span>
            <span className="document-title" title={fileName}>{fileName}</span>
            {dirty && <span className="dirty-dot" aria-label="Unsaved changes" title="Unsaved changes" />}
          </>
        )}
      </div>
      <div className="window-controls">
        <IconButton icon={Minus} label="Minimize" onClick={() => window.simple.minimize()} />
        <IconButton icon={maximized ? Copy : Square} label={maximized ? 'Restore' : 'Maximize'} onClick={() => window.simple.toggleMaximize()} />
        <IconButton icon={X} label="Close" className="window-close" onClick={onClose} />
      </div>
    </header>
  )
}
