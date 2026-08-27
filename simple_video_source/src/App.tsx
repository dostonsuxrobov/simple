import {
  Check,
  CircleAlert,
  ChevronDown,
  Expand,
  FileVideo2,
  FolderOpen,
  Gauge,
  Info,
  Keyboard,
  Maximize2,
  Minimize2,
  MoreHorizontal,
  Pause,
  Play,
  Plus,
  Square,
  Volume1,
  Volume2,
  VolumeX,
  X,
} from 'lucide-react'
import { useCallback, useEffect, useRef, useState, type CSSProperties, type DragEvent } from 'react'

const SUPPORTED_EXTENSIONS = ['mp4', 'm4v', 'webm', 'ogv', 'mov', 'mkv']
const PLAYBACK_RATES = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 1.75, 2]
type FitMode = 'contain' | 'cover' | 'actual'

function formatTime(seconds: number) {
  if (!Number.isFinite(seconds) || seconds < 0) return '0:00'
  const total = Math.floor(seconds)
  const hours = Math.floor(total / 3600)
  const minutes = Math.floor((total % 3600) / 60)
  const remaining = total % 60
  return hours
    ? `${hours}:${String(minutes).padStart(2, '0')}:${String(remaining).padStart(2, '0')}`
    : `${minutes}:${String(remaining).padStart(2, '0')}`
}

function formatBytes(bytes: number) {
  if (!Number.isFinite(bytes) || bytes < 1) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const unit = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  const value = bytes / 1024 ** unit
  return `${value.toFixed(value >= 10 || unit === 0 ? 0 : 1)} ${units[unit]}`
}

function friendlyDate(timestamp: number) {
  const value = new Date(timestamp)
  const today = new Date()
  if (value.toDateString() === today.toDateString()) {
    return value.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })
  }
  return value.toLocaleDateString([], { month: 'short', day: 'numeric', year: value.getFullYear() === today.getFullYear() ? undefined : 'numeric' })
}

function errorMessage(error: unknown) {
  if (error instanceof Error) return error.message.replace(/^Error invoking remote method '[^']+':\s*/i, '')
  return String(error || 'Something went wrong.')
}

function playbackErrorMessage(code?: number) {
  if (code === 1) return 'Playback was interrupted before it could begin.'
  if (code === 2) return 'The video could not be read. It may have moved or become unavailable.'
  if (code === 3) return 'The video appears damaged or could not be decoded.'
  return 'This video container opened, but its codec is not supported by this computer.'
}

function App() {
  const videoRef = useRef<HTMLVideoElement>(null)
  const hideControlsTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const dragDepth = useRef(0)
  const [file, setFile] = useState<VideoFilePayload | null>(null)
  const [openGeneration, setOpenGeneration] = useState(0)
  const [recents, setRecents] = useState<RecentVideo[]>([])
  const [loading, setLoading] = useState(false)
  const [playing, setPlaying] = useState(false)
  const [currentTime, setCurrentTime] = useState(0)
  const [duration, setDuration] = useState(0)
  const [volume, setVolume] = useState(0.8)
  const [muted, setMuted] = useState(false)
  const [rate, setRate] = useState(1)
  const [fit, setFit] = useState<FitMode>('contain')
  const [resolution, setResolution] = useState({ width: 0, height: 0 })
  const [dragging, setDragging] = useState(false)
  const [maximized, setMaximized] = useState(false)
  const [fullscreen, setFullscreen] = useState(false)
  const [controlsVisible, setControlsVisible] = useState(true)
  const [infoOpen, setInfoOpen] = useState(false)
  const [moreOpen, setMoreOpen] = useState(false)
  const [shortcutsOpen, setShortcutsOpen] = useState(false)
  const [toast, setToast] = useState<string | null>(null)
  const [mediaError, setMediaError] = useState<string | null>(null)

  const refreshRecents = useCallback(async () => {
    try {
      setRecents(await window.simpleVideo.listRecents())
    } catch {
      setRecents([])
    }
  }, [])

  const adoptFile = useCallback((payload: VideoFilePayload) => {
    setLoading(true)
    setPlaying(false)
    setCurrentTime(0)
    setDuration(0)
    setResolution({ width: 0, height: 0 })
    setMediaError(null)
    setInfoOpen(false)
    setMoreOpen(false)
    setFile(payload)
    setOpenGeneration((value) => value + 1)
    window.simpleVideo.setTitle(`${payload.name} — simple video`)
    void refreshRecents()
  }, [refreshRecents])

  const openPath = useCallback(async (filePath: string) => {
    try {
      adoptFile(await window.simpleVideo.openPath(filePath))
    } catch (error) {
      setLoading(false)
      setToast(errorMessage(error))
    }
  }, [adoptFile])

  const chooseFile = useCallback(async () => {
    try {
      const payload = await window.simpleVideo.openFile()
      if (payload) adoptFile(payload)
    } catch (error) {
      setToast(errorMessage(error))
    }
  }, [adoptFile])

  const togglePlay = useCallback(() => {
    const video = videoRef.current
    if (!video || !file) return
    if (video.paused) {
      void video.play().catch((error) => setToast(errorMessage(error)))
    } else {
      video.pause()
    }
  }, [file])

  const seekBy = useCallback((delta: number) => {
    const video = videoRef.current
    if (!video || !Number.isFinite(video.duration)) return
    video.currentTime = Math.max(0, Math.min(video.duration, video.currentTime + delta))
  }, [])

  const changeVolume = useCallback((nextVolume: number) => {
    const next = Math.max(0, Math.min(1, nextVolume))
    setVolume(next)
    if (next > 0) setMuted(false)
  }, [])

  const selectRate = useCallback((nextRate: number) => {
    const safeRate = PLAYBACK_RATES.reduce((best, candidate) => Math.abs(candidate - nextRate) < Math.abs(best - nextRate) ? candidate : best, 1)
    setRate(safeRate)
    if (videoRef.current) videoRef.current.playbackRate = safeRate
  }, [])

  const cycleFit = useCallback(() => {
    setFit((value) => value === 'contain' ? 'cover' : value === 'cover' ? 'actual' : 'contain')
  }, [])

  const revealControls = useCallback(() => {
    setControlsVisible(true)
    if (hideControlsTimer.current) clearTimeout(hideControlsTimer.current)
    if (playing && file) {
      hideControlsTimer.current = setTimeout(() => setControlsVisible(false), 2600)
    }
  }, [file, playing])

  useEffect(() => {
    void refreshRecents()
    const stopOpen = window.simpleVideo.onOpenExternal((filePath) => void openPath(filePath))
    const stopMaximized = window.simpleVideo.onMaximized(setMaximized)
    const stopFullscreen = window.simpleVideo.onFullscreen(setFullscreen)
    return () => {
      stopOpen()
      stopMaximized()
      stopFullscreen()
    }
  }, [openPath, refreshRecents])

  useEffect(() => {
    const video = videoRef.current
    if (!video) return
    video.volume = volume
    video.muted = muted
    video.playbackRate = rate
  }, [muted, rate, volume, file])

  useEffect(() => {
    revealControls()
    return () => {
      if (hideControlsTimer.current) clearTimeout(hideControlsTimer.current)
    }
  }, [playing, revealControls])

  useEffect(() => {
    if (!toast) return
    const timer = setTimeout(() => setToast(null), 4200)
    return () => clearTimeout(timer)
  }, [toast])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null
      const key = event.key.toLowerCase()
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'o') {
        event.preventDefault()
        void chooseFile()
        return
      }
      if (key === 'escape' && fullscreen) {
        event.preventDefault()
        window.simpleVideo.exitFullscreen()
        return
      }
      if (shortcutsOpen) {
        if (key === 'escape') setShortcutsOpen(false)
        return
      }
      if (target?.closest('input, select, textarea')) return
      if (target?.closest('button') && (event.code === 'Space' || key === 'enter')) return
      if (event.key === '?' || (event.key === '/' && event.shiftKey)) {
        event.preventDefault()
        setShortcutsOpen(true)
        return
      }
      if (!file) return
      if (event.code === 'Space' || key === 'k') {
        event.preventDefault()
        togglePlay()
      } else if (key === 'arrowleft') {
        event.preventDefault()
        seekBy(-5)
      } else if (key === 'arrowright') {
        event.preventDefault()
        seekBy(5)
      } else if (key === 'j') {
        seekBy(-10)
      } else if (key === 'l') {
        seekBy(10)
      } else if (key === 'arrowup') {
        event.preventDefault()
        changeVolume(volume + 0.05)
      } else if (key === 'arrowdown') {
        event.preventDefault()
        changeVolume(volume - 0.05)
      } else if (key === 'm') {
        setMuted((value) => !value)
      } else if (key === 'f') {
        window.simpleVideo.toggleFullscreen()
      } else if (key === 'home') {
        event.preventDefault()
        if (videoRef.current) videoRef.current.currentTime = 0
      } else if (key === 'end') {
        event.preventDefault()
        if (videoRef.current && Number.isFinite(videoRef.current.duration)) videoRef.current.currentTime = videoRef.current.duration
      } else if (key === '[') {
        selectRate(rate - 0.25)
      } else if (key === ']') {
        selectRate(rate + 0.25)
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [changeVolume, chooseFile, file, fullscreen, rate, seekBy, selectRate, shortcutsOpen, togglePlay, volume])

  const handleDragEnter = (event: DragEvent) => {
    event.preventDefault()
    dragDepth.current += 1
    setDragging(true)
  }

  const handleDragLeave = (event: DragEvent) => {
    event.preventDefault()
    dragDepth.current = Math.max(0, dragDepth.current - 1)
    if (!dragDepth.current) setDragging(false)
  }

  const handleDrop = async (event: DragEvent) => {
    event.preventDefault()
    dragDepth.current = 0
    setDragging(false)
    const paths = Array.from(event.dataTransfer.files)
      .map((item) => window.simpleVideo.pathForFile(item))
      .filter((item): item is string => Boolean(item))
      .filter((item) => SUPPORTED_EXTENSIONS.some((extension) => item.toLowerCase().endsWith(`.${extension}`)))
    if (!paths.length) {
      setToast('Drop a supported video file here.')
      return
    }
    await openPath(paths[0])
    if (paths.length > 1) {
      const results = await Promise.allSettled(paths.slice(1).map((item) => window.simpleVideo.openInNewWindow(item)))
      if (results.some((result) => result.status === 'rejected')) {
        setToast('One or more additional videos could not be opened.')
      }
    }
  }

  const removeRecent = async (event: React.MouseEvent, filePath: string) => {
    event.stopPropagation()
    try {
      setRecents(await window.simpleVideo.removeRecent(filePath))
    } catch (error) {
      setToast(errorMessage(error))
    }
  }

  const clearCurrent = () => {
    videoRef.current?.pause()
    if (fullscreen) window.simpleVideo.exitFullscreen()
    setFile(null)
    setPlaying(false)
    setCurrentTime(0)
    setDuration(0)
    setMediaError(null)
    setMoreOpen(false)
    setInfoOpen(false)
    window.simpleVideo.setTitle('simple video')
    void refreshRecents()
  }

  const fitLabel = fit === 'contain' ? 'Fit' : fit === 'cover' ? 'Fill' : 'Actual size'
  const volumeIcon = muted || volume === 0 ? <VolumeX /> : volume < 0.55 ? <Volume1 /> : <Volume2 />
  const progress = duration > 0 ? Math.min(100, (currentTime / duration) * 100) : 0
  const volumeProgress = muted ? 0 : volume * 100

  return (
    <main
      className={`app-shell${fullscreen ? ' is-fullscreen' : ''}`}
      onDragEnter={handleDragEnter}
      onDragOver={(event) => event.preventDefault()}
      onDragLeave={handleDragLeave}
      onDrop={(event) => void handleDrop(event)}
    >
      {!fullscreen && (
        <header className="titlebar">
          <div className="title-identity drag-region">
            <span className="app-mark"><FileVideo2 /></span>
            <span className="app-name">simple video</span>
            {file && <><span className="title-divider" /><span className="document-title">{file.name}</span></>}
          </div>
          <div className="title-actions">
            <button type="button" onClick={() => void chooseFile()} title="Open video (Ctrl+O)"><FolderOpen /></button>
            <button type="button" onClick={() => window.simpleVideo.newWindow()} title="New window"><Plus /></button>
          </div>
          <div className="window-actions">
            <button type="button" onClick={() => window.simpleVideo.minimize()} aria-label="Minimize"><Minimize2 /></button>
            <button type="button" onClick={() => window.simpleVideo.toggleMaximize()} aria-label={maximized ? 'Restore' : 'Maximize'}>{maximized ? <Square /> : <Maximize2 />}</button>
            <button className="close-button" type="button" onClick={() => window.simpleVideo.close()} aria-label="Close"><X /></button>
          </div>
        </header>
      )}

      <section className="workspace">
        {!file ? (
          <Welcome
            recents={recents}
            onOpen={() => void chooseFile()}
            onOpenRecent={(filePath) => void openPath(filePath)}
            onRemoveRecent={removeRecent}
            onClearRecents={async () => {
              try {
                setRecents(await window.simpleVideo.clearRecents())
              } catch (error) {
                setToast(errorMessage(error))
              }
            }}
            onShowShortcuts={() => setShortcutsOpen(true)}
          />
        ) : (
          <div className={`player-shell${controlsVisible ? '' : ' controls-hidden'}`} onMouseMove={revealControls} onMouseLeave={() => playing && setControlsVisible(false)}>
            <div className="video-stage" onDoubleClick={() => window.simpleVideo.toggleFullscreen()}>
              <video
                key={openGeneration}
                ref={videoRef}
                className={`video-element fit-${fit}`}
                src={file.url}
                preload="metadata"
                playsInline
                onClick={togglePlay}
                onLoadStart={() => setLoading(true)}
                onLoadedMetadata={(event) => {
                  const video = event.currentTarget
                  setDuration(Number.isFinite(video.duration) ? video.duration : 0)
                  setResolution({ width: video.videoWidth, height: video.videoHeight })
                  setMediaError(null)
                  setLoading(false)
                }}
                onCanPlay={() => setLoading(false)}
                onTimeUpdate={(event) => setCurrentTime(event.currentTarget.currentTime)}
                onDurationChange={(event) => setDuration(Number.isFinite(event.currentTarget.duration) ? event.currentTarget.duration : 0)}
                onPlay={() => setPlaying(true)}
                onPause={() => setPlaying(false)}
                onEnded={() => setPlaying(false)}
                onError={(event) => {
                  setLoading(false)
                  const message = playbackErrorMessage(event.currentTarget.error?.code)
                  setMediaError(message)
                }}
              />
              {loading && (
                <div className="loading-indicator" aria-live="polite">
                  <span className="spinner" />
                  <strong>Opening video</strong>
                </div>
              )}
              {mediaError && (
                <div className="media-error" role="alert">
                  <CircleAlert />
                  <strong>Couldn’t play this video</strong>
                  <span>{mediaError}</span>
                  <button type="button" onClick={() => void chooseFile()}>Open another video</button>
                </div>
              )}
              {!playing && !loading && !mediaError && (
                <button className="center-play" type="button" onClick={togglePlay} aria-label="Play video"><Play /></button>
              )}
            </div>

            <div className="player-toolbar top-toolbar">
              <div className="file-chip"><FileVideo2 /><span>{file.name}</span></div>
              <div className="toolbar-spacer" />
              <button type="button" className={fit !== 'contain' ? 'is-active' : ''} onClick={cycleFit} title="Cycle fit mode"><Expand /><span>{fitLabel}</span></button>
              <button type="button" className={infoOpen ? 'is-active' : ''} onClick={() => { setInfoOpen((value) => !value); setMoreOpen(false) }} title="Video information"><Info /></button>
              <div className="menu-anchor">
                <button type="button" className={moreOpen ? 'is-active' : ''} onClick={() => { setMoreOpen((value) => !value); setInfoOpen(false) }} title="More"><MoreHorizontal /></button>
                {moreOpen && (
                  <div className="popup-menu compact-menu">
                    <button type="button" onClick={() => void chooseFile()}><FolderOpen /><span>Open another video</span><kbd>Ctrl O</kbd></button>
                    <button type="button" onClick={() => window.simpleVideo.openInNewWindow(file.path)}><Plus /><span>Open in new window</span></button>
                    <button type="button" onClick={() => window.simpleVideo.revealFile(file.path)}><FileVideo2 /><span>Show in folder</span></button>
                    <button type="button" onClick={() => setShortcutsOpen(true)}><Keyboard /><span>Keyboard shortcuts</span><kbd>?</kbd></button>
                    <div className="menu-divider" />
                    <button type="button" onClick={clearCurrent}><X /><span>Close video</span></button>
                  </div>
                )}
              </div>
            </div>

            {infoOpen && (
              <aside className="info-panel">
                <div className="info-heading"><span><Info /> Video information</span><button type="button" onClick={() => setInfoOpen(false)}><X /></button></div>
                <div className="info-preview"><FileVideo2 /><span>{file.extension}</span></div>
                <h2 title={file.name}>{file.name}</h2>
                <dl>
                  <div><dt>Duration</dt><dd>{formatTime(duration)}</dd></div>
                  <div><dt>Resolution</dt><dd>{resolution.width ? `${resolution.width} × ${resolution.height}` : 'Reading…'}</dd></div>
                  <div><dt>File size</dt><dd>{formatBytes(file.size)}</dd></div>
                  <div><dt>Format</dt><dd>{file.extension}</dd></div>
                  <div><dt>Modified</dt><dd>{new Date(file.modifiedAt).toLocaleString()}</dd></div>
                </dl>
                <button className="reveal-button" type="button" onClick={() => window.simpleVideo.revealFile(file.path)}><FolderOpen /> Show in folder</button>
              </aside>
            )}

            <div className="controls-panel" onMouseEnter={revealControls}>
              <input
                className="seek-range"
                style={{ '--progress': `${progress}%` } as CSSProperties}
                type="range"
                min="0"
                max={duration || 1}
                step="0.01"
                value={Math.min(currentTime, duration || 0)}
                aria-label="Video position"
                onChange={(event) => {
                  const value = Number(event.target.value)
                  if (videoRef.current) videoRef.current.currentTime = value
                  setCurrentTime(value)
                }}
              />
              <div className="control-row">
                <button className="play-button" type="button" onClick={togglePlay} aria-label={playing ? 'Pause' : 'Play'}>{playing ? <Pause /> : <Play />}</button>
                <button type="button" onClick={() => setMuted((value) => !value)} aria-label={muted ? 'Unmute' : 'Mute'}>{volumeIcon}</button>
                <input
                  className="volume-range"
                  style={{ '--progress': `${volumeProgress}%` } as CSSProperties}
                  type="range"
                  min="0"
                  max="1"
                  step="0.01"
                  value={volume}
                  aria-label="Volume"
                  onChange={(event) => changeVolume(Number(event.target.value))}
                />
                <span className="time-readout">{formatTime(currentTime)} <i>/</i> {formatTime(duration)}</span>
                <span className="control-spacer" />
                <label className="rate-control" title="Playback speed"><Gauge /><select value={rate} onChange={(event) => selectRate(Number(event.target.value))}>{PLAYBACK_RATES.map((value) => <option key={value} value={value}>{value}×</option>)}</select><ChevronDown /></label>
                <button type="button" onClick={cycleFit} title={`Display: ${fitLabel}`}><Expand /></button>
                <button type="button" onClick={() => window.simpleVideo.toggleFullscreen()} title="Fullscreen (F)"><Maximize2 /></button>
              </div>
            </div>
          </div>
        )}

        {dragging && (
          <div className="drop-overlay">
            <div><FileVideo2 /><strong>Drop to open</strong><span>Additional videos open in their own windows</span></div>
          </div>
        )}
      </section>

      {shortcutsOpen && <Shortcuts onClose={() => setShortcutsOpen(false)} />}
      {toast && <div className="toast" role="status">{toast}<button type="button" onClick={() => setToast(null)}><X /></button></div>}
    </main>
  )
}

interface WelcomeProps {
  recents: RecentVideo[]
  onOpen: () => void
  onOpenRecent: (filePath: string) => void
  onRemoveRecent: (event: React.MouseEvent, filePath: string) => void
  onClearRecents: () => void
  onShowShortcuts: () => void
}

function Welcome({ recents, onOpen, onOpenRecent, onRemoveRecent, onClearRecents, onShowShortcuts }: WelcomeProps) {
  return (
    <div className="welcome">
      <div className="welcome-card">
        <div className="welcome-brand">
          <div className="welcome-mark"><FileVideo2 /></div>
          <div><h1>simple video</h1><p>Watch videos privately, right on this computer.</p></div>
        </div>
        <div className="start-actions">
          <button className="start-card primary-start" type="button" onClick={onOpen}>
            <span className="start-icon"><FolderOpen /></span><span><strong>Open a video</strong><small>Choose a file from this computer</small></span>
          </button>
          <button className="start-card" type="button" onClick={onShowShortcuts}>
            <span className="start-icon"><Keyboard /></span><span><strong>Playback shortcuts</strong><small>Play, seek, volume, speed and fullscreen</small></span>
          </button>
        </div>
        <section className="welcome-section">
          <div className="section-heading"><h2>Recent videos</h2>{recents.length > 0 && <button type="button" onClick={onClearRecents}>Clear list</button>}</div>
          <div className="recent-list">
            {recents.length ? recents.map((item) => (
              <button className="recent-row" type="button" key={item.path} onClick={() => onOpenRecent(item.path)}>
                <span className="recent-icon"><FileVideo2 /></span>
                <span className="recent-copy"><strong>{item.name}</strong><small>{item.path}</small></span>
                <time>{friendlyDate(item.openedAt)}</time>
                <span className="remove-recent" role="button" tabIndex={0} title="Remove from recent videos" onClick={(event) => onRemoveRecent(event, item.path)} onKeyDown={(event) => { if (event.key === 'Enter') onRemoveRecent(event as unknown as React.MouseEvent, item.path) }}><X /></span>
              </button>
            )) : <div className="empty-recents"><FileVideo2 /><strong>No recent videos yet</strong><span>Open a video or drop one anywhere in this window.</span></div>}
          </div>
        </section>
        <p className="drop-hint">Drop MP4, M4V, WebM, OGV, MOV or MKV files anywhere to open them.</p>
      </div>
    </div>
  )
}

function Shortcuts({ onClose }: { onClose: () => void }) {
  const rows = [
    ['Play or pause', 'Space / K'],
    ['Seek 5 seconds', '← / →'],
    ['Seek 10 seconds', 'J / L'],
    ['Volume', '↑ / ↓'],
    ['Mute', 'M'],
    ['Playback speed', '[ / ]'],
    ['Beginning / end', 'Home / End'],
    ['Fullscreen', 'F'],
    ['Open video', 'Ctrl O'],
  ]
  return (
    <div className="modal-backdrop" onMouseDown={(event) => event.currentTarget === event.target && onClose()}>
      <section className="shortcuts-modal" role="dialog" aria-modal="true" aria-labelledby="shortcuts-title">
        <div className="modal-heading"><span className="modal-mark"><Keyboard /></span><div><h2 id="shortcuts-title">Keyboard shortcuts</h2><p>Control playback without reaching for the mouse.</p></div><button type="button" onClick={onClose}><X /></button></div>
        <div className="shortcut-list">{rows.map(([action, keys]) => <div key={action}><span>{action}</span><kbd>{keys}</kbd></div>)}</div>
        <button className="done-button" type="button" onClick={onClose}><Check /> Done</button>
      </section>
    </div>
  )
}

export default App
