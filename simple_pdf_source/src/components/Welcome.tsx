import { FileImage, FileText, FileType2, FolderOpen, ShieldCheck, X } from 'lucide-react'
import type { DragEvent } from 'react'
import type { RecentFile } from '../types'
import { Button } from './ui'

interface WelcomeProps {
  recentFiles: RecentFile[]
  isDragging: boolean
  onOpen: () => void
  onOpenRecent: (file: RecentFile) => void
  onRemoveRecent: (path: string) => void
  onDrop: (event: DragEvent<HTMLDivElement>) => void
  onDragEnter: (event: DragEvent<HTMLDivElement>) => void
  onDragLeave: (event: DragEvent<HTMLDivElement>) => void
}

export function Welcome({ recentFiles, isDragging, onOpen, onOpenRecent, onRemoveRecent, onDrop, onDragEnter, onDragLeave }: WelcomeProps) {
  return (
    <main
      className="welcome"
      onDrop={onDrop}
      onDragOver={(event) => event.preventDefault()}
      onDragEnter={onDragEnter}
      onDragLeave={onDragLeave}
    >
      <section className="welcome-hero">
        <div className="welcome-kicker"><ShieldCheck size={14} /> Private by design · files stay on this device</div>
        <h1>Read, arrange, and refine PDFs.</h1>
        <p>A focused PDF workspace with the tools you use every day—and none of the clutter.</p>
        <div className={`drop-card ${isDragging ? 'is-dragging' : ''}`}>
          <div className="drop-icon"><FolderOpen size={24} strokeWidth={1.6} /></div>
          <strong>{isDragging ? 'Drop to open this file' : 'Open a document'}</strong>
          <span>Drop a file here, or choose one from your device.</span>
          <Button icon={FolderOpen} variant="primary" onClick={onOpen}>Choose file</Button>
          <div className="file-types">
            <span><FileType2 size={14} /> PDF</span>
            <span><FileImage size={14} /> JPG / PNG</span>
            <span><FileText size={14} /> Word / TXT</span>
          </div>
        </div>
      </section>

      <aside className="recent-panel">
        <div className="section-heading">
          <div>
            <span className="eyebrow">Your workspace</span>
            <h2>Recent files</h2>
          </div>
        </div>
        {recentFiles.length ? (
          <div className="recent-list">
            {recentFiles.map((file) => (
              <button key={file.path} type="button" className="recent-row" onClick={() => onOpenRecent(file)}>
                <span className="recent-file-icon"><FileType2 size={17} /></span>
                <span className="recent-copy">
                  <strong>{file.name}</strong>
                  <small>{file.path}</small>
                </span>
                <span
                  role="button"
                  tabIndex={0}
                  className="recent-remove"
                  aria-label={`Remove ${file.name} from recent files`}
                  onClick={(event) => { event.stopPropagation(); onRemoveRecent(file.path) }}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter' || event.key === ' ') {
                      event.stopPropagation()
                      onRemoveRecent(file.path)
                    }
                  }}
                ><X size={14} /></span>
              </button>
            ))}
          </div>
        ) : (
          <div className="recent-empty">
            <FileType2 size={20} />
            <strong>No recent files yet</strong>
            <span>Documents you open will appear here.</span>
          </div>
        )}
        <div className="privacy-note">
          <ShieldCheck size={15} />
          <p><strong>Local-first.</strong> simple never uploads your documents.</p>
        </div>
      </aside>
    </main>
  )
}
