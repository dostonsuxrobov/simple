import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { AlertTriangle, ChevronDown, ChevronRight, LoaderCircle, ScanText, X } from 'lucide-react'
import { preflightOcrAssets } from '../lib/ocr/assets'
import { classifyPages, type PageScanState } from '../lib/ocr/pageClassifier'
import {
  estimateOcrSeconds,
  formatOcrDuration,
  hasRecognizedText,
  pageNeedsRecognition,
  planOcrPage,
  warmUpOcr,
  type OcrRunProgress,
  type OcrScope,
} from '../lib/ocr/ocrRunner'
import type { OcrLanguageAsset } from '../lib/ocr/types'
import { errorMessage } from '../lib/utils'
import { Button, IconButton } from './ui'

export interface OcrDialogRequest {
  scope: OcrScope
  /** Candidate pages; the runner reads only those that need it. */
  pages: number[]
  /** Pages the user picked themselves (This page, Selected pages). */
  explicit: boolean
  language: string
  replaceExisting: boolean
}

interface OcrDialogProps {
  pdf: PDFDocumentProxy
  documentName: string
  currentPage: number
  selectedPages: number[]
  initialScope?: OcrScope
  signatureDetected: boolean
  /** Whether a page's decoded images may be released after checking it (pages off screen). */
  canReleasePage: (pageIndex: number) => boolean
  /** Set while recognising: the dialog shows progress and Stop. */
  progress: OcrRunProgress | null
  stopping?: boolean
  onRecognize: (request: OcrDialogRequest) => void
  onStop: () => void
  onClose: () => void
}

function focusTrap(event: ReactKeyboardEvent<HTMLElement>) {
  if (event.key !== 'Tab') return
  const focusable = [...event.currentTarget.querySelectorAll<HTMLElement>('button, input, select')]
    .filter((element) => !element.hasAttribute('disabled'))
  if (!focusable.length) return
  const first = focusable[0]
  const last = focusable[focusable.length - 1]
  if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
  else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
}

/** Recognize text: scope, language and options, then progress with Stop (Esc). */
export function OcrDialog({
  pdf, documentName, currentPage, selectedPages, initialScope, signatureDetected, canReleasePage,
  progress, stopping, onRecognize, onStop, onClose,
}: OcrDialogProps) {
  const pageCount = pdf.numPages
  const hasSelection = selectedPages.length > 1
  const [scope, setScope] = useState<OcrScope>(() => initialScope === 'selected' && !hasSelection ? 'needed' : initialScope ?? 'needed')
  const [states, setStates] = useState<ReadonlyMap<number, PageScanState>>(() => new Map())
  const [checked, setChecked] = useState(0)
  const [checkError, setCheckError] = useState('')
  const [languages, setLanguages] = useState<OcrLanguageAsset[] | null>(null)
  const [assetError, setAssetError] = useState('')
  const [language, setLanguage] = useState('eng')
  const [replaceExisting, setReplaceExisting] = useState(false)
  const [moreOpen, setMoreOpen] = useState(false)
  const canReleaseRef = useRef(canReleasePage)
  canReleaseRef.current = canReleasePage
  const running = Boolean(progress)
  const checking = checked < pageCount && !checkError

  // The components must be present before anything starts (a broken build fails here, at once).
  useEffect(() => {
    let active = true
    preflightOcrAssets().then((manifest) => {
      if (!active) return
      setLanguages(manifest.languages)
      const preferred = manifest.languages.some((item) => item.code === 'eng') ? 'eng' : manifest.languages[0]?.code
      if (preferred && preferred !== 'eng') setLanguage(preferred)
      // Start the engine while the options are read.
      if (preferred && !running) void warmUpOcr(preferred)
    }).catch((error) => {
      if (active) setAssetError(errorMessage(error))
    })
    return () => { active = false }
  }, [])

  // Count the pages that need recognising: this page first, then the selection, then the rest.
  useEffect(() => {
    if (running) return
    const controller = new AbortController()
    const order = [...new Set([currentPage, ...selectedPages, ...Array.from({ length: pageCount }, (_, index) => index)])]
    let lastPublished = 0
    classifyPages(pdf, order, {
      signal: controller.signal,
      canRelease: (index) => canReleaseRef.current(index),
      onProgress: (done, total, known) => {
        const now = performance.now()
        if (done < total && now - lastPublished < 120) return
        lastPublished = now
        setStates(new Map(known))
        setChecked(done)
      },
    }).catch((error) => {
      if (!controller.signal.aborted) setCheckError(errorMessage(error))
    })
    return () => controller.abort()
    // Classification is cached per page, so re-running after a run is cheap.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pdf, running])

  const candidates = useMemo(() => (
    scope === 'current' ? [currentPage] : scope === 'selected' ? selectedPages : Array.from({ length: pageCount }, (_, index) => index)
  ), [scope, currentPage, selectedPages, pageCount])
  const explicit = scope !== 'needed'
  const neededCount = useMemo(() => {
    let count = 0
    for (const state of states.values()) if (pageNeedsRecognition(state, replaceExisting)) count += 1
    return count
  }, [states, replaceExisting])
  const candidateStates = candidates.map((index) => states.get(index)).filter((state): state is PageScanState => Boolean(state))
  const candidatesKnown = candidateStates.length === candidates.length
  const plans = candidates.flatMap((index) => {
    const state = states.get(index)
    return state ? [planOcrPage(index, state, { explicit, replaceExisting })] : []
  })
  const toRecognize = plans.filter((plan) => plan.action === 'recognize').length
  const showReplace = scope === 'needed' ? hasRecognizedText(states.values()) : hasRecognizedText(candidateStates)
  const nothingToDo = candidatesKnown && toRecognize === 0
  const canRecognize = !running && !assetError && languages !== null && !nothingToDo && !checkError

  let note = ''
  if (nothingToDo && plans.length) {
    const reasons = new Set(plans.map((plan) => plan.reason))
    const single = candidates.length === 1
    if (reasons.size === 1 && reasons.has('has-ocr-text')) {
      note = `${single ? 'This page’s' : 'The text of these pages'} was already recognized. To read it again, open More and choose Replace text recognized earlier.`
    } else if (reasons.size === 1 && reasons.has('blank')) {
      note = single ? 'This page is blank.' : 'These pages are blank.'
    } else {
      note = scope === 'needed' ? 'Every page already has text.' : single ? 'This page already has text.' : 'These pages already have text.'
    }
  }
  const estimate = candidatesKnown || scope !== 'needed'
    ? formatOcrDuration(estimateOcrSeconds(scope === 'needed' ? neededCount : toRecognize))
    : ''

  function submit() {
    if (!canRecognize) return
    onRecognize({ scope, pages: candidates, explicit, language, replaceExisting: showReplace && replaceExisting })
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      if (running) onStop()
      else onClose()
      return
    }
    if (event.key === 'Enter' && !running && !(event.target instanceof HTMLButtonElement) && !(event.target instanceof HTMLSelectElement)) {
      event.preventDefault()
      event.stopPropagation()
      submit()
      return
    }
    focusTrap(event)
  }

  const selectedLanguage = languages?.find((item) => item.code === language)
  const status = progress
    ? progress.phase === 'checking'
      ? 'Checking pages…'
      : stopping
        ? 'Stopping…'
        : progress.starting
          ? 'Starting text recognition…'
          : `Recognizing page ${Math.min(progress.total, progress.done + 1)} of ${progress.total}…`
    : ''
  const remaining = progress?.phase === 'recognizing' && !stopping && progress.remainingSeconds !== undefined
    ? formatOcrDuration(progress.remainingSeconds)
    : ''

  return (
    <div className="export-dialog-overlay" onPointerDown={(event) => { if (event.target === event.currentTarget && !running) onClose() }}>
      <div className="export-dialog ocr-dialog" role="dialog" aria-modal="true" aria-label="Recognize text" aria-busy={running} onKeyDown={handleKeyDown}>
        <header className="export-dialog-header">
          <ScanText size={16} strokeWidth={1.8} aria-hidden="true" />
          <strong>Recognize text</strong>
          <span className="export-dialog-document">{documentName}</span>
          {!running && <IconButton icon={X} label="Close" compact onClick={onClose} />}
        </header>

        {running ? (
          <div className="export-dialog-body ocr-progress" role="status" aria-live="polite">
            <strong className="ocr-progress-status">{status}</strong>
            <div
              className="ocr-progress-bar"
              role="progressbar"
              aria-label="Text recognition progress"
              aria-valuemin={0}
              aria-valuemax={100}
              aria-valuenow={Math.round((progress?.fraction ?? 0) * 100)}
            >
              <span style={{ width: `${Math.round((progress?.fraction ?? 0) * 1000) / 10}%` }} />
            </div>
            <small>{remaining ? `${remaining} left` : progress?.phase === 'checking' ? `${progress.done} of ${progress.total} pages checked` : '\u00a0'}</small>
          </div>
        ) : (
          <div className="export-dialog-body">
            <p className="ocr-dialog-intro">Turns scanned pages into searchable, selectable text. The page images stay exactly as they are.</p>
            <fieldset className="export-field ocr-scope">
              <legend>Pages</legend>
              <label>
                <input type="radio" name="ocr-scope" checked={scope === 'current'} onChange={() => setScope('current')} />
                This page ({currentPage + 1})
              </label>
              <label>
                <input type="radio" name="ocr-scope" checked={scope === 'needed'} onChange={() => setScope('needed')} />
                <span>All pages that need it ({neededCount} of {pageCount})</span>
                {checking && <LoaderCircle className="spin ocr-counting" size={12} aria-label="Counting pages" />}
              </label>
              {hasSelection && (
                <label>
                  <input type="radio" name="ocr-scope" checked={scope === 'selected'} onChange={() => setScope('selected')} />
                  Selected pages ({selectedPages.length})
                </label>
              )}
            </fieldset>
            <div className="ocr-language">
              <span>Language</span>
              {languages && languages.length > 1 ? (
                <select aria-label="Recognition language" value={language} onChange={(event) => setLanguage(event.target.value)}>
                  {languages.map((item) => <option key={item.code} value={item.code}>{item.label}</option>)}
                </select>
              ) : (
                <strong>{selectedLanguage?.label ?? 'English'}</strong>
              )}
            </div>
            {showReplace && (
              <div className="ocr-more">
                <button type="button" className="ocr-more-toggle" aria-expanded={moreOpen} onClick={() => setMoreOpen((open) => !open)}>
                  {moreOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                  More
                </button>
                {moreOpen && (
                  <label className="ocr-replace">
                    <input type="checkbox" checked={replaceExisting} onChange={(event) => setReplaceExisting(event.target.checked)} />
                    Replace text recognized earlier
                  </label>
                )}
              </div>
            )}
            {note && <p className="ocr-note">{note}</p>}
            {signatureDetected && (
              <p className="export-layout-note ocr-warning"><AlertTriangle size={13} aria-hidden="true" />Recognizing text changes the document and invalidates its digital signature.</p>
            )}
            {(assetError || checkError) && <p className="inline-error ocr-error" role="alert">{assetError || checkError}</p>}
          </div>
        )}

        <footer className="export-dialog-footer">
          {running ? (
            <>
              <span />
              {/* Stop sits where Recognize was: the second click of a double-click must not stop the run. */}
              <Button variant="secondary" autoFocus disabled={stopping} onClick={(event) => { if (event.detail <= 1) onStop() }}>Stop</Button>
            </>
          ) : (
            <>
              <span>{checking && scope === 'needed' ? 'Counting pages…' : estimate}</span>
              <Button variant="secondary" onClick={onClose}>Cancel</Button>
              <Button icon={ScanText} variant="primary" autoFocus disabled={!canRecognize} onClick={submit}>Recognize</Button>
            </>
          )}
        </footer>
      </div>
    </div>
  )
}

const AUTO_OCR_KEY = 'simple:ocr:auto-on-edit'

/** "Do this automatically" from the Edit-mode offer. */
export function readAutoOcr() {
  try {
    return localStorage.getItem(AUTO_OCR_KEY) === '1'
  } catch {
    return false
  }
}

export function writeAutoOcr(enabled: boolean) {
  try {
    if (enabled) localStorage.setItem(AUTO_OCR_KEY, '1')
    else localStorage.removeItem(AUTO_OCR_KEY)
  } catch {
    // Storage unavailable: the choice applies to this offer only.
  }
}

interface OcrOfferProps {
  /** Where the page was clicked (window coordinates). */
  x: number
  y: number
  onRecognize: (automatically: boolean) => void
  /** Not now: this page is not offered again. */
  onDecline: () => void
  /** Escape, a click elsewhere, scrolling: the offer may come back. */
  onDismiss: () => void
}

/** Shown when a click in Edit mode lands on a scanned page that has no text. */
export function OcrOffer({ x, y, onRecognize, onDecline, onDismiss }: OcrOfferProps) {
  const ref = useRef<HTMLDivElement>(null)
  const [automatically, setAutomatically] = useState(false)
  const dismissRef = useRef(onDismiss)
  dismissRef.current = onDismiss

  useLayoutEffect(() => {
    const element = ref.current
    if (!element) return
    const bounds = element.getBoundingClientRect()
    element.style.left = `${Math.max(8, Math.min(x + 6, window.innerWidth - bounds.width - 8))}px`
    element.style.top = `${Math.max(46, Math.min(y + 10, window.innerHeight - bounds.height - 8))}px`
    element.querySelector<HTMLButtonElement>('.button-primary')?.focus({ preventScroll: true })
  }, [x, y])

  useEffect(() => {
    const dismiss = () => dismissRef.current()
    const outside = (event: PointerEvent) => { if (!ref.current?.contains(event.target as Node)) dismiss() }
    const viewer = document.querySelector('.viewer')
    document.addEventListener('pointerdown', outside, true)
    window.addEventListener('blur', dismiss)
    window.addEventListener('resize', dismiss)
    viewer?.addEventListener('scroll', dismiss, { passive: true })
    return () => {
      document.removeEventListener('pointerdown', outside, true)
      window.removeEventListener('blur', dismiss)
      window.removeEventListener('resize', dismiss)
      viewer?.removeEventListener('scroll', dismiss)
    }
  }, [])

  return (
    <div
      ref={ref}
      className="ocr-offer"
      role="dialog"
      aria-label="Recognize text on this page"
      style={{ left: x, top: y }}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.preventDefault()
          event.stopPropagation()
          onDismiss()
          return
        }
        focusTrap(event)
      }}
    >
      <p><strong>This page is a scanned image.</strong> Recognize its text to edit words directly.</p>
      <div className="ocr-offer-actions">
        <Button icon={ScanText} variant="primary" onClick={() => onRecognize(automatically)}>Recognize text</Button>
        <Button variant="ghost" onClick={onDecline}>Not now</Button>
      </div>
      <label className="ocr-offer-auto">
        <input type="checkbox" checked={automatically} onChange={(event) => setAutomatically(event.target.checked)} />
        Do this automatically
      </label>
    </div>
  )
}

interface OcrExportPromptProps {
  pageCount: number
  onRecognize: () => void
  onExportAnyway: () => void
  onCancel: () => void
}

/** Before a text export of pages that are scans without text. */
export function OcrExportPrompt({ pageCount, onRecognize, onExportAnyway, onCancel }: OcrExportPromptProps) {
  const single = pageCount === 1
  return (
    <div className="export-dialog-overlay" onPointerDown={(event) => { if (event.target === event.currentTarget) onCancel() }}>
      <div
        className="export-dialog ocr-dialog"
        role="alertdialog"
        aria-modal="true"
        aria-label="Scanned pages in this export"
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault()
            event.stopPropagation()
            onCancel()
            return
          }
          focusTrap(event)
        }}
      >
        <header className="export-dialog-header">
          <ScanText size={16} strokeWidth={1.8} aria-hidden="true" />
          <strong>Scanned pages</strong>
        </header>
        <div className="export-dialog-body">
          <p className="ocr-dialog-intro">
            {single
              ? '1 page in this export is a scanned image. Recognize its text first so it isn’t empty?'
              : `${pageCount} pages in this export are scanned images. Recognize their text first so they aren’t empty?`}
          </p>
        </div>
        <footer className="export-dialog-footer">
          <span />
          <Button variant="ghost" onClick={onCancel}>Cancel</Button>
          <Button variant="secondary" onClick={onExportAnyway}>Export anyway</Button>
          <Button icon={ScanText} variant="primary" autoFocus onClick={onRecognize}>Recognize and export</Button>
        </footer>
      </div>
    </div>
  )
}
