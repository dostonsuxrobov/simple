// src/advanced/useAdvancedMode.tsx (WP1)
// Host integration of the Advanced editor; main.tsx talks to Advanced mode only through this hook.
//   - Simple <-> Advanced state machine (design 7.1): enter from the Simple canvas or from a PSD,
//     exit flattens silently when nothing would be lost, otherwise asks (Cancel / Save as PSD / Flatten).
//   - Output routing (7.4): Save/Export/Print/Copy obtain pixels through withOutputCanvas; layered
//     documents save as PSD; the PSD overwrite policy (6.2) and the PSD fidelity prompt on open.
//   - Keyboard delegation (7.5): main.tsx keeps the only window keydown listener and hands every
//     non-host key to handleHostKey while Advanced is open or one of these dialogs is showing.
// The editor itself (./AdvancedEditor.tsx, WP6) and the PSD codec (./psd.ts, WP7) load lazily, so
// Simple start-up pays for neither.
import { Component, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { ComponentType, ErrorInfo, ReactNode } from 'react'
import { FileWarning, Layers, TriangleAlert } from 'lucide-react'
import { LIMITS } from './types.ts'
import type {
  AdvancedEditorHandle,
  AdvancedEditorProps,
  AdvancedHost,
  AdvancedInitialContent,
  ImportedDocument,
  ImportPsd,
  PsdFidelityIssue,
  PsdIssueCode,
} from './types.ts'
import type { PixelBuffer } from '../imaging/types.ts'
import type { AdvancedModeApi, AdvancedModeDeps, EditorMode, ExportFormat, OpenImage, SaveFormat } from '../shared/session.ts'
import { bufferToCanvas, releaseCanvas } from '../shared/canvas.ts'
import { trimImagingWorkers } from '../shared/workerClient.ts'
import './integration.css'

// ---------------------------------------------------------------------------------------------
// Optional extensions of the shared contracts (src/advanced/types.ts stays frozen)
// ---------------------------------------------------------------------------------------------

/** Optional members an AdvancedEditorHandle may implement; the host feature-checks each one. */
export interface AdvancedEditorHandleExtras {
  /** Keydown delegated by the host while Advanced is open (host file shortcuts excluded). Return true when consumed. */
  handleKeyDown?(event: KeyboardEvent): boolean
  /** Keyup while Advanced is open (spring-loaded tools). */
  handleKeyUp?(event: KeyboardEvent): boolean
  /** False when this editor cannot encode a layered PSD yet; hides the PSD save/export entries. Default true. */
  canEncodePsd?(): boolean
}

/** Extra host members passed to the editor on top of AdvancedHost. */
export interface AdvancedHostExtras {
  /**
   * The host content revision right now. A document store must start at this revision so that undoing
   * back to the opening state restores the host's Modified state exactly (it may be non-zero after Simple edits).
   */
  currentRevision(): number
}

export type AdvancedEditorComponent = ComponentType<AdvancedEditorProps>

type EditorHandle = AdvancedEditorHandle & AdvancedEditorHandleExtras

interface ExportOption {
  readonly format: ExportFormat
  readonly label: string
  readonly detail: string
}

export interface PsdImportResult {
  readonly document: ImportedDocument
  /** Issues of the layered import (kept even when the user opened a flattened copy). */
  readonly issues: readonly PsdFidelityIssue[]
}

/** What main.tsx gets: the shared AdvancedModeApi plus the integration helpers it needs. */
export interface AdvancedModeController extends AdvancedModeApi {
  /** Entering or leaving is in progress (disable the Advanced button). */
  readonly busy: boolean
  /** A dialog owned by this hook is open; main.tsx routes keys to handleHostKey while it is. */
  readonly modalOpen: boolean
  /** Layered PSD save/export can be offered right now. */
  readonly psdAvailable: boolean
  /** Extra Export As entries (PSD in Advanced). Empty in Simple mode. */
  readonly extraExportOptions: readonly ExportOption[]
  /** Commits open editor sessions (text, transform) so Save, Export and Close see every edit. */
  settle(): Promise<void>
  /** Format for "Save as..." in Advanced: the flattened image format (PNG for PSD sources). */
  flatSaveFormat(): SaveFormat
  /** True when saving in `format` drops layers (layered document in Advanced, non-PSD format). */
  wouldLoseLayers(format: SaveFormat): boolean
  /** PSD overwrite policy: asks when overwriting a PSD whose unsupported features would be lost. */
  confirmPsdOverwrite(image: OpenImage): Promise<'copy' | 'overwrite' | 'cancel'>
  /** Parses a PSD payload (lazy ./psd.ts) and asks about unsupported features. Null when cancelled. */
  importPsdPayload(payload: ImagePayload): Promise<PsdImportResult | null>
  /** Leaves Advanced without flattening because another document replaced this one. */
  discard(): void
  /** Pixel size the output paths produce right now. */
  outputSize(): { readonly width: number; readonly height: number } | null
}

// ---------------------------------------------------------------------------------------------
// Lazy modules
// ---------------------------------------------------------------------------------------------

interface PsdModule {
  readonly importPsd: ImportPsd
  readonly configurePsdCanvas?: (mode: 'dom' | 'node-test') => void
}

// The glob resolves at build time: empty until WP7 adds ./psd.ts, so the build never depends on it.
const PSD_MODULES = import.meta.glob<PsdModule>('./psd.ts')
const loadPsdModule: (() => Promise<PsdModule>) | undefined = PSD_MODULES['./psd.ts']
let psdCanvasConfigured = false

let editorPromise: Promise<AdvancedEditorComponent> | null = null
let loadedEditor: AdvancedEditorComponent | null = null

function loadEditor(): Promise<AdvancedEditorComponent> {
  if (!editorPromise) {
    editorPromise = import('./AdvancedEditor.tsx').then((module) => {
      loadedEditor = module.AdvancedEditor
      return module.AdvancedEditor
    })
    editorPromise.catch(() => { editorPromise = null })
  }
  return editorPromise
}

// ---------------------------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------------------------

const PSD_EXPORT_OPTION: ExportOption = Object.freeze({
  format: 'psd',
  label: 'Photoshop document (PSD)',
  detail: 'Keeps layers, masks and blend modes editable',
})

const ISSUE_LABELS: Readonly<Record<PsdIssueCode, string>> = Object.freeze({
  'layer-effects': 'Layer styles',
  'smart-object': 'Smart objects',
  'vector-mask': 'Vector masks',
  'unsupported-adjustment': 'Some adjustment layers',
  'group-flattened': 'Layer groups',
  'unsupported-blend-mode': 'Some blend modes',
  'fill-opacity': 'Fill opacity',
  'text-rerender': 'Text rendering',
  'bit-depth-reduced': '16/32-bit color',
  'color-profile': 'Color profile',
  'knockout-or-advanced-blending': 'Advanced blending',
})

/** "Layer styles (3), Smart objects (1)" for the issues not in `ignore`; empty when none remain. */
export function summarizePsdIssues(issues: readonly PsdFidelityIssue[], ignore: readonly PsdIssueCode[]): string {
  const counts = new Map<PsdIssueCode, number>()
  for (const issue of issues) {
    if (!ignore.includes(issue.code)) counts.set(issue.code, (counts.get(issue.code) ?? 0) + 1)
  }
  return [...counts].map(([code, count]) => `${ISSUE_LABELS[code] ?? code} (${count})`).join(', ')
}

export function pixelsHaveAlpha(buffer: PixelBuffer | null): boolean {
  if (!buffer) return false
  const data = buffer.data
  for (let index = 3; index < data.length; index += 4) if (data[index] < 255) return true
  return false
}

function messageOf(error: unknown, fallback: string): string {
  return error instanceof Error && error.message ? error.message : fallback
}

type HostChord = 'open' | 'save' | 'print' | 'export'

/** Shortcuts the host keeps in Advanced mode (Photoshop: Ctrl+Shift+E is Merge Visible there). */
function hostChord(event: KeyboardEvent): HostChord | null {
  if (!(event.ctrlKey || event.metaKey)) return null
  const key = event.key.toLowerCase()
  if (event.altKey) return event.shiftKey && (key === 'w' || event.code === 'KeyW') ? 'export' : null
  if (key === 's') return 'save'
  if (event.shiftKey) return null
  if (key === 'o') return 'open'
  if (key === 'p') return 'print'
  return null
}

// ---------------------------------------------------------------------------------------------
// Dialog state
// ---------------------------------------------------------------------------------------------

type DialogState =
  | { readonly kind: 'flatten'; readonly layers: number; readonly canSavePsd: boolean; readonly working: boolean }
  | { readonly kind: 'psd-overwrite'; readonly name: string; readonly summary: string; readonly resolve: (choice: 'copy' | 'overwrite' | 'cancel') => void }
  | { readonly kind: 'psd-fidelity'; readonly name: string; readonly summary: string; readonly resolve: (choice: 'layers' | 'flattened' | 'cancel') => void }

interface Session {
  readonly key: number
  readonly initial: AdvancedInitialContent
  /** Host revision when Advanced opened; the Simple canvas holds exactly this revision unless stale. */
  readonly entryRevision: number
  /** The Simple canvas does not hold this document (it was opened from a PSD). */
  readonly stale: boolean
}

// ---------------------------------------------------------------------------------------------
// Error boundary: an editor crash must never take the Simple shell (and its unsaved image) down
// ---------------------------------------------------------------------------------------------

interface BoundaryProps {
  readonly children: ReactNode
  readonly onError: (error: Error) => void
}

class EditorBoundary extends Component<BoundaryProps, { failed: boolean }> {
  state = { failed: false }

  static getDerivedStateFromError(): { failed: boolean } {
    return { failed: true }
  }

  componentDidCatch(error: Error, _info: ErrorInfo): void {
    this.props.onError(error)
  }

  render(): ReactNode {
    return this.state.failed ? null : this.props.children
  }
}

// ---------------------------------------------------------------------------------------------
// The hook
// ---------------------------------------------------------------------------------------------

export function useAdvancedMode(deps: AdvancedModeDeps): AdvancedModeController {
  const depsRef = useRef(deps)
  depsRef.current = deps

  const [mode, setModeState] = useState<EditorMode>('simple')
  const [session, setSessionState] = useState<Session | null>(null)
  const [Editor, setEditor] = useState<AdvancedEditorComponent | null>(() => loadedEditor)
  const [handle, setHandleState] = useState<EditorHandle | null>(null)
  const [dialog, setDialogState] = useState<DialogState | null>(null)
  const [busy, setBusyState] = useState(false)
  const [failure, setFailure] = useState<string | null>(null)

  const modeRef = useRef<EditorMode>('simple')
  const sessionRef = useRef<Session | null>(null)
  const handleRef = useRef<EditorHandle | null>(null)
  const dialogRef = useRef<DialogState | null>(null)
  const busyRef = useRef(false)
  const sessionKeyRef = useRef(0)
  const returnFocusRef = useRef(false)

  const setMode = (next: EditorMode) => { modeRef.current = next; setModeState(next) }
  const setSession = (next: Session | null) => { sessionRef.current = next; setSessionState(next) }
  const setHandle = (next: EditorHandle | null) => { handleRef.current = next; setHandleState(next) }
  const setDialog = (next: DialogState | null) => { dialogRef.current = next; setDialogState(next) }
  const setBusy = (next: boolean) => { busyRef.current = next; setBusyState(next) }

  const psdAvailable = useCallback((): boolean => {
    if (!loadPsdModule) return false
    const current = handleRef.current
    return !current || current.canEncodePsd?.() !== false
  }, [])

  const notify = (message: string, tone: 'normal' | 'error' = 'normal') => depsRef.current.notify(message, tone)

  // ----- session lifecycle -------------------------------------------------------------------

  const startSession = useCallback((initial: AdvancedInitialContent, stale: boolean) => {
    sessionKeyRef.current += 1
    setSession({ key: sessionKeyRef.current, initial, entryRevision: depsRef.current.revisions.current(), stale })
    setHandle(null)
    setFailure(null)
    if (loadedEditor) setEditor(() => loadedEditor)
    setMode('advanced')
  }, [])

  const teardown = useCallback((returnFocus: boolean) => {
    const pending = dialogRef.current
    if (pending && pending.kind !== 'flatten') pending.resolve('cancel')
    setDialog(null)
    setSession(null)
    setHandle(null)
    setFailure(null)
    setMode('simple')
    trimImagingWorkers()
    returnFocusRef.current = returnFocus
  }, [])

  // After a user-initiated exit, focus the Advanced button once the Simple toolbar is visible again.
  useEffect(() => {
    if (mode !== 'simple' || busy || !returnFocusRef.current) return
    returnFocusRef.current = false
    document.querySelector<HTMLElement>('.advanced-button:not(:disabled)')?.focus()
  }, [busy, mode])

  // The editor module can still be loading when a PSD opens; mount it as soon as it arrives.
  useEffect(() => {
    if (mode !== 'advanced' || Editor) return
    let cancelled = false
    loadEditor()
      .then((component) => { if (!cancelled) setEditor(() => component) })
      .catch((error: unknown) => { if (!cancelled) setFailure(messageOf(error, 'The Advanced editor could not be loaded.')) })
    return () => { cancelled = true }
  }, [Editor, mode])

  const enter = useCallback(async () => {
    if (modeRef.current !== 'simple' || busyRef.current || dialogRef.current) return
    const start = depsRef.current
    const image = start.getImage()
    const canvas = start.canvasRef.current
    if (!image || !canvas || !canvas.width || !canvas.height || start.isSuspended()) return
    setBusy(true)
    try {
      await start.settleSimple()
      const component = await loadEditor()
      const now = depsRef.current
      const current = now.getImage()
      const liveCanvas = now.canvasRef.current
      // Another image may have been opened while the editor loaded; only enter for the same document.
      if (modeRef.current !== 'simple' || !current || !liveCanvas || current.name !== image.name || current.path !== image.path) return
      setEditor(() => component)
      startSession({ kind: 'canvas', canvas: liveCanvas, name: current.name, hasAlpha: current.hasAlpha, ppi: 72 }, false)
    } catch (error) {
      notify(messageOf(error, 'The Advanced editor could not be opened.'), 'error')
    } finally {
      setBusy(false)
    }
  }, [startSession])

  const enterWithDocument = useCallback((document: ImportedDocument, name: string) => {
    startSession({ kind: 'document', document, name }, true)
  }, [startSession])

  /** Writes the flattened document into the Simple canvas (when needed) and returns to Simple. */
  const completeExit = useCallback(async (current: EditorHandle) => {
    const now = depsRef.current
    const active = sessionRef.current
    if (!active) return
    const changed = now.revisions.current() !== active.entryRevision
    if (active.stale || changed) {
      const hasAlpha = await current.hasTransparency()
      const flattened = await current.renderFlattened()
      try {
        // One Simple undo step that restores the pre-Advanced pixels and revision; none for a PSD.
        now.adoptFlattened(flattened, active.entryRevision, !active.stale && changed)
      } finally {
        releaseCanvas(flattened)
      }
      now.updateImage({ hasAlpha })
    }
    const image = now.getImage()
    const leavingPsd = image?.saveFormat === 'psd'
    // A later Simple Save must never overwrite the layered PSD with a flat image.
    if (leavingPsd) now.updateImage({ saveFormat: 'png', path: null })
    teardown(true)
    if (leavingPsd && image) notify(`Back in Simple. ${image.name} keeps its layers; Save now creates a PNG.`)
  }, [teardown])

  /** Exit when the editor never became ready or crashed: fall back to what the host still holds. */
  const exitWithoutEditor = useCallback(async () => {
    const now = depsRef.current
    const active = sessionRef.current
    if (!active) { teardown(true); return }
    const edited = now.revisions.current() !== active.entryRevision
    if (!active.stale) {
      // The Simple canvas still holds the image exactly as Advanced received it.
      if (edited) now.revisions.set(active.entryRevision)
      teardown(true)
      if (edited) notify('The Advanced editor stopped. Your image is back as it was before Advanced; edits made there were lost.', 'error')
      return
    }
    const composite = active.initial.kind === 'document' ? active.initial.document.composite : null
    if (!composite) {
      notify('This document cannot be shown in Simple mode. Open another image, or reopen this one.', 'error')
      return
    }
    const canvas = bufferToCanvas(composite)
    try {
      now.adoptFlattened(canvas, active.entryRevision, false)
    } finally {
      releaseCanvas(canvas)
    }
    if (edited) now.revisions.set(active.entryRevision)
    now.updateImage({ hasAlpha: pixelsHaveAlpha(composite), saveFormat: 'png', path: null })
    teardown(true)
    notify(edited
      ? 'The Advanced editor stopped. The document is shown as it was opened; edits made there were lost.'
      : 'Back in Simple with the document as it was opened. Save creates a PNG.', edited ? 'error' : 'normal')
  }, [teardown])

  const requestExit = useCallback(async () => {
    if (modeRef.current !== 'advanced' || busyRef.current || dialogRef.current) return
    if (depsRef.current.isSuspended()) return
    const current = handleRef.current
    setBusy(true)
    try {
      if (!current) {
        await exitWithoutEditor()
        return
      }
      await current.settle()
      if (current.isFlatEquivalent()) await completeExit(current)
      else setDialog({ kind: 'flatten', layers: current.layerCount(), canSavePsd: psdAvailable(), working: false })
    } catch (error) {
      notify(messageOf(error, 'Could not return to Simple mode.'), 'error')
    } finally {
      setBusy(false)
    }
  }, [completeExit, exitWithoutEditor, psdAvailable])

  const requestExitRef = useRef(requestExit)
  requestExitRef.current = requestExit

  const discard = useCallback(() => {
    if (modeRef.current === 'advanced' || sessionRef.current) teardown(false)
  }, [teardown])

  // ----- dialogs ---------------------------------------------------------------------------

  const cancelDialog = useCallback((current: DialogState) => {
    if (current.kind === 'flatten') {
      if (!current.working) setDialog(null)
      return
    }
    current.resolve('cancel')
  }, [])

  const ask = useCallback(<T extends string>(make: (resolve: (choice: T) => void) => DialogState): Promise<T> => {
    return new Promise<T>((resolve) => {
      const previous = dialogRef.current
      if (previous) cancelDialog(previous)
      let state: DialogState | null = null
      const finish = (choice: T) => {
        if (dialogRef.current === state) setDialog(null)
        resolve(choice)
      }
      state = make(finish)
      setDialog(state)
    })
  }, [cancelDialog])

  const flattenDialogAction = useCallback(async (action: 'save-psd' | 'flatten') => {
    const current = dialogRef.current
    const editor = handleRef.current
    if (!current || current.kind !== 'flatten' || current.working) return
    if (!editor) { setDialog(null); return }
    setDialog({ ...current, working: true })
    try {
      if (action === 'save-psd') {
        const saved = await depsRef.current.save(true)
        if (!saved) {
          setDialog({ ...current, working: false })
          return
        }
      }
      await completeExit(editor)
    } catch (error) {
      notify(messageOf(error, 'Could not return to Simple mode.'), 'error')
      if (dialogRef.current) setDialog({ ...current, working: false })
    }
  }, [completeExit])

  useEffect(() => () => {
    const pending = dialogRef.current
    if (pending && pending.kind !== 'flatten') pending.resolve('cancel')
  }, [])

  // ----- host adapter given to the editor ----------------------------------------------------

  const host = useMemo<AdvancedHost & AdvancedHostExtras>(() => ({
    nextRevision: () => depsRef.current.revisions.next(),
    setRevision: (revision) => depsRef.current.revisions.set(revision),
    currentRevision: () => depsRef.current.revisions.current(),
    notify: (message, tone) => depsRef.current.notify(message, tone),
    isSuspended: () => depsRef.current.isSuspended() || Boolean(dialogRef.current),
    requestExit: () => { void requestExitRef.current() },
    save: (forceDialog) => depsRef.current.save(forceDialog),
    openExportMenu: () => depsRef.current.openExportMenu(),
    print: () => depsRef.current.print(),
    copyPng: async (png) => { await window.simpleImage.copyPng(png) },
    readClipboardImage: async () => (typeof window.simpleImage.readClipboardImage === 'function'
      ? window.simpleImage.readClipboardImage()
      : null),
  }), [])

  const sessionKey = session?.key ?? 0
  const onReady = useCallback((ready: AdvancedEditorHandle) => {
    if (sessionRef.current?.key !== sessionKey) return
    setHandle(ready as EditorHandle)
    ready.focus()
  }, [sessionKey])

  const onEditorError = useCallback((error: Error) => {
    setHandle(null)
    setFailure(messageOf(error, 'The Advanced editor stopped unexpectedly.'))
  }, [])

  // Spring-loaded tools need key releases; the host owns no keyup listener, so forward them.
  useEffect(() => {
    if (mode !== 'advanced') return
    const onKeyUp = (event: KeyboardEvent) => {
      if (dialogRef.current || depsRef.current.isSuspended()) return
      handleRef.current?.handleKeyUp?.(event)
    }
    window.addEventListener('keyup', onKeyUp)
    return () => window.removeEventListener('keyup', onKeyUp)
  }, [mode])

  // ----- output routing ----------------------------------------------------------------------

  const settle = useCallback(async () => {
    if (modeRef.current !== 'advanced') return
    await handleRef.current?.settle()
  }, [])

  const withOutputCanvas = useCallback(async <T,>(use: (canvas: HTMLCanvasElement) => Promise<T>): Promise<T> => {
    const now = depsRef.current
    if (modeRef.current !== 'advanced') {
      const canvas = now.canvasRef.current
      if (!canvas || !canvas.width || !canvas.height) throw new Error('There is no image to use.')
      return use(canvas)
    }
    const current = handleRef.current
    if (current) {
      await current.settle()
      const flattened = await current.renderFlattened()
      try {
        return await use(flattened)
      } finally {
        releaseCanvas(flattened)
      }
    }
    // Editor not ready (still reading pixels, or failed): nothing can have been edited yet, so the
    // host's copy is exact.
    const active = sessionRef.current
    if (active && now.revisions.current() === active.entryRevision) {
      if (!active.stale) {
        const canvas = now.canvasRef.current
        if (canvas && canvas.width && canvas.height) return use(canvas)
      } else if (active.initial.kind === 'document' && active.initial.document.composite) {
        const canvas = bufferToCanvas(active.initial.document.composite)
        try {
          return await use(canvas)
        } finally {
          releaseCanvas(canvas)
        }
      }
    }
    throw new Error('The Advanced editor is still opening. Try again in a moment.')
  }, [])

  const effectiveSaveFormat = useCallback((): SaveFormat => {
    const format = depsRef.current.getImage()?.saveFormat ?? 'png'
    if (modeRef.current !== 'advanced' || format === 'psd') return format
    const current = handleRef.current
    return current && !current.isFlatEquivalent() && psdAvailable() ? 'psd' : format
  }, [psdAvailable])

  const flatSaveFormat = useCallback((): SaveFormat => {
    const format = depsRef.current.getImage()?.saveFormat ?? 'png'
    return format === 'psd' ? 'png' : format
  }, [])

  const wouldLoseLayers = useCallback((format: SaveFormat): boolean => {
    if (modeRef.current !== 'advanced' || format === 'psd') return false
    const current = handleRef.current
    return Boolean(current && !current.isFlatEquivalent())
  }, [])

  const encodePsd = useCallback(async (): Promise<Uint8Array> => {
    const current = modeRef.current === 'advanced' ? handleRef.current : null
    if (!current) throw new Error('Photoshop documents are saved from the Advanced editor.')
    await current.settle()
    return current.encodePsd()
  }, [])

  const outputSize = useCallback(() => {
    const now = depsRef.current
    if (modeRef.current === 'advanced') {
      const current = handleRef.current
      if (current) return current.size()
      const active = sessionRef.current
      if (active?.initial.kind === 'document') return { width: active.initial.document.width, height: active.initial.document.height }
    }
    const canvas = now.canvasRef.current
    return canvas ? { width: canvas.width, height: canvas.height } : null
  }, [])

  const confirmPsdOverwrite = useCallback((image: OpenImage): Promise<'copy' | 'overwrite' | 'cancel'> => {
    // Text layers stay editable (Photoshop re-renders them), so only real losses are listed.
    const summary = summarizePsdIssues(image.psdIssues, ['text-rerender'])
    if (!summary) return Promise.resolve('overwrite')
    return ask<'copy' | 'overwrite' | 'cancel'>((resolve) => ({ kind: 'psd-overwrite', name: image.name, summary, resolve }))
  }, [ask])

  const importPsdPayload = useCallback(async (payload: ImagePayload): Promise<PsdImportResult | null> => {
    if (!loadPsdModule) throw new Error('Photoshop documents (.psd) cannot be opened in this version.')
    void loadEditor().catch(() => {})
    const module = await loadPsdModule()
    if (!psdCanvasConfigured) {
      module.configurePsdCanvas?.('dom')
      psdCanvasConfigured = true
    }
    const limits = { maxPixels: LIMITS.maxPixels, maxDimension: LIMITS.maxDimension, memoryLimitBytes: LIMITS.psdMemoryLimitBytes }
    let document = await module.importPsd(payload.data, { ...limits, mode: 'layers' })
    const issues = document.issues
    const summary = summarizePsdIssues(issues, ['text-rerender', 'bit-depth-reduced'])
    if (summary) {
      const choice = await ask<'layers' | 'flattened' | 'cancel'>((resolve) => ({ kind: 'psd-fidelity', name: payload.name, summary, resolve }))
      if (choice === 'cancel') return null
      if (choice === 'flattened') document = await module.importPsd(payload.data, { ...limits, mode: 'flattened' })
    }
    await loadEditor()
    return { document, issues }
  }, [ask])

  // ----- keyboard ----------------------------------------------------------------------------

  const handleHostKey = useCallback((event: KeyboardEvent): boolean => {
    const current = dialogRef.current
    if (current) {
      if (event.key === 'Escape') {
        event.preventDefault()
        cancelDialog(current)
      } else if (event.key !== 'Tab' && event.key !== 'Enter' && event.key !== ' ') {
        event.preventDefault()
      }
      return true
    }
    if (modeRef.current !== 'advanced') return false
    const chord = hostChord(event)
    if (chord === 'export') {
      event.preventDefault()
      if (!event.repeat) depsRef.current.openExportMenu()
      return true
    }
    if (chord) return false
    if (busyRef.current || depsRef.current.isSuspended()) return true
    handleRef.current?.handleKeyDown?.(event)
    return true
  }, [cancelDialog])

  // ----- rendering ---------------------------------------------------------------------------

  const suspended = deps.isSuspended() || Boolean(dialog)

  const view: ReactNode = mode === 'advanced' && session ? (
    <section className="advanced-root" data-print-background aria-label="Advanced editor">
      {failure ? (
        <div className="advanced-status-panel" role="alert">
          <TriangleAlert />
          <strong>The Advanced editor stopped</strong>
          <p>{failure}</p>
          <button type="button" onClick={() => void requestExitRef.current()}>Back to Simple</button>
        </div>
      ) : Editor ? (
        <EditorBoundary key={session.key} onError={onEditorError}>
          <Editor key={session.key} host={host} initial={session.initial} suspended={suspended} onReady={onReady} />
        </EditorBoundary>
      ) : (
        <div className="advanced-status-panel" aria-live="polite"><Layers /><strong>Opening the Advanced editor…</strong></div>
      )}
    </section>
  ) : null

  let dialogNode: ReactNode = null
  if (dialog?.kind === 'flatten') {
    const layerText = dialog.layers === 1 ? 'Its layer settings' : `Its ${dialog.layers} layers`
    dialogNode = (
      <div className="modal-backdrop advanced-dialog-backdrop">
        <div className="modal advanced-dialog" role="dialog" aria-modal="true" aria-labelledby="advanced-flatten-title" aria-describedby="advanced-flatten-detail">
          <div className="modal-icon"><Layers /></div>
          <h2 id="advanced-flatten-title">Flatten to return to Simple?</h2>
          <p id="advanced-flatten-detail">
            Simple edits one flat image. {layerText} will be merged{dialog.canSavePsd ? '; save a Photoshop copy first to keep them editable.' : ' and cannot be split again.'}
          </p>
          <div className="modal-actions">
            <button type="button" disabled={dialog.working} onClick={() => setDialog(null)}>Cancel</button>
            {dialog.canSavePsd && (
              <button type="button" className="modal-primary" disabled={dialog.working} autoFocus onClick={() => void flattenDialogAction('save-psd')}>Save as PSD…</button>
            )}
            <button
              type="button"
              className={dialog.canSavePsd ? undefined : 'modal-primary'}
              disabled={dialog.working}
              autoFocus={!dialog.canSavePsd}
              onClick={() => void flattenDialogAction('flatten')}
            >Flatten</button>
          </div>
        </div>
      </div>
    )
  } else if (dialog?.kind === 'psd-overwrite') {
    const resolve = dialog.resolve
    dialogNode = (
      <div className="modal-backdrop advanced-dialog-backdrop">
        <div className="modal advanced-dialog" role="dialog" aria-modal="true" aria-labelledby="advanced-overwrite-title" aria-describedby="advanced-overwrite-detail">
          <div className="modal-icon"><FileWarning /></div>
          <h2 id="advanced-overwrite-title">Overwrite {dialog.name}?</h2>
          <p id="advanced-overwrite-detail">Simple can't keep: {dialog.summary}. Save a copy to leave the original untouched.</p>
          <div className="modal-actions">
            <button type="button" onClick={() => resolve('cancel')}>Cancel</button>
            <button type="button" onClick={() => resolve('overwrite')}>Overwrite</button>
            <button type="button" className="modal-primary" autoFocus onClick={() => resolve('copy')}>Save a copy…</button>
          </div>
        </div>
      </div>
    )
  } else if (dialog?.kind === 'psd-fidelity') {
    const resolve = dialog.resolve
    dialogNode = (
      <div className="modal-backdrop advanced-dialog-backdrop">
        <div className="modal advanced-dialog" role="dialog" aria-modal="true" aria-labelledby="advanced-fidelity-title" aria-describedby="advanced-fidelity-detail">
          <div className="modal-icon"><FileWarning /></div>
          <h2 id="advanced-fidelity-title">{dialog.name} uses features Simple can't edit</h2>
          <p id="advanced-fidelity-detail">{dialog.summary}. Open the layers to edit everything else, or open a flattened copy that looks exactly like the file.</p>
          <div className="modal-actions">
            <button type="button" onClick={() => resolve('cancel')}>Cancel</button>
            <button type="button" onClick={() => resolve('flattened')}>Open flattened</button>
            <button type="button" className="modal-primary" autoFocus onClick={() => resolve('layers')}>Open layers</button>
          </div>
        </div>
      </div>
    )
  }

  const psdNow = mode === 'advanced' && psdAvailable()

  return {
    mode,
    handle,
    busy,
    modalOpen: Boolean(dialog),
    psdAvailable: psdNow,
    extraExportOptions: psdNow ? [PSD_EXPORT_OPTION] : [],
    enter,
    enterWithDocument,
    requestExit: () => { void requestExit() },
    withOutputCanvas,
    effectiveSaveFormat,
    encodePsd,
    handleHostKey,
    settle,
    flatSaveFormat,
    wouldLoseLayers,
    confirmPsdOverwrite,
    importPsdPayload,
    discard,
    outputSize,
    view,
    dialog: dialogNode,
  }
}
