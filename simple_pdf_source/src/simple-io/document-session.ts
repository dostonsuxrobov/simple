// Vendored from simple/shared/renderer/document-session.ts by simple/scripts/sync-shared.cjs. Do not edit here.

// DocumentSession: the renderer half of the shared Save contract (design §3.1,
// §3.6) and of crash recovery (§4.2-§4.5). Create one session per open
// document; React workspaces wrap it in a hook (subscribe + state work with
// useSyncExternalStore), Docs uses it directly.
//
// Rules it enforces:
// - Pending edits are committed before every save, close, export, print and
//   open-replace. A blocked commit stops the action with the workspace's own
//   message; Ctrl+S is never ignored silently.
// - Dirty state is revision based. A save clears it only up to the revision it
//   captured, so an edit made while a save runs stays dirty.
// - One save runs at a time. Repeated Ctrl+S while a save runs queues exactly
//   one follow-up; a save waits for an export or print that is running.
// - Failures are never silent: the session shows the catalog's decision prompt
//   (through main) and follows the answer (Try Again, Save As, Replace, …).
// - Unsaved work is journaled for crash recovery: 3 s after the last change,
//   at least every 30 s while editing continues, immediately on blur, before
//   print or open-replace, and on demand.

import type {
  IoCode,
  IoFailure,
  IoResult,
  IoSuccess,
  LossInfo,
  ModuleId,
  NotifyAction,
  NotifyKind,
  RecoveryEntry,
  RecoveryPart,
  RecoveryPayload,
  RecoverySnapshot,
  RequestType,
  SimpleIO,
  StatusKey,
} from './io-client'
import {
  appName,
  baseName,
  capitalize,
  catalogText,
  failureFromError,
  folderLabel,
  folderOf,
  joinList,
  kindName,
  promptFailure,
  samePath,
  statusText,
  toIoResult,
} from './io-client'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** Save writes the bound file; Save As rebinds to a new file; Save a Copy writes a copy and changes nothing else. */
export type SaveMode = 'save' | 'save-as' | 'save-copy'

/** Non-save actions that must commit pending edits first. "replace" is an Open that replaces this document. */
export type SessionAction = 'export' | 'print' | 'replace'

/** Long-running actions a save waits for. */
export type ActivityKind = 'export' | 'print'

/** What the session asks the workspace to save. The adapter forwards it to its save IPC (main's performSave). */
export interface SaveRequest {
  mode: SaveMode
  /** Session revision captured after committing pending edits. Main echoes nothing back; the session tracks it. */
  revision: number
  /** True when the model equals the bound file, so main may write that file's bytes instead of re-serializing. */
  pristine: boolean
  /** Registry id of the target format; omitted for "keep the current format". */
  format?: string
  docId: string
  /** The user chose Replace in the changed-on-disk prompt. */
  force?: boolean
  /** The user chose Save Here Again after the file was moved or deleted. */
  recreate?: boolean
  /** Save As chosen from a failure prompt: main opens the fallback folder for this code (design §3.4, D4). */
  fallbackReason?: IoCode
  /** The path that failed, when known. */
  failedPath?: string
}

/** Result of commitPendingEdits(): nothing, or why the action can't run yet. */
export type CommitResult = void | undefined | { blocked: string; focus?: () => void }

/**
 * What a recovery snapshot contains. An adapter returns at least the parts;
 * the session fills in docId, revision, title and module.
 */
export type RecoveryContent = Pick<RecoverySnapshot, 'parts'> & Partial<Omit<RecoverySnapshot, 'parts'>>

/**
 * The workspace side of a session (design §3.6, §3.9, §4.3).
 */
export interface DocumentAdapter {
  /**
   * Applies every in-progress edit to the model (text box, cell editor,
   * formula bar, side-panel inputs, brush stroke, placed signature) and calls
   * session.noteChange() for each change it applies. Returns `{blocked}` with
   * a specific message when something can't be applied yet (an open dialog
   * with unapplied values, a crop drawn but not applied, a value that fails
   * validation); `focus` moves the user to it. Never discard pending edits.
   */
  commitPendingEdits(): Promise<CommitResult> | CommitResult
  /**
   * Optional: true when the model equals the last saved or opened file, for
   * example after undoing back to it (Docs: model === baseline; Image: an
   * explicit pristine flag). It must be cheap; it runs on every change.
   */
  isPristine?(): boolean
  /**
   * Saves through the workspace's save IPC and returns main's IoResult.
   * Expected failures are returned, not thrown. On success the adapter moves
   * its own pristine baseline to the model it serialized (not to edits made
   * while the save ran).
   */
  save(request: SaveRequest): Promise<IoResult>
  /**
   * Builds the recovery parts for the current model (design §4.3), or null
   * when there is nothing to keep yet (the next change tries again).
   */
  snapshotForRecovery(): Promise<RecoveryContent | null>
  /** Rebuilds the model from a recovery entry. The session marks the document dirty afterwards. */
  restoreFromRecovery(entry: RecoveryEntry, payload: RecoveryPayload): Promise<void>
  /** The document's display name, e.g. "Budget.xlsx". */
  title(): string
  /** Shows a toast. Only success and information; failures get decision prompts. */
  notify(kind: NotifyKind, text: string, actions?: NotifyAction[]): void
}

/** Timer functions, injectable for tests. Handles are opaque. */
export interface SessionTimers {
  setTimeout(callback: () => void, ms: number): unknown
  clearTimeout(handle: unknown): void
  now(): number
}

/** Autosave (recovery journal) cadence; defaults follow design §4.2. */
export interface RecoveryCadence {
  /** Write recovery copies at all (default true). */
  enabled?: boolean
  /** Idle time after the last change (default 3000 ms). */
  debounceMs?: number
  /** Longest time a change waits while editing continues (default 30000 ms). */
  maxDelayMs?: number
  /** Lower bound of the adaptive gap between snapshots (default 3000 ms). */
  minGapMs?: number
  /** Upper bound of the adaptive gap (default 120000 ms). */
  maxGapMs?: number
  /** Larger snapshots turn autosave off for the document (default 512 MiB). */
  maxBytes?: number
  /** Consecutive failures before the chip says "Autosave isn't working" (default 3). */
  failureThreshold?: number
  /** A snapshot or its write that takes longer counts as failed, so one stuck call can't stop autosave (default 120000 ms). */
  snapshotTimeoutMs?: number
}

/** Options for a DocumentSession. */
export interface DocumentSessionOptions {
  /** Noun for one document ("spreadsheet"); defaults to the catalog kind of the module. */
  kind?: string
  /** Workspace id; defaults to io.module. */
  module?: ModuleId
  /**
   * The id main's document registry gave this document when the workspace
   * opened it (openDocument() returns it; pass it through the open IPC).
   * Required for a document opened from a file: save requests and recovery
   * copies name the document by this id, and an unknown id is treated as a
   * new untitled document. Leave it out only for a new untitled document;
   * the session then makes one up and main registers it on first use.
   */
  docId?: string
  /** The bound file is read-only; the chip says "Read-only, use Save As". */
  readOnly?: boolean
  /**
   * Window title "• {name} — {appName}" while dirty (design §3.6 rule 7).
   * true (default when a document exists) sets document.title; a function
   * receives the title instead; false leaves the title alone.
   */
  windowTitle?: boolean | ((title: string) => void)
  /** Flush the journal on window blur and when the page is hidden (default true when a window exists). */
  windowEvents?: boolean
  /**
   * Answer main's close-query, save-now, discard and recovery-flush requests
   * (default true). One session per window answers them: dispose() the old
   * session before creating one for the next document, and pass false for any
   * other session the same window keeps.
   */
  handleRequests?: boolean
  recovery?: RecoveryCadence
  /** Longest wait for a running export or print before a save goes ahead anyway (default 120000 ms). */
  activityWaitMs?: number
  timers?: SessionTimers
  /** Receives problems the session handled itself (recovery write failures and the like). */
  log?: (message: string, detail?: unknown) => void
}

/** The session's lossy state: the bound file was saved in a format that doesn't keep everything. */
export interface LossyState {
  readonly format: string
  readonly formatLabel: string
  readonly lost: readonly string[]
  readonly fullFormat: string | null
  readonly fullLabel: string | null
  /** Revision of the lossy save. A full-format save at or after it clears the state. */
  readonly revision: number
}

/** Autosave health shown by the chip. */
export type AutosaveState = 'on' | 'off' | 'failing'

/** What the UI renders. The object is replaced only when something visible changes. */
export interface SessionState {
  readonly dirty: boolean
  /** A save is running (including its failure prompts). */
  readonly saving: boolean
  /** Another save is queued behind the running one. */
  readonly queued: boolean
  /** An export or print is running. */
  readonly activity: ActivityKind | null
  readonly lossy: LossyState | null
  /** The status chip's catalog key (status.*), or null for no chip. */
  readonly status: StatusKey | null
  readonly statusText: string
  /** Restored from the recovery journal and not saved since. */
  readonly recovered: boolean
  readonly readOnly: boolean
  readonly autosave: AutosaveState
  /** The last failed save, until a save succeeds. Clicking the chip calls reopenFailure(). */
  readonly lastFailure: IoFailure | null
}

/** The renderer's answer to main's close-query (design §3.7). */
export interface CloseQueryAnswer {
  dirty: boolean
  saving: boolean
  lossy: LossyState | null
  title: string
  kind: string
  docId: string
  /** Why pending edits could not be committed, when they could not. */
  blocked?: string
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const DEFAULT_CADENCE = Object.freeze({
  enabled: true,
  debounceMs: 3000,
  maxDelayMs: 30000,
  minGapMs: 3000,
  maxGapMs: 120000,
  maxBytes: 512 * 1024 * 1024,
  failureThreshold: 3,
  snapshotTimeoutMs: 120000,
})

/** Longest a flush waits when main asks for one (shutdown, close timeout). */
const REQUEST_FLUSH_BUDGET_MS = 2000
/** Most save attempts one command makes through failure prompts. */
const MAX_SAVE_ROUNDS = 25

const defaultTimers: SessionTimers = {
  setTimeout: (callback, ms) => globalThis.setTimeout(callback, ms),
  clearTimeout: (handle) => globalThis.clearTimeout(handle as never),
  now: () => Date.now(),
}

function defaultLog(message: string, detail?: unknown): void {
  console.warn('[simple-io]', message, detail === undefined ? '' : detail)
}

function newDocId(): string {
  const api = (globalThis as { crypto?: Crypto }).crypto
  if (api && typeof api.randomUUID === 'function') return api.randomUUID()
  const bytes = new Uint8Array(16)
  if (api && typeof api.getRandomValues === 'function') api.getRandomValues(bytes)
  else for (let index = 0; index < bytes.length; index += 1) bytes[index] = Math.floor(Math.random() * 256)
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('')
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

function cadenceFrom(options: RecoveryCadence | undefined): Required<RecoveryCadence> {
  const cadence: Required<RecoveryCadence> = { ...DEFAULT_CADENCE }
  if (!options) return cadence
  if (typeof options.enabled === 'boolean') cadence.enabled = options.enabled
  const numbers = ['debounceMs', 'maxDelayMs', 'minGapMs', 'maxGapMs', 'maxBytes', 'failureThreshold', 'snapshotTimeoutMs'] as const
  for (const key of numbers) {
    const value = options[key]
    if (typeof value === 'number' && Number.isFinite(value) && value >= 0) cadence[key] = value
  }
  cadence.maxGapMs = Math.max(cadence.maxGapMs, cadence.minGapMs)
  cadence.failureThreshold = Math.max(1, Math.round(cadence.failureThreshold))
  return cadence
}

function hasDom(): boolean {
  return typeof window !== 'undefined' && typeof document !== 'undefined'
}

/**
 * The window title for a document: "• Budget.xlsx — Simple Spreadsheets"
 * while it has unsaved changes (design §3.6 rule 7).
 * @param name document name
 * @param dirty unsaved changes
 * @param app workspace name
 */
export function windowTitleFor(name: string, dirty: boolean, app: string): string {
  const base = name && app ? `${name} — ${app}` : name || app
  return dirty ? `• ${base}` : base
}

/**
 * Rough size of a snapshot, used for the "autosave off for very large files" rule.
 * @param content recovery parts
 */
export function snapshotSize(content: { parts: readonly RecoveryPart[] }): number {
  let total = 0
  for (const part of content.parts) {
    if (typeof part.data === 'string') total += part.data.length
    else if (part.data) total += part.data.byteLength
  }
  return total
}

function formatLabelOf(info: { format: string; formatLabel?: string | undefined }): string {
  return info.formatLabel || info.format.toUpperCase()
}

function sameState(a: SessionState, b: SessionState): boolean {
  return a.dirty === b.dirty
    && a.saving === b.saving
    && a.queued === b.queued
    && a.activity === b.activity
    && a.lossy === b.lossy
    && a.status === b.status
    && a.statusText === b.statusText
    && a.recovered === b.recovered
    && a.readOnly === b.readOnly
    && a.autosave === b.autosave
    && a.lastFailure === b.lastFailure
}

interface QueuedSave {
  mode: SaveMode
  format: string | undefined
  promise: Promise<boolean>
  resolve: (value: boolean) => void
}

function queuedSave(mode: SaveMode, format: string | undefined): QueuedSave {
  let resolve: (value: boolean) => void = () => {}
  const promise = new Promise<boolean>((done) => { resolve = done })
  return { mode, format, promise, resolve }
}

// ---------------------------------------------------------------------------
// DocumentSession
// ---------------------------------------------------------------------------

/**
 * Save, close and recovery controller for one open document (design §3.6).
 *
 * Workspaces call noteChange() on every model mutation, save() for Save /
 * Save As / Save a Copy, prepare() before export, print or an open that
 * replaces the document, and track() around exports and prints.
 */
export class DocumentSession {
  readonly adapter: DocumentAdapter
  readonly io: SimpleIO | null
  readonly module: ModuleId | null
  readonly kind: string

  #docId: string
  #revision = 0
  #savedRevision = 0
  #everSaved = false
  #readOnly: boolean
  #recovered = false
  /**
   * The bound file does not hold the model: a restored recovery copy, or a
   * file an export or copy replaced. Until a save writes the file, the
   * adapter's isPristine() is ignored and saves are never sent as pristine.
   */
  #mustWrite = false
  #lossy: LossyState | null = null
  #lastFailure: IoFailure | null = null
  #lastRequest: SaveRequest | null = null
  #finishing = false
  #disposed = false

  #running: Promise<boolean> | null = null
  /** Whether the last save wrote a file (true) or failed, was canceled or blocked (false). */
  #lastWrote = false
  #queue: QueuedSave[] = []
  #activities = new Map<number, ActivityKind>()
  #activitySeq = 0
  #activityWaiters: Array<() => void> = []
  #waitingFor: ActivityKind | null = null
  readonly #activityWaitMs: number

  readonly #cadence: Required<RecoveryCadence>
  readonly #timers: SessionTimers
  readonly #log: (message: string, detail?: unknown) => void
  #autosave: AutosaveState = 'on'
  #snapshotRevision = 0
  #journalHasEntry = false
  #pendingSince: number | null = null
  #timer: unknown = null
  #timerDue = 0
  #snapshotLoop: Promise<void> | null = null
  #snapshotWanted = false
  #lastSnapshotStartedAt: number | null = null
  #lastSnapshotAt: number | null = null
  #gapMs: number
  #failures = 0
  #retryNotBefore = 0
  #flushReason = 'timer'

  #listeners = new Set<() => void>()
  #state: SessionState
  #cleanups: Array<() => void> = []
  readonly #titleTarget: ((title: string) => void) | null

  /**
   * @param adapter the workspace side (commit, save, snapshot, restore, toasts)
   * @param io window.simpleIO; null runs without prompts or recovery (tests, a plain browser)
   * @param options see DocumentSessionOptions
   */
  constructor(adapter: DocumentAdapter, io: SimpleIO | null, options: DocumentSessionOptions = {}) {
    this.adapter = adapter
    this.io = io
    this.module = options.module || (io ? io.module : null)
    this.kind = options.kind || (this.module ? kindName(this.module) : '') || 'document'
    this.#docId = options.docId || newDocId()
    this.#readOnly = Boolean(options.readOnly)
    this.#activityWaitMs = typeof options.activityWaitMs === 'number' && options.activityWaitMs >= 0 ? options.activityWaitMs : 120000
    this.#cadence = cadenceFrom(options.recovery)
    this.#timers = options.timers || defaultTimers
    this.#log = options.log || defaultLog
    this.#gapMs = this.#cadence.minGapMs
    if (!io) this.#cadence.enabled = false

    if (typeof options.windowTitle === 'function') this.#titleTarget = options.windowTitle
    else if (options.windowTitle !== false && hasDom()) this.#titleTarget = (title) => { document.title = title }
    else this.#titleTarget = null

    this.#state = this.#computeState()
    this.#updateTitle()
    if (io && options.handleRequests !== false) this.#bindRequests(io)
    if (options.windowEvents !== false && hasDom()) this.#bindWindowEvents()
  }

  // -- Read-only views -------------------------------------------------------

  /** Recovery journal id of this document. Changes once when a recovery entry is restored into the session. */
  get docId(): string {
    return this.#docId
  }

  /** Increments on every noteChange(). */
  get revision(): number {
    return this.#revision
  }

  /** The highest revision a completed Save or Save As wrote. */
  get savedRevision(): number {
    return this.#savedRevision
  }

  /** Unsaved changes exist: revision differs from savedRevision, unless the adapter reports the model pristine. */
  get dirty(): boolean {
    return this.#computeDirty()
  }

  /** The current state object (same object until something visible changes). */
  get state(): SessionState {
    return this.#state
  }

  /** When the last recovery copy was written (ms since the epoch), or null. Used by crash screens. */
  get lastRecoveryAt(): number | null {
    return this.#lastSnapshotAt
  }

  /**
   * Listens for state changes. Returns the unsubscribe function, so it plugs
   * straight into React's useSyncExternalStore(session.subscribe, () => session.state).
   */
  subscribe = (listener: () => void): (() => void) => {
    this.#listeners.add(listener)
    return () => { this.#listeners.delete(listener) }
  }

  // -- Changes ---------------------------------------------------------------

  /**
   * Call on EVERY model mutation, including undo and redo, from the place the
   * model changes (not from DOM guesses). Cheap: increments the revision,
   * schedules the recovery journal and notifies listeners only when the
   * visible state changes.
   */
  noteChange(): void {
    if (this.#disposed) return
    this.#revision += 1
    if (this.#recoveryActive()) {
      if (this.#pendingSince === null) this.#pendingSince = this.#timers.now()
      this.#schedule()
    }
    this.#emit()
  }

  /**
   * Call when the document's own file was replaced by something other than
   * the document (an export or Save a Copy written over it; main reports
   * ownFileChanged). The document becomes changed, so closing asks and the
   * next Save writes the file, even if isPristine() says the model is unchanged.
   */
  noteOwnFileChanged(): void {
    if (this.#disposed) return
    this.#mustWrite = true
    this.noteChange()
  }

  /**
   * Marks the bound file read-only (or writable again).
   * @param readOnly true when the file can only be saved with Save As
   */
  setReadOnly(readOnly: boolean): void {
    this.#readOnly = Boolean(readOnly)
    this.#emit()
  }

  // -- Save ------------------------------------------------------------------

  /**
   * Saves the document. Commits pending edits first, captures the revision,
   * asks the adapter to save, and handles failures with the catalog prompts
   * (Try Again, Save As in the fallback folder, Replace, Save Here Again, …).
   * While a save runs, further calls queue one follow-up per mode and format.
   * @param mode "save" (default), "save-as" or "save-copy"
   * @param format registry id of the target format, when the caller chose one
   * @returns true when everything up to now is on disk (for "save-copy": in the copy)
   */
  save(mode: SaveMode = 'save', format?: string): Promise<boolean> {
    if (this.#disposed) return Promise.resolve(false)
    if (this.#running) {
      const existing = this.#queue.find((entry) => entry.mode === mode && entry.format === format)
      if (existing) return existing.promise
      const entry = queuedSave(mode, format)
      this.#queue.push(entry)
      this.#emit()
      return entry.promise
    }
    return this.#start(mode, format, null)
  }

  /**
   * Shows the last save failure's prompt again (the "Not saved" chip's click)
   * and continues with the answer. Without a failure, it saves.
   * @returns the same as save()
   */
  reopenFailure(): Promise<boolean> {
    const failure = this.#lastFailure
    const request = this.#lastRequest
    if (!failure || !request) return this.save()
    if (this.#running) return this.save(request.mode, request.format)
    return this.#start(request.mode, request.format, { failure, request })
  }

  /**
   * Commits pending edits before a non-save action (design §3.1 step 1).
   * Before print and before an open that replaces this document it also
   * writes a recovery copy right away (design §4.2).
   * @param action "export", "print" or "replace"
   * @returns false when the commit was blocked; the user has been told why
   */
  async prepare(action: SessionAction): Promise<boolean> {
    if (this.#disposed) return false
    if (!(await this.#commit(true))) return false
    if (action === 'print' || action === 'replace') await this.flushRecovery(action, REQUEST_FLUSH_BUDGET_MS)
    return true
  }

  /**
   * Runs an export or print. A save requested meanwhile waits for it (the
   * chip says why) instead of being dropped; recovery copies continue.
   * Call prepare() first.
   * @param action "export" or "print"
   * @param work the export or print
   * @returns what `work` returns
   */
  async track<T>(action: ActivityKind, work: () => Promise<T> | T): Promise<T> {
    const id = ++this.#activitySeq
    this.#activities.set(id, action)
    this.#emit()
    try {
      return await work()
    } finally {
      this.#activities.delete(id)
      for (const wake of this.#activityWaiters.splice(0)) wake()
      this.#emit()
    }
  }

  // -- Close, discard, recovery ---------------------------------------------

  /**
   * Answers main's close-query (design §3.7): commits pending edits and
   * reports whether closing needs a decision.
   */
  async closeQuery(): Promise<CloseQueryAnswer> {
    let blocked: string | undefined
    if (!this.#running) {
      const outcome = await this.#runCommit()
      if (outcome !== true) blocked = outcome
    }
    const saving = Boolean(this.#running) || this.#queue.length > 0
    if (saving && !this.#finishing) {
      this.#finishing = true
      this.#emit()
    }
    const answer: CloseQueryAnswer = {
      dirty: this.#computeDirty() || blocked !== undefined,
      saving,
      lossy: this.#lossy,
      title: this.#title(),
      kind: this.kind,
      docId: this.#docId,
    }
    if (blocked !== undefined) answer.blocked = blocked
    return answer
  }

  /**
   * "Don't Save" or "Discard": drops this document's recovery copy. Changes
   * made afterwards are journaled again.
   */
  async discard(): Promise<void> {
    this.#clearTimer()
    this.#pendingSince = null
    this.#snapshotWanted = false
    if (this.#snapshotLoop) {
      try { await this.#snapshotLoop } catch { /* logged by the loop */ }
    }
    this.#snapshotRevision = this.#revision
    this.#journalHasEntry = false
    if (this.io) {
      try {
        await this.io.recovery.discard(this.#docId, this.#revision)
      } catch (error) {
        this.#log('Could not discard the recovery copy.', error)
      }
    }
    this.#emit()
  }

  /**
   * Writes a recovery copy now if there are changes the journal doesn't have.
   * Used on blur, before print and open-replace, by error boundaries, on
   * Windows shutdown and when a close request times out.
   * @param reason short label for logs
   * @param deadlineMs stop waiting after this long (the write may still finish later)
   */
  async flushRecovery(reason = 'flush', deadlineMs?: number): Promise<void> {
    if (!this.#recoveryActive()) return
    this.#clearTimer()
    this.#flushReason = reason
    const work = this.#requestSnapshot()
    if (deadlineMs === undefined) return work
    await new Promise<void>((resolve) => {
      const handle = this.#timers.setTimeout(resolve, Math.max(0, deadlineMs))
      work.then(() => {
        this.#timers.clearTimeout(handle)
        resolve()
      }, () => {
        this.#timers.clearTimeout(handle)
        resolve()
      })
    })
  }

  /** True when a recovery entry can replace this window's content (nothing unsaved, no save running). */
  canRestoreHere(): boolean {
    return !this.#computeDirty() && !this.#running
  }

  /**
   * Restores a recovery entry into this session (design §4.4). The document
   * becomes dirty with the chip "Recovered, not saved yet", keeps journaling
   * under the entry's id, and the entry is removed once a save succeeds.
   * Only after the adapter rebuilt the model does main bind the document to
   * the entry's source path with the stamp taken when the snapshot was made
   * (recovery.adopt), so a stale copy never overwrites newer work and a
   * failed restore leaves nothing bound. Until a save writes the file the
   * document stays changed, whatever the adapter's isPristine() says.
   * Refused while this document has unsaved changes or a save runs (see
   * canRestoreHere()): restore the entry in a new window instead.
   * @param entry the entry chosen on the welcome card or banner
   * @returns false when it was refused, or could not be read or restored (the user has been told)
   */
  async restore(entry: RecoveryEntry): Promise<boolean> {
    if (!this.io || this.#disposed) return false
    if (!this.canRestoreHere()) {
      this.#log('Restore refused: this document has unsaved changes. Restore the entry in a new window.')
      return false
    }
    let payload: RecoveryPayload
    try {
      payload = await this.io.recovery.read(entry.id)
      if (!payload || typeof payload !== 'object' || payload.ok === false) {
        const detail = payload && typeof payload === 'object' && payload.ok === false ? `${payload.code}${payload.reason ? `: ${payload.reason}` : ''}` : 'no payload'
        throw new Error(`The recovery copy could not be read (${detail}).`)
      }
      await this.adapter.restoreFromRecovery(entry, payload)
    } catch (error) {
      this.#log('Could not restore the recovery copy.', error)
      this.#notify('warning', catalogText('recovery.cannotRestore'))
      return false
    }
    // Now main binds this window to the entry's id, path and snapshot-time stamp.
    let adopted = (typeof payload.docId === 'string' && payload.docId) || entry.docId || entry.id
    if (typeof this.io.recovery.adopt === 'function') {
      let taken: { ok: boolean } | null = null
      try {
        taken = await this.io.recovery.adopt(adopted)
      } catch (error) {
        this.#log('Could not take the recovery copy over.', error)
      }
      if (!taken || taken.ok !== true) {
        // The content is restored but can't be bound to its file (another window
        // took it meanwhile): keep it as a new untitled document, never under the
        // id of what this window showed before.
        this.#log('The recovered document is kept as a new untitled document.')
        adopted = newDocId()
      }
    }
    if (adopted !== this.#docId) {
      if (this.#journalHasEntry) {
        try { await this.io.recovery.discard(this.#docId, this.#revision) } catch (error) { this.#log('Could not discard the previous recovery copy.', error) }
      }
      this.#docId = adopted
    }
    this.#clearTimer()
    this.#revision = Math.max(this.#revision, this.#savedRevision, typeof entry.revision === 'number' ? entry.revision : 0) + 1
    // The journal already holds exactly what was restored.
    this.#snapshotRevision = this.#revision
    this.#journalHasEntry = true
    this.#pendingSince = null
    this.#recovered = true
    this.#mustWrite = true
    this.#emit()
    return true
  }

  /**
   * Stops timers and request handlers. Does not touch the journal: a
   * document that still has unsaved work keeps its recovery copy.
   */
  dispose(): void {
    if (this.#disposed) return
    this.#disposed = true
    this.#clearTimer()
    for (const cleanup of this.#cleanups.splice(0)) {
      try { cleanup() } catch (error) { this.#log('Could not remove a listener.', error) }
    }
    this.#listeners.clear()
  }

  // -- Internals: state ------------------------------------------------------

  #title(): string {
    try {
      return String(this.adapter.title() || '')
    } catch {
      return ''
    }
  }

  #computeDirty(): boolean {
    if (this.#revision === this.#savedRevision) return false
    // The file doesn't hold the model (restored copy, file replaced by an export): pristine or not, it must be written.
    if (this.#mustWrite) return true
    if (this.adapter.isPristine) {
      try {
        if (this.adapter.isPristine()) return false
      } catch (error) {
        this.#log('isPristine() failed; treating the document as changed.', error)
      }
    }
    return true
  }

  #computeStatus(dirty: boolean, activity: ActivityKind | null): { key: StatusKey | null; vars: Record<string, string> } {
    const none: Record<string, string> = {}
    if (this.#running) {
      if (this.#waitingFor === 'export') return { key: 'waitingForExport', vars: none }
      if (this.#waitingFor === 'print') return { key: 'waitingForPrint', vars: none }
      if (this.#finishing) return { key: 'finishingSave', vars: none }
      if (this.#queue.length) return { key: 'willSaveAgain', vars: none }
      return { key: 'saving', vars: none }
    }
    if (activity === 'export') return { key: 'exporting', vars: none }
    if (dirty) {
      if (this.#lastFailure) return { key: 'notSaved', vars: none }
      if (this.#autosave === 'failing') return { key: 'autosaveFailing', vars: none }
      if (this.#autosave === 'off') return { key: 'autosaveOff', vars: none }
      if (this.#recovered) return { key: 'recovered', vars: none }
      if (this.#readOnly) return { key: 'readOnly', vars: none }
      return { key: 'unsaved', vars: none }
    }
    if (this.#lossy) return { key: 'savedAs', vars: { formatLabel: this.#lossy.formatLabel } }
    if (this.#readOnly) return { key: 'readOnly', vars: none }
    if (this.#everSaved) return { key: 'saved', vars: none }
    return { key: null, vars: none }
  }

  #computeState(): SessionState {
    const dirty = this.#computeDirty()
    const activity = this.#activities.size ? this.#activities.values().next().value ?? null : null
    const status = this.#computeStatus(dirty, activity)
    return Object.freeze({
      dirty,
      saving: Boolean(this.#running),
      queued: this.#queue.length > 0,
      activity,
      lossy: this.#lossy,
      status: status.key,
      statusText: statusText(status.key, status.vars),
      recovered: this.#recovered,
      readOnly: this.#readOnly,
      autosave: this.#autosave,
      lastFailure: this.#lastFailure,
    })
  }

  #emit(): void {
    const next = this.#computeState()
    if (sameState(next, this.#state)) return
    const titleChanged = next.dirty !== this.#state.dirty
    this.#state = next
    if (titleChanged) this.#updateTitle()
    for (const listener of [...this.#listeners]) {
      try { listener() } catch (error) { this.#log('A session listener failed.', error) }
    }
  }

  #updateTitle(): void {
    if (!this.#titleTarget) return
    try {
      this.#titleTarget(windowTitleFor(this.#title(), this.#state.dirty, this.module ? appName(this.module) : ''))
    } catch (error) {
      this.#log('Could not update the window title.', error)
    }
  }

  /** Re-reads the document name for the window title, after the workspace renamed it outside a save. */
  refreshTitle(): void {
    this.#updateTitle()
  }

  #notify(kind: NotifyKind, text: string, actions?: NotifyAction[]): void {
    if (!text) return
    try {
      if (actions && actions.length) this.adapter.notify(kind, text, actions)
      else this.adapter.notify(kind, text)
    } catch (error) {
      this.#log('Could not show a notification.', error)
    }
  }

  #showInFolderAction(path: string): NotifyAction[] {
    const io = this.io
    if (!io || !path) return []
    return [{
      label: catalogText('toast.actions.showInFolder'),
      run: () => { void io.shell.showItem(path).catch((error: unknown) => this.#log('Could not show the file in its folder.', error)) },
    }]
  }

  // -- Internals: commit -----------------------------------------------------

  /** Runs commitPendingEdits(). Returns true, or the message that blocked it. */
  async #runCommit(): Promise<true | string> {
    try {
      const result = await this.adapter.commitPendingEdits()
      if (result && typeof result === 'object' && typeof result.blocked === 'string') {
        if (typeof result.focus === 'function') {
          try { result.focus() } catch (error) { this.#log('Could not focus the blocking input.', error) }
        }
        return result.blocked || catalogText('notices.BLOCKED')
      }
      return true
    } catch (error) {
      this.#log('commitPendingEdits() failed.', error)
      return catalogText('notices.BLOCKED')
    }
  }

  async #commit(tell: boolean): Promise<boolean> {
    const outcome = await this.#runCommit()
    if (outcome === true) return true
    if (tell) this.#notify('warning', outcome)
    return false
  }

  // -- Internals: saving -----------------------------------------------------

  #start(mode: SaveMode, format: string | undefined, resume: { failure: IoFailure; request: SaveRequest } | null): Promise<boolean> {
    this.#lastWrote = false
    const run = this.#perform(mode, format, resume)
    this.#running = run
    this.#emit()
    const settle = () => {
      if (this.#running === run) this.#running = null
      this.#finishing = false
      this.#emit()
      this.#drainQueue(this.#lastWrote)
    }
    run.then(settle, (error: unknown) => {
      this.#log('Saving stopped unexpectedly.', error)
      this.#lastWrote = false
      settle()
    })
    return run.catch(() => false)
  }

  /**
   * Runs the saves queued behind the one that just ended.
   * @param wrote the save that ended wrote a file (false: it failed, was canceled or was blocked)
   */
  #drainQueue(wrote: boolean): void {
    while (this.#queue.length && !this.#running) {
      const next = this.#queue.shift()
      if (!next) break
      if (!wrote) {
        // The user just answered a failure prompt, canceled a dialog, or was
        // told why the save couldn't run; a queued Ctrl+S must not reopen
        // that unasked. The chip keeps showing the state.
        next.resolve(false)
        continue
      }
      if (next.mode === 'save' && next.format === undefined && !this.#computeDirty()) {
        next.resolve(true)
        continue
      }
      this.#start(next.mode, next.format, null).then(next.resolve, () => next.resolve(false))
    }
    this.#emit()
  }

  async #waitForActivities(): Promise<void> {
    if (!this.#activities.size) return
    const deadline = this.#timers.now() + this.#activityWaitMs
    while (this.#activities.size && this.#timers.now() < deadline) {
      this.#waitingFor = this.#activities.values().next().value ?? null
      this.#emit()
      await new Promise<void>((resolve) => {
        const handle = this.#timers.setTimeout(resolve, Math.max(0, deadline - this.#timers.now()))
        this.#activityWaiters.push(() => {
          this.#timers.clearTimeout(handle)
          resolve()
        })
      })
    }
    if (this.#activities.size) this.#log('An export or print is still running; saving anyway.')
    this.#waitingFor = null
    this.#emit()
  }

  async #perform(mode: SaveMode, format: string | undefined, resume: { failure: IoFailure; request: SaveRequest } | null): Promise<boolean> {
    await this.#waitForActivities()
    const nameBefore = this.#title()
    let request: SaveRequest = resume ? { ...resume.request } : { mode, revision: this.#revision, pristine: false, docId: this.#docId }
    if (!resume && format !== undefined) request.format = format
    let failure: IoFailure | null = resume ? resume.failure : null

    for (let round = 0; round < MAX_SAVE_ROUNDS; round += 1) {
      if (!failure) {
        if (!(await this.#commit(true))) return false
        request = { ...request, revision: this.#revision, pristine: !this.#computeDirty(), docId: this.#docId }
        let result: IoResult
        try {
          result = toIoResult(await this.adapter.save(request))
        } catch (error) {
          result = failureFromError(error)
        }
        if (result.ok) return this.#succeeded(result, request, nameBefore)
        if (result.code === 'CANCELED') return false
        if (result.code === 'BLOCKED') {
          this.#notify('warning', result.message || catalogText('notices.BLOCKED'))
          return false
        }
        failure = result
      }
      this.#lastFailure = failure
      this.#lastRequest = request
      this.#emit()
      const answer = await promptFailure(this.io, failure)
      if (answer === 'show-original') {
        const original = failure.backupPath || failure.asidePath
        if (original && this.io) {
          try { await this.io.shell.showItem(original) } catch (error) { this.#log('Could not show the original file.', error) }
        }
        continue
      }
      const next = nextRequest(answer, request, failure)
      if (!next) return false
      request = next
      failure = null
    }
    return false
  }

  #succeeded(result: IoSuccess, request: SaveRequest, nameBefore: string): boolean {
    this.#lastWrote = true
    const revision = request.revision
    const lossy: LossInfo | null = result.lossy && Array.isArray(result.lossy.lost) && result.lossy.lost.length
      ? { ...result.lossy, formatLabel: result.lossy.formatLabel || result.formatLabel || formatLabelOf(result.lossy) }
      : null
    if (request.mode !== 'save-copy') {
      this.#savedRevision = Math.max(this.#savedRevision, revision)
      this.#everSaved = true
      this.#recovered = false
      if (result.strategy !== 'unchanged') this.#mustWrite = false
      if (lossy) {
        this.#lossy = Object.freeze({
          format: lossy.format,
          formatLabel: formatLabelOf(lossy),
          lost: Object.freeze([...lossy.lost.map(String)]),
          fullFormat: lossy.fullFormat || null,
          fullLabel: lossy.fullLabel || null,
          revision,
        })
      } else if (this.#lossy && this.#lossy.revision <= revision) {
        this.#lossy = null
      }
    } else if (!lossy && this.#lossy && this.#lossy.revision <= revision) {
      // A full-format copy holds everything the lossy file left out.
      this.#lossy = null
    }
    this.#lastFailure = null
    this.#lastRequest = null
    if (request.mode !== 'save-copy' && !lossy) this.#dropJournalUpTo(revision)
    const copyComplete = this.#revision === revision
    if (request.mode === 'save-copy' && result.ownFileChanged) {
      // The copy went over the document's own file, which no longer holds the document.
      this.#mustWrite = true
      this.#revision += 1
      if (this.#recoveryActive()) {
        if (this.#pendingSince === null) this.#pendingSince = this.#timers.now()
        this.#schedule()
      }
    }
    this.#announce(result, request, lossy, nameBefore)
    this.#emit()
    this.#updateTitle()
    if (request.mode === 'save-copy') return copyComplete
    return !this.#computeDirty()
  }

  #announce(result: IoSuccess, request: SaveRequest, lossy: LossInfo | null, nameBefore: string): void {
    const name = result.name || baseName(result.path)
    const folder = folderLabel(folderOf(result.path))
    const show = this.#showInFolderAction(result.path)
    const newerPending = request.mode !== 'save-copy' && this.#revision > request.revision && this.#computeDirty()
    if (lossy) {
      this.#notify('info', catalogText('toast.savedLossy', {
        formatLabel: formatLabelOf(lossy),
        lostShort: capitalize(lossy.lostShort || joinList(lossy.lost)),
      }))
    } else if (result.sibling) {
      const siblingInfo = typeof result.sibling === 'object' ? result.sibling : null
      const original = result.originalName || (siblingInfo && siblingInfo.originalName) || nameBefore
      const key = result.originalFormatLabel ? 'toast.savedModernCopy' : 'toast.savedSibling'
      this.#notify('success', catalogText(key, { siblingName: name, name: original, formatLabel: result.originalFormatLabel }), show)
    } else if (newerPending) {
      this.#notify('info', catalogText('toast.savedNewerPending'))
    } else if (result.folderChanged === true || (result.folderChanged === undefined && request.mode !== 'save')) {
      this.#notify('success', catalogText('toast.savedTo', { folder }), show)
    } else {
      this.#notify('success', catalogText('toast.saved'))
    }
    if (result.backupFailed) this.#notify('info', catalogText('toast.savedNoBackup', { reason: result.backupFailed }))
    for (const warning of result.warnings || []) {
      if (typeof warning === 'string' && warning) this.#notify('warning', warning)
    }
  }

  // -- Internals: recovery journal ------------------------------------------

  #recoveryActive(): boolean {
    return Boolean(this.io) && this.#cadence.enabled && this.#autosave !== 'off' && !this.#disposed
  }

  #clearTimer(): void {
    if (this.#timer !== null) {
      this.#timers.clearTimeout(this.#timer)
      this.#timer = null
    }
  }

  /** Arms the snapshot timer: 3 s idle, at most 30 s after the first unsaved change, never inside the adaptive gap. */
  #schedule(): void {
    if (!this.#recoveryActive() || this.#pendingSince === null) return
    const now = this.#timers.now()
    let due = Math.min(now + this.#cadence.debounceMs, this.#pendingSince + this.#cadence.maxDelayMs)
    if (this.#lastSnapshotStartedAt !== null) due = Math.max(due, this.#lastSnapshotStartedAt + this.#gapMs)
    due = Math.max(due, this.#retryNotBefore)
    if (this.#timer !== null && this.#timerDue === due) return
    this.#clearTimer()
    this.#timerDue = due
    this.#timer = this.#timers.setTimeout(() => {
      this.#timer = null
      void this.#requestSnapshot()
    }, Math.max(0, due - now))
  }

  #requestSnapshot(): Promise<void> {
    this.#snapshotWanted = true
    if (!this.#snapshotLoop) {
      const loop = (async () => {
        try {
          while (this.#snapshotWanted && !this.#disposed) {
            this.#snapshotWanted = false
            await this.#snapshotOnce()
          }
        } finally {
          this.#snapshotLoop = null
        }
      })()
      this.#snapshotLoop = loop
    }
    return this.#snapshotLoop
  }

  async #snapshotOnce(): Promise<void> {
    const io = this.io
    if (!io || !this.#recoveryActive()) return
    if (!this.#computeDirty() && !this.#lossy) {
      // Saved in a full format, or undone back to the saved state: nothing to
      // recover. After a lossy save (CSV, …) the journal keeps the full model.
      this.#pendingSince = null
      if (this.#journalHasEntry) this.#dropJournalUpTo(this.#revision)
      return
    }
    const revision = this.#revision
    if (revision <= this.#snapshotRevision) {
      this.#pendingSince = null
      return
    }
    const firstPending = this.#pendingSince
    this.#pendingSince = null
    const started = this.#timers.now()
    const reason = this.#flushReason
    this.#flushReason = 'timer'
    this.#lastSnapshotStartedAt = started
    let failed = false
    try {
      const content = await this.#withTimeout(this.adapter.snapshotForRecovery(), 'The recovery snapshot')
      if (!content || !Array.isArray(content.parts) || !content.parts.length) {
        // Nothing to keep yet; the next change (or flush) tries again.
        return
      }
      if (snapshotSize(content) > this.#cadence.maxBytes) {
        this.#autosave = 'off'
        this.#log('Autosave is off for this document because it is too large.')
        return
      }
      const snapshot: RecoverySnapshot = {
        ...content,
        docId: this.#docId,
        revision,
        title: content.title || this.#title(),
        kind: content.kind || this.kind,
      }
      if (this.module) snapshot.module = this.module
      const written = await this.#withTimeout(io.recovery.write(snapshot), 'Writing the recovery copy')
      if (written && typeof written === 'object' && written.ok === false && !written.dropped) {
        if (written.code === 'TOO_LARGE') {
          this.#autosave = 'off'
          this.#log('Autosave is off for this document because it is too large.')
          return
        }
        if (written.code === 'UNSUPPORTED') {
          // This workspace keeps no recovery copies.
          this.#cadence.enabled = false
          return
        }
        throw new Error(`The recovery copy was not written (${written.code || 'unknown'}${written.message ? `: ${written.message}` : ''}).`)
      }
      // A write main dropped needs no retry: "discarded" means a save already
      // covered this revision, "stale" that the journal holds a newer one.
      const dropped = written && typeof written === 'object' && written.ok === false ? written.dropped : undefined
      this.#snapshotRevision = Math.max(this.#snapshotRevision, revision)
      if (dropped !== 'discarded') this.#journalHasEntry = true
      if (!dropped) this.#lastSnapshotAt = this.#timers.now()
      this.#failures = 0
      this.#retryNotBefore = 0
      if (this.#autosave === 'failing') this.#autosave = 'on'
    } catch (error) {
      failed = true
      this.#failures += 1
      this.#log(`Writing the recovery copy failed (${reason}; ${this.#failures} in a row).`, error)
      if (this.#failures >= this.#cadence.failureThreshold) this.#autosave = 'failing'
      this.#pendingSince = firstPending ?? started
    } finally {
      const now = this.#timers.now()
      const duration = Math.max(0, now - started)
      this.#gapMs = clamp(8 * duration, this.#cadence.minGapMs, this.#cadence.maxGapMs)
      if (failed) {
        const backoff = this.#cadence.debounceMs * 2 ** Math.min(this.#failures - 1, 6)
        this.#retryNotBefore = now + Math.min(backoff, this.#cadence.maxGapMs)
      }
      this.#emit()
      if (this.#pendingSince !== null && this.#revision > this.#snapshotRevision) this.#schedule()
    }
  }

  /** Rejects when `promise` takes longer than the snapshot timeout; the late result is ignored. */
  #withTimeout<T>(promise: Promise<T>, what: string): Promise<T> {
    const ms = this.#cadence.snapshotTimeoutMs
    if (!(ms > 0)) return promise
    return new Promise<T>((resolve, reject) => {
      const handle = this.#timers.setTimeout(() => reject(new Error(`${what} did not finish within ${Math.round(ms / 1000)} s.`)), ms)
      promise.then((value) => {
        this.#timers.clearTimeout(handle)
        resolve(value)
      }, (error: unknown) => {
        this.#timers.clearTimeout(handle)
        reject(error)
      })
    })
  }

  /**
   * Removes the journal entry up to a revision; bookkeeping errors never fail
   * a save. Main also drops any snapshot at or below that revision that is
   * still on its way, so a stale copy can't come back after the save.
   */
  #dropJournalUpTo(revision: number): void {
    const io = this.io
    if (!io || (!this.#journalHasEntry && !this.#cadence.enabled)) return
    if (revision >= this.#snapshotRevision) this.#journalHasEntry = false
    const docId = this.#docId
    void io.recovery.discard(docId, revision).catch((error: unknown) => this.#log('Could not remove the recovery copy.', error))
  }

  // -- Internals: requests from main and window events ----------------------

  #bindRequests(io: SimpleIO): void {
    const register = (type: RequestType, handler: (payload: unknown) => unknown) => {
      try {
        const off = io.onRequest(type, handler)
        if (typeof off === 'function') this.#cleanups.push(off)
      } catch (error) {
        this.#log(`Could not listen for ${type} requests.`, error)
      }
    }
    register('close-query', () => this.closeQuery())
    register('save-now', (payload) => {
      const request = (payload && typeof payload === 'object' ? payload : {}) as { mode?: unknown; format?: unknown }
      const mode: SaveMode = request.mode === 'save-as' || request.mode === 'save-copy' ? request.mode : 'save'
      return this.save(mode, typeof request.format === 'string' && request.format ? request.format : undefined)
    })
    register('discard', async () => {
      await this.discard()
      return true
    })
    register('recovery-flush', async (payload) => {
      const request = (payload && typeof payload === 'object' ? payload : {}) as { reason?: unknown; budgetMs?: unknown }
      const budget = typeof request.budgetMs === 'number' && request.budgetMs > 0 ? request.budgetMs : REQUEST_FLUSH_BUDGET_MS
      await this.flushRecovery(typeof request.reason === 'string' ? request.reason : 'request', budget)
      return { ok: true, revision: this.#snapshotRevision }
    })
  }

  #bindWindowEvents(): void {
    const onBlur = () => {
      if (this.#computeDirty()) void this.flushRecovery('blur')
    }
    const onVisibility = () => {
      if (document.visibilityState === 'hidden' && this.#computeDirty()) void this.flushRecovery('hidden')
    }
    window.addEventListener('blur', onBlur)
    document.addEventListener('visibilitychange', onVisibility)
    this.#cleanups.push(() => {
      window.removeEventListener('blur', onBlur)
      document.removeEventListener('visibilitychange', onVisibility)
    })
  }
}

/**
 * The next save request after the user answered a failure prompt, or null to stop.
 * Button ids come from io-catalog.json saveFailed.*.
 * @param answer the chosen button id
 * @param request the request that failed
 * @param failure the failure
 */
export function nextRequest(answer: string, request: SaveRequest, failure: IoFailure): SaveRequest | null {
  const asNew: SaveMode = request.mode === 'save-copy' ? 'save-copy' : 'save-as'
  const base: SaveRequest = { mode: request.mode, revision: request.revision, pristine: request.pristine, docId: request.docId }
  if (request.format !== undefined) base.format = request.format
  switch (answer) {
    case 'retry': {
      const retry: SaveRequest = { ...base }
      if (request.force) retry.force = true
      if (request.recreate) retry.recreate = true
      if (request.fallbackReason) retry.fallbackReason = request.fallbackReason
      if (request.failedPath) retry.failedPath = request.failedPath
      return retry
    }
    case 'save-as': {
      const next: SaveRequest = { ...base, mode: asNew, fallbackReason: failure.code }
      if (failure.path) next.failedPath = failure.path
      return next
    }
    case 'replace':
      return { ...base, force: true }
    case 'recreate':
      return { ...base, recreate: true }
    // The format was the problem, not the folder: a plain Save As in the
    // document's folder (a fallback reason would move the dialog elsewhere).
    case 'save-other-format':
      return { mode: asNew, revision: request.revision, pristine: request.pristine, docId: request.docId }
    // "Save as .xlsx": main writes the edits next to the original, which stays unchanged.
    case 'save-alternative': {
      const next: SaveRequest = { ...base }
      if (failure.altFormat) next.format = failure.altFormat
      return next
    }
    default:
      return null
  }
}

// ---------------------------------------------------------------------------
// Recovery card and banners (design §4.4)
// ---------------------------------------------------------------------------

/** One row of the "Recover unsaved work" card, ready to render. */
export interface RecoveryRow {
  id: string
  title: string
  /** "Unsaved changes from 14:02". */
  subtitle: string
  /** Folder of the source file ("" for an untitled document). */
  folder: string
  /** Extra lines: the file changed since, the file is gone, or it can't be restored. */
  notes: string[]
  restorable: boolean
  untitled: boolean
  entry: RecoveryEntry
}

function toDate(value: string | number | null | undefined): Date | null {
  if (value === undefined || value === null || value === '') return null
  const date = new Date(value)
  return Number.isNaN(date.getTime()) ? null : date
}

/**
 * Formats a snapshot time for the card and banners: the time of day for
 * today, otherwise the date and time, in the user's locale.
 * @param value ISO string or epoch milliseconds
 * @param now the current time (for tests)
 */
export function formatRecoveryTime(value: string | number | null | undefined, now: Date = new Date()): string {
  const date = toDate(value)
  if (!date) return ''
  const sameDay = date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth() && date.getDate() === now.getDate()
  if (sameDay) return date.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })
  const options: Intl.DateTimeFormatOptions = { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }
  if (date.getFullYear() !== now.getFullYear()) options.year = 'numeric'
  return date.toLocaleString(undefined, options)
}

/**
 * Builds the welcome-card row for a recovery entry.
 * @param entry from simpleIO.recovery.list()
 * @param now the current time (for tests)
 */
export function describeRecoveryEntry(entry: RecoveryEntry, now: Date = new Date()): RecoveryRow {
  const untitled = !entry.sourcePath
  const notes: string[] = []
  const restorable = entry.restorable !== false
  if (!restorable) notes.push(catalogText('recovery.cannotRestore'))
  else if (entry.sourceMissing) notes.push(catalogText('recovery.sourceMissing'))
  else if (entry.sourceChanged) notes.push(catalogText('recovery.changedSince'))
  const title = entry.title || (entry.sourcePath ? baseName(entry.sourcePath) : '')
    || catalogText('recovery.untitled', { kind: entry.kind || (entry.module ? kindName(entry.module) : '') || 'document' })
  return {
    id: entry.id,
    title,
    subtitle: catalogText('recovery.rowSubtitle', { time: formatRecoveryTime(entry.updatedAt ?? entry.createdAt, now) }),
    folder: entry.sourcePath ? folderOf(entry.sourcePath) : '',
    notes: notes.filter(Boolean),
    restorable,
    untitled,
    entry,
  }
}

/**
 * Lists recovery entries to offer: orphaned ones (from a crash, a kill or a
 * power loss), newest first. Never throws; a failed listing offers nothing.
 * @param io the bridge
 * @param filter optional source path: only entries for that file
 */
export async function listRecoveryEntries(io: SimpleIO | null, filter: { sourcePath?: string } = {}): Promise<RecoveryEntry[]> {
  if (!io) return []
  let entries: RecoveryEntry[]
  try {
    entries = await io.recovery.list()
  } catch {
    return []
  }
  if (!Array.isArray(entries)) return []
  const time = (entry: RecoveryEntry) => toDate(entry.updatedAt ?? entry.createdAt)?.getTime() ?? 0
  return entries
    .filter((entry) => entry && typeof entry.id === 'string' && entry.orphaned !== false)
    .filter((entry) => !filter.sourcePath || samePath(entry.sourcePath, filter.sourcePath))
    .sort((a, b) => time(b) - time(a))
}

/**
 * The entry for a file that was just opened, for the banner
 * "You have unsaved changes to this file from {time}." (design §4.4).
 * @param entries from listRecoveryEntries()
 * @param path the opened file
 */
export function findRecoveryForPath(entries: readonly RecoveryEntry[], path: string): RecoveryEntry | null {
  return entries.find((entry) => samePath(entry.sourcePath, path)) || null
}

/**
 * Banner text for an opened file with a pending recovery entry.
 * @param entry the entry
 * @param now the current time (for tests)
 */
export function recoveryBannerText(entry: RecoveryEntry, now: Date = new Date()): string {
  return catalogText('recovery.banner', { time: formatRecoveryTime(entry.updatedAt ?? entry.createdAt, now) })
}

/**
 * Banner text for other recovered documents: "Simple recovered 2 unsaved documents."
 * @param count number of entries
 */
export function orphanBannerText(count: number): string {
  if (count <= 0) return ''
  return count === 1 ? catalogText('recovery.orphanBannerOne') : catalogText('recovery.orphanBanner', { count })
}

/**
 * Removes a recovery entry (Discard on the card or banner). Never throws.
 * @param io the bridge
 * @param entry the entry
 * @returns false when main could not remove it
 */
export async function discardRecoveryEntry(io: SimpleIO | null, entry: RecoveryEntry): Promise<boolean> {
  if (!io) return false
  try {
    await io.recovery.discard(entry.id)
    return true
  } catch {
    return false
  }
}
