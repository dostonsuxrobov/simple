// Vendored from simple/shared/renderer/io-client.ts by simple/scripts/sync-shared.cjs. Do not edit here.

// Typed access to window.simpleIO, the bridge every Simple workspace's preload
// exposes (design §8.1), the shared user-facing strings in io-catalog.json, and
// the drop and paste helpers that keep real file paths (§6.5, §6.6).
//
// Framework-free: DOM types only, no React and no Node globals, so the same
// file compiles in every workspace renderer.
//
// Everything stays inside Simple. A file is handed to the matching Simple
// workspace through openInSimple(); "Show in Folder" (File Explorer) is the
// only hand-off to the operating system. Nothing here talks to the network.

import catalogData from './io-catalog.json'

// ---------------------------------------------------------------------------
// Types shared with the main process
// ---------------------------------------------------------------------------

/** The workspaces that expose window.simpleIO. */
export type ModuleId = 'pdf' | 'calc' | 'docs' | 'image' | 'video'

/** Every result code a shared I/O call can report (io-core.cjs IO_CODES). Save and open share one list. */
export type IoCode =
  | 'CANCELED'
  | 'BLOCKED'
  | 'LOCKED'
  | 'READ_ONLY'
  | 'NO_PERMISSION'
  | 'READ_ONLY_VOLUME'
  | 'FOLDER_MISSING'
  | 'FILE_UNAVAILABLE'
  | 'DISK_FULL'
  | 'NAME_TOO_LONG'
  | 'INVALID_NAME'
  | 'CHANGED_ON_DISK'
  | 'SOURCE_MISSING'
  | 'VERIFY_FAILED'
  | 'VALIDATION_FAILED'
  | 'SERIALIZE_FAILED'
  | 'NEEDS_OFFICE_ENGINE'
  | 'RESTORE_NEEDED'
  | 'NOT_FOUND'
  | 'DAMAGED'
  | 'ENCRYPTED'
  | 'TOO_LARGE'
  | 'UNSUPPORTED'
  | 'UNKNOWN'

/** Every IoCode, for checking values that crossed IPC. */
export const IO_CODES: readonly IoCode[] = Object.freeze([
  'CANCELED', 'BLOCKED', 'LOCKED', 'READ_ONLY', 'NO_PERMISSION', 'READ_ONLY_VOLUME', 'FOLDER_MISSING',
  'FILE_UNAVAILABLE', 'DISK_FULL', 'NAME_TOO_LONG', 'INVALID_NAME', 'CHANGED_ON_DISK', 'SOURCE_MISSING',
  'VERIFY_FAILED', 'VALIDATION_FAILED', 'SERIALIZE_FAILED', 'NEEDS_OFFICE_ENGINE', 'RESTORE_NEEDED',
  'NOT_FOUND', 'DAMAGED', 'ENCRYPTED', 'TOO_LARGE', 'UNSUPPORTED', 'UNKNOWN',
] as const)

/** File identity that main records at open and after every save. The renderer never makes one up. */
export interface Stamp {
  size: number
  mtimeMs: number
  sha256: string
}

/** How safeWriteFile replaced the target (design §3.2). */
export type WriteStrategy = 'rename' | 'swap' | 'in-place' | 'new'

/** What a lossy save or export left out (design §3.4 "Lossy formats"). */
export interface LossInfo {
  /** Registry id of the lossy format that was written, e.g. "csv". */
  format: string
  /** What the format did not keep, e.g. ["formulas", "formatting"]. */
  lost: string[]
  formatLabel?: string
  /** The lost items as one short phrase, e.g. "formulas and formatting". */
  lostShort?: string
  /** The full format that keeps everything, e.g. "xlsx", and its label. */
  fullFormat?: string | null
  fullLabel?: string | null
}

/**
 * A completed save or export (design §3.5): what main's performSave()
 * returns. `path`, `name` and `format` are always present.
 */
export interface IoSuccess {
  ok: true
  path: string
  name: string
  format: string
  formatLabel?: string
  docId?: string
  mode?: string
  /** The revision the request carried. */
  revision?: number | null
  stamp?: Stamp
  /** "unchanged": an untouched document saved to its own file; nothing was written. */
  strategy?: WriteStrategy | 'unchanged'
  attempts?: number
  ms?: number
  /** Items the serializer could not include. */
  warnings?: string[]
  lossy?: LossInfo | null
  /** Copy of the original kept in the versions store before the first overwrite. */
  backupPath?: string
  /** Set when the versions backup failed; the save itself still succeeded. */
  backupFailed?: string
  /** True when the file went to a different folder than the document's file. */
  folderChanged?: boolean
  /** True when the document now points at the written file. */
  rebound?: boolean
  /** Set when the edits went into a new file next to the original, which stayed unchanged (sibling save, §3.4). */
  sibling?: boolean | { originalPath?: string; originalName?: string } | null
  /** Path, name and format label of the original that a sibling save left unchanged. */
  originalPath?: string
  originalName?: string
  originalFormatLabel?: string
  /**
   * An export or Save a Copy replaced the document's own file, which no longer
   * holds the document: the session marks the document changed so the next
   * Save writes it.
   */
  ownFileChanged?: boolean
}

/** A failed save, export or open (design §3.5). Expected failures are returned, never thrown. */
export interface IoFailure {
  ok: false
  code: IoCode
  /** Plain catalog text from main. */
  message?: string
  context?: 'open' | 'save'
  name?: string
  path?: string
  folder?: string
  nearestFolder?: string
  needed?: string
  neededBytes?: number
  drive?: string
  reason?: string
  technical?: string
  backupPath?: string
  asidePath?: string
  formatLabel?: string
  altLabel?: string
  altExt?: string
  altFormat?: string
  ext?: string
  limit?: string
}

/** Result of every save and export call. */
export type IoResult = IoSuccess | IoFailure

/**
 * What the local office-file engine probe reports (office-engine.cjs). The
 * engine is optional and only used when it is already installed on this PC;
 * Simple never downloads or installs anything.
 */
export interface OfficeEngineStatus {
  available: boolean
  source: 'env' | 'user-runtime' | 'app-folder' | 'system' | null
  path: string | null
  version: string | null
  verified: boolean
  reason: string | null
  detail?: string | null
  checkedAt?: number
}

/** Result of simpleIO.capabilities(). */
export interface Capabilities {
  officeEngine: OfficeEngineStatus
  platform: string
  version?: number
  module?: string | null
  appName?: string
  /** True in the unified Simple app (one program for every workspace). */
  unified?: boolean
}

/** Why a save path is being chosen (design §8.1). */
export type SavePurpose = 'save' | 'save-as' | 'save-copy' | 'export' | 'fallback'

/** Request for simpleIO.chooseSavePath(). Main picks the folder; the renderer never does. */
export interface ChooseSavePathRequest {
  purpose: SavePurpose
  docId?: string
  /** Suggested file name, usually "<stem>.<ext>". */
  name: string
  /** Registry id of the format to preselect. */
  format: string
  /** Registry ids of every format the dialog may offer, the preselected one first. */
  formats: string[]
  /** With purpose "fallback": the failure that led here (LOCKED, READ_ONLY, …). */
  reason?: IoCode
}

/**
 * One part of a recovery snapshot (design §4.1, §4.3). Part names use
 * letters, digits, ".", "_" and "-" (at most 80 characters) and are unique.
 */
export interface RecoveryPart {
  /** Logical part name, e.g. "state.json", "workbook.json", "document.docx", "canvas.png". */
  name: string
  /** Text or bytes. Omit when `ref` is given. */
  data?: string | Uint8Array
  /**
   * Instead of bytes: "source" means this part equals the document's own file
   * as main read it (main records the path and stamp; the renderer never names a path).
   */
  ref?: 'source'
  /** With `ref`: the entry can't be restored once that file changed or is gone. */
  required?: boolean
  /** Ask main to compress the part before storing it. */
  compress?: boolean
}

/**
 * What a session sends to simpleIO.recovery.write(). Main adds the source
 * path and stamp from its document registry, the process and the times.
 */
export interface RecoverySnapshot {
  /** The document's registry id: letters, digits, "_" and "-". */
  docId: string
  module?: ModuleId
  title: string
  kind?: string
  format?: string | null
  /** Session revision the parts belong to; main drops writes at or below a discarded revision. */
  revision: number
  parts: RecoveryPart[]
  /** Small JSON (at most 256 KB), such as the active sheet or the selection. */
  extra?: unknown
}

/** What simpleIO.recovery.write() resolves to. A dropped write (stale or discarded) is not a failure. */
export type RecoveryWriteResult =
  | { ok: true; docId?: string; generation?: number; revision?: number }
  | { ok: false; docId?: string; dropped?: 'discarded' | 'stale'; code?: string; message?: string }

/** One recovery journal entry as listed by simpleIO.recovery.list(). */
export interface RecoveryEntry {
  id: string
  docId?: string
  module?: string | null
  title: string | null
  kind?: string | null
  format?: string | null
  sourcePath?: string | null
  suggestedPath?: string | null
  /** "same", "changed", "missing" or "unknown": the source file compared with the snapshot's stamp. */
  sourceState?: string | null
  revision?: number
  generation?: number
  createdAt?: string | number | null
  updatedAt?: string | number | null
  offeredAt?: string | null
  /** False for entries that belong to a document open right now. */
  orphaned?: boolean
  state?: string
  /** The source file changed after the snapshot; saving will ask before replacing it. */
  sourceChanged?: boolean
  /** The source file is not there anymore. */
  sourceMissing?: boolean
  /** False when a required file the entry refers to changed or is gone, so it can't be restored. */
  restorable?: boolean
  /** Why it can't be restored: "base-changed" or "base-missing". */
  reason?: string | null
  size?: number
}

/** A stored part as read back by simpleIO.recovery.read(). */
export interface RecoveredPart {
  name: string
  data?: string | Uint8Array
  ref?: { path: string; stamp?: Stamp | null; state?: string | null; matches?: boolean }
  required?: boolean
}

/**
 * Result of simpleIO.recovery.read(). `parts` is keyed by part name; a list
 * of parts is accepted too (see recoveredParts()).
 */
export type RecoveryPayload =
  | {
    ok?: true
    docId?: string
    entry: RecoveryEntry
    revision?: number
    generation?: number
    /** The newest copy was damaged, so the one before it was read. */
    fellBack?: boolean
    parts: Record<string, Uint8Array | string> | RecoveredPart[]
    extra?: unknown
  }
  | { ok: false; code: string; message?: string; reason?: string }

/**
 * What recovery.adopt() resolves to: the window took the entry over (same
 * document id; bound to the entry's file, or untitled when the file could not
 * be proven unchanged), or why it could not.
 */
export type RecoveryAdoptResult =
  | { ok: true; docId: string; untitled?: boolean; path?: string | null }
  | { ok: false; code: string; message?: string; reason?: string; technical?: string }

/** What recovery.discard() resolves to. */
export type RecoveryDiscardResult = { ok: boolean; removed?: boolean; kept?: boolean; code?: string; message?: string }

/** One saved earlier version of a file (design §3.8). */
export interface VersionEntry {
  id: string
  path: string
  name: string
  createdAt: string | number
  size: number
  sourcePath?: string | null
}

/** Result of simpleIO.clipboard.read() (design §6.6). Every part is capped by main. */
export interface ClipboardPayload {
  /** Real paths of copied files. */
  files: string[]
  png?: Uint8Array
  html?: string
  rtf?: string
  text?: string
}

/** Requests main sends to a renderer (design §3.7). */
export type RequestType = 'close-query' | 'save-now' | 'discard' | 'recovery-flush'

/** Toast kinds. Failures are never toasts; they get a decision prompt. */
export type NotifyKind = 'success' | 'info' | 'warning'

/** One toast button. */
export interface NotifyAction {
  label: string
  run: () => void
}

/**
 * Result of simpleIO.openInSimple(). `action` says what happened: the window
 * that shows the file was focused, this workspace opened it in a new window,
 * the unified app was asked to open it in its workspace, or (no Simple
 * workspace opens this format) File Explorer showed it.
 */
export interface OpenInSimpleResult {
  ok: boolean
  action?: 'focused' | 'opened' | 'launched' | 'shown-in-folder'
  /** Same as action "shown-in-folder". */
  shownInFolder?: boolean
  /** Workspace that received the file, when known. */
  mode?: string | null
  appName?: string | null
  path?: string
  reason?: string
  code?: string
  message?: string
}

/**
 * window.simpleIO, identical in every workspace (design §8.1). Two members of
 * the design are left out on purpose: `shell.openPath` (documents always open
 * inside Simple, through openInSimple) and `officeEngine.requestInstall`
 * (Simple never downloads or installs anything).
 */
export interface SimpleIO {
  readonly version: 1
  /** The workspace, or null when main did not name it. */
  readonly module: ModuleId | null
  capabilities(): Promise<Capabilities>
  onCapabilitiesChanged(callback: () => void): () => void
  /** The real path of a dropped or pasted File, or null when it has none (browser or mail drags). */
  pathForFile(file: File): string | null
  chooseSavePath(request: ChooseSavePathRequest): Promise<{ path: string; format: string } | null>
  chooseOpenPaths(request: { multi: boolean; purpose: 'open' | 'import' | 'insert' }): Promise<string[]>
  /** Shows a native decision prompt from the catalog and returns the chosen button id. */
  prompt(key: string, vars: Record<string, string>): Promise<string>
  recovery: {
    write(snapshot: RecoverySnapshot): Promise<RecoveryWriteResult | void>
    list(): Promise<RecoveryEntry[]>
    /** Reads an entry. Changes nothing: call adopt() once the document was rebuilt from it. */
    read(id: string): Promise<RecoveryPayload>
    /**
     * Takes a read entry over for this window after its content was restored:
     * main binds the document id to the entry's file and stops offering it.
     * Optional only for older bridges.
     */
    adopt?(id: string): Promise<RecoveryAdoptResult>
    discard(id: string, upToRevision?: number): Promise<RecoveryDiscardResult | void>
  }
  versions: {
    list(path: string): Promise<VersionEntry[]>
    open(id: string): Promise<unknown>
  }
  clipboard: { read(): Promise<ClipboardPayload> }
  shell: { showItem(path: string): Promise<unknown> }
  prefs: {
    get<T>(key: string): Promise<T | undefined>
    /** Resolves false when the preference could not be written; the key must be letters, digits, ".", "_" or "-". */
    set(key: string, value: unknown): Promise<boolean | void>
  }
  officeEngine: { status(): Promise<OfficeEngineStatus> }
  onRequest(type: RequestType, handler: (payload: unknown) => unknown): () => void
  /** Opens a file in the Simple workspace that owns its format. Never hands it to another program. */
  openInSimple(path: string): Promise<OpenInSimpleResult | void>
}

// ---------------------------------------------------------------------------
// The bridge
// ---------------------------------------------------------------------------

function isSimpleIO(value: unknown): value is SimpleIO {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<SimpleIO>
  return candidate.version === 1 && typeof candidate.prompt === 'function' && typeof candidate.chooseSavePath === 'function'
}

/**
 * Returns window.simpleIO, or null outside Simple (a plain browser, a unit test).
 * @returns The typed bridge, or null when the preload did not expose it.
 */
export function getSimpleIO(): SimpleIO | null {
  const candidate = (globalThis as { simpleIO?: unknown }).simpleIO
  return isSimpleIO(candidate) ? candidate : null
}

/**
 * Returns window.simpleIO or throws a clear error when the preload block is missing.
 * @throws {Error} when the bridge is not available
 */
export function requireSimpleIO(): SimpleIO {
  const io = getSimpleIO()
  if (!io) throw new Error('window.simpleIO is not available. The workspace preload must include the simple-io bridge block.')
  return io
}

/**
 * Checks whether a value is one of the shared I/O result codes.
 * @param value anything, usually a `code` that crossed IPC
 */
export function isIoCode(value: unknown): value is IoCode {
  return typeof value === 'string' && (IO_CODES as readonly string[]).includes(value)
}

/**
 * Turns anything an adapter or IPC call threw into an IoFailure, so callers
 * only ever handle IoResult values.
 * @param error what was thrown
 * @param fallback code to use when the error carries no known code
 */
export function failureFromError(error: unknown, fallback: IoCode = 'UNKNOWN'): IoFailure {
  if (error && typeof error === 'object') {
    const value = error as { ok?: unknown; code?: unknown; message?: unknown; name?: unknown }
    if (value.ok === false && isIoCode(value.code)) return error as IoFailure
    if (value.name === 'AbortError') return { ok: false, code: 'CANCELED' }
    const code = isIoCode(value.code) ? value.code : fallback
    const message = typeof value.message === 'string' ? value.message : ''
    const failure: IoFailure = { ok: false, code }
    if (message) failure.technical = message.length > 300 ? `${message.slice(0, 297)}…` : message
    return failure
  }
  return { ok: false, code: fallback, technical: String(error) }
}

/**
 * Normalizes what a save or export handler returned. A missing or malformed
 * result is a failure, never a silent success.
 * @param value the handler's return value
 */
export function toIoResult(value: unknown): IoResult {
  if (value && typeof value === 'object') {
    const result = value as { ok?: unknown; path?: unknown; code?: unknown }
    if (result.ok === true && typeof result.path === 'string' && result.path) return value as IoSuccess
    if (result.ok === false) return isIoCode(result.code) ? value as IoFailure : { ...(value as IoFailure), code: 'UNKNOWN' }
  }
  return { ok: false, code: 'UNKNOWN', technical: 'The save handler returned no result.' }
}

// ---------------------------------------------------------------------------
// Catalog
// ---------------------------------------------------------------------------

type CatalogValue = string | number | boolean | null | readonly CatalogValue[] | CatalogRecord
interface CatalogRecord { readonly [key: string]: CatalogValue }

const catalog = catalogData as unknown as CatalogRecord

function isRecord(value: CatalogValue | undefined): value is CatalogRecord {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value)
}

/** Values a catalog template can take. Missing values fall back to the catalog's neutral wording. */
export type TemplateVars = { readonly [key: string]: string | number | boolean | null | undefined }

/**
 * Looks up a raw catalog value by dotted key, such as "status.saving".
 * @param key dotted path into io-catalog.json
 * @returns the value, or undefined when the key does not exist
 */
export function catalogValue(key: string): CatalogValue | undefined {
  let node: CatalogValue | undefined = catalog
  for (const part of key.split('.')) {
    if (!isRecord(node) || !Object.prototype.hasOwnProperty.call(node, part)) return undefined
    node = node[part]
  }
  return node
}

/**
 * Replaces {placeholders} in a catalog string. A missing value falls back to
 * the catalog's neutral wording ("this file"), never to an empty gap. Same
 * rules as formatTemplate() in io-core.cjs.
 * @param template text with {name} placeholders
 * @param vars placeholder values
 */
export function formatTemplate(template: string, vars: TemplateVars = {}): string {
  return template.replace(/\{(\w+)\}/g, (match: string, key: string) => {
    const value = vars[key]
    if (value !== undefined && value !== null && value !== '') return String(value)
    const fallback = catalogValue(`placeholders.${key}`)
    return typeof fallback === 'string' ? fallback : match
  })
}

/**
 * Formats a catalog string by dotted key.
 * @param key for example "toast.saved" or "status.savedAs"
 * @param vars placeholder values
 * @returns the text, or "" when the key is missing or not a string
 */
export function catalogText(key: string, vars: TemplateVars = {}): string {
  const value = catalogValue(key)
  return typeof value === 'string' ? formatTemplate(value, vars) : ''
}

/** A decision prompt from the catalog (the same data main shows in a native message box). */
export interface CatalogPrompt {
  message: string
  detail: string
  buttons: Array<{ id: string; label: string }>
  defaultId: string | null
  cancelId: string | null
}

/**
 * Formats a catalog prompt such as "saveFailed.LOCKED" or "prompts.unsaved".
 * Decisions are shown by main as native prompts (simpleIO.prompt); this is
 * for in-app text such as the status chip's tooltip.
 * @param key dotted catalog key
 * @param vars placeholder values
 * @returns the formatted prompt, or null when the key does not name a prompt
 */
export function catalogPrompt(key: string, vars: TemplateVars = {}): CatalogPrompt | null {
  const entry = catalogValue(key)
  if (!isRecord(entry) || typeof entry.message !== 'string') return null
  const buttons: Array<{ id: string; label: string }> = []
  const list = entry.buttons
  if (Array.isArray(list)) {
    for (const button of list as readonly CatalogValue[]) {
      if (isRecord(button) && typeof button.id === 'string' && typeof button.label === 'string') {
        buttons.push({ id: button.id, label: formatTemplate(button.label, vars) })
      }
    }
  }
  return {
    message: formatTemplate(entry.message, vars),
    detail: typeof entry.detail === 'string' ? formatTemplate(entry.detail, vars) : '',
    buttons,
    defaultId: typeof entry.defaultId === 'string' ? entry.defaultId : null,
    cancelId: typeof entry.cancelId === 'string' ? entry.cancelId : null,
  }
}

/**
 * The user-facing workspace name, e.g. "Simple Spreadsheets".
 * @param module workspace id ("launcher" for the launcher window)
 */
export function appName(module: ModuleId | 'launcher' | string): string {
  return catalogText(`appNames.${module}`) || catalogText('placeholders.appName') || 'Simple'
}

/**
 * The noun for one document of a workspace, e.g. "spreadsheet".
 * @param module workspace id
 */
export function kindName(module: ModuleId | string): string {
  return catalogText(`kinds.${module}`)
}

/** Keys of the status chip texts in io-catalog.json (status.*). */
export type StatusKey =
  | 'saving'
  | 'saved'
  | 'unsaved'
  | 'notSaved'
  | 'savedAs'
  | 'finishingSave'
  | 'willSaveAgain'
  | 'waitingForExport'
  | 'waitingForPrint'
  | 'recovered'
  | 'readOnly'
  | 'autosaveOff'
  | 'autosaveFailing'
  | 'exporting'
  | 'exportingPercent'

/**
 * The status chip text for a key.
 * @param key a status.* catalog key, or null for no chip
 * @param vars placeholder values (formatLabel, percent)
 */
export function statusText(key: StatusKey | null, vars: TemplateVars = {}): string {
  return key ? catalogText(`status.${key}`, vars) : ''
}

const FAILURE_VAR_KEYS = [
  'name', 'path', 'folder', 'nearestFolder', 'needed', 'drive', 'reason', 'technical', 'backupPath',
  'formatLabel', 'altLabel', 'altExt', 'altFormat', 'ext', 'limit',
] as const

/**
 * Placeholder values for a failure prompt: every string detail of the failure.
 * @param failure the failed result
 * @returns string values only, ready for simpleIO.prompt()
 */
export function failureVars(failure: IoFailure): Record<string, string> {
  const vars: Record<string, string> = { code: failure.code }
  for (const key of FAILURE_VAR_KEYS) {
    const value = failure[key]
    if (typeof value === 'string' && value) vars[key] = value
  }
  if (!vars.backupPath && failure.asidePath) vars.backupPath = failure.asidePath
  if (!vars.name && failure.path) vars.name = baseName(failure.path)
  if (!vars.folder && failure.path) vars.folder = folderOf(failure.path)
  return vars
}

/**
 * Plain text for showing a failure inside the app (open errors are shown in an
 * error bar, not a dialog; design §6.7).
 * @param failure the failed result
 * @returns message, detail and the technical "Details:" line ("" when there is none)
 */
export function describeFailure(failure: IoFailure): { message: string; detail: string; details: string } {
  const table = failure.context === 'open' ? 'openFailed' : 'saveFailed'
  const vars = failureVars(failure)
  const key = isRecord(catalogValue(`${table}.${failure.code}`)) ? `${table}.${failure.code}` : `${table}.UNKNOWN`
  const prompt = catalogPrompt(key, vars)
  const message = prompt?.message || failure.message || catalogText(`notices.${failure.code}`, vars)
  return {
    message,
    detail: prompt?.detail || '',
    details: failure.technical ? catalogText('detailsLine', { technical: failure.technical }) : '',
  }
}

/**
 * Shows the native decision prompt for a failed save or export (catalog
 * saveFailed.<CODE>) and returns the chosen button id.
 * @param io the bridge, or null outside Simple
 * @param failure the failed result
 * @returns the button id, or "cancel" when no prompt could be shown
 */
export async function promptFailure(io: SimpleIO | null, failure: IoFailure): Promise<string> {
  if (!io) return 'cancel'
  let code: string = isRecord(catalogValue(`saveFailed.${failure.code}`)) ? failure.code : 'UNKNOWN'
  // Another Simple window editing the file is not "another program": it has its own wording.
  if (failure.reason === 'open-in-another-window' && isRecord(catalogValue('saveFailed.OPEN_IN_ANOTHER_WINDOW'))) code = 'OPEN_IN_ANOTHER_WINDOW'
  try {
    const answer = await io.prompt(`saveFailed.${code}`, failureVars(failure))
    return typeof answer === 'string' && answer ? answer : 'cancel'
  } catch {
    return 'cancel'
  }
}

/**
 * Opens a file in the Simple workspace that owns its format (simpleIO.openInSimple).
 * When the bridge can't do that (an older preload, a failed hand-off) the file
 * is only shown in File Explorer; it is never handed to another program.
 * @param io the bridge, or null outside Simple
 * @param path the file to open
 * @returns what happened; see openedInSimple()
 */
export async function openInSimple(io: SimpleIO | null, path: string): Promise<OpenInSimpleResult> {
  if (!io || !path) return { ok: false, code: 'UNSUPPORTED' }
  if (typeof io.openInSimple === 'function') {
    try {
      const result = await io.openInSimple(path)
      if (result && typeof result === 'object') return result
      return { ok: true, path }
    } catch {
      // Fall through: show the file instead.
    }
  }
  try {
    await io.shell.showItem(path)
    return { ok: true, action: 'shown-in-folder', shownInFolder: true, path, reason: 'no-bridge' }
  } catch {
    return { ok: false, code: 'UNKNOWN', path }
  }
}

/**
 * True when openInSimple() put the file in front of the user inside Simple
 * (not just in File Explorer).
 * @param result what openInSimple() returned
 */
export function openedInSimple(result: OpenInSimpleResult | null | undefined): boolean {
  return Boolean(result && result.ok !== false && result.action !== 'shown-in-folder' && result.shownInFolder !== true)
}

/**
 * The parts of a recovery payload as a list, whichever shape main sent
 * (keyed by name, or a list of parts).
 * @param payload what simpleIO.recovery.read() returned
 */
export function recoveredParts(payload: RecoveryPayload | null | undefined): RecoveredPart[] {
  if (!payload || payload.ok === false || !payload.parts) return []
  const parts = payload.parts
  if (Array.isArray(parts)) return parts.filter((part) => part && typeof part.name === 'string')
  return Object.entries(parts).map(([name, data]) => ({ name, data }))
}

/**
 * The contents of one recovered part, as text (bytes are read as UTF-8) or as
 * bytes, or null when the payload has no such part.
 * @param payload what simpleIO.recovery.read() returned
 * @param name the part name used when the snapshot was written
 * @param as "text" (default) or "bytes"
 */
export function recoveredPart(payload: RecoveryPayload | null | undefined, name: string, as?: 'text'): string | null
export function recoveredPart(payload: RecoveryPayload | null | undefined, name: string, as: 'bytes'): Uint8Array | null
export function recoveredPart(payload: RecoveryPayload | null | undefined, name: string, as: 'text' | 'bytes' = 'text'): string | Uint8Array | null {
  const part = recoveredParts(payload).find((candidate) => candidate.name === name)
  if (!part || part.data === undefined || part.data === null) return null
  if (as === 'bytes') return typeof part.data === 'string' ? new TextEncoder().encode(part.data) : part.data
  return typeof part.data === 'string' ? part.data : new TextDecoder('utf-8').decode(part.data)
}

// ---------------------------------------------------------------------------
// Small text and path helpers
// ---------------------------------------------------------------------------

/**
 * Joins items as "a, b and c".
 * @param items words or short phrases
 */
export function joinList(items: readonly string[]): string {
  const list = items.filter((item) => typeof item === 'string' && item.trim()).map((item) => item.trim())
  if (list.length <= 1) return list.join('')
  return `${list.slice(0, -1).join(', ')} and ${list[list.length - 1]}`
}

/**
 * Upper-cases the first letter, for list text that starts a sentence.
 * @param text any text
 */
export function capitalize(text: string): string {
  return text ? text.charAt(0).toUpperCase() + text.slice(1) : text
}

/**
 * The last segment of a Windows or POSIX path.
 * @param path a file path
 */
export function baseName(path: string): string {
  const trimmed = String(path || '').replace(/[\\/]+$/, '')
  const index = Math.max(trimmed.lastIndexOf('\\'), trimmed.lastIndexOf('/'))
  return index >= 0 ? trimmed.slice(index + 1) : trimmed
}

/**
 * The folder part of a path ("" when the path has no folder).
 * @param path a file path
 */
export function folderOf(path: string): string {
  const value = String(path || '')
  const index = Math.max(value.lastIndexOf('\\'), value.lastIndexOf('/'))
  if (index < 0) return ''
  if (index === 2 && /^[a-z]:/i.test(value)) return value.slice(0, 3)
  if (index === 0) return value.slice(0, 1)
  return value.slice(0, index)
}

/**
 * A short folder label for messages: "Documents › Finance" (the last two
 * folder names), or the drive for a root folder.
 * @param folder a folder path
 */
export function folderLabel(folder: string): string {
  const parts = String(folder || '').split(/[\\/]+/).filter(Boolean)
  if (!parts.length) return ''
  return parts.slice(-2).join(' › ')
}

/**
 * The lower-case extension of a file name or path, with its dot ("" when none).
 * @param name a file name or path
 */
export function fileExtension(name: string): string {
  const base = baseName(name)
  const index = base.lastIndexOf('.')
  if (index <= 0 || index === base.length - 1) return ''
  return base.slice(index).toLowerCase()
}

/**
 * Compares two Windows paths the way the file system does: case and slash
 * direction do not matter.
 * @param a a path
 * @param b another path
 */
export function samePath(a: string | null | undefined, b: string | null | undefined): boolean {
  if (!a || !b) return false
  const normalize = (value: string) => value.replace(/\//g, '\\').replace(/\\+$/, '').toLowerCase()
  return normalize(a) === normalize(b)
}

// ---------------------------------------------------------------------------
// Drop and paste (design §6.5, §6.6)
// ---------------------------------------------------------------------------

/** A dropped or pasted file, with its real path when it has one. */
export interface PickedFile {
  /** The File object; null for a path that came from the system clipboard. */
  file: File | null
  name: string
  /** Real path on disk; null for files dragged from a browser or mail program. */
  path: string | null
  /** Lower-case extension with its dot, "" when none. */
  extension: string
  /** Size in bytes, when known. */
  size: number | null
  type: string
}

function pathFor(io: SimpleIO | null | undefined, file: File): string | null {
  if (!io || typeof io.pathForFile !== 'function') return null
  try {
    const path = io.pathForFile(file)
    return typeof path === 'string' && path ? path : null
  } catch {
    return null
  }
}

function pickedFromFile(file: File, io: SimpleIO | null | undefined): PickedFile {
  return {
    file,
    name: file.name,
    path: pathFor(io, file),
    extension: fileExtension(file.name),
    size: typeof file.size === 'number' ? file.size : null,
    type: file.type || '',
  }
}

function pickedFromPath(path: string): PickedFile {
  const name = baseName(path)
  return { file: null, name, path, extension: fileExtension(name), size: null, type: '' }
}

/**
 * True when a drag carries files, for dragenter and dragover handlers.
 * @param transfer the event's dataTransfer
 */
export function dragHasFiles(transfer: DataTransfer | null | undefined): boolean {
  if (!transfer) return false
  const types = Array.from(transfer.types || [])
  return types.includes('Files') || (transfer.files ? transfer.files.length > 0 : false)
}

/**
 * Lists the files of a drop (or of a paste event's clipboardData) with their
 * real paths. A file with a path must be opened through its path, so Save
 * writes in place, Recent works and conflicts are detected.
 * @param transfer the event's dataTransfer or clipboardData
 * @param io the bridge, used for pathForFile
 */
export function filesFromTransfer(transfer: DataTransfer | null | undefined, io: SimpleIO | null | undefined = getSimpleIO()): PickedFile[] {
  if (!transfer || !transfer.files) return []
  return Array.from(transfer.files).map((file) => pickedFromFile(file, io))
}

/**
 * Splits picked files into those a workspace opens and the rest.
 * @param files picked files
 * @param accepted extensions with dots (".pdf"), or a predicate
 */
export function partitionFiles(
  files: readonly PickedFile[],
  accepted: readonly string[] | ((file: PickedFile) => boolean),
): { accepted: PickedFile[]; rejected: PickedFile[] } {
  const test = typeof accepted === 'function'
    ? accepted
    : (file: PickedFile) => accepted.map((extension) => extension.toLowerCase()).includes(file.extension)
  const result: { accepted: PickedFile[]; rejected: PickedFile[] } = { accepted: [], rejected: [] }
  for (const file of files) (test(file) ? result.accepted : result.rejected).push(file)
  return result
}

/**
 * The one summary toast for files that were not opened, or null when there are none.
 * @param names file names
 */
export function notOpenedMessage(names: readonly string[]): string | null {
  const list = names.filter(Boolean)
  if (!list.length) return null
  if (list.length === 1) return catalogText('toast.notOpenedOne', { name: list[0] })
  return catalogText('toast.notOpened', { count: list.length, list: list.join(', ') })
}

/** Where an Open, drop, paste or Recent pick goes (design §3.7 item 6). */
export type OpenPlacement = 'this-window' | 'new-window'

/**
 * Open never replaces a document with unsaved changes: it uses this window
 * only on its welcome screen with nothing unsaved, otherwise a new window.
 * @param state whether the current document is dirty and the window shows its welcome screen
 */
export function openPlacement(state: { dirty: boolean; onWelcomeScreen: boolean }): OpenPlacement {
  return state.onWelcomeScreen && !state.dirty ? 'this-window' : 'new-window'
}

/** What a drop does (design §6.5). The workspace labels its overlay accordingly. */
export type DropAction = 'open' | 'open-new-window' | 'insert'

/**
 * Decides what a drop does. Dropped on content that accepts inserts (PDF page
 * strip, Calc grid, Docs body, Image album): insert. Anywhere else: open, in
 * a new window unless this window is on its welcome screen with nothing unsaved.
 * @param state drop location and window state
 */
export function dropAction(state: { onInsertTarget: boolean; dirty: boolean; onWelcomeScreen: boolean }): DropAction {
  if (state.onInsertTarget) return 'insert'
  return openPlacement(state) === 'this-window' ? 'open' : 'open-new-window'
}

/**
 * Reads a path-less dropped or pasted file (dragged from a browser or mail
 * program) so it can open as an untitled copy that keeps its name.
 * @param file the File
 * @param maxBytes refuse larger files (default 64 MiB, like the clipboard cap)
 * @returns the bytes, or null when the file is larger than maxBytes
 */
export async function readFileBytes(file: File, maxBytes = 64 * 1024 * 1024): Promise<Uint8Array | null> {
  if (file.size > maxBytes) return null
  return new Uint8Array(await file.arrayBuffer())
}

/** Everything a paste offered, files first (design §6.6). */
export interface PastedContent {
  /** Pasted files; behave exactly like a drop of the same files. */
  files: PickedFile[]
  png: Uint8Array | null
  html: string | null
  rtf: string | null
  text: string | null
}

/**
 * Collects what a paste offers. Call it from the paste handler: the event's
 * clipboardData is read synchronously (it is only readable during the event),
 * then the system clipboard is read through simpleIO.clipboard.read() for
 * copied file paths and images.
 * @param event the paste event, or null for a "Paste" command without an event
 * @param io the bridge
 */
export async function readPaste(
  event: { readonly clipboardData: DataTransfer | null } | null,
  io: SimpleIO | null = getSimpleIO(),
): Promise<PastedContent> {
  const transfer = event ? event.clipboardData : null
  const content: PastedContent = { files: filesFromTransfer(transfer, io), png: null, html: null, rtf: null, text: null }
  if (transfer) {
    const text = transfer.getData('text/plain')
    const html = transfer.getData('text/html')
    const rtf = transfer.getData('text/rtf')
    if (text) content.text = text
    if (html) content.html = html
    if (rtf) content.rtf = rtf
  }
  if (!io) return content
  let system: ClipboardPayload | null = null
  try {
    system = await io.clipboard.read()
  } catch {
    system = null
  }
  if (!system) return content
  if (!content.files.length && Array.isArray(system.files)) {
    content.files = system.files.filter((path) => typeof path === 'string' && path).map(pickedFromPath)
  }
  if (system.png && system.png.byteLength) content.png = system.png
  if (content.html === null && system.html) content.html = system.html
  if (content.rtf === null && system.rtf) content.rtf = system.rtf
  if (content.text === null && system.text) content.text = system.text
  return content
}
