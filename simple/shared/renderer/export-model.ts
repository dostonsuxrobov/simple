// Vendored from simple/shared/renderer/export-model.ts by simple/scripts/sync-shared.cjs. Do not edit here.

// ExportModel: the state of the Export As dialog (design §7). Each workspace
// renders it with its own components; the model holds the rules:
// - a format list with one-line descriptions; a row that can't be produced on
//   this PC stays visible, says why, and can't be selected
// - a few contextual options per format, remembered per format
// - the last used format, folder and "open after exporting" are remembered
// - pending edits are committed first (session.prepare('export'))
// - the native Save dialog in main picks the path; strict extension rules
//   (no "photo.jpg.png"), and typing another offered extension switches format
// - failures get the same decision prompts as Save; nothing fails silently
// - warnings are reported; "Open" and "Open the file after exporting" open the
//   result in the matching Simple workspace through openInSimple(), never in
//   another program
// Export is a copy: it never rebinds the document or changes its dirty state.

import type {
  ChooseSavePathRequest,
  IoCode,
  IoFailure,
  IoResult,
  IoSuccess,
  ModuleId,
  NotifyAction,
  NotifyKind,
  SimpleIO,
} from './io-client'
import {
  baseName,
  catalogText,
  failureFromError,
  fileExtension,
  folderLabel,
  folderOf,
  joinList,
  kindName,
  openInSimple,
  openedInSimple,
  promptFailure,
  toIoResult,
} from './io-client'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/**
 * One Export As choice. Same shape as the rows of exportFormats() in
 * simple/shared/electron/formats.cjs, so a workspace can pass the registry's
 * rows straight through.
 */
export interface ExportFormatRow {
  /** Registry id, e.g. "pdf", "csv". */
  id: string
  label: string
  /** Extensions with dots, the default first. */
  extensions: readonly string[]
  /** One line shown under the label. */
  description: string
  /** "native", "values-only", "basic", … */
  mode?: string
  /** False when this PC can't produce the format now; the row stays visible but can't be selected. */
  available: boolean
  reason?: string | null
  usesEngine?: boolean
  lossy?: { keeps: string; loses: readonly string[] } | null
  /** One file per page or sheet; the exporter creates the folder (see multiFileFolderName). */
  multiFile?: boolean
}

/** Value of one export option. */
export type ExportOptionValue = string | number | boolean

/** How an option is edited. */
export type ExportOptionType = 'choice' | 'toggle' | 'number' | 'text'

/** One choice of a "choice" option. */
export interface ExportOptionChoice {
  id: string
  label: string
  description?: string
}

/** A contextual export option the workspace offers (design §7.3 "Main options"). */
export interface ExportOptionDef {
  id: string
  label: string
  type: ExportOptionType
  defaultValue: ExportOptionValue
  /** Formats the option applies to; omitted means every format. */
  formats?: readonly string[]
  choices?: readonly ExportOptionChoice[]
  min?: number
  max?: number
  step?: number
  description?: string
  /** Show the option only when this returns true (for example a page range only for "Range"). */
  visible?: (values: Readonly<Record<string, ExportOptionValue>>, formatId: string) => boolean
  /** Remember the value per format (default true). */
  remember?: boolean
}

/** What the workspace's exporter receives. It writes through main (safeWriteFile) and returns main's IoResult. */
export interface ExportRequest {
  path: string
  format: string
  /** Values of the options visible for this format. */
  options: Record<string, ExportOptionValue>
  docId?: string
  /** The user chose Replace after the target changed on disk. */
  force?: boolean
  /** The user chose Save Here Again after the target vanished. */
  recreate?: boolean
  /** Aborted when the user cancels; the exporter must leave no partial file and return CANCELED. */
  signal: AbortSignal
  /** Report progress from 0 to 100 for long exports. */
  onProgress(percent: number): void
}

/** The workspace's export function. */
export type Exporter = (request: ExportRequest) => Promise<IoResult>

/** The part of a DocumentSession the export model uses. */
export interface ExportSessionLike {
  readonly docId?: string
  prepare(action: 'export'): Promise<boolean>
  track?<T>(action: 'export', work: () => Promise<T>): Promise<T>
  /** The export replaced the document's own file: the document is marked changed (DocumentSession.noteOwnFileChanged). */
  noteOwnFileChanged?(): void
}

/** Everything the model needs from the workspace. */
export interface ExportModelConfig {
  io: SimpleIO | null
  /** Workspace id; defaults to io.module. Decides the first-run default format. */
  module?: ModuleId
  /** Display name of the document, e.g. "Budget.xlsx". */
  documentName: string
  /** The document's file, when it has one (its folder is the first-run export folder). */
  documentPath?: string | null
  docId?: string
  formats: readonly ExportFormatRow[]
  options?: readonly ExportOptionDef[]
  /** Format ids to list first, in this order (for example the document's own format). */
  order?: readonly string[]
  /** Preselected format when nothing is remembered; defaults to PDF (PNG in Images). */
  defaultFormat?: string
  exporter: Exporter
  session?: ExportSessionLike | null
  notify?: (kind: NotifyKind, text: string, actions?: NotifyAction[]) => void
  /** Shows the list of items an export left out (the toast's Details button). */
  onShowWarnings?: (warnings: readonly string[], name: string) => void
  /** Extensions that count as "known" when a typed name is checked (in addition to every row's). */
  knownExtensions?: readonly string[]
  log?: (message: string, detail?: unknown) => void
}

/** One row as rendered. */
export interface ExportRowView {
  readonly id: string
  readonly label: string
  /** Default extension, e.g. ".xlsx". */
  readonly extension: string
  readonly extensions: readonly string[]
  readonly description: string
  /** Short note next to the name: "values only", "basic formatting" or "Not available on this PC". */
  readonly tag: string
  readonly selectable: boolean
  readonly selected: boolean
  readonly lossy: boolean
  readonly multiFile: boolean
}

/** One option as rendered. */
export interface ExportOptionView {
  readonly id: string
  readonly label: string
  readonly type: ExportOptionType
  readonly value: ExportOptionValue
  readonly description: string
  readonly choices: readonly ExportOptionChoice[]
  readonly min: number | null
  readonly max: number | null
  readonly step: number | null
}

/** The dialog state. The object is replaced on every change. */
export interface ExportModelState {
  /** "Export "Budget.xlsx"". */
  readonly title: string
  readonly rows: readonly ExportRowView[]
  readonly selected: ExportRowView | null
  readonly options: readonly ExportOptionView[]
  /** "Keeps: values. Not kept: formulas, formatting and charts." for lossy formats, else "". */
  readonly lossNote: string
  readonly openAfter: boolean
  readonly openAfterLabel: string
  /** Remembered export folder, or the document's folder, or null (main then uses Documents). */
  readonly folder: string | null
  /** "Saves to: Documents › Finance", or "" when the folder is not known yet. */
  readonly savesTo: string
  /** Suggested file name, "<stem>.<ext>". */
  readonly fileName: string
  readonly busy: boolean
  readonly progress: number | null
  readonly statusText: string
  readonly canExport: boolean
  /** Items the last export left out. */
  readonly warnings: readonly string[]
}

/** Result of run(). */
export type ExportOutcome =
  | { ok: true; path: string; name: string; format: string; warnings: string[]; opened: boolean }
  | { ok: false; code: IoCode; failure?: IoFailure }

/** io-prefs.json keys the model reads and writes (main stores them per workspace). */
export const EXPORT_PREF_KEYS = Object.freeze({
  format: 'exportFormat',
  folder: 'lastExportFolder',
  openAfter: 'openAfterExport',
  options: 'exportOptions',
})

/** First-run default format per workspace (design §7.2 rule 1). */
export const DEFAULT_EXPORT_FORMATS: Readonly<Record<ModuleId, string | null>> = Object.freeze({
  pdf: 'pdf',
  calc: 'pdf',
  docs: 'pdf',
  image: 'png',
  video: null,
})

/** Extensions that are never kept as part of a file name's stem. */
export const COMMON_EXTENSIONS: readonly string[] = Object.freeze([
  '.pdf', '.docx', '.docm', '.dotx', '.dotm', '.doc', '.rtf', '.odt', '.txt', '.md', '.markdown', '.html', '.htm',
  '.xml', '.json', '.csv', '.tsv', '.tab', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xls', '.xlt', '.ods',
  '.fods', '.numbers', '.pptx', '.ppt', '.odp', '.epub', '.zip', '.png', '.apng', '.jpg', '.jpeg', '.jfif', '.jpe',
  '.jif', '.webp', '.gif', '.bmp', '.tif', '.tiff', '.svg', '.svgz', '.avif', '.ico', '.heic', '.heif', '.psd',
  '.mp4', '.m4v', '.mov', '.webm', '.ogv', '.mkv',
])

// ---------------------------------------------------------------------------
// Pure helpers
// ---------------------------------------------------------------------------

const MAX_EXPORT_ROUNDS = 25

function lower(values: readonly string[]): string[] {
  return values.map((value) => String(value).toLowerCase())
}

/**
 * The first-run default format of a workspace.
 * @param module workspace id
 */
export function defaultExportFormat(module: ModuleId | null | undefined): string | null {
  return module ? DEFAULT_EXPORT_FORMATS[module] ?? null : null
}

/**
 * The part of a document name that names an exported file: a known extension
 * is removed ("Budget.xlsx" → "Budget"), anything else is kept ("Q3 plan
 * v2.1" stays whole). Characters Windows forbids in names become "-".
 * @param name document name or path
 * @param knownExtensions extensions with dots that may be removed
 */
export function safeStem(name: string, knownExtensions: readonly string[] = COMMON_EXTENSIONS): string {
  let stem = baseName(String(name || '')).trim()
  const known = lower(knownExtensions)
  const extension = fileExtension(stem)
  if (extension && known.includes(extension)) stem = stem.slice(0, -extension.length)
  else if (known.includes(stem.toLowerCase())) stem = ''
  return stem
    .replace(/[<>:"/\\|?*\u0000-\u001f]+$/, '')
    .replace(/[<>:"/\\|?*\u0000-\u001f]+/g, '-')
    .replace(/[. ]+$/, '')
    .trim()
}

/**
 * The suggested file name for an export: "<stem>.<ext>".
 * @param documentName document name
 * @param row the chosen format
 * @param fallbackStem stem to use when the name has none
 * @param knownExtensions extensions that may be removed from the name
 */
export function exportFileName(
  documentName: string,
  row: Pick<ExportFormatRow, 'extensions'>,
  fallbackStem = 'export',
  knownExtensions: readonly string[] = COMMON_EXTENSIONS,
): string {
  const stem = safeStem(documentName, knownExtensions) || fallbackStem
  return `${stem}${row.extensions[0] ?? ''}`
}

/**
 * The folder name for a multi-file export: "<stem> - <format> pages".
 * @param documentName document name
 * @param row the chosen format
 * @param knownExtensions extensions that may be removed from the name
 */
export function multiFileFolderName(
  documentName: string,
  row: Pick<ExportFormatRow, 'label'>,
  knownExtensions: readonly string[] = COMMON_EXTENSIONS,
): string {
  return catalogText('export.multiFileFolder', { stem: safeStem(documentName, knownExtensions) || 'export', format: row.label })
}

/**
 * True when a path ends with one of the format's extensions.
 * @param path chosen path
 * @param row the format
 */
export function extensionMatches(path: string, row: Pick<ExportFormatRow, 'extensions'>): boolean {
  const extension = fileExtension(path)
  return Boolean(extension) && lower(row.extensions).includes(extension)
}

/** How a typed name was resolved (design §3.4 "Extension rules"). */
export interface TypedNameResolution {
  name: string
  format: string
  /** keep: already right; switch: another offered format; replace: a known extension replaced; append: extension added. */
  action: 'keep' | 'switch' | 'replace' | 'append'
}

/**
 * Applies the strict extension rules to a typed file name. Writable here →
 * switch to that format. Known but not offered → replace it, never append
 * ("photo.jpg" → "photo.png", not "photo.jpg.png"). Unknown → append ("Q3 plan
 * v2.1" → "Q3 plan v2.1.pdf"). Main applies the same rules to the native
 * dialog; this is for names typed inside the app.
 * @param typed the typed name
 * @param selected the selected format
 * @param rows every row of the dialog
 * @param knownExtensions extensions that count as known besides the rows'
 */
export function resolveTypedName(
  typed: string,
  selected: Pick<ExportFormatRow, 'id' | 'extensions'>,
  rows: readonly Pick<ExportFormatRow, 'id' | 'extensions' | 'available'>[],
  knownExtensions: readonly string[] = COMMON_EXTENSIONS,
): TypedNameResolution {
  const name = String(typed || '').trim()
  const extension = fileExtension(name)
  const primary = selected.extensions[0] ?? ''
  if (extension && lower(selected.extensions).includes(extension)) return { name, format: selected.id, action: 'keep' }
  if (extension) {
    const other = rows.find((row) => row.available && lower(row.extensions).includes(extension))
    if (other) return { name, format: other.id, action: 'switch' }
    const known = lower(knownExtensions).includes(extension) || rows.some((row) => lower(row.extensions).includes(extension))
    if (known) return { name: `${name.slice(0, -extension.length)}${primary}`, format: selected.id, action: 'replace' }
  }
  return { name: `${name}${primary}`, format: selected.id, action: 'append' }
}

/**
 * The short note shown next to a format's name.
 * @param row the format
 */
export function formatTag(row: Pick<ExportFormatRow, 'available' | 'mode'>): string {
  if (!row.available) return catalogText('export.notAvailable')
  if (row.mode === 'values-only') return catalogText('export.valuesOnly')
  if (row.mode === 'basic') return catalogText('export.basicFormatting')
  return ''
}

/**
 * "Keeps: … Not kept: …" for a lossy format, or "".
 * @param row the format
 */
export function lossNote(row: Pick<ExportFormatRow, 'lossy'>): string {
  const lossy = row.lossy
  if (!lossy || !Array.isArray(lossy.loses) || !lossy.loses.length) return ''
  return catalogText('export.keepsAndLoses', { keeps: lossy.keeps, lost: joinList(lossy.loses) })
}

function decimals(step: number): number {
  const text = String(step)
  const dot = text.indexOf('.')
  return dot < 0 ? 0 : text.length - dot - 1
}

/**
 * Checks a value against an option's type, choices and range.
 * @param def the option
 * @param value a value from the UI or from io-prefs.json
 * @returns the accepted value (numbers are clamped and snapped to the step), or undefined
 */
export function coerceOptionValue(def: ExportOptionDef, value: unknown): ExportOptionValue | undefined {
  switch (def.type) {
    case 'toggle':
      return typeof value === 'boolean' ? value : undefined
    case 'choice':
      return typeof value === 'string' && (def.choices || []).some((choice) => choice.id === value) ? value : undefined
    case 'text':
      return typeof value === 'string' ? value : undefined
    case 'number': {
      if (typeof value !== 'number' || !Number.isFinite(value)) return undefined
      let number = value
      if (typeof def.min === 'number') number = Math.max(def.min, number)
      if (typeof def.max === 'number') number = Math.min(def.max, number)
      if (typeof def.step === 'number' && def.step > 0) {
        const origin = typeof def.min === 'number' ? def.min : 0
        number = Number((Math.round((number - origin) / def.step) * def.step + origin).toFixed(decimals(def.step)))
        if (typeof def.max === 'number' && number > def.max) number -= def.step
      }
      return number
    }
    default:
      return undefined
  }
}

function sortRows(rows: readonly ExportFormatRow[], order: readonly string[] | undefined): ExportFormatRow[] {
  const valid = rows.filter((row) => row && typeof row.id === 'string' && row.id && Array.isArray(row.extensions) && row.extensions.length)
  if (!order || !order.length) return [...valid]
  const rank = (row: ExportFormatRow) => {
    const index = order.indexOf(row.id)
    return index < 0 ? order.length : index
  }
  return valid
    .map((row, index) => ({ row, index }))
    .sort((a, b) => rank(a.row) - rank(b.row) || a.index - b.index)
    .map((item) => item.row)
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value)
}

// ---------------------------------------------------------------------------
// ExportModel
// ---------------------------------------------------------------------------

/**
 * State and actions of the Export As dialog (design §7). Create it with
 * ExportModel.create(config), render `state`, and call select(), setOption(),
 * setOpenAfter(), run() and cancel() from the dialog's controls.
 */
export class ExportModel {
  readonly io: SimpleIO | null
  readonly module: ModuleId | null
  readonly documentName: string
  readonly documentPath: string | null

  readonly #config: ExportModelConfig
  readonly #rows: readonly ExportFormatRow[]
  readonly #defs: readonly ExportOptionDef[]
  readonly #known: readonly string[]
  #selectedId: string | null = null
  #values = new Map<string, Record<string, ExportOptionValue>>()
  #remembered: Record<string, Record<string, unknown>> = {}
  #openAfter = false
  #folder: string | null
  #busy = false
  #progress: number | null = null
  #warnings: readonly string[] = Object.freeze([])
  #abort: AbortController | null = null
  #listeners = new Set<() => void>()
  #state: ExportModelState
  #disposed = false

  /**
   * Builds the model and loads the remembered choices.
   * @param config see ExportModelConfig
   */
  static async create(config: ExportModelConfig): Promise<ExportModel> {
    const model = new ExportModel(config)
    await model.load()
    return model
  }

  /**
   * Builds the model with first-run defaults. Call load() (or use create()) to apply remembered choices.
   * @param config see ExportModelConfig
   */
  constructor(config: ExportModelConfig) {
    this.#config = config
    this.io = config.io
    this.module = config.module || (config.io ? config.io.module : null)
    this.documentName = String(config.documentName || '')
    this.documentPath = config.documentPath || null
    this.#rows = Object.freeze(sortRows(config.formats || [], config.order))
    this.#defs = Object.freeze([...(config.options || [])])
    this.#known = Object.freeze([...COMMON_EXTENSIONS, ...lower(config.knownExtensions || []), ...this.#rows.flatMap((row) => lower(row.extensions))])
    this.#folder = this.documentPath ? folderOf(this.documentPath) || null : null
    this.#selectedId = this.#initialFormat(undefined)
    this.#state = this.#computeState()
  }

  /** The dialog state (a new object after every change). */
  get state(): ExportModelState {
    return this.#state
  }

  /** The suggested file name for the selected format. */
  get fileName(): string {
    const row = this.#selectedRow()
    return row ? exportFileName(this.documentName, row, this.#fallbackStem(), this.#known) : ''
  }

  /**
   * Listens for state changes; returns the unsubscribe function
   * (works with React's useSyncExternalStore(model.subscribe, () => model.state)).
   */
  subscribe = (listener: () => void): (() => void) => {
    this.#listeners.add(listener)
    return () => { this.#listeners.delete(listener) }
  }

  /**
   * Reads the remembered format, folder, "open after exporting" and options
   * from io-prefs.json. Missing or invalid values keep the defaults.
   */
  async load(): Promise<void> {
    const io = this.io
    if (!io) return
    const read = async (key: string): Promise<unknown> => {
      try {
        return await io.prefs.get<unknown>(key)
      } catch (error) {
        this.#log(`Could not read the ${key} preference.`, error)
        return undefined
      }
    }
    const [format, folder, openAfter, options] = await Promise.all([
      read(EXPORT_PREF_KEYS.format),
      read(EXPORT_PREF_KEYS.folder),
      read(EXPORT_PREF_KEYS.openAfter),
      read(EXPORT_PREF_KEYS.options),
    ])
    if (this.#disposed) return
    if (typeof folder === 'string' && folder) this.#folder = folder
    if (typeof openAfter === 'boolean') this.#openAfter = openAfter
    if (isPlainObject(options)) {
      for (const [id, values] of Object.entries(options)) {
        if (isPlainObject(values)) this.#remembered[id] = { ...values }
      }
    }
    this.#selectedId = this.#initialFormat(typeof format === 'string' ? format : undefined)
    this.#emit()
  }

  /**
   * Selects a format. Rows that can't be produced on this PC can't be selected.
   * @param formatId registry id
   * @returns false when the row does not exist or is not available
   */
  select(formatId: string): boolean {
    const row = this.#rows.find((candidate) => candidate.id === formatId)
    if (!row || !row.available || this.#busy) return false
    if (this.#selectedId !== row.id) {
      this.#selectedId = row.id
      this.#emit()
    }
    return true
  }

  /**
   * Changes an option of the selected format. Invalid values are refused;
   * numbers are clamped to the option's range.
   * @param id option id
   * @param value new value
   * @returns false when the option does not apply or the value is invalid
   */
  setOption(id: string, value: ExportOptionValue): boolean {
    const formatId = this.#selectedId
    if (!formatId || this.#busy) return false
    const def = this.#defsFor(formatId).find((candidate) => candidate.id === id)
    if (!def) return false
    const accepted = coerceOptionValue(def, value)
    if (accepted === undefined) return false
    const values = { ...this.#currentValues(formatId), [id]: accepted }
    this.#values.set(formatId, values)
    this.#emit()
    return true
  }

  /**
   * Sets "Open the file after exporting". The file opens in the matching
   * Simple workspace, never in another program.
   * @param value checked or not
   */
  setOpenAfter(value: boolean): void {
    if (this.#openAfter === Boolean(value)) return
    this.#openAfter = Boolean(value)
    this.#emit()
  }

  /** Cancels a running export. The exporter must leave no partial file. */
  cancel(): void {
    if (this.#abort && !this.#abort.signal.aborted) this.#abort.abort()
  }

  /** Stops notifications; a running export continues to its end. */
  dispose(): void {
    this.#disposed = true
    this.#listeners.clear()
  }

  /**
   * Runs "Export…": commits pending edits, asks main for the target path
   * (remembered export folder, strict extensions), runs the exporter, handles
   * failures with the save-failure prompts, remembers the choices, reports
   * warnings, and opens the result in Simple when asked to.
   * @returns the outcome; `code` CANCELED when the user canceled
   */
  async run(): Promise<ExportOutcome> {
    if (this.#busy) return { ok: false, code: 'BLOCKED' }
    const row = this.#selectedRow()
    if (!row || !row.available) return { ok: false, code: 'UNSUPPORTED' }
    this.#busy = true
    this.#warnings = Object.freeze([])
    this.#emit()
    try {
      const session = this.#config.session
      if (session) {
        let ready = false
        try {
          ready = await session.prepare('export')
        } catch (error) {
          this.#log('Could not commit pending edits before exporting.', error)
        }
        if (!ready) return { ok: false, code: 'BLOCKED' }
      }
      const io = this.io
      if (!io) return { ok: false, code: 'UNKNOWN', failure: { ok: false, code: 'UNKNOWN', technical: 'window.simpleIO is not available.' } }

      let target = await this.#choosePath(io, 'export')
      if (!target) return { ok: false, code: 'CANCELED' }
      let force = false
      let recreate = false
      for (let round = 0; round < MAX_EXPORT_ROUNDS; round += 1) {
        let result: IoResult
        if ('ok' in target) result = target
        else result = this.#checkTarget(target) || await this.#runExporter(target, force, recreate)
        if (result.ok) return this.#succeeded(io, result)
        if (result.code === 'CANCELED') {
          this.#notify('info', catalogText('toast.exportCanceled'))
          return { ok: false, code: 'CANCELED', failure: result }
        }
        if (result.code === 'BLOCKED') {
          this.#notify('warning', result.message || catalogText('notices.BLOCKED'))
          return { ok: false, code: 'BLOCKED', failure: result }
        }
        let answer = await promptFailure(io, result)
        for (let shown = 0; answer === 'show-original' && shown < MAX_EXPORT_ROUNDS; shown += 1) {
          const original = result.backupPath || result.asidePath
          if (original) {
            try { await io.shell.showItem(original) } catch (error) { this.#log('Could not show the original file.', error) }
          }
          answer = await promptFailure(io, result)
        }
        if (answer === 'retry') {
          if ('ok' in target) target = await this.#choosePath(io, 'export')
        } else if (answer === 'replace') {
          force = true
        } else if (answer === 'recreate') {
          recreate = true
        } else if (answer === 'save-as' || answer === 'save-other-format') {
          // A folder problem (locked, read-only, missing, …) gets the fallback
          // folder; a name or format problem gets the export dialog again.
          const failedName = result.path ? baseName(result.path) : ''
          target = answer === 'save-as' && result.code !== 'INVALID_NAME'
            ? await this.#choosePath(io, 'fallback', result.code, failedName)
            : await this.#choosePath(io, 'export')
          force = false
          recreate = false
        } else {
          return { ok: false, code: result.code, failure: result }
        }
        if (!target) return { ok: false, code: 'CANCELED' }
      }
      return { ok: false, code: 'UNKNOWN' }
    } finally {
      this.#busy = false
      this.#progress = null
      this.#abort = null
      this.#emit()
    }
  }

  // -- Internals ---------------------------------------------------------------

  #log(message: string, detail?: unknown): void {
    if (this.#config.log) this.#config.log(message, detail)
    else console.warn('[simple-io]', message, detail === undefined ? '' : detail)
  }

  #notify(kind: NotifyKind, text: string, actions?: NotifyAction[]): void {
    const notify = this.#config.notify
    if (!notify || !text) return
    try {
      if (actions && actions.length) notify(kind, text, actions)
      else notify(kind, text)
    } catch (error) {
      this.#log('Could not show a notification.', error)
    }
  }

  #fallbackStem(): string {
    const kind = this.module ? kindName(this.module) : ''
    return catalogText('recovery.untitled', { kind: kind || 'document' }) || 'export'
  }

  #selectedRow(): ExportFormatRow | null {
    return this.#rows.find((row) => row.id === this.#selectedId) || null
  }

  #initialFormat(remembered: string | undefined): string | null {
    const selectable = (id: string | null | undefined) => Boolean(id) && this.#rows.some((row) => row.id === id && row.available)
    for (const id of [remembered, this.#config.defaultFormat, defaultExportFormat(this.module)]) {
      if (id && selectable(id)) return id
    }
    return this.#rows.find((row) => row.available)?.id ?? null
  }

  #defsFor(formatId: string): ExportOptionDef[] {
    return this.#defs.filter((def) => !def.formats || def.formats.includes(formatId))
  }

  /** Every applicable option's value: current, else remembered, else default. */
  #currentValues(formatId: string): Record<string, ExportOptionValue> {
    const current = this.#values.get(formatId)
    const remembered = this.#remembered[formatId] || {}
    const values: Record<string, ExportOptionValue> = {}
    for (const def of this.#defsFor(formatId)) {
      const fromCurrent = current ? current[def.id] : undefined
      const fromMemory = def.remember === false ? undefined : coerceOptionValue(def, remembered[def.id])
      const fallback = coerceOptionValue(def, def.defaultValue) ?? def.defaultValue
      values[def.id] = fromCurrent ?? fromMemory ?? fallback
    }
    return values
  }

  #visibleDefs(formatId: string, values: Readonly<Record<string, ExportOptionValue>>): ExportOptionDef[] {
    return this.#defsFor(formatId).filter((def) => {
      if (!def.visible) return true
      try {
        return Boolean(def.visible(values, formatId))
      } catch (error) {
        this.#log(`The visibility rule of the ${def.id} option failed.`, error)
        return true
      }
    })
  }

  #visibleValues(formatId: string): Record<string, ExportOptionValue> {
    const values = this.#currentValues(formatId)
    const visible: Record<string, ExportOptionValue> = {}
    for (const def of this.#visibleDefs(formatId, values)) {
      const value = values[def.id]
      if (value !== undefined) visible[def.id] = value
    }
    return visible
  }

  #computeState(): ExportModelState {
    const selected = this.#selectedRow()
    const rows = this.#rows.map((row): ExportRowView => Object.freeze({
      id: row.id,
      label: row.label,
      extension: row.extensions[0] ?? '',
      extensions: Object.freeze([...row.extensions]),
      description: row.description || '',
      tag: formatTag(row),
      selectable: Boolean(row.available),
      selected: row.id === this.#selectedId,
      lossy: Boolean(row.lossy && row.lossy.loses && row.lossy.loses.length),
      multiFile: Boolean(row.multiFile),
    }))
    let options: ExportOptionView[] = []
    if (selected) {
      const values = this.#currentValues(selected.id)
      options = this.#visibleDefs(selected.id, values).map((def) => Object.freeze({
        id: def.id,
        label: def.label,
        type: def.type,
        value: values[def.id] ?? def.defaultValue,
        description: def.description || '',
        choices: Object.freeze([...(def.choices || [])]),
        min: typeof def.min === 'number' ? def.min : null,
        max: typeof def.max === 'number' ? def.max : null,
        step: typeof def.step === 'number' ? def.step : null,
      }))
    }
    const progress = this.#progress
    return Object.freeze({
      title: catalogText('export.title', { name: this.documentName }),
      rows: Object.freeze(rows),
      selected: rows.find((row) => row.selected) || null,
      options: Object.freeze(options),
      lossNote: selected ? lossNote(selected) : '',
      openAfter: this.#openAfter,
      openAfterLabel: catalogText('export.openAfter'),
      folder: this.#folder,
      savesTo: this.#folder ? catalogText('export.savesTo', { folder: folderLabel(this.#folder) }) : '',
      fileName: this.fileName,
      busy: this.#busy,
      progress,
      statusText: this.#busy
        ? (progress === null ? catalogText('status.exporting') : catalogText('status.exportingPercent', { percent: Math.round(progress) }))
        : '',
      canExport: !this.#busy && Boolean(selected && selected.available),
      warnings: this.#warnings,
    })
  }

  #emit(): void {
    this.#state = this.#computeState()
    if (this.#disposed) return
    for (const listener of [...this.#listeners]) {
      try { listener() } catch (error) { this.#log('An export listener failed.', error) }
    }
  }

  /** Asks main for a target path; an exception becomes a failure the prompt loop handles. */
  async #choosePath(io: SimpleIO, purpose: 'export' | 'fallback', reason?: IoCode, failedName?: string): Promise<{ path: string; format: string } | IoFailure | null> {
    const row = this.#selectedRow()
    if (!row) return null
    const others = this.#rows.filter((candidate) => candidate.available && candidate.id !== row.id).map((candidate) => candidate.id)
    const request: ChooseSavePathRequest = {
      purpose,
      name: failedName || this.fileName,
      format: row.id,
      formats: [row.id, ...others],
    }
    const docId = this.#config.docId || this.#config.session?.docId
    if (docId) request.docId = docId
    if (reason) request.reason = reason
    try {
      const chosen = await io.chooseSavePath(request)
      if (!chosen || typeof chosen.path !== 'string' || !chosen.path) return null
      return { path: chosen.path, format: typeof chosen.format === 'string' && chosen.format ? chosen.format : row.id }
    } catch (error) {
      this.#log('Could not show the Save dialog.', error)
      return failureFromError(error)
    }
  }

  /**
   * Accepts the format main returned (typing another offered extension
   * switches the format) and refuses a path whose extension doesn't match it,
   * so a file is never written under a misleading name.
   */
  #checkTarget(target: { path: string; format: string }): IoFailure | null {
    const row = this.#rows.find((candidate) => candidate.id === target.format)
    const name = baseName(target.path)
    if (!row || !row.available) {
      return { ok: false, code: 'INVALID_NAME', name, path: target.path, technical: `Simple can't export the ${target.format || 'unknown'} format here.` }
    }
    if (row.id !== this.#selectedId) {
      this.#selectedId = row.id
      this.#emit()
    }
    if (!extensionMatches(target.path, row)) {
      return { ok: false, code: 'INVALID_NAME', name, path: target.path, technical: `The name doesn't end with ${row.extensions.join(' or ')}.` }
    }
    return null
  }

  async #runExporter(target: { path: string; format: string }, force: boolean, recreate: boolean): Promise<IoResult> {
    const abort = new AbortController()
    this.#abort = abort
    this.#progress = null
    this.#emit()
    const request: ExportRequest = {
      path: target.path,
      format: target.format,
      options: this.#visibleValues(target.format),
      signal: abort.signal,
      onProgress: (percent: number) => {
        if (this.#abort !== abort || typeof percent !== 'number' || !Number.isFinite(percent)) return
        this.#progress = Math.min(100, Math.max(0, percent))
        this.#emit()
      },
    }
    const docId = this.#config.docId || this.#config.session?.docId
    if (docId) request.docId = docId
    if (force) request.force = true
    if (recreate) request.recreate = true
    const work = () => this.#config.exporter(request)
    try {
      const session = this.#config.session
      const raw = session && session.track ? await session.track('export', work) : await work()
      const result = toIoResult(raw)
      if (!result.ok && abort.signal.aborted && result.code !== 'CANCELED') return { ...result, code: 'CANCELED' }
      return result
    } catch (error) {
      return abort.signal.aborted ? { ok: false, code: 'CANCELED' } : failureFromError(error)
    }
  }

  #remember(io: SimpleIO, formatId: string, folder: string): void {
    const set = (key: string, value: unknown) => {
      try {
        void io.prefs.set(key, value).catch((error: unknown) => this.#log(`Could not remember ${key}.`, error))
      } catch (error) {
        this.#log(`Could not remember ${key}.`, error)
      }
    }
    set(EXPORT_PREF_KEYS.format, formatId)
    if (folder) set(EXPORT_PREF_KEYS.folder, folder)
    set(EXPORT_PREF_KEYS.openAfter, this.#openAfter)
    const remembered: Record<string, ExportOptionValue> = {}
    const values = this.#currentValues(formatId)
    for (const def of this.#defsFor(formatId)) {
      const value = values[def.id]
      if (def.remember !== false && value !== undefined) remembered[def.id] = value
    }
    this.#remembered = { ...this.#remembered, [formatId]: remembered }
    set(EXPORT_PREF_KEYS.options, this.#remembered)
  }

  async #succeeded(io: SimpleIO, result: IoSuccess): Promise<ExportOutcome> {
    const path = result.path
    const name = result.name || baseName(path)
    const folder = folderOf(path)
    const format = result.format || this.#selectedId || ''
    const warnings = (result.warnings || []).filter((warning) => typeof warning === 'string' && warning)
    if (folder) this.#folder = folder
    this.#warnings = Object.freeze([...warnings])
    this.#remember(io, this.#selectedId || format, folder)
    if (result.ownFileChanged) {
      // The document's own file now holds the export, not the document: Save must write it again.
      try {
        this.#config.session?.noteOwnFileChanged?.()
      } catch (error) {
        this.#log('Could not mark the document changed after exporting over its file.', error)
      }
    }

    const actions: NotifyAction[] = []
    const onShowWarnings = this.#config.onShowWarnings
    if (warnings.length && onShowWarnings) {
      actions.push({ label: catalogText('toast.actions.details'), run: () => onShowWarnings(warnings, name) })
    }
    actions.push({ label: catalogText('toast.actions.open'), run: () => { void this.#open(io, path) } })
    actions.push({
      label: catalogText('toast.actions.showInFolder'),
      run: () => { void io.shell.showItem(path).catch((error: unknown) => this.#log('Could not show the file in its folder.', error)) },
    })
    if (warnings.length) this.#notify('warning', catalogText('toast.exportedWithWarnings', { name, count: warnings.length }), actions)
    else this.#notify('success', catalogText('toast.exported', { name, folder: folderLabel(folder) }), actions)

    let opened = false
    if (this.#openAfter) opened = await this.#open(io, path)
    return { ok: true, path, name, format, warnings: [...warnings], opened }
  }

  /**
   * Opens an exported file in the Simple workspace that owns its format, never
   * in another program. A format no Simple workspace opens is shown in File
   * Explorer instead, and the toast says why.
   */
  async #open(io: SimpleIO, path: string): Promise<boolean> {
    const result = await openInSimple(io, path)
    const opened = openedInSimple(result)
    if (opened && result.action === 'launched' && result.appName) {
      this.#notify('info', catalogText('toast.openedIn', { appName: result.appName }))
    } else if (result.ok && !opened && result.reason === 'unsupported') {
      this.#notify('info', catalogText('toast.shownInFolder', { name: baseName(path) }))
    }
    return opened
  }
}
