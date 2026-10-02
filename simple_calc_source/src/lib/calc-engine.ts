import type { CellData, CellScalar, DefinedName, SheetData, WorkbookModel } from '../spreadsheet-types'
import {
  diagnoseFormula,
  evaluateFormulaDetailed,
  formulaMayReturnArray,
  isFormulaError,
} from './formulas'
import type { FormulaDiagnostic, FormulaEvaluationHooks, FormulaRangeBounds, FormulaResult, FormulaTableInfo } from './formulas'

/**
 * Workbook calculation engine.
 *
 * Formula values are computed lazily and cached across workbook revisions. Every
 * evaluation records the rectangles it read, so an edit only invalidates the formulas
 * that (transitively) depend on the changed cells — the way Excel's calc chain works.
 * Dynamic arrays spill into neighbouring empty cells and report #SPILL! when blocked.
 */

const MAX_ROWS = 1_048_576
const MAX_COLUMNS = 16_384
const FALLBACK_ERRORS = new Set(['#NAME?', '#PARSE!', '#ERROR!'])

type Key = string

/**
 * Excel's calculation options (File > Options > Formulas), kept in the workbook's
 * `metadata.calcProperties` (the file's <calcPr>) so they travel with the file.
 */
export interface CalculationOptions {
  /** 'automatic' recalculates dependents on every change; 'manual' waits for recalculate() (F9). */
  mode: 'automatic' | 'manual'
  /** Resolve circular references by iteration instead of reporting them as #CIRC!. */
  iterate: boolean
  /** Most iterations per recalculation (Excel's default is 100). */
  maxIterations: number
  /** Iteration stops once no value in the cycle changes by more than this (default 0.001). */
  maxChange: number
}

export const DEFAULT_CALCULATION_OPTIONS: CalculationOptions = { mode: 'automatic', iterate: false, maxIterations: 100, maxChange: 0.001 }

function flagValue(value: unknown) {
  return value === true || value === 1 || value === '1' || (typeof value === 'string' && value.toLowerCase() === 'true')
}

/** The calculation options stored in a workbook (defaults for anything missing). */
export function calculationOptionsOf(workbook: WorkbookModel): CalculationOptions {
  const properties = workbook.metadata?.calcProperties || {}
  const count = properties.iterateCount === true ? 1 : Number(properties.iterateCount)
  const delta = Number(properties.iterateDelta)
  return {
    mode: String(properties.calcMode || '').toLowerCase() === 'manual' ? 'manual' : 'automatic',
    iterate: flagValue(properties.iterate),
    maxIterations: Number.isFinite(count) && count >= 1 ? Math.min(32_767, Math.trunc(count)) : DEFAULT_CALCULATION_OPTIONS.maxIterations,
    maxChange: Number.isFinite(delta) && delta >= 0 ? delta : DEFAULT_CALCULATION_OPTIONS.maxChange,
  }
}

/** A copy of the workbook with calculation options written to `metadata.calcProperties`. */
export function withCalculationOptions(workbook: WorkbookModel, options: Partial<CalculationOptions>): WorkbookModel {
  const next = { ...calculationOptionsOf(workbook), ...options }
  const calcProperties: Record<string, unknown> = {
    ...(workbook.metadata?.calcProperties || {}),
    calcMode: next.mode === 'manual' ? 'manual' : 'auto',
    iterate: next.iterate,
    iterateCount: next.maxIterations,
    iterateDelta: next.maxChange,
  }
  return { ...workbook, metadata: { ...(workbook.metadata || {}), calcProperties } }
}

/**
 * Imported formulas from a workbook file that Excel reads as legacy (pre-dynamic-array)
 * formulas carry `implicitIntersection: true`; they are intersected instead of spilling.
 */
type EngineCell = CellData & { implicitIntersection?: boolean }

function isLegacyFormula(cell: CellData | undefined) {
  return Boolean(cell?.formula) && (cell as EngineCell).implicitIntersection === true && cell!.formulaType !== 'array'
}

function iterationChange(before: CellScalar | undefined, after: CellScalar | undefined) {
  if (Object.is(before, after)) return 0
  if (typeof before === 'number' && typeof after === 'number') return Math.abs(after - before)
  return Infinity
}

interface Rect {
  sheetId: string
  top: number
  left: number
  bottom: number
  right: number
}

interface SpillInfo {
  top: number
  left: number
  bottom: number
  right: number
  columns: number
  values: FormulaResult[]
}

interface Frame {
  key: Key
  rects: Rect[]
  volatile: boolean
  visibility: boolean
  discard: boolean
}

/**
 * Spatial index of precedents on one sheet. Single cells are indexed by row/column;
 * narrow ranges (<= 16 columns) by column and 256-row bucket, or by column alone when they
 * are tall (A:A); everything wider is kept in a short list that is scanned directly.
 */
interface SheetDependencyIndex {
  cells: Map<number, Map<number, Set<Key>>>
  narrow: Map<number, Map<number, Set<Key>>>
  tall: Map<number, Set<Key>>
  wide: Map<Key, Rect[]>
}

const BUCKET_SHIFT = 8
const NARROW_COLUMNS = 16
const MAX_NARROW_BUCKETS = 64

function addToSet<K>(map: Map<K, Set<Key>>, key: K, value: Key) {
  let set = map.get(key)
  if (!set) {
    set = new Set()
    map.set(key, set)
  }
  set.add(value)
}

function removeFromSet<K>(map: Map<K, Set<Key>>, key: K, value: Key) {
  const set = map.get(key)
  if (!set) return
  set.delete(value)
  if (!set.size) map.delete(key)
}

export interface CellCoordinate {
  row: number
  col: number
}

export type ChangeHint = Map<string, Set<string> | 'all'>

function columnLabel(index: number) {
  let value = index + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

const COLUMN_LABELS: string[] = []
function cachedColumnLabel(index: number) {
  if (index < 2048) {
    let label = COLUMN_LABELS[index]
    if (label === undefined) {
      label = columnLabel(index)
      COLUMN_LABELS[index] = label
    }
    return label
  }
  return columnLabel(index)
}

export function cellAddress(row: number, col: number) {
  return `${cachedColumnLabel(col)}${row + 1}`
}

export function parseCellAddress(address: string): CellCoordinate | null {
  let index = 0
  let col = 0
  const length = address.length
  if (address.charCodeAt(0) === 36) index += 1
  const start = index
  while (index < length) {
    const code = address.charCodeAt(index) & ~32
    if (code < 65 || code > 90) break
    col = col * 26 + code - 64
    index += 1
  }
  if (index === start || index - start > 3) return null
  if (address.charCodeAt(index) === 36) index += 1
  let row = 0
  const digits = index
  while (index < length) {
    const code = address.charCodeAt(index)
    if (code < 48 || code > 57) return null
    row = row * 10 + code - 48
    index += 1
  }
  if (index === digits || row < 1) return null
  return { row: row - 1, col: col - 1 }
}

function parseRange(range: string): { top: number; left: number; bottom: number; right: number } | null {
  const [startText, endText = startText] = range.replace(/\$/g, '').split(':')
  const start = parseCellAddress(startText)
  const end = parseCellAddress(endText)
  if (!start || !end) return null
  return {
    top: Math.min(start.row, end.row),
    left: Math.min(start.col, end.col),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.col, end.col),
  }
}

function hasStoredValue(cell: CellData | undefined) {
  if (!cell) return false
  if (cell.formula) return true
  if (cell.arrayMember) return false
  return cell.value !== undefined && cell.value !== null && cell.value !== ''
}

function rectsIntersect(a: Rect, b: Rect) {
  return a.sheetId === b.sheetId && a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right
}

function isVisibilitySensitive(formula: string) {
  return /SUBTOTAL|AGGREGATE/i.test(formula)
}

export class CalculationEngine {
  private workbook: WorkbookModel
  private sheetById = new Map<string, SheetData>()
  private sheetByName = new Map<string, SheetData>()
  private sheetNames: string[] = []
  private values = new Map<Key, CellScalar>()
  private stale = new Set<Key>()
  private deps = new Map<Key, Rect[]>()
  private dependencyIndex = new Map<string, SheetDependencyIndex>()
  private volatileKeys = new Set<Key>()
  private visibilityKeys = new Set<Key>()
  private spills = new Map<Key, SpillInfo>()
  private spillAttempts = new Map<Key, Rect>()
  private spillCover = new Map<Key, Key>()
  private candidates = new Map<string, Set<string>>()
  private spillReady = new Set<string>()
  private usedRanges = new Map<string, { maxRow: number; maxCol: number }>()
  private sortedCells = new Map<string, { cells: SheetData['cells']; coordinates: Array<[number, number]> }>()
  private parsedMerges = new WeakMap<string[], Array<{ top: number; left: number; bottom: number; right: number }>>()
  private hiddenSets = new WeakMap<number[], Set<number>>()
  private stack: Frame[] = []
  private stackIndex = new Map<Key, number>()
  private speculativeRoot = -1
  private discardFrom = Infinity
  /**
   * A candidate value for one cell while evaluating a data-validation rule for a pending edit.
   * Formulas depending on that cell are recalculated (never from cache) and nothing computed
   * under the override is cached.
   */
  private override: { key: Key; sheetId: string; row: number; col: number; value: CellScalar; bypass: Set<Key> } | null = null
  private revision = 0
  private options: CalculationOptions
  /** Changes a manual-mode workbook has not recalculated yet. */
  private pendingRects: Rect[] = []
  private pendingKeys = new Set<Key>()
  private pendingVolatile = false
  private pendingAll = false
  /** Formula cells found on a circular reference, in the order they were found. */
  private circular = new Set<Key>()
  /** Circular cells awaiting iteration, and each one's value from the previous iteration. */
  private iterationPending = new Set<Key>()
  private iterationValues = new Map<Key, CellScalar>()
  private iterating = false
  /** Formulas whose result relies on array evaluation where a legacy formula would intersect. */
  private arrayEvaluated = new Set<Key>()

  constructor(workbook: WorkbookModel) {
    this.workbook = workbook
    this.options = calculationOptionsOf(workbook)
    this.indexSheets()
  }

  /** The workbook's calculation options (mode, iteration). */
  get calculationOptions(): CalculationOptions {
    return { ...this.options }
  }

  /** Manual calculation mode has changes that recalculate() has not applied yet ("Calculate"). */
  get needsRecalculation() {
    return this.options.mode === 'manual' &&
      (this.pendingRects.length > 0 || this.pendingKeys.size > 0 || this.pendingVolatile || this.pendingAll)
  }

  /** Current revision number; increases whenever any cached value may have changed. */
  get version() {
    return this.revision
  }

  /** The workbook revision the engine currently reflects. */
  get current() {
    return this.workbook
  }

  // ---- Workbook updates --------------------------------------------------------------------

  private indexSheets() {
    this.sheetById.clear()
    this.sheetByName.clear()
    this.sheetNames = []
    for (const sheet of this.workbook.sheets) {
      this.sheetById.set(sheet.id, sheet)
      this.sheetByName.set(sheet.name.toLocaleLowerCase(), sheet)
      this.sheetNames.push(sheet.name)
    }
  }

  private resetAll() {
    this.values.clear()
    this.stale.clear()
    this.deps.clear()
    this.dependencyIndex.clear()
    this.volatileKeys.clear()
    this.visibilityKeys.clear()
    this.spills.clear()
    this.spillAttempts.clear()
    this.spillCover.clear()
    this.candidates.clear()
    this.spillReady.clear()
    this.usedRanges.clear()
    this.sortedCells.clear()
    this.circular.clear()
    this.iterationPending.clear()
    this.iterationValues.clear()
    this.arrayEvaluated.clear()
    this.pendingRects = []
    this.pendingKeys.clear()
    this.pendingVolatile = false
    this.pendingAll = false
  }

  /**
   * Adopt a new workbook revision. `hint` lists the addresses known to have changed per
   * sheet id (from immer patches); sheets without a hint are diffed by object identity.
   */
  update(next: WorkbookModel, hint?: ChangeHint) {
    const previous = this.workbook
    if (next === previous) return
    this.workbook = next
    this.revision += 1

    const structural = previous.sheets.length !== next.sheets.length ||
      previous.sheets.some((sheet, index) => sheet.id !== next.sheets[index].id || sheet.name !== next.sheets[index].name) ||
      previous.definedNames !== next.definedNames ||
      previous.metadata?.definedNames !== next.metadata?.definedNames
    const previousById = new Map(previous.sheets.map((sheet) => [sheet.id, sheet]))
    this.indexSheets()
    let releasePending = false
    if (previous.metadata?.calcProperties !== next.metadata?.calcProperties) {
      const options = calculationOptionsOf(next)
      const iterationChanged = options.iterate !== this.options.iterate ||
        options.maxIterations !== this.options.maxIterations || options.maxChange !== this.options.maxChange
      releasePending = this.options.mode === 'manual' && options.mode === 'automatic'
      this.options = options
      if (iterationChanged) {
        this.resetAll()
        return
      }
    }
    if (structural) {
      this.resetAll()
      return
    }

    const manual = this.options.mode === 'manual'
    const changedRects: Rect[] = []
    let anyChange = false
    for (const sheet of next.sheets) {
      const old = previousById.get(sheet.id)
      if (!old || old === sheet) continue
      if (old.tables !== sheet.tables) {
        this.resetAll()
        return
      }
      if (old.merges !== sheet.merges) this.invalidateSpillAnchors(sheet.id, changedRects)
      if (old.hiddenRows !== sheet.hiddenRows || old.filteredRows !== sheet.filteredRows) {
        // SUBTOTAL/AGGREGATE formulas anywhere in the workbook that read this sheet's rows.
        for (const key of [...this.visibilityKeys]) {
          if (!key.startsWith(`${sheet.id}!`) && !this.deps.get(key)?.some((rect) => rect.sheetId === sheet.id)) continue
          if (manual) this.pendingKeys.add(key)
          else this.invalidateFormula(key, changedRects)
        }
        anyChange = true
      }
      if (old.cells === sheet.cells) continue
      anyChange = true
      this.sortedCells.delete(sheet.id)
      const sheetHint = hint?.get(sheet.id)
      const addresses = sheetHint && sheetHint !== 'all' ? sheetHint : this.diffCells(old.cells, sheet.cells)
      for (const address of addresses) this.cellChanged(sheet, old.cells[address], sheet.cells[address], address, changedRects)
    }
    if (releasePending) {
      // Switching back to automatic calculation applies everything manual mode held back.
      this.applyPending(changedRects)
      return
    }
    if (!anyChange && !changedRects.length) return
    if (manual) {
      // Manual calculation: the edited formulas themselves recalculate; their dependents and
      // volatile functions wait for recalculate() (F9).
      this.pendingVolatile = true
      if (this.pendingRects.length + changedRects.length > 50_000) this.pendingAll = true
      else this.pendingRects.push(...changedRects)
      return
    }
    for (const key of [...this.volatileKeys]) this.invalidateFormula(key, changedRects)
    this.propagate(changedRects)
  }

  /**
   * F9: recalculate everything changed since the last calculation (manual mode) and every
   * volatile function (NOW, RAND, OFFSET, ...) and circular reference.
   */
  recalculate() {
    if (this.stack.length) return
    this.revision += 1
    this.applyPending([])
  }

  private applyPending(extraRects: Rect[]) {
    if (this.pendingAll) {
      this.resetAll()
      return
    }
    const rects = [...this.pendingRects, ...extraRects]
    this.pendingRects = []
    this.pendingVolatile = false
    for (const key of this.pendingKeys) this.invalidateFormula(key, rects)
    this.pendingKeys.clear()
    for (const key of [...this.volatileKeys]) this.invalidateFormula(key, rects)
    this.propagate(rects)
  }

  /** Ctrl+Alt+F9: discard every calculated value and calculate the whole workbook again. */
  recalculateAll() {
    if (this.stack.length) return
    this.revision += 1
    this.resetAll()
  }

  /**
   * Shift+F9: recalculate the formulas on one sheet. In manual mode other sheets keep their
   * values until recalculate().
   */
  recalculateSheet(sheetId: string) {
    if (this.stack.length) return
    this.revision += 1
    const rects: Rect[] = []
    for (const key of [...this.values.keys()]) {
      if (key.startsWith(`${sheetId}!`)) this.invalidateFormula(key, rects)
    }
    this.spillReady.delete(sheetId)
    if (this.options.mode !== 'manual') this.propagate(rects)
  }

  /**
   * Formula cells found on circular references (each cell of a cycle, in the order found), as
   * Excel lists under Formulas > Error Checking > Circular References. Only calculated cells
   * are known; calculateAll() first for a complete list.
   */
  circularReferences(): Array<{ sheetId: string; address: string }> {
    return [...this.circular].map((key) => {
      const separator = key.indexOf('!')
      return { sheetId: key.slice(0, separator), address: key.slice(separator + 1) }
    })
  }

  private diffCells(previous: SheetData['cells'], next: SheetData['cells']) {
    const changed = new Set<string>()
    for (const address in next) {
      if (previous[address] !== next[address]) changed.add(address)
    }
    for (const address in previous) {
      if (!(address in next)) changed.add(address)
    }
    return changed
  }

  private cellChanged(sheet: SheetData, before: CellData | undefined, after: CellData | undefined, address: string, changedRects: Rect[]) {
    const coord = parseCellAddress(address)
    if (!coord) return
    const key = `${sheet.id}!${address}`
    const candidates = this.candidates.get(sheet.id)
    if (candidates) {
      if (after?.formula && !isLegacyFormula(after) && formulaMayReturnArray(after.formula)) candidates.add(address)
      else candidates.delete(address)
    }
    if (after && (after.formula || after.value !== undefined)) {
      const used = this.usedRanges.get(sheet.id)
      if (used) {
        used.maxRow = Math.max(used.maxRow, coord.row + 1)
        used.maxCol = Math.max(used.maxCol, coord.col + 1)
      }
    }
    if (before?.formula || after?.formula || this.values.has(key)) {
      this.invalidateFormula(key, changedRects)
    } else {
      changedRects.push({ sheetId: sheet.id, top: coord.row, left: coord.col, bottom: coord.row, right: coord.col })
    }
  }

  private invalidateSpillAnchors(sheetId: string, changedRects: Rect[]) {
    for (const key of [...this.spillAttempts.keys()]) {
      if (key.startsWith(`${sheetId}!`)) this.invalidateFormula(key, changedRects)
    }
    this.spillReady.delete(sheetId)
  }

  private indexRect(key: Key, rect: Rect, add: boolean) {
    let index = this.dependencyIndex.get(rect.sheetId)
    if (!index) {
      if (!add) return
      index = { cells: new Map(), narrow: new Map(), tall: new Map(), wide: new Map() }
      this.dependencyIndex.set(rect.sheetId, index)
    }
    if (rect.top === rect.bottom && rect.left === rect.right) {
      let row = index.cells.get(rect.top)
      if (add) {
        if (!row) {
          row = new Map()
          index.cells.set(rect.top, row)
        }
        addToSet(row, rect.left, key)
      } else if (row) {
        removeFromSet(row, rect.left, key)
        if (!row.size) index.cells.delete(rect.top)
      }
      return
    }
    const width = rect.right - rect.left + 1
    if (width <= NARROW_COLUMNS) {
      const firstBucket = rect.top >> BUCKET_SHIFT
      const lastBucket = rect.bottom >> BUCKET_SHIFT
      for (let col = rect.left; col <= rect.right; col += 1) {
        if (lastBucket - firstBucket + 1 <= MAX_NARROW_BUCKETS) {
          let buckets = index.narrow.get(col)
          if (!buckets) {
            if (!add) continue
            buckets = new Map()
            index.narrow.set(col, buckets)
          }
          for (let bucket = firstBucket; bucket <= lastBucket; bucket += 1) {
            if (add) addToSet(buckets, bucket, key)
            else removeFromSet(buckets, bucket, key)
          }
          if (!buckets.size) index.narrow.delete(col)
        } else if (add) {
          addToSet(index.tall, col, key)
        } else {
          removeFromSet(index.tall, col, key)
        }
      }
      return
    }
    if (add) {
      const list = index.wide.get(key)
      if (list) list.push(rect)
      else index.wide.set(key, [rect])
    } else {
      index.wide.delete(key)
    }
  }

  private removeDependencies(key: Key) {
    const rects = this.deps.get(key)
    if (!rects) return
    this.deps.delete(key)
    for (const rect of rects) this.indexRect(key, rect, false)
  }

  private addDependencies(key: Key, rects: Rect[]) {
    // Collapse duplicates (the same range read twice) to keep the index small.
    const unique: Rect[] = []
    const seen = new Set<string>()
    for (const rect of rects) {
      const signature = `${rect.sheetId}|${rect.top}|${rect.left}|${rect.bottom}|${rect.right}`
      if (seen.has(signature)) continue
      seen.add(signature)
      unique.push(rect)
    }
    this.deps.set(key, unique)
    for (const rect of unique) this.indexRect(key, rect, true)
  }

  /** Formula keys whose precedents intersect a rectangle. */
  private dependentsOfRect(rect: Rect): Key[] {
    const index = this.dependencyIndex.get(rect.sheetId)
    if (!index) return []
    const candidates = new Set<Key>()
    const height = rect.bottom - rect.top + 1
    if (height <= index.cells.size) {
      for (let row = rect.top; row <= rect.bottom; row += 1) {
        const columns = index.cells.get(row)
        if (!columns) continue
        if (rect.right - rect.left + 1 <= columns.size) {
          for (let col = rect.left; col <= rect.right; col += 1) {
            const set = columns.get(col)
            if (set) for (const key of set) candidates.add(key)
          }
        } else {
          for (const [col, set] of columns) if (col >= rect.left && col <= rect.right) for (const key of set) candidates.add(key)
        }
      }
    } else {
      for (const [row, columns] of index.cells) {
        if (row < rect.top || row > rect.bottom) continue
        for (const [col, set] of columns) if (col >= rect.left && col <= rect.right) for (const key of set) candidates.add(key)
      }
    }
    const visitColumn = (col: number) => {
      const buckets = index.narrow.get(col)
      if (buckets) {
        const firstBucket = rect.top >> BUCKET_SHIFT
        const lastBucket = rect.bottom >> BUCKET_SHIFT
        if (lastBucket - firstBucket + 1 <= buckets.size) {
          for (let bucket = firstBucket; bucket <= lastBucket; bucket += 1) {
            const set = buckets.get(bucket)
            if (set) for (const key of set) candidates.add(key)
          }
        } else {
          for (const [bucket, set] of buckets) if (bucket >= firstBucket && bucket <= lastBucket) for (const key of set) candidates.add(key)
        }
      }
      const tall = index.tall.get(col)
      if (tall) for (const key of tall) candidates.add(key)
    }
    const width = rect.right - rect.left + 1
    if (width <= index.narrow.size + index.tall.size) {
      for (let col = rect.left; col <= rect.right; col += 1) visitColumn(col)
    } else {
      const columns = new Set<number>([...index.narrow.keys(), ...index.tall.keys()])
      for (const col of columns) if (col >= rect.left && col <= rect.right) visitColumn(col)
    }
    for (const [key, rects] of index.wide) {
      if (rects.some((candidate) => rectsIntersect(candidate, rect))) candidates.add(key)
    }
    const output: Key[] = []
    for (const key of candidates) {
      const rects = this.deps.get(key)
      if (rects && rects.some((candidate) => rectsIntersect(candidate, rect))) output.push(key)
    }
    return output
  }

  private clearSpill(key: Key, changedRects?: Rect[]) {
    const spill = this.spills.get(key)
    if (!spill) return
    this.spills.delete(key)
    const sheetId = key.slice(0, key.indexOf('!'))
    for (let row = spill.top; row <= spill.bottom; row += 1) {
      for (let col = spill.left; col <= spill.right; col += 1) {
        const coveredKey = `${sheetId}!${cellAddress(row, col)}`
        if (this.spillCover.get(coveredKey) === key) this.spillCover.delete(coveredKey)
      }
    }
    changedRects?.push({ sheetId, top: spill.top, left: spill.left, bottom: spill.bottom, right: spill.right })
  }

  private invalidateFormula(key: Key, changedRects: Rect[]) {
    this.values.delete(key)
    this.stale.delete(key)
    this.removeDependencies(key)
    this.volatileKeys.delete(key)
    this.visibilityKeys.delete(key)
    this.circular.delete(key)
    this.arrayEvaluated.delete(key)
    this.clearSpill(key, changedRects)
    this.spillAttempts.delete(key)
    const separator = key.indexOf('!')
    const sheetId = key.slice(0, separator)
    this.spillReady.delete(sheetId)
    const coord = parseCellAddress(key.slice(separator + 1))
    if (coord) changedRects.push({ sheetId, top: coord.row, left: coord.col, bottom: coord.row, right: coord.col })
  }

  /** Invalidate every formula that depends on the changed rectangles, transitively. */
  private propagate(initial: Rect[]) {
    const queue = initial.slice()
    let guard = 0
    while (queue.length) {
      if ((guard += 1) > 2_000_000) {
        this.resetAll()
        return
      }
      const rect = queue.pop()!
      // An anchor whose spill area now contains a value (or lost one) must re-check blocking.
      for (const [anchor, attempt] of this.spillAttempts) {
        if (!rectsIntersect(attempt, rect)) continue
        const anchorCoord = parseCellAddress(anchor.slice(anchor.indexOf('!') + 1))
        if (anchorCoord && anchorCoord.row === rect.top && anchorCoord.col === rect.left && rect.top === rect.bottom && rect.left === rect.right) continue
        if (this.values.has(anchor)) this.invalidateFormula(anchor, queue)
      }
      for (const key of this.dependentsOfRect(rect)) {
        if (this.values.has(key) || this.deps.has(key)) this.invalidateFormula(key, queue)
      }
    }
  }

  // ---- Sheet lookups -----------------------------------------------------------------------

  private resolveSheet(reference: string): SheetData | undefined {
    return this.sheetById.get(reference) || this.sheetByName.get(reference.toLocaleLowerCase())
  }

  /** Extent of a sheet's cells and spilled arrays (whole-column/row references stop here). */
  private usedRangeOf(sheet: SheetData) {
    let used = this.usedRanges.get(sheet.id)
    if (!used) {
      let maxRow = 1
      let maxCol = 1
      for (const address in sheet.cells) {
        const coord = parseCellAddress(address)
        if (!coord) continue
        if (coord.row + 1 > maxRow) maxRow = coord.row + 1
        if (coord.col + 1 > maxCol) maxCol = coord.col + 1
      }
      // Spilled members are not cells, but A:A and 1:1 must still reach them.
      const prefix = `${sheet.id}!`
      for (const [key, spill] of this.spills) {
        if (!key.startsWith(prefix)) continue
        if (spill.bottom + 1 > maxRow) maxRow = spill.bottom + 1
        if (spill.right + 1 > maxCol) maxCol = spill.right + 1
      }
      used = { maxRow, maxCol }
      this.usedRanges.set(sheet.id, used)
    }
    return used
  }

  private sortedCoordinates(sheet: SheetData) {
    let entry = this.sortedCells.get(sheet.id)
    if (!entry || entry.cells !== sheet.cells) {
      const coordinates: Array<[number, number]> = []
      for (const address in sheet.cells) {
        const coord = parseCellAddress(address)
        if (coord) coordinates.push([coord.row + 1, coord.col + 1])
      }
      coordinates.sort((a, b) => a[0] - b[0] || a[1] - b[1])
      entry = { cells: sheet.cells, coordinates }
      this.sortedCells.set(sheet.id, entry)
    }
    return entry.coordinates
  }

  private mergesOf(sheet: SheetData) {
    const merges = sheet.merges || []
    let parsed = this.parsedMerges.get(merges)
    if (!parsed) {
      parsed = merges.map(parseRange).filter((value): value is NonNullable<typeof value> => Boolean(value))
      this.parsedMerges.set(merges, parsed)
    }
    return parsed
  }

  private hiddenSet(values: number[] | undefined) {
    if (!values?.length) return null
    let set = this.hiddenSets.get(values)
    if (!set) {
      set = new Set(values.map(Number))
      this.hiddenSets.set(values, set)
    }
    return set
  }

  private candidateSet(sheet: SheetData) {
    let set = this.candidates.get(sheet.id)
    if (!set) {
      set = new Set()
      for (const address in sheet.cells) {
        const cell = sheet.cells[address]
        if (cell?.formula && !isLegacyFormula(cell) && formulaMayReturnArray(cell.formula)) set.add(address)
      }
      this.candidates.set(sheet.id, set)
    }
    return set
  }

  private definedNameValue(name: string, sheetId?: string): string | undefined {
    const wanted = name.toLocaleLowerCase()
    const sheetIndex = sheetId ? this.workbook.sheets.findIndex((sheet) => sheet.id === sheetId) : -1
    const candidates = (this.workbook.definedNames || []).filter((item: DefinedName) => item.name.toLocaleLowerCase() === wanted)
    const local = candidates.find((item) => (item.localSheetIndex !== undefined && item.localSheetIndex === sheetIndex) ||
      (item.localSheetId !== undefined && item.localSheetId === sheetIndex))
    const global = candidates.find((item) => item.localSheetIndex === undefined && item.localSheetId === undefined)
    const entry = local || global || candidates[0]
    if (entry) {
      if (entry.ranges?.length > 1) return entry.ranges[0]
      return entry.ranges?.[0] || entry.ref
    }
    const metadata = this.workbook.metadata?.definedNames?.filter((item) => item.name.toLocaleLowerCase() === wanted) || []
    const metadataEntry = metadata.find((item) => item.localSheetId === sheetIndex) || metadata.find((item) => item.localSheetId === undefined) || metadata[0]
    return metadataEntry ? metadataEntry.ranges || metadataEntry.formula : undefined
  }

  private tableInfo(name: string | null, sheetReference: string, currentCell?: { row: number; column: number }): FormulaTableInfo | null {
    const sheets = name === null ? [this.resolveSheet(sheetReference)].filter(Boolean) as SheetData[] : this.workbook.sheets
    for (const sheet of sheets) {
      for (const table of sheet.tables || []) {
        const bounds = parseRange(table.ref)
        if (!bounds) continue
        if (name !== null) {
          const wanted = name.toLocaleLowerCase()
          if (table.name.toLocaleLowerCase() !== wanted && (table.displayName || '').toLocaleLowerCase() !== wanted) continue
        } else if (!currentCell || currentCell.row - 1 < bounds.top || currentCell.row - 1 > bounds.bottom || currentCell.column - 1 < bounds.left || currentCell.column - 1 > bounds.right) {
          continue
        }
        const width = bounds.right - bounds.left + 1
        const columns = Array.from({ length: width }, (_, offset) => {
          const declared = table.columns[offset]?.name
          if (declared) return declared
          const header = sheet.cells[cellAddress(bounds.top, bounds.left + offset)]
          return header?.value == null ? `Column${offset + 1}` : String(header.value)
        })
        return {
          sheetId: sheet.id,
          name: table.name,
          startRow: bounds.top + 1,
          endRow: bounds.bottom + 1,
          startColumn: bounds.left + 1,
          endColumn: bounds.right + 1,
          headerRowCount: table.headerRow === false ? 0 : 1,
          totalsRowCount: table.totalsRow ? 1 : 0,
          columns,
        }
      }
    }
    return null
  }

  // ---- Evaluation ---------------------------------------------------------------------------

  private readonly resolver = (sheetReference: string, address: string): CellScalar => {
    const sheet = this.resolveSheet(sheetReference)
    if (!sheet) return '#REF!'
    const normalized = address.charCodeAt(0) === 36 || address.includes('$') ? address.replace(/\$/g, '').toUpperCase() : address.toUpperCase()
    return this.valueOf(sheet, normalized)
  }

  private hooksFor(sheet: SheetData, coord: CellCoordinate, legacy = false): FormulaEvaluationHooks {
    return {
      currentCell: { row: coord.row + 1, column: coord.col + 1 },
      // The resolver returns calculated values: a text "=..." is text, never a formula.
      resolverReturnsValues: true,
      ...(legacy ? { implicitIntersection: true } : {}),
      getUsedRange: (reference) => {
        const target = this.resolveSheet(reference)
        if (!target) return null
        this.ensureSpills(target)
        return this.usedRangeOf(target)
      },
      forEachCellInRange: (reference, bounds, visit) => {
        const target = this.resolveSheet(reference)
        if (!target) return
        this.ensureSpills(target)
        const coordinates = this.sortedCoordinates(target)
        // Binary search the first populated row inside the bounds.
        let low = 0
        let high = coordinates.length
        while (low < high) {
          const middle = (low + high) >> 1
          if (coordinates[middle][0] < bounds.startRow) low = middle + 1
          else high = middle
        }
        const override = this.override
        let pending = override && override.sheetId === target.id &&
          override.row + 1 >= bounds.startRow && override.row + 1 <= bounds.endRow &&
          override.col + 1 >= bounds.startColumn && override.col + 1 <= bounds.endColumn
          ? [override.row + 1, override.col + 1] : null
        for (let index = low; index < coordinates.length; index += 1) {
          const [row, column] = coordinates[index]
          if (row > bounds.endRow) break
          if (column < bounds.startColumn || column > bounds.endColumn) continue
          if (pending && (row > pending[0] || (row === pending[0] && column >= pending[1]))) {
            if (row !== pending[0] || column !== pending[1]) visit(pending[0], pending[1])
            pending = null
          }
          visit(row, column)
        }
        if (pending) visit(pending[0], pending[1])
        // Spilled array members are not stored cells; visit the ones inside the bounds too.
        const prefix = `${target.id}!`
        for (const [key, spill] of this.spills) {
          if (!key.startsWith(prefix)) continue
          const top = Math.max(spill.top + 1, bounds.startRow)
          const bottom = Math.min(spill.bottom + 1, bounds.endRow)
          const left = Math.max(spill.left + 1, bounds.startColumn)
          const right = Math.min(spill.right + 1, bounds.endColumn)
          for (let row = top; row <= bottom; row += 1) {
            for (let column = left; column <= right; column += 1) {
              if (target.cells[cellAddress(row - 1, column - 1)]) continue
              if (override && override.sheetId === target.id && override.row + 1 === row && override.col + 1 === column) continue
              visit(row, column)
            }
          }
        }
      },
      resolveDefinedName: (name, sheetId) => this.definedNameValue(name, sheetId),
      isFormulaCell: (reference, address) => Boolean(this.resolveSheet(reference)?.cells[address.replace(/\$/g, '').toUpperCase()]?.formula),
      getCellFormula: (reference, address) => this.resolveSheet(reference)?.cells[address.replace(/\$/g, '').toUpperCase()]?.formula ?? null,
      trackRange: (reference, bounds) => this.track(reference, bounds),
      markVolatile: () => {
        const frame = this.stack[this.stack.length - 1]
        if (frame) frame.volatile = true
      },
      resolveSpill: (reference, row, column) => {
        const target = this.resolveSheet(reference)
        if (!target) return null
        const address = cellAddress(row - 1, column - 1)
        const anchor = target.cells[address]
        if (!anchor?.formula) return null
        this.valueOf(target, address)
        const spill = this.spills.get(`${target.id}!${address}`)
        return spill ? { startRow: spill.top + 1, endRow: spill.bottom + 1, startColumn: spill.left + 1, endColumn: spill.right + 1 } : null
      },
      resolveTable: (name, reference) => this.tableInfo(name, reference, { row: coord.row + 1, column: coord.col + 1 }),
      isRowHidden: (reference, row) => {
        const target = this.resolveSheet(reference)
        if (!target) return false
        const filtered = this.hiddenSet(target.filteredRows)
        return Boolean(this.hiddenSet(target.hiddenRows)?.has(row) && !filtered?.has(row))
      },
      isRowFiltered: (reference, row) => {
        const target = this.resolveSheet(reference)
        return Boolean(target && this.hiddenSet(target.filteredRows)?.has(row))
      },
      getSheetNames: () => this.sheetNames,
      getSheetName: (reference) => this.resolveSheet(reference)?.name ?? null,
    }
  }

  private track(reference: string, bounds: FormulaRangeBounds) {
    const frame = this.stack[this.stack.length - 1]
    if (!frame) return
    const sheet = this.resolveSheet(reference)
    if (!sheet) return
    frame.rects.push({
      sheetId: sheet.id,
      top: Math.max(0, bounds.startRow - 1),
      left: Math.max(0, bounds.startColumn - 1),
      bottom: Math.min(MAX_ROWS - 1, bounds.endRow - 1),
      right: Math.min(MAX_COLUMNS - 1, bounds.endColumn - 1),
    })
  }

  private evaluateFormulaCell(sheet: SheetData, address: string, cell: CellData): CellScalar {
    const key = `${sheet.id}!${address}`
    if (!this.override?.bypass.has(key)) {
      const cached = this.values.get(key)
      if (cached !== undefined || this.values.has(key)) return cached ?? null
    }
    const visitingIndex = this.stackIndex.get(key)
    if (visitingIndex !== undefined) {
      // A speculative spill probe that loops back outside itself must not cache anything.
      if (this.speculativeRoot >= 0 && visitingIndex < this.speculativeRoot) {
        this.discardFrom = Math.min(this.discardFrom, this.speculativeRoot)
        return '#CIRC!'
      }
      // A real circular reference: every formula from the re-entered cell up the stack is on it.
      if (!this.override) {
        for (let index = visitingIndex; index < this.stack.length; index += 1) {
          this.circular.add(this.stack[index].key)
          if (this.options.iterate) this.iterationPending.add(this.stack[index].key)
        }
      }
      if (!this.options.iterate) return '#CIRC!'
      // Iterative calculation: the cycle reads the value from the previous iteration.
      if (this.iterationValues.has(key)) return this.iterationValues.get(key) ?? null
      return typeof cell.result === 'number' ? cell.result : 0
    }
    const coord = parseCellAddress(address)
    if (!coord) return '#REF!'
    const frame: Frame = { key, rects: [], volatile: false, visibility: isVisibilitySensitive(cell.formula!), discard: false }
    this.stackIndex.set(key, this.stack.length)
    this.stack.push(frame)
    let result
    try {
      result = evaluateFormulaDetailed(cell.formula!, sheet.id, this.resolver, this.hooksFor(sheet, coord, isLegacyFormula(cell)))
    } catch {
      result = { value: '#ERROR!' as FormulaResult }
    } finally {
      this.stack.pop()
      this.stackIndex.delete(key)
    }
    const depth = this.stack.length
    let value: CellScalar = result.value
    let stale = false
    if (FALLBACK_ERRORS.has(String(value)) && cell.result !== undefined && cell.result !== null) {
      stale = !Object.is(cell.result, value)
      value = cell.result
    }
    if (depth >= this.discardFrom) {
      if (depth === this.discardFrom) this.discardFrom = Infinity
      return value
    }
    if (this.override) return value

    this.removeDependencies(key)
    this.addDependencies(key, frame.rects)
    if (frame.volatile) this.volatileKeys.add(key)
    if (frame.visibility) this.visibilityKeys.add(key)
    if (stale) this.stale.add(key)
    if (result.arrayEvaluation) this.arrayEvaluated.add(key)
    else this.arrayEvaluated.delete(key)
    this.clearSpill(key)
    this.spillAttempts.delete(key)

    if (result.array && !stale) {
      const { rowCount, columnCount, values } = result.array
      const region: Rect = {
        sheetId: sheet.id,
        top: coord.row,
        left: coord.col,
        bottom: coord.row + rowCount - 1,
        right: coord.col + columnCount - 1,
      }
      this.spillAttempts.set(key, region)
      if (this.spillBlocked(sheet, address, key, region)) {
        value = '#SPILL!'
      } else {
        this.spills.set(key, { top: region.top, left: region.left, bottom: region.bottom, right: region.right, columns: columnCount, values })
        for (let row = region.top; row <= region.bottom; row += 1) {
          for (let col = region.left; col <= region.right; col += 1) {
            if (row === coord.row && col === coord.col) continue
            this.spillCover.set(`${sheet.id}!${cellAddress(row, col)}`, key)
          }
        }
        // Whole-column/row references (A:A, 1:1) must reach the spilled cells.
        const used = this.usedRanges.get(sheet.id)
        if (used) {
          used.maxRow = Math.max(used.maxRow, region.bottom + 1)
          used.maxCol = Math.max(used.maxCol, region.right + 1)
        }
        // Formulas that read these cells before the array spilled here saw blanks.
        this.propagate([region])
      }
    }
    this.values.set(key, value)
    return value
  }

  private spillBlocked(sheet: SheetData, anchorAddress: string, key: Key, region: Rect) {
    if (region.bottom >= MAX_ROWS || region.right >= MAX_COLUMNS) return true
    for (const merge of this.mergesOf(sheet)) {
      if (merge.top <= region.bottom && region.top <= merge.bottom && merge.left <= region.right && region.left <= merge.right) return true
    }
    for (let row = region.top; row <= region.bottom; row += 1) {
      for (let col = region.left; col <= region.right; col += 1) {
        const address = cellAddress(row, col)
        if (address === anchorAddress) continue
        const cell = sheet.cells[address]
        if (cell && hasStoredValue(cell) && cell.arrayMember !== anchorAddress) return true
        const owner = this.spillCover.get(`${sheet.id}!${address}`)
        if (owner && owner !== key) return true
      }
    }
    return false
  }

  /** Make sure every potential spill anchor on the sheet has been evaluated. */
  private ensureSpills(sheet: SheetData) {
    if (this.spillReady.has(sheet.id)) return
    const candidates = this.candidateSet(sheet)
    let complete = true
    for (const address of candidates) {
      const key = `${sheet.id}!${address}`
      if (this.values.has(key)) continue
      if (this.stackIndex.has(key)) {
        complete = false
        continue
      }
      const cell = sheet.cells[address]
      if (!cell?.formula) continue
      const previousRoot = this.speculativeRoot
      if (this.stack.length && this.speculativeRoot < 0) this.speculativeRoot = this.stack.length
      try {
        this.evaluateFormulaCell(sheet, address, cell)
      } finally {
        this.speculativeRoot = previousRoot
      }
      if (!this.values.has(key)) complete = false
    }
    if (complete && !this.stack.length) this.spillReady.add(sheet.id)
  }

  private valueOf(sheet: SheetData, address: string): CellScalar {
    if (this.override && this.override.sheetId === sheet.id && this.override.key === `${sheet.id}!${address}`) return this.override.value
    const cell = sheet.cells[address]
    if (cell?.formula) return this.evaluateFormulaCell(sheet, address, cell)
    if (cell && hasStoredValue(cell)) return cell.value ?? null
    this.ensureSpills(sheet)
    const anchorKey = this.spillCover.get(`${sheet.id}!${address}`)
    if (anchorKey) {
      const spill = this.spills.get(anchorKey)
      const coord = parseCellAddress(address)
      if (spill && coord) return spill.values[(coord.row - spill.top) * spill.columns + (coord.col - spill.left)] ?? null
    }
    if (cell?.arrayMember) {
      // The saved array range no longer spills here; only an anchor we cannot calculate
      // (kept at its cached result) still shows the file's cached member values.
      const anchorKey = `${sheet.id}!${cell.arrayMember}`
      const anchor = sheet.cells[cell.arrayMember]
      if (anchor?.formula) {
        this.evaluateFormulaCell(sheet, cell.arrayMember, anchor)
        if (this.stale.has(anchorKey)) return cell.value ?? null
      }
      return null
    }
    return cell?.value ?? null
  }

  // ---- Public API ---------------------------------------------------------------------------

  /** Calculated value of any cell (formulas, constants, and spilled array members). */
  getValue(sheetId: string, address: string): CellScalar {
    const sheet = this.sheetById.get(sheetId)
    if (!sheet) return '#REF!'
    // Settle every spill before answering, so no value is computed against a spill that
    // is about to appear later in the same pass.
    if (!this.stack.length && this.spillReady.size < this.workbook.sheets.length) {
      for (const candidate of this.workbook.sheets) this.ensureSpills(candidate)
    }
    const value = this.valueOf(sheet, address)
    if (!this.iterationPending.size || this.stack.length || this.iterating) return value
    this.settleIterations()
    return this.valueOf(sheet, address)
  }

  /**
   * Iterative calculation (Excel's "Enable iterative calculation"): recalculate the cells of
   * the circular references found by the last evaluation, each pass reading the previous
   * pass's values, until no value changes by more than maxChange or maxIterations is reached.
   */
  private settleIterations() {
    if (this.iterating || this.stack.length || !this.iterationPending.size) return
    this.iterating = true
    try {
      const members = new Set(this.iterationPending)
      this.iterationPending.clear()
      // The evaluation that found the cycle was the first pass.
      for (let pass = 1; pass < this.options.maxIterations; pass += 1) {
        const before = new Map<Key, CellScalar>()
        for (const key of members) {
          const value = this.values.get(key) ?? null
          before.set(key, value)
          this.iterationValues.set(key, value)
        }
        const rects: Rect[] = []
        for (const key of members) this.invalidateFormula(key, rects)
        this.propagate(rects)
        let change = 0
        for (const key of members) change = Math.max(change, iterationChange(before.get(key), this.valueAtKey(key)))
        for (const key of this.iterationPending) members.add(key)
        this.iterationPending.clear()
        if (change <= this.options.maxChange) break
      }
      for (const key of members) {
        this.iterationValues.set(key, this.values.get(key) ?? null)
        // Circular cells take part in every recalculation, as in Excel.
        if (this.values.has(key)) this.volatileKeys.add(key)
      }
    } finally {
      this.iterating = false
    }
  }

  private valueAtKey(key: Key): CellScalar {
    const separator = key.indexOf('!')
    const sheet = this.sheetById.get(key.slice(0, separator))
    return sheet ? this.valueOf(sheet, key.slice(separator + 1)) : null
  }

  /**
   * Evaluate a formula as though it were entered at (row, col) on a sheet, without storing
   * it — for conditional formatting, data validation, and the Name Manager.
   */
  evaluateAt(sheetId: string, formula: string, row = 0, col = 0, override?: { row: number; col: number; value: CellScalar }): CellScalar {
    const sheet = this.sheetById.get(sheetId)
    if (!sheet) return '#REF!'
    if (!this.stack.length && this.spillReady.size < this.workbook.sheets.length) {
      for (const candidate of this.workbook.sheets) this.ensureSpills(candidate)
    }
    const previous = this.override
    if (override && !this.stack.length) {
      this.override = {
        key: `${sheet.id}!${cellAddress(override.row, override.col)}`,
        sheetId: sheet.id,
        row: override.row,
        col: override.col,
        value: override.value,
        bypass: this.transitiveDependents({ sheetId: sheet.id, top: override.row, left: override.col, bottom: override.row, right: override.col }),
      }
    }
    try {
      return evaluateFormulaDetailed(formula, sheet.id, this.resolver, this.hooksFor(sheet, { row, col })).value
    } catch {
      return '#VALUE!'
    } finally {
      this.override = previous
    }
  }

  /** Every calculated formula that (transitively) reads a rectangle, capped for safety. */
  private transitiveDependents(rect: Rect, limit = 20_000): Set<Key> {
    const seen = new Set<Key>()
    const queue: Rect[] = [rect]
    while (queue.length && seen.size < limit) {
      for (const key of this.dependentsOfRect(queue.pop()!)) {
        if (seen.has(key)) continue
        seen.add(key)
        const separator = key.indexOf('!')
        const coord = parseCellAddress(key.slice(separator + 1))
        if (!coord) continue
        const spill = this.spills.get(key)
        queue.push(spill
          ? { sheetId: key.slice(0, separator), top: spill.top, left: spill.left, bottom: spill.bottom, right: spill.right }
          : { sheetId: key.slice(0, separator), top: coord.row, left: coord.col, bottom: coord.row, right: coord.col })
      }
    }
    return seen
  }

  /** Whether the displayed value is the file's cached result of a formula we cannot calculate. */
  isStale(sheetId: string, address: string) {
    return this.stale.has(`${sheetId}!${address}`)
  }

  /**
   * Why a formula (a cell's, or `formula` as if entered on the sheet) cannot be calculated as
   * written: a syntax problem with its position and a repaired suggestion, or an unknown
   * function or name. Null when it is fine.
   */
  diagnose(sheetId: string, address: string, formula?: string): FormulaDiagnostic | null {
    const text = formula ?? this.sheetById.get(sheetId)?.cells[address]?.formula
    if (!text) return null
    return diagnoseFormula(text.startsWith('=') ? text : `=${text}`, {
      isDefinedName: (name) => this.definedNameValue(name, sheetId) !== undefined,
    })
  }

  /** Anchor address of the dynamic array covering a cell, if any (including the anchor itself). */
  spillAnchorOf(sheetId: string, address: string): string | null {
    const key = `${sheetId}!${address}`
    if (this.spills.has(key)) return address
    const anchor = this.spillCover.get(key)
    return anchor ? anchor.slice(anchor.indexOf('!') + 1) : null
  }

  /** Active spill range of an anchor as {top,left,bottom,right} (0-based), or null. */
  spillRange(sheetId: string, anchorAddress: string) {
    const sheet = this.sheetById.get(sheetId)
    if (sheet?.cells[anchorAddress]?.formula) this.valueOf(sheet, anchorAddress)
    const spill = this.spills.get(`${sheetId}!${anchorAddress}`)
    return spill ? { top: spill.top, left: spill.left, bottom: spill.bottom, right: spill.right } : null
  }

  /** Every active spill on a sheet (for drawing spill borders). */
  spillsOnSheet(sheetId: string) {
    const sheet = this.sheetById.get(sheetId)
    if (sheet) this.ensureSpills(sheet)
    const output: Array<{ anchor: string; top: number; left: number; bottom: number; right: number }> = []
    for (const [key, spill] of this.spills) {
      if (!key.startsWith(`${sheetId}!`)) continue
      output.push({ anchor: key.slice(sheetId.length + 1), top: spill.top, left: spill.left, bottom: spill.bottom, right: spill.right })
    }
    return output
  }

  /** Precedent rectangles (0-based) of a formula cell, after evaluating it. */
  precedentsOf(sheetId: string, address: string) {
    this.getValue(sheetId, address)
    return (this.deps.get(`${sheetId}!${address}`) || []).map((rect) => ({ ...rect }))
  }

  /** Formula cells that directly depend on a cell (0-based coordinates). */
  dependentsOf(sheetId: string, address: string) {
    const coord = parseCellAddress(address)
    if (!coord) return []
    this.calculateAll()
    const target: Rect = { sheetId, top: coord.row, left: coord.col, bottom: coord.row, right: coord.col }
    return this.dependentsOfRect(target).map((key) => {
      const separator = key.indexOf('!')
      return { sheetId: key.slice(0, separator), address: key.slice(separator + 1) }
    })
  }

  /** Evaluate every formula in the workbook (before saving, printing, or tracing). */
  calculateAll() {
    for (const sheet of this.workbook.sheets) {
      for (const address in sheet.cells) {
        const cell = sheet.cells[address]
        if (cell?.formula) this.evaluateFormulaCell(sheet, address, cell)
      }
    }
    this.settleIterations()
  }

  /**
   * A copy of the workbook with formula results, spill metadata, and spilled member values
   * written into the cells, ready for saving/exporting. Unchanged sheets keep identity.
   */
  withResults(options: { forSave?: boolean } = {}): WorkbookModel {
    // Saving or exporting recalculates a manual-mode workbook first, as Excel does ("Recalculate
    // workbook before saving", <calcPr calcOnSave>, on by default), so files never carry stale
    // results. Printing and previews (forSave false) show the values on screen.
    const calcOnSave = this.workbook.metadata?.calcProperties?.calcOnSave
    const recalculateOnSave = !(calcOnSave !== undefined && calcOnSave !== null && calcOnSave !== '' && !flagValue(calcOnSave))
    if (options.forSave !== false && this.needsRecalculation && recalculateOnSave) this.recalculate()
    this.calculateAll()
    let changed = false
    const sheets = this.workbook.sheets.map((sheet) => {
      let cells: SheetData['cells'] | null = null
      const write = (address: string, cell: CellData | null) => {
        if (!cells) cells = { ...sheet.cells }
        if (cell) cells[address] = cell
        else delete cells[address]
      }
      for (const address in sheet.cells) {
        const cell = sheet.cells[address]
        if (!cell) continue
        const key = `${sheet.id}!${address}`
        if (cell.formula) {
          if (this.stale.has(key)) continue
          const raw = this.values.get(key) ?? null
          // App-internal display markers (U+E000..., e.g. SPARKLINE) are saved as an empty result.
          const value = typeof raw === 'string' && raw.charCodeAt(0) === 0xe000 ? '' : raw
          const spill = this.spills.get(key)
          const next: CellData = { ...cell }
          let dirty = !Object.is(cell.result, value)
          next.result = value
          if (isFormulaError(value) || value === '#ERROR!') next.resultType = 'error'
          else delete next.resultType
          if (spill) {
            const range = `${cellAddress(spill.top, spill.left)}:${cellAddress(spill.bottom, spill.right)}`
            if (next.formulaRange !== range || next.formulaType !== 'array' || !next.dynamicFormula) dirty = true
            next.formulaType = 'array'
            next.formulaRange = range
            next.dynamicFormula = true
          } else if (this.arrayEvaluated.has(key) && !isLegacyFormula(cell) && cell.formulaType !== 'shared' && (cell.formulaType !== 'array' || cell.dynamicFormula)) {
            // =SUM(A1:A3*2) relies on array evaluation; saved as a plain formula Excel would
            // implicitly intersect A1:A3, so it is saved as a one-cell array formula.
            if (next.formulaRange !== address || next.formulaType !== 'array' || !next.dynamicFormula) dirty = true
            next.formulaType = 'array'
            next.formulaRange = address
            next.dynamicFormula = true
          } else if (cell.dynamicFormula || (cell.formulaType === 'array' && cell.formulaRange && parseRange(cell.formulaRange) && (() => {
            const bounds = parseRange(cell.formulaRange!)!
            return bounds.bottom > bounds.top || bounds.right > bounds.left
          })())) {
            // A dynamic array that no longer spills is saved as an ordinary formula.
            if (cell.dynamicFormula) {
              delete next.formulaType
              delete next.formulaRange
              delete next.dynamicFormula
              dirty = true
            }
          }
          if (dirty) {
            delete next.display
            write(address, next)
          }
        } else if (cell.arrayMember) {
          const anchorKey = `${sheet.id}!${cell.arrayMember}`
          if (this.stale.has(anchorKey)) continue
          if (this.spillCover.get(key) !== anchorKey) {
            const next: CellData = { ...cell }
            delete next.value
            delete next.arrayMember
            delete next.display
            write(address, next.style || next.numFmt || next.note ? next : null)
          }
        }
      }
      // Spilled members: cached values Excel expects in the covered cells.
      for (const [key, spill] of this.spills) {
        if (!key.startsWith(`${sheet.id}!`)) continue
        const anchorAddress = key.slice(sheet.id.length + 1)
        for (let row = spill.top; row <= spill.bottom; row += 1) {
          for (let col = spill.left; col <= spill.right; col += 1) {
            const address = cellAddress(row, col)
            if (address === anchorAddress) continue
            const value = spill.values[(row - spill.top) * spill.columns + (col - spill.left)] ?? null
            const existing = sheet.cells[address]
            if (existing?.arrayMember === anchorAddress && Object.is(existing.value, value)) continue
            const next: CellData = { ...(existing || {}), value: value as CellScalar, arrayMember: anchorAddress }
            if (isFormulaError(value)) next.type = 'error'
            else if (next.type === 'error') delete next.type
            delete next.display
            write(address, next)
          }
        }
      }
      if (!cells) return sheet
      changed = true
      return { ...sheet, cells: cells as SheetData['cells'] }
    })
    return changed ? { ...this.workbook, sheets } : this.workbook
  }
}

/**
 * Mark the cells covered by saved array/dynamic formula ranges so that the live spill can
 * replace their cached values. Returns the same sheet when nothing needed marking.
 */
export function markArrayMembers(sheet: SheetData): SheetData {
  let cells: SheetData['cells'] | null = null
  for (const [address, cell] of Object.entries(sheet.cells)) {
    if (!cell?.formula || !cell.formulaRange || (cell.formulaType !== 'array' && !cell.dynamicFormula)) continue
    const bounds = parseRange(cell.formulaRange)
    if (!bounds || (bounds.top === bounds.bottom && bounds.left === bounds.right)) continue
    if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > 1_000_000) continue
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        const member = cellAddress(row, col)
        if (member === address) continue
        const existing = (cells || sheet.cells)[member]
        if (existing?.formula || existing?.arrayMember === address) continue
        if (!cells) cells = { ...sheet.cells }
        cells[member] = { ...(existing || {}), arrayMember: address }
      }
    }
  }
  return cells ? { ...sheet, cells } : sheet
}
