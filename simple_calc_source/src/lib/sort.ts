/**
 * Multi-level range sort (pure) with Excel ordering: numbers < text < logicals < errors, blanks
 * always last in either direction, stable, optional case sensitivity, custom lists, sort by fill
 * or font colour, and left-to-right orientation. The app applies the resulting cell changes.
 */
import type { CellData, SheetData } from '../spreadsheet-types'
import { currentRegionAround } from './filter'
import {
  addressOf,
  boundsContain,
  boundsIntersect,
  cellHasContent,
  collationKeys,
  columnLabel,
  compareCollated,
  foldText,
  hostFillColor,
  hostFontColor,
  moveCell,
  normalizeColor,
  parseAddress,
  parseRange,
  scalarKind,
} from './data-tools-core'
import type { Bounds, CellChanges, DataHost, Scalar, ShiftFormula } from './data-tools-core'

export type SortOn = 'values' | 'fillColor' | 'fontColor'
/** 'rows' sorts top to bottom (keys are columns); 'columns' sorts left to right (keys are rows). */
export type SortOrientation = 'rows' | 'columns'

export interface SortLevel {
  /** 0-based offset within the range of the key column ('rows') or key row ('columns'). */
  key: number
  descending?: boolean
  sortOn?: SortOn
  /** For colour sorts: '#RRGGBB', or null for No Fill / Automatic font colour. */
  color?: string | null
  /** Where lines whose key matches `color` go (default 'top'). */
  position?: 'top' | 'bottom'
  /** Custom order such as Jan..Dec; list items sort first (reversed when `descending`). */
  customList?: string[]
}

export interface SortSpec {
  bounds: Bounds
  levels: SortLevel[]
  hasHeader?: boolean
  orientation?: SortOrientation
  caseSensitive?: boolean
  /** Absolute 0-based rows ('rows') or columns ('columns') that stay in place: hidden or filtered lines. */
  fixed?: Iterable<number> | ((line: number) => boolean)
}

export interface SortOrder {
  ok: true
  /** First data line (absolute row or column index). */
  start: number
  /** order[i] is the absolute source line that lands on line start + i. */
  order: number[]
  changed: boolean
}

export interface SortFailure { ok: false; error: string }

export const BUILTIN_CUSTOM_LISTS: Array<{ id: string; label: string; items: string[] }> = [
  { id: 'days-short', label: 'Sun, Mon, Tue, Wed, Thu, Fri, Sat', items: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'] },
  { id: 'days-long', label: 'Sunday, Monday, Tuesday, …', items: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'] },
  { id: 'months-short', label: 'Jan, Feb, Mar, Apr, …', items: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'] },
  { id: 'months-long', label: 'January, February, March, …', items: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'] },
]

const MAX_SORT_CELLS = 5_000_000

// Type ranks. LIST sorts before everything; BLANK is special-cased to stay last.
const LIST = 0
const NUMBER = 1
const TEXT = 2
const LOGICAL = 3
const ERROR = 4
const BLANK = 5

interface LevelKeys {
  rank: Uint8Array
  value: Float64Array
  descending: boolean
  custom: boolean
  color: boolean
}

function keyCell(spec: SortSpec, line: number, key: number): { row: number; col: number } {
  return (spec.orientation || 'rows') === 'rows'
    ? { row: line, col: spec.bounds.left + key }
    : { row: spec.bounds.top + key, col: line }
}

function lineRange(spec: SortSpec): { start: number; end: number } {
  const rows = (spec.orientation || 'rows') === 'rows'
  const first = rows ? spec.bounds.top : spec.bounds.left
  const last = rows ? spec.bounds.bottom : spec.bounds.right
  return { start: spec.hasHeader ? first + 1 : first, end: last }
}

function buildLevelKeys(spec: SortSpec, level: SortLevel, start: number, count: number, host: DataHost): LevelKeys {
  const rank = new Uint8Array(count)
  const value = new Float64Array(count)
  const sortOn = level.sortOn || 'values'
  if (sortOn !== 'values') {
    const wanted = normalizeColor(level.color)
    const matchRank = level.position === 'bottom' ? 1 : 0
    for (let index = 0; index < count; index += 1) {
      const { row, col } = keyCell(spec, start + index, level.key)
      const color = sortOn === 'fillColor' ? hostFillColor(host, row, col) : hostFontColor(host, row, col)
      value[index] = color === wanted ? matchRank : 1 - matchRank
      rank[index] = NUMBER
    }
    return { rank, value, descending: false, custom: false, color: true }
  }
  const listIndex = new Map<string, number>()
  if (level.customList?.length) {
    const items = level.descending ? [...level.customList].reverse() : level.customList
    items.forEach((item, index) => {
      const key = foldText(String(item).trim())
      if (!listIndex.has(key)) listIndex.set(key, index)
    })
  }
  const texts: Array<string | undefined> = new Array(count)
  let hasText = false
  for (let index = 0; index < count; index += 1) {
    const { row, col } = keyCell(spec, start + index, level.key)
    const cellValue: Scalar | undefined = host.valueAt(row, col)
    const kind = scalarKind(cellValue)
    if (kind === 'blank') { rank[index] = BLANK; continue }
    if (kind === 'number') { rank[index] = NUMBER; value[index] = cellValue as number }
    else if (kind === 'boolean') { rank[index] = LOGICAL; value[index] = cellValue ? 1 : 0 }
    else if (kind === 'error') { rank[index] = ERROR }
    else {
      const text = String(cellValue)
      if (text.trim() === '') { rank[index] = BLANK; continue }
      rank[index] = TEXT
      texts[index] = text
      hasText = true
    }
    if (listIndex.size) {
      const position = listIndex.get(foldText(host.displayAt(row, col).trim()))
      if (position !== undefined) { rank[index] = LIST; value[index] = position; texts[index] = undefined }
    }
  }
  if (hasText) {
    // Collate each distinct string once, then sort on integer ranks.
    const unique = [...new Set(texts.filter((text): text is string => text !== undefined))]
    const caseSensitive = Boolean(spec.caseSensitive)
    const collated = collationKeys(unique, caseSensitive)
    const sorted = unique.map((_, index) => index)
    sorted.sort((a, b) => compareCollated(unique[a], collated[a], unique[b], collated[b], caseSensitive))
    const ranks = new Map<string, number>()
    let current = 0
    for (let index = 0; index < sorted.length; index += 1) {
      const item = sorted[index]
      if (index > 0) {
        const previous = sorted[index - 1]
        if (compareCollated(unique[previous], collated[previous], unique[item], collated[item], caseSensitive) !== 0) current += 1
      }
      ranks.set(unique[item], current)
    }
    for (let index = 0; index < count; index += 1) {
      const text = texts[index]
      if (text !== undefined && rank[index] === TEXT) value[index] = ranks.get(text)!
    }
  }
  return { rank, value, descending: Boolean(level.descending) && !listIndex.size, custom: listIndex.size > 0, color: false }
}

function compareLevel(keys: LevelKeys, a: number, b: number): number {
  const ra = keys.rank[a]
  const rb = keys.rank[b]
  if (keys.color) return keys.value[a] - keys.value[b]
  if (ra === BLANK || rb === BLANK) return ra === rb ? 0 : ra === BLANK ? 1 : -1
  if (keys.custom) {
    if ((ra === LIST) !== (rb === LIST)) return ra === LIST ? -1 : 1
    if (ra === LIST) return keys.value[a] - keys.value[b]
  }
  const difference = ra - rb || keys.value[a] - keys.value[b]
  return keys.descending ? -difference : difference
}

function fixedPredicate(fixed: SortSpec['fixed']): (line: number) => boolean {
  if (!fixed) return () => false
  if (typeof fixed === 'function') return fixed
  const set = new Set(fixed)
  return set.size ? (line) => set.has(line) : () => false
}

/** Compute the stable sort order of the range's data lines (header and fixed lines excluded). */
export function computeSortOrder(spec: SortSpec, host: DataHost): SortOrder | SortFailure {
  const { start, end } = lineRange(spec)
  const count = end - start + 1
  if (count <= 0) return { ok: false, error: 'There is no data to sort below the header.' }
  const width = (spec.orientation || 'rows') === 'rows' ? spec.bounds.right - spec.bounds.left + 1 : spec.bounds.bottom - spec.bounds.top + 1
  const levels = spec.levels.filter((level) => Number.isInteger(level.key) && level.key >= 0 && level.key < width)
  if (!levels.length) return { ok: false, error: 'Choose a column to sort by.' }
  const isFixed = fixedPredicate(spec.fixed)
  const keys = levels.map((level) => buildLevelKeys(spec, level, start, count, host))
  const movable: number[] = []
  for (let index = 0; index < count; index += 1) if (!isFixed(start + index)) movable.push(index)
  movable.sort((a, b) => {
    for (let level = 0; level < keys.length; level += 1) {
      const result = compareLevel(keys[level], a, b)
      if (result) return result
    }
    return a - b
  })
  const order = new Array<number>(count)
  let next = 0
  let changed = false
  for (let index = 0; index < count; index += 1) {
    if (isFixed(start + index)) { order[index] = start + index; continue }
    const source = start + movable[next++]
    order[index] = source
    if (source !== start + index) changed = true
  }
  return { ok: true, start, order, changed }
}

/**
 * Why the range cannot be sorted, or null. Mirrors Excel: merged cells must be the same size in
 * every sorted line, and array / spilled formulas cannot be split.
 */
export function sortBlocker(sheet: Pick<SheetData, 'cells' | 'merges'>, spec: SortSpec): string | null {
  const orientation = spec.orientation || 'rows'
  const { bounds } = spec
  const cellCount = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
  if (cellCount > MAX_SORT_CELLS) return 'That range is too large to sort at once.'
  const { start, end } = lineRange(spec)
  if (end < start) return 'Select a range with at least two rows to sort.'
  const dataBounds: Bounds = orientation === 'rows' ? { ...bounds, top: start } : { ...bounds, left: start }
  const isFixed = fixedPredicate(spec.fixed)
  const signatures = new Map<number, string[]>()
  for (const range of sheet.merges || []) {
    const merged = parseRange(range)
    if (!merged || !boundsIntersect(merged, dataBounds)) continue
    if (!boundsContain(dataBounds, merged)) return 'To do this, all the merged cells need to be the same size.'
    const spansLines = orientation === 'rows' ? merged.top !== merged.bottom : merged.left !== merged.right
    if (spansLines) return 'To do this, all the merged cells need to be the same size.'
    const line = orientation === 'rows' ? merged.top : merged.left
    const span = orientation === 'rows' ? `${merged.left}:${merged.right}` : `${merged.top}:${merged.bottom}`
    const list = signatures.get(line) || []
    list.push(span)
    signatures.set(line, list)
  }
  if (signatures.size) {
    let expected: string | null = null
    for (let line = start; line <= end; line += 1) {
      if (isFixed(line)) continue
      const signature = (signatures.get(line) || []).sort().join(',')
      if (expected === null) expected = signature
      else if (signature !== expected) return 'To do this, all the merged cells need to be the same size.'
    }
  }
  for (const address in sheet.cells) {
    const cell = sheet.cells[address]
    const isArray = cell.formulaType === 'array' || cell.dynamicFormula || (cell.formulaRange && cell.formulaType !== 'shared')
    if (!isArray) continue
    const area = (cell.formulaRange && parseRange(cell.formulaRange)) || (parseAddress(address) && rangeOfAddress(address))
    if (area && boundsIntersect(area, dataBounds)) return "You can't sort part of an array formula. Select a range without array or spilled formulas."
  }
  return null
}

function rangeOfAddress(address: string): Bounds | null {
  const coord = parseAddress(address)
  return coord ? { top: coord.row, bottom: coord.row, left: coord.col, right: coord.col } : null
}

/**
 * Cell changes that move whole lines of the range into sorted order. Styles, notes and links
 * travel with their cells; formulas are shifted by the distance moved. Shared-formula groups
 * touching the range are converted to ordinary formulas so no child is orphaned.
 */
export function sortedCellChanges(sheet: Pick<SheetData, 'cells'>, spec: SortSpec, sorted: SortOrder, shiftFormula?: ShiftFormula): CellChanges {
  const changes: CellChanges = {}
  const rows = (spec.orientation || 'rows') === 'rows'
  const { bounds } = spec
  const crossStart = rows ? bounds.left : bounds.top
  const width = (rows ? bounds.right : bounds.bottom) - crossStart + 1
  const count = sorted.order.length
  // One address table and one cell snapshot for the whole range: every address is built and
  // looked up once, then reused as both a source and a target.
  const addresses = new Array<string>(count * width)
  const snapshot = new Array<(CellData & { sharedFormulaMaster?: string }) | undefined>(count * width)
  const masters = new Set<string>()
  for (let index = 0; index < count; index += 1) {
    const line = sorted.start + index
    for (let offset = 0; offset < width; offset += 1) {
      const address = rows ? addressOf(line, crossStart + offset) : addressOf(crossStart + offset, line)
      const cell = sheet.cells[address] as (CellData & { sharedFormulaMaster?: string }) | undefined
      addresses[index * width + offset] = address
      snapshot[index * width + offset] = cell
      if (cell?.formulaType === 'shared') masters.add(cell.sharedFormulaMaster || address)
    }
  }
  if (masters.size) {
    // Shared-formula groups touching the range become ordinary formulas (each cell already
    // holds its own formula text), so moving the master never orphans a child.
    for (const address in sheet.cells) {
      const cell = sheet.cells[address] as CellData & { sharedFormulaMaster?: string }
      if (cell.formulaType !== 'shared' || !masters.has(cell.sharedFormulaMaster || address)) continue
      const copy: CellData & { sharedFormulaMaster?: string } = { ...cell }
      delete copy.formulaType
      delete copy.formulaRange
      delete copy.sharedFormulaMaster
      changes[address] = copy
    }
  }
  for (let index = 0; index < count; index += 1) {
    const source = sorted.order[index] - sorted.start
    if (source === index) continue
    const delta = index - source
    for (let offset = 0; offset < width; offset += 1) {
      const target = addresses[index * width + offset]
      const cell = snapshot[source * width + offset]
      if (!cell) {
        if (snapshot[index * width + offset] || target in changes) changes[target] = null
        continue
      }
      changes[target] = rows ? moveCell(cell, delta, 0, shiftFormula) : moveCell(cell, 0, delta, shiftFormula)
    }
  }
  return changes
}

export type SortRangeResult =
  | { ok: true; changes: CellChanges; changed: boolean; order: SortOrder; dataBounds: Bounds }
  | SortFailure

/** Validate, order and build the cell changes in one call. */
export function sortRange(sheet: Pick<SheetData, 'cells' | 'merges'>, spec: SortSpec, host: DataHost, shiftFormula?: ShiftFormula): SortRangeResult {
  const blocker = sortBlocker(sheet, spec)
  if (blocker) return { ok: false, error: blocker }
  const order = computeSortOrder(spec, host)
  if (!order.ok) return order
  const rows = (spec.orientation || 'rows') === 'rows'
  const dataBounds: Bounds = rows ? { ...spec.bounds, top: order.start } : { ...spec.bounds, left: order.start }
  if (!order.changed) return { ok: true, changes: {}, changed: false, order, dataBounds }
  return { ok: true, changes: sortedCellChanges(sheet, spec, order, shiftFormula), changed: true, order, dataBounds }
}

// ---------------------------------------------------------------------------------------------
// Header detection and dialog helpers
// ---------------------------------------------------------------------------------------------

function isEmphasised(cell: CellData | undefined): string {
  const style = cell?.style
  if (!style) return ''
  return [
    style.font?.bold ? 'b' : '',
    style.font?.italic ? 'i' : '',
    style.font?.underline ? 'u' : '',
    style.fill && String(style.fill.pattern || '').toLocaleLowerCase() !== 'none' && (style.fill.fgColor || style.fill.color) ? 'f' : '',
    style.border?.bottom?.style ? 'l' : '',
  ].join('')
}

/**
 * Excel's "My data has headers" guess: the first line is all text while the lines beneath hold
 * numbers/dates/logicals in the same column, or the first line is formatted differently.
 */
export function detectHeaderRow(bounds: Bounds, host: DataHost, orientation: SortOrientation = 'rows'): boolean {
  const rows = orientation === 'rows'
  const lines = rows ? bounds.bottom - bounds.top + 1 : bounds.right - bounds.left + 1
  if (lines < 2) return false
  const crossStart = rows ? bounds.left : bounds.top
  const crossEnd = rows ? bounds.right : bounds.bottom
  const first = rows ? bounds.top : bounds.left
  const coordinate = (line: number, cross: number) => rows ? { row: line, col: cross } : { row: cross, col: line }
  const sample = Math.min(lines - 1, 20)
  let headerCells = 0
  let typedEvidence = false
  let formatEvidence = 0
  let formatChecks = 0
  for (let cross = crossStart; cross <= crossEnd; cross += 1) {
    const head = coordinate(first, cross)
    const headKind = scalarKind(host.valueAt(head.row, head.col))
    if (headKind === 'blank') continue
    if (headKind !== 'text') return false
    headerCells += 1
    let typed = 0
    let filled = 0
    for (let offset = 1; offset <= sample; offset += 1) {
      const cell = coordinate(first + offset, cross)
      const kind = scalarKind(host.valueAt(cell.row, cell.col))
      if (kind === 'blank') continue
      filled += 1
      if (kind === 'number' || kind === 'boolean') typed += 1
    }
    if (filled && typed / filled >= 0.5) typedEvidence = true
    if (host.cellAt) {
      const below = coordinate(first + 1, cross)
      const headFormat = isEmphasised(host.cellAt(head.row, head.col))
      const belowFormat = isEmphasised(host.cellAt(below.row, below.col))
      formatChecks += 1
      if (headFormat !== belowFormat && headFormat.length >= belowFormat.length) formatEvidence += 1
    }
  }
  if (!headerCells) return false
  return typedEvidence || (formatChecks > 0 && formatEvidence / formatChecks >= 0.5)
}

/** Labels for the Sort dialog's "Sort by" list: header text, or "Column B" / "Row 4". */
export function sortKeyLabels(bounds: Bounds, orientation: SortOrientation, hasHeader: boolean, host: DataHost): string[] {
  const rows = orientation === 'rows'
  const count = rows ? bounds.right - bounds.left + 1 : bounds.bottom - bounds.top + 1
  const labels: string[] = []
  for (let offset = 0; offset < count; offset += 1) {
    const fallback = rows ? `Column ${columnLabel(bounds.left + offset)}` : `Row ${bounds.top + offset + 1}`
    if (!hasHeader) { labels.push(fallback); continue }
    const text = rows ? host.displayAt(bounds.top, bounds.left + offset) : host.displayAt(bounds.top + offset, bounds.left)
    labels.push(text.trim() ? text.trim() : `(${fallback})`)
  }
  return labels
}

/** Distinct fill or font colours in a key line (for "Sort On: Cell Color" order choices). */
export function sortKeyColors(bounds: Bounds, orientation: SortOrientation, key: number, sortOn: Exclude<SortOn, 'values'>, hasHeader: boolean, host: DataHost): Array<string | null> {
  const spec: SortSpec = { bounds, orientation, hasHeader, levels: [] }
  const { start, end } = lineRange(spec)
  const seen = new Set<string | null>()
  for (let line = start; line <= end && seen.size < 64; line += 1) {
    const { row, col } = keyCell(spec, line, key)
    seen.add(sortOn === 'fillColor' ? hostFillColor(host, row, col) : hostFontColor(host, row, col))
  }
  return [...seen].sort((a, b) => (a === null ? 1 : 0) - (b === null ? 1 : 0))
}

/**
 * Excel's Sort Warning ("found data next to your selection"): sorting a one-column or partial
 * selection that has data right beside it would separate that column from the rest of its rows.
 * Returns the region to offer instead (the current region around the selection, Excel's
 * "Expand the selection") when it reaches beyond the selection's columns; null when the
 * selection can be sorted as it is (a single cell expands on its own).
 */
export function sortExpansionRegion(sheet: Pick<SheetData, 'cells'>, selection: Bounds): Bounds | null {
  if (selection.top === selection.bottom && selection.left === selection.right) return null
  // A whole-column selection is clipped to the data first, so the region is the data block.
  let lastRow = -1
  for (const address in sheet.cells) {
    if (!cellHasContent(sheet.cells[address])) continue
    const coord = parseAddress(address)
    if (coord && coord.col >= selection.left - 1 && coord.col <= selection.right + 1) lastRow = Math.max(lastRow, coord.row)
  }
  if (lastRow < selection.top) return null
  const clipped = { ...selection, bottom: Math.min(selection.bottom, lastRow) }
  const region = currentRegionAround(sheet, clipped)
  if (region.left >= selection.left && region.right <= selection.right) return null
  return { ...region, top: Math.min(region.top, selection.top), bottom: Math.max(region.bottom, clipped.bottom) }
}

/**
 * Spec for sorting an AutoFilter range from its column menu: the header row stays put and
 * `fixedRows` (1-based, e.g. hidden rows the sort must not move) keep their positions.
 */
export function sortSpecForFilter(filterRef: string, level: SortLevel, options: { fixedRows?: number[]; caseSensitive?: boolean } = {}): SortSpec | null {
  const bounds = parseRange(filterRef)
  if (!bounds || bounds.bottom <= bounds.top) return null
  const fixed = options.fixedRows?.length ? new Set(options.fixedRows.map((row) => row - 1)) : undefined
  return { bounds, hasHeader: true, levels: [level], caseSensitive: options.caseSensitive, fixed }
}
