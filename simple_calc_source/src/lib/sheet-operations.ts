import type {
  CellData,
  DefinedName,
  SheetData,
  SheetFilterState,
  WorkbookModel,
} from '../spreadsheet-types'
import {
  columnLabelToNumber,
  columnNumberToLabel,
  formatA1Address,
  parseA1Address,
  shiftFormulaReferences,
} from './formulas'
import { renameChartSheetReferences } from './charts'
import { adjustFilterForStructure, adjustRowNumbersForStructure, filterIsActive } from './filter'
import { moveReferencesInFormula, renameSheetInFormula, tokenizeFormulaText } from './formula-editing'
import { renameTablesInCopiedSheet, tableRegions, transformTablesForStructure } from './tables'
import { withValidation } from './validation'
import type { DataValidationModel } from './validation'

export const MAX_SHEET_ROWS = 1_048_576
export const MAX_SHEET_COLUMNS = 16_384

export type SheetAxis = 'row' | 'column'
export type SheetStructureKind = 'insert' | 'delete'

/** Zero-based structural edit coordinates, matching the grid selection model. */
export interface SheetStructureOperation {
  axis: SheetAxis
  kind: SheetStructureKind
  index: number
  count: number
  /**
   * Insert/delete cells rather than whole rows or columns: only this 0-based band of the other
   * axis shifts (columns for row operations, rows for column operations), as Excel's
   * "Shift cells down/right/up/left" does.
   */
  span?: { start: number; end: number }
}

export type CellShiftDirection = 'down' | 'right' | 'up' | 'left'

/** Whether a 1-based perpendicular extent lies outside the operation's band. */
function outsideSpan(operation: SheetStructureOperation, low: number, high: number) {
  return Boolean(operation.span) && (low < operation.span!.start + 1 || high > operation.span!.end + 1)
}

export interface GridCoordinate {
  row: number
  col: number
}

export interface GridSelection {
  anchor: GridCoordinate
  focus: GridCoordinate
}

export type SelectionStructureCommand =
  | 'insert-rows-above'
  | 'insert-rows-below'
  | 'delete-rows'
  | 'insert-columns-left'
  | 'insert-columns-right'
  | 'delete-columns'

export interface SelectionStructureResult {
  workbook: WorkbookModel
  selection: GridSelection
  operation: SheetStructureOperation
}

export interface InsertBlankSheetOptions {
  /** Insert after this sheet. When omitted, append to the workbook. */
  afterSheetId?: string
  /** Optional explicit name. The default is the first unused `SheetN`. */
  name?: string
  /** Optional explicit id. The default is a deterministic unused `sheet-N`. */
  id?: string
  rowCount?: number
  colCount?: number
  activate?: boolean
}

export interface InsertBlankSheetResult {
  workbook: WorkbookModel
  sheetId: string
}

type Bounds = { top: number; bottom: number; left: number; right: number }
type Interval = { start: number; end: number }

interface ParsedReference {
  kind: 'cell' | 'cell-range' | 'column-range' | 'row-range'
  first: string
  second?: string
}

interface FormulaReferenceMatch {
  length: number
  prefix: string
  reference: ParsedReference
}

export class SheetStructureError extends Error {
  constructor(
    message: string,
    readonly code:
      | 'INVALID_OPERATION'
      | 'SHEET_NOT_FOUND'
      | 'LIMIT_EXCEEDED'
      | 'ARRAY_RANGE_CONFLICT'
      | 'SHIFT_CONFLICT',
  ) {
    super(message)
    this.name = 'SheetStructureError'
  }
}

function clone<T>(value: T): T {
  return structuredClone(value) as T
}

function maxForAxis(axis: SheetAxis): number {
  return axis === 'row' ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS
}

function validateOperation(operation: SheetStructureOperation): void {
  const { axis, kind, index, count } = operation
  if (axis !== 'row' && axis !== 'column') {
    throw new SheetStructureError('The structural edit axis must be a row or column.', 'INVALID_OPERATION')
  }
  if (kind !== 'insert' && kind !== 'delete') {
    throw new SheetStructureError('The structural edit must insert or delete.', 'INVALID_OPERATION')
  }
  if (!Number.isSafeInteger(index) || index < 0 || index >= maxForAxis(axis)) {
    throw new SheetStructureError('The structural edit starts outside the worksheet.', 'INVALID_OPERATION')
  }
  if (!Number.isSafeInteger(count) || count <= 0) {
    throw new SheetStructureError('The structural edit count must be a positive integer.', 'INVALID_OPERATION')
  }
  if (index + count > maxForAxis(axis)) {
    throw new SheetStructureError('The edited row or column range would exceed the XLSX worksheet limit.', 'LIMIT_EXCEEDED')
  }
}

function normalizedBounds(selection: GridSelection): Bounds {
  const values = [selection.anchor.row, selection.anchor.col, selection.focus.row, selection.focus.col]
  if (!values.every((value) => Number.isSafeInteger(value) && value >= 0)) {
    throw new SheetStructureError('The selection contains an invalid coordinate.', 'INVALID_OPERATION')
  }
  const bounds = {
    top: Math.min(selection.anchor.row, selection.focus.row),
    bottom: Math.max(selection.anchor.row, selection.focus.row),
    left: Math.min(selection.anchor.col, selection.focus.col),
    right: Math.max(selection.anchor.col, selection.focus.col),
  }
  if (bounds.bottom >= MAX_SHEET_ROWS || bounds.right >= MAX_SHEET_COLUMNS) {
    throw new SheetStructureError('The selection is outside the worksheet.', 'INVALID_OPERATION')
  }
  return bounds
}

function coordinateFromAddress(address: string): GridCoordinate | null {
  const parsed = parseA1Address(address)
  return parsed ? { row: parsed.row - 1, col: parsed.column - 1 } : null
}

function addressFromCoordinate(coordinate: GridCoordinate): string {
  const address = formatA1Address({
    row: coordinate.row + 1,
    column: coordinate.col + 1,
    rowAbsolute: false,
    columnAbsolute: false,
  })
  if (!address) throw new SheetStructureError('A worksheet address exceeds the XLSX limit.', 'LIMIT_EXCEEDED')
  return address
}

function operationInterval(operation: SheetStructureOperation): Interval {
  const start = operation.index + 1
  return { start, end: start + operation.count - 1 }
}

/** Map one one-based row or column position. `null` means it was deleted. */
function transformPosition(position: number, operation: SheetStructureOperation): number | null {
  const { start, end } = operationInterval(operation)
  if (operation.kind === 'insert') return position >= start ? position + operation.count : position
  if (position < start) return position
  if (position > end) return position - operation.count
  return null
}

/** Navigation anchors inside deleted space land on the first surviving position. */
function transformNavigationPosition(position: number, operation: SheetStructureOperation): number {
  const transformed = transformPosition(position, operation)
  return transformed ?? Math.min(operation.index + 1, maxForAxis(operation.axis))
}

/**
 * Transform an inclusive interval. Insertions inside a range expand it;
 * deletions shrink it to the surviving cells. Direction is preserved.
 */
function transformInterval(interval: Interval, operation: SheetStructureOperation): Interval | null {
  const ascending = interval.start <= interval.end
  let low = Math.min(interval.start, interval.end)
  let high = Math.max(interval.start, interval.end)
  const { start, end } = operationInterval(operation)

  if (operation.kind === 'insert') {
    if (start <= low) {
      low += operation.count
      high += operation.count
    } else if (start <= high) {
      high += operation.count
    }
  } else if (high < start) {
    // The interval is before the deleted area.
  } else if (low > end) {
    low -= operation.count
    high -= operation.count
  } else {
    const beforeLow = low
    const beforeHigh = Math.min(high, start - 1)
    const afterLow = Math.max(low, end + 1)
    const afterHigh = high
    const hasBefore = beforeLow <= beforeHigh
    const hasAfter = afterLow <= afterHigh
    if (!hasBefore && !hasAfter) return null
    low = hasBefore ? beforeLow : afterLow - operation.count
    high = hasAfter ? afterHigh - operation.count : beforeHigh
  }

  return ascending ? { start: low, end: high } : { start: high, end: low }
}

function parseRange(range: string): Bounds | null {
  const match = /^\s*(\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d*))?\s*$/.exec(range)
  if (!match) return null
  const first = parseA1Address(match[1])
  const second = parseA1Address(match[2] || match[1])
  if (!first || !second) return null
  return {
    top: Math.min(first.row, second.row),
    bottom: Math.max(first.row, second.row),
    left: Math.min(first.column, second.column),
    right: Math.max(first.column, second.column),
  }
}

function boundsToRange(bounds: Bounds): string {
  const first = formatA1Address({
    row: bounds.top,
    column: bounds.left,
    rowAbsolute: false,
    columnAbsolute: false,
  })
  const second = formatA1Address({
    row: bounds.bottom,
    column: bounds.right,
    rowAbsolute: false,
    columnAbsolute: false,
  })
  if (!first || !second) throw new SheetStructureError('A range exceeds the XLSX worksheet limit.', 'LIMIT_EXCEEDED')
  return first === second ? first : `${first}:${second}`
}

function transformBounds(bounds: Bounds, operation: SheetStructureOperation): Bounds | null {
  const interval = operation.axis === 'row'
    ? { start: bounds.top, end: bounds.bottom }
    : { start: bounds.left, end: bounds.right }
  const transformed = transformInterval(interval, operation)
  if (!transformed) return null
  return operation.axis === 'row'
    ? { ...bounds, top: Math.min(transformed.start, transformed.end), bottom: Math.max(transformed.start, transformed.end) }
    : { ...bounds, left: Math.min(transformed.start, transformed.end), right: Math.max(transformed.start, transformed.end) }
}

function formatLike(original: string, value: number, axis: SheetAxis): string | null {
  if (axis === 'row') {
    if (value < 1 || value > MAX_SHEET_ROWS) return null
    return `${original.startsWith('$') ? '$' : ''}${value}`
  }
  const label = columnNumberToLabel(value)
  if (!label) return null
  const originalLabel = original.replace('$', '')
  const cased = originalLabel === originalLabel.toLowerCase() ? label.toLowerCase() : label
  return `${original.startsWith('$') ? '$' : ''}${cased}`
}

function transformCellToken(token: string, operation: SheetStructureOperation): string | null {
  const address = parseA1Address(token)
  if (!address) return token
  const across = operation.axis === 'row' ? address.column : address.row
  if (outsideSpan(operation, across, across)) return token
  const position = operation.axis === 'row' ? address.row : address.column
  const transformed = transformPosition(position, operation)
  if (transformed === null) return null
  const formatted = formatA1Address({
    ...address,
    row: operation.axis === 'row' ? transformed : address.row,
    column: operation.axis === 'column' ? transformed : address.column,
  })
  if (!formatted) return null
  const originalColumn = /^\$?([A-Za-z]+)/.exec(token)?.[1] || 'A'
  return originalColumn === originalColumn.toLowerCase()
    ? formatted.replace(/[A-Z]+/, (label) => label.toLowerCase())
    : formatted
}

function transformCellRange(firstToken: string, secondToken: string, operation: SheetStructureOperation): string | null {
  const first = parseA1Address(firstToken)
  const second = parseA1Address(secondToken)
  if (!first || !second) return `${firstToken}:${secondToken}`
  const acrossLow = operation.axis === 'row' ? Math.min(first.column, second.column) : Math.min(first.row, second.row)
  const acrossHigh = operation.axis === 'row' ? Math.max(first.column, second.column) : Math.max(first.row, second.row)
  if (outsideSpan(operation, acrossLow, acrossHigh)) return `${firstToken}:${secondToken}`
  const interval = operation.axis === 'row'
    ? { start: first.row, end: second.row }
    : { start: first.column, end: second.column }
  const transformed = transformInterval(interval, operation)
  if (!transformed) return null

  const firstPosition = transformed.start
  const secondPosition = transformed.end
  const firstFormatted = formatA1Address({
    ...first,
    row: operation.axis === 'row' ? firstPosition : first.row,
    column: operation.axis === 'column' ? firstPosition : first.column,
  })
  const secondFormatted = formatA1Address({
    ...second,
    row: operation.axis === 'row' ? secondPosition : second.row,
    column: operation.axis === 'column' ? secondPosition : second.column,
  })
  if (!firstFormatted || !secondFormatted) return null
  return `${preserveColumnCase(firstToken, firstFormatted)}:${preserveColumnCase(secondToken, secondFormatted)}`
}

function preserveColumnCase(original: string, formatted: string): string {
  const label = /^\$?([A-Za-z]+)/.exec(original)?.[1] || 'A'
  return label === label.toLowerCase()
    ? formatted.replace(/[A-Z]+/, (value) => value.toLowerCase())
    : formatted
}

function transformWholeRange(
  firstToken: string,
  secondToken: string,
  referenceAxis: SheetAxis,
  operation: SheetStructureOperation,
): string | null {
  if (referenceAxis !== operation.axis || operation.span) return `${firstToken}:${secondToken}`
  const first = referenceAxis === 'row'
    ? Number(firstToken.replace('$', ''))
    : columnLabelToNumber(firstToken.replace('$', ''))
  const second = referenceAxis === 'row'
    ? Number(secondToken.replace('$', ''))
    : columnLabelToNumber(secondToken.replace('$', ''))
  if (!first || !second) return `${firstToken}:${secondToken}`
  const transformed = transformInterval({ start: first, end: second }, operation)
  if (!transformed) return null
  const formattedFirst = formatLike(firstToken, transformed.start, referenceAxis)
  const formattedSecond = formatLike(secondToken, transformed.end, referenceAxis)
  return formattedFirst && formattedSecond ? `${formattedFirst}:${formattedSecond}` : null
}

function transformParsedReference(reference: ParsedReference, operation: SheetStructureOperation): string | null {
  if (reference.kind === 'cell') return transformCellToken(reference.first, operation)
  if (reference.kind === 'cell-range') return transformCellRange(reference.first, reference.second!, operation)
  if (reference.kind === 'column-range') return transformWholeRange(reference.first, reference.second!, 'column', operation)
  return transformWholeRange(reference.first, reference.second!, 'row', operation)
}

function copyDoubleQuotedString(source: string, start: number): number {
  let position = start + 1
  while (position < source.length) {
    if (source[position] !== '"') position += 1
    else if (source[position + 1] === '"') position += 2
    else return position + 1
  }
  return source.length
}

function sheetPrefixLength(source: string): number {
  if (source[0] === "'") {
    let position = 1
    while (position < source.length) {
      if (source[position] !== "'") position += 1
      else if (source[position + 1] === "'") position += 2
      else return source[position + 1] === '!' ? position + 2 : 0
    }
    return 0
  }
  const match = /^[A-Za-z_\\][A-Za-z0-9_.]*!/.exec(source)
  return match?.[0].length || 0
}

function referenceMatch(source: string): FormulaReferenceMatch | null {
  const prefixLength = sheetPrefixLength(source)
  const prefix = prefixLength ? source.slice(0, prefixLength) : ''
  const body = source.slice(prefixLength)
  const cellRange = /^(\$?[A-Za-z]{1,3}\$?[1-9]\d*):(\$?[A-Za-z]{1,3}\$?[1-9]\d*)/.exec(body)
  if (cellRange) {
    return {
      length: prefixLength + cellRange[0].length,
      prefix,
      reference: { kind: 'cell-range', first: cellRange[1], second: cellRange[2] },
    }
  }
  const columnRange = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})/.exec(body)
  if (columnRange) {
    return {
      length: prefixLength + columnRange[0].length,
      prefix,
      reference: { kind: 'column-range', first: columnRange[1], second: columnRange[2] },
    }
  }
  const rowRange = /^(\$?[1-9]\d*):(\$?[1-9]\d*)/.exec(body)
  if (rowRange) {
    return {
      length: prefixLength + rowRange[0].length,
      prefix,
      reference: { kind: 'row-range', first: rowRange[1], second: rowRange[2] },
    }
  }
  const cell = /^(\$?[A-Za-z]{1,3}\$?[1-9]\d*)/.exec(body)
  if (!cell) return null
  return {
    length: prefixLength + cell[0].length,
    prefix,
    reference: { kind: 'cell', first: cell[1] },
  }
}

function hasReferenceBoundaryBefore(source: string, position: number): boolean {
  return position === 0 || !/[A-Za-z0-9_.$]/.test(source[position - 1])
}

function hasReferenceBoundaryAfter(source: string, position: number): boolean {
  return !/[A-Za-z0-9_.]/.test(source[position] || '')
}

function nextNonWhitespace(source: string, start: number): string {
  let position = start
  while (/\s/.test(source[position] || '')) position += 1
  return source[position] || ''
}

function decodedSheetPrefix(prefix: string): string | null {
  if (!prefix) return null
  const withoutBang = prefix.slice(0, -1)
  if (withoutBang.startsWith("'") && withoutBang.endsWith("'")) {
    return withoutBang.slice(1, -1).replace(/''/g, "'")
  }
  return withoutBang
}

function referenceTargetsSheet(
  prefix: string,
  formulaSheetId: string | undefined,
  targetSheet: SheetData,
): boolean {
  if (!prefix) return formulaSheetId === targetSheet.id
  return decodedSheetPrefix(prefix)?.toLocaleLowerCase() === targetSheet.name.toLocaleLowerCase()
}

/**
 * Rewrite references for an insert/delete operation. Unlike copy-fill shifts,
 * structural edits move absolute and relative references alike. String literals
 * and structured-reference brackets are not interpreted as A1 addresses.
 */
export function rewriteFormulaForSheetStructure(
  formula: string,
  formulaSheetId: string | undefined,
  targetSheet: SheetData,
  operation: SheetStructureOperation,
): string {
  validateOperation(operation)
  let output = ''
  let position = 0

  while (position < formula.length) {
    if (formula[position] === '"') {
      const end = copyDoubleQuotedString(formula, position)
      output += formula.slice(position, end)
      position = end
      continue
    }
    if (formula[position] === '[') {
      const end = formula.indexOf(']', position + 1)
      if (end >= 0) {
        output += formula.slice(position, end + 1)
        position = end + 1
        continue
      }
    }
    if (hasReferenceBoundaryBefore(formula, position)) {
      const match = referenceMatch(formula.slice(position))
      if (match && hasReferenceBoundaryAfter(formula, position + match.length)) {
        // `[Book.xlsx]Sheet!A1` is an external-workbook reference. The bracket
        // was copied by the branch above, so consume its sheet/address token
        // unchanged rather than mistaking it for the local sheet of that name.
        if (match.prefix && formula[position - 1] === ']') {
          output += formula.slice(position, position + match.length)
          position += match.length
          continue
        }
        const singleCellFunction = match.reference.kind === 'cell'
          && nextNonWhitespace(formula, position + match.length) === '('
        if (!singleCellFunction && referenceTargetsSheet(match.prefix, formulaSheetId, targetSheet)) {
          const transformed = transformParsedReference(match.reference, operation)
          output += match.prefix + (transformed ?? '#REF!')
          position += match.length
          continue
        }
      }
    }
    output += formula[position]
    position += 1
  }
  return output
}

function transformLocalRangeText(text: string, operation: SheetStructureOperation): string | null {
  const cellMatch = /^\s*(\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d*))?\s*$/.exec(text)
  if (cellMatch) {
    if (!cellMatch[2]) return transformCellToken(cellMatch[1], operation)
    return transformCellRange(cellMatch[1], cellMatch[2], operation)
  }
  const columnMatch = /^\s*(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})\s*$/.exec(text)
  if (columnMatch) return transformWholeRange(columnMatch[1], columnMatch[2], 'column', operation)
  const rowMatch = /^\s*(\$?[1-9]\d*):(\$?[1-9]\d*)\s*$/.exec(text)
  if (rowMatch) return transformWholeRange(rowMatch[1], rowMatch[2], 'row', operation)
  return rewriteFormulaForSheetStructure(text, '__local__', { id: '__local__', name: '', rowCount: 1, colCount: 1, cells: {}, merges: [], colWidths: {}, rowHeights: {} }, operation)
}

function transformSqref(text: string, operation: SheetStructureOperation): string | null {
  const ranges = text.trim().split(/\s+/).filter(Boolean)
  if (!ranges.length) return text
  const transformed = ranges
    .map((range) => transformLocalRangeText(range, operation))
    .filter((range): range is string => Boolean(range && range !== '#REF!'))
  return transformed.length ? transformed.join(' ') : null
}

/** A shift may not move part of a merge, a table or an array range (Excel refuses too). */
function preflightShift(sheet: SheetData, operation: SheetStructureOperation): void {
  if (!operation.span) return
  const { start } = operationInterval(operation)
  const straddles = (bounds: Bounds) => {
    const along = operation.axis === 'row' ? bounds.bottom : bounds.right
    if (along < start) return false
    const low = operation.axis === 'row' ? bounds.left : bounds.top
    const high = operation.axis === 'row' ? bounds.right : bounds.bottom
    const intersects = high >= operation.span!.start + 1 && low <= operation.span!.end + 1
    return intersects && outsideSpan(operation, low, high)
  }
  for (const merge of sheet.merges || []) {
    const bounds = parseRange(merge)
    if (bounds && straddles(bounds)) throw new SheetStructureError(`Shifting these cells would split the merged cells ${merge}.`, 'SHIFT_CONFLICT')
  }
  for (const table of sheet.tables || []) {
    const bounds = parseRange(table.ref)
    if (bounds && straddles(bounds)) throw new SheetStructureError(`Shifting these cells would move part of the table ${table.name}. Select whole table columns or rows.`, 'SHIFT_CONFLICT')
  }
  for (const cell of Object.values(sheet.cells)) {
    const bounds = cell.formulaRange ? parseRange(cell.formulaRange) : null
    if (bounds && straddles(bounds)) throw new SheetStructureError(`Shifting these cells would split the array formula range ${cell.formulaRange}.`, 'SHIFT_CONFLICT')
  }
}

function preflightArrayRanges(sheet: SheetData, operation: SheetStructureOperation): void {
  const { start, end } = operationInterval(operation)
  const seen = new Set<string>()
  for (const cell of Object.values(sheet.cells)) {
    if (!cell.formulaRange || seen.has(cell.formulaRange)) continue
    seen.add(cell.formulaRange)
    const bounds = parseRange(cell.formulaRange)
    if (!bounds) continue
    if (operation.span && outsideSpan(operation, operation.axis === 'row' ? bounds.left : bounds.top, operation.axis === 'row' ? bounds.right : bounds.bottom)) continue
    const low = operation.axis === 'row' ? bounds.top : bounds.left
    const high = operation.axis === 'row' ? bounds.bottom : bounds.right
    const conflicts = operation.kind === 'insert'
      ? start > low && start <= high
      : Math.max(start, low) <= Math.min(end, high) && !(start <= low && end >= high)
    if (conflicts) {
      throw new SheetStructureError(
        `The edit would split the array/shared formula range ${cell.formulaRange}. Select the complete range or edit outside it.`,
        'ARRAY_RANGE_CONFLICT',
      )
    }
  }
}

function preflightOverflow(sheet: SheetData, operation: SheetStructureOperation): void {
  if (operation.kind !== 'insert') return
  const maximum = maxForAxis(operation.axis)
  const overflows = (position: number) => position >= operation.index + 1 && position + operation.count > maximum
  for (const address of Object.keys(sheet.cells)) {
    const coordinate = coordinateFromAddress(address)
    if (coordinate && overflows(operation.axis === 'row' ? coordinate.row + 1 : coordinate.col + 1)) {
      throw new SheetStructureError('The edit would push worksheet cells beyond the XLSX limit.', 'LIMIT_EXCEEDED')
    }
  }
  if (operation.span) return
  const dimensionRecord = operation.axis === 'row'
    ? { ...(sheet.rowHeights || {}), ...(sheet.rowProperties || {}) }
    : { ...(sheet.colWidths || {}), ...(sheet.columnProperties || {}) }
  if (Object.keys(dimensionRecord).some((key) => overflows(Number(key)))) {
    throw new SheetStructureError('The edit would push row or column metadata beyond the XLSX limit.', 'LIMIT_EXCEEDED')
  }
  const hidden = operation.axis === 'row' ? sheet.hiddenRows : sheet.hiddenCols
  if ((hidden || []).some((value) => overflows(Number(value)))) {
    throw new SheetStructureError('The edit would push hidden dimensions beyond the XLSX limit.', 'LIMIT_EXCEEDED')
  }
  for (const merge of sheet.merges || []) {
    const bounds = parseRange(merge)
    if (!bounds) continue
    const high = operation.axis === 'row' ? bounds.bottom : bounds.right
    if (overflows(high)) {
      throw new SheetStructureError('The edit would push a merged range beyond the XLSX limit.', 'LIMIT_EXCEEDED')
    }
  }
}

function transformCells(sheet: SheetData, operation: SheetStructureOperation): void {
  const cells: Record<string, CellData> = {}
  for (const [address, sourceCell] of Object.entries(sheet.cells)) {
    const coordinate = coordinateFromAddress(address)
    if (!coordinate) {
      cells[address] = sourceCell
      continue
    }
    const across = operation.axis === 'row' ? coordinate.col + 1 : coordinate.row + 1
    if (outsideSpan(operation, across, across)) {
      cells[address] = sourceCell
      continue
    }
    const position = operation.axis === 'row' ? coordinate.row + 1 : coordinate.col + 1
    const transformed = transformPosition(position, operation)
    if (transformed === null) continue
    const destination = operation.axis === 'row'
      ? { ...coordinate, row: transformed - 1 }
      : { ...coordinate, col: transformed - 1 }
    const cell = sourceCell
    if (cell.formulaRange) {
      const formulaRange = transformLocalRangeText(cell.formulaRange, operation)
      if (formulaRange) cell.formulaRange = formulaRange
      else delete cell.formulaRange
    }
    const extended = cell as CellData & { sharedFormulaMaster?: string }
    if (extended.sharedFormulaMaster) {
      const master = transformCellToken(extended.sharedFormulaMaster, operation)
      if (master) extended.sharedFormulaMaster = master
      else delete extended.sharedFormulaMaster
    }
    cells[addressFromCoordinate(destination)] = cell
  }
  sheet.cells = cells
}

function transformNumericRecord<T>(
  record: Record<string, T> | undefined,
  operation: SheetStructureOperation,
): Record<string, T> | undefined {
  if (!record) return undefined
  const result: Record<string, T> = {}
  for (const [key, value] of Object.entries(record)) {
    const index = Number(key)
    if (!Number.isSafeInteger(index) || index < 1) {
      result[key] = value
      continue
    }
    const transformed = transformPosition(index, operation)
    if (transformed !== null) result[String(transformed)] = value
  }
  return result
}

function transformHiddenDimensions(values: number[] | undefined, operation: SheetStructureOperation): number[] | undefined {
  if (!values) return undefined
  const transformed = values
    .map((value) => transformPosition(Number(value), operation))
    .filter((value): value is number => value !== null && Number.isSafeInteger(value) && value >= 1)
  return [...new Set(transformed)].sort((left, right) => left - right)
}

function transformAddressForNavigation(address: unknown, operation: SheetStructureOperation): string | undefined {
  if (typeof address !== 'string') return undefined
  const parsed = parseA1Address(address)
  if (!parsed) return address
  const position = operation.axis === 'row' ? parsed.row : parsed.column
  const transformed = transformNavigationPosition(position, operation)
  return formatA1Address({
    ...parsed,
    row: operation.axis === 'row' ? transformed : parsed.row,
    column: operation.axis === 'column' ? transformed : parsed.column,
  }) || address
}

function transformFrozenCount(value: unknown, operation: SheetStructureOperation): number {
  const current = Math.max(0, Math.floor(Number(value) || 0))
  const { start, end } = operationInterval(operation)
  if (operation.kind === 'insert') return start <= current ? current + operation.count : current
  const removed = Math.max(0, Math.min(current, end) - start + 1)
  return Math.max(0, current - removed)
}

function transformViewsAndFrozen(sheet: SheetData, operation: SheetStructureOperation): void {
  if (sheet.frozen) {
    if (operation.axis === 'row') sheet.frozen.rows = transformFrozenCount(sheet.frozen.rows, operation)
    else sheet.frozen.columns = transformFrozenCount(sheet.frozen.columns, operation)
    if (sheet.frozen.topLeftCell) sheet.frozen.topLeftCell = transformAddressForNavigation(sheet.frozen.topLeftCell, operation)
    if (sheet.frozen.activeCell) sheet.frozen.activeCell = transformAddressForNavigation(sheet.frozen.activeCell, operation)
  }
  if (!sheet.views) return
  for (const view of sheet.views) {
    if (operation.axis === 'row' && view.state === 'frozen') view.ySplit = transformFrozenCount(view.ySplit, operation)
    if (operation.axis === 'column' && view.state === 'frozen') view.xSplit = transformFrozenCount(view.xSplit, operation)
    if (view.topLeftCell) view.topLeftCell = transformAddressForNavigation(view.topLeftCell, operation)
    if (view.activeCell) view.activeCell = transformAddressForNavigation(view.activeCell, operation)
  }
}

function transformBreaks(value: unknown[] | undefined, operation: SheetStructureOperation): unknown[] | undefined {
  if (!value) return undefined
  return value.flatMap((item) => {
    if (!item || typeof item !== 'object') return [item]
    const record = item as Record<string, unknown>
    const key = 'id' in record ? 'id' : 'row' in record ? 'row' : 'index' in record ? 'index' : null
    if (!key) return [item]
    const position = Number(record[key])
    if (!Number.isSafeInteger(position) || position < 1) return [item]
    const transformed = transformPosition(position, operation)
    return transformed === null ? [] : [{ ...record, [key]: transformed }]
  })
}

function transformFormulaProperties(
  value: unknown,
  formulaSheetId: string,
  targetSheet: SheetData,
  operation: SheetStructureOperation,
  propertyName = '',
): unknown {
  if (Array.isArray(value)) {
    return value.map((item) => (
      typeof item === 'string' && /formulae?|formula[12]/i.test(propertyName)
        ? rewriteFormulaForSheetStructure(item, formulaSheetId, targetSheet, operation)
        : transformFormulaProperties(item, formulaSheetId, targetSheet, operation, propertyName)
    ))
  }
  if (!value || typeof value !== 'object') {
    return typeof value === 'string' && /formulae?|formula[12]/i.test(propertyName)
      ? rewriteFormulaForSheetStructure(value, formulaSheetId, targetSheet, operation)
      : value
  }
  const record = value as Record<string, unknown>
  // Conditional-format thresholds (cfvo) hold formulas in `value` for formula/num/percent types.
  const threshold = typeof record.value === 'string' && ['formula', 'num', 'percent', 'percentile'].includes(String(record.type)) &&
    !/^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/.test(record.value)
  const result: Record<string, unknown> = {}
  for (const [key, child] of Object.entries(record)) {
    result[key] = key === 'value' && threshold
      ? rewriteFormulaForSheetStructure(child as string, formulaSheetId, targetSheet, operation)
      : transformFormulaProperties(child, formulaSheetId, targetSheet, operation, key)
  }
  return result
}

function transformDataValidations(sheet: SheetData, targetSheet: SheetData, operation: SheetStructureOperation): void {
  if (!sheet.dataValidations) return
  const result: Record<string, unknown> = {}
  for (const [range, validation] of Object.entries(sheet.dataValidations)) {
    const transformedRange = transformSqref(range, operation)
    if (!transformedRange) continue
    result[transformedRange] = transformFormulaProperties(validation, sheet.id, targetSheet, operation)
  }
  sheet.dataValidations = result
}

function transformConditionalFormatting(sheet: SheetData, targetSheet: SheetData, operation: SheetStructureOperation): void {
  if (!sheet.conditionalFormattings) return
  sheet.conditionalFormattings = sheet.conditionalFormattings.flatMap((item) => {
    if (!item || typeof item !== 'object') return [item]
    const record = transformFormulaProperties(item, sheet.id, targetSheet, operation) as Record<string, unknown>
    if (typeof record.ref === 'string') {
      const ref = transformSqref(record.ref, operation)
      if (!ref) return []
      record.ref = ref
    }
    return [record]
  })
}

function endpointAddress(endpoint: unknown): string | null {
  if (typeof endpoint === 'string') return parseA1Address(endpoint) ? endpoint : null
  if (!endpoint || typeof endpoint !== 'object') return null
  const record = endpoint as Record<string, unknown>
  const row = Number(record.row)
  const column = Number(record.column)
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column)) return null
  return formatA1Address({ row, column, rowAbsolute: false, columnAbsolute: false })
}

function transformedEndpoint(original: unknown, address: string): unknown {
  if (typeof original === 'string') return address
  if (!original || typeof original !== 'object') return original
  const parsed = parseA1Address(address)
  return parsed ? { ...(original as Record<string, unknown>), row: parsed.row, column: parsed.column } : original
}

function transformAutoFilter(sheet: SheetData, operation: SheetStructureOperation): void {
  if (typeof sheet.autoFilter === 'string') {
    const transformed = transformSqref(sheet.autoFilter, operation)
    if (transformed) sheet.autoFilter = transformed
    else delete sheet.autoFilter
    return
  }
  if (!sheet.autoFilter || typeof sheet.autoFilter !== 'object') return
  const filter = sheet.autoFilter as Record<string, unknown>
  for (const key of ['ref', 'Ref']) {
    if (typeof filter[key] !== 'string') continue
    const transformed = transformSqref(filter[key] as string, operation)
    if (transformed) filter[key] = transformed
    else delete sheet.autoFilter
    return
  }
  const from = endpointAddress(filter.from)
  const to = endpointAddress(filter.to)
  if (!from || !to) return
  const transformed = transformLocalRangeText(`${from}:${to}`, operation)
  if (!transformed) {
    delete sheet.autoFilter
    return
  }
  const match = /^([^:]+):([^:]+)$/.exec(transformed)
  if (!match) return
  filter.from = transformedEndpoint(filter.from, match[1])
  filter.to = transformedEndpoint(filter.to, match[2])
}

/** The sheet AutoFilter's live state, or one built from an imported range string. */
function sheetFilterState(sheet: SheetData): SheetFilterState | null {
  if (sheet.filter && parseRange(sheet.filter.ref)) return sheet.filter
  return typeof sheet.autoFilter === 'string' && parseRange(sheet.autoFilter) ? { ref: sheet.autoFilter, columns: {} } : null
}

/** 1-based data rows (header excluded) of every filter with criteria: the sheet's and the tables'. */
function activeFilterRows(sheet: SheetData): Interval[] {
  const intervals: Interval[] = []
  if (sheet.filter && filterIsActive(sheet.filter)) {
    const bounds = parseRange(sheet.filter.ref)
    if (bounds) intervals.push({ start: bounds.top + 1, end: bounds.bottom })
  }
  for (const table of sheet.tables || []) {
    const regions = tableRegions(table)
    if (!regions || regions.header === null || !filterIsActive(table.filter)) continue
    intervals.push({ start: regions.dataTop + 1, end: regions.dataBottom + 1 })
  }
  return intervals
}

/**
 * Rows a filter hid come back when no filter owns them any more (the header row or the filtered
 * columns were deleted), as in Excel; without this they would stay hidden as if hidden by hand.
 */
function releaseOrphanedFilteredRows(sheet: SheetData): void {
  if (!sheet.filteredRows?.length) return
  const owned = activeFilterRows(sheet)
  const released = new Set(sheet.filteredRows.filter((row) => !owned.some((interval) => row >= interval.start && row <= interval.end)))
  if (!released.size) return
  sheet.filteredRows = sheet.filteredRows.filter((row) => !released.has(row))
  sheet.hiddenRows = (sheet.hiddenRows || []).filter((row) => !released.has(row))
  if (!sheet.filteredRows.length) delete sheet.filteredRows
}

/**
 * Keep the AutoFilter in step with inserted or deleted rows and columns: its range and column
 * criteria move (filter.ts adjustFilterForStructure) and the rows it hides shift with the grid,
 * so the header row is never treated as data and removing the filter shows every row again.
 * Deleting the header row (or every filter column) removes the filter, as in Excel.
 */
function transformSheetFilter(sheet: SheetData, operation: SheetStructureOperation): void {
  if (operation.axis === 'row' && sheet.filteredRows) sheet.filteredRows = adjustRowNumbersForStructure(sheet.filteredRows, operation)
  const state = sheetFilterState(sheet)
  if (!state) {
    transformAutoFilter(sheet, operation)
    return
  }
  const next = adjustFilterForStructure(state, operation)
  if (next) {
    if (sheet.filter) sheet.filter = next
    if (typeof sheet.autoFilter === 'string') sheet.autoFilter = next.ref
    else transformAutoFilter(sheet, operation)
    return
  }
  delete sheet.filter
  delete sheet.autoFilter
}

/**
 * Shifting cells (Insert/Delete cells) moves the AutoFilter only when the shifted band spans all
 * of its columns (or rows); a partial band leaves it where it is. Filtered lists refuse the
 * shift beforehand (see preflightFilteredShift).
 */
function transformShiftedFilter(sheet: SheetData, operation: SheetStructureOperation): void {
  const state = sheetFilterState(sheet)
  const bounds = state ? parseRange(state.ref) : null
  if (!state || !bounds) return
  const low = operation.axis === 'row' ? bounds.left : bounds.top
  const high = operation.axis === 'row' ? bounds.right : bounds.bottom
  if (outsideSpan(operation, low, high)) return
  const next = adjustFilterForStructure(state, { ...operation, span: undefined })
  if (next) {
    if (sheet.filter) sheet.filter = next
    if (typeof sheet.autoFilter === 'string') sheet.autoFilter = next.ref
    return
  }
  delete sheet.filter
  delete sheet.autoFilter
}

/**
 * Excel refuses to shift cells inside a filtered list ("This operation is not allowed"): the
 * hidden rows would no longer line up with the data. Whole-row and whole-column edits work.
 */
function preflightFilteredShift(sheet: SheetData, operation: SheetStructureOperation): void {
  if (!operation.span) return
  const { start } = operationInterval(operation)
  const ranges: Bounds[] = []
  const state = sheetFilterState(sheet)
  if (state && (filterIsActive(state) || Boolean(sheet.filteredRows?.length))) {
    const bounds = parseRange(state.ref)
    if (bounds) ranges.push(bounds)
  }
  for (const table of sheet.tables || []) {
    if (!filterIsActive(table.filter)) continue
    const bounds = parseRange(table.ref)
    if (bounds) ranges.push(bounds)
  }
  for (const bounds of ranges) {
    const along = operation.axis === 'row' ? bounds.bottom : bounds.right
    if (along < start) continue
    const low = operation.axis === 'row' ? bounds.left : bounds.top
    const high = operation.axis === 'row' ? bounds.right : bounds.bottom
    if (high >= operation.span.start + 1 && low <= operation.span.end + 1) {
      throw new SheetStructureError('Cells in a filtered list cannot be shifted. Clear the filter first, or insert or delete entire rows or columns.', 'SHIFT_CONFLICT')
    }
  }
}

function transformPageSetup(sheet: SheetData, operation: SheetStructureOperation): void {
  if (!sheet.pageSetup) return
  for (const key of ['printArea', 'printTitlesRow', 'printTitlesColumn']) {
    const value = sheet.pageSetup[key]
    if (typeof value !== 'string') continue
    if (key === 'printArea') {
      // A whole-column ($A:$F) or whole-row ($1:$20) area is kept as a bounded range plus its
      // original form (printAreaWhole, keyed by the range); the note moves with the range so the
      // area still prints to the used extent and is saved in its original form.
      const whole = sheet.pageSetup.printAreaWhole && typeof sheet.pageSetup.printAreaWhole === 'object'
        ? sheet.pageSetup.printAreaWhole as Record<string, unknown>
        : null
      const nextWhole: Record<string, string> = {}
      const areas: string[] = []
      for (const area of splitFormulaAreas(value.replace(/&&/g, ','))) {
        const moved = rewriteFormulaForSheetStructure(area, sheet.id, sheet, operation)
        if (moved.includes('#REF!')) continue
        areas.push(moved)
        const original = whole?.[area.split('!').pop()!.replace(/'/g, '').trim()]
        const bounds = typeof original === 'string' ? parseRange(moved.split('!').pop()!.replace(/'/g, '')) : null
        if (!bounds || typeof original !== 'string') continue
        const key = moved.split('!').pop()!.replace(/'/g, '').trim()
        if (/^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(original)) {
          nextWhole[key] = `$${columnNumberToLabel(bounds.left)}:$${columnNumberToLabel(bounds.right)}`
        } else if (/^\$?\d{1,7}:\$?\d{1,7}$/.test(original)) {
          nextWhole[key] = `$${bounds.top}:$${bounds.bottom}`
        }
      }
      if (areas.length) sheet.pageSetup.printArea = areas.join(',')
      else delete sheet.pageSetup.printArea
      if (Object.keys(nextWhole).length) sheet.pageSetup.printAreaWhole = nextWhole
      else delete sheet.pageSetup.printAreaWhole
      continue
    }
    const transformed = transformLocalRangeText(value, operation)
    if (transformed) sheet.pageSetup[key] = transformed
    else delete sheet.pageSetup[key]
  }
}

/** Split print-area unions without treating commas inside quoted sheet names as separators. */
function splitFormulaAreas(value: string): string[] {
  const result: string[] = []
  let current = ''
  let quoted = false
  for (let position = 0; position < value.length; position += 1) {
    const character = value[position]
    if (character === "'") {
      current += character
      if (quoted && value[position + 1] === "'") {
        current += value[++position]
      } else {
        quoted = !quoted
      }
    } else if (character === ',' && !quoted) {
      if (current.trim()) result.push(current.trim())
      current = ''
    } else {
      current += character
    }
  }
  if (current.trim()) result.push(current.trim())
  return result
}

/** Shifting cells moves only the band's cells, merges and tables; rows and columns stay. */
function transformShiftedCells(sheet: SheetData, operation: SheetStructureOperation): void {
  transformCells(sheet, operation)
  sheet.merges = (sheet.merges || []).flatMap((range) => {
    const bounds = parseRange(range)
    if (!bounds) return [range]
    const low = operation.axis === 'row' ? bounds.left : bounds.top
    const high = operation.axis === 'row' ? bounds.right : bounds.bottom
    if (outsideSpan(operation, low, high)) return [range]
    const transformed = transformBounds(bounds, operation)
    return transformed ? [boundsToRange(transformed)] : []
  })
  transformTablesForStructure(sheet, operation)
  transformShiftedFilter(sheet, operation)
  if (operation.kind === 'insert') {
    if (operation.axis === 'row') sheet.rowCount = Math.min(MAX_SHEET_ROWS, Math.max(1, Math.floor(Number(sheet.rowCount) || 1)) + operation.count)
    else sheet.colCount = Math.min(MAX_SHEET_COLUMNS, Math.max(1, Math.floor(Number(sheet.colCount) || 1)) + operation.count)
  }
}

function transformSheetStructure(sheet: SheetData, operation: SheetStructureOperation): void {
  if (operation.span) {
    transformShiftedCells(sheet, operation)
    return
  }
  transformCells(sheet, operation)
  sheet.merges = (sheet.merges || []).flatMap((range) => {
    const bounds = parseRange(range)
    if (!bounds) return [range]
    const transformed = transformBounds(bounds, operation)
    return transformed ? [boundsToRange(transformed)] : []
  })

  if (operation.axis === 'row') {
    sheet.rowHeights = transformNumericRecord(sheet.rowHeights, operation) || {}
    sheet.rowProperties = transformNumericRecord(sheet.rowProperties, operation)
    sheet.hiddenRows = transformHiddenDimensions(sheet.hiddenRows, operation)
    sheet.rowBreaks = transformBreaks(sheet.rowBreaks, operation)
    const oldCount = Math.max(1, Math.floor(Number(sheet.rowCount) || 1))
    sheet.rowCount = operation.kind === 'insert'
      ? Math.min(MAX_SHEET_ROWS, Math.max(oldCount + operation.count, operation.index + operation.count))
      : Math.max(1, oldCount - Math.max(0, Math.min(operation.count, oldCount - operation.index)))
  } else {
    sheet.colWidths = transformNumericRecord(sheet.colWidths, operation) || {}
    sheet.columnProperties = transformNumericRecord(sheet.columnProperties, operation)
    sheet.hiddenCols = transformHiddenDimensions(sheet.hiddenCols, operation)
    const extended = sheet as SheetData & { columnBreaks?: unknown[] }
    extended.columnBreaks = transformBreaks(extended.columnBreaks, operation)
    const oldCount = Math.max(1, Math.floor(Number(sheet.colCount) || 1))
    sheet.colCount = operation.kind === 'insert'
      ? Math.min(MAX_SHEET_COLUMNS, Math.max(oldCount + operation.count, operation.index + operation.count))
      : Math.max(1, oldCount - Math.max(0, Math.min(operation.count, oldCount - operation.index)))
  }

  transformViewsAndFrozen(sheet, operation)
  transformSheetFilter(sheet, operation)
  transformTablesForStructure(sheet, operation)
  releaseOrphanedFilteredRows(sheet)
  if (sheet.pivots?.length) {
    // Pivot blocks move with their anchor; one whose anchor row/column is deleted goes with it.
    sheet.pivots = sheet.pivots.flatMap((pivot) => {
      const position = (operation.axis === 'row' ? pivot.anchor.row : pivot.anchor.col) + 1
      const transformed = transformPosition(position, operation)
      if (transformed === null) return []
      return [{ ...pivot, anchor: operation.axis === 'row' ? { ...pivot.anchor, row: transformed - 1 } : { ...pivot.anchor, col: transformed - 1 } }]
    })
  }
  transformPageSetup(sheet, operation)
}

function localSheetIdForDefinedName(workbook: WorkbookModel, name: DefinedName): string | undefined {
  const extended = name as DefinedName & { localSheetRefId?: string }
  if (extended.localSheetRefId && workbook.sheets.some((sheet) => sheet.id === extended.localSheetRefId)) {
    return extended.localSheetRefId
  }
  const index = Number.isInteger(name.localSheetIndex)
    ? name.localSheetIndex
    : Number.isInteger(name.localSheetId) ? name.localSheetId : undefined
  return index !== undefined ? workbook.sheets[index]?.id : undefined
}

function transformDefinedNames(workbook: WorkbookModel, targetSheet: SheetData, operation: SheetStructureOperation): void {
  if (workbook.definedNames) {
    for (const name of workbook.definedNames) {
      const localSheetId = localSheetIdForDefinedName(workbook, name)
      if (name.ref) name.ref = rewriteFormulaForSheetStructure(name.ref, localSheetId, targetSheet, operation)
      if (name.ranges) name.ranges = name.ranges.map((range) => rewriteFormulaForSheetStructure(range, localSheetId, targetSheet, operation))
    }
  }
  const metadataNames = workbook.metadata?.definedNames
  if (metadataNames) {
    for (const name of metadataNames) {
      const localSheetId = Number.isInteger(name.localSheetId) ? workbook.sheets[name.localSheetId!]?.id : undefined
      if (name.ranges) name.ranges = rewriteFormulaForSheetStructure(name.ranges, localSheetId, targetSheet, operation)
      if (name.formula) name.formula = rewriteFormulaForSheetStructure(name.formula, localSheetId, targetSheet, operation)
    }
  }
}

function rewriteWorkbookFormulas(workbook: WorkbookModel, targetSheet: SheetData, operation: SheetStructureOperation): void {
  for (const sheet of workbook.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (!cell.formula) continue
      cell.formula = rewriteFormulaForSheetStructure(cell.formula, sheet.id, targetSheet, operation)
      delete cell.result
      delete cell.resultType
      delete cell.display
    }
  }
  transformDefinedNames(workbook, targetSheet, operation)
  for (const sheet of workbook.sheets) {
    for (const pivot of sheet.pivots || []) {
      if (/[!:]/.test(pivot.source)) pivot.source = rewriteFormulaForSheetStructure(pivot.source, sheet.id, targetSheet, operation)
    }
    if (sheet.sparklineGroups?.length) {
      sheet.sparklineGroups = sheet.sparklineGroups.map((group) => ({
        ...group,
        sparklines: group.sparklines.flatMap((item) => {
          const source = item.source ? rewriteFormulaForSheetStructure(item.source, sheet.id, targetSheet, operation) : item.source
          const cell = sheet.id === targetSheet.id ? transformLocalRangeText(item.cell, operation) : item.cell
          return cell ? [{ source, cell }] : []
        }),
      })).filter((group) => group.sparklines.length)
    }
  }
  for (const sheet of workbook.sheets) {
    if (sheet.id === targetSheet.id) {
      transformDataValidations(sheet, targetSheet, operation)
      transformConditionalFormatting(sheet, targetSheet, operation)
    } else {
      // Validation and conditional-format formulas on other sheets can still
      // refer explicitly to the structurally edited sheet. Their applied
      // ranges remain local to their own sheet and therefore do not move.
      if (sheet.dataValidations) {
        for (const [key, value] of Object.entries(sheet.dataValidations)) {
          sheet.dataValidations[key] = transformFormulaProperties(value, sheet.id, targetSheet, operation)
        }
      }
      if (sheet.conditionalFormattings) {
        sheet.conditionalFormattings = sheet.conditionalFormattings.map((item) => (
          transformFormulaProperties(item, sheet.id, targetSheet, operation)
        ))
      }
    }
  }
}

/**
 * Apply a row/column edit without mutating the supplied workbook.
 *
 * Cell payloads, cross-sheet formulas, merges, dimensions, hidden state,
 * frozen panes, filters, validation/conditional-format ranges, print ranges,
 * and defined names are moved together. Formula caches are invalidated and the
 * workbook is marked for a complete recalculation on its next compatible open.
 */
export function applySheetStructureOperation(
  workbook: WorkbookModel,
  sheetId: string,
  operation: SheetStructureOperation,
): WorkbookModel {
  validateOperation(operation)
  const sourceSheet = workbook.sheets.find((sheet) => sheet.id === sheetId)
  if (!sourceSheet) throw new SheetStructureError(`Worksheet ${sheetId} was not found.`, 'SHEET_NOT_FOUND')
  preflightShift(sourceSheet, operation)
  preflightFilteredShift(sourceSheet, operation)
  preflightArrayRanges(sourceSheet, operation)
  preflightOverflow(sourceSheet, operation)

  const next = clone(workbook)
  const targetSheet = next.sheets.find((sheet) => sheet.id === sheetId)!
  transformSheetStructure(targetSheet, operation)
  rewriteWorkbookFormulas(next, targetSheet, operation)
  next.metadata ||= {}
  next.metadata.calcProperties = {
    ...(next.metadata.calcProperties || {}),
    fullCalcOnLoad: true,
    forceFullCalc: true,
  }
  return next
}

/**
 * Excel's Insert/Delete cells: shift the cells in (and beyond) a block down/right to make room,
 * or delete the block and shift the following cells up/left. References move with the cells.
 */
export function shiftCells(workbook: WorkbookModel, sheetId: string, bounds: Bounds, direction: CellShiftDirection): WorkbookModel {
  const vertical = direction === 'down' || direction === 'up'
  const operation: SheetStructureOperation = {
    axis: vertical ? 'row' : 'column',
    kind: direction === 'down' || direction === 'right' ? 'insert' : 'delete',
    index: vertical ? bounds.top : bounds.left,
    count: vertical ? bounds.bottom - bounds.top + 1 : bounds.right - bounds.left + 1,
    span: vertical ? { start: bounds.left, end: bounds.right } : { start: bounds.top, end: bounds.bottom },
  }
  return applySheetStructureOperation(workbook, sheetId, operation)
}

export function insertRows(workbook: WorkbookModel, sheetId: string, index: number, count = 1): WorkbookModel {
  return applySheetStructureOperation(workbook, sheetId, { axis: 'row', kind: 'insert', index, count })
}

export function deleteRows(workbook: WorkbookModel, sheetId: string, index: number, count = 1): WorkbookModel {
  return applySheetStructureOperation(workbook, sheetId, { axis: 'row', kind: 'delete', index, count })
}

export function insertColumns(workbook: WorkbookModel, sheetId: string, index: number, count = 1): WorkbookModel {
  return applySheetStructureOperation(workbook, sheetId, { axis: 'column', kind: 'insert', index, count })
}

export function deleteColumns(workbook: WorkbookModel, sheetId: string, index: number, count = 1): WorkbookModel {
  return applySheetStructureOperation(workbook, sheetId, { axis: 'column', kind: 'delete', index, count })
}

/** Resolve menu wording against a zero-based rectangular selection. */
export function operationForSelection(
  selection: GridSelection,
  command: SelectionStructureCommand,
): SheetStructureOperation {
  const bounds = normalizedBounds(selection)
  switch (command) {
    case 'insert-rows-above':
      return { axis: 'row', kind: 'insert', index: bounds.top, count: bounds.bottom - bounds.top + 1 }
    case 'insert-rows-below':
      return { axis: 'row', kind: 'insert', index: bounds.bottom + 1, count: bounds.bottom - bounds.top + 1 }
    case 'delete-rows':
      return { axis: 'row', kind: 'delete', index: bounds.top, count: bounds.bottom - bounds.top + 1 }
    case 'insert-columns-left':
      return { axis: 'column', kind: 'insert', index: bounds.left, count: bounds.right - bounds.left + 1 }
    case 'insert-columns-right':
      return { axis: 'column', kind: 'insert', index: bounds.right + 1, count: bounds.right - bounds.left + 1 }
    case 'delete-columns':
      return { axis: 'column', kind: 'delete', index: bounds.left, count: bounds.right - bounds.left + 1 }
  }
}

function selectionAfterOperation(
  selection: GridSelection,
  operation: SheetStructureOperation,
): GridSelection {
  const bounds = normalizedBounds(selection)
  if (operation.axis === 'row') {
    const start = operation.index
    const end = operation.kind === 'insert' ? start + operation.count - 1 : start
    const row = Math.min(start, MAX_SHEET_ROWS - 1)
    return {
      anchor: { row, col: bounds.left },
      focus: { row: Math.min(end, MAX_SHEET_ROWS - 1), col: bounds.right },
    }
  }
  const start = operation.index
  const end = operation.kind === 'insert' ? start + operation.count - 1 : start
  const col = Math.min(start, MAX_SHEET_COLUMNS - 1)
  return {
    anchor: { row: bounds.top, col },
    focus: { row: bounds.bottom, col: Math.min(end, MAX_SHEET_COLUMNS - 1) },
  }
}

export function applySelectionStructureCommand(
  workbook: WorkbookModel,
  sheetId: string,
  selection: GridSelection,
  command: SelectionStructureCommand,
): SelectionStructureResult {
  const operation = operationForSelection(selection, command)
  const nextWorkbook = applySheetStructureOperation(workbook, sheetId, operation)
  const targetSheet = nextWorkbook.sheets.find((sheet) => sheet.id === sheetId)!
  const nextSelection = selectionAfterOperation(selection, operation)
  const clampCoordinate = (coordinate: GridCoordinate): GridCoordinate => ({
    row: Math.min(coordinate.row, Math.max(0, targetSheet.rowCount - 1)),
    col: Math.min(coordinate.col, Math.max(0, targetSheet.colCount - 1)),
  })
  return {
    workbook: nextWorkbook,
    selection: {
      anchor: clampCoordinate(nextSelection.anchor),
      focus: clampCoordinate(nextSelection.focus),
    },
    operation,
  }
}

/**
 * Apply a formula-text rewrite (a sheet rename or removal, say) to every formula-bearing part of
 * the workbook: cell formulas, validation and conditional-format formulas, defined names and
 * sparkline sources. Mutates `workbook` in place, so it can run on an immer draft.
 */
export function rewriteWorkbookFormulaText(workbook: WorkbookModel, rewrite: (formula: string) => string): void {
  for (const sheet of workbook.sheets) {
    for (const address in sheet.cells) {
      const cell = sheet.cells[address]
      if (!cell?.formula) continue
      const next = rewrite(cell.formula)
      if (next !== cell.formula) {
        cell.formula = next
        delete cell.result
        delete cell.display
      }
    }
    if (sheet.dataValidations) {
      for (const validation of Object.values(sheet.dataValidations)) {
        const formulae = (validation as { formulae?: unknown[] } | null)?.formulae
        if (!Array.isArray(formulae)) continue
        formulae.forEach((formula, index) => {
          if (typeof formula === 'string' && !formula.startsWith('"')) formulae[index] = rewrite(formula.replace(/^=/, ''))
        })
      }
    }
    for (const block of (sheet.conditionalFormattings || []) as Array<{ rules?: Array<{ formulae?: unknown[] }> }>) {
      for (const rule of block?.rules || []) {
        if (!Array.isArray(rule.formulae)) continue
        rule.formulae.forEach((formula, index) => {
          if (typeof formula === 'string') rule.formulae![index] = rewrite(formula.replace(/^=/, ''))
        })
      }
    }
    // Sparkline sources ("Sheet1!B2:F2") read through the sheet name like any formula; a changed
    // source also changes the group's signature, so the file writer regenerates its XML.
    for (const group of sheet.sparklineGroups || []) {
      for (const sparkline of group.sparklines) {
        if (typeof sparkline.source !== 'string' || !sparkline.source) continue
        const next = rewrite(sparkline.source.replace(/^=/, ''))
        if (next !== sparkline.source) sparkline.source = next
      }
    }
  }
  for (const name of workbook.definedNames || []) {
    if (Array.isArray(name.ranges)) name.ranges = name.ranges.map((range) => rewrite(range.replace(/^=/, '')))
    if (typeof name.ref === 'string') name.ref = rewrite(name.ref.replace(/^=/, ''))
  }
  for (const name of workbook.metadata?.definedNames || []) {
    if (typeof name.ranges === 'string') name.ranges = rewrite(name.ranges.replace(/^=/, ''))
    if (typeof name.formula === 'string') name.formula = rewrite(name.formula.replace(/^=/, ''))
  }
}

/** Shift sheet-scoped defined names (by sheet position) after sheets are inserted or removed. */
function shiftSheetScopedNames(workbook: WorkbookModel, fromIndex: number, delta: number): void {
  for (const definedName of workbook.definedNames || []) {
    const extended = definedName as DefinedName & { localSheetRefId?: string }
    if (extended.localSheetRefId) continue
    if (Number.isInteger(definedName.localSheetIndex) && definedName.localSheetIndex! >= fromIndex) definedName.localSheetIndex! += delta
    if (Number.isInteger(definedName.localSheetId) && definedName.localSheetId! >= fromIndex) definedName.localSheetId! += delta
  }
  for (const definedName of workbook.metadata?.definedNames || []) {
    if (Number.isInteger(definedName.localSheetId) && definedName.localSheetId! >= fromIndex) definedName.localSheetId! += delta
  }
}

/**
 * Remove the sheet at `index` from the workbook's sheet-scoped names: names that belonged to it
 * go, and names of the sheets after it keep pointing at their own sheet. Mutates `workbook`.
 */
export function removeSheetScopedNames(workbook: WorkbookModel, index: number): void {
  const belongs = (value: number | undefined) => Number.isInteger(value) && value === index
  if (workbook.definedNames) {
    workbook.definedNames = workbook.definedNames.filter((name) => {
      const extended = name as DefinedName & { localSheetRefId?: string }
      return Boolean(extended.localSheetRefId) || !(belongs(name.localSheetIndex) || (name.localSheetIndex === undefined && belongs(name.localSheetId)))
    })
  }
  if (workbook.metadata?.definedNames) workbook.metadata.definedNames = workbook.metadata.definedNames.filter((name) => !belongs(name.localSheetId))
  shiftSheetScopedNames(workbook, index + 1, -1)
}

export interface SheetCopyOptions {
  /** Id of the new sheet. */
  id: string
  /** Name of the copy; defaults to "<name> copy", "<name> copy 2", … */
  name?: string
  /** Ids for the copied tables, charts, pictures and pivot tables. */
  makeId?: (prefix: string) => string
}

function defaultCopyName(workbook: WorkbookModel, sourceName: string): string {
  const taken = new Set(workbook.sheets.map((sheet) => sheet.name.toLocaleLowerCase()))
  for (let suffix = 1; suffix < 10_000; suffix += 1) {
    const ending = suffix === 1 ? ' copy' : ` copy ${suffix}`
    const name = `${sourceName.slice(0, 31 - ending.length)}${ending}`
    if (!taken.has(name.toLocaleLowerCase())) return name
  }
  throw new SheetStructureError('No free sheet name was found for the copy.', 'INVALID_OPERATION')
}

/**
 * Excel's Duplicate sheet, as a new sheet object (the workbook is not changed; place it with
 * {@link insertSheetCopy}). The copy's own references to its source sheet (formulas,
 * validation and conditional formats, sparklines, charts) point at the copy, as in Excel; its
 * tables get workbook-unique names; and the import bookkeeping that ties a sheet, chart or
 * table to parts of the original file is dropped, so saving writes the copy as a new sheet.
 */
export function createSheetCopy(workbook: WorkbookModel, sourceSheetId: string, options: SheetCopyOptions): SheetData {
  const source = workbook.sheets.find((sheet) => sheet.id === sourceSheetId)
  if (!source) throw new SheetStructureError(`Worksheet ${sourceSheetId} was not found.`, 'SHEET_NOT_FOUND')
  const makeId = options.makeId ?? ((prefix: string) => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`)
  const name = (options.name ?? defaultCopyName(workbook, source.name)).slice(0, 31)
  if (!validSheetName(name) || workbook.sheets.some((sheet) => sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())) {
    throw new SheetStructureError(`The sheet name “${name}” is not available.`, 'INVALID_OPERATION')
  }
  const copy = clone(source)
  copy.id = options.id
  copy.name = name
  copy.state = 'visible'
  delete copy.sourceWorksheetId
  delete copy.sourceSheetName
  delete copy.sourceSheetIndex
  const selfRename = (formula: string) => renameSheetInFormula(formula, source.name, name)
  rewriteWorkbookFormulaText({ ...workbook, sheets: [copy], definedNames: [], metadata: {} }, selfRename)
  renameTablesInCopiedSheet(workbook, copy, makeId)
  if (copy.charts?.length) {
    copy.charts = (renameChartSheetReferences(copy.charts, source.name, name) || []).map((chart) => {
      const next = { ...chart, id: makeId('chart') }
      delete next.sourcePart
      delete next.sourceInfo
      return next
    })
  }
  if (copy.images?.length) copy.images = copy.images.map((image) => ({ ...image, id: makeId('image') }))
  if (copy.pivots?.length) copy.pivots = copy.pivots.map((pivot) => ({ ...pivot, id: makeId('pivot') }))
  return copy
}

/**
 * Put a sheet made by {@link createSheetCopy} right after its source, keep sheet-scoped names
 * on their own sheets, and give the copy its own copies of the source's sheet-scoped names (as
 * Excel does). Mutates `workbook`, so it can run on an immer draft.
 */
export function insertSheetCopy(workbook: WorkbookModel, copy: SheetData, sourceSheetId: string): void {
  const sourceIndex = workbook.sheets.findIndex((sheet) => sheet.id === sourceSheetId)
  if (sourceIndex < 0) throw new SheetStructureError(`Worksheet ${sourceSheetId} was not found.`, 'SHEET_NOT_FOUND')
  const sourceName = workbook.sheets[sourceIndex].name
  const insertionIndex = sourceIndex + 1
  workbook.sheets.splice(insertionIndex, 0, copy)
  shiftSheetScopedNames(workbook, insertionIndex, 1)
  const rename = (text: string) => renameSheetInFormula(text.replace(/^=/, ''), sourceName, copy.name)
  if (workbook.definedNames) {
    const local = workbook.definedNames.filter((name) => {
      const extended = name as DefinedName & { localSheetRefId?: string }
      return !extended.localSheetRefId && (name.localSheetIndex === sourceIndex || (name.localSheetIndex === undefined && name.localSheetId === sourceIndex))
    })
    for (const name of local) {
      const copied: DefinedName = { ...name, ranges: (name.ranges || []).map(rename) }
      if (typeof name.ref === 'string') copied.ref = rename(name.ref)
      if (name.localSheetIndex !== undefined) copied.localSheetIndex = insertionIndex
      if (name.localSheetId !== undefined) copied.localSheetId = insertionIndex
      workbook.definedNames.push(copied)
    }
  }
  const metadataNames = workbook.metadata?.definedNames
  if (metadataNames) {
    for (const name of metadataNames.filter((item) => item.localSheetId === sourceIndex)) {
      metadataNames.push({
        ...name,
        ...(typeof name.ranges === 'string' ? { ranges: rename(name.ranges) } : {}),
        ...(typeof name.formula === 'string' ? { formula: rename(name.formula) } : {}),
        localSheetId: insertionIndex,
      })
    }
  }
}

function validSheetName(name: string): boolean {
  return Boolean(name.trim()) && name.length <= 31 && !/[\\/?*:[\]]/.test(name)
}

function nextSheetNumber(workbook: WorkbookModel): number {
  let number = 1
  const names = new Set(workbook.sheets.map((sheet) => sheet.name.toLocaleLowerCase()))
  while (names.has(`sheet${number}`.toLocaleLowerCase())) number += 1
  return number
}

/** Insert a conventional blank sheet without mutating the supplied workbook. */
export function insertBlankSheet(
  workbook: WorkbookModel,
  options: InsertBlankSheetOptions = {},
): InsertBlankSheetResult {
  const next = clone(workbook)
  const number = nextSheetNumber(next)
  const name = (options.name ?? `Sheet${number}`).trim()
  if (!validSheetName(name)) {
    throw new SheetStructureError('Sheet names must be 1–31 characters and cannot contain \\, /, ?, *, [, ], or :.', 'INVALID_OPERATION')
  }
  if (next.sheets.some((sheet) => sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())) {
    throw new SheetStructureError(`A worksheet named “${name}” already exists.`, 'INVALID_OPERATION')
  }
  let id = options.id?.trim() || `sheet-${number}`
  if (!id || next.sheets.some((sheet) => sheet.id === id)) {
    let suffix = number
    do id = `sheet-${suffix++}`
    while (next.sheets.some((sheet) => sheet.id === id))
  }
  const rowCount = options.rowCount ?? 200
  const colCount = options.colCount ?? 40
  if (!Number.isSafeInteger(rowCount) || rowCount < 1 || rowCount > MAX_SHEET_ROWS
    || !Number.isSafeInteger(colCount) || colCount < 1 || colCount > MAX_SHEET_COLUMNS) {
    throw new SheetStructureError('The blank sheet dimensions are outside the XLSX worksheet limits.', 'INVALID_OPERATION')
  }
  const sheet: SheetData = {
    id,
    name,
    state: 'visible',
    rowCount,
    colCount,
    cells: {},
    merges: [],
    colWidths: {},
    rowHeights: {},
    frozen: { rows: 0, columns: 0 },
  }
  const afterIndex = options.afterSheetId === undefined
    ? next.sheets.length - 1
    : next.sheets.findIndex((candidate) => candidate.id === options.afterSheetId)
  if (options.afterSheetId !== undefined && afterIndex < 0) {
    throw new SheetStructureError(`Worksheet ${options.afterSheetId} was not found.`, 'SHEET_NOT_FOUND')
  }
  const insertionIndex = afterIndex + 1
  next.sheets.splice(insertionIndex, 0, sheet)
  for (const definedName of next.definedNames || []) {
    const extended = definedName as DefinedName & { localSheetRefId?: string }
    if (extended.localSheetRefId) continue
    if (Number.isInteger(definedName.localSheetIndex) && definedName.localSheetIndex! >= insertionIndex) {
      definedName.localSheetIndex! += 1
    }
    if (Number.isInteger(definedName.localSheetId) && definedName.localSheetId! >= insertionIndex) {
      definedName.localSheetId! += 1
    }
  }
  for (const definedName of next.metadata?.definedNames || []) {
    if (Number.isInteger(definedName.localSheetId) && definedName.localSheetId! >= insertionIndex) {
      definedName.localSheetId! += 1
    }
  }
  if (options.activate !== false) next.activeSheetId = id
  return { workbook: next, sheetId: id }
}

// ---------------------------------------------------------------------------------------------
// Drag-and-drop: move or copy a block of cells (Excel's drag of the selection border)
// ---------------------------------------------------------------------------------------------

export type CellBlockTransferMode = 'move' | 'copy'

export interface CellBlockTransfer {
  sheetId: string
  /** The dragged block, 0-based. */
  source: Bounds
  /** Where its top-left cell lands. */
  destination: GridCoordinate
  /** 'move' re-points references to the block (cut + paste); 'copy' shifts relative ones (copy + paste). */
  mode: CellBlockTransferMode
}

export interface CellBlockTransferPlan {
  /** Where the block lands, 0-based. */
  target: Bounds
  /** Why the drop cannot be done (Excel's wording); nothing is changed. */
  error?: string
  /** The landing area already holds data (other than the dragged cells): Excel asks first. */
  overwrites: boolean
  /** Dropped where it started: nothing to do. */
  unchanged: boolean
}

const CELL_KEY = /^([A-Z]{1,3})([1-9]\d*)$/

function cellKeyCoord(address: string): GridCoordinate | null {
  const match = CELL_KEY.exec(address)
  if (!match) return null
  let col = 0
  for (const character of match[1]) col = col * 26 + character.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, col: col - 1 }
}

function blockCellKey(row: number, col: number): string {
  return `${columnNumberToLabel(col + 1)}${row + 1}`
}

function blockLabel(bounds: Bounds): string {
  const start = blockCellKey(bounds.top, bounds.left)
  const end = blockCellKey(bounds.bottom, bounds.right)
  return start === end ? start : `${start}:${end}`
}

function withinBlock(bounds: Bounds, row: number, col: number) {
  return row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right
}

function blocksIntersect(a: Bounds, b: Bounds) {
  return a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left
}

function blockContains(outer: Bounds, inner: Bounds) {
  return inner.top >= outer.top && inner.bottom <= outer.bottom && inner.left >= outer.left && inner.right <= outer.right
}

/** Overlaps the area without lying inside it: the drop would split it. */
function straddles(region: Bounds, area: Bounds) {
  return blocksIntersect(region, area) && !blockContains(area, region)
}

function shiftBlock(bounds: Bounds, rowDelta: number, colDelta: number): Bounds {
  return { top: bounds.top + rowDelta, bottom: bounds.bottom + rowDelta, left: bounds.left + colDelta, right: bounds.right + colDelta }
}

function hasBlockContent(cell: CellData | undefined) {
  return Boolean(cell && (cell.formula || (cell.value !== undefined && cell.value !== null && cell.value !== '')))
}

function blockRange(range: string): Bounds | null {
  const text = range.replace(/\$/g, '').toUpperCase().trim()
  const columns = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(text)
  const [first, second = first] = (columns ? `${columns[1]}1:${columns[2]}${MAX_SHEET_ROWS}` : text).split(':')
  const start = cellKeyCoord(first)
  const end = cellKeyCoord(second)
  if (!start || !end) return null
  return { top: Math.min(start.row, end.row), bottom: Math.max(start.row, end.row), left: Math.min(start.col, end.col), right: Math.max(start.col, end.col) }
}

/** Array formula ranges and dynamic-array spills on a sheet, 0-based. */
function arrayRegions(sheet: SheetData): Bounds[] {
  const regions: Bounds[] = []
  for (const [address, cell] of Object.entries(sheet.cells)) {
    if (!cell) continue
    if (cell.formulaRange) {
      const range = blockRange(cell.formulaRange)
      if (range) regions.push(range)
    }
    if (cell.arrayMember) {
      const anchor = cellKeyCoord(cell.arrayMember.replace(/\$/g, '').toUpperCase())
      const member = cellKeyCoord(address)
      if (anchor && member) {
        regions.push({ top: Math.min(anchor.row, member.row), bottom: Math.max(anchor.row, member.row), left: Math.min(anchor.col, member.col), right: Math.max(anchor.col, member.col) })
      }
    }
  }
  return regions
}

/** Where a drop lands and whether it can be done, without changing anything. */
export function planCellBlockTransfer(workbook: WorkbookModel, transfer: CellBlockTransfer): CellBlockTransferPlan {
  const { source, destination, mode } = transfer
  const height = source.bottom - source.top + 1
  const width = source.right - source.left + 1
  const target: Bounds = { top: destination.row, left: destination.col, bottom: destination.row + height - 1, right: destination.col + width - 1 }
  const plan: CellBlockTransferPlan = { target, overwrites: false, unchanged: target.top === source.top && target.left === source.left }
  const sheet = workbook.sheets.find((item) => item.id === transfer.sheetId)
  if (!sheet) return { ...plan, error: 'The worksheet was not found.' }
  if (target.top < 0 || target.left < 0 || target.bottom >= MAX_SHEET_ROWS || target.right >= MAX_SHEET_COLUMNS) {
    return { ...plan, error: 'The cells would land outside the sheet.' }
  }
  if (plan.unchanged) return plan
  const areas = mode === 'move' ? [source, target] : [target]
  for (const range of sheet.merges || []) {
    const merged = blockRange(range)
    if (merged && areas.some((area) => straddles(merged, area))) return { ...plan, error: "We can't do that to a merged cell. Include the whole merged cell, or unmerge it first." }
  }
  if (arrayRegions(sheet).some((region) => areas.some((area) => straddles(region, area)))) {
    return { ...plan, error: "You can't change part of an array. Move or copy the whole array range." }
  }
  for (const pivot of sheet.pivots || []) {
    if (!pivot.extent) continue
    const block = { top: pivot.anchor.row, left: pivot.anchor.col, bottom: pivot.anchor.row + pivot.extent.rows - 1, right: pivot.anchor.col + pivot.extent.cols - 1 }
    if (areas.some((area) => blocksIntersect(block, area))) return { ...plan, error: "Pivot table cells can't be moved or overwritten. Change the pivot table in its editor." }
  }
  for (const table of sheet.tables || []) {
    const region = blockRange(table.ref)
    if (!region || !table.headerRow) continue
    const header = { ...region, bottom: region.top }
    if (areas.some((area) => blocksIntersect(header, area))) return { ...plan, error: `This would change the header row of the table ${table.displayName || table.name}. Move cells below its headers instead.` }
  }
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const coord = cellKeyCoord(address)
    if (!coord || !withinBlock(target, coord.row, coord.col) || !hasBlockContent(cell)) continue
    if (mode === 'move' && withinBlock(source, coord.row, coord.col)) continue
    plan.overwrites = true
    break
  }
  return plan
}

/** References that lay wholly inside cells a move overwrote become #REF!, as in Excel. */
function invalidateOverwritten(formula: string, formulaSheet: string, sheetName: string, overwritten: Bounds, source: Bounds): string {
  const text = `=${formula}`
  const wanted = sheetName.toLocaleLowerCase()
  let output = ''
  let last = 0
  for (const token of tokenizeFormulaText(text)) {
    const reference = token.reference
    if (token.kind !== 'reference' || !reference) continue
    if ((reference.sheet ?? formulaSheet).toLocaleLowerCase() !== wanted) continue
    if (reference.bottom - reference.top >= MAX_SHEET_ROWS - 1 || reference.right - reference.left >= MAX_SHEET_COLUMNS - 1) continue
    const bounds = { top: reference.top, bottom: reference.bottom, left: reference.left, right: reference.right }
    if (!blockContains(overwritten, bounds) || blockContains(source, bounds)) continue
    output += text.slice(last, token.start) + '#REF!'
    last = token.end
  }
  if (!last) return formula
  return (output + text.slice(last)).slice(1)
}

/** The validation rules over a block, cut into the pieces that lie inside it. */
function validationPieces(validations: Record<string, unknown> | undefined, block: Bounds): Array<{ bounds: Bounds; rule: DataValidationModel }> {
  const pieces: Array<{ bounds: Bounds; rule: DataValidationModel }> = []
  for (const [key, rule] of Object.entries(validations || {})) {
    if (!rule || typeof rule !== 'object') continue
    for (const part of key.trim().split(/\s+/)) {
      const bounds = blockRange(part)
      if (!bounds || !blocksIntersect(bounds, block)) continue
      pieces.push({
        bounds: { top: Math.max(bounds.top, block.top), bottom: Math.min(bounds.bottom, block.bottom), left: Math.max(bounds.left, block.left), right: Math.min(bounds.right, block.right) },
        rule: rule as DataValidationModel,
      })
    }
  }
  return pieces
}

/** A cell at its new place: array ranges and spill anchors follow; it leaves any shared formula group. */
function relocatedBlockCell(cell: CellData, rowDelta: number, colDelta: number): CellData {
  const next = { ...cell }
  if (next.formulaRange) {
    const range = blockRange(next.formulaRange)
    if (range) next.formulaRange = blockLabel(shiftBlock(range, rowDelta, colDelta))
  }
  if (next.arrayMember) {
    const anchor = cellKeyCoord(next.arrayMember.replace(/\$/g, '').toUpperCase())
    if (anchor) next.arrayMember = blockCellKey(anchor.row + rowDelta, anchor.col + colDelta)
  }
  // A shared formula's clones each carry their own text; a moved cell leaves the group.
  if (next.formulaType === 'shared') delete next.formulaType
  delete (next as CellData & { sharedFormulaMaster?: string }).sharedFormulaMaster
  return next
}

/**
 * Drop a dragged block (on a draft, inside mutateWorkbook, so it is one undo step). A move works
 * like cut + paste: the cells, merges and validation travel, every formula that pointed into the
 * block follows it, and references to cells it overwrote become #REF!. A copy works like copy +
 * paste: relative references shift by the distance moved. Returns the plan; when it has an
 * error, or the block did not move, the workbook is left unchanged.
 */
export function transferCellBlock(workbook: WorkbookModel, transfer: CellBlockTransfer): CellBlockTransferPlan {
  const plan = planCellBlockTransfer(workbook, transfer)
  if (plan.error || plan.unchanged) return plan
  const sheet = workbook.sheets.find((item) => item.id === transfer.sheetId)!
  const { source, mode } = transfer
  const { target } = plan
  const rowDelta = target.top - source.top
  const colDelta = target.left - source.left
  const landed: Array<[string, CellData]> = []
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const coord = cellKeyCoord(address)
    if (!coord || !cell || !withinBlock(source, coord.row, coord.col)) continue
    const key = blockCellKey(coord.row + rowDelta, coord.col + colDelta)
    if (mode === 'move') {
      landed.push([key, relocatedBlockCell(cell, rowDelta, colDelta)])
      continue
    }
    // A new top-level cell sharing its unchanged parts (styles, notes), as copy + paste does;
    // no structuredClone, since this runs on an immer draft.
    const copy = relocatedBlockCell(cell, rowDelta, colDelta)
    if (copy.formula) {
      copy.formula = shiftFormulaReferences(copy.formula.replace(/^=/, ''), rowDelta, colDelta)
      delete copy.result
      delete copy.resultType
      delete copy.display
    }
    landed.push([key, copy])
  }
  for (const address of Object.keys(sheet.cells)) {
    const coord = cellKeyCoord(address)
    if (coord && (withinBlock(target, coord.row, coord.col) || (mode === 'move' && withinBlock(source, coord.row, coord.col)))) delete sheet.cells[address]
  }
  for (const [address, cell] of landed) sheet.cells[address] = cell

  if (mode === 'move') {
    // Shared formulas elsewhere whose master moved keep their own text and leave the group.
    for (const cell of Object.values(sheet.cells)) {
      const extended = cell as CellData & { sharedFormulaMaster?: string }
      if (!extended?.sharedFormulaMaster) continue
      const master = cellKeyCoord(extended.sharedFormulaMaster.replace(/\$/g, '').toUpperCase())
      if (master && withinBlock(source, master.row, master.col)) {
        delete extended.sharedFormulaMaster
        if (extended.formulaType === 'shared') delete extended.formulaType
      }
    }
    for (const formulaSheet of workbook.sheets) {
      for (const address of Object.keys(formulaSheet.cells)) {
        const cell = formulaSheet.cells[address]
        if (!cell?.formula) continue
        const invalidated = invalidateOverwritten(cell.formula, formulaSheet.name, sheet.name, target, source)
        const formula = moveReferencesInFormula(invalidated, { formulaSheet: formulaSheet.name, sourceSheet: sheet.name, rect: source, rowDelta, colDelta })
        if (formula !== cell.formula) formulaSheet.cells[address] = { ...cell, formula }
      }
    }
  }

  // Merges travel with the block (or are copied); merges the block lands on give way.
  const carried: string[] = []
  const kept = (sheet.merges || []).filter((range) => {
    const merged = blockRange(range)
    if (!merged) return true
    const inSource = blockContains(source, merged)
    if (inSource) carried.push(blockLabel(shiftBlock(merged, rowDelta, colDelta)))
    if (mode === 'move' && inSource) return false
    return !blockContains(target, merged)
  })
  sheet.merges = [...kept, ...carried]

  // Validation travels (or is copied) the way the cells do.
  if (sheet.dataValidations && Object.keys(sheet.dataValidations).length) {
    const pieces = validationPieces(sheet.dataValidations, source)
    let validations: Record<string, unknown> = sheet.dataValidations
    if (mode === 'move') validations = withValidation(validations, source, null)
    validations = withValidation(validations, target, null)
    for (const piece of pieces) validations = withValidation(validations, shiftBlock(piece.bounds, rowDelta, colDelta), piece.rule)
    sheet.dataValidations = validations
  }

  sheet.rowCount = Math.max(sheet.rowCount, target.bottom + 1)
  sheet.colCount = Math.max(sheet.colCount, target.right + 1)
  return plan
}
