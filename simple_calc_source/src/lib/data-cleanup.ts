/**
 * Data cleanup tools (pure): Remove Duplicates, Trim whitespace, Split text to columns,
 * Change case and Fill blanks from above. Each returns a CellChanges map for the app to apply
 * (see applyCellChanges) so one undo step covers the whole operation.
 */
import type { CellData, SheetData } from '../spreadsheet-types'
import {
  addressOf,
  boundsContain,
  boundsIntersect,
  cellHasContent,
  foldText,
  moveCell,
  parseAddress,
  parseDateText,
  parseNumberText,
  parseRange,
  scalarKind,
  serialFromYMD,
} from './data-tools-core'
import type { Bounds, CellChanges, DataHost, ShiftFormula } from './data-tools-core'

type CellSheet = Pick<SheetData, 'cells'> & Partial<Pick<SheetData, 'merges'>>

export type CleanupFailure = { ok: false; error: string }

function blocker(sheet: CellSheet, bounds: Bounds): string | null {
  for (const range of sheet.merges || []) {
    const merged = parseRange(range)
    if (merged && boundsIntersect(merged, bounds) && !(merged.top === merged.bottom && merged.left === merged.right)) {
      return 'This command cannot be used on a range that contains merged cells. Unmerge them first.'
    }
  }
  for (const [address, cell] of Object.entries(sheet.cells)) {
    if (!(cell.formulaType === 'array' || cell.dynamicFormula || (cell.formulaRange && cell.formulaType !== 'shared'))) continue
    const area = (cell.formulaRange && parseRange(cell.formulaRange)) || parseRange(address)
    if (area && boundsIntersect(area, bounds) && !boundsContain(bounds, area)) return "You can't change part of an array formula."
    if (area && boundsIntersect(area, bounds)) return 'This command cannot be used on a range that contains array or spilled formulas.'
  }
  return null
}

function plainCell(previous: CellData | undefined): CellData {
  const next: CellData = { ...(previous || {}) }
  delete next.formula
  delete next.formulaType
  delete next.formulaRange
  delete next.dynamicFormula
  delete next.result
  delete next.resultType
  delete next.display
  delete next.richText
  delete next.type
  delete next.value
  delete (next as CellData & { sharedFormulaMaster?: string }).sharedFormulaMaster
  return next
}

// ---------------------------------------------------------------------------------------------
// Remove duplicates
// ---------------------------------------------------------------------------------------------

export interface RemoveDuplicatesOptions {
  /** 0-based column offsets within the range compared for duplicates (default: all). */
  columns?: number[]
  hasHeader?: boolean
}

export type RemoveDuplicatesResult =
  | { ok: true; removed: number; remaining: number; changes: CellChanges; /** Absolute 0-based rows kept, in order. */ keptRows: number[] }
  | CleanupFailure

/**
 * Excel's Remove Duplicates: rows whose selected columns show the same text (case-insensitive,
 * as displayed) as an earlier row are removed; the remaining rows shift up inside the range and
 * the vacated rows at its bottom are cleared. Cells outside the range are never touched.
 */
export function removeDuplicates(sheet: CellSheet, bounds: Bounds, options: RemoveDuplicatesOptions, host: DataHost, shiftFormula?: ShiftFormula): RemoveDuplicatesResult {
  const width = bounds.right - bounds.left + 1
  const columns = (options.columns?.length ? options.columns : Array.from({ length: width }, (_, index) => index))
    .filter((column) => Number.isInteger(column) && column >= 0 && column < width)
  if (!columns.length) return { ok: false, error: 'Select at least one column.' }
  const problem = blocker(sheet, bounds)
  if (problem) return { ok: false, error: problem }
  const start = options.hasHeader ? bounds.top + 1 : bounds.top
  if (start > bounds.bottom) return { ok: true, removed: 0, remaining: 0, changes: {}, keptRows: [] }
  const seen = new Set<string>()
  const kept: number[] = []
  for (let row = start; row <= bounds.bottom; row += 1) {
    let key = ''
    for (let index = 0; index < columns.length; index += 1) {
      if (index) key += '\u0000'
      key += foldText(host.displayAt(row, bounds.left + columns[index]))
    }
    if (seen.has(key)) continue
    seen.add(key)
    kept.push(row)
  }
  const total = bounds.bottom - start + 1
  const removed = total - kept.length
  const changes: CellChanges = {}
  if (removed) {
    kept.forEach((source, index) => {
      const target = start + index
      if (source === target) return
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        const cell = sheet.cells[addressOf(source, col)]
        const targetAddress = addressOf(target, col)
        if (cell) changes[targetAddress] = moveCell(cell, target - source, 0, shiftFormula)
        else if (sheet.cells[targetAddress]) changes[targetAddress] = null
      }
    })
    for (let row = start + kept.length; row <= bounds.bottom; row += 1) {
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        const address = addressOf(row, col)
        if (sheet.cells[address]) changes[address] = null
      }
    }
  }
  return { ok: true, removed, remaining: kept.length, changes, keptRows: kept }
}

// ---------------------------------------------------------------------------------------------
// Trim whitespace / change case
// ---------------------------------------------------------------------------------------------

/** Sheets' "Trim whitespace": strip leading/trailing spaces and collapse inner runs (incl. NBSP). */
export function trimText(text: string): string {
  return text
    .split(/\r\n|\r|\n/)
    .map((line) => line.replace(/[ \t  -​　]+/g, ' ').trim())
    .join('\n')
    .replace(/^\n+|\n+$/g, '')
}

function mapTextCells(sheet: CellSheet, bounds: Bounds, transform: (text: string) => string): { changes: CellChanges; count: number } {
  const changes: CellChanges = {}
  let count = 0
  const visit = (address: string, cell: CellData) => {
    if (cell.formula || typeof cell.value !== 'string') return
    const next = transform(cell.value)
    if (next === cell.value) return
    const updated: CellData = { ...cell, value: next }
    delete updated.richText
    delete updated.display
    changes[address] = updated
    count += 1
  }
  const area = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
  const keys = Object.keys(sheet.cells)
  if (area > keys.length) {
    for (const address of keys) {
      const coord = parseAddress(address)
      if (coord && coord.row >= bounds.top && coord.row <= bounds.bottom && coord.col >= bounds.left && coord.col <= bounds.right) visit(address, sheet.cells[address])
    }
  } else {
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        const address = addressOf(row, col)
        const cell = sheet.cells[address]
        if (cell) visit(address, cell)
      }
    }
  }
  return { changes, count }
}

/** Trim whitespace in every text constant of the range (formulas are left alone). */
export function trimWhitespace(sheet: CellSheet, bounds: Bounds): { changes: CellChanges; count: number } {
  return mapTextCells(sheet, bounds, trimText)
}

export type TextCase = 'upper' | 'lower' | 'proper' | 'sentence'

/** Excel PROPER semantics: a letter following a non-letter is capitalised, others lowered. */
export function properCase(text: string): string {
  let result = ''
  let previousLetter = false
  for (const character of text) {
    const isLetter = character.toLocaleLowerCase() !== character.toLocaleUpperCase()
    result += isLetter ? (previousLetter ? character.toLocaleLowerCase() : character.toLocaleUpperCase()) : character
    previousLetter = isLetter
  }
  return result
}

export function sentenceCase(text: string): string {
  let capitalise = true
  let result = ''
  for (const character of text.toLocaleLowerCase()) {
    const isLetter = character.toLocaleLowerCase() !== character.toLocaleUpperCase()
    if (isLetter && capitalise) { result += character.toLocaleUpperCase(); capitalise = false }
    else result += character
    if (/[.!?\n]/.test(character)) capitalise = true
  }
  return result
}

export function changeCase(sheet: CellSheet, bounds: Bounds, mode: TextCase): { changes: CellChanges; count: number } {
  const transform = mode === 'upper' ? (text: string) => text.toLocaleUpperCase()
    : mode === 'lower' ? (text: string) => text.toLocaleLowerCase()
      : mode === 'proper' ? properCase : sentenceCase
  return mapTextCells(sheet, bounds, transform)
}

// ---------------------------------------------------------------------------------------------
// Fill blanks from above
// ---------------------------------------------------------------------------------------------

/**
 * Fill each empty cell of the range with the nearest filled cell above it in the same column
 * (within the range). Formulas are copied with relative references shifted; number formats
 * travel with the value; the blank cell keeps its own style when it has one.
 */
export function fillBlanksFromAbove(sheet: CellSheet, bounds: Bounds, shiftFormula?: ShiftFormula): { changes: CellChanges; count: number } {
  const changes: CellChanges = {}
  let count = 0
  for (let col = bounds.left; col <= bounds.right; col += 1) {
    let source: { row: number; cell: CellData } | null = null
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      const address = addressOf(row, col)
      const cell = sheet.cells[address]
      if (cellHasContent(cell)) { source = { row, cell: cell! }; continue }
      if (!source) continue
      const copied = moveCell(source.cell, row - source.row, 0, shiftFormula)
      const next: CellData = { ...copied }
      if (cell?.style) {
        next.style = { ...cell.style }
        const format = source.cell.numFmt || source.cell.style?.numFmt
        if (format && !cell.numFmt && !cell.style.numFmt) next.numFmt = format
      }
      delete next.note
      delete next.hyperlink
      delete next.hyperlinkTooltip
      if (cell?.note) next.note = cell.note
      changes[address] = next
      count += 1
    }
  }
  return { changes, count }
}

// ---------------------------------------------------------------------------------------------
// Split text to columns
// ---------------------------------------------------------------------------------------------

export type SplitColumnFormat = 'general' | 'text' | 'date-mdy' | 'date-dmy' | 'date-ymd' | 'skip'

export interface SplitOptions {
  mode: 'delimited' | 'fixed'
  delimiters?: { tab?: boolean; comma?: boolean; semicolon?: boolean; space?: boolean; other?: string }
  treatConsecutiveAsOne?: boolean
  /** Quote character that protects delimiters inside a field ('' for none). Default '"'. */
  textQualifier?: '"' | "'" | ''
  /** Fixed width: character positions where new columns start, e.g. [5, 12]. */
  breaks?: number[]
  /** Per output column; missing entries are 'general'. */
  columnFormats?: SplitColumnFormat[]
  /** Top-left of the output (default: the source's first cell). */
  destination?: { row: number; col: number }
  /** Trim spaces around each field (default true for delimited). */
  trimFields?: boolean
}

function delimiterSet(options: SplitOptions): Set<string> {
  const set = new Set<string>()
  const delimiters = options.delimiters || {}
  if (delimiters.tab) set.add('\t')
  if (delimiters.comma) set.add(',')
  if (delimiters.semicolon) set.add(';')
  if (delimiters.space) set.add(' ')
  if (delimiters.other) set.add(delimiters.other[0])
  return set
}

/** Split one line into fields according to the options. */
export function splitLine(text: string, options: SplitOptions): string[] {
  if (options.mode === 'fixed') {
    const breaks = [...new Set((options.breaks || []).filter((value) => Number.isInteger(value) && value > 0))].sort((a, b) => a - b)
    const fields: string[] = []
    let position = 0
    for (const breakAt of breaks) {
      fields.push(text.slice(position, breakAt))
      position = breakAt
    }
    fields.push(text.slice(position))
    return options.trimFields === false ? fields : fields.map((field) => field.trim())
  }
  const delimiters = delimiterSet(options)
  if (!delimiters.size) return [text]
  const qualifier = options.textQualifier === undefined ? '"' : options.textQualifier
  const fields: string[] = []
  let field = ''
  let quoted = false
  let fieldStart = true
  let lastWasDelimiter = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (quoted) {
      if (character === qualifier) {
        if (text[index + 1] === qualifier) { field += qualifier; index += 1 }
        else quoted = false
      } else field += character
      continue
    }
    if (qualifier && character === qualifier && fieldStart) { quoted = true; fieldStart = false; lastWasDelimiter = false; field = ''; continue }
    if (delimiters.has(character)) {
      if (options.treatConsecutiveAsOne && lastWasDelimiter) continue
      fields.push(field)
      field = ''
      fieldStart = true
      lastWasDelimiter = true
      continue
    }
    field += character
    // Spaces before an opening qualifier are tolerated ('a, "b, c"').
    if (character.trim() !== '') fieldStart = false
    lastWasDelimiter = false
  }
  fields.push(field)
  return options.trimFields === false ? fields : fields.map((item) => item.trim())
}

/** First `limit` lines split for the preview table, and the widest field count. */
export function previewSplit(lines: string[], options: SplitOptions, limit = 10): { rows: string[][]; columns: number } {
  const rows = lines.slice(0, limit).map((line) => splitLine(line, options))
  return { rows, columns: rows.reduce((max, row) => Math.max(max, row.length), 0) }
}

function parseOrderedDate(text: string, order: 'mdy' | 'dmy' | 'ymd'): number | null {
  const parts = text.trim().split(/[/.\-\s]+/)
  if (parts.length !== 3 || parts.some((part) => !/^\d+$/.test(part))) return parseDateText(text)
  const numbers = parts.map(Number)
  const [a, b, c] = numbers
  let year: number, month: number, day: number
  if (order === 'mdy') { month = a; day = b; year = c }
  else if (order === 'dmy') { day = a; month = b; year = c }
  else { year = a; month = b; day = c }
  if (year < 100) year += year < 30 ? 2000 : 1900
  const date = new Date(Date.UTC(year, month - 1, day))
  if (month < 1 || month > 12 || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null
  return serialFromYMD(year, month, day)
}

/** Convert one split field into a cell value per the column's data format. */
export function convertField(text: string, format: SplitColumnFormat): { value: CellData['value']; numFmt?: string; type?: string } {
  if (format === 'text') return { value: text, numFmt: '@' }
  if (text === '') return { value: undefined }
  if (format === 'date-mdy' || format === 'date-dmy' || format === 'date-ymd') {
    const serial = parseOrderedDate(text, format.slice(5) as 'mdy' | 'dmy' | 'ymd')
    return serial === null ? { value: text } : { value: serial, numFmt: 'm/d/yyyy', type: 'date' }
  }
  if (/^(true|false)$/i.test(text)) return { value: text.toLocaleLowerCase() === 'true' }
  const number = parseNumberText(text)
  if (number !== null) return { value: number, numFmt: /%\s*$/.test(text) ? '0%' : undefined }
  const serial = parseDateText(text)
  if (serial !== null) return { value: serial, numFmt: serial % 1 ? 'm/d/yyyy h:mm' : 'm/d/yyyy', type: 'date' }
  return { value: text }
}

export type SplitResult =
  | { ok: true; changes: CellChanges; /** Output columns written. */ columns: number; rows: number; /** Non-empty cells (other than the sources) that will be replaced. */ overwrites: number; target: Bounds }
  | CleanupFailure

/** Source lines of a single-column range (display text of each cell). */
export function splitSourceLines(bounds: Bounds, host: DataHost): string[] {
  const lines: string[] = []
  for (let row = bounds.top; row <= bounds.bottom; row += 1) {
    const value = host.valueAt(row, bounds.left)
    lines.push(typeof value === 'string' ? value : host.displayAt(row, bounds.left))
  }
  return lines
}

/**
 * Excel's Text to Columns for one column: each source line is split and written across from
 * `destination`. Destination cells keep their style; values are typed per column format.
 */
export function splitTextToColumns(sheet: CellSheet, source: Bounds, options: SplitOptions, host: DataHost): SplitResult {
  if (source.left !== source.right) return { ok: false, error: 'Text to Columns converts one column at a time. Select cells in a single column.' }
  const destination = options.destination || { row: source.top, col: source.left }
  const lines = splitSourceLines(source, host)
  const formats = options.columnFormats || []
  let columns = 0
  const outputs: Array<Array<{ value: CellData['value']; numFmt?: string; type?: string }>> = []
  for (const line of lines) {
    const fields = scalarKind(line) === 'blank' ? [] : splitLine(line, options)
    const row: Array<{ value: CellData['value']; numFmt?: string; type?: string }> = []
    fields.forEach((field, index) => {
      const format = formats[index] || 'general'
      if (format === 'skip') return
      row.push(convertField(field, format))
    })
    columns = Math.max(columns, row.length)
    outputs.push(row)
  }
  if (!columns) return { ok: false, error: 'There is no text to split in the selected cells.' }
  const target: Bounds = { top: destination.row, bottom: destination.row + lines.length - 1, left: destination.col, right: destination.col + columns - 1 }
  const problem = blocker(sheet, target)
  if (problem) return { ok: false, error: problem }
  const changes: CellChanges = {}
  let overwrites = 0
  const sourceAddresses = new Set(lines.map((_, index) => addressOf(source.top + index, source.left)))
  outputs.forEach((fields, rowIndex) => {
    const row = destination.row + rowIndex
    const sourceAddress = addressOf(source.top + rowIndex, source.left)
    for (let index = 0; index < columns; index += 1) {
      const address = addressOf(row, destination.col + index)
      const existing = sheet.cells[address]
      const field = fields[index]
      if (!field) {
        // A source cell that produced no field in this position becomes empty (Excel clears it).
        if (address === sourceAddress && existing) {
          const cleared = plainCell(existing)
          changes[address] = Object.keys(cleared).length ? cleared : null
        }
        continue
      }
      if (existing && cellHasContent(existing) && !sourceAddresses.has(address)) overwrites += 1
      const next = plainCell(existing)
      if (field.value !== undefined) next.value = field.value
      if (field.numFmt) next.numFmt = field.numFmt
      if (field.type) next.type = field.type
      changes[address] = Object.keys(next).length ? next : null
    }
  })
  // With a separate destination the source column is left as it was (Excel copies, not moves).
  return { ok: true, changes, columns, rows: lines.length, overwrites, target }
}

/** Guess delimiters from sample lines (tab > semicolon > comma > space), for the dialog's first view. */
export function guessDelimiters(lines: string[]): NonNullable<SplitOptions['delimiters']> {
  const sample = lines.filter((line) => line.trim()).slice(0, 50)
  const score = (character: string) => {
    if (!sample.length) return 0
    const counts = sample.map((line) => line.split(character).length - 1)
    const hits = counts.filter((count) => count > 0).length
    return hits / sample.length
  }
  for (const [key, character] of [['tab', '\t'], ['semicolon', ';'], ['comma', ','], ['space', ' ']] as const) {
    if (score(character) >= 0.6) return { [key]: true }
  }
  return { comma: true }
}
