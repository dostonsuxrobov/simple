/**
 * Excel tables (ListObjects) as editing operations on the workbook model: creation, the
 * totals row, header/column-name sync, auto-expansion when typing next to a table,
 * calculated columns, renaming (with structured-reference rewriting) and conversion back
 * to a plain range. Functions mutate the sheet/workbook they are given, so the app can
 * call them on an immer draft.
 */
import type { CellBorderSide, CellData, SheetData, SheetTable, WorkbookModel } from '../spreadsheet-types'
import { moveReferencesInFormula, tokenizeFormulaText } from './formula-editing'
import { parseStructuredSpecifier, shiftFormulaReferences } from './formulas'
import { tableCellPaint } from './table-styles'

export interface TableBounds {
  top: number
  bottom: number
  left: number
  right: number
}

export interface TableRegions extends TableBounds {
  /** Header row, or null when the table has none. */
  header: number | null
  dataTop: number
  dataBottom: number
  /** Totals row, or null when it is off. */
  totals: number | null
}

const RANGE = /^\s*\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?\s*$/

function columnIndex(label: string) {
  let value = 0
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value - 1
}

function columnLabel(index: number) {
  let label = ''
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) label = String.fromCharCode(65 + ((value - 1) % 26)) + label
  return label
}

export function cellAddress(row: number, col: number) {
  return `${columnLabel(col)}${row + 1}`
}

export function parseTableRef(ref: string | undefined): TableBounds | null {
  const match = RANGE.exec(String(ref || ''))
  if (!match) return null
  const first = { row: Number(match[2]) - 1, col: columnIndex(match[1]) }
  const second = match[3] ? { row: Number(match[4]) - 1, col: columnIndex(match[3]) } : first
  return {
    top: Math.min(first.row, second.row),
    bottom: Math.max(first.row, second.row),
    left: Math.min(first.col, second.col),
    right: Math.max(first.col, second.col),
  }
}

export function formatTableRef(bounds: TableBounds) {
  return `${cellAddress(bounds.top, bounds.left)}:${cellAddress(bounds.bottom, bounds.right)}`
}

export function tableRegions(table: SheetTable): TableRegions | null {
  const bounds = parseTableRef(table.ref)
  if (!bounds) return null
  const header = table.headerRow === false ? null : bounds.top
  const totals = table.totalsRow ? bounds.bottom : null
  return {
    ...bounds,
    header,
    dataTop: header === null ? bounds.top : bounds.top + 1,
    dataBottom: totals === null ? bounds.bottom : bounds.bottom - 1,
    totals,
  }
}

export function tableContaining(sheet: Pick<SheetData, 'tables'>, row: number, col: number): SheetTable | undefined {
  return (sheet.tables || []).find((table) => {
    const bounds = parseTableRef(table.ref)
    return Boolean(bounds && row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right)
  })
}

function isBlank(cell: CellData | undefined) {
  return !cell || ((cell.value === undefined || cell.value === null || cell.value === '') && !cell.formula)
}

/** The text a header cell contributes as a column name (Excel converts headers to text). */
export function headerText(cell: CellData | undefined): string {
  if (!cell) return ''
  const value = cell.formula ? (cell.result ?? cell.value) : cell.value
  if (value === undefined || value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return String(value).replace(/[\r\n]+/g, ' ').trim()
}

/** Excel's naming: blanks become ColumnN, repeats get a numeric suffix (Amount, Amount2). */
export function uniqueColumnNames(raw: string[]): string[] {
  const used = new Set<string>()
  const result: string[] = []
  raw.forEach((text, index) => {
    let base = text.trim() || `Column${index + 1}`
    let name = base
    for (let suffix = 2; used.has(name.toLocaleLowerCase()); suffix += 1) {
      if (!text.trim()) { base = `Column${index + suffix}`; name = base; continue }
      name = `${base}${suffix}`
    }
    used.add(name.toLocaleLowerCase())
    result.push(name)
  })
  return result
}

function workbookTables(workbook: Pick<WorkbookModel, 'sheets'>) {
  return workbook.sheets.flatMap((sheet) => (sheet.tables || []).map((table) => ({ sheet, table })))
}

export function findTable(workbook: Pick<WorkbookModel, 'sheets'>, id: string) {
  return workbookTables(workbook).find((entry) => entry.table.id === id)
}

export function nextTableName(workbook: Pick<WorkbookModel, 'sheets' | 'definedNames'>, base = 'Table') {
  const taken = new Set([
    ...workbookTables(workbook).flatMap(({ table }) => [table.name, table.displayName || table.name].map((name) => name.toLocaleLowerCase())),
    ...(workbook.definedNames || []).map((name) => String(name.name).toLocaleLowerCase()),
  ])
  for (let index = 1; ; index += 1) if (!taken.has(`${base}${index}`.toLocaleLowerCase())) return `${base}${index}`
}

/** Excel's table-name rules; returns an error message, or null when the name is usable. */
export function validateTableName(workbook: Pick<WorkbookModel, 'sheets' | 'definedNames'>, name: string, exceptId?: string): string | null {
  const trimmed = name.trim()
  if (!trimmed) return 'Enter a table name.'
  if (trimmed.length > 255) return 'Table names can have at most 255 characters.'
  if (!/^[A-Za-z_\\À-￿][A-Za-z0-9_.\\À-￿]*$/.test(trimmed)) return 'Table names start with a letter or underscore and use only letters, numbers, periods and underscores.'
  if (/^[A-Za-z]{1,3}\d+$/.test(trimmed) || /^[Rr]\d*[Cc]\d*$/.test(trimmed) || /^[RrCc]$/.test(trimmed)) return 'A table name cannot look like a cell reference.'
  const wanted = trimmed.toLocaleLowerCase()
  const clash = workbookTables(workbook).some(({ table }) => table.id !== exceptId && (table.name.toLocaleLowerCase() === wanted || (table.displayName || '').toLocaleLowerCase() === wanted))
  if (clash || (workbook.definedNames || []).some((item) => String(item.name).toLocaleLowerCase() === wanted)) return `The name "${trimmed}" is already used in this workbook.`
  return null
}

function overlaps(a: TableBounds, b: TableBounds) {
  return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom
}

/** Why a range cannot become a table, or null. */
export function tableRangeError(sheet: Pick<SheetData, 'tables' | 'merges'>, bounds: TableBounds, exceptId?: string): string | null {
  if (bounds.bottom - bounds.top < 0 || bounds.right - bounds.left < 0) return 'Select a range for the table.'
  for (const table of sheet.tables || []) {
    if (table.id === exceptId) continue
    const other = parseTableRef(table.ref)
    if (other && overlaps(bounds, other)) return 'A table cannot overlap another table.'
  }
  for (const merge of sheet.merges || []) {
    const range = parseTableRef(merge)
    if (range && overlaps(bounds, range)) return 'A table cannot contain merged cells. Unmerge them first.'
  }
  return null
}

/**
 * Insert `count` blank rows' worth of cells at `row` inside columns left..right (Excel's
 * "Shift cells down"), keeping formula references to the moved block pointed at it.
 */
export function shiftCellsDown(workbook: Pick<WorkbookModel, 'sheets'>, sheet: SheetData, left: number, right: number, row: number, count = 1) {
  const moved: Array<[string, CellData]> = []
  let lastRow = row - 1
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const match = /^([A-Z]+)(\d+)$/.exec(address)
    if (!match) continue
    const col = columnIndex(match[1])
    const cellRow = Number(match[2]) - 1
    if (col < left || col > right || cellRow < row) continue
    moved.push([cellAddress(cellRow + count, col), cell])
    delete sheet.cells[address]
    lastRow = Math.max(lastRow, cellRow)
  }
  for (const [address, cell] of moved) sheet.cells[address] = cell
  if (lastRow >= row) {
    const rect = { top: row, bottom: lastRow, left, right }
    for (const other of workbook.sheets) {
      for (const cell of Object.values(other.cells)) {
        if (cell.formula) cell.formula = moveReferencesInFormula(cell.formula, { formulaSheet: other.name, sourceSheet: sheet.name, rect, rowDelta: count, colDelta: 0 })
      }
    }
  }
  sheet.rowCount = Math.max(sheet.rowCount, lastRow + count + 1)
}

export interface CreateTableOptions {
  hasHeaders: boolean
  style?: string
  name?: string
  id?: string
}

export function createTable(workbook: WorkbookModel, sheet: SheetData, range: TableBounds, options: CreateTableOptions): SheetTable | string {
  const bounds = { ...range }
  const error = tableRangeError(sheet, bounds)
  if (error) return error
  const name = options.name?.trim() || nextTableName(workbook)
  const nameError = validateTableName(workbook, name)
  if (nameError) return nameError
  if (!options.hasHeaders) {
    // Excel inserts a header row above the data, shifting the table's columns down.
    shiftCellsDown(workbook, sheet, bounds.left, bounds.right, bounds.top, 1)
    bounds.bottom += 1
  }
  const raw = Array.from({ length: bounds.right - bounds.left + 1 }, (_, offset) => (
    options.hasHeaders ? headerText(sheet.cells[cellAddress(bounds.top, bounds.left + offset)]) : ''
  ))
  const names = uniqueColumnNames(raw)
  names.forEach((columnName, offset) => writeHeaderCell(sheet, bounds.top, bounds.left + offset, columnName))
  if (bounds.bottom === bounds.top) {
    // A table always has at least one data row.
    bounds.bottom += 1
    sheet.rowCount = Math.max(sheet.rowCount, bounds.bottom + 1)
  }
  const table: SheetTable = {
    id: options.id || `table-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`,
    name,
    displayName: name,
    ref: formatTableRef(bounds),
    headerRow: true,
    totalsRow: false,
    columns: names.map((columnName) => ({ name: columnName })),
    style: { theme: options.style || 'TableStyleMedium2', showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
    showFilterButton: true,
  }
  sheet.tables = [...(sheet.tables || []), table]
  return table
}

function writeHeaderCell(sheet: SheetData, row: number, col: number, text: string) {
  const address = cellAddress(row, col)
  const previous = sheet.cells[address] || {}
  const next: CellData = { ...previous, value: text }
  delete next.formula
  delete next.result
  delete next.resultType
  delete next.display
  delete next.richText
  delete next.type
  sheet.cells[address] = next
}

// ---- Totals row -------------------------------------------------------------------------------

export const TOTAL_FUNCTIONS = [
  { id: 'none', label: 'None', code: 0 },
  { id: 'average', label: 'Average', code: 101 },
  { id: 'count', label: 'Count', code: 103 },
  { id: 'countNums', label: 'Count Numbers', code: 102 },
  { id: 'max', label: 'Max', code: 104 },
  { id: 'min', label: 'Min', code: 105 },
  { id: 'sum', label: 'Sum', code: 109 },
  { id: 'stdDev', label: 'StdDev', code: 107 },
  { id: 'var', label: 'Var', code: 110 },
] as const

export type TotalFunctionId = typeof TOTAL_FUNCTIONS[number]['id']

/** Column names inside brackets escape [ ] # ' with an apostrophe. */
export function escapeStructuredName(name: string) {
  return name.replace(/(['#[\]])/g, "'$1")
}

export function totalsFormula(tableName: string, columnName: string, fn: TotalFunctionId): string | null {
  const entry = TOTAL_FUNCTIONS.find((item) => item.id === fn)
  if (!entry || !entry.code) return null
  return `SUBTOTAL(${entry.code},${tableName}[${escapeStructuredName(columnName)}])`
}

function columnIsNumeric(sheet: SheetData, col: number, top: number, bottom: number) {
  let numbers = 0
  let others = 0
  for (let row = top; row <= bottom && numbers + others < 200; row += 1) {
    const cell = sheet.cells[cellAddress(row, col)]
    if (isBlank(cell)) continue
    const value = cell!.formula ? cell!.result : cell!.value
    if (typeof value === 'number' || cell!.formula) numbers += 1
    else others += 1
  }
  return numbers > 0 && numbers >= others
}

export function setTotalsRow(workbook: WorkbookModel, sheet: SheetData, table: SheetTable, on: boolean) {
  const regions = tableRegions(table)
  if (!regions || Boolean(table.totalsRow) === on) return
  if (on) {
    const row = regions.bottom + 1
    const occupied = Array.from({ length: regions.right - regions.left + 1 }, (_, offset) => sheet.cells[cellAddress(row, regions.left + offset)]).some((cell) => !isBlank(cell))
    if (occupied) shiftCellsDown(workbook, sheet, regions.left, regions.right, row, 1)
    const remembered = table.columns.some((column) => column.totalsRowFunction || column.totalsRowLabel)
    table.columns = table.columns.map((column, index) => {
      if (remembered) return column
      const last = index === table.columns.length - 1
      if (index === 0 && !last) return { ...column, totalsRowLabel: 'Total' }
      if (last) return { ...column, totalsRowFunction: columnIsNumeric(sheet, regions.left + index, regions.dataTop, regions.dataBottom) ? 'sum' : 'count' }
      return column
    })
    table.totalsRow = true
    table.ref = formatTableRef({ ...regions, bottom: row })
    sheet.rowCount = Math.max(sheet.rowCount, row + 1)
    writeTotalsCells(sheet, table)
  } else {
    for (let col = regions.left; col <= regions.right; col += 1) {
      const address = cellAddress(regions.bottom, col)
      const cell = sheet.cells[address]
      if (!cell) continue
      const kept: CellData = { ...cell }
      delete kept.value
      delete kept.formula
      delete kept.result
      delete kept.resultType
      delete kept.display
      if (kept.style || kept.numFmt || kept.note || kept.hyperlink) sheet.cells[address] = kept
      else delete sheet.cells[address]
    }
    table.totalsRow = false
    table.ref = formatTableRef({ ...regions, bottom: regions.bottom - 1 })
  }
}

/** Write the totals-row cells from the columns' totals functions and labels. */
export function writeTotalsCells(sheet: SheetData, table: SheetTable) {
  const regions = tableRegions(table)
  if (!regions || regions.totals === null) return
  table.columns.forEach((column, index) => {
    const address = cellAddress(regions.totals!, regions.left + index)
    const previous = sheet.cells[address] || {}
    const next: CellData = { ...previous }
    delete next.value
    delete next.formula
    delete next.result
    delete next.resultType
    delete next.display
    const fn = column.totalsRowFunction as TotalFunctionId | 'custom' | undefined
    if (fn === 'custom' && column.totalsRowFormula) next.formula = column.totalsRowFormula.replace(/^=/, '')
    else if (fn && fn !== 'custom') {
      const formula = totalsFormula(table.name, column.name, fn)
      if (formula) next.formula = formula
    } else if (column.totalsRowLabel) next.value = column.totalsRowLabel
    if (next.value !== undefined || next.formula || next.style || next.numFmt || next.note) sheet.cells[address] = next
    else delete sheet.cells[address]
  })
}

export function setTotalsFunction(sheet: SheetData, table: SheetTable, index: number, fn: TotalFunctionId) {
  const column = table.columns[index]
  if (!column) return
  const next = { ...column }
  delete next.totalsRowFormula
  delete next.totalsRowLabel
  if (fn === 'none') delete next.totalsRowFunction
  else next.totalsRowFunction = fn
  table.columns = table.columns.map((item, position) => (position === index ? next : item))
  writeTotalsCells(sheet, table)
}

// ---- Structured references --------------------------------------------------------------------

export interface StructuredRename {
  /** Renamed tables, old name -> new name (case-insensitive match). */
  tables?: Map<string, string>
  /** Renamed columns per table (keyed by lower-case table name, after any table rename). */
  columns?: Map<string, Map<string, string>>
}

function renameInSpecifier(specifier: string, columns: Map<string, string> | undefined) {
  if (!columns?.size || !specifier.startsWith('[') || !specifier.endsWith(']')) return specifier
  // Column items are renamed; #Specials and the @ this-row marker are kept as written.
  const rename = (text: string) => {
    const at = text.startsWith('@')
    const body = at ? text.slice(1) : text
    if (!body.trim() || body.trim().startsWith('#')) return text
    const renamed = columns.get(body.replace(/'(.)/g, '$1').trim().toLocaleLowerCase())
    return renamed === undefined ? text : `${at ? '@' : ''}${escapeStructuredName(renamed)}`
  }
  const inner = specifier.slice(1, -1)
  if (!inner.includes('[')) return `[${rename(inner)}]`
  let output = ''
  for (let index = 0; index < inner.length;) {
    if (inner[index] !== '[') { output += inner[index]; index += 1; continue }
    let end = index + 1
    let text = ''
    while (end < inner.length && inner[end] !== ']') {
      if (inner[end] === "'" && end + 1 < inner.length) { text += inner[end] + inner[end + 1]; end += 2; continue }
      text += inner[end]
      end += 1
    }
    output += `[${rename(text)}]`
    index = end + 1
  }
  return `[${output}]`
}

/**
 * Rewrite table and column names inside structured references. `hostTable` is the table the
 * formula sits in, which unqualified references like [@Amount] refer to.
 */
export function renameStructuredReferences(formula: string, rename: StructuredRename, hostTable?: string): string {
  if (!formula.includes('[')) return formula
  const text = `=${formula}`
  let output = ''
  let last = 0
  for (const token of tokenizeFormulaText(text)) {
    if (token.kind !== 'name' || !token.text.includes('[')) continue
    const bracket = token.text.indexOf('[')
    const tableName = token.text.slice(0, bracket)
    const specifier = token.text.slice(bracket)
    const renamedTable = tableName ? rename.tables?.get(tableName.toLocaleLowerCase()) : undefined
    const effective = (renamedTable ?? (tableName || hostTable || '')).toLocaleLowerCase()
    const nextSpecifier = renameInSpecifier(specifier, rename.columns?.get(effective))
    const nextText = `${renamedTable ?? tableName}${nextSpecifier}`
    if (nextText === token.text) continue
    output += text.slice(last, token.start) + nextText
    last = token.end
  }
  if (!last) return formula
  return (output + text.slice(last)).slice(1)
}

/** Apply a rename to every formula in the workbook (cells, names, validations, formats). */
export function applyStructuredRename(workbook: WorkbookModel, rename: StructuredRename) {
  const hasWork = Boolean(rename.tables?.size || [...(rename.columns?.values() || [])].some((map) => map.size))
  if (!hasWork) return
  for (const sheet of workbook.sheets) {
    const tables = (sheet.tables || []).map((table) => ({ table, bounds: parseTableRef(table.ref) }))
    for (const [address, cell] of Object.entries(sheet.cells)) {
      if (!cell.formula || !cell.formula.includes('[')) continue
      const match = /^([A-Z]+)(\d+)$/.exec(address)
      const row = match ? Number(match[2]) - 1 : -1
      const col = match ? columnIndex(match[1]) : -1
      const host = tables.find(({ bounds }) => bounds && row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right)?.table
      const next = renameStructuredReferences(cell.formula, rename, host?.name)
      if (next !== cell.formula) cell.formula = next
    }
    for (const table of sheet.tables || []) {
      table.columns = table.columns.map((column) => (column.totalsRowFormula
        ? { ...column, totalsRowFormula: renameStructuredReferences(column.totalsRowFormula, rename, table.name) }
        : column))
    }
    if (sheet.dataValidations) {
      for (const [key, validation] of Object.entries(sheet.dataValidations)) {
        const record = validation as { formulae?: unknown[] }
        if (Array.isArray(record?.formulae)) record.formulae = record.formulae.map((item) => (typeof item === 'string' ? renameStructuredReferences(item, rename) : item))
        sheet.dataValidations[key] = record
      }
    }
  }
  for (const name of workbook.definedNames || []) {
    if (typeof name.ref === 'string') name.ref = renameStructuredReferences(name.ref, rename)
  }
}

export function renameTable(workbook: WorkbookModel, tableId: string, name: string): string | null {
  const entry = findTable(workbook, tableId)
  if (!entry) return 'The table no longer exists.'
  const trimmed = name.trim()
  if (trimmed === entry.table.name) return null
  const error = validateTableName(workbook, trimmed, tableId)
  if (error) return error
  const old = entry.table.name
  entry.table.name = trimmed
  entry.table.displayName = trimmed
  applyStructuredRename(workbook, { tables: new Map([[old.toLocaleLowerCase(), trimmed]]) })
  if (entry.table.totalsRow) writeTotalsCells(entry.sheet, entry.table)
  return null
}

/**
 * Bring each table's column names in line with its header cells after an edit, renaming the
 * columns in structured references as Excel does. Returns true when anything changed.
 */
export function syncTableHeaders(workbook: WorkbookModel, sheet: SheetData, changed?: Iterable<string>): boolean {
  let any = false
  const changedSet = changed ? new Set(changed) : null
  const columnRenames = new Map<string, Map<string, string>>()
  for (const table of sheet.tables || []) {
    const regions = tableRegions(table)
    if (!regions || regions.header === null) continue
    const width = regions.right - regions.left + 1
    if (changedSet && !Array.from({ length: width }, (_, offset) => cellAddress(regions.header!, regions.left + offset)).some((address) => changedSet.has(address))) continue
    const raw = Array.from({ length: width }, (_, offset) => headerText(sheet.cells[cellAddress(regions.header!, regions.left + offset)]))
    const names = uniqueColumnNames(raw)
    const renames = new Map<string, string>()
    names.forEach((name, offset) => {
      const cell = sheet.cells[cellAddress(regions.header!, regions.left + offset)]
      if (raw[offset] !== name || cell?.formula || typeof cell?.value !== 'string') writeHeaderCell(sheet, regions.header!, regions.left + offset, name)
      const previous = table.columns[offset]?.name
      if (previous !== undefined && previous !== name) renames.set(previous.toLocaleLowerCase(), name)
    })
    if (renames.size || table.columns.length !== width) {
      table.columns = names.map((name, offset) => ({ ...(table.columns[offset] || {}), name }))
      if (renames.size) columnRenames.set(table.name.toLocaleLowerCase(), renames)
      any = true
    }
  }
  if (columnRenames.size) {
    applyStructuredRename(workbook, { columns: columnRenames })
    for (const table of sheet.tables || []) if (table.totalsRow && columnRenames.has(table.name.toLocaleLowerCase())) writeTotalsCells(sheet, table)
  }
  return any
}

// ---- Calculated columns and auto-expansion ----------------------------------------------------

/** The formula a calculated column repeats (as written in its first data row), or null. */
export function calculatedColumnFormula(sheet: SheetData, table: SheetTable, index: number, ignoreRow?: number): string | null {
  const regions = tableRegions(table)
  if (!regions) return null
  const col = regions.left + index
  let anchor: { formula: string; row: number } | null = null
  for (let row = regions.dataTop; row <= regions.dataBottom; row += 1) {
    if (row === ignoreRow) continue
    const cell = sheet.cells[cellAddress(row, col)]
    if (!cell?.formula) return null
    if (!anchor) { anchor = { formula: cell.formula, row }; continue }
    if (shiftFormulaReferences(anchor.formula, row - anchor.row, 0) !== cell.formula) return null
  }
  return anchor ? shiftFormulaReferences(anchor.formula, regions.dataTop - anchor.row, 0) : null
}

function columnIsEmpty(sheet: SheetData, col: number, top: number, bottom: number, ignoreRow: number) {
  for (let row = top; row <= bottom; row += 1) if (row !== ignoreRow && !isBlank(sheet.cells[cellAddress(row, col)])) return false
  return true
}

/**
 * Excel's calculated columns: a formula entered in a table column that is otherwise empty
 * (or already one consistent calculated column) fills the whole column.
 */
export function fillCalculatedColumn(sheet: SheetData, table: SheetTable, row: number, col: number): boolean {
  const regions = tableRegions(table)
  if (!regions || row < regions.dataTop || row > regions.dataBottom || col < regions.left || col > regions.right) return false
  const entered = sheet.cells[cellAddress(row, col)]
  if (!entered?.formula || regions.dataBottom === regions.dataTop) return false
  if (entered.formulaType === 'array' || entered.formulaRange) return false
  const index = col - regions.left
  const eligible = columnIsEmpty(sheet, col, regions.dataTop, regions.dataBottom, row) || calculatedColumnFormula(sheet, table, index, row) !== null
  if (!eligible) return false
  for (let target = regions.dataTop; target <= regions.dataBottom; target += 1) {
    if (target === row) continue
    const address = cellAddress(target, col)
    const previous = sheet.cells[address] || {}
    const next: CellData = { ...previous, formula: shiftFormulaReferences(entered.formula, target - row, 0) }
    delete next.value
    delete next.result
    delete next.resultType
    delete next.display
    if (entered.numFmt && !next.numFmt) next.numFmt = entered.numFmt
    sheet.cells[address] = next
  }
  return true
}

/**
 * Typing directly below a table (without a totals row) or directly right of it extends the
 * table, carrying calculated columns into the new row. Returns true when the table grew.
 */
export function expandTableForEntry(sheet: SheetData, row: number, col: number): boolean {
  const cell = sheet.cells[cellAddress(row, col)]
  if (isBlank(cell)) return false
  if (tableContaining(sheet, row, col)) return false
  for (const table of sheet.tables || []) {
    const regions = tableRegions(table)
    if (!regions) continue
    if (regions.totals === null && row === regions.bottom + 1 && col >= regions.left && col <= regions.right) {
      const calculated = table.columns.map((_, index) => calculatedColumnFormula(sheet, table, index))
      const nextBounds = { ...regions, bottom: row }
      if (tableRangeError(sheet, nextBounds, table.id)) return false
      table.ref = formatTableRef(nextBounds)
      calculated.forEach((formula, index) => {
        const address = cellAddress(row, regions.left + index)
        if (!formula || !isBlank(sheet.cells[address])) return
        const source = sheet.cells[cellAddress(regions.dataTop, regions.left + index)]
        sheet.cells[address] = { ...(sheet.cells[address] || {}), formula: shiftFormulaReferences(formula, row - regions.dataTop, 0), ...(source?.numFmt ? { numFmt: source.numFmt } : {}) }
      })
      return true
    }
    if (col === regions.right + 1 && row >= regions.top && row <= regions.bottom) {
      const nextBounds = { ...regions, right: col }
      if (tableRangeError(sheet, nextBounds, table.id)) return false
      table.ref = formatTableRef(nextBounds)
      const typed = regions.header !== null && row === regions.header ? headerText(cell) : ''
      const added = uniqueColumnNames([...table.columns.map((column) => column.name), typed])[table.columns.length]
      if (regions.header !== null) writeHeaderCell(sheet, regions.header, col, added)
      table.columns = [...table.columns, { name: added }]
      return true
    }
  }
  return false
}

// ---- Resize and convert -----------------------------------------------------------------------

/** Resize a table (the header row must stay in place, as in Excel). */
export function resizeTable(sheet: SheetData, table: SheetTable, bounds: TableBounds): string | null {
  const regions = tableRegions(table)
  if (!regions) return 'The table range is invalid.'
  if (bounds.top !== regions.top) return 'The header row must stay in the same row.'
  const minimumRows = (table.headerRow === false ? 0 : 1) + (table.totalsRow ? 1 : 0) + 1
  if (bounds.bottom - bounds.top + 1 < minimumRows) return 'A table needs at least one data row.'
  const error = tableRangeError(sheet, bounds, table.id)
  if (error) return error
  const width = bounds.right - bounds.left + 1
  const previousColumns = table.columns
  const offset = bounds.left - regions.left
  table.ref = formatTableRef(bounds)
  if (table.headerRow !== false) {
    const raw = Array.from({ length: width }, (_, index) => headerText(sheet.cells[cellAddress(bounds.top, bounds.left + index)]))
    const names = uniqueColumnNames(raw)
    names.forEach((name, index) => writeHeaderCell(sheet, bounds.top, bounds.left + index, name))
    table.columns = names.map((name, index) => ({ ...(previousColumns[index + offset] || {}), name }))
  } else {
    table.columns = uniqueColumnNames(Array.from({ length: width }, (_, index) => previousColumns[index + offset]?.name || '')).map((name, index) => ({ ...(previousColumns[index + offset] || {}), name }))
  }
  if (table.totalsRow) writeTotalsCells(sheet, table)
  return null
}

function structuredToRange(tokenText: string, host: { row: number; col: number }, workbook: WorkbookModel, formulaSheet: SheetData, onlyTable?: string): string | null {
  const bracket = tokenText.indexOf('[')
  const tableName = tokenText.slice(0, bracket)
  const specifier = parseStructuredSpecifier(tokenText.slice(bracket))
  if (!specifier) return null
  let entry: { sheet: SheetData; table: SheetTable } | undefined
  if (tableName) entry = workbookTables(workbook).find(({ table }) => table.name.toLocaleLowerCase() === tableName.toLocaleLowerCase())
  else {
    const table = tableContaining(formulaSheet, host.row, host.col)
    entry = table ? { sheet: formulaSheet, table } : undefined
  }
  if (!entry || (onlyTable && entry.table.id !== onlyTable)) return null
  const regions = tableRegions(entry.table)
  if (!regions) return null
  let first = regions.dataTop
  let last = regions.dataBottom
  if (specifier.specials.size) {
    const rows: Array<[number, number]> = []
    for (const special of specifier.specials) {
      if (special === '#all') rows.push([regions.top, regions.bottom])
      else if (special === '#data') rows.push([regions.dataTop, regions.dataBottom])
      else if (special === '#headers' && regions.header !== null) rows.push([regions.header, regions.header])
      else if (special === '#totals' && regions.totals !== null) rows.push([regions.totals, regions.totals])
      else return null
    }
    first = Math.min(...rows.map(([start]) => start))
    last = Math.max(...rows.map(([, end]) => end))
  }
  if (specifier.thisRow) { first = host.row; last = host.row }
  const indexOf = (name: string) => entry!.table.columns.findIndex((column) => column.name.toLocaleLowerCase() === name.toLocaleLowerCase())
  const indices: number[] = []
  if (specifier.span) indices.push(indexOf(specifier.span[0]), indexOf(specifier.span[1]))
  for (const column of specifier.columns) indices.push(indexOf(column))
  if (indices.some((index) => index < 0)) return null
  const left = indices.length ? regions.left + Math.min(...indices) : regions.left
  const right = indices.length ? regions.left + Math.max(...indices) : regions.right
  const absolute = !specifier.thisRow
  const ref = (row: number, col: number) => (absolute ? `$${columnLabel(col)}$${row + 1}` : `$${columnLabel(col)}${row + 1}`)
  const range = first === last && left === right ? ref(first, left) : `${ref(first, left)}:${ref(last, right)}`
  if (entry.sheet.id === formulaSheet.id) return range
  const sheetName = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(entry.sheet.name) ? entry.sheet.name : `'${entry.sheet.name.replace(/'/g, "''")}'`
  return `${sheetName}!${range}`
}

/** Replace structured references to one table (or every table) with A1 references. */
export function structuredReferencesToA1(formula: string, host: { row: number; col: number }, workbook: WorkbookModel, formulaSheet: SheetData, onlyTable?: string): string {
  if (!formula.includes('[')) return formula
  const text = `=${formula}`
  let output = ''
  let last = 0
  for (const token of tokenizeFormulaText(text)) {
    if (token.kind !== 'name' || !token.text.includes('[')) continue
    const replacement = structuredToRange(token.text, host, workbook, formulaSheet, onlyTable)
    if (!replacement) continue
    output += text.slice(last, token.start) + replacement
    last = token.end
  }
  if (!last) return formula
  return (output + text.slice(last)).slice(1)
}

function cssColorArgb(color: string) {
  const hex = color.replace('#', '').toUpperCase()
  return /^[0-9A-F]{6}$/.test(hex) ? `FF${hex}` : undefined
}

function borderSide(css: string | undefined): CellBorderSide | undefined {
  const match = /^(\d+)px\s+(solid|double|dashed|dotted)\s+(#[0-9a-f]{6})$/i.exec(css || '')
  if (!match) return undefined
  const width = Number(match[1])
  const kind = match[2].toLowerCase()
  const style = kind === 'double' ? 'double' : width >= 3 ? 'thick' : width === 2 ? 'medium' : kind === 'solid' ? 'thin' : kind
  const argb = cssColorArgb(match[3])
  return argb ? { style, color: { argb } } : { style }
}

/** Write a table style's look into the cells as direct formatting (kept by Convert to Range). */
export function bakeTableFormatting(sheet: SheetData, table: SheetTable, theme?: readonly string[]) {
  const regions = tableRegions(table)
  if (!regions) return
  const entry = { table, top: regions.top, bottom: regions.bottom, left: regions.left, right: regions.right }
  for (let row = regions.top; row <= regions.bottom; row += 1) {
    for (let col = regions.left; col <= regions.right; col += 1) {
      const paint = tableCellPaint(entry, row, col, theme)
      if (!paint) continue
      const address = cellAddress(row, col)
      const cell: CellData = { ...(sheet.cells[address] || {}) }
      const style = { ...(cell.style || {}) }
      if (paint.fill && !style.fill?.pattern && !style.fill?.type) {
        const argb = cssColorArgb(paint.fill)
        if (argb) style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } }
      }
      const fontColor = paint.color && !style.font?.color ? cssColorArgb(paint.color) : undefined
      const addBold = Boolean(paint.bold) && style.font?.bold === undefined
      if (fontColor || addBold) style.font = { ...(style.font || {}), ...(fontColor ? { color: { argb: fontColor } } : {}), ...(addBold ? { bold: true } : {}) }
      const border = { ...(style.border || {}) }
      for (const [side, css] of [['top', paint.borderTop], ['bottom', paint.borderBottom], ['left', paint.borderLeft], ['right', paint.borderRight]] as const) {
        const converted = borderSide(css)
        if (converted && !border[side]) border[side] = converted
      }
      if (Object.keys(border).length) style.border = border
      if (Object.keys(style).length) {
        cell.style = style
        sheet.cells[address] = cell
      }
    }
  }
}

/** Excel's "Convert to Range": the cells and their look stay, references become A1 references. */
export function convertTableToRange(workbook: WorkbookModel, tableId: string, theme?: readonly string[]) {
  const entry = findTable(workbook, tableId)
  if (!entry) return
  bakeTableFormatting(entry.sheet, entry.table, theme)
  for (const sheet of workbook.sheets) {
    for (const [address, cell] of Object.entries(sheet.cells)) {
      if (!cell.formula || !cell.formula.includes('[')) continue
      const match = /^([A-Z]+)(\d+)$/.exec(address)
      if (!match) continue
      const next = structuredReferencesToA1(cell.formula, { row: Number(match[2]) - 1, col: columnIndex(match[1]) }, workbook, sheet, tableId)
      if (next !== cell.formula) cell.formula = next
    }
  }
  entry.sheet.tables = (entry.sheet.tables || []).filter((table) => table.id !== tableId)
}

/** Keep table geometry aligned with inserted/deleted rows or columns (0-based operation). */
export function transformTablesForStructure(sheet: SheetData, operation: { axis: 'row' | 'column'; kind: 'insert' | 'delete'; index: number; count: number; span?: { start: number; end: number } }) {
  if (!sheet.tables?.length) return
  const start = operation.index
  const end = operation.index + operation.count - 1
  sheet.tables = sheet.tables.flatMap((table) => {
    const regions = tableRegions(table)
    if (!regions) return [table]
    if (operation.span) {
      // Shifting cells only moves tables inside the shifted band (others were refused).
      const low = operation.axis === 'row' ? regions.left : regions.top
      const high = operation.axis === 'row' ? regions.right : regions.bottom
      if (low < operation.span.start || high > operation.span.end) return [table]
    }
    const next: SheetTable = { ...table, columns: [...table.columns] }
    if (operation.axis === 'row') {
      let { top, bottom } = regions
      if (operation.kind === 'insert') {
        if (start <= top) { top += operation.count; bottom += operation.count }
        else if (start <= bottom) bottom += operation.count
      } else {
        const removedBefore = Math.max(0, Math.min(end, top - 1) - start + 1)
        const removedInside = Math.max(0, Math.min(end, bottom) - Math.max(start, top) + 1)
        if (removedInside >= bottom - top + 1) return []
        if (regions.header !== null && start <= regions.header && end >= regions.header) next.headerRow = false
        if (regions.totals !== null && start <= regions.totals && end >= regions.totals) next.totalsRow = false
        top -= removedBefore
        bottom -= removedBefore + removedInside
      }
      next.ref = formatTableRef({ top, bottom, left: regions.left, right: regions.right })
      return [next]
    }
    let { left, right } = regions
    if (operation.kind === 'insert') {
      if (start <= left) { left += operation.count; right += operation.count }
      else if (start <= right) {
        right += operation.count
        const offset = start - regions.left
        const names = uniqueColumnNames([...table.columns.map((column) => column.name), ...Array.from({ length: operation.count }, () => '')])
        const added = names.slice(table.columns.length).map((name) => ({ name }))
        next.columns.splice(offset, 0, ...added)
        if (regions.header !== null) added.forEach((column, index) => writeHeaderCell(sheet, regions.header!, start + index, column.name))
      }
    } else {
      const removedBefore = Math.max(0, Math.min(end, left - 1) - start + 1)
      const low = Math.max(start, left)
      const high = Math.min(end, right)
      const removedInside = Math.max(0, high - low + 1)
      if (removedInside >= right - left + 1) return []
      if (removedInside) next.columns.splice(low - left, removedInside)
      left -= removedBefore
      right -= removedBefore + removedInside
    }
    next.ref = formatTableRef({ top: regions.top, bottom: regions.bottom, left, right })
    return [next]
  })
}
