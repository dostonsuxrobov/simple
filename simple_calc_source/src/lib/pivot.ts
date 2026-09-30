/**
 * Pivot tables: summarise a source range by row/column fields into a static block of cells
 * (Excel's tabular layout: one column per row field, subtotals after each group, grand
 * totals). The definition lives on the output sheet (`sheet.pivots`) and the block is
 * rewritten on refresh.
 */
import type { CellData, CellScalar, CellStyle, PivotAxisField, PivotDateGroup, PivotSummarize, PivotTableModel, PivotValueField, SheetData } from '../spreadsheet-types'

export interface PivotSourceCell {
  value: CellScalar
  /** Displayed text (used for labels). */
  text: string
  /** The cell's number format (the first data row's is used for value fields). */
  numFmt?: string
  isDate?: boolean
}

export interface PivotSource {
  headers: string[]
  records: PivotSourceCell[][]
}

export const PIVOT_SUMMARIES: Array<{ id: PivotSummarize; label: string }> = [
  { id: 'sum', label: 'Sum' },
  { id: 'count', label: 'Count' },
  { id: 'average', label: 'Average' },
  { id: 'max', label: 'Max' },
  { id: 'min', label: 'Min' },
  { id: 'product', label: 'Product' },
  { id: 'countNums', label: 'Count Numbers' },
  { id: 'countDistinct', label: 'Distinct Count' },
  { id: 'median', label: 'Median' },
  { id: 'stdDev', label: 'StdDev' },
  { id: 'stdDevp', label: 'StdDevp' },
  { id: 'var', label: 'Var' },
  { id: 'varp', label: 'Varp' },
]

export const PIVOT_DATE_GROUPS: Array<{ id: PivotDateGroup; label: string }> = [
  { id: 'year', label: 'Years' },
  { id: 'quarter', label: 'Quarters' },
  { id: 'month', label: 'Months' },
  { id: 'yearQuarter', label: 'Year-Quarter' },
  { id: 'yearMonth', label: 'Year-Month' },
  { id: 'day', label: 'Days' },
]

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

export function valueFieldLabel(value: PivotValueField) {
  if (value.label) return value.label
  const summary = PIVOT_SUMMARIES.find((item) => item.id === value.summarize)?.label || 'Sum'
  return `${summary} of ${value.field}`
}

interface Key {
  id: string
  label: string
  sort: number | string
  /** Blank sorts last, as in Excel. */
  blank?: boolean
}

function serialDate(serial: number) {
  // Excel 1900 system (serial 60 is the phantom 29 Feb 1900).
  const days = Math.floor(serial) - (serial > 60 ? 25569 : 25568)
  return new Date(days * 86_400_000)
}

function keyOf(cell: PivotSourceCell | undefined, group: PivotDateGroup | undefined): Key {
  const value = cell?.value
  if (value === null || value === undefined || value === '') return { id: 'blank', label: '(blank)', sort: '', blank: true }
  if (group && typeof value === 'number' && Number.isFinite(value)) {
    const date = serialDate(value)
    const year = date.getUTCFullYear()
    const month = date.getUTCMonth()
    const quarter = Math.floor(month / 3) + 1
    switch (group) {
      case 'year': return { id: `y${year}`, label: String(year), sort: year }
      case 'quarter': return { id: `q${quarter}`, label: `Qtr${quarter}`, sort: quarter }
      case 'month': return { id: `m${month}`, label: MONTHS[month], sort: month }
      case 'yearQuarter': return { id: `yq${year}-${quarter}`, label: `${year} Qtr${quarter}`, sort: year * 10 + quarter }
      case 'yearMonth': return { id: `ym${year}-${month}`, label: `${MONTHS[month]} ${year}`, sort: year * 100 + month }
      case 'day': return { id: `d${Math.floor(value)}`, label: cell?.text || String(value), sort: Math.floor(value) }
    }
  }
  if (typeof value === 'number') return { id: `n${value}`, label: cell?.text || String(value), sort: value }
  if (typeof value === 'boolean') return { id: `b${value}`, label: value ? 'TRUE' : 'FALSE', sort: value ? 'true' : 'false' }
  const text = String(value)
  return { id: `s${text.toLocaleLowerCase()}`, label: cell?.text || text, sort: text.toLocaleLowerCase() }
}

function compareKeys(a: Key, b: Key) {
  if (a.blank !== b.blank) return a.blank ? 1 : -1
  if (typeof a.sort === 'number' && typeof b.sort === 'number') return a.sort - b.sort
  if (typeof a.sort === 'number') return -1
  if (typeof b.sort === 'number') return 1
  return String(a.sort).localeCompare(String(b.sort), undefined, { numeric: true, sensitivity: 'base' })
}

class Accumulator {
  sum = 0
  count = 0
  counta = 0
  product = 1
  min = Infinity
  max = -Infinity
  sumSquares = 0
  numbers: number[] | null = null
  distinct: Set<string> | null = null

  constructor(summarize: PivotSummarize) {
    if (summarize === 'median') this.numbers = []
    if (summarize === 'countDistinct') this.distinct = new Set()
  }

  add(cell: PivotSourceCell | undefined) {
    const value = cell?.value
    if (value === null || value === undefined || value === '') return
    this.counta += 1
    this.distinct?.add(`${typeof value}:${typeof value === 'string' ? value.toLocaleLowerCase() : String(value)}`)
    if (typeof value !== 'number' || !Number.isFinite(value)) return
    this.count += 1
    this.sum += value
    this.product *= value
    this.sumSquares += value * value
    if (value < this.min) this.min = value
    if (value > this.max) this.max = value
    this.numbers?.push(value)
  }

  result(summarize: PivotSummarize): CellScalar {
    switch (summarize) {
      case 'sum': return this.sum
      case 'count': return this.counta
      case 'countNums': return this.count
      case 'countDistinct': return this.distinct?.size ?? 0
      case 'average': return this.count ? this.sum / this.count : '#DIV/0!'
      case 'max': return this.count ? this.max : 0
      case 'min': return this.count ? this.min : 0
      case 'product': return this.count ? this.product : 0
      case 'median': {
        const sorted = (this.numbers || []).slice().sort((a, b) => a - b)
        if (!sorted.length) return '#NUM!'
        const middle = sorted.length >> 1
        return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2
      }
      case 'var':
      case 'stdDev':
      case 'varp':
      case 'stdDevp': {
        const population = summarize === 'varp' || summarize === 'stdDevp'
        const denominator = population ? this.count : this.count - 1
        if (denominator <= 0) return '#DIV/0!'
        const variance = Math.max(0, (this.sumSquares - (this.sum * this.sum) / this.count) / denominator)
        return summarize.startsWith('std') ? Math.sqrt(variance) : variance
      }
    }
    return this.sum
  }
}

export type PivotCellKind = 'corner' | 'fieldHeader' | 'valueHeader' | 'columnKey' | 'rowKey' | 'value' | 'subtotalLabel' | 'subtotal' | 'totalLabel' | 'total' | 'blank'

export interface PivotOutputCell {
  value: CellScalar
  kind: PivotCellKind
  numFmt?: string
}

export interface PivotOutput {
  cells: PivotOutputCell[][]
  rows: number
  columns: number
  /** Why nothing could be summarised (missing fields), for the editor. */
  message?: string
}

interface TreeNode {
  key: Key
  path: string
  children: Map<string, TreeNode>
}

function fieldIndexes(source: PivotSource, fields: Array<{ field: string }>) {
  return fields.map((item) => source.headers.findIndex((header) => header.toLocaleLowerCase() === item.field.toLocaleLowerCase()))
}

/** Leaves of an axis tree in display order, with subtotal entries after each inner group. */
function axisEntries(root: TreeNode, fields: PivotAxisField[], subtotals: boolean) {
  const entries: Array<{ path: string; keys: Key[]; kind: 'leaf' | 'subtotal' }> = []
  const walk = (node: TreeNode, depth: number, keys: Key[]) => {
    const order = fields[depth]?.order === 'desc' ? -1 : 1
    const children = [...node.children.values()].sort((a, b) => order * compareKeys(a.key, b.key))
    for (const child of children) {
      const nextKeys = [...keys, child.key]
      if (depth === fields.length - 1) entries.push({ path: child.path, keys: nextKeys, kind: 'leaf' })
      else {
        walk(child, depth + 1, nextKeys)
        if (subtotals && fields[depth]?.showTotals !== false) entries.push({ path: child.path, keys: nextKeys, kind: 'subtotal' })
      }
    }
  }
  if (fields.length) walk(root, 0, [])
  return entries
}

export function computePivot(source: PivotSource, model: PivotTableModel): PivotOutput {
  const rowFields = model.rows
  const columnFields = model.columns
  const values = model.values
  const rowIndexes = fieldIndexes(source, rowFields)
  const columnIndexes = fieldIndexes(source, columnFields)
  const valueIndexes = fieldIndexes(source, values)
  const filterIndexes = fieldIndexes(source, model.filters)
  const missing = [...rowIndexes, ...columnIndexes, ...valueIndexes, ...filterIndexes].some((index) => index < 0)
  if (missing) return { cells: [[{ value: 'A field of this pivot table is no longer in the source range.', kind: 'corner' }]], rows: 1, columns: 1, message: 'missing field' }
  if (!rowFields.length && !columnFields.length && !values.length) {
    return { cells: [[{ value: 'Add rows, columns or values in the pivot table editor.', kind: 'corner' }]], rows: 1, columns: 1, message: 'empty' }
  }
  const excluded = model.filters.map((filter) => new Set(filter.exclude || []))

  const rowRoot: TreeNode = { key: { id: '', label: '', sort: '' }, path: '', children: new Map() }
  const columnRoot: TreeNode = { key: { id: '', label: '', sort: '' }, path: '', children: new Map() }
  const cells = new Map<string, Accumulator[]>()
  const accumulatorsAt = (key: string) => {
    let list = cells.get(key)
    if (!list) { list = values.map((value) => new Accumulator(value.summarize)); cells.set(key, list) }
    return list
  }
  const insert = (root: TreeNode, keys: Key[]) => {
    const paths = ['']
    let node = root
    for (const key of keys) {
      const path = `${node.path}\u0001${key.id}`
      let child = node.children.get(key.id)
      if (!child) { child = { key, path, children: new Map() }; node.children.set(key.id, child) }
      node = child
      paths.push(path)
    }
    return paths
  }

  for (const record of source.records) {
    let skip = false
    for (let index = 0; index < filterIndexes.length; index += 1) {
      if (excluded[index].size && excluded[index].has(keyOf(record[filterIndexes[index]], undefined).id)) { skip = true; break }
    }
    if (skip) continue
    const rowPaths = insert(rowRoot, rowIndexes.map((column, index) => keyOf(record[column], rowFields[index].dateGroup)))
    const columnPaths = insert(columnRoot, columnIndexes.map((column, index) => keyOf(record[column], columnFields[index].dateGroup)))
    for (const rowPath of rowPaths) {
      for (const columnPath of columnPaths) {
        const accumulators = accumulatorsAt(`${rowPath}\u0002${columnPath}`)
        for (let index = 0; index < values.length; index += 1) accumulators[index].add(record[valueIndexes[index]])
      }
    }
  }

  const subtotals = true
  const rowEntries = axisEntries(rowRoot, rowFields, subtotals)
  const columnEntries = axisEntries(columnRoot, columnFields, subtotals)
  const showRowGrand = model.showRowGrandTotal !== false && rowFields.length > 0
  const showColumnGrand = model.showColumnGrandTotal !== false && columnFields.length > 0
  const columnLeaves: Array<{ path: string; keys: Key[]; kind: 'leaf' | 'subtotal' | 'grand' }> = columnFields.length
    ? [...columnEntries, ...(showColumnGrand ? [{ path: '', keys: [], kind: 'grand' as const }] : [])]
    : [{ path: '', keys: [], kind: 'leaf' }]
  const rowLeaves: Array<{ path: string; keys: Key[]; kind: 'leaf' | 'subtotal' | 'grand' }> = rowFields.length
    ? [...rowEntries, ...(showRowGrand ? [{ path: '', keys: [], kind: 'grand' as const }] : [])]
    : [{ path: '', keys: [], kind: 'leaf' }]

  const numberFormatOf = (index: number) => {
    const value = values[index]
    const column = valueIndexes[index]
    if (value.showAs && value.showAs !== 'normal' && value.showAs !== 'runningTotal') return '0.00%'
    if (['count', 'countNums', 'countDistinct'].includes(value.summarize)) return '0'
    return source.records.find((record) => record[column]?.numFmt)?.[column]?.numFmt
  }
  const formats = values.map((_, index) => numberFormatOf(index))
  const raw = (rowPath: string, columnPath: string, index: number) => cells.get(`${rowPath}\u0002${columnPath}`)?.[index]?.result(values[index].summarize) ?? null
  const shown = (rowPath: string, columnPath: string, index: number, runningBase?: Map<string, number>) => {
    const value = raw(rowPath, columnPath, index)
    const mode = values[index].showAs || 'normal'
    if (mode === 'normal' || typeof value !== 'number') return value
    if (mode === 'runningTotal') {
      if (!runningBase) return value
      const next = (runningBase.get(`${columnPath}#${index}`) || 0) + value
      runningBase.set(`${columnPath}#${index}`, next)
      return next
    }
    const base = mode === 'percentOfGrandTotal' ? raw('', '', index) : mode === 'percentOfRowTotal' ? raw(rowPath, '', index) : raw('', columnPath, index)
    return typeof base === 'number' && base !== 0 ? value / base : '#DIV/0!'
  }

  const valueCount = Math.max(1, values.length)
  const labelColumns = Math.max(1, rowFields.length)
  const valueHeaderRow = values.length > 1 || !columnFields.length
  const dataColumns = columnLeaves.length * (values.length ? valueCount : 0)
  const width = labelColumns + dataColumns
  const grid: PivotOutputCell[][] = []
  const blankRow = () => Array.from({ length: width }, (): PivotOutputCell => ({ value: null, kind: 'blank' }))

  // ---- Header rows ----
  if (columnFields.length) {
    // Row 0: the value caption (single value) and the column field names.
    const top = blankRow()
    top[0] = { value: values.length === 1 ? valueFieldLabel(values[0]) : values.length ? 'Values' : null, kind: 'corner' }
    columnFields.forEach((field, index) => { if (labelColumns + index < width) top[labelColumns + index] = { value: field.field, kind: 'fieldHeader' } })
    grid.push(top)
    // One row per column level with its keys; the last level row also carries the row field names.
    for (let level = 0; level < columnFields.length; level += 1) {
      const row = blankRow()
      const last = level === columnFields.length - 1
      if (last && !valueHeaderRow) rowFields.forEach((field, index) => { row[index] = { value: field.field, kind: 'fieldHeader' } })
      columnLeaves.forEach((leaf, leafIndex) => {
        const start = labelColumns + leafIndex * valueCount
        let label: CellScalar = null
        let kind: PivotCellKind = 'columnKey'
        if (leaf.kind === 'grand') { if (level === 0) { label = values.length > 1 ? 'Total' : 'Grand Total'; kind = 'totalLabel' } }
        else if (leaf.kind === 'subtotal') { if (level === leaf.keys.length - 1) { label = `${leaf.keys[level].label} Total`; kind = 'subtotalLabel' } }
        else {
          const previous = columnLeaves[leafIndex - 1]
          const repeated = previous && previous.kind === 'leaf' && previous.keys.slice(0, level + 1).every((key, index) => key.id === leaf.keys[index]?.id)
          if (!repeated || level === leaf.keys.length - 1) label = leaf.keys[level]?.label ?? null
        }
        if (label !== null) row[start] = { value: label, kind }
      })
      grid.push(row)
    }
    if (valueHeaderRow) {
      const row = blankRow()
      rowFields.forEach((field, index) => { row[index] = { value: field.field, kind: 'fieldHeader' } })
      columnLeaves.forEach((leaf, leafIndex) => values.forEach((value, index) => {
        const prefix = leaf.kind === 'grand' ? 'Total ' : ''
        row[labelColumns + leafIndex * valueCount + index] = { value: `${prefix}${valueFieldLabel(value)}`, kind: 'valueHeader' }
      }))
      grid.push(row)
    }
  } else {
    const row = blankRow()
    if (rowFields.length) rowFields.forEach((field, index) => { row[index] = { value: field.field, kind: 'fieldHeader' } })
    else row[0] = { value: null, kind: 'corner' }
    values.forEach((value, index) => { row[labelColumns + index] = { value: valueFieldLabel(value), kind: 'valueHeader' } })
    grid.push(row)
  }

  // ---- Body rows ----
  const running = new Map<string, number>()
  rowLeaves.forEach((leaf, leafIndex) => {
    const row = blankRow()
    if (!rowFields.length) row[0] = { value: values.length ? 'Total' : null, kind: 'totalLabel' }
    else if (leaf.kind === 'grand') row[0] = { value: 'Grand Total', kind: 'totalLabel' }
    else if (leaf.kind === 'subtotal') row[leaf.keys.length - 1] = { value: `${leaf.keys[leaf.keys.length - 1].label} Total`, kind: 'subtotalLabel' }
    else {
      const previous = rowLeaves[leafIndex - 1]
      leaf.keys.forEach((key, level) => {
        const repeated = previous && previous.kind === 'leaf' && previous.keys.slice(0, level + 1).every((item, index) => item.id === leaf.keys[index]?.id)
        if (!repeated || level === leaf.keys.length - 1) row[level] = { value: key.label, kind: 'rowKey' }
      })
    }
    const bodyKind: PivotCellKind = leaf.kind === 'grand' ? 'total' : leaf.kind === 'subtotal' ? 'subtotal' : 'value'
    columnLeaves.forEach((columnLeaf, columnIndex) => values.forEach((_, index) => {
      const kind: PivotCellKind = columnLeaf.kind === 'grand' ? 'total' : columnLeaf.kind === 'subtotal' && bodyKind === 'value' ? 'subtotal' : bodyKind
      const value = shown(leaf.path, columnLeaf.path, index, leaf.kind === 'leaf' ? running : undefined)
      row[labelColumns + columnIndex * valueCount + index] = { value, kind, ...(formats[index] ? { numFmt: formats[index] } : {}) }
    }))
    grid.push(row)
  })

  return { cells: grid, rows: grid.length, columns: width }
}

// ---- Writing the block ---------------------------------------------------------------------------

function address(row: number, col: number) {
  let label = ''
  for (let value = col + 1; value > 0; value = Math.floor((value - 1) / 26)) label = String.fromCharCode(65 + ((value - 1) % 26)) + label
  return `${label}${row + 1}`
}

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } }
const THIN_BLUE = { style: 'thin', color: { argb: 'FF9BC2E6' } }

function styleFor(kind: PivotCellKind): CellStyle | undefined {
  switch (kind) {
    case 'corner':
    case 'fieldHeader':
    case 'valueHeader':
      return { font: { bold: true }, fill: HEADER_FILL, border: { bottom: THIN_BLUE } } as CellStyle
    case 'columnKey':
      return { font: { bold: true }, fill: HEADER_FILL } as CellStyle
    case 'subtotalLabel':
    case 'subtotal':
      return { font: { bold: true } } as CellStyle
    case 'totalLabel':
    case 'total':
      return { font: { bold: true }, fill: HEADER_FILL, border: { top: THIN_BLUE } } as CellStyle
    default:
      return undefined
  }
}

/** Cells the new block would cover that hold data outside the pivot's previous block. */
export function pivotOverwriteConflicts(sheet: SheetData, model: PivotTableModel, output: PivotOutput) {
  const previous = model.extent
  const conflicts: string[] = []
  for (let row = 0; row < output.rows; row += 1) {
    for (let col = 0; col < output.columns; col += 1) {
      const target = { row: model.anchor.row + row, col: model.anchor.col + col }
      if (previous && row < previous.rows && col < previous.cols) continue
      const cell = sheet.cells[address(target.row, target.col)]
      if (cell && ((cell.value !== undefined && cell.value !== null && cell.value !== '') || cell.formula)) conflicts.push(address(target.row, target.col))
      if (conflicts.length >= 5) return conflicts
    }
  }
  return conflicts
}

/** Clear the previous block and write the new one, remembering its extent. */
export function writePivotOutput(sheet: SheetData, model: PivotTableModel, output: PivotOutput) {
  const previous = model.extent
  if (previous) {
    for (let row = 0; row < previous.rows; row += 1) {
      for (let col = 0; col < previous.cols; col += 1) delete sheet.cells[address(model.anchor.row + row, model.anchor.col + col)]
    }
  }
  output.cells.forEach((cells, row) => cells.forEach((cell, col) => {
    const style = styleFor(cell.kind)
    if ((cell.value === null || cell.value === undefined) && !style) return
    const data: CellData = {}
    if (cell.value !== null && cell.value !== undefined) data.value = cell.value
    if (cell.numFmt) data.numFmt = cell.numFmt
    if (style) data.style = style
    sheet.cells[address(model.anchor.row + row, model.anchor.col + col)] = data
  }))
  model.extent = { rows: output.rows, cols: output.columns }
  sheet.rowCount = Math.max(sheet.rowCount, model.anchor.row + output.rows + 1)
  sheet.colCount = Math.max(sheet.colCount, model.anchor.col + output.columns + 1)
}

/** Distinct keys of a source field, for the filter checklist (label, id, count). */
export function pivotFieldKeys(source: PivotSource, field: string, dateGroup?: PivotDateGroup) {
  const index = source.headers.findIndex((header) => header.toLocaleLowerCase() === field.toLocaleLowerCase())
  if (index < 0) return []
  const keys = new Map<string, { key: Key; count: number }>()
  for (const record of source.records) {
    const key = keyOf(record[index], dateGroup)
    const entry = keys.get(key.id)
    if (entry) entry.count += 1
    else keys.set(key.id, { key, count: 1 })
  }
  return [...keys.values()].sort((a, b) => compareKeys(a.key, b.key)).map(({ key, count }) => ({ id: key.id, label: key.label, count }))
}

/** Whether a field holds dates (so date grouping is offered). */
export function pivotFieldIsDate(source: PivotSource, field: string) {
  const index = source.headers.findIndex((header) => header.toLocaleLowerCase() === field.toLocaleLowerCase())
  if (index < 0) return false
  let dates = 0
  let others = 0
  for (const record of source.records.slice(0, 200)) {
    const cell = record[index]
    if (cell?.value === null || cell?.value === undefined || cell.value === '') continue
    if (cell.isDate && typeof cell.value === 'number') dates += 1
    else others += 1
  }
  return dates > 0 && dates >= others
}

export function nextPivotName(sheets: SheetData[]) {
  const taken = new Set(sheets.flatMap((sheet) => (sheet.pivots || []).map((pivot) => pivot.name.toLocaleLowerCase())))
  for (let index = 1; ; index += 1) if (!taken.has(`pivottable${index}`)) return `PivotTable${index}`
}
