/*
 * Paste Special (Excel's dialog / Google Sheets' submenu) as a pure function.
 *
 * `applyPasteSpecial` takes the copied matrix (internal clipboard or parsed HTML/text), the
 * destination selection and an accessor for existing destination cells, and returns the
 * cell changes plus column widths, merges and validation rectangles for the caller to apply.
 * It never mutates its inputs.
 */
import type { CellData, CellScalar } from '../spreadsheet-types'
import {
  CLIPBOARD_MAX_CELLS,
  boundsToRange,
  cellAddress,
  rangeToBounds,
  shiftFormulaA1,
  transposeFormula,
  type ClipboardBounds,
  type ClipboardCoord,
} from './clipboard-html'

export type PasteSpecialPasteType =
  | 'all'
  | 'formulas'
  | 'values'
  | 'formats'
  | 'comments'
  | 'validation'
  | 'allExceptBorders'
  | 'columnWidths'
  | 'formulasAndNumberFormats'
  | 'valuesAndNumberFormats'

export type PasteSpecialOperation = 'none' | 'add' | 'subtract' | 'multiply' | 'divide'

export interface PasteSpecialOptions {
  paste: PasteSpecialPasteType
  operation: PasteSpecialOperation
  skipBlanks: boolean
  transpose: boolean
  /** Write `=Sheet!A1` links to the source cells instead of their contents. */
  pasteLink: boolean
}

export const DEFAULT_PASTE_SPECIAL_OPTIONS: Readonly<PasteSpecialOptions> = Object.freeze({
  paste: 'all',
  operation: 'none',
  skipBlanks: false,
  transpose: false,
  pasteLink: false,
})

export function resolvePasteSpecialOptions(options: Partial<PasteSpecialOptions> = {}): PasteSpecialOptions {
  return { ...DEFAULT_PASTE_SPECIAL_OPTIONS, ...options }
}

/** Paste types an arithmetic operation can combine with (Excel greys the others out). */
export function pasteTypeSupportsOperation(paste: PasteSpecialPasteType): boolean {
  return paste === 'all' || paste === 'allExceptBorders' || paste === 'formulas' || paste === 'values'
    || paste === 'formulasAndNumberFormats' || paste === 'valuesAndNumberFormats'
}

/** Excel enables Paste Link only for "All" / "All except borders" with no operation. */
export function pasteTypeSupportsLink(options: Pick<PasteSpecialOptions, 'paste' | 'operation'>): boolean {
  return (options.paste === 'all' || options.paste === 'allExceptBorders') && options.operation === 'none'
}

export const PASTE_SPECIAL_PRESETS = {
  all: { paste: 'all' },
  values: { paste: 'values' },
  formulas: { paste: 'formulas' },
  formats: { paste: 'formats' },
  comments: { paste: 'comments' },
  validation: { paste: 'validation' },
  allExceptBorders: { paste: 'allExceptBorders' },
  columnWidths: { paste: 'columnWidths' },
  formulasAndNumberFormats: { paste: 'formulasAndNumberFormats' },
  valuesAndNumberFormats: { paste: 'valuesAndNumberFormats' },
  transpose: { paste: 'all', transpose: true },
  pasteLink: { paste: 'all', pasteLink: true },
} as const satisfies Record<string, Partial<PasteSpecialOptions>>

export type PasteSpecialPresetId = keyof typeof PASTE_SPECIAL_PRESETS

export interface PasteSpecialMenuItem {
  id: string
  label: string
  shortcut?: string
  /** Preset to run directly; absent for the item that opens the dialog. */
  preset?: PasteSpecialPresetId
  opensDialog?: boolean
  /** Needs the internal clipboard (links / validation cannot come from another app). */
  requiresInternalSource?: boolean
  separatorBefore?: boolean
  description: string
}

/** Context-menu / Edit-menu entries in the order Google Sheets and Excel present them. */
export const PASTE_SPECIAL_MENU: readonly PasteSpecialMenuItem[] = [
  { id: 'paste-values', label: 'Values only', shortcut: 'Ctrl+Shift+V', preset: 'values', description: 'Paste the displayed values without formulas or formatting.' },
  { id: 'paste-formats', label: 'Formatting only', preset: 'formats', description: 'Paste fonts, fills, borders, alignment and number formats.' },
  { id: 'paste-formulas', label: 'Formulas only', preset: 'formulas', description: 'Paste formulas and constants, keeping the destination formatting.' },
  { id: 'paste-values-number-formats', label: 'Values and number formats', preset: 'valuesAndNumberFormats', description: 'Paste values with their number formats only.' },
  { id: 'paste-transposed', label: 'Transposed', preset: 'transpose', description: 'Turn copied rows into columns and columns into rows.' },
  { id: 'paste-column-widths', label: 'Column widths only', preset: 'columnWidths', description: 'Apply the copied column widths.' },
  { id: 'paste-no-borders', label: 'All except borders', preset: 'allExceptBorders', description: 'Paste everything but keep the destination borders.' },
  { id: 'paste-notes', label: 'Notes only', preset: 'comments', description: 'Paste cell notes and comments.' },
  { id: 'paste-validation', label: 'Data validation only', preset: 'validation', requiresInternalSource: true, description: 'Paste data validation rules.' },
  { id: 'paste-link', label: 'Paste link', preset: 'pasteLink', requiresInternalSource: true, description: 'Insert formulas that link to the copied cells.' },
  { id: 'paste-special-dialog', label: 'Paste special…', shortcut: 'Ctrl+Alt+V', opensDialog: true, separatorBefore: true, description: 'Choose exactly what to paste.' },
]

export type PasteSpecialValidation = Record<string, unknown>

export interface PasteSpecialSource {
  cells: ReadonlyArray<ReadonlyArray<CellData | undefined> | undefined>
  /** Absolute sheet coordinate of cells[0][0]; formulas are shifted from here. */
  origin: ClipboardCoord
  /** Source worksheet name, for Paste Link. */
  sheetName?: string
  /** Merged areas relative to cells[0][0] ("A1:B2" = first two rows/columns). */
  merges?: readonly string[]
  /** Per source column, in whatever unit the caller uses for widths. */
  columnWidths?: ReadonlyArray<number | undefined>
  /** Per source cell validation rule (undefined = none). Omit when unknown (external sources). */
  validations?: ReadonlyArray<ReadonlyArray<PasteSpecialValidation | undefined> | undefined>
}

export interface PasteSpecialInput {
  source: PasteSpecialSource
  /** Destination selection. A single cell pastes once; exact multiples of the copied size tile. */
  destination: { top: number; left: number; bottom?: number; right?: number }
  getDestinationCell: (row: number, col: number) => CellData | undefined
  options?: Partial<PasteSpecialOptions>
  /** Formula shifter, e.g. `shiftFormulaReferences` from lib/formulas. Defaults to a built-in A1 shifter. */
  shiftFormula?: (formula: string, rowDelta: number, colDelta: number) => string
  /** Destination worksheet name; Paste Link omits the sheet prefix when it matches the source. */
  destinationSheetName?: string
  /** Live value of a source formula cell (defaults to its cached `result`). */
  resolveValue?: (cell: CellData, sourceRow: number, sourceCol: number) => CellScalar | undefined
  maxCells?: number
  sheetLimits?: { rows: number; cols: number }
}

export interface PasteSpecialValidationChange {
  range: string
  /** null clears validation from the range. */
  validation: PasteSpecialValidation | null
}

export interface PasteSpecialResult {
  /** Cell writes keyed by A1 address; null deletes the cell. */
  changes: Record<string, CellData | null>
  /** Absolute column index → width (source units). */
  columnWidths?: Record<number, number>
  /** Absolute A1 merge ranges to add. */
  merges?: string[]
  /** Remove existing merges intersecting this area before adding `merges`. */
  clearMergesIn?: ClipboardBounds
  validations?: PasteSpecialValidationChange[]
  /** Area the paste covers; select it afterwards. */
  selection: ClipboardBounds
  cellCount: number
  error?: 'empty' | 'too-large'
}

const CONTENT_KEYS = ['value', 'formula', 'formulaType', 'formulaRange', 'dynamicFormula', 'result', 'resultType', 'display', 'richText', 'arrayMember'] as const
const FORMAT_CELL_TYPES = new Set(['checkbox', 'dropdown'])
const OPERATORS: Record<Exclude<PasteSpecialOperation, 'none'>, string> = { add: '+', subtract: '-', multiply: '*', divide: '/' }

interface Content {
  value?: CellScalar
  formula?: string
  type?: string
  richText?: CellData['richText']
}

function clone<T>(value: T): T {
  return value === undefined ? value : structuredClone(value)
}

function withoutContent(cell: CellData | undefined): CellData {
  const next = cell ? clone(cell) : {}
  for (const key of CONTENT_KEYS) delete next[key]
  if (next.type && !FORMAT_CELL_TYPES.has(next.type)) delete next.type
  return next
}

function withoutFormats(cell: CellData | undefined): CellData {
  const next = cell ? clone(cell) : {}
  delete next.style
  delete next.numFmt
  if (next.type && FORMAT_CELL_TYPES.has(next.type)) delete next.type
  return next
}

function applyFormats(target: CellData, source: CellData | undefined) {
  if (source?.style) target.style = clone(source.style)
  if (source?.numFmt) target.numFmt = source.numFmt
  if (source?.type && FORMAT_CELL_TYPES.has(source.type)) target.type = source.type
}

function applyNumberFormat(target: CellData, source: CellData | undefined) {
  const format = source?.numFmt ?? source?.style?.numFmt
  delete target.numFmt
  if (target.style) {
    delete target.style.numFmt
    if (!Object.keys(target.style).length) delete target.style
  }
  if (format) target.numFmt = format
}

function assignContent(target: CellData, content: Content) {
  if (content.formula !== undefined) target.formula = content.formula
  else if (content.value !== undefined) target.value = content.value
  if (content.type) target.type = content.type
  if (content.richText?.length && content.formula === undefined) target.richText = clone(content.richText)
}

function isBlank(cell: CellData | undefined) {
  return !cell || (!cell.formula && (cell.value === undefined || cell.value === null || cell.value === ''))
}

function hasContent(cell: CellData) {
  if (cell.style && !Object.keys(cell.style).length) delete cell.style
  return cell.value !== undefined || Boolean(cell.formula) || Boolean(cell.style) || Boolean(cell.numFmt)
    || Boolean(cell.note) || Boolean(cell.hyperlink) || Boolean(cell.type && FORMAT_CELL_TYPES.has(cell.type))
}

function roundResult(value: number) {
  return Number.isFinite(value) ? Number.parseFloat(value.toPrecision(15)) : value
}

function numberLiteral(value: number) {
  return String(value)
}

export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name)) return name
  return `'${name.replace(/'/g, "''")}'`
}

type OperationOutcome = Content | 'unchanged'

/** Excel's Operation semantics: numbers combine, formulas are wrapped, text is left alone. */
function combine(content: Content, destination: CellData | undefined, operation: PasteSpecialOperation): OperationOutcome {
  if (operation === 'none') return content
  const operator = OPERATORS[operation]
  let sourceExpression: string | null = null
  let sourceNumber = 0
  if (content.formula !== undefined) sourceExpression = `(${content.formula})`
  else if (typeof content.value === 'number') sourceNumber = content.value
  else if (content.value === undefined || content.value === null || content.value === '') sourceNumber = 0
  else return content // text / booleans are pasted as-is

  if (destination?.formula) {
    return { formula: `(${destination.formula.replace(/^=/, '')})${operator}${sourceExpression ?? numberLiteral(sourceNumber)}` }
  }
  const current = destination?.value
  let destinationNumber: number
  if (current === undefined || current === null || current === '') destinationNumber = 0
  else if (typeof current === 'number') destinationNumber = current
  else return 'unchanged' // text / boolean destinations keep their value

  if (sourceExpression) {
    if (destinationNumber === 0 && operation === 'add') return { formula: content.formula }
    return { formula: `${numberLiteral(destinationNumber)}${operator}${sourceExpression}` }
  }
  switch (operation) {
    case 'add': return { value: roundResult(destinationNumber + sourceNumber) }
    case 'subtract': return { value: roundResult(destinationNumber - sourceNumber) }
    case 'multiply': return { value: roundResult(destinationNumber * sourceNumber) }
    default: return sourceNumber === 0 ? { value: '#DIV/0!' } : { value: roundResult(destinationNumber / sourceNumber) }
  }
}

function rectanglesFromCells(cells: Array<[number, number]>): ClipboardBounds[] {
  const byRow = new Map<number, number[]>()
  for (const [row, col] of cells) {
    const list = byRow.get(row) ?? []
    list.push(col)
    byRow.set(row, list)
  }
  const rows = [...byRow.keys()].sort((a, b) => a - b)
  const done: ClipboardBounds[] = []
  let open = new Map<string, ClipboardBounds>()
  let previousRow = Number.NaN
  for (const row of rows) {
    const cols = (byRow.get(row) as number[]).sort((a, b) => a - b)
    const runs: Array<[number, number]> = []
    for (const col of cols) {
      const last = runs[runs.length - 1]
      if (last && col === last[1] + 1) last[1] = col
      else if (!last || col > last[1]) runs.push([col, col])
    }
    const next = new Map<string, ClipboardBounds>()
    for (const [left, right] of runs) {
      const key = `${left}:${right}`
      const extending = row === previousRow + 1 ? open.get(key) : undefined
      if (extending) {
        extending.bottom = row
        next.set(key, extending)
        open.delete(key)
      } else next.set(key, { top: row, bottom: row, left, right })
    }
    open.forEach((rect) => done.push(rect))
    open = next
    previousRow = row
  }
  open.forEach((rect) => done.push(rect))
  return done
}

/**
 * Computes a Paste Special operation. Tiling follows Excel: when the destination selection is
 * an exact multiple of the (possibly transposed) copied block, the block repeats to fill it;
 * otherwise it pastes once at the top-left cell.
 */
export function applyPasteSpecial(input: PasteSpecialInput): PasteSpecialResult {
  const options = resolvePasteSpecialOptions(input.options)
  const { source } = input
  const limits = input.sheetLimits ?? { rows: 1_048_576, cols: 16_384 }
  const maxCells = input.maxCells ?? CLIPBOARD_MAX_CELLS
  const shift = input.shiftFormula ?? shiftFormulaA1
  const top = Math.max(0, input.destination.top)
  const left = Math.max(0, input.destination.left)
  const sourceRows = source.cells.length
  let sourceCols = 0
  for (const row of source.cells) if (row && row.length > sourceCols) sourceCols = row.length
  const emptySelection = { top, left, bottom: top, right: left }
  if (!sourceRows || !sourceCols) return { changes: {}, selection: emptySelection, cellCount: 0, error: 'empty' }

  const transpose = options.transpose
  const blockRows = transpose ? sourceCols : sourceRows
  const blockCols = transpose ? sourceRows : sourceCols
  const selectedRows = Math.max(1, (input.destination.bottom ?? top) - top + 1)
  const selectedCols = Math.max(1, (input.destination.right ?? left) - left + 1)
  const tiles = selectedRows % blockRows === 0 && selectedCols % blockCols === 0
    ? { rows: selectedRows / blockRows, cols: selectedCols / blockCols }
    : { rows: 1, cols: 1 }
  const bottom = Math.min(limits.rows - 1, top + blockRows * tiles.rows - 1)
  const right = Math.min(limits.cols - 1, left + blockCols * tiles.cols - 1)
  const selection: ClipboardBounds = { top, left, bottom, right }
  const cellCount = (bottom - top + 1) * (right - left + 1)
  if (cellCount > maxCells) return { changes: {}, selection, cellCount, error: 'too-large' }

  const paste = options.paste
  const operation = pasteTypeSupportsOperation(paste) && !options.pasteLink ? options.operation : 'none'
  const skipBlanks = options.skipBlanks && !options.pasteLink
  const changes: Record<string, CellData | null> = {}
  const linkPrefix = source.sheetName && source.sheetName !== input.destinationSheetName ? `${quoteSheetName(source.sheetName)}!` : ''
  const writesCells = paste !== 'validation' && paste !== 'columnWidths'
  const writesValidation = !options.pasteLink && (paste === 'all' || paste === 'allExceptBorders' || paste === 'validation') && Boolean(source.validations)
  const validationCells = new Map<string, { validation: PasteSpecialValidation | null; cells: Array<[number, number]> }>()

  const moveFormula = (formula: string, sourceRow: number, sourceCol: number, destRow: number, destCol: number) => {
    const fromRow = source.origin.row + sourceRow
    const fromCol = source.origin.col + sourceCol
    return transpose
      ? transposeFormula(formula, fromRow, fromCol, destRow, destCol)
      : shift(formula, destRow - fromRow, destCol - fromCol)
  }

  const sourceContent = (cell: CellData | undefined, sourceRow: number, sourceCol: number, destRow: number, destCol: number, values: boolean, rich: boolean): Content => {
    if (!cell) return {}
    const type = cell.type && !FORMAT_CELL_TYPES.has(cell.type) ? cell.type : undefined
    if (cell.formula) {
      if (!values) return { formula: moveFormula(cell.formula.replace(/^=/, ''), sourceRow, sourceCol, destRow, destCol) }
      const value = input.resolveValue?.(cell, sourceRow, sourceCol) ?? cell.result
      return { value: value === null ? undefined : value }
    }
    const content: Content = { value: cell.value === null ? undefined : cell.value, type }
    if (rich && cell.richText?.length) content.richText = cell.richText
    return content
  }

  for (let tileRow = 0; tileRow < tiles.rows; tileRow += 1) {
    for (let tileCol = 0; tileCol < tiles.cols; tileCol += 1) {
      for (let blockRow = 0; blockRow < blockRows; blockRow += 1) {
        const destRow = top + tileRow * blockRows + blockRow
        if (destRow >= limits.rows) break
        for (let blockCol = 0; blockCol < blockCols; blockCol += 1) {
          const destCol = left + tileCol * blockCols + blockCol
          if (destCol >= limits.cols) break
          const sourceRow = transpose ? blockCol : blockRow
          const sourceCol = transpose ? blockRow : blockCol
          const sourceCell = source.cells[sourceRow]?.[sourceCol]
          if (skipBlanks && isBlank(sourceCell)) continue

          if (writesValidation) {
            const validation = source.validations?.[sourceRow]?.[sourceCol] ?? null
            const key = validation ? JSON.stringify(validation) : 'null'
            const group = validationCells.get(key) ?? { validation, cells: [] }
            group.cells.push([destRow, destCol])
            validationCells.set(key, group)
          }
          if (!writesCells) continue

          const address = cellAddress(destRow, destCol)
          const destination = input.getDestinationCell(destRow, destCol)
          let next: CellData | undefined
          if (options.pasteLink) {
            next = withoutContent(destination)
            next.formula = `${linkPrefix}${cellAddress(source.origin.row + sourceRow, source.origin.col + sourceCol)}`
          } else if (paste === 'comments') {
            if (sourceCell?.note === undefined) continue
            next = destination ? clone(destination) : {}
            next.note = clone(sourceCell.note)
          } else if (paste === 'formats') {
            next = withoutFormats(destination)
            applyFormats(next, sourceCell)
          } else {
            const values = paste === 'values' || paste === 'valuesAndNumberFormats'
            const everything = paste === 'all' || paste === 'allExceptBorders'
            const content = sourceContent(sourceCell, sourceRow, sourceCol, destRow, destCol, values, everything && operation === 'none')
            const outcome = combine(content, destination, operation)
            if (everything) {
              next = withoutContent(sourceCell)
              // A text destination keeps its value under an operation; the formats still paste.
              assignContent(next, outcome === 'unchanged' ? sourceContent(destination, 0, 0, destRow, destCol, true, true) : outcome)
              if (paste === 'allExceptBorders') {
                if (next.style) delete next.style.border
                const border = destination?.style?.border
                if (border) next.style = { ...(next.style ?? {}), border: clone(border) }
              }
            } else {
              if (outcome === 'unchanged') continue
              next = withoutContent(destination)
              assignContent(next, outcome)
              if (paste === 'formulasAndNumberFormats' || paste === 'valuesAndNumberFormats') applyNumberFormat(next, sourceCell)
            }
          }
          changes[address] = hasContent(next) ? next : null
        }
      }
    }
  }

  const result: PasteSpecialResult = { changes, selection, cellCount }

  if (paste === 'columnWidths' && source.columnWidths && !transpose) {
    const widths: Record<number, number> = {}
    // Only the column count matters for widths: C1:F1 takes a two-column copy twice.
    const widthTiles = selectedCols % blockCols === 0 ? selectedCols / blockCols : 1
    for (let tileCol = 0; tileCol < widthTiles; tileCol += 1) {
      for (let col = 0; col < blockCols; col += 1) {
        const width = source.columnWidths[col]
        const destCol = left + tileCol * blockCols + col
        if (width !== undefined && Number.isFinite(width) && width > 0 && destCol < limits.cols) widths[destCol] = width
      }
    }
    if (Object.keys(widths).length) result.columnWidths = widths
  }

  if (!options.pasteLink && (paste === 'all' || paste === 'allExceptBorders' || paste === 'formats')) {
    const merges: string[] = []
    for (const range of source.merges ?? []) {
      const bounds = rangeToBounds(range)
      if (!bounds || bounds.bottom >= sourceRows || bounds.right >= sourceCols) continue
      const relative = transpose
        ? { top: bounds.left, left: bounds.top, bottom: bounds.right, right: bounds.bottom }
        : bounds
      for (let tileRow = 0; tileRow < tiles.rows; tileRow += 1) {
        for (let tileCol = 0; tileCol < tiles.cols; tileCol += 1) {
          const rowOffset = top + tileRow * blockRows
          const colOffset = left + tileCol * blockCols
          const merged = { top: relative.top + rowOffset, bottom: relative.bottom + rowOffset, left: relative.left + colOffset, right: relative.right + colOffset }
          if (merged.bottom < limits.rows && merged.right < limits.cols) merges.push(boundsToRange(merged))
        }
      }
    }
    result.clearMergesIn = selection
    if (merges.length) result.merges = merges
  }

  if (writesValidation && validationCells.size) {
    const validations: PasteSpecialValidationChange[] = []
    validationCells.forEach(({ validation, cells }) => {
      for (const rect of rectanglesFromCells(cells)) validations.push({ range: boundsToRange(rect), validation: validation ? clone(validation) : null })
    })
    validations.sort((a, b) => Number(a.validation !== null) - Number(b.validation !== null))
    result.validations = validations
  }
  return result
}
