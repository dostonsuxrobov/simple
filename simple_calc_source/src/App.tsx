import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent, CSSProperties, DragEvent, KeyboardEvent, MouseEvent, MutableRefObject, PointerEvent as ReactPointerEvent, RefObject } from 'react'
import { applyPatches, enablePatches, produceWithPatches, setAutoFreeze } from 'immer'
import type { Patch } from 'immer'
import type { ChartAnchor, PivotTableModel, SheetFilterState, SheetImage, SheetTable, SparklineGroup } from './spreadsheet-types'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDownAZ,
  ArrowUpZA,
  Bold,
  Check,
  CheckSquare,
  ChevronDown,
  CircleAlert,
  ClipboardPaste,
  Columns3,
  Copy,
  DollarSign,
  Download,
  Eraser,
  EyeOff,
  FilePlus2,
  FileSpreadsheet,
  FolderOpen,
  FunctionSquare,
  Grid2X2,
  Info,
  Italic,
  Link,
  ListPlus,
  Maximize2,
  Menu,
  MessageSquarePlus,
  Minus,
  PaintBucket,
  Percent,
  RotateCw,
  Rows3,
  Pencil,
  Plus,
  Printer,
  Redo2,
  Replace,
  Save,
  Scissors,
  Search,
  BarChart3,
  Filter,
  FilterX,
  Highlighter,
  Paintbrush,
  Palette,
  Sigma,
  SlidersHorizontal,
  Tag,
  Square,
  StickyNote,
  Strikethrough,
  ChartSpline,
  ImagePlus,
  Lock,
  LockOpen,
  TableProperties,
  Table2,
  TableCellsMerge,
  TableCellsSplit,
  Trash2,
  Type,
  Underline,
  Undo2,
  WrapText,
  X,
} from 'lucide-react'
import appIcon from '../icon.png'
import { SpreadsheetContextMenu, SpreadsheetMenus } from './components/SpreadsheetMenus'
import type { SpreadsheetMenuDefinition, SpreadsheetMenuItem } from './components/SpreadsheetMenus'
import { SpreadsheetExportDialog } from './components/SpreadsheetExportDialog'
import { FormulaAssist, FormulaHighlight } from './components/FormulaAssist'
import { NameManagerDialog, validateDefinedName } from './components/NameManagerDialog'
import { FormatCellsDialog } from './components/FormatCellsDialog'
import type { FormatCellsTab } from './components/FormatCellsDialog'
import { ColorDropdownButton } from './components/ColorPicker'
import { BorderPicker, BorderPresetIcon, BORDER_PRESET_ITEMS } from './components/BorderPicker'
import { CellStylesGallery } from './components/CellStylesGallery'
import { CommandBar } from './components/CommandBar'
import { ToolbarMenuButton } from './components/ToolbarMenuButton'
import type { ToolbarMenuItem } from './components/ToolbarMenuButton'
import { applyCellStylePresetToCell, applyFormatChangeToCell, applyNeighborBorderChange, applyTint, borderPresetChange, cellPosition } from './lib/cell-styles'
import type { BorderPreset, CellStylePreset, FormatBorderChange, FormatCellsChange } from './lib/cell-styles'
import { adjustDecimals as adjustFormatDecimals } from './lib/format-codes'
import { FONT_FAMILIES, FONT_SIZES } from './lib/fonts'
import type { NameDraft } from './components/NameManagerDialog'
import type { AssistItem, FormulaAssistState } from './components/FormulaAssist'
import { SpreadsheetPrintDialog } from './components/SpreadsheetPrintDialog'
import { createAutofillPatch } from './lib/autofill'
import { readRichClipboard, writeRichClipboard } from './lib/clipboard'
import { isClipboardTooLarge, parseClipboardPayload, serializeSelectionToClipboard } from './lib/clipboard-html'
import { applyPasteSpecial, PASTE_SPECIAL_MENU, PASTE_SPECIAL_PRESETS } from './lib/paste-special'
import type { PasteSpecialOptions } from './lib/paste-special'
import { PasteSpecialDialog } from './components/PasteSpecialDialog'
import { ChartLayer } from './components/ChartLayer'
import { ImageLayer } from './components/ImageLayer'
import { TraceArrowLayer } from './components/TraceArrowLayer'
import type { TraceArrow } from './components/TraceArrowLayer'
import { ChartEditorPanel } from './components/ChartEditorPanel'
import { createChartFromRange, renameWorkbookChartReferences, resolveChartData, transformWorkbookChartsForStructure, workbookChartAccessor, workbookChartPalette } from './lib/charts'
import type { ChartGeometry, ChartWorkbookAccessor } from './lib/charts'
import { buildPrintChartPayload } from './lib/chart-render'
import { isFormulaError, shiftFormulaReferences } from './lib/formulas'
import { CalculationEngine, markArrayMembers } from './lib/calc-engine'
import { fromFileWorkbook, toFileWorkbook } from './lib/formula-file-format'
import { callContextAt, canInsertReferenceAt, completionContextAt, formulaReferences, insertReference, moveReferencesInFormula, removeSheetFromFormula, renameSheetInFormula, tokenColor, tokenizeFormulaText, toggleAbsoluteReference } from './lib/formula-editing'
import { functionSignature, getFunctionInfo, searchFunctions } from './lib/function-catalog'
import './lib/formula-library'
import { inferFormulaNumberFormat, parseCellInput } from './lib/input-parsing'
import { fitNumericText, measureTextWidth } from './lib/cell-text-fit'
import { tableAt, tableCellPaint } from './lib/table-styles'
import { computeConditionalFormats, dataBarBackground } from './lib/conditional-format'
import type { ConditionalCellFormat } from './lib/conditional-format'
import { ConditionalFormatPanel } from './components/ConditionalFormatPanel'
import { ConditionalIcon } from './components/ConditionalIcon'
import { AlertDialog } from './components/AlertDialog'
import type { AlertRequest } from './components/AlertDialog'
import { FilterMenu } from './components/FilterMenu'
import { SortDialog } from './components/SortDialog'
import type { SortDialogResult } from './components/SortDialog'
import { DataValidationDialog } from './components/DataValidationDialog'
import { RemoveDuplicatesDialog, TextToColumnsDialog } from './components/DataCleanupDialogs'
import { applyCellChanges, createSheetHost } from './lib/data-tools-core'
import type { DataHost } from './lib/data-tools-core'
import { applyFilterResult, clearFilterCriteria, computeFilteredRows, createFilterState, criteriaIsActive, distinctColumnValues, extendFilterRange, filterIsActive, filterRangeForSelection, hiddenRowsWithoutFilter, setColumnCriteria } from './lib/filter'
import { convertTableToRange, createTable, expandTableForEntry, fillCalculatedColumn, formatTableRef, parseTableRef, renameTable, resizeTable, setTotalsFunction, setTotalsRow, shiftCellsDown, syncTableHeaders, tableContaining, tableRegions, writeTotalsCells } from './lib/tables'
import type { TotalFunctionId } from './lib/tables'
import { CreateTableDialog, TableDesignPanel } from './components/TableTools'
import { ShiftCellsDialog } from './components/ShiftCellsDialog'
import { GoalSeekDialog } from './components/GoalSeekDialog'
import { goalSeek } from './lib/goal-seek'
import { PROTECTED_MESSAGE, cellLocked, protectionAllows, protectionViolation, sheetProtection } from './lib/protection'
import { ProtectSheetDialog } from './components/ProtectSheetDialog'
import type { ProtectSheetResult } from './components/ProtectSheetDialog'
import { isSparklineValue, parseSparkline } from './lib/formula-lib-sparkline'
import { groupSparklineSpec, renderSparklineSvg } from './lib/sparkline-render'
import type { SparklineSpec } from './lib/formula-lib-sparkline'
import { CreateSparklinesDialog } from './components/CreateSparklinesDialog'
import type { SparklineKind } from './components/CreateSparklinesDialog'
import { computePivot, nextPivotName, pivotOverwriteConflicts, writePivotOutput } from './lib/pivot'
import type { PivotOutput, PivotSource } from './lib/pivot'
import { CreatePivotDialog, PivotEditorPanel } from './components/PivotEditorPanel'
import { isDateNumberFormat } from './lib/data-tools-core'
import { uniqueColumnNames } from './lib/tables'
import { changeOutline, clearOutline, maxOutlineLevel, outlineGroups, showOutlineLevel, toggleGroup } from './lib/outline'
import type { OutlineAxis, OutlineGroup } from './lib/outline'
import type { ShiftCellsChoice } from './components/ShiftCellsDialog'
import type { TableOption } from './components/TableTools'
import { detectHeaderRow, sortKeyColors, sortKeyLabels, sortRange, sortSpecForFilter } from './lib/sort'
import { findInvalidCells, findValidation, listOptionsForValidation, rangesWithSameValidation, validateValue, withValidation } from './lib/validation'
import { changeCase, fillBlanksFromAbove, removeDuplicates, splitSourceLines, splitTextToColumns, trimWhitespace } from './lib/data-cleanup'
import { isDateTimeFormat } from './lib/format-codes'
import type { ChangeHint } from './lib/calc-engine'
import { accountingDisplayParts, formatColor, formatScalar, isAccountingNumberFormat } from './lib/number-format'
import { registerRecoverySave } from './lib/recovery'
import { applySelectionStructureCommand, shiftCells } from './lib/sheet-operations'
import type { SelectionStructureCommand } from './lib/sheet-operations'
import type {
  CellBorder,
  CellBorderSide,
  CellData,
  CellFill,
  CellFont,
  CellScalar,
  CellStyle,
  RecentWorkbook,
  SheetChart,
  SheetData,
  SpreadsheetExportFormat,
  SpreadsheetPrintOptions,
  SpreadsheetColor,
  SpreadsheetDisplayParts,
  WorkbookModel,
  WorkbookPayload,
} from './spreadsheet-types'

enablePatches()
setAutoFreeze(false)

const RECENT_KEY = 'simple-calc:recent:v1'
const DEFAULT_ROWS = 200
const DEFAULT_COLS = 40
const DEFAULT_ROW_HEIGHT = 26
const DEFAULT_COL_WIDTH = 108
const IMPORTED_ROW_HEIGHT = 20
const IMPORTED_COL_WIDTH = 64
const HEADER_HEIGHT = 27
const HEADER_WIDTH = 46
const MAX_HISTORY = 200
const MIN_COLUMN_PIXELS = 20
const MAX_COLUMN_PIXELS = 1_000
const MIN_ROW_PIXELS = 12
const MAX_ROW_POINTS = 409

interface Coord {
  row: number
  col: number
}

interface Selection {
  anchor: Coord
  focus: Coord
}

interface GridContextTarget {
  kind: 'cell' | 'row-header' | 'column-header' | 'corner'
  coord: Coord
}

const NON_TEXT_INPUT_TYPES = new Set(['button', 'checkbox', 'color', 'file', 'hidden', 'image', 'radio', 'range', 'reset', 'submit'])

function isNativeTextEditingTarget(target: EventTarget | null) {
  if (!(target instanceof Element)) return false
  const editable = target.closest<HTMLElement>('[contenteditable]')
  if (editable?.isContentEditable) return true
  if (target instanceof HTMLTextAreaElement) return true
  return target instanceof HTMLInputElement && !NON_TEXT_INPUT_TYPES.has(target.type.toLocaleLowerCase())
}

type AggregateMode = 'sum' | 'average' | 'minimum' | 'maximum' | 'count' | 'numeric'
const STATUS_STATS: Array<{ key: AggregateMode; label: string }> = [
  { key: 'average', label: 'Average' },
  { key: 'count', label: 'Count' },
  { key: 'numeric', label: 'Numerical count' },
  { key: 'minimum', label: 'Min' },
  { key: 'maximum', label: 'Max' },
  { key: 'sum', label: 'Sum' },
]
const STATUS_STATS_KEY = 'simple-calc:status-stats:v1'

function readStatusStats(): Set<AggregateMode> {
  try {
    const stored = JSON.parse(localStorage.getItem(STATUS_STATS_KEY) || 'null')
    if (Array.isArray(stored)) return new Set(stored.filter((item): item is AggregateMode => STATUS_STATS.some((stat) => stat.key === item)))
  } catch {
    // Fall back to Excel's defaults.
  }
  return new Set<AggregateMode>(['average', 'count', 'sum'])
}

interface OpenWorkbook {
  documentId: string
  path: string | null
  name: string
  sourceFormat: string
  requiresSaveAs: boolean
  warnings: string[]
  backupPath?: string
  stats?: WorkbookPayload['stats']
}

interface EditingState {
  address: string
  draft: string
  /** Excel's Enter mode (started by typing: arrows commit or point) vs Edit mode (F2: arrows move the caret). */
  mode?: 'enter' | 'edit'
}

interface OverlayBounds {
  top: number
  left: number
  bottom: number
  right: number
}

interface ReferenceHighlight extends OverlayBounds {
  color: string
}

interface HighlightPart {
  text: string
  color?: string
}

interface InternalClipboard {
  text: string
  origin: Coord
  cells: CellData[][]
  sheetId: string
  sheetName: string
  /** Merges relative to the copied block ("A1" = first copied cell). */
  merges: string[]
  columnWidths: Array<number | undefined>
  validations?: Array<Array<Record<string, unknown> | undefined>>
  /** Ctrl+X: the next paste moves the block instead of copying it. */
  cut: boolean
}

interface HistoryEntry {
  patches: Patch[]
  inversePatches: Patch[]
}

interface TextPromptRequest {
  title: string
  label: string
  initialValue?: string
  multiline?: boolean
}

interface TextPromptState extends TextPromptRequest {
  id: number
  resolve: (value: string | null) => void
}

function makeId(prefix = 'sheet') {
  return `${prefix}-${globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2)}`}`
}

function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}

function columnName(index: number) {
  let value = index + 1
  let result = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    value = Math.floor((value - 1) / 26)
  }
  return result
}

function columnIndex(name: string) {
  return name.toUpperCase().split('').reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0) - 1
}

function addressOf(coord: Coord) {
  return `${columnName(coord.col)}${coord.row + 1}`
}

function coordOf(address: string): Coord | null {
  const match = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(address.trim())
  if (!match) return null
  return { col: columnIndex(match[1]), row: Number(match[2]) - 1 }
}

function selectionBounds(selection: Selection) {
  return {
    top: Math.min(selection.anchor.row, selection.focus.row),
    bottom: Math.max(selection.anchor.row, selection.focus.row),
    left: Math.min(selection.anchor.col, selection.focus.col),
    right: Math.max(selection.anchor.col, selection.focus.col),
  }
}

function rangeAddress(bounds: { top: number; bottom: number; left: number; right: number }) {
  const start = addressOf({ row: bounds.top, col: bounds.left })
  const end = addressOf({ row: bounds.bottom, col: bounds.right })
  return start === end ? start : `${start}:${end}`
}

function adjustedDecimalFormat(format: string | undefined, delta: 1 | -1) {
  const current = format && format !== 'General' ? format : '0'
  if (/[dmyhs]/i.test(current.replace(/\[[^\]]*]/g, ''))) return null
  const match = /\.([0#]+)/.exec(current)
  const decimals = match?.[1].length || 0
  const nextDecimals = clamp(decimals + delta, 0, 12)
  if (nextDecimals === decimals) return current
  if (match) return nextDecimals ? current.replace(/\.([0#]+)/, `.${'0'.repeat(nextDecimals)}`) : current.replace(/\.([0#]+)/, '')
  const percentIndex = current.indexOf('%')
  const suffixIndex = percentIndex >= 0 ? percentIndex : current.length
  return `${current.slice(0, suffixIndex)}.${'0'.repeat(nextDecimals)}${current.slice(suffixIndex)}`
}

function fillSelectionForTarget(selection: Selection, target: Coord): Selection {
  const bounds = selectionBounds(selection)
  const rowDistance = target.row < bounds.top
    ? bounds.top - target.row
    : target.row > bounds.bottom ? target.row - bounds.bottom : 0
  const colDistance = target.col < bounds.left
    ? bounds.left - target.col
    : target.col > bounds.right ? target.col - bounds.right : 0

  if (rowDistance === 0 && colDistance === 0) return selection
  if (rowDistance >= colDistance) {
    return {
      anchor: { row: Math.min(bounds.top, target.row), col: bounds.left },
      focus: { row: Math.max(bounds.bottom, target.row), col: bounds.right },
    }
  }
  return {
    anchor: { row: bounds.top, col: Math.min(bounds.left, target.col) },
    focus: { row: bounds.bottom, col: Math.max(bounds.right, target.col) },
  }
}

function rangeAddresses(selection: Selection, limit = 100_000) {
  const bounds = selectionBounds(selection)
  const total = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
  if (total > limit) return []
  const addresses: string[] = []
  for (let row = bounds.top; row <= bounds.bottom; row += 1) {
    for (let col = bounds.left; col <= bounds.right; col += 1) addresses.push(addressOf({ row, col }))
  }
  return addresses
}

function hiddenIndexSet(values?: number[]) {
  return new Set((values || []).map((value) => Number(value) - 1).filter((value) => Number.isInteger(value) && value >= 0))
}

function stepPastHidden(position: number, delta: number, max: number, hidden: Set<number>) {
  if (!delta) return clamp(position, 0, max)
  const step = delta > 0 ? 1 : -1
  let index = clamp(position, 0, max)
  for (let remaining = Math.abs(delta); remaining > 0; remaining -= 1) {
    let next = index + step
    while (next >= 0 && next <= max && hidden.has(next)) next += step
    if (next < 0 || next > max) break
    index = next
  }
  return index
}

function edgeJumpCoord(sheet: SheetData, from: Coord, rowDelta: number, colDelta: number, hidden: Set<number>): Coord {
  const vertical = rowDelta !== 0
  const step = (vertical ? rowDelta : colDelta) > 0 ? 1 : -1
  const position = vertical ? from.row : from.col
  const filled = new Set<number>()
  let farthest = Math.max(0, (vertical ? sheet.rowCount : sheet.colCount) - 1)
  for (const [address, cell] of Object.entries(sheet.cells)) {
    if (!cell.formula && (cell.value === undefined || cell.value === null || cell.value === '')) continue
    const coord = coordOf(address)
    if (!coord || (vertical ? coord.col !== from.col : coord.row !== from.row)) continue
    const index = vertical ? coord.row : coord.col
    farthest = Math.max(farthest, index)
    if (!hidden.has(index)) filled.add(index)
  }
  const edge = step > 0 ? farthest : 0
  let target = position
  if (filled.has(position) && filled.has(position + step)) {
    while (target !== edge && filled.has(target + step)) target += step
  } else {
    target = step > 0 ? Math.max(edge, position) : Math.min(edge, position)
    for (let index = position + step; step > 0 ? index <= edge : index >= edge; index += step) {
      if (filled.has(index)) { target = index; break }
    }
  }
  while (hidden.has(target) && (step > 0 ? target > position : target < position)) target -= step
  return vertical ? { row: target, col: from.col } : { row: from.row, col: target }
}

function contiguousRegionBounds(sheet: SheetData, origin: Coord) {
  const filled: Coord[] = []
  for (const [address, cell] of Object.entries(sheet.cells)) {
    if (!cell.formula && (cell.value === undefined || cell.value === null || cell.value === '')) continue
    const coord = coordOf(address)
    if (coord) filled.push(coord)
  }
  const bounds = { top: origin.row, bottom: origin.row, left: origin.col, right: origin.col }
  let grew = true
  while (grew) {
    grew = false
    for (const coord of filled) {
      if (coord.row < bounds.top - 1 || coord.row > bounds.bottom + 1 || coord.col < bounds.left - 1 || coord.col > bounds.right + 1) continue
      if (coord.row < bounds.top) { bounds.top = coord.row; grew = true }
      if (coord.row > bounds.bottom) { bounds.bottom = coord.row; grew = true }
      if (coord.col < bounds.left) { bounds.left = coord.col; grew = true }
      if (coord.col > bounds.right) { bounds.right = coord.col; grew = true }
    }
  }
  return bounds
}

function compareCellScalars(a: CellScalar, b: CellScalar) {
  const rank = (value: CellScalar) => typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2
  if (rank(a) !== rank(b)) return rank(a) - rank(b)
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b)
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' })
}

interface SearchOptions {
  matchCase: boolean
  entireCell: boolean
  workbook: boolean
  lookIn: 'values' | 'formulas'
}

function textMatches(text: string, query: string, options: SearchOptions) {
  if (!query) return false
  if (options.matchCase) return options.entireCell ? text === query : text.includes(query)
  const haystack = text.toLocaleLowerCase()
  const needle = query.toLocaleLowerCase()
  return options.entireCell ? haystack === needle : haystack.includes(needle)
}

const sortedAddressCache = new WeakMap<SheetData['cells'], string[]>()
function sortedCellAddresses(cells: SheetData['cells']) {
  let sorted = sortedAddressCache.get(cells)
  if (!sorted) {
    sorted = Object.keys(cells).map((address) => ({ address, coord: coordOf(address) || { row: 0, col: 0 } }))
      .sort((a, b) => a.coord.row - b.coord.row || a.coord.col - b.coord.col)
      .map((item) => item.address)
    sortedAddressCache.set(cells, sorted)
  }
  return sorted
}

function replaceAllOccurrences(source: string, query: string, replacement: string, matchCase = false) {
  if (!query) return source
  if (matchCase) return source.split(query).join(replacement)
  // Compare per candidate position instead of index-mapping across a
  // lowercased copy: toLocaleLowerCase can change string length (e.g. 'İ'),
  // which would scatter haystack indices across the original source.
  const needle = query.toLocaleLowerCase()
  let output = ''
  let position = 0
  let index = 0
  while (index + query.length <= source.length) {
    if (source.slice(index, index + query.length).toLocaleLowerCase() === needle) {
      output += source.slice(position, index) + replacement
      index += query.length
      position = index
    } else {
      index += 1
    }
  }
  return output + source.slice(position)
}

function replacementDraft(cell: CellData, query: string, replacement: string, options?: SearchOptions) {
  const matchCase = Boolean(options?.matchCase)
  if (cell.formula) {
    if (options?.entireCell) return textMatches(`=${cell.formula}`, query, options) ? replacement : null
    const replaced = replaceAllOccurrences(cell.formula, query, replacement, matchCase)
    return replaced === cell.formula ? null : `=${replaced}`
  }
  const source = cell.value === null || cell.value === undefined ? '' : String(cell.value)
  if (options?.entireCell) return textMatches(source, query, options) ? replacement : null
  const replaced = replaceAllOccurrences(source, query, replacement, matchCase)
  return replaced === source ? null : replaced
}

function normalizeFormula(formula: string) {
  return formula.trim().replace(/^=/, '')
}

function dateSerialFromDraft(draft: string) {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(draft)
  const local = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(draft)
  const year = Number(iso?.[1] || local?.[3])
  const month = Number(iso?.[2] || local?.[1])
  const day = Number(iso?.[3] || local?.[2])
  if (!year || !month || !day) return null
  const timestamp = Date.UTC(year, month - 1, day)
  const date = new Date(timestamp)
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null
  return (timestamp - Date.UTC(1899, 11, 30)) / 86_400_000
}

function rawCellValue(cell?: CellData) {
  if (!cell) return ''
  if (cell.formula) return `=${normalizeFormula(cell.formula)}`
  if (cell.value === null || cell.value === undefined) return ''
  // Never stringify a structured value into "[object Object]" — fall back to the text the
  // file rendered.
  if (typeof cell.value === 'object') return cell.display ?? ''
  return String(cell.value)
}

function parseDraft(draft: string, previous: CellData = {}): CellData {
  const next: CellData = { ...previous }
  delete next.display
  delete next.result
  delete next.resultType
  delete next.formula
  delete next.formulaType
  delete next.formulaRange
  delete next.dynamicFormula
  delete next.value
  delete next.type
  delete next.richText
  delete next.arrayMember

  if (draft.startsWith("'") || (previous.numFmt || previous.style?.numFmt) === '@') {
    next.value = draft.startsWith("'") ? draft.slice(1) : draft
    return next
  }
  if (draft.startsWith('=')) {
    const formula = normalizeFormula(draft)
    if (formula) next.formula = formula
    return next
  }
  if (!draft.length) return next
  const parsed = parseCellInput(draft)
  if (!parsed) {
    next.value = draft
    return next
  }
  next.value = parsed.value
  if (parsed.type) next.type = parsed.type
  const existingFormat = previous.numFmt || previous.style?.numFmt
  // An entry only brings its implied format into a General cell (a typed "15%" keeps an
  // existing percent format; a date typed into a currency cell keeps the currency format).
  if (parsed.numFmt && (!existingFormat || existingFormat === 'General')) next.numFmt = parsed.numFmt
  return next
}

const EMU_PER_PIXEL = 9525

/** A cell anchor covering `width` x `height` pixels (100% zoom) from the top-left of (row, col). */
function anchorForSize(sheet: SheetData, row: number, col: number, width: number, height: number): ChartAnchor {
  const columnPixels = (index: number) => (sheet.hiddenCols?.includes(index + 1) ? 0 : sheet.colWidths[String(index + 1)] === undefined ? IMPORTED_COL_WIDTH : columnPixelWidth(sheet.colWidths[String(index + 1)], 1))
  const rowPixels = (index: number) => (sheet.hiddenRows?.includes(index + 1) ? 0 : sheet.rowHeights[String(index + 1)] === undefined ? IMPORTED_ROW_HEIGHT : rowPixelHeight(sheet.rowHeights[String(index + 1)], 1))
  const walk = (start: number, length: number, sizeOf: (index: number) => number) => {
    let index = start
    let remaining = length
    for (let guard = 0; guard < 16_384 && remaining > sizeOf(index); guard += 1) { remaining -= sizeOf(index); index += 1 }
    return { index, offset: Math.round(remaining * EMU_PER_PIXEL) }
  }
  const end = { col: walk(col, width, columnPixels), row: walk(row, height, rowPixels) }
  return { from: { row, col, rowOffsetEmu: 0, colOffsetEmu: 0 }, to: { row: end.row.index, col: end.col.index, rowOffsetEmu: end.row.offset, colOffsetEmu: end.col.offset }, editAs: 'oneCell' }
}

/** Load a picture as a data URL Excel can store (PNG/JPEG/GIF; others are converted to PNG) with its natural size. */
async function preparePicture(source: Blob | string): Promise<{ src: string; width: number; height: number }> {
  const dataUrl = typeof source === 'string' ? source : await new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result))
    reader.onerror = () => reject(reader.error || new Error('The picture could not be read.'))
    reader.readAsDataURL(source)
  })
  const image = await new Promise<HTMLImageElement>((resolve, reject) => {
    const element = new Image()
    element.onload = () => resolve(element)
    element.onerror = () => reject(new Error('That file is not a picture this app can show.'))
    element.src = dataUrl
  })
  const width = image.naturalWidth || 96
  const height = image.naturalHeight || 96
  if (/^data:image\/(png|jpe?g|gif);/i.test(dataUrl)) return { src: dataUrl, width, height }
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  canvas.getContext('2d')?.drawImage(image, 0, 0)
  return { src: canvas.toDataURL('image/png'), width, height }
}

/** Excel autofits a pivot table's columns on every update (never narrower than the default). */
function autofitPivotColumns(sheet: SheetData, pivot: PivotTableModel, output: PivotOutput) {
  for (let offset = 0; offset < output.columns; offset += 1) {
    let widest = 0
    for (const row of output.cells) {
      const cell = row[offset]
      if (!cell || cell.value === null || cell.value === undefined) continue
      const text = typeof cell.value === 'number' ? formatScalar(cell.value, cell.numFmt || 'General') : String(cell.value)
      const bold = cell.kind !== 'value' && cell.kind !== 'rowKey'
      widest = Math.max(widest, measureTextWidth(text, `${bold ? '700' : '400'} 14.667px Calibri, Aptos, "Segoe UI", sans-serif`))
    }
    const pixels = Math.min(MAX_COLUMN_PIXELS, Math.max(IMPORTED_COL_WIDTH, Math.ceil(widest + 14)))
    sheet.colWidths[String(pivot.anchor.col + offset + 1)] = modelColumnWidth(pixels, 1)
  }
}

/** Sheets whose table header cells a change touched (from immer patches), with the addresses. */
function tableHeaderTouches(workbook: WorkbookModel, patches: Patch[]) {
  const touched = new Map<string, Set<string>>()
  for (const patch of patches) {
    if (patch.path[0] !== 'sheets' || patch.path[2] !== 'cells') continue
    const sheet = workbook.sheets[Number(patch.path[1])]
    if (!sheet?.tables?.length) continue
    const address = patch.path.length > 3 ? String(patch.path[3]) : null
    const coord = address ? coordOf(address) : null
    const hitsHeader = sheet.tables.some((table) => {
      const regions = tableRegions(table)
      return regions && regions.header !== null && (!coord || (coord.row === regions.header && coord.col >= regions.left && coord.col <= regions.right))
    })
    if (!hitsHeader) continue
    const set = touched.get(sheet.id) || new Set<string>()
    if (address) set.add(address)
    else set.add('*')
    touched.set(sheet.id, set)
  }
  return touched
}

/** Excel's typing behaviours around tables: auto-expansion, header names, calculated columns. */
function applyTableEntry(workbook: WorkbookModel, sheet: SheetData, address: string, coord: Coord) {
  expandTableForEntry(sheet, coord.row, coord.col)
  const table = tableContaining(sheet, coord.row, coord.col)
  if (!table) return
  const regions = tableRegions(table)
  if (regions?.header === coord.row) syncTableHeaders(workbook, sheet, [address])
  else if (sheet.cells[address]?.formula && regions && coord.row >= regions.dataTop && coord.row <= regions.dataBottom) fillCalculatedColumn(sheet, table, coord.row, coord.col)
}

/** A table's AutoFilter state, with its range following the table's current header and data rows. */
function tableFilterState(table: SheetTable): SheetFilterState | null {
  const regions = tableRegions(table)
  if (!regions || regions.header === null) return null
  const ref = formatTableRef({ top: regions.header, bottom: Math.max(regions.header, regions.dataBottom), left: regions.left, right: regions.right })
  return { ref, columns: table.filter?.columns || {}, ...(table.filter?.sort ? { sort: table.filter.sort } : {}) }
}

/** Hidden rows for every filter on a sheet (its AutoFilter and each table's), as a union. */
function refreshAllFilters(sheet: SheetData, host: DataHost, overrides: { sheet?: SheetFilterState | null; tables?: Map<string, SheetFilterState | null> } = {}) {
  const rows = new Set<number>()
  const sheetState = overrides.sheet !== undefined ? overrides.sheet : (sheet.filter ?? null)
  if (sheetState && filterIsActive(sheetState)) for (const row of computeFilteredRows(sheetState, host)) rows.add(row)
  for (const table of sheet.tables || []) {
    const state = overrides.tables?.has(table.id) ? overrides.tables.get(table.id) ?? null : tableFilterState(table)
    if (state && filterIsActive(state)) for (const row of computeFilteredRows(state, host)) rows.add(row)
  }
  return applyFilterResult(sheet, [...rows])
}

function hasCellContent(cell: CellData) {
  return cell.value !== undefined || Boolean(cell.formula) || Boolean(cell.style) || Boolean(cell.numFmt) || Boolean(cell.note) || Boolean(cell.hyperlink)
}

function blankWorkbook(): WorkbookModel {
  const sheetId = makeId()
  return {
    version: 1,
    name: 'Untitled.xlsx',
    activeSheetId: sheetId,
    sheets: [{
      id: sheetId,
      name: 'Sheet1',
      state: 'visible',
      rowCount: DEFAULT_ROWS,
      colCount: DEFAULT_COLS,
      cells: {},
      merges: [],
      colWidths: {},
      rowHeights: {},
      frozen: { rows: 0, columns: 0 },
    }],
    metadata: { creator: 'simple_calc' },
  }
}

function normalizeWorkbook(input: WorkbookModel, name = 'Untitled.xlsx'): WorkbookModel {
  const workbook = fromFileWorkbook(structuredClone(input || blankWorkbook()) as WorkbookModel)
  workbook.version = 1
  workbook.name ||= name
  workbook.sheets = Array.isArray(workbook.sheets) && workbook.sheets.length ? workbook.sheets : blankWorkbook().sheets
  const usedIds = new Set<string>()
  workbook.sheets.forEach((sheet, index) => {
    if (!sheet.id || usedIds.has(sheet.id)) sheet.id = makeId()
    usedIds.add(sheet.id)
    sheet.name ||= `Sheet${index + 1}`
    sheet.state ||= 'visible'
    sheet.cells ||= {}
    sheet.merges ||= []
    sheet.colWidths ||= {}
    sheet.rowHeights ||= {}
    sheet.rowCount = clamp(Math.max(DEFAULT_ROWS, Number(sheet.rowCount) || 0), DEFAULT_ROWS, 1_048_576)
    sheet.colCount = clamp(Math.max(DEFAULT_COLS, Number(sheet.colCount) || 0), DEFAULT_COLS, 16_384)
    const normalizedCells: Record<string, CellData> = {}
    Object.entries(sheet.cells).forEach(([address, cell]) => { normalizedCells[address.toUpperCase()] = cell || {} })
    sheet.cells = markArrayMembers({ ...sheet, cells: normalizedCells }).cells
  })
  const active = workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId)
  if (!active || active.state !== 'visible') workbook.activeSheetId = workbook.sheets.find((sheet) => sheet.state === 'visible')?.id || workbook.sheets[0].id
  return workbook
}

function readRecent(): RecentWorkbook[] {
  try {
    const parsed = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]')
    return Array.isArray(parsed) ? parsed.filter((item) => item?.path && item?.name).slice(0, 8) : []
  } catch {
    return []
  }
}

function rememberRecent(current: RecentWorkbook[], item: RecentWorkbook) {
  return [item, ...current.filter((entry) => entry.path.toLowerCase() !== item.path.toLowerCase())].slice(0, 8)
}

function writeRecent(items: RecentWorkbook[]) {
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(items))
  } catch {
    return
  }
}

function errorMessage(error: unknown) {
  if (error instanceof Error) return error.message
  if (typeof error === 'string') return error
  return 'Something went wrong.'
}

const OFFICE_THEME_COLORS = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31',
  'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72',
]

// Workbooks address theme colours by index, and every file carries its own palette — an
// Excel 2007/2010 file resolves accent2 to brick red where the modern default is orange.
// Held at module scope because cssColor is a pure style helper reached from every cell and a
// window shows one workbook at a time.
let activeThemeColors: readonly string[] = OFFICE_THEME_COLORS

function applyWorkbookThemeColors(colors: unknown) {
  activeThemeColors = Array.isArray(colors) && colors.length
    ? OFFICE_THEME_COLORS.map((fallback, index) => {
      const value = String(colors[index] ?? '').replace(/^#/, '')
      return /^[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : fallback
    })
    : OFFICE_THEME_COLORS
}

const EXCEL_INDEXED_COLORS: Record<number, string> = {
  0: '000000', 1: 'FFFFFF', 2: 'FF0000', 3: '00FF00', 4: '0000FF', 5: 'FFFF00', 6: 'FF00FF', 7: '00FFFF',
  8: '000000', 9: 'FFFFFF', 10: 'FF0000', 11: '00FF00', 12: '0000FF', 13: 'FFFF00', 14: 'FF00FF', 15: '00FFFF',
  16: '800000', 17: '008000', 18: '000080', 19: '808000', 20: '800080', 21: '008080', 22: 'C0C0C0', 23: '808080',
  24: '9999FF', 25: '993366', 26: 'FFFFCC', 27: 'CCFFFF', 28: '660066', 29: 'FF8080', 30: '0066CC', 31: 'CCCCFF',
  32: '000080', 33: 'FF00FF', 34: 'FFFF00', 35: '00FFFF', 36: '800080', 37: '800000', 38: '008080', 39: '0000FF',
  40: '00CCFF', 41: 'CCFFFF', 42: 'CCFFCC', 43: 'FFFF99', 44: '99CCFF', 45: 'FF99CC', 46: 'CC99FF', 47: 'FFCC99',
  48: '3366FF', 49: '33CCCC', 50: '99CC00', 51: 'FFCC00', 52: 'FF9900', 53: 'FF6600', 54: '666699', 55: '969696',
  56: '003366', 57: '339966', 58: '003300', 59: '333300', 60: '993300', 61: '993366', 62: '333399', 63: '333333',
}

function tintHex(hex: string, tint = 0) {
  if (!Number.isFinite(tint) || tint === 0) return hex
  // Excel applies tints in HSL lightness space.
  return applyTint(hex, clamp(tint, -1, 1)).replace(/^#/, '').toUpperCase()
}

function cssColor(value: unknown, fallback = '') {
  let color = ''
  let tint = 0
  if (typeof value === 'string') color = value
  else if (value && typeof value === 'object') {
    const object = value as SpreadsheetColor
    color = object.argb || object.rgb || ''
    tint = Number(object.tint) || 0
    if (!color && Number.isInteger(object.theme)) color = activeThemeColors[Number(object.theme)] || ''
    if (!color && Number.isInteger(object.indexed)) color = EXCEL_INDEXED_COLORS[Number(object.indexed)] || ''
  }
  color = color.replace(/^#/, '')
  if (color.length === 8) color = color.slice(2)
  if (/^[0-9a-f]{6}$/i.test(color)) return `#${tintHex(color.toUpperCase(), tint)}`
  return fallback
}

function fontCss(font: CellFont = {}, scale = 1): CSSProperties {
  const underline = font.underline
  const decorations = [underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ')
  const family = String(font.name || '').replace(/["\\\r\n]/g, '').trim()
  const fallback = Number(font.family) === 1 ? '"Times New Roman", serif'
    : Number(font.family) === 3 ? 'Consolas, monospace' : 'Calibri, Aptos, "Segoe UI", Arial, sans-serif'
  return {
    fontFamily: family ? `"${family}", ${fallback}` : undefined,
    fontSize: font.size ? `${font.size * (4 / 3) * scale}px` : undefined,
    fontWeight: font.bold ? 700 : undefined,
    fontStyle: font.italic ? 'italic' : undefined,
    fontStretch: font.condense ? 'condensed' : font.extend ? 'expanded' : undefined,
    textDecorationLine: decorations || undefined,
    textDecorationStyle: typeof underline === 'string' && /double/i.test(underline) ? 'double' : undefined,
    textShadow: font.shadow ? '0.08em 0.1em 0.12em rgba(0, 0, 0, 0.28)' : undefined,
    WebkitTextStroke: font.outline ? '0.35px currentColor' : undefined,
    color: cssColor(font.color) || undefined,
  }
}

function gradientStop(stop: unknown, index: number, total: number) {
  const item = stop && typeof stop === 'object' ? stop as { position?: number; color?: unknown } : {}
  const position = Number.isFinite(item.position) ? clamp(Number(item.position), 0, 1) : total > 1 ? index / (total - 1) : 0
  return `${cssColor(item.color, '#ffffff')} ${Math.round(position * 100)}%`
}

function fillCss(fill: CellFill = {}): CSSProperties {
  const foreground = cssColor(fill.color || fill.fgColor)
  const background = cssColor(fill.bgColor, '#ffffff')
  const stops = Array.isArray(fill.stops) ? fill.stops : []
  if (String(fill.type).toLocaleLowerCase() === 'gradient' && stops.length) {
    const renderedStops = stops.map((stop, index) => gradientStop(stop, index, stops.length)).join(', ')
    if (fill.center) {
      const centerLeft = Number(fill.center.left)
      const centerTop = Number(fill.center.top)
      const x = clamp((Number.isFinite(centerLeft) ? centerLeft : 0.5) * 100, 0, 100)
      const y = clamp((Number.isFinite(centerTop) ? centerTop : 0.5) * 100, 0, 100)
      return { backgroundColor: background, backgroundImage: `radial-gradient(circle at ${x}% ${y}%, ${renderedStops})` }
    }
    return { backgroundColor: background, backgroundImage: `linear-gradient(${90 + (Number(fill.degree) || 0)}deg, ${renderedStops})` }
  }

  const pattern = String(fill.pattern || '').toLocaleLowerCase()
  if (!pattern || pattern === 'none') return foreground ? { backgroundColor: foreground } : {}
  if (pattern === 'solid') return { backgroundColor: foreground || background }
  const ink = foreground || '#808080'
  const sparse = /light|gray0625/.test(pattern)
  const step = sparse ? 8 : 5
  const line = sparse ? 1 : 2
  const direction = /vertical/.test(pattern) ? '90deg'
    : /down/.test(pattern) ? '45deg'
      : /up/.test(pattern) ? '-45deg' : '0deg'
  const stripe = `repeating-linear-gradient(${direction}, transparent 0, transparent ${step - line}px, ${ink} ${step - line}px, ${ink} ${step}px)`
  if (/grid|trellis/.test(pattern)) {
    const cross = `repeating-linear-gradient(${direction === '90deg' ? '0deg' : '90deg'}, transparent 0, transparent ${step - line}px, ${ink} ${step - line}px, ${ink} ${step}px)`
    return { backgroundColor: background, backgroundImage: `${stripe}, ${cross}` }
  }
  if (/gray/.test(pattern)) {
    return { backgroundColor: background, backgroundImage: `radial-gradient(${ink} ${sparse ? '0.65px' : '1px'}, transparent ${sparse ? '0.8px' : '1.2px'})`, backgroundSize: `${step}px ${step}px` }
  }
  return { backgroundColor: background, backgroundImage: stripe }
}

// OOXML draws nothing for a side with no style, and nothing for style="none" — even when
// the side carries a colour.  Defaulting those to a thin rule invents lines through the
// middle of imported sheets.
function drawableBorderSide(side: CellBorderSide | undefined): side is CellBorderSide {
  if (!side || typeof side !== 'object') return false
  const style = String(side.style || '').toLocaleLowerCase()
  return Boolean(style) && style !== 'none'
}

function borderCss(side: CellBorderSide | undefined) {
  if (!drawableBorderSide(side)) return undefined
  const style = String(side.style).toLocaleLowerCase()
  const color = cssColor(side.color, '#4b4b4b')
  const width = /thick/.test(style) ? 3 : /medium|double/.test(style) ? 2 : 1
  const line = /double/.test(style) ? 'double' : /dot/.test(style) ? 'dotted' : /dash/.test(style) ? 'dashed' : 'solid'
  return `${line === 'double' ? Math.max(3, width) : width}px ${line} ${color}`
}

function cellCss(cell?: CellData, resolved?: CellScalar, scale = 1, height?: number): CSSProperties {
  const style = cell?.style || {}
  const alignment = style.alignment || {}
  const border = style.border || {}
  // Formula cells are laid out from the value the sheet currently evaluates to.  Workbooks
  // written by tools other than Excel routinely ship formulas with an empty <v/>, so the
  // cached result is absent and alignment would silently fall back to the text defaults.
  const rawValue = resolved !== undefined ? resolved : cell?.formula ? cell.result : cell?.value
  const horizontal = String(alignment.horizontal || '').toLocaleLowerCase()
  const vertical = String(alignment.vertical || '').toLocaleLowerCase()
  // Imported rows use point sizes, but browser line boxes include leading. Keep
  // that leading inside a tight row so 12pt numerals fit a normal 15pt row.
  // Do not resize the source row, font, or the user's chosen vertical alignment.
  const fontPixels = (Number(style.font?.size) || 11) * (4 / 3) * scale
  const borderPixels = (side: CellBorderSide | undefined, fallback = 0) => {
    if (!drawableBorderSide(side)) return fallback
    return /thick|double/i.test(String(side?.style)) ? 3 : /medium/i.test(String(side?.style)) ? 2 : 1
  }
  const availableHeight = height === undefined ? Infinity : height - 2 * scale - borderPixels(border.top) - borderPixels(border.bottom, 1)
  const compactLineHeight = height !== undefined && style.font?.size
    ? `${Math.min(fontPixels * 1.16, Math.max(fontPixels, availableHeight))}px` : undefined
  const indent = Math.max(0, Number(alignment.indent || 0) + Number(alignment.relativeIndent || 0))
  // Excel's General alignment: numbers right, text left, logical and error values centred.
  const inferredAlignment = typeof rawValue === 'number' ? 'right' : typeof rawValue === 'boolean' || isFormulaError(rawValue) ? 'center' : undefined
  const accountingAlignment = typeof rawValue === 'number' && isAccountingNumberFormat(cell?.numFmt || style.numFmt)
  const textAlign = accountingAlignment ? 'right'
    : horizontal === 'right' ? 'right'
    : horizontal === 'center' || horizontal === 'centercontinuous' ? 'center'
      : horizontal === 'justify' || horizontal === 'distributed' ? 'justify'
        : horizontal === 'left' ? 'left' : inferredAlignment
  const diagonal = border.diagonal
  const diagonalColor = cssColor(diagonal && typeof diagonal === 'object' ? diagonal.color : undefined, '#4b4b4b')
  const diagonalImages = [
    border.diagonalUp ? `linear-gradient(to top right, transparent calc(50% - .5px), ${diagonalColor} 50%, transparent calc(50% + .5px))` : '',
    border.diagonalDown ? `linear-gradient(to bottom right, transparent calc(50% - .5px), ${diagonalColor} 50%, transparent calc(50% + .5px))` : '',
  ].filter(Boolean)
  const fill = fillCss(style.fill)
  const backgroundImage = [...diagonalImages, fill.backgroundImage ? String(fill.backgroundImage) : ''].filter(Boolean).join(', ') || undefined
  return {
    ...fontCss(style.font, scale),
    ...fill,
    backgroundImage,
    borderTop: borderCss(border.top),
    borderLeft: borderCss(border.left),
    borderBottom: borderCss(border.bottom),
    borderRight: borderCss(border.right),
    alignItems: vertical === 'top' ? 'flex-start' : vertical === 'middle' || vertical === 'center' ? 'safe center' : vertical === 'bottom' ? 'safe flex-end' : undefined,
    lineHeight: compactLineHeight,
    textAlign,
    whiteSpace: alignment.wrapText ? 'pre-wrap' : 'nowrap',
    overflowWrap: alignment.wrapText ? 'anywhere' : undefined,
    direction: alignment.readingOrder === 'rtl' || alignment.readingOrder === 2 ? 'rtl' : alignment.readingOrder === 'ltr' || alignment.readingOrder === 1 ? 'ltr' : undefined,
    paddingLeft: indent && textAlign !== 'right' ? `${6 + indent * 9 * scale}px` : undefined,
    paddingRight: indent && textAlign === 'right' ? `${6 + indent * 9 * scale}px` : undefined,
  }
}

function cellContentCss(cell: CellData | undefined, width?: number): CSSProperties {
  const alignment = cell?.style?.alignment || {}
  const font = cell?.style?.font || {}
  const rotation = alignment.textRotation
  const isScript = font.vertAlign === 'superscript' || font.vertAlign === 'subscript'
  const transforms: string[] = []
  if (typeof rotation === 'number' && rotation !== 0 && rotation !== 255) transforms.push(`rotate(${-rotation}deg)`)
  if (font.vertAlign === 'superscript') transforms.push('translateY(-0.22em)')
  else if (font.vertAlign === 'subscript') transforms.push('translateY(0.22em)')
  return {
    width: width ? `${width}px` : '100%',
    maxWidth: width ? 'none' : '100%',
    writingMode: rotation === 'vertical' || rotation === 255 ? 'vertical-rl' : undefined,
    textOrientation: rotation === 'vertical' || rotation === 255 ? 'upright' : undefined,
    transform: transforms.length ? transforms.join(' ') : undefined,
    transformOrigin: typeof rotation === 'number' && rotation !== 0 && rotation !== 255 ? 'center' : undefined,
    fontSize: isScript ? '82%' : undefined,
    fontStretch: alignment.shrinkToFit ? 'condensed' : undefined,
  }
}

/**
 * Cell addresses touched by a set of immer patches, per sheet id, so the calculation engine
 * can skip diffing large sheets. A patch that replaces a whole sheet or cell map marks the
 * sheet 'all' (the engine then diffs it by identity).
 */
/** Apply a formula rewrite to every formula-bearing part of the workbook (in an immer draft). */
function rewriteWorkbookFormulas(workbook: WorkbookModel, rewrite: (formula: string) => string) {
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

const TAB_COLORS = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0', '4472C4', 'ED7D31', 'A5A5A5', '70AD47']

function formulaHighlightParts(text: string): HighlightPart[] {
  const references = formulaReferences(text)
  return tokenizeFormulaText(text).map((token) => ({ text: token.text, color: tokenColor(token, references) }))
}

function changeHintFromPatches(patches: Patch[], next: WorkbookModel): ChangeHint {
  const hint: ChangeHint = new Map()
  for (const patch of patches) {
    const [root, index, field, address] = patch.path
    if (root !== 'sheets' || typeof index !== 'number') continue
    const sheet = next.sheets[index]
    if (!sheet) continue
    if (field === undefined || (field === 'cells' && address === undefined)) {
      hint.set(sheet.id, 'all')
      continue
    }
    if (field !== 'cells' || typeof address !== 'string') continue
    const existing = hint.get(sheet.id)
    if (existing === 'all') continue
    if (existing) existing.add(address)
    else hint.set(sheet.id, new Set([address]))
  }
  return hint
}

function mergeChangeHints(first: ChangeHint, second: ChangeHint): ChangeHint {
  const merged: ChangeHint = new Map(first)
  for (const [sheetId, addresses] of second) {
    const existing = merged.get(sheetId)
    if (existing === 'all' || addresses === 'all') merged.set(sheetId, 'all')
    else merged.set(sheetId, new Set([...(existing || []), ...addresses]))
  }
  return merged
}

function parseClipboardText(text: string) {
  const rows: string[][] = [[]]
  let value = ''
  let quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === '"') {
      if (quoted && text[index + 1] === '"') { value += '"'; index += 1 }
      else quoted = !quoted
    } else if (character === '\t' && !quoted) {
      rows[rows.length - 1].push(value)
      value = ''
    } else if ((character === '\n' || character === '\r') && !quoted) {
      if (character === '\r' && text[index + 1] === '\n') index += 1
      rows[rows.length - 1].push(value)
      value = ''
      if (index < text.length - 1) rows.push([])
    } else value += character
  }
  rows[rows.length - 1].push(value)
  return rows
}

function escapeClipboard(value: string) {
  return /[\t\n\r"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value
}

/** Maximum digit width (px) of Calibri/Aptos 11pt, the unit Excel column widths are stored in. */
const EXCEL_DIGIT_WIDTH = 7

/** Excel's column width → pixels: Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW). */
function columnPixelWidth(value: number | undefined, zoom: number) {
  if (!value) return IMPORTED_COL_WIDTH * zoom
  const pixels = Math.trunc(((256 * value + Math.trunc(128 / EXCEL_DIGIT_WIDTH)) / 256) * EXCEL_DIGIT_WIDTH)
  return clamp(pixels, 2, 1_000) * zoom
}

function rowPixelHeight(value: number | undefined, zoom: number) {
  if (!value) return DEFAULT_ROW_HEIGHT * zoom
  return clamp(value * (4 / 3), 2, 640) * zoom
}

function modelColumnWidth(pixelWidth: number, zoom: number) {
  // Store Excel's own granularity (1/256 of a character) so a pixel width round-trips exactly.
  return Math.round(clamp(pixelWidth / zoom / EXCEL_DIGIT_WIDTH, 0.2, 255) * 256) / 256
}

function modelRowHeight(pixelHeight: number, zoom: number) {
  return Math.round(clamp(pixelHeight / zoom * (3 / 4), 2, MAX_ROW_POINTS) * 100) / 100
}

interface DimensionResize {
  axis: 'column' | 'row'
  index: number
  pointerId: number
  startClient: number
  startPixels: number
  pixels: number
}

interface FreezeDrag {
  axis: 'column' | 'row'
  pointerId: number
  count: number
}

type CellPane = 'body' | 'frozen-row' | 'frozen-column' | 'frozen-corner'

const NO_MERGES: ReadonlyMap<string, ReturnType<typeof mergeBounds>> = new Map()

function cellPane(frozenRow: boolean, frozenColumn: boolean): CellPane {
  return frozenRow && frozenColumn ? 'frozen-corner' : frozenRow ? 'frozen-row' : frozenColumn ? 'frozen-column' : 'body'
}

// Imported merges may cross a saved freeze boundary. Each pane paints a clipped
// portion of the same full cell; changing the merge or freeze metadata is unnecessary.
function intersectingPanes(bounds: { top: number; bottom: number; left: number; right: number }, frozenRows: number, frozenColumns: number): CellPane[] {
  const rowParts = [true, false].filter((frozen) => frozen ? bounds.top < frozenRows : bounds.bottom >= frozenRows)
  const columnParts = [true, false].filter((frozen) => frozen ? bounds.left < frozenColumns : bounds.right >= frozenColumns)
  return rowParts.flatMap((row) => columnParts.map((column) => cellPane(row, column)))
}

interface AxisMetric {
  index: number
  size: number
}

interface AxisMetricIndex {
  defaultSize: number
  sizeByIndex: Map<number, number>
  prefixDelta: number[]
}

const axisMetricIndexCache = new WeakMap<AxisMetric[], AxisMetricIndex>()

function indexedAxisMetrics(defaultSize: number, custom: AxisMetric[]) {
  const cached = axisMetricIndexCache.get(custom)
  if (cached?.defaultSize === defaultSize) return cached
  let delta = 0
  const prefixDelta = custom.map((metric) => {
    delta += metric.size - defaultSize
    return delta
  })
  const indexed = {
    defaultSize,
    sizeByIndex: new Map(custom.map((metric) => [metric.index, metric.size])),
    prefixDelta,
  }
  axisMetricIndexCache.set(custom, indexed)
  return indexed
}

function axisSize(index: number, defaultSize: number, custom: AxisMetric[]) {
  if (!custom.length) return defaultSize
  return indexedAxisMetrics(defaultSize, custom).sizeByIndex.get(index) ?? defaultSize
}

function axisOffset(index: number, defaultSize: number, custom: AxisMetric[]) {
  if (!custom.length || index <= 0) return Math.max(0, index) * defaultSize
  let low = 0
  let high = custom.length
  while (low < high) {
    const middle = Math.floor((low + high) / 2)
    if (custom[middle].index < index) low = middle + 1
    else high = middle
  }
  const prefix = indexedAxisMetrics(defaultSize, custom).prefixDelta
  return index * defaultSize + (low > 0 ? prefix[low - 1] : 0)
}

function axisIndexAt(pixel: number, count: number, defaultSize: number, custom: AxisMetric[]) {
  let low = 0
  let high = count
  while (low < high) {
    const middle = Math.floor((low + high + 1) / 2)
    if (axisOffset(middle, defaultSize, custom) <= pixel) low = middle
    else high = middle - 1
  }
  return clamp(low, 0, Math.max(0, count - 1))
}

function axisBoundaryAt(pixel: number, count: number, defaultSize: number, custom: AxisMetric[]) {
  if (pixel <= 0 || count <= 0) return 0
  const index = axisIndexAt(pixel, count, defaultSize, custom)
  const start = axisOffset(index, defaultSize, custom)
  const end = axisOffset(Math.min(count, index + 1), defaultSize, custom)
  return clamp(pixel - start < end - pixel ? index : index + 1, 0, count)
}

function mergeBounds(range: string) {
  const [start, end = start] = range.split(':')
  const a = coordOf(start)
  const b = coordOf(end)
  if (!a || !b) return null
  return { top: Math.min(a.row, b.row), bottom: Math.max(a.row, b.row), left: Math.min(a.col, b.col), right: Math.max(a.col, b.col) }
}

// Excel treats a merge as a single cell: reading, editing and committing all go through the
// master, whichever covered address the selection focus happens to sit on.
function mergeMasterAddress(sheet: SheetData | null | undefined, address: string) {
  const coord = coordOf(address)
  if (!sheet || !coord) return address
  for (const range of sheet.merges || []) {
    const bounds = mergeBounds(range)
    if (!bounds) continue
    if (coord.row >= bounds.top && coord.row <= bounds.bottom && coord.col >= bounds.left && coord.col <= bounds.right) {
      return addressOf({ row: bounds.top, col: bounds.left })
    }
  }
  return address
}

function noteText(note: CellData['note']) {
  if (!note) return ''
  if (typeof note === 'string') return note
  if (Array.isArray(note.comments)) return note.comments.map((comment) => `${comment.author ? `${comment.author}: ` : ''}${comment.text || ''}`).join('\n')
  if (Array.isArray(note.texts)) return note.texts.map((item) => item.text || '').join('')
  return typeof note.text === 'string' ? note.text : ''
}

interface ValidationRange {
  bounds: NonNullable<ReturnType<typeof mergeBounds>>
  validation: Record<string, unknown>
}

const validationRangeCache = new WeakMap<Record<string, unknown>, ValidationRange[]>()

function indexedValidationRanges(validations: Record<string, unknown>) {
  const cached = validationRangeCache.get(validations)
  if (cached) return cached
  const ranges: ValidationRange[] = []
  for (const [range, validation] of Object.entries(validations)) {
    if (!validation || typeof validation !== 'object') continue
    const bounds = mergeBounds(range)
    if (bounds) ranges.push({ bounds, validation: validation as Record<string, unknown> })
  }
  validationRangeCache.set(validations, ranges)
  return ranges
}

function validationForCell(sheet: SheetData, address: string, known?: Coord) {
  const validations = sheet.dataValidations
  if (!validations) return undefined
  if (validations[address] && typeof validations[address] === 'object') return validations[address] as Record<string, unknown>
  const coord = known || coordOf(address)
  if (!coord) return undefined
  for (const { bounds, validation } of indexedValidationRanges(validations)) {
    if (coord.row >= bounds.top && coord.row <= bounds.bottom && coord.col >= bounds.left && coord.col <= bounds.right) return validation
  }
  return undefined
}

function validationListOptions(validation: Record<string, unknown> | undefined) {
  if (!validation || validation.type !== 'list' || !Array.isArray(validation.formulae)) return []
  const source = validation.formulae[0]
  if (typeof source !== 'string') return []
  const literal = source.startsWith('"') && source.endsWith('"') ? source.slice(1, -1).replace(/""/g, '"') : ''
  return literal ? literal.split(',').map((item) => item.trim()).filter(Boolean) : []
}

function setValidationForBounds(sheet: SheetData, target: { top: number; bottom: number; left: number; right: number }, validation?: Record<string, unknown>) {
  const next: Record<string, unknown> = {}
  for (const [range, existing] of Object.entries(sheet.dataValidations || {})) {
    const bounds = mergeBounds(range)
    if (!bounds || bounds.bottom < target.top || bounds.top > target.bottom || bounds.right < target.left || bounds.left > target.right) {
      next[range] = existing
      continue
    }
    const pieces = [
      bounds.top < target.top ? { top: bounds.top, bottom: target.top - 1, left: bounds.left, right: bounds.right } : null,
      bounds.bottom > target.bottom ? { top: target.bottom + 1, bottom: bounds.bottom, left: bounds.left, right: bounds.right } : null,
      bounds.left < target.left ? { top: Math.max(bounds.top, target.top), bottom: Math.min(bounds.bottom, target.bottom), left: bounds.left, right: target.left - 1 } : null,
      bounds.right > target.right ? { top: Math.max(bounds.top, target.top), bottom: Math.min(bounds.bottom, target.bottom), left: target.right + 1, right: bounds.right } : null,
    ].filter((piece): piece is { top: number; bottom: number; left: number; right: number } => Boolean(piece && piece.top <= piece.bottom && piece.left <= piece.right))
    pieces.forEach((piece) => { next[rangeAddress(piece)] = existing })
  }
  if (validation) next[rangeAddress(target)] = validation
  sheet.dataValidations = next
}

interface GridProps {
  sheet: SheetData
  selection: Selection
  editing: EditingState | null
  zoom: number
  useImportedDefaults: boolean
  showGridlines: boolean
  showNotes: boolean
  displayValue: (sheetId: string, address: string) => string
  resolvedValue: (sheetId: string, address: string) => CellScalar | undefined
  staleValue: (sheetId: string, address: string) => boolean
  onSelection: (selection: Selection) => void
  onBeginEdit: (draft?: string) => void
  onDraft: (draft: string) => void
  onCommitEdit: (direction?: 'up' | 'down' | 'left' | 'right') => void
  onCancelEdit: () => void
  onFill: (target: Coord) => void
  onCellValue: (address: string, value: CellData['value']) => void
  onOpenHyperlink: (target: string) => void
  onColumnResize: (column: number, width: number) => void
  onRowResize: (row: number, height: number) => void
  onFreeze: (axis: 'rows' | 'columns', count: number) => void
  onKeyDown: (event: KeyboardEvent<HTMLDivElement>) => void
  onContextTarget: (target: GridContextTarget, position: { x: number; y: number }) => void
  referenceHighlights: ReferenceHighlight[]
  spillOutline: OverlayBounds | null
  editorParts: HighlightPart[] | null
  editorSelectionRequest: { start: number; end: number; id: number } | null
  /** Point mode: a click/drag on a cell while editing a formula inserts a reference. */
  onPointCell: (coord: Coord, phase: 'down' | 'drag', extend: boolean) => boolean
  onEditorKeyDown: (event: KeyboardEvent<HTMLTextAreaElement>) => boolean
  onEditorCaret: (start: number, end: number) => void
  onEditorFocusChange: (focused: boolean, element: HTMLTextAreaElement | null) => void
  /** Conditional formats for the cells in a window of the active sheet (0-based bounds). */
  conditionalFormats: (visible: OverlayBounds) => Map<string, ConditionalCellFormat> | null
  /** AutoFilter header: row index, column span, and columns with active criteria. */
  /** Header filter buttons: the sheet AutoFilter ('sheet') and each table with buttons. */
  filterHeaders: Array<{ key: string; row: number; left: number; right: number; active: Set<number>; sort?: { column: number; descending: boolean } }>
  onFilterButton: (key: string, col: number, anchor: { left: number; top: number; right: number; bottom: number }) => void
  inputMessage: { row: number; col: number; title: string; text: string } | null
  invalidCells: Coord[]
  cutRange: OverlayBounds | null
  /** Floating chart layer, drawn in canvas coordinates. */
  renderCharts?: (geometry: ChartGeometry, viewport: { left: number; top: number; width: number; height: number }, geometryVersion: unknown) => React.ReactNode
  listOptionsFor: (validation: Record<string, unknown> | undefined, address: string) => string[]
  /** Excel sparkline for a cell of the active sheet (only passed when the sheet has sparkline groups). */
  sparklineFor?: (address: string) => SparklineSpec | null
  /** Row/column groups (Excel's outline); the outline bars render beside the grid when present. */
  outline?: { rows: OutlineGroup[]; columns: OutlineGroup[]; rowDepth: number; columnDepth: number } | null
  onOutlineToggle?: (axis: OutlineAxis, group: OutlineGroup) => void
  onOutlineLevel?: (axis: OutlineAxis, level: number) => void
}

interface GridCellViewProps {
  sheet: SheetData
  row: number
  col: number
  pane: CellPane
  zoom: number
  merge?: NonNullable<ReturnType<typeof mergeBounds>>
  columns: number
  frozenRowCount: number
  frozenColumnCount: number
  defaultColumn: number
  defaultRow: number
  columnMetrics: AxisMetric[]
  rowMetrics: AxisMetric[]
  mergeRanges: Array<NonNullable<ReturnType<typeof mergeBounds>>>
  selectionRef: MutableRefObject<Selection>
  viewportRef: RefObject<HTMLDivElement | null>
  fillDraggingRef: RefObject<boolean>
  displayValue: GridProps['displayValue']
  resolvedValue: GridProps['resolvedValue']
  staleValue: GridProps['staleValue']
  onSelection: GridProps['onSelection']
  onBeginEdit: GridProps['onBeginEdit']
  onCellValue: GridProps['onCellValue']
  onOpenHyperlink: GridProps['onOpenHyperlink']
  onPointCell: GridProps['onPointCell']
  pointDraggingRef: MutableRefObject<boolean>
  conditional?: ConditionalCellFormat
  listOptionsFor: GridProps['listOptionsFor']
  sparklineFor?: GridProps['sparklineFor']
}

function mergedDisplayCell(
  sheet: SheetData,
  merge: NonNullable<ReturnType<typeof mergeBounds>> | undefined,
  source: CellData | undefined,
) {
  if (!merge) return source
  const sideFrom = (coordinates: Coord[], side: keyof CellBorder) => {
    for (const coordinate of coordinates) {
      const candidate = sheet.cells[addressOf(coordinate)]?.style?.border?.[side]
      if (drawableBorderSide(candidate as CellBorderSide | undefined)) return candidate as CellBorderSide
    }
    return undefined
  }
  const top = Array.from({ length: merge.right - merge.left + 1 }, (_, offset) => ({ row: merge.top, col: merge.left + offset }))
  const bottom = Array.from({ length: merge.right - merge.left + 1 }, (_, offset) => ({ row: merge.bottom, col: merge.left + offset }))
  const left = Array.from({ length: merge.bottom - merge.top + 1 }, (_, offset) => ({ row: merge.top + offset, col: merge.left }))
  const right = Array.from({ length: merge.bottom - merge.top + 1 }, (_, offset) => ({ row: merge.top + offset, col: merge.right }))
  const sourceStyle = source?.style || {}
  const sourceBorder = sourceStyle.border || {}
  const border = {
    ...sourceBorder,
    top: sideFrom(top, 'top') || sourceBorder.top,
    bottom: sideFrom(bottom, 'bottom') || sourceBorder.bottom,
    left: sideFrom(left, 'left') || sourceBorder.left,
    right: sideFrom(right, 'right') || sourceBorder.right,
  }
  return { ...(source || {}), style: { ...sourceStyle, border } }
}

const GridCellView = memo(function GridCellView({
  sheet,
  row,
  col,
  pane,
  zoom,
  merge,
  columns,
  frozenRowCount,
  frozenColumnCount,
  defaultColumn,
  defaultRow,
  columnMetrics,
  rowMetrics,
  mergeRanges,
  selectionRef,
  viewportRef,
  fillDraggingRef,
  displayValue,
  resolvedValue,
  staleValue,
  onSelection,
  onBeginEdit,
  onCellValue,
  onOpenHyperlink,
  onPointCell,
  pointDraggingRef,
  conditional,
  listOptionsFor,
  sparklineFor,
}: GridCellViewProps) {
  const address = addressOf({ row, col })
  const primaryPane = cellPane(row < frozenRowCount, col < frozenColumnCount)
  const continuation = pane !== primaryPane
  const left = (pane === 'body' ? HEADER_WIDTH : 0) + axisOffset(col, defaultColumn, columnMetrics)
  const top = (pane === 'body' ? HEADER_HEIGHT : 0) + axisOffset(row, defaultRow, rowMetrics)
  const width = merge
    ? axisOffset(merge.right + 1, defaultColumn, columnMetrics) - axisOffset(merge.left, defaultColumn, columnMetrics)
    : axisSize(col, defaultColumn, columnMetrics)
  const height = merge
    ? axisOffset(merge.bottom + 1, defaultRow, rowMetrics) - axisOffset(merge.top, defaultRow, rowMetrics)
    : axisSize(row, defaultRow, rowMetrics)
  const cell = mergedDisplayCell(sheet, merge, sheet.cells[address])
  const rawValue = resolvedValue(sheet.id, address)
  const pixelRatio = window.devicePixelRatio || 1
  const pixel = (value: number) => Math.round(value * pixelRatio) / pixelRatio
  const paintedLeft = pixel(left), paintedTop = pixel(top)
  const tableEntry = sheet.tables?.length ? tableAt(sheet.tables, row, col) : null
  const validation = sheet.dataValidations ? validationForCell(sheet, address, { row, col }) : undefined
  const handleMouseDown = (event: MouseEvent<HTMLDivElement>, isDropdown: boolean) => {
    if (event.button !== 0) return
    if (onPointCell(merge ? { row: merge.top, col: merge.left } : { row, col }, 'down', event.shiftKey)) {
      // Keep focus in the formula being edited; the click inserted a reference.
      event.preventDefault()
      pointDraggingRef.current = true
      return
    }
    event.preventDefault()
    viewportRef.current?.focus({ preventScroll: true })
    const target = merge ? { row: merge.top, col: merge.left } : { row, col }
    const anchor = merge ? { row: merge.bottom, col: merge.right } : target
    onSelection(event.shiftKey
      ? { anchor: selectionRef.current.anchor, focus: merge ? { row: merge.bottom, col: merge.right } : target }
      : { anchor, focus: target })
    if (continuation && isDropdown) {
      // Only the master participates in focus/validation. A clipped visual
      // copy still opens that same native picker from the pointer gesture.
      const master = document.getElementById(`cell-${sheet.id}-${address}`)?.querySelector<HTMLSelectElement>('.cell-dropdown')
      master?.focus({ preventScroll: true })
      try { master?.showPicker() } catch { /* Keyboard selection remains available on the focused master. */ }
    }
  }
  const handleMouseEnter = (event: MouseEvent<HTMLDivElement>) => {
    if (pointDraggingRef.current) {
      if (event.buttons === 1) onPointCell(merge ? { row: merge.bottom, col: merge.right } : { row, col }, 'drag', true)
      else pointDraggingRef.current = false
      return
    }
    if (!fillDraggingRef.current && event.buttons === 1) onSelection({ anchor: selectionRef.current.anchor, focus: merge ? { row: merge.bottom, col: merge.right } : { row, col } })
  }
  const excelSparkline = sparklineFor ? sparklineFor(address) : null
  if (!cell && !merge && !conditional && !tableEntry && !validation && !excelSparkline && (rawValue === null || rawValue === undefined)) {
    // Most of a visible window is blank: paint it with the stylesheet defaults and skip the
    // formatting pipeline and content node, which keeps scrolling cheap on large windows.
    return (
      <div
        id={`cell-${sheet.id}-${address}`}
        data-cell-address={address}
        data-cell-pane={pane}
        className={pane !== 'body' ? 'grid-cell is-frozen' : 'grid-cell'}
        style={{ left: paintedLeft, top: paintedTop, width: pixel(left + width) - paintedLeft, height: pixel(top + height) - paintedTop }}
        role="gridcell"
        aria-rowindex={row + 1}
        aria-colindex={col + 1}
        onMouseDown={(event) => handleMouseDown(event, false)}
        onMouseEnter={handleMouseEnter}
        onDoubleClick={() => onBeginEdit()}
      />
    )
  }
  const paintedStyle = cellCss(cell, rawValue, zoom, height)
  if (tableEntry) {
    // Table styles paint underneath direct cell formatting, as in Excel.
    const paint = tableCellPaint(tableEntry, row, col, activeThemeColors)
    const own = cell?.style
    if (paint) {
      if (paint.fill && !own?.fill?.pattern && !own?.fill?.type) paintedStyle.backgroundColor = paint.fill
      if (paint.color && !own?.font?.color) paintedStyle.color = paint.color
      if (paint.bold && own?.font?.bold === undefined) paintedStyle.fontWeight = 700
      if (paint.borderTop && !paintedStyle.borderTop) paintedStyle.borderTop = paint.borderTop
      if (paint.borderBottom && !paintedStyle.borderBottom) paintedStyle.borderBottom = paint.borderBottom
      if (paint.borderLeft && !paintedStyle.borderLeft) paintedStyle.borderLeft = paint.borderLeft
      if (paint.borderRight && !paintedStyle.borderRight) paintedStyle.borderRight = paint.borderRight
    }
  }
  const formatTextColor = typeof rawValue === 'number' || typeof rawValue === 'string' ? formatColor(rawValue, conditional?.numFmt || cell?.numFmt || cell?.style?.numFmt) : undefined
  if (formatTextColor) paintedStyle.color = formatTextColor
  if (conditional) {
    if (conditional.fill || conditional.colorScale) {
      paintedStyle.backgroundColor = conditional.colorScale || conditional.fill
      paintedStyle.backgroundImage = undefined
    }
    if (conditional.font?.color) paintedStyle.color = conditional.font.color
    if (conditional.font?.bold !== undefined) paintedStyle.fontWeight = conditional.font.bold ? 700 : 400
    if (conditional.font?.italic !== undefined) paintedStyle.fontStyle = conditional.font.italic ? 'italic' : 'normal'
    if (conditional.font?.underline || conditional.font?.strike) {
      paintedStyle.textDecorationLine = [conditional.font.underline ? 'underline' : '', conditional.font.strike ? 'line-through' : ''].filter(Boolean).join(' ')
    }
    if (conditional.border?.top) paintedStyle.borderTop = conditional.border.top
    if (conditional.border?.right) paintedStyle.borderRight = conditional.border.right
    if (conditional.border?.bottom) paintedStyle.borderBottom = conditional.border.bottom
    if (conditional.border?.left) paintedStyle.borderLeft = conditional.border.left
    if (conditional.dataBar) {
      const bar = dataBarBackground(conditional.dataBar)
      const base = paintedStyle.backgroundColor ? String(paintedStyle.backgroundColor) : ''
      paintedStyle.backgroundImage = bar.backgroundImage
      paintedStyle.backgroundSize = bar.backgroundSize
      paintedStyle.backgroundPosition = bar.backgroundPosition
      paintedStyle.backgroundRepeat = bar.backgroundRepeat
      if (base) paintedStyle.backgroundColor = base
    }
  }
  const hideConditionalValue = Boolean(conditional && ((conditional.dataBar && !conditional.dataBar.showValue) || (conditional.icon && !conditional.icon.showValue)))
  if (!merge && cell?.style?.border) {
    // A shared edge is painted once when both cells specify the same border.
    // Different borders retain their explicit source appearance.
    const aboveBottom = sheet.cells[addressOf({ row: row - 1, col })]?.style?.border?.bottom
    const previousRight = sheet.cells[addressOf({ row, col: col - 1 })]?.style?.border?.right
    if (row > 0 && !sheet.hiddenRows?.includes(row) && cell.style.border.top && aboveBottom && borderCss(cell.style.border.top) === borderCss(aboveBottom)) paintedStyle.borderTop = 'none'
    if (col > 0 && !sheet.hiddenCols?.includes(col) && cell.style.border.left && previousRight && borderCss(cell.style.border.left) === borderCss(previousRight)) paintedStyle.borderLeft = 'none'
  }
  const formattedText = conditional?.numFmt && typeof rawValue === 'number' ? formatScalar(rawValue, conditional.numFmt) : displayValue(sheet.id, address)
  const staleResult = Boolean(cell?.formula) && staleValue(sheet.id, address)
  const cellFont = cell?.style?.font
  const cellNumFmt = cell?.numFmt || cell?.style?.numFmt
  // Numbers never truncate: General numbers drop precision to fit, others show "####".
  const cellText = typeof rawValue === 'number' && formattedText && !formattedText.startsWith('=') && !cell?.style?.alignment?.textRotation
    ? fitNumericText(
      formattedText,
      rawValue,
      !cellNumFmt || cellNumFmt === 'General',
      width - 7 * zoom - Math.max(0, Number(cell?.style?.alignment?.indent || 0)) * 9 * zoom,
      `${cellFont?.italic ? 'italic ' : ''}${cellFont?.bold ? '700' : '400'} ${(Number(cellFont?.size) || 11) * (4 / 3) * zoom}px ${cellFont?.name ? `"${String(cellFont.name).replace(/"/g, '')}", ` : ''}Calibri, Aptos, "Segoe UI", sans-serif`,
      (Number(cellFont?.size) || 11) * (4 / 3) * zoom,
    )
    : formattedText
  const accounting = cellText === formattedText ? accountingDisplayParts(cellText, cellNumFmt) : null
  const validationOptions = validation?.type === 'list' && validation.showDropDown !== true ? listOptionsFor(validation, address) : []
  const isCheckbox = cell?.type === 'checkbox' || (
    typeof rawValue === 'boolean' && validationOptions.length === 2 &&
    validationOptions[0].toLocaleUpperCase() === 'TRUE' && validationOptions[1].toLocaleUpperCase() === 'FALSE'
  )
  const isDropdown = !isCheckbox && (cell?.type === 'dropdown' || validationOptions.length > 0)
  const horizontal = String(cell?.style?.alignment?.horizontal || '').toLocaleLowerCase()
  const canOverflow = Boolean(
    cellText && !merge && !cell?.style?.alignment?.wrapText && !cell?.style?.alignment?.clipText && !isCheckbox && !isDropdown &&
    typeof rawValue !== 'number' && typeof rawValue !== 'boolean' && !isFormulaError(rawValue) &&
    !['right', 'center', 'centercontinuous', 'justify', 'distributed', 'fill'].includes(horizontal) &&
    !cell?.style?.alignment?.textRotation
  )
  let contentWidth = width
  if (canOverflow) {
    for (let nextCol = col + 1; nextCol < columns && nextCol <= col + 50; nextCol += 1) {
      if (nextCol === frozenColumnCount) break
      const nextAddress = addressOf({ row, col: nextCol })
      const nextCell = sheet.cells[nextAddress]
      const blockedByMerge = mergeRanges.some((range) => row >= range.top && row <= range.bottom && nextCol >= range.left && nextCol <= range.right)
      if (blockedByMerge || nextCell?.formula || (nextCell?.value !== null && nextCell?.value !== undefined && nextCell?.value !== '')) break
      contentWidth += axisSize(nextCol, defaultColumn, columnMetrics)
    }
  }
  return (
    <div
      id={`cell-${sheet.id}-${address}${continuation ? `-${pane}` : ''}`}
      data-cell-address={address}
      data-cell-pane={pane}
      className={`grid-cell${cell?.formula ? ' has-formula' : ''}${staleResult ? ' stale-result' : ''}${cell?.note ? ' has-note' : ''}${cell?.style?.fill ? ' has-fill' : ''}${paintedStyle.borderTop !== undefined || paintedStyle.borderBottom !== undefined || paintedStyle.borderLeft !== undefined || paintedStyle.borderRight !== undefined ? ' has-border' : ''}${canOverflow && contentWidth > width ? ' can-overflow' : ''}${pane !== 'body' ? ' is-frozen' : ''}`}
      style={{ left: paintedLeft, top: paintedTop, width: pixel(left + width) - paintedLeft, height: pixel(top + height) - paintedTop, ...paintedStyle }}
      role={continuation ? 'presentation' : 'gridcell'}
      aria-hidden={continuation || undefined}
      aria-rowspan={merge && !continuation ? merge.bottom - merge.top + 1 : undefined}
      aria-colspan={merge && !continuation ? merge.right - merge.left + 1 : undefined}
      aria-rowindex={row + 1}
      aria-colindex={col + 1}
      title={staleResult ? 'Value from last file save — formula not recalculated' : noteText(cell?.note) || cell?.hyperlinkTooltip || (cell?.formula ? `=${cell.formula}` : cellText || undefined)}
      onMouseDown={(event) => handleMouseDown(event, isDropdown)}
      onMouseEnter={handleMouseEnter}
      onDoubleClick={() => onBeginEdit()}
    >
      {excelSparkline && (
        <span
          className="cell-sparkline is-background"
          aria-hidden="true"
          // Markup from numbers and validated colours only (sparkline-render.ts).
          dangerouslySetInnerHTML={{ __html: renderSparklineSvg(excelSparkline, Math.max(4, width - 6 * zoom), Math.max(4, height - 4 * zoom)) }}
        />
      )}
      {isCheckbox ? (
        <input
          type="checkbox"
          className="cell-checkbox"
          tabIndex={continuation ? -1 : undefined}
          aria-label={`${address} checkbox`}
          checked={rawValue === true || String(rawValue).toLocaleUpperCase() === 'TRUE'}
          onMouseDown={(event) => {
            if (continuation) event.preventDefault()
            event.stopPropagation()
            viewportRef.current?.focus({ preventScroll: true })
            onSelection({ anchor: { row, col }, focus: { row, col } })
          }}
          onChange={(event) => onCellValue(address, event.target.checked)}
        />
      ) : isDropdown ? (
        <select
          className="cell-dropdown"
          ref={(node) => { if (node) node.inert = continuation }}
          tabIndex={continuation ? -1 : undefined}
          aria-label={`${address} dropdown`}
          value={rawValue == null ? '' : String(rawValue)}
          onMouseDown={(event) => {
            event.stopPropagation()
            onSelection({ anchor: { row, col }, focus: { row, col } })
          }}
          onChange={(event) => onCellValue(address, event.target.value)}
        >
          <option value="">Select…</option>
          {validationOptions.map((option) => <option key={option} value={option}>{option}</option>)}
        </select>
      ) : cell?.hyperlink ? (
        <button
          type="button"
          className="cell-content cell-hyperlink"
          tabIndex={continuation ? -1 : undefined}
          // Shrink-wrapped rather than stretched across the overflow run: a stretched button
          // would arm the link — and swallow the click — over every empty cell it spans.
          style={{
            ...cellContentCss(cell, canOverflow ? contentWidth : undefined),
            ...(canOverflow && contentWidth > width ? { width: 'max-content', maxWidth: `${contentWidth}px` } : null),
          }}
          onClick={(event) => { event.stopPropagation(); onOpenHyperlink(cell.hyperlink!) }}
        >{cellText || cell.hyperlink}</button>
      ) : isSparklineValue(rawValue) ? (
        <span
          className="cell-sparkline"
          aria-label="Sparkline"
          // eslint-disable-next-line react/no-danger -- markup built from numbers and an allowlisted palette
          dangerouslySetInnerHTML={{ __html: (() => { const spec = parseSparkline(rawValue); return spec ? renderSparklineSvg(spec, Math.max(4, width - 6 * zoom), Math.max(4, height - 4 * zoom)) : '' })() }}
        />
      ) : (
        <>
          {conditional?.icon && <ConditionalIcon className="cell-cf-icon" set={conditional.icon.set} index={conditional.icon.index} size={Math.max(10, Math.round(13 * zoom))} />}
          <span className={`cell-content${accounting ? ' cell-accounting' : ''}`} style={cellContentCss(cell, canOverflow ? contentWidth : undefined)}>
            {hideConditionalValue ? null : accounting
              ? <><span>{accounting.symbol}</span><span>{accounting.amount}</span></>
              : cellText}
          </span>
        </>
      )}
    </div>
  )
})

type GridCellRowProps = Omit<GridCellViewProps, 'row' | 'col' | 'merge' | 'conditional'> & {
  coordinates: Array<{ row: number; col: number }>
  mergeMap: ReadonlyMap<string, ReturnType<typeof mergeBounds>>
  conditionalMap: Map<string, ConditionalCellFormat> | null
}

const GridCellRow = memo(function GridCellRow({ coordinates, mergeMap, conditionalMap, ...cellProps }: GridCellRowProps) {
  return <>{coordinates.map(({ row, col }) => {
    const address = addressOf({ row, col })
    return (
      <GridCellView
        key={address}
        {...cellProps}
        row={row}
        col={col}
        merge={mergeMap.get(address) ?? undefined}
        conditional={conditionalMap?.get(address)}
      />
    )
  })}</>
}, gridRowPropsEqual)

/** Rows re-render only when something they draw changed (merges follow `mergeRanges`). */
function gridRowPropsEqual(previous: GridCellRowProps, next: GridCellRowProps) {
  if (
    previous.conditionalMap !== next.conditionalMap ||
    previous.sheet !== next.sheet || previous.pane !== next.pane || previous.zoom !== next.zoom || previous.columns !== next.columns ||
    previous.frozenRowCount !== next.frozenRowCount || previous.frozenColumnCount !== next.frozenColumnCount || previous.defaultColumn !== next.defaultColumn ||
    previous.defaultRow !== next.defaultRow || previous.columnMetrics !== next.columnMetrics ||
    previous.rowMetrics !== next.rowMetrics || previous.mergeRanges !== next.mergeRanges ||
    previous.selectionRef !== next.selectionRef || previous.viewportRef !== next.viewportRef ||
    previous.fillDraggingRef !== next.fillDraggingRef || previous.displayValue !== next.displayValue ||
    previous.resolvedValue !== next.resolvedValue || previous.staleValue !== next.staleValue ||
    previous.onSelection !== next.onSelection || previous.onBeginEdit !== next.onBeginEdit ||
    previous.onCellValue !== next.onCellValue || previous.onOpenHyperlink !== next.onOpenHyperlink ||
    previous.onPointCell !== next.onPointCell || previous.pointDraggingRef !== next.pointDraggingRef ||
    previous.listOptionsFor !== next.listOptionsFor ||
    previous.sparklineFor !== next.sparklineFor ||
    previous.coordinates.length !== next.coordinates.length
  ) return false
  return previous.coordinates === next.coordinates || previous.coordinates.every((coordinate, index) => (
    coordinate.row === next.coordinates[index].row && coordinate.col === next.coordinates[index].col
  ))
}

interface DimensionHeaderHandlers {
  select: (axis: 'column' | 'row', index: number, event: MouseEvent<HTMLDivElement>) => void
  extend: (axis: 'column' | 'row', index: number, event: MouseEvent<HTMLDivElement>) => void
  beginResize: (axis: 'column' | 'row', index: number, event: ReactPointerEvent<HTMLButtonElement>) => void
  autoFit: (axis: 'column' | 'row', index: number) => void
  resizeWithKey: (axis: 'column' | 'row', index: number, event: KeyboardEvent<HTMLButtonElement>) => void
}

interface DimensionHeaderProps {
  axis: 'column' | 'row'
  index: number
  offset: number
  size: number
  zoom: number
  frozen: boolean
  selection: 'none' | 'partial' | 'full'
  handlersRef: MutableRefObject<DimensionHeaderHandlers>
}

const DimensionHeader = memo(function DimensionHeader({ axis, index, offset, size, zoom, frozen, selection, handlersRef }: DimensionHeaderProps) {
  const row = axis === 'row'
  const label = row ? String(index + 1) : columnName(index)
  const noun = row ? 'row' : 'column'
  return (
    <div
      className={`${row ? 'row-header' : 'column-header'}${frozen ? ' is-frozen' : ''}${selection === 'full' ? ' is-fully-selected' : selection === 'partial' ? ' is-selected' : ''}`}
      style={row ? { left: 0, top: offset, width: HEADER_WIDTH, height: size } : { left: offset, top: 0, width: size, height: HEADER_HEIGHT }}
      role={row ? 'rowheader' : 'columnheader'}
      aria-label={`${row ? 'Row' : 'Column'} ${label}`}
      aria-selected={selection !== 'none' || undefined}
      onMouseDown={(event) => handlersRef.current.select(axis, index, event)}
      onMouseEnter={(event) => handlersRef.current.extend(axis, index, event)}
    >
      <span className="dimension-header-label">{label}</span>
      <button
        type="button"
        role="separator"
        aria-orientation={row ? 'horizontal' : 'vertical'}
        aria-label={`Resize ${noun} ${label}`}
        aria-valuemin={row ? MIN_ROW_PIXELS : MIN_COLUMN_PIXELS}
        aria-valuemax={row ? Math.round(MAX_ROW_POINTS * (4 / 3)) : MAX_COLUMN_PIXELS}
        aria-valuenow={Math.round(size / zoom)}
        title={`Drag to resize ${noun} ${label}. Double-click to auto-fit.`}
        className={`dimension-resize-handle is-${noun}`}
        {...(row ? { 'data-resize-row': index + 1 } : { 'data-resize-column': label })}
        onMouseDown={(event) => { event.preventDefault(); event.stopPropagation() }}
        onClick={(event) => event.stopPropagation()}
        onPointerDown={(event) => handlersRef.current.beginResize(axis, index, event)}
        onDoubleClick={(event) => { event.preventDefault(); event.stopPropagation(); handlersRef.current.autoFit(axis, index) }}
        onKeyDown={(event) => handlersRef.current.resizeWithKey(axis, index, event)}
      />
    </div>
  )
})

function SpreadsheetGrid({
  sheet,
  selection,
  editing,
  zoom,
  useImportedDefaults,
  showGridlines,
  showNotes,
  displayValue,
  resolvedValue,
  staleValue,
  onSelection,
  onBeginEdit,
  onDraft,
  onCommitEdit,
  onCancelEdit,
  onFill,
  onCellValue,
  onOpenHyperlink,
  onColumnResize,
  onRowResize,
  onFreeze,
  onKeyDown,
  onContextTarget,
  referenceHighlights,
  spillOutline,
  editorParts,
  editorSelectionRequest,
  onPointCell,
  onEditorKeyDown,
  onEditorCaret,
  onEditorFocusChange,
  conditionalFormats,
  filterHeaders,
  onFilterButton,
  inputMessage,
  invalidCells,
  cutRange,
  renderCharts,
  listOptionsFor,
  sparklineFor,
  outline,
  onOutlineToggle,
  onOutlineLevel,
}: GridProps) {
  const viewportRef = useRef<HTMLDivElement | null>(null)
  const rowOutlineContentRef = useRef<HTMLDivElement>(null)
  const columnOutlineContentRef = useRef<HTMLDivElement>(null)
  const pointDraggingRef = useRef(false)
  const headerDragRef = useRef<'column' | 'row' | null>(null)
  const editorHighlightRef = useRef<HTMLDivElement>(null)
  const selectionRef = useRef(selection)
  selectionRef.current = selection
  const fillDraggingRef = useRef(false)
  const fillTargetRef = useRef<Coord | null>(null)
  const fillPointerIdRef = useRef<number | null>(null)
  const fillPointerRef = useRef<{ x: number; y: number } | null>(null)
  const fillFrameRef = useRef<number | null>(null)
  const updateFillFrameRef = useRef<() => void>(() => {})
  const resizeRef = useRef<DimensionResize | null>(null)
  const freezeDragRef = useRef<FreezeDrag | null>(null)
  const viewportFrameRef = useRef<number | null>(null)
  const pendingViewportRef = useRef<{ scrollLeft: number; scrollTop: number } | null>(null)
  const fillPreviewOverlayRef = useRef<HTMLDivElement>(null)
  const editorOverlayRef = useRef<HTMLTextAreaElement>(null)
  // Off-DOM mirror for draft width/height. Measuring on the live textarea (width → 0)
  // makes Chromium chase the caret and shove .sheet-viewport to the far edge of the grid.
  const editorMeasureRef = useRef<HTMLTextAreaElement | null>(null)
  // Frozen while a cell edit is open so caret focus / selection-range cannot scroll the sheet.
  const editScrollLockRef = useRef<{ scrollLeft: number; scrollTop: number } | null>(null)
  const editorCommitViaKeyRef = useRef(false)
  const [fillTarget, setFillTarget] = useState<Coord | null>(null)
  const [resize, setResize] = useState<DimensionResize | null>(null)
  const [freezeDrag, setFreezeDrag] = useState<FreezeDrag | null>(null)
  const [viewport, setViewport] = useState({ scrollLeft: 0, scrollTop: 0, width: 1000, height: 600 })
  const sourceColumnWidth = Number(sheet.properties?.defaultColWidth)
  const sourceRowHeight = Number(sheet.properties?.defaultRowHeight)
  const defaultColumn = Number.isFinite(sourceColumnWidth) && sourceColumnWidth > 0
    ? columnPixelWidth(sourceColumnWidth, zoom) : (useImportedDefaults ? IMPORTED_COL_WIDTH : DEFAULT_COL_WIDTH) * zoom
  const defaultRow = Number.isFinite(sourceRowHeight) && sourceRowHeight > 0
    ? rowPixelHeight(sourceRowHeight, zoom) : (useImportedDefaults ? IMPORTED_ROW_HEIGHT : DEFAULT_ROW_HEIGHT) * zoom
  const syncPinnedOverlays = useCallback((scrollLeft: number, scrollTop: number) => {
    // Outline bars sit outside the scroller and follow it directly.
    if (rowOutlineContentRef.current) rowOutlineContentRef.current.style.transform = `translate3d(0, ${-scrollTop}px, 0)`
    if (columnOutlineContentRef.current) columnOutlineContentRef.current.style.transform = `translate3d(${-scrollLeft}px, 0, 0)`
    for (const element of [fillPreviewOverlayRef.current, editorOverlayRef.current, editorHighlightRef.current]) {
      if (!element) continue
      const x = element.classList.contains('pin-x') ? scrollLeft : 0
      const y = element.classList.contains('pin-y') ? scrollTop : 0
      const transform = `translate3d(${x}px, ${y}px, 0)`
      if (element.style.transform !== transform) element.style.transform = transform
    }
  }, [])
  useLayoutEffect(() => {
    const element = viewportRef.current
    if (!element || (!rowOutlineContentRef.current && !columnOutlineContentRef.current)) return
    if (rowOutlineContentRef.current) rowOutlineContentRef.current.style.transform = `translate3d(0, ${-element.scrollTop}px, 0)`
    if (columnOutlineContentRef.current) columnOutlineContentRef.current.style.transform = `translate3d(${-element.scrollLeft}px, 0, 0)`
  })
  const hiddenColumns = useMemo(() => new Set((sheet.hiddenCols || []).map((value) => Number(value) - 1).filter((value) => Number.isInteger(value) && value >= 0)), [sheet.hiddenCols])
  const hiddenRows = useMemo(() => new Set((sheet.hiddenRows || []).map((value) => Number(value) - 1).filter((value) => Number.isInteger(value) && value >= 0)), [sheet.hiddenRows])
  const columnMetrics = useMemo(() => {
    const indices = new Set<number>([
      ...Object.keys(sheet.colWidths || {}).map((key) => Number(key) - 1),
      ...hiddenColumns,
      ...(resize?.axis === 'column' ? [resize.index] : []),
    ])
    return [...indices]
      .filter((index) => Number.isInteger(index) && index >= 0)
      .map((index) => ({ index, size: hiddenColumns.has(index) ? 0 : resize?.axis === 'column' && resize.index === index ? resize.pixels : columnPixelWidth(Number(sheet.colWidths?.[String(index + 1)]), zoom) }))
      .sort((a, b) => a.index - b.index)
  }, [hiddenColumns, resize, sheet.colWidths, zoom])
  const rowMetrics = useMemo(() => {
    const indices = new Set<number>([
      ...Object.keys(sheet.rowHeights || {}).map((key) => Number(key) - 1),
      ...hiddenRows,
      ...(resize?.axis === 'row' ? [resize.index] : []),
    ])
    return [...indices]
      .filter((index) => Number.isInteger(index) && index >= 0)
      .map((index) => ({ index, size: hiddenRows.has(index) ? 0 : resize?.axis === 'row' && resize.index === index ? resize.pixels : rowPixelHeight(Number(sheet.rowHeights?.[String(index + 1)]), zoom) }))
      .sort((a, b) => a.index - b.index)
  }, [hiddenRows, resize, sheet.rowHeights, zoom])
  const rows = clamp(Math.max(DEFAULT_ROWS, sheet.rowCount + 25), DEFAULT_ROWS, 1_048_576)
  const columns = clamp(Math.max(DEFAULT_COLS, sheet.colCount + 8), DEFAULT_COLS, 16_384)
  const mergeRanges = useMemo(() => (sheet.merges || []).map(mergeBounds).filter((value): value is NonNullable<ReturnType<typeof mergeBounds>> => Boolean(value)), [sheet.merges])
  const frozenRowCount = clamp(Math.trunc(Number(sheet.frozen?.rows) || 0), 0, rows)
  const frozenColumnCount = clamp(Math.trunc(Number(sheet.frozen?.columns) || 0), 0, columns)
  const frozenHeight = axisOffset(frozenRowCount, defaultRow, rowMetrics)
  const frozenWidth = axisOffset(frozenColumnCount, defaultColumn, columnMetrics)
  const totalWidth = HEADER_WIDTH + axisOffset(columns, defaultColumn, columnMetrics)
  const totalHeight = HEADER_HEIGHT + axisOffset(rows, defaultRow, rowMetrics)
  const columnOverscan = Math.max(defaultColumn * 1.5, viewport.width * 0.08)
  const rowOverscan = Math.max(defaultRow * 8, viewport.height * 0.26)
  const startCol = axisIndexAt(Math.max(0, viewport.scrollLeft - HEADER_WIDTH - columnOverscan), columns, defaultColumn, columnMetrics)
  const startRow = axisIndexAt(Math.max(0, viewport.scrollTop - HEADER_HEIGHT - rowOverscan), rows, defaultRow, rowMetrics)
  const endCol = axisIndexAt(viewport.scrollLeft + viewport.width - HEADER_WIDTH + columnOverscan, columns, defaultColumn, columnMetrics)
  const endRow = axisIndexAt(viewport.scrollTop + viewport.height - HEADER_HEIGHT + rowOverscan, rows, defaultRow, rowMetrics)

  const [viewportNode, setViewportNode] = useState<HTMLDivElement | null>(null)
  useEffect(() => {
    const viewportElement = viewportNode
    if (!viewportElement) return
    const update = () => {
      // A replaced (detached) viewport reports 0x0; never let that shrink the rendered window.
      if (!viewportElement.isConnected) return
      syncPinnedOverlays(viewportElement.scrollLeft, viewportElement.scrollTop)
      setViewport((current) => {
        const next = {
          width: viewportElement.clientWidth,
          height: viewportElement.clientHeight,
          scrollLeft: viewportElement.scrollLeft,
          scrollTop: viewportElement.scrollTop,
        }
        return current.width === next.width && current.height === next.height && current.scrollLeft === next.scrollLeft && current.scrollTop === next.scrollTop ? current : next
      })
    }
    update()
    const observer = new ResizeObserver(update)
    observer.observe(viewportElement)
    return () => {
      observer.disconnect()
      if (viewportFrameRef.current !== null) window.cancelAnimationFrame(viewportFrameRef.current)
      viewportFrameRef.current = null
      pendingViewportRef.current = null
    }
  }, [syncPinnedOverlays, viewportNode])

  useEffect(() => {
    const moveResize = (event: globalThis.PointerEvent) => {
      const current = resizeRef.current
      if (!current || current.pointerId !== event.pointerId) return
      event.preventDefault()
      const delta = (current.axis === 'column' ? event.clientX : event.clientY) - current.startClient
      const minimum = (current.axis === 'column' ? MIN_COLUMN_PIXELS : MIN_ROW_PIXELS) * zoom
      const maximum = (current.axis === 'column' ? MAX_COLUMN_PIXELS : MAX_ROW_POINTS * (4 / 3)) * zoom
      const next = { ...current, pixels: clamp(current.startPixels + delta, minimum, maximum) }
      resizeRef.current = next
      setResize(next)
    }
    const finishResize = (commit: boolean, pointerId?: number) => {
      const current = resizeRef.current
      if (!current || (pointerId !== undefined && current.pointerId !== pointerId)) return
      resizeRef.current = null
      setResize(null)
      if (!commit || Math.abs(current.pixels - current.startPixels) < 0.5) return
      if (current.axis === 'column') onColumnResize(current.index, modelColumnWidth(current.pixels, zoom))
      else onRowResize(current.index, modelRowHeight(current.pixels, zoom))
      viewportRef.current?.focus({ preventScroll: true })
    }
    const endResize = (event: globalThis.PointerEvent) => finishResize(true, event.pointerId)
    const cancelPointerResize = (event: globalThis.PointerEvent) => finishResize(false, event.pointerId)
    const cancelResize = () => finishResize(false)
    const cancelWithEscape = (event: globalThis.KeyboardEvent) => {
      if (event.key !== 'Escape' || !resizeRef.current) return
      event.preventDefault()
      finishResize(false)
    }
    window.addEventListener('pointermove', moveResize)
    window.addEventListener('pointerup', endResize)
    window.addEventListener('pointercancel', cancelPointerResize)
    window.addEventListener('blur', cancelResize)
    window.addEventListener('keydown', cancelWithEscape)
    return () => {
      window.removeEventListener('pointermove', moveResize)
      window.removeEventListener('pointerup', endResize)
      window.removeEventListener('pointercancel', cancelPointerResize)
      window.removeEventListener('blur', cancelResize)
      window.removeEventListener('keydown', cancelWithEscape)
    }
  }, [onColumnResize, onRowResize, zoom])

  useEffect(() => {
    const release = () => { pointDraggingRef.current = false; headerDragRef.current = null }
    window.addEventListener('pointerup', release)
    window.addEventListener('blur', release)
    return () => {
      window.removeEventListener('pointerup', release)
      window.removeEventListener('blur', release)
    }
  }, [])

  useEffect(() => {
    if (!resize) return
    const previous = document.documentElement.style.cursor
    document.documentElement.style.cursor = resize.axis === 'column' ? 'col-resize' : 'row-resize'
    return () => { document.documentElement.style.cursor = previous }
  }, [resize?.axis])

  const safeFreezeCount = useCallback((axis: 'row' | 'column', requested: number) => {
    const count = axis === 'row' ? rows : columns
    const defaultSize = axis === 'row' ? defaultRow : defaultColumn
    const metrics = axis === 'row' ? rowMetrics : columnMetrics
    const viewportSize = axis === 'row' ? viewport.height - HEADER_HEIGHT : viewport.width - HEADER_WIDTH
    const minimumScrollable = Math.max(defaultSize, (axis === 'row' ? 42 : 64) * zoom)
    const maximumPixels = Math.max(0, viewportSize - minimumScrollable)
    const maximumCount = axisIndexAt(maximumPixels, count, defaultSize, metrics)
    const current = axis === 'row' ? frozenRowCount : frozenColumnCount
    let next = clamp(Math.trunc(requested), 0, maximumCount)

    // A pane boundary cannot split a merged cell. When freezing more, include
    // the whole merge; when dragging back, snap before it.
    for (let attempt = 0; attempt <= mergeRanges.length; attempt += 1) {
      const split = mergeRanges.find((range) => axis === 'row'
        ? range.top < next && next <= range.bottom
        : range.left < next && next <= range.right)
      if (!split) break
      const outward = axis === 'row' ? split.bottom + 1 : split.right + 1
      const inward = axis === 'row' ? split.top : split.left
      next = clamp(next >= current && outward <= maximumCount ? outward : inward, 0, maximumCount)
    }
    return next
  }, [columnMetrics, columns, defaultColumn, defaultRow, frozenColumnCount, frozenRowCount, mergeRanges, rowMetrics, rows, viewport.height, viewport.width, zoom])

  const freezeCountAtClient = useCallback((axis: 'row' | 'column', clientPosition: number) => {
    const element = viewportRef.current
    if (!element) return axis === 'row' ? frozenRowCount : frozenColumnCount
    const rect = element.getBoundingClientRect()
    const header = axis === 'row' ? HEADER_HEIGHT : HEADER_WIDTH
    const local = clientPosition - (axis === 'row' ? rect.top : rect.left) - header
    const frozenPixels = axis === 'row' ? frozenHeight : frozenWidth
    const scroll = axis === 'row' ? element.scrollTop : element.scrollLeft
    const sheetPixel = Math.max(0, local <= frozenPixels ? local : local + scroll)
    return safeFreezeCount(axis, axisBoundaryAt(
      sheetPixel,
      axis === 'row' ? rows : columns,
      axis === 'row' ? defaultRow : defaultColumn,
      axis === 'row' ? rowMetrics : columnMetrics,
    ))
  }, [columnMetrics, columns, defaultColumn, defaultRow, frozenColumnCount, frozenHeight, frozenRowCount, frozenWidth, rowMetrics, rows, safeFreezeCount])

  useEffect(() => {
    const finishFreeze = (commit: boolean, pointerId?: number) => {
      const current = freezeDragRef.current
      if (!current || (pointerId !== undefined && current.pointerId !== pointerId)) return
      freezeDragRef.current = null
      setFreezeDrag(null)
      const previous = current.axis === 'row' ? frozenRowCount : frozenColumnCount
      if (commit && current.count !== previous) onFreeze(current.axis === 'row' ? 'rows' : 'columns', current.count)
      viewportRef.current?.focus({ preventScroll: true })
    }
    const moveFreeze = (event: globalThis.PointerEvent) => {
      const current = freezeDragRef.current
      if (!current || current.pointerId !== event.pointerId) return
      event.preventDefault()
      const count = freezeCountAtClient(current.axis, current.axis === 'row' ? event.clientY : event.clientX)
      if (count === current.count) return
      const next = { ...current, count }
      freezeDragRef.current = next
      setFreezeDrag(next)
    }
    const endFreeze = (event: globalThis.PointerEvent) => finishFreeze(true, event.pointerId)
    const cancelFreeze = (event: globalThis.PointerEvent) => finishFreeze(false, event.pointerId)
    const cancelOnBlur = () => finishFreeze(false)
    const cancelWithEscape = (event: globalThis.KeyboardEvent) => {
      if (event.key !== 'Escape' || !freezeDragRef.current) return
      event.preventDefault()
      finishFreeze(false)
    }
    window.addEventListener('pointermove', moveFreeze)
    window.addEventListener('pointerup', endFreeze)
    window.addEventListener('pointercancel', cancelFreeze)
    window.addEventListener('blur', cancelOnBlur)
    window.addEventListener('keydown', cancelWithEscape)
    return () => {
      window.removeEventListener('pointermove', moveFreeze)
      window.removeEventListener('pointerup', endFreeze)
      window.removeEventListener('pointercancel', cancelFreeze)
      window.removeEventListener('blur', cancelOnBlur)
      window.removeEventListener('keydown', cancelWithEscape)
    }
  }, [freezeCountAtClient, frozenColumnCount, frozenRowCount, onFreeze])

  useEffect(() => {
    if (!freezeDrag) return
    const previous = document.documentElement.style.cursor
    document.documentElement.style.cursor = freezeDrag.axis === 'row' ? 'row-resize' : 'col-resize'
    return () => { document.documentElement.style.cursor = previous }
  }, [freezeDrag?.axis])

  useEffect(() => {
    const element = viewportRef.current
    if (!element) return
    const topLeft = coordOf(sheet.frozen?.topLeftCell || '')
    element.scrollLeft = topLeft ? Math.max(0, axisOffset(topLeft.col, defaultColumn, columnMetrics) - frozenWidth) : 0
    element.scrollTop = topLeft ? Math.max(0, axisOffset(topLeft.row, defaultRow, rowMetrics) - frozenHeight) : 0
    syncPinnedOverlays(element.scrollLeft, element.scrollTop)
  // A newly activated sheet owns its saved view. Zooming or editing must not
  // unexpectedly snap that view back to the file's original top-left cell.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet.id])

  useLayoutEffect(() => {
    const element = viewportRef.current
    if (!editing) {
      editScrollLockRef.current = null
      return
    }
    if (!element) return
    // Capture once when the edit opens so later caret/focus work cannot move the sheet.
    if (!editScrollLockRef.current) {
      editScrollLockRef.current = { scrollLeft: element.scrollLeft, scrollTop: element.scrollTop }
    }
    const lock = editScrollLockRef.current
    const pin = () => {
      if (!editScrollLockRef.current) return
      if (element.scrollLeft === lock.scrollLeft && element.scrollTop === lock.scrollTop) return
      element.scrollLeft = lock.scrollLeft
      element.scrollTop = lock.scrollTop
      syncPinnedOverlays(lock.scrollLeft, lock.scrollTop)
    }
    pin()
    // setSelectionRange / caret scroll can land after the focus handler returns.
    // Keep pinning for the life of the edit so Chromium cannot yank the virtual grid.
    let frame = window.requestAnimationFrame(function tick() {
      pin()
      frame = window.requestAnimationFrame(tick)
    })
    return () => window.cancelAnimationFrame(frame)
  }, [editing, syncPinnedOverlays])

  useEffect(() => {
    // While a cell is being edited, Chromium will also try to keep the caret in view.
    // That caret can sit far past the cell box on a long unwrapped draft, so following
    // it (or the selection) would shove the sheet to the far edge. Stay put until edit ends.
    if (editing) return
    const element = viewportRef.current
    if (!element) return
    const x = HEADER_WIDTH + axisOffset(selection.focus.col, defaultColumn, columnMetrics)
    const y = HEADER_HEIGHT + axisOffset(selection.focus.row, defaultRow, rowMetrics)
    const width = axisSize(selection.focus.col, defaultColumn, columnMetrics)
    const height = axisSize(selection.focus.row, defaultRow, rowMetrics)
    if (selection.focus.col >= frozenColumnCount) {
      if (x < element.scrollLeft + HEADER_WIDTH + frozenWidth) element.scrollLeft = Math.max(0, x - HEADER_WIDTH - frozenWidth)
      else if (x + width > element.scrollLeft + element.clientWidth) element.scrollLeft = x + width - element.clientWidth
    }
    if (selection.focus.row >= frozenRowCount) {
      if (y < element.scrollTop + HEADER_HEIGHT + frozenHeight) element.scrollTop = Math.max(0, y - HEADER_HEIGHT - frozenHeight)
      else if (y + height > element.scrollTop + element.clientHeight) element.scrollTop = y + height - element.clientHeight
    }
  }, [editing, selection.focus.col, selection.focus.row, defaultColumn, defaultRow, columnMetrics, rowMetrics, frozenColumnCount, frozenRowCount, frozenHeight, frozenWidth])

  const frozenRenderRowEnd = frozenRowCount
    ? Math.min(frozenRowCount - 1, axisIndexAt(Math.max(0, viewport.height - HEADER_HEIGHT), rows, defaultRow, rowMetrics) + 2)
    : -1
  const frozenRenderColumnEnd = frozenColumnCount
    ? Math.min(frozenColumnCount - 1, axisIndexAt(Math.max(0, viewport.width - HEADER_WIDTH), columns, defaultColumn, columnMetrics) + 2)
    : -1
  const renderedRows = useMemo(() => [...new Set([
    ...Array.from({ length: Math.max(0, endRow - startRow + 1) }, (_, offset) => startRow + offset),
    ...Array.from({ length: frozenRenderRowEnd + 1 }, (_, row) => row),
  ])].filter((row) => !hiddenRows.has(row)).sort((a, b) => a - b), [endRow, frozenRenderRowEnd, hiddenRows, startRow])
  const renderedColumns = useMemo(() => [...new Set([
    ...Array.from({ length: Math.max(0, endCol - startCol + 1) }, (_, offset) => startCol + offset),
    ...Array.from({ length: frozenRenderColumnEnd + 1 }, (_, col) => col),
  ])].filter((col) => !hiddenColumns.has(col)).sort((a, b) => a - b), [endCol, frozenRenderColumnEnd, hiddenColumns, startCol])

  // The rendered window grouped into per-pane rows. Merges resolve row by row: covered cells
  // drop out and each merge is drawn from its top-left master in every pane it crosses. Row
  // lists are reused while a row's columns are unchanged, so scrolling only renders new rows.
  const rowCoordinateCacheRef = useRef(new Map<string, Coord[]>())
  const { rowsByPane, mergeMap } = useMemo(() => {
    const sortedHas = (values: number[], low: number, high: number) => {
      let lo = 0, hi = values.length
      while (lo < hi) { const mid = (lo + hi) >> 1; if (values[mid] < low) lo = mid + 1; else hi = mid }
      return lo < values.length && values[lo] <= high
    }
    const active = mergeRanges.filter((merge) => sortedHas(renderedRows, merge.top, merge.bottom) && sortedHas(renderedColumns, merge.left, merge.right))
    const merges = active.length ? new Map<string, ReturnType<typeof mergeBounds>>() : NO_MERGES
    const groups: Record<CellPane, Map<number, Coord[]>> = { body: new Map(), 'frozen-row': new Map(), 'frozen-column': new Map(), 'frozen-corner': new Map() }
    const push = (pane: CellPane, coordinate: Coord) => {
      const row = groups[pane].get(coordinate.row)
      if (row) row.push(coordinate)
      else groups[pane].set(coordinate.row, [coordinate])
    }
    for (const row of renderedRows) {
      const frozenRow = row < frozenRowCount
      const rowMerges = active.length ? active.filter((merge) => row >= merge.top && row <= merge.bottom) : active
      for (const col of renderedColumns) {
        const merge = rowMerges.length ? rowMerges.find((item) => col >= item.left && col <= item.right) : undefined
        if (merge) (merges as Map<string, ReturnType<typeof mergeBounds>>).set(addressOf({ row, col }), merge)
        else push(cellPane(frozenRow, col < frozenColumnCount), { row, col })
      }
    }
    for (const merge of active) {
      const master = { row: merge.top, col: merge.left }
      ;(merges as Map<string, ReturnType<typeof mergeBounds>>).set(addressOf(master), merge)
      for (const pane of intersectingPanes(merge, frozenRowCount, frozenColumnCount)) push(pane, master)
    }
    const previousRows = rowCoordinateCacheRef.current
    const nextRows = new Map<string, Coord[]>()
    const panes = {} as Record<CellPane, Coord[][]>
    for (const pane of Object.keys(groups) as CellPane[]) {
      const entries = [...groups[pane].entries()]
      if (active.length) {
        entries.sort((a, b) => a[0] - b[0])
        for (const [, coordinates] of entries) coordinates.sort((a, b) => a.col - b.col)
      }
      panes[pane] = entries.map(([row, coordinates]) => {
        const key = `${pane}:${row}`
        const cached = previousRows.get(key)
        const reuse = cached && cached.length === coordinates.length && cached.every((item, index) => item.col === coordinates[index].col)
        const stable = reuse ? cached : coordinates
        nextRows.set(key, stable)
        return stable
      })
    }
    rowCoordinateCacheRef.current = nextRows
    return { rowsByPane: panes, mergeMap: merges }
  }, [frozenColumnCount, frozenRowCount, mergeRanges, renderedColumns, renderedRows])

  const measurementContext = useMemo(() => document.createElement('canvas').getContext('2d'), [])
  const measureText = (text: string, cell?: CellData) => {
    const font = cell?.style?.font || {}
    const fontSize = Math.max(8, Number(font.size) || 11) * (4 / 3) * zoom
    if (measurementContext) {
      const family = fontCss(font).fontFamily || 'Calibri, Aptos, "Segoe UI", Arial, sans-serif'
      measurementContext.font = `${font.italic ? 'italic ' : ''}${font.bold ? '700 ' : '400 '}${fontSize}px ${family}`
    }
    const lines = String(text || '').split(/\r?\n/)
    const width = Math.max(0, ...lines.map((line) => measurementContext?.measureText(line).width || line.length * fontSize * 0.55))
    return { width, fontSize, lineCount: Math.max(1, lines.length) }
  }

  const autoFitColumn = (col: number) => {
    if (hiddenColumns.has(col)) return
    const header = measureText(columnName(col))
    let targetPixels = Math.max(40 * zoom, header.width + 18 * zoom)
    let inspected = 0
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const coord = coordOf(address)
      if (!coord || coord.col !== col) continue
      const merged = mergeRanges.find((range) => coord.row >= range.top && coord.row <= range.bottom && coord.col >= range.left && coord.col <= range.right)
      if (merged && merged.left !== merged.right) continue
      const text = displayValue(sheet.id, address)
      if (!text) continue
      const measured = measureText(text, cell)
      const alignment = cell.style?.alignment || {}
      const indent = Math.max(0, Number(alignment.indent || 0) + Number(alignment.relativeIndent || 0)) * 9 * zoom
      let required = measured.width + indent + 14 * zoom
      const rotation = typeof alignment.textRotation === 'number' ? alignment.textRotation : alignment.textRotation === 'vertical' ? 90 : 0
      if (rotation) {
        const radians = Math.abs(rotation) * Math.PI / 180
        required = Math.abs(measured.width * Math.cos(radians)) + Math.abs(measured.fontSize * 1.25 * measured.lineCount * Math.sin(radians)) + 14 * zoom
      }
      targetPixels = Math.max(targetPixels, required)
      inspected += 1
      if (inspected >= 50_000 || targetPixels >= MAX_COLUMN_PIXELS * zoom) break
    }
    onColumnResize(col, modelColumnWidth(clamp(targetPixels, MIN_COLUMN_PIXELS * zoom, MAX_COLUMN_PIXELS * zoom), zoom))
  }

  const autoFitRow = (row: number) => {
    if (hiddenRows.has(row)) return
    let targetPixels = defaultRow
    let inspected = 0
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const coord = coordOf(address)
      if (!coord || coord.row !== row) continue
      const merged = mergeRanges.find((range) => coord.row >= range.top && coord.row <= range.bottom && coord.col >= range.left && coord.col <= range.right)
      if (merged && merged.top !== merged.bottom) continue
      const text = displayValue(sheet.id, address)
      if (!text) continue
      const measured = measureText(text, cell)
      const left = merged ? merged.left : coord.col
      const right = merged ? merged.right : coord.col
      const available = Math.max(12 * zoom, axisOffset(right + 1, defaultColumn, columnMetrics) - axisOffset(left, defaultColumn, columnMetrics) - 12 * zoom)
      const shouldWrap = Boolean(cell.style?.alignment?.wrapText)
      const logicalLines = String(text).split(/\r?\n/).reduce((count, line) => (
        count + (shouldWrap ? Math.max(1, Math.ceil((measurementContext?.measureText(line).width || line.length * measured.fontSize * 0.55) / available)) : 1)
      ), 0)
      const rotation = typeof cell.style?.alignment?.textRotation === 'number' ? cell.style.alignment.textRotation : cell.style?.alignment?.textRotation === 'vertical' ? 90 : 0
      const required = rotation
        ? Math.abs(measured.width * Math.sin(Math.abs(rotation) * Math.PI / 180)) + measured.fontSize * 1.25 + 8 * zoom
        : Math.max(1, logicalLines) * measured.fontSize * 1.25 + 6 * zoom
      targetPixels = Math.max(targetPixels, required)
      inspected += 1
      if (inspected >= 50_000 || targetPixels >= MAX_ROW_POINTS * (4 / 3) * zoom) break
    }
    onRowResize(row, modelRowHeight(clamp(targetPixels, MIN_ROW_PIXELS * zoom, MAX_ROW_POINTS * (4 / 3) * zoom), zoom))
  }

  const beginDimensionResize = (axis: 'column' | 'row', index: number, event: ReactPointerEvent<HTMLButtonElement>) => {
    if (event.button !== 0 || (axis === 'column' ? hiddenColumns.has(index) : hiddenRows.has(index))) return
    event.preventDefault()
    event.stopPropagation()
    const startPixels = axis === 'column'
      ? axisSize(index, defaultColumn, columnMetrics)
      : axisSize(index, defaultRow, rowMetrics)
    const next: DimensionResize = {
      axis,
      index,
      pointerId: event.pointerId,
      startClient: axis === 'column' ? event.clientX : event.clientY,
      startPixels,
      pixels: startPixels,
    }
    resizeRef.current = next
    setResize(next)
    try { event.currentTarget.setPointerCapture(event.pointerId) } catch { /* Window listeners still finish the resize. */ }
  }

  const resizeDimensionWithKey = (axis: 'column' | 'row', index: number, event: KeyboardEvent<HTMLButtonElement>) => {
    const positive = axis === 'column' ? event.key === 'ArrowRight' : event.key === 'ArrowDown'
    const negative = axis === 'column' ? event.key === 'ArrowLeft' : event.key === 'ArrowUp'
    if (event.key === 'Enter') {
      event.preventDefault()
      if (axis === 'column') autoFitColumn(index)
      else autoFitRow(index)
      return
    }
    if (!positive && !negative) return
    event.preventDefault()
    const current = axis === 'column'
      ? axisSize(index, defaultColumn, columnMetrics)
      : axisSize(index, defaultRow, rowMetrics)
    const step = (event.shiftKey ? 20 : axis === 'column' ? 7 : 4) * zoom
    const pixels = current + (positive ? step : -step)
    if (axis === 'column') onColumnResize(index, modelColumnWidth(clamp(pixels, MIN_COLUMN_PIXELS * zoom, MAX_COLUMN_PIXELS * zoom), zoom))
    else onRowResize(index, modelRowHeight(clamp(pixels, MIN_ROW_PIXELS * zoom, MAX_ROW_POINTS * (4 / 3) * zoom), zoom))
  }

  const beginFreezeDrag = (axis: 'column' | 'row', event: ReactPointerEvent<HTMLButtonElement>) => {
    if (event.button !== 0) return
    event.preventDefault()
    event.stopPropagation()
    const next: FreezeDrag = {
      axis,
      pointerId: event.pointerId,
      count: axis === 'row' ? frozenRowCount : frozenColumnCount,
    }
    freezeDragRef.current = next
    setFreezeDrag(next)
    try { event.currentTarget.setPointerCapture(event.pointerId) } catch { /* Window listeners still complete the drag. */ }
  }

  const freezeHandleKeyDown = (axis: 'column' | 'row', event: KeyboardEvent<HTMLButtonElement>) => {
    const current = axis === 'row' ? frozenRowCount : frozenColumnCount
    const decrement = axis === 'row' ? event.key === 'ArrowUp' : event.key === 'ArrowLeft'
    const increment = axis === 'row' ? event.key === 'ArrowDown' : event.key === 'ArrowRight'
    if (event.key !== 'Home' && !decrement && !increment) return
    event.preventDefault()
    const requested = event.key === 'Home' ? 0 : current + (increment ? (event.shiftKey ? 5 : 1) : -(event.shiftKey ? 5 : 1))
    const next = safeFreezeCount(axis, requested)
    if (next !== current) onFreeze(axis === 'row' ? 'rows' : 'columns', next)
  }

  const handleViewportScroll = (target: HTMLDivElement) => {
    const lock = editScrollLockRef.current
    if (lock && (target.scrollLeft !== lock.scrollLeft || target.scrollTop !== lock.scrollTop)) {
      // Cell edit focus/selection-range and Chromium caret scrolling try to yank the
      // viewport to the grid edge. Hold the position the edit started at until commit.
      target.scrollLeft = lock.scrollLeft
      target.scrollTop = lock.scrollTop
      syncPinnedOverlays(lock.scrollLeft, lock.scrollTop)
      return
    }
    const scrollLeft = target.scrollLeft
    const scrollTop = target.scrollTop
    // Keep the handful of pinned overlays on the compositor without changing
    // an inherited style on every grid cell.
    syncPinnedOverlays(scrollLeft, scrollTop)
    const visibleLeft = Math.max(0, scrollLeft - HEADER_WIDTH)
    const visibleTop = Math.max(0, scrollTop - HEADER_HEIGHT)
    const visibleRight = Math.max(0, scrollLeft + target.clientWidth - HEADER_WIDTH)
    const visibleBottom = Math.max(0, scrollTop + target.clientHeight - HEADER_HEIGHT)
    const columnGuard = columnOverscan * 0.42
    const rowGuard = rowOverscan * 0.42
    const needsColumns = (startCol > 0 && visibleLeft < axisOffset(startCol, defaultColumn, columnMetrics) + columnGuard) ||
      (endCol < columns - 1 && visibleRight > axisOffset(endCol + 1, defaultColumn, columnMetrics) - columnGuard)
    const needsRows = (startRow > 0 && visibleTop < axisOffset(startRow, defaultRow, rowMetrics) + rowGuard) ||
      (endRow < rows - 1 && visibleBottom > axisOffset(endRow + 1, defaultRow, rowMetrics) - rowGuard)
    if (!needsColumns && !needsRows) return
    pendingViewportRef.current = { scrollLeft, scrollTop }
    if (viewportFrameRef.current !== null) return
    viewportFrameRef.current = window.requestAnimationFrame(() => {
      viewportFrameRef.current = null
      const pending = pendingViewportRef.current
      pendingViewportRef.current = null
      if (!pending) return
      setViewport((current) => current.scrollLeft === pending.scrollLeft && current.scrollTop === pending.scrollTop
        ? current
        : { ...current, ...pending })
    })
  }

  const handleViewportContextMenu = (event: MouseEvent<HTMLDivElement>) => {
    const element = viewportRef.current
    if (!element) return
    if (isNativeTextEditingTarget(event.target)) return
    const rect = element.getBoundingClientRect()
    const localX = event.clientX - rect.left
    const localY = event.clientY - rect.top
    if (localX < HEADER_WIDTH && localY < HEADER_HEIGHT) {
      event.preventDefault()
      onContextTarget({ kind: 'corner', coord: { row: 0, col: 0 } }, { x: event.clientX, y: event.clientY })
      return
    }
    event.preventDefault()
    const columnLocal = localX - HEADER_WIDTH
    const rowLocal = localY - HEADER_HEIGHT
    const columnPixel = columnLocal < frozenWidth ? columnLocal : columnLocal + element.scrollLeft
    const rowPixel = rowLocal < frozenHeight ? rowLocal : rowLocal + element.scrollTop
    const coord = {
      row: axisIndexAt(Math.max(0, rowPixel), rows, defaultRow, rowMetrics),
      col: axisIndexAt(Math.max(0, columnPixel), columns, defaultColumn, columnMetrics),
    }
    const kind = localX < HEADER_WIDTH ? 'row-header' as const : localY < HEADER_HEIGHT ? 'column-header' as const : 'cell' as const
    onContextTarget({ kind, coord }, { x: event.clientX, y: event.clientY })
  }

  const selectedBounds = selectionBounds(selection)
  const selectedMerge = selectedBounds.left === selectedBounds.right && selectedBounds.top === selectedBounds.bottom
    ? mergeRanges.find((merge) => selectedBounds.top >= merge.top && selectedBounds.top <= merge.bottom && selectedBounds.left >= merge.left && selectedBounds.left <= merge.right)
    : undefined
  const bounds = selectedMerge || selectedBounds
  const selectionPanes = intersectingPanes(bounds, frozenRowCount, frozenColumnCount)
  const selectionHandlePane = cellPane(bounds.bottom < frozenRowCount, bounds.right < frozenColumnCount)
  const selectionLeft = axisOffset(bounds.left, defaultColumn, columnMetrics)
  const selectionTop = axisOffset(bounds.top, defaultRow, rowMetrics)
  const selectionWidth = axisOffset(bounds.right + 1, defaultColumn, columnMetrics) - axisOffset(bounds.left, defaultColumn, columnMetrics)
  const selectionHeight = axisOffset(bounds.bottom + 1, defaultRow, rowMetrics) - axisOffset(bounds.top, defaultRow, rowMetrics)
  const fillPreview = fillTarget ? fillSelectionForTarget(selection, fillTarget) : null
  const fillBounds = fillPreview ? selectionBounds(fillPreview) : null
  const fillPinnedX = Boolean(fillBounds && fillBounds.right < frozenColumnCount)
  const fillPinnedY = Boolean(fillBounds && fillBounds.bottom < frozenRowCount)
  const fillLeft = fillBounds ? HEADER_WIDTH + axisOffset(fillBounds.left, defaultColumn, columnMetrics) : 0
  const fillTop = fillBounds ? HEADER_HEIGHT + axisOffset(fillBounds.top, defaultRow, rowMetrics) : 0
  const fillWidth = fillBounds ? axisOffset(fillBounds.right + 1, defaultColumn, columnMetrics) - axisOffset(fillBounds.left, defaultColumn, columnMetrics) : 0
  const fillHeight = fillBounds ? axisOffset(fillBounds.bottom + 1, defaultRow, rowMetrics) - axisOffset(fillBounds.top, defaultRow, rowMetrics) : 0
  // The in-cell editor is anchored to the cell it edits: it starts at exactly the cell
  // rectangle, wears the cell's own font and alignment, and only grows once the typed
  // content no longer fits — the way Excel and Sheets behave.
  const editorGeometry = useMemo(() => {
    if (!editing) return null
    const coord = coordOf(editing.address)
    if (!coord) return null
    const merge = mergeMap.get(editing.address)
    // A merge is one cell: its master carries the content, the style and the rectangle,
    // even when the selection focus landed on a covered address.
    const anchor = merge ? { row: merge.top, col: merge.left } : coord
    const cell = sheet.cells[addressOf(anchor)]
    const alignment = cell?.style?.alignment || {}
    const horizontal = String(alignment.horizontal || '').toLocaleLowerCase()
    const numeric = typeof cell?.value === 'number' || (Boolean(cell?.formula) && typeof cell?.result === 'number')
    const textAlign: CSSProperties['textAlign'] = horizontal === 'right' ? 'right'
      : horizontal === 'center' || horizontal === 'centercontinuous' ? 'center'
        : horizontal === 'justify' || horizontal === 'distributed' ? 'justify'
          : horizontal === 'left' ? 'left' : numeric && !editing.draft.startsWith('=') ? 'right' : undefined
    return {
      frozenColumn: anchor.col < frozenColumnCount,
      frozenRow: anchor.row < frozenRowCount,
      left: HEADER_WIDTH + axisOffset(anchor.col, defaultColumn, columnMetrics),
      top: HEADER_HEIGHT + axisOffset(anchor.row, defaultRow, rowMetrics),
      width: merge
        ? axisOffset(merge.right + 1, defaultColumn, columnMetrics) - axisOffset(merge.left, defaultColumn, columnMetrics)
        : axisSize(coord.col, defaultColumn, columnMetrics),
      height: merge
        ? axisOffset(merge.bottom + 1, defaultRow, rowMetrics) - axisOffset(merge.top, defaultRow, rowMetrics)
        : axisSize(coord.row, defaultRow, rowMetrics),
      font: cell?.style?.font,
      // Wearing the cell's own solid fill keeps the editor opaque over its neighbours while
      // staying readable with the cell's own font colour.  Patterns and gradients fall back
      // to plain paper.
      background: String(cell?.style?.fill?.pattern || '').toLocaleLowerCase() === 'solid'
        ? cssColor(cell?.style?.fill?.color || cell?.style?.fill?.fgColor) || undefined
        : undefined,
      wrapText: Boolean(alignment.wrapText),
      textAlign,
      growLeft: textAlign === 'right',
    }
  }, [columnMetrics, defaultColumn, defaultRow, editing, frozenColumnCount, frozenRowCount, mergeMap, rowMetrics, sheet])

  // Runs after every render so a scroll or an unrelated state change cannot snap the
  // editor back to the cell rectangle while a longer draft is being typed.
  useLayoutEffect(() => {
    const node = editorOverlayRef.current
    if (!node || !editorGeometry) return
    const { left, top, width, height, wrapText, growLeft, frozenColumn, frozenRow } = editorGeometry
    const element = viewportRef.current
    const lock = editScrollLockRef.current
    const scrollLeft = lock?.scrollLeft ?? element?.scrollLeft ?? 0
    const scrollTop = lock?.scrollTop ?? element?.scrollTop ?? 0
    const clientWidth = element?.clientWidth || viewport.width
    const clientHeight = element?.clientHeight || viewport.height
    // A pinned editor also carries translate3d(scrollLeft, scrollTop) from pinnedTransform,
    // so its painted origin is left + scrollLeft, not left.
    const originX = left + (frozenColumn ? scrollLeft : 0)
    const originY = top + (frozenRow ? scrollTop : 0)
    const roomRight = Math.max(0, scrollLeft + clientWidth - originX - 4)
    const roomLeft = Math.max(0, originX - scrollLeft - HEADER_WIDTH)
    const maxWidth = Math.max(width, growLeft ? roomLeft + width : roomRight)
    const maxHeight = Math.max(height, scrollTop + clientHeight - originY - 4)

    // Measure on an off-DOM clone. Collapsing the live editor to width 0 lets Chromium
    // scroll .sheet-viewport so the caret stays in view — often the far edge of a large grid.
    let measure = editorMeasureRef.current
    if (!measure) {
      measure = document.createElement('textarea')
      measure.setAttribute('aria-hidden', 'true')
      measure.tabIndex = -1
      measure.rows = 1
      measure.spellcheck = false
      Object.assign(measure.style, {
        position: 'fixed',
        left: '0',
        top: '0',
        visibility: 'hidden',
        pointerEvents: 'none',
        zIndex: '-1',
        padding: '0 6px',
        border: '0',
        outline: '0',
        resize: 'none',
        overflow: 'hidden',
        boxSizing: 'border-box',
      })
      document.body.appendChild(measure)
      editorMeasureRef.current = measure
    }
    const computed = window.getComputedStyle(node)
    measure.style.font = computed.font
    measure.style.fontFamily = computed.fontFamily
    measure.style.fontSize = computed.fontSize
    measure.style.fontWeight = computed.fontWeight
    measure.style.fontStyle = computed.fontStyle
    measure.style.letterSpacing = computed.letterSpacing
    measure.style.lineHeight = computed.lineHeight
    measure.style.fontVariantNumeric = computed.fontVariantNumeric
    measure.value = node.value
    measure.style.whiteSpace = 'pre'
    measure.style.width = '0px'
    measure.style.height = `${height}px`
    const needed = measure.scrollWidth
    const nextWidth = wrapText ? width : Math.min(Math.max(width, needed), maxWidth)
    measure.style.width = `${nextWidth}px`
    measure.style.whiteSpace = wrapText || needed > nextWidth ? 'pre-wrap' : 'pre'
    measure.style.height = 'auto'
    const neededHeight = measure.scrollHeight
    const nextHeight = Math.min(Math.max(height, neededHeight), maxHeight)

    node.style.left = `${left}px`
    if (growLeft && nextWidth > width) node.style.left = `${Math.max(left - roomLeft, left - (nextWidth - width))}px`
    node.style.width = `${nextWidth}px`
    node.style.height = `${nextHeight}px`
    node.style.whiteSpace = wrapText || needed > nextWidth ? 'pre-wrap' : 'pre'
    node.style.overflowY = neededHeight > nextHeight + 1 ? 'auto' : 'hidden'
    const highlight = editorHighlightRef.current
    if (highlight) {
      highlight.style.left = node.style.left
      highlight.style.width = node.style.width
      highlight.style.height = node.style.height
      highlight.style.whiteSpace = node.style.whiteSpace
      highlight.scrollTop = node.scrollTop
      highlight.scrollLeft = node.scrollLeft
    }
    if (document.activeElement !== node) node.focus({ preventScroll: true })
    if (element) {
      element.scrollLeft = scrollLeft
      element.scrollTop = scrollTop
      if (lock) {
        lock.scrollLeft = scrollLeft
        lock.scrollTop = scrollTop
      }
    }
  })

  useEffect(() => () => {
    editorMeasureRef.current?.remove()
    editorMeasureRef.current = null
  }, [])

  useLayoutEffect(() => {
    const node = editorOverlayRef.current
    if (!node || !editorSelectionRequest) return
    node.setSelectionRange(editorSelectionRequest.start, editorSelectionRequest.end)
    onEditorCaret(editorSelectionRequest.start, editorSelectionRequest.end)
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editorSelectionRequest?.id])

  const pinnedTransform = (pinX: boolean, pinY: boolean) => {
    const element = viewportRef.current
    return `translate3d(${pinX ? element?.scrollLeft || 0 : 0}px, ${pinY ? element?.scrollTop || 0 : 0}px, 0)`
  }
  const freezePreviewOffset = freezeDrag
    ? axisOffset(freezeDrag.count, freezeDrag.axis === 'row' ? defaultRow : defaultColumn, freezeDrag.axis === 'row' ? rowMetrics : columnMetrics)
    : 0

  updateFillFrameRef.current = () => {
    const element = viewportRef.current
    const pointer = fillPointerRef.current
    if (!element || !pointer || !fillDraggingRef.current) return
    const rect = element.getBoundingClientRect()
    const edge = 32
    const horizontalStart = rect.left + HEADER_WIDTH
    const verticalStart = rect.top + HEADER_HEIGHT
    const horizontalEnd = rect.right - 12
    const verticalEnd = rect.bottom - 12
    let deltaX = 0
    let deltaY = 0
    if (pointer.x < horizontalStart + edge) deltaX = -Math.min(24, Math.max(4, (horizontalStart + edge - pointer.x) * 0.36))
    else if (pointer.x > horizontalEnd - edge) deltaX = Math.min(24, Math.max(4, (pointer.x - horizontalEnd + edge) * 0.36))
    if (pointer.y < verticalStart + edge) deltaY = -Math.min(24, Math.max(4, (verticalStart + edge - pointer.y) * 0.36))
    else if (pointer.y > verticalEnd - edge) deltaY = Math.min(24, Math.max(4, (pointer.y - verticalEnd + edge) * 0.36))

    if (deltaX || deltaY) {
      element.scrollLeft = Math.max(0, element.scrollLeft + deltaX)
      element.scrollTop = Math.max(0, element.scrollTop + deltaY)
    }
    const localX = Math.max(0, pointer.x - rect.left - HEADER_WIDTH)
    const localY = Math.max(0, pointer.y - rect.top - HEADER_HEIGHT)
    const colPixel = localX < frozenWidth ? localX : localX + element.scrollLeft
    const rowPixel = localY < frozenHeight ? localY : localY + element.scrollTop
    const target = {
      row: axisIndexAt(rowPixel, rows, defaultRow, rowMetrics),
      col: axisIndexAt(colPixel, columns, defaultColumn, columnMetrics),
    }
    if (fillTargetRef.current?.row !== target.row || fillTargetRef.current?.col !== target.col) {
      fillTargetRef.current = target
      setFillTarget(target)
    }
    if (deltaX || deltaY) fillFrameRef.current = window.requestAnimationFrame(() => updateFillFrameRef.current())
  }

  useEffect(() => {
    const finishFill = (apply: boolean, pointerId?: number) => {
      if (!fillDraggingRef.current || (pointerId !== undefined && fillPointerIdRef.current !== pointerId)) return
      fillDraggingRef.current = false
      fillPointerIdRef.current = null
      fillPointerRef.current = null
      if (fillFrameRef.current !== null) window.cancelAnimationFrame(fillFrameRef.current)
      fillFrameRef.current = null
      const target = fillTargetRef.current
      fillTargetRef.current = null
      setFillTarget(null)
      if (apply && target) onFill(target)
      viewportRef.current?.focus({ preventScroll: true })
    }
    const moveFill = (event: globalThis.PointerEvent) => {
      if (!fillDraggingRef.current || fillPointerIdRef.current !== event.pointerId) return
      fillPointerRef.current = { x: event.clientX, y: event.clientY }
      if (fillFrameRef.current !== null) window.cancelAnimationFrame(fillFrameRef.current)
      updateFillFrameRef.current()
    }
    const endFill = (event: globalThis.PointerEvent) => finishFill(true, event.pointerId)
    const cancelFill = (event: globalThis.PointerEvent) => finishFill(false, event.pointerId)
    const moveMouseFill = (event: globalThis.MouseEvent) => {
      if (!fillDraggingRef.current || fillPointerIdRef.current !== -1) return
      fillPointerRef.current = { x: event.clientX, y: event.clientY }
      if (fillFrameRef.current !== null) window.cancelAnimationFrame(fillFrameRef.current)
      updateFillFrameRef.current()
    }
    const endMouseFill = () => finishFill(true, -1)
    const cancelOnBlur = () => finishFill(false)
    window.addEventListener('pointermove', moveFill)
    window.addEventListener('pointerup', endFill)
    window.addEventListener('pointercancel', cancelFill)
    window.addEventListener('mousemove', moveMouseFill)
    window.addEventListener('mouseup', endMouseFill)
    window.addEventListener('blur', cancelOnBlur)
    return () => {
      window.removeEventListener('pointermove', moveFill)
      window.removeEventListener('pointerup', endFill)
      window.removeEventListener('pointercancel', cancelFill)
      window.removeEventListener('mousemove', moveMouseFill)
      window.removeEventListener('mouseup', endMouseFill)
      window.removeEventListener('blur', cancelOnBlur)
      if (fillFrameRef.current !== null) window.cancelAnimationFrame(fillFrameRef.current)
    }
  }, [onFill])

  // Conditional formats are computed for the scrolled window (in 64-row x 16-column chunks,
  // so ordinary scrolling reuses the result) plus the frozen bands, never the rows between.
  const conditionalWindowTop = Math.floor(startRow / 64) * 64
  const conditionalWindowBottom = (Math.floor(endRow / 64) + 1) * 64
  const conditionalWindowLeft = Math.floor(startCol / 16) * 16
  const conditionalWindowRight = (Math.floor(endCol / 16) + 1) * 16
  const conditionalMap = useMemo(() => {
    const windows: OverlayBounds[] = [{ top: conditionalWindowTop, bottom: conditionalWindowBottom, left: conditionalWindowLeft, right: conditionalWindowRight }]
    if (frozenRowCount) windows.push({ top: 0, bottom: frozenRowCount - 1, left: conditionalWindowLeft, right: conditionalWindowRight })
    if (frozenColumnCount) windows.push({ top: conditionalWindowTop, bottom: conditionalWindowBottom, left: 0, right: frozenColumnCount - 1 })
    if (frozenRowCount && frozenColumnCount) windows.push({ top: 0, bottom: frozenRowCount - 1, left: 0, right: frozenColumnCount - 1 })
    let merged: Map<string, ConditionalCellFormat> | null = null
    for (const window of windows) {
      const result = conditionalFormats(window)
      if (!result) return null
      if (!merged) merged = result
      else for (const [address, format] of result) if (!merged.has(address)) merged.set(address, format)
    }
    return merged
  }, [conditionalFormats, conditionalWindowBottom, conditionalWindowLeft, conditionalWindowRight, conditionalWindowTop, frozenColumnCount, frozenRowCount])


  // Unchanged rows reuse last render's element, so React skips them without even creating
  // (and, in development, validating) a new element.
  const rowElementCacheRef = useRef(new Map<string, { props: GridCellRowProps; element: JSX.Element }>())
  const nextRowElements = new Map<string, { props: GridCellRowProps; element: JSX.Element }>()
  const renderGridRows = (pane: CellPane) => rowsByPane[pane].map((coordinates) => {
    const key = `${pane}-${coordinates[0].row}`
    const props: GridCellRowProps = {
      coordinates, mergeMap, conditionalMap, sheet, pane, zoom, columns, frozenRowCount, frozenColumnCount, defaultColumn, defaultRow,
      columnMetrics, rowMetrics, mergeRanges, selectionRef, viewportRef, fillDraggingRef, displayValue, resolvedValue, staleValue,
      onSelection, onBeginEdit, onCellValue, onOpenHyperlink, onPointCell, pointDraggingRef, listOptionsFor, sparklineFor,
    }
    const cached = rowElementCacheRef.current.get(key)
    const entry = cached && gridRowPropsEqual(cached.props, props) ? cached : { props, element: <GridCellRow key={key} {...props} /> }
    nextRowElements.set(key, entry)
    return entry.element
  })

  const renderSelection = (pane: CellPane) => selectionPanes.includes(pane) ? (
        <div
          className={`selection-outline${pane === 'body' ? ' in-body-pane' : ''}`}
          data-selection-pane={pane}
          style={{ left: selectionLeft + (pane === 'body' ? HEADER_WIDTH : 0), top: selectionTop + (pane === 'body' ? HEADER_HEIGHT : 0), width: selectionWidth, height: selectionHeight }}
          aria-hidden="true"
        >{pane === selectionHandlePane && <span
          className="fill-handle"
          onPointerDown={(event) => {
            if (event.button !== 0) return
            event.preventDefault()
            event.stopPropagation()
            fillDraggingRef.current = true
            fillPointerIdRef.current = event.pointerId
            fillPointerRef.current = { x: event.clientX, y: event.clientY }
            fillTargetRef.current = null
            setFillTarget(null)
            try { event.currentTarget.setPointerCapture(event.pointerId) } catch { /* Window listeners still complete the drag. */ }
            updateFillFrameRef.current()
          }}
          onMouseDown={(event) => {
            if (event.button !== 0 || fillDraggingRef.current) return
            event.preventDefault()
            event.stopPropagation()
            fillDraggingRef.current = true
            fillPointerIdRef.current = -1
            fillPointerRef.current = { x: event.clientX, y: event.clientY }
            fillTargetRef.current = null
            setFillTarget(null)
            updateFillFrameRef.current()
          }}
        />}</div>
  ) : null

  const chartGeometryVersion = useMemo(() => ({ zoom, columnMetrics, rowMetrics, defaultColumn, defaultRow }), [columnMetrics, defaultColumn, defaultRow, rowMetrics, zoom])
  const overlayRect = (bounds: OverlayBounds, pane: CellPane) => {
    const top = Math.min(bounds.top, rows - 1)
    const bottom = Math.min(bounds.bottom, rows - 1)
    const left = Math.min(bounds.left, columns - 1)
    const right = Math.min(bounds.right, columns - 1)
    return {
      left: axisOffset(left, defaultColumn, columnMetrics) + (pane === 'body' ? HEADER_WIDTH : 0),
      top: axisOffset(top, defaultRow, rowMetrics) + (pane === 'body' ? HEADER_HEIGHT : 0),
      width: axisOffset(right + 1, defaultColumn, columnMetrics) - axisOffset(left, defaultColumn, columnMetrics),
      height: axisOffset(bottom + 1, defaultRow, rowMetrics) - axisOffset(top, defaultRow, rowMetrics),
    }
  }
  const renderOverlays = (pane: CellPane) => (
    <>
      {spillOutline && intersectingPanes(spillOutline, frozenRowCount, frozenColumnCount).includes(pane) && (
        <div className="spill-outline" data-spill-outline={pane} style={overlayRect(spillOutline, pane)} aria-hidden="true" />
      )}
      {referenceHighlights.map((item, index) => intersectingPanes(item, frozenRowCount, frozenColumnCount).includes(pane) ? (
        <div key={`${index}-${item.top}-${item.left}`} className="range-finder" data-range-finder={index} style={{ ...overlayRect(item, pane), color: item.color }} aria-hidden="true" />
      ) : null)}
      {filterHeaders.map((filterHeader) => Array.from({ length: filterHeader.right - filterHeader.left + 1 }, (_, offset) => filterHeader.left + offset)
        .filter((col) => !hiddenColumns.has(col) && cellPane(filterHeader.row < frozenRowCount, col < frozenColumnCount) === pane && ((col >= startCol - 2 && col <= endCol + 2) || col < frozenColumnCount))
        .map((col) => {
          const rect = overlayRect({ top: filterHeader.row, bottom: filterHeader.row, left: col, right: col }, pane)
          const size = Math.max(14, Math.min(20, rect.height - 3))
          const active = filterHeader.active.has(col)
          const sorted = filterHeader.sort?.column === col
          return (
            <button
              key={`filter-${filterHeader.key}-${col}`}
              type="button"
              className={`filter-button${active ? ' is-active' : ''}${sorted ? ' is-sorted' : ''}`}
              data-filter-column={columnName(col)}
              data-filter-key={filterHeader.key}
              aria-label={`Filter ${columnName(col)}`}
              title={active ? 'Filtered — click to change' : 'Filter and sort'}
              style={{ left: rect.left + rect.width - size - 2, top: rect.top + (rect.height - size) / 2, width: size, height: size }}
              onMouseDown={(event) => { event.preventDefault(); event.stopPropagation() }}
              onClick={(event) => {
                event.stopPropagation()
                const box = event.currentTarget.getBoundingClientRect()
                onFilterButton(filterHeader.key, col, { left: box.left, top: box.top, right: box.right, bottom: box.bottom })
              }}
            >{active ? <Filter size={Math.round(size * 0.6)} /> : <ChevronDown size={Math.round(size * 0.7)} />}</button>
          )
        }))}
      {cutRange && intersectingPanes(cutRange, frozenRowCount, frozenColumnCount).includes(pane) && (
        <div className="cut-marquee" style={overlayRect(cutRange, pane)} aria-hidden="true" />
      )}
      {invalidCells.map((coord) => cellPane(coord.row < frozenRowCount, coord.col < frozenColumnCount) === pane ? (
        <div key={`invalid-${coord.row}-${coord.col}`} className="invalid-circle" style={overlayRect({ top: coord.row, bottom: coord.row, left: coord.col, right: coord.col }, pane)} aria-hidden="true" />
      ) : null)}
      {inputMessage && !editing && cellPane(inputMessage.row < frozenRowCount, inputMessage.col < frozenColumnCount) === pane && (() => {
        const rect = overlayRect({ top: inputMessage.row, bottom: inputMessage.row, left: inputMessage.col, right: inputMessage.col }, pane)
        return (
          <div className="validation-input-message" role="tooltip" style={{ left: rect.left + Math.min(rect.width, 40), top: rect.top + rect.height + 4 }}>
            {inputMessage.title && <strong>{inputMessage.title}</strong>}
            {inputMessage.text && <span>{inputMessage.text}</span>}
          </div>
        )
      })()}
    </>
  )

  const lastSelectableRow = Math.max(0, sheet.rowCount - 1)
  const lastSelectableColumn = Math.max(0, sheet.colCount - 1)
  const wholeColumnsSelected = selectedBounds.top === 0 && selectedBounds.bottom >= lastSelectableRow
  const wholeRowsSelected = selectedBounds.left === 0 && selectedBounds.right >= lastSelectableColumn
  // Header callbacks go through a ref so memoized headers only re-render when their own
  // geometry or selection state changes, not on every scroll-driven grid render.
  const headerHandlersRef = useRef<DimensionHeaderHandlers>(null!)
  headerHandlersRef.current = {
    select: (axis, index, event) => {
      if (event.button !== 0) return
      event.preventDefault()
      viewportRef.current?.focus({ preventScroll: true })
      headerDragRef.current = axis
      if (axis === 'row') {
        const anchorRow = event.shiftKey ? selectionRef.current.anchor.row : index
        onSelection({ anchor: { row: anchorRow, col: lastSelectableColumn }, focus: { row: index, col: 0 } })
      } else {
        const anchorColumn = event.shiftKey ? selectionRef.current.anchor.col : index
        onSelection({ anchor: { row: lastSelectableRow, col: anchorColumn }, focus: { row: 0, col: index } })
      }
    },
    extend: (axis, index, event) => {
      if (headerDragRef.current !== axis || event.buttons !== 1) return
      onSelection({ anchor: selectionRef.current.anchor, focus: axis === 'row' ? { row: index, col: 0 } : { row: 0, col: index } })
    },
    beginResize: beginDimensionResize,
    autoFit: (axis, index) => (axis === 'row' ? autoFitRow(index) : autoFitColumn(index)),
    resizeWithKey: resizeDimensionWithKey,
  }
  const headerSelection = (inSelection: boolean, whole: boolean): DimensionHeaderProps['selection'] => (inSelection ? (whole ? 'full' : 'partial') : 'none')
  const renderColumnHeader = (col: number) => (
    <DimensionHeader
      key={`ch-${col}`}
      axis="column"
      index={col}
      offset={axisOffset(col, defaultColumn, columnMetrics)}
      size={axisSize(col, defaultColumn, columnMetrics)}
      zoom={zoom}
      frozen={col < frozenColumnCount}
      selection={headerSelection(col >= selectedBounds.left && col <= selectedBounds.right, wholeColumnsSelected)}
      handlersRef={headerHandlersRef}
    />
  )
  const renderRowHeader = (row: number) => (
    <DimensionHeader
      key={`rh-${row}`}
      axis="row"
      index={row}
      offset={axisOffset(row, defaultRow, rowMetrics)}
      size={axisSize(row, defaultRow, rowMetrics)}
      zoom={zoom}
      frozen={row < frozenRowCount}
      selection={headerSelection(row >= selectedBounds.top && row <= selectedBounds.bottom, wholeRowsSelected)}
      handlersRef={headerHandlersRef}
    />
  )

  const gridElement = (
    <div
      key="sheet-grid"
      ref={(node) => {
        viewportRef.current = node
        if (node && node !== viewportNode) setViewportNode(node)
      }}
      className={`sheet-viewport${editing ? ' is-editing' : ''}${useImportedDefaults ? ' uses-imported-defaults' : ''}${showGridlines ? '' : ' hides-gridlines'}${showNotes ? '' : ' hides-notes'}`}
      style={{ '--cell-zoom': zoom } as CSSProperties}
      tabIndex={0}
      role="grid"
      aria-label={`${sheet.name} spreadsheet grid`}
      aria-rowcount={rows}
      aria-colcount={columns}
      aria-activedescendant={`cell-${sheet.id}-${addressOf(selection.focus)}`}
      onKeyDown={onKeyDown}
      onScroll={(event) => handleViewportScroll(event.currentTarget)}
      onContextMenu={handleViewportContextMenu}
    >
      <div className="sheet-canvas" style={{ width: totalWidth, height: totalHeight }}>
        <div className="column-header-layer" style={{ width: totalWidth - HEADER_WIDTH, height: HEADER_HEIGHT }}>
          {renderedColumns.filter((col) => col >= frozenColumnCount).map(renderColumnHeader)}
        </div>
        {frozenColumnCount > 0 && (
          <div className="column-header-layer is-frozen" style={{ width: frozenWidth, height: HEADER_HEIGHT }}>
            {renderedColumns.filter((col) => col < frozenColumnCount).map(renderColumnHeader)}
          </div>
        )}
        <div className="row-header-layer" style={{ width: HEADER_WIDTH, height: totalHeight - HEADER_HEIGHT }}>
          {renderedRows.filter((row) => row >= frozenRowCount).map(renderRowHeader)}
        </div>
        {frozenRowCount > 0 && (
          <div className="row-header-layer is-frozen" style={{ width: HEADER_WIDTH, height: frozenHeight }}>
            {renderedRows.filter((row) => row < frozenRowCount).map(renderRowHeader)}
          </div>
        )}

        {renderGridRows('body')}
        {renderSelection('body')}
        {renderOverlays('body')}
        {frozenRowCount > 0 && (
          <div className="frozen-pane-layer is-row" style={{ width: totalWidth - HEADER_WIDTH, height: frozenHeight }}>
            {renderGridRows('frozen-row')}
            {renderSelection('frozen-row')}
            {renderOverlays('frozen-row')}
          </div>
        )}
        {frozenColumnCount > 0 && (
          <div className="frozen-pane-layer is-column" style={{ width: frozenWidth, height: totalHeight - HEADER_HEIGHT }}>
            {renderGridRows('frozen-column')}
            {renderSelection('frozen-column')}
            {renderOverlays('frozen-column')}
          </div>
        )}
        {frozenRowCount > 0 && frozenColumnCount > 0 && (
          <div className="frozen-pane-layer is-corner" style={{ width: frozenWidth, height: frozenHeight }}>
            {renderGridRows('frozen-corner')}
            {renderSelection('frozen-corner')}
            {renderOverlays('frozen-corner')}
          </div>
        )}

        {frozenColumnCount > 0 && frozenWidth > 0 && (
          <button
            type="button"
            role="separator"
            aria-orientation="vertical"
            aria-label="Adjust frozen columns"
            aria-valuemin={0}
            aria-valuemax={columns}
            aria-valuenow={frozenColumnCount}
            aria-valuetext={`${frozenColumnCount} frozen ${frozenColumnCount === 1 ? 'column' : 'columns'}`}
            className="freeze-divider is-vertical"
            data-freeze-divider="columns"
            style={{ left: HEADER_WIDTH + frozenWidth - 2, top: HEADER_HEIGHT, height: Math.max(0, viewport.height - HEADER_HEIGHT) }}
            onPointerDown={(event) => beginFreezeDrag('column', event)}
            onKeyDown={(event) => freezeHandleKeyDown('column', event)}
          />
        )}
        {frozenRowCount > 0 && frozenHeight > 0 && (
          <button
            type="button"
            role="separator"
            aria-orientation="horizontal"
            aria-label="Adjust frozen rows"
            aria-valuemin={0}
            aria-valuemax={rows}
            aria-valuenow={frozenRowCount}
            aria-valuetext={`${frozenRowCount} frozen ${frozenRowCount === 1 ? 'row' : 'rows'}`}
            className="freeze-divider is-horizontal"
            data-freeze-divider="rows"
            style={{ left: HEADER_WIDTH, top: HEADER_HEIGHT + frozenHeight - 2, width: Math.max(0, viewport.width - HEADER_WIDTH) }}
            onPointerDown={(event) => beginFreezeDrag('row', event)}
            onKeyDown={(event) => freezeHandleKeyDown('row', event)}
          />
        )}

        {freezeDrag && (
          <div
            className={`freeze-drag-preview ${freezeDrag.axis === 'row' ? 'is-horizontal' : 'is-vertical'}`}
            style={freezeDrag.axis === 'row'
              ? { left: HEADER_WIDTH, top: HEADER_HEIGHT + freezePreviewOffset - 2, width: Math.max(0, viewport.width - HEADER_WIDTH) }
              : { left: HEADER_WIDTH + freezePreviewOffset - 2, top: HEADER_HEIGHT, height: Math.max(0, viewport.height - HEADER_HEIGHT) }}
            aria-hidden="true"
          />
        )}

        {fillPreview && (
          <div
            ref={fillPreviewOverlayRef}
            className={`fill-preview${fillPinnedX || fillPinnedY ? ' scroll-pin' : ''}${fillPinnedX ? ' pin-x' : ''}${fillPinnedY ? ' pin-y' : ''}`}
            style={{ left: fillLeft, top: fillTop, width: fillWidth, height: fillHeight, transform: pinnedTransform(fillPinnedX, fillPinnedY) }}
            aria-hidden="true"
          />
        )}



        {editing && editorGeometry && (() => {
          const { frozenColumn, frozenRow, left, top, width, height, font, background, wrapText, textAlign } = editorGeometry
          const pinClasses = `${frozenColumn || frozenRow ? ' scroll-pin' : ''}${frozenColumn ? ' pin-x' : ''}${frozenRow ? ' pin-y' : ''}`
          return (
            <>
            <textarea
              ref={editorOverlayRef}
              rows={1}
              spellCheck={false}
              className={`cell-editor${pinClasses}${editorParts ? ' is-formula-draft' : ''}`}
              style={{
                left,
                top,
                width,
                height,
                ...fontCss(font, zoom),
                background,
                textAlign,
                whiteSpace: wrapText ? 'pre-wrap' : 'pre',
                transform: pinnedTransform(frozenColumn, frozenRow),
              }}
              value={editing.draft}
              onChange={(event) => {
                onDraft(event.target.value)
                onEditorCaret(event.target.selectionStart, event.target.selectionEnd)
              }}
              onSelect={(event) => onEditorCaret(event.currentTarget.selectionStart, event.currentTarget.selectionEnd)}
              onScroll={(event) => {
                const highlight = editorHighlightRef.current
                if (highlight) {
                  highlight.scrollTop = event.currentTarget.scrollTop
                  highlight.scrollLeft = event.currentTarget.scrollLeft
                }
              }}
              onMouseDown={(event) => event.stopPropagation()}
              onFocus={(event) => {
                onEditorFocusChange(true, event.currentTarget)
                editorCommitViaKeyRef.current = false
                // Typing a character opens the editor with that character already seeded,
                // and F2 opens it on the existing text.  Selecting it would make the next
                // keystroke replace it, so park the caret at the end instead.
                const viewport = viewportRef.current
                const lock = editScrollLockRef.current
                const scrollLeft = lock?.scrollLeft ?? viewport?.scrollLeft ?? 0
                const scrollTop = lock?.scrollTop ?? viewport?.scrollTop ?? 0
                const end = event.currentTarget.value.length
                event.currentTarget.setSelectionRange(end, end)
                const restore = () => {
                  if (!viewport) return
                  viewport.scrollLeft = scrollLeft
                  viewport.scrollTop = scrollTop
                  if (lock) {
                    lock.scrollLeft = scrollLeft
                    lock.scrollTop = scrollTop
                  }
                  syncPinnedOverlays(scrollLeft, scrollTop)
                }
                // setSelectionRange often scrolls the ancestor after this handler returns.
                restore()
                queueMicrotask(restore)
                window.requestAnimationFrame(() => {
                  restore()
                  window.requestAnimationFrame(restore)
                })
              }}
              onBlur={() => {
                onEditorFocusChange(false, null)
                if (editorCommitViaKeyRef.current) { editorCommitViaKeyRef.current = false; return }
                onCommitEdit()
              }}
              onKeyDown={(event) => {
                if (onEditorKeyDown(event)) {
                  if (event.defaultPrevented && !editorOverlayRef.current) editorCommitViaKeyRef.current = true
                  return
                }
                if (event.key === 'Escape') { event.preventDefault(); onCancelEdit() }
                else if (event.key === 'Enter' && !event.altKey) { event.preventDefault(); editorCommitViaKeyRef.current = true; onCommitEdit(event.shiftKey ? 'up' : 'down'); viewportRef.current?.focus({ preventScroll: true }) }
                else if (event.key === 'Tab') { event.preventDefault(); editorCommitViaKeyRef.current = true; onCommitEdit(event.shiftKey ? 'left' : 'right'); viewportRef.current?.focus({ preventScroll: true }) }
              }}
            />
            {editorParts && (
              <FormulaHighlight
                ref={editorHighlightRef}
                className={`cell-editor fa-highlight-layer${pinClasses}`}
                parts={editorParts}
                style={{
                  left,
                  top,
                  width,
                  height,
                  ...fontCss(font, zoom),
                  textAlign,
                  whiteSpace: wrapText ? 'pre-wrap' : 'pre',
                  transform: pinnedTransform(frozenColumn, frozenRow),
                }}
              />
            )}
            </>
          )
        })()}

        {renderCharts?.(
          {
            cellRect: (row, col) => ({
              left: HEADER_WIDTH + axisOffset(col, defaultColumn, columnMetrics),
              top: HEADER_HEIGHT + axisOffset(row, defaultRow, rowMetrics),
              width: axisSize(col, defaultColumn, columnMetrics),
              height: axisSize(row, defaultRow, rowMetrics),
            }),
            cellAtPoint: (x, y) => ({
              row: axisIndexAt(Math.max(0, y - HEADER_HEIGHT), rows, defaultRow, rowMetrics),
              col: axisIndexAt(Math.max(0, x - HEADER_WIDTH), columns, defaultColumn, columnMetrics),
            }),
          },
          { left: viewport.scrollLeft, top: viewport.scrollTop, width: viewport.width, height: viewport.height },
          chartGeometryVersion,
        )}

        <div className="grid-corner" style={{ width: HEADER_WIDTH, height: HEADER_HEIGHT }}>
          <button
            type="button"
            role="separator"
            aria-orientation="horizontal"
            aria-label="Freeze rows"
            aria-valuemin={0}
            aria-valuemax={rows}
            aria-valuenow={frozenRowCount}
            aria-valuetext={`${frozenRowCount} frozen ${frozenRowCount === 1 ? 'row' : 'rows'}`}
            title="Drag down to freeze rows. Drag a frozen divider back here to unfreeze."
            className="freeze-origin-handle is-horizontal"
            data-freeze-handle="rows"
            onPointerDown={(event) => beginFreezeDrag('row', event)}
            onKeyDown={(event) => freezeHandleKeyDown('row', event)}
          />
          <button
            type="button"
            role="separator"
            aria-orientation="vertical"
            aria-label="Freeze columns"
            aria-valuemin={0}
            aria-valuemax={columns}
            aria-valuenow={frozenColumnCount}
            aria-valuetext={`${frozenColumnCount} frozen ${frozenColumnCount === 1 ? 'column' : 'columns'}`}
            title="Drag right to freeze columns. Drag a frozen divider back here to unfreeze."
            className="freeze-origin-handle is-vertical"
            data-freeze-handle="columns"
            onPointerDown={(event) => beginFreezeDrag('column', event)}
            onKeyDown={(event) => freezeHandleKeyDown('column', event)}
          />
        </div>
      </div>
    </div>
  )
  rowElementCacheRef.current = nextRowElements
  // The frame is always rendered (and the grid keyed) so switching between sheets with and
  // without outlines never remounts the scrolling viewport.
  if (!outline || (!outline.rowDepth && !outline.columnDepth)) {
    return <div className="sheet-outline-frame" style={{ gridTemplateColumns: '0px minmax(0, 1fr)', gridTemplateRows: '0px minmax(0, 1fr)' }}>{gridElement}</div>
  }

  const OUTLINE_STEP = 14
  const rowStrip = outline.rowDepth ? 6 + (outline.rowDepth + 1) * OUTLINE_STEP : 0
  const columnStrip = outline.columnDepth ? 6 + (outline.columnDepth + 1) * OUTLINE_STEP : 0
  const levelButtons = (axis: OutlineAxis, depth: number) => Array.from({ length: depth + 1 }, (_, index) => (
    <button
      key={`${axis}-level-${index + 1}`}
      type="button"
      className="outline-level"
      data-outline-level={`${axis}-${index + 1}`}
      title={`Show ${axis === 'row' ? 'row' : 'column'} level ${index + 1}`}
      aria-label={`Show ${axis === 'row' ? 'row' : 'column'} outline level ${index + 1}`}
      style={axis === 'row' ? { left: 3 + index * OUTLINE_STEP, top: (HEADER_HEIGHT - 13) / 2 } : { top: 3 + index * OUTLINE_STEP, left: (HEADER_WIDTH - 13) / 2 }}
      onMouseDown={(event) => event.preventDefault()}
      onClick={() => onOutlineLevel?.(axis, index + 1)}
    >{index + 1}</button>
  ))
  const groupMarks = (axis: OutlineAxis, groups: OutlineGroup[], first: number, last: number, base: number) => groups
    .filter((group) => group.end >= first && group.start <= last)
    .map((group) => {
      const row = axis === 'row'
      const offset = (index: number) => row ? axisOffset(index, defaultRow, rowMetrics) : axisOffset(index, defaultColumn, columnMetrics)
      const size = (index: number) => row ? axisSize(index, defaultRow, rowMetrics) : axisSize(index, defaultColumn, columnMetrics)
      const lane = 3 + (group.level - 1) * OUTLINE_STEP + 6
      const from = offset(group.start) - base
      const to = offset(group.end + 1) - base
      const button = group.summary !== null ? offset(group.summary) + size(group.summary) / 2 - base : to
      const key = `${axis}-${group.level}-${group.start}`
      const toggle = (
        <button
          key={`${key}-toggle`}
          type="button"
          className={`outline-toggle${group.collapsed ? ' is-collapsed' : ''}`}
          data-outline-group={`${axis}-${group.level}-${group.start + 1}-${group.end + 1}`}
          aria-expanded={!group.collapsed}
          aria-label={`${group.collapsed ? 'Expand' : 'Collapse'} ${row ? 'rows' : 'columns'} ${row ? group.start + 1 : columnName(group.start)}–${row ? group.end + 1 : columnName(group.end)}`}
          title={group.collapsed ? 'Show detail' : 'Hide detail'}
          style={row ? { left: lane - 6, top: button - 6 } : { top: lane - 6, left: button - 6 }}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => onOutlineToggle?.(axis, group)}
        >{group.collapsed ? '+' : '−'}</button>
      )
      if (to - from < 2) return toggle
      const lineStart = from + 2
      const lineEnd = group.summary !== null ? Math.min(to, button - 7) : to - 7
      return [
        lineEnd > lineStart ? <span key={`${key}-line`} className="outline-line" style={row ? { left: lane, top: lineStart, width: 1, height: lineEnd - lineStart } : { top: lane, left: lineStart, height: 1, width: lineEnd - lineStart }} /> : null,
        lineEnd > lineStart ? <span key={`${key}-tick`} className="outline-line" style={row ? { left: lane - 4, top: lineStart, width: 5, height: 1 } : { top: lane - 4, left: lineStart, height: 5, width: 1 }} /> : null,
        toggle,
      ]
    })
  return (
    <div className="sheet-outline-frame" style={{ gridTemplateColumns: `${rowStrip}px minmax(0, 1fr)`, gridTemplateRows: `${columnStrip}px minmax(0, 1fr)` }}>
      {rowStrip > 0 && columnStrip > 0 && <div key="outline-corner" className="outline-corner" />}
      {columnStrip > 0 && (
        <div key="outline-columns" className="outline-strip is-columns" aria-label="Column outline">
          <div className="outline-layer" style={{ left: 0, width: HEADER_WIDTH, top: 0, bottom: 0 }}>{levelButtons('column', outline.columnDepth)}</div>
          {frozenColumnCount > 0 && (
            <div className="outline-layer" style={{ left: HEADER_WIDTH, width: frozenWidth, top: 0, bottom: 0 }}>
              {groupMarks('column', outline.columns, 0, frozenColumnCount - 1, 0)}
            </div>
          )}
          <div className="outline-layer" style={{ left: HEADER_WIDTH + frozenWidth, right: 0, top: 0, bottom: 0 }}>
            <div ref={columnOutlineContentRef} className="outline-content">
              {groupMarks('column', outline.columns, Math.max(frozenColumnCount, startCol - 20), endCol + 20, frozenWidth)}
            </div>
          </div>
        </div>
      )}
      {rowStrip > 0 && (
        <div key="outline-rows" className="outline-strip is-rows" aria-label="Row outline">
          <div className="outline-layer" style={{ top: 0, height: HEADER_HEIGHT, left: 0, right: 0 }}>{levelButtons('row', outline.rowDepth)}</div>
          {frozenRowCount > 0 && (
            <div className="outline-layer" style={{ top: HEADER_HEIGHT, height: frozenHeight, left: 0, right: 0 }}>
              {groupMarks('row', outline.rows, 0, frozenRowCount - 1, 0)}
            </div>
          )}
          <div className="outline-layer" style={{ top: HEADER_HEIGHT + frozenHeight, bottom: 0, left: 0, right: 0 }}>
            <div ref={rowOutlineContentRef} className="outline-content">
              {groupMarks('row', outline.rows, Math.max(frozenRowCount, startRow - 40), endRow + 40, frozenHeight)}
            </div>
          </div>
        </div>
      )}
      {gridElement}
    </div>
  )
}

function TextPromptDialog({ request, onClose }: { request: TextPromptRequest; onClose: (value: string | null) => void }) {
  const [value, setValue] = useState(request.initialValue || '')
  const fieldRef = useRef<HTMLInputElement | HTMLTextAreaElement>(null)
  useEffect(() => {
    fieldRef.current?.focus()
    fieldRef.current?.select()
  }, [])
  return (
    <div className="prompt-overlay" onPointerDown={(event) => { if (event.target === event.currentTarget) onClose(null) }}>
      <form
        className="prompt-card"
        role="dialog"
        aria-modal="true"
        aria-label={request.title}
        onSubmit={(event) => { event.preventDefault(); onClose(value) }}
        onKeyDown={(event) => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onClose(null) } }}
      >
        <h2>{request.title}</h2>
        <label>
          <span>{request.label}</span>
          {request.multiline ? (
            <textarea
              ref={fieldRef as RefObject<HTMLTextAreaElement>}
              rows={5}
              value={value}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => { if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) { event.preventDefault(); onClose(value) } }}
            />
          ) : (
            <input
              ref={fieldRef as RefObject<HTMLInputElement>}
              value={value}
              onChange={(event) => setValue(event.target.value)}
            />
          )}
        </label>
        <div className="prompt-actions">
          {request.multiline && <span className="prompt-hint">Ctrl+Enter to save</span>}
          <button type="button" className="secondary-action" onClick={() => onClose(null)}>Cancel</button>
          <button type="submit" className="primary-action">OK</button>
        </div>
      </form>
    </div>
  )
}

function IconButton({ label, children, active, disabled, onClick }: {
  label: string
  children: React.ReactNode
  active?: boolean
  disabled?: boolean
  onClick?: () => void
}) {
  return (
    <button type="button" className={`tool-button${active ? ' is-active' : ''}`} aria-label={label} title={label} disabled={disabled} onClick={onClick}>
      {children}
    </button>
  )
}

function TitleBar({ fileName, dirty, onClose }: { fileName?: string; dirty: boolean; onClose: () => void }) {
  const [maximized, setMaximized] = useState(false)
  useEffect(() => window.simpleCalc.onMaximized(setMaximized), [])
  return (
    <header className="titlebar">
      <div className="titlebar-drag">
        <img src={appIcon} alt="" />
        <span className="product-name">simple_calc</span>
        {fileName && <><span className="title-separator">/</span><span className="document-title">{fileName}</span></>}
        {dirty && <span className="dirty-dot" title="Unsaved changes" />}
      </div>
      <div className="window-controls">
        <button type="button" aria-label="Minimize" onClick={() => window.simpleCalc.minimize()}><Minus size={14} /></button>
        <button type="button" aria-label={maximized ? 'Restore' : 'Maximize'} onClick={() => window.simpleCalc.toggleMaximize()}>
          {maximized ? <Square size={11} /> : <Maximize2 size={13} />}
        </button>
        <button type="button" className="window-close" aria-label="Close" onClick={onClose}><X size={15} /></button>
      </div>
    </header>
  )
}

function Welcome({ recent, busy, onNew, onOpen, onRecent, onRemoveRecent, onDrop }: {
  recent: RecentWorkbook[]
  busy: boolean
  onNew: () => void
  onOpen: () => void
  onRecent: (item: RecentWorkbook) => void
  onRemoveRecent: (path: string) => void
  onDrop: (event: DragEvent<HTMLElement>) => void
}) {
  const [dragging, setDragging] = useState(false)
  const dragDepth = useRef(0)
  return (
    <main
      className="welcome"
      onDragEnter={(event) => { event.preventDefault(); dragDepth.current += 1; setDragging(true) }}
      onDragOver={(event) => event.preventDefault()}
      onDragLeave={(event) => { event.preventDefault(); dragDepth.current -= 1; if (dragDepth.current <= 0) setDragging(false) }}
      onDrop={(event) => { dragDepth.current = 0; setDragging(false); onDrop(event) }}
    >
      <section className="welcome-hero">
        <div className="hero-mark"><img src={appIcon} alt="simple_calc" /></div>
        <span className="eyebrow">A calm spreadsheet workspace</span>
        <h1>Numbers, without the noise.</h1>
        <p>Open familiar spreadsheet formats, keep formulas intact, and share clean modern workbooks across Excel, Sheets, Numbers, and Calc.</p>
        <div className="hero-actions">
          <button type="button" className="primary-action" onClick={onNew} disabled={busy}><FilePlus2 size={17} /> New spreadsheet</button>
          <button type="button" className="secondary-action" onClick={onOpen} disabled={busy}><FolderOpen size={17} /> Open a file</button>
        </div>
        <div className={`drop-zone${dragging ? ' is-dragging' : ''}`}>
          <FileSpreadsheet size={20} />
          <span>{dragging ? 'Drop to open' : 'Or drop a spreadsheet anywhere'}</span>
        </div>
        <div className="format-line"><span>XLSX</span><span>XLS</span><span>XLSB</span><span>ODS</span><span>CSV</span><span>TSV</span><span>NUMBERS</span></div>
      </section>
      <aside className="recent-panel">
        <div className="recent-heading">
          <div><span className="eyebrow">Your workspace</span><h2>Recent files</h2></div>
          <span>{recent.length || '—'}</span>
        </div>
        <div className="recent-list">
          {recent.map((item) => (
            <div className="recent-item" key={item.path}>
              <button type="button" className="recent-main" onClick={() => onRecent(item)}>
                <span className="recent-icon"><FileSpreadsheet size={17} /></span>
                <span className="recent-copy"><strong>{item.name}</strong><small>{item.path}</small></span>
                <span className="recent-format">{item.format.toUpperCase()}</span>
              </button>
              <button type="button" className="recent-remove" aria-label={`Remove ${item.name} from recents`} onClick={() => onRemoveRecent(item.path)}><X size={13} /></button>
            </div>
          ))}
          {!recent.length && (
            <div className="recent-empty"><FileSpreadsheet size={22} /><strong>No recent spreadsheets</strong><span>Files stay on this computer and appear here after you open them.</span></div>
          )}
        </div>
        <div className="local-note"><Check size={14} /><p><strong>Local by default.</strong> Your workbook is opened and saved on this computer—there is no upload step.</p></div>
      </aside>
    </main>
  )
}

export default function App() {
  const [documentFile, setDocumentFile] = useState<OpenWorkbook | null>(null)
  const [workbook, setWorkbook] = useState<WorkbookModel | null>(null)
  const [selection, setSelection] = useState<Selection>({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
  const [editing, setEditing] = useState<EditingState | null>(null)
  const [formulaDraft, setFormulaDraft] = useState('')
  const [dirty, setDirty] = useState(false)
  const [busy, setBusy] = useState('')
  const [toast, setToast] = useState('')
  const [recent, setRecent] = useState<RecentWorkbook[]>(readRecent)
  const [zoom, setZoom] = useState(1)
  const [warningOpen, setWarningOpen] = useState(false)
  const [exportOpen, setExportOpen] = useState(false)
  const [printOpen, setPrintOpen] = useState(false)
  const [historyTick, setHistoryTick] = useState(0)
  const [statusStats, setStatusStats] = useState<Set<AggregateMode>>(readStatusStats)
  const [tabDropTarget, setTabDropTarget] = useState<string | null>(null)
  const [nameManagerOpen, setNameManagerOpen] = useState(false)
  const [formatDialogTab, setFormatDialogTab] = useState<FormatCellsTab | null>(null)
  const [conditionalPanelOpen, setConditionalPanelOpen] = useState(false)
  const [cutRange, setCutRange] = useState<({ sheetId: string } & OverlayBounds) | null>(null)
  const [pasteSpecialOpen, setPasteSpecialOpen] = useState(false)
  const [selectedChartId, setSelectedChartId] = useState<string | null>(null)
  const [chartEditorId, setChartEditorId] = useState<string | null>(null)
  const [searchOpen, setSearchOpen] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [searchIndex, setSearchIndex] = useState(-1)
  const [replaceOpen, setReplaceOpen] = useState(false)
  const [replaceValue, setReplaceValue] = useState('')
  const [searchOptions, setSearchOptions] = useState<SearchOptions>({ matchCase: false, entireCell: false, workbook: false, lookIn: 'values' })
  const [searchOptionsOpen, setSearchOptionsOpen] = useState(false)
  const [contextMenu, setContextMenu] = useState<{ kind: GridContextTarget['kind'] | 'sheet-tab' | 'status' | 'image'; x: number; y: number; sheetId?: string; imageId?: string } | null>(null)
  const [selectedImageId, setSelectedImageId] = useState<string | null>(null)
  const [protectDialogOpen, setProtectDialogOpen] = useState(false)
  const [sparklineDialog, setSparklineDialog] = useState<{ kind: SparklineKind; data: string; location: string } | null>(null)
  // Formula auditing arrows on one sheet; each Trace press follows one more level.
  const [traceState, setTraceState] = useState<{ sheetId: string; arrows: TraceArrow[]; precedents: string[]; dependents: string[] } | null>(null)
  const [pivotDialog, setPivotDialog] = useState<{ range: string; location: string } | null>(null)
  // The editor follows the active cell into a pivot table; closing hides it for that pivot.
  const [pivotEditorHidden, setPivotEditorHidden] = useState<string | null>(null)
  const [showFormulaBar, setShowFormulaBar] = useState(true)
  const [showFormulas, setShowFormulas] = useState(false)
  const [showNotes, setShowNotes] = useState(true)
  const [immersive, setImmersive] = useState(false)
  const [textPrompt, setTextPrompt] = useState<TextPromptState | null>(null)
  const [nameBoxDraft, setNameBoxDraft] = useState<string | null>(null)
  const [nameBoxInvalid, setNameBoxInvalid] = useState(false)
  const [statsSelection, setStatsSelection] = useState(selection)
  const [formulaInput, setFormulaInput] = useState<'cell' | 'bar' | null>(null)
  const [assistCaret, setAssistCaret] = useState(0)
  const [assistIndex, setAssistIndex] = useState(0)
  const [assistDismissed, setAssistDismissed] = useState(false)
  const [assistAnchor, setAssistAnchor] = useState<{ left: number; top: number; bottom: number; width: number } | null>(null)
  const [editorSelectionRequest, setEditorSelectionRequest] = useState<{ start: number; end: number; id: number } | null>(null)
  const formulaBarRef = useRef<HTMLTextAreaElement>(null)
  const formulaBarSelectionRef = useRef<{ start: number; end: number } | null>(null)
  const formulaInputRef = useRef<'cell' | 'bar' | null>(null)
  const formulaInputElementRef = useRef<HTMLTextAreaElement | null>(null)
  const assistCaretRef = useRef(0)
  const selectionRequestIdRef = useRef(0)
  /** The reference most recently inserted by pointing, so arrows/drags replace it. */
  const pointRef = useRef<{ span: { start: number; end: number }; anchor: Coord; focus: Coord } | null>(null)
  const moveSelectionRef = useRef<(rowDelta: number, colDelta: number, extend?: boolean) => void>(() => {})
  const toggleFilterRef = useRef<() => void>(() => {})
  const openShiftDialogRef = useRef<(mode: 'insert' | 'delete') => void>(() => {})
  const refreshPivotsRef = useRef<(all: boolean) => void>(() => {})
  const groupSelectionRef = useRef<(axis: OutlineAxis | null, delta: 1 | -1) => void>(() => {})
  const openCreateTableRef = useRef<() => void>(() => {})
  const [pointerDown, setPointerDown] = useState(false)
  const historyRef = useRef<HistoryEntry[]>([])
  const futureRef = useRef<HistoryEntry[]>([])
  const internalClipboard = useRef<InternalClipboard | null>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const searchOriginRef = useRef('A1')
  const promptIdRef = useRef(0)
  const formulaBarDirtyRef = useRef<{ sheetId: string; address: string; draft: string } | null>(null)
  const workbookRef = useRef(workbook)
  const documentRef = useRef(documentFile)
  const dirtyRef = useRef(dirty)
  const selectionRef = useRef(selection)
  const editingRef = useRef(editing)
  // Assigned once commitCell exists; held in a ref so save/close/discard keep stable
  // identities and never re-subscribe their IPC listeners on every keystroke.
  const flushPendingEditsRef = useRef<() => void>(() => {})
  // In-app confirmations (native confirm() dialogs block the renderer and look foreign).
  const confirmRef = useRef<(message: string) => Promise<boolean>>(async (message) => window.confirm(message))
  // One calculation engine follows the open workbook across revisions, recalculating only
  // what an edit invalidates. Hints from immer patches spare it from diffing big sheets.
  const engineRef = useRef<CalculationEngine | null>(null)
  const engineHintRef = useRef<{ from: WorkbookModel; to: WorkbookModel; hint: ChangeHint } | null>(null)

  useEffect(() => { workbookRef.current = workbook }, [workbook])
  useEffect(() => { documentRef.current = documentFile }, [documentFile])
  useEffect(() => { dirtyRef.current = dirty }, [dirty])
  useEffect(() => { selectionRef.current = selection }, [selection])
  useEffect(() => { editingRef.current = editing }, [editing])
  const formulaDraftRef = useRef(formulaDraft)
  formulaDraftRef.current = formulaDraft
  useEffect(() => { writeRecent(recent) }, [recent])
  useEffect(() => {
    if (!toast) return
    const timeout = window.setTimeout(() => setToast(''), 3800)
    return () => window.clearTimeout(timeout)
  }, [toast])

  const activeSheet = useMemo(() => workbook?.sheets.find((sheet) => sheet.id === workbook.activeSheetId) || workbook?.sheets[0] || null, [workbook])
  const gridlinesVisible = activeSheet?.views?.[0]?.showGridLines !== false
  const activeAddress = mergeMasterAddress(activeSheet, addressOf(selection.focus))
  const activeCell = activeSheet?.cells[activeAddress]
  const hiddenRowSet = useMemo(() => hiddenIndexSet(activeSheet?.hiddenRows), [activeSheet?.hiddenRows])
  const hiddenColSet = useMemo(() => hiddenIndexSet(activeSheet?.hiddenCols), [activeSheet?.hiddenCols])

  const fitSheetView = useCallback((mode: 'width' | 'sheet') => {
    if (!activeSheet) return
    const viewport = document.querySelector<HTMLElement>('.sheet-viewport')
    if (!viewport) return
    // Fit the actual form/content, not the editor's extra blank rows/columns.
    // This changes the view only; source dimensions and formatting stay intact.
    let bottom = 0
    let right = 0
    for (const address of Object.keys(activeSheet.cells)) {
      const position = coordOf(address)
      if (position) { bottom = Math.max(bottom, position.row); right = Math.max(right, position.col) }
    }
    for (const merge of activeSheet.merges || []) {
      const range = mergeBounds(merge)
      if (range) { bottom = Math.max(bottom, range.bottom); right = Math.max(right, range.right) }
    }
    const imported = true
    const sourceWidth = Number(activeSheet.properties?.defaultColWidth)
    const sourceHeight = Number(activeSheet.properties?.defaultRowHeight)
    const defaultWidth = sourceWidth > 0 ? columnPixelWidth(sourceWidth, 1) : imported ? IMPORTED_COL_WIDTH : DEFAULT_COL_WIDTH
    const defaultHeight = sourceHeight > 0 ? rowPixelHeight(sourceHeight, 1) : imported ? IMPORTED_ROW_HEIGHT : DEFAULT_ROW_HEIGHT
    const extent = (count: number, fallback: number, values: Record<string, number>, hidden: number[] | undefined, pixels: (value: number, zoom: number) => number) => {
      const hiddenSet = new Set(hidden || [])
      let total = count * fallback
      for (const [key, value] of Object.entries(values || {})) {
        if (Number(key) >= 1 && Number(key) <= count && !hiddenSet.has(Number(key))) total += pixels(value, 1) - fallback
      }
      for (const index of hiddenSet) if (index >= 1 && index <= count) total -= fallback
      return Math.max(1, total)
    }
    const width = extent(right + 1, defaultWidth, activeSheet.colWidths, activeSheet.hiddenCols, columnPixelWidth)
    const height = extent(bottom + 1, defaultHeight, activeSheet.rowHeights, activeSheet.hiddenRows, rowPixelHeight)
    const ratio = Math.min((viewport.clientWidth - HEADER_WIDTH - 14) / width, mode === 'sheet' ? (viewport.clientHeight - HEADER_HEIGHT - 14) / height : 2)
    setZoom(clamp(Math.floor(ratio * 100) / 100, 0.1, 2))
    viewport.scrollTo({ top: 0, left: 0 })
  }, [activeSheet, workbook?.metadata?.sourceName])

  const syncEngine = useCallback((target: WorkbookModel) => {
    let engine = engineRef.current
    if (!engine) {
      engine = new CalculationEngine(target)
      engineRef.current = engine
      engineHintRef.current = null
      return engine
    }
    const pending = engineHintRef.current
    const hint = pending && pending.to === target && pending.from === engine.current ? pending.hint : undefined
    engine.update(target, hint)
    if (pending?.to === target) engineHintRef.current = null
    return engine
  }, [])
  const recordEngineHint = useCallback((from: WorkbookModel, to: WorkbookModel, patches: Patch[]) => {
    const hint = changeHintFromPatches(patches, to)
    const pending = engineHintRef.current
    engineHintRef.current = pending && pending.to === from
      ? { from: pending.from, to, hint: mergeChangeHints(pending.hint, hint) }
      : { from, to, hint }
  }, [])
  const resetEngine = useCallback((target: WorkbookModel | null) => {
    engineRef.current = target ? new CalculationEngine(target) : null
    engineHintRef.current = null
  }, [])
  /** Workbook with fresh formula results and spill metadata, for save/print/export. */
  const calculatedWorkbook = useCallback((target: WorkbookModel) => syncEngine(target).withResults(), [syncEngine])
  const calcEngine = useMemo(() => (workbook ? syncEngine(workbook) : null), [syncEngine, workbook])
  const sheetById = useMemo(() => new Map((workbook?.sheets || []).map((sheet) => [sheet.id, sheet])), [workbook])
  const displayValue = useCallback((sheetId: string, address: string) => {
    if (!calcEngine) return ''
    const sheet = sheetById.get(sheetId)
    const cell = sheet?.cells[address]
    if (showFormulas && cell?.formula) return `=${cell.formula}`
    const value = calcEngine.getValue(sheetId, address)
    // In-cell charts have no text (they are drawn by the grid).
    if (isSparklineValue(value)) return ''
    if (!cell) {
      if (value === null || value === undefined) return ''
      // A spilled array member takes the anchor's number format when it has none of its own.
      const anchor = calcEngine.spillAnchorOf(sheetId, address)
      const anchorCell = anchor ? sheet?.cells[anchor] : undefined
      return formatScalar(value, anchorCell?.numFmt || anchorCell?.style?.numFmt)
    }
    const cachedResultMatches = cell.formula ? Object.is(value, cell.result) : !cell.arrayMember && Object.is(value, cell.value)
    // A formula saved without a cached value arrives with display === '', which must never
    // win over the freshly evaluated number.
    const fallback = cachedResultMatches ? cell.display || undefined : undefined
    let numFmt = cell.numFmt || cell.style?.numFmt
    if (!numFmt && cell.arrayMember) {
      const anchorCell = sheet?.cells[cell.arrayMember]
      numFmt = anchorCell?.numFmt || anchorCell?.style?.numFmt
    }
    return formatScalar(value, numFmt, fallback)
  }, [calcEngine, sheetById, showFormulas])
  const resolvedValue = useCallback((sheetId: string, address: string): CellScalar | undefined => {
    if (!calcEngine) return undefined
    const value = calcEngine.getValue(sheetId, address)
    return value === null && !sheetById.get(sheetId)?.cells[address] ? undefined : value
  }, [calcEngine, sheetById])
  const staleValue = useCallback((sheetId: string, address: string) => Boolean(calcEngine?.isStale(sheetId, address)), [calcEngine])
  /** Values, display text, colours, and formula evaluation for the data tools. */
  const dataHostFor = useCallback((sheet: SheetData): DataHost => {
    const engine = calcEngine
    const resolveReference = (reference: string, depth = 0): Array<{ value: CellScalar; text: string }> | null => {
      const text = reference.trim().replace(/^=/, '')
      const match = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(text)
      const current = workbookRef.current
      if (!match) {
        if (depth > 3 || !current) return null
        const named = (current.definedNames || []).find((item) => item.name.toLocaleLowerCase() === text.toLocaleLowerCase())
        return named?.ranges?.[0] ? resolveReference(named.ranges[0], depth + 1) : null
      }
      const sheetName = match[1]?.replace(/''/g, "'") ?? match[2]
      const target = sheetName ? current?.sheets.find((item) => item.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()) : sheet
      if (!target) return null
      const top = Number(match[4]) - 1
      const left = columnIndex(match[3])
      const bottom = match[6] ? Number(match[6]) - 1 : top
      const right = match[5] ? columnIndex(match[5]) : left
      const output: Array<{ value: CellScalar; text: string }> = []
      for (let row = Math.min(top, bottom); row <= Math.max(top, bottom) && output.length < 10_000; row += 1) {
        for (let col = Math.min(left, right); col <= Math.max(left, right); col += 1) {
          const address = addressOf({ row, col })
          const value = engine ? engine.getValue(target.id, address) : target.cells[address]?.value ?? null
          output.push({ value: value ?? null, text: displayValue(target.id, address) })
        }
      }
      return output
    }
    return createSheetHost(sheet, {
      valueAt: (row, col) => (engine ? engine.getValue(sheet.id, addressOf({ row, col })) : sheet.cells[addressOf({ row, col })]?.value ?? null),
      displayAt: (row, col) => displayValue(sheet.id, addressOf({ row, col })),
      cellAt: (row, col) => sheet.cells[addressOf({ row, col })],
      fillColorAt: (row, col) => {
        const fill = sheet.cells[addressOf({ row, col })]?.style?.fill
        if (!fill || String(fill.pattern || '').toLocaleLowerCase() === 'none' || (!fill.fgColor && !fill.color)) return null
        return cssColor(fill.fgColor || fill.color) || null
      },
      fontColorAt: (row, col) => cssColor(sheet.cells[addressOf({ row, col })]?.style?.font?.color) || null,
      isDateAt: (row, col) => {
        const cell = sheet.cells[addressOf({ row, col })]
        return isDateTimeFormat(cell?.numFmt || cell?.style?.numFmt)
      },
      evaluate: (formula, row, col, override) => (engine ? engine.evaluateAt(sheet.id, formula, row, col, override) : null),
      resolveReference: (reference) => resolveReference(reference),
      shiftFormula: shiftFormulaReferences,
      date1904: Boolean(workbookRef.current?.metadata?.date1904),
    })
  }, [calcEngine, displayValue])
  const listOptionsFor = useCallback((validation: Record<string, unknown> | undefined, address: string) => {
    const sheet = workbookRef.current?.sheets.find((item) => item.id === workbookRef.current?.activeSheetId)
    const coord = coordOf(address)
    if (!sheet || !validation) return []
    try {
      return listOptionsForValidation(validation, dataHostFor(sheet), coord || undefined)
    } catch {
      return validationListOptions(validation)
    }
  }, [dataHostFor])

  const chartAccessor = useMemo<ChartWorkbookAccessor | null>(() => (
    workbook && calcEngine
      ? workbookChartAccessor(
        workbook,
        (sheetId, address) => { const value = calcEngine.getValue(sheetId, address); return value === null ? undefined : value },
        (sheetId, address) => displayValue(sheetId, address),
      )
      : null
  ), [calcEngine, displayValue, workbook])
  const chartPalette = useMemo(() => workbookChartPalette(workbook), [workbook])
  const resolveChart = useCallback((chart: Parameters<typeof resolveChartData>[0]) => (
    chartAccessor ? resolveChartData(chart, chartAccessor, { palette: chartPalette }) : { categories: [], series: [] }
  ), [chartAccessor, chartPalette])

  const activeSheetForFormats = workbook?.sheets.find((sheet) => sheet.id === workbook.activeSheetId)
  const conditionalFormats = useCallback((visible: OverlayBounds) => {
    const sheet = activeSheetForFormats
    if (!calcEngine || !sheet?.conditionalFormattings?.length) return null
    const now = new Date()
    const today = (Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - Date.UTC(1899, 11, 30)) / 86_400_000
    try {
      return computeConditionalFormats(sheet.conditionalFormattings, {
        sheetId: sheet.id,
        valueAt: (row, col) => {
          const value = calcEngine.getValue(sheet.id, addressOf({ row, col }))
          return value === undefined ? null : value
        },
        evaluate: (formula, row, col) => {
          const value = calcEngine.evaluateAt(sheet.id, formula, row, col)
          return value === null || value === undefined ? 0 : value
        },
        today,
        cssColor: (color, fallback) => cssColor(color, fallback),
        rowCount: sheet.rowCount,
        colCount: sheet.colCount,
      }, visible)
    } catch {
      return null
    }
  // Recalculate whenever any workbook value may have changed.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeSheetForFormats, calcEngine, workbook])

  const resetHistory = useCallback(() => {
    historyRef.current = []
    futureRef.current = []
    setHistoryTick((value) => value + 1)
  }, [])

  const mutateWorkbook = useCallback((mutator: (next: WorkbookModel) => void) => {
    const current = workbookRef.current
    if (!current) return
    let [next, patches, inversePatches] = produceWithPatches(current, (draft) => { mutator(draft as WorkbookModel) })
    if (next === current) return
    // Sheet protection: every edit path goes through here, so this is the one gate.
    const violation = protectionViolation(current, next, patches)
    if (violation) { setToast(violation); return }
    const headerTouches = tableHeaderTouches(next, patches)
    if (headerTouches.size) {
      // Header edits rename table columns (and their structured references) in the same step.
      const [synced, syncPatches, syncInverse] = produceWithPatches(next, (draft) => {
        for (const [sheetId, addresses] of headerTouches) {
          const sheet = draft.sheets.find((item) => item.id === sheetId)
          if (sheet) syncTableHeaders(draft as WorkbookModel, sheet as SheetData, addresses.has('*') ? undefined : addresses)
        }
      })
      if (synced !== next) {
        next = synced
        patches = [...patches, ...syncPatches]
        inversePatches = [...syncInverse, ...inversePatches]
      }
    }
    historyRef.current.push({ patches, inversePatches })
    if (historyRef.current.length > MAX_HISTORY) historyRef.current.shift()
    futureRef.current = []
    recordEngineHint(current, next, patches)
    workbookRef.current = next
    setWorkbook(next)
    // Stamp the ref alongside the state: a save triggered in the same tick as the mutation
    // (Ctrl+S straight after typing) reads the ref, and the state effect has not run yet.
    dirtyRef.current = true
    setDirty(true)
    setHistoryTick((value) => value + 1)
  }, [recordEngineHint])

  const resizeColumn = useCallback((column: number, width: number) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId) || next.sheets[0]
      if (!sheet) return
      const key = String(column + 1)
      sheet.colWidths[key] = width
      const properties = sheet.columnProperties?.[key]
      if (properties?.bestFit === true) sheet.columnProperties![key] = { ...properties, bestFit: false }
    })
  }, [mutateWorkbook])

  const resizeRow = useCallback((row: number, height: number) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId) || next.sheets[0]
      if (!sheet) return
      sheet.rowHeights[String(row + 1)] = height
    })
  }, [mutateWorkbook])

  const undo = useCallback(() => {
    const current = workbookRef.current
    const entry = historyRef.current.pop()
    if (!current || !entry) return
    futureRef.current.push(entry)
    const previous = applyPatches(current, entry.inversePatches)
    recordEngineHint(current, previous, entry.inversePatches)
    workbookRef.current = previous
    setWorkbook(previous)
    setDirty(true)
    setEditing(null)
    setHistoryTick((value) => value + 1)
  }, [recordEngineHint])

  const redo = useCallback(() => {
    const current = workbookRef.current
    const entry = futureRef.current.pop()
    if (!current || !entry) return
    historyRef.current.push(entry)
    const next = applyPatches(current, entry.patches)
    recordEngineHint(current, next, entry.patches)
    workbookRef.current = next
    setWorkbook(next)
    setDirty(true)
    setEditing(null)
    setHistoryTick((value) => value + 1)
  }, [recordEngineHint])

  const canDiscard = useCallback(async () => {
    flushPendingEditsRef.current()
    return !dirtyRef.current || (await confirmRef.current('Discard the unsaved changes in this spreadsheet?'))
  }, [])

  const askText = useCallback((request: TextPromptRequest) => new Promise<string | null>((resolve) => {
    promptIdRef.current += 1
    setTextPrompt({ ...request, id: promptIdRef.current, resolve })
  }), [])

  const applyPayload = useCallback((payload: WorkbookPayload) => {
    const normalized = normalizeWorkbook(payload.workbook, payload.name)
    applyWorkbookThemeColors(normalized.metadata?.themeColors)
    resetEngine(normalized)
    workbookRef.current = normalized
    setWorkbook(normalized)
    setDocumentFile({
      documentId: payload.documentId,
      path: payload.path,
      name: payload.name,
      sourceFormat: payload.sourceFormat,
      requiresSaveAs: payload.requiresSaveAs,
      warnings: payload.warnings || [],
      stats: payload.stats,
    })
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
    setEditing(null)
    setDirty(false)
    setWarningOpen(false)
    resetHistory()
    if (payload.path) {
      setRecent((current) => rememberRecent(current, {
        path: payload.path!, name: payload.name, format: payload.sourceFormat, openedAt: Date.now(),
      }))
    }
    if (payload.warnings?.length) setToast(`Opened with ${payload.warnings.length} compatibility ${payload.warnings.length === 1 ? 'note' : 'notes'}`)
  }, [resetEngine, resetHistory])

  const newWorkbook = useCallback(async () => {
    if (!(await canDiscard())) return
    setBusy('Creating a clean spreadsheet…')
    try {
      const { documentId } = await window.simpleCalc.createWorkbook()
      const next = blankWorkbook()
      applyWorkbookThemeColors(undefined)
      resetEngine(next)
      workbookRef.current = next
      setWorkbook(next)
      setDocumentFile({ documentId, path: null, name: 'Untitled.xlsx', sourceFormat: 'xlsx', requiresSaveAs: true, warnings: [] })
      setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
      setEditing(null)
      setDirty(false)
      resetHistory()
    } catch (error) {
      setToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }, [canDiscard, resetEngine, resetHistory])

  const openWorkbook = useCallback(async () => {
    if (!(await canDiscard())) return
    setBusy('Opening spreadsheet…')
    try {
      const payload = await window.simpleCalc.openWorkbook()
      if (payload) applyPayload(payload)
    } catch (error) {
      setToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }, [applyPayload, canDiscard])

  const openPath = useCallback(async (filePath: string) => {
    if (!(await canDiscard())) return
    setBusy('Opening spreadsheet…')
    try {
      applyPayload(await window.simpleCalc.openPath(filePath))
    } catch (error) {
      setRecent((current) => current.filter((item) => item.path !== filePath))
      setToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }, [applyPayload, canDiscard])

  const openPathRef = useRef(openPath)
  openPathRef.current = openPath

  const handleDrop = useCallback(async (event: DragEvent<HTMLElement>) => {
    event.preventDefault()
    const file = event.dataTransfer.files[0]
    if (!file || !(await canDiscard())) return
    setBusy(`Opening ${file.name}…`)
    try {
      applyPayload(await window.simpleCalc.openBytes(file.name, await file.arrayBuffer()))
    } catch (error) {
      setToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }, [applyPayload, canDiscard])

  const saveWorkbook = useCallback(async (saveAs = false, requestedFormat?: string) => {
    flushPendingEditsRef.current()
    const currentWorkbook = workbookRef.current
    const currentDocument = documentRef.current
    if (!currentWorkbook || !currentDocument) return
    let format = requestedFormat || currentDocument.sourceFormat || 'xlsx'
    if (!dirtyRef.current && !saveAs && currentDocument.path && format === currentDocument.sourceFormat) {
      setToast('No changes to save')
      return
    }
    if (dirtyRef.current && !['xlsx', 'xls', 'ods', 'csv', 'tsv'].includes(format)) {
      if (!(await confirmRef.current(`Edited ${format.toUpperCase()} files cannot be saved in their original format yet. Save your edits as an XLSX copy? The original file will stay unchanged.`))) return
      format = 'xlsx'
      saveAs = true
    }
    if ((format === 'csv' || format === 'tsv') && currentWorkbook.sheets.length > 1 && !(await confirmRef.current(`${format.toUpperCase()} keeps only the active sheet and cannot store formatting, merges, or multiple sheets. Continue?`))) return
    if (format === 'xlsx' && currentDocument.sourceFormat !== 'xlsx' && currentDocument.warnings.length && !(await confirmRef.current(`Convert this ${currentDocument.sourceFormat.toUpperCase()} workbook to XLSX? Values, formulas, and common formatting will be carried across, but unsupported native features may not have a cross-format equivalent and can be simplified or omitted.`))) return
    if (format === 'xlsx' && currentDocument.sourceFormat === 'xlsx' && dirtyRef.current && currentDocument.warnings.length && !(await confirmRef.current('This workbook contains features listed in the compatibility note that may be simplified after editing. Save the edited workbook anyway?'))) return
    if (format === 'xls' && dirtyRef.current && currentDocument.warnings.some((warning) => /contains VBA|truncated|dropped|could not/i.test(warning)) && !(await confirmRef.current('This workbook contains unsupported features described in the compatibility note. Save the edited XLS workbook? An original backup will be kept before overwriting it.'))) return
    setBusy(`Saving ${format.toUpperCase()} workbook…`)
    try {
      const result = await window.simpleCalc.saveWorkbook({
        documentId: currentDocument.documentId,
        workbook: toFileWorkbook(calculatedWorkbook(currentWorkbook)),
        saveAs: saveAs || !currentDocument.path || format !== currentDocument.sourceFormat,
        format,
        suggestedName: currentDocument.name,
        sourceUnmodified: !dirtyRef.current,
      })
      if (!result) return
      if (documentRef.current?.documentId !== currentDocument.documentId) return
      const savedLatestRevision = workbookRef.current === currentWorkbook
      setDocumentFile((current) => current ? {
        ...current,
        path: result.path,
        name: result.name,
        sourceFormat: result.format,
        requiresSaveAs: false,
        warnings: result.format === 'xlsx' ? [] : current.warnings,
        backupPath: result.backupPath,
      } : current)
      setWorkbook((current) => current ? { ...current, name: result.name } : current)
      if (savedLatestRevision) setDirty(false)
      setRecent((current) => rememberRecent(current, { path: result.path, name: result.name, format: result.format, openedAt: Date.now() }))
      setToast(savedLatestRevision ? `Saved ${result.name}${result.backupPath ? ' · Original backup kept' : ''}` : `Saved ${result.name} · Newer edits still need saving`)
    } catch (error) {
      setToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }, [calculatedWorkbook])

  const openPrintDialog = useCallback(() => {
    flushPendingEditsRef.current()
    if (!workbookRef.current || !documentRef.current) {
      setToast('Open a spreadsheet before printing')
      return
    }
    setPrintOpen(true)
  }, [])

  const createPrintPayload = useCallback((options: SpreadsheetPrintOptions) => {
    flushPendingEditsRef.current()
    const currentWorkbook = workbookRef.current
    const currentDocument = documentRef.current
    if (!currentWorkbook || !currentDocument) throw new Error('There is no spreadsheet to print.')
    let printableWorkbook = calculatedWorkbook(currentWorkbook)
    const normalFont = currentWorkbook.metadata?.normalFont as { name?: string; size?: number; bold?: boolean; italic?: boolean } | undefined
    if (normalFont?.name && Number(normalFont.size) > 0) {
      const context = document.createElement('canvas').getContext('2d')
      if (context) {
        context.font = `${normalFont.italic ? 'italic ' : ''}${normalFont.bold ? 'bold ' : ''}${Number(normalFont.size) * 4 / 3}px ${JSON.stringify(normalFont.name)}, Arial, sans-serif`
        const printDigitWidth = Math.max(...'0123456789'.split('').map(digit => context.measureText(digit).width))
        printableWorkbook = { ...printableWorkbook, sheets: printableWorkbook.sheets.map(sheet => ({ ...sheet, properties: { ...sheet.properties, printDigitWidth } })) }
      }
    }
    const displayValues: Record<string, Record<string, string>> = {}
    const displayParts: SpreadsheetDisplayParts = {}
    const printableSheetIds = new Set(options.scope === 'workbook'
      ? printableWorkbook.sheets.filter((sheet) => sheet.state !== 'hidden' && sheet.state !== 'veryHidden').map((sheet) => sheet.id)
      : [printableWorkbook.activeSheetId])
    const values = (sheetId: string, address: string) => displayValues[sheetId]?.[address] ?? ''
    for (const sheet of printableWorkbook.sheets) {
      if (!printableSheetIds.has(sheet.id)) continue
      const values: Record<string, string> = {}
      const parts: SpreadsheetDisplayParts[string] = {}
      for (const [address, cell] of Object.entries(sheet.cells)) {
        if (showFormulas && cell.formula) values[address] = `=${cell.formula}`
        else {
          const value = cell.formula ? cell.result : cell.value
          values[address] = formatScalar(value, cell.numFmt || cell.style?.numFmt, cell.display || undefined)
          const accounting = typeof value === 'number' ? accountingDisplayParts(values[address], cell.numFmt || cell.style?.numFmt) : null
          parts[address] = { type: typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'text', ...(accounting ? { accounting } : {}) }
        }
      }
      displayValues[sheet.id] = values
      displayParts[sheet.id] = parts
    }
    const selected = selectionBounds(selectionRef.current)
    let charts: ReturnType<typeof buildPrintChartPayload> | undefined
    if (printableWorkbook.sheets.some((sheet) => sheet.charts?.length) && calcEngine) {
      const accessor = workbookChartAccessor(printableWorkbook, (sheetId, address) => { const value = calcEngine.getValue(sheetId, address); return value === null ? undefined : value }, (sheetId, address) => values(sheetId, address))
      const palette = workbookChartPalette(printableWorkbook)
      charts = buildPrintChartPayload(printableWorkbook, (chart) => resolveChartData(chart, accessor, { palette }), { sheetIds: [...printableSheetIds] })
    }
    // Pictures print in their anchored place, like charts (PNG/JPEG/GIF only).
    const images: Record<string, Array<{ src: string; from: ChartAnchor['from']; to: ChartAnchor['to'] }>> = {}
    for (const sheet of printableWorkbook.sheets) {
      if (!printableSheetIds.has(sheet.id) || !sheet.images?.length) continue
      const list = sheet.images
        .filter((image) => image.anchor?.from && image.anchor.to && /^data:image\/(png|jpe?g|gif);base64,/i.test(image.src))
        .map((image) => ({ src: image.src, from: image.anchor.from, to: image.anchor.to }))
      if (list.length) images[sheet.id] = list
    }
    return {
      documentId: currentDocument.documentId,
      name: currentDocument.name,
      workbook: printableWorkbook,
      displayValues,
      displayParts,
      selection: selected,
      options,
      ...(charts ? { charts } : {}),
      ...(Object.keys(images).length ? { images } : {}),
    }
  }, [calcEngine, calculatedWorkbook, showFormulas])

  const renderPrintPreview = useCallback(async (options: SpreadsheetPrintOptions) => {
    return window.simpleCalc.renderPrintPreview(createPrintPayload(options))
  }, [createPrintPayload])

  const printWorkbook = useCallback(async (options: SpreadsheetPrintOptions) => {
    return window.simpleCalc.printWorkbook(createPrintPayload(options))
  }, [createPrintPayload])

  const exportWorkbook = useCallback(async (format: SpreadsheetExportFormat, options: SpreadsheetPrintOptions) => {
    const payload = createPrintPayload(options)
    const currentDocument = documentRef.current
    const currentWorkbook = workbookRef.current
    if (!currentDocument || !currentWorkbook) throw new Error('There is no spreadsheet to export.')
    if ((format === 'csv' || format === 'tsv') && currentWorkbook.sheets.length > 1 && !(await confirmRef.current(`${format.toUpperCase()} exports only the active sheet. Other sheets, formatting, merges, formulas, charts, and images will not be included. Continue?`))) return false
    if (format === 'xlsx' && currentDocument.sourceFormat !== 'xlsx' && currentDocument.warnings.length && !(await confirmRef.current(`Export this ${currentDocument.sourceFormat.toUpperCase()} workbook as XLSX? Values, formulas, and common formatting will be carried across, but unsupported native features may be simplified or omitted.`))) return false
    if (format === 'xlsx' && currentDocument.sourceFormat === 'xlsx' && dirtyRef.current && currentDocument.warnings.length && !(await confirmRef.current('This workbook contains features listed in the compatibility note that may be simplified in an edited XLSX export. Continue?'))) return false
    setBusy(`Exporting ${format.toUpperCase()}…`)
    try {
      const result = await window.simpleCalc.exportWorkbook({
        ...payload,
        workbook: toFileWorkbook(payload.workbook),
        format,
        suggestedName: currentDocument.name,
        sourceUnmodified: !dirtyRef.current,
      })
      if (!result) return false
      setToast(`Exported ${result.name}`)
      return true
    } catch (error) {
      setToast(errorMessage(error))
      throw error
    } finally {
      setBusy('')
    }
  }, [createPrintPayload])

  const closeWindow = useCallback(async () => {
    flushPendingEditsRef.current()
    if (dirtyRef.current && !(await confirmRef.current('Close simple_calc and discard unsaved changes?'))) return
    window.simpleCalc.close()
  }, [])

  useEffect(() => window.simpleCalc.onCloseRequested(closeWindow), [closeWindow])
  useEffect(() => {
    if (!import.meta.env.DEV) return
    // Development-only automation hook used by the QA scripts.
    ;(window as unknown as { __calcQA?: unknown }).__calcQA = {
      openPath: (filePath: string) => { dirtyRef.current = false; return openPathRef.current(filePath) },
      workbook: () => workbookRef.current,
      value: (sheetId: string, address: string) => engineRef.current?.getValue(sheetId, address),
    }
  }, [])
  useEffect(() => window.simpleCalc.onOpenExternal((filePath) => { void openPath(filePath) }), [openPath])

  const commitCell = useCallback((address: string, draft: string) => {
    if (!workbookRef.current) return
    const sheetId = workbookRef.current.activeSheetId
    const sheet = workbookRef.current.sheets.find((item) => item.id === sheetId)
    if (!sheet || rawCellValue(sheet.cells[address]) === draft) { setEditing(null); return }
    const target = coordOf(address)
    if (target && sheet.pivots?.some((pivot) => pivot.extent && target.row >= pivot.anchor.row && target.row < pivot.anchor.row + pivot.extent.rows && target.col >= pivot.anchor.col && target.col < pivot.anchor.col + pivot.extent.cols)) {
      setEditing(null)
      setToast("Pivot table cells can't be edited. Change the pivot table in its editor, or refresh it.")
      return
    }
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === next.activeSheetId)!
      const previousCell = target.cells[address]
      const cell = parseDraft(draft, previousCell)
      if (cell.formula && !previousCell?.formula && !(cell.numFmt || cell.style?.numFmt)) {
        const inferred = inferFormulaNumberFormat(cell.formula, (reference, sheetName) => {
          const source = sheetName ? next.sheets.find((item) => item.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()) : target
          const referenced = source?.cells[reference]
          return referenced?.numFmt || referenced?.style?.numFmt
        })
        if (inferred) cell.numFmt = inferred
      }
      if (hasCellContent(cell)) target.cells[address] = cell
      else delete target.cells[address]
      const coord = coordOf(address)
      if (coord && target.tables?.length) applyTableEntry(next, target, address, coord)
      if (coord) {
        target.rowCount = Math.max(target.rowCount, coord.row + 1)
        target.colCount = Math.max(target.colCount, coord.col + 1)
        // Excel widens a default-width column to fit a newly entered date/currency value.
        const columnKey = String(coord.col + 1)
        const format = cell.numFmt || cell.style?.numFmt
        const autoSized = target.colWidths[columnKey] === undefined || target.columnProperties?.[columnKey]?.bestFit === true
        if (typeof cell.value === 'number' && format && format !== 'General' && !cell.formula && autoSized && !Number(target.properties?.defaultColWidth)) {
          const font = cell.style?.font
          const size = (Number(font?.size) || 11) * (4 / 3)
          const text = formatScalar(cell.value, format)
          const needed = measureTextWidth(text, `${font?.bold ? '700' : '400'} ${size}px ${font?.name ? `"${font.name}", ` : ''}Calibri, Aptos, sans-serif`) + 10
          const current = target.colWidths[columnKey] === undefined ? IMPORTED_COL_WIDTH : columnPixelWidth(target.colWidths[columnKey], 1)
          if (needed > current) {
            target.colWidths[columnKey] = modelColumnWidth(Math.min(needed, MAX_COLUMN_PIXELS), 1)
            target.columnProperties = { ...(target.columnProperties || {}), [columnKey]: { ...(target.columnProperties?.[columnKey] || {}), bestFit: true } }
          }
        }
      }
    })
    setEditing(null)
  }, [mutateWorkbook])

  // Saving, closing or replacing the workbook has to see the text the user is still typing.
  flushPendingEditsRef.current = () => {
    const pendingFormulaBar = formulaBarDirtyRef.current
    if (pendingFormulaBar) {
      formulaBarDirtyRef.current = null
      if (pendingFormulaBar.sheetId === workbookRef.current?.activeSheetId) {
        commitCell(pendingFormulaBar.address, pendingFormulaBar.draft)
      }
    }
    const pendingEditor = editingRef.current
    if (pendingEditor) {
      editingRef.current = null
      commitCell(pendingEditor.address, pendingEditor.draft)
    }
  }

  useEffect(() => {
    const pending = formulaBarDirtyRef.current
    if (pending) {
      if (pending.sheetId === activeSheet?.id && pending.address === activeAddress) return
      formulaBarDirtyRef.current = null
      if (pending.sheetId === activeSheet?.id) commitCell(pending.address, pending.draft)
    }
    setFormulaDraft(rawCellValue(activeCell))
  }, [activeSheet?.id, activeAddress, activeCell, commitCell])

  useEffect(() => {
    registerRecoverySave(async () => {
      const currentWorkbook = workbookRef.current
      const currentDocument = documentRef.current
      if (!currentWorkbook || !currentDocument) throw new Error('No workbook is open to recover.')
      return window.simpleCalc.saveWorkbook({
        documentId: currentDocument.documentId,
        workbook: toFileWorkbook(calculatedWorkbook(currentWorkbook)),
        saveAs: true,
        format: 'xlsx',
        suggestedName: currentDocument.name,
        sourceUnmodified: false,
      })
    })
  }, [calculatedWorkbook])

  const setCellValue = useCallback((address: string, value: CellData['value']) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      const cell = { ...(sheet.cells[address] || {}), value }
      delete cell.formula
      delete cell.formulaType
      delete cell.formulaRange
      delete cell.dynamicFormula
      delete cell.result
      delete cell.resultType
      delete cell.display
      sheet.cells[address] = cell
    })
  }, [mutateWorkbook])

  const openHyperlink = useCallback((target: string) => {
    void window.simpleCalc.openExternal(target).catch((error) => setToast(errorMessage(error)))
  }, [])

  const beginEdit = useCallback((initialDraft?: string) => {
    const current = workbookRef.current
    const sheet = current?.sheets.find((item) => item.id === current.activeSheetId) || current?.sheets[0]
    if (!sheet) return
    const address = mergeMasterAddress(sheet, addressOf(selectionRef.current.focus))
    pointRef.current = null
    setEditing({ address, draft: initialDraft === undefined ? rawCellValue(sheet.cells[address]) : initialDraft, mode: initialDraft === undefined ? 'edit' : 'enter' })
    setAssistDismissed(false)
  }, [])

  const handleGridSelection = useCallback((next: Selection) => {
    setSelection(next)
    setEditing(null)
  }, [])

  const closeContextMenu = useCallback((restoreGridFocus = false) => {
    setContextMenu(null)
    if (!restoreGridFocus) return
    window.requestAnimationFrame(() => {
      if (document.querySelector('[aria-modal="true"]')) return
      document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })
    })
  }, [])

  const openGridContextMenu = useCallback((target: GridContextTarget, position: { x: number; y: number }) => {
    const current = workbookRef.current
    const sheet = current?.sheets.find((item) => item.id === current.activeSheetId)
    if (!sheet) return
    const bounds = selectionBounds(selectionRef.current)
    const lastRow = Math.max(0, sheet.rowCount - 1)
    const lastColumn = Math.max(0, sheet.colCount - 1)
    if (target.kind === 'cell') {
      const inside = target.coord.row >= bounds.top && target.coord.row <= bounds.bottom && target.coord.col >= bounds.left && target.coord.col <= bounds.right
      if (!inside) setSelection({ anchor: target.coord, focus: target.coord })
    } else if (target.kind === 'row-header') {
      const alreadyWholeRows = bounds.left === 0 && bounds.right === lastColumn
      if (!alreadyWholeRows || target.coord.row < bounds.top || target.coord.row > bounds.bottom) {
        setSelection({ anchor: { row: target.coord.row, col: lastColumn }, focus: { row: target.coord.row, col: 0 } })
      }
    } else if (target.kind === 'column-header') {
      const alreadyWholeColumns = bounds.top === 0 && bounds.bottom === lastRow
      if (!alreadyWholeColumns || target.coord.col < bounds.left || target.coord.col > bounds.right) {
        setSelection({ anchor: { row: lastRow, col: target.coord.col }, focus: { row: 0, col: target.coord.col } })
      }
    } else {
      setSelection({ anchor: { row: lastRow, col: lastColumn }, focus: { row: 0, col: 0 } })
    }
    setEditing(null)
    setContextMenu({ kind: target.kind, x: position.x, y: position.y })
  }, [])

  const clearSelection = useCallback(() => {
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large to clear at once.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      addresses.forEach((address) => {
        const cell = sheet.cells[address]
        if (!cell) return
        const cleared = parseDraft('', cell)
        if (hasCellContent(cleared)) sheet.cells[address] = cleared
        else delete sheet.cells[address]
      })
    })
  }, [mutateWorkbook, selection])

  const applyStyle = useCallback((update: (style: CellStyle) => void) => {
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large to format at once.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      addresses.forEach((address) => {
        const cell = sheet.cells[address] ||= {}
        const style = cell.style ||= {}
        update(style)
      })
    })
  }, [mutateWorkbook, selection])

  const setNumberFormat = useCallback((numFmt: string) => {
    const addresses = rangeAddresses(selection)
    if (!addresses.length) return
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      addresses.forEach((address) => { sheet.cells[address] = { ...(sheet.cells[address] || {}), numFmt } })
    })
  }, [mutateWorkbook, selection])

  const adjustDecimals = useCallback((delta: 1 | -1) => {
    const activeValue = activeSheet && calcEngine ? calcEngine.getValue(activeSheet.id, activeAddress) : undefined
    const nextFormat = adjustFormatDecimals(activeCell?.numFmt || activeCell?.style?.numFmt, delta, activeValue)
    if (!nextFormat) { setToast('Decimal controls are unavailable for date and time formats.'); return }
    setNumberFormat(nextFormat)
  }, [activeAddress, activeCell, activeSheet, calcEngine, setNumberFormat])

  const clearFormatting = useCallback(() => {
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large to format at once.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      addresses.forEach((address) => {
        const cell = sheet.cells[address]
        if (!cell) return
        const cleared = { ...cell }
        delete cleared.style
        delete cleared.numFmt
        if (hasCellContent(cleared)) sheet.cells[address] = cleared
        else delete sheet.cells[address]
      })
    })
  }, [mutateWorkbook, selection])

  const setBorderPreset = useCallback((preset: 'all' | 'outer' | 'bottom' | 'clear') => {
    const bounds = selectionBounds(selection)
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large to format at once.'); return }
    const side: CellBorderSide = { style: 'thin', color: { argb: 'FFD5D9D2' } }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      addresses.forEach((address) => {
        const coord = coordOf(address)!
        const cell = sheet.cells[address] ||= {}
        const style = cell.style ||= {}
        if (preset === 'clear') delete style.border
        else {
          const border = style.border ||= {}
          if (preset === 'all') border.top = border.bottom = border.left = border.right = side
          if (preset === 'bottom') border.bottom = side
          if (preset === 'outer') {
            if (coord.row === bounds.top) border.top = side
            if (coord.row === bounds.bottom) border.bottom = side
            if (coord.col === bounds.left) border.left = side
            if (coord.col === bounds.right) border.right = side
          }
        }
      })
    })
  }, [mutateWorkbook, selection])

  /** Apply a Format Cells change to the selection with Excel's edge/inside-border semantics. */
  const applyFormatChangeToSelection = useCallback((change: FormatCellsChange) => {
    const bounds = selectionBounds(selectionRef.current)
    const total = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
    if (total > 100_000) { setToast('That selection is too large to format at once.'); return }
    const { merge, ...cellChange } = change
    const hasCellChange = Object.values(cellChange).some((value) => value !== undefined)
    if (hasCellChange) {
      mutateWorkbook((next) => {
        const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
        for (let row = bounds.top; row <= bounds.bottom; row += 1) {
          for (let col = bounds.left; col <= bounds.right; col += 1) {
            const address = addressOf({ row, col })
            const updated = applyFormatChangeToCell(sheet.cells[address], cellChange, cellPosition(row, col, bounds))
            if (hasCellContent(updated)) sheet.cells[address] = updated
            else delete sheet.cells[address]
          }
        }
        // A changed outer edge also clears the facing edge of the neighbouring cell.
        const borders: FormatBorderChange | undefined = cellChange.borders
        if (borders) {
          const edge = (outer: 'top' | 'bottom' | 'left' | 'right') => borders[outer] !== undefined || borders.outline !== undefined
          const touch = (row: number, col: number, direction: 'above' | 'below' | 'left' | 'right') => {
            if (row < 0 || col < 0) return
            const address = addressOf({ row, col })
            const cell = sheet.cells[address]
            if (!cell?.style?.border) return
            const style = applyNeighborBorderChange(cell.style, borders, direction)
            sheet.cells[address] = { ...cell, style }
          }
          if (edge('top')) for (let col = bounds.left; col <= bounds.right; col += 1) touch(bounds.top - 1, col, 'above')
          if (edge('bottom')) for (let col = bounds.left; col <= bounds.right; col += 1) touch(bounds.bottom + 1, col, 'below')
          if (edge('left')) for (let row = bounds.top; row <= bounds.bottom; row += 1) touch(row, bounds.left - 1, 'left')
          if (edge('right')) for (let row = bounds.top; row <= bounds.bottom; row += 1) touch(row, bounds.right + 1, 'right')
        }
      })
    }
    if (merge === true) mergeSelectionRef.current()
    else if (merge === false) unmergeSelectionRef.current()
  }, [mutateWorkbook])

  const applyBorderPresetToSelection = useCallback((preset: BorderPreset, side: CellBorderSide) => {
    applyFormatChangeToSelection({ borders: borderPresetChange(preset, side) })
  }, [applyFormatChangeToSelection])

  const applyCellStyleToSelection = useCallback((preset: CellStylePreset) => {
    const addresses = rangeAddresses(selectionRef.current)
    if (!addresses.length) { setToast('That selection is too large to format at once.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      addresses.forEach((address) => {
        const updated = applyCellStylePresetToCell(sheet.cells[address], preset)
        if (hasCellContent(updated)) sheet.cells[address] = updated
        else delete sheet.cells[address]
      })
    })
  }, [mutateWorkbook])

  const [formatPainter, setFormatPainter] = useState<{ cells: Array<Array<Pick<CellData, 'style' | 'numFmt'>>>; sticky: boolean } | null>(null)
  const formatPainterRef = useRef(formatPainter)
  formatPainterRef.current = formatPainter
  const startFormatPainter = useCallback((sticky: boolean) => {
    const sheet = workbookRef.current?.sheets.find((item) => item.id === workbookRef.current?.activeSheetId)
    if (!sheet) return
    const bounds = selectionBounds(selectionRef.current)
    if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > 10_000) { setToast('Select a smaller range to copy its formatting.'); return }
    const cells: Array<Array<Pick<CellData, 'style' | 'numFmt'>>> = []
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      const line: Array<Pick<CellData, 'style' | 'numFmt'>> = []
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        const cell = sheet.cells[addressOf({ row, col })]
        line.push({ style: cell?.style ? structuredClone(cell.style) : undefined, numFmt: cell?.numFmt })
      }
      cells.push(line)
    }
    setFormatPainter({ cells, sticky })
    setToast(sticky ? 'Format painter locked: click cells to paint, Esc to stop' : 'Select cells to paint the copied formatting')
  }, [])
  const paintFormatsOnSelection = useCallback(() => {
    const painter = formatPainterRef.current
    if (!painter) return
    const bounds = selectionBounds(selectionRef.current)
    const height = painter.cells.length
    const width = painter.cells[0]?.length || 1
    if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > 100_000) { setToast('That selection is too large to format at once.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      for (let row = bounds.top; row <= bounds.bottom; row += 1) {
        for (let col = bounds.left; col <= bounds.right; col += 1) {
          const source = painter.cells[(row - bounds.top) % height][(col - bounds.left) % width]
          const address = addressOf({ row, col })
          const cell: CellData = { ...(sheet.cells[address] || {}) }
          if (source.style) cell.style = structuredClone(source.style)
          else delete cell.style
          if (source.numFmt) cell.numFmt = source.numFmt
          else delete cell.numFmt
          if (hasCellContent(cell)) sheet.cells[address] = cell
          else delete sheet.cells[address]
        }
      }
    })
    if (!painter.sticky) setFormatPainter(null)
  }, [mutateWorkbook])
  useEffect(() => {
    if (!formatPainter) return
    const viewport = document.querySelector<HTMLElement>('.sheet-viewport')
    const apply = (event: globalThis.PointerEvent) => {
      if (!(event.target instanceof Element) || !event.target.closest('.sheet-viewport')) return
      window.setTimeout(() => paintFormatsOnSelection(), 0)
    }
    const cancel = (event: globalThis.KeyboardEvent) => { if (event.key === 'Escape') setFormatPainter(null) }
    viewport?.classList.add('is-format-painting')
    window.addEventListener('pointerup', apply)
    window.addEventListener('keydown', cancel)
    return () => {
      viewport?.classList.remove('is-format-painting')
      window.removeEventListener('pointerup', apply)
      window.removeEventListener('keydown', cancel)
    }
  }, [formatPainter, paintFormatsOnSelection])

  const mergeSelectionRef = useRef<() => void>(() => {})
  const unmergeSelectionRef = useRef<() => void>(() => {})

  const unmergeSelection = useCallback(() => {
    if (!activeSheet) return
    const bounds = selectionBounds(selection)
    const matching = (activeSheet.merges || []).filter((range) => {
      const merged = mergeBounds(range)
      return Boolean(merged && merged.bottom >= bounds.top && merged.top <= bounds.bottom && merged.right >= bounds.left && merged.left <= bounds.right)
    })
    if (!matching.length) { setToast('No merged cells are selected.'); return }
    const remove = new Set(matching)
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      sheet.merges = (sheet.merges || []).filter((range) => !remove.has(range))
    })
  }, [activeSheet, mutateWorkbook, selection])

  const mergeSelection = useCallback(async () => {
    if (!activeSheet) return
    const bounds = selectionBounds(selection)
    if (bounds.top === bounds.bottom && bounds.left === bounds.right) { setToast('Select two or more cells to merge.'); return }
    const overlaps = (activeSheet.merges || []).filter((range) => {
      const merged = mergeBounds(range)
      return Boolean(merged && merged.bottom >= bounds.top && merged.top <= bounds.bottom && merged.right >= bounds.left && merged.left <= bounds.right)
    })
    const requested = rangeAddress(bounds)
    if (overlaps.length === 1 && overlaps[0] === requested) { unmergeSelection(); return }
    if (overlaps.length) { setToast('Unmerge the overlapping cells before creating this merge.'); return }
    let extraValues = 0
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        if (row === bounds.top && col === bounds.left) continue
        const cell = activeSheet.cells[addressOf({ row, col })]
        if (cell && (cell.value !== undefined || cell.formula)) extraValues += 1
      }
    }
    if (extraValues && !(await confirmRef.current(`Merge these cells? ${extraValues} non-leading ${extraValues === 1 ? 'value' : 'values'} will be removed.`))) return
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      sheet.merges = [...(sheet.merges || []), requested]
      for (let row = bounds.top; row <= bounds.bottom; row += 1) {
        for (let col = bounds.left; col <= bounds.right; col += 1) {
          if (row === bounds.top && col === bounds.left) continue
          const address = addressOf({ row, col })
          const cell = sheet.cells[address]
          if (!cell) continue
          const kept: CellData = {}
          if (cell.style) kept.style = cell.style
          if (cell.numFmt) kept.numFmt = cell.numFmt
          if (hasCellContent(kept)) sheet.cells[address] = kept
          else delete sheet.cells[address]
        }
      }
    })
    setSelection({ anchor: { row: bounds.top, col: bounds.left }, focus: { row: bounds.bottom, col: bounds.right } })
  }, [activeSheet, mutateWorkbook, selection, unmergeSelection])

  mergeSelectionRef.current = mergeSelection
  unmergeSelectionRef.current = unmergeSelection

  /** Merge each row (across) or each column (down) of the selection separately. */
  const mergeSelectionBy = useCallback((mode: 'across' | 'down' | 'center') => {
    if (!activeSheet) return
    const bounds = selectionBounds(selection)
    if (mode === 'center') {
      mergeSelection()
      applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'center' } })
      return
    }
    const ranges: string[] = []
    if (mode === 'across') {
      if (bounds.left === bounds.right) { setToast('Select two or more columns to merge across.'); return }
      for (let row = bounds.top; row <= bounds.bottom; row += 1) ranges.push(rangeAddress({ top: row, bottom: row, left: bounds.left, right: bounds.right }))
    } else {
      if (bounds.top === bounds.bottom) { setToast('Select two or more rows to merge down.'); return }
      for (let col = bounds.left; col <= bounds.right; col += 1) ranges.push(rangeAddress({ top: bounds.top, bottom: bounds.bottom, left: col, right: col }))
    }
    const overlaps = (activeSheet.merges || []).some((range) => {
      const merged = mergeBounds(range)
      return Boolean(merged && merged.bottom >= bounds.top && merged.top <= bounds.bottom && merged.right >= bounds.left && merged.left <= bounds.right)
    })
    if (overlaps) { setToast('Unmerge the overlapping cells first.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      sheet.merges = [...(sheet.merges || []), ...ranges]
      for (const range of ranges) {
        const merged = mergeBounds(range)!
        for (let row = merged.top; row <= merged.bottom; row += 1) {
          for (let col = merged.left; col <= merged.right; col += 1) {
            if (row === merged.top && col === merged.left) continue
            const address = addressOf({ row, col })
            const cell = sheet.cells[address]
            if (!cell) continue
            const kept: CellData = {}
            if (cell.style) kept.style = cell.style
            if (cell.numFmt) kept.numFmt = cell.numFmt
            if (hasCellContent(kept)) sheet.cells[address] = kept
            else delete sheet.cells[address]
          }
        }
      }
    })
  }, [activeSheet, applyStyle, mergeSelection, mutateWorkbook, selection])

  const applyAlternatingColors = useCallback(() => {
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large to format at once.'); return }
    const bounds = selectionBounds(selection)
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      addresses.forEach((address) => {
        const coord = coordOf(address)!
        const cell = sheet.cells[address] ||= {}
        const style = cell.style ||= {}
        style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: coord.row === bounds.top ? 'FFDCE9E0' : (coord.row - bounds.top) % 2 ? 'FFF4F7F4' : 'FFFFFFFF' } }
        if (coord.row === bounds.top) style.font = { ...(style.font || {}), bold: true, color: { argb: 'FF294936' } }
      })
    })
  }, [mutateWorkbook, selection])

  const insertLink = useCallback(async () => {
    if (!activeSheet) return
    const existing = activeCell?.hyperlink || 'https://'
    const entered = (await askText({ title: 'Insert link', label: 'Link URL or email address', initialValue: existing }))?.trim()
    if (!entered) return
    const target = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(entered) ? `mailto:${entered}` : /^[a-z][a-z0-9+.-]*:/i.test(entered) ? entered : `https://${entered}`
    try {
      const parsed = new URL(target)
      if (!['http:', 'https:', 'mailto:'].includes(parsed.protocol)) throw new Error('Unsupported protocol')
    } catch {
      setToast('Enter a valid web or email link.')
      return
    }
    if (activeCell?.formula) { setToast('A formula cell cannot also contain a direct hyperlink.'); return }
    const currentText = activeCell?.value == null ? '' : String(activeCell.value)
    const text = await askText({ title: 'Insert link', label: 'Text to display', initialValue: currentText || entered })
    if (text === null) return
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      sheet.cells[activeAddress] = { ...(sheet.cells[activeAddress] || {}), value: text || entered, hyperlink: target, hyperlinkTooltip: target }
    })
  }, [activeAddress, activeCell, activeSheet, askText, mutateWorkbook])

  const editAnnotation = useCallback(async (kind: 'note' | 'comment') => {
    const current = noteText(activeCell?.note)
    const value = await askText({
      title: kind === 'note' ? 'Cell note' : 'Offline comment',
      label: kind === 'note' ? 'Note for the selected cell' : 'Comment for the selected cell (saved as an Excel note)',
      initialValue: current,
      multiline: true,
    })
    if (value === null) return
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      const cell = { ...(sheet.cells[activeAddress] || {}) }
      if (!value.trim()) delete cell.note
      else cell.note = kind === 'note' ? value.trim() : { comments: [{ author: 'simple_calc', text: value.trim() }] }
      if (hasCellContent(cell)) sheet.cells[activeAddress] = cell
      else delete sheet.cells[activeAddress]
    })
  }, [activeAddress, activeCell?.note, askText, mutateWorkbook])

  const insertCheckboxes = useCallback(() => {
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large for checkboxes.'); return }
    const bounds = selectionBounds(selection)
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      setValidationForBounds(sheet, bounds, { type: 'list', allowBlank: false, formulae: ['"TRUE,FALSE"'] })
      addresses.forEach((address) => {
        const existing = sheet.cells[address] || {}
        const value = typeof existing.value === 'boolean' ? existing.value : false
        const cell: CellData = { ...existing, value, type: 'checkbox' }
        delete cell.formula
        delete cell.formulaType
        delete cell.formulaRange
        delete cell.dynamicFormula
        delete cell.result
        delete cell.display
        sheet.cells[address] = cell
      })
    })
  }, [mutateWorkbook, selection])

  const insertDropdown = useCallback(async () => {
    const entered = await askText({ title: 'Insert dropdown', label: 'Dropdown options, separated by commas', initialValue: 'Option 1, Option 2' })
    if (entered === null) return
    const options = [...new Set(entered.split(',').map((item) => item.trim()).filter(Boolean))].slice(0, 100)
    if (options.length < 2) { setToast('Enter at least two dropdown options.'); return }
    const literal = `"${options.join(',').replace(/"/g, '""')}"`
    if (literal.length > 255) { setToast('Excel-compatible dropdown lists are limited to 255 characters.'); return }
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large for a dropdown.'); return }
    const bounds = selectionBounds(selection)
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      setValidationForBounds(sheet, bounds, { type: 'list', allowBlank: true, showErrorMessage: true, formulae: [literal] })
      addresses.forEach((address) => {
        const existing = sheet.cells[address] || {}
        const current = existing.value == null ? '' : String(existing.value)
        const cell: CellData = { ...existing, value: options.includes(current) ? current : '', type: 'dropdown' }
        delete cell.formula
        delete cell.formulaType
        delete cell.formulaRange
        delete cell.dynamicFormula
        delete cell.result
        delete cell.display
        sheet.cells[address] = cell
      })
    })
  }, [askText, mutateWorkbook, selection])

  const removeValidation = useCallback(() => {
    const bounds = selectionBounds(selection)
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large to update at once.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      setValidationForBounds(sheet, bounds)
      addresses.forEach((address) => {
        const cell = sheet.cells[address]
        if (!cell || (cell.type !== 'checkbox' && cell.type !== 'dropdown')) return
        const updated = { ...cell }
        delete updated.type
        sheet.cells[address] = updated
      })
    })
  }, [mutateWorkbook, selection])

  /**
   * Excel AutoSum: for a single cell, sum the run of numbers directly above (or else to the
   * left) and open the formula for confirmation; for a range, write totals below each column
   * (into the selection's empty last row when there is one).
   */
  const insertFunction = useCallback((name: string) => {
    const sheet = activeSheet
    if (!sheet) return
    const bounds = selectionBounds(selection)
    const multi = bounds.top !== bounds.bottom || bounds.left !== bounds.right
    const isNumberAt = (row: number, col: number) => {
      if (row < 0 || col < 0) return false
      const value = calcEngine?.getValue(sheet.id, addressOf({ row, col }))
      return typeof value === 'number'
    }
    const isEmptyAt = (row: number, col: number) => {
      const value = calcEngine?.getValue(sheet.id, addressOf({ row, col }))
      return value === null || value === undefined || value === ''
    }
    if (!['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX'].includes(name)) { beginEdit(`=${name}(`); return }
    if (!multi) {
      const { row, col } = selection.focus
      let top = row - 1
      // Skip at most one blank/label cell directly above, as Excel does for headed columns.
      while (top >= 0 && isNumberAt(top, col)) top -= 1
      if (top < row - 1) {
        beginEdit(`=${name}(${rangeAddress({ top: top + 1, bottom: row - 1, left: col, right: col })})`)
        return
      }
      let left = col - 1
      while (left >= 0 && isNumberAt(row, left)) left -= 1
      if (left < col - 1) {
        beginEdit(`=${name}(${rangeAddress({ top: row, bottom: row, left: left + 1, right: col - 1 })})`)
        return
      }
      beginEdit(`=${name}(`)
      return
    }
    const lastRowEmpty = bounds.bottom > bounds.top && Array.from({ length: bounds.right - bounds.left + 1 }, (_, offset) => bounds.left + offset).every((col) => isEmptyAt(bounds.bottom, col))
    const targetRow = lastRowEmpty ? bounds.bottom : Math.min(1_048_575, bounds.bottom + 1)
    const sourceBottom = lastRowEmpty ? bounds.bottom - 1 : bounds.bottom
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === next.activeSheetId)!
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        const address = addressOf({ row: targetRow, col })
        const cell = parseDraft(`=${name}(${rangeAddress({ top: bounds.top, bottom: sourceBottom, left: col, right: col })})`, target.cells[address])
        const sourceFormat = target.cells[addressOf({ row: bounds.top, col })]?.numFmt || target.cells[addressOf({ row: bounds.top, col })]?.style?.numFmt
        if (sourceFormat && !cell.numFmt && name !== 'COUNT' && name !== 'COUNTA') cell.numFmt = sourceFormat
        target.cells[address] = cell
      }
      target.rowCount = Math.max(target.rowCount, targetRow + 1)
    })
    setSelection({ anchor: { row: bounds.top, col: bounds.left }, focus: { row: targetRow, col: bounds.right } })
  }, [activeSheet, beginEdit, calcEngine, mutateWorkbook, selection])

  const applyStructureCommand = useCallback((command: SelectionStructureCommand) => {
    const current = workbookRef.current
    if (!current) return
    const protection = sheetProtection(current.sheets.find((sheet) => sheet.id === current.activeSheetId))
    const needed = command.startsWith('insert-rows') ? 'insertRows' : command.startsWith('insert-columns') ? 'insertColumns' : command === 'delete-rows' ? 'deleteRows' : 'deleteColumns'
    if (protection && !protectionAllows(protection, needed)) { setToast(PROTECTED_MESSAGE); return }
    try {
      const result = applySelectionStructureCommand(current, current.activeSheetId, selection, command)
      const withCharts = transformWorkbookChartsForStructure(result.workbook, current.activeSheetId, result.operation)
      const [next, patches, inversePatches] = produceWithPatches(current, (draft) => { Object.assign(draft, withCharts) })
      historyRef.current.push({ patches, inversePatches })
      if (historyRef.current.length > MAX_HISTORY) historyRef.current.shift()
      futureRef.current = []
      workbookRef.current = next
      setWorkbook(next)
      setSelection(result.selection)
      setEditing(null)
      setDirty(true)
      setHistoryTick((value) => value + 1)
      const amount = result.operation.count
      const noun = result.operation.axis === 'row' ? (amount === 1 ? 'row' : 'rows') : (amount === 1 ? 'column' : 'columns')
      setToast(`${result.operation.kind === 'insert' ? 'Inserted' : 'Deleted'} ${amount} ${noun}`)
    } catch (error) {
      setToast(errorMessage(error))
    }
  }, [selection])

  /** Excel's Insert/Delete cells with a shift direction (partial rows or columns). */
  const applyCellShift = useCallback((direction: 'down' | 'right' | 'up' | 'left') => {
    const current = workbookRef.current
    if (!current) return
    if (sheetProtection(current.sheets.find((sheet) => sheet.id === current.activeSheetId))) { setToast(PROTECTED_MESSAGE); return }
    const bounds = selectionBounds(selection)
    try {
      const result = shiftCells(current, current.activeSheetId, bounds, direction)
      const [next, patches, inversePatches] = produceWithPatches(current, (draft) => { Object.assign(draft, result) })
      historyRef.current.push({ patches, inversePatches })
      if (historyRef.current.length > MAX_HISTORY) historyRef.current.shift()
      futureRef.current = []
      workbookRef.current = next
      setWorkbook(next)
      setEditing(null)
      setDirty(true)
      setHistoryTick((value) => value + 1)
      setToast(direction === 'down' || direction === 'right' ? `Inserted cells, shifted ${direction}` : `Deleted cells, shifted ${direction}`)
    } catch (error) {
      setToast(errorMessage(error))
    }
  }, [selection])

  const openShiftDialog = useCallback((mode: 'insert' | 'delete') => {
    const sheet = activeSheet
    if (!sheet) return
    const bounds = selectionBounds(selection)
    const wholeRows = bounds.left === 0 && bounds.right >= Math.max(0, sheet.colCount - 1)
    const wholeColumns = bounds.top === 0 && bounds.bottom >= Math.max(0, sheet.rowCount - 1)
    // Whole rows or columns go straight through, as in Excel.
    if (wholeRows && !wholeColumns) { applyStructureCommand(mode === 'insert' ? 'insert-rows-above' : 'delete-rows'); return }
    if (wholeColumns && !wholeRows) { applyStructureCommand(mode === 'insert' ? 'insert-columns-left' : 'delete-columns'); return }
    const wide = bounds.right - bounds.left > bounds.bottom - bounds.top
    setShiftDialog({ mode, initial: mode === 'insert' ? (wide ? 'shift-down' : 'shift-right') : (wide ? 'shift-up' : 'shift-left') })
  }, [activeSheet, applyStructureCommand, selection])
  openShiftDialogRef.current = openShiftDialog

  const applyShiftChoice = useCallback((mode: 'insert' | 'delete', choice: ShiftCellsChoice) => {
    setShiftDialog(null)
    if (choice === 'entire-row') applyStructureCommand(mode === 'insert' ? 'insert-rows-above' : 'delete-rows')
    else if (choice === 'entire-column') applyStructureCommand(mode === 'insert' ? 'insert-columns-left' : 'delete-columns')
    else applyCellShift(choice.slice('shift-'.length) as 'down' | 'right' | 'up' | 'left')
    window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true }))
  }, [applyCellShift, applyStructureCommand])

  /** Read a pivot's source (A1 range with its header row, or a table name) from calculated values. */
  const readPivotSource = useCallback((book: WorkbookModel, text: string, fallbackSheetId: string): { source: PivotSource | null; error: string | null } => {
    const engine = calcEngine
    if (!engine) return { source: null, error: 'The workbook is still calculating.' }
    const trimmed = text.trim().replace(/^=/, '')
    let sheet: SheetData | undefined
    let bounds: { top: number; bottom: number; left: number; right: number } | null = null
    for (const candidate of book.sheets) {
      const table = candidate.tables?.find((item) => item.name.toLocaleLowerCase() === trimmed.toLocaleLowerCase())
      const regions = table ? tableRegions(table) : null
      if (regions && regions.header !== null) { sheet = candidate; bounds = { top: regions.header, bottom: regions.dataBottom, left: regions.left, right: regions.right }; break }
    }
    if (!sheet) {
      const match = /^(?:(?:'((?:[^']|'')+)'|([^!'"]+))!)?(\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+)$/.exec(trimmed)
      if (!match) return { source: null, error: 'Enter a range such as Sheet1!A1:E200 or a table name.' }
      const name = match[1] ? match[1].replace(/''/g, "'") : match[2]
      sheet = name ? book.sheets.find((item) => item.name.toLocaleLowerCase() === name.trim().toLocaleLowerCase()) : book.sheets.find((item) => item.id === fallbackSheetId)
      if (!sheet) return { source: null, error: `There is no sheet named "${name}".` }
      bounds = parseTableRef(match[3])
    }
    if (!bounds || bounds.bottom <= bounds.top) return { source: null, error: 'The data range needs a header row and at least one row of data.' }
    if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > 2_000_000) return { source: null, error: 'That data range is too large for a pivot table.' }
    const read = (row: number, col: number) => {
      const address = addressOf({ row, col })
      const value = engine.getValue(sheet!.id, address)
      const cell = sheet!.cells[address]
      const numFmt = cell?.numFmt || cell?.style?.numFmt
      return { value, text: displayValue(sheet!.id, address), ...(numFmt ? { numFmt } : {}), ...(typeof value === 'number' && isDateNumberFormat(numFmt) ? { isDate: true } : {}) }
    }
    const headers = uniqueColumnNames(Array.from({ length: bounds.right - bounds.left + 1 }, (_, offset) => read(bounds!.top, bounds!.left + offset).text.trim()))
    const records: PivotSource['records'] = []
    for (let row = bounds.top + 1; row <= bounds.bottom; row += 1) {
      const record = Array.from({ length: bounds.right - bounds.left + 1 }, (_, offset) => read(row, bounds!.left + offset))
      if (record.some((item) => item.value !== null && item.value !== undefined && item.value !== '')) records.push(record)
    }
    return { source: { headers, records }, error: null }
  }, [calcEngine, displayValue])

  /** Recompute a pivot (optionally with a new definition) and rewrite its block. */
  const refreshPivot = useCallback(async (pivotId: string, definition?: PivotTableModel) => {
    const book = workbookRef.current
    if (!book) return false
    const host = book.sheets.find((sheet) => sheet.pivots?.some((pivot) => pivot.id === pivotId))
    const current = host?.pivots?.find((pivot) => pivot.id === pivotId)
    if (!host || !current) return false
    const model = definition || current
    const { source, error } = readPivotSource(book, model.source, host.id)
    if (!source) { setToast(error || 'The pivot table source could not be read.'); if (definition) mutateWorkbook((next) => { const sheet = next.sheets.find((item) => item.id === host.id); if (sheet?.pivots) sheet.pivots = sheet.pivots.map((pivot) => (pivot.id === pivotId ? { ...definition, extent: current.extent } : pivot)) }); return false }
    const output = computePivot(source, model)
    const conflicts = pivotOverwriteConflicts(host, { ...model, extent: current.extent }, output)
    if (conflicts.length && !(await confirmRef.current(`There's already data in ${conflicts.join(', ')}${conflicts.length >= 5 ? '…' : ''}. Do you want to replace it?`))) return false
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === host.id)
      const target = sheet?.pivots?.find((pivot) => pivot.id === pivotId)
      if (!sheet || !target) return
      const updated: PivotTableModel = { ...model, extent: target.extent }
      writePivotOutput(sheet as SheetData, updated, output)
      autofitPivotColumns(sheet as SheetData, updated, output)
      sheet.pivots = sheet.pivots!.map((pivot) => (pivot.id === pivotId ? updated : pivot))
    })
    return true
  }, [mutateWorkbook, readPivotSource])

  const refreshPivots = useCallback(async (all: boolean) => {
    const book = workbookRef.current
    if (!book) return
    const ids = all
      ? book.sheets.flatMap((sheet) => (sheet.pivots || []).map((pivot) => pivot.id))
      : (() => {
        const sheet = book.sheets.find((item) => item.id === book.activeSheetId)
        const found = sheet?.pivots?.find((pivot) => pivot.extent && selection.focus.row >= pivot.anchor.row && selection.focus.row < pivot.anchor.row + pivot.extent.rows && selection.focus.col >= pivot.anchor.col && selection.focus.col < pivot.anchor.col + pivot.extent.cols)
        return found ? [found.id] : []
      })()
    if (!ids.length) { setToast(all ? 'This workbook has no pivot tables.' : 'Select a cell in a pivot table to refresh it.'); return }
    let done = 0
    for (const id of ids) if (await refreshPivot(id)) done += 1
    if (done) setToast(done === 1 ? 'Refreshed the pivot table' : `Refreshed ${done} pivot tables`)
  }, [refreshPivot, selection.focus.col, selection.focus.row])
  refreshPivotsRef.current = (all) => { void refreshPivots(all) }

  const openCreatePivot = useCallback(() => {
    const sheet = activeSheet
    if (!sheet) return
    const table = tableContaining(sheet, selection.focus.row, selection.focus.col)
    const selected = selectionBounds(selection)
    const region = (selected.top === selected.bottom && selected.left === selected.right ? filterRangeForSelection(sheet, selected, selection.focus) : selected) || selected
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet.name) ? sheet.name : `'${sheet.name.replace(/'/g, "''")}'`
    const range = table ? table.name : `${quoted}!${rangeAddress(region).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`
    setPivotDialog({ range, location: `${quoted}!${columnName(region.right + 2)}${region.top + 1}` })
  }, [activeSheet, selection])

  const createPivot = useCallback((range: string, destination: 'new' | string): string | null => {
    const book = workbookRef.current
    const sheet = activeSheet
    if (!book || !sheet) return 'There is no workbook open.'
    const { source, error } = readPivotSource(book, range, sheet.id)
    if (!source) return error
    let hostId = sheet.id
    let anchor = { row: 0, col: 0 }
    if (destination !== 'new') {
      const match = /^(?:(?:'((?:[^']|'')+)'|([^!'"]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(destination)
      if (!match) return 'Enter a location such as Sheet1!H2.'
      const name = match[1] ? match[1].replace(/''/g, "'") : match[2]
      const host = name ? book.sheets.find((item) => item.name.toLocaleLowerCase() === name.trim().toLocaleLowerCase()) : sheet
      if (!host) return `There is no sheet named "${name}".`
      hostId = host.id
      anchor = coordOf(`${match[3].toUpperCase()}${match[4]}`) || anchor
    }
    const pivot: PivotTableModel = {
      id: `pivot-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`,
      name: nextPivotName(book.sheets),
      source: range,
      anchor,
      rows: [],
      columns: [],
      values: [],
      filters: [],
    }
    const output = computePivot(source, pivot)
    if (destination !== 'new') {
      const host = book.sheets.find((item) => item.id === hostId)!
      if (pivotOverwriteConflicts(host, pivot, output).length) return 'That location already has data. Choose an empty place for the pivot table.'
    }
    mutateWorkbook((next) => {
      let host = next.sheets.find((item) => item.id === hostId)!
      if (destination === 'new') {
        let index = 1
        while (next.sheets.some((item) => item.name.toLocaleLowerCase() === `pivot table ${index}`)) index += 1
        const id = makeId()
        const position = next.sheets.findIndex((item) => item.id === sheet.id)
        const created: SheetData = { id, name: `Pivot Table ${index}`, state: 'visible', rowCount: DEFAULT_ROWS, colCount: DEFAULT_COLS, cells: {}, merges: [], colWidths: {}, rowHeights: {}, frozen: {} }
        next.sheets.splice(position + 1, 0, created)
        host = next.sheets[position + 1]
        hostId = id
      }
      const model = { ...pivot }
      writePivotOutput(host as SheetData, model, output)
      autofitPivotColumns(host as SheetData, model, output)
      host.pivots = [...(host.pivots || []), model]
      next.activeSheetId = hostId
    })
    setSelection({ anchor, focus: anchor })
    setPivotDialog(null)
    setPivotEditorHidden(null)
    return null
  }, [activeSheet, mutateWorkbook, readPivotSource])

  const deletePivot = useCallback((pivotId: string) => {
    mutateWorkbook((next) => {
      for (const sheet of next.sheets) {
        const pivot = sheet.pivots?.find((item) => item.id === pivotId)
        if (!pivot) continue
        if (pivot.extent) {
          for (let row = 0; row < pivot.extent.rows; row += 1) {
            for (let col = 0; col < pivot.extent.cols; col += 1) delete sheet.cells[addressOf({ row: pivot.anchor.row + row, col: pivot.anchor.col + col })]
          }
        }
        sheet.pivots = sheet.pivots!.filter((item) => item.id !== pivotId)
      }
    })
    setToast('Deleted the pivot table')
  }, [mutateWorkbook])

  const activePivot = activeSheet?.pivots?.find((pivot) => {
    const extent = pivot.extent || { rows: 1, cols: 1 }
    return selection.focus.row >= pivot.anchor.row && selection.focus.row < pivot.anchor.row + extent.rows && selection.focus.col >= pivot.anchor.col && selection.focus.col < pivot.anchor.col + extent.cols
  })
  const activePivotSource = useMemo(() => (activePivot && workbook && pivotEditorHidden !== activePivot.id
    ? readPivotSource(workbook, activePivot.source, activeSheet!.id)
    : null
  // The source is re-read when the pivot definition or workbook changes.
  ), [activePivot, activeSheet, pivotEditorHidden, readPivotSource, workbook])

  const traceCells = useCallback((direction: 'precedents' | 'dependents') => {
    const sheet = activeSheet
    const engine = calcEngine
    if (!sheet || !engine || !workbook) return
    const current = traceState?.sheetId === sheet.id ? traceState : null
    const frontier = current?.[direction].length ? current[direction] : [addressOf(selection.focus)]
    const arrows = [...(current?.arrows || [])]
    const seen = new Set(arrows.map((arrow) => `${arrow.from.top},${arrow.from.left},${arrow.from.bottom},${arrow.from.right}>${arrow.to.row},${arrow.to.col}>${arrow.external || ''}`))
    const next: string[] = []
    const push = (arrow: TraceArrow) => {
      const key = `${arrow.from.top},${arrow.from.left},${arrow.from.bottom},${arrow.from.right}>${arrow.to.row},${arrow.to.col}>${arrow.external || ''}`
      if (seen.has(key)) return false
      seen.add(key)
      arrows.push(arrow)
      return true
    }
    const hasError = (sheetId: string, bounds: { top: number; left: number; bottom: number; right: number }) => {
      for (let row = bounds.top; row <= Math.min(bounds.bottom, bounds.top + 30); row += 1) {
        for (let col = bounds.left; col <= Math.min(bounds.right, bounds.left + 30); col += 1) if (isFormulaError(engine.getValue(sheetId, addressOf({ row, col })))) return true
      }
      return false
    }
    let added = 0
    for (const address of frontier) {
      const coord = coordOf(address)
      if (!coord) continue
      if (direction === 'precedents') {
        if (!sheet.cells[address]?.formula) continue
        for (const rect of engine.precedentsOf(sheet.id, address)) {
          const bounds = { top: rect.top, left: rect.left, bottom: Math.min(rect.bottom, 1_048_575), right: Math.min(rect.right, 16_383) }
          const external = rect.sheetId !== sheet.id ? workbook.sheets.find((item) => item.id === rect.sheetId)?.name || 'Another sheet' : undefined
          if (push({ from: external ? { top: coord.row, left: coord.col, bottom: coord.row, right: coord.col } : bounds, to: coord, external, error: hasError(rect.sheetId, bounds) })) added += 1
          if (!external && (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) <= 400) {
            for (let row = bounds.top; row <= bounds.bottom; row += 1) for (let col = bounds.left; col <= bounds.right; col += 1) {
              const precedent = addressOf({ row, col })
              if (sheet.cells[precedent]?.formula) next.push(precedent)
            }
          }
        }
      } else {
        for (const dependent of engine.dependentsOf(sheet.id, address)) {
          const target = coordOf(dependent.address)
          if (!target) continue
          const external = dependent.sheetId !== sheet.id ? workbook.sheets.find((item) => item.id === dependent.sheetId)?.name || 'Another sheet' : undefined
          const error = isFormulaError(engine.getValue(dependent.sheetId, dependent.address))
          if (push(external
            ? { from: { top: coord.row, left: coord.col, bottom: coord.row, right: coord.col }, to: coord, external: `${external}!${dependent.address}`, error, outgoing: true }
            : { from: { top: coord.row, left: coord.col, bottom: coord.row, right: coord.col }, to: target, error })) added += 1
          if (!external) next.push(dependent.address)
        }
      }
    }
    if (!added) {
      setToast(direction === 'precedents'
        ? (current ? 'There are no more precedents to trace.' : 'The active cell has no formula precedents on this sheet or others.')
        : (current ? 'There are no more dependents to trace.' : 'No formulas refer to the active cell.'))
      return
    }
    setTraceState({ sheetId: sheet.id, arrows, precedents: direction === 'precedents' ? next : current?.precedents || [], dependents: direction === 'dependents' ? next : current?.dependents || [] })
  }, [activeSheet, calcEngine, selection.focus, traceState, workbook])

  const activeProtection = sheetProtection(activeSheet)

  const protectActiveSheet = useCallback(async ({ password, allow }: ProtectSheetResult): Promise<string | null> => {
    const sheet = activeSheet
    if (!sheet) return 'There is no sheet to protect.'
    let hash: Record<string, unknown> = {}
    if (password) {
      if (!window.simpleCalc.protection) return 'Password protection is not available in this build.'
      try { hash = await window.simpleCalc.protection.hash(password) } catch (error) { return errorMessage(error) }
    }
    const protection: Record<string, unknown> = { sheet: true, ...hash }
    for (const [key, value] of Object.entries(allow)) {
      if (key === 'selectLockedCells' || key === 'selectUnlockedCells') { if (!value) protection[key] = false }
      else if (key === 'objects') { if (!value) protection.objects = false }
      else if (value) protection[key] = true
    }
    protection.scenarios = false
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === sheet.id)
      if (target) target.sheetProtection = protection
    })
    setProtectDialogOpen(false)
    setToast(`Protected “${sheet.name}”`)
    return null
  }, [activeSheet, mutateWorkbook])

  const unprotectActiveSheet = useCallback(async () => {
    const sheet = activeSheet
    const protection = sheetProtection(sheet)
    if (!sheet || !protection) return
    if (protection.hashValue) {
      const password = await askText({ title: 'Unprotect sheet', label: 'Password' })
      if (password === null) return
      const ok = window.simpleCalc.protection ? await window.simpleCalc.protection.verify(protection, password).catch(() => false) : false
      if (!ok) { setToast("The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization."); return }
    }
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === sheet.id)
      if (target) target.sheetProtection = null
    })
    setToast(`Unprotected “${sheet.name}”`)
  }, [activeSheet, askText, mutateWorkbook])

  const selectionLocked = activeSheet ? cellLocked(activeSheet, addressOf(selection.focus)) : true
  const setSelectionLocked = useCallback((locked: boolean) => {
    const bounds = selectionBounds(selection)
    if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > 200_000) { setToast('Select a smaller range to change cell protection.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)
      if (!sheet) return
      for (let row = bounds.top; row <= bounds.bottom; row += 1) {
        for (let col = bounds.left; col <= bounds.right; col += 1) {
          const address = addressOf({ row, col })
          const cell = { ...(sheet.cells[address] || {}) }
          const style = { ...(cell.style || {}) }
          const protection = { ...((style.protection as Record<string, unknown> | undefined) || {}) }
          if (locked) delete protection.locked
          else protection.locked = false
          if (Object.keys(protection).length) style.protection = protection
          else delete style.protection
          if (Object.keys(style).length) cell.style = style
          else delete cell.style
          if (hasCellContent(cell)) sheet.cells[address] = cell
          else delete sheet.cells[address]
        }
      }
    })
    setToast(locked ? 'Cells locked (takes effect when the sheet is protected)' : 'Cells unlocked: they stay editable on a protected sheet')
  }, [mutateWorkbook, selection])

  // ---- Excel sparkline groups ------------------------------------------------------------------
  const sparklineIndex = useMemo(() => {
    const map = new Map<string, { group: SparklineGroup; groupIndex: number; source: string }>()
    ;(activeSheet?.sparklineGroups || []).forEach((group, groupIndex) => {
      for (const item of group.sparklines) map.set(item.cell.replace(/\$/g, '').toUpperCase(), { group, groupIndex, source: item.source })
    })
    return map
  }, [activeSheet?.sparklineGroups])

  const readSparklineSource = useCallback((source: string): Array<number | null> => {
    const book = workbookRef.current
    const engine = calcEngine
    if (!book || !engine) return []
    const match = /^(?:(?:'((?:[^']|'')+)'|([^!'"]+))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(source.trim())
    if (!match) return []
    const name = match[1] ? match[1].replace(/''/g, "'") : match[2]
    const sheet = name ? book.sheets.find((item) => item.name.toLocaleLowerCase() === name.trim().toLocaleLowerCase()) : book.sheets.find((item) => item.id === book.activeSheetId)
    const bounds = parseTableRef(match[3])
    if (!sheet || !bounds) return []
    const values: Array<number | null> = []
    for (let row = bounds.top; row <= bounds.bottom && values.length < 2000; row += 1) {
      for (let col = bounds.left; col <= bounds.right && values.length < 2000; col += 1) {
        const value = engine.getValue(sheet.id, addressOf({ row, col }))
        values.push(typeof value === 'number' && Number.isFinite(value) ? value : null)
      }
    }
    return values
  }, [calcEngine])

  const sparklineFor = useMemo(() => {
    if (!sparklineIndex.size || !calcEngine) return undefined
    const groupAxis = new Map<number, { min?: number; max?: number }>()
    const axisFor = (group: SparklineGroup, groupIndex: number) => {
      if (group.minAxisType !== 'group' && group.maxAxisType !== 'group' && group.minAxisType !== 'custom' && group.maxAxisType !== 'custom') return undefined
      let cached = groupAxis.get(groupIndex)
      if (!cached) {
        const all = group.sparklines.flatMap((item) => readSparklineSource(item.source)).filter((value): value is number => value !== null)
        cached = {
          ...(group.minAxisType === 'custom' ? { min: group.manualMin } : group.minAxisType === 'group' && all.length ? { min: Math.min(...all) } : {}),
          ...(group.maxAxisType === 'custom' ? { max: group.manualMax } : group.maxAxisType === 'group' && all.length ? { max: Math.max(...all) } : {}),
        }
        groupAxis.set(groupIndex, cached)
      }
      return cached
    }
    return (address: string) => {
      const entry = sparklineIndex.get(address)
      if (!entry) return null
      return groupSparklineSpec(entry.group, readSparklineSource(entry.source), axisFor(entry.group, entry.groupIndex))
    }
  // The workbook dependency redraws sparklines when their data changes.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [calcEngine, readSparklineSource, sparklineIndex, workbook])

  const openCreateSparklines = useCallback((kind: SparklineKind) => {
    const sheet = activeSheet
    if (!sheet) return
    const selected = selectionBounds(selection)
    const region = (selected.top === selected.bottom && selected.left === selected.right ? filterRangeForSelection(sheet, selected, selection.focus) : selected) || selected
    const location = region.bottom > region.top
      ? rangeAddress({ top: region.top, bottom: region.bottom, left: region.right + 1, right: region.right + 1 })
      : rangeAddress({ top: region.top, bottom: region.top, left: region.right + 1, right: region.right + 1 })
    setSparklineDialog({ kind, data: rangeAddress(region), location })
  }, [activeSheet, selection])

  const createSparklines = useCallback((dataText: string, locationText: string, kind: SparklineKind): string | null => {
    const sheet = activeSheet
    if (!sheet) return 'There is no sheet.'
    const dataMatch = /^(?:(?:'((?:[^']|'')+)'|([^!'"]+))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(dataText)
    const data = dataMatch ? parseTableRef(dataMatch[3]) : null
    const location = parseTableRef(locationText.replace(/^[^!]*!/, ''))
    if (!data) return 'Enter a data range such as A2:E10.'
    if (!location) return 'Enter a location range such as F2:F10.'
    const sheetName = dataMatch?.[1] || dataMatch?.[2] ? (dataMatch[1] ? dataMatch[1].replace(/''/g, "'") : dataMatch[2]) : sheet.name
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`
    const cells: Coord[] = []
    for (let row = location.top; row <= location.bottom; row += 1) for (let col = location.left; col <= location.right; col += 1) cells.push({ row, col })
    const rows = data.bottom - data.top + 1
    const columns = data.right - data.left + 1
    let sources: string[]
    if (cells.length === 1) sources = [rangeAddress(data)]
    else if (location.left === location.right && cells.length === rows) sources = Array.from({ length: rows }, (_, index) => rangeAddress({ top: data.top + index, bottom: data.top + index, left: data.left, right: data.right }))
    else if (location.top === location.bottom && cells.length === columns) sources = Array.from({ length: columns }, (_, index) => rangeAddress({ top: data.top, bottom: data.bottom, left: data.left + index, right: data.left + index }))
    else return 'The location range must be one cell, or one cell per row or column of the data.'
    const group: SparklineGroup = {
      type: kind,
      colors: { series: '#376092', negative: '#D00000', markers: '#D00000', high: '#D00000', low: '#D00000', first: '#D00000', last: '#D00000', axis: '#000000' },
      displayEmptyCellsAs: 'gap',
      ...(kind === 'stacked' ? { negative: true } : {}),
      sparklines: cells.map((cell, index) => ({ source: `${quoted}!${sources[index]}`, cell: addressOf(cell) })),
    }
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === sheet.id)
      if (!target) return
      const replaced = new Set(group.sparklines.map((item) => item.cell))
      const kept = (target.sparklineGroups || []).map((existing) => ({ ...existing, sparklines: existing.sparklines.filter((item) => !replaced.has(item.cell)) })).filter((existing) => existing.sparklines.length)
      target.sparklineGroups = [...kept, group]
    })
    setSparklineDialog(null)
    setToast(`Added ${group.sparklines.length} ${group.sparklines.length === 1 ? 'sparkline' : 'sparklines'}`)
    return null
  }, [activeSheet, mutateWorkbook])

  const activeSparkline = sparklineIndex.get(addressOf(selection.focus))
  const updateSparklineGroup = useCallback((groupIndex: number, patch: Partial<SparklineGroup> | null) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)
      if (!sheet?.sparklineGroups?.[groupIndex]) return
      if (patch === null) sheet.sparklineGroups = sheet.sparklineGroups.filter((_, index) => index !== groupIndex)
      else sheet.sparklineGroups[groupIndex] = { ...sheet.sparklineGroups[groupIndex], ...patch }
    })
  }, [mutateWorkbook])

  const insertPicture = useCallback(async (source: Blob | string, name?: string) => {
    const sheet = activeSheet
    if (!sheet) return
    try {
      const picture = await preparePicture(source)
      const scale = Math.min(1, 640 / picture.width, 480 / picture.height)
      const width = Math.max(16, Math.round(picture.width * scale))
      const height = Math.max(16, Math.round(picture.height * scale))
      const image: SheetImage = {
        id: `image-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`,
        ...(name ? { name } : {}),
        src: picture.src,
        anchor: anchorForSize(sheet, selection.focus.row, selection.focus.col, width, height),
      }
      mutateWorkbook((next) => {
        const target = next.sheets.find((item) => item.id === sheet.id)
        if (target) target.images = [...(target.images || []), image]
      })
      setSelectedImageId(image.id)
      setToast(name ? `Inserted ${name}` : 'Inserted picture')
    } catch (error) {
      setToast(errorMessage(error))
    }
  }, [activeSheet, mutateWorkbook, selection.focus.col, selection.focus.row])

  const openPicturePicker = useCallback(() => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = 'image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml'
    input.onchange = () => {
      const file = input.files?.[0]
      if (file) void insertPicture(file, file.name)
    }
    input.click()
  }, [insertPicture])

  const updateImage = useCallback((image: SheetImage) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)
      if (sheet?.images) sheet.images = sheet.images.map((item) => (item.id === image.id ? image : item))
    })
  }, [mutateWorkbook])

  const deleteImage = useCallback((id: string) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)
      if (sheet?.images) sheet.images = sheet.images.filter((item) => item.id !== id)
    })
    setSelectedImageId(null)
    window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true }))
  }, [mutateWorkbook])

  const outlineState = useMemo(() => {
    if (!activeSheet || (!activeSheet.rowProperties && !activeSheet.columnProperties)) return null
    const rowDepth = maxOutlineLevel(activeSheet, 'row')
    const columnDepth = maxOutlineLevel(activeSheet, 'column')
    if (!rowDepth && !columnDepth) return null
    return { rows: outlineGroups(activeSheet, 'row'), columns: outlineGroups(activeSheet, 'column'), rowDepth, columnDepth }
  }, [activeSheet])

  const updateOutline = useCallback((recipe: (sheet: SheetData) => void) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)
      if (sheet) recipe(sheet)
    })
  }, [mutateWorkbook])

  const groupSelection = useCallback((axis: OutlineAxis | null, delta: 1 | -1) => {
    const sheet = activeSheet
    if (!sheet) return
    const bounds = selectionBounds(selection)
    const wholeColumns = bounds.top === 0 && bounds.bottom >= Math.max(0, sheet.rowCount - 1)
    const chosen: OutlineAxis = axis || (wholeColumns ? 'column' : 'row')
    const start = chosen === 'row' ? bounds.top : bounds.left
    const end = chosen === 'row' ? bounds.bottom : bounds.right
    let changed = false
    updateOutline((draft) => { changed = changeOutline(draft, chosen, start, end, delta) })
    const noun = chosen === 'row' ? (start === end ? `row ${start + 1}` : `rows ${start + 1}–${end + 1}`) : (start === end ? `column ${columnName(start)}` : `columns ${columnName(start)}–${columnName(end)}`)
    setToast(changed ? `${delta > 0 ? 'Grouped' : 'Ungrouped'} ${noun}` : delta > 0 ? 'Groups can be nested up to 7 levels.' : `There is no group on ${noun}.`)
  }, [activeSheet, selection, updateOutline])
  groupSelectionRef.current = groupSelection

  const runGoalSeek = useCallback((setCellText: string, toValueText: string, changingText: string): string | null => {
    const sheet = activeSheet
    const engine = calcEngine
    if (!sheet || !engine) return 'There is no sheet to calculate.'
    const clean = (text: string) => text.trim().replace(/^=/, '').replace(/\$/g, '').toUpperCase()
    const setAddress = clean(setCellText)
    const changingAddress = clean(changingText)
    const changing = coordOf(changingAddress)
    if (!coordOf(setAddress)) return 'Set cell must be a single cell reference, such as B5.'
    if (!changing) return 'By changing cell must be a single cell reference, such as B2.'
    if (!sheet.cells[setAddress]?.formula) return 'Set cell must contain a formula.'
    if (sheet.cells[changingAddress]?.formula) return 'By changing cell must contain a value, not a formula.'
    const target = Number(toValueText.trim().replace(/,/g, ''))
    if (!toValueText.trim() || !Number.isFinite(target)) return 'To value must be a number.'
    const evaluate = (x: number) => {
      const value = engine.evaluateAt(sheet.id, setAddress, 0, 0, { row: changing.row, col: changing.col, value: x })
      return typeof value === 'number' ? value : null
    }
    const current = engine.getValue(sheet.id, changingAddress)
    const start = typeof current === 'number' ? current : 0
    const probe = [start, start + 1, start * 2 + 3].map(evaluate)
    if (probe.every((value) => value === probe[0])) return `${setAddress} does not depend on ${changingAddress}.`
    const result = goalSeek(evaluate, target, start)
    if (!result.found) return `Goal Seek may not have found a solution. The closest was ${changingAddress} = ${formatScalar(result.value, 'General')}, giving ${formatScalar(result.achieved, 'General')}.`
    mutateWorkbook((next) => {
      const draft = next.sheets.find((item) => item.id === sheet.id)!
      const cell: CellData = { ...(draft.cells[changingAddress] || {}), value: result.value }
      delete cell.formula
      delete cell.result
      delete cell.display
      draft.cells[changingAddress] = cell
    })
    setGoalSeekOpen(false)
    setToast(`Goal Seek found a solution: ${changingAddress} = ${formatScalar(result.value, 'General')} makes ${setAddress} = ${formatScalar(result.achieved, 'General')}`)
    return null
  }, [activeSheet, calcEngine, mutateWorkbook])

  /** Copy (or cut) the selection: TSV of displayed values plus styled HTML for other apps. */
  const copySelection = useCallback(async (cut = false) => {
    if (!activeSheet) return false
    const bounds = selectionBounds(selection)
    const total = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
    if (total > 100_000) { setToast('That selection is too large to copy at once.'); return false }
    const matrix: CellData[][] = []
    const validations: Array<Array<Record<string, unknown> | undefined>> = []
    const hasValidations = Boolean(activeSheet.dataValidations && Object.keys(activeSheet.dataValidations).length)
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      const cells: CellData[] = []
      const rules: Array<Record<string, unknown> | undefined> = []
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        cells.push(structuredClone(activeSheet.cells[addressOf({ row, col })] || {}) as CellData)
        if (hasValidations) rules.push(findValidation(activeSheet.dataValidations, row, col)?.validation as Record<string, unknown> | undefined)
      }
      matrix.push(cells)
      if (hasValidations) validations.push(rules)
    }
    const merges = (activeSheet.merges || []).flatMap((range) => {
      const merged = mergeBounds(range)
      if (!merged || merged.top < bounds.top || merged.left < bounds.left || merged.bottom > bounds.bottom || merged.right > bounds.right) return []
      return [rangeAddress({ top: merged.top - bounds.top, bottom: merged.bottom - bounds.top, left: merged.left - bounds.left, right: merged.right - bounds.left })]
    })
    const columnWidths = Array.from({ length: bounds.right - bounds.left + 1 }, (_, offset) => activeSheet.colWidths[String(bounds.left + offset + 1)])
    const serialized = serializeSelectionToClipboard(matrix, {
      displayAt: (row, col) => displayValue(activeSheet.id, addressOf({ row: bounds.top + row, col: bounds.left + col })),
      cssColor: (color) => cssColor(color),
      columnWidthsPx: columnWidths.map((width) => (width === undefined ? IMPORTED_COL_WIDTH : columnPixelWidth(width, 1))),
      rowHeightsPx: Array.from({ length: bounds.bottom - bounds.top + 1 }, (_, offset) => {
        const height = activeSheet.rowHeights[String(bounds.top + offset + 1)]
        return height === undefined ? IMPORTED_ROW_HEIGHT : rowPixelHeight(height, 1)
      }),
      merges,
      sheetName: activeSheet.name,
      origin: { row: bounds.top, col: bounds.left },
      defaultFont: { name: String((workbookRef.current?.metadata?.normalFont as { name?: string } | undefined)?.name || 'Calibri'), size: Number((workbookRef.current?.metadata?.normalFont as { size?: number } | undefined)?.size || 11) },
    })
    const written = await writeRichClipboard(serialized)
    if (!written) { setToast('The clipboard is not available.'); return false }
    internalClipboard.current = {
      text: serialized.text,
      origin: { row: bounds.top, col: bounds.left },
      cells: matrix,
      sheetId: activeSheet.id,
      sheetName: activeSheet.name,
      merges,
      columnWidths,
      validations: hasValidations ? validations : undefined,
      cut,
    }
    setCutRange(cut ? { sheetId: activeSheet.id, ...bounds } : null)
    setToast(`${total.toLocaleString()} ${total === 1 ? 'cell' : 'cells'} ${cut ? 'cut' : 'copied'}`)
    return true
  }, [activeSheet, displayValue, selection])

  /** Excel's cut + paste: move the block, keep its formulas, and re-point references to it. */
  const moveClipboardBlock = useCallback((clip: InternalClipboard, destination: Coord) => {
    const current = workbookRef.current
    if (!current) return
    const source = current.sheets.find((sheet) => sheet.id === clip.sheetId)
    const target = current.sheets.find((sheet) => sheet.id === current.activeSheetId)
    if (!source || !target) return
    const height = clip.cells.length
    const width = clip.cells[0]?.length || 0
    const rect = { top: clip.origin.row, left: clip.origin.col, bottom: clip.origin.row + height - 1, right: clip.origin.col + width - 1 }
    const rowDelta = destination.row - rect.top
    const colDelta = destination.col - rect.left
    if (!rowDelta && !colDelta && source.id === target.id) return
    if (destination.row + height > 1_048_576 || destination.col + width > 16_384) { setToast('The paste area extends beyond the sheet.'); return }
    mutateWorkbook((next) => {
      const from = next.sheets.find((sheet) => sheet.id === source.id)!
      const to = next.sheets.find((sheet) => sheet.id === target.id)!
      const moved: Array<[string, CellData]> = []
      for (let row = 0; row < height; row += 1) {
        for (let col = 0; col < width; col += 1) {
          const address = addressOf({ row: rect.top + row, col: rect.left + col })
          const cell = from.cells[address]
          if (cell) moved.push([addressOf({ row: destination.row + row, col: destination.col + col }), cell])
          delete from.cells[address]
        }
      }
      for (let row = 0; row < height; row += 1) {
        for (let col = 0; col < width; col += 1) delete to.cells[addressOf({ row: destination.row + row, col: destination.col + col })]
      }
      for (const [address, cell] of moved) to.cells[address] = cell
      // Merges travel with the block.
      const movedMerges: string[] = []
      from.merges = (from.merges || []).filter((range) => {
        const merged = mergeBounds(range)
        if (!merged || merged.top < rect.top || merged.left < rect.left || merged.bottom > rect.bottom || merged.right > rect.right) return true
        movedMerges.push(rangeAddress({ top: merged.top + rowDelta, bottom: merged.bottom + rowDelta, left: merged.left + colDelta, right: merged.right + colDelta }))
        return false
      })
      to.merges = [...(to.merges || []), ...movedMerges]
      to.rowCount = Math.max(to.rowCount, destination.row + height)
      to.colCount = Math.max(to.colCount, destination.col + width)
      // Every formula that referred to the moved cells now refers to their new place.
      for (const sheet of next.sheets) {
        for (const address in sheet.cells) {
          const cell = sheet.cells[address]
          if (!cell?.formula) continue
          const formula = moveReferencesInFormula(cell.formula, { formulaSheet: sheet.name, sourceSheet: source.name, rect, rowDelta, colDelta, destinationSheet: target.name })
          if (formula !== cell.formula) sheet.cells[address] = { ...cell, formula }
        }
      }
    })
    setSelection({ anchor: destination, focus: { row: destination.row + height - 1, col: destination.col + width - 1 } })
    internalClipboard.current = null
    setCutRange(null)
  }, [mutateWorkbook])

  /** Paste (or Paste Special) at the selection: internal copies keep everything; other apps' HTML keeps formatting. */
  const pasteSelection = useCallback(async (options?: Partial<PasteSpecialOptions>) => {
    if (!activeSheet) return
    try {
      const payload = await readRichClipboard()
      if (!options && !payload.text.trim() && !payload.html && payload.image) { await insertPicture(payload.image); return }
      const clip = internalClipboard.current && internalClipboard.current.text === payload.text ? internalClipboard.current : null
      const bounds = selectionBounds(selection)
      if (clip?.cut && !options) {
        moveClipboardBlock(clip, { row: bounds.top, col: bounds.left })
        return
      }
      let sourceCells: Array<Array<CellData | undefined>>
      let sourceOrigin: Coord
      let merges: string[] = []
      let columnWidths: Array<number | undefined> | undefined
      let validations: InternalClipboard['validations']
      let sourceSheetName = activeSheet.name
      if (clip) {
        sourceCells = clip.cells
        sourceOrigin = clip.origin
        merges = clip.merges
        columnWidths = clip.columnWidths
        validations = clip.validations
        sourceSheetName = clip.sheetName
      } else {
        const parsed = parseClipboardPayload(payload, { origin: { row: bounds.top, col: bounds.left }, defaultFont: { name: 'Calibri', size: 11 } })
        if (!parsed || !parsed.cells.length) return
        sourceCells = parsed.cells
        sourceOrigin = parsed.formulaOrigin
        merges = parsed.merges
        columnWidths = parsed.columnWidths?.map((width) => (width === undefined ? undefined : modelColumnWidth(width, 1)))
        if (parsed.sourceSheetName) sourceSheetName = parsed.sourceSheetName
      }
      const clipSheet = clip ? workbookRef.current?.sheets.find((sheet) => sheet.id === clip.sheetId) : undefined
      const result = applyPasteSpecial({
        source: { cells: sourceCells, origin: sourceOrigin, sheetName: sourceSheetName, merges, columnWidths, validations },
        destination: bounds,
        getDestinationCell: (row, col) => activeSheet.cells[addressOf({ row, col })],
        options: options || { paste: 'all' },
        shiftFormula: shiftFormulaReferences,
        destinationSheetName: activeSheet.name,
        resolveValue: clipSheet && calcEngine ? (_cell, row, col) => calcEngine.getValue(clipSheet.id, addressOf({ row, col })) : undefined,
      })
      if (result.error) { setToast(result.error); return }
      mutateWorkbook((next) => {
        const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
        applyCellChanges(sheet.cells, result.changes)
        if (result.clearMergesIn) {
          const area = result.clearMergesIn
          sheet.merges = (sheet.merges || []).filter((range) => {
            const merged = mergeBounds(range)
            return !merged || merged.bottom < area.top || merged.top > area.bottom || merged.right < area.left || merged.left > area.right
          })
        }
        if (result.merges?.length) sheet.merges = [...(sheet.merges || []), ...result.merges]
        for (const [column, width] of Object.entries(result.columnWidths || {})) {
          if (Number.isFinite(width) && width > 0) sheet.colWidths[String(Number(column) + 1)] = width
        }
        for (const change of result.validations || []) {
          const range = mergeBounds(change.range)
          if (range) sheet.dataValidations = withValidation(sheet.dataValidations, range, change.validation as never)
        }
        const area = result.selection
        sheet.rowCount = Math.max(sheet.rowCount, area.bottom + 1)
        sheet.colCount = Math.max(sheet.colCount, area.right + 1)
      })
      setSelection({ anchor: { row: result.selection.top, col: result.selection.left }, focus: { row: result.selection.bottom, col: result.selection.right } })
    } catch (error) {
      setToast(isClipboardTooLarge(error) ? error.message : errorMessage(error))
    }
  }, [activeSheet, calcEngine, insertPicture, moveClipboardBlock, mutateWorkbook, selection])

  const autofillSelection = useCallback((target: Coord) => {
    if (!activeSheet) return
    const source = selectionBounds(selection)
    const expandedSelection = fillSelectionForTarget(selection, target)
    const expanded = selectionBounds(expandedSelection)
    if (expanded.top === source.top && expanded.bottom === source.bottom && expanded.left === source.left && expanded.right === source.right) return

    const overlapsExpanded = (bounds: { top: number; bottom: number; left: number; right: number }) => (
      bounds.bottom >= expanded.top && bounds.top <= expanded.bottom && bounds.right >= expanded.left && bounds.left <= expanded.right
    )
    if ((activeSheet.merges || []).some((range) => {
      const bounds = mergeBounds(range)
      return Boolean(bounds && overlapsExpanded(bounds))
    })) {
      setToast('Autofill is unavailable across merged cells. Unmerge them first.')
      return
    }
    const protectedFormula = Object.entries(activeSheet.cells).some(([address, cell]) => {
      if (!(cell.formulaType === 'array' || cell.formulaRange || cell.dynamicFormula)) return false
      const formulaBounds = cell.formulaRange ? mergeBounds(cell.formulaRange) : null
      if (formulaBounds) return overlapsExpanded(formulaBounds)
      const coord = coordOf(address)
      return Boolean(coord && overlapsExpanded({ top: coord.row, bottom: coord.row, left: coord.col, right: coord.col }))
    })
    if (protectedFormula) {
      setToast('Autofill cannot overwrite an array formula range.')
      return
    }

    const destination = expanded.top < source.top
      ? { top: expanded.top, bottom: source.top - 1, left: source.left, right: source.right }
      : expanded.bottom > source.bottom
        ? { top: source.bottom + 1, bottom: expanded.bottom, left: source.left, right: source.right }
        : expanded.left < source.left
          ? { top: source.top, bottom: source.bottom, left: expanded.left, right: source.left - 1 }
          : { top: source.top, bottom: source.bottom, left: source.right + 1, right: expanded.right }

    try {
      const patch = createAutofillPatch({ cells: activeSheet.cells, source, destination })
      const changeCount = Object.keys(patch.changes).length
      mutateWorkbook((next) => {
        const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
        Object.entries(patch.changes).forEach(([address, cell]) => {
          if (cell && hasCellContent(cell)) sheet.cells[address] = cell
          else delete sheet.cells[address]
        })
        sheet.rowCount = Math.max(sheet.rowCount, expanded.bottom + 1)
        sheet.colCount = Math.max(sheet.colCount, expanded.right + 1)
      })
      setSelection(expandedSelection)
      setEditing(null)
      setToast(`Filled ${changeCount.toLocaleString()} ${changeCount === 1 ? 'cell' : 'cells'}`)
    } catch (error) {
      setToast(errorMessage(error))
    }
  }, [activeSheet, mutateWorkbook, selection])

  const fillSelectedRange = useCallback((direction: 'down' | 'right') => {
    if (!activeSheet) return
    const bounds = selectionBounds(selection)
    if ((direction === 'down' && bounds.top === bounds.bottom) || (direction === 'right' && bounds.left === bounds.right)) {
      setToast(`Select at least two ${direction === 'down' ? 'rows' : 'columns'} to fill ${direction}.`)
      return
    }
    const overlaps = (range: { top: number; bottom: number; left: number; right: number }) => (
      range.bottom >= bounds.top && range.top <= bounds.bottom && range.right >= bounds.left && range.left <= bounds.right
    )
    if ((activeSheet.merges || []).some((range) => {
      const merged = mergeBounds(range)
      return Boolean(merged && overlaps(merged))
    })) {
      setToast('Fill is unavailable across merged cells. Unmerge them first.')
      return
    }
    if (Object.entries(activeSheet.cells).some(([address, cell]) => {
      if (!(cell.formulaType === 'array' || cell.formulaRange || cell.dynamicFormula)) return false
      const formulaBounds = cell.formulaRange ? mergeBounds(cell.formulaRange) : null
      if (formulaBounds) return overlaps(formulaBounds)
      const coord = coordOf(address)
      return Boolean(coord && overlaps({ top: coord.row, bottom: coord.row, left: coord.col, right: coord.col }))
    })) {
      setToast('Fill cannot overwrite an array formula range.')
      return
    }
    const source = direction === 'down'
      ? { top: bounds.top, bottom: bounds.top, left: bounds.left, right: bounds.right }
      : { top: bounds.top, bottom: bounds.bottom, left: bounds.left, right: bounds.left }
    const destination = direction === 'down'
      ? { top: bounds.top + 1, bottom: bounds.bottom, left: bounds.left, right: bounds.right }
      : { top: bounds.top, bottom: bounds.bottom, left: bounds.left + 1, right: bounds.right }
    try {
      const patch = createAutofillPatch({ cells: activeSheet.cells, source, destination })
      mutateWorkbook((next) => {
        const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
        Object.entries(patch.changes).forEach(([address, cell]) => {
          if (cell && hasCellContent(cell)) sheet.cells[address] = cell
          else delete sheet.cells[address]
        })
      })
      setToast(`Filled ${Object.keys(patch.changes).length.toLocaleString()} cells ${direction}`)
    } catch (error) {
      setToast(errorMessage(error))
    }
  }, [activeSheet, mutateWorkbook, selection])

  /** Rows (0-based) that stay in place while sorting: hidden and filtered rows. */
  const fixedRowsForSort = useCallback((sheet: SheetData) => new Set((sheet.hiddenRows || []).map((row) => Number(row) - 1)), [])

  const applySortResult = useCallback((sheet: SheetData, result: ReturnType<typeof sortRange>, label: string) => {
    if (!result.ok) { setToast(result.error); return false }
    if (!result.changed) { setToast('The range is already sorted.'); return false }
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === sheet.id)!
      applyCellChanges(target.cells, result.changes)
    })
    setToast(label)
    return true
  }, [mutateWorkbook])

  /** Excel's quick sort: by the active cell's column, over the selection or current region. */
  const sortSelectionRange = useCallback((direction: 'asc' | 'desc') => {
    const sheet = activeSheet
    if (!sheet) return
    const selected = selectionBounds(selection)
    const region = selected.top === selected.bottom && selected.left === selected.right
      ? filterRangeForSelection(sheet, selected, selection.focus)
      : selected
    if (!region) { setToast('Select a range with data to sort.'); return }
    if (region.top === region.bottom) { setToast('Select a range with at least two rows to sort.'); return }
    if ((region.bottom - region.top + 1) * (region.right - region.left + 1) > 2_000_000) { setToast('That range is too large to sort at once.'); return }
    const host = dataHostFor(sheet)
    const hasHeader = detectHeaderRow(region, host)
    const key = clamp(selection.focus.col - region.left, 0, region.right - region.left)
    const result = sortRange(sheet, { bounds: region, levels: [{ key, descending: direction === 'desc' }], hasHeader, fixed: fixedRowsForSort(sheet) }, host, shiftFormulaReferences)
    if (applySortResult(sheet, result, `Sorted ${rangeAddress(region)} by column ${columnName(region.left + key)} ${direction === 'asc' ? 'A→Z' : 'Z→A'}`)) {
      setSelection({ anchor: { row: region.top, col: region.left }, focus: { row: region.bottom, col: region.right } })
      setEditing(null)
    }
  }, [activeSheet, applySortResult, dataHostFor, fixedRowsForSort, selection])

  const [filterMenu, setFilterMenu] = useState<{ key: string; col: number; anchor: { left: number; top: number; right: number; bottom: number } } | null>(null)
  const [tableDialog, setTableDialog] = useState<{ range: string; hasHeaders: boolean; style?: string } | null>(null)
  const [tableDesignOpen, setTableDesignOpen] = useState(false)
  const [shiftDialog, setShiftDialog] = useState<{ mode: 'insert' | 'delete'; initial: ShiftCellsChoice } | null>(null)
  const [goalSeekOpen, setGoalSeekOpen] = useState(false)
  const [sortDialogOpen, setSortDialogOpen] = useState(false)
  const [validationDialogOpen, setValidationDialogOpen] = useState(false)
  const [cleanupDialog, setCleanupDialog] = useState<'duplicates' | 'split' | null>(null)
  const [invalidCells, setInvalidCells] = useState<Coord[]>([])
  const [alertRequest, setAlertRequest] = useState<(AlertRequest & { resolve: (button: string) => void }) | null>(null)
  const askAlert = useCallback((request: AlertRequest) => new Promise<string>((resolve) => setAlertRequest({ ...request, resolve })), [])
  confirmRef.current = (message: string) => askAlert({
    title: 'simple_calc',
    message,
    tone: 'warning',
    buttons: [{ id: 'ok', label: 'Continue', primary: true }, { id: 'cancel', label: 'Cancel' }],
  }).then((button) => button === 'ok')

  const insertChart = useCallback(() => {
    const sheet = activeSheet
    if (!sheet || !chartAccessor) return
    const bounds = selectionBounds(selection)
    const chart = createChartFromRange(sheet.name, bounds, chartAccessor, 'column', {
      metrics: {
        columnWidth: (col) => (sheet.hiddenCols?.includes(col + 1) ? 0 : sheet.colWidths[String(col + 1)] === undefined ? IMPORTED_COL_WIDTH : columnPixelWidth(sheet.colWidths[String(col + 1)], 1)),
        rowHeight: (row) => (sheet.hiddenRows?.includes(row + 1) ? 0 : sheet.rowHeights[String(row + 1)] === undefined ? IMPORTED_ROW_HEIGHT : rowPixelHeight(sheet.rowHeights[String(row + 1)], 1)),
      },
    })
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === sheet.id)!
      target.charts = [...(target.charts || []), chart]
    })
    setSelectedChartId(chart.id)
    setChartEditorId(chart.id)
  }, [activeSheet, chartAccessor, mutateWorkbook, selection])

  const updateChart = useCallback((chart: SheetChart) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      sheet.charts = (sheet.charts || []).map((item) => (item.id === chart.id ? structuredClone(chart) as typeof item : item))
    })
  }, [mutateWorkbook])

  const deleteChart = useCallback((id: string) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      sheet.charts = (sheet.charts || []).filter((item) => item.id !== id)
    })
    setSelectedChartId(null)
    setChartEditorId((current) => (current === id ? null : current))
  }, [mutateWorkbook])

  const toggleFilter = useCallback(() => {
    const sheet = activeSheet
    if (!sheet) return
    if (sheet.filter || sheet.autoFilter) {
      mutateWorkbook((next) => {
        const target = next.sheets.find((item) => item.id === sheet.id)!
        target.hiddenRows = hiddenRowsWithoutFilter(target)
        delete target.filteredRows
        delete target.filter
        target.autoFilter = undefined
      })
      setFilterMenu(null)
      setToast('Filter removed')
      return
    }
    const range = filterRangeForSelection(sheet, selectionBounds(selection), selection.focus)
    if (!range) { setToast('Select a range with data to filter.'); return }
    const state = createFilterState(range)
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === sheet.id)!
      target.filter = state
      target.autoFilter = state.ref
    })
    setToast(`Filter created for ${state.ref}`)
  }, [activeSheet, mutateWorkbook, selection])

  toggleFilterRef.current = toggleFilter
  const activeFilter = activeSheet?.filter || (typeof activeSheet?.autoFilter === 'string' && activeSheet.autoFilter ? createFilterState(mergeBounds(activeSheet.autoFilter as string) || { top: 0, bottom: 0, left: 0, right: 0 }) : null)
  const filterBounds = activeFilter ? mergeBounds(activeFilter.ref) : null
  const activeTable = activeSheet ? tableContaining(activeSheet, selection.focus.row, selection.focus.col) : undefined

  /** Every filter with header buttons on the active sheet: its AutoFilter ('sheet') and tables. */
  const filterTargets = useMemo(() => {
    const targets: Array<{ key: string; state: SheetFilterState; bounds: { top: number; bottom: number; left: number; right: number }; table?: SheetTable }> = []
    if (activeFilter && filterBounds) targets.push({ key: 'sheet', state: activeFilter, bounds: filterBounds })
    for (const table of activeSheet?.tables || []) {
      if (table.showFilterButton === false) continue
      const state = tableFilterState(table)
      const bounds = state ? parseTableRef(state.ref) : null
      if (state && bounds) targets.push({ key: table.id, state, bounds, table })
    }
    return targets
  }, [activeFilter, activeSheet?.tables, filterBounds])

  const writeFilterState = useCallback((target: SheetData, key: string, state: SheetFilterState) => {
    if (key === 'sheet') {
      target.filter = state
      target.autoFilter = state.ref
      return
    }
    const table = target.tables?.find((item) => item.id === key)
    if (table) table.filter = state
  }, [])

  const applyFilterCriteria = useCallback((key: string, col: number, criteria: Parameters<typeof setColumnCriteria>[2]) => {
    const sheet = activeSheet
    const target = filterTargets.find((item) => item.key === key)
    if (!sheet || !target) return
    let state = setColumnCriteria(target.state, col - target.bounds.left, criteria)
    if (key === 'sheet') state = extendFilterRange(state, sheet)
    const host = dataHostFor(sheet)
    const fields = refreshAllFilters(sheet, host, key === 'sheet' ? { sheet: state } : { tables: new Map([[key, state]]) })
    mutateWorkbook((next) => {
      const draft = next.sheets.find((item) => item.id === sheet.id)!
      writeFilterState(draft, key, state)
      draft.hiddenRows = fields.hiddenRows
      draft.filteredRows = fields.filteredRows
    })
  }, [activeSheet, dataHostFor, filterTargets, mutateWorkbook, writeFilterState])

  const reapplyFilter = useCallback((clear = false) => {
    const sheet = activeSheet
    if (!sheet || !filterTargets.length) { setToast('This sheet has no filter.'); return }
    const sheetState = activeFilter ? extendFilterRange(clear ? clearFilterCriteria(activeFilter) : activeFilter, sheet) : null
    const tableStates = new Map<string, SheetFilterState | null>()
    for (const target of filterTargets) if (target.table) tableStates.set(target.key, clear ? clearFilterCriteria(target.state) : target.state)
    const fields = refreshAllFilters(sheet, dataHostFor(sheet), { ...(activeFilter ? { sheet: sheetState } : {}), tables: tableStates })
    mutateWorkbook((next) => {
      const draft = next.sheets.find((item) => item.id === sheet.id)!
      if (sheetState) writeFilterState(draft, 'sheet', sheetState)
      for (const [key, state] of tableStates) if (state) writeFilterState(draft, key, state)
      draft.hiddenRows = fields.hiddenRows
      draft.filteredRows = fields.filteredRows
    })
  }, [activeFilter, activeSheet, dataHostFor, filterTargets, mutateWorkbook, writeFilterState])

  const sortFilterColumn = useCallback((key: string, col: number, descending: boolean) => {
    const sheet = activeSheet
    const target = filterTargets.find((item) => item.key === key)
    if (!sheet || !target) return
    const spec = sortSpecForFilter(target.state.ref, { key: col - target.bounds.left, descending }, { fixedRows: [...fixedRowsForSort(sheet)] })
    if (!spec) return
    const result = sortRange(sheet, spec, dataHostFor(sheet), shiftFormulaReferences)
    if (!result.ok) { setToast(result.error); return }
    mutateWorkbook((next) => {
      const draft = next.sheets.find((item) => item.id === sheet.id)!
      if (result.changed) applyCellChanges(draft.cells, result.changes)
      writeFilterState(draft, key, { ...target.state, sort: { column: col - target.bounds.left, descending } })
    })
  }, [activeSheet, dataHostFor, filterTargets, fixedRowsForSort, mutateWorkbook, writeFilterState])

  // ---- Tables -----------------------------------------------------------------------------------
  const openCreateTable = useCallback((style?: string) => {
    const sheet = activeSheet
    if (!sheet) return
    const existing = tableContaining(sheet, selection.focus.row, selection.focus.col)
    if (existing) {
      if (style) {
        mutateWorkbook((next) => {
          const table = next.sheets.find((item) => item.id === sheet.id)?.tables?.find((item) => item.id === existing.id)
          if (table) table.style = { ...(table.style || {}), theme: style }
        })
      } else setTableDesignOpen(true)
      return
    }
    const selected = selectionBounds(selection)
    const region = (selected.top === selected.bottom && selected.left === selected.right ? filterRangeForSelection(sheet, selected, selection.focus) : selected) || selected
    setTableDialog({ range: rangeAddress(region), hasHeaders: detectHeaderRow(region, dataHostFor(sheet)), style })
  }, [activeSheet, dataHostFor, mutateWorkbook, selection])
  openCreateTableRef.current = () => openCreateTable()

  const createTableFromDialog = useCallback((range: string, hasHeaders: boolean, style: string): string | null => {
    const sheet = activeSheet
    const current = workbookRef.current
    if (!sheet || !current) return 'There is no sheet to add a table to.'
    const bounds = parseTableRef(range)
    if (!bounds) return 'Enter a range such as A1:D20.'
    let error: string | null = null
    let createdTable: SheetTable | null = null
    const trial = structuredClone(current)
    const trialSheet = trial.sheets.find((item) => item.id === sheet.id)!
    const outcome = createTable(trial, trialSheet, bounds, { hasHeaders, style })
    if (typeof outcome === 'string') error = outcome
    if (error) return error
    const overlapsSheetFilter = filterBounds && !(filterBounds.right < bounds.left || bounds.right < filterBounds.left || filterBounds.bottom < bounds.top || bounds.bottom < filterBounds.top)
    mutateWorkbook((next) => {
      const draft = next.sheets.find((item) => item.id === sheet.id)!
      if (overlapsSheetFilter) {
        // The table takes over the range's AutoFilter, as Excel converts it.
        draft.hiddenRows = hiddenRowsWithoutFilter(draft)
        delete draft.filteredRows
        delete draft.filter
        draft.autoFilter = undefined
      }
      const result = createTable(next, draft, bounds, { hasHeaders, style, id: (outcome as SheetTable).id })
      if (typeof result !== 'string') createdTable = result
    })
    setTableDialog(null)
    const made = createdTable as SheetTable | null
    if (made) {
      const regions = tableRegions(made)
      if (regions) setSelection({ anchor: { row: regions.top, col: regions.left }, focus: { row: regions.top, col: regions.left } })
      setToast(`Created ${made.name} (${made.ref})`)
    }
    return null
  }, [activeSheet, filterBounds, mutateWorkbook])

  const updateTable = useCallback((tableId: string, recipe: (workbook: WorkbookModel, sheet: SheetData, table: SheetTable) => void) => {
    const sheet = activeSheet
    if (!sheet) return
    mutateWorkbook((next) => {
      const draft = next.sheets.find((item) => item.id === sheet.id)
      const table = draft?.tables?.find((item) => item.id === tableId)
      if (draft && table) recipe(next, draft, table)
    })
  }, [activeSheet, mutateWorkbook])

  const setTableOption = useCallback((tableId: string, option: TableOption, value: boolean) => {
    updateTable(tableId, (workbook, sheet, table) => {
      if (option === 'totalsRow') { setTotalsRow(workbook, sheet, table, value); return }
      if (option === 'showFilterButton') {
        table.showFilterButton = value
        if (!value && table.filter && filterIsActive(table.filter)) {
          delete table.filter
          const fields = refreshAllFilters(sheet, dataHostFor(sheet))
          sheet.hiddenRows = fields.hiddenRows
          sheet.filteredRows = fields.filteredRows
        }
        return
      }
      if (option === 'headerRow') {
        const regions = tableRegions(table)
        if (!regions || (table.headerRow !== false) === value) return
        if (!value) {
          // The header row leaves the table; its cells are cleared (Excel keeps the names).
          for (let col = regions.left; col <= regions.right; col += 1) {
            const address = addressOf({ row: regions.top, col })
            const cell = sheet.cells[address]
            if (!cell) continue
            const kept: CellData = { ...cell }
            delete kept.value
            if (kept.style || kept.numFmt || kept.note) sheet.cells[address] = kept
            else delete sheet.cells[address]
          }
          table.headerRow = false
          table.ref = formatTableRef({ ...regions, top: regions.top + 1 })
          delete table.filter
          return
        }
        let top = regions.top - 1
        const aboveFree = top >= 0 && Array.from({ length: regions.right - regions.left + 1 }, (_, offset) => sheet.cells[addressOf({ row: top, col: regions.left + offset })])
          .every((cell) => !cell || (cell.value == null || cell.value === '') && !cell.formula)
        if (!aboveFree) {
          shiftCellsDown(workbook, sheet, regions.left, regions.right, regions.top, 1)
          top = regions.top
          table.ref = formatTableRef({ ...regions, top, bottom: regions.bottom + 1 })
        } else table.ref = formatTableRef({ ...regions, top })
        table.headerRow = true
        table.columns.forEach((column, offset) => { sheet.cells[addressOf({ row: top, col: regions.left + offset })] = { ...(sheet.cells[addressOf({ row: top, col: regions.left + offset })] || {}), value: column.name } })
        if (table.totalsRow) writeTotalsCells(sheet, table)
        return
      }
      table.style = { ...(table.style || {}), [option]: value }
    })
  }, [dataHostFor, updateTable])

  const renameActiveTable = useCallback((tableId: string, name: string): string | null => {
    const current = workbookRef.current
    if (!current) return null
    const trial = structuredClone(current)
    const error = renameTable(trial, tableId, name)
    if (error) return error
    mutateWorkbook((next) => { renameTable(next, tableId, name) })
    return null
  }, [mutateWorkbook])

  const resizeActiveTable = useCallback((tableId: string, range: string): string | null => {
    const sheet = activeSheet
    const bounds = parseTableRef(range)
    if (!sheet || !bounds) return 'Enter a range such as A1:D20.'
    const trial = structuredClone(sheet)
    const trialTable = trial.tables?.find((item) => item.id === tableId)
    if (!trialTable) return 'The table no longer exists.'
    const error = resizeTable(trial, trialTable, bounds)
    if (error) return error
    updateTable(tableId, (_workbook, draft, table) => { resizeTable(draft, table, bounds) })
    return null
  }, [activeSheet, updateTable])

  const convertActiveTable = useCallback((tableId: string) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.tables?.some((table) => table.id === tableId))
      const table = sheet?.tables?.find((item) => item.id === tableId)
      if (sheet && table?.filter && filterIsActive(table.filter)) {
        delete table.filter
        const fields = refreshAllFilters(sheet, dataHostFor(sheet))
        sheet.hiddenRows = fields.hiddenRows
        sheet.filteredRows = fields.filteredRows
      }
      convertTableToRange(next, tableId, activeThemeColors)
    })
    setTableDesignOpen(false)
    setToast('Converted the table to a normal range')
  }, [dataHostFor, mutateWorkbook])

  const runCustomSort = useCallback((result: SortDialogResult) => {
    const sheet = activeSheet
    if (!sheet) return
    const selected = selectionBounds(selection)
    const region = selected.top === selected.bottom && selected.left === selected.right ? filterRangeForSelection(sheet, selected, selection.focus) : selected
    if (!region) return
    const outcome = sortRange(sheet, { bounds: region, levels: result.levels, hasHeader: result.hasHeader, orientation: result.orientation, caseSensitive: result.caseSensitive, fixed: result.orientation === 'rows' ? fixedRowsForSort(sheet) : undefined }, dataHostFor(sheet), shiftFormulaReferences)
    if (applySortResult(sheet, outcome, `Sorted ${rangeAddress(region)}`)) setSortDialogOpen(false)
  }, [activeSheet, applySortResult, dataHostFor, fixedRowsForSort, selection])

  const applyCleanup = useCallback((kind: 'trim' | 'upper' | 'lower' | 'proper' | 'sentence' | 'fill-blanks') => {
    const sheet = activeSheet
    if (!sheet) return
    const bounds = selectionBounds(selection)
    const outcome = kind === 'trim' ? trimWhitespace(sheet, bounds)
      : kind === 'fill-blanks' ? fillBlanksFromAbove(sheet, bounds, shiftFormulaReferences)
        : changeCase(sheet, bounds, kind)
    if (!outcome.count) { setToast('Nothing to change in the selection.'); return }
    mutateWorkbook((next) => applyCellChanges(next.sheets.find((item) => item.id === sheet.id)!.cells, outcome.changes))
    setToast(`Updated ${outcome.count.toLocaleString()} ${outcome.count === 1 ? 'cell' : 'cells'}`)
  }, [activeSheet, mutateWorkbook, selection])

  const circleInvalidData = useCallback(() => {
    const sheet = activeSheet
    if (!sheet?.dataValidations) { setToast('This sheet has no data validation.'); return }
    const found = findInvalidCells(sheet.dataValidations, dataHostFor(sheet), { extent: { rows: sheet.rowCount, cols: sheet.colCount }, limit: 5000 })
    setInvalidCells(found.map((address) => coordOf(address)).filter((coord): coord is Coord => Boolean(coord)))
    setToast(found.length ? `${found.length.toLocaleString()} invalid ${found.length === 1 ? 'cell' : 'cells'} circled` : 'No invalid data found')
  }, [activeSheet, dataHostFor])

  /**
   * Commit an edit through the cell's data validation. Returns false when the entry was
   * rejected or needs confirmation (the alert then decides and may reopen the editor).
   */
  const commitWithValidation = useCallback((address: string, draft: string) => {
    const sheet = workbookRef.current?.sheets.find((item) => item.id === workbookRef.current?.activeSheetId)
    const coord = coordOf(address)
    if (!sheet || !coord || draft.startsWith('=')) { commitCell(address, draft); return true }
    const found = findValidation(sheet.dataValidations, coord.row, coord.col)
    if (!found || rawCellValue(sheet.cells[address]) === draft) { commitCell(address, draft); return true }
    const parsed = parseDraft(draft, sheet.cells[address]).value
    const outcome = validateValue(found.validation, draft, parsed ?? null, dataHostFor(sheet), { row: coord.row, col: coord.col, anchor: found.anchor })
    if (outcome.ok || outcome.alert === 'none') {
      if (outcome.value !== undefined && outcome.value !== parsed && typeof outcome.value !== 'object') {
        commitCell(address, typeof outcome.value === 'number' && !Number.isNaN(outcome.value) ? String(outcome.value) : String(outcome.value ?? ''))
      } else commitCell(address, draft)
      return true
    }
    const tone = outcome.alert === 'warning' ? 'warning' : outcome.alert === 'information' ? 'information' : 'stop'
    const buttons = tone === 'stop'
      ? [{ id: 'retry', label: 'Retry', primary: true }, { id: 'cancel', label: 'Cancel' }]
      : tone === 'warning'
        ? [{ id: 'yes', label: 'Yes', primary: true }, { id: 'no', label: 'No' }, { id: 'cancel', label: 'Cancel' }]
        : [{ id: 'ok', label: 'OK', primary: true }, { id: 'cancel', label: 'Cancel' }]
    const message = outcome.message?.text || 'This value doesn’t match the data validation restrictions defined for this cell.'
    setEditing(null)
    void askAlert({ title: outcome.message?.title || 'simple_calc', message: tone === 'warning' ? `${message}\n\nContinue?` : message, tone, buttons }).then((button) => {
      if (button === 'yes' || button === 'ok') commitCell(address, draft)
      else if (button === 'retry' || button === 'no') {
        setSelection({ anchor: coord, focus: coord })
        setEditing({ address, draft, mode: 'edit' })
      }
    })
    return false
  }, [askAlert, commitCell, dataHostFor])

  const searchMatches = useMemo(() => {
    const query = searchQuery.trim()
    if (!workbook || !activeSheet || !query) return [] as Array<{ sheetId: string; address: string }>
    const sheets = searchOptions.workbook
      ? workbook.sheets.filter((sheet) => sheet.state === 'visible')
      : [activeSheet]
    const output: Array<{ sheetId: string; address: string }> = []
    for (const sheet of sheets) {
      for (const address of sortedCellAddresses(sheet.cells)) {
        const cell = sheet.cells[address]
        if (!cell) continue
        const candidates = searchOptions.lookIn === 'formulas'
          ? [rawCellValue(cell)]
          : [displayValue(sheet.id, address), cell.formula ? '' : rawCellValue(cell)]
        if (candidates.some((value) => value && textMatches(value, query, searchOptions))) output.push({ sheetId: sheet.id, address })
        if (output.length >= 50_000) return output
      }
    }
    return output
  }, [activeSheet, displayValue, searchOptions, searchQuery, workbook])

  const revealSearchMatch = useCallback((index: number) => {
    if (!searchMatches.length) return
    const normalized = (index + searchMatches.length) % searchMatches.length
    const match = searchMatches[normalized]
    const coord = coordOf(match.address)
    if (!coord) return
    setSearchIndex(normalized)
    if (workbookRef.current?.activeSheetId !== match.sheetId) setWorkbook((current) => current ? { ...current, activeSheetId: match.sheetId } : current)
    setSelection({ anchor: coord, focus: coord })
    setEditing(null)
  }, [searchMatches])

  const moveSearch = useCallback((direction: 1 | -1) => {
    revealSearchMatch(searchIndex < 0 ? (direction === 1 ? 0 : searchMatches.length - 1) : searchIndex + direction)
  }, [revealSearchMatch, searchIndex, searchMatches.length])

  const openSearch = useCallback(() => {
    searchOriginRef.current = activeAddress
    setSearchOpen(true)
    window.requestAnimationFrame(() => {
      searchInputRef.current?.focus()
      searchInputRef.current?.select()
    })
  }, [activeAddress])
  const closeSearch = useCallback(() => {
    setSearchOpen(false)
    setReplaceOpen(false)
    window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true }))
  }, [])

  const openReplace = useCallback(() => {
    setReplaceOpen(true)
    openSearch()
  }, [openSearch])

  const replaceCurrentMatch = useCallback(() => {
    if (!workbook || !searchMatches.length) return
    const query = searchQuery.trim()
    if (!query) return
    const match = searchMatches[searchIndex < 0 ? 0 : Math.min(searchIndex, searchMatches.length - 1)]
    const sheet = workbook.sheets.find((item) => item.id === match.sheetId)
    const cell = sheet?.cells[match.address]
    const draft = cell ? replacementDraft(cell, query, replaceValue, searchOptions) : null
    if (draft === null) { setToast('That match only appears in a formatted result and cannot be replaced.'); return }
    searchOriginRef.current = match.address
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === match.sheetId)!
      const updated = parseDraft(draft, target.cells[match.address])
      if (hasCellContent(updated)) target.cells[match.address] = updated
      else delete target.cells[match.address]
    })
  }, [mutateWorkbook, replaceValue, searchIndex, searchMatches, searchOptions, searchQuery, workbook])

  const replaceAllMatches = useCallback(() => {
    if (!workbook || !searchMatches.length) return
    const query = searchQuery.trim()
    if (!query) return
    const updates: Array<{ sheetId: string; address: string; draft: string }> = []
    searchMatches.forEach(({ sheetId, address }) => {
      const cell = workbook.sheets.find((item) => item.id === sheetId)?.cells[address]
      const draft = cell ? replacementDraft(cell, query, replaceValue, searchOptions) : null
      if (draft !== null) updates.push({ sheetId, address, draft })
    })
    if (!updates.length) { setToast('No replaceable matches.'); return }
    mutateWorkbook((next) => {
      updates.forEach(({ sheetId, address, draft }) => {
        const sheet = next.sheets.find((item) => item.id === sheetId)!
        const updated = parseDraft(draft, sheet.cells[address])
        if (hasCellContent(updated)) sheet.cells[address] = updated
        else delete sheet.cells[address]
      })
    })
    setToast(`Replaced ${updates.length} ${updates.length === 1 ? 'cell' : 'cells'}`)
  }, [mutateWorkbook, replaceValue, searchMatches, searchOptions, searchQuery, workbook])

  useEffect(() => {
    if (!searchOpen) return
    const frame = window.requestAnimationFrame(() => searchInputRef.current?.focus())
    return () => window.cancelAnimationFrame(frame)
  }, [searchOpen])

  useEffect(() => {
    if (!searchOpen || !searchQuery.trim() || !searchMatches.length) {
      setSearchIndex(-1)
      return
    }
    const origin = coordOf(searchOriginRef.current)
    const activeId = workbookRef.current?.activeSheetId
    const nextIndex = origin
      ? searchMatches.findIndex(({ sheetId, address }) => {
          const match = coordOf(address)
          return sheetId === activeId && Boolean(match && (match.row > origin.row || (match.row === origin.row && match.col > origin.col)))
        })
      : 0
    revealSearchMatch(nextIndex >= 0 ? nextIndex : 0)
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revealSearchMatch, searchMatches, searchOpen, searchQuery])

  const addSheet = useCallback(() => {
    mutateWorkbook((next) => {
      let index = next.sheets.length + 1
      while (next.sheets.some((sheet) => sheet.name.toLocaleLowerCase() === `sheet${index}`.toLocaleLowerCase())) index += 1
      const id = makeId()
      next.sheets.push({ id, name: `Sheet${index}`, state: 'visible', rowCount: DEFAULT_ROWS, colCount: DEFAULT_COLS, cells: {}, merges: [], colWidths: {}, rowHeights: {}, frozen: {} })
      next.activeSheetId = id
    })
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
  }, [mutateWorkbook])

  const renameSheet = useCallback(async (sheetId: string) => {
    if (!workbook) return
    const sheet = workbook.sheets.find((item) => item.id === sheetId)
    if (!sheet) return
    const name = (await askText({ title: 'Rename sheet', label: 'Sheet name', initialValue: sheet.name }))?.trim()
    if (!name || name === sheet.name) return
    if (workbook.sheets.some((item) => item.id !== sheetId && item.name.toLocaleLowerCase() === name.toLocaleLowerCase())) { setToast('Sheet names must be unique.'); return }
    if (/[\\/?*[\]:]/.test(name) || name.startsWith("'") || name.endsWith("'")) { setToast('Sheet names cannot contain \\ / ? * [ ] : or start/end with an apostrophe.'); return }
    const finalName = name.slice(0, 31)
    mutateWorkbook((next) => {
      next.sheets.find((item) => item.id === sheetId)!.name = finalName
      // Excel keeps every reference pointing at the renamed sheet.
      rewriteWorkbookFormulas(next, (formula) => renameSheetInFormula(formula, sheet.name, finalName))
      for (const item of next.sheets) for (const pivot of item.pivots || []) pivot.source = renameSheetInFormula(pivot.source, sheet.name, finalName)
      const renamed = renameWorkbookChartReferences(next as WorkbookModel, sheet.name, finalName)
      renamed.sheets.forEach((item, index) => { if (item.charts !== next.sheets[index].charts) next.sheets[index].charts = item.charts })
    })
  }, [askText, mutateWorkbook, workbook])

  const deleteActiveSheet = useCallback(async () => {
    if (!workbook || workbook.sheets.length <= 1) { setToast('A workbook needs at least one sheet.'); return }
    if (!(await confirmRef.current(`Delete “${activeSheet?.name}”? This can be undone.`))) return
    mutateWorkbook((next) => {
      const index = next.sheets.findIndex((sheet) => sheet.id === next.activeSheetId)
      const [removed] = next.sheets.splice(index, 1)
      next.activeSheetId = (next.sheets.slice(Math.max(0, index - 1)).find((sheet) => sheet.state === 'visible') || next.sheets.find((sheet) => sheet.state === 'visible') || next.sheets[0]).id
      if (removed) rewriteWorkbookFormulas(next, (formula) => removeSheetFromFormula(formula, removed.name))
    })
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
  }, [activeSheet?.name, mutateWorkbook, workbook])

  const moveSheet = useCallback((sheetId: string, targetIndex: number) => {
    const current = workbookRef.current
    if (!current) return
    const from = current.sheets.findIndex((sheet) => sheet.id === sheetId)
    const to = clamp(targetIndex, 0, current.sheets.length - 1)
    if (from < 0 || from === to) return
    mutateWorkbook((next) => {
      const [moved] = next.sheets.splice(from, 1)
      next.sheets.splice(to, 0, moved)
    })
  }, [mutateWorkbook])

  const setSheetTabColor = useCallback((sheetId: string, color: string | null) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === sheetId)
      if (!sheet) return
      const properties = { ...(sheet.properties || {}) }
      if (color) properties.tabColor = { argb: `FF${color}` }
      else delete properties.tabColor
      sheet.properties = properties
    })
  }, [mutateWorkbook])

  const duplicateActiveSheet = useCallback(() => {
    if (!activeSheet) return
    const copy = structuredClone(activeSheet) as SheetData
    copy.id = makeId()
    mutateWorkbook((next) => {
      const sourceIndex = next.sheets.findIndex((sheet) => sheet.id === next.activeSheetId)
      const base = `${copy.name} copy`
      let name = base
      let suffix = 2
      while (next.sheets.some((sheet) => sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())) name = `${base} ${suffix++}`
      copy.name = name.slice(0, 31)
      next.sheets.splice(sourceIndex + 1, 0, copy)
      next.activeSheetId = copy.id
    })
  }, [activeSheet, mutateWorkbook])

  const setFreeze = useCallback((axis: 'rows' | 'columns' | 'both', count: number, columns = count) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      const frozen = axis === 'both'
        ? { ...(sheet.frozen || {}), rows: Math.max(0, Math.trunc(count)), columns: Math.max(0, Math.trunc(columns)) }
        : { ...(sheet.frozen || {}), [axis]: Math.max(0, Math.trunc(count)) }
      if ((frozen.rows || 0) || (frozen.columns || 0)) frozen.topLeftCell = addressOf({ row: frozen.rows || 0, col: frozen.columns || 0 })
      else delete frozen.topLeftCell
      sheet.frozen = frozen
    })
  }, [mutateWorkbook])

  const toggleGridlines = useCallback(() => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      const views = Array.isArray(sheet.views) ? [...sheet.views] : []
      views[0] = { ...(views[0] || {}), showGridLines: !(views[0]?.showGridLines !== false) }
      sheet.views = views
    })
  }, [mutateWorkbook])

  const hideSelectedDimension = useCallback((axis: 'rows' | 'columns') => {
    const bounds = selectionBounds(selection)
    const start = axis === 'rows' ? bounds.top : bounds.left
    const end = axis === 'rows' ? bounds.bottom : bounds.right
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      const hidden = new Set((axis === 'rows' ? sheet.hiddenRows : sheet.hiddenCols) || [])
      for (let index = start; index <= end; index += 1) hidden.add(index + 1)
      const sorted = [...hidden].sort((a, b) => a - b)
      if (axis === 'rows') sheet.hiddenRows = sorted
      else sheet.hiddenCols = sorted
    })
  }, [mutateWorkbook, selection])

  const unhideNearSelection = useCallback((axis: 'rows' | 'columns') => {
    const bounds = selectionBounds(selection)
    const start = axis === 'rows' ? bounds.top : bounds.left
    const end = axis === 'rows' ? bounds.bottom : bounds.right
    const hidden = new Set(((axis === 'rows' ? activeSheet?.hiddenRows : activeSheet?.hiddenCols) || []).map((value) => Number(value) - 1))
    const reveal = new Set<number>()
    hidden.forEach((index) => { if (index >= start && index <= end) reveal.add(index) })
    for (let index = start - 1; hidden.has(index); index -= 1) reveal.add(index)
    for (let index = end + 1; hidden.has(index); index += 1) reveal.add(index)
    if (!reveal.size) { setToast(`No hidden ${axis} touch the selection.`); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      const kept = ((axis === 'rows' ? sheet.hiddenRows : sheet.hiddenCols) || []).filter((value) => !reveal.has(Number(value) - 1))
      if (axis === 'rows') sheet.hiddenRows = kept
      else sheet.hiddenCols = kept
    })
  }, [activeSheet, mutateWorkbook, selection])

  const hideActiveSheet = useCallback(() => {
    if (!workbook || !activeSheet) return
    const visible = workbook.sheets.filter((sheet) => sheet.state === 'visible')
    if (visible.length <= 1) { setToast('At least one sheet must remain visible.'); return }
    mutateWorkbook((next) => {
      const index = next.sheets.findIndex((sheet) => sheet.id === next.activeSheetId)
      next.sheets[index].state = 'hidden'
      const replacement = next.sheets.find((sheet, sheetIndex) => sheetIndex > index && sheet.state === 'visible') || [...next.sheets].reverse().find((sheet) => sheet.state === 'visible')!
      next.activeSheetId = replacement.id
    })
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
  }, [activeSheet, mutateWorkbook, workbook])

  const unhideSheet = useCallback((sheetId: string) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === sheetId)
      if (sheet && sheet.state === 'hidden') sheet.state = 'visible'
    })
  }, [mutateWorkbook])

  const activateSheet = useCallback((sheetId: string) => {
    setWorkbook((current) => current ? { ...current, activeSheetId: sheetId } : current)
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
    setEditing(null)
  }, [])

  const switchSheetBy = useCallback((delta: number) => {
    const current = workbookRef.current
    if (!current) return
    const visible = current.sheets.filter((sheet) => sheet.state === 'visible')
    const index = visible.findIndex((sheet) => sheet.id === current.activeSheetId)
    const next = visible[clamp(index + delta, 0, visible.length - 1)]
    if (next && next.id !== current.activeSheetId) activateSheet(next.id)
  }, [activateSheet])

  const toggleFullscreen = useCallback(async () => {
    try {
      if (document.fullscreenElement) await document.exitFullscreen()
      else await document.documentElement.requestFullscreen()
    } catch {
      setImmersive((current) => !current)
    }
  }, [])

  useEffect(() => {
    const update = () => setImmersive(Boolean(document.fullscreenElement))
    document.addEventListener('fullscreenchange', update)
    return () => document.removeEventListener('fullscreenchange', update)
  }, [])

  useEffect(() => {
    const timeout = window.setTimeout(() => setStatsSelection(selection), 120)
    return () => window.clearTimeout(timeout)
  }, [selection])

  useEffect(() => {
    const press = (event: globalThis.PointerEvent) => { if (event.button === 0) setPointerDown(true) }
    const release = () => setPointerDown(false)
    window.addEventListener('pointerdown', press)
    window.addEventListener('pointerup', release)
    window.addEventListener('pointercancel', release)
    window.addEventListener('blur', release)
    return () => {
      window.removeEventListener('pointerdown', press)
      window.removeEventListener('pointerup', release)
      window.removeEventListener('pointercancel', release)
      window.removeEventListener('blur', release)
    }
  }, [])

  const selectionStats = useMemo(() => {
    if (!activeSheet) return { selected: 0, count: 0, numeric: 0, sum: 0, average: 0, minimum: 0, maximum: 0, numFmt: undefined as string | undefined }
    const bounds = selectionBounds(statsSelection)
    const selected = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
    let count = 0
    let numeric = 0
    let sum = 0
    let minimum = Number.POSITIVE_INFINITY
    let maximum = Number.NEGATIVE_INFINITY
    let commonNumFmt: string | undefined
    let mixedNumFmt = false
    const includeCell = (address: string, cell: CellData | undefined) => {
      if (!cell) return
      const value = calcEngine ? calcEngine.getValue(activeSheet.id, address) : cell.value
      if (cell.formula || (value !== null && value !== undefined && value !== '')) count += 1
      if (typeof value === 'number' && Number.isFinite(value)) {
        const numFmt = cell.numFmt || cell.style?.numFmt
        if (numeric === 0) commonNumFmt = numFmt
        else if (numFmt !== commonNumFmt) mixedNumFmt = true
        numeric += 1
        sum += value
        minimum = Math.min(minimum, value)
        maximum = Math.max(maximum, value)
      }
    }
    if (selected <= 20_000) {
      for (let row = bounds.top; row <= bounds.bottom; row += 1) {
        for (let col = bounds.left; col <= bounds.right; col += 1) {
          const address = addressOf({ row, col })
          includeCell(address, activeSheet.cells[address])
        }
      }
    } else if (!pointerDown) {
      Object.entries(activeSheet.cells).forEach(([address, cell]) => {
        const coord = coordOf(address)
        if (!coord || coord.row < bounds.top || coord.row > bounds.bottom || coord.col < bounds.left || coord.col > bounds.right) return
        includeCell(address, cell)
      })
    }
    return {
      selected,
      count,
      numeric,
      sum,
      average: numeric ? sum / numeric : 0,
      minimum: numeric ? minimum : 0,
      maximum: numeric ? maximum : 0,
      numFmt: mixedNumFmt ? undefined : commonNumFmt,
    }
  }, [activeSheet, calcEngine, pointerDown, statsSelection])

  const statusItems = STATUS_STATS.filter((stat) => statusStats.has(stat.key)).flatMap((stat) => {
    const counting = stat.key === 'count' || stat.key === 'numeric'
    if (!counting && selectionStats.numeric === 0) return []
    if (stat.key === 'count' && selectionStats.count < 2 && selectionStats.numeric === 0) return []
    const value = stat.key === 'count' ? selectionStats.count
      : stat.key === 'numeric' ? selectionStats.numeric
        : stat.key === 'average' ? selectionStats.average
          : stat.key === 'minimum' ? selectionStats.minimum
            : stat.key === 'maximum' ? selectionStats.maximum
              : selectionStats.sum
    const text = counting
      ? value.toLocaleString()
      : formatScalar(value, selectionStats.numFmt, value.toLocaleString(undefined, { maximumFractionDigits: 10 }))
    return [{ ...stat, value, text }]
  })

  const moveSelectionImpl = useCallback((rowDelta: number, colDelta: number, extend = false) => {
    const next = {
      row: stepPastHidden(selection.focus.row, rowDelta, 1_048_575, hiddenRowSet),
      col: stepPastHidden(selection.focus.col, colDelta, 16_383, hiddenColSet),
    }
    setSelection((current) => ({ anchor: extend ? current.anchor : next, focus: next }))
  }, [hiddenColSet, hiddenRowSet, selection.focus.col, selection.focus.row])
  const moveSelection = moveSelectionImpl
  moveSelectionRef.current = moveSelectionImpl

  const jumpSelection = useCallback((rowDelta: number, colDelta: number, extend = false) => {
    if (!activeSheet) return
    const next = edgeJumpCoord(activeSheet, selection.focus, rowDelta, colDelta, rowDelta !== 0 ? hiddenRowSet : hiddenColSet)
    setSelection((current) => ({ anchor: extend ? current.anchor : next, focus: next }))
  }, [activeSheet, hiddenColSet, hiddenRowSet, selection.focus])

  const selectAllSheet = useCallback(() => {
    const current = workbookRef.current
    const sheet = current?.sheets.find((item) => item.id === current.activeSheetId)
    if (!sheet) return
    setSelection({
      anchor: { row: 0, col: 0 },
      focus: { row: Math.max(0, sheet.rowCount - 1), col: Math.max(0, sheet.colCount - 1) },
    })
    setEditing(null)
  }, [])

  /** Qualified absolute reference for the selection, e.g. 'My Sheet'!$A$1:$C$9. */
  const selectionReferenceText = useCallback(() => {
    const sheet = activeSheet
    if (!sheet) return ''
    const bounds = selectionBounds(selectionRef.current)
    const absolute = (coord: Coord) => `$${columnName(coord.col)}$${coord.row + 1}`
    const range = bounds.top === bounds.bottom && bounds.left === bounds.right
      ? absolute({ row: bounds.top, col: bounds.left })
      : `${absolute({ row: bounds.top, col: bounds.left })}:${absolute({ row: bounds.bottom, col: bounds.right })}`
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet.name) && !/^[A-Za-z]{1,3}\d+$/.test(sheet.name) ? sheet.name : `'${sheet.name.replace(/'/g, "''")}'`
    return `${quoted}!${range}`
  }, [activeSheet])

  const selectReferenceText = useCallback((reference: string) => {
    const current = workbookRef.current
    if (!current) return false
    const match = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(reference.trim().replace(/^=/, ''))
    if (!match) return false
    const sheetName = match[1]?.replace(/''/g, "'") ?? match[2]
    const target = sheetName ? current.sheets.find((sheet) => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()) : current.sheets.find((sheet) => sheet.id === current.activeSheetId)
    if (!target) return false
    const start = { row: Number(match[4]) - 1, col: columnIndex(match[3]) }
    const end = match[5] ? { row: Number(match[6]) - 1, col: columnIndex(match[5]) } : start
    if (target.id !== current.activeSheetId) setWorkbook((value) => value ? { ...value, activeSheetId: target.id } : value)
    setSelection({ anchor: start, focus: end })
    setEditing(null)
    return true
  }, [])

  const saveDefinedNames = useCallback((drafts: NameDraft[]) => {
    const current = workbookRef.current
    if (!current) return
    mutateWorkbook((next) => {
      const builtIns = (next.definedNames || []).filter((item) => /^_xlnm\./i.test(item.name))
      next.definedNames = [...builtIns, ...drafts.map((draft) => {
        const scopeSheet = draft.scope === null ? undefined : next.sheets[draft.scope]
        return {
          name: draft.name,
          ranges: [draft.refersTo.replace(/^=/, '')],
          ...(scopeSheet ? { localSheetIndex: draft.scope!, localSheetRefId: scopeSheet.id } : {}),
          ...(draft.comment ? { comment: draft.comment } : {}),
        }
      })]
      if (next.metadata?.definedNames) delete next.metadata.definedNames
    })
    setNameManagerOpen(false)
    setToast(`Saved ${drafts.length} ${drafts.length === 1 ? 'name' : 'names'}`)
  }, [mutateWorkbook])

  const commitNameBox = useCallback((raw: string) => {
    const sheet = activeSheet
    const text = raw.trim()
    const parts = text.split(':')
    const start = coordOf(parts[0] || '')
    const end = parts.length === 2 ? coordOf(parts[1] || '') : start
    if (sheet && text && (!start || !end || parts.length > 2)) {
      // A defined name or table jumps to its range; a new valid name names the selection.
      const current = workbookRef.current
      const wanted = text.toLocaleLowerCase()
      const sheetIndex = current ? current.sheets.findIndex((item) => item.id === current.activeSheetId) : -1
      const names = current?.definedNames || []
      const defined = names.find((item) => item.name.toLocaleLowerCase() === wanted && (item.localSheetIndex === undefined || item.localSheetIndex === sheetIndex))
        || current?.metadata?.definedNames?.find((item) => item.name.toLocaleLowerCase() === wanted)
      const definedRef = defined ? ('ranges' in defined && Array.isArray(defined.ranges) ? defined.ranges[0] : (defined as { ranges?: string; formula?: string }).ranges || (defined as { formula?: string }).formula) : undefined
      if (typeof definedRef === 'string' && selectReferenceText(definedRef)) {
        setNameBoxDraft(null)
        setNameBoxInvalid(false)
        return true
      }
      const table = current?.sheets.flatMap((item) => (item.tables || []).map((entry) => ({ sheet: item, table: entry }))).find((entry) => entry.table.name.toLocaleLowerCase() === wanted)
      if (table && selectReferenceText(`'${table.sheet.name.replace(/'/g, "''")}'!${table.table.ref}`)) {
        setNameBoxDraft(null)
        setNameBoxInvalid(false)
        return true
      }
      const existing: NameDraft[] = names.map((item) => ({ name: item.name, refersTo: item.ranges?.[0] || '', scope: item.localSheetIndex ?? null }))
      if (!validateDefinedName(text, existing, null, null)) {
        const refersTo = selectionReferenceText()
        mutateWorkbook((next) => {
          next.definedNames = [...(next.definedNames || []), { name: text, ranges: [refersTo] }]
        })
        setToast(`Named ${refersTo} as ${text}`)
        setNameBoxDraft(null)
        setNameBoxInvalid(false)
        return true
      }
    }
    if (!sheet || parts.length > 2 || !start || !end || Math.min(start.row, end.row) < 0) {
      setNameBoxInvalid(true)
      return false
    }
    const maxRow = clamp(Math.max(DEFAULT_ROWS, sheet.rowCount + 25), DEFAULT_ROWS, 1_048_576) - 1
    const maxCol = clamp(Math.max(DEFAULT_COLS, sheet.colCount + 8), DEFAULT_COLS, 16_384) - 1
    setSelection({
      anchor: { row: clamp(start.row, 0, maxRow), col: clamp(start.col, 0, maxCol) },
      focus: { row: clamp(end.row, 0, maxRow), col: clamp(end.col, 0, maxCol) },
    })
    setEditing(null)
    setNameBoxDraft(null)
    setNameBoxInvalid(false)
    return true
  }, [activeSheet, mutateWorkbook, selectReferenceText, selectionReferenceText])

  useEffect(() => {
    if (!nameBoxInvalid) return
    const timeout = window.setTimeout(() => setNameBoxInvalid(false), 700)
    return () => window.clearTimeout(timeout)
  }, [nameBoxInvalid])

  // ---- Formula editing: autocomplete, argument hints, range finder, pointing, F4 ----------

  const activeFormulaDraft = formulaInput === 'cell' ? editing?.draft ?? null : formulaInput === 'bar' ? formulaDraft : null

  const updateAnchorFromInput = useCallback(() => {
    const element = formulaInputElementRef.current
    if (!element) { setAssistAnchor(null); return }
    const rect = element.getBoundingClientRect()
    setAssistAnchor((current) => current && current.left === rect.left && current.top === rect.top && current.bottom === rect.bottom && current.width === rect.width
      ? current
      : { left: rect.left, top: rect.top, bottom: rect.bottom, width: rect.width })
  }, [])

  const handleFormulaCaret = useCallback((start: number, _end: number) => {
    assistCaretRef.current = start
    setAssistCaret(start)
    updateAnchorFromInput()
  }, [updateAnchorFromInput])

  const handleEditorFocusChange = useCallback((focused: boolean, element: HTMLTextAreaElement | null) => {
    if (focused) {
      formulaInputRef.current = 'cell'
      formulaInputElementRef.current = element
      setFormulaInput('cell')
      if (element) handleFormulaCaret(element.selectionStart, element.selectionEnd)
    } else if (formulaInputRef.current === 'cell') {
      formulaInputRef.current = null
      formulaInputElementRef.current = null
      setFormulaInput(null)
      setAssistAnchor(null)
    }
  }, [handleFormulaCaret])

  const definedNameList = useMemo(() => {
    const names = new Set<string>()
    for (const entry of workbook?.definedNames || []) if (!entry.hidden && !/^_xlnm\./i.test(entry.name)) names.add(entry.name)
    for (const entry of workbook?.metadata?.definedNames || []) if (!/^_xlnm\./i.test(entry.name)) names.add(entry.name)
    return [...names].sort((a, b) => a.localeCompare(b))
  }, [workbook?.definedNames, workbook?.metadata?.definedNames])
  const tableNameList = useMemo(() => (workbook?.sheets || []).flatMap((sheet) => (sheet.tables || []).map((table) => table.name)), [workbook?.sheets])

  const assistState = useMemo<FormulaAssistState>(() => {
    const empty: FormulaAssistState = { items: [], activeIndex: 0, call: null }
    if (!activeFormulaDraft?.startsWith('=')) return empty
    const completion = assistDismissed ? null : completionContextAt(activeFormulaDraft, assistCaret)
    let items: AssistItem[] = []
    if (completion) {
      const prefix = completion.prefix.toUpperCase()
      items = searchFunctions(completion.prefix, 14).map((info) => ({ kind: 'function' as const, label: info.name, detail: info.description, insertText: `${info.name}(` }))
      for (const name of definedNameList) if (name.toUpperCase().startsWith(prefix)) items.push({ kind: 'name', label: name, detail: 'Defined name', insertText: name })
      for (const name of tableNameList) if (name.toUpperCase().startsWith(prefix)) items.push({ kind: 'table', label: name, detail: 'Table', insertText: `${name}[` })
      items = items.slice(0, 24)
      // A fully typed name that is already a complete function call needs no list.
      if (items.length === 1 && items[0].label.toUpperCase() === prefix && items[0].kind !== 'function') items = []
    }
    const call = callContextAt(activeFormulaDraft, assistCaret)
    const info = call ? getFunctionInfo(call.name) : undefined
    return {
      items,
      activeIndex: items.length ? clamp(assistIndex, 0, items.length - 1) : 0,
      call: info && call ? {
        fn: { name: info.name, description: info.description, signature: functionSignature(info), args: info.args, category: String(info.category) },
        argumentIndex: call.argumentIndex,
      } : null,
    }
  }, [activeFormulaDraft, assistCaret, assistDismissed, assistIndex, definedNameList, tableNameList])

  useEffect(() => { setAssistIndex(0) }, [assistState.items.length, activeFormulaDraft])
  useEffect(() => {
    // Unmounting a focused editor fires no blur; drop its assist state explicitly.
    if (editing || formulaInputRef.current !== 'cell') return
    formulaInputRef.current = null
    formulaInputElementRef.current = null
    pointRef.current = null
    setFormulaInput(null)
    setAssistAnchor(null)
  }, [editing])

  /** Replace the active formula input's text and caret (cell editor or formula bar). */
  const updateActiveFormulaDraft = useCallback((text: string, start: number, end = start) => {
    const input = formulaInputRef.current
    if (input === 'cell') {
      setEditing((current) => current ? { ...current, draft: text } : current)
      selectionRequestIdRef.current += 1
      setEditorSelectionRequest({ start, end, id: selectionRequestIdRef.current })
    } else if (input === 'bar') {
      setFormulaDraft(text)
      const sheet = workbookRef.current?.sheets.find((item) => item.id === workbookRef.current?.activeSheetId)
      if (sheet) formulaBarDirtyRef.current = { sheetId: sheet.id, address: mergeMasterAddress(sheet, addressOf(selectionRef.current.focus)), draft: text }
      formulaBarSelectionRef.current = { start, end }
    }
    assistCaretRef.current = start
    setAssistCaret(start)
  }, [])

  useLayoutEffect(() => {
    const request = formulaBarSelectionRef.current
    const element = formulaBarRef.current
    if (!request || !element || document.activeElement !== element) return
    formulaBarSelectionRef.current = null
    element.setSelectionRange(request.start, request.end)
  })

  const acceptAssistItem = useCallback((item: AssistItem) => {
    const input = formulaInputRef.current
    const draft = input === 'cell' ? editingRef.current?.draft : input === 'bar' ? formulaDraftRef.current : null
    if (!draft) return
    const context = completionContextAt(draft, assistCaretRef.current)
    if (!context) return
    // Don't double the "(" when the user already typed it.
    const insert = draft[context.end] === '(' && item.insertText.endsWith('(') ? item.insertText.slice(0, -1) : item.insertText
    const text = draft.slice(0, context.start) + insert + draft.slice(context.end)
    const caret = context.start + insert.length + (insert !== item.insertText ? 1 : 0)
    pointRef.current = null
    updateActiveFormulaDraft(text, caret)
  }, [updateActiveFormulaDraft])

  const pointCell = useCallback((coord: Coord, phase: 'down' | 'drag', extend: boolean) => {
    const input = formulaInputRef.current
    if (!input) return false
    const draft = input === 'cell' ? editingRef.current?.draft : formulaDraftRef.current
    if (!draft?.startsWith('=')) return false
    const caret = assistCaretRef.current
    const point = pointRef.current
    const replacing = Boolean(point && point.span.end === caret)
    if (phase === 'down' && !replacing && !canInsertReferenceAt(draft, caret)) return false
    if (phase === 'drag' && !point) return false
    const anchor = phase === 'drag' || (extend && point) ? point!.anchor : coord
    const bounds = {
      top: Math.min(anchor.row, coord.row),
      bottom: Math.max(anchor.row, coord.row),
      left: Math.min(anchor.col, coord.col),
      right: Math.max(anchor.col, coord.col),
    }
    const inserted = insertReference(draft, caret, bounds, replacing ? point!.span : null)
    pointRef.current = { span: inserted.span, anchor, focus: coord }
    updateActiveFormulaDraft(inserted.text, inserted.caret)
    return true
  }, [updateActiveFormulaDraft])

  const fillSelectionWithDraft = useCallback((draft: string) => {
    const bounds = selectionBounds(selectionRef.current)
    const origin = coordOf(editingRef.current?.address || '') || { row: bounds.top, col: bounds.left }
    if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > 100_000) { setToast('That selection is too large to fill at once.'); return }
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      for (let row = bounds.top; row <= bounds.bottom; row += 1) {
        for (let col = bounds.left; col <= bounds.right; col += 1) {
          const address = addressOf({ row, col })
          const text = draft.startsWith('=') ? `=${shiftFormulaReferences(normalizeFormula(draft), row - origin.row, col - origin.col)}` : draft
          const cell = parseDraft(text, sheet.cells[address])
          if (hasCellContent(cell)) sheet.cells[address] = cell
          else delete sheet.cells[address]
        }
      }
      sheet.rowCount = Math.max(sheet.rowCount, bounds.bottom + 1)
      sheet.colCount = Math.max(sheet.colCount, bounds.right + 1)
    })
    setEditing(null)
  }, [mutateWorkbook])

  /** App-level keys inside the cell editor. Returns true when the key was handled. */
  const handleEditorKeyDown = useCallback((event: KeyboardEvent<HTMLTextAreaElement>) => {
    const current = editingRef.current
    if (!current) return false
    const element = event.currentTarget
    const draft = element.value
    const caret = element.selectionStart
    if (assistState.items.length) {
      if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
        event.preventDefault()
        const delta = event.key === 'ArrowDown' ? 1 : -1
        setAssistIndex((index) => (index + delta + assistState.items.length) % assistState.items.length)
        return true
      }
      if (event.key === 'Tab' || (event.key === 'Enter' && !event.altKey && !event.ctrlKey && !event.shiftKey)) {
        event.preventDefault()
        acceptAssistItem(assistState.items[assistState.activeIndex])
        return true
      }
      if (event.key === 'Escape') {
        event.preventDefault()
        setAssistDismissed(true)
        return true
      }
    }
    if (event.key === 'F4') {
      event.preventDefault()
      const toggled = toggleAbsoluteReference(draft, caret, element.selectionEnd)
      if (toggled) {
        pointRef.current = null
        updateActiveFormulaDraft(toggled.text, toggled.selectionStart, toggled.selectionEnd)
      }
      return true
    }
    if (event.key === 'F2') {
      event.preventDefault()
      setEditing((state) => state ? { ...state, mode: state.mode === 'edit' ? 'enter' : 'edit' } : state)
      return true
    }
    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey) && !event.altKey) {
      event.preventDefault()
      fillSelectionWithDraft(draft)
      return true
    }
    const arrow = event.key === 'ArrowUp' ? [-1, 0] : event.key === 'ArrowDown' ? [1, 0] : event.key === 'ArrowLeft' ? [0, -1] : event.key === 'ArrowRight' ? [0, 1] : null
    if (arrow && current.mode !== 'edit' && !event.altKey) {
      const point = pointRef.current
      const replacing = Boolean(point && point.span.end === caret)
      if (draft.startsWith('=') && (replacing || canInsertReferenceAt(draft, caret))) {
        event.preventDefault()
        const origin = replacing ? point!.focus : (coordOf(current.address) || selectionRef.current.focus)
        const focus = { row: clamp(origin.row + arrow[0], 0, 1_048_575), col: clamp(origin.col + arrow[1], 0, 16_383) }
        const anchor = event.shiftKey ? (replacing ? point!.anchor : origin) : focus
        const bounds = { top: Math.min(anchor.row, focus.row), bottom: Math.max(anchor.row, focus.row), left: Math.min(anchor.col, focus.col), right: Math.max(anchor.col, focus.col) }
        const inserted = insertReference(draft, caret, bounds, replacing ? point!.span : null)
        pointRef.current = { span: inserted.span, anchor, focus }
        updateActiveFormulaDraft(inserted.text, inserted.caret)
        return true
      }
      if (current.mode === 'enter' && !event.ctrlKey && !event.shiftKey) {
        event.preventDefault()
        commitCell(current.address, draft)
        moveSelectionRef.current(arrow[0], arrow[1])
        document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })
        return true
      }
    }
    return false
  }, [acceptAssistItem, assistState, commitCell, fillSelectionWithDraft, updateActiveFormulaDraft])

  const referenceHighlights = useMemo<ReferenceHighlight[]>(() => {
    if (!activeFormulaDraft?.startsWith('=') || !activeSheet) return []
    const sheetName = activeSheet.name.toLocaleLowerCase()
    return formulaReferences(activeFormulaDraft)
      .filter((reference) => !reference.sheet || reference.sheet.toLocaleLowerCase() === sheetName)
      .map((reference) => ({ top: reference.top, left: reference.left, bottom: reference.bottom, right: reference.right, color: reference.color }))
  }, [activeFormulaDraft, activeSheet])
  const editorParts = useMemo(() => (editing?.draft.startsWith('=') ? formulaHighlightParts(editing.draft) : null), [editing?.draft])
  const formulaBarParts = useMemo(() => (formulaDraft.startsWith('=') ? formulaHighlightParts(formulaDraft) : null), [formulaDraft])
  const spillOutline = useMemo(() => {
    if (!calcEngine || !activeSheet) return null
    const anchor = calcEngine.spillAnchorOf(activeSheet.id, activeAddress)
    return anchor ? calcEngine.spillRange(activeSheet.id, anchor) : null
  }, [activeAddress, activeSheet, calcEngine])
  const spillGhostFormula = useMemo(() => {
    if (!calcEngine || !activeSheet || activeCell?.formula || (activeCell?.value !== undefined && activeCell?.value !== null && !activeCell.arrayMember)) return ''
    const anchor = calcEngine.spillAnchorOf(activeSheet.id, activeAddress)
    const formula = anchor && anchor !== activeAddress ? activeSheet.cells[anchor]?.formula : undefined
    return formula ? `=${formula}` : ''
  }, [activeAddress, activeCell, activeSheet, calcEngine])

  const handleGridKeyDown = useCallback((event: KeyboardEvent<HTMLDivElement>) => {
    if (editing) return
    const control = event.ctrlKey || event.metaKey
    if (event.altKey && event.key === 'Enter' && activeCell?.hyperlink) { event.preventDefault(); openHyperlink(activeCell.hyperlink); return }
    if (event.key === ' ' && (control || event.shiftKey)) {
      event.preventDefault()
      const bounds = selectionBounds(selection)
      const lastRow = Math.max(0, (activeSheet?.rowCount || 1) - 1)
      const lastCol = Math.max(0, (activeSheet?.colCount || 1) - 1)
      if (control && event.shiftKey) setSelection({ anchor: { row: 0, col: 0 }, focus: { row: lastRow, col: lastCol } })
      else if (event.shiftKey) setSelection({ anchor: { row: bounds.top, col: 0 }, focus: { row: bounds.bottom, col: lastCol } })
      else setSelection({ anchor: { row: 0, col: bounds.left }, focus: { row: lastRow, col: bounds.right } })
      return
    }
    if (event.key === ' ' && activeSheet && activeCell) {
      const options = validationListOptions(validationForCell(activeSheet, activeAddress))
      const checkbox = activeCell.type === 'checkbox' || (options[0]?.toLocaleUpperCase() === 'TRUE' && options[1]?.toLocaleUpperCase() === 'FALSE')
      if (checkbox) { event.preventDefault(); setCellValue(activeAddress, !(activeCell.value === true || String(activeCell.value).toLocaleUpperCase() === 'TRUE')); return }
    }
    const code = event.code
    if (control && !event.shiftKey && (event.key === ';')) {
      event.preventDefault()
      const today = new Date()
      beginEdit(`${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`)
      return
    }
    if (control && event.shiftKey && (event.key === ':' || code === 'Semicolon')) {
      event.preventDefault()
      const now = new Date()
      const hours = now.getHours() % 12 || 12
      beginEdit(`${hours}:${String(now.getMinutes()).padStart(2, '0')} ${now.getHours() < 12 ? 'AM' : 'PM'}`)
      return
    }
    if (event.altKey && (event.key === '=' || code === 'Equal') && !control) { event.preventDefault(); insertFunction('SUM'); return }
    if (control && event.shiftKey && /^Digit[1-6]$/.test(code)) {
      event.preventDefault()
      const formats: Record<string, string> = { Digit1: '#,##0.00', Digit2: 'h:mm AM/PM', Digit3: 'd-mmm-yy', Digit4: '$#,##0.00', Digit5: '0%', Digit6: '0.00E+00' }
      setNumberFormat(formats[code])
      return
    }
    if (control && event.shiftKey && code === 'Backquote') { event.preventDefault(); setNumberFormat('General'); return }
    if (control && !event.shiftKey && code === 'Digit5') { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), strike: !activeCell?.style?.font?.strike } }); return }
    if (control && !event.shiftKey && code === 'Digit9') { event.preventDefault(); hideSelectedDimension('rows'); return }
    if (control && !event.shiftKey && code === 'Digit0') { event.preventDefault(); hideSelectedDimension('columns'); return }
    if (control && event.shiftKey && code === 'Digit9') { event.preventDefault(); unhideNearSelection('rows'); return }
    if (control && event.shiftKey && code === 'Digit0') { event.preventDefault(); unhideNearSelection('columns'); return }
    if (control && event.shiftKey && code === 'Digit7') { event.preventDefault(); setBorderPreset('outer'); return }
    if (control && event.shiftKey && code === 'Minus') { event.preventDefault(); setBorderPreset('clear'); return }
    if (control && ((event.shiftKey && code === 'Equal') || code === 'NumpadAdd')) { event.preventDefault(); openShiftDialogRef.current('insert'); return }
    if (event.altKey && event.shiftKey && !control && (event.key === 'ArrowRight' || event.key === 'ArrowLeft')) { event.preventDefault(); groupSelectionRef.current(null, event.key === 'ArrowRight' ? 1 : -1); return }
    if (control && !event.shiftKey && (code === 'Minus' || code === 'NumpadSubtract')) { event.preventDefault(); openShiftDialogRef.current('delete'); return }
    if (control && (event.key === 'PageDown' || event.key === 'PageUp')) {
      event.preventDefault()
      switchSheetBy(event.key === 'PageDown' ? 1 : -1)
      return
    }
    if (control && !event.shiftKey && (event.key === "'" || code === 'Quote')) {
      event.preventDefault()
      const above = activeSheet?.cells[addressOf({ row: Math.max(0, selection.focus.row - 1), col: selection.focus.col })]
      beginEdit(above?.formula ? `=${above.formula}` : rawCellValue(above))
      return
    }
    if (control && event.shiftKey && (event.key === '"' || code === 'Quote')) {
      event.preventDefault()
      const aboveAddress = addressOf({ row: Math.max(0, selection.focus.row - 1), col: selection.focus.col })
      const value = activeSheet ? calcEngine?.getValue(activeSheet.id, aboveAddress) : null
      beginEdit(value === null || value === undefined ? '' : String(value))
      return
    }
    if ((event.key === 'F5' || (control && event.key.toLocaleLowerCase() === 'g')) && !event.shiftKey) {
      event.preventDefault()
      const nameBox = document.querySelector<HTMLInputElement>('.name-box')
      nameBox?.focus()
      nameBox?.select()
      return
    }
    if (event.key === 'Home' && !control) {
      event.preventDefault()
      const target = { row: selection.focus.row, col: hiddenColSet.has(0) ? stepPastHidden(0, 1, 16_383, hiddenColSet) : 0 }
      setSelection((current) => ({ anchor: event.shiftKey ? current.anchor : target, focus: target }))
      return
    }
    if (control && event.key.toLocaleLowerCase() === 'd') { event.preventDefault(); fillSelectedRange('down'); return }
    if (control && event.key.toLocaleLowerCase() === 'r') { event.preventDefault(); fillSelectedRange('right'); return }
    if (control && event.key.toLocaleLowerCase() === 'c') { event.preventDefault(); void copySelection(); return }
    if (control && event.key.toLocaleLowerCase() === 'x') { event.preventDefault(); void copySelection(true); return }
    if (control && event.altKey && event.key.toLocaleLowerCase() === 'v') { event.preventDefault(); setPasteSpecialOpen(true); return }
    if (control && event.shiftKey && event.key.toLocaleLowerCase() === 'v') { event.preventDefault(); void pasteSelection(PASTE_SPECIAL_PRESETS.values); return }
    if (control && event.key.toLocaleLowerCase() === 'v') { event.preventDefault(); void pasteSelection(); return }
    if (event.key === 'Escape' && cutRange) { event.preventDefault(); setCutRange(null); internalClipboard.current = null; return }
    if (control && event.key.toLocaleLowerCase() === 'a') {
      event.preventDefault()
      selectAllSheet()
      return
    }
    if (control && event.key === 'Home') {
      event.preventDefault()
      const home = {
        row: hiddenRowSet.has(0) ? stepPastHidden(0, 1, 1_048_575, hiddenRowSet) : 0,
        col: hiddenColSet.has(0) ? stepPastHidden(0, 1, 16_383, hiddenColSet) : 0,
      }
      setSelection((current) => ({ anchor: event.shiftKey ? current.anchor : home, focus: home }))
      return
    }
    if (control && event.key === 'End') {
      event.preventDefault()
      const last = { row: 0, col: 0 }
      Object.keys(activeSheet?.cells || {}).forEach((address) => {
        const coord = coordOf(address)
        if (!coord) return
        last.row = Math.max(last.row, coord.row)
        last.col = Math.max(last.col, coord.col)
      })
      setSelection((current) => ({ anchor: event.shiftKey ? current.anchor : last, focus: last }))
      return
    }
    if (event.key === 'PageUp' || event.key === 'PageDown') {
      event.preventDefault()
      const direction = event.key === 'PageDown' ? 1 : -1
      const importedDefaults = true
      if (event.altKey) {
        const pageColumns = Math.max(1, Math.floor((event.currentTarget.clientWidth - HEADER_WIDTH) / ((importedDefaults ? IMPORTED_COL_WIDTH : DEFAULT_COL_WIDTH) * zoom)))
        moveSelection(0, direction * pageColumns, event.shiftKey)
      } else {
        const pageRows = Math.max(1, Math.floor((event.currentTarget.clientHeight - HEADER_HEIGHT) / ((importedDefaults ? IMPORTED_ROW_HEIGHT : DEFAULT_ROW_HEIGHT) * zoom)))
        moveSelection(direction * pageRows, 0, event.shiftKey)
      }
      return
    }
    if (event.key === 'F2') { event.preventDefault(); beginEdit(); return }
    if (event.key === 'Enter') { event.preventDefault(); beginEdit(); return }
    if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); clearSelection(); return }
    if (event.key === 'ArrowUp') { event.preventDefault(); if (control) jumpSelection(-1, 0, event.shiftKey); else moveSelection(-1, 0, event.shiftKey); return }
    if (event.key === 'ArrowDown') { event.preventDefault(); if (control) jumpSelection(1, 0, event.shiftKey); else moveSelection(1, 0, event.shiftKey); return }
    if (event.key === 'ArrowLeft') { event.preventDefault(); if (control) jumpSelection(0, -1, event.shiftKey); else moveSelection(0, -1, event.shiftKey); return }
    if (event.key === 'ArrowRight') { event.preventDefault(); if (control) jumpSelection(0, 1, event.shiftKey); else moveSelection(0, 1, event.shiftKey); return }
    if (event.key === 'Tab') { event.preventDefault(); moveSelection(0, event.shiftKey ? -1 : 1); return }
    if (!control && !event.altKey && event.key.length === 1) { event.preventDefault(); beginEdit(event.key) }
  }, [activeAddress, activeCell, activeSheet, applyStyle, beginEdit, calcEngine, clearSelection, copySelection, cutRange, editing, fillSelectedRange, hiddenColSet, hiddenRowSet, hideSelectedDimension, insertFunction, jumpSelection, moveSelection, openHyperlink, pasteSelection, selectAllSheet, selection, setBorderPreset, setCellValue, setNumberFormat, switchSheetBy, unhideNearSelection, zoom])

  useEffect(() => {
    const handleKey = (event: globalThis.KeyboardEvent) => {
      const control = event.ctrlKey || event.metaKey
      const typing = event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement || event.target instanceof HTMLSelectElement
      if (document.querySelector('[aria-modal="true"]')) {
        const appShortcut = control && ['o', 's', 'p', 'e', 'z', 'y', 'f', 'h', 'k', 'b', 'i', 'u'].includes(event.key.toLocaleLowerCase())
        if (appShortcut || (event.altKey && event.shiftKey && event.key === '5') || (event.shiftKey && event.key === 'F11')) event.preventDefault()
        return
      }
      if (control && event.key.toLocaleLowerCase() === 'o') { event.preventDefault(); void openWorkbook(); return }
      if (control && event.shiftKey && event.key.toLocaleLowerCase() === 'e') { event.preventDefault(); setExportOpen(true); return }
      if (control && event.key.toLocaleLowerCase() === 's') { event.preventDefault(); void saveWorkbook(event.shiftKey); return }
      if (control && event.key.toLocaleLowerCase() === 'p') { event.preventDefault(); openPrintDialog(); return }
      if (control && event.key.toLocaleLowerCase() === 'z' && !typing) { event.preventDefault(); undo(); return }
      if (control && event.key.toLocaleLowerCase() === 'y' && !typing) { event.preventDefault(); redo(); return }
      if (control && event.key.toLocaleLowerCase() === 'f') { event.preventDefault(); openSearch(); return }
      if (control && event.key.toLocaleLowerCase() === 'h') { event.preventDefault(); openReplace(); return }
      if (control && event.key.toLocaleLowerCase() === 'k' && !typing) { event.preventDefault(); void insertLink(); return }
      if (control && (event.key === '`' || event.key === '~') && !typing) { event.preventDefault(); setShowFormulas((value) => !value); return }
      if (control && event.key === '\\' && !typing) { event.preventDefault(); clearFormatting(); return }
      if (control && event.key.toLocaleLowerCase() === 'b' && !typing) { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), bold: !activeCell?.style?.font?.bold } }); return }
      if (control && event.key.toLocaleLowerCase() === 'i' && !typing) { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), italic: !activeCell?.style?.font?.italic } }); return }
      if (control && event.key.toLocaleLowerCase() === 'u' && !typing) { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), underline: !activeCell?.style?.font?.underline } }); return }
      if (event.altKey && event.shiftKey && event.key === '5' && !typing) { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), strike: !activeCell?.style?.font?.strike } }); return }
      if (event.ctrlKey && event.altKey && event.key.toLocaleLowerCase() === 'm' && !typing) { event.preventDefault(); void editAnnotation('comment'); return }
      if (event.shiftKey && event.key === 'F2' && !typing) { event.preventDefault(); void editAnnotation('note'); return }
      if (event.shiftKey && event.key === 'F11' && !typing) { event.preventDefault(); addSheet(); return }
      if (control && event.key === 'F3') { event.preventDefault(); setNameManagerOpen(true); return }
      if (control && event.shiftKey && event.key.toLocaleLowerCase() === 'l' && !typing) { event.preventDefault(); toggleFilterRef.current(); return }
      if (event.altKey && event.key === 'F5' && !typing) { event.preventDefault(); refreshPivotsRef.current(control); return }
      if (control && !event.shiftKey && !event.altKey && ['t', 'l'].includes(event.key.toLocaleLowerCase()) && !typing) { event.preventDefault(); openCreateTableRef.current(); return }
      if (control && !event.shiftKey && (event.code === 'Digit1' || event.key === '1') && !typing) { event.preventDefault(); setFormatDialogTab('number'); return }
    }
    window.addEventListener('keydown', handleKey)
    return () => window.removeEventListener('keydown', handleKey)
  }, [activeCell, addSheet, applyStyle, clearFormatting, editAnnotation, insertLink, openPrintDialog, openReplace, openSearch, openWorkbook, redo, saveWorkbook, undo])

  const removeRecent = useCallback((filePath: string) => {
    setRecent((current) => current.filter((item) => item.path !== filePath))
  }, [])

  const titleBar = <TitleBar fileName={documentFile?.name} dirty={dirty} onClose={closeWindow} />
  if (!workbook || !documentFile || !activeSheet) {
    return (
      <div className="app-root">
        {titleBar}
        <Welcome recent={recent} busy={Boolean(busy)} onNew={() => { void newWorkbook() }} onOpen={() => { void openWorkbook() }} onRecent={(item) => { void openPath(item.path) }} onRemoveRecent={removeRecent} onDrop={handleDrop} />
        {busy && <div className="busy-overlay"><div className="busy-card"><span className="spinner" />{busy}</div></div>}
        {toast && <div className="toast"><span>{toast}</span><button type="button" aria-label="Dismiss" onClick={() => setToast('')}><X size={13} /></button></div>}
      </div>
    )
  }

  const activeStyle = activeCell?.style || {}
  const activeFontName = activeStyle.font?.name || String((workbook.metadata?.normalFont as { name?: string } | undefined)?.name || 'Calibri')
  const activeFontSize = Number.isFinite(activeStyle.font?.size) && Number(activeStyle.font?.size) > 0 ? Number(activeStyle.font?.size) : 11
  const fontChoices = [...new Set([...FONT_FAMILIES.map((font) => font.name), activeFontName])]
  const themePalette = activeThemeColors.slice()
  const activeValueForFormats = calcEngine?.getValue(activeSheet.id, activeAddress)
  const sampleNumber = typeof activeValueForFormats === 'number' ? activeValueForFormats : 1000.12
  const currentFormat = activeCell?.numFmt || activeStyle.numFmt || 'General'
  const numberFormatMenuItems: ToolbarMenuItem[] = [
    { label: 'Automatic', code: 'General' },
    { label: 'Plain text', code: '@' },
    { label: 'Number', code: '#,##0.00', separatorBefore: true },
    { label: 'Percent', code: '0.00%' },
    { label: 'Scientific', code: '0.00E+00' },
    { label: 'Accounting', code: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', separatorBefore: true },
    { label: 'Financial', code: '#,##0.00;(#,##0.00)' },
    { label: 'Currency', code: '$#,##0.00' },
    { label: 'Currency rounded', code: '$#,##0' },
    { label: 'Date', code: 'm/d/yyyy', separatorBefore: true },
    { label: 'Time', code: 'h:mm:ss AM/PM' },
    { label: 'Date time', code: 'm/d/yyyy h:mm:ss' },
    { label: 'Duration', code: '[h]:mm:ss' },
  ].map((entry) => ({
    id: `number-format-${entry.label.toLocaleLowerCase().replace(/\s+/g, '-')}`,
    label: `${entry.label}   ${entry.code === '@' ? '' : formatScalar(entry.code.startsWith('m/d') || entry.code.includes('h:') ? 45306.5 : sampleNumber, entry.code)}`,
    checked: currentFormat === entry.code,
    separatorBefore: entry.separatorBefore,
    action: () => setNumberFormat(entry.code),
  }))
  numberFormatMenuItems.push({ id: 'number-format-custom', label: 'Custom number format…', separatorBefore: true, action: () => setFormatDialogTab('number') })
  const fontSizeChoices = [...new Set([8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, activeFontSize])].sort((a, b) => a - b)
  const bounds = selectionBounds(selection)
  const selectedRowCount = bounds.bottom - bounds.top + 1
  const selectedColumnCount = bounds.right - bounds.left + 1
  const hiddenSheets = workbook.sheets.filter((sheet) => sheet.state === 'hidden')
  const menuDefinitions: SpreadsheetMenuDefinition[] = [
    {
      id: 'file',
      label: 'File',
      items: [
        { id: 'file-new', label: 'New spreadsheet', icon: <FilePlus2 size={13} />, action: () => { void newWorkbook() } },
        { id: 'file-open', label: 'Open…', shortcut: 'Ctrl+O', icon: <FolderOpen size={13} />, action: () => { void openWorkbook() } },
        { id: 'file-save', label: 'Save', shortcut: 'Ctrl+S', icon: <Save size={13} />, action: () => { void saveWorkbook(false) } },
        { id: 'file-save-as', label: 'Save as…', shortcut: 'Ctrl+Shift+S', action: () => { void saveWorkbook(true) } },
        { id: 'file-export', label: 'Export As…', shortcut: 'Ctrl+Shift+E', icon: <Download size={13} />, action: () => setExportOpen(true) },
        { id: 'file-print', label: 'Print…', shortcut: 'Ctrl+P', separatorBefore: true, icon: <Printer size={13} />, action: openPrintDialog },
      ],
    },
    {
      id: 'view',
      label: 'View',
      items: [
        {
          id: 'view-show', label: 'Show', icon: <Menu size={13} />, children: [
            { id: 'view-show-formulas', label: 'Formulas', shortcut: 'Ctrl+~', checked: showFormulas, action: () => setShowFormulas((value) => !value) },
            { id: 'view-show-formula-bar', label: 'Formula bar', checked: showFormulaBar, action: () => setShowFormulaBar((value) => !value) },
            { id: 'view-show-gridlines', label: 'Gridlines', checked: gridlinesVisible, action: toggleGridlines },
            { id: 'view-show-notes', label: 'Note indicators', checked: showNotes, action: () => setShowNotes((value) => !value) },
          ],
        },
        {
          id: 'view-freeze', label: 'Freeze', icon: <Square size={13} />, children: [
            {
              id: 'freeze-panes',
              label: (activeSheet.frozen?.rows || activeSheet.frozen?.columns) ? 'Unfreeze panes' : `Freeze panes at ${addressOf(selection.focus)}`,
              action: () => ((activeSheet.frozen?.rows || activeSheet.frozen?.columns) ? setFreeze('both', 0, 0) : setFreeze('both', selection.focus.row, selection.focus.col)),
            },
            { id: 'freeze-top-row', label: 'Freeze top row', action: () => setFreeze('both', 1, 0) },
            { id: 'freeze-first-column', label: 'Freeze first column', action: () => setFreeze('both', 0, 1) },
            {
              separatorBefore: true,
              id: 'view-freeze-rows', label: 'Rows', children: [
                { id: 'freeze-no-rows', label: 'No rows', checked: !(activeSheet.frozen?.rows || 0), action: () => setFreeze('rows', 0) },
                { id: 'freeze-one-row', label: '1 row', checked: activeSheet.frozen?.rows === 1, action: () => setFreeze('rows', 1) },
                { id: 'freeze-to-row', label: `Up to row ${selection.focus.row + 1}`, checked: activeSheet.frozen?.rows === selection.focus.row + 1, action: () => setFreeze('rows', selection.focus.row + 1) },
              ],
            },
            {
              id: 'view-freeze-columns', label: 'Columns', children: [
                { id: 'freeze-no-columns', label: 'No columns', checked: !(activeSheet.frozen?.columns || 0), action: () => setFreeze('columns', 0) },
                { id: 'freeze-one-column', label: '1 column', checked: activeSheet.frozen?.columns === 1, action: () => setFreeze('columns', 1) },
                { id: 'freeze-to-column', label: `Up to column ${columnName(selection.focus.col)}`, checked: activeSheet.frozen?.columns === selection.focus.col + 1, action: () => setFreeze('columns', selection.focus.col + 1) },
              ],
            },
          ],
        },
        {
          id: 'view-hidden-sheets', label: 'Hidden sheets', icon: <EyeOff size={13} />, children: hiddenSheets.length
            ? hiddenSheets.map((sheet) => ({ id: `unhide-${sheet.id}`, label: `Show ${sheet.name}`, action: () => unhideSheet(sheet.id) }))
            : [{ id: 'no-hidden-sheets', label: 'No hidden sheets', disabled: true }],
        },
        { id: 'view-hide-sheet', label: 'Hide current sheet', action: hideActiveSheet },
        { id: 'view-fit-width', label: 'Fit data width', action: () => fitSheetView('width') },
        { id: 'view-fit-sheet', label: 'Fit sheet', action: () => fitSheetView('sheet') },
        {
          id: 'view-zoom', label: 'Zoom', separatorBefore: true, children: [50, 75, 90, 100, 125, 150, 200].map((percent) => ({
            id: `zoom-${percent}`,
            label: `${percent}%`,
            checked: Math.round(zoom * 100) === percent,
            action: () => setZoom(percent / 100),
          })),
        },
        { id: 'view-fullscreen', label: immersive ? 'Exit full screen' : 'Full screen', shortcut: 'Esc to exit', icon: <Maximize2 size={13} />, action: () => { void toggleFullscreen() } },
      ],
    },
    {
      id: 'insert',
      label: 'Insert',
      items: [
        {
          id: 'insert-rows', label: 'Rows', icon: <Rows3 size={13} />, children: [
            { id: 'insert-rows-above', label: `Insert ${selectedRowCount === 1 ? 'row' : `${selectedRowCount} rows`} above`, action: () => applyStructureCommand('insert-rows-above') },
            { id: 'insert-rows-below', label: `Insert ${selectedRowCount === 1 ? 'row' : `${selectedRowCount} rows`} below`, action: () => applyStructureCommand('insert-rows-below') },
            { id: 'delete-rows', label: `Delete selected ${selectedRowCount === 1 ? 'row' : 'rows'}`, action: () => applyStructureCommand('delete-rows'), separatorBefore: true },
          ],
        },
        {
          id: 'insert-columns', label: 'Columns', icon: <Columns3 size={13} />, children: [
            { id: 'insert-columns-left', label: `Insert ${selectedColumnCount === 1 ? 'column' : `${selectedColumnCount} columns`} left`, action: () => applyStructureCommand('insert-columns-left') },
            { id: 'insert-columns-right', label: `Insert ${selectedColumnCount === 1 ? 'column' : `${selectedColumnCount} columns`} right`, action: () => applyStructureCommand('insert-columns-right') },
            { id: 'delete-columns', label: `Delete selected ${selectedColumnCount === 1 ? 'column' : 'columns'}`, action: () => applyStructureCommand('delete-columns'), separatorBefore: true },
          ],
        },
        { id: 'insert-cells', label: 'Cells…', shortcut: 'Ctrl+Shift+=', icon: <Grid2X2 size={13} />, action: () => openShiftDialog('insert') },
        { id: 'insert-sheet', label: 'Sheet', shortcut: 'Shift+F11', icon: <FilePlus2 size={13} />, action: addSheet },
        {
          id: 'insert-function', label: 'Function', separatorBefore: true, icon: <FunctionSquare size={13} />, children: ['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'IF'].map((name) => ({ id: `function-${name.toLocaleLowerCase()}`, label: name, action: () => insertFunction(name) })),
        },
        { id: 'insert-table', label: 'Table', shortcut: 'Ctrl+T', icon: <Table2 size={13} />, action: () => openCreateTable() },
        { id: 'insert-pivot', label: 'Pivot table…', icon: <TableProperties size={13} />, action: openCreatePivot },
        { id: 'insert-chart', label: 'Chart', icon: <BarChart3 size={13} />, action: insertChart },
        { id: 'insert-image', label: 'Image…', icon: <ImagePlus size={13} />, action: openPicturePicker },
        {
          id: 'insert-sparklines', label: 'Sparklines', icon: <ChartSpline size={13} />, children: [
            { id: 'sparkline-line', label: 'Line…', action: () => openCreateSparklines('line') },
            { id: 'sparkline-column', label: 'Column…', action: () => openCreateSparklines('column') },
            { id: 'sparkline-winloss', label: 'Win/Loss…', action: () => openCreateSparklines('stacked') },
          ],
        },
        { id: 'insert-link', label: 'Link', shortcut: 'Ctrl+K', icon: <Link size={13} />, action: () => { void insertLink() } },
        { id: 'insert-checkbox', label: 'Checkbox', separatorBefore: true, icon: <CheckSquare size={13} />, action: insertCheckboxes },
        { id: 'insert-dropdown', label: 'Dropdown', icon: <ListPlus size={13} />, action: () => { void insertDropdown() } },
        { id: 'insert-comment', label: 'Comment', shortcut: 'Ctrl+Alt+M', separatorBefore: true, icon: <MessageSquarePlus size={13} />, action: () => { void editAnnotation('comment') } },
        { id: 'insert-note', label: 'Note', shortcut: 'Shift+F2', icon: <StickyNote size={13} />, action: () => { void editAnnotation('note') } },
      ],
    },
    {
      id: 'format',
      label: 'Format',
      items: [
        { id: 'format-cells', label: 'Format cells…', shortcut: 'Ctrl+1', icon: <SlidersHorizontal size={13} />, action: () => setFormatDialogTab('number') },
        {
          id: 'format-number', label: 'Number', icon: <DollarSign size={13} />, children: [
            { id: 'number-general', label: 'Automatic', checked: (activeCell?.numFmt || activeStyle.numFmt || 'General') === 'General', action: () => setNumberFormat('General') },
            { id: 'number-number', label: 'Number', action: () => setNumberFormat('#,##0.00') },
            { id: 'number-percent', label: 'Percent', action: () => setNumberFormat('0.00%') },
            { id: 'number-currency', label: 'Currency', action: () => setNumberFormat('$#,##0.00') },
            { id: 'number-scientific', label: 'Scientific', action: () => setNumberFormat('0.00E+00') },
            { id: 'number-date', label: 'Date', action: () => setNumberFormat('m/d/yyyy') },
            { id: 'number-time', label: 'Time', action: () => setNumberFormat('h:mm:ss AM/PM') },
            { id: 'number-less-decimal', label: 'Decrease decimal places', action: () => adjustDecimals(-1), separatorBefore: true },
            { id: 'number-more-decimal', label: 'Increase decimal places', action: () => adjustDecimals(1) },
            { id: 'number-custom', label: 'Custom number format…', separatorBefore: true, action: () => setFormatDialogTab('number') },
          ],
        },
        {
          id: 'format-text', label: 'Text', icon: <Bold size={13} />, children: [
            { id: 'text-bold', label: 'Bold', shortcut: 'Ctrl+B', checked: Boolean(activeStyle.font?.bold), action: () => applyStyle((style) => { style.font = { ...(style.font || {}), bold: !activeStyle.font?.bold } }) },
            { id: 'text-italic', label: 'Italic', shortcut: 'Ctrl+I', checked: Boolean(activeStyle.font?.italic), action: () => applyStyle((style) => { style.font = { ...(style.font || {}), italic: !activeStyle.font?.italic } }) },
            { id: 'text-underline', label: 'Underline', shortcut: 'Ctrl+U', checked: Boolean(activeStyle.font?.underline), action: () => applyStyle((style) => { style.font = { ...(style.font || {}), underline: !activeStyle.font?.underline } }) },
            { id: 'text-strike', label: 'Strikethrough', shortcut: 'Alt+Shift+5', checked: Boolean(activeStyle.font?.strike), action: () => applyStyle((style) => { style.font = { ...(style.font || {}), strike: !activeStyle.font?.strike } }) },
          ],
        },
        {
          id: 'format-alignment', label: 'Alignment', icon: <AlignLeft size={13} />, children: [
            { id: 'align-left', label: 'Left', checked: activeStyle.alignment?.horizontal === 'left', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'left' } }) },
            { id: 'align-center', label: 'Center', checked: activeStyle.alignment?.horizontal === 'center', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'center' } }) },
            { id: 'align-right', label: 'Right', checked: activeStyle.alignment?.horizontal === 'right', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'right' } }) },
            { id: 'align-top', label: 'Top', separatorBefore: true, checked: activeStyle.alignment?.vertical === 'top', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), vertical: 'top' } }) },
            { id: 'align-middle', label: 'Middle', checked: activeStyle.alignment?.vertical === 'middle', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), vertical: 'middle' } }) },
            { id: 'align-bottom', label: 'Bottom', checked: activeStyle.alignment?.vertical === 'bottom', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), vertical: 'bottom' } }) },
          ],
        },
        {
          id: 'format-wrapping', label: 'Wrapping', icon: <WrapText size={13} />, children: [
            { id: 'wrap-overflow', label: 'Overflow', checked: !activeStyle.alignment?.wrapText && !activeStyle.alignment?.clipText, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), wrapText: false, clipText: false } }) },
            { id: 'wrap-wrap', label: 'Wrap', checked: Boolean(activeStyle.alignment?.wrapText), action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), wrapText: true, clipText: false } }) },
            { id: 'wrap-clip', label: 'Clip', checked: Boolean(activeStyle.alignment?.clipText), action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), wrapText: false, clipText: true } }) },
          ],
        },
        {
          id: 'format-rotation', label: 'Rotation', icon: <RotateCw size={13} />, children: [
            { id: 'rotate-none', label: 'None', checked: !activeStyle.alignment?.textRotation, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: 0 } }) },
            { id: 'rotate-up', label: 'Tilt up', checked: activeStyle.alignment?.textRotation === 45, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: 45 } }) },
            { id: 'rotate-down', label: 'Tilt down', checked: activeStyle.alignment?.textRotation === -45, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: -45 } }) },
            { id: 'rotate-vertical', label: 'Stack vertically', checked: activeStyle.alignment?.textRotation === 'vertical', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: 'vertical' } }) },
          ],
        },
        {
          id: 'format-font-size', label: 'Font size', children: fontSizeChoices.map((size) => ({ id: `font-size-${size}`, label: `${size}`, checked: activeFontSize === size, action: () => applyStyle((style) => { style.font = { ...(style.font || {}), size } }) })),
        },
        {
          id: 'format-merge', label: 'Merge cells', separatorBefore: true, icon: <TableCellsMerge size={13} />, children: [
            { id: 'merge-all', label: 'Merge all', action: mergeSelection },
            { id: 'merge-center-menu', label: 'Merge & center', action: () => mergeSelectionBy('center') },
            { id: 'merge-across-menu', label: 'Merge horizontally', action: () => mergeSelectionBy('across') },
            { id: 'merge-down-menu', label: 'Merge vertically', action: () => mergeSelectionBy('down') },
            { id: 'unmerge', label: 'Unmerge', icon: <TableCellsSplit size={13} />, action: unmergeSelection },
          ],
        },
        {
          id: 'format-borders', label: 'Borders', icon: <Grid2X2 size={13} />, children: [
            { id: 'borders-all', label: 'All borders', action: () => setBorderPreset('all') },
            { id: 'borders-outer', label: 'Outer border', action: () => setBorderPreset('outer') },
            { id: 'borders-bottom', label: 'Bottom border', action: () => setBorderPreset('bottom') },
            { id: 'borders-inner', label: 'Inner borders', action: () => applyBorderPresetToSelection('inner', { style: 'thin', color: { argb: 'FF000000' } }) },
            { id: 'borders-thick-outer', label: 'Thick outside border', action: () => applyBorderPresetToSelection('outer', { style: 'medium', color: { argb: 'FF000000' } }) },
            { id: 'borders-double-bottom', label: 'Double bottom border', action: () => applyBorderPresetToSelection('bottom', { style: 'double', color: { argb: 'FF000000' } }) },
            { id: 'borders-clear', label: 'Clear borders', action: () => setBorderPreset('clear'), separatorBefore: true },
            { id: 'borders-more', label: 'More borders…', action: () => setFormatDialogTab('border') },
          ],
        },
        {
          id: 'format-hide', label: 'Hide & unhide', separatorBefore: true, icon: <EyeOff size={13} />, children: [
            { id: 'hide-rows', label: selectedRowCount === 1 ? `Hide row ${bounds.top + 1}` : `Hide rows ${bounds.top + 1}–${bounds.bottom + 1}`, action: () => hideSelectedDimension('rows') },
            { id: 'hide-columns', label: selectedColumnCount === 1 ? `Hide column ${columnName(bounds.left)}` : `Hide columns ${columnName(bounds.left)}–${columnName(bounds.right)}`, action: () => hideSelectedDimension('columns') },
            { id: 'unhide-rows', label: 'Unhide rows near selection', separatorBefore: true, disabled: !(activeSheet.hiddenRows || []).length, action: () => unhideNearSelection('rows') },
            { id: 'unhide-columns', label: 'Unhide columns near selection', disabled: !(activeSheet.hiddenCols || []).length, action: () => unhideNearSelection('columns') },
          ],
        },
        { id: 'format-conditional', label: 'Conditional formatting…', separatorBefore: true, icon: <Highlighter size={13} />, action: () => setConditionalPanelOpen(true) },
        { id: 'format-as-table', label: activeTable ? 'Table design…' : 'Format as table…', icon: <Table2 size={13} />, action: () => openCreateTable() },
        { id: 'format-alternating', label: 'Alternating colors', action: applyAlternatingColors },
        { id: 'format-clear', label: 'Clear formatting', shortcut: 'Ctrl+\\', icon: <Eraser size={13} />, action: clearFormatting },
        { id: 'format-lock', label: 'Lock cells', checked: selectionLocked, separatorBefore: true, icon: <Lock size={13} />, action: () => setSelectionLocked(!selectionLocked) },
      ],
    },
    {
      id: 'data',
      label: 'Data',
      items: [
        { id: 'data-named-ranges', label: 'Named ranges…', shortcut: 'Ctrl+F3', icon: <Tag size={13} />, action: () => setNameManagerOpen(true) },
        { id: 'data-sort-asc', label: 'Sort range A→Z', separatorBefore: true, icon: <ArrowDownAZ size={13} />, action: () => sortSelectionRange('asc') },
        { id: 'data-sort-desc', label: 'Sort range Z→A', icon: <ArrowUpZA size={13} />, action: () => sortSelectionRange('desc') },
        { id: 'data-sort-custom', label: 'Custom sort…', action: () => setSortDialogOpen(true) },
        { id: 'data-filter', label: activeFilter ? 'Remove filter' : 'Create a filter', shortcut: 'Ctrl+Shift+L', separatorBefore: true, icon: activeFilter ? <FilterX size={13} /> : <Filter size={13} />, action: toggleFilter },
        { id: 'data-filter-reapply', label: 'Reapply filter', disabled: !filterTargets.length, action: () => reapplyFilter(false) },
        { id: 'data-filter-clear', label: 'Clear filter criteria', disabled: !filterTargets.length, action: () => reapplyFilter(true) },
        { id: 'data-validation', label: 'Data validation…', separatorBefore: true, icon: <ListPlus size={13} />, action: () => setValidationDialogOpen(true) },
        { id: 'data-checkbox', label: 'Insert checkboxes', icon: <CheckSquare size={13} />, action: insertCheckboxes },
        { id: 'data-dropdown', label: 'Insert dropdown', action: () => { void insertDropdown() } },
        { id: 'data-circle-invalid', label: invalidCells.length ? 'Clear validation circles' : 'Circle invalid data', action: () => invalidCells.length ? setInvalidCells([]) : circleInvalidData() },
        { id: 'data-remove-validation', label: 'Remove validation', action: removeValidation },
        {
          id: 'data-cleanup', label: 'Data cleanup', separatorBefore: true, children: [
            { id: 'cleanup-duplicates', label: 'Remove duplicates…', action: () => setCleanupDialog('duplicates') },
            { id: 'cleanup-trim', label: 'Trim whitespace', action: () => applyCleanup('trim') },
            { id: 'cleanup-fill', label: 'Fill blanks from above', action: () => applyCleanup('fill-blanks') },
            { id: 'cleanup-upper', label: 'UPPERCASE', separatorBefore: true, action: () => applyCleanup('upper') },
            { id: 'cleanup-lower', label: 'lowercase', action: () => applyCleanup('lower') },
            { id: 'cleanup-proper', label: 'Proper Case', action: () => applyCleanup('proper') },
            { id: 'cleanup-sentence', label: 'Sentence case', action: () => applyCleanup('sentence') },
          ],
        },
        { id: 'data-split', label: 'Split text to columns…', action: () => setCleanupDialog('split') },
        {
          id: 'data-pivot', label: 'Pivot table', separatorBefore: true, icon: <TableProperties size={13} />, children: [
            { id: 'pivot-create', label: 'Create pivot table…', action: openCreatePivot },
            { id: 'pivot-editor', label: 'Pivot table editor', disabled: !activePivot, action: () => setPivotEditorHidden(null) },
            { id: 'pivot-refresh', label: 'Refresh', shortcut: 'Alt+F5', separatorBefore: true, disabled: !activePivot, action: () => { void refreshPivots(false) } },
            { id: 'pivot-refresh-all', label: 'Refresh all', shortcut: 'Ctrl+Alt+F5', disabled: !workbook.sheets.some((sheet) => sheet.pivots?.length), action: () => { void refreshPivots(true) } },
          ],
        },
        {
          id: 'data-group', label: 'Group and outline', children: [
            { id: 'group-rows', label: 'Group rows', shortcut: 'Alt+Shift+→', action: () => groupSelection('row', 1) },
            { id: 'group-columns', label: 'Group columns', action: () => groupSelection('column', 1) },
            { id: 'ungroup-rows', label: 'Ungroup rows', shortcut: 'Alt+Shift+←', separatorBefore: true, action: () => groupSelection('row', -1) },
            { id: 'ungroup-columns', label: 'Ungroup columns', action: () => groupSelection('column', -1) },
            { id: 'outline-collapse-all', label: 'Collapse all groups', separatorBefore: true, disabled: !outlineState, action: () => updateOutline((sheet) => { showOutlineLevel(sheet, 'row', 1); showOutlineLevel(sheet, 'column', 1) }) },
            { id: 'outline-expand-all', label: 'Expand all groups', disabled: !outlineState, action: () => updateOutline((sheet) => { showOutlineLevel(sheet, 'row', 8); showOutlineLevel(sheet, 'column', 8) }) },
            { id: 'outline-clear', label: 'Clear outline', disabled: !outlineState, action: () => updateOutline((sheet) => { clearOutline(sheet, 'row'); clearOutline(sheet, 'column') }) },
          ],
        },
        {
          id: 'data-what-if', label: 'What-if analysis', children: [
            { id: 'data-goal-seek', label: 'Goal seek…', action: () => setGoalSeekOpen(true) },
          ],
        },
      ],
    },
    {
      id: 'tools',
      label: 'Tools',
      items: [
        { id: 'tools-paste-special', label: 'Paste special…', shortcut: 'Ctrl+Alt+V', icon: <ClipboardPaste size={13} />, action: () => setPasteSpecialOpen(true) },
        activeProtection
          ? { id: 'tools-unprotect', label: 'Unprotect sheet…', icon: <LockOpen size={13} />, separatorBefore: true, action: () => { void unprotectActiveSheet() } }
          : { id: 'tools-protect', label: 'Protect sheet…', icon: <Lock size={13} />, separatorBefore: true, action: () => setProtectDialogOpen(true) },
        {
          id: 'tools-auditing', label: 'Formula auditing', children: [
            { id: 'audit-precedents', label: 'Trace precedents', action: () => traceCells('precedents') },
            { id: 'audit-dependents', label: 'Trace dependents', action: () => traceCells('dependents') },
            { id: 'audit-remove', label: 'Remove arrows', separatorBefore: true, disabled: !traceState, action: () => setTraceState(null) },
          ],
        },
        { id: 'tools-paste-values', label: 'Paste values only', shortcut: 'Ctrl+Shift+V', action: () => { void pasteSelection(PASTE_SPECIAL_PRESETS.values) } },
        { id: 'tools-find', label: 'Find in this sheet', shortcut: 'Ctrl+F', separatorBefore: true, icon: <Search size={13} />, action: openSearch },
        { id: 'tools-replace', label: 'Find and replace', shortcut: 'Ctrl+H', icon: <Replace size={13} />, action: openReplace },
        { id: 'tools-show-formulas', label: 'Show formulas', shortcut: 'Ctrl+~', checked: showFormulas, icon: <FunctionSquare size={13} />, action: () => setShowFormulas((value) => !value) },
      ],
    },
  ]
  const contextMenuItems: SpreadsheetMenuItem[] = (() => {
    if (!contextMenu) return []
    if (contextMenu.kind === 'image') {
      const imageId = contextMenu.imageId
      const image = activeSheet.images?.find((item) => item.id === imageId)
      return [
        { id: 'image-delete', label: 'Delete picture', shortcut: 'Delete', icon: <Trash2 size={13} />, action: () => { if (imageId) deleteImage(imageId) } },
        {
          id: 'image-reset', label: 'Reset size', disabled: !image, action: () => {
            if (!image) return
            void preparePicture(image.src).then((picture) => {
              updateImage({ ...image, anchor: { ...anchorForSize(activeSheet, image.anchor.from.row, image.anchor.from.col, picture.width, picture.height), from: image.anchor.from, editAs: image.anchor.editAs } })
            }).catch(() => undefined)
          },
        },
        {
          id: 'image-alt', label: 'Alt text…', separatorBefore: true, disabled: !image, action: () => {
            if (!image) return
            void askText({ title: 'Alt text', label: 'Describe the picture for screen readers', initialValue: image.altText || '' }).then((text) => {
              if (text !== null) updateImage({ ...image, altText: text.trim() || undefined })
            })
          },
        },
      ]
    }
    if (contextMenu.kind === 'status') {
      return STATUS_STATS.map((stat) => ({
        id: `status-${stat.key}`,
        label: stat.label,
        checked: statusStats.has(stat.key),
        action: () => setStatusStats((current) => {
          const next = new Set(current)
          if (next.has(stat.key)) next.delete(stat.key)
          else next.add(stat.key)
          try { localStorage.setItem(STATUS_STATS_KEY, JSON.stringify([...next])) } catch { /* Preference only. */ }
          return next
        }),
      }))
    }
    if (contextMenu.kind === 'sheet-tab') {
      const sheetId = contextMenu.sheetId || workbook.activeSheetId
      return [
        { id: 'context-rename-sheet', label: 'Rename sheet', icon: <Pencil size={13} />, action: () => { void renameSheet(sheetId) } },
        { id: 'context-duplicate-sheet', label: 'Duplicate sheet', icon: <Copy size={13} />, action: duplicateActiveSheet },
        {
          id: 'context-tab-color', label: 'Tab color', icon: <PaintBucket size={13} />, children: [
            { id: 'tab-color-none', label: 'No color', action: () => setSheetTabColor(sheetId, null) },
            ...TAB_COLORS.map((color) => ({
              id: `tab-color-${color}`,
              label: `#${color}`,
              icon: <span className="tab-color-swatch" style={{ background: `#${color}` }} />,
              checked: String((workbook.sheets.find((item) => item.id === sheetId)?.properties?.tabColor as { argb?: string } | undefined)?.argb || '').toUpperCase().endsWith(color),
              action: () => setSheetTabColor(sheetId, color),
            })),
          ],
        },
        { id: 'context-move-sheet-left', label: 'Move left', separatorBefore: true, disabled: workbook.sheets.findIndex((item) => item.id === sheetId) <= 0, action: () => moveSheet(sheetId, workbook.sheets.findIndex((item) => item.id === sheetId) - 1) },
        { id: 'context-move-sheet-right', label: 'Move right', disabled: workbook.sheets.findIndex((item) => item.id === sheetId) >= workbook.sheets.length - 1, action: () => moveSheet(sheetId, workbook.sheets.findIndex((item) => item.id === sheetId) + 1) },
        { id: 'context-hide-sheet', label: 'Hide sheet', icon: <EyeOff size={13} />, action: hideActiveSheet },
        { id: 'context-delete-sheet', label: 'Delete sheet', separatorBefore: true, icon: <Trash2 size={13} />, action: deleteActiveSheet },
      ]
    }
    const formatItems: SpreadsheetMenuItem[] = [
      { id: 'context-format-cells', label: 'Format cells…', shortcut: 'Ctrl+1', separatorBefore: true, icon: <SlidersHorizontal size={13} />, action: () => setFormatDialogTab('number') },
    ]
    const clipboardItems: SpreadsheetMenuItem[] = [
      { id: 'context-cut', label: 'Cut', shortcut: 'Ctrl+X', icon: <Scissors size={13} />, action: () => { void copySelection(true) } },
      { id: 'context-copy', label: 'Copy', shortcut: 'Ctrl+C', icon: <Copy size={13} />, action: () => { void copySelection() } },
      { id: 'context-paste', label: 'Paste', shortcut: 'Ctrl+V', icon: <ClipboardPaste size={13} />, action: () => { void pasteSelection() } },
      {
        id: 'context-paste-special', label: 'Paste special', children: PASTE_SPECIAL_MENU.map((item) => ({
          id: `context-${item.id}`,
          label: item.label,
          shortcut: item.shortcut,
          separatorBefore: item.separatorBefore,
          disabled: Boolean(item.requiresInternalSource && !internalClipboard.current),
          action: () => item.opensDialog ? setPasteSpecialOpen(true) : void pasteSelection(item.preset ? PASTE_SPECIAL_PRESETS[item.preset] : undefined),
        })),
      },
      { id: 'context-clear', label: 'Clear contents', shortcut: 'Delete', icon: <Eraser size={13} />, action: clearSelection },
      { id: 'context-select-all', label: 'Select all', shortcut: 'Ctrl+A', separatorBefore: true, icon: <Grid2X2 size={13} />, action: selectAllSheet },
    ]
    if (contextMenu.kind === 'corner') return clipboardItems
    const sortItems: SpreadsheetMenuItem[] = [
      { id: 'context-sort-asc', label: 'Sort range A→Z', separatorBefore: true, icon: <ArrowDownAZ size={13} />, action: () => sortSelectionRange('asc') },
      { id: 'context-sort-desc', label: 'Sort range Z→A', icon: <ArrowUpZA size={13} />, action: () => sortSelectionRange('desc') },
    ]
    const insertRowItems: SpreadsheetMenuItem[] = [
      { id: 'context-insert-rows-above', label: `Insert ${selectedRowCount === 1 ? 'row' : `${selectedRowCount} rows`} above`, separatorBefore: true, icon: <Rows3 size={13} />, action: () => applyStructureCommand('insert-rows-above') },
      { id: 'context-insert-rows-below', label: `Insert ${selectedRowCount === 1 ? 'row' : `${selectedRowCount} rows`} below`, action: () => applyStructureCommand('insert-rows-below') },
    ]
    const insertColumnItems: SpreadsheetMenuItem[] = [
      { id: 'context-insert-columns-left', label: `Insert ${selectedColumnCount === 1 ? 'column' : `${selectedColumnCount} columns`} left`, separatorBefore: contextMenu.kind === 'column-header', icon: <Columns3 size={13} />, action: () => applyStructureCommand('insert-columns-left') },
      { id: 'context-insert-columns-right', label: `Insert ${selectedColumnCount === 1 ? 'column' : `${selectedColumnCount} columns`} right`, action: () => applyStructureCommand('insert-columns-right') },
    ]
    const deleteRowItem: SpreadsheetMenuItem = { id: 'context-delete-rows', label: `Delete selected ${selectedRowCount === 1 ? 'row' : 'rows'}`, separatorBefore: true, icon: <Trash2 size={13} />, action: () => applyStructureCommand('delete-rows') }
    const deleteColumnItem: SpreadsheetMenuItem = { id: 'context-delete-columns', label: `Delete selected ${selectedColumnCount === 1 ? 'column' : 'columns'}`, icon: <Trash2 size={13} />, action: () => applyStructureCommand('delete-columns') }
    const shiftCellItems: SpreadsheetMenuItem[] = [
      { id: 'context-insert-cells', label: 'Insert cells…', shortcut: 'Ctrl+Shift+=', separatorBefore: true, action: () => openShiftDialog('insert') },
      { id: 'context-delete-cells', label: 'Delete cells…', shortcut: 'Ctrl+-', action: () => openShiftDialog('delete') },
    ]
    if (contextMenu.kind === 'row-header') {
      return [
        ...clipboardItems,
        ...insertRowItems,
        deleteRowItem,
        { id: 'context-hide-rows', label: selectedRowCount === 1 ? `Hide row ${bounds.top + 1}` : `Hide rows ${bounds.top + 1}–${bounds.bottom + 1}`, separatorBefore: true, icon: <EyeOff size={13} />, action: () => hideSelectedDimension('rows') },
        { id: 'context-unhide-rows', label: 'Unhide rows near selection', disabled: !(activeSheet.hiddenRows || []).length, action: () => unhideNearSelection('rows') },
        { id: 'context-group-rows', label: selectedRowCount === 1 ? `Group row ${bounds.top + 1}` : `Group rows ${bounds.top + 1}–${bounds.bottom + 1}`, separatorBefore: true, action: () => groupSelection('row', 1) },
        { id: 'context-ungroup-rows', label: 'Ungroup rows', disabled: !outlineState?.rowDepth, action: () => groupSelection('row', -1) },
        ...sortItems,
      ]
    }
    if (contextMenu.kind === 'column-header') {
      return [
        ...clipboardItems,
        ...insertColumnItems,
        { ...deleteColumnItem, separatorBefore: true },
        { id: 'context-hide-columns', label: selectedColumnCount === 1 ? `Hide column ${columnName(bounds.left)}` : `Hide columns ${columnName(bounds.left)}–${columnName(bounds.right)}`, separatorBefore: true, icon: <EyeOff size={13} />, action: () => hideSelectedDimension('columns') },
        { id: 'context-unhide-columns', label: 'Unhide columns near selection', disabled: !(activeSheet.hiddenCols || []).length, action: () => unhideNearSelection('columns') },
        { id: 'context-group-columns', label: selectedColumnCount === 1 ? `Group column ${columnName(bounds.left)}` : `Group columns ${columnName(bounds.left)}–${columnName(bounds.right)}`, separatorBefore: true, action: () => groupSelection('column', 1) },
        { id: 'context-ungroup-columns', label: 'Ungroup columns', disabled: !outlineState?.columnDepth, action: () => groupSelection('column', -1) },
        ...sortItems,
      ]
    }
    const sparklineItems: SpreadsheetMenuItem[] = activeSparkline ? [{
      id: 'context-sparklines', label: 'Sparklines', separatorBefore: true, icon: <ChartSpline size={13} />, children: [
        { id: 'sparkline-type-line', label: 'Line', checked: activeSparkline.group.type === 'line', action: () => updateSparklineGroup(activeSparkline.groupIndex, { type: 'line' }) },
        { id: 'sparkline-type-column', label: 'Column', checked: activeSparkline.group.type === 'column', action: () => updateSparklineGroup(activeSparkline.groupIndex, { type: 'column' }) },
        { id: 'sparkline-type-winloss', label: 'Win/Loss', checked: activeSparkline.group.type === 'stacked', action: () => updateSparklineGroup(activeSparkline.groupIndex, { type: 'stacked' }) },
        ...(['high', 'low', 'first', 'last', 'negative', 'markers'] as const).map((key, index): SpreadsheetMenuItem => ({
          id: `sparkline-flag-${key}`,
          label: key === 'markers' ? 'Markers' : key === 'negative' ? 'Negative points' : `${key[0].toUpperCase()}${key.slice(1)} point`,
          separatorBefore: index === 0,
          checked: Boolean(activeSparkline.group[key]),
          action: () => updateSparklineGroup(activeSparkline.groupIndex, { [key]: !activeSparkline.group[key] }),
        })),
        { id: 'sparkline-axis', label: 'Show axis', checked: Boolean(activeSparkline.group.displayXAxis), action: () => updateSparklineGroup(activeSparkline.groupIndex, { displayXAxis: !activeSparkline.group.displayXAxis }) },
        { id: 'sparkline-clear', label: 'Clear sparkline group', separatorBefore: true, action: () => updateSparklineGroup(activeSparkline.groupIndex, null) },
      ],
    }] : []
    const tableItems: SpreadsheetMenuItem[] = activeTable ? [{
      id: 'context-table', label: 'Table', separatorBefore: true, icon: <Table2 size={13} />, children: [
        { id: 'context-table-design', label: 'Table design…', action: () => setTableDesignOpen(true) },
        { id: 'context-table-totals', label: 'Totals row', checked: Boolean(activeTable.totalsRow), action: () => setTableOption(activeTable.id, 'totalsRow', !activeTable.totalsRow) },
        { id: 'context-table-convert', label: 'Convert to range', separatorBefore: true, action: () => convertActiveTable(activeTable.id) },
      ],
    }] : []
    return [
      ...clipboardItems,
      ...insertRowItems,
      ...insertColumnItems,
      deleteRowItem,
      deleteColumnItem,
      ...shiftCellItems,
      ...sortItems,
      ...sparklineItems,
      ...tableItems,
      ...formatItems,
    ]
  })()
  return (
    <div className={`app-root workbook-app${immersive ? ' is-fullscreen' : ''}`} onDragOver={(event) => event.preventDefault()} onDrop={handleDrop}>
      {titleBar}
      <SpreadsheetMenus menus={menuDefinitions} />
      <CommandBar
        leading={<>
          <IconButton label="New spreadsheet" onClick={() => { void newWorkbook() }}><FilePlus2 size={16} /></IconButton>
          <IconButton label="Open spreadsheet (Ctrl+O)" onClick={() => { void openWorkbook() }}><FolderOpen size={16} /></IconButton>
          <button type="button" className="save-command" onClick={() => { void saveWorkbook(false) }} title="Save (Ctrl+S)"><Save size={16} /><span>Save</span></button>
          <IconButton label="Export As (Ctrl+Shift+E) — XLSX, ODS, CSV, TSV, PDF, or HTML" onClick={() => setExportOpen(true)}><Download size={16} /></IconButton>
          <IconButton label="Print (Ctrl+P)" onClick={openPrintDialog}><Printer size={16} /></IconButton>
        </>}
        groups={[
          {
            id: 'history',
            content: <>
              <IconButton label="Undo (Ctrl+Z)" disabled={!historyRef.current.length} onClick={undo}><Undo2 size={16} /></IconButton>
              <IconButton label="Redo (Ctrl+Y)" disabled={!futureRef.current.length} onClick={redo}><Redo2 size={16} /></IconButton>
              <button
                type="button"
                className={`tool-button${formatPainter ? ' is-active' : ''}`}
                aria-label="Paint format (double-click to keep painting)"
                title="Paint format (double-click to keep painting)"
                onClick={() => { if (formatPainter) setFormatPainter(null); else startFormatPainter(false) }}
                onDoubleClick={() => startFormatPainter(true)}
              ><Paintbrush size={15} /></button>
              <IconButton label="Clear formatting (Ctrl+\\)" onClick={clearFormatting}><Eraser size={15} /></IconButton>
            </>,
          },
          {
            id: 'zoom',
            content: <select className="toolbar-zoom-select" aria-label="Zoom" value={Math.round(zoom * 100)} onChange={(event) => event.target.value === 'fit-width' ? fitSheetView('width') : event.target.value === 'fit-sheet' ? fitSheetView('sheet') : setZoom(Number(event.target.value) / 100)}>
              {[...new Set([50, 75, 90, 100, 125, 150, 200, Math.round(zoom * 100)])].sort((a, b) => a - b).map((percent) => <option key={percent} value={percent}>{percent}%</option>)}
              <option value="fit-width">Fit data width</option>
              <option value="fit-sheet">Fit sheet</option>
            </select>,
          },
          {
            id: 'number',
            content: <>
              <IconButton label="Format as currency" onClick={() => setNumberFormat('$#,##0.00')}><DollarSign size={14} /></IconButton>
              <IconButton label="Format as percent" onClick={() => setNumberFormat('0.00%')}><Percent size={14} /></IconButton>
              <IconButton label="Decrease decimal places" onClick={() => adjustDecimals(-1)}><span className="decimal-tool">.0‹</span></IconButton>
              <IconButton label="Increase decimal places" onClick={() => adjustDecimals(1)}><span className="decimal-tool">.00›</span></IconButton>
              <ToolbarMenuButton label="More number formats" icon={<span className="number-format-tool">123</span>} items={numberFormatMenuItems} />
            </>,
          },
          {
            id: 'font',
            content: <>
              <select className="font-select" aria-label="Font" value={activeFontName} style={{ fontFamily: `"${activeFontName}", Calibri, sans-serif` }} onChange={(event) => applyStyle((style) => { style.font = { ...(style.font || {}), name: event.target.value } })}>
                {fontChoices.map((font) => <option key={font} style={{ fontFamily: `"${font}", sans-serif` }}>{font}</option>)}
              </select>
              <IconButton label="Decrease font size" onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), size: clamp((style.font?.size || activeStyle.font?.size || 11) - 1, 1, 409) } })}><Minus size={13} /></IconButton>
              <input
                className="size-input"
                aria-label="Font size"
                list="font-size-options"
                key={`${activeAddress}-${activeFontSize}`}
                defaultValue={String(activeFontSize)}
                onKeyDown={(event) => {
                  if (event.key !== 'Enter') return
                  event.preventDefault()
                  const size = Number(event.currentTarget.value)
                  if (Number.isFinite(size) && size >= 1 && size <= 409) applyStyle((style) => { style.font = { ...(style.font || {}), size: Math.round(size * 2) / 2 } })
                  document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })
                }}
                onChange={(event) => {
                  const size = Number(event.target.value)
                  if (FONT_SIZES.includes(size)) applyStyle((style) => { style.font = { ...(style.font || {}), size } })
                }}
              />
              <datalist id="font-size-options">{FONT_SIZES.map((size) => <option key={size} value={size} />)}</datalist>
              <IconButton label="Increase font size" onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), size: clamp((style.font?.size || activeStyle.font?.size || 11) + 1, 1, 409) } })}><Plus size={13} /></IconButton>
            </>,
          },
          {
            id: 'text',
            content: <>
              <IconButton label="Bold (Ctrl+B)" active={Boolean(activeStyle.font?.bold)} onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), bold: !activeStyle.font?.bold } })}><Bold size={15} /></IconButton>
              <IconButton label="Italic (Ctrl+I)" active={Boolean(activeStyle.font?.italic)} onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), italic: !activeStyle.font?.italic } })}><Italic size={15} /></IconButton>
              <IconButton label="Underline (Ctrl+U)" active={Boolean(activeStyle.font?.underline)} onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), underline: !activeStyle.font?.underline } })}><Underline size={15} /></IconButton>
              <IconButton label="Strikethrough (Ctrl+5)" active={Boolean(activeStyle.font?.strike)} onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), strike: !activeStyle.font?.strike } })}><Strikethrough size={15} /></IconButton>
              <ColorDropdownButton label="Text color" mode="text" compact icon={<Type size={14} />} themeColors={themePalette} value={(activeStyle.font?.color as SpreadsheetColor | string | undefined) ?? null} onChange={(color) => applyFormatChangeToSelection({ font: { color } })} />
            </>,
          },
          {
            id: 'cell',
            content: <>
              <ColorDropdownButton label="Fill color" mode="fill" compact icon={<PaintBucket size={14} />} themeColors={themePalette} value={(activeStyle.fill?.fgColor as SpreadsheetColor | string | undefined) ?? (activeStyle.fill?.color as string | undefined) ?? null} onChange={(color) => applyFormatChangeToSelection({ fill: color ? { type: 'pattern', pattern: 'solid', fgColor: color } : null })} />
              <ToolbarMenuButton
                label="Borders"
                icon={<Grid2X2 size={15} />}
                render={(close) => (
                  <BorderPicker
                    themeColors={themePalette}
                    multiRow={selectedRowCount > 1}
                    multiColumn={selectedColumnCount > 1}
                    onApply={(preset, side) => { applyBorderPresetToSelection(preset, side); close() }}
                    onMoreBorders={() => { close(); setFormatDialogTab('border') }}
                  />
                )}
              />
              <ToolbarMenuButton
                label="Merge cells"
                icon={<TableCellsMerge size={15} />}
                onPrimary={mergeSelection}
                items={[
                  { id: 'merge-all', label: 'Merge all', icon: <TableCellsMerge size={14} />, action: mergeSelection },
                  { id: 'merge-center', label: 'Merge & center', action: () => mergeSelectionBy('center') },
                  { id: 'merge-across', label: 'Merge horizontally', action: () => mergeSelectionBy('across') },
                  { id: 'merge-down', label: 'Merge vertically', action: () => mergeSelectionBy('down') },
                  { id: 'merge-undo', label: 'Unmerge', icon: <TableCellsSplit size={14} />, separatorBefore: true, action: unmergeSelection },
                ]}
              />
            </>,
          },
          {
            id: 'align',
            content: <>
              <ToolbarMenuButton label="Horizontal align" icon={activeStyle.alignment?.horizontal === 'center' ? <AlignCenter size={15} /> : activeStyle.alignment?.horizontal === 'right' ? <AlignRight size={15} /> : <AlignLeft size={15} />} items={[
                { id: 'h-left', label: 'Left', icon: <AlignLeft size={14} />, checked: activeStyle.alignment?.horizontal === 'left', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'left' } }) },
                { id: 'h-center', label: 'Center', icon: <AlignCenter size={14} />, checked: activeStyle.alignment?.horizontal === 'center', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'center' } }) },
                { id: 'h-right', label: 'Right', icon: <AlignRight size={14} />, checked: activeStyle.alignment?.horizontal === 'right', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'right' } }) },
                { id: 'h-general', label: 'General (by data type)', separatorBefore: true, checked: !activeStyle.alignment?.horizontal || activeStyle.alignment.horizontal === 'general', action: () => applyStyle((style) => { if (style.alignment) delete style.alignment.horizontal }) },
                { id: 'h-more', label: 'More alignment…', action: () => setFormatDialogTab('alignment') },
              ]} />
              <ToolbarMenuButton label="Vertical align" icon={activeStyle.alignment?.vertical === 'top' ? <AlignVerticalJustifyStart size={15} /> : activeStyle.alignment?.vertical === 'middle' ? <AlignVerticalJustifyCenter size={15} /> : <AlignVerticalJustifyEnd size={15} />} items={[
                { id: 'v-top', label: 'Top', icon: <AlignVerticalJustifyStart size={14} />, checked: activeStyle.alignment?.vertical === 'top', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), vertical: 'top' } }) },
                { id: 'v-middle', label: 'Middle', icon: <AlignVerticalJustifyCenter size={14} />, checked: activeStyle.alignment?.vertical === 'middle', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), vertical: 'middle' } }) },
                { id: 'v-bottom', label: 'Bottom', icon: <AlignVerticalJustifyEnd size={14} />, checked: !activeStyle.alignment?.vertical || activeStyle.alignment.vertical === 'bottom', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), vertical: 'bottom' } }) },
              ]} />
              <ToolbarMenuButton label="Text wrapping" icon={<WrapText size={15} />} items={[
                { id: 'wrap-overflow', label: 'Overflow', checked: !activeStyle.alignment?.wrapText && !activeStyle.alignment?.clipText, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), wrapText: false, clipText: false } }) },
                { id: 'wrap-wrap', label: 'Wrap', checked: Boolean(activeStyle.alignment?.wrapText), action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), wrapText: true, clipText: false } }) },
                { id: 'wrap-clip', label: 'Clip', checked: Boolean(activeStyle.alignment?.clipText), action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), wrapText: false, clipText: true } }) },
              ]} />
              <ToolbarMenuButton label="Text rotation" icon={<RotateCw size={15} />} items={[
                { id: 'rotate-none', label: 'None', checked: !activeStyle.alignment?.textRotation, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: 0 } }) },
                { id: 'rotate-up', label: 'Tilt up', checked: activeStyle.alignment?.textRotation === 45, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: 45 } }) },
                { id: 'rotate-down', label: 'Tilt down', checked: activeStyle.alignment?.textRotation === -45, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: -45 } }) },
                { id: 'rotate-vertical', label: 'Stack vertically', checked: activeStyle.alignment?.textRotation === 'vertical', action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: 'vertical' } }) },
                { id: 'rotate-90', label: 'Rotate up', checked: activeStyle.alignment?.textRotation === 90, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: 90 } }) },
                { id: 'rotate-neg-90', label: 'Rotate down', checked: activeStyle.alignment?.textRotation === -90, action: () => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: -90 } }) },
                { id: 'rotate-more', label: 'Custom angle…', separatorBefore: true, action: () => setFormatDialogTab('alignment') },
              ]} />
            </>,
          },
          {
            id: 'insert',
            content: <>
              <ToolbarMenuButton label={activeTable ? `Table design — ${activeTable.name}` : 'Format as table (Ctrl+T)'} icon={<Table2 size={15} />} active={Boolean(activeTable && tableDesignOpen)} onPrimary={() => (activeTable ? setTableDesignOpen((open) => !open) : openCreateTable())} items={[
                { id: 'table-create', label: activeTable ? 'Table design…' : 'Format as table…', shortcut: 'Ctrl+T', action: () => (activeTable ? setTableDesignOpen(true) : openCreateTable()) },
                ...(activeTable ? [
                  { id: 'table-totals', label: 'Totals row', checked: Boolean(activeTable.totalsRow), separatorBefore: true, action: () => setTableOption(activeTable.id, 'totalsRow', !activeTable.totalsRow) },
                  { id: 'table-banded', label: 'Banded rows', checked: activeTable.style?.showRowStripes !== false, action: () => setTableOption(activeTable.id, 'showRowStripes', activeTable.style?.showRowStripes === false) },
                  { id: 'table-filter-buttons', label: 'Filter buttons', checked: activeTable.showFilterButton !== false && activeTable.headerRow !== false, action: () => setTableOption(activeTable.id, 'showFilterButton', activeTable.showFilterButton === false) },
                  { id: 'table-convert', label: 'Convert to range', separatorBefore: true, action: () => convertActiveTable(activeTable.id) },
                ] : []),
              ]} />
              <IconButton label="Insert chart" onClick={insertChart}><BarChart3 size={15} /></IconButton>
              <IconButton label="Insert image" onClick={openPicturePicker}><ImagePlus size={15} /></IconButton>
              <IconButton label={activeFilter ? 'Remove filter (Ctrl+Shift+L)' : 'Create a filter (Ctrl+Shift+L)'} active={Boolean(activeFilter)} onClick={toggleFilter}><Filter size={15} /></IconButton>
              <IconButton label="Sort A→Z" onClick={() => sortSelectionRange('asc')}><ArrowDownAZ size={15} /></IconButton>
              <IconButton label="Insert link (Ctrl+K)" onClick={() => { void insertLink() }}><Link size={15} /></IconButton>
              <IconButton label="Add note (Shift+F2)" onClick={() => { void editAnnotation('note') }}><StickyNote size={15} /></IconButton>
              <ToolbarMenuButton label="Functions (Alt+=)" icon={<Sigma size={15} />} onPrimary={() => insertFunction('SUM')} items={[
                ...['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((name): ToolbarMenuItem => ({ id: `fn-${name}`, label: name, action: () => insertFunction(name) })),
                ...['IF', 'SUMIF', 'COUNTIF', 'XLOOKUP', 'VLOOKUP', 'FILTER', 'UNIQUE', 'SORT', 'TEXT', 'CONCAT', 'TODAY'].map((name, index): ToolbarMenuItem => ({ id: `fn-${name}`, label: name, separatorBefore: index === 0, action: () => beginEdit(`=${name}(`) })),
              ]} />
            </>,
          },
          {
            id: 'styles',
            content: <>
              <ToolbarMenuButton
                label="Cell styles"
                icon={<Palette size={15} />}
                render={(close) => <CellStylesGallery themeColors={themePalette} onSelect={(preset) => { applyCellStyleToSelection(preset); close() }} />}
              />
              <IconButton label="Conditional formatting" active={conditionalPanelOpen} onClick={() => setConditionalPanelOpen(true)}><Highlighter size={15} /></IconButton>
              <IconButton label="Format cells (Ctrl+1)" onClick={() => setFormatDialogTab('number')}><SlidersHorizontal size={15} /></IconButton>
            </>,
          },
        ]}
        trailing={<span className="format-badge">{documentFile.sourceFormat.toUpperCase()}</span>}
      />

      {documentFile.warnings.length > 0 && (
        <div className="compatibility-bar">
          <CircleAlert size={14} />
          <span><strong>Compatibility note:</strong> {documentFile.warnings[0]}</span>
          {documentFile.backupPath && <button type="button" aria-label="Show original backup" onClick={() => { void window.simpleCalc.showItem(documentFile.backupPath!) }}>Original backup</button>}
          {documentFile.warnings.length > 1 && <button type="button" onClick={() => setWarningOpen((open) => !open)}>{warningOpen ? 'Hide' : `Show all ${documentFile.warnings.length}`}</button>}
          {warningOpen && <div className="compatibility-details">{documentFile.warnings.map((warning) => <p key={warning}>{warning}</p>)}</div>}
        </div>
      )}

      {(showFormulaBar || searchOpen) && <div className={`formula-bar${showFormulaBar ? '' : ' is-search-only'}`}>
        <input
          className={`name-box${nameBoxInvalid ? ' is-invalid' : ''}`}
          aria-label="Name box"
          spellCheck={false}
          value={nameBoxDraft ?? rangeAddress(bounds)}
          onChange={(event) => setNameBoxDraft(event.target.value)}
          onFocus={(event) => event.currentTarget.select()}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault()
              if (commitNameBox(event.currentTarget.value)) document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })
            } else if (event.key === 'Escape') {
              event.preventDefault()
              setNameBoxDraft(null)
              setNameBoxInvalid(false)
              event.currentTarget.blur()
            }
          }}
          onBlur={() => setNameBoxDraft(null)}
        />
        <Sigma size={14} />
        {/* A textarea, not an input: HTML strips newlines out of an input's value, which would
            silently flatten every multi-line cell the moment the bar is touched. */}
        <div className="formula-input-shell">
        <textarea
          ref={formulaBarRef}
          className={`formula-bar-input${(editing && formulaInput !== 'bar' ? editorParts : formulaBarParts) ? ' is-formula-draft' : ''}`}
          aria-label="Formula bar"
          rows={1}
          spellCheck={false}
          wrap="off"
          value={editing && formulaInput !== 'bar' ? editing.draft : formulaDraft}
          placeholder={spillGhostFormula || undefined}
          readOnly={Boolean(spillGhostFormula) && !formulaDraft}
          onChange={(event) => {
            setFormulaDraft(event.target.value)
            pointRef.current = null
            setAssistDismissed(false)
            formulaBarDirtyRef.current = { sheetId: activeSheet.id, address: activeAddress, draft: event.target.value }
            handleFormulaCaret(event.target.selectionStart, event.target.selectionEnd)
          }}
          onSelect={(event) => handleFormulaCaret(event.currentTarget.selectionStart, event.currentTarget.selectionEnd)}
          onScroll={(event) => {
            const highlight = event.currentTarget.nextElementSibling as HTMLElement | null
            if (highlight) { highlight.scrollLeft = event.currentTarget.scrollLeft; highlight.scrollTop = event.currentTarget.scrollTop }
          }}
          onFocus={(event) => {
            formulaInputRef.current = 'bar'
            formulaInputElementRef.current = event.currentTarget
            setFormulaInput('bar')
            handleFormulaCaret(event.currentTarget.selectionStart, event.currentTarget.selectionEnd)
          }}
          onKeyDown={(event) => {
            if (assistState.items.length) {
              if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                event.preventDefault()
                const delta = event.key === 'ArrowDown' ? 1 : -1
                setAssistIndex((index) => (index + delta + assistState.items.length) % assistState.items.length)
                return
              }
              if (event.key === 'Tab' || (event.key === 'Enter' && !event.altKey && !event.ctrlKey)) { event.preventDefault(); acceptAssistItem(assistState.items[assistState.activeIndex]); return }
              if (event.key === 'Escape') { event.preventDefault(); setAssistDismissed(true); return }
            }
            if (event.key === 'F4') {
              event.preventDefault()
              const toggled = toggleAbsoluteReference(formulaDraft, event.currentTarget.selectionStart, event.currentTarget.selectionEnd)
              if (toggled) updateActiveFormulaDraft(toggled.text, toggled.selectionStart, toggled.selectionEnd)
              return
            }
            if (event.key === 'Enter' && (event.ctrlKey || event.metaKey) && !event.altKey) { event.preventDefault(); formulaBarDirtyRef.current = null; fillSelectionWithDraft(formulaDraft); document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true }); return }
            if (event.key === 'Enter' && !event.altKey) { event.preventDefault(); formulaBarDirtyRef.current = null; commitCell(activeAddress, formulaDraft); moveSelection(event.shiftKey ? -1 : 1, 0); document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true }) }
            else if (event.key === 'Escape') { formulaBarDirtyRef.current = null; setFormulaDraft(rawCellValue(activeCell)); event.currentTarget.blur() }
          }}
          onBlur={() => {
            if (formulaInputRef.current === 'bar') {
              formulaInputRef.current = null
              formulaInputElementRef.current = null
              setFormulaInput(null)
              setAssistAnchor(null)
            }
            pointRef.current = null
            const pending = formulaBarDirtyRef.current
            if (!pending) return
            formulaBarDirtyRef.current = null
            if (pending.sheetId === workbookRef.current?.activeSheetId) commitCell(pending.address, pending.draft)
          }}
        />
        {(editing && formulaInput !== 'bar' ? editorParts : formulaBarParts) && (
          <FormulaHighlight className="formula-bar-input fa-highlight-layer" parts={(editing && formulaInput !== 'bar' ? editorParts : formulaBarParts)!} />
        )}
        </div>
        {searchOpen && (
          <div
            className={`search-panel${replaceOpen ? ' has-replace' : ''}`}
            role="search"
            aria-label={replaceOpen ? 'Find and replace in this sheet' : 'Find in this sheet'}
            onKeyDown={(event) => {
              if (event.key === 'Escape') { event.preventDefault(); closeSearch() }
            }}
          >
            <Search size={14} aria-hidden="true" />
            <input
              ref={searchInputRef}
              aria-label="Find in this sheet"
              placeholder="Find in this sheet"
              value={searchQuery}
              onChange={(event) => setSearchQuery(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') { event.preventDefault(); moveSearch(event.shiftKey ? -1 : 1) }
              }}
            />
            <span className="search-count" aria-live="polite">
              {searchQuery.trim() ? (searchMatches.length ? `${searchIndex + 1} of ${searchMatches.length}` : 'No matches') : 'Type to search'}
            </span>
            <button type="button" aria-label="Previous match" title="Previous match (Shift+Enter)" disabled={!searchMatches.length} onClick={() => moveSearch(-1)}>↑</button>
            <button type="button" aria-label="Next match" title="Next match (Enter)" disabled={!searchMatches.length} onClick={() => moveSearch(1)}>↓</button>
            <button type="button" className={replaceOpen ? 'is-active' : undefined} aria-label="Toggle replace" aria-pressed={replaceOpen} title="Replace (Ctrl+H)" onClick={() => setReplaceOpen((value) => !value)}><Replace size={13} /></button>
            <button type="button" className={searchOptionsOpen ? 'is-active' : undefined} aria-label="Search options" aria-pressed={searchOptionsOpen} title="Options" onClick={() => setSearchOptionsOpen((value) => !value)}><SlidersHorizontal size={13} /></button>
            <button type="button" aria-label="Close find" title="Close (Esc)" onClick={closeSearch}><X size={13} /></button>
            {searchOptionsOpen && (
              <div className="search-options-row">
                <label><input type="checkbox" checked={searchOptions.matchCase} onChange={(event) => setSearchOptions((value) => ({ ...value, matchCase: event.target.checked }))} /> Match case</label>
                <label><input type="checkbox" checked={searchOptions.entireCell} onChange={(event) => setSearchOptions((value) => ({ ...value, entireCell: event.target.checked }))} /> Entire cell</label>
                <label><input type="checkbox" checked={searchOptions.workbook} onChange={(event) => setSearchOptions((value) => ({ ...value, workbook: event.target.checked }))} /> All sheets</label>
                <label>Look in <select value={searchOptions.lookIn} onChange={(event) => setSearchOptions((value) => ({ ...value, lookIn: event.target.value as SearchOptions['lookIn'] }))}><option value="values">Values</option><option value="formulas">Formulas</option></select></label>
              </div>
            )}
            {replaceOpen && (
              <div className="search-replace-row">
                <input
                  aria-label="Replace with"
                  placeholder="Replace with"
                  value={replaceValue}
                  onChange={(event) => setReplaceValue(event.target.value)}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') { event.preventDefault(); replaceCurrentMatch() }
                  }}
                />
                <button type="button" disabled={!searchMatches.length} onClick={replaceCurrentMatch}>Replace</button>
                <button type="button" disabled={!searchMatches.length} onClick={replaceAllMatches}>Replace all</button>
              </div>
            )}
          </div>
        )}
      </div>}

      <div className="workspace-body">
      <SpreadsheetGrid
        sheet={activeSheet}
        selection={selection}
        editing={editing}
        zoom={zoom}
        useImportedDefaults
        showGridlines={gridlinesVisible}
        showNotes={showNotes}
        displayValue={displayValue}
        resolvedValue={resolvedValue}
        staleValue={staleValue}
        onSelection={handleGridSelection}
        onBeginEdit={beginEdit}
        onDraft={(draft) => {
          pointRef.current = null
          setAssistDismissed(false)
          setEditing((current) => current ? { ...current, draft } : current)
        }}
        onCommitEdit={(direction) => {
          if (editing && !commitWithValidation(editing.address, editing.draft)) return
          if (direction === 'down') moveSelection(1, 0)
          else if (direction === 'up') moveSelection(-1, 0)
          else if (direction === 'right') moveSelection(0, 1)
          else if (direction === 'left') moveSelection(0, -1)
        }}
        onCancelEdit={() => setEditing(null)}
        onFill={autofillSelection}
        onCellValue={setCellValue}
        onOpenHyperlink={openHyperlink}
        onColumnResize={resizeColumn}
        onRowResize={resizeRow}
        onFreeze={setFreeze}
        onKeyDown={handleGridKeyDown}
        onContextTarget={openGridContextMenu}
        referenceHighlights={referenceHighlights}
        spillOutline={editing ? null : spillOutline}
        editorParts={editorParts}
        editorSelectionRequest={editorSelectionRequest}
        onPointCell={pointCell}
        onEditorKeyDown={handleEditorKeyDown}
        onEditorCaret={handleFormulaCaret}
        onEditorFocusChange={handleEditorFocusChange}
        conditionalFormats={conditionalFormats}
        filterHeaders={filterTargets.map((target) => ({
          key: target.key,
          row: target.bounds.top,
          left: target.bounds.left,
          right: target.bounds.right,
          active: new Set(Object.entries(target.state.columns || {}).filter(([, criteria]) => criteriaIsActive(criteria)).map(([offset]) => target.bounds.left + Number(offset))),
          sort: target.state.sort ? { column: target.bounds.left + target.state.sort.column, descending: target.state.sort.descending } : undefined,
        }))}
        onFilterButton={(key, col, anchor) => setFilterMenu({ key, col, anchor })}
        inputMessage={(() => {
          const found = findValidation(activeSheet.dataValidations, selection.focus.row, selection.focus.col)
          const validation = found?.validation as { showInputMessage?: boolean; promptTitle?: string; prompt?: string } | undefined
          return validation?.showInputMessage && (validation.promptTitle || validation.prompt)
            ? { row: selection.focus.row, col: selection.focus.col, title: validation.promptTitle || '', text: validation.prompt || '' }
            : null
        })()}
        invalidCells={invalidCells}
        cutRange={cutRange && cutRange.sheetId === activeSheet.id ? cutRange : null}
        renderCharts={activeSheet.charts?.length || activeSheet.images?.length || traceState?.sheetId === activeSheet.id ? (geometry, viewport, geometryVersion) => (<>
          {traceState?.sheetId === activeSheet.id ? <TraceArrowLayer arrows={traceState.arrows} geometry={geometry} /> : null}
          {activeSheet.images?.length ? (
            <ImageLayer
              images={activeSheet.images}
              geometry={geometry}
              zoom={zoom}
              selectedImageId={selectedImageId}
              viewport={viewport}
              onSelect={setSelectedImageId}
              onChange={updateImage}
              onDelete={deleteImage}
              onContextMenu={(id, position) => setContextMenu({ kind: 'image', x: position.clientX, y: position.clientY, imageId: id })}
            />
          ) : null}
          {activeSheet.charts?.length ? <ChartLayer
            charts={activeSheet.charts}
            geometry={geometry}
            geometryVersion={geometryVersion}
            zoom={zoom}
            selectedChartId={selectedChartId}
            resolveData={resolveChart}
            dataVersion={workbook}
            onSelect={setSelectedChartId}
            onChange={updateChart}
            onDelete={deleteChart}
            onEdit={(id) => setChartEditorId(id)}
            viewport={viewport}
          /> : null}
        </>) : undefined}
        listOptionsFor={listOptionsFor}
        sparklineFor={sparklineFor}
        outline={outlineState}
        onOutlineToggle={(axis, group) => updateOutline((sheet) => toggleGroup(sheet, axis, group))}
        onOutlineLevel={(axis, level) => updateOutline((sheet) => showOutlineLevel(sheet, axis, level))}
      />
      {activePivot && pivotEditorHidden !== activePivot.id && !chartEditorId && (
        <PivotEditorPanel
          pivot={activePivot}
          source={activePivotSource?.source || null}
          sourceError={activePivotSource?.error}
          onChange={(next) => { void refreshPivot(activePivot.id, next) }}
          onRefresh={() => { void refreshPivot(activePivot.id).then((done) => { if (done) setToast('Refreshed the pivot table') }) }}
          onDelete={() => deletePivot(activePivot.id)}
          onClose={() => setPivotEditorHidden(activePivot.id)}
        />
      )}
      {tableDesignOpen && activeTable && !chartEditorId && !activePivot && (
        <TableDesignPanel
          table={activeTable}
          theme={activeThemeColors}
          onRename={(name) => renameActiveTable(activeTable.id, name)}
          onResize={(range) => resizeActiveTable(activeTable.id, range)}
          onOption={(option, value) => setTableOption(activeTable.id, option, value)}
          onStyle={(style) => updateTable(activeTable.id, (_workbook, _sheet, table) => { table.style = { ...(table.style || {}), theme: style } })}
          onTotalsFunction={(index, fn: TotalFunctionId) => updateTable(activeTable.id, (_workbook, sheet, table) => setTotalsFunction(sheet, table, index, fn))}
          onConvert={() => convertActiveTable(activeTable.id)}
          onRemoveDuplicates={() => {
            const regions = tableRegions(activeTable)
            if (regions) setSelection({ anchor: { row: regions.header ?? regions.dataTop, col: regions.left }, focus: { row: regions.dataBottom, col: regions.right } })
            setCleanupDialog('duplicates')
          }}
          onClose={() => setTableDesignOpen(false)}
        />
      )}
      {chartEditorId && (() => {
        const chart = activeSheet.charts?.find((item) => item.id === chartEditorId)
        if (!chart) return null
        return (
          <ChartEditorPanel
            chart={chart}
            sheetNames={workbook.sheets.map((sheet) => sheet.name)}
            activeSheetName={activeSheet.name}
            accessor={chartAccessor}
            data={resolveChart(chart)}
            themePalette={chartPalette}
            onChange={(next) => updateChart({ ...next, modified: true })}
            onClose={() => setChartEditorId(null)}
            onDelete={() => deleteChart(chart.id)}
          />
        )
      })()}
      </div>
      {formulaInput && (
        <FormulaAssist
          anchor={assistAnchor}
          state={assistState}
          onPick={acceptAssistItem}
          onHover={setAssistIndex}
        />
      )}

      <div className="sheet-strip">
        <div className="sheet-actions">
          <IconButton label="Add sheet" onClick={addSheet}><Plus size={15} /></IconButton>
        </div>
        <div className="sheet-tabs">
          {workbook.sheets.filter((sheet) => sheet.state === 'visible').map((sheet) => {
            const tabArgb = String((sheet.properties?.tabColor as { argb?: string } | undefined)?.argb || '')
            const tabColor = /^[0-9a-f]{8}$/i.test(tabArgb) ? `#${tabArgb.slice(2)}` : cssColor(sheet.properties?.tabColor)
            return (
            <button
              type="button"
              key={sheet.id}
              className={`sheet-tab${sheet.id === workbook.activeSheetId ? ' is-active' : ''}${tabColor ? ' has-color' : ''}${tabDropTarget === sheet.id ? ' is-drop-target' : ''}`}
              style={tabColor ? { '--tab-color': tabColor } as CSSProperties : undefined}
              draggable
              onDragStart={(event) => {
                event.dataTransfer.setData('application/x-simple-calc-sheet', sheet.id)
                event.dataTransfer.effectAllowed = 'move'
              }}
              onDragOver={(event) => {
                if (!event.dataTransfer.types.includes('application/x-simple-calc-sheet')) return
                event.preventDefault()
                event.stopPropagation()
                if (tabDropTarget !== sheet.id) setTabDropTarget(sheet.id)
              }}
              onDragLeave={() => setTabDropTarget((current) => current === sheet.id ? null : current)}
              onDrop={(event) => {
                const dragged = event.dataTransfer.getData('application/x-simple-calc-sheet')
                if (!dragged) return
                event.preventDefault()
                event.stopPropagation()
                setTabDropTarget(null)
                moveSheet(dragged, workbook.sheets.findIndex((item) => item.id === sheet.id))
              }}
              onDragEnd={() => setTabDropTarget(null)}
              onClick={() => activateSheet(sheet.id)}
              onDoubleClick={() => { void renameSheet(sheet.id) }}
              onContextMenu={(event) => {
                event.preventDefault()
                if (workbook.activeSheetId !== sheet.id) {
                  setWorkbook((current) => current ? { ...current, activeSheetId: sheet.id } : current)
                  setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
                  setEditing(null)
                }
                setContextMenu({ kind: 'sheet-tab', sheetId: sheet.id, x: event.clientX, y: event.clientY })
              }}
              title={sheetProtection(sheet) ? 'Protected sheet · double-click to rename · drag to reorder' : 'Double-click to rename · drag to reorder'}
            >{sheetProtection(sheet) && <Lock size={11} className="sheet-tab-lock" aria-label="Protected" />}{sheet.name}</button>
            )
          })}
        </div>
        <div className="sheet-menu-actions">
          <IconButton label="Duplicate active sheet" onClick={duplicateActiveSheet}><Copy size={14} /></IconButton>
          <IconButton label="Delete active sheet" onClick={deleteActiveSheet}><Trash2 size={14} /></IconButton>
        </div>
      </div>

      <footer className="statusbar">
        <div className="status-left"><span className="ready-dot" /> <span>{dirty ? 'Unsaved changes' : 'Ready'}</span><span className="status-divider" /><span>{workbook.sheets.length} {workbook.sheets.length === 1 ? 'sheet' : 'sheets'}</span>{activeCell?.formula && <><span className="status-divider" /><span>Formula</span></>}{activeProtection && <><span className="status-divider" /><span className="status-protected"><Lock size={11} aria-hidden="true" /> Protected</span></>}</div>
        <div className="status-right">
          {selectionStats.selected > 1 && (
            <div
              className="status-stats"
              aria-live="polite"
              title="Click a value to copy it. Right-click to choose which statistics appear."
              onContextMenu={(event) => {
                event.preventDefault()
                setContextMenu({ kind: 'status', x: event.clientX, y: event.clientY })
              }}
            >
              {statusItems.map((item) => (
                <button
                  type="button"
                  key={item.key}
                  className="status-stat"
                  data-status-stat={item.key}
                  onClick={() => { void navigator.clipboard.writeText(String(item.value)).then(() => setToast(`Copied ${item.label.toLocaleLowerCase()} ${item.text}`)).catch(() => undefined) }}
                >
                  <span>{item.label}:</span> <strong>{item.text}</strong>
                </button>
              ))}
            </div>
          )}
          <div className="zoom-control">
            <button type="button" aria-label="Zoom out" onClick={() => setZoom((value) => clamp(value - 0.1, 0.1, 2))}><Minus size={13} /></button>
            <input type="range" min="10" max="200" step="1" value={Math.round(zoom * 100)} aria-label="Zoom" onChange={(event) => setZoom(Number(event.target.value) / 100)} />
            <button type="button" aria-label="Zoom in" onClick={() => setZoom((value) => clamp(value + 0.1, 0.1, 2))}><Plus size={13} /></button>
            <span>{Math.round(zoom * 100)}%</span>
          </div>
        </div>
      </footer>

      {contextMenu && (
        <SpreadsheetContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          label={contextMenu.kind === 'image' ? 'Picture menu' : contextMenu.kind === 'status' ? 'Status bar statistics' : contextMenu.kind === 'sheet-tab' ? 'Sheet menu' : contextMenu.kind === 'row-header' ? 'Row menu' : contextMenu.kind === 'column-header' ? 'Column menu' : contextMenu.kind === 'corner' ? 'All cells menu' : 'Cell menu'}
          items={contextMenuItems}
          onClose={closeContextMenu}
        />
      )}
      {textPrompt && <TextPromptDialog key={textPrompt.id} request={textPrompt} onClose={(value) => { setTextPrompt(null); textPrompt.resolve(value) }} />}
      {formatDialogTab && (
        <FormatCellsDialog
          initialTab={formatDialogTab}
          style={{ ...(activeCell?.style || {}), numFmt: activeCell?.numFmt || activeCell?.style?.numFmt }}
          sampleValue={(() => { const value = calcEngine?.getValue(activeSheet.id, activeAddress); return value === undefined ? null : value })()}
          multiRow={selectedRowCount > 1}
          multiColumn={selectedColumnCount > 1}
          themeColors={themePalette}
          merged={(activeSheet.merges || []).some((range) => { const merged = mergeBounds(range); return Boolean(merged && merged.top === bounds.top && merged.left === bounds.left && merged.bottom === bounds.bottom && merged.right === bounds.right) })}
          defaultFont={{ name: String((workbook.metadata?.normalFont as { name?: string } | undefined)?.name || 'Calibri'), size: Number((workbook.metadata?.normalFont as { size?: number } | undefined)?.size || 11) }}
          onApply={applyFormatChangeToSelection}
          onClose={() => { setFormatDialogTab(null); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })) }}
        />
      )}
      {alertRequest && <AlertDialog request={alertRequest} onClose={(button) => { const request = alertRequest; setAlertRequest(null); request.resolve(button); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.cell-editor, .sheet-viewport')?.focus({ preventScroll: true })) }} />}
      {filterMenu && (() => {
        const target = filterTargets.find((item) => item.key === filterMenu.key)
        if (!target) return null
        const offset = filterMenu.col - target.bounds.left
        return (
          <FilterMenu
            columnLabel={displayValue(activeSheet.id, addressOf({ row: target.bounds.top, col: filterMenu.col })) || `Column ${columnName(filterMenu.col)}`}
            criteria={target.state.columns[offset] || null}
            values={distinctColumnValues(target.state, offset, dataHostFor(activeSheet))}
            anchor={filterMenu.anchor}
            sortDirection={target.state.sort?.column === offset ? (target.state.sort.descending ? 'desc' : 'asc') : null}
            onApply={(criteria) => { applyFilterCriteria(target.key, filterMenu.col, criteria); setFilterMenu(null) }}
            onSort={(descending) => { sortFilterColumn(target.key, filterMenu.col, descending); setFilterMenu(null) }}
            onClose={() => setFilterMenu(null)}
          />
        )
      })()}
      {pivotDialog && (
        <CreatePivotDialog
          initialRange={pivotDialog.range}
          initialLocation={pivotDialog.location}
          onCreate={createPivot}
          onClose={() => { setPivotDialog(null); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })) }}
        />
      )}
      {sparklineDialog && (
        <CreateSparklinesDialog
          kind={sparklineDialog.kind}
          initialData={sparklineDialog.data}
          initialLocation={sparklineDialog.location}
          onCreate={createSparklines}
          onClose={() => { setSparklineDialog(null); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })) }}
        />
      )}
      {protectDialogOpen && (
        <ProtectSheetDialog
          sheetName={activeSheet.name}
          onProtect={protectActiveSheet}
          onClose={() => { setProtectDialogOpen(false); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })) }}
        />
      )}
      {goalSeekOpen && (
        <GoalSeekDialog
          initialSetCell={activeCell?.formula ? addressOf(selection.focus) : ''}
          initialChangingCell=""
          onSeek={runGoalSeek}
          onClose={() => { setGoalSeekOpen(false); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })) }}
        />
      )}
      {shiftDialog && (
        <ShiftCellsDialog
          mode={shiftDialog.mode}
          initial={shiftDialog.initial}
          onApply={(choice) => applyShiftChoice(shiftDialog.mode, choice)}
          onClose={() => { setShiftDialog(null); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })) }}
        />
      )}
      {tableDialog && (
        <CreateTableDialog
          initialRange={tableDialog.range}
          initialHasHeaders={tableDialog.hasHeaders}
          initialStyle={tableDialog.style}
          theme={activeThemeColors}
          onCreate={createTableFromDialog}
          onClose={() => { setTableDialog(null); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })) }}
        />
      )}
      {sortDialogOpen && (() => {
        const selected = selectionBounds(selection)
        const region = (selected.top === selected.bottom && selected.left === selected.right ? filterRangeForSelection(activeSheet, selected, selection.focus) : selected) || selected
        const host = dataHostFor(activeSheet)
        return (
          <SortDialog
            rangeLabel={rangeAddress(region)}
            initialHasHeader={detectHeaderRow(region, host)}
            initialLevels={[{ key: clamp(selection.focus.col - region.left, 0, region.right - region.left) }]}
            labelsFor={(orientation, hasHeader) => sortKeyLabels(region, orientation, hasHeader, host)}
            colorsFor={(orientation, key, sortOn, hasHeader) => sortKeyColors(region, orientation, key, sortOn, hasHeader, host)}
            onSort={runCustomSort}
            onClose={() => setSortDialogOpen(false)}
          />
        )
      })()}
      {validationDialogOpen && (() => {
        const found = findValidation(activeSheet.dataValidations, selection.focus.row, selection.focus.col)
        return (
          <DataValidationDialog
            rangeLabel={rangeAddress(bounds)}
            initial={found?.validation ?? null}
            sameSettingsCount={found ? Math.max(0, rangesWithSameValidation(activeSheet.dataValidations, found.validation).length - 1) : 0}
            date1904={Boolean(workbook.metadata?.date1904)}
            onApply={(validation, options) => {
              mutateWorkbook((next) => {
                const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
                let validations = withValidation(sheet.dataValidations, bounds, validation)
                if (options.applyToSameSettings && found) {
                  for (const key of rangesWithSameValidation(sheet.dataValidations, found.validation)) {
                    const range = mergeBounds(key)
                    if (range) validations = withValidation(validations, range, validation)
                  }
                }
                sheet.dataValidations = validations
              })
              setValidationDialogOpen(false)
            }}
            onClose={() => setValidationDialogOpen(false)}
          />
        )
      })()}
      {cleanupDialog === 'duplicates' && (() => {
        const selected = selectionBounds(selection)
        const region = (selected.top === selected.bottom && selected.left === selected.right ? filterRangeForSelection(activeSheet, selected, selection.focus) : selected) || selected
        const host = dataHostFor(activeSheet)
        return (
          <RemoveDuplicatesDialog
            rangeLabel={rangeAddress(region)}
            initialHasHeader={detectHeaderRow(region, host)}
            labelsFor={(hasHeader) => sortKeyLabels(region, 'rows', hasHeader, host)}
            onApply={({ columns, hasHeader }) => {
              const outcome = removeDuplicates(activeSheet, region, { columns, hasHeader }, host, shiftFormulaReferences)
              if (!outcome.ok) return { error: outcome.error }
              if (outcome.removed) mutateWorkbook((next) => applyCellChanges(next.sheets.find((item) => item.id === next.activeSheetId)!.cells, outcome.changes))
              return { removed: outcome.removed, remaining: outcome.remaining }
            }}
            onClose={() => setCleanupDialog(null)}
          />
        )
      })()}
      {cleanupDialog === 'split' && (() => {
        const selected = selectionBounds(selection)
        const source = { top: selected.top, bottom: selected.bottom, left: selected.left, right: selected.left }
        const host = dataHostFor(activeSheet)
        return (
          <TextToColumnsDialog
            sourceLabel={rangeAddress(source)}
            sampleLines={splitSourceLines(source, host)}
            defaultDestination={addressOf({ row: source.top, col: source.left })}
            onApply={(options, confirmOverwrite) => {
              const outcome = splitTextToColumns(activeSheet, source, options, host)
              if (!outcome.ok) return { error: outcome.error }
              if (outcome.overwrites && !confirmOverwrite) return { confirm: 'There’s already data here. Do you want to replace it?' }
              mutateWorkbook((next) => {
                const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
                applyCellChanges(sheet.cells, outcome.changes)
                sheet.colCount = Math.max(sheet.colCount, source.left + 32)
              })
              setCleanupDialog(null)
            }}
            onClose={() => setCleanupDialog(null)}
          />
        )
      })()}
      {pasteSpecialOpen && (
        <PasteSpecialDialog
          sourceLabel={internalClipboard.current ? `${rangeAddress({ top: internalClipboard.current.origin.row, left: internalClipboard.current.origin.col, bottom: internalClipboard.current.origin.row + internalClipboard.current.cells.length - 1, right: internalClipboard.current.origin.col + (internalClipboard.current.cells[0]?.length || 1) - 1 })} on ${internalClipboard.current.sheetName}` : 'Clipboard'}
          canPasteLink={Boolean(internalClipboard.current)}
          disabledPasteTypes={internalClipboard.current ? [] : ['validation']}
          onCancel={() => setPasteSpecialOpen(false)}
          onPaste={(options) => { setPasteSpecialOpen(false); void pasteSelection(options) }}
        />
      )}
      {conditionalPanelOpen && (
        <ConditionalFormatPanel
          sheetName={activeSheet.name}
          selectionRef={rangeAddress(bounds)}
          conditionalFormattings={(activeSheet.conditionalFormattings as unknown[] | undefined) || []}
          onChange={(nextRules) => mutateWorkbook((next) => {
            const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
            sheet.conditionalFormattings = nextRules
            delete sheet.conditionalFormattingsTruncated
          })}
          onClose={() => { setConditionalPanelOpen(false); window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus({ preventScroll: true })) }}
          cssColor={(color, fallback) => cssColor(color, fallback)}
        />
      )}
      {nameManagerOpen && (
        <NameManagerDialog
          names={workbook.definedNames || (workbook.metadata?.definedNames || []).map((item) => ({ name: item.name, ranges: [String(item.ranges || item.formula || '')], localSheetIndex: item.localSheetId }))}
          sheetNames={workbook.sheets.map((sheet) => sheet.name)}
          selectionReference={`=${selectionReferenceText()}`}
          evaluate={(refersTo) => {
            if (!calcEngine) return ''
            const value = calcEngine.evaluateAt(activeSheet.id, refersTo)
            return value === null ? '' : typeof value === 'number' ? value.toLocaleString(undefined, { maximumFractionDigits: 6 }) : String(value)
          }}
          onSave={saveDefinedNames}
          onClose={() => setNameManagerOpen(false)}
        />
      )}
      {exportOpen && (
        <SpreadsheetExportDialog
          sheetCount={workbook.sheets.filter((sheet) => sheet.state !== 'hidden' && sheet.state !== 'veryHidden').length || 1}
          selectionLabel={rangeAddress(bounds)}
          defaultGridlines={gridlinesVisible}
          compatibilityWarning={Boolean(documentFile.warnings.some((warning) => /chart|drawing|image|pivot|slicer|object/i.test(warning)) || workbook.sheets.some((sheet) => (sheet.requiresSourcePackage || []).some((feature) => /chart|drawing|image|pivot|slicer|object/i.test(feature))))}
          onClose={() => setExportOpen(false)}
          onExport={exportWorkbook}
        />
      )}
      {printOpen && (
        <SpreadsheetPrintDialog
          sheetCount={workbook.sheets.filter((sheet) => sheet.state !== 'hidden' && sheet.state !== 'veryHidden').length || 1}
          selectionLabel={rangeAddress(bounds)}
          defaultGridlines={gridlinesVisible}
          compatibilityWarning={Boolean(documentFile.warnings.some((warning) => /chart|drawing|image|pivot|slicer|object/i.test(warning)) || workbook.sheets.some((sheet) => (sheet.requiresSourcePackage || []).some((feature) => /chart|drawing|image|pivot|slicer|object/i.test(feature))))}
          onClose={() => setPrintOpen(false)}
          onRenderPreview={renderPrintPreview}
          onPrint={printWorkbook}
        />
      )}
      {busy && <div className="busy-overlay"><div className="busy-card"><span className="spinner" />{busy}</div></div>}
      {toast && <div className="toast"><span>{toast}</span><button type="button" aria-label="Dismiss" onClick={() => setToast('')}><X size={13} /></button></div>}
      <span className="history-signal" aria-hidden="true">{historyTick}</span>
    </div>
  )
}
