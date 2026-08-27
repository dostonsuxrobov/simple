import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent, CSSProperties, DragEvent, KeyboardEvent, MouseEvent, PointerEvent as ReactPointerEvent, RefObject } from 'react'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  Bold,
  Check,
  CheckSquare,
  ChevronDown,
  CircleAlert,
  Columns3,
  Copy,
  DollarSign,
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
  Plus,
  Redo2,
  Save,
  Search,
  Sigma,
  Square,
  StickyNote,
  Strikethrough,
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
import { SpreadsheetMenus } from './components/SpreadsheetMenus'
import type { SpreadsheetMenuDefinition } from './components/SpreadsheetMenus'
import { createAutofillPatch } from './lib/autofill'
import { evaluateFormula, shiftFormulaReferences } from './lib/formulas'
import { accountingDisplayParts, formatScalar, isAccountingNumberFormat } from './lib/number-format'
import { applySelectionStructureCommand } from './lib/sheet-operations'
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
  SheetData,
  SpreadsheetColor,
  WorkbookModel,
  WorkbookPayload,
} from './spreadsheet-types'

const RECENT_KEY = 'simple-calc:recent:v1'
const DEFAULT_ROWS = 200
const DEFAULT_COLS = 40
const DEFAULT_ROW_HEIGHT = 26
const DEFAULT_COL_WIDTH = 108
const IMPORTED_ROW_HEIGHT = 20
const IMPORTED_COL_WIDTH = 64
const HEADER_HEIGHT = 27
const HEADER_WIDTH = 46
const MAX_HISTORY = 24
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

type AggregateMode = 'sum' | 'average' | 'minimum' | 'maximum' | 'count' | 'numeric'

interface OpenWorkbook {
  documentId: string
  path: string | null
  name: string
  sourceFormat: string
  requiresSaveAs: boolean
  warnings: string[]
  stats?: WorkbookPayload['stats']
}

interface EditingState {
  address: string
  draft: string
}

interface InternalClipboard {
  text: string
  origin: Coord
  cells: CellData[][]
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

  if (draft.startsWith('=')) {
    const formula = normalizeFormula(draft)
    if (formula) next.formula = formula
    return next
  }
  if (!draft.length) return next
  if (/^(true|false)$/i.test(draft)) {
    next.value = draft.toLowerCase() === 'true'
    return next
  }
  const dateSerial = dateSerialFromDraft(draft.trim())
  if (dateSerial !== null) {
    next.value = dateSerial
    next.type = 'date'
    next.numFmt ||= 'm/d/yyyy'
    return next
  }
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(draft) && !/^[-+]?0\d/.test(draft)) {
    next.value = Number(draft)
    return next
  }
  next.value = draft
  return next
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
  const workbook = structuredClone(input || blankWorkbook()) as WorkbookModel
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
    sheet.cells = normalizedCells
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
  const next = [item, ...current.filter((entry) => entry.path.toLowerCase() !== item.path.toLowerCase())].slice(0, 8)
  localStorage.setItem(RECENT_KEY, JSON.stringify(next))
  return next
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
  const amount = clamp(tint, -1, 1)
  return [0, 2, 4].map((offset) => {
    const channel = Number.parseInt(hex.slice(offset, offset + 2), 16)
    const adjusted = amount < 0
      ? channel * (1 + amount)
      : channel + (255 - channel) * amount
    return Math.round(clamp(adjusted, 0, 255)).toString(16).padStart(2, '0')
  }).join('').toUpperCase()
}

function cssColor(value: unknown, fallback = '') {
  let color = ''
  let tint = 0
  if (typeof value === 'string') color = value
  else if (value && typeof value === 'object') {
    const object = value as SpreadsheetColor
    color = object.argb || object.rgb || ''
    tint = Number(object.tint) || 0
    if (!color && Number.isInteger(object.theme)) color = OFFICE_THEME_COLORS[Number(object.theme)] || ''
    if (!color && Number.isInteger(object.indexed)) color = EXCEL_INDEXED_COLORS[Number(object.indexed)] || ''
  }
  color = color.replace(/^#/, '')
  if (color.length === 8) color = color.slice(2)
  if (/^[0-9a-f]{6}$/i.test(color)) return `#${tintHex(color.toUpperCase(), tint)}`
  return fallback
}

function fontCss(font: CellFont = {}): CSSProperties {
  const underline = font.underline
  const decorations = [underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ')
  return {
    fontFamily: font.name || undefined,
    fontSize: font.size ? `${font.size * (4 / 3)}px` : undefined,
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

function borderCss(side: CellBorderSide | undefined) {
  if (!side || (!side.style && !side.color)) return undefined
  const style = String(side.style || 'thin').toLocaleLowerCase()
  const color = cssColor(side.color, '#4b4b4b')
  const width = /thick/.test(style) ? 3 : /medium|double/.test(style) ? 2 : 1
  const line = /double/.test(style) ? 'double' : /dot/.test(style) ? 'dotted' : /dash/.test(style) ? 'dashed' : 'solid'
  return `${line === 'double' ? Math.max(3, width) : width}px ${line} ${color}`
}

function cellCss(cell?: CellData): CSSProperties {
  const style = cell?.style || {}
  const alignment = style.alignment || {}
  const border = style.border || {}
  const rawValue = cell?.formula ? cell.result : cell?.value
  const horizontal = String(alignment.horizontal || '').toLocaleLowerCase()
  const vertical = String(alignment.vertical || '').toLocaleLowerCase()
  const indent = Math.max(0, Number(alignment.indent || 0) + Number(alignment.relativeIndent || 0))
  const inferredAlignment = typeof rawValue === 'number' ? 'right' : typeof rawValue === 'boolean' ? 'center' : undefined
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
    ...fontCss(style.font),
    ...fill,
    backgroundImage,
    borderTop: borderCss(border.top),
    borderLeft: borderCss(border.left),
    borderBottom: borderCss(border.bottom),
    borderRight: borderCss(border.right),
    alignItems: vertical === 'top' ? 'flex-start' : vertical === 'middle' || vertical === 'center' ? 'center' : vertical === 'bottom' ? 'flex-end' : undefined,
    textAlign,
    whiteSpace: alignment.wrapText ? 'pre-wrap' : 'nowrap',
    overflowWrap: alignment.wrapText ? 'anywhere' : undefined,
    direction: alignment.readingOrder === 'rtl' || alignment.readingOrder === 2 ? 'rtl' : alignment.readingOrder === 'ltr' || alignment.readingOrder === 1 ? 'ltr' : undefined,
    paddingLeft: indent && textAlign !== 'right' ? `${6 + indent * 9}px` : undefined,
    paddingRight: indent && textAlign === 'right' ? `${6 + indent * 9}px` : undefined,
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

function createValueResolver(workbook: WorkbookModel) {
  const cache = new Map<string, CellScalar>()
  const visiting = new Set<string>()

  const resolveSheet = (reference: string) => workbook.sheets.find((sheet) => (
    sheet.id === reference || sheet.name.toLocaleLowerCase() === reference.toLocaleLowerCase()
  ))

  const resolve = (sheetReference: string, address: string): CellScalar => {
    const sheet = resolveSheet(sheetReference)
    if (!sheet) return '#REF!'
    const normalized = address.replace(/\$/g, '').toUpperCase()
    const key = `${sheet.id}!${normalized}`
    if (cache.has(key)) return cache.get(key) ?? null
    if (visiting.has(key)) return '#CIRC!'
    const cell = sheet.cells[normalized]
    if (!cell) return null
    if (!cell.formula) return cell.value ?? cell.result ?? null
    visiting.add(key)
    let value: CellScalar
    try {
      const evaluated = evaluateFormula(cell.formula, sheet.id, resolve)
      value = (evaluated === undefined ? null : evaluated) as CellScalar
      if (typeof value === 'string' && value.startsWith('#') && cell.result !== undefined && cell.result !== null && !['#DIV/0!', '#CIRC!', '#REF!'].includes(value)) {
        value = cell.result
      }
    } catch {
      value = cell.result ?? '#ERROR!'
    }
    visiting.delete(key)
    cache.set(key, value)
    return value
  }
  return resolve
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

function columnPixelWidth(value: number | undefined, zoom: number) {
  if (!value) return DEFAULT_COL_WIDTH * zoom
  const pixels = value * 8
  return clamp(pixels, 2, 1_000) * zoom
}

function rowPixelHeight(value: number | undefined, zoom: number) {
  if (!value) return DEFAULT_ROW_HEIGHT * zoom
  return clamp(value * (4 / 3), 2, 640) * zoom
}

function modelColumnWidth(pixelWidth: number, zoom: number) {
  return Math.round(clamp(pixelWidth / zoom / 8, 0.2, 125) * 100) / 100
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

function noteText(note: CellData['note']) {
  if (!note) return ''
  if (typeof note === 'string') return note
  if (Array.isArray(note.comments)) return note.comments.map((comment) => `${comment.author ? `${comment.author}: ` : ''}${comment.text || ''}`).join('\n')
  if (Array.isArray(note.texts)) return note.texts.map((item) => item.text || '').join('')
  return typeof note.text === 'string' ? note.text : ''
}

function validationForCell(sheet: SheetData, address: string) {
  const validations = sheet.dataValidations || {}
  if (validations[address] && typeof validations[address] === 'object') return validations[address] as Record<string, unknown>
  const coord = coordOf(address)
  if (!coord) return undefined
  for (const [range, validation] of Object.entries(validations)) {
    const bounds = mergeBounds(range)
    if (!bounds || coord.row < bounds.top || coord.row > bounds.bottom || coord.col < bounds.left || coord.col > bounds.right) continue
    if (validation && typeof validation === 'object') return validation as Record<string, unknown>
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
  onSelection: (selection: Selection) => void
  onBeginEdit: (draft?: string) => void
  onDraft: (draft: string) => void
  onCommitEdit: () => void
  onCancelEdit: () => void
  onFill: (target: Coord) => void
  onCellValue: (address: string, value: CellData['value']) => void
  onOpenHyperlink: (target: string) => void
  onColumnResize: (column: number, width: number) => void
  onRowResize: (row: number, height: number) => void
  onFreeze: (axis: 'rows' | 'columns', count: number) => void
  onKeyDown: (event: KeyboardEvent<HTMLDivElement>) => void
}

interface GridCellViewProps {
  sheet: SheetData
  row: number
  col: number
  pane: CellPane
  merge?: NonNullable<ReturnType<typeof mergeBounds>>
  columns: number
  frozenColumnCount: number
  defaultColumn: number
  defaultRow: number
  columnMetrics: AxisMetric[]
  rowMetrics: AxisMetric[]
  mergeRanges: Array<NonNullable<ReturnType<typeof mergeBounds>>>
  selection: Selection
  viewportRef: RefObject<HTMLDivElement | null>
  fillDraggingRef: RefObject<boolean>
  displayValue: GridProps['displayValue']
  onSelection: GridProps['onSelection']
  onBeginEdit: GridProps['onBeginEdit']
  onCellValue: GridProps['onCellValue']
  onOpenHyperlink: GridProps['onOpenHyperlink']
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
      if (candidate && typeof candidate === 'object') return candidate as CellBorderSide
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
  merge,
  columns,
  frozenColumnCount,
  defaultColumn,
  defaultRow,
  columnMetrics,
  rowMetrics,
  mergeRanges,
  selection,
  viewportRef,
  fillDraggingRef,
  displayValue,
  onSelection,
  onBeginEdit,
  onCellValue,
  onOpenHyperlink,
}: GridCellViewProps) {
  const address = addressOf({ row, col })
  const left = (pane === 'body' ? HEADER_WIDTH : 0) + axisOffset(col, defaultColumn, columnMetrics)
  const top = (pane === 'body' ? HEADER_HEIGHT : 0) + axisOffset(row, defaultRow, rowMetrics)
  const width = merge
    ? axisOffset(merge.right + 1, defaultColumn, columnMetrics) - axisOffset(merge.left, defaultColumn, columnMetrics)
    : axisSize(col, defaultColumn, columnMetrics)
  const height = merge
    ? axisOffset(merge.bottom + 1, defaultRow, rowMetrics) - axisOffset(merge.top, defaultRow, rowMetrics)
    : axisSize(row, defaultRow, rowMetrics)
  const cell = mergedDisplayCell(sheet, merge, sheet.cells[address])
  const cellText = displayValue(sheet.id, address)
  const accounting = accountingDisplayParts(cellText, cell?.numFmt || cell?.style?.numFmt)
  const rawValue = cell?.formula ? cell.result : cell?.value
  const validation = validationForCell(sheet, address)
  const validationOptions = validationListOptions(validation)
  const isCheckbox = cell?.type === 'checkbox' || (
    typeof rawValue === 'boolean' && validationOptions.length === 2 &&
    validationOptions[0].toLocaleUpperCase() === 'TRUE' && validationOptions[1].toLocaleUpperCase() === 'FALSE'
  )
  const isDropdown = !isCheckbox && (cell?.type === 'dropdown' || validationOptions.length > 0)
  const horizontal = String(cell?.style?.alignment?.horizontal || '').toLocaleLowerCase()
  const canOverflow = Boolean(
    cellText && !merge && !cell?.style?.alignment?.wrapText && !cell?.style?.alignment?.clipText && !isCheckbox && !isDropdown &&
    typeof rawValue !== 'number' && typeof rawValue !== 'boolean' &&
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
      id={`cell-${sheet.id}-${address}`}
      className={`grid-cell${cell?.formula ? ' has-formula' : ''}${cell?.note ? ' has-note' : ''}${cell?.style?.fill ? ' has-fill' : ''}${canOverflow && contentWidth > width ? ' can-overflow' : ''}${pane !== 'body' ? ' is-frozen' : ''}`}
      style={{ left, top, width, height, ...cellCss(cell) }}
      role="gridcell"
      aria-rowindex={row + 1}
      aria-colindex={col + 1}
      title={noteText(cell?.note) || cell?.hyperlinkTooltip || (cell?.formula ? `=${cell.formula}` : cellText || undefined)}
      onMouseDown={(event) => {
        if (event.button !== 0) return
        event.preventDefault()
        viewportRef.current?.focus()
        const target = merge ? { row: merge.top, col: merge.left } : { row, col }
        const anchor = merge ? { row: merge.bottom, col: merge.right } : target
        onSelection(event.shiftKey
          ? { anchor: selection.anchor, focus: merge ? { row: merge.bottom, col: merge.right } : target }
          : { anchor, focus: target })
      }}
      onMouseEnter={(event) => {
        if (!fillDraggingRef.current && event.buttons === 1) onSelection({ anchor: selection.anchor, focus: merge ? { row: merge.bottom, col: merge.right } : { row, col } })
      }}
      onDoubleClick={() => onBeginEdit()}
    >
      {isCheckbox ? (
        <input
          type="checkbox"
          className="cell-checkbox"
          aria-label={`${address} checkbox`}
          checked={rawValue === true || String(rawValue).toLocaleUpperCase() === 'TRUE'}
          onMouseDown={(event) => {
            event.stopPropagation()
            viewportRef.current?.focus()
            onSelection({ anchor: { row, col }, focus: { row, col } })
          }}
          onChange={(event) => onCellValue(address, event.target.checked)}
        />
      ) : isDropdown ? (
        <select
          className="cell-dropdown"
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
        <button type="button" className="cell-content cell-hyperlink" style={cellContentCss(cell, canOverflow ? contentWidth : undefined)} onClick={(event) => { event.stopPropagation(); onOpenHyperlink(cell.hyperlink!) }}>{cellText || cell.hyperlink}</button>
      ) : (
        <span className={`cell-content${accounting ? ' cell-accounting' : ''}`} style={cellContentCss(cell, canOverflow ? contentWidth : undefined)}>
          {accounting
            ? <><span>{accounting.symbol}</span><span>{accounting.amount}</span></>
            : cellText}
        </span>
      )}
    </div>
  )
})

type GridCellRowProps = Omit<GridCellViewProps, 'row' | 'col' | 'merge'> & {
  coordinates: Array<{ row: number; col: number }>
  mergeMap: Map<string, ReturnType<typeof mergeBounds>>
}

const GridCellRow = memo(function GridCellRow({ coordinates, mergeMap, ...cellProps }: GridCellRowProps) {
  return <>{coordinates.map(({ row, col }) => (
    <GridCellView
      key={addressOf({ row, col })}
      {...cellProps}
      row={row}
      col={col}
      merge={mergeMap.get(addressOf({ row, col })) ?? undefined}
    />
  ))}</>
}, (previous, next) => {
  if (
    previous.sheet !== next.sheet || previous.pane !== next.pane || previous.columns !== next.columns ||
    previous.frozenColumnCount !== next.frozenColumnCount || previous.defaultColumn !== next.defaultColumn ||
    previous.defaultRow !== next.defaultRow || previous.columnMetrics !== next.columnMetrics ||
    previous.rowMetrics !== next.rowMetrics || previous.mergeRanges !== next.mergeRanges ||
    previous.selection !== next.selection || previous.viewportRef !== next.viewportRef ||
    previous.fillDraggingRef !== next.fillDraggingRef || previous.displayValue !== next.displayValue ||
    previous.onSelection !== next.onSelection || previous.onBeginEdit !== next.onBeginEdit ||
    previous.onCellValue !== next.onCellValue || previous.onOpenHyperlink !== next.onOpenHyperlink ||
    previous.coordinates.length !== next.coordinates.length
  ) return false
  return previous.coordinates.every((coordinate, index) => (
    coordinate.row === next.coordinates[index].row && coordinate.col === next.coordinates[index].col
  ))
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
}: GridProps) {
  const viewportRef = useRef<HTMLDivElement>(null)
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
  const selectionOutlineRef = useRef<HTMLDivElement>(null)
  const fillPreviewOverlayRef = useRef<HTMLDivElement>(null)
  const editorOverlayRef = useRef<HTMLTextAreaElement>(null)
  const [fillTarget, setFillTarget] = useState<Coord | null>(null)
  const [resize, setResize] = useState<DimensionResize | null>(null)
  const [freezeDrag, setFreezeDrag] = useState<FreezeDrag | null>(null)
  const [viewport, setViewport] = useState({ scrollLeft: 0, scrollTop: 0, width: 1000, height: 600 })
  const defaultColumn = (useImportedDefaults ? IMPORTED_COL_WIDTH : DEFAULT_COL_WIDTH) * zoom
  const defaultRow = (useImportedDefaults ? IMPORTED_ROW_HEIGHT : DEFAULT_ROW_HEIGHT) * zoom
  const syncPinnedOverlays = useCallback((scrollLeft: number, scrollTop: number) => {
    for (const element of [selectionOutlineRef.current, fillPreviewOverlayRef.current, editorOverlayRef.current]) {
      if (!element) continue
      const x = element.classList.contains('pin-x') ? scrollLeft : 0
      const y = element.classList.contains('pin-y') ? scrollTop : 0
      const transform = `translate3d(${x}px, ${y}px, 0)`
      if (element.style.transform !== transform) element.style.transform = transform
    }
  }, [])
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

  useEffect(() => {
    const viewportElement = viewportRef.current
    if (!viewportElement) return
    const update = () => {
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
  }, [syncPinnedOverlays])

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
      viewportRef.current?.focus()
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
      viewportRef.current?.focus()
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

  useEffect(() => {
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
  }, [selection.focus.col, selection.focus.row, defaultColumn, defaultRow, columnMetrics, rowMetrics, frozenColumnCount, frozenRowCount, frozenHeight, frozenWidth])

  const frozenRenderRowEnd = frozenRowCount
    ? Math.min(frozenRowCount - 1, axisIndexAt(Math.max(0, viewport.height - HEADER_HEIGHT), rows, defaultRow, rowMetrics) + 2)
    : -1
  const frozenRenderColumnEnd = frozenColumnCount
    ? Math.min(frozenColumnCount - 1, axisIndexAt(Math.max(0, viewport.width - HEADER_WIDTH), columns, defaultColumn, columnMetrics) + 2)
    : -1
  const renderedRows = [...new Set([
    ...Array.from({ length: Math.max(0, endRow - startRow + 1) }, (_, offset) => startRow + offset),
    ...Array.from({ length: frozenRenderRowEnd + 1 }, (_, row) => row),
  ])].filter((row) => !hiddenRows.has(row)).sort((a, b) => a - b)
  const renderedColumns = [...new Set([
    ...Array.from({ length: Math.max(0, endCol - startCol + 1) }, (_, offset) => startCol + offset),
    ...Array.from({ length: frozenRenderColumnEnd + 1 }, (_, col) => col),
  ])].filter((col) => !hiddenColumns.has(col)).sort((a, b) => a - b)

  const visibleCellMap = new Map<string, { row: number; col: number }>()
  for (const row of renderedRows) {
    for (const col of renderedColumns) visibleCellMap.set(addressOf({ row, col }), { row, col })
  }

  const mergeMap = new Map<string, ReturnType<typeof mergeBounds>>()
  for (const merge of mergeRanges) {
    const intersectsRows = renderedRows.some((row) => row >= merge.top && row <= merge.bottom)
    const intersectsColumns = renderedColumns.some((col) => col >= merge.left && col <= merge.right)
    if (!intersectsRows || !intersectsColumns) continue
    for (const row of renderedRows) {
      if (row < merge.top || row > merge.bottom) continue
      for (const col of renderedColumns) {
        if (col >= merge.left && col <= merge.right) mergeMap.set(addressOf({ row, col }), merge)
      }
    }
    const master = { row: merge.top, col: merge.left }
    visibleCellMap.set(addressOf(master), master)
    mergeMap.set(addressOf(master), merge)
  }
  const visibleCells = [...visibleCellMap.values()].sort((a, b) => a.row - b.row || a.col - b.col)

  const measurementContext = useMemo(() => document.createElement('canvas').getContext('2d'), [])
  const measureText = (text: string, cell?: CellData) => {
    const font = cell?.style?.font || {}
    const fontSize = Math.max(8, Number(font.size) || 11) * (4 / 3) * zoom
    if (measurementContext) {
      const family = String(font.name || 'Calibri').replace(/["\\]/g, '')
      measurementContext.font = `${font.italic ? 'italic ' : ''}${font.bold ? '700 ' : '400 '}${fontSize}px "${family}"`
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

  const bounds = selectionBounds(selection)
  const selectionPinnedX = bounds.right < frozenColumnCount
  const selectionPinnedY = bounds.bottom < frozenRowCount
  const selectionLeft = HEADER_WIDTH + axisOffset(bounds.left, defaultColumn, columnMetrics)
  const selectionTop = HEADER_HEIGHT + axisOffset(bounds.top, defaultRow, rowMetrics)
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
      viewportRef.current?.focus()
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

  const cellsByPane: Record<CellPane, Array<{ row: number; col: number }>> = {
    body: [],
    'frozen-row': [],
    'frozen-column': [],
    'frozen-corner': [],
  }
  for (const coordinate of visibleCells) {
    const address = addressOf(coordinate)
    const merge = mergeMap.get(address)
    if (merge && (merge.top !== coordinate.row || merge.left !== coordinate.col)) continue
    const frozenColumn = merge ? merge.right < frozenColumnCount : coordinate.col < frozenColumnCount
    const frozenRow = merge ? merge.bottom < frozenRowCount : coordinate.row < frozenRowCount
    const pane: CellPane = frozenRow && frozenColumn ? 'frozen-corner' : frozenRow ? 'frozen-row' : frozenColumn ? 'frozen-column' : 'body'
    cellsByPane[pane].push(coordinate)
  }

  const rowsByPane = Object.fromEntries((Object.keys(cellsByPane) as CellPane[]).map((pane) => {
    const grouped = new Map<number, Array<{ row: number; col: number }>>()
    for (const coordinate of cellsByPane[pane]) {
      const row = grouped.get(coordinate.row)
      if (row) row.push(coordinate)
      else grouped.set(coordinate.row, [coordinate])
    }
    return [pane, [...grouped.values()]]
  })) as Record<CellPane, Array<Array<{ row: number; col: number }>>>

  const renderGridRows = (pane: CellPane) => rowsByPane[pane].map((coordinates) => (
    <GridCellRow
      key={`${pane}-${coordinates[0].row}`}
      coordinates={coordinates}
      mergeMap={mergeMap}
      sheet={sheet}
      pane={pane}
      columns={columns}
      frozenColumnCount={frozenColumnCount}
      defaultColumn={defaultColumn}
      defaultRow={defaultRow}
      columnMetrics={columnMetrics}
      rowMetrics={rowMetrics}
      mergeRanges={mergeRanges}
      selection={selection}
      viewportRef={viewportRef}
      fillDraggingRef={fillDraggingRef}
      displayValue={displayValue}
      onSelection={onSelection}
      onBeginEdit={onBeginEdit}
      onCellValue={onCellValue}
      onOpenHyperlink={onOpenHyperlink}
    />
  ))

  const renderColumnHeader = (col: number) => {
    const width = axisSize(col, defaultColumn, columnMetrics)
    return (
      <div
        key={`ch-${col}`}
        className={`column-header${col < frozenColumnCount ? ' is-frozen' : ''}`}
        style={{ left: axisOffset(col, defaultColumn, columnMetrics), top: 0, width, height: HEADER_HEIGHT }}
        role="columnheader"
        aria-label={`Column ${columnName(col)}`}
        onMouseDown={(event) => {
          event.preventDefault()
          viewportRef.current?.focus()
          const usedBottom = Math.max(0, sheet.rowCount - 1)
          onSelection({ anchor: { row: 0, col }, focus: { row: usedBottom, col } })
        }}
      >
        <span className="dimension-header-label">{columnName(col)}</span>
        <button
          type="button"
          role="separator"
          aria-orientation="vertical"
          aria-label={`Resize column ${columnName(col)}`}
          aria-valuemin={MIN_COLUMN_PIXELS}
          aria-valuemax={MAX_COLUMN_PIXELS}
          aria-valuenow={Math.round(width / zoom)}
          title={`Drag to resize column ${columnName(col)}. Double-click to auto-fit.`}
          className="dimension-resize-handle is-column"
          data-resize-column={columnName(col)}
          onMouseDown={(event) => { event.preventDefault(); event.stopPropagation() }}
          onClick={(event) => event.stopPropagation()}
          onPointerDown={(event) => beginDimensionResize('column', col, event)}
          onDoubleClick={(event) => { event.preventDefault(); event.stopPropagation(); autoFitColumn(col) }}
          onKeyDown={(event) => resizeDimensionWithKey('column', col, event)}
        />
      </div>
    )
  }

  const renderRowHeader = (row: number) => {
    const height = axisSize(row, defaultRow, rowMetrics)
    return (
      <div
        key={`rh-${row}`}
        className={`row-header${row < frozenRowCount ? ' is-frozen' : ''}`}
        style={{ left: 0, top: axisOffset(row, defaultRow, rowMetrics), width: HEADER_WIDTH, height }}
        role="rowheader"
        aria-label={`Row ${row + 1}`}
        onMouseDown={(event) => {
          event.preventDefault()
          viewportRef.current?.focus()
          const usedRight = Math.max(0, sheet.colCount - 1)
          onSelection({ anchor: { row, col: 0 }, focus: { row, col: usedRight } })
        }}
      >
        <span className="dimension-header-label">{row + 1}</span>
        <button
          type="button"
          role="separator"
          aria-orientation="horizontal"
          aria-label={`Resize row ${row + 1}`}
          aria-valuemin={MIN_ROW_PIXELS}
          aria-valuemax={Math.round(MAX_ROW_POINTS * (4 / 3))}
          aria-valuenow={Math.round(height / zoom)}
          title={`Drag to resize row ${row + 1}. Double-click to auto-fit.`}
          className="dimension-resize-handle is-row"
          data-resize-row={row + 1}
          onMouseDown={(event) => { event.preventDefault(); event.stopPropagation() }}
          onClick={(event) => event.stopPropagation()}
          onPointerDown={(event) => beginDimensionResize('row', row, event)}
          onDoubleClick={(event) => { event.preventDefault(); event.stopPropagation(); autoFitRow(row) }}
          onKeyDown={(event) => resizeDimensionWithKey('row', row, event)}
        />
      </div>
    )
  }

  return (
    <div
      ref={viewportRef}
      className={`sheet-viewport${useImportedDefaults ? ' uses-imported-defaults' : ''}${showGridlines ? '' : ' hides-gridlines'}${showNotes ? '' : ' hides-notes'}`}
      tabIndex={0}
      role="grid"
      aria-label={`${sheet.name} spreadsheet grid`}
      aria-rowcount={rows}
      aria-colcount={columns}
      aria-activedescendant={`cell-${sheet.id}-${addressOf(selection.focus)}`}
      onKeyDown={onKeyDown}
      onScroll={(event) => handleViewportScroll(event.currentTarget)}
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
        {frozenRowCount > 0 && (
          <div className="frozen-pane-layer is-row" style={{ width: totalWidth - HEADER_WIDTH, height: frozenHeight }}>
            {renderGridRows('frozen-row')}
          </div>
        )}
        {frozenColumnCount > 0 && (
          <div className="frozen-pane-layer is-column" style={{ width: frozenWidth, height: totalHeight - HEADER_HEIGHT }}>
            {renderGridRows('frozen-column')}
          </div>
        )}
        {frozenRowCount > 0 && frozenColumnCount > 0 && (
          <div className="frozen-pane-layer is-corner" style={{ width: frozenWidth, height: frozenHeight }}>
            {renderGridRows('frozen-corner')}
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

        <div
          ref={selectionOutlineRef}
          className={`selection-outline${selectionPinnedX || selectionPinnedY ? ' scroll-pin' : ''}${selectionPinnedX ? ' pin-x' : ''}${selectionPinnedY ? ' pin-y' : ''}`}
          style={{ left: selectionLeft, top: selectionTop, width: selectionWidth, height: selectionHeight, transform: pinnedTransform(selectionPinnedX, selectionPinnedY) }}
          aria-hidden="true"
        ><span
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
        /></div>

        {editing && (() => {
          const coord = coordOf(editing.address)
          if (!coord) return null
          const merge = mergeMap.get(editing.address)
          const frozenColumn = merge ? merge.right < frozenColumnCount : coord.col < frozenColumnCount
          const frozenRow = merge ? merge.bottom < frozenRowCount : coord.row < frozenRowCount
          const left = HEADER_WIDTH + axisOffset(coord.col, defaultColumn, columnMetrics)
          const top = HEADER_HEIGHT + axisOffset(coord.row, defaultRow, rowMetrics)
          const width = merge
            ? axisOffset(merge.right + 1, defaultColumn, columnMetrics) - axisOffset(merge.left, defaultColumn, columnMetrics)
            : axisSize(coord.col, defaultColumn, columnMetrics)
          const height = merge
            ? axisOffset(merge.bottom + 1, defaultRow, rowMetrics) - axisOffset(merge.top, defaultRow, rowMetrics)
            : axisSize(coord.row, defaultRow, rowMetrics)
          return (
            <textarea
              ref={editorOverlayRef}
              autoFocus
              className={`cell-editor${frozenColumn || frozenRow ? ' scroll-pin' : ''}${frozenColumn ? ' pin-x' : ''}${frozenRow ? ' pin-y' : ''}`}
              style={{ left, top, width: Math.max(width, 160), minHeight: height, transform: pinnedTransform(frozenColumn, frozenRow) }}
              value={editing.draft}
              onChange={(event) => onDraft(event.target.value)}
              onFocus={(event) => event.currentTarget.select()}
              onBlur={onCommitEdit}
              onKeyDown={(event) => {
                if (event.key === 'Escape') { event.preventDefault(); onCancelEdit() }
                else if (event.key === 'Enter' && !event.altKey && !event.shiftKey) { event.preventDefault(); event.currentTarget.blur() }
                else if (event.key === 'Tab') { event.preventDefault(); event.currentTarget.blur() }
              }}
            />
          )
        })()}

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
  const [historyTick, setHistoryTick] = useState(0)
  const [aggregateMode, setAggregateMode] = useState<AggregateMode>('sum')
  const [searchOpen, setSearchOpen] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [searchIndex, setSearchIndex] = useState(-1)
  const [showFormulaBar, setShowFormulaBar] = useState(true)
  const [showFormulas, setShowFormulas] = useState(false)
  const [showNotes, setShowNotes] = useState(true)
  const [immersive, setImmersive] = useState(false)
  const historyRef = useRef<WorkbookModel[]>([])
  const futureRef = useRef<WorkbookModel[]>([])
  const internalClipboard = useRef<InternalClipboard | null>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const searchOriginRef = useRef('A1')
  const workbookRef = useRef(workbook)
  const documentRef = useRef(documentFile)
  const dirtyRef = useRef(dirty)

  useEffect(() => { workbookRef.current = workbook }, [workbook])
  useEffect(() => { documentRef.current = documentFile }, [documentFile])
  useEffect(() => { dirtyRef.current = dirty }, [dirty])
  useEffect(() => {
    if (!toast) return
    const timeout = window.setTimeout(() => setToast(''), 3800)
    return () => window.clearTimeout(timeout)
  }, [toast])

  const activeSheet = useMemo(() => workbook?.sheets.find((sheet) => sheet.id === workbook.activeSheetId) || workbook?.sheets[0] || null, [workbook])
  const gridlinesVisible = activeSheet?.views?.[0]?.showGridLines !== false
  const activeAddress = addressOf(selection.focus)
  const activeCell = activeSheet?.cells[activeAddress]

  useEffect(() => setFormulaDraft(rawCellValue(activeCell)), [activeSheet?.id, activeAddress, activeCell])

  const valueResolver = useMemo(() => workbook ? createValueResolver(workbook) : () => null, [workbook])
  const displayValue = useCallback((sheetId: string, address: string) => {
    if (!workbook) return ''
    const sheet = workbook.sheets.find((item) => item.id === sheetId)
    const cell = sheet?.cells[address]
    if (!cell) return ''
    if (showFormulas && cell.formula) return `=${cell.formula}`
    const value = cell.formula ? valueResolver(sheetId, address) : cell.value
    const cachedResultMatches = !cell.formula || Object.is(value, cell.result)
    return formatScalar(value, cell.numFmt || cell.style?.numFmt, cachedResultMatches ? cell.display : undefined)
  }, [showFormulas, workbook, valueResolver])

  const resetHistory = useCallback(() => {
    historyRef.current = []
    futureRef.current = []
    setHistoryTick((value) => value + 1)
  }, [])

  const mutateWorkbook = useCallback((mutator: (next: WorkbookModel) => void) => {
    const current = workbookRef.current
    if (!current) return
    historyRef.current.push(current)
    if (historyRef.current.length > MAX_HISTORY) historyRef.current.shift()
    futureRef.current = []
    const next = structuredClone(current) as WorkbookModel
    mutator(next)
    workbookRef.current = next
    setWorkbook(next)
    setDirty(true)
    setHistoryTick((value) => value + 1)
  }, [])

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
    const previous = historyRef.current.pop()
    if (!current || !previous) return
    futureRef.current.push(current)
    workbookRef.current = previous
    setWorkbook(previous)
    setDirty(true)
    setEditing(null)
    setHistoryTick((value) => value + 1)
  }, [])

  const redo = useCallback(() => {
    const current = workbookRef.current
    const next = futureRef.current.pop()
    if (!current || !next) return
    historyRef.current.push(current)
    workbookRef.current = next
    setWorkbook(next)
    setDirty(true)
    setEditing(null)
    setHistoryTick((value) => value + 1)
  }, [])

  const canDiscard = useCallback(() => !dirtyRef.current || window.confirm('Discard the unsaved changes in this spreadsheet?'), [])

  const applyPayload = useCallback((payload: WorkbookPayload) => {
    const normalized = normalizeWorkbook(payload.workbook, payload.name)
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
  }, [resetHistory])

  const newWorkbook = useCallback(async () => {
    if (!canDiscard()) return
    setBusy('Creating a clean spreadsheet…')
    try {
      const { documentId } = await window.simpleCalc.createWorkbook()
      const next = blankWorkbook()
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
  }, [canDiscard, resetHistory])

  const openWorkbook = useCallback(async () => {
    if (!canDiscard()) return
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
    if (!canDiscard()) return
    setBusy('Opening spreadsheet…')
    try {
      applyPayload(await window.simpleCalc.openPath(filePath))
    } catch (error) {
      setRecent((current) => {
        const next = current.filter((item) => item.path !== filePath)
        localStorage.setItem(RECENT_KEY, JSON.stringify(next))
        return next
      })
      setToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }, [applyPayload, canDiscard])

  const handleDrop = useCallback(async (event: DragEvent<HTMLElement>) => {
    event.preventDefault()
    const file = event.dataTransfer.files[0]
    if (!file || !canDiscard()) return
    setBusy(`Opening ${file.name}…`)
    try {
      applyPayload(await window.simpleCalc.openBytes(file.name, await file.arrayBuffer()))
    } catch (error) {
      setToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }, [applyPayload, canDiscard])

  const saveWorkbook = useCallback(async (saveAs = false, format: 'xlsx' | 'ods' | 'csv' | 'tsv' = 'xlsx') => {
    const currentWorkbook = workbookRef.current
    const currentDocument = documentRef.current
    if (!currentWorkbook || !currentDocument) return
    if (!dirtyRef.current && !saveAs && !currentDocument.requiresSaveAs && format === 'xlsx') {
      setToast('No changes to save')
      return
    }
    if ((format === 'csv' || format === 'tsv') && currentWorkbook.sheets.length > 1 && !window.confirm(`${format.toUpperCase()} keeps only the active sheet and cannot store formatting, merges, or multiple sheets. Continue?`)) return
    if (format === 'xlsx' && currentDocument.sourceFormat !== 'xlsx' && currentDocument.warnings.length && !window.confirm(`Convert this ${currentDocument.sourceFormat.toUpperCase()} workbook to XLSX? Values, formulas, and common formatting will be carried across, but unsupported native features may not have a cross-format equivalent and can be simplified or omitted.`)) return
    if (format === 'xlsx' && currentDocument.sourceFormat === 'xlsx' && dirtyRef.current && currentDocument.warnings.length && !window.confirm('This workbook contains features listed in the compatibility note that may be simplified after editing. Save the edited workbook anyway?')) return
    setBusy(format === 'xlsx' ? 'Saving workbook…' : `Exporting ${format.toUpperCase()}…`)
    try {
      const result = await window.simpleCalc.saveWorkbook({
        documentId: currentDocument.documentId,
        workbook: currentWorkbook,
        saveAs: saveAs || currentDocument.requiresSaveAs || format !== 'xlsx',
        format,
        suggestedName: currentDocument.name,
        sourceUnmodified: !dirtyRef.current,
      })
      if (!result) return
      setDocumentFile((current) => current ? {
        ...current,
        path: result.path,
        name: result.name,
        sourceFormat: result.format,
        requiresSaveAs: false,
        warnings: result.format === 'xlsx' ? [] : current.warnings,
      } : current)
      setWorkbook((current) => current ? { ...current, name: result.name } : current)
      setDirty(false)
      setRecent((current) => rememberRecent(current, { path: result.path, name: result.name, format: result.format, openedAt: Date.now() }))
      setToast(`Saved ${result.name}`)
    } catch (error) {
      setToast(errorMessage(error))
    } finally {
      setBusy('')
    }
  }, [])

  const closeWindow = useCallback(() => {
    if (dirtyRef.current && !window.confirm('Close simple_calc and discard unsaved changes?')) return
    window.simpleCalc.close()
  }, [])

  useEffect(() => window.simpleCalc.onCloseRequested(closeWindow), [closeWindow])
  useEffect(() => window.simpleCalc.onOpenExternal((filePath) => { void openPath(filePath) }), [openPath])

  const commitCell = useCallback((address: string, draft: string) => {
    if (!workbookRef.current) return
    const sheetId = workbookRef.current.activeSheetId
    const sheet = workbookRef.current.sheets.find((item) => item.id === sheetId)
    if (!sheet || rawCellValue(sheet.cells[address]) === draft) { setEditing(null); return }
    mutateWorkbook((next) => {
      const target = next.sheets.find((item) => item.id === next.activeSheetId)!
      const cell = parseDraft(draft, target.cells[address])
      if (hasCellContent(cell)) target.cells[address] = cell
      else delete target.cells[address]
      const coord = coordOf(address)
      if (coord) {
        target.rowCount = Math.max(target.rowCount, coord.row + 1)
        target.colCount = Math.max(target.colCount, coord.col + 1)
      }
    })
    setEditing(null)
  }, [mutateWorkbook])

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
    if (!activeSheet) return
    setEditing({ address: activeAddress, draft: initialDraft === undefined ? rawCellValue(activeSheet.cells[activeAddress]) : initialDraft })
  }, [activeAddress, activeSheet])

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
        const cell = sheet.cells[address] || {}
        const style = structuredClone(cell.style || {}) as CellStyle
        update(style)
        sheet.cells[address] = { ...cell, style }
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
    const nextFormat = adjustedDecimalFormat(activeCell?.numFmt || activeCell?.style?.numFmt, delta)
    if (!nextFormat) { setToast('Decimal controls are unavailable for date and time formats.'); return }
    setNumberFormat(nextFormat)
  }, [activeCell, setNumberFormat])

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
        const cell = sheet.cells[address] || {}
        const style = structuredClone(cell.style || {}) as CellStyle
        if (preset === 'clear') delete style.border
        else {
          const border = structuredClone(style.border || {}) as CellBorder
          if (preset === 'all') border.top = border.bottom = border.left = border.right = side
          if (preset === 'bottom') border.bottom = side
          if (preset === 'outer') {
            if (coord.row === bounds.top) border.top = side
            if (coord.row === bounds.bottom) border.bottom = side
            if (coord.col === bounds.left) border.left = side
            if (coord.col === bounds.right) border.right = side
          }
          style.border = border
        }
        sheet.cells[address] = { ...cell, style }
      })
    })
  }, [mutateWorkbook, selection])

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

  const mergeSelection = useCallback(() => {
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
    if (extraValues && !window.confirm(`Merge these cells? ${extraValues} non-leading ${extraValues === 1 ? 'value' : 'values'} will be removed.`)) return
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

  const applyAlternatingColors = useCallback(() => {
    const addresses = rangeAddresses(selection)
    if (!addresses.length) { setToast('That selection is too large to format at once.'); return }
    const bounds = selectionBounds(selection)
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      addresses.forEach((address) => {
        const coord = coordOf(address)!
        const cell = sheet.cells[address] || {}
        const style = structuredClone(cell.style || {}) as CellStyle
        style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: coord.row === bounds.top ? 'FFDCE9E0' : (coord.row - bounds.top) % 2 ? 'FFF4F7F4' : 'FFFFFFFF' } }
        if (coord.row === bounds.top) style.font = { ...(style.font || {}), bold: true, color: { argb: 'FF294936' } }
        sheet.cells[address] = { ...cell, style }
      })
    })
  }, [mutateWorkbook, selection])

  const insertLink = useCallback(() => {
    if (!activeSheet) return
    const existing = activeCell?.hyperlink || 'https://'
    const entered = window.prompt('Link URL or email address:', existing)?.trim()
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
    const text = window.prompt('Text to display:', currentText || entered)
    if (text === null) return
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      sheet.cells[activeAddress] = { ...(sheet.cells[activeAddress] || {}), value: text || entered, hyperlink: target, hyperlinkTooltip: target }
    })
  }, [activeAddress, activeCell, activeSheet, mutateWorkbook])

  const editAnnotation = useCallback((kind: 'note' | 'comment') => {
    const current = noteText(activeCell?.note)
    const value = window.prompt(kind === 'note' ? 'Cell note:' : 'Offline comment (saved as an Excel note):', current)
    if (value === null) return
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      const cell = { ...(sheet.cells[activeAddress] || {}) }
      if (!value.trim()) delete cell.note
      else cell.note = kind === 'note' ? value.trim() : { comments: [{ author: 'simple_calc', text: value.trim() }] }
      if (hasCellContent(cell)) sheet.cells[activeAddress] = cell
      else delete sheet.cells[activeAddress]
    })
  }, [activeAddress, activeCell?.note, mutateWorkbook])

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

  const insertDropdown = useCallback(() => {
    const entered = window.prompt('Dropdown options, separated by commas:', 'Option 1, Option 2')
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
  }, [mutateWorkbook, selection])

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

  const insertFunction = useCallback((name: string) => {
    const bounds = selectionBounds(selection)
    const multi = bounds.top !== bounds.bottom || bounds.left !== bounds.right
    if (!multi || name === 'IF') { beginEdit(`=${name}(`); return }
    const target = { row: Math.min(1_048_575, bounds.bottom + 1), col: bounds.left }
    const address = addressOf(target)
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      sheet.cells[address] = parseDraft(`=${name}(${rangeAddress(bounds)})`, sheet.cells[address])
      sheet.rowCount = Math.max(sheet.rowCount, target.row + 1)
    })
    setSelection({ anchor: target, focus: target })
  }, [beginEdit, mutateWorkbook, selection])

  const applyStructureCommand = useCallback((command: SelectionStructureCommand) => {
    const current = workbookRef.current
    if (!current) return
    try {
      const result = applySelectionStructureCommand(current, current.activeSheetId, selection, command)
      historyRef.current.push(current)
      if (historyRef.current.length > MAX_HISTORY) historyRef.current.shift()
      futureRef.current = []
      workbookRef.current = result.workbook
      setWorkbook(result.workbook)
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

  const copySelection = useCallback(async () => {
    if (!activeSheet) return
    const bounds = selectionBounds(selection)
    const total = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
    if (total > 100_000) { setToast('That selection is too large to copy at once.'); return }
    const matrix: CellData[][] = []
    const textRows: string[] = []
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      const cells: CellData[] = []
      const values: string[] = []
      for (let col = bounds.left; col <= bounds.right; col += 1) {
        const cell = structuredClone(activeSheet.cells[addressOf({ row, col })] || {}) as CellData
        cells.push(cell)
        values.push(escapeClipboard(rawCellValue(cell)))
      }
      matrix.push(cells)
      textRows.push(values.join('\t'))
    }
    const text = textRows.join('\r\n')
    try {
      await navigator.clipboard.writeText(text)
      internalClipboard.current = { text, origin: { row: bounds.top, col: bounds.left }, cells: matrix }
      setToast(`${total} ${total === 1 ? 'cell' : 'cells'} copied`)
    } catch (error) {
      setToast(errorMessage(error))
    }
  }, [activeSheet, selection])

  const pasteSelection = useCallback(async () => {
    if (!activeSheet) return
    try {
      const text = await navigator.clipboard.readText()
      const internal = internalClipboard.current?.text === text ? internalClipboard.current : null
      const rows = internal ? internal.cells : parseClipboardText(text).map((row) => row.map((value) => parseDraft(value)))
      if (!rows.length) return
      const origin = selection.focus
      mutateWorkbook((next) => {
        const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
        rows.forEach((rowCells, rowOffset) => rowCells.forEach((sourceCell, colOffset) => {
          const destination = { row: origin.row + rowOffset, col: origin.col + colOffset }
          const address = addressOf(destination)
          const cell = structuredClone(sourceCell) as CellData
          if (internal && cell.formula) {
            const source = { row: internal.origin.row + rowOffset, col: internal.origin.col + colOffset }
            cell.formula = shiftFormulaReferences(cell.formula, destination.row - source.row, destination.col - source.col)
            delete cell.formulaType
            delete cell.formulaRange
            delete cell.dynamicFormula
            delete cell.result
            delete cell.display
          }
          if (hasCellContent(cell)) sheet.cells[address] = cell
          else delete sheet.cells[address]
          sheet.rowCount = Math.max(sheet.rowCount, destination.row + 1)
          sheet.colCount = Math.max(sheet.colCount, destination.col + 1)
        }))
      })
      setSelection({ anchor: origin, focus: { row: origin.row + rows.length - 1, col: origin.col + Math.max(0, ...rows.map((row) => row.length - 1)) } })
    } catch (error) {
      setToast(errorMessage(error))
    }
  }, [activeSheet, mutateWorkbook, selection.focus])

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

  const searchMatches = useMemo(() => {
    const query = searchQuery.trim().toLocaleLowerCase()
    if (!activeSheet || !query) return []
    return Object.entries(activeSheet.cells).sort(([a], [b]) => {
      const ca = coordOf(a) || { row: 0, col: 0 }
      const cb = coordOf(b) || { row: 0, col: 0 }
      return ca.row - cb.row || ca.col - cb.col
    }).filter(([address, cell]) => [displayValue(activeSheet.id, address), cell.formula ? '' : rawCellValue(cell)]
      .some((value) => value.toLocaleLowerCase().includes(query)))
      .map(([address]) => address)
  }, [activeSheet, displayValue, searchQuery])

  const revealSearchMatch = useCallback((index: number) => {
    if (!searchMatches.length) return
    const normalized = (index + searchMatches.length) % searchMatches.length
    const coord = coordOf(searchMatches[normalized])
    if (!coord) return
    setSearchIndex(normalized)
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
    window.requestAnimationFrame(() => document.querySelector<HTMLElement>('.sheet-viewport')?.focus())
  }, [])

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
    const nextIndex = origin
      ? searchMatches.findIndex((address) => {
          const match = coordOf(address)
          return Boolean(match && (match.row > origin.row || (match.row === origin.row && match.col > origin.col)))
        })
      : 0
    revealSearchMatch(nextIndex >= 0 ? nextIndex : 0)
  }, [activeSheet?.id, revealSearchMatch, searchMatches, searchOpen, searchQuery])

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

  const renameSheet = useCallback((sheetId: string) => {
    if (!workbook) return
    const sheet = workbook.sheets.find((item) => item.id === sheetId)
    if (!sheet) return
    const name = window.prompt('Rename sheet:', sheet.name)?.trim()
    if (!name || name === sheet.name) return
    if (workbook.sheets.some((item) => item.id !== sheetId && item.name.toLocaleLowerCase() === name.toLocaleLowerCase())) { setToast('Sheet names must be unique.'); return }
    mutateWorkbook((next) => { next.sheets.find((item) => item.id === sheetId)!.name = name.slice(0, 31) })
  }, [mutateWorkbook, workbook])

  const deleteActiveSheet = useCallback(() => {
    if (!workbook || workbook.sheets.length <= 1) { setToast('A workbook needs at least one sheet.'); return }
    if (!window.confirm(`Delete “${activeSheet?.name}”? This can be undone.`)) return
    mutateWorkbook((next) => {
      const index = next.sheets.findIndex((sheet) => sheet.id === next.activeSheetId)
      next.sheets.splice(index, 1)
      next.activeSheetId = next.sheets[Math.max(0, index - 1)].id
    })
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
  }, [activeSheet?.name, mutateWorkbook, workbook])

  const duplicateActiveSheet = useCallback(() => {
    if (!activeSheet) return
    mutateWorkbook((next) => {
      const sourceIndex = next.sheets.findIndex((sheet) => sheet.id === next.activeSheetId)
      const copy = structuredClone(next.sheets[sourceIndex]) as SheetData
      copy.id = makeId()
      const base = `${copy.name} copy`
      let name = base
      let suffix = 2
      while (next.sheets.some((sheet) => sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())) name = `${base} ${suffix++}`
      copy.name = name.slice(0, 31)
      next.sheets.splice(sourceIndex + 1, 0, copy)
      next.activeSheetId = copy.id
    })
  }, [activeSheet, mutateWorkbook])

  const setFreeze = useCallback((axis: 'rows' | 'columns', count: number) => {
    mutateWorkbook((next) => {
      const sheet = next.sheets.find((item) => item.id === next.activeSheetId)!
      const frozen = { ...(sheet.frozen || {}), [axis]: Math.max(0, Math.trunc(count)) }
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

  const selectionStats = useMemo(() => {
    if (!activeSheet) return { selected: 0, count: 0, numeric: 0, sum: 0, average: 0, minimum: 0, maximum: 0, numFmt: undefined as string | undefined }
    const bounds = selectionBounds(selection)
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
      const value = cell.formula ? valueResolver(activeSheet.id, address) : cell.value
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
    } else {
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
  }, [activeSheet, selection, valueResolver])

  const aggregateValue = aggregateMode === 'count'
    ? selectionStats.count
    : aggregateMode === 'numeric'
      ? selectionStats.numeric
      : aggregateMode === 'average'
        ? selectionStats.average
        : aggregateMode === 'minimum'
          ? selectionStats.minimum
          : aggregateMode === 'maximum'
            ? selectionStats.maximum
            : selectionStats.sum
  const aggregateHasValue = aggregateMode === 'count' || aggregateMode === 'numeric' || selectionStats.numeric > 0
  const aggregateText = aggregateHasValue
    ? (aggregateMode === 'count' || aggregateMode === 'numeric'
        ? aggregateValue.toLocaleString()
        : formatScalar(aggregateValue, selectionStats.numFmt, aggregateValue.toLocaleString(undefined, { maximumFractionDigits: 6 })))
    : '—'

  const moveSelection = useCallback((rowDelta: number, colDelta: number, extend = false) => {
    const next = {
      row: clamp(selection.focus.row + rowDelta, 0, 1_048_575),
      col: clamp(selection.focus.col + colDelta, 0, 16_383),
    }
    setSelection((current) => ({ anchor: extend ? current.anchor : next, focus: next }))
  }, [selection.focus.col, selection.focus.row])

  const handleGridKeyDown = useCallback((event: KeyboardEvent<HTMLDivElement>) => {
    if (editing) return
    const control = event.ctrlKey || event.metaKey
    if (event.altKey && event.key === 'Enter' && activeCell?.hyperlink) { event.preventDefault(); openHyperlink(activeCell.hyperlink); return }
    if (event.key === ' ' && activeSheet && activeCell) {
      const options = validationListOptions(validationForCell(activeSheet, activeAddress))
      const checkbox = activeCell.type === 'checkbox' || (options[0]?.toLocaleUpperCase() === 'TRUE' && options[1]?.toLocaleUpperCase() === 'FALSE')
      if (checkbox) { event.preventDefault(); setCellValue(activeAddress, !(activeCell.value === true || String(activeCell.value).toLocaleUpperCase() === 'TRUE')); return }
    }
    if (control && event.key.toLocaleLowerCase() === 'd') { event.preventDefault(); fillSelectedRange('down'); return }
    if (control && event.key.toLocaleLowerCase() === 'r') { event.preventDefault(); fillSelectedRange('right'); return }
    if (control && event.key.toLocaleLowerCase() === 'c') { event.preventDefault(); void copySelection(); return }
    if (control && event.key.toLocaleLowerCase() === 'x') { event.preventDefault(); void copySelection().then(clearSelection); return }
    if (control && event.key.toLocaleLowerCase() === 'v') { event.preventDefault(); void pasteSelection(); return }
    if (control && event.key.toLocaleLowerCase() === 'a') {
      event.preventDefault()
      setSelection({ anchor: { row: 0, col: 0 }, focus: { row: Math.max(0, (activeSheet?.rowCount || 1) - 1), col: Math.max(0, (activeSheet?.colCount || 1) - 1) } })
      return
    }
    if (event.key === 'F2') { event.preventDefault(); beginEdit(); return }
    if (event.key === 'Enter') { event.preventDefault(); beginEdit(); return }
    if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); clearSelection(); return }
    if (event.key === 'ArrowUp') { event.preventDefault(); moveSelection(-1, 0, event.shiftKey); return }
    if (event.key === 'ArrowDown') { event.preventDefault(); moveSelection(1, 0, event.shiftKey); return }
    if (event.key === 'ArrowLeft') { event.preventDefault(); moveSelection(0, -1, event.shiftKey); return }
    if (event.key === 'ArrowRight') { event.preventDefault(); moveSelection(0, 1, event.shiftKey); return }
    if (event.key === 'Tab') { event.preventDefault(); moveSelection(0, event.shiftKey ? -1 : 1); return }
    if (!control && !event.altKey && event.key.length === 1) { event.preventDefault(); beginEdit(event.key) }
  }, [activeAddress, activeCell, activeSheet, beginEdit, clearSelection, copySelection, editing, fillSelectedRange, moveSelection, openHyperlink, pasteSelection, setCellValue])

  useEffect(() => {
    const handleKey = (event: globalThis.KeyboardEvent) => {
      const control = event.ctrlKey || event.metaKey
      const typing = event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement || event.target instanceof HTMLSelectElement
      if (control && event.key.toLocaleLowerCase() === 'o') { event.preventDefault(); void openWorkbook(); return }
      if (control && event.key.toLocaleLowerCase() === 's') { event.preventDefault(); void saveWorkbook(event.shiftKey); return }
      if (control && event.key.toLocaleLowerCase() === 'z' && !typing) { event.preventDefault(); undo(); return }
      if (control && event.key.toLocaleLowerCase() === 'y' && !typing) { event.preventDefault(); redo(); return }
      if (control && event.key.toLocaleLowerCase() === 'f') { event.preventDefault(); openSearch(); return }
      if (control && event.key.toLocaleLowerCase() === 'k' && !typing) { event.preventDefault(); insertLink(); return }
      if (control && (event.key === '`' || event.key === '~') && !typing) { event.preventDefault(); setShowFormulas((value) => !value); return }
      if (control && event.key === '\\' && !typing) { event.preventDefault(); clearFormatting(); return }
      if (control && event.key.toLocaleLowerCase() === 'b' && !typing) { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), bold: !activeCell?.style?.font?.bold } }); return }
      if (control && event.key.toLocaleLowerCase() === 'i' && !typing) { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), italic: !activeCell?.style?.font?.italic } }); return }
      if (control && event.key.toLocaleLowerCase() === 'u' && !typing) { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), underline: !activeCell?.style?.font?.underline } }); return }
      if (event.altKey && event.shiftKey && event.key === '5' && !typing) { event.preventDefault(); applyStyle((style) => { style.font = { ...(style.font || {}), strike: !activeCell?.style?.font?.strike } }); return }
      if (event.ctrlKey && event.altKey && event.key.toLocaleLowerCase() === 'm' && !typing) { event.preventDefault(); editAnnotation('comment'); return }
      if (event.shiftKey && event.key === 'F2' && !typing) { event.preventDefault(); editAnnotation('note'); return }
      if (event.shiftKey && event.key === 'F11' && !typing) { event.preventDefault(); addSheet(); return }
    }
    window.addEventListener('keydown', handleKey)
    return () => window.removeEventListener('keydown', handleKey)
  }, [activeCell, addSheet, applyStyle, clearFormatting, editAnnotation, insertLink, openSearch, openWorkbook, redo, saveWorkbook, undo])

  const removeRecent = useCallback((filePath: string) => {
    setRecent((current) => {
      const next = current.filter((item) => item.path !== filePath)
      localStorage.setItem(RECENT_KEY, JSON.stringify(next))
      return next
    })
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
  const bounds = selectionBounds(selection)
  const selectedRowCount = bounds.bottom - bounds.top + 1
  const selectedColumnCount = bounds.right - bounds.left + 1
  const hiddenSheets = workbook.sheets.filter((sheet) => sheet.state === 'hidden')
  const menuDefinitions: SpreadsheetMenuDefinition[] = [
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
        { id: 'insert-sheet', label: 'Sheet', shortcut: 'Shift+F11', icon: <FilePlus2 size={13} />, action: addSheet },
        {
          id: 'insert-function', label: 'Function', separatorBefore: true, icon: <FunctionSquare size={13} />, children: ['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'IF'].map((name) => ({ id: `function-${name.toLocaleLowerCase()}`, label: name, action: () => insertFunction(name) })),
        },
        { id: 'insert-link', label: 'Link', shortcut: 'Ctrl+K', icon: <Link size={13} />, action: insertLink },
        { id: 'insert-checkbox', label: 'Checkbox', separatorBefore: true, icon: <CheckSquare size={13} />, action: insertCheckboxes },
        { id: 'insert-dropdown', label: 'Dropdown', icon: <ListPlus size={13} />, action: insertDropdown },
        { id: 'insert-comment', label: 'Comment', shortcut: 'Ctrl+Alt+M', separatorBefore: true, icon: <MessageSquarePlus size={13} />, action: () => editAnnotation('comment') },
        { id: 'insert-note', label: 'Note', shortcut: 'Shift+F2', icon: <StickyNote size={13} />, action: () => editAnnotation('note') },
      ],
    },
    {
      id: 'format',
      label: 'Format',
      items: [
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
          id: 'format-font-size', label: 'Font size', children: [8, 9, 10, 11, 12, 14, 16, 18, 24, 32].map((size) => ({ id: `font-size-${size}`, label: `${size}`, checked: (activeStyle.font?.size || 11) === size, action: () => applyStyle((style) => { style.font = { ...(style.font || {}), size } }) })),
        },
        {
          id: 'format-merge', label: 'Merge cells', separatorBefore: true, icon: <TableCellsMerge size={13} />, children: [
            { id: 'merge-all', label: 'Merge all', action: mergeSelection },
            { id: 'unmerge', label: 'Unmerge', icon: <TableCellsSplit size={13} />, action: unmergeSelection },
          ],
        },
        {
          id: 'format-borders', label: 'Borders', icon: <Grid2X2 size={13} />, children: [
            { id: 'borders-all', label: 'All borders', action: () => setBorderPreset('all') },
            { id: 'borders-outer', label: 'Outer border', action: () => setBorderPreset('outer') },
            { id: 'borders-bottom', label: 'Bottom border', action: () => setBorderPreset('bottom') },
            { id: 'borders-clear', label: 'Clear borders', action: () => setBorderPreset('clear'), separatorBefore: true },
          ],
        },
        { id: 'format-alternating', label: 'Alternating colors', separatorBefore: true, action: applyAlternatingColors },
        { id: 'format-clear', label: 'Clear formatting', shortcut: 'Ctrl+\\', icon: <Eraser size={13} />, action: clearFormatting },
      ],
    },
    {
      id: 'data',
      label: 'Data',
      items: [
        { id: 'data-checkbox', label: 'Add checkbox validation', icon: <CheckSquare size={13} />, action: insertCheckboxes },
        { id: 'data-dropdown', label: 'Add dropdown validation', icon: <ListPlus size={13} />, action: insertDropdown },
        { id: 'data-remove-validation', label: 'Remove validation', separatorBefore: true, action: removeValidation },
      ],
    },
    {
      id: 'tools',
      label: 'Tools',
      items: [
        { id: 'tools-find', label: 'Find in this sheet', shortcut: 'Ctrl+F', icon: <Search size={13} />, action: openSearch },
        { id: 'tools-show-formulas', label: 'Show formulas', shortcut: 'Ctrl+~', checked: showFormulas, icon: <FunctionSquare size={13} />, action: () => setShowFormulas((value) => !value) },
      ],
    },
  ]
  return (
    <div className={`app-root workbook-app${immersive ? ' is-fullscreen' : ''}`} onDragOver={(event) => event.preventDefault()} onDrop={handleDrop}>
      {titleBar}
      <SpreadsheetMenus menus={menuDefinitions} />
      <div className="command-bar">
        <div className="command-group file-group">
          <IconButton label="New spreadsheet" onClick={() => { void newWorkbook() }}><FilePlus2 size={16} /></IconButton>
          <IconButton label="Open spreadsheet (Ctrl+O)" onClick={() => { void openWorkbook() }}><FolderOpen size={16} /></IconButton>
          <button type="button" className="save-command" onClick={() => { void saveWorkbook(false) }} title="Save (Ctrl+S)"><Save size={16} /><span>Save</span></button>
          <label className="export-select" title="Save as or export">
            <ChevronDown size={13} />
            <select aria-label="Save as or export" value="" onChange={(event) => {
              const value = event.target.value as 'xlsx' | 'ods' | 'csv' | 'tsv'
              if (value) void saveWorkbook(true, value)
              event.target.value = ''
            }}>
              <option value="" disabled>Save as…</option>
              <option value="xlsx">Excel workbook (.xlsx)</option>
              <option value="ods">OpenDocument (.ods)</option>
              <option value="csv">Active sheet (.csv)</option>
              <option value="tsv">Active sheet (.tsv)</option>
            </select>
          </label>
        </div>
        <span className="command-separator" />
        <div className="command-group">
          <IconButton label="Undo (Ctrl+Z)" disabled={!historyRef.current.length} onClick={undo}><Undo2 size={16} /></IconButton>
          <IconButton label="Redo (Ctrl+Y)" disabled={!futureRef.current.length} onClick={redo}><Redo2 size={16} /></IconButton>
          <IconButton label="Copy" onClick={() => { void copySelection() }}><Copy size={15} /></IconButton>
          <IconButton label="Find (Ctrl+F)" active={searchOpen} onClick={openSearch}><Search size={15} /></IconButton>
        </div>
        <span className="command-separator" />
        <div className="command-group quick-number-group">
          <select className="toolbar-zoom-select" aria-label="Zoom" value={Math.round(zoom * 100)} onChange={(event) => setZoom(Number(event.target.value) / 100)}>
            {[50, 75, 90, 100, 125, 150, 200].map((percent) => <option key={percent} value={percent}>{percent}%</option>)}
          </select>
          <IconButton label="Currency format" onClick={() => setNumberFormat('$#,##0.00')}><DollarSign size={14} /></IconButton>
          <IconButton label="Percent format" onClick={() => setNumberFormat('0.00%')}><Percent size={14} /></IconButton>
          <IconButton label="Decrease decimal places" onClick={() => adjustDecimals(-1)}><span className="decimal-tool">.0‹</span></IconButton>
          <IconButton label="Increase decimal places" onClick={() => adjustDecimals(1)}><span className="decimal-tool">.00›</span></IconButton>
        </div>
        <span className="command-separator" />
        <div className="command-group format-group">
          <select
            className="font-select"
            aria-label="Font"
            value={activeStyle.font?.name || 'Aptos'}
            onChange={(event) => applyStyle((style) => { style.font = { ...(style.font || {}), name: event.target.value } })}
          >
            <option>Aptos</option><option>Arial</option><option>Calibri</option><option>Georgia</option><option>Segoe UI</option><option>Times New Roman</option>
          </select>
          <IconButton label="Decrease font size" onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), size: clamp((style.font?.size || activeStyle.font?.size || 11) - 1, 6, 72) } })}><Minus size={13} /></IconButton>
          <select
            className="size-select"
            aria-label="Font size"
            value={String(activeStyle.font?.size || 11)}
            onChange={(event) => applyStyle((style) => { style.font = { ...(style.font || {}), size: Number(event.target.value) } })}
          >
            {[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32].map((size) => <option key={size}>{size}</option>)}
          </select>
          <IconButton label="Increase font size" onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), size: clamp((style.font?.size || activeStyle.font?.size || 11) + 1, 6, 72) } })}><Plus size={13} /></IconButton>
          <IconButton label="Bold (Ctrl+B)" active={Boolean(activeStyle.font?.bold)} onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), bold: !activeStyle.font?.bold } })}><Bold size={15} /></IconButton>
          <IconButton label="Italic (Ctrl+I)" active={Boolean(activeStyle.font?.italic)} onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), italic: !activeStyle.font?.italic } })}><Italic size={15} /></IconButton>
          <IconButton label="Underline (Ctrl+U)" active={Boolean(activeStyle.font?.underline)} onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), underline: !activeStyle.font?.underline } })}><Underline size={15} /></IconButton>
          <IconButton label="Strikethrough (Alt+Shift+5)" active={Boolean(activeStyle.font?.strike)} onClick={() => applyStyle((style) => { style.font = { ...(style.font || {}), strike: !activeStyle.font?.strike } })}><Strikethrough size={15} /></IconButton>
          <label className="color-tool" title="Text color"><Type size={15} /><input type="color" value={cssColor(activeStyle.font?.color, '#20211f')} onChange={(event) => applyStyle((style) => { style.font = { ...(style.font || {}), color: { argb: `FF${event.target.value.slice(1).toUpperCase()}` } } })} /></label>
          <label className="color-tool" title="Fill color"><PaintBucket size={15} /><input type="color" value={cssColor(activeStyle.fill?.color || activeStyle.fill?.fgColor, '#ffffff')} onChange={(event) => applyStyle((style) => { style.fill = { ...(style.fill || {}), type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${event.target.value.slice(1).toUpperCase()}` } } })} /></label>
        </div>
        <span className="command-separator" />
        <div className="command-group">
          <IconButton label="Align left" active={!activeStyle.alignment?.horizontal || activeStyle.alignment.horizontal === 'left'} onClick={() => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'left' } })}><AlignLeft size={15} /></IconButton>
          <IconButton label="Align center" active={activeStyle.alignment?.horizontal === 'center'} onClick={() => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'center' } })}><AlignCenter size={15} /></IconButton>
          <IconButton label="Align right" active={activeStyle.alignment?.horizontal === 'right'} onClick={() => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), horizontal: 'right' } })}><AlignRight size={15} /></IconButton>
          <select className="number-select" aria-label="Number format" value={activeCell?.numFmt || activeStyle.numFmt || 'General'} onChange={(event) => setNumberFormat(event.target.value)}>
            <option value="General">General</option>
            <option value="0">Number</option>
            <option value="0.00">Number · 2 decimals</option>
            <option value="#,#0">Thousands</option>
            <option value="0%">Percent</option>
            <option value="0.00%">Percent · 2 decimals</option>
            <option value="$#,##0.00">Currency</option>
            <option value="m/d/yyyy">Date</option>
          </select>
          <IconButton label="All borders" onClick={() => setBorderPreset('all')}><Grid2X2 size={15} /></IconButton>
          <IconButton label="Merge selected cells" onClick={mergeSelection}><TableCellsMerge size={15} /></IconButton>
          <label className="toolbar-icon-select" title="Vertical alignment">
            {activeStyle.alignment?.vertical === 'top' ? <AlignVerticalJustifyStart size={15} /> : activeStyle.alignment?.vertical === 'middle' ? <AlignVerticalJustifyCenter size={15} /> : <AlignVerticalJustifyEnd size={15} />}
            <select aria-label="Vertical alignment" value={activeStyle.alignment?.vertical || 'bottom'} onChange={(event) => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), vertical: event.target.value } })}>
              <option value="top">Top</option><option value="middle">Middle</option><option value="bottom">Bottom</option>
            </select>
          </label>
          <label className="toolbar-icon-select" title="Text wrapping">
            <WrapText size={15} />
            <select aria-label="Text wrapping" value={activeStyle.alignment?.wrapText ? 'wrap' : activeStyle.alignment?.clipText ? 'clip' : 'overflow'} onChange={(event) => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), wrapText: event.target.value === 'wrap', clipText: event.target.value === 'clip' } })}>
              <option value="overflow">Overflow</option><option value="wrap">Wrap</option><option value="clip">Clip</option>
            </select>
          </label>
          <label className="toolbar-icon-select" title="Text rotation">
            <RotateCw size={15} />
            <select aria-label="Text rotation" value={String(activeStyle.alignment?.textRotation || 0)} onChange={(event) => applyStyle((style) => { style.alignment = { ...(style.alignment || {}), textRotation: event.target.value === 'vertical' ? 'vertical' : Number(event.target.value) } })}>
              <option value="0">None</option><option value="45">Tilt up</option><option value="-45">Tilt down</option><option value="vertical">Vertical</option>
            </select>
          </label>
          <IconButton label="Insert link (Ctrl+K)" onClick={insertLink}><Link size={15} /></IconButton>
          <IconButton label="Add note (Shift+F2)" onClick={() => editAnnotation('note')}><StickyNote size={15} /></IconButton>
          <label className="toolbar-icon-select" title="Insert function">
            <Sigma size={15} />
            <select aria-label="Insert function" value="" onChange={(event) => { if (event.target.value) insertFunction(event.target.value); event.target.value = '' }}>
              <option value="" disabled>Function</option>{['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'IF'].map((name) => <option key={name} value={name}>{name}</option>)}
            </select>
          </label>
        </div>
        <div className="command-spacer" />
        <span className="format-badge">{documentFile.sourceFormat.toUpperCase()}</span>
      </div>

      {documentFile.warnings.length > 0 && (
        <div className="compatibility-bar">
          <CircleAlert size={14} />
          <span><strong>Compatibility note:</strong> {documentFile.warnings[0]}</span>
          {documentFile.warnings.length > 1 && <button type="button" onClick={() => setWarningOpen((open) => !open)}>{warningOpen ? 'Hide' : `Show all ${documentFile.warnings.length}`}</button>}
          {warningOpen && <div className="compatibility-details">{documentFile.warnings.map((warning) => <p key={warning}>{warning}</p>)}</div>}
        </div>
      )}

      {(showFormulaBar || searchOpen) && <div className={`formula-bar${showFormulaBar ? '' : ' is-search-only'}`}>
        <div className="name-box">{activeAddress}</div>
        <Sigma size={14} />
        <input
          aria-label="Formula bar"
          value={formulaDraft}
          onChange={(event) => setFormulaDraft(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') { event.preventDefault(); commitCell(activeAddress, formulaDraft); (event.currentTarget as HTMLInputElement).blur() }
            else if (event.key === 'Escape') { setFormulaDraft(rawCellValue(activeCell)); (event.currentTarget as HTMLInputElement).blur() }
          }}
        />
        {searchOpen && (
          <div
            className="search-panel"
            role="search"
            aria-label="Find in this sheet"
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
            <button type="button" aria-label="Close find" title="Close (Esc)" onClick={closeSearch}><X size={13} /></button>
          </div>
        )}
      </div>}

      <SpreadsheetGrid
        sheet={activeSheet}
        selection={selection}
        editing={editing}
        zoom={zoom}
        useImportedDefaults={Boolean(workbook.metadata?.sourceName)}
        showGridlines={gridlinesVisible}
        showNotes={showNotes}
        displayValue={displayValue}
        onSelection={(next) => { setSelection(next); if (editing) setEditing(null) }}
        onBeginEdit={beginEdit}
        onDraft={(draft) => setEditing((current) => current ? { ...current, draft } : current)}
        onCommitEdit={() => { if (editing) commitCell(editing.address, editing.draft) }}
        onCancelEdit={() => setEditing(null)}
        onFill={autofillSelection}
        onCellValue={setCellValue}
        onOpenHyperlink={openHyperlink}
        onColumnResize={resizeColumn}
        onRowResize={resizeRow}
        onFreeze={setFreeze}
        onKeyDown={handleGridKeyDown}
      />

      <div className="sheet-strip">
        <div className="sheet-actions">
          <IconButton label="Add sheet" onClick={addSheet}><Plus size={15} /></IconButton>
        </div>
        <div className="sheet-tabs">
          {workbook.sheets.filter((sheet) => sheet.state === 'visible').map((sheet) => (
            <button
              type="button"
              key={sheet.id}
              className={`sheet-tab${sheet.id === workbook.activeSheetId ? ' is-active' : ''}`}
              onClick={() => {
                setWorkbook((current) => current ? { ...current, activeSheetId: sheet.id } : current)
                setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
                setEditing(null)
              }}
              onDoubleClick={() => renameSheet(sheet.id)}
              title="Double-click to rename"
            >{sheet.name}</button>
          ))}
        </div>
        <div className="sheet-menu-actions">
          <IconButton label="Duplicate active sheet" onClick={duplicateActiveSheet}><Copy size={14} /></IconButton>
          <IconButton label="Delete active sheet" onClick={deleteActiveSheet}><Trash2 size={14} /></IconButton>
        </div>
      </div>

      <footer className="statusbar">
        <div className="status-left"><span className="ready-dot" /> <span>{dirty ? 'Unsaved changes' : 'Ready'}</span><span className="status-divider" /><span>{workbook.sheets.length} {workbook.sheets.length === 1 ? 'sheet' : 'sheets'}</span>{activeCell?.formula && <><span className="status-divider" /><span>Formula</span></>}</div>
        <div className="status-right">
          {selectionStats.selected > 1 && (
            <label className="aggregate-picker" title="Quick calculation for the selected cells">
              <select aria-label="Selected cell calculation" aria-describedby="quick-calc-output" value={aggregateMode} onChange={(event) => setAggregateMode(event.target.value as AggregateMode)}>
                <option value="sum">Sum</option>
                <option value="average">Average</option>
                <option value="minimum">Minimum</option>
                <option value="maximum">Maximum</option>
                <option value="count">Count</option>
                <option value="numeric">Count numbers</option>
              </select>
              <output id="quick-calc-output" aria-live="polite">{aggregateText}</output>
            </label>
          )}
          <div className="zoom-control">
            <button type="button" aria-label="Zoom out" onClick={() => setZoom((value) => clamp(value - 0.1, 0.5, 2))}><Minus size={13} /></button>
            <input type="range" min="50" max="200" step="10" value={Math.round(zoom * 100)} aria-label="Zoom" onChange={(event) => setZoom(Number(event.target.value) / 100)} />
            <button type="button" aria-label="Zoom in" onClick={() => setZoom((value) => clamp(value + 0.1, 0.5, 2))}><Plus size={13} /></button>
            <span>{Math.round(zoom * 100)}%</span>
          </div>
        </div>
      </footer>

      {busy && <div className="busy-overlay"><div className="busy-card"><span className="spinner" />{busy}</div></div>}
      {toast && <div className="toast"><span>{toast}</span><button type="button" aria-label="Dismiss" onClick={() => setToast('')}><X size={13} /></button></div>}
      <span className="history-signal" aria-hidden="true">{historyTick}</span>
    </div>
  )
}
