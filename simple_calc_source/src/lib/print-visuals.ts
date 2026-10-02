/**
 * `input.visuals` for renderPrintPreview / printWorkbook / exportWorkbook (PDF and HTML):
 * conditional formatting, table styles and banding, number-format colours, checkboxes and
 * sparklines, resolved per printed cell with the grid's own rules (resolveCellVisual) and
 * serialized for electron/spreadsheet-print.cjs, which validates every value again.
 *
 * Styles are sent once in an indexed table (cells refer to them by index) so a large print
 * stays small; icons are serialized ConditionalIcon glyphs, sparklines SVG drawn at the
 * printed cell size.
 */
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import type { CellScalar, SheetData, WorkbookModel } from '../spreadsheet-types'
import { ConditionalIcon } from '../components/ConditionalIcon'
import { printedAnchorSize } from './chart-render'
import { columnIndex, computeConditionalFormats, dataBarBackground, iconSetSize, normalizeIconSetName } from './conditional-format'
import type { CellBounds, ConditionalCellFormat } from './conditional-format'
import { formatScalar } from './number-format'
import { renderSparklineSvg } from './sparkline-render'
import { parsedTables, tableAt, tableCellPaint } from './table-styles'
import { findValidation, literalListItems } from './validation'
import { conditionalFormatHost, createSparklineResolver, excelToday, formulaSparkline, resolveCellVisual, workbookCssColor, workbookThemeColors } from './visual-style'
import type { CellVisual, VisualBorderSide, VisualValueEngine } from './visual-style'

/** One resolved style as the print engine receives it. */
export interface PrintCellVisual {
  fill?: string
  color?: string
  bold?: boolean
  italic?: boolean
  decoration?: string
  borders?: Partial<Record<VisualBorderSide, string>>
  /** Data bar as CSS background layers (conditional-format dataBarBackground). */
  bar?: { image: string; size: string; position: string; repeat: string }
  /** Key into PrintVisualPayload.icons ("3Arrows:2"). */
  icon?: string
  /** Show bar/icon only: the value is not printed. */
  hide?: boolean
  /** Printed as a checkbox; the value is its state. */
  checkbox?: boolean
}

export interface PrintSheetVisuals {
  styles: PrintCellVisual[]
  /** Address -> index into styles. */
  cells: Record<string, number>
  /** Address -> the value as a conditional number format shows it. */
  text?: Record<string, string>
  /** Address -> sparkline SVG at the printed cell size. */
  sparklines?: Record<string, string>
}

export interface PrintVisualPayload {
  /** "set:index" -> ConditionalIcon SVG markup. */
  icons: Record<string, string>
  sheets: Record<string, PrintSheetVisuals>
}

export interface PrintVisualOptions {
  engine: VisualValueEngine
  /** Sheets that will print (default: every visible sheet). */
  sheetIds?: readonly string[]
  /** 0-based bounds when printing the selection (the active sheet only). */
  selection?: CellBounds | null
  /** The grid's colour resolver; defaults to the workbook theme. */
  cssColor?: (color: unknown, fallback?: string) => string
  /** Theme colours for table styles (OOXML order); defaults to workbook.metadata.themeColors. */
  themeColors?: readonly string[]
  /** Options of a dropdown list rule (the grid's listOptionsFor); defaults to literal items. */
  listOptions?: (sheet: SheetData, validation: Record<string, unknown>, row: number, col: number) => readonly string[]
  /** View > Show formulas: formula text prints instead of conditionally formatted values. */
  showFormulas?: boolean
  /** Excel serial for "today" (timePeriod rules); defaults to the current date. */
  today?: number
}

/** Icon size on paper, in CSS px at 100% (the grid draws 13px at 100% zoom). */
const PRINT_ICON_SIZE = 13
const MAX_VISUAL_CELLS = 400_000
const CELL_ADDRESS = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/

function position(address: string): { row: number; col: number } | null {
  const match = CELL_ADDRESS.exec(address)
  return match ? { row: Number(match[2]) - 1, col: columnIndex(match[1]) } : null
}

function addressOf(row: number, col: number) {
  let value = col + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return `${label}${row + 1}`
}

function parseBounds(reference: string): CellBounds | null {
  const [start, end = start] = reference.replace(/^.*!/, '').replace(/\$/g, '').trim().toUpperCase().split(':')
  const first = position(start)
  const second = position(end)
  if (!first || !second) return null
  return {
    top: Math.min(first.row, second.row),
    bottom: Math.max(first.row, second.row),
    left: Math.min(first.col, second.col),
    right: Math.max(first.col, second.col),
  }
}

function include(target: CellBounds | null, bounds: CellBounds): CellBounds {
  if (!target) return { ...bounds }
  return {
    top: Math.min(target.top, bounds.top),
    bottom: Math.max(target.bottom, bounds.bottom),
    left: Math.min(target.left, bounds.left),
    right: Math.max(target.right, bounds.right),
  }
}

/** What a sheet can print: its used cells and merges, its print areas and sparklines. */
export function printableBounds(sheet: SheetData): CellBounds {
  let bounds: CellBounds | null = null
  for (const address of Object.keys(sheet.cells || {})) {
    const cell = position(address)
    if (cell) bounds = include(bounds, { top: cell.row, bottom: cell.row, left: cell.col, right: cell.col })
  }
  for (const merge of sheet.merges || []) {
    const range = parseBounds(merge)
    if (range) bounds = include(bounds, range)
  }
  const printArea = sheet.pageSetup?.printArea
  if (typeof printArea === 'string') {
    for (const area of printArea.split(/&&|,/)) {
      const range = parseBounds(area)
      if (range) bounds = include(bounds, range)
    }
  }
  for (const group of sheet.sparklineGroups || []) {
    for (const item of group.sparklines) {
      const range = parseBounds(item.cell)
      if (range) bounds = include(bounds, range)
    }
  }
  return bounds || { top: 0, bottom: 0, left: 0, right: 0 }
}

function within(bounds: CellBounds, row: number, col: number) {
  return row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right
}

/** The print engine's serialized form of a resolved visual; null when it adds nothing. */
export function printCellVisual(visual: CellVisual, iconKey?: string): PrintCellVisual | null {
  const wire: PrintCellVisual = {}
  if (visual.fill) wire.fill = visual.fill
  if (visual.color) wire.color = visual.color
  if (visual.bold !== undefined) wire.bold = visual.bold
  if (visual.italic !== undefined) wire.italic = visual.italic
  if (visual.decoration) wire.decoration = visual.decoration
  if (visual.borders && Object.keys(visual.borders).length) wire.borders = { ...visual.borders }
  if (visual.dataBar) {
    const bar = dataBarBackground(visual.dataBar)
    wire.bar = { image: bar.backgroundImage, size: bar.backgroundSize, position: bar.backgroundPosition, repeat: bar.backgroundRepeat }
  }
  if (iconKey) wire.icon = iconKey
  if (visual.hideValue) wire.hide = true
  if (visual.checkbox) wire.checkbox = visual.checkbox.checked
  return Object.keys(wire).length ? wire : null
}

/** ConditionalIcon as SVG markup (the same glyph the grid draws). */
export function conditionalIconMarkup(set: string, index: number, size = PRINT_ICON_SIZE): string {
  return renderToStaticMarkup(createElement(ConditionalIcon, { set, index, size }))
}

/**
 * Everything a payload depends on besides the options key: each sheet's cells (cell values
 * live there and in formulas over them), its formatting rules and its printed geometry.
 */
function workbookSignature(workbook: WorkbookModel): unknown[] {
  const signature: unknown[] = [workbook.metadata?.themeColors]
  for (const sheet of workbook.sheets) {
    signature.push(
      sheet.id, sheet.name, sheet.state, sheet.cells, sheet.conditionalFormattings, sheet.tables, sheet.dataValidations,
      sheet.sparklineGroups, sheet.merges, sheet.colWidths, sheet.rowHeights, sheet.hiddenRows, sheet.hiddenCols,
      sheet.pageSetup?.printArea, sheet.properties?.printDigitWidth, sheet.properties?.defaultColWidth, sheet.properties?.defaultRowHeight,
    )
  }
  return signature
}

const payloadCache = new WeakMap<object, { signature: unknown[]; key: string; options: PrintVisualOptions; payload: PrintVisualPayload }>()

/**
 * Build the print visuals for a workbook. Pass the printable workbook (the one sent as
 * `workbook`, with printDigitWidth applied) so sparklines are drawn at their printed size.
 * Repeated calls for an unchanged workbook (every preview refresh) reuse the last result;
 * pass stable `cssColor` / `listOptions` functions to benefit.
 */
export function buildPrintVisuals(workbook: WorkbookModel, options: PrintVisualOptions): PrintVisualPayload {
  const today = options.today ?? excelToday()
  const key = JSON.stringify([options.sheetIds ?? null, options.selection ?? null, Boolean(options.showFormulas), today, options.themeColors ?? null])
  const signature = workbookSignature(workbook)
  const cached = payloadCache.get(options.engine)
  if (cached && cached.key === key && cached.options.cssColor === options.cssColor && cached.options.listOptions === options.listOptions &&
    cached.signature.length === signature.length && cached.signature.every((item, index) => item === signature[index])) {
    return cached.payload
  }
  const payload = computePrintVisuals(workbook, options, today)
  payloadCache.set(options.engine, { signature, key, options, payload })
  return payload
}

function computePrintVisuals(workbook: WorkbookModel, options: PrintVisualOptions, today: number): PrintVisualPayload {
  const payload: PrintVisualPayload = { icons: {}, sheets: {} }
  const cssColor = options.cssColor || workbookCssColor(workbook.metadata?.themeColors)
  const themeColors = workbookThemeColors(options.themeColors || workbook.metadata?.themeColors)
  const engine = options.engine
  const valueAt = (sheetId: string, address: string): CellScalar | undefined => engine.getValue(sheetId, address)
  const wanted = options.sheetIds ? new Set(options.sheetIds) : null
  for (const sheet of workbook.sheets) {
    if (wanted ? !wanted.has(sheet.id) : sheet.state === 'hidden' || sheet.state === 'veryHidden') continue
    const bounds = options.selection || printableBounds(sheet)
    let conditional: Map<string, ConditionalCellFormat> | null = null
    if (sheet.conditionalFormattings?.length) {
      try {
        conditional = computeConditionalFormats(sheet.conditionalFormattings, conditionalFormatHost(sheet, engine, { today, cssColor }), bounds)
      } catch {
        conditional = null
      }
    }
    const sparklineFor = createSparklineResolver(workbook, sheet, valueAt)
    // Every cell that can carry a visual: stored cells, conditional results, table cells and
    // sparkline targets, inside what prints (printableBounds already holds every stored cell).
    const candidates = new Set<string>()
    const add = (address: string) => { if (candidates.size < MAX_VISUAL_CELLS) candidates.add(address) }
    for (const address of Object.keys(sheet.cells || {})) {
      if (!options.selection) { add(address); continue }
      const cell = position(address)
      if (cell && within(bounds, cell.row, cell.col)) add(address)
    }
    for (const address of conditional?.keys() || []) add(address)
    for (const entry of parsedTables(sheet.tables)) {
      const top = Math.max(entry.top, bounds.top), bottom = Math.min(entry.bottom, bounds.bottom)
      const left = Math.max(entry.left, bounds.left), right = Math.min(entry.right, bounds.right)
      for (let row = top; row <= bottom; row += 1) for (let col = left; col <= right; col += 1) add(addressOf(row, col))
    }
    for (const group of sheet.sparklineGroups || []) {
      for (const item of group.sparklines) {
        const address = item.cell.replace(/\$/g, '').toUpperCase()
        const cell = position(address)
        if (cell && within(bounds, cell.row, cell.col)) add(address)
      }
    }
    if (!candidates.size) continue

    const merges = new Map<string, CellBounds>()
    for (const merge of sheet.merges || []) {
      const range = parseBounds(merge)
      if (range) merges.set(addressOf(range.top, range.left), range)
    }
    const styles: PrintCellVisual[] = []
    const styleIndex = new Map<string, number>()
    const cells: Record<string, number> = {}
    const text: Record<string, string> = {}
    const sparklines: Record<string, string> = {}
    for (const address of candidates) {
      const cell = position(address)
      if (!cell) continue
      const data = sheet.cells?.[address]
      const raw = valueAt(sheet.id, address)
      const value = raw === undefined ? null : raw
      const entry = sheet.tables?.length ? tableAt(sheet.tables, cell.row, cell.col) : null
      const paint = entry ? tableCellPaint(entry, cell.row, cell.col, themeColors) : null
      let listOptions: readonly string[] | undefined
      const validation = sheet.dataValidations ? findValidation(sheet.dataValidations, cell.row, cell.col)?.validation as Record<string, unknown> | undefined : undefined
      if (validation?.type === 'list' && validation.showDropDown !== true) {
        listOptions = options.listOptions?.(sheet, validation, cell.row, cell.col)
          ?? literalListItems(Array.isArray(validation.formulae) ? validation.formulae[0] : undefined)
          ?? []
      }
      const visual = resolveCellVisual({ cell: data, value, table: paint, conditional: conditional?.get(address), listOptions })
      let iconKey: string | undefined
      if (visual.icon && !data?.hyperlink) {
        const set = normalizeIconSetName(visual.icon.set)
        const index = Math.floor(visual.icon.index)
        if (index >= 0 && index < iconSetSize(set)) {
          iconKey = `${set}:${index}`
          if (!Object.hasOwn(payload.icons, iconKey)) {
            const markup = conditionalIconMarkup(set, index)
            if (markup) payload.icons[iconKey] = markup
            else iconKey = undefined
          }
        }
      }
      const wire = printCellVisual(visual, iconKey)
      if (wire) {
        const key = JSON.stringify(wire)
        let index = styleIndex.get(key)
        if (index === undefined) {
          index = styles.length
          styles.push(wire)
          styleIndex.set(key, index)
        }
        cells[address] = index
      }
      if (visual.numFmt && typeof value === 'number' && !visual.hideValue && !options.showFormulas) text[address] = formatScalar(value, visual.numFmt)
      const spec = formulaSparkline(value) || sparklineFor?.(address) || null
      if (spec) {
        const merge = merges.get(address)
        const size = printedAnchorSize(sheet, { row: cell.row, col: cell.col }, { row: (merge?.bottom ?? cell.row) + 1, col: (merge?.right ?? cell.col) + 1 })
        if (size.width > 0 && size.height > 0) sparklines[address] = renderSparklineSvg(spec, Math.max(4, size.width - 6), Math.max(4, size.height - 4))
      }
    }
    if (!styles.length && !Object.keys(text).length && !Object.keys(sparklines).length) continue
    payload.sheets[sheet.id] = {
      styles,
      cells,
      ...(Object.keys(text).length ? { text } : {}),
      ...(Object.keys(sparklines).length ? { sparklines } : {}),
    }
  }
  return payload
}
