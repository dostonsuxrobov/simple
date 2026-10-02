/**
 * The resolved look of one cell: what table styles (header, banding, borders), the number
 * format's [Color], conditional formatting (fills, colour scales, fonts, borders, data bars,
 * icon sets, number formats), checkboxes and sparklines add on top of the cell's own
 * formatting.
 *
 * The grid and the print / PDF / HTML path share this one resolution, so a printed page
 * shows exactly what the sheet shows. The layering follows Excel: a table style paints
 * under the cell's direct formatting, a format colour replaces the text colour, and a
 * conditional format wins over both.
 *
 * Framework-free and pure: the callers pass the cell, its calculated value and the
 * already-evaluated table paint and conditional format.
 */
import type { CellData, CellScalar, SheetData, SparklineGroup, WorkbookModel } from '../spreadsheet-types'
import type { ConditionalCellFormat, ConditionalDataBar, ConditionalFormatHost, ConditionalScalar } from './conditional-format'
import type { SparklineSpec } from './formula-lib-sparkline'
import type { TableCellPaint } from './table-styles'
import { applyTint } from './cell-styles'
import { columnLabel, defaultConditionalCssColor } from './conditional-format'
import { isSparklineValue, parseSparkline } from './formula-lib-sparkline'
import { formatColor } from './number-format'
import { groupSparklineSpec } from './sparkline-render'
import { parseTableRef } from './tables'

export type VisualBorderSide = 'top' | 'right' | 'bottom' | 'left'

export const VISUAL_BORDER_SIDES: readonly VisualBorderSide[] = ['top', 'right', 'bottom', 'left']

export interface CellVisual {
  /** Background colour over the cell's own fill: a table band or header, a CF fill or a colour scale. */
  fill?: string
  /** A conditional fill replaces any pattern, gradient or diagonal image of the cell's own fill. */
  replacesFillImage?: boolean
  /** Text colour: the table style, then the number format's [Color], then a conditional font colour. */
  color?: string
  /** Bold from a table style (on only) or a conditional format (on or off). */
  bold?: boolean
  /** Italic from a conditional format (on or off). */
  italic?: boolean
  /** CSS text-decoration-line from a conditional format ('underline', 'line-through' or both). */
  decoration?: string
  /** CSS border shorthands ("1px solid #4472C4") from a table style or a conditional format. */
  borders?: Partial<Record<VisualBorderSide, string>>
  dataBar?: ConditionalDataBar
  icon?: { set: string; index: number }
  /** "Show bar only" / "Show icon only": the value text is not drawn. */
  hideValue?: boolean
  /** The number format a conditional format applies (numbers only). */
  numFmt?: string
  /** The cell is drawn as a checkbox. */
  checkbox?: { checked: boolean }
}

export interface CellVisualInput {
  cell?: CellData
  /** The calculated value the grid shows (formula result or constant). */
  value?: CellScalar
  /** tableCellPaint(...) for a cell inside a table. */
  table?: TableCellPaint | null
  conditional?: ConditionalCellFormat | null
  /**
   * Options of a list validation with an in-cell dropdown (validation.showDropDown !== true);
   * exactly TRUE and FALSE make a boolean cell a checkbox, as in Google Sheets.
   */
  listOptions?: readonly string[]
}

const TABLE_BORDER_KEYS: Record<VisualBorderSide, keyof TableCellPaint> = {
  top: 'borderTop',
  right: 'borderRight',
  bottom: 'borderBottom',
  left: 'borderLeft',
}

/** OOXML draws nothing for a side without a style or with style="none". */
function drawableBorder(side: unknown): boolean {
  if (!side || typeof side !== 'object') return false
  const style = String((side as { style?: unknown }).style || '').toLocaleLowerCase()
  return Boolean(style) && style !== 'none'
}

export function checkboxChecked(value: unknown): boolean {
  return value === true || String(value).toLocaleUpperCase() === 'TRUE'
}

/** A checkbox cell control, or a TRUE/FALSE value under a TRUE,FALSE dropdown list. */
export function isCheckboxCell(cell: CellData | undefined, value: unknown, listOptions?: readonly string[]): boolean {
  if (cell?.type === 'checkbox') return true
  return typeof value === 'boolean' && listOptions?.length === 2 &&
    listOptions[0].toLocaleUpperCase() === 'TRUE' && listOptions[1].toLocaleUpperCase() === 'FALSE'
}

/** Resolve everything that paints a cell beyond its own style. Empty object: nothing to add. */
export function resolveCellVisual({ cell, value, table, conditional, listOptions }: CellVisualInput): CellVisual {
  const visual: CellVisual = {}
  const own = cell?.style
  const borders: Partial<Record<VisualBorderSide, string>> = {}
  if (table) {
    // Table styles paint underneath direct cell formatting, as in Excel.
    if (table.fill && !own?.fill?.pattern && !own?.fill?.type) visual.fill = table.fill
    if (table.color && !own?.font?.color) visual.color = table.color
    if (table.bold && own?.font?.bold === undefined) visual.bold = true
    for (const side of VISUAL_BORDER_SIDES) {
      const paint = table[TABLE_BORDER_KEYS[side]]
      if (typeof paint === 'string' && paint && !drawableBorder(own?.border?.[side])) borders[side] = paint
    }
  }
  const formatTextColor = typeof value === 'number' || typeof value === 'string'
    ? formatColor(value, conditional?.numFmt || cell?.numFmt || own?.numFmt)
    : undefined
  if (formatTextColor) visual.color = formatTextColor
  if (conditional) {
    if (conditional.fill || conditional.colorScale) {
      visual.fill = conditional.colorScale || conditional.fill
      visual.replacesFillImage = true
    }
    if (conditional.font?.color) visual.color = conditional.font.color
    if (conditional.font?.bold !== undefined) visual.bold = conditional.font.bold
    if (conditional.font?.italic !== undefined) visual.italic = conditional.font.italic
    if (conditional.font?.underline || conditional.font?.strike) {
      visual.decoration = [conditional.font.underline ? 'underline' : '', conditional.font.strike ? 'line-through' : ''].filter(Boolean).join(' ')
    }
    for (const side of VISUAL_BORDER_SIDES) if (conditional.border?.[side]) borders[side] = conditional.border[side]
    if (conditional.dataBar) visual.dataBar = conditional.dataBar
    if (conditional.icon) visual.icon = { set: conditional.icon.set, index: conditional.icon.index }
    if ((conditional.dataBar && !conditional.dataBar.showValue) || (conditional.icon && !conditional.icon.showValue)) visual.hideValue = true
    if (conditional.numFmt && typeof value === 'number') visual.numFmt = conditional.numFmt
  }
  if (Object.keys(borders).length) visual.borders = borders
  if (isCheckboxCell(cell, value, listOptions)) visual.checkbox = { checked: checkboxChecked(value) }
  return visual
}

export function isEmptyVisual(visual: CellVisual): boolean {
  return Object.keys(visual).length === 0
}

/** Mutable style record the grid paints a cell with (React CSSProperties-compatible keys). */
export type VisualStyleRecord = Record<string, string | number | undefined>

/**
 * Apply a resolved visual to the grid's base cell style (cellCss output), exactly as the
 * grid layers table paint, format colour and conditional formatting. `dataBarCss` is
 * conditional-format's dataBarBackground (passed in to keep this module CSS-free).
 */
export function applyCellVisualStyle(
  style: VisualStyleRecord,
  visual: CellVisual,
  dataBarCss: (bar: ConditionalDataBar) => { backgroundImage: string; backgroundSize: string; backgroundPosition: string; backgroundRepeat: string },
): VisualStyleRecord {
  if (visual.fill) style.backgroundColor = visual.fill
  if (visual.replacesFillImage) style.backgroundImage = undefined
  if (visual.color) style.color = visual.color
  if (visual.bold !== undefined) style.fontWeight = visual.bold ? 700 : 400
  if (visual.italic !== undefined) style.fontStyle = visual.italic ? 'italic' : 'normal'
  if (visual.decoration) style.textDecorationLine = visual.decoration
  if (visual.borders?.top) style.borderTop = visual.borders.top
  if (visual.borders?.right) style.borderRight = visual.borders.right
  if (visual.borders?.bottom) style.borderBottom = visual.borders.bottom
  if (visual.borders?.left) style.borderLeft = visual.borders.left
  if (visual.dataBar) {
    const bar = dataBarCss(visual.dataBar)
    style.backgroundImage = bar.backgroundImage
    style.backgroundSize = bar.backgroundSize
    style.backgroundPosition = bar.backgroundPosition
    style.backgroundRepeat = bar.backgroundRepeat
  }
  return style
}

// ---------------------------------------------------------------------------------------------
// Inputs the grid and print both need
// ---------------------------------------------------------------------------------------------

/** The calculation engine surface the visuals need (CalculationEngine satisfies it). */
export interface VisualValueEngine {
  getValue(sheetId: string, address: string): CellScalar | undefined
  evaluateAt(sheetId: string, formula: string, row?: number, col?: number): CellScalar | undefined
}

const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']

/** A workbook's theme colours (OOXML order) as 6-digit hex, the Office default filling gaps. */
export function workbookThemeColors(themeColors?: readonly unknown[]): string[] {
  return OFFICE_THEME.map((fallback, index) => {
    const value = String(themeColors?.[index] ?? '').replace(/^#/, '')
    return /^[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : fallback
  })
}

/**
 * A colour resolver for a workbook's own theme (OOXML theme-index order), with Excel's HSL
 * tint; plain argb/rgb strings and `{ argb }` objects pass through.
 */
export function workbookCssColor(themeColors?: readonly unknown[]): (color: unknown, fallback?: string) => string {
  const theme = workbookThemeColors(themeColors)
  return (color, fallback = '') => {
    if (color && typeof color === 'object') {
      const record = color as { argb?: string; rgb?: string; theme?: number; tint?: number }
      if (!record.argb && !record.rgb && Number.isInteger(record.theme) && theme[Number(record.theme)]) {
        const tint = Number(record.tint) || 0
        return `#${(tint ? applyTint(theme[Number(record.theme)], Math.max(-1, Math.min(1, tint))) : theme[Number(record.theme)]).replace(/^#/, '').toUpperCase()}`
      }
    }
    return defaultConditionalCssColor(color, fallback)
  }
}

/** Excel serial for today (1900 date system), for timePeriod rules. */
export function excelToday(now = new Date()): number {
  return (Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - Date.UTC(1899, 11, 30)) / 86_400_000
}

/** The conditional-format host for one sheet, as the grid builds it. */
export function conditionalFormatHost(
  sheet: SheetData,
  engine: VisualValueEngine,
  options: { today?: number; cssColor?: (color: unknown, fallback?: string) => string } = {},
): ConditionalFormatHost {
  return {
    sheetId: sheet.id,
    valueAt: (row, col) => {
      const value = engine.getValue(sheet.id, `${columnLabel(col)}${row + 1}`)
      return value === undefined ? null : value as ConditionalScalar
    },
    evaluate: (formula, row, col) => {
      const value = engine.evaluateAt(sheet.id, formula, row, col)
      return value === null || value === undefined ? 0 : value
    },
    today: options.today ?? excelToday(),
    cssColor: options.cssColor,
    rowCount: sheet.rowCount,
    colCount: sheet.colCount,
  }
}

const SPARKLINE_SOURCE = /^(?:(?:'((?:[^']|'')+)'|([^!'"]+))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/

/** The numbers of an Excel sparkline's source range (blanks and text are gaps). */
export function readSparklineSource(
  workbook: WorkbookModel,
  defaultSheet: SheetData,
  source: string,
  valueAt: (sheetId: string, address: string) => CellScalar | undefined,
): Array<number | null> {
  const match = SPARKLINE_SOURCE.exec(source.trim())
  if (!match) return []
  const name = match[1] ? match[1].replace(/''/g, "'") : match[2]
  const sheet = name ? workbook.sheets.find((item) => item.name.toLocaleLowerCase() === name.trim().toLocaleLowerCase()) : defaultSheet
  const bounds = parseTableRef(match[3])
  if (!sheet || !bounds) return []
  const values: Array<number | null> = []
  for (let row = bounds.top; row <= bounds.bottom && values.length < 2000; row += 1) {
    for (let col = bounds.left; col <= bounds.right && values.length < 2000; col += 1) {
      const value = valueAt(sheet.id, `${columnLabel(col)}${row + 1}`)
      values.push(typeof value === 'number' && Number.isFinite(value) ? value : null)
    }
  }
  return values
}

/**
 * The sparkline drawn in each target cell of a sheet's Excel sparkline groups, with the
 * group/custom axis shared across the group; null when the sheet has none.
 */
export function createSparklineResolver(
  workbook: WorkbookModel,
  sheet: SheetData,
  valueAt: (sheetId: string, address: string) => CellScalar | undefined,
): ((address: string) => SparklineSpec | null) | null {
  const groups = sheet.sparklineGroups || []
  if (!groups.length) return null
  const index = new Map<string, { group: SparklineGroup; groupIndex: number; source: string }>()
  groups.forEach((group, groupIndex) => {
    for (const item of group.sparklines) index.set(item.cell.replace(/\$/g, '').toUpperCase(), { group, groupIndex, source: item.source })
  })
  const read = (source: string) => readSparklineSource(workbook, sheet, source, valueAt)
  const groupAxis = new Map<number, { min?: number; max?: number }>()
  const axisFor = (group: SparklineGroup, groupIndex: number) => {
    if (group.minAxisType !== 'group' && group.maxAxisType !== 'group' && group.minAxisType !== 'custom' && group.maxAxisType !== 'custom') return undefined
    let cached = groupAxis.get(groupIndex)
    if (!cached) {
      const all = group.sparklines.flatMap((item) => read(item.source)).filter((value): value is number => value !== null)
      cached = {
        ...(group.minAxisType === 'custom' ? { min: group.manualMin } : group.minAxisType === 'group' && all.length ? { min: Math.min(...all) } : {}),
        ...(group.maxAxisType === 'custom' ? { max: group.manualMax } : group.maxAxisType === 'group' && all.length ? { max: Math.max(...all) } : {}),
      }
      groupAxis.set(groupIndex, cached)
    }
    return cached
  }
  return (address) => {
    const entry = index.get(address)
    if (!entry) return null
    return groupSparklineSpec(entry.group, read(entry.source), axisFor(entry.group, entry.groupIndex))
  }
}

/** The in-cell chart of a SPARKLINE() result, if the value is one. */
export function formulaSparkline(value: unknown): SparklineSpec | null {
  return isSparklineValue(value) ? parseSparkline(value) : null
}
