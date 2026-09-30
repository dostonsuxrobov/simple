// Pure chart model helpers: range parsing, data resolution, auto-detection of
// series from a selected range, anchor geometry and structural-edit shifting.
// No DOM or React dependencies (also bundled into Node QA scripts).
import type {
  CellScalar,
  ChartAnchor,
  ChartAnchorPoint,
  ChartGrouping,
  ChartSeries,
  ChartSeriesType,
  ChartType,
  SheetChart,
  SheetData,
  WorkbookModel,
} from '../spreadsheet-types'
import type { SheetStructureOperation } from './sheet-operations'

export const EMU_PER_PIXEL = 9525
const MAX_ROWS = 1_048_576
const MAX_COLS = 16_384
/** Hard cap on points read per reference so a whole-column reference cannot stall rendering. */
export const MAX_CHART_POINTS = 20_000

// ---------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------

/** Office 2013+ theme accents 1-6. */
export const OFFICE_ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

export const CHART_PALETTES: Array<{ id: string; label: string; colors: string[] }> = [
  { id: 'office', label: 'Office', colors: OFFICE_ACCENTS },
  { id: 'sheets', label: 'Bright', colors: ['#4285F4', '#EA4335', '#FBBC04', '#34A853', '#FF6D01', '#46BDC6'] },
  { id: 'simple', label: 'Simple', colors: ['#476B57', '#8FB3A0', '#C9A227', '#5C7FA3', '#B5654A', '#7E6BA8'] },
  { id: 'ocean', label: 'Ocean', colors: ['#1F4E79', '#2E75B6', '#5B9BD5', '#9DC3E6', '#00B0F0', '#0070C0'] },
  { id: 'warm', label: 'Warm', colors: ['#C00000', '#ED7D31', '#FFC000', '#BF8F00', '#843C0C', '#F4B183'] },
  { id: 'mono', label: 'Grey', colors: ['#404040', '#7F7F7F', '#A6A6A6', '#262626', '#595959', '#BFBFBF'] },
]

export function normalizeHex(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  const match = /^#?([0-9a-f]{6})(?:[0-9a-f]{2})?$/i.exec(value.trim())
  if (match) return `#${match[1].toUpperCase()}`
  const short = /^#?([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(value.trim())
  return short ? `#${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`.toUpperCase() : undefined
}

function hexToRgb(hex: string) {
  const value = parseInt(hex.slice(1), 16)
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255]
}

function rgbToHex(r: number, g: number, b: number) {
  const part = (v: number) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')
  return `#${part(r)}${part(g)}${part(b)}`.toUpperCase()
}

function rgbToHsl(r: number, g: number, b: number) {
  r /= 255; g /= 255; b /= 255
  const max = Math.max(r, g, b), min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
  return [h / 6, s, l]
}

function hslToRgb(h: number, s: number, l: number) {
  if (s === 0) return [l * 255, l * 255, l * 255]
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255]
}

/** DrawingML lumMod/lumOff (fractions, e.g. 0.6 / 0.4). */
export function adjustLuminance(hex: string, lumMod = 1, lumOff = 0) {
  const normalized = normalizeHex(hex) || '#000000'
  const [h, s, l] = rgbToHsl(...(hexToRgb(normalized) as [number, number, number]))
  const [r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, l * lumMod + lumOff)))
  return rgbToHex(r, g, b)
}

// Excel's variations for series 7+: darker, lighter, and so on.
const PALETTE_VARIANTS: Array<[number, number]> = [[1, 0], [0.6, 0], [0.6, 0.4], [0.8, 0], [0.8, 0.2], [0.5, 0], [0.4, 0.6]]

/** Colour for the n-th series (0-based) following Excel's accent cycling. */
export function paletteColor(index: number, palette: string[] = OFFICE_ACCENTS) {
  const colors = palette.length ? palette : OFFICE_ACCENTS
  const base = colors[((index % colors.length) + colors.length) % colors.length]
  const cycle = Math.floor(Math.max(0, index) / colors.length) % PALETTE_VARIANTS.length
  const [mod, off] = PALETTE_VARIANTS[cycle]
  return cycle === 0 ? normalizeHex(base) || '#4472C4' : adjustLuminance(base, mod, off)
}

/** Series palette for a workbook: its own theme accents (theme slots 4-9) when present. */
export function workbookChartPalette(workbook?: Pick<WorkbookModel, 'metadata'> | null): string[] {
  const theme = workbook?.metadata?.themeColors
  if (Array.isArray(theme) && theme.length >= 10) {
    const accents = theme.slice(4, 10).map((value) => normalizeHex(value))
    if (accents.every(Boolean)) return accents as string[]
  }
  return OFFICE_ACCENTS
}

export function chartPalette(chart: Pick<SheetChart, 'style'>, fallback?: string[]) {
  const custom = chart.style?.palette?.map((color) => normalizeHex(color)).filter(Boolean) as string[] | undefined
  return custom?.length ? custom : fallback?.length ? fallback : OFFICE_ACCENTS
}

// ---------------------------------------------------------------------------
// A1 references
// ---------------------------------------------------------------------------

export interface CellBounds { top: number; bottom: number; left: number; right: number }
export interface ChartRefArea extends CellBounds { sheet?: string }

export function columnLabel(index: number) {
  let value = Math.max(0, Math.trunc(index)) + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

export function columnIndex(label: string) {
  let value = 0
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value - 1
}

export function cellAddress(row: number, col: number) {
  return `${columnLabel(col)}${row + 1}`
}

export function quoteSheetName(name: string) {
  return /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(true|false)$/i.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`
}

/** "Sheet1!$A$1:$B$4" (absolute) or "A1:B4" when `sheet` is empty. */
export function formatChartRef(sheet: string | undefined, bounds: CellBounds, absolute = true) {
  const d = absolute ? '$' : ''
  const start = `${d}${columnLabel(bounds.left)}${d}${bounds.top + 1}`
  const end = `${d}${columnLabel(bounds.right)}${d}${bounds.bottom + 1}`
  const body = bounds.top === bounds.bottom && bounds.left === bounds.right ? start : `${start}:${end}`
  return sheet ? `${quoteSheetName(sheet)}!${body}` : body
}

const CELL_PART = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/
const COLUMN_PART = /^\$?([A-Za-z]{1,3})$/
const ROW_PART = /^\$?(\d{1,7})$/

function parseAreaBody(body: string): CellBounds | null {
  const [first, second = first, extra] = body.split(':')
  if (extra !== undefined) return null
  const a = CELL_PART.exec(first), b = CELL_PART.exec(second)
  if (a && b) {
    const r1 = Number(a[2]) - 1, r2 = Number(b[2]) - 1, c1 = columnIndex(a[1]), c2 = columnIndex(b[1])
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS || c1 >= MAX_COLS || c2 >= MAX_COLS) return null
    return { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: Math.min(c1, c2), right: Math.max(c1, c2) }
  }
  const ca = COLUMN_PART.exec(first), cb = COLUMN_PART.exec(second)
  if (ca && cb && body.includes(':')) {
    const c1 = columnIndex(ca[1]), c2 = columnIndex(cb[1])
    return { top: 0, bottom: MAX_ROWS - 1, left: Math.min(c1, c2), right: Math.max(c1, c2) }
  }
  const ra = ROW_PART.exec(first), rb = ROW_PART.exec(second)
  if (ra && rb && body.includes(':')) {
    const r1 = Number(ra[1]) - 1, r2 = Number(rb[1]) - 1
    if (r1 < 0 || r2 < 0) return null
    return { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: 0, right: MAX_COLS - 1 }
  }
  return null
}

function splitTopLevel(text: string, separator: string) {
  const parts: string[] = []
  let current = '', quoted = false, depth = 0
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === "'") {
      if (quoted && text[index + 1] === "'") { current += "''"; index += 1; continue }
      quoted = !quoted
    } else if (!quoted && character === '(') depth += 1
    else if (!quoted && character === ')') depth -= 1
    if (!quoted && depth === 0 && character === separator) { parts.push(current); current = '' } else current += character
  }
  parts.push(current)
  return parts
}

/**
 * Parse a chart reference: "Sheet1!$B$2:$B$10", "'My Sheet'!A1", "(Sheet1!$A$1:$A$3,Sheet1!$C$1:$C$3)",
 * "B2:B10" (relative to `defaultSheet`). Returns null for #REF!, external or malformed references.
 */
export function parseChartRef(ref: string | undefined | null, defaultSheet?: string): ChartRefArea[] | null {
  if (typeof ref !== 'string') return null
  let text = ref.trim()
  if (!text || /#REF!/i.test(text) || text.length > 4096) return null
  if (text.startsWith('=')) text = text.slice(1).trim()
  while (text.startsWith('(') && text.endsWith(')')) text = text.slice(1, -1).trim()
  const pieces = splitTopLevel(text, ',')
  const areas: ChartRefArea[] = []
  let lastSheet = defaultSheet
  for (const rawPiece of pieces) {
    const piece = rawPiece.trim()
    if (!piece) return null
    let sheet = lastSheet
    let body = piece
    const bang = piece.lastIndexOf('!')
    if (bang >= 0) {
      let sheetText = piece.slice(0, bang).trim()
      body = piece.slice(bang + 1).trim()
      if (sheetText.startsWith("'") && sheetText.endsWith("'")) sheetText = sheetText.slice(1, -1).replace(/''/g, "'")
      if (/^\[\d+\]/.test(sheetText) || sheetText.includes('[')) return null // external workbook
      sheet = sheetText
      lastSheet = sheet
    }
    const bounds = parseAreaBody(body.replace(/\s+/g, ''))
    if (!bounds) return null
    areas.push({ ...bounds, sheet })
  }
  return areas.length ? areas : null
}

export function refCellCount(areas: ChartRefArea[]) {
  return areas.reduce((sum, area) => sum + (area.bottom - area.top + 1) * (area.right - area.left + 1), 0)
}

// ---------------------------------------------------------------------------
// Workbook accessor
// ---------------------------------------------------------------------------

/** Read-only view of calculated cell values used by charts. Addresses are A1 ("B7"). */
export interface ChartWorkbookAccessor {
  /** Calculated value (formula result or literal). */
  value(sheetName: string, address: string): CellScalar | undefined
  /** Formatted display text; defaults to String(value). */
  text?(sheetName: string, address: string): string
  numberFormat?(sheetName: string, address: string): string | undefined
  /** 1-based row / column hidden state (for "plot visible cells only"). */
  isRowHidden?(sheetName: string, row: number): boolean
  isColumnHidden?(sheetName: string, col: number): boolean
  /** Resolve a defined name to a reference such as "Sheet1!$A$1:$A$9". */
  resolveName?(name: string, sheetName?: string): string | undefined
  /** False for deleted/unknown sheets, so their references fall back to the cached points. */
  hasSheet?(sheetName: string): boolean
}

function scalarOf(value: unknown): CellScalar | undefined {
  if (value === null || value === undefined) return undefined
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') return value
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>
    if (record.type === 'date' && record.value !== undefined) {
      const date = new Date(String(record.value))
      if (!Number.isNaN(date.getTime())) return date.getTime() / 86_400_000 + 25_569
    }
    if (typeof record.value === 'string' || typeof record.value === 'number') return record.value
  }
  return undefined
}

/**
 * Build an accessor straight from the workbook model (cached formula results and literals).
 * `calculated` can supply live engine results (sheetId, A1) => value and `display` formatted text.
 */
export function workbookChartAccessor(
  workbook: WorkbookModel,
  calculated?: (sheetId: string, address: string) => CellScalar | undefined,
  display?: (sheetId: string, address: string) => string | undefined,
): ChartWorkbookAccessor {
  const byName = new Map<string, SheetData>()
  for (const sheet of workbook.sheets) byName.set(sheet.name.toLocaleLowerCase(), sheet)
  const hiddenRows = new WeakMap<SheetData, Set<number>>()
  const hiddenCols = new WeakMap<SheetData, Set<number>>()
  const find = (name: string) => byName.get(String(name || '').toLocaleLowerCase())
  const valueOf = (sheet: SheetData, address: string) => {
    if (calculated) {
      const live = calculated(sheet.id, address)
      if (live !== undefined) return live
    }
    const cell = sheet.cells?.[address]
    if (!cell) return undefined
    return scalarOf(cell.formula ? cell.result ?? cell.value : cell.value)
  }
  return {
    hasSheet: (sheetName) => Boolean(find(sheetName)),
    value(sheetName, address) {
      const sheet = find(sheetName)
      return sheet ? valueOf(sheet, address) : undefined
    },
    text(sheetName, address) {
      const sheet = find(sheetName)
      if (!sheet) return ''
      const shown = display?.(sheet.id, address) ?? sheet.cells?.[address]?.display
      if (typeof shown === 'string') return shown
      const value = valueOf(sheet, address)
      return value === undefined || value === null ? '' : String(value)
    },
    numberFormat(sheetName, address) {
      const cell = find(sheetName)?.cells?.[address]
      return cell?.numFmt || cell?.style?.numFmt || undefined
    },
    isRowHidden(sheetName, row) {
      const sheet = find(sheetName)
      if (!sheet?.hiddenRows?.length) return false
      let set = hiddenRows.get(sheet)
      if (!set) hiddenRows.set(sheet, set = new Set(sheet.hiddenRows))
      return set.has(row)
    },
    isColumnHidden(sheetName, col) {
      const sheet = find(sheetName)
      if (!sheet?.hiddenCols?.length) return false
      let set = hiddenCols.get(sheet)
      if (!set) hiddenCols.set(sheet, set = new Set(sheet.hiddenCols))
      return set.has(col)
    },
    resolveName(name, sheetName) {
      const lower = name.toLocaleLowerCase()
      const sheetIndex = sheetName ? workbook.sheets.findIndex((sheet) => sheet.name.toLocaleLowerCase() === String(sheetName).toLocaleLowerCase()) : -1
      const names = workbook.definedNames || []
      const match = names.find((item) => item.name.toLocaleLowerCase() === lower && item.localSheetIndex === sheetIndex)
        || names.find((item) => item.name.toLocaleLowerCase() === lower && item.localSheetIndex === undefined)
      return match?.ref || match?.ranges?.[0]
    },
  }
}

interface RefRead {
  values: Array<CellScalar | undefined>
  texts: string[]
  numFmt?: string
  /** Rows (for column vectors) or columns of the first area, used for multi-level labels. */
  levels: number
}

function readRef(ref: string | undefined, accessor: ChartWorkbookAccessor, defaultSheet: string | undefined, visibleOnly: boolean): RefRead | null {
  if (!ref) return null
  let areas = parseChartRef(ref, defaultSheet)
  if (!areas && accessor.resolveName) {
    const bare = ref.trim().replace(/^=/, '')
    const bang = bare.lastIndexOf('!')
    const name = bang >= 0 ? bare.slice(bang + 1) : bare
    const scope = bang >= 0 ? bare.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") : defaultSheet
    if (/^[A-Za-z_\\][\w.\\]*$/.test(name)) {
      const resolved = accessor.resolveName(name, scope)
      if (resolved && resolved !== ref) areas = parseChartRef(resolved, scope)
    }
  }
  if (!areas) return null
  if (accessor.hasSheet && areas.some((area) => !accessor.hasSheet!(area.sheet || defaultSheet || ''))) return null
  const values: Array<CellScalar | undefined> = []
  const texts: string[] = []
  let numFmt: string | undefined
  let levels = 1
  for (const [areaIndex, area] of areas.entries()) {
    const sheet = area.sheet || defaultSheet
    if (!sheet) return null
    const rows = area.bottom - area.top + 1
    const cols = area.right - area.left + 1
    // A 2-D category range is multi-level: the innermost level (last row/column) labels points.
    const byColumn = rows >= cols
    if (areaIndex === 0) levels = byColumn ? cols : rows
    const outer = byColumn ? rows : cols
    for (let i = 0; i < outer && values.length < MAX_CHART_POINTS; i += 1) {
      const row = byColumn ? area.top + i : area.bottom
      const col = byColumn ? area.right : area.left + i
      if (visibleOnly && (accessor.isRowHidden?.(sheet, row + 1) || accessor.isColumnHidden?.(sheet, col + 1))) continue
      const address = cellAddress(row, col)
      const value = accessor.value(sheet, address)
      values.push(value)
      texts.push(accessor.text ? accessor.text(sheet, address) : value === undefined || value === null ? '' : String(value))
      if (numFmt === undefined && typeof value === 'number') numFmt = accessor.numberFormat?.(sheet, address)
    }
    // Whole-column/row references ("B:B") stop at the last non-empty cell.
    if (area.bottom === MAX_ROWS - 1 || area.right === MAX_COLS - 1) {
      while (values.length && (values[values.length - 1] === undefined || values[values.length - 1] === null || values[values.length - 1] === '')) { values.pop(); texts.pop() }
    }
  }
  return { values, texts, numFmt, levels }
}

function numericOf(value: CellScalar | undefined): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

// ---------------------------------------------------------------------------
// Resolved data
// ---------------------------------------------------------------------------

export interface ResolvedChartSeries {
  name: string
  values: Array<number | null>
  x?: Array<number | null>
  color: string
  /** Colours per point when points vary (pie/doughnut/varyColors). */
  pointColors?: string[]
  numFmt?: string
}

export interface ResolvedChartData {
  categories: string[]
  /** Numeric category values when every category is a number (scatter X / date axes). */
  categoryNumbers?: Array<number | null>
  categoryNumFmt?: string
  series: ResolvedChartSeries[]
}

export function chartSeriesType(chart: SheetChart, series: ChartSeries | undefined): ChartSeriesType | 'pie' | 'doughnut' | 'radar' {
  if (chart.type === 'combo') return series?.type || 'column'
  if (chart.type === 'pie' || chart.type === 'doughnut' || chart.type === 'radar') return chart.type
  if (chart.type === 'unsupported') return 'column'
  return chart.type
}

export function chartVariesColors(chart: SheetChart) {
  return chart.type === 'pie' || chart.type === 'doughnut' || (chart.varyColors === true && chart.series.length === 1 && chart.type !== 'combo' && chart.type !== 'scatter')
}

/**
 * Resolve the chart's references against calculated values. Unresolvable
 * references fall back to the imported caches, so charts over external or
 * deleted data keep their last known shape.
 */
export function resolveChartData(
  chart: SheetChart,
  accessor: ChartWorkbookAccessor | null | undefined,
  options: { palette?: string[]; defaultSheet?: string } = {},
): ResolvedChartData {
  const palette = chartPalette(chart, options.palette)
  const visibleOnly = chart.plotVisibleOnly !== false
  const series = Array.isArray(chart.series) ? chart.series : []
  let categories: string[] | undefined
  let categoryNumbers: Array<number | null> | undefined
  let categoryNumFmt: string | undefined
  const resolved: ResolvedChartSeries[] = series.map((item, index) => {
    const valuesRead = accessor ? readRef(item.valuesRef, accessor, options.defaultSheet, visibleOnly) : null
    const values = valuesRead ? valuesRead.values.map(numericOf) : (item.valuesCache || []).map((value) => (typeof value === 'number' && Number.isFinite(value) ? value : null))
    let name = item.name || ''
    if (item.nameRef && accessor) {
      const nameRead = readRef(item.nameRef, accessor, options.defaultSheet, false)
      if (nameRead) name = nameRead.texts.filter(Boolean).join(' ') || name
    }
    if (!name) name = `Series${index + 1}`
    let x: Array<number | null> | undefined
    if (chartSeriesType(chart, item) === 'scatter') {
      const xRead = accessor ? readRef(item.xValuesRef || item.categoriesRef, accessor, options.defaultSheet, visibleOnly) : null
      const raw = xRead ? xRead.values.map(numericOf) : item.xValuesCache
      x = raw && raw.some((value) => value !== null) ? raw.slice(0, values.length) : values.map((_, point) => point + 1)
    }
    if (!categories) {
      const catRead = accessor ? readRef(item.categoriesRef, accessor, options.defaultSheet, visibleOnly) : null
      if (catRead) {
        categories = catRead.texts.map((text, point) => text || (catRead.values[point] === undefined ? '' : String(catRead.values[point])))
        const numbers = catRead.values.map(numericOf)
        if (numbers.length && numbers.every((value, point) => value !== null || catRead.values[point] === undefined)) categoryNumbers = numbers
        categoryNumFmt = catRead.numFmt
      } else if (item.categoriesCache?.length) {
        categories = item.categoriesCache.map((value) => String(value ?? ''))
      }
    }
    const color = normalizeHex(item.color) || paletteColor(index, palette)
    const output: ResolvedChartSeries = { name, values, color }
    if (x) output.x = x
    const numFmt = item.valuesNumFmt || valuesRead?.numFmt
    if (numFmt) output.numFmt = numFmt
    return output
  })
  const pointCount = resolved.reduce((max, item) => Math.max(max, item.values.length), 0)
  if (!categories) categories = Array.from({ length: pointCount }, (_, index) => String(index + 1))
  else if (categories.length < pointCount) categories = [...categories, ...Array.from({ length: pointCount - categories.length }, (_, index) => String(categories!.length + index + 1))]
  if (chartVariesColors(chart)) {
    resolved.forEach((item, seriesIndex) => {
      const overrides = series[seriesIndex]?.pointColors || {}
      item.pointColors = item.values.map((_, point) => normalizeHex(overrides[String(point)]) || paletteColor(point, palette))
    })
  } else {
    resolved.forEach((item, seriesIndex) => {
      const overrides = series[seriesIndex]?.pointColors
      if (overrides && Object.keys(overrides).length) item.pointColors = item.values.map((_, point) => normalizeHex(overrides[String(point)]) || item.color)
    })
  }
  const data: ResolvedChartData = { categories, series: resolved }
  if (categoryNumbers) data.categoryNumbers = categoryNumbers
  if (categoryNumFmt) data.categoryNumFmt = categoryNumFmt
  return data
}

/** Title shown on the chart, including Excel's automatic single-series title. */
export function chartDisplayTitle(chart: SheetChart, data?: ResolvedChartData | null) {
  if (typeof chart.title === 'string' && chart.title.trim()) return chart.title
  if (chart.autoTitleDeleted) return ''
  if (chart.title === '') return ''
  if (chart.series.length === 1 && chart.type !== 'unsupported') return data?.series[0]?.name || chart.series[0].name || ''
  return ''
}

// ---------------------------------------------------------------------------
// Creating charts from a selection
// ---------------------------------------------------------------------------

let idCounter = 0
export function newChartId(prefix = 'chart') {
  idCounter = (idCounter + 1) % 1_000_000
  const random = Math.random().toString(36).slice(2, 8)
  return `${prefix}-${Date.now().toString(36)}-${idCounter.toString(36)}${random}`
}

interface CellProbe { value: CellScalar | undefined; text: string }

function probe(accessor: ChartWorkbookAccessor, sheet: string, row: number, col: number): CellProbe {
  const address = cellAddress(row, col)
  const value = accessor.value(sheet, address)
  return { value, text: accessor.text ? accessor.text(sheet, address) : value == null ? '' : String(value) }
}

const isBlank = (cell: CellProbe) => cell.value === undefined || cell.value === null || cell.value === ''
const isNumber = (cell: CellProbe) => typeof cell.value === 'number' && Number.isFinite(cell.value)

/** Expand a single selected cell to its contiguous data block (Excel's current region). */
export function expandToCurrentRegion(sheetName: string, bounds: CellBounds, accessor: ChartWorkbookAccessor, limit = 2_000): CellBounds {
  if (bounds.top !== bounds.bottom || bounds.left !== bounds.right) return bounds
  const region = { ...bounds }
  const filled = (row: number, col: number) => row >= 0 && col >= 0 && !isBlank(probe(accessor, sheetName, row, col))
  let grew = true
  let guard = 0
  while (grew && guard++ < 64) {
    grew = false
    const rowHas = (row: number) => { for (let col = Math.max(0, region.left - 1); col <= region.right + 1; col += 1) if (filled(row, col)) return true; return false }
    const colHas = (col: number) => { for (let row = Math.max(0, region.top - 1); row <= region.bottom + 1; row += 1) if (filled(row, col)) return true; return false }
    while (region.top > 0 && region.bottom - region.top < limit && rowHas(region.top - 1)) { region.top -= 1; grew = true }
    while (region.bottom - region.top < limit && rowHas(region.bottom + 1)) { region.bottom += 1; grew = true }
    while (region.left > 0 && region.right - region.left < 200 && colHas(region.left - 1)) { region.left -= 1; grew = true }
    while (region.right - region.left < 200 && colHas(region.right + 1)) { region.right += 1; grew = true }
  }
  return region
}

export interface RangeLayout {
  seriesIn: 'columns' | 'rows'
  firstRowHeaders: boolean
  firstColumnLabels: boolean
}

/** Excel/Sheets-like detection of headers, labels and series orientation. */
export function detectRangeLayout(sheetName: string, bounds: CellBounds, accessor: ChartWorkbookAccessor, type?: ChartType): RangeLayout {
  const rows = bounds.bottom - bounds.top + 1
  const cols = bounds.right - bounds.left + 1
  const sampleRows = Math.min(rows, 200)
  const sampleCols = Math.min(cols, 60)
  const at = (r: number, c: number) => probe(accessor, sheetName, bounds.top + r, bounds.left + c)
  const corner = at(0, 0)
  // Header row: the first row holds text/blank while later rows hold numbers.
  let headerText = 0, headerNumbers = 0
  for (let c = 0; c < sampleCols; c += 1) {
    const cell = at(0, c)
    if (isNumber(cell)) headerNumbers += 1
    else if (!isBlank(cell)) headerText += 1
  }
  let bodyNumbers = 0
  for (let r = 1; r < sampleRows; r += 1) for (let c = 0; c < sampleCols; c += 1) if (isNumber(at(r, c))) bodyNumbers += 1
  const firstRowHeaders = rows > 1 && headerText > 0 && (headerNumbers === 0 || headerNumbers < headerText) && bodyNumbers > 0

  // Label column: text/dates in the first column, or a blank corner above labels.
  const start = firstRowHeaders ? 1 : 0
  let labelText = 0, labelNumbers = 0, labelBlank = 0
  const numbers: number[] = []
  for (let r = start; r < sampleRows; r += 1) {
    const cell = at(r, 0)
    if (isNumber(cell)) { labelNumbers += 1; numbers.push(cell.value as number) } else if (isBlank(cell)) labelBlank += 1
    else labelText += 1
  }
  const dateLike = (() => {
    const fmt = accessor.numberFormat?.(sheetName, cellAddress(bounds.top + start, bounds.left)) || ''
    return /(^|[^\\])[dmy]{1,4}/i.test(fmt.replace(/"[^"]*"/g, '')) && !/^[#0.,%\s]*$/.test(fmt)
  })()
  const yearLike = numbers.length >= 2 && numbers.every((value, index) => Number.isInteger(value) && value >= 1800 && value <= 2200 && (index === 0 || value === numbers[index - 1] + 1))
  const cornerBlank = isBlank(corner)
  const firstColumnLabels = cols > 1 && (
    (labelText > 0 && labelText >= labelNumbers) ||
    (firstRowHeaders && cornerBlank && labelBlank < sampleRows - start) ||
    (labelNumbers > 0 && (dateLike || yearLike)) ||
    (type === 'scatter' && labelNumbers > 0)
  )
  const dataRows = rows - (firstRowHeaders ? 1 : 0)
  const dataCols = cols - (firstColumnLabels ? 1 : 0)
  // Excel: series follow the smaller dimension; ties plot one series per column.
  const seriesIn: 'columns' | 'rows' = type === 'scatter' ? 'columns' : dataRows >= dataCols ? 'columns' : 'rows'
  return { seriesIn, firstRowHeaders, firstColumnLabels }
}

/** Build series references for a range and layout. */
export function buildSeriesFromRange(sheetName: string, bounds: CellBounds, accessor: ChartWorkbookAccessor | null, layout: RangeLayout, type: ChartType): ChartSeries[] {
  const { seriesIn, firstRowHeaders, firstColumnLabels } = layout
  const top = bounds.top + (firstRowHeaders ? 1 : 0)
  const left = bounds.left + (firstColumnLabels ? 1 : 0)
  const series: ChartSeries[] = []
  const nameOf = (row: number, col: number) => (accessor ? probe(accessor, sheetName, row, col).text : '')
  if (seriesIn === 'columns') {
    if (top > bounds.bottom) return []
    // Excel: every leading text-only column is a category level (multi-level axis), and a
    // column without numbers is never plotted as a series.
    const hasNumbers = (col: number) => {
      if (!accessor) return true
      for (let row = top; row <= Math.min(bounds.bottom, top + 199); row += 1) {
        const cell = probe(accessor, sheetName, row, col)
        if (isNumber(cell)) return true
      }
      return false
    }
    let labelRight = firstColumnLabels ? bounds.left : bounds.left - 1
    if (firstColumnLabels && type !== 'scatter') {
      while (labelRight + 1 < bounds.right && !hasNumbers(labelRight + 1)) labelRight += 1
    }
    const categoriesRef = firstColumnLabels ? formatChartRef(sheetName, { top, bottom: bounds.bottom, left: bounds.left, right: labelRight }) : undefined
    for (let col = Math.max(left, labelRight + 1); col <= bounds.right && series.length < 255; col += 1) {
      if (!hasNumbers(col)) continue
      const item: ChartSeries = { id: newChartId('series'), valuesRef: formatChartRef(sheetName, { top, bottom: bounds.bottom, left: col, right: col }) }
      if (firstRowHeaders) { item.nameRef = formatChartRef(sheetName, { top: bounds.top, bottom: bounds.top, left: col, right: col }); item.name = nameOf(bounds.top, col) || undefined }
      if (categoriesRef) {
        if (type === 'scatter') item.xValuesRef = categoriesRef
        else item.categoriesRef = categoriesRef
      }
      series.push(item)
    }
  } else {
    if (left > bounds.right) return []
    const categoriesRef = firstRowHeaders ? formatChartRef(sheetName, { top: bounds.top, bottom: bounds.top, left, right: bounds.right }) : undefined
    for (let row = top; row <= bounds.bottom && series.length < 255; row += 1) {
      const item: ChartSeries = { id: newChartId('series'), valuesRef: formatChartRef(sheetName, { top: row, bottom: row, left, right: bounds.right }) }
      if (firstColumnLabels) { item.nameRef = formatChartRef(sheetName, { top: row, bottom: row, left: bounds.left, right: bounds.left }); item.name = nameOf(row, bounds.left) || undefined }
      if (categoriesRef) {
        if (type === 'scatter') item.xValuesRef = categoriesRef
        else item.categoriesRef = categoriesRef
      }
      series.push(item)
    }
  }
  if (type === 'pie' && series.length > 1) return series.slice(0, 1)
  return series
}

export interface AnchorMetrics {
  /** Column width / row height in px at zoom 1 (0 for hidden). */
  columnWidth(col: number): number
  rowHeight(row: number): number
}

/** Anchor covering `widthPx` x `heightPx` starting at a cell. */
export function anchorForSize(start: { row: number; col: number }, widthPx: number, heightPx: number, metrics?: AnchorMetrics): ChartAnchor {
  const colWidth = (col: number) => metrics?.columnWidth(col) ?? 64
  const rowHeight = (row: number) => metrics?.rowHeight(row) ?? 20
  let col = start.col, remainingX = widthPx
  while (col < MAX_COLS - 1 && remainingX > colWidth(col)) { remainingX -= colWidth(col); col += 1 }
  let row = start.row, remainingY = heightPx
  while (row < MAX_ROWS - 1 && remainingY > rowHeight(row)) { remainingY -= rowHeight(row); row += 1 }
  return {
    from: { row: start.row, col: start.col, rowOffsetEmu: 0, colOffsetEmu: 0 },
    to: { row, col, rowOffsetEmu: Math.round(Math.max(0, remainingY) * EMU_PER_PIXEL), colOffsetEmu: Math.round(Math.max(0, remainingX) * EMU_PER_PIXEL) },
  }
}

export const DEFAULT_CHART_WIDTH = 480
export const DEFAULT_CHART_HEIGHT = 288

export interface CreateChartOptions {
  /** Place the chart here instead of beside the data. */
  anchor?: ChartAnchor
  /** Grid metrics used to size the default anchor (px at zoom 1). */
  metrics?: AnchorMetrics
  layout?: Partial<RangeLayout>
  grouping?: ChartGrouping
  title?: string
  /** Expand a single-cell selection to its data block (default true). */
  expandSelection?: boolean
}

function autoTitle(series: ChartSeries[], type: ChartType) {
  const names = series.map((item) => item.name).filter(Boolean) as string[]
  if (type === 'pie' || type === 'doughnut' || names.length === 1) return undefined // Excel's automatic series-name title
  if (names.length === 2) return `${names[0]} and ${names[1]}`
  if (names.length === 3) return `${names[0]}, ${names[1]} and ${names[2]}`
  return 'Chart Title'
}

/**
 * Create a chart from a selected range with Excel/Sheets-like auto detection.
 * `bounds` are 0-based and inclusive; `sheetName` is the sheet holding the data.
 */
export function createChartFromRange(sheetName: string, bounds: CellBounds, accessor: ChartWorkbookAccessor, type: ChartType = 'column', options: CreateChartOptions = {}): SheetChart {
  const range = options.expandSelection === false ? bounds : expandToCurrentRegion(sheetName, bounds, accessor)
  const detected = detectRangeLayout(sheetName, range, accessor, type)
  const layout: RangeLayout = { ...detected, ...options.layout }
  const chartType: ChartType = type === 'unsupported' ? 'column' : type
  const seriesType: ChartType = chartType === 'combo' ? 'column' : chartType
  const series = buildSeriesFromRange(sheetName, range, accessor, layout, seriesType)
  // Excel's "Clustered Column - Line": the last series is a line, the others columns.
  if (chartType === 'combo') series.forEach((item, index) => { item.type = series.length > 1 && index === series.length - 1 ? 'line' : 'column' })
  if (chartType === 'line' || chartType === 'radar') series.forEach((item) => { item.marker = 'none' })
  if (chartType === 'combo') series.forEach((item) => { if (item.type === 'line') item.marker = 'none' })
  if (chartType === 'scatter') series.forEach((item) => { item.marker = 'circle'; item.showLine = false })
  const anchor = options.anchor || anchorForSize({ row: range.top, col: range.right + 2 }, DEFAULT_CHART_WIDTH, DEFAULT_CHART_HEIGHT, options.metrics)
  const chart: SheetChart = {
    id: newChartId(),
    type: chartType,
    anchor,
    series,
    dataRange: formatChartRef(sheetName, range, false),
    seriesIn: layout.seriesIn,
    firstRowHeaders: layout.firstRowHeaders,
    firstColumnLabels: layout.firstColumnLabels,
    legend: chartType === 'pie' || chartType === 'doughnut' || series.length > 1 ? 'bottom' : 'none',
    axes: chartType === 'pie' || chartType === 'doughnut' ? undefined : { x: {}, y: { gridlines: true } },
    modified: true,
  }
  if (chartType === 'column' || chartType === 'bar' || chartType === 'line' || chartType === 'area') chart.grouping = options.grouping || 'clustered'
  if (chartType === 'doughnut') chart.holeSize = 60
  if (chartType === 'scatter') chart.axes = { x: { gridlines: false }, y: { gridlines: true } }
  const title = options.title ?? autoTitle(series, chartType)
  if (title !== undefined) chart.title = title
  return chart
}

/** Rebuild the series of a chart for a new data range text, keeping per-series styling by position. */
export function applyChartDataRange(chart: SheetChart, rangeText: string, accessor: ChartWorkbookAccessor | null, defaultSheet: string, layout?: Partial<RangeLayout>): SheetChart | null {
  const areas = parseChartRef(rangeText, defaultSheet)
  if (!areas || areas.length !== 1) return null
  const area = areas[0]
  const sheet = area.sheet || defaultSheet
  const effectiveType = chart.type === 'combo' ? 'column' : chart.type
  const detected = accessor ? detectRangeLayout(sheet, area, accessor, effectiveType) : { seriesIn: 'columns' as const, firstRowHeaders: true, firstColumnLabels: true }
  const resolvedLayout: RangeLayout = {
    seriesIn: layout?.seriesIn ?? chart.seriesIn ?? detected.seriesIn,
    firstRowHeaders: layout?.firstRowHeaders ?? chart.firstRowHeaders ?? detected.firstRowHeaders,
    firstColumnLabels: layout?.firstColumnLabels ?? chart.firstColumnLabels ?? detected.firstColumnLabels,
  }
  const rebuilt = buildSeriesFromRange(sheet, area, accessor, resolvedLayout, effectiveType)
  const series = rebuilt.map((item, index) => {
    const previous = chart.series[index]
    if (!previous) return chart.type === 'combo' ? { ...item, type: 'line' as ChartSeriesType } : item
    const { id, color, type, secondaryAxis, marker, markerSize, smooth, showLine, lineWidth, dataLabels, invertIfNegative } = previous
    return { ...item, id, color, type, secondaryAxis, marker, markerSize, smooth, showLine, lineWidth, dataLabels, invertIfNegative }
  })
  return {
    ...chart,
    dataRange: rangeText.trim(),
    ...resolvedLayout,
    series: chart.type === 'pie' ? series.slice(0, 1) : series,
    modified: true,
  }
}

/**
 * Best-effort source range for a chart without `dataRange` (e.g. imported from
 * Excel): the bounding box of its name/category/value references when they all
 * live on one sheet and form a regular block.
 */
export function inferChartDataRange(chart: SheetChart): (RangeLayout & { range: string }) | null {
  const series = chart.series || []
  if (!series.length) return null
  let sheet: string | undefined
  const box = { top: Infinity, bottom: -1, left: Infinity, right: -1 }
  let columnVectors = 0, rowVectors = 0
  let hasNames = false, hasCategories = false
  for (const item of series) {
    for (const [key, ref] of [['name', item.nameRef], ['cat', item.categoriesRef || item.xValuesRef], ['val', item.valuesRef]] as const) {
      if (!ref) continue
      const areas = parseChartRef(ref)
      if (!areas || areas.length !== 1 || !areas[0].sheet) return null
      const area = areas[0]
      if (sheet === undefined) sheet = area.sheet
      else if (sheet.toLocaleLowerCase() !== area.sheet!.toLocaleLowerCase()) return null
      box.top = Math.min(box.top, area.top); box.bottom = Math.max(box.bottom, area.bottom)
      box.left = Math.min(box.left, area.left); box.right = Math.max(box.right, area.right)
      if (key === 'val') { if (area.left === area.right) columnVectors += 1; else if (area.top === area.bottom) rowVectors += 1 }
      if (key === 'name') hasNames = true
      if (key === 'cat') hasCategories = true
    }
  }
  if (!sheet || box.bottom < 0) return null
  const seriesIn: 'columns' | 'rows' = rowVectors > columnVectors ? 'rows' : 'columns'
  return {
    range: formatChartRef(sheet, box, false),
    seriesIn,
    firstRowHeaders: seriesIn === 'columns' ? hasNames : hasCategories,
    firstColumnLabels: seriesIn === 'columns' ? hasCategories : hasNames,
  }
}

// ---------------------------------------------------------------------------
// Chart type changes
// ---------------------------------------------------------------------------

export interface ChartTypeOption {
  key: string
  label: string
  type: ChartType
  grouping?: ChartGrouping
}

export const CHART_TYPE_OPTIONS: ChartTypeOption[] = [
  { key: 'column', label: 'Column', type: 'column', grouping: 'clustered' },
  { key: 'column-stacked', label: 'Stacked column', type: 'column', grouping: 'stacked' },
  { key: 'column-percent', label: '100% column', type: 'column', grouping: 'percentStacked' },
  { key: 'bar', label: 'Bar', type: 'bar', grouping: 'clustered' },
  { key: 'bar-stacked', label: 'Stacked bar', type: 'bar', grouping: 'stacked' },
  { key: 'bar-percent', label: '100% bar', type: 'bar', grouping: 'percentStacked' },
  { key: 'line', label: 'Line', type: 'line', grouping: 'clustered' },
  { key: 'line-stacked', label: 'Stacked line', type: 'line', grouping: 'stacked' },
  { key: 'area', label: 'Area', type: 'area', grouping: 'clustered' },
  { key: 'area-stacked', label: 'Stacked area', type: 'area', grouping: 'stacked' },
  { key: 'area-percent', label: '100% area', type: 'area', grouping: 'percentStacked' },
  { key: 'pie', label: 'Pie', type: 'pie' },
  { key: 'doughnut', label: 'Doughnut', type: 'doughnut' },
  { key: 'scatter', label: 'Scatter', type: 'scatter' },
  { key: 'combo', label: 'Combo', type: 'combo' },
  { key: 'radar', label: 'Radar', type: 'radar' },
]

export function chartTypeKey(chart: Pick<SheetChart, 'type' | 'grouping'>) {
  const grouping = chart.grouping || 'clustered'
  const match = CHART_TYPE_OPTIONS.find((option) => option.type === chart.type && (option.grouping || 'clustered') === grouping)
    || CHART_TYPE_OPTIONS.find((option) => option.type === chart.type)
  return match?.key || 'column'
}

/** Switch chart type, adapting series fields the way Excel does. */
export function changeChartType(chart: SheetChart, type: ChartType, grouping?: ChartGrouping): SheetChart {
  if (chart.type === 'unsupported') return chart
  const next: SheetChart = { ...chart, type, modified: true }
  if (type === 'column' || type === 'bar' || type === 'line' || type === 'area') next.grouping = grouping || 'clustered'
  else delete next.grouping
  const wasPie = chart.type === 'pie' || chart.type === 'doughnut'
  const isPie = type === 'pie' || type === 'doughnut'
  next.series = chart.series.map((series, index) => {
    const item = { ...series }
    if (type === 'combo') { item.type = item.type || (chart.series.length > 1 && index === chart.series.length - 1 ? 'line' : 'column'); if (item.type === 'line' && !item.marker) item.marker = 'none' } else delete item.type
    if (type !== 'combo') delete item.secondaryAxis
    if (type === 'scatter') {
      if (!item.xValuesRef && item.categoriesRef) item.xValuesRef = item.categoriesRef
      if (item.showLine === undefined) item.showLine = false
      if (!item.marker || item.marker === 'none') item.marker = 'circle'
    } else if (chart.type === 'scatter') {
      if (!item.categoriesRef && item.xValuesRef) item.categoriesRef = item.xValuesRef
      delete item.xValuesRef
      delete item.showLine
      if (type === 'line') item.marker = 'none'
    }
    if ((type === 'line' || type === 'radar') && chart.type !== 'line' && chart.type !== 'scatter' && !item.marker) item.marker = 'none'
    return item
  })
  if (isPie && !wasPie) {
    next.legend = chart.legend === 'none' || !chart.legend ? 'right' : chart.legend
    delete next.axes
  } else if (!isPie && wasPie) {
    next.axes = { x: {}, y: { gridlines: true } }
    if (next.series.length === 1) next.legend = 'none'
  }
  if (type === 'doughnut' && !next.holeSize) next.holeSize = 60
  if (!isPie && !next.axes) next.axes = { x: {}, y: { gridlines: true } }
  return next
}

/** Mark a user edit (anything other than a move/resize). */
export function updateChart(chart: SheetChart, patch: Partial<SheetChart>): SheetChart {
  return { ...chart, ...patch, modified: true }
}

/** Defensive normalisation for charts coming from files or older sessions. */
export function normalizeChart(input: Partial<SheetChart> & { id?: string }): SheetChart {
  const point = (value: Partial<ChartAnchorPoint> | undefined, fallback: ChartAnchorPoint): ChartAnchorPoint => ({
    row: clampInt(value?.row, 0, MAX_ROWS - 1, fallback.row),
    col: clampInt(value?.col, 0, MAX_COLS - 1, fallback.col),
    rowOffsetEmu: Math.max(0, Math.round(Number(value?.rowOffsetEmu) || 0)),
    colOffsetEmu: Math.max(0, Math.round(Number(value?.colOffsetEmu) || 0)),
  })
  const from = point(input.anchor?.from, { row: 0, col: 0 })
  let to = point(input.anchor?.to, { row: from.row + 15, col: from.col + 8 })
  if (to.row < from.row || (to.row === from.row && (to.rowOffsetEmu || 0) <= (from.rowOffsetEmu || 0))) to = { ...to, row: from.row + 1 }
  if (to.col < from.col || (to.col === from.col && (to.colOffsetEmu || 0) <= (from.colOffsetEmu || 0))) to = { ...to, col: from.col + 1 }
  const types: ChartType[] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'radar', 'combo', 'unsupported']
  return {
    ...input,
    id: typeof input.id === 'string' && input.id ? input.id : newChartId(),
    type: types.includes(input.type as ChartType) ? input.type as ChartType : 'unsupported',
    anchor: { ...input.anchor, from, to },
    series: Array.isArray(input.series) ? input.series.filter((item) => item && typeof item === 'object').map((item) => ({ ...item, id: item.id || newChartId('series') })) : [],
  } as SheetChart
}

function clampInt(value: unknown, min: number, max: number, fallback: number) {
  const number = Math.trunc(Number(value))
  return Number.isFinite(number) ? Math.max(min, Math.min(max, number)) : fallback
}

// ---------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------

export interface ChartRect { left: number; top: number; width: number; height: number }

/** Grid geometry supplied by the host grid, in sheet-canvas pixels at the current zoom. */
export interface ChartGeometry {
  /** Rectangle of a 0-based cell (hidden rows/columns report 0 height/width). */
  cellRect(row: number, col: number): ChartRect
  /** Cell containing a canvas point (clamped to the sheet). */
  cellAtPoint(x: number, y: number): { row: number; col: number }
}

export function anchorToRect(anchor: ChartAnchor, geometry: ChartGeometry, zoom = 1): ChartRect {
  const place = (point: ChartAnchorPoint) => {
    const cell = geometry.cellRect(point.row, point.col)
    const dx = Math.min(cell.width, ((point.colOffsetEmu || 0) / EMU_PER_PIXEL) * zoom)
    const dy = Math.min(cell.height, ((point.rowOffsetEmu || 0) / EMU_PER_PIXEL) * zoom)
    return { x: cell.left + dx, y: cell.top + dy }
  }
  const a = place(anchor.from)
  const b = place(anchor.to)
  return { left: a.x, top: a.y, width: Math.max(1, b.x - a.x), height: Math.max(1, b.y - a.y) }
}

export function pointToAnchorPoint(x: number, y: number, geometry: ChartGeometry, zoom = 1): ChartAnchorPoint {
  const cell = geometry.cellAtPoint(Math.max(0, x), Math.max(0, y))
  const rect = geometry.cellRect(cell.row, cell.col)
  const dx = Math.max(0, Math.min(rect.width, x - rect.left))
  const dy = Math.max(0, Math.min(rect.height, y - rect.top))
  return {
    row: cell.row,
    col: cell.col,
    rowOffsetEmu: Math.round((dy / Math.max(zoom, 0.01)) * EMU_PER_PIXEL),
    colOffsetEmu: Math.round((dx / Math.max(zoom, 0.01)) * EMU_PER_PIXEL),
  }
}

export function rectToAnchor(rect: ChartRect, geometry: ChartGeometry, zoom = 1, editAs?: ChartAnchor['editAs']): ChartAnchor {
  const anchor: ChartAnchor = {
    from: pointToAnchorPoint(rect.left, rect.top, geometry, zoom),
    to: pointToAnchorPoint(rect.left + rect.width, rect.top + rect.height, geometry, zoom),
  }
  if (editAs && editAs !== 'twoCell') anchor.editAs = editAs
  return anchor
}

/** Snap a rectangle's edges to the nearest cell boundaries (Alt-drag in Excel). */
export function snapRectToCells(rect: ChartRect, geometry: ChartGeometry): ChartRect {
  const snap = (x: number, y: number) => {
    const cell = geometry.cellAtPoint(Math.max(0, x), Math.max(0, y))
    const r = geometry.cellRect(cell.row, cell.col)
    return { x: x - r.left < r.width / 2 ? r.left : r.left + r.width, y: y - r.top < r.height / 2 ? r.top : r.top + r.height }
  }
  const a = snap(rect.left, rect.top)
  const b = snap(rect.left + rect.width, rect.top + rect.height)
  return { left: a.x, top: a.y, width: Math.max(8, b.x - a.x), height: Math.max(8, b.y - a.y) }
}

/** Cell bounds covered by an anchor (for print areas / used range). */
export function anchorBounds(anchor: ChartAnchor): CellBounds {
  return { top: anchor.from.row, left: anchor.from.col, bottom: anchor.to.row, right: anchor.to.col }
}

// ---------------------------------------------------------------------------
// Structural edits and sheet renames
// ---------------------------------------------------------------------------

function shiftInterval(start: number, end: number, operation: SheetStructureOperation): [number, number] | null {
  const { kind, index, count } = operation
  if (kind === 'insert') {
    if (index <= start) return [start + count, end + count]
    if (index <= end) return [start, end + count]
    return [start, end]
  }
  const last = index + count - 1
  if (last < start) return [start - count, end - count]
  if (index > end) return [start, end]
  if (index <= start && last >= end) return null
  const newStart = index <= start ? index : start
  const removedInside = Math.min(end, last) - Math.max(start, index) + 1
  return [newStart, end - removedInside]
}

function sameSheet(a: string | undefined, b: string) {
  return (a || '').toLocaleLowerCase() === b.toLocaleLowerCase()
}

/** Rewrite one chart reference for a row/column insert/delete on `targetSheet`. */
export function shiftChartRef(ref: string | undefined, targetSheet: string, operation: SheetStructureOperation, defaultSheet?: string): string | undefined {
  if (!ref) return ref
  const areas = parseChartRef(ref, defaultSheet)
  if (!areas) return ref
  if (!areas.some((area) => sameSheet(area.sheet, targetSheet))) return ref
  const rebuilt: string[] = []
  for (const area of areas) {
    if (!sameSheet(area.sheet, targetSheet)) { rebuilt.push(formatChartRef(area.sheet, area)); continue }
    const isRow = operation.axis === 'row'
    const whole = isRow ? area.top === 0 && area.bottom === MAX_ROWS - 1 : area.left === 0 && area.right === MAX_COLS - 1
    if (whole) { rebuilt.push(formatChartRef(area.sheet, area)); continue }
    const shifted = isRow ? shiftInterval(area.top, area.bottom, operation) : shiftInterval(area.left, area.right, operation)
    if (!shifted) return `${quoteSheetName(area.sheet || targetSheet)}!#REF!`
    const next = isRow ? { ...area, top: shifted[0], bottom: Math.min(MAX_ROWS - 1, shifted[1]) } : { ...area, left: shifted[0], right: Math.min(MAX_COLS - 1, shifted[1]) }
    rebuilt.push(formatChartRef(area.sheet, next))
  }
  return rebuilt.length > 1 ? `(${rebuilt.join(',')})` : rebuilt[0]
}

function shiftAnchorPoint(point: ChartAnchorPoint, operation: SheetStructureOperation, isEnd: boolean): ChartAnchorPoint {
  const key = operation.axis === 'row' ? 'row' : 'col'
  const offsetKey = operation.axis === 'row' ? 'rowOffsetEmu' : 'colOffsetEmu'
  const value = point[key]
  const { index, count, kind } = operation
  if (kind === 'insert') {
    if (index <= value && !(isEnd && index === value && !(point[offsetKey] || 0))) return { ...point, [key]: value + count }
    return point
  }
  const last = index + count - 1
  if (last < value) return { ...point, [key]: value - count }
  if (index > value) return point
  // The anchor cell itself was deleted: collapse onto the first surviving cell.
  return { ...point, [key]: index, [offsetKey]: 0 }
}

/**
 * Shift charts for a row/column insert/delete on `targetSheet`. Charts drawn on
 * that sheet move (and size, for twoCell anchors) with the cells; every
 * chart's references to that sheet are rewritten.
 */
/** A drawing object's anchor after a row/column insert/delete (twoCell sizes, oneCell moves, absolute stays). */
export function transformAnchorForStructure(anchor: ChartAnchor, operation: SheetStructureOperation): ChartAnchor {
  if (anchor.editAs === 'absolute') return anchor
  const from = shiftAnchorPoint(anchor.from, operation, false)
  const key = operation.axis === 'row' ? 'row' : 'col'
  const offsetKey = operation.axis === 'row' ? 'rowOffsetEmu' : 'colOffsetEmu'
  let to = anchor.editAs === 'oneCell'
    ? { ...anchor.to, [key]: anchor.to[key] + (from[key] - anchor.from[key]) } // move only: keep the span
    : shiftAnchorPoint(anchor.to, operation, true)
  if (to[key] < from[key]) to = { ...to, [key]: from[key] }
  if (to[key] === from[key] && (to[offsetKey] || 0) <= (from[offsetKey] || 0)) to = { ...to, [key]: from[key] + 1, [offsetKey]: 0 }
  return { ...anchor, from, to }
}

export function transformChartsForStructureChange(charts: SheetChart[] | undefined, chartSheetName: string, targetSheetName: string, operation: SheetStructureOperation): SheetChart[] | undefined {
  if (!charts?.length) return charts
  return charts.map((chart) => {
    let next: SheetChart = chart
    if (sameSheet(chartSheetName, targetSheetName) && chart.anchor && chart.anchor.editAs !== 'absolute') {
      next = { ...next, anchor: transformAnchorForStructure(chart.anchor, operation) }
    }
    const shift = (ref: string | undefined) => shiftChartRef(ref, targetSheetName, operation, chartSheetName)
    let refsChanged = false
    const series = chart.series.map((item) => {
      const updated = { ...item, nameRef: shift(item.nameRef), categoriesRef: shift(item.categoriesRef), valuesRef: shift(item.valuesRef), xValuesRef: shift(item.xValuesRef) }
      if (updated.nameRef !== item.nameRef || updated.categoriesRef !== item.categoriesRef || updated.valuesRef !== item.valuesRef || updated.xValuesRef !== item.xValuesRef) refsChanged = true
      for (const key of ['nameRef', 'categoriesRef', 'valuesRef', 'xValuesRef'] as const) if (updated[key] === undefined) delete updated[key]
      return updated
    })
    const dataRange = chart.dataRange ? shiftChartRef(chart.dataRange, targetSheetName, operation, chartSheetName) : chart.dataRange
    const titleRef = shift(chart.titleRef)
    if (refsChanged || dataRange !== chart.dataRange || titleRef !== chart.titleRef) {
      next = { ...next, series, modified: true }
      if (dataRange !== undefined) next.dataRange = dataRange
      if (titleRef !== undefined) next.titleRef = titleRef
    }
    return next
  })
}

/** Apply `transformChartsForStructureChange` across every sheet of a workbook (call after applySheetStructureOperation). */
export function transformWorkbookChartsForStructure(workbook: WorkbookModel, sheetId: string, operation: SheetStructureOperation): WorkbookModel {
  const target = workbook.sheets.find((sheet) => sheet.id === sheetId)
  if (!target || !workbook.sheets.some((sheet) => sheet.charts?.length || (sheet.id === sheetId && sheet.images?.length))) return workbook
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      let next = sheet
      if (sheet.charts?.length) next = { ...next, charts: transformChartsForStructureChange(sheet.charts, sheet.name, target.name, operation) }
      // Pictures move (and size) with the cells of their own sheet.
      if (sheet.id === sheetId && sheet.images?.length && !operation.span) next = { ...next, images: sheet.images.map((image) => ({ ...image, anchor: transformAnchorForStructure(image.anchor, operation) })) }
      return next
    }),
  }
}

function renameInRef(ref: string | undefined, oldName: string, newName: string) {
  if (!ref) return ref
  const areas = parseChartRef(ref)
  if (!areas || !areas.some((area) => sameSheet(area.sheet, oldName))) return ref
  const parts = areas.map((area) => formatChartRef(area.sheet && sameSheet(area.sheet, oldName) ? newName : area.sheet, area))
  return parts.length > 1 ? `(${parts.join(',')})` : parts[0]
}

/** Update chart references after a sheet rename (Excel rewrites them; so must we). */
export function renameChartSheetReferences(charts: SheetChart[] | undefined, oldName: string, newName: string): SheetChart[] | undefined {
  if (!charts?.length || oldName === newName) return charts
  return charts.map((chart) => {
    let changed = false
    const rename = (ref: string | undefined) => { const next = renameInRef(ref, oldName, newName); if (next !== ref) changed = true; return next }
    const series = chart.series.map((item) => {
      const updated = { ...item }
      for (const key of ['nameRef', 'categoriesRef', 'valuesRef', 'xValuesRef'] as const) if (item[key]) updated[key] = rename(item[key])
      return updated
    })
    const dataRange = chart.dataRange && /!/.test(chart.dataRange) ? rename(chart.dataRange) : chart.dataRange
    const titleRef = rename(chart.titleRef)
    return changed ? { ...chart, series, dataRange, ...(titleRef ? { titleRef } : {}), modified: true } : chart
  })
}

/** Apply a sheet rename to every chart in the workbook. */
export function renameWorkbookChartReferences(workbook: WorkbookModel, oldName: string, newName: string): WorkbookModel {
  if (!workbook.sheets.some((sheet) => sheet.charts?.length)) return workbook
  return { ...workbook, sheets: workbook.sheets.map((sheet) => (sheet.charts?.length ? { ...sheet, charts: renameChartSheetReferences(sheet.charts, oldName, newName) } : sheet)) }
}
