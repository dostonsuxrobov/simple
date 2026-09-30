/**
 * Pure helpers behind the formatting UI:
 *  - colour maths (theme palette, Excel tints, resolution to CSS),
 *  - `FormatCellsChange` and Excel's semantics for applying it to every cell of a selection,
 *  - Google Sheets-style border presets,
 *  - Excel's built-in Cell Styles gallery.
 */
import type { CSSProperties } from 'react'
import type {
  CellAlignment, CellBorder, CellBorderSide, CellData, CellFill, CellFont, CellStyle, SpreadsheetColor,
} from '../spreadsheet-types'

export type ColorValue = string | SpreadsheetColor

// ---------------------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------------------

/** Office 2013-2022 theme in OOXML theme-index order (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink). */
export const DEFAULT_THEME_COLORS: readonly string[] = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31',
  'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72',
]

export const THEME_COLOR_NAMES: readonly string[] = [
  'Background 1', 'Text 1', 'Background 2', 'Text 2',
  'Accent 1', 'Accent 2', 'Accent 3', 'Accent 4', 'Accent 5', 'Accent 6',
  'Hyperlink', 'Followed Hyperlink',
]

export const STANDARD_COLORS: readonly { hex: string; name: string }[] = [
  { hex: 'C00000', name: 'Dark Red' },
  { hex: 'FF0000', name: 'Red' },
  { hex: 'FFC000', name: 'Orange' },
  { hex: 'FFFF00', name: 'Yellow' },
  { hex: '92D050', name: 'Light Green' },
  { hex: '00B050', name: 'Green' },
  { hex: '00B0F0', name: 'Light Blue' },
  { hex: '0070C0', name: 'Blue' },
  { hex: '002060', name: 'Dark Blue' },
  { hex: '7030A0', name: 'Purple' },
]

/** Excel's legacy indexed palette (0-63); 64 is the system foreground, 65 the system background. */
export const EXCEL_INDEXED_COLORS: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
  '000000', 'FFFFFF',
]

const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value))

/** Normalizes `#rgb`, `#rrggbb`, `rrggbb` or `aarrggbb` to uppercase `RRGGBB`; null if invalid. */
export function normalizeHex(input: string | undefined | null): string | null {
  let value = String(input ?? '').trim().replace(/^#/, '')
  if (/^[0-9a-f]{3}$/i.test(value)) value = value.split('').map((char) => char + char).join('')
  if (/^[0-9a-f]{8}$/i.test(value)) value = value.slice(2)
  return /^[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : null
}

export interface Rgb { r: number; g: number; b: number }
export interface Hsl { h: number; s: number; l: number }
export interface Hsv { h: number; s: number; v: number }

export function hexToRgb(hex: string): Rgb {
  const value = normalizeHex(hex) || '000000'
  return { r: parseInt(value.slice(0, 2), 16), g: parseInt(value.slice(2, 4), 16), b: parseInt(value.slice(4, 6), 16) }
}

export function rgbToHex({ r, g, b }: Rgb): string {
  return [r, g, b].map((channel) => Math.round(clamp(channel, 0, 255)).toString(16).padStart(2, '0')).join('').toUpperCase()
}

/** h in degrees 0-360, s/l in 0-1. */
export function rgbToHsl({ r, g, b }: Rgb): Hsl {
  const red = r / 255, green = g / 255, blue = b / 255
  const max = Math.max(red, green, blue), min = Math.min(red, green, blue)
  const l = (max + min) / 2
  if (max === min) return { h: 0, s: 0, l }
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const h = max === red ? (green - blue) / d + (green < blue ? 6 : 0) : max === green ? (blue - red) / d + 2 : (red - green) / d + 4
  return { h: h * 60, s, l }
}

export function hslToRgb({ h, s, l }: Hsl): Rgb {
  const hue = (((h % 360) + 360) % 360) / 360
  if (s === 0) return { r: l * 255, g: l * 255, b: l * 255 }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const channel = (t: number) => {
    let x = t
    if (x < 0) x += 1
    if (x > 1) x -= 1
    if (x < 1 / 6) return p + (q - p) * 6 * x
    if (x < 1 / 2) return q
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6
    return p
  }
  return { r: channel(hue + 1 / 3) * 255, g: channel(hue) * 255, b: channel(hue - 1 / 3) * 255 }
}

/** h in degrees 0-360, s/v in 0-1. */
export function rgbToHsv({ r, g, b }: Rgb): Hsv {
  const red = r / 255, green = g / 255, blue = b / 255
  const max = Math.max(red, green, blue), min = Math.min(red, green, blue)
  const d = max - min
  let h = 0
  if (d) h = max === red ? ((green - blue) / d) % 6 : max === green ? (blue - red) / d + 2 : (red - green) / d + 4
  return { h: ((h * 60) + 360) % 360, s: max ? d / max : 0, v: max }
}

export function hsvToRgb({ h, s, v }: Hsv): Rgb {
  const hue = (((h % 360) + 360) % 360) / 60
  const c = v * s
  const x = c * (1 - Math.abs((hue % 2) - 1))
  const m = v - c
  const [r, g, b] = hue < 1 ? [c, x, 0] : hue < 2 ? [x, c, 0] : hue < 3 ? [0, c, x] : hue < 4 ? [0, x, c] : hue < 5 ? [x, 0, c] : [c, 0, x]
  return { r: (r + m) * 255, g: (g + m) * 255, b: (b + m) * 255 }
}

/** Excel's tint: scales HLS luminance toward black (tint < 0) or white (tint > 0). */
export function applyTint(hex: string, tint = 0): string {
  const base = normalizeHex(hex) || '000000'
  if (!Number.isFinite(tint) || tint === 0) return base
  const amount = clamp(tint, -1, 1)
  const hsl = rgbToHsl(hexToRgb(base))
  const l = amount < 0 ? hsl.l * (1 + amount) : hsl.l * (1 - amount) + amount
  return rgbToHex(hslToRgb({ ...hsl, l: clamp(l, 0, 1) }))
}

function themeHex(themeColors: readonly string[] | undefined, index: number) {
  return normalizeHex(themeColors?.[index]) || DEFAULT_THEME_COLORS[index] || ''
}

/** Resolves a spreadsheet colour to `#RRGGBB`, or '' when it is automatic/unknown. */
export function resolveColorHex(color: ColorValue | null | undefined, themeColors: readonly string[] = DEFAULT_THEME_COLORS): string {
  if (!color) return ''
  let hex: string | null = null
  let tint = 0
  if (typeof color === 'string') hex = normalizeHex(color)
  else {
    tint = Number(color.tint) || 0
    hex = normalizeHex(color.argb || color.rgb)
    if (!hex && Number.isInteger(color.theme)) hex = themeHex(themeColors, Number(color.theme)) || null
    if (!hex && Number.isInteger(color.indexed)) hex = EXCEL_INDEXED_COLORS[Number(color.indexed)] || null
  }
  return hex ? `#${applyTint(hex, tint)}` : ''
}

/** An explicit RGB colour in the ExcelJS shape the app writes. */
export function argbColor(hex: string): SpreadsheetColor {
  return { argb: `FF${normalizeHex(hex) || '000000'}` }
}

/** Excel's five tint/shade steps for a theme colour, chosen from its luminance. */
export function themeTintSteps(hex: string): number[] {
  const { l } = rgbToHsl(hexToRgb(hex))
  if (l <= 0.005) return [0.5, 0.35, 0.25, 0.15, 0.05]
  if (l >= 0.995) return [-0.05, -0.15, -0.25, -0.35, -0.5]
  if (l < 0.2) return [0.9, 0.75, 0.5, 0.25, 0.1]
  if (l > 0.8) return [-0.1, -0.25, -0.5, -0.75, -0.9]
  return [0.8, 0.6, 0.4, -0.25, -0.5]
}

export function tintLabel(tint: number) {
  if (!tint) return ''
  return `${tint > 0 ? 'Lighter' : 'Darker'} ${Math.round(Math.abs(tint) * 100)}%`
}

export interface PaletteSwatch {
  color: SpreadsheetColor
  hex: string
  label: string
}

/** Excel's theme grid: 10 columns (Background 1 … Accent 6) × (base row + 5 tint rows). */
export function themePalette(themeColors: readonly string[] = DEFAULT_THEME_COLORS): PaletteSwatch[][] {
  const columns = Array.from({ length: 10 }, (_, index) => themeHex(themeColors, index))
  const base = columns.map((hex, theme) => ({ color: { theme }, hex: `#${hex}`, label: THEME_COLOR_NAMES[theme] }))
  const rows = [0, 1, 2, 3, 4].map((step) => columns.map((hex, theme) => {
    const tint = themeTintSteps(hex)[step]
    return { color: { theme, tint }, hex: `#${applyTint(hex, tint)}`, label: `${THEME_COLOR_NAMES[theme]}, ${tintLabel(tint)}` }
  }))
  return [base, ...rows]
}

export function standardPalette(): PaletteSwatch[] {
  return STANDARD_COLORS.map((entry) => ({ color: argbColor(entry.hex), hex: `#${entry.hex}`, label: entry.name }))
}

/** Same colour: theme references compare by index and tint, everything else by resolved RGB. */
export function sameColor(a: ColorValue | null | undefined, b: ColorValue | null | undefined, themeColors: readonly string[] = DEFAULT_THEME_COLORS) {
  if (!a || !b) return !a && !b
  const aTheme = typeof a === 'object' && Number.isInteger(a.theme)
  const bTheme = typeof b === 'object' && Number.isInteger(b.theme)
  if (aTheme && bTheme) {
    const left = a as SpreadsheetColor, right = b as SpreadsheetColor
    return left.theme === right.theme && Math.abs((Number(left.tint) || 0) - (Number(right.tint) || 0)) < 0.005
  }
  if (aTheme !== bTheme) return false
  const left = resolveColorHex(a, themeColors)
  return Boolean(left) && left === resolveColorHex(b, themeColors)
}

export function colorLabel(color: ColorValue | null | undefined, themeColors: readonly string[] = DEFAULT_THEME_COLORS) {
  if (!color) return 'Automatic'
  if (typeof color === 'object' && Number.isInteger(color.theme)) {
    const tint = tintLabel(Number(color.tint) || 0)
    return `${THEME_COLOR_NAMES[Number(color.theme)] || `Theme ${color.theme}`}${tint ? `, ${tint}` : ''}`
  }
  const hex = resolveColorHex(color, themeColors)
  const standard = STANDARD_COLORS.find((entry) => `#${entry.hex}` === hex)
  return standard ? standard.name : hex || 'Automatic'
}

// ---------------------------------------------------------------------------------------
// Format Cells change model
// ---------------------------------------------------------------------------------------

/** Each key: undefined = unchanged, null = remove the property, value = set it. */
export type NullablePartial<T> = { [K in keyof T]?: T[K] | null }

/**
 * Border edits relative to the whole selection. Each key: undefined = unchanged,
 * null = remove, a side = draw it. `outline` is shorthand for all four outer edges;
 * an explicit `top`/`bottom`/`left`/`right` wins over it.
 */
export interface FormatBorderChange {
  outline?: CellBorderSide | null
  top?: CellBorderSide | null
  bottom?: CellBorderSide | null
  left?: CellBorderSide | null
  right?: CellBorderSide | null
  /** Edges between rows inside the selection. */
  insideHorizontal?: CellBorderSide | null
  /** Edges between columns inside the selection. */
  insideVertical?: CellBorderSide | null
  /** Bottom-left → top-right. OOXML stores one diagonal style per cell, shared by both. */
  diagonalUp?: CellBorderSide | null
  /** Top-left → bottom-right. */
  diagonalDown?: CellBorderSide | null
}

/**
 * Everything the Format Cells dialog changed — only what the user touched is present.
 * Apply to each selected cell with `applyFormatChange` (style) or `applyFormatChangeToCell`
 * (also moves `numFmt` to the cell, where it takes precedence over `style.numFmt`).
 */
export interface FormatCellsChange {
  /** Excel format code; 'General' clears the number format. */
  numFmt?: string
  alignment?: NullablePartial<CellAlignment>
  font?: NullablePartial<CellFont>
  /** null = No Color (remove the fill). */
  fill?: CellFill | null
  borders?: FormatBorderChange
  protection?: { locked?: boolean; hidden?: boolean }
  /** Selection-level: true = merge the selection, false = unmerge. Not applied per cell. */
  merge?: boolean
}

/** Where a cell sits in the selection: true when it lies on that outer edge. */
export interface CellPosition {
  top: boolean
  bottom: boolean
  left: boolean
  right: boolean
}

export const SINGLE_CELL_POSITION: CellPosition = { top: true, bottom: true, left: true, right: true }

/** Builds a CellPosition from 0-based coordinates and inclusive selection bounds. */
export function cellPosition(row: number, col: number, bounds: { top: number; bottom: number; left: number; right: number }): CellPosition {
  return { top: row === bounds.top, bottom: row === bounds.bottom, left: col === bounds.left, right: col === bounds.right }
}

function clonePlain<T>(value: T): T {
  if (Array.isArray(value)) return value.map((item) => clonePlain(item)) as T
  if (value && typeof value === 'object') {
    const copy: Record<string, unknown> = {}
    for (const [key, item] of Object.entries(value as Record<string, unknown>)) copy[key] = clonePlain(item)
    return copy as T
  }
  return value
}

function mergeNullable<T extends object>(current: T | undefined, change: NullablePartial<T>): T | undefined {
  const next: Record<string, unknown> = { ...(current ? clonePlain(current) : {}) }
  for (const [key, value] of Object.entries(change)) {
    if (value === undefined) continue
    if (value === null) delete next[key]
    else next[key] = clonePlain(value)
  }
  return Object.keys(next).length ? next as T : undefined
}

type Edge = 'top' | 'bottom' | 'left' | 'right'

function edgeChange(borders: FormatBorderChange, edge: Edge, position: CellPosition) {
  if (position[edge]) return borders[edge] !== undefined ? borders[edge] : borders.outline
  return edge === 'top' || edge === 'bottom' ? borders.insideHorizontal : borders.insideVertical
}

function hasBorderContent(border: CellBorder) {
  return Object.entries(border).some(([key, value]) => (
    (key === 'diagonalUp' || key === 'diagonalDown') ? value === true : Boolean(value && typeof value === 'object')
  ))
}

function applyBorderChange(border: CellBorder | undefined, borders: FormatBorderChange, position: CellPosition): CellBorder | undefined {
  const next: CellBorder = border ? clonePlain(border) : {}
  for (const edge of ['top', 'bottom', 'left', 'right'] as const) {
    const side = edgeChange(borders, edge, position)
    if (side === undefined) continue
    if (side === null || !side.style || side.style === 'none') delete next[edge]
    else next[edge] = clonePlain(side)
  }
  for (const key of ['diagonalUp', 'diagonalDown'] as const) {
    const side = borders[key]
    if (side === undefined) continue
    if (side === null || !side.style || side.style === 'none') delete next[key]
    else {
      next[key] = true
      next.diagonal = clonePlain(side)
    }
  }
  if (!next.diagonalUp && !next.diagonalDown) {
    delete next.diagonal
    delete next.diagonalUp
    delete next.diagonalDown
  }
  return hasBorderContent(next) ? next : undefined
}

/**
 * Applies a Format Cells change to one cell's style with Excel's semantics: the outline
 * draws only on the selection's outer edges, inside borders only on interior edges. Pure —
 * the input is never mutated. `style.numFmt` is written for `numFmt`; prefer
 * `applyFormatChangeToCell` when the cell-level `numFmt` must follow.
 */
export function applyFormatChange(style: CellStyle | undefined, change: FormatCellsChange, position: CellPosition = SINGLE_CELL_POSITION): CellStyle {
  const next: CellStyle = style ? { ...style } : {}
  if (change.numFmt !== undefined) {
    const code = change.numFmt.trim()
    if (!code || /^general$/i.test(code)) delete next.numFmt
    else next.numFmt = code
  }
  if (change.alignment) {
    const alignment = mergeNullable<CellAlignment>(next.alignment, change.alignment)
    if (alignment) next.alignment = alignment
    else delete next.alignment
  }
  if (change.font) {
    const font = mergeNullable<CellFont>(next.font, change.font)
    if (font) next.font = font
    else delete next.font
  }
  if (change.fill !== undefined) {
    if (change.fill === null) delete next.fill
    else next.fill = clonePlain(change.fill)
  }
  if (change.borders) {
    const border = applyBorderChange(next.border, change.borders, position)
    if (border) next.border = border
    else delete next.border
  }
  if (change.protection) {
    const protection = { ...(next.protection || {}), ...change.protection }
    // Locked and visible is the default; omit it to keep files lean.
    if (protection.locked !== false && !protection.hidden) delete next.protection
    else next.protection = protection
  }
  return next
}

function cellHasStyle(style: CellStyle) {
  return Object.values(style).some((value) => value !== undefined)
}

/** Like `applyFormatChange`, but on a whole cell: the number format lands on `cell.numFmt`. */
export function applyFormatChangeToCell(cell: CellData | undefined, change: FormatCellsChange, position: CellPosition = SINGLE_CELL_POSITION): CellData {
  const next: CellData = cell ? { ...cell } : {}
  const { numFmt, ...rest } = change
  const style = applyFormatChange(next.style, rest, position)
  if (numFmt !== undefined) {
    const code = numFmt.trim()
    delete style.numFmt
    if (!code || /^general$/i.test(code)) delete next.numFmt
    else next.numFmt = code
  }
  if (cellHasStyle(style)) next.style = style
  else delete next.style
  return next
}

export type NeighborDirection = 'above' | 'below' | 'left' | 'right'

/**
 * Excel draws a shared edge once. When the selection's outer edge is restyled or removed,
 * the facing side of the cell just outside the selection is cleared so no stale line remains
 * (e.g. the row above loses its bottom border). Returns the neighbour's new style.
 */
export function applyNeighborBorderChange(style: CellStyle | undefined, borders: FormatBorderChange, neighbor: NeighborDirection): CellStyle | undefined {
  const selectionEdge: Edge = neighbor === 'above' ? 'top' : neighbor === 'below' ? 'bottom' : neighbor
  const facing: Edge = neighbor === 'above' ? 'bottom' : neighbor === 'below' ? 'top' : neighbor === 'left' ? 'right' : 'left'
  const edge = borders[selectionEdge] !== undefined ? borders[selectionEdge] : borders.outline
  if (edge === undefined || !style?.border?.[facing]) return style
  const border = { ...style.border }
  delete border[facing]
  const next = { ...style }
  if (hasBorderContent(border)) next.border = border
  else delete next.border
  return next
}

// ---------------------------------------------------------------------------------------
// Border presets (Google Sheets border menu)
// ---------------------------------------------------------------------------------------

export type BorderPreset = 'all' | 'inner' | 'horizontal' | 'vertical' | 'outer' | 'left' | 'top' | 'right' | 'bottom' | 'clear'

export const BORDER_LINE_STYLES = [
  'thin', 'hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot',
  'medium', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot', 'thick', 'double',
] as const

export type BorderLineStyle = typeof BORDER_LINE_STYLES[number]

export const DEFAULT_BORDER_SIDE: CellBorderSide = { style: 'thin', color: { argb: 'FF000000' } }

/** The selection-relative border change a preset makes. */
export function borderPresetChange(preset: BorderPreset, side: CellBorderSide = DEFAULT_BORDER_SIDE): FormatBorderChange {
  switch (preset) {
    case 'all': return { outline: side, insideHorizontal: side, insideVertical: side }
    case 'inner': return { insideHorizontal: side, insideVertical: side }
    case 'horizontal': return { insideHorizontal: side }
    case 'vertical': return { insideVertical: side }
    case 'outer': return { outline: side }
    case 'left': return { left: side }
    case 'top': return { top: side }
    case 'right': return { right: side }
    case 'bottom': return { bottom: side }
    case 'clear': return { outline: null, insideHorizontal: null, insideVertical: null, diagonalUp: null, diagonalDown: null }
  }
}

/**
 * Applies a border preset to one cell of the selection. `position` says which outer edges
 * the cell lies on; the result is the cell's new style (the input is not mutated).
 */
export function applyBorderPreset(preset: BorderPreset, side: CellBorderSide, position: CellPosition, style?: CellStyle): CellStyle {
  return applyFormatChange(style, { borders: borderPresetChange(preset, side) }, position)
}

// ---------------------------------------------------------------------------------------
// Rendering helpers for previews
// ---------------------------------------------------------------------------------------

export function borderSideCss(side: CellBorderSide | undefined, themeColors: readonly string[] = DEFAULT_THEME_COLORS) {
  if (!side?.style || side.style === 'none') return undefined
  const style = side.style
  const color = resolveColorHex(side.color, themeColors) || '#000000'
  const width = /thick/.test(style) ? 3 : /medium/.test(style) || style === 'slantDashDot' ? 2 : 1
  if (style === 'double') return `3px double ${color}`
  const line = /dot/i.test(style) && !/dash/i.test(style) || style === 'hair' ? 'dotted' : /dash/i.test(style) ? 'dashed' : 'solid'
  return `${width}px ${line} ${color}`
}

/** OOXML pattern fills in Excel's Pattern Style gallery order (Solid is the plain background). */
export const PATTERN_FILLS: readonly { id: string; label: string }[] = [
  { id: 'none', label: 'Solid (no pattern)' },
  { id: 'darkGray', label: '75% Gray' },
  { id: 'mediumGray', label: '50% Gray' },
  { id: 'lightGray', label: '25% Gray' },
  { id: 'gray125', label: '12.5% Gray' },
  { id: 'gray0625', label: '6.25% Gray' },
  { id: 'darkHorizontal', label: 'Horizontal Stripe' },
  { id: 'darkVertical', label: 'Vertical Stripe' },
  { id: 'darkDown', label: 'Reverse Diagonal Stripe' },
  { id: 'darkUp', label: 'Diagonal Stripe' },
  { id: 'darkGrid', label: 'Diagonal Crosshatch' },
  { id: 'darkTrellis', label: 'Thick Diagonal Crosshatch' },
  { id: 'lightHorizontal', label: 'Thin Horizontal Stripe' },
  { id: 'lightVertical', label: 'Thin Vertical Stripe' },
  { id: 'lightDown', label: 'Thin Reverse Diagonal Stripe' },
  { id: 'lightUp', label: 'Thin Diagonal Stripe' },
  { id: 'lightGrid', label: 'Thin Horizontal Crosshatch' },
  { id: 'lightTrellis', label: 'Thin Diagonal Crosshatch' },
]

const mod = (value: number, divisor: number) => ((value % divisor) + divisor) % divisor

/** Pixel rule per pattern: [tile size, is-ink(x, y)]. */
const PATTERN_PIXELS: Record<string, [number, (x: number, y: number) => boolean]> = {
  darkGray: [4, (x, y) => x !== (y % 2 === 0 ? 0 : 2)],
  mediumGray: [2, (x, y) => (x + y) % 2 === 0],
  lightGray: [4, (x, y) => x === (y % 2 === 0 ? 0 : 2)],
  gray125: [4, (x, y) => (x === 0 && y === 0) || (x === 2 && y === 2)],
  gray0625: [8, (x, y) => y % 2 === 0 && x === (y % 4 === 0 ? 0 : 4)],
  darkHorizontal: [4, (_x, y) => y < 2],
  darkVertical: [4, (x) => x < 2],
  darkDown: [4, (x, y) => mod(x - y, 4) < 2],
  darkUp: [4, (x, y) => mod(x + y, 4) < 2],
  darkGrid: [4, (x, y) => (Math.floor(x / 2) + Math.floor(y / 2)) % 2 === 0],
  darkTrellis: [4, (x, y) => !((x === 1 && y === 0) || (x === 3 && y === 2))],
  lightHorizontal: [4, (_x, y) => y === 0],
  lightVertical: [4, (x) => x === 0],
  lightDown: [4, (x, y) => mod(x - y, 4) === 0],
  lightUp: [4, (x, y) => mod(x + y, 4) === 3],
  lightGrid: [4, (x, y) => x === 0 || y === 0],
  lightTrellis: [4, (x, y) => mod(x - y, 4) === 0 || mod(x + y, 4) === 2],
}

const patternUriCache = new Map<string, string>()

/** A repeating SVG tile (data URI) drawing a pattern's ink pixels; transparent elsewhere. */
export function patternTileUri(pattern: string, inkHex: string, scale = 1): string {
  const rule = PATTERN_PIXELS[pattern]
  if (!rule) return ''
  const key = `${pattern}|${inkHex}|${scale}`
  const cached = patternUriCache.get(key)
  if (cached) return cached
  const [size, ink] = rule
  let rects = ''
  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) if (ink(x, y)) rects += `<rect x="${x}" y="${y}" width="1" height="1"/>`
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size * scale}" height="${size * scale}" viewBox="0 0 ${size} ${size}" shape-rendering="crispEdges"><g fill="${inkHex}">${rects}</g></svg>`
  const uri = `url("data:image/svg+xml,${encodeURIComponent(svg)}")`
  patternUriCache.set(key, uri)
  return uri
}

/** CSS for a fill: solid colour, or background colour plus a pixel-accurate pattern tile. */
export function fillPreviewCss(fill: CellFill | null | undefined, themeColors: readonly string[] = DEFAULT_THEME_COLORS, scale = 1): CSSProperties {
  if (!fill || fill.type === 'gradient') return {}
  const pattern = String(fill.pattern || 'solid')
  if (pattern === 'none') return {}
  if (pattern === 'solid') {
    const color = resolveColorHex(fill.fgColor || fill.color, themeColors)
    return color ? { backgroundColor: color } : {}
  }
  const background = resolveColorHex(fill.bgColor, themeColors) || '#FFFFFF'
  const ink = resolveColorHex(fill.fgColor || fill.color, themeColors) || '#000000'
  const tile = patternTileUri(pattern, ink, scale)
  return tile ? { backgroundColor: background, backgroundImage: tile, backgroundRepeat: 'repeat' } : { backgroundColor: background }
}

/** Approximate CSS for a style tile or sample (fonts, solid fills, borders). */
export function stylePreviewCss(style: CellStyle | undefined, themeColors: readonly string[] = DEFAULT_THEME_COLORS): CSSProperties {
  const css: CSSProperties = {}
  const font = style?.font
  if (font) {
    const color = resolveColorHex(font.color, themeColors)
    if (color) css.color = color
    if (font.bold) css.fontWeight = 700
    if (font.italic) css.fontStyle = 'italic'
    const lines = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean)
    if (lines.length) css.textDecorationLine = lines.join(' ')
    if (font.underline === 'double' || font.underline === 'doubleAccounting') css.textDecorationStyle = 'double'
    if (font.name) css.fontFamily = `"${String(font.name).replace(/["\\]/g, '')}", "Segoe UI", sans-serif`
  }
  Object.assign(css, fillPreviewCss(style?.fill, themeColors))
  const border = style?.border
  if (border) {
    css.borderTop = borderSideCss(border.top, themeColors)
    css.borderBottom = borderSideCss(border.bottom, themeColors)
    css.borderLeft = borderSideCss(border.left, themeColors)
    css.borderRight = borderSideCss(border.right, themeColors)
  }
  return css
}

// ---------------------------------------------------------------------------------------
// Excel built-in Cell Styles (Excel 2013+ Office theme)
// ---------------------------------------------------------------------------------------

export type CellStyleGroup = 'good-bad' | 'data' | 'titles' | 'themed' | 'number'
export type CellStyleInclude = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection'

export interface CellStylePreset {
  id: string
  name: string
  group: CellStyleGroup
  style: CellStyle
  /** Formatting categories the style sets (Excel's "Style Includes"); others are left alone. */
  includes: CellStyleInclude[]
}

export const CELL_STYLE_GROUPS: readonly { id: CellStyleGroup; label: string }[] = [
  { id: 'good-bad', label: 'Good, Bad and Neutral' },
  { id: 'data', label: 'Data and Model' },
  { id: 'titles', label: 'Titles and Headings' },
  { id: 'themed', label: 'Themed Cell Styles' },
  { id: 'number', label: 'Number Format' },
]

const rgb = (hex: string): SpreadsheetColor => ({ argb: `FF${hex}` })
const solid = (color: SpreadsheetColor): CellFill => ({ type: 'pattern', pattern: 'solid', fgColor: color })
const box = (style: string, color: SpreadsheetColor): CellBorder => ({
  top: { style, color }, left: { style, color }, bottom: { style, color }, right: { style, color },
})

/** Heading font of the Office theme (Title style). */
export const HEADING_FONT_NAME = 'Aptos Display'

const ACCENT_STYLES: CellStylePreset[] = [0, 1, 2, 3, 4, 5].flatMap((offset) => {
  const theme = 4 + offset
  const accent = offset + 1
  return [
    { id: `accent${accent}-20`, name: `20% - Accent${accent}`, group: 'themed' as const, includes: ['font', 'fill'] as CellStyleInclude[], style: { font: { color: { theme: 1 } }, fill: solid({ theme, tint: 0.7999816888943144 }) } },
    { id: `accent${accent}-40`, name: `40% - Accent${accent}`, group: 'themed' as const, includes: ['font', 'fill'] as CellStyleInclude[], style: { font: { color: { theme: 1 } }, fill: solid({ theme, tint: 0.5999938962981048 }) } },
    { id: `accent${accent}-60`, name: `60% - Accent${accent}`, group: 'themed' as const, includes: ['font', 'fill'] as CellStyleInclude[], style: { font: { color: { theme: 0 } }, fill: solid({ theme, tint: 0.3999755851924192 }) } },
    { id: `accent${accent}`, name: `Accent${accent}`, group: 'themed' as const, includes: ['font', 'fill'] as CellStyleInclude[], style: { font: { color: { theme: 0 } }, fill: solid({ theme }) } },
  ]
})

// Order themed styles the way Excel's gallery lays them out: one row per intensity.
const THEMED_ORDER = ['-20', '-40', '-60', ''].flatMap((suffix) => [1, 2, 3, 4, 5, 6].map((accent) => `accent${accent}${suffix}`))

export const CELL_STYLES: readonly CellStylePreset[] = [
  { id: 'normal', name: 'Normal', group: 'good-bad', includes: ['number', 'alignment', 'font', 'border', 'fill', 'protection'], style: {} },
  { id: 'bad', name: 'Bad', group: 'good-bad', includes: ['font', 'fill'], style: { font: { color: rgb('9C0006') }, fill: solid(rgb('FFC7CE')) } },
  { id: 'good', name: 'Good', group: 'good-bad', includes: ['font', 'fill'], style: { font: { color: rgb('006100') }, fill: solid(rgb('C6EFCE')) } },
  { id: 'neutral', name: 'Neutral', group: 'good-bad', includes: ['font', 'fill'], style: { font: { color: rgb('9C5700') }, fill: solid(rgb('FFEB9C')) } },

  { id: 'calculation', name: 'Calculation', group: 'data', includes: ['font', 'fill', 'border'], style: { font: { bold: true, color: rgb('FA7D00') }, fill: solid(rgb('F2F2F2')), border: box('thin', rgb('7F7F7F')) } },
  { id: 'check-cell', name: 'Check Cell', group: 'data', includes: ['font', 'fill', 'border'], style: { font: { bold: true, color: { theme: 0 } }, fill: solid(rgb('A5A5A5')), border: box('double', rgb('3F3F3F')) } },
  { id: 'explanatory', name: 'Explanatory Text', group: 'data', includes: ['font'], style: { font: { italic: true, color: rgb('7F7F7F') } } },
  { id: 'input', name: 'Input', group: 'data', includes: ['font', 'fill', 'border'], style: { font: { color: rgb('3F3F76') }, fill: solid(rgb('FFCC99')), border: box('thin', rgb('7F7F7F')) } },
  { id: 'linked-cell', name: 'Linked Cell', group: 'data', includes: ['font', 'border'], style: { font: { color: rgb('FA7D00') }, border: { bottom: { style: 'double', color: rgb('FF8001') } } } },
  { id: 'note', name: 'Note', group: 'data', includes: ['fill', 'border'], style: { fill: solid(rgb('FFFFCC')), border: box('thin', rgb('B2B2B2')) } },
  { id: 'output', name: 'Output', group: 'data', includes: ['font', 'fill', 'border'], style: { font: { bold: true, color: rgb('3F3F3F') }, fill: solid(rgb('F2F2F2')), border: box('thin', rgb('3F3F3F')) } },
  { id: 'warning', name: 'Warning Text', group: 'data', includes: ['font'], style: { font: { color: rgb('FF0000') } } },

  { id: 'heading1', name: 'Heading 1', group: 'titles', includes: ['font', 'border'], style: { font: { bold: true, size: 15, color: { theme: 3 } }, border: { bottom: { style: 'thick', color: { theme: 4 } } } } },
  { id: 'heading2', name: 'Heading 2', group: 'titles', includes: ['font', 'border'], style: { font: { bold: true, size: 13, color: { theme: 3 } }, border: { bottom: { style: 'thick', color: { theme: 4, tint: 0.499984740745262 } } } } },
  { id: 'heading3', name: 'Heading 3', group: 'titles', includes: ['font', 'border'], style: { font: { bold: true, size: 11, color: { theme: 3 } }, border: { bottom: { style: 'medium', color: { theme: 4, tint: 0.3999755851924192 } } } } },
  { id: 'heading4', name: 'Heading 4', group: 'titles', includes: ['font'], style: { font: { bold: true, size: 11, color: { theme: 3 } } } },
  { id: 'title', name: 'Title', group: 'titles', includes: ['font'], style: { font: { name: HEADING_FONT_NAME, size: 18, color: { theme: 3 } } } },
  { id: 'total', name: 'Total', group: 'titles', includes: ['font', 'border'], style: { font: { bold: true, color: { theme: 1 } }, border: { top: { style: 'thin', color: { theme: 4 } }, bottom: { style: 'double', color: { theme: 4 } } } } },

  ...THEMED_ORDER.map((id) => ACCENT_STYLES.find((preset) => preset.id === id)!),

  { id: 'comma', name: 'Comma', group: 'number', includes: ['number'], style: { numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' } },
  { id: 'comma0', name: 'Comma [0]', group: 'number', includes: ['number'], style: { numFmt: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)' } },
  { id: 'currency', name: 'Currency', group: 'number', includes: ['number'], style: { numFmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' } },
  { id: 'currency0', name: 'Currency [0]', group: 'number', includes: ['number'], style: { numFmt: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)' } },
  { id: 'percent', name: 'Percent', group: 'number', includes: ['number'], style: { numFmt: '0%' } },
]

export function cellStylePreset(id: string) {
  return CELL_STYLES.find((preset) => preset.id === id)
}

/**
 * Applies a cell style like Excel: every category the style includes is replaced, the rest
 * is kept. The cell's font face and size survive unless the style sets them, so workbooks
 * whose Normal font isn't the app default keep their typeface.
 */
export function applyCellStylePreset(style: CellStyle | undefined, preset: CellStylePreset): CellStyle {
  const next: CellStyle = style ? clonePlain(style) : {}
  const source = preset.style
  for (const include of preset.includes) {
    if (include === 'number') {
      if (source.numFmt) next.numFmt = source.numFmt
      else delete next.numFmt
    } else if (include === 'font') {
      const keep: CellFont = {}
      if (next.font?.name) keep.name = next.font.name
      if (next.font?.size) keep.size = next.font.size
      if (next.font?.family !== undefined) keep.family = next.font.family
      if (next.font?.scheme && !source.font?.name) keep.scheme = next.font.scheme
      const font = { ...keep, ...clonePlain(source.font || {}) }
      if (source.font?.name) delete font.scheme
      if (Object.keys(font).length) next.font = font
      else delete next.font
    } else if (include === 'fill') {
      if (source.fill) next.fill = clonePlain(source.fill)
      else delete next.fill
    } else if (include === 'border') {
      if (source.border) next.border = clonePlain(source.border)
      else delete next.border
    } else if (include === 'alignment') {
      if (source.alignment) next.alignment = clonePlain(source.alignment)
      else delete next.alignment
    } else if (include === 'protection') {
      if (source.protection) next.protection = clonePlain(source.protection)
      else delete next.protection
    }
  }
  return next
}

/** Cell-level variant: the style's number format lands on `cell.numFmt` (which wins over style.numFmt). */
export function applyCellStylePresetToCell(cell: CellData | undefined, preset: CellStylePreset): CellData {
  const next: CellData = cell ? { ...cell } : {}
  const style = applyCellStylePreset(next.style, preset)
  if (preset.includes.includes('number')) {
    delete style.numFmt
    if (preset.style.numFmt) next.numFmt = preset.style.numFmt
    else delete next.numFmt
  }
  if (cellHasStyle(style)) next.style = style
  else delete next.style
  return next
}
