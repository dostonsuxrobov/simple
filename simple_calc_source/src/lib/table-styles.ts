import type { SheetTable } from '../spreadsheet-types'

/**
 * Painting for Excel's built-in table styles (TableStyleLight1–21, Medium1–28, Dark1–11).
 * Each family follows Excel's pattern — the first style uses the dark text colour and the
 * next six use accents 1–6 — with header, banding, border, and total-row treatments
 * derived from the workbook theme, so imported tables look the way Excel draws them.
 * Direct cell formatting is layered on top by the grid, as in Excel.
 */

export interface TableCellPaint {
  fill?: string
  color?: string
  bold?: boolean
  borderTop?: string
  borderBottom?: string
  borderLeft?: string
  borderRight?: string
}

export interface ParsedTable {
  table: SheetTable
  top: number
  left: number
  bottom: number
  right: number
}

type Family = 'light-a' | 'light-b' | 'light-c' | 'medium-a' | 'medium-b' | 'medium-c' | 'medium-d' | 'dark-a' | 'dark-b'

const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

function tint(hex: string, amount: number): string {
  const channels = [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16))
  const adjusted = channels.map((channel) => Math.round(amount >= 0 ? channel + (255 - channel) * amount : channel * (1 + amount)))
  return `#${adjusted.map((channel) => Math.max(0, Math.min(255, channel)).toString(16).padStart(2, '0')).join('')}`
}

export function parseTableStyleName(name: string | undefined): { family: Family; colorIndex: number } | null {
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/i.exec(name || 'TableStyleMedium2')
  if (!match) return null
  const kind = match[1].toLowerCase()
  const number = Number(match[2])
  if (kind === 'light') {
    if (number < 1 || number > 21) return null
    return { family: number <= 7 ? 'light-a' : number <= 14 ? 'light-b' : 'light-c', colorIndex: (number - 1) % 7 }
  }
  if (kind === 'medium') {
    if (number < 1 || number > 28) return null
    return { family: number <= 7 ? 'medium-a' : number <= 14 ? 'medium-b' : number <= 21 ? 'medium-c' : 'medium-d', colorIndex: (number - 1) % 7 }
  }
  if (number < 1 || number > 11) return null
  return { family: number <= 7 ? 'dark-a' : 'dark-b', colorIndex: number <= 7 ? (number - 1) % 7 : (number - 8) % 4 }
}

/** Hex (no #) of the style's base colour: dk1 for the first style of a family, else an accent. */
function baseColor(colorIndex: number, theme: readonly string[]): string {
  if (colorIndex === 0) return theme[1] || '000000'
  return theme[3 + colorIndex] || DEFAULT_THEME[3 + colorIndex]
}

const tableCache = new WeakMap<SheetTable[], ParsedTable[]>()

function decode(ref: string) {
  const match = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(ref.trim())
  if (!match) return null
  const column = (label: string) => label.toUpperCase().split('').reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0) - 1
  const top = Number(match[2]) - 1
  const left = column(match[1])
  const bottom = match[4] ? Number(match[4]) - 1 : top
  const right = match[3] ? column(match[3]) : left
  return { top: Math.min(top, bottom), bottom: Math.max(top, bottom), left: Math.min(left, right), right: Math.max(left, right) }
}

export function parsedTables(tables: SheetTable[] | undefined): ParsedTable[] {
  if (!tables?.length) return []
  let parsed = tableCache.get(tables)
  if (!parsed) {
    parsed = tables.flatMap((table) => {
      const bounds = decode(table.ref)
      return bounds ? [{ table, ...bounds }] : []
    })
    tableCache.set(tables, parsed)
  }
  return parsed
}

export function tableAt(tables: SheetTable[] | undefined, row: number, col: number): ParsedTable | null {
  for (const entry of parsedTables(tables)) {
    if (row >= entry.top && row <= entry.bottom && col >= entry.left && col <= entry.right) return entry
  }
  return null
}

/** The table-style paint for one cell (0-based coordinates), or null outside tables. */
export function tableCellPaint(entry: ParsedTable, row: number, col: number, theme: readonly string[] = DEFAULT_THEME): TableCellPaint | null {
  const { table } = entry
  const style = parseTableStyleName(table.style?.theme)
  if (!style) return null
  const base = baseColor(style.colorIndex, theme)
  const dark = `#${base}`
  const header = table.headerRow !== false && row === entry.top
  const total = table.totalsRow && row === entry.bottom
  const bodyTop = entry.top + (table.headerRow !== false ? 1 : 0)
  const bodyIndex = row - bodyTop
  const columnIndex = col - entry.left
  const rowBand = table.style?.showRowStripes !== false && !header && !total && bodyIndex % 2 === 0
  const columnBand = Boolean(table.style?.showColumnStripes) && !header && !total && columnIndex % 2 === 0
  const firstColumn = Boolean(table.style?.showFirstColumn) && col === entry.left
  const lastColumn = Boolean(table.style?.showLastColumn) && col === entry.right
  const thin = (color: string) => `1px solid ${color}`
  const paint: TableCellPaint = {}
  const band = rowBand || columnBand
  switch (style.family) {
    case 'light-a': {
      const line = dark
      if (header) { paint.bold = true; paint.borderBottom = thin(line); paint.borderTop = thin(line) }
      if (total) { paint.bold = true; paint.borderTop = `3px double ${line}` }
      if (row === entry.bottom && !total) paint.borderBottom = thin(line)
      if (band) paint.fill = tint(base, style.colorIndex === 0 ? 0.85 : 0.8)
      if (style.colorIndex > 0) paint.color = tint(base, -0.25)
      break
    }
    case 'light-b': {
      if (header) { paint.fill = dark; paint.color = '#ffffff'; paint.bold = true }
      if (total) { paint.bold = true; paint.borderTop = `3px double ${dark}` }
      if (band) { paint.borderTop = thin(dark); paint.borderBottom = thin(dark) }
      if (col === entry.left) paint.borderLeft = thin(dark)
      if (col === entry.right) paint.borderRight = thin(dark)
      if (row === entry.bottom) paint.borderBottom = thin(dark)
      break
    }
    case 'light-c': {
      const line = dark
      paint.borderTop = thin(line); paint.borderBottom = thin(line); paint.borderLeft = thin(line); paint.borderRight = thin(line)
      if (header) { paint.bold = true; paint.borderBottom = `2px solid ${line}` }
      if (total) { paint.bold = true; paint.borderTop = `3px double ${line}` }
      if (band) paint.fill = tint(base, style.colorIndex === 0 ? 0.85 : 0.8)
      break
    }
    case 'medium-a': {
      const line = tint(base, 0.4)
      if (header) { paint.fill = dark; paint.color = '#ffffff'; paint.bold = true }
      else if (total) { paint.bold = true; paint.borderTop = `3px double ${dark}` }
      else { paint.borderTop = thin(line); paint.borderBottom = thin(line) }
      if (band) paint.fill = tint(base, style.colorIndex === 0 ? 0.85 : 0.8)
      if (col === entry.left) paint.borderLeft = thin(line)
      if (col === entry.right) paint.borderRight = thin(line)
      break
    }
    case 'medium-b': {
      paint.borderTop = thin('#ffffff'); paint.borderBottom = thin('#ffffff'); paint.borderLeft = thin('#ffffff'); paint.borderRight = thin('#ffffff')
      if (header) { paint.fill = dark; paint.color = '#ffffff'; paint.bold = true; paint.borderBottom = '3px solid #ffffff' }
      else if (total) { paint.fill = dark; paint.color = '#ffffff'; paint.bold = true; paint.borderTop = '3px solid #ffffff' }
      else paint.fill = band ? tint(base, style.colorIndex === 0 ? 0.65 : 0.6) : tint(base, style.colorIndex === 0 ? 0.85 : 0.8)
      if ((firstColumn || lastColumn) && !header && !total) { paint.fill = dark; paint.color = '#ffffff'; paint.bold = true }
      break
    }
    case 'medium-c': {
      const line = dark
      if (header) { paint.fill = '#000000'; paint.color = '#ffffff'; paint.bold = true }
      if (total) { paint.bold = true; paint.borderTop = `3px double ${line}` }
      if (band) paint.fill = tint(style.colorIndex === 0 ? '000000' : base, style.colorIndex === 0 ? 0.85 : 0.8)
      paint.borderBottom = paint.borderBottom || thin(tint(base, 0.4))
      if (row === entry.bottom) paint.borderBottom = thin(line)
      break
    }
    case 'medium-d': {
      const line = tint(base, 0.4)
      paint.borderTop = thin(line); paint.borderBottom = thin(line); paint.borderLeft = thin(line); paint.borderRight = thin(line)
      paint.fill = band ? tint(base, style.colorIndex === 0 ? 0.65 : 0.6) : tint(base, style.colorIndex === 0 ? 0.85 : 0.8)
      if (header) { paint.bold = true; paint.fill = tint(base, style.colorIndex === 0 ? 0.85 : 0.8) }
      if (total) { paint.bold = true; paint.borderTop = `3px double ${dark}` }
      break
    }
    case 'dark-a': {
      const shade = tint(base, style.colorIndex === 0 ? 0.25 : -0.25)
      paint.color = '#ffffff'
      paint.fill = band ? tint(base, style.colorIndex === 0 ? 0.15 : -0.5) : shade
      if (header) { paint.fill = '#000000'; paint.bold = true; paint.borderBottom = '2px solid #ffffff' }
      if (total) { paint.bold = true; paint.fill = tint(base, -0.5); paint.borderTop = '2px solid #ffffff' }
      break
    }
    case 'dark-b': {
      const accents = [[1, 2], [3, 4], [5, 6], [7, 8]][style.colorIndex] || [1, 2]
      const first = style.colorIndex === 0 ? tint(theme[1] || '000000', 0.25) : `#${theme[3 + accents[0]] || '4472C4'}`
      paint.fill = band ? tint(first.replace('#', ''), 0.6) : tint(first.replace('#', ''), 0.8)
      if (header) { paint.fill = `#${theme[3 + accents[1]] || 'ED7D31'}`; paint.color = '#ffffff'; paint.bold = true }
      if (total) { paint.bold = true; paint.borderTop = `3px double ${first}` }
      break
    }
  }
  if ((firstColumn || lastColumn) && !paint.bold) paint.bold = true
  return paint
}

/** The built-in style names in gallery order. */
export const TABLE_STYLE_GALLERY: string[] = [
  ...Array.from({ length: 21 }, (_, index) => `TableStyleLight${index + 1}`),
  ...Array.from({ length: 28 }, (_, index) => `TableStyleMedium${index + 1}`),
  ...Array.from({ length: 11 }, (_, index) => `TableStyleDark${index + 1}`),
]
