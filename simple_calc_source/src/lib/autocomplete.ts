/**
 * Excel's AutoComplete for cell values and its Alt+Down "Pick From Drop-down List": the text
 * entries of the contiguous block of cells above and below the active cell in its column.
 * Numbers, dates, logicals and formulas never take part (Excel completes text only), and a
 * blank cell ends the block, as in Excel.
 */
import type { CellData, SheetData } from '../spreadsheet-types'

/** How far the block is followed in each direction (Excel stops at the block edge). */
const MAX_SCAN = 1_000

function columnName(index: number) {
  let value = index + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function occupied(cell: CellData | undefined) {
  return Boolean(cell && (cell.formula || (cell.value !== null && cell.value !== undefined && cell.value !== '')))
}

function entryText(cell: CellData | undefined): string | null {
  if (!cell || cell.formula || typeof cell.value !== 'string') return null
  const text = cell.value
  return text.trim() ? text : null
}

/**
 * Text entries of the column block around (row, col), the edited cell itself excluded,
 * nearest first and without case-insensitive duplicates.
 */
export function columnTextEntries(cells: SheetData['cells'], row: number, col: number, limit = MAX_SCAN): string[] {
  const label = columnName(col)
  const seen = new Set<string>()
  const output: string[] = []
  const add = (cell: CellData | undefined) => {
    const text = entryText(cell)
    if (text === null) return
    const key = text.toLocaleLowerCase()
    if (seen.has(key)) return
    seen.add(key)
    output.push(text)
  }
  for (let current = row - 1, steps = 0; current >= 0 && steps < limit; current -= 1, steps += 1) {
    const cell = cells[`${label}${current + 1}`]
    if (!occupied(cell)) break
    add(cell)
  }
  for (let current = row + 1, steps = 0; steps < limit; current += 1, steps += 1) {
    const cell = cells[`${label}${current + 1}`]
    if (!occupied(cell)) break
    add(cell)
  }
  return output
}

/**
 * The single entry that completes `typed` (case-insensitive prefix), or null when no entry or
 * more than one distinct entry starts with it (Excel completes only an unambiguous prefix).
 */
export function autoCompleteMatch(typed: string, entries: readonly string[]): string | null {
  if (!typed || !typed.trim() || /[\r\n]/.test(typed)) return null
  const prefix = typed.toLocaleLowerCase()
  let match: string | null = null
  let matchKey = ''
  for (const entry of entries) {
    const key = entry.toLocaleLowerCase()
    if (!key.startsWith(prefix)) continue
    if (match !== null && key !== matchKey) return null
    if (match === null) {
      match = entry
      matchKey = key
    }
  }
  return match !== null && match.length > typed.length ? match : null
}

/**
 * What the editor shows while AutoComplete suggests `match`: the characters typed so far, then
 * the rest of the entry (selected, so the next keystroke replaces it).
 */
export function autoCompleteDraft(typed: string, match: string): string {
  return typed + match.slice(typed.length)
}

/**
 * Accepting a suggestion stores the entry as it is written in the column (Excel matches the
 * case of the existing entry): "new y" + "ork" becomes "New York".
 */
export function acceptedAutoComplete(draft: string, suggestion: { typed: string; match: string } | null): string {
  if (!suggestion) return draft
  return draft === autoCompleteDraft(suggestion.typed, suggestion.match) ? suggestion.match : draft
}

/** Alt+Down: the column block's distinct text entries, sorted as Excel's pick list. */
export function pickListEntries(cells: SheetData['cells'], row: number, col: number, limit = MAX_SCAN): string[] {
  return columnTextEntries(cells, row, col, limit)
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }))
}
