/**
 * Page setup a sheet carries into its file and every printout: the print area, rows and
 * columns repeated on each page, the header and footer, and manual page breaks. The print
 * dialog edits these through SpreadsheetPageSetupPatch; applyPageSetupPatch writes them in
 * the model shape the XLSX serializer already saves (pageSetup.printArea / printTitlesRow /
 * printTitlesColumn as _xlnm.Print_Area and _xlnm.Print_Titles, headerFooter, rowBreaks).
 */
import type { SheetData } from '../spreadsheet-types'

export interface PageSetupBounds {
  top: number
  bottom: number
  left: number
  right: number
}

/** What the print dialog shows for the active sheet. */
export interface SpreadsheetPageSetupState {
  sheetName: string
  /** Print area in A1 form ("A1:D20", several areas comma-separated), '' for the whole sheet. */
  printArea: string
  /** Rows repeated at the top of each page ("1:2"), ''. */
  printTitlesRow: string
  /** Columns repeated at the left of each page ("A:B"), ''. */
  printTitlesColumn: string
  frozenRows: number
  frozenColumns: number
  /** The current selection, 0-based. */
  selection: PageSetupBounds
  /** OOXML header/footer codes for odd (normal) pages. */
  oddHeader: string
  oddFooter: string
  /** The first page or even pages use their own header/footer (kept as they are). */
  differentFirst: boolean
  differentOddEven: boolean
  /** 1-based first row of each page that starts at a manual page break, ascending. */
  rowBreaks: number[]
}

/** A change to the active sheet's page setup; null clears a setting. */
export interface SpreadsheetPageSetupPatch {
  printArea?: string | null
  printTitlesRow?: string | null
  printTitlesColumn?: string | null
  oddHeader?: string
  oddFooter?: string
  /** Replaces all manual row breaks: 1-based first row of each new page. */
  rowBreaks?: number[]
}

const MAX_ROWS = 1_048_576
const MAX_COLUMNS = 16_384

export function columnLabel(index: number): string {
  let value = Math.max(0, Math.trunc(index)) + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function columnNumber(label: string): number {
  let value = 0
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value
}

/** "B2:D10" (or "B2" for one cell) for 0-based bounds. */
export function rangeLabel(bounds: PageSetupBounds): string {
  const start = `${columnLabel(bounds.left)}${bounds.top + 1}`
  const end = `${columnLabel(bounds.right)}${bounds.bottom + 1}`
  return start === end ? start : `${start}:${end}`
}

/**
 * Rows to repeat typed by the user ("1", "1:2", "$1:$3"): "first:last", '' for none, or
 * null when the text is not a row range.
 */
export function normalizeTitleRows(text: string): string | null {
  const value = text.trim().replace(/\$/g, '')
  if (!value) return ''
  const match = /^(\d{1,7})(?::(\d{1,7}))?$/.exec(value)
  if (!match) return null
  const first = Number(match[1]), last = Number(match[2] ?? match[1])
  if (first < 1 || last < 1 || first > MAX_ROWS || last > MAX_ROWS) return null
  return `${Math.min(first, last)}:${Math.max(first, last)}`
}

/** Columns to repeat ("A", "A:B", "$A:$C"): "A:B", '' for none, or null when invalid. */
export function normalizeTitleColumns(text: string): string | null {
  const value = text.trim().replace(/\$/g, '').toUpperCase()
  if (!value) return ''
  const match = /^([A-Z]{1,3})(?::([A-Z]{1,3}))?$/.exec(value)
  if (!match) return null
  const first = columnNumber(match[1]), last = columnNumber(match[2] ?? match[1])
  if (first > MAX_COLUMNS || last > MAX_COLUMNS) return null
  return `${columnLabel(Math.min(first, last) - 1)}:${columnLabel(Math.max(first, last) - 1)}`
}

function rowBreakList(sheet: SheetData): number[] {
  const rows = new Set<number>()
  for (const entry of Array.isArray(sheet.rowBreaks) ? sheet.rowBreaks : []) {
    // ExcelJS shape: { id: 1-based row above the break, max, man }.
    const id = Math.trunc(Number(entry && typeof entry === 'object' ? (entry as { id?: unknown }).id : entry))
    if (Number.isFinite(id) && id >= 1 && id < MAX_ROWS) rows.add(id + 1)
  }
  return [...rows].sort((a, b) => a - b)
}

export function pageSetupStateFor(sheet: SheetData, selection: PageSetupBounds): SpreadsheetPageSetupState {
  const setup = (sheet.pageSetup || {}) as Record<string, unknown>
  const headerFooter = (sheet.headerFooter || {}) as Record<string, unknown>
  const text = (value: unknown) => (typeof value === 'string' ? value : '')
  const printArea = text(setup.printArea).split('&&').map((area) => area.trim()).filter(Boolean).join(', ')
  return {
    sheetName: sheet.name,
    printArea,
    printTitlesRow: text(setup.printTitlesRow).replace(/\$/g, ''),
    printTitlesColumn: text(setup.printTitlesColumn).replace(/\$/g, ''),
    frozenRows: Math.max(0, Math.trunc(Number(sheet.frozen?.rows) || 0)),
    frozenColumns: Math.max(0, Math.trunc(Number(sheet.frozen?.columns) || 0)),
    selection: { ...selection },
    oddHeader: text(headerFooter.oddHeader),
    oddFooter: text(headerFooter.oddFooter),
    differentFirst: headerFooter.differentFirst === true,
    differentOddEven: headerFooter.differentOddEven === true,
    rowBreaks: rowBreakList(sheet),
  }
}

/**
 * Apply a page-setup change to a sheet (an immer draft inside mutateWorkbook), in the model
 * shape workbooks.cjs saves. Invalid values are ignored rather than stored.
 */
export function applyPageSetupPatch(sheet: SheetData, patch: SpreadsheetPageSetupPatch): void {
  const touchesSetup = 'printArea' in patch || 'printTitlesRow' in patch || 'printTitlesColumn' in patch
  if (touchesSetup) {
    const setup = { ...(sheet.pageSetup || {}) } as Record<string, unknown>
    if ('printArea' in patch) {
      const areas = String(patch.printArea || '').split(/[,&]+/).map((area) => area.trim().replace(/\$/g, '').toUpperCase()).filter((area) => /^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$/.test(area))
      // A new area replaces the whole-column form remembered from the opened file.
      delete setup.printAreaWhole
      if (areas.length) setup.printArea = areas.join('&&')
      else delete setup.printArea
    }
    if ('printTitlesRow' in patch) {
      const rows = patch.printTitlesRow ? normalizeTitleRows(patch.printTitlesRow) : ''
      if (rows) setup.printTitlesRow = rows
      else if (rows === '') delete setup.printTitlesRow
    }
    if ('printTitlesColumn' in patch) {
      const columns = patch.printTitlesColumn ? normalizeTitleColumns(patch.printTitlesColumn) : ''
      if (columns) setup.printTitlesColumn = columns
      else if (columns === '') delete setup.printTitlesColumn
    }
    sheet.pageSetup = setup
  }
  if ('oddHeader' in patch || 'oddFooter' in patch) {
    const headerFooter = { ...(sheet.headerFooter || {}) } as Record<string, unknown>
    for (const key of ['oddHeader', 'oddFooter'] as const) {
      if (!(key in patch)) continue
      const value = String(patch[key] ?? '').slice(0, 2_000)
      if (value) headerFooter[key] = value
      else delete headerFooter[key]
    }
    sheet.headerFooter = headerFooter
  }
  if (patch.rowBreaks) {
    const rows = [...new Set(patch.rowBreaks.map((row) => Math.trunc(Number(row))).filter((row) => Number.isFinite(row) && row >= 2 && row <= MAX_ROWS))].sort((a, b) => a - b).slice(0, 1_026)
    // ExcelJS's row.addPageBreak shape; id is the last row before the break.
    sheet.rowBreaks = rows.map((row) => ({ id: row - 1, max: MAX_COLUMNS - 1, man: 1 }))
  }
}

// ---------------------------------------------------------------------------------------------
// Header and footer text
// ---------------------------------------------------------------------------------------------

export type HeaderFooterSection = 'left' | 'center' | 'right'

export const HEADER_FOOTER_SECTIONS: readonly HeaderFooterSection[] = ['left', 'center', 'right']

/** Excel's field tokens as its header editor shows them, with their OOXML codes. */
export const HEADER_FOOTER_FIELDS: ReadonlyArray<{ token: string; code: string; label: string }> = [
  { token: '&[Page]', code: '&P', label: 'Page' },
  { token: '&[Pages]', code: '&N', label: 'Pages' },
  { token: '&[Tab]', code: '&A', label: 'Sheet' },
  { token: '&[File]', code: '&F', label: 'File' },
  { token: '&[Date]', code: '&D', label: 'Date' },
  { token: '&[Time]', code: '&T', label: 'Time' },
]

const FIELD_TOKENS: Record<string, string> = { P: '&[Page]', N: '&[Pages]', A: '&[Tab]', F: '&[File]', D: '&[Date]', T: '&[Time]', Z: '&[Path]', G: '&[Picture]' }
const TOKEN_CODES: Record<string, string> = { page: '&P', pages: '&N', tab: '&A', sheet: '&A', file: '&F', date: '&D', time: '&T', path: '&Z', picture: '&G' }
const FORMAT_TOGGLES = new Set(['B', 'I', 'U', 'E', 'S', 'X', 'Y', 'O', 'H'])

/** Length of the formatting code at `index` ("&B", "&14", "&\"Arial,Bold\"", "&KFF0000"), or 0. */
function formattingCodeLength(text: string, index: number): number {
  if (text[index] !== '&') return 0
  const code = text[index + 1]
  if (code === undefined) return 0
  if (FORMAT_TOGGLES.has(code)) return 2
  if (code === '"') {
    const end = text.indexOf('"', index + 2)
    return end < 0 ? 0 : end - index + 1
  }
  const size = /^&\d{1,3}/.exec(text.slice(index))
  if (size) return size[0].length
  const color = /^&K(?:[0-9A-Fa-f]{6}|\d{2}[+-]\d{3})/.exec(text.slice(index))
  return color ? color[0].length : 0
}

/** Split an OOXML header/footer into its raw left/center/right sections (text without &L/&C/&R is centred). */
export function splitHeaderFooter(code: string): Record<HeaderFooterSection, string> {
  const sections: Record<HeaderFooterSection, string> = { left: '', center: '', right: '' }
  let section: HeaderFooterSection = 'center'
  const text = String(code || '')
  for (let index = 0; index < text.length;) {
    if (text[index] === '&' && index + 1 < text.length) {
      const next = text[index + 1]
      if (next === 'L' || next === 'C' || next === 'R') {
        section = next === 'L' ? 'left' : next === 'C' ? 'center' : 'right'
        index += 2
        continue
      }
      const length = next === '&' ? 2 : Math.max(2, formattingCodeLength(text, index))
      sections[section] += text.slice(index, index + length)
      index += length
      continue
    }
    sections[section] += text[index]
    index += 1
  }
  return sections
}

export function joinHeaderFooter(sections: Record<HeaderFooterSection, string>): string {
  return `${sections.left ? `&L${sections.left}` : ''}${sections.center ? `&C${sections.center}` : ''}${sections.right ? `&R${sections.right}` : ''}`
}

/** A raw section as editable text: fields become Excel tokens (&[Page]), formatting codes are hidden. */
export function headerSectionText(raw: string): string {
  let result = ''
  for (let index = 0; index < raw.length;) {
    if (raw[index] !== '&' || index + 1 >= raw.length) { result += raw[index]; index += 1; continue }
    const next = raw[index + 1]
    if (next === '&') { result += '&'; index += 2; continue }
    if (Object.hasOwn(FIELD_TOKENS, next)) { result += FIELD_TOKENS[next]; index += 2; continue }
    const length = formattingCodeLength(raw, index)
    if (length) { index += length; continue }
    result += raw.slice(index, index + 2)
    index += 2
  }
  return result
}

/**
 * Editable text back to OOXML. Excel tokens become field codes and any other "&" a literal
 * ampersand; the section's leading font, size and colour codes from `previousRaw` are kept.
 */
export function headerSectionCode(text: string, previousRaw = ''): string {
  let prefix = ''
  for (let index = 0; index < previousRaw.length;) {
    const length = formattingCodeLength(previousRaw, index)
    if (!length) break
    prefix += previousRaw.slice(index, index + length)
    index += length
  }
  let body = ''
  const value = String(text || '').replace(/\r\n?/g, '\n')
  for (let index = 0; index < value.length;) {
    if (value[index] !== '&') { body += value[index]; index += 1; continue }
    const token = /^&\[([A-Za-z]+)\]/.exec(value.slice(index))
    const code = token ? TOKEN_CODES[token[1].toLowerCase()] : undefined
    if (token && code) { body += code; index += token[0].length; continue }
    body += '&&'
    index += 1
  }
  // A font-size code swallows following digits ("&14" + "2026"): separate them.
  if (/&\d{1,3}$/.test(prefix) && /^\d/.test(body)) body = ` ${body}`
  return body ? prefix + body : ''
}

/** Editable left/center/right texts of an OOXML header or footer. */
export function headerFooterTexts(code: string): Record<HeaderFooterSection, string> {
  const raw = splitHeaderFooter(code)
  return { left: headerSectionText(raw.left), center: headerSectionText(raw.center), right: headerSectionText(raw.right) }
}

/**
 * The OOXML for edited section texts. Sections whose text did not change keep their
 * original codes byte for byte (formatting in the middle of a line survives).
 */
export function headerFooterCode(original: string, texts: Record<HeaderFooterSection, string>): string {
  const raw = splitHeaderFooter(original)
  const sections = { ...raw }
  for (const section of HEADER_FOOTER_SECTIONS) {
    if (texts[section] === headerSectionText(raw[section])) continue
    sections[section] = headerSectionCode(texts[section], raw[section])
  }
  const next = joinHeaderFooter(sections)
  // Untouched input stays exactly as the file had it (section order, repeated sections).
  return HEADER_FOOTER_SECTIONS.every((section) => sections[section] === raw[section]) ? original : next
}
