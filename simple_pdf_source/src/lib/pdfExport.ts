import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { ExportTextPage } from '../types'

type ExportImageFormat = 'png' | 'jpeg' | 'webp'

// ---------------------------------------------------------------------------
// The layout model shared by the Word, web page, Markdown and text exports.
// pdf.js text items are rebuilt into lines (baselines and its end-of-line
// markers), lines into paragraphs, aligned columns into simple tables, and
// headings are ranked by font size across the exported pages. Pictures come
// from the operator list. The model is plain data, so it crosses IPC as is.
// ---------------------------------------------------------------------------

export interface ExportRun {
  text: string
  bold?: boolean
  italic?: boolean
  /** Font size in points. */
  size?: number
  /** Font family, when the PDF names a recognisable one. */
  font?: string
  superscript?: boolean
}

export interface ExportLine {
  runs: ExportRun[]
  /** The line ends with a forced break (the next word would have fitted). */
  hardBreak?: boolean
}

export type ExportBlock =
  | { type: 'heading'; level: 1 | 2 | 3; lines: ExportLine[]; align?: 'center' | 'right' }
  | { type: 'paragraph'; lines: ExportLine[]; align?: 'center' | 'right'; list?: 'bullet' }
  /** rows, then cells, then the cell's lines; widths in points. */
  | { type: 'table'; rows: ExportLine[][][]; widths: number[] }
  /** width and height are the drawn size in points. */
  | { type: 'image'; data: Uint8Array; mime: 'image/png' | 'image/jpeg'; width: number; height: number }

export interface StructuredTextPage extends ExportTextPage {
  /** Page size in points, as displayed. */
  width: number
  height: number
  blocks: ExportBlock[]
}

export interface ExtractTextOptions {
  /**
   * Include pictures (default true). Without them the operator list is not
   * read, which makes plain-text exports of long documents much faster; bold
   * and italic then come only from fonts already loaded by rendering.
   */
  images?: boolean
  /** Called after each page, for a progress message. */
  onProgress?: (done: number, total: number) => void
}

interface TextContentItem {
  str: string
  hasEOL?: boolean
  transform?: number[]
  width?: number
  height?: number
  fontName?: string
}

interface FontInfo {
  bold: boolean
  italic: boolean
  family: string
}

interface Fragment {
  text: string
  u: number
  v: number
  right: number
  size: number
  bold: boolean
  italic: boolean
  font: string
  superscript?: boolean
}

interface Segment {
  u: number
  right: number
  fragments: Fragment[]
}

interface Line {
  segments: Segment[]
  u: number
  right: number
  v: number
  size: number
}

interface DraftText {
  type: 'text'
  lines: Line[]
  top: number
  size: number
  bold: boolean
  chars: number
  list?: 'bullet'
  align?: 'center' | 'right'
  hardBreaks: boolean[]
}

interface DraftTable {
  type: 'table'
  rows: Line[][][]
  widths: number[]
  top: number
}

interface DraftImage {
  type: 'image'
  data: Uint8Array
  mime: 'image/png' | 'image/jpeg'
  width: number
  height: number
  top: number
}

type Draft = DraftText | DraftTable | DraftImage

/** One page of the model before headings are ranked across the document. */
export interface PageLayoutDraft {
  pageNumber: number
  width: number
  height: number
  drafts: Draft[]
  /** Characters per font size (rounded to 0.5 pt), to find the body size. */
  sizes: Record<string, number>
}

export interface PlacedImage {
  data: Uint8Array
  mime: 'image/png' | 'image/jpeg'
  /** The image's transformation matrix in PDF user space. */
  matrix: number[]
}

const BULLET_PATTERN = /^\s*([•◦▪▫●○■□‣⁃∙·\uF0B7\uF0A7\uF076\uF0D8\uF0FC\uF06C\uF06E\uF0A8])(\s+|$)/
const PAGE_NUMBER_PATTERN = /^\s*(page\s*)?\d{1,4}(\s*(of|\/)\s*\d{1,4})?\s*$/i

function isTextItem(value: unknown): value is TextContentItem {
  return Boolean(value && typeof value === 'object' && typeof (value as TextContentItem).str === 'string')
}

function cleanText(value: string) {
  // PDF text streams occasionally carry NULs and other C0 controls.
  return value.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '')
}

function itemSize(item: TextContentItem) {
  const [a = 0, b = 0, c = 0, d = 0] = item.transform || []
  return Math.hypot(c, d) || Math.hypot(a, b) || Number(item.height) || 0
}

/** Page text with its line breaks (pdf.js end-of-line markers and baseline changes). */
export function textContentToPlainText(items: unknown[]) {
  let output = ''
  let lastY: number | null = null
  let lastRight: number | null = null
  for (const rawItem of items) {
    if (!isTextItem(rawItem)) continue
    const value = cleanText(rawItem.str)
    // pdf.js 4 reports a line end as an empty item with hasEOL, so the break
    // is applied even when the item carries no text.
    if (value) {
      const transform = rawItem.transform
      const size = itemSize(rawItem) || 10
      const y = transform ? transform[5] : null
      const x = transform ? transform[4] : null
      if (output && !output.endsWith('\n') && y !== null && lastY !== null && Math.abs(y - lastY) > size * 0.5) output = `${output.trimEnd()}\n`
      const previous = output.at(-1) || ''
      if (output && !/\s/.test(previous) && !/^\s|^[,.;:!?%)\]}]/.test(value)) {
        const gap = x !== null && lastRight !== null ? x - lastRight : null
        if (gap === null || gap > size * 0.12) output += ' '
      }
      output += value
      if (y !== null) lastY = y
      lastRight = x !== null && Number.isFinite(rawItem.width) ? x + Number(rawItem.width) : null
    }
    if (rawItem.hasEOL) output = `${output.trimEnd()}\n`
  }
  return output
    .split('\n')
    .map((line) => line.trimEnd())
    .join('\n')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim()
}

// ---------------------------------------------------------------------------
// Fonts
// ---------------------------------------------------------------------------

const KNOWN_FAMILIES: Record<string, string> = {
  arial: 'Arial', arialnarrow: 'Arial Narrow', helvetica: 'Arial', helveticaneue: 'Helvetica Neue', timesnewroman: 'Times New Roman',
  times: 'Times New Roman', timesroman: 'Times New Roman', couriernew: 'Courier New', courier: 'Courier New', calibri: 'Calibri',
  calibrilight: 'Calibri Light', cambria: 'Cambria', cambriamath: 'Cambria Math', georgia: 'Georgia', verdana: 'Verdana',
  tahoma: 'Tahoma', segoeui: 'Segoe UI', garamond: 'Garamond', bookantiqua: 'Book Antiqua', centurygothic: 'Century Gothic',
  trebuchetms: 'Trebuchet MS', consolas: 'Consolas', aptos: 'Aptos', palatinolinotype: 'Palatino Linotype', palatino: 'Palatino Linotype',
  lucidaconsole: 'Lucida Console', candara: 'Candara', constantia: 'Constantia', corbel: 'Corbel', franklingothic: 'Franklin Gothic',
  symbol: 'Symbol', wingdings: 'Wingdings', notosans: 'Noto Sans', notoserif: 'Noto Serif', roboto: 'Roboto', opensans: 'Open Sans',
}

/** Bold, italic and a family name from a PDF font name such as "ABCDEF+Georgia-BoldItalic". */
export function fontInfoFromName(rawName: string): FontInfo {
  const name = String(rawName || '').replace(/^[A-Z]{6}\+/, '')
  // Recognised (OCR) text is drawn with GlyphLessFont, which names no typeface
  // and no style: the exported text keeps the document's default font.
  if (/^GlyphLessFont$/i.test(name)) return { bold: false, italic: false, family: '' }
  const style = name.split(/[-,]/).slice(1).join('-') || name
  const bold = /bold|black|heavy|semibold|demi|extrabold|ultrabold/i.test(style) || /(bold|black|heavy)$/i.test(name.split(/[-,]/)[0])
  const italic = /italic|oblique|slanted|(^|[-,])it$/i.test(style) || /italic|oblique/i.test(name.split(/[-,]/)[0])
  let base = name.split(/[-,]/)[0].replace(/(PSMT|PS|MT|Std|Pro|LT|Regular|Bold|Italic)+$/g, '')
  const key = base.toLowerCase().replace(/[^a-z]/g, '')
  let family = KNOWN_FAMILIES[key] || ''
  if (!family && base && !/^(g_d\d|f\d+$|t\d+$|font\d*$|cid)/i.test(base) && /[a-z]{3}/i.test(base)) {
    base = base.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/_/g, ' ').trim()
    family = base.length <= 40 ? base : ''
  }
  return { bold, italic, family }
}

// ---------------------------------------------------------------------------
// Lines
// ---------------------------------------------------------------------------

interface Frame {
  cos: number
  sin: number
}

function dominantFrame(items: TextContentItem[]): Frame {
  const weights = new Map<number, number>()
  for (const item of items) {
    if (!item.str || !item.transform) continue
    const [a, b] = item.transform
    const angle = Math.round(Math.atan2(b, a) * 180 / Math.PI)
    const snapped = ((Math.round(angle / 90) * 90) % 360 + 360) % 360
    const key = Math.abs(angle - Math.round(angle / 90) * 90) <= 3 ? snapped : ((angle % 360) + 360) % 360
    weights.set(key, (weights.get(key) || 0) + item.str.length)
  }
  const angle = [...weights].sort((left, right) => right[1] - left[1])[0]?.[0] ?? 0
  return { cos: Math.cos(angle * Math.PI / 180), sin: Math.sin(angle * Math.PI / 180) }
}

function inFrame(item: TextContentItem, frame: Frame) {
  const [a = 1, b = 0] = item.transform || []
  const length = Math.hypot(a, b) || 1
  return Math.abs(a / length * frame.cos + b / length * frame.sin - 1) < 0.002
}

function fragmentOf(item: TextContentItem, frame: Frame, fonts: (fontName: string) => FontInfo): Fragment | null {
  const text = cleanText(item.str)
  if (!text || !item.transform) return null
  const [, , , , x = 0, y = 0] = item.transform
  const size = itemSize(item)
  if (!(size > 0)) return null
  const u = x * frame.cos + y * frame.sin
  const v = -x * frame.sin + y * frame.cos
  const width = Math.max(0, Number(item.width) || 0)
  const info = fonts(item.fontName || '')
  return { text, u, v, right: u + width, size, bold: info.bold, italic: info.italic, font: info.family }
}

function lineText(line: Line, separator = '\t') {
  return line.segments.map((segment) => segment.fragments.map((fragment) => fragment.text).join('')).join(separator)
}

function lineChars(line: Line) {
  return line.segments.reduce((sum, segment) => sum + segment.fragments.reduce((count, fragment) => count + fragment.text.trim().length, 0), 0)
}

function finishLine(line: Line) {
  const weights = new Map<number, number>()
  for (const segment of line.segments) {
    for (const fragment of segment.fragments) {
      if (fragment.superscript) continue
      const key = Math.round(fragment.size * 2) / 2
      weights.set(key, (weights.get(key) || 0) + Math.max(1, fragment.text.trim().length))
    }
  }
  line.size = [...weights].sort((left, right) => right[1] - left[1])[0]?.[0] ?? line.size
  line.u = line.segments[0].u
  line.right = Math.max(...line.segments.map((segment) => segment.right))
  return line
}

/** Lines in content order. A big horizontal gap starts a new segment (a column). */
function buildLines(items: TextContentItem[], frame: Frame, fonts: (fontName: string) => FontInfo) {
  const lines: Line[] = []
  let line: Line | null = null
  let pendingBreak = false
  let pendingGap = false
  for (const item of items) {
    if (!item.str) {
      if (item.hasEOL) pendingBreak = true
      continue
    }
    const fragment = fragmentOf(item, frame, fonts)
    if (!fragment) {
      if (item.hasEOL) pendingBreak = true
      continue
    }
    const blank = !fragment.text.trim()
    if (line) {
      const reference = Math.max(fragment.size, line.size)
      const sameBaseline = Math.abs(fragment.v - line.v) <= reference * 0.4
      const smaller = fragment.size < line.size * 0.85
      const raised = smaller && fragment.v > line.v - line.size * 0.35 && fragment.v < line.v + line.size * 0.75
      const continues = fragment.u >= line.right - reference * 1.2
      const joins = (sameBaseline || raised) && continues && (!pendingBreak || (raised && !sameBaseline))
      if (!joins) {
        lines.push(finishLine(line))
        line = null
      } else if (blank) {
        // pdf.js marks wide gaps (table cells, tab stops) with a wide space.
        if (Number(item.width) > reference * 2.2) pendingGap = true
        else {
          const segment = line.segments[line.segments.length - 1]
          const last = segment.fragments[segment.fragments.length - 1]
          if (!/\s$/.test(last.text)) segment.fragments.push({ ...last, text: ' ', u: fragment.u, right: Math.max(fragment.right, last.right), superscript: false })
          segment.right = Math.max(segment.right, fragment.right)
          line.right = Math.max(line.right, fragment.right)
        }
        pendingBreak = Boolean(item.hasEOL)
        continue
      } else {
        const gap = fragment.u - line.right
        fragment.superscript = smaller && fragment.v > line.v + line.size * 0.15
        const segment = line.segments[line.segments.length - 1]
        if (pendingGap || gap > reference * 2.2) {
          line.segments.push({ u: fragment.u, right: fragment.right, fragments: [fragment] })
        } else {
          const last = segment.fragments[segment.fragments.length - 1]
          if (gap > reference * 0.12 && !/\s$/.test(last.text) && !/^\s/.test(fragment.text)) segment.fragments.push({ ...last, text: ' ', superscript: false })
          segment.fragments.push(fragment)
          segment.right = Math.max(segment.right, fragment.right)
        }
        line.right = Math.max(line.right, fragment.right)
        pendingGap = false
        pendingBreak = Boolean(item.hasEOL)
        continue
      }
    }
    pendingGap = false
    pendingBreak = Boolean(item.hasEOL)
    if (blank) continue
    line = { segments: [{ u: fragment.u, right: fragment.right, fragments: [fragment] }], u: fragment.u, right: fragment.right, v: fragment.v, size: fragment.size }
  }
  if (line) lines.push(finishLine(line))
  for (const finished of lines) {
    for (const segment of finished.segments) {
      // Trim the spaces pdf.js leaves at segment edges.
      while (segment.fragments.length > 1 && !segment.fragments[segment.fragments.length - 1].text.trim()) segment.fragments.pop()
      while (segment.fragments.length > 1 && !segment.fragments[0].text.trim()) segment.fragments.shift()
    }
  }
  return lines.filter((candidate) => lineChars(candidate) > 0)
}

// ---------------------------------------------------------------------------
// Tables, paragraphs, pictures
// ---------------------------------------------------------------------------

function modalPitch(lines: Line[]) {
  const counts = new Map<string, Map<number, number>>()
  for (let index = 1; index < lines.length; index += 1) {
    const previous = lines[index - 1]
    const line = lines[index]
    if (Math.abs(line.size - previous.size) > 0.5) continue
    const pitch = previous.v - line.v
    if (!(pitch > 0 && pitch < line.size * 3)) continue
    const key = String(Math.round(line.size * 2) / 2)
    const bucket = counts.get(key) || new Map<number, number>()
    const rounded = Math.round(pitch * 2) / 2
    bucket.set(rounded, (bucket.get(rounded) || 0) + 1)
    counts.set(key, bucket)
  }
  return (size: number) => {
    const bucket = counts.get(String(Math.round(size * 2) / 2))
    const best = bucket ? [...bucket].sort((left, right) => right[1] - left[1] || left[0] - right[0])[0]?.[0] : undefined
    return best || size * 1.2
  }
}

interface TableCandidate {
  start: number
  end: number
  columns: number[]
  rows: { v: number, cells: Line[][] }[]
}

/** Column index of a position, or -1. */
function columnAt(columns: number[], u: number, tolerance: number) {
  let best = -1
  let distance = Infinity
  columns.forEach((column, index) => {
    const delta = Math.abs(column - u)
    if (delta <= tolerance && delta < distance) { best = index; distance = delta }
  })
  return best
}

/**
 * Runs of lines whose segments sit on shared column positions: a line with two
 * or more segments starts (or continues) a table, lines starting under a
 * column continue the cell above, and a line on an earlier row's baseline
 * joins that row (a tall neighbouring cell came first in the content).
 */
function findTables(lines: Line[]) {
  const tables: TableCandidate[] = []
  let index = 0
  while (index < lines.length) {
    const first = lines[index]
    if (first.segments.length < 2) { index += 1; continue }
    const tolerance = Math.max(3, first.size * 0.3)
    const table: TableCandidate = { start: index, end: index, columns: first.segments.map((segment) => segment.u), rows: [] }
    const addRow = (line: Line) => {
      const cells: Line[][] = table.columns.map(() => [])
      for (const segment of line.segments) {
        const column = Math.max(0, columnAt(table.columns, segment.u, tolerance))
        cells[column].push({ ...line, segments: [segment], u: segment.u, right: segment.right })
      }
      table.rows.push({ v: line.v, cells })
    }
    addRow(first)
    let cursor = index + 1
    for (; cursor < lines.length; cursor += 1) {
      const line = lines[cursor]
      const starts = line.segments.map((segment) => columnAt(table.columns, segment.u, tolerance))
      if (line.segments.length >= 2) {
        if (starts.every((column) => column >= 0) && new Set(starts).size === starts.length) {
          const sameRow = table.rows.find((row) => Math.abs(row.v - line.v) <= line.size * 0.4)
          if (sameRow) {
            for (const segment of line.segments) sameRow.cells[columnAt(table.columns, segment.u, tolerance)].push({ ...line, segments: [segment], u: segment.u, right: segment.right })
          } else addRow(line)
          continue
        }
        break
      }
      const column = starts[0]
      if (column < 0) break
      const nextColumn = table.columns[column + 1]
      if (nextColumn !== undefined && line.right > nextColumn + tolerance) break
      const sameRow = table.rows.find((row) => Math.abs(row.v - line.v) <= line.size * 0.4)
      const lastRow = table.rows[table.rows.length - 1]
      const owner = sameRow || (line.v <= lastRow.v + line.size * 0.4 ? lastRow : null)
      if (!owner) break
      // A first-column line far below the row is the text after the table.
      const above = owner.cells[column].at(-1)
      if (column === 0 && !sameRow && above && above.v - line.v > line.size * 2.2) break
      owner.cells[column].push({ ...line, segments: [line.segments[0]] })
    }
    const filled = table.columns.map((_, column) => table.rows.some((row) => row.cells[column].length))
    if (table.rows.length >= 2 && filled.filter(Boolean).length >= 2) {
      table.end = cursor - 1
      tables.push(table)
      index = cursor
    } else index += 1
  }
  return tables
}

function lineToExport(line: Line, hardBreak = false): ExportLine {
  const runs: ExportRun[] = []
  line.segments.forEach((segment, segmentIndex) => {
    if (segmentIndex) runs.push({ text: '\t' })
    for (const fragment of segment.fragments) {
      const run: ExportRun = { text: fragment.text }
      if (fragment.bold) run.bold = true
      if (fragment.italic) run.italic = true
      run.size = Math.round(fragment.size * 2) / 2
      if (fragment.font) run.font = fragment.font
      if (fragment.superscript) run.superscript = true
      const last = runs[runs.length - 1]
      if (last && last.text !== '\t' && Boolean(last.bold) === Boolean(run.bold) && Boolean(last.italic) === Boolean(run.italic)
        && last.size === run.size && last.font === run.font && Boolean(last.superscript) === Boolean(run.superscript)) last.text += run.text
      else runs.push(run)
    }
  })
  if (runs.length) {
    runs[0].text = runs[0].text.replace(/^\s+/, '')
    runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '')
  }
  const result: ExportLine = { runs: runs.filter((run) => run.text) }
  if (hardBreak) result.hardBreak = true
  return result
}

function allBold(line: Line) {
  return line.segments.every((segment) => segment.fragments.every((fragment) => fragment.bold || !fragment.text.trim()))
}

/** Text lines into paragraphs: spacing, size, indentation and short lines end one. */
function groupParagraphs(lines: Line[], pitchFor: (size: number) => number, column: { left: number, right: number }) {
  const blocks: DraftText[] = []
  let block: DraftText | null = null
  let left = 0
  let right = 0
  const start = (line: Line) => {
    block = { type: 'text', lines: [line], top: line.v + line.size * 0.8, size: line.size, bold: allBold(line), chars: lineChars(line), hardBreaks: [], list: BULLET_PATTERN.test(lineText(line)) ? 'bullet' : undefined }
    blocks.push(block)
    left = line.u
    right = line.right
  }
  for (const line of lines) {
    const current = block as DraftText | null
    if (!current) { start(line); continue }
    const previous = current.lines[current.lines.length - 1]
    const gap = previous.v - line.v
    const pitch = pitchFor(previous.size)
    const text = lineText(line)
    let split = gap < -0.5 || gap > Math.max(pitch * 1.3, previous.size * 1.15)
    if (!split && Math.abs(line.size - previous.size) > Math.max(line.size, previous.size) * 0.15) split = true
    if (!split && BULLET_PATTERN.test(text)) split = true
    if (!split && current.bold && !allBold(line) && current.lines.length === 1 && lineChars(previous) <= 120) split = true
    if (!split) {
      const indent = line.u - (current.lines.length === 1 ? previous.u : left)
      const centered = Math.abs((line.u + line.right) / 2 - (previous.u + previous.right) / 2) < 2.5
        && line.right - line.u < (column.right - column.left) * 0.9
      if (Math.abs(indent) > line.size * 1.1 && !centered) {
        // A first line may be indented (or hang out for a bullet).
        const firstLineIndent = current.lines.length === 1 && indent < 0
        const hanging = current.lines.length === 1 && indent > 0 && current.list === 'bullet'
        if (!firstLineIndent && !hanging) split = true
      }
      // A short line before ends its paragraph.
      if (!split && current.lines.length >= 2 && previous.right < right - Math.max(line.size * 3, (right - left) * 0.2) && !centered) split = true
    }
    if (split) { start(line); continue }
    if (current.lines.length === 1) left = line.u
    else left = Math.min(left, line.u)
    right = Math.max(right, line.right)
    current.lines.push(line)
    current.chars += lineChars(line)
    current.bold = current.bold && allBold(line)
  }
  // Forced breaks: the next line's first word would have fitted on this one.
  // A block of short lines is measured against the page's text column.
  for (const draft of blocks) {
    const blockRight = Math.max(...draft.lines.map((line) => line.right))
    const blockWidth = blockRight - Math.min(...draft.lines.map((line) => line.u))
    const available = blockWidth < (column.right - column.left) * 0.85 ? Math.max(blockRight, column.right) : blockRight
    draft.hardBreaks = draft.lines.map((line, index) => {
      const next = draft.lines[index + 1]
      if (!next) return false
      const nextText = lineText(next)
      const average = (next.right - next.u) / Math.max(1, nextText.length)
      const firstWord = (nextText.trim().split(/\s/)[0] || '').length * average
      return line.right + average + firstWord < available - line.size * 0.5
    })
    const centers = draft.lines.map((line) => (line.u + line.right) / 2)
    const columnCenter = (column.left + column.right) / 2
    const widest = Math.max(...draft.lines.map((line) => line.right - line.u))
    if (widest < (column.right - column.left) * 0.85) {
      if (centers.every((center) => Math.abs(center - columnCenter) < 3)) draft.align = 'center'
      else if (draft.lines.length > 1 && draft.lines.every((line) => Math.abs(line.right - blockRight) < 2) && draft.lines.some((line) => Math.abs(line.u - draft.lines[0].u) > 3)) draft.align = 'right'
    }
    if (draft.list === 'bullet') {
      const first = draft.lines[0].segments[0]
      first.fragments[0] = { ...first.fragments[0], text: first.fragments[0].text.replace(BULLET_PATTERN, '') }
      if (!first.fragments[0].text && first.fragments.length > 1) first.fragments.shift()
      if (!first.fragments[0].text.trim() && draft.lines[0].segments.length > 1) draft.lines[0].segments.shift()
    }
  }
  return blocks
}

function multiply(m1: number[], m2: number[]) {
  return [
    m1[0] * m2[0] + m1[1] * m2[2],
    m1[0] * m2[1] + m1[1] * m2[3],
    m1[2] * m2[0] + m1[3] * m2[2],
    m1[2] * m2[1] + m1[3] * m2[3],
    m1[4] * m2[0] + m1[5] * m2[2] + m2[4],
    m1[4] * m2[1] + m1[5] * m2[3] + m2[5],
  ]
}

/**
 * One page of the layout model. `fonts` maps pdf.js font ids to PDF font
 * names; `images` are pictures already encoded as PNG or JPEG.
 */
export function buildPageLayout(input: {
  pageNumber: number
  width: number
  height: number
  items: unknown[]
  fonts?: Record<string, string>
  images?: PlacedImage[]
}): PageLayoutDraft {
  const items = (input.items || []).filter(isTextItem)
  const frame = dominantFrame(items)
  const fontCache = new Map<string, FontInfo>()
  const fonts = (fontName: string) => {
    if (!fontCache.has(fontName)) fontCache.set(fontName, fontInfoFromName(input.fonts?.[fontName] || ''))
    return fontCache.get(fontName) as FontInfo
  }
  const framed = items.filter((item) => !item.str || inFrame(item, frame))
  const others = items.filter((item) => item.str && !inFrame(item, frame))
  let lines = buildLines(framed, frame, fonts)
  // Running headers and footers that are only a page number are left out.
  const vertical = Math.abs(frame.sin) > 0.5 ? input.width : input.height
  lines = lines.filter((line) => {
    const text = lineText(line, ' ')
    if (!PAGE_NUMBER_PATTERN.test(text)) return true
    const fromBottom = frame.cos >= 0 && frame.sin >= 0 ? line.v : vertical + line.v
    return !(fromBottom < vertical * 0.08 || fromBottom > vertical * 0.92)
  })
  const sizes: Record<string, number> = {}
  for (const line of lines) {
    for (const segment of line.segments) {
      for (const fragment of segment.fragments) {
        const key = String(Math.round(fragment.size * 2) / 2)
        sizes[key] = (sizes[key] || 0) + fragment.text.trim().length
      }
    }
  }
  const column = {
    left: lines.length ? Math.min(...lines.map((line) => line.u)) : 0,
    right: lines.length ? Math.max(...lines.map((line) => line.right)) : input.width,
  }
  const pitchFor = modalPitch(lines)
  const drafts: Draft[] = []
  const tables = findTables(lines)
  let cursor = 0
  const flushText = (end: number) => {
    if (end > cursor) drafts.push(...groupParagraphs(lines.slice(cursor, end), pitchFor, column))
  }
  for (const table of tables) {
    flushText(table.start)
    const widths = table.columns.map((start, index) => {
      const next = table.columns[index + 1]
      const end = next !== undefined ? next : Math.max(...table.rows.flatMap((row) => row.cells[index].map((line) => line.right)), start + 36)
      return Math.max(18, end - start)
    })
    drafts.push({ type: 'table', rows: table.rows.map((row) => row.cells), widths, top: table.rows[0].v + lines[table.start].size })
    cursor = table.end + 1
  }
  flushText(lines.length)
  if (others.length) {
    // Text at another angle (a stamp or a side label) follows the page text.
    const otherFrame = dominantFrame(others)
    drafts.push(...groupParagraphs(buildLines(others, otherFrame, fonts), modalPitch([]), column).map((draft) => ({ ...draft, top: -Infinity })))
  }
  // Pictures go before the first block that starts below them.
  for (const image of input.images || []) {
    const [a, b, c, d, e, f] = image.matrix
    const corners = [[e, f], [a + e, b + f], [c + e, d + f], [a + c + e, b + d + f]]
    const top = Math.max(...corners.map(([x, y]) => -x * frame.sin + y * frame.cos))
    const width = Math.hypot(a, b)
    const height = Math.hypot(c, d)
    const draft: DraftImage = { type: 'image', data: image.data, mime: image.mime, width, height, top }
    const at = drafts.findIndex((candidate) => candidate.top < top - 1)
    if (at < 0) drafts.push(draft)
    else drafts.splice(at, 0, draft)
  }
  return { pageNumber: input.pageNumber, width: input.width, height: input.height, drafts, sizes }
}

function blocksToText(blocks: ExportBlock[]) {
  const lineString = (line: ExportLine) => line.runs.map((run) => run.text).join('')
  return blocks.map((block) => {
    if (block.type === 'image') return ''
    if (block.type === 'table') return block.rows.map((row) => row.map((cell) => cell.map(lineString).join(' ')).join('\t')).join('\n')
    if (block.type === 'heading') return block.lines.map(lineString).join(' ')
    const lines = block.lines.map(lineString)
    if (block.list === 'bullet' && lines.length) lines[0] = `• ${lines[0]}`
    return lines.join('\n')
  }).filter(Boolean).join('\n\n')
}

/**
 * Turns page drafts into export pages: the most common text size is the body
 * size, and short paragraphs set larger than it become headings, ranked by
 * size across all the pages (largest = Heading 1).
 */
export function finishLayout(pages: PageLayoutDraft[]): StructuredTextPage[] {
  const totals = new Map<number, number>()
  for (const page of pages) for (const [size, count] of Object.entries(page.sizes)) totals.set(Number(size), (totals.get(Number(size)) || 0) + count)
  const body = [...totals].sort((left, right) => right[1] - left[1] || left[0] - right[0])[0]?.[0] || 11
  const isHeading = (draft: Draft): boolean => draft.type === 'text' && !draft.list && draft.size >= body * 1.15
    && draft.lines.length <= 3 && draft.chars <= 200 && draft.chars > 0
  const headingSizes = [...new Set(pages.flatMap((page) => page.drafts.filter((draft): draft is DraftText => isHeading(draft)).map((draft) => Math.round(draft.size * 2) / 2)))].sort((left, right) => right - left)
  const levelFor = (size: number) => Math.min(3, headingSizes.indexOf(Math.round(size * 2) / 2) + 1) as 1 | 2 | 3
  return pages.map((page) => {
    const blocks: ExportBlock[] = page.drafts.map((draft): ExportBlock => {
      if (draft.type === 'image') return { type: 'image', data: draft.data, mime: draft.mime, width: draft.width, height: draft.height }
      if (draft.type === 'table') {
        return { type: 'table', widths: draft.widths, rows: draft.rows.map((row) => row.map((cell) => cell.map((line) => lineToExport(line)))) }
      }
      const lines = draft.lines.map((line, index) => lineToExport(line, draft.hardBreaks[index]))
      const text: DraftText = draft
      if (isHeading(text)) return { type: 'heading', level: levelFor(text.size), lines, ...(text.align ? { align: text.align } : {}) }
      return { type: 'paragraph', lines, ...(text.align ? { align: text.align } : {}), ...(text.list ? { list: text.list } : {}) }
    })
    return { pageNumber: page.pageNumber, width: page.width, height: page.height, blocks, text: blocksToText(blocks) }
  })
}

// ---------------------------------------------------------------------------
// Pictures from the operator list
// ---------------------------------------------------------------------------

// pdf.js OPS codes (stable across pdf.js 3 and 4); importing pdf.js here would
// pull the library into the Node tests that load this module.
const OP_SAVE = 10
const OP_RESTORE = 11
const OP_TRANSFORM = 12
const OP_FORM_BEGIN = 74
const OP_FORM_END = 75
const OP_IMAGE = 85
const OP_INLINE_IMAGE = 86
const MAX_IMAGE_BYTES = 96 * 1024 * 1024
const MAX_IMAGES = 400

interface DecodedImage {
  width: number
  height: number
  kind?: number
  data?: Uint8Array | Uint8ClampedArray
  bitmap?: CanvasImageSource & { width: number; height: number }
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let index = 0; index < 256; index += 1) {
    let value = index
    for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
    table[index] = value >>> 0
  }
  return table
})()

function crc32(parts: Uint8Array[]) {
  let value = 0xffffffff
  for (const part of parts) for (let index = 0; index < part.length; index += 1) value = CRC_TABLE[(value ^ part[index]) & 0xff] ^ (value >>> 8)
  return (value ^ 0xffffffff) >>> 0
}

async function deflate(data: Uint8Array) {
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new CompressionStream('deflate'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

/** An RGBA PNG without a canvas (used where no canvas exists). */
async function encodeRgbaPng(width: number, height: number, rgba: Uint8Array | Uint8ClampedArray) {
  const opaque = rgba.every((value, index) => index % 4 !== 3 || value === 255)
  const channels = opaque ? 3 : 4
  const raw = new Uint8Array((width * channels + 1) * height)
  for (let row = 0; row < height; row += 1) {
    let at = row * (width * channels + 1) + 1
    for (let x = 0; x < width; x += 1) {
      const from = (row * width + x) * 4
      raw[at++] = rgba[from]
      raw[at++] = rgba[from + 1]
      raw[at++] = rgba[from + 2]
      if (channels === 4) raw[at++] = rgba[from + 3]
    }
  }
  const chunk = (type: string, body: Uint8Array) => {
    const head = new Uint8Array(8)
    const view = new DataView(head.buffer)
    view.setUint32(0, body.length)
    for (let index = 0; index < 4; index += 1) head[4 + index] = type.charCodeAt(index)
    const tail = new Uint8Array(4)
    new DataView(tail.buffer).setUint32(0, crc32([head.subarray(4), body]))
    return [head, body, tail]
  }
  const header = new Uint8Array(13)
  const view = new DataView(header.buffer)
  view.setUint32(0, width)
  view.setUint32(4, height)
  header[8] = 8
  header[9] = channels === 4 ? 6 : 2
  const parts = [Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a), ...chunk('IHDR', header), ...chunk('IDAT', await deflate(raw)), ...chunk('IEND', new Uint8Array(0))]
  const output = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) { output.set(part, offset); offset += part.length }
  return output
}

function toRgba(image: DecodedImage) {
  const { width, height, data, kind } = image
  if (!data) return null
  const rgba = new Uint8ClampedArray(width * height * 4)
  if (kind === 3) rgba.set(data.subarray(0, rgba.length))
  else if (kind === 2) {
    for (let pixel = 0, from = 0; pixel < width * height; pixel += 1, from += 3) {
      rgba[pixel * 4] = data[from]; rgba[pixel * 4 + 1] = data[from + 1]; rgba[pixel * 4 + 2] = data[from + 2]; rgba[pixel * 4 + 3] = 255
    }
  } else if (kind === 1) {
    const rowBytes = Math.ceil(width / 8)
    for (let row = 0; row < height; row += 1) {
      for (let x = 0; x < width; x += 1) {
        const value = (data[row * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1 ? 255 : 0
        const at = (row * width + x) * 4
        rgba[at] = value; rgba[at + 1] = value; rgba[at + 2] = value; rgba[at + 3] = 255
      }
    }
  } else return null
  return rgba
}

function makeCanvas(width: number, height: number): { canvas: OffscreenCanvas | HTMLCanvasElement, context: OffscreenCanvasRenderingContext2D | CanvasRenderingContext2D } | null {
  if (typeof OffscreenCanvas !== 'undefined') {
    const canvas = new OffscreenCanvas(width, height)
    const context = canvas.getContext('2d')
    return context ? { canvas, context } : null
  }
  if (typeof document !== 'undefined' && typeof document.createElement === 'function') {
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const context = canvas.getContext('2d')
    return context ? { canvas, context } : null
  }
  return null
}

async function canvasToBytes(canvas: OffscreenCanvas | HTMLCanvasElement, mime: 'image/png' | 'image/jpeg') {
  const blob = 'convertToBlob' in canvas
    ? await canvas.convertToBlob({ type: mime, quality: 0.9 })
    : await new Promise<Blob | null>((resolve) => (canvas as HTMLCanvasElement).toBlob(resolve, mime, 0.9))
  return blob ? new Uint8Array(await blob.arrayBuffer()) : null
}

/** A decoded pdf.js image as PNG, or JPEG for large opaque pictures. */
async function encodeImage(image: DecodedImage): Promise<{ data: Uint8Array, mime: 'image/png' | 'image/jpeg' } | null> {
  const width = Math.round(image.width)
  const height = Math.round(image.height)
  if (!(width >= 8 && height >= 8) || width * height > 40_000_000) return null
  const rgba = image.bitmap ? null : toRgba(image)
  if (!image.bitmap && !rgba) return null
  const drawn = makeCanvas(width, height)
  if (drawn) {
    if (image.bitmap) drawn.context.drawImage(image.bitmap, 0, 0, width, height)
    else drawn.context.putImageData(new ImageData(rgba as Uint8ClampedArray<ArrayBuffer>, width, height), 0, 0)
    const pixels = drawn.context.getImageData(0, 0, width, height).data
    let opaque = true
    for (let index = 3; index < pixels.length; index += 4) if (pixels[index] !== 255) { opaque = false; break }
    const mime = opaque && width * height > 160_000 ? 'image/jpeg' : 'image/png'
    const data = await canvasToBytes(drawn.canvas, mime)
    return data ? { data, mime } : null
  }
  if (!rgba || typeof CompressionStream === 'undefined') return null
  return { data: await encodeRgbaPng(width, height, rgba), mime: 'image/png' }
}

interface PdfObjects {
  has?: (id: string) => boolean
  get: (id: string, callback?: (value: unknown) => void) => unknown
}

function resolvedObject(store: PdfObjects | undefined, id: string) {
  return new Promise<unknown>((resolve) => {
    if (!store) { resolve(null); return }
    let settled = false
    const timer = setTimeout(() => { settled = true; resolve(null) }, 4000)
    try {
      store.get(id, (value) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        resolve(value)
      })
    } catch {
      clearTimeout(timer)
      resolve(null)
    }
  })
}

interface ImageBudget { bytes: number, count: number }

interface OperatorPage {
  getOperatorList?: () => Promise<{ fnArray: number[], argsArray: unknown[][] }>
  objs?: PdfObjects
  commonObjs?: PdfObjects
}

/** Pictures drawn on a page with their placement, encoded for the export. */
async function pageImages(page: OperatorPage, list: { fnArray: number[], argsArray: unknown[][] }, pageArea: number, textChars: number, budget: ImageBudget) {
  const placed: { id?: string, inline?: DecodedImage, matrix: number[] }[] = []
  let matrix = [1, 0, 0, 1, 0, 0]
  const stack: number[][] = []
  for (let index = 0; index < list.fnArray.length; index += 1) {
    const operator = list.fnArray[index]
    const args = list.argsArray[index] || []
    if (operator === OP_SAVE) stack.push(matrix)
    else if (operator === OP_RESTORE) matrix = stack.pop() || matrix
    else if (operator === OP_TRANSFORM) matrix = multiply(args as number[], matrix)
    else if (operator === OP_FORM_BEGIN) {
      stack.push(matrix)
      if (Array.isArray(args[0]) && args[0].length === 6) matrix = multiply(args[0] as number[], matrix)
    } else if (operator === OP_FORM_END) matrix = stack.pop() || matrix
    else if (operator === OP_IMAGE && typeof args[0] === 'string') placed.push({ id: args[0], matrix })
    else if (operator === OP_INLINE_IMAGE && args[0] && typeof args[0] === 'object') placed.push({ inline: args[0] as DecodedImage, matrix })
  }
  const images: PlacedImage[] = []
  for (const item of placed) {
    if (budget.count >= MAX_IMAGES || budget.bytes >= MAX_IMAGE_BYTES) break
    const [a, b, c, d] = item.matrix
    const width = Math.hypot(a, b)
    const height = Math.hypot(c, d)
    if (width < 6 || height < 6) continue
    // A picture covering the page under real text is a scan or a background.
    if (width * height >= pageArea * 0.85 && textChars >= 20) continue
    const decoded = item.inline || await resolvedObject(item.id?.startsWith('g_') ? page.commonObjs : page.objs, item.id || '') as DecodedImage | null
    if (!decoded || typeof decoded !== 'object') continue
    const encoded = await encodeImage(decoded).catch(() => null)
    if (!encoded) continue
    budget.bytes += encoded.data.length
    budget.count += 1
    images.push({ ...encoded, matrix: item.matrix })
  }
  return images
}

interface ExtractPage {
  getTextContent: () => Promise<{ items: unknown[], styles?: unknown }>
  getViewport?: (options: { scale: number }) => { width: number, height: number }
  view?: number[]
  rotate?: number
  commonObjs?: PdfObjects
  objs?: PdfObjects
  getOperatorList?: () => Promise<{ fnArray: number[], argsArray: unknown[][] }>
  cleanup: () => void
}

function fontNames(page: ExtractPage, items: unknown[]) {
  const names: Record<string, string> = {}
  for (const item of items) {
    const fontName = isTextItem(item) ? item.fontName : undefined
    if (!fontName || fontName in names) continue
    names[fontName] = ''
    try {
      if (page.commonObjs && (!page.commonObjs.has || page.commonObjs.has(fontName))) {
        const font = page.commonObjs.get(fontName) as { name?: string } | null
        names[fontName] = String(font?.name || '')
      }
    } catch { /* not loaded */ }
  }
  return names
}

/**
 * The text and layout of the given pages, for the Word, web page, Markdown
 * and text exports. Each page keeps `text` (its plain text with line breaks).
 */
export async function extractPdfText(pdf: PDFDocumentProxy, pageIndices: number[], options: ExtractTextOptions = {}): Promise<StructuredTextPage[]> {
  const drafts: PageLayoutDraft[] = []
  const budget: ImageBudget = { bytes: 0, count: 0 }
  for (const pageIndex of pageIndices) {
    const page = await pdf.getPage(pageIndex + 1) as unknown as ExtractPage
    try {
      const content = await page.getTextContent()
      const viewport = typeof page.getViewport === 'function' ? page.getViewport({ scale: 1 }) : null
      const view = Array.isArray(page.view) ? page.view : [0, 0, 612, 792]
      const width = viewport?.width || Math.abs(view[2] - view[0]) || 612
      const height = viewport?.height || Math.abs(view[3] - view[1]) || 792
      const textChars = content.items.reduce((sum: number, item) => sum + (isTextItem(item) ? item.str.trim().length : 0), 0)
      let images: PlacedImage[] = []
      try {
        // The operator list also loads the fonts whose names give bold and italic.
        const list = options.images !== false && typeof page.getOperatorList === 'function' ? await page.getOperatorList() : null
        if (list) images = await pageImages(page, list, width * height, textChars, budget)
      } catch {
        images = []
      }
      drafts.push(buildPageLayout({ pageNumber: pageIndex + 1, width, height, items: content.items, fonts: fontNames(page, content.items), images }))
    } finally {
      page.cleanup()
    }
    options.onProgress?.(drafts.length, pageIndices.length)
  }
  return finishLayout(drafts)
}

function canvasBlob(canvas: HTMLCanvasElement, mimeType: string, quality?: number) {
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error('The rendered page could not be encoded.')), mimeType, quality)
  })
}

export async function renderPdfPageImage(
  pdf: PDFDocumentProxy,
  pageIndex: number,
  format: ExportImageFormat,
  requestedScale: number,
  jpegQuality: number,
) {
  const page = await pdf.getPage(pageIndex + 1)
  let canvas: HTMLCanvasElement | null = null
  try {
    const unitViewport = page.getViewport({ scale: 1, rotation: page.rotate || 0 })
    const pixelLimitScale = Math.sqrt(64_000_000 / Math.max(1, unitViewport.width * unitViewport.height))
    const dimensionLimitScale = 8_192 / Math.max(1, unitViewport.width, unitViewport.height)
    // Apply safety caps last: a minimum of 0.1 *after* the cap still creates
    // enormous canvases for large-format PDFs and makes every image codec fail.
    const scale = Math.min(Math.max(0.01, Number(requestedScale) || 1), pixelLimitScale, dimensionLimitScale)
    const viewport = page.getViewport({ scale, rotation: page.rotate || 0 })
    canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.floor(viewport.width))
    canvas.height = Math.max(1, Math.floor(viewport.height))
    const context = canvas.getContext('2d', { alpha: format !== 'jpeg' })
    if (!context) throw new Error('The page image canvas could not be created.')
    if (format === 'jpeg') {
      context.fillStyle = '#ffffff'
      context.fillRect(0, 0, canvas.width, canvas.height)
    }
    await page.render({ canvasContext: context, viewport,
      transform: [canvas.width / viewport.width, 0, 0, canvas.height / viewport.height, 0, 0],
    }).promise
    const mimeType = format === 'jpeg' ? 'image/jpeg' : format === 'webp' ? 'image/webp' : 'image/png'
    const blob = await canvasBlob(canvas, mimeType, format === 'jpeg' || format === 'webp' ? jpegQuality : undefined)
    return new Uint8Array(await blob.arrayBuffer())
  } finally {
    if (canvas) {
      canvas.width = 0
      canvas.height = 0
    }
    page.cleanup()
  }
}
