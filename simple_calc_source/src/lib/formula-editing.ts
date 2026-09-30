/**
 * Tolerant formula-text helpers for the editor: syntax tokens for highlighting, the
 * references a formula mentions (Excel's colour-coded "range finder"), the function call
 * and argument under the caret, autocomplete context, point-mode insertion, and F4
 * absolute/relative toggling. Nothing here throws on incomplete input like `=SUM(A1:`.
 */

export type EditorTokenKind =
  | 'equals'
  | 'function'
  | 'reference'
  | 'name'
  | 'number'
  | 'string'
  | 'error'
  | 'operator'
  | 'open'
  | 'close'
  | 'separator'
  | 'brace'
  | 'boolean'
  | 'space'
  | 'text'

export interface EditorToken {
  kind: EditorTokenKind
  start: number
  end: number
  text: string
  /** For references: the parsed target. */
  reference?: FormulaReference
}

export interface FormulaReference {
  /** Sheet name as written (unquoted), when the reference is qualified. */
  sheet?: string
  /** 0-based inclusive bounds. Whole columns/rows use the sheet limits. */
  top: number
  left: number
  bottom: number
  right: number
  /** Text of the reference as written, including any sheet prefix. */
  text: string
  start: number
  end: number
  spill?: boolean
}

export const REFERENCE_COLORS = ['#1f6fd1', '#c0392b', '#7d3cb5', '#1e8449', '#b9770e', '#a93270', '#117a8b', '#6e4b2a']

const MAX_ROWS = 1_048_576
const MAX_COLUMNS = 16_384

function columnIndex(label: string) {
  let value = 0
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value - 1
}

function columnLabel(index: number) {
  let value = index + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

const SHEET_PREFIX = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!`
const CELL = String.raw`\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}`
const REFERENCE_RE = new RegExp(
  String.raw`^(${SHEET_PREFIX})?(?:(${CELL})(?::(${CELL}))?(#)?|(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})|(\$?[1-9]\d{0,6}):(\$?[1-9]\d{0,6}))`,
)

function parseCell(text: string) {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(text)
  if (!match) return null
  const col = columnIndex(match[1])
  const row = Number(match[2]) - 1
  if (col < 0 || col >= MAX_COLUMNS || row < 0 || row >= MAX_ROWS) return null
  return { row, col }
}

function unquoteSheet(prefix: string | undefined) {
  if (!prefix) return undefined
  const name = prefix.slice(0, -1)
  return name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name
}

/** Split formula text into editor tokens. Works on partial input. */
export function tokenizeFormulaText(text: string): EditorToken[] {
  const tokens: EditorToken[] = []
  if (!text.startsWith('=')) return [{ kind: 'text', start: 0, end: text.length, text }]
  tokens.push({ kind: 'equals', start: 0, end: 1, text: '=' })
  let position = 1
  const push = (kind: EditorTokenKind, end: number, reference?: FormulaReference) => {
    tokens.push({ kind, start: position, end, text: text.slice(position, end), reference })
    position = end
  }
  while (position < text.length) {
    const character = text[position]
    const rest = text.slice(position)
    if (/\s/.test(character)) {
      let end = position + 1
      while (end < text.length && /\s/.test(text[end])) end += 1
      push('space', end)
      continue
    }
    if (character === '"') {
      let end = position + 1
      while (end < text.length) {
        if (text[end] === '"') {
          if (text[end + 1] === '"') end += 2
          else { end += 1; break }
        } else end += 1
      }
      push('string', Math.min(end, text.length))
      continue
    }
    if (character === '#') {
      const match = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)/i.exec(rest)
      if (match) { push('error', position + match[0].length); continue }
    }
    // References (optionally sheet-qualified) — but not function names like LOG10(.
    const reference = REFERENCE_RE.exec(rest)
    if (reference && (position === 1 || !/[A-Za-z0-9_.$]/.test(text[position - 1]))) {
      const end = position + reference[0].length
      const following = text[end] ?? ''
      // LOG10( and ATAN2( look like cells; a bare cell followed by "(" is a function name.
      const looksLikeFunction = !reference[1] && /^[A-Za-z]{1,3}\d+$/.test(reference[0]) && /^\s*\(/.test(text.slice(end))
      if (!/[A-Za-z0-9_.]/.test(following) && !looksLikeFunction) {
        const sheet = unquoteSheet(reference[1])
        let bounds: { top: number; left: number; bottom: number; right: number } | null = null
        if (reference[2]) {
          const first = parseCell(reference[2])
          const second = reference[3] ? parseCell(reference[3]) : first
          if (first && second) bounds = { top: Math.min(first.row, second.row), left: Math.min(first.col, second.col), bottom: Math.max(first.row, second.row), right: Math.max(first.col, second.col) }
        } else if (reference[5]) {
          const a = columnIndex(reference[5].replace('$', ''))
          const b = columnIndex(reference[6].replace('$', ''))
          bounds = { top: 0, bottom: MAX_ROWS - 1, left: Math.min(a, b), right: Math.max(a, b) }
        } else if (reference[7]) {
          const a = Number(reference[7].replace('$', '')) - 1
          const b = Number(reference[8].replace('$', '')) - 1
          bounds = { top: Math.min(a, b), bottom: Math.max(a, b), left: 0, right: MAX_COLUMNS - 1 }
        }
        if (bounds) {
          push('reference', end, { ...bounds, sheet, text: reference[0], start: position, end, spill: Boolean(reference[4]) })
          continue
        }
      }
    }
    const identifier = /^(?:_xlfn\.|_xlws\.|_xlpm\.)*[A-Za-z_\\À-￿][A-Za-z0-9_.?À-￿]*/.exec(rest)
    if (identifier) {
      const end = position + identifier[0].length
      if (text[end] === '[') {
        // Structured reference: Table1[Column] / Table1[[#Headers],[A]]
        let depth = 0
        let close = end
        for (; close < text.length; close += 1) {
          if (text[close] === "'") { close += 1; continue }
          if (text[close] === '[') depth += 1
          else if (text[close] === ']') { depth -= 1; if (depth === 0) { close += 1; break } }
        }
        push('name', Math.min(close, text.length))
        continue
      }
      const upper = identifier[0].toUpperCase()
      if (/^\s*\(/.test(text.slice(end))) push('function', end)
      else if (upper === 'TRUE' || upper === 'FALSE') push('boolean', end)
      else push('name', end)
      continue
    }
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?%?/.exec(rest)
    if (number) { push('number', position + number[0].length); continue }
    if (character === '[') {
      let depth = 0
      let close = position
      for (; close < text.length; close += 1) {
        if (text[close] === '[') depth += 1
        else if (text[close] === ']') { depth -= 1; if (depth === 0) { close += 1; break } }
      }
      push('name', Math.min(close, text.length))
      continue
    }
    if (character === '(') { push('open', position + 1); continue }
    if (character === ')') { push('close', position + 1); continue }
    if (character === ',' || character === ';') { push('separator', position + 1); continue }
    if (character === '{' || character === '}') { push('brace', position + 1); continue }
    const operator = /^(?:<=|>=|<>|[-+*/^&=<>%:@!])/.exec(rest)
    if (operator) { push('operator', position + operator[0].length); continue }
    push('text', position + 1)
  }
  return tokens
}

export interface ColoredReference extends FormulaReference {
  color: string
  colorIndex: number
}

/** References in order of appearance, coloured like Excel (same text → same colour). */
export function formulaReferences(text: string): ColoredReference[] {
  const tokens = tokenizeFormulaText(text)
  const colors = new Map<string, number>()
  const output: ColoredReference[] = []
  for (const token of tokens) {
    if (token.kind !== 'reference' || !token.reference) continue
    const key = token.reference.text.replace(/\$/g, '').toUpperCase()
    let index = colors.get(key)
    if (index === undefined) {
      index = colors.size % REFERENCE_COLORS.length
      colors.set(key, index)
    }
    output.push({ ...token.reference, color: REFERENCE_COLORS[index], colorIndex: index })
  }
  return output
}

export interface CallContext {
  name: string
  /** 0-based argument index at the caret. */
  argumentIndex: number
  /** Offset of the function name. */
  start: number
}

/** Innermost function call containing the caret, and which argument the caret is in. */
export function callContextAt(text: string, caret: number): CallContext | null {
  if (!text.startsWith('=')) return null
  const tokens = tokenizeFormulaText(text.slice(0, caret))
  const stack: Array<{ name: string | null; argumentIndex: number; start: number; braceDepth: number }> = []
  let braceDepth = 0
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index]
    if (token.kind === 'brace') braceDepth += token.text === '{' ? 1 : -1
    else if (token.kind === 'open') {
      const previous = tokens[index - 1]
      const name = previous?.kind === 'function' ? previous.text : null
      stack.push({ name, argumentIndex: 0, start: previous?.kind === 'function' ? previous.start : token.start, braceDepth })
    } else if (token.kind === 'close') stack.pop()
    else if (token.kind === 'separator' && stack.length && braceDepth === stack[stack.length - 1].braceDepth) stack[stack.length - 1].argumentIndex += 1
  }
  for (let index = stack.length - 1; index >= 0; index -= 1) {
    const frame = stack[index]
    if (frame.name) return { name: frame.name.replace(/^(?:_xlfn\.)?(?:_xlws\.)?/i, '').toUpperCase(), argumentIndex: frame.argumentIndex, start: frame.start }
  }
  return null
}

export interface CompletionContext {
  prefix: string
  start: number
  end: number
}

/** The identifier being typed at the caret, when a function/name completion makes sense. */
export function completionContextAt(text: string, caret: number): CompletionContext | null {
  if (!text.startsWith('=')) return null
  const before = text.slice(0, caret)
  const tokens = tokenizeFormulaText(before)
  const last = tokens[tokens.length - 1]
  if (!last || last.end !== caret) return null
  if (last.kind !== 'name' && last.kind !== 'reference' && last.kind !== 'boolean') return null
  if (last.kind === 'reference' && (last.text.includes(':') || last.text.includes('$') || last.text.includes('!'))) return null
  if (!/^[A-Za-z_]/.test(last.text)) return null
  if (text[caret] && /[A-Za-z0-9_.(]/.test(text[caret])) return null
  return { prefix: last.text, start: last.start, end: last.end }
}

/** Whether a reference may be inserted at the caret by clicking/arrowing (Excel "Point" mode). */
export function canInsertReferenceAt(text: string, caret: number): boolean {
  if (!text.startsWith('=')) return false
  let index = caret - 1
  while (index >= 0 && text[index] === ' ') index -= 1
  if (index < 0) return false
  const character = text[index]
  if ('=(,;+-*/^&<>:{'.includes(character)) {
    // Not inside a string literal.
    const quotes = (text.slice(0, caret).match(/"/g) || []).length
    return quotes % 2 === 0
  }
  return false
}

/** The reference token that ends exactly at (or contains) the caret, if any. */
export function referenceAt(text: string, caret: number): FormulaReference | null {
  for (const token of tokenizeFormulaText(text)) {
    if (token.kind === 'reference' && token.reference && token.start <= caret && caret <= token.end) return token.reference
  }
  return null
}

export function referenceText(bounds: { top: number; left: number; bottom: number; right: number }, sheet?: string) {
  const start = `${columnLabel(bounds.left)}${bounds.top + 1}`
  const end = `${columnLabel(bounds.right)}${bounds.bottom + 1}`
  const range = start === end ? start : `${start}:${end}`
  if (!sheet) return range
  const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) && !/^[A-Za-z]{1,3}\d+$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`
  return `${quoted}!${range}`
}

/**
 * Insert or replace a pointed reference. `replace` is the span of a reference previously
 * inserted by pointing (so arrow keys move it instead of appending another).
 */
export function insertReference(
  text: string,
  caret: number,
  bounds: { top: number; left: number; bottom: number; right: number },
  replace: { start: number; end: number } | null,
  sheet?: string,
) {
  const reference = referenceText(bounds, sheet)
  const start = replace ? replace.start : caret
  const end = replace ? replace.end : caret
  const next = text.slice(0, start) + reference + text.slice(end)
  return { text: next, caret: start + reference.length, span: { start, end: start + reference.length } }
}

function cycleAnchors(reference: string) {
  // A1 -> $A$1 -> A$1 -> $A1 -> A1
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(reference)
  if (!match) return reference
  const [, columnAnchor, column, rowAnchor, row] = match
  if (!columnAnchor && !rowAnchor) return `$${column}$${row}`
  if (columnAnchor && rowAnchor) return `${column}$${row}`
  if (!columnAnchor && rowAnchor) return `$${column}${row}`
  return `${column}${row}`
}

/** F4: cycle absolute/relative anchors of the reference at (or just before) the caret. */
export function toggleAbsoluteReference(text: string, caret: number, selectionEnd = caret) {
  const tokens = tokenizeFormulaText(text)
  const targets = tokens.filter((token) => token.kind === 'reference' && (
    (token.start <= caret && caret <= token.end) || (token.start < selectionEnd && token.end > caret)
  ))
  if (!targets.length) return null
  let output = ''
  let last = 0
  let newCaret = caret
  let newEnd = selectionEnd
  for (const token of targets) {
    const reference = token.reference!
    const prefixLength = reference.text.lastIndexOf('!') + 1
    const prefix = reference.text.slice(0, prefixLength)
    const body = reference.text.slice(prefixLength)
    const hash = body.endsWith('#') ? '#' : ''
    const core = hash ? body.slice(0, -1) : body
    const parts = core.split(':')
    const first = cycleAnchors(parts[0])
    // Excel toggles both corners together, deriving the second from the first's state.
    const state = { column: first.startsWith('$'), row: /\$\d/.test(first) }
    const apply = (part: string) => {
      const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(part)
      if (!match) return part
      return `${state.column ? '$' : ''}${match[1]}${state.row ? '$' : ''}${match[2]}`
    }
    const replaced = prefix + [first, ...parts.slice(1).map(apply)].join(':') + hash
    output += text.slice(last, token.start) + replaced
    const delta = replaced.length - reference.text.length
    if (token.start < caret) newCaret = Math.max(token.start, caret + (caret >= token.end ? delta : 0))
    if (targets.length === 1) {
      newCaret = token.start
      newEnd = token.start + replaced.length
    } else newEnd += delta
    last = token.end
  }
  output += text.slice(last)
  return { text: output, selectionStart: newCaret, selectionEnd: newEnd }
}

/** Colour class for a token in the highlighted editor layer. */
export function tokenColor(token: EditorToken, references: ColoredReference[]): string | undefined {
  if (token.kind === 'reference') return references.find((reference) => reference.start === token.start)?.color
  if (token.kind === 'string') return '#1c7a3c'
  if (token.kind === 'number' || token.kind === 'boolean') return '#1d5fa8'
  if (token.kind === 'error') return '#b3261e'
  if (token.kind === 'function') return '#222421'
  return undefined
}

function quoteSheetName(name: string) {
  return /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`
}

/** Rewrite references to a renamed sheet (Excel updates every formula on rename). */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  const text = `=${formula}`
  const wanted = oldName.toLocaleLowerCase()
  let output = ''
  let last = 0
  const tokens = tokenizeFormulaText(text)
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index]
    if (token.kind === 'reference' && token.reference?.sheet !== undefined) {
      if (token.reference.sheet.toLocaleLowerCase() !== wanted) continue
      const bang = token.text.lastIndexOf('!')
      output += text.slice(last, token.start) + quoteSheetName(newName) + token.text.slice(bang)
      last = token.end
    } else if (token.kind === 'name' && tokens[index + 1]?.kind === 'operator' && tokens[index + 1].text === '!') {
      // Sheet-qualified names such as Sheet1!TaxRate.
      if (token.text.toLocaleLowerCase() !== wanted) continue
      output += text.slice(last, token.start) + quoteSheetName(newName)
      last = token.end
    }
  }
  if (!last) return formula
  output += text.slice(last)
  return output.slice(1)
}

/** Replace references to a deleted sheet with #REF! (as Excel does). */
export function removeSheetFromFormula(formula: string, sheetName: string): string {
  const text = `=${formula}`
  const wanted = sheetName.toLocaleLowerCase()
  let output = ''
  let last = 0
  for (const token of tokenizeFormulaText(text)) {
    if (token.kind !== 'reference' || token.reference?.sheet?.toLocaleLowerCase() !== wanted) continue
    output += text.slice(last, token.start) + '#REF!'
    last = token.end
  }
  if (!last) return formula
  output += text.slice(last)
  return output.slice(1)
}

export interface ReferenceMove {
  /** Sheet the formula lives on. */
  formulaSheet: string
  /** Sheet the moved block came from, and its 0-based bounds. */
  sourceSheet: string
  rect: { top: number; left: number; bottom: number; right: number }
  rowDelta: number
  colDelta: number
  /** Sheet the block moved to (defaults to the source sheet). */
  destinationSheet?: string
}

function movedCell(text: string, rowDelta: number, colDelta: number) {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(text)
  if (!match) return text
  const column = columnIndex(match[2]) + colDelta
  const row = Number(match[4]) + rowDelta
  if (column < 0 || row < 1) return '#REF!'
  return `${match[1]}${columnLabel(column)}${match[3]}${row}`
}

/**
 * Cut-and-paste semantics: references that lie entirely inside the moved block follow it
 * (absolute anchors too, as in Excel), including references from other sheets.
 */
export function moveReferencesInFormula(formula: string, move: ReferenceMove): string {
  const text = `=${formula}`
  const source = move.sourceSheet.toLocaleLowerCase()
  const destination = (move.destinationSheet ?? move.sourceSheet)
  const formulaSheet = move.formulaSheet.toLocaleLowerCase()
  let output = ''
  let last = 0
  for (const token of tokenizeFormulaText(text)) {
    const reference = token.reference
    if (token.kind !== 'reference' || !reference) continue
    const sheet = (reference.sheet ?? move.formulaSheet).toLocaleLowerCase()
    if (sheet !== source) continue
    const { rect } = move
    if (reference.top < rect.top || reference.bottom > rect.bottom || reference.left < rect.left || reference.right > rect.right) continue
    // Whole-row/column references never "move".
    if (reference.bottom - reference.top >= 1_048_575 || reference.right - reference.left >= 16_383) continue
    const bang = reference.text.lastIndexOf('!')
    const body = reference.text.slice(bang + 1).replace(/#$/, '')
    const spill = reference.text.endsWith('#') ? '#' : ''
    const moved = body.split(':').map((part) => movedCell(part, move.rowDelta, move.colDelta)).join(':')
    let prefix = bang >= 0 ? reference.text.slice(0, bang + 1) : ''
    if (destination.toLocaleLowerCase() !== sheet) prefix = destination.toLocaleLowerCase() === formulaSheet ? '' : `${quoteSheetName(destination)}!`
    output += text.slice(last, token.start) + prefix + moved + spill
    last = token.end
  }
  if (!last) return formula
  return (output + text.slice(last)).slice(1)
}
