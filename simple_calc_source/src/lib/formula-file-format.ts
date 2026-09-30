import { tokenizeFormulaText } from './formula-editing'
import type { EditorToken } from './formula-editing'

/**
 * Excel stores functions introduced after Excel 2007 with an `_xlfn.` prefix (and the
 * worksheet-namespace ones with `_xlfn._xlws.`), LET/LAMBDA parameter names with `_xlpm.`,
 * `A1#` as `_xlfn.ANCHORARRAY(A1)` and `@x` as `_xlfn.SINGLE(x)`. Without the prefixes
 * Excel treats the functions as unknown (#NAME?) when it opens the file. Formulas are held
 * in their friendly form in the editor and converted at the file boundary.
 */
const WORKSHEET_NAMESPACE = new Set(['FILTER', 'SORT'])
export const FUTURE_FUNCTIONS = new Set([
  'ACOT', 'ACOTH', 'AGGREGATE', 'ARABIC', 'ARRAYTOTEXT', 'BASE', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST',
  'BINOM.DIST.RANGE', 'BINOM.INV', 'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR', 'BYCOL', 'BYROW',
  'CEILING.MATH', 'CEILING.PRECISE', 'CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST',
  'CHOOSECOLS', 'CHOOSEROWS', 'COMBINA', 'CONCAT', 'CONFIDENCE.NORM', 'CONFIDENCE.T', 'COT', 'COTH',
  'COVARIANCE.P', 'COVARIANCE.S', 'CSC', 'CSCH', 'DAYS', 'DECIMAL', 'DROP', 'ERF.PRECISE', 'ERFC.PRECISE',
  'EXPAND', 'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST', 'FIELDVALUE', 'FILTER',
  'FILTERXML', 'FLOOR.MATH', 'FLOOR.PRECISE', 'FORECAST.ETS', 'FORECAST.ETS.CONFINT', 'FORECAST.ETS.SEASONALITY',
  'FORECAST.ETS.STAT', 'FORECAST.LINEAR', 'FORMULATEXT', 'GAMMA', 'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN.PRECISE',
  'GAUSS', 'GROUPBY', 'HSTACK', 'HYPGEOM.DIST', 'IFNA', 'IFS', 'IMAGE', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH',
  'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN', 'ISFORMULA', 'ISOMITTED', 'ISOWEEKNUM', 'LAMBDA', 'LET',
  'LOGNORM.DIST', 'LOGNORM.INV', 'MAKEARRAY', 'MAP', 'MAXIFS', 'MINIFS', 'MODE.MULT', 'MODE.SNGL', 'MUNIT',
  'NEGBINOM.DIST', 'NETWORKDAYS.INTL', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'NUMBERVALUE',
  'PDURATION', 'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTOF', 'PERCENTRANK.EXC', 'PERCENTRANK.INC',
  'PERMUTATIONA', 'PHI', 'PIVOTBY', 'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC', 'QUERYSTRING',
  'RANDARRAY', 'RANK.AVG', 'RANK.EQ', 'REDUCE', 'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST', 'RRI', 'SCAN',
  'SEC', 'SECH', 'SEQUENCE', 'SHEET', 'SHEETS', 'SKEW.P', 'SORT', 'SORTBY', 'STDEV.P', 'STDEV.S', 'SWITCH',
  'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'TAKE', 'TEXTAFTER', 'TEXTBEFORE',
  'TEXTJOIN', 'TEXTSPLIT', 'TOCOL', 'TOROW', 'TRIMRANGE', 'UNICHAR', 'UNICODE', 'UNIQUE', 'VALUETOTEXT',
  'VAR.P', 'VAR.S', 'VSTACK', 'WEBSERVICE', 'WEIBULL.DIST', 'WRAPCOLS', 'WRAPROWS', 'XLOOKUP', 'XMATCH',
  'XOR', 'Z.TEST', 'ANCHORARRAY', 'SINGLE',
])

const PREFIX_RE = /^(?:_xlfn\.|_xlws\.|_xlpm\.)+/i

function significant(tokens: EditorToken[]) {
  return tokens.filter((token) => token.kind !== 'space')
}

/** Index of the token closing the parenthesis opened at `openIndex`. */
function matchingClose(tokens: EditorToken[], openIndex: number) {
  let depth = 0
  for (let index = openIndex; index < tokens.length; index += 1) {
    if (tokens[index].kind === 'open') depth += 1
    else if (tokens[index].kind === 'close') {
      depth -= 1
      if (depth === 0) return index
    }
  }
  return -1
}

/** Names declared by LET/LAMBDA in the formula (upper-cased, without prefixes). */
function declaredLocalNames(tokens: EditorToken[]) {
  const names = new Set<string>()
  const list = significant(tokens)
  for (let index = 0; index < list.length; index += 1) {
    const token = list[index]
    if (token.kind !== 'function') continue
    const name = token.text.replace(PREFIX_RE, '').toUpperCase()
    if (name !== 'LET' && name !== 'LAMBDA') continue
    const open = index + 1
    if (list[open]?.kind !== 'open') continue
    const close = matchingClose(list, open)
    if (close < 0) continue
    // Split top-level arguments.
    const args: EditorToken[][] = [[]]
    let depth = 0
    for (let cursor = open + 1; cursor < close; cursor += 1) {
      const item = list[cursor]
      if (item.kind === 'open') depth += 1
      if (item.kind === 'close') depth -= 1
      if (item.kind === 'separator' && depth === 0) args.push([])
      else args[args.length - 1].push(item)
    }
    const candidates = name === 'LET' ? args.filter((_arg, position) => position % 2 === 0 && position < args.length - 1) : args.slice(0, -1)
    for (const arg of candidates) {
      if (arg.length === 1 && (arg[0].kind === 'name' || arg[0].kind === 'reference')) names.add(arg[0].text.replace(PREFIX_RE, '').toUpperCase())
    }
  }
  return names
}

/** Friendly editor form of a formula stored in a workbook file (no leading "="). */
export function fromFileFormula(formula: string): string {
  if (!/_xl(?:fn|ws|pm)\./i.test(formula)) return formula
  const text = `=${formula}`
  const tokens = tokenizeFormulaText(text)
  let output = ''
  let last = 0
  const list = tokens
  for (let index = 0; index < list.length; index += 1) {
    const token = list[index]
    if ((token.kind === 'function' || token.kind === 'name') && PREFIX_RE.test(token.text)) {
      const bare = token.text.replace(PREFIX_RE, '')
      const upper = bare.toUpperCase()
      // _xlfn.ANCHORARRAY(A1) -> A1#
      if (token.kind === 'function' && upper === 'ANCHORARRAY') {
        const rest = significant(list.slice(index + 1))
        if (rest[0]?.kind === 'open' && rest[1]?.kind === 'reference' && rest[2]?.kind === 'close') {
          output += text.slice(last, token.start) + rest[1].text + '#'
          last = rest[2].end
          index = list.indexOf(rest[2])
          continue
        }
      }
      // _xlfn.SINGLE(x) -> @x for a single reference, name, or call argument.
      if (token.kind === 'function' && upper === 'SINGLE') {
        const openIndex = list.findIndex((item, position) => position > index && item.kind === 'open')
        const closeIndex = openIndex >= 0 ? matchingClose(list, openIndex) : -1
        if (closeIndex > openIndex) {
          const inner = text.slice(list[openIndex].end, list[closeIndex].start).trim()
          const innerTokens = significant(tokenizeFormulaText(`=${inner}`).slice(1))
          const simple = innerTokens.length === 1 && ['reference', 'name'].includes(innerTokens[0].kind)
          const call = innerTokens[0]?.kind === 'function' && innerTokens[1]?.kind === 'open' && matchingClose(innerTokens, 1) === innerTokens.length - 1
          if (simple || call) {
            output += text.slice(last, token.start) + '@' + fromFileFormula(inner)
            last = list[closeIndex].end
            index = closeIndex
            continue
          }
        }
      }
      output += text.slice(last, token.start) + bare
      last = token.end
    }
  }
  output += text.slice(last)
  return output.slice(1)
}

/** File form of an editor formula (no leading "="): adds the prefixes Excel requires. */
export function toFileFormula(formula: string): string {
  const text = `=${formula}`
  const tokens = tokenizeFormulaText(text)
  const locals = declaredLocalNames(tokens)
  let output = ''
  let last = 0
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index]
    if (token.kind === 'function') {
      if (PREFIX_RE.test(token.text)) continue
      const upper = token.text.toUpperCase()
      if (locals.has(upper)) {
        output += text.slice(last, token.start) + `_xlpm.${token.text}`
        last = token.end
      } else if (FUTURE_FUNCTIONS.has(upper)) {
        output += text.slice(last, token.start) + (WORKSHEET_NAMESPACE.has(upper) ? '_xlfn._xlws.' : '_xlfn.') + token.text
        last = token.end
      }
      continue
    }
    if ((token.kind === 'name' || token.kind === 'reference') && locals.has(token.text.toUpperCase()) && !token.text.includes('!') && !token.text.includes(':')) {
      output += text.slice(last, token.start) + `_xlpm.${token.text}`
      last = token.end
      continue
    }
    if (token.kind === 'reference' && token.reference?.spill) {
      output += text.slice(last, token.start) + `_xlfn.ANCHORARRAY(${token.text.slice(0, -1)})`
      last = token.end
      continue
    }
    if (token.kind === 'operator' && token.text === '@') {
      // @operand -> _xlfn.SINGLE(operand) for the following reference/name/call.
      const rest = tokens.slice(index + 1)
      const firstIndex = rest.findIndex((item) => item.kind !== 'space')
      const first = rest[firstIndex]
      if (!first) continue
      let endToken: EditorToken | undefined = first
      if (first.kind === 'function') {
        const openIndex = index + 1 + firstIndex + 1
        const closeIndex = tokens[openIndex]?.kind === 'open' ? matchingClose(tokens, openIndex) : -1
        endToken = closeIndex > 0 ? tokens[closeIndex] : undefined
      } else if (first.kind !== 'reference' && first.kind !== 'name') endToken = undefined
      if (!endToken) continue
      const operand = text.slice(first.start, endToken.end)
      output += text.slice(last, token.start) + `_xlfn.SINGLE(${toFileFormula(operand)})`
      last = endToken.end
      index = tokens.indexOf(endToken)
    }
  }
  output += text.slice(last)
  return output.slice(1)
}

interface FormulaCellLike { formula?: string }
interface SheetLike<C extends FormulaCellLike> { cells: Record<string, C> }
interface WorkbookLike<C extends FormulaCellLike, S extends SheetLike<C>> {
  sheets: S[]
  definedNames?: Array<{ ranges?: string[]; ref?: string }>
}

function mapFormulas<C extends FormulaCellLike, S extends SheetLike<C>, W extends WorkbookLike<C, S>>(workbook: W, convert: (formula: string) => string): W {
  let changed = false
  const sheets = workbook.sheets.map((sheet) => {
    let cells: Record<string, C> | null = null
    for (const address in sheet.cells) {
      const cell = sheet.cells[address]
      if (!cell?.formula) continue
      const next = convert(cell.formula)
      if (next === cell.formula) continue
      if (!cells) cells = { ...sheet.cells }
      cells[address] = { ...cell, formula: next }
    }
    if (!cells) return sheet
    changed = true
    return { ...sheet, cells }
  })
  const definedNames = workbook.definedNames?.map((name) => {
    const ranges = name.ranges?.map((range) => (range.startsWith('=') ? `=${convert(range.slice(1))}` : convert(range)))
    if (!ranges || ranges.every((range, index) => range === name.ranges![index])) return name
    changed = true
    return { ...name, ranges }
  })
  return changed ? { ...workbook, sheets, ...(definedNames ? { definedNames } : {}) } : workbook
}

/** Workbook with every formula in Excel's file form (for saving/exporting). */
export function toFileWorkbook<W extends WorkbookLike<FormulaCellLike, SheetLike<FormulaCellLike>>>(workbook: W): W {
  return mapFormulas(workbook, toFileFormula)
}

/** Workbook with every formula in the friendly editor form (after opening). */
export function fromFileWorkbook<W extends WorkbookLike<FormulaCellLike, SheetLike<FormulaCellLike>>>(workbook: W): W {
  return mapFormulas(workbook, fromFileFormula)
}
