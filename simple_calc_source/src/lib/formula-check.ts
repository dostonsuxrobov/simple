/**
 * Excel syntax this app does not calculate but Excel accepts: a typed formula using it is
 * entered as written (it shows an error here and calculates in Excel) instead of being refused
 * with "There's a problem with this formula".
 *
 * - the intersection operator, a space between two references: =A1:C3 B2:B9
 * - a union in parentheses: =SUM((A1:A2,C1:C2))
 * - a reference to another workbook: =[Budget.xlsx]Sheet1!A1, ='C:\x\[B.xlsx]S'!A1
 */
import { tokenizeFormulaText } from './formula-editing'
import type { EditorToken } from './formula-editing'

function isReferenceLike(token: EditorToken | undefined) {
  return Boolean(token && (token.kind === 'reference' || token.kind === 'name'))
}

export function usesExcelOnlySyntax(draft: string): boolean {
  if (!draft.startsWith('=')) return false
  // Another workbook: "[file]" starting a sheet reference (not a table's [Column]), outside
  // text in quotes.
  const unquoted = draft.replace(/"(?:[^"]|"")*"?/g, '""')
  if (/(?:^=|[=(,;+\-*/&^<>'\s\\])\[[^\]\r\n]+\][^!"(),+\-*/&^<>=\r\n]*!/.test(unquoted)) return true
  const tokens = tokenizeFormulaText(draft)
  for (let index = 1; index < tokens.length - 1; index += 1) {
    // Intersection: reference, whitespace, reference.
    if (tokens[index].kind === 'space' && isReferenceLike(tokens[index - 1]) && isReferenceLike(tokens[index + 1])) return true
  }
  // Union: "(" not opening a function call, holding only references separated by commas.
  for (let index = 0; index < tokens.length; index += 1) {
    if (tokens[index].kind !== 'open' || tokens[index - 1]?.kind === 'function') continue
    let depth = 0
    let separators = 0
    let onlyReferences = true
    for (let inner = index + 1; inner < tokens.length; inner += 1) {
      const token = tokens[inner]
      if (token.kind === 'open') { depth += 1; onlyReferences = false; continue }
      if (token.kind === 'close') {
        if (depth === 0) {
          if (separators > 0 && onlyReferences) return true
          break
        }
        depth -= 1
        continue
      }
      if (depth > 0) continue
      if (token.kind === 'separator') separators += 1
      else if (token.kind !== 'space' && !isReferenceLike(token)) onlyReferences = false
    }
  }
  return false
}
