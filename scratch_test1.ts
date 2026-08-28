import { evaluateFormula, shiftFormulaReferences } from './src/lib/formulas.ts'
import type { FormulaEvaluationHooks } from './src/lib/formulas.ts'

const cells: Record<string, number | string | boolean | null> = { A1: 1, A2: 2, A3: 3, B1: 'x' }
const resolver = (_s: string, a: string) => cells[a] ?? null
const f = (s: string) => evaluateFormula(s, 'Sheet1', resolver)

// shift whole ranges off-grid
console.log('shift A:A left ->', JSON.stringify(shiftFormulaReferences('SUM(A:A)', 0, -1)))
console.log('eval of shifted:', f('SUM(#REF!:#REF!)'))
console.log('shift $A:B left ->', JSON.stringify(shiftFormulaReferences('SUM($A:B)', 0, -1)))
console.log('eval:', f('SUM($A:#REF!)'))
console.log('shift 1:2 up ->', JSON.stringify(shiftFormulaReferences('SUM(1:2)', -1, 0)))
console.log('baseline cell off-grid ->', JSON.stringify(shiftFormulaReferences('SUM(A1:B2)', -1, 0)))
console.log('eval baseline off-grid:', f('SUM(#REF!:B1)'))

// range compare / criterion
console.log('COUNTIF blank crit:', f('COUNTIF(A1:A3,B2)'))

// name with paren: shift function-looking
console.log('shift LOG10(A1):', shiftFormulaReferences('LOG10(A1)', 1, 1))
console.log('shift AB1(x) fn-like:', shiftFormulaReferences('AB1(A1)', 1, 1))

// sparse ordering: TEXTJOIN over >100k sparse
const sparseCells: Record<string, number | string> = { A200000: 'later', A1: 'first' }
// insertion order: A200000 inserted before A1
const sparseResolver = (_s: string, a: string) => sparseCells[a] ?? null
const sparseHooks: FormulaEvaluationHooks = {
  getUsedRange: () => ({ maxRow: 200000, maxCol: 2 }),
  forEachCellInRange: (_s, b, visit) => {
    for (const key of Object.keys(sparseCells)) {
      const m = /^([A-Z])(\d+)$/.exec(key)!
      const col = m[1].charCodeAt(0) - 64, row = Number(m[2])
      if (row >= b.startRow && row <= b.endRow && col >= b.startColumn && col <= b.endColumn) visit(row, col)
    }
  },
}
console.log('TEXTJOIN sparse order:', evaluateFormula('TEXTJOIN(",",TRUE,A:A)', 'S', sparseResolver, sparseHooks))
console.log('CONCAT sparse order:', evaluateFormula('CONCAT(A:A)', 'S', sparseResolver, sparseHooks))
