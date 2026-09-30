import assert from 'node:assert/strict'
import {
  evaluateFormula,
  evaluateFormulaDetailed,
  formulaMayReturnArray,
  type FormulaEvaluationHooks,
  type FormulaTableInfo,
} from '../src/lib/formulas.ts'

// A small sheet: A1:A5 = 1..5, B1:B5 = 10..50, C1 = "x", D1:D3 = names.
const cells: Record<string, unknown> = {
  A1: 1, A2: 2, A3: 3, A4: 4, A5: 5,
  B1: 10, B2: 20, B3: 30, B4: 40, B5: 50,
  C1: 'x', D1: 'North', D2: 'South', D3: 'North',
  // A table at F1:H4 with header row: Region | Units | Price
  F1: 'Region', G1: 'Units', H1: 'Price',
  F2: 'North', G2: 3, H2: 2.5,
  F3: 'South', G3: 5, H3: 4,
  F4: 'East', G4: 2, H4: 10,
}
const resolver = (_sheet: string, address: string) => cells[address] as never
const table: FormulaTableInfo = {
  sheetId: 'Sheet1', name: 'Sales', startRow: 1, endRow: 4, startColumn: 6, endColumn: 8,
  headerRowCount: 1, totalsRowCount: 0, columns: ['Region', 'Units', 'Price'],
}
const tracked: string[] = []
let volatile = 0
const hooks = (row = 1, column = 1): FormulaEvaluationHooks => ({
  currentCell: { row, column },
  resolveTable: (name) => (name === null || name.toLowerCase() === 'sales' ? table : null),
  resolveSpill: (_sheet, r, c) => (r === 1 && c === 10 ? { startRow: 1, endRow: 3, startColumn: 10, endColumn: 10 } : null),
  trackRange: (sheet, b) => { tracked.push(`${sheet}:${b.startRow},${b.startColumn}:${b.endRow},${b.endColumn}`) },
  markVolatile: () => { volatile += 1 },
  resolveDefinedName: (name) => ({ DOUBLE: '=LAMBDA(x, x * 2)', RATE: '0.5' } as Record<string, string>)[name.toUpperCase()],
  getSheetNames: () => ['Sheet1', 'Data'],
  getSheetName: (id) => id,
})
const run = (formula: string, row?: number, column?: number) => evaluateFormula(formula, 'Sheet1', resolver, hooks(row, column))
const array = (formula: string, row?: number, column?: number) => {
  const result = evaluateFormulaDetailed(formula, 'Sheet1', resolver, hooks(row, column))
  return result.array ? { rows: result.array.rowCount, columns: result.array.columnCount, values: result.array.values } : result.value
}

// Comparison semantics match Excel.
assert.equal(run('1="1"'), false)
assert.equal(run('"abc"="ABC"'), true)
assert.equal(run('"a"<"b"'), true)
assert.equal(run('1<"a"'), true)
assert.equal(run('"a"<TRUE'), true)
assert.equal(run('Z99=0'), true)
assert.equal(run('Z99=""'), true)
// Text conversion keeps 15 significant digits.
assert.equal(run('0.1+0.2&""'), '0.3')
assert.equal(run('1/3&""'), '0.333333333333333')
assert.equal(run('10^15&""'), '1E+15')
assert.equal(run('"$1,200"+1'), 1201)
assert.equal(run('"15%"*2'), 0.3)

// LET / LAMBDA / helpers.
assert.equal(run('LET(x, 2, y, x*3, x+y)'), 8)
assert.equal(run('LET(_xlpm.x, 5, _xlpm.x*2)'), 10)
assert.equal(run('LAMBDA(a, b, a+b)(2, 3)'), 5)
assert.equal(run('LET(f, LAMBDA(n, n*n), f(4))'), 16)
assert.equal(run('DOUBLE(21)'), 42)
assert.equal(run('RATE*4'), 2)
assert.equal(run('LAMBDA(a, b, IF(ISOMITTED(b), a, a+b))(7)'), 7)
assert.equal(run('LAMBDA(x, x)'), '#CALC!')
assert.deepEqual(array('MAP(A1:A3, LAMBDA(v, v*10))'), { rows: 3, columns: 1, values: [10, 20, 30] })
assert.equal(run('REDUCE(0, A1:A5, LAMBDA(acc, v, acc+v))'), 15)
assert.deepEqual(array('SCAN(0, A1:A4, LAMBDA(acc, v, acc+v))'), { rows: 4, columns: 1, values: [1, 3, 6, 10] })
assert.deepEqual(array('BYROW(A1:B2, LAMBDA(r, SUM(r)))'), { rows: 2, columns: 1, values: [11, 22] })
assert.deepEqual(array('BYCOL(A1:B2, LAMBDA(c, SUM(c)))'), { rows: 1, columns: 2, values: [3, 30] })
assert.deepEqual(array('MAKEARRAY(2, 3, LAMBDA(r, c, r*c))'), { rows: 2, columns: 3, values: [1, 2, 3, 2, 4, 6] })

// Lifting and broadcasting.
assert.deepEqual(array('LEN(D1:D3)'), { rows: 3, columns: 1, values: [5, 5, 5] })
assert.deepEqual(array('A1:A3*B1:B3'), { rows: 3, columns: 1, values: [10, 40, 90] })
assert.deepEqual(array('SEQUENCE(2)*SEQUENCE(1,3)'), { rows: 2, columns: 3, values: [1, 2, 3, 2, 4, 6] })
assert.deepEqual(array('IF(A1:A4>2, "big", "small")'), { rows: 4, columns: 1, values: ['small', 'small', 'big', 'big'] })
assert.deepEqual(array('IFERROR(1/(A1:A3-2), "none")'), { rows: 3, columns: 1, values: [-1, 'none', 1] })
assert.deepEqual(array('ROW(A2:A4)'), { rows: 3, columns: 1, values: [2, 3, 4] })
assert.equal(run('SUMPRODUCT(--(MOD(ROW(A1:A5),2)=0), B1:B5)'), 60)
assert.deepEqual(array('MATCH({3,1}, A1:A5, 0)'), { rows: 1, columns: 2, values: [3, 1] })
assert.equal(run('SUM(COUNTIF(D1:D3, {"North","South"}))'), 3)
assert.equal(run('LEN(A1:A1)'), 1)

// References computed from functions.
assert.equal(run('SUM(A1:INDEX(A1:A5, 3))'), 6)
assert.equal(run('SUM(INDEX(A1:B5, 0, 2))'), 150)
assert.equal(run('ROWS(A2:INDEX(A:A, 5))'), 4)
assert.equal(run('SUM(OFFSET(A1, 1, 0, 3, 1))'), 9)
assert.equal(run('SUM(OFFSET(A5, 0, 0, -2, 1))'), 9)
assert.equal(run('SUM(INDIRECT("A1:A" & 4))'), 10)
assert.equal(run('INDIRECT("R2C2", FALSE)'), 20)
assert.equal(run('ISREF(A1)'), true)
assert.equal(run('ISREF(1)'), false)
assert.ok(volatile > 0, 'OFFSET/INDIRECT mark the formula volatile')

// Structured references.
assert.equal(run('SUM(Sales[Units])'), 10)
assert.equal(run('SUMPRODUCT(Sales[Units], Sales[Price])'), 7.5 + 20 + 20)
assert.equal(run('COUNTA(Sales[#Headers])'), 3)
assert.equal(run('ROWS(Sales[#All])'), 4)
assert.equal(run('[@Units]*[@Price]', 3, 9), 20)
assert.equal(run('Sales[@Region]', 4, 9), 'East')
assert.equal(run('SUM(Sales[[Units]:[Price]])'), 26.5)
assert.equal(run('Missing[Units]'), '#REF!')

// Spill references and implicit intersection.
cells.J1 = 7; cells.J2 = 8; cells.J3 = 9
assert.equal(run('SUM(J1#)'), 24)
assert.equal(run('SUM(ANCHORARRAY(J1))'), 24)
assert.equal(run('SUM(A1#)'), '#REF!')
assert.equal(run('@A1:A5', 3, 3), 3)
assert.equal(run('SINGLE(B1:B5)', 2, 5), 20)

// Array detection for spill anchors.
assert.equal(formulaMayReturnArray('SUM(A1:A5)'), false)
assert.equal(formulaMayReturnArray('A1:A5'), true)
assert.equal(formulaMayReturnArray('FILTER(A1:A5, A1:A5>2)'), true)
assert.equal(formulaMayReturnArray('A1+1'), false)
assert.equal(formulaMayReturnArray('LEN(A1:A3)'), true)
assert.equal(formulaMayReturnArray('LEN(A1)'), false)

// Dependency tracking covers whole columns without clamping.
tracked.length = 0
run('SUM(A:A)')
assert.ok(tracked.includes('Sheet1:1,1:1048576,1'), 'A:A is tracked as the full column')

console.log('Formula engine QA passed: comparisons, text conversion, LET/LAMBDA, lifting, references, tables, and spills.')
