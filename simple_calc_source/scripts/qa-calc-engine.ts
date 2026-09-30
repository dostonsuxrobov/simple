import assert from 'node:assert/strict'
import { produce } from 'immer'
import { CalculationEngine, markArrayMembers } from '../src/lib/calc-engine'
import type { CellData, SheetData, WorkbookModel } from '../src/spreadsheet-types'

function sheet(id: string, name: string, cells: Record<string, CellData>, extra: Partial<SheetData> = {}): SheetData {
  return { id, name, rowCount: 200, colCount: 40, cells, merges: [], colWidths: {}, rowHeights: {}, ...extra }
}
function workbook(...sheets: SheetData[]): WorkbookModel {
  return { version: 1, name: 'test.xlsx', activeSheetId: sheets[0].id, sheets }
}
const f = (formula: string): CellData => ({ formula })
const v = (value: CellData['value']): CellData => ({ value })

// ---- Basic evaluation and incremental invalidation ------------------------------------------
let book = workbook(
  sheet('s1', 'Sheet1', { A1: v(1), A2: f('A1*2'), A3: f('SUM(A1:A2)'), B1: v(10), B2: f('B1+1'), C1: f('Data!A1*3') }),
  sheet('s2', 'Data', { A1: v(7) }),
)
const engine = new CalculationEngine(book)
assert.equal(engine.getValue('s1', 'A2'), 2)
assert.equal(engine.getValue('s1', 'A3'), 3)
assert.equal(engine.getValue('s1', 'B2'), 11)
assert.equal(engine.getValue('s1', 'C1'), 21)
assert.equal(engine.getValue('s1', 'Z99'), null)

const edit = (mutator: (draft: WorkbookModel) => void) => {
  book = produce(book, mutator)
  engine.update(book)
}
edit((draft) => { draft.sheets[0].cells.A1 = v(5) })
assert.equal(engine.getValue('s1', 'A2'), 10)
assert.equal(engine.getValue('s1', 'A3'), 15)
assert.equal(engine.getValue('s1', 'B2'), 11, 'unrelated formula keeps its value')
edit((draft) => { draft.sheets[1].cells.A1 = v(2) })
assert.equal(engine.getValue('s1', 'C1'), 6, 'cross-sheet dependency recalculates')
edit((draft) => { draft.sheets[0].cells.A2 = f('A1*100') })
assert.equal(engine.getValue('s1', 'A3'), 505, 'changed formula propagates')
edit((draft) => { delete draft.sheets[0].cells.A1 })
assert.equal(engine.getValue('s1', 'A3'), 0)

// Circular references are reported, and recover when broken.
edit((draft) => { draft.sheets[0].cells.D1 = f('D2+1'); draft.sheets[0].cells.D2 = f('D1+1') })
assert.equal(engine.getValue('s1', 'D1'), '#CIRC!')
edit((draft) => { draft.sheets[0].cells.D2 = v(4) })
assert.equal(engine.getValue('s1', 'D1'), 5)

// Sheet rename resets caches and keeps references by name working after the formula is updated.
edit((draft) => { draft.sheets[1].name = 'Inputs'; draft.sheets[0].cells.C1 = f('Inputs!A1*3') })
assert.equal(engine.getValue('s1', 'C1'), 6)

// Whole-column references see rows added far below.
edit((draft) => { draft.sheets[0].cells.E1 = f('SUM(F:F)'); draft.sheets[0].cells.F3 = v(4) })
assert.equal(engine.getValue('s1', 'E1'), 4)
edit((draft) => { draft.sheets[0].cells.F5000 = v(6) })
assert.equal(engine.getValue('s1', 'E1'), 10)

// ---- Dynamic arrays -------------------------------------------------------------------------
edit((draft) => {
  draft.sheets[0].cells.H1 = f('SEQUENCE(3)')
  draft.sheets[0].cells.I1 = f('SUM(H1#)')
  draft.sheets[0].cells.J1 = f('H2*10')
})
assert.equal(engine.getValue('s1', 'H1'), 1)
assert.equal(engine.getValue('s1', 'H2'), 2)
assert.equal(engine.getValue('s1', 'H3'), 3)
assert.equal(engine.getValue('s1', 'H4'), null)
assert.equal(engine.getValue('s1', 'I1'), 6)
assert.equal(engine.getValue('s1', 'J1'), 20, 'plain references read spilled cells')
assert.deepEqual(engine.spillRange('s1', 'H1'), { top: 0, left: 7, bottom: 2, right: 7 })
assert.equal(engine.spillAnchorOf('s1', 'H3'), 'H1')

edit((draft) => { draft.sheets[0].cells.H1 = f('SEQUENCE(4)') })
assert.equal(engine.getValue('s1', 'I1'), 10, 'spill resize propagates to A1# readers')
assert.equal(engine.getValue('s1', 'H4'), 4)

edit((draft) => { draft.sheets[0].cells.H3 = v('blocker') })
assert.equal(engine.getValue('s1', 'H1'), '#SPILL!')
assert.equal(engine.getValue('s1', 'H2'), null)
assert.equal(engine.getValue('s1', 'J1'), 0)
assert.equal(engine.getValue('s1', 'I1'), '#REF!')
edit((draft) => { delete draft.sheets[0].cells.H3 })
assert.equal(engine.getValue('s1', 'H1'), 1, 'clearing the blocker restores the spill')
assert.equal(engine.getValue('s1', 'J1'), 20)

// Merged cells block spills.
edit((draft) => { draft.sheets[0].merges = ['H2:H3'] })
assert.equal(engine.getValue('s1', 'H1'), '#SPILL!')
edit((draft) => { draft.sheets[0].merges = [] })
assert.equal(engine.getValue('s1', 'H1'), 1)

// A formula that is not array-shaped never spills.
edit((draft) => { draft.sheets[0].cells.K1 = f('SUM(SEQUENCE(4))') })
assert.equal(engine.getValue('s1', 'K1'), 10)
assert.equal(engine.getValue('s1', 'K2'), null)

// Saved array ranges: cached member values yield to the live spill.
const saved = markArrayMembers(sheet('s3', 'Saved', {
  A1: v(3), A2: v(1), A3: v(2),
  B1: { formula: 'SORT(A1:A3)', formulaType: 'array', formulaRange: 'B1:B3', dynamicFormula: true, result: 99 },
  B2: v(99), B3: v(99),
}))
assert.equal(saved.cells.B2.arrayMember, 'B1')
const savedEngine = new CalculationEngine(workbook(saved))
assert.deepEqual(['B1', 'B2', 'B3'].map((address) => savedEngine.getValue('s3', address)), [1, 2, 3])

// Unsupported anchors fall back to cached values for the whole saved range.
const unsupported = markArrayMembers(sheet('s4', 'Legacy', {
  B1: { formula: 'MYADDIN(A1:A3)', formulaType: 'array', formulaRange: 'B1:B2', result: 5 },
  B2: v(6),
}))
const unsupportedEngine = new CalculationEngine(workbook(unsupported))
assert.equal(unsupportedEngine.getValue('s4', 'B1'), 5)
assert.equal(unsupportedEngine.getValue('s4', 'B2'), 6)
assert.equal(unsupportedEngine.isStale('s4', 'B1'), true)

// withResults writes the spill range and member values for saving.
const results = engine.withResults()
const resultSheet = results.sheets[0]
assert.equal(resultSheet.cells.H1.formulaRange, 'H1:H4')
assert.equal(resultSheet.cells.H1.dynamicFormula, true)
assert.equal(resultSheet.cells.H3.value, 3)
assert.equal(resultSheet.cells.H3.arrayMember, 'H1')
assert.equal(resultSheet.cells.A3.result, 0)

// ---- Tables, filtered rows, volatile functions -----------------------------------------------
let tableBook = workbook(sheet('t1', 'Sales', {
  A1: v('Region'), B1: v('Units'), A2: v('North'), B2: v(5), A3: v('South'), B3: v(7), A4: v('East'), B4: v(1),
  D1: f('SUM(Orders[Units])'), D2: f('SUBTOTAL(9, B2:B4)'), D3: f('SUBTOTAL(109, B2:B4)'), D4: f('SUM(B2:B4)'),
  E1: f('NOW()'), E2: f('Orders[@Units]*2'),
}, {
  tables: [{ id: 'tb1', name: 'Orders', ref: 'A1:B4', headerRow: true, totalsRow: false, columns: [{ name: 'Region' }, { name: 'Units' }] }],
}))
const tableEngine = new CalculationEngine(tableBook)
assert.equal(tableEngine.getValue('t1', 'D1'), 13)
assert.equal(tableEngine.getValue('t1', 'E2'), 10, '[@Units] reads the row of the formula')
tableBook = produce(tableBook, (draft) => { draft.sheets[0].hiddenRows = [3]; draft.sheets[0].filteredRows = [3] })
tableEngine.update(tableBook)
assert.equal(tableEngine.getValue('t1', 'D2'), 6, 'SUBTOTAL skips filtered rows')
assert.equal(tableEngine.getValue('t1', 'D4'), 13, 'SUM still counts filtered rows')
tableBook = produce(tableBook, (draft) => { draft.sheets[0].hiddenRows = [3, 4]; draft.sheets[0].filteredRows = [3] })
tableEngine.update(tableBook)
assert.equal(tableEngine.getValue('t1', 'D2'), 6, 'SUBTOTAL 9 includes manually hidden rows')
assert.equal(tableEngine.getValue('t1', 'D3'), 5, 'SUBTOTAL 109 excludes manually hidden rows')

// ---- Performance: 10k running totals ---------------------------------------------------------
const cells: Record<string, CellData> = {}
for (let row = 1; row <= 10_000; row += 1) {
  cells[`B${row}`] = v(1)
  cells[`C${row}`] = f(`SUM($B$1:B${row})`)
  cells[`D${row}`] = f(`B${row}*2`)
}
let big = workbook(sheet('p1', 'Perf', cells))
const perf = new CalculationEngine(big)
let started = performance.now()
for (let row = 1; row <= 10_000; row += 1) perf.getValue('p1', `D${row}`)
assert.equal(perf.getValue('p1', 'C10000'), 10_000)
const initial = performance.now() - started
started = performance.now()
big = produce(big, (draft) => { draft.sheets[0].cells.B9000 = v(5) })
perf.update(big, new Map([['p1', new Set(['B9000'])]]))
assert.equal(perf.getValue('p1', 'C10000'), 10_004)
assert.equal(perf.getValue('p1', 'C8999'), 8_999)
assert.equal(perf.getValue('p1', 'D9000'), 10)
const incremental = performance.now() - started
started = performance.now()
big = produce(big, (draft) => { draft.sheets[0].cells.B20 = v(3) })
perf.update(big)
assert.equal(perf.getValue('p1', 'D20'), 6)
const diffed = performance.now() - started
console.log(`  perf: first calc ${initial.toFixed(0)} ms, hinted edit ${incremental.toFixed(0)} ms, diffed edit ${diffed.toFixed(0)} ms`)
assert.ok(incremental < 400, `incremental recalculation took ${incremental} ms`)

// ---- Candidate values for data validation --------------------------------------------------
{
  const book = workbook(sheet('v1', 'Rules', { A1: v('x'), A2: v('y'), A3: v('w'), B1: f('A1&"!"'), C1: f('SUM(D1:D3)'), D1: v(1), D2: f('D1*2') }))
  const rules = new CalculationEngine(book)
  assert.equal(rules.getValue('v1', 'B1'), 'x!')
  assert.equal(rules.getValue('v1', 'C1'), 3)
  const unique = (value: string) => rules.evaluateAt('v1', 'COUNTIF($A$1:$A$10,A4)=1', 3, 0, { row: 3, col: 0, value })
  assert.equal(unique('x'), false, 'a duplicate entry in a blank cell is counted')
  assert.equal(unique('z'), true)
  assert.equal(rules.evaluateAt('v1', 'B1="q!"', 0, 1, { row: 0, col: 0, value: 'q' }), true, 'cached dependents see the candidate')
  assert.equal(rules.evaluateAt('v1', 'C1', 0, 2, { row: 0, col: 3, value: 10 }), 30, 'transitive dependents recalculate')
  assert.equal(rules.getValue('v1', 'B1'), 'x!', 'nothing computed under an override is cached')
  assert.equal(rules.getValue('v1', 'C1'), 3)
  assert.equal(rules.getValue('v1', 'D2'), 2)
  assert.equal(rules.evaluateAt('v1', 'LEN(A5)<3', 4, 0, { row: 4, col: 0, value: 'abcd' }), false)
}

console.log('Calculation engine QA passed: incremental recalculation, spills, saved arrays, tables, visibility, scaling, and validation candidates.')
