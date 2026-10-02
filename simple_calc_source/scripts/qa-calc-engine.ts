import assert from 'node:assert/strict'
import { produce } from 'immer'
import { CalculationEngine, calculationOptionsOf, markArrayMembers, withCalculationOptions } from '../src/lib/calc-engine'
import { diagnoseFormula } from '../src/lib/formulas'
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

// ---- Regression cases: calc-formula-engine-2/3/4/5 ------------------------------------------
{
  // Whole-column and whole-row references reach spilled array members (calc-formula-engine-2).
  let spillBook = workbook(
    sheet('w1', 'Sheet1', { A1: f('SEQUENCE(5)'), C1: f('SUM(A:A)'), C2: f('COUNT(A:A)'), E1: f('MAX(A:A)'), D1: f('SUM(A1:A10)'), G2: f('COUNTA(1:1)') }),
    sheet('w2', 'Other', { A1: f('SUM(Sheet1!A:A)'), A2: f('MATCH(5,Sheet1!A:A,0)') }),
  )
  const spills = new CalculationEngine(spillBook)
  assert.deepEqual(['C1', 'C2', 'E1', 'D1'].map((address) => spills.getValue('w1', address)), [15, 5, 5, 15])
  assert.deepEqual(['A1', 'A2'].map((address) => spills.getValue('w2', address)), [15, 5])
  spillBook = produce(spillBook, (draft) => { draft.sheets[0].cells.A1 = f('SEQUENCE(8)') })
  spills.update(spillBook)
  assert.equal(spills.getValue('w1', 'C1'), 36, 'a growing spill widens whole-column readers')
  assert.equal(spills.getValue('w2', 'A1'), 36)
  spillBook = produce(spillBook, (draft) => { draft.sheets[0].cells.J1 = f('SEQUENCE(1,4)') })
  spills.update(spillBook)
  assert.equal(spills.getValue('w1', 'G2'), 8, 'whole-row references reach a row spill')
  // Sparse (very large) whole-column ranges visit spilled members too.
  const sparseBook = workbook(sheet('w3', 'Big', { A1: f('SEQUENCE(4)'), A200000: v(10), C1: f('SUM(A:B)') }))
  assert.equal(new CalculationEngine(sparseBook).getValue('w3', 'C1'), 20)

  // Text that starts with "=" stays text (calc-formula-engine-3).
  const textBook = workbook(sheet('x1', 'Sheet1', {
    A1: v('=== Q1 ==='), A2: v(10), A3: v(20), B1: f('SUM(A1:A3)'), B2: f('A1&" total"'),
    C1: f('SUM(D1:D2)'), D1: v(3), D2: v(4), E1: f('FORMULATEXT(C1)'), E2: f('E1'), E3: f('LEN(E1)'),
    G1: v('=1+1'), G2: f('ISTEXT(G1)'), G3: f('G1'),
  }))
  const text = new CalculationEngine(textBook)
  assert.deepEqual(['B1', 'B2', 'E1', 'E2', 'E3', 'G2', 'G3'].map((address) => text.getValue('x1', address)),
    [30, '=== Q1 === total', '=SUM(D1:D2)', '=SUM(D1:D2)', 11, true, '=1+1'])

  // SUBTOTAL/AGGREGATE on another sheet follow hidden rows; ISFORMULA tracks its target
  // (calc-formula-engine-4).
  let visibilityBook = workbook(
    sheet('y1', 'Data', { A1: v(1), A2: v(2), A3: v(3), B1: f('SUBTOTAL(109,A1:A3)'), D1: v(5), C1: f('ISFORMULA(D1)') }),
    sheet('y2', 'Summary', { A1: f('SUBTOTAL(109,Data!A1:A3)'), A2: f('AGGREGATE(9,5,Data!A1:A3)'), A3: f('A1*10') }),
  )
  const visibility = new CalculationEngine(visibilityBook)
  assert.deepEqual(['A1', 'A2', 'A3'].map((address) => visibility.getValue('y2', address)), [6, 6, 60])
  assert.equal(visibility.getValue('y1', 'C1'), false)
  visibilityBook = produce(visibilityBook, (draft) => { draft.sheets[0].hiddenRows = [2] })
  visibility.update(visibilityBook)
  assert.equal(visibility.getValue('y1', 'B1'), 4)
  assert.deepEqual(['A1', 'A2', 'A3'].map((address) => visibility.getValue('y2', address)), [4, 4, 40])
  visibilityBook = produce(visibilityBook, (draft) => { draft.sheets[0].hiddenRows = []; draft.sheets[0].filteredRows = [3] })
  visibility.update(visibilityBook)
  assert.deepEqual(['A1', 'A2'].map((address) => visibility.getValue('y2', address)), [3, 3])
  visibilityBook = produce(visibilityBook, (draft) => { draft.sheets[0].cells.D1 = f('1+1') })
  visibility.update(visibilityBook)
  assert.equal(visibility.getValue('y1', 'C1'), true)
  assert.equal(new CalculationEngine(visibilityBook).getValue('y2', 'A3'), visibility.getValue('y2', 'A3'))

  // Legacy formulas from workbook files intersect instead of spilling, and keep their kind on
  // save; formulas typed here keep dynamic-array meaning (calc-formula-engine-5).
  const legacy = (formula: string): CellData => ({ formula, implicitIntersection: true } as CellData)
  const legacyBook: WorkbookModel = {
    ...workbook(sheet('z1', 'Sheet1', {
      A1: v(10), A2: v(20), A3: v(30),
      B2: legacy('A1:A3*2'), B3: legacy('A1:A3*2'), C2: legacy('Price'), D2: legacy('SUMPRODUCT(A1:A3*2)'), D3: legacy('SUM(A1:A3*2)'),
      E2: f('A1:A3*2'), F1: f('SUM(A1:A3*2)'), G1: f('SUM(A1:A3)'), H1: { formula: 'A1:A3*2', formulaType: 'array', formulaRange: 'H1:H3' } as CellData,
    })),
    definedNames: [{ name: 'Price', ranges: ['Sheet1!$A$1:$A$3'] }],
  }
  const legacyEngine = new CalculationEngine(legacyBook)
  assert.deepEqual(['B2', 'B3', 'B4', 'C2', 'C3', 'D2', 'D3'].map((address) => legacyEngine.getValue('z1', address)), [40, 60, null, 20, null, 120, 60])
  assert.deepEqual(['E2', 'E3', 'E4', 'F1', 'H1', 'H3'].map((address) => legacyEngine.getValue('z1', address)), [20, 40, 60, 120, 20, 60])
  const savedCells = legacyEngine.withResults().sheets[0].cells
  assert.equal(savedCells.B3.formulaType, undefined, 'a legacy formula is not turned into an array formula')
  assert.equal(savedCells.B3.formulaRange, undefined)
  assert.equal(savedCells.B3.result, 60)
  assert.equal(savedCells.C2.dynamicFormula, undefined)
  assert.equal(savedCells.E2.formulaRange, 'E2:E4')
  assert.equal(savedCells.F1.formulaType, 'array', 'a typed SUM(A1:A3*2) keeps its array meaning in Excel')
  assert.equal(savedCells.F1.formulaRange, 'F1')
  assert.equal(savedCells.G1.formulaType, undefined)
}

// ---- CALC-027: calculation options, F9, circular references, iteration --------------------------
{
  // Circular references are reported with their cells.
  let circularBook = workbook(sheet('c1', 'Sheet1', { A1: f('B1+1'), B1: f('A1+1'), C1: f('A1*2'), D1: v(1), E1: f('E1+D1') }))
  const circular = new CalculationEngine(circularBook)
  assert.equal(circular.getValue('c1', 'C1'), '#CIRC!')
  assert.deepEqual(circular.circularReferences(), [{ sheetId: 'c1', address: 'A1' }, { sheetId: 'c1', address: 'B1' }])
  assert.equal(circular.getValue('c1', 'E1'), '#CIRC!')
  assert.equal(circular.circularReferences().length, 3)
  circularBook = produce(circularBook, (draft) => { draft.sheets[0].cells.B1 = v(4) })
  circular.update(circularBook)
  assert.equal(circular.getValue('c1', 'C1'), 10)
  assert.deepEqual(circular.circularReferences().map((cell) => cell.address), ['E1'])

  // Iterative calculation (Excel defaults: 100 iterations, 0.001 maximum change).
  circularBook = withCalculationOptions(circularBook, { iterate: true })
  assert.deepEqual(calculationOptionsOf(circularBook), { mode: 'automatic', iterate: true, maxIterations: 100, maxChange: 0.001 })
  circular.update(circularBook)
  assert.equal(circular.getValue('c1', 'E1'), 100, 'E1=E1+1 runs 100 iterations')
  let iterationBook = withCalculationOptions(workbook(sheet('i1', 'Sheet1', {
    A1: f('0.5*A1+1'), B1: f('A1*10'), C1: v(1000), D1: f('C1*0.1+E1'), E1: f('D1*0.05'),
  })), { iterate: true, maxIterations: 100, maxChange: 0.001 })
  const iteration = new CalculationEngine(iterationBook)
  const converged = iteration.getValue('i1', 'A1') as number
  assert.ok(Math.abs(converged - 2) <= 0.001, `0.5*A1+1 converges to 2 (got ${converged})`)
  assert.ok(Math.abs((iteration.getValue('i1', 'B1') as number) - converged * 10) < 1e-9)
  assert.ok(Math.abs((iteration.getValue('i1', 'D1') as number) - 100 / 0.95) < 0.01, 'interest-style circularity converges')
  iterationBook = produce(iterationBook, (draft) => { draft.sheets[0].cells.C1 = v(2000) })
  iteration.update(iterationBook)
  assert.ok(Math.abs((iteration.getValue('i1', 'D1') as number) - 200 / 0.95) < 0.01, 'iteration continues after an edit')
  iterationBook = withCalculationOptions(iterationBook, { iterate: false })
  iteration.update(iterationBook)
  assert.equal(iteration.getValue('i1', 'A1'), '#CIRC!')

  // Manual calculation: edits wait for recalculate() (F9); entered formulas calculate at once.
  let manualBook = withCalculationOptions(workbook(sheet('m1', 'Sheet1', { A1: v(1), B1: f('A1*10'), C1: f('B1+1') })), { mode: 'manual' })
  const manual = new CalculationEngine(manualBook)
  assert.equal(manual.calculationOptions.mode, 'manual')
  assert.equal(manual.getValue('m1', 'C1'), 11)
  assert.equal(manual.needsRecalculation, false)
  manualBook = produce(manualBook, (draft) => { draft.sheets[0].cells.A1 = v(5) })
  manual.update(manualBook)
  assert.equal(manual.getValue('m1', 'B1'), 10, 'dependents keep their value in manual mode')
  assert.equal(manual.needsRecalculation, true)
  manualBook = produce(manualBook, (draft) => { draft.sheets[0].cells.D1 = f('A1+100') })
  manual.update(manualBook)
  assert.equal(manual.getValue('m1', 'D1'), 105, 'an entered formula calculates')
  manual.recalculate()
  assert.deepEqual([manual.getValue('m1', 'B1'), manual.getValue('m1', 'C1'), manual.needsRecalculation], [50, 51, false])
  manualBook = produce(manualBook, (draft) => { draft.sheets[0].cells.A1 = v(7) })
  manual.update(manualBook)
  manual.recalculateSheet('m1')
  assert.equal(manual.getValue('m1', 'C1'), 71, 'Shift+F9 calculates the sheet')
  manualBook = produce(manualBook, (draft) => { draft.sheets[0].cells.A1 = v(2) })
  manual.update(manualBook)
  manualBook = withCalculationOptions(manualBook, { mode: 'automatic' })
  manual.update(manualBook)
  assert.deepEqual([manual.getValue('m1', 'B1'), manual.getValue('m1', 'C1'), manual.needsRecalculation], [20, 21, false], 'switching to automatic applies held-back changes')
  const before = manual.version
  manual.recalculateAll()
  assert.ok(manual.version > before)
  assert.equal(manual.getValue('m1', 'C1'), 21)

  // Review F4: saving / exporting a manual-mode workbook recalculates first (Excel's
  // "Recalculate workbook before saving"), unless the file turned that option off.
  let saveBook = withCalculationOptions(workbook(sheet('s1', 'Sheet1', { A1: v(5), B1: f('A1*2') })), { mode: 'manual' })
  const saveEngine = new CalculationEngine(saveBook)
  assert.equal(saveEngine.getValue('s1', 'B1'), 10)
  saveBook = produce(saveBook, (draft) => { draft.sheets[0].cells.A1 = v(50) })
  saveEngine.update(saveBook)
  assert.equal(saveEngine.needsRecalculation, true)
  assert.equal(saveEngine.withResults({ forSave: false }).sheets[0].cells.B1.result, 10, 'printing shows the values on screen')
  assert.equal(saveEngine.needsRecalculation, true)
  assert.equal(saveEngine.withResults().sheets[0].cells.B1.result, 100, 'the saved result is recalculated')
  assert.equal(saveEngine.needsRecalculation, false)
  let noCalcOnSave = withCalculationOptions(workbook(sheet('s2', 'Sheet1', { A1: v(5), B1: f('A1*2') })), { mode: 'manual' })
  noCalcOnSave = { ...noCalcOnSave, metadata: { ...noCalcOnSave.metadata, calcProperties: { ...noCalcOnSave.metadata!.calcProperties, calcOnSave: false } } }
  const noCalcEngine = new CalculationEngine(noCalcOnSave)
  noCalcEngine.getValue('s2', 'B1')
  noCalcOnSave = produce(noCalcOnSave, (draft) => { draft.sheets[0].cells.A1 = v(50) })
  noCalcEngine.update(noCalcOnSave)
  assert.equal(noCalcEngine.withResults().sheets[0].cells.B1.result, 10, 'calcOnSave="0" keeps the last calculated value, as in Excel')

  // CALC-007: engine diagnostics resolve workbook names.
  const named: WorkbookModel = { ...workbook(sheet('n1', 'Sheet1', { A1: f('SUMM(1)'), A2: f('Rate*2'), A3: f('Missing*2'), A4: f('SUM(1') })), definedNames: [{ name: 'Rate', ranges: ['0.5'] }] }
  const diagnostics = new CalculationEngine(named)
  assert.equal(diagnostics.getValue('n1', 'A4'), '#NAME?')
  assert.equal(diagnostics.diagnose('n1', 'A1')?.kind, 'unknown-function')
  assert.equal(diagnostics.diagnose('n1', 'A2'), null)
  assert.equal(diagnostics.diagnose('n1', 'A3')?.kind, 'unknown-name')
  assert.equal(diagnostics.diagnose('n1', 'A4')?.suggestion, '=SUM(1)')
  assert.deepEqual(diagnostics.diagnose('n1', 'B1', '=IF(A1,1'), diagnoseFormula('=IF(A1,1'))
}

console.log('Calculation engine QA passed: incremental recalculation, spills, saved arrays, tables, visibility, scaling, validation candidates, legacy formulas, calculation options, and circular references.')
