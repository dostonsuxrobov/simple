import assert from 'node:assert/strict'
import { computePivot, pivotFieldIsDate, pivotFieldKeys, pivotOverwriteConflicts, writePivotOutput } from '../src/lib/pivot'
import type { PivotSource } from '../src/lib/pivot'
import type { PivotTableModel, SheetData } from '../src/spreadsheet-types'

const serial = (y: number, m: number, d: number) => Math.round(Date.UTC(y, m - 1, d) / 86_400_000) + 25569
const cell = (value: string | number | null, extra: Record<string, unknown> = {}) => ({ value, text: value === null ? '' : String(value), ...extra })
const source: PivotSource = {
  headers: ['Region', 'Product', 'Date', 'Units', 'Price'],
  records: [
    ['East', 'Pen', serial(2024, 1, 15), 3, 2],
    ['East', 'Ink', serial(2024, 2, 10), 5, 4],
    ['West', 'Pen', serial(2024, 1, 20), 7, 2],
    ['West', 'Pen', serial(2023, 12, 5), 1, 2],
    ['West', 'Ink', null, 2, 4],
  ].map((row) => row.map((value, index) => cell(value as string | number | null, index === 2 && value !== null ? { isDate: true } : index === 4 ? { numFmt: '$#,##0.00' } : {}))),
}
const model = (overrides: Partial<PivotTableModel>): PivotTableModel => ({ id: 'p', name: 'PivotTable1', source: 'Data!A1:E6', anchor: { row: 0, col: 0 }, rows: [], columns: [], values: [], filters: [], ...overrides })
const grid = (output: ReturnType<typeof computePivot>) => output.cells.map((row) => row.map((item) => item.value))

assert.deepEqual(grid(computePivot(source, model({ rows: [{ field: 'Region' }], values: [{ field: 'Units', summarize: 'sum' }] }))), [
  ['Region', 'Sum of Units'],
  ['East', 8],
  ['West', 10],
  ['Grand Total', 18],
])

assert.deepEqual(grid(computePivot(source, model({ rows: [{ field: 'Region' }, { field: 'Product' }], values: [{ field: 'Units', summarize: 'sum' }] }))), [
  ['Region', 'Product', 'Sum of Units'],
  ['East', 'Ink', 5],
  [null, 'Pen', 3],
  ['East Total', null, 8],
  ['West', 'Ink', 2],
  [null, 'Pen', 8],
  ['West Total', null, 10],
  ['Grand Total', null, 18],
], 'nested rows with subtotals and repeated labels suppressed')

assert.deepEqual(grid(computePivot(source, model({ rows: [{ field: 'Product' }], columns: [{ field: 'Region' }], values: [{ field: 'Units', summarize: 'sum' }] }))), [
  ['Sum of Units', 'Region', null, null],
  ['Product', 'East', 'West', 'Grand Total'],
  ['Ink', 5, 2, 7],
  ['Pen', 3, 8, 11],
  ['Grand Total', 8, 10, 18],
])

{
  const output = computePivot(source, model({ rows: [{ field: 'Product' }], columns: [{ field: 'Region' }], values: [{ field: 'Units', summarize: 'sum' }, { field: 'Price', summarize: 'average' }] }))
  const rows = grid(output)
  assert.deepEqual(rows[1], [null, 'East', null, 'West', null, 'Total', null])
  assert.deepEqual(rows[2], ['Product', 'Sum of Units', 'Average of Price', 'Sum of Units', 'Average of Price', 'Total Sum of Units', 'Total Average of Price'])
  assert.deepEqual(rows[3], ['Ink', 5, 4, 2, 4, 7, 4])
  assert.equal(output.cells[3][2].numFmt, '$#,##0.00', 'value fields keep the source number format')
  assert.equal(output.cells[3][1].numFmt, undefined)
}

assert.equal(pivotFieldIsDate(source, 'Date'), true)
assert.deepEqual(grid(computePivot(source, model({ rows: [{ field: 'Date', dateGroup: 'year' }], values: [{ field: 'Units', summarize: 'sum' }] }))).slice(1), [
  ['2023', 1],
  ['2024', 15],
  ['(blank)', 2],
  ['Grand Total', 18],
], 'years sort chronologically with blanks last')
assert.deepEqual(grid(computePivot(source, model({ rows: [{ field: 'Date', dateGroup: 'month' }], values: [{ field: 'Units', summarize: 'count' }] }))).slice(1, 4), [
  ['Jan', 2],
  ['Feb', 1],
  ['Dec', 1],
])

{
  const output = computePivot(source, model({ rows: [{ field: 'Region' }], values: [{ field: 'Units', summarize: 'sum', showAs: 'percentOfGrandTotal' }] }))
  assert.ok(Math.abs((output.cells[1][1].value as number) - 8 / 18) < 1e-12)
  assert.equal(output.cells[1][1].numFmt, '0.00%')
}
{
  const keys = pivotFieldKeys(source, 'Product')
  assert.deepEqual(keys.map((key) => [key.label, key.count]), [['Ink', 2], ['Pen', 3]])
  const output = computePivot(source, model({ rows: [{ field: 'Region' }], values: [{ field: 'Units', summarize: 'sum' }], filters: [{ field: 'Product', exclude: [keys[0].id] }] }))
  assert.deepEqual(grid(output).slice(1), [['East', 3], ['West', 8], ['Grand Total', 11]], 'filtered-out values are excluded')
}
assert.deepEqual(grid(computePivot(source, model({ rows: [{ field: 'Region', order: 'desc' }], values: [{ field: 'Product', summarize: 'countDistinct' }, { field: 'Units', summarize: 'median' }] }))).slice(1), [
  ['West', 2, 2],
  ['East', 2, 4],
  ['Grand Total', 2, 3],
])
assert.deepEqual(grid(computePivot(source, model({ values: [{ field: 'Units', summarize: 'sum' }, { field: 'Units', summarize: 'max' }] }))), [
  [null, 'Sum of Units', 'Max of Units'],
  ['Total', 18, 7],
], 'values alone give a single total row')
assert.match(String(computePivot(source, model({ rows: [{ field: 'Missing' }] })).cells[0][0].value), /no longer/)

// ---- Writing the block --------------------------------------------------------------------------
{
  const sheet: SheetData = { id: 's', name: 'Pivot', rowCount: 10, colCount: 5, cells: { C3: { value: 'keep me' } }, merges: [], colWidths: {}, rowHeights: {} }
  const pivot = model({ anchor: { row: 0, col: 0 }, rows: [{ field: 'Region' }, { field: 'Product' }], values: [{ field: 'Units', summarize: 'sum' }] })
  const big = computePivot(source, pivot)
  assert.deepEqual(pivotOverwriteConflicts(sheet, pivot, big), ['C3'])
  delete sheet.cells.C3
  writePivotOutput(sheet, pivot, big)
  assert.deepEqual(pivot.extent, { rows: 8, cols: 3 })
  assert.equal(sheet.cells.A1.value, 'Region')
  assert.equal(sheet.cells.A1.style?.font?.bold, true)
  assert.equal(sheet.cells.C8.value, 18)
  assert.deepEqual(pivotOverwriteConflicts(sheet, pivot, big), [], 'the pivot may overwrite its own block')
  pivot.rows = [{ field: 'Region' }]
  writePivotOutput(sheet, pivot, computePivot(source, pivot))
  assert.equal(sheet.cells.C8, undefined, 'cells of the previous, larger block are cleared')
  assert.equal(sheet.cells.B4.value, 18)
}

console.log('Pivot QA passed: rows, nested rows, columns, multiple values, date grouping, show-as, filters, sort, distinct/median, writing and conflicts.')
