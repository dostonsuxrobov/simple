import assert from 'node:assert/strict'
import type { SheetData, WorkbookModel } from '../src/spreadsheet-types.ts'
import {
  applySelectionStructureCommand,
  createSheetCopy,
  deleteColumns,
  deleteRows,
  insertBlankSheet,
  insertColumns,
  insertRows,
  insertSheetCopy,
  removeSheetScopedNames,
  rewriteFormulaForSheetStructure,
  rewriteWorkbookFormulaText,
  SheetStructureError,
  shiftCells,
} from '../src/lib/sheet-operations.ts'
import { hiddenRowsWithoutFilter } from '../src/lib/filter.ts'
import { removeSheetFromFormula, renameSheetInFormula } from '../src/lib/formula-editing.ts'

function sheet(overrides: Partial<SheetData> = {}): SheetData {
  return {
    id: 'main',
    name: 'Main',
    state: 'visible',
    rowCount: 10,
    colCount: 8,
    cells: {},
    merges: [],
    colWidths: {},
    rowHeights: {},
    ...overrides,
  }
}

function workbook(main: SheetData, other?: SheetData): WorkbookModel {
  return {
    version: 1,
    name: 'operations.xlsx',
    activeSheetId: main.id,
    sheets: [main, other || sheet({ id: 'other', name: 'Other' })],
    metadata: { calcProperties: { calcMode: 'auto', forceFullCalc: false } },
  }
}

function rowInsertionFixture(): WorkbookModel {
  const main = sheet({
    cells: {
      A1: { value: 'Header' },
      B1: { value: 1 },
      B2: { value: 2 },
      B3: { value: 3 },
      C2: {
        formula: 'SUM(B1:B3)+$C$2+"A2"+Table1[A2]',
        result: 6,
        resultType: 'date',
        display: '6',
      },
      F7: { formula: 'TRANSPOSE(A1:C1)', formulaType: 'array', formulaRange: 'F7:H7' },
    },
    merges: ['A2:C3'],
    rowHeights: { '2': 30, '5': 40 },
    rowProperties: { '2': { outlineLevel: 1 }, '5': { numFmt: '0.00' } },
    hiddenRows: [2, 5],
    frozen: { rows: 2, columns: 2, topLeftCell: 'C3', activeCell: 'D4' },
    views: [{ state: 'frozen', ySplit: 2, xSplit: 2, topLeftCell: 'C3', activeCell: 'D4' }],
    rowBreaks: [{ id: 5, min: 0, max: 16_383 }],
    dataValidations: {
      'D2:D4': { type: 'custom', formulae: ['B2:B4', '"A1,A2"'] },
    },
    conditionalFormattings: [{ ref: 'E2:E4', rules: [
      { type: 'expression', formulae: ['B2>0'] },
      { type: 'colorScale', cfvo: [{ type: 'formula', value: '$B$3' }, { type: 'num', value: 5 }, { type: 'percent', value: '50' }] },
    ] }],
    autoFilter: 'A1:D4',
    pageSetup: { printArea: "'Main'!A1:H10", printTitlesRow: '1:1' },
  })
  const other = sheet({
    id: 'other',
    name: 'Other',
    cells: {
      A1: { formula: 'Main!B2+$C$1+\'Main\'!A1+"Main!A2"+[Other.xlsx]Main!B2', result: 3, display: '3' },
    },
    conditionalFormattings: [{ ref: 'A1:A2', rules: [{ formulae: ['Main!B2>0'] }] }],
  })
  const model = workbook(main, other)
  model.definedNames = [
    { name: 'GlobalInput', ranges: ['Main!$B$2:$B$3'], ref: 'Main!$B$2:$B$3' },
    { name: 'LocalRows', ranges: ['$A$1:$A$3'], ref: '$A$1:$A$3', localSheetIndex: 0 },
  ]
  model.metadata!.definedNames = [{ name: 'MetadataInput', ranges: 'Main!$B$2:$B$3' }]
  return model
}

{
  const source = rowInsertionFixture()
  const snapshot = structuredClone(source)
  const result = insertRows(source, 'main', 1, 2)
  const main = result.sheets[0]
  const other = result.sheets[1]

  assert.deepEqual(source, snapshot, 'row insertion must not mutate its input')
  assert.equal(main.rowCount, 12)
  assert.equal(main.cells.C4.formula, 'SUM(B1:B5)+$C$4+"A2"+Table1[A2]')
  assert.equal(main.cells.C4.result, undefined)
  assert.equal(main.cells.C4.resultType, undefined)
  assert.equal(main.cells.C4.display, undefined)
  assert.equal(main.cells.F9.formulaRange, 'F9:H9')
  assert.deepEqual(main.merges, ['A4:C5'])
  assert.deepEqual(main.rowHeights, { '4': 30, '7': 40 })
  assert.deepEqual(main.rowProperties, { '4': { outlineLevel: 1 }, '7': { numFmt: '0.00' } })
  assert.deepEqual(main.hiddenRows, [4, 7])
  assert.deepEqual(main.frozen, { rows: 4, columns: 2, topLeftCell: 'C5', activeCell: 'D6' })
  assert.equal(main.views?.[0].ySplit, 4)
  assert.equal(main.views?.[0].topLeftCell, 'C5')
  assert.deepEqual(main.rowBreaks, [{ id: 7, min: 0, max: 16_383 }])
  assert.deepEqual(Object.keys(main.dataValidations || {}), ['D4:D6'])
  assert.deepEqual((main.dataValidations?.['D4:D6'] as { formulae: string[] }).formulae, ['B4:B6', '"A1,A2"'])
  assert.equal((main.conditionalFormattings?.[0] as { ref: string }).ref, 'E4:E6')
  assert.deepEqual(
    ((main.conditionalFormattings?.[0] as { rules: Array<{ formulae: string[] }> }).rules[0]).formulae,
    ['B4>0'],
  )
  assert.deepEqual(
    ((main.conditionalFormattings?.[0] as { rules: Array<{ cfvo?: Array<{ value: unknown }> }> }).rules[1]).cfvo?.map((item) => item.value),
    ['$B$5', 5, '50'],
    'formula thresholds follow inserted rows; numbers stay',
  )
  assert.equal(main.autoFilter, 'A1:D6')
  assert.equal(main.pageSetup?.printArea, "'Main'!A1:H12")
  assert.equal(main.pageSetup?.printTitlesRow, '1:1')
  assert.equal(other.cells.A1.formula, 'Main!B4+$C$1+\'Main\'!A1+"Main!A2"+[Other.xlsx]Main!B2')
  assert.equal(other.cells.A1.result, undefined)
  assert.deepEqual(
    ((other.conditionalFormattings?.[0] as { rules: Array<{ formulae: string[] }> }).rules[0]).formulae,
    ['Main!B4>0'],
  )
  assert.equal(result.definedNames?.[0].ref, 'Main!$B$4:$B$5')
  assert.equal(result.definedNames?.[1].ref, '$A$1:$A$5')
  assert.equal(result.metadata?.definedNames?.[0].ranges, 'Main!$B$4:$B$5')
  assert.equal(result.metadata?.calcProperties?.fullCalcOnLoad, true)
  assert.equal(result.metadata?.calcProperties?.forceFullCalc, true)
}

{
  const target = sheet()
  const operation = { axis: 'row', kind: 'insert', index: 1, count: 2 } as const
  assert.equal(
    rewriteFormulaForSheetStructure('SUM(1:3)+SUM(A:C)+Main!$A$2', 'main', target, operation),
    'SUM(1:5)+SUM(A:C)+Main!$A$4',
  )
}

{
  const source = workbook(sheet({ merges: ['A2:C4'] }))
  const result = insertRows(source, 'main', 2, 1)
  assert.deepEqual(result.sheets[0].merges, ['A2:C5'], 'inserting inside a merge expands it')
}

{
  const main = sheet({
    cells: {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
      A4: { value: 4 },
      A5: { value: 5 },
      C6: { formula: 'SUM(A1:A5)+A2+A5', result: 10, display: '10' },
    },
    merges: ['A1:B5'],
    rowHeights: { '2': 20, '4': 40, '6': 60 },
    hiddenRows: [2, 4, 6],
    frozen: { rows: 3, columns: 1, topLeftCell: 'C4', activeCell: 'C4' },
    dataValidations: {
      'D1:D5': { formulae: ['A1:A5'] },
      'E2:E3': { type: 'whole' },
    },
    autoFilter: { ref: 'A1:D5' },
  })
  const other = sheet({
    id: 'other', name: 'Other', cells: { A1: { formula: 'Main!A2+Main!A5', result: 7 } },
  })
  const result = deleteRows(workbook(main, other), 'main', 1, 2)
  const output = result.sheets[0]

  assert.equal(output.rowCount, 8)
  assert.deepEqual(Object.keys(output.cells).sort(), ['A1', 'A2', 'A3', 'C4'])
  assert.equal(output.cells.A2.value, 4)
  assert.equal(output.cells.A3.value, 5)
  assert.equal(output.cells.C4.formula, 'SUM(A1:A3)+#REF!+A3')
  assert.deepEqual(output.merges, ['A1:B3'])
  assert.deepEqual(output.rowHeights, { '2': 40, '4': 60 })
  assert.deepEqual(output.hiddenRows, [2, 4])
  assert.deepEqual(output.frozen, { rows: 1, columns: 1, topLeftCell: 'C2', activeCell: 'C2' })
  assert.deepEqual(Object.keys(output.dataValidations || {}), ['D1:D3'])
  assert.deepEqual((output.dataValidations?.['D1:D3'] as { formulae: string[] }).formulae, ['A1:A3'])
  assert.deepEqual(output.autoFilter, { ref: 'A1:D3' })
  assert.equal(result.sheets[1].cells.A1.formula, 'Main!#REF!+Main!A3')
}

{
  const main = sheet({
    cells: {
      A1: { value: 1 }, B1: { value: 2 }, C1: { value: 3 }, D1: { value: 4 },
      E1: { formula: 'SUM(A1:D1)+B1+D1', result: 10 },
    },
    merges: ['A1:D2'],
    colWidths: { '2': 20, '4': 40, '5': 50 },
    columnProperties: { '2': { outlineLevel: 1 }, '4': { numFmt: '0' } },
    hiddenCols: [2, 4, 5],
    frozen: { rows: 1, columns: 3, topLeftCell: 'D2' },
  })
  const deleted = deleteColumns(workbook(main), 'main', 1, 2).sheets[0]
  assert.deepEqual(Object.keys(deleted.cells).sort(), ['A1', 'B1', 'C1'])
  assert.equal(deleted.cells.B1.value, 4)
  assert.equal(deleted.cells.C1.formula, 'SUM(A1:B1)+#REF!+B1')
  assert.deepEqual(deleted.merges, ['A1:B2'])
  assert.deepEqual(deleted.colWidths, { '2': 40, '3': 50 })
  assert.deepEqual(deleted.columnProperties, { '2': { numFmt: '0' } })
  assert.deepEqual(deleted.hiddenCols, [2, 3])
  assert.deepEqual(deleted.frozen, { rows: 1, columns: 1, topLeftCell: 'B2' })

  const inserted = insertColumns(workbook(sheet({
    cells: { A1: { value: 1 }, B1: { formula: 'SUM(A1:B1)+$B$1' } },
    merges: ['B2:C3'],
    colWidths: { '2': 20 },
    hiddenCols: [2],
    frozen: { rows: 1, columns: 2, topLeftCell: 'C3' },
  })), 'main', 1, 1).sheets[0]
  assert.equal(inserted.cells.C1.formula, 'SUM(A1:C1)+$C$1')
  assert.deepEqual(inserted.merges, ['C2:D3'])
  assert.deepEqual(inserted.colWidths, { '3': 20 })
  assert.deepEqual(inserted.hiddenCols, [3])
  assert.deepEqual(inserted.frozen, { rows: 1, columns: 3, topLeftCell: 'D3' })
}

{
  const source = workbook(sheet({
    cells: {
      A2: { formula: 'TRANSPOSE(D1:F2)', formulaType: 'array', formulaRange: 'A2:C3' },
    },
  }))
  assert.throws(
    () => insertRows(source, 'main', 2, 1),
    (error: unknown) => error instanceof SheetStructureError && error.code === 'ARRAY_RANGE_CONFLICT',
  )
  assert.throws(
    () => deleteRows(source, 'main', 1, 1),
    (error: unknown) => error instanceof SheetStructureError && error.code === 'ARRAY_RANGE_CONFLICT',
  )
  const removed = deleteRows(source, 'main', 1, 2)
  assert.equal(Object.keys(removed.sheets[0].cells).length, 0)
}

{
  const source = workbook(sheet({ cells: { XFD1: { value: 'edge' } }, colCount: 16_384 }))
  assert.throws(
    () => insertColumns(source, 'main', 0, 1),
    (error: unknown) => error instanceof SheetStructureError && error.code === 'LIMIT_EXCEEDED',
  )
}

{
  const source = workbook(sheet({ name: 'Sheet1' }), sheet({ id: 'other', name: 'Sheet3' }))
  source.definedNames = [{ name: 'LocalOnThird', ranges: ['$A$1'], localSheetIndex: 1 }]
  source.metadata!.definedNames = [{ name: 'MetadataLocalOnThird', ranges: '$A$1', localSheetId: 1 }]
  const snapshot = structuredClone(source)
  const added = insertBlankSheet(source, { afterSheetId: 'main' })
  assert.deepEqual(source, snapshot)
  assert.equal(added.workbook.sheets[1].name, 'Sheet2')
  assert.equal(added.workbook.sheets[1].id, added.sheetId)
  assert.equal(added.workbook.activeSheetId, added.sheetId)
  assert.equal(added.workbook.sheets[1].rowCount, 200)
  assert.equal(added.workbook.sheets[1].colCount, 40)
  assert.equal(added.workbook.definedNames?.[0].localSheetIndex, 2)
  assert.equal(added.workbook.metadata?.definedNames?.[0].localSheetId, 2)
}

{
  const source = workbook(sheet())
  const result = applySelectionStructureCommand(
    source,
    'main',
    { anchor: { row: 2, col: 1 }, focus: { row: 4, col: 3 } },
    'insert-rows-below',
  )
  assert.deepEqual(result.operation, { axis: 'row', kind: 'insert', index: 5, count: 3 })
  assert.deepEqual(result.selection, {
    anchor: { row: 5, col: 1 },
    focus: { row: 7, col: 3 },
  })
  assert.equal(result.workbook.sheets[0].rowCount, 13)
}

// ---- Insert/delete cells (shift) -------------------------------------------------------------
{
  const base = workbook(sheet({
    cells: {
      A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: { value: 10 }, B3: { value: 30 }, C1: { value: 100 },
      D1: { formula: 'SUM(A1:A5)' }, E1: { formula: 'B3' }, F1: { formula: 'SUM(A1:C1)' }, G1: { formula: 'A3*2' }, H1: { formula: 'C1+B1' },
    },
    rowHeights: { '2': 40 },
    dataValidations: { 'A2:A3': { type: 'whole', formulae: ['0'] } },
  }))
  const down = shiftCells(base, 'main', { top: 1, bottom: 1, left: 0, right: 0 }, 'down').sheets[0]
  assert.equal(down.cells.A2, undefined, 'the inserted cell is blank')
  assert.equal(down.cells.A3.value, 2)
  assert.equal(down.cells.A4.value, 3)
  assert.equal(down.cells.B3.value, 30, 'cells outside the band stay')
  assert.equal(down.cells.D1.formula, 'SUM(A1:A6)', 'ranges inside the band grow')
  assert.equal(down.cells.E1.formula, 'B3')
  assert.equal(down.cells.F1.formula, 'SUM(A1:C1)', 'ranges wider than the band are untouched')
  assert.equal(down.cells.G1.formula, 'A4*2', 'references follow shifted cells')
  assert.deepEqual(down.rowHeights, { '2': 40 }, 'row heights do not move with cells')
  assert.deepEqual(Object.keys(down.dataValidations || {}), ['A3:A4'])

  const left = shiftCells(base, 'main', { top: 0, bottom: 0, left: 1, right: 1 }, 'left').sheets[0]
  assert.equal(left.cells.B1.value, 100, 'C1 moves into B1')
  assert.equal(left.cells.B3.value, 30, 'rows outside the band stay')
  assert.equal(left.cells.C1.formula, 'SUM(A1:A5)')
  assert.equal(left.cells.D1.formula, 'B3')
  assert.equal(left.cells.E1.formula, 'SUM(A1:B1)', 'a range across the deleted cell shrinks')
  assert.equal(left.cells.F1.formula, 'A3*2')
  assert.equal(left.cells.G1.formula, 'B1+#REF!', 'references to the deleted cell become #REF!')
  assert.equal(left.cells.H1, undefined)
}
{
  const merged = workbook(sheet({ cells: { A1: { value: 1 } }, merges: ['A2:B2'] }))
  assert.throws(() => shiftCells(merged, 'main', { top: 0, bottom: 0, left: 0, right: 0 }, 'down'), (error: unknown) => error instanceof SheetStructureError && error.code === 'SHIFT_CONFLICT')
  const inside = shiftCells(merged, 'main', { top: 0, bottom: 0, left: 0, right: 1 }, 'down').sheets[0]
  assert.deepEqual(inside.merges, ['A3:B3'], 'merges inside the band move')
}

// calc-grid-interaction-5: inserting or deleting rows/columns moves the AutoFilter with the data.
// Header in row 3, data in rows 4-7, filter column A = 'East' (rows 5 and 7 hidden).
function filteredFixture(extra: Partial<SheetData> = {}): WorkbookModel {
  return workbook(sheet({
    rowCount: 20,
    cells: {
      A3: { value: 'Region' }, B3: { value: 'Sales' },
      A4: { value: 'East' }, B4: { value: 1 },
      A5: { value: 'West' }, B5: { value: 2 },
      A6: { value: 'East' }, B6: { value: 3 },
      A7: { value: 'West' }, B7: { value: 4 },
    },
    filter: { ref: 'A3:B7', columns: { 0: { values: ['East'] } } },
    autoFilter: 'A3:B7',
    filteredRows: [5, 7],
    hiddenRows: [5, 7],
    ...extra,
  }))
}
{
  const above = insertRows(filteredFixture(), 'main', 1, 1).sheets[0]
  assert.equal(above.filter?.ref, 'A4:B8', 'the filter range follows its header down')
  assert.equal(above.autoFilter, 'A4:B8')
  assert.deepEqual(above.filteredRows, [6, 8], 'filtered rows move with their data')
  assert.deepEqual(above.hiddenRows, [6, 8])
  assert.deepEqual(hiddenRowsWithoutFilter(above), [], 'removing the filter afterwards shows every row')
  assert.ok(!above.filteredRows?.includes(4), 'the header row is never hidden')

  const inside = insertRows(filteredFixture(), 'main', 5, 2).sheets[0]
  assert.equal(inside.filter?.ref, 'A3:B9', 'rows inserted inside the list extend it')
  assert.deepEqual(inside.filteredRows, [5, 9], 'new rows are visible; later filtered rows shift')

  const removedData = deleteRows(filteredFixture(), 'main', 4, 1).sheets[0]
  assert.equal(removedData.filter?.ref, 'A3:B6')
  assert.deepEqual(removedData.filteredRows, [6], 'a deleted filtered row is gone; the next one shifts up')

  const noHeader = deleteRows(filteredFixture({ hiddenRows: [2, 5, 7] }), 'main', 2, 1).sheets[0]
  assert.equal(noHeader.filter, undefined, 'deleting the header row removes the filter')
  assert.equal(noHeader.autoFilter, undefined)
  assert.equal(noHeader.filteredRows, undefined)
  assert.deepEqual(noHeader.hiddenRows, [2], 'rows the filter hid come back; a row hidden by hand stays hidden')

  const leftColumn = insertColumns(filteredFixture(), 'main', 0, 1).sheets[0]
  assert.equal(leftColumn.filter?.ref, 'B3:C7')
  assert.deepEqual(Object.keys(leftColumn.filter?.columns || {}), ['0'], 'criteria stay on their column')
  assert.deepEqual(leftColumn.filteredRows, [5, 7], 'column edits do not move rows')

  const between = filteredFixture({ filter: { ref: 'A3:B7', columns: { 1: { condition: { operator: 'greaterThan', value: 2 } } } } })
  const widened = insertColumns(between, 'main', 1, 1).sheets[0]
  assert.equal(widened.filter?.ref, 'A3:C7')
  assert.deepEqual(Object.keys(widened.filter?.columns || {}), ['2'], 'a criterion right of an inserted column moves with its column')

  const withoutColumns = deleteColumns(filteredFixture(), 'main', 0, 2).sheets[0]
  assert.equal(withoutColumns.filter, undefined, 'deleting every filter column removes the filter')
  assert.deepEqual(withoutColumns.hiddenRows, [], 'and its rows come back')

  assert.throws(
    () => shiftCells(filteredFixture(), 'main', { top: 4, bottom: 4, left: 0, right: 0 }, 'down'),
    (error: unknown) => error instanceof SheetStructureError && error.code === 'SHIFT_CONFLICT',
    'cells inside a filtered list cannot be shifted',
  )
  const below = shiftCells(filteredFixture(), 'main', { top: 10, bottom: 10, left: 0, right: 1 }, 'down').sheets[0]
  assert.equal(below.filter?.ref, 'A3:B7', 'a shift below the list leaves the filter alone')
  const unfiltered = filteredFixture({ filter: { ref: 'A3:B7', columns: {} }, filteredRows: undefined, hiddenRows: undefined })
  const grown = shiftCells(unfiltered, 'main', { top: 4, bottom: 4, left: 0, right: 1 }, 'down').sheets[0]
  assert.equal(grown.filter?.ref, 'A3:B8', 'with no criteria, a full-width shift inside the list grows it')
}
{
  // A table's own filter follows its columns.
  const tableSheet = sheet({
    cells: { A1: { value: 'Name' }, B1: { value: 'Qty' }, A2: { value: 'a' }, B2: { value: 1 } },
    tables: [{ id: 't1', name: 'Sales', ref: 'A1:B2', headerRow: true, totalsRow: false, columns: [{ name: 'Name' }, { name: 'Qty' }], filter: { ref: 'A1:B2', columns: { 1: { values: ['1'] } } } }],
  })
  const inserted = insertColumns(workbook(tableSheet), 'main', 1, 1).sheets[0]
  assert.equal(inserted.tables?.[0].ref, 'A1:C2')
  assert.deepEqual(Object.keys(inserted.tables?.[0].filter?.columns || {}), ['2'], 'the Qty criterion moves with the Qty column')
  assert.equal(inserted.tables?.[0].filter?.ref, 'A1:C2')
}

// calc-file-io-objects-10: a duplicated sheet's tables get unique names; its own references and
// charts point at the copy; import bookkeeping is dropped.
{
  const data = sheet({
    id: 'data',
    name: 'Data',
    sourceWorksheetId: 1,
    sourceSheetName: 'Data',
    sourceSheetIndex: 0,
    cells: {
      A1: { value: 'Name' }, B1: { value: 'Qty' }, A2: { value: 'a' }, B2: { value: 5 }, A3: { value: 'b' }, B3: { value: 7 },
      C1: { formula: 'SUM(Sales[Qty])', result: 12 },
      D1: { formula: 'Data!A1&Other!A1' },
    },
    tables: [{ id: 'table-Sales', name: 'Sales', displayName: 'Sales', ref: 'A1:B3', headerRow: true, totalsRow: false, columns: [{ name: 'Name' }, { name: 'Qty' }], imported: true }],
    dataValidations: { 'E1:E3': { type: 'list', formulae: ['Sales[Name]'] } },
    sparklineGroups: [{ type: 'line', colors: {}, sparklines: [{ source: 'Data!B2:B3', cell: 'F1' }], sourceXml: '<x14:sparklineGroup/>', signature: 'sig' }],
    charts: [{ id: 'chart-1', type: 'column', anchor: { from: { row: 5, col: 0 }, to: { row: 15, col: 6 } }, series: [{ id: 's1', valuesRef: 'Data!$B$2:$B$3', categoriesRef: 'Data!$A$2:$A$3' }], sourcePart: 'xl/charts/chart1.xml', sourceInfo: { drawingPart: 'xl/drawings/drawing1.xml', anchorIndex: 0, fingerprint: 'f', kind: 'chart' } }],
  })
  const other = sheet({ id: 'other', name: 'Other', cells: { A1: { formula: 'SUM(Sales[Qty])' } } })
  const model: WorkbookModel = {
    ...workbook(data, other),
    definedNames: [
      { name: 'Local', ranges: ['Data!$A$1'], localSheetIndex: 0 },
      { name: 'OtherLocal', ranges: ['Other!$A$1'], localSheetIndex: 1 },
      { name: 'Global', ranges: ['Data!$B$2'] },
    ],
  }
  let ids = 0
  const copy = createSheetCopy(model, 'data', { id: 'copy', makeId: (prefix) => `${prefix}-${++ids}` })
  assert.equal(copy.name, 'Data copy')
  assert.equal(copy.tables?.[0].name, 'Sales2', 'the copied table gets a workbook-unique name')
  assert.equal(copy.tables?.[0].displayName, 'Sales2')
  assert.notEqual(copy.tables?.[0].id, 'table-Sales')
  assert.equal(copy.tables?.[0].imported, undefined)
  assert.equal(copy.cells.C1.formula, 'SUM(Sales2[Qty])', "the copy's own structured references follow its table")
  assert.equal(copy.cells.D1.formula, "'Data copy'!A1&Other!A1", 'references to the source sheet point at the copy')
  assert.deepEqual((copy.dataValidations?.['E1:E3'] as { formulae: string[] }).formulae, ['Sales2[Name]'])
  assert.equal(copy.sparklineGroups?.[0].sparklines[0].source, "'Data copy'!B2:B3", 'sparklines read the copy')
  assert.equal(copy.charts?.[0].series[0].valuesRef, "'Data copy'!$B$2:$B$3", 'charts plot the copy')
  assert.equal(copy.charts?.[0].sourcePart, undefined, 'the chart is written fresh, never byte-copied from the original')
  assert.equal(copy.charts?.[0].sourceInfo, undefined)
  assert.notEqual(copy.charts?.[0].id, 'chart-1')
  assert.equal(copy.sourceWorksheetId, undefined, 'the copy is a new worksheet when saved')
  assert.equal(copy.sourceSheetName, undefined)
  assert.equal(copy.sourceSheetIndex, undefined)
  assert.equal(model.sheets[0].cells.C1.formula, 'SUM(Sales[Qty])', 'the source sheet is unchanged')
  assert.equal(model.sheets[0].tables?.[0].name, 'Sales')

  const next = structuredClone(model)
  insertSheetCopy(next, copy, 'data')
  assert.deepEqual(next.sheets.map((item) => item.name), ['Data', 'Data copy', 'Other'])
  assert.equal(next.definedNames?.find((name) => name.name === 'OtherLocal')?.localSheetIndex, 2, 'sheet-scoped names stay on their sheet')
  const copiedLocal = next.definedNames?.filter((name) => name.name === 'Local')
  assert.deepEqual(copiedLocal?.map((name) => [name.localSheetIndex, name.ranges[0]]), [[0, 'Data!$A$1'], [1, "'Data copy'!$A$1"]], 'the copy gets its own sheet-scoped names')
  assert.equal(next.definedNames?.filter((name) => name.name === 'Global').length, 1, 'workbook names are not duplicated')
  assert.equal(createSheetCopy(next, 'data', { id: 'copy2' }).name, 'Data copy 2')
  assert.equal(createSheetCopy(next, 'data', { id: 'copy3', makeId: (prefix) => `${prefix}-x` }).tables?.[0].name, 'Sales3', 'Sales2 is taken by the first copy')
}

// calc-file-io-objects-11: renaming or deleting a sheet rewrites the sparklines that read it.
{
  const model: WorkbookModel = workbook(
    sheet({ id: 'main', name: 'Sheet1', cells: { B2: { value: 1 } } }),
    sheet({ id: 'other', name: 'Sales', sparklineGroups: [{ type: 'line', colors: {}, sparklines: [{ source: 'Sheet1!B2:F2', cell: 'A1' }, { source: 'Sales!B2:F2', cell: 'A2' }], sourceXml: '<x/>', signature: 'old' }] }),
  )
  const renamed = structuredClone(model)
  rewriteWorkbookFormulaText(renamed, (formula) => renameSheetInFormula(formula, 'Sheet1', 'Data 2026'))
  assert.deepEqual(renamed.sheets[1].sparklineGroups?.[0].sparklines.map((item) => item.source), ["'Data 2026'!B2:F2", 'Sales!B2:F2'])
  const removed = structuredClone(model)
  rewriteWorkbookFormulaText(removed, (formula) => removeSheetFromFormula(formula, 'Sheet1'))
  assert.equal(removed.sheets[1].sparklineGroups?.[0].sparklines[0].source, '#REF!', 'a deleted source sheet leaves #REF!, never another sheet')
}
{
  const model: WorkbookModel = {
    ...workbook(sheet({ id: 'a', name: 'A' }), sheet({ id: 'b', name: 'B' })),
    definedNames: [
      { name: 'OnA', ranges: ['A!$A$1'], localSheetIndex: 0 },
      { name: 'OnB', ranges: ['B!$A$1'], localSheetIndex: 1 },
      { name: 'OnC', ranges: ['C!$A$1'], localSheetIndex: 2 },
    ],
  }
  removeSheetScopedNames(model, 1)
  assert.deepEqual(model.definedNames?.map((name) => [name.name, name.localSheetIndex]), [['OnA', 0], ['OnC', 1]], "a deleted sheet's names go; later sheets keep theirs")
}

{
  // Review F7: an opened whole-column / whole-row print area keeps its original form when rows
  // or columns are inserted or deleted.
  const columns = workbook(sheet({ cells: { A1: { value: 1 } }, pageSetup: { printArea: 'A1:B5&&D1:D5', printAreaWhole: { 'A1:B5': '$A:$B' } } }))
  const afterInsert = insertRows(columns, 'main', 2, 1).sheets[0].pageSetup!
  assert.equal(afterInsert.printArea, 'A1:B6,D1:D6')
  assert.deepEqual(afterInsert.printAreaWhole, { 'A1:B6': '$A:$B' })
  const afterDelete = deleteRows(columns, 'main', 1, 2).sheets[0].pageSetup!
  assert.deepEqual(afterDelete.printAreaWhole, { 'A1:B3': '$A:$B' })
  const afterColumn = insertColumns(columns, 'main', 0, 1).sheets[0].pageSetup!
  assert.deepEqual(afterColumn.printAreaWhole, { 'B1:C5': '$B:$C' }, 'a column insert moves the whole-column form')
  const rows = workbook(sheet({ pageSetup: { printArea: 'A2:H3', printAreaWhole: { 'A2:H3': '$2:$3' } } }))
  assert.deepEqual(insertRows(rows, 'main', 0, 2).sheets[0].pageSetup!.printAreaWhole, { 'A4:H5': '$4:$5' })
  const plain = insertRows(workbook(sheet({ pageSetup: { printArea: 'A1:B5' } })), 'main', 0, 1).sheets[0].pageSetup!
  assert.equal(plain.printAreaWhole, undefined)
}

process.stdout.write('Sheet operations QA passed: pure row/column edits and cell shifts preserve formulas, ranges, dimensions, panes, filters, and sheet metadata; sheet copies, renames and deletions keep references, tables and names consistent.\n')
