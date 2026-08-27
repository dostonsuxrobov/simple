import assert from 'node:assert/strict'
import type { SheetData, WorkbookModel } from '../src/spreadsheet-types.ts'
import {
  applySelectionStructureCommand,
  deleteColumns,
  deleteRows,
  insertBlankSheet,
  insertColumns,
  insertRows,
  rewriteFormulaForSheetStructure,
  SheetStructureError,
} from '../src/lib/sheet-operations.ts'

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
    conditionalFormattings: [{ ref: 'E2:E4', rules: [{ type: 'expression', formulae: ['B2>0'] }] }],
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

process.stdout.write('Sheet operations QA passed: pure row/column edits preserve formulas, ranges, dimensions, panes, and sheet metadata.\n')
