import assert from 'node:assert/strict'
import { CalculationEngine } from '../src/lib/calc-engine'
import { applySheetStructureOperation } from '../src/lib/sheet-operations'
import {
  calculatedColumnFormula,
  convertTableToRange,
  createTable,
  expandTableForEntry,
  fillCalculatedColumn,
  renameStructuredReferences,
  renameTable,
  renameTablesInCopiedSheet,
  resizeTable,
  setTotalsFunction,
  setTotalsRow,
  structuredReferencesToA1,
  syncTableHeaders,
  tableRegions,
  uniqueColumnNames,
  validateTableName,
} from '../src/lib/tables'
import type { CellData, SheetData, SheetTable, WorkbookModel } from '../src/spreadsheet-types'

function sheet(id: string, name: string, cells: Record<string, CellData>, extra: Partial<SheetData> = {}): SheetData {
  return { id, name, rowCount: 100, colCount: 26, cells, merges: [], colWidths: {}, rowHeights: {}, ...extra }
}
function workbook(...sheets: SheetData[]): WorkbookModel {
  return { version: 1, name: 'tables.xlsx', activeSheetId: sheets[0].id, sheets }
}
const v = (value: CellData['value']): CellData => ({ value })
const f = (formula: string): CellData => ({ formula })
const created = (result: SheetTable | string) => {
  assert.notEqual(typeof result, 'string', String(result))
  return result as SheetTable
}

// ---- Naming -----------------------------------------------------------------------------------
assert.deepEqual(uniqueColumnNames(['Item', '', 'Item', 'item', '']), ['Item', 'Column2', 'Item2', 'item3', 'Column5'])
{
  const book = workbook(sheet('s1', 'Sheet1', {}))
  assert.equal(validateTableName(book, 'Sales'), null)
  assert.match(String(validateTableName(book, 'A1')), /cell reference/)
  assert.match(String(validateTableName(book, 'R1C1')), /cell reference/)
  assert.match(String(validateTableName(book, 'my table')), /letters, numbers/)
  assert.match(String(validateTableName(book, '1st')), /start with/)
  book.definedNames = [{ name: 'Rates', ref: 'Sheet1!$A$1' }]
  assert.match(String(validateTableName(book, 'rates')), /already used/)
}

// ---- Creation --------------------------------------------------------------------------------
{
  const data = sheet('s1', 'Sheet1', {
    A1: v('Item'), B1: v('Units'), C1: v(''), D1: v('Units'),
    A2: v('Pen'), B2: v(3), D2: v(1),
    A3: v('Ink'), B3: v(5), D3: v(2),
  })
  const book = workbook(data)
  const table = created(createTable(book, data, { top: 0, bottom: 2, left: 0, right: 3 }, { hasHeaders: true }))
  assert.equal(table.name, 'Table1')
  assert.equal(table.ref, 'A1:D3')
  assert.deepEqual(table.columns.map((column) => column.name), ['Item', 'Units', 'Column3', 'Units2'])
  assert.equal(data.cells.C1.value, 'Column3', 'blank headers are written back')
  assert.equal(data.cells.D1.value, 'Units2')
  assert.equal(table.style?.theme, 'TableStyleMedium2')
  assert.match(String(createTable(book, data, { top: 2, bottom: 4, left: 3, right: 4 }, { hasHeaders: true })), /overlap/)
  const second = created(createTable(book, data, { top: 10, bottom: 12, left: 0, right: 1 }, { hasHeaders: true }))
  assert.equal(second.name, 'Table2')
}
{
  // Without headers Excel inserts a header row and shifts the data down.
  const data = sheet('s1', 'Sheet1', { A1: v(1), B1: v(2), A2: v(3), B2: v(4), C1: f('SUM(A1:B2)') })
  const book = workbook(data)
  const table = created(createTable(book, data, { top: 0, bottom: 1, left: 0, right: 1 }, { hasHeaders: false }))
  assert.equal(table.ref, 'A1:B3')
  assert.deepEqual(table.columns.map((column) => column.name), ['Column1', 'Column2'])
  assert.equal(data.cells.A2.value, 1)
  assert.equal(data.cells.B3.value, 4)
  assert.equal(data.cells.C1.formula, 'SUM(A2:B3)', 'references follow the shifted block')
}

// ---- Totals row -------------------------------------------------------------------------------
{
  const data = sheet('s1', 'Sheet1', {
    A1: v('Item'), B1: v('Amount'),
    A2: v('Pen'), B2: v(3),
    A3: v('Ink'), B3: v(5),
    A4: v('below'),
  })
  const book = workbook(data)
  const table = created(createTable(book, data, { top: 0, bottom: 2, left: 0, right: 1 }, { hasHeaders: true, name: 'Sales' }))
  setTotalsRow(book, data, table, true)
  assert.equal(table.ref, 'A1:B4')
  assert.equal(data.cells.A4.value, 'Total')
  assert.equal(data.cells.B4.formula, 'SUBTOTAL(109,Sales[Amount])')
  assert.equal(data.cells.A5.value, 'below', 'occupied cells below shift down')
  const engine = new CalculationEngine(book)
  assert.equal(engine.getValue('s1', 'B4'), 8)
  setTotalsFunction(data, table, 1, 'average')
  assert.equal(data.cells.B4.formula, 'SUBTOTAL(101,Sales[Amount])')
  setTotalsRow(book, data, table, false)
  assert.equal(table.ref, 'A1:B3')
  assert.equal(data.cells.B4, undefined)
  setTotalsRow(book, data, table, true)
  assert.equal(data.cells.B4.formula, 'SUBTOTAL(101,Sales[Amount])', 'totals functions are remembered')
}

// ---- Header sync and structured-reference renames ---------------------------------------------
{
  const data = sheet('s1', 'Sheet1', {
    A1: v('Item'), B1: v('Amount'), C1: v('Double'),
    A2: v('Pen'), B2: v(3), C2: f('[@Amount]*2'),
    A3: v('Ink'), B3: v(5), C3: f('[@Amount]*2'),
    E1: f('SUM(Sales[Amount])'), E2: f('Sales[[#Headers],[Amount]]'), E3: f('SUM(Sales[@[Amount]:[Double]])'),
  })
  const book = workbook(data)
  created(createTable(book, data, { top: 0, bottom: 2, left: 0, right: 2 }, { hasHeaders: true, name: 'Sales' }))
  data.cells.B1 = v('Revenue')
  assert.equal(syncTableHeaders(book, data, ['B1']), true)
  assert.equal(data.cells.E1.formula, 'SUM(Sales[Revenue])')
  assert.equal(data.cells.E2.formula, 'Sales[[#Headers],[Revenue]]')
  assert.equal(data.cells.E3.formula, 'SUM(Sales[@[Revenue]:[Double]])')
  assert.equal(data.cells.C2.formula, '[@Revenue]*2', 'unqualified references inside the table follow')
  assert.equal(syncTableHeaders(book, data, ['Z9']), false, 'edits outside header rows are ignored')
  data.cells.C1 = v('revenue')
  syncTableHeaders(book, data, ['C1'])
  assert.equal(data.cells.C1.value, 'revenue2', 'duplicate header text gets a suffix')
  assert.equal(renameTable(book, data.tables![0].id, 'Orders'), null)
  assert.equal(data.cells.E1.formula, 'SUM(Orders[Revenue])')
  assert.match(String(renameTable(book, data.tables![0].id, 'B2')), /cell reference/)
  assert.equal(renameStructuredReferences('"Sales[Amount]"&Sales[Amount]', { tables: new Map([['sales', 'X']]) }), '"Sales[Amount]"&X[Amount]', 'strings are untouched')
  const engine = new CalculationEngine(book)
  assert.equal(engine.getValue('s1', 'E1'), 8)
  assert.equal(engine.getValue('s1', 'C3'), 10)
}

// ---- Calculated columns and auto-expansion -------------------------------------------------------
{
  const data = sheet('s1', 'Sheet1', {
    A1: v('Item'), B1: v('Units'), C1: v('Double'),
    A2: v('Pen'), B2: v(3),
    A3: v('Ink'), B3: v(5),
    A4: v('Pad'), B4: v(7),
  })
  const book = workbook(data)
  const table = created(createTable(book, data, { top: 0, bottom: 3, left: 0, right: 2 }, { hasHeaders: true }))
  data.cells.C3 = f('B3*2')
  assert.equal(fillCalculatedColumn(data, table, 2, 2), true)
  assert.equal(data.cells.C2.formula, 'B2*2')
  assert.equal(data.cells.C4.formula, 'B4*2')
  assert.equal(calculatedColumnFormula(data, table, 2), 'B2*2')
  data.cells.C2 = f('B2*3')
  assert.equal(fillCalculatedColumn(data, table, 1, 2), true, 'a consistent calculated column is rewritten')
  assert.equal(data.cells.C4.formula, 'B4*3')
  data.cells.A3 = f('"x"')
  assert.equal(fillCalculatedColumn(data, table, 2, 0), false, 'a column with values is left alone')

  // Typing below the table extends it and carries the calculated column.
  data.cells.A5 = v('Tape')
  assert.equal(expandTableForEntry(data, 4, 0), true)
  assert.equal(table.ref, 'A1:C5')
  assert.equal(data.cells.C5.formula, 'B5*3')
  // Typing to the right adds a column named from the header or ColumnN.
  data.cells.D3 = v(1)
  assert.equal(expandTableForEntry(data, 2, 3), true)
  assert.equal(table.ref, 'A1:D5')
  assert.equal(data.cells.D1.value, 'Column4')
  assert.deepEqual(table.columns.map((column) => column.name), ['Item', 'Units', 'Double', 'Column4'])
  data.cells.E1 = v('Notes')
  assert.equal(expandTableForEntry(data, 0, 4), true)
  assert.equal(table.columns[4].name, 'Notes')
  data.cells.H9 = v(1)
  assert.equal(expandTableForEntry(data, 8, 7), false)
  assert.equal(tableRegions(table)!.dataBottom, 4)
}

// ---- Resize, structure edits, convert -------------------------------------------------------------
{
  const data = sheet('s1', 'Sheet1', {
    A1: v('Item'), B1: v('Units'), C1: v('Price'),
    A2: v('Pen'), B2: v(3), C2: v(2),
    A3: v('Ink'), B3: v(5), C3: v(4),
    E2: f('SUMPRODUCT(Sales[Units],Sales[Price])'), E3: f('Sales[@Units]'),
  })
  let book = workbook(data)
  const table = created(createTable(book, data, { top: 0, bottom: 2, left: 0, right: 2 }, { hasHeaders: true, name: 'Sales' }))
  assert.match(String(resizeTable(data, table, { top: 1, bottom: 2, left: 0, right: 2 })), /header row/)
  assert.equal(resizeTable(data, table, { top: 0, bottom: 3, left: 0, right: 2 }), null)
  assert.equal(table.ref, 'A1:C4')
  assert.equal(resizeTable(data, table, { top: 0, bottom: 2, left: 0, right: 2 }), null)

  book = applySheetStructureOperation(book, 's1', { axis: 'row', kind: 'insert', index: 2, count: 2 })
  assert.equal(book.sheets[0].tables![0].ref, 'A1:C5', 'rows inserted inside a table extend it')
  book = applySheetStructureOperation(book, 's1', { axis: 'column', kind: 'insert', index: 1, count: 1 })
  const widened = book.sheets[0].tables![0]
  assert.equal(widened.ref, 'A1:D5')
  assert.deepEqual(widened.columns.map((column) => column.name), ['Item', 'Column4', 'Units', 'Price'])
  assert.equal(book.sheets[0].cells.B1.value, 'Column4', 'the new header cell is named')
  book = applySheetStructureOperation(book, 's1', { axis: 'column', kind: 'delete', index: 1, count: 1 })
  assert.deepEqual(book.sheets[0].tables![0].columns.map((column) => column.name), ['Item', 'Units', 'Price'])
  book = applySheetStructureOperation(book, 's1', { axis: 'row', kind: 'delete', index: 2, count: 2 })
  assert.equal(book.sheets[0].tables![0].ref, 'A1:C3')

  const converted = structuredClone(book)
  convertTableToRange(converted, converted.sheets[0].tables![0].id)
  assert.equal(converted.sheets[0].tables!.length, 0)
  assert.equal(converted.sheets[0].cells.E2.formula, 'SUMPRODUCT($B$2:$B$3,$C$2:$C$3)')
  assert.equal(converted.sheets[0].cells.E3.formula, '$B3', 'a this-row reference beside the table intersects its row, as in Excel')
  assert.equal(structuredReferencesToA1('[@Units]*2', { row: 2, col: 3 }, book, book.sheets[0]), '[@Units]*2')
  assert.equal(converted.sheets[0].cells.A1.style?.font?.bold, true, 'the header look is kept as direct formatting')
  assert.ok(converted.sheets[0].cells.A1.style?.fill, 'the header fill is kept')
  const engine = new CalculationEngine(converted)
  assert.equal(engine.getValue('s1', 'E2'), 26)

  book = applySheetStructureOperation(book, 's1', { axis: 'row', kind: 'delete', index: 0, count: 3 })
  assert.equal(book.sheets[0].tables!.length, 0, 'deleting every table row removes the table')
}

// ---- Convert keeps this-row references working inside the table -------------------------------
{
  const data = sheet('s1', 'Sheet1', {
    A1: v('Units'), B1: v('Double'),
    A2: v(3), B2: f('[@Units]*2'),
    A3: v(5), B3: f('Sales[@Units]*2'),
  })
  const book = workbook(data)
  created(createTable(book, data, { top: 0, bottom: 2, left: 0, right: 1 }, { hasHeaders: true, name: 'Sales' }))
  convertTableToRange(book, data.tables![0].id)
  assert.equal(data.cells.B2.formula, '$A2*2')
  assert.equal(data.cells.B3.formula, '$A3*2')
}

// ---- A copied sheet's tables get unique names (calc-file-io-objects-10) ------------------------
{
  const totals = { ...v(12), formula: 'SUBTOTAL(109,Sales[Qty])' }
  const original = sheet('s1', 'Data', {
    A1: v('Name'), B1: v('Qty'), A2: v('a'), B2: v(5), A3: v('b'), B3: v(7), A4: v('Total'), B4: totals,
    C2: f('[@Qty]*2'), D1: f('COUNT(Sales[Qty])+COUNT(Other[Qty])'),
  }, {
    tables: [
      { id: 'table-Sales', name: 'Sales', displayName: 'Sales', ref: 'A1:B4', headerRow: true, totalsRow: true, columns: [{ name: 'Name', totalsRowLabel: 'Total' }, { name: 'Qty', totalsRowFunction: 'sum' }] },
      { id: 'table-Table1', name: 'Table1', displayName: 'Table1', ref: 'F1:F3', headerRow: true, totalsRow: false, columns: [{ name: 'Col' }] },
    ],
    conditionalFormattings: [{ ref: 'B2:B3', rules: [{ type: 'expression', formulae: ['B2>AVERAGE(Sales[Qty])'] }] }],
  })
  const other = sheet('s2', 'Other', {}, { tables: [{ id: 'table-Other', name: 'Other', ref: 'A1:A2', headerRow: true, totalsRow: false, columns: [{ name: 'Qty' }] }, { id: 'table-Sales2', name: 'Sales2', ref: 'C1:C2', headerRow: true, totalsRow: false, columns: [{ name: 'X' }] }] })
  const book = workbook(original, other)
  const copy = structuredClone(original)
  let id = 0
  const renames = renameTablesInCopiedSheet(book, copy, (prefix) => `${prefix}-copy-${++id}`)
  assert.deepEqual([...renames.entries()], [['sales', 'Sales3'], ['table1', 'Table2']], 'Sales2 is taken elsewhere, so the copy is Sales3')
  assert.deepEqual(copy.tables?.map((table) => [table.id, table.name, table.displayName]), [['table-copy-1', 'Sales3', 'Sales3'], ['table-copy-2', 'Table2', 'Table2']])
  assert.equal(copy.cells.B4.formula, 'SUBTOTAL(109,Sales3[Qty])', 'the totals row follows the renamed table')
  assert.equal(copy.cells.C2.formula, '[@Qty]*2', 'this-row references stay unqualified')
  assert.equal(copy.cells.D1.formula, 'COUNT(Sales3[Qty])+COUNT(Other[Qty])', 'references to tables elsewhere stay')
  assert.deepEqual((copy.conditionalFormattings?.[0] as { rules: Array<{ formulae: string[] }> }).rules[0].formulae, ['B2>AVERAGE(Sales3[Qty])'])
  assert.equal(original.cells.B4.formula, 'SUBTOTAL(109,Sales[Qty])', 'the original sheet keeps its names')
  assert.equal(validateTableName({ sheets: [original, other, copy] }, 'Sales3'), 'The name "Sales3" is already used in this workbook.')
}

console.log('Tables QA passed: naming, creation, totals, header sync, renames, calculated columns, auto-expand, resize, structure edits, convert to range, copied-sheet names.')
