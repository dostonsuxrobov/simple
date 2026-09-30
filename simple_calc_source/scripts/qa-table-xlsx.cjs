'use strict'

// Tables created in the editor survive an XLSX save and reload: geometry, header and totals
// rows, totals functions and labels, styles, and whether the filter buttons are shown.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')

function workbook(tables) {
  return {
    version: 1,
    name: 'tables.xlsx',
    activeSheetId: 's1',
    sheets: [{
      id: 's1',
      name: 'Sales',
      state: 'visible',
      rowCount: 20,
      colCount: 8,
      merges: [],
      colWidths: {},
      rowHeights: {},
      cells: {
        A1: { value: 'Item' }, B1: { value: 'Units' }, C1: { value: 'Price' },
        A2: { value: 'Pen' }, B2: { value: 3 }, C2: { value: 2 },
        A3: { value: 'Ink' }, B3: { value: 5 }, C3: { value: 4 },
        A4: { value: 'Total' }, B4: { formula: 'SUBTOTAL(109,Orders[Units])' }, C4: { formula: 'SUBTOTAL(101,Orders[Price])' },
        A7: { value: 'Name' }, B7: { value: 'Score' },
        A8: { value: 'Ann' }, B8: { value: 9 },
        E1: { formula: 'SUM(Orders[Units])' },
      },
      tables,
    }],
  }
}

async function tableXml(bytes) {
  const zip = await JSZip.loadAsync(bytes)
  const names = Object.keys(zip.files).filter((name) => /^xl\/tables\/table\d+\.xml$/.test(name)).sort()
  return Promise.all(names.map((name) => zip.file(name).async('string')))
}

async function main() {
  const model = workbook([
    {
      id: 't1', name: 'Orders', displayName: 'Orders', ref: 'A1:C4', headerRow: true, totalsRow: true,
      columns: [{ name: 'Item', totalsRowLabel: 'Total' }, { name: 'Units', totalsRowFunction: 'sum' }, { name: 'Price', totalsRowFunction: 'average' }],
      style: { theme: 'TableStyleMedium9', showRowStripes: true, showColumnStripes: false, showFirstColumn: true, showLastColumn: false },
      showFilterButton: true,
    },
    {
      id: 't2', name: 'Scores', displayName: 'Scores', ref: 'A7:B8', headerRow: true, totalsRow: false,
      columns: [{ name: 'Name' }, { name: 'Score' }],
      style: { theme: 'TableStyleLight1', showRowStripes: false },
      showFilterButton: false,
    },
  ])
  const bytes = await serializeWorkbook(model, 'xlsx')
  const xml = await tableXml(bytes)
  assert.equal(xml.length, 2)
  const orders = xml.find((text) => text.includes('name="Orders"'))
  const scores = xml.find((text) => text.includes('name="Scores"'))
  assert.match(orders, /ref="A1:C4"/)
  assert.match(orders, /totalsRowCount="1"/)
  assert.match(orders, /headerRowCount="1"/)
  assert.match(orders, /<autoFilter ref="A1:C3"/, 'the AutoFilter covers the header and data rows only')
  assert.match(orders, /totalsRowFunction="sum"/)
  assert.match(orders, /totalsRowFunction="average"/)
  assert.match(orders, /totalsRowLabel="Total"/)
  assert.match(orders, /TableStyleMedium9/)
  assert.match(orders, /showFirstColumn="1"/)
  assert.doesNotMatch(scores, /<autoFilter/, 'hidden filter buttons write no AutoFilter')

  const payload = await workbookPayloadFromBytes('tables.xlsx', bytes)
  const tables = payload.workbook.sheets[0].tables
  const reOrders = tables.find((table) => table.name === 'Orders')
  const reScores = tables.find((table) => table.name === 'Scores')
  assert.equal(reOrders.ref, 'A1:C4')
  assert.equal(reOrders.totalsRow, true)
  assert.equal(reOrders.headerRow, true)
  assert.deepEqual(reOrders.columns.map((column) => column.totalsRowFunction || column.totalsRowLabel), ['Total', 'sum', 'average'])
  assert.equal(reOrders.style.theme, 'TableStyleMedium9')
  assert.equal(reOrders.showFilterButton, undefined, 'filter buttons stay on')
  assert.equal(reScores.showFilterButton, false, 'hidden filter buttons are read back')
  assert.equal(payload.workbook.sheets[0].cells.E1.formula, 'SUM(Orders[Units])')
  process.stdout.write('Table XLSX QA passed: geometry, totals functions and labels, styles, and filter buttons round-trip.\n')
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
