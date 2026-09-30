'use strict'

// Outline levels and collapsed summary flags round-trip through XLSX exactly.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')

async function main() {
  const model = {
    version: 1, name: 'outline.xlsx', activeSheetId: 's1',
    sheets: [{
      id: 's1', name: 'Plan', state: 'visible', rowCount: 20, colCount: 10, merges: [], colWidths: {}, rowHeights: {},
      cells: { A1: { value: 'Header' }, A2: { value: 1 }, A3: { value: 2 }, A4: { value: 3 }, A5: { formula: 'SUM(A2:A4)' } },
      rowProperties: { 2: { outlineLevel: 1 }, 3: { outlineLevel: 2 }, 4: { outlineLevel: 2 }, 5: { collapsed: true }, 8: { outlineLevel: 1 } },
      columnProperties: { 2: { outlineLevel: 1 }, 3: { outlineLevel: 1 }, 4: { collapsed: true } },
      hiddenRows: [3, 4, 9],
      hiddenCols: [2, 3],
    }],
  }
  const bytes = await serializeWorkbook(model, 'xlsx')
  const xml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml').async('string')
  assert.match(xml, /outlineLevelRow="2"/)
  assert.match(xml, /outlineLevelCol="1"/)
  const row = (number) => new RegExp(`<row r="${number}"[^>]*>`).exec(xml)?.[0] || ''
  assert.doesNotMatch(row(3), /collapsed/, 'detail rows are not marked collapsed')
  assert.doesNotMatch(row(4), /collapsed/)
  assert.match(row(5), /collapsed="1"/, 'the summary row keeps its flag')
  assert.match(row(3), /outlineLevel="2"/)
  const payload = await workbookPayloadFromBytes('outline.xlsx', bytes)
  const sheet = payload.workbook.sheets[0]
  assert.equal(sheet.rowProperties['3'].outlineLevel, 2)
  assert.equal(sheet.rowProperties['3'].collapsed, undefined)
  assert.equal(sheet.rowProperties['5'].collapsed, true)
  assert.equal(sheet.columnProperties['2'].outlineLevel, 1)
  assert.equal(sheet.columnProperties['2'].collapsed, undefined)
  assert.equal(sheet.columnProperties['4'].collapsed, true)
  assert.deepEqual(sheet.hiddenRows, [3, 4, 9], 'an empty hidden row survives')
  assert.equal(sheet.rowProperties['8'].outlineLevel, 1, 'an empty grouped row survives')
  process.stdout.write('Outline XLSX QA passed: levels, depths, and collapsed summary flags round-trip.\n')
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
