'use strict'
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')
;(async () => {
  const source = {
    version: 1, name: 'Fractional fonts', activeSheetId: 's1', metadata: { normalFont: { name: 'Arial', size: 9.5 } },
    sheets: [{ id: 's1', name: 'Fonts', rowCount: 5, colCount: 3, merges: [],
      cells: { A1: { value: 'Half point', style: { font: { name: 'Cambria', size: 13.5 } } }, A2: { value: 'Quarter point', style: { font: { name: 'Cambria', size: 13.25 } } }, A3: { value: 'Whole point', style: { font: { name: 'Cambria', size: 13 } } }, A4: { value: 'Row font' }, B1: { value: 'Column font' } },
      rowProperties: { 4: { style: { font: { name: 'Arial', size: 8.5 } } } },
      columnProperties: { 2: { style: { font: { name: 'Arial', size: 17.75 } } } },
    }],
  }
  let bytes = await serializeWorkbook(source, 'xlsx')
  const xml = await (await JSZip.loadAsync(bytes)).file('xl/styles.xml').async('string')
  assert.match(xml, /<sz val="13\.5"/)
  assert.match(xml, /<sz val="13\.25"/)
  for (let iteration = 0; iteration < 2; iteration++) {
    const { workbook } = await workbookPayloadFromBytes('fractional.xlsx', bytes), sheet = workbook.sheets[0]
    assert.equal(sheet.cells.A1.style.font.size, 13.5)
    assert.equal(sheet.cells.A2.style.font.size, 13.25, 'same family/rounded size still resolves the correct font ID')
    assert.equal(sheet.cells.A3.style.font.size, 13)
    assert.equal(sheet.rowProperties['4'].style.font.size, 8.5)
    assert.equal(sheet.columnProperties['2'].style.font.size, 17.75)
    assert.equal(sheet.cells.B1.style.font.size, 17.75)
    assert.equal(workbook.metadata.normalFont.size, 9.5)
    sheet.cells.A1.value = 'Edited without changing font'
    bytes = await serializeWorkbook(workbook, 'xlsx', { baseBytes: bytes })
  }
  console.log('Fractional font QA passed: exact13.5/13.25 cell fonts, integer font, row/column defaults, Normal font and two source-backed save/reopen cycles.')
})().catch(error => { console.error(error); process.exitCode = 1 })
