'use strict'

// Sheet protection (flags, password hash) and unlocked cells round-trip through XLSX, and the
// saved password still verifies.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')
const { hashPassword, verifyPassword } = require('../electron/protection.cjs')

async function main() {
  const hash = hashPassword('open sesame')
  const model = {
    version: 1, name: 'protected.xlsx', activeSheetId: 's1',
    sheets: [{
      id: 's1', name: 'Form', state: 'visible', rowCount: 10, colCount: 5, merges: [], colWidths: {}, rowHeights: {},
      cells: { A1: { value: 'Name' }, B1: { value: 'Ann', style: { protection: { locked: false } } } },
      sheetProtection: { sheet: true, ...hash, formatCells: true, autoFilter: true, objects: false, scenarios: false },
    }],
  }
  const bytes = await serializeWorkbook(model, 'xlsx')
  const xml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml').async('string')
  const tag = /<sheetProtection\b[^>]*\/>/.exec(xml)[0]
  assert.match(tag, /sheet="1"/)
  assert.match(tag, /algorithmName="SHA-512"/)
  assert.match(tag, /formatCells="0"/, 'allowed formatting is written as formatCells="0"')
  assert.match(tag, /autoFilter="0"/)
  assert.doesNotMatch(tag, /insertRows="0"/, 'row insertion stays protected')
  assert.match(tag, /objects="1"/)
  const reopened = await workbookPayloadFromBytes('protected.xlsx', bytes)
  const sheet = reopened.workbook.sheets[0]
  assert.equal(sheet.sheetProtection.sheet, true)
  assert.equal(sheet.sheetProtection.formatCells, true)
  assert.equal(sheet.sheetProtection.objects, false)
  assert.equal(sheet.cells.B1.style.protection.locked, false)
  assert.equal(verifyPassword(sheet.sheetProtection, 'open sesame'), true)
  assert.equal(verifyPassword(sheet.sheetProtection, 'Open sesame'), false)
  process.stdout.write('Protection XLSX QA passed: flags, SHA-512 password, unlocked cells, and verification round-trip.\n')
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
