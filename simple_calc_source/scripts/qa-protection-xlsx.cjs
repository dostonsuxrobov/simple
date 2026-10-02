'use strict'

// Sheet protection (flags, password hash) and unlocked cells round-trip through XLSX, and the
// saved password still verifies.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const ExcelJS = require('exceljs')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')
const { hashPassword, verifyPassword, legacyPasswordHash } = require('../electron/protection.cjs')

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
  await legacyPassword()
  process.stdout.write('Protection XLSX QA passed: flags, SHA-512 password, unlocked cells, verification round-trip, and legacy 16-bit sheet passwords.\n')
}

/**
 * calc-file-io-objects-7: a sheet protected with the legacy password attribute (Excel
 * 2007/2010, XlsxWriter, openpyxl, LibreOffice) still needs its password and keeps it on save.
 */
async function legacyPassword() {
  assert.equal(legacyPasswordHash('secret'), 'DAA7')
  assert.equal(legacyPasswordHash('password'), '83AF')
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Payroll').getCell('A1').value = 100
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string')
  zip.file('xl/worksheets/sheet1.xml', xml.replace('</sheetData>', `</sheetData><sheetProtection password="${legacyPasswordHash('secret')}" sheet="1" objects="1" scenarios="1"/>`))
  const source = await zip.generateAsync({ type: 'nodebuffer' })
  const model = (await workbookPayloadFromBytes('legacy.xlsx', source)).workbook
  const protection = model.sheets[0].sheetProtection
  assert.equal(protection.sheet, true)
  assert.equal(protection.password, 'DAA7')
  assert.equal(verifyPassword(protection, 'wrong'), false, 'a wrong password does not unprotect')
  assert.equal(verifyPassword(protection, 'secret'), true)
  for (const baseBytes of [source, null]) {
    const saved = await serializeWorkbook(model, 'xlsx', { baseBytes })
    const tag = /<sheetProtection\b[^>]*\/>/.exec(await (await JSZip.loadAsync(saved)).file('xl/worksheets/sheet1.xml').async('string'))[0]
    assert.match(tag, /password="DAA7"/)
    assert.match(tag, /sheet="1"/)
  }
  // A new modern password replaces the legacy one.
  const reprotected = structuredClone(model)
  reprotected.sheets[0].sheetProtection = { sheet: true, ...hashPassword('new one') }
  const tag = /<sheetProtection\b[^>]*\/>/.exec(await (await JSZip.loadAsync(await serializeWorkbook(reprotected, 'xlsx', { baseBytes: source }))).file('xl/worksheets/sheet1.xml').async('string'))[0]
  assert.doesNotMatch(tag, /password=/)
  assert.match(tag, /algorithmName="SHA-512"/)
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
