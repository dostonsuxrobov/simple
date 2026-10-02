'use strict'

// Data validations keep formula/reference bounds, date bounds, and the hidden-dropdown flag
// through an XLSX save and reload.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const ExcelJS = require('exceljs')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')

async function main() {
  const model = {
    version: 1,
    name: 'validation.xlsx',
    activeSheetId: 's1',
    sheets: [{
      id: 's1', name: 'Rules', state: 'visible', rowCount: 20, colCount: 10, merges: [], colWidths: {}, rowHeights: {},
      cells: { Z1: { value: 10 }, Y1: { value: 45000 }, A1: { value: 'x' } },
      dataValidations: {
        B2: { type: 'whole', operator: 'between', allowBlank: true, formulae: ['$Z$1', '100'] },
        C2: { type: 'date', operator: 'greaterThan', allowBlank: true, formulae: ['TODAY()'] },
        D2: { type: 'date', operator: 'between', allowBlank: true, formulae: [{ type: 'date', value: '2024-01-01T00:00:00.000Z' }, '$Y$1'] },
        E2: { type: 'list', allowBlank: true, showDropDown: true, formulae: ['"Low,High"'] },
        F2: { type: 'decimal', operator: 'lessThan', formulae: [2.5] },
        G2: { type: 'textLength', operator: 'lessThanOrEqual', formulae: ['LEN($A$1)+3'] },
      },
    }],
  }
  const bytes = await serializeWorkbook(model, 'xlsx')
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string')
  const rule = (ref) => {
    const match = new RegExp(`<dataValidation[^>]*sqref="${ref}"[^>]*>([\\s\\S]*?)</dataValidation>`).exec(xml)
    assert.ok(match, `validation for ${ref}`)
    return match[0]
  }
  assert.match(rule('B2'), /<formula1>\$Z\$1<\/formula1><formula2>100<\/formula2>/)
  assert.match(rule('C2'), /<formula1>TODAY\(\)<\/formula1>/)
  assert.match(rule('D2'), /<formula1>45292<\/formula1><formula2>\$Y\$1<\/formula2>/)
  assert.match(rule('E2'), /showDropDown="1"/)
  assert.match(rule('F2'), /<formula1>2.5<\/formula1>/)
  assert.match(rule('G2'), /<formula1>LEN\(\$A\$1\)\+3<\/formula1>/)
  assert.doesNotMatch(xml, /NaN/)

  const payload = await workbookPayloadFromBytes('validation.xlsx', bytes)
  const read = payload.workbook.sheets[0].dataValidations
  assert.deepEqual(read.B2.formulae, ['$Z$1', 100])
  assert.deepEqual(read.C2.formulae, ['TODAY()'])
  assert.equal(read.D2.formulae[1], '$Y$1')
  const firstDate = read.D2.formulae[0]
  const firstIso = firstDate instanceof Date ? firstDate.toISOString() : firstDate && firstDate.value
  assert.equal(String(firstIso).slice(0, 10), '2024-01-01')
  assert.equal(read.E2.showDropDown, true)
  assert.deepEqual(read.G2.formulae, ['LEN($A$1)+3'])

  // A second save keeps everything the first one wrote.
  const again = await serializeWorkbook(payload.workbook, 'xlsx')
  const xml2 = await (await JSZip.loadAsync(again)).file('xl/worksheets/sheet1.xml').async('string')
  assert.doesNotMatch(xml2, /NaN/)
  assert.match(xml2, /TODAY\(\)/)
  assert.match(xml2, /showDropDown="1"/)
  await largeRanges()
  process.stdout.write('Validation XLSX QA passed: formula and reference bounds, date bounds, hidden dropdowns, and rules on very large ranges round-trip and stay editable.\n')
}

/**
 * CALC-SIE-12: a rule on a whole column is one range rule, not tens of thousands of cell
 * entries, so editing, clearing or shifting it on such a sheet is saved.
 */
async function largeRanges() {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Big')
  sheet.getCell('A1').value = 'Status'
  sheet.dataValidations.add('D2:D20001', { type: 'list', allowBlank: true, formulae: ['"Open,Closed"'] })
  sheet.dataValidations.add('E2', { type: 'whole', operator: 'between', formulae: [1, 10] })
  const source = Buffer.from(await workbook.xlsx.writeBuffer())
  const started = Date.now()
  const payload = await workbookPayloadFromBytes('big.xlsx', source)
  assert.ok(Date.now() - started < 10_000, 'a whole-column rule opens quickly')
  const model = payload.workbook.sheets[0]
  assert.deepEqual(Object.keys(model.dataValidations).sort(), ['D2:D20001', 'E2'])
  assert.equal(model.dataValidationsTruncated, undefined)
  assert.ok(!payload.warnings.some((warning) => /validation/i.test(warning)))
  // Clear the rule and save over the source: it is gone after reopening.
  const cleared = structuredClone(payload.workbook)
  delete cleared.sheets[0].dataValidations['D2:D20001']
  const clearedBytes = await serializeWorkbook(cleared, 'xlsx', { baseBytes: source })
  const clearedXml = await (await JSZip.loadAsync(clearedBytes)).file('xl/worksheets/sheet1.xml').async('string')
  assert.doesNotMatch(clearedXml, /sqref="D2:D20001"/)
  assert.deepEqual(Object.keys((await workbookPayloadFromBytes('big.xlsx', clearedBytes)).workbook.sheets[0].dataValidations), ['E2'])
  // All rules removed (the key dropped): nothing comes back from the source package.
  const none = structuredClone(payload.workbook)
  delete none.sheets[0].dataValidations
  const noneXml = await (await JSZip.loadAsync(await serializeWorkbook(none, 'xlsx', { baseBytes: source }))).file('xl/worksheets/sheet1.xml').async('string')
  assert.doesNotMatch(noneXml, /<dataValidation\b/)
  // Five rows inserted at the top: the editor shifts the range key; the save writes it.
  const shifted = structuredClone(payload.workbook)
  const rule = shifted.sheets[0].dataValidations['D2:D20001']
  shifted.sheets[0].dataValidations = { 'D7:D20006': rule, E7: shifted.sheets[0].dataValidations.E2 }
  const shiftedXml = await (await JSZip.loadAsync(await serializeWorkbook(shifted, 'xlsx', { baseBytes: source }))).file('xl/worksheets/sheet1.xml').async('string')
  assert.match(shiftedXml, /sqref="D7:D20006"/)
  assert.match(shiftedXml, /sqref="E7"/)
  // A multi-area sqref keeps every area; cell lookups still find the rule inside a range.
  const multi = await workbookPayloadFromBytes('multi.xlsx', await patchSqref(source, 'D2:D20001', 'D2:D5 F2:F5'))
  assert.deepEqual(Object.keys(multi.workbook.sheets[0].dataValidations).sort(), ['D2:D5', 'E2', 'F2:F5'])
  const reread = new ExcelJS.Workbook()
  await reread.xlsx.load(source)
  assert.equal(reread.getWorksheet('Big').getCell('D500').dataValidation.type, 'list')
}

async function patchSqref(bytes, from, to) {
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string')
  zip.file('xl/worksheets/sheet1.xml', xml.replace(`sqref="${from}"`, `sqref="${to}"`))
  return zip.generateAsync({ type: 'nodebuffer' })
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
