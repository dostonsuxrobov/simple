'use strict'

// Data validations keep formula/reference bounds, date bounds, and the hidden-dropdown flag
// through an XLSX save and reload.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
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
  process.stdout.write('Validation XLSX QA passed: formula and reference bounds, date bounds, and hidden dropdowns round-trip.\n')
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
