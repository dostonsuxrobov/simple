'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const ExcelJS = require('exceljs')
const XLSX = require('xlsx')
const {
  workbookPayloadFromPath,
  workbookPayloadFromBytes,
  serializeWorkbook,
} = require('../electron/workbooks.cjs')

async function createRichFixture(filePath) {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'simple_calc QA'
  workbook.created = new Date('2025-01-02T00:00:00Z')
  workbook.calcProperties.fullCalcOnLoad = true
  const summary = workbook.addWorksheet('Summary', { views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }] })
  summary.columns = [{ width: 18 }, { width: 15 }, { width: 14 }, { width: 14 }]
  summary.getRow(1).height = 25
  summary.getCell('A1').value = 'Revenue'
  summary.getCell('A1').font = { bold: true, color: { argb: 'FFFFFFFF' } }
  summary.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF476B57' } }
  summary.getCell('A2').value = 1200
  summary.getCell('A3').value = 800
  summary.getCell('A4').value = { formula: 'SUM(A2:A3)', result: 2000 }
  summary.getCell('A4').numFmt = '$#,##0.00'
  summary.getCell('B2').value = { formula: 'A2*20%', result: 240 }
  summary.getCell('B2').note = 'Calculated tax'
  summary.getCell('B3').value = { text: 'OpenAI', hyperlink: 'https://openai.com' }
  summary.getCell('C1').value = 'Quarterly model'
  summary.mergeCells('C1:D1')
  const inputs = workbook.addWorksheet('Inputs')
  inputs.getCell('A1').value = 'Rate'
  inputs.getCell('A2').value = 0.2
  inputs.getCell('A2').numFmt = '0%'
  const hidden = workbook.addWorksheet('Hidden data')
  hidden.state = 'hidden'
  hidden.getCell('A1').value = 'kept'
  workbook.views = [{ activeTab: 0, firstSheet: 0, visibility: 'visible' }]
  await workbook.xlsx.writeFile(filePath)
}

async function main() {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-calc-qa-'))
  try {
    const richPath = path.join(directory, 'rich.xlsx')
    await createRichFixture(richPath)
    const opened = await workbookPayloadFromPath(richPath)
    assert.equal(opened.sourceFormat, 'xlsx')
    assert.equal(opened.requiresSaveAs, false)
    assert.equal(opened.workbook.sheets.length, 3)
    const summary = opened.workbook.sheets.find((sheet) => sheet.name === 'Summary')
    assert(summary)
    assert.equal(summary.cells.A4.formula, 'SUM(A2:A3)')
    assert.equal(summary.cells.A4.result, 2000)
    assert.equal(summary.cells.A1.style.font.bold, true)
    assert.equal(summary.cells.B2.note, 'Calculated tax')
    assert.equal(summary.cells.B3.hyperlink, 'https://openai.com')
    assert.deepEqual(summary.merges, ['C1:D1'])
    assert.equal(summary.frozen.rows, 1)
    assert.equal(summary.frozen.columns, 1)

    summary.cells.A2.value = 1500
    const savedBytes = await serializeWorkbook(opened.workbook, 'xlsx')
    assert(Buffer.isBuffer(savedBytes) && savedBytes.length > 1000)
    const reopened = await workbookPayloadFromBytes('roundtrip.xlsx', savedBytes)
    const reopenedSummary = reopened.workbook.sheets.find((sheet) => sheet.name === 'Summary')
    assert.equal(reopenedSummary.cells.A2.value, 1500)
    assert.equal(reopenedSummary.cells.A4.formula, 'SUM(A2:A3)')
    assert.equal(reopenedSummary.cells.A1.style.font.bold, true)
    assert.equal(reopenedSummary.cells.B2.note, 'Calculated tax')
    assert.equal(reopenedSummary.cells.B3.hyperlink, 'https://openai.com')
    assert.deepEqual(reopenedSummary.merges, ['C1:D1'])
    assert.equal(reopened.workbook.sheets.find((sheet) => sheet.name === 'Hidden data').state, 'hidden')

    const csv = await serializeWorkbook(opened.workbook, 'csv')
    assert.match(csv.toString('utf8'), /Revenue/)
    assert.match(csv.toString('utf8'), /2,000\.00|2000/)
    const tsv = await serializeWorkbook(opened.workbook, 'tsv')
    assert.match(tsv.toString('utf8'), /\t/)
    const ods = await serializeWorkbook(opened.workbook, 'ods')
    const odsOpened = await workbookPayloadFromBytes('roundtrip.ods', ods)
    assert.equal(odsOpened.workbook.sheets[0].cells.A2.value, 1500)

    const legacyWorkbook = XLSX.utils.book_new()
    const legacySheet = XLSX.utils.aoa_to_sheet([
      ['Item', 'Price', 'Total'],
      ['Book', 12.5, { t: 'n', f: 'B2*2', v: 25 }],
      ['Leading zero', '00123', null],
    ])
    legacySheet['!merges'] = [XLSX.utils.decode_range('A4:C4')]
    XLSX.utils.book_append_sheet(legacyWorkbook, legacySheet, 'Legacy')
    const xlsBytes = XLSX.write(legacyWorkbook, { type: 'buffer', bookType: 'biff8' })
    const xlsOpened = await workbookPayloadFromBytes('legacy.xls', xlsBytes)
    assert.equal(xlsOpened.requiresSaveAs, true)
    // SheetJS' legacy BIFF writer does not emit formula records; real XLS
    // formula imports are covered by the decoder, while this generated fixture
    // still exercises values, text types, merges, and XLS-to-XLSX conversion.
    assert.equal(xlsOpened.workbook.sheets[0].cells.C2.value, 25)
    assert.equal(xlsOpened.workbook.sheets[0].cells.B3.value, '00123')
    assert(xlsOpened.warnings.length > 0)

    const csvOpened = await workbookPayloadFromBytes('plain.csv', Buffer.from('name,amount\r\nalpha,42\r\nβeta,7', 'utf8'))
    assert.equal(csvOpened.workbook.sheets[0].cells.A2.value, 'alpha')
    assert.equal(csvOpened.workbook.sheets[0].cells.B2.value, 42)

    process.stdout.write(`Workbook QA passed: ${opened.stats.cells} cells, ${opened.stats.formulas} formulas, XLSX/XLS/ODS/CSV/TSV.\n`)
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
