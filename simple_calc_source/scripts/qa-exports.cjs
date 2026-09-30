'use strict'

const assert = require('node:assert/strict')
const { workbookPayloadFromBytes } = require('../electron/workbooks.cjs')
const {
  createSpreadsheetExport,
  exportFilter,
  normalizeExportFormat,
} = require('../electron/spreadsheet-export.cjs')

function request() {
  const sheets = [
    {
      id: 'sheet-1', name: 'Summary <2026>', state: 'visible', rowCount: 2, colCount: 2,
      cells: { A1: { value: '<Revenue>', style: { font: { bold: true } } }, B1: { value: 42 }, A2: { value: 'active only' } },
      merges: [], colWidths: {}, rowHeights: {}, hiddenRows: [], hiddenCols: [],
    },
    {
      id: 'sheet-2', name: 'Second', state: 'visible', rowCount: 1, colCount: 1,
      cells: { A1: { value: 'second sheet marker' } }, merges: [], colWidths: {}, rowHeights: {}, hiddenRows: [], hiddenCols: [],
    },
    {
      id: 'sheet-3', name: 'Hidden', state: 'hidden', rowCount: 1, colCount: 1,
      cells: { A1: { value: 'hidden marker' } }, merges: [], colWidths: {}, rowHeights: {}, hiddenRows: [], hiddenCols: [],
    },
  ]
  return {
    documentId: 'document-1',
    name: 'Plan.xlsx',
    workbook: { version: 1, name: 'Plan.xlsx', activeSheetId: 'sheet-1', sheets, metadata: {} },
    displayValues: {
      'sheet-1': { A1: '<Revenue>', B1: '$42.00', A2: 'active only' },
      'sheet-2': { A1: 'second sheet marker' },
      'sheet-3': { A1: 'hidden marker' },
    },
    selection: { top: 0, bottom: 0, left: 0, right: 1 },
    options: { scope: 'workbook', orientation: 'landscape', scaling: 'fit-width', paperSize: 'a4', gridlines: true, headings: true },
  }
}

async function main() {
  assert.equal(normalizeExportFormat('.PDF'), 'pdf')
  assert.deepEqual(exportFilter('html'), [{ name: 'Web page', extensions: ['html'] }])
  assert.throws(() => normalizeExportFormat('exe'), /choose xlsx/i)

  const html = await createSpreadsheetExport(request(), 'html')
  const htmlText = html.bytes.toString('utf8')
  assert.equal(html.format, 'html')
  assert.equal(html.printDocument.sheetCount, 2, 'HTML workbook export should include visible sheets')
  assert.equal(html.printDocument.pageCount, 2, 'the HTML export must use the same physical-page model as print')
  assert.match(htmlText, /<!doctype html>/i)
  assert.match(htmlText, /data-sheet-name="Second"/)
  assert.doesNotMatch(htmlText, /hidden marker/)
  assert.match(htmlText, /&lt;Revenue&gt;/, 'HTML values must be escaped')
  assert.match(htmlText, /@page \{ size: A4 landscape;/)
  assert.equal((htmlText.match(/class="print-page print-sheet"/g) || []).length, 2)

  const pdf = await createSpreadsheetExport(request(), 'pdf')
  assert.equal(pdf.bytes, null, 'Electron renders PDF bytes from the validated print document')
  assert.equal(pdf.printDocument.sheetCount, 2)
  assert.equal(pdf.printDocument.pageCount, html.printDocument.pageCount, 'PDF and HTML must share pagination')
  assert.equal(pdf.printDocument.html, html.printDocument.html, 'PDF, preview, and HTML export must share the validated print document')
  assert.equal(pdf.printDocument.options.scope, 'workbook')

  const csv = await createSpreadsheetExport(request(), 'csv')
  assert.match(csv.bytes.toString('utf8'), /active only/)
  assert.doesNotMatch(csv.bytes.toString('utf8'), /second sheet marker/, 'CSV must contain the active sheet only')

  const tsv = await createSpreadsheetExport(request(), 'tsv')
  assert.match(tsv.bytes.toString('utf8'), /\t/)
  assert.doesNotMatch(tsv.bytes.toString('utf8'), /second sheet marker/)

  const xlsx = await createSpreadsheetExport(request(), 'xlsx')
  assert.ok(Buffer.isBuffer(xlsx.bytes) && xlsx.bytes.length > 500)
  const reopened = await workbookPayloadFromBytes('export.xlsx', xlsx.bytes)
  assert.equal(reopened.workbook.sheets.length, 3)
  assert.equal(reopened.workbook.sheets[0].cells.A1.value, '<Revenue>')

  const ods = await createSpreadsheetExport(request(), 'ods')
  assert.ok(Buffer.isBuffer(ods.bytes) && ods.bytes.length > 500)
  const reopenedOds = await workbookPayloadFromBytes('export.ods', ods.bytes)
  assert.equal(reopenedOds.workbook.sheets.length, 3)

  await assert.rejects(() => createSpreadsheetExport({}, 'xlsx'), /invalid spreadsheet export request/i)
  console.log('Export QA passed: XLSX, ODS, CSV, TSV, PDF print input, HTML, escaping, visible-sheet scope, and format validation.')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
