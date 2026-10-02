'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const ExcelJS = require('exceljs')
const {
  workbookPayloadFromBytes,
  serializeWorkbook,
} = require('../electron/workbooks.cjs')

const PIXEL_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZlV8AAAAASUVORK5CYII='

async function createFidelityFixture() {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'simple_calc fidelity QA'
  workbook.lastModifiedBy = 'round-trip verifier'
  workbook.title = 'Workbook fidelity fixture'
  workbook.created = new Date('2026-01-02T03:04:05Z')
  workbook.modified = new Date('2026-02-03T04:05:06Z')
  workbook.calcProperties.fullCalcOnLoad = false
  workbook.calcProperties.forceFullCalc = false
  workbook.calcProperties.calcMode = 'auto'
  workbook.views = [{ activeTab: 0, firstSheet: 0, visibility: 'visible' }]

  const sheet = workbook.addWorksheet('Fidelity', {
    properties: {
      defaultRowHeight: 19,
      tabColor: { argb: 'FF476B57' },
      outlineProperties: { summaryBelow: false, summaryRight: false },
    },
    views: [{
      state: 'frozen',
      xSplit: 1,
      ySplit: 2,
      topLeftCell: 'B3',
      activeCell: 'B3',
      showGridLines: false,
      zoomScale: 125,
    }],
    pageSetup: {
      paperSize: 9,
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      printArea: 'A1:H20',
      printTitlesRow: '1:2',
      margins: { left: 0.3, right: 0.3, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
    },
    headerFooter: {
      oddHeader: '&Csimple_calc fidelity',
      oddFooter: '&LConfidential&RPage &P of &N',
    },
  })

  sheet.columns = [
    { width: 24 },
    { width: 14 },
    { width: 14 },
    { width: 16 },
    { width: 18 },
    { width: 12 },
    { width: 11, hidden: true },
    { width: 13 },
  ]
  sheet.getColumn(6).outlineLevel = 1
  sheet.getRow(1).height = 30
  sheet.getRow(5).hidden = true
  // ExcelJS only emits an otherwise-empty hidden row when it also has a
  // concrete row record such as a height.
  sheet.getRow(5).height = 19
  sheet.getRow(6).outlineLevel = 1

  sheet.mergeCells('A1:C1')
  sheet.getCell('A1').value = 'Fidelity test'
  sheet.getCell('A1').font = { name: 'Aptos Display', size: 16, bold: true, color: { argb: 'FFFFFFFF' } }
  sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF476B57' } }
  sheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
  sheet.getCell('A1').border = {
    bottom: { style: 'double', color: { argb: 'FF20362B' } },
    right: { style: 'mediumDashed', color: { theme: 1, tint: 0.2 } },
  }

  sheet.getCell('A2').value = 'Input'
  sheet.getCell('B2').value = 10
  sheet.getCell('B2').numFmt = '$#,##0.00;[Red]-$#,##0.00'
  sheet.getCell('B2').name = 'InputAmount'
  sheet.getCell('B2').note = {
    texts: [
      { text: 'Owner: ', font: { bold: true, color: { argb: 'FF476B57' } } },
      { text: 'Finance' },
    ],
  }
  sheet.getCell('B2').protection = { locked: false, hidden: false }
  sheet.getCell('C2').value = { formula: 'B2*2', result: 20 }
  sheet.getCell('C2').font = { italic: true, color: { theme: 4, tint: -0.2 } }
  sheet.getCell('C2').alignment = { horizontal: 'right', vertical: 'top', indent: 1 }
  sheet.getCell('D2').value = 'Open'
  sheet.getCell('D2').dataValidation = {
    type: 'list',
    allowBlank: false,
    showInputMessage: true,
    promptTitle: 'Status',
    prompt: 'Choose a status',
    formulae: ['"Open,Closed,Pending"'],
  }
  sheet.getCell('E2').value = {
    richText: [
      { font: { bold: true, color: { argb: 'FFE05A47' } }, text: 'Rich ' },
      { font: { italic: true, color: { theme: 4 } }, text: 'text' },
    ],
  }
  sheet.getCell('F2').value = { text: 'OpenAI', hyperlink: 'https://openai.com', tooltip: 'Open link' }

  sheet.getCell('A3').value = 'Wrapped and rotated text'
  sheet.getCell('A3').alignment = { wrapText: true, vertical: 'middle', textRotation: 25, shrinkToFit: true }
  sheet.getRow(3).height = 42
  sheet.getCell('B3').value = 0.25
  sheet.getCell('B3').numFmt = '0.00%'
  sheet.getCell('C3').value = { formula: 'SUM(B2:B3)', result: 10.25 }
  sheet.getCell('D3').value = 'Pattern'
  sheet.getCell('D3').fill = {
    type: 'pattern',
    pattern: 'darkTrellis',
    fgColor: { indexed: 10 },
    bgColor: { argb: 'FFFFFFCC' },
  }
  sheet.getCell('E3').value = 'Gradient'
  sheet.getCell('E3').fill = {
    type: 'gradient',
    gradient: 'angle',
    degree: 45,
    stops: [
      { position: 0, color: { argb: 'FFFFFFFF' } },
      { position: 1, color: { argb: 'FF4F81BD' } },
    ],
  }
  sheet.getCell('F3').value = 'Borders'
  sheet.getCell('F3').border = {
    top: { style: 'thin', color: { argb: 'FFFF0000' } },
    left: { style: 'dashed', color: { argb: 'FF00AA00' } },
    bottom: { style: 'double', color: { argb: 'FF0000FF' } },
    right: { style: 'thick', color: { indexed: 8 } },
    diagonal: { style: 'thin', color: { theme: 5 } },
    diagonalUp: true,
  }
  sheet.getCell('H20').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } }

  sheet.getCell('C6').value = {
    formula: 'TRANSPOSE(B2:B4)',
    result: 10,
    shareType: 'array',
    ref: 'C6:E6',
  }

  sheet.addTable({
    name: 'SalesTable',
    ref: 'A10',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium4', showRowStripes: true },
    columns: [
      { name: 'Item', totalsRowLabel: 'Total' },
      { name: 'Amount', totalsRowFunction: 'sum' },
    ],
    rows: [['Alpha', 5], ['Beta', 12], ['Gamma', 8]],
  })
  sheet.addConditionalFormatting({
    ref: 'B11:B13',
    rules: [{
      type: 'cellIs',
      operator: 'greaterThan',
      formulae: [9],
      style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC6EFCE' } } },
    }],
  })
  sheet.autoFilter = 'A10:B13'

  const imageId = workbook.addImage({ base64: PIXEL_PNG, extension: 'png' })
  sheet.addImage(imageId, { tl: { col: 5, row: 4 }, ext: { width: 32, height: 32 }, editAs: 'oneCell' })

  const hidden = workbook.addWorksheet('Hidden formulas')
  hidden.state = 'veryHidden'
  hidden.getCell('A1').value = { formula: 'Fidelity!B2+1', result: 11 }

  return Buffer.from(await workbook.xlsx.writeBuffer({ useStyles: true, useSharedStrings: true }))
}

async function loadExcelJS(bytes) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(bytes)
  return workbook
}

const RED_PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==', 'base64')

/** A source sheet loaded with everything a sheet can carry besides cells. */
async function decoratedSheetSource() {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Summary').getCell('A1').value = 'summary'
  const old = workbook.addWorksheet('Old', { properties: { tabColor: { argb: 'FFFF0000' } } })
  old.getCell('A1').value = 'old data'
  old.getCell('A2').value = 'x'
  old.headerFooter.oddHeader = '&CCONFIDENTIAL - Old payroll'
  old.dataValidations.add('B2:B20', { type: 'list', allowBlank: true, formulae: ['"Yes,No"'] })
  old.addConditionalFormatting({ ref: 'A1:C20', rules: [{ type: 'expression', formulae: ['TRUE'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } }] })
  old.addImage(workbook.addImage({ buffer: RED_PNG, extension: 'png' }), 'D2:F6')
  old.addTable({ name: 'Salaries', ref: 'H1', headerRow: true, columns: [{ name: 'Name' }, { name: 'Salary' }], rows: [['Ann', 9000]] })
  old.autoFilter = 'A1:C5'
  await old.protect('secret', {})
  return Buffer.from(await workbook.xlsx.writeBuffer({ useStyles: true, useSharedStrings: true }))
}

function assertBareSheet(sheet, label) {
  assert(sheet, `${label} exists`)
  assert.equal(sheet.properties.tabColor, undefined, `${label}: no inherited tab colour`)
  assert.ok(!sheet.headerFooter || !sheet.headerFooter.oddHeader, `${label}: no inherited header`)
  assert.ok(!sheet.sheetProtection || !sheet.sheetProtection.sheet, `${label}: not protected`)
  assert.equal(Object.keys(sheet.dataValidations.model).length, 0, `${label}: no inherited dropdowns`)
  assert.equal(sheet.conditionalFormattings.length, 0, `${label}: no inherited conditional formats`)
  assert.equal(sheet.getImages().length, 0, `${label}: no inherited pictures`)
  assert.equal(sheet.getTables().length, 0, `${label}: no inherited tables`)
  assert.ok(!sheet.autoFilter, `${label}: no inherited filter`)
}

/**
 * calc-file-io-objects-2 / CALC-SIE-11: a sheet added after deleting another one never
 * inherits the deleted sheet's objects, and settings the model removed do not come back.
 */
async function sheetIdentity() {
  const source = await decoratedSheetSource()
  const opened = (await workbookPayloadFromBytes('identity.xlsx', source)).workbook
  // (A) delete "Old", add a fresh "Sheet2" (the shape App.addSheet creates); (A2) named "Old".
  for (const name of ['Sheet2', 'Old']) {
    const model = structuredClone(opened)
    model.sheets.splice(1, 1)
    model.sheets.push({ id: `new-${name}`, name, state: 'visible', rowCount: 100, colCount: 26, cells: { A1: { value: 'my new data' } }, merges: [], colWidths: {}, rowHeights: {}, frozen: {} })
    const saved = await loadExcelJS(await serializeWorkbook(model, 'xlsx', { baseBytes: source }))
    const sheet = saved.getWorksheet(name)
    assertBareSheet(sheet, `new sheet "${name}"`)
    assert.equal(sheet.getCell('A1').value, 'my new data')
    assert.equal(saved.media.length, 0, 'the deleted sheet\'s picture is not kept in the package')
  }
  // (C) settings removed in the editor (keys dropped from the model) do not come back.
  const cleared = structuredClone(opened)
  const old = cleared.sheets[1]
  for (const key of ['autoFilter', 'sheetProtection', 'dataValidations', 'conditionalFormattings', 'tables', 'images']) delete old[key]
  old.properties = { ...old.properties }
  delete old.properties.tabColor
  old.headerFooter = {}
  const clearedBook = await loadExcelJS(await serializeWorkbook(cleared, 'xlsx', { baseBytes: source }))
  assertBareSheet(clearedBook.getWorksheet('Old'), 'cleared "Old"')
  // An untouched matched sheet keeps everything.
  const kept = await loadExcelJS(await serializeWorkbook(opened, 'xlsx', { baseBytes: source }))
  const keptOld = kept.getWorksheet('Old')
  assert.equal(keptOld.properties.tabColor.argb, 'FFFF0000')
  assert.ok(keptOld.sheetProtection.sheet)
  assert.equal(keptOld.getImages().length, 1)
  assert.equal(keptOld.getTables().length, 1)
  assert.ok(Object.keys(keptOld.dataValidations.model).length > 0)
  // (Dup) a duplicate moved before its original: the original keeps the source worksheet
  // (here: pictures the editor did not load), the copy gets a fresh one.
  const duplicated = structuredClone(opened)
  const original = duplicated.sheets[1]
  delete original.images
  original.requiresSourcePackage = ['images']
  delete original.tables
  original.tables = []
  const copy = { ...structuredClone(original), id: 'copy', name: 'Old (2)' }
  duplicated.sheets.splice(1, 0, copy)
  const duplicateBook = await loadExcelJS(await serializeWorkbook(duplicated, 'xlsx', { baseBytes: source }))
  assert.equal(duplicateBook.getWorksheet('Old').getImages().length, 1, 'the original keeps its source pictures')
  assert.equal(duplicateBook.getWorksheet('Old (2)').getImages().length, 0, 'the copy does not take the original\'s source worksheet')
  assert.deepEqual(duplicateBook.worksheets.map((sheet) => sheet.name), ['Summary', 'Old (2)', 'Old'])
}

async function main() {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-calc-fidelity-'))
  try {
    const sourceBytes = await createFidelityFixture()
    const sourcePath = path.join(directory, 'fidelity.xlsx')
    await fs.writeFile(sourcePath, sourceBytes)

    const opened = await workbookPayloadFromBytes('fidelity.xlsx', sourceBytes)
    assert.equal(opened.sourceFormat, 'xlsx')
    assert.equal(opened.workbook.metadata.importedWith, 'exceljs')
    const modelSheet = opened.workbook.sheets.find((item) => item.name === 'Fidelity')
    assert(modelSheet)
    assert.equal(modelSheet.cells.C2.formula, 'B2*2')
    assert.equal(modelSheet.cells.C2.result, 20)
    assert.equal(modelSheet.cells.C6.formulaType, 'array')
    assert.equal(modelSheet.cells.C6.formulaRange, 'C6:E6')
    assert.equal(modelSheet.cells.A1.style.border.bottom.style, 'double')
    assert.equal(modelSheet.cells.C2.style.font.color.theme, 4)
    assert.equal(modelSheet.cells.D3.style.fill.fgColor.indexed, 10)
    assert.equal(modelSheet.cells.E3.style.fill.type, 'gradient')
    assert.equal(modelSheet.cells.A3.style.alignment.textRotation, 25)
    assert(modelSheet.hiddenRows.includes(5))
    assert(modelSheet.hiddenCols.includes(7))
    assert.deepEqual(modelSheet.frozen && { rows: modelSheet.frozen.rows, columns: modelSheet.frozen.columns }, { rows: 2, columns: 1 })
    assert.equal(modelSheet.rowHeights['3'], 42)
    assert.equal(modelSheet.colWidths['1'], 24)
    assert(modelSheet.cells.H20.style.fill)

    // The same OOXML reader must be used for macro/template package families,
    // so styles and formulas are not downgraded merely because of extension.
    const macroNamed = await workbookPayloadFromBytes('fidelity.xlsm', sourceBytes)
    assert.equal(macroNamed.workbook.metadata.importedWith, 'exceljs')
    assert.equal(macroNamed.workbook.sheets[0].cells.C2.formula, 'B2*2')
    assert.equal(macroNamed.workbook.sheets[0].cells.A1.style.fill.fgColor.argb, 'FF476B57')

    modelSheet.cells.B2.value = 15
    modelSheet.cells.B2.display = '$15.00'
    const overlaidBytes = await serializeWorkbook(opened.workbook, 'xlsx', { baseBytes: sourceBytes, sourceFormat: 'xlsx' })
    assert(Buffer.isBuffer(overlaidBytes) && overlaidBytes.length > 5_000)
    const overlaid = await loadExcelJS(overlaidBytes)
    const savedSheet = overlaid.getWorksheet('Fidelity')
    assert(savedSheet)
    assert.equal(savedSheet.getCell('B2').value, 15)
    assert.equal(savedSheet.getCell('B2').numFmt, '$#,##0.00;[Red]-$#,##0.00')
    assert.equal(savedSheet.getCell('C2').formula, 'B2*2')
    assert.equal(savedSheet.getCell('A1').fill.fgColor.argb, 'FF476B57')
    assert.equal(savedSheet.getCell('F3').border.bottom.style, 'double')
    assert.equal(savedSheet.getCell('D2').dataValidation.type, 'list')
    assert(savedSheet.conditionalFormattings.length > 0)
    assert(savedSheet.getTable('SalesTable'))
    assert(savedSheet.getImages().length > 0)
    assert.equal(savedSheet.pageSetup.orientation, 'landscape')
    assert.equal(savedSheet.pageSetup.printArea, 'A1:H20')
    assert.equal(savedSheet.headerFooter.oddHeader, '&Csimple_calc fidelity')
    assert.equal(savedSheet.views[0].state, 'frozen')
    assert.equal(savedSheet.views[0].xSplit, 1)
    assert.equal(savedSheet.views[0].ySplit, 2)
    assert.equal(savedSheet.getRow(5).hidden, true)
    assert.equal(savedSheet.getColumn(7).hidden, true)
    assert.equal(overlaid.getWorksheet('Hidden formulas').state, 'veryHidden')

    const reopened = await workbookPayloadFromBytes('overlaid.xlsx', overlaidBytes)
    const reopenedSheet = reopened.workbook.sheets.find((item) => item.name === 'Fidelity')
    assert.equal(reopenedSheet.cells.B2.value, 15)
    assert.equal(reopenedSheet.cells.C2.formula, 'B2*2')
    assert.equal(reopenedSheet.cells.A1.style.fill.fgColor.argb, 'FF476B57')
    // Pictures are now modelled (shown and editable), so they need no compatibility warning.
    assert(!reopened.warnings.some((warning) => /images/i.test(warning)))
    assert(reopenedSheet.images && reopenedSheet.images.length > 0, 'pictures are loaded into the editor')

    await sheetIdentity()

    process.stdout.write('Fidelity QA passed: source-backed XLSX overlay retained styles, formulas, structure, validations, conditional formatting, tables, images, and print/view settings; new sheets never inherit a deleted sheet\'s objects and removed settings stay removed.\n')
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
