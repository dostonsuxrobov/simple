'use strict'

const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')
const ExcelJS = require('exceljs')
const XLSX = require('xlsx')
const { extractLegacyBiffStyles } = require('../electron/legacy-biff-styles.cjs')
const { serializeWorkbook, workbookPayloadFromPath } = require('../electron/workbooks.cjs')
const { findOfficeConverter } = require('../electron/office-converter.cjs')

const FIXTURE_PATH = path.join(__dirname, 'fixtures', 'schedule_template.xls')
const FIXTURE_SHA256 = '0b7c348b35dc5624ae134c09043708b85b980648c89bfda587522f4f9c0c8787'
const ACCOUNTING_INTEGER = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
const ACCOUNTING_CURRENCY = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'

function approximately(actual, expected, label) {
  assert.equal(typeof actual, 'number', `${label} must be numeric`)
  assert.ok(Math.abs(actual - expected) < 1e-9, `${label}: expected ${expected}, received ${actual}`)
}

function borderStyles(cell) {
  const border = cell && cell.style && cell.style.border
  return Object.fromEntries(['top', 'right', 'bottom', 'left'].map((side) => [side, border && border[side] && border[side].style]))
}

function assertBorder(sheet, address, style) {
  assert.deepEqual(borderStyles(sheet.cells[address]), {
    top: style,
    right: style,
    bottom: style,
    left: style,
  }, `${address} should retain its four ${style} BIFF borders`)
}

async function main() {
  const fixtureBytes = fs.readFileSync(FIXTURE_PATH)
  assert.equal(
    crypto.createHash('sha256').update(fixtureBytes).digest('hex'),
    FIXTURE_SHA256,
    'The legacy XLS regression fixture changed unexpectedly.',
  )

  const parsedFixture = XLSX.read(fixtureBytes, { type: 'buffer', cellNF: true })
  const explicitAccounting = '0.0000 "per-workbook"'
  const explicitFormats = { ...parsedFixture.SSF, 44: explicitAccounting }
  const directlyExtracted = extractLegacyBiffStyles(fixtureBytes, explicitFormats)
  assert.ok(directlyExtracted, 'The supplied workbook should be recognized as BIFF8')
  assert.equal(
    directlyExtracted.sheets[0].cells.E2.numFmt,
    explicitAccounting,
    'The BIFF extractor must use its explicit per-workbook format table rather than global SSF state.',
  )

  const olderBiffCfb = XLSX.CFB.read(fixtureBytes, { type: 'buffer' })
  const workbookEntry = olderBiffCfb.FileIndex.find((entry) => (
    entry && entry.type === 2 && /^(?:Workbook|Book)$/i.test(String(entry.name || '')) && entry.content
  ))
  assert.ok(workbookEntry, 'The regression fixture should contain a BIFF Workbook stream')
  workbookEntry.content = Buffer.from(workbookEntry.content)
  assert.equal(workbookEntry.content.readUInt16LE(0), 0x0809, 'The Workbook stream should begin with a BIFF8 BOF record')
  assert.equal(workbookEntry.content.readUInt16LE(4), 0x0600, 'The fixture should declare BIFF8')
  workbookEntry.content.writeUInt16LE(0x0500, 4)
  workbookEntry.size = workbookEntry.content.length
  const olderBiffBytes = Buffer.from(XLSX.CFB.write(olderBiffCfb, { type: 'buffer' }))
  assert.equal(
    extractLegacyBiffStyles(olderBiffBytes, parsedFixture.SSF),
    null,
    'The BIFF8-only style parser must reject older BIFF layouts instead of decoding them incorrectly.',
  )

  const opened = await workbookPayloadFromPath(FIXTURE_PATH)
  assert.equal(opened.sourceFormat, 'xls')
  // Without the document engine an edited .xls is saved as an .xlsx next to the original.
  const officeEngine = Boolean(await findOfficeConverter())
  assert.equal(opened.requiresSaveAs, !officeEngine)
  assert.equal(opened.workbook.sheets.length, 1)
  assert.equal(opened.stats.formulas, 2)

  const sheet = opened.workbook.sheets[0]
  assert.equal(sheet.name, 'Sheet1')
  assert.equal(sheet.rowCount, 38)
  assert.equal(sheet.colCount, 6)
  assert.deepEqual(sheet.hiddenCols, [], 'The BIFF-limit hidden tail should not collapse the editable columns after F')

  assert.deepEqual([...sheet.merges].sort(), [
    'A32:E32',
    'A35:B35',
    'A37:B37',
    'D37:E37',
  ])

  approximately(sheet.colWidths['1'], 34.6640625, 'column A width')
  approximately(sheet.colWidths['2'], 13.83203125, 'column B width')
  approximately(sheet.colWidths['3'], 13.5, 'column C width')
  approximately(sheet.colWidths['4'], 21.5, 'column D width')
  approximately(sheet.colWidths['5'], 15.33203125, 'column E width')
  approximately(sheet.colWidths['6'], 3, 'column F width')
  approximately(sheet.rowHeights['1'], 15.75, 'row 1 height')
  approximately(sheet.rowHeights['2'], 16, 'automatic row 2 height')
  approximately(sheet.rowHeights['26'], 16, 'automatic row 26 height')
  approximately(sheet.rowHeights['34'], 16, 'automatic row 34 height')
  approximately(sheet.rowHeights['30'], 15.5, 'row 30 height')
  approximately(sheet.rowHeights['32'], 77.25, 'row 32 height')
  approximately(sheet.rowHeights['37'], 18, 'row 37 height')
  assert.equal(opened.workbook.metadata.normalFont.name, 'Arial')
  assert.equal(opened.workbook.metadata.normalFont.size, 10)
  assert.equal(sheet.pageSetup.printArea, 'A1:E37')
  assert.equal(sheet.pageSetup.scale, 91)
  assert.equal(sheet.pageSetup.fitToPage, false, 'stored fit1x1 must not activate disabled fit-to-page mode')
  assert.equal(sheet.pageSetup.orientation, 'portrait')
  assert.match(sheet.headerFooter.oddHeader, /SCHEDULE OF ACCOUNTS/)
  assert.match(sheet.headerFooter.oddHeader, /Page &P of &N/)

  // Blank BIFF cells carry the table's layout. They must not be discarded or
  // reduced to their white fill by the compatibility reader.
  for (const address of ['A2', 'B2', 'C2', 'D2', 'E2', 'A26', 'E26']) {
    assert.ok(sheet.cells[address], `${address} should remain as a styled blank cell`)
    assertBorder(sheet, address, 'thin')
  }
  for (const address of ['A1', 'B1', 'C1', 'D1', 'E1']) assertBorder(sheet, address, 'thin')
  assertBorder(sheet, 'E30', 'medium')
  assert.deepEqual(borderStyles(sheet.cells.A27), {
    top: undefined,
    right: undefined,
    bottom: undefined,
    left: undefined,
  }, 'The separator below the invoice table should remain borderless')

  assert.equal(sheet.cells.A1.style.font.name, 'Times New Roman')
  assert.equal(sheet.cells.A1.style.font.size, 10)
  assert.equal(sheet.cells.A2.style.font.name, 'Arial')
  assert.equal(sheet.cells.A2.style.font.size, 12)
  assert.equal(sheet.cells.A32.style.font.name, 'Times New Roman')
  assert.equal(sheet.cells.A32.style.font.size, 10)

  assert.equal(sheet.cells.A1.style.alignment.horizontal, 'center')
  assert.equal(sheet.cells.A1.style.alignment.vertical, 'bottom')
  assert.equal(sheet.cells.A1.style.alignment.wrapText, true)
  assert.equal(sheet.cells.A2.style.alignment.horizontal, 'left')
  assert.equal(sheet.cells.A2.style.alignment.vertical, 'bottom')
  assert.equal(sheet.cells.B2.style.alignment.horizontal, 'center')
  assert.equal(sheet.cells.B2.style.alignment.vertical, 'bottom')
  assert.equal(sheet.cells.A32.style.alignment.wrapText, true)

  assert.equal(sheet.cells.B1.numFmt, '@')
  assert.equal(sheet.cells.C1.numFmt, 'm/d/yy;@')
  assert.equal(sheet.cells.D1.numFmt, '@')
  assert.equal(sheet.cells.B2.numFmt, '@')
  assert.equal(sheet.cells.C2.numFmt, 'm/d/yy;@')
  assert.equal(sheet.cells.D2.numFmt, '@')
  assert.equal(sheet.cells.E2.numFmt, ACCOUNTING_CURRENCY)
  assert.equal(sheet.cells.E26.numFmt, ACCOUNTING_CURRENCY)
  assert.equal(sheet.cells.A30.numFmt, ACCOUNTING_INTEGER)
  assert.equal(sheet.cells.E30.numFmt, ACCOUNTING_CURRENCY)

  assert.equal(sheet.cells.A30.formula, 'COUNT(E2:E26)')
  assert.equal(sheet.cells.A30.result, 0)
  assert.equal(sheet.cells.A30.display.replace(/\s/g, ''), '-')
  assert.equal(sheet.cells.E30.formula, 'SUM(E2:E26)')
  assert.equal(sheet.cells.E30.result, 0)
  assert.equal(sheet.cells.E30.display.replace(/\s/g, ''), '$-')

  const savedBytes = await serializeWorkbook(opened.workbook, 'xlsx')
  const savedWorkbook = new ExcelJS.Workbook()
  await savedWorkbook.xlsx.load(savedBytes)
  const savedSheet = savedWorkbook.worksheets[0]
  approximately(savedSheet.getRow(2).height, 16, 'saved automatic row height')
  assert.equal(savedSheet.pageSetup.scale, 91)
  assert.equal(savedSheet.pageSetup.fitToPage, false)
  assert.equal(savedSheet.headerFooter.oddHeader, sheet.headerFooter.oddHeader)
  for (const [range, leftAddress, rightAddress] of [
    ['A35:B35', 'A35', 'B35'],
    ['A37:B37', 'A37', 'B37'],
    ['D37:E37', 'D37', 'E37'],
  ]) {
    const leftBorder = savedSheet.getCell(leftAddress).border
    const rightBorder = savedSheet.getCell(rightAddress).border
    assert.equal(leftBorder.left && leftBorder.left.style, 'thin', `${range} should retain its left edge after Save As XLSX`)
    assert.equal(leftBorder.top && leftBorder.top.style, 'thin', `${range} should retain its top edge after Save As XLSX`)
    assert.equal(leftBorder.bottom && leftBorder.bottom.style, 'thin', `${range} should retain its bottom edge after Save As XLSX`)
    assert.equal(rightBorder.right && rightBorder.right.style, 'thin', `${range} should retain its right edge after Save As XLSX`)
  }

  process.stdout.write('Legacy XLS QA passed: BIFF styles, formats, formulas, merges, and dimensions retained.\n')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
