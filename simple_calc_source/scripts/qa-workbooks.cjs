'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const ExcelJS = require('exceljs')
const JSZip = require('jszip')
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

    // Native readers require absolute rows as well as columns in Print_Area.
    // Also verify layout edits replace the imported built-in instead of reviving it.
    const printModel = structuredClone(opened.workbook)
    printModel.sheets[0].name = "Owner's summary"
    printModel.sheets[0].pageSetup = { printArea: 'A1:D4&&A8:B9', printTitlesRow: '1:2', printTitlesColumn: 'A:B' }
    printModel.definedNames = [{ name: '_xlnm.Print_Area', ranges: ["'Summary'!$A$1:$Z$99"], localSheetIndex: 0 }]
    const printBytes = await serializeWorkbook(printModel, 'xlsx')
    const printZip = await JSZip.loadAsync(printBytes)
    const printXml = await printZip.file('xl/workbook.xml').async('string')
    assert.match(printXml, /Owner&apos;s|Owner&apos;&apos;s|Owner''s/)
    assert.ok(printXml.includes('$A$1:$D$4') && printXml.includes('$A$8:$B$9'))
    assert.ok(!printXml.includes('$Z$99'), 'stale imported print area must not override the saved layout')
    assert.ok(printXml.includes('$1:$2') && printXml.includes('$A:$B'))
    const printReopened = await workbookPayloadFromBytes('print.xlsx', printBytes)
    assert.equal(printReopened.workbook.sheets[0].pageSetup.printArea, 'A1:D4&&A8:B9')

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
    assert.equal(xlsOpened.requiresSaveAs, false)
    // SheetJS' legacy BIFF writer does not emit formula records; real XLS
    // formula imports are covered by the decoder, while this generated fixture
    // still exercises values, text types, merges, and XLS-to-XLSX conversion.
    assert.equal(xlsOpened.workbook.sheets[0].cells.C2.value, 25)
    assert.equal(xlsOpened.workbook.sheets[0].cells.B3.value, '00123')
    assert(xlsOpened.warnings.length > 0)

    const csvOpened = await workbookPayloadFromBytes('plain.csv', Buffer.from('name,amount\r\nalpha,42\r\nβeta,7', 'utf8'))
    assert.equal(csvOpened.workbook.sheets[0].cells.A2.value, 'alpha')
    assert.equal(csvOpened.workbook.sheets[0].cells.B2.value, 42)

    const euCsv = await workbookPayloadFromBytes('euro.csv', Buffer.from('name;value;price\r\nwidget;1,5;1.234,56\r\ngadget;2;00123', 'utf8'))
    const euSheet = euCsv.workbook.sheets[0]
    assert.equal(euCsv.workbook.metadata.delimiter, ';')
    assert.equal(euSheet.cells.C1.value, 'price')
    assert.equal(euSheet.cells.B2.value, 1.5)
    assert.equal(euSheet.cells.B2.display, '1,5')
    assert.equal(euSheet.cells.C2.value, 1234.56)
    assert.equal(euSheet.cells.B3.value, 2)
    assert.equal(euSheet.cells.C3.value, '00123')
    assert.equal(euSheet.cells.C3.numFmt, '@')
    const euTsv = await serializeWorkbook(euCsv.workbook, 'tsv')
    assert.match(euTsv.toString('utf8'), /widget\t1\.5\t1,234\.56/)

    const usCsv = await workbookPayloadFromBytes('us.csv', Buffer.from('label,qty,price,pct\r\nwidget,42,"1,234.5",20%', 'utf8'))
    const usSheet = usCsv.workbook.sheets[0]
    assert.equal(usCsv.workbook.metadata.delimiter, ',')
    assert.equal(usSheet.cells.B2.value, 42)
    assert.equal(usSheet.cells.C2.value, 1234.5)
    assert.equal(usSheet.cells.C2.display, '1,234.5')
    assert.equal(usSheet.cells.D2.value, 0.2)

    // Quoted "n,n" text pairs must not flip a comma-delimited US file to
    // decimal-comma parsing (1.234 stays 1.234 and "6,8" stays text), while
    // unambiguous EU grouped values still enable it.
    const pairsCsv = await workbookPayloadFromBytes('pairs.csv', Buffer.from('sku,sizes,weight\r\nA1,"6,8",1.234\r\nB2,"8,10",2.5\r\nC3,"10,12",0.75', 'utf8'))
    const pairsSheet = pairsCsv.workbook.sheets[0]
    assert.equal(pairsSheet.cells.C2.value, 1.234)
    assert.equal(pairsSheet.cells.B2.value, '6,8')
    const euGroupedCsv = await workbookPayloadFromBytes('grouped.csv', Buffer.from('a,b\r\nx,"1.234,56"', 'utf8'))
    assert.equal(euGroupedCsv.workbook.sheets[0].cells.B2.value, 1234.56)

    const pipeTxt = await workbookPayloadFromBytes('pipes.txt', Buffer.from('a|b|c\r\n1|2|3\r\n4|5|6', 'utf8'))
    assert.equal(pipeTxt.workbook.metadata.delimiter, '|')
    assert.equal(pipeTxt.workbook.sheets[0].cells.C2.value, 3)

    const isoSerial = (Date.UTC(2024, 0, 15) - Date.UTC(1899, 11, 30)) / 86_400_000
    const idsCsv = await workbookPayloadFromBytes('ids.csv', Buffer.from('id,when\r\n00123,2024-01-15\r\n42,1/15/2024', 'utf8'))
    const idsSheet = idsCsv.workbook.sheets[0]
    assert.equal(idsSheet.cells.A2.value, '00123')
    assert.equal(idsSheet.cells.A2.numFmt, '@')
    assert.equal(idsSheet.cells.B2.value, isoSerial)
    assert.equal(idsSheet.cells.B2.numFmt, 'yyyy-mm-dd')
    assert.equal(idsSheet.cells.A3.value, 42)
    assert.equal(idsSheet.cells.B3.value, isoSerial)
    const idsExported = await serializeWorkbook(idsCsv.workbook, 'csv')
    assert.match(idsExported.toString('utf8'), /\r\n00123,2024-01-15\r\n/)
    const idsReopened = await workbookPayloadFromBytes('ids-roundtrip.csv', idsExported)
    assert.equal(idsReopened.workbook.sheets[0].cells.A2.value, '00123')
    assert.equal(idsReopened.workbook.sheets[0].cells.B2.value, isoSerial)

    // Excel's reserved builtin number formats (Currency / Comma / Accounting), the locale
    // short-date builtins, rich-text runs, phantom fills and colour-only borders all reach
    // the model the way the renderer expects them.
    const stylesPath = path.join(directory, 'styles.xlsx')
    const stylesBook = new ExcelJS.Workbook()
    const styled = stylesBook.addWorksheet('Styled')
    styled.getCell('A1').value = 1234.5
    styled.getCell('A1').numFmt = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    styled.getCell('A2').value = new Date(Date.UTC(2023, 2, 15))
    styled.getCell('A3').value = { richText: [{ text: 'plain ' }, { text: 'bold', font: { bold: true } }] }
    styled.getCell('A4').value = 'bordered'
    styled.getCell('A4').border = { right: { style: 'none', color: { argb: 'FF000000' } } }
    styled.getCell('A5').value = 'unfilled'
    styled.getCell('A5').fill = { type: 'pattern', pattern: 'none' }
    await stylesBook.xlsx.writeFile(stylesPath)
    const stylesOpened = await workbookPayloadFromPath(stylesPath)
    const styledSheet = stylesOpened.workbook.sheets[0]
    assert.equal(styledSheet.cells.A2.numFmt, 'm/d/yyyy', 'builtin short date resolves to the locale format')
    assert.equal(typeof styledSheet.cells.A3.value, 'string', 'rich text keeps a scalar value')
    assert.equal(styledSheet.cells.A3.value, 'plain bold')
    assert.equal(styledSheet.cells.A3.richText.length, 2, 'rich-text runs travel beside the plain text')
    // Import keeps style="none" verbatim for round-trip fidelity; the renderer is what
    // refuses to draw it (see drawableBorderSide in src/App.tsx).
    assert.equal(styledSheet.cells.A4.style.border.right.style, 'none')
    assert.equal(styledSheet.cells.A5.style && styledSheet.cells.A5.style.fill, undefined, 'a paint-nothing fill is dropped')
    const styledRoundTrip = await workbookPayloadFromBytes(
      'styles-roundtrip.xlsx',
      await serializeWorkbook(stylesOpened.workbook, 'xlsx'),
    )
    assert.equal(styledRoundTrip.workbook.sheets[0].cells.A3.richText.length, 2, 'untouched rich text survives a save')

    // ExcelJS omits the reserved ids Excel uses for its Currency/Comma/Accounting styles,
    // and those ids never appear in <numFmts>, so they have to be seeded from a builtin table.
    const reservedPath = path.join(directory, 'reserved-formats.xlsx')
    const reservedZip = await JSZip.loadAsync(await fs.readFile(stylesPath))
    const reservedStyles = await reservedZip.file('xl/styles.xml').async('string')
    const cellXfsMatch = /<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/.exec(reservedStyles)
    assert(cellXfsMatch, 'the fixture has a cellXfs table')
    const accountingXfIndex = Number(cellXfsMatch[1])
    reservedZip.file('xl/styles.xml', reservedStyles.replace(
      cellXfsMatch[0],
      `<cellXfs count="${accountingXfIndex + 1}">${cellXfsMatch[2]}` +
        '<xf numFmtId="44" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>',
    ))
    const reservedSheet = await reservedZip.file('xl/worksheets/sheet1.xml').async('string')
    reservedZip.file('xl/worksheets/sheet1.xml', reservedSheet.replace(
      '</sheetData>',
      `<row r="20"><c r="A20" s="${accountingXfIndex}"><v>1234.5</v></c></row></sheetData>`,
    ))
    await fs.writeFile(reservedPath, await reservedZip.generateAsync({ type: 'nodebuffer' }))
    const reservedOpened = await workbookPayloadFromPath(reservedPath)
    assert.match(
      String(reservedOpened.workbook.sheets[0].cells.A20.numFmt),
      /^_\("\$"\* #,##0\.00_\)/,
      'builtin accounting format id 44 resolves',
    )

    // A formula with no cached result must survive an ODS export.
    const odsSource = {
      version: 1,
      name: 'ods-formulas.xlsx',
      activeSheetId: 's1',
      sheets: [{
        id: 's1',
        name: 'Sheet1',
        rowCount: 4,
        colCount: 3,
        cells: { A1: { value: 1 }, B1: { value: 2 }, C1: { formula: 'A1+B1' } },
        merges: [],
      }],
      definedNames: [],
      metadata: {},
    }
    const odsBytes = await serializeWorkbook(odsSource, 'ods')
    const odsReopened = await workbookPayloadFromBytes('ods-formulas.ods', odsBytes)
    assert.equal(odsReopened.workbook.sheets[0].cells.C1.formula, 'A1+B1', 'an uncached formula survives an ODS export')

    const encrypted = Buffer.concat([
      Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
      Buffer.alloc(504),
      Buffer.from('EncryptedPackage', 'utf16le'),
    ])
    for (const name of ['secret.xlsx', 'secret.xls', 'secret.xlsb']) {
      await assert.rejects(workbookPayloadFromBytes(name, encrypted), /password-protected/)
    }

    process.stdout.write(`Workbook QA passed: ${opened.stats.cells} cells, ${opened.stats.formulas} formulas, XLSX/XLS/ODS/CSV/TSV.\n`)
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
