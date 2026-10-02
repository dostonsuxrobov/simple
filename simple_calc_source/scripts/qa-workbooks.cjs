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
  importWorkbookBytes,
  serializeWorkbook,
} = require('../electron/workbooks.cjs')
const { findOfficeConverter } = require('../electron/office-converter.cjs')

const SAMPLES = path.resolve(__dirname, '..', '..', 'Simple test examples')
const notes = []

async function sheetXml(bytes, part = 'xl/worksheets/sheet1.xml') {
  return (await JSZip.loadAsync(bytes)).file(part).async('string')
}

/** Rewrite one part of an OOXML package (used to build fixtures ExcelJS cannot write). */
async function patchPart(bytes, part, change) {
  const zip = await JSZip.loadAsync(bytes)
  zip.file(part, change(await zip.file(part).async('string')))
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function sharedFormulas() {
  // CALC-SIE-1: editing, clearing or retyping the first cell of a filled-down (shared)
  // formula must never block the save or rewrite the other cells.
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Data')
  for (let row = 2; row <= 6; row += 1) sheet.getCell(`A${row}`).value = row - 1
  sheet.getCell('B2').value = { formula: 'A2*2', result: 2, shareType: 'shared', ref: 'B2:B6' }
  for (let row = 3; row <= 6; row += 1) sheet.getCell(`B${row}`).value = { sharedFormula: 'B2', result: (row - 1) * 2 }
  const base = Buffer.from(await workbook.xlsx.writeBuffer())
  assert.match(await sheetXml(base), /t="shared"/, 'the fixture uses a shared formula')
  const imported = (await workbookPayloadFromBytes('shared.xlsx', base)).workbook
  const cells = imported.sheets[0].cells
  assert.equal(cells.B3.formula, 'A3*2', 'each cell owns its translated formula')
  assert.equal(cells.B6.formula, 'A6*2')
  for (const cell of Object.values(cells)) {
    assert.notEqual(cell.formulaType, 'shared')
    assert.equal(cell.sharedFormulaMaster, undefined)
  }
  const edits = {
    formula: (model) => { model.sheets[0].cells.B2 = { formula: 'A2*3' } },
    number: (model) => { model.sheets[0].cells.B2 = { value: 5 } },
    cleared: (model) => { delete model.sheets[0].cells.B2 },
    // Row 2 deleted: every cell below moves up with its own (already shifted) formula.
    rowDeleted: (model) => {
      const next = {}
      for (const [address, cell] of Object.entries(model.sheets[0].cells)) {
        const match = /^([A-Z]+)(\d+)$/.exec(address)
        const row = Number(match[2])
        if (row === 2) continue
        const moved = { ...cell }
        if (moved.formula) moved.formula = moved.formula.replace(/A(\d+)/g, (_all, number) => `A${Number(number) - 1}`)
        next[`${match[1]}${row > 2 ? row - 1 : row}`] = moved
      }
      model.sheets[0].cells = next
    },
  }
  for (const [label, edit] of Object.entries(edits)) {
    for (const baseBytes of [base, null]) {
      const model = structuredClone(imported)
      edit(model)
      const saved = await serializeWorkbook(model, 'xlsx', { baseBytes })
      assert.doesNotMatch(await sheetXml(saved), /t="shared"/, `${label}: no shared group is re-emitted`)
      const back = (await workbookPayloadFromBytes('back.xlsx', saved)).workbook.sheets[0].cells
      if (label === 'rowDeleted') {
        assert.equal(back.B2.formula, 'A2*2')
        assert.equal(back.B5.formula, 'A5*2')
      } else {
        assert.equal(back.B3.formula, 'A3*2', `${label}: the next cell keeps its formula`)
        assert.equal(back.B6.formula, 'A6*2')
      }
      if (label === 'formula') assert.equal(back.B2.formula, 'A2*3')
      if (label === 'number') assert.equal(back.B2.value, 5)
    }
  }
  await sharedFormulaLiterals()
}

async function sharedFormulaLiterals() {
  // Review F1: a filled-down formula is translated from its master without touching string
  // literals, quoted sheet names or structured references ("Q1" never becomes "Q2").
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Data')
  workbook.addWorksheet('Q1 Data')
  const master = 'IF(A1="Q1",1,0)+COUNTIF(A1:A3,"SKU100")+SUM(T1[Q1])+SUM(Sales[[#This Row],[Q1]])+\'Q1 Data\'!B1+LOG10(10)+$A$1'
  sheet.getCell('B1').value = { formula: master, result: 1, shareType: 'shared', ref: 'B1:B3' }
  for (let row = 2; row <= 3; row += 1) sheet.getCell(`B${row}`).value = { sharedFormula: 'B1', result: 0 }
  sheet.getCell('D1').value = { formula: 'SUM(IF(A1:A3>5,1,0))', result: 0 }
  sheet.getCell('E1').value = { formula: 'SUM(IF(A1:A3>5,1,0))', result: 0, shareType: 'array', ref: 'E1' }
  const base = Buffer.from(await workbook.xlsx.writeBuffer())
  assert.match(await sheetXml(base), /t="shared"/, 'the literal fixture uses a shared formula')
  const cells = (await workbookPayloadFromBytes('shared-literals.xlsx', base)).workbook.sheets[0].cells
  assert.equal(cells.B1.formula, master)
  // Review F10: a plain formula from a file calculates with implicit intersection (legacy), an
  // array formula does not.
  assert.equal(cells.D1.implicitIntersection, true, 'a plain opened formula is a legacy formula')
  assert.equal(cells.B3.implicitIntersection, true)
  assert.equal(cells.E1.implicitIntersection, undefined, 'an array formula is not')
  assert.equal(

    cells.B3.formula,
    'IF(A3="Q1",1,0)+COUNTIF(A3:A5,"SKU100")+SUM(T1[Q1])+SUM(Sales[[#This Row],[Q1]])+\'Q1 Data\'!B3+LOG10(10)+$A$1',
    'quoted text, table names and table columns keep their text in a filled cell',
  )
  const saved = await serializeWorkbook((await workbookPayloadFromBytes('shared-literals.xlsx', base)).workbook, 'xlsx', { baseBytes: base })
  const back = (await workbookPayloadFromBytes('back.xlsx', saved)).workbook.sheets[0].cells
  assert.equal(back.B2.formula, cells.B2.formula, 'the translated formula survives a save')
  assert.match(back.B2.formula, /"Q1"/)
}

async function printAreas() {
  // calc-file-io-objects-5: whole-column and whole-row print areas open, print a bounded
  // range, and are written back in their original form; an unreadable area never blocks Save.
  const build = async (reference) => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Sheet1')
    for (let row = 1; row <= 30; row += 1) sheet.getCell(`A${row}`).value = row
    sheet.getCell('H1').value = 'edge'
    sheet.pageSetup.printArea = 'A1:B2'
    const bytes = Buffer.from(await workbook.xlsx.writeBuffer())
    return patchPart(bytes, 'xl/workbook.xml', (xml) => xml.replace(/(<definedName name="_xlnm\.Print_Area"[^>]*>)[^<]*(<\/definedName>)/, (_all, open, close) => `${open}${reference}${close}`))
  }
  for (const [reference, bounded, written] of [
    ["'Sheet1'!$A:$D", 'A1:D30', "'Sheet1'!$A:$D"],
    ["'Sheet1'!$1:$20", 'A1:H20', "'Sheet1'!$1:$20"],
    ["'Sheet1'!$A$1:$C$5,'Sheet1'!$E:$F", 'A1:C5&&E1:F30', "'Sheet1'!$A$1:$C$5,'Sheet1'!$E:$F"],
  ]) {
    const base = await build(reference)
    const payload = await workbookPayloadFromBytes('print.xlsx', base)
    const setup = payload.workbook.sheets[0].pageSetup
    assert.equal(setup.printArea, bounded, `${reference} is bounded to the used size`)
    assert.doesNotMatch(setup.printArea, /NaN/)
    for (const baseBytes of [base, null]) {
      const saved = await serializeWorkbook(payload.workbook, 'xlsx', { baseBytes })
      const xml = await sheetXml(saved, 'xl/workbook.xml')
      assert.ok(xml.includes(`>${written.replace(/'/g, '&apos;')}<`) || xml.includes(`>${written}<`), `${reference} is written back as ${written}: ${xml.match(/<definedNames>[\s\S]*?<\/definedNames>/)}`)
    }
  }
  const unreadable = await build("'Sheet1'!#REF!")
  const opened = await workbookPayloadFromBytes('print.xlsx', unreadable)
  assert.equal(opened.workbook.sheets[0].pageSetup.printArea, undefined)
  assert.ok(opened.warnings.some((warning) => /print area/.test(warning)))
  await serializeWorkbook(opened.workbook, 'xlsx', { baseBytes: unreadable })
  const invalid = structuredClone(opened.workbook)
  invalid.sheets[0].pageSetup = { ...invalid.sheets[0].pageSetup, printArea: 'not a range' }
  const warnings = []
  await serializeWorkbook(invalid, 'xlsx', { warnings })
  assert.ok(warnings.some((warning) => /not a cell range/.test(warning)), 'an invalid print area is skipped with a note')
}

async function hyperlinks() {
  // calc-file-io-objects-8 and CALC-SIE-25: links to a place in the workbook, links on formula
  // cells and links whose ref is a range survive open and save.
  const workbook = new ExcelJS.Workbook()
  const toc = workbook.addWorksheet('TOC')
  workbook.addWorksheet('Data').getCell('B5').value = 5
  toc.getCell('A1').value = { text: 'Go to data', hyperlink: 'https://placeholder.invalid/' }
  toc.getCell('A2').value = { formula: '1+1', result: 2 }
  toc.getCell('A3').value = 'range one'
  toc.getCell('B3').value = 'range two'
  let bytes = Buffer.from(await workbook.xlsx.writeBuffer())
  bytes = await patchPart(bytes, 'xl/worksheets/sheet1.xml', (xml) => xml.replace(/<hyperlinks>[\s\S]*<\/hyperlinks>/, '<hyperlinks><hyperlink ref="A1" location="\'Data\'!B5" display="Go to data"/><hyperlink ref="A2" r:id="rIdX1" tooltip="Example"/><hyperlink ref="A3:B3" r:id="rIdX2"/></hyperlinks>'))
  bytes = await patchPart(bytes, 'xl/worksheets/_rels/sheet1.xml.rels', () => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdX1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/><Relationship Id="rIdX2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/page" TargetMode="External"/></Relationships>')
  const model = (await workbookPayloadFromBytes('links.xlsx', bytes)).workbook
  const cells = model.sheets[0].cells
  assert.equal(cells.A1.hyperlink, "#'Data'!B5")
  assert.equal(cells.A2.formula, '1+1', 'a linked formula keeps its formula')
  assert.equal(cells.A2.hyperlink, 'https://example.com/')
  assert.equal(cells.A2.hyperlinkTooltip, 'Example')
  assert.equal(cells.B3.hyperlink, 'https://example.org/page')
  model.sheets[0].cells.C1 = { formula: 'Data!B5*2', result: 10, hyperlink: '#Data!B5' }
  for (const baseBytes of [bytes, null]) {
    const saved = await serializeWorkbook(model, 'xlsx', { baseBytes })
    const xml = await sheetXml(saved)
    assert.match(xml, /<hyperlink ref="A1" location="&apos;Data&apos;!B5"\/>/)
    assert.doesNotMatch(xml, /<hyperlink ref="A1"[^>]*r:id=/, 'an in-workbook link needs no relationship')
    const back = (await workbookPayloadFromBytes('back.xlsx', saved)).workbook.sheets[0].cells
    assert.equal(back.A1.hyperlink, "#'Data'!B5")
    assert.equal(back.A2.formula, '1+1')
    assert.equal(back.A2.hyperlink, 'https://example.com/')
    assert.equal(back.A2.hyperlinkTooltip, 'Example')
    assert.equal(back.A3.hyperlink, 'https://example.org/page')
    assert.equal(back.C1.formula, 'Data!B5*2')
    assert.equal(back.C1.hyperlink, '#Data!B5')
  }
}

async function checkboxes() {
  // CALC-SIE-25: a checkbox reopens as a checkbox, not as a TRUE/FALSE dropdown.
  const model = {
    version: 1, name: 'checks', activeSheetId: 's1', metadata: {},
    sheets: [{ id: 's1', name: 'Tasks', rowCount: 3, colCount: 2, merges: [], colWidths: {}, rowHeights: {},
      cells: { A1: { value: true, type: 'checkbox' }, A2: { value: false, type: 'checkbox' }, B1: { value: 'Option 1', type: 'dropdown' } },
      dataValidations: { 'A1:A2': { type: 'list', allowBlank: false, formulae: ['"TRUE,FALSE"'] }, B1: { type: 'list', allowBlank: true, formulae: ['"Option 1,Option 2"'] } } }],
  }
  const back = (await workbookPayloadFromBytes('checks.xlsx', await serializeWorkbook(model, 'xlsx'))).workbook.sheets[0]
  assert.equal(back.cells.A1.type, 'checkbox')
  assert.equal(back.cells.A1.value, true)
  assert.equal(back.cells.A2.type, 'checkbox')
  assert.notEqual(back.cells.B1.type, 'checkbox')
}

async function workbookProtection() {
  // calc-file-io-objects-7: "Protect Workbook" and the file-sharing settings are written back
  // in schema order, so very hidden sheets stay locked away after a save.
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Visible').getCell('A1').value = 1
  const secret = workbook.addWorksheet('Secret')
  secret.state = 'veryHidden'
  secret.getCell('A1').value = 'salary table'
  let bytes = Buffer.from(await workbook.xlsx.writeBuffer())
  bytes = await patchPart(bytes, 'xl/workbook.xml', (xml) => xml
    .replace('<sheets>', '<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="abc=" workbookSaltValue="def=" workbookSpinCount="100000" lockStructure="1"/><sheets>')
    .replace('<workbookPr', '<fileSharing readOnlyRecommended="1" userName="Owner"/><workbookPr'))
  const model = (await workbookPayloadFromBytes('locked.xlsx', bytes)).workbook
  assert.equal(model.metadata.workbookProtection.lockStructure, '1')
  for (const baseBytes of [bytes, null]) {
    const xml = await sheetXml(await serializeWorkbook(model, 'xlsx', { baseBytes }), 'xl/workbook.xml')
    assert.match(xml, /<workbookProtection [^>]*lockStructure="1"[^>]*workbookHashValue="abc="|<workbookProtection [^>]*workbookHashValue="abc="[^>]*lockStructure="1"/)
    const order = xml.replace(/<definedNames>[\s\S]*<\/definedNames>/, '').match(/<(fileVersion|fileSharing|workbookPr|workbookProtection|bookViews|sheets)\b/g).map((tag) => tag.slice(1))
    assert.deepEqual(order.filter((tag) => tag !== 'bookViews'), ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'sheets'])
    assert.ok(order.indexOf('workbookProtection') < order.indexOf('sheets'))
  }
}

async function contentSniffing() {
  // CALC-SIE-20: the reader follows the bytes, not the extension, and never shows an empty
  // workbook for a non-empty file. A mismatched name always needs Save As.
  const legacy = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(legacy, XLSX.utils.aoa_to_sheet([['Item', 1]]), 'Legacy')
  const xlsBytes = XLSX.write(legacy, { type: 'buffer', bookType: 'biff8' })
  const xlsxBook = new ExcelJS.Workbook()
  xlsxBook.addWorksheet('Modern').getCell('A1').value = 'ok'
  const xlsxBytes = Buffer.from(await xlsxBook.xlsx.writeBuffer())
  const odsBytes = XLSX.write(legacy, { type: 'buffer', bookType: 'ods' })
  const cases = [
    ['legacy.xlsx', xlsBytes, 'xls', 'Legacy'],
    ['data.xlsx', Buffer.from('a,b\r\n1,2\r\n'), 'csv', 'Sheet1'],
    ['modern.xls', xlsxBytes, 'xlsx', 'Modern'],
    ['sheet.xlsx', odsBytes, 'ods', 'Legacy'],
    ['page.xls', Buffer.from('<html><body><table><tr><td>1</td><td>2</td></tr></table></body></html>'), 'html', null],
  ]
  for (const [name, bytes, format, firstSheet] of cases) {
    const payload = await workbookPayloadFromBytes(name, bytes, { officeEngine: false })
    assert.equal(payload.sourceFormat, format, `${name} is read as ${format}`)
    assert.equal(payload.requiresSaveAs, true, `${name} must not be overwritten in place`)
    assert.match(payload.warnings[0], /actually/)
    if (firstSheet) assert.equal(payload.workbook.sheets[0].name, firstSheet)
  }
  assert.equal((await workbookPayloadFromBytes('modern.xls', xlsxBytes)).workbook.metadata.importedWith, 'exceljs')
  const noExtension = await workbookPayloadFromBytes('export', Buffer.from('a\tb\r\n1\t2\r\n'))
  assert.equal(noExtension.sourceFormat, 'tsv')
  assert.equal(noExtension.workbook.sheets[0].cells.B2.value, 2)
  const emptyZip = await new JSZip().file('readme.txt', 'x').generateAsync({ type: 'nodebuffer' })
  await assert.rejects(workbookPayloadFromBytes('broken.xlsx', emptyZip), /not a valid XLSX package/)
  const docx = await new JSZip().file('word/document.xml', '<w:document/>').generateAsync({ type: 'nodebuffer' })
  await assert.rejects(workbookPayloadFromBytes('letter.xlsx', docx), /Word document/)
  // A package without worksheets (or without a workbook part) is never used as a save base.
  const noSheets = await new JSZip().file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>').file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>').generateAsync({ type: 'nodebuffer' })
  const model = { version: 1, name: 'x', activeSheetId: 's', metadata: {}, sheets: [{ id: 's', name: 'S', rowCount: 1, colCount: 1, merges: [], cells: { A1: { value: 1 } } }] }
  for (const baseBytes of [noSheets, emptyZip]) {
    const saved = await serializeWorkbook(model, 'xlsx', { baseBytes })
    assert.equal((await workbookPayloadFromBytes('x.xlsx', saved)).workbook.sheets[0].cells.A1.value, 1)
  }
}

async function nativeOpenDocument() {
  // CALC-SIE-13: .ods opens without the document engine, with real sheet names, A1 formulas,
  // hidden sheets and column widths. Compared with the XLSX version of the same sample.
  const odsPath = path.join(SAMPLES, 'Complex workbook.ods')
  const xlsxPath = path.join(SAMPLES, 'Complex workbook.xlsx')
  let odsBytes
  let xlsxBytes
  try {
    odsBytes = await fs.readFile(odsPath)
    xlsxBytes = await fs.readFile(xlsxPath)
  } catch {
    notes.push('native ODS sample comparison skipped (Simple test examples not found)')
    return
  }
  const ods = await workbookPayloadFromBytes('Complex workbook.ods', odsBytes, { officeEngine: false })
  const xlsx = await workbookPayloadFromBytes('Complex workbook.xlsx', xlsxBytes, { officeEngine: false })
  assert.equal(ods.workbook.metadata.importedWith, 'sheetjs-ods')
  assert.equal(ods.requiresSaveAs, true)
  assert.match(ods.warnings.join(' '), /without the document engine/)
  assert.deepEqual(ods.workbook.sheets.map((sheet) => sheet.name), xlsx.workbook.sheets.map((sheet) => sheet.name))
  assert.ok(ods.workbook.sheets.some((sheet) => sheet.name === 'Inputs & notes'), 'entities in sheet names are decoded')
  assert.deepEqual(ods.workbook.sheets.map((sheet) => sheet.state), xlsx.workbook.sheets.map((sheet) => sheet.state), 'hidden sheets stay hidden')
  const summary = ods.workbook.sheets.find((sheet) => sheet.name === 'Summary')
  assert.equal(summary.cells.B5.formula, 'SUM(Transactions!F4:F203)')
  assert.equal(summary.cells.B11.formula, "IFERROR('Inputs & notes'!B12,0)")
  assert.ok(Object.keys(summary.colWidths).length >= 3, 'column widths are read')
  let formulas = 0
  ods.workbook.sheets.forEach((sheet, index) => {
    const other = xlsx.workbook.sheets[index]
    for (const [address, cell] of Object.entries(other.cells)) {
      if (!cell.formula) continue
      formulas += 1
      assert.equal(sheet.cells[address] && sheet.cells[address].formula, cell.formula, `${sheet.name}!${address}`)
    }
    for (const [address, cell] of Object.entries(other.cells)) {
      if (cell.formula || cell.value == null || cell.value === '' || typeof cell.value === 'boolean' || cell.type === 'error' || (sheet.cells[address] && sheet.cells[address].formula)) continue
      if (sheet.name === 'Inputs & notes' && address === 'B3') continue // the two sample files use different tax rates
      const mine = sheet.cells[address] && sheet.cells[address].value
      if (typeof cell.value === 'number') assert.ok(Math.abs(mine - cell.value) < 1e-9, `${sheet.name}!${address}: ${mine} vs ${cell.value}`)
      else assert.equal(mine, cell.value, `${sheet.name}!${address}`)
    }
  })
  assert.ok(formulas > 500, `every formula is compared (${formulas})`)
  // Flat ODS with an entity-escaped name and OpenFormula references.
  const fods = `<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:automatic-styles><style:style style:name="taH" style:family="table"><style:table-properties table:display="false"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="R&amp;D"><table:table-row><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell><table:table-cell table:formula="of:=[.A1]*[&apos;Other sheet&apos;.$A$1]" office:value-type="float" office:value="6"><text:p>6</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Other sheet" table:style-name="taH"><table:table-row><table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document>`
  const flat = await workbookPayloadFromBytes('flat.fods', Buffer.from(fods, 'utf8'))
  assert.equal(flat.workbook.sheets[0].name, 'R&D')
  assert.equal(flat.workbook.sheets[0].cells.B1.formula, "A1*'Other sheet'!$A$1")
  assert.equal(flat.workbook.sheets[1].state, 'hidden')
}

async function legacyWithoutEngine(officeEngine) {
  // CALC-SIE-2 (open side): without the engine an .xls/.ods says that Save writes an .xlsx copy.
  const legacy = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(legacy, XLSX.utils.aoa_to_sheet([['a', 1]]), 'Legacy')
  const bytes = XLSX.write(legacy, { type: 'buffer', bookType: 'biff8' })
  const without = await workbookPayloadFromBytes('legacy.xls', bytes, { officeEngine: false })
  assert.equal(without.requiresSaveAs, true)
  assert.equal(without.saveAsFormat, 'xlsx')
  assert.match(without.warnings.join(' '), /creates an \.xlsx copy next to the original/)
  assert.doesNotMatch(without.warnings.join(' '), /Save keeps \.xls/)
  const imported = await importWorkbookBytes('legacy.xls', bytes, { officeEngine: false })
  assert.equal(imported.mergeBase, null, 'an XLS model has no OOXML merge base')
  if (officeEngine) {
    const withEngine = await workbookPayloadFromBytes('legacy.xls', bytes, { officeEngine: true })
    assert.equal(withEngine.requiresSaveAs, false)
  }
}

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
  // The suite runs on machines with and without the optional document engine (LibreOffice):
  // XLS/ODS go through the engine when present and through the native fallbacks otherwise.
  const officeEngine = Boolean(await findOfficeConverter())
  if (!officeEngine) notes.push('LibreOffice not found: XLS/ODS checks ran against the native (no-engine) readers and writers')
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
    assert.equal(xlsOpened.requiresSaveAs, !officeEngine, 'without the engine an edited .xls is saved as an .xlsx copy')
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

    await sharedFormulas()
    await printAreas()
    await hyperlinks()
    await checkboxes()
    await workbookProtection()
    await contentSniffing()
    await nativeOpenDocument()
    await legacyWithoutEngine(officeEngine)

    for (const note of notes) process.stdout.write(`Note: ${note}.\n`)
    process.stdout.write(`Workbook QA passed: ${opened.stats.cells} cells, ${opened.stats.formulas} formulas, XLSX/XLS/ODS/CSV/TSV (${officeEngine ? 'document engine' : 'native fallbacks'}), shared formulas, print areas, links, checkboxes, workbook protection, content sniffing.\n`)
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
