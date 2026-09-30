const assert = require('node:assert/strict')
const { createSpreadsheetPrintDocument, parseRange, usedBounds } = require('../electron/spreadsheet-print.cjs')

function sheet(id, name, cells = {}, extras = {}) {
  return {
    id,
    name,
    state: 'visible',
    rowCount: 200,
    colCount: 40,
    cells,
    merges: [],
    colWidths: {},
    rowHeights: {},
    ...extras,
  }
}

function request(options = {}, overrides = {}) {
  const first = sheet('sheet-1', 'Main <sheet>', {
    A1: { value: '<Revenue & cost>', style: { font: { bold: true, color: { argb: 'FF123456' } }, fill: { pattern: 'solid', fgColor: { argb: 'FFFFEEDD' } } } },
    B2: { value: 42 },
    C3: { value: 'outside' },
  })
  const second = sheet('sheet-2', 'Second', { A1: { value: 'other sheet' } })
  const hidden = sheet('sheet-3', 'Hidden', { A1: { value: 'do not print' } }, { state: 'hidden' })
  return {
    documentId: 'document-1',
    name: 'Quarterly <Plan>.xlsx',
    workbook: {
      version: 1,
      name: 'Quarterly Plan.xlsx',
      activeSheetId: 'sheet-1',
      sheets: [first, second, hidden],
      metadata: { themeColors: ['FFFFFFFF', 'FF000000'] },
    },
    displayValues: {
      'sheet-1': { A1: '<Revenue & cost>', B2: '$42.00', C3: 'outside' },
      'sheet-2': { A1: 'other sheet' },
      'sheet-3': { A1: 'do not print' },
    },
    selection: { top: 0, bottom: 1, left: 0, right: 1 },
    options: {
      scope: 'active-sheet',
      orientation: 'portrait',
      scaling: 'fit-width',
      paperSize: 'letter',
      margins: 'normal',
      gridlines: true,
      headings: false,
      ...options,
    },
    ...overrides,
  }
}

{
  const result = createSpreadsheetPrintDocument(request())
  assert.equal(result.sheetCount, 1)
  assert.equal(result.pageCount, 1)
  assert.equal(result.pageBreaks, 0)
  assert.equal(result.oversizedDimensions, 0)
  assert.deepEqual(result.paper, { label: 'Letter', widthInches: 8.5, heightInches: 11 })
  assert.equal(result.printedCells, 9, 'active sheet should use the populated A1:C3 range, not the 200x40 viewport')
  assert.match(result.html, /data-print-scope="active-sheet"/)
  assert.match(result.html, /data-orientation="portrait"/)
  assert.match(result.html, /@page \{ size: Letter portrait;/)
  assert.match(result.html, /@page \{ size: Letter portrait; margin: 0; \}/)
  assert.match(result.html, /class="print-page print-sheet"/)
  assert.match(result.html, /--page-width:816\.00px/)
  assert.match(result.html, /--margin-left:48\.00px/)
  assert.match(result.html, /@media screen/)
  assert.match(result.html, /@media print/)
  assert.match(result.html, /class="sheet-table show-gridlines"/)
  assert.doesNotMatch(result.html, /class="column-heading"/)
  assert.match(result.html, /&lt;Revenue &amp; cost&gt;/)
  assert.doesNotMatch(result.html, /<Revenue & cost>/)
  assert.match(result.html, /font-weight:700/)
  assert.match(result.html, /color:#123456/)
  assert.match(result.html, /background-color:#FFEEDD/i)
  assert.doesNotMatch(result.html, /font-size:6pt/, 'cells without an explicit font size should inherit the readable print default')
}

{
  const oversized = sheet('oversized', 'Oversized', { A1: { value: 'too wide and tall' } }, {
    colWidths: { 1: 125 },
    rowHeights: { 1: 480 },
  })
  const result = createSpreadsheetPrintDocument(request({ scaling: 'actual', paperSize: 'a4', orientation: 'landscape', margins: 'wide', headings: true }, {
    workbook: { version: 1, name: 'Oversized.xlsx', activeSheetId: 'oversized', sheets: [oversized] },
    displayValues: { oversized: { A1: 'too wide and tall' } },
  }))
  assert.equal(result.oversizedDimensions, 2, 'actual-size output must disclose every unpageable row and column')
}

{
  const cells = {}
  const display = {}
  for (let row = 1; row <= 50; row += 1) {
    cells[`A${row}`] = { value: row }
    display[`A${row}`] = String(row)
  }
  const tall = sheet('tall', 'Tall', cells)
  const narrow = createSpreadsheetPrintDocument(request({ scaling: 'actual', margins: 'narrow' }, {
    workbook: { version: 1, name: 'Tall.xlsx', activeSheetId: 'tall', sheets: [tall] },
    displayValues: { tall: display },
  }))
  const wide = createSpreadsheetPrintDocument(request({ scaling: 'actual', margins: 'wide' }, {
    workbook: { version: 1, name: 'Tall.xlsx', activeSheetId: 'tall', sheets: [tall] },
    displayValues: { tall: display },
  }))
  assert.equal(narrow.pageCount, 1, 'narrow margins should keep fifty default-height rows on one Letter page')
  assert.equal(wide.pageCount, 2, 'wide margins should create a visible page break for the same rows')
  assert.equal(wide.pageBreaks, 1)
  assert.match(wide.html, /data-page-number="2"/)
}

{
  const wideSheet = sheet('wide-pages', 'Wide pages', { A1: { value: 'left' }, Z1: { value: 'right' } })
  const portrait = createSpreadsheetPrintDocument(request({ scaling: 'actual', orientation: 'portrait' }, {
    workbook: { version: 1, name: 'Wide pages.xlsx', activeSheetId: 'wide-pages', sheets: [wideSheet] },
    displayValues: { 'wide-pages': { A1: 'left', Z1: 'right' } },
  }))
  const landscape = createSpreadsheetPrintDocument(request({ scaling: 'actual', orientation: 'landscape' }, {
    workbook: { version: 1, name: 'Wide pages.xlsx', activeSheetId: 'wide-pages', sheets: [wideSheet] },
    displayValues: { 'wide-pages': { A1: 'left', Z1: 'right' } },
  }))
  const fitSheet = createSpreadsheetPrintDocument(request({ scaling: 'fit-sheet' }, {
    workbook: { version: 1, name: 'Wide pages.xlsx', activeSheetId: 'wide-pages', sheets: [wideSheet] },
    displayValues: { 'wide-pages': { A1: 'left', Z1: 'right' } },
  }))
  assert.ok(portrait.pageCount > 1, 'actual size should create horizontal paper pages for wide content')
  assert.ok(landscape.pageCount < portrait.pageCount, 'landscape should fit more columns on each physical page')
  assert.equal(fitSheet.pageCount, 1, 'fit sheet should produce exactly one physical page')
  assert.match(portrait.html, /data-column-range="A:K"/)
  assert.match(portrait.html, /data-column-range="W:Z"/)
}

{
  const hiddenOnly = sheet('hidden-only', 'Hidden only', { A1: { value: 'private' } }, { hiddenRows: [1], hiddenCols: [1] })
  const result = createSpreadsheetPrintDocument(request({}, {
    workbook: { version: 1, name: 'Hidden.xlsx', activeSheetId: 'hidden-only', sheets: [hiddenOnly] },
    displayValues: { 'hidden-only': { A1: 'private' } },
  }))
  assert.equal(result.printedCells, 0)
  assert.doesNotMatch(result.html, /data-address="A1"/)
  assert.doesNotMatch(result.html, />private</)
}

{
  const result = createSpreadsheetPrintDocument(request({
    scope: 'selection',
    orientation: 'landscape',
    scaling: 'actual',
    paperSize: 'a4',
    gridlines: false,
    headings: true,
  }))
  assert.equal(result.printedCells, 4)
  assert.match(result.html, /@page \{ size: A4 landscape;/)
  assert.match(result.html, /data-scaling="actual"/)
  assert.match(result.html, /data-scale="1\.0000"/)
  assert.match(result.html, /class="sheet-table show-headings"/)
  assert.doesNotMatch(result.html, /sheet-table show-gridlines/)
  assert.match(result.html, /class="column-heading" scope="col">A</)
  assert.match(result.html, /class="row-heading" scope="row"><span[^>]*>2<\/span>/)
  assert.doesNotMatch(result.html, />outside</)
}

{
  const input = request({ scope: 'workbook', scaling: 'fit-sheet' })
  const result = createSpreadsheetPrintDocument(input)
  assert.equal(result.sheetCount, 2, 'hidden sheets must stay hidden when printing the workbook')
  assert.match(result.html, /data-sheet-name="Second"/)
  assert.doesNotMatch(result.html, /data-sheet-name="Hidden"/)
  assert.doesNotMatch(result.html, />do not print</)
}

{
  const mergedSheet = sheet('merged', 'Merged', {
    A1: { value: 'merged value' },
    C3: { value: 'visible' },
  }, {
    merges: ['A1:B2'],
    hiddenRows: [3],
    hiddenCols: [3],
  })
  const input = request({ scope: 'selection' }, {
    workbook: { version: 1, name: 'Merged.xlsx', activeSheetId: 'merged', sheets: [mergedSheet] },
    displayValues: { merged: { A1: 'merged value', C3: 'visible' } },
    selection: { top: 1, bottom: 3, left: 0, right: 3 },
  })
  const result = createSpreadsheetPrintDocument(input)
  assert.equal(result.printedCells, 6, 'hidden row 3 and column C should not occupy print cells')
  assert.match(result.html, /data-address="A2"[^>]*colspan="2"/)
  assert.match(result.html, />merged value</)
  assert.doesNotMatch(result.html, /data-address="C3"/)
}

{
  const wide = sheet('wide', 'Wide', { A1: { value: 'a' }, Z1: { value: 'z' } })
  const result = createSpreadsheetPrintDocument(request({ scaling: 'fit-width' }, {
    workbook: { version: 1, name: 'Wide.xlsx', activeSheetId: 'wide', sheets: [wide] },
    displayValues: { wide: { A1: 'a', Z1: 'z' } },
  }))
  const scale = Number(/data-scale="([0-9.]+)"/.exec(result.html)?.[1])
  assert.ok(scale > 0 && scale < 1, `wide sheet should scale down, received ${scale}`)
}

{
  assert.deepEqual(parseRange("'Sheet 1'!$B$2:$D$8"), { top: 1, bottom: 7, left: 1, right: 3 })
  assert.deepEqual(usedBounds(sheet('used', 'Used', { D9: { value: 1 } }, { merges: ['B2:C4'] })), { top: 1, bottom: 8, left: 1, right: 3 })
}

{
  assert.throws(() => createSpreadsheetPrintDocument(request({ scope: 'selection' }, {
    selection: { top: 0, bottom: 1000, left: 0, right: 1000 },
  })), /too large to preview/i)
  const tooManySheets = Array.from({ length: 101 }, (_, index) => sheet(`sheet-${index}`, `Sheet ${index + 1}`, { A1: { value: index } }))
  assert.throws(() => createSpreadsheetPrintDocument(request({ scope: 'workbook' }, {
    workbook: { version: 1, name: 'Too many.xlsx', activeSheetId: 'sheet-0', sheets: tooManySheets },
  })), /more than 100 visible sheets/i, 'the print engine must not silently omit visible worksheets')
  assert.throws(() => createSpreadsheetPrintDocument({}), /invalid print request/i)
}

{
  const source = request({ useSavedLayout: true })
  source.workbook.sheets[0].pageSetup = {
    printArea: '$A$1:$B$2', paperSize: 9, orientation: 'landscape',
    fitToPage: false, fitToWidth: 1, fitToHeight: 1, scale: 91,
    margins: { top: 1.81, bottom: 0.5, left: 0.75, right: 0.75, header: 0.66, footer: 0.48 },
    showGridLines: false, showRowColHeaders: false,
  }
  source.workbook.sheets[0].headerFooter = {
    oddHeader: '&L&K000000&G&R&"Arial,Bold"&14&K000000\nSCHEDULE OF ACCOUNTS\n&12\n&"Arial,Regular"Page &P of &N',
    oddFooter: '&L&F&C&A&R&"Arial,Italic"&8I',
  }
  const result = createSpreadsheetPrintDocument(source)
  assert.equal(result.printedCells, 4, 'stored print area excludes cells outside the source form')
  assert.equal(result.options.paperSize, 'a4', 'native host gets the effective saved paper size')
  assert.equal(result.options.orientation, 'landscape')
  assert.match(result.html, /--margin-top:173\.76px/)
  assert.match(result.html, /--header-top:63\.36px/)
  assert.match(result.html, /--footer-bottom:46\.08px/)
  assert.match(result.html, /data-scale="0\.9100"/, 'inactive fit1x1 must not override saved91%')
  assert.match(result.html, /font-size:12\.74pt[^>]*font-weight:700/)
  assert.match(result.html, /SCHEDULE OF ACCOUNTS/)
  assert.match(result.html, /Page 1 of 1/)
  assert.match(result.html, /Quarterly &lt;Plan&gt;\.xlsx/)
  assert.match(result.html, /Main &lt;sheet&gt;/)
  assert.doesNotMatch(result.html, /&amp;K000000|&amp;G|show-gridlines"/)
  assert.equal(result.warnings.length, 1)
  assert.match(result.warnings[0], /header or footer picture cannot be rendered/)
  assert.doesNotMatch(result.html, /data-address="C3"/)
  source.options.scope = 'selection'
  source.selection = { top: 2, bottom: 2, left: 2, right: 2 }
  assert.match(createSpreadsheetPrintDocument(source).html, /data-address="C3"/, 'explicit selection takes priority over a saved print area')
  source.options.scope = 'active-sheet'
  source.options.useSavedLayout = false
  const manual = createSpreadsheetPrintDocument(source)
  assert.equal(manual.options.paperSize, 'letter')
  assert.match(manual.html, /--margin-top:48\.00px/)
  assert.equal(manual.printedCells, 4, 'manual page settings still retain the print area')
}

{
  const source = request({ useSavedLayout: true })
  source.workbook.sheets[0].pageSetup = { printArea: "'Main, name'!$A$1:$A$1,'Main, name'!$C$3:$C$3", scale: 100, fitToPage: false }
  source.workbook.sheets[0].headerFooter = { differentFirst: true, firstHeader: '&CFirst', oddHeader: '&CPage &P of &N', differentOddEven: true, evenHeader: '&CEven &P of &N' }
  const result = createSpreadsheetPrintDocument(source)
  assert.equal(result.pageCount, 2, 'non-contiguous print areas are separate pages')
  assert.equal(result.printedCells, 2)
  assert.match(result.html, />First</)
  assert.match(result.html, />Even 2 of 2</)
  assert.doesNotMatch(result.html, /data-address="B2"/)
  source.workbook.sheets[0].pageSetup.printArea = 'A1:B2:C3'
  assert.throws(() => createSpreadsheetPrintDocument(source), /saved print area.*Select the cells/)
}

{
  const source = request({ scope: 'workbook', useSavedLayout: true })
  source.workbook.sheets[0].pageSetup = { paperSize: 9, orientation: 'landscape' }
  source.workbook.sheets[1].pageSetup = { paperSize: 5, orientation: 'portrait' }
  const result = createSpreadsheetPrintDocument(source)
  assert.match(result.html, /@page sheet-0 \{ size: A4 landscape;/)
  assert.match(result.html, /@page sheet-1 \{ size: Legal portrait;/)
  assert.equal(result.mixedPaperSizes, true)
  assert.match(result.warnings.join(' '), /different paper sizes or orientations/)
}

{
  const source = request({ useSavedLayout: true })
  const wide = source.workbook.sheets[0]
  wide.cells.Z80 = { value: 'last' }
  wide.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 1, scale: 91 }
  const fit = createSpreadsheetPrintDocument(source)
  assert.equal(fit.pageCount, 1)
  assert.ok(fit.minimumScale < 0.91)
  wide.pageSetup.fitToWidth = 2
  wide.pageSetup.fitToHeight = 2
  assert.ok(createSpreadsheetPrintDocument(source).pageCount <= 4, 'saved multi-page fit respects both axes')
  wide.pageSetup = { fitToPage: false, scale: 200, margins: { left: 5, right: 5 } }
  assert.throws(() => createSpreadsheetPrintDocument(source), /margins leave no room/)
}

{
  const source = request()
  source.workbook.sheets[0].headerFooter = {
    oddHeader: '&L&"Arial; color:red}<script>,Bold"&12<svg onload="alert(1)">&& &P+2 &N &F &A &Q',
    oddFooter: '&C&Bbold&B &Iitalic&I &Uunder&U &Sstrike&S',
  }
  const result = createSpreadsheetPrintDocument(source)
  assert.doesNotMatch(result.html, /<script>|<svg|<img|;color:red/)
  assert.match(result.html, /&lt;svg onload=&quot;alert\(1\)&quot;&gt;&amp; 3 1/)
  assert.match(result.html, /&amp;Q/)
  assert.match(result.warnings.join(' '), /unsupported header or footer code/)
}

{
  const source = request()
  const s = source.workbook.sheets[0]
  s.cells = { A1: { value: 'Repeated heading' }, A2: { value: 'First data row' }, A70: { value: 'Last data row' } }
  s.rowHeights = {}; s.properties = { defaultRowHeight: 20 }
  s.pageSetup = { printArea: 'A1:B70', printTitlesRow: '1:1' }
  source.displayValues = {}
  source.options = { scope: 'active-sheet', scaling: 'actual', gridlines: false }
  const result = createSpreadsheetPrintDocument(source)
  assert.equal((result.html.match(/Repeated heading/g) || []).length, result.pageCount, 'every continuation page repeats its saved heading')
  assert.equal((result.html.match(/First data row/g) || []).length, 1)
  assert.equal((result.html.match(/Last data row/g) || []).length, 1)
  const rectangle = { style: 'thin', color: { argb: 'FF102030' } }
  s.cells = { A1: { value: 'Master-only border', style: { border: { top: rectangle, right: rectangle, bottom: rectangle, left: rectangle } } } }
  s.merges = ['A1:B2']; s.pageSetup = { printArea: 'A1:B2' }
  const box = createSpreadsheetPrintDocument(source).html
  for (const side of ['top','right','bottom','left']) assert.match(box, new RegExp(`border-${side}:1px solid #102030`))
  s.headerFooter = { oddFooter: '&P of &N' }
  s.pageSetup = { printArea: 'A1:B70', firstPageNumber: 1, useFirstPageNumber: false }
  const numbered = createSpreadsheetPrintDocument(source)
  assert.match(numbered.html, new RegExp(`>${numbered.pageCount} of ${numbered.pageCount}<`), 'dormant start number does not restart each page')
}

async function verifyLegacyForm() {
  const path = require('node:path')
  const { workbookPayloadFromPath } = require('../electron/workbooks.cjs')
  const payload = await workbookPayloadFromPath(path.join(__dirname, 'fixtures', 'schedule_template.xls'))
  const workbook = payload.workbook
  const first = workbook.sheets[0]
  assert.equal(first.pageSetup.printArea, 'A1:E37')
  assert.equal(first.pageSetup.fitToPage, false)
  // Arial10 Normal-style digits measured by Chromium at96dpi; the app sends
  // this metric from Canvas for the actual installed font, not this fixture.
  first.properties.printDigitWidth = 7.415
  const source = { name: payload.name, workbook, options: { useSavedLayout: true } }
  const result = createSpreadsheetPrintDocument(source)
  assert.equal(result.printedCells, 185)
  assert.equal(result.minimumScale, 0.91)
  assert.equal(result.pageCount, 1, 'source Normal font keeps the complete A:E form together at the stored91% scale')
  assert.match(result.html, /data-row-range="1:37" data-column-range="A:E"/)
  assert.match(result.html, /--margin-top:173\.76px/)
  assert.match(result.html, /SCHEDULE OF ACCOUNTS/)
  assert.match(result.html, /Page 1 of 1/)
  assert.match(result.warnings.join(' '), /header or footer picture cannot be rendered/)
  assert.doesNotMatch(result.html, /&amp;K000000/)
  assert.doesNotMatch(result.html, /&nbsp;/, 'empty spacer rows must not acquire an intrinsic text line box')
  const cellHtml = (html, address) => new RegExp(`<td data-address="${address}"[\\s\\S]*?<\\/td>`).exec(html)?.[0] || ''
  const invoiceLabel = cellHtml(result.html, 'B30')
  assert.match(invoiceLabel, /class="print-overflow"/, 'the full schedule label can flow through empty C30/D30')
  assert.match(invoiceLabel, /INVOICES - THIS SCHEDULE:/)
  const permittedWidth = Number(/;width:([0-9.]+)px/.exec(invoiceLabel)?.[1])
  assert.ok(permittedWidth > 350 && permittedWidth < 356, 'overflow stops before the occupied E30 total')
  for (const address of ['A35', 'A37', 'D37']) {
    assert.match(cellHtml(result.html, address), /border-right:1px solid #000000/, `${address} must retain the closing edge stored on its last merged cell`)
  }
  first.cells.E2.value = 123.45
  first.cells.E3.value = 246.9
  first.cells.E30.result = undefined
  source.displayValues = { [first.id]: { E2: '$123.45', E3: '$246.90', E30: '$370.35' } }
  source.displayParts = { [first.id]: { E30: { type: 'number', accounting: { symbol: '$', amount: '370.35' } } } }
  const edited = createSpreadsheetPrintDocument(source)
  for (const address of ['E2', 'E3', 'E30']) {
    assert.match(cellHtml(edited.html, address), /text-align:right/, `${address} should use numeric alignment even with a currency-prefixed formatted value`)
    assert.match(cellHtml(edited.html, address), /class="accounting-symbol">\$<\/span><span class="accounting-amount">/, `${address} prints the same separated currency/amount layout as the grid`)
  }
  assert.match(cellHtml(edited.html, 'E30'), />370\.35<\/span>/, 'live formula display parts do not depend on a cached model result')
  first.cells.E2.value = '$123.45'
  assert.doesNotMatch(cellHtml(createSpreadsheetPrintDocument(source).html, 'E2'), /class="print-accounting"/, 'text resembling currency is still text')
  first.properties.printDigitWidth = NaN
  assert.equal(createSpreadsheetPrintDocument(source).pageCount, 2, 'invalid font measurements safely fall back to the current grid metric')
}

verifyLegacyForm().then(() => {
  console.log('Print layout QA passed: real legacy form and Normal-font widths, scopes, exact paper pages, saved print areas, header/footer fields and escaping, source scale vs fit, saved margins, mixed paper sizes, page breaks, orientation, scaling, headings, gridlines, merges, hidden dimensions, and size limits.')
}).catch((error) => { console.error(error); process.exitCode = 1 })
