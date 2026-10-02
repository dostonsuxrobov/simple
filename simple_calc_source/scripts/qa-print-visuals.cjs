'use strict'

// Print visuals and page setup QA (run by npm run test:print through qa-print-layout.cjs).
// The renderer side (src/lib/visual-style.ts, print-visuals.ts, print-page-setup.ts) is
// bundled with esbuild; the print engine is the real electron/spreadsheet-print.cjs.
// Never sends a print job.
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const JSZip = require('jszip')
const { createSpreadsheetPrintDocument } = require('../electron/spreadsheet-print.cjs')

const ROOT = path.resolve(__dirname, '..')

function loadVisualLibrary() {
  const esbuild = require('esbuild')
  const outfile = path.join(ROOT, 'tmp', 'qa-print-visuals-lib.cjs')
  fs.mkdirSync(path.dirname(outfile), { recursive: true })
  esbuild.buildSync({
    entryPoints: [path.join(__dirname, 'qa-print-visuals-lib.ts')],
    bundle: true, platform: 'node', format: 'cjs', outfile, logLevel: 'silent',
    define: { 'process.env.NODE_ENV': '"production"' },
  })
  delete require.cache[outfile]
  return require(outfile)
}

function cellHtml(html, address) {
  return new RegExp(`<td data-address="${address}"[\\s\\S]*?</td>`).exec(html)?.[0] || ''
}

function checkPrintedVisuals(lib) {
  const input = lib.checkVisualPayload()
  const result = createSpreadsheetPrintDocument(input)
  const html = result.html
  const colors = lib.EXPECTED_COLORS
  assert.match(cellHtml(html, 'B2'), new RegExp(`background-color:${colors.scaleLow}`), 'colour scale prints')
  assert.match(cellHtml(html, 'B6'), new RegExp(`background-color:${colors.scaleHigh}`))
  assert.match(cellHtml(html, 'C4'), /background-image:linear-gradient\(90deg, #638EC6, #[0-9A-F]{6}\);background-size:50\.000% calc\(100% - 6px\);background-position:0\.000% 50%;background-repeat:no-repeat/, 'data bar prints at its length')
  assert.match(cellHtml(html, 'D2'), /<span class="print-cf"[^>]*><span class="print-cf-icon" aria-hidden="true"><svg[^>]*><circle[^>]*fill="#D6392B"/, 'icon set prints the grid glyph before the value')
  assert.match(cellHtml(html, 'D2'), /<span class="print-cf-text">5<\/span>/)
  assert.match(cellHtml(html, 'A1'), /background-color:#4472C4;color:#ffffff;font-weight:700/, 'table header style prints')
  assert.match(cellHtml(html, 'A2'), /background-color:#dae3f3/, 'banded row prints')
  assert.doesNotMatch(cellHtml(html, 'A3'), /background-color/, 'unbanded row stays white')
  assert.match(cellHtml(html, 'A2'), /border-top:1px solid #8faadc/, 'table borders print')
  assert.match(cellHtml(html, 'G2'), /class="print-checkbox"[^>]*><svg[^>]*aria-label="Checked"/, 'checkbox prints checked')
  assert.match(cellHtml(html, 'G3'), /aria-label="Not checked"/)
  assert.doesNotMatch(cellHtml(html, 'G2'), />TRUE</, 'a checkbox prints no TRUE text')
  assert.match(cellHtml(html, 'H2'), /aria-label="Checked"/, 'a TRUE/FALSE dropdown prints as a checkbox')
  assert.match(cellHtml(html, 'I2'), /color:#FF0000/, 'number format colour prints')
  assert.match(cellHtml(html, 'J2'), />25\.0%</, 'conditional number format prints its text')
  assert.match(cellHtml(html, 'F2'), /class="has-sparkline"[\s\S]*<span class="print-sparkline" aria-hidden="true"><svg[^>]*><polyline[^>]*stroke="#1a73e8"/, 'SPARKLINE() prints as a chart')
  assert.doesNotMatch(html, /\uE000|sparkline:\{/, 'the SPARKLINE marker is never printed as text')
  for (const row of [2, 3, 4, 5, 6]) assert.match(cellHtml(html, `E${row}`), /<polyline[^>]*stroke="#376092"/, `sparkline group E${row}`)
  assert.equal(result.pageCount, 1)
  return input
}

function checkVisualSanitization(input) {
  const hostile = JSON.parse(JSON.stringify(input))
  const sheet = hostile.visuals.sheets.report
  const index = sheet.styles.length
  sheet.styles.push(
    { fill: 'red;background:url(http://example.invalid/x.png)', color: '#12345', bold: 'yes', decoration: 'blink', borders: { top: '1px solid red', left: '1px solid #000000;x:url(a)' } },
    { bar: { image: 'url(http://example.invalid/bar.png)', size: '50% 100%', position: '0 0', repeat: 'no-repeat' } },
    { bar: { image: 'linear-gradient(#000000, #ffffff), image-set(x)', size: '50% 100%', position: '0 0', repeat: 'no-repeat' } },
    { bar: { image: 'linear-gradient(#000000, #ffffff)', size: '50% 100%', position: '(0) 0', repeat: 'no-repeat' } },
    { icon: 'missing:1', fill: '#ABCDEF' },
  )
  sheet.cells.K1 = index
  sheet.cells.K2 = index + 1
  sheet.cells.K3 = index + 2
  sheet.cells.K4 = index + 3
  sheet.cells.K5 = index + 4
  sheet.cells['not an address'] = 0
  sheet.text = { ...sheet.text, K1: '<b onclick="x()">bold</b>' }
  sheet.sparklines = { ...sheet.sparklines, K2: '<svg><script>alert(1)</script></svg>', K3: '<svg onload="alert(1)"><rect/></svg>' }
  hostile.visuals.icons['3TrafficLights1:1'] = '<svg><image href="http://example.invalid/i.png"/></svg>'
  hostile.workbook.sheets[0].cells.K1 = { value: 'x' }
  const html = createSpreadsheetPrintDocument(hostile).html
  assert.doesNotMatch(html, /url\(|image-set|<script|<image|example\.invalid|blink/, 'unsafe visuals never reach the page')
  assert.doesNotMatch(html, /<[^>]+\son[a-z]+=/i, 'no event-handler attribute reaches the page')
  assert.doesNotMatch(cellHtml(html, 'K1'), /background-color|color:#12345|font-weight:yes|border-top/)
  assert.match(cellHtml(html, 'K1'), /&lt;b onclick=&quot;x\(\)&quot;&gt;bold&lt;\/b&gt;/, 'text overrides are escaped')
  assert.doesNotMatch(cellHtml(html, 'K2'), /background-image/)
  assert.doesNotMatch(cellHtml(html, 'K3'), /background-image/)
  assert.doesNotMatch(cellHtml(html, 'K4'), /background-image/)
  assert.match(cellHtml(html, 'K5'), /background-color:#ABCDEF/, 'a safe property survives next to an unknown icon')
  assert.doesNotMatch(cellHtml(html, 'D3'), /print-cf-icon/, 'an unsafe icon glyph is dropped, the value still prints')
  assert.match(cellHtml(html, 'D3'), />35</)
  // Without renderer visuals the engine still never prints the marker (a direct caller may send
  // the raw SPARKLINE value), and checkbox controls print as boxes.
  const raw = JSON.parse(JSON.stringify(input))
  raw.displayValues.report.F2 = '\uE000sparkline:{"d":[1,2,3],"o":{}}'
  raw.workbook.sheets[0].cells.F2.result = raw.displayValues.report.F2
  const plain = createSpreadsheetPrintDocument({ ...raw, visuals: undefined }).html
  assert.match(cellHtml(plain, 'F2'), /<span style="max-height:[0-9.]+px"><\/span>/, 'the raw marker prints as an empty cell')
  assert.doesNotMatch(plain, /\uE000|sparkline:\{/)
  assert.match(cellHtml(plain, 'G2'), /aria-label="Checked"/)
  assert.match(cellHtml(plain, 'G3'), /aria-label="Not checked"/)
}

function checkAuthoredPageSetup(lib) {
  const cells = {}
  for (let row = 1; row <= 60; row += 1) {
    cells[`A${row}`] = { value: `Label ${row}` }
    for (let col = 1; col < 20; col += 1) cells[`${String.fromCharCode(65 + col)}${row}`] = { value: row * col }
  }
  const sheet = { id: 'setup', name: 'Setup', rowCount: 60, colCount: 20, cells, merges: [], colWidths: {}, rowHeights: {}, frozen: { rows: 1, columns: 1 } }
  lib.applyPageSetupPatch(sheet, { printArea: 'B2:T40', printTitlesRow: '1:1', printTitlesColumn: 'A:A', oddHeader: '&L&A&R&D', oddFooter: '&CPage &P of &N', rowBreaks: [21] })
  const request = (options = {}) => ({ name: 'Setup.xlsx', workbook: { version: 1, name: 'Setup.xlsx', activeSheetId: 'setup', sheets: [sheet] }, options: { scope: 'active-sheet', scaling: 'actual', margins: 'normal', gridlines: false, ...options } })
  const result = createSpreadsheetPrintDocument(request())
  const pages = [...result.html.matchAll(/data-row-range="([^"]*)" data-column-range="([^"]*)"/g)].map((match) => `${match[1]}|${match[2]}`)
  // Titles outside the print area lead every page, the first included (as in Excel); the
  // manual break above row 21 ends the first band of rows; columns B:T split after K.
  assert.deepEqual(pages, ['1:20|A:K', '1:20|A:T', '1:40|A:K', '1:40|A:T'], `pages: ${pages.join(' ')}`)
  for (const page of result.html.split('<section').slice(1)) {
    assert.match(page, /<td data-address="A1"/, 'every page carries title row 1')
    assert.match(page, /<td data-address="A\d+"[^>]*><span[^>]*>Label \d+</, 'every page carries title column A')
  }
  assert.doesNotMatch(result.html.split('<section')[2], /data-address="K\d+"/, 'the second column page holds L:T after the title column')
  assert.match(result.html, /class="header-footer-left"><span[^>]*>Setup<\/span>/, 'the authored header prints the sheet name')
  assert.match(result.html, new RegExp(`Page 1 of ${result.pageCount}`), 'the authored footer numbers pages')
  // Rows to repeat that fill the page are reported, not silently dropped.
  sheet.rowHeights = Object.fromEntries(Array.from({ length: 30 }, (_, index) => [String(index + 1), 40]))
  lib.applyPageSetupPatch(sheet, { printTitlesRow: '1:30' })
  assert.throws(() => createSpreadsheetPrintDocument(request()), /repeated heading rows fill the page/)
  sheet.rowHeights = {}
  lib.applyPageSetupPatch(sheet, { printArea: null, printTitlesRow: null, printTitlesColumn: null, rowBreaks: [] })
  const state = lib.pageSetupStateFor(sheet, { top: 0, bottom: 0, left: 0, right: 0 })
  assert.equal(state.printArea, '')
  assert.deepEqual(state.rowBreaks, [])
  assert.equal(createSpreadsheetPrintDocument(request()).printedCells, 60 * 20, 'clearing the print area prints the used range again')
}

async function checkPageSetupSaves(lib) {
  const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')
  const cells = {}
  for (let row = 1; row <= 30; row += 1) cells[`A${row}`] = { value: row }
  const sheet = { id: 's1', name: 'Report', rowCount: 30, colCount: 4, cells, merges: [], colWidths: {}, rowHeights: {} }
  lib.applyPageSetupPatch(sheet, { printArea: 'A1:B20', printTitlesRow: '1', printTitlesColumn: 'A', oddHeader: '&CQuarterly &A', oddFooter: '&RPage &P of &N', rowBreaks: [11] })
  const workbook = { version: 1, name: 'Saved.xlsx', activeSheetId: 's1', sheets: [sheet], metadata: {} }
  const bytes = await serializeWorkbook(workbook, 'xlsx', {})
  const zip = await JSZip.loadAsync(bytes)
  const book = await zip.file('xl/workbook.xml').async('string')
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string')
  assert.match(book, /<definedName name="_xlnm\.Print_Area" localSheetId="0">'Report'!\$A\$1:\$B\$20<\/definedName>/, 'the print area is saved as Excel\'s defined name')
  assert.match(book, /<definedName name="_xlnm\.Print_Titles" localSheetId="0">'Report'!\$1:\$1,'Report'!\$A:\$A<\/definedName>/, 'rows and columns to repeat are saved')
  assert.match(xml, /<oddHeader>&amp;CQuarterly &amp;A<\/oddHeader><oddFooter>&amp;RPage &amp;P of &amp;N<\/oddFooter>/, 'header and footer are saved')
  assert.match(xml, /<rowBreaks count="1" manualBreakCount="1"><brk id="10" max="16383" man="1"\/><\/rowBreaks>/, 'the manual break above row 11 is saved')
  const reopened = (await workbookPayloadFromBytes('Saved.xlsx', bytes)).workbook.sheets[0]
  assert.equal(reopened.pageSetup.printArea, 'A1:B20')
  assert.equal(reopened.pageSetup.printTitlesRow, '1:1')
  assert.equal(reopened.pageSetup.printTitlesColumn, 'A:A')
  assert.equal(reopened.headerFooter.oddHeader, '&CQuarterly &A')
  // Known gap outside the print module: the importer does not read <rowBreaks> back (ExcelJS's
  // reader never parses them), so a reopened file loses manual breaks. Reported, not hidden.
  return { breaksReimported: Array.isArray(reopened.rowBreaks) && reopened.rowBreaks.length > 0 }
}

async function run() {
  const lib = loadVisualLibrary()
  const parityCases = lib.checkGridParity()
  lib.checkVisualCache()
  lib.checkPageSetupHelpers()
  const input = checkPrintedVisuals(lib)
  checkVisualSanitization(input)
  checkAuthoredPageSetup(lib)
  const saved = await checkPageSetupSaves(lib)
  console.log(`Print visuals QA passed: ${parityCases} grid-parity cases (table styles, number-format colours, conditional fills, fonts, borders, data bars, icons, checkboxes), colour scale, data bar, icon set, banded table, sparklines, checkboxes and CF number formats on paper; hostile visuals rejected; authored print area, title rows/columns, header/footer and page breaks printed and saved${saved.breaksReimported ? ' and reopened' : ' (manual breaks are not re-imported by workbooks.cjs yet)'}.`)
}

module.exports = { run, loadVisualLibrary }

if (require.main === module) {
  run().catch((error) => { console.error(error); process.exitCode = 1 })
}
