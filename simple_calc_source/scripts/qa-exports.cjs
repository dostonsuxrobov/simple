'use strict'

const assert = require('node:assert/strict')
const JSZip = require('jszip')
const XLSX = require('xlsx')
const { workbookPayloadFromBytes } = require('../electron/workbooks.cjs')
const {
  createSpreadsheetExport,
  exportFilter,
  exportLosses,
  normalizeExportFormat,
} = require('../electron/spreadsheet-export.cjs')
const { findOfficeConverter } = require('../electron/office-converter.cjs')

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

  const officeEngine = Boolean(await findOfficeConverter())
  if (officeEngine) {
    const ods = await createSpreadsheetExport(request(), 'ods')
    assert.ok(Buffer.isBuffer(ods.bytes) && ods.bytes.length > 500)
    const reopenedOds = await workbookPayloadFromBytes('export.ods', ods.bytes)
    assert.equal(reopenedOds.workbook.sheets.length, 3)
  } else {
    console.log('Note: LibreOffice not found; the engine-converted XLS/ODS exports were skipped and the native writers were tested.')
  }
  await nativeExports()
  const visuals = await visualExports()

  await assert.rejects(() => createSpreadsheetExport({}, 'xlsx'), /invalid spreadsheet export request/i)
  console.log(`Export QA passed: XLSX, ODS, CSV, TSV, PDF print input, HTML, escaping, visible-sheet scope, format validation, ${officeEngine ? 'engine and ' : ''}native XLS/ODS writers, and conditional formatting, banded tables, sparklines and checkboxes in ${visuals}.`)
}

// ---------------------------------------------------------------------------
// CALC-008: what the grid paints beyond cell formatting reaches HTML and PDF.
// ---------------------------------------------------------------------------
const path = require('node:path')
const os = require('node:os')
const fs = require('node:fs')
const zlib = require('node:zlib')
const { spawn, spawnSync } = require('node:child_process')

function cellHtml(html, address) {
  return new RegExp(`<td data-address="${address}"[\\s\\S]*?</td>`).exec(html)?.[0] || ''
}

/** Decompressed PDF content: page streams, object streams and the file itself. */
function pdfContent(bytes) {
  const raw = bytes.toString('latin1')
  const parts = [raw]
  const marker = /stream\r?\n/g
  let match
  while ((match = marker.exec(raw))) {
    const start = match.index + match[0].length
    const end = raw.indexOf('endstream', start)
    if (end < 0) break
    try { parts.push(zlib.inflateSync(bytes.subarray(start, end), { finishFlush: zlib.constants.Z_SYNC_FLUSH }).toString('latin1')) } catch {}
    marker.lastIndex = end
  }
  return parts.join('\n')
}

function pdfColors(content) {
  const colors = []
  for (const op of content.matchAll(/(\d*\.?\d+)\s+(\d*\.?\d+)\s+(\d*\.?\d+)\s+(?:rg|RG|sc|scn|SC|SCN)\b/g)) colors.push([Number(op[1]), Number(op[2]), Number(op[3])])
  return colors
}

function hasPdfColor(colors, hex, tolerance = 0.012) {
  const value = Number.parseInt(hex.replace(/^#/, ''), 16)
  const target = [((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255]
  return colors.some((color) => color.every((channel, index) => Math.abs(channel - target[index]) <= tolerance))
}

function killTree(child) {
  if (!child || child.exitCode !== null) return
  if (process.platform === 'win32') spawnSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true, stdio: 'ignore' })
  else child.kill('SIGKILL')
}

/** Render a print document to PDF in a hidden Electron window (same options as main.cjs). */
async function renderInElectron(printDocument, probes, timeoutMs = 90_000) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'calc-export-visuals-'))
  try {
    const input = path.join(directory, 'visuals.html')
    const output = path.join(directory, 'visuals.pdf')
    const config = path.join(directory, 'render.json')
    fs.writeFileSync(input, printDocument.html, 'utf8')
    const profile = path.join(directory, 'profile')
    fs.writeFileSync(config, JSON.stringify({ input, output, profile, landscape: printDocument.options.orientation === 'landscape', probes }))
    const env = { ...process.env }
    delete env.ELECTRON_RUN_AS_NODE
    const child = spawn(require('electron'), [path.join(__dirname, 'qa-print-render.cjs'), config, `--user-data-dir=${profile}`], { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
    let log = ''
    child.stdout.on('data', (chunk) => { log += chunk })
    child.stderr.on('data', (chunk) => { log += chunk })
    const code = await new Promise((resolve) => {
      const timer = setTimeout(() => { killTree(child); resolve('timeout') }, timeoutMs)
      child.on('error', (error) => { clearTimeout(timer); log += String(error); resolve('error') })
      child.on('exit', (exitCode) => { clearTimeout(timer); resolve(exitCode) })
    })
    killTree(child)
    if (code !== 0) return { failed: code, log }
    return { pdf: fs.readFileSync(output), report: JSON.parse(fs.readFileSync(`${output}.json`, 'utf8')) }
  } finally {
    // Chromium's helper processes may hold the profile for a moment after the main process exits.
    try { fs.rmSync(directory, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 }) } catch {}
  }
}

async function visualExports() {
  const lib = require('./qa-print-visuals.cjs').loadVisualLibrary()
  const input = lib.visualPrintInput()
  const colors = lib.EXPECTED_COLORS
  const html = await createSpreadsheetExport(input, 'html')
  const text = html.bytes.toString('utf8')
  assert.match(cellHtml(text, 'B2'), new RegExp(`background-color:${colors.scaleLow}`), 'HTML: colour scale')
  assert.match(cellHtml(text, 'C6'), /background-image:linear-gradient\(90deg, #638EC6/, 'HTML: data bar')
  assert.match(cellHtml(text, 'D6'), /print-cf-icon[\s\S]*fill="#2F9E44"/, 'HTML: icon set')
  assert.match(cellHtml(text, 'A1'), /background-color:#4472C4/, 'HTML: table header')
  assert.match(cellHtml(text, 'A2'), /background-color:#dae3f3/, 'HTML: banded row')
  assert.match(cellHtml(text, 'G2'), /aria-label="Checked"/, 'HTML: checkbox')
  assert.match(cellHtml(text, 'E3'), /<polyline[^>]*stroke="#376092"/, 'HTML: sparkline group')
  assert.match(cellHtml(text, 'F2'), /<polyline[^>]*stroke="#1a73e8"/, 'HTML: SPARKLINE()')
  assert.match(text, /print-color-adjust: exact/, 'HTML keeps its colours when printed from a browser')
  assert.doesNotMatch(text, /\uE000sparkline/)

  const pdf = await createSpreadsheetExport(input, 'pdf')
  assert.equal(pdf.printDocument.html, text, 'PDF renders the same document as the HTML export')
  const probes = [
    { name: 'colour scale low', selector: 'td[data-address="B2"]', colors: [colors.scaleLow] },
    { name: 'colour scale high', selector: 'td[data-address="B6"]', colors: [colors.scaleHigh] },
    { name: 'data bar', selector: 'td[data-address="C6"]', colors: [colors.dataBar] },
    { name: 'icon red', selector: 'td[data-address="D2"] .print-cf-icon', colors: [colors.iconRed] },
    { name: 'icon green', selector: 'td[data-address="D6"] .print-cf-icon', colors: [colors.iconGreen] },
    { name: 'table header', selector: 'td[data-address="A1"]', colors: [colors.tableHeader] },
    { name: 'band', selector: 'td[data-address="A2"]', colors: ['#DAE3F3'] },
    { name: 'checkbox', selector: 'td[data-address="G2"] .print-checkbox', colors: [colors.checkbox] },
    { name: 'sparkline formula', selector: 'td[data-address="F2"] .print-sparkline', colors: [colors.sparklineFormula], tolerance: 24 },
    { name: 'sparkline group', selector: 'td[data-address="E4"] .print-sparkline', colors: [colors.sparklineGroup], tolerance: 24 },
    { name: 'negative red', selector: 'td[data-address="I2"]', colors: [colors.negativeRed], tolerance: 60 },
  ]
  let rendered = await renderInElectron(pdf.printDocument, probes)
  if (rendered.failed === 'timeout') rendered = await renderInElectron(pdf.printDocument, probes, 150_000)
  if (rendered.failed !== undefined) throw new Error(`Electron PDF render failed (${rendered.failed}): ${rendered.log}`)
  for (const probe of probes) {
    const result = rendered.report.probes[probe.name]
    assert.ok(result && result.rect, `${probe.name}: element rendered`)
    for (const color of probe.colors) assert.equal(result.found[color], true, `${probe.name}: ${color} is painted`)
  }
  const content = pdfContent(rendered.pdf)
  assert.equal(rendered.pdf.subarray(0, 5).toString('latin1'), '%PDF-')
  assert.equal((content.match(/\/Type\s*\/Page\b(?!s)/g) || []).length, 1, 'one PDF page, as in the preview')
  const painted = pdfColors(content)
  for (const [name, hex] of Object.entries({
    'colour scale low': colors.scaleLow, 'colour scale high': colors.scaleHigh, 'table header': colors.tableHeader, band: '#DAE3F3',
    'icon red': colors.iconRed, 'icon green': colors.iconGreen, checkbox: colors.checkbox,
    'sparkline group': colors.sparklineGroup, 'sparkline formula': colors.sparklineFormula, 'negative red': colors.negativeRed,
  })) assert.ok(hasPdfColor(painted, hex), `PDF paints ${name} (${hex})`)
  assert.match(content, /\/ShadingType\s+\d/, 'PDF draws the gradient data bars as shadings')
  return 'HTML and a real Electron PDF'
}

/** CALC-SIE-14: XLS (values only, after consent) and ODS (basic formatting) without the engine. */
async function nativeExports() {
  const workbook = {
    version: 1, name: 'Plan.xlsx', activeSheetId: 'sheet-1', metadata: {},
    definedNames: [{ name: 'Rate', ranges: ['Data!$B$2'] }],
    sheets: [
      {
        id: 'sheet-1', name: 'Data', state: 'visible', rowCount: 4, colCount: 3, merges: ['A4:B4'], colWidths: { 1: 20 }, rowHeights: {},
        frozen: { rows: 1, columns: 0 },
        cells: {
          A1: { value: 'When', style: { font: { bold: true } } }, B1: { value: 'Amount' }, C1: { value: 'Total' },
          A2: { value: 45366, numFmt: 'yyyy-mm-dd' }, B2: { value: 12.5, numFmt: '#,##0.00' }, C2: { formula: 'B2*2', result: 25 },
          A3: { value: 45367, numFmt: 'yyyy-mm-dd' }, B3: { value: 7 }, C3: { formula: 'SUM(B2:B3)', result: 19.5 },
          A4: { value: 'merged note' },
        },
      },
      { id: 'sheet-2', name: 'Audit', state: 'hidden', rowCount: 1, colCount: 1, merges: [], cells: { A1: { value: 'secret' } } },
    ],
  }
  const input = { workbook }
  // ODS: basic formatting, formulas kept, hidden sheet stays hidden, dates keep their format.
  const warnings = []
  const ods = await createSpreadsheetExport(input, 'ods', { officeEngine: false, warnings })
  const zip = await JSZip.loadAsync(ods.bytes)
  assert.equal(Object.keys(zip.files)[0], 'mimetype', 'ODF mimetype comes first')
  assert.equal(await zip.file('mimetype').async('string'), 'application/vnd.oasis.opendocument.spreadsheet')
  const sheetJs = XLSX.read(ods.bytes, { type: 'buffer', cellFormula: true, cellNF: true })
  assert.deepEqual(sheetJs.SheetNames, ['Data', 'Audit'])
  assert.equal(sheetJs.Sheets.Data.C2.f, 'B2*2')
  assert.match(String(sheetJs.Sheets.Data.A2.z), /^yyyy"?-"?mm"?-"?dd$/, `date format kept: ${sheetJs.Sheets.Data.A2.z}`)
  const reopened = await workbookPayloadFromBytes('export.ods', ods.bytes, { officeEngine: false })
  assert.equal(reopened.workbook.sheets[1].state, 'hidden', 'a hidden sheet is not revealed by the export')
  assert.equal(reopened.workbook.sheets[0].cells.C3.formula, 'SUM(B2:B3)')
  assert.equal(reopened.workbook.sheets[0].cells.A2.value, 45366)
  assert.ok(warnings.some((warning) => /fonts, fills/.test(warning)), 'the formatting loss is reported')
  assert.ok(exportLosses(workbook, 'ods', { officeEngine: false }).includes('frozen panes'))
  assert.deepEqual(exportLosses(workbook, 'ods', { officeEngine: true }), [], 'nothing is lost through the engine')
  // XLS: refused without consent (the loss list names formulas, styles and names), then values only.
  const losses = exportLosses(workbook, 'xls', { officeEngine: false })
  assert.ok(losses.some((loss) => /2 formulas/.test(loss)))
  assert.ok(losses.includes('named ranges'))
  await assert.rejects(() => createSpreadsheetExport(input, 'xls', { officeEngine: false }), (error) => error.code === 'LOSSY_CONFIRM_REQUIRED' && Array.isArray(error.losses) && error.losses.length > 0)
  const xls = await createSpreadsheetExport(input, 'xls', { officeEngine: false, valuesOnly: true })
  const legacy = XLSX.read(xls.bytes, { type: 'buffer', cellFormula: true, cellNF: true })
  assert.deepEqual(legacy.SheetNames, ['Data', 'Audit'])
  assert.equal(legacy.Workbook.Sheets[1].Hidden, 1)
  for (const [address, cell] of Object.entries(workbook.sheets[0].cells)) {
    const expected = cell.formula ? cell.result : cell.value
    assert.equal(legacy.Sheets.Data[address].v, expected, `XLS keeps the value of ${address}`)
  }
  assert.equal(legacy.Sheets.Data.C2.f, undefined, 'values only')
  const reopenedXls = await workbookPayloadFromBytes('export.xls', xls.bytes, { officeEngine: false })
  assert.equal(reopenedXls.workbook.sheets[0].cells.B2.value, 12.5)
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
