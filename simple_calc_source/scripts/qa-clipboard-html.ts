import assert from 'node:assert/strict'
import {
  ClipboardTooLargeError,
  a1ToR1C1,
  parseClipboardHtml,
  parseExcelXmlSpreadsheet,
  parseClipboardPayload,
  parseDelimitedText,
  parsePastedText,
  parsePastedValue,
  r1c1ToA1,
  serializeSelectionToClipboard,
  shiftFormulaA1,
  transposeFormula,
} from '../src/lib/clipboard-html'
import { applyPasteSpecial, PASTE_SPECIAL_MENU, PASTE_SPECIAL_PRESETS, type PasteSpecialInput } from '../src/lib/paste-special'
import { shiftFormulaReferences } from '../src/lib/formulas'
import type { CellData } from '../src/spreadsheet-types'
import { EXCEL_365, EXCEL_XML_SPREADSHEET, GOOGLE_SHEETS, LIBREOFFICE, WEB_PAGE, WORD } from './fixtures/clipboard-html-fixtures'

const builtin = { parser: 'builtin' as const }
let checks = 0
function test(name: string, run: () => void) {
  try {
    run()
    checks += 1
  } catch (error) {
    console.error(`FAILED: ${name}`)
    throw error
  }
}

// ---------------------------------------------------------------------------------------------
// Excel
// ---------------------------------------------------------------------------------------------

test('Excel 365 clipboard HTML', () => {
  const parsed = parseClipboardHtml(EXCEL_365, builtin)
  assert.ok(parsed)
  assert.equal(parsed.source, 'excel')
  assert.equal(parsed.cells.length, 4)
  assert.equal(parsed.cells[0].length, 4)
  const [row1, row2, row3, row4] = parsed.cells

  const header = row1[0]
  assert.equal(header.value, 'Name')
  assert.deepEqual(header.style?.font, { bold: true }, 'Calibri 11 black is the default font and is not stored')
  assert.deepEqual(header.style?.fill, { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } })
  assert.deepEqual(header.style?.alignment, { horizontal: 'center', vertical: 'middle' })
  for (const side of ['top', 'right', 'bottom', 'left']) assert.deepEqual(header.style?.border?.[side], { style: 'thin', color: { argb: 'FF000000' } })
  assert.equal(row1[1].style?.border?.left, undefined, 'inline border-left:none overrides the class border')
  assert.ok(row1[1].style?.border?.top)

  assert.equal(row2[0].value, 'Alpha\nBeta')
  assert.equal(row2[0].style?.alignment?.wrapText, true)
  assert.equal(row2[1].value, 1234.5)
  assert.equal(row2[1].numFmt, '"$"#,##0.00', 'mso-number-format CSS escapes are decoded')
  assert.equal(row2[1].style, undefined, 'align=right on an Excel number is general alignment, not explicit')
  assert.equal(row2[2].value, 0.125, 'x:num carries the precise value behind the rounded display')
  assert.equal(row2[2].numFmt, '0.00%')
  assert.equal(row2[3].value, 45306)
  assert.equal(row2[3].numFmt, 'm/d/yyyy')
  assert.equal(row2[3].type, 'date')

  assert.equal(row3[0].value, '007', 'text number format keeps leading zeros')
  assert.equal(row3[0].numFmt, '@')
  assert.deepEqual(row3[0].style?.font, { name: 'Times New Roman', size: 14, italic: true, underline: true, color: { argb: 'FFFF0000' } })
  assert.equal(row3[1].value, -42.5)
  assert.equal(row3[1].numFmt, '0.00_);[Red]\\(0.00\\)')
  assert.deepEqual(row3[1].style?.border, { bottom: { style: 'double', color: { argb: 'FF000000' } } })
  assert.equal(row3[2].value, 'Mixed bold blue')
  assert.deepEqual(row3[2].richText, [{ text: 'Mixed ' }, { text: 'bold blue', font: { bold: true, color: { argb: 'FF0070C0' } } }])

  assert.equal(row4[0].value, true)
  assert.equal(row4[1].formula, 'B2*2')
  assert.equal(row4[1].result, 2469)
  assert.equal(row4[2].value, 'Merged total')
  assert.deepEqual(parsed.merges, ['C4:D4'], 'mso-ignore:colspan overflow is not a merge')
  assert.deepEqual(parsed.columnWidths, [64, 128, 64, 64])
  assert.deepEqual(parsed.rowHeights, [20, 40, 20, 20])
})

test('Excel XML Spreadsheet flavour (formulas)', () => {
  const parsed = parseExcelXmlSpreadsheet(EXCEL_XML_SPREADSHEET, { origin: { row: 9, col: 1 } })
  assert.ok(parsed)
  assert.equal(parsed.source, 'excel')
  assert.equal(parsed.sourceSheetName, 'Budget 2024')
  assert.deepEqual(parsed.merges, ['A1:B1'])
  assert.deepEqual(parsed.columnWidths, [64, 128, 128, undefined])
  assert.deepEqual(parsed.rowHeights, [undefined, 40, undefined, undefined])
  const [row1, row2, row3, row4] = parsed.cells
  assert.equal(row1[0].value, 'Header')
  assert.deepEqual(row1[0].style, {
    font: { bold: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
    border: {
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
      top: { style: 'medium', color: { argb: 'FFFF0000' } },
    },
    alignment: { horizontal: 'center', vertical: 'middle' },
  })
  assert.deepEqual(row1[1], {})
  assert.deepEqual(row1[2], { value: 1234.5, numFmt: '"$"#,##0.00' })
  assert.deepEqual(row1[3], { value: 45306, numFmt: 'm/d/yyyy', type: 'date' })
  assert.deepEqual(row2[0], { value: 'Line 1\nLine 2', style: { alignment: { wrapText: true } } })
  assert.deepEqual(row2[1], { formula: 'D10*2', result: 2469 }, 'R1C1 resolved at the paste position (C11)')
  assert.equal(row2[2].formula, 'SUM($C$1:D10)')
  assert.equal(row2[2].numFmt, '"$"#,##0.00', 'ss:Parent style inheritance')
  assert.deepEqual(row2[2].style?.font, { name: 'Times New Roman', size: 12, italic: true, underline: 'double', color: { argb: 'FF0070C0' } })
  assert.deepEqual(row2[3], { formula: '1/0', result: '#DIV/0!', resultType: 'error' })
  assert.equal(row3[0].value, true)
  assert.equal(row3[1].hyperlink, 'https://example.com/')
  assert.equal(row3[1].note, 'Ana:\nCheck this')
  assert.deepEqual(row3[2], {})
  assert.deepEqual(row3[3].richText, [{ text: 'Bold', font: { bold: true } }, { text: ' red', font: { color: { argb: 'FFFF0000' } } }])
  assert.equal(row4[0].value, 0.5625, 'time-only DateTime (1899-12-31) is a fraction')
  const preferred = parseClipboardPayload({ text: 'x', html: EXCEL_365, excelXml: EXCEL_XML_SPREADSHEET }, { parser: 'builtin' })
  assert.equal(preferred?.cells[1][1].formula, 'C1*2', 'the XML flavour wins over HTML because it has formulas')
  assert.equal(parseClipboardPayload({ text: 'x', html: EXCEL_365, excelXml: '<Workbook><broken' }, { parser: 'builtin' })?.cells[0][0].value, 'Name', 'broken XML falls back to HTML')
  assert.throws(() => parseExcelXmlSpreadsheet(EXCEL_XML_SPREADSHEET, { maxCells: 10 }), ClipboardTooLargeError)
})

// ---------------------------------------------------------------------------------------------
// Google Sheets
// ---------------------------------------------------------------------------------------------

test('Google Sheets clipboard HTML', () => {
  const parsed = parseClipboardHtml(GOOGLE_SHEETS, { ...builtin, origin: { row: 4, col: 3 } })
  assert.ok(parsed)
  assert.equal(parsed.source, 'google-sheets')
  assert.deepEqual(parsed.formulaOrigin, { row: 4, col: 3 })
  const [row1, row2, row3, row4] = parsed.cells
  assert.equal(row1[0].value, 'Item')
  assert.deepEqual(row1[0].style?.font, { name: 'Arial', size: 10, bold: true }, 'table-level Arial 10pt is inherited')
  assert.deepEqual(row1[0].style?.fill, { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAD3' } })
  assert.equal(row1[0].style?.border, undefined, 'Sheets gridline borders from the td rule are ignored')
  assert.equal(row1[1].style?.alignment?.horizontal, 'center')

  assert.equal(row2[0].value, 'Widget')
  assert.equal(row2[0].note, 'Check supplier')
  assert.deepEqual(row2[0].style?.font, { name: 'Arial', size: 10, italic: true, color: { argb: 'FFFF0000' } })
  assert.equal(row2[1].value, 19.99)
  assert.equal(row2[1].numFmt, '"$"#,##0.00')
  assert.deepEqual(row2[1].style?.border, { bottom: { style: 'medium', color: { argb: 'FF000000' } } })
  assert.equal(row2[2].formula, 'E6*2', 'R1C1 formula is resolved at the paste destination (F6)')
  assert.equal(row2[2].result, 39.98)

  assert.equal(row3[0].value, 'Merged')
  assert.equal(row3[0].style?.alignment?.vertical, 'middle')
  assert.equal(row3[1].value, 45306)
  assert.equal(row3[1].numFmt, 'yyyy-mm-dd')
  assert.equal(row3[1].type, 'date')
  assert.equal(row3[2].value, true)

  assert.deepEqual(row4[0], {}, 'covered merge cell')
  assert.equal(row4[1].formula, 'SUM(F6:F7)/$C$2')
  assert.equal(row4[1].result, 0.25)
  assert.equal(row4[1].numFmt, '0%')
  assert.equal(row4[2].value, 'Docs')
  assert.equal(row4[2].hyperlink, 'https://example.com/docs')
  assert.deepEqual(parsed.merges, ['A3:A4'])
  assert.deepEqual(parsed.columnWidths, [100, 120, 100])
  assert.deepEqual(parsed.rowHeights, [21, 21, 21, 21])
})

// ---------------------------------------------------------------------------------------------
// LibreOffice Calc
// ---------------------------------------------------------------------------------------------

test('LibreOffice Calc clipboard HTML', () => {
  const parsed = parseClipboardHtml(LIBREOFFICE, builtin)
  assert.ok(parsed)
  assert.equal(parsed.source, 'libreoffice')
  const [row1, row2, row3, row4] = parsed.cells
  const lo = { name: 'Liberation Sans', size: 10 }
  assert.equal(row1[0].value, 'Region')
  assert.deepEqual(row1[0].style?.font, { ...lo, bold: true })
  assert.deepEqual(row1[0].style?.fill?.fgColor, { argb: 'FFFFFF00' })
  assert.equal(row1[0].style?.alignment, undefined, 'ALIGN=LEFT on text is general alignment')
  assert.deepEqual(row1[1].style?.font, { ...lo, bold: true, italic: true })
  assert.equal(row1[1].style?.alignment?.horizontal, 'center')
  assert.deepEqual(row1[2].style?.font, { ...lo, color: { argb: 'FFC9211E' } })

  assert.equal(row2[1].value, 1234.5)
  assert.equal(row2[1].numFmt, '[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00')
  assert.equal(row2[2].value, 'Line one\nline two')
  assert.equal(row2[2].style?.alignment?.wrapText, true)
  assert.equal(row3[1].value, 45306)
  assert.equal(row3[1].type, 'date')
  assert.equal(row3[1].numFmt, 'DD/MM/YYYY')
  assert.equal(row3[2].value, true)
  assert.equal(row4[0].value, 'Wide')
  assert.deepEqual(row4[0].style?.font, { name: 'DejaVu Serif', size: 14 })
  assert.deepEqual(row4[0].style?.alignment, { horizontal: 'center', vertical: 'middle' })
  assert.equal(row4[2].value, 0.5)
  assert.equal(row4[2].numFmt, '0%')
  assert.deepEqual(parsed.merges, ['A4:B4'])
  assert.deepEqual(parsed.columnWidths, [86, 86, 113])
})

// ---------------------------------------------------------------------------------------------
// Generic web page (hostile markup)
// ---------------------------------------------------------------------------------------------

test('generic web table', () => {
  const parsed = parseClipboardHtml(WEB_PAGE, builtin)
  assert.ok(parsed)
  assert.equal(parsed.source, 'html')
  const serialized = JSON.stringify(parsed)
  for (const needle of ['evil', 'alert', 'cookie', 'javascript', 'x.png', 'steal']) assert.ok(!serialized.includes(needle), `${needle} must not survive parsing`)
  assert.equal(parsed.cells.length, 6)
  assert.equal(parsed.cells[0][0].value, 'Quarterly report', 'text around the table becomes its own rows')
  assert.equal(parsed.cells[5][0].value, 'Source: internal')
  const [, header, first, second, third] = parsed.cells
  assert.deepEqual(header.map((cell) => cell.value), ['Product', 'Revenue', 'Growth', 'Launch', 'Link'])
  assert.deepEqual(header[0].style?.font, { bold: true }, 'th is bold by default')

  assert.equal(first[0].value, 'Café & Bar')
  assert.deepEqual(first[0].style?.font, { color: { argb: 'FF008000' } }, 'stylesheet td rules apply')
  assert.equal(first[1].value, 12345.67)
  assert.equal(first[1].numFmt, '"$"#,##0.00')
  assert.equal(first[2].value, 0.125)
  assert.equal(first[2].numFmt, '0.0%')
  assert.equal(first[3].value, 45356)
  assert.equal(first[3].numFmt, 'mmm d, yyyy')
  assert.equal(first[4].value, 'bad')
  assert.equal(first[4].hyperlink, undefined, 'javascript: links are dropped')

  assert.equal(second[0].value, 'Hot')
  assert.deepEqual(second[0].style?.fill?.fgColor, { argb: 'FFFF0000' })
  assert.deepEqual(second[0].style?.font, { bold: true, color: { argb: 'FFFFFFFF' } })
  assert.equal(second[1].value, -1500)
  assert.equal(second[1].numFmt, '#,##0;(#,##0)')
  assert.equal(second[2].value, -0.03)
  assert.equal(second[3].value, 45351)
  assert.equal(second[4].hyperlink, 'https://ok.example/x?a=1&b=2')

  assert.equal(third[0].value, undefined, '&nbsp; alone is an empty cell')
  assert.equal(third[1].value, '007')
  assert.equal(third[2].value, 1500)
  assert.ok(Math.abs((third[3].value as number) - (13 * 60 + 45) / 1440) < 1e-12)
  assert.equal(third[3].numFmt, 'h:mm')
  assert.equal(third[4].style?.alignment?.horizontal, 'center')
})

test('Word table', () => {
  const parsed = parseClipboardHtml(WORD, builtin)
  assert.ok(parsed)
  assert.equal(parsed.source, 'word')
  assert.equal(parsed.cells.length, 1)
  const [header, number] = parsed.cells[0]
  assert.equal(header.value, 'Header')
  assert.deepEqual(header.style?.font, { name: 'Arial', size: 9, bold: true })
  assert.equal(header.style?.alignment?.vertical, 'top')
  assert.equal(header.style?.border?.left?.style, 'medium')
  assert.equal(number.value, 42)
  assert.deepEqual(number.style?.fill?.fgColor, { argb: 'FFD9E2F3' })
  assert.equal(number.style?.border?.left, undefined)
  assert.equal(number.style?.border?.right?.style, 'medium')
})

test('tolerant parsing', () => {
  const loose = parseClipboardHtml('<table><tr><td>1<td>2<tr><td>3</table>', builtin)
  assert.deepEqual(loose?.cells.map((row) => row.map((cell) => cell.value)), [[1, 2], [3, undefined]])
  const nested = parseClipboardHtml('<table><tr><td>Outer<table><tr><td>in1<td>in2</table></td><td>B</td></tr></table>', builtin)
  assert.deepEqual(nested?.cells.map((row) => row.map((cell) => cell.value)), [['Outer\nin1 in2', 'B']])
  const entities = parseClipboardHtml('<table><tr><td>&lt;b&gt;</td><td>&#8364;5</td><td>&amp;copy</td></tr></table>', builtin)
  assert.deepEqual(entities?.cells[0].map((cell) => cell.value), ['<b>', 5, '&copy'])
  assert.equal(entities?.cells[0][1].numFmt, '"€"#,##0')
  assert.equal(parseClipboardHtml('<p>No table here</p>', builtin), null)
  assert.equal(parseClipboardHtml('', builtin), null)
  const imported = parseClipboardHtml('<style>@import url(x.css); @charset "utf-8"; td.k { font-weight: bold }</style><table><tr><td class="k">x</td></tr></table>', builtin)
  assert.deepEqual(imported?.cells[0][0].style?.font, { bold: true }, '@import does not swallow the next rule')
  const spans = parseClipboardHtml('<table><tr><td rowspan=0 colspan=9999>x</td></tr></table>', builtin)
  assert.equal(spans?.cells[0].length, 1000, 'colspan is clamped')
  assert.throws(() => parseClipboardHtml(`<table>${'<tr><td>1</td></tr>'.repeat(10)}</table>`, { ...builtin, maxCells: 5 }), ClipboardTooLargeError)
})

// ---------------------------------------------------------------------------------------------
// Plain text
// ---------------------------------------------------------------------------------------------

test('plain text parsing', () => {
  assert.deepEqual(parseDelimitedText('a\tb\r\nc\td\r\n'), [['a', 'b'], ['c', 'd']])
  assert.deepEqual(parseDelimitedText('"line1\nline2"\tx\r\n'), [['line1\nline2', 'x']])
  assert.deepEqual(parseDelimitedText('5" pipe\t"quoted ""x"""\n'), [['5" pipe', 'quoted "x"']])
  assert.deepEqual(parseDelimitedText('"Hello" she said'), [['"Hello" she said']])
  assert.deepEqual(parseDelimitedText('a\t\n\tb'), [['a', ''], ['', 'b']])
  assert.deepEqual(parseDelimitedText('x\t'), [['x', '']])
  const values = (text: string) => parsePastedValue(text)
  assert.deepEqual(values('1,234.50'), { value: 1234.5, numFmt: '#,##0.00' })
  assert.deepEqual(values('$5'), { value: 5, numFmt: '"$"#,##0' })
  assert.deepEqual(values('-$1,234.5'), { value: -1234.5, numFmt: '"$"#,##0.00' })
  assert.deepEqual(values('12%'), { value: 0.12, numFmt: '0%' })
  assert.deepEqual(values('1/15/2024'), { value: 45306, numFmt: 'm/d/yyyy', type: 'date' })
  assert.deepEqual(values('2024-01-15 13:30'), { value: 45306 + 13.5 / 24, numFmt: 'm/d/yyyy h:mm', type: 'date' })
  assert.deepEqual(values('TRUE'), { value: true })
  assert.deepEqual(values("'007"), { value: '007' })
  assert.deepEqual(values('007'), { value: '007' })
  assert.deepEqual(values('=SUM(A1:A2)'), { formula: 'SUM(A1:A2)' })
  assert.deepEqual(values('(1,234)'), { value: -1234, numFmt: '#,##0;(#,##0)' })
  assert.deepEqual(values(' 42 '), { value: 42 })
  assert.deepEqual(values('1.5e3'), { value: 1500 })
  assert.deepEqual(values('12 345'), { value: '12 345' })
  assert.deepEqual(values('1,5'), { value: '1,5' })
  assert.deepEqual(values('2/30/2024'), { value: '2/30/2024' }, 'impossible dates stay text')
  assert.deepEqual(values(''), {})
  assert.deepEqual(parsePastedValue('1.234,50 €', { decimalSeparator: ',' }), { value: 1234.5, numFmt: '#,##0.00 "€"' })
  assert.deepEqual(parsePastedValue('15/01/2024', { dayFirst: true }), { value: 45306, numFmt: 'd/m/yyyy', type: 'date' })

  const matrix = parsePastedText('Name\tAmount\nA\t1\nB')
  assert.deepEqual(matrix.map((row) => row.map((cell) => cell.value)), [['Name', 'Amount'], ['A', 1], ['B', undefined]])
  assert.deepEqual(parsePastedText('a\nb\n').map((row) => row.map((cell) => cell.value)), [['a'], ['b']])
  assert.deepEqual(parsePastedText(''), [])
  assert.throws(() => parsePastedText('1\n'.repeat(100_001)), ClipboardTooLargeError)
  const payload = parseClipboardPayload({ text: 'x\ty', html: '<b>no table</b>' })
  assert.equal(payload?.source, 'text')
  assert.deepEqual(payload?.cells[0].map((cell) => cell.value), ['x', 'y'])
})

// ---------------------------------------------------------------------------------------------
// R1C1 / A1
// ---------------------------------------------------------------------------------------------

test('R1C1 and A1 formula conversion', () => {
  assert.equal(a1ToR1C1('=A1+$B$2+C$3+$D4', 4, 4), '=R[-4]C[-4]+R2C2+R3C[-2]+R[-1]C4')
  assert.equal(a1ToR1C1('=SUM(A1:B2)', 0, 0), '=SUM(R[0]C[0]:R[1]C[1])')
  assert.equal(a1ToR1C1('=SUM(C:C)+SUM(2:3)', 0, 0), '=SUM(C[2]:C[2])+SUM(R[1]:R[2])')
  assert.equal(a1ToR1C1('="A1"&A1', 0, 0), '="A1"&R[0]C[0]')
  assert.equal(a1ToR1C1('=LOG10(A1)+ATAN2(1,2)+1.5E3', 0, 0), '=LOG10(R[0]C[0])+ATAN2(1,2)+1.5E3')
  assert.equal(a1ToR1C1("='My Sheet'!A1+Sheet2!B2", 0, 0), "='My Sheet'!R[0]C[0]+Sheet2!R[1]C[1]")
  assert.equal(a1ToR1C1('=SUM(Table1[Amount])+A1', 0, 0), '=SUM(Table1[Amount])+R[0]C[0]')
  assert.equal(a1ToR1C1('=_xlfn.XLOOKUP(A1,B:B,C:C)', 0, 0), '=_xlfn.XLOOKUP(R[0]C[0],C[1]:C[1],C[2]:C[2])')

  const samples = ['=A1+$B$2+C$3+$D4', '=SUM(A1:B2)*2', "='My Sheet'!C3", '=IF(A1>0,"A1",B$9)', '=SUM(C:C)', '=SUM(2:3)', '=ROUND(E5,2)']
  for (const formula of samples) assert.equal(r1c1ToA1(a1ToR1C1(formula, 6, 5), 6, 5), formula, `round trip ${formula}`)
  assert.equal(r1c1ToA1('=RC[-1]*2', 0, 1), '=A1*2', 'Excel-style R1C1 without brackets')
  assert.equal(r1c1ToA1('=R[0]C[-1]', 0, 0), '=#REF!')
  assert.equal(r1c1ToA1('=ROUND(R[0]C[1],2)+COUNT(C[1]:C[1])+ROW()+COLUMN()', 0, 0), '=ROUND(B1,2)+COUNT(B:B)+ROW()+COLUMN()')

  for (const formula of ['A1+B2', '$A$1+A$1+$A1', 'SUM(A1:B3)', "'Other sheet'!B2*C3", 'SUM(A:A)', 'IF(A1="B2",C3,D4)']) {
    assert.equal(shiftFormulaA1(formula, 2, 3), shiftFormulaReferences(formula, 2, 3), `shift ${formula}`)
  }
  assert.equal(transposeFormula('SUM(A1:A3)', 3, 0, 0, 3), 'SUM(A1:C1)')
  assert.equal(transposeFormula('A1+B1', 0, 2, 6, 4), 'E5+E6')
  assert.equal(transposeFormula('$A$1*2', 5, 5, 9, 9), '$A$1*2')
})

// ---------------------------------------------------------------------------------------------
// Serialize and round trip
// ---------------------------------------------------------------------------------------------

const SOURCE: CellData[][] = [
  [
    { value: 'Title <script>', style: { font: { bold: true, size: 14, color: { argb: 'FF1F4E79' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } }, alignment: { horizontal: 'center', vertical: 'middle' } } },
    {},
    { value: 0.256, numFmt: '0.0%' },
  ],
  [
    { value: 1234.5, numFmt: '"$"#,##0.00', style: { border: { bottom: { style: 'thin', color: { argb: 'FF000000' } }, top: { style: 'medium', color: { argb: 'FFFF0000' } } } } },
    { formula: 'A2*2', result: 2469, numFmt: '#,##0' },
    { value: 'two\nlines', style: { alignment: { wrapText: true } }, note: 'Remember this' },
  ],
  [
    { value: true },
    { value: 'Docs', hyperlink: 'https://example.com', richText: [{ text: 'Do', font: { bold: true } }, { text: 'cs', font: { italic: true } }] },
    { value: 45306, numFmt: 'm/d/yyyy', type: 'date', style: { font: { name: 'Georgia', underline: 'double', strike: true } } },
  ],
]

test('serialize to TSV + HTML', () => {
  const displays = [['Title <script>', '', '25.6%'], ['$1,234.50', '2,469', 'two\nlines'], ['TRUE', 'Docs', '1/15/2024']]
  const { text, html } = serializeSelectionToClipboard(SOURCE, {
    displayAt: (row, col) => displays[row][col],
    merges: ['A1:B1'],
    origin: { row: 1, col: 1 },
    sheetName: 'Data Sheet',
    columnWidthsPx: [120, 80, 90],
    rowHeightsPx: [30, 20, 20],
  })
  assert.equal(text, 'Title <script>\t\t25.6%\r\n$1,234.50\t2,469\t"two\nlines"\r\nTRUE\tDocs\t1/15/2024')
  assert.ok(!html.includes('<script>'), 'values are escaped')
  assert.ok(html.includes('Title &lt;script&gt;'))
  assert.ok(html.includes('colspan="2"'))
  assert.ok(html.includes('<col width="120" style="width:120px">'))
  assert.ok(html.includes('x:num="1234.5"'))
  assert.ok(html.includes(`mso-number-format:'&quot;$&quot;#,##0.00'`))
  assert.ok(html.includes('x:fmla="=B3*2"') === false && html.includes('x:fmla="=A2*2"'))
  assert.ok(html.includes('data-sheets-formula="=R[-1]C[-2]*2"'), 'Sheets formula is relative to the source position (C3)')
  assert.ok(html.includes('x:bool="TRUE"'))
  assert.ok(html.includes('border-top:1pt solid #FF0000'))
  assert.ok(html.includes('two<br>lines'))
  assert.ok(html.includes('<a href="https://example.com">'))
  assert.ok(html.includes('data-sheets-note="Remember this"'))
  assert.ok(html.includes('data-sc-origin="B2"') && html.includes('data-sc-sheet="Data Sheet"'))
})

test('round trip through simple_calc HTML (lossless internal data)', () => {
  const { html } = serializeSelectionToClipboard(SOURCE, { merges: ['A1:B1'], origin: { row: 1, col: 1 }, sheetName: 'Data Sheet', columnWidthsPx: [120, 80, 90] })
  const parsed = parseClipboardHtml(html, builtin)
  assert.ok(parsed)
  assert.equal(parsed.source, 'simple-calc')
  assert.deepEqual(parsed.formulaOrigin, { row: 1, col: 1 })
  assert.equal(parsed.sourceSheetName, 'Data Sheet')
  assert.deepEqual(parsed.merges, ['A1:B1'])
  assert.deepEqual(parsed.columnWidths, [120, 80, 90])
  const expected = SOURCE.map((row) => row.map((cell) => ({ ...cell })))
  expected[0][1] = {}
  assert.deepEqual(parsed.cells, expected)
})

test('round trip through the portable HTML (what other apps see)', () => {
  const { html } = serializeSelectionToClipboard(SOURCE, { merges: ['A1:B1'], origin: { row: 1, col: 1 }, includeInternalData: false })
  const parsed = parseClipboardHtml(html, builtin)
  assert.ok(parsed)
  const [row1, row2, row3] = parsed.cells
  assert.equal(row1[0].value, 'Title <script>')
  assert.deepEqual(row1[0].style, SOURCE[0][0].style)
  assert.equal(row1[2].value, 0.256)
  assert.equal(row1[2].numFmt, '0.0%')
  assert.equal(row2[0].value, 1234.5)
  assert.equal(row2[0].numFmt, '"$"#,##0.00')
  assert.deepEqual(row2[0].style, SOURCE[1][0].style)
  assert.equal(row2[1].formula, 'A2*2', 'formula comes back at the source origin')
  assert.equal(row2[1].result, 2469)
  assert.equal(row2[2].value, 'two\nlines')
  assert.equal(row2[2].note, 'Remember this')
  assert.deepEqual(row2[2].style, { alignment: { wrapText: true } })
  assert.equal(row3[0].value, true)
  assert.equal(row3[1].hyperlink, 'https://example.com')
  assert.deepEqual(row3[1].richText, SOURCE[2][1].richText)
  assert.equal(row3[2].value, 45306)
  assert.equal(row3[2].type, 'date')
  assert.deepEqual(row3[2].style?.font, { name: 'Georgia', underline: 'double', strike: true })
  assert.deepEqual(parsed.merges, ['A1:B1'])
})

// ---------------------------------------------------------------------------------------------
// Paste Special
// ---------------------------------------------------------------------------------------------

function grid(cells: Record<string, CellData>) {
  return (row: number, col: number) => cells[`${String.fromCharCode(65 + col)}${row + 1}`]
}

const PS_SOURCE: CellData[][] = [
  [{ value: 10, numFmt: '0.00', style: { font: { bold: true } } }, { formula: 'A1*2', result: 20 }],
  [{}, { value: 'text', note: 'hello', style: { border: { top: { style: 'thin' } } } }],
]

function paste(options: PasteSpecialInput['options'], destination: PasteSpecialInput['destination'], existing: Record<string, CellData> = {}, extra: Partial<PasteSpecialInput> = {}) {
  return applyPasteSpecial({
    source: { cells: PS_SOURCE, origin: { row: 0, col: 0 }, sheetName: 'Sheet1' },
    destination,
    getDestinationCell: grid(existing),
    options,
    shiftFormula: shiftFormulaReferences,
    ...extra,
  })
}

test('paste special: values / formulas / formats / skip blanks', () => {
  const values = paste({ paste: 'values' }, { top: 0, left: 3 }, { D2: { value: 5 }, E1: { value: 1, style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } } } })
  assert.deepEqual(values.changes.D1, { value: 10 })
  assert.deepEqual(values.changes.E1, { value: 20, style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } } }, 'values keep destination formatting')
  assert.equal(values.changes.D2, null, 'a blank source clears the destination')
  assert.deepEqual(values.selection, { top: 0, left: 3, bottom: 1, right: 4 })

  const skipped = paste({ paste: 'values', skipBlanks: true }, { top: 0, left: 3 }, { D2: { value: 5 } })
  assert.ok(!('D2' in skipped.changes), 'skip blanks leaves the destination alone')

  const formulas = paste({ paste: 'formulas' }, { top: 0, left: 3 })
  assert.deepEqual(formulas.changes.E1, { formula: 'D1*2' })
  assert.deepEqual(formulas.changes.E2, { value: 'text' })

  const formats = paste({ paste: 'formats' }, { top: 0, left: 3 }, { D1: { value: 99 } })
  assert.deepEqual(formats.changes.D1, { value: 99, numFmt: '0.00', style: { font: { bold: true } } })
  assert.deepEqual(formats.merges, undefined)

  const all = paste({ paste: 'all' }, { top: 0, left: 3 }, { D1: { value: 99, note: 'old' } })
  assert.deepEqual(all.changes.D1, { value: 10, numFmt: '0.00', style: { font: { bold: true } } })
  assert.deepEqual(all.changes.E2, { value: 'text', note: 'hello', style: { border: { top: { style: 'thin' } } } })

  const noBorders = paste({ paste: 'allExceptBorders' }, { top: 0, left: 3 }, { E2: { style: { border: { left: { style: 'thick' } } } } })
  assert.deepEqual(noBorders.changes.E2, { value: 'text', note: 'hello', style: { border: { left: { style: 'thick' } } } })

  const notes = paste({ paste: 'comments' }, { top: 0, left: 3 }, { E2: { value: 7 } })
  assert.deepEqual(Object.keys(notes.changes), ['E2'])
  assert.deepEqual(notes.changes.E2, { value: 7, note: 'hello' })

  const valuesAndFormats = paste({ paste: 'valuesAndNumberFormats' }, { top: 0, left: 3 }, { D1: { style: { font: { italic: true } } } })
  assert.deepEqual(valuesAndFormats.changes.D1, { value: 10, numFmt: '0.00', style: { font: { italic: true } } })
})

test('paste special: operations', () => {
  const one = (value: CellData, options: PasteSpecialInput['options'], existing: CellData | undefined) => applyPasteSpecial({
    source: { cells: [[value]], origin: { row: 0, col: 0 } },
    destination: { top: 2, left: 2 },
    getDestinationCell: () => existing,
    options,
    shiftFormula: shiftFormulaReferences,
  }).changes.C3
  assert.deepEqual(one({ value: 10 }, { paste: 'values', operation: 'add' }, { value: 5 }), { value: 15 })
  assert.deepEqual(one({ value: 10 }, { paste: 'values', operation: 'multiply' }, { value: 5, numFmt: '0.0' }), { value: 50, numFmt: '0.0' })
  assert.deepEqual(one({ value: 4 }, { paste: 'values', operation: 'divide' }, { value: 10 }), { value: 2.5 })
  assert.deepEqual(one({ value: 0 }, { paste: 'values', operation: 'divide' }, { value: 10 }), { value: '#DIV/0!' })
  assert.deepEqual(one({ value: 0.2 }, { paste: 'values', operation: 'add' }, { value: 0.1 }), { value: 0.3 })
  assert.deepEqual(one({ value: 10 }, { paste: 'values', operation: 'subtract' }, { formula: 'B9' }), { formula: '(B9)-10' })
  assert.deepEqual(one({ formula: 'A1*2', result: 2 }, { paste: 'formulas', operation: 'add' }, { value: 4 }), { formula: '4+(C3*2)' })
  assert.deepEqual(one({ value: 10 }, { paste: 'values', operation: 'multiply' }, undefined), { value: 0 })
  assert.equal(one({ value: 10 }, { paste: 'values', operation: 'add' }, { value: 'label' }), undefined, 'text destinations are untouched')
  assert.deepEqual(one({ value: 'abc' }, { paste: 'values', operation: 'add' }, { value: 3 }), { value: 'abc' }, 'text sources paste as-is')
  assert.deepEqual(one({ value: 10 }, { paste: 'formats', operation: 'add' }, { value: 3 }), { value: 3 }, 'operations do not apply to formats')
})

test('paste special: transpose, tiling, link, widths, merges, validation', () => {
  const row: CellData[][] = [[{ value: 1 }, { value: 2 }, { formula: 'A1+B1', result: 3 }]]
  const transposed = applyPasteSpecial({
    source: { cells: row, origin: { row: 0, col: 0 }, merges: ['A1:B1'] },
    destination: { top: 4, left: 4 },
    getDestinationCell: () => undefined,
    options: PASTE_SPECIAL_PRESETS.transpose,
    shiftFormula: shiftFormulaReferences,
  })
  assert.deepEqual(transposed.changes, { E5: { value: 1 }, E6: { value: 2 }, E7: { formula: 'E5+E6' } })
  assert.deepEqual(transposed.merges, ['E5:E6'])
  assert.deepEqual(transposed.selection, { top: 4, left: 4, bottom: 6, right: 4 })

  const tiled = applyPasteSpecial({
    source: { cells: [[{ value: 1 }, { formula: 'A1*2' }]], origin: { row: 0, col: 0 } },
    destination: { top: 0, left: 0, bottom: 3, right: 3 },
    getDestinationCell: () => undefined,
    options: { paste: 'all' },
    shiftFormula: shiftFormulaReferences,
  })
  assert.equal(Object.keys(tiled.changes).length, 16)
  assert.deepEqual(tiled.changes.C4, { value: 1 })
  assert.deepEqual(tiled.changes.D4, { formula: 'C4*2' })
  const once = applyPasteSpecial({
    source: { cells: [[{ value: 1 }, { value: 2 }]], origin: { row: 0, col: 0 } },
    destination: { top: 0, left: 0, bottom: 2, right: 2 },
    getDestinationCell: () => undefined,
  })
  assert.deepEqual(Object.keys(once.changes), ['A1', 'B1'], 'non-multiple selections paste once')

  const linked = applyPasteSpecial({
    source: { cells: PS_SOURCE, origin: { row: 1, col: 1 }, sheetName: 'Sheet 1' },
    destination: { top: 0, left: 0 },
    getDestinationCell: grid({ A1: { value: 3, style: { font: { bold: true } } } }),
    options: PASTE_SPECIAL_PRESETS.pasteLink,
    destinationSheetName: 'Sheet2',
  })
  assert.deepEqual(linked.changes.A1, { formula: "'Sheet 1'!B2", style: { font: { bold: true } } })
  assert.deepEqual(linked.changes.B2, { formula: "'Sheet 1'!C3" })
  const sameSheet = applyPasteSpecial({
    source: { cells: PS_SOURCE, origin: { row: 1, col: 1 }, sheetName: 'Sheet1' },
    destination: { top: 0, left: 0 },
    getDestinationCell: () => undefined,
    options: { pasteLink: true },
    destinationSheetName: 'Sheet1',
  })
  assert.deepEqual(sameSheet.changes.A2, { formula: 'B3' })

  const widths = applyPasteSpecial({
    source: { cells: PS_SOURCE, origin: { row: 0, col: 0 }, columnWidths: [12, 20] },
    destination: { top: 0, left: 2, bottom: 0, right: 5 },
    getDestinationCell: () => undefined,
    options: PASTE_SPECIAL_PRESETS.columnWidths,
  })
  assert.deepEqual(widths.changes, {})
  assert.deepEqual(widths.columnWidths, { 2: 12, 3: 20, 4: 12, 5: 20 })

  const listRule = { type: 'list', formulae: ['"Yes,No"'] }
  const validated = applyPasteSpecial({
    source: { cells: [[{ value: 'Yes' }, { value: 'No' }], [{ value: 'x' }, {}]], origin: { row: 0, col: 0 }, validations: [[listRule, listRule], [listRule, undefined]] },
    destination: { top: 0, left: 0, bottom: 3, right: 1 },
    getDestinationCell: () => undefined,
    options: PASTE_SPECIAL_PRESETS.validation,
  })
  assert.deepEqual(validated.changes, {})
  const coverage = (withRule: boolean) => (validated.validations ?? [])
    .filter((item) => (item.validation !== null) === withRule)
    .flatMap((item) => {
      const [start, end = start] = item.range.split(':')
      const cells: string[] = []
      for (let row = Number(start.slice(1)); row <= Number(end.slice(1)); row += 1) {
        for (let col = start.charCodeAt(0); col <= end.charCodeAt(0); col += 1) cells.push(`${String.fromCharCode(col)}${row}`)
      }
      return cells
    }).sort()
  assert.deepEqual(coverage(false), ['B2', 'B4'], 'cells without a source rule are cleared')
  assert.deepEqual(coverage(true), ['A1', 'A2', 'A3', 'A4', 'B1', 'B3'])
  assert.ok((validated.validations ?? []).every((item) => item.validation === null || item.validation !== listRule), 'rules are copied, not shared')
  assert.equal(validated.validations?.[0].validation, null, 'clears come first')
  assert.ok((validated.validations?.length ?? 0) <= 6, 'rectangles are coalesced')

  const tooLarge = applyPasteSpecial({
    source: { cells: [[{ value: 1 }]], origin: { row: 0, col: 0 } },
    destination: { top: 0, left: 0, bottom: 999, right: 199 },
    getDestinationCell: () => undefined,
  })
  assert.equal(tooLarge.error, 'too-large')
  assert.deepEqual(tooLarge.changes, {})
  assert.ok(PASTE_SPECIAL_MENU.some((item) => item.shortcut === 'Ctrl+Shift+V' && item.preset === 'values'))
  assert.ok(PASTE_SPECIAL_MENU.every((item) => item.opensDialog || (item.preset && item.preset in PASTE_SPECIAL_PRESETS)))
})

test('pasting parsed external content through paste special', () => {
  const parsed = parseClipboardHtml(GOOGLE_SHEETS, { ...builtin, origin: { row: 4, col: 3 } })
  assert.ok(parsed)
  const result = applyPasteSpecial({
    source: { cells: parsed.cells, origin: parsed.formulaOrigin, merges: parsed.merges },
    destination: { top: 4, left: 3 },
    getDestinationCell: () => undefined,
    shiftFormula: shiftFormulaReferences,
  })
  assert.equal(result.changes.F6?.formula, 'E6*2')
  assert.equal(result.changes.E8?.formula, 'SUM(F6:F7)/$C$2')
  assert.deepEqual(result.merges, ['D7:D8'])
  assert.deepEqual(result.clearMergesIn, { top: 4, left: 3, bottom: 7, right: 5 })
})

// ---------------------------------------------------------------------------------------------
// Performance
// ---------------------------------------------------------------------------------------------

test('performance at the 100k-cell cap', () => {
  const rows = 400
  const cols = 250
  const cells: CellData[][] = Array.from({ length: rows }, (_, row) => Array.from({ length: cols }, (_, col) => (
    col % 3 === 0 ? { value: row * cols + col, numFmt: '#,##0.00', style: { font: { bold: row === 0 } } }
      : col % 3 === 1 ? { value: `r${row}c${col}` }
        : { formula: `A${row + 1}*2`, result: row * 2 }
  )))
  let started = performance.now()
  const { html, text } = serializeSelectionToClipboard(cells, { includeInternalData: false })
  const serializeMs = performance.now() - started
  started = performance.now()
  const parsed = parseClipboardHtml(html, builtin)
  const parseMs = performance.now() - started
  assert.equal(parsed?.cells.length, rows)
  assert.equal(parsed?.cells[399][248].formula, 'A400*2')
  started = performance.now()
  const fromText = parsePastedText(text)
  const textMs = performance.now() - started
  assert.equal(fromText[10][0].value, 2500)
  assert.ok(serializeMs < 4000 && parseMs < 8000 && textMs < 3000, `too slow: serialize ${serializeMs}ms, parse ${parseMs}ms, text ${textMs}ms`)
  console.log(`  100k cells: serialize ${serializeMs.toFixed(0)}ms (${(html.length / 1e6).toFixed(1)} MB), parse HTML ${parseMs.toFixed(0)}ms, parse TSV ${textMs.toFixed(0)}ms`)
  assert.throws(() => parseClipboardHtml(html.replace('</tbody>', `<tr>${'<td>1</td>'.repeat(cols)}</tr></tbody>`), builtin), ClipboardTooLargeError)
})

console.log(`Clipboard HTML QA passed: ${checks} groups (Excel, Google Sheets, LibreOffice, Word, web tables, TSV, R1C1, round trips, paste special).`)
