'use strict'

// Delimited text (CSV/TSV): Excel's quote rules, per-column decimal and date inference,
// encodings, values written as displayed with four-digit years, RFC 4180 quoting, and saving a
// file back in the dialect it was read in.
const assert = require('node:assert/strict')
const { workbookPayloadFromBytes, serializeWorkbook, delimitedDialectFor } = require('../electron/workbooks.cjs')
const { encodeText } = require('../electron/delimited-text.cjs')

const serial = (year, month, day) => (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86_400_000
const open = (name, text, options = { locale: 'en-US' }) => workbookPayloadFromBytes(name, Buffer.isBuffer(text) ? text : Buffer.from(text, 'utf8'), options)
const cellsOf = (payload) => payload.workbook.sheets[0].cells
const value = (payload, address) => cellsOf(payload)[address] && cellsOf(payload)[address].value

/** Save the way main does for an in-format save: the dialect recorded at open is kept. */
async function saveInPlace(payload) {
  const format = payload.sourceFormat
  return serializeWorkbook(payload.workbook, format, { dialect: delimitedDialectFor(payload.workbook, format) })
}

async function quotes() {
  // calc-file-io-objects-6: a quote inside an unquoted field is a literal character.
  const inches = await open('inventory.csv', 'Product,Price,Qty\r\nTV 55" OLED,1299.99,2\r\nMonitor 27",249.50,5\r\nCable,9.99,10\r\n')
  assert.equal(value(inches, 'A2'), 'TV 55" OLED')
  assert.equal(value(inches, 'B2'), 1299.99)
  assert.equal(value(inches, 'C2'), 2)
  assert.equal(value(inches, 'A3'), 'Monitor 27"')
  assert.equal(value(inches, 'B3'), 249.5)
  assert.equal(value(inches, 'A4'), 'Cable')
  assert.equal(inches.workbook.sheets[0].rowCount, 4)
  // A quoted field may follow spaces; text after a closing quote is kept; "" is a quote.
  const lenient = await open('lenient.csv', 'a, "b, c" ,d\r\n"x"y,"say ""hi""",z\r\n')
  assert.equal(value(lenient, 'B1'), 'b, c')
  assert.equal(value(lenient, 'C1'), 'd')
  assert.equal(value(lenient, 'A2'), 'xy')
  assert.equal(value(lenient, 'B2'), 'say "hi"')
}

async function numbers() {
  // calc-file-io-objects-3: a semicolon file with dot decimals is not read as thousands.
  const readings = await open('readings.csv', 'sensor;reading;ratio\r\nA;1.5;0.125\r\nB;2.25;3.142\r\nC;1234.567;12.500\r\n')
  assert.equal(value(readings, 'C2'), 0.125)
  assert.equal(value(readings, 'C3'), 3.142)
  assert.equal(value(readings, 'C4'), 12.5)
  assert.equal(cellsOf(readings).C4.numFmt, '0.000', 'trailing zeros of the source stay visible')
  assert.equal(value(readings, 'B4'), 1234.567)
  assert.equal(readings.workbook.metadata.decimalComma, false)
  // Ambiguous values with only weak, uncorroborated evidence stay text (never scaled 1000x).
  const prices = await open('prices.csv', 'Item;Price\r\nA;1.250\r\nB;0.5\r\nC;12.000')
  assert.equal(value(prices, 'B2'), '1.250')
  assert.equal(value(prices, 'B3'), 0.5)
  assert.equal(value(prices, 'B4'), '12.000')
  assert.ok(prices.warnings.some((warning) => /Column B/.test(warning)), 'the user is told why the column stayed text')
  const tabbed = await open('amounts.tsv', 'k\tv\r\na\t2,000\r\nb\t1,234\r\nc\t1,5')
  assert.equal(value(tabbed, 'B2'), '2,000')
  assert.equal(value(tabbed, 'B3'), '1,234')
  assert.equal(value(tabbed, 'B4'), 1.5)
  // Decimal commas and grouping (dot, space, no-break space, apostrophe).
  const eu = await open('eu.csv', 'name;value;price;spaced;swiss\r\nwidget;1,5;1.234,56;1 234,56;1\'234\r\ngadget;2;00123;12 500;2\'500')
  assert.equal(eu.workbook.metadata.decimalComma, true)
  assert.equal(value(eu, 'B2'), 1.5)
  assert.equal(value(eu, 'C2'), 1234.56)
  assert.equal(value(eu, 'D2'), 1234.56)
  assert.equal(value(eu, 'E2'), 1234)
  assert.equal(value(eu, 'D3'), 12500)
  assert.equal(value(eu, 'C3'), '00123')
  // No evidence at all: a semicolon file reads EU grouping; a comma file reads US decimals.
  assert.equal(value(await open('menge.csv', 'k;v\r\na;1.250\r\nb;2.500'), 'B2'), 1250)
  assert.equal(value(await open('menge.csv', 'k,v\r\na,1.250\r\nb,2.500'), 'B2'), 1.25)
  // Quoted "n,n" pairs in a comma file stay text unless the file shows EU grouping.
  const pairs = await open('pairs.csv', 'sku,sizes,weight\r\nA1,"6,8",1.234\r\nB2,"8,10",2.5\r\nC3,"10,12",0.75')
  assert.equal(value(pairs, 'B2'), '6,8')
  assert.equal(value(pairs, 'C2'), 1.234)
  assert.equal(value(await open('grouped.csv', 'a,b\r\nx,"1.234,56"'), 'B2'), 1234.56)
  // calc-file-io-objects-4: %, $ and grouped values keep a format that shows the source text.
  const formatted = await open('orders.csv', 'Item,Price,Discount,Total,Neg\r\nWidget,"$1,234.50",12.5%,"1,080.19","(1,000.00)"\r\n')
  assert.equal(value(formatted, 'B2'), 1234.5)
  assert.equal(cellsOf(formatted).B2.numFmt, '"$"#,##0.00')
  assert.equal(value(formatted, 'C2'), 0.125)
  assert.equal(cellsOf(formatted).C2.numFmt, '0.0%')
  assert.equal(cellsOf(formatted).D2.numFmt, '#,##0.00')
  assert.equal(value(formatted, 'E2'), -1000)
  const euro = await open('euro.csv', 'a;b\r\n1.234,56 €;-12,5 %')
  assert.equal(value(euro, 'A2'), 1234.56)
  assert.equal(cellsOf(euro).A2.numFmt, '#,##0.00 "€"')
  assert.equal(value(euro, 'B2'), -0.125)
}

async function dates() {
  // Per-column day/month order: one unambiguous date decides the column.
  const dmy = await open('dates.csv', 'd\r\n15/03/2024\r\n01/02/2024\r\n12/11/2024')
  assert.deepEqual(['A2', 'A3', 'A4'].map((address) => value(dmy, address)), [serial(2024, 3, 15), serial(2024, 2, 1), serial(2024, 11, 12)])
  const mdy = await open('dates-us.csv', 'd\r\n03/15/2024\r\n01/02/2024')
  assert.equal(value(mdy, 'A3'), serial(2024, 1, 2))
  // European dot dates, with decimal commas elsewhere.
  const dots = await open('dots.csv', 'a;b\r\n15.03.2024;1,5\r\n01.02.2024;2')
  assert.equal(value(dots, 'A2'), serial(2024, 3, 15))
  assert.equal(value(dots, 'A3'), serial(2024, 2, 1))
  assert.equal(cellsOf(dots).A2.numFmt, 'dd\\.mm\\.yyyy')
  // calc-file-io-objects-4 / CALC-SIE-7: four-digit years keep a four-digit format.
  const born = await open('people.csv', 'Name,Born\r\nAda,01/02/1949\r\nBob,12/31/1925\r\n')
  assert.equal(value(born, 'B2'), serial(1949, 1, 2))
  assert.equal(cellsOf(born).B2.numFmt, 'mm/dd/yyyy')
  const saved = (await saveInPlace(born)).toString('utf8')
  assert.equal(saved, 'Name,Born\r\nAda,01/02/1949\r\nBob,12/31/1925\r\n', 'a save in place is byte-identical')
  const reopened = await open('people.csv', saved)
  assert.equal(value(reopened, 'B2'), serial(1949, 1, 2), '1949 does not become 2049')
  // Excel's two-digit window: 29 -> 2029, 30 -> 1930.
  const short = await open('short.csv', 'd\r\n1/2/29\r\n1/2/30')
  assert.equal(value(short, 'A2'), serial(2029, 1, 2))
  assert.equal(value(short, 'A3'), serial(1930, 1, 2))
  // Times and date-times.
  const times = await open('times.csv', 't,dt\r\n14:30,3/15/2024 2:30 PM')
  assert.equal(value(times, 'A2'), 14.5 / 24)
  assert.equal(value(times, 'B2'), serial(2024, 3, 15) + 14.5 / 24)
}

async function encodings() {
  const utf16 = await open('u16.csv', Buffer.from('Name,Value\r\nA,1\r\n', 'utf16le'))
  assert.equal(value(utf16, 'A1'), 'Name')
  assert.equal(utf16.workbook.metadata.dialect.encoding, 'utf-16le')
  const sjis = await open('sjis.csv', Buffer.from([0x82, 0xa0, 0x2c, 0x82, 0xa2, 0x0d, 0x0a, 0x31, 0x2c, 0x32]))
  assert.equal(value(sjis, 'A1'), 'あ')
  assert.equal(value(sjis, 'B1'), 'い')
  assert.equal(sjis.workbook.metadata.dialect.encoding, 'shift_jis')
  const cyrillic = await open('cyr.csv', Buffer.from('\xcf\xf0\xe8\xe2\xe5\xf2;\xec\xe8\xf0\r\n1;2\r\n', 'latin1'))
  assert.equal(value(cyrillic, 'A1'), 'Привет')
  const latin = await open('latin.csv', Buffer.from('Name;Ort\r\nM\xfcller;K\xf6ln\r\n', 'latin1'))
  assert.equal(value(latin, 'A2'), 'Müller')
  assert.equal(latin.workbook.metadata.dialect.encoding, 'windows-1252')
  assert.equal(encodeText('あ,い', 'shift_jis', false).toString('hex'), '82a02c82a2')
  assert.equal(encodeText('€', 'windows-1252', false).toString('hex'), '80')
  assert.equal(encodeText('あ', 'windows-1252', false), null, 'an unrepresentable character is reported')
}

async function writer() {
  // CALC-SIE-7: RFC 4180 quoting (line breaks, delimiters, quotes, edge spaces) and 4-digit years.
  const model = {
    version: 1, name: 'rfc', activeSheetId: 's', metadata: {},
    sheets: [{ id: 's', name: 'S', rowCount: 4, colCount: 4, merges: [], cells: {
      A1: { value: 'Name' }, B1: { value: 'Note' }, C1: { value: 'Born' }, D1: { value: 'Text' },
      A2: { value: 'a,b' }, B2: { value: 'line1\nline2' }, C2: { value: serial(1945, 1, 1), numFmt: 'm/d/yy' }, D2: { value: ' lead' },
      A3: { value: 'say "hi"' }, B3: { value: 'x' }, C3: { value: serial(2060, 12, 31), numFmt: 'm/d/yyyy' }, D3: { value: 'trail ' },
      A4: { formula: 'SUM(1,2)', result: 3 }, B4: { value: true }, C4: { value: 0.1 + 0.2 }, D4: { value: '#N/A', type: 'error' },
    } }],
  }
  const csv = await serializeWorkbook(model, 'csv')
  assert.deepEqual([...csv.subarray(0, 3)], [0xef, 0xbb, 0xbf], 'a new CSV is UTF-8 with BOM for Excel')
  const text = csv.toString('utf8').replace(/^﻿/, '')
  assert.equal(text, 'Name,Note,Born,Text\r\n"a,b","line1\nline2",1/1/1945," lead"\r\n"say ""hi""",x,12/31/2060,"trail "\r\n3,TRUE,0.3,#N/A\r\n')
  const back = await open('rfc.csv', csv)
  assert.equal(back.workbook.sheets[0].rowCount, 4, 'an embedded line break does not split the row')
  assert.equal(value(back, 'B2'), 'line1\nline2')
  assert.equal(value(back, 'A2'), 'a,b')
  assert.equal(value(back, 'D2'), ' lead')
  assert.equal(value(back, 'A3'), 'say "hi"')
  assert.equal(value(back, 'C2'), serial(1945, 1, 1))
  assert.equal(value(back, 'C3'), serial(2060, 12, 31))
  const tsv = (await serializeWorkbook(model, 'tsv')).toString('utf8')
  assert.match(tsv, /^﻿Name\tNote\tBorn\tText\r\n/)
  assert.match(tsv, /\r\na,b\t"line1\nline2"\t/)
}

async function dialects() {
  // CALC-SIE-18 / calc-file-io-objects-4: a save in place keeps delimiter, decimal comma,
  // grouping, encoding (cp1252, no BOM) and line endings; only the edited cell changes.
  const source = Buffer.from('Name;Betrag;Datum\r\nM\xfcller;1.234,56;15.03.2024\r\nSch\xf6n;-7,50;01.02.2024\r\n', 'latin1')
  const payload = await open('konto.csv', source)
  const dialect = payload.workbook.metadata.dialect
  assert.equal(dialect.delimiter, ';')
  assert.equal(dialect.decimalComma, true)
  assert.equal(dialect.encoding, 'windows-1252')
  assert.equal(dialect.hadBom, false)
  payload.workbook.sheets[0].cells.A3 = { value: 'Sch\xf6ner' }
  const saved = await saveInPlace(payload)
  assert.ok(saved.equals(Buffer.from('Name;Betrag;Datum\r\nM\xfcller;1.234,56;15.03.2024\r\nSch\xf6ner;-7,50;01.02.2024\r\n', 'latin1')), `cp1252 round trip: ${JSON.stringify(saved.toString('latin1'))}`)
  // New values follow the dialect: decimal comma and dot grouping.
  payload.workbook.sheets[0].cells.B4 = { value: 9876.5, numFmt: '#,##0.00' }
  payload.workbook.sheets[0].cells.C4 = { value: 0.25 }
  assert.match((await saveInPlace(payload)).toString('latin1'), /\r\n;9\.876,50;0,25\r\n$/)
  // A character the source encoding cannot hold falls back to UTF-8 with a note, never '?'.
  payload.workbook.sheets[0].cells.A4 = { value: 'あ' }
  const warnings = []
  const fallback = await serializeWorkbook(payload.workbook, 'csv', { dialect: delimitedDialectFor(payload.workbook, 'csv'), warnings })
  assert.match(fallback.toString('utf8'), /あ/)
  assert.ok(warnings.some((warning) => /UTF-8/.test(warning)))
  // A tab-delimited .csv stays tab-delimited; UTF-8 without BOM and LF endings are kept.
  const tab = Buffer.from('a\tb\r\n1,5\t2\r\n')
  assert.ok((await saveInPlace(await open('tab.csv', tab))).equals(tab))
  const lf = Buffer.from('x,y\n1,2\n')
  assert.ok((await saveInPlace(await open('lf.csv', lf))).equals(lf))
  const spaced = Buffer.from('a;b\r\nx;1 234,56\r\n')
  assert.ok((await saveInPlace(await open('spaced.csv', spaced))).equals(spaced))
  const ragged = Buffer.from('a,b,c\r\nd\r\n\r\ne,f\r\n')
  assert.ok((await saveInPlace(await open('ragged.csv', ragged))).equals(ragged), 'ragged rows are not padded')
  // A save in another format (or an export) uses predictable defaults.
  assert.equal(delimitedDialectFor(payload.workbook, 'tsv'), undefined)
  const asDefault = (await serializeWorkbook((await open('konto.csv', source)).workbook, 'csv')).toString('utf8')
  assert.match(asDefault, /^﻿Name,Betrag,Datum\r\nMüller,"1,234.56",15.03.2024\r\n/)
}

async function main() {
  await quotes()
  await numbers()
  await dates()
  await encodings()
  await writer()
  await dialects()
  console.log('Delimited QA passed: literal quotes, per-column decimals and dates, grouping, encodings, RFC 4180 writing, four-digit years, and dialect-preserving saves.')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
