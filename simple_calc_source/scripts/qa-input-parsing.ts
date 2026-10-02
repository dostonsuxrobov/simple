import assert from 'node:assert/strict'
import { formulaFromSignedEntry, inferFormulaNumberFormat, parseCellInput, textNeedsQuotePrefix } from '../src/lib/input-parsing'

const now = new Date(2026, 8, 28)
const p = (text: string) => parseCellInput(text, now)
assert.deepEqual(p('42'), { value: 42 })
assert.deepEqual(p('-3.5e2'), { value: -350 })
assert.equal(p('007'), null)
assert.deepEqual(p('1,234'), { value: 1234, numFmt: '#,##0' })
assert.deepEqual(p('-12,345.67'), { value: -12345.67, numFmt: '#,##0.00' })
assert.deepEqual(p('$1,234.50'), { value: 1234.5, numFmt: '$#,##0.00' })
assert.deepEqual(p('$5'), { value: 5, numFmt: '$#,##0' })
assert.deepEqual(p('-$5'), { value: -5, numFmt: '$#,##0' })
assert.deepEqual(p('($5.25)'), { value: -5.25, numFmt: '$#,##0.00' })
assert.deepEqual(p('€12'), { value: 12, numFmt: '"€"#,##0' })
assert.deepEqual(p('(123.45)'), { value: -123.45 })
assert.deepEqual(p('15%'), { value: 0.15, numFmt: '0%' })
assert.deepEqual(p('2.5%'), { value: 0.025, numFmt: '0.0%' })
assert.deepEqual(p('1 1/2'), { value: 1.5, numFmt: '# ?/?' })
assert.deepEqual(p('TRUE'), { value: true })
assert.deepEqual(p('#n/a'), { value: '#N/A', type: 'error' })
assert.deepEqual(p('2024-01-15'), { value: 45306, numFmt: 'yyyy-mm-dd', type: 'date' })
assert.deepEqual(p('1/15/2024'), { value: 45306, numFmt: 'm/d/yyyy', type: 'date' })
assert.deepEqual(p('1/2'), { value: 46024, numFmt: 'd-mmm', type: 'date' })
assert.deepEqual(p('15-Jan-2024'), { value: 45306, numFmt: 'd-mmm-yy', type: 'date' })
assert.deepEqual(p('Jan 15, 2024'), { value: 45306, numFmt: 'mmm d, yyyy', type: 'date' })
assert.deepEqual(p('January 15 2024'), { value: 45306, numFmt: 'mmm d, yyyy', type: 'date' })
assert.deepEqual(p('Mar-24'), { value: 45352, numFmt: 'mmm-yy', type: 'date' })
assert.deepEqual(p('10:30'), { value: 0.4375, numFmt: 'h:mm' })
assert.deepEqual(p('10:30 PM'), { value: 22.5 / 24, numFmt: 'h:mm AM/PM' })
assert.deepEqual(p('7 am'), { value: 7 / 24, numFmt: 'h:mm AM/PM' })
assert.deepEqual(p('1/15/2024 10:30'), { value: 45306 + 0.4375, numFmt: 'm/d/yyyy h:mm', type: 'date' })
assert.equal(p('2/30/2024'), null)
assert.equal(p('hello'), null)
assert.equal(p('Mayday'), null)
assert.equal(p('A1'), null)

const formats: Record<string, string> = { A1: 'm/d/yyyy', B2: '$#,##0.00', C3: 'General' }
const formatOf = (address: string) => formats[address]
assert.equal(inferFormulaNumberFormat('=TODAY()', formatOf), 'm/d/yyyy')
assert.equal(inferFormulaNumberFormat('=NOW()', formatOf), 'm/d/yyyy h:mm')
assert.equal(inferFormulaNumberFormat('=A1+7', formatOf), 'm/d/yyyy')
assert.equal(inferFormulaNumberFormat('=B2*1.1', formatOf), '$#,##0.00')
assert.equal(inferFormulaNumberFormat('=SUM(B2:B9)', formatOf), '$#,##0.00')
assert.equal(inferFormulaNumberFormat('=COUNT(B2:B9)', formatOf), undefined)
assert.equal(inferFormulaNumberFormat('=B2*C3', formatOf), undefined)
assert.equal(inferFormulaNumberFormat('=A1>5', formatOf), undefined)
assert.equal(inferFormulaNumberFormat('=C3+1', formatOf), undefined)

// calc-formula-engine-12: typed dates follow Excel's 1900 date system (serial 60 is 1900-02-29).
assert.deepEqual(p('1/1/1900'), { value: 1, numFmt: 'm/d/yyyy', type: 'date' })
assert.deepEqual(p('2/28/1900'), { value: 59, numFmt: 'm/d/yyyy', type: 'date' })
assert.deepEqual(p('2/29/1900'), { value: 60, numFmt: 'm/d/yyyy', type: 'date' })
assert.deepEqual(p('3/1/1900'), { value: 61, numFmt: 'm/d/yyyy', type: 'date' })
assert.deepEqual(p('1900-01-02'), { value: 2, numFmt: 'yyyy-mm-dd', type: 'date' })
assert.equal(p('2/29/1901'), null)
assert.equal(p('12/31/1899'), null)

// calc-formula-engine-10: text that would be read back as something else keeps its apostrophe
// in the editor, so editing it again keeps it text.
for (const text of ['12345678901234567890', '3/4', 'TRUE', '=A1', "'quoted", '15%', '#N/A', '+B1+C1', '-A1', '1,000']) {
  assert.equal(textNeedsQuotePrefix(text), true, `"${text}" needs an apostrophe`)
}
for (const text of ['hello', '- item', 'Mayday', '+1-555-0100', 'A1', '007', '']) {
  assert.equal(textNeedsQuotePrefix(text), false, `"${text}" stays as typed`)
}

// calc-formula-engine-7: "+B1+C1" and "-A1" typed without "=" are formulas, as in Excel.
assert.equal(formulaFromSignedEntry('+B1+C1'), '+B1+C1')
assert.equal(formulaFromSignedEntry('-A1*2'), '-A1*2')
assert.equal(formulaFromSignedEntry('+SUM(A1:A3)'), '+SUM(A1:A3)')
assert.equal(formulaFromSignedEntry('-5'), null, 'numbers stay numbers')
assert.equal(formulaFromSignedEntry('- item'), null)
assert.equal(formulaFromSignedEntry('+1-555-0100'), null)
assert.equal(formulaFromSignedEntry('-total'), null)
assert.equal(formulaFromSignedEntry('A1+1'), null)
assert.equal(formulaFromSignedEntry('+SUM('), null)
console.log('Input parsing QA passed: numbers, currency, percentages, fractions, dates, times, inferred formula formats, 1900 dates, apostrophe text, and +/- formula entry.')
