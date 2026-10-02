import assert from 'node:assert/strict'
import { accountingDisplayParts, formatScalar, isAccountingNumberFormat } from '../src/lib/number-format'

const accountingCurrency = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
const accountingInteger = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
const compact = (value: string) => value.replace(/\s/g, '')

assert.equal(compact(formatScalar(1200, accountingCurrency)), '$1,200.00')
assert.equal(compact(formatScalar(0, accountingCurrency)), '$-')
assert.equal(compact(formatScalar(-1200, accountingCurrency)), '$(1,200.00)')
assert.equal(compact(formatScalar(0, accountingInteger)), '-')
assert.equal(compact(formatScalar(1, accountingInteger)), '1')
assert.equal(formatScalar(0.125, '0.00%'), '12.50%')
assert.equal(formatScalar(45_292, 'm/d/yy'), '1/1/24')

assert.deepEqual(accountingDisplayParts(formatScalar(1200, accountingCurrency), accountingCurrency), {
  symbol: '$',
  amount: '1,200.00',
})
assert.equal(accountingDisplayParts('$1,200.00', '$#,##0.00'), null)
assert.equal(isAccountingNumberFormat(accountingInteger), true)
assert.equal(isAccountingNumberFormat('$#,##0.00'), false)

// calc-formula-engine-9: unformatted numbers show in General (no binary noise), logicals as
// TRUE/FALSE, and formatted numbers round half away from zero on their decimal value.
assert.equal(formatScalar(0.1 + 0.2), '0.3')
assert.equal(formatScalar(4.35 * 100), '435')
assert.equal(formatScalar(1 / 3), '0.333333333')
assert.equal(formatScalar(123456789012), '1.23457E+11')
assert.equal(formatScalar(-3.25, ''), '-3.25')
assert.equal(formatScalar(42), '42')
assert.equal(formatScalar(true), 'TRUE')
assert.equal(formatScalar(false, 'General'), 'FALSE')
assert.equal(formatScalar(0.1 + 0.2, undefined, '0.3000'), '0.3000', "an imported cell's own text still wins")
assert.equal(formatScalar(-2.5, '0'), '-3')
assert.equal(formatScalar(-0.5, '0'), '-1')
assert.equal(formatScalar(2.5, '0'), '3')
assert.equal(formatScalar(-0.125, '0.00'), '-0.13')
assert.equal(formatScalar(-0.005, '0.00'), '-0.01')
assert.equal(formatScalar(1.005, '0.00'), '1.01')
assert.equal(formatScalar(2.675, '0.00'), '2.68')
assert.equal(formatScalar(-0.125, '0%'), '-13%')
assert.equal(formatScalar(0.12345, '0.00%'), '12.35%')
assert.equal(formatScalar(-2.5, '#,##0'), '-3')
assert.equal(formatScalar(1234.565, '$#,##0.00'), '$1,234.57')
assert.equal(formatScalar(12345678, '#,##0.0,'), '12,345.7')
assert.equal(formatScalar(-1.005, '#,##0.00_);(#,##0.00)'), '(1.01)')
assert.equal(formatScalar(-0.004, '0.00;(0.00)'), '(0.00)', 'the negative section is kept for values that round to zero')
assert.equal(formatScalar(99.999, '[>=100]0;0.00'), '100.00', 'conditions use the unrounded value')
assert.equal(compact(formatScalar(-1234.565, accountingCurrency)), '$(1,234.57)')

process.stdout.write('Number-format QA passed: Excel accounting, date, and percentage masks render correctly; General display and Excel rounding.\n')
