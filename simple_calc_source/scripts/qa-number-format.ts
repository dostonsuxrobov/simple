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

process.stdout.write('Number-format QA passed: Excel accounting, date, and percentage masks render correctly.\n')
