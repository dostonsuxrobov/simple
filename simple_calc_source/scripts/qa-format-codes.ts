import assert from 'node:assert/strict'
import { accountingCurrencySymbol, accountingDisplayParts, formatColor, formatScalar, formatScalarDetailed } from '../src/lib/number-format'
import {
  CURRENCY_SYMBOLS, CUSTOM_FORMAT_PRESETS, DATE_FORMATS, FRACTION_FORMATS, NEGATIVE_NUMBER_STYLES, SPECIAL_FORMATS, TIME_FORMATS,
  adjustDecimals, buildNumberFormat, detectNumberFormat, formatTypeSample, isDateTimeFormat, validateNumberFormat,
} from '../src/lib/format-codes'
import type { NumberFormatCategory, NumberFormatOptions } from '../src/lib/format-codes'
import {
  CELL_STYLES, applyBorderPreset, applyCellStylePreset, applyCellStylePresetToCell, applyFormatChange, applyFormatChangeToCell,
  applyNeighborBorderChange, applyTint, cellPosition, fillPreviewCss, patternTileUri, resolveColorHex, sameColor, themePalette, themeTintSteps,
} from '../src/lib/cell-styles'
import type { CellPosition, FormatCellsChange } from '../src/lib/cell-styles'
import { FONT_FAMILIES, fontStack, parseFontSize } from '../src/lib/fonts'
import type { CellBorderSide, CellStyle } from '../src/spreadsheet-types'

let checks = 0
const eq = (actual: unknown, expected: unknown, message?: string) => { checks += 1; assert.deepEqual(actual, expected, message) }
const ok = (value: unknown, message?: string) => { checks += 1; assert.ok(value, message) }
const compact = (value: string) => value.replace(/\s+/g, ' ').trim()
const detailed = (value: Parameters<typeof formatScalarDetailed>[0], code: string) => formatScalarDetailed(value, code)

// ---------------------------------------------------------------------------------------
// formatScalarDetailed: sections, colours, conditions
// ---------------------------------------------------------------------------------------

eq(detailed(-1234.5, '#,##0.00;[Red](#,##0.00)'), { text: '(1,234.50)', color: '#FF0000' })
eq(detailed(1234.5, '#,##0.00;[Red](#,##0.00)'), { text: '1,234.50' })
eq(detailed(0, '#,##0.00;[Red](#,##0.00)'), { text: '0.00' })
eq(detailed(-5, '0.00;[Red]0.00'), { text: '5.00', color: '#FF0000' })
eq(detailed(-5, '#,##0.00_);[Red](#,##0.00)'), { text: '(5.00)', color: '#FF0000' })
eq(detailed(0, '[Blue]#,##0;[Red]-#,##0;[Green]0'), { text: '0', color: '#00FF00' })
eq(detailed(7, '[Blue]#,##0;[Red]-#,##0;[Green]0'), { text: '7', color: '#0000FF' })
eq(detailed(5, '[Color10]0').color, '#008000', '[ColorN] uses the 56-colour palette (Color10 = dark green)')
eq(detailed(5, '[Color3]0').color, '#FF0000')
eq(detailed(5, '[magenta]0.00').color, '#FF00FF', 'colour names are case-insensitive')
eq(detailed(5, '[Cyan]0').color, '#00FFFF')
eq(detailed(5, '[Yellow]0').color, '#FFFF00')
eq(detailed(5, '[White]0').color, '#FFFFFF')
eq(detailed(5, '[Black]0').color, '#000000')
eq(detailed(-3, '[Red]General;[Blue]-General'), { text: '-3', color: '#0000FF' }, 'General negative section shows one minus sign')

// Conditions
eq(detailed(150, '[>=100][Red]0;[Blue]0'), { text: '150', color: '#FF0000' })
eq(detailed(50, '[>=100][Red]0;[Blue]0'), { text: '50', color: '#0000FF' })
eq(detailed(150, '[Red][<=100]0;[Blue][>100]0'), { text: '150', color: '#0000FF' })
eq(detailed(80, '[Red][<=100]0;[Blue][>100]0'), { text: '80', color: '#FF0000' })
eq(detailed(5551234, '[<=9999999]###-####;(###) ###-####').text, '555-1234')
eq(detailed(2125551234, '[<=9999999]###-####;(###) ###-####').text, '(212) 555-1234')
eq(detailed(0, '[=0]"zero";0').text, 'zero')
eq(detailed(1500, '[>=1000]#,##0.0,"K";0').text, '1.5K')

// Zero / hidden sections
eq(detailed(0, '0;-0;;@').text, '')
eq(detailed(-5, '0.00;;').text, '')
eq(detailed(0, '#,##0.00;(#,##0.00);"-"').text, '-')

// Percent, scientific, engineering
eq(formatScalar(0.1234, '0.0%'), '12.3%')
eq(formatScalar(0.5, '0%'), '50%')
eq(formatScalar(12345.678, '0.00E+00'), '1.23E+04')
eq(formatScalar(-0.00012, '0.00E+00'), '-1.20E-04')
eq(formatScalar(0, '0.00E+00'), '0.00E+00')
eq(formatScalar(12345.678, '##0.0E+0'), '12.3E+3')
eq(formatScalar(0.00012345, '##0.00E+00'), '123.45E-06', 'engineering notation with a negative exponent')
eq(formatScalar(-0.00012345, '##0.0E+0'), '-123.5E-6')
eq(formatScalar(999.96, '##0.0E+0'), '1.0E+3', 'mantissa rounding carries into the exponent')
eq(formatScalar(123, '000.00E+00'), '123.00E+00')

// Fractions
eq(compact(formatScalar(1.25, '# ?/?')), '1 1/4')
eq(compact(formatScalar(0.3333, '# ?/?')), '1/3')
eq(compact(formatScalar(3.14159, '# ??/??')), '3 1/7')
eq(compact(formatScalar(-3.14159, '# ??/??')), '-3 1/7')
eq(compact(formatScalar(3.14159, '# ???/???')), '3 16/113')
eq(compact(formatScalar(1.3, '# ?/4')), '1 1/4')
eq(compact(formatScalar(1.3, '# ?/8')), '1 2/8')
eq(compact(formatScalar(1.3, '# ??/16')), '1 5/16')
eq(compact(formatScalar(1.3, '# ?/10')), '1 3/10')
eq(compact(formatScalar(1.337, '# ??/100')), '1 34/100')
eq(compact(formatScalar(0.75, '# ?/2')), '1')
eq(compact(formatScalar(0.25, '0 ?/?')), '0 1/4', 'a zero integer placeholder before a fraction')
eq(compact(formatScalar(-0.25, '0 ?/?')), '-0 1/4')
eq(compact(formatScalar(1.5, '0 ??/??')), '1 1/2')

// Dates and times
eq(formatScalar(45292, 'mmm d, yyyy'), 'Jan 1, 2024')
eq(formatScalar(45292, 'dddd, mmmm d, yyyy'), 'Monday, January 1, 2024')
eq(formatScalar(45292, 'yyyy-mm-dd'), '2024-01-01')
eq(formatScalar(45292, 'd-mmm-yy'), '1-Jan-24')
eq(formatScalar(45292, 'mmmmm'), 'J')
eq(formatScalar(45292, '[$-409]mmmm d, yyyy;@'), 'January 1, 2024')
eq(formatScalar(0.5, '[$-409]h:mm:ss AM/PM'), '12:00:00 PM')
eq(formatScalar(1.5, '[h]:mm:ss'), '36:00:00', 'elapsed hours')
eq(formatScalar(2.25, '[h]:mm'), '54:00')
eq(formatScalar(0.05, '[mm]:ss'), '72:00', 'elapsed minutes')
eq(formatScalar(0.5, '[ss]'), '43200', 'elapsed seconds')
eq(formatScalar(0.75, 'h:mm AM/PM'), '6:00 PM')
eq(formatScalar(45292.5, 'm/d/yyyy h:mm'), '1/1/2024 12:00')
eq(formatScalar(40982.5631365741, 'yyyy-mm-ddThh:mm:ss'), '2012-03-14T13:30:55', 'unquoted literal letters are tolerated')
eq(detailed(-0.5, '[h]:mm:ss'), { text: '########', overflow: true }, 'negative times fill with #')
eq(detailed(3_000_000, 'm/d/yyyy'), { text: '########', overflow: true }, 'serials past 9999-12-31 fill with #')
eq(formatScalar(60, 'm/d/yyyy'), '2/29/1900', 'Excel\'s 1900 leap-year bug is preserved')
eq(formatScalar(40982.5631365741, 'dd.mm.yyyy'), '14.03.2012', 'dotted European dates (SSF rejects them unquoted)')
eq(formatScalar(40982.5631365741, 'd. mmmm yyyy'), '14. March 2012')
eq(formatScalar(40982.5631365741, '[$-407]dd.mm.yyyy hh:mm'), '14.03.2012 13:30')
eq(formatScalar(40982.5631365741, 'hh:mm:ss.00'), '13:30:55.00', 'the seconds fraction keeps its decimal point')

// Text and the text section
eq(formatScalar(12, '@'), '12')
eq(detailed('abc', '[Red]@'), { text: 'abc', color: '#FF0000' })
eq(detailed('abc', '0;-0;0;[Blue]@'), { text: 'abc', color: '#0000FF' })
eq(detailed('abc', ';;;'), { text: '' }, ';;; hides text too')
eq(detailed('abc', '"Item: "@'), { text: 'Item: abc' })
eq(detailed('abc', '0.00'), { text: 'abc' }, 'numeric-only formats leave text alone')
eq(formatScalar('abc', '_(@_)', 'abc'), 'abc', 'padding-only text sections keep the stored display')
eq(formatColor('abc', '[Red]@'), '#FF0000')
eq(formatColor(-1, '0;[Red]-0'), '#FF0000')
eq(formatColor(1, '0;[Red]-0'), undefined)

// Currency, accounting, locale tags
eq(formatScalar(1234.5, '"$"#,##0'), '$1,235')
eq(formatScalar(1234.5, '[$€-x-euro2] #,##0.00'), '€ 1,234.50')
eq(formatScalar(1234.5, '#,##0.00 [$€-x-euro1]'), '1,234.50 €')
eq(formatScalar(1234, '[$¥-411]#,##0'), '¥1,234')
eq(formatScalar(1234, '[$£-809]#,##0.00'), '£1,234.00')
eq(formatScalar(1234, '#,##0.00 [$₽-419]'), '1,234.00 ₽')
eq(formatScalar(5, '[$CHF-807] #,##0.00'), 'CHF 5.00')
const usAccounting = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
eq(compact(formatScalar(1234.5, usAccounting)), '$1,234.50')
eq(compact(formatScalar(-1234.5, usAccounting)), '$(1,234.50)')
eq(compact(formatScalar(0, usAccounting)), '$-')
eq(accountingDisplayParts(formatScalar(1234.5, usAccounting), usAccounting), { symbol: '$', amount: '1,234.50' })
eq(accountingCurrencySymbol('_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * "-"??_-;_-@_-'), '€')
eq(accountingCurrencySymbol('_-[$CHF-807] * #,##0.00_-;-[$CHF-807] * #,##0.00_-;_-[$CHF-807] * "-"??_-;_-@_-'), 'CHF')
eq(accountingCurrencySymbol('_-* #,##0.00 [$€-x-euro1]_-;-* #,##0.00 [$€-x-euro1]_-;_-* "-"?? [$€-x-euro1]_-;_-@_-'), '', 'a trailing symbol stays attached to the number')

// ---------------------------------------------------------------------------------------
// Build / detect round trips
// ---------------------------------------------------------------------------------------

const roundTrip = (category: NumberFormatCategory, options: NumberFormatOptions) => {
  const code = buildNumberFormat(category, options)
  const detected = detectNumberFormat(code)
  eq(detected.category, category, `category round-trips for ${code}`)
  for (const [key, value] of Object.entries(options)) {
    eq((detected.options as Record<string, unknown>)[key], value, `${key} round-trips for ${code}`)
  }
  eq(buildNumberFormat(detected.category, detected.options), code, `rebuild is stable for ${code}`)
  ok(validateNumberFormat(code).valid, `${code} is valid`)
}

eq(buildNumberFormat('number', { decimals: 2, thousands: true, negative: 'red-parens' }), '#,##0.00_);[Red](#,##0.00)')
eq(buildNumberFormat('number', { decimals: 0, thousands: false, negative: 'red' }), '0;[Red]0')
eq(buildNumberFormat('currency', { decimals: 2, symbol: 'usd', negative: 'parens' }), '$#,##0.00_);($#,##0.00)')
eq(buildNumberFormat('currency', { decimals: 2, symbol: 'eur-prefix', negative: 'minus' }), '[$€-x-euro2] #,##0.00')
eq(buildNumberFormat('accounting', { decimals: 2, symbol: 'usd' }), usAccounting)
eq(buildNumberFormat('accounting', { decimals: 0, symbol: 'none' }), '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)')
eq(buildNumberFormat('accounting', { decimals: 2, symbol: 'eur-suffix' }), '_-* #,##0.00 [$€-x-euro1]_-;-* #,##0.00 [$€-x-euro1]_-;_-* "-"?? [$€-x-euro1]_-;_-@_-')
eq(buildNumberFormat('accounting', { decimals: 2, symbol: 'gbp' }), '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-')
eq(buildNumberFormat('percentage', { decimals: 1 }), '0.0%')
eq(buildNumberFormat('scientific', { decimals: 2 }), '0.00E+00')
eq(buildNumberFormat('scientific', { decimals: 0 }), '0E+00')
eq(buildNumberFormat('special', { type: 'phone' }), '[<=9999999]###-####;(###) ###-####')
eq(buildNumberFormat('text'), '@')
eq(buildNumberFormat('general'), 'General')
eq(buildNumberFormat('custom', { code: ' 0.0" kg" ' }), '0.0" kg"')
eq(buildNumberFormat('number', { decimals: 99 }), `0.${'0'.repeat(30)}`, 'decimals clamp to 30')

for (const decimals of [0, 1, 2, 5]) {
  for (const thousands of [false, true]) {
    for (const negative of NEGATIVE_NUMBER_STYLES) roundTrip('number', { decimals, thousands, negative: negative.id })
  }
  for (const symbol of CURRENCY_SYMBOLS.filter((entry) => entry.id !== 'none')) {
    for (const negative of NEGATIVE_NUMBER_STYLES) roundTrip('currency', { decimals, symbol: symbol.id, negative: negative.id })
  }
  for (const symbol of CURRENCY_SYMBOLS) roundTrip('accounting', { decimals, symbol: symbol.id })
  roundTrip('percentage', { decimals })
  roundTrip('scientific', { decimals })
}
for (const entry of DATE_FORMATS) roundTrip('date', { type: entry.id })
for (const entry of TIME_FORMATS.filter((time) => !DATE_FORMATS.some((date) => date.code === time.code))) roundTrip('time', { type: entry.id })
for (const entry of FRACTION_FORMATS) roundTrip('fraction', { type: entry.id })
for (const entry of SPECIAL_FORMATS) roundTrip('special', { type: entry.id })

// Excel's own spellings of builtin formats are recognized.
eq(detectNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'), detectNumberFormat(usAccounting))
eq(detectNumberFormat('"$"#,##0.00_);[Red]\\("$"#,##0.00\\)').category, 'currency')
eq(detectNumberFormat('"$"#,##0.00_);[Red]\\("$"#,##0.00\\)').options.negative, 'red-parens')
eq(detectNumberFormat('#,##0.00;[red]#,##0.00').options.negative, 'red', 'colour names are case-insensitive')
eq(detectNumberFormat('m/d/yyyy').category, 'date')
eq(detectNumberFormat('mmmm d, yyyy').options.type, '[$-409]mmmm d, yyyy;@', 'locale-less spelling maps to the list entry')
eq(detectNumberFormat('h:mm:ss AM/PM').category, 'time')
eq(detectNumberFormat('General').category, 'general')
eq(detectNumberFormat('').category, 'general')
eq(detectNumberFormat('@').category, 'text')
eq(detectNumberFormat('0.0,"K"').category, 'custom')
eq(detectNumberFormat('0.0,"K"').options.code, '0.0,"K"')
eq(detectNumberFormat('yyyy "Q"q').category, 'custom')
ok(isDateTimeFormat('[h]:mm:ss'))
ok(isDateTimeFormat('m/d;@'))
ok(!isDateTimeFormat('0.00E+00'))
ok(!isDateTimeFormat('#,##0.00 "days"'))
ok(!isDateTimeFormat('General'))
for (const preset of CUSTOM_FORMAT_PRESETS) ok(validateNumberFormat(preset).valid, `preset ${preset} renders`)
ok(!validateNumberFormat('0;0;0;0;0').valid, 'five sections are invalid')
ok(!validateNumberFormat('0.00"kg').valid, 'unterminated quote is invalid')
ok(!validateNumberFormat('[Red0').valid, 'unclosed bracket is invalid')
eq(formatTypeSample('date', DATE_FORMATS[0]), '*3/14/2012')
eq(formatTypeSample('time', TIME_FORMATS[0]), '*1:30:55 PM')
eq(formatTypeSample('date', DATE_FORMATS.find((entry) => entry.id === '[$-409]mmmm d, yyyy;@')!), 'March 14, 2012')
eq(formatTypeSample('time', TIME_FORMATS.find((entry) => entry.id === '[h]:mm:ss;@')!), '37:30:55', 'elapsed samples use one day plus the sample time, as in Excel')
for (const entry of [...DATE_FORMATS, ...TIME_FORMATS]) {
  eq(formatTypeSample(DATE_FORMATS.includes(entry) ? 'date' : 'time', entry), `${entry.system ? '*' : ''}${entry.label}`, `list sample for ${entry.code}`)
}

// ---------------------------------------------------------------------------------------
// Decimal increase / decrease
// ---------------------------------------------------------------------------------------

eq(adjustDecimals('0', 1), '0.0')
eq(adjustDecimals('0.00', 1), '0.000')
eq(adjustDecimals('0.0', -1), '0')
eq(adjustDecimals('0', -1), '0')
eq(adjustDecimals('#,##0.00_);[Red](#,##0.00)', -1), '#,##0.0_);[Red](#,##0.0)')
eq(adjustDecimals('$#,##0_);($#,##0)', 1), '$#,##0.0_);($#,##0.0)')
eq(adjustDecimals('0%', 1), '0.0%')
eq(adjustDecimals('0.00%', -1), '0.0%')
eq(adjustDecimals('0.00E+00', 1), '0.000E+00', 'only the mantissa gains a decimal')
eq(adjustDecimals('0E+00', 1), '0.0E+00')
eq(adjustDecimals('#,##0,"K"', 1), '#,##0.0,"K"', 'decimals go before thousands scaling')
eq(adjustDecimals('0.00" kg"', 1), '0.000" kg"')
eq(adjustDecimals('[$€-x-euro2] #,##0.00', 1), '[$€-x-euro2] #,##0.000')
eq(adjustDecimals(usAccounting, 1), '_($* #,##0.000_);_($* (#,##0.000);_($* "-"???_);_(@_)', 'accounting zero section keeps its padding in step')
eq(adjustDecimals(usAccounting, -1), '_($* #,##0.0_);_($* (#,##0.0);_($* "-"?_);_(@_)')
eq(adjustDecimals(buildNumberFormat('accounting', { decimals: 0, symbol: 'usd' }), 1), buildNumberFormat('accounting', { decimals: 1, symbol: 'usd' }))
eq(adjustDecimals(buildNumberFormat('accounting', { decimals: 1, symbol: 'usd' }), -1), buildNumberFormat('accounting', { decimals: 0, symbol: 'usd' }))
eq(adjustDecimals('General', 1, 1.5), '0.00', 'General starts from the decimals the value shows')
eq(adjustDecimals('General', -1, 1.2345), '0.000')
eq(adjustDecimals('General', 1, 7), '0.0')
eq(adjustDecimals(undefined, 1), '0.0')
eq(adjustDecimals('m/d/yyyy', 1), null)
eq(adjustDecimals('[h]:mm:ss', 1), null)
eq(adjustDecimals('# ?/?', 1), null)
eq(adjustDecimals('@', 1), null)
eq(adjustDecimals(`0.${'0'.repeat(30)}`, 1), `0.${'0'.repeat(30)}`, 'capped at 30 decimals')
for (let decimals = 0; decimals < 5; decimals += 1) {
  for (const negative of NEGATIVE_NUMBER_STYLES) {
    eq(adjustDecimals(buildNumberFormat('currency', { decimals, symbol: 'usd', negative: negative.id }), 1), buildNumberFormat('currency', { decimals: decimals + 1, symbol: 'usd', negative: negative.id }))
    eq(adjustDecimals(buildNumberFormat('number', { decimals: decimals + 1, thousands: true, negative: negative.id }), -1), buildNumberFormat('number', { decimals, thousands: true, negative: negative.id }))
  }
}

// ---------------------------------------------------------------------------------------
// applyFormatChange / applyBorderPreset semantics
// ---------------------------------------------------------------------------------------

const thin: CellBorderSide = { style: 'thin', color: { argb: 'FF000000' } }
const thick: CellBorderSide = { style: 'thick', color: { argb: 'FFFF0000' } }
const bounds = { top: 0, bottom: 2, left: 0, right: 2 }
const at = (row: number, col: number): CellPosition => cellPosition(row, col, bounds)

// Outline + inside on a 3×3 selection.
const outlineInside: FormatCellsChange = { borders: { outline: thick, insideHorizontal: thin, insideVertical: thin } }
const corner = applyFormatChange(undefined, outlineInside, at(0, 0))
eq(corner.border, { top: thick, left: thick, bottom: thin, right: thin })
const center = applyFormatChange(undefined, outlineInside, at(1, 1))
eq(center.border, { top: thin, left: thin, bottom: thin, right: thin })
const bottomRight = applyFormatChange(undefined, outlineInside, at(2, 2))
eq(bottomRight.border, { top: thin, left: thin, bottom: thick, right: thick })
eq(applyFormatChange(undefined, outlineInside, at(1, 0)).border, { top: thin, bottom: thin, left: thick, right: thin })

// Explicit edges override the outline; untouched sides survive; null removes.
const existing: CellStyle = { border: { top: thin, bottom: thin, left: thin, right: thin }, font: { bold: true } }
eq(applyFormatChange(existing, { borders: { outline: thick, bottom: null } }, at(2, 0)).border, { top: thin, left: thick, right: thin })
eq(applyFormatChange(existing, { borders: { insideHorizontal: null } }, at(1, 1)).border, { left: thin, right: thin })
eq(applyFormatChange(existing, { borders: { outline: null, insideHorizontal: null, insideVertical: null } }, at(1, 1)).border, undefined, 'all edges cleared removes the border object')
eq(existing.border?.top, thin, 'input is not mutated')

// Diagonals share one style.
const diagonal = applyFormatChange(undefined, { borders: { diagonalUp: thin, diagonalDown: thin } })
eq(diagonal.border, { diagonalUp: true, diagonal: thin, diagonalDown: true })
eq(applyFormatChange(diagonal, { borders: { diagonalUp: null, diagonalDown: null } }).border, undefined)

// Neighbour cleanup across a shared edge.
eq(applyNeighborBorderChange({ border: { bottom: thin, top: thin } }, { outline: null }, 'above'), { border: { top: thin } })
eq(applyNeighborBorderChange({ border: { left: thin } }, { right: thick }, 'right'), {})
const untouched = { border: { bottom: thin } }
eq(applyNeighborBorderChange(untouched, { left: thick }, 'above'), untouched, 'unrelated edges leave neighbours alone')

// Non-border categories.
const styled: CellStyle = { font: { name: 'Arial', size: 10, bold: true }, alignment: { horizontal: 'center', wrapText: true }, numFmt: '0.00', fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } }
const changed = applyFormatChange(styled, {
  numFmt: 'General',
  font: { bold: null, italic: true, color: { theme: 4 } },
  alignment: { horizontal: null, textRotation: 45 },
  fill: null,
  protection: { locked: false },
})
eq(changed, {
  font: { name: 'Arial', size: 10, italic: true, color: { theme: 4 } },
  alignment: { wrapText: true, textRotation: 45 },
  protection: { locked: false },
})
eq(applyFormatChange({ protection: { locked: false } }, { protection: { locked: true } }).protection, undefined, 'default protection is omitted')
eq(applyFormatChange({ alignment: { indent: 2 } }, { alignment: { indent: null } }).alignment, undefined)

// Cell-level number format.
eq(applyFormatChangeToCell({ value: 1, style: { numFmt: '0' }, numFmt: '0.0' }, { numFmt: '#,##0.00' }), { value: 1, numFmt: '#,##0.00' })
eq(applyFormatChangeToCell({ value: 1, numFmt: '0.0' }, { numFmt: 'General' }), { value: 1 })
eq(applyFormatChangeToCell(undefined, { font: { bold: true } }), { style: { font: { bold: true } } })

// Border presets (Google Sheets semantics).
eq(applyBorderPreset('all', thin, at(1, 1)).border, { top: thin, bottom: thin, left: thin, right: thin })
eq(applyBorderPreset('outer', thin, at(1, 1)).border, undefined, 'interior cells get no outer border')
eq(applyBorderPreset('outer', thin, at(0, 1)).border, { top: thin })
eq(applyBorderPreset('inner', thin, at(0, 0)).border, { bottom: thin, right: thin })
eq(applyBorderPreset('horizontal', thin, at(0, 2)).border, { bottom: thin })
eq(applyBorderPreset('vertical', thin, at(2, 0)).border, { right: thin })
eq(applyBorderPreset('left', thin, at(1, 0)).border, { left: thin })
eq(applyBorderPreset('left', thin, at(1, 1)).border, undefined)
eq(applyBorderPreset('top', thin, at(0, 2)).border, { top: thin })
eq(applyBorderPreset('right', thin, at(2, 2)).border, { right: thin })
eq(applyBorderPreset('bottom', thin, at(2, 1)).border, { bottom: thin })
eq(applyBorderPreset('clear', thin, at(1, 1), { border: { top: thick, diagonalUp: true, diagonal: thin }, font: { bold: true } }), { font: { bold: true } })
eq(applyBorderPreset('all', thick, { top: true, bottom: true, left: true, right: true }, { border: { top: thin } }).border, { top: thick, bottom: thick, left: thick, right: thick })

// ---------------------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------------------

const near = (actual: string, expected: string, tolerance = 2) => {
  const channels = (hex: string) => [1, 3, 5].map((offset) => parseInt(hex.replace('#', '').slice(offset - 1, offset + 1), 16))
  const a = channels(actual)
  const b = channels(expected)
  ok(a.every((value, index) => Math.abs(value - b[index]) <= tolerance), `${actual} ≈ ${expected}`)
}
near(`#${applyTint('4472C4', 0.8)}`, '#D9E1F2') // Blue, Accent 1, Lighter 80%
near(`#${applyTint('4472C4', 0.6)}`, '#B4C6E7')
near(`#${applyTint('4472C4', 0.4)}`, '#8EA9DB')
near(`#${applyTint('4472C4', -0.25)}`, '#305496')
near(`#${applyTint('4472C4', -0.5)}`, '#203764')
near(`#${applyTint('FFFFFF', -0.05)}`, '#F2F2F2')
near(`#${applyTint('000000', 0.5)}`, '#808080')
near(`#${applyTint('E7E6E6', -0.1)}`, '#D0CECE')
eq(themeTintSteps('FFFFFF'), [-0.05, -0.15, -0.25, -0.35, -0.5])
eq(themeTintSteps('000000'), [0.5, 0.35, 0.25, 0.15, 0.05])
eq(themeTintSteps('E7E6E6'), [-0.1, -0.25, -0.5, -0.75, -0.9])
eq(themeTintSteps('44546A'), [0.8, 0.6, 0.4, -0.25, -0.5])
const palette = themePalette()
eq(palette.length, 6)
ok(palette.every((row) => row.length === 10))
eq(palette[0][4].color, { theme: 4 })
eq(palette[1][4].color, { theme: 4, tint: 0.8 })
eq(resolveColorHex({ theme: 4 }), '#4472C4')
eq(resolveColorHex({ argb: 'FF123456' }), '#123456')
eq(resolveColorHex({ indexed: 10 }), '#FF0000')
eq(resolveColorHex({ theme: 4 }, ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD']), '#4F81BD', 'workbook theme palette is honoured')
eq(resolveColorHex(null), '')
ok(sameColor({ theme: 4, tint: 0.79998 }, { theme: 4, tint: 0.8 }))
ok(!sameColor({ theme: 4 }, { argb: 'FF4472C4' }), 'theme and RGB colours are different choices')
ok(sameColor('#FF0000', { argb: 'FFFF0000' }))
ok(patternTileUri('darkGrid', '#000000').startsWith('url("data:image/svg+xml,'))
eq(fillPreviewCss({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }), { backgroundColor: '#FFFF00' })
eq(fillPreviewCss(null), {})

// ---------------------------------------------------------------------------------------
// Cell styles
// ---------------------------------------------------------------------------------------

const style = (id: string) => CELL_STYLES.find((preset) => preset.id === id)!
const expectedNames = ['Normal', 'Bad', 'Good', 'Neutral', 'Calculation', 'Check Cell', 'Explanatory Text', 'Input', 'Linked Cell', 'Note', 'Output', 'Warning Text', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Title', 'Total', 'Comma', 'Comma [0]', 'Currency', 'Currency [0]', 'Percent']
for (const name of expectedNames) ok(CELL_STYLES.some((preset) => preset.name === name), `cell style ${name}`)
for (let accent = 1; accent <= 6; accent += 1) {
  for (const prefix of ['20% - ', '40% - ', '60% - ', '']) ok(CELL_STYLES.some((preset) => preset.name === `${prefix}Accent${accent}`), `${prefix}Accent${accent}`)
}
eq(new Set(CELL_STYLES.map((preset) => preset.id)).size, CELL_STYLES.length, 'style ids are unique')
eq(style('good').style, { font: { color: { argb: 'FF006100' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } } })
const base: CellStyle = { font: { name: 'Arial', size: 10, italic: true }, border: { top: thin }, numFmt: '0.00', alignment: { horizontal: 'right' } }
eq(applyCellStylePreset(base, style('good')), {
  font: { name: 'Arial', size: 10, color: { argb: 'FF006100' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } },
  border: { top: thin },
  numFmt: '0.00',
  alignment: { horizontal: 'right' },
}, 'Good replaces font and fill only; the typeface survives')
eq(applyCellStylePreset(base, style('heading1')).font, { name: 'Arial', size: 15, bold: true, color: { theme: 3 } })
eq(applyCellStylePreset(base, style('heading1')).border, { bottom: { style: 'thick', color: { theme: 4 } } })
eq(applyCellStylePreset(base, style('normal')), { font: { name: 'Arial', size: 10 } }, 'Normal resets everything but the typeface')
eq(applyCellStylePresetToCell({ value: 5, numFmt: '0.0', style: base }, style('percent')).numFmt, '0%')
eq(applyCellStylePresetToCell({ value: 5, numFmt: '0.0', style: base }, style('percent')).style?.numFmt, undefined)
eq(applyCellStylePreset(undefined, style('title')).font?.name, 'Aptos Display')

// ---------------------------------------------------------------------------------------
// Fonts
// ---------------------------------------------------------------------------------------

ok(FONT_FAMILIES.length >= 60, 'at least 60 curated fonts')
for (const name of ['Aptos', 'Calibri', 'Cambria', 'Arial', 'Arial Black', 'Segoe UI', 'Times New Roman', 'Georgia', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Consolas', 'Courier New', 'Garamond', 'Century Gothic', 'Book Antiqua', 'Palatino Linotype', 'Lucida Console', 'Comic Sans MS', 'Impact', 'Candara', 'Constantia', 'Corbel', 'Ebrima', 'Gabriola', 'Leelawadee UI', 'Malgun Gothic', 'Microsoft YaHei', 'Yu Gothic', 'MS Gothic', 'SimSun', 'Nirmala UI']) {
  ok(FONT_FAMILIES.some((font) => font.name === name), `font ${name}`)
}
eq(new Set(FONT_FAMILIES.map((font) => font.name)).size, FONT_FAMILIES.length, 'font names are unique')
eq(fontStack('Consolas'), '"Consolas", Consolas, "Courier New", monospace')
eq(parseFontSize('10.3'), 10.5)
eq(parseFontSize('0'), null)
eq(parseFontSize('410'), null)
eq(parseFontSize('11'), 11)

process.stdout.write(`Format-codes QA passed: ${checks} checks (sections, colours, conditions, fractions, elapsed time, accounting, build/detect round trips, decimals, border and style semantics).\n`)
