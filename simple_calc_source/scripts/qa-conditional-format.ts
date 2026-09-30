/**
 * Conditional formatting QA: evaluation of every rule type, priority/stopIfTrue merging,
 * relative formula shifting, range statistics, authoring/management helpers, the editor
 * draft model, performance, and ExcelJS / workbook-pipeline round trips.
 *
 *   esbuild scripts/qa-conditional-format.ts --bundle --platform=node --format=esm --packages=external --outfile=tmp/qa-conditional-format.mjs && node tmp/qa-conditional-format.mjs
 *
 * `--packages=external` keeps ExcelJS a single runtime instance, so the fidelity patches in
 * electron/conditional-format-exceljs.cjs apply to the same classes this script uses.
 */
import assert from 'node:assert/strict'
import { createRequire } from 'node:module'
import { performance } from 'node:perf_hooks'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { evaluateFormula, shiftFormulaReferences } from '../src/lib/formulas'
import {
  CONDITIONAL_STYLE_PRESETS,
  ICON_SET_NAMES,
  addRule,
  cellAddress,
  clearRulesInRange,
  computeConditionalFormats,
  conditionalPresetStyle,
  conditionalRuleKind,
  createAverageRule,
  createBlankRule,
  createCellIsRule,
  createColorScaleRule,
  createConditionalStyle,
  createDataBarRule,
  createDateRule,
  createDuplicateRule,
  createErrorRule,
  createFormulaRule,
  createIconSetRule,
  createTextRule,
  createTopBottomRule,
  dataBarBackground,
  deleteRule,
  listRulesForRange,
  matchConditionalStylePreset,
  moveRulePriority,
  normalizeConditionalFormattings,
  parseRangeInput,
  shiftConditionalFormula,
  updateRule,
} from '../src/lib/conditional-format'
import type { ConditionalFormatHost, ConditionalRule, ConditionalScalar, TimePeriod } from '../src/lib/conditional-format'
import { CONDITION_OPTIONS, defaultRuleDraft, describeRule, draftToRule, ruleToDraft, validateDraft, withIconSet } from '../src/lib/conditional-format-editor'
import type { RuleDraft } from '../src/lib/conditional-format-editor'

const require = createRequire(import.meta.url)
const { installConditionalFormattingPatches } = require('../electron/conditional-format-exceljs.cjs') as { installConditionalFormattingPatches: () => boolean }

type Cells = Record<string, ConditionalScalar>

let passed = 0
const failures: string[] = []

async function test(name: string, body: () => void | Promise<void>) {
  try {
    await body()
    passed += 1
  } catch (error) {
    failures.push(`${name}\n    ${error instanceof Error ? error.stack?.split('\n').slice(0, 3).join('\n    ') : String(error)}`)
  }
}

function host(cells: Cells, options: { today?: number; rowCount?: number; colCount?: number; counter?: { evaluate: number; value: number } } = {}): ConditionalFormatHost {
  const upper: Cells = {}
  for (const [key, value] of Object.entries(cells)) upper[key.toUpperCase()] = value
  const resolver = (_sheet: string, address: string) => {
    const value = upper[address.toUpperCase()]
    if (typeof value === 'string' && /^#(N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!)$/.test(value)) return { error: value as '#N/A' }
    return value ?? null
  }
  return {
    sheetId: 's1',
    today: options.today ?? 45_308,
    rowCount: options.rowCount,
    colCount: options.colCount,
    valueAt(row, col) {
      if (options.counter) options.counter.value += 1
      return upper[cellAddress(row, col)] ?? null
    },
    evaluate(formula, row, col) {
      if (options.counter) options.counter.evaluate += 1
      return evaluateFormula(formula, 's1', resolver, { currentCell: { row: row + 1, column: col + 1 } })
    },
  }
}

function plain<T>(value: T): T {
  return value === undefined ? value : JSON.parse(JSON.stringify(value))
}

function block(ref: string, ...rules: ConditionalRule[]) {
  return { ref, rules }
}

const RED_FILL = createConditionalStyle({ fillColor: '#FFC7CE' })
const BOLD = createConditionalStyle({ bold: true })
const GREEN_FILL = createConditionalStyle({ fillColor: '#C6EFCE' })

function matched(result: Map<string, unknown>, addresses: string[]) {
  return addresses.filter((address) => result.has(address))
}

function column(values: ConditionalScalar[], letter = 'A', start = 1): Cells {
  const cells: Cells = {}
  values.forEach((value, index) => { cells[`${letter}${start + index}`] = value })
  return cells
}

function fills(result: Map<string, { fill?: string }>, addresses: string[]) {
  return addresses.map((address) => result.get(address)?.fill ?? null)
}

// ---------------------------------------------------------------------------------------------
// Relative formula shifting
// ---------------------------------------------------------------------------------------------

await test('relative formula shifting matches shiftFormulaReferences', () => {
  const formulas = [
    'A1>5', '$A1>$B$1', 'A$1+B2*C3', 'SUM(A1:B2)>10', 'AND(A1<>"",A1="x A1")', 'COUNTIF($A:$A,A1)>1', 'SUM(A:C)', 'SUM(1:3)',
    "'My Sheet'!A1+Sheet2!B2", 'LOG10(A1)', 'a1+b$2', 'MOD(ROW(),2)=0', 'Table1[Col]+A1', 'ISNUMBER(SEARCH("B2",C3))',
    'IF(A1>0,A1,-A1)', 'Sales2024+A2', 'XFD1+A1', 'SUM($A$1:A1)', 'INDEX($A:$A,ROW())', "'A1 x'!C3", 'Sheet1!$A1:B$2',
  ]
  const offsets = [[0, 0], [1, 0], [0, 1], [5, 3], [100, 26], [2, 700]]
  for (const formula of formulas) {
    for (const [rowDelta, colDelta] of offsets) {
      const expected = shiftFormulaReferences(formula, rowDelta, colDelta)
      if (expected.includes('#REF!')) continue
      assert.equal(shiftConditionalFormula(formula, rowDelta, colDelta), expected, `${formula} shifted by ${rowDelta},${colDelta}`)
    }
  }
})

await test('relative references wrap around sheet edges like Excel', () => {
  assert.equal(shiftConditionalFormula('A1', -1, 0), 'A1048576')
  assert.equal(shiftConditionalFormula('A1', 0, -1), 'XFD1')
  assert.equal(shiftConditionalFormula('$A$1+B2', -5, -5), '$A$1+XFA1048573')
})

// ---------------------------------------------------------------------------------------------
// cellIs
// ---------------------------------------------------------------------------------------------

await test('cellIs operators with literals', () => {
  const cells = column([1, 5, 10, 15, 20])
  const addresses = ['A1', 'A2', 'A3', 'A4', 'A5']
  const run = (operator: Parameters<typeof createCellIsRule>[0], values: Array<string | number>) =>
    matched(computeConditionalFormats([block('A1:A5', createCellIsRule(operator, values, { style: RED_FILL }))], host(cells)), addresses)
  assert.deepEqual(run('greaterThan', [10]), ['A4', 'A5'])
  assert.deepEqual(run('greaterThanOrEqual', [10]), ['A3', 'A4', 'A5'])
  assert.deepEqual(run('lessThan', [10]), ['A1', 'A2'])
  assert.deepEqual(run('lessThanOrEqual', [10]), ['A1', 'A2', 'A3'])
  assert.deepEqual(run('equal', [10]), ['A3'])
  assert.deepEqual(run('notEqual', [10]), ['A1', 'A2', 'A4', 'A5'])
  assert.deepEqual(run('between', [5, 15]), ['A2', 'A3', 'A4'])
  assert.deepEqual(run('between', [15, 5]), ['A2', 'A3', 'A4'], 'between is order-insensitive')
  assert.deepEqual(run('notBetween', [5, 15]), ['A1', 'A5'])
})

await test('cellIs Excel comparison semantics (text, blanks, types, errors)', () => {
  const cells: Cells = { A1: 'Apple', A2: 'apple', A3: null, A4: 'zebra', A5: true, A6: '#N/A', A7: 0.1 + 0.2 }
  const addresses = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7']
  const run = (rule: ConditionalRule, ref = 'A1:A7') => matched(computeConditionalFormats([block(ref, rule)], host(cells)), addresses)
  assert.deepEqual(run(createCellIsRule('equal', ['"APPLE"'], { style: RED_FILL })), ['A1', 'A2'], 'text equality is case-insensitive')
  // Blank compares as 0 with numbers; text sorts after numbers and logicals after text.
  assert.deepEqual(run(createCellIsRule('lessThan', [1], { style: RED_FILL })), ['A3', 'A7'])
  assert.deepEqual(run(createCellIsRule('greaterThan', [1000], { style: RED_FILL })), ['A1', 'A2', 'A4', 'A5'])
  assert.deepEqual(run(createCellIsRule('equal', [0.3], { style: RED_FILL })), ['A7'], '15-digit equality')
  assert.deepEqual(run(createCellIsRule('notEqual', ['"x"'], { style: RED_FILL })).includes('A6'), false, 'errors never match')
})

await test('cellIs operands are formulas: absolute and relative references', () => {
  const cells: Cells = { A1: 1, A2: 5, A3: 9, B1: 2, B2: 4, B3: 10, D1: 5 }
  const addresses = ['A1', 'A2', 'A3']
  const absolute = computeConditionalFormats([block('A1:A3', createCellIsRule('greaterThanOrEqual', ['$D$1'], { style: RED_FILL }))], host(cells))
  assert.deepEqual(matched(absolute, addresses), ['A2', 'A3'])
  const relative = computeConditionalFormats([block('A1:A3', createCellIsRule('greaterThan', ['=B1'], { style: RED_FILL }))], host(cells))
  assert.deepEqual(matched(relative, addresses), ['A2'])
  const math = computeConditionalFormats([block('A1:A3', createCellIsRule('between', ['B1-1', 'B1+1'], { style: RED_FILL }))], host(cells))
  assert.deepEqual(matched(math, addresses), ['A1', 'A2', 'A3'], 'each cell compared with its own B-1 .. B+1')
})

// ---------------------------------------------------------------------------------------------
// expression
// ---------------------------------------------------------------------------------------------

await test('expression rules shift relative references from the top-left cell', () => {
  const cells: Cells = { B2: 5, B3: 50, C2: 10, C3: 20, B4: 1, C4: 1 }
  const result = computeConditionalFormats([block('B2:B4', createFormulaRule('=B2>C2', { style: RED_FILL }))], host(cells))
  assert.deepEqual(matched(result, ['B2', 'B3', 'B4']), ['B3'])
  const striped = computeConditionalFormats([block('A1:A6', createFormulaRule('MOD(ROW(),2)=0', { style: RED_FILL }))], host({}))
  assert.deepEqual(matched(striped, ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']), ['A2', 'A4', 'A6'])
  const mixed = computeConditionalFormats([block('A1:C2', createFormulaRule('$A1="x"', { style: BOLD }))], host({ A1: 'x', A2: 'y' }))
  assert.deepEqual(matched(mixed, ['A1', 'B1', 'C1', 'A2', 'B2', 'C2']), ['A1', 'B1', 'C1'])
  const text = computeConditionalFormats([block('A1', createFormulaRule('"text"', { style: BOLD }))], host({}))
  assert.equal(text.size, 0, 'text results are not TRUE')
  const numeric = computeConditionalFormats([block('A1', createFormulaRule('2', { style: BOLD }))], host({}))
  assert.equal(numeric.size, 1, 'non-zero numbers are TRUE')
})

await test('multi-range refs anchor on the first range', () => {
  const cells: Cells = { C5: 1, C6: 2, A1: 1, D5: 1, D6: 9, B1: 1 }
  // Formula written for C5 compares with the cell to its right.
  const result = computeConditionalFormats([block('C5:C6 A1', createFormulaRule('C5=D5', { style: BOLD }))], host(cells))
  assert.deepEqual(matched(result, ['C5', 'C6', 'A1']), ['C5', 'A1'])
})

// ---------------------------------------------------------------------------------------------
// Text, blanks, errors, dates
// ---------------------------------------------------------------------------------------------

await test('text rules: contains / not contains / begins / ends (case-insensitive, wildcards)', () => {
  const cells: Cells = { A1: 'Invoice 2024', A2: 'receipt', A3: 12345, A4: '#N/A', A5: null, A6: 'INV-99' }
  const addresses = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']
  const run = (rule: ConditionalRule) => matched(computeConditionalFormats([block('A1:A6', rule)], host(cells)), addresses)
  assert.deepEqual(run(createTextRule('containsText', 'inv', { style: RED_FILL, ref: 'A1:A6' })), ['A1', 'A6'])
  assert.deepEqual(run(createTextRule('containsText', '23', { style: RED_FILL })), ['A3'], 'numbers are searched as text')
  assert.deepEqual(run(createTextRule('containsText', 'in*4', { style: RED_FILL })), ['A1'], '* wildcard')
  assert.deepEqual(run(createTextRule('containsText', 'r?c', { style: RED_FILL })), ['A2'], '? wildcard')
  assert.deepEqual(run(createTextRule('notContainsText', 'inv', { style: RED_FILL })), ['A2', 'A3', 'A4', 'A5'], 'errors do not contain text')
  assert.deepEqual(run(createTextRule('beginsWith', 'INV', { style: RED_FILL })), ['A1', 'A6'])
  assert.deepEqual(run(createTextRule('endsWith', '99', { style: RED_FILL })), ['A6'])
  // A text rule without `text` falls back to its formula.
  const formulaOnly: ConditionalRule = { type: 'containsText', operator: 'containsText', formulae: ['NOT(ISERROR(SEARCH("rec",A1)))'], style: RED_FILL }
  assert.deepEqual(run(formulaOnly), ['A2'])
})

await test('blank and error rules', () => {
  const cells: Cells = { A1: null, A2: '   ', A3: 'x', A4: '#DIV/0!', A5: 0 }
  const addresses = ['A1', 'A2', 'A3', 'A4', 'A5']
  const run = (rule: ConditionalRule) => matched(computeConditionalFormats([block('A1:A5', rule)], host(cells)), addresses)
  assert.deepEqual(run(createBlankRule(true, { style: RED_FILL })), ['A1', 'A2'])
  assert.deepEqual(run(createBlankRule(false, { style: RED_FILL })), ['A3', 'A5'])
  assert.deepEqual(run(createErrorRule(true, { style: RED_FILL })), ['A4'])
  assert.deepEqual(run(createErrorRule(false, { style: RED_FILL })), ['A1', 'A2', 'A3', 'A5'])
  // ExcelJS-native type names are accepted too.
  assert.deepEqual(run({ type: 'containsBlanks', style: RED_FILL }), ['A1', 'A2'])
})

await test('time period rules against a fixed today (Wed 2024-01-17 = 45308)', () => {
  const today = 45_308
  const serial = (offset: number) => today + offset
  const cells: Cells = {}
  const offsets = [-40, -31, -17, -10, -8, -7, -6, -3, -1, 0, 0.75, 1, 3, 4, 5, 11, 12, 15, 20, 45]
  offsets.forEach((offset, index) => { cells[`A${index + 1}`] = serial(offset) })
  cells.A30 = 'today'
  const addresses = Object.keys(cells)
  const run = (period: TimePeriod) => matched(computeConditionalFormats([block('A1:A30', createDateRule(period, { style: RED_FILL }))], host(cells, { today })), addresses)
    .map((address) => (typeof cells[address] === 'number' ? Number(cells[address]) - today : cells[address]))
  assert.deepEqual(run('today'), [0, 0.75])
  assert.deepEqual(run('yesterday'), [-1])
  assert.deepEqual(run('tomorrow'), [1])
  assert.deepEqual(run('last7Days'), [-6, -3, -1, 0, 0.75])
  assert.deepEqual(run('thisWeek'), [-3, -1, 0, 0.75, 1, 3], 'Sunday 14th .. Saturday 20th')
  assert.deepEqual(run('lastWeek'), [-10, -8, -7, -6], 'Sunday 7th .. Saturday 13th')
  assert.deepEqual(run('nextWeek'), [4, 5], 'Sunday 21st .. Saturday 27th')
  assert.deepEqual(run('thisMonth'), [-10, -8, -7, -6, -3, -1, 0, 0.75, 1, 3, 4, 5, 11, 12], 'January 2024')
  assert.deepEqual(run('lastMonth'), [-40, -31, -17].filter((offset) => today + offset >= 45_261 && today + offset <= 45_291))
  assert.deepEqual(run('nextMonth'), [15, 20].filter((offset) => today + offset >= 45_323))
})

// ---------------------------------------------------------------------------------------------
// Range statistics rules
// ---------------------------------------------------------------------------------------------

await test('top/bottom N and N% with ties', () => {
  const values = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 100, 'x', null]
  const cells = column(values)
  const addresses = values.map((_, index) => `A${index + 1}`)
  const run = (config: Parameters<typeof createTopBottomRule>[0]) =>
    matched(computeConditionalFormats([block('A1:A13', createTopBottomRule(config, { style: RED_FILL }))], host(cells)), addresses).map((address) => cells[address])
  assert.deepEqual(run({ rank: 2 }), [100, 100])
  assert.deepEqual(run({ rank: 3 }), [90, 100, 100])
  assert.deepEqual(run({ rank: 2, bottom: true }), [10, 20])
  assert.deepEqual(run({ rank: 20, percent: true }), [100, 100], '20% of 11 numbers = 2')
  assert.deepEqual(run({ rank: 1, percent: true }), [100, 100], 'at least one item (ties included)')
  assert.deepEqual(run({ rank: 30, percent: true, bottom: true }), [10, 20, 30])
})

await test('above/below average, equal, standard deviations (population)', () => {
  const values = [2, 4, 4, 4, 5, 5, 7, 9] // mean 5, population sd 2
  const cells = column(values)
  const addresses = values.map((_, index) => `A${index + 1}`)
  const run = (config: Parameters<typeof createAverageRule>[0]) =>
    matched(computeConditionalFormats([block('A1:A8', createAverageRule(config, { style: RED_FILL }))], host(cells)), addresses).map((address) => cells[address])
  assert.deepEqual(run({}), [7, 9])
  assert.deepEqual(run({ below: true }), [2, 4, 4, 4])
  assert.deepEqual(run({ equal: true }), [5, 5, 7, 9])
  assert.deepEqual(run({ below: true, equal: true }), [2, 4, 4, 4, 5, 5])
  assert.deepEqual(run({ stdDev: 1 }), [9], 'value > mean + 1 sd (7)')
  assert.deepEqual(run({ stdDev: 1, below: true }), [2], 'value < mean - 1 sd (3)')
})

await test('duplicate and unique values (case-insensitive, blanks ignored)', () => {
  const cells: Cells = { A1: 'Apple', A2: 'apple', A3: 'pear', A4: 1, A5: 1, A6: '1', A7: null, A8: null, A9: true }
  const addresses = Object.keys(cells)
  const duplicate = matched(computeConditionalFormats([block('A1:A9', createDuplicateRule(false, { style: RED_FILL }))], host(cells)), addresses)
  assert.deepEqual(duplicate, ['A1', 'A2', 'A4', 'A5'])
  const unique = matched(computeConditionalFormats([block('A1:A9', createDuplicateRule(true, { style: RED_FILL }))], host(cells)), addresses)
  assert.deepEqual(unique, ['A3', 'A6', 'A9'])
})

// ---------------------------------------------------------------------------------------------
// Priority, stopIfTrue and merging
// ---------------------------------------------------------------------------------------------

await test('priority order, property merging and stopIfTrue', () => {
  const cells = column([5, 50, 500])
  const rules = [
    block('A1:A3', { ...createCellIsRule('greaterThan', [10], { style: createConditionalStyle({ fillColor: '#FF0000', bold: true }) }), priority: 2 }),
    block('A1:A3', { ...createCellIsRule('greaterThan', [100], { style: createConditionalStyle({ fillColor: '#00FF00', italic: true }) }), priority: 1 }),
    block('A1:A3', { ...createCellIsRule('greaterThan', [0], { style: createConditionalStyle({ fontColor: '#0000FF', fillColor: '#999999' }) }), priority: 3 }),
  ]
  const result = computeConditionalFormats(rules, host(cells))
  assert.deepEqual(result.get('A3'), { fill: '#00FF00', font: { italic: true, bold: true, color: '#0000FF' } }, 'higher priority wins; others add unset properties')
  assert.deepEqual(result.get('A2'), { fill: '#FF0000', font: { bold: true, color: '#0000FF' } })
  assert.deepEqual(result.get('A1'), { fill: '#999999', font: { color: '#0000FF' } })

  const stopping = [
    rules[0],
    block('A1:A3', { ...rules[1].rules[0], stopIfTrue: true }),
    rules[2],
  ]
  const stopped = computeConditionalFormats(stopping, host(cells))
  assert.deepEqual(stopped.get('A3'), { fill: '#00FF00', font: { italic: true } }, 'stopIfTrue blocks lower rules')
  assert.deepEqual(stopped.get('A2'), { fill: '#FF0000', font: { bold: true, color: '#0000FF' } })

  // A rule with no format still stops lower rules (Excel's "stop if true" trick).
  const silent = computeConditionalFormats([
    block('A1:A3', { ...createCellIsRule('equal', [5]), priority: 1, stopIfTrue: true }),
    block('A1:A3', { ...createCellIsRule('greaterThan', [0], { style: RED_FILL }), priority: 2 }),
  ], host(cells))
  assert.deepEqual(matched(silent, ['A1', 'A2', 'A3']), ['A2', 'A3'])

  // Rules without a priority come after prioritised ones, in document order.
  const unprioritised = computeConditionalFormats([
    block('A1', { type: 'expression', formulae: ['TRUE'], style: createConditionalStyle({ fillColor: '#111111' }) }),
    block('A1', { type: 'expression', formulae: ['TRUE'], priority: 4, style: createConditionalStyle({ fillColor: '#222222' }) }),
  ], host(cells))
  assert.equal(unprioritised.get('A1')?.fill, '#222222')
})

await test('dxf style resolution: bgColor for solid fills, borders, numFmt, theme colours via host', () => {
  const style = {
    font: { color: { theme: 4 }, underline: 'double', strike: true },
    fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFEB9C' } },
    border: { top: { style: 'thin', color: { argb: 'FF9C0006' } }, bottom: { style: 'medium' } },
    numFmt: '0.0%',
  }
  const cssColor = (color: unknown, fallback = '') => {
    const record = color as { theme?: number; argb?: string } | undefined
    if (record?.theme === 4) return '#123456'
    if (record?.argb) return `#${record.argb.slice(2)}`
    return fallback
  }
  const withHost = { ...host({ A1: 1 }), cssColor }
  const result = computeConditionalFormats([block('A1', { type: 'expression', formulae: ['TRUE'], style: style as never })], withHost)
  assert.deepEqual(result.get('A1'), {
    font: { color: '#123456', underline: true, strike: true },
    fill: '#FFEB9C',
    border: { top: '1px solid #9C0006', bottom: '2px solid #000000' },
    numFmt: '0.0%',
  })
})

// ---------------------------------------------------------------------------------------------
// Colour scales, data bars, icon sets
// ---------------------------------------------------------------------------------------------

await test('2- and 3-colour scales with every cfvo type', () => {
  const cells = column([0, 25, 50, 75, 100, 'text', null])
  cells.D1 = 100
  const two = computeConditionalFormats([block('A1:A7', createColorScaleRule([{ type: 'min', color: '#000000' }, { type: 'max', color: '#FFFFFF' }]))], host(cells))
  assert.deepEqual(fills(two, ['A1', 'A2', 'A3', 'A5', 'A6', 'A7']), ['#000000', '#404040', '#808080', '#FFFFFF', null, null])
  assert.equal(two.get('A3')?.colorScale, '#808080')
  const three = computeConditionalFormats([block('A1:A7', createColorScaleRule([
    { type: 'min', color: '#F8696B' }, { type: 'percentile', value: 50, color: '#FFEB84' }, { type: 'max', color: '#63BE7B' },
  ]))], host(cells))
  assert.deepEqual(fills(three, ['A1', 'A3', 'A5']), ['#F8696B', '#FFEB84', '#63BE7B'])
  const numeric = computeConditionalFormats([block('A1:A5', createColorScaleRule([{ type: 'num', value: 50, color: '#000000' }, { type: 'formula', value: '$D$1', color: '#FFFFFF' }]))], host(cells))
  assert.deepEqual(fills(numeric, ['A1', 'A3', 'A4', 'A5']), ['#000000', '#000000', '#808080', '#FFFFFF'], 'num 50 .. formula 100, clamped')
  const percent = computeConditionalFormats([block('A1:A5', createColorScaleRule([{ type: 'percent', value: 20, color: '#000000' }, { type: 'percent', value: 60, color: '#FFFFFF' }]))], host(cells))
  assert.deepEqual(fills(percent, ['A1', 'A2', 'A3', 'A4']), ['#000000', '#202020', '#BFBFBF', '#FFFFFF'])
  // A higher-priority fill wins over a lower-priority scale.
  const layered = computeConditionalFormats([
    block('A1:A5', { ...createCellIsRule('equal', [50], { style: GREEN_FILL }), priority: 1 }),
    block('A1:A5', { ...createColorScaleRule([{ type: 'min', color: '#000000' }, { type: 'max', color: '#FFFFFF' }]), priority: 2 }),
  ], host(cells))
  assert.equal(layered.get('A3')?.fill, '#C6EFCE')
  assert.equal(layered.get('A3')?.colorScale, undefined)
  assert.equal(layered.get('A1')?.fill, '#000000')
})

await test('data bars: automatic bounds, lengths, negatives, axis, options', () => {
  const positive = computeConditionalFormats([block('A1:A4', createDataBarRule({ color: '#638EC6' }))], host(column([1, 5, 10, 'x'])))
  assert.equal(positive.get('A1')?.dataBar?.fraction, 0.1, 'autoMin is 0 for positive data')
  assert.equal(positive.get('A3')?.dataBar?.fraction, 1)
  assert.equal(positive.get('A1')?.dataBar?.axis, undefined)
  assert.equal(positive.get('A1')?.dataBar?.gradient, true)
  assert.equal(positive.get('A1')?.dataBar?.border, '#638EC6')
  assert.equal(positive.has('A4'), false)

  const lowest = computeConditionalFormats([block('A1:A3', createDataBarRule({ min: { type: 'min' }, max: { type: 'max' }, gradient: false, showValue: false }))], host(column([1, 5, 9])))
  assert.equal(lowest.get('A1')?.dataBar?.fraction, 0)
  assert.equal(lowest.get('A2')?.dataBar?.fraction, 0.5)
  assert.equal(lowest.get('A2')?.dataBar?.gradient, false)
  assert.equal(lowest.get('A2')?.dataBar?.showValue, false)
  assert.equal(lowest.get('A2')?.dataBar?.border, undefined)

  const legacy = computeConditionalFormats([block('A1:A3', { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } })], host(column([1, 5, 9])))
  assert.ok(Math.abs((legacy.get('A1')?.dataBar?.fraction ?? 0) - 0.1) < 1e-9, 'Excel 2007 bars default to 10%..90%')
  assert.ok(Math.abs((legacy.get('A3')?.dataBar?.fraction ?? 0) - 0.9) < 1e-9)

  const mixed = computeConditionalFormats([block('A1:A3', createDataBarRule({ negativeColor: '#FF0000' }))], host(column([-10, 30, 0])))
  const negative = mixed.get('A1')?.dataBar
  assert.equal(negative?.negative, true)
  assert.equal(negative?.color, '#FF0000')
  assert.equal(negative?.axis, 0.25)
  assert.equal(negative?.fraction, 0.25)
  assert.equal(mixed.get('A2')?.dataBar?.fraction, 0.75)
  assert.equal(mixed.get('A3')?.dataBar?.fraction, 0)

  const middle = computeConditionalFormats([block('A1:A2', createDataBarRule({ axisPosition: 'middle' }))], host(column([-10, 30])))
  assert.equal(middle.get('A1')?.dataBar?.axis, 0.5)
  assert.equal(middle.get('A1')?.dataBar?.fraction, 0.5)
  assert.equal(middle.get('A2')?.dataBar?.fraction, 0.5)

  const background = dataBarBackground(mixed.get('A1')!.dataBar!)
  assert.match(background.backgroundImage, /linear-gradient/)
  assert.equal(background.backgroundImage.split('gradient(').length - 1, 3 + 0, 'axis + fill + border layers')
})

await test('icon sets: default thresholds, gte, reverse, showValue, number/percentile/formula, custom icons', () => {
  const values = [0, 10, 33, 50, 67, 90, 100]
  const cells = column(values)
  cells.D1 = 60
  const addresses = values.map((_, index) => `A${index + 1}`)
  const icons = (rule: ConditionalRule) => {
    const result = computeConditionalFormats([block('A1:A7', rule)], host(cells))
    return addresses.map((address) => result.get(address)?.icon?.index ?? null)
  }
  assert.deepEqual(icons(createIconSetRule('3Arrows')), [0, 0, 1, 1, 2, 2, 2], 'percent 33 / 67, >=')
  assert.deepEqual(icons(createIconSetRule('3Arrows', { thresholds: [{ type: 'percent', value: 33, gte: false }, { type: 'percent', value: 67, gte: false }] })), [0, 0, 0, 1, 1, 2, 2])
  assert.deepEqual(icons(createIconSetRule('3Arrows', { reverse: true })), [2, 2, 1, 1, 0, 0, 0])
  assert.deepEqual(icons(createIconSetRule('5Quarters')), [0, 0, 1, 2, 3, 4, 4])
  assert.deepEqual(icons(createIconSetRule('4Rating', { thresholds: [{ type: 'num', value: 10 }, { type: 'percentile', value: 50 }, { type: 'formula', value: '$D$1' }] })), [0, 1, 1, 2, 3, 3, 3])
  const hidden = computeConditionalFormats([block('A1:A7', createIconSetRule('3Flags', { showValue: false }))], host(cells))
  assert.deepEqual(hidden.get('A1')?.icon, { set: '3Flags', index: 0, showValue: false })
  const custom = computeConditionalFormats([block('A1:A7', { ...createIconSetRule('3Arrows'), icons: [{ iconSet: '3Flags', iconId: 0 }, { iconSet: 'NoIcons', iconId: 0 }, { iconSet: '3Stars', iconId: 2 }] })], host(cells))
  assert.deepEqual(custom.get('A1')?.icon, { set: '3Flags', index: 0, showValue: true })
  assert.equal(custom.get('A3')?.icon, undefined, 'NoIcons')
  assert.deepEqual(custom.get('A7')?.icon, { set: '3Stars', index: 2, showValue: true })
  // ExcelJS reads a missing iconSet attribute as "3TrafficLights".
  const legacy = computeConditionalFormats([block('A1:A7', { type: 'iconSet', iconSet: '3TrafficLights', cfvo: [] })], host(cells))
  assert.equal(legacy.get('A7')?.icon?.set, '3TrafficLights1')
  for (const set of ICON_SET_NAMES) {
    const result = computeConditionalFormats([block('A1:A7', createIconSetRule(set))], host(cells))
    assert.equal(result.get('A7')?.icon?.index, Number(set[0]) - 1, set)
  }
})

// ---------------------------------------------------------------------------------------------
// Visible window, clamping and performance
// ---------------------------------------------------------------------------------------------

await test('visible bounds limit per-cell work but statistics cover the whole range', () => {
  const cells: Cells = {}
  for (let row = 1; row <= 1000; row += 1) cells[`A${row}`] = row
  const counter = { evaluate: 0, value: 0 }
  const rules = [
    block('A1:A1000', { ...createTopBottomRule({ rank: 5 }, { style: RED_FILL }), priority: 1 }),
    block('A1:A1000', { ...createFormulaRule('A1>990', { style: BOLD }), priority: 2 }),
  ]
  const visible = computeConditionalFormats(rules, host(cells, { counter }), { top: 990, bottom: 999, left: 0, right: 5 })
  assert.deepEqual([...visible.keys()].sort(), ['A1000', 'A991', 'A992', 'A993', 'A994', 'A995', 'A996', 'A997', 'A998', 'A999'].sort())
  assert.equal(visible.get('A996')?.fill, '#FFC7CE', 'top 5 computed over the whole range')
  assert.equal(visible.get('A995')?.fill, undefined)
  assert.equal(counter.evaluate, 10, 'formula evaluated only for visible cells')

  // Whole-column rules are clamped to the used range.
  const clamped = { evaluate: 0, value: 0 }
  const result = computeConditionalFormats([block('A:A', createColorScaleRule([{ type: 'min', color: '#000000' }, { type: 'max', color: '#FFFFFF' }]))], host(cells, { counter: clamped, rowCount: 1000, colCount: 3 }))
  assert.equal(result.size, 1000)
  assert.ok(clamped.value <= 2000, `used-range clamp (${clamped.value} reads)`)
})

await test('performance: 50k covered cells under 50 ms', () => {
  const rows = 12_500
  const grid: ConditionalScalar[][] = []
  for (let row = 0; row < rows; row += 1) grid.push([row % 97, (row * 7) % 101, row % 13 === 0 ? 'x' : row, (row * 31) % 1000])
  const fastHost: ConditionalFormatHost = {
    sheetId: 's1',
    today: 45_308,
    rowCount: rows,
    colCount: 4,
    valueAt: (row, col) => grid[row]?.[col] ?? null,
    evaluate: () => false,
  }
  const rules = [
    block(`A1:A${rows}`, createColorScaleRule([{ type: 'min', color: '#F8696B' }, { type: 'percentile', value: 50, color: '#FFEB84' }, { type: 'max', color: '#63BE7B' }])),
    block(`B1:B${rows}`, createDataBarRule()),
    block(`C1:C${rows}`, createCellIsRule('greaterThan', [500], { style: RED_FILL })),
    block(`D1:D${rows}`, createIconSetRule('3Arrows')),
  ]
  const normalized = normalizeConditionalFormattings(rules)
  const timings: number[] = []
  let size = 0
  for (let run = 0; run < 7; run += 1) {
    const start = performance.now()
    size = computeConditionalFormats(normalized, fastHost).size
    timings.push(performance.now() - start)
  }
  timings.sort((a, b) => a - b)
  const median = timings[3]
  process.stdout.write(`  perf: ${size} formatted cells from 50,000 covered, median ${median.toFixed(1)} ms (best ${timings[0].toFixed(1)} ms)\n`)
  assert.ok(median < 50, `median ${median.toFixed(1)} ms`)
})

// ---------------------------------------------------------------------------------------------
// Authoring and management helpers
// ---------------------------------------------------------------------------------------------

await test('authoring helpers produce ExcelJS model shapes', () => {
  assert.deepEqual(createCellIsRule('between', [1, '=$B$1'], { style: BOLD }), { type: 'cellIs', style: { font: { bold: true } }, operator: 'between', formulae: ['1', '$B$1'] })
  assert.deepEqual(createTextRule('containsText', 'a"b', { ref: 'B2:B9' }), { type: 'containsText', operator: 'containsText', text: 'a"b', formulae: ['NOT(ISERROR(SEARCH("a""b",B2)))'] })
  assert.deepEqual(createTextRule('notContainsText', 'x'), { type: 'notContainsText', operator: 'notContains', text: 'x', formulae: ['ISERROR(SEARCH("x",A1))'] })
  assert.deepEqual(createTextRule('beginsWith', 'x').formulae, ['LEFT(A1,LEN("x"))="x"'])
  assert.deepEqual(createTextRule('endsWith', 'x').formulae, ['RIGHT(A1,LEN("x"))="x"'])
  assert.deepEqual(createBlankRule(true), { type: 'containsText', operator: 'containsBlanks', formulae: ['LEN(TRIM(A1))=0'] })
  assert.deepEqual(createErrorRule(false), { type: 'containsText', operator: 'notContainsErrors', formulae: ['NOT(ISERROR(A1))'] })
  assert.deepEqual(createDateRule('last7Days', { ref: 'C3' }), { type: 'timePeriod', timePeriod: 'last7Days', formulae: ['AND(TODAY()-FLOOR(C3,1)<=6,FLOOR(C3,1)<=TODAY())'] })
  assert.deepEqual(createTopBottomRule({ rank: 5, percent: true, bottom: true }), { type: 'top10', rank: 5, percent: true, bottom: true })
  assert.deepEqual(createAverageRule({ below: true, stdDev: 2 }), { type: 'aboveAverage', aboveAverage: false, stdDev: 2 })
  assert.deepEqual(createAverageRule({ equal: true }), { type: 'aboveAverage', equalAverage: true })
  assert.deepEqual(createDuplicateRule(true), { type: 'uniqueValues' })
  assert.deepEqual(createFormulaRule('=A1>1'), { type: 'expression', formulae: ['A1>1'] })
  assert.deepEqual(createColorScaleRule([{ type: 'min', color: '#F8696B' }, { type: 'num', value: '=$A$1', color: 'FFEB84' }]), {
    type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'num', value: '$A$1' }], color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }],
  })
  assert.deepEqual(createIconSetRule('4Arrows', { reverse: true, showValue: false }).cfvo, [
    { type: 'percent', value: 0 }, { type: 'percent', value: 25 }, { type: 'percent', value: 50 }, { type: 'percent', value: 75 },
  ])
  const bar = createDataBarRule({ color: '#63C384', gradient: false })
  assert.equal(bar.border, false)
  assert.equal(bar.gradient, false)
  assert.deepEqual(plain(bar.cfvo), [{ type: 'autoMin' }, { type: 'autoMax' }])
  assert.deepEqual(conditionalPresetStyle('lightRedFillDarkRedText'), {
    font: { color: { argb: 'FF9C0006' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' }, bgColor: { argb: 'FFFFC7CE' } },
  })
  for (const preset of CONDITIONAL_STYLE_PRESETS) assert.equal(matchConditionalStylePreset(preset.style), preset.id)
  assert.equal(matchConditionalStylePreset(createConditionalStyle({ fillColor: '#123456' })), 'custom')
})

await test('range parsing and validation', () => {
  assert.deepEqual(parseRangeInput('b2:d9, F1 $H$3:$H$4'), { ref: 'B2:D9 F1 H3:H4', ranges: [{ top: 1, bottom: 8, left: 1, right: 3 }, { top: 0, bottom: 0, left: 5, right: 5 }, { top: 2, bottom: 3, left: 7, right: 7 }] })
  assert.equal((parseRangeInput('A:A') as { ref: string }).ref, 'A1:A1048576')
  assert.ok('error' in parseRangeInput('B2:??'))
  assert.ok('error' in parseRangeInput(''))
})

await test('add / list / update / delete / move rules', () => {
  let model = addRule([], 'A1:A10', createCellIsRule('greaterThan', [5], { style: RED_FILL }))
  model = addRule(model, 'B1:B10', createDuplicateRule())
  model = addRule(model, 'C5', createTextRule('containsText', 'x'))
  const list = listRulesForRange(model)
  assert.deepEqual(list.map((entry) => [entry.id, entry.ref, entry.kind]), [[1, 'C5', 'containsText'], [2, 'B1:B10', 'duplicateValues'], [3, 'A1:A10', 'cellIs']])
  assert.deepEqual(list[0].rule.formulae, ['NOT(ISERROR(SEARCH("x",C5)))'], 'derived formula follows the real range')
  assert.deepEqual(listRulesForRange(model, 'A5:B6').map((entry) => entry.id), [2, 3])

  model = moveRulePriority(model, 3, -2)
  assert.deepEqual(listRulesForRange(model).map((entry) => entry.ref), ['A1:A10', 'C5', 'B1:B10'])
  model = moveRulePriority(model, 1, 5)
  assert.deepEqual(listRulesForRange(model).map((entry) => entry.ref), ['C5', 'B1:B10', 'A1:A10'])

  model = updateRule(model, 1, { ref: 'D1:D3' })
  assert.deepEqual(listRulesForRange(model)[0].rule.formulae, ['NOT(ISERROR(SEARCH("x",D1)))'])
  model = updateRule(model, 3, { rule: createCellIsRule('lessThan', [0], { style: BOLD }) })
  assert.equal(listRulesForRange(model)[2].rule.operator, 'lessThan')
  assert.equal(listRulesForRange(model)[2].rule.priority, 3)

  // Splitting a shared block when only one rule's range changes.
  const shared = [{ ref: 'E1:E5', rules: [{ ...createDuplicateRule(), priority: 1 }, { ...createBlankRule(true), priority: 2 }] }]
  const split = updateRule(shared, 2, { ref: 'F1:F5' })
  assert.equal(split.length, 2)
  assert.deepEqual(listRulesForRange(split).map((entry) => [entry.id, entry.ref]), [[1, 'E1:E5'], [2, 'F1:F5']])

  model = deleteRule(model, 2)
  assert.deepEqual(listRulesForRange(model).map((entry) => [entry.id, entry.ref]), [[1, 'D1:D3'], [2, 'A1:A10']])
  const frozen = Object.freeze([Object.freeze({ ref: 'A1', rules: Object.freeze([Object.freeze({ type: 'duplicateValues', priority: 1 })]) })])
  assert.doesNotThrow(() => deleteRule(frozen, 1), 'helpers never mutate their input')
  assert.equal(frozen[0].rules.length, 1)
})

await test('clearRulesInRange splits ranges and re-anchors relative formulas', () => {
  const model = [
    { ref: 'A1:C3', rules: [{ ...createFormulaRule('A1>B1', { style: BOLD }), priority: 1 }] },
    { ref: 'E1:E2', rules: [{ ...createDuplicateRule(), priority: 2 }] },
    { ref: 'A1:A3', rules: [{ ...createTextRule('beginsWith', 'x', { ref: 'A1:A3' }), priority: 3 }] },
  ]
  const cleared = clearRulesInRange(model, 'A1:B1 E1:E2')
  assert.deepEqual(cleared.map((item) => item.ref), ['C1 A2:C3', 'A2:A3'])
  const formula = cleared[0].rules[0].formulae?.[0]
  assert.equal(formula, 'C1>D1', 'formula re-anchored from A1 to C1')
  assert.deepEqual(cleared[1].rules[0].formulae, ['LEFT(A2,LEN("x"))="x"'])
  assert.deepEqual(cleared.flatMap((item) => item.rules.map((rule) => rule.priority)), [1, 2])
  // The re-anchored rule still formats the same cells.
  const cells: Cells = { A2: 5, B2: 1, C3: 9, D3: 1, C1: 3, D1: 1 }
  const before = computeConditionalFormats([model[0]], host(cells))
  const after = computeConditionalFormats([cleared[0]], host(cells))
  for (const address of ['C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']) assert.equal(after.has(address), before.has(address), address)
})

// ---------------------------------------------------------------------------------------------
// Editor draft model
// ---------------------------------------------------------------------------------------------

await test('editor drafts round-trip every rule type', () => {
  const rules: Array<[ConditionalRule, string]> = [
    [createCellIsRule('between', [1, '=$B$1'], { style: conditionalPresetStyle('greenFillDarkGreenText'), stopIfTrue: true }), 'Cell value between 1 and =$B$1'],
    [createCellIsRule('equal', ['"Done"'], { style: BOLD }), 'Cell value = Done'],
    [createTextRule('beginsWith', 'INV', { style: RED_FILL }), 'Text begins with “INV”'],
    [createDateRule('lastMonth', { style: RED_FILL }), 'Date is last month'],
    [createTopBottomRule({ rank: 15, percent: true }, { style: RED_FILL }), 'Top 15%'],
    [createAverageRule({ below: true, stdDev: 2 }, { style: RED_FILL }), '2 std dev below average'],
    [createAverageRule({ equal: true }, { style: RED_FILL }), 'Equal to or above average'],
    [createDuplicateRule(true, { style: RED_FILL }), 'Unique values'],
    [createBlankRule(false, { style: RED_FILL }), 'Cell is not empty'],
    [createErrorRule(true, { style: RED_FILL }), 'Cell contains an error'],
    [createFormulaRule('=AND($A1>0,B1<5)', { style: createConditionalStyle({ fontColor: '#0000FF', italic: true, borderColor: '#9C0006' }) }), 'Formula: =AND($A1>0,B1<5)'],
    [createColorScaleRule([{ type: 'num', value: 5, color: '#F8696B' }, { type: 'percent', value: 40, color: '#FFEB84' }, { type: 'formula', value: '$Z$1', color: '#63BE7B' }]), '3-color scale'],
    [createColorScaleRule([{ type: 'min', color: '#FFFFFF' }, { type: 'max', color: '#63BE7B' }]), '2-color scale'],
    [createDataBarRule({ color: '#FF555A', gradient: false, showValue: false, axisPosition: 'middle', direction: 'rightToLeft', min: { type: 'num', value: 0 }, max: { type: 'percentile', value: 90 } }), 'Data bar (bar only)'],
    [createIconSetRule('4TrafficLights', { reverse: true, thresholds: [{ type: 'num', value: 1 }, { type: 'percentile', value: 50, gte: false }, { type: 'formula', value: '$A$1' }] }), 'Icon set: 4 Traffic lights'],
  ]
  for (const [rule, description] of rules) {
    const ref = 'B2:B20'
    const prepared = addRule([], ref, rule)[0].rules[0]
    assert.equal(describeRule(prepared), description)
    const draft = ruleToDraft(prepared, ref)
    assert.equal(validateDraft(draft), null, description)
    const back = draftToRule(draft, prepared)
    assert.ok(!('error' in back), description)
    if ('error' in back) continue
    assert.equal(back.ref, ref)
    assert.deepEqual(back.rule, prepared, description)
  }
})

await test('editor validation and defaults', () => {
  const draft = defaultRuleDraft('A1:A5')
  assert.equal(draft.condition, 'cellIs:greaterThan')
  assert.equal(validateDraft(draft), 'Enter a value.')
  assert.equal(validateDraft({ ...draft, ref: 'nope!' }), '"nope!" is not a valid range.')
  assert.equal(validateDraft({ ...draft, condition: 'top:topPercent', value1: '101' }), 'Enter a whole number from 1 to 100.')
  const scale: RuleDraft = { ...draft, format: 'colorScale', scale: { ...draft.scale, points: [draft.scale.points[0], { ...draft.scale.points[1], value: '150' }, draft.scale.points[2]] } }
  assert.equal(validateDraft(scale), 'The midpoint percentile must be between 0 and 100.')
  assert.equal(withIconSet(draft, '5Rating').icons.thresholds.length, 4)
  const text = draftToRule({ ...draft, value1: 'hello' })
  assert.ok(!('error' in text) && text.rule.formulae?.[0] === '"hello"', 'plain text values become text literals')
  for (const option of CONDITION_OPTIONS) {
    const value = option.input === 'rank' || option.input === 'percentRank' ? '5' : option.input === 'formula' ? '=A1>1' : 'x'
    const result = draftToRule({ ...draft, condition: option.id, value1: value, value2: 'y' })
    assert.ok(!('error' in result), option.id)
    if (!('error' in result)) assert.notEqual(conditionalRuleKind(result.rule), null, option.id)
  }
})

// ---------------------------------------------------------------------------------------------
// ExcelJS round trips
// ---------------------------------------------------------------------------------------------

assert.equal(installConditionalFormattingPatches(), true, 'ExcelJS patches installed')

const PRESET = conditionalPresetStyle('lightRedFillDarkRedText')

function authoredModel() {
  let model: unknown[] = []
  const add = (ref: string, rule: ConditionalRule) => { model = addRule(model, ref, rule) }
  add('A1:A20', createCellIsRule('between', [1, '$Z$1'], { style: PRESET, stopIfTrue: true }))
  add('B1:B20', createFormulaRule('B1>A1', { style: { ...conditionalPresetStyle('yellowFillDarkYellowText'), numFmt: '0.0%' } }))
  add('C1:C20', createTextRule('containsText', 'abc', { style: PRESET }))
  add('C1:C20', createTextRule('notContainsText', 'zz', { style: PRESET }))
  add('D1:D20', createTextRule('beginsWith', 'Q', { style: conditionalPresetStyle('redBorder') }))
  add('D1:D20', createTextRule('endsWith', '!', { style: conditionalPresetStyle('redText') }))
  add('E1:E20', createBlankRule(true, { style: PRESET }))
  add('E1:E20', createBlankRule(false, { style: PRESET }))
  add('E1:E20', createErrorRule(true, { style: PRESET }))
  add('E1:E20', createErrorRule(false, { style: PRESET }))
  add('F1:F20', createDateRule('thisWeek', { style: PRESET }))
  add('G1:G20', createTopBottomRule({ rank: 3, bottom: true, percent: true }, { style: PRESET }))
  add('H1:H20', createAverageRule({ below: true, stdDev: 2 }, { style: PRESET }))
  add('H1:H20', createAverageRule({ equal: true }, { style: PRESET }))
  add('I1:I20 K1:K5', createDuplicateRule(false, { style: PRESET }))
  add('J1:J20', createDuplicateRule(true, { style: createConditionalStyle({ bold: true, italic: true, underline: true, strike: true, fontColor: '#123456' }) }))
  add('L1:L20', createColorScaleRule([{ type: 'min', color: '#F8696B' }, { type: 'percentile', value: 50, color: '#FFEB84' }, { type: 'max', color: '#63BE7B' }]))
  add('M1:M20', createColorScaleRule([{ type: 'formula', value: '$Z$2', color: '#FFFFFF' }, { type: 'num', value: 90, color: '#5A8AC6' }]))
  add('N1:N20', createDataBarRule({ color: '#638EC6', showValue: false }))
  add('O1:O20', createDataBarRule({ color: '#FFB628', gradient: false, axisPosition: 'middle', direction: 'rightToLeft', min: { type: 'num', value: -5 }, max: { type: 'percentile', value: 95 } }))
  add('P1:P20', createIconSetRule('3Arrows', { reverse: true, showValue: false, thresholds: [{ type: 'num', value: 10 }, { type: 'percent', value: 70, gte: false }] }))
  add('Q1:Q20', createIconSetRule('3Stars'))
  add('R1:R20', createIconSetRule('5Boxes', { thresholds: [{ type: 'percentile', value: 20 }, { type: 'formula', value: '$Z$3' }, { type: 'num', value: 60 }, { type: 'percent', value: 80 }] }))
  add('S1:S20', createIconSetRule('3Triangles'))
  add('T1:T20', { ...createIconSetRule('3Arrows'), icons: [{ iconSet: '3Flags', iconId: 0 }, { iconSet: 'NoIcons', iconId: 0 }, { iconSet: '3Stars', iconId: 2 }] })
  add('U1:U20', createIconSetRule('5Quarters'))
  return model as Array<{ ref: string; rules: ConditionalRule[] }>
}

/** Comparable form: drop writer bookkeeping (dxfId, x14Id, ref) and undefined keys; sort by priority. */
function comparable(model: unknown): unknown {
  const blocks = normalizeConditionalFormattings(model)
  const clean = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(clean)
    if (!value || typeof value !== 'object') return value
    const result: Record<string, unknown> = {}
    for (const [key, item] of Object.entries(value as Record<string, unknown>)) {
      if (item === undefined || key === 'dxfId' || key === 'x14Id' || key === 'ref' || key === 'numFmtId') continue
      result[key] = clean(item)
    }
    return result
  }
  return listRulesForRange(blocks).map((entry) => ({ ref: entry.ref, rule: clean(entry.rule) }))
}

async function writeAndRead(model: unknown[]) {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Rules')
  sheet.getCell('A1').value = 1
  ;(sheet as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings = JSON.parse(JSON.stringify(model))
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer())
  const back = new ExcelJS.Workbook()
  await back.xlsx.load(buffer)
  return { buffer, model: (back.worksheets[0] as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings }
}

await test('ExcelJS round trip preserves every authored rule', async () => {
  const model = authoredModel()
  const { buffer, model: read } = await writeAndRead(model)
  assert.deepEqual(comparable(read), comparable(model))
  const zip = await JSZip.loadAsync(buffer)
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  for (const fragment of [
    'type="duplicateValues"', 'type="uniqueValues"', 'type="notContainsText"', 'operator="notContains"', 'type="beginsWith"', 'type="endsWith"',
    'type="containsBlanks"', 'type="notContainsErrors"', 'stopIfTrue="1"', 'text="abc"', 'equalAverage="1"', 'stdDev="2"', 'gte="0"',
    'x14:dataBar', 'axisPosition="middle"', 'direction="rightToLeft"', 'iconSet="3Stars"', 'iconSet="5Boxes"',
    'custom="1"', 'iconId="2"', '<xm:f>$Z$3</xm:f>', '<cfvo type="formula" val="$Z$2"/>', 'sqref="I1:I20 K1:K5"',
  ]) assert.ok(xml.includes(fragment), `sheet XML contains ${fragment}`)
  assert.ok(!/operator="containsBlanks"|operator="notContainsErrors"/.test(xml), 'no invalid operator attributes')
  assert.match(xml, /<dataBar[^>]*showValue="0"/)
  assert.ok(!/<cfvo type="autoM/.test(xml), 'autoMin/autoMax only inside the x14 extension')
  const bordered = [...xml.matchAll(/<x14:dataBar[^>]*>.*?<\/x14:dataBar>/gs)].map((match) => match[0]).find((item) => item.includes('<x14:borderColor')) ?? ''
  const at = (tag: string) => bordered.indexOf(`<x14:${tag}`)
  assert.ok(at('cfvo') >= 0 && at('cfvo') < at('borderColor') && at('borderColor') < at('negativeFillColor') && at('negativeFillColor') < at('negativeBorderColor') && at('negativeBorderColor') < at('axisColor'), 'x14 child order')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  assert.ok(styles.includes('formatCode="0.0%"') && !styles.includes('[object Object]'))
  // Second generation is stable.
  const again = await writeAndRead(read as unknown[])
  assert.deepEqual(comparable(again.model), comparable(model))
})

const EXCEL_CF_XML = `<conditionalFormatting sqref="A1:A10"><cfRule type="containsText" dxfId="0" priority="1" stopIfTrue="1" operator="containsText" text="ok"><formula>NOT(ISERROR(SEARCH("ok",A1)))</formula></cfRule></conditionalFormatting>`
  + `<conditionalFormatting sqref="B1:B10"><cfRule type="aboveAverage" dxfId="1" priority="2" aboveAverage="0" equalAverage="1"/><cfRule type="duplicateValues" dxfId="0" priority="3"/></conditionalFormatting>`
  + `<conditionalFormatting sqref="C1:C10"><cfRule type="dataBar" priority="4"><dataBar showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{AAAAAAAA-0000-0000-0000-000000000001}</x14:id></ext></extLst></cfRule></conditionalFormatting>`
  + `<conditionalFormatting sqref="D1:D10"><cfRule type="iconSet" priority="5"><iconSet iconSet="4Arrows" reverse="1"><cfvo type="percent" val="0"/><cfvo type="formula" val="$Z$1"/><cfvo type="percentile" val="50" gte="0"/><cfvo type="num" val="75"/></iconSet></cfRule></conditionalFormatting>`
  + `<conditionalFormatting sqref="E1:E10"><cfRule type="colorScale" priority="6"><colorScale><cfvo type="min"/><cfvo type="max"/><color theme="4" tint="0.39997558519241921"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>`
  + `<conditionalFormatting sqref="F1:F10"><cfRule type="endsWith" dxfId="1" priority="7" operator="endsWith" text="z"><formula>RIGHT(F1,LEN("z"))="z"</formula></cfRule></conditionalFormatting>`

const EXCEL_EXT_XML = `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings>`
  + `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{AAAAAAAA-0000-0000-0000-000000000001}"><x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>C1:C10</xm:sqref></x14:conditionalFormatting>`
  + `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="iconSet" priority="8" id="{AAAAAAAA-0000-0000-0000-000000000002}"><x14:iconSet iconSet="3Stars" showValue="0"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="formula"><xm:f>Other!$A$1</xm:f></x14:cfvo><x14:cfvo type="percent" gte="0"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>G1:G10</xm:sqref></x14:conditionalFormatting>`
  + `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="expression" priority="9" id="{AAAAAAAA-0000-0000-0000-000000000003}"><xm:f>H1&gt;Other!$B$1</xm:f><x14:dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>H1:H10</xm:sqref></x14:conditionalFormatting>`
  + `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="iconSet" priority="10" id="{AAAAAAAA-0000-0000-0000-000000000004}"><x14:iconSet iconSet="3Arrows" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo><x14:cfIcon iconSet="3Flags" iconId="0"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>I1:I10</xm:sqref></x14:conditionalFormatting>`
  + `</x14:conditionalFormattings></ext></extLst>`

const EXCEL_DXFS = `<dxfs count="2"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf><dxf><font><b/><i/></font><numFmt numFmtId="164" formatCode="0.000"/><border><left style="thin"><color rgb="FF9C0006"/></left><right style="thin"><color rgb="FF9C0006"/></right><top style="thin"><color rgb="FF9C0006"/></top><bottom style="thin"><color rgb="FF9C0006"/></bottom></border></dxf></dxfs>`

async function excelAuthoredWorkbook(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Rules').getCell('A1').value = 1
  workbook.addWorksheet('Other').getCell('A1').value = 50
  const zip = await JSZip.loadAsync(Buffer.from(await workbook.xlsx.writeBuffer()))
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  sheet = sheet.replace(/<pageMargins/, `${EXCEL_CF_XML}<pageMargins`).replace(/<\/worksheet>/, `${EXCEL_EXT_XML}</worksheet>`)
  zip.file('xl/worksheets/sheet1.xml', sheet)
  let styles = await zip.file('xl/styles.xml')!.async('string')
  styles = styles.replace(/<dxfs[^>]*\/>|<dxfs[^>]*>[\s\S]*?<\/dxfs>/, EXCEL_DXFS)
  zip.file('xl/styles.xml', styles)
  return Buffer.from(await zip.generateAsync({ type: 'nodebuffer' }))
}

function excelAuthoredExpectation() {
  const red = { font: { color: { argb: 'FF9C0006' } }, fill: { type: 'pattern', bgColor: { argb: 'FFFFC7CE' } } }
  const bold = {
    font: { bold: true, italic: true },
    numFmt: '0.000',
    border: {
      left: { style: 'thin', color: { argb: 'FF9C0006' } }, right: { style: 'thin', color: { argb: 'FF9C0006' } },
      top: { style: 'thin', color: { argb: 'FF9C0006' } }, bottom: { style: 'thin', color: { argb: 'FF9C0006' } },
    },
  }
  return { red, bold }
}

await test('Excel-authored conditional formats import without loss', async () => {
  const buffer = await excelAuthoredWorkbook()
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(buffer)
  const rules = listRulesForRange((workbook.worksheets[0] as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings)
  const byPriority = new Map(rules.map((entry) => [entry.id, entry]))
  const { red, bold } = excelAuthoredExpectation()
  const text = byPriority.get(1)!.rule
  assert.equal(text.stopIfTrue, true)
  assert.equal(text.text, 'ok')
  assert.deepEqual(plain(text.style?.font), red.font)
  assert.equal((text.style?.fill as { bgColor?: unknown })?.bgColor !== undefined, true)
  const average = byPriority.get(2)!.rule
  assert.equal(average.aboveAverage, false)
  assert.equal(average.equalAverage, true)
  assert.equal(average.style?.numFmt, '0.000', 'dxf numFmt normalised to its format code')
  assert.deepEqual(plain(average.style?.border), bold.border)
  assert.equal(byPriority.get(3)!.rule.type, 'duplicateValues')
  const bar = byPriority.get(4)!.rule
  assert.deepEqual(plain(bar.cfvo), [{ type: 'autoMin' }, { type: 'autoMax' }], 'x14 thresholds win')
  assert.equal(bar.showValue, false)
  assert.equal(bar.minLength, 0)
  assert.equal(bar.border, true)
  assert.equal(bar.negativeBarBorderColorSameAsPositive, false)
  assert.equal(bar.negativeBarColorSameAsPositive, undefined, 'absent means the schema default (false)')
  assert.deepEqual(plain(bar.negativeFillColor), { argb: 'FFFF0000' })
  const icons = byPriority.get(5)!.rule
  assert.deepEqual(plain(icons.cfvo), [{ type: 'percent', value: 0 }, { type: 'formula', value: '$Z$1' }, { type: 'percentile', value: 50, gte: false }, { type: 'num', value: 75 }])
  assert.equal(icons.reverse, true)
  assert.deepEqual(plain((byPriority.get(6)!.rule.color as unknown[])[0]), { theme: 4, tint: 0.39997558519241921 })
  assert.equal(byPriority.get(7)!.rule.type, 'endsWith')
  assert.equal(byPriority.get(7)!.rule.text, 'z')
  const stars = byPriority.get(8)!.rule
  assert.equal(stars.iconSet, '3Stars')
  assert.equal(stars.showValue, false)
  assert.deepEqual(plain(stars.cfvo), [{ type: 'percent', value: 0 }, { type: 'formula', value: 'Other!$A$1' }, { type: 'percent', value: 67, gte: false }])
  const crossSheet = byPriority.get(9)!.rule
  assert.equal(crossSheet.type, 'expression')
  assert.deepEqual(plain(crossSheet.formulae), ['H1>Other!$B$1'])
  assert.deepEqual(plain(crossSheet.style?.font), { bold: true, color: { argb: 'FF9C0006' } })
  const custom = byPriority.get(10)!.rule
  assert.deepEqual(plain(custom.icons), [{ iconSet: '3Flags', iconId: 0 }, { iconSet: 'NoIcons', iconId: 0 }, { iconSet: '3Stars', iconId: 2 }])

  // Evaluate the imported model directly.
  const cells: Cells = { A1: 'OK go', A2: 'no', C1: -5, C2: 10, D1: 1, D2: 100, Z1: 30 }
  const result = computeConditionalFormats(rules.map((entry) => ({ ref: entry.ref, rules: [entry.rule] })), host(cells))
  assert.equal(result.get('A1')?.fill, '#FFC7CE')
  assert.equal(result.get('C1')?.dataBar?.negative, true)
  assert.equal(result.get('C1')?.dataBar?.color, '#FF0000')
  assert.equal(result.get('C2')?.dataBar?.showValue, false)
  assert.equal(result.get('D2')?.icon?.index, 0, '4Arrows reversed: top value gets the lowest icon')

  // Re-save the imported model and read it again: nothing lost.
  const { model: again } = await writeAndRead(rules.map((entry) => ({ ref: entry.ref, rules: [entry.rule] })))
  assert.deepEqual(comparable(again), comparable(rules.map((entry) => ({ ref: entry.ref, rules: [entry.rule] }))))
})

await test('workbook pipeline (workbooks.cjs import -> model -> source-backed save) keeps rules', async () => {
  const workbooks = require('../electron/workbooks.cjs') as {
    workbookPayloadFromBytes: (name: string, data: Buffer) => Promise<{ workbook: { sheets: Array<{ conditionalFormattings?: unknown[] }> } }>
    serializeWorkbook: (model: unknown, format?: string, options?: { baseBytes?: Buffer | null }) => Promise<Buffer>
  }
  const source = await excelAuthoredWorkbook()
  const payload = await workbooks.workbookPayloadFromBytes('rules.xlsx', source)
  const imported = payload.workbook.sheets[0].conditionalFormattings ?? []
  assert.equal(listRulesForRange(imported).length, 10)
  // Edit through the helpers (as the panel would), then save both source-backed and fresh.
  const edited = addRule(imported, 'J1:J5', createDuplicateRule(true, { style: PRESET }))
  payload.workbook.sheets[0].conditionalFormattings = edited
  for (const baseBytes of [source, null]) {
    const saved = await workbooks.serializeWorkbook(payload.workbook, 'xlsx', { baseBytes })
    const reread = await workbooks.workbookPayloadFromBytes('rules.xlsx', saved)
    assert.deepEqual(comparable(reread.workbook.sheets[0].conditionalFormattings), comparable(edited), baseBytes ? 'source-backed save' : 'fresh save')
  }
})

// ---------------------------------------------------------------------------------------------

if (failures.length) {
  process.stderr.write(`Conditional-format QA: ${failures.length} failed, ${passed} passed\n\n${failures.map((failure) => `  x ${failure}`).join('\n\n')}\n`)
  process.exit(1)
}
process.stdout.write(`Conditional-format QA passed: ${passed} checks (evaluation, priorities, statistics, authoring, editor drafts, ExcelJS and workbook round trips).\n`)
