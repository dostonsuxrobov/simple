import assert from 'node:assert/strict'
import { createElement } from 'react'
import { renderToString } from 'react-dom/server.browser'
import { FilterMenu } from '../src/components/FilterMenu'
import { SortDialog } from '../src/components/SortDialog'
import { DataValidationDialog } from '../src/components/DataValidationDialog'
import { RemoveDuplicatesDialog, TextToColumnsDialog } from '../src/components/DataCleanupDialogs'
import type { CellData, SheetData, SheetFilterState } from '../src/spreadsheet-types'
import {
  addressOf,
  applyCellChanges,
  createSheetHost,
  isDateNumberFormat,
  normalizeColor,
  parseDateText,
  parseNumberText,
  parseRange,
  parseTimeText,
  serialFromYMD,
  wildcardRegExp,
} from '../src/lib/data-tools-core'
import type { DataHost, Scalar } from '../src/lib/data-tools-core'
import {
  FILTER_OPERATORS,
  adjustFilterForStructure,
  adjustRowNumbersForStructure,
  applyFilterResult,
  computeFilteredRows,
  createFilterState,
  describeCriteria,
  detectCurrentRegion,
  distinctColumnValues,
  extendFilterRange,
  filterRangeForSelection,
  filterStateFromSheet,
  hiddenRowsWithoutFilter,
  refreshSheetFilter,
  setColumnCriteria,
  valuesCriteria,
} from '../src/lib/filter'
import {
  BUILTIN_CUSTOM_LISTS,
  computeSortOrder,
  detectHeaderRow,
  sortBlocker,
  sortKeyColors,
  sortKeyLabels,
  sortRange,
  sortSpecForFilter,
} from '../src/lib/sort'
import type { SortSpec } from '../src/lib/sort'
import {
  asValidation,
  describeValidation,
  dialogStateToValidation,
  findInvalidCells,
  findValidation,
  hasInCellDropdown,
  listOptionsForValidation,
  rangesWithSameValidation,
  validateValue,
  validationToDialogState,
  validationsEqual,
  withValidation,
} from '../src/lib/validation'
import type { DataValidationModel } from '../src/lib/validation'
import { listEntryForOption } from '../src/lib/validation'
import { currentRegionAround, filterHiddenRowSet } from '../src/lib/filter'
import { sortExpansionRegion } from '../src/lib/sort'
import { addMonthsToSerial, addWeekdaysToSerial, createAutofillPatch, fillHandleDoubleClickBottom } from '../src/lib/autofill'
import type { AutofillMode } from '../src/lib/autofill'
import {
  changeCase,
  convertField,
  fillBlanksFromAbove,
  guessDelimiters,
  previewSplit,
  properCase,
  removeDuplicates,
  splitLine,
  splitTextToColumns,
  trimText,
  trimWhitespace,
} from '../src/lib/data-cleanup'

let checks = 0
function test(name: string, run: () => void) {
  try {
    run()
    checks += 1
  } catch (error) {
    process.stderr.write(`\nFAILED: ${name}\n`)
    throw error
  }
}

type Input = Scalar | CellData | undefined

/** Build a sheet from rows of values (or full CellData) starting at A1 (or `origin`). */
function sheetFrom(rows: Input[][], origin = { row: 0, col: 0 }): SheetData {
  const cells: Record<string, CellData> = {}
  rows.forEach((row, r) => row.forEach((input, c) => {
    if (input === undefined) return
    const cell: CellData = input !== null && typeof input === 'object' ? input : { value: input }
    cells[addressOf(origin.row + r, origin.col + c)] = cell
  }))
  return { id: 's', name: 'Sheet1', rowCount: 1000, colCount: 50, cells, merges: [], colWidths: {}, rowHeights: {} }
}

/** Minimal A1 reference shifter for tests (relative refs only; $ anchors kept). */
function shift(formula: string, rowDelta: number, colDelta: number): string {
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (_match, colAbs: string, col: string, rowAbs: string, row: string) => {
    let colIndex = col.split('').reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0) - 1
    let rowIndex = Number(row) - 1
    if (!colAbs) colIndex += colDelta
    if (!rowAbs) rowIndex += rowDelta
    return `${colAbs}${addressOf(0, colIndex).replace(/\d+$/, '')}${rowAbs}${rowIndex + 1}`
  })
}

const NOW = new Date(2024, 4, 15, 10, 30) // Wednesday 15 May 2024
const TODAY = serialFromYMD(2024, 5, 15)

function hostFor(sheet: SheetData, overrides: Partial<DataHost> = {}): DataHost {
  return createSheetHost(sheet, { now: () => NOW, shiftFormula: shift, ...overrides })
}

function column(sheet: SheetData, col: number, from: number, to: number) {
  const host = hostFor(sheet)
  const values: Array<Scalar | undefined> = []
  for (let row = from; row <= to; row += 1) values.push(host.valueAt(row, col))
  return values
}

// ---------------------------------------------------------------------------------------------
// Core helpers
// ---------------------------------------------------------------------------------------------

test('core parsing', () => {
  assert.deepEqual(parseRange('$B$2:$D$9'), { top: 1, bottom: 8, left: 1, right: 3 })
  assert.deepEqual(parseRange("'My Sheet'!A1:B2"), { top: 0, bottom: 1, left: 0, right: 1 })
  assert.deepEqual(parseRange('C:D', { rows: 100 }), { top: 0, bottom: 99, left: 2, right: 3 })
  assert.equal(parseRange('nope'), null)
  assert.equal(parseNumberText('1,234.5'), 1234.5)
  assert.equal(parseNumberText('$1,200'), 1200)
  assert.equal(parseNumberText('(12)'), -12)
  assert.equal(parseNumberText('12%'), 0.12)
  assert.equal(parseNumberText('1e3'), 1000)
  assert.equal(parseNumberText('12abc'), null)
  assert.equal(parseNumberText('1,23'), null)
  assert.equal(parseDateText('2024-01-01'), 45292)
  assert.equal(parseDateText('1/1/2024'), 45292)
  assert.equal(parseDateText('1/1/24'), 45292)
  assert.equal(parseDateText('15-Jan-2024'), 45306)
  assert.equal(parseDateText('Jan 15, 2024'), 45306)
  assert.equal(parseDateText('January 15 2024'), 45306)
  assert.equal(parseDateText('2/30/2024'), null)
  assert.equal(parseDateText('2024-01-01T12:00:00.000Z'), 45292.5)
  assert.ok(Math.abs(parseDateText('1/1/2024 6:00 PM')! - 45292.75) < 1e-9)
  assert.equal(parseTimeText('12:00'), 0.5)
  assert.equal(parseTimeText('6 PM'), 0.75)
  assert.equal(parseTimeText('12:00 AM'), 0)
  assert.equal(parseTimeText('25:00'), null)
  assert.equal(parseTimeText('9'), null)
  assert.equal(isDateNumberFormat('m/d/yyyy'), true)
  assert.equal(isDateNumberFormat('[h]:mm:ss'), true)
  assert.equal(isDateNumberFormat('#,##0.00'), false)
  assert.equal(isDateNumberFormat('"Days: "0'), false)
  assert.equal(isDateNumberFormat('0.00%'), false)
  assert.equal(normalizeColor('FFFF0000'), '#FF0000')
  assert.equal(normalizeColor('#0f0'), '#00FF00')
  assert.equal(normalizeColor({ theme: 4 }), '#4472C4')
  assert.equal(normalizeColor({ indexed: 2 }), '#FF0000')
  assert.equal(wildcardRegExp('a*e').test('Apple'), true)
  assert.equal(wildcardRegExp('?pple').test('apple'), true)
  assert.equal(wildcardRegExp('~*').test('*'), true)
  assert.equal(wildcardRegExp('~*').test('x'), false)
  assert.equal(wildcardRegExp('a.c').test('abc'), false)
})

// ---------------------------------------------------------------------------------------------
// Filter
// ---------------------------------------------------------------------------------------------

const red = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } } }
const blueFont = { font: { color: { argb: 'FF0000FF' } } }

function filterFixture() {
  // Header row 1; data rows 2..11.
  return sheetFrom([
    ['Name', 'Qty', 'When', 'Flag'],
    ['Apple', 10, { value: TODAY, numFmt: 'm/d/yyyy' }, true],
    ['banana', 5, { value: TODAY - 1, numFmt: 'm/d/yyyy' }, false],
    [{ value: 'Cherry', style: red }, 20, { value: TODAY + 1, numFmt: 'm/d/yyyy' }, true],
    ['apple', 7, { value: serialFromYMD(2024, 5, 12), numFmt: 'm/d/yyyy' }, false],
    [null, 15, { value: serialFromYMD(2024, 5, 20), numFmt: 'm/d/yyyy' }, true],
    [{ value: 'Date*', style: blueFont }, '#N/A', { value: serialFromYMD(2024, 4, 30), numFmt: 'm/d/yyyy' }, false],
    ['grape', 'n/a', { value: serialFromYMD(2024, 2, 10), numFmt: 'm/d/yyyy' }, true],
    [{ value: 'kiwi', style: red }, 20, { value: serialFromYMD(2023, 12, 31), numFmt: 'm/d/yyyy' }, false],
    ['Avocado', -3, { value: serialFromYMD(2025, 1, 2), numFmt: 'm/d/yyyy' }, true],
    ['fig', 0, undefined, false],
  ])
}

function hiddenFor(criteria: SheetFilterState['columns'][number], columnIndex = 0, sheet = filterFixture(), overrides: Partial<DataHost> = {}) {
  const state: SheetFilterState = { ref: 'A1:D11', columns: { [columnIndex]: criteria } }
  return computeFilteredRows(state, hostFor(sheet, overrides))
}

/** Visible data rows (1-based) for a criteria. */
function shownFor(criteria: SheetFilterState['columns'][number], columnIndex = 0, sheet = filterFixture(), overrides: Partial<DataHost> = {}) {
  const hidden = new Set(hiddenFor(criteria, columnIndex, sheet, overrides))
  return [2, 3, 4, 5, 6, 7, 8, 9, 10, 11].filter((row) => !hidden.has(row))
}

test('filter: values checklist (case-insensitive) and blanks', () => {
  assert.deepEqual(shownFor({ values: ['APPLE'] }), [2, 5])
  assert.deepEqual(shownFor({ values: ['apple'], blanks: true }), [2, 5, 6])
  assert.deepEqual(shownFor({ values: [] , blanks: true }), [6])
  assert.deepEqual(shownFor({ blanks: false }), [2, 3, 4, 5, 7, 8, 9, 10, 11])
  assert.deepEqual(shownFor({ values: ['10', '#N/A'] }, 1), [2, 7])
  assert.deepEqual(shownFor({ values: ['TRUE'] }, 3), [2, 4, 6, 8, 10])
  // Dates match by display text.
  assert.deepEqual(shownFor({ values: ['5/15/2024'] }, 2, filterFixture(), { displayAt: (row, col) => {
    const value = hostFor(filterFixture()).valueAt(row, col)
    if (col === 2 && typeof value === 'number') { const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000); return `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${date.getUTCFullYear()}` }
    return hostFor(filterFixture()).displayAt(row, col)
  } }), [2])
})

test('filter: text operators with wildcards', () => {
  const op = (operator: string, value?: string | number, value2?: string | number) => shownFor({ condition: { operator, value, value2 } })
  assert.deepEqual(op('equals', 'apple'), [2, 5])
  assert.deepEqual(op('equals', 'a*'), [2, 5, 10])
  assert.deepEqual(op('equals', '?ig'), [11])
  assert.deepEqual(op('equals', 'Date~*'), [7])
  assert.deepEqual(op('notEquals', 'apple'), [3, 4, 6, 7, 8, 9, 10, 11])
  assert.deepEqual(op('beginsWith', 'a'), [2, 5, 10])
  assert.deepEqual(op('notBeginsWith', 'a'), [3, 4, 6, 7, 8, 9, 11])
  assert.deepEqual(op('endsWith', 'e'), [2, 5, 8])
  assert.deepEqual(op('notEndsWith', 'e'), [3, 4, 6, 7, 9, 10, 11])
  assert.deepEqual(op('contains', 'an'), [3])
  assert.deepEqual(op('contains', 'p*e'), [2, 5, 8])
  assert.deepEqual(op('notContains', 'p'), [3, 4, 6, 7, 9, 10, 11])
  assert.deepEqual(op('isEmpty'), [6])
  assert.deepEqual(op('isNotEmpty'), [2, 3, 4, 5, 7, 8, 9, 10, 11])
  assert.deepEqual(op('equals', ''), [6])
  assert.deepEqual(op('greaterThan', 'c'), [4, 7, 8, 9, 11])
})

test('filter: number operators', () => {
  const op = (operator: string, value?: string | number, value2?: string | number) => shownFor({ condition: { operator, value, value2 } }, 1)
  assert.deepEqual(op('equals', 20), [4, 9])
  assert.deepEqual(op('equals', '20'), [4, 9])
  assert.deepEqual(op('greaterThan', 10), [4, 6, 9])
  assert.deepEqual(op('greaterThanOrEqual', '10'), [2, 4, 6, 9])
  assert.deepEqual(op('lessThan', 5), [10, 11])
  assert.deepEqual(op('lessThanOrEqual', 5), [3, 10, 11])
  assert.deepEqual(op('between', 5, 15), [2, 3, 5, 6])
  assert.deepEqual(op('between', 15, 5), [2, 3, 5, 6])
  assert.deepEqual(op('notBetween', 5, 15), [4, 7, 8, 9, 10, 11])
  assert.deepEqual(op('top', 2), [4, 9]) // ties at 20 both shown
  assert.deepEqual(op('top', 3), [4, 6, 9])
  assert.deepEqual(op('bottom', 2), [10, 11])
  assert.deepEqual(op('topPercent', 25), [4, 9]) // 8 numbers -> 2 items
  assert.deepEqual(op('bottomPercent', 50), [3, 5, 10, 11])
  // mean of 10,5,20,7,15,20,-3,0 = 74/8 = 9.25
  assert.deepEqual(op('aboveAverage'), [2, 4, 6, 9])
  assert.deepEqual(op('belowAverage'), [3, 5, 10, 11])
})

test('filter: AND / OR conditions', () => {
  assert.deepEqual(shownFor({ condition: { operator: 'greaterThan', value: 5, operator2: 'lessThan', value2: 20, join: 'and' } }, 1), [2, 5, 6])
  assert.deepEqual(shownFor({ condition: { operator: 'lessThan', value: 0, operator2: 'greaterThanOrEqual', value2: 20, join: 'or' } }, 1), [4, 9, 10])
  assert.deepEqual(shownFor({ condition: { operator: 'beginsWith', value: 'a', operator2: 'endsWith', value2: 'o', join: 'and' } }), [10])
  assert.deepEqual(shownFor({ condition: { operator: 'equals', value: 'fig', operator2: 'equals', value2: 'kiwi', join: 'or' } }), [9, 11])
  // Non-combinable operators ignore a stray second condition.
  assert.deepEqual(shownFor({ condition: { operator: 'between', value: 5, value2: 7, operator2: 'equals', join: 'or' } }, 1), [3, 5])
})

test('filter: date operators and dynamic periods', () => {
  const op = (operator: string, value?: string | number, value2?: string | number) => shownFor({ condition: { operator, value, value2 } }, 2)
  assert.deepEqual(op('today'), [2])
  assert.deepEqual(op('yesterday'), [3])
  assert.deepEqual(op('tomorrow'), [4])
  assert.deepEqual(op('thisWeek'), [2, 3, 4, 5]) // Sun 12 .. Sat 18 May
  assert.deepEqual(op('nextWeek'), [6])
  assert.deepEqual(op('lastWeek'), [])
  assert.deepEqual(op('thisMonth'), [2, 3, 4, 5, 6])
  assert.deepEqual(op('lastMonth'), [7])
  assert.deepEqual(op('thisQuarter'), [2, 3, 4, 5, 6, 7])
  assert.deepEqual(op('lastQuarter'), [8])
  assert.deepEqual(op('thisYear'), [2, 3, 4, 5, 6, 7, 8])
  assert.deepEqual(op('lastYear'), [9])
  assert.deepEqual(op('nextYear'), [10])
  assert.deepEqual(op('yearToDate'), [2, 3, 5, 7, 8])
  assert.deepEqual(op('Q2'), [2, 3, 4, 5, 6, 7])
  assert.deepEqual(op('Q4'), [9])
  assert.deepEqual(op('M2'), [8])
  assert.deepEqual(op('M12'), [9])
  assert.deepEqual(op('dateEquals', '5/15/2024'), [2])
  assert.deepEqual(op('dateBefore', '2024-05-01'), [7, 8, 9])
  assert.deepEqual(op('dateAfter', 'May 15, 2024'), [4, 6, 10])
  assert.deepEqual(op('dateBetween', '5/12/2024', '5/15/2024'), [2, 3, 5])
})

test('filter: colours, custom formula and combined parts', () => {
  assert.deepEqual(shownFor({ fillColor: '#FF0000' }), [4, 9])
  assert.deepEqual(shownFor({ fillColor: 'none' }), [2, 3, 5, 6, 7, 8, 10, 11])
  assert.deepEqual(shownFor({ fontColor: '#0000ff' }), [7])
  assert.deepEqual(shownFor({ fillColor: '#FF0000', values: ['kiwi'] }), [9])
  const evaluate = (formula: string, row: number) => {
    // Custom formula "=B2>10" shifted per row.
    const match = /^B(\d+)>(\d+)$/.exec(formula)
    assert.ok(match, formula)
    assert.equal(Number(match[1]) - 1, row)
    const value = hostFor(filterFixture()).valueAt(row, 1)
    return typeof value === 'number' && value > Number(match[2])
  }
  assert.deepEqual(shownFor({ condition: { operator: 'customFormula', value: '=B2>10' } }, 0, filterFixture(), { evaluate }), [4, 6, 9])
})

test('filter: multi-column and checklist cross-filtering', () => {
  const sheet = filterFixture()
  const host = hostFor(sheet)
  let state: SheetFilterState = createFilterState({ top: 0, bottom: 10, left: 0, right: 3 })
  assert.equal(state.ref, 'A1:D11')
  state = setColumnCriteria(state, 3, { values: ['TRUE'] })
  state = setColumnCriteria(state, 1, { condition: { operator: 'greaterThan', value: 10 } })
  assert.deepEqual(computeFilteredRows(state, host), [2, 3, 5, 7, 8, 9, 10, 11])
  // Checklist for column A lists only values from rows visible under B and D.
  const names = distinctColumnValues(state, 0, host)
  assert.deepEqual(names.items.map((item) => item.text), ['Cherry'])
  assert.equal(names.blanks?.count, 1)
  // Checklist for column B ignores B's own filter but applies D's.
  const qty = distinctColumnValues(state, 1, host)
  assert.deepEqual(qty.items.map((item) => [item.text, item.checked]), [['-3', false], ['10', false], ['15', true], ['20', true], ['n/a', false]])
  assert.equal(setColumnCriteria(state, 1, null).columns[1], undefined)
})

test('filter: checklist ordering, counts, kinds and colours', () => {
  const sheet = sheetFrom([
    ['Mixed'],
    ['pear'], [10], [2], ['Apple'], [true], [false], ['#DIV/0!'], [null], ['apple'], [2], [{ value: 45292, numFmt: 'm/d/yyyy', display: '1/1/2024' }], [''],
  ])
  const host = hostFor(sheet)
  const result = distinctColumnValues(createFilterState({ top: 0, bottom: 12, left: 0, right: 0 }), 0, host)
  assert.deepEqual(result.items.map((item) => item.text), ['2', '10', '1/1/2024', 'Apple', 'pear', 'FALSE', 'TRUE', '#DIV/0!'])
  assert.deepEqual(result.items.map((item) => item.count), [2, 1, 1, 2, 1, 1, 1, 1])
  assert.deepEqual(result.items.map((item) => item.kind), ['number', 'number', 'date', 'text', 'text', 'boolean', 'boolean', 'error'])
  assert.deepEqual(result.items[2].date, { year: 2024, month: 1, day: 1 })
  assert.equal(result.blanks?.count, 2)
  assert.equal(result.rows, 12)
  assert.ok(result.items.every((item) => item.checked))
  const colours = distinctColumnValues({ ref: 'A1:D11', columns: {} }, 0, hostFor(filterFixture()))
  assert.deepEqual(colours.fillColors, [{ color: null, count: 8 }, { color: '#FF0000', count: 2 }])
  assert.deepEqual(colours.fontColors.map((entry) => entry.color), [null, '#0000FF'])
  // Checklist -> criteria.
  const selected = result.items.filter((item) => item.kind !== 'text').map((item) => item.key)
  assert.deepEqual(valuesCriteria(result, selected, false), { values: ['2', '10', '1/1/2024', 'FALSE', 'TRUE', '#DIV/0!'], blanks: false })
  assert.equal(valuesCriteria(result, result.items.map((item) => item.key), true), null)
})

test('filter: range detection, extension and sheet state', () => {
  const sheet = sheetFrom([
    ['H1', 'H2', undefined, 'X'],
    [1, 2],
    [3, undefined],
    [undefined, 4, 5],
    [],
    [9],
  ])
  assert.deepEqual(detectCurrentRegion(sheet, 1, 0), { top: 0, bottom: 3, left: 0, right: 3 })
  assert.deepEqual(detectCurrentRegion(sheet, 5, 0), { top: 5, bottom: 5, left: 0, right: 0 })
  assert.deepEqual(detectCurrentRegion(sheet, 4, 0), { top: 0, bottom: 5, left: 0, right: 3 }) // blank cell touching both blocks
  assert.equal(detectCurrentRegion(sheet, 10, 10), null)
  assert.deepEqual(filterRangeForSelection(sheet, { top: 0, bottom: 0, left: 0, right: 1 }), { top: 0, bottom: 3, left: 0, right: 1 })
  assert.deepEqual(filterRangeForSelection(sheet, { top: 0, bottom: 2, left: 0, right: 1 }), { top: 0, bottom: 2, left: 0, right: 1 })
  const extended = extendFilterRange({ ref: 'A1:B2', columns: {} }, sheet)
  assert.equal(extended.ref, 'A1:B4')
  assert.deepEqual(filterStateFromSheet({ autoFilter: 'A1:C9' }), { ref: 'A1:C9', columns: {} })
  assert.deepEqual(filterStateFromSheet({ autoFilter: { from: { row: 2, column: 1 }, to: { row: 5, column: 3 } } }), { ref: 'A2:C5', columns: {} })
})

test('filter: hidden rows keep manual hides separate', () => {
  const first = applyFilterResult({ hiddenRows: [3, 20], filteredRows: [] }, [3, 5, 6])
  assert.deepEqual(first, { hiddenRows: [3, 5, 6, 20], filteredRows: [5, 6] })
  const second = applyFilterResult({ ...first }, [7])
  assert.deepEqual(second, { hiddenRows: [3, 7, 20], filteredRows: [7] })
  assert.deepEqual(hiddenRowsWithoutFilter(second), [3, 20])
  const sheet = filterFixture()
  const refreshed = refreshSheetFilter({ ...sheet, hiddenRows: [30] }, { ref: 'A1:D11', columns: { 1: { condition: { operator: 'greaterThan', value: 12 } } } }, hostFor(sheet))
  assert.deepEqual(refreshed.filteredRows, [2, 3, 5, 7, 8, 10, 11])
  assert.deepEqual(refreshed.hiddenRows, [2, 3, 5, 7, 8, 10, 11, 30])
  assert.equal(refreshed.autoFilter, 'A1:D11')
})

test('filter: structural edits', () => {
  const state: SheetFilterState = { ref: 'B2:E10', columns: { 0: { values: ['a'] }, 2: { blanks: false }, 3: { fillColor: '#FF0000' } }, sort: { column: 3, descending: true } }
  assert.equal(adjustFilterForStructure(state, { axis: 'row', kind: 'insert', index: 0, count: 2 })!.ref, 'B4:E12')
  assert.equal(adjustFilterForStructure(state, { axis: 'row', kind: 'insert', index: 5, count: 3 })!.ref, 'B2:E13')
  assert.equal(adjustFilterForStructure(state, { axis: 'row', kind: 'insert', index: 11, count: 3 })!.ref, 'B2:E10')
  assert.equal(adjustFilterForStructure(state, { axis: 'row', kind: 'delete', index: 1, count: 1 }), null)
  assert.equal(adjustFilterForStructure(state, { axis: 'row', kind: 'delete', index: 4, count: 2 })!.ref, 'B2:E8')
  assert.equal(adjustFilterForStructure(state, { axis: 'row', kind: 'delete', index: 8, count: 10 })!.ref, 'B2:E8')
  assert.equal(adjustFilterForStructure(state, { axis: 'row', kind: 'delete', index: 0, count: 1 })!.ref, 'B1:E9')
  const inserted = adjustFilterForStructure(state, { axis: 'column', kind: 'insert', index: 3, count: 1 })!
  assert.equal(inserted.ref, 'B2:F10')
  assert.deepEqual(Object.keys(inserted.columns), ['0', '3', '4'])
  assert.equal(inserted.sort?.column, 4)
  const deleted = adjustFilterForStructure(state, { axis: 'column', kind: 'delete', index: 3, count: 1 })!
  assert.equal(deleted.ref, 'B2:D10')
  assert.deepEqual(Object.keys(deleted.columns), ['0', '2'])
  assert.equal(deleted.columns[2].fillColor, '#FF0000')
  assert.equal(adjustFilterForStructure(state, { axis: 'column', kind: 'delete', index: 0, count: 1 })!.ref, 'A2:D10')
  assert.equal(adjustFilterForStructure(state, { axis: 'column', kind: 'delete', index: 0, count: 10 }), null)
  assert.deepEqual(adjustRowNumbersForStructure([3, 5, 9], { axis: 'row', kind: 'delete', index: 4, count: 1 }), [3, 8])
  assert.deepEqual(adjustRowNumbersForStructure([3, 5, 9], { axis: 'row', kind: 'insert', index: 3, count: 2 }), [3, 7, 11])
})

test('filter: descriptions and operator catalogue', () => {
  assert.equal(describeCriteria(undefined), 'Showing all')
  assert.equal(describeCriteria({ values: ['a', 'b'], blanks: true }), 'Showing "a", "b", (Blanks)')
  assert.equal(describeCriteria({ condition: { operator: 'greaterThan', value: 5, operator2: 'lessThan', value2: 9, join: 'and' } }), 'Greater than 5 and less than 9')
  assert.equal(describeCriteria({ condition: { operator: 'between', value: 1, value2: 3 } }), 'Is between 1 and 3')
  assert.ok(FILTER_OPERATORS.length > 50)
  assert.equal(new Set(FILTER_OPERATORS.map((item) => item.id)).size, FILTER_OPERATORS.length)
})

// ---------------------------------------------------------------------------------------------
// Sort
// ---------------------------------------------------------------------------------------------

function sortColumn(values: Input[], spec: Partial<SortSpec> = {}, extra: Input[] = []) {
  const sheet = sheetFrom(values.map((value, index) => [value, extra[index] ?? index]))
  const bounds = { top: 0, bottom: values.length - 1, left: 0, right: 1 }
  const result = sortRange(sheet, { bounds, levels: [{ key: 0 }], ...spec }, hostFor(sheet))
  assert.ok(result.ok, !result.ok ? result.error : '')
  applyCellChanges(sheet.cells, result.changes)
  return { sheet, values: column(sheet, 0, 0, values.length - 1), tags: column(sheet, 1, 0, values.length - 1) }
}

test('sort: Excel type order, blanks last in both directions', () => {
  const input: Input[] = ['b', 3, null, true, '#N/A', 'A', -1, false, '', 'a', 10]
  const asc = sortColumn(input)
  assert.deepEqual(asc.values.slice(0, 9), [-1, 3, 10, 'A', 'a', 'b', false, true, '#N/A'])
  assert.deepEqual(asc.tags.slice(0, 9), [6, 1, 10, 5, 9, 0, 7, 3, 4]) // 'A' before 'a' by stability
  assert.deepEqual(asc.tags.slice(9).sort(), [2, 8])
  const desc = sortColumn(input, { levels: [{ key: 0, descending: true }] })
  assert.deepEqual(desc.values.slice(0, 9), ['#N/A', true, false, 'b', 'A', 'a', 10, 3, -1])
  assert.deepEqual(desc.tags.slice(9).sort(), [2, 8])
})

test('sort: stability, case sensitivity and text collation', () => {
  const stable = sortColumn(['x', 'y', 'x', 'y', 'x'])
  assert.deepEqual(stable.tags, [0, 2, 4, 1, 3])
  const insensitive = sortColumn(['b', 'B', 'a', 'A'])
  assert.deepEqual(insensitive.values, ['a', 'A', 'b', 'B'])
  const sensitive = sortColumn(['B', 'b', 'A', 'a'], { caseSensitive: true })
  assert.deepEqual(sensitive.values, ['a', 'A', 'b', 'B'])
  const text = sortColumn(['item10', 'item2', 'Item1', 'éclair', 'eclair', 'zebra', '10', '9'])
  assert.deepEqual(text.values, ['10', '9', 'eclair', 'éclair', 'Item1', 'item10', 'item2', 'zebra'])
})

test('sort: multi-level, custom lists, colours', () => {
  const sheet = sheetFrom([
    ['Dept', 'Month', 'Amount'],
    ['Sales', 'Mar', 5],
    ['Ops', 'Jan', 7],
    ['Sales', 'Jan', 9],
    ['Ops', 'Feb', 1],
    ['Sales', 'Jan', 3],
    ['Ops', 'unknown', 2],
  ])
  const bounds = { top: 0, bottom: 6, left: 0, right: 2 }
  const host = hostFor(sheet)
  const months = BUILTIN_CUSTOM_LISTS.find((list) => list.id === 'months-short')!.items
  const result = sortRange(sheet, { bounds, hasHeader: true, levels: [{ key: 0 }, { key: 1, customList: months }, { key: 2, descending: true }] }, host)
  assert.ok(result.ok)
  applyCellChanges(sheet.cells, result.changes)
  assert.deepEqual(column(sheet, 0, 0, 6), ['Dept', 'Ops', 'Ops', 'Ops', 'Sales', 'Sales', 'Sales'])
  assert.deepEqual(column(sheet, 1, 0, 6), ['Month', 'Jan', 'Feb', 'unknown', 'Jan', 'Jan', 'Mar'])
  assert.deepEqual(column(sheet, 2, 0, 6), ['Amount', 7, 1, 2, 9, 3, 5])
  const reversed = computeSortOrder({ bounds, hasHeader: true, levels: [{ key: 1, customList: months, descending: true }] }, hostFor(sheet))
  assert.ok(reversed.ok)
  assert.deepEqual(reversed.order.map((row) => hostFor(sheet).valueAt(row, 1)), ['Mar', 'Feb', 'Jan', 'Jan', 'Jan', 'unknown'])
  const coloured = sheetFrom([[1], [{ value: 2, style: red }], [3], [{ value: 4, style: red }]])
  const byColor = computeSortOrder({ bounds: { top: 0, bottom: 3, left: 0, right: 0 }, levels: [{ key: 0, sortOn: 'fillColor', color: '#ff0000' }] }, hostFor(coloured))
  assert.ok(byColor.ok)
  assert.deepEqual(byColor.order, [1, 3, 0, 2])
  const bottom = computeSortOrder({ bounds: { top: 0, bottom: 3, left: 0, right: 0 }, levels: [{ key: 0, sortOn: 'fillColor', color: '#FF0000', position: 'bottom' }] }, hostFor(coloured))
  assert.ok(bottom.ok)
  assert.deepEqual(bottom.order, [0, 2, 1, 3])
  assert.deepEqual(sortKeyColors({ top: 0, bottom: 3, left: 0, right: 0 }, 'rows', 0, 'fillColor', false, hostFor(coloured)), ['#FF0000', null])
})

test('sort: left to right, fixed (hidden) lines, header labels', () => {
  const sheet = sheetFrom([
    ['Label', 3, 1, 2],
    ['Tag', 'c', 'a', 'b'],
  ])
  const bounds = { top: 0, bottom: 1, left: 0, right: 3 }
  const result = sortRange(sheet, { bounds, orientation: 'columns', hasHeader: true, levels: [{ key: 0 }] }, hostFor(sheet))
  assert.ok(result.ok)
  applyCellChanges(sheet.cells, result.changes)
  assert.deepEqual([0, 1, 2, 3].map((col) => hostFor(sheet).valueAt(1, col)), ['Tag', 'a', 'b', 'c'])
  assert.deepEqual(sortKeyLabels(bounds, 'columns', true, hostFor(sheet)), ['Label', 'Tag'])
  assert.deepEqual(sortKeyLabels({ top: 0, bottom: 3, left: 1, right: 2 }, 'rows', false, hostFor(sheet)), ['Column B', 'Column C'])
  const fixed = sortColumn([5, 4, 3, 2, 1], { fixed: [1, 3] })
  assert.deepEqual(fixed.values, [1, 4, 3, 2, 5])
  const fn = sortColumn([5, 4, 3, 2, 1], { fixed: (row) => row === 0 })
  assert.deepEqual(fn.values, [5, 1, 2, 3, 4])
  const unchanged = sortRange(sheetFrom([[1], [2]]), { bounds: { top: 0, bottom: 1, left: 0, right: 0 }, levels: [{ key: 0 }] }, hostFor(sheetFrom([[1], [2]])))
  assert.ok(unchanged.ok && !unchanged.changed && Object.keys(unchanged.changes).length === 0)
})

test('sort: moves styles, shifts formulas, unshares shared formulas', () => {
  const sheet = sheetFrom([
    [3, { formula: 'A1*2', result: 6 }, { value: 'x', style: red }],
    [1, { formula: 'A2*2', result: 2 }, undefined],
    [2, { formula: 'A3*2+$D$1', result: 4, formulaType: 'shared', formulaRange: 'B3:B4' }, 'z'],
  ])
  sheet.cells.B4 = { formula: 'A4*2+$D$1', formulaType: 'shared', sharedFormulaMaster: 'B3' } as CellData
  const bounds = { top: 0, bottom: 2, left: 0, right: 2 }
  const result = sortRange(sheet, { bounds, levels: [{ key: 0 }] }, hostFor(sheet), shift)
  assert.ok(result.ok)
  applyCellChanges(sheet.cells, result.changes)
  assert.equal(sheet.cells.B1.formula, 'A1*2')
  assert.equal(sheet.cells.B2.formula, 'A2*2+$D$1')
  assert.equal(sheet.cells.B3.formula, 'A3*2')
  assert.equal(sheet.cells.B1.result, undefined)
  assert.equal(sheet.cells.B2.formulaType, undefined)
  assert.equal(sheet.cells.B2.formulaRange, undefined)
  assert.equal(sheet.cells.B4.formulaType, undefined)
  assert.equal((sheet.cells.B4 as CellData & { sharedFormulaMaster?: string }).sharedFormulaMaster, undefined)
  assert.equal(sheet.cells.B4.formula, 'A4*2+$D$1')
  assert.equal(sheet.cells.C1, undefined)
  assert.deepEqual(sheet.cells.C3.style, red)
  assert.equal(sheet.cells.C2.value, 'z')
})

test('sort: refuses unequal merges and array formulas', () => {
  const sheet = sheetFrom([[3, 'a'], [1, 'b'], [2, 'c']])
  const spec: SortSpec = { bounds: { top: 0, bottom: 2, left: 0, right: 1 }, levels: [{ key: 0 }] }
  assert.equal(sortBlocker(sheet, spec), null)
  assert.match(sortBlocker({ ...sheet, merges: ['A1:A2'] }, spec)!, /same size/)
  assert.match(sortBlocker({ ...sheet, merges: ['A1:B1'] }, spec)!, /same size/)
  assert.equal(sortBlocker({ ...sheet, merges: ['A1:B1', 'A2:B2', 'A3:B3'] }, spec), null)
  assert.match(sortBlocker({ ...sheet, merges: ['B3:C3'] }, spec)!, /same size/)
  assert.equal(sortBlocker({ ...sheet, merges: ['A1:B1'] }, { ...spec, fixed: [1, 2] }), null)
  const array = sheetFrom([[3, { formula: 'SEQUENCE(2)', formulaType: 'array', formulaRange: 'B1:B2' }], [1], [2]])
  assert.match(sortBlocker(array, spec)!, /array/)
  const spill = sheetFrom([[3], [1], [2]])
  spill.cells.E1 = { formula: 'SEQUENCE(5)', dynamicFormula: true, formulaRange: 'E1:E5' }
  assert.equal(sortBlocker(spill, spec), null)
  assert.match(sortBlocker(spill, { ...spec, bounds: { top: 0, bottom: 2, left: 0, right: 4 } })!, /array/)
  const failed = sortRange({ ...sheet, merges: ['A1:A2'] }, spec, hostFor(sheet))
  assert.ok(!failed.ok)
  assert.ok(!computeSortOrder({ ...spec, bounds: { top: 0, bottom: 0, left: 0, right: 1 }, hasHeader: true }, hostFor(sheet)).ok)
})

test('sort: from the filter menu', () => {
  const sheet = sheetFrom([['H'], [3], [1], [2], [0]])
  const spec = sortSpecForFilter('A1:A5', { key: 0 }, { fixedRows: [3] })!
  assert.deepEqual(spec.bounds, { top: 0, bottom: 4, left: 0, right: 0 })
  const result = sortRange(sheet, spec, hostFor(sheet))
  assert.ok(result.ok)
  applyCellChanges(sheet.cells, result.changes)
  assert.deepEqual(column(sheet, 0, 0, 4), ['H', 0, 1, 2, 3])
  assert.equal(sortSpecForFilter('A1:A1', { key: 0 }), null)
})

test('sort: header detection', () => {
  const host = (rows: Input[][]) => hostFor(sheetFrom(rows))
  assert.equal(detectHeaderRow({ top: 0, bottom: 3, left: 0, right: 1 }, host([['Name', 'Qty'], ['a', 1], ['b', 2], ['c', 3]])), true)
  assert.equal(detectHeaderRow({ top: 0, bottom: 2, left: 0, right: 1 }, host([['a', 1], ['b', 2], ['c', 3]])), false)
  assert.equal(detectHeaderRow({ top: 0, bottom: 2, left: 0, right: 0 }, host([['x'], ['y'], ['z']])), false)
  const bold = { font: { bold: true } }
  assert.equal(detectHeaderRow({ top: 0, bottom: 2, left: 0, right: 0 }, host([[{ value: 'Name', style: bold }], ['y'], ['z']])), true)
  assert.equal(detectHeaderRow({ top: 0, bottom: 0, left: 0, right: 1 }, host([['a', 'b']])), false)
})

// ---------------------------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------------------------

const vhost = hostFor(sheetFrom([
  ['Red', 1, 'x'],
  ['Green', 2, 'y'],
  [null, 3, 'Red'],
  ['Blue', 4, 'z'],
]), {
  resolveReference: (reference) => reference.toLowerCase() === 'colors' || reference === 'Lists!$A$1:$A$2'
    ? [{ value: 'Cyan', text: 'Cyan' }, { value: 'Magenta', text: 'Magenta' }, { value: 'cyan', text: 'cyan' }]
    : null,
  evaluate: (formula, _row, _col, override) => {
    if (formula === '$B$4') return 4
    if (formula === 'MAXLEN') return 3
    const match = /^ISNUMBER\(([A-Z]+\d+)\)$/.exec(formula)
    if (match && override) return typeof override.value === 'number'
    const even = /^MOD\(([A-Z]+)(\d+),2\)=0$/.exec(formula)
    if (even && override) return typeof override.value === 'number' && override.value % 2 === 0
    return '#NAME?'
  },
})

const rule = (model: Partial<DataValidationModel>): DataValidationModel => ({ type: 'any', showErrorMessage: true, ...model } as DataValidationModel)
const ok = (model: Partial<DataValidationModel>, raw: string | undefined, value: Scalar | undefined, target = { row: 0, col: 0 }) => validateValue(rule(model), raw, value, vhost, target).ok

test('validation: whole / decimal / text length with every operator', () => {
  const whole = (operator: DataValidationModel['operator'], formulae: DataValidationModel['formulae']) => rule({ type: 'whole', operator, formulae })
  const cases: Array<[DataValidationModel['operator'], DataValidationModel['formulae'], number, boolean]> = [
    ['between', [1, 10], 5, true], ['between', [1, 10], 11, false], ['between', [1, 10], 1, true],
    ['notBetween', [1, 10], 5, false], ['notBetween', [1, 10], 0, true],
    ['equal', [7], 7, true], ['equal', [7], 8, false],
    ['notEqual', [7], 7, false], ['notEqual', [7], 8, true],
    ['greaterThan', [7], 7, false], ['greaterThan', [7], 8, true],
    ['lessThan', [7], 6, true], ['lessThan', [7], 7, false],
    ['greaterThanOrEqual', [7], 7, true], ['greaterThanOrEqual', [7], 6, false],
    ['lessThanOrEqual', [7], 7, true], ['lessThanOrEqual', [7], 8, false],
  ]
  for (const [operator, formulae, value, expected] of cases) {
    assert.equal(validateValue(whole(operator, formulae), String(value), value, vhost).ok, expected, `${operator} ${value}`)
  }
  assert.equal(ok({ type: 'whole', operator: 'between', formulae: [1, 10] }, '2.5', 2.5), false)
  assert.equal(ok({ type: 'whole', operator: 'between', formulae: [1, 10] }, 'abc', 'abc'), false)
  const coerced = validateValue(rule({ type: 'whole', operator: 'between', formulae: [1, 2000] }), '1,000', '1,000', vhost)
  assert.deepEqual(coerced, { ok: true, value: 1000 })
  assert.equal(ok({ type: 'decimal', operator: 'between', formulae: [0.5, 1.5] }, '1.25', 1.25), true)
  assert.equal(ok({ type: 'decimal', operator: 'greaterThan', formulae: ['0.5'] }, '0.25', 0.25), false)
  assert.equal(ok({ type: 'decimal', operator: 'lessThanOrEqual', formulae: ['$B$4'] }, '4', 4), true)
  assert.equal(ok({ type: 'decimal', operator: 'lessThanOrEqual', formulae: ['$B$4'] }, '4.1', 4.1), false)
  assert.equal(ok({ type: 'textLength', operator: 'lessThanOrEqual', formulae: [3] }, 'abc', 'abc'), true)
  assert.equal(ok({ type: 'textLength', operator: 'lessThanOrEqual', formulae: [3] }, 'abcd', 'abcd'), false)
  assert.equal(ok({ type: 'textLength', operator: 'equal', formulae: ['MAXLEN'] }, '12.5', 12.5), false)
  assert.equal(ok({ type: 'textLength', operator: 'equal', formulae: ['MAXLEN'] }, '125', 125), true)
})

test('validation: dates and times', () => {
  const jan1 = { type: 'date', value: new Date(Date.UTC(2024, 0, 1)).toISOString() }
  const dec31 = new Date(Date.UTC(2024, 11, 31))
  const dates = { type: 'date' as const, operator: 'between' as const, formulae: [jan1, dec31] as DataValidationModel['formulae'] }
  assert.equal(ok(dates, '6/1/2024', serialFromYMD(2024, 6, 1)), true)
  assert.equal(ok(dates, '1/1/2025', serialFromYMD(2025, 1, 1)), false)
  assert.equal(ok(dates, 'hello', 'hello'), false)
  const typed = validateValue(rule(dates), 'Mar 3, 2024', 'Mar 3, 2024', vhost)
  assert.deepEqual(typed, { ok: true, value: serialFromYMD(2024, 3, 3) })
  assert.equal(ok({ type: 'date', operator: 'greaterThan', formulae: ['2024-01-01T00:00:00.000Z'] }, '', serialFromYMD(2024, 1, 2)), true)
  assert.equal(ok({ type: 'date', operator: 'lessThan', formulae: [45292] }, '', 45292), false)
  const times = { type: 'time' as const, operator: 'between' as const, formulae: ['0.375', '0.708333333333333'] as DataValidationModel['formulae'] }
  const nine = validateValue(rule(times), '9:30 AM', '9:30 AM', vhost)
  assert.equal(nine.ok, true)
  assert.ok(Math.abs((nine.value as number) - 0.3958333) < 1e-6)
  assert.equal(ok(times, '8:00', '8:00'), false)
  assert.equal(ok(times, '', 0.5), true)
  assert.equal(ok(times, '', 45292.5), true) // date-time compared on its time of day
  assert.equal(ok({ type: 'time', operator: 'greaterThan', formulae: ['17:00'] }, '6 PM', '6 PM'), true)
})

test('validation: lists (literal, range, name), case-insensitive, dropdown options', () => {
  const literal = rule({ type: 'list', formulae: ['"Yes,No, Maybe"'] })
  assert.equal(validateValue(literal, 'yes', 'yes', vhost).ok, true)
  assert.equal(validateValue(literal, 'Maybe', 'Maybe', vhost).ok, true)
  assert.equal(validateValue(literal, 'nah', 'nah', vhost).ok, false)
  assert.deepEqual(listOptionsForValidation(literal, vhost), ['Yes', 'No', 'Maybe'])
  const numbers = rule({ type: 'list', formulae: ['"1,2,3"'] })
  assert.equal(validateValue(numbers, '2', 2, vhost).ok, true)
  assert.equal(validateValue(numbers, '4', 4, vhost).ok, false)
  const range = rule({ type: 'list', formulae: ['$A$1:$A$4'] })
  assert.deepEqual(listOptionsForValidation(range, vhost), ['Red', 'Green', 'Blue'])
  assert.equal(validateValue(range, 'green', 'green', vhost).ok, true)
  assert.equal(validateValue(range, 'Purple', 'Purple', vhost).ok, false)
  const dupes = rule({ type: 'list', formulae: ['$C$1:$C$4'] })
  assert.deepEqual(listOptionsForValidation(dupes, vhost), ['x', 'y', 'Red', 'z'])
  const named = rule({ type: 'list', formulae: ['Colors'] })
  assert.deepEqual(listOptionsForValidation(named, vhost), ['Cyan', 'Magenta'])
  assert.equal(validateValue(named, 'MAGENTA', 'MAGENTA', vhost).ok, true)
  assert.deepEqual(listOptionsForValidation(rule({ type: 'list', formulae: ['Lists!$A$1:$A$2'] }), vhost), ['Cyan', 'Magenta'])
  assert.deepEqual(listOptionsForValidation(rule({ type: 'list', formulae: ['Missing!A1:A3'] }), vhost), [])
  assert.deepEqual(listOptionsForValidation(rule({ type: 'whole' }), vhost), [])
  const many = rule({ type: 'list', formulae: ['A1:A5000'] })
  const bigHost = hostFor(sheetFrom(Array.from({ length: 5000 }, (_, index) => [`item ${index}`])))
  assert.equal(listOptionsForValidation(many, bigHost).length, 1000)
  assert.equal(hasInCellDropdown(literal), true)
  assert.equal(hasInCellDropdown({ ...literal, showDropDown: true }), false)
})

test('validation: custom formulas, blanks, alerts and relative anchors', () => {
  const custom = rule({ type: 'custom', formulae: ['ISNUMBER(A1)'] })
  assert.equal(validateValue(custom, '5', 5, vhost).ok, true)
  assert.equal(validateValue(custom, 'x', 'x', vhost).ok, false)
  // Relative formula written for the anchor (B2) shifts to the edited cell (B5).
  const even = rule({ type: 'custom', formulae: ['MOD(B2,2)=0'] })
  let seen = ''
  const tracking: DataHost = { ...vhost, evaluate: (formula, row, col, override) => { seen = formula; return vhost.evaluate!(formula, row, col, override) } }
  assert.equal(validateValue(even, '4', 4, tracking, { row: 4, col: 1, anchor: { row: 1, col: 1 } }).ok, true)
  assert.equal(seen, 'MOD(B5,2)=0')
  assert.equal(validateValue(even, '3', 3, tracking, { row: 4, col: 1, anchor: { row: 1, col: 1 } }).ok, false)
  // Blanks.
  assert.equal(ok({ type: 'whole', operator: 'between', formulae: [1, 2], allowBlank: true }, '', null), true)
  assert.equal(ok({ type: 'whole', operator: 'between', formulae: [1, 2], allowBlank: false }, '', null), false)
  assert.equal(ok({ type: 'any' }, 'anything', 'anything'), true)
  // Alerts and messages.
  const stop = validateValue(rule({ type: 'whole', operator: 'between', formulae: [1, 10] }), '50', 50, vhost)
  assert.deepEqual(stop.message, { title: 'Invalid entry', text: 'Enter a whole number between 1 and 10.', style: 'stop' })
  assert.equal(stop.alert, 'stop')
  const warning = validateValue(rule({ type: 'list', formulae: ['"a,b"'], errorStyle: 'warning', errorTitle: 'Hmm', error: 'Pick a or b' }), 'c', 'c', vhost)
  assert.deepEqual([warning.alert, warning.message], ['warning', { title: 'Hmm', text: 'Pick a or b', style: 'warning' }])
  const silent = validateValue({ type: 'whole', operator: 'equal', formulae: [1] }, '2', 2, vhost)
  assert.equal(silent.ok, false)
  assert.equal(silent.alert, 'none')
  assert.equal(validateValue(null, 'x', 'x', vhost).ok, true)
})

test('validation: descriptions', () => {
  assert.equal(describeValidation({ type: 'whole', operator: 'between', formulae: [1, 10] }), 'A whole number between 1 and 10')
  assert.equal(describeValidation({ type: 'decimal', operator: 'greaterThan', formulae: ['$B$1'] }), 'A decimal number greater than =$B$1')
  assert.equal(describeValidation({ type: 'list', formulae: ['"a,b,c"'] }), 'List: a, b, c')
  assert.equal(describeValidation({ type: 'list', formulae: ['$A$1:$A$9'] }), 'List from =$A$1:$A$9')
  assert.equal(describeValidation({ type: 'date', operator: 'greaterThan', formulae: [{ type: 'date', value: '2024-01-01T00:00:00.000Z' }] }), 'A date after 1/1/2024')
  assert.equal(describeValidation({ type: 'time', operator: 'between', formulae: ['0.375', '0.75'] }), 'A time between 9:00 AM and 6:00 PM')
  assert.equal(describeValidation({ type: 'textLength', operator: 'lessThanOrEqual', formulae: [20] }), 'Text length less than or equal to 20')
  assert.equal(describeValidation({ type: 'custom', formulae: ['ISNUMBER(A1)'] }), 'Custom formula =ISNUMBER(A1)')
  assert.equal(describeValidation({ type: 'any' }), 'Any value')
})

test('validation: dialog state round trips and input checks', () => {
  const base = validationToDialogState(null)
  assert.equal(base.type, 'any')
  assert.equal(base.allowBlank, true)
  const whole = dialogStateToValidation({ ...base, type: 'whole', operator: 'between', value1: '1', value2: '10', promptTitle: 'Qty', prompt: '1 to 10' })
  assert.ok(whole.ok)
  assert.deepEqual(whole.validation, { type: 'whole', operator: 'between', formulae: [1, 10], allowBlank: true, showInputMessage: true, promptTitle: 'Qty', prompt: '1 to 10', showErrorMessage: true })
  assert.deepEqual(validationToDialogState(whole.validation), { ...base, type: 'whole', value1: '1', value2: '10', promptTitle: 'Qty', prompt: '1 to 10' })
  const swapped = dialogStateToValidation({ ...base, type: 'whole', value1: '10', value2: '1' })
  assert.deepEqual(swapped, { ok: false, error: 'The Maximum must be greater than or equal to the Minimum.', field: 'value2' })
  const fractional = dialogStateToValidation({ ...base, type: 'whole', operator: 'equal', value1: '1.5' })
  assert.ok(!fractional.ok && fractional.field === 'value1')
  const reference = dialogStateToValidation({ ...base, type: 'decimal', operator: 'lessThan', value1: '=$B$1' })
  assert.ok(reference.ok)
  assert.deepEqual(reference.validation.formulae, ['$B$1'])
  const date = dialogStateToValidation({ ...base, type: 'date', operator: 'between', value1: '1/1/2024', value2: '2024-12-31' })
  assert.ok(date.ok)
  assert.deepEqual(date.validation.formulae, [{ type: 'date', value: '2024-01-01T00:00:00.000Z' }, { type: 'date', value: '2024-12-31T00:00:00.000Z' }])
  assert.deepEqual(validationToDialogState(date.validation).value2, '12/31/2024')
  assert.ok(!dialogStateToValidation({ ...base, type: 'date', value1: 'soon', value2: '1/1/2024' }).ok)
  const time = dialogStateToValidation({ ...base, type: 'time', operator: 'greaterThan', value1: '9:00 AM' })
  assert.ok(time.ok)
  assert.deepEqual(time.validation.formulae, ['0.375'])
  assert.equal(validationToDialogState(time.validation).value1, '9:00 AM')
  const list = dialogStateToValidation({ ...base, type: 'list', value1: 'Yes, No ,Maybe', inCellDropdown: false, errorStyle: 'warning', errorTitle: 'T', error: 'E' })
  assert.ok(list.ok)
  assert.deepEqual(list.validation, { type: 'list', formulae: ['"Yes,No,Maybe"'], showDropDown: true, allowBlank: true, showInputMessage: true, showErrorMessage: true, errorStyle: 'warning', errorTitle: 'T', error: 'E' })
  assert.equal(validationToDialogState(list.validation).value1, 'Yes,No,Maybe')
  assert.equal(validationToDialogState(list.validation).inCellDropdown, false)
  const listRef = dialogStateToValidation({ ...base, type: 'list', value1: '=Sheet2!$A$1:$A$9' })
  assert.ok(listRef.ok)
  assert.deepEqual(listRef.validation.formulae, ['Sheet2!$A$1:$A$9'])
  assert.equal(validationToDialogState(listRef.validation).value1, '=Sheet2!$A$1:$A$9')
  assert.ok(!dialogStateToValidation({ ...base, type: 'list', value1: '=A1:B5' }).ok)
  assert.ok(!dialogStateToValidation({ ...base, type: 'list', value1: 'x'.repeat(300) }).ok)
  assert.ok(!dialogStateToValidation({ ...base, type: 'list', value1: ' , ' }).ok)
  const custom = dialogStateToValidation({ ...base, type: 'custom', value1: '=ISNUMBER(A1)' })
  assert.ok(custom.ok)
  assert.deepEqual(custom.validation.formulae, ['ISNUMBER(A1)'])
  assert.ok(!dialogStateToValidation({ ...base, type: 'custom', value1: '=' }).ok)
  const length = dialogStateToValidation({ ...base, type: 'textLength', operator: 'lessThanOrEqual', value1: '-1' })
  assert.ok(!length.ok)
  const any = dialogStateToValidation({ ...base, prompt: 'Hello' })
  assert.ok(any.ok)
  assert.deepEqual(any.validation, { type: 'any', showInputMessage: true, prompt: 'Hello', showErrorMessage: true })
  // Imported rules without showErrorMessage keep alerts off in the dialog? Excel shows the checkbox cleared.
  assert.equal(validationToDialogState({ type: 'whole', operator: 'equal', formulae: [1], showErrorMessage: false }).showErrorMessage, false)
  assert.equal(asValidation({ type: 'bogus' }), null)
})

test('validation: ranges, same-settings, circle invalid data', () => {
  const list = { type: 'list', formulae: ['"a,b"'], showErrorMessage: true }
  const number = { type: 'whole', operator: 'between', formulae: [1, 5], showErrorMessage: true }
  const validations: Record<string, unknown> = { 'A1:A10': list, B2: number, 'C1:C3 E1:E3': list }
  assert.deepEqual(findValidation(validations, 4, 0), { key: 'A1:A10', validation: list, anchor: { row: 0, col: 0 } })
  assert.equal(findValidation(validations, 2, 4)?.key, 'C1:C3 E1:E3')
  assert.equal(findValidation(validations, 5, 5), null)
  assert.deepEqual(rangesWithSameValidation(validations, { ...list, prompt: '' }), ['A1:A10', 'C1:C3 E1:E3'])
  assert.equal(validationsEqual({ type: 'list', allowBlank: false, formulae: ['"a"'] }, { type: 'list', formulae: ['"a"'] }), true)
  const split = withValidation(validations, { top: 3, bottom: 4, left: 0, right: 1 }, null)
  assert.deepEqual(Object.keys(split).sort(), ['A1:A3', 'A6:A10', 'B2', 'C1:C3', 'E1:E3'].sort())
  const replaced = withValidation(validations, { top: 1, bottom: 1, left: 1, right: 1 }, list as DataValidationModel)
  assert.deepEqual(replaced.B2, list)
  const sheet = sheetFrom([
    ['a', 3],
    ['c', 9],
    [null, 'x'],
    ['B', 5],
  ])
  const invalid = findInvalidCells({ 'A1:A1048576': list, 'B1:B4': number }, hostFor(sheet), { extent: { rows: 4, cols: 2 } })
  assert.deepEqual(invalid, ['A2', 'B2', 'B3'])
  assert.deepEqual(findInvalidCells({ 'A1:A4': list }, hostFor(sheet), { extent: { rows: 4, cols: 2 }, limit: 1 }), ['A2'])
  assert.deepEqual(findInvalidCells({ 'A1:B4': number }, hostFor(sheet), { extent: { rows: 4, cols: 2 }, within: { top: 0, bottom: 1, left: 1, right: 1 } }), ['B2'])
})

// ---------------------------------------------------------------------------------------------
// Cleanup
// ---------------------------------------------------------------------------------------------

test('cleanup: remove duplicates', () => {
  const sheet = sheetFrom([
    ['Name', 'City', 'Total'],
    ['Ann', 'Paris', { formula: 'C2', result: 1 }],
    ['ann', 'PARIS', 2],
    ['Bob', 'Rome', { formula: 'A4&B4', result: 'BobRome', style: red }],
    ['Ann', 'Rome', 4],
    ['Bob', 'Rome', 5],
    [null, null, null],
    [null, null, 7],
  ])
  const bounds = { top: 0, bottom: 7, left: 0, right: 2 }
  const all = removeDuplicates(sheet, bounds, { hasHeader: true }, hostFor(sheet), shift)
  assert.ok(all.ok)
  assert.equal(all.removed, 0) // Total differs everywhere
  const byNameCity = removeDuplicates(sheet, bounds, { hasHeader: true, columns: [0, 1] }, hostFor(sheet), shift)
  assert.ok(byNameCity.ok)
  assert.equal(byNameCity.removed, 3)
  assert.equal(byNameCity.remaining, 4)
  assert.deepEqual(byNameCity.keptRows, [1, 3, 4, 6])
  applyCellChanges(sheet.cells, byNameCity.changes)
  assert.deepEqual(column(sheet, 0, 0, 7), ['Name', 'Ann', 'Bob', 'Ann', null, null, null, null])
  assert.deepEqual(column(sheet, 1, 0, 4), ['City', 'Paris', 'Rome', 'Rome', null])
  assert.equal(sheet.cells.C3.formula, 'A3&B3')
  assert.deepEqual(sheet.cells.C3.style, red)
  assert.equal(sheet.cells.C6, undefined)
  const byName = removeDuplicates(sheetFrom([['x'], ['X'], ['y']]), { top: 0, bottom: 2, left: 0, right: 0 }, { hasHeader: false }, hostFor(sheetFrom([['x'], ['X'], ['y']])))
  assert.ok(byName.ok && byName.removed === 1)
  const merged = removeDuplicates({ ...sheetFrom([['a', 'b']]), merges: ['A1:B1'] }, { top: 0, bottom: 0, left: 0, right: 1 }, {}, hostFor(sheetFrom([['a', 'b']])))
  assert.ok(!merged.ok)
  assert.ok(!removeDuplicates(sheet, bounds, { columns: [9] }, hostFor(sheet)).ok)
})

test('cleanup: trim, case, fill blanks', () => {
  assert.equal(trimText('  a   b  c  '), 'a b c')
  assert.equal(trimText(' line 1 \n\n  line  2 \n'), 'line 1\n\nline 2')
  const sheet = sheetFrom([['  x  ', { formula: '" y "', result: ' y ' }, 5, 'ok']])
  const trimmed = trimWhitespace(sheet, { top: 0, bottom: 0, left: 0, right: 3 })
  assert.equal(trimmed.count, 1)
  assert.equal(trimmed.changes.A1?.value, 'x')
  assert.equal(properCase("o'neil mcDONALD-smith 2nd"), "O'Neil Mcdonald-Smith 2Nd")
  const cased = changeCase(sheetFrom([['hello world. bye NOW']]), { top: 0, bottom: 0, left: 0, right: 0 }, 'sentence')
  assert.equal(cased.changes.A1?.value, 'Hello world. Bye now')
  assert.equal(changeCase(sheetFrom([['abc']]), { top: 0, bottom: 0, left: 0, right: 0 }, 'upper').changes.A1?.value, 'ABC')
  assert.equal(changeCase(sheetFrom([['ABC']]), { top: 0, bottom: 0, left: 0, right: 0 }, 'lower').changes.A1?.value, 'abc')
  const blanks = sheetFrom([
    ['Region', { value: 45292, numFmt: 'm/d/yyyy' }],
    ['North', undefined],
    [undefined, { formula: 'B1+1' }],
    [{ value: null, style: red }, undefined],
    ['South', 1],
    [undefined, undefined],
  ])
  const filled = fillBlanksFromAbove(blanks, { top: 1, bottom: 5, left: 0, right: 1 }, shift)
  applyCellChanges(blanks.cells, filled.changes)
  assert.deepEqual(column(blanks, 0, 1, 5), ['North', 'North', 'North', 'South', 'South'])
  assert.deepEqual(blanks.cells.A4.style, red)
  assert.equal(blanks.cells.B4.formula, 'B2+1')
  assert.equal(blanks.cells.B6.value, 1)
  assert.equal(blanks.cells.B2, undefined) // nothing above B2 inside the range
  assert.equal(filled.count, 5)
})

test('cleanup: split text to columns', () => {
  const csv: Parameters<typeof splitLine>[1] = { mode: 'delimited', delimiters: { comma: true } }
  assert.deepEqual(splitLine('a,b,,c', csv), ['a', 'b', '', 'c'])
  assert.deepEqual(splitLine('a,b,,c', { ...csv, treatConsecutiveAsOne: true }), ['a', 'b', 'c'])
  assert.deepEqual(splitLine('"Smith, John",42,"say ""hi"""', csv), ['Smith, John', '42', 'say "hi"'])
  assert.deepEqual(splitLine('Smith, "Lee, Ann", 29', csv), ['Smith', 'Lee, Ann', '29'])
  assert.deepEqual(splitLine("'a,b',c", { ...csv, textQualifier: "'" }), ['a,b', 'c'])
  assert.deepEqual(splitLine('"a,b",c', { ...csv, textQualifier: '' }), ['"a', 'b"', 'c'])
  assert.deepEqual(splitLine('a  b\tc;d|e', { mode: 'delimited', delimiters: { space: true, tab: true, semicolon: true, other: '|' }, treatConsecutiveAsOne: true }), ['a', 'b', 'c', 'd', 'e'])
  assert.deepEqual(splitLine('ABCDE12345xyz', { mode: 'fixed', breaks: [5, 10] }), ['ABCDE', '12345', 'xyz'])
  assert.deepEqual(splitLine('ab', { mode: 'fixed', breaks: [5] }), ['ab', ''])
  assert.deepEqual(splitLine('abc', { mode: 'delimited' }), ['abc'])
  assert.deepEqual(previewSplit(['a,b', 'c,d,e', 'f'], csv, 2), { rows: [['a', 'b'], ['c', 'd', 'e']], columns: 3 })
  assert.deepEqual(convertField('1,234', 'general'), { value: 1234, numFmt: undefined })
  assert.deepEqual(convertField('50%', 'general'), { value: 0.5, numFmt: '0%' })
  assert.deepEqual(convertField('007', 'text'), { value: '007', numFmt: '@' })
  assert.deepEqual(convertField('2024-01-01', 'general'), { value: 45292, numFmt: 'm/d/yyyy', type: 'date' })
  assert.deepEqual(convertField('01/02/2024', 'date-dmy'), { value: serialFromYMD(2024, 2, 1), numFmt: 'm/d/yyyy', type: 'date' })
  assert.deepEqual(convertField('2024.02.01', 'date-ymd'), { value: serialFromYMD(2024, 2, 1), numFmt: 'm/d/yyyy', type: 'date' })
  assert.deepEqual(convertField('31/31/2024', 'date-mdy'), { value: '31/31/2024' })
  assert.deepEqual(convertField('TRUE', 'general'), { value: true })
  assert.deepEqual(guessDelimiters(['a;b;c', 'd;e']), { semicolon: true })
  assert.deepEqual(guessDelimiters(['a\tb', 'c\td']), { tab: true })
  const sheet = sheetFrom([
    ['Smith,John,42'],
    [{ value: 'Doe,Jane,7', style: red }],
    [null],
    ['Solo'],
  ])
  sheet.cells.C4 = { value: 'occupied' }
  const result = splitTextToColumns(sheet, { top: 0, bottom: 3, left: 0, right: 0 }, { ...csv, columnFormats: ['general', 'skip', 'text'] }, hostFor(sheet))
  assert.ok(result.ok)
  assert.equal(result.columns, 2)
  assert.equal(result.overwrites, 0)
  applyCellChanges(sheet.cells, result.changes)
  assert.deepEqual([sheet.cells.A1.value, sheet.cells.B1.value, sheet.cells.B1.numFmt], ['Smith', '42', '@'])
  assert.deepEqual(sheet.cells.A2.style, red)
  assert.equal(sheet.cells.A2.value, 'Doe')
  assert.equal(sheet.cells.A4.value, 'Solo')
  const moved = sheetFrom([['1;2'], ['3;4']])
  moved.cells.D2 = { value: 'old' }
  const elsewhere = splitTextToColumns(moved, { top: 0, bottom: 1, left: 0, right: 0 }, { mode: 'delimited', delimiters: { semicolon: true }, destination: { row: 0, col: 2 } }, hostFor(moved))
  assert.ok(elsewhere.ok)
  assert.equal(elsewhere.overwrites, 1)
  applyCellChanges(moved.cells, elsewhere.changes)
  assert.deepEqual([moved.cells.A1.value, moved.cells.C1.value, moved.cells.D1.value, moved.cells.D2.value], ['1;2', 1, 2, 4])
  assert.ok(!splitTextToColumns(moved, { top: 0, bottom: 1, left: 0, right: 1 }, csv, hostFor(moved)).ok)
})

// ---------------------------------------------------------------------------------------------
// Fill (fill handle, Fill Down/Right, Auto Fill Options) — CALC-019, calc-grid-interaction-8, CALC-006
// ---------------------------------------------------------------------------------------------

/** Fill column A (seeds from A1) down by `count` rows; returns the filled values. */
function fillDown(seeds: Input[], count: number, mode?: AutofillMode, options: { date1904?: boolean; existing?: Record<string, CellData> } = {}) {
  const sheet = sheetFrom(seeds.map((seed) => [seed]))
  Object.assign(sheet.cells, options.existing || {})
  const patch = createAutofillPatch({
    cells: sheet.cells,
    source: { top: 0, bottom: seeds.length - 1, left: 0, right: 0 },
    destination: { top: seeds.length, bottom: seeds.length + count - 1, left: 0, right: 0 },
    mode,
    date1904: options.date1904,
  })
  const values = Array.from({ length: count }, (_, index) => patch.changes[addressOf(seeds.length + index, 0)]?.value)
  return { patch, values }
}

const DATE = { numFmt: 'm/d/yyyy' }
const dateCell = (year: number, month: number, day: number): CellData => ({ value: serialFromYMD(year, month, day), ...DATE })
const ymd = (year: number, month: number, day: number) => serialFromYMD(year, month, day)

test('fill: day and month names continue around their list, keeping case', () => {
  assert.deepEqual(fillDown(['Mon'], 7).values, ['Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun', 'Mon'])
  assert.deepEqual(fillDown(['monday'], 2).values, ['tuesday', 'wednesday'])
  assert.deepEqual(fillDown(['FRIDAY'], 3).values, ['SATURDAY', 'SUNDAY', 'MONDAY'])
  assert.deepEqual(fillDown(['November'], 3).values, ['December', 'January', 'February'])
  assert.deepEqual(fillDown(['Jan'], 2).values, ['Feb', 'Mar'])
  assert.deepEqual(fillDown(['May'], 1).values, ['Jun'], 'May reads as the short month list first, as in Excel')
  assert.deepEqual(fillDown(['April', 'May'], 1).values, ['June'])
  assert.deepEqual(fillDown(['Jan', 'Mar'], 3).values, ['May', 'Jul', 'Sep'], 'two seeds keep their step')
  assert.deepEqual(fillDown(['Wed', 'Mon'], 2).values, ['Sat', 'Thu'], 'a backwards step continues backwards')
  assert.deepEqual(fillDown(['Sun', 'Sept'], 2).values, ['Sun', 'Sept'], 'unknown names repeat')
})

test('fill: quarters wrap after Q4 in every spelling', () => {
  assert.deepEqual(fillDown(['Q1'], 4).values, ['Q2', 'Q3', 'Q4', 'Q1'])
  assert.deepEqual(fillDown(['Qtr 3'], 2).values, ['Qtr 4', 'Qtr 1'])
  assert.deepEqual(fillDown(['Quarter 4'], 1).values, ['Quarter 1'])
  assert.deepEqual(fillDown(['1st Quarter'], 4).values, ['2nd Quarter', '3rd Quarter', '4th Quarter', '1st Quarter'])
  assert.deepEqual(fillDown(['q2', 'q4'], 2).values, ['q2', 'q4'], 'a two-quarter step repeats around the year')
  assert.deepEqual(fillDown(['Q5'], 1).values, ['Q6'], 'Q5 is not a quarter: its number counts up')
})

test('fill: text with a trailing number counts up from a single seed (Excel)', () => {
  assert.deepEqual(fillDown(['Item 1'], 3).values, ['Item 2', 'Item 3', 'Item 4'])
  assert.deepEqual(fillDown(['A-009'], 2).values, ['A-010', 'A-011'])
  assert.deepEqual(fillDown(['Item 1', 'Item 3'], 2).values, ['Item 5', 'Item 7'])
  assert.deepEqual(fillDown(['Item 1'], 2, 'toggle').values, ['Item 1', 'Item 1'], 'Ctrl+drag copies')
  assert.deepEqual(fillDown(['Total'], 2).values, ['Total', 'Total'])
})

test('fill: numbers copy alone, count with Ctrl or Fill Series, and extend steps and trends', () => {
  assert.deepEqual(fillDown([5], 3).values, [5, 5, 5])
  assert.deepEqual(fillDown([5], 3, 'toggle').values, [6, 7, 8], 'Ctrl+drag a lone number counts up')
  assert.deepEqual(fillDown([5], 2, 'series').values, [6, 7])
  assert.deepEqual(fillDown([1, 3], 2).values, [5, 7])
  assert.deepEqual(fillDown([1, 3], 2, 'toggle').values, [1, 3], 'Ctrl+drag a series copies it')
  assert.deepEqual(fillDown([0.1, 0.2], 2).values, [0.3, 0.4], 'steps are stored to 15 digits')
  const trend = fillDown([1, 2, 4], 2).values as number[]
  assert.ok(Math.abs(trend[0] - 16 / 3) < 1e-9 && Math.abs(trend[1] - 41 / 6) < 1e-9, `uneven seeds follow the best-fit line (${trend})`)
  assert.deepEqual(fillDown([2, 2], 2).values, [2, 2])
})

test('fill: dates step by days, months and years', () => {
  const single = fillDown([dateCell(2026, 1, 15)], 2)
  assert.deepEqual(single.values, [ymd(2026, 1, 16), ymd(2026, 1, 17)], 'a lone date advances a day')
  assert.equal(single.patch.changes.A2?.numFmt, 'm/d/yyyy', 'the date format travels with the value')
  assert.deepEqual(fillDown([dateCell(2026, 1, 15)], 2, 'toggle').values, [ymd(2026, 1, 15), ymd(2026, 1, 15)], 'Ctrl+drag copies a date')
  assert.deepEqual(fillDown([dateCell(2026, 1, 15), dateCell(2026, 2, 15)], 2).values, [ymd(2026, 3, 15), ymd(2026, 4, 15)], 'one month apart steps by months')
  assert.deepEqual(fillDown([dateCell(2026, 1, 31), dateCell(2026, 2, 28)], 3).values, [ymd(2026, 3, 31), ymd(2026, 4, 30), ymd(2026, 5, 31)], 'month ends stay month ends (EDATE)')
  assert.deepEqual(fillDown([dateCell(2024, 1, 1), dateCell(2025, 1, 1)], 1).values, [ymd(2026, 1, 1)], 'a year apart steps by years')
  assert.deepEqual(fillDown([dateCell(2026, 1, 1), dateCell(2026, 1, 8)], 1).values, [ymd(2026, 1, 15)], 'a week apart steps by days')
  assert.deepEqual(fillDown([dateCell(2026, 1, 2)], 3, 'weekdays').values, [ymd(2026, 1, 5), ymd(2026, 1, 6), ymd(2026, 1, 7)], 'Fill Weekdays skips the weekend after Friday 2 Jan 2026')
  assert.deepEqual(fillDown([dateCell(2026, 1, 31)], 2, 'months').values, [ymd(2026, 2, 28), ymd(2026, 3, 31)])
  assert.deepEqual(fillDown([dateCell(2024, 2, 29)], 1, 'years').values, [ymd(2025, 2, 28)])
  assert.deepEqual(fillDown([dateCell(2026, 1, 15)], 1, 'days').values, [ymd(2026, 1, 16)])
  assert.equal(addMonthsToSerial(ymd(2026, 3, 31), -1), ymd(2026, 2, 28))
  assert.equal(addWeekdaysToSerial(ymd(2026, 1, 5), -1), ymd(2026, 1, 2))
  const time = fillDown([{ value: 0.375, numFmt: 'h:mm' }], 2).values as number[]
  assert.ok(Math.abs(time[0] - 10 / 24) < 1e-9 && Math.abs(time[1] - 11 / 24) < 1e-9, 'a lone time adds an hour')
  assert.deepEqual(single.patch.options, ['copy', 'series', 'formats', 'values', 'days', 'weekdays', 'months', 'years'])
  assert.deepEqual(fillDown([5], 1).patch.options, ['copy', 'series', 'formats', 'values'])
})

test('fill: Copy cells (Fill Down / Right) copies dates verbatim and shifts formulas', () => {
  const sheet = sheetFrom([[dateCell(2026, 1, 15), { formula: 'A1*2' }]])
  const patch = createAutofillPatch({ cells: sheet.cells, source: { top: 0, bottom: 0, left: 0, right: 1 }, destination: { top: 1, bottom: 3, left: 0, right: 1 }, mode: 'copy' })
  assert.equal(patch.changes.A2?.value, ymd(2026, 1, 15))
  assert.equal(patch.changes.A4?.value, ymd(2026, 1, 15), 'Ctrl+D never turns a date into a series')
  assert.equal(patch.changes.B3?.formula, 'A3*2')
  const right = createAutofillPatch({ cells: sheetFrom([['Mon']]).cells, source: { top: 0, bottom: 0, left: 0, right: 0 }, destination: { top: 0, bottom: 0, left: 1, right: 2 }, mode: 'copy' })
  assert.deepEqual([right.changes.B1?.value, right.changes.C1?.value], ['Mon', 'Mon'])
})

test('fill: formatting only and without formatting', () => {
  const bold = { font: { bold: true } }
  const seeds = [{ value: 1, style: bold, numFmt: '0.00' }, { value: 2, style: bold, numFmt: '0.00' }]
  const existing = { A3: { value: 'keep', style: { fill: { pattern: 'solid', fgColor: { argb: 'FFFF0000' } } } } }
  const formats = fillDown(seeds, 2, 'formats', { existing })
  assert.deepEqual(formats.patch.changes.A3, { value: 'keep', style: bold, numFmt: '0.00' }, 'contents stay, the look comes from the seeds')
  assert.deepEqual(formats.patch.changes.A4, { style: bold, numFmt: '0.00' })
  const values = fillDown(seeds, 2, 'values', { existing })
  assert.deepEqual(values.patch.changes.A3, { value: 3, style: existing.A3.style }, 'the series fills, the destination keeps its look')
  assert.deepEqual(values.patch.changes.A4, { value: 4 })
})

test('fill: up and left run the series backwards', () => {
  const sheet = sheetFrom([[], [], ['Wed']])
  const patch = createAutofillPatch({ cells: sheet.cells, source: { top: 2, bottom: 2, left: 0, right: 0 }, destination: { top: 0, bottom: 1, left: 0, right: 0 } })
  assert.deepEqual([patch.changes.A1?.value, patch.changes.A2?.value], ['Mon', 'Tue'])
})

test('fill: double-clicking the fill handle follows the neighbouring data', () => {
  const sheet = sheetFrom(Array.from({ length: 10 }, (_, row) => [row + 1, row === 0 ? { formula: 'A1*2' } : undefined]))
  const has = (cells: Record<string, CellData>) => (row: number, col: number) => cellFilled(cells[addressOf(row, col)])
  assert.equal(fillHandleDoubleClickBottom(has(sheet.cells), { top: 0, bottom: 0, left: 1, right: 1 }), 9, 'fills B2:B10 beside A1:A10')
  sheet.cells.B6 = { value: 'x' }
  assert.equal(fillHandleDoubleClickBottom(has(sheet.cells), { top: 0, bottom: 0, left: 1, right: 1 }), 4, 'stops above data already in the column')
  const right = sheetFrom([[{ formula: 'B1' }, 1], [undefined, 2], [undefined, 3]])
  assert.equal(fillHandleDoubleClickBottom(has(right.cells), { top: 0, bottom: 0, left: 0, right: 0 }), 2, 'uses the column to the right when the left is empty')
  assert.equal(fillHandleDoubleClickBottom(has(sheetFrom([[1]]).cells), { top: 0, bottom: 0, left: 0, right: 0 }), null, 'nothing beside: nothing to fill')
})

function cellFilled(cell: CellData | undefined) {
  return Boolean(cell && (cell.formula || (cell.value !== undefined && cell.value !== null && cell.value !== '')))
}

// ---------------------------------------------------------------------------------------------
// Filtered ranges: visible cells only (calc-grid-interaction-2, CALC-002)
// ---------------------------------------------------------------------------------------------

test('filter: filter-hidden rows for visible-cells-only commands', () => {
  assert.deepEqual([...filterHiddenRowSet({ hiddenRows: [3, 5, 7], filteredRows: [3, 5, 9] })].sort((a, b) => a - b), [2, 4], 'only rows a filter hides (and that are still hidden), 0-based')
  assert.equal(filterHiddenRowSet({ hiddenRows: [3], filteredRows: undefined }).size, 0, 'manually hidden rows are not skipped')
})

// ---------------------------------------------------------------------------------------------
// Sort Warning (CALC-001)
// ---------------------------------------------------------------------------------------------

test('sort: a one-column selection beside data offers to expand', () => {
  const sheet = sheetFrom([['Name', 'Qty', 'Price'], ['b', 2, 20], ['a', 1, 10], ['c', 3, 30]])
  assert.deepEqual(sortExpansionRegion(sheet, { top: 0, bottom: 3, left: 1, right: 1 }), { top: 0, bottom: 3, left: 0, right: 2 })
  assert.deepEqual(sortExpansionRegion(sheet, { top: 1, bottom: 3, left: 0, right: 1 }), { top: 0, bottom: 3, left: 0, right: 2 }, 'a partial block grows to the whole region')
  assert.equal(sortExpansionRegion(sheet, { top: 0, bottom: 3, left: 0, right: 2 }), null, 'the whole region needs no warning')
  assert.equal(sortExpansionRegion(sheet, { top: 1, bottom: 1, left: 1, right: 1 }), null, 'a single cell expands on its own')
  assert.deepEqual(sortExpansionRegion(sheet, { top: 0, bottom: 1_048_575, left: 1, right: 1 }), { top: 0, bottom: 3, left: 0, right: 2 }, 'a whole column is clipped to the data')
  const alone = sheetFrom([[3, undefined, 'x'], [1, undefined, 'y'], [2]])
  assert.equal(sortExpansionRegion(alone, { top: 0, bottom: 2, left: 0, right: 0 }), null, 'a blank column apart: nothing beside it')
  assert.deepEqual(currentRegionAround(sheet, { top: 2, bottom: 2, left: 1, right: 1 }), { top: 0, bottom: 3, left: 0, right: 2 })
})

// ---------------------------------------------------------------------------------------------
// Dropdown picks store typed values (calc-grid-interaction-9)
// ---------------------------------------------------------------------------------------------

test('validation: a dropdown pick maps to the source value', () => {
  const sheet = sheetFrom([
    [{ value: 10, numFmt: '$#,##0.00', display: '$10.00' }],
    [{ value: serialFromYMD(2026, 1, 15), numFmt: 'm/d/yyyy', display: '1/15/2026' }],
    [{ value: true }],
    [{ value: '00123' }],
  ])
  const host = hostFor(sheet)
  const ranged: DataValidationModel = { type: 'list', formulae: ['$A$1:$A$4'] }
  assert.deepEqual(listEntryForOption(ranged, '$10.00', host), { value: 10, text: '$10.00', numFmt: '$#,##0.00', literal: false })
  assert.deepEqual(listEntryForOption(ranged, '1/15/2026', host), { value: serialFromYMD(2026, 1, 15), text: '1/15/2026', numFmt: 'm/d/yyyy', literal: false })
  assert.equal(listEntryForOption(ranged, 'TRUE', host)?.value, true)
  assert.equal(listEntryForOption(ranged, '00123', host)?.value, '00123', 'text in the source stays text')
  const literal: DataValidationModel = { type: 'list', formulae: ['"1,2,3"'] }
  assert.deepEqual(listEntryForOption(literal, '2', host), { value: 2, text: '2', literal: true })
  assert.equal(listEntryForOption(literal, '9', host), null)
})

// ---------------------------------------------------------------------------------------------
// Performance (100k rows)
// ---------------------------------------------------------------------------------------------

test('performance: 100k-row filter, checklist and sort', () => {
  const rows = 100_000
  const names = ['alpha', 'Bravo', 'charlie', 'Delta', 'echo', 'foxtrot', 'Golf', 'hotel']
  const values: Scalar[] = new Array(rows)
  const texts: string[] = new Array(rows)
  const numbers: number[] = new Array(rows)
  for (let index = 0; index < rows; index += 1) {
    values[index] = index
    texts[index] = `${names[index % names.length]} ${(index * 7919) % 50_000}`
    numbers[index] = ((index * 104_729) % 99_991) / 7
  }
  // Array-backed host: the lead's host reads the engine's caches, so string addressing is not measured here.
  const host: DataHost = {
    valueAt: (row, col) => row === 0 ? 'Header' : col === 0 ? texts[row - 1] : numbers[row - 1],
    displayAt: (row, col) => row === 0 ? 'Header' : col === 0 ? texts[row - 1] : String(numbers[row - 1]),
    fillColorAt: () => null,
    fontColorAt: () => null,
    isDateAt: () => false,
  }
  const state: SheetFilterState = { ref: `A1:B${rows + 1}`, columns: { 0: { condition: { operator: 'contains', value: 'a' } }, 1: { condition: { operator: 'top', value: 5000 } } } }
  const time = (label: string, run: () => void, budget: number) => {
    run() // warm up the JIT
    const started = performance.now()
    run()
    const elapsed = performance.now() - started
    process.stdout.write(`  ${label}: ${elapsed.toFixed(1)} ms\n`)
    assert.ok(elapsed < budget, `${label} took ${elapsed.toFixed(1)} ms (budget ${budget} ms)`)
  }
  let hidden: number[] = []
  time('filter 100k rows (contains + top 5000)', () => { hidden = computeFilteredRows(state, host) }, 100)
  assert.ok(hidden.length > 90_000)
  const valuesState: SheetFilterState = { ref: `A1:B${rows + 1}`, columns: { 0: { values: texts.slice(0, 2000) } } }
  time('filter 100k rows (2000-value checklist)', () => { computeFilteredRows(valuesState, host) }, 100)
  let distinct = 0
  time('checklist 100k rows (~50k distinct)', () => { distinct = distinctColumnValues({ ref: state.ref, columns: {} }, 0, host).items.length }, 250)
  assert.ok(distinct > 40_000)
  const bounds = { top: 0, bottom: rows, left: 0, right: 1 }
  let order: number[] = []
  time('sort 100k rows by text then number', () => {
    const result = computeSortOrder({ bounds, hasHeader: true, levels: [{ key: 0 }, { key: 1, descending: true }] }, host)
    assert.ok(result.ok)
    order = result.order
  }, 150)
  time('sort 100k rows by number', () => {
    const result = computeSortOrder({ bounds, hasHeader: true, levels: [{ key: 1 }] }, host)
    assert.ok(result.ok)
  }, 100)
  for (let index = 1; index < 200; index += 1) {
    const previous = texts[order[index - 1] - 1]
    const current = texts[order[index] - 1]
    assert.ok(previous.localeCompare(current, undefined, { sensitivity: 'accent' }) <= 0)
  }
  void values
})

test('performance: sheet-backed current region and full sort with cell changes', () => {
  const rows = 100_000
  const cells: Record<string, CellData> = {}
  for (let row = 0; row < rows; row += 1) {
    cells[addressOf(row, 0)] = { value: `name ${(row * 7919) % 30_000}` }
    cells[addressOf(row, 1)] = { value: (row * 104_729) % 99_991 }
    cells[addressOf(row, 2)] = row % 3 ? { formula: `B${row + 1}*2` } : { value: 'x', style: red }
  }
  const sheet: SheetData = { id: 's', name: 'Big', rowCount: rows, colCount: 3, cells, merges: [], colWidths: {}, rowHeights: {} }
  const host = createSheetHost(sheet)
  let started = performance.now()
  const region = detectCurrentRegion(sheet, 500, 1)
  const regionMs = performance.now() - started
  assert.deepEqual(region, { top: 0, bottom: rows - 1, left: 0, right: 2 })
  started = performance.now()
  const result = sortRange(sheet, { bounds: region!, levels: [{ key: 0 }, { key: 1, descending: true }] }, host, shift)
  const sortMs = performance.now() - started
  assert.ok(result.ok && result.changed)
  started = performance.now()
  applyCellChanges(sheet.cells, result.changes)
  const applyMs = performance.now() - started
  process.stdout.write(`  current region 100k×3 (string-addressed): ${regionMs.toFixed(1)} ms
  sortRange 100k×3 incl. cell changes: ${sortMs.toFixed(1)} ms (apply ${applyMs.toFixed(1)} ms)
`)
  assert.ok(regionMs < 400 && sortMs < 1500)
  const check = createSheetHost(sheet)
  for (let row = 1; row < 500; row += 1) {
    const previous = String(check.valueAt(row - 1, 0))
    const current = String(check.valueAt(row, 0))
    assert.ok(previous.localeCompare(current, undefined, { sensitivity: 'accent' }) <= 0, `${previous} > ${current}`)
    const formula = sheet.cells[addressOf(row, 2)]?.formula
    if (formula) assert.equal(formula, `B${row + 1}*2`)
  }
})

// ---------------------------------------------------------------------------------------------
// Components render (server-side smoke test: markup, roles, labels)
// ---------------------------------------------------------------------------------------------

test('components: render with typical props', () => {
  const originalError = console.error
  const logged: string[] = []
  console.error = (...args: unknown[]) => { logged.push(args.map(String).join(' ')) }
  try {
    const sheet = filterFixture()
    const host = hostFor(sheet)
    const state: SheetFilterState = { ref: 'A1:D11', columns: { 1: { condition: { operator: 'greaterThan', value: 5 } } } }
    const noop = () => undefined
    const filterHtml = renderToString(createElement(FilterMenu, { columnLabel: 'Name', criteria: { values: ['apple'] }, values: distinctColumnValues(state, 0, host), anchor: { left: 10, top: 10, right: 30, bottom: 30 }, onApply: noop, onSort: noop, onSortByColor: noop, onClose: noop }))
    assert.match(filterHtml, /role="dialog"/)
    assert.match(filterHtml, /aria-modal="true"/)
    assert.match(filterHtml, /Filter by values/)
    assert.match(filterHtml, /Clear filter from/)
    const conditionHtml = renderToString(createElement(FilterMenu, { columnLabel: 'Qty', criteria: state.columns[1], values: distinctColumnValues(state, 1, host), anchor: { left: 10, top: 10, right: 30, bottom: 30 }, onApply: noop, onSort: noop, onClose: noop }))
    assert.match(conditionHtml, /Sort smallest to largest/)
    assert.match(conditionHtml, /value="greaterThan" selected=""/)
    const bounds = { top: 0, bottom: 10, left: 0, right: 3 }
    const sortHtml = renderToString(createElement(SortDialog, { rangeLabel: 'A1:D11', initialHasHeader: true, initialLevels: [{ key: 1, descending: true }, { key: 0, sortOn: 'fillColor', color: '#FF0000' }], labelsFor: (orientation, hasHeader) => sortKeyLabels(bounds, orientation, hasHeader, host), colorsFor: (orientation, key, sortOn, hasHeader) => sortKeyColors(bounds, orientation, key, sortOn, hasHeader, host), onSort: noop, onClose: noop }))
    assert.match(sortHtml, /Then by/)
    assert.match(sortHtml, /My data has headers/)
    assert.match(sortHtml, />Qty</)
    const validationHtml = renderToString(createElement(DataValidationDialog, { rangeLabel: 'B2:B11', initial: { type: 'list', formulae: ['"a,b"'], showErrorMessage: true }, sameSettingsCount: 1, onApply: noop, onClose: noop }))
    assert.match(validationHtml, /role="tablist"/)
    assert.match(validationHtml, /value="a,b"/)
    assert.match(validationHtml, /In-cell dropdown/)
    const dupesHtml = renderToString(createElement(RemoveDuplicatesDialog, { rangeLabel: 'A1:D11', initialHasHeader: true, labelsFor: (hasHeader: boolean) => sortKeyLabels(bounds, 'rows', hasHeader, host), onApply: () => ({ removed: 1, remaining: 9 }), onClose: noop }))
    assert.match(dupesHtml, /Name/)
    const splitHtml = renderToString(createElement(TextToColumnsDialog, { sourceLabel: 'A1:A3', sampleLines: ['a,b', 'c,d'], defaultDestination: 'A1', onApply: noop, onClose: noop }))
    assert.match(splitHtml, /Delimited/)
    assert.match(splitHtml, /<td[^>]*>b<\/td>/)
  } finally {
    console.error = originalError
  }
  const unexpected = logged.filter((message) => !/useLayoutEffect does nothing on the server/.test(message))
  assert.deepEqual(unexpected, [])
})

process.stdout.write(`Data-tools QA passed: ${checks} groups (filter, sort, validation, cleanup, performance, components).\n`)
