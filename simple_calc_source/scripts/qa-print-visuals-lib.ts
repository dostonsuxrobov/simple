// QA library for print/export visuals (bundled by scripts/qa-print-visuals.cjs and
// scripts/qa-exports.cjs). Builds a sheet with conditional formatting (colour scale, data
// bar, icon set, number format), a banded table, sparklines and checkboxes, resolves its
// print visuals exactly as the app does, and checks that the shared visual resolution
// paints cells the way the grid's inline layering does.
import assert from 'node:assert/strict'
import { CalculationEngine } from '../src/lib/calc-engine'
import '../src/lib/formula-library'
import { dataBarBackground } from '../src/lib/conditional-format'
import type { ConditionalCellFormat } from '../src/lib/conditional-format'
import { formatColor } from '../src/lib/number-format'
import { buildPrintVisuals, conditionalIconMarkup, printableBounds } from '../src/lib/print-visuals'
import {
  applyPageSetupPatch,
  headerFooterCode,
  headerFooterTexts,
  normalizeTitleColumns,
  normalizeTitleRows,
  pageSetupStateFor,
  splitHeaderFooter,
} from '../src/lib/print-page-setup'
import type { TableCellPaint } from '../src/lib/table-styles'
import { applyCellVisualStyle, isEmptyVisual, resolveCellVisual } from '../src/lib/visual-style'
import type { VisualStyleRecord } from '../src/lib/visual-style'
import type { CellData, SheetData, WorkbookModel } from '../src/spreadsheet-types'

import { EXPECTED_COLORS, visualPrintInput, visualWorkbook } from './qa-print-visuals-fixture'

export { applyPageSetupPatch, buildPrintVisuals, pageSetupStateFor, printableBounds }
export { EXPECTED_COLORS, visualPrintInput, visualWorkbook }

// ---------------------------------------------------------------------------------------------
// Grid parity: the grid's inline layering (App.tsx GridCellView) as an oracle
// ---------------------------------------------------------------------------------------------

type Style = Record<string, string | number | undefined>

function gridOracle(base: Style, cell: CellData | undefined, rawValue: unknown, tablePaint: TableCellPaint | null | undefined, conditional: ConditionalCellFormat | undefined, validationOptions: string[]) {
  const paintedStyle: Style = { ...base }
  if (tablePaint !== undefined) {
    const paint = tablePaint
    const own = cell?.style
    if (paint) {
      if (paint.fill && !own?.fill?.pattern && !own?.fill?.type) paintedStyle.backgroundColor = paint.fill
      if (paint.color && !own?.font?.color) paintedStyle.color = paint.color
      if (paint.bold && own?.font?.bold === undefined) paintedStyle.fontWeight = 700
      if (paint.borderTop && !paintedStyle.borderTop) paintedStyle.borderTop = paint.borderTop
      if (paint.borderBottom && !paintedStyle.borderBottom) paintedStyle.borderBottom = paint.borderBottom
      if (paint.borderLeft && !paintedStyle.borderLeft) paintedStyle.borderLeft = paint.borderLeft
      if (paint.borderRight && !paintedStyle.borderRight) paintedStyle.borderRight = paint.borderRight
    }
  }
  const formatTextColor = typeof rawValue === 'number' || typeof rawValue === 'string' ? formatColor(rawValue, conditional?.numFmt || cell?.numFmt || cell?.style?.numFmt) : undefined
  if (formatTextColor) paintedStyle.color = formatTextColor
  if (conditional) {
    if (conditional.fill || conditional.colorScale) {
      paintedStyle.backgroundColor = conditional.colorScale || conditional.fill
      paintedStyle.backgroundImage = undefined
    }
    if (conditional.font?.color) paintedStyle.color = conditional.font.color
    if (conditional.font?.bold !== undefined) paintedStyle.fontWeight = conditional.font.bold ? 700 : 400
    if (conditional.font?.italic !== undefined) paintedStyle.fontStyle = conditional.font.italic ? 'italic' : 'normal'
    if (conditional.font?.underline || conditional.font?.strike) {
      paintedStyle.textDecorationLine = [conditional.font.underline ? 'underline' : '', conditional.font.strike ? 'line-through' : ''].filter(Boolean).join(' ')
    }
    if (conditional.border?.top) paintedStyle.borderTop = conditional.border.top
    if (conditional.border?.right) paintedStyle.borderRight = conditional.border.right
    if (conditional.border?.bottom) paintedStyle.borderBottom = conditional.border.bottom
    if (conditional.border?.left) paintedStyle.borderLeft = conditional.border.left
    if (conditional.dataBar) {
      const bar = dataBarBackground(conditional.dataBar)
      const baseColor = paintedStyle.backgroundColor ? String(paintedStyle.backgroundColor) : ''
      paintedStyle.backgroundImage = bar.backgroundImage
      paintedStyle.backgroundSize = bar.backgroundSize
      paintedStyle.backgroundPosition = bar.backgroundPosition
      paintedStyle.backgroundRepeat = bar.backgroundRepeat
      if (baseColor) paintedStyle.backgroundColor = baseColor
    }
  }
  const hide = Boolean(conditional && ((conditional.dataBar && !conditional.dataBar.showValue) || (conditional.icon && !conditional.icon.showValue)))
  const isCheckbox = cell?.type === 'checkbox' || (
    typeof rawValue === 'boolean' && validationOptions.length === 2 &&
    validationOptions[0].toLocaleUpperCase() === 'TRUE' && validationOptions[1].toLocaleUpperCase() === 'FALSE'
  )
  const checked = rawValue === true || String(rawValue).toLocaleUpperCase() === 'TRUE'
  return { style: paintedStyle, hide, isCheckbox, checked, icon: conditional?.icon ? { set: conditional.icon.set, index: conditional.icon.index } : undefined, numFmt: conditional?.numFmt && typeof rawValue === 'number' ? conditional.numFmt : undefined }
}

function defined(style: Style): Style {
  return Object.fromEntries(Object.entries(style).filter(([, value]) => value !== undefined))
}

/** Compares resolveCellVisual + applyCellVisualStyle with the grid oracle over a matrix of cases. */
export function checkGridParity(): number {
  const ownStyles: Array<{ cell?: CellData; base: Style }> = [
    { cell: undefined, base: {} },
    { cell: { value: 1 }, base: {} },
    { cell: { value: 1, style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEEDD' } } } }, base: { backgroundColor: '#FFEEDD' } },
    { cell: { value: 1, style: { fill: { type: 'pattern', pattern: 'none' } } }, base: {} },
    { cell: { value: 1, style: { font: { bold: false, color: { argb: 'FF123456' } } } }, base: { color: '#123456' } },
    { cell: { value: 1, style: { font: { bold: true } } }, base: { fontWeight: 700 } },
    { cell: { value: 1, style: { border: { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'none' } } } }, base: { borderTop: '1px solid #000000' } },
    { cell: { value: 1, style: { border: { diagonalUp: true, diagonal: { style: 'thin' } } as never } }, base: { backgroundImage: 'linear-gradient(to top right, transparent calc(50% - .5px), #4b4b4b 50%, transparent calc(50% + .5px))' } },
    { cell: { value: -3, numFmt: '0;[Red]-0' }, base: {} },
    { cell: { value: 'text', numFmt: '@;[Blue]@' }, base: {} },
    { cell: { value: true, type: 'checkbox' }, base: {} },
    { cell: { value: 'TRUE', type: 'checkbox' }, base: {} },
    { cell: { value: true }, base: {} },
    { cell: { value: 1, hyperlink: 'https://example.com' }, base: {} },
  ]
  const tables: Array<TableCellPaint | null | undefined> = [
    undefined,
    null,
    { fill: '#dae3f3', borderTop: '1px solid #8faadc', borderBottom: '1px solid #8faadc' },
    { fill: '#4472C4', color: '#ffffff', bold: true, borderLeft: '1px solid #8faadc', borderRight: '2px solid #ffffff' },
  ]
  const conditionals: Array<ConditionalCellFormat | undefined> = [
    undefined,
    { fill: '#FFC7CE', font: { color: '#9C0006' } },
    { fill: '#F8696B', colorScale: '#F8696B' },
    { font: { bold: false, italic: true, underline: true } },
    { font: { strike: true, underline: false }, border: { top: '2px solid #000000', left: '1px dotted #FF0000' } },
    { dataBar: { fraction: 0.4, color: '#638EC6', gradient: true, showValue: true } },
    { dataBar: { fraction: 0.6, color: '#FF0000', negative: true, axis: 0.3, axisColor: '#000000', gradient: false, border: '#AA0000', showValue: false, rtl: true }, fill: '#FFFFCC' },
    { icon: { set: '3Arrows', index: 2, showValue: true } },
    { icon: { set: '5Rating', index: 0, showValue: false }, numFmt: '0.0%' },
    { numFmt: '[Green]0.00' },
  ]
  const listOptions: string[][] = [[], ['TRUE', 'FALSE'], ['true', 'false'], ['Yes', 'No']]
  let cases = 0
  for (const { cell, base } of ownStyles) {
    for (const table of tables) {
      for (const conditional of conditionals) {
        for (const options of listOptions) {
          const value = cell?.formula ? cell.result : cell?.value
          const oracle = gridOracle(base, cell, value, table, conditional, options)
          const visual = resolveCellVisual({ cell, value, table: table ?? null, conditional, listOptions: options })
          const style = applyCellVisualStyle({ ...base } as VisualStyleRecord, visual, dataBarBackground)
          const label = JSON.stringify({ cell, table, conditional, options })
          assert.deepEqual(defined(style), defined(oracle.style), `grid style parity: ${label}`)
          // replacesFillImage clears the image even where the oracle had none.
          assert.equal('backgroundImage' in style && style.backgroundImage === undefined, Boolean(conditional?.fill || conditional?.colorScale) && !conditional?.dataBar, `fill image cleared: ${label}`)
          assert.equal(Boolean(visual.hideValue), oracle.hide, `hidden value parity: ${label}`)
          assert.equal(Boolean(visual.checkbox), oracle.isCheckbox, `checkbox parity: ${label}`)
          if (visual.checkbox) assert.equal(visual.checkbox.checked, oracle.checked, `checkbox state parity: ${label}`)
          assert.deepEqual(visual.icon, oracle.icon, `icon parity: ${label}`)
          assert.equal(visual.numFmt, oracle.numFmt, `conditional number format parity: ${label}`)
          cases += 1
        }
      }
    }
  }
  assert.equal(isEmptyVisual(resolveCellVisual({ cell: { value: 'plain' }, value: 'plain' })), true, 'a plain cell adds nothing')
  return cases
}

/** The resolved payload for the fixture: what the print engine receives. */
export function checkVisualPayload() {
  const input = visualPrintInput()
  const sheet = input.visuals.sheets.report
  assert.ok(sheet, 'the printed sheet has visuals')
  const style = (address: string) => sheet.styles[sheet.cells[address]]
  assert.equal(style('B2').fill, EXPECTED_COLORS.scaleLow, 'colour scale minimum')
  assert.equal(style('B6').fill, EXPECTED_COLORS.scaleHigh, 'colour scale maximum')
  assert.match(style('C2').bar?.image || '', /linear-gradient\(90deg, #638EC6/, 'data bar')
  assert.match(style('C6').bar?.size || '', /^90\.000% /, 'the largest bar reaches Excel\'s 90% maximum length')
  assert.equal(style('D2').icon, '3TrafficLights1:0')
  assert.equal(style('D6').icon, '3TrafficLights1:2')
  assert.match(input.visuals.icons['3TrafficLights1:0'], /fill="#D6392B"/, 'icons are the grid\'s ConditionalIcon glyphs')
  assert.equal(input.visuals.icons['3TrafficLights1:2'], conditionalIconMarkup('3TrafficLights1', 2))
  assert.equal(style('A1').fill, EXPECTED_COLORS.tableHeader, 'table header fill')
  assert.equal(style('A1').bold, true)
  assert.equal(style('A1').color?.toLowerCase(), '#ffffff')
  assert.ok(style('A2').fill && style('A2').fill !== style('A3').fill, 'banded rows alternate')
  assert.equal(style('A4').fill, style('A2').fill, 'every other row shares the band colour')
  assert.equal(style('G2').checkbox, true, 'checkbox control (checked)')
  assert.equal(style('G3').checkbox, false, 'checkbox control (clear)')
  assert.equal(style('H2').checkbox, true, 'TRUE/FALSE dropdown prints as a checkbox, as in Google Sheets')
  assert.equal(style('H3').checkbox, false)
  assert.equal(style('I2').color, EXPECTED_COLORS.negativeRed, 'number format colour')
  assert.equal(sheet.text?.J2, '25.0%', 'a conditional number format changes the printed value')
  assert.match(sheet.sparklines?.F2 || '', /<polyline[^>]+stroke="#1a73e8"/, 'SPARKLINE() result')
  for (const row of [2, 3, 4, 5, 6]) assert.match(sheet.sparklines?.[`E${row}`] || '', /stroke="#376092"/, `sparkline group cell E${row}`)
  assert.equal(new Set(Object.values(sheet.cells)).size, sheet.styles.length, 'styles are sent once and shared by index')
  return input
}

/** Preview refreshes reuse the visuals until the workbook or the request changes. */
export function checkVisualCache() {
  const workbook = visualWorkbook()
  const engine = new CalculationEngine(workbook)
  const first = buildPrintVisuals(workbook, { engine, today: 46_000 })
  assert.equal(buildPrintVisuals(workbook, { engine, today: 46_000 }), first, 'an unchanged workbook reuses its visuals')
  assert.notEqual(buildPrintVisuals(workbook, { engine, today: 46_000, showFormulas: true }), first, 'a different request recomputes')
  const sheet = workbook.sheets[0]
  const edited: WorkbookModel = { ...workbook, sheets: [{ ...sheet, cells: { ...sheet.cells, B2: { value: 95 } } }] }
  engine.update(edited)
  const after = buildPrintVisuals(edited, { engine, today: 46_000 })
  assert.notEqual(after, first, 'an edit recomputes')
  const style = (payload: typeof after, address: string) => payload.sheets.report.styles[payload.sheets.report.cells[address]]
  assert.equal(style(after, 'B3').fill, EXPECTED_COLORS.scaleLow, 'the edited values re-scale the colours')
}

/** Page-setup helpers: title ranges, header/footer text, and the model they write. */
export function checkPageSetupHelpers(): number {
  assert.equal(normalizeTitleRows('1'), '1:1')
  assert.equal(normalizeTitleRows(' $3:$2 '), '2:3')
  assert.equal(normalizeTitleRows(''), '')
  assert.equal(normalizeTitleRows('A1'), null)
  assert.equal(normalizeTitleColumns('b'), 'B:B')
  assert.equal(normalizeTitleColumns('$C:$A'), 'A:C')
  assert.equal(normalizeTitleColumns('A1'), null)

  const original = '&L&"Arial,Bold"&14Title&C&P of &N&R&D'
  assert.deepEqual(splitHeaderFooter(original), { left: '&"Arial,Bold"&14Title', center: '&P of &N', right: '&D' })
  const texts = headerFooterTexts(original)
  assert.deepEqual(texts, { left: 'Title', center: '&[Page] of &[Pages]', right: '&[Date]' })
  assert.equal(headerFooterCode(original, texts), original, 'untouched text keeps the file\'s codes exactly')
  assert.equal(headerFooterCode(original, { ...texts, left: 'R&D plan' }), '&L&"Arial,Bold"&14R&&D plan&C&P of &N&R&D', 'edited text keeps its font and a literal & is escaped')
  assert.equal(headerFooterCode(original, { ...texts, left: '2026 plan' }), '&L&"Arial,Bold"&14 2026 plan&C&P of &N&R&D', 'leading digits stay apart from the font size code')
  assert.equal(headerFooterCode('', { left: '', center: 'Page &[Page] of &[Pages] · &[Tab]', right: '' }), '&CPage &P of &N · &A')
  assert.equal(headerFooterCode('Plain', { left: '', center: '', right: '' }), '', 'clearing every section removes the header')

  const sheet: SheetData = { id: 's', name: 'Sheet 1', rowCount: 50, colCount: 8, cells: {}, merges: [], colWidths: {}, rowHeights: {}, frozen: { rows: 1, columns: 1 }, pageSetup: { printArea: 'A1:F100', printAreaWhole: { 'A1:F100': '$A:$F' } } }
  applyPageSetupPatch(sheet, { printArea: '$B$2:$D$10', printTitlesRow: '1', printTitlesColumn: 'a', oddHeader: '&C&A', oddFooter: '&CPage &P', rowBreaks: [21, 11, 21, 1] })
  assert.equal(sheet.pageSetup?.printArea, 'B2:D10')
  assert.equal(sheet.pageSetup?.printAreaWhole, undefined, 'a new print area forgets the whole-column form')
  assert.equal(sheet.pageSetup?.printTitlesRow, '1:1')
  assert.equal(sheet.pageSetup?.printTitlesColumn, 'A:A')
  assert.deepEqual(sheet.headerFooter, { oddHeader: '&C&A', oddFooter: '&CPage &P' })
  assert.deepEqual(sheet.rowBreaks, [{ id: 10, max: 16383, man: 1 }, { id: 20, max: 16383, man: 1 }], 'breaks use ExcelJS\'s shape (row 1 cannot start a page)')
  const state = pageSetupStateFor(sheet, { top: 4, bottom: 9, left: 1, right: 3 })
  assert.equal(state.printArea, 'B2:D10')
  assert.deepEqual(state.rowBreaks, [11, 21])
  assert.equal(state.frozenRows, 1)
  applyPageSetupPatch(sheet, { printArea: null, printTitlesRow: null, printTitlesColumn: 'not a column', oddHeader: '', rowBreaks: [] })
  assert.equal(sheet.pageSetup?.printArea, undefined)
  assert.equal(sheet.pageSetup?.printTitlesRow, undefined)
  assert.equal(sheet.pageSetup?.printTitlesColumn, 'A:A', 'invalid text never replaces a setting')
  assert.deepEqual(sheet.headerFooter, { oddFooter: '&CPage &P' })
  assert.deepEqual(sheet.rowBreaks, [])
  return 1
}
