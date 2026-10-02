// Print/export visuals fixture shared by the QA scripts (Node) and the print dialog QA
// (bundled for a browser window): a sheet with conditional formatting (colour scale, data
// bar, icon set, number format), a banded table, sparklines and checkboxes, and the print
// request App.tsx builds for it, with the visuals resolved as the app resolves them.
import { CalculationEngine } from '../src/lib/calc-engine'
import '../src/lib/formula-library'
import { accountingDisplayParts, formatScalar } from '../src/lib/number-format'
import { buildPrintVisuals } from '../src/lib/print-visuals'
import type { CellData, SheetData, WorkbookModel } from '../src/spreadsheet-types'

/** Expected colours (CSS hex) the printed pages must carry. */
export const EXPECTED_COLORS = {
  scaleLow: '#F8696B',
  scaleHigh: '#63BE7B',
  dataBar: '#638EC6',
  tableHeader: '#4472C4',
  iconRed: '#D6392B',
  iconGreen: '#2F9E44',
  sparklineGroup: '#376092',
  sparklineFormula: '#1A73E8',
  checkbox: '#476B57',
  negativeRed: '#FF0000',
}

function sheetOf(cells: Record<string, CellData>, extra: Partial<SheetData> = {}): SheetData {
  return { id: 'report', name: 'Report', rowCount: 12, colCount: 12, cells, merges: [], colWidths: {}, rowHeights: {}, ...extra }
}

/** The fixture workbook: every visual kind the grid draws beyond plain cell formatting. */
export function visualWorkbook(): WorkbookModel {
  const cells: Record<string, CellData> = {
    A1: { value: 'Region' }, B1: { value: 'Sales' }, C1: { value: 'Growth' }, D1: { value: 'Score' },
    E1: { value: 'Trend' }, F1: { value: 'Mini' }, G1: { value: 'Done' }, H1: { value: 'Flag' }, I1: { value: 'Delta' }, J1: { value: 'Share' },
    I2: { value: -5, numFmt: '0;[Red]-0' },
    J2: { value: 0.25 },
    G2: { value: true, type: 'checkbox' },
    G3: { value: false, type: 'checkbox' },
    H2: { value: true },
    H3: { value: false },
    F2: { formula: 'SPARKLINE(B2:B6)' },
  }
  const rows: Array<[string, number, number, number]> = [
    ['North', 10, 20, 5],
    ['South', 30, 40, 35],
    ['East', 50, 60, 55],
    ['West', 70, 80, 75],
    ['Central', 90, 100, 95],
  ]
  rows.forEach(([region, sales, growth, score], index) => {
    const row = index + 2
    cells[`A${row}`] = { value: region }
    cells[`B${row}`] = { value: sales }
    cells[`C${row}`] = { value: growth }
    cells[`D${row}`] = { value: score }
  })
  const sheet = sheetOf(cells, {
    conditionalFormattings: [
      { ref: 'B2:B6', rules: [{ type: 'colorScale', priority: 1, cfvo: [{ type: 'min' }, { type: 'max' }], color: [{ argb: 'FFF8696B' }, { argb: 'FF63BE7B' }] }] },
      { ref: 'C2:C6', rules: [{ type: 'dataBar', priority: 2, cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } }] },
      { ref: 'D2:D6', rules: [{ type: 'iconSet', priority: 3, iconSet: '3TrafficLights1' }] },
      { ref: 'J2', rules: [{ type: 'expression', priority: 4, formulae: ['J2>0'], style: { numFmt: '0.0%' } }] },
    ],
    tables: [{
      id: 'table-1', name: 'Sales', ref: 'A1:D6', headerRow: true, totalsRow: false,
      columns: [{ name: 'Region' }, { name: 'Sales' }, { name: 'Growth' }, { name: 'Score' }],
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
    }],
    sparklineGroups: [{
      type: 'line', colors: { series: '#376092' },
      sparklines: [2, 3, 4, 5, 6].map((row) => ({ source: `Report!B${row}:D${row}`, cell: `E${row}` })),
    }],
    dataValidations: { 'H2:H3': { type: 'list', formulae: ['"TRUE,FALSE"'], allowBlank: true } },
  })
  return { version: 1, name: 'Visuals.xlsx', activeSheetId: sheet.id, sheets: [sheet], metadata: {} }
}

/**
 * A print request as App.tsx's createPrintPayload builds it, plus `visuals`. Display values
 * use the app's formatting, including the raw SPARKLINE marker the engine must never print.
 */
export function visualPrintInput(options: Record<string, unknown> = {}) {
  const workbook = visualWorkbook()
  const engine = new CalculationEngine(workbook)
  const printable = engine.withResults()
  const displayValues: Record<string, Record<string, string>> = {}
  const displayParts: Record<string, Record<string, { type: 'number' | 'boolean' | 'text'; accounting?: { symbol: string; amount: string } }>> = {}
  for (const sheet of printable.sheets) {
    const values: Record<string, string> = {}
    const parts: Record<string, { type: 'number' | 'boolean' | 'text'; accounting?: { symbol: string; amount: string } }> = {}
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const value = cell.formula ? cell.result : cell.value
      values[address] = formatScalar(value, cell.numFmt || cell.style?.numFmt, cell.display || undefined)
      const accounting = typeof value === 'number' ? accountingDisplayParts(values[address], cell.numFmt || cell.style?.numFmt) : null
      parts[address] = { type: typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'text', ...(accounting ? { accounting } : {}) }
    }
    displayValues[sheet.id] = values
    displayParts[sheet.id] = parts
  }
  const visuals = buildPrintVisuals(printable, { engine, sheetIds: [printable.activeSheetId], today: 46_000 })
  return {
    documentId: 'qa-visuals',
    name: 'Visuals.xlsx',
    workbook: printable,
    displayValues,
    displayParts,
    selection: { top: 0, bottom: 0, left: 0, right: 0 },
    options: { scope: 'active-sheet', orientation: 'landscape', scaling: 'actual', paperSize: 'letter', margins: 'normal', gridlines: true, headings: false, ...options },
    visuals,
  }
}
