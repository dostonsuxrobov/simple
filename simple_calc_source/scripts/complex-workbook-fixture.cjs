'use strict'

// Synthetic source model for creation through Simple Calc's real workbook API.
// No external or user documents are read by this fixture.
function complexWorkbook() {
  const money = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
  const thin = { style: 'thin', color: { argb: 'FF80968C' } }
  const border = { top: thin, right: thin, bottom: thin, left: thin }
  const base = { font: { name: 'Arial', size: 10 }, alignment: { vertical: 'middle' } }
  const heading = { font: { name: 'Georgia', size: 21, bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF234E42' } }, alignment: { vertical: 'middle' } }
  const tableHead = { font: { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF40715C' } }, alignment: { vertical: 'middle', wrapText: true }, border }
  const sheet = (id, name, rows, cols) => ({ id, name, rowCount: rows, colCount: cols, state: 'visible', cells: {}, merges: [], colWidths: {}, rowHeights: {}, hiddenRows: [], hiddenCols: [], properties: { defaultRowHeight: 18, defaultColWidth: 12 }, frozen: { rows: 3, columns: 1 }, pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, showGridLines: false, margins: { left: .4, right: .4, top: .6, bottom: .6, header: .2, footer: .2 } }, headerFooter: { oddHeader: '&LComplex workbook test&R&A', oddFooter: '&LConfidential sample&RPage &P of &N' } })
  const summary = sheet('summary', 'Summary', 22, 6)
  summary.merges = ['A1:F2', 'A13:F15', 'A18:C19', 'D18:F19']
  summary.colWidths = { 1: 28, 2: 18, 3: 12, 4: 16, 5: 12, 6: 12 }
  summary.rowHeights = { 1: 23, 2: 23, 3: 8, 13: 22, 14: 22, 15: 22, 17: 10, 18: 18, 19: 18 }
  summary.pageSetup.printArea = 'A1:F20'
  summary.cells.A1 = { value: 'Quarterly operations', style: heading }
  const metrics = [
    ['Transactions', 'COUNTA(Transactions!A4:A203)'],
    ['Net total', 'SUM(Transactions!F4:F203)'],
    ['Tax total', 'SUM(Transactions!G4:G203)'],
    ['Gross total', 'SUM(Transactions!H4:H203)'],
    ['North net', 'SUMIF(Transactions!C4:C203,"North",Transactions!F4:F203)'],
    ['Reconciliation', 'ROUND(B5+B6-B7,2)'],
    ['Blank result', 'IF(B4>0,"",0)'],
    ['Handled input error', 'IFERROR(\'Inputs & notes\'!B12,0)'],
  ]
  metrics.forEach(([label, formula], i) => { summary.cells[`A${i + 4}`] = { value: label, style: base }; summary.cells[`B${i + 4}`] = { formula, style: { ...base, font: { ...base.font, bold: true } }, numFmt: i > 0 && i < 5 ? money : '0.00' } })
  summary.cells.A13 = { value: 'This synthetic workbook checks linked calculations, readable imported geometry, and export fidelity. The paragraph must remain completely visible across three merged rows. Final sentence: the complete paragraph is retained.', style: { ...base, alignment: { wrapText: true, vertical: 'middle' }, border } }
  summary.cells.A17 = { value: 'Prepared by', style: base }
  summary.cells.D17 = { value: 'Reviewed by', style: base }
  summary.cells.A18 = { value: '', style: { ...base, border } }
  summary.cells.D18 = { value: '', style: { ...base, border } }
  const transactions = sheet('transactions', 'Transactions', 203, 8)
  transactions.merges = ['A1:H2']
  transactions.colWidths = { 1: 15, 2: 14, 3: 13, 4: 10, 5: 13, 6: 15, 7: 15, 8: 15 }
  transactions.rowHeights = { 1: 23, 2: 23, 3: 30 }
  transactions.pageSetup = { ...transactions.pageSetup, orientation: 'landscape', printArea: 'A1:H63', printTitlesRow: '3:3' }
  transactions.cells.A1 = { value: 'Transaction detail', style: heading }
  ;['Invoice', 'Date', 'Region', 'Quantity', 'Unit price', 'Net', 'Tax', 'Gross'].forEach((value, index) => { transactions.cells[`${String.fromCharCode(65 + index)}3`] = { value, style: tableHead } })
  let net = 0, tax = 0, north = 0
  for (let r = 4; r <= 203; r++) {
    const quantity = (r % 11) + 1, price = 10 + (r % 7) * 1.25, line = quantity * price, charge = Math.round(line * .075 * 100) / 100
    const region = ['North', 'South', 'East', 'West'][r % 4]
    net += line; tax += charge; if (region === 'North') north += line
    const style = { ...base, border: { bottom: thin }, ...(r % 2 === 0 ? { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF1ED' } } } : {}) }
    transactions.cells[`A${r}`] = { value: `INV-${String(r - 3).padStart(5, '0')}`, style }
    transactions.cells[`B${r}`] = { value: 46023 + r - 4, numFmt: 'yyyy-mm-dd', style }
    transactions.cells[`C${r}`] = { value: region, style }
    transactions.cells[`D${r}`] = { value: quantity, style }
    transactions.cells[`E${r}`] = { value: price, numFmt: money, style }
    transactions.cells[`F${r}`] = { formula: `D${r}*E${r}`, numFmt: money, style }
    transactions.cells[`G${r}`] = { formula: `ROUND(F${r}*'Inputs & notes'!$B$3,2)`, numFmt: money, style }
    transactions.cells[`H${r}`] = { formula: `F${r}+G${r}`, numFmt: money, style }
  }
  const notes = sheet('notes', 'Inputs & notes', 26, 6)
  notes.merges = ['A1:F2', 'B5:F5', 'B6:F6', 'B7:F7', 'A16:F18']
  notes.colWidths = { 1: 22, 2: 19, 3: 13, 4: 13, 5: 13, 6: 13 }
  notes.rowHeights = { 1: 23, 2: 23, 5: 27, 6: 27, 7: 27, 9: 12, 10: 30, 14: 30, 16: 24, 17: 24, 18: 24 }
  notes.pageSetup = { ...notes.pageSetup, paperSize: 1, printArea: 'A1:F20' }
  notes.cells = {
    A1: { value: 'Inputs and layout checks', style: heading }, A3: { value: 'Tax rate', style: base }, B3: { value: .075, numFmt: '0.00%', style: base },
    A5: { value: 'Japanese', style: base }, B5: { value: '請求書と支払い — 日本語の確認', style: { ...base, font: { name: 'Yu Gothic', size: 12 } } },
    A6: { value: 'Arabic', style: base }, B6: { value: 'فاتورة تجريبية باللغة العربية', style: { ...base, font: { name: 'Arial', size: 12 }, alignment: { readingOrder: 'rtl', horizontal: 'right', vertical: 'middle' } } },
    A7: { value: 'Accents', style: base }, B7: { value: 'Crème brûlée · é · €1.234,56 · 🧾', style: { ...base, font: { name: 'Segoe UI', size: 12 } } },
    A9: { value: 'Tight row', style: base }, B9: { value: 12345.67, numFmt: '#,##0.00', style: { ...base, font: { name: 'Cambria', size: 10 }, alignment: { vertical: 'bottom' } } },
    A10: { value: 'Top alignment', style: base }, B10: { value: 89012.34, numFmt: '#,##0.00', style: { ...base, font: { name: 'Arial', size: 11 }, alignment: { vertical: 'top' } } },
    A11: { value: 'Identifier as text', style: base }, B11: { value: '00001234567890123456', numFmt: '@', style: base },
    A12: { value: 'Native error cell', style: base }, B12: { value: '#DIV/0!', type: 'error', style: base },
    A13: { value: 'Boolean', style: base }, B13: { value: false, style: base },
    A14: { value: 'CSV quoting', style: base }, B14: { value: 'Comma, quote " and\nnew line', style: { ...base, alignment: { wrapText: true } } },
    A16: { value: 'Long merged text with an explicit line break.\nThe final line must remain inside the merged range when printing and exporting. Complete final line.', style: { ...base, alignment: { wrapText: true, vertical: 'top' }, border } },
  }
  const audit = sheet('audit', 'Audit hidden', 2, 2)
  audit.state = 'hidden'; audit.cells.A1 = { value: 'HIDDEN_AUDIT_SENTINEL' }; audit.cells.B1 = { formula: 'Summary!B9' }
  return { workbook: { version: 1, name: 'Complex operations.xlsx', activeSheetId: 'summary', sheets: [summary, transactions, notes, audit], metadata: { creator: 'Simple QA', normalFont: { name: 'Arial', size: 10 } } }, expected: { net, tax: Math.round(tax * 100) / 100, north, count: 200 }, money }
}
module.exports = { complexWorkbook }
