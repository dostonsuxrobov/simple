import assert from 'node:assert/strict'
import { clearAfterSuccessfulCopy, clipboardSourceIntact, relocateMovedFormula } from '../src/lib/clipboard'
import { moveReferencesInFormula } from '../src/lib/formula-editing'
import { applyPasteSpecial, snapshotValueResolver } from '../src/lib/paste-special'
import { shiftFormulaReferences } from '../src/lib/formulas'
import type { CellData, CellScalar } from '../src/spreadsheet-types'

{
  let clears = 0
  const result = await clearAfterSuccessfulCopy(async () => false, () => { clears += 1 })
  assert.equal(result, false)
  assert.equal(clears, 0, 'a reported clipboard failure must preserve every selected cell')
}

{
  let clears = 0
  const result = await clearAfterSuccessfulCopy(async () => { throw new Error('Clipboard permission denied') }, () => { clears += 1 })
  assert.equal(result, false)
  assert.equal(clears, 0, 'a rejected clipboard write must preserve every selected cell')
}

{
  let clears = 0
  const result = await clearAfterSuccessfulCopy(async () => true, () => { clears += 1 })
  assert.equal(result, true)
  assert.equal(clears, 1, 'a successful clipboard write should clear the selection exactly once')
}

// calc-grid-interaction-1: Paste Values pastes the copied cells' own values, wherever the copy
// started. Sheet: A1='Name', B2=10, C2=5, D2='=B2*C2' (50). Copy D2, paste values at F2.
{
  const sheet: Record<string, CellScalar> = { A1: 'Name', B2: 10, C2: 5, D2: 50 }
  const copied: CellData[][] = [[{ formula: 'B2*C2' }]]
  // The snapshot the app takes at copy time (calculated values, laid out like the copy).
  const values: CellScalar[][] = [[sheet.D2]]
  const result = applyPasteSpecial({
    source: { cells: copied, origin: { row: 1, col: 3 }, sheetName: 'Sheet1' },
    destination: { top: 1, left: 5 },
    getDestinationCell: () => undefined,
    options: { paste: 'values' },
    shiftFormula: shiftFormulaReferences,
    resolveValue: snapshotValueResolver(values),
  })
  assert.deepEqual(result.changes.F2, { value: 50 }, 'F2 receives the value of D2 (50), not the value of A1 ("Name")')

  // The resolver is called with indices inside the copied block, never sheet coordinates.
  const seen: Array<[number, number]> = []
  applyPasteSpecial({
    source: { cells: [[{ formula: 'A1' }], [{ formula: 'A2' }]], origin: { row: 7, col: 3 } },
    destination: { top: 0, left: 0 },
    getDestinationCell: () => undefined,
    options: { paste: 'values' },
    resolveValue: (_cell, row, col) => { seen.push([row, col]); return 1 },
  })
  assert.deepEqual(seen, [[0, 0], [1, 0]], 'resolveValue receives block-relative positions')
}

// Copying D2:D10 and pasting values writes the values of D2:D10, not those of A1:A9.
{
  const copied: CellData[][] = Array.from({ length: 9 }, (_, index) => [{ formula: `B${index + 2}*C${index + 2}`, result: -1 }])
  const values: CellScalar[][] = Array.from({ length: 9 }, (_, index) => [(index + 1) * 100])
  const result = applyPasteSpecial({
    source: { cells: copied, origin: { row: 1, col: 3 } },
    destination: { top: 1, left: 5 },
    getDestinationCell: () => undefined,
    options: { paste: 'valuesAndNumberFormats' },
    resolveValue: snapshotValueResolver(values),
  })
  for (let index = 0; index < 9; index += 1) assert.equal(result.changes[`F${index + 2}`]?.value, (index + 1) * 100)
}

// Transposed values read the snapshot of the source cell they come from.
{
  const copied: CellData[][] = [[{ formula: 'X1' }, { formula: 'X2' }]]
  const values: CellScalar[][] = [[1, 2]]
  const result = applyPasteSpecial({
    source: { cells: copied, origin: { row: 4, col: 4 } },
    destination: { top: 0, left: 0 },
    getDestinationCell: () => undefined,
    options: { paste: 'values', transpose: true },
    resolveValue: snapshotValueResolver(values),
  })
  assert.equal(result.changes.A1?.value, 1)
  assert.equal(result.changes.A2?.value, 2)
}

// A formula without a snapshot value (an external copy) falls back to its cached result.
{
  const result = applyPasteSpecial({
    source: { cells: [[{ formula: 'B2', result: 7 }]], origin: { row: 3, col: 3 } },
    destination: { top: 0, left: 0 },
    getDestinationCell: () => undefined,
    options: { paste: 'values' },
    resolveValue: snapshotValueResolver(undefined),
  })
  assert.equal(result.changes.A1?.value, 7)
}

// calc-grid-interaction-3: a cut only moves cells while they are exactly where they were.
{
  const cells = { A5: { value: 1 } }
  const stamp = { documentId: 'doc-1', sheetId: 'sheet-1', cells }
  assert.equal(clipboardSourceIntact(stamp, { documentId: 'doc-1', sheets: [{ id: 'sheet-1', cells }] }), true)
  assert.equal(clipboardSourceIntact(stamp, { documentId: 'doc-1', sheets: [{ id: 'sheet-1', cells: { ...cells } }] }), false, 'an insert/delete/sort/edit replaces the cell map')
  assert.equal(clipboardSourceIntact(stamp, { documentId: 'doc-2', sheets: [{ id: 'sheet-1', cells }] }), false, 'another workbook with the same sheet ids')
  assert.equal(clipboardSourceIntact(stamp, { documentId: 'doc-1', sheets: [{ id: 'sheet-2', cells: {} }] }), false, 'the source sheet was deleted')
  assert.equal(clipboardSourceIntact(null, { documentId: 'doc-1', sheets: [] }), false)
}

// calc-grid-interaction-4: cut on Sheet1 A1:A3, paste on Sheet2 at C1.
{
  const move = { sourceSheet: 'Sheet1', destinationSheet: 'Sheet2', rect: { top: 0, left: 0, bottom: 2, right: 0 }, rowDelta: 0, colDelta: 2 }
  assert.equal(relocateMovedFormula('A1*2', move), 'C1*2', 'a reference inside the block moves with it')
  assert.equal(relocateMovedFormula('B5+1', move), 'Sheet1!B5+1', 'a reference outside the block keeps pointing at Sheet1')
  assert.equal(relocateMovedFormula('SUM(A1:A2)+$A$3', move), 'SUM(C1:C2)+$C$3', 'ranges and absolute references inside the block move')
  assert.equal(relocateMovedFormula('Sheet1!A2+Sheet1!D9', move), 'Sheet2!C2+Sheet1!D9', 'qualified references follow the same rule')
  assert.equal(relocateMovedFormula('Sheet3!A1+Sheet2!B1', move), 'Sheet3!A1+Sheet2!B1', 'references to other sheets stay as they are')
  assert.equal(relocateMovedFormula('SUM(A:A)+A1:A9', move), 'SUM(Sheet1!A:A)+Sheet1!A1:A9', 'whole columns and ranges partly outside stay on Sheet1')
  assert.equal(relocateMovedFormula('A1#*2', move), 'C1#*2', 'a spill reference inside the block moves')
  assert.equal(relocateMovedFormula('"A1"&TaxRate', move), '"A1"&TaxRate', 'text and names are untouched')
  const quoted = { ...move, sourceSheet: 'My Data' }
  assert.equal(relocateMovedFormula('B5', quoted), "'My Data'!B5", 'sheet names that need quotes are quoted')
  // The rest of the workbook re-points to the block's new place (formulas outside the block).
  assert.equal(moveReferencesInFormula('Sheet1!A1+1', { formulaSheet: 'Sheet2', sourceSheet: 'Sheet1', rect: move.rect, rowDelta: 0, colDelta: 2, destinationSheet: 'Sheet2' }), 'C1+1')
  assert.equal(moveReferencesInFormula('A2*3', { formulaSheet: 'Sheet1', sourceSheet: 'Sheet1', rect: move.rect, rowDelta: 0, colDelta: 2, destinationSheet: 'Sheet2' }), 'Sheet2!C2*3')
}

console.log('Clipboard QA passed: cut clears only after a confirmed write; Paste Values uses the copied cells\' own values; cut mode ends when its cells change; cross-sheet cut keeps formulas pointing at the right cells.')
