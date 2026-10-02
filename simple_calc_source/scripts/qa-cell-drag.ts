/*
 * Drag-and-drop of the selection border (CALC-029) and the in-cell list dropdown helpers
 * (CALC-018), as pure functions: moving a block re-points references like cut + paste, a
 * Ctrl-drag copies like copy + paste, merges/arrays/pivots/table headers are protected, and
 * multiple-selection lists accept any set of their options.
 *
 *   npm run test:cell-drag
 */
import assert from 'node:assert/strict'
import { produce } from 'immer'
import type { SheetData, WorkbookModel } from '../src/spreadsheet-types.ts'
import { planCellBlockTransfer, transferCellBlock } from '../src/lib/sheet-operations.ts'
import type { CellBlockTransfer } from '../src/lib/sheet-operations.ts'
import { createSheetHost } from '../src/lib/data-tools-core.ts'
import {
  chipTextColor,
  dropdownOptionColor,
  dropdownPresentation,
  joinMultipleSelection,
  splitMultipleSelection,
  toggleMultipleSelection,
  validateValue,
} from '../src/lib/validation.ts'

function sheet(overrides: Partial<SheetData> = {}): SheetData {
  return { id: 'main', name: 'Main', state: 'visible', rowCount: 20, colCount: 10, cells: {}, merges: [], colWidths: {}, rowHeights: {}, ...overrides }
}

function book(main: SheetData, other?: SheetData): WorkbookModel {
  return {
    version: 1,
    name: 'drag.xlsx',
    activeSheetId: main.id,
    sheets: [main, other || sheet({ id: 'other', name: 'Other' })],
    metadata: {},
  } as WorkbookModel
}

/** Runs the drop on an immer draft, exactly as the app does inside mutateWorkbook. */
function drop(model: WorkbookModel, transfer: Omit<CellBlockTransfer, 'sheetId'> & { sheetId?: string }) {
  let plan: ReturnType<typeof transferCellBlock> | null = null
  const next = produce(model, (draft) => { plan = transferCellBlock(draft as WorkbookModel, { sheetId: 'main', ...transfer }) })
  return { next, plan: plan! as ReturnType<typeof transferCellBlock> }
}

const A1B2 = { top: 0, bottom: 1, left: 0, right: 1 }

// ---- Move: cells, formulas inside and outside the block, references from elsewhere ------------
{
  const model = book(
    sheet({
      cells: {
        A1: { value: 1, style: { font: { bold: true } } },
        B1: { formula: 'A1*2', result: 2 },
        A2: { value: 'note me', note: 'kept' },
        B2: { formula: 'D9+A2', result: 0 },
        C5: { formula: 'SUM(A1:B2)+$A$1+B1', result: 5 },
      },
    }),
    sheet({ id: 'other', name: 'Other', cells: { A1: { formula: 'Main!B1+Main!A1:A2', result: 0 } } }),
  )
  const { next, plan } = drop(model, { source: A1B2, destination: { row: 4, col: 4 }, mode: 'move' })
  assert.equal(plan.error, undefined)
  assert.deepEqual(plan.target, { top: 4, bottom: 5, left: 4, right: 5 })
  const cells = next.sheets[0].cells
  assert.equal(cells.A1, undefined, 'the source is emptied')
  assert.equal(cells.B2, undefined)
  assert.equal(cells.E5.value, 1)
  assert.deepEqual(cells.E5.style, { font: { bold: true } }, 'formatting travels')
  assert.equal(cells.E6.note, 'kept', 'notes travel')
  assert.equal(cells.F5.formula, 'E5*2', 'a reference inside the block follows it')
  assert.equal(cells.F6.formula, 'D9+E6', 'a reference outside the block keeps pointing at the same cell')
  assert.equal(cells.C5.formula, 'SUM(E5:F6)+$E$5+F5', 'formulas elsewhere follow the moved cells, absolute ones too')
  assert.equal(next.sheets[1].cells.A1.formula, 'Main!F5+Main!E5:E6', 'other sheets follow too')
  assert.equal(model.sheets[0].cells.A1.value, 1, 'the input workbook is untouched')
}

// ---- Move onto data: Excel asks first; references to the overwritten cells become #REF! ------
{
  const model = book(sheet({ cells: { A1: { value: 'x' }, C1: { value: 'old' }, D1: { formula: 'C1&"!"', result: 'old!' }, D2: { formula: 'SUM(C1:C3)', result: 0 } } }))
  const plan = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 0, bottom: 0, left: 0, right: 0 }, destination: { row: 0, col: 2 }, mode: 'move' })
  assert.equal(plan.overwrites, true, 'C1 holds data')
  const { next } = drop(model, { source: { top: 0, bottom: 0, left: 0, right: 0 }, destination: { row: 0, col: 2 }, mode: 'move' })
  assert.equal(next.sheets[0].cells.C1.value, 'x')
  assert.equal(next.sheets[0].cells.D1.formula, '#REF!&"!"', 'a reference to the replaced cell is #REF!, as in Excel')
  assert.equal(next.sheets[0].cells.D2.formula, 'SUM(C1:C3)', 'a range only partly replaced is kept')
  const empty = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 0, bottom: 0, left: 0, right: 0 }, destination: { row: 5, col: 5 }, mode: 'move' })
  assert.equal(empty.overwrites, false)
}

// ---- Overlapping move (one row down): the block's own cells are not "data in the way" ---------
{
  const model = book(sheet({ cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: { formula: 'A1+A2+A3', result: 6 } } }))
  const source = { top: 0, bottom: 2, left: 0, right: 0 }
  const plan = planCellBlockTransfer(model, { sheetId: 'main', source, destination: { row: 1, col: 0 }, mode: 'move' })
  assert.equal(plan.overwrites, false)
  const { next } = drop(model, { source, destination: { row: 1, col: 0 }, mode: 'move' })
  const cells = next.sheets[0].cells
  assert.deepEqual([cells.A1?.value, cells.A2?.value, cells.A3?.value, cells.A4?.value], [undefined, 1, 2, 3])
  assert.equal(cells.B1.formula, 'A2+A3+A4')
}

// ---- Dropped where it started: nothing happens ------------------------------------------------
{
  const model = book(sheet({ cells: { A1: { value: 1 } } }))
  const { next, plan } = drop(model, { source: A1B2, destination: { row: 0, col: 0 }, mode: 'move' })
  assert.equal(plan.unchanged, true)
  assert.equal(next, model, 'no change, so no undo step')
}

// ---- Copy (Ctrl): relative references shift, absolute ones stay; the source stays ------------
{
  const model = book(sheet({ cells: { A1: { value: 2 }, B1: { formula: 'A1*$A$1', result: 4, display: '4' }, C9: { formula: 'B1', result: 4 } } }))
  const { next } = drop(model, { source: { top: 0, bottom: 0, left: 0, right: 1 }, destination: { row: 3, col: 2 }, mode: 'copy' })
  const cells = next.sheets[0].cells
  assert.equal(cells.A1.value, 2, 'the source stays')
  assert.equal(cells.B1.formula, 'A1*$A$1')
  assert.equal(cells.C4.value, 2)
  assert.equal(cells.D4.formula, 'C4*$A$1', 'relative references shift like copy + paste')
  assert.equal(cells.D4.result, undefined, 'the copy is recalculated, not shown with the old result')
  assert.equal(cells.C9.formula, 'B1', 'references to the source are not re-pointed by a copy')
}

// ---- Merges travel; a block that splits a merge is refused -----------------------------------
{
  const model = book(sheet({ cells: { A1: { value: 'm' } }, merges: ['A1:B1', 'H1:I2'] }))
  const { next } = drop(model, { source: A1B2, destination: { row: 10, col: 0 }, mode: 'move' })
  assert.deepEqual([...next.sheets[0].merges].sort(), ['A11:B11', 'H1:I2'])
  const copied = drop(model, { source: A1B2, destination: { row: 10, col: 0 }, mode: 'copy' }).next
  assert.deepEqual([...copied.sheets[0].merges].sort(), ['A11:B11', 'A1:B1', 'H1:I2'].sort())
  const split = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 0, bottom: 0, left: 0, right: 0 }, destination: { row: 5, col: 5 }, mode: 'move' })
  assert.match(split.error || '', /merged cell/)
  const onto = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 5, bottom: 5, left: 0, right: 0 }, destination: { row: 0, col: 8 }, mode: 'copy' })
  assert.match(onto.error || '', /merged cell/, 'landing on part of a merge is refused too')
}

// ---- Arrays, pivots and table headers are protected -------------------------------------------
{
  const model = book(sheet({
    cells: { A1: { formula: 'TRANSPOSE(D1:D2)', formulaType: 'array', formulaRange: 'A1:B1' }, B1: { value: 3 } },
    pivots: [{ id: 'p', name: 'Pivot', source: 'Main!A1:B5', anchor: { row: 10, col: 5 }, rows: [], columns: [], values: [], filters: [], extent: { rows: 3, cols: 2 } }],
    tables: [{ id: 't', name: 'Sales', ref: 'H5:I8', headerRow: true, totalsRow: false, columns: [{ name: 'A' }, { name: 'B' }] }],
  }))
  const part = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 0, bottom: 0, left: 0, right: 0 }, destination: { row: 3, col: 0 }, mode: 'move' })
  assert.match(part.error || '', /part of an array/)
  const whole = drop(model, { source: { top: 0, bottom: 0, left: 0, right: 1 }, destination: { row: 3, col: 0 }, mode: 'move' }).next
  assert.equal(whole.sheets[0].cells.A4.formulaRange, 'A4:B4', 'a whole array range moves with its range')
  const pivot = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 0, bottom: 0, left: 3, right: 3 }, destination: { row: 11, col: 5 }, mode: 'copy' })
  assert.match(pivot.error || '', /Pivot/)
  const header = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 4, bottom: 4, left: 7, right: 7 }, destination: { row: 15, col: 0 }, mode: 'move' })
  assert.match(header.error || '', /header row/)
  const body = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 5, bottom: 6, left: 7, right: 7 }, destination: { row: 15, col: 0 }, mode: 'move' })
  assert.equal(body.error, undefined, 'cells below the headers can move')
  const edge = planCellBlockTransfer(model, { sheetId: 'main', source: { top: 0, bottom: 1, left: 0, right: 0 }, destination: { row: 1_048_575, col: 0 }, mode: 'move' })
  assert.match(edge.error || '', /outside the sheet/)
}

// ---- Validation travels with a move and is copied with a copy ---------------------------------
{
  const rule = { type: 'list', allowBlank: true, formulae: ['"Yes,No"'] }
  const model = book(sheet({ cells: { A1: { value: 'Yes' } }, dataValidations: { 'A1:A5': rule, 'F1:F2': { type: 'whole', formulae: [1] } } }))
  const moved = drop(model, { source: { top: 0, bottom: 1, left: 0, right: 0 }, destination: { row: 0, col: 5 }, mode: 'move' }).next
  assert.deepEqual(Object.keys(moved.sheets[0].dataValidations || {}).sort(), ['A3:A5', 'F1:F2'], 'the moved rows take the list; the target loses its old rule')
  assert.deepEqual(moved.sheets[0].dataValidations?.['F1:F2'], rule)
  const copied = drop(model, { source: { top: 0, bottom: 1, left: 0, right: 0 }, destination: { row: 0, col: 5 }, mode: 'copy' }).next
  assert.deepEqual(Object.keys(copied.sheets[0].dataValidations || {}).sort(), ['A1:A5', 'F1:F2'])
  assert.deepEqual(copied.sheets[0].dataValidations?.['F1:F2'], rule)
}

// ---- Shared formulas: a moved cell leaves its group ------------------------------------------
{
  const model = book(sheet({ cells: {
    A1: { formula: 'B1*2', formulaType: 'shared', result: 0 },
    A2: { formula: 'B2*2', formulaType: 'shared', sharedFormulaMaster: 'A1', result: 0 } as SheetData['cells'][string],
  } }))
  const { next } = drop(model, { source: { top: 0, bottom: 0, left: 0, right: 0 }, destination: { row: 0, col: 4 }, mode: 'move' })
  const cells = next.sheets[0].cells as Record<string, SheetData['cells'][string] & { sharedFormulaMaster?: string }>
  assert.equal(cells.E1.formulaType, undefined)
  assert.equal(cells.E1.formula, 'B1*2')
  assert.equal(cells.A2.sharedFormulaMaster, undefined, 'a clone whose master moved keeps its own formula')
  assert.equal(cells.A2.formula, 'B2*2')
}

// ---- Dropdown presentation and multiple selections (CALC-018) ---------------------------------
{
  const plain = dropdownPresentation({ type: 'list', formulae: ['"a,b"'] })
  assert.deepEqual([plain.style, plain.multiple, plain.colors.size], ['arrow', false, 0], "a rule without Simple's settings behaves as in Excel")
  const rule = { type: 'list', formulae: ['"Open,In progress,Done"'], simpleDropdown: { style: 'chip', multiple: true, colors: { done: '#B7E1CD', Open: 'red', 'In progress': '#fce8b2' } } }
  const chips = dropdownPresentation(rule)
  assert.equal(chips.style, 'chip')
  assert.equal(chips.multiple, true)
  assert.equal(dropdownOptionColor(chips, 'Done'), '#b7e1cd', 'colours match options case-insensitively')
  assert.equal(dropdownOptionColor(chips, 'Open'), undefined, 'only hex colours are used')
  assert.equal(dropdownPresentation(rule), chips, 'cached per rule')
  assert.equal(chipTextColor('#b7e1cd'), '#222421')
  assert.equal(chipTextColor('#1a3d6b'), '#ffffff')

  assert.deepEqual(splitMultipleSelection(' Done, open ,, done'), ['Done', 'open'])
  assert.equal(joinMultipleSelection(['Done', 'Open', 'legacy'], ['Open', 'In progress', 'Done']), 'Open, Done, legacy', 'list order, unknown values kept last')
  assert.equal(toggleMultipleSelection('Open', 'Done', ['Open', 'In progress', 'Done']), 'Open, Done')
  assert.equal(toggleMultipleSelection('Open, Done', 'open', ['Open', 'In progress', 'Done']), 'Done')

  const host = createSheetHost({ cells: {} })
  assert.equal(validateValue(rule, 'Open, Done', 'Open, Done', host).ok, true, 'a multiple-selection rule accepts a set of its options')
  assert.equal(validateValue(rule, 'Open, Lost', 'Open, Lost', host).ok, false, 'but not a value outside the list')
  assert.equal(validateValue({ ...rule, simpleDropdown: { style: 'chip' } }, 'Open, Done', 'Open, Done', host).ok, false, 'a single-choice list rejects a set')
  assert.equal(validateValue(rule, 'Done', 'Done', host).ok, true)
}

process.stdout.write('Cell drag QA passed: dragging the selection border moves cells like cut + paste (references follow, overwritten references become #REF!), Ctrl-drag copies like copy + paste, merges/arrays/pivots/table headers are protected, validation travels, and list dropdowns support chips, colours and multiple selections.\n')
