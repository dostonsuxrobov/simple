import assert from 'node:assert/strict'
import { enablePatches, produceWithPatches } from 'immer'
import { cellLocked, protectionAllows, protectionViolation, sheetProtection } from '../src/lib/protection'
import type { WorkbookModel } from '../src/spreadsheet-types'

enablePatches()
const base = (protection: Record<string, unknown> | null): WorkbookModel => ({
  version: 1, name: 'p.xlsx', activeSheetId: 's',
  sheets: [{
    id: 's', name: 'S', rowCount: 20, colCount: 10, merges: [], colWidths: {}, rowHeights: {},
    cells: { A1: { value: 1 }, B1: { value: 2, style: { protection: { locked: false } } }, C1: { formula: 'A1*2', result: 2 } },
    columnProperties: { '4': { style: { protection: { locked: false } } } },
    sheetProtection: protection,
  }],
})
const check = (book: WorkbookModel, recipe: (draft: WorkbookModel) => void) => {
  const [next, patches] = produceWithPatches(book, recipe)
  return protectionViolation(book, next, patches)
}
const locked = base({ sheet: true, objects: false })
assert.ok(sheetProtection(locked.sheets[0]))
assert.equal(cellLocked(locked.sheets[0], 'A1'), true, 'cells are locked by default')
assert.equal(cellLocked(locked.sheets[0], 'B1'), false)
assert.equal(cellLocked(locked.sheets[0], 'D7'), false, 'a column style can unlock its cells')
assert.ok(check(locked, (d) => { d.sheets[0].cells.A1 = { value: 9 } }), 'a locked cell cannot be edited')
assert.ok(check(locked, (d) => { d.sheets[0].cells.E5 = { value: 'x' } }), 'nor can a new locked cell')
assert.equal(check(locked, (d) => { d.sheets[0].cells.B1.value = 5 }), null, 'unlocked cells stay editable')
assert.equal(check(locked, (d) => { d.sheets[0].cells.D3 = { value: 'ok' } }), null)
assert.ok(check(locked, (d) => { d.sheets[0].cells.B1.style = { font: { bold: true }, protection: { locked: false } } }), 'formatting needs permission')
assert.equal(check(locked, (d) => { d.sheets[0].cells.C1.result = 4; d.sheets[0].cells.C1.display = '4' }), null, 'recalculated results are not edits')
assert.ok(check(locked, (d) => { d.sheets[0].colWidths['2'] = 20 }))
assert.ok(check(locked, (d) => { d.sheets[0].images = [] }), 'objects are protected')
assert.ok(check(locked, (d) => { d.sheets[0].merges = ['A5:B5'] }))
assert.equal(check(locked, (d) => { d.sheets[0].sheetProtection = null }), null, 'unprotecting is always allowed (the command checks the password)')
assert.equal(check(locked, (d) => { d.sheets[0].frozen = { rows: 1, columns: 0 } }), null, 'view settings are not protected')

const permissive = base({ sheet: true, formatCells: true, formatColumns: true, autoFilter: true })
assert.equal(check(permissive, (d) => { d.sheets[0].cells.B1.style = { font: { bold: true }, protection: { locked: false } } }), null)
assert.equal(check(permissive, (d) => { d.sheets[0].colWidths['2'] = 20 }), null)
assert.equal(check(permissive, (d) => { d.sheets[0].hiddenRows = [3]; d.sheets[0].filteredRows = [3] }), null, 'filtering is allowed with AutoFilter permission')
assert.equal(check(permissive, (d) => { d.sheets[0].images = [] }), null, 'objects are editable unless protected')
assert.equal(protectionAllows(sheetProtection(permissive.sheets[0]), 'insertRows'), false)

const open = base(null)
assert.equal(check(open, (d) => { d.sheets[0].cells.A1 = { value: 9 }; d.sheets[0].merges = ['A1:B1'] }), null, 'unprotected sheets are never limited')
assert.equal(check(open, (d) => { d.sheets[0].sheetProtection = { sheet: true } }), null, 'protecting is allowed')
console.log('Protection QA passed: locked defaults, row/column unlocks, content vs formatting, derived results, objects, permissions, and toggling.')
