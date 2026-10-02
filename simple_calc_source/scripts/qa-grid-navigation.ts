import assert from 'node:assert/strict'
import {
  SHEET_LAST_COL,
  SHEET_LAST_ROW,
  continueTabRun,
  currentRegion,
  edgeJumpIndex,
  enterAfterTabRun,
  firstPositionAfter,
  indexFilledCells,
  lastPositionBefore,
  nextSelectionCorner,
  progressiveSelectAll,
  selectionExtension,
  stepWithinBounds,
  virtualExtent,
} from '../src/lib/grid-navigation'
import type { GridCoord } from '../src/lib/grid-navigation'

let checks = 0
const check = (label: string, run: () => void) => {
  try {
    run()
    checks += 1
  } catch (error) {
    console.error(`FAILED: ${label}`)
    throw error
  }
}
const at = (row: number, col: number): GridCoord => ({ row, col })

// CALC-020: Enter and Tab move the active cell inside a multi-cell selection and wrap.
check('Enter walks down each column of B2:C3, then wraps to the top-left', () => {
  const block = { top: 1, bottom: 2, left: 1, right: 2 }
  const path: GridCoord[] = [at(1, 1)]
  for (let step = 0; step < 4; step += 1) path.push(stepWithinBounds(block, path[path.length - 1], 'columns', false))
  assert.deepEqual(path, [at(1, 1), at(2, 1), at(1, 2), at(2, 2), at(1, 1)])
})
check('Tab walks along each row of B2:C3, then wraps', () => {
  const block = { top: 1, bottom: 2, left: 1, right: 2 }
  const path: GridCoord[] = [at(1, 1)]
  for (let step = 0; step < 4; step += 1) path.push(stepWithinBounds(block, path[path.length - 1], 'rows', false))
  assert.deepEqual(path, [at(1, 1), at(1, 2), at(2, 1), at(2, 2), at(1, 1)])
})
check('Shift+Enter and Shift+Tab walk backwards and wrap to the last cell', () => {
  const block = { top: 1, bottom: 2, left: 1, right: 2 }
  assert.deepEqual(stepWithinBounds(block, at(1, 1), 'columns', true), at(2, 2))
  assert.deepEqual(stepWithinBounds(block, at(1, 1), 'rows', true), at(2, 2))
  assert.deepEqual(stepWithinBounds(block, at(1, 2), 'rows', true), at(1, 1))
})
check('Enter skips hidden rows and cells covered by a merge', () => {
  const block = { top: 0, bottom: 3, left: 0, right: 0 }
  // Row 2 (index 1) hidden: A1 -> A3.
  assert.deepEqual(stepWithinBounds(block, at(0, 0), 'columns', false, (row) => row === 1), at(2, 0))
  // A3:A4 merged (master A3): from A3 the next cell is A1 (A4 is covered).
  assert.deepEqual(stepWithinBounds(block, at(2, 0), 'columns', false, (row) => row === 3), at(0, 0))
})
check('A single cell, or a selection whose cells are all skipped, keeps the active cell', () => {
  assert.deepEqual(stepWithinBounds({ top: 4, bottom: 4, left: 4, right: 4 }, at(4, 4), 'columns', false), at(4, 4))
  assert.deepEqual(stepWithinBounds({ top: 0, bottom: 1, left: 0, right: 0 }, at(0, 0), 'columns', false, () => true), at(0, 0))
})

// CALC-020: a Tab run returns Enter to the column where it started.
check('Tab, Tab, Enter goes to the next row under the first column', () => {
  let run = continueTabRun(null, at(4, 0), at(4, 1))
  run = continueTabRun(run, at(4, 1), at(4, 2))
  assert.equal(run.startCol, 0)
  assert.deepEqual(enterAfterTabRun(run, at(4, 2)), at(5, 0))
})
check('A move between the Tabs ends the run', () => {
  const run = continueTabRun(null, at(4, 0), at(4, 1))
  assert.equal(enterAfterTabRun(run, at(7, 1)), null)
  const restarted = continueTabRun(run, at(9, 3), at(9, 4))
  assert.equal(restarted.startCol, 3)
})

// CALC-021: Ctrl+. cycles the corners clockwise and keeps the range.
check('Ctrl+. cycles B2:D5 clockwise', () => {
  const block = { top: 1, bottom: 4, left: 1, right: 3 }
  const first = nextSelectionCorner(block, at(1, 1))
  assert.deepEqual(first.focus, at(1, 3))
  assert.deepEqual(first.anchor, at(4, 1))
  const second = nextSelectionCorner(block, first.focus)
  assert.deepEqual(second.focus, at(4, 3))
  const third = nextSelectionCorner(block, second.focus)
  assert.deepEqual(third.focus, at(4, 1))
  assert.deepEqual(nextSelectionCorner(block, third.focus).focus, at(1, 1))
})
check('Ctrl+. on a one-column selection alternates its two ends', () => {
  const column = { top: 0, bottom: 5, left: 2, right: 2 }
  assert.deepEqual(nextSelectionCorner(column, at(0, 2)).focus, at(5, 2))
  assert.deepEqual(nextSelectionCorner(column, at(5, 2)).focus, at(0, 2))
})

// Review F3: Shift+Arrow after Shift+Space / Ctrl+Space / Ctrl+A grows the block, not the active cell.
check('Shift+Space on C5 then Shift+Down extends whole rows', () => {
  const row = { anchor: at(4, 2), focus: at(4, 2), range: { top: 4, bottom: 4, left: 0, right: SHEET_LAST_COL } }
  const ends = selectionExtension(row)
  assert.deepEqual(ends.anchor, at(4, 0))
  assert.deepEqual(ends.from, at(4, SHEET_LAST_COL))
  // One step down from the moving corner keeps every column: rows 5:6.
  const next = { row: ends.from.row + 1, col: ends.from.col }
  assert.deepEqual([Math.min(ends.anchor.row, next.row), Math.max(ends.anchor.row, next.row), Math.min(ends.anchor.col, next.col), Math.max(ends.anchor.col, next.col)], [4, 5, 0, SHEET_LAST_COL])
})
check('Ctrl+Space on C5 then Shift+Right extends whole columns', () => {
  const column = { anchor: at(4, 2), focus: at(4, 2), range: { top: 0, bottom: SHEET_LAST_ROW, left: 2, right: 2 } }
  const ends = selectionExtension(column)
  assert.deepEqual(ends.anchor, at(0, 2))
  assert.deepEqual(ends.from, at(SHEET_LAST_ROW, 2))
})
check('Ctrl+A on B2 inside B2:D6 then Shift+Down grows from the far corner', () => {
  const block = { anchor: at(1, 1), focus: at(1, 1), range: { top: 1, bottom: 5, left: 1, right: 3 } }
  const ends = selectionExtension(block)
  assert.deepEqual(ends.anchor, at(1, 1))
  assert.deepEqual(ends.from, at(5, 3))
})
check('Active cell at the bottom-right of a block anchors the top-left', () => {
  const ends = selectionExtension({ anchor: at(5, 3), focus: at(5, 3), range: { top: 1, bottom: 5, left: 1, right: 3 } })
  assert.deepEqual(ends.anchor, at(5, 3))
  assert.deepEqual(ends.from, at(1, 1))
})
check('A plain selection extends its focus and keeps its anchor', () => {
  const ends = selectionExtension({ anchor: at(0, 0), focus: at(3, 2) })
  assert.deepEqual(ends, { anchor: at(0, 0), from: at(3, 2) })
})

// CALC-015: Ctrl+Arrow runs to the data edge, the next value, or the sheet edge.
check('Ctrl+Down in an empty column lands on row 1,048,576', () => {
  assert.equal(edgeJumpIndex({ filled: [], last: SHEET_LAST_ROW }, 0, 1), SHEET_LAST_ROW)
})
check('Ctrl+Down below the last value goes to the sheet edge, not the used range', () => {
  assert.equal(edgeJumpIndex({ filled: [0, 1, 2], last: SHEET_LAST_ROW }, 2, 1), SHEET_LAST_ROW)
})
check('Ctrl+Down inside a block runs to its end, then jumps to the next block', () => {
  const lane = { filled: [0, 1, 2, 10, 11], last: SHEET_LAST_ROW }
  assert.equal(edgeJumpIndex(lane, 0, 1), 2)
  assert.equal(edgeJumpIndex(lane, 2, 1), 10)
  assert.equal(edgeJumpIndex(lane, 10, 1), 11)
  assert.equal(edgeJumpIndex(lane, 5, 1), 10)
})
check('Ctrl+Up runs back to the block start and to row 1', () => {
  const lane = { filled: [0, 1, 2, 10, 11], last: SHEET_LAST_ROW }
  assert.equal(edgeJumpIndex(lane, 11, -1), 10)
  assert.equal(edgeJumpIndex(lane, 10, -1), 2)
  assert.equal(edgeJumpIndex(lane, 2, -1), 0)
  assert.equal(edgeJumpIndex({ filled: [], last: SHEET_LAST_ROW }, 500, -1), 0)
})
check('Ctrl+Right in an empty row lands on column XFD', () => {
  assert.equal(edgeJumpIndex({ filled: [], last: SHEET_LAST_COL }, 3, 1), SHEET_LAST_COL)
})
check('Hidden rows are stepped over and never a stop', () => {
  // Values in rows 1-4 (indices 0-3); row 3 (index 2) hidden: the run continues past it.
  assert.equal(edgeJumpIndex({ filled: [0, 1, 2, 3], hidden: new Set([2]), last: SHEET_LAST_ROW }, 0, 1), 3)
  // A hidden value is not a stop: from row 1, the next visible value is row 9.
  assert.equal(edgeJumpIndex({ filled: [4, 8], hidden: new Set([4]), last: SHEET_LAST_ROW }, 0, 1), 8)
  // The sheet edge itself hidden: stop on the last visible row.
  assert.equal(edgeJumpIndex({ filled: [], hidden: new Set([SHEET_LAST_ROW]), last: SHEET_LAST_ROW }, 0, 1), SHEET_LAST_ROW - 1)
})

// CALC-021: current region (Ctrl+Shift+8) and progressive Ctrl+A.
const table = indexFilledCells([at(1, 1), at(1, 2), at(2, 1), at(2, 2), at(3, 3), at(9, 9)])
check('The current region grows across touching (also diagonal) values', () => {
  assert.deepEqual(currentRegion(table, at(1, 1)), { top: 1, bottom: 3, left: 1, right: 3 })
  assert.deepEqual(currentRegion(table, at(3, 3)), { top: 1, bottom: 3, left: 1, right: 3 })
})
check('A blank cell next to data takes the data block; a blank isolated cell has no region', () => {
  assert.deepEqual(currentRegion(table, at(4, 4)), { top: 1, bottom: 4, left: 1, right: 4 })
  assert.equal(currentRegion(table, at(20, 20)), null)
  assert.deepEqual(currentRegion(table, at(9, 9)), { top: 9, bottom: 9, left: 9, right: 9 })
})
check('Ctrl+A selects the region first, then the sheet', () => {
  const sheet = { top: 0, left: 0, bottom: 99, right: 25 }
  const region = { top: 1, bottom: 3, left: 1, right: 3 }
  assert.deepEqual(progressiveSelectAll({ top: 1, bottom: 1, left: 1, right: 1 }, region, sheet), region)
  assert.deepEqual(progressiveSelectAll(region, region, sheet), sheet)
  assert.deepEqual(progressiveSelectAll({ top: 20, bottom: 20, left: 20, right: 20 }, null, sheet), sheet)
})
check('Region growth stays fast on a large contiguous block', () => {
  const cells: GridCoord[] = []
  for (let row = 0; row < 20_000; row += 1) for (let col = 0; col < 5; col += 1) cells.push(at(row, col))
  const started = performance.now()
  const region = currentRegion(indexFilledCells(cells), at(10_000, 2))
  assert.deepEqual(region, { top: 0, bottom: 19_999, left: 0, right: 4 })
  assert.ok(performance.now() - started < 2_000, 'a 100,000-cell region resolves in well under two seconds')
})

// calc-grid-interaction-12 / CALC-015: the grid's laid-out extent follows the selection and the view.
check('A new sheet lays out 200 rows; a jump to A5000 lays out row 5000 and a margin', () => {
  const base = { minimum: 200, used: 200, usedMargin: 25, selectionEnd: 0, selectionMargin: 25, viewEnd: 30, viewMargin: 30, limit: 1_048_576 }
  assert.equal(virtualExtent(base), 225)
  assert.ok(virtualExtent({ ...base, selectionEnd: 4_999 }) >= 5_000 + 25)
  assert.ok(virtualExtent({ ...base, viewEnd: 900 }) >= 930, 'scrolling near the end lays out another page')
  assert.equal(virtualExtent({ ...base, selectionEnd: SHEET_LAST_ROW }), 1_048_576)
  assert.equal(virtualExtent({ ...base, selectionEnd: SHEET_LAST_ROW, limit: 800_000 }), 800_000)
})

// calc-grid-interaction-11: Find order across sheets ('All sheets').
check('Find Next walks the matches in sheet order and reaches other sheets', () => {
  // Sheet1 matches A1 and A5, Sheet2 matches B2; Find opened at Sheet1!A3.
  const matches = [{ sheet: 0, row: 0, col: 0 }, { sheet: 0, row: 4, col: 0 }, { sheet: 1, row: 1, col: 1 }]
  assert.equal(firstPositionAfter(matches, { sheet: 0, row: 2, col: 0 }), 1)
  assert.equal(firstPositionAfter(matches, { sheet: 0, row: 4, col: 0 }), 2, 'from Sheet1!A5 the next match is on Sheet2')
  assert.equal(firstPositionAfter(matches, { sheet: 1, row: 1, col: 1 }), 0, 'after the last match Find wraps to the first')
  assert.equal(firstPositionAfter(matches, { sheet: 0, row: 4, col: 0 }, true), 1, 'an inclusive origin keeps the current match')
  assert.equal(lastPositionBefore(matches, { sheet: 1, row: 1, col: 1 }), 1)
  assert.equal(lastPositionBefore(matches, { sheet: 0, row: 0, col: 0 }), 2)
  assert.equal(firstPositionAfter([], { sheet: 0, row: 0, col: 0 }), -1)
})

console.log(`Grid navigation QA passed: ${checks} checks (Enter/Tab inside selections, Tab runs, Ctrl+., Ctrl+Arrow to the sheet edge, current region, progressive Ctrl+A, grid extent, Find order).`)
