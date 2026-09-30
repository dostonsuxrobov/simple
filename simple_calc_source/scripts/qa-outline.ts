import assert from 'node:assert/strict'
import { changeOutline, clearOutline, collapseGroup, expandGroup, innermostGroupAt, maxOutlineLevel, outlineGroups, showOutlineLevel, summaryAfter } from '../src/lib/outline'
import type { SheetData } from '../src/spreadsheet-types'

const sheet = (extra: Partial<SheetData> = {}): SheetData => ({ id: 's', name: 'S', rowCount: 50, colCount: 10, cells: {}, merges: [], colWidths: {}, rowHeights: {}, ...extra })

{
  const data = sheet()
  // Detail rows 2-9 (0-based 1..8) with a nested group 3-5, summaries below.
  assert.equal(changeOutline(data, 'row', 1, 8, 1), true)
  changeOutline(data, 'row', 2, 4, 1)
  assert.equal(maxOutlineLevel(data, 'row'), 2)
  assert.equal(summaryAfter(data, 'row'), true)
  let groups = outlineGroups(data, 'row')
  assert.deepEqual(groups.map(({ level, start, end, summary, collapsed }) => [level, start, end, summary, collapsed]), [[1, 1, 8, 9, false], [2, 2, 4, 5, false]])

  collapseGroup(data, 'row', groups[1])
  assert.deepEqual(data.hiddenRows, [3, 4, 5])
  assert.equal(data.rowProperties?.['6']?.collapsed, true, 'the summary row carries the collapsed flag')
  groups = outlineGroups(data, 'row')
  assert.equal(groups[1].collapsed, true)
  collapseGroup(data, 'row', groups[0])
  assert.deepEqual(data.hiddenRows, [2, 3, 4, 5, 6, 7, 8, 9])
  expandGroup(data, 'row', outlineGroups(data, 'row')[0])
  assert.deepEqual(data.hiddenRows, [3, 4, 5], 'a nested collapsed group stays collapsed')
  expandGroup(data, 'row', outlineGroups(data, 'row')[1])
  assert.deepEqual(data.hiddenRows, [])

  showOutlineLevel(data, 'row', 1)
  assert.deepEqual(data.hiddenRows, [2, 3, 4, 5, 6, 7, 8, 9], 'level 1 shows only summaries')
  showOutlineLevel(data, 'row', 2)
  assert.deepEqual(data.hiddenRows, [3, 4, 5])
  showOutlineLevel(data, 'row', 3)
  assert.deepEqual(data.hiddenRows, [])
  assert.equal(innermostGroupAt(outlineGroups(data, 'row'), 3)?.level, 2)

  data.filteredRows = [8]
  data.hiddenRows = [8]
  showOutlineLevel(data, 'row', 1)
  showOutlineLevel(data, 'row', 3)
  assert.deepEqual(data.hiddenRows, [8], 'filter-hidden rows stay hidden when detail is shown')

  changeOutline(data, 'row', 2, 4, -1)
  assert.equal(maxOutlineLevel(data, 'row'), 1)
  clearOutline(data, 'row')
  assert.equal(maxOutlineLevel(data, 'row'), 0)
  assert.deepEqual(outlineGroups(data, 'row'), [])
}
{
  // Summary columns to the left, and adjacent groups at the same level stay separate.
  const data = sheet({ properties: { outlineProperties: { summaryRight: false } } })
  changeOutline(data, 'column', 2, 3, 1)
  changeOutline(data, 'column', 5, 6, 1)
  assert.equal(summaryAfter(data, 'column'), false)
  assert.deepEqual(outlineGroups(data, 'column').map(({ start, end, summary }) => [start, end, summary]), [[2, 3, 1], [5, 6, 4]])
  collapseGroup(data, 'column', outlineGroups(data, 'column')[1])
  assert.deepEqual(data.hiddenCols, [6, 7])
  assert.equal(data.columnProperties?.['5']?.collapsed, true)
  for (let index = 0; index < 9; index += 1) changeOutline(data, 'column', 0, 0, 1)
  assert.equal(data.columnProperties?.['1']?.outlineLevel, 7, 'levels stop at 7')
}

console.log('Outline QA passed: grouping, nesting, collapse/expand, level buttons, filters, summary position, and limits.')
