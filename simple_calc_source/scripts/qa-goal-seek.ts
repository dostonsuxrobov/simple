import assert from 'node:assert/strict'
import { CalculationEngine } from '../src/lib/calc-engine'
import { goalSeek } from '../src/lib/goal-seek'
import type { WorkbookModel } from '../src/spreadsheet-types'

const close = (actual: number, expected: number, tolerance = 1e-3) => assert.ok(Math.abs(actual - expected) <= tolerance, `${actual} is not ${expected}`)

{
  const result = goalSeek((x) => 2 * x + 3, 11, 0)
  assert.equal(result.found, true)
  close(result.value, 4)
}
{
  const result = goalSeek((x) => x * x, 2, 1)
  assert.equal(result.found, true)
  close(result.value, Math.SQRT2)
}
{
  const result = goalSeek((x) => x * x, -1, 3)
  assert.equal(result.found, false, 'an unreachable target is reported')
  assert.ok(Math.abs(result.value) < 0.1, 'the closest point is returned')
}
{
  const result = goalSeek((x) => (x > 5 ? null : x * 10), 30, 0)
  assert.equal(result.found, true, 'errors (null) along the way are stepped around')
  close(result.value, 3)
}

// Through the calculation engine: what rate makes the monthly payment 20?
{
  const book: WorkbookModel = {
    version: 1,
    name: 'loan.xlsx',
    activeSheetId: 's1',
    sheets: [{
      id: 's1', name: 'Loan', rowCount: 20, colCount: 5, merges: [], colWidths: {}, rowHeights: {},
      cells: { B1: { value: 1000 }, B2: { value: 0.05 }, B3: { formula: '-PMT(B2/12,60,B1)' }, B4: { formula: 'B3*60' } },
    }],
  }
  const engine = new CalculationEngine(book)
  const before = engine.getValue('s1', 'B3') as number
  close(before, 18.871, 0.01)
  const result = goalSeek((x) => {
    const value = engine.evaluateAt('s1', 'B3', 0, 0, { row: 1, col: 1, value: x })
    return typeof value === 'number' ? value : null
  }, 20, 0.05)
  assert.equal(result.found, true)
  close(result.achieved, 20)
  close(result.value, 0.0741, 0.0005)
  assert.equal(engine.getValue('s1', 'B3'), before, 'the search never changes the workbook')
  const total = goalSeek((x) => engine.evaluateAt('s1', 'B4', 0, 0, { row: 0, col: 1, value: x }) as number, 1500, 1000)
  assert.equal(total.found, true, 'transitive dependents recalculate under the trial value')
  close(total.value, 1000 * 1500 / (before * 60), 0.5)
}

console.log('Goal Seek QA passed: linear, nonlinear, unreachable, error recovery, and engine-backed searches.')
