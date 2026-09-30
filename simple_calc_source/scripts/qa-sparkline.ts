import assert from 'node:assert/strict'
import { CalculationEngine } from '../src/lib/calc-engine'
import '../src/lib/formula-library'
import { isSparklineValue, parseSparkline } from '../src/lib/formula-lib-sparkline'
import { renderSparklineSvg } from '../src/lib/sparkline-render'
import type { WorkbookModel } from '../src/spreadsheet-types'

const book: WorkbookModel = {
  version: 1, name: 'spark.xlsx', activeSheetId: 's',
  sheets: [{ id: 's', name: 'S', rowCount: 10, colCount: 10, merges: [], colWidths: {}, rowHeights: {}, cells: {
    A1: { value: 3 }, B1: { value: -2 }, C1: { value: 5 }, D1: { value: 'x' }, E1: { value: 7 },
    F1: { formula: 'SPARKLINE(A1:E1)' },
    F2: { formula: 'SPARKLINE(A1:E1,{"charttype","column";"color","red";"negcolor","javascript:alert(1)";"highcolor","#00ff00"})' },
    F3: { formula: 'SPARKLINE(A1:E1,{"charttype","winloss"})' },
    F4: { formula: 'SPARKLINE({40,60},{"charttype","bar";"max",100})' },
    F5: { formula: 'SPARKLINE(A1:E1,{"charttype"})' },
  } }],
}
const engine = new CalculationEngine(book)
const line = parseSparkline(engine.getValue('s', 'F1'))
assert.ok(line)
assert.deepEqual(line.data, [3, -2, 5, null, 7], 'text becomes a gap')
const lineSvg = renderSparklineSvg(line, 80, 18)
assert.equal((lineSvg.match(/<polyline/g) || []).length, 1, 'a gap splits the line')
assert.equal((lineSvg.match(/<circle/g) || []).length, 1, 'an isolated point after a gap is drawn as a dot')
const column = parseSparkline(engine.getValue('s', 'F2'))!
assert.equal(column.options.charttype, 'column')
const columnSvg = renderSparklineSvg(column, 80, 18)
assert.equal((columnSvg.match(/<rect/g) || []).length, 4)
assert.match(columnSvg, /fill="#d93025"/)
assert.match(columnSvg, /fill="#00ff00"/, 'the highest point gets highcolor')
assert.doesNotMatch(columnSvg, /javascript/, 'unknown colours never reach the markup')
assert.equal((renderSparklineSvg(parseSparkline(engine.getValue('s', 'F3'))!, 80, 18).match(/<rect/g) || []).length, 4)
const barSvg = renderSparklineSvg(parseSparkline(engine.getValue('s', 'F4'))!, 100, 20)
assert.match(barSvg, /width="40"/)
assert.match(barSvg, /x="40" y="3" width="60"/)
assert.equal(engine.getValue('s', 'F5'), '#VALUE!', 'options need two columns')
assert.ok(isSparklineValue(engine.getValue('s', 'F1')))
const saved = engine.withResults()
assert.equal(saved.sheets[0].cells.F1.result, '', 'the marker is never written into a file')
console.log('Sparkline QA passed: line gaps, column colours, winloss, bar, option validation, and saved results.')
