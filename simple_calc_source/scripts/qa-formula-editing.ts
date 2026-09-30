import assert from 'node:assert/strict'
import {
  callContextAt, canInsertReferenceAt, completionContextAt, formulaReferences, insertReference,
  toggleAbsoluteReference, tokenizeFormulaText,
} from '../src/lib/formula-editing'
import { fromFileFormula, toFileFormula } from '../src/lib/formula-file-format'

const kinds = (text: string) => tokenizeFormulaText(text).filter((token) => token.kind !== 'space').map((token) => `${token.kind}:${token.text}`)
assert.deepEqual(kinds('=SUM(A1:B2, Sheet2!C3, \'My Sheet\'!$D$4) + LOG10(5) & "x"'), [
  'equals:=', 'function:SUM', 'open:(', 'reference:A1:B2', 'separator:,', 'reference:Sheet2!C3', 'separator:,',
  "reference:'My Sheet'!$D$4", 'close:)', 'operator:+', 'function:LOG10', 'open:(', 'number:5', 'close:)', 'operator:&', 'string:"x"',
])
assert.deepEqual(kinds('=A:A+3:3+A1#'), ['equals:=', 'reference:A:A', 'operator:+', 'reference:3:3', 'operator:+', 'reference:A1#'])
assert.deepEqual(kinds('=SUM(A1:'), ['equals:=', 'function:SUM', 'open:(', 'reference:A1', 'operator::'])
assert.deepEqual(kinds('="unterminated'), ['equals:=', 'string:"unterminated'])
assert.deepEqual(kinds('=Sales[Units]*TaxRate'), ['equals:=', 'name:Sales[Units]', 'operator:*', 'name:TaxRate'])

const references = formulaReferences('=A1+B2*A1+$a$1')
assert.deepEqual(references.map((reference) => reference.colorIndex), [0, 1, 0, 0])
assert.deepEqual(references[1], { ...references[1], top: 1, left: 1, bottom: 1, right: 1 })

assert.deepEqual(callContextAt('=IF(A1>2, SUM(B1, ', 18), { name: 'SUM', argumentIndex: 1, start: 10 })
assert.deepEqual(callContextAt('=IF(A1>2, SUM(B1), ', 19), { name: 'IF', argumentIndex: 2, start: 1 })
assert.deepEqual(callContextAt('=_xlfn.XLOOKUP(A1, {1,2}, ', 26)?.argumentIndex, 2)
assert.equal(callContextAt('=A1+1', 5), null)

assert.deepEqual(completionContextAt('=SU', 3), { prefix: 'SU', start: 1, end: 3 })
assert.deepEqual(completionContextAt('=A1+VLO', 7), { prefix: 'VLO', start: 4, end: 7 })
assert.equal(completionContextAt('="SU', 4), null)
assert.equal(completionContextAt('=SUM(', 5), null)

assert.equal(canInsertReferenceAt('=', 1), true)
assert.equal(canInsertReferenceAt('=SUM(', 5), true)
assert.equal(canInsertReferenceAt('=A1+', 4), true)
assert.equal(canInsertReferenceAt('=A1', 3), false)
assert.equal(canInsertReferenceAt('="a+', 4), false)
assert.equal(canInsertReferenceAt('hello', 5), false)

assert.deepEqual(insertReference('=SUM(', 5, { top: 0, left: 1, bottom: 4, right: 1 }, null), { text: '=SUM(B1:B5', caret: 10, span: { start: 5, end: 10 } })
assert.deepEqual(insertReference('=SUM(B1', 7, { top: 1, left: 1, bottom: 1, right: 1 }, { start: 5, end: 7 }).text, '=SUM(B2')
assert.equal(insertReference('=', 1, { top: 0, left: 0, bottom: 0, right: 0 }, null, 'My Sheet').text, "='My Sheet'!A1")

assert.equal(toggleAbsoluteReference('=A1+B2', 2)?.text, '=$A$1+B2')
assert.equal(toggleAbsoluteReference('=$A$1+B2', 3)?.text, '=A$1+B2')
assert.equal(toggleAbsoluteReference('=A$1', 3)?.text, '=$A1')
assert.equal(toggleAbsoluteReference('=$A1', 3)?.text, '=A1')
assert.equal(toggleAbsoluteReference('=SUM(A1:B2)', 9)?.text, '=SUM($A$1:$B$2)')
assert.equal(toggleAbsoluteReference('=Sheet2!C3', 10)?.text, '=Sheet2!$C$3')
assert.equal(toggleAbsoluteReference('=5+5', 3), null)

// File format prefixes.
assert.equal(toFileFormula('XLOOKUP(A1,B:B,C:C)'), '_xlfn.XLOOKUP(A1,B:B,C:C)')
assert.equal(toFileFormula('FILTER(A1:A9,A1:A9>2)'), '_xlfn._xlws.FILTER(A1:A9,A1:A9>2)')
assert.equal(toFileFormula('LET(x,1,y,x+1,x*y)'), '_xlfn.LET(_xlpm.x,1,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)')
assert.equal(toFileFormula('LAMBDA(a,b,a+b)(1,2)'), '_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)(1,2)')
assert.equal(toFileFormula('SUM(A1#)'), 'SUM(_xlfn.ANCHORARRAY(A1))')
assert.equal(toFileFormula('@A1:A9*2'), '_xlfn.SINGLE(A1:A9)*2')
assert.equal(toFileFormula('SUM(A1:A3)+"FILTER("'), 'SUM(A1:A3)+"FILTER("')
assert.equal(fromFileFormula('_xlfn.XLOOKUP(A1,B:B,C:C)'), 'XLOOKUP(A1,B:B,C:C)')
assert.equal(fromFileFormula('_xlfn._xlws.SORT(A1:A3)'), 'SORT(A1:A3)')
assert.equal(fromFileFormula('_xlfn.LET(_xlpm.x,1,_xlpm.x*2)'), 'LET(x,1,x*2)')
assert.equal(fromFileFormula('SUM(_xlfn.ANCHORARRAY(A1))'), 'SUM(A1#)')
assert.equal(fromFileFormula('_xlfn.SINGLE(A1:A9)*2'), '@A1:A9*2')
for (const formula of ['XLOOKUP(A1,B:B,C:C)', 'LET(x,1,x*2)', 'SUM(A1#)+SORT(B1:B4)', 'IFS(A1>1,"a",TRUE,"b")']) {
  assert.equal(fromFileFormula(toFileFormula(formula)), formula)
}
console.log('Formula editing QA passed: tokens, references, call context, completion, pointing, F4, and file prefixes.')
import { removeSheetFromFormula, renameSheetInFormula } from '../src/lib/formula-editing'
assert.equal(renameSheetInFormula('SUM(Data!A1:A5)+Data!B2', 'Data', 'Sales 2024'), "SUM('Sales 2024'!A1:A5)+'Sales 2024'!B2")
assert.equal(renameSheetInFormula("'My Data'!A1*2", 'my data', 'Totals'), 'Totals!A1*2')
assert.equal(renameSheetInFormula('Other!A1+"Data!A1"', 'Data', 'X'), 'Other!A1+"Data!A1"')
assert.equal(renameSheetInFormula('SUM(A1:A3)', 'Data', 'X'), 'SUM(A1:A3)')
assert.equal(renameSheetInFormula('Data!TaxRate*2', 'Data', 'Q1'), "'Q1'!TaxRate*2")
assert.equal(renameSheetInFormula('A1', 'Data', 'A1'), 'A1')
assert.equal(renameSheetInFormula('Data!A1', 'Data', 'B2'), "'B2'!A1")
assert.equal(removeSheetFromFormula('SUM(Data!A1:A5)+B1', 'Data'), 'SUM(#REF!)+B1')
console.log('Sheet rename QA passed.')
import { moveReferencesInFormula } from '../src/lib/formula-editing'
const move = { formulaSheet: 'Sheet1', sourceSheet: 'Sheet1', rect: { top: 0, left: 0, bottom: 1, right: 1 }, rowDelta: 4, colDelta: 2 }
assert.equal(moveReferencesInFormula('A1+B2*$A$2', move), 'C5+D6*$C$6')
assert.equal(moveReferencesInFormula('SUM(A1:B2)+A3', move), 'SUM(C5:D6)+A3')
assert.equal(moveReferencesInFormula('SUM(A1:B3)', move), 'SUM(A1:B3)')
assert.equal(moveReferencesInFormula('Sheet1!A1', { ...move, formulaSheet: 'Other' }), 'Sheet1!C5')
assert.equal(moveReferencesInFormula('A1*2', { ...move, destinationSheet: 'Data' }), 'Data!C5*2')
assert.equal(moveReferencesInFormula('Sheet1!B1', { ...move, formulaSheet: 'Data', destinationSheet: 'Data' }), 'D5')
console.log('Reference move QA passed.')
