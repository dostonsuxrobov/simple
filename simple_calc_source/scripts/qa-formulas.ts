import assert from 'node:assert/strict'
import { evaluateFormula, shiftFormulaReferences } from '../src/lib/formulas.ts'

const cells: Record<string, number | string | boolean | null> = {
  A1: 1, A2: 2, A3: 3,
  B1: 'East', B2: 'West', B3: 'East',
  C1: 10, C2: 20, C3: 30,
  E1: 1, F1: 2, G1: 3,
  E2: 'one', F2: 'two', G2: 'three',
}
const resolver = (_sheet: string, address: string) => cells[address] ?? null
const formula = (source: string) => evaluateFormula(source, 'Sheet1', resolver)

assert.equal(formula('SUMIF(B1:B3,"East",C1:C3)'), 40)
assert.equal(formula('SUMIFS(C1:C3,B1:B3,"East",A1:A3,">1")'), 30)
assert.equal(formula('COUNTIF(C1:C3,">=20")'), 2)
assert.equal(formula('COUNTIFS(B1:B3,"East",C1:C3,">10")'), 1)
assert.equal(formula('AVERAGEIF(B1:B3,"East",C1:C3)'), 20)
assert.equal(formula('AVERAGEIFS(C1:C3,B1:B3,"East",A1:A3,">1")'), 30)
assert.equal(formula('PRODUCT(A1:A3)'), 6)
assert.equal(formula('MEDIAN(C1:C3)'), 20)
assert.equal(formula('LARGE(C1:C3,2)'), 20)
assert.equal(formula('SMALL(C1:C3,2)'), 20)
assert.equal(formula('RANK(20,C1:C3)'), 2)
assert.equal(formula('VAR.S(A1:A3)'), 1)
assert.equal(formula('STDEV.S(A1:A3)'), 1)
assert.equal(formula('ROUNDUP(1.231,2)'), 1.24)
assert.equal(formula('ROUNDDOWN(1.239,2)'), 1.23)
assert.equal(formula('INT(-1.2)'), -2)
assert.equal(formula('MOD(-3,2)'), 1)
assert.equal(formula('POWER(3,2)'), 9)
assert.equal(formula('SQRT(81)'), 9)
assert.equal(formula('IFERROR(1/0,"safe")'), 'safe')
assert.equal(formula('IFS(A1=2,"no",A1=1,"yes")'), 'yes')
assert.equal(formula('XOR(TRUE,FALSE,FALSE)'), true)
assert.equal(formula('LEN("a😀")'), 2)
assert.equal(formula('LEFT("simple_calc",6)'), 'simple')
assert.equal(formula('RIGHT("simple_calc",4)'), 'calc')
assert.equal(formula('MID("simple_calc",8,4)'), 'calc')
assert.equal(formula('TRIM("  simple   calc  ")'), 'simple calc')
assert.equal(formula('PROPER("simple calc")'), 'Simple Calc')
assert.equal(formula('TEXTJOIN("-",TRUE,"a","",B1)'), 'a-East')
assert.equal(formula('SUBSTITUTE("a-b-b","b","x",2)'), 'a-b-x')
assert.equal(formula('FIND("Calc","Simple Calc")'), 8)
assert.equal(formula('SEARCH("c*","Simple Calc")'), 8)
assert.equal(formula('EXACT("Calc","calc")'), false)
assert.equal(formula('VALUE("$1,250.50")'), 1250.5)
assert.equal(formula('YEAR(DATE(2024,2,29))'), 2024)
assert.equal(formula('MONTH(EDATE(DATE(2024,1,31),1))'), 2)
assert.equal(formula('DAY(EOMONTH(DATE(2024,2,5),0))'), 29)
assert.equal(formula('INDEX(C1:C3,2)'), 20)
assert.equal(formula('MATCH("West",B1:B3,0)'), 2)
assert.equal(formula('VLOOKUP("West",B1:C3,2,FALSE)'), 20)
assert.equal(formula('HLOOKUP(2,E1:G2,2,FALSE)'), 'two')
assert.equal(shiftFormulaReferences('A1+$B1+C$1+$D$1', 2, 3), 'D3+$B3+F$1+$D$1')

process.stdout.write('Formula QA passed: 42 Google Sheets-compatible formula and reference checks.\n')
