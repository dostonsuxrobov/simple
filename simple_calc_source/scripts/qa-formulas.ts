import assert from 'node:assert/strict'
import { describeFormulaError, diagnoseFormula, evaluateFormula, findLookupIndex, shiftFormulaReferences } from '../src/lib/formulas.ts'
import type { FormulaEvaluationHooks } from '../src/lib/formulas.ts'

const cells: Record<string, number | string | boolean | null> = {
  A1: 1, A2: 2, A3: 3,
  B1: 'East', B2: 'West', B3: 'East',
  C1: 10, C2: 20, C3: 30,
  E1: 1, F1: 2, G1: 3,
  E2: 'one', F2: 'two', G2: 'three',
  F5: '=A1+41',
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
assert.equal(formula('SWITCH(A2,1,"one",2,"two")'), 'two')
assert.equal(formula('SWITCH(B1,"West","w","East","e")'), 'e')
assert.equal(formula('SWITCH(A3,1,"one",2,"two","other")'), 'other')
assert.equal(formula('SWITCH(A3,1,"one",2,"two")'), '#N/A')
assert.equal(formula('SUBTOTAL(9,C1:C3)'), 60)
assert.equal(formula('SUBTOTAL(109,C1:C3)'), 60)
assert.equal(formula('SUBTOTAL(1,C1:C3)'), 20)
assert.equal(formula('SUBTOTAL(2,C1:C3)'), 3)
assert.equal(formula('SUBTOTAL(4,C1:C3)'), 30)
assert.equal(formula('SUBTOTAL(5,C1:C3)'), 10)
assert.equal(formula('SUBTOTAL(6,A1:A3)'), 6)
assert.equal(formula('SUBTOTAL(12,C1:C3)'), '#VALUE!')
assert.equal(formula('_xlfn.IFS(A1=2,"no",A1=1,"yes")'), 'yes')
assert.deepEqual(formula('_xlfn._xlws.SORT(A1:A3)'), formula('SORT(A1:A3)'))
assert.deepEqual(formula('_xlws.SORT(A1:A3)'), formula('SORT(A1:A3)'))
// Omitted arguments keep their positional meaning as the empty value, like Excel.
assert.equal(formula('IF(A1>50,1,)'), 0)
assert.equal(formula('IF(A1>0,,1)'), 0)
assert.equal(formula('SUM(A1,)'), 1)
assert.equal(formula('SUM(,A1)'), 1)
assert.equal(formula('MATCH("West",B1:B3,)'), 2)
assert.equal(formula('IFERROR(IF(A1>50,1,),9)'), 0)
// Operators broadcast over a range so boolean-mask idioms work.
assert.equal(formula('SUMPRODUCT((B1:B3="East")*(C1:C3))'), 40)
assert.equal(formula('SUMPRODUCT(--(B1:B3="East"),C1:C3)'), 40)
assert.equal(formula('SUMPRODUCT(A1:A3*C1:C3)'), 140)
assert.equal(formula('SUM((B1:B3="East")*1)'), 2)
assert.equal(formula('SUMPRODUCT((B1:B3="East")*(C1:C3>10)*(C1:C3))'), 30)
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

const close = (actual: unknown, expected: number, tolerance = 1e-9) => {
  assert.equal(typeof actual, 'number')
  assert.ok(Math.abs((actual as number) - expected) <= tolerance, `${String(actual)} !~ ${expected}`)
}
const round2 = (value: unknown) => Math.round((value as number) * 100) / 100

const definedNames: Record<string, string | number> = {
  SALES: 'C1:C3',
  TAXRATE: 0.1,
  EASTLABEL: 'East',
  GREETING: 'hello world',
  WHOLECOL: 'A:A',
}
const hooks: FormulaEvaluationHooks = {
  getUsedRange: () => ({ maxRow: 3, maxCol: 7 }),
  resolveDefinedName: (name) => definedNames[name.toUpperCase()],
  currentCell: { row: 5, column: 2 },
}
const hooked = (source: string) => evaluateFormula(source, 'Sheet1', resolver, hooks)

// Whole-column/row references clamped to the used range.
assert.equal(hooked('SUM(A:A)'), 6)
assert.equal(hooked('SUM($A:$A)'), 6)
assert.equal(hooked('SUM(A:C)'), 66)
assert.equal(hooked('SUM(1:1)'), 17)
assert.equal(hooked('SUM(1:2)'), 39)
assert.equal(hooked('SUM(2:2)'), 22)
assert.equal(hooked('COUNT(A:A)'), 3)
assert.equal(hooked('COUNTA(B:B)'), 3)
assert.equal(hooked('MIN(C:C)'), 10)
assert.equal(hooked('MAX(C:C)'), 30)
assert.equal(hooked('AVERAGE(A:A)'), 2)
assert.equal(hooked('SUM(Sheet1!A:A)'), 6)
assert.equal(hooked("SUM('Sheet1'!C:C)"), 60)
assert.equal(hooked('SUMIF(B:B,"East",C:C)'), 40)
assert.equal(hooked('A:A'), 1)
assert.equal(formula('SUM(A:A)'), 6)
assert.equal(formula('SUM(1:1)'), 17)
assert.equal(formula('SUM(1:1048577)'), '#REF!')
assert.equal(formula('SUM(XFE:XFE)'), '#REF!')

// Sparse iteration over mostly-empty whole columns.
const sparseCells: Record<string, number | string> = { A1: 1, A2: 2, A3: 3, A500000: 7, B2: 9 }
const sparseResolver = (_sheet: string, address: string) => sparseCells[address] ?? null
const sparseHooks: FormulaEvaluationHooks = {
  getUsedRange: () => ({ maxRow: 500_000, maxCol: 4 }),
  forEachCellInRange: (_sheet, bounds, visit) => {
    for (const key of Object.keys(sparseCells)) {
      const match = /^([A-Z])(\d+)$/.exec(key)
      if (!match) continue
      const column = match[1].charCodeAt(0) - 64
      const row = Number(match[2])
      if (
        row >= bounds.startRow && row <= bounds.endRow &&
        column >= bounds.startColumn && column <= bounds.endColumn
      ) {
        visit(row, column)
      }
    }
  },
}
const sparse = (source: string) => evaluateFormula(source, 'Sheet1', sparseResolver, sparseHooks)

assert.equal(sparse('SUM(A:A)'), 13)
assert.equal(sparse('COUNT(A:A)'), 4)
assert.equal(sparse('MAX(A:A)'), 7)
assert.equal(sparse('AVERAGE(A:A)'), 3.25)
assert.equal(sparse('COUNTA(A:B)'), 5)
assert.equal(sparse('COUNTBLANK(A:A)'), 499_996)
assert.equal(sparse('SUM(A1:B500000)'), 22)
assert.equal(sparse('SUM(A1:A3)'), 6)
assert.equal(sparse('VLOOKUP(1,A:B,2)'), '#VALUE!')
const cappedHooks: FormulaEvaluationHooks = { getUsedRange: () => ({ maxRow: 500_000, maxCol: 4 }) }
assert.equal(evaluateFormula('SUM(A:A)', 'Sheet1', sparseResolver, cappedHooks), '#VALUE!')
const denseHooks: FormulaEvaluationHooks = {
  getUsedRange: () => ({ maxRow: 1_000_000, maxCol: 1 }),
  forEachCellInRange: (_sheet, bounds, visit) => {
    const last = Math.min(bounds.endRow, bounds.startRow + 150_000)
    for (let row = bounds.startRow; row <= last; row += 1) visit(row, bounds.startColumn)
  },
}
assert.equal(evaluateFormula('SUM(A:A)', 'Sheet1', () => 1, denseHooks), 150_001)

// Defined names.
assert.equal(hooked('SUM(SALES)'), 60)
assert.equal(hooked('TAXRATE'), 0.1)
assert.equal(hooked('TAXRATE*C1'), 1)
assert.equal(hooked('GREETING'), 'hello world')
assert.equal(hooked('COUNTIF(B1:B3,EASTLABEL)'), 2)
assert.equal(hooked('SUM(WHOLECOL)'), 6)
assert.equal(hooked('BOGUS'), '#NAME?')
assert.equal(formula('SALES'), '#NAME?')

// Array literals.
assert.equal(formula('SUM({1,2;3,4})'), 10)
assert.equal(formula('{1,2;3,4}'), 1)
assert.equal(formula('INDEX({1,2;3,4},2,2)'), 4)
assert.equal(formula('SUM({1,"2",TRUE})'), 1)
assert.equal(formula('SUM({-1,-2})'), -3)
assert.equal(formula('MEDIAN({3,1,2})'), 2)
assert.equal(formula('MATCH("b",{"a","b","c"},0)'), 2)
assert.equal(formula('SUM({1,#DIV/0!})'), '#DIV/0!')
assert.equal(formula('{1,2;3}'), '#VALUE!')
// Unreadable formulas evaluate to an Excel error (never the non-standard #PARSE!).
assert.equal(formula('{}'), '#NAME?')
assert.equal(formula('SUM(1;2;3)'), 6)
assert.equal(formula('C1:C3'), 10)

// shiftFormulaReferences with whole-column/row ranges.
assert.equal(shiftFormulaReferences('SUM(A:A)', 5, 2), 'SUM(C:C)')
assert.equal(shiftFormulaReferences('SUM($A:B)', 0, 1), 'SUM($A:C)')
assert.equal(shiftFormulaReferences('SUM(1:2)', 3, 0), 'SUM(4:5)')
assert.equal(shiftFormulaReferences('SUM($1:$2)', 3, 0), 'SUM($1:$2)')
assert.equal(shiftFormulaReferences('Sheet2!A:A', 0, 1), 'Sheet2!B:B')
assert.equal(shiftFormulaReferences("'My Sheet'!A:B", 0, 1), "'My Sheet'!B:C")
assert.equal(shiftFormulaReferences('SUM(A:B)', 0, -1), 'SUM(#REF!:A)')
assert.equal(shiftFormulaReferences('"A:B"&A1', 1, 1), '"A:B"&B2')
assert.equal(shiftFormulaReferences('sum(a:b)', 0, 1), 'sum(b:c)')

// TEXT via SSF.
assert.equal(formula('TEXT(0.285,"0.0%")'), '28.5%')
assert.equal(formula('TEXT(1234.567,"$#,##0.00")'), '$1,234.57')
assert.equal(formula('TEXT(DATE(2024,3,9),"yyyy-mm-dd")'), '2024-03-09')
assert.equal(formula('TEXT("abc","0.0")'), 'abc')
assert.equal(formula('TEXT("0.5","0%")'), '50%')

// SUMPRODUCT.
assert.equal(formula('SUMPRODUCT(A1:A3,C1:C3)'), 140)
assert.equal(formula('SUMPRODUCT(A1:A3)'), 6)
assert.equal(formula('SUMPRODUCT(A1:A3,C1:C2)'), '#VALUE!')

// XLOOKUP.
assert.equal(formula('XLOOKUP("West",B1:B3,C1:C3)'), 20)
assert.equal(formula('XLOOKUP("Z",B1:B3,C1:C3)'), '#N/A')
assert.equal(formula('XLOOKUP("Z",B1:B3,C1:C3,"missing")'), 'missing')
assert.equal(formula('XLOOKUP(25,C1:C3,A1:A3,"nf",-1)'), 2)
assert.equal(formula('XLOOKUP(25,C1:C3,A1:A3,"nf",1)'), 3)
assert.equal(formula('XLOOKUP("East",B1:B3,C1:C3)'), 10)
assert.equal(formula('XLOOKUP("East",B1:B3,C1:C3,"nf",0,-1)'), 30)
assert.equal(formula('XLOOKUP("W*",B1:B3,C1:C3,"nf",2)'), 20)
assert.equal(formula('INDEX(XLOOKUP("West",B1:B3,B1:C3),1,2)'), 20)

// CHOOSE.
assert.equal(formula('CHOOSE(2,"a","b","c")'), 'b')
assert.equal(formula('CHOOSE(1,10,1/0)'), 10)
assert.equal(formula('CHOOSE(5,"a")'), '#VALUE!')

// OFFSET.
assert.equal(formula('OFFSET(A1,1,2)'), 20)
assert.equal(formula('SUM(OFFSET(A1,0,0,3,1))'), 6)
assert.equal(formula('SUM(OFFSET(A1:A2,1,2,2,1))'), 50)
assert.equal(formula('OFFSET(A1,-1,0)'), '#REF!')
assert.equal(formula('OFFSET(A1,0,0,0,1)'), '#REF!')

// INDIRECT.
assert.equal(formula('INDIRECT("C2")'), 20)
assert.equal(formula('SUM(INDIRECT("A1:A3"))'), 6)
assert.equal(formula('INDIRECT("Sheet1!C3")'), 30)
assert.equal(formula('SUM(INDIRECT("A:A"))'), 6)
assert.equal(formula('INDIRECT("junk")'), '#REF!')

// ROW/COLUMN/ROWS/COLUMNS.
assert.equal(formula('ROW(B7)'), 7)
assert.equal(formula('COLUMN(B7)'), 2)
assert.equal(formula('ROW(A3:A5)'), 3)
assert.equal(hooked('ROW()'), 5)
assert.equal(hooked('COLUMN()'), 2)
assert.equal(formula('ROW()'), '#VALUE!')
assert.equal(formula('ROWS(A1:B3)'), 3)
assert.equal(formula('COLUMNS(A1:B3)'), 2)
assert.equal(formula('ROWS({1;2;3})'), 3)
assert.equal(formula('ROWS(A:A)'), 1048576)
assert.equal(formula('COLUMNS(A:C)'), 3)

// IS* family.
assert.equal(formula('ISBLANK(D1)'), true)
assert.equal(formula('ISBLANK(A1)'), false)
assert.equal(formula('ISNUMBER(A1)'), true)
assert.equal(formula('ISNUMBER(B1)'), false)
assert.equal(formula('ISNUMBER(1/0)'), false)
assert.equal(formula('ISTEXT(B1)'), true)
assert.equal(formula('ISTEXT(A1)'), false)
assert.equal(formula('ISLOGICAL(TRUE)'), true)
assert.equal(formula('ISLOGICAL(A1)'), false)
assert.equal(formula('ISERROR(1/0)'), true)
assert.equal(formula('ISERROR(A1)'), false)
assert.equal(formula('ISERR(1/0)'), true)
assert.equal(formula('ISERR(NA())'), false)
assert.equal(formula('ISNA(NA())'), true)
assert.equal(formula('ISNA(1/0)'), false)
assert.equal(formula('ISFORMULA(F5)'), true)
assert.equal(formula('ISFORMULA(A1)'), false)
// Hosts whose resolvers return computed values report formulas via the hook.
const formulaCellHooked = (source: string) => evaluateFormula(source, 'Sheet1', resolver, {
  isFormulaCell: (_sheet, address) => address === 'F5',
})
assert.equal(formulaCellHooked('ISFORMULA(F5)'), true)
assert.equal(formulaCellHooked('ISFORMULA(A1)'), false)
assert.equal(formula('F5'), 42)

// IFNA / NA.
assert.equal(formula('IFNA(NA(),"x")'), 'x')
assert.equal(formula('IFNA(1/0,"x")'), '#DIV/0!')
assert.equal(formula('IFNA(5,"x")'), 5)
assert.equal(formula('NA()'), '#N/A')

// COUNTBLANK.
assert.equal(formula('COUNTBLANK(A1:D3)'), 3)
assert.equal(formula('COUNTBLANK(A1:A3)'), 0)

// MINIFS / MAXIFS.
assert.equal(formula('MINIFS(C1:C3,B1:B3,"East")'), 10)
assert.equal(formula('MAXIFS(C1:C3,B1:B3,"East")'), 30)
assert.equal(formula('MAXIFS(C1:C3,B1:B3,"East",A1:A3,">1")'), 30)
assert.equal(formula('MINIFS(C1:C3,B1:B3,"Z")'), 0)

// CEILING / FLOOR family.
assert.equal(formula('CEILING(2.5,1)'), 3)
assert.equal(formula('CEILING(-2.5,-2)'), -4)
assert.equal(formula('CEILING(-2.5,2)'), -2)
close(formula('CEILING(1.5,0.1)'), 1.5, 1e-9)
assert.equal(formula('CEILING(2,-1)'), '#NUM!')
assert.equal(formula('CEILING(5,0)'), 0)
assert.equal(formula('CEILING.MATH(24.3,5)'), 25)
assert.equal(formula('CEILING.MATH(6.7)'), 7)
assert.equal(formula('CEILING.MATH(-5.5,2)'), -4)
assert.equal(formula('CEILING.MATH(-5.5,2,1)'), -6)
assert.equal(formula('FLOOR(2.5,1)'), 2)
assert.equal(formula('FLOOR(-2.5,-2)'), -2)
assert.equal(formula('FLOOR(-2.5,2)'), -4)
assert.equal(formula('FLOOR(2.5,0)'), '#DIV/0!')
assert.equal(formula('FLOOR.MATH(24.3,5)'), 20)
assert.equal(formula('FLOOR.MATH(6.7)'), 6)
assert.equal(formula('FLOOR.MATH(-5.5,2)'), -6)
assert.equal(formula('FLOOR.MATH(-5.5,2,1)'), -4)
assert.equal(formula('MROUND(10,3)'), 9)
close(formula('MROUND(1.3,0.2)'), 1.4, 1e-9)
assert.equal(formula('MROUND(-10,-3)'), -9)
assert.equal(formula('MROUND(5,-2)'), '#NUM!')
assert.equal(formula('TRUNC(8.9)'), 8)
assert.equal(formula('TRUNC(-8.9)'), -8)
assert.equal(formula('TRUNC(3.14159,2)'), 3.14)
assert.equal(formula('EVEN(1.5)'), 2)
assert.equal(formula('EVEN(3)'), 4)
assert.equal(formula('EVEN(2)'), 2)
assert.equal(formula('EVEN(-1)'), -2)
assert.equal(formula('EVEN(0)'), 0)
assert.equal(formula('ODD(1.5)'), 3)
assert.equal(formula('ODD(3)'), 3)
assert.equal(formula('ODD(2)'), 3)
assert.equal(formula('ODD(-2)'), -3)
assert.equal(formula('ODD(0)'), 1)

// Date and time parsing.
assert.equal(formula('DATEVALUE("2024-02-29")'), formula('DATE(2024,2,29)'))
assert.equal(formula('DATEVALUE("3/9/2024")'), formula('DATE(2024,3,9)'))
assert.equal(formula('DATEVALUE("9 Mar 2024")'), formula('DATE(2024,3,9)'))
assert.equal(formula('DATEVALUE("March 9, 2024")'), formula('DATE(2024,3,9)'))
assert.equal(formula('DATEVALUE("2024-01-01 12:30")'), formula('DATE(2024,1,1)'))
assert.equal(formula('DATEVALUE("2024-02-30")'), '#VALUE!')
assert.equal(formula('DATEVALUE("hello")'), '#VALUE!')
assert.equal(formula('TIMEVALUE("6:00")'), 0.25)
assert.equal(formula('TIMEVALUE("12:00 PM")'), 0.5)
close(formula('TIMEVALUE("12:30:00 AM")'), 30 / 1440, 1e-12)
assert.equal(formula('TIMEVALUE("2024-01-01 06:00")'), 0.25)
assert.equal(formula('TIMEVALUE("25:00")'), '#VALUE!')
assert.equal(formula('TIME(6,0,0)'), 0.25)
close(formula('TIME(18,30,0)'), 0.7708333333333334, 1e-12)
close(formula('TIME(25,0,0)'), 1 / 24, 1e-12)
assert.equal(formula('TIME(-1,0,0)'), '#NUM!')
assert.equal(formula('HOUR(TIME(18,30,45))'), 18)
assert.equal(formula('MINUTE(TIME(18,30,45))'), 30)
assert.equal(formula('SECOND(TIME(18,30,45))'), 45)
assert.equal(formula('HOUR(0.75)'), 18)
assert.equal(formula('WEEKDAY(DATE(2024,1,1))'), 2)
assert.equal(formula('WEEKDAY(DATE(2024,1,1),2)'), 1)
assert.equal(formula('WEEKDAY(DATE(2024,1,1),3)'), 0)
assert.equal(formula('WEEKDAY(DATE(2024,1,7))'), 1)
assert.equal(formula('WEEKDAY(DATE(2024,1,1),14)'), 5)
assert.equal(formula('WEEKDAY(DATE(2024,1,1),9)'), '#NUM!')
assert.equal(formula('WEEKNUM(DATE(2024,1,1))'), 1)
assert.equal(formula('WEEKNUM(DATE(2024,1,7))'), 2)
assert.equal(formula('WEEKNUM(DATE(2024,1,7),2)'), 1)
assert.equal(formula('WEEKNUM(DATE(2023,1,1),21)'), 52)
assert.equal(formula('WEEKNUM(DATE(2024,1,4),21)'), 1)
assert.equal(formula('DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Y")'), 2)
assert.equal(formula('DATEDIF(DATE(2001,1,1),DATE(2001,3,15),"M")'), 2)
assert.equal(formula('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"D")'), 440)
assert.equal(formula('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"MD")'), 14)
assert.equal(formula('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YM")'), 2)
assert.equal(formula('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YD")'), 75)
assert.equal(formula('DATEDIF(DATE(2002,1,1),DATE(2001,1,1),"D")'), '#NUM!')
assert.equal(formula('DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Q")'), '#NUM!')
assert.equal(formula('NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))'), 110)
assert.equal(formula('NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),DATE(2012,11,22))'), 109)
assert.equal(formula('NETWORKDAYS(DATE(2013,3,1),DATE(2012,10,1))'), -110)
assert.equal(formula('WORKDAY(DATE(2008,10,1),151)'), formula('DATE(2009,4,30)'))
assert.equal(formula('WORKDAY(DATE(2024,1,8),-5)'), formula('DATE(2024,1,1)'))
assert.equal(formula('WORKDAY(DATE(2024,1,5),1,DATE(2024,1,8))'), formula('DATE(2024,1,9)'))
assert.equal(formula('DAYS(DATE(2021,3,15),DATE(2021,2,1))'), 42)
assert.equal(formula('DAYS(367,1)'), 366)

// Financial functions.
assert.equal(round2(formula('PMT(0.08/12,10,10000)')), -1037.03)
assert.equal(formula('PMT(0,12,1200)'), -100)
assert.equal(round2(formula('FV(0.06/12,10,-200,-500,1)')), 2581.4)
assert.equal(formula('FV(0,10,-100)'), 1000)
assert.equal(round2(formula('PV(0.08/12,240,500)')), -59777.15)
close(formula('NPER(0.12/12,-100,-1000,10000,1)'), 59.6738657, 1e-6)
assert.equal(round2(formula('IPMT(0.1/12,1,36,8000)')), -66.67)
assert.equal(round2(formula('PPMT(0.1/12,1,24,2000)')), -75.62)
close(
  (formula('IPMT(0.1/12,2,36,8000)') as number) + (formula('PPMT(0.1/12,2,36,8000)') as number),
  formula('PMT(0.1/12,36,8000)') as number,
  1e-9,
)
assert.equal(formula('IPMT(0.1/12,0,36,8000)'), '#NUM!')
assert.equal(round2(formula('NPV(0.1,-10000,3000,4200,6800)')), 1188.44)
const irr = formula('IRR({-70000,12000,15000,18000,21000})')
close(irr, -0.0212, 1e-3)
const irrNpv = [-70000, 12000, 15000, 18000, 21000].reduce(
  (sum, value, index) => sum + value / (1 + (irr as number)) ** index,
  0,
)
assert.ok(Math.abs(irrNpv) < 1e-6)
close(formula('IRR({-100,60,60})'), 0.13066, 1e-4)
assert.equal(formula('IRR({100,200})'), '#NUM!')
const rate = formula('RATE(48,-200,8000)')
close(rate, 0.0077, 1e-4)
close(formula(`PV(${rate as number},48,-200)`), 8000, 0.01)

// Text functions.
assert.equal(formula('REPT("ab",3)'), 'ababab')
assert.equal(formula('REPT("x",0)'), '')
assert.equal(formula('REPT("x",-1)'), '#VALUE!')
assert.equal(formula('CHAR(65)'), 'A')
assert.equal(formula('CODE("A")'), 65)
assert.equal(formula('CHAR(CODE("a"))'), 'a')
assert.equal(formula('CHAR(0)'), '#VALUE!')
assert.equal(formula('CODE("")'), '#VALUE!')
assert.equal(formula('CLEAN("a"&CHAR(10)&"b")'), 'ab')
assert.equal(formula('REPLACE("abcdef",2,3,"XY")'), 'aXYef')
assert.equal(formula('REPLACE("abc",1,0,"Z")'), 'Zabc')
assert.equal(formula('FIXED(1234.567,1)'), '1,234.6')
assert.equal(formula('FIXED(1234.567,-2)'), '1,200')
assert.equal(formula('FIXED(1234.567,1,TRUE)'), '1234.6')
assert.equal(formula('FIXED(0.5)'), '0.50')
assert.equal(formula('FIXED(-1234.567)'), '-1,234.57')
assert.equal(formula('DOLLAR(1234.567)'), '$1,234.57')
assert.equal(formula('DOLLAR(-1234.567,-2)'), '($1,200)')
assert.equal(formula('DOLLAR(-0.123,4)'), '($0.1230)')
assert.equal(formula('TEXTBEFORE("red-blue-green","-")'), 'red')
assert.equal(formula('TEXTBEFORE("red-blue-green","-",2)'), 'red-blue')
assert.equal(formula('TEXTBEFORE("red-blue-green","-",-1)'), 'red-blue')
assert.equal(formula('TEXTAFTER("red-blue-green","-")'), 'blue-green')
assert.equal(formula('TEXTAFTER("red-blue-green","-",-1)'), 'green')
assert.equal(formula('TEXTAFTER("abc","x")'), '#N/A')
assert.equal(formula('TEXTBEFORE("abc","")'), '#VALUE!')
assert.equal(formula('TEXTAFTER("abc","b",0)'), '#VALUE!')

// RAND / RANDBETWEEN.
const randValue = formula('RAND()')
assert.equal(typeof randValue, 'number')
assert.ok((randValue as number) >= 0 && (randValue as number) < 1)
assert.equal(formula('RANDBETWEEN(5,5)'), 5)
const randBetween = formula('RANDBETWEEN(1,10)')
assert.ok(Number.isInteger(randBetween) && (randBetween as number) >= 1 && (randBetween as number) <= 10)
assert.equal(formula('RANDBETWEEN(10,1)'), '#NUM!')

// Math functions.
assert.equal(formula('EXP(0)'), 1)
close(formula('EXP(1)'), Math.E, 1e-12)
close(formula('LN(EXP(2))'), 2, 1e-12)
assert.equal(formula('LN(0)'), '#NUM!')
assert.equal(formula('LOG(8,2)'), 3)
close(formula('LOG(100)'), 2, 1e-12)
assert.equal(formula('LOG(8,1)'), '#DIV/0!')
assert.equal(formula('LOG10(1000)'), 3)
assert.equal(formula('LOG10(0)'), '#NUM!')
close(formula('PI()'), Math.PI, 1e-15)
assert.equal(formula('SIGN(-5)'), -1)
assert.equal(formula('SIGN(0)'), 0)
assert.equal(formula('SIGN(3)'), 1)
assert.equal(formula('SIN(0)'), 0)
close(formula('SIN(PI()/2)'), 1, 1e-12)
close(formula('COS(0)'), 1, 1e-12)
close(formula('TAN(PI()/4)'), 1, 1e-12)
close(formula('ASIN(1)'), Math.PI / 2, 1e-12)
close(formula('ACOS(1)'), 0, 1e-12)
close(formula('ATAN(1)'), Math.PI / 4, 1e-12)
close(formula('ATAN2(1,1)'), Math.PI / 4, 1e-12)
assert.equal(formula('ASIN(2)'), '#NUM!')
assert.equal(formula('ATAN2(0,0)'), '#DIV/0!')
close(formula('DEGREES(PI())'), 180, 1e-12)
close(formula('RADIANS(180)'), Math.PI, 1e-12)
close(formula('SQRTPI(1)'), Math.sqrt(Math.PI), 1e-12)
assert.equal(formula('SQRTPI(-1)'), '#NUM!')

// Statistics.
assert.equal(formula('PERCENTILE(C1:C3,0.5)'), 20)
assert.equal(formula('PERCENTILE(C1:C3,0.25)'), 15)
assert.equal(formula('PERCENTILE.INC(C1:C3,0.25)'), 15)
assert.equal(formula('PERCENTILE(C1:C3,2)'), '#NUM!')
assert.equal(formula('QUARTILE(C1:C3,2)'), 20)
assert.equal(formula('QUARTILE(C1:C3,0)'), 10)
assert.equal(formula('QUARTILE.INC(C1:C3,4)'), 30)
assert.equal(formula('QUARTILE(C1:C3,5)'), '#NUM!')
assert.equal(formula('MODE({1,2,2,3})'), 2)
assert.equal(formula('MODE.SNGL({5,5,6,6})'), 5)
assert.equal(formula('MODE({1,2,3})'), '#N/A')
assert.equal(formula('VARP({2,4})'), 1)
assert.equal(formula('VAR.P({2,4})'), 1)
assert.equal(formula('STDEVP({2,4})'), 1)
close(formula('STDEV.P(A1:A3)'), Math.sqrt(2 / 3), 1e-12)
assert.equal(formula('VAR.P({5})'), 0)
assert.equal(formula('AVERAGEA(A1:A3)'), 2)
assert.equal(formula('AVERAGEA(B1:C1)'), 5)
assert.equal(formula('AVERAGEA({1,TRUE})'), 1)

// Dynamic arrays (top-left display; no spill).
assert.equal(formula('ROWS(UNIQUE(B1:B3))'), 2)
assert.equal(formula('INDEX(UNIQUE(B1:B3),2,1)'), 'West')
assert.equal(formula('UNIQUE(B1:B3)'), 'East')
assert.equal(formula('ROWS(UNIQUE(B1:B3,FALSE,TRUE))'), 1)
assert.equal(formula('INDEX(UNIQUE(B1:B3,FALSE,TRUE),1,1)'), 'West')
assert.equal(formula('INDEX(SORT(C1:C3,1,-1),1,1)'), 30)
assert.equal(formula('SORT(C1:C3,1,-1)'), 30)
assert.equal(formula('INDEX(SORT(B1:C3,2,-1),1,2)'), 30)
assert.equal(formula('SORT(C1:C3,1,2)'), '#VALUE!')
assert.equal(formula('SUM(FILTER(C1:C3,{TRUE;FALSE;TRUE}))'), 40)
assert.equal(formula('FILTER(C1:C3,{FALSE;FALSE;FALSE},"none")'), 'none')
assert.equal(formula('FILTER(C1:C3,{FALSE;FALSE;FALSE})'), '#CALC!')
assert.equal(formula('FILTER(C1:C3,{TRUE;FALSE})'), '#VALUE!')
assert.equal(formula('SUM(SEQUENCE(4))'), 10)
assert.equal(formula('SEQUENCE(2,2,10,5)'), 10)
assert.equal(formula('INDEX(SEQUENCE(2,2,10,5),2,2)'), 25)
assert.equal(formula('ROWS(SEQUENCE(3,2))'), 3)
assert.equal(formula('SEQUENCE(0)'), '#VALUE!')
assert.equal(formula('INDEX(TRANSPOSE(B1:C3),1,2)'), 'West')
assert.equal(formula('ROWS(TRANSPOSE(A1:B3))'), 2)
assert.equal(formula('COLUMNS(TRANSPOSE(A1:B3))'), 3)

// ---- Regression cases for the verified engine bugs (calc-formula-engine-1 … -12, CALC-007) ----
{
  const sheet: Record<string, unknown> = {
    // Dates 2023-03-15, 2024-01-01, 2024-04-19, text, blank, 2024-12-31, 2025-01-01.
    C1: 45000, C2: 45292, C3: 45400, C4: 'x', C6: 45657, C7: 45658,
    B1: 1, B2: 2, B3: 3, B4: '10',
    E1: 'apple', E2: 'banana', E3: 'cherry', F1: 1, F2: 2, F3: 3,
    H1: 2, H2: 'Hello', H3: true, H4: '20240115',
    J1: '5', J2: true, J3: '10', J4: 4, K1: 1, K2: '#DIV/0!',
    M1: 'a*c', M2: 'abc', N1: 10, N2: 20,
  }
  const run = (source: string) => evaluateFormula(source, 'Sheet1', (_s, address) => sheet[address] as never, { resolverReturnsValues: true })

  // calc-formula-engine-1: criteria read dates, %, currency and thousands like typed input;
  // text criteria with < or > only compare text.
  assert.equal(run('COUNTIFS(C1:C7,">=1/1/2024",C1:C7,"<=12/31/2024")'), 3)
  assert.equal(run('SUMIFS(C1:C7,C1:C7,">=1/1/2024")'), 182007)
  assert.equal(run('COUNTIF(C1:C7,"1/1/2024")'), 1)
  assert.equal(run('COUNTIF(C1:C7,"Jan 1, 2024")'), 1)
  assert.equal(run('COUNTIF(C1:C7,">=2024-01-01")'), 4)
  assert.equal(run('COUNTIF(B1:B5,">50%")'), 3)
  assert.equal(run('COUNTIF(B1:B5,">$1")'), 2)
  assert.equal(run('COUNTIF(B1:B5,"<1,000")'), 3)
  assert.equal(run('COUNTIF(C1:C7,"<m")'), 0)
  assert.equal(run('COUNTIF(E1:E3,"<c")'), 2)
  assert.equal(run('COUNTIF(B1:B5,"10")'), 1, 'equality still matches numeric text')
  assert.equal(run('COUNTIF(B1:B5,"<>10")'), 4)
  assert.equal(run('COUNTIF(B1:B5,"")'), 1)
  assert.equal(run('COUNTIF(B1:B5,"<>")'), 4)
  assert.equal(run('COUNTIF(E1:E3,"b*")'), 1)
  assert.equal(run('COUNTIF(F1:F3,"1*")'), 0, 'wildcards only match text')
  assert.equal(run('COUNTIF(M1:M2,"a~*c")'), 1)
  assert.equal(run('COUNTIF(H1:H4,TRUE)'), 1)
  assert.equal(run('COUNTIF(K1:K2,">0")'), 1, 'errors in the range are skipped, not returned')
  assert.equal(run('COUNTIF(K1:K2,"#DIV/0!")'), 1)
  assert.equal(run('SUMIF(K1:K2,"<>#DIV/0!")'), 1)
  assert.equal(run('MAXIFS(C1:C7,C1:C7,"<1/1/2025")'), 45657)
  assert.equal(run('AVERAGEIF(N1:N2,">$15")'), 20)

  // calc-formula-engine-6: wildcards in exact-match lookups, fractional index arguments.
  assert.equal(run('VLOOKUP("ban*",E1:F3,2,FALSE)'), 2)
  assert.equal(run('VLOOKUP("*err*",E1:F3,2,FALSE)'), 3)
  assert.equal(run('MATCH("?pple",E1:E3,0)'), 1)
  assert.equal(run('MATCH("a~*c",M1:M2,0)'), 1)
  assert.equal(run('MATCH("a*c",M1:M2,0)'), 1)
  assert.equal(run('HLOOKUP("ap*",E1:E3,2,FALSE)'), 'banana')
  assert.equal(run('VLOOKUP("zz*",E1:F3,2,FALSE)'), '#N/A')
  assert.equal(run('INDEX(E1:E3,2.5)'), 'banana')
  assert.equal(run('INDEX(E1:F3,2.9,1.2)'), 'banana')
  assert.equal(run('VLOOKUP("banana",E1:F3,2.5,FALSE)'), 2)
  assert.equal(run('INDEX(E1:E3,-1)'), '#VALUE!')
  assert.equal(run('VLOOKUP("banana",E1:F3,0.5,FALSE)'), '#VALUE!')
  assert.equal(run('INDEX({1,2,3},2)'), 2)
  {
    // Review F8: the wildcard step limit is per cell, so a long range is searched to the end.
    const long = Array.from({ length: 150_000 }, (_, index) => `item ${index + 1} of the list`)
    assert.deepEqual(findLookupIndex(long, '*zzz*', 0), { kind: 'evaluationError', code: '#N/A' }, 'no match is #N/A, not #VALUE!')
    assert.equal(findLookupIndex(long, '*149999 of*', 0), 149_998)
  }

  // calc-formula-engine-7: Excel operator precedence and the no-op unary plus.
  assert.equal(run('-2^2'), 4)
  assert.equal(run('-H1^2'), 4)
  assert.equal(run('2^3^2'), 64)
  assert.equal(run('2^-2'), 0.25)
  assert.equal(run('2*-3^2'), 18)
  assert.equal(run('-2%'), -0.02)
  assert.equal(run('2^3%'), 2 ** 0.03)
  assert.equal(run('+H2'), 'Hello')
  assert.equal(run('+H3'), true)
  assert.equal(run('--H3'), 1)

  // calc-formula-engine-8: DATE coerces text and blanks; VALUE reads dates, times, "-$5".
  assert.equal(run('DATE(LEFT(H4,4),MID(H4,5,2),RIGHT(H4,2))'), 45306)
  assert.equal(run('DATE(2024,1,Z9)'), 45291)
  assert.equal(run('DATE("2024","1","15")'), 45306)
  assert.equal(run('DATE("x",1,1)'), '#VALUE!')
  assert.equal(run('DATE(10000,1,1)'), '#NUM!')
  close(run('VALUE("12:30")'), 0.5208333333333334, 1e-12)
  assert.equal(run('VALUE("2024-01-15")'), 45306)
  close(run('VALUE("2024-01-15 06:00")'), 45306.25, 1e-9)
  assert.equal(run('VALUE("-$5")'), -5)
  assert.equal(run('VALUE("0x10")'), '#VALUE!')
  assert.equal(run('"0x10"+0'), '#VALUE!')

  // calc-formula-engine-11: a single-cell reference follows the reference rule like a range.
  assert.equal(run('SUM(J1)'), 0)
  assert.equal(run('SUM(J1,J2)'), 0)
  assert.equal(run('SUM(J1:J2)'), 0)
  assert.equal(run('SUM(J1,J4)'), 4)
  assert.equal(run('COUNT(J3)'), 0)
  assert.equal(run('MAX(J3)'), 0)
  assert.equal(run('AVERAGE(J3,J4)'), 4)
  assert.equal(run('SUM("5",TRUE)'), 6, 'typed arguments still coerce')
  assert.equal(run('COUNT("5",J1)'), 1)
  assert.equal(run('COUNTA(K2)'), 1, 'COUNTA counts an error in a referenced cell')
  assert.equal(run('AVERAGEA(J1,J4)'), 2, 'AVERAGEA counts referenced text as 0')
  assert.equal(run('SUM(INDEX(J1:J4,1))'), 0)

  // calc-formula-engine-12: Excel's 1900 date system for serials 0-60.
  assert.equal(run('DATE(1900,1,1)'), 1)
  assert.equal(run('DATE(1900,2,28)'), 59)
  assert.equal(run('DATE(1900,2,29)'), 60)
  assert.equal(run('DATE(1900,3,1)'), 61)
  assert.equal(run('DATE(1900,3,0)'), 60)
  assert.equal(run('DATE(1900,1,0)'), 0)
  assert.equal(run('DATE(1900,1,-1)'), '#NUM!')
  assert.equal(run('DAY(60)') + '-' + run('MONTH(60)'), '29-2')
  assert.equal(run('DAY(59)'), 28)
  assert.equal(run('DAY(61)') + '-' + run('MONTH(61)'), '1-3')
  assert.equal(run('YEAR(Z9)') + '/' + run('MONTH(Z9)') + '/' + run('DAY(Z9)'), '1900/1/0')
  assert.equal(run('YEAR(-1)'), '#NUM!')
  assert.equal(run('YEAR(2958466)'), '#NUM!')
  assert.equal(run('DATEVALUE("1/1/1900")'), 1)
  assert.equal(run('DATEVALUE("2/29/1900")'), 60)
  assert.equal(run('DATEVALUE("12/31/1899")'), '#VALUE!')
  assert.equal(run('EDATE(60,1)'), 89)
  assert.equal(run('EOMONTH(DATE(1900,2,1),0)'), 60)
  assert.equal(run('DATEDIF(DATE(1900,1,1),DATE(1900,3,1),"D")'), 60)
  assert.equal(run('WEEKNUM(0)'), 0)
  assert.equal(run('DATE(2024,1,15)'), 45306, 'modern dates are unchanged')

  // CALC-007: unreadable formulas give Excel errors, never #PARSE!.
  assert.equal(run('SUM(1'), '#NAME?')
  assert.equal(run('"abc'), '#NAME?')
  assert.equal(run('1+'), '#NAME?')
  assert.equal(run('A1048577'), '#REF!')
  assert.equal(run('1e999'), '#NUM!')
  assert.equal(run('{1,2;3}'), '#VALUE!')
  assert.equal(run('#PARSE!'), '#NAME?')
}

// CALC-007: a parse-problem object the editor can show, with a repaired suggestion.
{
  const missing = diagnoseFormula('=SUM(A1:A3')
  assert.equal(missing?.error, '#NAME?')
  assert.equal(missing?.kind, 'syntax')
  assert.equal(missing?.message, 'A closing parenthesis is missing.')
  assert.equal(missing?.position, 4)
  assert.equal(missing?.suggestion, '=SUM(A1:A3)')
  assert.equal(diagnoseFormula('="abc')?.suggestion, '="abc"')
  assert.equal(diagnoseFormula('=IF(A1>0,"yes')?.suggestion, '=IF(A1>0,"yes")')
  assert.equal(diagnoseFormula('=SUM(1,2))')?.suggestion, '=SUM(1,2)')
  assert.equal(diagnoseFormula('=SUM(1,2))')?.message, "There's an extra closing parenthesis.")
  assert.equal(diagnoseFormula('=1+')?.suggestion, undefined)
  const unknown = diagnoseFormula('=1+SUMM(A1)')
  assert.equal(unknown?.kind, 'unknown-function')
  assert.equal(unknown?.position, 3)
  assert.equal(unknown?.length, 4)
  assert.equal(diagnoseFormula('=_xlfn.XLOOKUP(1,A1:A2,B1:B2)'), null)
  assert.equal(diagnoseFormula('=LET(x,1,f,LAMBDA(y,y+x),f(2))'), null)
  assert.equal(diagnoseFormula('=SUM(A1:A3)'), null)
  assert.equal(diagnoseFormula('=Total*2'), null, 'names are only checked when the host can resolve them')
  assert.equal(diagnoseFormula('=Total*2', { isDefinedName: () => false })?.kind, 'unknown-name')
  assert.equal(diagnoseFormula('=Total*2', { isDefinedName: (name) => name === 'Total' }), null)
  assert.equal(diagnoseFormula('=A1048577+1')?.error, '#REF!')
  assert.equal(diagnoseFormula('=A1048577+1')?.kind, 'reference')
  assert.equal(diagnoseFormula('=')?.message, 'The formula is empty.')
  assert.match(describeFormulaError('#DIV/0!') ?? '', /divides by zero/)
  assert.equal(describeFormulaError(42), null)
}

process.stdout.write('Formula QA passed: 42 legacy checks plus 326 new formula, grammar, and hook checks, and the engine regression cases.\n')
