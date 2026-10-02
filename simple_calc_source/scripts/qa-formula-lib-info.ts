import type { Cells, Harness } from './qa-formula-library.ts'

export default function run(h: Harness): void {
  h.section('info: IS* and TYPE')
  h.eq('ISEVEN(-1)', false)
  h.eq('ISEVEN(2.5)', true)
  h.eq('ISEVEN(5)', false)
  h.eq('ISEVEN(0)', true)
  h.eq('ISEVEN(-2.9)', true)
  h.eq('ISEVEN(Z99)', true)
  h.eq('ISEVEN("4")', true)
  h.eq('ISEVEN("a")', '#VALUE!')
  h.eq('ISEVEN(TRUE)', '#VALUE!')
  h.eq('ISODD(-1)', true)
  h.eq('ISODD(2.5)', false)
  h.eq('ISODD(5)', true)
  h.eq('ISODD("3")', true)
  h.eq('ISODD(1/0)', '#DIV/0!')
  h.rowsEq('ISODD(A1:A3)', [[true], [false], [true]])
  h.eq('ISNONTEXT(1)', true)
  h.eq('ISNONTEXT("a")', false)
  h.eq('ISNONTEXT(Z99)', true)
  h.eq('ISNONTEXT(1/0)', true)
  h.eq('ISNONTEXT(F1)', false)
  h.eq('TYPE(1)', 1)
  h.eq('TYPE("a")', 2)
  h.eq('TYPE(TRUE)', 4)
  h.eq('TYPE(1/0)', 16)
  h.eq('TYPE({1,2})', 64)
  h.eq('TYPE(A1:A2)', 64)
  h.eq('TYPE(A1)', 1)
  h.eq('TYPE(F1)', 2)
  h.eq('TYPE(Z99)', 1)
  h.eq('TYPE(LAMBDA(x,x))', 128)
  h.eq('ERROR.TYPE(1/0)', 2)
  h.eq('ERROR.TYPE(NA())', 7)
  h.eq('ERROR.TYPE(#NULL!)', 1)
  h.eq('ERROR.TYPE(#REF!)', 4)
  h.eq('ERROR.TYPE(#NAME?)', 5)
  h.eq('ERROR.TYPE(SQRT(-1))', 6)
  h.eq('ERROR.TYPE("a"+1)', 3)
  h.eq('ERROR.TYPE(#CALC!)', 14)
  h.eq('ERROR.TYPE(1)', '#N/A')
  h.eq('IF(ERROR.TYPE(1/0)<3,CHOOSE(ERROR.TYPE(1/0),"null","div"),"other")', 'div')
  h.eq('TRUE()', true)
  h.eq('FALSE()', false)
  h.eq('AND(TRUE(),NOT(FALSE()))', true)
  h.eq('N(7)', 7)
  h.eq('N("7")', 0)
  h.eq('N(TRUE)', 1)
  h.eq('N(F1)', 0)
  h.eq('N(A3)', 3)
  h.eq('N(1/0)', '#DIV/0!')
  h.eq('N(DATE(2024,1,1))', 45292)
  h.eq('INFO("system")', 'pcdos')
  h.eq('INFO("recalc")', 'Automatic')
  h.eq('INFO("numfile")', 2)
  h.eq('INFO("release")', '16.0')
  h.eq('INFO("bogus")', '#VALUE!')
  h.eq('ISEMAIL("a@b.com")', true)
  h.eq('ISEMAIL("first.last+tag@mail.example.org")', true)
  h.eq('ISEMAIL("a@b")', false)
  h.eq('ISEMAIL("not an email")', false)
  h.eq('ISEMAIL(5)', false)
  h.eq('ISURL("google.com")', true)
  h.eq('ISURL("https://x.org/a?b=1")', true)
  h.eq('ISURL("http://localhost")', false)
  h.eq('ISURL("not a url")', false)
  h.eq('ISBETWEEN(5,1,10)', true)
  h.eq('ISBETWEEN(1,1,10,FALSE)', false)
  h.eq('ISBETWEEN(10,1,10,TRUE,FALSE)', false)
  h.eq('ISBETWEEN(10,1,10)', true)
  h.eq('ISBETWEEN(11,1,10)', false)
  h.eq('ISBETWEEN("b","a","c")', true)
  h.rowsEq('ISBETWEEN(A1:A3,2,3)', [[false], [true], [true]])

  h.section('info: Sheets operators')
  h.eq('ADD(2,3)', 5)
  h.eq('MINUS(2,3)', -1)
  h.eq('MULTIPLY(2,3)', 6)
  h.eq('DIVIDE(1,4)', 0.25)
  h.eq('DIVIDE(1,0)', '#DIV/0!')
  h.eq('POW(2,10)', 1024)
  h.eq('EQ("a","A")', true)
  h.eq('NE(1,2)', true)
  h.eq('GT(2,1)', true)
  h.eq('GTE(1,1)', true)
  h.eq('LT(1,2)', true)
  h.eq('LTE(2,1)', false)
  h.eq('UMINUS(3)', -3)
  h.eq('UPLUS(4)', 4)
  h.eq('UNARY_PERCENT(50)', 0.5)
  h.rowsEq('ADD(A1:A3,1)', [[2], [3], [4]])

  h.section('info: database functions')
  const db: Cells = {
    A1: 'Tree', B1: 'Height', C1: 'Age', D1: 'Yield', E1: 'Profit',
    A2: 'Apple', B2: 18, C2: 20, D2: 14, E2: 105,
    A3: 'Pear', B3: 12, C3: 12, D3: 10, E3: 96,
    A4: 'Cherry', B4: 13, C4: 14, D4: 9, E4: 105,
    A5: 'Apple', B5: 14, C5: 15, D5: 10, E5: 75,
    A6: 'Pear', B6: 9, C6: 8, D6: 8, E6: 76.8,
    A7: 'Apple', B7: 8, C7: 9, D7: 6, E7: 45,
    // Apple with 10 < Height < 16 (AND across a row).
    H1: 'Tree', I1: 'Height', J1: 'Height', H2: { value: '=Apple' }, I2: '>10', J2: '<16',
    // Apple OR Pear (OR across rows).
    H4: 'tree', H5: { value: '=Apple' }, H6: { value: '=Pear' },
    // Text without an operator matches values that begin with it.
    L1: 'Tree', L2: 'P',
    N1: 'Profit', N2: '>=100',
    // A blank criteria row matches every record.
    P1: 'Tree',
    R1: 'Tree', R2: { value: '=Plum' },
    T1: 'Tree', T2: 'App',
    V1: 'Tree', V2: { value: '=App' },
    X1: 'Tree', X2: '*rr*',
    Z1: 'Tree', Z2: 'apple',
    // Apple rows OR any tree with Yield > 9.
    AB1: 'Tree', AC1: 'Yield', AB2: { value: '=Apple' }, AC3: '>9',
    AE1: 'Height', AE2: 12,
    AG1: 'Tree', AG2: { value: '<>Apple' },
  }
  h.eq('DSUM(A1:E7,"Profit",H1:J2)', 75, db)
  h.eq('DGET(A1:E7,"Yield",H1:J2)', 10, db)
  h.eq('DCOUNT(A1:E7,"Age",H1:J2)', 1, db)
  h.near('DSUM(A1:E7,"Profit",H4:H6)', 397.8, 1e-12, db)
  h.eq('DSUM(A1:E7,5,H4:H6)', h.value('DSUM(A1:E7,"Profit",H4:H6)', db), db)
  h.eq('DAVERAGE(A1:E7,"Yield",H4:H6)', 9.6, db)
  h.eq('DMAX(A1:E7,"Profit",H4:H6)', 105, db)
  h.eq('DMIN(A1:E7,"Profit",H4:H6)', 45, db)
  h.eq('DCOUNT(A1:E7,,H4:H6)', 5, db)
  h.eq('DCOUNTA(A1:E7,"Tree",H4:H6)', 5, db)
  h.eq('DCOUNTA(A1:E7,,H4:H6)', 5, db)
  h.eq('DPRODUCT(A1:E7,"Yield",H4:H6)', 67200, db)
  h.eq('DGET(A1:E7,"Yield",H4:H6)', '#NUM!', db)
  h.near('DSTDEV(A1:E7,"Yield",H4:H6)', 2.9664793948382653, 1e-12, db)
  h.near('DVAR(A1:E7,"Yield",H4:H6)', 8.8, 1e-12, db)
  h.near('DVARP(A1:E7,"Yield",H4:H6)', 7.04, 1e-12, db)
  h.near('DSTDEVP(A1:E7,"Yield",H4:H6)', 2.6532998322843198, 1e-12, db)
  h.near('DSUM(A1:E7,"Profit",L1:L2)', 172.8, 1e-12, db)
  h.eq('DCOUNT(A1:E7,"Height",L1:L2)', 2, db)
  h.eq('DSUM(A1:E7,"Yield",N1:N2)', 23, db)
  h.eq('DAVERAGE(A1:E7,"Height",N1:N2)', 15.5, db)
  h.near('DSUM(A1:E7,"Profit",P1:P2)', 502.8, 1e-12, db)
  h.eq('DCOUNT(A1:E7,"Tree",P1:P2)', 0, db)
  h.eq('DCOUNTA(A1:E7,"Tree",P1:P2)', 6, db)
  h.eq('DGET(A1:E7,"Yield",R1:R2)', '#VALUE!', db)
  h.eq('DCOUNTA(A1:E7,"Tree",T1:T2)', 3, db)
  h.eq('DCOUNTA(A1:E7,"Tree",V1:V2)', 0, db)
  h.eq('DGET(A1:E7,"Tree",X1:X2)', 'Cherry', db)
  h.eq('DCOUNTA(A1:E7,"Tree",Z1:Z2)', 3, db)
  h.eq('DSUM(A1:E7,"Profit",AB1:AC3)', 105 + 96 + 75 + 45, db)
  h.eq('DGET(A1:E7,"Tree",AE1:AE2)', 'Pear', db)
  h.eq('DCOUNTA(A1:E7,"Tree",AG1:AG2)', 3, db)
  h.eq('DSUM(A1:E7,"Nope",H4:H6)', '#VALUE!', db)
  h.eq('DSUM(A1:E7,6,H4:H6)', '#VALUE!', db)
  h.eq('DSUM(A1:E7,,H4:H6)', '#VALUE!', db)
  h.eq('DAVERAGE(A1:E7,"Yield",R1:R2)', '#DIV/0!', db)
  h.eq('DMAX(A1:E7,"Yield",R1:R2)', 0, db)
  h.eq('DSTDEV(A1:E7,"Yield",H1:J2)', '#DIV/0!', db)
  h.eq('DSUM(A1:E7,"profit",H4:H6)', h.value('DSUM(A1:E7,"Profit",H4:H6)', db), db)

  // calc-formula-engine-1: database criteria read dates, percentages, currency, and
  // thousands like typed input, and "<"/">" with text only compare text.
  const dated: Cells = {
    A1: 'Day', B1: 'Amount', C1: 'Rate',
    A2: 45000, B2: 1000, C2: 0.25,
    A3: 45292, B3: 2000, C3: 0.5,
    A4: 45400, B4: 3000, C4: 0.75,
    A5: 'pending', B5: 4000, C5: 1,
    E1: 'Day', E2: '>=1/1/2024',
    F1: 'Day', F2: '1/1/2024',
    G1: 'Rate', G2: '>50%',
    H1: 'Amount', H2: '>$1,500',
    I1: 'Day', I2: '<q',
  }
  h.eq('DSUM(A1:C5,"Amount",E1:E2)', 5000, dated)
  h.eq('DSUM(A1:C5,"Amount",F1:F2)', 2000, dated)
  h.eq('DCOUNT(A1:C5,"Rate",G1:G2)', 2, dated)
  h.eq('DSUM(A1:C5,"Amount",H1:H2)', 9000, dated)
  h.eq('DSUM(A1:C5,"Amount",I1:I2)', 4000, dated)

  {
    // 20,000-row database with a two-row OR criteria range stays fast.
    const big: Cells = { A1: 'Key', B1: 'Amount', D1: 'Key', D2: 'k1', D3: 'k7' }
    for (let row = 2; row <= 20_001; row += 1) {
      big[`A${row}`] = `k${row % 10}`
      big[`B${row}`] = row
    }
    const started = performance.now()
    const expected = Array.from({ length: 20_000 }, (_value, index) => index + 2)
      .filter((row) => row % 10 === 1 || row % 10 === 7)
      .reduce((sum, row) => sum + row, 0)
    h.eq('DSUM(A1:B20001,"Amount",D1:D3)', expected, big)
    h.check('DSUM over 20k rows is fast', performance.now() - started < 2000, `${Math.round(performance.now() - started)}ms`)
  }
}
