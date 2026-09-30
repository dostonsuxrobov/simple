import type { Harness } from './qa-formula-library.ts'

// Expected values come from Microsoft's documentation examples for each function (quoted to
// the digits shown there, hence the looser tolerances) or from exact closed forms.
const DOC = 1e-6
const LOOSE = 5e-5

export default function run(h: Harness): void {
  h.section('stats')
  const { eq, near, rowsEq, rowsNear } = h

  // Sample sheet for this part.
  const cells: Record<string, unknown> = {
    // PERCENTILE.EXC / PERCENTRANK.EXC docs: 1,2,3,6,6,6,7,8,9
    A1: 1, A2: 2, A3: 3, A4: 6, A5: 6, A6: 6, A7: 7, A8: 8, A9: 9,
    // PERCENTRANK.INC docs: 13,12,11,8,4,3,2,1,1,1
    B1: 13, B2: 12, B3: 11, B4: 8, B5: 4, B6: 3, B7: 2, B8: 1, B9: 1, B10: 1,
    // QUARTILE.EXC docs: 6,7,15,36,39,40,41,42,43,47,49
    C1: 6, C2: 7, C3: 15, C4: 36, C5: 39, C6: 40, C7: 41, C8: 42, C9: 43, C10: 47, C11: 49,
    // LINEST multiple regression docs (office buildings): x1..x4 in D:G, y in H.
    D1: 2310, E1: 2, F1: 2, G1: 20, H1: 142000,
    D2: 2333, E2: 2, F2: 2, G2: 12, H2: 144000,
    D3: 2356, E3: 3, F3: 1.5, G3: 33, H3: 151000,
    D4: 2379, E4: 3, F4: 2, G4: 43, H4: 150000,
    D5: 2402, E5: 2, F5: 3, G5: 53, H5: 139000,
    D6: 2425, E6: 4, F6: 2, G6: 23, H6: 169000,
    D7: 2448, E7: 2, F7: 1.5, G7: 99, H7: 126000,
    D8: 2471, E8: 2, F8: 2, G8: 34, H8: 142900,
    D9: 2494, E9: 3, F9: 3, G9: 23, H9: 163000,
    D10: 2517, E10: 4, F10: 4, G10: 55, H10: 169000,
    D11: 2540, E11: 2, F11: 3, G11: 22, H11: 149000,
    // LOGEST / GROWTH docs: months 11..16, units
    I1: 11, J1: 33100, I2: 12, J2: 47300, I3: 13, J3: 69000, I4: 14, J4: 102000, I5: 15, J5: 150000, I6: 16, J6: 220000,
    I7: 17, I8: 18,
    // FREQUENCY docs: scores and bins
    K1: 79, K2: 85, K3: 78, K4: 85, K5: 50, K6: 81, K7: 95, K8: 88, K9: 97,
    L1: 70, L2: 79, L3: 89,
    // T.TEST docs
    M1: 3, M2: 4, M3: 5, M4: 8, M5: 9, M6: 1, M7: 2, M8: 4, M9: 5,
    N1: 6, N2: 19, N3: 3, N4: 2, N5: 14, N6: 4, N7: 5, N8: 17, N9: 1,
    // CHISQ.TEST docs: actual O1:P3, expected Q1:R3
    O1: 58, P1: 35, O2: 11, P2: 25, O3: 10, P3: 23,
    Q1: 45.35, R1: 47.65, Q2: 17.56, R2: 18.44, Q3: 16.09, R3: 16.91,
    // Z.TEST docs
    S1: 3, S2: 6, S3: 7, S4: 8, S5: 6, S6: 5, S7: 4, S8: 2, S9: 1, S10: 9,
    // Mixed data for *A functions: 1, "x", TRUE, blank, 5
    T1: 1, T2: 'x', T3: true, T5: 5,
    // PROB docs
    U1: 0, U2: 1, U3: 2, U4: 3, V1: 0.2, V2: 0.3, V3: 0.1, V4: 0.4,
    // MODE.MULT docs
    W1: 1, W2: 2, W3: 3, W4: 4, W5: 3, W6: 2, W7: 1, W8: 2, W9: 3, W10: 5, W11: 6, W12: 1,
    // TREND docs: months 1..12 and costs
    X1: 133890, X2: 135000, X3: 135790, X4: 137300, X5: 138130, X6: 139100, X7: 139900, X8: 141120,
    X9: 141890, X10: 143230, X11: 144000, X12: 145290,
    Y1: 1, Y2: 2, Y3: 3, Y4: 4, Y5: 5, Y6: 6, Y7: 7, Y8: 8, Y9: 9, Y10: 10, Y11: 11, Y12: 12,
    Z1: 13, Z2: 14, Z3: 15, Z4: 16, Z5: 17,
  }
  const v = (formula: string, expected: number, tolerance = DOC) => near(formula, expected, tolerance, cells)
  const e = (formula: string, expected: number | string | boolean) => eq(formula, expected, cells)

  // ---- erf / erfc -------------------------------------------------------------------------
  v('ERF(0.745)', 0.70792892)
  v('ERF(1)', 0.84270079)
  near('ERF(0.5)', 0.5204998778130465, 1e-14)
  near('ERF(-1)', -0.8427007929497149, 1e-14)
  near('ERF(0,1)', 0.8427007929497149, 1e-14)
  near('ERF(1,2)', 0.9953222650189527 - 0.8427007929497149, 1e-13)
  near('ERF.PRECISE(0.745)', 0.7079289200957377, 1e-12)
  near('ERF.PRECISE(1)', 0.8427007929497149, 1e-14)
  near('ERFC(1)', 0.15729920705028513, 1e-14)
  near('ERFC(2)', 0.004677734981047266, 1e-13)
  near('ERFC(5)', 1.5374597944280349e-12, 1e-12)
  near('ERFC(-1)', 1.8427007929497148, 1e-14)
  near('ERFC.PRECISE(1)', 0.15729920705028513, 1e-14)
  eq('ERF(0)', 0)
  eq('ERFC(0)', 1)
  eq('ERF("a")', '#VALUE!')
  rowsNear('ERF({0.5;1})', [[0.5204998778130465], [0.8427007929497149]], 1e-13)

  // ---- *A aggregates ---------------------------------------------------------------------
  e('MAXA(T1:T5)', 5)
  e('MINA(T1:T5)', 0) // "x" counts as 0 inside a range
  e('MAXA(-1,-2,TRUE)', 1)
  e('MINA(2,FALSE)', 0)
  e('MAXA(T4)', 0)
  e('MAXA("3",1)', 3)
  e('MAXA("abc",1)', '#VALUE!')
  // T1:T5 as *A values: 1, 0, 1, 5
  v('AVERAGEA(T1:T5)', 7 / 4, 1e-12)
  v('STDEVA(T1:T5)', Math.sqrt(((1 - 1.75) ** 2 * 2 + 1.75 ** 2 + 3.25 ** 2) / 3), 1e-12)
  v('STDEVPA(T1:T5)', Math.sqrt(((1 - 1.75) ** 2 * 2 + 1.75 ** 2 + 3.25 ** 2) / 4), 1e-12)
  v('VARA(T1:T5)', ((1 - 1.75) ** 2 * 2 + 1.75 ** 2 + 3.25 ** 2) / 3, 1e-12)
  v('VARPA(T1:T5)', ((1 - 1.75) ** 2 * 2 + 1.75 ** 2 + 3.25 ** 2) / 4, 1e-12)
  eq('STDEVPA(B1:B8)', 2)
  eq('VARPA(B1:B8)', 4)
  e('STDEVA(1)', '#DIV/0!')
  e('VARPA(T4)', '#DIV/0!')
  // Docs example: STDEVA over 1345,1301,1368,1322,1310,1370,1318,1350,1303,1299
  near('STDEVA(1345,1301,1368,1322,1310,1370,1318,1350,1303,1299)', 27.46391572, 1e-8)
  near('VARA(1345,1301,1368,1322,1310,1370,1318,1350,1303,1299)', 754.2666667, 1e-8)
  near('VARPA(1345,1301,1368,1322,1310,1370,1318,1350,1303,1299)', 678.84, 1e-10)
  near('STDEVPA(1345,1301,1368,1322,1310,1370,1318,1350,1303,1299)', 26.05455814, 1e-8)

  // ---- descriptive ---------------------------------------------------------------------------
  near('GEOMEAN(4,5,8,7,11,4,3)', 5.476986969656962, 1e-12)
  near('HARMEAN(4,5,8,7,11,4,3)', 5.028375962061728, 1e-12)
  eq('GEOMEAN(4,0)', '#NUM!')
  eq('GEOMEAN(4,-1)', '#NUM!')
  eq('HARMEAN(0,1)', '#NUM!')
  eq('GEOMEAN(C1)', '#NUM!')
  eq('GEOMEAN(2,8)', 4)
  near('AVEDEV(4,5,6,7,5,4,3)', 1.020408163, 1e-8)
  eq('AVEDEV(A1:A10)', 2.5)
  eq('DEVSQ(4,5,8,7,11,4,3)', 48)
  eq('DEVSQ(B1:B8)', 32)
  eq('DEVSQ(C1)', '#NUM!')
  near('KURT(3,4,5,2,3,4,5,6,4,7)', -0.151799637, 1e-8)
  near('SKEW(3,4,5,2,3,4,5,6,4,7)', 0.359543071, 1e-8)
  near('SKEW.P(3,4,5,2,3,4,5,6,4,7)', 0.303193339, 1e-8)
  eq('KURT(1,2,3)', '#DIV/0!')
  eq('KURT(1,1,1,1)', '#DIV/0!')
  eq('SKEW(1,2)', '#DIV/0!')
  eq('SKEW(2,2,2)', '#DIV/0!')
  eq('SKEW.P(1,2)', '#DIV/0!')
  eq('SKEW(A1:A10)', 0)
  near('STANDARDIZE(42,40,1.5)', 1.333333333, 1e-9)
  eq('STANDARDIZE(1,1,0)', '#NUM!')
  eq('STANDARDIZE(1,1,-1)', '#NUM!')
  near('TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)', 3.777777778, 1e-9)
  eq('TRIMMEAN(A1:A10,0)', 5.5)
  eq('TRIMMEAN(A1:A10,0.2)', 5.5)
  eq('TRIMMEAN({1,2,3,100},0.5)', 2.5)
  eq('TRIMMEAN(A1:A10,1)', '#NUM!')
  eq('TRIMMEAN(A1:A10,-0.1)', '#NUM!')

  // ---- percentiles / ranks -------------------------------------------------------------------
  e('PERCENTILE.EXC(A1:A9,0.25)', 2.5)
  e('PERCENTILE.EXC(A1:A9,0)', '#NUM!')
  e('PERCENTILE.EXC(A1:A9,0.01)', '#NUM!')
  e('PERCENTILE.EXC(A1:A9,2)', '#NUM!')
  e('PERCENTILE.EXC(A1:A9,0.5)', 6)
  v('PERCENTILE.EXC(A1:A9,0.9)', 9, 1e-12)
  v('PERCENTILE.EXC(A1:A9,0.75)', 7.5, 1e-12)
  e('QUARTILE.EXC(C1:C11,1)', 15)
  e('QUARTILE.EXC(C1:C11,3)', 43)
  e('QUARTILE.EXC(C1:C11,2)', 40)
  e('QUARTILE.EXC(C1:C11,0)', '#NUM!')
  e('QUARTILE.EXC(C1:C11,4)', '#NUM!')
  rowsEq('QUARTILE.EXC(C1:C11,{1;2;3})', [[15], [40], [43]], cells)
  // AGGREGATE 18/19 route to PERCENTILE.EXC/QUARTILE.EXC.
  e('AGGREGATE(18,6,A1:A9,0.25)', 2.5)
  e('AGGREGATE(19,6,C1:C11,1)', 15)
  e('PERCENTRANK.INC(B1:B10,2)', 0.333)
  e('PERCENTRANK.INC(B1:B10,4)', 0.555)
  e('PERCENTRANK.INC(B1:B10,8)', 0.666)
  e('PERCENTRANK.INC(B1:B10,5)', 0.583)
  e('PERCENTRANK(B1:B10,5)', 0.583)
  e('PERCENTRANK.INC(B1:B10,5,1)', 0.5)
  e('PERCENTRANK.INC(B1:B10,13)', 1)
  e('PERCENTRANK.INC(B1:B10,1)', 0)
  e('PERCENTRANK.INC(B1:B10,0)', '#N/A')
  e('PERCENTRANK.INC(B1:B10,14)', '#N/A')
  e('PERCENTRANK.INC(B1:B10,5,0)', '#NUM!')
  e('PERCENTRANK.EXC(A1:A9,7)', 0.7)
  e('PERCENTRANK.EXC(A1:A9,5.43)', 0.381)
  e('PERCENTRANK.EXC(A1:A9,5.43,1)', 0.3)
  e('PERCENTRANK.EXC(A1:A9,1)', 0.1)
  e('PERCENTRANK.EXC(A1:A9,10)', '#N/A')
  e('RANK.AVG(94,{89,88,92,101,94,97,95})', 4)
  // Default sheet B1:B8 = 2,4,4,4,5,5,7,9
  eq('RANK.AVG(4,B1:B8)', 6)
  eq('RANK.AVG(4,B1:B8,1)', 3)
  eq('RANK.AVG(3,B1:B8)', '#N/A')
  e('RANK.AVG(1,B8:B10)', 2)
  e('RANK.AVG(4,B1:B8)', 5)
  rowsEq('RANK.AVG({2;9},B1:B8)', [[8], [1]])
  // RANK/RANK.EQ return #N/A when the number is absent (core fix).
  eq('RANK(3,B1:B8)', '#N/A')
  eq('RANK.EQ(3,B1:B8,1)', '#N/A')
  eq('RANK.EQ(4,B1:B8)', 5)
  eq('RANK(4,B1:B8,1)', 2)
  rowsEq('MODE.MULT(W1:W12)', [[1], [2], [3]], cells)
  rowsEq('MODE.MULT(1,2,2,3,3)', [[2], [3]], cells)
  e('MODE.MULT(1,2,3)', '#N/A')
  e('MODE.MULT(C1)', '#N/A')
  rowsEq('FREQUENCY(K1:K9,L1:L3)', [[1], [2], [4], [2]], cells)
  rowsEq('FREQUENCY(K1:K9,{89,70,79})', [[4], [1], [2], [2]], cells)
  rowsEq('FREQUENCY(A1:A10,5)', [[5], [5]])
  rowsEq('FREQUENCY({1,2,2,3},{2,2})', [[3], [0], [1]], cells)
  rowsEq('FREQUENCY(A1:A10,C1)', [[10]])

  // ---- paired / regression -------------------------------------------------------------------
  near('CORREL({3,2,4,5,6},{9,7,12,15,17})', 0.997054486, 1e-8)
  near('PEARSON({9,7,5,3,1},{10,6,1,5,3})', 0.699379, 1e-6)
  near('RSQ(D1:D7,E1:E7)', 0.05795, 1e-4)
  near('SLOPE(D1:D7,E1:E7)', 0.305555556, 1e-8)
  near('INTERCEPT(D1:D7,E1:E7)', 5 - (11 / 36) * 6, 1e-12)
  near('INTERCEPT({2,3,9,1,8},{6,5,11,7,5})', 0.0483871, 1e-6)
  near('STEYX(D1:D7,E1:E7)', 3.305718950210041, 1e-9)
  eq('COVAR({3,2,4,5,6},{9,7,12,15,17})', 5.2)
  eq('COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})', 5.2)
  near('COVARIANCE.S({2,4,8},{5,11,12})', 9.666666667, 1e-9)
  near('FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})', 10.607253, 1e-7)
  near('FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})', 10.607253, 1e-7)
  rowsNear('FORECAST({30;20},{6,7,9,15,21},{20,28,31,38,40})', [[10.607253086], [3.516203704]], 1e-9)
  eq('CORREL({1,2},{1,2,3})', '#N/A')
  eq('CORREL({1,1,1},{1,2,3})', '#DIV/0!')
  eq('SLOPE({1,2,3},{1,1,1})', '#DIV/0!')
  eq('STEYX({1,2},{1,2})', '#DIV/0!')
  eq('COVARIANCE.S({1},{2})', '#DIV/0!')
  eq('CORREL(A1:A3,{1,#N/A,3})', '#N/A')
  // Non-numeric pairs are skipped.
  near('CORREL({1,2,"x",4},{2,4,100,8})', 1, 1e-12)
  eq('SLOPE(A1:A10,A1:A10)', 1)
  eq('INTERCEPT({3,5,7},{1,2,3})', 1)

  // LINEST: simple fit and docs multiple regression with stats.
  rowsNear('LINEST({1,9,5,7},{0,4,2,3})', [[2, 1]], 1e-12)
  rowsNear('LINEST({3,5,7})', [[2, 1]], 1e-12)
  rowsNear('LINEST({1;9;5;7},{0;4;2;3},FALSE)', [[(4 * 9 + 2 * 5 + 3 * 7) / (16 + 4 + 9), 0]], 1e-12)
  near('SUM(LINEST({3100,4500,4400,5400,7500,8100},{1,2,3,4,5,6})*{9,1})', 11000, 1e-9)
  rowsNear('LINEST(H1:H11,D1:G11,TRUE,TRUE)', [
    [-234.2371645, 2553.21066, 12529.76817, 27.64138737, 52317.83051],
    [13.26801148, 530.6691519, 400.0668382, 5.429374042, 12237.3616],
    [0.996747993, 970.5784629, '#N/A', '#N/A', '#N/A'],
    [459.7536742, 6, '#N/A', '#N/A', '#N/A'],
    [1732393319, 5652135.316, '#N/A', '#N/A', '#N/A'],
  ] as never, 1e-8, cells)
  // Single x with stats: y = 1,9,5,7; x = 0,4,2,3 → exact fit? residuals 0,0,0,0 → check se row.
  // Hand-derived: x̄=2.5, ȳ=3.5, Sxx=5, Sxy=4, SSresid=1.8, SSreg=3.2.
  rowsNear('LINEST({2,3,5,4},{1,2,3,4},TRUE,TRUE)', [
    [0.8, 1.5],
    [Math.sqrt(0.9 / 5), Math.sqrt(0.9 * (0.25 + 6.25 / 5))],
    [0.64, Math.sqrt(0.9)],
    [3.2 / 0.9, 2],
    [3.2, 1.8],
  ], 1e-12)
  // Through the origin: m = Σxy/Σx² = 39/30, SSresid = 3.3, SStotal = Σy² = 54.
  rowsNear('LINEST({2,3,5,4},{1,2,3,4},FALSE,TRUE)', [
    [1.3, 0],
    [Math.sqrt(1.1 / 30), '#N/A'],
    [50.7 / 54, Math.sqrt(1.1)],
    [50.7 / 1.1, 3],
    [50.7, 3.3],
  ] as never, 1e-12)
  // Collinear columns are dropped (coefficient 0).
  rowsNear('LINEST({1;2;3;4},{1,2;2,4;3,6;4,8})', [[0, 1, 0]], 1e-10)
  eq('LINEST({1,2,3},{1,2})', '#REF!')
  eq('LINEST({1,"a",3})', '#VALUE!')
  rowsNear('LOGEST(J1:J6,I1:I6)', [[1.463275628, 495.3047702]], 1e-8, cells)
  rowsNear('LOGEST(J1:J6,I1:I6,TRUE,TRUE)', [
    [1.463275628, 495.3047702],
    [0.002633403, 0.035834282],
    [0.99980862, 0.011016315],
    [20896.8011, 4],
    [2.53601883, 0.000485437],
  ], 1e-6, cells)
  eq('LOGEST({1,-2,3})', '#NUM!')
  rowsNear('GROWTH(J1:J6,I1:I6,I7:I8)', [[320196.7184], [468536.0539]], 1e-9, cells)
  rowsNear('GROWTH(J1:J6,I1:I6)', [
    [32618.20377], [47729.42261], [69841.30086], [102197.0734], [149542.4867], [218821.8762],
  ], 1e-9, cells)
  rowsEq('ROUND(TREND(X1:X12,Y1:Y12,Z1:Z5),0)', [[146172], [147190], [148208], [149226], [150244]], cells)
  near('INDEX(TREND(X1:X12,Y1:Y12),2)', 134971.5152, 1e-9, cells)
  near('INDEX(TREND(X1:X12,Y1:Y12),1)', 133953.3333, 1e-9, cells)
  near('INDEX(TREND(X1:X12),12)', 145153.3333, 1e-9, cells)
  rowsNear('TREND({1,2,3},{1,2,3},{4,5})', [[4, 5]], 1e-12)
  rowsNear('TREND({1;2;3},,{4;5})', [[4], [5]], 1e-12)
  rowsNear('TREND(H1:H11,D1:G11,{2500,3,2,25})', [[52317.83051 + 27.64138737 * 2500 + 12529.76817 * 3 + 2553.21066 * 2 - 234.2371645 * 25]], 1e-8, cells)

  // ---- tests ---------------------------------------------------------------------------------
  near('T.TEST(M1:M9,N1:N9,2,1)', 0.196016, 1e-5, cells)
  near('TTEST(M1:M9,N1:N9,2,1)', 0.196016, 1e-5, cells)
  near('T.TEST(M1:M9,N1:N9,1,1)', 0.098008, 1e-5, cells)
  near('T.TEST(M1:M9,N1:N9,2,2)', 0.191995, 1e-5, cells)
  near('T.TEST(M1:M9,N1:N9,2,3)', 0.202293, 1e-5, cells)
  eq('T.TEST(M1:M9,N1:N9,3,1)', '#NUM!', cells)
  eq('T.TEST(M1:M9,N1:N9,2,4)', '#NUM!', cells)
  eq('T.TEST(M1:M9,N1:N8,2,1)', '#N/A', cells)
  near('F.TEST({6,7,9,15,21},{20,28,31,38,40})', 0.64831785, 1e-7)
  near('FTEST({6,7,9,15,21},{20,28,31,38,40})', 0.64831785, 1e-7)
  eq('F.TEST({1},{1,2})', '#DIV/0!')
  near('CHISQ.TEST(O1:P3,Q1:R3)', 0.000308192, 1e-5, cells)
  near('CHITEST(O1:P3,Q1:R3)', 0.000308192, 1e-5, cells)
  eq('CHISQ.TEST(O1:P3,Q1:R2)', '#N/A', cells)
  near('Z.TEST(S1:S10,4)', 0.090574, 1e-5, cells)
  near('Z.TEST(S1:S10,6)', 0.863043, 1e-5, cells)
  near('ZTEST(S1:S10,4)', 0.090574, 1e-5, cells)
  near('Z.TEST(S1:S10,4,3)', 0.12312584984626351, 1e-12, cells) // 1-Φ(1.1/(3/√10))
  eq('Z.TEST(C1,4)', '#N/A')
  eq('Z.TEST({1},1)', '#DIV/0!')

  // ---- normal ----------------------------------------------------------------------------------
  v('NORM.DIST(42,40,1.5,TRUE)', 0.9087888)
  v('NORM.DIST(42,40,1.5,FALSE)', 0.10934005)
  v('NORMDIST(42,40,1.5,TRUE)', 0.9087888)
  eq('NORM.DIST(42,40,0,TRUE)', '#NUM!')
  eq('NORM.DIST(42,40,-1,TRUE)', '#NUM!')
  eq('NORMDIST(42,40,1.5)', '#VALUE!')
  v('NORM.INV(0.908789,40,1.5)', 42.000002)
  v('NORMINV(0.908789,40,1.5)', 42.000002)
  eq('NORM.INV(0,40,1.5)', '#NUM!')
  eq('NORM.INV(1,40,1.5)', '#NUM!')
  eq('NORM.INV(0.5,40,0)', '#NUM!')
  eq('NORM.INV(0.5,40,1.5)', 40)
  v('NORM.S.DIST(1.333333,TRUE)', 0.908788726)
  v('NORM.S.DIST(1.333333,FALSE)', 0.164010148)
  v('NORMSDIST(1.333333)', 0.908788726)
  near('NORM.S.DIST(0,TRUE)', 0.5, 1e-15)
  near('NORM.S.DIST(-8,TRUE)', 6.22096057427178e-16, 1e-10)
  near('NORM.S.DIST(1.96,TRUE)', 0.9750021048517795, 1e-13)
  v('NORM.S.INV(0.908789)', 1.3333347)
  v('NORMSINV(0.908789)', 1.3333347)
  near('NORM.S.INV(0.975)', 1.959963984540054, 1e-13)
  near('NORM.S.INV(0.5)', 0, 1e-15)
  near('NORM.S.INV(1E-10)', -6.361340902404056, 1e-12)
  eq('NORM.S.INV(0)', '#NUM!')
  eq('NORM.S.INV(1)', '#NUM!')
  rowsNear('NORM.S.DIST({-1;0;1},TRUE)', [[0.15865525393145707], [0.5], [0.8413447460685429]], 1e-13)
  v('LOGNORM.DIST(4,3.5,1.2,TRUE)', 0.0390836)
  v('LOGNORM.DIST(4,3.5,1.2,FALSE)', 0.0176176)
  v('LOGNORMDIST(4,3.5,1.2)', 0.0390836)
  eq('LOGNORM.DIST(0,3.5,1.2,TRUE)', '#NUM!')
  eq('LOGNORM.DIST(4,3.5,0,TRUE)', '#NUM!')
  v('LOGNORM.INV(0.039084,3.5,1.2)', 4.0000252)
  v('LOGINV(0.039084,3.5,1.2)', 4.0000252)
  eq('LOGNORM.INV(1,3.5,1.2)', '#NUM!')
  near('PHI(0.75)', 0.301137432, 1e-8)
  near('PHI(0)', 0.3989422804014327, 1e-14)
  near('GAUSS(2)', 0.477249868, 1e-8)
  near('GAUSS(-2)', -0.477249868, 1e-8)
  near('FISHER(0.75)', 0.972955075, 1e-8)
  near('FISHERINV(0.972955)', 0.75, 1e-6)
  eq('FISHER(1)', '#NUM!')
  eq('FISHER(-1)', '#NUM!')
  near('CONFIDENCE.NORM(0.05,2.5,50)', 0.692951912, 1e-8)
  near('CONFIDENCE(0.05,2.5,50)', 0.692951912, 1e-8)
  near('CONFIDENCE.T(0.05,1,50)', 0.284196855, 1e-8)
  eq('CONFIDENCE.NORM(0,2.5,50)', '#NUM!')
  eq('CONFIDENCE.NORM(0.05,2.5,0.5)', '#NUM!')
  eq('CONFIDENCE.T(0.05,1,1)', '#DIV/0!')

  // ---- Student t -------------------------------------------------------------------------------
  v('T.DIST(60,1,TRUE)', 0.99469533)
  v('T.DIST(8,3,FALSE)', 0.00073691)
  near('T.DIST(0,5,TRUE)', 0.5, 1e-15)
  near('T.DIST(-1,1,TRUE)', 0.25, 1e-14)
  near('T.DIST(1,2,TRUE)', 0.5 + 0.5 / Math.sqrt(3), 1e-13)
  near('T.DIST(1,2.9,TRUE)', 0.5 + 0.5 / Math.sqrt(3), 1e-13) // df truncated
  eq('T.DIST(1,0.5,TRUE)', '#NUM!')
  v('T.DIST.2T(1.959999998,60)', 0.05464493)
  v('T.DIST.RT(1.959999998,60)', 0.027322465)
  near('T.DIST.RT(-1,1)', 0.75, 1e-14)
  eq('T.DIST.2T(-1,60)', '#NUM!')
  v('TDIST(1.959999998,60,2)', 0.05464493)
  v('TDIST(1.959999998,60,1)', 0.027322465)
  eq('TDIST(1,60,3)', '#NUM!')
  eq('TDIST(-1,60,1)', '#NUM!')
  v('T.INV(0.75,2)', 0.8164966)
  near('T.INV(0.25,2)', -0.816496580927726, 1e-12)
  near('T.INV(0.5,7)', 0, 1e-15)
  near('T.INV(0.975,1)', 12.706204736174698, 1e-11)
  eq('T.INV(0,2)', '#NUM!')
  eq('T.INV(1,2)', '#NUM!')
  v('T.INV.2T(0.546449,60)', 0.606533)
  v('TINV(0.546449,60)', 0.606533)
  near('T.INV.2T(0.05,10)', 2.228138851986274, 1e-11)
  near('T.INV.2T(1,10)', 0, 1e-15)
  eq('T.INV.2T(0,10)', '#NUM!')
  eq('T.INV.2T(1.5,10)', '#NUM!')
  near('T.DIST.2T(T.INV.2T(0.01,4),4)', 0.01, 1e-12)
  near('T.DIST(T.INV(1E-8,3),3,TRUE)', 1e-8, 1e-9)

  // ---- chi-square -----------------------------------------------------------------------------
  v('CHISQ.DIST(0.5,1,TRUE)', 0.52049988)
  v('CHISQ.DIST(2,3,FALSE)', 0.20755375)
  near('CHISQ.DIST(2,2,TRUE)', 1 - Math.exp(-1), 1e-14)
  near('CHISQ.DIST(0,2,FALSE)', 0.5, 1e-14)
  eq('CHISQ.DIST(-1,2,TRUE)', '#NUM!')
  eq('CHISQ.DIST(1,0,TRUE)', '#NUM!')
  v('CHISQ.DIST.RT(18.307,10)', 0.0500006)
  v('CHIDIST(18.307,10)', 0.0500006)
  eq('CHIDIST(-1,10)', '#NUM!')
  v('CHISQ.INV(0.93,1)', 3.283020287)
  v('CHISQ.INV(0.6,2)', 1.832581464)
  near('CHISQ.INV(0,2)', 0, 1e-15)
  eq('CHISQ.INV(1,2)', '#NUM!')
  eq('CHISQ.INV(-0.1,2)', '#NUM!')
  v('CHISQ.INV.RT(0.050001,10)', 18.30697346)
  v('CHIINV(0.050001,10)', 18.30697346)
  near('CHISQ.INV.RT(1,10)', 0, 1e-15)
  eq('CHISQ.INV.RT(0,10)', '#NUM!')
  near('CHISQ.DIST.RT(CHISQ.INV.RT(1E-12,5),5)', 1e-12, 1e-9)

  // ---- F -----------------------------------------------------------------------------------------
  v('F.DIST(15.2069,6,4,TRUE)', 0.99)
  v('F.DIST(15.2069,6,4,FALSE)', 0.0012238)
  v('F.DIST.RT(15.2069,6,4)', 0.01)
  v('FDIST(15.2069,6,4)', 0.01)
  eq('F.DIST(-1,6,4,TRUE)', '#NUM!')
  eq('F.DIST(1,0,4,TRUE)', '#NUM!')
  near('F.DIST(1,2,2,TRUE)', 0.5, 1e-14)
  v('F.INV(0.01,6,4)', 0.10930991)
  v('F.INV.RT(0.01,6,4)', 15.20686)
  v('FINV(0.01,6,4)', 15.20686)
  near('F.INV(0.5,2,2)', 1, 1e-13)
  eq('F.INV(1,6,4)', '#NUM!')
  eq('F.INV.RT(0,6,4)', '#NUM!')
  near('F.DIST.RT(F.INV.RT(0.001,3,7),3,7)', 0.001, 1e-10)

  // ---- binomial and friends -------------------------------------------------------------------
  v('BINOM.DIST(6,10,0.5,FALSE)', 0.2050781)
  near('BINOM.DIST(6,10,0.5,FALSE)', 210 / 1024, 1e-13)
  near('BINOM.DIST(6,10,0.5,TRUE)', 848 / 1024, 1e-13)
  near('BINOMDIST(6.9,10.2,0.5,FALSE)', 210 / 1024, 1e-13)
  eq('BINOM.DIST(11,10,0.5,FALSE)', '#NUM!')
  eq('BINOM.DIST(-1,10,0.5,FALSE)', '#NUM!')
  eq('BINOM.DIST(1,10,1.5,FALSE)', '#NUM!')
  eq('BINOM.DIST(0,10,0,FALSE)', 1)
  eq('BINOM.DIST(10,10,1,TRUE)', 1)
  near('BINOM.DIST(5000,10000,0.5,TRUE)', 0.5039893230696174, 1e-9)
  v('BINOM.DIST.RANGE(60,0.75,48)', 0.083974967, 1e-7)
  v('BINOM.DIST.RANGE(60,0.75,45,50)', 0.523629793, 1e-7)
  eq('BINOM.DIST.RANGE(60,0.75,50,45)', '#NUM!')
  eq('BINOM.DIST.RANGE(60,1.2,50)', '#NUM!')
  eq('BINOM.INV(6,0.5,0.75)', 4)
  eq('CRITBINOM(6,0.5,0.75)', 4)
  eq('BINOM.INV(6,0.5,0.34375)', 2)
  eq('BINOM.INV(6,0.5,0)', 0)
  eq('BINOM.INV(6,0.5,1)', 6)
  eq('BINOM.INV(6,0.5,1.5)', '#NUM!')
  eq('BINOM.INV(100000,0.5,0.5)', 50000)
  v('POISSON.DIST(2,5,TRUE)', 0.124652)
  v('POISSON.DIST(2,5,FALSE)', 0.084224)
  v('POISSON(2,5,TRUE)', 0.124652)
  eq('POISSON.DIST(-1,5,TRUE)', '#NUM!')
  eq('POISSON.DIST(1,-5,TRUE)', '#NUM!')
  eq('POISSON.DIST(0,0,FALSE)', 1)
  near('POISSON.DIST(2000,2000,TRUE)', 0.5059471, 1e-6)
  v('EXPON.DIST(0.2,10,TRUE)', 0.86466472)
  v('EXPON.DIST(0.2,10,FALSE)', 1.35335283)
  v('EXPONDIST(0.2,10,TRUE)', 0.86466472)
  eq('EXPON.DIST(-1,10,TRUE)', '#NUM!')
  eq('EXPON.DIST(1,0,TRUE)', '#NUM!')
  v('WEIBULL.DIST(105,20,100,TRUE)', 0.929581)
  v('WEIBULL.DIST(105,20,100,FALSE)', 0.035589)
  v('WEIBULL(105,20,100,TRUE)', 0.929581)
  eq('WEIBULL.DIST(-1,20,100,TRUE)', '#NUM!')
  eq('WEIBULL.DIST(1,0,100,TRUE)', '#NUM!')
  v('HYPGEOM.DIST(1,4,8,20,TRUE)', 0.4654, 2e-4)
  v('HYPGEOM.DIST(1,4,8,20,FALSE)', 0.3633, 2e-4)
  near('HYPGEOM.DIST(1,4,8,20,FALSE)', (8 * 220) / 4845, 1e-13)
  near('HYPGEOMDIST(1,4,8,20)', (8 * 220) / 4845, 1e-13)
  eq('HYPGEOM.DIST(5,4,8,20,FALSE)', '#NUM!')
  eq('HYPGEOM.DIST(1,21,8,20,FALSE)', '#NUM!')
  eq('HYPGEOM.DIST(1,4,0,20,FALSE)', '#NUM!')
  v('NEGBINOM.DIST(10,5,0.25,TRUE)', 0.3135141)
  v('NEGBINOM.DIST(10,5,0.25,FALSE)', 0.0550487)
  v('NEGBINOMDIST(10,5,0.25)', 0.0550487)
  eq('NEGBINOM.DIST(10,0,0.25,TRUE)', '#NUM!')
  eq('NEGBINOM.DIST(-1,5,0.25,TRUE)', '#NUM!')
  eq('NEGBINOMDIST(10,5,1.25)', '#NUM!')

  // ---- gamma / beta ---------------------------------------------------------------------------
  near('GAMMA(2.5)', 1.329340388, 1e-9)
  near('GAMMA(-3.75)', 0.267866129, 1e-8)
  near('GAMMA(0.5)', Math.sqrt(Math.PI), 1e-14)
  eq('GAMMA(5)', 24)
  eq('GAMMA(0)', '#NUM!')
  eq('GAMMA(-2)', '#NUM!')
  eq('GAMMA(172)', '#NUM!')
  near('GAMMALN(4)', 1.791759469, 1e-9)
  near('GAMMALN.PRECISE(4)', 1.791759469, 1e-9)
  near('GAMMALN(0.5)', Math.log(Math.sqrt(Math.PI)), 1e-14)
  near('GAMMALN(100)', 359.1342053695754, 1e-14)
  eq('GAMMALN(0)', '#NUM!')
  eq('GAMMALN(-1)', '#NUM!')
  v('GAMMA.DIST(10.00001131,9,2,FALSE)', 0.032639)
  v('GAMMA.DIST(10.00001131,9,2,TRUE)', 0.068094)
  v('GAMMADIST(10.00001131,9,2,TRUE)', 0.068094)
  eq('GAMMA.DIST(-1,9,2,TRUE)', '#NUM!')
  eq('GAMMA.DIST(1,0,2,TRUE)', '#NUM!')
  near('GAMMA.DIST(2,1,1,TRUE)', 1 - Math.exp(-2), 1e-14)
  v('GAMMA.INV(0.068094,9,2)', 10.0000112)
  v('GAMMAINV(0.068094,9,2)', 10.0000112)
  eq('GAMMA.INV(1,9,2)', '#NUM!')
  eq('GAMMA.INV(0.5,0,2)', '#NUM!')
  near('GAMMA.INV(0,9,2)', 0, 1e-15)
  v('BETA.DIST(2,8,10,TRUE,1,3)', 0.6854706)
  v('BETA.DIST(2,8,10,FALSE,1,3)', 1.4837646)
  v('BETADIST(2,8,10,1,3)', 0.6854706)
  near('BETA.DIST(0.5,2,2,TRUE)', 0.5, 1e-14)
  near('BETA.DIST(0.25,1,1,TRUE)', 0.25, 1e-14)
  eq('BETA.DIST(4,8,10,TRUE,1,3)', '#NUM!')
  eq('BETA.DIST(2,0,10,TRUE,1,3)', '#NUM!')
  eq('BETA.DIST(2,8,10,TRUE,3,3)', '#NUM!')
  v('BETA.INV(0.685470581,8,10,1,3)', 2)
  v('BETAINV(0.685470581,8,10,1,3)', 2)
  near('BETA.INV(0.5,2,2)', 0.5, 1e-13)
  eq('BETA.INV(0,8,10)', '#NUM!')
  eq('BETA.INV(0.5,-1,10)', '#NUM!')
  near('BETA.DIST(BETA.INV(0.999,0.5,0.5),0.5,0.5,TRUE)', 0.999, 1e-12)

  // ---- misc ------------------------------------------------------------------------------------
  near('PROB(U1:U4,V1:V4,2)', 0.1, 1e-12, cells)
  near('PROB(U1:U4,V1:V4,1,3)', 0.8, 1e-12, cells)
  eq('PROB(U1:U4,V1:V3,1,3)', '#N/A', cells)
  eq('PROB({0,1},{0.5,0.6},1)', '#NUM!')
  eq('PROB({0,1},{0,1},1)', '#NUM!')
  eq('PERMUT(100,3)', 970200)
  eq('PERMUT(3,2)', 6)
  eq('PERMUT(3.9,2.2)', 6)
  eq('PERMUT(2,3)', '#NUM!')
  eq('PERMUT(0,0)', '#NUM!')
  eq('PERMUT(5,-1)', '#NUM!')
  eq('PERMUTATIONA(3,2)', 9)
  eq('PERMUTATIONA(2,2)', 4)
  eq('PERMUTATIONA(0,0)', 1)
  eq('PERMUTATIONA(-1,2)', '#NUM!')
  rowsEq('PERMUT(5,{1,2,3})', [[5, 20, 60]])
  eq('AVERAGE.WEIGHTED({1,2,3},{1,1,2})', 2.25)
  eq('AVERAGE.WEIGHTED(A1:A2,{3;1},10,2)', (3 + 2 + 20) / 6)
  eq('AVERAGE.WEIGHTED({1,2},{0,0})', '#DIV/0!')
  eq('AVERAGE.WEIGHTED({1,2},{1})', '#VALUE!')
  eq('AVERAGE.WEIGHTED({1,2},{1,1},3)', '#VALUE!')

  // Consistency: inverse(cdf(x)) round trips.
  for (const [dist, inv] of [
    ['NORM.DIST(1.7,0.3,2,TRUE)', 'NORM.INV(NORM.DIST(1.7,0.3,2,TRUE),0.3,2)'],
    ['LOGNORM.DIST(1.7,0.3,2,TRUE)', 'LOGNORM.INV(LOGNORM.DIST(1.7,0.3,2,TRUE),0.3,2)'],
    ['T.DIST(1.7,9,TRUE)', 'T.INV(T.DIST(1.7,9,TRUE),9)'],
    ['CHISQ.DIST(1.7,9,TRUE)', 'CHISQ.INV(CHISQ.DIST(1.7,9,TRUE),9)'],
    ['F.DIST(1.7,9,3,TRUE)', 'F.INV(F.DIST(1.7,9,3,TRUE),9,3)'],
    ['GAMMA.DIST(1.7,2.5,0.7,TRUE)', 'GAMMA.INV(GAMMA.DIST(1.7,2.5,0.7,TRUE),2.5,0.7)'],
    ['BETA.DIST(1.7,2.5,0.7,TRUE,1,2)', 'BETA.INV(BETA.DIST(1.7,2.5,0.7,TRUE,1,2),2.5,0.7,1,2)'],
  ]) {
    h.check(`${dist} is a probability`, typeof h.value(dist) === 'number')
    near(inv, 1.7, 1e-10)
  }
}
