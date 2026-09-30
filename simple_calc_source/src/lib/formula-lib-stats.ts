// Statistical worksheet functions for the extended library (see formula-library.ts):
// descriptive statistics, regression, and probability distributions (plus the legacy
// compatibility names Excel still accepts).

import { finiteResult, isEvaluationError, scalarArgument, toNumber } from "./formulas";
import type { EvaluationError, EvaluationScalar } from "./formulas";
import {
  arrayResult,
  divError,
  integerArg,
  isMissing,
  naError,
  numberArg,
  numericFunction,
  numericPairs,
  numError,
  optionalBooleanArg,
  optionalIntegerArg,
  readNumbers,
  readNumbersA,
  rectArg,
  refError,
  spec,
  valueError,
} from "./formula-lib-shared";
import type { Scalar, Specs, Value } from "./formula-lib-shared";
import * as S from "./formula-lib-special";

type NumberResult = number | EvaluationError;

const truncate = Math.trunc;

// ---- Descriptive helpers -------------------------------------------------------------------

function mean(values: number[]): number {
  let sum = 0;
  for (const value of values) sum += value;
  return sum / values.length;
}

/** Sum of squared deviations from the mean (two-pass for accuracy). */
function sumSquaredDeviations(values: number[]): number {
  const average = mean(values);
  let sum = 0;
  for (const value of values) sum += (value - average) ** 2;
  return sum;
}

function sampleVariance(values: number[]): NumberResult {
  if (values.length < 2) return divError();
  return finiteResult(sumSquaredDeviations(values) / (values.length - 1));
}

function populationVariance(values: number[]): NumberResult {
  if (values.length < 1) return divError();
  return finiteResult(sumSquaredDeviations(values) / values.length);
}

function sortedNumbers(values: number[]): number[] {
  return Float64Array.from(values).sort() as unknown as number[];
}

/** Number of sorted values strictly less than x (binary search). */
function countLess(sorted: ArrayLike<number>, x: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (sorted[mid] < x) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** Number of sorted values less than or equal to x. */
function countLessOrEqual(sorted: ArrayLike<number>, x: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (sorted[mid] <= x) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** Excel's significance truncation (PERCENTRANK rounds down to `digits` places). */
function truncateDigits(value: number, digits: number): number {
  const factor = 10 ** digits;
  return Math.floor(Number((value * factor).toPrecision(15))) / factor;
}

function aggregate(compute: (numbers: number[]) => Value | NumberResult, readA = false) {
  return spec(1, 255, (values) => {
    const numbers = readA ? readNumbersA(values) : readNumbers(values);
    if (isEvaluationError(numbers)) return numbers;
    return compute(numbers);
  });
}

function percentileExclusive(sorted: ArrayLike<number>, k: number): NumberResult {
  const n = sorted.length;
  if (n === 0 || !(k > 0 && k < 1)) return numError();
  const position = k * (n + 1);
  if (position < 1 || position > n) return numError();
  const lower = Math.floor(position);
  const fraction = position - lower;
  const base = sorted[lower - 1];
  if (lower >= n || fraction === 0) return base;
  return finiteResult(base + fraction * (sorted[lower] - base));
}

function percentRank(values: number[], x: number, significance: number, exclusive: boolean): NumberResult {
  if (values.length === 0) return numError();
  if (significance < 1) return numError();
  const sorted = sortedNumbers(values);
  const n = sorted.length;
  if (x < sorted[0] || x > sorted[n - 1]) return naError();
  const smaller = countLess(sorted, x);
  let rank: number;
  if (sorted[smaller] === x) {
    if (exclusive) rank = (smaller + 1) / (n + 1);
    else rank = n === 1 ? 1 : smaller / (n - 1);
  } else {
    const low = sorted[smaller - 1];
    const high = sorted[smaller];
    const fraction = (x - low) / (high - low);
    rank = exclusive ? (smaller + fraction) / (n + 1) : (smaller - 1 + fraction) / (n - 1);
  }
  return truncateDigits(rank, significance);
}

function percentRankSpec(exclusive: boolean) {
  return spec(
    2,
    3,
    (values) => {
      const numbers = readNumbers([values[0]]);
      if (isEvaluationError(numbers)) return numbers;
      const x = numberArg(values[1]);
      if (isEvaluationError(x)) return x;
      const significance = integerArg(values[2], 3);
      if (isEvaluationError(significance)) return significance;
      return percentRank(numbers, x, significance, exclusive);
    },
    { liftArgs: [1, 2] },
  );
}

// ---- Paired data / regression helpers --------------------------------------------------------

interface PairedMoments {
  n: number;
  meanX: number;
  meanY: number;
  sxx: number;
  syy: number;
  sxy: number;
}

function pairedMoments(xs: number[], ys: number[]): PairedMoments {
  const n = xs.length;
  const meanX = n ? mean(xs) : NaN;
  const meanY = n ? mean(ys) : NaN;
  let sxx = 0;
  let syy = 0;
  let sxy = 0;
  for (let index = 0; index < n; index += 1) {
    const dx = xs[index] - meanX;
    const dy = ys[index] - meanY;
    sxx += dx * dx;
    syy += dy * dy;
    sxy += dx * dy;
  }
  return { n, meanX, meanY, sxx, syy, sxy };
}

/** Moments of (known_x, known_y) given Excel's (known_y, known_x) argument order. */
function regressionMoments(knownY: Value | undefined, knownX: Value | undefined): PairedMoments | EvaluationError {
  const pairs = numericPairs(knownX, knownY);
  if (isEvaluationError(pairs)) return pairs;
  return pairedMoments(pairs.xs, pairs.ys);
}

function pairedSpec(compute: (moments: PairedMoments) => NumberResult, yFirst = false) {
  return spec(2, 2, (values) => {
    const moments = yFirst ? regressionMoments(values[0], values[1]) : (() => {
      const pairs = numericPairs(values[0], values[1]);
      return isEvaluationError(pairs) ? pairs : pairedMoments(pairs.xs, pairs.ys);
    })();
    if (isEvaluationError(moments)) return moments;
    const result = compute(moments);
    return typeof result === "number" ? finiteResult(result) : result;
  });
}

function correlation(moments: PairedMoments): NumberResult {
  if (moments.n === 0 || moments.sxx === 0 || moments.syy === 0) return divError();
  return Math.max(-1, Math.min(1, moments.sxy / Math.sqrt(moments.sxx * moments.syy)));
}

interface RegressionData {
  y: number[];
  /** n rows of k x-values. */
  x: number[][];
  k: number;
  /** y was supplied as a single row (x variables are then rows of known_x). */
  byRow: boolean;
  /** Shape of known_y for default new_x. */
  yRows: number;
  yColumns: number;
}

function numericGrid(value: Value | undefined): { values: number[]; rowCount: number; columnCount: number } | EvaluationError {
  const rect = rectArg(value);
  if (isEvaluationError(rect)) return rect;
  const values: number[] = [];
  for (const entry of rect.values) {
    if (isEvaluationError(entry)) return entry;
    if (typeof entry !== "number") return valueError();
    values.push(entry);
  }
  return { values, rowCount: rect.rowCount, columnCount: rect.columnCount };
}

function regressionData(knownY: Value | undefined, knownX: Value | undefined): RegressionData | EvaluationError {
  const y = numericGrid(knownY);
  if (isEvaluationError(y)) return y;
  const n = y.values.length;
  const byRow = y.rowCount === 1 && y.columnCount > 1;
  if (isMissing(knownX)) {
    return {
      y: y.values,
      x: y.values.map((_value, index) => [index + 1]),
      k: 1,
      byRow,
      yRows: y.rowCount,
      yColumns: y.columnCount,
    };
  }
  const x = numericGrid(knownX);
  if (isEvaluationError(x)) return x;
  const rows: number[][] = [];
  let k: number;
  if (x.values.length === n && (x.rowCount === y.rowCount || x.columnCount === y.columnCount || n === 1 || (y.rowCount > 1 && y.columnCount > 1))) {
    // Same number of cells in a matching orientation: one x variable.
    if (y.rowCount > 1 && y.columnCount > 1 && (x.rowCount !== y.rowCount || x.columnCount !== y.columnCount)) return refError();
    k = 1;
    for (const value of x.values) rows.push([value]);
  } else if (!byRow && y.columnCount === 1 && x.rowCount === n) {
    k = x.columnCount;
    for (let row = 0; row < n; row += 1) rows.push(x.values.slice(row * k, row * k + k));
  } else if (byRow && x.columnCount === n) {
    k = x.rowCount;
    for (let column = 0; column < n; column += 1) {
      const row: number[] = [];
      for (let variable = 0; variable < k; variable += 1) row.push(x.values[variable * n + column]);
      rows.push(row);
    }
  } else {
    return refError();
  }
  return { y: y.values, x: rows, k, byRow, yRows: y.rowCount, yColumns: y.columnCount };
}

interface RegressionFit {
  /** Coefficients m_1..m_k (dropped collinear variables are 0). */
  slopes: number[];
  intercept: number;
  n: number;
  k: number;
  /** Variables actually estimated (collinear ones dropped). */
  used: number;
  df: number;
  ssReg: number;
  ssResid: number;
  /** Standard errors of m_1..m_k and b (NaN when undefined). */
  slopeErrors: number[];
  interceptError: number;
  constant: boolean;
}

/** Least squares via Householder QR, dropping collinear columns like Excel's LINEST. */
function fitRegression(data: RegressionData, constant: boolean): RegressionFit | EvaluationError {
  const { y, x, k } = data;
  const n = y.length;
  if (n === 0) return valueError();
  const xMeans = new Array<number>(k).fill(0);
  let yMean = 0;
  if (constant) {
    for (let row = 0; row < n; row += 1) {
      yMean += y[row];
      for (let column = 0; column < k; column += 1) xMeans[column] += x[row][column];
    }
    yMean /= n;
    for (let column = 0; column < k; column += 1) xMeans[column] /= n;
  }
  // Column-major working copy of the (centered) design matrix.
  const a: Float64Array[] = [];
  for (let column = 0; column < k; column += 1) {
    const values = new Float64Array(n);
    for (let row = 0; row < n; row += 1) values[row] = x[row][column] - xMeans[column];
    a.push(values);
  }
  const b = Float64Array.from(y, (value) => value - yMean);
  const accepted: number[] = [];
  const diagonal: number[] = [];
  const maxRank = Math.max(0, n - (constant ? 1 : 0));
  for (let column = 0; column < k; column += 1) {
    const r = accepted.length;
    const values = a[column];
    let originalNorm = 0;
    for (let row = 0; row < n; row += 1) originalNorm += (x[row][column] - xMeans[column]) ** 2;
    let norm = 0;
    for (let row = r; row < n; row += 1) norm += values[row] * values[row];
    norm = Math.sqrt(norm);
    if (r >= maxRank || norm <= 1e-10 * Math.sqrt(originalNorm) || norm === 0) continue;
    const alpha = values[r] > 0 ? -norm : norm;
    const v = new Float64Array(n);
    for (let row = r; row < n; row += 1) v[row] = values[row];
    v[r] -= alpha;
    let vNorm = 0;
    for (let row = r; row < n; row += 1) vNorm += v[row] * v[row];
    if (vNorm === 0) continue;
    const reflect = (target: Float64Array) => {
      let dot = 0;
      for (let row = r; row < n; row += 1) dot += v[row] * target[row];
      const factor = (2 * dot) / vNorm;
      for (let row = r; row < n; row += 1) target[row] -= factor * v[row];
    };
    for (let other = column; other < k; other += 1) reflect(a[other]);
    reflect(b);
    accepted.push(column);
    diagonal.push(a[column][r]);
  }
  const used = accepted.length;
  // Back substitution: R beta = (Q^T b)[0..used).
  const beta = new Array<number>(used).fill(0);
  for (let i = used - 1; i >= 0; i -= 1) {
    let sum = b[i];
    for (let j = i + 1; j < used; j += 1) sum -= a[accepted[j]][i] * beta[j];
    beta[i] = sum / diagonal[i];
  }
  let ssReg = 0;
  for (let i = 0; i < used; i += 1) ssReg += b[i] * b[i];
  let ssResid = 0;
  for (let i = used; i < n; i += 1) ssResid += b[i] * b[i];
  // R^-1 (upper triangular) for coefficient standard errors.
  const inverse: number[][] = Array.from({ length: used }, () => new Array<number>(used).fill(0));
  for (let j = 0; j < used; j += 1) {
    inverse[j][j] = 1 / diagonal[j];
    for (let i = j - 1; i >= 0; i -= 1) {
      let sum = 0;
      for (let m = i + 1; m <= j; m += 1) sum += a[accepted[m]][i] * inverse[m][j];
      inverse[i][j] = -sum / diagonal[i];
    }
  }
  const df = n - used - (constant ? 1 : 0);
  const variance = df > 0 ? ssResid / df : NaN;
  const slopes = new Array<number>(k).fill(0);
  const slopeErrors = new Array<number>(k).fill(0);
  for (let i = 0; i < used; i += 1) {
    slopes[accepted[i]] = beta[i];
    let rowNorm = 0;
    for (let j = i; j < used; j += 1) rowNorm += inverse[i][j] ** 2;
    slopeErrors[accepted[i]] = Math.sqrt(variance * rowNorm);
  }
  let intercept = 0;
  let interceptError = NaN;
  if (constant) {
    intercept = yMean;
    for (let i = 0; i < used; i += 1) intercept -= beta[i] * xMeans[accepted[i]];
    // Var(b) = s² (1/n + x̄ᵀ (XcᵀXc)⁻¹ x̄) with (XcᵀXc)⁻¹ = R⁻¹ R⁻ᵀ.
    let quadratic = 0;
    for (let j = 0; j < used; j += 1) {
      let sum = 0;
      for (let i = 0; i <= j; i += 1) sum += inverse[i][j] * xMeans[accepted[i]];
      quadratic += sum * sum;
    }
    interceptError = Math.sqrt(variance * (1 / n + quadratic));
  }
  return {
    slopes,
    intercept,
    n,
    k,
    used,
    df,
    ssReg,
    ssResid,
    slopeErrors,
    interceptError,
    constant,
  };
}

function regressionOutput(fit: RegressionFit, stats: boolean, transform: (value: number) => number): Value {
  const columns = fit.k + 1;
  const firstRow: Scalar[] = [];
  for (let index = fit.k - 1; index >= 0; index -= 1) firstRow.push(finiteOrNum(transform(fit.slopes[index])));
  firstRow.push(finiteOrNum(transform(fit.intercept)));
  if (!stats) return arrayResult(firstRow, 1, columns);
  const na = naError();
  const fill = (row: Scalar[]): Scalar[] => {
    while (row.length < columns) row.push(na);
    return row;
  };
  const errorsRow: Scalar[] = [];
  for (let index = fit.k - 1; index >= 0; index -= 1) errorsRow.push(finiteOrNum(fit.slopeErrors[index]));
  errorsRow.push(fit.constant ? finiteOrNum(fit.interceptError) : na);
  const ssTotal = fit.ssReg + fit.ssResid;
  const rSquared = ssTotal === 0 ? numError() : finiteOrNum(fit.ssReg / ssTotal);
  const standardErrorY = fit.df > 0 ? finiteOrNum(Math.sqrt(fit.ssResid / fit.df)) : numError();
  const fStatistic =
    fit.df > 0 && fit.used > 0 && fit.ssResid > 0 ? finiteOrNum(fit.ssReg / fit.used / (fit.ssResid / fit.df)) : numError();
  const values: Scalar[] = [
    ...firstRow,
    ...errorsRow,
    ...fill([rSquared, standardErrorY]),
    ...fill([fStatistic, fit.df]),
    ...fill([finiteOrNum(fit.ssReg), finiteOrNum(fit.ssResid)]),
  ];
  return arrayResult(values, 5, columns);
}

function finiteOrNum(value: number): Scalar {
  return Number.isFinite(value) ? (Object.is(value, -0) ? 0 : value) : numError();
}

function linestSpec(logarithmic: boolean) {
  return spec(
    1,
    4,
    (values) => {
      const data = regressionData(values[0], values[1]);
      if (isEvaluationError(data)) return data;
      if (logarithmic) {
        if (data.y.some((value) => value <= 0)) return numError();
        data.y = data.y.map(Math.log);
      }
      const constant = optionalBooleanArg(values[2], true);
      if (isEvaluationError(constant)) return constant;
      const stats = optionalBooleanArg(values[3], false);
      if (isEvaluationError(stats)) return stats;
      const fit = fitRegression(data, constant);
      if (isEvaluationError(fit)) return fit;
      return regressionOutput(fit, stats, logarithmic ? Math.exp : (value) => value);
    },
    { returnsArray: true },
  );
}

function trendSpec(exponential: boolean) {
  return spec(
    1,
    4,
    (values) => {
      const data = regressionData(values[0], values[1]);
      if (isEvaluationError(data)) return data;
      if (exponential) {
        if (data.y.some((value) => value <= 0)) return numError();
        data.y = data.y.map(Math.log);
      }
      const constant = optionalBooleanArg(values[3], true);
      if (isEvaluationError(constant)) return constant;
      const fit = fitRegression(data, constant);
      if (isEvaluationError(fit)) return fit;
      const output = (prediction: number): Scalar => finiteOrNum(exponential ? Math.exp(prediction) : prediction);
      let newX: { values: number[]; rowCount: number; columnCount: number };
      if (isMissing(values[2])) {
        if (isMissing(values[1])) {
          newX = {
            values: data.y.map((_value, index) => index + 1),
            rowCount: data.yRows,
            columnCount: data.yColumns,
          };
        } else {
          const knownX = numericGrid(values[1]);
          if (isEvaluationError(knownX)) return knownX;
          newX = knownX;
        }
      } else {
        const parsed = numericGrid(values[2]);
        if (isEvaluationError(parsed)) return parsed;
        newX = parsed;
      }
      if (data.k === 1) {
        return arrayResult(
          newX.values.map((value) => output(fit.intercept + fit.slopes[0] * value)),
          newX.rowCount,
          newX.columnCount,
        );
      }
      if (!data.byRow) {
        if (newX.columnCount !== data.k) return refError();
        const predictions: Scalar[] = [];
        for (let row = 0; row < newX.rowCount; row += 1) {
          let sum = fit.intercept;
          for (let column = 0; column < data.k; column += 1) sum += fit.slopes[column] * newX.values[row * data.k + column];
          predictions.push(output(sum));
        }
        return arrayResult(predictions, newX.rowCount, 1);
      }
      if (newX.rowCount !== data.k) return refError();
      const predictions: Scalar[] = [];
      for (let column = 0; column < newX.columnCount; column += 1) {
        let sum = fit.intercept;
        for (let variable = 0; variable < data.k; variable += 1) {
          sum += fit.slopes[variable] * newX.values[variable * newX.columnCount + column];
        }
        predictions.push(output(sum));
      }
      return arrayResult(predictions, 1, newX.columnCount);
    },
    { returnsArray: true },
  );
}

function forecastSpec() {
  return spec(
    3,
    3,
    (values) => {
      const x = numberArg(values[0]);
      if (isEvaluationError(x)) return x;
      const moments = regressionMoments(values[1], values[2]);
      if (isEvaluationError(moments)) return moments;
      if (moments.n === 0 || moments.sxx === 0) return divError();
      const slope = moments.sxy / moments.sxx;
      return finiteResult(moments.meanY - slope * moments.meanX + slope * x);
    },
    { liftArgs: [0] },
  );
}

// ---- Hypothesis tests --------------------------------------------------------------------------

function tTest(values: Value[]): NumberResult {
  const tails = integerArg(values[2]);
  if (isEvaluationError(tails)) return tails;
  const type = integerArg(values[3]);
  if (isEvaluationError(type)) return type;
  if (tails !== 1 && tails !== 2) return numError();
  if (type < 1 || type > 3) return numError();
  let t: number;
  let df: number;
  if (type === 1) {
    const pairs = numericPairs(values[0], values[1]);
    if (isEvaluationError(pairs)) return pairs;
    const differences = pairs.xs.map((value, index) => value - pairs.ys[index]);
    const n = differences.length;
    if (n < 2) return divError();
    const variance = sumSquaredDeviations(differences) / (n - 1);
    if (variance === 0) return divError();
    t = mean(differences) / Math.sqrt(variance / n);
    df = n - 1;
  } else {
    const first = readNumbers([values[0]]);
    if (isEvaluationError(first)) return first;
    const second = readNumbers([values[1]]);
    if (isEvaluationError(second)) return second;
    const n1 = first.length;
    const n2 = second.length;
    if (n1 < 2 || n2 < 2) return divError();
    const v1 = sumSquaredDeviations(first) / (n1 - 1);
    const v2 = sumSquaredDeviations(second) / (n2 - 1);
    const difference = mean(first) - mean(second);
    if (type === 2) {
      df = n1 + n2 - 2;
      const pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / df;
      const denominator = Math.sqrt(pooled * (1 / n1 + 1 / n2));
      if (denominator === 0) return divError();
      t = difference / denominator;
    } else {
      const a = v1 / n1;
      const b = v2 / n2;
      if (a + b === 0) return divError();
      t = difference / Math.sqrt(a + b);
      df = (a + b) ** 2 / (a * a / (n1 - 1) + b * b / (n2 - 1));
    }
  }
  return finiteResult(tails * S.studentTSf(Math.abs(t), df));
}

function fTest(values: Value[]): NumberResult {
  const first = readNumbers([values[0]]);
  if (isEvaluationError(first)) return first;
  const second = readNumbers([values[1]]);
  if (isEvaluationError(second)) return second;
  if (first.length < 2 || second.length < 2) return divError();
  const v1 = sumSquaredDeviations(first) / (first.length - 1);
  const v2 = sumSquaredDeviations(second) / (second.length - 1);
  if (v1 === 0 || v2 === 0) return divError();
  const upper = S.fDistribution(first.length - 1, second.length - 1).sf(v1 / v2);
  return finiteResult(2 * Math.min(upper, 1 - upper));
}

function chiSquareTest(values: Value[]): NumberResult {
  const actual = rectArg(values[0]);
  if (isEvaluationError(actual)) return actual;
  const expected = rectArg(values[1]);
  if (isEvaluationError(expected)) return expected;
  if (actual.rowCount !== expected.rowCount || actual.columnCount !== expected.columnCount) return naError();
  const rows = actual.rowCount;
  const columns = actual.columnCount;
  let df: number;
  if (rows > 1 && columns > 1) df = (rows - 1) * (columns - 1);
  else if (rows * columns > 1) df = rows * columns - 1;
  else return naError();
  let statistic = 0;
  for (let index = 0; index < actual.values.length; index += 1) {
    const observed = actual.values[index];
    const expectation = expected.values[index];
    if (isEvaluationError(observed)) return observed;
    if (isEvaluationError(expectation)) return expectation;
    if (typeof observed !== "number" || typeof expectation !== "number") continue;
    if (expectation === 0) return divError();
    statistic += (observed - expectation) ** 2 / expectation;
  }
  return finiteResult(S.gammaQ(df / 2, statistic / 2));
}

function zTest(values: Value[]): NumberResult {
  const numbers = readNumbers([values[0]]);
  if (isEvaluationError(numbers)) return numbers;
  const x = numberArg(values[1]);
  if (isEvaluationError(x)) return x;
  const n = numbers.length;
  if (n === 0) return naError();
  let sigma: number;
  if (isMissing(values[2])) {
    const variance = sampleVariance(numbers);
    if (isEvaluationError(variance)) return variance;
    sigma = Math.sqrt(variance);
  } else {
    const given = numberArg(values[2]);
    if (isEvaluationError(given)) return given;
    sigma = given;
  }
  if (sigma <= 0) return sigma === 0 ? divError() : numError();
  return finiteResult(1 - S.normalCdf((mean(numbers) - x) / (sigma / Math.sqrt(n))));
}

// ---- Discrete distribution helpers -------------------------------------------------------------

function binomialInverse(trials: number, p: number, alpha: number): NumberResult {
  if (trials < 0 || p < 0 || p > 1 || alpha < 0 || alpha > 1) return numError();
  const target = alpha * (1 - 1e-14);
  if (trials <= 10_000) {
    let cumulative = 0;
    for (let k = 0; k <= trials; k += 1) {
      cumulative += S.binomialPmf(k, trials, p);
      if (cumulative >= target) return k;
    }
    return trials;
  }
  let lo = 0;
  let hi = trials;
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2);
    if (S.binomialCdf(mid, trials, p) >= target) hi = mid;
    else lo = mid + 1;
  }
  return lo;
}

function hypergeometric(
  sample: number,
  sampleSize: number,
  successes: number,
  population: number,
  cumulative: boolean,
): NumberResult {
  if (
    sample < 0 ||
    sample > Math.min(sampleSize, successes) ||
    sample < Math.max(0, sampleSize - population + successes) ||
    sampleSize <= 0 ||
    sampleSize > population ||
    successes <= 0 ||
    successes > population ||
    population <= 0
  ) {
    return numError();
  }
  const denominator = S.lnCombination(population, sampleSize);
  const pmf = (k: number) =>
    Math.exp(S.lnCombination(successes, k) + S.lnCombination(population - successes, sampleSize - k) - denominator);
  if (!cumulative) return finiteResult(pmf(sample));
  let sum = 0;
  for (let k = Math.max(0, sampleSize - population + successes); k <= sample; k += 1) sum += pmf(k);
  return finiteResult(Math.min(1, sum));
}

function negativeBinomialPmf(failures: number, successes: number, p: number): number {
  if (p === 0) return 0;
  if (p === 1) return failures === 0 ? 1 : 0;
  return Math.exp(
    S.lnCombination(failures + successes - 1, successes - 1) + successes * Math.log(p) + failures * Math.log1p(-p),
  );
}

// ---- Distribution specs --------------------------------------------------------------------------

const normalDist = numericFunction([undefined, undefined, undefined, undefined], (x, m, s, c) => {
  if (s <= 0) return numError();
  const z = (x - m) / s;
  return c ? S.normalCdf(z) : S.normalPdf(z) / s;
});
const normalInv = numericFunction([undefined, undefined, undefined], (p, m, s) =>
  p <= 0 || p >= 1 || s <= 0 ? numError() : m + s * S.normalInv(p),
);
const standardNormalInv = numericFunction([undefined], (p) => (p <= 0 || p >= 1 ? numError() : S.normalInv(p)));
const logNormalInv = numericFunction([undefined, undefined, undefined], (p, m, s) =>
  p <= 0 || p >= 1 || s <= 0 ? numError() : Math.exp(m + s * S.normalInv(p)),
);

const tDist = numericFunction([undefined, undefined, undefined], (x, rawDf, c) => {
  const df = truncate(rawDf);
  if (df < 1) return numError();
  return c ? S.studentTCdf(x, df) : S.studentTPdf(x, df);
});
const tDistTwoTailed = numericFunction([undefined, undefined], (x, rawDf) => {
  const df = truncate(rawDf);
  if (x < 0 || df < 1) return numError();
  return Math.min(1, 2 * S.studentTSf(x, df));
});
const tDistRightTailed = numericFunction([undefined, undefined], (x, rawDf) => {
  const df = truncate(rawDf);
  if (df < 1) return numError();
  return S.studentTSf(x, df);
});
const tDistLegacy = numericFunction([undefined, undefined, undefined], (x, rawDf, rawTails) => {
  const df = truncate(rawDf);
  const tails = truncate(rawTails);
  if (x < 0 || df < 1 || (tails !== 1 && tails !== 2)) return numError();
  return Math.min(1, tails * S.studentTSf(x, df));
});
const tInv = numericFunction([undefined, undefined], (p, rawDf) => {
  const df = truncate(rawDf);
  if (p <= 0 || p >= 1 || df < 1) return numError();
  if (p === 0.5) return 0;
  return p < 0.5 ? -S.studentTInvUpper(p, df) : S.studentTInvUpper(1 - p, df);
});
const tInvTwoTailed = numericFunction([undefined, undefined], (p, rawDf) => {
  const df = truncate(rawDf);
  if (p <= 0 || p > 1 || df < 1) return numError();
  return S.studentTInvUpper(p / 2, df);
});

function chiSquareDf(raw: number): number | null {
  const df = truncate(raw);
  return df < 1 || df > 1e10 ? null : df;
}
const chiDist = numericFunction([undefined, undefined, undefined], (x, rawDf, c) => {
  const df = chiSquareDf(rawDf);
  if (df === null || x < 0) return numError();
  return c ? S.gammaP(df / 2, x / 2) : S.chiSquarePdf(x, df);
});
const chiDistRight = numericFunction([undefined, undefined], (x, rawDf) => {
  const df = chiSquareDf(rawDf);
  if (df === null || x < 0) return numError();
  return S.gammaQ(df / 2, x / 2);
});
const chiInv = numericFunction([undefined, undefined], (p, rawDf) => {
  const df = chiSquareDf(rawDf);
  if (df === null || p < 0 || p >= 1) return numError();
  return 2 * S.gammaInv(p, 1 - p, df / 2);
});
const chiInvRight = numericFunction([undefined, undefined], (q, rawDf) => {
  const df = chiSquareDf(rawDf);
  if (df === null || q <= 0 || q > 1) return numError();
  return 2 * S.gammaInv(1 - q, q, df / 2);
});

function fDegrees(raw1: number, raw2: number): [number, number] | null {
  const d1 = truncate(raw1);
  const d2 = truncate(raw2);
  return d1 < 1 || d2 < 1 || d1 >= 1e10 || d2 >= 1e10 ? null : [d1, d2];
}
const fDist = numericFunction([undefined, undefined, undefined, undefined], (x, raw1, raw2, c) => {
  const degrees = fDegrees(raw1, raw2);
  if (!degrees || x < 0) return numError();
  return c ? S.fDistribution(degrees[0], degrees[1]).cdf(x) : S.fPdf(x, degrees[0], degrees[1]);
});
const fDistRight = numericFunction([undefined, undefined, undefined], (x, raw1, raw2) => {
  const degrees = fDegrees(raw1, raw2);
  if (!degrees || x < 0) return numError();
  return S.fDistribution(degrees[0], degrees[1]).sf(x);
});
const fInv = numericFunction([undefined, undefined, undefined], (p, raw1, raw2) => {
  const degrees = fDegrees(raw1, raw2);
  if (!degrees || p < 0 || p >= 1) return numError();
  return S.fInv(p, 1 - p, degrees[0], degrees[1]);
});
const fInvRight = numericFunction([undefined, undefined, undefined], (q, raw1, raw2) => {
  const degrees = fDegrees(raw1, raw2);
  if (!degrees || q <= 0 || q > 1) return numError();
  return S.fInv(1 - q, q, degrees[0], degrees[1]);
});

function betaDistribution(x: number, alpha: number, beta: number, cumulative: boolean, low: number, high: number): NumberResult {
  if (alpha <= 0 || beta <= 0 || low >= high || x < low || x > high) return numError();
  const z = (x - low) / (high - low);
  return cumulative ? S.betaRegularized(z, alpha, beta) : S.betaPdf(z, alpha, beta) / (high - low);
}
const betaDist = numericFunction([undefined, undefined, undefined, undefined, 0, 1], (x, a, b, c, low, high) =>
  betaDistribution(x, a, b, c !== 0, low, high),
);
const betaDistLegacy = numericFunction([undefined, undefined, undefined, 0, 1], (x, a, b, low, high) =>
  betaDistribution(x, a, b, true, low, high),
);
const betaInv = numericFunction([undefined, undefined, undefined, 0, 1], (p, a, b, low, high) => {
  if (p <= 0 || p > 1 || a <= 0 || b <= 0 || low >= high) return numError();
  return low + (high - low) * S.betaInv(p, 1 - p, a, b);
});

const gammaDist = numericFunction([undefined, undefined, undefined, undefined], (x, alpha, beta, c) => {
  if (x < 0 || alpha <= 0 || beta <= 0) return numError();
  return c ? S.gammaP(alpha, x / beta) : S.gammaPdf(x, alpha, beta);
});
const gammaInv = numericFunction([undefined, undefined, undefined], (p, alpha, beta) => {
  if (p < 0 || p >= 1 || alpha <= 0 || beta <= 0) return numError();
  return S.gammaInv(p, 1 - p, alpha, beta);
});

const exponentialDist = numericFunction([undefined, undefined, undefined], (x, lambda, c) => {
  if (x < 0 || lambda <= 0) return numError();
  return c ? -Math.expm1(-lambda * x) : lambda * Math.exp(-lambda * x);
});
const weibullDist = numericFunction([undefined, undefined, undefined, undefined], (x, alpha, beta, c) => {
  if (x < 0 || alpha <= 0 || beta <= 0) return numError();
  const scaled = (x / beta) ** alpha;
  if (c) return -Math.expm1(-scaled);
  if (x === 0) return alpha < 1 ? numError() : alpha === 1 ? 1 / beta : 0;
  return (alpha / beta) * (x / beta) ** (alpha - 1) * Math.exp(-scaled);
});
const poissonDist = numericFunction([undefined, undefined, undefined], (rawX, mean, c) => {
  const x = truncate(rawX);
  if (x < 0 || mean < 0) return numError();
  return c ? S.poissonCdf(x, mean) : S.poissonPmf(x, mean);
});
const binomialDist = numericFunction([undefined, undefined, undefined, undefined], (rawS, rawN, p, c) => {
  const successes = truncate(rawS);
  const trials = truncate(rawN);
  if (successes < 0 || successes > trials || p < 0 || p > 1) return numError();
  return c ? S.binomialCdf(successes, trials, p) : S.binomialPmf(successes, trials, p);
});
const binomialRange = numericFunction([undefined, undefined, undefined, NaN], (rawN, p, rawS, rawS2) => {
  const trials = truncate(rawN);
  const first = truncate(rawS);
  const last = Number.isNaN(rawS2) ? first : truncate(rawS2);
  if (trials < 0 || p < 0 || p > 1 || first < 0 || first > trials || last < first || last > trials) return numError();
  if (last - first <= 1_000) {
    let sum = 0;
    for (let k = first; k <= last; k += 1) sum += S.binomialPmf(k, trials, p);
    return Math.min(1, sum);
  }
  return Math.max(0, S.binomialCdf(last, trials, p) - S.binomialCdf(first - 1, trials, p));
});
const binomialInv = numericFunction([undefined, undefined, undefined], (rawN, p, alpha) =>
  binomialInverse(truncate(rawN), p, alpha),
);
const negativeBinomialDist = numericFunction([undefined, undefined, undefined, undefined], (rawF, rawS, p, c) => {
  const failures = truncate(rawF);
  const successes = truncate(rawS);
  if (p < 0 || p > 1 || failures < 0 || successes < 1) return numError();
  if (!c) return negativeBinomialPmf(failures, successes, p);
  if (p === 0) return 0;
  if (p === 1) return 1;
  return S.betaRegularized(p, successes, failures + 1, 1 - p);
});
const negativeBinomialLegacy = numericFunction([undefined, undefined, undefined], (rawF, rawS, p) => {
  const failures = truncate(rawF);
  const successes = truncate(rawS);
  if (p < 0 || p > 1 || failures < 0 || successes < 1) return numError();
  return negativeBinomialPmf(failures, successes, p);
});
const hypergeometricDist = numericFunction([undefined, undefined, undefined, undefined, undefined], (k, n, m, N, c) =>
  hypergeometric(truncate(k), truncate(n), truncate(m), truncate(N), c !== 0),
);
const hypergeometricLegacy = numericFunction([undefined, undefined, undefined, undefined], (k, n, m, N) =>
  hypergeometric(truncate(k), truncate(n), truncate(m), truncate(N), false),
);

const confidenceNorm = numericFunction([undefined, undefined, undefined], (alpha, sd, rawSize) => {
  const size = truncate(rawSize);
  if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) return numError();
  return (-S.normalInv(alpha / 2) * sd) / Math.sqrt(size);
});
const confidenceT = numericFunction([undefined, undefined, undefined], (alpha, sd, rawSize) => {
  const size = truncate(rawSize);
  if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) return numError();
  if (size === 1) return divError();
  return (S.studentTInvUpper(alpha / 2, size - 1) * sd) / Math.sqrt(size);
});

const erfSpec = numericFunction([undefined, NaN], (lower, upper) =>
  Number.isNaN(upper) ? S.erf(lower) : S.erf(upper) - S.erf(lower),
);
const erfPrecise = numericFunction([undefined], (x) => S.erf(x));
const erfcSpec = numericFunction([undefined], (x) => S.erfc(x));
const gammaLn = numericFunction([undefined], (x) => (x <= 0 ? numError() : S.lnGamma(x)));

// ---- Specs -------------------------------------------------------------------------------------

function varianceSpec(population: boolean, root: boolean) {
  return aggregate((numbers) => {
    const variance = population ? populationVariance(numbers) : sampleVariance(numbers);
    if (isEvaluationError(variance)) return variance;
    return root ? Math.sqrt(variance) : variance;
  }, true);
}

const percentileExc = spec(
  2,
  2,
  (values) => {
    const numbers = readNumbers([values[0]]);
    if (isEvaluationError(numbers)) return numbers;
    const k = numberArg(values[1]);
    if (isEvaluationError(k)) return k;
    return percentileExclusive(sortedNumbers(numbers), k);
  },
  { liftArgs: [1] },
);

const quartileExc = spec(
  2,
  2,
  (values) => {
    const numbers = readNumbers([values[0]]);
    if (isEvaluationError(numbers)) return numbers;
    const quart = integerArg(values[1]);
    if (isEvaluationError(quart)) return quart;
    if (quart <= 0 || quart >= 4) return numError();
    return percentileExclusive(sortedNumbers(numbers), quart / 4);
  },
  { liftArgs: [1] },
);

const rankAverage = spec(
  2,
  3,
  (values) => {
    const number = numberArg(values[0]);
    if (isEvaluationError(number)) return number;
    const numbers = readNumbers([values[1]]);
    if (isEvaluationError(numbers)) return numbers;
    const order = optionalIntegerArg(values[2], 0);
    if (isEvaluationError(order)) return order;
    let before = 0;
    let equal = 0;
    for (const value of numbers) {
      if (value === number) equal += 1;
      else if (order === 0 ? value > number : value < number) before += 1;
    }
    if (equal === 0) return naError();
    return before + (equal + 1) / 2;
  },
  { liftArgs: [0] },
);

const modeMultiple = spec(
  1,
  255,
  (values) => {
    const numbers = readNumbers(values);
    if (isEvaluationError(numbers)) return numbers;
    const counts = new Map<number, number>();
    let best = 1;
    for (const value of numbers) {
      const count = (counts.get(value) ?? 0) + 1;
      counts.set(value, count);
      if (count > best) best = count;
    }
    if (best < 2) return naError();
    const modes: Scalar[] = [];
    for (const [value, count] of counts) if (count === best) modes.push(value);
    return arrayResult(modes, modes.length, 1);
  },
  { returnsArray: true },
);

const frequency = spec(
  2,
  2,
  (values) => {
    const data = readNumbers([values[0]]);
    if (isEvaluationError(data)) return data;
    const bins = readNumbers([values[1]]);
    if (isEvaluationError(bins)) return bins;
    const order = bins.map((_bin, index) => index).sort((a, b) => bins[a] - bins[b] || a - b);
    const sortedBins = order.map((index) => bins[index]);
    const counts = new Array<number>(bins.length + 1).fill(0);
    for (const value of data) {
      // First bin (in ascending order) that is >= value.
      let lo = 0;
      let hi = sortedBins.length;
      while (lo < hi) {
        const mid = (lo + hi) >>> 1;
        if (sortedBins[mid] < value) lo = mid + 1;
        else hi = mid;
      }
      counts[lo] += 1;
    }
    const result: Scalar[] = new Array(bins.length + 1).fill(0);
    order.forEach((original, sortedIndex) => {
      result[original] = counts[sortedIndex];
    });
    result[bins.length] = counts[bins.length];
    return arrayResult(result, result.length, 1);
  },
  { returnsArray: true },
);

const prob = spec(
  3,
  4,
  (values) => {
    const pairs = numericPairs(values[0], values[1]);
    if (isEvaluationError(pairs)) return pairs;
    const lower = numberArg(values[2]);
    if (isEvaluationError(lower)) return lower;
    let upper = lower;
    if (!isMissing(values[3])) {
      const given = numberArg(values[3]);
      if (isEvaluationError(given)) return given;
      upper = given;
    }
    let total = 0;
    let sum = 0;
    for (let index = 0; index < pairs.xs.length; index += 1) {
      const probability = pairs.ys[index];
      if (probability <= 0 || probability > 1) return numError();
      total += probability;
      const x = pairs.xs[index];
      if (x >= lower && x <= upper) sum += probability;
    }
    if (Math.abs(total - 1) > 1e-9) return numError();
    return finiteResult(sum);
  },
  { liftArgs: [2, 3] },
);

const averageWeighted = spec(2, 254, (values) => {
  if (values.length % 2 !== 0) return valueError();
  let weighted = 0;
  let totalWeight = 0;
  for (let index = 0; index < values.length; index += 2) {
    const valueRect = rectArg(values[index]);
    if (isEvaluationError(valueRect)) return valueRect;
    const weightRect = rectArg(values[index + 1]);
    if (isEvaluationError(weightRect)) return weightRect;
    if (valueRect.values.length !== weightRect.values.length) return valueError();
    for (let position = 0; position < valueRect.values.length; position += 1) {
      const value = valueRect.values[position];
      const weight = weightRect.values[position];
      if (isEvaluationError(value)) return value;
      if (isEvaluationError(weight)) return weight;
      if (value === null && weight === null) continue;
      const number = typeof value === "number" ? value : toNumber(scalarArgument(value as EvaluationScalar));
      const factor = typeof weight === "number" ? weight : toNumber(scalarArgument(weight as EvaluationScalar));
      if (isEvaluationError(number)) return number;
      if (isEvaluationError(factor)) return factor;
      if (factor < 0) return numError();
      weighted += number * factor;
      totalWeight += factor;
    }
  }
  if (totalWeight === 0) return divError();
  return finiteResult(weighted / totalWeight);
});

export const STATS_FUNCTIONS: Specs = {
  ERF: erfSpec,
  "ERF.PRECISE": erfPrecise,
  ERFC: erfcSpec,
  "ERFC.PRECISE": erfcSpec,

  MAXA: spec(0, 255, (values) => {
    const numbers = readNumbersA(values);
    if (isEvaluationError(numbers)) return numbers;
    let best = -Infinity;
    for (const value of numbers) if (value > best) best = value;
    return numbers.length ? best : 0;
  }),
  MINA: spec(0, 255, (values) => {
    const numbers = readNumbersA(values);
    if (isEvaluationError(numbers)) return numbers;
    let best = Infinity;
    for (const value of numbers) if (value < best) best = value;
    return numbers.length ? best : 0;
  }),
  STDEVA: varianceSpec(false, true),
  STDEVPA: varianceSpec(true, true),
  VARA: varianceSpec(false, false),
  VARPA: varianceSpec(true, false),
  GEOMEAN: aggregate((numbers) => {
    if (numbers.length === 0) return numError();
    let logSum = 0;
    for (const value of numbers) {
      if (value <= 0) return numError();
      logSum += Math.log(value);
    }
    return finiteResult(Math.exp(logSum / numbers.length));
  }),
  HARMEAN: aggregate((numbers) => {
    if (numbers.length === 0) return numError();
    let inverseSum = 0;
    for (const value of numbers) {
      if (value <= 0) return numError();
      inverseSum += 1 / value;
    }
    return finiteResult(numbers.length / inverseSum);
  }),
  AVEDEV: aggregate((numbers) => {
    if (numbers.length === 0) return numError();
    const average = mean(numbers);
    let sum = 0;
    for (const value of numbers) sum += Math.abs(value - average);
    return finiteResult(sum / numbers.length);
  }),
  DEVSQ: aggregate((numbers) => (numbers.length === 0 ? numError() : finiteResult(sumSquaredDeviations(numbers)))),
  KURT: aggregate((numbers) => {
    const n = numbers.length;
    if (n < 4) return divError();
    const variance = sumSquaredDeviations(numbers) / (n - 1);
    if (variance === 0) return divError();
    const average = mean(numbers);
    let fourth = 0;
    for (const value of numbers) fourth += ((value - average) ** 2 / variance) ** 2;
    return finiteResult(
      ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * fourth - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3)),
    );
  }),
  SKEW: aggregate((numbers) => {
    const n = numbers.length;
    if (n < 3) return divError();
    const variance = sumSquaredDeviations(numbers) / (n - 1);
    if (variance === 0) return divError();
    const sd = Math.sqrt(variance);
    const average = mean(numbers);
    let third = 0;
    for (const value of numbers) third += ((value - average) / sd) ** 3;
    return finiteResult((n / ((n - 1) * (n - 2))) * third);
  }),
  "SKEW.P": aggregate((numbers) => {
    const n = numbers.length;
    if (n < 3) return divError();
    const variance = sumSquaredDeviations(numbers) / n;
    if (variance === 0) return divError();
    const sd = Math.sqrt(variance);
    const average = mean(numbers);
    let third = 0;
    for (const value of numbers) third += ((value - average) / sd) ** 3;
    return finiteResult(third / n);
  }),
  STANDARDIZE: numericFunction([undefined, undefined, undefined], (x, m, s) => (s <= 0 ? numError() : (x - m) / s)),
  TRIMMEAN: spec(
    2,
    2,
    (values) => {
      const numbers = readNumbers([values[0]]);
      if (isEvaluationError(numbers)) return numbers;
      const percent = numberArg(values[1]);
      if (isEvaluationError(percent)) return percent;
      if (percent < 0 || percent >= 1 || numbers.length === 0) return numError();
      const trim = Math.floor((numbers.length * percent) / 2 + 1e-9);
      const sorted = sortedNumbers(numbers);
      let sum = 0;
      for (let index = trim; index < sorted.length - trim; index += 1) sum += sorted[index];
      return finiteResult(sum / (sorted.length - 2 * trim));
    },
    { liftArgs: [1] },
  ),
  "PERCENTILE.EXC": percentileExc,
  "QUARTILE.EXC": quartileExc,
  PERCENTRANK: percentRankSpec(false),
  "PERCENTRANK.INC": percentRankSpec(false),
  "PERCENTRANK.EXC": percentRankSpec(true),
  "RANK.AVG": rankAverage,
  "MODE.MULT": modeMultiple,
  FREQUENCY: frequency,

  CORREL: pairedSpec(correlation),
  PEARSON: pairedSpec(correlation),
  RSQ: pairedSpec((moments) => {
    const r = correlation(moments);
    return isEvaluationError(r) ? r : r * r;
  }, true),
  COVAR: pairedSpec((moments) => (moments.n === 0 ? divError() : moments.sxy / moments.n)),
  "COVARIANCE.P": pairedSpec((moments) => (moments.n === 0 ? divError() : moments.sxy / moments.n)),
  "COVARIANCE.S": pairedSpec((moments) => (moments.n < 2 ? divError() : moments.sxy / (moments.n - 1))),
  SLOPE: pairedSpec((moments) => (moments.n === 0 || moments.sxx === 0 ? divError() : moments.sxy / moments.sxx), true),
  INTERCEPT: pairedSpec(
    (moments) =>
      moments.n === 0 || moments.sxx === 0 ? divError() : moments.meanY - (moments.sxy / moments.sxx) * moments.meanX,
    true,
  ),
  STEYX: pairedSpec((moments) => {
    if (moments.n < 3 || moments.sxx === 0) return divError();
    const residual = Math.max(0, moments.syy - (moments.sxy * moments.sxy) / moments.sxx);
    return Math.sqrt(residual / (moments.n - 2));
  }, true),
  FORECAST: forecastSpec(),
  "FORECAST.LINEAR": forecastSpec(),
  TREND: trendSpec(false),
  GROWTH: trendSpec(true),
  LINEST: linestSpec(false),
  LOGEST: linestSpec(true),

  "NORM.DIST": normalDist,
  NORMDIST: normalDist,
  "NORM.INV": normalInv,
  NORMINV: normalInv,
  "NORM.S.DIST": numericFunction([undefined, undefined], (z, c) => (c ? S.normalCdf(z) : S.normalPdf(z))),
  NORMSDIST: numericFunction([undefined], (z) => S.normalCdf(z)),
  "NORM.S.INV": standardNormalInv,
  NORMSINV: standardNormalInv,
  "LOGNORM.DIST": numericFunction([undefined, undefined, undefined, undefined], (x, m, s, c) => {
    if (x <= 0 || s <= 0) return numError();
    const z = (Math.log(x) - m) / s;
    return c ? S.normalCdf(z) : S.normalPdf(z) / (x * s);
  }),
  LOGNORMDIST: numericFunction([undefined, undefined, undefined], (x, m, s) => {
    if (x <= 0 || s <= 0) return numError();
    return S.normalCdf((Math.log(x) - m) / s);
  }),
  "LOGNORM.INV": logNormalInv,
  LOGINV: logNormalInv,
  PHI: numericFunction([undefined], (x) => S.normalPdf(x)),
  GAUSS: numericFunction([undefined], (z) => 0.5 * S.erf(z / Math.SQRT2)),
  FISHER: numericFunction([undefined], (x) => (x <= -1 || x >= 1 ? numError() : Math.atanh(x))),
  FISHERINV: numericFunction([undefined], (y) => Math.tanh(y)),

  "T.DIST": tDist,
  "T.DIST.2T": tDistTwoTailed,
  "T.DIST.RT": tDistRightTailed,
  TDIST: tDistLegacy,
  "T.INV": tInv,
  "T.INV.2T": tInvTwoTailed,
  TINV: tInvTwoTailed,
  "T.TEST": spec(4, 4, tTest),
  TTEST: spec(4, 4, tTest),

  "CHISQ.DIST": chiDist,
  "CHISQ.DIST.RT": chiDistRight,
  CHIDIST: chiDistRight,
  "CHISQ.INV": chiInv,
  "CHISQ.INV.RT": chiInvRight,
  CHIINV: chiInvRight,
  "CHISQ.TEST": spec(2, 2, chiSquareTest),
  CHITEST: spec(2, 2, chiSquareTest),

  "F.DIST": fDist,
  "F.DIST.RT": fDistRight,
  FDIST: fDistRight,
  "F.INV": fInv,
  "F.INV.RT": fInvRight,
  FINV: fInvRight,
  "F.TEST": spec(2, 2, fTest),
  FTEST: spec(2, 2, fTest),

  "BINOM.DIST": binomialDist,
  BINOMDIST: binomialDist,
  "BINOM.DIST.RANGE": binomialRange,
  "BINOM.INV": binomialInv,
  CRITBINOM: binomialInv,
  "POISSON.DIST": poissonDist,
  POISSON: poissonDist,
  "EXPON.DIST": exponentialDist,
  EXPONDIST: exponentialDist,
  GAMMA: numericFunction([undefined], (x) => (Number.isInteger(x) && x <= 0 ? numError() : S.gamma(x))),
  GAMMALN: gammaLn,
  "GAMMALN.PRECISE": gammaLn,
  "GAMMA.DIST": gammaDist,
  GAMMADIST: gammaDist,
  "GAMMA.INV": gammaInv,
  GAMMAINV: gammaInv,
  "BETA.DIST": betaDist,
  BETADIST: betaDistLegacy,
  "BETA.INV": betaInv,
  BETAINV: betaInv,
  "WEIBULL.DIST": weibullDist,
  WEIBULL: weibullDist,
  "HYPGEOM.DIST": hypergeometricDist,
  HYPGEOMDIST: hypergeometricLegacy,
  "NEGBINOM.DIST": negativeBinomialDist,
  NEGBINOMDIST: negativeBinomialLegacy,
  "CONFIDENCE.NORM": confidenceNorm,
  CONFIDENCE: confidenceNorm,
  "CONFIDENCE.T": confidenceT,
  "Z.TEST": spec(2, 3, zTest, { liftArgs: [1, 2] }),
  ZTEST: spec(2, 3, zTest, { liftArgs: [1, 2] }),
  PROB: prob,
  PERMUT: numericFunction([undefined, undefined], (rawN, rawK) => {
    const n = truncate(rawN);
    const k = truncate(rawK);
    if (n <= 0 || k < 0 || n < k) return numError();
    if (k <= 1_000) {
      let product = 1;
      for (let index = 0; index < k; index += 1) {
        product *= n - index;
        if (!Number.isFinite(product)) return numError();
      }
      return product;
    }
    return Math.exp(S.lnGamma(n + 1) - S.lnGamma(n - k + 1));
  }),
  PERMUTATIONA: numericFunction([undefined, undefined], (rawN, rawK) => {
    const n = truncate(rawN);
    const k = truncate(rawK);
    if (n < 0 || k < 0) return numError();
    return n ** k;
  }),
  "AVERAGE.WEIGHTED": averageWeighted,
};

