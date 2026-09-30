// Math & trigonometry worksheet functions for the extended library (see formula-library.ts):
// sums of squares, number theory, combinatorics, radix/roman conversions, series, the
// reciprocal/hyperbolic trig family, precise ceiling/floor, matrix functions, and the Google
// Sheets operator functions (ADD, MINUS, ...).

import {
  applyBinaryOperator,
  finiteResult,
  isEvaluationError,
  MAX_RANGE_CELLS,
  scalarArgument,
  significanceRound,
  toNumber,
  toText,
} from "./formulas";
import type { EvaluationError } from "./formulas";
import {
  arrayResult,
  divError,
  integerArg,
  numberArg,
  numError,
  numericFunction,
  numericPairs,
  readNumbers,
  rectArg,
  spec,
  valueError,
} from "./formula-lib-shared";
import type { Rect, Scalar, Specs, Value } from "./formula-lib-shared";

const MAX_SAFE = 2 ** 53;
const TRIG_LIMIT = 2 ** 27;

// ---- helpers -------------------------------------------------------------------------------

function gcdPair(left: number, right: number): number {
  let a = left;
  let b = right;
  while (b > 0) {
    const remainder = a % b;
    a = b;
    b = remainder;
  }
  return a;
}

/** Non-negative integer operands for GCD/LCM/MULTINOMIAL (truncated; negatives are #NUM!). */
function wholeNumbers(values: Value[]): number[] | EvaluationError {
  const numbers = readNumbers(values);
  if (isEvaluationError(numbers)) return numbers;
  const result: number[] = [];
  for (const number of numbers) {
    if (number < 0 || number >= MAX_SAFE) return numError();
    result.push(Math.trunc(number));
  }
  return result;
}

export function combinations(n: number, k: number): number {
  const chosen = Math.min(k, n - k);
  let result = 1;
  for (let index = 1; index <= chosen; index += 1) {
    result = (result * (n - chosen + index)) / index;
    if (!Number.isFinite(result)) return Infinity;
  }
  return result < MAX_SAFE ? Math.round(result) : result;
}

export function factorial(n: number): number {
  let result = 1;
  for (let index = 2; index <= n; index += 1) {
    result *= index;
    if (!Number.isFinite(result)) return Infinity;
  }
  return result;
}

/** Every cell of an array argument as a number (errors propagate, anything else is #VALUE!). */
function numericMatrix(value: Value | undefined): (Rect & { numbers: number[] }) | EvaluationError {
  const range = rectArg(value);
  if (isEvaluationError(range)) return range;
  const numbers: number[] = new Array(range.values.length);
  for (let index = 0; index < range.values.length; index += 1) {
    const entry = range.values[index];
    if (isEvaluationError(entry)) return entry;
    if (typeof entry !== "number") return valueError();
    numbers[index] = entry;
  }
  return { ...range, numbers };
}

function clean(value: number): number {
  return Object.is(value, -0) ? 0 : value;
}

// Matrix inverse by Gauss-Jordan elimination with partial pivoting. Null when singular.
function invertMatrix(size: number, source: number[]): number[] | null {
  const width = size * 2;
  const work = new Float64Array(size * width);
  let scale = 0;
  for (let row = 0; row < size; row += 1) {
    for (let column = 0; column < size; column += 1) {
      const value = source[row * size + column];
      work[row * width + column] = value;
      scale = Math.max(scale, Math.abs(value));
    }
    work[row * width + size + row] = 1;
  }
  if (scale === 0) return null;
  const tolerance = scale * size * 1e-15;
  for (let pivotColumn = 0; pivotColumn < size; pivotColumn += 1) {
    let pivotRow = pivotColumn;
    let best = Math.abs(work[pivotColumn * width + pivotColumn]);
    for (let row = pivotColumn + 1; row < size; row += 1) {
      const candidate = Math.abs(work[row * width + pivotColumn]);
      if (candidate > best) {
        best = candidate;
        pivotRow = row;
      }
    }
    if (!(best > tolerance)) return null;
    if (pivotRow !== pivotColumn) {
      for (let column = 0; column < width; column += 1) {
        const swap = work[pivotRow * width + column];
        work[pivotRow * width + column] = work[pivotColumn * width + column];
        work[pivotColumn * width + column] = swap;
      }
    }
    const pivot = work[pivotColumn * width + pivotColumn];
    for (let column = 0; column < width; column += 1) work[pivotColumn * width + column] /= pivot;
    for (let row = 0; row < size; row += 1) {
      if (row === pivotColumn) continue;
      const factor = work[row * width + pivotColumn];
      if (factor === 0) continue;
      for (let column = 0; column < width; column += 1) {
        work[row * width + column] -= factor * work[pivotColumn * width + column];
      }
    }
  }
  const inverse: number[] = new Array(size * size);
  for (let row = 0; row < size; row += 1) {
    for (let column = 0; column < size; column += 1) {
      inverse[row * size + column] = clean(work[row * width + size + column]);
    }
  }
  return inverse;
}

function determinant(size: number, source: number[]): number {
  const work = Float64Array.from(source);
  let result = 1;
  for (let pivotColumn = 0; pivotColumn < size; pivotColumn += 1) {
    let pivotRow = pivotColumn;
    let best = Math.abs(work[pivotColumn * size + pivotColumn]);
    for (let row = pivotColumn + 1; row < size; row += 1) {
      const candidate = Math.abs(work[row * size + pivotColumn]);
      if (candidate > best) {
        best = candidate;
        pivotRow = row;
      }
    }
    if (best === 0) return 0;
    if (pivotRow !== pivotColumn) {
      for (let column = 0; column < size; column += 1) {
        const swap = work[pivotRow * size + column];
        work[pivotRow * size + column] = work[pivotColumn * size + column];
        work[pivotColumn * size + column] = swap;
      }
      result = -result;
    }
    const pivot = work[pivotColumn * size + pivotColumn];
    result *= pivot;
    for (let row = pivotColumn + 1; row < size; row += 1) {
      const factor = work[row * size + pivotColumn] / pivot;
      if (factor === 0) continue;
      for (let column = pivotColumn; column < size; column += 1) {
        work[row * size + column] -= factor * work[pivotColumn * size + column];
      }
    }
  }
  return result;
}

// ---- Roman numerals ------------------------------------------------------------------------

const ROMAN_CHARACTERS = ["M", "D", "C", "L", "X", "V", "I"];
const ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];

/** Excel's ROMAN with forms 0 (classic) through 4 (simplified), as the analysis add-in does. */
function romanText(number: number, form: number): string {
  let value = number;
  let result = "";
  const maxIndex = ROMAN_VALUES.length - 1;
  for (let step = 0; step <= maxIndex / 2; step += 1) {
    let index = step * 2;
    const digit = Math.floor(value / ROMAN_VALUES[index]);
    if (digit % 5 === 4) {
      const index2 = digit === 4 ? index - 1 : index - 2;
      let steps = 0;
      while (steps < form && index < maxIndex) {
        steps += 1;
        if (ROMAN_VALUES[index2] - ROMAN_VALUES[index + 1] <= value) index += 1;
        else steps = form;
      }
      result += ROMAN_CHARACTERS[index] + ROMAN_CHARACTERS[index2];
      value = value + ROMAN_VALUES[index] - ROMAN_VALUES[index2];
    } else {
      if (digit > 4) result += ROMAN_CHARACTERS[index - 1];
      result += ROMAN_CHARACTERS[index].repeat(digit % 5);
      value %= ROMAN_VALUES[index];
    }
  }
  return result;
}

const ARABIC_VALUES: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };

// ---- specs ---------------------------------------------------------------------------------

function pairSum(combine: (x: number, y: number) => number) {
  return spec(2, 2, (values) => {
    const pairs = numericPairs(values[0], values[1]);
    if (isEvaluationError(pairs)) return pairs;
    let sum = 0;
    for (let index = 0; index < pairs.xs.length; index += 1) sum += combine(pairs.xs[index], pairs.ys[index]);
    return finiteResult(sum);
  });
}

function reciprocalTrig(compute: (x: number) => number, zeroIsDivision: boolean) {
  return numericFunction([undefined], (x) => {
    if (Math.abs(x) >= TRIG_LIMIT) return numError();
    if (zeroIsDivision && x === 0) return divError();
    return compute(x);
  });
}

function preciseRound(direction: "up" | "down") {
  return numericFunction([undefined, 1], (number, significance) => {
    const step = Math.abs(significance);
    if (step === 0 || number === 0) return 0;
    return significanceRound(number, step, direction);
  });
}

function operatorFunction(operator: string) {
  return spec(2, 2, (values) => applyBinaryOperator(operator, scalarArgument(values[0]), scalarArgument(values[1])), {
    liftArgs: "all",
  });
}

export const MATH_FUNCTIONS: Specs = {
  SUMSQ: spec(1, 255, (values) => {
    const numbers = readNumbers(values);
    if (isEvaluationError(numbers)) return numbers;
    let sum = 0;
    for (const number of numbers) sum += number * number;
    return finiteResult(sum);
  }),
  SUMX2MY2: pairSum((x, y) => x * x - y * y),
  SUMX2PY2: pairSum((x, y) => x * x + y * y),
  SUMXMY2: pairSum((x, y) => (x - y) * (x - y)),
  GCD: spec(1, 255, (values) => {
    const numbers = wholeNumbers(values);
    if (isEvaluationError(numbers)) return numbers;
    let result = 0;
    for (const number of numbers) result = gcdPair(result, number);
    return result;
  }),
  LCM: spec(1, 255, (values) => {
    const numbers = wholeNumbers(values);
    if (isEvaluationError(numbers)) return numbers;
    if (numbers.length === 0) return 0;
    let result = 1;
    for (const number of numbers) {
      if (number === 0) return 0;
      result = (result / gcdPair(result, number)) * number;
      if (result >= MAX_SAFE) return numError();
    }
    return result;
  }),
  FACT: numericFunction([undefined], (value) => {
    if (value < 0) return numError();
    return factorial(Math.trunc(value));
  }),
  FACTDOUBLE: numericFunction([undefined], (value) => {
    const n = Math.trunc(value);
    if (n < -1) return numError();
    let result = 1;
    for (let index = n; index > 1; index -= 2) {
      result *= index;
      if (!Number.isFinite(result)) return numError();
    }
    return result;
  }),
  COMBIN: numericFunction([undefined, undefined], (rawN, rawK) => {
    const n = Math.trunc(rawN);
    const k = Math.trunc(rawK);
    if (n < 0 || k < 0 || n < k) return numError();
    return combinations(n, k);
  }),
  COMBINA: numericFunction([undefined, undefined], (rawN, rawK) => {
    const n = Math.trunc(rawN);
    const k = Math.trunc(rawK);
    if (n < 0 || k < 0) return numError();
    if (k === 0) return 1;
    if (n === 0) return 0;
    return combinations(n + k - 1, k);
  }),
  MULTINOMIAL: spec(1, 255, (values) => {
    const numbers = wholeNumbers(values);
    if (isEvaluationError(numbers)) return numbers;
    let result = 1;
    let total = 0;
    for (const number of numbers) {
      for (let index = 1; index <= number; index += 1) {
        total += 1;
        result = (result * total) / index;
        if (!Number.isFinite(result)) return numError();
      }
    }
    return result < MAX_SAFE ? Math.round(result) : result;
  }),
  QUOTIENT: numericFunction([undefined, undefined], (numerator, denominator) =>
    denominator === 0 ? divError() : Math.trunc(numerator / denominator),
  ),
  BASE: numericFunction([undefined, undefined, 0], (rawNumber, rawRadix, rawLength) => {
    const number = Math.trunc(rawNumber);
    const radix = Math.trunc(rawRadix);
    const minLength = Math.trunc(rawLength);
    if (number < 0 || number >= MAX_SAFE) return numError();
    if (radix < 2 || radix > 36) return numError();
    if (minLength < 0 || minLength > 255) return numError();
    return number.toString(radix).toUpperCase().padStart(minLength, "0");
  }),
  DECIMAL: spec(2, 2, (values) => {
    const text = toText(scalarArgument(values[0]));
    if (isEvaluationError(text)) return text;
    const radix = integerArg(values[1]);
    if (isEvaluationError(radix)) return radix;
    if (radix < 2 || radix > 36) return numError();
    if (text.length > 255) return valueError();
    let result = 0;
    for (const character of text.trim()) {
      const digit = parseInt(character, 36);
      if (Number.isNaN(digit) || digit >= radix) return numError();
      result = result * radix + digit;
    }
    return finiteResult(result);
  }, { liftArgs: "all" }),
  ROMAN: spec(1, 2, (values) => {
    const number = numberArg(values[0]);
    if (isEvaluationError(number)) return number;
    let form = 0;
    if (values[1] !== undefined) {
      const raw = scalarArgument(values[1]);
      if (isEvaluationError(raw)) return raw;
      if (typeof raw === "boolean") form = raw ? 0 : 4;
      else {
        const numeric = toNumber(raw);
        if (isEvaluationError(numeric)) return numeric;
        form = Math.trunc(numeric);
      }
    }
    const value = Math.trunc(number);
    if (value < 0 || value > 3999 || form < 0 || form > 4) return valueError();
    return romanText(value, form);
  }, { liftArgs: "all" }),
  ARABIC: spec(1, 1, (values) => {
    const raw = scalarArgument(values[0]);
    if (isEvaluationError(raw)) return raw;
    if (typeof raw === "number" || typeof raw === "boolean") return valueError();
    let text = (raw ?? "").trim().toUpperCase();
    if (text.length > 255) return valueError();
    let sign = 1;
    if (text.startsWith("-")) {
      sign = -1;
      text = text.slice(1);
    }
    let total = 0;
    let largest = 0;
    for (let index = text.length - 1; index >= 0; index -= 1) {
      const value = ARABIC_VALUES[text[index]];
      if (value === undefined) return valueError();
      if (value < largest) total -= value;
      else {
        total += value;
        largest = value;
      }
    }
    return sign * total;
  }, { liftArgs: "all" }),
  SERIESSUM: spec(4, 4, (values) => {
    const x = numberArg(values[0]);
    if (isEvaluationError(x)) return x;
    const n = numberArg(values[1]);
    if (isEvaluationError(n)) return n;
    const m = numberArg(values[2]);
    if (isEvaluationError(m)) return m;
    const coefficients = rectArg(values[3]);
    if (isEvaluationError(coefficients)) return coefficients;
    let sum = 0;
    for (let index = 0; index < coefficients.values.length; index += 1) {
      const coefficient = coefficients.values[index];
      if (isEvaluationError(coefficient)) return coefficient;
      if (typeof coefficient !== "number") return valueError();
      const power = n + index * m;
      if (x === 0 && power < 0) return numError();
      sum += coefficient * x ** power;
    }
    return finiteResult(sum);
  }),
  SEC: reciprocalTrig((x) => 1 / Math.cos(x), false),
  CSC: reciprocalTrig((x) => 1 / Math.sin(x), true),
  COT: reciprocalTrig((x) => 1 / Math.tan(x), true),
  COTH: reciprocalTrig((x) => 1 / Math.tanh(x), true),
  CSCH: reciprocalTrig((x) => 1 / Math.sinh(x), true),
  SECH: reciprocalTrig((x) => 1 / Math.cosh(x), false),
  SINH: numericFunction([undefined], (x) => Math.sinh(x)),
  COSH: numericFunction([undefined], (x) => Math.cosh(x)),
  TANH: numericFunction([undefined], (x) => Math.tanh(x)),
  ASINH: numericFunction([undefined], (x) => Math.asinh(x)),
  ACOSH: numericFunction([undefined], (x) => (x < 1 ? numError() : Math.acosh(x))),
  ATANH: numericFunction([undefined], (x) => (Math.abs(x) >= 1 ? numError() : Math.atanh(x))),
  ACOT: numericFunction([undefined], (x) => Math.PI / 2 - Math.atan(x)),
  ACOTH: numericFunction([undefined], (x) => (Math.abs(x) <= 1 ? numError() : 0.5 * Math.log((x + 1) / (x - 1)))),
  "CEILING.PRECISE": preciseRound("up"),
  "ISO.CEILING": preciseRound("up"),
  "FLOOR.PRECISE": preciseRound("down"),
  MMULT: spec(2, 2, (values) => {
    const left = numericMatrix(values[0]);
    if (isEvaluationError(left)) return left;
    const right = numericMatrix(values[1]);
    if (isEvaluationError(right)) return right;
    if (left.columnCount !== right.rowCount) return valueError();
    const rows = left.rowCount;
    const columns = right.columnCount;
    if (rows > MAX_RANGE_CELLS / columns) return valueError();
    const inner = left.columnCount;
    const output = new Float64Array(rows * columns);
    for (let row = 0; row < rows; row += 1) {
      for (let k = 0; k < inner; k += 1) {
        const factor = left.numbers[row * inner + k];
        if (factor === 0) continue;
        for (let column = 0; column < columns; column += 1) {
          output[row * columns + column] += factor * right.numbers[k * columns + column];
        }
      }
    }
    const result: Scalar[] = new Array(output.length);
    for (let index = 0; index < output.length; index += 1) {
      const value = finiteResult(output[index]);
      if (isEvaluationError(value)) return value;
      result[index] = value;
    }
    return arrayResult(result, rows, columns);
  }, { returnsArray: true }),
  MINVERSE: spec(1, 1, (values) => {
    const matrix = numericMatrix(values[0]);
    if (isEvaluationError(matrix)) return matrix;
    if (matrix.rowCount !== matrix.columnCount) return valueError();
    const inverse = invertMatrix(matrix.rowCount, matrix.numbers);
    if (!inverse || inverse.some((value) => !Number.isFinite(value))) return numError();
    return arrayResult(inverse, matrix.rowCount, matrix.columnCount);
  }, { returnsArray: true }),
  MDETERM: spec(1, 1, (values) => {
    const matrix = numericMatrix(values[0]);
    if (isEvaluationError(matrix)) return matrix;
    if (matrix.rowCount !== matrix.columnCount) return valueError();
    // Excel reports 15 significant digits, so LU round-off (0.9999999999999998) reads as 1.
    const value = determinant(matrix.rowCount, matrix.numbers);
    return finiteResult(value === 0 ? 0 : Number(value.toPrecision(15)));
  }),
  MUNIT: spec(1, 1, (values) => {
    const raw = integerArg(values[0]);
    if (isEvaluationError(raw)) return raw;
    if (raw < 1 || raw > Math.sqrt(MAX_RANGE_CELLS)) return valueError();
    const output: Scalar[] = new Array(raw * raw).fill(0);
    for (let index = 0; index < raw; index += 1) output[index * raw + index] = 1;
    return arrayResult(output, raw, raw);
  }, { returnsArray: true }),

  // Google Sheets operator functions.
  ADD: operatorFunction("+"),
  MINUS: operatorFunction("-"),
  MULTIPLY: operatorFunction("*"),
  DIVIDE: operatorFunction("/"),
  POW: operatorFunction("^"),
  EQ: operatorFunction("="),
  NE: operatorFunction("<>"),
  GT: operatorFunction(">"),
  GTE: operatorFunction(">="),
  LT: operatorFunction("<"),
  LTE: operatorFunction("<="),
  UMINUS: numericFunction([undefined], (x) => -x),
  UPLUS: spec(1, 1, (values) => scalarArgument(values[0]), { liftArgs: "all" }),
  UNARY_PERCENT: numericFunction([undefined], (x) => x / 100),
};
