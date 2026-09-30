// Shared argument/result helpers for the extended worksheet function library
// (formula-library.ts and the formula-lib-*.ts domain modules).

import {
  asRectangularValues,
  collectNumbers,
  collectValues,
  evaluationError,
  finiteResult,
  isEvaluationError,
  isEvaluationRange,
  isLambdaValue,
  MAX_RANGE_CELLS,
  scalarArgument,
  toBoolean,
  toNumber,
  toText,
} from "./formulas";
import type {
  EvaluationError,
  EvaluationRange,
  EvaluationScalar,
  EvaluationValue,
  FormulaError,
  FunctionEvaluation,
  FunctionSpec,
  RectangularValues,
} from "./formulas";

export type Value = EvaluationValue;
export type Scalar = EvaluationScalar;
export type Spec = FunctionSpec;
export type Call = FunctionEvaluation;
export type Rect = RectangularValues;
export type Specs = Record<string, FunctionSpec>;

export const error = (code: FormulaError): EvaluationError => evaluationError(code);
export const valueError = (): EvaluationError => evaluationError("#VALUE!");
export const numError = (): EvaluationError => evaluationError("#NUM!");
export const naError = (): EvaluationError => evaluationError("#N/A");
export const divError = (): EvaluationError => evaluationError("#DIV/0!");
export const refError = (): EvaluationError => evaluationError("#REF!");
export const calcError = (): EvaluationError => evaluationError("#CALC!");

export { isEvaluationError as isError, isEvaluationRange as isRange, isLambdaValue as isLambda };

/** An argument that was not passed (`undefined`) or passed empty/blank (`null`). */
export function isMissing(value: Value | undefined): boolean {
  return value === undefined || value === null;
}

/**
 * A single numeric argument. A missing argument (not passed) yields `fallback`
 * (or #VALUE! when no fallback is given); an omitted/blank one coerces to 0, like Excel.
 */
export function numberArg(value: Value | undefined, fallback?: number): number | EvaluationError {
  if (value === undefined) return fallback === undefined ? valueError() : fallback;
  return toNumber(scalarArgument(value));
}

/** Like numberArg, but an omitted/blank argument (`FN(x,,y)`) also takes the default. */
export function optionalNumberArg(value: Value | undefined, fallback: number): number | EvaluationError {
  if (value === undefined || value === null) return fallback;
  return toNumber(scalarArgument(value));
}

/** A numeric argument truncated toward zero (Excel's integer coercion for counts/indices). */
export function integerArg(value: Value | undefined, fallback?: number): number | EvaluationError {
  const number = numberArg(value, fallback);
  return isEvaluationError(number) ? number : Math.trunc(number);
}

export function optionalIntegerArg(value: Value | undefined, fallback: number): number | EvaluationError {
  const number = optionalNumberArg(value, fallback);
  return isEvaluationError(number) ? number : Math.trunc(number);
}

export function booleanArg(value: Value | undefined, fallback: boolean): boolean | EvaluationError {
  if (value === undefined) return fallback;
  return toBoolean(scalarArgument(value));
}

export function optionalBooleanArg(value: Value | undefined, fallback: boolean): boolean | EvaluationError {
  if (value === undefined || value === null) return fallback;
  return toBoolean(scalarArgument(value));
}

export function textArg(value: Value | undefined, fallback?: string): string | EvaluationError {
  if (value === undefined) return fallback === undefined ? valueError() : fallback;
  return toText(scalarArgument(value));
}

/** Numbers for aggregate functions: ranges skip text/logicals/blanks, direct arguments coerce. */
export function readNumbers(values: Value[]): number[] | EvaluationError {
  const collected = collectNumbers(values);
  return collected.error ?? collected.values;
}

/**
 * Numbers for the *A aggregates (MAXA, STDEVA, ...): inside ranges text counts as 0 and
 * logicals as 1/0 (blanks skipped); direct arguments coerce, non-numeric text is #VALUE!.
 */
export function readNumbersA(values: Value[]): number[] | EvaluationError {
  const numbers: number[] = [];
  for (const entry of collectValues(values)) {
    const value = entry.value;
    if (isEvaluationError(value)) return value;
    if (value === null) continue;
    if (typeof value === "number") numbers.push(value);
    else if (typeof value === "boolean") numbers.push(value ? 1 : 0);
    else if (entry.fromRange) numbers.push(0);
    else {
      const converted = toNumber(value);
      if (isEvaluationError(converted)) return converted;
      numbers.push(converted);
    }
  }
  return numbers;
}

/** A rectangular array argument (scalars become 1x1). Sparse ranges are #VALUE!. */
export function rectArg(value: Value | undefined): Rect | EvaluationError {
  if (value === undefined) return valueError();
  return asRectangularValues(value);
}

export function arrayResult(values: Scalar[], rowCount: number, columnCount: number): Value {
  if (rowCount < 1 || columnCount < 1 || values.length === 0) return calcError();
  return { kind: "evaluationRange", values, rowCount, columnCount } satisfies EvaluationRange;
}

export function rowsResult(rows: Scalar[][]): Value {
  if (rows.length === 0 || rows[0].length === 0) return calcError();
  const columnCount = rows[0].length;
  const values: Scalar[] = [];
  for (const row of rows) {
    for (let column = 0; column < columnCount; column += 1) values.push(row[column] ?? null);
  }
  return arrayResult(values, rows.length, columnCount);
}

export function tooManyCells(rowCount: number, columnCount: number): boolean {
  return columnCount > 0 && rowCount > MAX_RANGE_CELLS / columnCount;
}

export function checkedNumber(value: number): number | EvaluationError {
  return finiteResult(value);
}

/**
 * Paired numeric data for regression/correlation functions (SLOPE, CORREL, ...): both
 * arguments must hold the same number of cells (#N/A otherwise); errors propagate; only
 * positions where both values are numbers are kept.
 */
export function numericPairs(
  left: Value | undefined,
  right: Value | undefined,
): { xs: number[]; ys: number[] } | EvaluationError {
  const a = rectArg(left);
  if (isEvaluationError(a)) return a;
  const b = rectArg(right);
  if (isEvaluationError(b)) return b;
  if (a.values.length !== b.values.length) return naError();
  const xs: number[] = [];
  const ys: number[] = [];
  for (let index = 0; index < a.values.length; index += 1) {
    const x = a.values[index];
    const y = b.values[index];
    if (isEvaluationError(x)) return x;
    if (isEvaluationError(y)) return y;
    if (typeof x === "number" && typeof y === "number") {
      xs.push(x);
      ys.push(y);
    }
  }
  return { xs, ys };
}

/**
 * A scalar numeric function lifted over arrays. `defaults[i]` is the value of optional
 * argument i when it is not passed (`undefined` marks a required argument); omitted/blank
 * arguments coerce to 0 as in Excel. Numeric results pass through finiteResult (#NUM! on
 * overflow/NaN).
 */
export function numericFunction(
  defaults: Array<number | undefined>,
  compute: (...args: number[]) => number | EvaluationError | Value,
  options: Partial<Pick<FunctionSpec, "volatile" | "liftArgs">> = {},
): FunctionSpec {
  let minArgs = 0;
  while (minArgs < defaults.length && defaults[minArgs] === undefined) minArgs += 1;
  return {
    minArgs,
    maxArgs: defaults.length,
    liftArgs: options.liftArgs ?? "all",
    volatile: options.volatile,
    impl: (values) => {
      const numbers: number[] = [];
      for (let index = 0; index < defaults.length; index += 1) {
        const number = numberArg(values[index], defaults[index]);
        if (isEvaluationError(number)) return number;
        numbers.push(number);
      }
      const result = compute(...numbers);
      return typeof result === "number" ? finiteResult(result) : result;
    },
  };
}

/** Build a spec with the common shape. */
export function spec(
  minArgs: number,
  maxArgs: number,
  impl: FunctionSpec["impl"],
  options: Partial<Omit<FunctionSpec, "minArgs" | "maxArgs" | "impl">> = {},
): FunctionSpec {
  return { minArgs, maxArgs, impl, ...options };
}

/** Relative/absolute tolerance comparison used by iterative solvers. */
export function nearlyEqual(left: number, right: number, tolerance = 1e-12): boolean {
  return Math.abs(left - right) <= tolerance * Math.max(1, Math.abs(left), Math.abs(right));
}
