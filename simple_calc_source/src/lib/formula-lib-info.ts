// Information, logical, and database (D*) functions for the extended
// library (see formula-library.ts).

import {
  compareValues,
  createCriterionTest,
  criterionOperandValue,
  isEvaluationError,
  isEvaluationRange,
  isLambdaValue,
  scalarArgument,
  toNumber,
  toText,
} from "./formulas";
import type { CriterionTest, EvaluationError, FormulaError } from "./formulas";
import {
  divError,
  naError,
  numError,
  optionalBooleanArg,
  rectArg,
  spec,
  textArg,
  valueError,
} from "./formula-lib-shared";
import type { Scalar, Spec, Specs, Value } from "./formula-lib-shared";

const ERROR_TYPE_CODES: Partial<Record<FormulaError, number>> = {
  "#NULL!": 1,
  "#DIV/0!": 2,
  "#VALUE!": 3,
  "#REF!": 4,
  "#NAME?": 5,
  "#NUM!": 6,
  "#N/A": 7,
  "#SPILL!": 9,
  "#CALC!": 14,
};

function parityTest(odd: boolean): Spec {
  return spec(
    1,
    1,
    (values) => {
      const value = scalarArgument(values[0]);
      if (isEvaluationError(value)) return value;
      if (typeof value === "boolean") return valueError();
      const number = toNumber(value);
      if (isEvaluationError(number)) return number;
      const remainder = Math.abs(Math.trunc(number)) % 2;
      return odd ? remainder === 1 : remainder === 0;
    },
    { liftArgs: "all" },
  );
}

function detectPlatformIsMac(): boolean {
  const navigatorLike = (globalThis as { navigator?: { platform?: string; userAgent?: string } }).navigator;
  const platform = navigatorLike?.platform ?? navigatorLike?.userAgent ?? "";
  if (platform) return /mac/i.test(platform);
  const processLike = (globalThis as { process?: { platform?: string } }).process;
  return processLike?.platform === "darwin";
}

// ---- Database functions -----------------------------------------------------------------

interface DatabaseSelection {
  /** Values of the chosen field for matching records, or null when no field was given. */
  fieldValues: Scalar[] | null;
  /** Number of matching records. */
  count: number;
}

function headerKey(value: Scalar): string {
  const text = value === null ? "" : isEvaluationError(value) ? value.code : toText(value);
  return (typeof text === "string" ? text : "").trim().toLocaleLowerCase();
}

/**
 * A database criteria cell. Text without a comparison operator matches values that begin with
 * it (Excel's database/advanced-filter rule); "=text" is an exact match; "=" alone matches blank
 * cells and "<>" alone non-blank cells.
 */
function databaseCriterion(cell: Scalar): CriterionTest | EvaluationError | null {
  if (cell === null || cell === "") return null;
  if (isEvaluationError(cell)) return cell;
  // Only real text matches as a prefix; "1/1/2024", "50%" or "$1" compare as values.
  if (typeof cell === "string" && !/^(<=|>=|<>|=|<|>)/.test(cell) && typeof criterionOperandValue(cell) === "string") {
    return createCriterionTest(`${cell}*`);
  }
  return createCriterionTest(cell);
}

function selectRecords(values: Value[], fieldOptional: boolean): DatabaseSelection | EvaluationError {
  const database = rectArg(values[0]);
  if (isEvaluationError(database)) return database;
  if (database.rowCount < 1) return valueError();
  const columns = database.columnCount;
  const headers = database.values.slice(0, columns).map(headerKey);

  let fieldIndex = -1;
  const field = values[1];
  if (field === undefined || field === null) {
    if (!fieldOptional) return valueError();
  } else {
    const scalar = scalarArgument(field);
    if (isEvaluationError(scalar)) return scalar;
    if (typeof scalar === "number") {
      fieldIndex = Math.trunc(scalar) - 1;
      if (fieldIndex < 0 || fieldIndex >= columns) return valueError();
    } else {
      const wanted = headerKey(scalar);
      fieldIndex = headers.indexOf(wanted);
      if (fieldIndex < 0) return valueError();
    }
  }

  const criteria = rectArg(values[2]);
  if (isEvaluationError(criteria)) return criteria;
  const criteriaColumns: number[] = [];
  for (let column = 0; column < criteria.columnCount; column += 1) {
    criteriaColumns.push(headers.indexOf(headerKey(criteria.values[column])));
  }
  // OR across criteria rows, AND across the columns of one row. A row with no conditions
  // matches every record.
  const rowTests: Array<Array<{ column: number; test: CriterionTest }>> = [];
  for (let row = 1; row < criteria.rowCount; row += 1) {
    const tests: Array<{ column: number; test: CriterionTest }> = [];
    for (let column = 0; column < criteria.columnCount; column += 1) {
      const target = criteriaColumns[column];
      const cell = criteria.values[row * criteria.columnCount + column];
      // Computed criteria (a header that is not a field name) cannot be re-evaluated per record.
      if (target < 0) continue;
      const test = databaseCriterion(cell);
      if (test === null) continue;
      if (isEvaluationError(test)) return test;
      tests.push({ column: target, test });
    }
    rowTests.push(tests);
  }
  const matchesAll = rowTests.length === 0 || rowTests.some((tests) => tests.length === 0);

  const fieldValues: Scalar[] = [];
  let count = 0;
  for (let row = 1; row < database.rowCount; row += 1) {
    const offset = row * columns;
    let matched = matchesAll;
    if (!matched) {
      for (const tests of rowTests) {
        let all = true;
        for (const { column, test } of tests) {
          const result = test(database.values[offset + column]);
          if (isEvaluationError(result) || !result) {
            all = false;
            break;
          }
        }
        if (all) {
          matched = true;
          break;
        }
      }
    }
    if (!matched) continue;
    count += 1;
    if (fieldIndex >= 0) fieldValues.push(database.values[offset + fieldIndex]);
  }
  return { fieldValues: fieldIndex >= 0 ? fieldValues : null, count };
}

function numericFieldValues(selection: DatabaseSelection): number[] | EvaluationError {
  const numbers: number[] = [];
  for (const value of selection.fieldValues ?? []) {
    if (isEvaluationError(value)) return value;
    if (typeof value === "number") numbers.push(value);
  }
  return numbers;
}

function databaseSpec(compute: (numbers: number[], selection: DatabaseSelection) => Value, fieldOptional = false): Spec {
  return spec(3, 3, (values) => {
    const selection = selectRecords(values, fieldOptional);
    if (isEvaluationError(selection)) return selection;
    const numbers = numericFieldValues(selection);
    if (isEvaluationError(numbers)) return numbers;
    return compute(numbers, selection);
  });
}

function variance(numbers: number[], sample: boolean): number | EvaluationError {
  const count = numbers.length;
  if (count === 0 || (sample && count < 2)) return divError();
  const mean = numbers.reduce((sum, value) => sum + value, 0) / count;
  const squares = numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0);
  return squares / (sample ? count - 1 : count);
}

function sqrtOf(value: number | EvaluationError): number | EvaluationError {
  return isEvaluationError(value) ? value : Math.sqrt(value);
}

// ---- Registry ---------------------------------------------------------------------------

export const INFO_FUNCTIONS: Specs = {
  ISEVEN: parityTest(false),
  ISODD: parityTest(true),
  ISNONTEXT: spec(1, 1, (values) => typeof scalarArgument(values[0]) !== "string", { liftArgs: "all" }),
  TYPE: spec(1, 1, (values) => {
    const value = values[0];
    if (isLambdaValue(value)) return 128;
    if (isEvaluationRange(value)) {
      if (value.sparse || value.values.length !== 1) return 64;
      return typeCode(value.values[0]);
    }
    return typeCode(value);
  }),
  "ERROR.TYPE": spec(
    1,
    1,
    (values) => {
      const value = scalarArgument(values[0]);
      if (!isEvaluationError(value)) return naError();
      return ERROR_TYPE_CODES[value.code] ?? naError();
    },
    { liftArgs: "all" },
  ),
  TRUE: spec(0, 0, () => true),
  FALSE: spec(0, 0, () => false),
  N: spec(
    1,
    1,
    (values) => {
      const value = scalarArgument(values[0]);
      if (isEvaluationError(value)) return value;
      if (typeof value === "number") return value;
      if (typeof value === "boolean") return value ? 1 : 0;
      return 0;
    },
    { liftArgs: "all" },
  ),
  INFO: spec(
    1,
    1,
    (values, call) => {
      const kind = textArg(values[0]);
      if (isEvaluationError(kind)) return kind;
      const mac = detectPlatformIsMac();
      switch (kind.trim().toLowerCase()) {
        case "directory":
          return "";
        case "numfile":
          return Math.max(1, call.hooks.getSheetNames?.().length ?? 1);
        case "origin":
          return "$A:$A$1";
        case "osversion":
          return mac ? "Macintosh (Intel) Version 14.0" : "Windows (64-bit) NT 10.00";
        case "recalc":
          return "Automatic";
        case "release":
          return "16.0";
        case "system":
          return mac ? "mac" : "pcdos";
        default:
          return valueError();
      }
    },
    { volatile: true },
  ),
  ISEMAIL: spec(
    1,
    1,
    (values) => {
      const value = scalarArgument(values[0]);
      return typeof value === "string" && /^[^\s@"(),:;<>[\]\\]+@[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)*\.[A-Za-z]{2,}$/.test(value.trim());
    },
    { liftArgs: "all" },
  ),
  ISURL: spec(
    1,
    1,
    (values) => {
      const value = scalarArgument(values[0]);
      return (
        typeof value === "string" &&
        /^(?:(?:https?|ftp):\/\/)?(?:[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?\.)+[A-Za-z]{2,}(?::\d{1,5})?(?:[/?#]\S*)?$/i.test(value.trim())
      );
    },
    { liftArgs: "all" },
  ),
  ISBETWEEN: spec(
    3,
    5,
    (values) => {
      const value = scalarArgument(values[0]);
      const lower = scalarArgument(values[1]);
      const upper = scalarArgument(values[2]);
      const lowerInclusive = optionalBooleanArg(values[3], true);
      if (isEvaluationError(lowerInclusive)) return lowerInclusive;
      const upperInclusive = optionalBooleanArg(values[4], true);
      if (isEvaluationError(upperInclusive)) return upperInclusive;
      const low = compareValues(value, lower);
      if (isEvaluationError(low)) return low;
      const high = compareValues(value, upper);
      if (isEvaluationError(high)) return high;
      return (lowerInclusive ? low >= 0 : low > 0) && (upperInclusive ? high <= 0 : high < 0);
    },
    { liftArgs: "all" },
  ),

  // Database functions: D*(database, field, criteria).
  DSUM: databaseSpec((numbers) => numbers.reduce((sum, value) => sum + value, 0)),
  DAVERAGE: databaseSpec((numbers) =>
    numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : divError(),
  ),
  DCOUNT: databaseSpec((numbers, selection) => (selection.fieldValues === null ? selection.count : numbers.length), true),
  DCOUNTA: databaseSpec(
    (_numbers, selection) =>
      selection.fieldValues === null
        ? selection.count
        : selection.fieldValues.filter((value) => value !== null && value !== "").length,
    true,
  ),
  DGET: spec(3, 3, (values) => {
    const selection = selectRecords(values, false);
    if (isEvaluationError(selection)) return selection;
    const found = selection.fieldValues ?? [];
    if (found.length === 0) return valueError();
    if (found.length > 1) return numError();
    return found[0];
  }),
  DMAX: databaseSpec((numbers) => (numbers.length ? numbers.reduce((best, value) => Math.max(best, value)) : 0)),
  DMIN: databaseSpec((numbers) => (numbers.length ? numbers.reduce((best, value) => Math.min(best, value)) : 0)),
  DPRODUCT: databaseSpec((numbers) => (numbers.length ? numbers.reduce((product, value) => product * value, 1) : 0)),
  DSTDEV: databaseSpec((numbers) => sqrtOf(variance(numbers, true))),
  DSTDEVP: databaseSpec((numbers) => sqrtOf(variance(numbers, false))),
  DVAR: databaseSpec((numbers) => variance(numbers, true)),
  DVARP: databaseSpec((numbers) => variance(numbers, false)),
};

function typeCode(value: Scalar): number {
  if (value === null || typeof value === "number") return 1;
  if (typeof value === "string") return 2;
  if (typeof value === "boolean") return 4;
  return 16;
}
