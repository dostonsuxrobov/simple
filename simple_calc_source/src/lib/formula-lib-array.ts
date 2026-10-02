// Dynamic-array shaping, lookup/reference, and Google Sheets array functions for the extended
// library (see formula-library.ts).

import {
  arrayEntryKey,
  collectValues,
  columnNumberToLabel,
  compareValues,
  hasFormulaFunction,
  isEvaluationError,
  isEvaluationRange,
  isLambdaValue,
  lambdaScalarResult,
  lookupComparison,
  MAX_EXCEL_COLUMN,
  MAX_EXCEL_ROW,
  numberToText,
  rectangleRows,
  scalarArgument,
  sortScalarCompare,
  toBoolean,
  toNumber,
  toText,
  transposedRows,
  wildcardMatches,
  wildcardTokens,
} from "./formulas";
import type { EvaluationError, EvaluationRange, FormulaNode, RectangleBounds } from "./formulas";
import {
  arrayResult,
  calcError,
  integerArg,
  naError,
  numError,
  optionalBooleanArg,
  optionalIntegerArg,
  optionalNumberArg,
  readNumbers,
  rectArg,
  rowsResult,
  spec,
  textArg,
  tooManyCells,
  valueError,
} from "./formula-lib-shared";
import type { Call, Rect, Scalar, Specs, Value } from "./formula-lib-shared";
import { compileRegex } from "./formula-lib-text";

// ---- Lazy-argument helpers --------------------------------------------------------------

/** Evaluate argument `index` of a lazy call: undefined when absent, null when omitted. */
function lazyArg(call: Call, index: number): Value | undefined {
  const node = call.argumentNodes[index];
  if (!node) return undefined;
  if (node.kind === "omitted") return null;
  return call.evaluate(node);
}

function isOmittedNode(node: FormulaNode | undefined): boolean {
  return !node || node.kind === "omitted";
}

// Derived column/row vectors are memoized per source array so lifted lookups reuse one vector
// (and its lookup cache) across calls.
const VECTOR_CACHE = new WeakMap<Scalar[], Map<string, Scalar[]>>();

function cachedVector(range: Rect, key: string, build: () => Scalar[]): Scalar[] {
  let vectors = VECTOR_CACHE.get(range.values);
  if (!vectors) {
    vectors = new Map();
    VECTOR_CACHE.set(range.values, vectors);
  }
  let vector = vectors.get(key);
  if (!vector) {
    vector = build();
    vectors.set(key, vector);
  }
  return vector;
}

function columnValues(range: Rect, column: number): Scalar[] {
  if (range.columnCount === 1) return range.values;
  return cachedVector(range, `c${column}`, () => {
    const values: Scalar[] = new Array(range.rowCount);
    for (let row = 0; row < range.rowCount; row += 1) values[row] = range.values[row * range.columnCount + column];
    return values;
  });
}

function rowValues(range: Rect, row: number): Scalar[] {
  if (range.rowCount === 1) return range.values;
  return cachedVector(range, `r${row}`, () => range.values.slice(row * range.columnCount, (row + 1) * range.columnCount));
}

/** Numbers from index arguments that may be scalars or arrays (CHOOSECOLS/CHOOSEROWS). */
function indexNumbers(values: Value[]): number[] | EvaluationError {
  const numbers: number[] = [];
  for (const entry of collectValues(values)) {
    const number = toNumber(entry.value);
    if (isEvaluationError(number)) return number;
    numbers.push(Math.trunc(number));
  }
  return numbers;
}

function padValue(value: Value | undefined): Scalar {
  if (value === undefined) return naError();
  const scalar = scalarArgument(value);
  return scalar;
}

// ---- Lookup helpers ---------------------------------------------------------------------

interface LookupCache {
  seen: number;
  first?: Map<string, number>;
  last?: Map<string, number>;
  sortedNumeric?: boolean;
}

// Lifted lookups (XMATCH(A1:A50000, B1:B50000, 0)) call the function once per lookup value
// with the same candidate array. After the second call against the same array, build a hash
// index so the whole lifted call is O(n) instead of O(n²).
const LOOKUP_CACHE = new WeakMap<Scalar[], LookupCache>();
const INDEX_THRESHOLD = 32;

function lookupCache(candidates: Scalar[]): LookupCache {
  let entry = LOOKUP_CACHE.get(candidates);
  if (!entry) {
    entry = { seen: 0 };
    LOOKUP_CACHE.set(candidates, entry);
  }
  entry.seen += 1;
  return entry;
}

function exactKey(value: Scalar): string | null {
  if (value === null) return "s:";
  if (typeof value === "number") return `n:${value}`;
  if (typeof value === "string") return `s:${value.toLocaleLowerCase()}`;
  if (typeof value === "boolean") return `b:${value}`;
  return null;
}

function exactIndex(candidates: Scalar[], key: Scalar, reverse: boolean): number {
  const cache = lookupCache(candidates);
  if (!cache.first && cache.seen >= 2 && candidates.length > INDEX_THRESHOLD) {
    const first = new Map<string, number>();
    const last = new Map<string, number>();
    for (let index = 0; index < candidates.length; index += 1) {
      const candidateKey = exactKey(candidates[index]);
      if (candidateKey === null) continue;
      if (!first.has(candidateKey)) first.set(candidateKey, index);
      last.set(candidateKey, index);
    }
    cache.first = first;
    cache.last = last;
  }
  if (cache.first && cache.last) {
    const searchKey = exactKey(key);
    if (searchKey === null) return -1;
    return (reverse ? cache.last : cache.first).get(searchKey) ?? -1;
  }
  const count = candidates.length;
  for (let offset = 0; offset < count; offset += 1) {
    const index = reverse ? count - 1 - offset : offset;
    if (lookupComparison(candidates[index], key) === 0) return index;
  }
  return -1;
}

/** Excel approximate (sorted) lookup: the last candidate <= key, skipping errors and other types. */
function approximateIndex(candidates: Scalar[], key: Scalar): number {
  if (typeof key === "number" && candidates.length > INDEX_THRESHOLD) {
    const cache = lookupCache(candidates);
    if (cache.sortedNumeric === undefined && cache.seen >= 2) {
      let sorted = true;
      for (let index = 0; index < candidates.length; index += 1) {
        const value = candidates[index];
        if (typeof value !== "number" || (index > 0 && value < (candidates[index - 1] as number))) {
          sorted = false;
          break;
        }
      }
      cache.sortedNumeric = sorted;
    }
    if (cache.sortedNumeric) {
      let low = 0;
      let high = candidates.length - 1;
      let found = -1;
      while (low <= high) {
        const middle = (low + high) >> 1;
        if ((candidates[middle] as number) <= key) {
          found = middle;
          low = middle + 1;
        } else {
          high = middle - 1;
        }
      }
      return found;
    }
  }
  let found = -1;
  for (let index = 0; index < candidates.length; index += 1) {
    const comparison = lookupComparison(candidates[index], key);
    if (comparison === null || isEvaluationError(comparison)) continue;
    if (comparison <= 0) found = index;
    else break;
  }
  return found;
}

function binarySearchIndex(candidates: Scalar[], key: Scalar, matchMode: number, descending: boolean): number {
  // Lower bound under the search order, comparing with Excel's cross-type ordering.
  const order = (value: Scalar): number => {
    const comparison = compareValues(value, key);
    if (isEvaluationError(comparison)) return 1;
    return descending ? -comparison : comparison;
  };
  let low = 0;
  let high = candidates.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (order(candidates[middle]) < 0) low = middle + 1;
    else high = middle;
  }
  if (low < candidates.length && lookupComparison(candidates[low], key) === 0) return low;
  if (matchMode === 0) return -1;
  // Next larger (1) / next smaller (-1) in value terms.
  const wantsLarger = matchMode === 1;
  const candidate = descending ? (wantsLarger ? low - 1 : low) : wantsLarger ? low : low - 1;
  if (candidate < 0 || candidate >= candidates.length) return -1;
  const comparison = lookupComparison(candidates[candidate], key);
  if (comparison === null || isEvaluationError(comparison)) return -1;
  return (wantsLarger ? comparison > 0 : comparison < 0) ? candidate : -1;
}

/** XMATCH/XLOOKUP-style match. Returns a zero-based index, -1 for no match, or an error. */
export function extendedMatch(
  candidates: Scalar[],
  key: Scalar,
  matchMode: number,
  searchMode: number,
): number | EvaluationError {
  if (isEvaluationError(key)) return key;
  if (searchMode === 2 || searchMode === -2) {
    if (matchMode === 2 || matchMode === 3) return valueError();
    return binarySearchIndex(candidates, key, matchMode, searchMode === -2);
  }
  const reverse = searchMode === -1;
  const count = candidates.length;
  if (matchMode === 2 && typeof key === "string") {
    const tokens = wildcardTokens(key, true);
    for (let offset = 0; offset < count; offset += 1) {
      const index = reverse ? count - 1 - offset : offset;
      const candidate = candidates[index];
      if (typeof candidate !== "string") continue;
      // The step limit applies to each candidate, not to the whole lookup range.
      const matched = wildcardMatches(candidate, tokens, true);
      if (matched === null) return valueError();
      if (matched) return index;
    }
    return -1;
  }
  if (matchMode === 3) {
    const pattern = toText(key);
    if (isEvaluationError(pattern)) return pattern;
    const regex = compileRegex(pattern, false);
    if (isEvaluationError(regex)) return regex;
    for (let offset = 0; offset < count; offset += 1) {
      const index = reverse ? count - 1 - offset : offset;
      const candidate = candidates[index];
      const text = typeof candidate === "string" ? candidate : typeof candidate === "number" ? numberToText(candidate) : null;
      if (text === null) continue;
      regex.lastIndex = 0;
      if (regex.test(text)) return index;
    }
    return -1;
  }
  if (matchMode !== -1 && matchMode !== 1) return exactIndex(candidates, key, reverse);
  let best = -1;
  for (let offset = 0; offset < count; offset += 1) {
    const index = reverse ? count - 1 - offset : offset;
    const comparison = lookupComparison(candidates[index], key);
    if (comparison === null || isEvaluationError(comparison)) continue;
    if (comparison === 0) return index;
    if (matchMode === 1 ? comparison < 0 : comparison > 0) continue;
    if (best < 0) {
      best = index;
      continue;
    }
    const relative = lookupComparison(candidates[index], candidates[best]);
    if (typeof relative === "number" && (matchMode === 1 ? relative < 0 : relative > 0)) best = index;
  }
  return best;
}

function quoteSheetName(name: string): string {
  const withoutBook = name.replace(/^\[[^\]]*\]/, "");
  const plain =
    /^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(withoutBook) &&
    !/^[A-Za-z]{1,3}\d+$/.test(withoutBook) &&
    !/^[Rr]\d*[Cc]\d*$/.test(withoutBook) &&
    !/^(TRUE|FALSE)$/i.test(withoutBook);
  return plain ? name : `'${name.replace(/'/g, "''")}'`;
}

// ---- GROUPBY / PIVOTBY ------------------------------------------------------------------

type Aggregator = (group: EvaluationRange, all: EvaluationRange) => Scalar;

function columnRange(values: Scalar[]): EvaluationRange {
  return { kind: "evaluationRange", values, rowCount: values.length, columnCount: 1 };
}

function percentOf(subset: EvaluationRange, all: EvaluationRange): Scalar {
  const part = readNumbers([subset]);
  if (isEvaluationError(part)) return part;
  const whole = readNumbers([all]);
  if (isEvaluationError(whole)) return whole;
  const total = whole.reduce((sum, value) => sum + value, 0);
  if (total === 0) return { kind: "evaluationError", code: "#DIV/0!" };
  return part.reduce((sum, value) => sum + value, 0) / total;
}

/** A LAMBDA, or an eta-reduced function name such as SUM (stored as `_xleta.SUM` in files). */
function resolveAggregator(call: Call, node: FormulaNode | undefined): Aggregator | EvaluationError {
  if (!node || node.kind === "omitted") return valueError();
  if (node.kind === "name") {
    const bare = node.name.replace(/^_xleta\./i, "").replace(/^_xlfn\./i, "").toUpperCase();
    const evaluated = /^_xleta\./i.test(node.name) ? null : call.evaluate(node);
    if (evaluated !== null && isLambdaValue(evaluated)) return lambdaAggregator(call, evaluated);
    if (bare === "PERCENTOF") return percentOf;
    if (hasFormulaFunction(bare)) {
      return (group) =>
        lambdaScalarResult(
          call.evaluate({
            kind: "call",
            name: bare,
            arguments: [{ kind: "array", values: group.values, rowCount: group.rowCount, columnCount: group.columnCount }],
          }),
        );
    }
    return isEvaluationError(evaluated) ? evaluated : valueError();
  }
  const value = call.evaluate(node);
  if (isLambdaValue(value)) return lambdaAggregator(call, value);
  return isEvaluationError(value) ? value : valueError();
}

function lambdaAggregator(call: Call, lambda: Value): Aggregator {
  const parameterCount = isLambdaValue(lambda) ? lambda.parameters.length : 1;
  return (group, all) => lambdaScalarResult(call.invokeLambda(lambda, parameterCount >= 2 ? [group, all] : [group]));
}

interface FieldTable {
  /** Header cells per field column (null when the data has no header row). */
  headers: Scalar[] | null;
  /** Data rows (header removed). */
  rows: Scalar[][];
}

function splitHeaders(range: Rect, hasHeaders: boolean): FieldTable {
  const rows = rectangleRows(range);
  return hasHeaders ? { headers: rows[0], rows: rows.slice(1) } : { headers: null, rows };
}

/** Excel's header auto-detection: text in the first values row, numbers in the second. */
function detectHeaders(values: Rect): boolean {
  if (values.rowCount < 2) return false;
  for (let column = 0; column < values.columnCount; column += 1) {
    const first = values.values[column];
    const second = values.values[values.columnCount + column];
    if (typeof first !== "string" || typeof second !== "number") return false;
  }
  return true;
}

interface HeaderMode {
  hasHeaders: boolean;
  showHeaders: boolean;
  generate: boolean;
}

function headerMode(raw: Value | undefined, values: Rect): HeaderMode | EvaluationError {
  if (raw === undefined || raw === null) {
    const detected = detectHeaders(values);
    return { hasHeaders: detected, showHeaders: detected, generate: false };
  }
  const mode = toNumber(scalarArgument(raw));
  if (isEvaluationError(mode)) return mode;
  switch (Math.trunc(mode)) {
    case 0:
      return { hasHeaders: false, showHeaders: false, generate: false };
    case 1:
      return { hasHeaders: true, showHeaders: false, generate: false };
    case 2:
      return { hasHeaders: false, showHeaders: true, generate: true };
    case 3:
      return { hasHeaders: true, showHeaders: true, generate: false };
    default:
      return valueError();
  }
}

function filterFlags(raw: Value | undefined, dataRows: number, totalRows: number): boolean[] | null | EvaluationError {
  if (raw === undefined || raw === null) return null;
  const range = rectArg(raw);
  if (isEvaluationError(range)) return range;
  if (range.columnCount !== 1 && range.rowCount !== 1) return valueError();
  const flags = range.values;
  if (flags.length !== dataRows && flags.length !== totalRows) return valueError();
  const offset = flags.length === totalRows ? totalRows - dataRows : 0;
  const result: boolean[] = [];
  for (let index = 0; index < dataRows; index += 1) {
    const flag = toBoolean(flags[index + offset]);
    if (isEvaluationError(flag)) return flag;
    result.push(flag);
  }
  return result;
}

interface Group {
  fields: Scalar[];
  rows: number[];
}

function groupRows(rows: Scalar[][], include: (index: number) => boolean): Group[] {
  const groups = new Map<string, Group>();
  for (let index = 0; index < rows.length; index += 1) {
    if (!include(index)) continue;
    const fields = rows[index];
    const key = fields.map(arrayEntryKey).join("\u0001");
    let group = groups.get(key);
    if (!group) {
      group = { fields, rows: [] };
      groups.set(key, group);
    }
    group.rows.push(index);
  }
  return [...groups.values()];
}

function compareFieldLists(left: Scalar[], right: Scalar[], upTo: number, direction: number): number {
  for (let index = 0; index < upTo; index += 1) {
    const comparison = sortScalarCompare(left[index], right[index]);
    if (comparison !== 0) return comparison * direction;
  }
  return 0;
}

function sortOrders(raw: Value | undefined): number[] | EvaluationError {
  if (raw === undefined || raw === null) return [1];
  const orders: number[] = [];
  for (const entry of collectValues([raw])) {
    const number = toNumber(entry.value);
    if (isEvaluationError(number)) return number;
    const order = Math.trunc(number);
    if (order === 0) return valueError();
    orders.push(order);
  }
  return orders.length ? orders : [1];
}

function aggregateGroup(
  aggregators: Aggregator,
  valueColumns: Scalar[][],
  rows: number[],
  include: (index: number) => boolean,
): Scalar[] {
  return valueColumns.map((column) => {
    const all: Scalar[] = [];
    for (let index = 0; index < column.length; index += 1) if (include(index)) all.push(column[index]);
    const subset = rows.map((row) => column[row]);
    return aggregators(columnRange(subset), columnRange(all));
  });
}

function evaluateGroupBy(call: Call): Value {
  const rowFieldsValue = lazyArg(call, 0);
  const valuesValue = lazyArg(call, 1);
  const rowFields = rectArg(rowFieldsValue);
  if (isEvaluationError(rowFields)) return rowFields;
  const values = rectArg(valuesValue);
  if (isEvaluationError(values)) return values;
  if (rowFields.rowCount !== values.rowCount) return valueError();
  const aggregator = resolveAggregator(call, call.argumentNodes[2]);
  if (isEvaluationError(aggregator)) return aggregator;
  const headers = headerMode(lazyArg(call, 3), values);
  if (isEvaluationError(headers)) return headers;
  const totalDepth = optionalIntegerArg(lazyArg(call, 4), 1);
  if (isEvaluationError(totalDepth)) return totalDepth;
  const orders = sortOrders(lazyArg(call, 5));
  if (isEvaluationError(orders)) return orders;
  const fieldTable = splitHeaders(rowFields, headers.hasHeaders);
  const valueTable = splitHeaders(values, headers.hasHeaders);
  const dataRows = fieldTable.rows.length;
  const flags = filterFlags(lazyArg(call, 6), dataRows, rowFields.rowCount);
  if (isEvaluationError(flags)) return flags;
  const include = (index: number) => (flags ? flags[index] : true);
  const fieldCount = rowFields.columnCount;
  const valueCount = values.columnCount;
  const valueColumns: Scalar[][] = [];
  for (let column = 0; column < valueCount; column += 1) valueColumns.push(valueTable.rows.map((row) => row[column]));

  const groups = groupRows(fieldTable.rows, include);
  if (groups.length === 0) return calcError();
  const results = groups.map((group) => ({ group, aggregates: aggregateGroup(aggregator, valueColumns, group.rows, include) }));

  const depth = Math.abs(totalDepth);
  const hierarchical = depth >= 2 && fieldCount >= 2;
  const primary = orders[0];
  const sortIndex = Math.abs(primary) - 1;
  const direction = primary < 0 ? -1 : 1;
  if (sortIndex >= fieldCount + valueCount) return valueError();
  results.sort((left, right) => {
    if (hierarchical) {
      const parent = sortScalarCompare(left.group.fields[0], right.group.fields[0]);
      if (parent !== 0 && sortIndex !== 0) return parent;
    }
    const byKey =
      sortIndex < fieldCount
        ? sortScalarCompare(left.group.fields[sortIndex], right.group.fields[sortIndex])
        : sortScalarCompare(left.aggregates[sortIndex - fieldCount], right.aggregates[sortIndex - fieldCount]);
    if (byKey !== 0) return byKey * direction;
    return compareFieldLists(left.group.fields, right.group.fields, fieldCount, 1);
  });

  const width = fieldCount + valueCount;
  const output: Scalar[][] = [];
  if (headers.showHeaders) {
    const header: Scalar[] = [];
    for (let column = 0; column < fieldCount; column += 1) {
      header.push(headers.generate || !fieldTable.headers ? `Row Field ${column + 1}` : fieldTable.headers[column]);
    }
    for (let column = 0; column < valueCount; column += 1) {
      header.push(headers.generate || !valueTable.headers ? `Value ${column + 1}` : valueTable.headers[column]);
    }
    output.push(header);
  }
  const allRows: number[] = [];
  for (let index = 0; index < dataRows; index += 1) if (include(index)) allRows.push(index);
  const totalRow = (): Scalar[] => {
    const row: Scalar[] = new Array(width).fill("");
    row[0] = "Total";
    aggregateGroup(aggregator, valueColumns, allRows, include).forEach((value, index) => {
      row[fieldCount + index] = value;
    });
    return row;
  };
  if (totalDepth < 0) output.push(totalRow());
  let index = 0;
  while (index < results.length) {
    if (!hierarchical) {
      output.push([...results[index].group.fields, ...results[index].aggregates]);
      index += 1;
      continue;
    }
    // A block of groups sharing the first field, with its subtotal before (-2) or after (2).
    const parentKey = arrayEntryKey(results[index].group.fields[0]);
    let end = index;
    const blockRows: number[] = [];
    while (end < results.length && arrayEntryKey(results[end].group.fields[0]) === parentKey) {
      blockRows.push(...results[end].group.rows);
      end += 1;
    }
    const subtotal: Scalar[] = new Array(width).fill("");
    subtotal[0] = results[index].group.fields[0];
    aggregateGroup(aggregator, valueColumns, blockRows, include).forEach((value, position) => {
      subtotal[fieldCount + position] = value;
    });
    if (totalDepth < 0) output.push(subtotal);
    for (let position = index; position < end; position += 1) {
      output.push([...results[position].group.fields, ...results[position].aggregates]);
    }
    if (totalDepth > 0) output.push(subtotal);
    index = end;
  }
  if (totalDepth > 0) output.push(totalRow());
  if (tooManyCells(output.length, width)) return valueError();
  return rowsResult(output);
}

function evaluatePivotBy(call: Call): Value {
  const rowFields = rectArg(lazyArg(call, 0));
  if (isEvaluationError(rowFields)) return rowFields;
  const colFields = rectArg(lazyArg(call, 1));
  if (isEvaluationError(colFields)) return colFields;
  const values = rectArg(lazyArg(call, 2));
  if (isEvaluationError(values)) return values;
  if (rowFields.rowCount !== values.rowCount || colFields.rowCount !== values.rowCount) return valueError();
  const aggregator = resolveAggregator(call, call.argumentNodes[3]);
  if (isEvaluationError(aggregator)) return aggregator;
  const headers = headerMode(lazyArg(call, 4), values);
  if (isEvaluationError(headers)) return headers;
  const rowTotalDepth = optionalIntegerArg(lazyArg(call, 5), 1);
  if (isEvaluationError(rowTotalDepth)) return rowTotalDepth;
  const rowOrders = sortOrders(lazyArg(call, 6));
  if (isEvaluationError(rowOrders)) return rowOrders;
  const colTotalDepth = optionalIntegerArg(lazyArg(call, 7), 1);
  if (isEvaluationError(colTotalDepth)) return colTotalDepth;
  const colOrders = sortOrders(lazyArg(call, 8));
  if (isEvaluationError(colOrders)) return colOrders;
  const rowTable = splitHeaders(rowFields, headers.hasHeaders);
  const colTable = splitHeaders(colFields, headers.hasHeaders);
  const valueTable = splitHeaders(values, headers.hasHeaders);
  const dataRows = rowTable.rows.length;
  const flags = filterFlags(lazyArg(call, 9), dataRows, values.rowCount);
  if (isEvaluationError(flags)) return flags;
  const relativeTo = optionalIntegerArg(lazyArg(call, 10), 0);
  if (isEvaluationError(relativeTo)) return relativeTo;
  const include = (index: number) => (flags ? flags[index] : true);

  const rowGroups = groupRows(rowTable.rows, include);
  const colGroups = groupRows(colTable.rows, include);
  if (rowGroups.length === 0 || colGroups.length === 0) return calcError();
  const rowDirection = rowOrders[0] < 0 ? -1 : 1;
  const colDirection = colOrders[0] < 0 ? -1 : 1;
  rowGroups.sort((left, right) => compareFieldLists(left.fields, right.fields, rowFields.columnCount, rowDirection));
  colGroups.sort((left, right) => compareFieldLists(left.fields, right.fields, colFields.columnCount, colDirection));

  const rowOf = new Int32Array(dataRows).fill(-1);
  rowGroups.forEach((group, index) => group.rows.forEach((row) => (rowOf[row] = index)));
  const colOf = new Int32Array(dataRows).fill(-1);
  colGroups.forEach((group, index) => group.rows.forEach((row) => (colOf[row] = index)));

  const valueCount = values.columnCount;
  const valueColumns: Scalar[][] = [];
  for (let column = 0; column < valueCount; column += 1) valueColumns.push(valueTable.rows.map((row) => row[column]));

  // Bucket row indices per (row group, column group) in one pass.
  const buckets = new Map<number, number[]>();
  const colCount = colGroups.length;
  for (let index = 0; index < dataRows; index += 1) {
    if (!include(index)) continue;
    const key = rowOf[index] * colCount + colOf[index];
    let bucket = buckets.get(key);
    if (!bucket) buckets.set(key, (bucket = []));
    bucket.push(index);
  }
  const allRows: number[] = [];
  for (let index = 0; index < dataRows; index += 1) if (include(index)) allRows.push(index);
  const pick = (column: number, rows: number[]) => columnRange(rows.map((row) => valueColumns[column][row]));
  const relativeRows = (rowGroup: number | null, colGroup: number | null): number[] => {
    if (relativeTo === 2) return allRows;
    if (relativeTo === 1 || relativeTo === 4) return rowGroup === null ? allRows : rowGroups[rowGroup].rows;
    return colGroup === null ? allRows : colGroups[colGroup].rows;
  };
  const cell = (column: number, rows: number[], rowGroup: number | null, colGroup: number | null): Scalar => {
    if (rows.length === 0) return "";
    return aggregator(pick(column, rows), pick(column, relativeRows(rowGroup, colGroup)));
  };

  const rowFieldCount = rowFields.columnCount;
  const colFieldCount = colFields.columnCount;
  const showRowTotals = colTotalDepth !== 0;
  const showColTotals = rowTotalDepth !== 0;
  const width = rowFieldCount + (colGroups.length + (showRowTotals ? 1 : 0)) * valueCount;
  const output: Scalar[][] = [];
  // Column-key header rows (one per column field), plus a value-name row for several values.
  for (let level = 0; level < colFieldCount; level += 1) {
    const header: Scalar[] = new Array(rowFieldCount).fill("");
    if (level === colFieldCount - 1 && headers.showHeaders && valueCount === 1) {
      for (let column = 0; column < rowFieldCount; column += 1) {
        header[column] = headers.generate || !rowTable.headers ? `Row Field ${column + 1}` : rowTable.headers[column];
      }
    }
    for (const group of colGroups) for (let value = 0; value < valueCount; value += 1) header.push(value === 0 ? group.fields[level] : "");
    if (showRowTotals) for (let value = 0; value < valueCount; value += 1) header.push(value === 0 && level === 0 ? "Total" : "");
    output.push(header);
  }
  if (valueCount > 1 && headers.showHeaders) {
    const header: Scalar[] = new Array(rowFieldCount).fill("");
    for (let column = 0; column < rowFieldCount; column += 1) {
      header[column] = headers.generate || !rowTable.headers ? `Row Field ${column + 1}` : rowTable.headers[column];
    }
    const names = valueTable.headers ?? valueColumns.map((_column, index) => `Value ${index + 1}`);
    for (let group = 0; group < colGroups.length + (showRowTotals ? 1 : 0); group += 1) header.push(...names);
    output.push(header);
  }
  const bodyRow = (rowGroup: number | null, fields: Scalar[]): Scalar[] => {
    const row: Scalar[] = [...fields];
    for (let colGroup = 0; colGroup < colGroups.length; colGroup += 1) {
      const rows = rowGroup === null ? colGroups[colGroup].rows.filter(include) : buckets.get(rowGroup * colCount + colGroup) ?? [];
      for (let value = 0; value < valueCount; value += 1) row.push(cell(value, rows, rowGroup, colGroup));
    }
    if (showRowTotals) {
      const rows = rowGroup === null ? allRows : rowGroups[rowGroup].rows;
      for (let value = 0; value < valueCount; value += 1) row.push(cell(value, rows, rowGroup, null));
    }
    return row;
  };
  const grandTotal = () => {
    const fields: Scalar[] = new Array(rowFieldCount).fill("");
    fields[0] = "Total";
    return bodyRow(null, fields);
  };
  if (rowTotalDepth < 0) output.push(grandTotal());
  rowGroups.forEach((group, index) => output.push(bodyRow(index, group.fields)));
  if (showColTotals && rowTotalDepth > 0) output.push(grandTotal());
  if (tooManyCells(output.length, width)) return valueError();
  return rowsResult(output);
}

// ---- TRIMRANGE --------------------------------------------------------------------------

function trimBounds(
  isBlank: (row: number, column: number) => boolean,
  rows: number,
  columns: number,
  trimRows: number,
  trimColumns: number,
): { top: number; bottom: number; left: number; right: number } | null {
  let top = 0;
  let bottom = rows - 1;
  let left = 0;
  let right = columns - 1;
  const rowBlank = (row: number) => {
    for (let column = left; column <= right; column += 1) if (!isBlank(row, column)) return false;
    return true;
  };
  const columnBlank = (column: number) => {
    for (let row = top; row <= bottom; row += 1) if (!isBlank(row, column)) return false;
    return true;
  };
  if (trimRows === 1 || trimRows === 3) while (top <= bottom && rowBlank(top)) top += 1;
  if (trimRows === 2 || trimRows === 3) while (bottom >= top && rowBlank(bottom)) bottom -= 1;
  if (top > bottom) return null;
  if (trimColumns === 1 || trimColumns === 3) while (left <= right && columnBlank(left)) left += 1;
  if (trimColumns === 2 || trimColumns === 3) while (right >= left && columnBlank(right)) right -= 1;
  if (left > right) return null;
  return { top, bottom, left, right };
}

function trimRangeBounds(call: Call): RectangleBounds | EvaluationError | null {
  const bounds = call.referenceBounds(call.argumentNodes[0]);
  if (bounds === null || isEvaluationError(bounds)) return bounds;
  const trimRows = optionalIntegerArg(lazyArg(call, 1), 3);
  if (isEvaluationError(trimRows)) return trimRows;
  const trimColumns = optionalIntegerArg(lazyArg(call, 2), 3);
  if (isEvaluationError(trimColumns)) return trimColumns;
  if (trimRows < 0 || trimRows > 3 || trimColumns < 0 || trimColumns > 3) return valueError();
  const visit = call.hooks.forEachCellInRange;
  const rows = bounds.lastRow - bounds.firstRow + 1;
  const columns = bounds.lastColumn - bounds.firstColumn + 1;
  let extent: { top: number; bottom: number; left: number; right: number } | null;
  if (visit && tooManyCells(rows, columns)) {
    // Sparse sheets: only populated cells matter, so take their bounding box.
    let top = Infinity;
    let bottom = -Infinity;
    let left = Infinity;
    let right = -Infinity;
    visit(
      bounds.sheetId,
      { startRow: bounds.firstRow, endRow: bounds.lastRow, startColumn: bounds.firstColumn, endColumn: bounds.lastColumn },
      (row, column) => {
        top = Math.min(top, row);
        bottom = Math.max(bottom, row);
        left = Math.min(left, column);
        right = Math.max(right, column);
      },
    );
    if (top === Infinity) return refErrorValue();
    extent = {
      top: trimRows === 1 || trimRows === 3 ? top - bounds.firstRow : 0,
      bottom: trimRows === 2 || trimRows === 3 ? bottom - bounds.firstRow : rows - 1,
      left: trimColumns === 1 || trimColumns === 3 ? left - bounds.firstColumn : 0,
      right: trimColumns === 2 || trimColumns === 3 ? right - bounds.firstColumn : columns - 1,
    };
  } else {
    const value = call.resolveBounds(bounds);
    if (isEvaluationError(value)) return value;
    if (!isEvaluationRange(value) || value.sparse) {
      return isEvaluationRange(value) ? valueError() : bounds;
    }
    extent = trimBounds(
      (row, column) => {
        const entry = value.values[row * value.columnCount + column];
        return entry === null || entry === "";
      },
      value.rowCount,
      value.columnCount,
      trimRows,
      trimColumns,
    );
    if (!extent) return refErrorValue();
  }
  return {
    sheetId: bounds.sheetId,
    firstRow: bounds.firstRow + extent.top,
    lastRow: bounds.firstRow + extent.bottom,
    firstColumn: bounds.firstColumn + extent.left,
    lastColumn: bounds.firstColumn + extent.right,
  };
}

function refErrorValue(): EvaluationError {
  return { kind: "evaluationError", code: "#REF!" };
}

// ---- Registry ---------------------------------------------------------------------------

export const ARRAY_FUNCTIONS: Specs = {
  SORTBY: spec(
    2,
    253,
    (values) => {
      const source = rectArg(values[0]);
      if (isEvaluationError(source)) return source;
      const keys: Array<{ values: Scalar[]; order: number }> = [];
      let byColumns: boolean | null = null;
      for (let index = 1; index < values.length; index += 2) {
        const by = rectArg(values[index]);
        if (isEvaluationError(by)) return by;
        let columns: boolean;
        if (by.columnCount === 1 && by.rowCount === source.rowCount) columns = false;
        else if (by.rowCount === 1 && by.columnCount === source.columnCount) columns = true;
        else return valueError();
        if (byColumns === null) byColumns = columns;
        else if (byColumns !== columns && !(by.rowCount === 1 && by.columnCount === 1)) return valueError();
        const order = optionalIntegerArg(values[index + 1], 1);
        if (isEvaluationError(order)) return order;
        if (order !== 1 && order !== -1) return valueError();
        keys.push({ values: by.values, order });
      }
      let rows = rectangleRows(source);
      if (byColumns) rows = transposedRows(rows);
      const indices = rows.map((_row, index) => index);
      indices.sort((left, right) => {
        for (const key of keys) {
          const a = key.values[left];
          const b = key.values[right];
          const aBlank = a === null;
          const bBlank = b === null;
          if (aBlank || bBlank) {
            if (aBlank !== bBlank) return aBlank ? 1 : -1;
            continue;
          }
          const comparison = sortScalarCompare(a, b) * key.order;
          if (comparison !== 0) return comparison;
        }
        return 0;
      });
      const sorted = indices.map((index) => rows[index]);
      return rowsResult(byColumns ? transposedRows(sorted) : sorted);
    },
    { returnsArray: true },
  ),

  XMATCH: spec(
    2,
    4,
    (values) => {
      const key = scalarArgument(values[0]);
      if (isEvaluationError(key)) return key;
      const lookup = rectArg(values[1]);
      if (isEvaluationError(lookup)) return lookup;
      if (lookup.rowCount > 1 && lookup.columnCount > 1) return valueError();
      const matchMode = optionalIntegerArg(values[2], 0);
      if (isEvaluationError(matchMode)) return matchMode;
      if (![0, -1, 1, 2, 3].includes(matchMode)) return valueError();
      const searchMode = optionalIntegerArg(values[3], 1);
      if (isEvaluationError(searchMode)) return searchMode;
      if (![1, -1, 2, -2].includes(searchMode)) return valueError();
      const index = extendedMatch(lookup.values, key, matchMode, searchMode);
      if (isEvaluationError(index)) return index;
      return index < 0 ? naError() : index + 1;
    },
    { liftArgs: [0] },
  ),

  LOOKUP: spec(
    2,
    3,
    (values) => {
      const key = scalarArgument(values[0]);
      if (isEvaluationError(key)) return key;
      const lookup = rectArg(values[1]);
      if (isEvaluationError(lookup)) return lookup;
      let candidates: Scalar[];
      let results: Scalar[];
      if (values.length >= 3) {
        if (lookup.rowCount > 1 && lookup.columnCount > 1) return naError();
        const result = rectArg(values[2]);
        if (isEvaluationError(result)) return result;
        if (result.rowCount > 1 && result.columnCount > 1) return naError();
        candidates = lookup.values;
        results = result.values;
      } else if (lookup.columnCount > lookup.rowCount) {
        candidates = rowValues(lookup, 0);
        results = rowValues(lookup, lookup.rowCount - 1);
      } else {
        candidates = columnValues(lookup, 0);
        results = columnValues(lookup, lookup.columnCount - 1);
      }
      const index = approximateIndex(candidates, key);
      if (index < 0 || index >= results.length) return naError();
      return results[index];
    },
    { liftArgs: [0] },
  ),

  RANDARRAY: spec(
    0,
    5,
    (values) => {
      const rows = optionalIntegerArg(values[0], 1);
      if (isEvaluationError(rows)) return rows;
      const columns = optionalIntegerArg(values[1], 1);
      if (isEvaluationError(columns)) return columns;
      const min = optionalNumberArg(values[2], 0);
      if (isEvaluationError(min)) return min;
      const max = optionalNumberArg(values[3], 1);
      if (isEvaluationError(max)) return max;
      const integer = optionalBooleanArg(values[4], false);
      if (isEvaluationError(integer)) return integer;
      if (rows < 0 || columns < 0) return valueError();
      if (rows === 0 || columns === 0) return calcError();
      if (min > max) return valueError();
      if (integer && (!Number.isInteger(min) || !Number.isInteger(max))) return valueError();
      if (tooManyCells(rows, columns)) return valueError();
      const output: Scalar[] = new Array(rows * columns);
      for (let index = 0; index < output.length; index += 1) {
        output[index] = integer ? min + Math.floor(Math.random() * (max - min + 1)) : min + Math.random() * (max - min);
      }
      return arrayResult(output, rows, columns);
    },
    { returnsArray: true, volatile: true },
  ),

  CHOOSECOLS: spec(
    2,
    254,
    (values) => {
      const source = rectArg(values[0]);
      if (isEvaluationError(source)) return source;
      const picks = indexNumbers(values.slice(1));
      if (isEvaluationError(picks)) return picks;
      const columns: number[] = [];
      for (const pick of picks) {
        if (pick === 0 || Math.abs(pick) > source.columnCount) return valueError();
        columns.push(pick > 0 ? pick - 1 : source.columnCount + pick);
      }
      if (tooManyCells(source.rowCount, columns.length)) return valueError();
      const output: Scalar[] = [];
      for (let row = 0; row < source.rowCount; row += 1) {
        for (const column of columns) output.push(source.values[row * source.columnCount + column]);
      }
      return arrayResult(output, source.rowCount, columns.length);
    },
    { returnsArray: true },
  ),

  CHOOSEROWS: spec(
    2,
    254,
    (values) => {
      const source = rectArg(values[0]);
      if (isEvaluationError(source)) return source;
      const picks = indexNumbers(values.slice(1));
      if (isEvaluationError(picks)) return picks;
      const output: Scalar[] = [];
      for (const pick of picks) {
        if (pick === 0 || Math.abs(pick) > source.rowCount) return valueError();
        const row = pick > 0 ? pick - 1 : source.rowCount + pick;
        for (let column = 0; column < source.columnCount; column += 1) {
          output.push(source.values[row * source.columnCount + column]);
        }
      }
      if (tooManyCells(picks.length, source.columnCount)) return valueError();
      return arrayResult(output, picks.length, source.columnCount);
    },
    { returnsArray: true },
  ),

  TAKE: spec(
    2,
    3,
    (values) => {
      const source = rectArg(values[0]);
      if (isEvaluationError(source)) return source;
      const rows = optionalIntegerArg(values[1], source.rowCount);
      if (isEvaluationError(rows)) return rows;
      const columns = optionalIntegerArg(values[2], source.columnCount);
      if (isEvaluationError(columns)) return columns;
      if (rows === 0 || columns === 0) return calcError();
      const rowCount = Math.min(Math.abs(rows), source.rowCount);
      const columnCount = Math.min(Math.abs(columns), source.columnCount);
      const firstRow = rows > 0 ? 0 : source.rowCount - rowCount;
      const firstColumn = columns > 0 ? 0 : source.columnCount - columnCount;
      const output: Scalar[] = [];
      for (let row = firstRow; row < firstRow + rowCount; row += 1) {
        for (let column = firstColumn; column < firstColumn + columnCount; column += 1) {
          output.push(source.values[row * source.columnCount + column]);
        }
      }
      return arrayResult(output, rowCount, columnCount);
    },
    { returnsArray: true },
  ),

  DROP: spec(
    2,
    3,
    (values) => {
      const source = rectArg(values[0]);
      if (isEvaluationError(source)) return source;
      const rows = optionalIntegerArg(values[1], 0);
      if (isEvaluationError(rows)) return rows;
      const columns = optionalIntegerArg(values[2], 0);
      if (isEvaluationError(columns)) return columns;
      const rowCount = source.rowCount - Math.abs(rows);
      const columnCount = source.columnCount - Math.abs(columns);
      if (rowCount <= 0 || columnCount <= 0) return calcError();
      const firstRow = rows > 0 ? rows : 0;
      const firstColumn = columns > 0 ? columns : 0;
      const output: Scalar[] = [];
      for (let row = firstRow; row < firstRow + rowCount; row += 1) {
        for (let column = firstColumn; column < firstColumn + columnCount; column += 1) {
          output.push(source.values[row * source.columnCount + column]);
        }
      }
      return arrayResult(output, rowCount, columnCount);
    },
    { returnsArray: true },
  ),

  VSTACK: spec(
    1,
    254,
    (values) => {
      const parts: Rect[] = [];
      for (const value of values) {
        const part = rectArg(value);
        if (isEvaluationError(part)) return part;
        parts.push(part);
      }
      const width = Math.max(...parts.map((part) => part.columnCount));
      const height = parts.reduce((sum, part) => sum + part.rowCount, 0);
      if (tooManyCells(height, width)) return valueError();
      const output: Scalar[] = [];
      for (const part of parts) {
        for (let row = 0; row < part.rowCount; row += 1) {
          for (let column = 0; column < width; column += 1) {
            output.push(column < part.columnCount ? part.values[row * part.columnCount + column] : naError());
          }
        }
      }
      return arrayResult(output, height, width);
    },
    { returnsArray: true },
  ),

  HSTACK: spec(
    1,
    254,
    (values) => {
      const parts: Rect[] = [];
      for (const value of values) {
        const part = rectArg(value);
        if (isEvaluationError(part)) return part;
        parts.push(part);
      }
      const height = Math.max(...parts.map((part) => part.rowCount));
      const width = parts.reduce((sum, part) => sum + part.columnCount, 0);
      if (tooManyCells(height, width)) return valueError();
      const output: Scalar[] = [];
      for (let row = 0; row < height; row += 1) {
        for (const part of parts) {
          for (let column = 0; column < part.columnCount; column += 1) {
            output.push(row < part.rowCount ? part.values[row * part.columnCount + column] : naError());
          }
        }
      }
      return arrayResult(output, height, width);
    },
    { returnsArray: true },
  ),

  TOCOL: spec(1, 3, (values) => flattenSpec(values, "column"), { returnsArray: true }),
  TOROW: spec(1, 3, (values) => flattenSpec(values, "row"), { returnsArray: true }),

  WRAPROWS: spec(2, 3, (values) => wrapSpec(values, "rows"), { returnsArray: true }),
  WRAPCOLS: spec(2, 3, (values) => wrapSpec(values, "columns"), { returnsArray: true }),

  EXPAND: spec(
    2,
    4,
    (values) => {
      const source = rectArg(values[0]);
      if (isEvaluationError(source)) return source;
      const rows = optionalIntegerArg(values[1], source.rowCount);
      if (isEvaluationError(rows)) return rows;
      const columns = optionalIntegerArg(values[2], source.columnCount);
      if (isEvaluationError(columns)) return columns;
      if (rows < source.rowCount || columns < source.columnCount) return valueError();
      if (tooManyCells(rows, columns)) return valueError();
      const pad = padValue(values[3]);
      const output: Scalar[] = [];
      for (let row = 0; row < rows; row += 1) {
        for (let column = 0; column < columns; column += 1) {
          output.push(
            row < source.rowCount && column < source.columnCount ? source.values[row * source.columnCount + column] : pad,
          );
        }
      }
      return arrayResult(output, rows, columns);
    },
    { returnsArray: true },
  ),

  TRIMRANGE: {
    minArgs: 1,
    maxArgs: 3,
    lazy: true,
    returnsArray: true,
    reference: trimRangeBounds,
    impl: (_values, call) => {
      const bounds = trimRangeBounds(call);
      if (bounds === null) {
        // Not a reference: trim the computed array instead.
        const source = rectArg(call.evaluate(call.argumentNodes[0]));
        if (isEvaluationError(source)) return source;
        const trimRows = optionalIntegerArg(lazyArg(call, 1), 3);
        if (isEvaluationError(trimRows)) return trimRows;
        const trimColumns = optionalIntegerArg(lazyArg(call, 2), 3);
        if (isEvaluationError(trimColumns)) return trimColumns;
        const extent = trimBounds(
          (row, column) => {
            const entry = source.values[row * source.columnCount + column];
            return entry === null || entry === "";
          },
          source.rowCount,
          source.columnCount,
          trimRows,
          trimColumns,
        );
        if (!extent) return refErrorValue();
        const output: Scalar[] = [];
        for (let row = extent.top; row <= extent.bottom; row += 1) {
          for (let column = extent.left; column <= extent.right; column += 1) output.push(source.values[row * source.columnCount + column]);
        }
        return arrayResult(output, extent.bottom - extent.top + 1, extent.right - extent.left + 1);
      }
      if (isEvaluationError(bounds)) return bounds;
      return call.resolveBounds(bounds);
    },
  },

  GROUPBY: { minArgs: 3, maxArgs: 8, lazy: true, returnsArray: true, impl: (_values, call) => evaluateGroupBy(call) },
  PIVOTBY: { minArgs: 4, maxArgs: 11, lazy: true, returnsArray: true, impl: (_values, call) => evaluatePivotBy(call) },
  PERCENTOF: spec(2, 2, (values) => {
    const subset = rectArg(values[0]);
    if (isEvaluationError(subset)) return subset;
    const all = rectArg(values[1]);
    if (isEvaluationError(all)) return all;
    return percentOf(
      { kind: "evaluationRange", ...subset },
      { kind: "evaluationRange", ...all },
    );
  }),

  ADDRESS: spec(
    2,
    5,
    (values) => {
      const row = integerArg(values[0]);
      if (isEvaluationError(row)) return row;
      const column = integerArg(values[1]);
      if (isEvaluationError(column)) return column;
      const absolute = optionalIntegerArg(values[2], 1);
      if (isEvaluationError(absolute)) return absolute;
      const a1 = optionalBooleanArg(values[3], true);
      if (isEvaluationError(a1)) return a1;
      const sheet = values[4] === undefined || values[4] === null ? "" : textArg(values[4]);
      if (isEvaluationError(sheet)) return sheet;
      if (row < 1 || row > MAX_EXCEL_ROW || column < 1 || column > MAX_EXCEL_COLUMN) return valueError();
      if (absolute < 1 || absolute > 4) return valueError();
      const absoluteRow = absolute === 1 || absolute === 2;
      const absoluteColumn = absolute === 1 || absolute === 3;
      const reference = a1
        ? `${absoluteColumn ? "$" : ""}${columnNumberToLabel(column)}${absoluteRow ? "$" : ""}${row}`
        : `${absoluteRow ? `R${row}` : `R[${row}]`}${absoluteColumn ? `C${column}` : `C[${column}]`}`;
      return sheet ? `${quoteSheetName(sheet)}!${reference}` : reference;
    },
    { liftArgs: "all" },
  ),

  AREAS: {
    minArgs: 1,
    maxArgs: 1,
    lazy: true,
    impl: (_values, call) => {
      const bounds = call.referenceBounds(call.argumentNodes[0]);
      if (bounds === null) return valueError();
      return isEvaluationError(bounds) ? bounds : 1;
    },
  },

  HYPERLINK: spec(
    1,
    2,
    (values) => {
      const link = scalarArgument(values[0]);
      if (isEvaluationError(link)) return link;
      if (values[1] === undefined) return link === null ? 0 : link;
      const friendly = scalarArgument(values[1]);
      return friendly === null ? 0 : friendly;
    },
    { liftArgs: "all" },
  ),

  CELL: {
    minArgs: 1,
    maxArgs: 2,
    lazy: true,
    volatile: true,
    impl: (_values, call) => {
      const infoType = textArg(call.evaluate(call.argumentNodes[0]));
      if (isEvaluationError(infoType)) return infoType;
      let bounds: RectangleBounds;
      if (!isOmittedNode(call.argumentNodes[1])) {
        const reference = call.referenceBounds(call.argumentNodes[1]);
        if (reference === null) return valueError();
        if (isEvaluationError(reference)) return reference;
        bounds = reference;
      } else {
        const cell = call.hooks.currentCell;
        if (!cell) return valueError();
        bounds = { sheetId: call.currentSheetId, firstRow: cell.row, lastRow: cell.row, firstColumn: cell.column, lastColumn: cell.column };
      }
      const topLeft: RectangleBounds = { ...bounds, lastRow: bounds.firstRow, lastColumn: bounds.firstColumn };
      const sheetName = call.hooks.getSheetName?.(bounds.sheetId) ?? bounds.sheetId;
      const readValue = (): Scalar => {
        const value = call.resolveBounds(topLeft);
        if (isEvaluationRange(value)) return value.values[0] ?? null;
        return isLambdaValue(value) ? valueError() : value;
      };
      switch (infoType.trim().toLowerCase()) {
        case "address": {
          const address = `$${columnNumberToLabel(bounds.firstColumn)}$${bounds.firstRow}`;
          return bounds.sheetId.toLowerCase() === call.currentSheetId.toLowerCase() ? address : `${quoteSheetName(sheetName)}!${address}`;
        }
        case "row":
          return bounds.firstRow;
        case "col":
          return bounds.firstColumn;
        case "contents": {
          const value = readValue();
          return value;
        }
        case "type": {
          const value = readValue();
          if (value === null) return "b";
          return typeof value === "string" ? "l" : "v";
        }
        case "filename":
          return `[Book1]${sheetName}`;
        case "format":
          return "G";
        case "parentheses":
        case "color":
          return 0;
        case "prefix":
          return "";
        case "protect":
          return 1;
        case "width":
          return 8;
        default:
          return valueError();
      }
    },
  },

  COUNTUNIQUE: spec(1, 255, (values) => {
    const seen = new Set<string>();
    for (const entry of collectValues(values)) {
      const value = entry.value;
      if (isEvaluationError(value)) return value;
      if (value === null || value === "") continue;
      // Google Sheets compares text case-sensitively here.
      seen.add(typeof value === "string" ? `s:${value}` : arrayEntryKey(value));
    }
    return seen.size;
  }),

  FLATTEN: spec(
    1,
    255,
    (values) => {
      const output: Scalar[] = [];
      for (const value of values) {
        const range = rectArg(value);
        if (isEvaluationError(range)) return range;
        output.push(...range.values);
      }
      if (output.length > 100_000) return valueError();
      return arrayResult(output, output.length, 1);
    },
    { returnsArray: true },
  ),

  ARRAYFORMULA: spec(1, 1, (values) => values[0] ?? null, { returnsArray: true }),

  ARRAY_CONSTRAIN: spec(
    3,
    3,
    (values) => {
      const source = rectArg(values[0]);
      if (isEvaluationError(source)) return source;
      const rows = integerArg(values[1]);
      if (isEvaluationError(rows)) return rows;
      const columns = integerArg(values[2]);
      if (isEvaluationError(columns)) return columns;
      if (rows < 1 || columns < 1) return numError();
      const rowCount = Math.min(rows, source.rowCount);
      const columnCount = Math.min(columns, source.columnCount);
      const output: Scalar[] = [];
      for (let row = 0; row < rowCount; row += 1) {
        for (let column = 0; column < columnCount; column += 1) output.push(source.values[row * source.columnCount + column]);
      }
      return arrayResult(output, rowCount, columnCount);
    },
    { returnsArray: true },
  ),

  SORTN: spec(
    1,
    255,
    (values) => {
      const source = rectArg(values[0]);
      if (isEvaluationError(source)) return source;
      const count = optionalIntegerArg(values[1], 1);
      if (isEvaluationError(count)) return count;
      const tiesMode = optionalIntegerArg(values[2], 0);
      if (isEvaluationError(tiesMode)) return tiesMode;
      if (count < 0 || tiesMode < 0 || tiesMode > 3) return valueError();
      const rows = rectangleRows(source);
      const keys: Array<{ values: Scalar[]; direction: number }> = [];
      for (let index = 3; index < values.length; index += 2) {
        const column = values[index];
        const ascending = optionalBooleanArg(values[index + 1], true);
        if (isEvaluationError(ascending)) return ascending;
        const direction = ascending ? 1 : -1;
        if (isEvaluationRange(column) && column.values.length > 1) {
          const range = rectArg(column);
          if (isEvaluationError(range)) return range;
          if (range.values.length !== rows.length) return valueError();
          keys.push({ values: range.values, direction });
        } else {
          const position = integerArg(column);
          if (isEvaluationError(position)) return position;
          if (position < 1 || position > source.columnCount) return valueError();
          keys.push({ values: rows.map((row) => row[position - 1]), direction });
        }
      }
      if (keys.length === 0) keys.push({ values: rows.map((row) => row[0]), direction: 1 });
      const compare = (left: number, right: number) => {
        for (const key of keys) {
          const comparison = sortScalarCompare(key.values[left], key.values[right]) * key.direction;
          if (comparison !== 0) return comparison;
        }
        return 0;
      };
      const order = rows.map((_row, index) => index).sort(compare);
      const rowKey = (index: number) => keys.map((key) => arrayEntryKey(key.values[index])).join("\u0001");
      const kept: number[] = [];
      if (tiesMode === 0) kept.push(...order.slice(0, count));
      else if (tiesMode === 1) {
        for (const index of order) {
          if (kept.length < count || (kept.length > 0 && compare(kept[kept.length - 1], index) === 0)) kept.push(index);
          else break;
        }
      } else {
        // 2: the first n rows after removing duplicates; 3: every row of the first n unique keys.
        const allowed = new Set<string>();
        for (const index of order) {
          const key = rowKey(index);
          if (allowed.has(key)) {
            if (tiesMode === 3) kept.push(index);
            continue;
          }
          if (allowed.size >= count) break;
          allowed.add(key);
          kept.push(index);
        }
      }
      if (kept.length === 0) return calcError();
      return rowsResult(kept.map((index) => rows[index]));
    },
    { returnsArray: true },
  ),
};

function flattenSpec(values: Value[], shape: "column" | "row"): Value {
  const source = rectArg(values[0]);
  if (isEvaluationError(source)) return source;
  const ignore = optionalIntegerArg(values[1], 0);
  if (isEvaluationError(ignore)) return ignore;
  if (ignore < 0 || ignore > 3) return valueError();
  const byColumn = optionalBooleanArg(values[2], false);
  if (isEvaluationError(byColumn)) return byColumn;
  const skipBlanks = ignore === 1 || ignore === 3;
  const skipErrors = ignore === 2 || ignore === 3;
  const output: Scalar[] = [];
  const visit = (value: Scalar) => {
    if (skipBlanks && value === null) return;
    if (skipErrors && isEvaluationError(value)) return;
    output.push(value);
  };
  if (byColumn) {
    for (let column = 0; column < source.columnCount; column += 1) {
      for (let row = 0; row < source.rowCount; row += 1) visit(source.values[row * source.columnCount + column]);
    }
  } else {
    for (const value of source.values) visit(value);
  }
  if (output.length === 0) return calcError();
  return shape === "column" ? arrayResult(output, output.length, 1) : arrayResult(output, 1, output.length);
}

function wrapSpec(values: Value[], shape: "rows" | "columns"): Value {
  const source = rectArg(values[0]);
  if (isEvaluationError(source)) return source;
  if (source.rowCount > 1 && source.columnCount > 1) return valueError();
  const wrap = integerArg(values[1]);
  if (isEvaluationError(wrap)) return wrap;
  if (wrap < 1) return numError();
  const pad = padValue(values[2]);
  const count = source.values.length;
  const other = Math.ceil(count / wrap);
  const rowCount = shape === "rows" ? other : wrap;
  const columnCount = shape === "rows" ? wrap : other;
  if (tooManyCells(rowCount, columnCount)) return valueError();
  const output: Scalar[] = new Array(rowCount * columnCount);
  for (let row = 0; row < rowCount; row += 1) {
    for (let column = 0; column < columnCount; column += 1) {
      const index = shape === "rows" ? row * wrap + column : column * wrap + row;
      output[row * columnCount + column] = index < count ? source.values[index] : pad;
    }
  }
  return arrayResult(output, rowCount, columnCount);
}
