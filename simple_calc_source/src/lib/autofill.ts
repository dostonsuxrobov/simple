import type { CellData } from "../spreadsheet-types";
import { shiftFormulaReferences } from "./formulas";

const MAX_EXCEL_ROWS = 1_048_576;
const MAX_EXCEL_COLUMNS = 16_384;
const DEFAULT_MAX_PATCH_CELLS = 100_000;

/** A zero-based spreadsheet coordinate. */
export interface AutofillCoordinate {
  row: number;
  col: number;
}

/**
 * An inclusive, zero-based rectangle. `top <= bottom` and `left <= right` are
 * required by {@link createAutofillPatch}.
 */
export interface AutofillRectangle {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export type AutofillDirection = "up" | "down" | "left" | "right";

/** Input for a copy-handle fill operation. */
export interface AutofillRequest {
  /** Cells from the active sheet, keyed by A1 address. */
  cells: Readonly<Record<string, CellData | undefined>>;
  /** The selected seed cells. */
  source: AutofillRectangle;
  /**
   * The adjacent extension to fill. This rectangle excludes `source`: for
   * example, dragging A1:A2 down through A5 uses A3:A5 as `destination`.
   */
  destination: AutofillRectangle;
  /** Safety limit for generated changes. Defaults to 100,000 cells. */
  maxCells?: number;
}

/** The result of an autofill operation, ready to apply in one workbook edit. */
export interface AutofillPatch {
  direction: AutofillDirection;
  source: AutofillRectangle;
  destination: AutofillRectangle;
  /**
   * Destination changes keyed by A1 address. `null` means delete/clear that
   * cell; a style-only seed remains a non-null style-only cell.
   */
  changes: Record<string, CellData | null>;
}

interface LaneEntry {
  coord: AutofillCoordinate;
  cell: CellData;
}

interface NumericSeries {
  kind: "numeric";
  first: number;
  step: number;
}

interface TextSuffixSeries {
  kind: "text-suffix";
  prefix: string;
  suffix: string;
  first: number;
  step: number;
  digitWidth: number;
  padWithZeroes: boolean;
  explicitPlus: boolean;
}

interface RepeatingSeries {
  kind: "repeat";
}

type LaneSeries = NumericSeries | TextSuffixSeries | RepeatingSeries;

interface ParsedNumericSuffix {
  prefix: string;
  suffix: string;
  value: number;
  digits: string;
  explicitPlus: boolean;
}

function assertRectangle(rectangle: AutofillRectangle, label: string): void {
  const values = [rectangle.top, rectangle.bottom, rectangle.left, rectangle.right];
  if (!values.every((value) => Number.isSafeInteger(value))) {
    throw new TypeError(`${label} must contain safe integer coordinates.`);
  }
  if (
    rectangle.top < 0 ||
    rectangle.left < 0 ||
    rectangle.top > rectangle.bottom ||
    rectangle.left > rectangle.right
  ) {
    throw new RangeError(`${label} must be a non-empty, ordered rectangle.`);
  }
  if (rectangle.bottom >= MAX_EXCEL_ROWS || rectangle.right >= MAX_EXCEL_COLUMNS) {
    throw new RangeError(`${label} exceeds the XLSX worksheet limits.`);
  }
}

function rectangleSize(rectangle: AutofillRectangle): number {
  return (
    (rectangle.bottom - rectangle.top + 1) *
    (rectangle.right - rectangle.left + 1)
  );
}

function detectDirection(
  source: AutofillRectangle,
  destination: AutofillRectangle,
): AutofillDirection {
  const sameColumns =
    source.left === destination.left && source.right === destination.right;
  const sameRows =
    source.top === destination.top && source.bottom === destination.bottom;

  if (sameColumns && destination.top === source.bottom + 1) return "down";
  if (sameColumns && destination.bottom === source.top - 1) return "up";
  if (sameRows && destination.left === source.right + 1) return "right";
  if (sameRows && destination.right === source.left - 1) return "left";

  throw new RangeError(
    "destination must be directly adjacent to source and share its columns (vertical fill) or rows (horizontal fill).",
  );
}

function columnName(column: number): string {
  let value = column + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function addressOf(coord: AutofillCoordinate): string {
  return `${columnName(coord.col)}${coord.row + 1}`;
}

function positiveModulo(value: number, divisor: number): number {
  return ((value % divisor) + divisor) % divisor;
}

function cloneCell(cell: CellData): CellData {
  return structuredClone(cell) as CellData;
}

function cellAt(
  cells: Readonly<Record<string, CellData | undefined>>,
  coord: AutofillCoordinate,
): CellData {
  return cells[addressOf(coord)] ?? {};
}

function hasCellPayload(cell: CellData): boolean {
  return Object.keys(cell).length > 0;
}

/**
 * Remove values that belong to the old formula calculation or array/shared
 * formula structure. A newly filled cell must be recalculated in its new
 * position and must not remain attached to its seed's formula range.
 */
function clearStaleCalculationMetadata(cell: CellData): void {
  delete cell.formulaType;
  delete cell.formulaRange;
  delete cell.dynamicFormula;
  delete cell.result;
  delete cell.resultType;
  delete cell.display;
}

function stripNumberFormatLiterals(format: string): string {
  let result = "";
  let position = 0;

  while (position < format.length) {
    const character = format[position];
    if (character === '"') {
      position += 1;
      while (position < format.length) {
        if (format[position] !== '"') position += 1;
        else if (format[position + 1] === '"') position += 2;
        else {
          position += 1;
          break;
        }
      }
      continue;
    }

    // Excel uses backslash, underscore, and asterisk to escape or lay out the
    // next character. That escaped character must not look like a date token.
    if (character === "\\" || character === "_" || character === "*") {
      position += 2;
      continue;
    }

    if (character === "[") {
      const end = format.indexOf("]", position + 1);
      if (end >= 0) {
        const bracket = format.slice(position + 1, end).trim();
        // Elapsed-time formats ([h], [mm], [ss]) are date/time formats. Other
        // brackets are colors, conditions, locales, or currency declarations.
        if (/^[hms]+$/i.test(bracket)) result += `[${bracket}]`;
        position = end + 1;
        continue;
      }
    }

    result += character;
    position += 1;
  }

  return result;
}

/**
 * Return whether an Excel/Sheets number-format code represents a calendar date
 * or time. Quoted words, colors, conditions, locales, and escaped characters
 * are ignored so formats such as `0 \"days\"` are not mistaken for dates.
 */
export function isDateLikeNumberFormat(format: string | undefined): boolean {
  if (!format || format.trim().toLocaleLowerCase() === "general") return false;
  const visible = stripNumberFormatLiterals(format).toLocaleLowerCase();
  if (/\[(?:h+|m+|s+)\]/.test(visible)) return true;
  if (/am\s*\/\s*pm|a\s*\/\s*p/.test(visible)) return true;

  // Date/time letters are meaningful as runs in spreadsheet format strings.
  // Checking complete runs avoids interpreting scientific notation's E token.
  return /(^|[^a-z])(?:y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2})(?=$|[^a-z])/.test(
    visible,
  );
}

function isDateCell(cell: CellData): boolean {
  return (
    cell.type === "date" ||
    cell.resultType === "date" ||
    isDateLikeNumberFormat(cell.numFmt ?? cell.style?.numFmt)
  );
}

function nearlyEqual(left: number, right: number): boolean {
  const scale = Math.max(1, Math.abs(left), Math.abs(right));
  return Math.abs(left - right) <= Number.EPSILON * 64 * scale;
}

function arithmeticStep(values: readonly number[]): number | null {
  if (values.length < 2) return null;
  const step = values[1] - values[0];
  for (let index = 2; index < values.length; index += 1) {
    if (!nearlyEqual(values[index] - values[index - 1], step)) return null;
  }
  return step;
}

function parseNumericSuffix(value: string): ParsedNumericSuffix | null {
  const match = /^(.*?)([+-]?)(\d+)(\s*)$/.exec(value);
  if (!match) return null;
  const numericValue = Number(`${match[2]}${match[3]}`);
  if (!Number.isSafeInteger(numericValue)) return null;
  return {
    prefix: match[1],
    suffix: match[4],
    value: numericValue,
    digits: match[3],
    explicitPlus: match[2] === "+",
  };
}

function inferTextSuffixSeries(entries: readonly LaneEntry[]): TextSuffixSeries | null {
  // A single text value copies as-is. Two seeds opt into a numeric-suffix
  // sequence ("Item 1", "Item 2"), matching spreadsheet fill-handle behavior.
  if (entries.length < 2) return null;
  const parsed = entries.map(({ cell }) =>
    typeof cell.value === "string" ? parseNumericSuffix(cell.value) : null,
  );
  if (parsed.some((item) => item === null)) return null;
  const items = parsed as ParsedNumericSuffix[];
  const first = items[0];
  if (
    !items.every(
      (item) => item.prefix === first.prefix && item.suffix === first.suffix,
    )
  ) {
    return null;
  }

  const step = arithmeticStep(items.map((item) => item.value));
  if (step === null || !Number.isSafeInteger(step)) return null;
  const digitWidth = Math.max(...items.map((item) => item.digits.length));
  const padWithZeroes = items.some(
    (item) => item.digits.length > 1 && item.digits.startsWith("0"),
  );
  return {
    kind: "text-suffix",
    prefix: first.prefix,
    suffix: first.suffix,
    first: first.value,
    step,
    digitWidth,
    padWithZeroes,
    explicitPlus: items.every((item) => item.explicitPlus),
  };
}

function inferLaneSeries(entries: readonly LaneEntry[]): LaneSeries {
  // Formula lanes copy their seed formulas instead of extrapolating cached
  // results. Mixed formula/literal lanes repeat the source pattern as well.
  if (entries.some(({ cell }) => Boolean(cell.formula))) return { kind: "repeat" };

  const numericValues = entries.map(({ cell }) => cell.value);
  if (
    numericValues.length > 0 &&
    numericValues.every((value) => typeof value === "number" && Number.isFinite(value))
  ) {
    const values = numericValues as number[];
    const allDates = entries.every(({ cell }) => isDateCell(cell));
    if (values.length === 1 && allDates) {
      return { kind: "numeric", first: values[0], step: 1 };
    }
    const step = arithmeticStep(values);
    if (step !== null) return { kind: "numeric", first: values[0], step };
  }

  return inferTextSuffixSeries(entries) ?? { kind: "repeat" };
}

function formatTextSuffix(series: TextSuffixSeries, logicalIndex: number): string {
  const value = series.first + series.step * logicalIndex;
  if (!Number.isSafeInteger(value)) {
    // This can only occur after extrapolating beyond the safe-integer range.
    // Keeping a deterministic string is safer than silently losing precision.
    return `${series.prefix}${String(value)}${series.suffix}`;
  }
  const magnitude = String(Math.abs(value));
  const digits = series.padWithZeroes
    ? magnitude.padStart(series.digitWidth, "0")
    : magnitude;
  const sign = value < 0 ? "-" : series.explicitPlus ? "+" : "";
  return `${series.prefix}${sign}${digits}${series.suffix}`;
}

function filledCell(
  entries: readonly LaneEntry[],
  series: LaneSeries,
  logicalIndex: number,
  destination: AutofillCoordinate,
): CellData | null {
  const templateIndex = positiveModulo(logicalIndex, entries.length);
  const template = entries[templateIndex];
  const output = cloneCell(template.cell);
  clearStaleCalculationMetadata(output);

  if (output.formula) {
    delete output.value;
    output.formula = shiftFormulaReferences(
      output.formula,
      destination.row - template.coord.row,
      destination.col - template.coord.col,
    );
  } else if (series.kind === "numeric") {
    delete output.formula;
    output.value = series.first + series.step * logicalIndex;
  } else if (series.kind === "text-suffix") {
    delete output.formula;
    output.value = formatTextSuffix(series, logicalIndex);
  }

  return hasCellPayload(output) ? output : null;
}

function verticalLane(
  cells: Readonly<Record<string, CellData | undefined>>,
  source: AutofillRectangle,
  col: number,
): LaneEntry[] {
  const result: LaneEntry[] = [];
  for (let row = source.top; row <= source.bottom; row += 1) {
    const coord = { row, col };
    result.push({ coord, cell: cellAt(cells, coord) });
  }
  return result;
}

function horizontalLane(
  cells: Readonly<Record<string, CellData | undefined>>,
  source: AutofillRectangle,
  row: number,
): LaneEntry[] {
  const result: LaneEntry[] = [];
  for (let col = source.left; col <= source.right; col += 1) {
    const coord = { row, col };
    result.push({ coord, cell: cellAt(cells, coord) });
  }
  return result;
}

/**
 * Build the rectangular workbook patch produced by dragging a selection's fill
 * handle vertically or horizontally.
 *
 * - Each column is an independent series for vertical fills; each row is an
 *   independent series for horizontal fills.
 * - Two or more arithmetic numeric seeds are extrapolated. A lone date-formatted
 *   serial advances one day; a lone ordinary number repeats.
 * - Two or more matching text values with arithmetic integer suffixes are
 *   extrapolated, including zero-padded suffixes.
 * - Other values repeat cyclically. Formula references shift from the repeated
 *   seed cell to each destination, honoring `$` anchors.
 * - Entire cells are cloned, preserving formatting, notes, and links, while
 *   cached formula results and array/shared-formula metadata are removed.
 *
 * The function is pure: neither the request nor its source cells are mutated.
 * Apply each non-null `changes[address]` to the sheet and delete addresses whose
 * change is `null`.
 */
export function createAutofillPatch(request: AutofillRequest): AutofillPatch {
  const { cells, source, destination } = request;
  assertRectangle(source, "source");
  assertRectangle(destination, "destination");
  const direction = detectDirection(source, destination);
  const total = rectangleSize(destination);
  const maxCells = request.maxCells ?? DEFAULT_MAX_PATCH_CELLS;
  if (!Number.isSafeInteger(maxCells) || maxCells <= 0) {
    throw new RangeError("maxCells must be a positive safe integer.");
  }
  if (total > maxCells) {
    throw new RangeError(
      `Autofill would change ${total.toLocaleString()} cells, exceeding the ${maxCells.toLocaleString()}-cell limit.`,
    );
  }

  const changes: Record<string, CellData | null> = {};
  const vertical = direction === "up" || direction === "down";
  const laneCache = new Map<number, { entries: LaneEntry[]; series: LaneSeries }>();

  for (let row = destination.top; row <= destination.bottom; row += 1) {
    for (let col = destination.left; col <= destination.right; col += 1) {
      const laneKey = vertical ? col : row;
      let lane = laneCache.get(laneKey);
      if (!lane) {
        const entries = vertical
          ? verticalLane(cells, source, col)
          : horizontalLane(cells, source, row);
        lane = { entries, series: inferLaneSeries(entries) };
        laneCache.set(laneKey, lane);
      }

      const logicalIndex = vertical ? row - source.top : col - source.left;
      const coord = { row, col };
      changes[addressOf(coord)] = filledCell(
        lane.entries,
        lane.series,
        logicalIndex,
        coord,
      );
    }
  }

  return {
    direction,
    source: { ...source },
    destination: { ...destination },
    changes,
  };
}

