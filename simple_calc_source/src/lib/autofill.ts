import type { CellData } from "../spreadsheet-types";
import { shiftFormulaReferences } from "./formulas";
import { serialFromYMD, ymdFromSerial } from "./data-tools-core";

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

/**
 * How a fill is built, as Excel's fill handle and its Auto Fill Options button choose:
 *
 * - `auto`: the fill handle's default (a lone number copies; a lone date, a day/month name,
 *   a quarter or "Item 1" continues as a series; two or more seeds extend their step).
 * - `toggle`: Ctrl+drag, which swaps copying and continuing a series.
 * - `copy`: Copy Cells (also Fill Down / Fill Right): seeds repeat, formulas shift.
 * - `series`: Fill Series (a lone number counts up by 1).
 * - `formats`: Fill Formatting Only (destination contents stay).
 * - `values`: Fill Without Formatting (destination formatting stays).
 * - `days` / `weekdays` / `months` / `years`: date steps for date seeds.
 */
export type AutofillMode =
  | "auto"
  | "toggle"
  | "copy"
  | "series"
  | "formats"
  | "values"
  | "days"
  | "weekdays"
  | "months"
  | "years";

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
  /** Defaults to `auto`. */
  mode?: AutofillMode;
  /** The workbook uses the 1904 date system (month and weekday arithmetic). */
  date1904?: boolean;
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
  /** The Auto Fill Options that apply to this fill, in Excel's menu order. */
  options: AutofillMode[];
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

interface TrendSeries {
  kind: "trend";
  /** Value at logical index 0 and the change per index (least-squares line). */
  intercept: number;
  slope: number;
}

interface MonthSeries {
  kind: "months";
  first: number;
  /** Whole months per index (12 for a yearly series). */
  step: number;
}

interface WeekdaySeries {
  kind: "weekdays";
  first: number;
  /** Working days per index. */
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

type Casing = "upper" | "lower" | "title" | "as-is";

interface ListSeries {
  kind: "list";
  items: readonly string[];
  first: number;
  step: number;
  casing: Casing;
}

interface RepeatingSeries {
  kind: "repeat";
}

type LaneSeries =
  | NumericSeries
  | TrendSeries
  | MonthSeries
  | WeekdaySeries
  | TextSuffixSeries
  | ListSeries
  | RepeatingSeries;

interface ParsedNumericSuffix {
  prefix: string;
  suffix: string;
  value: number;
  digits: string;
  explicitPlus: boolean;
}

/** Excel's built-in fill lists (Tools > Custom Lists): checked in this order. */
export const BUILTIN_FILL_LISTS: ReadonlyArray<readonly string[]> = [
  ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
  ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
  ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
];

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
  delete cell.arrayMember;
  delete (cell as CellData & { sharedFormulaMaster?: string }).sharedFormulaMaster;
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

/** A calendar date format (not a pure time or duration format such as h:mm or [h]:mm). */
function isCalendarDateFormat(format: string | undefined): boolean {
  if (!isDateLikeNumberFormat(format)) return false;
  const visible = stripNumberFormatLiterals(String(format)).toLocaleLowerCase();
  return /(^|[^a-z])(?:y{1,4}|d{1,4})(?=$|[^a-z])/.test(visible) || /(^|[^a-z])m{3,5}(?=$|[^a-z])/.test(visible);
}

function isDateCell(cell: CellData): boolean {
  return (
    cell.type === "date" ||
    cell.resultType === "date" ||
    isDateLikeNumberFormat(cell.numFmt ?? cell.style?.numFmt)
  );
}

/** Date cells whose day/month/year steps are meaningful (times of day step linearly). */
function isCalendarDateCell(cell: CellData): boolean {
  return cell.type === "date" || cell.resultType === "date" || isCalendarDateFormat(cell.numFmt ?? cell.style?.numFmt);
}

function nearlyEqual(left: number, right: number): boolean {
  const scale = Math.max(1, Math.abs(left), Math.abs(right));
  return Math.abs(left - right) <= Number.EPSILON * 64 * scale;
}

/** Excel stores 15 significant digits: 0.1 + 0.2 fills as 0.3, not 0.30000000000000004. */
function tidy(value: number): number {
  if (!Number.isFinite(value) || value === 0) return value;
  return Number(value.toPrecision(15));
}

function arithmeticStep(values: readonly number[]): number | null {
  if (values.length < 2) return null;
  const step = values[1] - values[0];
  for (let index = 2; index < values.length; index += 1) {
    if (!nearlyEqual(values[index] - values[index - 1], step)) return null;
  }
  return step;
}

/** Excel's fill handle extends a best-fit (least-squares) line through uneven seeds. */
function linearTrend(values: readonly number[]): TrendSeries {
  const count = values.length;
  const meanX = (count - 1) / 2;
  const meanY = values.reduce((sum, value) => sum + value, 0) / count;
  let numerator = 0;
  let denominator = 0;
  values.forEach((value, index) => {
    numerator += (index - meanX) * (value - meanY);
    denominator += (index - meanX) ** 2;
  });
  const slope = denominator ? numerator / denominator : 0;
  return { kind: "trend", intercept: meanY - slope * meanX, slope };
}

function parseNumericSuffix(value: string): ParsedNumericSuffix | null {
  const match = /^(.*?)(\d+)(\s*)$/.exec(value);
  if (!match) return null;
  // A sign only belongs to the number in text that is just a signed number ("+5", "-12");
  // in "A-009" or "INV-0042" the dash separates the code from its number.
  const sign = /^\s*[+-]$/.test(match[1]) ? match[1].trim() : "";
  const numericValue = Number(`${sign}${match[2]}`);
  if (!Number.isSafeInteger(numericValue)) return null;
  return {
    prefix: sign ? match[1].slice(0, match[1].length - 1) : match[1],
    suffix: match[3],
    value: numericValue,
    digits: match[2],
    explicitPlus: sign === "+",
  };
}

/**
 * "Item 1" continues as "Item 2", "Item 3" (a lone seed counts up by one, as in Excel);
 * two or more seeds with the same text around an arithmetic number keep their step.
 */
function inferTextSuffixSeries(entries: readonly LaneEntry[]): TextSuffixSeries | null {
  if (!entries.length) return null;
  const parsed = entries.map(({ cell }) =>
    typeof cell.value === "string" && !cell.formula ? parseNumericSuffix(cell.value) : null,
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

  const step = items.length === 1 ? 1 : arithmeticStep(items.map((item) => item.value));
  if (step === null || !Number.isSafeInteger(step) || step === 0) return null;
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

function casingOf(text: string): Casing {
  const letters = text.replace(/[^A-Za-zÀ-ÿ]/g, "");
  if (!letters) return "as-is";
  if (letters === letters.toLocaleUpperCase() && letters.length > 1) return "upper";
  if (letters === letters.toLocaleLowerCase()) return "lower";
  const head = letters[0];
  const tail = letters.slice(1);
  if (head === head.toLocaleUpperCase() && tail === tail.toLocaleLowerCase()) return "title";
  return "as-is";
}

function applyCasing(item: string, casing: Casing): string {
  if (casing === "upper") return item.toLocaleUpperCase();
  if (casing === "lower") return item.toLocaleLowerCase();
  if (casing === "title") return item.charAt(0).toLocaleUpperCase() + item.slice(1).toLocaleLowerCase();
  return item;
}

function ordinalSuffix(value: number): string {
  return value === 1 ? "st" : value === 2 ? "nd" : value === 3 ? "rd" : "th";
}

/**
 * The four items of a quarter series written like `text` ("Q1", "Qtr 2", "Quarter 3",
 * "4th Quarter"), and the seed's position, or null.
 */
function quarterList(text: string): { items: string[]; index: number } | null {
  const named = /^(Q|Qtr\.?|Quarter)(\s*)([1-4])$/i.exec(text);
  if (named) {
    const items = [1, 2, 3, 4].map((quarter) => `${named[1]}${named[2]}${quarter}`);
    return { items, index: Number(named[3]) - 1 };
  }
  const ordinal = /^([1-4])(st|nd|rd|th)(\s*)(Quarter|Qtr\.?|Q)$/i.exec(text);
  if (ordinal) {
    const upper = ordinal[2] === ordinal[2].toLocaleUpperCase();
    const items = [1, 2, 3, 4].map((quarter) => {
      const suffix = ordinalSuffix(quarter);
      return `${quarter}${upper ? suffix.toLocaleUpperCase() : suffix}${ordinal[3]}${ordinal[4]}`;
    });
    return { items, index: Number(ordinal[1]) - 1 };
  }
  return null;
}

/** Day and month names (full or three-letter, any case) and quarters continue around their list. */
function inferListSeries(entries: readonly LaneEntry[]): ListSeries | null {
  if (!entries.length) return null;
  const texts: string[] = [];
  for (const { cell } of entries) {
    if (cell.formula || typeof cell.value !== "string") return null;
    const text = cell.value.trim();
    if (!text) return null;
    texts.push(text);
  }
  const casing = casingOf(texts[0]);
  const fromPositions = (items: readonly string[], positions: number[], listCasing: Casing): ListSeries | null => {
    const size = items.length;
    const step = positions.length === 1 ? 1 : positiveModulo(positions[1] - positions[0], size);
    if (step === 0) return null;
    for (let index = 1; index < positions.length; index += 1) {
      if (positiveModulo(positions[index] - positions[index - 1], size) !== step) return null;
    }
    // Two seeds that step backwards (Wed, Mon) keep going backwards.
    const signed = step > size / 2 && positions.length > 1 ? step - size : step;
    return { kind: "list", items, first: positions[0], step: signed, casing: listCasing };
  };
  const quarters = texts.map(quarterList);
  if (quarters.every(Boolean)) {
    const items = quarters[0]!.items;
    const sameSpelling = quarters.every((item) => item!.items.join("|").toLocaleLowerCase() === items.join("|").toLocaleLowerCase());
    if (sameSpelling) return fromPositions(items, quarters.map((item) => item!.index), "as-is");
    return null;
  }
  for (const list of BUILTIN_FILL_LISTS) {
    const lower = list.map((item) => item.toLocaleLowerCase());
    const positions = texts.map((text) => lower.indexOf(text.toLocaleLowerCase()));
    if (positions.some((position) => position < 0)) continue;
    return fromPositions(list, positions, casing);
  }
  return null;
}

// ---- Calendar arithmetic (serials) ----------------------------------------------------------

function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/** Excel's EDATE: the same day `months` later, clamped to the month's end (time kept). */
export function addMonthsToSerial(serial: number, months: number, date1904 = false): number {
  const day = Math.floor(serial);
  const time = serial - day;
  const { year, month, day: dayOfMonth } = ymdFromSerial(day, date1904);
  const total = year * 12 + (month - 1) + months;
  const nextYear = Math.floor(total / 12);
  const nextMonth = positiveModulo(total, 12) + 1;
  return serialFromYMD(nextYear, nextMonth, Math.min(dayOfMonth, daysInMonth(nextYear, nextMonth)), date1904) + time;
}

/** Months from `from` to `to` when `to` is exactly EDATE(from, n); otherwise null. */
function wholeMonthsBetween(from: number, to: number, date1904: boolean): number | null {
  if (!nearlyEqual(from - Math.floor(from), to - Math.floor(to))) return null;
  const a = ymdFromSerial(Math.floor(from), date1904);
  const b = ymdFromSerial(Math.floor(to), date1904);
  const months = (b.year - a.year) * 12 + (b.month - a.month);
  return nearlyEqual(addMonthsToSerial(from, months, date1904), to) ? months : null;
}

function isWeekend(serial: number, date1904: boolean): boolean {
  const { weekday } = ymdFromSerial(Math.floor(serial), date1904);
  return weekday === 0 || weekday === 6;
}

/** Excel's WORKDAY without holidays: `count` working days after (or before) `serial`. */
export function addWeekdaysToSerial(serial: number, count: number, date1904 = false): number {
  let current = serial;
  const direction = count < 0 ? -1 : 1;
  let remaining = Math.abs(Math.trunc(count));
  while (remaining > 0) {
    current += direction;
    if (!isWeekend(current, date1904)) remaining -= 1;
  }
  return current;
}

function weekdaysBetween(from: number, to: number, date1904: boolean): number | null {
  if (isWeekend(from, date1904) || isWeekend(to, date1904)) return null;
  const direction = to >= from ? 1 : -1;
  let count = 0;
  for (let current = from; direction > 0 ? current < to : current > to; current += direction) {
    if (!isWeekend(current + direction, date1904)) count += direction;
    if (Math.abs(count) > 10_000) return null;
  }
  return count || null;
}

// ---- Lane inference --------------------------------------------------------------------------

interface LaneFacts {
  formula: boolean;
  numbers: number[] | null;
  dates: boolean;
}

function laneFacts(entries: readonly LaneEntry[]): LaneFacts {
  const formula = entries.some(({ cell }) => Boolean(cell.formula));
  const values = entries.map(({ cell }) => cell.value);
  const numbers = !formula && values.length > 0 && values.every((value) => typeof value === "number" && Number.isFinite(value))
    ? (values as number[])
    : null;
  return { formula, numbers, dates: Boolean(numbers) && entries.every(({ cell }) => isCalendarDateCell(cell)) };
}

function numericSeries(values: readonly number[], singleStep: number | null): LaneSeries {
  if (values.length === 1) return singleStep === null ? { kind: "repeat" } : { kind: "numeric", first: values[0], step: singleStep };
  const step = arithmeticStep(values);
  if (step !== null) return step === 0 ? { kind: "repeat" } : { kind: "numeric", first: values[0], step };
  return linearTrend(values);
}

/** Two or more dates a whole number of months apart (1/15, 2/15 or 1/31, 2/28) step by months. */
function monthSeries(values: readonly number[], date1904: boolean): MonthSeries | null {
  if (values.length < 2) return null;
  const step = wholeMonthsBetween(values[0], values[1], date1904);
  if (step === null || step === 0) return null;
  for (let index = 2; index < values.length; index += 1) {
    if (!nearlyEqual(addMonthsToSerial(values[0], step * index, date1904), values[index])) return null;
  }
  return { kind: "months", first: values[0], step };
}

/** A lone time of day (h:mm, [h]:mm) steps by an hour, as in Excel; null when not a time. */
function timeStep(entries: readonly LaneEntry[]): number | null {
  return entries.every(({ cell }) => isDateCell(cell)) ? 1 / 24 : null;
}

/** What the fill handle does by default (Excel). */
function inferAutoSeries(entries: readonly LaneEntry[], facts: LaneFacts, date1904: boolean): LaneSeries {
  // Formula lanes copy their seed formulas instead of extrapolating cached
  // results. Mixed formula/literal lanes repeat the source pattern as well.
  if (facts.formula) return { kind: "repeat" };
  if (facts.numbers) {
    // A lone date advances a day; two or more dates whole months apart step by months.
    if (facts.dates) return monthSeries(facts.numbers, date1904) ?? numericSeries(facts.numbers, 1);
    // A lone number copies (Ctrl+drag or Fill Series counts it up); a lone time adds an hour.
    return numericSeries(facts.numbers, timeStep(entries));
  }
  return inferListSeries(entries) ?? inferTextSuffixSeries(entries) ?? { kind: "repeat" };
}

/** Fill Series: like the default, except that a lone number counts up by one. */
function inferSeries(entries: readonly LaneEntry[], facts: LaneFacts, date1904: boolean): LaneSeries {
  if (facts.formula) return { kind: "repeat" };
  if (facts.numbers && !facts.dates) return numericSeries(facts.numbers, timeStep(entries) ?? 1);
  return inferAutoSeries(entries, facts, date1904);
}

function inferDateSeries(mode: "days" | "weekdays" | "months" | "years", entries: readonly LaneEntry[], facts: LaneFacts, date1904: boolean): LaneSeries {
  if (!facts.numbers || !facts.dates || facts.formula) return inferAutoSeries(entries, facts, date1904);
  const values = facts.numbers;
  if (mode === "days") return numericSeries(values, 1);
  if (mode === "weekdays") {
    const step = values.length > 1 ? weekdaysBetween(values[0], values[1], date1904) ?? 1 : 1;
    return { kind: "weekdays", first: values[0], step };
  }
  const months = values.length > 1 ? wholeMonthsBetween(values[0], values[1], date1904) : null;
  if (mode === "months") return { kind: "months", first: values[0], step: months || 1 };
  return { kind: "months", first: values[0], step: months && months % 12 === 0 ? months : 12 };
}

function inferLaneSeries(entries: readonly LaneEntry[], mode: AutofillMode, date1904: boolean): LaneSeries {
  const facts = laneFacts(entries);
  switch (mode) {
    case "copy":
    case "formats":
      return { kind: "repeat" };
    case "series":
      return inferSeries(entries, facts, date1904);
    case "toggle": {
      // Ctrl+drag: a lane that would copy becomes a series, and a series becomes a copy.
      const automatic = inferAutoSeries(entries, facts, date1904);
      return automatic.kind === "repeat" ? inferSeries(entries, facts, date1904) : { kind: "repeat" };
    }
    case "days":
    case "weekdays":
    case "months":
    case "years":
      return inferDateSeries(mode, entries, facts, date1904);
    default:
      return inferAutoSeries(entries, facts, date1904);
  }
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

function seriesValue(series: LaneSeries, logicalIndex: number, date1904: boolean): CellData["value"] | undefined {
  switch (series.kind) {
    case "numeric":
      return tidy(series.first + series.step * logicalIndex);
    case "trend":
      return tidy(series.intercept + series.slope * logicalIndex);
    case "months":
      return addMonthsToSerial(series.first, series.step * logicalIndex, date1904);
    case "weekdays":
      return addWeekdaysToSerial(series.first, series.step * logicalIndex, date1904);
    case "text-suffix":
      return formatTextSuffix(series, logicalIndex);
    case "list":
      return applyCasing(series.items[positiveModulo(series.first + series.step * logicalIndex, series.items.length)], series.casing);
    default:
      return undefined;
  }
}

const FORMAT_KEYS = ["style", "numFmt"] as const;

function withoutFormats(cell: CellData | undefined): CellData {
  const output = cell ? cloneCell(cell) : {};
  for (const key of FORMAT_KEYS) delete output[key];
  return output;
}

function filledCell(
  entries: readonly LaneEntry[],
  series: LaneSeries,
  logicalIndex: number,
  destination: AutofillCoordinate,
  mode: AutofillMode,
  existing: CellData | undefined,
  date1904: boolean,
): CellData | null {
  const templateIndex = positiveModulo(logicalIndex, entries.length);
  const template = entries[templateIndex];

  if (mode === "formats") {
    // Fill Formatting Only: the destination keeps its contents and takes the seed's look.
    const output = withoutFormats(existing);
    if (template.cell.style) output.style = structuredClone(template.cell.style);
    if (template.cell.numFmt) output.numFmt = template.cell.numFmt;
    return hasCellPayload(output) ? output : null;
  }

  const output = cloneCell(template.cell);
  clearStaleCalculationMetadata(output);

  if (output.formula) {
    delete output.value;
    output.formula = shiftFormulaReferences(
      output.formula,
      destination.row - template.coord.row,
      destination.col - template.coord.col,
    );
  } else {
    const value = seriesValue(series, logicalIndex, date1904);
    if (value !== undefined) {
      delete output.formula;
      delete output.richText;
      output.value = value;
    }
  }

  if (mode === "values") {
    // Fill Without Formatting: contents only; the destination keeps its own look.
    for (const key of FORMAT_KEYS) delete output[key];
    if (existing?.style) output.style = structuredClone(existing.style);
    if (existing?.numFmt) output.numFmt = existing.numFmt;
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
 * - Two or more numeric seeds extend their step (uneven seeds follow Excel's
 *   best-fit line). A lone date advances one day; a lone ordinary number repeats
 *   unless the mode asks for a series (Fill Series, Ctrl+drag).
 * - Dates a whole number of months apart step by months (or years).
 * - Day and month names (full or short, matching case) and quarters (Q1, Qtr 1,
 *   Quarter 1, 1st Quarter) continue around their list.
 * - Text ending in a number ("Item 1", "A-009") counts up, keeping zero padding.
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
  const mode = request.mode ?? "auto";
  const date1904 = Boolean(request.date1904);
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
  let dateLanes = false;

  for (let row = destination.top; row <= destination.bottom; row += 1) {
    for (let col = destination.left; col <= destination.right; col += 1) {
      const laneKey = vertical ? col : row;
      let lane = laneCache.get(laneKey);
      if (!lane) {
        const entries = vertical
          ? verticalLane(cells, source, col)
          : horizontalLane(cells, source, row);
        lane = { entries, series: inferLaneSeries(entries, mode, date1904) };
        if (laneFacts(entries).dates) dateLanes = true;
        laneCache.set(laneKey, lane);
      }

      const logicalIndex = vertical ? row - source.top : col - source.left;
      const coord = { row, col };
      const address = addressOf(coord);
      changes[address] = filledCell(
        lane.entries,
        lane.series,
        logicalIndex,
        coord,
        mode,
        cells[address],
        date1904,
      );
    }
  }

  const options: AutofillMode[] = ["copy", "series", "formats", "values"];
  if (dateLanes) options.push("days", "weekdays", "months", "years");
  return {
    direction,
    source: { ...source },
    destination: { ...destination },
    changes,
    options,
  };
}

/** Menu labels for the Auto Fill Options button. */
export const AUTOFILL_MODE_LABELS: Record<Exclude<AutofillMode, "auto" | "toggle">, string> = {
  copy: "Copy cells",
  series: "Fill series",
  formats: "Fill formatting only",
  values: "Fill without formatting",
  days: "Fill days",
  weekdays: "Fill weekdays",
  months: "Fill months",
  years: "Fill years",
};

/**
 * Excel's double-click on the fill handle: the last row to fill down to, following the
 * contiguous data in the column left of the selection (or, when that is empty, the column to
 * its right). The fill stops above the first cell in the filled columns that already holds
 * data. Returns null when there is nothing to fill.
 */
export function fillHandleDoubleClickBottom(
  hasContent: (row: number, col: number) => boolean,
  selection: AutofillRectangle,
  options: { maxRow?: number } = {},
): number | null {
  const maxRow = Math.min(MAX_EXCEL_ROWS - 1, options.maxRow ?? MAX_EXCEL_ROWS - 1);
  const start = selection.bottom + 1;
  if (start > maxRow) return null;
  const neighbours = [selection.left - 1, selection.right + 1].filter((col) => col >= 0 && col < MAX_EXCEL_COLUMNS);
  let bottom: number | null = null;
  for (const col of neighbours) {
    if (!hasContent(start, col)) continue;
    let row = start;
    while (row + 1 <= maxRow && hasContent(row + 1, col)) row += 1;
    bottom = row;
    break;
  }
  if (bottom === null) return null;
  for (let row = start; row <= bottom; row += 1) {
    for (let col = selection.left; col <= selection.right; col += 1) {
      if (hasContent(row, col)) return row > start ? row - 1 : null;
    }
  }
  return bottom;
}
