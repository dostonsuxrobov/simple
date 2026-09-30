// Date worksheet functions for the extended library (see formula-library.ts):
// DAYS360, YEARFRAC, ISOWEEKNUM, NETWORKDAYS.INTL, WORKDAY.INTL, EPOCHTODATE, TO_DATE.
// Also exports the day-count helpers (civil dates, 30/360, YEARFRAC) used by the
// financial module's bond and depreciation functions.

import {
  collectValues,
  dowFromSerial,
  isEvaluationError,
  isEvaluationRange,
  MAX_DATE_SERIAL,
  scalarArgument,
  toNumber,
} from "./formulas";
import type { EvaluationError, EvaluationScalar } from "./formulas";
import { integerArg, numberArg, numError, spec, valueError } from "./formula-lib-shared";
import type { Specs, Value } from "./formula-lib-shared";

// ---- Civil date arithmetic (proleptic Gregorian, serial 0 = 1899-12-30) ---------------------

export interface CivilDate {
  y: number;
  /** 1-12 */
  m: number;
  d: number;
}

function daysFromCivil(year: number, month: number, day: number): number {
  const y = year - (month <= 2 ? 1 : 0);
  const era = Math.floor(y / 400);
  const yoe = y - era * 400;
  const mp = (month + 9) % 12;
  const doy = Math.floor((153 * mp + 2) / 5) + day - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146097 + doe - 719468;
}

const EPOCH_DAYS = daysFromCivil(1899, 12, 30);

/** Excel serial (integer part) to a civil date. */
export function civilFromSerial(serial: number): CivilDate {
  const z = Math.floor(serial) + EPOCH_DAYS + 719468;
  const era = Math.floor(z / 146097);
  const doe = z - era * 146097;
  const yoe = Math.floor((doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365);
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const d = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const m = mp < 10 ? mp + 3 : mp - 9;
  const y = yoe + era * 400 + (m <= 2 ? 1 : 0);
  return { y, m, d };
}

/** Civil date to Excel serial; out-of-range days/months roll over like DATE(). */
export function serialFromCivil(year: number, month: number, day: number): number {
  const totalMonths = year * 12 + (month - 1);
  const y = Math.floor(totalMonths / 12);
  const m = totalMonths - y * 12 + 1;
  return daysFromCivil(y, m, 1) - EPOCH_DAYS + (day - 1);
}

export function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

export function daysInMonthOf(year: number, month: number): number {
  if (month === 2) return isLeapYear(year) ? 29 : 28;
  return [4, 6, 9, 11].includes(month) ? 30 : 31;
}

function isLastDayOfMonth(date: CivilDate): boolean {
  return date.d === daysInMonthOf(date.y, date.m);
}

/**
 * A date argument: numbers or date text, truncated to a whole day. Negative or beyond
 * 9999-12-31 is #NUM!.
 */
export function dateArg(value: Value | undefined, fallback?: number): number | EvaluationError {
  const number = numberArg(value, fallback);
  if (isEvaluationError(number)) return number;
  const serial = Math.floor(number);
  if (serial < 0 || serial > MAX_DATE_SERIAL) return numError();
  return serial;
}

/** DAYS360: US (NASD, Excel's variant) or European 30/360 day difference; may be negative. */
export function days360(start: number, end: number, european: boolean): number {
  const a = civilFromSerial(start);
  const b = civilFromSerial(end);
  let day1 = a.d;
  let day2 = b.d;
  let month2 = b.m;
  let year2 = b.y;
  if (day1 === 31) day1 = 30;
  else if (!european && a.m === 2 && isLastDayOfMonth(a)) day1 = 30;
  if (day2 === 31) {
    if (!european && day1 !== 30) {
      day2 = 1;
      if (month2 === 12) {
        year2 += 1;
        month2 = 1;
      } else {
        month2 += 1;
      }
    } else {
      day2 = 30;
    }
  }
  return day2 + month2 * 30 + year2 * 360 - day1 - a.m * 30 - a.y * 360;
}

/** The 30/360 US variant YEARFRAC basis 0 (and the bond functions) use; start <= end. */
export function days360Nasd(start: number, end: number): number {
  const a = civilFromSerial(start);
  const b = civilFromSerial(end);
  let day1 = a.d;
  let day2 = b.d;
  if (day1 === 31 && day2 === 31) {
    day1 = 30;
    day2 = 30;
  } else if (day1 === 31) {
    day1 = 30;
  } else if (day1 === 30 && day2 === 31) {
    day2 = 30;
  } else if (a.m === 2 && b.m === 2 && isLastDayOfMonth(a) && isLastDayOfMonth(b)) {
    day1 = 30;
    day2 = 30;
  } else if (a.m === 2 && isLastDayOfMonth(a)) {
    day1 = 30;
  }
  return day2 + b.m * 30 + b.y * 360 - (day1 + a.m * 30 + a.y * 360);
}

/** European 30/360 (basis 4). */
export function days360European(start: number, end: number): number {
  return days360(start, end, true);
}

function appearsWithinYear(a: CivilDate, b: CivilDate): boolean {
  if (a.y === b.y) return true;
  return a.y + 1 === b.y && (a.m > b.m || (a.m === b.m && a.d >= b.d));
}

/** Excel's YEARFRAC (start/end serials, basis 0-4). */
export function yearFrac(startSerial: number, endSerial: number, basis: number): number {
  let start = Math.floor(startSerial);
  let end = Math.floor(endSerial);
  if (start === end) return 0;
  if (start > end) [start, end] = [end, start];
  switch (basis) {
    case 0:
      return days360Nasd(start, end) / 360;
    case 2:
      return (end - start) / 360;
    case 3:
      return (end - start) / 365;
    case 4:
      return days360European(start, end) / 360;
    default: {
      const a = civilFromSerial(start);
      const b = civilFromSerial(end);
      if (appearsWithinYear(a, b)) {
        let yearLength = 365;
        if (a.y === b.y && isLeapYear(a.y)) yearLength = 366;
        else if (
          (isLeapYear(a.y) && start < serialFromCivil(a.y, 3, 1)) ||
          (isLeapYear(b.y) && end >= serialFromCivil(b.y, 3, 1)) ||
          (b.m === 2 && b.d === 29)
        ) {
          yearLength = 366;
        }
        return (end - start) / yearLength;
      }
      const years = b.y - a.y + 1;
      const days = serialFromCivil(b.y + 1, 1, 1) - serialFromCivil(a.y, 1, 1);
      return (end - start) / (days / years);
    }
  }
}

/** Days in a year for a day-count basis (bond functions). */
export function basisYearDays(basis: number, year?: number): number {
  if (basis === 3) return 365;
  if (basis === 1) return year !== undefined && isLeapYear(year) ? 366 : 365;
  return 360;
}

/** Day count between two dates under a basis (30/360 variants or actual days). */
export function basisDayCount(start: number, end: number, basis: number): number {
  if (basis === 0) return start <= end ? days360Nasd(start, end) : -days360Nasd(end, start);
  if (basis === 4) return days360European(start, end);
  return end - start;
}

/** Read and validate a basis argument (default 0): integers 0-4, #NUM! otherwise. */
export function basisArg(value: Value | undefined): number | EvaluationError {
  const basis = integerArg(value, 0);
  if (isEvaluationError(basis)) return basis;
  return basis < 0 || basis > 4 ? numError() : basis;
}

// ---- Weekends and holidays -------------------------------------------------------------------

/** weekend[dow] (0 = Sunday) from a NETWORKDAYS.INTL/WORKDAY.INTL weekend argument. */
function weekendMask(value: Value | undefined): boolean[] | EvaluationError {
  const mask = [false, false, false, false, false, false, false];
  if (value === undefined || value === null) {
    mask[0] = true;
    mask[6] = true;
    return mask;
  }
  const scalar = scalarArgument(value);
  if (isEvaluationError(scalar)) return scalar;
  if (typeof scalar === "string") {
    if (!/^[01]{7}$/.test(scalar)) return valueError();
    for (let index = 0; index < 7; index += 1) {
      // The string starts on Monday.
      if (scalar[index] === "1") mask[(index + 1) % 7] = true;
    }
    return mask;
  }
  const code = toNumber(scalar);
  if (isEvaluationError(code)) return code;
  const whole = Math.trunc(code);
  if (whole >= 1 && whole <= 7) {
    mask[(whole + 5) % 7] = true;
    mask[(whole + 6) % 7] = true;
    return mask;
  }
  if (whole >= 11 && whole <= 17) {
    mask[whole - 11] = true;
    return mask;
  }
  return numError();
}

/** Sorted, de-duplicated holiday serials. Date text is coerced; other text is #VALUE!. */
function holidayList(value: Value | undefined): number[] | EvaluationError {
  if (value === undefined) return [];
  const serials = new Set<number>();
  for (const entry of collectValues([value])) {
    const item: EvaluationScalar = entry.value;
    if (isEvaluationError(item)) return item;
    if (item === null) continue;
    if (typeof item === "boolean") return valueError();
    const number = typeof item === "number" ? item : toNumber(item);
    if (isEvaluationError(number)) return valueError();
    const serial = Math.floor(number);
    if (serial < 0 || serial > MAX_DATE_SERIAL) return numError();
    serials.add(serial);
  }
  return [...serials].sort((left, right) => left - right);
}

function workdaysInSpan(first: number, last: number, weekend: boolean[]): number {
  // Inclusive span [first, last] with first <= last.
  const span = last - first + 1;
  const workdaysPerWeek = weekend.filter((flag) => !flag).length;
  const fullWeeks = Math.floor(span / 7);
  let count = fullWeeks * workdaysPerWeek;
  for (let day = first + fullWeeks * 7; day <= last; day += 1) {
    if (!weekend[dowFromSerial(day)]) count += 1;
  }
  return count;
}

function lowerBound(sorted: number[], target: number): number {
  let low = 0;
  let high = sorted.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (sorted[middle] < target) low = middle + 1;
    else high = middle;
  }
  return low;
}

function networkDays(
  startValue: Value | undefined,
  endValue: Value | undefined,
  weekendValue: Value | undefined,
  holidayValue: Value | undefined,
): number | EvaluationError {
  const start = dateArg(startValue);
  if (isEvaluationError(start)) return start;
  const end = dateArg(endValue);
  if (isEvaluationError(end)) return end;
  const weekend = weekendMask(weekendValue);
  if (isEvaluationError(weekend)) return weekend;
  const holidays = holidayList(holidayValue);
  if (isEvaluationError(holidays)) return holidays;
  const sign = start <= end ? 1 : -1;
  const first = Math.min(start, end);
  const last = Math.max(start, end);
  let count = workdaysInSpan(first, last, weekend);
  for (let index = lowerBound(holidays, first); index < holidays.length && holidays[index] <= last; index += 1) {
    if (!weekend[dowFromSerial(holidays[index])]) count -= 1;
  }
  return sign * count;
}

function workday(
  startValue: Value | undefined,
  daysValue: Value | undefined,
  weekendValue: Value | undefined,
  holidayValue: Value | undefined,
): number | EvaluationError {
  const start = dateArg(startValue);
  if (isEvaluationError(start)) return start;
  const rawDays = numberArg(daysValue);
  if (isEvaluationError(rawDays)) return rawDays;
  const weekend = weekendMask(weekendValue);
  if (isEvaluationError(weekend)) return weekend;
  const workdaysPerWeek = weekend.filter((flag) => !flag).length;
  if (workdaysPerWeek === 0) return valueError();
  const holidays = holidayList(holidayValue);
  if (isEvaluationError(holidays)) return holidays;
  const workingHolidays = holidays.filter((serial) => !weekend[dowFromSerial(serial)]);
  let remaining = Math.trunc(rawDays);
  if (remaining === 0) return start;
  const step = remaining > 0 ? 1 : -1;
  remaining = Math.abs(remaining);
  let current = start;
  const holidaysBetween = (from: number, to: number): number => {
    // Holidays strictly after `from` up to and including `to` (in travel direction).
    const low = Math.min(from, to);
    const high = Math.max(from, to);
    const lowIndex = step > 0 ? lowerBound(workingHolidays, low + 1) : lowerBound(workingHolidays, low);
    const highIndex = step > 0 ? lowerBound(workingHolidays, high + 1) : lowerBound(workingHolidays, high);
    return highIndex - lowIndex;
  };
  // Jump whole weeks: every 7 consecutive days hold exactly `workdaysPerWeek` workdays,
  // minus any (working-day) holidays inside the jump, which are added back to `remaining`.
  while (remaining > workdaysPerWeek) {
    const weeks = Math.floor((remaining - 1) / workdaysPerWeek);
    const next = current + step * weeks * 7;
    if (next < 0 || next > MAX_DATE_SERIAL) return numError();
    remaining -= weeks * workdaysPerWeek;
    remaining += holidaysBetween(current, next);
    current = next;
  }
  const holidaySet = new Set(workingHolidays);
  while (remaining > 0) {
    current += step;
    if (current < 0 || current > MAX_DATE_SERIAL) return numError();
    if (!weekend[dowFromSerial(current)] && !holidaySet.has(current)) remaining -= 1;
  }
  return current;
}

export const DATE_FUNCTIONS: Specs = {
  DAYS360: spec(
    2,
    3,
    (values) => {
      const start = dateArg(values[0]);
      if (isEvaluationError(start)) return start;
      const end = dateArg(values[1]);
      if (isEvaluationError(end)) return end;
      let european = false;
      if (values[2] !== undefined) {
        const scalar = scalarArgument(values[2]);
        if (isEvaluationError(scalar)) return scalar;
        const number = toNumber(scalar);
        if (isEvaluationError(number)) return number;
        european = number !== 0;
      }
      return days360(start, end, european);
    },
    { liftArgs: "all" },
  ),
  YEARFRAC: spec(
    2,
    3,
    (values) => {
      const start = dateArg(values[0]);
      if (isEvaluationError(start)) return start;
      const end = dateArg(values[1]);
      if (isEvaluationError(end)) return end;
      const basis = basisArg(values[2]);
      if (isEvaluationError(basis)) return basis;
      return yearFrac(start, end, basis);
    },
    { liftArgs: "all" },
  ),
  ISOWEEKNUM: spec(
    1,
    1,
    (values) => {
      const serial = dateArg(values[0]);
      if (isEvaluationError(serial)) return serial;
      // ISO weeks start on Monday; week 1 holds the year's first Thursday.
      const dow = (dowFromSerial(serial) + 6) % 7; // Monday = 0
      const thursday = serial - dow + 3;
      const year = civilFromSerial(thursday).y;
      const firstThursdayWeekStart = (() => {
        const jan4 = serialFromCivil(year, 1, 4);
        return jan4 - ((dowFromSerial(jan4) + 6) % 7);
      })();
      return Math.floor((serial - firstThursdayWeekStart) / 7) + 1;
    },
    { liftArgs: "all" },
  ),
  "NETWORKDAYS.INTL": spec(2, 4, (values) => networkDays(values[0], values[1], values[2], values[3]), {
    liftArgs: [0, 1],
  }),
  "WORKDAY.INTL": spec(2, 4, (values) => workday(values[0], values[1], values[2], values[3]), {
    liftArgs: [0, 1],
  }),
  EPOCHTODATE: spec(
    1,
    2,
    (values) => {
      const timestamp = numberArg(values[0]);
      if (isEvaluationError(timestamp)) return timestamp;
      const unit = integerArg(values[1], 1);
      if (isEvaluationError(unit)) return unit;
      const divisor = unit === 1 ? 86_400 : unit === 2 ? 86_400_000 : unit === 3 ? 86_400_000_000 : 0;
      if (!divisor || timestamp < 0) return numError();
      const serial = 25_569 + timestamp / divisor;
      return serial > MAX_DATE_SERIAL + 1 ? numError() : serial;
    },
    { liftArgs: "all" },
  ),
  TO_DATE: spec(
    1,
    1,
    (values) => {
      const value = scalarArgument(values[0]);
      if (isEvaluationRange(values[0]) && isEvaluationError(value)) return value;
      // Numbers are already date serials; Sheets returns anything else unchanged.
      return value;
    },
    { liftArgs: "all" },
  ),
};
