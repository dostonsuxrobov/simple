// Financial worksheet functions for the extended library (see formula-library.ts):
// interest-rate conversion, dated cash flows (XNPV/XIRR/MIRR), depreciation (SLN, SYD, DB,
// DDB, VDB, AMORLINC, AMORDEGRC), cumulative loan parts, dollar fractions, and coupon bonds,
// discounted securities and T-bills.

import {
  annuityFutureValue,
  annuityPayment,
  collectValues,
  finiteResult,
  isEvaluationError,
  toNumber,
} from "./formulas";
import type { EvaluationError } from "./formulas";
import {
  divError,
  numberArg,
  numError,
  readNumbers,
  rectArg,
  spec,
  valueError,
} from "./formula-lib-shared";
import type { Specs, Value } from "./formula-lib-shared";
import {
  basisDayCount,
  civilFromSerial,
  dateArg,
  daysInMonthOf,
  days360Nasd,
  serialFromCivil,
  yearFrac,
} from "./formula-lib-date";

type NumberOrError = number | EvaluationError;

/** Read positional numeric arguments (`defaults[i]` for arguments that were not passed). */
function readArgs(values: Value[], defaults: Array<number | undefined>): number[] | EvaluationError {
  const numbers: number[] = [];
  for (let index = 0; index < defaults.length; index += 1) {
    const number = numberArg(values[index], defaults[index]);
    if (isEvaluationError(number)) return number;
    numbers.push(number);
  }
  return numbers;
}

/**
 * A scalar financial function: all arguments lifted, read as numbers (defaults for optional
 * arguments that were not passed), numeric results checked for overflow.
 */
function financial(
  defaults: Array<number | undefined>,
  compute: (...args: number[]) => NumberOrError,
): Specs[string] {
  let minArgs = 0;
  while (minArgs < defaults.length && defaults[minArgs] === undefined) minArgs += 1;
  return spec(
    minArgs,
    defaults.length,
    (values) => {
      const args = readArgs(values, defaults);
      if (isEvaluationError(args)) return args;
      const result = compute(...args);
      return typeof result === "number" ? finiteResult(result) : result;
    },
    { liftArgs: "all" },
  );
}

/** A date argument that arrived as a number: whole days, 0..9999-12-31. */
function wholeDate(value: number): NumberOrError {
  return dateArg(value);
}

function validBasis(value: number): NumberOrError {
  const basis = Math.trunc(value);
  return basis < 0 || basis > 4 ? numError() : basis;
}

function validFrequency(value: number): NumberOrError {
  const frequency = Math.trunc(value);
  return frequency === 1 || frequency === 2 || frequency === 4 ? frequency : numError();
}

interface BondTerms {
  settlement: number;
  maturity: number;
  frequency: number;
  basis: number;
}

function bondTerms(settlementValue: number, maturityValue: number, frequencyValue: number, basisValue: number): BondTerms | EvaluationError {
  const settlement = wholeDate(settlementValue);
  if (isEvaluationError(settlement)) return settlement;
  const maturity = wholeDate(maturityValue);
  if (isEvaluationError(maturity)) return maturity;
  const frequency = validFrequency(frequencyValue);
  if (isEvaluationError(frequency)) return frequency;
  const basis = validBasis(basisValue);
  if (isEvaluationError(basis)) return basis;
  if (settlement >= maturity) return numError();
  return { settlement, maturity, frequency, basis };
}

function datedTerms(settlementValue: number, maturityValue: number, basisValue: number): BondTerms | EvaluationError {
  return bondTerms(settlementValue, maturityValue, 1, basisValue);
}

// ---- Coupon schedule -------------------------------------------------------------------------

/** Coupon dates step back from maturity; an end-of-month maturity keeps coupons at month end. */
function couponDateBefore(maturity: number, periods: number, frequency: number): number {
  const date = civilFromSerial(maturity);
  const endOfMonth = date.d === daysInMonthOf(date.y, date.m);
  const totalMonths = date.y * 12 + (date.m - 1) - periods * (12 / frequency);
  const year = Math.floor(totalMonths / 12);
  const month = totalMonths - year * 12 + 1;
  const last = daysInMonthOf(year, month);
  return serialFromCivil(year, month, endOfMonth ? last : Math.min(date.d, last));
}

interface CouponPeriod {
  /** Previous coupon date (on or before settlement). */
  pcd: number;
  /** Next coupon date (after settlement). */
  ncd: number;
  /** Coupons payable between settlement and maturity. */
  count: number;
}

function couponPeriod(terms: BondTerms): CouponPeriod {
  const { settlement, maturity, frequency } = terms;
  const from = civilFromSerial(settlement);
  const to = civilFromSerial(maturity);
  const monthsApart = (to.y - from.y) * 12 + (to.m - from.m);
  let periods = Math.max(1, Math.floor(monthsApart / (12 / frequency)));
  while (couponDateBefore(maturity, periods, frequency) > settlement) periods += 1;
  while (periods > 1 && couponDateBefore(maturity, periods - 1, frequency) <= settlement) periods -= 1;
  return {
    pcd: couponDateBefore(maturity, periods, frequency),
    ncd: couponDateBefore(maturity, periods - 1, frequency),
    count: periods,
  };
}

/** COUPDAYS: days in the coupon period containing settlement. */
function couponDays(terms: BondTerms, period = couponPeriod(terms)): number {
  if (terms.basis === 1) return period.ncd - period.pcd;
  return (terms.basis === 3 ? 365 : 360) / terms.frequency;
}

/** COUPDAYBS: days from the start of the coupon period to settlement. */
function couponDaysBefore(terms: BondTerms, period = couponPeriod(terms)): number {
  return basisDayCount(period.pcd, terms.settlement, terms.basis);
}

/** COUPDAYSNC: days from settlement to the next coupon date. */
function couponDaysNext(terms: BondTerms, period = couponPeriod(terms)): number {
  if (terms.basis === 0 || terms.basis === 4) {
    return couponDays(terms, period) - couponDaysBefore(terms, period);
  }
  return period.ncd - terms.settlement;
}

function bondPrice(terms: BondTerms, rate: number, yld: number, redemption: number): number {
  const period = couponPeriod(terms);
  const e = couponDays(terms, period);
  const a = couponDaysBefore(terms, period);
  const dsc = couponDaysNext(terms, period);
  const coupon = (100 * rate) / terms.frequency;
  const n = period.count;
  if (n === 1) {
    return (coupon + redemption) / (1 + (yld / terms.frequency) * (dsc / e)) - coupon * (a / e);
  }
  const base = 1 + yld / terms.frequency;
  const offset = dsc / e;
  let price = redemption / base ** (n - 1 + offset);
  for (let k = 1; k <= n; k += 1) price += coupon / base ** (k - 1 + offset);
  return price - coupon * (a / e);
}

function bondYield(terms: BondTerms, rate: number, price: number, redemption: number): NumberOrError {
  const period = couponPeriod(terms);
  if (period.count === 1) {
    const e = couponDays(terms, period);
    const a = couponDaysBefore(terms, period);
    const dsr = couponDaysNext(terms, period);
    const paid = price / 100 + (a / e) * (rate / terms.frequency);
    return ((redemption / 100 + rate / terms.frequency - paid) / paid) * ((terms.frequency * e) / dsr);
  }
  const f = (yld: number) => bondPrice(terms, rate, yld, redemption) - price;
  return solveRate(f, rate > 0 ? rate : 0.05, -terms.frequency + 1e-10, 1e3);
}

/** Macaulay duration in years. */
function bondDuration(terms: BondTerms, coupon: number, yld: number): number {
  const period = couponPeriod(terms);
  const e = couponDays(terms, period);
  const offset = couponDaysNext(terms, period) / e;
  const cash = (100 * coupon) / terms.frequency;
  const base = 1 + yld / terms.frequency;
  let weighted = 0;
  let total = 0;
  for (let k = 1; k <= period.count; k += 1) {
    const time = k - 1 + offset;
    const flow = k === period.count ? cash + 100 : cash;
    const present = flow / base ** time;
    weighted += time * present;
    total += present;
  }
  return weighted / total / terms.frequency;
}

/**
 * Root of a monotone-ish function: Newton (numeric derivative) from `guess`, falling back to
 * a bracket scan plus bisection inside (low, high).
 */
function solveRate(f: (rate: number) => number, guess: number, low: number, high: number): NumberOrError {
  let rate = guess;
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const value = f(rate);
    if (!Number.isFinite(value)) break;
    if (Math.abs(value) < 1e-12) return rate;
    const step = Math.max(1e-7, Math.abs(rate) * 1e-7);
    const derivative = (f(rate + step) - f(rate - step)) / (2 * step);
    if (!Number.isFinite(derivative) || derivative === 0) break;
    const next = rate - value / derivative;
    if (!Number.isFinite(next) || next <= low || next >= high) break;
    if (Math.abs(next - rate) < 1e-13 * Math.max(1, Math.abs(rate))) return next;
    rate = next;
  }
  return bisectRoot(f, low, high);
}

function bisectRoot(f: (rate: number) => number, low: number, high: number): NumberOrError {
  // Scan a grid for a sign change, then bisect.
  const points: number[] = [];
  const span = [-0.999, -0.99, -0.9, -0.75, -0.5, -0.25, -0.1, 0, 0.05, 0.1, 0.2, 0.35, 0.5, 0.75, 1, 1.5, 2, 3, 5, 10, 20, 50, 100, 1e3];
  for (const point of span) if (point > low && point < high) points.push(point);
  let previous = Number.NaN;
  let previousValue = Number.NaN;
  for (const point of points) {
    const value = f(point);
    if (!Number.isFinite(value)) continue;
    if (value === 0) return point;
    if (Number.isFinite(previousValue) && Math.sign(value) !== Math.sign(previousValue)) {
      let a = previous;
      let b = point;
      let fa = previousValue;
      for (let iteration = 0; iteration < 200; iteration += 1) {
        const middle = (a + b) / 2;
        const fm = f(middle);
        if (!Number.isFinite(fm)) return numError();
        if (fm === 0 || Math.abs(b - a) < 1e-15 * Math.max(1, Math.abs(middle))) return middle;
        if (Math.sign(fm) === Math.sign(fa)) {
          a = middle;
          fa = fm;
        } else {
          b = middle;
        }
      }
      return (a + b) / 2;
    }
    previous = point;
    previousValue = value;
  }
  return numError();
}

// ---- Dated cash flows ------------------------------------------------------------------------

interface DatedFlows {
  values: number[];
  dates: number[];
}

function flowNumber(value: unknown): NumberOrError {
  if (isEvaluationError(value)) return value;
  if (typeof value === "number") return value;
  if (typeof value === "string") {
    const converted = toNumber(value);
    return isEvaluationError(converted) ? valueError() : converted;
  }
  return valueError();
}

function datedFlows(valuesArg: Value | undefined, datesArg: Value | undefined): DatedFlows | EvaluationError {
  const valueRange = rectArg(valuesArg);
  if (isEvaluationError(valueRange)) return valueRange;
  const dateRange = rectArg(datesArg);
  if (isEvaluationError(dateRange)) return dateRange;
  if (valueRange.values.length !== dateRange.values.length) return numError();
  const values: number[] = [];
  const dates: number[] = [];
  for (let index = 0; index < valueRange.values.length; index += 1) {
    const value = flowNumber(valueRange.values[index]);
    if (isEvaluationError(value)) return value;
    const rawDate = flowNumber(dateRange.values[index]);
    if (isEvaluationError(rawDate)) return rawDate;
    const date = Math.floor(rawDate);
    if (date < 0) return numError();
    values.push(value);
    dates.push(date);
  }
  if (values.length === 0) return numError();
  for (const date of dates) if (date < dates[0]) return numError();
  return { values, dates };
}

function xnpv(rate: number, flows: DatedFlows): number {
  let sum = 0;
  for (let index = 0; index < flows.values.length; index += 1) {
    sum += flows.values[index] / (1 + rate) ** ((flows.dates[index] - flows.dates[0]) / 365);
  }
  return sum;
}

function xirr(flows: DatedFlows, guess: number): NumberOrError {
  if (!flows.values.some((value) => value > 0) || !flows.values.some((value) => value < 0)) {
    return numError();
  }
  const derivative = (rate: number): number => {
    let sum = 0;
    for (let index = 0; index < flows.values.length; index += 1) {
      const time = (flows.dates[index] - flows.dates[0]) / 365;
      sum -= (time * flows.values[index]) / (1 + rate) ** (time + 1);
    }
    return sum;
  };
  let rate = guess;
  for (let iteration = 0; iteration < 100 && rate > -1; iteration += 1) {
    const value = xnpv(rate, flows);
    if (!Number.isFinite(value)) break;
    const slope = derivative(rate);
    if (!Number.isFinite(slope) || slope === 0) break;
    const next = rate - value / slope;
    if (!Number.isFinite(next) || next <= -1) break;
    if (Math.abs(next - rate) < 1e-12 * Math.max(1, Math.abs(rate))) {
      return Math.abs(xnpv(next, flows)) < 1e-6 * Math.max(1, ...flows.values.map(Math.abs)) ? next : numError();
    }
    rate = next;
  }
  return bisectRoot((candidate) => xnpv(candidate, flows), -1, 1e3);
}

// ---- Depreciation ----------------------------------------------------------------------------

function ddbPeriod(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life;
  let oldValue: number;
  if (rate >= 1) {
    rate = 1;
    oldValue = period === 1 ? cost : 0;
  } else {
    oldValue = cost * (1 - rate) ** (period - 1);
  }
  const newValue = cost * (1 - rate) ** period;
  const depreciation = newValue < salvage ? oldValue - salvage : oldValue - newValue;
  return depreciation < 0 ? 0 : depreciation;
}

/** Declining balance with a switch to straight line once that is larger (LibreOffice's ScInterVDB). */
function interVdb(cost: number, salvage: number, life: number, remainingLife: number, period: number, factor: number): number {
  let total = 0;
  const loopEnd = Math.ceil(period - 1e-12);
  let remaining = cost - salvage;
  let straight = 0;
  let switched = false;
  for (let index = 1; index <= loopEnd; index += 1) {
    let term: number;
    if (!switched) {
      const declining = ddbPeriod(cost, salvage, life, index, factor);
      straight = remaining / (remainingLife - (index - 1));
      if (straight > declining) {
        term = straight;
        switched = true;
      } else {
        term = declining;
        remaining -= declining;
      }
    } else {
      term = straight;
    }
    if (index === loopEnd) term *= period + 1 - loopEnd;
    total += term;
  }
  return total;
}

function vdb(cost: number, salvage: number, life: number, start: number, end: number, factor: number, noSwitch: boolean): number {
  const intStart = Math.floor(start + 1e-12);
  const intEnd = Math.ceil(end - 1e-12);
  if (noSwitch) {
    let total = 0;
    for (let index = intStart + 1; index <= intEnd; index += 1) {
      let term = ddbPeriod(cost, salvage, life, index, factor);
      if (index === intStart + 1) term *= Math.min(end, intStart + 1) - start;
      else if (index === intEnd) term *= end + 1 - intEnd;
      total += term;
    }
    return total;
  }
  let part = 0;
  if (Math.abs(start - intStart) > 1e-12) {
    const value = cost - interVdb(cost, salvage, life, life, intStart, factor);
    part += (start - intStart) * interVdb(value, salvage, life, life - intStart, 1, factor);
  }
  if (Math.abs(end - intEnd) > 1e-12) {
    const tempStart = intEnd - 1;
    const value = cost - interVdb(cost, salvage, life, life, tempStart, factor);
    part += (intEnd - end) * interVdb(value, salvage, life, life - tempStart, intEnd - tempStart, factor);
  }
  const remainingCost = cost - interVdb(cost, salvage, life, life, intStart, factor);
  return interVdb(remainingCost, salvage, life, life - intStart, intEnd - intStart, factor) - part;
}

function amorBasis(value: number): NumberOrError {
  const basis = Math.trunc(value);
  return basis === 0 || basis === 1 || basis === 3 || basis === 4 ? basis : numError();
}

// ---- Accrued interest ------------------------------------------------------------------------

/** Σ accrued days / normal quasi-coupon period length, with quasi-coupons anchored at `anchor`. */
function accrualPeriods(start: number, end: number, anchor: number, frequency: number, basis: number): number {
  const anchorDate = civilFromSerial(anchor);
  const endOfMonth = anchorDate.d === daysInMonthOf(anchorDate.y, anchorDate.m);
  const step = 12 / frequency;
  const couponAt = (k: number): number => {
    const totalMonths = anchorDate.y * 12 + (anchorDate.m - 1) + k * step;
    const year = Math.floor(totalMonths / 12);
    const month = totalMonths - year * 12 + 1;
    const last = daysInMonthOf(year, month);
    return serialFromCivil(year, month, endOfMonth ? last : Math.min(anchorDate.d, last));
  };
  const startDate = civilFromSerial(start);
  let k = Math.floor(((startDate.y - anchorDate.y) * 12 + (startDate.m - anchorDate.m)) / step) - 1;
  while (couponAt(k) > start) k -= 1;
  while (couponAt(k + 1) <= start) k += 1;
  let total = 0;
  for (let guard = 0; guard < 100_000; guard += 1, k += 1) {
    const periodStart = couponAt(k);
    const periodEnd = couponAt(k + 1);
    if (periodStart >= end) break;
    const from = Math.max(start, periodStart);
    const to = Math.min(end, periodEnd);
    const length = basis === 1 ? periodEnd - periodStart : (basis === 3 ? 365 : 360) / frequency;
    const days = basis === 0 ? days360Nasd(from, to) : basisDayCount(from, to, basis);
    total += days / length;
  }
  return total;
}

// ---- Odd first/last coupon periods -----------------------------------------------------------

/** Quasi-coupon date `k` periods after `anchor` (negative steps back), month-end preserving. */
function quasiCouponDate(anchor: number, k: number, frequency: number): number {
  const date = civilFromSerial(anchor);
  const endOfMonth = date.d === daysInMonthOf(date.y, date.m);
  const totalMonths = date.y * 12 + (date.m - 1) + k * (12 / frequency);
  const year = Math.floor(totalMonths / 12);
  const month = totalMonths - year * 12 + 1;
  const last = daysInMonthOf(year, month);
  return serialFromCivil(year, month, endOfMonth ? last : Math.min(date.d, last));
}

function oddDayCount(from: number, to: number, basis: number): number {
  return basis === 0 ? days360Nasd(from, to) : basisDayCount(from, to, basis);
}

function oddPeriodLength(start: number, end: number, basis: number, frequency: number): number {
  return basis === 1 ? end - start : (basis === 3 ? 365 : 360) / frequency;
}

interface OddFirstTerms {
  settlement: number;
  maturity: number;
  issue: number;
  firstCoupon: number;
  frequency: number;
  basis: number;
}

function oddFirstTerms(values: number[]): OddFirstTerms | EvaluationError {
  const [settlementValue, maturityValue, issueValue, firstValue] = values;
  const dates: number[] = [];
  for (const value of [settlementValue, maturityValue, issueValue, firstValue]) {
    const date = wholeDate(value);
    if (isEvaluationError(date)) return date;
    dates.push(date);
  }
  const [settlement, maturity, issue, firstCoupon] = dates;
  const frequency = validFrequency(values[4]);
  if (isEvaluationError(frequency)) return frequency;
  const basis = validBasis(values[5]);
  if (isEvaluationError(basis)) return basis;
  if (!(maturity > firstCoupon && firstCoupon > settlement && settlement > issue)) return numError();
  return { settlement, maturity, issue, firstCoupon, frequency, basis };
}

/** ODDFPRICE: cash flows of a bond whose first coupon period is short or long. */
function oddFirstPrice(terms: OddFirstTerms, rate: number, yld: number, redemption: number): number {
  const { settlement, maturity, issue, firstCoupon, frequency, basis } = terms;
  const coupon = (100 * rate) / frequency;
  const base = 1 + yld / frequency;
  const first = civilFromSerial(firstCoupon);
  const last = civilFromSerial(maturity);
  const regular = Math.max(0, Math.round(((last.y - first.y) * 12 + (last.m - first.m)) / (12 / frequency)));
  const oddFraction = accrualPeriods(issue, firstCoupon, firstCoupon, frequency, basis);
  const accrued = accrualPeriods(issue, settlement, firstCoupon, frequency, basis);
  let index = -1;
  while (quasiCouponDate(firstCoupon, index, frequency) > settlement) index -= 1;
  const periodStart = quasiCouponDate(firstCoupon, index, frequency);
  const periodEnd = quasiCouponDate(firstCoupon, index + 1, frequency);
  const length = oddPeriodLength(periodStart, periodEnd, basis, frequency);
  const toNext =
    basis === 0 || basis === 4 ? length - oddDayCount(periodStart, settlement, basis) : oddDayCount(settlement, periodEnd, basis);
  const firstTime = -(index + 1) + toNext / length;
  let price = (coupon * oddFraction) / base ** firstTime;
  for (let k = 1; k <= regular; k += 1) price += coupon / base ** (firstTime + k);
  price += redemption / base ** (firstTime + regular);
  return price - coupon * accrued;
}

interface OddLastTerms {
  settlement: number;
  maturity: number;
  lastInterest: number;
  frequency: number;
  basis: number;
}

function oddLastTerms(values: number[]): OddLastTerms | EvaluationError {
  const dates: number[] = [];
  for (const value of values.slice(0, 3)) {
    const date = wholeDate(value);
    if (isEvaluationError(date)) return date;
    dates.push(date);
  }
  const [settlement, maturity, lastInterest] = dates;
  const frequency = validFrequency(values[3]);
  if (isEvaluationError(frequency)) return frequency;
  const basis = validBasis(values[4]);
  if (isEvaluationError(basis)) return basis;
  if (!(maturity > settlement && settlement > lastInterest)) return numError();
  return { settlement, maturity, lastInterest, frequency, basis };
}

/** Σ DC/NL (whole odd period), Σ A/NL (accrued) and Σ DSC/NL (remaining) for an odd last period. */
function oddLastFractions(terms: OddLastTerms): { total: number; accrued: number; remaining: number } {
  const { settlement, maturity, lastInterest, frequency, basis } = terms;
  return {
    total: accrualPeriods(lastInterest, maturity, lastInterest, frequency, basis),
    accrued: accrualPeriods(lastInterest, settlement, lastInterest, frequency, basis),
    remaining: accrualPeriods(settlement, maturity, lastInterest, frequency, basis),
  };
}

function tbillDays(settlementValue: number, maturityValue: number): NumberOrError {
  const settlement = wholeDate(settlementValue);
  if (isEvaluationError(settlement)) return settlement;
  const maturity = wholeDate(maturityValue);
  if (isEvaluationError(maturity)) return maturity;
  if (settlement >= maturity) return numError();
  const date = civilFromSerial(settlement);
  const oneYearLater = serialFromCivil(date.y + 1, date.m, Math.min(date.d, daysInMonthOf(date.y + 1, date.m)));
  if (maturity > oneYearLater) return numError();
  return maturity - settlement;
}

function cumulativeParts(
  kind: "interest" | "principal",
  rate: number,
  rawNper: number,
  pv: number,
  rawStart: number,
  rawEnd: number,
  rawType: number,
): NumberOrError {
  const nper = Math.trunc(rawNper);
  const start = Math.trunc(rawStart);
  const end = Math.trunc(rawEnd);
  if (rate <= 0 || nper <= 0 || pv <= 0 || start < 1 || end < start || end > nper) return numError();
  if (rawType !== 0 && rawType !== 1) return numError();
  const type = rawType;
  const payment = annuityPayment(rate, nper, pv, 0, type);
  if (isEvaluationError(payment)) return payment;
  let interest = 0;
  let principal = 0;
  let first = start;
  if (first === 1) {
    if (type === 0) {
      interest = -pv;
      principal = payment + pv * rate;
    } else {
      principal = payment;
    }
    first = 2;
  }
  for (let period = first; period <= end; period += 1) {
    const periodInterest =
      type === 1
        ? annuityFutureValue(rate, period - 2, payment, pv, 1) - payment
        : annuityFutureValue(rate, period - 1, payment, pv, 0);
    interest += periodInterest;
    principal += payment - periodInterest * rate;
  }
  return kind === "interest" ? interest * rate : principal;
}

/** Excel treats an omitted (empty) par as $1,000 and still requires the argument slot. */
function parDefault(base: Specs[string], index: number, minArgs: number): Specs[string] {
  return {
    ...base,
    minArgs,
    impl: (values, call) =>
      base.impl(values.map((value, position) => (position === index && value === null ? 1000 : value)), call),
  };
}

function dollarFraction(value: number, rawFraction: number, direction: "decimal" | "fraction"): NumberOrError {
  const fraction = Math.trunc(rawFraction);
  if (fraction < 0) return numError();
  if (fraction === 0) return divError();
  const integer = Math.trunc(value);
  const remainder = value - integer;
  const scale = 10 ** Math.ceil(Math.log10(fraction) - 1e-12);
  const result = direction === "decimal" ? integer + (remainder * scale) / fraction : integer + (remainder * fraction) / scale;
  return Number(result.toPrecision(15));
}

export const FINANCIAL_FUNCTIONS: Specs = {
  EFFECT: financial([undefined, undefined], (nominal, rawPeriods) => {
    const periods = Math.trunc(rawPeriods);
    if (nominal <= 0 || periods < 1) return numError();
    return (1 + nominal / periods) ** periods - 1;
  }),
  NOMINAL: financial([undefined, undefined], (effective, rawPeriods) => {
    const periods = Math.trunc(rawPeriods);
    if (effective <= 0 || periods < 1) return numError();
    return periods * ((1 + effective) ** (1 / periods) - 1);
  }),
  XNPV: spec(3, 3, (values) => {
    const rate = numberArg(values[0]);
    if (isEvaluationError(rate)) return rate;
    if (rate <= -1) return numError();
    const flows = datedFlows(values[1], values[2]);
    if (isEvaluationError(flows)) return flows;
    return finiteResult(xnpv(rate, flows));
  }),
  XIRR: spec(2, 3, (values) => {
    const flows = datedFlows(values[0], values[1]);
    if (isEvaluationError(flows)) return flows;
    const guess = values[2] === undefined || values[2] === null ? 0.1 : numberArg(values[2]);
    if (isEvaluationError(guess)) return guess;
    if (guess <= -1) return numError();
    const result = xirr(flows, guess);
    return isEvaluationError(result) ? result : finiteResult(result);
  }),
  MIRR: spec(3, 3, (values) => {
    const flows = readNumbers([values[0]]);
    if (isEvaluationError(flows)) return flows;
    const financeRate = numberArg(values[1]);
    if (isEvaluationError(financeRate)) return financeRate;
    const reinvestRate = numberArg(values[2]);
    if (isEvaluationError(reinvestRate)) return reinvestRate;
    const n = flows.length;
    let positive = 0;
    let negative = 0;
    for (let index = 0; index < n; index += 1) {
      const flow = flows[index];
      if (flow > 0) positive += flow * (1 + reinvestRate) ** (n - 1 - index);
      else if (flow < 0) negative += flow / (1 + financeRate) ** index;
    }
    if (positive === 0 || negative === 0 || n < 2) return divError();
    return finiteResult((positive / -negative) ** (1 / (n - 1)) - 1);
  }),
  SLN: financial([undefined, undefined, undefined], (cost, salvage, life) =>
    life === 0 ? divError() : (cost - salvage) / life,
  ),
  SYD: financial([undefined, undefined, undefined, undefined], (cost, salvage, life, period) => {
    if (life <= 0 || period <= 0 || period > life) return numError();
    return ((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1));
  }),
  DDB: financial([undefined, undefined, undefined, undefined, 2], (cost, salvage, life, period, factor) => {
    if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0 || period > life) return numError();
    return ddbPeriod(cost, salvage, life, period, factor);
  }),
  DB: financial([undefined, undefined, undefined, undefined, 12], (cost, salvage, life, period, months) => {
    if (months < 1 || months > 12 || salvage < 0 || salvage > cost || cost <= 0 || life <= 0 || period <= 0 || period > life + 1) {
      return numError();
    }
    let rate = 1 - (salvage / cost) ** (1 / life);
    rate = Math.floor(rate * 1000 + 0.5) / 1000;
    const first = (cost * rate * months) / 12;
    if (Math.floor(period) === 1) return first;
    let total = first;
    let depreciation = 0;
    const last = Math.floor(Math.min(life, period));
    for (let index = 2; index <= last; index += 1) {
      depreciation = (cost - total) * rate;
      total += depreciation;
    }
    if (period > life) depreciation = ((cost - total) * rate * (12 - months)) / 12;
    return depreciation;
  }),
  VDB: spec(
    5,
    7,
    (values) => {
      const args = readArgs(values.slice(0, 6), [undefined, undefined, undefined, undefined, undefined, 2]);
      if (isEvaluationError(args)) return args;
      const [cost, salvage, life, start, end, factor] = args;
      const noSwitch = values[6] === undefined ? 0 : numberArg(values[6]);
      if (isEvaluationError(noSwitch)) return noSwitch;
      if (start < 0 || end < start || end > life || cost < 0 || salvage > cost || factor <= 0 || life <= 0) {
        return numError();
      }
      return finiteResult(vdb(cost, salvage, life, start, end, factor, noSwitch !== 0));
    },
    { liftArgs: "all" },
  ),
  CUMIPMT: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined],
    (rate, nper, pv, start, end, type) => cumulativeParts("interest", rate, nper, pv, start, end, type),
  ),
  CUMPRINC: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined],
    (rate, nper, pv, start, end, type) => cumulativeParts("principal", rate, nper, pv, start, end, type),
  ),
  PDURATION: financial([undefined, undefined, undefined], (rate, pv, fv) => {
    if (rate <= 0 || pv <= 0 || fv <= 0) return numError();
    return (Math.log(fv) - Math.log(pv)) / Math.log(1 + rate);
  }),
  RRI: financial([undefined, undefined, undefined], (nper, pv, fv) => {
    if (nper <= 0 || pv === 0) return numError();
    return (fv / pv) ** (1 / nper) - 1;
  }),
  FVSCHEDULE: spec(2, 2, (values) => {
    const principal = numberArg(values[0]);
    if (isEvaluationError(principal)) return principal;
    let result = principal;
    for (const entry of collectValues([values[1]])) {
      const rate = entry.value;
      if (isEvaluationError(rate)) return rate;
      if (rate === null) continue;
      if (typeof rate === "number") result *= 1 + rate;
      else if (!entry.fromRange && typeof rate === "string") {
        const converted = toNumber(rate);
        if (isEvaluationError(converted)) return valueError();
        result *= 1 + converted;
      } else {
        return valueError();
      }
    }
    return finiteResult(result);
  }),
  DOLLARDE: financial([undefined, undefined], (value, fraction) => dollarFraction(value, fraction, "decimal")),
  DOLLARFR: financial([undefined, undefined], (value, fraction) => dollarFraction(value, fraction, "fraction")),
  ISPMT: financial([undefined, undefined, undefined, undefined], (rate, period, nper, pv) =>
    nper === 0 ? divError() : pv * rate * (period / nper - 1),
  ),

  // ---- Coupon bonds ----------------------------------------------------------------------------
  COUPDAYBS: financial([undefined, undefined, undefined, 0], (settlement, maturity, frequency, basis) => {
    const terms = bondTerms(settlement, maturity, frequency, basis);
    return isEvaluationError(terms) ? terms : couponDaysBefore(terms);
  }),
  COUPDAYS: financial([undefined, undefined, undefined, 0], (settlement, maturity, frequency, basis) => {
    const terms = bondTerms(settlement, maturity, frequency, basis);
    return isEvaluationError(terms) ? terms : couponDays(terms);
  }),
  COUPDAYSNC: financial([undefined, undefined, undefined, 0], (settlement, maturity, frequency, basis) => {
    const terms = bondTerms(settlement, maturity, frequency, basis);
    return isEvaluationError(terms) ? terms : couponDaysNext(terms);
  }),
  COUPNCD: financial([undefined, undefined, undefined, 0], (settlement, maturity, frequency, basis) => {
    const terms = bondTerms(settlement, maturity, frequency, basis);
    return isEvaluationError(terms) ? terms : couponPeriod(terms).ncd;
  }),
  COUPPCD: financial([undefined, undefined, undefined, 0], (settlement, maturity, frequency, basis) => {
    const terms = bondTerms(settlement, maturity, frequency, basis);
    return isEvaluationError(terms) ? terms : couponPeriod(terms).pcd;
  }),
  COUPNUM: financial([undefined, undefined, undefined, 0], (settlement, maturity, frequency, basis) => {
    const terms = bondTerms(settlement, maturity, frequency, basis);
    return isEvaluationError(terms) ? terms : couponPeriod(terms).count;
  }),
  PRICE: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, rate, yld, redemption, frequency, basis) => {
      const terms = bondTerms(settlement, maturity, frequency, basis);
      if (isEvaluationError(terms)) return terms;
      if (rate < 0 || yld < 0 || redemption <= 0) return numError();
      return bondPrice(terms, rate, yld, redemption);
    },
  ),
  YIELD: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, rate, price, redemption, frequency, basis) => {
      const terms = bondTerms(settlement, maturity, frequency, basis);
      if (isEvaluationError(terms)) return terms;
      if (rate < 0 || price <= 0 || redemption <= 0) return numError();
      return bondYield(terms, rate, price, redemption);
    },
  ),
  DURATION: financial(
    [undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, coupon, yld, frequency, basis) => {
      const terms = bondTerms(settlement, maturity, frequency, basis);
      if (isEvaluationError(terms)) return terms;
      if (coupon < 0 || yld < 0) return numError();
      return bondDuration(terms, coupon, yld);
    },
  ),
  MDURATION: financial(
    [undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, coupon, yld, frequency, basis) => {
      const terms = bondTerms(settlement, maturity, frequency, basis);
      if (isEvaluationError(terms)) return terms;
      if (coupon < 0 || yld < 0) return numError();
      return bondDuration(terms, coupon, yld) / (1 + yld / terms.frequency);
    },
  ),
  ACCRINT: parDefault(financial(
    [undefined, undefined, undefined, undefined, 1000, undefined, 0, 1],
    (issueValue, firstInterestValue, settlementValue, rate, par, frequencyValue, basisValue, calcMethod) => {
      const issue = wholeDate(issueValue);
      if (isEvaluationError(issue)) return issue;
      const firstInterest = wholeDate(firstInterestValue);
      if (isEvaluationError(firstInterest)) return firstInterest;
      const settlement = wholeDate(settlementValue);
      if (isEvaluationError(settlement)) return settlement;
      const frequency = validFrequency(frequencyValue);
      if (isEvaluationError(frequency)) return frequency;
      const basis = validBasis(basisValue);
      if (isEvaluationError(basis)) return basis;
      if (rate <= 0 || par <= 0 || issue >= settlement) return numError();
      const start = calcMethod === 0 && settlement > firstInterest ? firstInterest : issue;
      return ((par * rate) / frequency) * accrualPeriods(start, settlement, firstInterest, frequency, basis);
    },
  ), 4, 6),
  ACCRINTM: parDefault(financial([undefined, undefined, undefined, 1000, 0], (issueValue, settlementValue, rate, par, basisValue) => {
    const terms = datedTerms(issueValue, settlementValue, basisValue);
    if (isEvaluationError(terms)) return terms;
    if (rate <= 0 || par <= 0) return numError();
    return par * rate * yearFrac(terms.settlement, terms.maturity, terms.basis);
  }), 3, 4),
  DISC: financial([undefined, undefined, undefined, undefined, 0], (settlement, maturity, price, redemption, basis) => {
    const terms = datedTerms(settlement, maturity, basis);
    if (isEvaluationError(terms)) return terms;
    if (price <= 0 || redemption <= 0) return numError();
    return (1 - price / redemption) / yearFrac(terms.settlement, terms.maturity, terms.basis);
  }),
  INTRATE: financial([undefined, undefined, undefined, undefined, 0], (settlement, maturity, investment, redemption, basis) => {
    const terms = datedTerms(settlement, maturity, basis);
    if (isEvaluationError(terms)) return terms;
    if (investment <= 0 || redemption <= 0) return numError();
    return (redemption / investment - 1) / yearFrac(terms.settlement, terms.maturity, terms.basis);
  }),
  RECEIVED: financial([undefined, undefined, undefined, undefined, 0], (settlement, maturity, investment, discount, basis) => {
    const terms = datedTerms(settlement, maturity, basis);
    if (isEvaluationError(terms)) return terms;
    if (investment <= 0 || discount <= 0) return numError();
    const denominator = 1 - discount * yearFrac(terms.settlement, terms.maturity, terms.basis);
    return denominator <= 0 ? numError() : investment / denominator;
  }),
  PRICEDISC: financial([undefined, undefined, undefined, undefined, 0], (settlement, maturity, discount, redemption, basis) => {
    const terms = datedTerms(settlement, maturity, basis);
    if (isEvaluationError(terms)) return terms;
    if (discount <= 0 || redemption <= 0) return numError();
    return redemption * (1 - discount * yearFrac(terms.settlement, terms.maturity, terms.basis));
  }),
  YIELDDISC: financial([undefined, undefined, undefined, undefined, 0], (settlement, maturity, price, redemption, basis) => {
    const terms = datedTerms(settlement, maturity, basis);
    if (isEvaluationError(terms)) return terms;
    if (price <= 0 || redemption <= 0) return numError();
    return (redemption / price - 1) / yearFrac(terms.settlement, terms.maturity, terms.basis);
  }),
  PRICEMAT: financial(
    [undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, issueValue, rate, yld, basis) => {
      const terms = datedTerms(settlement, maturity, basis);
      if (isEvaluationError(terms)) return terms;
      const issue = wholeDate(issueValue);
      if (isEvaluationError(issue)) return issue;
      if (rate < 0 || yld < 0 || issue > terms.settlement) return numError();
      const toMaturity = yearFrac(issue, terms.maturity, terms.basis);
      const accrued = yearFrac(issue, terms.settlement, terms.basis);
      const remaining = yearFrac(terms.settlement, terms.maturity, terms.basis);
      return (100 + toMaturity * rate * 100) / (1 + remaining * yld) - accrued * rate * 100;
    },
  ),
  YIELDMAT: financial(
    [undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, issueValue, rate, price, basis) => {
      const terms = datedTerms(settlement, maturity, basis);
      if (isEvaluationError(terms)) return terms;
      const issue = wholeDate(issueValue);
      if (isEvaluationError(issue)) return issue;
      if (rate < 0 || price <= 0 || issue > terms.settlement) return numError();
      const toMaturity = yearFrac(issue, terms.maturity, terms.basis);
      const accrued = yearFrac(issue, terms.settlement, terms.basis);
      const remaining = yearFrac(terms.settlement, terms.maturity, terms.basis);
      const paid = price / 100 + accrued * rate;
      return ((1 + toMaturity * rate - paid) / paid) / remaining;
    },
  ),
  TBILLPRICE: financial([undefined, undefined, undefined], (settlement, maturity, discount) => {
    const days = tbillDays(settlement, maturity);
    if (isEvaluationError(days)) return days;
    if (discount <= 0) return numError();
    const price = 100 * (1 - (discount * days) / 360);
    return price <= 0 ? numError() : price;
  }),
  TBILLYIELD: financial([undefined, undefined, undefined], (settlement, maturity, price) => {
    const days = tbillDays(settlement, maturity);
    if (isEvaluationError(days)) return days;
    if (price <= 0) return numError();
    return ((100 - price) / price) * (360 / days);
  }),
  TBILLEQ: financial([undefined, undefined, undefined], (settlement, maturity, discount) => {
    const days = tbillDays(settlement, maturity);
    if (isEvaluationError(days)) return days;
    if (discount <= 0) return numError();
    if (days <= 182) return (365 * discount) / (360 - discount * days);
    const price = 100 * (1 - (discount * days) / 360);
    if (price <= 0) return numError();
    const term = days / 365;
    return (-term + Math.sqrt(term * term - (2 * term - 1) * (1 - 100 / price))) / (term - 0.5);
  }),
  AMORLINC: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined, 0],
    (cost, purchasedValue, firstPeriodValue, salvage, rawPeriod, rate, basisValue) => {
      const purchased = wholeDate(purchasedValue);
      if (isEvaluationError(purchased)) return purchased;
      const firstPeriod = wholeDate(firstPeriodValue);
      if (isEvaluationError(firstPeriod)) return firstPeriod;
      const basis = amorBasis(basisValue);
      if (isEvaluationError(basis)) return basis;
      if (rate <= 0 || rawPeriod < 0 || cost < 0 || salvage < 0 || salvage > cost || purchased > firstPeriod) return numError();
      const period = Math.trunc(rawPeriod);
      const fullRate = cost * rate;
      const firstRate = yearFrac(purchased, firstPeriod, basis) * rate * cost;
      const fullPeriods = Math.trunc((cost - salvage - firstRate) / fullRate);
      let result = 0;
      if (period === 0) result = firstRate;
      else if (period <= fullPeriods) result = fullRate;
      else if (period === fullPeriods + 1) result = cost - salvage - fullRate * fullPeriods - firstRate;
      return result > 0 ? result : 0;
    },
  ),
  AMORDEGRC: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined, 0],
    (costValue, purchasedValue, firstPeriodValue, salvage, rawPeriod, rateValue, basisValue) => {
      const purchased = wholeDate(purchasedValue);
      if (isEvaluationError(purchased)) return purchased;
      const firstPeriod = wholeDate(firstPeriodValue);
      if (isEvaluationError(firstPeriod)) return firstPeriod;
      const basis = amorBasis(basisValue);
      if (isEvaluationError(basis)) return basis;
      if (rateValue <= 0 || rawPeriod < 0 || costValue < 0 || salvage < 0 || salvage > costValue || purchased > firstPeriod) {
        return numError();
      }
      const usefulLife = 1 / rateValue;
      // Excel rejects a useful life strictly between 0-1, 1-2, 2-3 and 4-5 years.
      if ((usefulLife > 0 && usefulLife < 1) || (usefulLife > 1 && usefulLife < 2) || (usefulLife > 2 && usefulLife < 3) || (usefulLife > 4 && usefulLife < 5)) {
        return numError();
      }
      const coefficient = usefulLife < 3 ? 1 : usefulLife < 5 ? 1.5 : usefulLife <= 6 ? 2 : 2.5;
      const rate = rateValue * coefficient;
      const period = Math.trunc(rawPeriod);
      let cost = costValue;
      let depreciation = Math.round(yearFrac(purchased, firstPeriod, basis) * rate * cost);
      cost -= depreciation;
      let rest = cost - salvage;
      for (let index = 0; index < period; index += 1) {
        depreciation = Math.round(rate * cost);
        rest -= depreciation;
        if (rest < 0) {
          return period - index <= 1 ? Math.round(cost * 0.5) : 0;
        }
        cost -= depreciation;
      }
      return depreciation;
    },
  ),
  ODDFPRICE: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis) => {
      const terms = oddFirstTerms([settlement, maturity, issue, firstCoupon, frequency, basis]);
      if (isEvaluationError(terms)) return terms;
      if (rate < 0 || yld < 0 || redemption <= 0) return numError();
      return oddFirstPrice(terms, rate, yld, redemption);
    },
  ),
  ODDFYIELD: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, issue, firstCoupon, rate, price, redemption, frequency, basis) => {
      const terms = oddFirstTerms([settlement, maturity, issue, firstCoupon, frequency, basis]);
      if (isEvaluationError(terms)) return terms;
      if (rate < 0 || price <= 0 || redemption <= 0) return numError();
      return solveRate(
        (yld) => oddFirstPrice(terms, rate, yld, redemption) - price,
        rate > 0 ? rate : 0.05,
        -terms.frequency + 1e-10,
        1e3,
      );
    },
  ),
  ODDLPRICE: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis) => {
      const terms = oddLastTerms([settlement, maturity, lastInterest, frequency, basis]);
      if (isEvaluationError(terms)) return terms;
      if (rate < 0 || yld < 0 || redemption <= 0) return numError();
      const coupon = (100 * rate) / terms.frequency;
      const { total, accrued, remaining } = oddLastFractions(terms);
      return (redemption + total * coupon) / (1 + (remaining * yld) / terms.frequency) - accrued * coupon;
    },
  ),
  ODDLYIELD: financial(
    [undefined, undefined, undefined, undefined, undefined, undefined, undefined, 0],
    (settlement, maturity, lastInterest, rate, price, redemption, frequency, basis) => {
      const terms = oddLastTerms([settlement, maturity, lastInterest, frequency, basis]);
      if (isEvaluationError(terms)) return terms;
      if (rate < 0 || price <= 0 || redemption <= 0) return numError();
      const coupon = (100 * rate) / terms.frequency;
      const { total, accrued, remaining } = oddLastFractions(terms);
      const paid = price + accrued * coupon;
      return ((redemption + total * coupon - paid) / paid) * (terms.frequency / remaining);
    },
  ),
};
