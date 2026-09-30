// Engineering worksheet functions for the extended library (see formula-library.ts): radix
// conversions (BIN/OCT/DEC/HEX with 10-digit two's complement), bitwise operations, DELTA and
// GESTEP, CONVERT, complex-number text functions (COMPLEX, IM*), and the Bessel functions.
// ERF/ERFC live with the statistics module.

import {
  collectValues,
  finiteResult,
  isEvaluationError,
  numberToText,
  scalarArgument,
  toNumber,
  toText,
} from "./formulas";
import type { EvaluationError } from "./formulas";
import {
  divError,
  naError,
  numberArg,
  numError,
  numericFunction,
  spec,
  valueError,
} from "./formula-lib-shared";
import type { Specs, Value } from "./formula-lib-shared";

// ---- Radix conversions ---------------------------------------------------------------------

interface Radix {
  base: number;
  /** Bits in a 10-digit two's complement number of this radix. */
  bits: number;
  digits: RegExp;
}

const BIN: Radix = { base: 2, bits: 10, digits: /^[01]*$/ };
const OCT: Radix = { base: 8, bits: 30, digits: /^[0-7]*$/ };
const DEC: Radix = { base: 10, bits: 0, digits: /^\d*$/ };
const HEX: Radix = { base: 16, bits: 40, digits: /^[0-9A-Fa-f]*$/ };

/** Parse a BIN/OCT/HEX argument (text or number, at most 10 digits) as a signed value. */
function parseRadix(value: Value | undefined, radix: Radix): number | EvaluationError {
  const scalar = scalarArgument(value ?? null);
  if (isEvaluationError(scalar)) return scalar;
  if (typeof scalar === "boolean") return valueError();
  let text: string;
  if (typeof scalar === "number") {
    if (scalar < 0 || !Number.isInteger(scalar)) return numError();
    text = numberToText(scalar);
  } else {
    text = scalar === null ? "" : scalar.trim();
  }
  if (text.length > 10 || !radix.digits.test(text)) return numError();
  if (text === "") return 0;
  const unsigned = parseInt(text, radix.base);
  if (text.length === 10 && unsigned >= 2 ** (radix.bits - 1)) return unsigned - 2 ** radix.bits;
  return unsigned;
}

/** Format a signed value in a radix, honoring Excel's [places] rules and range limits. */
function formatRadix(value: number, radix: Radix, placesValue: Value | undefined): Value {
  const limit = 2 ** (radix.bits - 1);
  if (value < -limit || value >= limit) return numError();
  if (value < 0) return (2 ** radix.bits + value).toString(radix.base).toUpperCase();
  const text = value.toString(radix.base).toUpperCase();
  if (placesValue === undefined) return text;
  const rawPlaces = toNumber(scalarArgument(placesValue));
  if (isEvaluationError(rawPlaces)) return rawPlaces;
  const places = Math.trunc(rawPlaces);
  if (places <= 0 || places > 10 || text.length > places) return numError();
  return text.padStart(places, "0");
}

function radixConversion(from: Radix, to: Radix) {
  const maxArgs = to === DEC ? 1 : 2;
  return spec(1, maxArgs, (values) => {
    let number: number | EvaluationError;
    if (from === DEC) {
      const raw = scalarArgument(values[0]);
      if (isEvaluationError(raw)) return raw;
      if (typeof raw === "boolean") return valueError();
      const numeric = toNumber(raw);
      if (isEvaluationError(numeric)) return numeric;
      number = Math.trunc(numeric);
    } else {
      number = parseRadix(values[0], from);
    }
    if (isEvaluationError(number)) return number;
    if (to === DEC) return number;
    return formatRadix(number, to, values[1]);
  }, { liftArgs: "all" });
}

// ---- Bitwise operations (0 <= n < 2^48) ------------------------------------------------------

const BIT_LIMIT = 2 ** 48;

function bitOperand(value: number): number | EvaluationError {
  if (value < 0 || value >= BIT_LIMIT || !Number.isInteger(value)) return numError();
  return value;
}

function bitwise(combine: (left: bigint, right: bigint) => bigint) {
  return numericFunction([undefined, undefined], (rawLeft, rawRight) => {
    const left = bitOperand(rawLeft);
    if (isEvaluationError(left)) return left;
    const right = bitOperand(rawRight);
    if (isEvaluationError(right)) return right;
    return Number(combine(BigInt(left), BigInt(right)));
  });
}

function bitShift(direction: 1 | -1) {
  return numericFunction([undefined, undefined], (rawNumber, rawShift) => {
    const number = bitOperand(rawNumber);
    if (isEvaluationError(number)) return number;
    const shift = Math.trunc(rawShift) * direction;
    if (Math.abs(shift) > 53) return numError();
    const result = shift >= 0 ? number * 2 ** shift : Math.floor(number / 2 ** -shift);
    return result >= BIT_LIMIT ? numError() : result;
  });
}

// ---- CONVERT ---------------------------------------------------------------------------------

type UnitCategory =
  | "mass" | "length" | "time" | "pressure" | "force" | "energy" | "power" | "magnetism"
  | "temperature" | "volume" | "area" | "speed" | "information";

interface UnitDefinition {
  category: UnitCategory;
  /** Size of one unit in the category's base unit. */
  factor: number;
  /** Accepts metric prefixes ("km", "mg", ...). */
  prefix?: boolean;
  /** Squared/cubed units raise the prefix to this power ("km2" = 1e6 m2). */
  power?: number;
}

const INCH = 0.0254;
const FOOT = 0.3048;
const YARD = 0.9144;
const MILE = 1609.344;
const NAUTICAL_MILE = 1852;
const LIGHT_YEAR = 9460730472580800;
const PICA_POINT = INCH / 72;
const PICA = INCH / 6;
const US_GALLON = 3.785411784e-3;
const UK_GALLON = 4.54609e-3;
const POUND = 453.59237;
const LBF = 4.4482216152605;
const HORSEPOWER = 745.69987158227;

const UNITS: Record<string, UnitDefinition> = {};
function define(category: UnitCategory, factor: number, names: string[], options: { prefix?: boolean; power?: number } = {}) {
  for (const name of names) UNITS[name] = { category, factor, ...options };
}

// Mass (gram).
define("mass", 1, ["g"], { prefix: true });
define("mass", 14593.9029372064, ["sg"]);
define("mass", POUND, ["lbm"]);
define("mass", 1.660538782e-24, ["u"], { prefix: true });
define("mass", POUND / 16, ["ozm"]);
define("mass", 0.06479891, ["grain"]);
define("mass", POUND * 100, ["cwt", "shweight"]);
define("mass", POUND * 112, ["uk_cwt", "lcwt", "hweight"]);
define("mass", POUND * 14, ["stone"]);
define("mass", POUND * 2000, ["ton"]);
define("mass", POUND * 2240, ["uk_ton", "LTON", "brton"]);
// Length (meter).
define("length", 1, ["m"], { prefix: true });
define("length", MILE, ["mi"]);
define("length", NAUTICAL_MILE, ["Nmi"]);
define("length", INCH, ["in"]);
define("length", FOOT, ["ft"]);
define("length", YARD, ["yd"]);
define("length", 1e-10, ["ang"], { prefix: true });
define("length", 1.143, ["ell"]);
define("length", LIGHT_YEAR, ["ly"], { prefix: true });
define("length", 30856775812815500, ["parsec", "pc"], { prefix: true });
define("length", PICA_POINT, ["Pica", "Picapt"]);
define("length", PICA, ["pica"]);
define("length", (5280 * 1200) / 3937, ["survey_mi"]);
// Time (second).
define("time", 365.25 * 86400, ["yr"]);
define("time", 86400, ["day", "d"]);
define("time", 3600, ["hr"]);
define("time", 60, ["mn", "min"]);
define("time", 1, ["sec", "s"], { prefix: true });
// Pressure (pascal).
define("pressure", 1, ["Pa", "p"], { prefix: true });
define("pressure", 101325, ["atm", "at"], { prefix: true });
define("pressure", 101325 / 760, ["mmHg"], { prefix: true });
define("pressure", 101325 / 760, ["Torr"]);
define("pressure", LBF / (INCH * INCH), ["psi"]);
// Force (newton).
define("force", 1, ["N"], { prefix: true });
define("force", 1e-5, ["dyn", "dy"], { prefix: true });
define("force", LBF, ["lbf"]);
define("force", 0.00980665, ["pond"], { prefix: true });
// Energy (joule).
define("energy", 1, ["J"], { prefix: true });
define("energy", 1e-7, ["e"], { prefix: true });
define("energy", 4.184, ["c"], { prefix: true });
define("energy", 4.1868, ["cal"], { prefix: true });
define("energy", 1.602176487e-19, ["eV", "ev"], { prefix: true });
define("energy", HORSEPOWER * 3600, ["HPh", "hh"]);
define("energy", 3600, ["Wh", "wh"], { prefix: true });
define("energy", 0.0421401100938048, ["flb"]);
define("energy", 1055.05585262, ["BTU", "btu"]);
// Power (watt).
define("power", 1, ["W", "w"], { prefix: true });
define("power", HORSEPOWER, ["HP", "h"]);
define("power", 735.49875, ["PS"]);
// Magnetism (tesla).
define("magnetism", 1, ["T"], { prefix: true });
define("magnetism", 1e-4, ["ga"], { prefix: true });
// Temperature is affine; the factor field is unused.
define("temperature", 1, ["C", "cel"]);
define("temperature", 1, ["F", "fah"]);
define("temperature", 1, ["K", "kel"], { prefix: true });
define("temperature", 1, ["Rank"]);
define("temperature", 1, ["Reau"]);
// Volume (cubic meter).
define("volume", 1e-3, ["l", "L", "lt"], { prefix: true });
define("volume", US_GALLON / 768, ["tsp"]);
define("volume", 5e-6, ["tspm"]);
define("volume", US_GALLON / 256, ["tbs"]);
define("volume", US_GALLON / 128, ["oz"]);
define("volume", US_GALLON / 16, ["cup"]);
define("volume", US_GALLON / 8, ["pt", "us_pt"]);
define("volume", UK_GALLON / 8, ["uk_pt"]);
define("volume", US_GALLON / 4, ["qt"]);
define("volume", UK_GALLON / 4, ["uk_qt"]);
define("volume", US_GALLON, ["gal"]);
define("volume", UK_GALLON, ["uk_gal"]);
define("volume", 1, ["m3"], { prefix: true, power: 3 });
define("volume", MILE ** 3, ["mi3"]);
define("volume", NAUTICAL_MILE ** 3, ["Nmi3"]);
define("volume", INCH ** 3, ["in3"]);
define("volume", FOOT ** 3, ["ft3"]);
define("volume", YARD ** 3, ["yd3"]);
define("volume", 1e-30, ["ang3"], { prefix: true, power: 3 });
define("volume", PICA_POINT ** 3, ["Pica3", "Picapt3"]);
define("volume", PICA ** 3, ["pica3"]);
define("volume", US_GALLON * 42, ["barrel"]);
define("volume", 0.03523907016688, ["bushel"]);
define("volume", FOOT ** 3 * 100, ["regton", "GRT"]);
define("volume", FOOT ** 3 * 40, ["MTON"]);
define("volume", LIGHT_YEAR ** 3, ["ly3"]);
// Area (square meter).
define("area", 1, ["m2"], { prefix: true, power: 2 });
define("area", MILE ** 2, ["mi2"]);
define("area", NAUTICAL_MILE ** 2, ["Nmi2"]);
define("area", INCH ** 2, ["in2"]);
define("area", FOOT ** 2, ["ft2"]);
define("area", YARD ** 2, ["yd2"]);
define("area", 1e-20, ["ang2"], { prefix: true, power: 2 });
define("area", PICA_POINT ** 2, ["Pica2", "Picapt2"]);
define("area", PICA ** 2, ["pica2"]);
define("area", 2500, ["Morgen"]);
define("area", 100, ["ar"], { prefix: true });
define("area", 4046.8564224, ["uk_acre"]);
define("area", 4046.87260987425, ["us_acre"]);
define("area", LIGHT_YEAR ** 2, ["ly2"]);
define("area", 10000, ["ha"]);
// Speed (meter per second).
define("speed", 1, ["m/s", "m/sec"], { prefix: true });
define("speed", 1 / 3600, ["m/h", "m/hr"], { prefix: true });
define("speed", MILE / 3600, ["mph"]);
define("speed", NAUTICAL_MILE / 3600, ["kn"]);
define("speed", (6080 * FOOT) / 3600, ["admkn"]);
// Information (bit).
define("information", 1, ["bit"], { prefix: true });
define("information", 8, ["byte"], { prefix: true });

const METRIC_PREFIXES: Record<string, number> = {
  Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 1e1, e: 1e1,
  d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24,
};
const BINARY_PREFIXES: Record<string, number> = {
  Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10,
};

interface ResolvedUnit {
  name: string;
  unit: UnitDefinition;
  scale: number;
}

function resolveUnit(name: string): ResolvedUnit | null {
  const exact = UNITS[name];
  if (exact) return { name, unit: exact, scale: 1 };
  for (const [prefix, size] of Object.entries(BINARY_PREFIXES)) {
    if (!name.startsWith(prefix)) continue;
    const unit = UNITS[name.slice(prefix.length)];
    if (unit?.prefix && unit.category === "information") return { name: name.slice(prefix.length), unit, scale: size };
  }
  for (const prefix of ["da", ...Object.keys(METRIC_PREFIXES).filter((key) => key !== "da")]) {
    if (!name.startsWith(prefix) || name.length === prefix.length) continue;
    const baseName = name.slice(prefix.length);
    const unit = UNITS[baseName];
    if (unit?.prefix) return { name: baseName, unit, scale: METRIC_PREFIXES[prefix] ** (unit.power ?? 1) };
  }
  return null;
}

function toKelvin(value: number, name: string): number {
  switch (name) {
    case "C":
    case "cel":
      return value + 273.15;
    case "F":
    case "fah":
      return ((value - 32) * 5) / 9 + 273.15;
    case "Rank":
      return (value * 5) / 9;
    case "Reau":
      return value * 1.25 + 273.15;
    default:
      return value;
  }
}

function convertTemperature(value: number, from: ResolvedUnit, to: ResolvedUnit): number {
  const source = value * from.scale;
  const fromCelsius = from.name === "C" || from.name === "cel";
  const toCelsius = to.name === "C" || to.name === "cel";
  const fromFahrenheit = from.name === "F" || from.name === "fah";
  const toFahrenheit = to.name === "F" || to.name === "fah";
  // Direct Celsius/Fahrenheit conversions avoid the 273.15 round trip.
  if (fromFahrenheit && toCelsius) return ((source - 32) * 5) / 9;
  if (fromCelsius && toFahrenheit) return (source * 9) / 5 + 32;
  const kelvin = toKelvin(source, from.name);
  let result: number;
  switch (to.name) {
    case "C":
    case "cel":
      result = kelvin - 273.15;
      break;
    case "F":
    case "fah":
      result = ((kelvin - 273.15) * 9) / 5 + 32;
      break;
    case "Rank":
      result = (kelvin * 9) / 5;
      break;
    case "Reau":
      result = (kelvin - 273.15) * 0.8;
      break;
    default:
      result = kelvin;
  }
  return result / to.scale;
}

/** Round away binary noise so CONVERT(1,"lbm","kg") is 0.45359237 as in Excel. */
function significant15(value: number): number {
  if (value === 0 || !Number.isFinite(value)) return value;
  return Number(value.toPrecision(15));
}

// ---- Complex numbers -----------------------------------------------------------------------

interface Complex {
  re: number;
  im: number;
}

const REAL_PATTERN = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

function parseComplexText(text: string): { value: Complex; suffix: "i" | "j" | null } | null {
  const source = text;
  if (source === "") return { value: { re: 0, im: 0 }, suffix: null };
  const last = source[source.length - 1];
  if (last !== "i" && last !== "j") {
    return REAL_PATTERN.test(source) ? { value: { re: Number(source), im: 0 }, suffix: null } : null;
  }
  const body = source.slice(0, -1);
  let split = -1;
  for (let index = body.length - 1; index > 0; index -= 1) {
    const character = body[index];
    if ((character === "+" || character === "-") && body[index - 1] !== "e" && body[index - 1] !== "E") {
      split = index;
      break;
    }
  }
  const realText = split >= 0 ? body.slice(0, split) : "";
  const imaginaryText = split >= 0 ? body.slice(split) : body;
  if (realText !== "" && !REAL_PATTERN.test(realText)) return null;
  let imaginary: number;
  if (imaginaryText === "" || imaginaryText === "+") imaginary = 1;
  else if (imaginaryText === "-") imaginary = -1;
  else if (REAL_PATTERN.test(imaginaryText)) imaginary = Number(imaginaryText);
  else return null;
  return { value: { re: realText === "" ? 0 : Number(realText), im: imaginary }, suffix: last };
}

interface ComplexArgument {
  value: Complex;
  suffix: "i" | "j" | null;
}

function complexScalar(scalar: unknown): ComplexArgument | EvaluationError {
  if (isEvaluationError(scalar)) return scalar;
  if (typeof scalar === "boolean") return valueError();
  if (scalar === null || scalar === undefined) return { value: { re: 0, im: 0 }, suffix: null };
  if (typeof scalar === "number") return { value: { re: scalar, im: 0 }, suffix: null };
  const parsed = parseComplexText(String(scalar));
  return parsed ?? numError();
}

function complexArg(value: Value | undefined): ComplexArgument | EvaluationError {
  return complexScalar(scalarArgument(value ?? null));
}

function mergeSuffix(current: "i" | "j" | null, next: "i" | "j" | null): "i" | "j" | null | EvaluationError {
  if (next === null) return current;
  if (current !== null && current !== next) return valueError();
  return next;
}

function formatComplex(value: Complex, suffix: "i" | "j" | null): Value {
  const re = finiteResult(value.re);
  const im = finiteResult(value.im);
  if (isEvaluationError(re)) return re;
  if (isEvaluationError(im)) return im;
  const unit = suffix ?? "i";
  const realText = numberToText(re);
  if (im === 0) return realText;
  const imaginaryText = im === 1 ? "" : im === -1 ? "-" : numberToText(im);
  if (re === 0) return `${imaginaryText}${unit}`;
  const sign = im > 0 ? "+" : "";
  return `${realText}${sign}${imaginaryText}${unit}`;
}

const complex = {
  add: (a: Complex, b: Complex): Complex => ({ re: a.re + b.re, im: a.im + b.im }),
  sub: (a: Complex, b: Complex): Complex => ({ re: a.re - b.re, im: a.im - b.im }),
  mul: (a: Complex, b: Complex): Complex => ({ re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re }),
  div: (a: Complex, b: Complex): Complex => {
    const denominator = b.re * b.re + b.im * b.im;
    return { re: (a.re * b.re + a.im * b.im) / denominator, im: (a.im * b.re - a.re * b.im) / denominator };
  },
  abs: (a: Complex): number => Math.hypot(a.re, a.im),
  exp: (a: Complex): Complex => {
    const scale = Math.exp(a.re);
    return { re: scale * Math.cos(a.im), im: scale * Math.sin(a.im) };
  },
  ln: (a: Complex): Complex => ({ re: Math.log(Math.hypot(a.re, a.im)), im: Math.atan2(a.im, a.re) }),
  sin: (a: Complex): Complex => ({ re: Math.sin(a.re) * Math.cosh(a.im), im: Math.cos(a.re) * Math.sinh(a.im) }),
  cos: (a: Complex): Complex => ({ re: Math.cos(a.re) * Math.cosh(a.im), im: -Math.sin(a.re) * Math.sinh(a.im) }),
  sinh: (a: Complex): Complex => ({ re: Math.sinh(a.re) * Math.cos(a.im), im: Math.cosh(a.re) * Math.sin(a.im) }),
  cosh: (a: Complex): Complex => ({ re: Math.cosh(a.re) * Math.cos(a.im), im: Math.sinh(a.re) * Math.sin(a.im) }),
};

const ONE: Complex = { re: 1, im: 0 };
const isZero = (a: Complex) => a.re === 0 && a.im === 0;

/** A one-argument complex function returning complex text (or a number via `numeric`). */
function complexUnary(compute: (value: Complex) => Complex | EvaluationError) {
  return spec(1, 1, (values) => {
    const argument = complexArg(values[0]);
    if (isEvaluationError(argument)) return argument;
    const result = compute(argument.value);
    return isEvaluationError(result) ? result : formatComplex(result, argument.suffix);
  }, { liftArgs: "all" });
}

function complexNumeric(compute: (value: Complex) => number | EvaluationError) {
  return spec(1, 1, (values) => {
    const argument = complexArg(values[0]);
    if (isEvaluationError(argument)) return argument;
    const result = compute(argument.value);
    return isEvaluationError(result) ? result : finiteResult(result);
  }, { liftArgs: "all" });
}

/** IMSUM/IMPRODUCT: any number of complex values or ranges of them (blank cells skipped). */
function complexFold(initial: Complex, combine: (left: Complex, right: Complex) => Complex) {
  return spec(1, 255, (values) => {
    let result = initial;
    let suffix: "i" | "j" | null = null;
    for (const entry of collectValues(values)) {
      if (entry.fromRange && entry.value === null) continue;
      const argument = complexScalar(entry.value);
      if (isEvaluationError(argument)) return argument;
      const merged = mergeSuffix(suffix, argument.suffix);
      if (isEvaluationError(merged)) return merged;
      suffix = merged;
      result = combine(result, argument.value);
    }
    return formatComplex(result, suffix);
  });
}

function complexBinary(compute: (left: Complex, right: Complex) => Complex | EvaluationError) {
  return spec(2, 2, (values) => {
    const left = complexArg(values[0]);
    if (isEvaluationError(left)) return left;
    const right = complexArg(values[1]);
    if (isEvaluationError(right)) return right;
    const suffix = mergeSuffix(left.suffix, right.suffix);
    if (isEvaluationError(suffix)) return suffix;
    const result = compute(left.value, right.value);
    return isEvaluationError(result) ? result : formatComplex(result, suffix);
  }, { liftArgs: "all" });
}

function reciprocal(compute: (value: Complex) => Complex) {
  return (value: Complex): Complex | EvaluationError => {
    const denominator = compute(value);
    if (isZero(denominator)) return numError();
    return complex.div(ONE, denominator);
  };
}

// ---- Bessel functions ----------------------------------------------------------------------

const GAUSS_NODES = [0.1488743389816312, 0.4333953941292472, 0.6794095682990244, 0.8650633666889845, 0.9739065285171717];
const GAUSS_WEIGHTS = [0.2955242247147529, 0.2692667193099963, 0.219086362515982, 0.1494513491505806, 0.0666713443086881];

/** Composite 10-point Gauss-Legendre quadrature. */
function integrate(f: (t: number) => number, from: number, to: number, panels: number): number {
  const width = (to - from) / panels;
  let total = 0;
  for (let panel = 0; panel < panels; panel += 1) {
    const middle = from + (panel + 0.5) * width;
    const half = width / 2;
    let sum = 0;
    for (let index = 0; index < GAUSS_NODES.length; index += 1) {
      const offset = half * GAUSS_NODES[index];
      sum += GAUSS_WEIGHTS[index] * (f(middle - offset) + f(middle + offset));
    }
    total += sum * half;
  }
  return total;
}

/** Power series sum_k (-1)^k (x/2)^(2k+n) / (k!(k+n)!) (sign = -1) or its I variant (sign = 1). */
function besselSeries(x: number, n: number, sign: 1 | -1): number {
  const half = x / 2;
  let term = 1;
  for (let index = 1; index <= n; index += 1) term *= half / index;
  let sum = term;
  const square = half * half;
  for (let k = 1; k < 10_000; k += 1) {
    term *= (sign * square) / (k * (k + n));
    sum += term;
    if (Math.abs(term) <= Math.abs(sum) * 1e-17) break;
  }
  return sum;
}

/** Hankel asymptotic expansion for J_n and Y_n at large x. */
function besselAsymptotic(x: number, n: number): { j: number; y: number } {
  const mu = 4 * n * n;
  let p = 1;
  let q = 0;
  let term = 1;
  let previous = Infinity;
  for (let k = 1; k < 200; k += 1) {
    term *= (mu - (2 * k - 1) ** 2) / (k * 8 * x);
    if (Math.abs(term) > previous) break;
    previous = Math.abs(term);
    if (k % 2 === 1) q += (k % 4 === 1 ? 1 : -1) * term;
    else p += (k % 4 === 2 ? -1 : 1) * term;
    if (Math.abs(term) < 1e-17) break;
  }
  const chi = x - (n / 2 + 0.25) * Math.PI;
  const scale = Math.sqrt(2 / (Math.PI * x));
  return { j: scale * (p * Math.cos(chi) - q * Math.sin(chi)), y: scale * (p * Math.sin(chi) + q * Math.cos(chi)) };
}

function besselJ(x: number, n: number): number {
  if (x < 0) return (n % 2 === 0 ? 1 : -1) * besselJ(-x, n);
  if (x === 0) return n === 0 ? 1 : 0;
  if (x <= 5 || (x * x) / 4 < n + 1) return besselSeries(x, n, -1);
  if (x > 25 && n <= Math.sqrt(x)) return besselAsymptotic(x, n).j;
  // J_n(x) = (1/2pi) * integral over a full period of cos(n t - x sin t): the trapezoid rule
  // converges exponentially for this smooth periodic integrand.
  const points = 2 * Math.ceil(x + n) + 64;
  let sum = 0;
  for (let index = 0; index < points; index += 1) {
    const t = (2 * Math.PI * index) / points;
    sum += Math.cos(n * t - x * Math.sin(t));
  }
  return sum / points;
}

function besselY01(x: number, n: 0 | 1): number {
  if (x > 25) return besselAsymptotic(x, n).y;
  const oscillating = integrate((t) => Math.sin(x * Math.sin(t) - n * t), 0, Math.PI, Math.ceil(x) + 8) / Math.PI;
  const limit = Math.asinh(50 / x);
  const decaying =
    integrate((t) => (n === 0 ? 2 : 2 * Math.sinh(t)) * Math.exp(-x * Math.sinh(t)), 0, limit, 64) / Math.PI;
  return oscillating - decaying;
}

function besselY(x: number, n: number): number {
  let previous = besselY01(x, 0);
  if (n === 0) return previous;
  let current = besselY01(x, 1);
  for (let order = 1; order < n; order += 1) {
    const next = ((2 * order) / x) * current - previous;
    previous = current;
    current = next;
    if (!Number.isFinite(current)) return current;
  }
  return current;
}

function besselK01(x: number, n: 0 | 1): number {
  // K_n(x) = integral_0^inf exp(-x cosh t) cosh(n t) dt, scaled by exp(-x) for range.
  const limit = Math.acosh(1 + 50 / x);
  const scaled = integrate((t) => Math.exp(-x * (Math.cosh(t) - 1)) * (n === 0 ? 1 : Math.cosh(t)), 0, limit, 64);
  return scaled * Math.exp(-x);
}

function besselK(x: number, n: number): number {
  let previous = besselK01(x, 0);
  if (n === 0) return previous;
  let current = besselK01(x, 1);
  for (let order = 1; order < n; order += 1) {
    const next = previous + ((2 * order) / x) * current;
    previous = current;
    current = next;
    if (!Number.isFinite(current)) return current;
  }
  return current;
}

function besselSpec(compute: (x: number, n: number) => number, positiveOnly: boolean) {
  return numericFunction([undefined, undefined], (x, rawN) => {
    const n = Math.trunc(rawN);
    if (n < 0) return numError();
    if (positiveOnly && x <= 0) return numError();
    return compute(x, n);
  });
}

// ---- specs ---------------------------------------------------------------------------------

export const ENGINEERING_FUNCTIONS: Specs = {
  BIN2DEC: radixConversion(BIN, DEC),
  BIN2HEX: radixConversion(BIN, HEX),
  BIN2OCT: radixConversion(BIN, OCT),
  DEC2BIN: radixConversion(DEC, BIN),
  DEC2HEX: radixConversion(DEC, HEX),
  DEC2OCT: radixConversion(DEC, OCT),
  HEX2BIN: radixConversion(HEX, BIN),
  HEX2DEC: radixConversion(HEX, DEC),
  HEX2OCT: radixConversion(HEX, OCT),
  OCT2BIN: radixConversion(OCT, BIN),
  OCT2DEC: radixConversion(OCT, DEC),
  OCT2HEX: radixConversion(OCT, HEX),
  BITAND: bitwise((left, right) => left & right),
  BITOR: bitwise((left, right) => left | right),
  BITXOR: bitwise((left, right) => left ^ right),
  BITLSHIFT: bitShift(1),
  BITRSHIFT: bitShift(-1),
  DELTA: numericFunction([undefined, 0], (left, right) => (left === right ? 1 : 0)),
  GESTEP: numericFunction([undefined, 0], (number, step) => (number >= step ? 1 : 0)),
  CONVERT: spec(3, 3, (values) => {
    const number = numberArg(values[0]);
    if (isEvaluationError(number)) return number;
    const fromName = toText(scalarArgument(values[1]));
    if (isEvaluationError(fromName)) return fromName;
    const toName = toText(scalarArgument(values[2]));
    if (isEvaluationError(toName)) return toName;
    const from = resolveUnit(fromName);
    const to = resolveUnit(toName);
    if (!from || !to || from.unit.category !== to.unit.category) return naError();
    const result =
      from.unit.category === "temperature"
        ? convertTemperature(number, from, to)
        : (number * from.unit.factor * from.scale) / (to.unit.factor * to.scale);
    return finiteResult(significant15(result));
  }, { liftArgs: "all" }),

  COMPLEX: spec(2, 3, (values) => {
    const re = numberArg(values[0]);
    if (isEvaluationError(re)) return re;
    const im = numberArg(values[1]);
    if (isEvaluationError(im)) return im;
    let suffix: "i" | "j" = "i";
    if (values[2] !== undefined) {
      const text = toText(scalarArgument(values[2]));
      if (isEvaluationError(text)) return text;
      if (text !== "" && text !== "i" && text !== "j") return valueError();
      if (text === "j") suffix = "j";
    }
    return formatComplex({ re, im }, suffix);
  }, { liftArgs: "all" }),
  IMABS: complexNumeric((value) => complex.abs(value)),
  IMAGINARY: complexNumeric((value) => value.im),
  IMREAL: complexNumeric((value) => value.re),
  IMARGUMENT: complexNumeric((value) => (isZero(value) ? divError() : Math.atan2(value.im, value.re))),
  IMCONJUGATE: complexUnary((value) => ({ re: value.re, im: -value.im })),
  IMCOS: complexUnary(complex.cos),
  IMCOSH: complexUnary(complex.cosh),
  IMSIN: complexUnary(complex.sin),
  IMSINH: complexUnary(complex.sinh),
  IMTAN: complexUnary((value) => {
    const denominator = complex.cos(value);
    return isZero(denominator) ? numError() : complex.div(complex.sin(value), denominator);
  }),
  IMCOT: complexUnary((value) => {
    const denominator = complex.sin(value);
    return isZero(denominator) ? numError() : complex.div(complex.cos(value), denominator);
  }),
  IMSEC: complexUnary(reciprocal(complex.cos)),
  IMCSC: complexUnary(reciprocal(complex.sin)),
  IMSECH: complexUnary(reciprocal(complex.cosh)),
  IMCSCH: complexUnary(reciprocal(complex.sinh)),
  IMEXP: complexUnary(complex.exp),
  IMLN: complexUnary((value) => (isZero(value) ? numError() : complex.ln(value))),
  IMLOG10: complexUnary((value) => {
    if (isZero(value)) return numError();
    const ln = complex.ln(value);
    return { re: ln.re / Math.LN10, im: ln.im / Math.LN10 };
  }),
  IMLOG2: complexUnary((value) => {
    if (isZero(value)) return numError();
    const ln = complex.ln(value);
    return { re: ln.re / Math.LN2, im: ln.im / Math.LN2 };
  }),
  IMSQRT: complexUnary((value) => {
    const modulus = Math.sqrt(complex.abs(value));
    const angle = Math.atan2(value.im, value.re) / 2;
    return { re: modulus * Math.cos(angle), im: modulus * Math.sin(angle) };
  }),
  IMPOWER: spec(2, 2, (values) => {
    const argument = complexArg(values[0]);
    if (isEvaluationError(argument)) return argument;
    const power = numberArg(values[1]);
    if (isEvaluationError(power)) return power;
    const { value } = argument;
    if (isZero(value)) return power > 0 ? formatComplex({ re: 0, im: 0 }, argument.suffix) : numError();
    // Polar form, as Excel computes it (IMPOWER("2+3i",3) = -46+9.00000000000001i).
    const modulus = complex.abs(value) ** power;
    const angle = Math.atan2(value.im, value.re) * power;
    return formatComplex({ re: modulus * Math.cos(angle), im: modulus * Math.sin(angle) }, argument.suffix);
  }, { liftArgs: "all" }),
  IMDIV: complexBinary((left, right) => (isZero(right) ? numError() : complex.div(left, right))),
  IMSUB: complexBinary(complex.sub),
  IMSUM: complexFold({ re: 0, im: 0 }, complex.add),
  IMPRODUCT: complexFold(ONE, complex.mul),

  BESSELJ: besselSpec(besselJ, false),
  BESSELY: besselSpec(besselY, true),
  BESSELI: besselSpec((x, n) => besselSeries(x, n, 1), false),
  BESSELK: besselSpec(besselK, true),
};
