// Special functions and distribution primitives for the statistical worksheet functions:
// log-gamma/gamma, regularized incomplete gamma and beta, erf/erfc (W. J. Cody), the normal
// quantile (Wichura AS241), and a safeguarded Newton/bisection solver for quantiles.

const LN_SQRT_2PI = 0.91893853320467274178; // ln(sqrt(2*pi))
const SQRT2 = Math.SQRT2;
const FPMIN = 1e-300;
const EPS = 1e-16;

// Lanczos approximation, g = 7, n = 9.
const LANCZOS_G = 7;
const LANCZOS = [
  0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313,
  -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6,
  1.5056327351493116e-7,
];

const FACTORIALS: number[] = [1];
for (let index = 1; index <= 170; index += 1) FACTORIALS.push(FACTORIALS[index - 1] * index);

/** n! for integer 0 <= n <= 170 (Infinity beyond). */
export function factorial(n: number): number {
  if (n < 0 || !Number.isInteger(n)) return NaN;
  return n <= 170 ? FACTORIALS[n] : Infinity;
}

/** ln Γ(x) for x > 0. */
export function lnGamma(x: number): number {
  if (!(x > 0) || !Number.isFinite(x)) return x === Infinity ? Infinity : NaN;
  if (Number.isInteger(x) && x <= 171) return Math.log(FACTORIALS[x - 1]);
  if (x < 0.5) {
    // Reflection: Γ(x)Γ(1-x) = π / sin(πx).
    return Math.log(Math.PI / Math.sin(Math.PI * x)) - lnGamma(1 - x);
  }
  if (x >= 10) {
    const inverse = 1 / x;
    const inverse2 = inverse * inverse;
    const series =
      inverse *
      (1 / 12 +
        inverse2 *
          (-1 / 360 +
            inverse2 *
              (1 / 1260 +
                inverse2 * (-1 / 1680 + inverse2 * (1 / 1188 + inverse2 * (-691 / 360360 + inverse2 / 156))))));
    return (x - 0.5) * Math.log(x) - x + LN_SQRT_2PI + series;
  }
  const z = x - 1;
  let sum = LANCZOS[0];
  for (let index = 1; index < LANCZOS.length; index += 1) sum += LANCZOS[index] / (z + index);
  const t = z + LANCZOS_G + 0.5;
  return LN_SQRT_2PI + (z + 0.5) * Math.log(t) - t + Math.log(sum);
}

/** Γ(x); NaN at 0 and the negative integers, ±Infinity on overflow. */
export function gamma(x: number): number {
  if (Number.isInteger(x)) {
    if (x <= 0) return NaN;
    return x <= 171 ? FACTORIALS[x - 1] : Infinity;
  }
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
  if (x > 171.7) return Infinity;
  if (x < 10) {
    const z = x - 1;
    let sum = LANCZOS[0];
    for (let index = 1; index < LANCZOS.length; index += 1) sum += LANCZOS[index] / (z + index);
    const t = z + LANCZOS_G + 0.5;
    return Math.sqrt(2 * Math.PI) * Math.pow(t, z + 0.5) * Math.exp(-t) * sum;
  }
  return Math.exp(lnGamma(x));
}

/** ln B(a, b). */
export function lnBeta(a: number, b: number): number {
  return lnGamma(a) + lnGamma(b) - lnGamma(a + b);
}

/** ln C(n, k) for real n >= k >= 0. */
export function lnCombination(n: number, k: number): number {
  return lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1);
}

function gammaSeries(a: number, x: number): number {
  let term = 1 / a;
  let sum = term;
  let ap = a;
  for (let iteration = 0; iteration < 100_000; iteration += 1) {
    ap += 1;
    term *= x / ap;
    sum += term;
    if (Math.abs(term) < Math.abs(sum) * EPS) break;
  }
  return sum * Math.exp(-x + a * Math.log(x) - lnGamma(a));
}

function gammaContinuedFraction(a: number, x: number): number {
  let b = x + 1 - a;
  let c = 1 / FPMIN;
  let d = 1 / b;
  let h = d;
  for (let index = 1; index < 100_000; index += 1) {
    const an = -index * (index - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = b + an / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) < EPS) break;
  }
  return Math.exp(-x + a * Math.log(x) - lnGamma(a)) * h;
}

/** Regularized lower incomplete gamma P(a, x). */
export function gammaP(a: number, x: number): number {
  if (x <= 0) return 0;
  if (x === Infinity) return 1;
  return x < a + 1 ? gammaSeries(a, x) : 1 - gammaContinuedFraction(a, x);
}

/** Regularized upper incomplete gamma Q(a, x) = 1 - P(a, x). */
export function gammaQ(a: number, x: number): number {
  if (x <= 0) return 1;
  if (x === Infinity) return 0;
  return x < a + 1 ? 1 - gammaSeries(a, x) : gammaContinuedFraction(a, x);
}

function betaContinuedFraction(x: number, a: number, b: number): number {
  const qab = a + b;
  const qap = a + 1;
  const qam = a - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d;
  let h = d;
  for (let m = 1; m < 100_000; m += 1) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) < EPS) break;
  }
  return h;
}

/**
 * Regularized incomplete beta I_x(a, b). `y` is 1 - x, passed explicitly when the caller
 * knows it more precisely than 1 - x (e.g. t²/(ν+t²)).
 */
export function betaRegularized(x: number, a: number, b: number, y = 1 - x): number {
  if (x <= 0) return 0;
  if (y <= 0) return 1;
  const front = Math.exp(a * Math.log(x) + b * Math.log(y) - lnBeta(a, b));
  if (x < (a + 1) / (a + b + 2)) return (front * betaContinuedFraction(x, a, b)) / a;
  return 1 - (front * betaContinuedFraction(y, b, a)) / b;
}

/** 1 - I_x(a, b), computed without cancellation. */
export function betaRegularizedComplement(x: number, a: number, b: number, y = 1 - x): number {
  return betaRegularized(y, b, a, x);
}

// ---- erf / erfc (W. J. Cody, "Rational Chebyshev approximations for the error function") ----

const ERF_A = [3.1611237438705656, 113.86415415105016, 377.485237685302, 3209.3775891384694, 0.18577770618460315];
const ERF_B = [23.601290952344122, 244.02463793444417, 1282.6165260773723, 2844.236833439171];
const ERF_C = [
  0.5641884969886701, 8.883149794388377, 66.11919063714163, 298.635138197400131, 881.952221241769090,
  1712.04761263407058, 2051.07837782607147, 1230.33935479799725, 2.15311535474403846e-8,
];
const ERF_D = [
  15.744926110709835, 117.6939508913125, 537.1811018620099, 1621.3895745666903, 3290.7992357334597,
  4362.619090143247, 3439.3676741437216, 1230.3393548037495,
];
const ERF_P = [0.30532663496123236, 0.36034489994980445, 0.12578172611122926, 0.016083785148742275, 6.587491615298378e-4, 0.016315387137302097];
const ERF_Q = [2.568520192289822, 1.8729528499234604, 0.5279051029514285, 0.06051834131244132, 0.0023352049762686918];
const ONE_OVER_SQRT_PI = 0.56418958354775628695;

/** erfc(|x|) for |x| > 0.46875 (Cody's second and third ranges). */
function erfcTail(y: number): number {
  let result: number;
  if (y <= 4) {
    let numerator = ERF_C[8] * y;
    let denominator = y;
    for (let index = 0; index < 7; index += 1) {
      numerator = (numerator + ERF_C[index]) * y;
      denominator = (denominator + ERF_D[index]) * y;
    }
    result = (numerator + ERF_C[7]) / (denominator + ERF_D[7]);
  } else {
    if (y >= 27) return 0;
    const inverse = 1 / (y * y);
    let numerator = ERF_P[5] * inverse;
    let denominator = inverse;
    for (let index = 0; index < 4; index += 1) {
      numerator = (numerator + ERF_P[index]) * inverse;
      denominator = (denominator + ERF_Q[index]) * inverse;
    }
    result = (inverse * (numerator + ERF_P[4])) / (denominator + ERF_Q[4]);
    result = (ONE_OVER_SQRT_PI - result) / y;
  }
  const rounded = Math.trunc(y * 16) / 16;
  const delta = (y - rounded) * (y + rounded);
  return Math.exp(-rounded * rounded) * Math.exp(-delta) * result;
}

function erfSmall(x: number): number {
  const square = Math.abs(x) > 1.11e-16 ? x * x : 0;
  let numerator = ERF_A[4] * square;
  let denominator = square;
  for (let index = 0; index < 3; index += 1) {
    numerator = (numerator + ERF_A[index]) * square;
    denominator = (denominator + ERF_B[index]) * square;
  }
  return (x * (numerator + ERF_A[3])) / (denominator + ERF_B[3]);
}

export function erf(x: number): number {
  const y = Math.abs(x);
  if (y <= 0.46875) return erfSmall(x);
  const value = 1 - erfcTail(y);
  return x < 0 ? -value : value;
}

export function erfc(x: number): number {
  const y = Math.abs(x);
  if (y <= 0.46875) return 1 - erfSmall(x);
  const tail = erfcTail(y);
  return x < 0 ? 2 - tail : tail;
}

// ---- Normal distribution -----------------------------------------------------------------

export function normalPdf(z: number): number {
  return Math.exp(-0.5 * z * z - LN_SQRT_2PI);
}

export function normalCdf(z: number): number {
  return 0.5 * erfc(-z / SQRT2);
}

/** Standard normal quantile (Wichura, AS241 PPND16; ~1e-16 relative accuracy). */
export function normalInv(p: number): number {
  if (!(p > 0 && p < 1)) return p === 0 ? -Infinity : p === 1 ? Infinity : NaN;
  const q = p - 0.5;
  if (Math.abs(q) <= 0.425) {
    const r = 0.180625 - q * q;
    return (
      (q *
        (((((((r * 2509.0809287301226727 + 33430.575583588128105) * r + 67265.770927008700853) * r +
          45921.953931549871457) * r + 13731.693765509461125) * r + 1971.5909503065514427) * r +
          133.14166789178437745) * r + 3.387132872796366608)) /
      (((((((r * 5226.495278852545925 + 28729.085735721942674) * r + 39307.89580009271061) * r +
        21213.794301586595867) * r + 5394.1960214247511077) * r + 687.1870074920579083) * r +
        42.313330701600911252) * r + 1)
    );
  }
  let r = q < 0 ? p : 1 - p;
  r = Math.sqrt(-Math.log(r));
  let value: number;
  if (r <= 5) {
    r -= 1.6;
    value =
      (((((((r * 7.7454501427834140764e-4 + 0.0227238449892691845833) * r + 0.24178072517745061177) * r +
        1.27045825245236838258) * r + 3.64784832476320460504) * r + 5.7694972214606914055) * r +
        4.6303378461565452959) * r + 1.42343711074968357734) /
      (((((((r * 1.05075007164441684324e-9 + 5.475938084995344946e-4) * r + 0.0151986665636164571966) * r +
        0.14810397642748007459) * r + 0.68976733498510000455) * r + 1.6763848301838038494) * r +
        2.05319162663775882187) * r + 1);
  } else {
    r -= 5;
    value =
      (((((((r * 2.01033439929228813265e-7 + 2.71155556874348757815e-5) * r + 0.0012426609473880784386) * r +
        0.026532189526576123093) * r + 0.29656057182850489123) * r + 1.7848265399172913358) * r +
        5.4637849111641143699) * r + 6.6579046435011037772) /
      (((((((r * 2.04426310338993978564e-15 + 1.4215117583164458887e-7) * r + 1.8463183175100546818e-5) * r +
        7.868691311456132591e-4) * r + 0.0148753612908506148525) * r + 0.13692988092273580531) * r +
        0.59983220655588793769) * r + 1);
  }
  return q < 0 ? -value : value;
}

// ---- Root finding --------------------------------------------------------------------------

/**
 * Root of an increasing function g on [lo, hi] (g(lo) <= 0 <= g(hi)), using Newton steps with
 * the derivative `dg` when they stay inside the bracket and shrink fast enough (rtsafe),
 * bisection otherwise.
 */
export function solveIncreasing(
  g: (x: number) => number,
  lo: number,
  hi: number,
  dg?: (x: number) => number,
): number {
  let x = 0.5 * (lo + hi);
  let dxOld = hi - lo;
  let dx = dxOld;
  for (let iteration = 0; iteration < 1_000; iteration += 1) {
    const value = g(x);
    if (Number.isNaN(value)) return NaN;
    if (value === 0) return x;
    if (value < 0) lo = x;
    else hi = x;
    const derivative = dg ? dg(x) : NaN;
    let next: number;
    if (
      Number.isFinite(derivative) &&
      derivative > 0 &&
      x - value / derivative > lo &&
      x - value / derivative < hi &&
      Math.abs(2 * value) <= Math.abs(dxOld * derivative)
    ) {
      dxOld = dx;
      dx = value / derivative;
      next = x - dx;
    } else {
      dxOld = dx;
      next = lo + 0.5 * (hi - lo);
      dx = x - next;
    }
    if (next === x || Math.abs(dx) <= 4e-16 * Math.abs(next) || Math.abs(dx) < 1e-300) return next;
    if (hi - lo <= 4e-16 * Math.max(Math.abs(lo), Math.abs(hi))) return next;
    x = next;
  }
  return x;
}

export interface ContinuousDistribution {
  cdf: (x: number) => number;
  /** Survival function 1 - cdf(x), computed without cancellation. */
  sf: (x: number) => number;
  pdf: (x: number) => number;
}

/**
 * Quantile of a continuous distribution with support [lower, upper]: the x with cdf(x) = p
 * (q = 1 - p, passed separately so upper-tail inverses keep full precision). `start` seeds
 * the upward bracket search when the support is unbounded.
 */
export function invertDistribution(
  distribution: ContinuousDistribution,
  p: number,
  q: number,
  lower: number,
  upper: number,
  start: number,
): number {
  if (p <= 0) return lower;
  if (q <= 0) return upper;
  const useLower = p <= q;
  const g = useLower
    ? (x: number) => distribution.cdf(x) - p
    : (x: number) => q - distribution.sf(x);
  let lo = lower;
  let hi = upper;
  if (!Number.isFinite(upper)) {
    lo = lower;
    hi = Math.max(start, Number.MIN_VALUE);
    let guard = 0;
    while (g(hi) < 0) {
      lo = hi;
      hi *= 2;
      guard += 1;
      if (!Number.isFinite(hi) || guard > 2_100) return Infinity;
    }
  }
  if (!Number.isFinite(lower)) {
    let low = Math.min(-1, -Math.abs(start));
    let guard = 0;
    while (g(low) > 0) {
      hi = low;
      low *= 2;
      guard += 1;
      if (!Number.isFinite(low) || guard > 2_100) return -Infinity;
    }
    lo = low;
  }
  return solveIncreasing(g, lo, hi, distribution.pdf);
}

// ---- Distribution primitives ---------------------------------------------------------------

/** Upper tail P(T > t) of Student's t with ν degrees of freedom. */
export function studentTSf(t: number, df: number): number {
  if (t === 0) return 0.5;
  const square = t * t;
  const x = df / (df + square);
  const y = square / (df + square);
  const half = 0.5 * betaRegularized(x, df / 2, 0.5, y);
  return t > 0 ? half : 1 - half;
}

export function studentTCdf(t: number, df: number): number {
  return studentTSf(-t, df);
}

export function studentTPdf(t: number, df: number): number {
  return Math.exp(
    lnGamma((df + 1) / 2) - lnGamma(df / 2) - 0.5 * Math.log(df * Math.PI) - ((df + 1) / 2) * Math.log1p((t * t) / df),
  );
}

export function studentT(df: number): ContinuousDistribution {
  return { cdf: (t) => studentTCdf(t, df), sf: (t) => studentTSf(t, df), pdf: (t) => studentTPdf(t, df) };
}

/** t >= 0 with P(T > t) = q, for 0 < q <= 0.5. */
export function studentTInvUpper(q: number, df: number): number {
  if (q >= 0.5) return 0;
  return invertDistribution(studentT(df), 1 - q, q, 0, Infinity, 1);
}

export function gammaPdf(x: number, shape: number, scale: number): number {
  if (x < 0) return 0;
  if (x === 0) return shape < 1 ? Infinity : shape === 1 ? 1 / scale : 0;
  return Math.exp((shape - 1) * Math.log(x) - x / scale - lnGamma(shape) - shape * Math.log(scale));
}

export function gammaDistribution(shape: number, scale: number): ContinuousDistribution {
  return {
    cdf: (x) => gammaP(shape, x / scale),
    sf: (x) => gammaQ(shape, x / scale),
    pdf: (x) => gammaPdf(x, shape, scale),
  };
}

/** Quantile of Gamma(shape, scale) for probability p (q = 1 - p). */
export function gammaInv(p: number, q: number, shape: number, scale = 1): number {
  return invertDistribution(gammaDistribution(shape, scale), p, q, 0, Infinity, Math.max(shape * scale, 1e-3));
}

export function betaPdf(x: number, a: number, b: number): number {
  if (x < 0 || x > 1) return 0;
  if (x === 0) return a < 1 ? Infinity : a === 1 ? b : 0;
  if (x === 1) return b < 1 ? Infinity : b === 1 ? a : 0;
  return Math.exp((a - 1) * Math.log(x) + (b - 1) * Math.log1p(-x) - lnBeta(a, b));
}

export function betaDistribution(a: number, b: number): ContinuousDistribution {
  return {
    cdf: (x) => betaRegularized(x, a, b),
    sf: (x) => betaRegularizedComplement(x, a, b),
    pdf: (x) => betaPdf(x, a, b),
  };
}

export function betaInv(p: number, q: number, a: number, b: number): number {
  return invertDistribution(betaDistribution(a, b), p, q, 0, 1, 0.5);
}

export function fPdf(x: number, d1: number, d2: number): number {
  if (x < 0) return 0;
  if (x === 0) return d1 < 2 ? Infinity : d1 === 2 ? 1 : 0;
  return Math.exp(
    0.5 * (d1 * Math.log(d1 * x) + d2 * Math.log(d2) - (d1 + d2) * Math.log(d1 * x + d2)) - Math.log(x) - lnBeta(d1 / 2, d2 / 2),
  );
}

export function fDistribution(d1: number, d2: number): ContinuousDistribution {
  return {
    cdf: (x) => (x <= 0 ? 0 : betaRegularized((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2, d2 / (d1 * x + d2))),
    sf: (x) => (x <= 0 ? 1 : betaRegularized(d2 / (d1 * x + d2), d2 / 2, d1 / 2, (d1 * x) / (d1 * x + d2))),
    pdf: (x) => fPdf(x, d1, d2),
  };
}

export function fInv(p: number, q: number, d1: number, d2: number): number {
  return invertDistribution(fDistribution(d1, d2), p, q, 0, Infinity, 1);
}

export function chiSquarePdf(x: number, df: number): number {
  return gammaPdf(x, df / 2, 2);
}

// ---- Discrete distributions ----------------------------------------------------------------

export function binomialPmf(k: number, n: number, p: number): number {
  if (k < 0 || k > n) return 0;
  if (p === 0) return k === 0 ? 1 : 0;
  if (p === 1) return k === n ? 1 : 0;
  return Math.exp(lnCombination(n, k) + k * Math.log(p) + (n - k) * Math.log1p(-p));
}

/** P(X <= k) for X ~ Binomial(n, p). */
export function binomialCdf(k: number, n: number, p: number): number {
  if (k < 0) return 0;
  if (k >= n) return 1;
  if (n <= 1_000) {
    let sum = 0;
    for (let index = 0; index <= k; index += 1) sum += binomialPmf(index, n, p);
    return Math.min(1, sum);
  }
  if (p === 0) return 1;
  if (p === 1) return 0;
  return betaRegularized(1 - p, n - k, k + 1, p);
}

export function poissonPmf(k: number, mean: number): number {
  if (k < 0) return 0;
  if (mean === 0) return k === 0 ? 1 : 0;
  return Math.exp(-mean + k * Math.log(mean) - lnGamma(k + 1));
}

export function poissonCdf(k: number, mean: number): number {
  if (k < 0) return 0;
  if (mean === 0) return 1;
  if (k < 1_000 && mean < 700) {
    let term = Math.exp(-mean);
    let sum = term;
    for (let index = 1; index <= k; index += 1) {
      term *= mean / index;
      sum += term;
    }
    return Math.min(1, sum);
  }
  return gammaQ(k + 1, mean);
}
