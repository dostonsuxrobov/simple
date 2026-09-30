/**
 * Goal Seek: find x so that f(x) reaches a target. Secant steps from the current value,
 * switching to bisection once a sign change brackets the root; up to Excel's default of
 * 100 iterations with a 0.001 maximum change.
 */
export interface GoalSeekResult {
  found: boolean
  value: number
  /** f(value) at the end of the search. */
  achieved: number
  iterations: number
}

export interface GoalSeekOptions {
  maxIterations?: number
  /** Stop once |f(x) - target| is at most this (Excel's "Maximum change"). */
  tolerance?: number
}

export function goalSeek(evaluate: (x: number) => number | null, target: number, start: number, options: GoalSeekOptions = {}): GoalSeekResult {
  const maxIterations = options.maxIterations ?? 100
  const tolerance = options.tolerance ?? 0.001
  // Keep refining well past the reporting tolerance: secant steps converge fast, and a
  // result like 19.99965 for a target of 20 looks wrong even when it is "within 0.001".
  const tight = Math.min(tolerance, Math.max(1e-12, Math.abs(target) * 1e-10))
  const g = (x: number) => {
    const value = evaluate(x)
    return value === null || !Number.isFinite(value) ? null : value - target
  }
  let x0 = Number.isFinite(start) ? start : 0
  let f0 = g(x0)
  if (f0 === null) {
    x0 = 0
    f0 = g(0)
    if (f0 === null) return { found: false, value: start, achieved: NaN, iterations: 0 }
  }
  if (Math.abs(f0) <= tight) return { found: true, value: x0, achieved: f0 + target, iterations: 0 }
  let x1 = x0 === 0 ? 0.01 : x0 * 1.01
  let f1 = g(x1)
  let bracket: [number, number, number, number] | null = null
  let best = { x: x0, f: f0 }
  for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
    if (!bracket && f1 !== null) {
      if (Math.abs(f1) < Math.abs(best.f)) best = { x: x1, f: f1 }
      if (Math.abs(f1) <= tight) return { found: true, value: x1, achieved: f1 + target, iterations: iteration }
      if (Math.sign(f1) !== Math.sign(f0)) bracket = [x0, f0, x1, f1]
    }
    let next: number
    if (bracket) {
      // Regula falsi with a bisection guard keeps the root bracketed.
      const [a, fa, b, fb]: [number, number, number, number] = bracket
      const secant = b - fb * (b - a) / (fb - fa)
      next = Number.isFinite(secant) && secant > Math.min(a, b) && secant < Math.max(a, b) && iteration % 3 ? secant : (a + b) / 2
      const fn = g(next)
      if (fn === null) return { found: false, value: best.x, achieved: best.f + target, iterations: iteration }
      if (Math.abs(fn) < Math.abs(best.f)) best = { x: next, f: fn }
      if (Math.abs(fn) <= tight || Math.abs(b - a) < 1e-12 * Math.max(1, Math.abs(next))) return { found: Math.abs(fn) <= tolerance, value: next, achieved: fn + target, iterations: iteration }
      bracket = Math.sign(fn) === Math.sign(fa) ? [next, fn, b, fb] : [a, fa, next, fn]
      continue
    }
    if (f1 === null) {
      // Step back towards the last good point.
      x1 = (x0 + x1) / 2
      f1 = g(x1)
      continue
    }
    const slope = (f1 - f0) / (x1 - x0)
    next = slope !== 0 && Number.isFinite(slope) ? x1 - f1 / slope : x1 + (x1 - x0) * 2 + 1
    // Damp wild jumps so a flat region does not throw the search to infinity.
    const limit = Math.max(1, Math.abs(x1)) * 1e6
    if (!Number.isFinite(next) || Math.abs(next - x1) > limit) next = x1 + Math.sign(next - x1 || 1) * limit
    x0 = x1
    f0 = f1
    x1 = next
    f1 = g(x1)
  }
  return { found: Math.abs(best.f) <= tolerance, value: best.x, achieved: best.f + target, iterations: maxIterations }
}
