// src/shared/worker-ops/color.ts
// Owner: WP2. Pixel-colour and geometry handlers for the imaging worker:
// adjust, filter, look, resample, rotate, warp, histogram, autoEnhance.
// Contract read by ./index.ts: keep the named export `handlers` (an ImagingHandlers table).
// Handlers run in a module worker or inline in Node tests: no DOM, check `options.signal` between
// rows/stripes (throw an Error named 'AbortError'), report `options.onProgress` at most ~20 times.
//
// Long operations run as resumable row runs (imaging/buffer.ts RowRun) and yield to the event loop between
// chunks: a worker only handles its next message (the runtime's 'abort') between tasks, so a handler that
// never yields could not be cancelled. Chunking never changes the result (row runs are chunk-exact).
// Invalid input throws synchronously; a signal that is already aborted throws AbortError synchronously.
// Inline runs receive the caller's own objects, so no handler mutates its input; every result is a new buffer.
// Stripe safety for workerClient striping (stripes do not know their absolute row): 'adjust' is a point
// operation (margin 0; only a dithered Gradient Map's +-0.5 level noise pattern restarts per stripe);
// 'filter' is stripe-exact with imaging/filters.ts stripeMargin(spec) (null when it needs absolute rows);
// 'look' must not be striped (its vignette depends on the position in the whole image).
import type { ImagingHandlers, OpOptions } from '../../imaging/types.ts'
import { startAdjustments } from '../../imaging/adjustments.ts'
import { autoEnhance } from '../../imaging/autoEnhance.ts'
import { runRowsAsync, throwIfAborted } from '../../imaging/buffer.ts'
import { startFilter } from '../../imaging/filters.ts'
import { computeHistogram } from '../../imaging/histogram.ts'
import { startLook } from '../../imaging/looks.ts'
import { startResample } from '../../imaging/resample.ts'
import { startRotate, startWarpPerspective } from '../../imaging/transform.ts'

function requireInput<T>(input: T, op: string): T {
  if (!input || typeof input !== 'object') throw new RangeError(`The imaging operation "${op}" received no input.`)
  return input
}

function checkpoint(options: OpOptions | undefined): void {
  throwIfAborted(options?.signal)
}

export const handlers: ImagingHandlers = Object.freeze({
  adjust: (input, options) => {
    const { src, specs, mask, opacity } = requireInput(input, 'adjust')
    checkpoint(options)
    return runRowsAsync(startAdjustments(src, specs ?? [], mask ?? null, typeof opacity === 'number' ? opacity : 1), options)
  },
  filter: (input, options) => {
    const { src, spec, mask } = requireInput(input, 'filter')
    checkpoint(options)
    return runRowsAsync(startFilter(src, spec, mask ?? null), options)
  },
  look: (input, options) => {
    const { src, look, intensity } = requireInput(input, 'look')
    checkpoint(options)
    return runRowsAsync(startLook(src, look, typeof intensity === 'number' ? intensity : 100), options)
  },
  resample: (input, options) => {
    const { src, width, height, method } = requireInput(input, 'resample')
    checkpoint(options)
    return runRowsAsync(startResample(src, width, height, method ?? 'auto'), options)
  },
  rotate: (input, options) => {
    const { src, degrees, fit, interpolation, background } = requireInput(input, 'rotate')
    checkpoint(options)
    return runRowsAsync(startRotate(src, degrees, fit, interpolation ?? 'bicubic', background ?? { r: 0, g: 0, b: 0, a: 0 }), options)
  },
  warp: (input, options) => {
    const { src, inverse, out, interpolation } = requireInput(input, 'warp')
    checkpoint(options)
    return runRowsAsync(startWarpPerspective(src, inverse, out, interpolation ?? 'bicubic'), options)
  },
  histogram: (input, options) => {
    const { src, mask } = requireInput(input, 'histogram')
    checkpoint(options)
    return computeHistogram(src, mask ?? null)
  },
  autoEnhance: (input, options) => {
    const { src } = requireInput(input, 'autoEnhance')
    checkpoint(options)
    return autoEnhance(src)
  },
})
