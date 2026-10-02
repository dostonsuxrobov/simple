// src/shared/worker-ops/mask.ts
// Owner: WP4 (seeded empty by WP1). Selection and retouching handlers for the imaging worker:
// flood, feather, expand, contract, heal.
// Contract read by ./index.ts: keep the named export `handlers` (an ImagingHandlers table).
// Handlers run in a module worker or inline in Node tests: no DOM, check `options.signal` between
// rows/stripes (throw an Error named 'AbortError'), report `options.onProgress` at most ~20 times.
// None of these ops is stripe-safe (flood fill, distance transforms and healing read the whole input),
// so they must not be registered with the client's stripeMargins.
// The algorithms are synchronous: inside a worker an 'abort' message is only dispatched when the event
// loop gets a turn, so every handler yields one task before starting. A job cancelled while it waits
// in the worker's message queue therefore stops at once; a job already computing runs to completion
// (its result is discarded by the client).
import type { ImagingHandlers, MaskBuffer, OpOptions, PixelBuffer } from '../../imaging/types.ts'
import { floodFill } from '../../imaging/floodFill.ts'
import { contractMask, expandMask, featherMask } from '../../imaging/mask.ts'
import { healRegion } from '../../imaging/inpaint.ts'

function abortError(): Error {
  const error = new Error('The operation was cancelled.')
  error.name = 'AbortError'
  return error
}

/** Lets already-queued messages (an 'abort' sent right after 'run') reach the worker, then checks. */
async function start(options: OpOptions): Promise<void> {
  if (options?.signal?.aborted) throw abortError()
  await new Promise<void>((resolve) => setTimeout(resolve, 0))
  if (options?.signal?.aborted) throw abortError()
}

function pixels(value: PixelBuffer, label: string): PixelBuffer {
  const data = value?.data as ArrayBufferView | undefined
  if (!value || !data || !ArrayBuffer.isView(data) || data.byteLength !== value.width * value.height * 4) {
    throw new TypeError(`The ${label} pixels are missing or do not match their size.`)
  }
  // Structured clone keeps the typed array kind; accept any byte view (e.g. a Uint8Array) too.
  if (data instanceof Uint8ClampedArray) return value
  return { width: value.width, height: value.height, data: new Uint8ClampedArray(data.buffer, data.byteOffset, data.byteLength) }
}

function mask(value: MaskBuffer, label: string): MaskBuffer {
  const data = value?.data as ArrayBufferView | undefined
  if (!value || !data || !ArrayBuffer.isView(data) || data.byteLength !== value.width * value.height) {
    throw new TypeError(`The ${label} mask is missing or does not match its size.`)
  }
  if (data instanceof Uint8Array) return value
  return { width: value.width, height: value.height, data: new Uint8Array(data.buffer, data.byteOffset, data.byteLength) }
}

function amount(value: number, label: string, max: number): number {
  if (!Number.isFinite(value) || value < 0 || value > max) throw new RangeError(`The ${label} must be between 0 and ${max}.`)
  return value
}

export const handlers: ImagingHandlers = {
  flood: async (input, options) => {
    const src = pixels(input.src, 'source')
    await start(options)
    return floodFill(src, input.seed, input.options, options)
  },
  feather: async (input, options) => {
    const selection = mask(input.mask, 'selection')
    const radius = amount(input.radius, 'feather radius', 1000)
    await start(options)
    return featherMask(selection, radius, options)
  },
  expand: async (input, options) => {
    const selection = mask(input.mask, 'selection')
    const pixelCount = amount(input.pixels, 'expand amount', 500)
    await start(options)
    return expandMask(selection, pixelCount, options)
  },
  contract: async (input, options) => {
    const selection = mask(input.mask, 'selection')
    const pixelCount = amount(input.pixels, 'contract amount', 500)
    await start(options)
    return contractMask(selection, pixelCount, options)
  },
  heal: async (input, options) => {
    const src = pixels(input.src, 'source')
    const hole = mask(input.hole, 'healing')
    await start(options)
    return healRegion(src, hole, input.options ?? {}, options)
  },
}
