// src/shared/workerClient.ts (WP1)
// Main-thread client for the imaging worker pool (design section 5.16).
//   - Lazy pool of min(4, hardwareConcurrency - 1) module workers; a job is only posted to a
//     worker after it reported 'ready', so a worker that cannot start never swallows input.
//   - Input buffers are transferred; views that do not own their whole ArrayBuffer are copied to
//     a compact buffer first (structured clone would otherwise copy the entire backing buffer).
//   - Abort: queued jobs leave the queue, running jobs get an 'abort' message; run() rejects with
//     an Error named 'AbortError' immediately.
//   - Large jobs of stripe-safe ops (registered margins; 'adjust' by default) are split into
//     horizontal stripes with overlap and spread across the pool, which also yields honest progress.
//   - Without Worker (Node tests), or when workers cannot start (CSP, crash on load), handlers run
//     inline on the calling thread.
// DOM-free at module level so Node tests can load it.
import type {
  ImagingClient,
  ImagingHandlers,
  ImagingInput,
  ImagingOp,
  ImagingOutput,
  ImagingRunOptions,
  MaskBuffer,
  OpOptions,
  PixelBuffer,
} from '../imaging/types.ts'
import type { ImagingWorkerRequest, ImagingWorkerResponse } from './worker-ops/index.ts'

/** Ops whose input carries `src` (and optionally a same-size `mask`) and whose output is a same-size PixelBuffer. */
export type StripeableOp = 'adjust' | 'filter' | 'look'
/** Rows of context a stripe needs above and below (0 for point operations), or null to run unsplit. */
export type StripeMargin<K extends StripeableOp> = (input: ImagingInput<K>) => number | null
export type StripeMargins = { readonly [K in StripeableOp]?: StripeMargin<K> }

export interface ImagingClientOptions {
  /** Maximum pool size. Default min(4, hardwareConcurrency - 1), at least 1. */
  readonly workers?: number
  /** Worker factory. Default: the bundled module worker when `Worker` exists. */
  readonly createWorker?: () => Worker
  /** Handlers to run on the calling thread (Node tests; also the fallback when workers fail). */
  readonly inline?: ImagingHandlers
  /** Loads handlers for the inline fallback. Default: dynamic import of ./worker-ops/index.ts. */
  readonly loadInline?: () => Promise<ImagingHandlers>
  /** Stripe-safe ops and their margins. */
  readonly stripeMargins?: StripeMargins
  /** Jobs larger than this many pixels are striped (when a margin is registered). Default 4 MP. */
  readonly stripeThresholdPixels?: number
  /** How long a new worker may take to report 'ready'. Default 15 s. */
  readonly startTimeoutMs?: number
}

export interface PooledImagingClient extends ImagingClient {
  /** 'workers' (pool), 'inline' (calling thread) or 'disposed'. */
  readonly mode: 'workers' | 'inline' | 'disposed'
  /** Jobs queued or running. */
  readonly pending: number
  /** Registers (or with null removes) the stripe margin of an op. */
  setStripeMargin<K extends StripeableOp>(op: K, margin: StripeMargin<K> | null): void
  /** Terminates idle workers; they restart lazily on the next job. */
  trim(): void
}

const DEFAULT_STRIPE_THRESHOLD = 4_000_000
const TARGET_STRIPE_PIXELS = 2_000_000
const MAX_STRIPES = 64
const MAX_START_FAILURES = 3

export function abortError(message = 'The operation was cancelled.'): Error {
  const error = new Error(message)
  error.name = 'AbortError'
  return error
}

function namedError(name: string, message: string): Error {
  const error = new Error(message)
  error.name = name
  return error
}

export function defaultWorkerCount(): number {
  const cores = typeof navigator !== 'undefined' && Number(navigator.hardwareConcurrency) > 0 ? Number(navigator.hardwareConcurrency) : 2
  return Math.max(1, Math.min(4, cores - 1))
}

// ---------------------------------------------------------------------------------------------
// Transfer preparation
// ---------------------------------------------------------------------------------------------

function ownsWholeBuffer(view: ArrayBufferView): boolean {
  return view.buffer instanceof ArrayBuffer && view.byteOffset === 0 && view.byteLength === view.buffer.byteLength
}

function compactCopy(view: ArrayBufferView): ArrayBufferView {
  const bytes = new Uint8Array(view.byteLength)
  bytes.set(new Uint8Array(view.buffer, view.byteOffset, view.byteLength))
  if (view instanceof DataView) return new DataView(bytes.buffer)
  const Constructor = view.constructor as new (buffer: ArrayBuffer) => ArrayBufferView
  return new Constructor(bytes.buffer)
}

function isPlainObject(value: object): boolean {
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

/**
 * Returns a structured-clone-ready copy of `value` and collects the buffers to transfer.
 * Owning views are transferred as-is when `transfer` is true (the caller's view detaches);
 * non-owning views are always replaced by a compact copy, which is then transferred.
 * Class instances shaped like pixel containers (ImageData) become plain { width, height, data }.
 */
export function prepareTransfer(value: unknown, transfer: boolean, buffers: Set<ArrayBuffer> = new Set(), depth = 0): unknown {
  if (value === null || typeof value !== 'object' || depth > 8) return value
  if (ArrayBuffer.isView(value)) {
    if (ownsWholeBuffer(value)) {
      if (transfer && value.byteLength > 0) buffers.add(value.buffer as ArrayBuffer)
      return value
    }
    const copy = compactCopy(value)
    if (copy.byteLength > 0) buffers.add(copy.buffer as ArrayBuffer)
    return copy
  }
  if (value instanceof ArrayBuffer) {
    if (transfer && value.byteLength > 0) buffers.add(value)
    return value
  }
  if (Array.isArray(value)) return value.map((item) => prepareTransfer(item, transfer, buffers, depth + 1))
  if (isPlainObject(value)) {
    const result: Record<string, unknown> = {}
    for (const [key, item] of Object.entries(value)) result[key] = prepareTransfer(item, transfer, buffers, depth + 1)
    return result
  }
  const shaped = value as { width?: unknown; height?: unknown; data?: unknown }
  if (typeof shaped.width === 'number' && typeof shaped.height === 'number' && ArrayBuffer.isView(shaped.data)) {
    return { width: shaped.width, height: shaped.height, data: prepareTransfer(shaped.data, transfer, buffers, depth + 1) }
  }
  return value
}

// ---------------------------------------------------------------------------------------------
// Striping
// ---------------------------------------------------------------------------------------------

interface StripeSource {
  readonly src: PixelBuffer
  readonly mask?: MaskBuffer | null
}

interface Stripe {
  readonly top: number
  readonly bottom: number
}

function isPixelBuffer(value: unknown): value is PixelBuffer {
  const buffer = value as PixelBuffer | null
  return Boolean(buffer && typeof buffer.width === 'number' && typeof buffer.height === 'number'
    && ArrayBuffer.isView(buffer.data) && buffer.data.length === buffer.width * buffer.height * 4)
}

/** Equal stripes of at least 4 x margin rows (overlap overhead <= 50%), or null when splitting does not pay. */
export function planStripes(width: number, height: number, margin: number, workers: number): Stripe[] | null {
  const pixels = width * height
  let count = Math.max(workers, Math.ceil(pixels / TARGET_STRIPE_PIXELS))
  count = Math.min(count, MAX_STRIPES, height)
  if (margin > 0) count = Math.min(count, Math.floor(height / (4 * margin)))
  if (count < 2) return null
  const rows = Math.ceil(height / count)
  const stripes: Stripe[] = []
  for (let top = 0; top < height; top += rows) stripes.push({ top, bottom: Math.min(height, top + rows) })
  return stripes
}

function sliceRows(buffer: PixelBuffer, top: number, bottom: number): PixelBuffer {
  const rowBytes = buffer.width * 4
  return { width: buffer.width, height: bottom - top, data: buffer.data.slice(top * rowBytes, bottom * rowBytes) }
}

function sliceMaskRows(mask: MaskBuffer, top: number, bottom: number): MaskBuffer {
  return { width: mask.width, height: bottom - top, data: mask.data.slice(top * mask.width, bottom * mask.width) }
}

// ---------------------------------------------------------------------------------------------
// Client
// ---------------------------------------------------------------------------------------------

interface Job {
  readonly id: number
  readonly op: ImagingOp
  input: unknown
  readonly transfer: boolean
  readonly signal?: AbortSignal
  readonly onProgress?: (fraction: number) => void
  readonly resolve: (output: unknown) => void
  readonly reject: (error: unknown) => void
  slot: Slot | null
  done: boolean
  detach: () => void
}

interface Slot {
  readonly worker: Worker
  ready: boolean
  /** Job running in this worker (kept after an abort until the worker answers). */
  job: Job | null
  timer: ReturnType<typeof setTimeout> | null
}

type AnyHandler = (input: unknown, options: OpOptions) => unknown

function defaultWorkerFactory(): (() => Worker) | null {
  if (typeof Worker === 'undefined') return null
  return () => new Worker(new URL('./imaging.worker.ts', import.meta.url), { type: 'module', name: 'simple-imaging' })
}

function defaultInlineLoader(): Promise<ImagingHandlers> {
  return import('./worker-ops/index.ts').then((module) => module.HANDLERS)
}

export function createImagingClient(options: ImagingClientOptions = {}): PooledImagingClient {
  const createWorker = options.createWorker ?? (options.inline ? null : defaultWorkerFactory())
  const loadInline = options.loadInline ?? defaultInlineLoader
  const stripeThreshold = Math.max(1, options.stripeThresholdPixels ?? DEFAULT_STRIPE_THRESHOLD)
  const startTimeout = Math.max(1, options.startTimeoutMs ?? 15_000)
  const margins = new Map<ImagingOp, (input: never) => number | null>()
  for (const [op, margin] of Object.entries(options.stripeMargins ?? {})) {
    if (typeof margin === 'function') margins.set(op as ImagingOp, margin as (input: never) => number | null)
  }

  let maxWorkers = Math.max(1, Math.floor(options.workers ?? defaultWorkerCount()))
  let state: 'workers' | 'inline' | 'disposed' = createWorker ? 'workers' : 'inline'
  let slots: Slot[] = []
  const queue: Job[] = []
  let nextId = 1
  let startFailures = 0
  let inlineHandlers: Promise<ImagingHandlers> | null = options.inline ? Promise.resolve(options.inline) : null

  const settle = (job: Job, outcome: () => void) => {
    if (job.done) return
    job.done = true
    job.detach()
    job.input = null
    outcome()
  }

  const inlineTable = () => {
    if (!inlineHandlers) {
      inlineHandlers = loadInline()
      // A failed load must not poison later attempts.
      inlineHandlers.catch(() => { inlineHandlers = null })
    }
    return inlineHandlers
  }

  const runInline = async (job: Job) => {
    try {
      const handlers = await inlineTable()
      if (job.done) return
      const handler = handlers[job.op] as AnyHandler | undefined
      if (typeof handler !== 'function') {
        throw namedError('NotSupportedError', `The imaging operation "${String(job.op)}" is not available in this build.`)
      }
      const output = await handler(job.input, { signal: job.signal, onProgress: job.onProgress })
      if (job.signal?.aborted) throw abortError()
      settle(job, () => job.resolve(output))
    } catch (error) {
      settle(job, () => job.reject(job.signal?.aborted ? abortError() : error))
    }
  }

  const terminate = (slot: Slot) => {
    if (slot.timer) clearTimeout(slot.timer)
    slot.timer = null
    try { slot.worker.terminate() } catch { /* already gone */ }
  }

  const fallBackToInline = () => {
    state = 'inline'
    for (const slot of slots) terminate(slot)
    slots = []
    pump()
  }

  const onSlotFailure = (slot: Slot) => {
    if (!slots.includes(slot)) return
    slots = slots.filter((entry) => entry !== slot)
    terminate(slot)
    const job = slot.job
    slot.job = null
    if (!slot.ready) {
      // Jobs are only posted to ready workers, so nothing was transferred to this one.
      startFailures += 1
      // Nothing else running or starting: workers cannot run here (CSP, missing chunk) -> inline.
      if (!slots.length) { fallBackToInline(); return }
      if (startFailures >= MAX_START_FAILURES) maxWorkers = Math.max(1, slots.length)
      pump()
      return
    }
    if (job) settle(job, () => job.reject(new Error('The image worker stopped unexpectedly. Try again; very large images may need more free memory.')))
    pump()
  }

  const onMessage = (slot: Slot, message: ImagingWorkerResponse | null | undefined) => {
    if (!message || typeof message !== 'object') return
    if (message.type === 'ready') {
      if (slot.timer) clearTimeout(slot.timer)
      slot.timer = null
      slot.ready = true
      pump()
      return
    }
    const job = slot.job
    if (!job || job.id !== message.id) return
    if (message.type === 'progress') {
      if (!job.done) job.onProgress?.(message.fraction)
      return
    }
    slot.job = null
    if (message.type === 'result') settle(job, () => job.resolve(message.output))
    else if (message.type === 'error') settle(job, () => job.reject(namedError(message.name || 'Error', message.message || 'The imaging operation failed.')))
    pump()
  }

  const spawn = (): boolean => {
    if (!createWorker) return false
    let worker: Worker
    try {
      worker = createWorker()
    } catch {
      startFailures += 1
      if (!slots.length) fallBackToInline()
      else maxWorkers = Math.max(1, slots.length)
      return false
    }
    const slot: Slot = { worker, ready: false, job: null, timer: null }
    const timer: unknown = setTimeout(() => onSlotFailure(slot), startTimeout)
    // Node (tests): do not keep the process alive for a worker that never starts.
    const unref = (timer as { unref?: () => void }).unref
    if (typeof unref === 'function') unref.call(timer)
    slot.timer = timer as ReturnType<typeof setTimeout>
    worker.addEventListener('message', (event: MessageEvent) => onMessage(slot, event.data as ImagingWorkerResponse))
    worker.addEventListener('error', (event: Event) => {
      event.preventDefault?.()
      onSlotFailure(slot)
    })
    worker.addEventListener('messageerror', () => onSlotFailure(slot))
    slots.push(slot)
    return true
  }

  const dispatch = (slot: Slot, job: Job) => {
    const buffers = new Set<ArrayBuffer>()
    try {
      const input = prepareTransfer(job.input, job.transfer, buffers)
      job.input = null
      slot.job = job
      job.slot = slot
      const request: ImagingWorkerRequest = { type: 'run', id: job.id, op: job.op, input }
      slot.worker.postMessage(request, [...buffers])
    } catch (error) {
      slot.job = null
      settle(job, () => job.reject(error))
    }
  }

  function pump(): void {
    if (state === 'disposed') return
    if (state === 'inline') {
      while (queue.length) void runInline(queue.shift() as Job)
      return
    }
    while (queue.length) {
      const slot = slots.find((entry) => entry.ready && !entry.job)
      if (!slot) break
      dispatch(slot, queue.shift() as Job)
    }
    const starting = slots.filter((entry) => !entry.ready).length
    let wanted = queue.length - starting
    while (wanted > 0 && slots.length < maxWorkers && state === 'workers') {
      if (!spawn()) return
      wanted -= 1
    }
  }

  const enqueue = (op: ImagingOp, input: unknown, runOptions: ImagingRunOptions): Promise<unknown> => {
    if (state === 'disposed') return Promise.reject(abortError('The imaging service was closed.'))
    const signal = runOptions.signal
    if (signal?.aborted) return Promise.reject(abortError())
    return new Promise((resolve, reject) => {
      const job: Job = {
        id: nextId++,
        op,
        input,
        transfer: runOptions.transfer !== false,
        signal,
        onProgress: runOptions.onProgress,
        resolve,
        reject,
        slot: null,
        done: false,
        detach: () => {},
      }
      if (signal) {
        const onAbort = () => {
          if (job.done) return
          const index = queue.indexOf(job)
          if (index >= 0) queue.splice(index, 1)
          const slot = job.slot
          if (slot && slot.job === job) {
            const request: ImagingWorkerRequest = { type: 'abort', id: job.id }
            try { slot.worker.postMessage(request) } catch { /* the worker is gone; failure handling frees the slot */ }
          }
          settle(job, () => job.reject(abortError()))
        }
        signal.addEventListener('abort', onAbort, { once: true })
        job.detach = () => signal.removeEventListener('abort', onAbort)
      }
      queue.push(job)
      pump()
    })
  }

  const runStriped = async (op: ImagingOp, input: StripeSource, stripes: Stripe[], margin: number, runOptions: ImagingRunOptions): Promise<PixelBuffer> => {
    const { src, mask } = input
    const width = src.width
    const output = new Uint8ClampedArray(width * src.height * 4)
    const fractions = new Float64Array(stripes.length)
    let reported = -1
    const report = () => {
      if (!runOptions.onProgress) return
      let done = 0
      for (let index = 0; index < stripes.length; index += 1) done += fractions[index] * (stripes[index].bottom - stripes[index].top)
      const fraction = Math.min(1, done / src.height)
      if (fraction < 1 && fraction - reported < 0.01) return
      if (fraction <= reported) return
      reported = fraction
      runOptions.onProgress(fraction)
    }
    const controller = new AbortController()
    const onAbort = () => controller.abort()
    runOptions.signal?.addEventListener('abort', onAbort, { once: true })
    try {
      await Promise.all(stripes.map(async (stripe, index) => {
        const top = Math.max(0, stripe.top - margin)
        const bottom = Math.min(src.height, stripe.bottom + margin)
        const part = { ...input, src: sliceRows(src, top, bottom), ...(mask ? { mask: sliceMaskRows(mask, top, bottom) } : {}) }
        const result = await enqueue(op, part, {
          signal: controller.signal,
          transfer: true,
          onProgress: (fraction) => { fractions[index] = fraction; report() },
        })
        if (!isPixelBuffer(result) || result.width !== width || result.height !== bottom - top) {
          throw new Error(`The imaging operation "${op}" returned a buffer of the wrong size.`)
        }
        const rowBytes = width * 4
        output.set(result.data.subarray((stripe.top - top) * rowBytes, (stripe.bottom - top) * rowBytes), stripe.top * rowBytes)
        fractions[index] = 1
        report()
      }))
    } catch (error) {
      controller.abort()
      throw runOptions.signal?.aborted ? abortError() : error
    } finally {
      runOptions.signal?.removeEventListener('abort', onAbort)
    }
    return { width, height: src.height, data: output }
  }

  const stripePlan = (op: ImagingOp, input: unknown): { stripes: Stripe[]; margin: number } | null => {
    if (state !== 'workers' || maxWorkers < 2) return null
    const marginOf = margins.get(op)
    const source = input as StripeSource | null
    if (!marginOf || !source || !isPixelBuffer(source.src)) return null
    const { width, height } = source.src
    if (width * height <= stripeThreshold) return null
    const mask = source.mask
    if (mask && (mask.width !== width || mask.height !== height || mask.data.length !== width * height)) return null
    let margin: number | null
    try { margin = marginOf(input as never) } catch { return null }
    if (margin === null || !Number.isFinite(margin) || margin < 0) return null
    const rows = Math.ceil(margin)
    const stripes = planStripes(width, height, rows, maxWorkers)
    return stripes ? { stripes, margin: rows } : null
  }

  const client: PooledImagingClient = {
    get mode() {
      return state
    },
    get pending() {
      return queue.length + slots.filter((slot) => slot.job && !slot.job.done).length
    },
    run<K extends ImagingOp>(op: K, input: ImagingInput<K>, runOptions: ImagingRunOptions = {}): Promise<ImagingOutput<K>> {
      if (state === 'disposed') return Promise.reject(abortError('The imaging service was closed.'))
      if (runOptions.signal?.aborted) return Promise.reject(abortError())
      const plan = stripePlan(op, input)
      if (plan) return runStriped(op, input as unknown as StripeSource, plan.stripes, plan.margin, runOptions) as Promise<ImagingOutput<K>>
      return enqueue(op, input, runOptions) as Promise<ImagingOutput<K>>
    },
    setStripeMargin(op, margin) {
      if (margin) margins.set(op, margin as (input: never) => number | null)
      else margins.delete(op)
    },
    trim() {
      if (state !== 'workers' || queue.length) return
      const idle = slots.filter((slot) => !slot.job)
      slots = slots.filter((slot) => slot.job)
      for (const slot of idle) terminate(slot)
    },
    dispose() {
      if (state === 'disposed') return
      state = 'disposed'
      const jobs = [...queue, ...slots.map((slot) => slot.job).filter((job): job is Job => Boolean(job))]
      queue.length = 0
      for (const slot of slots) terminate(slot)
      slots = []
      for (const job of jobs) settle(job, () => job.reject(abortError('The imaging service was closed.')))
    },
  }
  return client
}

let sharedClient: PooledImagingClient | null = null

/** The app-wide pool (created lazily; workers start on the first job). 'adjust' is striped by default. */
export function getImagingClient(): PooledImagingClient {
  if (!sharedClient || sharedClient.mode === 'disposed') {
    sharedClient = createImagingClient({ stripeMargins: { adjust: () => 0 } })
  }
  return sharedClient
}

/** Releases idle workers of the shared pool, if it exists (called when leaving Advanced mode). */
export function trimImagingWorkers(): void {
  sharedClient?.trim()
}
