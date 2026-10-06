import Tesseract from 'tesseract.js/dist/tesseract.esm.min.js'
import type { ImageLike, LoggerMessage, Page as TesseractPage, Worker as TesseractWorker, WorkerParams } from 'tesseract.js'
import { tesseractParameters } from '../../../electron/ocr-preprocess.mjs'
import { ocrAssetLocation, ocrLanguageCodes, preflightOcrAssets } from './assets'
import { DEFAULT_OCR_WORKERS, OcrError, ocrAbortError, type TesseractPsm } from './types'

export type { TesseractPage }
export type OcrEnginePhase = 'loading-core' | 'loading-language' | 'initializing' | 'recognizing'

export interface OcrEngineOptions {
  /** 'eng', or several joined with '+', e.g. 'eng+deu'. */
  language: string
  /** Upper bound on parallel workers (default DEFAULT_OCR_WORKERS); also capped by the CPU count. */
  maxWorkers?: number
  /** Diagnostics and tests only: another ocr/ asset folder (default ./ocr/). */
  assetBase?: string
  /** Tests only: how long a worker may take to start (default 30 s); applies when the engine is created. */
  initTimeoutMs?: number
}

export interface RecognizeRequest {
  /** A P5 PGM as produced by the prep worker. */
  pgm: Uint8Array
  /** Resolution of the PGM; Tesseract assumes 70 DPI otherwise. */
  dpi: number
  psm?: TesseractPsm
  signal?: AbortSignal
  onProgress?: (progress: number, phase: OcrEnginePhase) => void
}

export interface OcrEngine {
  readonly language: string
  /** Resolves with Tesseract's page: `blocks` (block > paragraph > line > word > symbol) and `confidence`. */
  recognize(request: RecognizeRequest): Promise<TesseractPage>
  dispose(): Promise<void>
}

/** createWorker must settle within this time; tesseract.js 7 can otherwise wait forever. */
const INIT_TIMEOUT_MS = 30_000
/** A page that takes longer is treated as a hung worker. */
const JOB_TIMEOUT_MS = 180_000
const IDLE_TIMEOUT_MS = 60_000

const PHASES: Record<string, OcrEnginePhase> = {
  'loading tesseract core': 'loading-core',
  'initializing tesseract': 'loading-core',
  'loading language traineddata': 'loading-language',
  'initializing api': 'initializing',
  'recognizing text': 'recognizing',
}

interface Job {
  request: RecognizeRequest
  resolve: (page: TesseractPage) => void
  reject: (error: unknown) => void
  settled: boolean
  slot: Slot | null
  timer: ReturnType<typeof setTimeout> | null
  detach: () => void
}

interface Slot {
  worker: TesseractWorker | null
  /** The Web Worker under `worker`, kept so a worker whose start never settles can still be terminated. */
  native: Worker | null
  starting: Promise<TesseractWorker> | null
  cancelStart: ((error: unknown) => void) | null
  job: Job | null
  idleTimer: ReturnType<typeof setTimeout> | null
  generation: number
}

/**
 * Run `spawn` while recording the Web Workers it constructs. tesseract.js
 * creates its worker synchronously inside createWorker() but only exposes it
 * once initialisation succeeds; on a failed language load the promise never
 * settles, and without this handle the worker (and its wasm heap) would leak.
 */
function captureWorkers<T>(spawn: () => T): { result: T; native: Worker | null } {
  const NativeWorker = globalThis.Worker
  const created: Worker[] = []
  const Capturing = function CapturingWorker(url: string | URL, options?: WorkerOptions) {
    const instance = new NativeWorker(url, options)
    created.push(instance)
    return instance
  } as unknown as typeof Worker
  Capturing.prototype = NativeWorker.prototype
  globalThis.Worker = Capturing
  try {
    return { result: spawn(), native: created[0] ?? null }
  } finally {
    globalThis.Worker = NativeWorker
  }
}

function engineError(code: 'engine-init' | 'recognize-failed' | 'engine-crashed' | 'engine-timeout', detail: unknown) {
  const text = detail instanceof Error ? detail.message : String(detail ?? '')
  const message = code === 'engine-init'
    ? 'The text recognition engine could not start.'
    : code === 'engine-timeout'
      ? 'The text recognition engine did not start in time.'
      : code === 'engine-crashed'
        ? 'The text recognition engine stopped responding.'
        : 'Text recognition failed on this page.'
  return new OcrError(code, message, text)
}

class TesseractEngine implements OcrEngine {
  readonly language: string
  private readonly assetBase?: string
  private readonly key: string
  private readonly initTimeoutMs: number
  private maxWorkers: number
  private slots: Slot[] = []
  private queue: Job[] = []
  private disposed = false

  constructor(key: string, language: string, maxWorkers: number, assetBase?: string, initTimeoutMs = INIT_TIMEOUT_MS) {
    this.key = key
    this.language = language
    this.maxWorkers = maxWorkers
    this.assetBase = assetBase
    this.initTimeoutMs = initTimeoutMs
  }

  setMaxWorkers(maxWorkers: number) {
    this.maxWorkers = maxWorkers
    this.pump()
  }

  recognize(request: RecognizeRequest): Promise<TesseractPage> {
    if (this.disposed) return Promise.reject(ocrAbortError())
    if (request.signal?.aborted) return Promise.reject(ocrAbortError())
    if (!(request.pgm instanceof Uint8Array) || request.pgm.length < 16) {
      return Promise.reject(new OcrError('recognize-failed', 'Text recognition failed on this page.', 'Empty OCR image'))
    }
    try {
      tesseractParameters({ dpi: request.dpi, psm: request.psm ?? '3' })
    } catch (error) {
      return Promise.reject(new OcrError('recognize-failed', 'Text recognition failed on this page.', String(error)))
    }
    return new Promise<TesseractPage>((resolve, reject) => {
      const job: Job = { request, resolve, reject, settled: false, slot: null, timer: null, detach: () => {} }
      const signal = request.signal
      if (signal) {
        const onAbort = () => this.abort(job)
        signal.addEventListener('abort', onAbort, { once: true })
        job.detach = () => signal.removeEventListener('abort', onAbort)
      }
      this.queue.push(job)
      this.pump()
    })
  }

  async dispose() {
    if (this.disposed) return
    this.disposed = true
    if (registry.get(this.key) === this) registry.delete(this.key)
    for (const job of this.queue.splice(0)) this.settle(job, ocrAbortError())
    for (const slot of this.slots) {
      const job = slot.job
      slot.job = null
      this.stopSlot(slot, ocrAbortError())
      if (job) this.settle(job, ocrAbortError())
    }
    this.slots = []
  }

  private poolSize() {
    const cores = typeof navigator !== 'undefined' && navigator.hardwareConcurrency ? navigator.hardwareConcurrency : 2
    return Math.max(1, Math.min(this.maxWorkers, cores - 2))
  }

  private pump() {
    if (this.disposed) return
    while (this.queue.length) {
      let slot = this.slots.find((candidate) => !candidate.job && candidate.worker)
        ?? this.slots.find((candidate) => !candidate.job)
      if (!slot && this.slots.length < this.poolSize()) {
        slot = { worker: null, native: null, starting: null, cancelStart: null, job: null, idleTimer: null, generation: 0 }
        this.slots.push(slot)
      }
      if (!slot) return
      const job = this.queue.shift()!
      if (slot.idleTimer) { clearTimeout(slot.idleTimer); slot.idleTimer = null }
      slot.job = job
      job.slot = slot
      void this.run(slot, job)
    }
  }

  private async run(slot: Slot, job: Job) {
    const generation = slot.generation
    const current = () => !job.settled && slot.generation === generation && slot.job === job
    try {
      const worker = await this.ensureWorker(slot, job)
      if (!current()) return
      job.timer = setTimeout(() => {
        if (slot.job !== job) return
        slot.job = null
        this.stopSlot(slot, ocrAbortError())
        this.settle(job, engineError('engine-crashed', `No result after ${JOB_TIMEOUT_MS / 1000} s`))
        this.pump()
      }, JOB_TIMEOUT_MS)
      // The typings declare PSM as an enum and omit typed arrays, but the PSM
      // enum values are these strings and loadImage() copies any Uint8Array.
      const parameters = tesseractParameters({ dpi: job.request.dpi, psm: job.request.psm ?? '3' }) as unknown as Partial<WorkerParams>
      await worker.setParameters(parameters)
      if (!current()) return
      job.request.onProgress?.(0, 'recognizing')
      const { data } = await worker.recognize(job.request.pgm as unknown as ImageLike, {}, { text: false, blocks: true })
      if (!current()) return
      this.settle(job, null, data)
    } catch (error) {
      // settle() ignores jobs that were already aborted or timed out; a failed
      // start has already stopped the slot (new generation) and must still settle.
      if (error instanceof OcrError || (error as { name?: string })?.name === 'AbortError') this.settle(job, error)
      else this.settle(job, engineError('recognize-failed', error))
    } finally {
      if (slot.job === job) {
        slot.job = null
        this.scheduleIdle(slot)
      }
      this.pump()
    }
  }

  private settle(job: Job, error: unknown, page?: TesseractPage) {
    if (job.settled) return
    job.settled = true
    if (job.timer) clearTimeout(job.timer)
    job.detach()
    if (error) job.reject(error)
    else job.resolve(page!)
  }

  private abort(job: Job) {
    if (job.settled) return
    const queued = this.queue.indexOf(job)
    if (queued >= 0) this.queue.splice(queued, 1)
    const slot = job.slot
    if (slot && slot.job === job) {
      // tesseract.js cannot cancel a running job: end the worker; the slot restarts lazily.
      slot.job = null
      this.stopSlot(slot, ocrAbortError())
    }
    this.settle(job, ocrAbortError())
    this.pump()
  }

  private scheduleIdle(slot: Slot) {
    if (slot.idleTimer) clearTimeout(slot.idleTimer)
    if (!slot.worker) return
    slot.idleTimer = setTimeout(() => {
      slot.idleTimer = null
      if (!slot.job) this.stopSlot(slot, ocrAbortError())
    }, IDLE_TIMEOUT_MS)
  }

  private stopSlot(slot: Slot, reason: unknown) {
    slot.generation += 1
    if (slot.idleTimer) { clearTimeout(slot.idleTimer); slot.idleTimer = null }
    const { worker, native, cancelStart } = slot
    slot.worker = null
    slot.native = null
    slot.starting = null
    slot.cancelStart = null
    cancelStart?.(reason)
    try { void worker?.terminate() } catch { /* already gone */ }
    try { native?.terminate() } catch { /* already gone */ }
  }

  private ensureWorker(slot: Slot, job: Job): Promise<TesseractWorker> {
    if (slot.worker) return Promise.resolve(slot.worker)
    if (!slot.starting) {
      const generation = slot.generation
      slot.starting = this.startWorker(slot).then((worker) => {
        if (slot.generation !== generation) {
          void worker.terminate()
          throw ocrAbortError()
        }
        slot.worker = worker
        slot.starting = null
        slot.cancelStart = null
        return worker
      }, (error) => {
        if (slot.generation === generation) this.stopSlot(slot, error)
        throw error
      })
    }
    job.request.onProgress?.(0, 'loading-core')
    return slot.starting
  }

  private async startWorker(slot: Slot): Promise<TesseractWorker> {
    const generation = slot.generation
    await preflightOcrAssets({ base: this.assetBase, languages: ocrLanguageCodes(this.language) })
    if (slot.generation !== generation) throw ocrAbortError()
    const location = ocrAssetLocation(this.assetBase)
    return new Promise<TesseractWorker>((resolve, reject) => {
      let settled = false
      const finish = (error: unknown, worker?: TesseractWorker) => {
        if (settled) {
          if (worker) void worker.terminate()
          return
        }
        settled = true
        clearTimeout(timer)
        if (error) {
          if (worker) void worker.terminate()
          reject(error)
        } else resolve(worker!)
      }
      const timer = setTimeout(() => finish(engineError('engine-timeout', `createWorker did not settle within ${this.initTimeoutMs} ms`)), this.initTimeoutMs)
      slot.cancelStart = (reason) => finish(reason)
      let captured: { result: Promise<TesseractWorker>; native: Worker | null }
      try {
        captured = captureWorkers(() => Tesseract.createWorker(this.language, Tesseract.OEM.LSTM_ONLY, {
          workerPath: location.workerPath,
          corePath: location.corePath,
          // Always a langPath: tesseract.js 7 mishandles {code, data} language objects.
          langPath: location.langPath,
          workerBlobURL: false,
          cacheMethod: 'none',
          gzip: true,
          logger: (message: LoggerMessage) => this.log(slot, message),
          // Called for a failed load/language/initialise step (createWorker then never
          // settles) and for every later job rejection (the job promise rejects too).
          errorHandler: (error: unknown) => {
            if (!settled) finish(engineError('engine-init', error))
          },
        }))
      } catch (error) {
        finish(engineError('engine-init', error))
        return
      }
      slot.native = captured.native
      if (!captured.native) console.warn('OCR: the recognition worker could not be tracked; a failed start may leave it running.')
      captured.native?.addEventListener('error', (event) => {
        const message = event instanceof ErrorEvent ? event.message : 'Worker error'
        if (!settled) {
          finish(engineError('engine-init', message))
          return
        }
        if (slot.native !== captured.native) return
        const job = slot.job
        slot.job = null
        this.stopSlot(slot, ocrAbortError())
        if (job) this.settle(job, engineError('engine-crashed', message))
        this.pump()
      })
      captured.result.then((worker) => finish(null, worker), (error) => finish(engineError('engine-init', error)))
    })
  }

  private log(slot: Slot, message: LoggerMessage) {
    const job = slot.job
    if (!job?.request.onProgress) return
    const phase = PHASES[message.status]
    if (!phase) return
    const progress = Number.isFinite(message.progress) ? Math.max(0, Math.min(1, message.progress)) : 0
    job.request.onProgress(phase === 'recognizing' ? progress : 0, phase)
  }
}

const registry = new Map<string, TesseractEngine>()
let unloadHooked = false

/** One engine (worker pool) per language and asset folder. */
export function getOcrEngine(options: OcrEngineOptions): OcrEngine {
  const language = ocrLanguageCodes(options.language).join('+')
  const maxWorkers = Math.max(1, Math.floor(options.maxWorkers ?? DEFAULT_OCR_WORKERS))
  const key = `${language}|${options.assetBase ?? ''}`
  let engine = registry.get(key)
  if (engine) {
    engine.setMaxWorkers(maxWorkers)
    return engine
  }
  const initTimeoutMs = Number.isFinite(options.initTimeoutMs) && options.initTimeoutMs! > 0 ? options.initTimeoutMs : undefined
  engine = new TesseractEngine(key, language, maxWorkers, options.assetBase, initTimeoutMs)
  registry.set(key, engine)
  if (!unloadHooked && typeof window !== 'undefined') {
    unloadHooked = true
    // pagehide, unlike beforeunload, only fires when the page really goes away.
    window.addEventListener('pagehide', () => { void disposeAllOcrEngines() })
  }
  return engine
}

/** Stop every OCR worker, e.g. when another document is opened. */
export async function disposeAllOcrEngines(): Promise<void> {
  await Promise.all([...registry.values()].map((engine) => engine.dispose()))
}
