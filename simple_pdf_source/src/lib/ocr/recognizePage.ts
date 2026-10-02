import type { PDFPageProxy } from 'pdfjs-dist'
import {
  needsOrientationCheck,
  pickOrientation,
  readingScore,
  tesseractReadingStats,
  tesseractToOcrPage,
  type OcrReading,
} from '../../../electron/ocr-geometry.mjs'
import { engineSignature } from '../../../electron/ocr-preprocess.mjs'
import { getOcrEngine, type OcrEngine, type OcrEnginePhase } from './engine'
import { normalizeRotation, renderForOcr, type OcrRenderOptions } from './renderForOcr'
import {
  OcrError,
  ocrAbortError,
  type OcrPageDiagnostics,
  type OcrPageProgress,
  type OcrPageResult,
  type OcrPreparedPage,
  type OcrPrepRequest,
  type OcrPrepResponse,
  type OcrRect,
  type OcrRotation,
  type TesseractPsm,
} from './types'

export interface RecognizePageOptions {
  /** Zero-based page index, recorded in the result. */
  pageIndex: number
  /** The user's pending view rotation for this page. */
  pageRotation?: number
  /** 'eng' (default) or several codes joined with '+'. */
  language?: string
  psm?: TesseractPsm
  /** Resolution of the page's scan image, when known (see chooseOcrDpi). */
  nativeDpi?: number
  /** Explicit render resolution; overrides nativeDpi. */
  dpi?: number
  /** Extra rotation for a page whose scan is sideways or upside down. */
  orientationCorrection?: OcrRotation
  /** Retry weak results in the other three orientations (see recognizePage). */
  orientationFallback?: boolean
  /** Mixed pages: visible native text rects, so OCR words over them are dropped. */
  nativeText?: OcrRect[]
  /** Mixed pages: keep only words inside these image rects. */
  imageRects?: OcrRect[]
  signal?: AbortSignal
  onProgress?: (progress: OcrPageProgress) => void
  /** Defaults to the shared engine for `language`. */
  engine?: OcrEngine
  /** Cache lookup by content key, consulted after preparation and before recognition; a hit is final. */
  lookup?: (contentKey: string) => OcrPageResult | undefined
  /**
   * Remember a final result under the key a later lookup for this page will
   * use: the first pass's key, also when the orientation was corrected (the
   * result's own contentKey then names the corrected render).
   */
  store?: (contentKey: string, result: OcrPageResult) => void
  intent?: OcrRenderOptions['intent']
  optionalContentConfigPromise?: OcrRenderOptions['optionalContentConfigPromise']
  cleanup?: boolean
  diagnostics?: OcrPageDiagnostics
}

const PREP_IDLE_MS = 60_000
let prepWorker: Worker | null = null
let prepIdleTimer: ReturnType<typeof setTimeout> | null = null
let prepSequence = 0
const pendingPrep = new Map<number, { resolve: (page: OcrPreparedPage) => void; reject: (error: unknown) => void }>()

function failPendingPrep(error: unknown) {
  const pending = [...pendingPrep.values()]
  pendingPrep.clear()
  for (const request of pending) request.reject(error)
}

/** Stop the preparation worker (it restarts on demand). */
export function disposeOcrPrepWorker() {
  if (prepIdleTimer) { clearTimeout(prepIdleTimer); prepIdleTimer = null }
  prepWorker?.terminate()
  prepWorker = null
  failPendingPrep(ocrAbortError())
}

function schedulePrepIdle() {
  if (prepIdleTimer) clearTimeout(prepIdleTimer)
  prepIdleTimer = setTimeout(() => {
    prepIdleTimer = null
    if (!pendingPrep.size) {
      prepWorker?.terminate()
      prepWorker = null
    }
  }, PREP_IDLE_MS)
}

function prepWorkerInstance(): Worker {
  if (prepWorker) return prepWorker
  const worker = new Worker(new URL('./ocrPrep.worker.ts', import.meta.url), { type: 'module', name: 'simple-ocr-prep' })
  worker.addEventListener('message', (event: MessageEvent<OcrPrepResponse>) => {
    const message = event.data
    const pending = pendingPrep.get(message.id)
    if (pending) {
      pendingPrep.delete(message.id)
      if (message.type === 'error') {
        pending.reject(new OcrError('prepare-failed', 'This page could not be prepared for text recognition.', message.message))
      } else pending.resolve(message)
    }
    // Also after a result nobody waits for (its page was stopped).
    if (!pendingPrep.size) schedulePrepIdle()
  })
  worker.addEventListener('error', (event) => {
    if (prepWorker === worker) prepWorker = null
    worker.terminate()
    failPendingPrep(new OcrError('prepare-failed', 'This page could not be prepared for text recognition.', event.message))
  })
  prepWorker = worker
  return worker
}

/** Run the preprocessing pipeline in the prep worker. `rgba` is transferred (detached). */
export function preparePageInWorker(input: { rgba: Uint8ClampedArray; width: number; height: number; dpi: number }, signal?: AbortSignal): Promise<OcrPreparedPage> {
  if (signal?.aborted) return Promise.reject(ocrAbortError())
  const id = ++prepSequence
  const buffer = input.rgba.byteOffset === 0 && input.rgba.byteLength === input.rgba.buffer.byteLength
    ? input.rgba.buffer as ArrayBuffer
    : input.rgba.slice().buffer as ArrayBuffer
  return new Promise<OcrPreparedPage>((resolve, reject) => {
    const onAbort = () => {
      // The worker finishes this page in well under a second; its result is ignored.
      pendingPrep.delete(id)
      reject(ocrAbortError())
    }
    signal?.addEventListener('abort', onAbort, { once: true })
    pendingPrep.set(id, {
      resolve: (page) => { signal?.removeEventListener('abort', onAbort); resolve(page) },
      reject: (error) => { signal?.removeEventListener('abort', onAbort); reject(error) },
    })
    if (prepIdleTimer) { clearTimeout(prepIdleTimer); prepIdleTimer = null }
    const request: OcrPrepRequest = { type: 'prepare-page', id, rgba: buffer, width: input.width, height: input.height, dpi: input.dpi }
    try {
      prepWorkerInstance().postMessage(request, [buffer])
    } catch (error) {
      pendingPrep.delete(id)
      signal?.removeEventListener('abort', onAbort)
      reject(new OcrError('prepare-failed', 'This page could not be prepared for text recognition.', String(error)))
    }
  })
}

const round6 = (value: number) => Math.round(value * 1e6) / 1e6

const now = () => performance.now()

/** Orientation probes only need to tell readable from unreadable, so they run at a low resolution. */
const ORIENTATION_PROBE_DPI = 150

interface PagePass {
  result: OcrPageResult
  blank: boolean
  /** Served by `lookup`: final, no orientation check. */
  cached: boolean
  contentKey: string
  reading: OcrReading
}

/**
 * Recognise one page: render (main thread) -> prepare (worker) -> Tesseract
 * (engine pool) -> PDF geometry. Blank pages skip recognition and return a
 * result without words.
 *
 * With `orientationFallback`, a weak or mostly vertical first pass (see
 * needsOrientationCheck in ocr-geometry) is probed at the other three
 * orientations at 150 DPI; the probe pickOrientation() selects is recognised
 * again at full resolution and kept when it scores higher. Its
 * orientationCorrectedBy then names the correction (a "page looks sideways"
 * hint).
 */
export async function recognizePage(page: PDFPageProxy, options: RecognizePageOptions): Promise<OcrPageResult> {
  const first = await recognizeOnce(page, options)
  if (first.cached) return first.result
  const result = options.orientationFallback && !first.blank && needsOrientationCheck(first.reading)
    ? await withOrientationFallback(page, options, first)
    : first.result
  options.store?.(first.contentKey, result)
  return result
}

async function withOrientationFallback(page: PDFPageProxy, options: RecognizePageOptions, first: PagePass): Promise<OcrPageResult> {
  const base = normalizeRotation(options.orientationCorrection ?? 0)
  // The probes share the engine's worker pool, so they overlap.
  let finished = 0
  options.onProgress?.({ phase: 'orientation', progress: 0 })
  const probes = await Promise.all([90, 180, 270].map(async (extra) => {
    const correction = normalizeRotation(base + extra)
    const probe = await recognizeOnce(page, {
      ...options,
      orientationCorrection: correction,
      dpi: ORIENTATION_PROBE_DPI,
      lookup: undefined,
      store: undefined,
      diagnostics: undefined,
      onProgress: undefined,
      cleanup: false,
    })
    finished += 1
    options.onProgress?.({ phase: 'orientation', progress: finished / 3 })
    return { correction, reading: probe.reading }
  }))
  const best = pickOrientation(first.reading, probes)
  if (options.diagnostics) {
    options.diagnostics.orientation = {
      probes: probes.map(({ correction, reading }) => ({ correction, tesseractConfidence: reading.tesseractConfidence, wordCount: reading.wordCount, verticalLines: reading.verticalLines })),
      chosen: base,
    }
  }
  if (!best) {
    if (options.cleanup !== false) page.cleanup()
    return first.result
  }
  const corrected = await recognizeOnce(page, { ...options, orientationCorrection: best.correction, lookup: undefined, store: undefined })
  if (readingScore(corrected.reading) < readingScore(first.reading)) return first.result
  if (options.diagnostics?.orientation) options.diagnostics.orientation.chosen = best.correction
  return corrected.result
}

async function recognizeOnce(page: PDFPageProxy, options: RecognizePageOptions): Promise<PagePass> {
  const { signal } = options
  const diagnostics = options.diagnostics
  const language = options.language ?? 'eng'
  const psm = options.psm ?? '3'
  const correction = normalizeRotation(options.orientationCorrection ?? 0)
  const report = (phase: OcrPageProgress['phase'], progress: number) => options.onProgress?.({ phase, progress: Math.max(0, Math.min(1, progress)) })
  const started = now()
  if (signal?.aborted) throw ocrAbortError()

  report('render', 0)
  let lap = now()
  const render = await renderForOcr(page, {
    pageRotation: (options.pageRotation ?? 0) + correction,
    nativeDpi: options.nativeDpi,
    dpi: options.dpi,
    intent: options.intent,
    optionalContentConfigPromise: options.optionalContentConfigPromise,
    cleanup: options.cleanup,
    signal,
  })
  if (diagnostics) Object.assign(diagnostics, { renderMs: Math.round(now() - lap), width: render.width, height: render.height, dpi: render.dpi, intent: render.intent })

  report('prepare', 0.08)
  lap = now()
  const prepared = await preparePageInWorker({ rgba: render.rgba, width: render.width, height: render.height, dpi: render.dpi }, signal)
  if (diagnostics) Object.assign(diagnostics, { prepareMs: Math.round(now() - lap), prep: prepared.stats, blank: prepared.blank })

  // Same pixels, engine settings and pixel-to-PDF mapping give the same result.
  const contentKey = `${prepared.imageSha256}|${engineSignature({ language, psm })}|${render.pdfFromPixel.map(round6).join(',')}`
  const cached = options.lookup?.(contentKey)
  if (cached) {
    if (diagnostics) Object.assign(diagnostics, { cacheHit: true, totalMs: Math.round(now() - started) })
    report('map', 1)
    const result = cached.pageIndex === options.pageIndex ? cached : { ...cached, pageIndex: options.pageIndex }
    const lines = result.paragraphs.reduce((sum, paragraph) => sum + paragraph.lines.length, 0)
    const reading = { tesseractConfidence: result.meanConfidence, meanConfidence: result.meanConfidence, wordCount: result.wordCount, lines, verticalLines: 0 }
    return { result, blank: prepared.blank, cached: true, contentKey, reading }
  }

  let blocks: Parameters<typeof tesseractToOcrPage>[0] = []
  let tesseractConfidence = 0
  if (!prepared.blank && prepared.pgm) {
    report('engine', 0.15)
    lap = now()
    const engine = options.engine ?? getOcrEngine({ language })
    const recognized = await engine.recognize({
      pgm: prepared.pgm,
      dpi: render.dpi,
      psm,
      signal,
      onProgress: (progress: number, phase: OcrEnginePhase) => report(phase === 'recognizing' ? 'recognize' : 'engine', 0.2 + 0.75 * progress),
    })
    blocks = recognized.blocks ?? []
    tesseractConfidence = Number.isFinite(recognized.confidence) ? recognized.confidence : 0
    if (diagnostics) diagnostics.recognizeMs = Math.round(now() - lap)
  }

  report('map', 0.97)
  lap = now()
  const result = tesseractToOcrPage(blocks, {
    toPdf: render.toPdf,
    dpi: render.dpi,
    pageIndex: options.pageIndex,
    contentKey,
    rotation: render.rotation,
    deskew: prepared.deskew,
    language,
    orientationCorrectedBy: correction,
    nativeText: options.nativeText,
    imageRects: options.imageRects,
  })
  if (diagnostics) Object.assign(diagnostics, { mapMs: Math.round(now() - lap), cacheHit: false, totalMs: Math.round(now() - started) })
  report('map', 1)
  const stats = tesseractReadingStats(blocks)
  const reading = { tesseractConfidence, meanConfidence: result.meanConfidence, wordCount: result.wordCount, lines: stats.lines, verticalLines: stats.verticalLines }
  return { result, blank: prepared.blank, cached: false, contentKey, reading }
}
