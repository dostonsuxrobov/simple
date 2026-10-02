import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfRect } from '../../types'
import { getCachedOcrResult, storeOcrResult } from './ocrCache'
import { classifyPages, holdsDecodedObjects, PAGE_CLASSIFIER, type PageScanState } from './pageClassifier'
import { isOcrAbort, ocrAbortError, OcrError, type OcrLayerOperation, type OcrPageResult } from './types'

// "Recognize text": decide which pages to read and how, read them through the
// engine's worker pool (rendering, preparation and recognition of different
// pages overlap), report progress, and on Stop keep every page finished so far.

// The recognition runtime (tesseract.js, the engine pool, preprocessing and
// geometry) loads on first use, so opening the app does not pay for it.
interface OcrRuntime {
  recognizePage: typeof import('./recognizePage').recognizePage
  disposeOcrPrepWorker: typeof import('./recognizePage').disposeOcrPrepWorker
  getOcrEngine: typeof import('./engine').getOcrEngine
  disposeAllOcrEngines: typeof import('./engine').disposeAllOcrEngines
  buildOcrLayerOperation: typeof import('../../../electron/ocr-geometry.mjs').buildOcrLayerOperation
}

let runtime: Promise<OcrRuntime> | null = null

function loadOcrRuntime(): Promise<OcrRuntime> {
  if (!runtime) {
    const loading = Promise.all([import('./recognizePage'), import('./engine'), import('../../../electron/ocr-geometry.mjs')])
      .then(([page, engine, geometry]) => ({
        recognizePage: page.recognizePage,
        disposeOcrPrepWorker: page.disposeOcrPrepWorker,
        getOcrEngine: engine.getOcrEngine,
        disposeAllOcrEngines: engine.disposeAllOcrEngines,
        buildOcrLayerOperation: geometry.buildOcrLayerOperation,
      }))
    runtime = loading
    loading.catch(() => { if (runtime === loading) runtime = null })
  }
  return runtime
}

/** Stop every recognition worker (another document opened). Nothing to do when OCR never ran. */
export async function disposeOcrRuntime() {
  const loaded = runtime ? await runtime.catch(() => null) : null
  if (!loaded) return
  loaded.disposeOcrPrepWorker()
  await loaded.disposeAllOcrEngines()
}

// An 8 x 8 white page: reading it starts a recognition worker and returns at once.
const WARM_UP_PGM = (() => {
  const header = new TextEncoder().encode('P5\n8 8\n255\n')
  const image = new Uint8Array(header.length + 64).fill(255)
  image.set(header)
  return image
})()
let warmingUp: Promise<void> | null = null

/**
 * Start the recognition engine ahead of a likely run (the offer or the dialog
 * is on screen): the first page then does not wait for the engine to load.
 * An idle engine stops by itself after a minute. Failures are left for the
 * run to report.
 */
export function warmUpOcr(language = 'eng') {
  warmingUp ??= loadOcrRuntime()
    .then((ocr) => ocr.getOcrEngine({ language }).recognize({ pgm: WARM_UP_PGM.slice(), dpi: 300, psm: '6' }))
    .then(() => {}, () => {})
    .finally(() => { warmingUp = null })
  return warmingUp
}

/** 'needed': every page without text; 'current' and 'selected' were chosen by the user. */
export type OcrScope = 'needed' | 'current' | 'selected'

export type OcrSkipReason = 'has-text' | 'has-ocr-text' | 'blank'

export interface OcrPagePlan {
  pageIndex: number
  action: 'recognize' | 'skip'
  reason?: OcrSkipReason
  /** Mixed pages: visible text that recognised words must not duplicate. */
  nativeText?: PdfRect[]
  /** Mixed pages: only words on these images are kept. */
  imageRects?: PdfRect[]
  nativeDpi?: number
}

/**
 * What to do with one page. `explicit` is a page the user picked (This page,
 * Selected pages); "pages that need it" reads only pages without any text.
 */
export function planOcrPage(pageIndex: number, state: PageScanState, options: { explicit: boolean; replaceExisting: boolean }): OcrPagePlan {
  const skip = (reason: OcrSkipReason): OcrPagePlan => ({ pageIndex, action: 'skip', reason })
  const recognize = (extra: Partial<OcrPagePlan> = {}): OcrPagePlan => ({
    pageIndex,
    action: 'recognize',
    ...(state.nativeDpi ? { nativeDpi: state.nativeDpi } : {}),
    ...extra,
  })
  if (state.kind === 'blank') return skip('blank')
  // Recognised text is only ever replaced, never doubled.
  if (state.hasOcrText && !options.replaceExisting) return skip('has-ocr-text')
  // A scan's few visible words (a stamp, a Bates number) are not read a second time.
  if (state.kind === 'scan' || state.kind === 'searchable-scan' || state.kind === 'vector-only') {
    return recognize(state.visibleTextRects.length ? { nativeText: state.visibleTextRects } : {})
  }
  if (!options.explicit) return skip('has-text')
  // Visible text stays as it is; only pictures on the page are read.
  if (!state.imageRects.length) return skip('has-text')
  return recognize({ nativeText: state.visibleTextRects, imageRects: state.imageRects })
}

/** Pages "Recognize text" would read for the scope "pages that need it". */
export function pageNeedsRecognition(state: PageScanState, replaceExisting: boolean) {
  return planOcrPage(0, state, { explicit: false, replaceExisting }).action === 'recognize'
}

/** Pages with recognised text among these states. */
export function hasRecognizedText(states: Iterable<PageScanState>) {
  for (const state of states) if (state.invisibleChars >= PAGE_CLASSIFIER.minChars) return true
  return false
}

export const OCR_SECONDS_PER_PAGE = 1.8

/** Parallel recognition workers: the engine pool's size. */
export function ocrWorkerCount(maxWorkers = 2) {
  const cores = typeof navigator !== 'undefined' && navigator.hardwareConcurrency ? navigator.hardwareConcurrency : 2
  return Math.max(1, Math.min(maxWorkers, cores - 2))
}

export function estimateOcrSeconds(pages: number, workers = ocrWorkerCount()) {
  return pages <= 0 ? 0 : (pages * OCR_SECONDS_PER_PAGE) / Math.max(1, workers)
}

/** "About 40 seconds", "About 3 minutes". */
export function formatOcrDuration(seconds: number) {
  if (!Number.isFinite(seconds) || seconds <= 0) return ''
  if (seconds < 8) return 'A few seconds'
  if (seconds < 55) return `About ${Math.max(10, Math.round(seconds / 5) * 5)} seconds`
  const minutes = Math.max(1, Math.round(seconds / 60))
  return `About ${minutes} ${minutes === 1 ? 'minute' : 'minutes'}`
}

export interface OcrRunProgress {
  phase: 'checking' | 'recognizing'
  /** Pages finished (recognised or failed). */
  done: number
  /** Pages to recognise (pages checked, while checking). */
  total: number
  /** Overall share, 0..1. */
  fraction: number
  remainingSeconds?: number
  /** The recognition engine is still starting (first use in this window). */
  starting?: boolean
}

export interface OcrRunOptions {
  /** Candidate pages, already resolved from the scope. */
  pages: readonly number[]
  explicit: boolean
  language: string
  replaceExisting: boolean
  /** The user's pending view rotation per page: pages are read as they are shown. */
  pageRotations?: Record<number, number>
  signal: AbortSignal
  onProgress?: (progress: OcrRunProgress) => void
  /** See ClassifyPagesOptions.canRelease. */
  canReleasePage?: (pageIndex: number) => boolean
}

export interface OcrRunOutcome {
  /** Finished pages in page order, including pages where no text was found. */
  results: OcrPageResult[]
  /** The cache key each finished page was stored under (for the new document's index). */
  contentKeys: Map<number, string>
  skipped: Array<{ pageIndex: number; reason: OcrSkipReason }>
  failed: Array<{ pageIndex: number; message: string }>
  /** Pages that were to be recognised. */
  planned: number
  stopped: boolean
  /** The engine could not run at all (missing or damaged components); the run ended early. */
  fatal?: Error
  /** The document change for the finished pages, or null when there is nothing to write. */
  operation: OcrLayerOperation | null
}

/**
 * The part of a cache key that names a mixed page's filters: the same pixels
 * read with other visible-text or picture regions keep other words.
 */
export function regionSignature(plan: Pick<OcrPagePlan, 'nativeText' | 'imageRects'>) {
  if (!plan.nativeText?.length && !plan.imageRects?.length) return ''
  const round = (rect: PdfRect) => [rect.x, rect.y, rect.width, rect.height].map((value) => Math.round(value * 10)).join(',')
  const text = `${(plan.nativeText ?? []).map(round).join(';')}/${(plan.imageRects ?? []).map(round).join(';')}`
  // FNV-1a, 32 bit: short and stable.
  let hash = 0x811c9dc5
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index)
    hash = Math.imul(hash, 0x01000193) >>> 0
  }
  return `|regions:${hash.toString(16).padStart(8, '0')}`
}

const FATAL_CODES = new Set(['assets-missing', 'assets-invalid', 'language-unavailable', 'engine-init', 'engine-timeout'])

function isFatal(error: unknown) {
  return error instanceof OcrError && FATAL_CODES.has(error.code)
}

function failureMessage(error: unknown) {
  if (error instanceof OcrError) return error.message
  return error instanceof Error && error.message ? error.message : 'Text recognition failed on this page.'
}

/**
 * Recognise the pages the options select. Resolves with what was finished,
 * also when stopped (`stopped`) or when the engine cannot start (`fatal`); it
 * rejects only for a document that cannot be read at all.
 */
export async function runOcr(pdf: PDFDocumentProxy, options: OcrRunOptions): Promise<OcrRunOutcome> {
  const { signal } = options
  const outcome: OcrRunOutcome = { results: [], contentKeys: new Map(), skipped: [], failed: [], planned: 0, stopped: false, operation: null }
  // Load the runtime while the pages are checked.
  const loadingRuntime = loadOcrRuntime()
  loadingRuntime.catch(() => {})
  const report = options.onProgress ?? (() => {})
  const candidates = [...new Set(options.pages)].filter((index) => Number.isInteger(index) && index >= 0 && index < pdf.numPages).sort((a, b) => a - b)

  let states: Map<number, PageScanState>
  // Pages whose decoded images someone else holds (the viewer) keep them after being read.
  const held = new Set<number>()
  try {
    report({ phase: 'checking', done: 0, total: candidates.length, fraction: 0 })
    for (const pageIndex of candidates) {
      if (holdsDecodedObjects(await pdf.getPage(pageIndex + 1))) held.add(pageIndex)
    }
    states = await classifyPages(pdf, candidates, {
      signal,
      canRelease: options.canReleasePage,
      onProgress: (done, total) => report({ phase: 'checking', done, total, fraction: 0 }),
    })
  } catch (error) {
    if (isOcrAbort(error) || signal.aborted) {
      outcome.stopped = true
      return outcome
    }
    throw error
  }

  const plans = candidates.map((pageIndex) => planOcrPage(pageIndex, states.get(pageIndex)!, {
    explicit: options.explicit,
    replaceExisting: options.replaceExisting,
  }))
  for (const plan of plans) if (plan.action === 'skip') outcome.skipped.push({ pageIndex: plan.pageIndex, reason: plan.reason! })
  const queue = plans.filter((plan) => plan.action === 'recognize')
  outcome.planned = queue.length
  if (!queue.length) return outcome

  // A private controller: Stop (the caller's signal) and a fatal engine error both end every page in flight.
  const controller = new AbortController()
  const stop = () => controller.abort()
  if (signal.aborted) stop()
  signal.addEventListener('abort', stop, { once: true })

  let ocr: OcrRuntime
  try {
    ocr = await loadingRuntime
  } catch (error) {
    outcome.fatal = new OcrError('assets-missing', 'Text recognition (OCR) components are missing from this build.', String(error))
    return outcome
  }

  const total = queue.length
  const active = new Map<number, number>()
  const started = performance.now()
  let done = 0
  let engineRunning = false
  const publish = () => {
    let partial = 0
    for (const share of active.values()) partial += share
    const completed = Math.min(total, done + partial)
    const fraction = total ? completed / total : 1
    // Measured speed once a fair part of a page is done; the typical speed before.
    const remainingSeconds = completed > 0.25
      ? Math.max(0, ((performance.now() - started) / 1000 / completed) * (total - completed))
      : estimateOcrSeconds(total - completed)
    report({ phase: 'recognizing', done, total, fraction, remainingSeconds, starting: !engineRunning && done === 0 })
  }
  publish()

  const results = new Map<number, OcrPageResult>()
  let cursor = 0
  const worker = async () => {
    while (cursor < queue.length && !controller.signal.aborted) {
      const plan = queue[cursor]
      cursor += 1
      active.set(plan.pageIndex, 0)
      let finished = true
      try {
        const page = await pdf.getPage(plan.pageIndex + 1)
        const regions = regionSignature(plan)
        let storedKey: string | undefined
        const result = await ocr.recognizePage(page, {
          pageIndex: plan.pageIndex,
          pageRotation: options.pageRotations?.[plan.pageIndex] ?? 0,
          language: options.language,
          nativeDpi: plan.nativeDpi,
          orientationFallback: true,
          nativeText: plan.nativeText,
          imageRects: plan.imageRects,
          signal: controller.signal,
          cleanup: !held.has(plan.pageIndex) && (options.canReleasePage?.(plan.pageIndex) ?? true),
          lookup: (key) => {
            const cached = getCachedOcrResult(key + regions)
            if (cached) storedKey = key + regions
            return cached
          },
          store: (key, value) => {
            storedKey = key + regions
            storeOcrResult(key + regions, value)
          },
          onProgress: ({ phase, progress }) => {
            if (!active.has(plan.pageIndex)) return
            if (phase === 'recognize' || phase === 'map' || phase === 'orientation') engineRunning = true
            active.set(plan.pageIndex, progress)
            publish()
          },
        })
        // A page that finished keeps its text, even when another page has just stopped the run.
        results.set(plan.pageIndex, result)
        if (storedKey) outcome.contentKeys.set(plan.pageIndex, storedKey)
      } catch (error) {
        if (isOcrAbort(error) && controller.signal.aborted) {
          finished = false
        } else if (isFatal(error)) {
          outcome.fatal ??= error as Error
          finished = false
          stop()
        } else {
          outcome.failed.push({ pageIndex: plan.pageIndex, message: failureMessage(error) })
        }
      } finally {
        active.delete(plan.pageIndex)
      }
      if (!finished) break
      done += 1
      publish()
    }
  }

  try {
    await Promise.all(Array.from({ length: Math.min(ocrWorkerCount(), queue.length) }, worker))
  } finally {
    signal.removeEventListener('abort', stop)
  }
  outcome.stopped = signal.aborted
  outcome.results = [...results.keys()].sort((a, b) => a - b).map((pageIndex) => results.get(pageIndex)!)
  outcome.failed.sort((a, b) => a.pageIndex - b.pageIndex)
  const written = pagesToWrite(outcome.results, options.replaceExisting)
  outcome.operation = written.length
    ? ocr.buildOcrLayerOperation(written, { replaceExisting: options.replaceExisting, language: options.language })
    : null
  return outcome
}

/** Pages whose text is written: pages without words matter only when they replace earlier text. */
export function pagesToWrite(results: readonly OcrPageResult[], replaceExisting: boolean) {
  return results.filter((result) => replaceExisting || result.wordCount > 0)
}

/** "page 4", "pages 4 and 7", "pages 4, 7, 9 and 2 more" (1-based). */
export function describePages(pageIndices: readonly number[]) {
  const numbers = [...pageIndices].sort((a, b) => a - b).map((index) => index + 1)
  if (numbers.length === 1) return `page ${numbers[0]}`
  if (numbers.length <= 3) return `pages ${numbers.slice(0, -1).join(', ')} and ${numbers.at(-1)}`
  return `pages ${numbers.slice(0, 3).join(', ')} and ${numbers.length - 3} more`
}

export interface OcrSummary {
  /** The toast text. */
  message: string
  /** Pages that look sideways and the rotation that makes them upright. */
  sideways: Array<{ pageIndex: number; degrees: number }>
}

/** The completion message (design microcopy). */
export function summarizeOcrRun(outcome: OcrRunOutcome, applied: boolean): OcrSummary {
  const withText = outcome.results.filter((result) => result.wordCount > 0)
  const weak = withText.filter((result) => result.meanConfidence < 60).map((result) => result.pageIndex)
  const sideways = withText
    .filter((result) => result.orientationCorrectedBy !== 0)
    .map((result) => ({ pageIndex: result.pageIndex, degrees: result.orientationCorrectedBy }))
  const notes: string[] = []
  if (weak.length) notes.push(`some words on ${describePages(weak)} may be inaccurate`)
  if (sideways.length) notes.push(`${describePages(sideways.map((item) => item.pageIndex)).replace(/^p/, 'P')} ${sideways.length === 1 ? 'looks' : 'look'} sideways`)
  if (outcome.failed.length) notes.push(`${describePages(outcome.failed.map((item) => item.pageIndex))} could not be read`)
  const suffix = notes.length ? ` · ${notes.join(' · ')}` : ''
  let message: string
  if (outcome.stopped) {
    message = applied && withText.length
      ? `Stopped — text added to ${withText.length} of ${outcome.planned} ${outcome.planned === 1 ? 'page' : 'pages'}.`
      : 'Stopped — no text was added.'
  } else if (applied && withText.length) {
    message = `Text recognized on ${withText.length} ${withText.length === 1 ? 'page' : 'pages'}. You can now search, select, and edit it.`
  } else if (outcome.planned && !outcome.failed.length) {
    message = outcome.planned === 1 ? 'No text was found on this page.' : 'No text was found on these pages.'
  } else if (outcome.failed.length && !withText.length) {
    message = outcome.failed[0].message
    return { message: outcome.failed.length > 1 ? `${message} (${describePages(outcome.failed.map((item) => item.pageIndex))})` : message, sideways: [] }
  } else {
    const reasons = new Set(outcome.skipped.map((item) => item.reason))
    message = reasons.has('has-ocr-text') && !reasons.has('has-text')
      ? 'This text was already recognized. Choose “Replace text recognized earlier” to read it again.'
      : reasons.has('blank') && reasons.size === 1
        ? 'There is nothing to recognize on these pages.'
        : 'These pages already have text.'
  }
  return { message: `${message}${suffix}`, sideways }
}

/** Ends a run that must not continue (another document opened, window closing). */
export function abortOcrRun(controller: AbortController | null) {
  if (controller && !controller.signal.aborted) controller.abort(ocrAbortError())
}
