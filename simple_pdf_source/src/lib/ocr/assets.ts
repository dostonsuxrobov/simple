import { OcrError, ocrAbortError, type OcrAssetManifest } from './types'

// The OCR runtime ships next to index.html (dist/ocr/, emitted by the
// simpleOcrAssets() Vite plugin) and is read from file:// (or app.asar) at run
// time, so recognition never needs the network.

export const OCR_WORKER_FILE = 'worker.min.js'
/**
 * A file, not a directory: given a directory tesseract.js picks the relaxed-SIMD
 * core when the CPU supports it, and that core crashes with float models.
 */
export const OCR_CORE_FILE = 'tesseract-core-simd-lstm.wasm.js'
export const OCR_MISSING_MESSAGE = 'Text recognition (OCR) components are missing from this build.'
const OCR_INVALID_MESSAGE = 'Text recognition (OCR) components in this build are damaged.'
const DEFAULT_PREFLIGHT_TIMEOUT_MS = 8_000

export interface OcrAssetLocation {
  /** Absolute URL of the ocr/ folder, with a trailing slash. */
  base: string
  manifestUrl: string
  workerPath: string
  corePath: string
  /** tesseract.js appends /<lang>.traineddata.gz. */
  langPath: string
}

export function ocrAssetLocation(base?: string): OcrAssetLocation {
  const root = new URL(base ?? './ocr/', document.baseURI)
  if (!root.pathname.endsWith('/')) root.pathname += '/'
  return {
    base: root.href,
    manifestUrl: new URL('manifest.json', root).href,
    workerPath: new URL(OCR_WORKER_FILE, root).href,
    corePath: new URL(`core/${OCR_CORE_FILE}`, root).href,
    langPath: new URL('tessdata', root).href,
  }
}

export function ocrLanguageCodes(language: string): string[] {
  const codes = String(language).split('+').map((code) => code.trim()).filter(Boolean)
  if (!codes.length || codes.some((code) => !/^[a-z]{3}(?:_[a-z]+)?$/.test(code))) {
    throw new OcrError('language-unavailable', 'This text recognition language is not supported.', String(language))
  }
  return [...new Set(codes)]
}

function validateManifest(value: unknown): OcrAssetManifest {
  const manifest = value as Partial<OcrAssetManifest> | null
  const languagesValid = Array.isArray(manifest?.languages) && manifest.languages.every((language) => language
    && typeof language.code === 'string' && typeof language.file === 'string' && Number.isFinite(language.bytes))
  if (!manifest || manifest.schema !== 1 || manifest.worker !== OCR_WORKER_FILE || manifest.core !== OCR_CORE_FILE
    || !languagesValid || !manifest.files || typeof manifest.files !== 'object') {
    throw new OcrError('assets-invalid', OCR_INVALID_MESSAGE, 'manifest.json has an unexpected shape')
  }
  return manifest as OcrAssetManifest
}

type ProbeKind = 'script' | 'gzip'

async function probe(url: string, expectedBytes: number | undefined, kind: ProbeKind, signal: AbortSignal) {
  let response: Response
  try {
    response = await fetch(url, { signal })
  } catch (error) {
    if (signal.aborted) throw error
    throw new OcrError('assets-missing', OCR_MISSING_MESSAGE, `${url}: ${String(error)}`)
  }
  if (!response.ok) throw new OcrError('assets-missing', OCR_MISSING_MESSAGE, `${url}: HTTP ${response.status}`)
  const length = Number(response.headers.get('content-length'))
  if (expectedBytes !== undefined && Number.isFinite(length) && length > 0 && length !== expectedBytes) {
    void response.body?.cancel().catch(() => {})
    throw new OcrError('assets-invalid', OCR_INVALID_MESSAGE, `${url}: ${length} bytes, expected ${expectedBytes}`)
  }
  // Only the first chunk is needed to tell a real asset from an empty or HTML error response.
  let head: Uint8Array | undefined
  const reader = response.body?.getReader()
  if (reader) {
    head = (await reader.read()).value
    void reader.cancel().catch(() => {})
  } else {
    head = new Uint8Array(await response.arrayBuffer())
  }
  const validHead = !!head && head.length > 0 && (kind === 'gzip' ? head[0] === 0x1f && head[1] === 0x8b : head[0] !== 0x3c)
  if (!validHead) throw new OcrError('assets-invalid', OCR_INVALID_MESSAGE, `${url}: unexpected content`)
}

async function runPreflight(location: OcrAssetLocation, languages: string[], timeoutMs: number): Promise<OcrAssetManifest> {
  const controller = new AbortController()
  let timedOut = false
  const timer = setTimeout(() => { timedOut = true; controller.abort() }, timeoutMs)
  try {
    let manifest: OcrAssetManifest
    try {
      const response = await fetch(location.manifestUrl, { signal: controller.signal })
      if (!response.ok) throw new Error(`HTTP ${response.status}`)
      manifest = validateManifest(await response.json())
    } catch (error) {
      if (error instanceof OcrError || timedOut) throw error
      throw new OcrError('assets-missing', OCR_MISSING_MESSAGE, `${location.manifestUrl}: ${String(error)}`)
    }
    const missing = languages.filter((code) => !manifest.languages.some((language) => language.code === code))
    if (missing.length) {
      throw new OcrError('language-unavailable', 'Text recognition for this language is not included in this build.', missing.join(', '))
    }
    await Promise.all([
      probe(location.workerPath, manifest.files[OCR_WORKER_FILE]?.bytes, 'script', controller.signal),
      probe(location.corePath, manifest.files[`core/${OCR_CORE_FILE}`]?.bytes, 'script', controller.signal),
      ...languages.map((code) => {
        const language = manifest.languages.find((candidate) => candidate.code === code)!
        return probe(new URL(language.file, location.base).href, language.bytes, 'gzip', controller.signal)
      }),
    ])
    return manifest
  } catch (error) {
    if (timedOut) throw new OcrError('assets-missing', OCR_MISSING_MESSAGE, `No answer from ${location.base} within ${timeoutMs} ms`)
    throw error instanceof OcrError ? error : new OcrError('assets-missing', OCR_MISSING_MESSAGE, String(error))
  } finally {
    clearTimeout(timer)
    // Stop any probe still reading once one has failed.
    controller.abort()
  }
}

const preflights = new Map<string, Promise<OcrAssetManifest>>()

/**
 * Check, before any tesseract.js worker is created, that the manifest and the
 * worker, core and language files are present and look right. tesseract.js 7
 * never settles createWorker when a language file fails to load, so a broken
 * build must be caught here. A success is remembered for the page lifetime; a
 * failure is retried on the next call.
 */
export function preflightOcrAssets(options: { base?: string; languages?: string[]; timeoutMs?: number; signal?: AbortSignal } = {}): Promise<OcrAssetManifest> {
  if (options.signal?.aborted) return Promise.reject(ocrAbortError())
  let location: OcrAssetLocation
  let languages: string[]
  try {
    location = ocrAssetLocation(options.base)
    languages = (options.languages ?? ['eng']).flatMap(ocrLanguageCodes)
  } catch (error) {
    return Promise.reject(error)
  }
  const key = `${location.base}|${languages.join('+')}`
  let pending = preflights.get(key)
  if (!pending) {
    const started = runPreflight(location, languages, options.timeoutMs ?? DEFAULT_PREFLIGHT_TIMEOUT_MS)
    pending = started
    preflights.set(key, started)
    started.catch(() => { if (preflights.get(key) === started) preflights.delete(key) })
  }
  const signal = options.signal
  if (!signal) return pending
  return new Promise<OcrAssetManifest>((resolve, reject) => {
    const onAbort = () => reject(ocrAbortError())
    signal.addEventListener('abort', onAbort, { once: true })
    pending!.then(resolve, reject).finally(() => signal.removeEventListener('abort', onAbort))
  })
}
