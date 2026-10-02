// Shared OCR contracts for the renderer. The result and layer types are declared
// next to their implementation (electron/ocr-geometry.d.mts, which main-process
// code shares) and re-exported here so the app has one place to import from.
export type {
  OcrEngineInfo,
  OcrGeometryContext,
  OcrLayerLine,
  OcrLayerOperation,
  OcrLayerPage,
  OcrLayerWord,
  OcrLine,
  OcrPageResult,
  OcrParagraph,
  OcrRect,
  OcrRotation,
  OcrWord,
  PdfPoint,
} from '../../../electron/ocr-geometry.mjs'
export type { OcrDeskew, PreparePageStats, TesseractPsm } from '../../../electron/ocr-preprocess.mjs'

import type { OcrRotation } from '../../../electron/ocr-geometry.mjs'
import type { OcrDeskew, PreparePageStats } from '../../../electron/ocr-preprocess.mjs'

/** One recognisable language listed in dist/ocr/manifest.json. */
export interface OcrLanguageAsset {
  code: string
  label: string
  model: 'best_int'
  /** Path under ocr/, e.g. tessdata/eng.traineddata.gz. */
  file: string
  bytes: number
  sha256: string
}

/** dist/ocr/manifest.json as written by the simpleOcrAssets() Vite plugin. */
export interface OcrAssetManifest {
  schema: 1
  tesseractJs: string
  tesseractCore: string
  /** Under ocr/. */
  worker: 'worker.min.js'
  /** Under ocr/core/. Pinned: the relaxed-SIMD cores crash with float models. */
  core: 'tesseract-core-simd-lstm.wasm.js'
  languages: OcrLanguageAsset[]
  /** Every other file under ocr/ (paths relative to ocr/), with its size and sha256. */
  files: Record<string, { bytes: number; sha256: string }>
}

export type OcrErrorCode =
  | 'assets-missing'
  | 'assets-invalid'
  | 'language-unavailable'
  | 'engine-init'
  | 'engine-timeout'
  | 'engine-crashed'
  | 'recognize-failed'
  | 'prepare-failed'
  | 'render-failed'

/** A recognition failure with a message that can be shown to the user as is. */
export class OcrError extends Error {
  readonly code: OcrErrorCode
  /** Technical detail for logs (URLs, engine messages); not for the UI. */
  readonly detail?: string

  constructor(code: OcrErrorCode, message: string, detail?: string) {
    super(message)
    this.name = 'OcrError'
    this.code = code
    this.detail = detail
  }
}

export function ocrAbortError(): DOMException {
  return new DOMException('Text recognition was stopped.', 'AbortError')
}

export function isOcrAbort(error: unknown): boolean {
  return (error as { name?: unknown } | null)?.name === 'AbortError'
}

/** Main thread -> prep worker. `rgba` is transferred. */
export interface OcrPrepRequest {
  type: 'prepare-page'
  id: number
  rgba: ArrayBuffer
  width: number
  height: number
  dpi: number
}

export interface OcrPreparedPage {
  type: 'prepared'
  id: number
  blank: boolean
  width: number
  height: number
  /** The P5 PGM to recognise; null for blank pages. Transferred. */
  pgm: Uint8Array | null
  /** sha256 of the PGM (header + pixels): the content part of the cache key. */
  imageSha256: string
  deskew: OcrDeskew | null
  stats: PreparePageStats
}

export type OcrPrepResponse = OcrPreparedPage | { type: 'error'; id: number; message: string }

export type OcrPagePhase = 'render' | 'prepare' | 'engine' | 'recognize' | 'map' | 'orientation'

export interface OcrPageProgress {
  phase: OcrPagePhase
  /** Fraction of this page's work, 0..1. */
  progress: number
}

/** Filled in by recognizePage when the caller passes an object (benchmarks, smoke tests). */
export interface OcrPageDiagnostics {
  renderMs?: number
  prepareMs?: number
  recognizeMs?: number
  mapMs?: number
  totalMs?: number
  width?: number
  height?: number
  dpi?: number
  intent?: 'display' | 'print'
  blank?: boolean
  cacheHit?: boolean
  prep?: PreparePageStats
  /** Present when the orientation fallback ran. */
  orientation?: {
    probes: Array<{ correction: OcrRotation; tesseractConfidence: number; wordCount: number; verticalLines: number }>
    chosen: OcrRotation
  }
}
