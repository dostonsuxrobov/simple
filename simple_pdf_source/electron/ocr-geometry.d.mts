import type { OcrDeskew } from './ocr-preprocess.mjs'

/** A point in unrotated PDF user space. */
export interface PdfPoint { x: number; y: number }
/** An axis-aligned rectangle in unrotated PDF user space (same shape as src/types PdfRect). */
export interface OcrRect { x: number; y: number; width: number; height: number }
export type OcrRotation = 0 | 90 | 180 | 270

export interface OcrWord {
  text: string
  confidence: number
  /** Baseline start of the word. */
  origin: PdfPoint
  /** Unit text direction (the line's, snapped to the page axis below 0.5 deg). */
  dir: PdfPoint
  /** Ink extent along dir. */
  width: number
  /** Distance along dir to the next word's origin; 0 for the last word. */
  gap: number
  /** Box corners in reading orientation: top-left, top-right, bottom-right, bottom-left. */
  quad: [PdfPoint, PdfPoint, PdfPoint, PdfPoint]
  bbox: OcrRect
}
export interface OcrLine {
  id: string
  text: string
  confidence: number
  /** Font size for the invisible layer: Tesseract's row height in PDF units. */
  fontSize: number
  baseline: { origin: PdfPoint; dir: PdfPoint }
  bbox: OcrRect
  words: OcrWord[]
}
export interface OcrParagraph { id: string; bbox: OcrRect; lines: OcrLine[]; ltr: boolean }
export interface OcrEngineInfo { name: 'tesseract.js'; version: '7.0.0'; core: 'simd-lstm'; language: string; model: 'best_int' }
export interface OcrPageResult {
  schema: 1
  /** Page index at recognition time; cache consumers rebind it after page moves. */
  pageIndex: number
  /** sha256 of the OCR input bitmap | engine signature | pixel-to-PDF mapping. */
  contentKey: string
  engine: OcrEngineInfo
  /** Rotation the page was rendered at for recognition. */
  rotation: OcrRotation
  dpi: number
  deskewDegrees: number
  orientationCorrectedBy: OcrRotation
  meanConfidence: number
  wordCount: number
  paragraphs: OcrParagraph[]
}

export interface TesseractBoxLike { x0: number; y0: number; x1: number; y1: number }
export interface TesseractSymbolLike { text: string; bbox: TesseractBoxLike; is_superscript?: boolean | number; is_subscript?: boolean | number; is_dropcap?: boolean | number }
/** `symbols` (when present) refine the line baseline from glyph bottoms. */
export interface TesseractWordLike { text: string; confidence: number; bbox: TesseractBoxLike; symbols?: TesseractSymbolLike[] }
export interface TesseractLineLike { words: TesseractWordLike[]; baseline?: TesseractBoxLike; rowAttributes?: { rowHeight?: number }; bbox?: TesseractBoxLike }
export interface TesseractParagraphLike { lines: TesseractLineLike[]; is_ltr?: boolean | number }
export interface TesseractBlockLike { paragraphs: TesseractParagraphLike[] }

export interface OcrFilterSettings {
  minWordConfidence: number
  weakWordConfidence: number
  weakWordMinAlphanumerics: number
  weakWordMinLineConfidence: number
  minLineConfidence: number
  snapDegrees: number
  duplicateIou: number
}

export interface OcrGeometryContext {
  /** pdf.js viewport.convertToPdfPoint of the OCR render (OCR pixels -> PDF user space). */
  toPdf(x: number, y: number): PdfPoint | number[]
  dpi: number
  pageIndex: number
  contentKey: string
  rotation: OcrRotation
  deskew?: OcrDeskew | null
  language?: string
  orientationCorrectedBy?: OcrRotation
  /** Visible native text rects (mixed pages): overlapping OCR words are duplicates. */
  nativeText?: OcrRect[]
  /** When given, only words centred inside these rects are kept (mixed pages). */
  imageRects?: OcrRect[]
  filter?: Partial<OcrFilterSettings>
}

export interface OcrLayerWord { text: string; x: number; y: number; dx: number; dy: number; width: number; gap: number }
export interface OcrLayerLine { fontSize: number; words: OcrLayerWord[] }
export interface OcrLayerPage { pageIndex: number; lines: OcrLayerLine[] }
export interface OcrLayerOperation {
  type: 'ocr-text-layer'
  replaceExisting: boolean
  meta: { engine: string; language: string }
  pages: OcrLayerPage[]
}

export const OCR_RESULT_SCHEMA: 1
export const OCR_ENGINE_INFO: Readonly<Omit<OcrEngineInfo, 'language'>>
export const OCR_FILTER: Readonly<OcrFilterSettings>
export function normalizeOcrText(text: string): string
/** Recognised lines/words and how many lines Tesseract read vertically (text it turned by itself). */
export function tesseractReadingStats(blocks: readonly TesseractBlockLike[] | null | undefined): { lines: number; verticalLines: number; words: number }

/** One recognition pass, as the orientation fallback sees it. */
export interface OcrReading {
  /** Tesseract's page confidence (MeanTextConf), before this module's filters. */
  tesseractConfidence: number
  /** OcrPageResult.meanConfidence / wordCount (after filtering). */
  meanConfidence: number
  wordCount: number
  /** Tesseract's lines with text, and how many of them run vertically. */
  lines: number
  verticalLines: number
}
export const ORIENTATION_POLICY: Readonly<{ minConfidence: number; verticalShare: number; improvement: number; verticalTolerance: number }>
export function needsOrientationCheck(reading: OcrReading | null | undefined): boolean
export function readingScore(reading: OcrReading): number
export function pickOrientation<T extends { reading: OcrReading }>(first: OcrReading, probes: readonly T[]): T | null
export function tesseractToOcrPage(blocks: readonly TesseractBlockLike[] | null | undefined, ctx: OcrGeometryContext): OcrPageResult
export function ocrPageToLayerPayload(result: OcrPageResult): OcrLayerPage
export function buildOcrLayerOperation(results: readonly OcrPageResult[], options?: { replaceExisting?: boolean; language?: string }): OcrLayerOperation
