import type { PixelBaseline } from './ocr-retouch.mjs'

export type ScanFontClass = 'serif' | 'sans' | 'mono'

export interface ScanLineFeatures {
  /** Pixels. */
  xHeight: number
  ascender: number
  stroke: number
  strokeRatio: number
  serifScore: number
  serifFeet: number
  stems: number
  pitchRatio: number
  evenness: number
  italicDegrees: number
  baselineShift: number
  components: number
  resting: number
  /** 0..1 RGB. */
  color: number[]
  background: number[]
  contrast: number
}

export interface ScanStyleEstimate {
  fontClass: ScanFontClass
  fontFamily: string
  fontWeight: 400 | 700
  italic: boolean
  /** Points. */
  fontSize: number
  xHeight: number
  ascender?: number
  strokeRatio: number
  color: [number, number, number]
  background: [number, number, number]
  confidence: number
}

export interface ScanStyleOptions {
  channels?: 1 | 3 | 4
  dpi?: number
  baseline?: PixelBaseline
  length?: number
  /** Em size in pixels (bands only). */
  fontSize?: number
  text?: string
  words?: Array<{ u0: number; u1: number }>
}

export const SCAN_FONT_FAMILIES: Readonly<Record<ScanFontClass, string>>
export const SCAN_FONT_METRICS: Readonly<Record<ScanFontClass, { regular: { ascender: number; xHeight: number }; bold: { ascender: number; xHeight: number } }>>
export const SCAN_STYLE: Readonly<Record<string, unknown>>
export function measureLineFeatures(pixels: ArrayLike<number>, width: number, height: number, options?: ScanStyleOptions): ScanLineFeatures
export function classifyScanStyle(features: ScanLineFeatures, options?: { dpi?: number; fontSizeHint?: number }): ScanStyleEstimate
export function estimateScanStyle(pixels: ArrayLike<number>, width: number, height: number, options?: ScanStyleOptions): { style: ScanStyleEstimate; features: ScanLineFeatures }
