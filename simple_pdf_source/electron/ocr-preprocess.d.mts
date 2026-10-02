/** One grey byte per pixel, rows top to bottom. */
export interface GrayImage { width: number; height: number; data: Uint8Array }
/** The OCR image was rotated by -degrees about (cx, cy); deskewPoint maps OCR points back. */
export interface OcrDeskew { degrees: number; cx: number; cy: number }
export interface InkLevels { ink: number | null; paper: number }
export interface PreparePageStats {
  dpi?: number
  inkLevel: number | null
  paperLevel: number
  contrast: number
  skewDegrees: number
  noiseSigma: number
  denoised: boolean
  backgroundSigma: number
  binarized: boolean
  inkPixels: number
  timings: Record<string, number>
}
export interface PreparedPage { image: GrayImage; blank: boolean; deskew: OcrDeskew | null; stats: PreparePageStats }
export type TesseractPsm = '3' | '4' | '6' | '7' | '11'

export const PREP_VERSION: string
export const OCR_ENGINE_TAG: string
export const DESKEW_MIN_DEGREES: number
export const NOISE_SIGMA_THRESHOLD: number
export const SAUVOLA_BACKGROUND_SIGMA: number
export const LOW_CONTRAST_LEVELS: number
export const SAUVOLA_INK_FLOOR: number

export function toGray(pixels: ArrayLike<number> & { subarray?: (start: number, end: number) => ArrayLike<number> }, width: number, height: number, channels?: 1 | 3 | 4): GrayImage
export function histogram(image: GrayImage): Uint32Array
export function normalizeBackground(image: GrayImage, options?: { block?: number }): GrayImage
export function measureInkLevels(image: GrayImage, options?: { low?: number; high?: number }): InkLevels
export function stretchContrast(image: GrayImage, options?: { low?: number; high?: number; minRange?: number; levels?: InkLevels }): GrayImage
export function estimateNoise(image: GrayImage): number
export function median3(image: GrayImage): GrayImage
export function estimateSkew(image: GrayImage, options?: { maxDegrees?: number; threshold?: number; maxPoints?: number; step?: number }): number
export function rotateGray(image: GrayImage, degrees: number, options?: { fill?: number }): GrayImage
export function deskewPoint(x: number, y: number, deskew: OcrDeskew | null | undefined): { x: number; y: number }
export function estimateBackgroundSigma(image: GrayImage): number
export function sauvola(image: GrayImage, options?: { window?: number; k?: number; range?: number; inkFloor?: number }): GrayImage
export function countInkPixels(image: GrayImage, threshold?: number): number
export function minimumInkPixels(width: number, height: number): number
export function isBlank(image: GrayImage, options?: { minInkPixels?: number }): boolean
export function encodePgm(image: GrayImage): Uint8Array
export function decodePgm(bytes: Uint8Array): GrayImage
export function engineSignature(options?: { language?: string; psm?: TesseractPsm }): string
export function tesseractParameters(options: { dpi: number; psm?: TesseractPsm }): { user_defined_dpi: string; tessedit_pageseg_mode: string; preserve_interword_spaces: string }
export function sha256Hex(bytes: Uint8Array): Promise<string>
export function preparePage(pixels: ArrayLike<number>, width: number, height: number, options?: { channels?: 1 | 3 | 4; dpi?: number; minInkPixels?: number }): PreparedPage
