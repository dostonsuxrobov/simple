/** Pixel position in a region, top-left origin. */
export interface PixelPoint { x: number; y: number }
/** Baseline origin and unit reading direction in region pixels (y down). */
export interface PixelBaseline { x: number; y: number; dx: number; dy: number }
export interface PixelBox { x0: number; y0: number; x1: number; y1: number }

export interface RetouchWord {
  /** First mask label of the word (its own label for target boxes). */
  label: number
  /** Every mask label that belongs to the word (segmented words may span several clusters). */
  labels: number[]
  /** The recognised word, when segmentation matched the text. */
  text?: string
  /** Ink box (pixels) and extent along the baseline (u0..u1), or null without ink. */
  box: (PixelBox & { u0: number; u1: number }) | null
}

export interface RetouchResult {
  width: number
  height: number
  channels: 4
  bilevel: boolean
  paper: number
  ink: number
  /** Region pixels with the masked ones inpainted (RGBA, alpha 255). */
  filled: Uint8ClampedArray
  /** Word label of each masked pixel; -1 outside the mask. */
  labels: Int16Array
  /** Pixels a patch never changes: ink that is kept and a one-pixel rim around it. */
  blocked: Uint8Array
  words: RetouchWord[]
  inkBox: PixelBox | null
  maskPixels: number
  noiseSigma: number[]
  paperColor: number[]
  /** True when words were found in the pixels and matched to the text. */
  segmented: boolean
  radius?: number
  reason?: string
}

export interface RetouchOptions {
  channels?: 1 | 3 | 4
  dpi?: number
  baseline?: PixelBaseline
  /** Line extent along the baseline, pixels. */
  length?: number
  /** Em size, pixels. */
  fontSize?: number
  /** x-height, pixels (default half the em). */
  xHeight?: number
  /** Word boxes to remove as quads [top-left, top-right, bottom-right, bottom-left]; default: the whole line. */
  targets?: Array<{ quad: [PixelPoint, PixelPoint, PixelPoint, PixelPoint] } | [PixelPoint, PixelPoint, PixelPoint, PixelPoint]>
  /** Find the words in the pixels (no targets) and match them to `text`. */
  segment?: boolean
  text?: string
  /** How wide the recognised words should be (default sans). */
  fontClass?: 'serif' | 'sans' | 'mono'
  /** Grain seed: the same seed gives the same patch. */
  seed?: string | number
}

export interface PatchImage {
  rgba: Uint8ClampedArray
  width: number
  height: number
  /** Top-left of the patch in the region. */
  x: number
  y: number
}

export const RETOUCH_VERSION: string
export const RETOUCH: Readonly<Record<string, number>>
export function grayOf(pixels: ArrayLike<number>, width: number, height: number, channels?: 1 | 3 | 4): Uint8Array
export function rgbaOf(pixels: ArrayLike<number>, width: number, height: number, channels?: 1 | 3 | 4): Uint8ClampedArray
export function grayHistogram(gray: Uint8Array): Uint32Array
export function histogramQuantile(hist: Uint32Array, total: number, fraction: number): number
export function estimateLevels(gray: Uint8Array): { paper: number; ink: number; contrast: number; bilevel: boolean; hist: Uint32Array }
export function localPaper(gray: Uint8Array, width: number, height: number, block?: number): Float32Array
export function lineFrame(baseline: PixelBaseline): { dx: number; dy: number; u(x: number, y: number): number; v(x: number, y: number): number; point(u: number, v: number): PixelPoint }
export function labelComponents(mask: ArrayLike<number>, width: number, height: number): { labels: Int32Array; count: number; area: Int32Array; minX: Int32Array; minY: Int32Array; maxX: Int32Array; maxY: Int32Array }
export function dilateMask(mask: ArrayLike<number>, width: number, height: number, radius: number): Uint8Array
export function pushPullFill(values: Float32Array, known: ArrayLike<number>, width: number, height: number): Float32Array
export function seedOf(seed: unknown): number
export function createRandom(seed: unknown): () => number
export function grainModel(rgba: Uint8ClampedArray, known: ArrayLike<number>, width: number, height: number, limit?: number): { samples: number; sigma: number[]; sample(random: () => number, out: number[]): number[] } | null
export function clusterWords(items: Array<{ u0: number; u1: number }>, gap: number): Array<{ u0: number; u1: number; members: number[] }>
export function expectedWordWidth(token: string, fontClass?: 'serif' | 'sans' | 'mono'): number
export function alignClustersToTokens(clusters: Array<{ u0: number; u1: number }>, tokens: string[], options?: { em?: number; fontClass?: 'serif' | 'sans' | 'mono' }): Array<{ u0: number; u1: number; clusters: number[] }> | null
export function textTokens(text: string): string[]
export function retouchLine(pixels: ArrayLike<number>, width: number, height: number, options?: RetouchOptions): RetouchResult
export function composePatch(
  retouch: Pick<RetouchResult, 'width' | 'height' | 'labels' | 'blocked' | 'filled' | 'bilevel'> & { paperColor?: number[] },
  include?: Set<number> | null,
): PatchImage | null
export function compositePatch(regionRgba: Uint8ClampedArray, width: number, height: number, patch: PatchImage | null): Uint8ClampedArray
