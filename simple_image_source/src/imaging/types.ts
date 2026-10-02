// src/imaging/types.ts
// Pure, DOM-free contracts shared by Simple mode, Advanced mode, the imaging worker and Node tests.
// Rules for every file under src/imaging/:
//   - erasable TypeScript only (no enum, namespace, parameter properties, decorators);
//   - relative imports carry an explicit `.ts` extension and type-only imports use `import type`;
//   - never reference document/window/HTMLCanvasElement (workers and Node must load these modules).

// ---------------------------------------------------------------------------------------------
// Geometry and pixel containers
// ---------------------------------------------------------------------------------------------

/** Straight (non-premultiplied) RGBA8, row-major, top-left origin. Structurally compatible with ImageData. */
export interface PixelBuffer {
  readonly width: number
  readonly height: number
  /** length === width * height * 4 */
  readonly data: Uint8ClampedArray
}

/** One 8-bit coverage channel: 0 = none / hidden / unselected, 255 = full. */
export interface MaskBuffer {
  readonly width: number
  readonly height: number
  /** length === width * height */
  readonly data: Uint8Array
}

/** Integer pixel rectangle. width/height >= 0. */
export interface IntRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

/** Sub-pixel rectangle (selection marquees, crop boxes before rounding). */
export interface Rect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export interface Point {
  readonly x: number
  readonly y: number
}

export interface Size {
  readonly width: number
  readonly height: number
}

/** 0..255 per channel. */
export interface Rgb8 {
  readonly r: number
  readonly g: number
  readonly b: number
}

/** 0..255 per channel, alpha included. */
export interface Rgba8 extends Rgb8 {
  readonly a: number
}

/** Canvas / DOMMatrix order: x' = a*x + c*y + e, y' = b*x + d*y + f. */
export type Affine = readonly [a: number, b: number, c: number, d: number, e: number, f: number]

/** Row-major 3x3 matrix mapping [x, y, 1] to [X, Y, W]; the mapped point is (X / W, Y / W). */
export type Homography = readonly [number, number, number, number, number, number, number, number, number]

export type Interpolation = 'nearest' | 'bilinear' | 'bicubic' | 'lanczos3'

/** 'auto' = area pre-filter then lanczos3 when shrinking by more than 2x, lanczos3 when shrinking, bicubic when enlarging. */
export type ResampleMethod = Interpolation | 'area' | 'auto'

export type RotateFit = 'expand' | 'crop-inscribed' | 'same-size'

export interface OpOptions {
  /** 0..1; implementations call it at most ~20 times per operation. */
  readonly onProgress?: (fraction: number) => void
  /** Checked between rows/stripes. On abort, throw an Error whose name is 'AbortError'. */
  readonly signal?: AbortSignal
}

// ---------------------------------------------------------------------------------------------
// Blend modes (Photoshop's full list, minus group-only 'pass through')
// ---------------------------------------------------------------------------------------------

export type BlendMode =
  | 'normal' | 'dissolve'
  | 'darken' | 'multiply' | 'color-burn' | 'linear-burn' | 'darker-color'
  | 'lighten' | 'screen' | 'color-dodge' | 'linear-dodge' | 'lighter-color'
  | 'overlay' | 'soft-light' | 'hard-light' | 'vivid-light' | 'linear-light' | 'pin-light' | 'hard-mix'
  | 'difference' | 'exclusion' | 'subtract' | 'divide'
  | 'hue' | 'saturation' | 'color' | 'luminosity'

// ---------------------------------------------------------------------------------------------
// Adjustments (shared by Simple quick adjust, destructive Image > Adjustments and adjustment layers)
// ---------------------------------------------------------------------------------------------

/** Photoshop Levels channel. Inputs/outputs 0..255; gamma 0.1..9.99 (1 = neutral). */
export interface LevelsChannel {
  readonly inBlack: number
  readonly inWhite: number
  readonly gamma: number
  readonly outBlack: number
  readonly outWhite: number
}

/** Curve control point, both axes 0..255. Arrays are sorted by x, 2..16 points, unique x. */
export interface CurvePoint {
  readonly x: number
  readonly y: number
}

export type HueRange = 'reds' | 'yellows' | 'greens' | 'cyans' | 'blues' | 'magentas'

export interface HslShift {
  /** -180..180 degrees */
  readonly hue: number
  /** -100..100 */
  readonly saturation: number
  /** -100..100 */
  readonly lightness: number
}

/** Each -100..100 (Photoshop Color Balance sliders). */
export interface ColorBalanceShift {
  readonly cyanRed: number
  readonly magentaGreen: number
  readonly yellowBlue: number
}

export interface GradientStop {
  /** 0..1 along the gradient. */
  readonly position: number
  readonly color: Rgb8
  /** 0..1 interpolation midpoint to the next stop; default 0.5. */
  readonly midpoint?: number
}

export interface OpacityStop {
  readonly position: number
  /** 0..1 */
  readonly opacity: number
}

/** Simple-mode sliders, each -100..100. `auto` is set by Auto-enhance and is not a visible slider. */
export interface QuickAdjust {
  readonly exposure: number
  readonly brightness: number
  readonly contrast: number
  readonly highlights: number
  readonly shadows: number
  readonly saturation: number
  readonly warmth: number
  readonly auto: LevelsChannel | null
}

export type AdjustmentSpec =
  | { readonly type: 'brightness-contrast'; readonly brightness: number /* -150..150 */; readonly contrast: number /* -50..100 */; readonly legacy: boolean }
  | { readonly type: 'levels'; readonly rgb: LevelsChannel; readonly red: LevelsChannel; readonly green: LevelsChannel; readonly blue: LevelsChannel }
  | { readonly type: 'curves'; readonly rgb: readonly CurvePoint[]; readonly red: readonly CurvePoint[]; readonly green: readonly CurvePoint[]; readonly blue: readonly CurvePoint[] }
  | { readonly type: 'exposure'; readonly exposure: number /* -20..20 stops */; readonly offset: number /* -0.5..0.5 */; readonly gamma: number /* 0.01..9.99 */ }
  | { readonly type: 'vibrance'; readonly vibrance: number /* -100..100 */; readonly saturation: number /* -100..100 */ }
  | { readonly type: 'hue-saturation'; readonly colorize: boolean; readonly master: HslShift; readonly ranges: Readonly<Partial<Record<HueRange, HslShift>>> }
  | { readonly type: 'color-balance'; readonly shadows: ColorBalanceShift; readonly midtones: ColorBalanceShift; readonly highlights: ColorBalanceShift; readonly preserveLuminosity: boolean }
  | { readonly type: 'black-white'; readonly reds: number; readonly yellows: number; readonly greens: number; readonly cyans: number; readonly blues: number; readonly magentas: number /* each -200..300 */; readonly tint: Rgb8 | null }
  | { readonly type: 'photo-filter'; readonly color: Rgb8; readonly density: number /* 0..100 */; readonly preserveLuminosity: boolean }
  | { readonly type: 'invert' }
  | { readonly type: 'posterize'; readonly levels: number /* 2..255 */ }
  | { readonly type: 'threshold'; readonly level: number /* 1..255 */ }
  | { readonly type: 'gradient-map'; readonly stops: readonly GradientStop[]; readonly reverse: boolean; readonly dither: boolean }
  | ({ readonly type: 'quick' } & QuickAdjust)

export type AdjustmentType = AdjustmentSpec['type']

/** A compiled, reusable per-pixel transform. Pure point operation: never reads neighbours. */
export interface AdjustmentKernel {
  readonly spec: AdjustmentSpec
  readonly isIdentity: boolean
  /** In place on straight RGBA; alpha untouched. Processes pixels [startPixel, endPixel). */
  apply(data: Uint8ClampedArray, startPixel?: number, endPixel?: number, ditherSeed?: number): void
}

// ---------------------------------------------------------------------------------------------
// Filters (neighbourhood operations) and one-click looks
// ---------------------------------------------------------------------------------------------

export type FilterSpec =
  | { readonly type: 'gaussian-blur'; readonly radius: number /* 0.1..250, = sigma in px */ }
  | { readonly type: 'motion-blur'; readonly angle: number /* -90..90 deg */; readonly distance: number /* 1..999 px */ }
  | { readonly type: 'unsharp-mask'; readonly amount: number /* 1..500 % */; readonly radius: number /* 0.1..250 */; readonly threshold: number /* 0..255 */ }
  | { readonly type: 'sharpen'; readonly strength: 'normal' | 'more' }
  | { readonly type: 'add-noise'; readonly amount: number /* 0.1..400 % */; readonly distribution: 'uniform' | 'gaussian'; readonly monochromatic: boolean; readonly seed: number }
  | { readonly type: 'median'; readonly radius: number /* 1..100 */ }
  | { readonly type: 'reduce-noise'; readonly strength: number /* 0..10 */; readonly preserveDetails: number /* 0..100 % */ }
  | { readonly type: 'pixelate'; readonly cellSize: number /* 2..200 */ }
  | { readonly type: 'emboss'; readonly angle: number /* -180..180 */; readonly height: number /* 1..10 */; readonly amount: number /* 1..500 % */ }
  | { readonly type: 'find-edges' }

export type FilterType = FilterSpec['type']

export type LookId = 'none' | 'vivid' | 'warm' | 'cool' | 'mono' | 'sepia' | 'vintage' | 'dramatic' | 'fade'

// ---------------------------------------------------------------------------------------------
// Selection masks, flood fill, brushes, healing, gradients
// ---------------------------------------------------------------------------------------------

export type SelectionOp = 'replace' | 'add' | 'subtract' | 'intersect'
export type FillRule = 'nonzero' | 'evenodd'

export interface FloodOptions {
  /** 0..255; a pixel matches when max(|dr|,|dg|,|db|[,|da|]) <= tolerance. Photoshop default 32. */
  readonly tolerance: number
  readonly contiguous: boolean
  readonly antiAlias: boolean
  readonly compareAlpha: boolean
}

export interface BrushTip {
  /** px, 1..5000 */
  readonly diameter: number
  /** 0..1 (1 = hard edge, still anti-aliased) */
  readonly hardness: number
  /** 0.01..1 (1 = circle) */
  readonly roundness: number
  /** degrees */
  readonly angle: number
}

export interface StrokeSample {
  readonly x: number
  readonly y: number
  /** 0..1; mouse input reports 1 */
  readonly pressure: number
  readonly time: number
}

export interface Dab {
  readonly x: number
  readonly y: number
  /** effective diameter after pressure */
  readonly diameter: number
  /** 0..1 per-dab flow after pressure */
  readonly flow: number
}

export interface HealOptions {
  readonly seed?: number
  readonly maxIterations?: number
  /** Candidate source rings, as multiples of the hole diameter. Default [1, 1.5, 2, 3]. */
  readonly rings?: readonly number[]
}

export type GradientKind = 'linear' | 'radial' | 'angle' | 'reflected' | 'diamond'

export interface GradientSpec {
  readonly kind: GradientKind
  readonly from: Point
  readonly to: Point
  readonly stops: readonly GradientStop[]
  readonly opacityStops: readonly OpacityStop[]
  readonly reverse: boolean
  readonly dither: boolean
}

export interface Histogram {
  readonly red: Uint32Array
  readonly green: Uint32Array
  readonly blue: Uint32Array
  /** Rec. 709 luma of gamma-encoded values, 256 bins */
  readonly luma: Uint32Array
  readonly alpha: Uint32Array
  /** pixels counted (alpha > 0 and mask > 0) */
  readonly count: number
}

export interface AutoEnhanceResult {
  readonly quick: QuickAdjust
  /** 0..1, how far the image was from the target statistics */
  readonly strength: number
}

// ---------------------------------------------------------------------------------------------
// Worker protocol (src/shared/imaging.worker.ts routes these to src/imaging/* functions)
// ---------------------------------------------------------------------------------------------

export interface ImagingJobMap {
  readonly adjust: { readonly input: { readonly src: PixelBuffer; readonly specs: readonly AdjustmentSpec[]; readonly mask: MaskBuffer | null; readonly opacity: number }; readonly output: PixelBuffer }
  readonly filter: { readonly input: { readonly src: PixelBuffer; readonly spec: FilterSpec; readonly mask: MaskBuffer | null }; readonly output: PixelBuffer }
  readonly look: { readonly input: { readonly src: PixelBuffer; readonly look: LookId; readonly intensity: number }; readonly output: PixelBuffer }
  readonly resample: { readonly input: { readonly src: PixelBuffer; readonly width: number; readonly height: number; readonly method: ResampleMethod }; readonly output: PixelBuffer }
  readonly rotate: { readonly input: { readonly src: PixelBuffer; readonly degrees: number; readonly fit: RotateFit; readonly interpolation: Interpolation; readonly background: Rgba8 }; readonly output: PixelBuffer }
  readonly warp: { readonly input: { readonly src: PixelBuffer; readonly inverse: Homography; readonly out: IntRect; readonly interpolation: Interpolation }; readonly output: PixelBuffer }
  readonly histogram: { readonly input: { readonly src: PixelBuffer; readonly mask: MaskBuffer | null }; readonly output: Histogram }
  readonly autoEnhance: { readonly input: { readonly src: PixelBuffer }; readonly output: AutoEnhanceResult }
  readonly flood: { readonly input: { readonly src: PixelBuffer; readonly seed: Point; readonly options: FloodOptions }; readonly output: MaskBuffer }
  readonly feather: { readonly input: { readonly mask: MaskBuffer; readonly radius: number }; readonly output: MaskBuffer }
  readonly expand: { readonly input: { readonly mask: MaskBuffer; readonly pixels: number }; readonly output: MaskBuffer }
  readonly contract: { readonly input: { readonly mask: MaskBuffer; readonly pixels: number }; readonly output: MaskBuffer }
  readonly heal: { readonly input: { readonly src: PixelBuffer; readonly hole: MaskBuffer; readonly options: HealOptions }; readonly output: PixelBuffer }
}

export type ImagingOp = keyof ImagingJobMap
export type ImagingInput<K extends ImagingOp> = ImagingJobMap[K]['input']
export type ImagingOutput<K extends ImagingOp> = ImagingJobMap[K]['output']

/** A worker-op module (src/shared/worker-ops/*.ts) exports a partial table of handlers. */
export type ImagingHandlers = {
  readonly [K in ImagingOp]?: (input: ImagingInput<K>, options: OpOptions) => ImagingOutput<K> | Promise<ImagingOutput<K>>
}

export interface ImagingRunOptions {
  readonly onProgress?: (fraction: number) => void
  readonly signal?: AbortSignal
  /** Transfer input buffers to the worker (caller must not reuse them). Default true. */
  readonly transfer?: boolean
}

export interface ImagingClient {
  run<K extends ImagingOp>(op: K, input: ImagingInput<K>, options?: ImagingRunOptions): Promise<ImagingOutput<K>>
  /** Terminates workers; pending runs reject with AbortError. */
  dispose(): void
}
