// src/imaging/autoEnhance.ts (WP2)
// One-click corrections (design section 5.10): Simple mode's Auto (a QuickAdjust with hidden Levels) and
// Photoshop's Auto Tone / Auto Contrast / Auto Color (Levels specs). Statistics come from a proxy of at most
// 1 MP (regular sampling), so the cost is bounded for any image size. Pure and DOM-free.
import type { AdjustmentSpec, AutoEnhanceResult, LevelsChannel, PixelBuffer, QuickAdjust } from './types.ts'
import { assertBuffer } from './buffer.ts'
import { SRGB_TO_LINEAR } from './color.ts'
import { histogramPercentile } from './histogram.ts'

const PROXY_PIXELS = 1_000_000

function identityLevels(): LevelsChannel {
  return { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 }
}

function levelsSpec(rgb: LevelsChannel, red: LevelsChannel, green: LevelsChannel, blue: LevelsChannel): AdjustmentSpec {
  return { type: 'levels', rgb, red, green, blue }
}

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value
}

function round2(value: number): number {
  return Math.round(value * 100) / 100
}

/** Visits about PROXY_PIXELS evenly spaced pixels (all of them for images up to 1 MP). */
function forEachSample(src: PixelBuffer, visit: (i: number) => void): void {
  const { width, height } = src
  const step = Math.max(1, Math.ceil(Math.sqrt((width * height) / PROXY_PIXELS)))
  const offset = step >> 1
  for (let y = step > 1 ? Math.min(offset, height - 1) : 0; y < height; y += step) {
    const row = y * width
    for (let x = step > 1 ? Math.min(offset, width - 1) : 0; x < width; x += step) visit((row + x) * 4)
  }
}

function luma709(r: number, g: number, b: number): number {
  return (2126 * r + 7152 * g + 722 * b) / 10000
}

/** Levels channel that maps [black, white] to [0, 255] with gamma; identity when the range is too small. */
function stretchChannel(black: number, white: number, gamma = 1, minimumRange = 2): LevelsChannel {
  if (!(white - black >= minimumRange)) return identityLevels()
  return { inBlack: clamp(Math.round(black), 0, 254), inWhite: clamp(Math.round(white), 1, 255), gamma: round2(clamp(gamma, 0.1, 9.99)), outBlack: 0, outWhite: 255 }
}

/**
 * Simple mode's Auto: luma percentiles 0.4 / 99.6 become hidden Levels black and white points, the stretched
 * median sets gamma (0.75..1.35), muted images get +12 saturation, a colour cast in near-neutral mid-tones sets
 * warmth (-30..30), clipped highlights (> 2%) get -10 highlights and crushed shadows (> 3%) +10 shadows.
 * Pixels with alpha below 128 are ignored.
 */
export function autoEnhance(src: PixelBuffer): AutoEnhanceResult {
  assertBuffer(src)
  const data = src.data
  const luma = new Uint32Array(256)
  let count = 0
  let chromaSum = 0
  let clipped = 0
  let crushed = 0
  let neutralCount = 0
  let neutralRed = 0
  let neutralBlue = 0
  forEachSample(src, (i) => {
    if (data[i + 3] < 128) return
    const r = data[i]
    const g = data[i + 1]
    const b = data[i + 2]
    const max = r > g ? (r > b ? r : b) : (g > b ? g : b)
    const min = r < g ? (r < b ? r : b) : (g < b ? g : b)
    const y = luma709(r, g, b)
    luma[Math.round(y)] += 1
    count += 1
    chromaSum += (max - min) / 255
    if (max >= 254) clipped += 1
    if (max <= 2) crushed += 1
    if (max - min <= 38 && y >= 51 && y <= 217) {
      neutralCount += 1
      neutralRed += SRGB_TO_LINEAR[r]
      neutralBlue += SRGB_TO_LINEAR[b]
    }
  })
  const quiet: QuickAdjust = { exposure: 0, brightness: 0, contrast: 0, highlights: 0, shadows: 0, saturation: 0, warmth: 0, auto: null }
  if (count === 0) return { quick: quiet, strength: 0 }

  const black = histogramPercentile(luma, 0.004)
  const white = histogramPercentile(luma, 0.996)
  const median = histogramPercentile(luma, 0.5)
  let auto: LevelsChannel | null = null
  let gamma = 1
  if (white - black >= 16) {
    const m = clamp((median - black) / (white - black), 0.01, 0.99)
    gamma = round2(clamp(Math.log(m) / Math.log(0.5), 0.75, 1.35))
    if (black > 0 || white < 255 || gamma !== 1) auto = { inBlack: black, inWhite: white, gamma, outBlack: 0, outWhite: 255 }
  }
  const meanChroma = chromaSum / count
  const saturation = meanChroma < 0.18 ? 12 : 0
  let warmth = 0
  if (neutralCount >= Math.max(64, count * 0.005) && neutralRed > 0 && neutralBlue > 0) {
    warmth = clamp(Math.round(60 * Math.log2(neutralBlue / neutralRed)), -30, 30)
  }
  const highlights = clipped / count > 0.02 ? -10 : 0
  const shadows = crushed / count > 0.03 ? 10 : 0
  const stretch = auto ? (auto.inBlack + (255 - auto.inWhite)) / 255 : 0
  const strength = clamp(Math.max(
    stretch * 2.5,
    (Math.abs(Math.log(gamma)) / Math.log(1.35)) * 0.6,
    (Math.abs(warmth) / 30) * 0.6,
    saturation ? 0.2 : 0,
    (highlights ? 0.15 : 0) + (shadows ? 0.15 : 0),
  ), 0, 1)
  return { quick: { exposure: 0, brightness: 0, contrast: 0, highlights, shadows, saturation, warmth: warmth || 0, auto }, strength }
}

/** Per-channel histograms (alpha > 0) of the proxy. */
function channelHistograms(src: PixelBuffer): { red: Uint32Array; green: Uint32Array; blue: Uint32Array; luma: Uint32Array; count: number } {
  const data = src.data
  const red = new Uint32Array(256)
  const green = new Uint32Array(256)
  const blue = new Uint32Array(256)
  const luma = new Uint32Array(256)
  let count = 0
  forEachSample(src, (i) => {
    if (data[i + 3] === 0) return
    red[data[i]] += 1
    green[data[i + 1]] += 1
    blue[data[i + 2]] += 1
    luma[Math.round(luma709(data[i], data[i + 1], data[i + 2]))] += 1
    count += 1
  })
  return { red, green, blue, luma, count }
}

const CLIP = 0.001

/** Photoshop Auto Tone ("Enhance Per Channel Contrast"): each channel stretched with 0.1% clipping. */
export function autoTone(src: PixelBuffer): AdjustmentSpec {
  assertBuffer(src)
  const h = channelHistograms(src)
  if (h.count === 0) return levelsSpec(identityLevels(), identityLevels(), identityLevels(), identityLevels())
  const channel = (bins: Uint32Array) => stretchChannel(histogramPercentile(bins, CLIP), histogramPercentile(bins, 1 - CLIP))
  return levelsSpec(identityLevels(), channel(h.red), channel(h.green), channel(h.blue))
}

/**
 * Photoshop Auto Contrast ("Enhance Monochromatic Contrast"): one black and white point for all channels
 * (0.1% clipping of the combined R, G and B values), applied on the RGB master so colours keep their balance.
 */
export function autoContrast(src: PixelBuffer): AdjustmentSpec {
  assertBuffer(src)
  const h = channelHistograms(src)
  if (h.count === 0) return levelsSpec(identityLevels(), identityLevels(), identityLevels(), identityLevels())
  const combined = new Uint32Array(256)
  for (let value = 0; value < 256; value += 1) combined[value] = h.red[value] + h.green[value] + h.blue[value]
  const rgb = stretchChannel(histogramPercentile(combined, CLIP), histogramPercentile(combined, 1 - CLIP))
  return levelsSpec(rgb, identityLevels(), identityLevels(), identityLevels())
}

/**
 * Photoshop Auto Color ("Find Dark & Light Colors" + "Snap Neutral Midtones"): the average colours of the
 * darkest and lightest 0.1% (by luma) become per-channel black and white points; then per-channel gamma moves
 * the average near-neutral mid-tone to neutral grey.
 */
export function autoColor(src: PixelBuffer): AdjustmentSpec {
  assertBuffer(src)
  const data = src.data
  const h = channelHistograms(src)
  const none = levelsSpec(identityLevels(), identityLevels(), identityLevels(), identityLevels())
  if (h.count === 0) return none
  const darkLimit = histogramPercentile(h.luma, CLIP)
  const lightLimit = histogramPercentile(h.luma, 1 - CLIP)
  const dark = [0, 0, 0]
  const light = [0, 0, 0]
  let darkCount = 0
  let lightCount = 0
  forEachSample(src, (i) => {
    if (data[i + 3] === 0) return
    const y = Math.round(luma709(data[i], data[i + 1], data[i + 2]))
    if (y <= darkLimit) {
      dark[0] += data[i]; dark[1] += data[i + 1]; dark[2] += data[i + 2]
      darkCount += 1
    }
    if (y >= lightLimit) {
      light[0] += data[i]; light[1] += data[i + 1]; light[2] += data[i + 2]
      lightCount += 1
    }
  })
  if (!darkCount || !lightCount) return none
  const channels = [0, 1, 2].map((c) => stretchChannel(dark[c] / darkCount, light[c] / lightCount, 1, 8))
  // Snap neutral mid-tones: mean of near-neutral mid-tone pixels after the stretch.
  const sums = [0, 0, 0]
  let neutral = 0
  const stretched = [0, 0, 0]
  forEachSample(src, (i) => {
    if (data[i + 3] === 0) return
    for (let c = 0; c < 3; c += 1) {
      const level = channels[c]
      const span = level.inWhite - level.inBlack
      stretched[c] = span > 0 ? clamp((data[i + c] - level.inBlack) / span, 0, 1) : data[i + c] / 255
    }
    const max = Math.max(stretched[0], stretched[1], stretched[2])
    const min = Math.min(stretched[0], stretched[1], stretched[2])
    const y = 0.2126 * stretched[0] + 0.7152 * stretched[1] + 0.0722 * stretched[2]
    if (max - min > 0.12 || y < 0.2 || y > 0.8) return
    sums[0] += stretched[0]; sums[1] += stretched[1]; sums[2] += stretched[2]
    neutral += 1
  })
  if (neutral >= Math.max(32, h.count * 0.002)) {
    const means = sums.map((sum) => sum / neutral)
    const target = (means[0] + means[1] + means[2]) / 3
    for (let c = 0; c < 3; c += 1) {
      if (!(means[c] > 0 && means[c] < 1 && target > 0 && target < 1)) continue
      // ((v - ib) / (iw - ib))^(1 / gamma) = target at the channel mean.
      channels[c] = { ...channels[c], gamma: round2(clamp(Math.log(means[c]) / Math.log(target), 0.6, 1.6)) }
    }
  }
  return levelsSpec(identityLevels(), channels[0], channels[1], channels[2])
}
