// Font style of a scanned text line, measured from its pixels (design section
// 4.8.3 and the calibration in 2.5): serif / sans / monospace, weight, italic,
// size, ink and paper colour. Tesseract's LSTM engine reports no font
// information, so an edit of scanned text picks a matching Windows font from
// these measurements. Pure: typed arrays in, plain objects out.

import { estimateLevels, grayOf, labelComponents, lineFrame, localPaper } from './ocr-retouch.mjs'

/** The Windows family that stands in for each class (also known to the PDF writer). */
export const SCAN_FONT_FAMILIES = Object.freeze({ serif: 'Times New Roman', sans: 'Arial', mono: 'Courier New' })

/**
 * Ink heights of the stand-in families as a share of the em: `ascender` is
 * what the tall glyphs of running text (capitals, digits, b d f h k l)
 * measure on average, `xHeight` the lowercase body. Calibrated on lines
 * rendered at 300 DPI (tests/scan-style.test.cjs).
 */
export const SCAN_FONT_METRICS = Object.freeze({
  serif: Object.freeze({ regular: { ascender: 0.677, xHeight: 0.45 }, bold: { ascender: 0.677, xHeight: 0.47 } }),
  sans: Object.freeze({ regular: { ascender: 0.731, xHeight: 0.532 }, bold: { ascender: 0.731, xHeight: 0.541 } }),
  mono: Object.freeze({ regular: { ascender: 0.597, xHeight: 0.419 }, bold: { ascender: 0.597, xHeight: 0.434 } }),
})

export const SCAN_STYLE = Object.freeze({
  /** Monospace when a glyph's advance is this many x-heights or more... */
  monoPitch: 1.22,
  /** ...or this many with evenly spaced glyphs (coefficient of variation below monoEvenness). */
  monoPitchEven: 1.1,
  monoEvenness: 0.12,
  /** Serif when at least this share of stem-like glyphs (and two or more) has feet on both sides. */
  serifScore: 0.2,
  minSerifFeet: 2,
  /** Bold when the stroke is wider than this share of the x-height. */
  bold: Object.freeze({ serif: 0.235, sans: 0.205, mono: 0.16 }),
  /** Italic when the sharpest shear is at least this many degrees... */
  italicDegrees: 7,
  /** ...and sharpens the vertical strokes by at least this factor. */
  italicEnergy: 1.05,
  /** Size weighting of the ascender against the x-height (design: 2:1). */
  ascenderWeight: 2,
  minFontSize: 4,
  maxFontSize: 96,
})

// Characters by where their ink sits on the line (Latin text).
const X_HEIGHT_CHARS = /[acemnorsuvwxzıµ]/gu
const TALL_CHARS = /[A-Zbdfhklt0-9ßÀ-ÖØ-Þ]/gu

const quantile = (sorted, fraction) => {
  if (!sorted.length) return NaN
  const position = Math.max(0, Math.min(sorted.length - 1, fraction * (sorted.length - 1)))
  const low = Math.floor(position)
  const high = Math.ceil(position)
  return sorted[low] + (sorted[high] - sorted[low]) * (position - low)
}
const countMatches = (text, pattern) => (String(text).match(pattern) || []).length
// Pixels at least a quarter covered count as ink, so measured extents run about half a pixel long.
const ANTIALIAS_BIAS = 0.5

/**
 * Measure a line: `pixels` (RGBA or `channels`-channel) of a region around
 * it, `baseline` {x, y, dx, dy} and `length` (pixels along it), `fontSize`
 * (the em size in pixels, e.g. from the OCR layer; used for bands only) and
 * the recognised `text`. `words` (optional) are the words' [u0, u1] extents
 * along the baseline.
 */
export function measureLineFeatures(pixels, width, height, options = {}) {
  const channels = options.channels ?? 4
  const count = width * height
  const gray = grayOf(pixels, width, height, channels)
  const levels = estimateLevels(gray)
  const fontSize = Math.max(4, Number(options.fontSize) || height * 0.5)
  const length = Math.max(1, Number(options.length) || width)
  const frame = lineFrame(options.baseline ?? { x: 0, y: height * 0.75, dx: 1, dy: 0 })
  const text = String(options.text ?? '')
  // Coverage against the paper around each pixel, so stains and shading are not ink.
  const paperAt = localPaper(gray, width, height, Math.max(12, 0.45 * fontSize))
  const coverage = new Float32Array(count)
  for (let index = 0; index < count; index += 1) {
    const paper = Math.max(paperAt[index], levels.ink + 24)
    coverage[index] = Math.max(0, Math.min(1, (paper - gray[index]) / (paper - levels.ink)))
  }
  const mask = new Uint8Array(count)
  for (let index = 0; index < count; index += 1) mask[index] = coverage[index] >= 0.25 ? 1 : 0

  const components = labelComponents(mask, width, height)
  const stats = []
  for (let id = 0; id <= components.count; id += 1) stats.push({ id, area: components.area[id], inside: 0, u0: Infinity, u1: -Infinity, v0: Infinity, v1: -Infinity })
  const bandTop = 1.05 * fontSize
  const bandBottom = -0.4 * fontSize
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const id = components.labels[y * width + x]
      if (!id) continue
      const entry = stats[id]
      const u = frame.u(x, y)
      const v = frame.v(x, y)
      if (v <= bandTop && v >= bandBottom && u >= -0.5 * fontSize && u <= length + 0.5 * fontSize) entry.inside += 1
      // Pixel extents in the line frame (centres, widened by half a pixel each way).
      if (u - 0.5 < entry.u0) entry.u0 = u - 0.5
      if (u + 0.5 > entry.u1) entry.u1 = u + 0.5
      if (v - 0.5 < entry.v0) entry.v0 = v - 0.5
      if (v + 0.5 > entry.v1) entry.v1 = v + 0.5
    }
  }
  const line = stats.filter((entry) => entry.id && entry.inside * 2 > entry.area && entry.area >= 3)
  const lineIds = new Uint8Array(components.count + 1)
  for (const entry of line) lineIds[entry.id] = 1

  // The baseline as the glyphs show it: OCR baselines can be a pixel or two off.
  // Flat-bottomed glyphs sit on it, round ones overshoot below: take the upper bottoms.
  const bottoms = line.filter((entry) => entry.v1 - entry.v0 >= 0.3 * fontSize && entry.v0 > -0.12 * fontSize && entry.v0 < 0.15 * fontSize).map((entry) => entry.v0).sort((a, b) => a - b)
  const shift = bottoms.length ? quantile(bottoms, 0.7) : 0
  const tolerance = Math.max(1.5, 0.07 * fontSize)
  const resting = line.filter((entry) => Math.abs(entry.v0 - shift) <= tolerance && entry.v1 - shift >= 0.2 * fontSize)
  // Heights above the baseline; antialiasing widens every glyph by about half a pixel.
  const heights = resting.map((entry) => entry.v1 - shift - ANTIALIAS_BIAS).sort((a, b) => a - b)

  // Which share of the resting glyphs should be lowercase bodies (from the text).
  const xChars = countMatches(text, X_HEIGHT_CHARS) + countMatches(text, /[i]/gu)
  const tallChars = countMatches(text, TALL_CHARS)
  const restingChars = Math.max(1, xChars + tallChars)
  const xShare = xChars / restingChars
  let xHeight = NaN
  let ascender = NaN
  if (heights.length) {
    if (xChars >= 2 && xShare >= 0.12) xHeight = quantile(heights, 0.5 * xShare)
    if (tallChars >= 1 && 1 - xShare >= 0.06) ascender = quantile(heights, Math.min(0.97, xShare + 0.6 * (1 - xShare)))
    // Without text (or a line of one kind), fall back to the design's quantiles.
    if (!Number.isFinite(xHeight) && !Number.isFinite(ascender)) {
      xHeight = quantile(heights, 0.35)
      ascender = quantile(heights, 0.9)
    }
  }

  // Stroke: ink across rows at 35-55 % of the x-height, summed as coverage
  // (sub-pixel accurate), median over the runs.
  const body = Number.isFinite(xHeight) ? xHeight : Number.isFinite(ascender) ? ascender * 0.68 : fontSize * 0.5
  const sample = (u, v) => {
    const point = frame.point(u, v)
    const x = Math.floor(point.x)
    const y = Math.floor(point.y)
    if (x < 0 || y < 0 || x >= width || y >= height) return 0
    const index = y * width + x
    return lineIds[components.labels[index]] ? coverage[index] : 0
  }
  const runs = []
  for (const fraction of [0.35, 0.45, 0.55]) {
    const v = shift + fraction * body
    let run = 0
    let peak = 0
    for (let u = -0.5 * fontSize; u <= length + 0.5 * fontSize; u += 1) {
      const value = sample(u, v)
      if (value >= 0.15) {
        run += value
        peak = Math.max(peak, value)
      } else if (run) {
        if (peak >= 0.6) runs.push(run)
        run = 0
        peak = 0
      }
    }
    if (run && peak >= 0.6) runs.push(run)
  }
  runs.sort((a, b) => a - b)
  const stroke = runs.length ? quantile(runs, 0.5) : NaN

  // Italic: the shear that makes the vertical strokes sharpest (largest
  // energy of the ink's projection along the baseline), as in skew estimation.
  let slant = 0
  let italicDegrees = 0
  if (line.length) {
    const points = []
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const index = y * width + x
        if (!lineIds[components.labels[index]] || coverage[index] < 0.5) continue
        const v = frame.v(x, y) - shift
        if (v < 0 || v > body) continue
        points.push(frame.u(x, y), v)
      }
    }
    if (points.length >= 40) {
      let low = Infinity
      let high = -Infinity
      let tallest = 0
      for (let index = 0; index < points.length; index += 2) {
        if (points[index] < low) low = points[index]
        if (points[index] > high) high = points[index]
        if (points[index + 1] > tallest) tallest = points[index + 1]
      }
      const reach = Math.ceil(tallest * Math.tan((20 * Math.PI) / 180)) + 2
      const histogram = new Int32Array(Math.ceil(high - low) + 2 * reach + 3)
      const offset = reach + 1 - low
      const energy = (degrees) => {
        const slope = Math.tan((degrees * Math.PI) / 180)
        histogram.fill(0)
        for (let index = 0; index < points.length; index += 2) histogram[Math.round(points[index] - points[index + 1] * slope + offset)] += 1
        let total = 0
        for (let index = 0; index < histogram.length; index += 1) total += histogram[index] * histogram[index]
        return total
      }
      const upright = energy(0)
      let best = 0
      let bestEnergy = upright
      for (let degrees = -20; degrees <= 20; degrees += 1) {
        const value = energy(degrees)
        if (value > bestEnergy) { bestEnergy = value; best = degrees }
      }
      if (best >= SCAN_STYLE.italicDegrees && bestEnergy >= upright * SCAN_STYLE.italicEnergy) {
        italicDegrees = best
        slant = Math.tan((best * Math.PI) / 180)
      }
    }
  }
  // Positions along the line with the slant taken out (upright glyphs: unchanged).
  const along = (x, y) => frame.u(x, y) - (frame.v(x, y) - shift) * slant
  const extents = new Map()
  if (slant) {
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const id = components.labels[y * width + x]
        if (!id || !lineIds[id]) continue
        const u = along(x, y)
        const entry = extents.get(id)
        if (entry) { if (u - 0.5 < entry.u0) entry.u0 = u - 0.5; if (u + 0.5 > entry.u1) entry.u1 = u + 0.5 }
        else extents.set(id, { u0: u - 0.5, u1: u + 0.5 })
      }
    }
  }
  const uprightExtent = (entry) => extents.get(entry.id) ?? entry

  // Serifs: stem-like glyphs (narrow, at least body high) with feet on both sides.
  let stems = 0
  let feet = 0
  if (Number.isFinite(stroke) && stroke > 0) {
    for (const entry of resting) {
      const upright = uprightExtent(entry)
      const glyphWidth = upright.u1 - upright.u0
      const glyphHeight = entry.v1 - shift - ANTIALIAS_BIAS
      if (glyphWidth > Math.min(4.5 * stroke, 0.8 * body) || glyphHeight < 0.8 * body) continue
      const extent = (v) => {
        let first = Infinity
        let last = -Infinity
        const offset = (v - shift) * slant
        for (let u = upright.u0 - 1; u <= upright.u1 + 1; u += 0.5) {
          const point = frame.point(u + offset, v)
          const x = Math.floor(point.x)
          const y = Math.floor(point.y)
          if (x < 0 || y < 0 || x >= width || y >= height) continue
          const index = y * width + x
          if (components.labels[index] !== entry.id || coverage[index] < 0.35) continue
          if (u < first) first = u
          if (u > last) last = u
        }
        return Number.isFinite(first) ? { first, last } : null
      }
      const middle = extent(shift + 0.5 * body)
      // A serif is often a single pixel row: look at the glyph's two lowest rows.
      const feetRows = [extent(shift + 0.5), extent(shift + 1.5)].filter(Boolean)
      const foot = feetRows.length ? { first: Math.min(...feetRows.map((row) => row.first)), last: Math.max(...feetRows.map((row) => row.last)) } : null
      if (!middle || !foot) continue
      stems += 1
      const reach = Math.max(1, 0.18 * stroke)
      if (middle.first - foot.first >= reach && foot.last - middle.last >= reach) feet += 1
    }
  }
  const serifScore = stems ? feet / stems : 0
  const serifFeet = feet

  // Pitch: advance per character over the words (or the whole line), in x-heights.
  const letters = String(text).replace(/\s+/gu, '')
  const glyphs = Array.from(letters).length
  let inkWidth = 0
  if (Array.isArray(options.words) && options.words.length) {
    for (const word of options.words) inkWidth += Math.max(0, word.u1 - word.u0)
  } else if (line.length) {
    inkWidth = Math.max(...line.map((entry) => uprightExtent(entry).u1)) - Math.min(...line.map((entry) => uprightExtent(entry).u0))
  }
  const wordCount = Array.isArray(options.words) && options.words.length ? options.words.length : Math.max(1, String(text).trim().split(/\s+/u).filter(Boolean).length)
  const spaces = Array.isArray(options.words) && options.words.length ? 0 : Math.max(0, wordCount - 1)
  const pitch = glyphs ? inkWidth / Math.max(1, glyphs + spaces) : NaN
  // Evenness of glyph spacing: centre-to-centre steps of neighbouring glyphs.
  const centres = resting.map((entry) => { const upright = uprightExtent(entry); return (upright.u0 + upright.u1) / 2 }).sort((a, b) => a - b)
  const steps = []
  for (let index = 1; index < centres.length; index += 1) {
    const step = centres[index] - centres[index - 1]
    if (Number.isFinite(pitch) && step > 0.4 * pitch && step < 1.6 * pitch) steps.push(step)
  }
  const stepMean = steps.reduce((sum, step) => sum + step, 0) / Math.max(1, steps.length)
  const stepSpread = Math.sqrt(steps.reduce((sum, step) => sum + (step - stepMean) ** 2, 0) / Math.max(1, steps.length))
  const evenness = steps.length >= 6 && stepMean > 0 ? stepSpread / stepMean : NaN

  // Colour: the solid core of the strokes, and the paper around them.
  const rgb = (index, channel) => (channels === 1 ? pixels[index] : pixels[index * channels + channel])
  const cores = []
  for (let index = 0; index < count; index += 1) if (lineIds[components.labels[index]]) cores.push(coverage[index])
  cores.sort((a, b) => a - b)
  const solid = cores.length ? Math.max(0.5, quantile(cores, 0.9) * 0.85) : 1
  const inkSamples = [[], [], []]
  const paperSamples = [[], [], []]
  for (let index = 0; index < count; index += 1) {
    if (lineIds[components.labels[index]] && coverage[index] >= solid) {
      for (let channel = 0; channel < 3; channel += 1) inkSamples[channel].push(rgb(index, channel))
    } else if (coverage[index] < 0.08 && !mask[index]) {
      for (let channel = 0; channel < 3; channel += 1) paperSamples[channel].push(rgb(index, channel))
    }
  }
  // Medians of byte values, from histograms (the paper has many pixels).
  const byteMedian = (values) => {
    const hist = new Uint32Array(256)
    for (const value of values) hist[Math.max(0, Math.min(255, Math.round(value)))] += 1
    let seen = 0
    for (let level = 0; level < 256; level += 1) {
      seen += hist[level]
      if (seen * 2 > values.length) return level
    }
    return 255
  }
  const colorOf = (samples, fallback) => samples.map((values, channel) => (values.length ? byteMedian(values) / 255 : fallback[channel]))

  return {
    xHeight,
    ascender,
    stroke,
    strokeRatio: Number.isFinite(stroke) && Number.isFinite(body) && body > 0 ? stroke / body : NaN,
    serifScore,
    serifFeet,
    stems,
    pitchRatio: Number.isFinite(pitch) && body > 0 ? pitch / body : NaN,
    evenness,
    italicDegrees,
    baselineShift: shift,
    components: line.length,
    resting: resting.length,
    color: colorOf(inkSamples, [0, 0, 0]),
    background: colorOf(paperSamples, [1, 1, 1]),
    contrast: levels.contrast,
  }
}

/**
 * Turn line measurements into a style: class, stand-in family, weight,
 * italic, size in points (`dpi` is the resolution the pixels were measured
 * at), colours and a 0..1 confidence. `fontSizeHint` (points) is used only
 * when the glyphs give no usable height.
 */
export function classifyScanStyle(features, options = {}) {
  const dpi = Number(options.dpi) > 0 ? Number(options.dpi) : 300
  const toPoints = (pixels) => (pixels * 72) / dpi
  const pitch = features.pitchRatio
  const mono = Number.isFinite(pitch) && (pitch >= SCAN_STYLE.monoPitch
    || (pitch >= SCAN_STYLE.monoPitchEven && Number.isFinite(features.evenness) && features.evenness <= SCAN_STYLE.monoEvenness))
  const serif = features.serifScore >= SCAN_STYLE.serifScore && (features.serifFeet ?? 0) >= SCAN_STYLE.minSerifFeet
  const fontClass = mono ? 'mono' : serif ? 'serif' : 'sans'
  const bold = Number.isFinite(features.strokeRatio) && features.strokeRatio > SCAN_STYLE.bold[fontClass]
  const metrics = SCAN_FONT_METRICS[fontClass][bold ? 'bold' : 'regular']
  const estimates = []
  if (Number.isFinite(features.ascender) && features.ascender > 0) estimates.push([toPoints(features.ascender) / metrics.ascender, SCAN_STYLE.ascenderWeight])
  if (Number.isFinite(features.xHeight) && features.xHeight > 0) estimates.push([toPoints(features.xHeight) / metrics.xHeight, 1])
  let fontSize = estimates.length
    ? estimates.reduce((sum, [value, weight]) => sum + value * weight, 0) / estimates.reduce((sum, [, weight]) => sum + weight, 0)
    : Number(options.fontSizeHint) || 11
  // Two estimates that disagree badly: one quantile landed on the wrong
  // glyph group; trust the one nearer the OCR size when there is one.
  if (estimates.length === 2 && Number(options.fontSizeHint) > 0) {
    const [a, b] = estimates.map(([value]) => value)
    if (Math.abs(a - b) / Math.max(a, b) > 0.22) fontSize = Math.abs(a - options.fontSizeHint) <= Math.abs(b - options.fontSizeHint) ? a : b
  }
  fontSize = Math.max(SCAN_STYLE.minFontSize, Math.min(SCAN_STYLE.maxFontSize, Math.round(fontSize * 10) / 10))
  const italic = Math.abs(features.italicDegrees) >= SCAN_STYLE.italicDegrees && features.italicDegrees > 0
  // Confidence: how much evidence there was, and how far from each decision boundary.
  const evidence = Math.min(1, (features.resting || 0) / 14)
  const margin = (value, threshold, scale) => (Number.isFinite(value) ? Math.min(1, Math.abs(value - threshold) / scale) : 0)
  const certainty = Math.min(
    mono ? margin(pitch, SCAN_STYLE.monoPitch, 0.15) || 0.5 : Math.max(margin(features.serifScore, SCAN_STYLE.serifScore, 0.3), 0.2),
    margin(features.strokeRatio, SCAN_STYLE.bold[fontClass], 0.06) || 0.3,
  )
  return {
    fontClass,
    fontFamily: SCAN_FONT_FAMILIES[fontClass],
    fontWeight: bold ? 700 : 400,
    italic,
    fontSize,
    xHeight: Number.isFinite(features.xHeight) ? Math.round(toPoints(features.xHeight) * 100) / 100 : Math.round(fontSize * metrics.xHeight * 100) / 100,
    ...(Number.isFinite(features.ascender) ? { ascender: Math.round(toPoints(features.ascender) * 100) / 100 } : {}),
    strokeRatio: Number.isFinite(features.strokeRatio) ? Math.round(features.strokeRatio * 1000) / 1000 : 0,
    color: features.color.map((channel) => Math.round(channel * 1000) / 1000),
    background: features.background.map((channel) => Math.round(channel * 1000) / 1000),
    confidence: Math.round(evidence * Math.max(0.2, certainty) * 100) / 100,
  }
}

/** Measure and classify in one step; also returns the raw measurements. */
export function estimateScanStyle(pixels, width, height, options = {}) {
  const features = measureLineFeatures(pixels, width, height, options)
  const dpi = Number(options.dpi) > 0 ? Number(options.dpi) : 300
  const style = classifyScanStyle(features, { dpi, fontSizeHint: Number(options.fontSize) > 0 ? (Number(options.fontSize) * 72) / dpi : undefined })
  return { style, features }
}
