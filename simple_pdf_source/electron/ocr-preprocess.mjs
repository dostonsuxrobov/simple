// OCR input preparation (WP1). Pure functions over typed arrays so the same code
// runs in the renderer's prep worker and in Node tests: no DOM, no Node built-ins.
//
// Images are { width, height, data: Uint8Array } with one grey byte per pixel,
// rows top to bottom. Coordinates are continuous pixel-edge coordinates: pixel
// (i, j) covers [i, i + 1) x [j, j + 1) and its centre is (i + .5, j + .5). That
// matches Tesseract's word boxes and pdf.js viewport coordinates.

/** Bump whenever a change here can alter the OCR input bitmap (part of cache keys). */
export const PREP_VERSION = 'prep1'
/** tesseract.js 7 with the pinned SIMD (non-relaxed) LSTM core. */
export const OCR_ENGINE_TAG = 'tjs7|core-simd-lstm'
/** Deskewing below this angle cost accuracy on the fixtures (1.2 deg: 99.2% vs 100%). */
export const DESKEW_MIN_DEGREES = 1.5
/** Noise sigma (grey levels, after stretching) above which the 3x3 median runs. */
export const NOISE_SIGMA_THRESHOLD = 6
/** Background sigma after normalisation above which the page is binarised with Sauvola. */
export const SAUVOLA_BACKGROUND_SIGMA = 18
/**
 * Paper-to-ink contrast (after normalisation) below which global thresholding
 * loses faint strokes, so the page is binarised with Sauvola instead. Faded
 * toner, pencil and carbon copies fall below it; the stained low-contrast
 * fixture measured 88 levels, ordinary scans 200 or more.
 */
export const LOW_CONTRAST_LEVELS = 160
/** Pixels darker than this stay ink under Sauvola, so thick strokes do not hollow out. */
export const SAUVOLA_INK_FLOOR = 64

function assertDimensions(width, height) {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) {
    throw new Error('OCR image dimensions must be positive integers.')
  }
  if (width * height > 80_000_000) throw new Error('OCR image is too large.')
}

function assertImage(image) {
  if (!image || !(image.data instanceof Uint8Array)) throw new Error('Expected a greyscale OCR image.')
  assertDimensions(image.width, image.height)
  if (image.data.length < image.width * image.height) throw new Error('OCR image data is truncated.')
}

const clampByte = (value) => (value <= 0 ? 0 : value >= 255 ? 255 : value)

/**
 * BT.601 luma of an 8-bit RGBA (canvas), RGB or grey buffer. Partially
 * transparent pixels are composited over white paper.
 */
export function toGray(pixels, width, height, channels = 4) {
  assertDimensions(width, height)
  if (![1, 3, 4].includes(channels)) throw new Error('OCR pixels must have 1, 3 or 4 channels.')
  const count = width * height
  if (!pixels || pixels.length < count * channels) throw new Error('OCR pixel data is truncated.')
  const out = new Uint8Array(count)
  if (channels === 1) {
    out.set(pixels.subarray ? pixels.subarray(0, count) : Array.prototype.slice.call(pixels, 0, count))
    return { width, height, data: out }
  }
  for (let i = 0, j = 0; i < count; i += 1, j += channels) {
    // 0.299 R + 0.587 G + 0.114 B in 16.16 fixed point (weights sum to 65536).
    let y = (pixels[j] * 19595 + pixels[j + 1] * 38470 + pixels[j + 2] * 7471 + 32768) >> 16
    if (channels === 4) {
      const alpha = pixels[j + 3]
      if (alpha !== 255) y = ((y * alpha + 255 * (255 - alpha)) / 255 + 0.5) | 0
    }
    out[i] = y
  }
  return { width, height, data: out }
}

export function histogram(image) {
  assertImage(image)
  const hist = new Uint32Array(256)
  const { data } = image
  const count = image.width * image.height
  for (let i = 0; i < count; i += 1) hist[data[i]] += 1
  return hist
}

function percentile(hist, total, fraction) {
  const target = Math.max(1, total * fraction)
  let accumulated = 0
  for (let value = 0; value < 256; value += 1) {
    accumulated += hist[value]
    if (accumulated >= target) return value
  }
  return 255
}

/**
 * Flatten uneven illumination, stains and paper tint: estimate the paper level
 * per block (90th percentile), close dark blocks with a max filter, smooth, and
 * divide every pixel by the bilinearly upsampled background.
 */
export function normalizeBackground(image, options = {}) {
  assertImage(image)
  const block = Math.max(4, Math.round(options.block ?? 16))
  const { width: w, height: h, data } = image
  const bw = Math.ceil(w / block)
  const bh = Math.ceil(h / block)
  const levels = new Float32Array(bw * bh)
  const hist = new Uint32Array(256)
  for (let by = 0; by < bh; by += 1) {
    const y0 = by * block
    const y1 = Math.min(h, y0 + block)
    for (let bx = 0; bx < bw; bx += 1) {
      const x0 = bx * block
      const x1 = Math.min(w, x0 + block)
      hist.fill(0)
      for (let y = y0; y < y1; y += 1) {
        const row = y * w
        for (let x = x0; x < x1; x += 1) hist[data[row + x]] += 1
      }
      // The 90th percentile counted from the bright end: paper dominates blocks.
      const skip = ((y1 - y0) * (x1 - x0)) * 0.1
      let accumulated = 0
      let level = 0
      for (let value = 255; value >= 0; value -= 1) {
        accumulated += hist[value]
        if (accumulated > skip) { level = value; break }
      }
      levels[by * bw + bx] = level
    }
  }
  const filtered = new Float32Array(levels.length)
  // 5x5 max closes blocks that are mostly ink, then a 5x5 mean smooths steps.
  for (let y = 0; y < bh; y += 1) {
    for (let x = 0; x < bw; x += 1) {
      let maximum = 0
      for (let dy = -2; dy <= 2; dy += 1) {
        const yy = y + dy
        if (yy < 0 || yy >= bh) continue
        for (let dx = -2; dx <= 2; dx += 1) {
          const xx = x + dx
          if (xx < 0 || xx >= bw) continue
          const value = levels[yy * bw + xx]
          if (value > maximum) maximum = value
        }
      }
      filtered[y * bw + x] = maximum
    }
  }
  for (let y = 0; y < bh; y += 1) {
    for (let x = 0; x < bw; x += 1) {
      let sum = 0
      let count = 0
      for (let dy = -2; dy <= 2; dy += 1) {
        const yy = y + dy
        if (yy < 0 || yy >= bh) continue
        for (let dx = -2; dx <= 2; dx += 1) {
          const xx = x + dx
          if (xx < 0 || xx >= bw) continue
          sum += filtered[yy * bw + xx]
          count += 1
        }
      }
      levels[y * bw + x] = Math.max(16, sum / count)
    }
  }
  // Bilinear upsampling with block centres as samples.
  const column0 = new Int32Array(w)
  const column1 = new Int32Array(w)
  const columnT = new Float32Array(w)
  for (let x = 0; x < w; x += 1) {
    const f = Math.min(bw - 1, Math.max(0, (x + 0.5) / block - 0.5))
    const i = Math.floor(f)
    column0[x] = i
    column1[x] = Math.min(bw - 1, i + 1)
    columnT[x] = f - i
  }
  const out = new Uint8Array(w * h)
  const rowLevels = new Float32Array(w)
  for (let y = 0; y < h; y += 1) {
    const f = Math.min(bh - 1, Math.max(0, (y + 0.5) / block - 0.5))
    const r0 = Math.floor(f)
    const r1 = Math.min(bh - 1, r0 + 1)
    const ty = f - r0
    const top = r0 * bw
    const bottom = r1 * bw
    for (let x = 0; x < w; x += 1) {
      const tx = columnT[x]
      const a = levels[top + column0[x]] + (levels[top + column1[x]] - levels[top + column0[x]]) * tx
      const b = levels[bottom + column0[x]] + (levels[bottom + column1[x]] - levels[bottom + column0[x]]) * tx
      rowLevels[x] = 255 / (a + (b - a) * ty)
    }
    const row = y * w
    for (let x = 0; x < w; x += 1) {
      const value = data[row + x] * rowLevels[x] + 0.5
      out[row + x] = value >= 255 ? 255 : value
    }
  }
  return { width: w, height: h, data: out }
}

/**
 * Ink and paper levels: the 0.5th and 60th percentiles. On sparse pages, where
 * the darkest 0.5% is still paper, the ink level is the 10th percentile of the
 * pixels clearly darker than paper. `ink` is null when nothing looks like ink.
 */
export function measureInkLevels(image, options = {}) {
  const hist = histogram(image)
  const total = image.width * image.height
  const paper = percentile(hist, total, options.high ?? 0.6)
  let ink = percentile(hist, total, options.low ?? 0.005)
  if (paper - ink < 48) {
    const cutoff = paper - 24
    let inkCount = 0
    for (let value = 0; value < Math.max(0, cutoff); value += 1) inkCount += hist[value]
    if (cutoff <= 0 || inkCount < Math.max(16, total * 2e-6)) return { ink: null, paper }
    const target = inkCount * 0.1
    let accumulated = 0
    ink = 0
    for (let value = 0; value < cutoff; value += 1) {
      accumulated += hist[value]
      if (accumulated >= target) { ink = value; break }
    }
  }
  return { ink, paper }
}

/**
 * Map the ink level to black and the paper level to white (see
 * measureInkLevels). Returns the input unchanged when there is nothing to
 * stretch, so a nearly blank page is never turned black.
 */
export function stretchContrast(image, options = {}) {
  assertImage(image)
  const minRange = options.minRange ?? 32
  const total = image.width * image.height
  const { ink, paper } = options.levels ?? measureInkLevels(image, options)
  if (ink === null) return image
  const white = Math.min(255, Math.max(paper, ink + minRange))
  const black = Math.min(ink, white - minRange)
  if (black <= 0 && white >= 255) return image
  const lut = new Uint8Array(256)
  const scale = 255 / (white - black)
  for (let value = 0; value < 256; value += 1) lut[value] = clampByte(Math.round((value - black) * scale))
  const out = new Uint8Array(total)
  const { data } = image
  for (let i = 0; i < total; i += 1) out[i] = lut[data[i]]
  return { width: image.width, height: image.height, data: out }
}

/**
 * Robust noise sigma (grey levels) from the 4-neighbour Laplacian at background
 * pixels on a 2-px grid: sigma = 1.4826 * median(|L|) / sqrt(20). Glyph edges are
 * a small minority of samples, so the median ignores them.
 */
export function estimateNoise(image) {
  assertImage(image)
  const { width: w, height: h, data } = image
  if (w < 3 || h < 3) return 0
  const counts = new Uint32Array(2041)
  let samples = 0
  for (let y = 1; y < h - 1; y += 2) {
    let i = y * w + 1
    for (let x = 1; x < w - 1; x += 2, i += 2) {
      const c = data[i]
      const n = data[i - w]
      const s = data[i + w]
      const west = data[i - 1]
      const east = data[i + 1]
      if (c < 128 || n < 128 || s < 128 || west < 128 || east < 128) continue
      const laplacian = 4 * c - n - s - west - east
      counts[laplacian < 0 ? -laplacian : laplacian] += 1
      samples += 1
    }
  }
  if (samples < 64) return 0
  const half = samples / 2
  let accumulated = 0
  for (let value = 0; value < counts.length; value += 1) {
    accumulated += counts[value]
    if (accumulated >= half) return (1.4826 * value) / Math.sqrt(20)
  }
  return 0
}

/**
 * 3x3 median with the 19-comparator network (Paeth): removes isolated specks
 * and roughly halves grain. Like any 3x3 median it erases features thinner than
 * 2 px, which is why preparePage only runs it on measurably noisy pages.
 * Border pixels are copied.
 */
export function median3(image) {
  assertImage(image)
  const { width: w, height: h, data } = image
  const out = new Uint8Array(data.subarray(0, w * h))
  if (w < 3 || h < 3) return { width: w, height: h, data: out }
  for (let y = 1; y < h - 1; y += 1) {
    let i = y * w + 1
    for (let x = 1; x < w - 1; x += 1, i += 1) {
      let p0 = data[i - w - 1]
      let p1 = data[i - w]
      let p2 = data[i - w + 1]
      let p3 = data[i - 1]
      let p4 = data[i]
      let p5 = data[i + 1]
      let p6 = data[i + w - 1]
      let p7 = data[i + w]
      let p8 = data[i + w + 1]
      if (p0 === p4 && p1 === p4 && p2 === p4 && p3 === p4 && p5 === p4 && p6 === p4 && p7 === p4 && p8 === p4) continue
      let t
      if (p1 > p2) { t = p1; p1 = p2; p2 = t }
      if (p4 > p5) { t = p4; p4 = p5; p5 = t }
      if (p7 > p8) { t = p7; p7 = p8; p8 = t }
      if (p0 > p1) { t = p0; p0 = p1; p1 = t }
      if (p3 > p4) { t = p3; p3 = p4; p4 = t }
      if (p6 > p7) { t = p6; p6 = p7; p7 = t }
      if (p1 > p2) { t = p1; p1 = p2; p2 = t }
      if (p4 > p5) { t = p4; p4 = p5; p5 = t }
      if (p7 > p8) { t = p7; p7 = p8; p8 = t }
      if (p0 > p3) { t = p0; p0 = p3; p3 = t }
      if (p5 > p8) { t = p5; p5 = p8; p8 = t }
      if (p4 > p7) { t = p4; p4 = p7; p7 = t }
      if (p3 > p6) { t = p3; p3 = p6; p6 = t }
      if (p1 > p4) { t = p1; p1 = p4; p4 = t }
      if (p2 > p5) { t = p2; p2 = p5; p5 = t }
      if (p4 > p7) { t = p4; p4 = p7; p7 = t }
      if (p4 > p2) { t = p4; p4 = p2; p2 = t }
      if (p6 > p4) { t = p6; p6 = p4; p4 = t }
      if (p4 > p2) { t = p4; p4 = p2; p2 = t }
      out[i] = p4
    }
  }
  return { width: w, height: h, data: out }
}

/**
 * Skew of the text lines in degrees: positive when lines run downhill to the
 * right (clockwise on screen). Projection-profile variance over at most
 * `maxPoints` dark points, searched over +/-maxDegrees in 0.5 deg steps and then
 * refined in 0.05 deg steps. Pages without enough ink report 0.
 */
export function estimateSkew(image, options = {}) {
  assertImage(image)
  const { width: w, height: h, data } = image
  const maxDegrees = options.maxDegrees ?? 6
  const threshold = options.threshold ?? 128
  const maxPoints = options.maxPoints ?? 40_000
  const step = options.step ?? 2
  let dark = 0
  for (let y = 0; y < h; y += step) {
    const row = y * w
    for (let x = 0; x < w; x += step) if (data[row + x] < threshold) dark += 1
  }
  if (dark < 64) return 0
  const stride = Math.max(1, Math.ceil(dark / maxPoints))
  const capacity = Math.ceil(dark / stride)
  const xs = new Float64Array(capacity)
  const ys = new Float64Array(capacity)
  const cx = w / 2
  const cy = h / 2
  let count = 0
  let seen = 0
  for (let y = 0; y < h; y += step) {
    const row = y * w
    for (let x = 0; x < w; x += step) {
      if (data[row + x] >= threshold) continue
      if (seen % stride === 0 && count < capacity) {
        xs[count] = x + 0.5 - cx
        ys[count] = y + 0.5 - cy
        count += 1
      }
      seen += 1
    }
  }
  const radius = Math.hypot(w, h) / 2 + 2
  const binSize = 2
  const bins = new Float64Array(Math.ceil((2 * radius) / binSize) + 2)
  const score = (degrees) => {
    const angle = (degrees * Math.PI) / 180
    const sin = Math.sin(angle)
    const cos = Math.cos(angle)
    bins.fill(0)
    for (let i = 0; i < count; i += 1) bins[((ys[i] * cos - xs[i] * sin + radius) / binSize) | 0] += 1
    let sum = 0
    for (let i = 0; i < bins.length; i += 1) sum += bins[i] * bins[i]
    return sum
  }
  let best = 0
  let bestScore = -1
  for (let degrees = -maxDegrees; degrees <= maxDegrees + 1e-9; degrees += 0.5) {
    const value = score(degrees)
    if (value > bestScore) { bestScore = value; best = degrees }
  }
  const coarse = best
  for (let degrees = coarse - 0.5; degrees <= coarse + 0.5 + 1e-9; degrees += 0.05) {
    const value = score(degrees)
    if (value > bestScore) { bestScore = value; best = degrees }
  }
  return Math.round(best * 100) / 100
}

/**
 * Rotate the image content by `degrees` (positive = clockwise on screen) about
 * the image centre, bilinear, keeping the size and filling with paper white.
 * rotateGray(image, -estimateSkew(image)) levels the text lines.
 */
export function rotateGray(image, degrees, options = {}) {
  assertImage(image)
  const { width: w, height: h, data } = image
  const fill = options.fill ?? 255
  const angle = (degrees * Math.PI) / 180
  const cos = Math.cos(angle)
  const sin = Math.sin(angle)
  const cx = w / 2
  const cy = h / 2
  const out = new Uint8Array(w * h)
  for (let y = 0; y < h; y += 1) {
    const dy = y + 0.5 - cy
    // Source point q = R(-degrees) (p - c) + c, shifted to pixel-centre indices.
    let sx = cos * (0.5 - cx) + sin * dy + cx - 0.5
    let sy = -sin * (0.5 - cx) + cos * dy + cy - 0.5
    const row = y * w
    for (let x = 0; x < w; x += 1, sx += cos, sy -= sin) {
      const x0 = Math.floor(sx)
      const y0 = Math.floor(sy)
      if (x0 < -1 || y0 < -1 || x0 >= w || y0 >= h) { out[row + x] = fill; continue }
      const fx = sx - x0
      const fy = sy - y0
      const inX0 = x0 >= 0
      const inX1 = x0 + 1 < w
      const inY0 = y0 >= 0
      const inY1 = y0 + 1 < h
      const a = inX0 && inY0 ? data[y0 * w + x0] : fill
      const b = inX1 && inY0 ? data[y0 * w + x0 + 1] : fill
      const c = inX0 && inY1 ? data[(y0 + 1) * w + x0] : fill
      const d = inX1 && inY1 ? data[(y0 + 1) * w + x0 + 1] : fill
      const top = a + (b - a) * fx
      const bottom = c + (d - c) * fx
      out[row + x] = top + (bottom - top) * fy + 0.5
    }
  }
  return { width: w, height: h, data: out }
}

/**
 * Map a point of a deskewed OCR image back to the rendered page image:
 * p0 = R(+degrees) (p - c) + c, the inverse of rotateGray(image, -degrees).
 */
export function deskewPoint(x, y, deskew) {
  if (!deskew || !deskew.degrees) return { x, y }
  const angle = (deskew.degrees * Math.PI) / 180
  const cos = Math.cos(angle)
  const sin = Math.sin(angle)
  const dx = x - deskew.cx
  const dy = y - deskew.cy
  return { x: cos * dx - sin * dy + deskew.cx, y: sin * dx + cos * dy + deskew.cy }
}

/**
 * Robust sigma of the paper (pixels >= 128) from its lower tail, which survives
 * the clipping at 255 that normalisation causes: median - 15.87th percentile.
 */
export function estimateBackgroundSigma(image) {
  const hist = histogram(image)
  let paper = 0
  for (let value = 128; value < 256; value += 1) paper += hist[value]
  if (paper < 64) return 0
  const lowerTail = paper * 0.1587
  const half = paper * 0.5
  let accumulated = 0
  let tail = -1
  for (let value = 128; value < 256; value += 1) {
    accumulated += hist[value]
    if (tail < 0 && accumulated >= lowerTail) tail = value
    if (accumulated >= half) return value - tail
  }
  return 0
}

/**
 * Sauvola binarisation (window 31, k 0.2, R 128) with sliding column sums, so
 * memory stays O(width) instead of two full-size integral images. Pixels darker
 * than `inkFloor` are always ink: inside strokes wider than half the window the
 * local threshold would otherwise drop below the stroke and hollow it out.
 */
export function sauvola(image, options = {}) {
  assertImage(image)
  const { width: w, height: h, data } = image
  const radius = Math.max(1, Math.floor((options.window ?? 31) / 2))
  const k = options.k ?? 0.2
  const range = options.range ?? 128
  const inkFloor = options.inkFloor ?? 0
  const columnSum = new Float64Array(w)
  const columnSquares = new Float64Array(w)
  const out = new Uint8Array(w * h)
  const addRow = (y, sign) => {
    const row = y * w
    for (let x = 0; x < w; x += 1) {
      const value = data[row + x]
      columnSum[x] += sign * value
      columnSquares[x] += sign * value * value
    }
  }
  for (let y = 0; y <= Math.min(h - 1, radius); y += 1) addRow(y, 1)
  for (let y = 0; y < h; y += 1) {
    if (y > 0) {
      if (y + radius < h) addRow(y + radius, 1)
      if (y - radius - 1 >= 0) addRow(y - radius - 1, -1)
    }
    const rows = Math.min(h - 1, y + radius) - Math.max(0, y - radius) + 1
    let sum = 0
    let squares = 0
    for (let x = 0; x <= Math.min(w - 1, radius); x += 1) { sum += columnSum[x]; squares += columnSquares[x] }
    const row = y * w
    for (let x = 0; x < w; x += 1) {
      if (x > 0) {
        if (x + radius < w) { sum += columnSum[x + radius]; squares += columnSquares[x + radius] }
        if (x - radius - 1 >= 0) { sum -= columnSum[x - radius - 1]; squares -= columnSquares[x - radius - 1] }
      }
      const value = data[row + x]
      if (value < inkFloor) continue
      const n = rows * (Math.min(w - 1, x + radius) - Math.max(0, x - radius) + 1)
      const mean = sum / n
      const deviation = Math.sqrt(Math.max(0, squares / n - mean * mean))
      out[row + x] = value > mean * (1 + k * (deviation / range - 1)) ? 255 : 0
    }
  }
  return { width: w, height: h, data: out }
}

/** Dark pixels that belong to a stroke (at least two dark 4-neighbours); isolated specks do not count. */
export function countInkPixels(image, threshold = 128) {
  assertImage(image)
  const { width: w, height: h, data } = image
  let count = 0
  for (let y = 1; y < h - 1; y += 1) {
    let i = y * w + 1
    for (let x = 1; x < w - 1; x += 1, i += 1) {
      if (data[i] >= threshold) continue
      const neighbours = (data[i - 1] < threshold ? 1 : 0) + (data[i + 1] < threshold ? 1 : 0)
        + (data[i - w] < threshold ? 1 : 0) + (data[i + w] < threshold ? 1 : 0)
      if (neighbours >= 2) count += 1
    }
  }
  return count
}

/** Default minimum stroke pixels for a page to be worth recognising: about two 10 pt characters. */
export function minimumInkPixels(width, height) {
  return Math.max(48, Math.round(width * height * 1.5e-5))
}

export function isBlank(image, options = {}) {
  return countInkPixels(image) < (options.minInkPixels ?? minimumInkPixels(image.width, image.height))
}

/** Binary PGM (P5), the cheapest encoding Leptonica reads. */
export function encodePgm(image) {
  assertImage(image)
  const header = `P5\n${image.width} ${image.height}\n255\n`
  const out = new Uint8Array(header.length + image.width * image.height)
  for (let i = 0; i < header.length; i += 1) out[i] = header.charCodeAt(i)
  out.set(image.data.subarray(0, image.width * image.height), header.length)
  return out
}

export function decodePgm(bytes) {
  const head = String.fromCharCode(...bytes.subarray(0, Math.min(64, bytes.length)))
  const match = /^P5\s+(\d+)\s+(\d+)\s+255\s/.exec(head)
  if (!match) throw new Error('Not an 8-bit binary PGM.')
  const width = Number(match[1])
  const height = Number(match[2])
  assertDimensions(width, height)
  const data = bytes.slice(match[0].length, match[0].length + width * height)
  if (data.length < width * height) throw new Error('PGM data is truncated.')
  return { width, height, data }
}

/** Identifies everything besides the pixels that changes recognition output. */
export function engineSignature({ language = 'eng', psm = '3' } = {}) {
  return `${OCR_ENGINE_TAG}|${language}-best_int|psm${psm}|${PREP_VERSION}`
}

/**
 * Tesseract parameters for one page. Never set thresholding_method to 1: on the
 * noisy fixture it took 104 s and produced garbage. Otsu (default) after
 * normalisation scored 100% on every fixture.
 */
export function tesseractParameters({ dpi, psm = '3' }) {
  if (!Number.isFinite(dpi) || dpi < 30 || dpi > 2400) throw new Error('Invalid OCR resolution.')
  if (!['3', '4', '6', '7', '11'].includes(String(psm))) throw new Error('Unsupported page segmentation mode.')
  return {
    user_defined_dpi: String(Math.round(dpi)),
    tessedit_pageseg_mode: String(psm),
    preserve_interword_spaces: '0',
  }
}

export async function sha256Hex(bytes) {
  const subtle = globalThis.crypto && globalThis.crypto.subtle
  if (!subtle) throw new Error('SHA-256 is not available in this context.')
  const digest = new Uint8Array(await subtle.digest('SHA-256', bytes))
  let hex = ''
  for (let i = 0; i < digest.length; i += 1) hex += digest[i].toString(16).padStart(2, '0')
  return hex
}

const now = () => (globalThis.performance && typeof globalThis.performance.now === 'function' ? globalThis.performance.now() : Date.now())

/**
 * The full OCR input pipeline: grey -> background normalisation -> contrast
 * stretch -> 3x3 median when noisy -> deskew when |skew| >= 1.5 deg -> Sauvola
 * when the ink is faint (low contrast) or the paper is still heavily textured
 * -> blank test. Otherwise Tesseract's own Otsu threshold is used.
 *
 * Binarisation is for OCR input only; the page image itself is never touched.
 * `deskew`, when present, maps OCR-image points back with deskewPoint().
 */
export function preparePage(pixels, width, height, options = {}) {
  const timings = {}
  let started = now()
  const lap = (name) => {
    const time = now()
    timings[name] = Math.round((time - started) * 10) / 10
    started = time
  }
  let image = toGray(pixels, width, height, options.channels ?? 4)
  lap('gray')
  image = normalizeBackground(image)
  lap('normalize')
  const levels = measureInkLevels(image)
  const contrast = levels.ink === null ? 0 : levels.paper - levels.ink
  image = stretchContrast(image, { levels })
  lap('stretch')
  const noiseSigma = estimateNoise(image)
  const denoised = noiseSigma > NOISE_SIGMA_THRESHOLD
  if (denoised) image = median3(image)
  lap('denoise')
  const skewDegrees = estimateSkew(image)
  let deskew = null
  if (Math.abs(skewDegrees) >= DESKEW_MIN_DEGREES) {
    image = rotateGray(image, -skewDegrees)
    deskew = { degrees: skewDegrees, cx: width / 2, cy: height / 2 }
  }
  lap('deskew')
  const backgroundSigma = estimateBackgroundSigma(image)
  const lowContrast = levels.ink !== null && contrast < LOW_CONTRAST_LEVELS
  const binarized = lowContrast || backgroundSigma > SAUVOLA_BACKGROUND_SIGMA
  if (binarized) image = sauvola(image, { inkFloor: SAUVOLA_INK_FLOOR })
  lap('binarize')
  const inkPixels = countInkPixels(image)
  const blank = inkPixels < (options.minInkPixels ?? minimumInkPixels(width, height))
  lap('blank')
  return {
    image,
    blank,
    deskew,
    stats: {
      dpi: options.dpi,
      inkLevel: levels.ink,
      paperLevel: levels.paper,
      contrast,
      skewDegrees,
      noiseSigma: Math.round(noiseSigma * 100) / 100,
      denoised,
      backgroundSigma,
      binarized,
      inkPixels,
      timings,
    },
  }
}
