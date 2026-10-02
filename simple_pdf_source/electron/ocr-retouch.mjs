// Retouch ("inpaint") the scanned glyphs of one text line, design section
// 4.8.3 (WP4, "Edit scanned text"). Pure: typed arrays in, typed arrays out,
// so the renderer's preparation worker and the Node tests run the same code.
//
// The result is a patch image whose alpha covers only the old glyph pixels
// (plus a thin feather), filled with the surrounding paper and its grain. The
// scan image itself is never touched, so every pixel outside the patch stays
// exactly as it was. This is retouching, not redaction: the original pixels
// remain in the file underneath the patch.

export const RETOUCH_VERSION = 'retouch1'

/** Tunable constants (fractions are of the region's paper-to-ink contrast or of the em size). */
export const RETOUCH = Object.freeze({
  /** Paper level: this quantile of the region's grey values. */
  paperQuantile: 0.7,
  /** Ink level: this quantile. */
  inkQuantile: 0.02,
  /** A pixel darker than paper - threshold * (paper - ink) is ink. */
  inkThreshold: 0.25,
  /** Below this paper-to-ink contrast (grey levels) the region holds no ink worth removing. */
  minContrast: 24,
  /** Target word boxes grow by this share of the x-height before they restrict the mask. */
  quadGrowth: 0.12,
  /** The ink mask grows by this many points (3 px at 300 DPI) to swallow antialiased edges. */
  dilatePoints: 0.72,
  /** Smallest dilation in pixels. */
  minDilate: 2,
  /** The line band, in em above (ascent) and below (descent) the baseline. */
  bandAscent: 1.02,
  bandDescent: 0.36,
  /** Bilevel scan: this share of pixels within `bilevelSlack` levels of black or white. */
  bilevelShare: 0.97,
  bilevelSlack: 10,
  /** Word segmentation without recognised text: a gap wider than this share of the x-height separates words. */
  wordGap: 0.33,
  /** Feather (one pixel) alpha. */
  rampAlpha: 128,
})

const clampByte = (value) => (value <= 0 ? 0 : value >= 255 ? 255 : Math.round(value))

/** Grey (BT.601 luma) of 1-, 3- or 4-channel pixels. */
export function grayOf(pixels, width, height, channels = 4) {
  const count = width * height
  if (!(count > 0) || pixels.length < count * channels) throw new Error('retouch: the pixel buffer does not match its size.')
  const gray = new Uint8Array(count)
  if (channels === 1) {
    gray.set(pixels.subarray ? pixels.subarray(0, count) : Array.prototype.slice.call(pixels, 0, count))
    return gray
  }
  for (let index = 0, offset = 0; index < count; index += 1, offset += channels) {
    gray[index] = (pixels[offset] * 299 + pixels[offset + 1] * 587 + pixels[offset + 2] * 114 + 500) / 1000
  }
  return gray
}

/** RGBA copy (alpha 255) of 1-, 3- or 4-channel pixels. */
export function rgbaOf(pixels, width, height, channels = 4) {
  const count = width * height
  const out = new Uint8ClampedArray(count * 4)
  for (let index = 0, offset = 0; index < count; index += 1, offset += channels) {
    const o = index * 4
    if (channels === 1) {
      out[o] = out[o + 1] = out[o + 2] = pixels[offset]
    } else {
      out[o] = pixels[offset]
      out[o + 1] = pixels[offset + 1]
      out[o + 2] = pixels[offset + 2]
    }
    out[o + 3] = 255
  }
  return out
}

export function grayHistogram(gray) {
  const hist = new Uint32Array(256)
  for (let index = 0; index < gray.length; index += 1) hist[gray[index]] += 1
  return hist
}

export function histogramQuantile(hist, total, fraction) {
  const target = Math.max(0, Math.min(total - 1, Math.floor(total * fraction)))
  let seen = 0
  for (let level = 0; level < 256; level += 1) {
    seen += hist[level]
    if (seen > target) return level
  }
  return 255
}

/** Paper and ink levels of a region (design: p70 and p2 of its grey values). */
export function estimateLevels(gray) {
  const hist = grayHistogram(gray)
  const paper = histogramQuantile(hist, gray.length, RETOUCH.paperQuantile)
  const ink = histogramQuantile(hist, gray.length, RETOUCH.inkQuantile)
  const slack = RETOUCH.bilevelSlack
  let extremes = 0
  for (let level = 0; level <= slack; level += 1) extremes += hist[level] + hist[255 - level]
  return { paper, ink, contrast: paper - ink, bilevel: extremes >= gray.length * RETOUCH.bilevelShare, hist }
}

/**
 * The paper level around every pixel (stains, shading and tinted paper vary
 * across a line): the 90th percentile of `block`-pixel blocks, closed with a
 * 3x3 max (blocks that are mostly ink), smoothed with a 3x3 mean and
 * upsampled bilinearly. Returns a Float32Array of grey levels.
 */
export function localPaper(gray, width, height, block = 16) {
  const size = Math.max(4, Math.round(block))
  const bw = Math.ceil(width / size)
  const bh = Math.ceil(height / size)
  const levels = new Float32Array(bw * bh)
  const hist = new Uint32Array(256)
  for (let by = 0; by < bh; by += 1) {
    for (let bx = 0; bx < bw; bx += 1) {
      hist.fill(0)
      const y1 = Math.min(height, (by + 1) * size)
      const x1 = Math.min(width, (bx + 1) * size)
      for (let y = by * size; y < y1; y += 1) for (let x = bx * size; x < x1; x += 1) hist[gray[y * width + x]] += 1
      const skip = (y1 - by * size) * (x1 - bx * size) * 0.1
      let seen = 0
      let level = 255
      for (let value = 255; value >= 0; value -= 1) {
        seen += hist[value]
        if (seen > skip) { level = value; break }
      }
      levels[by * bw + bx] = level
    }
  }
  const pass = (source, reduce) => {
    const out = new Float32Array(source.length)
    for (let y = 0; y < bh; y += 1) {
      for (let x = 0; x < bw; x += 1) {
        const values = []
        for (let dy = -1; dy <= 1; dy += 1) {
          for (let dx = -1; dx <= 1; dx += 1) {
            const yy = y + dy
            const xx = x + dx
            if (yy >= 0 && yy < bh && xx >= 0 && xx < bw) values.push(source[yy * bw + xx])
          }
        }
        out[y * bw + x] = reduce(values)
      }
    }
    return out
  }
  const closed = pass(levels, (values) => Math.max(...values))
  const smooth = pass(closed, (values) => values.reduce((sum, value) => sum + value, 0) / values.length)
  const out = new Float32Array(width * height)
  for (let y = 0; y < height; y += 1) {
    const fy = Math.min(bh - 1, Math.max(0, (y + 0.5) / size - 0.5))
    const r0 = Math.floor(fy)
    const r1 = Math.min(bh - 1, r0 + 1)
    const ty = fy - r0
    for (let x = 0; x < width; x += 1) {
      const fx = Math.min(bw - 1, Math.max(0, (x + 0.5) / size - 0.5))
      const c0 = Math.floor(fx)
      const c1 = Math.min(bw - 1, c0 + 1)
      const tx = fx - c0
      const top = smooth[r0 * bw + c0] + (smooth[r0 * bw + c1] - smooth[r0 * bw + c0]) * tx
      const bottom = smooth[r1 * bw + c0] + (smooth[r1 * bw + c1] - smooth[r1 * bw + c0]) * tx
      out[y * width + x] = top + (bottom - top) * ty
    }
  }
  return out
}

/**
 * Position of pixel centres in a text line's frame: `u` along the baseline
 * from its origin, `v` up from the baseline (towards ascenders), in pixels.
 * `dx, dy` is the unit reading direction in pixel space (y down).
 */
export function lineFrame(baseline) {
  const length = Math.hypot(baseline.dx, baseline.dy) || 1
  const dx = baseline.dx / length
  const dy = baseline.dy / length
  const bx = baseline.x
  const by = baseline.y
  return {
    dx,
    dy,
    u: (x, y) => (x + 0.5 - bx) * dx + (y + 0.5 - by) * dy,
    v: (x, y) => (x + 0.5 - bx) * dy - (y + 0.5 - by) * dx,
    /** Pixel position of frame point (u, v). */
    point: (u, v) => ({ x: bx + u * dx + v * dy, y: by + u * dy - v * dx }),
  }
}

/**
 * Connected components (8-neighbour) of a 0/1 mask, two passes with union-find.
 * Returns labels (0 = background, 1..count) and per-component pixel counts and boxes.
 */
export function labelComponents(mask, width, height) {
  const labels = new Int32Array(width * height)
  const parent = [0]
  const find = (label) => {
    let root = label
    while (parent[root] !== root) root = parent[root]
    while (parent[label] !== root) {
      const next = parent[label]
      parent[label] = root
      label = next
    }
    return root
  }
  const union = (a, b) => {
    const ra = find(a)
    const rb = find(b)
    if (ra !== rb) parent[Math.max(ra, rb)] = Math.min(ra, rb)
    return Math.min(ra, rb)
  }
  for (let y = 0; y < height; y += 1) {
    const row = y * width
    for (let x = 0; x < width; x += 1) {
      const index = row + x
      if (!mask[index]) continue
      let label = 0
      if (x > 0 && labels[index - 1]) label = labels[index - 1]
      if (y > 0) {
        const up = row - width + x
        for (const neighbour of [x > 0 ? up - 1 : -1, up, x + 1 < width ? up + 1 : -1]) {
          if (neighbour < 0 || !labels[neighbour]) continue
          label = label ? union(label, labels[neighbour]) : labels[neighbour]
        }
      }
      if (!label) {
        label = parent.length
        parent.push(label)
      }
      labels[index] = label
    }
  }
  // Flatten to consecutive ids.
  const remap = new Int32Array(parent.length)
  let count = 0
  for (let label = 1; label < parent.length; label += 1) {
    const root = find(label)
    if (!remap[root]) remap[root] = ++count
    remap[label] = remap[root]
  }
  const area = new Int32Array(count + 1)
  const minX = new Int32Array(count + 1).fill(width)
  const minY = new Int32Array(count + 1).fill(height)
  const maxX = new Int32Array(count + 1).fill(-1)
  const maxY = new Int32Array(count + 1).fill(-1)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const index = y * width + x
      if (!labels[index]) continue
      const id = remap[labels[index]]
      labels[index] = id
      area[id] += 1
      if (x < minX[id]) minX[id] = x
      if (x > maxX[id]) maxX[id] = x
      if (y < minY[id]) minY[id] = y
      if (y > maxY[id]) maxY[id] = y
    }
  }
  return { labels, count, area, minX, minY, maxX, maxY }
}

/** Disk offsets of radius r (pixels), as [dx, dy, distance²]. */
function diskOffsets(radius) {
  const offsets = []
  const limit = radius * radius + 0.5
  for (let dy = -radius; dy <= radius; dy += 1) {
    for (let dx = -radius; dx <= radius; dx += 1) {
      const distance = dx * dx + dy * dy
      if (distance <= limit) offsets.push([dx, dy, distance])
    }
  }
  return offsets
}

/** Grow a 0/1 mask by a disk of `radius` pixels. */
export function dilateMask(mask, width, height, radius) {
  if (radius <= 0) return Uint8Array.from(mask)
  const out = new Uint8Array(width * height)
  const offsets = diskOffsets(radius)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (!mask[y * width + x]) continue
      for (const [dx, dy] of offsets) {
        const xx = x + dx
        const yy = y + dy
        if (xx >= 0 && yy >= 0 && xx < width && yy < height) out[yy * width + xx] = 1
      }
    }
  }
  return out
}

/**
 * Fill the unknown pixels of one channel from the known ones: a pyramid of
 * masked means, each level filled from the next coarser one by bilinear
 * upsampling (push-pull). Known pixels keep their value.
 */
export function pushPullFill(values, known, width, height) {
  const sum0 = new Float32Array(width * height)
  const weight0 = new Float32Array(width * height)
  for (let index = 0; index < sum0.length; index += 1) {
    if (!known[index]) continue
    sum0[index] = values[index]
    weight0[index] = 1
  }
  const levels = [{ width, height, sum: sum0, weight: weight0 }]
  while (levels.at(-1).width > 1 || levels.at(-1).height > 1) {
    const fine = levels.at(-1)
    const w2 = Math.ceil(fine.width / 2)
    const h2 = Math.ceil(fine.height / 2)
    const sum = new Float32Array(w2 * h2)
    const weight = new Float32Array(w2 * h2)
    for (let y = 0; y < fine.height; y += 1) {
      const row = (y >> 1) * w2
      for (let x = 0; x < fine.width; x += 1) {
        const source = y * fine.width + x
        const target = row + (x >> 1)
        sum[target] += fine.sum[source]
        weight[target] += fine.weight[source]
      }
    }
    levels.push({ width: w2, height: h2, sum, weight })
  }
  // Pull: estimate every cell from the coarser level where it has no weight;
  // cells with partial weight blend their own mean with the coarser estimate.
  let estimate = null
  for (let level = levels.length - 1; level >= 0; level -= 1) {
    const { width: w, height: h, sum, weight } = levels[level]
    const cellArea = 4 ** level
    const next = new Float32Array(w * h)
    for (let y = 0; y < h; y += 1) {
      for (let x = 0; x < w; x += 1) {
        const index = y * w + x
        const own = weight[index] > 0 ? sum[index] / weight[index] : 0
        if (!estimate) { next[index] = own; continue }
        const coarse = levels[level + 1]
        const fx = Math.min(coarse.width - 1, Math.max(0, (x - 0.5) / 2))
        const fy = Math.min(coarse.height - 1, Math.max(0, (y - 0.5) / 2))
        const x0 = Math.floor(fx)
        const y0 = Math.floor(fy)
        const x1 = Math.min(coarse.width - 1, x0 + 1)
        const y1 = Math.min(coarse.height - 1, y0 + 1)
        const tx = fx - x0
        const ty = fy - y0
        const up = estimate[y0 * coarse.width + x0] * (1 - tx) * (1 - ty) + estimate[y0 * coarse.width + x1] * tx * (1 - ty)
          + estimate[y1 * coarse.width + x0] * (1 - tx) * ty + estimate[y1 * coarse.width + x1] * tx * ty
        const share = level === 0 ? (weight[index] > 0 ? 1 : 0) : Math.min(1, weight[index] / cellArea)
        next[index] = share * own + (1 - share) * up
      }
    }
    estimate = next
  }
  const out = Float32Array.from(values)
  for (let index = 0; index < out.length; index += 1) if (!known[index]) out[index] = estimate[index]
  return out
}

/** 32-bit hash of a string or number (FNV-1a), the grain seed. */
export function seedOf(seed) {
  const text = String(seed ?? '')
  let hash = 0x811c9dc5
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index)
    hash = Math.imul(hash, 0x01000193) >>> 0
  }
  return hash >>> 0
}

/** Deterministic PRNG (mulberry32): the same key always gives the same grain. */
export function createRandom(seed) {
  let state = seedOf(seed) || 1
  return () => {
    state = (state + 0x6d2b79f5) >>> 0
    let t = state
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/**
 * The paper's grain as it really is: the residuals (pixel - 5x5 mean) of
 * known pixels whose whole window is known, kept per pixel for all three
 * channels. Drawing from them reproduces the texture, its colour and its
 * skew (paper near white is clipped at 255, so its noise is lopsided), which
 * symmetric synthetic noise would not. Returns null without enough paper.
 */
export function grainModel(rgba, known, width, height, limit = 40_000) {
  const stride = width + 1
  const sums = [0, 1, 2].map(() => new Float64Array(stride * (height + 1)))
  const count = new Int32Array(stride * (height + 1))
  for (let y = 0; y < height; y += 1) {
    const rowSums = [0, 0, 0]
    let rowCount = 0
    for (let x = 0; x < width; x += 1) {
      const index = y * width + x
      if (known[index]) {
        for (let channel = 0; channel < 3; channel += 1) rowSums[channel] += rgba[index * 4 + channel]
        rowCount += 1
      }
      const cell = (y + 1) * stride + x + 1
      for (let channel = 0; channel < 3; channel += 1) sums[channel][cell] = sums[channel][y * stride + x + 1] + rowSums[channel]
      count[cell] = count[y * stride + x + 1] + rowCount
    }
  }
  const residuals = []
  for (let y = 2; y < height - 2; y += 1) {
    for (let x = 2; x < width - 2; x += 1) {
      const index = y * width + x
      if (!known[index]) continue
      const a = (y - 2) * stride + x - 2
      const b = (y - 2) * stride + x + 3
      const c = (y + 3) * stride + x - 2
      const d = (y + 3) * stride + x + 3
      if (count[d] - count[b] - count[c] + count[a] !== 25) continue
      for (let channel = 0; channel < 3; channel += 1) {
        const total = sums[channel]
        residuals.push(rgba[index * 4 + channel] - (total[d] - total[b] - total[c] + total[a]) / 25)
      }
    }
  }
  const samples = residuals.length / 3
  if (samples < 16) return null
  // Keep an even subsample of a large region.
  const step = Math.max(1, Math.floor(samples / limit))
  const kept = new Float32Array(Math.ceil(samples / step) * 3)
  let used = 0
  for (let sample = 0; sample < samples; sample += step) {
    kept[used * 3] = residuals[sample * 3]
    kept[used * 3 + 1] = residuals[sample * 3 + 1]
    kept[used * 3 + 2] = residuals[sample * 3 + 2]
    used += 1
  }
  // A residual against a mean that includes its own pixel is smaller by sqrt(24/25).
  const gain = 1 / Math.sqrt(24 / 25)
  const sigma = [0, 1, 2].map((channel) => {
    const magnitudes = new Float32Array(used)
    for (let index = 0; index < used; index += 1) magnitudes[index] = Math.abs(kept[index * 3 + channel])
    magnitudes.sort()
    return 1.4826 * magnitudes[used >> 1] * gain
  })
  return {
    samples: used,
    sigma,
    sample(random, out) {
      const at = Math.min(used - 1, Math.floor(random() * used)) * 3
      out[0] = kept[at] * gain
      out[1] = kept[at + 1] * gain
      out[2] = kept[at + 2] * gain
      return out
    },
  }
}

/** Point-in-quad test for a word box grown by `grow` pixels. Quads are [tl, tr, br, bl] as {x, y}. */
function quadTester(quad, grow) {
  const [tl, tr, , bl] = quad
  const ex = { x: tr.x - tl.x, y: tr.y - tl.y }
  const ey = { x: bl.x - tl.x, y: bl.y - tl.y }
  const lx = Math.hypot(ex.x, ex.y) || 1
  const ly = Math.hypot(ey.x, ey.y) || 1
  const mx = grow / lx
  const my = grow / ly
  const xs = quad.map((corner) => corner.x)
  const ys = quad.map((corner) => corner.y)
  return {
    box: { x0: Math.min(...xs) - grow, y0: Math.min(...ys) - grow, x1: Math.max(...xs) + grow, y1: Math.max(...ys) + grow },
    contains(px, py) {
      const rx = px - tl.x
      const ry = py - tl.y
      const s = (rx * ex.x + ry * ex.y) / (lx * lx)
      const t = (rx * ey.x + ry * ey.y) / (ly * ly)
      return s >= -mx && s <= 1 + mx && t >= -my && t <= 1 + my
    },
  }
}

/**
 * Group a line's glyph components into words: components sorted along the
 * baseline join while the gap to the next one is at most `gap` pixels.
 * `items` are { u0, u1 }; returns clusters { u0, u1, members: index[] }.
 */
export function clusterWords(items, gap) {
  const order = items.map((item, index) => ({ ...item, index })).sort((a, b) => a.u0 - b.u0)
  const clusters = []
  for (const item of order) {
    const last = clusters.at(-1)
    if (last && item.u0 - last.u1 <= gap) {
      last.u1 = Math.max(last.u1, item.u1)
      last.members.push(item.index)
    } else {
      clusters.push({ u0: item.u0, u1: item.u1, members: [item.index] })
    }
  }
  return clusters
}

// Advance widths (1/1000 em) of printable ASCII (32-126) in the three
// classic faces, used only to tell how wide each recognised word should be.
const ADVANCES = {
  sans: '278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 278 278 584 584 584 556 1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 556 556 333 500 278 556 500 722 500 500 500 334 260 334 584',
  serif: '250 333 408 500 500 833 778 180 333 333 500 564 250 333 250 278 500 500 500 500 500 500 500 500 500 500 278 278 564 564 564 444 921 722 667 667 722 611 556 722 722 333 389 722 611 889 722 722 556 722 667 556 611 722 722 944 722 722 611 333 278 333 469 500 333 444 500 444 500 444 333 500 500 278 278 500 278 778 500 500 500 500 333 389 278 500 500 722 500 500 444 480 200 480 541',
}
const advanceTables = Object.fromEntries(Object.entries(ADVANCES).map(([key, list]) => [key, list.split(' ').map(Number)]))

/** Expected ink width of a word in em (advance widths, less the outer side bearings). */
export function expectedWordWidth(token, fontClass = 'sans') {
  const table = advanceTables[fontClass]
  let width = 0
  for (const character of String(token)) {
    const code = character.codePointAt(0)
    width += fontClass === 'mono' ? 600 : code >= 32 && code <= 126 && table ? table[code - 32] : fontClass === 'serif' ? 500 : 556
  }
  return Math.max(0.1, width / 1000 - 0.06)
}

/**
 * Match glyph clusters (sorted along the line, { u0, u1 }) to the recognised
 * words: every word takes a run of consecutive clusters, chosen so that each
 * run is as wide as its word should be (dynamic programming over the split
 * points). `em` is the em size in pixels. Returns, per token, its clusters
 * and extent, or null when the two cannot be matched (fewer clusters than
 * words, or widths that do not fit).
 */
export function alignClustersToTokens(clusters, tokens, options = {}) {
  const n = tokens.length
  const m = clusters.length
  if (!n || m < n) return null
  const em = Math.max(1, Number(options.em) || 1)
  const fontClass = options.fontClass || 'sans'
  const expected = tokens.map((token) => expectedWordWidth(token, fontClass))
  const space = fontClass === 'mono' ? 0.6 : fontClass === 'serif' ? 0.25 : 0.278
  // Scale the expectations to this line's length (fonts differ in width).
  const scale = (clusters[m - 1].u1 - clusters[0].u0) / em / (expected.reduce((sum, value) => sum + value, 0) + space * (n - 1))
  const target = expected.map((value) => value * scale * em)
  const gapAfter = (index) => clusters[index + 1].u0 - clusters[index].u1
  // cost[i][j]: tokens 0..i-1 over clusters 0..j-1.
  const cost = Array.from({ length: n + 1 }, () => new Float64Array(m + 1).fill(Infinity))
  const from = Array.from({ length: n + 1 }, () => new Int32Array(m + 1).fill(-1))
  cost[0][0] = 0
  for (let i = 1; i <= n; i += 1) {
    for (let j = i; j <= m - (n - i); j += 1) {
      for (let k = i - 1; k < j; k += 1) {
        if (!Number.isFinite(cost[i - 1][k])) continue
        const width = clusters[j - 1].u1 - clusters[k].u0
        const error = (width - target[i - 1]) / em
        // A split inside a narrow gap is unlikely to be a word space.
        const gap = j < m ? gapAfter(j - 1) / em : 1
        const split = i < n ? 0.25 * Math.max(0, 0.22 - gap) / 0.22 : 0
        const value = cost[i - 1][k] + error * error + split
        if (value < cost[i][j]) { cost[i][j] = value; from[i][j] = k }
      }
    }
  }
  if (!Number.isFinite(cost[n][m])) return null
  const groups = []
  for (let i = n, j = m; i > 0; i -= 1) {
    const k = from[i][j]
    groups.unshift({ u0: clusters[k].u0, u1: Math.max(...clusters.slice(k, j).map((cluster) => cluster.u1)), clusters: Array.from({ length: j - k }, (_, offset) => k + offset) })
    j = k
  }
  // Plausible only when the words came out about as wide as expected.
  for (let index = 0; index < n; index += 1) {
    const width = groups[index].u1 - groups[index].u0
    const expectedWidth = target[index]
    if (Math.abs(width - expectedWidth) > Math.max(0.6 * em, 0.45 * expectedWidth)) return null
  }
  return groups
}

/** Whitespace-separated tokens of a line of text. */
export function textTokens(text) {
  return String(text ?? '').trim().split(/\s+/u).filter(Boolean)
}

/**
 * Retouch one scanned text line in a rendered region.
 *
 * `pixels` are the region's RGBA (or `channels`-channel) pixels. `baseline`
 * gives the line's baseline origin and unit reading direction in region
 * pixels (y down), `length` its extent along the baseline and `fontSize` its
 * em size (all in pixels). `targets` are the word boxes to remove, as quads
 * [top-left, top-right, bottom-right, bottom-left]; without them the whole
 * line is the target and, with `segment`, its words are found from the
 * pixels (`text` gives the recognised words to match, `fontClass` how wide
 * they should be). `seed` makes the grain reproducible.
 *
 * Returns everything composePatch() needs to build a patch for any subset of
 * the words, the words' ink extents (`words`, each with the mask labels it
 * owns) and diagnostics. `labels` is -1 outside the retouch mask.
 */
export function retouchLine(pixels, width, height, options = {}) {
  const channels = options.channels ?? 4
  const count = width * height
  const gray = grayOf(pixels, width, height, channels)
  const rgba = rgbaOf(pixels, width, height, channels)
  const dpi = Number(options.dpi) > 0 ? Number(options.dpi) : 300
  const fontSize = Math.max(4, Number(options.fontSize) || 0.15 * dpi)
  const xHeight = Math.max(2, Number(options.xHeight) || fontSize * 0.5)
  const length = Math.max(1, Number(options.length) || width)
  const frame = lineFrame(options.baseline ?? { x: 0, y: height * 0.75, dx: 1, dy: 0 })
  const levels = estimateLevels(gray)
  const empty = (reason) => ({
    width, height, channels: 4, bilevel: levels.bilevel, paper: levels.paper, ink: levels.ink, reason,
    filled: rgba, labels: new Int16Array(count).fill(-1), blocked: new Uint8Array(count), words: [], inkBox: null,
    maskPixels: 0, noiseSigma: [0, 0, 0], paperColor: [levels.paper, levels.paper, levels.paper], segmented: false,
  })
  if (levels.contrast < RETOUCH.minContrast) return empty('no-ink')

  const bilevel = levels.bilevel
  // Ink is darker than the paper around it (not than the region's paper):
  // stains and shading must not turn into ink.
  const paperAt = bilevel ? null : localPaper(gray, width, height, Math.max(12, 0.45 * fontSize))
  const inkAll = new Uint8Array(count)
  for (let index = 0; index < count; index += 1) {
    const paper = paperAt ? Math.max(paperAt[index], levels.ink + RETOUCH.minContrast) : 255
    const threshold = bilevel ? 128 : paper - RETOUCH.inkThreshold * (paper - levels.ink)
    inkAll[index] = gray[index] < threshold ? 1 : 0
  }

  // Where this line's glyphs may be: the band around its baseline.
  const bandTop = RETOUCH.bandAscent * fontSize
  const bandBottom = -RETOUCH.bandDescent * fontSize
  const margin = 0.6 * fontSize
  const inBand = (x, y) => {
    const v = frame.v(x, y)
    const u = frame.u(x, y)
    return v <= bandTop && v >= bandBottom && u >= -margin && u <= length + margin
  }

  // Target area: the word boxes (grown), or the whole line band.
  const grow = RETOUCH.quadGrowth * xHeight
  const targets = Array.isArray(options.targets) && options.targets.length
    ? options.targets.map((target) => quadTester(target.quad ?? target, grow))
    : null
  const targetOf = new Int16Array(count).fill(-1)
  if (targets) {
    targets.forEach((target, label) => {
      const x0 = Math.max(0, Math.floor(target.box.x0))
      const y0 = Math.max(0, Math.floor(target.box.y0))
      const x1 = Math.min(width, Math.ceil(target.box.x1) + 1)
      const y1 = Math.min(height, Math.ceil(target.box.y1) + 1)
      for (let y = y0; y < y1; y += 1) {
        for (let x = x0; x < x1; x += 1) {
          const index = y * width + x
          if (targetOf[index] >= 0 || !target.contains(x + 0.5, y + 0.5)) continue
          targetOf[index] = label
        }
      }
    })
  } else {
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) if (inBand(x, y)) targetOf[y * width + x] = 0
    }
  }

  // Components of all ink: one that lies mostly outside this line's band
  // belongs to a neighbouring line (a descender above, an ascender below) and
  // is protected, even where it reaches into a target box.
  const components = labelComponents(inkAll, width, height)
  const inside = new Int32Array(components.count + 1)
  const uMin = new Float64Array(components.count + 1).fill(Infinity)
  const uMax = new Float64Array(components.count + 1).fill(-Infinity)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const id = components.labels[y * width + x]
      if (!id) continue
      if (inBand(x, y)) inside[id] += 1
      const u = frame.u(x, y)
      if (u < uMin[id]) uMin[id] = u
      if (u > uMax[id]) uMax[id] = u
    }
  }
  const lineComponent = new Uint8Array(components.count + 1)
  for (let id = 1; id <= components.count; id += 1) lineComponent[id] = inside[id] * 2 > components.area[id] ? 1 : 0

  const target = new Uint8Array(count)
  const labels = new Int16Array(count).fill(-1)
  for (let index = 0; index < count; index += 1) {
    if (!inkAll[index] || targetOf[index] < 0) continue
    if (!lineComponent[components.labels[index]]) continue
    target[index] = 1
  }

  // Words: the target boxes, or clusters of the line's components.
  let words = []
  let segmented = false
  if (targets) {
    for (let index = 0; index < count; index += 1) if (target[index]) labels[index] = targetOf[index]
    words = targets.map((_, label) => ({ label, labels: [label] }))
  } else {
    const minArea = Math.max(4, 0.012 * xHeight * xHeight)
    const used = []
    for (let id = 1; id <= components.count; id += 1) {
      if (!lineComponent[id] || components.area[id] < minArea) continue
      used.push({ u0: uMin[id], u1: uMax[id] + 1, id })
    }
    // Glyph clusters: components that overlap along the line (an i and its
    // dot, accents, touching letters) form one glyph.
    const glyphs = clusterWords(used, 0.5)
    const tokens = options.segment ? textTokens(options.text) : []
    const groups = tokens.length ? alignClustersToTokens(glyphs, tokens, { em: fontSize, fontClass: options.fontClass }) : null
    // Without a match (or without text), words are what the gaps suggest.
    const clusters = groups
      ? groups.map((group) => ({ u0: group.u0, u1: group.u1, members: group.clusters.flatMap((index) => glyphs[index].members) }))
      : clusterWords(used, RETOUCH.wordGap * xHeight)
    const clusterOfComponent = new Int32Array(components.count + 1).fill(-1)
    clusters.forEach((cluster, clusterIndex) => {
      for (const member of cluster.members) clusterOfComponent[used[member].id] = clusterIndex
    })
    // Specks and dots too small to place a word go to the nearest cluster.
    const nearest = (u) => {
      let best = 0
      let distance = Infinity
      clusters.forEach((cluster, clusterIndex) => {
        const d = u < cluster.u0 ? cluster.u0 - u : u > cluster.u1 ? u - cluster.u1 : 0
        if (d < distance) { distance = d; best = clusterIndex }
      })
      return best
    }
    for (let id = 1; id <= components.count; id += 1) {
      if (clusterOfComponent[id] < 0 && lineComponent[id] && clusters.length) clusterOfComponent[id] = nearest((uMin[id] + uMax[id]) / 2)
    }
    for (let index = 0; index < count; index += 1) {
      if (!target[index]) continue
      const cluster = clusterOfComponent[components.labels[index]]
      labels[index] = cluster >= 0 ? cluster : 0
    }
    if (groups) {
      segmented = true
      words = groups.map((_, index) => ({ label: index, labels: [index], text: tokens[index] }))
    } else {
      words = clusters.length ? [{ label: 0, labels: clusters.map((_, index) => index) }] : []
    }
  }

  // Ink extents per word (and of everything targeted).
  const extents = new Map()
  let inkBox = null
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const index = y * width + x
      if (!target[index]) continue
      const label = labels[index]
      let box = extents.get(label)
      if (!box) extents.set(label, box = { x0: x, y0: y, x1: x + 1, y1: y + 1, u0: Infinity, u1: -Infinity })
      if (x < box.x0) box.x0 = x
      if (y < box.y0) box.y0 = y
      if (x + 1 > box.x1) box.x1 = x + 1
      if (y + 1 > box.y1) box.y1 = y + 1
      const u = frame.u(x, y)
      if (u < box.u0) box.u0 = u
      if (u + 1 > box.u1) box.u1 = u + 1
      if (!inkBox) inkBox = { x0: x, y0: y, x1: x + 1, y1: y + 1 }
      else {
        if (x < inkBox.x0) inkBox.x0 = x
        if (y < inkBox.y0) inkBox.y0 = y
        if (x + 1 > inkBox.x1) inkBox.x1 = x + 1
        if (y + 1 > inkBox.y1) inkBox.y1 = y + 1
      }
    }
  }
  words = words.map((word) => {
    let box = null
    for (const label of word.labels) {
      const own = extents.get(label)
      if (!own) continue
      box = box
        ? { x0: Math.min(box.x0, own.x0), y0: Math.min(box.y0, own.y0), x1: Math.max(box.x1, own.x1), y1: Math.max(box.y1, own.y1), u0: Math.min(box.u0, own.u0), u1: Math.max(box.u1, own.u1) }
        : { ...own }
    }
    return { ...word, box }
  })
  if (!inkBox) return { ...empty('no-target-ink'), words }

  // The mask: target ink grown over its antialiased edge, each grown pixel
  // owned by the nearest word. Ink that is not targeted (a neighbouring line,
  // a word that is kept) and a one-pixel rim around it are never touched.
  const radius = bilevel ? 1 : Math.max(RETOUCH.minDilate, Math.round(RETOUCH.dilatePoints * dpi / 72))
  const blocked = dilateMask(Uint8Array.from(inkAll, (flag, index) => (flag && !target[index] ? 1 : 0)), width, height, 1)
  const distance = new Float32Array(count).fill(Infinity)
  const offsets = diskOffsets(radius)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const index = y * width + x
      if (!target[index]) continue
      const label = labels[index]
      for (const [dx, dy, d] of offsets) {
        const xx = x + dx
        const yy = y + dy
        if (xx < 0 || yy < 0 || xx >= width || yy >= height) continue
        const other = yy * width + xx
        if (blocked[other] && !target[other]) continue
        if (target[other] && other !== index) continue
        if (d < distance[other]) { distance[other] = d; labels[other] = label }
      }
    }
  }
  let maskPixels = 0
  for (let index = 0; index < count; index += 1) {
    if (Number.isFinite(distance[index])) maskPixels += 1
    else labels[index] = -1
  }

  // Fill from the paper only: every ink pixel (and its grown edge) is unknown.
  const known = dilateMask(inkAll, width, height, radius)
  for (let index = 0; index < count; index += 1) known[index] = known[index] ? 0 : 1
  let knownCount = 0
  for (let index = 0; index < count; index += 1) knownCount += known[index]
  const paperColor = [0, 1, 2].map((channel) => {
    const hist = new Uint32Array(256)
    for (let index = 0; index < count; index += 1) if (known[index]) hist[rgba[index * 4 + channel]] += 1
    return knownCount ? histogramQuantile(hist, knownCount, 0.5) : levels.paper
  })
  const filled = new Uint8ClampedArray(rgba)
  const noiseSigma = [0, 0, 0]
  if (bilevel || knownCount < 16) {
    const paper = bilevel ? (levels.paper >= 128 ? 255 : levels.paper) : levels.paper
    for (let index = 0; index < count; index += 1) {
      if (labels[index] < 0) continue
      for (let channel = 0; channel < 3; channel += 1) filled[index * 4 + channel] = bilevel ? paper : paperColor[channel]
    }
  } else {
    const random = createRandom(options.seed ?? 'simple')
    let grey = true
    for (let index = 0; index < count && grey; index += 1) {
      const offset = index * 4
      grey = rgba[offset] === rgba[offset + 1] && rgba[offset] === rgba[offset + 2]
    }
    const fills = []
    for (let channel = 0; channel < (grey ? 1 : 3); channel += 1) {
      const values = new Float32Array(count)
      for (let index = 0; index < count; index += 1) values[index] = rgba[index * 4 + channel]
      fills.push(pushPullFill(values, known, width, height))
    }
    while (fills.length < 3) fills.push(fills[0])
    // Grain: residuals of the surrounding paper, drawn in pixel order from a
    // generator seeded by the edit, so preview and save get the same patch.
    const grain = grainModel(rgba, known, width, height)
    if (grain) for (let channel = 0; channel < 3; channel += 1) noiseSigma[channel] = grain.sigma[channel]
    const noise = [0, 0, 0]
    for (let index = 0; index < count; index += 1) {
      if (grain) grain.sample(random, noise)
      if (labels[index] < 0) continue
      for (let channel = 0; channel < 3; channel += 1) {
        filled[index * 4 + channel] = clampByte(fills[channel][index] + noise[channel])
      }
    }
  }

  return {
    width, height, channels: 4, bilevel, paper: levels.paper, ink: levels.ink,
    filled, labels, blocked, words, inkBox, maskPixels, noiseSigma, paperColor, segmented, radius,
  }
}

/**
 * The patch for some of the words (all of them when `include` is null): RGBA
 * cropped to the masked pixels, alpha 255 on them, a one-pixel feather of
 * alpha 128 around them (not on bilevel scans) and 0 everywhere else, so the
 * pixels around the old glyphs are never changed. `include` holds mask
 * labels (see retouchLine words[].labels). Returns null when nothing is masked.
 */
export function composePatch(retouch, include = null) {
  const { width, height, labels, blocked, filled, bilevel } = retouch
  const count = width * height
  const selected = new Uint8Array(count)
  let any = false
  for (let index = 0; index < count; index += 1) {
    const label = labels[index]
    if (label < 0 || (include && !include.has(label))) continue
    selected[index] = 1
    any = true
  }
  if (!any) return null
  const ramp = new Uint8Array(count)
  if (!bilevel) {
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const index = y * width + x
        if (selected[index] || labels[index] >= 0 || blocked[index]) continue
        let near = false
        for (let dy = -1; dy <= 1 && !near; dy += 1) {
          for (let dx = -1; dx <= 1 && !near; dx += 1) {
            const xx = x + dx
            const yy = y + dy
            if ((dx || dy) && xx >= 0 && yy >= 0 && xx < width && yy < height && selected[yy * width + xx]) near = true
          }
        }
        if (near) ramp[index] = 1
      }
    }
  }
  let x0 = width
  let y0 = height
  let x1 = -1
  let y1 = -1
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const index = y * width + x
      if (!selected[index] && !ramp[index]) continue
      if (x < x0) x0 = x
      if (x > x1) x1 = x
      if (y < y0) y0 = y
      if (y > y1) y1 = y
    }
  }
  const patchWidth = x1 - x0 + 1
  const patchHeight = y1 - y0 + 1
  const out = new Uint8ClampedArray(patchWidth * patchHeight * 4)
  const paper = retouch.paperColor ?? [255, 255, 255]
  for (let y = 0; y < patchHeight; y += 1) {
    for (let x = 0; x < patchWidth; x += 1) {
      const source = (y + y0) * width + x + x0
      const target = (y * patchWidth + x) * 4
      if (selected[source] || ramp[source]) {
        out[target] = filled[source * 4]
        out[target + 1] = filled[source * 4 + 1]
        out[target + 2] = filled[source * 4 + 2]
        out[target + 3] = selected[source] ? 255 : RETOUCH.rampAlpha
      } else {
        // Transparent: a flat paper colour compresses best and never fringes.
        out[target] = paper[0]
        out[target + 1] = paper[1]
        out[target + 2] = paper[2]
        out[target + 3] = 0
      }
    }
  }
  return { rgba: out, width: patchWidth, height: patchHeight, x: x0, y: y0 }
}

/** Alpha-composite a patch over region pixels (RGBA): what a viewer shows. Returns new RGBA. */
export function compositePatch(regionRgba, width, height, patch) {
  const out = new Uint8ClampedArray(regionRgba)
  if (!patch) return out
  for (let y = 0; y < patch.height; y += 1) {
    for (let x = 0; x < patch.width; x += 1) {
      const tx = x + patch.x
      const ty = y + patch.y
      if (tx < 0 || ty < 0 || tx >= width || ty >= height) continue
      const source = (y * patch.width + x) * 4
      const alpha = patch.rgba[source + 3] / 255
      if (!alpha) continue
      const target = (ty * width + tx) * 4
      for (let channel = 0; channel < 3; channel += 1) {
        out[target + channel] = clampByte(out[target + channel] * (1 - alpha) + patch.rgba[source + channel] * alpha)
      }
    }
  }
  return out
}
