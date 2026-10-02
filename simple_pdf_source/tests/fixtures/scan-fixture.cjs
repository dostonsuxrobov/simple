'use strict'

// Deterministic synthetic "scanned" pages for the OCR tests: a page of known
// text is laid out with pdf-lib standard fonts, rasterised with mupdf, optionally
// degraded (skew, grain, specks, uneven light, stain, blur, low contrast) and
// wrapped as an image-only PDF. Ground truth is the laid-out text plus, per
// word, its ink box and baseline mapped into the scan page's PDF space.

const PAGE_WIDTH = 612
const PAGE_HEIGHT = 792

const TEXT_BLOCKS = [
  { font: 'HelveticaBold', size: 18, text: 'Quarterly Operations Report' },
  { font: 'TimesRoman', size: 11, text: 'The committee reviewed the proposed budget for the third quarter and approved the revised allocation for maintenance, staffing, and training. Several members noted that the warehouse lease expires on September 30, 2026, and recommended that negotiations begin immediately.' },
  { font: 'Helvetica', size: 10.5, text: 'Invoice number 48213 was paid in full on July 14. The remaining balance of $1,250.75 will be transferred to account 7741-0093 before the end of the month. Please contact the finance office with any questions about these figures.' },
  { font: 'Courier', size: 10, text: 'SKU-0042  Widget, blue     12 units   $3.50\nSKU-0107  Bracket, steel   40 units   $0.85\nSKU-0311  Gasket (large)    7 units  $12.00' },
  { font: 'TimesRomanBold', size: 12, text: 'Action items' },
  { font: 'TimesRoman', size: 11, text: 'Ms. Alvarez will prepare a summary of vendor quotes; Mr. Chen will schedule the safety inspection; the board will vote on the expansion plan at its next meeting.' },
]

/** Fixture variants used by the accuracy script (design section 6.1). */
const VARIANTS = {
  clean300: { dpi: 300 },
  clean200: { dpi: 200 },
  low150: { dpi: 150 },
  noisy300: { dpi: 300, degrade: { skewDeg: 1.2, noise: 14, specks: 0.0005, gradient: true, seed: 7 } },
  skew4: { dpi: 300, degrade: { skewDeg: 4 } },
  hard300: { dpi: 300, hard: { seed: 99 } },
  rotated90Crop: { dpi: 300, layout: 'rotated90Crop' },
  color300: { dpi: 300, color: [0.1, 0.2, 0.6] },
  /** Scanned on its side, text turned counter-clockwise: a landscape page without /Rotate (needs the orientation fallback). */
  sideways300: { dpi: 300, layout: 'sideways' },
  /** Text turned clockwise: Tesseract reads it as vertical lines by itself. */
  sidewaysCw300: { dpi: 300, layout: 'sidewaysCw' },
}

let mupdfModule
async function loadMupdf() {
  mupdfModule ??= import('mupdf')
  return mupdfModule
}

/** Linear congruential generator; the same seed always gives the same page. */
function createRandom(seed = 1) {
  let state = seed >>> 0
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    return state / 4294967296
  }
}

function wrapLines(font, size, text, maxWidth = 470) {
  const lines = []
  for (const paragraph of text.split('\n')) {
    let line = ''
    for (const word of paragraph.split(' ')) {
      const candidate = line ? `${line} ${word}` : word
      if (font.widthOfTextAtSize(candidate, size) > maxWidth && line) {
        lines.push(line)
        line = word
      } else line = candidate
    }
    if (line) lines.push(line)
  }
  return lines
}

/**
 * The vector source page and its layout: lines with baseline y and per-word
 * advance extents, in PDF user space (612 x 792, origin bottom-left).
 */
async function vectorPage(options = {}) {
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib')
  const color = options.color ?? [0.08, 0.08, 0.1]
  const doc = await PDFDocument.create()
  const page = doc.addPage([PAGE_WIDTH, PAGE_HEIGHT])
  const lines = []
  let y = 720
  for (const block of TEXT_BLOCKS) {
    const font = await doc.embedFont(StandardFonts[block.font])
    for (const text of wrapLines(font, block.size, block.text)) {
      page.drawText(text, { x: 72, y, size: block.size, font, color: rgb(color[0], color[1], color[2]) })
      const words = []
      const pattern = /\S+/g
      let match
      while ((match = pattern.exec(text))) {
        const x0 = 72 + font.widthOfTextAtSize(text.slice(0, match.index), block.size)
        words.push({ text: match[0], x0, x1: x0 + font.widthOfTextAtSize(match[0], block.size) })
      }
      lines.push({ text, font: block.font, size: block.size, x: 72, baseline: y, words })
      y -= block.size * 1.35
    }
    y -= block.size * 0.8
  }
  return { bytes: await doc.save(), lines }
}

/** Render one page with mupdf: { width, height, channels (1 or 3), data }. Applies /Rotate and the CropBox. */
async function rasterize(bytes, dpi, options = {}) {
  const mupdf = await loadMupdf()
  const document = mupdf.Document.openDocument(bytes, 'application/pdf')
  try {
    const page = document.loadPage(options.pageIndex ?? 0)
    try {
      const colorspace = options.color ? mupdf.ColorSpace.DeviceRGB : mupdf.ColorSpace.DeviceGray
      const pixmap = page.toPixmap(mupdf.Matrix.scale(dpi / 72, dpi / 72), colorspace, false, true)
      try {
        const width = pixmap.getWidth()
        const height = pixmap.getHeight()
        const channels = options.color ? 3 : 1
        const stride = pixmap.getStride()
        const source = pixmap.getPixels()
        const data = new Uint8Array(width * height * channels)
        for (let row = 0; row < height; row += 1) data.set(source.subarray(row * stride, row * stride + width * channels), row * width * channels)
        return { width, height, channels, data }
      } finally {
        pixmap.destroy()
      }
    } finally {
      page.destroy()
    }
  } finally {
    document.destroy()
  }
}

/** Bilinear rotation of the content by `degrees` (clockwise on screen) about the image centre; white fill. */
function rotateImage(image, degrees) {
  const { width: w, height: h, channels, data } = image
  const out = new Uint8Array(data.length).fill(255)
  const angle = (degrees * Math.PI) / 180
  const cos = Math.cos(angle)
  const sin = Math.sin(angle)
  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      const px = x + 0.5 - w / 2
      const py = y + 0.5 - h / 2
      const sx = cos * px + sin * py + w / 2 - 0.5
      const sy = -sin * px + cos * py + h / 2 - 0.5
      const x0 = Math.floor(sx)
      const y0 = Math.floor(sy)
      if (x0 < 0 || y0 < 0 || x0 + 1 >= w || y0 + 1 >= h) continue
      const fx = sx - x0
      const fy = sy - y0
      const top = (y0 * w + x0) * channels
      const bottom = top + w * channels
      for (let c = 0; c < channels; c += 1) {
        out[(y * w + x) * channels + c] = data[top + c] * (1 - fx) * (1 - fy) + data[top + channels + c] * fx * (1 - fy)
          + data[bottom + c] * (1 - fx) * fy + data[bottom + channels + c] * fx * fy + 0.5
      }
    }
  }
  return { width: w, height: h, channels, data: out }
}

/** Scanner-like degradation of a grey image: skew, grain (sigma = noise), specks and uneven light. */
function degrade(image, options = {}) {
  if (image.channels !== 1) throw new Error('degrade() expects a grey image')
  let out = options.skewDeg ? rotateImage(image, options.skewDeg) : { ...image, data: new Uint8Array(image.data) }
  const { width: w, height: h } = out
  const data = out.data
  const random = createRandom(options.seed ?? 7)
  if (options.noise || options.specks) {
    for (let i = 0; i < data.length; i += 1) {
      let value = data[i]
      if (options.noise) value += (random() + random() + random() - 1.5) * options.noise * 2
      if (options.specks && random() < options.specks) value -= 180
      data[i] = Math.max(0, Math.min(255, value))
    }
  }
  if (options.gradient) {
    for (let y = 0; y < h; y += 1) {
      for (let x = 0; x < w; x += 1) data[y * w + x] = Math.max(0, data[y * w + x] - (x / w) * 18 - (y / h) * 10)
    }
  }
  if (options.blur) out = boxBlur(out)
  return out
}

function boxBlur(image) {
  const { width: w, height: h, data } = image
  const out = new Uint8Array(data)
  for (let y = 1; y < h - 1; y += 1) {
    for (let x = 1; x < w - 1; x += 1) {
      let total = 0
      for (let dy = -1; dy <= 1; dy += 1) for (let dx = -1; dx <= 1; dx += 1) total += data[(y + dy) * w + x + dx]
      out[y * w + x] = total / 9
    }
  }
  return { ...image, data: out }
}

/** Faded toner on grey paper with a coffee stain, shading, grain and a soft focus. */
function hardScan(image, options = {}) {
  if (image.channels !== 1) throw new Error('hardScan() expects a grey image')
  const { width: w, height: h, data } = image
  const scale = w / 2550
  const random = createRandom(options.seed ?? 99)
  const out = new Uint8Array(data.length)
  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      const i = y * w + x
      const value = data[i] / 255
      let paper = 205 - 40 * Math.exp(-(((x - 1300 * scale) / (500 * scale)) ** 2 + ((y - 800 * scale) / (260 * scale)) ** 2))
      paper -= (y / h) * 25
      const ink = 115
      out[i] = Math.max(0, Math.min(255, ink + (paper - ink) * value + (random() + random() - 1) * 10))
    }
  }
  return boxBlur({ ...image, data: out })
}

/** Rotate a raster 90 degrees counter-clockwise (pixel (x, y) -> (y, width - 1 - x)). */
function rotateRaster90ccw(image) {
  const { width: w, height: h, channels, data } = image
  const out = new Uint8Array(data.length)
  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      for (let c = 0; c < channels; c += 1) out[((w - 1 - x) * h + y) * channels + c] = data[(y * w + x) * channels + c]
    }
  }
  return { width: h, height: w, channels, data: out }
}

/** Rotate a raster 90 degrees clockwise (pixel (x, y) -> (height - 1 - y, x)). */
function rotateRaster90cw(image) {
  const { width: w, height: h, channels, data } = image
  const out = new Uint8Array(data.length)
  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      for (let c = 0; c < channels; c += 1) out[(x * h + (h - 1 - y)) * channels + c] = data[(y * w + x) * channels + c]
    }
  }
  return { width: h, height: w, channels, data: out }
}

/** Wrap a raster as an image-only PDF page (Flate image XObject, no re-encoding). */
async function imagePdf(image, options = {}) {
  const { PDFDocument, degrees, pushGraphicsState, popGraphicsState, concatTransformationMatrix, drawObject } = require('pdf-lib')
  const pageSize = options.pageSize ?? [PAGE_WIDTH, PAGE_HEIGHT]
  const rect = options.rect ?? { x: 0, y: 0, width: pageSize[0], height: pageSize[1] }
  const doc = await PDFDocument.create()
  const page = doc.addPage(pageSize)
  const stream = doc.context.flateStream(image.data, {
    Type: 'XObject',
    Subtype: 'Image',
    Width: image.width,
    Height: image.height,
    ColorSpace: image.channels === 3 ? 'DeviceRGB' : 'DeviceGray',
    BitsPerComponent: 8,
  })
  const name = page.node.newXObject('Scan', doc.context.register(stream))
  page.pushOperators(pushGraphicsState(), concatTransformationMatrix(rect.width, 0, 0, rect.height, rect.x, rect.y), drawObject(name), popGraphicsState())
  if (options.cropBox) page.setCropBox(options.cropBox.x, options.cropBox.y, options.cropBox.width, options.cropBox.height)
  if (options.rotate) page.setRotation(degrees(options.rotate))
  return doc.save({ useObjectStreams: false })
}

let layoutPromise
/** The vector layout with each word's ink box measured on a clean 300 DPI raster. */
async function groundTruth() {
  layoutPromise ??= (async () => {
    const vector = await vectorPage()
    const dpi = 300
    const raster = await rasterize(vector.bytes, dpi)
    const s = dpi / 72
    const { width: w, height: h, data } = raster
    for (const line of vector.lines) {
      for (const word of line.words) {
        const px0 = Math.max(0, Math.floor((word.x0 - 0.05 * line.size) * s))
        const px1 = Math.min(w, Math.ceil((word.x1 + 0.05 * line.size) * s))
        const py0 = Math.max(0, Math.floor((PAGE_HEIGHT - (line.baseline + 0.95 * line.size)) * s))
        const py1 = Math.min(h, Math.ceil((PAGE_HEIGHT - (line.baseline - 0.3 * line.size)) * s))
        let minX = Infinity
        let minY = Infinity
        let maxX = -Infinity
        let maxY = -Infinity
        for (let y = py0; y < py1; y += 1) {
          for (let x = px0; x < px1; x += 1) {
            if (data[y * w + x] >= 128) continue
            if (x < minX) minX = x
            if (x > maxX) maxX = x
            if (y < minY) minY = y
            if (y > maxY) maxY = y
          }
        }
        if (minX === Infinity) throw new Error(`No ink found for fixture word ${word.text}`)
        word.inkBox = { x: minX / s, y: PAGE_HEIGHT - (maxY + 1) / s, width: (maxX + 1 - minX) / s, height: (maxY + 1 - minY) / s }
      }
    }
    return vector
  })()
  return layoutPromise
}

/**
 * Build a named variant: { name, pdf, nativeDpi, truth }. `truth.words` are in
 * reading order with `centre` (ink box centre) and `baseline` ({ origin, dir })
 * in the scan page's unrotated PDF space.
 */
async function buildScanVariant(name) {
  const variant = VARIANTS[name]
  if (!variant) throw new Error(`Unknown scan fixture ${name}`)
  const truthLayout = await groundTruth()
  const source = variant.color ? await vectorPage({ color: variant.color }) : truthLayout
  let image = await rasterize(source.bytes, variant.dpi, { color: Boolean(variant.color) })
  let map = (u, v) => ({ x: u, y: v })
  let pdf
  if (variant.degrade) {
    image = degrade(image, variant.degrade)
    const s = variant.dpi / 72
    const angle = ((variant.degrade.skewDeg ?? 0) * Math.PI) / 180
    const cx = image.width / 2
    const cy = image.height / 2
    map = (u, v) => {
      const qx = u * s - cx
      const qy = (PAGE_HEIGHT - v) * s - cy
      const px = Math.cos(angle) * qx - Math.sin(angle) * qy + cx
      const py = Math.sin(angle) * qx + Math.cos(angle) * qy + cy
      return { x: px / s, y: PAGE_HEIGHT - py / s }
    }
  }
  if (variant.hard) image = hardScan(image, variant.hard)
  if (variant.layout === 'rotated90Crop') {
    // Landscape media 900 x 700 holding the portrait scan turned on its side,
    // CropBox [50 40 842 652] and /Rotate 90 so it displays upright.
    const rotated = rotateRaster90ccw(image)
    pdf = await imagePdf(rotated, {
      pageSize: [900, 700],
      rect: { x: 50, y: 40, width: 792, height: 612 },
      cropBox: { x: 50, y: 40, width: 792, height: 612 },
      rotate: 90,
    })
    map = (u, v) => ({ x: 50 + (PAGE_HEIGHT - v), y: 40 + u })
  } else if (variant.layout === 'sideways') {
    pdf = await imagePdf(rotateRaster90ccw(image), { pageSize: [PAGE_HEIGHT, PAGE_WIDTH] })
    map = (u, v) => ({ x: PAGE_HEIGHT - v, y: u })
  } else if (variant.layout === 'sidewaysCw') {
    pdf = await imagePdf(rotateRaster90cw(image), { pageSize: [PAGE_HEIGHT, PAGE_WIDTH] })
    map = (u, v) => ({ x: v, y: PAGE_WIDTH - u })
  } else {
    pdf = await imagePdf(image)
  }
  const words = []
  const lines = truthLayout.lines.map((line, lineIndex) => {
    const lineWords = line.words.map((word) => {
      const centre = map(word.inkBox.x + word.inkBox.width / 2, word.inkBox.y + word.inkBox.height / 2)
      const origin = map(word.x0, line.baseline)
      const along = map(word.x0 + 1, line.baseline)
      const length = Math.hypot(along.x - origin.x, along.y - origin.y)
      const truthWord = {
        text: word.text,
        lineIndex,
        centre,
        baseline: { origin, dir: { x: (along.x - origin.x) / length, y: (along.y - origin.y) / length } },
        inkWidth: word.inkBox.width,
      }
      words.push(truthWord)
      return truthWord
    })
    return { text: line.text, size: line.size, font: line.font, words: lineWords }
  })
  return {
    name,
    pdf,
    nativeDpi: variant.dpi,
    truth: { text: lines.map((line) => line.text).join('\n'), lines, words },
  }
}

module.exports = {
  PAGE_WIDTH,
  PAGE_HEIGHT,
  TEXT_BLOCKS,
  VARIANTS,
  createRandom,
  vectorPage,
  rasterize,
  rotateImage,
  degrade,
  hardScan,
  rotateRaster90ccw,
  rotateRaster90cw,
  imagePdf,
  groundTruth,
  buildScanVariant,
}
