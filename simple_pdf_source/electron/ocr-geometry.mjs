// Tesseract result -> PDF geometry (WP1). Pure: the caller supplies toPdf, the
// pdf.js viewport mapping of the OCR render (convertToPdfPoint), which already
// handles /Rotate, CropBox/MediaBox origins and UserUnit. OCR-image points are
// first mapped back through the deskew rotation, then through toPdf.

import { deskewPoint } from './ocr-preprocess.mjs'

export const OCR_RESULT_SCHEMA = 1
export const OCR_ENGINE_INFO = Object.freeze({ name: 'tesseract.js', version: '7.0.0', core: 'simd-lstm', model: 'best_int' })

export const OCR_FILTER = Object.freeze({
  /** Words below this confidence are dropped. */
  minWordConfidence: 15,
  /** Words below this confidence need >= 2 letters/digits and a confident line. */
  weakWordConfidence: 40,
  weakWordMinAlphanumerics: 2,
  weakWordMinLineConfidence: 50,
  /** Lines whose mean word confidence is below this are dropped (junk over photos and specks). */
  minLineConfidence: 25,
  /** Baselines within this angle of a page axis are written exactly along the axis. */
  snapDegrees: 0.5,
  /** Mixed pages: OCR words overlapping native text by more than this IoU are duplicates. */
  duplicateIou: 0.3,
})

const LIGATURES = new Map([
  ['ﬀ', 'ff'], ['ﬁ', 'fi'], ['ﬂ', 'fl'], ['ﬃ', 'ffi'], ['ﬄ', 'ffl'], ['ﬅ', 'st'], ['ﬆ', 'st'],
])
const ROTATIONS = new Set([0, 90, 180, 270])
const DESCENDERS = /[gjpqy,;()[\]{}|/\\@$Q]/
/** Letters whose glyphs rest on the baseline in Latin fonts (no descenders, tails or old-style figures). */
const BASELINE_LETTERS = /^[A-IK-PR-Za-fh-ik-or-xz]$/

/**
 * Text that the 2-byte GlyphLessFont encoding can carry: ligatures decomposed,
 * NFC, characters outside the BMP and lone surrogates replaced with U+FFFD, and
 * control characters removed.
 */
export function normalizeOcrText(text) {
  let value = String(text ?? '').replace(/[ﬀ-ﬆ]/g, (ligature) => LIGATURES.get(ligature) ?? ligature)
  value = value.normalize('NFC')
  let out = ''
  for (let i = 0; i < value.length; i += 1) {
    const code = value.charCodeAt(i)
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(i + 1)
      if (next >= 0xdc00 && next <= 0xdfff) i += 1
      out += '�'
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      out += '�'
    } else if (code < 0x20 || (code >= 0x7f && code <= 0x9f)) {
      continue
    } else {
      out += value[i]
    }
  }
  return out
}

const finite = (value) => typeof value === 'number' && Number.isFinite(value)
const median = (values) => {
  if (!values.length) return NaN
  const sorted = [...values].sort((a, b) => a - b)
  const middle = sorted.length >> 1
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2
}
const dot = (a, b) => a.x * b.x + a.y * b.y
const alphanumerics = (text) => (text.match(/[\p{L}\p{N}]/gu) ?? []).length
const validBox = (box) => box && finite(box.x0) && finite(box.y0) && finite(box.x1) && finite(box.y1) && box.x1 > box.x0 && box.y1 > box.y0

function rectOfPoints(points) {
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const point of points) {
    if (point.x < minX) minX = point.x
    if (point.y < minY) minY = point.y
    if (point.x > maxX) maxX = point.x
    if (point.y > maxY) maxY = point.y
  }
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY }
}

function unionRects(rects) {
  return rectOfPoints(rects.flatMap((rect) => [{ x: rect.x, y: rect.y }, { x: rect.x + rect.width, y: rect.y + rect.height }]))
}

function rectContains(rect, point) {
  return point.x >= rect.x && point.x <= rect.x + rect.width && point.y >= rect.y && point.y <= rect.y + rect.height
}

function intersectionOverUnion(a, b) {
  const width = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x)
  const height = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y)
  if (width <= 0 || height <= 0) return 0
  const intersection = width * height
  return intersection / (a.width * a.height + b.width * b.height - intersection)
}

const validRects = (rects) => (Array.isArray(rects) ? rects.filter((rect) => rect && finite(rect.x) && finite(rect.y) && finite(rect.width) && finite(rect.height) && rect.width > 0 && rect.height > 0) : [])

/** Unit direction snapped to the nearest page axis when within `snapDegrees`. */
function snapDirection(direction, snapDegrees) {
  const angle = Math.atan2(direction.y, direction.x)
  const axis = Math.round(angle / (Math.PI / 2))
  const delta = angle - axis * (Math.PI / 2)
  if (Math.abs(delta) >= (snapDegrees * Math.PI) / 180) return { direction, snapped: false }
  const quadrant = ((axis % 4) + 4) % 4
  return { direction: [{ x: 1, y: 0 }, { x: 0, y: 1 }, { x: -1, y: 0 }, { x: 0, y: -1 }][quadrant], snapped: true }
}

/**
 * Tesseract's reported baseline as { origin, u } (u = unit reading direction),
 * or null when it is missing or diagonal. Lines are horizontal, or vertical
 * when Tesseract recognised rotated text by itself (it reads text turned 90
 * degrees clockwise that way, with empty glyph boxes).
 */
function reportedBaseline(line) {
  const reported = line?.baseline
  if (!reported || !finite(reported.x0) || !finite(reported.y0) || !finite(reported.x1) || !finite(reported.y1)) return null
  const dx = reported.x1 - reported.x0
  const dy = reported.y1 - reported.y0
  const length = Math.hypot(dx, dy)
  if (length < 1 || (Math.abs(dy) > 0.6 * Math.abs(dx) && Math.abs(dx) > 0.6 * Math.abs(dy))) return null
  return { origin: { x: reported.x0, y: reported.y0 }, u: { x: dx / length, y: dy / length } }
}

/**
 * The line's baseline frame in OCR pixels: a point on the baseline, the unit
 * reading direction `u`, `up` (towards the ascenders) and whether the line runs
 * vertically. For horizontal lines Tesseract's fitted baseline can sit a few
 * pixels low or tilt where tails and descenders pull it (the 18 pt fixture
 * heading: 4 px low at one end), so when enough letters that rest on the
 * baseline are recognised, their glyph bottoms refine it: Theil-Sen slope when
 * they span the line, then the median offset. Otherwise Tesseract's baseline is
 * used, and a degenerate one falls back to the median bottom of words without
 * descenders.
 */
function lineFrame(line, words) {
  let frame = reportedBaseline(line)
  if (frame && Math.abs(frame.u.y) > Math.abs(frame.u.x)) {
    return { ...frame, up: { x: frame.u.y, y: -frame.u.x }, vertical: true }
  }
  if (frame && frame.u.x < 0) frame = { origin: frame.origin, u: { x: -frame.u.x, y: -frame.u.y } }
  const horizontal = (anchorX, anchorY, slope) => {
    const length = Math.hypot(1, slope)
    const u = { x: 1 / length, y: slope / length }
    return { origin: { x: anchorX, y: anchorY }, u, up: { x: u.y, y: -u.x }, vertical: false }
  }
  let slope = frame ? frame.u.y / frame.u.x : null
  const anchorX = frame ? frame.origin.x : words[0].box.x0
  const points = []
  for (const word of words) {
    for (const symbol of Array.isArray(word.symbols) ? word.symbols : []) {
      if (!validBox(symbol?.bbox) || symbol.is_superscript || symbol.is_subscript || symbol.is_dropcap) continue
      if (!BASELINE_LETTERS.test(String(symbol.text ?? ''))) continue
      points.push({ x: (symbol.bbox.x0 + symbol.bbox.x1) / 2, y: symbol.bbox.y1 })
    }
  }
  if (points.length >= 3) {
    const span = Math.max(...points.map((point) => point.x)) - Math.min(...points.map((point) => point.x))
    const lineWidth = words[words.length - 1].box.x1 - words[0].box.x0
    if (points.length >= 6 && span >= lineWidth * 0.4) {
      const slopes = []
      for (let i = 0; i < points.length; i += 1) {
        for (let j = i + 1; j < points.length; j += 1) {
          const dx = points[j].x - points[i].x
          if (Math.abs(dx) >= span * 0.25) slopes.push((points[j].y - points[i].y) / dx)
        }
      }
      const fitted = median(slopes)
      if (slopes.length >= 3 && Math.abs(fitted) < 0.6) slope = fitted
    }
    const usedSlope = slope ?? 0
    return horizontal(anchorX, median(points.map((point) => point.y - usedSlope * (point.x - anchorX))), usedSlope)
  }
  if (frame) return horizontal(anchorX, frame.origin.y, slope)
  const plain = words.filter((word) => !DESCENDERS.test(word.text))
  const bottoms = (plain.length ? plain : words).map((word) => word.box.y1)
  return horizontal(anchorX, median(bottoms), 0)
}

/** Where a word starts and ends along its line's baseline (frame parameter t). */
function wordSpan(frame, box) {
  const a = frame.vertical ? (box.y0 - frame.origin.y) / frame.u.y : (box.x0 - frame.origin.x) / frame.u.x
  const b = frame.vertical ? (box.y1 - frame.origin.y) / frame.u.y : (box.x1 - frame.origin.x) / frame.u.x
  return { start: Math.min(a, b), end: Math.max(a, b) }
}

/** The box corners in reading orientation: top-left, top-right, bottom-right, bottom-left. */
function readingCorners(frame, box) {
  const { x0, y0, x1, y1 } = box
  if (!frame.vertical) return [{ x: x0, y: y0 }, { x: x1, y: y0 }, { x: x1, y: y1 }, { x: x0, y: y1 }]
  return frame.u.y > 0
    ? [{ x: x1, y: y0 }, { x: x1, y: y1 }, { x: x0, y: y1 }, { x: x0, y: y0 }]
    : [{ x: x0, y: y1 }, { x: x0, y: y0 }, { x: x1, y: y0 }, { x: x1, y: y1 }]
}

/**
 * Orientation fallback policy (design 4.4, recalibrated on the fixtures). In a
 * wrong orientation Tesseract's own page confidence drops to about 30-35 while
 * upright pages score 94-95; the mean of the words this module keeps stays near
 * 65 because the filters drop the worst junk, so the raw confidence decides.
 * Text turned clockwise is read by Tesseract as vertical lines with confidence
 * 95 but split reading order, so a mostly vertical reading is retried too.
 */
export const ORIENTATION_POLICY = Object.freeze({
  /** Below this Tesseract page confidence the page may be sideways or upside down. */
  minConfidence: 60,
  /** Above this share of vertical lines Tesseract turned the text itself. */
  verticalShare: 0.5,
  /** A probe must beat a weak first pass by this many confidence points. */
  improvement: 25,
  /** A horizontal reading may be this much less confident than a vertical one. */
  verticalTolerance: 10,
})

const verticalShare = (reading) => (reading.lines > 0 ? reading.verticalLines / reading.lines : 0)

/**
 * `reading`: { tesseractConfidence, meanConfidence, wordCount, lines,
 * verticalLines } of one recognition pass. True when other orientations should
 * be probed: a weak pass (Tesseract confidence, or the design's thresholds on
 * the kept words) or a mostly vertical one. Pages without recognised lines are
 * never probed.
 */
export function needsOrientationCheck(reading) {
  if (!reading || !(reading.lines > 0)) return false
  return reading.tesseractConfidence < ORIENTATION_POLICY.minConfidence
    || verticalShare(reading) > ORIENTATION_POLICY.verticalShare
    || (reading.meanConfidence < 45 && reading.wordCount < 20)
    || reading.meanConfidence < 30
}

/** Confidence x words, halved for an all-vertical reading. */
export function readingScore(reading) {
  return reading.tesseractConfidence * reading.wordCount * (1 - 0.5 * verticalShare(reading))
}

/**
 * The probe ({ correction, reading }) worth recognising again at full
 * resolution, or null. Horizontal readings are preferred; the winner must beat
 * a weak first pass by `improvement` points, or come within
 * `verticalTolerance` of a vertical one.
 */
export function pickOrientation(first, probes) {
  const usable = probes.filter((probe) => probe.reading.wordCount > 0)
  const horizontal = usable.filter((probe) => verticalShare(probe.reading) <= ORIENTATION_POLICY.verticalShare)
  const candidates = horizontal.length ? horizontal : usable
  if (!candidates.length) return null
  const best = candidates.reduce((winner, probe) => (readingScore(probe.reading) > readingScore(winner.reading) ? probe : winner))
  const required = verticalShare(first) > ORIENTATION_POLICY.verticalShare && verticalShare(best.reading) <= ORIENTATION_POLICY.verticalShare
    ? first.tesseractConfidence - ORIENTATION_POLICY.verticalTolerance
    : first.tesseractConfidence + ORIENTATION_POLICY.improvement
  return best.reading.tesseractConfidence >= required ? best : null
}

/**
 * How Tesseract read the page: recognised lines and words, and how many lines
 * run vertically (text Tesseract turned by itself). Used to decide whether a
 * page might be sideways or upside down.
 */
export function tesseractReadingStats(blocks) {
  let lines = 0
  let verticalLines = 0
  let words = 0
  for (const block of Array.isArray(blocks) ? blocks : []) {
    for (const paragraph of Array.isArray(block?.paragraphs) ? block.paragraphs : []) {
      for (const line of Array.isArray(paragraph?.lines) ? paragraph.lines : []) {
        const count = (Array.isArray(line?.words) ? line.words : []).filter((word) => normalizeOcrText(word?.text).trim()).length
        if (!count) continue
        lines += 1
        words += count
        const frame = reportedBaseline(line)
        if (frame ? Math.abs(frame.u.y) > Math.abs(frame.u.x) : validBox(line.bbox) && line.bbox.y1 - line.bbox.y0 > 2 * (line.bbox.x1 - line.bbox.x0)) verticalLines += 1
      }
    }
  }
  return { lines, verticalLines, words }
}

/**
 * Convert Tesseract `blocks` (pixel space of the OCR image) to an
 * OcrPageResult in PDF user space. Reading order is Tesseract's block ->
 * paragraph -> line order.
 */
export function tesseractToOcrPage(blocks, ctx) {
  if (!ctx || typeof ctx.toPdf !== 'function') throw new Error('OCR geometry needs a toPdf mapping.')
  if (!finite(ctx.dpi) || ctx.dpi <= 0) throw new Error('OCR geometry needs the render resolution.')
  if (!Number.isInteger(ctx.pageIndex) || ctx.pageIndex < 0) throw new Error('OCR geometry needs a page index.')
  const rotation = ROTATIONS.has(ctx.rotation) ? ctx.rotation : 0
  const orientationCorrectedBy = ROTATIONS.has(ctx.orientationCorrectedBy) ? ctx.orientationCorrectedBy : 0
  const deskew = ctx.deskew && finite(ctx.deskew.degrees) && ctx.deskew.degrees !== 0 ? ctx.deskew : null
  const filter = { ...OCR_FILTER, ...(ctx.filter ?? {}) }
  const nativeText = validRects(ctx.nativeText)
  const imageRects = validRects(ctx.imageRects)
  const restrictToImages = Array.isArray(ctx.imageRects)

  const toPage = (x, y) => {
    const source = deskew ? deskewPoint(x, y, deskew) : { x, y }
    const point = ctx.toPdf(source.x, source.y)
    const px = Array.isArray(point) ? point[0] : point?.x
    const py = Array.isArray(point) ? point[1] : point?.y
    if (!finite(px) || !finite(py)) throw new Error('OCR geometry mapping returned an invalid point.')
    return { x: px, y: py }
  }
  // PDF units per OCR pixel (a similarity transform: rotation + uniform scale).
  const origin = toPage(0, 0)
  const unit = toPage(1, 0)
  const unitsPerPixel = Math.hypot(unit.x - origin.x, unit.y - origin.y)
  if (!(unitsPerPixel > 0)) throw new Error('OCR geometry mapping is degenerate.')

  const paragraphs = []
  let lineNumber = 0
  let paragraphNumber = 0
  let confidenceSum = 0
  let wordCount = 0
  for (const block of Array.isArray(blocks) ? blocks : []) {
    for (const paragraph of Array.isArray(block?.paragraphs) ? block.paragraphs : []) {
      const lines = []
      for (const line of Array.isArray(paragraph?.lines) ? paragraph.lines : []) {
        const candidates = []
        for (const word of Array.isArray(line?.words) ? line.words : []) {
          const text = normalizeOcrText(word?.text).trim()
          if (!text || /\s/.test(text) || !validBox(word.bbox)) continue
          candidates.push({ text, confidence: finite(word.confidence) ? word.confidence : 0, box: word.bbox, symbols: word.symbols })
        }
        if (!candidates.length) continue
        const lineConfidence = candidates.reduce((sum, word) => sum + word.confidence, 0) / candidates.length
        if (lineConfidence < filter.minLineConfidence) continue
        const confident = candidates.filter((word) => word.confidence >= filter.minWordConfidence
          && (word.confidence >= filter.weakWordConfidence
            || (alphanumerics(word.text) >= filter.weakWordMinAlphanumerics && lineConfidence >= filter.weakWordMinLineConfidence)))
        if (!confident.length) continue

        const frame = lineFrame(line, confident)
        const on = (t) => toPage(frame.origin.x + t * frame.u.x, frame.origin.y + t * frame.u.y)
        const start = on(wordSpan(frame, confident[0].box).start)
        const end = on(wordSpan(frame, confident[confident.length - 1].box).end)
        let direction = { x: end.x - start.x, y: end.y - start.y }
        const length = Math.hypot(direction.x, direction.y)
        if (length > 1e-9) direction = { x: direction.x / length, y: direction.y / length }
        else {
          // A one-glyph line: the reading direction mapped to the page.
          const ahead = on(1)
          const here = on(0)
          direction = { x: (ahead.x - here.x) / unitsPerPixel, y: (ahead.y - here.y) / unitsPerPixel }
        }
        ;({ direction } = snapDirection(direction, filter.snapDegrees))

        let words = confident.map((word) => {
          const span = wordSpan(frame, word.box)
          const wordOrigin = on(span.start)
          const wordEnd = on(span.end)
          const quad = readingCorners(frame, word.box).map((corner) => toPage(corner.x, corner.y))
          return {
            text: word.text,
            confidence: word.confidence,
            origin: wordOrigin,
            end: wordEnd,
            dir: direction,
            width: Math.max(0, dot({ x: wordEnd.x - wordOrigin.x, y: wordEnd.y - wordOrigin.y }, direction)),
            gap: 0,
            quad,
            bbox: rectOfPoints(quad),
            heightPx: frame.vertical ? word.box.x1 - word.box.x0 : word.box.y1 - word.box.y0,
          }
        })
        if (nativeText.length) {
          words = words.filter((word) => {
            const centre = { x: word.bbox.x + word.bbox.width / 2, y: word.bbox.y + word.bbox.height / 2 }
            return !nativeText.some((rect) => rectContains(rect, centre) || intersectionOverUnion(rect, word.bbox) > filter.duplicateIou)
          })
        }
        if (restrictToImages) {
          words = words.filter((word) => {
            const centre = { x: word.bbox.x + word.bbox.width / 2, y: word.bbox.y + word.bbox.height / 2 }
            return imageRects.some((rect) => rectContains(rect, centre))
          })
        }
        if (!words.length) continue
        for (let index = 0; index < words.length - 1; index += 1) {
          const word = words[index]
          const next = words[index + 1]
          word.gap = Math.max(0, dot({ x: next.origin.x - word.end.x, y: next.origin.y - word.end.y }, direction))
        }

        const heights = words.map((word) => word.heightPx)
        const medianHeight = median(heights)
        const rowHeight = line?.rowAttributes?.rowHeight
        const rowPixels = finite(rowHeight) && rowHeight > 0 && rowHeight >= medianHeight * 0.5 && rowHeight <= medianHeight * 2
          ? rowHeight
          : medianHeight
        const outputWords = words.map(({ end, heightPx, ...word }) => word)
        const confidence = outputWords.reduce((sum, word) => sum + word.confidence, 0) / outputWords.length
        confidenceSum += outputWords.reduce((sum, word) => sum + word.confidence, 0)
        wordCount += outputWords.length
        lineNumber += 1
        lines.push({
          id: `ocr-${ctx.pageIndex + 1}-l${lineNumber}`,
          text: outputWords.map((word) => word.text).join(' '),
          confidence,
          fontSize: rowPixels * unitsPerPixel,
          baseline: { origin: outputWords[0].origin, dir: direction },
          bbox: unionRects(outputWords.map((word) => word.bbox)),
          words: outputWords,
        })
      }
      if (!lines.length) continue
      paragraphNumber += 1
      paragraphs.push({
        id: `ocr-${ctx.pageIndex + 1}-p${paragraphNumber}`,
        bbox: unionRects(lines.map((line) => line.bbox)),
        lines,
        ltr: !(paragraph?.is_ltr === false || paragraph?.is_ltr === 0),
      })
    }
  }
  return {
    schema: OCR_RESULT_SCHEMA,
    pageIndex: ctx.pageIndex,
    contentKey: String(ctx.contentKey ?? ''),
    engine: { ...OCR_ENGINE_INFO, language: String(ctx.language ?? 'eng') },
    rotation,
    dpi: ctx.dpi,
    deskewDegrees: deskew ? deskew.degrees : 0,
    orientationCorrectedBy,
    meanConfidence: wordCount ? confidenceSum / wordCount : 0,
    wordCount,
    paragraphs,
  }
}

const round = (value, digits) => {
  const factor = 10 ** digits
  const rounded = Math.round(value * factor) / factor
  return Object.is(rounded, -0) ? 0 : rounded
}

/** The compact per-page payload of the 'ocr-text-layer' mutation (validated again in main). */
export function ocrPageToLayerPayload(result) {
  return {
    pageIndex: result.pageIndex,
    lines: result.paragraphs.flatMap((paragraph) => paragraph.lines).map((line) => ({
      fontSize: round(line.fontSize, 3),
      words: line.words.map((word) => ({
        text: word.text,
        x: round(word.origin.x, 3),
        y: round(word.origin.y, 3),
        dx: round(word.dir.x, 6),
        dy: round(word.dir.y, 6),
        width: round(word.width, 3),
        gap: round(word.gap, 3),
      })),
    })),
  }
}

/**
 * The 'ocr-text-layer' document mutation for a set of recognised pages. Pages
 * without words stay in the list: with replaceExisting their earlier layer must
 * still be removed.
 */
export function buildOcrLayerOperation(results, options = {}) {
  const language = String(options.language ?? results[0]?.engine?.language ?? 'eng')
  return {
    type: 'ocr-text-layer',
    replaceExisting: options.replaceExisting === true,
    meta: { engine: `${OCR_ENGINE_INFO.name} ${OCR_ENGINE_INFO.version}`, language },
    pages: results.map(ocrPageToLayerPayload),
  }
}
