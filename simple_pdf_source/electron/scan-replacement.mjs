// Minimal-footprint replacement for an edit of scanned text, design section
// 4.8.5 (WP4). When only some words of a scanned line change, only those
// words are retouched and redrawn: the unchanged words keep their scanned
// pixels and their recognised (searchable) text, which is what makes an edit
// blend into the scan. Pure: no DOM, no PDF.

/** A replacement may be squeezed to fit its room by at most this factor (8 %). */
export const MAX_FIT_COMPRESSION = 1.08

/** Whitespace-separated tokens of a line (leading and trailing whitespace ignored). */
export function replacementTokens(text) {
  const trimmed = String(text ?? '').trim()
  return trimmed ? trimmed.split(/\s+/u) : []
}

function unionRects(rects) {
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -Infinity
  let y1 = -Infinity
  for (const rect of rects) {
    if (!rect) continue
    x0 = Math.min(x0, rect.x)
    y0 = Math.min(y0, rect.y)
    x1 = Math.max(x1, rect.x + rect.width)
    y1 = Math.max(y1, rect.y + rect.height)
  }
  return Number.isFinite(x0) ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null
}

/**
 * Decide what a changed scanned line needs.
 *
 * `words` are the line's recognised words with their boxes (PDF space, in
 * reading order, aligned with the tokens of `originalText`). Positions along
 * the line are measured by `x` (the left edge) unless `along` gives them
 * (`start`/`end` of each word along a slanted baseline). `measure(text)` is
 * the width the new text will take (points, at the edit's size, font and
 * horizontal scale) and `spaceWidth` one space. `lineEnd` is where the text
 * may run to at most (the line end plus whatever room the page leaves).
 *
 * Returns
 *   { kind: 'none' }   nothing changed,
 *   { kind: 'words', first, last, originalText, text, rect, start, end, room }
 *                      replace words first..last (inclusive; last = first - 1
 *                      for a pure insertion) with `text`; `rect` is the union
 *                      of the replaced boxes, `start`/`end` the replaced
 *                      extent and `room` the width the new text may use,
 *   { kind: 'line' }   replace the whole line.
 */
export function planScanReplacement(input) {
  const { words, measure } = input
  const spaceWidth = Math.max(0, Number(input.spaceWidth) || 0)
  const oldTokens = replacementTokens(input.originalText)
  const newTokens = replacementTokens(input.newText)
  const sameTokens = oldTokens.length === newTokens.length && oldTokens.every((token, index) => token === newTokens[index])
  if (sameTokens && !/\n/u.test(String(input.newText ?? '').trim())) return { kind: 'none' }
  // Line breaks, or words that do not line up with the boxes: the whole line.
  if (/\n/u.test(String(input.newText ?? '').trim()) || !Array.isArray(words) || words.length !== oldTokens.length || !oldTokens.length || typeof measure !== 'function') {
    return { kind: 'line' }
  }
  const start = (index) => (input.along?.[index]?.start ?? words[index].rect.x)
  const end = (index) => (input.along?.[index]?.end ?? words[index].rect.x + words[index].rect.width)
  const lineEnd = Number.isFinite(input.lineEnd) ? Number(input.lineEnd) : end(words.length - 1)

  // Common prefix and suffix (never overlapping).
  let prefix = 0
  while (prefix < oldTokens.length && prefix < newTokens.length && oldTokens[prefix] === newTokens[prefix]) prefix += 1
  let suffix = 0
  while (suffix < oldTokens.length - prefix && suffix < newTokens.length - prefix
    && oldTokens[oldTokens.length - 1 - suffix] === newTokens[newTokens.length - 1 - suffix]) suffix += 1

  const plan = (first, stop, replacement) => {
    // Old words first..stop-1 become `replacement`.
    const left = first < stop ? start(first) : first > 0 ? end(first - 1) + spaceWidth : start(0)
    const right = stop < words.length ? start(stop) - spaceWidth : lineEnd
    const room = Math.max(0, right - left)
    const text = replacement.join(' ')
    const fits = measure(text) <= room * MAX_FIT_COMPRESSION + 1e-6
    return {
      fits,
      result: {
        kind: 'words',
        first,
        last: stop - 1,
        originalText: oldTokens.slice(first, stop).join(' '),
        text,
        rect: first < stop ? unionRects(words.slice(first, stop).map((word) => word.rect)) : null,
        start: left,
        end: first < stop ? end(stop - 1) : left,
        room,
      },
    }
  }

  const first = prefix
  const stop = oldTokens.length - suffix
  const replacement = newTokens.slice(prefix, newTokens.length - suffix)
  const narrow = plan(first, stop, replacement)
  if (narrow.fits) return narrow.result
  // Too long for its room: re-typeset to the end of the line.
  if (stop < words.length) {
    const extended = plan(first, words.length, newTokens.slice(prefix))
    if (extended.fits) return extended.result
  }
  return { kind: 'line' }
}
