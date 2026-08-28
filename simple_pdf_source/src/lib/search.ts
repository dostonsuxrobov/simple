export interface SearchMatchRect {
  left: number
  top: number
  width: number
  height: number
}

interface TextPoint {
  node: Text
  offset: number
}

interface NormalizedTextMap {
  text: string
  starts: TextPoint[]
  ends: TextPoint[]
}

export function normalizeSearchValue(value: string) {
  return value
    .normalize('NFKC')
    .replace(/\u00ad/g, '')
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[‐‑‒–—]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .toLocaleLowerCase()
}

function normalizeMappedFragment(value: string) {
  return value
    .normalize('NFKC')
    .replace(/\u00ad/g, '')
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[‐‑‒–—]/g, '-')
    .toLocaleLowerCase()
}

/**
 * Build the same whitespace-normalized string used by document search while
 * retaining a DOM boundary for every normalized UTF-16 code unit. PDF.js puts
 * words and phrases in separate absolutely positioned spans, so the virtual
 * space between spans is important for phrase matches.
 */
function mapTextLayer(layer: HTMLElement): NormalizedTextMap {
  // The text layer appends spans in visual reading order, but occurrence
  // indices are counted against getTextContent's content-stream order
  // (Sidebar searches pages that have no text layer mounted). Restore stream
  // order here so the k-th counted occurrence resolves to the k-th mapped one.
  const nodes = Array.from(layer.querySelectorAll<HTMLElement>('[data-text-item="true"]'))
    .sort((a, b) => Number(a.dataset.textItemIndex || 0) - Number(b.dataset.textItemIndex || 0))
    .map((span) => span.firstChild)
    .filter((node): node is Text => node instanceof Text && Boolean(node.data))
  const output: string[] = []
  const starts: TextPoint[] = []
  const ends: TextPoint[] = []
  let pendingSpace: { start: TextPoint; end: TextPoint } | null = null

  const append = (value: string, start: TextPoint, end: TextPoint) => {
    for (let index = 0; index < value.length; index += 1) {
      output.push(value[index])
      starts.push(start)
      ends.push(end)
    }
  }

  const consume = (value: string, start: TextPoint, end: TextPoint) => {
    const normalized = normalizeMappedFragment(value)
    for (let index = 0; index < normalized.length; index += 1) {
      const character = normalized[index]
      if (/\s/.test(character)) {
        pendingSpace = pendingSpace
          ? { start: pendingSpace.start, end }
          : { start, end }
        continue
      }
      if (pendingSpace && output.length) append(' ', pendingSpace.start, pendingSpace.end)
      pendingSpace = null
      append(character, start, end)
    }
  }

  nodes.forEach((node, nodeIndex) => {
    if (nodeIndex) {
      const previous = nodes[nodeIndex - 1]
      consume(' ', { node: previous, offset: previous.length }, { node, offset: 0 })
    }
    for (let offset = 0; offset < node.data.length;) {
      const codePoint = node.data.codePointAt(offset)
      if (codePoint === undefined) break
      let character = String.fromCodePoint(codePoint)
      let nextOffset = offset + character.length
      // Keep combining marks with their base so NFKC mapping stays identical
      // to the whole-string normalization used by the search index.
      while (nextOffset < node.data.length) {
        const nextCodePoint = node.data.codePointAt(nextOffset)
        if (nextCodePoint === undefined) break
        const nextCharacter = String.fromCodePoint(nextCodePoint)
        if (!/\p{Mark}/u.test(nextCharacter)) break
        character += nextCharacter
        nextOffset += nextCharacter.length
      }
      consume(character, { node, offset }, { node, offset: nextOffset })
      offset = nextOffset
    }
  })

  // A pending separator is deliberately not flushed: normalizeSearchValue
  // trims trailing whitespace.
  return { text: output.join(''), starts, ends }
}

function occurrenceOffset(haystack: string, needle: string, occurrenceIndex: number) {
  let matchOffset = -1
  let nextOffset = 0
  for (let index = 0; index <= occurrenceIndex; index += 1) {
    matchOffset = haystack.indexOf(needle, nextOffset)
    if (matchOffset < 0) return -1
    nextOffset = matchOffset + needle.length
  }
  return matchOffset
}

/** Resolve one indexed word/phrase match into visible page-relative boxes. */
export function findTextLayerSearchRects(
  textLayer: HTMLElement,
  pageSurface: HTMLElement,
  query: string,
  occurrenceIndex: number,
): SearchMatchRect[] {
  const needle = normalizeSearchValue(query)
  if (!needle || occurrenceIndex < 0) return []
  const mapped = mapTextLayer(textLayer)
  const matchOffset = occurrenceOffset(mapped.text, needle, occurrenceIndex)
  if (matchOffset < 0) return []

  // A match's nodes may be far apart — or even reversed — in DOM order, since
  // the mapping runs in stream order while the layer is laid out visually.
  // Measure one single-node Range per touched text node instead of a single
  // start-to-end Range, which would collapse or cover interposed spans.
  // Virtual inter-node spaces have differing start/end nodes and no ink.
  const nodeIntervals = new Map<Text, { start: number; end: number }>()
  for (let index = matchOffset; index < matchOffset + needle.length; index += 1) {
    const start = mapped.starts[index]
    const end = mapped.ends[index]
    if (!start || !end || start.node !== end.node) continue
    const interval = nodeIntervals.get(start.node)
    if (interval) {
      interval.start = Math.min(interval.start, start.offset)
      interval.end = Math.max(interval.end, end.offset)
    } else {
      nodeIntervals.set(start.node, { start: start.offset, end: end.offset })
    }
  }

  const surfaceBounds = pageSurface.getBoundingClientRect()
  const rectangles: SearchMatchRect[] = []
  const range = document.createRange()
  for (const [node, interval] of nodeIntervals) {
    range.setStart(node, interval.start)
    range.setEnd(node, interval.end)
    for (const rect of range.getClientRects()) {
      if (rect.width > 0.25 && rect.height > 0.25) {
        rectangles.push({
          left: rect.left - surfaceBounds.left,
          top: rect.top - surfaceBounds.top,
          width: rect.width,
          height: rect.height,
        })
      }
    }
  }
  range.detach()
  return rectangles
}
