export interface ClientRectBox {
  left: number
  top: number
  right: number
  bottom: number
  width: number
  height: number
}

export interface TextSlice {
  start: number
  end: number
  text: string
  rect: ClientRectBox
}

export interface TextCaretPoint {
  node: Text
  offset: number
  span: HTMLSpanElement
}

function intersectsSelection(range: Range, node: Text) {
  // comparePoint returns -1 before the range, 0 inside it, and 1 after it.
  // Checking both ends avoids Range#getClientRects returning the containing
  // text item (often a complete PDF line) when only a few glyphs are selected.
  try {
    return range.comparePoint(node, 0) !== 1
      && range.comparePoint(node, node.length) !== -1
  } catch {
    return false
  }
}

function clippedRect(rect: DOMRect, bounds: DOMRect): ClientRectBox | null {
  const left = Math.max(rect.left, bounds.left)
  const top = Math.max(rect.top, bounds.top)
  const right = Math.min(rect.right, bounds.right)
  const bottom = Math.min(rect.bottom, bounds.bottom)
  const width = right - left
  const height = bottom - top
  if (![left, top, right, bottom].every(Number.isFinite) || width < 0.5 || height < 1) return null
  return { left, top, right, bottom, width, height }
}

function rangeRectWithin(range: Range, bounds: DOMRect): ClientRectBox | null {
  const pieces = Array.from(range.getClientRects())
    .map((rect) => clippedRect(rect, bounds))
    .filter((rect): rect is ClientRectBox => Boolean(rect))
  if (!pieces.length) return null
  const left = Math.min(...pieces.map((rect) => rect.left))
  const top = Math.min(...pieces.map((rect) => rect.top))
  const right = Math.max(...pieces.map((rect) => rect.right))
  const bottom = Math.max(...pieces.map((rect) => rect.bottom))
  return { left, top, right, bottom, width: right - left, height: bottom - top }
}

function sliceForOffsets(span: HTMLSpanElement, start: number, end: number): TextSlice | null {
  const node = span.firstChild
  if (!(node instanceof Text)) return null
  const safeStart = Math.max(0, Math.min(node.length, start))
  const safeEnd = Math.max(safeStart, Math.min(node.length, end))
  if (safeEnd <= safeStart) return null
  const range = document.createRange()
  try {
    range.setStart(node, safeStart)
    range.setEnd(node, safeEnd)
    const rect = rangeRectWithin(range, span.getBoundingClientRect())
    return rect ? { start: safeStart, end: safeEnd, text: node.data.slice(safeStart, safeEnd), rect } : null
  } catch {
    return null
  } finally {
    range.detach()
  }
}

/** Return a precise single-text-item slice when the user selected part of a PDF line. */
export function selectedTextSliceWithinSpan(selection: Selection | null, span: HTMLSpanElement): TextSlice | null {
  if (!selection || selection.isCollapsed || selection.rangeCount < 1) return null
  const node = span.firstChild
  if (!(node instanceof Text)) return null
  const range = selection.getRangeAt(0)
  // A PageTextEdit has one baseline and rectangle. Do not collapse a multi-line
  // selection into one destructive cover rectangle.
  if (range.startContainer !== node || range.endContainer !== node) return null
  return sliceForOffsets(span, range.startOffset, range.endOffset)
}

/**
 * Pick one Unicode word at the caret. This keeps a normal edit-mode click from
 * replacing the complete PDF.js text item, which is commonly an entire line.
 */
export function wordTextSliceAtPoint(span: HTMLSpanElement, clientX: number, clientY: number): TextSlice | null {
  const text = span.textContent || ''
  if (!text) return null
  const rawOffset = caretOffsetAtPoint(span, clientX, clientY)
  let index = Math.max(0, Math.min(text.length - 1, rawOffset === text.length ? rawOffset - 1 : rawOffset))
  const isWordCharacter = (character: string) => /[\p{L}\p{M}\p{N}'\u2019-]/u.test(character)

  if (!isWordCharacter(text[index])) {
    let next = index
    while (next < text.length && !isWordCharacter(text[next])) next += 1
    let previous = index - 1
    while (previous >= 0 && !isWordCharacter(text[previous])) previous -= 1
    if (next < text.length && (previous < 0 || next - index <= index - previous)) index = next
    else if (previous >= 0) index = previous
    else return sliceForOffsets(span, index, Math.min(text.length, index + 1))
  }

  let start = index
  let end = index + 1
  while (start > 0 && isWordCharacter(text[start - 1])) start -= 1
  while (end < text.length && isWordCharacter(text[end])) end += 1
  return sliceForOffsets(span, start, end)
}

function mergeLineRects(rects: ClientRectBox[]) {
  const sorted = [...rects].sort((a, b) => {
    const sameLine = Math.abs(a.top - b.top) <= Math.max(2, Math.min(a.height, b.height) * 0.3)
    return sameLine ? a.left - b.left : a.top - b.top
  })
  const output: ClientRectBox[] = []

  for (const rect of sorted) {
    const previous = output.at(-1)
    if (previous) {
      const overlap = Math.min(previous.bottom, rect.bottom) - Math.max(previous.top, rect.top)
      const sameLine = overlap >= Math.min(previous.height, rect.height) * 0.55
      const gap = rect.left - previous.right
      if (sameLine && gap <= Math.max(2, Math.min(previous.height, rect.height) * 0.16)) {
        previous.left = Math.min(previous.left, rect.left)
        previous.top = Math.min(previous.top, rect.top)
        previous.right = Math.max(previous.right, rect.right)
        previous.bottom = Math.max(previous.bottom, rect.bottom)
        previous.width = previous.right - previous.left
        previous.height = previous.bottom - previous.top
        continue
      }
    }
    output.push({ ...rect })
  }
  return output
}

/**
 * Return only the glyph runs selected inside `container`.
 *
 * PDF.js commonly places a complete visual line in one absolutely-positioned
 * span. Walking its text nodes and making a sub-range for the selected offsets
 * keeps markup character-precise instead of falling back to that whole span.
 */
export function selectionClientRectsWithin(selection: Selection, container: HTMLElement): ClientRectBox[] {
  if (selection.isCollapsed || selection.rangeCount < 1) return []
  const source = selection.getRangeAt(0)
  const bounds = container.getBoundingClientRect()
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT)
  const rects: ClientRectBox[] = []

  for (let current = walker.nextNode(); current; current = walker.nextNode()) {
    const node = current as Text
    if (!node.length || !intersectsSelection(source, node)) continue
    const start = node === source.startContainer ? source.startOffset : 0
    const end = node === source.endContainer ? source.endOffset : node.length
    if (end <= start) continue

    const part = document.createRange()
    try {
      part.setStart(node, start)
      part.setEnd(node, end)
      for (const rawRect of Array.from(part.getClientRects())) {
        const rect = clippedRect(rawRect, bounds)
        if (rect) rects.push(rect)
      }
    } catch {
      // Ignore a node invalidated by a concurrent text-layer refresh.
    } finally {
      part.detach()
    }
  }

  return mergeLineRects(rects)
}

export function caretOffsetAtPoint(span: HTMLSpanElement, clientX: number, clientY: number) {
  const text = span.firstChild
  if (!(text instanceof Text)) return span.textContent?.length ?? 0

  // Use the glyph geometry that is actually painted in the text layer. Native
  // caretPositionFromPoint can be a few characters off inside transformed PDF
  // spans (especially during a portable build's cold font load). Projecting
  // character centres onto the run's baseline is deterministic for horizontal,
  // vertical, rotated, LTR, and RTL text; items are short, so scanning every
  // glyph midpoint stays cheap and keeps bidi runs (whose projections are not
  // monotonic along the axis) correct.
  const angle = Number(span.dataset.textAngle) || 0
  const axisX = Math.cos(angle)
  const axisY = Math.sin(angle)
  const pointerProjection = clientX * axisX + clientY * axisY
  const characterProjection = (offset: number) => {
    if (offset < 0 || offset >= text.length) return Number.NaN
    const probe = document.createRange()
    try {
      probe.setStart(text, offset)
      probe.setEnd(text, offset + 1)
      const rect = probe.getBoundingClientRect()
      if (rect.width < 0.01 && rect.height < 0.01) return Number.NaN
      return (rect.left + rect.right) / 2 * axisX + (rect.top + rect.bottom) / 2 * axisY
    } catch {
      return Number.NaN
    } finally {
      probe.detach()
    }
  }

  if (text.length) {
    const projections: number[] = []
    for (let offset = 0; offset < text.length; offset += 1) projections.push(characterProjection(offset))
    let nearest = -1
    let nearestDistance = Number.POSITIVE_INFINITY
    for (let offset = 0; offset < text.length; offset += 1) {
      if (!Number.isFinite(projections[offset])) continue
      const distance = Math.abs(pointerProjection - projections[offset])
      if (distance < nearestDistance) {
        nearest = offset
        nearestDistance = distance
      }
    }
    if (nearest >= 0) {
      // The local run direction decides which side of the nearest glyph maps
      // to the next logical offset; an RTL glyph's start edge is its right.
      let forward: boolean | null = null
      for (let offset = nearest + 1; offset < text.length && forward === null; offset += 1) {
        if (Number.isFinite(projections[offset]) && projections[offset] !== projections[nearest]) {
          forward = projections[offset] > projections[nearest]
        }
      }
      for (let offset = nearest - 1; offset >= 0 && forward === null; offset -= 1) {
        if (Number.isFinite(projections[offset]) && projections[offset] !== projections[nearest]) {
          forward = projections[nearest] > projections[offset]
        }
      }
      if (forward === null) forward = span.dir !== 'rtl'
      return (pointerProjection > projections[nearest]) === forward ? nearest + 1 : nearest
    }
  }

  const documentWithCaret = document as Document & {
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null
    caretRangeFromPoint?: (x: number, y: number) => Range | null
  }
  const position = documentWithCaret.caretPositionFromPoint?.(clientX, clientY)
  if (position?.offsetNode === text) return Math.max(0, Math.min(text.length, position.offset))
  const caretRange = documentWithCaret.caretRangeFromPoint?.(clientX, clientY)
  if (caretRange?.startContainer === text) return Math.max(0, Math.min(text.length, caretRange.startOffset))

  // Chromium normally supplies one of the APIs above. This final horizontal
  // fallback covers malformed runs whose individual glyphs have no geometry.
  let low = 0
  let high = text.length
  while (low < high) {
    const middle = Math.floor((low + high) / 2)
    const probe = document.createRange()
    probe.setStart(text, middle)
    probe.setEnd(text, Math.min(text.length, middle + 1))
    const rect = probe.getBoundingClientRect()
    probe.detach()
    if (clientX < rect.left + rect.width / 2) high = middle
    else low = middle + 1
  }
  return low
}

/**
 * Resolve a pointer to a caret only when it is actually over a PDF text item.
 *
 * Chromium's native drag selection tries to infer a visual line from the DOM.
 * PDF text layers are instead made from independently positioned spans, and
 * that inference can briefly jump to an adjacent line while the pointer is in
 * the gap between two spans. Requiring a real span hit keeps the focus caret on
 * the last glyph the pointer touched until it reaches another text item.
 */
export function textCaretAtPoint(
  container: HTMLElement,
  clientX: number,
  clientY: number,
  previousSpan?: HTMLSpanElement | null,
): TextCaretPoint | null {
  const candidates: Array<{ span: HTMLSpanElement; rect: DOMRect }> = []
  const seen = new Set<HTMLSpanElement>()

  for (const element of document.elementsFromPoint(clientX, clientY)) {
    const span = element.closest<HTMLSpanElement>('[data-text-item="true"]')
    if (!span || seen.has(span) || !container.contains(span)) continue
    const rect = span.getBoundingClientRect()
    if (rect.width < 0.5 || rect.height < 1) continue
    // elementsFromPoint can include an ancestor after transforms or clipping.
    // Keep a small sub-pixel allowance but reject neighbouring visual lines.
    if (clientX < rect.left - 0.75 || clientX > rect.right + 0.75
      || clientY < rect.top - 0.75 || clientY > rect.bottom + 0.75) continue
    seen.add(span)
    candidates.push({ span, rect })
  }

  if (!candidates.length) return null

  const caretAt = (span: HTMLSpanElement, rect: DOMRect): TextCaretPoint | null => {
    const node = span.firstChild
    if (!(node instanceof Text)) return null
    const x = Math.max(rect.left + 0.01, Math.min(rect.right - 0.01, clientX))
    const y = Math.max(rect.top + 0.01, Math.min(rect.bottom - 0.01, clientY))
    return { node, offset: caretOffsetAtPoint(span, x, y), span }
  }

  // Hysteresis: while the pointer is still inside the previous focus span's
  // vertical band (with a small margin), keep that span instead of letting an
  // overlapping neighbour steal the caret mid-drag.
  if (previousSpan?.isConnected && container.contains(previousSpan)) {
    const rect = candidates.find((candidate) => candidate.span === previousSpan)?.rect
      ?? previousSpan.getBoundingClientRect()
    const margin = Math.max(1, rect.height * 0.25)
    if (rect.width >= 0.5 && rect.height >= 1
      && clientX >= rect.left - 0.75 && clientX <= rect.right + 0.75
      && clientY >= rect.top - margin && clientY <= rect.bottom + margin) {
      const caret = caretAt(previousSpan, rect)
      if (caret) return caret
    }
  }

  // Transformed PDF spans can overlap slightly. Prefer the spans whose band
  // actually contains the pointer; only rank by normalized centre distance
  // beyond that, because that ranking alone flip-flops in overlap bands.
  const banded = candidates.filter(({ rect }) => clientY >= rect.top && clientY <= rect.bottom)
  const pool = banded.length ? banded : candidates
  pool.sort((a, b) => {
    const aDistance = Math.abs(clientY - (a.rect.top + a.rect.bottom) / 2) / a.rect.height
    const bDistance = Math.abs(clientY - (b.rect.top + b.rect.bottom) / 2) / b.rect.height
    return aDistance - bDistance || a.rect.height - b.rect.height
  })
  return caretAt(pool[0].span, pool[0].rect)
}
