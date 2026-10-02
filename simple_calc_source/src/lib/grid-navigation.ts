/*
 * Excel-style grid navigation as pure functions (no React, no DOM), so the keyboard
 * behaviour can be tested on its own:
 *
 * - Enter / Tab move the active cell inside a multi-cell selection and wrap.
 * - A Tab run remembers its first column, so Enter returns to it on the next row.
 * - Ctrl+. cycles the active cell through the selection corners.
 * - Ctrl+Arrow runs to the edge of a data block, else to the next value, else to the sheet edge.
 * - The current region (Ctrl+Shift+8) and the progressive Ctrl+A (region first, then the sheet).
 * - The rendered extent of the virtual grid follows the data, the selection and the scroll.
 */

/** Last 0-based row and column of an Excel worksheet (XFD1048576). */
export const SHEET_LAST_ROW = 1_048_575
export const SHEET_LAST_COL = 16_383

export interface GridCoord { row: number; col: number }
export interface GridBounds { top: number; left: number; bottom: number; right: number }

export function boundsOf(anchor: GridCoord, focus: GridCoord): GridBounds {
  return {
    top: Math.min(anchor.row, focus.row),
    bottom: Math.max(anchor.row, focus.row),
    left: Math.min(anchor.col, focus.col),
    right: Math.max(anchor.col, focus.col),
  }
}

/**
 * Where Shift+navigation extends from. A plain selection extends its moving end (focus) and
 * keeps its anchor. A stored block (Ctrl+A, Shift+Space, Ctrl+Space, Go To) is anchored at the
 * block corner nearest the active cell and grows or shrinks from the opposite corner, as in
 * Excel: Shift+Space then Shift+Down selects two whole rows.
 */
export function selectionExtension(selection: { anchor: GridCoord; focus: GridCoord; range?: GridBounds }): { anchor: GridCoord; from: GridCoord } {
  if (!selection.range) return { anchor: selection.anchor, from: selection.focus }
  const { top, bottom, left, right } = selection.range
  const active = selection.focus
  const anchorRow = active.row - top <= bottom - active.row ? top : bottom
  const anchorCol = active.col - left <= right - active.col ? left : right
  return {
    anchor: { row: anchorRow, col: anchorCol },
    from: { row: anchorRow === top ? bottom : top, col: anchorCol === left ? right : left },
  }
}

export function sameBounds(a: GridBounds, b: GridBounds) {
  return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right
}

export function containsBounds(outer: GridBounds, inner: GridBounds) {
  return inner.top >= outer.top && inner.bottom <= outer.bottom && inner.left >= outer.left && inner.right <= outer.right
}

export function boundsContain(bounds: GridBounds, coord: GridCoord) {
  return coord.row >= bounds.top && coord.row <= bounds.bottom && coord.col >= bounds.left && coord.col <= bounds.right
}

/**
 * The next active cell inside a selection, as Excel moves it with Enter (`order: 'columns'`:
 * down the column, then the top of the next column) or Tab (`order: 'rows'`: along the row,
 * then the start of the next row). `backwards` is Shift+Enter / Shift+Tab. Both wrap around
 * the whole selection. `skip` rules out cells that cannot become active (hidden rows or
 * columns, cells covered by a merge); when every cell is skipped the active cell stays put.
 */
export function stepWithinBounds(
  bounds: GridBounds,
  active: GridCoord,
  order: 'rows' | 'columns',
  backwards: boolean,
  skip?: (row: number, col: number) => boolean,
): GridCoord {
  const height = bounds.bottom - bounds.top + 1
  const width = bounds.right - bounds.left + 1
  const total = height * width
  if (total <= 1) return active
  const row = Math.min(bounds.bottom, Math.max(bounds.top, active.row)) - bounds.top
  const col = Math.min(bounds.right, Math.max(bounds.left, active.col)) - bounds.left
  let index = order === 'rows' ? row * width + col : col * height + row
  const step = backwards ? -1 : 1
  // A long run of hidden rows inside a huge selection must not hang the keyboard.
  const limit = Math.min(total, 2_000_000)
  for (let attempt = 0; attempt < limit; attempt += 1) {
    index = (index + step + total) % total
    const next = order === 'rows'
      ? { row: bounds.top + Math.floor(index / width), col: bounds.left + (index % width) }
      : { row: bounds.top + (index % height), col: bounds.left + Math.floor(index / height) }
    if (!skip || !skip(next.row, next.col)) return next
  }
  return active
}

/**
 * Ctrl+. : moves the active cell clockwise to the next distinct corner of the selection
 * (top-left, top-right, bottom-right, bottom-left) and keeps the selected range. Returns the
 * new focus (active cell) and the opposite corner as the anchor.
 */
export function nextSelectionCorner(bounds: GridBounds, active: GridCoord): { anchor: GridCoord; focus: GridCoord } {
  const corners: GridCoord[] = [
    { row: bounds.top, col: bounds.left },
    { row: bounds.top, col: bounds.right },
    { row: bounds.bottom, col: bounds.right },
    { row: bounds.bottom, col: bounds.left },
  ]
  const same = (a: GridCoord, b: GridCoord) => a.row === b.row && a.col === b.col
  const at = corners.findIndex((corner) => same(corner, active))
  let focus = corners[0]
  if (at >= 0) {
    for (let offset = 1; offset <= 4; offset += 1) {
      const candidate = corners[(at + offset) % 4]
      if (!same(candidate, active)) { focus = candidate; break }
    }
  }
  const anchor = {
    row: focus.row === bounds.top ? bounds.bottom : bounds.top,
    col: focus.col === bounds.left ? bounds.right : bounds.left,
  }
  return { anchor, focus }
}

/** One column (or row) of a sheet, as Ctrl+Arrow sees it. */
export interface EdgeLane {
  /** Indices along the lane that hold a value (any order, duplicates allowed). */
  filled: Iterable<number>
  /** Hidden indices: never a stop, stepped over. */
  hidden?: ReadonlySet<number>
  /** Last valid index (the sheet edge). */
  last: number
}

/**
 * Excel's Ctrl+Arrow along one lane: inside a block of values it runs to the block's last
 * value; otherwise it jumps to the next value, or to the sheet edge when there is none
 * (Ctrl+Down in an empty column lands on row 1,048,576).
 */
export function edgeJumpIndex(lane: EdgeLane, position: number, step: 1 | -1): number {
  const hidden = lane.hidden
  const last = Math.max(0, lane.last)
  const visibleFilled = [...new Set(lane.filled)].filter((index) => index >= 0 && index <= last && !hidden?.has(index)).sort((a, b) => a - b)
  const filled = new Set(visibleFilled)
  const nextVisible = (index: number) => {
    let next = index + step
    while (next >= 0 && next <= last && hidden?.has(next)) next += step
    return next >= 0 && next <= last ? next : null
  }
  const start = Math.min(last, Math.max(0, position))
  const neighbour = nextVisible(start)
  if (filled.has(start) && neighbour !== null && filled.has(neighbour)) {
    let target = neighbour
    for (let next = nextVisible(target); next !== null && filled.has(next); next = nextVisible(target)) target = next
    return target
  }
  // First value strictly beyond the position in the step direction (binary search).
  let low = 0
  let high = visibleFilled.length
  while (low < high) {
    const middle = (low + high) >> 1
    if (visibleFilled[middle] <= start) low = middle + 1
    else high = middle
  }
  const found = step > 0 ? visibleFilled[low] : visibleFilled[(visibleFilled[low - 1] === start ? low - 2 : low - 1)]
  if (found !== undefined) return found
  // No value beyond: the sheet edge, or the last visible index before it.
  let edge = step > 0 ? last : 0
  while (hidden?.has(edge) && (step > 0 ? edge > start : edge < start)) edge -= step
  return edge
}

/** Filled cells indexed both ways, for region growth without rescanning the whole sheet. */
export interface FilledIndex {
  byRow: Map<number, number[]>
  byCol: Map<number, number[]>
}

export function indexFilledCells(coords: Iterable<GridCoord>): FilledIndex {
  const byRow = new Map<number, number[]>()
  const byCol = new Map<number, number[]>()
  for (const { row, col } of coords) {
    const columns = byRow.get(row)
    if (columns) columns.push(col)
    else byRow.set(row, [col])
    const rows = byCol.get(col)
    if (rows) rows.push(row)
    else byCol.set(col, [row])
  }
  for (const list of byRow.values()) list.sort((a, b) => a - b)
  for (const list of byCol.values()) list.sort((a, b) => a - b)
  return { byRow, byCol }
}

function anyBetween(sorted: number[] | undefined, low: number, high: number) {
  if (!sorted?.length) return false
  let lo = 0
  let hi = sorted.length
  while (lo < hi) {
    const middle = (lo + hi) >> 1
    if (sorted[middle] < low) lo = middle + 1
    else hi = middle
  }
  return lo < sorted.length && sorted[lo] <= high
}

/**
 * Excel's current region around `origin`: the rectangle bounded by blank rows and blank
 * columns (diagonal neighbours count). Returns null when the origin is a blank cell with no
 * value around it, which is when Ctrl+A selects the whole sheet straight away.
 */
export function currentRegion(index: FilledIndex, origin: GridCoord, limits: GridCoord = { row: SHEET_LAST_ROW, col: SHEET_LAST_COL }): GridBounds | null {
  const bounds: GridBounds = { top: origin.row, bottom: origin.row, left: origin.col, right: origin.col }
  let grew = true
  while (grew) {
    grew = false
    const left = Math.max(0, bounds.left - 1)
    const right = Math.min(limits.col, bounds.right + 1)
    const top = Math.max(0, bounds.top - 1)
    const bottom = Math.min(limits.row, bounds.bottom + 1)
    if (bounds.top > 0 && anyBetween(index.byRow.get(bounds.top - 1), left, right)) { bounds.top -= 1; grew = true }
    if (bounds.bottom < limits.row && anyBetween(index.byRow.get(bounds.bottom + 1), left, right)) { bounds.bottom += 1; grew = true }
    if (bounds.left > 0 && anyBetween(index.byCol.get(bounds.left - 1), top, bottom)) { bounds.left -= 1; grew = true }
    if (bounds.right < limits.col && anyBetween(index.byCol.get(bounds.right + 1), top, bottom)) { bounds.right += 1; grew = true }
  }
  const single = bounds.top === bounds.bottom && bounds.left === bounds.right
  if (single && !anyBetween(index.byRow.get(origin.row), origin.col, origin.col)) return null
  return bounds
}

/**
 * Progressive Ctrl+A: the current region first, then the whole sheet once the selection
 * already covers the region (or when there is no region around the active cell).
 */
export function progressiveSelectAll(current: GridBounds, region: GridBounds | null, sheet: GridBounds): GridBounds {
  if (!region || containsBounds(current, region)) return sheet
  return region
}

/**
 * How many rows (or columns) the virtual grid lays out: the used area plus a margin, the
 * selection plus a margin, and the scrolled-to area plus a page, so the grid behaves as
 * if it were unbounded while only the visible window is rendered. Never past the sheet
 * limit, never below the minimum a blank sheet shows.
 */
export function virtualExtent(options: {
  minimum: number
  used: number
  usedMargin: number
  selectionEnd: number
  selectionMargin: number
  viewEnd: number
  viewMargin: number
  limit: number
}): number {
  const wanted = Math.max(
    options.minimum,
    options.used + options.usedMargin,
    options.selectionEnd + 1 + options.selectionMargin,
    options.viewEnd + 1 + options.viewMargin,
  )
  return Math.max(Math.min(options.minimum, options.limit), Math.min(options.limit, Math.ceil(wanted)))
}

/** A row being typed with Tab: Enter goes to the next row under the column it started in. */
export interface TabRun {
  startCol: number
  /** Where the last Tab left the active cell; any other move ends the run. */
  expected: GridCoord
}

export function continueTabRun(run: TabRun | null, from: GridCoord, to: GridCoord): TabRun {
  const continuing = run && run.expected.row === from.row && run.expected.col === from.col
  return { startCol: continuing ? run.startCol : from.col, expected: { ...to } }
}

/** Target of Enter after a Tab run, or null when the run does not apply. */
export function enterAfterTabRun(run: TabRun | null, from: GridCoord, lastRow = SHEET_LAST_ROW): GridCoord | null {
  if (!run || run.expected.row !== from.row || run.expected.col !== from.col) return null
  return { row: Math.min(lastRow, from.row + 1), col: run.startCol }
}

/** A cell position in workbook order: sheet index, then row, then column (Find's order). */
export interface OrderedPosition { sheet: number; row: number; col: number }

function comparePositions(a: OrderedPosition, b: OrderedPosition) {
  return a.sheet - b.sheet || a.row - b.row || a.col - b.col
}

/**
 * Find Next from a cell: the index of the first position after `origin` (or at it, when
 * `inclusive`), wrapping to the first position. `positions` must be in workbook order.
 * Returns -1 for an empty list.
 */
export function firstPositionAfter(positions: readonly OrderedPosition[], origin: OrderedPosition, inclusive = false): number {
  if (!positions.length) return -1
  const index = positions.findIndex((position) => {
    const order = comparePositions(position, origin)
    return inclusive ? order >= 0 : order > 0
  })
  return index >= 0 ? index : 0
}

/** Find Previous from a cell: the last position before `origin`, wrapping to the last one. */
export function lastPositionBefore(positions: readonly OrderedPosition[], origin: OrderedPosition): number {
  if (!positions.length) return -1
  for (let index = positions.length - 1; index >= 0; index -= 1) {
    if (comparePositions(positions[index], origin) < 0) return index
  }
  return positions.length - 1
}
