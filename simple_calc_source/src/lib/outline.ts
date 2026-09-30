/**
 * Row/column grouping (Excel's outline): levels live in `rowProperties` / `columnProperties`
 * as `outlineLevel` (1-7) with `collapsed` on the summary row, and collapsing hides the detail
 * rows through `hiddenRows` / `hiddenCols` (1-based), exactly as XLSX stores them. Indexes in
 * this module are 0-based. Functions mutate the sheet (an immer draft in the app).
 */
import type { SheetData } from '../spreadsheet-types'

export type OutlineAxis = 'row' | 'column'

export interface OutlineGroup {
  level: number
  start: number
  end: number
  /** The summary row/column the +/- button sits on (null at the sheet edge). */
  summary: number | null
  collapsed: boolean
}

export const MAX_OUTLINE_LEVEL = 7

type PropertyRecord = Record<string, Record<string, unknown>>

function properties(sheet: SheetData, axis: OutlineAxis): PropertyRecord | undefined {
  return axis === 'row' ? sheet.rowProperties : sheet.columnProperties
}

function ensureProperties(sheet: SheetData, axis: OutlineAxis): PropertyRecord {
  if (axis === 'row') return (sheet.rowProperties ||= {})
  return (sheet.columnProperties ||= {})
}

function hiddenList(sheet: SheetData, axis: OutlineAxis): number[] {
  return (axis === 'row' ? sheet.hiddenRows : sheet.hiddenCols) || []
}

function setHidden(sheet: SheetData, axis: OutlineAxis, values: Set<number>) {
  const sorted = [...values].sort((a, b) => a - b)
  if (axis === 'row') sheet.hiddenRows = sorted
  else sheet.hiddenCols = sorted
}

export function outlineLevelMap(sheet: SheetData, axis: OutlineAxis): Map<number, number> {
  const map = new Map<number, number>()
  for (const [key, record] of Object.entries(properties(sheet, axis) || {})) {
    const level = Math.min(MAX_OUTLINE_LEVEL, Math.max(0, Math.floor(Number(record?.outlineLevel) || 0)))
    const index = Number(key) - 1
    if (level > 0 && Number.isInteger(index) && index >= 0) map.set(index, level)
  }
  return map
}

export function maxOutlineLevel(sheet: SheetData, axis: OutlineAxis): number {
  let max = 0
  for (const level of outlineLevelMap(sheet, axis).values()) max = Math.max(max, level)
  return max
}

/** Excel's default puts summary rows below and summary columns to the right of their detail. */
export function summaryAfter(sheet: SheetData, axis: OutlineAxis): boolean {
  const outline = (sheet as SheetData & { outline?: { above?: boolean; left?: boolean } }).outline
  const props = (sheet.properties as { outlineProperties?: { summaryBelow?: boolean; summaryRight?: boolean } } | undefined)?.outlineProperties
  if (axis === 'row') return !(outline?.above || props?.summaryBelow === false)
  return !(outline?.left || props?.summaryRight === false)
}

export function outlineGroups(sheet: SheetData, axis: OutlineAxis): OutlineGroup[] {
  const levels = outlineLevelMap(sheet, axis)
  if (!levels.size) return []
  const hidden = new Set(hiddenList(sheet, axis).map((value) => Number(value) - 1))
  const after = summaryAfter(sheet, axis)
  const indexes = [...levels.keys()].sort((a, b) => a - b)
  const max = Math.max(...levels.values())
  const groups: OutlineGroup[] = []
  for (let level = 1; level <= max; level += 1) {
    let start = -1
    let previous = -2
    const close = () => {
      if (start < 0) return
      const end = previous
      const summary = after ? end + 1 : start - 1
      let collapsed = true
      for (let index = start; index <= end; index += 1) if (!hidden.has(index)) { collapsed = false; break }
      groups.push({ level, start, end, summary: summary >= 0 ? summary : null, collapsed })
      start = -1
    }
    for (const index of indexes) {
      if ((levels.get(index) || 0) < level) { close(); previous = index; continue }
      if (start >= 0 && index !== previous + 1) close()
      if (start < 0) start = index
      previous = index
    }
    close()
  }
  return groups
}

/** Group (delta 1) or ungroup (delta -1) a block of rows or columns. Returns true on change. */
export function changeOutline(sheet: SheetData, axis: OutlineAxis, start: number, end: number, delta: 1 | -1): boolean {
  const record = ensureProperties(sheet, axis)
  let changed = false
  for (let index = Math.max(0, start); index <= end; index += 1) {
    const key = String(index + 1)
    const current = Math.max(0, Math.floor(Number(record[key]?.outlineLevel) || 0))
    const next = Math.min(MAX_OUTLINE_LEVEL, Math.max(0, current + delta))
    if (next === current) continue
    changed = true
    const entry = { ...(record[key] || {}) }
    if (next) entry.outlineLevel = next
    else delete entry.outlineLevel
    if (Object.keys(entry).length) record[key] = entry
    else delete record[key]
  }
  return changed
}

function setCollapsedFlag(sheet: SheetData, axis: OutlineAxis, index: number | null, collapsed: boolean) {
  if (index === null) return
  const record = ensureProperties(sheet, axis)
  const key = String(index + 1)
  const entry = { ...(record[key] || {}) }
  if (collapsed) entry.collapsed = true
  else delete entry.collapsed
  if (Object.keys(entry).length) record[key] = entry
  else delete record[key]
}

export function collapseGroup(sheet: SheetData, axis: OutlineAxis, group: OutlineGroup) {
  const hidden = new Set(hiddenList(sheet, axis))
  for (let index = group.start; index <= group.end; index += 1) hidden.add(index + 1)
  setHidden(sheet, axis, hidden)
  setCollapsedFlag(sheet, axis, group.summary, true)
}

/** Show a group's detail, keeping nested groups that are themselves collapsed (and filtered rows) hidden. */
export function expandGroup(sheet: SheetData, axis: OutlineAxis, group: OutlineGroup) {
  const groups = outlineGroups(sheet, axis)
  const filtered = axis === 'row' ? new Set(sheet.filteredRows || []) : new Set<number>()
  const stayHidden = new Set<number>()
  for (const inner of groups) {
    if (inner.level <= group.level || inner.start < group.start || inner.end > group.end) continue
    const flagged = inner.summary !== null && properties(sheet, axis)?.[String(inner.summary + 1)]?.collapsed === true
    if (flagged) for (let index = inner.start; index <= inner.end; index += 1) stayHidden.add(index)
  }
  const hidden = new Set(hiddenList(sheet, axis))
  for (let index = group.start; index <= group.end; index += 1) {
    if (!stayHidden.has(index) && !filtered.has(index + 1)) hidden.delete(index + 1)
  }
  setHidden(sheet, axis, hidden)
  setCollapsedFlag(sheet, axis, group.summary, false)
}

export function toggleGroup(sheet: SheetData, axis: OutlineAxis, group: OutlineGroup) {
  if (group.collapsed) expandGroup(sheet, axis, group)
  else collapseGroup(sheet, axis, group)
}

/** Excel's level buttons: show everything above `level`, hide deeper detail. */
export function showOutlineLevel(sheet: SheetData, axis: OutlineAxis, level: number) {
  const levels = outlineLevelMap(sheet, axis)
  const filtered = axis === 'row' ? new Set(sheet.filteredRows || []) : new Set<number>()
  const hidden = new Set(hiddenList(sheet, axis))
  for (const [index, value] of levels) {
    if (value >= level) hidden.add(index + 1)
    else if (!filtered.has(index + 1)) hidden.delete(index + 1)
  }
  setHidden(sheet, axis, hidden)
  for (const group of outlineGroups(sheet, axis)) setCollapsedFlag(sheet, axis, group.summary, group.level >= level)
}

export function clearOutline(sheet: SheetData, axis: OutlineAxis) {
  const record = properties(sheet, axis)
  if (!record) return
  for (const [key, entry] of Object.entries(record)) {
    const next = { ...entry }
    delete next.outlineLevel
    delete next.collapsed
    if (Object.keys(next).length) record[key] = next
    else delete record[key]
  }
}

/** The group a row/column belongs to at its deepest level, if any. */
export function innermostGroupAt(groups: OutlineGroup[], index: number): OutlineGroup | undefined {
  let found: OutlineGroup | undefined
  for (const group of groups) if (index >= group.start && index <= group.end && (!found || group.level > found.level)) found = group
  return found
}
