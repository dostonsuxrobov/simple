// src/advanced/history.ts (WP3)
// Undo history of the Advanced editor (design 5.4). The DocumentStore records entries; tools and commands
// never touch history directly.
//   - Entries hold inverse operations. Pixel edits are copy-on-write tiles: the first touch of a tile stores
//     a copy (or null when it did not exist); undo swaps stored and live tiles, redo swaps back. Whole-buffer
//     operations swap surfaces by reference. Metadata changes keep the immutable before/after objects.
//   - Budget: at most `maxEntries` undoable steps and `budgetBytes`. The oldest steps are evicted first; the
//     base entry ("Open" / "Advanced editor") has no ops and is relabelled once steps were trimmed. The
//     step at the cursor is never evicted, so the last action stays undoable even when it alone is larger
//     than the budget.
//   - Revisions: every entry knows the host content revision before and after it; undo/redo/jump set the
//     host revision of the state they land on, which drives the Modified badge.
//   - Byte accounting is exact for what history keeps alive and is recomputed whenever an entry is
//     undone or redone (an undone insert, for example, keeps its layer alive only through history).
// Pure and DOM-free.
import type { Point } from '../imaging/types.ts'
import type {
  AdvancedHost,
  EditTarget,
  History,
  HistoryEntryInfo,
  HistoryIcon,
  HistoryState,
  Layer,
  LayerId,
  LayerMask,
  SelectionSnapshot,
  TiledMask,
  TiledSurface,
} from './types.ts'
import { COALESCE_MS, LIMITS } from './types.ts'
import { layerBytes } from './memory.ts'
import { snapshotBytes } from './selection.ts'

export const TRIMMED_LABEL = 'Earlier steps were trimmed'
/** Fixed bookkeeping cost charged per entry and per small op. */
export const ENTRY_OVERHEAD_BYTES = 256
export const SMALL_OP_BYTES = 64

export interface ActiveRef {
  readonly id: LayerId | null
  readonly target: EditTarget
}

export interface CanvasInfo {
  readonly width: number
  readonly height: number
  readonly ppi: number
}

export type TileData = Uint8ClampedArray | Uint8Array

export interface TilesOp {
  readonly kind: 'tiles'
  readonly layerId: LayerId
  readonly target: EditTarget
  /** The live surface the tiles belong to (undo swaps into exactly this object). */
  readonly surface: TiledSurface | TiledMask
  /** tileKey -> the tile not currently live (null = absent). Swapped on undo and redo. */
  readonly stored: Map<number, TileData | null>
}

export interface SurfaceOp {
  readonly kind: 'surface'
  readonly layerId: LayerId
  readonly target: EditTarget
  /** The surface not currently live, and the offset that goes with it (swapped on undo and redo). */
  other: TiledSurface | TiledMask
  otherOffset: Point
}

export type HistoryOp =
  | TilesOp
  | SurfaceOp
  | { readonly kind: 'insert' | 'remove'; readonly layer: Layer; readonly index: number }
  | { readonly kind: 'move'; readonly layerId: LayerId; readonly from: number; readonly to: number }
  | { readonly kind: 'layer'; readonly before: Layer; readonly after: Layer }
  | { readonly kind: 'mask'; readonly layerId: LayerId; readonly before: LayerMask | null; readonly after: LayerMask | null }
  | { readonly kind: 'active'; readonly before: ActiveRef; readonly after: ActiveRef }
  | { readonly kind: 'selection'; readonly before: SelectionSnapshot; readonly after: SelectionSnapshot }
  | { readonly kind: 'canvas'; readonly before: CanvasInfo; readonly after: CanvasInfo }

export interface HistoryEntry {
  readonly id: number
  label: string
  readonly icon: HistoryIcon
  bytes: number
  readonly revisionBefore: number
  revisionAfter: number
  ops: HistoryOp[]
  readonly coalesceKey: string | null
  /** Time of the last transaction merged into this entry (coalescing window). */
  time: number
}

export type HistoryDirection = 'undo' | 'redo'

export interface HistoryOptions {
  readonly baseLabel: string
  readonly baseRevision: number
  readonly host: Pick<AdvancedHost, 'setRevision'>
  /** Applies one entry: its ops in reverse order for undo, in order for redo. */
  readonly apply: (entry: HistoryEntry, direction: HistoryDirection) => void
  /** Runs before any navigation or reset (the store commits an open stroke here). */
  readonly beforeNavigate?: () => void
  /** Runs after undo/redo/jump moved the cursor, once the host revision was set. */
  readonly afterNavigate?: (revision: number) => void
  /** Runs after any change of the history state (record, navigation, trim, reset). */
  readonly onChange?: () => void
  readonly budgetBytes?: number
  readonly maxEntries?: number
}

export interface RecordInput {
  readonly label: string
  readonly icon: HistoryIcon
  readonly ops: readonly HistoryOp[]
  readonly revisionBefore: number
  readonly revisionAfter: number
  readonly coalesceKey?: string | null
  readonly time: number
}

export interface HistoryController extends History {
  /** Appends an entry (dropping the redo tail), then enforces the budget. */
  record(input: RecordInput): HistoryEntry
  /** The newest entry when a transaction with `key` at `time` should merge into it, else null. */
  coalesceTarget(key: string | null | undefined, time: number): HistoryEntry | null
  /** Merges more ops into `entry` (the newest entry). */
  extend(entry: HistoryEntry, ops: readonly HistoryOp[], revisionAfter: number, time: number): void
  /** Host revision of the current state. */
  currentRevision(): number
  /** Every entry, oldest first (read-only use). */
  entriesForDebug(): readonly HistoryEntry[]
  dispose(): void
}

// ---------------------------------------------------------------------------------------------
// Byte accounting
// ---------------------------------------------------------------------------------------------

function totalLayerBytes(layer: Layer): number {
  const bytes = layerBytes(layer)
  return bytes.pixels + bytes.mask + SMALL_OP_BYTES
}

/** Resources `held` keeps alive that `live` does not share (raster caches, masks). */
function exclusiveLayerBytes(held: Layer, live: Layer): number {
  let bytes = 0
  if ((held.kind === 'text' || held.kind === 'shape') && held.raster
    && !((live.kind === 'text' || live.kind === 'shape') && live.raster && live.raster.surface === held.raster.surface)) {
    bytes += held.raster.surface.byteSize
  }
  if (held.kind === 'raster' && !(live.kind === 'raster' && live.surface === held.surface)) bytes += held.surface.byteSize
  if (held.mask && held.mask.surface !== live.mask?.surface) bytes += held.mask.surface.byteSize
  return bytes
}

/**
 * Bytes `op` keeps alive. `applied` = the op's effect is in the live document (entries at or before the
 * cursor); otherwise it was undone and history holds what redo needs.
 */
export function opBytes(op: HistoryOp, applied: boolean): number {
  switch (op.kind) {
    case 'tiles': {
      let bytes = SMALL_OP_BYTES
      for (const tile of op.stored.values()) if (tile) bytes += tile.byteLength
      return bytes
    }
    case 'surface':
      return SMALL_OP_BYTES + op.other.byteSize
    case 'insert':
      return applied ? SMALL_OP_BYTES : totalLayerBytes(op.layer)
    case 'remove':
      return applied ? totalLayerBytes(op.layer) : SMALL_OP_BYTES
    case 'layer':
      return SMALL_OP_BYTES + (applied ? exclusiveLayerBytes(op.before, op.after) : exclusiveLayerBytes(op.after, op.before))
    case 'mask': {
      const held = applied ? op.before : op.after
      const live = applied ? op.after : op.before
      return SMALL_OP_BYTES + (held && held.surface !== live?.surface ? held.surface.byteSize : 0)
    }
    case 'selection':
      return snapshotBytes(op.before) + snapshotBytes(op.after)
    default:
      return SMALL_OP_BYTES
  }
}

function entryBytes(entry: HistoryEntry, applied: boolean): number {
  if (!entry.ops.length) return 0
  let bytes = ENTRY_OVERHEAD_BYTES
  for (const op of entry.ops) bytes += opBytes(op, applied)
  return bytes
}

/** True when an op changes what Save/Export would write (selection and active-layer changes do not). */
export function opAffectsOutput(op: HistoryOp): boolean {
  return op.kind !== 'selection' && op.kind !== 'active'
}

/** Appends `next` to `ops`, collapsing consecutive metadata changes of the same layer into one op. */
function appendOps(ops: HistoryOp[], next: readonly HistoryOp[]): void {
  for (const op of next) {
    const last = ops[ops.length - 1]
    if (last && last.kind === 'layer' && op.kind === 'layer' && last.after === op.before) {
      ops[ops.length - 1] = { kind: 'layer', before: last.before, after: op.after }
      continue
    }
    ops.push(op)
  }
}

// ---------------------------------------------------------------------------------------------
// The history controller
// ---------------------------------------------------------------------------------------------

export function createHistory(options: HistoryOptions): HistoryController {
  let nextId = 1
  let entries: HistoryEntry[] = []
  let cursor = 0
  let totalBytes = 0
  let trimmed = false
  let budgetBytes = sanitizeBudget(options.budgetBytes ?? LIMITS.historyBudgetBytes)
  let maxEntries = sanitizeMax(options.maxEntries ?? LIMITS.historyMaxEntries)
  let toggleFrom: number | null = null
  let navigating = false
  let cached: HistoryState | null = null
  let disposed = false

  function sanitizeBudget(bytes: number): number {
    return Number.isFinite(bytes) && bytes >= 0 ? Math.floor(bytes) : LIMITS.historyBudgetBytes
  }

  function sanitizeMax(count: number): number {
    return Number.isFinite(count) && count >= 1 ? Math.floor(count) : LIMITS.historyMaxEntries
  }

  function makeBase(label: string, revision: number): HistoryEntry {
    const entry: HistoryEntry = {
      id: nextId,
      label: label || 'Open',
      icon: 'open',
      bytes: 0,
      revisionBefore: revision,
      revisionAfter: revision,
      ops: [],
      coalesceKey: null,
      time: 0,
    }
    nextId += 1
    return entry
  }

  entries = [makeBase(options.baseLabel, options.baseRevision)]

  function changed(): void {
    cached = null
    options.onChange?.()
  }

  function recount(entry: HistoryEntry, applied: boolean): void {
    const bytes = entryBytes(entry, applied)
    totalBytes += bytes - entry.bytes
    entry.bytes = bytes
  }

  /** Evicts the oldest undoable steps, then redo steps, until the budget holds; the cursor step stays. */
  function enforce(): void {
    let evicted = false
    while (cursor >= 2 && (entries.length - 1 > maxEntries || totalBytes > budgetBytes)) {
      const oldest = entries[1]
      entries.splice(1, 1)
      cursor -= 1
      totalBytes -= oldest.bytes
      // The base now stands for the state after the evicted step.
      entries[0].revisionAfter = oldest.revisionAfter
      evicted = true
      if (toggleFrom !== null) toggleFrom -= 1
    }
    while (entries.length - 1 > cursor && (entries.length - 1 > maxEntries || totalBytes > budgetBytes)) {
      const dropped = entries.pop() as HistoryEntry
      totalBytes -= dropped.bytes
    }
    if (evicted) {
      trimmed = true
      entries[0].label = TRIMMED_LABEL
    }
    if (toggleFrom !== null && (toggleFrom < 1 || toggleFrom >= entries.length)) toggleFrom = null
  }

  function assertUsable(): void {
    if (disposed) throw new Error('This document has been closed.')
  }

  function navigate(target: number, keepToggle = false): void {
    assertUsable()
    if (navigating) throw new Error('History is already changing.')
    options.beforeNavigate?.()
    const index = Math.max(0, Math.min(entries.length - 1, Math.floor(target)))
    if (!keepToggle) toggleFrom = null
    if (index === cursor) return
    navigating = true
    try {
      while (cursor > index) {
        const entry = entries[cursor]
        options.apply(entry, 'undo')
        cursor -= 1
        recount(entry, false)
      }
      while (cursor < index) {
        const entry = entries[cursor + 1]
        options.apply(entry, 'redo')
        cursor += 1
        recount(entry, true)
      }
    } finally {
      navigating = false
    }
    const revision = entries[cursor].revisionAfter
    options.host.setRevision(revision)
    options.afterNavigate?.(revision)
    changed()
  }

  const controller: HistoryController = {
    getState(): HistoryState {
      if (!cached) {
        cached = Object.freeze({
          entries: Object.freeze(entries.map((entry): HistoryEntryInfo => Object.freeze({
            id: entry.id,
            label: entry.label,
            icon: entry.icon,
            bytes: entry.bytes,
          }))),
          cursor,
          totalBytes,
          budgetBytes,
          maxEntries,
          trimmed,
        })
      }
      return cached
    },
    canUndo: () => !disposed && cursor > 0,
    canRedo: () => !disposed && cursor < entries.length - 1,
    // Each navigation first lets the store commit an open stroke (it becomes the newest step), so a
    // Ctrl+Z pressed mid-stroke undoes that stroke instead of losing it or corrupting the stack.
    undo() {
      assertUsable()
      options.beforeNavigate?.()
      if (cursor > 0) navigate(cursor - 1)
    },
    redo() {
      assertUsable()
      options.beforeNavigate?.()
      if (cursor < entries.length - 1) navigate(cursor + 1)
    },
    jumpTo(index: number) {
      assertUsable()
      if (!Number.isFinite(index)) return
      navigate(index)
    },
    toggleLast() {
      assertUsable()
      options.beforeNavigate?.()
      if (toggleFrom !== null && cursor === toggleFrom - 1 && toggleFrom < entries.length) {
        navigate(toggleFrom)
        return
      }
      if (cursor > 0) {
        const from = cursor
        navigate(cursor - 1, true)
        toggleFrom = from
        return
      }
      if (cursor < entries.length - 1) navigate(cursor + 1)
    },
    reset(label: string) {
      assertUsable()
      options.beforeNavigate?.()
      const revision = entries[cursor].revisionAfter
      entries = [makeBase(label, revision)]
      cursor = 0
      totalBytes = 0
      trimmed = false
      toggleFrom = null
      changed()
    },
    setBudget(bytes: number, count: number) {
      budgetBytes = sanitizeBudget(bytes)
      maxEntries = sanitizeMax(count)
      enforce()
      changed()
    },
    record(input: RecordInput): HistoryEntry {
      assertUsable()
      if (navigating) throw new Error('Cannot record while history is changing.')
      while (entries.length - 1 > cursor) {
        const dropped = entries.pop() as HistoryEntry
        totalBytes -= dropped.bytes
      }
      const entry: HistoryEntry = {
        id: nextId,
        label: input.label || 'Edit',
        icon: input.icon,
        bytes: 0,
        revisionBefore: input.revisionBefore,
        revisionAfter: input.revisionAfter,
        ops: [],
        coalesceKey: input.coalesceKey ?? null,
        time: input.time,
      }
      nextId += 1
      appendOps(entry.ops, input.ops)
      entries.push(entry)
      cursor = entries.length - 1
      recount(entry, true)
      toggleFrom = null
      enforce()
      changed()
      return entry
    },
    coalesceTarget(key: string | null | undefined, time: number): HistoryEntry | null {
      if (!key || disposed) return null
      if (cursor !== entries.length - 1 || cursor === 0) return null
      const entry = entries[cursor]
      if (entry.coalesceKey !== key) return null
      if (!(time - entry.time <= COALESCE_MS) || time < entry.time) return null
      return entry
    },
    extend(entry: HistoryEntry, ops: readonly HistoryOp[], revisionAfter: number, time: number) {
      assertUsable()
      if (entries[cursor] !== entry || cursor !== entries.length - 1) throw new Error('Only the newest history step can be extended.')
      appendOps(entry.ops, ops)
      entry.revisionAfter = revisionAfter
      entry.time = time
      recount(entry, true)
      toggleFrom = null
      enforce()
      changed()
    },
    currentRevision: () => entries[cursor].revisionAfter,
    entriesForDebug: () => entries,
    dispose() {
      if (disposed) return
      const revision = entries[cursor].revisionAfter
      disposed = true
      // Release every stored tile and surface.
      entries = [makeBase(entries[0].label, revision)]
      cursor = 0
      totalBytes = 0
      toggleFrom = null
      cached = null
    },
  }
  return controller
}
