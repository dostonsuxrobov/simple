// src/advanced/memory.ts (WP3)
// Memory estimates and the refusal policy for the Advanced editor (design 5.7).
//   - The whole session is budgeted at LIMITS.documentBudgetBytes (3 GiB): layer surfaces, masks, derived
//     caches (pyramid levels, composite tiles), history, plus what the host reports for the Simple canvas
//     and its undo stack (setExternalBytes).
//   - ensureMemory() is called before operations that allocate a lot (new layer, duplicate, paste, image
//     size up, PSD import): it frees rebuildable caches first, then trims the oldest history steps, and only
//     then refuses with a clear message. Nothing the user made is ever dropped to make room.
// Pure and DOM-free; caches register themselves here so the store can report them without importing them.
import type { DocumentState, DocumentStore, Layer, MemoryUsage } from './types.ts'
import { LIMITS } from './types.ts'

// ---------------------------------------------------------------------------------------------
// Cache registry
// ---------------------------------------------------------------------------------------------

export interface CacheReporter {
  readonly name: string
  /** Bytes held right now. */
  bytes(): number
  /** Drops what can be rebuilt; called under memory pressure. */
  trim?(): void
}

const reporters = new Set<CacheReporter>()
let external = 0

/** Registers a rebuildable cache; returns the unregister function. */
export function registerCache(reporter: CacheReporter): () => void {
  reporters.add(reporter)
  return () => { reporters.delete(reporter) }
}

/** Bytes held by every registered cache. */
export function cacheBytes(): number {
  let total = 0
  for (const reporter of reporters) {
    try {
      const bytes = reporter.bytes()
      if (Number.isFinite(bytes) && bytes > 0) total += bytes
    } catch {
      // A broken reporter must not break memory accounting.
    }
  }
  return total
}

/** Asks every cache to drop what it can rebuild. Returns the bytes freed (estimate). */
export function trimCaches(): number {
  const before = cacheBytes()
  for (const reporter of reporters) {
    try {
      reporter.trim?.()
    } catch {
      // Ignore: trimming is best effort.
    }
  }
  return Math.max(0, before - cacheBytes())
}

/** Host-reported bytes outside the Advanced document (the Simple canvas and its undo stack). */
export function setExternalBytes(bytes: number): void {
  external = Number.isFinite(bytes) && bytes > 0 ? Math.round(bytes) : 0
}

export function externalBytes(): number {
  return external
}

// ---------------------------------------------------------------------------------------------
// Estimates
// ---------------------------------------------------------------------------------------------

/** Pixel and mask bytes one layer holds (allocated tiles only). */
export function layerBytes(layer: Layer): { readonly pixels: number; readonly mask: number } {
  let pixels = 0
  if (layer.kind === 'raster') pixels = layer.surface.byteSize
  else if (layer.kind === 'text' || layer.kind === 'shape') pixels = layer.raster.surface.byteSize
  const mask = layer.mask ? layer.mask.surface.byteSize : 0
  return { pixels, mask }
}

/** Layer and mask bytes of a whole document. */
export function documentBytes(state: Pick<DocumentState, 'layers'>): { readonly layers: number; readonly masks: number } {
  let layers = 0
  let masks = 0
  for (const layer of state.layers) {
    const bytes = layerBytes(layer)
    layers += bytes.pixels
    masks += bytes.mask
  }
  return { layers, masks }
}

/** Worst-case bytes of a fully covered width x height RGBA layer (whole tiles, as allocated). */
export function estimateRasterBytes(width: number, height: number): number {
  if (!(width > 0 && height > 0)) return 0
  const tiles = Math.ceil(width / 256) * Math.ceil(height / 256)
  return tiles * 256 * 256 * 4
}

/** Worst-case bytes of a fully painted width x height mask. */
export function estimateMaskBytes(width: number, height: number): number {
  return estimateRasterBytes(width, height) / 4
}

/** Builds a MemoryUsage record (budget = LIMITS.documentBudgetBytes). */
export function usageOf(parts: { readonly layers: number; readonly masks: number; readonly history: number; readonly caches?: number }): MemoryUsage {
  const caches = parts.caches ?? cacheBytes()
  const total = parts.layers + parts.masks + caches + parts.history + external
  return { layers: parts.layers, masks: parts.masks, caches, history: parts.history, total, budget: LIMITS.documentBudgetBytes }
}

// ---------------------------------------------------------------------------------------------
// Refusal policy
// ---------------------------------------------------------------------------------------------

export type MemoryCheck =
  | { readonly ok: true; readonly freedCaches: number; readonly freedHistory: number }
  | { readonly ok: false; readonly message: string; readonly neededBytes: number; readonly availableBytes: number }

/** "412 MB", "1.2 GB". */
export function formatBytes(bytes: number): string {
  const value = Math.max(0, Number(bytes) || 0)
  if (value >= 1024 ** 3) return `${(value / 1024 ** 3).toFixed(1)} GB`
  if (value >= 1024 ** 2) return `${Math.round(value / 1024 ** 2)} MB`
  if (value >= 1024) return `${Math.round(value / 1024)} KB`
  return `${Math.round(value)} bytes`
}

function refusal(usage: MemoryUsage, extra: number, action: string): MemoryCheck {
  const available = Math.max(0, usage.budget - usage.total)
  return {
    ok: false,
    neededBytes: extra,
    availableBytes: available,
    message: `There isn't enough memory to ${action} (needs about ${formatBytes(extra)}, ${formatBytes(available)} free). `
      + 'Merge or delete some layers, or make the image smaller, then try again.',
  }
}

/**
 * Makes room for `extraBytes` more: frees rebuildable caches, then trims the oldest history steps (the
 * step at the cursor always stays undoable), then refuses. History is only trimmed when that actually
 * makes enough room, so a refused operation never costs the user undo steps. `action` names the
 * operation in the message, e.g. "duplicate this layer".
 */
export function ensureMemory(store: DocumentStore, extraBytes: number, action = 'do this'): MemoryCheck {
  const extra = Math.max(0, Number(extraBytes) || 0)
  let usage = store.memoryUsage()
  if (usage.total + extra <= usage.budget) return { ok: true, freedCaches: 0, freedHistory: 0 }
  const freedCaches = trimCaches()
  usage = store.memoryUsage()
  if (usage.total + extra <= usage.budget) return { ok: true, freedCaches, freedHistory: 0 }
  const history = store.history
  const state = history.getState()
  const kept = state.entries[state.cursor]?.bytes ?? 0
  const freeable = Math.max(0, state.totalBytes - kept)
  if (usage.total - freeable + extra > usage.budget) return refusal(usage, extra, action)
  const deficit = usage.total + extra - usage.budget
  // Lowering the budget evicts the oldest steps (then redo steps); restoring it does not bring them back.
  history.setBudget(Math.max(0, state.totalBytes - deficit), state.maxEntries)
  history.setBudget(state.budgetBytes, state.maxEntries)
  const freedHistory = Math.max(0, state.totalBytes - history.getState().totalBytes)
  usage = store.memoryUsage()
  if (usage.total + extra <= usage.budget) return { ok: true, freedCaches, freedHistory }
  return refusal(usage, extra, action)
}
