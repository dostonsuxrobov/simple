// src/advanced/panels/HistoryPanel.tsx (WP6)
// The History panel (design 5.4, 5.14):
//   - every history state with its tool icon; a click (or Up / Down / Home / End in the list) jumps to it,
//     exactly like repeated Undo / Redo, so nothing is lost: later states stay listed (dimmed) until a new
//     edit replaces them, as in Photoshop;
//   - a notice once old steps were trimmed to keep memory in budget, and the steps / memory in the footer;
//   - snapshots (the camera button): a full copy of the document at that moment, listed above the steps.
//     Clicking a snapshot brings the document back to it as one new, undoable step ("Revert to
//     Snapshot 1"), so reverting never throws away the steps in between. Snapshots live as long as the
//     editor session, count towards the editor's memory budget (a memory.ts reporter) and are released
//     when the editor closes (releaseHistorySnapshots, called by AdvancedEditor).
import { useCallback, useEffect, useRef, useSyncExternalStore } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react'
import {
  Blend,
  Brush,
  Camera,
  ClipboardPaste,
  Crop,
  Eraser,
  FolderOpen,
  Hand,
  Image as ImageIcon,
  Lasso,
  LassoSelect,
  Layers,
  Maximize2,
  Merge,
  Move,
  PaintBucket,
  Pipette,
  Scaling,
  Shapes,
  SlidersHorizontal,
  Sparkles,
  SquareDashed,
  CircleDashed,
  Stamp,
  Bandage,
  Trash2,
  Type as TypeIcon,
  WandSparkles,
  ZoomIn,
} from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import type { PixelBuffer } from '../../imaging/types.ts'
import type { DocumentState, DocumentStore, EditTarget, HistoryIcon, Layer, SelectionSnapshot } from '../types.ts'
import { newLayerId } from '../document.ts'
import { compositeRect } from '../composite.ts'
import { maxPyramidLevel } from '../pyramid.ts'
import { documentBytes, ensureMemory, formatBytes, registerCache } from '../memory.ts'
import { restoreSelection, snapshotBytes, snapshotSelection } from '../selection.ts'
import { nextSelectionVersion } from '../tools/shared.ts'
import { commandsBusy } from '../commands.ts'
import type { PanelContext, PanelDefinition, PanelProps } from './PanelHost.tsx'
import { useHistoryState } from './PanelHost.tsx'
import { revealInList } from './AdjustmentControls.tsx'

// ---------------------------------------------------------------------------------------------
// Icons
// ---------------------------------------------------------------------------------------------

const HISTORY_ICONS: Readonly<Record<string, LucideIcon>> = Object.freeze({
  open: FolderOpen,
  layer: Layers,
  adjustment: SlidersHorizontal,
  filter: Sparkles,
  image: Maximize2,
  selection: SquareDashed,
  paste: ClipboardPaste,
  text: TypeIcon,
  transform: Scaling,
  merge: Merge,
  fill: PaintBucket,
  move: Move,
  'marquee-rect': SquareDashed,
  'marquee-ellipse': CircleDashed,
  lasso: Lasso,
  'lasso-polygon': LassoSelect,
  'magic-wand': WandSparkles,
  crop: Crop,
  eyedropper: Pipette,
  'spot-healing': Bandage,
  brush: Brush,
  'clone-stamp': Stamp,
  eraser: Eraser,
  gradient: Blend,
  'paint-bucket': PaintBucket,
  shape: Shapes,
  hand: Hand,
  zoom: ZoomIn,
})

export function historyIcon(icon: HistoryIcon): LucideIcon {
  return HISTORY_ICONS[icon] ?? Layers
}

// ---------------------------------------------------------------------------------------------
// Snapshots (per document store)
// ---------------------------------------------------------------------------------------------

export interface HistorySnapshot {
  readonly id: number
  readonly name: string
  readonly width: number
  readonly height: number
  readonly ppi: number
  /** Private copies (never inserted into the document; restoring inserts fresh copies). */
  readonly layers: readonly Layer[]
  readonly activeIndex: number
  readonly editTarget: EditTarget
  readonly selection: SelectionSnapshot
  readonly bytes: number
  readonly thumbnail: PixelBuffer | null
}

interface SnapshotHolder {
  list: readonly HistorySnapshot[]
  counter: number
  readonly listeners: Set<() => void>
  unregister: (() => void) | null
}

const holders = new WeakMap<DocumentStore, SnapshotHolder>()
const EMPTY: readonly HistorySnapshot[] = Object.freeze([])

function holderOf(store: DocumentStore): SnapshotHolder {
  let holder = holders.get(store)
  if (!holder) {
    holder = { list: EMPTY, counter: 0, listeners: new Set(), unregister: null }
    holders.set(store, holder)
  }
  return holder
}

function emit(holder: SnapshotHolder): void {
  for (const listener of [...holder.listeners]) listener()
}

/** A private deep copy: same properties and id, its own pixel storage. */
function cloneLayer(layer: Layer): Layer {
  const mask = layer.mask ? { ...layer.mask, surface: layer.mask.surface.clone() } : null
  switch (layer.kind) {
    case 'raster': return { ...layer, mask, surface: layer.surface.clone() }
    case 'text': return { ...layer, mask, raster: { ...layer.raster, surface: layer.raster.surface.clone() } }
    case 'shape': return { ...layer, mask, raster: { ...layer.raster, surface: layer.raster.surface.clone() } }
    default: return { ...layer, mask }
  }
}

const THUMB_EDGE = 56

function thumbnailOf(state: DocumentState): PixelBuffer | null {
  try {
    const maxLevel = maxPyramidLevel(state.width, state.height)
    let level = 0
    while (Math.max(state.width, state.height) / 2 ** level > THUMB_EDGE * 2 && level < maxLevel) level += 1
    const width = Math.max(1, Math.ceil(state.width / 2 ** level))
    const height = Math.max(1, Math.ceil(state.height / 2 ** level))
    return compositeRect(state, { x: 0, y: 0, width, height }, { level })
  } catch (error) {
    console.error(error)
    return null
  }
}

/** Snapshot > New Snapshot. Refuses (with the reason) when the copy would not fit in memory. */
export function createHistorySnapshot(store: DocumentStore, notify: (message: string, tone?: 'normal' | 'error') => void): HistorySnapshot | null {
  const state = store.getState()
  const usage = documentBytes(state)
  const bytes = usage.layers + usage.masks
  const check = ensureMemory(store, bytes, 'make a snapshot')
  if (!check.ok) {
    notify(check.message, 'error')
    return null
  }
  const holder = holderOf(store)
  holder.counter += 1
  const layers = state.layers.map(cloneLayer)
  const snapshot: HistorySnapshot = {
    id: holder.counter,
    name: `Snapshot ${holder.counter}`,
    width: state.width,
    height: state.height,
    ppi: state.ppi,
    layers,
    activeIndex: state.layers.findIndex((layer) => layer.id === state.activeLayerId),
    editTarget: state.editTarget,
    selection: snapshotSelection(state.selection),
    bytes: bytes + snapshotBytes(snapshotSelection(state.selection)),
    thumbnail: thumbnailOf(state),
  }
  holder.list = Object.freeze([...holder.list, snapshot])
  if (!holder.unregister) {
    // Snapshots are the user's: reported to the budget, never trimmed to make room.
    holder.unregister = registerCache({ name: 'history-snapshots', bytes: () => holder.list.reduce((sum, item) => sum + item.bytes, 0) })
  }
  emit(holder)
  return snapshot
}

/** Brings the document back to a snapshot as one undoable step. */
export function revertToSnapshot(store: DocumentStore, snapshot: HistorySnapshot, notify: (message: string, tone?: 'normal' | 'error') => void): boolean {
  const check = ensureMemory(store, snapshot.bytes, `go back to ${snapshot.name}`)
  if (!check.ok) {
    notify(check.message, 'error')
    return false
  }
  try {
    const fresh = snapshot.layers.map((layer) => ({ ...cloneLayer(layer), id: newLayerId() }) as Layer)
    const active = snapshot.activeIndex >= 0 ? fresh[snapshot.activeIndex] ?? null : null
    store.transact(`Revert to ${snapshot.name}`, 'open', (tx) => {
      for (const layer of [...tx.state.layers].reverse()) tx.removeLayer(layer.id)
      tx.setCanvas(snapshot.width, snapshot.height, snapshot.ppi)
      fresh.forEach((layer, index) => tx.insertLayer(layer, index))
      tx.setActiveLayer(active ? active.id : null, active ? snapshot.editTarget : 'pixels')
      tx.setSelection(restoreSelection(snapshot.selection, snapshot.width, snapshot.height, nextSelectionVersion(tx.state)))
    })
    return true
  } catch (error) {
    notify(error instanceof Error ? error.message : `${snapshot.name} could not be restored.`, 'error')
    return false
  }
}

export function deleteHistorySnapshot(store: DocumentStore, id: number): void {
  const holder = holders.get(store)
  if (!holder) return
  holder.list = Object.freeze(holder.list.filter((snapshot) => snapshot.id !== id))
  emit(holder)
}

/** Frees every snapshot of a document (the editor calls this when it closes). */
export function releaseHistorySnapshots(store: DocumentStore): void {
  const holder = holders.get(store)
  if (!holder) return
  holder.list = EMPTY
  holder.unregister?.()
  holder.unregister = null
  emit(holder)
  holder.listeners.clear()
  holders.delete(store)
}

export function useHistorySnapshots(store: DocumentStore): readonly HistorySnapshot[] {
  const subscribe = useCallback((listener: () => void) => {
    const holder = holderOf(store)
    holder.listeners.add(listener)
    return () => { holder.listeners.delete(listener) }
  }, [store])
  return useSyncExternalStore(subscribe, () => holders.get(store)?.list ?? EMPTY)
}

// ---------------------------------------------------------------------------------------------
// The panel
// ---------------------------------------------------------------------------------------------

/**
 * Why history cannot move right now, or null. An open free transform, crop or polygon, or a filter still
 * running in the worker, would land after the jump and replace the later steps, so the user finishes
 * (or cancels) it first, like Photoshop's modal transform.
 */
export function historyBlocker(context: PanelContext): string | null {
  const services = context.commands.services
  const transform = services?.transformSession?.()
  if (transform && transform.hasSession()) return 'Finish the transform first: Enter applies it, Esc cancels it.'
  const tool = services?.activeTool?.()
  if (tool && tool.hasSession()) return tool.id === 'text' ? 'Finish typing first (Ctrl+Enter), or press Esc.' : 'Finish the current tool first: Enter applies it, Esc cancels it.'
  if (commandsBusy(context.store) || services?.busy?.()) return 'Wait a moment: the last change is still being applied.'
  return null
}

function SnapshotThumb({ pixels }: { readonly pixels: PixelBuffer | null }) {
  const ref = useRef<HTMLCanvasElement>(null)
  useEffect(() => {
    const canvas = ref.current
    const context = canvas?.getContext('2d')
    if (!canvas || !context || !pixels) return
    const scratch = new OffscreenCanvas(pixels.width, pixels.height)
    try {
      scratch.getContext('2d')?.putImageData(new ImageData(new Uint8ClampedArray(pixels.data), pixels.width, pixels.height), 0, 0)
      const scale = Math.min(canvas.width / pixels.width, canvas.height / pixels.height)
      const w = Math.max(1, Math.round(pixels.width * scale))
      const h = Math.max(1, Math.round(pixels.height * scale))
      context.clearRect(0, 0, canvas.width, canvas.height)
      context.imageSmoothingQuality = 'high'
      context.drawImage(scratch, Math.floor((canvas.width - w) / 2), Math.floor((canvas.height - h) / 2), w, h)
    } finally {
      scratch.width = 1
      scratch.height = 1
    }
  }, [pixels])
  useEffect(() => {
    const canvas = ref.current
    return () => {
      if (canvas) {
        canvas.width = 1
        canvas.height = 1
      }
    }
  }, [])
  return <canvas ref={ref} className="ae-snap-thumb" width={48} height={34} aria-hidden="true" />
}

function HistoryPanel({ context, suspended }: PanelProps): ReactNode {
  const { store, host } = context
  const history = useHistoryState(store)
  const snapshots = useHistorySnapshots(store)
  const listRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    revealInList(listRef.current, listRef.current?.querySelector('[aria-selected="true"]') ?? null)
  }, [history.cursor, history.entries.length])

  const jump = (index: number, focusCanvas = true) => {
    if (suspended || index === history.cursor) return
    const blocked = historyBlocker(context)
    if (blocked) {
      host.notify(blocked)
      return
    }
    try {
      store.history.jumpTo(index)
    } catch (error) {
      host.notify(error instanceof Error ? error.message : 'History could not move to that step.', 'error')
    }
    if (focusCanvas) context.focusCanvas()
  }

  const onListKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    let target: number | null = null
    if (event.key === 'ArrowUp') target = history.cursor - 1
    else if (event.key === 'ArrowDown') target = history.cursor + 1
    else if (event.key === 'Home') target = 0
    else if (event.key === 'End') target = history.entries.length - 1
    if (target === null) return
    event.preventDefault()
    event.stopPropagation()
    if (target >= 0 && target < history.entries.length && target !== history.cursor) jump(target, false)
  }

  const snapshot = () => {
    if (suspended) return
    const made = createHistorySnapshot(store, host.notify)
    if (made) host.notify(`${made.name} saved. Click it to come back to this state.`)
    context.focusCanvas()
  }

  const bytes = history.totalBytes + snapshots.reduce((sum, item) => sum + item.bytes, 0)
  return (
    <div className="ae-history-panel">
      {snapshots.length > 0 && (
        <div className="ae-snapshots" role="list" aria-label="Snapshots">
          {snapshots.map((item) => (
            <div key={item.id} className="ae-snapshot" role="listitem">
              <button
                type="button"
                className="ae-snapshot-main"
                title={`Go back to ${item.name} (${item.width} × ${item.height} px, ${item.layers.length} layer${item.layers.length === 1 ? '' : 's'}). Undo returns here.`}
                disabled={suspended}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => {
                  const blocked = historyBlocker(context)
                  if (blocked) {
                    host.notify(blocked)
                    return
                  }
                  if (revertToSnapshot(store, item, host.notify)) context.focusCanvas()
                }}
              >
                <SnapshotThumb pixels={item.thumbnail} />
                <span>{item.name}</span>
              </button>
              <button
                type="button"
                className="ae-snapshot-delete"
                title={`Delete ${item.name}`}
                aria-label={`Delete ${item.name}`}
                disabled={suspended}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => deleteHistorySnapshot(store, item.id)}
              ><Trash2 aria-hidden="true" /></button>
            </div>
          ))}
        </div>
      )}
      <div ref={listRef} className="ae-history" role="listbox" aria-label="History" tabIndex={0} onKeyDown={onListKey}>
        {history.trimmed && <p className="ae-muted ae-history-note">Older steps were removed to stay within memory.</p>}
        {history.entries.map((entry, index) => {
          const Icon = index === 0 ? ImageIcon : historyIcon(entry.icon)
          return (
            <div
              key={entry.id}
              role="option"
              aria-selected={index === history.cursor}
              aria-disabled={suspended || undefined}
              className={`ae-history-row ${index === history.cursor ? 'is-current' : ''} ${index > history.cursor ? 'is-future' : ''} ${index === 0 ? 'is-base' : ''}`}
              title={index > history.cursor ? 'Undone: click to redo up to here (a new edit removes it)' : undefined}
              onClick={() => jump(index)}
            >
              <Icon aria-hidden="true" />
              <span>{entry.label}</span>
            </div>
          )
        })}
      </div>
      <div className="ae-panel-footer ae-history-footer">
        <span className="ae-history-stats" title="Undo steps kept and the memory they use">
          {history.entries.length - 1} step{history.entries.length === 2 ? '' : 's'} · {formatBytes(bytes)}
        </span>
        <button type="button" title="New snapshot (keeps a copy of the document you can come back to)" aria-label="New snapshot" disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={snapshot}>
          <Camera aria-hidden="true" />
        </button>
      </div>
    </div>
  )
}

export const panel: PanelDefinition = { id: 'history', title: 'History', component: HistoryPanel, grow: 1 }
