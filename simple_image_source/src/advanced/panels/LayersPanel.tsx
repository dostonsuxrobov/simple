// src/advanced/panels/LayersPanel.tsx (WP6)
// The Layers panel (design 5.14), top layer first:
//   - blend mode and opacity of the active layer (drag the "Opacity" label to scrub), and the four locks
//     (transparency, pixels, position, all); the Background keeps its locks (click its lock badge, or
//     Layer from Background, to make it a normal layer);
//   - each row: eye (Alt+click shows only this layer, Alt+click again brings the others back), thumbnail
//     (Ctrl+click loads its transparency as a selection; Shift / Alt / Shift+Alt with Ctrl add, subtract,
//     intersect), the mask thumbnail when there is one (click targets the mask, Shift+click disables it,
//     Alt+click shows it large, Ctrl+click loads it as a selection), the name (double-click or F2
//     renames), lock badge, clipping arrow (clipped layers are indented and their base is underlined);
//   - drag a row to change the stacking order (nothing goes below the Background);
//   - right-click (or the ... button, or Shift+F10) for Duplicate, Delete, Merge Down / Visible, Flatten,
//     Rasterize, Layer from Background, clipping and mask commands, Select Pixels;
//   - footer: new adjustment layer, add mask, create / release clipping mask, more, duplicate, new, delete.
// Every change is one store transaction with Photoshop's history label; choosing a layer is not a step.
// Thumbnails come from the pyramid level that fits the box and are redrawn at most every 300 ms, only for
// layers whose pixels changed. Keyboard: the list is a listbox (Up / Down / Home / End choose, F2 renames,
// Delete removes the layer, Space toggles visibility).
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, MouseEvent as ReactMouseEvent, PointerEvent as ReactPointerEvent, ReactNode } from 'react'
import { createPortal } from 'react-dom'
import {
  Brush,
  Copy,
  CornerLeftDown,
  Ellipsis,
  Eye,
  EyeOff,
  Grid3x3,
  Lock,
  Move,
  Plus,
  Shapes,
  SlidersHorizontal,
  SquareDot,
  Trash2,
  Type as TypeIcon,
} from 'lucide-react'
import type { AdjustmentType, BlendMode, PixelBuffer, SelectionOp } from '../../imaging/types.ts'
import type { DocumentState, DocumentStore, Layer, LayerId, LayerLocks, LayerMask, Selection } from '../types.ts'
import { BLEND_MODE_MENU, DEFAULT_LOCKS } from '../types.ts'
import type { EditorCommandId } from '../commands.ts'
import { ADJUSTMENT_TYPES, adjustmentLabel } from '../../imaging/adjustments.ts'
import { blendModeLabel } from '../../imaging/blend.ts'
import { levelOffset, maxPyramidLevel, readMaskLevel, readSurfaceLevel } from '../pyramid.ts'
import { applySelectionOp, selectionFromAlpha, selectionFromMask } from '../selection.ts'
import {
  canvasRect,
  isBackgroundLayer,
  layerPixelBounds,
  nextSelectionVersion,
  pixelSourceOf,
  readLayerRect,
  readMaskRect,
  setInteractive,
} from '../tools/shared.ts'
import type { PanelContext, PanelDefinition, PanelProps } from './PanelHost.tsx'
import { useDocumentSelector, useEditorState, usePixelVersion } from './PanelHost.tsx'
import { NumberInput, clampNumber, revealInList } from './AdjustmentControls.tsx'

// ---------------------------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------------------------

const THUMB_CSS = 36
const MASK_CSS = 30
const MASK_VIEW_CSS = 168

function notifyError(context: PanelContext, error: unknown, fallback: string): void {
  context.host.notify(error instanceof Error && error.message ? error.message : fallback, 'error')
}

function transact(context: PanelContext, label: string, run: Parameters<DocumentStore['transact']>[2], options?: { coalesceKey?: string; affectsOutput?: boolean }, icon: 'layer' | 'selection' = 'layer'): boolean {
  try {
    context.store.transact(label, icon, run, options)
    return true
  } catch (error) {
    notifyError(context, error, 'That could not be changed.')
    return false
  }
}

function kindLabel(layer: Layer): string {
  if (layer.kind === 'raster') return layer.isBackground ? 'Background' : 'Pixel layer'
  if (layer.kind === 'text') return 'Type layer'
  if (layer.kind === 'shape') return 'Shape layer'
  return `${adjustmentLabel(layer.adjustment.type)} adjustment layer`
}

function anyLock(locks: LayerLocks): boolean {
  return locks.pixels || locks.position || locks.transparency
}

function allLocked(locks: LayerLocks): boolean {
  return locks.pixels && locks.position && locks.transparency
}

/** Pyramid level whose image of the whole document fits an `edge`-pixel box. */
function thumbLevel(width: number, height: number, edge: number): number {
  const maxLevel = maxPyramidLevel(width, height)
  let level = 0
  while (Math.max(width, height) / 2 ** level > edge && level < maxLevel) level += 1
  return level
}

/** Draws pixels into a canvas, fitted and centred (transparent elsewhere; CSS shows a checkerboard). */
function drawFitted(canvas: HTMLCanvasElement, pixels: PixelBuffer): void {
  const context = canvas.getContext('2d')
  if (!context) return
  context.clearRect(0, 0, canvas.width, canvas.height)
  if (pixels.width <= 0 || pixels.height <= 0) return
  const scratch = new OffscreenCanvas(pixels.width, pixels.height)
  try {
    scratch.getContext('2d')?.putImageData(new ImageData(pixels.data as Uint8ClampedArray<ArrayBuffer>, pixels.width, pixels.height), 0, 0)
    const scale = Math.min(canvas.width / pixels.width, canvas.height / pixels.height)
    const width = Math.max(1, Math.round(pixels.width * scale))
    const height = Math.max(1, Math.round(pixels.height * scale))
    context.imageSmoothingEnabled = true
    context.imageSmoothingQuality = 'high'
    context.drawImage(scratch, Math.floor((canvas.width - width) / 2), Math.floor((canvas.height - height) / 2), width, height)
  } finally {
    scratch.width = 1
    scratch.height = 1
  }
}

function grayPixels(values: { readonly width: number; readonly height: number; readonly data: Uint8Array }): PixelBuffer {
  const data = new Uint8ClampedArray(values.width * values.height * 4)
  for (let i = 0, p = 0; i < values.data.length; i += 1, p += 4) {
    const v = values.data[i]
    data[p] = v
    data[p + 1] = v
    data[p + 2] = v
    data[p + 3] = 255
  }
  return { width: values.width, height: values.height, data }
}

/** Releases a canvas element when its component unmounts (the app's temporary-canvas rule). */
function useReleaseOnUnmount(ref: { readonly current: HTMLCanvasElement | null }): void {
  useEffect(() => {
    const canvas = ref.current
    return () => {
      if (canvas) {
        canvas.width = 1
        canvas.height = 1
      }
    }
  }, [ref])
}

function deviceScale(): number {
  return typeof window !== 'undefined' && window.devicePixelRatio > 0 ? Math.min(3, window.devicePixelRatio) : 1
}

function LayerThumb({ layer, width, height, version }: { readonly layer: Layer; readonly width: number; readonly height: number; readonly version: number }) {
  const ref = useRef<HTMLCanvasElement>(null)
  const drawn = useRef('')
  const size = Math.round(THUMB_CSS * deviceScale())
  useEffect(() => {
    const canvas = ref.current
    const source = pixelSourceOf(layer)
    if (!canvas || !source) return
    const key = `${source.surface.version}|${source.offsetX}|${source.offsetY}|${width}|${height}|${size}`
    if (key === drawn.current) return
    drawn.current = key
    try {
      const level = thumbLevel(width, height, size)
      const pixels = readSurfaceLevel(source.surface, level, {
        x: -levelOffset(source.offsetX, level),
        y: -levelOffset(source.offsetY, level),
        width: Math.max(1, Math.ceil(width / 2 ** level)),
        height: Math.max(1, Math.ceil(height / 2 ** level)),
      })
      drawFitted(canvas, pixels)
    } catch (error) {
      console.error(error)
    }
  }, [height, layer, size, version, width])
  useReleaseOnUnmount(ref)
  return <canvas ref={ref} className="ae-thumb" width={size} height={size} style={{ width: THUMB_CSS, height: THUMB_CSS }} aria-hidden="true" />
}

function MaskThumb({ mask, width, height, version, css = MASK_CSS }: { readonly mask: LayerMask; readonly width: number; readonly height: number; readonly version: number; readonly css?: number }) {
  const ref = useRef<HTMLCanvasElement>(null)
  const drawn = useRef('')
  const size = Math.round(css * deviceScale())
  useEffect(() => {
    const canvas = ref.current
    if (!canvas) return
    const key = `${mask.surface.version}|${mask.offsetX}|${mask.offsetY}|${width}|${height}|${size}`
    if (key === drawn.current) return
    drawn.current = key
    try {
      const level = thumbLevel(width, height, size)
      const values = readMaskLevel(mask.surface, level, {
        x: -levelOffset(mask.offsetX, level),
        y: -levelOffset(mask.offsetY, level),
        width: Math.max(1, Math.ceil(width / 2 ** level)),
        height: Math.max(1, Math.ceil(height / 2 ** level)),
      })
      drawFitted(canvas, grayPixels(values))
    } catch (error) {
      console.error(error)
    }
  }, [height, mask, size, version, width])
  useReleaseOnUnmount(ref)
  return <canvas ref={ref} className="ae-mask-thumb" width={size} height={size} style={{ width: css, height: css }} aria-hidden="true" />
}

// ---------------------------------------------------------------------------------------------
// Selections from layers (Ctrl+click)
// ---------------------------------------------------------------------------------------------

function opFromModifiers(event: { readonly shiftKey: boolean; readonly altKey: boolean }): SelectionOp {
  if (event.shiftKey && event.altKey) return 'intersect'
  if (event.shiftKey) return 'add'
  if (event.altKey) return 'subtract'
  return 'replace'
}

function selectionOfLayer(state: DocumentState, layer: Layer): Selection | null {
  if (layer.kind === 'adjustment') return layer.mask ? selectionOfMask(state, layer.mask) : null
  const bounds = layerPixelBounds(layer)
  if (!bounds) return null
  return selectionFromAlpha(readLayerRect(layer, bounds), bounds.x, bounds.y, state.width, state.height, nextSelectionVersion(state))
}

function selectionOfMask(state: DocumentState, mask: LayerMask): Selection | null {
  return selectionFromMask(readMaskRect(mask, canvasRect(state)), nextSelectionVersion(state))
}

function loadSelection(context: PanelContext, shape: Selection | null, op: SelectionOp, emptyMessage: string): void {
  const state = context.store.getState()
  if (!shape && op === 'replace') {
    context.host.notify(emptyMessage)
    return
  }
  const version = nextSelectionVersion(state)
  const next = shape
    ? (op === 'replace' ? { ...shape, version } : applySelectionOp(state.selection, shape.mask, op, version))
    : op === 'intersect' ? null : state.selection
  if (next === state.selection) return
  const label = op === 'replace' ? 'Load Selection' : op === 'add' ? 'Add to Selection' : op === 'subtract' ? 'Subtract from Selection' : 'Intersect Selection'
  transact(context, label, (tx) => tx.setSelection(next), { affectsOutput: false }, 'selection')
}

// ---------------------------------------------------------------------------------------------
// Popup menu (context menu, adjustment-layer menu)
// ---------------------------------------------------------------------------------------------

type MenuEntry =
  | { readonly kind: 'command'; readonly command: EditorCommandId; readonly label?: string }
  | { readonly kind: 'action'; readonly label: string; readonly run: () => void; readonly disabled?: boolean }
  | { readonly kind: 'separator' }

const cmd = (command: EditorCommandId, label?: string): MenuEntry => ({ kind: 'command', command, label })
const SEPARATOR: MenuEntry = { kind: 'separator' }

function PopupMenu({ x, y, entries, context, label, onClose }: {
  readonly x: number
  readonly y: number
  readonly entries: readonly MenuEntry[]
  readonly context: PanelContext
  readonly label: string
  readonly onClose: (runCommand: EditorCommandId | null) => void
}) {
  const ref = useRef<HTMLDivElement>(null)
  const [position, setPosition] = useState({ left: x, top: y })
  const closeRef = useRef(onClose)
  closeRef.current = onClose

  useLayoutEffect(() => {
    const element = ref.current
    if (!element) return
    const width = element.offsetWidth
    const height = element.offsetHeight
    setPosition({
      left: Math.max(6, Math.min(x, window.innerWidth - width - 6)),
      top: Math.max(6, y + height > window.innerHeight - 6 ? y - height : y),
    })
    element.querySelector<HTMLElement>('[role="menuitem"]:not([aria-disabled="true"])')?.focus({ preventScroll: true })
  }, [x, y])

  useEffect(() => {
    const onPointerDown = (event: PointerEvent) => {
      if (event.target instanceof Node && ref.current?.contains(event.target)) return
      closeRef.current(null)
    }
    const onBlur = () => closeRef.current(null)
    document.addEventListener('pointerdown', onPointerDown, true)
    window.addEventListener('blur', onBlur)
    window.addEventListener('resize', onBlur)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true)
      window.removeEventListener('blur', onBlur)
      window.removeEventListener('resize', onBlur)
    }
  }, [])

  const items = entries.map((entry) => {
    if (entry.kind === 'separator') return { entry, enabled: false, text: '', keys: '' }
    if (entry.kind === 'action') return { entry, enabled: !entry.disabled, text: entry.label, keys: '' }
    return { entry, enabled: context.isEnabled(entry.command), text: entry.label ?? context.label(entry.command), keys: context.shortcut(entry.command) }
  })

  const activate = (index: number) => {
    const item = items[index]
    if (!item || !item.enabled) return
    if (item.entry.kind === 'command') closeRef.current(item.entry.command)
    else if (item.entry.kind === 'action') {
      closeRef.current(null)
      item.entry.run()
    }
  }

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const focusables = Array.from(ref.current?.querySelectorAll<HTMLElement>('[role="menuitem"]:not([aria-disabled="true"])') ?? [])
    const at = focusables.indexOf(document.activeElement as HTMLElement)
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault()
      const next = focusables[(at + (event.key === 'ArrowDown' ? 1 : -1) + focusables.length) % Math.max(1, focusables.length)]
      next?.focus({ preventScroll: true })
    } else if (event.key === 'Home' || event.key === 'End') {
      event.preventDefault()
      focusables[event.key === 'Home' ? 0 : focusables.length - 1]?.focus({ preventScroll: true })
    } else if (event.key === 'Escape' || event.key === 'Tab') {
      event.preventDefault()
      closeRef.current(null)
    } else if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault()
      const index = Number((document.activeElement as HTMLElement | null)?.dataset.index)
      if (Number.isFinite(index)) activate(index)
    }
    // Menu keys never become editor shortcuts.
    event.stopPropagation()
  }

  const host = (typeof document !== 'undefined' && document.querySelector('.ae-shell')) || document.body
  return createPortal(
    <div ref={ref} className="ae-menu ae-popup-menu" role="menu" aria-label={label} style={{ left: position.left, top: position.top }} onKeyDown={onKeyDown} onContextMenu={(event) => event.preventDefault()}>
      {items.map((item, index) => item.entry.kind === 'separator'
        ? <div key={`sep-${index}`} className="ae-menu-separator" role="separator" />
        : (
          <button
            key={`${index}-${item.text}`}
            type="button"
            role="menuitem"
            data-index={index}
            className="ae-menu-item"
            aria-disabled={!item.enabled || undefined}
            tabIndex={-1}
            onMouseEnter={(event) => { if (item.enabled) event.currentTarget.focus({ preventScroll: true }) }}
            onClick={() => activate(index)}
          >
            <span className="ae-menu-check" />
            <span className="ae-menu-label">{item.text}</span>
            <span className="ae-menu-keys">{item.keys}</span>
          </button>
        ))}
    </div>,
    host,
  )
}

// ---------------------------------------------------------------------------------------------
// Opacity control
// ---------------------------------------------------------------------------------------------

function OpacityControl({ context, layer, disabled }: { readonly context: PanelContext; readonly layer: Layer | null; readonly disabled: boolean }) {
  const [sliderOpen, setSliderOpen] = useState(false)
  const percent = layer ? Math.round(layer.opacity * 100) : 100
  const set = (value: number) => {
    if (!layer) return
    const opacity = clampNumber(Math.round(value), 0, 100) / 100
    if (opacity === layer.opacity) return
    transact(context, 'Opacity Change', (tx) => tx.updateLayer(layer.id, { opacity }), { coalesceKey: `layer-opacity:${layer.id}` })
  }
  const scrub = (event: ReactPointerEvent<HTMLSpanElement>) => {
    if (disabled || !layer || event.button !== 0) return
    event.preventDefault()
    const element = event.currentTarget
    const startX = event.clientX
    const start = percent
    try {
      element.setPointerCapture(event.pointerId)
    } catch {
      // ignore
    }
    setInteractive(context.commands, true)
    const move = (moveEvent: PointerEvent) => set(start + Math.round((moveEvent.clientX - startX) / (moveEvent.shiftKey ? 0.25 : 1)))
    const end = () => {
      element.removeEventListener('pointermove', move)
      element.removeEventListener('pointerup', end)
      element.removeEventListener('pointercancel', end)
      setInteractive(context.commands, false)
    }
    element.addEventListener('pointermove', move)
    element.addEventListener('pointerup', end)
    element.addEventListener('pointercancel', end)
  }
  return (
    <div className="ae-opacity-control">
      <span className="ae-scrub" title="Opacity: drag left or right to change it" onPointerDown={scrub}>Opacity</span>
      <NumberInput label="Layer opacity in percent" value={percent} min={0} max={100} disabled={disabled || !layer} onChange={set} />
      <button type="button" className="ae-opacity-more" aria-label="Opacity slider" aria-expanded={sliderOpen} title="Show the opacity slider" disabled={disabled || !layer} onClick={() => setSliderOpen((open) => !open)}>%</button>
      {sliderOpen && layer && !disabled && (
        <div className="ae-opacity-pop" onKeyDown={(event) => { if (event.key === 'Escape' || event.key === 'Enter') setSliderOpen(false); event.stopPropagation() }}>
          <input
            type="range"
            className="ae-slider"
            min={0}
            max={100}
            value={percent}
            aria-label="Layer opacity"
            autoFocus
            onPointerDown={() => setInteractive(context.commands, true)}
            onPointerUp={() => setInteractive(context.commands, false)}
            onPointerCancel={() => setInteractive(context.commands, false)}
            onBlur={() => setSliderOpen(false)}
            onChange={(event) => set(Number(event.currentTarget.value))}
          />
        </div>
      )}
    </div>
  )
}

// ---------------------------------------------------------------------------------------------
// Solo memory (Alt+click on an eye brings back exactly what was visible before)
// ---------------------------------------------------------------------------------------------

const soloMemory = new WeakMap<DocumentStore, { readonly layerId: LayerId; readonly visible: ReadonlyMap<LayerId, boolean> }>()

// ---------------------------------------------------------------------------------------------
// The panel
// ---------------------------------------------------------------------------------------------

interface DragState {
  readonly id: LayerId
  readonly pointerId: number
  readonly startY: number
  active: boolean
}

interface DropIndicator {
  readonly id: LayerId
  /** Gap in display order (0 = above the top row). */
  readonly gap: number
  readonly top: number
}

interface MenuState {
  readonly x: number
  readonly y: number
  readonly kind: 'layer' | 'adjustment'
}

const BLEND_OPTIONS = BLEND_MODE_MENU

function LayersPanel({ context, suspended }: PanelProps): ReactNode {
  const { store } = context
  const layers = useDocumentSelector(store, (state) => state.layers)
  const activeId = useDocumentSelector(store, (state) => state.activeLayerId)
  const editTarget = useDocumentSelector(store, (state) => state.editTarget)
  const docWidth = useDocumentSelector(store, (state) => state.width)
  const docHeight = useDocumentSelector(store, (state) => state.height)
  const editor = useEditorState(context.editor)
  const version = usePixelVersion(store, 300)
  const [renaming, setRenaming] = useState<LayerId | null>(null)
  const [menu, setMenu] = useState<MenuState | null>(null)
  const [maskView, setMaskView] = useState<{ readonly id: LayerId; readonly top: number } | null>(null)
  const [drop, setDrop] = useState<DropIndicator | null>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const dragRef = useRef<DragState | null>(null)
  const suppressClick = useRef(false)
  // The blend menu hands the keyboard back to the canvas after a mouse choice, not while arrowing through it.
  const blendByKeyboard = useRef(false)

  const active = layers.find((layer) => layer.id === activeId) ?? null
  const background = Boolean(active && isBackgroundLayer(active))
  const display = [...layers].reverse()

  useEffect(() => {
    if (!activeId) return
    revealInList(listRef.current, listRef.current?.querySelector(`[data-layer-id="${CSS.escape(activeId)}"]`) ?? null)
  }, [activeId])

  useEffect(() => {
    if (maskView && !layers.some((layer) => layer.id === maskView.id && layer.mask)) setMaskView(null)
  }, [layers, maskView])

  const select = useCallback((layer: Layer, target: 'pixels' | 'mask' | 'keep' = 'keep') => {
    const state = store.getState()
    const nextTarget = target === 'keep' ? (state.activeLayerId === layer.id ? state.editTarget : 'pixels') : target
    if (state.activeLayerId === layer.id && state.editTarget === nextTarget) return
    transact(context, 'Select Layer', (tx) => tx.setActiveLayer(layer.id, nextTarget), { affectsOutput: false })
  }, [context, store])

  const toggleVisible = (layer: Layer, solo: boolean) => {
    const state = store.getState()
    if (solo) {
      const remembered = soloMemory.get(store)
      const others = state.layers.filter((entry) => entry.id !== layer.id)
      const isSolo = layer.visible && others.every((entry) => !entry.visible)
      if (isSolo) {
        // Bring back what was visible before (everything, when that is unknown).
        transact(context, 'Show Layers', (tx) => {
          for (const entry of state.layers) {
            const visible = remembered && remembered.layerId === layer.id ? remembered.visible.get(entry.id) ?? true : true
            if (entry.visible !== visible) tx.updateLayer(entry.id, { visible })
          }
        })
        soloMemory.delete(store)
      } else {
        soloMemory.set(store, { layerId: layer.id, visible: new Map(state.layers.map((entry) => [entry.id, entry.visible])) })
        transact(context, 'Hide Layers', (tx) => {
          for (const entry of state.layers) {
            const visible = entry.id === layer.id
            if (entry.visible !== visible) tx.updateLayer(entry.id, { visible })
          }
        })
      }
      return
    }
    transact(context, layer.visible ? 'Hide Layer' : 'Show Layer', (tx) => tx.updateLayer(layer.id, { visible: !layer.visible }))
  }

  const rename = (layer: Layer, name: string) => {
    setRenaming(null)
    const trimmed = name.trim()
    // Enter and the blur that follows both arrive here: compare with the layer as it is now.
    const current = store.getState().layers.find((entry) => entry.id === layer.id)
    if (trimmed && current && trimmed !== current.name) transact(context, 'Rename Layer', (tx) => tx.updateLayer(layer.id, { name: trimmed }))
    context.focusCanvas()
  }

  const setLocks = (layer: Layer, locks: LayerLocks) => {
    const label = anyLock(locks) && !anyLock(layer.locks) ? 'Lock Layer' : !anyLock(locks) ? 'Unlock Layer' : 'Change Layer Locks'
    transact(context, label, (tx) => tx.updateLayer(layer.id, { locks }))
  }

  const toggleMaskEnabled = (layer: Layer) => {
    if (!layer.mask) return
    const enabled = !layer.mask.enabled
    transact(context, enabled ? 'Enable Layer Mask' : 'Disable Layer Mask', (tx) => tx.updateLayer(layer.id, { maskEnabled: enabled }))
  }

  const runOn = (layer: Layer | null, command: EditorCommandId) => {
    if (layer) select(layer)
    context.run(command)
  }

  // ----- drag to reorder ---------------------------------------------------------------------

  const gapAt = (clientY: number): { gap: number; top: number } => {
    const list = listRef.current
    if (!list) return { gap: 0, top: 0 }
    const rows = Array.from(list.querySelectorAll<HTMLElement>('[data-layer-row]'))
    const listBox = list.getBoundingClientRect()
    for (let index = 0; index < rows.length; index += 1) {
      const box = rows[index].getBoundingClientRect()
      if (clientY < box.top + box.height / 2) return { gap: index, top: box.top - listBox.top + list.scrollTop }
    }
    const last = rows[rows.length - 1]
    return { gap: rows.length, top: last ? last.getBoundingClientRect().bottom - listBox.top + list.scrollTop : 0 }
  }

  const finishDrag = (id: LayerId, gap: number) => {
    const state = store.getState()
    const count = state.layers.length
    const from = state.layers.findIndex((layer) => layer.id === id)
    if (from < 0) return
    const fromDisplay = count - 1 - from
    const finalDisplay = fromDisplay < gap ? gap - 1 : gap
    let to = count - 1 - finalDisplay
    if (isBackgroundLayer(state.layers[0])) to = Math.max(1, to)
    to = clampNumber(to, 0, count - 1)
    if (to === from) return
    transact(context, 'Layer Order', (tx) => tx.moveLayer(id, to))
  }

  const onRowPointerDown = (event: ReactPointerEvent<HTMLDivElement>, layer: Layer) => {
    if (suspended || event.button !== 0 || renaming) return
    if (event.target instanceof Element && event.target.closest('button, input')) return
    if (isBackgroundLayer(layer)) return
    dragRef.current = { id: layer.id, pointerId: event.pointerId, startY: event.clientY, active: false }
    const move = (moveEvent: PointerEvent) => {
      const drag = dragRef.current
      if (!drag || moveEvent.pointerId !== drag.pointerId) return
      if (!drag.active) {
        if (Math.abs(moveEvent.clientY - drag.startY) < 5) return
        drag.active = true
      }
      const list = listRef.current
      if (list) {
        const box = list.getBoundingClientRect()
        if (moveEvent.clientY < box.top + 18) list.scrollTop -= 10
        else if (moveEvent.clientY > box.bottom - 18) list.scrollTop += 10
      }
      const { gap, top } = gapAt(moveEvent.clientY)
      setDrop({ id: drag.id, gap, top })
    }
    const end = (upEvent: PointerEvent) => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', end)
      window.removeEventListener('pointercancel', cancel)
      const drag = dragRef.current
      dragRef.current = null
      setDrop(null)
      if (drag?.active) {
        suppressClick.current = true
        setTimeout(() => { suppressClick.current = false }, 0)
        finishDrag(drag.id, gapAt(upEvent.clientY).gap)
        context.focusCanvas()
      }
    }
    const cancel = () => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', end)
      window.removeEventListener('pointercancel', cancel)
      dragRef.current = null
      setDrop(null)
    }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', end)
    window.addEventListener('pointercancel', cancel)
  }

  // ----- keyboard ----------------------------------------------------------------------------

  const onListKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (renaming || suspended) return
    if (!(event.target instanceof HTMLElement) || event.target !== event.currentTarget) return
    const index = display.findIndex((layer) => layer.id === activeId)
    let handled = true
    if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
      const next = display[index < 0 ? 0 : index + (event.key === 'ArrowUp' ? -1 : 1)]
      if (next) select(next, 'pixels')
    } else if (event.key === 'Home' || event.key === 'End') {
      const next = display[event.key === 'Home' ? 0 : display.length - 1]
      if (next) select(next, 'pixels')
    } else if (event.key === 'F2' && active) {
      setRenaming(active.id)
    } else if ((event.key === 'Delete' || event.key === 'Backspace') && active) {
      context.run('layer.delete')
    } else if (event.key === ' ' && active) {
      toggleVisible(active, event.altKey)
    } else if ((event.key === 'F10' && event.shiftKey) || event.key === 'ContextMenu') {
      const row = active ? listRef.current?.querySelector(`[data-layer-id="${CSS.escape(active.id)}"]`) : null
      const box = (row ?? event.currentTarget).getBoundingClientRect()
      setMenu({ x: box.left + 24, y: box.bottom - 4, kind: 'layer' })
    } else {
      handled = false
    }
    if (handled) {
      event.preventDefault()
      event.stopPropagation()
    }
  }

  // ----- menus -------------------------------------------------------------------------------

  const layerMenu = (): MenuEntry[] => {
    const entries: MenuEntry[] = [cmd('layer.duplicate'), cmd('layer.delete'), SEPARATOR]
    if (active && isBackgroundLayer(active)) entries.push(cmd('layer.from-background'))
    if (active && (active.kind === 'text' || active.kind === 'shape')) entries.push(cmd('layer.rasterize'))
    entries.push(cmd('layer.toggle-clipping'), SEPARATOR)
    if (active?.mask) entries.push(cmd('layer.toggle-mask'), cmd('layer.apply-mask'), cmd('layer.delete-mask'))
    else entries.push(cmd('layer.add-mask'))
    entries.push(cmd('select.load-layer-alpha', 'Select Pixels'), SEPARATOR)
    entries.push(cmd('layer.to-front'), cmd('layer.to-back'), SEPARATOR)
    entries.push(cmd('layer.merge-down'), cmd('layer.merge-visible'), cmd('layer.stamp-visible'), cmd('layer.flatten'))
    return entries
  }

  const adjustmentMenu = (): MenuEntry[] => ADJUSTMENT_TYPES.flatMap((type: AdjustmentType, index): MenuEntry[] => {
    const entry = cmd(`adjustment-layer.${type}` as EditorCommandId, adjustmentLabel(type))
    return index === 4 || index === 9 ? [SEPARATOR, entry] : [entry]
  })

  const closeMenu = (command: EditorCommandId | null) => {
    setMenu(null)
    if (command) context.run(command)
    else context.focusCanvas()
  }

  const openMenuAt = (event: ReactMouseEvent<HTMLElement>, kind: MenuState['kind']) => {
    const box = event.currentTarget.getBoundingClientRect()
    setMenu((current) => (current ? null : { x: box.left, y: kind === 'adjustment' ? box.top - 4 : box.top - 4, kind }))
  }

  // ----- rows --------------------------------------------------------------------------------

  const baseOfClip = new Set<LayerId>()
  layers.forEach((layer, index) => {
    if (!layer.clipped && layers[index + 1]?.clipped && !isBackgroundLayer(layers[index + 1])) baseOfClip.add(layer.id)
  })

  const locks = active?.locks ?? DEFAULT_LOCKS
  const lockDisabled = suspended || !active
  const backgroundLocks = Boolean(active && isBackgroundLayer(active))
  const adjustmentActive = active?.kind === 'adjustment'
  const viewed = maskView ? layers.find((layer) => layer.id === maskView.id) ?? null : null

  return (
    <div className="ae-layers">
      <div className="ae-layer-props">
        <select
          aria-label="Blend mode"
          title={background ? 'The Background is always Normal. Use Layer from Background to change it.' : 'Blend mode'}
          value={active?.blendMode ?? 'normal'}
          disabled={!active || background || suspended}
          onKeyDown={() => { blendByKeyboard.current = true }}
          onPointerDown={() => { blendByKeyboard.current = false }}
          onChange={(event) => {
            const blendMode = event.currentTarget.value as BlendMode
            if (active && active.blendMode !== blendMode) transact(context, 'Blending Change', (tx) => tx.updateLayer(active.id, { blendMode }))
            if (!blendByKeyboard.current) context.focusCanvas()
          }}
        >
          {BLEND_OPTIONS.map((mode, index) => mode === '-'
            ? <option key={`sep-${index}`} disabled>──────────</option>
            : <option key={mode} value={mode}>{blendModeLabel(mode)}</option>)}
        </select>
        <OpacityControl context={context} layer={active} disabled={suspended || background} />
      </div>
      <div className="ae-lock-row" role="group" aria-label="Lock">
        <span>Lock</span>
        <button
          type="button"
          className={`ae-lock-toggle ${locks.transparency ? 'is-on' : ''}`}
          title="Lock transparent pixels (painting keeps transparency)"
          aria-label="Lock transparent pixels"
          aria-pressed={locks.transparency}
          disabled={lockDisabled || backgroundLocks || adjustmentActive}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => active && setLocks(active, { ...active.locks, transparency: !active.locks.transparency })}
        ><Grid3x3 aria-hidden="true" /></button>
        <button
          type="button"
          className={`ae-lock-toggle ${locks.pixels ? 'is-on' : ''}`}
          title="Lock image pixels (no painting, filters or adjustments)"
          aria-label="Lock image pixels"
          aria-pressed={locks.pixels}
          disabled={lockDisabled || adjustmentActive}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => active && setLocks(active, { ...active.locks, pixels: !active.locks.pixels })}
        ><Brush aria-hidden="true" /></button>
        <button
          type="button"
          className={`ae-lock-toggle ${locks.position ? 'is-on' : ''}`}
          title="Lock position (no moving or transforming)"
          aria-label="Lock position"
          aria-pressed={locks.position}
          disabled={lockDisabled || backgroundLocks}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => active && setLocks(active, { ...active.locks, position: !active.locks.position })}
        ><Move aria-hidden="true" /></button>
        <button
          type="button"
          className={`ae-lock-toggle ${active && allLocked(active.locks) ? 'is-on' : ''}`}
          title="Lock all"
          aria-label="Lock all"
          aria-pressed={Boolean(active && allLocked(active.locks))}
          disabled={lockDisabled || backgroundLocks}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            if (!active) return
            const on = !allLocked(active.locks)
            setLocks(active, { pixels: on, position: on, transparency: on })
          }}
        ><Lock aria-hidden="true" /></button>
      </div>
      <div
        ref={listRef}
        className="ae-layer-list"
        role="listbox"
        aria-label="Layers"
        aria-activedescendant={activeId ? `ae-layer-${activeId}` : undefined}
        tabIndex={0}
        onKeyDown={onListKey}
      >
        {display.map((layer) => {
          const selected = layer.id === activeId
          const isBackground = isBackgroundLayer(layer)
          const locked = anyLock(layer.locks)
          const dragging = drop?.id === layer.id
          return (
            <div
              key={layer.id}
              id={`ae-layer-${layer.id}`}
              data-layer-row=""
              data-layer-id={layer.id}
              role="option"
              aria-selected={selected}
              className={`ae-layer-row ${selected ? 'is-active' : ''} ${layer.visible ? '' : 'is-hidden'} ${layer.clipped && !isBackground ? 'is-clipped' : ''} ${dragging ? 'is-dragging' : ''}`}
              title={kindLabel(layer)}
              onPointerDown={(event) => onRowPointerDown(event, layer)}
              onClick={() => {
                if (suppressClick.current) return
                select(layer)
              }}
              onContextMenu={(event) => {
                event.preventDefault()
                if (suspended) return
                select(layer)
                setMenu({ x: event.clientX, y: event.clientY, kind: 'layer' })
              }}
            >
              <button
                type="button"
                className="ae-eye"
                aria-label={layer.visible ? `Hide ${layer.name}` : `Show ${layer.name}`}
                aria-pressed={layer.visible}
                title={layer.visible ? 'Hide (Alt+click: show only this layer)' : 'Show (Alt+click: show only this layer)'}
                disabled={suspended}
                onMouseDown={(event) => event.preventDefault()}
                onClick={(event) => {
                  event.stopPropagation()
                  toggleVisible(layer, event.altKey)
                }}
              >{layer.visible ? <Eye aria-hidden="true" /> : <EyeOff aria-hidden="true" />}</button>
              <span className="ae-layer-clip" aria-hidden="true">{layer.clipped && !isBackground ? <CornerLeftDown /> : null}</span>
              <span
                className={`ae-thumb-wrap ${selected && editTarget === 'pixels' ? 'is-target' : ''}`}
                title={layer.kind === 'adjustment' ? 'Adjustment layer: its settings are in Properties' : 'Ctrl+click: select the pixels of this layer'}
                onClick={(event) => {
                  if (event.ctrlKey || event.metaKey) {
                    event.stopPropagation()
                    if (suppressClick.current) return
                    loadSelection(context, selectionOfLayer(store.getState(), layer), opFromModifiers(event), `"${layer.name}" has no pixels to select.`)
                    context.focusCanvas()
                    return
                  }
                  event.stopPropagation()
                  if (suppressClick.current) return
                  select(layer, 'pixels')
                }}
                onDoubleClick={() => {
                  if (layer.kind === 'adjustment' && !editor.panels.properties) context.run('view.panel-properties')
                }}
              >
                {layer.kind === 'adjustment'
                  ? <span className="ae-thumb is-icon" style={{ width: THUMB_CSS, height: THUMB_CSS }}><SlidersHorizontal aria-hidden="true" /></span>
                  : <LayerThumb layer={layer} width={docWidth} height={docHeight} version={version} />}
              </span>
              {layer.mask ? (
                <span
                  className={`ae-thumb-wrap is-mask ${selected && editTarget === 'mask' ? 'is-target' : ''} ${layer.mask.enabled ? '' : 'is-disabled'}`}
                  role="button"
                  tabIndex={-1}
                  aria-label={`Layer mask of ${layer.name}${layer.mask.enabled ? '' : ' (disabled)'}`}
                  title={`Layer mask${layer.mask.enabled ? '' : ' (disabled)'}: click to edit it, Shift+click to ${layer.mask.enabled ? 'disable' : 'enable'}, Alt+click to view, Ctrl+click to select`}
                  onClick={(event) => {
                    event.stopPropagation()
                    if (suppressClick.current || suspended) return
                    if (event.ctrlKey || event.metaKey) {
                      loadSelection(context, layer.mask ? selectionOfMask(store.getState(), layer.mask) : null, opFromModifiers(event), 'The mask hides everything; there is nothing to select.')
                      context.focusCanvas()
                    } else if (event.shiftKey) {
                      toggleMaskEnabled(layer)
                    } else if (event.altKey) {
                      const row = event.currentTarget.closest('[data-layer-row]')
                      const box = row?.getBoundingClientRect()
                      setMaskView((current) => (current?.id === layer.id ? null : { id: layer.id, top: box ? box.bottom : 0 }))
                    } else {
                      select(layer, 'mask')
                    }
                  }}
                >
                  <MaskThumb mask={layer.mask} width={docWidth} height={docHeight} version={version} />
                </span>
              ) : null}
              {renaming === layer.id ? (
                <input
                  className="ae-rename"
                  aria-label="Layer name"
                  defaultValue={layer.name}
                  autoFocus
                  onFocus={(event) => event.currentTarget.select()}
                  onClick={(event) => event.stopPropagation()}
                  onPointerDown={(event) => event.stopPropagation()}
                  onBlur={(event) => rename(layer, event.currentTarget.value)}
                  onKeyDown={(event) => {
                    event.stopPropagation()
                    if (event.key === 'Enter') rename(layer, event.currentTarget.value)
                    if (event.key === 'Escape') {
                      setRenaming(null)
                      context.focusCanvas()
                    }
                  }}
                />
              ) : (
                <span
                  className={`ae-layer-name ${baseOfClip.has(layer.id) ? 'is-clip-base' : ''}`}
                  onDoubleClick={(event) => {
                    event.stopPropagation()
                    if (!suspended) setRenaming(layer.id)
                  }}
                >
                  {layer.kind === 'text' && <TypeIcon className="ae-kind" aria-hidden="true" />}
                  {layer.kind === 'shape' && <Shapes className="ae-kind" aria-hidden="true" />}
                  <span className={isBackground ? 'is-italic' : undefined}>{layer.name}</span>
                </span>
              )}
              {locked ? (
                <button
                  type="button"
                  className="ae-lock-badge"
                  title={isBackground ? 'The Background is locked. Click to turn it into a normal layer.' : 'Locked. Click to unlock.'}
                  aria-label={isBackground ? 'Turn the Background into a normal layer' : `Unlock ${layer.name}`}
                  disabled={suspended}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={(event) => {
                    event.stopPropagation()
                    if (isBackground) runOn(layer, 'layer.from-background')
                    else setLocks(layer, DEFAULT_LOCKS)
                  }}
                ><Lock aria-hidden="true" /></button>
              ) : <span />}
            </div>
          )
        })}
        {drop && <div className="ae-drop-line" style={{ top: drop.top }} aria-hidden="true" />}
      </div>
      <div className="ae-panel-footer ae-layers-footer">
        <button type="button" title="New fill or adjustment layer" aria-label="New adjustment layer" aria-haspopup="menu" aria-expanded={menu?.kind === 'adjustment'} disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={(event) => openMenuAt(event, 'adjustment')}>
          <SlidersHorizontal aria-hidden="true" />
        </button>
        <button type="button" title={active?.mask ? `${context.label('layer.toggle-mask')}` : 'Add layer mask (from the selection when there is one)'} aria-label={active?.mask ? context.label('layer.toggle-mask') : 'Add layer mask'} disabled={suspended || !(active?.mask ? context.isEnabled('layer.toggle-mask') : context.isEnabled('layer.add-mask'))} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run(active?.mask ? 'layer.toggle-mask' : 'layer.add-mask')}>
          <SquareDot aria-hidden="true" />
        </button>
        <button type="button" title={`${context.label('layer.toggle-clipping')} (${context.shortcut('layer.toggle-clipping')})`} aria-label={context.label('layer.toggle-clipping')} aria-pressed={Boolean(active?.clipped)} disabled={suspended || !context.isEnabled('layer.toggle-clipping')} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('layer.toggle-clipping')}>
          <CornerLeftDown aria-hidden="true" />
        </button>
        <span className="ae-footer-spacer" />
        <button type="button" title="More layer commands" aria-label="More layer commands" aria-haspopup="menu" aria-expanded={menu?.kind === 'layer'} disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={(event) => openMenuAt(event, 'layer')}>
          <Ellipsis aria-hidden="true" />
        </button>
        <button type="button" title={`Duplicate layer (${context.shortcut('layer.via-copy')} without a selection)`} aria-label="Duplicate layer" disabled={suspended || !context.isEnabled('layer.duplicate')} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('layer.duplicate')}>
          <Copy aria-hidden="true" />
        </button>
        <button type="button" title={`New layer (${context.shortcut('layer.new')})`} aria-label="New layer" disabled={suspended || !context.isEnabled('layer.new')} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('layer.new')}>
          <Plus aria-hidden="true" />
        </button>
        <button type="button" title="Delete layer" aria-label="Delete layer" disabled={suspended || !context.isEnabled('layer.delete')} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('layer.delete')}>
          <Trash2 aria-hidden="true" />
        </button>
      </div>
      {menu && !suspended && (
        <PopupMenu
          x={menu.x}
          y={menu.y}
          context={context}
          label={menu.kind === 'adjustment' ? 'New adjustment layer' : 'Layer'}
          entries={menu.kind === 'adjustment' ? adjustmentMenu() : layerMenu()}
          onClose={closeMenu}
        />
      )}
      {viewed?.mask && maskView && (
        <div className="ae-mask-view" role="dialog" aria-label={`Layer mask of ${viewed.name}`} onClick={() => setMaskView(null)}>
          <MaskThumb mask={viewed.mask} width={docWidth} height={docHeight} version={version} css={MASK_VIEW_CSS} />
          <span>{viewed.name}: mask{viewed.mask.enabled ? '' : ' (disabled)'}. White shows the layer, black hides it. Click to close.</span>
        </div>
      )}
    </div>
  )
}

export const panel: PanelDefinition = { id: 'layers', title: 'Layers', component: LayersPanel, grow: 3 }
