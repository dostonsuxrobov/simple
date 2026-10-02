// src/advanced/commands.ts (WP6)
// The command registry of the Advanced editor (design 5.14, 5.15, 8.4): every menu item, shortcut and panel
// button runs through runCommand(id, context), and menus grey items out with isCommandEnabled(id, context).
//   - EditorCommandId = the shared CommandId set plus editor-only commands (opacity / flow digits, panel
//     toggles, Sharpen More) that the frozen contract has no ids for.
//   - Commands change the document only through store.transact (one history step each, Photoshop labels);
//     selection-only and active-layer changes never mark the document modified.
//   - Refusals explain themselves through host.notify (Photoshop's "could not ... because" messages).
//   - Document operations the dialogs need are exported (applyAdjustmentToLayer, applyFilterToLayer,
//     resizeImage, resizeCanvas, rotateCanvas, modifySelection, placePixelsAsLayer), so a dialog only
//     collects parameters. Commands that need parameters open a dialog through CommandServices.openDialog.
//   - Worker jobs (filters, adjustments, image size, rotation, selection modify) are tracked per document
//     (commandsBusy / whenCommandsIdle) and write back only when the layer is exactly as it was when the job
//     started, so a long filter can never overwrite later edits.
// DOM-free at module load (Node tests drive it with the real document store and an inline imaging client);
// DOM-only services (clipboard encode/decode, dialogs, free transform) come in through CommandServices.
import type {
  AdjustmentSpec,
  AdjustmentType,
  Affine,
  FilterSpec,
  FilterType,
  IntRect,
  MaskBuffer,
  PixelBuffer,
  Point,
  ResampleMethod,
  Rgb8,
} from '../imaging/types.ts'
import type {
  CommandId,
  DocumentState,
  HistoryIcon,
  Layer,
  LayerId,
  LayerMask,
  RasterLayer,
  Selection,
  SelectionSnapshot,
  ShapeLayer,
  TextLayer,
  TiledMask,
  TiledSurface,
  ToolContext,
  ToolId,
  Transaction,
} from './types.ts'
import { DEFAULT_LOCKS, LIMITS } from './types.ts'
import {
  createAdjustmentLayer,
  createRasterLayer,
  duplicateLayer,
  layerContentBounds,
  maskFromSelection,
  nextLayerName,
  pixelEditBlocker,
  rasterizeVectorLayer,
} from './document.ts'
import { createMaskSurface, createSurface, maskFromBuffer, surfaceFromBuffer } from './tiles.ts'
import {
  invertSelection,
  restoreSelection,
  selectAll,
  selectionCoverage,
  selectionFromAlpha,
  selectionFromMask,
  snapshotSelection,
} from './selection.ts'
import { compositeRect } from './composite.ts'
import { ensureMemory } from './memory.ts'
import { stepZoom } from './viewport.ts'
import { composeAffine } from './toolGeometry.ts'
import type { AdvancedToolController, JobTracker, PaintTarget } from './tools/shared.ts'
import {
  activeLayerOf,
  blitOver,
  canvasRect,
  composeFloating,
  createJobTracker,
  findLayer,
  grayOf,
  intersect,
  isAbortError,
  isBackgroundLayer,
  layerIndexOf,
  layerPixelBounds,
  liftFloating,
  messageOf,
  nextSelectionVersion,
  paintMask,
  paintPixels,
  readLayerRect,
  readMaskRect,
  resolvePaintTarget,
  union,
  writeRegion,
} from './tools/shared.ts'
import { TOOL_META } from './tools/index.ts'
import { warpMask, warpPixels } from './tools/transform.ts'
import { ADJUSTMENT_TYPES, adjustmentLabel, defaultAdjustment } from '../imaging/adjustments.ts'
import { FILTER_TYPES, defaultFilter, filterLabel, filterMargin } from '../imaging/filters.ts'
import { autoColor, autoContrast, autoTone } from '../imaging/autoEnhance.ts'
import { flipHorizontal, flipVertical, rotate180, rotate90 } from '../imaging/buffer.ts'
import { createMaskBuffer } from '../imaging/mask.ts'
import type { AdvancedEditorStore, DigitMemory } from './editorState.ts'
import {
  DEFAULT_BACKGROUND,
  DEFAULT_FOREGROUND,
  TOOL_IDS,
  TOOL_SLOTS,
  flowOptionKey,
  isToolId,
  opacityFromDigit,
  opacityOptionKey,
  sizeOptionKey,
  stepBrushSize,
  stepHardness,
} from './editorState.ts'

// ---------------------------------------------------------------------------------------------
// Ids
// ---------------------------------------------------------------------------------------------

export const DIGITS = Object.freeze(['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] as const)
export type Digit = typeof DIGITS[number]

export type PanelCommandId = 'view.panel-layers' | 'view.panel-properties' | 'view.panel-history' | 'view.panel-color'

/** Commands of this editor that the shared CommandId set has no ids for. */
export type EditorOnlyCommandId =
  | `tool.opacity-${Digit}`
  | `tool.flow-${Digit}`
  | PanelCommandId
  | 'filter.sharpen-more'

export type EditorCommandId = CommandId | EditorOnlyCommandId

const STATIC_COMMANDS = [
  'edit.undo', 'edit.redo', 'edit.toggle-last',
  'edit.cut', 'edit.copy', 'edit.copy-merged', 'edit.paste', 'edit.paste-in-place',
  'edit.clear', 'edit.fill-foreground', 'edit.fill-background', 'edit.fill-foreground-preserve',
  'edit.free-transform', 'edit.swap-colors', 'edit.default-colors',
  'select.all', 'select.deselect', 'select.reselect', 'select.inverse',
  'select.feather', 'select.expand', 'select.contract', 'select.load-layer-alpha',
  'layer.new', 'layer.duplicate', 'layer.delete', 'layer.via-copy', 'layer.via-cut',
  'layer.merge-down', 'layer.merge-visible', 'layer.stamp-visible', 'layer.flatten',
  'layer.add-mask', 'layer.delete-mask', 'layer.apply-mask', 'layer.toggle-mask', 'layer.rasterize',
  'layer.from-background', 'layer.toggle-clipping',
  'layer.raise', 'layer.lower', 'layer.to-front', 'layer.to-back',
  'layer.select-above', 'layer.select-below',
  'image.size', 'image.canvas-size', 'image.crop-to-selection', 'image.trim',
  'image.rotate-cw', 'image.rotate-ccw', 'image.rotate-180', 'image.rotate-arbitrary',
  'image.flip-horizontal', 'image.flip-vertical',
  'image.auto-tone', 'image.auto-contrast', 'image.auto-color', 'image.desaturate',
  'filter.repeat',
  'view.zoom-in', 'view.zoom-out', 'view.fit', 'view.actual-pixels', 'view.toggle-panels',
  'brush.smaller', 'brush.larger', 'brush.softer', 'brush.harder',
  'tool.cycle-marquee', 'tool.cycle-lasso', 'tool.cycle-gradient', 'tool.cycle-shape',
  'session.commit', 'session.cancel',
  'view.panel-layers', 'view.panel-properties', 'view.panel-history', 'view.panel-color',
  'filter.sharpen-more',
] as const satisfies readonly EditorCommandId[]

/** Every command the editor knows (menus, shortcuts and tests enumerate it). */
export const COMMAND_IDS: readonly EditorCommandId[] = Object.freeze([
  ...STATIC_COMMANDS,
  ...ADJUSTMENT_TYPES.map((type): EditorCommandId => `adjust.${type}`),
  ...ADJUSTMENT_TYPES.map((type): EditorCommandId => `adjustment-layer.${type}`),
  ...FILTER_TYPES.map((type): EditorCommandId => `filter.${type}`),
  ...TOOL_IDS.map((tool): EditorCommandId => `tool.${tool}`),
  ...DIGITS.map((digit): EditorCommandId => `tool.opacity-${digit}`),
  ...DIGITS.map((digit): EditorCommandId => `tool.flow-${digit}`),
])

const KNOWN = new Set<string>(COMMAND_IDS)

export function isEditorCommand(value: unknown): value is EditorCommandId {
  return typeof value === 'string' && KNOWN.has(value)
}

// ---------------------------------------------------------------------------------------------
// Context and services
// ---------------------------------------------------------------------------------------------

export type ModifySelectionOperation = 'feather' | 'expand' | 'contract'
export type CanvasAnchor = 'top-left' | 'top' | 'top-right' | 'left' | 'center' | 'right' | 'bottom-left' | 'bottom' | 'bottom-right'

export const CANVAS_ANCHORS: readonly CanvasAnchor[] = Object.freeze([
  'top-left', 'top', 'top-right', 'left', 'center', 'right', 'bottom-left', 'bottom', 'bottom-right',
] as const)

/** Commands that need parameters ask the editor for one of these dialogs. */
export type DialogRequest =
  | { readonly kind: 'image-size' }
  | { readonly kind: 'canvas-size' }
  | { readonly kind: 'rotate-canvas' }
  | { readonly kind: 'modify-selection'; readonly operation: ModifySelectionOperation }
  | { readonly kind: 'adjustment'; readonly type: AdjustmentType }
  | { readonly kind: 'filter'; readonly type: FilterType }

export type DialogKind = DialogRequest['kind']

/** DOM-side services the editor shell provides; every member is optional (Node tests run without them). */
export interface CommandServices {
  /** The controller receiving input (sessions: crop, polygonal lasso, text). */
  readonly activeTool?: () => AdvancedToolController | null
  /** The free-transform controller while a Ctrl+T session is open, or null. */
  readonly transformSession?: () => AdvancedToolController | null
  /** Starts Free Transform on the active layer (the controller tells the user when it cannot). */
  readonly startTransform?: () => void
  /** Opens a parameter dialog; false when this build has no dialog for the request. */
  readonly openDialog?: (request: DialogRequest) => boolean
  readonly canOpenDialog?: (request: DialogRequest) => boolean
  /** System clipboard (PNG encode / decode happens in the editor). */
  readonly clipboard?: {
    write(pixels: PixelBuffer): Promise<void>
    read(): Promise<PixelBuffer | null>
  }
  /** A tool is finishing an asynchronous commit (worker): document commands wait. */
  readonly busy?: () => boolean
}

/** ToolContext plus the editor's optional services. Any ToolContext is a valid CommandContext. */
export interface CommandContext extends ToolContext {
  readonly services?: CommandServices
}

// ---------------------------------------------------------------------------------------------
// Per-document runtime (Reselect memory, digit chains, the internal clipboard, worker jobs)
// ---------------------------------------------------------------------------------------------

interface ClipboardEntry {
  readonly pixels: PixelBuffer
  /** Document rectangle the pixels were copied from. */
  readonly rect: IntRect
}

interface CommandRuntime {
  readonly jobs: JobTracker
  /** The selection before the last deselect (Select > Reselect). */
  lastSelection: { readonly snapshot: SelectionSnapshot; readonly width: number; readonly height: number } | null
  previousSelection: Selection | null
  digits: DigitMemory | null
  clipboard: ClipboardEntry | null
  /** Panels shown before Tab hid them all. */
  hiddenPanels: { layers: boolean; properties: boolean; history: boolean; color: boolean } | null
  unsubscribe: (() => void) | null
}

type StoreKey = ToolContext['store']
const runtimes = new WeakMap<StoreKey, CommandRuntime>()

function runtimeOf(store: StoreKey): CommandRuntime {
  let runtime = runtimes.get(store)
  if (!runtime) {
    runtime = {
      jobs: createJobTracker(),
      lastSelection: null,
      previousSelection: store.getState().selection,
      digits: null,
      clipboard: null,
      hiddenPanels: null,
      unsubscribe: null,
    }
    runtimes.set(store, runtime)
  }
  return runtime
}

/**
 * Starts tracking the document for commands (call once when the editor opens): remembers every selection
 * that is removed, whoever removes it, so Select > Reselect brings it back. Returns the detach function.
 */
export function attachCommands(store: StoreKey): () => void {
  const runtime = runtimeOf(store)
  if (!runtime.unsubscribe) {
    runtime.previousSelection = store.getState().selection
    runtime.unsubscribe = store.subscribe((change) => {
      if (!change.selection) return
      const state = store.getState()
      const before = runtime.previousSelection
      if (before && !state.selection) {
        runtime.lastSelection = { snapshot: snapshotSelection(before), width: before.mask.width, height: before.mask.height }
      }
      runtime.previousSelection = state.selection
    })
  }
  return () => {
    runtime.unsubscribe?.()
    runtime.unsubscribe = null
  }
}

/** True while a command's worker job runs for this document. */
export function commandsBusy(store: StoreKey): boolean {
  return runtimes.get(store)?.jobs.busy ?? false
}

/** Resolves when no command job of this document is running (settle / Save wait for it). */
export function whenCommandsIdle(store: StoreKey): Promise<void> {
  return runtimes.get(store)?.jobs.whenIdle() ?? Promise.resolve()
}

/** The pixels copied last inside this editor (Ctrl+C / Ctrl+X), or null. */
export function internalClipboard(store: StoreKey): ClipboardEntry | null {
  return runtimes.get(store)?.clipboard ?? null
}

// ---------------------------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------------------------

function servicesOf(ctx: ToolContext): CommandServices {
  return (ctx as CommandContext).services ?? {}
}

function fail(ctx: ToolContext, message: string): false {
  ctx.host.notify(message, 'error')
  return false
}

function report(ctx: ToolContext, error: unknown, fallback: string): void {
  if (isAbortError(error)) return
  ctx.host.notify(messageOf(error, fallback), 'error')
}

function effectiveTool(ctx: ToolContext): ToolId {
  const editor = ctx.editor.getState()
  return editor.springFrom ?? editor.tool
}

function hasPixels(layer: Layer | null): layer is RasterLayer | TextLayer | ShapeLayer {
  return Boolean(layer && layer.kind !== 'adjustment')
}

/** Index a new layer goes to: just above the active layer (top when nothing is active). */
function indexAbove(state: DocumentState): number {
  const index = layerIndexOf(state, state.activeLayerId)
  return index >= 0 ? index + 1 : state.layers.length
}

function roomForLayers(ctx: ToolContext, count = 1): boolean {
  if (ctx.store.getState().layers.length + count <= LIMITS.maxLayers) return true
  return fail(ctx, `A document can have at most ${LIMITS.maxLayers} layers. Merge or delete some layers first.`)
}

function roomForBytes(ctx: ToolContext, bytes: number, action: string): boolean {
  const check = ensureMemory(ctx.store, bytes, action)
  return check.ok ? true : fail(ctx, check.message)
}

function rectBytes(rect: IntRect | null): number {
  return rect ? rect.width * rect.height * 4 : 0
}

function emptyPixels(width: number, height: number): PixelBuffer {
  return { width, height, data: new Uint8ClampedArray(Math.max(0, width * height * 4)) }
}

function solidPixels(width: number, height: number, color: Rgb8): PixelBuffer {
  const out = emptyPixels(width, height)
  const d = out.data
  for (let p = 0; p < d.length; p += 4) {
    d[p] = color.r
    d[p + 1] = color.g
    d[p + 2] = color.b
    d[p + 3] = 255
  }
  return out
}

/** Grey, opaque RGBA view of mask values (colour operations then work on masks too). */
function maskToPixels(mask: MaskBuffer): PixelBuffer {
  const out = emptyPixels(mask.width, mask.height)
  for (let i = 0, p = 0; i < mask.data.length; i += 1, p += 4) {
    const v = mask.data[i]
    out.data[p] = v
    out.data[p + 1] = v
    out.data[p + 2] = v
    out.data[p + 3] = 255
  }
  return out
}

/** Mask values back from RGBA (Rec. 601 luma). */
function pixelsToMask(pixels: PixelBuffer): MaskBuffer {
  const out = createMaskBuffer(pixels.width, pixels.height)
  const d = pixels.data
  for (let i = 0, p = 0; i < out.data.length; i += 1, p += 4) {
    out.data[i] = Math.round(0.299 * d[p] + 0.587 * d[p + 1] + 0.114 * d[p + 2])
  }
  return out
}

function cropPixels(src: PixelBuffer, rect: IntRect): PixelBuffer {
  const out = emptyPixels(rect.width, rect.height)
  for (let y = 0; y < rect.height; y += 1) {
    const sy = rect.y + y
    if (sy < 0 || sy >= src.height) continue
    const x0 = Math.max(0, rect.x)
    const x1 = Math.min(src.width, rect.x + rect.width)
    if (x1 <= x0) continue
    out.data.set(src.data.subarray((sy * src.width + x0) * 4, (sy * src.width + x1) * 4), (y * rect.width + (x0 - rect.x)) * 4)
  }
  return out
}

/** Effective visibility: hidden layers, and layers clipped to a hidden base, are not shown. */
function shownLayers(state: Pick<DocumentState, 'layers'>): Set<LayerId> {
  const shown = new Set<LayerId>()
  let baseVisible = true
  for (const layer of state.layers) {
    if (!layer.clipped || isBackgroundLayer(layer)) {
      baseVisible = layer.visible
      if (layer.visible) shown.add(layer.id)
    } else if (baseVisible && layer.visible) {
      shown.add(layer.id)
    }
  }
  return shown
}

/** Document-space area a layer contributes (the canvas for the Background), or null. */
function layerArea(state: Pick<DocumentState, 'width' | 'height'>, layer: Layer): IntRect | null {
  if (isBackgroundLayer(layer)) return canvasRect(state)
  return layerPixelBounds(layer)
}

/** A copy of `layer` moved by (dx, dy) with its mask (for compositing into a buffer at another origin). */
function shiftedLayer(layer: Layer, dx: number, dy: number): Layer {
  const mask = layer.mask ? { ...layer.mask, offsetX: layer.mask.offsetX + dx, offsetY: layer.mask.offsetY + dy } : null
  if (layer.kind === 'raster') return { ...layer, offsetX: layer.offsetX + dx, offsetY: layer.offsetY + dy, mask }
  if (layer.kind === 'text') return { ...layer, raster: { ...layer.raster, offsetX: layer.raster.offsetX + dx, offsetY: layer.raster.offsetY + dy }, mask }
  if (layer.kind === 'shape') return { ...layer, raster: { ...layer.raster, offsetX: layer.raster.offsetX + dx, offsetY: layer.raster.offsetY + dy }, mask }
  return { ...layer, mask }
}

/** Composites `layers` (bottom to top) over a document rectangle that may extend past the canvas. */
function compositeArea(layers: readonly Layer[], rect: IntRect): PixelBuffer {
  if (!(rect.width > 0 && rect.height > 0)) return emptyPixels(0, 0)
  const shifted = layers.map((layer) => shiftedLayer(layer, -rect.x, -rect.y))
  return compositeRect({ width: rect.width, height: rect.height, layers: shifted }, { x: 0, y: 0, width: rect.width, height: rect.height })
}

function alphaBounds(pixels: PixelBuffer): IntRect | null {
  let x0 = pixels.width
  let y0 = pixels.height
  let x1 = -1
  let y1 = -1
  const d = pixels.data
  for (let y = 0; y < pixels.height; y += 1) {
    let p = y * pixels.width * 4 + 3
    for (let x = 0; x < pixels.width; x += 1, p += 4) {
      if (d[p] === 0) continue
      if (x < x0) x0 = x
      if (x > x1) x1 = x
      if (y < y0) y0 = y
      y1 = y
    }
  }
  return x1 < 0 ? null : { x: x0, y: y0, width: x1 - x0 + 1, height: y1 - y0 + 1 }
}

// ---------------------------------------------------------------------------------------------
// Region edits (fills, destructive adjustments and filters on the active layer or its mask)
// ---------------------------------------------------------------------------------------------

interface Region {
  readonly layerId: LayerId
  readonly target: PaintTarget
  /** Document rectangle being changed. */
  readonly rect: IntRect
  readonly selection: Selection | null
  readonly surface: TiledSurface | TiledMask
  readonly version: number
}

/**
 * The part of the active layer (or its targeted mask) a command changes: the selection bounds, otherwise
 * the whole layer (the canvas for the Background). Null after telling the user why nothing can change.
 */
function regionOf(ctx: ToolContext, whole: 'content' | 'canvas'): Region | null {
  const target = resolvePaintTarget(ctx)
  if (!target) return null
  const state = ctx.store.getState()
  const layer = target.layer
  const selection = state.selection
  let rect: IntRect | null
  if (selection) {
    rect = selection.bounds
  } else if (target.channels === 1) {
    const mask = layer.mask as LayerMask
    const bounds = mask.surface.contentBounds()
    rect = union(canvasRect(state), bounds ? { x: bounds.x + mask.offsetX, y: bounds.y + mask.offsetY, width: bounds.width, height: bounds.height } : null)
  } else if (whole === 'canvas' || isBackgroundLayer(layer)) {
    rect = canvasRect(state)
  } else {
    rect = layerPixelBounds(layer)
  }
  if (!rect) {
    fail(ctx, `"${layer.name}" is empty; there are no pixels to change.`)
    return null
  }
  const surface = target.channels === 1 ? (layer.mask as LayerMask).surface : (layer as RasterLayer).surface
  return { layerId: layer.id, target, rect, selection, surface, version: surface.version }
}

function readRegion(region: Region, rect: IntRect = region.rect): PixelBuffer {
  if (region.target.channels === 1) return maskToPixels(readMaskRect(region.target.layer.mask as LayerMask, rect))
  return readLayerRect(region.target.layer, rect)
}

/** True when the layer (or mask) still holds exactly what the region was read from. */
function regionIsCurrent(state: DocumentState, region: Region): boolean {
  const layer = findLayer(state, region.layerId)
  if (!layer) return false
  if (region.target.channels === 1) {
    const mask = layer.mask
    return Boolean(mask && mask.surface === region.surface && mask.surface.version === region.version
      && mask.offsetX === region.target.origin.x && mask.offsetY === region.target.origin.y)
  }
  return layer.kind === 'raster' && layer.surface === region.surface && layer.surface.version === region.version
    && layer.offsetX === region.target.origin.x && layer.offsetY === region.target.origin.y
}

const CHANGED_MESSAGE = 'The layer changed while this was being applied, so nothing was changed. Try again.'

/** Writes `pixels` (region.rect-sized RGBA) back as one history step. */
function commitRegion(ctx: ToolContext, region: Region, pixels: PixelBuffer, label: string, icon: HistoryIcon): boolean {
  const data = region.target.channels === 1 ? pixelsToMask(pixels) : pixels
  ctx.store.transact(label, icon, (tx) => {
    if (!regionIsCurrent(tx.state, region)) throw new Error(CHANGED_MESSAGE)
    const editor = tx.editPixels(region.layerId, region.target.target)
    writeRegion(editor, region.target, region.rect, data)
  })
  return true
}

const STRIPE_ROWS = 256

/** Paints a solid colour through the selection in stripes (bounded memory), as one history step. */
function fillRegion(ctx: ToolContext, region: Region, color: Rgb8, options: { readonly erase?: boolean; readonly preserveAlpha?: boolean }, label: string): boolean {
  const { rect, target } = region
  ctx.store.transact(label, 'fill', (tx) => {
    if (!regionIsCurrent(tx.state, region)) throw new Error(CHANGED_MESSAGE)
    const editor = tx.editPixels(region.layerId, target.target)
    for (let top = rect.y; top < rect.y + rect.height; top += STRIPE_ROWS) {
      const stripe: IntRect = { x: rect.x, y: top, width: rect.width, height: Math.min(STRIPE_ROWS, rect.y + rect.height - top) }
      const coverage = new Float32Array(stripe.width * stripe.height).fill(1)
      const selection = region.selection ? selectionCoverage(region.selection, stripe) : null
      if (target.channels === 1) {
        const before = readMaskRect(target.layer.mask as LayerMask, stripe)
        writeRegion(editor, target, stripe, paintMask(before, coverage, grayOf(color), 1, selection))
      } else {
        const before = readLayerRect(target.layer, stripe)
        const after = paintPixels(before, coverage, {
          color,
          opacity: 1,
          mode: 'normal',
          erase: Boolean(options.erase),
          preserveAlpha: Boolean(options.preserveAlpha),
          selection,
        })
        writeRegion(editor, target, stripe, after)
      }
    }
  })
  return true
}

/** Fill (Alt+Backspace / Ctrl+Backspace): the selection, or the whole layer, with a colour. */
export function fillWithColor(ctx: ToolContext, color: Rgb8, preserveTransparency: boolean, label = 'Fill'): boolean {
  const region = regionOf(ctx, 'canvas')
  if (!region) return false
  return fillRegion(ctx, region, color, { preserveAlpha: preserveTransparency || region.target.preserveAlpha }, label)
}

/**
 * Image > Adjustments (destructive): applies `spec` to the active layer (or its mask) inside the selection,
 * as one history step. Resolves false when nothing was changed (the user was told why).
 */
export async function applyAdjustmentToLayer(ctx: ToolContext, spec: AdjustmentSpec, label = adjustmentLabel(spec.type)): Promise<boolean> {
  const region = regionOf(ctx, 'content')
  if (!region) return false
  return runtimeOf(ctx.store).jobs.run(async () => {
    try {
      const src = readRegion(region)
      const mask = region.selection ? selectionCoverage(region.selection, region.rect) : null
      const out = await ctx.imaging.run('adjust', { src, specs: [spec], mask, opacity: 1 })
      return commitRegion(ctx, region, out, label, 'adjustment')
    } catch (error) {
      report(ctx, error, `${label} could not be applied.`)
      return false
    }
  })
}

/**
 * Filter menu: applies `spec` to the active layer (or its mask) inside the selection, reading a margin of
 * surrounding pixels so edges blur into their neighbours like Photoshop. Remembers it for Last Filter.
 */
export async function applyFilterToLayer(ctx: ToolContext, spec: FilterSpec, label = filterLabel(spec.type, spec)): Promise<boolean> {
  const region = regionOf(ctx, 'content')
  if (!region) return false
  ctx.editor.update({ lastFilter: spec })
  return runtimeOf(ctx.store).jobs.run(async () => {
    try {
      const margin = Math.max(0, Math.ceil(filterMargin(spec)))
      const outer: IntRect = { x: region.rect.x - margin, y: region.rect.y - margin, width: region.rect.width + 2 * margin, height: region.rect.height + 2 * margin }
      const src = readRegion(region, outer)
      const alpha = region.target.channels === 4 && region.target.preserveAlpha ? alphaChannel(src) : null
      let mask: MaskBuffer | null = null
      if (region.selection || margin > 0) {
        // Coverage 0 in the margin: those pixels feed the filter but are never changed.
        mask = createMaskBuffer(outer.width, outer.height)
        const inner = region.selection ? selectionCoverage(region.selection, region.rect) : null
        for (let y = 0; y < region.rect.height; y += 1) {
          const row = (y + margin) * outer.width + margin
          if (inner) mask.data.set(inner.data.subarray(y * region.rect.width, (y + 1) * region.rect.width), row)
          else mask.data.fill(255, row, row + region.rect.width)
        }
      }
      const filtered = await ctx.imaging.run('filter', { src, spec, mask })
      if (alpha) restoreAlpha(filtered, alpha)
      const out = margin > 0 ? cropPixels(filtered, { x: margin, y: margin, width: region.rect.width, height: region.rect.height }) : filtered
      return commitRegion(ctx, region, out, label, 'filter')
    } catch (error) {
      report(ctx, error, `${label} could not be applied.`)
      return false
    }
  })
}

function alphaChannel(pixels: PixelBuffer): Uint8Array {
  const out = new Uint8Array(pixels.width * pixels.height)
  for (let i = 0, p = 3; i < out.length; i += 1, p += 4) out[i] = pixels.data[p]
  return out
}

function restoreAlpha(pixels: PixelBuffer, alpha: Uint8Array): void {
  for (let i = 0, p = 3; i < alpha.length; i += 1, p += 4) pixels.data[p] = alpha[i]
}

type AutoKind = 'tone' | 'contrast' | 'color'

async function autoAdjust(ctx: ToolContext, kind: AutoKind): Promise<boolean> {
  const region = regionOf(ctx, 'content')
  if (!region) return false
  const label = kind === 'tone' ? 'Auto Tone' : kind === 'contrast' ? 'Auto Contrast' : 'Auto Color'
  let spec: AdjustmentSpec
  try {
    const sample = readRegion(region)
    spec = kind === 'tone' ? autoTone(sample) : kind === 'contrast' ? autoContrast(sample) : autoColor(sample)
  } catch (error) {
    report(ctx, error, `${label} could not be applied.`)
    return false
  }
  return applyAdjustmentToLayer(ctx, spec, label)
}

// ---------------------------------------------------------------------------------------------
// Selection
// ---------------------------------------------------------------------------------------------

function setSelection(ctx: ToolContext, label: string, next: Selection | null): void {
  ctx.store.transact(label, 'selection', (tx) => tx.setSelection(next), { affectsOutput: false })
}

/** Select > Modify > Feather / Expand / Contract (in the imaging worker). */
export async function modifySelection(ctx: ToolContext, operation: ModifySelectionOperation, amount: number): Promise<boolean> {
  const start = ctx.store.getState()
  const selection = start.selection
  if (!selection) return fail(ctx, 'Make a selection first.')
  const value = Number(amount)
  if (!(Number.isFinite(value) && value > 0)) return fail(ctx, 'Enter a positive number of pixels.')
  const label = operation === 'feather' ? 'Feather' : operation === 'expand' ? 'Expand' : 'Contract'
  return runtimeOf(ctx.store).jobs.run(async () => {
    try {
      const mask: MaskBuffer = { width: selection.mask.width, height: selection.mask.height, data: new Uint8Array(selection.mask.data) }
      const result = operation === 'feather'
        ? await ctx.imaging.run('feather', { mask, radius: Math.min(1000, value) })
        : operation === 'expand'
          ? await ctx.imaging.run('expand', { mask, pixels: Math.min(500, Math.round(value)) })
          : await ctx.imaging.run('contract', { mask, pixels: Math.min(500, Math.round(value)) })
      const now = ctx.store.getState()
      if (now.selection !== selection || now.width !== start.width || now.height !== start.height) return false
      const next = selectionFromMask(result, nextSelectionVersion(now))
      if (!next && operation === 'contract') ctx.host.notify('Contracting removed the whole selection.')
      setSelection(ctx, label, next)
      return true
    } catch (error) {
      report(ctx, error, `${label} could not be applied.`)
      return false
    }
  })
}

// ---------------------------------------------------------------------------------------------
// Layers
// ---------------------------------------------------------------------------------------------

/** Adds an image as a new layer above the active one, centred in the view (or at `at`). Returns its id. */
export function placePixelsAsLayer(ctx: ToolContext, pixels: PixelBuffer, name: string, at: 'center' | Point = 'center', label = 'Paste'): LayerId | null {
  if (!(pixels && pixels.width > 0 && pixels.height > 0)) {
    fail(ctx, 'There is no image to place.')
    return null
  }
  if (!roomForLayers(ctx) || !roomForBytes(ctx, pixels.width * pixels.height * 4, 'place this image')) return null
  const state = ctx.store.getState()
  let x: number
  let y: number
  if (at === 'center') {
    const size = ctx.view.getViewportSize()
    const center = size.width > 0 && size.height > 0
      ? ctx.view.screenToDoc({ x: size.width / 2, y: size.height / 2 })
      : { x: state.width / 2, y: state.height / 2 }
    x = Math.round(center.x - pixels.width / 2)
    y = Math.round(center.y - pixels.height / 2)
  } else {
    x = Math.round(at.x)
    y = Math.round(at.y)
  }
  const layer = createRasterLayer({ name: nextLayerName(state, String(name || '').trim() || 'Layer'), surface: surfaceFromBuffer(pixels), offsetX: x, offsetY: y })
  ctx.store.transact(label, 'paste', (tx) => {
    tx.insertLayer(layer, indexAbove(tx.state))
    tx.setActiveLayer(layer.id)
  })
  return layer.id
}

function newLayer(ctx: ToolContext): void {
  if (!roomForLayers(ctx)) return
  const state = ctx.store.getState()
  const layer = createRasterLayer({ name: nextLayerName(state, 'Layer') })
  ctx.store.transact('New Layer', 'layer', (tx) => {
    tx.insertLayer(layer, indexAbove(tx.state))
    tx.setActiveLayer(layer.id)
  })
}

function duplicate(ctx: ToolContext, label = 'Duplicate Layer'): void {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!active) {
    fail(ctx, 'Select a layer to duplicate.')
    return
  }
  if (!roomForLayers(ctx)) return
  const bytes = active.kind === 'raster' ? active.surface.byteSize : active.kind === 'adjustment' ? 0 : active.raster.surface.byteSize
  if (!roomForBytes(ctx, bytes + (active.mask ? active.mask.surface.byteSize : 0), 'duplicate this layer')) return
  const copy = duplicateLayer(active, nextLayerName(state, `${active.name} copy`))
  ctx.store.transact(label, 'layer', (tx) => {
    tx.insertLayer(copy, layerIndexOf(tx.state, active.id) + 1)
    tx.setActiveLayer(copy.id)
  })
}

function deleteLayer(ctx: ToolContext): void {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!active) {
    fail(ctx, 'Select a layer to delete.')
    return
  }
  if (state.layers.length <= 1) {
    fail(ctx, 'A document needs at least one layer.')
    return
  }
  ctx.store.transact('Delete Layer', 'layer', (tx) => tx.removeLayer(active.id))
}

/** Ctrl+J / Ctrl+Shift+J: the selected pixels of the active layer on a new layer (cut leaves a hole). */
function layerVia(ctx: ToolContext, cut: boolean): void {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  const selection = state.selection
  if (!selection) {
    if (cut) {
      fail(ctx, 'Layer via Cut needs a selection.')
      return
    }
    duplicate(ctx, 'Layer Via Copy')
    return
  }
  if (!active) {
    fail(ctx, 'Select a layer first.')
    return
  }
  if (active.kind !== 'raster') {
    fail(ctx, active.kind === 'adjustment'
      ? `"${active.name}" is an adjustment layer; it has no pixels to copy.`
      : `Rasterize "${active.name}" first (Layer > Rasterize) to copy part of it.`)
    return
  }
  if (cut) {
    const blocker = pixelEditBlocker(active, 'pixels')
    if (blocker) {
      fail(ctx, blocker)
      return
    }
  }
  if (!roomForLayers(ctx) || !roomForBytes(ctx, rectBytes(selection.bounds) * (cut ? 2 : 1), 'copy the selection to a new layer')) return
  const fill = cut && (active.isBackground || active.locks.transparency) ? ctx.editor.getState().background : null
  const floating = liftFloating(state, active, selection, { cut, fill })
  if (!floating) {
    fail(ctx, 'The selected area is empty.')
    return
  }
  const layer = createRasterLayer({ name: nextLayerName(state, 'Layer'), surface: surfaceFromBuffer(floating.lifted), offsetX: floating.rect.x, offsetY: floating.rect.y })
  const hole = cut ? composeFloating(floating, null) : null
  ctx.store.transact(cut ? 'Layer Via Cut' : 'Layer Via Copy', 'layer', (tx) => {
    if (hole) {
      const editor = tx.editPixels(active.id, 'pixels')
      writeRegion(editor, { origin: { x: active.offsetX, y: active.offsetY }, channels: 4 }, hole.rect, hole.pixels)
    }
    tx.insertLayer(layer, layerIndexOf(tx.state, active.id) + 1)
    tx.setActiveLayer(layer.id)
  })
}

/** A raster layer showing `pixels` at `rect`, keeping the look-and-feel properties of `like`. */
function rasterLike(like: Layer, pixels: PixelBuffer, rect: IntRect, name = like.name): RasterLayer {
  return createRasterLayer({
    name,
    surface: surfaceFromBuffer(pixels),
    offsetX: rect.x,
    offsetY: rect.y,
    visible: like.visible,
    opacity: like.opacity,
    blendMode: like.blendMode,
    locks: isBackgroundLayer(like) ? DEFAULT_LOCKS : like.locks,
    clipped: like.clipped,
  })
}

function mergeDown(ctx: ToolContext): void {
  const state = ctx.store.getState()
  const upper = activeLayerOf(state)
  const index = layerIndexOf(state, state.activeLayerId)
  if (!upper || index <= 0) {
    fail(ctx, 'There is no layer below to merge into.')
    return
  }
  const lower = state.layers[index - 1]
  if (!upper.visible || !lower.visible) {
    fail(ctx, 'Show both layers before merging them.')
    return
  }
  if (lower.kind === 'adjustment') {
    fail(ctx, `"${lower.name}" is an adjustment layer; a layer cannot be merged into it.`)
    return
  }
  if (lower.locks.pixels) {
    fail(ctx, `"${lower.name}" is locked. Unlock its pixels first.`)
    return
  }
  const lowerArea = layerArea(state, lower)
  const area = isBackgroundLayer(lower)
    ? canvasRect(state)
    : upper.kind === 'adjustment' ? lowerArea : union(lowerArea, layerArea(state, upper))
  if (!area) {
    // Both layers are empty: merging only removes the upper one.
    ctx.store.transact('Merge Down', 'merge', (tx) => {
      tx.removeLayer(upper.id)
      tx.setActiveLayer(lower.id)
    })
    return
  }
  if (!roomForBytes(ctx, rectBytes(area), 'merge these layers')) return
  // The lower layer's own opacity, blend mode and mask stay its properties; its mask is applied to the result.
  const base = { ...lower, visible: true, opacity: 1, blendMode: 'normal' as const, clipped: false } as Layer
  const pixels = compositeArea([base, { ...upper, visible: true } as Layer], area)
  ctx.store.transact('Merge Down', 'merge', (tx) => {
    tx.removeLayer(upper.id)
    if (lower.kind === 'raster') {
      tx.replaceSurface(lower.id, 'pixels', surfaceFromBuffer(pixels), { x: area.x, y: area.y })
      if (lower.mask) tx.setMask(lower.id, null)
      tx.setActiveLayer(lower.id)
    } else {
      const at = layerIndexOf(tx.state, lower.id)
      const merged = rasterLike(lower, pixels, area)
      tx.removeLayer(lower.id)
      tx.insertLayer(merged, at)
      tx.setActiveLayer(merged.id)
    }
  })
}

function visibleComposite(ctx: ToolContext, state: DocumentState, label: string): { readonly area: IntRect; readonly pixels: PixelBuffer; readonly shown: Set<LayerId> } | null {
  const shown = shownLayers(state)
  let area: IntRect | null = null
  for (const layer of state.layers) if (shown.has(layer.id)) area = union(area, layerArea(state, layer))
  if (!area) {
    fail(ctx, `${label} needs at least one visible layer with pixels.`)
    return null
  }
  if (!roomForBytes(ctx, rectBytes(area), label.toLowerCase())) return null
  return { area, pixels: compositeArea(state.layers, area), shown }
}

function mergeVisible(ctx: ToolContext): void {
  const state = ctx.store.getState()
  const shown = shownLayers(state)
  if (shown.size < 2) {
    fail(ctx, 'Merge Visible needs at least two visible layers.')
    return
  }
  const bottom = state.layers[0]
  const keepBackground = isBackgroundLayer(bottom) && shown.has(bottom.id)
  const composite = visibleComposite(ctx, state, 'Merge Visible')
  if (!composite) return
  const active = activeLayerOf(state)
  const nameFrom = active && shown.has(active.id) ? active : [...state.layers].reverse().find((layer) => shown.has(layer.id)) as Layer
  ctx.store.transact('Merge Visible', 'merge', (tx) => {
    if (keepBackground) {
      const pixels = cropPixels(composite.pixels, { x: -composite.area.x, y: -composite.area.y, width: state.width, height: state.height })
      for (const layer of state.layers) if (layer.id !== bottom.id && shown.has(layer.id)) tx.removeLayer(layer.id)
      tx.replaceSurface(bottom.id, 'pixels', surfaceFromBuffer(pixels), { x: 0, y: 0 })
      tx.setActiveLayer(bottom.id)
      return
    }
    const lowest = state.layers.findIndex((layer) => shown.has(layer.id))
    const below = state.layers.slice(0, lowest).length
    for (const layer of state.layers) if (shown.has(layer.id)) tx.removeLayer(layer.id)
    const merged = createRasterLayer({ name: nameFrom.name, surface: surfaceFromBuffer(composite.pixels), offsetX: composite.area.x, offsetY: composite.area.y })
    tx.insertLayer(merged, below)
    tx.setActiveLayer(merged.id)
  })
}

function stampVisible(ctx: ToolContext): void {
  const state = ctx.store.getState()
  if (!roomForLayers(ctx)) return
  const composite = visibleComposite(ctx, state, 'Stamp Visible')
  if (!composite) return
  const layer = createRasterLayer({ name: nextLayerName(state, 'Layer'), surface: surfaceFromBuffer(composite.pixels), offsetX: composite.area.x, offsetY: composite.area.y })
  ctx.store.transact('Stamp Visible', 'merge', (tx) => {
    tx.insertLayer(layer, indexAbove(tx.state))
    tx.setActiveLayer(layer.id)
  })
}

function flatten(ctx: ToolContext): void {
  const state = ctx.store.getState()
  if (!roomForBytes(ctx, state.width * state.height * 4, 'flatten the image')) return
  // Photoshop flattens onto white and discards hidden layers.
  const canvas = canvasRect(state)
  const out = solidPixels(state.width, state.height, { r: 255, g: 255, b: 255 })
  blitOver(out, canvas, compositeRect(state, canvas), canvas)
  const background = createRasterLayer({ name: 'Background', isBackground: true, surface: surfaceFromBuffer(out) })
  ctx.store.transact('Flatten Image', 'merge', (tx) => {
    for (const layer of [...tx.state.layers].reverse()) tx.removeLayer(layer.id)
    tx.insertLayer(background, 0)
    tx.setActiveLayer(background.id)
    tx.setSelection(null)
  })
}

function layerFromBackground(ctx: ToolContext, tx: Transaction, layer: Layer): void {
  tx.updateLayer(layer.id, { isBackground: false, locks: DEFAULT_LOCKS, name: 'Layer 0' })
}

function addMask(ctx: ToolContext): void {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!active) {
    fail(ctx, 'Select a layer first.')
    return
  }
  if (active.mask) {
    fail(ctx, `"${active.name}" already has a layer mask.`)
    return
  }
  const mask = maskFromSelection(state.selection)
  ctx.store.transact('Add Layer Mask', 'layer', (tx) => {
    // Photoshop turns the Background into "Layer 0" when it gets a mask.
    if (isBackgroundLayer(active)) layerFromBackground(ctx, tx, active)
    tx.setMask(active.id, mask)
    tx.setActiveLayer(active.id, 'mask')
    if (state.selection) tx.setSelection(null)
  })
}

function applyMask(ctx: ToolContext): void {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!active || !active.mask) {
    fail(ctx, 'The active layer has no layer mask.')
    return
  }
  if (active.kind !== 'raster') {
    fail(ctx, active.kind === 'adjustment'
      ? 'An adjustment layer\'s mask cannot be applied; it is how the adjustment is limited.'
      : `Rasterize "${active.name}" first to apply its mask.`)
    return
  }
  if (active.locks.pixels) {
    fail(ctx, `"${active.name}" is locked. Unlock its pixels first.`)
    return
  }
  const mask = active.mask
  const bounds = layerPixelBounds(active)
  ctx.store.transact('Apply Layer Mask', 'layer', (tx) => {
    if (bounds) {
      const editor = tx.editPixels(active.id, 'pixels')
      for (let top = bounds.y; top < bounds.y + bounds.height; top += STRIPE_ROWS) {
        const stripe: IntRect = { x: bounds.x, y: top, width: bounds.width, height: Math.min(STRIPE_ROWS, bounds.y + bounds.height - top) }
        const pixels = readLayerRect(active, stripe)
        const values = readMaskRect(mask, stripe)
        for (let i = 0, p = 3; i < values.data.length; i += 1, p += 4) {
          const value = values.data[i]
          if (value === 255) continue
          pixels.data[p] = Math.round((pixels.data[p] * value) / 255)
        }
        writeRegion(editor, { origin: { x: active.offsetX, y: active.offsetY }, channels: 4 }, stripe, pixels)
      }
    }
    tx.setMask(active.id, null)
  })
}

function rasterize(ctx: ToolContext): void {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!active || (active.kind !== 'text' && active.kind !== 'shape')) {
    fail(ctx, 'Only type and shape layers can be rasterized.')
    return
  }
  const raster = rasterizeVectorLayer(active)
  ctx.store.transact('Rasterize Layer', 'layer', (tx) => {
    const at = layerIndexOf(tx.state, active.id)
    tx.removeLayer(active.id)
    tx.insertLayer(raster, at)
    tx.setActiveLayer(raster.id)
  })
}

function moveActive(ctx: ToolContext, where: 'raise' | 'lower' | 'front' | 'back'): void {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!active) return
  if (isBackgroundLayer(active)) {
    fail(ctx, 'The Background is locked at the bottom. Use Layer from Background to move it.')
    return
  }
  const index = layerIndexOf(state, active.id)
  const floor = isBackgroundLayer(state.layers[0]) ? 1 : 0
  const top = state.layers.length - 1
  const to = where === 'raise' ? index + 1 : where === 'lower' ? index - 1 : where === 'front' ? top : floor
  if (to < floor || to > top || to === index) return
  const label = where === 'raise' ? 'Bring Forward' : where === 'lower' ? 'Send Backward' : where === 'front' ? 'Bring to Front' : 'Send to Back'
  ctx.store.transact(label, 'layer', (tx) => tx.moveLayer(active.id, to))
}

function selectNeighbour(ctx: ToolContext, direction: 1 | -1): void {
  const state = ctx.store.getState()
  const index = layerIndexOf(state, state.activeLayerId)
  const next = state.layers[index < 0 ? (direction > 0 ? 0 : state.layers.length - 1) : index + direction]
  if (!next) return
  ctx.store.transact('Select Layer', 'layer', (tx) => tx.setActiveLayer(next.id), { affectsOutput: false })
}

function setLayerOpacity(ctx: ToolContext, opacity: number): void {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!active) return
  if (isBackgroundLayer(active)) {
    fail(ctx, 'The Background has no opacity. Use Layer from Background first.')
    return
  }
  ctx.store.transact('Opacity Change', 'layer', (tx) => tx.updateLayer(active.id, { opacity }), { coalesceKey: `layer-opacity:${active.id}` })
}

function newAdjustmentLayer(ctx: ToolContext, type: AdjustmentType): void {
  if (!roomForLayers(ctx)) return
  const state = ctx.store.getState()
  const label = adjustmentLabel(type)
  // Photoshop: the active selection becomes the adjustment layer's mask (and is then deselected).
  const layer = createAdjustmentLayer(nextLayerName(state, label), defaultAdjustment(type), maskFromSelection(state.selection))
  ctx.store.transact(`${label} Layer`, 'adjustment', (tx) => {
    tx.insertLayer(layer, indexAbove(tx.state))
    tx.setActiveLayer(layer.id)
    if (tx.state.selection) tx.setSelection(null)
  })
}

// ---------------------------------------------------------------------------------------------
// Whole-image operations
// ---------------------------------------------------------------------------------------------

export type CanvasTurn = 'cw' | 'ccw' | '180' | 'flip-h' | 'flip-v'

/** Where a document rectangle lands after turning a width x height canvas. */
export function turnRect(rect: IntRect, turn: CanvasTurn, width: number, height: number): IntRect {
  switch (turn) {
    case 'cw': return { x: height - (rect.y + rect.height), y: rect.x, width: rect.height, height: rect.width }
    case 'ccw': return { x: rect.y, y: width - (rect.x + rect.width), width: rect.height, height: rect.width }
    case '180': return { x: width - (rect.x + rect.width), y: height - (rect.y + rect.height), width: rect.width, height: rect.height }
    case 'flip-h': return { x: width - (rect.x + rect.width), y: rect.y, width: rect.width, height: rect.height }
    case 'flip-v': return { x: rect.x, y: height - (rect.y + rect.height), width: rect.width, height: rect.height }
  }
}

/** The document-space affine of a canvas turn (applies to text and shape transforms). */
export function turnAffine(turn: CanvasTurn, width: number, height: number): Affine {
  switch (turn) {
    case 'cw': return [0, 1, -1, 0, height, 0]
    case 'ccw': return [0, -1, 1, 0, 0, width]
    case '180': return [-1, 0, 0, -1, width, height]
    case 'flip-h': return [-1, 0, 0, 1, width, 0]
    case 'flip-v': return [1, 0, 0, -1, 0, height]
  }
}

function turnPixels(src: PixelBuffer, turn: CanvasTurn): PixelBuffer {
  switch (turn) {
    case 'cw': return rotate90(src, true)
    case 'ccw': return rotate90(src, false)
    case '180': return rotate180(src)
    case 'flip-h': return flipHorizontal(src)
    case 'flip-v': return flipVertical(src)
  }
}

/** The same exact permutation for one-channel masks. */
export function turnMask(src: MaskBuffer, turn: CanvasTurn): MaskBuffer {
  const { width: w, height: h, data } = src
  const quarter = turn === 'cw' || turn === 'ccw'
  const out = createMaskBuffer(quarter ? h : w, quarter ? w : h)
  const o = out.data
  for (let y = 0; y < h; y += 1) {
    const row = y * w
    for (let x = 0; x < w; x += 1) {
      const value = data[row + x]
      switch (turn) {
        case 'cw': o[x * h + (h - 1 - y)] = value; break
        case 'ccw': o[(w - 1 - x) * h + y] = value; break
        case '180': o[(h - 1 - y) * w + (w - 1 - x)] = value; break
        case 'flip-h': o[row + (w - 1 - x)] = value; break
        case 'flip-v': o[(h - 1 - y) * w + x] = value; break
      }
    }
  }
  return out
}

const TURN_LABELS: Readonly<Record<CanvasTurn, string>> = Object.freeze({
  cw: 'Rotate Canvas',
  ccw: 'Rotate Canvas',
  '180': 'Rotate Canvas',
  'flip-h': 'Flip Canvas Horizontal',
  'flip-v': 'Flip Canvas Vertical',
})

/** Image > Image Rotation (90 / 180 / flips): exact pixel permutations of every layer, mask and the selection. */
export function turnCanvas(ctx: ToolContext, turn: CanvasTurn): boolean {
  const state = ctx.store.getState()
  const width = state.width
  const height = state.height
  let bytes = 0
  for (const layer of state.layers) {
    if (layer.kind === 'raster') bytes += layer.surface.byteSize
    if (layer.mask) bytes += layer.mask.surface.byteSize
  }
  if (!roomForBytes(ctx, bytes, 'rotate the image')) return false
  const affine = turnAffine(turn, width, height)
  type Step = (tx: Transaction) => void
  const steps: Step[] = []
  for (const layer of state.layers) {
    if (layer.kind === 'raster') {
      const bounds = layer.surface.contentBounds()
      if (bounds) {
        const placed = turnRect({ x: bounds.x + layer.offsetX, y: bounds.y + layer.offsetY, width: bounds.width, height: bounds.height }, turn, width, height)
        const surface = surfaceFromBuffer(turnPixels(layer.surface.read(bounds), turn))
        steps.push((tx) => tx.replaceSurface(layer.id, 'pixels', surface, { x: placed.x, y: placed.y }))
      }
    } else if (layer.kind === 'text') {
      const text = { ...layer.text, transform: composeAffine(affine, layer.text.transform) }
      steps.push((tx) => tx.updateLayer(layer.id, { text }))
    } else if (layer.kind === 'shape') {
      const shape = { ...layer.shape, transform: composeAffine(affine, layer.shape.transform) }
      steps.push((tx) => tx.updateLayer(layer.id, { shape }))
    }
    const mask = layer.mask
    if (mask) {
      const bounds = mask.surface.contentBounds()
      if (bounds) {
        const placed = turnRect({ x: bounds.x + mask.offsetX, y: bounds.y + mask.offsetY, width: bounds.width, height: bounds.height }, turn, width, height)
        const surface = maskFromBuffer(turnMask(mask.surface.read(bounds), turn), mask.surface.defaultValue)
        steps.push((tx) => tx.replaceSurface(layer.id, 'mask', surface, { x: placed.x, y: placed.y }))
      }
    }
  }
  const selection = state.selection ? selectionFromMask(turnMask(state.selection.mask, turn), nextSelectionVersion(state)) : null
  const quarter = turn === 'cw' || turn === 'ccw'
  ctx.store.transact(TURN_LABELS[turn], 'image', (tx) => {
    for (const step of steps) step(tx)
    tx.setCanvas(quarter ? height : width, quarter ? width : height)
    if (selection || tx.state.selection) tx.setSelection(selection)
  })
  return true
}

/**
 * Crops the canvas to `rect` (document space): pixel layers are trimmed to it (Photoshop's "delete cropped
 * pixels"), type and shape layers and masks move with the canvas; the selection is dropped.
 */
export function cropCanvasTo(ctx: ToolContext, rect: IntRect, label = 'Crop'): boolean {
  const state = ctx.store.getState()
  const crop = intersect(rect, canvasRect(state))
  if (!crop) return fail(ctx, 'The crop area is outside the image.')
  const bx = crop.x
  const by = crop.y
  ctx.store.transact(label, 'crop', (tx) => {
    for (const layer of state.layers) {
      if (layer.kind === 'raster') {
        const bounds = isBackgroundLayer(layer) ? crop : layerPixelBounds(layer)
        const keep = bounds ? intersect(bounds, crop) : null
        if (keep) {
          const pixels = readLayerRect(layer, keep)
          tx.replaceSurface(layer.id, 'pixels', surfaceFromBuffer(pixels), { x: keep.x - bx, y: keep.y - by })
        } else {
          tx.replaceSurface(layer.id, 'pixels', createSurface(), { x: 0, y: 0 })
        }
      } else if (layer.kind === 'text' || layer.kind === 'shape') {
        tx.updateLayer(layer.id, { offset: { x: layer.raster.offsetX - bx, y: layer.raster.offsetY - by } })
      }
      if (layer.mask) {
        const current = findLayer(tx.state, layer.id)?.mask
        const x = layer.mask.offsetX - bx
        const y = layer.mask.offsetY - by
        if (current && (current.offsetX !== x || current.offsetY !== y)) tx.setMask(layer.id, { ...current, offsetX: x, offsetY: y })
      }
    }
    tx.setCanvas(crop.width, crop.height)
    tx.setSelection(null)
  })
  return true
}

async function trim(ctx: ToolContext): Promise<boolean> {
  const state = ctx.store.getState()
  if (isBackgroundLayer(state.layers[0]) && state.layers[0].visible) return fail(ctx, 'The image has no transparent pixels to trim.')
  const flat = await ctx.compositor.flatten()
  if (ctx.store.getState() !== state) return false
  const bounds = alphaBounds(flat)
  if (!bounds) return fail(ctx, 'The image is completely transparent; there is nothing to keep.')
  if (bounds.width === state.width && bounds.height === state.height) return fail(ctx, 'The image has no transparent edges to trim.')
  return cropCanvasTo(ctx, bounds, 'Trim')
}

function anchorFactors(anchor: CanvasAnchor): { readonly fx: number; readonly fy: number } {
  const fx = anchor.endsWith('left') ? 0 : anchor.endsWith('right') ? 1 : 0.5
  const fy = anchor.startsWith('top') ? 0 : anchor.startsWith('bottom') ? 1 : 0.5
  return { fx, fy }
}

function validSize(width: number, height: number): boolean {
  return Number.isInteger(width) && Number.isInteger(height) && width > 0 && height > 0
    && width <= LIMITS.maxDimension && height <= LIMITS.maxDimension && width * height <= LIMITS.maxPixels
}

const SIZE_MESSAGE = `The size must be whole pixels, at most ${LIMITS.maxDimension.toLocaleString('en-US')} per side and ${LIMITS.maxPixels / 1_000_000} megapixels in total.`

/**
 * Image > Canvas Size: the canvas grows or shrinks around `anchor`; every layer keeps its pixels (outside
 * the canvas too), and the Background is extended with the background colour (and trimmed to the canvas).
 */
export function resizeCanvas(ctx: ToolContext, width: number, height: number, anchor: CanvasAnchor = 'center'): boolean {
  if (!validSize(width, height)) return fail(ctx, SIZE_MESSAGE)
  const state = ctx.store.getState()
  if (width === state.width && height === state.height) return true
  const { fx, fy } = anchorFactors(anchor)
  const dx = Math.round((width - state.width) * fx)
  const dy = Math.round((height - state.height) * fy)
  const fill = ctx.editor.getState().background
  const bottom = state.layers[0]
  let backgroundPixels: PixelBuffer | null = null
  if (bottom && isBackgroundLayer(bottom)) {
    if (!roomForBytes(ctx, width * height * 4, 'resize the canvas')) return false
    backgroundPixels = solidPixels(width, height, fill)
    const old = readLayerRect(bottom, canvasRect(state))
    const destination: IntRect = { x: dx, y: dy, width: state.width, height: state.height }
    const target = canvasRect({ width, height })
    const overlap = intersect(destination, target)
    if (overlap) {
      for (let y = overlap.y; y < overlap.y + overlap.height; y += 1) {
        const from = ((y - dy) * state.width + (overlap.x - dx)) * 4
        backgroundPixels.data.set(old.data.subarray(from, from + overlap.width * 4), (y * width + overlap.x) * 4)
      }
    }
  }
  ctx.store.transact('Canvas Size', 'image', (tx) => {
    for (const layer of state.layers) {
      if (layer.kind === 'raster') {
        if (isBackgroundLayer(layer) && backgroundPixels) tx.replaceSurface(layer.id, 'pixels', surfaceFromBuffer(backgroundPixels), { x: 0, y: 0 })
        else tx.updateLayer(layer.id, { offset: { x: layer.offsetX + dx, y: layer.offsetY + dy } })
      } else if (layer.kind === 'text' || layer.kind === 'shape') {
        tx.updateLayer(layer.id, { offset: { x: layer.raster.offsetX + dx, y: layer.raster.offsetY + dy } })
      }
      if (layer.mask) {
        const current = findLayer(tx.state, layer.id)?.mask
        const x = layer.mask.offsetX + dx
        const y = layer.mask.offsetY + dy
        if (current && (current.offsetX !== x || current.offsetY !== y)) tx.setMask(layer.id, { ...current, offsetX: x, offsetY: y })
      }
    }
    tx.setCanvas(width, height)
  })
  return true
}

function scaleRect(rect: IntRect, sx: number, sy: number): IntRect {
  const x0 = Math.round(rect.x * sx)
  const y0 = Math.round(rect.y * sy)
  const x1 = Math.round((rect.x + rect.width) * sx)
  const y1 = Math.round((rect.y + rect.height) * sy)
  return { x: x0, y: y0, width: Math.max(1, x1 - x0), height: Math.max(1, y1 - y0) }
}

/**
 * Image > Image Size: resamples every pixel layer and mask (imaging worker), scales type and shape layers
 * (they stay editable) and offsets; one history step. `method` defaults to 'auto' (area when shrinking).
 */
export async function resizeImage(ctx: ToolContext, width: number, height: number, options: { readonly method?: ResampleMethod; readonly ppi?: number } = {}): Promise<boolean> {
  if (!validSize(width, height)) return fail(ctx, SIZE_MESSAGE)
  const start = ctx.store.getState()
  const ppi = Number.isFinite(options.ppi) && Number(options.ppi) > 0 ? Number(options.ppi) : start.ppi
  if (width === start.width && height === start.height) {
    if (ppi !== start.ppi) ctx.store.transact('Image Size', 'image', (tx) => tx.setCanvas(width, height, ppi))
    return true
  }
  const sx = width / start.width
  const sy = height / start.height
  let bytes = 0
  for (const layer of start.layers) {
    if (layer.kind === 'raster') {
      const bounds = layerPixelBounds(layer)
      if (bounds) bytes += rectBytes(scaleRect(bounds, sx, sy))
    }
  }
  if (!roomForBytes(ctx, bytes, 'resize the image')) return false
  const method = options.method ?? 'auto'
  return runtimeOf(ctx.store).jobs.run(async () => {
    try {
      type Step = (tx: Transaction) => void
      const steps: Step[] = []
      const scale: Affine = [sx, 0, 0, sy, 0, 0]
      for (const layer of start.layers) {
        if (layer.kind === 'raster') {
          const bounds = isBackgroundLayer(layer) ? canvasRect(start) : layerPixelBounds(layer)
          if (bounds) {
            const target = isBackgroundLayer(layer) ? { x: 0, y: 0, width, height } : scaleRect(bounds, sx, sy)
            const src = readLayerRect(layer, bounds)
            const out = await ctx.imaging.run('resample', { src, width: target.width, height: target.height, method })
            const surface = surfaceFromBuffer(out)
            steps.push((tx) => tx.replaceSurface(layer.id, 'pixels', surface, { x: target.x, y: target.y }))
          }
        } else if (layer.kind === 'text') {
          const text = { ...layer.text, transform: composeAffine(scale, layer.text.transform) }
          steps.push((tx) => tx.updateLayer(layer.id, { text }))
        } else if (layer.kind === 'shape') {
          const shape = { ...layer.shape, transform: composeAffine(scale, layer.shape.transform) }
          steps.push((tx) => tx.updateLayer(layer.id, { shape }))
        }
        const mask = layer.mask
        if (mask) {
          const local = mask.surface.contentBounds()
          if (local) {
            const bounds = { x: local.x + mask.offsetX, y: local.y + mask.offsetY, width: local.width, height: local.height }
            const target = scaleRect(bounds, sx, sy)
            const src = maskToPixels(mask.surface.read(local))
            const out = await ctx.imaging.run('resample', { src, width: target.width, height: target.height, method })
            const surface = maskFromBuffer(pixelsToMask(out), mask.surface.defaultValue)
            steps.push((tx) => tx.replaceSurface(layer.id, 'mask', surface, { x: target.x, y: target.y }))
          } else {
            const x = Math.round(mask.offsetX * sx)
            const y = Math.round(mask.offsetY * sy)
            steps.push((tx) => {
              const current = findLayer(tx.state, layer.id)?.mask
              if (current) tx.setMask(layer.id, { ...current, offsetX: x, offsetY: y })
            })
          }
        }
      }
      if (ctx.store.getState().pixelVersion !== start.pixelVersion || ctx.store.getState().layers !== start.layers) {
        throw new Error('The image changed while it was being resized, so nothing was changed. Try again.')
      }
      ctx.store.transact('Image Size', 'image', (tx) => {
        for (const step of steps) step(tx)
        tx.setCanvas(width, height, ppi)
      })
      return true
    } catch (error) {
      report(ctx, error, 'The image could not be resized.')
      return false
    }
  })
}

/**
 * Image > Image Rotation > Arbitrary: the canvas grows to hold the turned image; every layer turns about
 * the canvas centre (bicubic, in the worker); the Background's new corners take the background colour.
 */
export async function rotateCanvas(ctx: ToolContext, degrees: number): Promise<boolean> {
  const value = Number(degrees)
  if (!Number.isFinite(value)) return fail(ctx, 'Enter an angle in degrees.')
  const normalized = ((value % 360) + 360) % 360
  if (Math.abs(normalized) < 1e-9 || Math.abs(normalized - 360) < 1e-9) return true
  for (const [angle, turn] of [[90, 'cw'], [180, '180'], [270, 'ccw']] as const) {
    if (Math.abs(normalized - angle) < 1e-9) return turnCanvas(ctx, turn)
  }
  const start = ctx.store.getState()
  const radians = (normalized * Math.PI) / 180
  const cos = Math.cos(radians)
  const sin = Math.sin(radians)
  const width = Math.max(1, Math.round(Math.abs(start.width * cos) + Math.abs(start.height * sin)))
  const height = Math.max(1, Math.round(Math.abs(start.width * sin) + Math.abs(start.height * cos)))
  if (!validSize(width, height)) return fail(ctx, SIZE_MESSAGE)
  // Document -> new document: centre, rotate, re-centre in the larger canvas.
  const cx = start.width / 2
  const cy = start.height / 2
  const e = width / 2 - (cos * cx - sin * cy)
  const f = height / 2 - (sin * cx + cos * cy)
  const affine: Affine = [cos, sin, -sin, cos, e, f]
  const forward = [cos, -sin, e, sin, cos, f, 0, 0, 1] as const
  let bytes = 0
  for (const layer of start.layers) {
    if (layer.kind === 'raster') bytes += layer.surface.byteSize * 2
  }
  if (!roomForBytes(ctx, bytes, 'rotate the image')) return false
  const fill = ctx.editor.getState().background
  return runtimeOf(ctx.store).jobs.run(async () => {
    try {
      type Step = (tx: Transaction) => void
      const steps: Step[] = []
      for (const layer of start.layers) {
        if (layer.kind === 'raster') {
          const bounds = isBackgroundLayer(layer) ? canvasRect(start) : layerPixelBounds(layer)
          if (bounds) {
            const placed = await warpPixels(ctx, readLayerRect(layer, bounds), bounds.x, bounds.y, forward)
            if (isBackgroundLayer(layer)) {
              const out = solidPixels(width, height, fill)
              if (placed) blitOver(out, { x: 0, y: 0, width, height }, placed.pixels, placed.rect)
              const surface = surfaceFromBuffer(out)
              steps.push((tx) => tx.replaceSurface(layer.id, 'pixels', surface, { x: 0, y: 0 }))
            } else if (placed) {
              const surface = surfaceFromBuffer(placed.pixels)
              steps.push((tx) => tx.replaceSurface(layer.id, 'pixels', surface, { x: placed.rect.x, y: placed.rect.y }))
            }
          }
        } else if (layer.kind === 'text') {
          const text = { ...layer.text, transform: composeAffine(affine, layer.text.transform) }
          steps.push((tx) => tx.updateLayer(layer.id, { text }))
        } else if (layer.kind === 'shape') {
          const shape = { ...layer.shape, transform: composeAffine(affine, layer.shape.transform) }
          steps.push((tx) => tx.updateLayer(layer.id, { shape }))
        }
        if (layer.mask) {
          const warped = await warpMask(ctx, layer.mask, forward)
          if (warped) steps.push((tx) => tx.replaceSurface(layer.id, 'mask', warped.surface, warped.offset))
        }
      }
      if (ctx.store.getState().pixelVersion !== start.pixelVersion || ctx.store.getState().layers !== start.layers) {
        throw new Error('The image changed while it was being rotated, so nothing was changed. Try again.')
      }
      ctx.store.transact('Rotate Canvas', 'image', (tx) => {
        for (const step of steps) step(tx)
        tx.setCanvas(width, height)
      })
      return true
    } catch (error) {
      report(ctx, error, 'The image could not be rotated.')
      return false
    }
  })
}

// ---------------------------------------------------------------------------------------------
// Clipboard
// ---------------------------------------------------------------------------------------------

/** The pixels Copy takes: the active layer (or the merged image) inside the selection; null when empty. */
export function copyPixels(state: DocumentState, merged: boolean): ClipboardEntry | null {
  const selection = state.selection
  const active = activeLayerOf(state)
  let rect: IntRect | null
  if (selection) rect = selection.bounds
  else if (merged || !active) rect = canvasRect(state)
  else rect = intersect(layerArea(state, active) ?? canvasRect(state), canvasRect(state))
  if (!rect) return null
  let pixels: PixelBuffer
  if (merged) pixels = compositeRect(state, rect)
  else if (active && active.kind !== 'adjustment') pixels = readLayerRect(active, rect)
  else return null
  if (selection) {
    const coverage = selectionCoverage(selection, rect)
    for (let i = 0, p = 3; i < coverage.data.length; i += 1, p += 4) {
      const c = coverage.data[i]
      if (c === 255) continue
      pixels.data[p] = Math.round((pixels.data[p] * c) / 255)
    }
  }
  // Photoshop copies only the visible pixels: trim fully transparent edges.
  const bounds = alphaBounds(pixels)
  if (!bounds) return null
  if (bounds.width !== pixels.width || bounds.height !== pixels.height) {
    pixels = cropPixels(pixels, bounds)
    rect = { x: rect.x + bounds.x, y: rect.y + bounds.y, width: bounds.width, height: bounds.height }
  }
  return { pixels, rect }
}

async function copy(ctx: ToolContext, merged: boolean): Promise<boolean> {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!merged && (!active || active.kind === 'adjustment')) {
    return fail(ctx, active ? `"${active.name}" is an adjustment layer; use Copy Merged to copy what you see.` : 'Select a layer to copy.')
  }
  const entry = copyPixels(state, merged)
  if (!entry) return fail(ctx, 'The selected area is empty.')
  runtimeOf(ctx.store).clipboard = entry
  const clipboard = servicesOf(ctx).clipboard
  if (clipboard) await clipboard.write(entry.pixels)
  return true
}

async function cut(ctx: ToolContext): Promise<boolean> {
  const state = ctx.store.getState()
  const active = activeLayerOf(state)
  if (!active || active.kind !== 'raster') {
    return fail(ctx, active && active.kind !== 'adjustment' ? `Rasterize "${active.name}" first to cut from it.` : 'Select a pixel layer to cut from.')
  }
  const blocker = pixelEditBlocker(active, 'pixels')
  if (blocker) return fail(ctx, blocker)
  if (!state.selection) {
    // Photoshop: Cut without a selection takes the whole layer.
    if (!(await copy(ctx, false))) return false
    if (isBackgroundLayer(active) || active.locks.transparency) {
      const region = regionOf(ctx, 'canvas')
      return region ? fillRegion(ctx, region, ctx.editor.getState().background, { preserveAlpha: true }, 'Cut') : false
    }
    ctx.store.transact('Cut', 'layer', (tx) => tx.replaceSurface(active.id, 'pixels', createSurface(), { x: active.offsetX, y: active.offsetY }))
    return true
  }
  if (!(await copy(ctx, false))) return false
  return clearSelected(ctx, 'Cut')
}

async function paste(ctx: ToolContext, inPlace: boolean): Promise<boolean> {
  const clipboard = servicesOf(ctx).clipboard
  const internal = runtimeOf(ctx.store).clipboard
  let pixels: PixelBuffer | null = null
  if (clipboard) {
    try {
      pixels = await clipboard.read()
    } catch (error) {
      report(ctx, error, 'The clipboard could not be read.')
    }
  }
  let rect: IntRect | null = null
  if (!pixels && internal) pixels = internal.pixels
  if (pixels && internal && samePixels(pixels, internal.pixels)) rect = internal.rect
  if (!pixels) return fail(ctx, 'The clipboard has no image to paste.')
  const at: 'center' | Point = inPlace && rect ? { x: rect.x, y: rect.y } : 'center'
  return placePixelsAsLayer(ctx, pixels, 'Layer', at, inPlace ? 'Paste in Place' : 'Paste') !== null
}

function samePixels(a: PixelBuffer, b: PixelBuffer): boolean {
  if (a.width !== b.width || a.height !== b.height || a.data.length !== b.data.length) return false
  for (let i = 0; i < a.data.length; i += 1) if (a.data[i] !== b.data[i]) return false
  return true
}

/** Clears the selected pixels (the Background and locked transparency take the background colour). */
function clearSelected(ctx: ToolContext, label = 'Clear'): boolean {
  const region = regionOf(ctx, 'canvas')
  if (!region) return false
  const background = ctx.editor.getState().background
  if (region.target.channels === 1) return fillRegion(ctx, region, background, {}, label)
  if (region.target.preserveAlpha) return fillRegion(ctx, region, background, { preserveAlpha: true }, label)
  return fillRegion(ctx, region, background, { erase: true }, label)
}

function clear(ctx: ToolContext): void {
  const state = ctx.store.getState()
  if (state.selection) {
    clearSelected(ctx)
    return
  }
  // Photoshop: Delete without a selection deletes the active layer.
  deleteLayer(ctx)
}

// ---------------------------------------------------------------------------------------------
// Tools, brushes, colours, view
// ---------------------------------------------------------------------------------------------

function selectTool(ctx: ToolContext, tool: ToolId): void {
  ctx.editor.update({ tool, springFrom: null })
}

function cycleSlot(ctx: ToolContext, slotIndex: number): void {
  const slot = TOOL_SLOTS[slotIndex]
  const editor = ctx.editor as Partial<AdvancedEditorStore> & ToolContext['editor']
  const current = typeof editor.slotTool === 'function' ? editor.slotTool(slot[0]) : (slot.includes(effectiveTool(ctx)) ? effectiveTool(ctx) : slot[0])
  const active = effectiveTool(ctx)
  // Photoshop: Shift+key moves to the next tool of the slot; when another tool is active it first shows the slot's tool.
  const next = slot.includes(active) ? slot[(slot.indexOf(current) + 1) % slot.length] : current
  selectTool(ctx, next)
}

const SHAPE_ORDER = ['rectangle', 'ellipse', 'line', 'arrow'] as const

function cycleShape(ctx: ToolContext): void {
  const options = ctx.editor.getState().options.shape
  if (effectiveTool(ctx) === 'shape') {
    const next = SHAPE_ORDER[(SHAPE_ORDER.indexOf(options.kind) + 1) % SHAPE_ORDER.length]
    ctx.editor.updateOptions('shape', { kind: next })
  }
  selectTool(ctx, 'shape')
}

function stepBrush(ctx: ToolContext, kind: 'size' | 'hardness', direction: 1 | -1): void {
  const key = sizeOptionKey(effectiveTool(ctx))
  if (!key) return
  const options = ctx.editor.getState().options[key]
  if (kind === 'size') ctx.editor.updateOptions(key, { size: stepBrushSize(options.size, direction) })
  else ctx.editor.updateOptions(key, { hardness: stepHardness(options.hardness, direction) })
  ctx.view.requestOverlay()
}

function digitCommand(ctx: ToolContext, kind: 'opacity' | 'flow', digit: number): void {
  const runtime = runtimeOf(ctx.store)
  const tool = effectiveTool(ctx)
  const now = Date.now()
  const optionKey = kind === 'opacity' ? opacityOptionKey(tool) : flowOptionKey(tool)
  const target = optionKey ? `${kind}:${optionKey}` : `layer:${ctx.store.getState().activeLayerId ?? ''}`
  const { value, memory } = opacityFromDigit(digit, now, target, runtime.digits)
  runtime.digits = memory
  if (optionKey) {
    if (kind === 'opacity') ctx.editor.updateOptions(optionKey, { opacity: value })
    else ctx.editor.updateOptions(optionKey as 'brush' | 'eraser' | 'clone', { flow: Math.max(0.01, value) })
    return
  }
  // Photoshop: with a tool that has no opacity (Move, selections...) the digits set the layer opacity.
  if (kind === 'opacity') setLayerOpacity(ctx, value)
}

function zoomStep(ctx: ToolContext, direction: 1 | -1): void {
  const size = ctx.view.getViewportSize()
  const view = ctx.view.getView()
  ctx.view.zoomAt(stepZoom(view.zoom, size.dpr, direction), { x: size.width / 2, y: size.height / 2 })
}

function togglePanels(ctx: ToolContext): void {
  const runtime = runtimeOf(ctx.store)
  const panels = ctx.editor.getState().panels
  const anyShown = panels.layers || panels.properties || panels.history || panels.color
  if (anyShown) {
    runtime.hiddenPanels = { ...panels }
    ctx.editor.update({ panels: { layers: false, properties: false, history: false, color: false } })
  } else {
    const restore = runtime.hiddenPanels ?? { layers: true, properties: true, history: true, color: true }
    runtime.hiddenPanels = null
    ctx.editor.update({ panels: restore })
  }
}

function togglePanel(ctx: ToolContext, panel: 'layers' | 'properties' | 'history' | 'color'): void {
  const panels = ctx.editor.getState().panels
  ctx.editor.update({ panels: { ...panels, [panel]: !panels[panel] } })
}

function sessionController(ctx: ToolContext): AdvancedToolController | null {
  const services = servicesOf(ctx)
  const transform = services.transformSession?.() ?? null
  if (transform && transform.hasSession()) return transform
  const tool = services.activeTool?.() ?? null
  return tool && tool.hasSession() ? tool : null
}

function openDialog(ctx: ToolContext, request: DialogRequest, name: string): void {
  const services = servicesOf(ctx)
  if (services.openDialog && services.openDialog(request)) return
  fail(ctx, `${name} is not available in this version yet.`)
}

function dialogAvailable(ctx: ToolContext, request: DialogRequest): boolean {
  const services = servicesOf(ctx)
  if (!services.openDialog) return false
  return services.canOpenDialog ? services.canOpenDialog(request) : true
}

// ---------------------------------------------------------------------------------------------
// The registry
// ---------------------------------------------------------------------------------------------

type Run = (ctx: CommandContext) => void | boolean | Promise<unknown>

interface CommandSpec {
  readonly run: Run
  readonly enabled?: (ctx: CommandContext, state: DocumentState) => boolean
  /** Changes the document: refused while a worker job or an asynchronous tool commit is running. */
  readonly document?: boolean
}

function activeIs(state: DocumentState, test: (layer: Layer) => boolean): boolean {
  const layer = activeLayerOf(state)
  return Boolean(layer && test(layer))
}

function editablePixels(state: DocumentState): boolean {
  const layer = activeLayerOf(state)
  if (!layer) return false
  if (state.editTarget === 'mask' && layer.mask) return true
  return !pixelEditBlocker(layer, 'pixels')
}

function canMergeDown(state: DocumentState): boolean {
  const index = layerIndexOf(state, state.activeLayerId)
  if (index <= 0) return false
  const lower = state.layers[index - 1]
  return lower.kind !== 'adjustment' && lower.visible && state.layers[index].visible
}

function hasSessionNow(ctx: CommandContext): boolean {
  return sessionController(ctx) !== null
}

const doc = (run: Run, enabled?: CommandSpec['enabled']): CommandSpec => ({ run, enabled, document: true })
const ui = (run: Run, enabled?: CommandSpec['enabled']): CommandSpec => ({ run, enabled })

const SPECS: Readonly<Record<(typeof STATIC_COMMANDS)[number], CommandSpec>> = {
  'edit.undo': ui((ctx) => {
    const session = sessionController(ctx)
    if (session) {
      // Typed text is never thrown away by Undo: the text field undoes its own typing.
      if (session.id === 'text') return
      session.cancelSession()
      return
    }
    if (ctx.store.history.canUndo()) ctx.store.history.undo()
  }, (ctx) => hasSessionNow(ctx) || ctx.store.history.canUndo()),
  'edit.redo': ui((ctx) => {
    if (ctx.store.history.canRedo()) ctx.store.history.redo()
  }, (ctx) => ctx.store.history.canRedo()),
  'edit.toggle-last': ui((ctx) => ctx.store.history.toggleLast(), (ctx) => ctx.store.history.canUndo() || ctx.store.history.canRedo()),
  'edit.cut': doc((ctx) => cut(ctx), (_ctx, state) => activeIs(state, (layer) => layer.kind === 'raster' && !pixelEditBlocker(layer, 'pixels'))),
  'edit.copy': ui((ctx) => copy(ctx, false), (_ctx, state) => activeIs(state, hasPixels)),
  'edit.copy-merged': ui((ctx) => copy(ctx, true)),
  'edit.paste': doc((ctx) => paste(ctx, false)),
  'edit.paste-in-place': doc((ctx) => paste(ctx, true)),
  'edit.clear': doc((ctx) => clear(ctx), (_ctx, state) => state.selection ? editablePixels(state) : state.layers.length > 1 && Boolean(activeLayerOf(state))),
  'edit.fill-foreground': doc((ctx) => fillWithColor(ctx, ctx.editor.getState().foreground, false, 'Fill'), (_ctx, state) => editablePixels(state)),
  'edit.fill-background': doc((ctx) => fillWithColor(ctx, ctx.editor.getState().background, false, 'Fill'), (_ctx, state) => editablePixels(state)),
  'edit.fill-foreground-preserve': doc((ctx) => fillWithColor(ctx, ctx.editor.getState().foreground, true, 'Fill'), (_ctx, state) => editablePixels(state)),
  'edit.free-transform': doc((ctx) => {
    const start = servicesOf(ctx).startTransform
    if (start) start()
    else fail(ctx, 'Free Transform is not available here.')
  }, (ctx, state) => Boolean(servicesOf(ctx).startTransform) && activeIs(state, (layer) => layer.kind !== 'adjustment')),
  'edit.swap-colors': ui((ctx) => {
    const { foreground, background } = ctx.editor.getState()
    ctx.editor.update({ foreground: background, background: foreground })
  }),
  'edit.default-colors': ui((ctx) => ctx.editor.update({ foreground: DEFAULT_FOREGROUND, background: DEFAULT_BACKGROUND })),

  'select.all': ui((ctx) => {
    const state = ctx.store.getState()
    setSelection(ctx, 'Select All', selectAll(state.width, state.height, nextSelectionVersion(state)))
  }),
  'select.deselect': ui((ctx) => {
    if (ctx.store.getState().selection) setSelection(ctx, 'Deselect', null)
  }, (_ctx, state) => Boolean(state.selection)),
  'select.reselect': ui((ctx) => {
    const state = ctx.store.getState()
    const last = runtimeOf(ctx.store).lastSelection
    if (!last) return
    if (last.width !== state.width || last.height !== state.height) {
      fail(ctx, 'The last selection no longer fits this image size.')
      return
    }
    setSelection(ctx, 'Reselect', restoreSelection(last.snapshot, state.width, state.height, nextSelectionVersion(state)))
  }, (ctx, state) => Boolean(runtimeOf(ctx.store).lastSelection) && !state.selection),
  'select.inverse': ui((ctx) => {
    const state = ctx.store.getState()
    if (!state.selection) return
    setSelection(ctx, 'Inverse', invertSelection(state.selection, nextSelectionVersion(state)))
  }, (_ctx, state) => Boolean(state.selection)),
  'select.feather': ui((ctx) => openDialog(ctx, { kind: 'modify-selection', operation: 'feather' }, 'Feather'),
    (ctx, state) => Boolean(state.selection) && dialogAvailable(ctx, { kind: 'modify-selection', operation: 'feather' })),
  'select.expand': ui((ctx) => openDialog(ctx, { kind: 'modify-selection', operation: 'expand' }, 'Expand'),
    (ctx, state) => Boolean(state.selection) && dialogAvailable(ctx, { kind: 'modify-selection', operation: 'expand' })),
  'select.contract': ui((ctx) => openDialog(ctx, { kind: 'modify-selection', operation: 'contract' }, 'Contract'),
    (ctx, state) => Boolean(state.selection) && dialogAvailable(ctx, { kind: 'modify-selection', operation: 'contract' })),
  'select.load-layer-alpha': ui((ctx) => {
    const state = ctx.store.getState()
    const active = activeLayerOf(state)
    if (!active || active.kind === 'adjustment') {
      fail(ctx, 'Select a layer with pixels to load its transparency.')
      return
    }
    const bounds = layerPixelBounds(active)
    const next = bounds ? selectionFromAlpha(readLayerRect(active, bounds), bounds.x, bounds.y, state.width, state.height, nextSelectionVersion(state)) : null
    if (!next) {
      fail(ctx, `"${active.name}" has no visible pixels on the canvas.`)
      return
    }
    setSelection(ctx, 'Load Selection', next)
  }, (_ctx, state) => activeIs(state, hasPixels)),

  'layer.new': doc((ctx) => newLayer(ctx), (_ctx, state) => state.layers.length < LIMITS.maxLayers),
  'layer.duplicate': doc((ctx) => duplicate(ctx), (_ctx, state) => Boolean(activeLayerOf(state)) && state.layers.length < LIMITS.maxLayers),
  'layer.delete': doc((ctx) => deleteLayer(ctx), (_ctx, state) => Boolean(activeLayerOf(state)) && state.layers.length > 1),
  'layer.via-copy': doc((ctx) => layerVia(ctx, false), (_ctx, state) => state.selection ? activeIs(state, (layer) => layer.kind === 'raster') : Boolean(activeLayerOf(state))),
  'layer.via-cut': doc((ctx) => layerVia(ctx, true), (_ctx, state) => Boolean(state.selection) && activeIs(state, (layer) => layer.kind === 'raster' && !pixelEditBlocker(layer, 'pixels'))),
  'layer.merge-down': doc((ctx) => mergeDown(ctx), (_ctx, state) => canMergeDown(state)),
  'layer.merge-visible': doc((ctx) => mergeVisible(ctx), (_ctx, state) => shownLayers(state).size >= 2),
  'layer.stamp-visible': doc((ctx) => stampVisible(ctx), (_ctx, state) => shownLayers(state).size >= 1 && state.layers.length < LIMITS.maxLayers),
  'layer.flatten': doc((ctx) => flatten(ctx), (_ctx, state) => state.layers.length > 1 || !isBackgroundLayer(state.layers[0])),
  'layer.add-mask': doc((ctx) => addMask(ctx), (_ctx, state) => activeIs(state, (layer) => !layer.mask)),
  'layer.delete-mask': doc((ctx) => {
    const active = activeLayerOf(ctx.store.getState())
    if (active?.mask) ctx.store.transact('Delete Layer Mask', 'layer', (tx) => tx.setMask(active.id, null))
  }, (_ctx, state) => activeIs(state, (layer) => Boolean(layer.mask))),
  'layer.apply-mask': doc((ctx) => applyMask(ctx), (_ctx, state) => activeIs(state, (layer) => Boolean(layer.mask) && layer.kind === 'raster')),
  'layer.toggle-mask': doc((ctx) => {
    const active = activeLayerOf(ctx.store.getState())
    if (!active?.mask) return
    const enabled = !active.mask.enabled
    ctx.store.transact(enabled ? 'Enable Layer Mask' : 'Disable Layer Mask', 'layer', (tx) => tx.updateLayer(active.id, { maskEnabled: enabled }))
  }, (_ctx, state) => activeIs(state, (layer) => Boolean(layer.mask))),
  'layer.rasterize': doc((ctx) => rasterize(ctx), (_ctx, state) => activeIs(state, (layer) => layer.kind === 'text' || layer.kind === 'shape')),
  'layer.from-background': doc((ctx) => {
    const active = activeLayerOf(ctx.store.getState())
    if (active && isBackgroundLayer(active)) ctx.store.transact('Layer from Background', 'layer', (tx) => layerFromBackground(ctx, tx, active))
  }, (_ctx, state) => activeIs(state, isBackgroundLayer)),
  'layer.toggle-clipping': doc((ctx) => {
    const state = ctx.store.getState()
    const active = activeLayerOf(state)
    if (!active) return
    if (isBackgroundLayer(active) || layerIndexOf(state, active.id) === 0) {
      fail(ctx, 'A clipping mask needs a layer below to clip to.')
      return
    }
    const clipped = !active.clipped
    ctx.store.transact(clipped ? 'Create Clipping Mask' : 'Release Clipping Mask', 'layer', (tx) => tx.updateLayer(active.id, { clipped }))
  }, (_ctx, state) => activeIs(state, (layer) => !isBackgroundLayer(layer) && layerIndexOf(state, layer.id) > 0)),
  'layer.raise': doc((ctx) => moveActive(ctx, 'raise'), (_ctx, state) => activeIs(state, (layer) => !isBackgroundLayer(layer) && layerIndexOf(state, layer.id) < state.layers.length - 1)),
  'layer.lower': doc((ctx) => moveActive(ctx, 'lower'), (_ctx, state) => activeIs(state, (layer) => !isBackgroundLayer(layer) && layerIndexOf(state, layer.id) > (isBackgroundLayer(state.layers[0]) ? 1 : 0))),
  'layer.to-front': doc((ctx) => moveActive(ctx, 'front'), (_ctx, state) => activeIs(state, (layer) => !isBackgroundLayer(layer) && layerIndexOf(state, layer.id) < state.layers.length - 1)),
  'layer.to-back': doc((ctx) => moveActive(ctx, 'back'), (_ctx, state) => activeIs(state, (layer) => !isBackgroundLayer(layer) && layerIndexOf(state, layer.id) > (isBackgroundLayer(state.layers[0]) ? 1 : 0))),
  'layer.select-above': ui((ctx) => selectNeighbour(ctx, 1), (_ctx, state) => layerIndexOf(state, state.activeLayerId) < state.layers.length - 1),
  'layer.select-below': ui((ctx) => selectNeighbour(ctx, -1), (_ctx, state) => layerIndexOf(state, state.activeLayerId) !== 0 && state.layers.length > 0),

  'image.size': doc((ctx) => openDialog(ctx, { kind: 'image-size' }, 'Image Size'), (ctx) => dialogAvailable(ctx, { kind: 'image-size' })),
  'image.canvas-size': doc((ctx) => openDialog(ctx, { kind: 'canvas-size' }, 'Canvas Size'), (ctx) => dialogAvailable(ctx, { kind: 'canvas-size' })),
  'image.crop-to-selection': doc((ctx) => {
    const selection = ctx.store.getState().selection
    if (!selection) {
      fail(ctx, 'Make a selection to crop to.')
      return
    }
    cropCanvasTo(ctx, selection.bounds, 'Crop')
  }, (_ctx, state) => Boolean(state.selection)),
  'image.trim': doc((ctx) => trim(ctx), (_ctx, state) => !(isBackgroundLayer(state.layers[0]) && state.layers[0].visible)),
  'image.rotate-cw': doc((ctx) => turnCanvas(ctx, 'cw')),
  'image.rotate-ccw': doc((ctx) => turnCanvas(ctx, 'ccw')),
  'image.rotate-180': doc((ctx) => turnCanvas(ctx, '180')),
  'image.rotate-arbitrary': doc((ctx) => openDialog(ctx, { kind: 'rotate-canvas' }, 'Rotate Canvas'), (ctx) => dialogAvailable(ctx, { kind: 'rotate-canvas' })),
  'image.flip-horizontal': doc((ctx) => turnCanvas(ctx, 'flip-h')),
  'image.flip-vertical': doc((ctx) => turnCanvas(ctx, 'flip-v')),
  'image.auto-tone': doc((ctx) => autoAdjust(ctx, 'tone'), (_ctx, state) => editablePixels(state)),
  'image.auto-contrast': doc((ctx) => autoAdjust(ctx, 'contrast'), (_ctx, state) => editablePixels(state)),
  'image.auto-color': doc((ctx) => autoAdjust(ctx, 'color'), (_ctx, state) => editablePixels(state)),
  'image.desaturate': doc((ctx) => applyAdjustmentToLayer(ctx, { ...defaultAdjustment('hue-saturation'), master: { hue: 0, saturation: -100, lightness: 0 } }, 'Desaturate'),
    (_ctx, state) => editablePixels(state)),

  'filter.repeat': doc((ctx) => {
    const last = ctx.editor.getState().lastFilter
    if (!last) {
      fail(ctx, 'No filter has been used yet.')
      return
    }
    return applyFilterToLayer(ctx, last)
  }, (ctx, state) => Boolean(ctx.editor.getState().lastFilter) && editablePixels(state)),
  'filter.sharpen-more': doc((ctx) => applyFilterToLayer(ctx, { type: 'sharpen', strength: 'more' }), (_ctx, state) => editablePixels(state)),

  'view.zoom-in': ui((ctx) => zoomStep(ctx, 1)),
  'view.zoom-out': ui((ctx) => zoomStep(ctx, -1)),
  'view.fit': ui((ctx) => ctx.view.fit()),
  'view.actual-pixels': ui((ctx) => ctx.view.actualPixels()),
  'view.toggle-panels': ui((ctx) => togglePanels(ctx)),
  'view.panel-layers': ui((ctx) => togglePanel(ctx, 'layers')),
  'view.panel-properties': ui((ctx) => togglePanel(ctx, 'properties')),
  'view.panel-history': ui((ctx) => togglePanel(ctx, 'history')),
  'view.panel-color': ui((ctx) => togglePanel(ctx, 'color')),

  'brush.smaller': ui((ctx) => stepBrush(ctx, 'size', -1), (ctx) => sizeOptionKey(effectiveTool(ctx)) !== null),
  'brush.larger': ui((ctx) => stepBrush(ctx, 'size', 1), (ctx) => sizeOptionKey(effectiveTool(ctx)) !== null),
  'brush.softer': ui((ctx) => stepBrush(ctx, 'hardness', -1), (ctx) => sizeOptionKey(effectiveTool(ctx)) !== null),
  'brush.harder': ui((ctx) => stepBrush(ctx, 'hardness', 1), (ctx) => sizeOptionKey(effectiveTool(ctx)) !== null),

  'tool.cycle-marquee': ui((ctx) => cycleSlot(ctx, 0)),
  'tool.cycle-lasso': ui((ctx) => cycleSlot(ctx, 1)),
  'tool.cycle-gradient': ui((ctx) => cycleSlot(ctx, 2)),
  'tool.cycle-shape': ui((ctx) => cycleShape(ctx)),

  'session.commit': ui((ctx) => sessionController(ctx)?.commitSession(), (ctx) => hasSessionNow(ctx)),
  'session.cancel': ui((ctx) => sessionController(ctx)?.cancelSession(), (ctx) => hasSessionNow(ctx)),
}

const IMMEDIATE_FILTERS: ReadonlySet<FilterType> = new Set<FilterType>(['sharpen', 'find-edges'])

function specFor(id: EditorCommandId): CommandSpec | null {
  const fixed = (SPECS as Record<string, CommandSpec>)[id]
  if (fixed) return fixed
  const dot = id.indexOf('.')
  const group = id.slice(0, dot)
  const name = id.slice(dot + 1)
  if (group === 'tool') {
    if (isToolId(name)) return ui((ctx) => selectTool(ctx, name))
    const digit = /^(opacity|flow)-([0-9])$/.exec(name)
    if (digit) return ui((ctx) => digitCommand(ctx, digit[1] as 'opacity' | 'flow', Number(digit[2])))
    return null
  }
  if (group === 'adjust' && (ADJUSTMENT_TYPES as readonly string[]).includes(name)) {
    const type = name as AdjustmentType
    if (type === 'invert') return doc((ctx) => applyAdjustmentToLayer(ctx, { type: 'invert' }), (_ctx, state) => editablePixels(state))
    return doc((ctx) => openDialog(ctx, { kind: 'adjustment', type }, adjustmentLabel(type)),
      (ctx, state) => editablePixels(state) && dialogAvailable(ctx, { kind: 'adjustment', type }))
  }
  if (group === 'adjustment-layer' && (ADJUSTMENT_TYPES as readonly string[]).includes(name)) {
    const type = name as AdjustmentType
    return doc((ctx) => newAdjustmentLayer(ctx, type), (_ctx, state) => state.layers.length < LIMITS.maxLayers)
  }
  if (group === 'filter' && (FILTER_TYPES as readonly string[]).includes(name)) {
    const type = name as FilterType
    if (IMMEDIATE_FILTERS.has(type)) return doc((ctx) => applyFilterToLayer(ctx, defaultFilter(type)), (_ctx, state) => editablePixels(state))
    return doc((ctx) => openDialog(ctx, { kind: 'filter', type }, filterLabel(type)),
      (ctx, state) => editablePixels(state) && dialogAvailable(ctx, { kind: 'filter', type }))
  }
  return null
}

function blocked(ctx: CommandContext): boolean {
  return commandsBusy(ctx.store) || Boolean(servicesOf(ctx).busy?.())
}

/** Whether a command can run right now (menus grey out the rest). Never throws. */
export function isCommandEnabled(id: EditorCommandId, context: CommandContext): boolean {
  const spec = specFor(id)
  if (!spec) return false
  try {
    if (spec.document && blocked(context)) return false
    return spec.enabled ? spec.enabled(context, context.store.getState()) : true
  } catch {
    return false
  }
}

/**
 * Runs a command and resolves when it (and any worker job it started) has finished. Problems are shown to
 * the user through host.notify; the promise itself never rejects.
 */
export async function runCommandAsync(id: EditorCommandId, context: CommandContext): Promise<void> {
  const spec = specFor(id)
  if (!spec) {
    context.host.notify(`"${id}" is not a command of this editor.`, 'error')
    return
  }
  if (spec.document && blocked(context)) {
    context.host.notify('Wait a moment: the last change is still being applied.')
    return
  }
  try {
    await spec.run(context)
  } catch (error) {
    report(context, error, 'That could not be done.')
  }
}

/** Fire-and-forget form used by menus and shortcuts (design 8.4 signature). */
export function runCommand(id: EditorCommandId, context: CommandContext): void {
  void runCommandAsync(id, context)
}

// ---------------------------------------------------------------------------------------------
// Labels (menus, tooltips, history)
// ---------------------------------------------------------------------------------------------

const LABELS: Readonly<Partial<Record<EditorCommandId, string>>> = Object.freeze({
  'edit.undo': 'Undo',
  'edit.redo': 'Redo',
  'edit.toggle-last': 'Toggle Last State',
  'edit.cut': 'Cut',
  'edit.copy': 'Copy',
  'edit.copy-merged': 'Copy Merged',
  'edit.paste': 'Paste',
  'edit.paste-in-place': 'Paste in Place',
  'edit.clear': 'Clear',
  'edit.fill-foreground': 'Fill with Foreground Color',
  'edit.fill-background': 'Fill with Background Color',
  'edit.fill-foreground-preserve': 'Fill Foreground, Preserve Transparency',
  'edit.free-transform': 'Free Transform',
  'edit.swap-colors': 'Swap Colors',
  'edit.default-colors': 'Default Colors',
  'select.all': 'All',
  'select.deselect': 'Deselect',
  'select.reselect': 'Reselect',
  'select.inverse': 'Inverse',
  'select.feather': 'Feather…',
  'select.expand': 'Expand…',
  'select.contract': 'Contract…',
  'select.load-layer-alpha': 'Load Selection from Layer',
  'layer.new': 'New Layer',
  'layer.duplicate': 'Duplicate Layer',
  'layer.delete': 'Delete Layer',
  'layer.via-copy': 'Layer via Copy',
  'layer.via-cut': 'Layer via Cut',
  'layer.merge-down': 'Merge Down',
  'layer.merge-visible': 'Merge Visible',
  'layer.stamp-visible': 'Stamp Visible',
  'layer.flatten': 'Flatten Image',
  'layer.add-mask': 'Add Layer Mask',
  'layer.delete-mask': 'Delete Layer Mask',
  'layer.apply-mask': 'Apply Layer Mask',
  'layer.toggle-mask': 'Disable Layer Mask',
  'layer.rasterize': 'Rasterize Layer',
  'layer.from-background': 'Layer from Background',
  'layer.toggle-clipping': 'Create Clipping Mask',
  'layer.raise': 'Bring Forward',
  'layer.lower': 'Send Backward',
  'layer.to-front': 'Bring to Front',
  'layer.to-back': 'Send to Back',
  'layer.select-above': 'Select Layer Above',
  'layer.select-below': 'Select Layer Below',
  'image.size': 'Image Size…',
  'image.canvas-size': 'Canvas Size…',
  'image.crop-to-selection': 'Crop to Selection',
  'image.trim': 'Trim Transparent Pixels',
  'image.rotate-cw': '90° Clockwise',
  'image.rotate-ccw': '90° Counter Clockwise',
  'image.rotate-180': '180°',
  'image.rotate-arbitrary': 'Arbitrary…',
  'image.flip-horizontal': 'Flip Canvas Horizontal',
  'image.flip-vertical': 'Flip Canvas Vertical',
  'image.auto-tone': 'Auto Tone',
  'image.auto-contrast': 'Auto Contrast',
  'image.auto-color': 'Auto Color',
  'image.desaturate': 'Desaturate',
  'filter.repeat': 'Last Filter',
  'filter.sharpen-more': 'Sharpen More',
  'view.zoom-in': 'Zoom In',
  'view.zoom-out': 'Zoom Out',
  'view.fit': 'Fit on Screen',
  'view.actual-pixels': '100%',
  'view.toggle-panels': 'Show / Hide Panels',
  'view.panel-layers': 'Layers',
  'view.panel-properties': 'Properties',
  'view.panel-history': 'History',
  'view.panel-color': 'Color',
  'brush.smaller': 'Decrease Brush Size',
  'brush.larger': 'Increase Brush Size',
  'brush.softer': 'Decrease Brush Hardness',
  'brush.harder': 'Increase Brush Hardness',
  'tool.cycle-marquee': 'Next Marquee Tool',
  'tool.cycle-lasso': 'Next Lasso Tool',
  'tool.cycle-gradient': 'Gradient / Paint Bucket',
  'tool.cycle-shape': 'Next Shape',
  'session.commit': 'Commit',
  'session.cancel': 'Cancel',
})

/** Menu label of a command; some labels follow the document ("Release Clipping Mask", "Last Filter (Gaussian Blur)"). */
export function commandLabel(id: EditorCommandId, context?: CommandContext): string {
  if (context) {
    const state = context.store.getState()
    const active = activeLayerOf(state)
    if (id === 'layer.toggle-clipping' && active?.clipped) return 'Release Clipping Mask'
    if (id === 'layer.toggle-mask' && active?.mask && !active.mask.enabled) return 'Enable Layer Mask'
    if (id === 'filter.repeat') {
      const last = context.editor.getState().lastFilter
      if (last) return `Last Filter (${filterLabel(last.type, last)})`
    }
  }
  const fixed = LABELS[id]
  if (fixed) return fixed
  const dot = id.indexOf('.')
  const group = id.slice(0, dot)
  const name = id.slice(dot + 1)
  if (group === 'tool') {
    if (isToolId(name)) return TOOL_META[name].label
    const digit = /^(opacity|flow)-([0-9])$/.exec(name)
    if (digit) return `${digit[1] === 'opacity' ? 'Opacity' : 'Flow'} ${digit[2] === '0' ? 100 : Number(digit[2]) * 10}%`
  }
  if (group === 'adjust') return name === 'invert' ? 'Invert' : `${adjustmentLabel(name as AdjustmentType)}…`
  if (group === 'adjustment-layer') return `${adjustmentLabel(name as AdjustmentType)}…`
  if (group === 'filter') return IMMEDIATE_FILTERS.has(name as FilterType) ? filterLabel(name as FilterType) : `${filterLabel(name as FilterType)}…`
  return id
}
