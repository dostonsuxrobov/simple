// src/advanced/types.ts
// The single shared contract for the Advanced editor. Owned by WP1; later changes need agreement of
// every package that imports it. Runtime content is limited to the frozen constants at the bottom.
// Pure engine modules (tiles, document, history, composite, pyramid, selection, psdMapping) follow the
// same erasable-TypeScript rules as src/imaging so Node tests can import them directly.

import type {
  AdjustmentSpec,
  AdjustmentType,
  Affine,
  BlendMode,
  FilterSpec,
  FilterType,
  GradientKind,
  ImagingClient,
  IntRect,
  MaskBuffer,
  OpOptions,
  PixelBuffer,
  Point,
  Rect,
  Rgb8,
  Rgba8,
  SelectionOp,
  Size,
} from '../imaging/types.ts'

// ---------------------------------------------------------------------------------------------
// Sparse tiled storage (src/advanced/tiles.ts)
// ---------------------------------------------------------------------------------------------

/** Tiles are TILE_SIZE x TILE_SIZE, addressed by integer tile coordinates that may be negative. */
export interface TiledSurface {
  readonly channels: 4
  /** Monotonic; bumps on every mutation (used by caches). */
  readonly version: number
  /** Allocated tiles x TILE_SIZE^2 x 4. */
  readonly byteSize: number
  readonly tileCount: number
  /** Tight bounds of pixels with alpha > 0 in layer-local coordinates, or null. Cached per version. */
  contentBounds(): IntRect | null
  /** Copy a region out; unallocated areas read as transparent black. */
  read(rect: IntRect, target?: PixelBuffer): PixelBuffer
  /** Copy a region in at (x, y); allocates tiles; frees tiles that become fully transparent. Returns touched tile-aligned rect. */
  write(x: number, y: number, src: PixelBuffer, srcRect?: IntRect): IntRect
  tile(tx: number, ty: number): Uint8ClampedArray | undefined
  /** Replaces a whole tile (length TILE_SIZE^2 * 4) or deletes it with undefined. */
  setTile(tx: number, ty: number, data: Uint8ClampedArray | undefined): void
  tileVersion(tx: number, ty: number): number
  forEachTile(visit: (tx: number, ty: number, data: Uint8ClampedArray) => void): void
  clone(): TiledSurface
}

export interface TiledMask {
  readonly channels: 1
  /** Value of every unallocated pixel: 255 = reveal all (new layer mask), 0 = hide all. */
  readonly defaultValue: 0 | 255
  readonly version: number
  readonly byteSize: number
  readonly tileCount: number
  /** Tight bounds of pixels that differ from defaultValue, or null. */
  contentBounds(): IntRect | null
  read(rect: IntRect, target?: MaskBuffer): MaskBuffer
  write(x: number, y: number, src: MaskBuffer, srcRect?: IntRect): IntRect
  tile(tx: number, ty: number): Uint8Array | undefined
  setTile(tx: number, ty: number, data: Uint8Array | undefined): void
  tileVersion(tx: number, ty: number): number
  forEachTile(visit: (tx: number, ty: number, data: Uint8Array) => void): void
  clone(): TiledMask
}

// ---------------------------------------------------------------------------------------------
// Document model (layers are ordered bottom -> top, index 0 = bottom; same order as ag-psd)
// ---------------------------------------------------------------------------------------------

export type LayerId = string

export interface LayerLocks {
  /** Lock image pixels: painting, filters and adjustments are refused. */
  readonly pixels: boolean
  /** Lock position: move and transform are refused. */
  readonly position: boolean
  /** Lock transparent pixels: painting preserves alpha; eraser paints the background colour. */
  readonly transparency: boolean
}

export interface LayerMask {
  readonly surface: TiledMask
  /** Document position of mask-local (0, 0). */
  readonly offsetX: number
  readonly offsetY: number
  readonly enabled: boolean
  /** When true, moving/transforming the layer moves the mask too (Photoshop chain icon). */
  readonly linked: boolean
}

export interface LayerCommon {
  readonly id: LayerId
  readonly name: string
  readonly visible: boolean
  /** 0..1; exported to PSD as round(opacity * 255) / 255. */
  readonly opacity: number
  readonly blendMode: BlendMode
  readonly locks: LayerLocks
  readonly mask: LayerMask | null
  /** Clipping mask onto the nearest non-clipped layer below. v1 renders it; UI exposes Alt+click / Ctrl+Alt+G. */
  readonly clipped: boolean
}

export interface RasterLayer extends LayerCommon {
  readonly kind: 'raster'
  /** Photoshop "Background": always bottom, opaque, locks.position and locks.transparency set. */
  readonly isBackground: boolean
  /** Document position of layer-local (0, 0). Moving a layer only changes these. */
  readonly offsetX: number
  readonly offsetY: number
  /** Mutable pixel content with a stable identity. Mutate only through Transaction/StrokeSession. */
  readonly surface: TiledSurface
}

export interface TextStyle {
  /** CSS family name, e.g. 'Segoe UI'. PSD import maps PostScript names via queryLocalFonts() when available. */
  readonly fontFamily: string
  /** px at document resolution (pt == px when ppi is 72). */
  readonly fontSize: number
  readonly fontWeight: 400 | 700
  readonly italic: boolean
  readonly underline: boolean
  readonly color: Rgb8
  readonly align: 'left' | 'center' | 'right'
  /** Multiple of fontSize, e.g. 1.2. */
  readonly lineHeight: number
  /** px; canvas letterSpacing. */
  readonly letterSpacing: number
}

export interface TextSpec {
  readonly text: string
  readonly style: TextStyle
  /** null = point text (no wrapping); number = paragraph box width in px (wraps). */
  readonly boxWidth: number | null
  /** Maps text-local space (origin = top-left of the first line box) to document space. */
  readonly transform: Affine
}

export type ShapeKind = 'rectangle' | 'ellipse' | 'line' | 'arrow'

export interface ShapeSpec {
  readonly kind: ShapeKind
  /** Shape-local box (rectangle/ellipse) or segment endpoints (line/arrow): (x1, y1) -> (x2, y2). */
  readonly x1: number
  readonly y1: number
  readonly x2: number
  readonly y2: number
  readonly fill: Rgba8 | null
  readonly stroke: Rgba8 | null
  readonly strokeWidth: number
  readonly cornerRadius: number
  readonly arrowHeads: 'none' | 'end' | 'both'
  /** Shape-local to document space. */
  readonly transform: Affine
}

/** Derived pixels for text/shape layers; regenerated whenever specKey differs from the spec's key. */
export interface RasterCache {
  readonly surface: TiledSurface
  readonly offsetX: number
  readonly offsetY: number
  readonly specKey: string
}

export interface TextLayer extends LayerCommon {
  readonly kind: 'text'
  readonly text: TextSpec
  readonly raster: RasterCache
}

export interface ShapeLayer extends LayerCommon {
  readonly kind: 'shape'
  readonly shape: ShapeSpec
  readonly raster: RasterCache
}

export interface AdjustmentLayer extends LayerCommon {
  readonly kind: 'adjustment'
  readonly adjustment: AdjustmentSpec
}

export type Layer = RasterLayer | TextLayer | ShapeLayer | AdjustmentLayer
export type LayerKind = Layer['kind']

/** Dense, document-sized selection. bounds = tight bounds of non-zero coverage. */
export interface Selection {
  readonly mask: MaskBuffer
  readonly bounds: IntRect
  readonly version: number
}

/** Compact selection snapshot used by history and Reselect (Ctrl+Shift+D). */
export type SelectionSnapshot =
  | { readonly kind: 'none' }
  | { readonly kind: 'all' }
  | { readonly kind: 'region'; readonly rect: IntRect; readonly data: Uint8Array }

/** Immutable snapshot consumed by React through useSyncExternalStore. */
export interface DocumentState {
  readonly width: number
  readonly height: number
  /** Pixels per inch (PSD resolutionInfo, JPEG JFIF density); 72 when unknown. */
  readonly ppi: number
  readonly layers: readonly Layer[]
  readonly activeLayerId: LayerId | null
  /** Photoshop "target the mask" state for the active layer. */
  readonly editTarget: EditTarget
  readonly selection: Selection | null
  /** Host content revision this state corresponds to (drives the Modified badge). */
  readonly revision: number
  /** Bumps on any pixel change; thumbnails and caches key on it. */
  readonly pixelVersion: number
}

export type EditTarget = 'pixels' | 'mask'

export type LayerPatch = {
  readonly name?: string
  readonly visible?: boolean
  readonly opacity?: number
  readonly blendMode?: BlendMode
  readonly locks?: LayerLocks
  readonly clipped?: boolean
  readonly offset?: Point
  readonly maskEnabled?: boolean
  readonly maskLinked?: boolean
  readonly isBackground?: boolean
  readonly text?: TextSpec
  readonly shape?: ShapeSpec
  readonly adjustment?: AdjustmentSpec
}

export interface PixelEditor {
  readonly layerId: LayerId
  readonly target: EditTarget
  /** Live surface; layer-local coordinates. */
  readonly surface: TiledSurface | TiledMask
  /** Call with the layer-local rect BEFORE writing inside it: snapshots untouched tiles for undo. */
  touch(rect: IntRect): void
  /** touch + surface.write + invalidate the composite. */
  writePixels(x: number, y: number, src: PixelBuffer): void
  writeMask(x: number, y: number, src: MaskBuffer): void
  /** Pre-edit content of a region (stroke opacity model, clone source); falls back to live pixels where untouched. */
  readBefore(rect: IntRect): PixelBuffer | MaskBuffer
  /** Marks a document-space rect dirty for compositing without a write (e.g. after setTile). */
  invalidate(rect: IntRect): void
}

export interface Transaction {
  /** Live state inside the transaction (already reflects earlier calls in this transaction). */
  readonly state: DocumentState
  insertLayer(layer: Layer, index?: number): void
  removeLayer(id: LayerId): void
  moveLayer(id: LayerId, toIndex: number): void
  updateLayer(id: LayerId, patch: LayerPatch): void
  setMask(id: LayerId, mask: LayerMask | null): void
  setActiveLayer(id: LayerId | null, target?: EditTarget): void
  setSelection(selection: Selection | null): void
  editPixels(id: LayerId, target: EditTarget): PixelEditor
  /** Whole-buffer swap (transform, image size, rotate canvas). History keeps the old surface alive. */
  replaceSurface(id: LayerId, target: EditTarget, next: TiledSurface | TiledMask, offset?: Point): void
  setCanvas(width: number, height: number, ppi?: number): void
  /** Abort the transaction: every change made so far is reverted and nothing is recorded. */
  rollback(): never
}

export interface TransactOptions {
  /** Consecutive transactions with the same key within COALESCE_MS merge into one history entry (slider drags, nudges). */
  readonly coalesceKey?: string
  /** false for selection-only / active-layer changes: no new content revision, so no Modified badge. Default true. */
  readonly affectsOutput?: boolean
}

export interface StrokeSession {
  readonly editor: PixelEditor
  /** Records one history entry (if anything changed) and allocates a revision. */
  commit(): void
  /** Restores every touched tile; records nothing. */
  cancel(): void
}

export interface DocumentChange {
  readonly structure: boolean
  readonly selection: boolean
  readonly history: boolean
  /** Document-space rects whose composite is stale. */
  readonly dirty: readonly IntRect[] | 'all'
}

export interface DocumentStore {
  getState(): DocumentState
  subscribe(listener: (change: DocumentChange) => void): () => void
  transact<T>(label: string, icon: HistoryIcon, run: (tx: Transaction) => T, options?: TransactOptions): T
  beginStroke(layerId: LayerId, target: EditTarget, label: string, icon: HistoryIcon): StrokeSession
  readonly history: History
  /** Estimated bytes held by layer surfaces, masks, caches and history. */
  memoryUsage(): MemoryUsage
  dispose(): void
}

// ---------------------------------------------------------------------------------------------
// History (src/advanced/history.ts)
// ---------------------------------------------------------------------------------------------

export type HistoryIcon =
  | ToolId
  | 'open' | 'layer' | 'adjustment' | 'filter' | 'image' | 'selection' | 'paste' | 'text' | 'transform' | 'merge' | 'fill'

export interface HistoryEntryInfo {
  readonly id: number
  readonly label: string
  readonly icon: HistoryIcon
  readonly bytes: number
}

export interface HistoryState {
  /** entries[0] is the base state ("Open" / "Advanced editor"), never evicted. */
  readonly entries: readonly HistoryEntryInfo[]
  /** Index of the current state in entries. */
  readonly cursor: number
  readonly totalBytes: number
  readonly budgetBytes: number
  readonly maxEntries: number
  /** True once old entries have been evicted to respect the budget. */
  readonly trimmed: boolean
}

export interface History {
  getState(): HistoryState
  canUndo(): boolean
  canRedo(): boolean
  undo(): void
  redo(): void
  /** History panel click; equivalent to repeated undo/redo. */
  jumpTo(index: number): void
  /** Photoshop Ctrl+Alt+Z: toggles between the current and the previous state. */
  toggleLast(): void
  /** Drops everything and records a new base entry with the given label. */
  reset(label: string): void
  setBudget(bytes: number, maxEntries: number): void
}

export interface MemoryUsage {
  readonly layers: number
  readonly masks: number
  readonly caches: number
  readonly history: number
  readonly total: number
  readonly budget: number
}

// ---------------------------------------------------------------------------------------------
// Compositing (pure: src/advanced/composite.ts + pyramid.ts; DOM: src/advanced/compositor.ts)
// ---------------------------------------------------------------------------------------------

export type SampleSource = 'current' | 'current-below' | 'all'

/**
 * Live, uncommitted previews. 'adjustment' on an adjustment layer replaces its spec; on a raster layer it
 * applies the spec to that layer's pixels inside `selection` (destructive-dialog preview). Free transform
 * previews hide the layer with 'layer-props' and draw a warped snapshot on the overlay instead.
 */
export type CompositePreview =
  | { readonly kind: 'adjustment'; readonly layerId: LayerId; readonly spec: AdjustmentSpec; readonly selection: Selection | null }
  | { readonly kind: 'layer-props'; readonly layerId: LayerId; readonly opacity?: number; readonly blendMode?: BlendMode; readonly visible?: boolean }
  | { readonly kind: 'layer-pixels'; readonly layerId: LayerId; readonly level: number; readonly rect: IntRect; readonly pixels: PixelBuffer }

export interface CompositeOptions {
  /** Pyramid level: 0 = full resolution, n = 1 / 2^n. Default 0. rect is in level-n pixels. */
  readonly level?: number
  readonly preview?: CompositePreview | null
  /** Composite only layers strictly below this layer (Sample "current & below" uses the inclusive variant). */
  readonly belowLayerId?: LayerId
  readonly includeBelowLayer?: boolean
  /** Composite a single layer in isolation (thumbnails, Ctrl+click alpha). */
  readonly onlyLayerId?: LayerId
}

/** Pure, synchronous, deterministic. Used by the display compositor, flatten/export and Node tests. */
export type CompositeRectFn = (
  doc: Pick<DocumentState, 'width' | 'height' | 'layers'>,
  rect: IntRect,
  options?: CompositeOptions,
) => PixelBuffer

/** Screen (CSS px) = document px * zoom + offset. */
export interface ViewTransform {
  readonly zoom: number
  readonly offsetX: number
  readonly offsetY: number
}

export interface ViewportSize {
  readonly width: number
  readonly height: number
  readonly dpr: number
}

export interface Compositor {
  attach(canvas: HTMLCanvasElement): void
  setView(view: ViewTransform, size: ViewportSize): void
  invalidate(dirty: readonly IntRect[] | 'all'): void
  setPreview(preview: CompositePreview | null): void
  /** Resolves once every visible tile at the active level is current. */
  settle(): Promise<void>
  /** Full-resolution composite of the whole document, computed in stripes with progress. */
  flatten(options?: OpOptions): Promise<PixelBuffer>
  /** Same as flatten but streamed into a new canvas; the caller releases it (width = height = 1). */
  renderToCanvas(options?: OpOptions): Promise<HTMLCanvasElement>
  /** Eyedropper/Info sampling at document coordinates, averaging size x size pixels. */
  sample(x: number, y: number, size: number, source: SampleSource): Rgba8
  readonly level: number
  dispose(): void
}

// ---------------------------------------------------------------------------------------------
// Tools (src/advanced/tools/*)
// ---------------------------------------------------------------------------------------------

export type ToolId =
  | 'move'
  | 'marquee-rect' | 'marquee-ellipse'
  | 'lasso' | 'lasso-polygon'
  | 'magic-wand'
  | 'crop'
  | 'eyedropper'
  | 'spot-healing'
  | 'brush'
  | 'clone-stamp'
  | 'eraser'
  | 'gradient' | 'paint-bucket'
  | 'text'
  | 'shape'
  | 'hand'
  | 'zoom'

export interface ToolPointerEvent {
  /** Document coordinates (float). */
  readonly doc: Point
  /** CSS px relative to the view canvas. */
  readonly screen: Point
  /** 0..1; 1 for mouse buttons, 0 for hover. */
  readonly pressure: number
  readonly button: number
  readonly buttons: number
  readonly shift: boolean
  readonly alt: boolean
  readonly ctrl: boolean
  readonly pointerType: 'mouse' | 'pen' | 'touch'
  readonly time: number
  /** Coalesced samples since the previous event, oldest first, in document coordinates (getCoalescedEvents). */
  readonly coalesced: readonly { readonly doc: Point; readonly pressure: number; readonly time: number }[]
}

export interface ViewController {
  getView(): ViewTransform
  getViewportSize(): ViewportSize
  docToScreen(point: Point): Point
  screenToDoc(point: Point): Point
  /** Zoom keeping the document point under `anchor` (CSS px) fixed. */
  zoomAt(zoom: number, anchor: Point): void
  panBy(dx: number, dy: number): void
  fit(): void
  actualPixels(): void
  /** Overlay redraw on the next animation frame. */
  requestOverlay(): void
}

export interface ToolContext {
  readonly store: DocumentStore
  readonly editor: EditorStore
  readonly view: ViewController
  readonly compositor: Compositor
  readonly imaging: ImagingClient
  readonly host: AdvancedHost
  setCursor(cursor: string): void
  /** Status-bar hint, e.g. "Alt-click to set the clone source". */
  setHint(text: string | null): void
  /** Mount a DOM overlay (text editing textarea, transform numeric fields) above the canvas. */
  setOverlayElement(element: HTMLElement | null): void
}

export interface ToolController {
  readonly id: ToolId
  activate(ctx: ToolContext): void
  /** Must commit or cancel any session (text: commit; transform/crop: commit only when asked, otherwise cancel). */
  deactivate(): void
  pointerDown(event: ToolPointerEvent): void
  /** Also called for hover (buttons === 0) so tools can update cursors and outlines. */
  pointerMove(event: ToolPointerEvent): void
  pointerUp(event: ToolPointerEvent): void
  pointerCancel(): void
  /** Return true when handled (Enter / Escape / arrows for crop, transform, polygon lasso, text). */
  keyDown(event: KeyboardEvent): boolean
  keyUp(event: KeyboardEvent): boolean
  /** Screen-space overlay (marching ants are drawn by CanvasView, not tools). */
  drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void
  /** True while a crop, free transform, polygon lasso or text edit is open. */
  hasSession(): boolean
  commitSession(): void
  cancelSession(): void
}

export type ToolFactory = () => ToolController

// ---------------------------------------------------------------------------------------------
// Editor state (tool options, colours, view, panels) - src/advanced/editorState.ts
// ---------------------------------------------------------------------------------------------

export interface BrushOptions {
  readonly size: number
  readonly hardness: number
  readonly opacity: number
  readonly flow: number
  /** Fraction of the diameter between dabs; Photoshop default 0.25. */
  readonly spacing: number
  /** 0..1 input smoothing. */
  readonly smoothing: number
  readonly pressureSize: boolean
  readonly pressureOpacity: boolean
  readonly blendMode: BlendMode
}

export type AspectPreset = 'free' | 'original' | '1:1' | '4:3' | '3:2' | '16:9' | '5:4' | '7:5'

export interface ToolOptions {
  readonly brush: BrushOptions
  readonly eraser: BrushOptions
  readonly clone: BrushOptions & { readonly aligned: boolean; readonly sample: SampleSource }
  readonly heal: { readonly size: number; readonly hardness: number; readonly sample: SampleSource }
  readonly marquee: { readonly op: SelectionOp; readonly feather: number; readonly antiAlias: boolean; readonly style: 'normal' | 'fixed-ratio' | 'fixed-size'; readonly ratio: Size; readonly fixedSize: Size }
  readonly lasso: { readonly op: SelectionOp; readonly feather: number; readonly antiAlias: boolean }
  readonly wand: { readonly op: SelectionOp; readonly tolerance: number; readonly contiguous: boolean; readonly antiAlias: boolean; readonly sample: SampleSource }
  readonly bucket: { readonly tolerance: number; readonly contiguous: boolean; readonly antiAlias: boolean; readonly sample: SampleSource; readonly opacity: number; readonly blendMode: BlendMode }
  readonly gradient: { readonly kind: GradientKind; readonly preset: 'foreground-background' | 'foreground-transparent' | 'black-white' | 'custom'; readonly reverse: boolean; readonly dither: boolean; readonly opacity: number; readonly blendMode: BlendMode }
  readonly crop: { readonly aspect: AspectPreset; readonly portrait: boolean; readonly deleteCroppedPixels: boolean; readonly overlay: 'thirds' | 'grid' | 'none' }
  readonly text: TextStyle
  readonly shape: { readonly kind: ShapeKind; readonly fill: Rgba8 | null; readonly stroke: Rgba8 | null; readonly strokeWidth: number; readonly cornerRadius: number }
  readonly eyedropper: { readonly size: 1 | 3 | 5 | 11; readonly sample: SampleSource }
  readonly move: { readonly autoSelect: boolean; readonly showTransformControls: boolean }
}

export interface PanelVisibility {
  readonly layers: boolean
  readonly properties: boolean
  readonly history: boolean
  readonly color: boolean
}

export interface EditorState {
  readonly tool: ToolId
  /** Tool to return to after a spring-loaded key (Space = hand, Alt in paint tools = eyedropper) is released. */
  readonly springFrom: ToolId | null
  readonly foreground: Rgb8
  readonly background: Rgb8
  readonly options: ToolOptions
  readonly swatches: readonly Rgb8[]
  readonly recentColors: readonly Rgb8[]
  readonly view: ViewTransform
  readonly panels: PanelVisibility
  readonly lastFilter: FilterSpec | null
}

export interface EditorStore {
  getState(): EditorState
  subscribe(listener: () => void): () => void
  update(patch: Partial<Omit<EditorState, 'options'>>): void
  updateOptions<K extends keyof ToolOptions>(tool: K, patch: Partial<ToolOptions[K]>): void
}

// ---------------------------------------------------------------------------------------------
// Commands and shortcuts (src/advanced/commands.ts, shortcuts.ts)
// ---------------------------------------------------------------------------------------------

export type CommandId =
  | 'edit.undo' | 'edit.redo' | 'edit.toggle-last'
  | 'edit.cut' | 'edit.copy' | 'edit.copy-merged' | 'edit.paste' | 'edit.paste-in-place'
  | 'edit.clear' | 'edit.fill-foreground' | 'edit.fill-background' | 'edit.fill-foreground-preserve'
  | 'edit.free-transform' | 'edit.swap-colors' | 'edit.default-colors'
  | 'select.all' | 'select.deselect' | 'select.reselect' | 'select.inverse'
  | 'select.feather' | 'select.expand' | 'select.contract' | 'select.load-layer-alpha'
  | 'layer.new' | 'layer.duplicate' | 'layer.delete' | 'layer.via-copy' | 'layer.via-cut'
  | 'layer.merge-down' | 'layer.merge-visible' | 'layer.stamp-visible' | 'layer.flatten'
  | 'layer.add-mask' | 'layer.delete-mask' | 'layer.apply-mask' | 'layer.toggle-mask' | 'layer.rasterize'
  | 'layer.from-background' | 'layer.toggle-clipping'
  | 'layer.raise' | 'layer.lower' | 'layer.to-front' | 'layer.to-back'
  | 'layer.select-above' | 'layer.select-below'
  | 'image.size' | 'image.canvas-size' | 'image.crop-to-selection' | 'image.trim'
  | 'image.rotate-cw' | 'image.rotate-ccw' | 'image.rotate-180' | 'image.rotate-arbitrary'
  | 'image.flip-horizontal' | 'image.flip-vertical'
  | 'image.auto-tone' | 'image.auto-contrast' | 'image.auto-color' | 'image.desaturate'
  | `adjust.${AdjustmentType}`
  | `adjustment-layer.${AdjustmentType}`
  | `filter.${FilterType}`
  | 'filter.repeat'
  | 'view.zoom-in' | 'view.zoom-out' | 'view.fit' | 'view.actual-pixels' | 'view.toggle-panels'
  | 'brush.smaller' | 'brush.larger' | 'brush.softer' | 'brush.harder'
  | `tool.${ToolId}`
  | 'tool.cycle-marquee' | 'tool.cycle-lasso' | 'tool.cycle-gradient' | 'tool.cycle-shape'
  | 'session.commit' | 'session.cancel'

export interface ShortcutBinding {
  /** Normalised chord: modifiers in the order ctrl+alt+shift, then the key, e.g. 'ctrl+shift+e', 'alt+backspace', '['. */
  readonly chord: string
  readonly command: CommandId
  /** Active only when no text input or text-tool session has focus (single-key tool shortcuts). */
  readonly requiresCanvasFocus: boolean
  /** Fire on key repeat ([ and ] do; tool switches do not). */
  readonly repeat: boolean
}

// ---------------------------------------------------------------------------------------------
// Host integration (the only surface main.tsx uses)
// ---------------------------------------------------------------------------------------------

export interface AdvancedHost {
  /** Allocates a new content revision from the host-wide counter shared with Simple mode. */
  nextRevision(): number
  /** The document is now at `revision` (after an edit, undo or redo). Host recomputes Modified. */
  setRevision(revision: number): void
  notify(message: string, tone?: 'normal' | 'error'): void
  /** True while a host modal (print dialog, unsaved-changes prompt) owns input. */
  isSuspended(): boolean
  /** The user pressed "Simple". The host runs the exit flow. */
  requestExit(): void
  /** Host save path (Save / Save As), shared with the title bar buttons. */
  save(forceDialog: boolean): Promise<boolean>
  openExportMenu(): void
  print(): void
  /** PNG bytes to the system clipboard. */
  copyPng(png: Uint8Array): Promise<void>
  /** System clipboard image as PNG bytes, or null. */
  readClipboardImage(): Promise<Uint8Array | null>
}

export interface AdvancedEditorHandle {
  /** Commit text edits; cancel or commit open crop/transform sessions; wait for pending worker jobs. */
  settle(): Promise<void>
  /** Full-resolution composite in a new canvas. The caller must release it (width = height = 1). */
  renderFlattened(): Promise<HTMLCanvasElement>
  /** Layered PSD bytes with the merged composite embedded. */
  encodePsd(): Promise<Uint8Array>
  /** True when flattening loses nothing (see design section 7.3). */
  isFlatEquivalent(): boolean
  layerCount(): number
  size(): Size
  /** Whether the composite has any alpha < 255 (drives JPEG warnings and the Simple 'Color' row). */
  hasTransparency(): Promise<boolean>
  /** Active-layer selection (or merged when merged = true) as a new canvas; caller releases. Null when empty. */
  copySelection(merged: boolean): Promise<HTMLCanvasElement | null>
  /** Places an image as a new layer centred in the view (paste, drop). */
  placeImage(pixels: PixelBuffer, name: string): void
  /** PSD features that would be lost by overwriting the source file. */
  fidelityIssues(): readonly PsdFidelityIssue[]
  focus(): void
}

export type AdvancedInitialContent =
  | {
    readonly kind: 'canvas'
    /** The live Simple canvas; read once in 256-row stripes, never modified. */
    readonly canvas: HTMLCanvasElement
    readonly name: string
    readonly hasAlpha: boolean
    readonly ppi: number
  }
  | {
    readonly kind: 'document'
    readonly document: ImportedDocument
    readonly name: string
  }

export interface AdvancedEditorProps {
  readonly host: AdvancedHost
  readonly initial: AdvancedInitialContent
  /** Mirrors host.isSuspended() for rendering (inert, no shortcuts). */
  readonly suspended: boolean
  readonly onReady?: (handle: AdvancedEditorHandle) => void
}

// ---------------------------------------------------------------------------------------------
// PSD (src/advanced/psd.ts + psdMapping.ts)
// ---------------------------------------------------------------------------------------------

export type PsdIssueCode =
  | 'layer-effects'
  | 'smart-object'
  | 'vector-mask'
  | 'unsupported-adjustment'
  | 'group-flattened'
  | 'unsupported-blend-mode'
  | 'fill-opacity'
  | 'text-rerender'
  | 'bit-depth-reduced'
  | 'color-profile'
  | 'knockout-or-advanced-blending'

export interface PsdFidelityIssue {
  readonly code: PsdIssueCode
  readonly layerName: string | null
  readonly detail: string
}

export interface ImportedDocument {
  readonly width: number
  readonly height: number
  readonly ppi: number
  readonly layers: readonly Layer[]
  /** The file's merged image (exact Photoshop appearance), when present. */
  readonly composite: PixelBuffer | null
  readonly issues: readonly PsdFidelityIssue[]
}

export interface PsdImportOptions {
  readonly maxPixels: number
  readonly maxDimension: number
  /** Passed to ag-psd totalMemoryLimit and checked per decoded layer. */
  readonly memoryLimitBytes: number
  /** 'flattened' opens only the merged image as a Background layer (exact look). */
  readonly mode: 'layers' | 'flattened'
  /** Maps PostScript font names to CSS families; defaults to a heuristic. */
  readonly resolveFont?: (postScriptName: string) => { readonly family: string; readonly weight: 400 | 700; readonly italic: boolean }
}

export interface PsdExportOptions {
  /** Ask Photoshop to re-render text layers on open (keeps them editable there). Default true. */
  readonly invalidateText: boolean
}

export type ImportPsd = (bytes: Uint8Array, options?: Partial<PsdImportOptions>) => Promise<ImportedDocument>
export type ExportPsd = (
  doc: Pick<DocumentState, 'width' | 'height' | 'ppi' | 'layers'>,
  composite: PixelBuffer,
  options?: Partial<PsdExportOptions>,
) => Promise<Uint8Array>

// ---------------------------------------------------------------------------------------------
// Runtime constants (kept here so every package agrees on limits)
// ---------------------------------------------------------------------------------------------

export const TILE_SIZE = 256
export const TILE_SHIFT = 8
export const COALESCE_MS = 1200

export const LIMITS = Object.freeze({
  maxDimension: 20_000,
  maxPixels: 50_000_000,
  maxLayers: 200,
  historyMaxEntries: 50,
  historyBudgetBytes: 768 * 1024 * 1024,
  documentBudgetBytes: 3 * 1024 * 1024 * 1024,
  compositeTileCacheTiles: 512,
  psdMemoryLimitBytes: 1536 * 1024 * 1024,
  frameBudgetMs: 8,
})

export const DEFAULT_LOCKS: LayerLocks = Object.freeze({ pixels: false, position: false, transparency: false })
export const BACKGROUND_LOCKS: LayerLocks = Object.freeze({ pixels: false, position: true, transparency: true })

/** Photoshop menu order; '-' marks a separator. */
export const BLEND_MODE_MENU: readonly (BlendMode | '-')[] = Object.freeze([
  'normal', 'dissolve', '-',
  'darken', 'multiply', 'color-burn', 'linear-burn', 'darker-color', '-',
  'lighten', 'screen', 'color-dodge', 'linear-dodge', 'lighter-color', '-',
  'overlay', 'soft-light', 'hard-light', 'vivid-light', 'linear-light', 'pin-light', 'hard-mix', '-',
  'difference', 'exclusion', 'subtract', 'divide', '-',
  'hue', 'saturation', 'color', 'luminosity',
] as const)

// Re-exported so feature code can import every shared type from one place.
export type { Rect }
