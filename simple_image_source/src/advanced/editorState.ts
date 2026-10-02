// src/advanced/editorState.ts (WP6)
// The editor store of the Advanced editor (design 5.14, 8.2 EditorState): the active tool, foreground and
// background colours, tool options, swatches, panels and the last filter.
//   - createEditorStore(storageKey): a tiny external store (getState / subscribe / update / updateOptions)
//     that React reads with useSyncExternalStore. Tool options, colours, swatches, panel visibility and the
//     tool last chosen in each toolbox slot persist in localStorage (a per-viewer convenience): written a
//     moment after the last change and on dispose; anything unreadable or out of range falls back to the
//     defaults field by field, so a damaged entry can never stop the editor from opening.
//   - Pure helpers the commands share: Photoshop's [ / ] size ladder, hardness steps, the option set a tool
//     paints with, digit-key opacity.
// Pure and DOM-free at module load (Node tests import it); storage is injected or found on globalThis.
import type { BlendMode, GradientKind, Rgb8, Rgba8, SelectionOp } from '../imaging/types.ts'
import type {
  AspectPreset,
  BrushOptions,
  EditorState,
  EditorStore,
  PanelVisibility,
  SampleSource,
  ShapeKind,
  TextStyle,
  ToolId,
  ToolOptions,
  ViewTransform,
} from './types.ts'
import { BLEND_MODE_MENU } from './types.ts'

// ---------------------------------------------------------------------------------------------
// Defaults
// ---------------------------------------------------------------------------------------------

function freeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const key of Object.keys(value as object)) freeze((value as Record<string, unknown>)[key])
    Object.freeze(value)
  }
  return value
}

const BRUSH_DEFAULTS: BrushOptions = {
  size: 24,
  hardness: 0.8,
  opacity: 1,
  flow: 1,
  spacing: 0.25,
  smoothing: 0.1,
  pressureSize: true,
  pressureOpacity: false,
  blendMode: 'normal',
}

export const DEFAULT_TEXT_STYLE: TextStyle = freeze({
  fontFamily: 'Segoe UI',
  fontSize: 48,
  fontWeight: 400,
  italic: false,
  underline: false,
  color: { r: 0, g: 0, b: 0 },
  align: 'left',
  lineHeight: 1.2,
  letterSpacing: 0,
})

/** Photoshop's factory defaults for every tool. */
export const DEFAULT_TOOL_OPTIONS: ToolOptions = freeze({
  brush: { ...BRUSH_DEFAULTS },
  eraser: { ...BRUSH_DEFAULTS, hardness: 1, pressureSize: false },
  clone: { ...BRUSH_DEFAULTS, hardness: 0.5, aligned: true, sample: 'current' },
  heal: { size: 30, hardness: 0.5, sample: 'current' },
  marquee: { op: 'replace', feather: 0, antiAlias: true, style: 'normal', ratio: { width: 1, height: 1 }, fixedSize: { width: 64, height: 64 } },
  lasso: { op: 'replace', feather: 0, antiAlias: true },
  wand: { op: 'replace', tolerance: 32, contiguous: true, antiAlias: true, sample: 'current' },
  bucket: { tolerance: 32, contiguous: true, antiAlias: true, sample: 'current', opacity: 1, blendMode: 'normal' },
  gradient: { kind: 'linear', preset: 'foreground-background', reverse: false, dither: true, opacity: 1, blendMode: 'normal' },
  crop: { aspect: 'free', portrait: false, deleteCroppedPixels: true, overlay: 'thirds' },
  text: { ...DEFAULT_TEXT_STYLE, color: { r: 0, g: 0, b: 0 } },
  shape: { kind: 'rectangle', fill: { r: 64, g: 128, b: 230, a: 255 }, stroke: null, strokeWidth: 4, cornerRadius: 0 },
  eyedropper: { size: 1, sample: 'all' },
  move: { autoSelect: false, showTransformControls: false },
} as ToolOptions)

export const DEFAULT_PANELS: PanelVisibility = freeze({ layers: true, properties: true, history: true, color: true })

export const DEFAULT_FOREGROUND: Rgb8 = freeze({ r: 0, g: 0, b: 0 })
export const DEFAULT_BACKGROUND: Rgb8 = freeze({ r: 255, g: 255, b: 255 })

export const DEFAULT_EDITOR_STATE: EditorState = freeze({
  tool: 'move',
  springFrom: null,
  foreground: DEFAULT_FOREGROUND,
  background: DEFAULT_BACKGROUND,
  options: DEFAULT_TOOL_OPTIONS,
  swatches: [
    { r: 0, g: 0, b: 0 }, { r: 255, g: 255, b: 255 }, { r: 128, g: 128, b: 128 }, { r: 237, g: 28, b: 36 },
    { r: 255, g: 127, b: 39 }, { r: 255, g: 242, b: 0 }, { r: 34, g: 177, b: 76 }, { r: 0, g: 162, b: 232 },
    { r: 63, g: 72, b: 204 }, { r: 163, g: 73, b: 164 },
  ],
  recentColors: [],
  view: { zoom: 1, offsetX: 0, offsetY: 0 },
  panels: DEFAULT_PANELS,
  lastFilter: null,
} as EditorState)

export const EDITOR_STORAGE_KEY = 'simple-image:advanced-editor'

export const TOOL_IDS: readonly ToolId[] = Object.freeze([
  'move', 'marquee-rect', 'marquee-ellipse', 'lasso', 'lasso-polygon', 'magic-wand', 'crop', 'eyedropper',
  'spot-healing', 'brush', 'clone-stamp', 'eraser', 'gradient', 'paint-bucket', 'text', 'shape', 'hand', 'zoom',
] as const)

/** Toolbox slots whose tools cycle with Shift + the key (same grouping as tools/index.ts TOOL_GROUPS). */
export const TOOL_SLOTS: readonly (readonly ToolId[])[] = Object.freeze([
  Object.freeze(['marquee-rect', 'marquee-ellipse'] as ToolId[]),
  Object.freeze(['lasso', 'lasso-polygon'] as ToolId[]),
  Object.freeze(['gradient', 'paint-bucket'] as ToolId[]),
])

const SHAPE_KINDS: readonly ShapeKind[] = Object.freeze(['rectangle', 'ellipse', 'line', 'arrow'] as const)
const SAMPLE_SOURCES: readonly SampleSource[] = Object.freeze(['current', 'current-below', 'all'] as const)
const SELECTION_OPS: readonly SelectionOp[] = Object.freeze(['replace', 'add', 'subtract', 'intersect'] as const)
const GRADIENT_KINDS: readonly GradientKind[] = Object.freeze(['linear', 'radial', 'angle', 'reflected', 'diamond'] as const)
const ASPECTS: readonly AspectPreset[] = Object.freeze(['free', 'original', '1:1', '4:3', '3:2', '16:9', '5:4', '7:5'] as const)
const BLEND_MODES: readonly BlendMode[] = Object.freeze(BLEND_MODE_MENU.filter((mode): mode is BlendMode => mode !== '-'))

export function isToolId(value: unknown): value is ToolId {
  return typeof value === 'string' && (TOOL_IDS as readonly string[]).includes(value)
}

/** The toolbox slot (cycle group) of a tool, or null for single-tool slots. */
export function slotOf(tool: ToolId): readonly ToolId[] | null {
  return TOOL_SLOTS.find((slot) => slot.includes(tool)) ?? null
}

// ---------------------------------------------------------------------------------------------
// Validation (stored values are untrusted: every field is checked against the default's shape)
// ---------------------------------------------------------------------------------------------

function num(value: unknown, fallback: number, min: number, max: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? Math.min(max, Math.max(min, value)) : fallback
}

function int(value: unknown, fallback: number, min: number, max: number): number {
  return Math.round(num(value, fallback, min, max))
}

function bool(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback
}

function oneOf<T extends string | number>(value: unknown, allowed: readonly T[], fallback: T): T {
  return (allowed as readonly unknown[]).includes(value) ? value as T : fallback
}

function record(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
}

export function sanitizeRgb(value: unknown, fallback: Rgb8): Rgb8 {
  const source = record(value)
  if (!('r' in source) || !('g' in source) || !('b' in source)) return fallback
  return { r: int(source.r, fallback.r, 0, 255), g: int(source.g, fallback.g, 0, 255), b: int(source.b, fallback.b, 0, 255) }
}

function sanitizeRgba(value: unknown, fallback: Rgba8 | null): Rgba8 | null {
  if (value === null) return null
  const source = record(value)
  if (!('r' in source)) return fallback
  const base = fallback ?? { r: 0, g: 0, b: 0, a: 255 }
  return { ...sanitizeRgb(source, base), a: int(source.a, 255, 0, 255) }
}

function sanitizeBrush(value: unknown, fallback: BrushOptions): BrushOptions {
  const source = record(value)
  return {
    size: num(source.size, fallback.size, 1, 5000),
    hardness: num(source.hardness, fallback.hardness, 0, 1),
    opacity: num(source.opacity, fallback.opacity, 0, 1),
    flow: num(source.flow, fallback.flow, 0.01, 1),
    spacing: num(source.spacing, fallback.spacing, 0.01, 10),
    smoothing: num(source.smoothing, fallback.smoothing, 0, 1),
    pressureSize: bool(source.pressureSize, fallback.pressureSize),
    pressureOpacity: bool(source.pressureOpacity, fallback.pressureOpacity),
    blendMode: oneOf(source.blendMode, BLEND_MODES, fallback.blendMode),
  }
}

function sanitizeSize(value: unknown, fallback: { readonly width: number; readonly height: number }, min: number): { width: number; height: number } {
  const source = record(value)
  return { width: num(source.width, fallback.width, min, 100_000), height: num(source.height, fallback.height, min, 100_000) }
}

function sanitizeText(value: unknown, fallback: TextStyle): TextStyle {
  const source = record(value)
  const family = typeof source.fontFamily === 'string' && source.fontFamily.trim() ? source.fontFamily.trim().slice(0, 120) : fallback.fontFamily
  return {
    fontFamily: family,
    fontSize: num(source.fontSize, fallback.fontSize, 1, 5000),
    fontWeight: oneOf(source.fontWeight, [400, 700] as const, fallback.fontWeight),
    italic: bool(source.italic, fallback.italic),
    underline: bool(source.underline, fallback.underline),
    color: sanitizeRgb(source.color, fallback.color),
    align: oneOf(source.align, ['left', 'center', 'right'] as const, fallback.align),
    lineHeight: num(source.lineHeight, fallback.lineHeight, 0.5, 5),
    letterSpacing: num(source.letterSpacing, fallback.letterSpacing, -100, 1000),
  }
}

/** Tool options from untrusted input: every missing, mistyped or out-of-range field takes its default. */
export function sanitizeOptions(value: unknown, defaults: ToolOptions = DEFAULT_TOOL_OPTIONS): ToolOptions {
  const source = record(value)
  const d = defaults
  const clone = record(source.clone)
  const heal = record(source.heal)
  const marquee = record(source.marquee)
  const lasso = record(source.lasso)
  const wand = record(source.wand)
  const bucket = record(source.bucket)
  const gradient = record(source.gradient)
  const crop = record(source.crop)
  const shape = record(source.shape)
  const eyedropper = record(source.eyedropper)
  const move = record(source.move)
  return {
    brush: sanitizeBrush(source.brush, d.brush),
    eraser: sanitizeBrush(source.eraser, d.eraser),
    clone: {
      ...sanitizeBrush(clone, d.clone),
      aligned: bool(clone.aligned, d.clone.aligned),
      sample: oneOf(clone.sample, SAMPLE_SOURCES, d.clone.sample),
    },
    heal: {
      size: num(heal.size, d.heal.size, 1, 5000),
      hardness: num(heal.hardness, d.heal.hardness, 0, 1),
      sample: oneOf(heal.sample, SAMPLE_SOURCES, d.heal.sample),
    },
    marquee: {
      op: oneOf(marquee.op, SELECTION_OPS, d.marquee.op),
      feather: num(marquee.feather, d.marquee.feather, 0, 1000),
      antiAlias: bool(marquee.antiAlias, d.marquee.antiAlias),
      style: oneOf(marquee.style, ['normal', 'fixed-ratio', 'fixed-size'] as const, d.marquee.style),
      ratio: sanitizeSize(marquee.ratio, d.marquee.ratio, 0.001),
      fixedSize: sanitizeSize(marquee.fixedSize, d.marquee.fixedSize, 1),
    },
    lasso: {
      op: oneOf(lasso.op, SELECTION_OPS, d.lasso.op),
      feather: num(lasso.feather, d.lasso.feather, 0, 1000),
      antiAlias: bool(lasso.antiAlias, d.lasso.antiAlias),
    },
    wand: {
      op: oneOf(wand.op, SELECTION_OPS, d.wand.op),
      tolerance: int(wand.tolerance, d.wand.tolerance, 0, 255),
      contiguous: bool(wand.contiguous, d.wand.contiguous),
      antiAlias: bool(wand.antiAlias, d.wand.antiAlias),
      sample: oneOf(wand.sample, SAMPLE_SOURCES, d.wand.sample),
    },
    bucket: {
      tolerance: int(bucket.tolerance, d.bucket.tolerance, 0, 255),
      contiguous: bool(bucket.contiguous, d.bucket.contiguous),
      antiAlias: bool(bucket.antiAlias, d.bucket.antiAlias),
      sample: oneOf(bucket.sample, SAMPLE_SOURCES, d.bucket.sample),
      opacity: num(bucket.opacity, d.bucket.opacity, 0, 1),
      blendMode: oneOf(bucket.blendMode, BLEND_MODES, d.bucket.blendMode),
    },
    gradient: {
      kind: oneOf(gradient.kind, GRADIENT_KINDS, d.gradient.kind),
      preset: oneOf(gradient.preset, ['foreground-background', 'foreground-transparent', 'black-white', 'custom'] as const, d.gradient.preset),
      reverse: bool(gradient.reverse, d.gradient.reverse),
      dither: bool(gradient.dither, d.gradient.dither),
      opacity: num(gradient.opacity, d.gradient.opacity, 0, 1),
      blendMode: oneOf(gradient.blendMode, BLEND_MODES, d.gradient.blendMode),
    },
    crop: {
      aspect: oneOf(crop.aspect, ASPECTS, d.crop.aspect),
      portrait: bool(crop.portrait, d.crop.portrait),
      deleteCroppedPixels: bool(crop.deleteCroppedPixels, d.crop.deleteCroppedPixels),
      overlay: oneOf(crop.overlay, ['thirds', 'grid', 'none'] as const, d.crop.overlay),
    },
    text: sanitizeText(source.text, d.text),
    shape: {
      kind: oneOf(shape.kind, SHAPE_KINDS, d.shape.kind),
      fill: 'fill' in shape ? sanitizeRgba(shape.fill, d.shape.fill) : d.shape.fill,
      stroke: 'stroke' in shape ? sanitizeRgba(shape.stroke, d.shape.stroke) : d.shape.stroke,
      strokeWidth: num(shape.strokeWidth, d.shape.strokeWidth, 0, 1000),
      cornerRadius: num(shape.cornerRadius, d.shape.cornerRadius, 0, 10_000),
    },
    eyedropper: {
      size: oneOf(eyedropper.size, [1, 3, 5, 11] as const, d.eyedropper.size),
      sample: oneOf(eyedropper.sample, SAMPLE_SOURCES, d.eyedropper.sample),
    },
    move: {
      autoSelect: bool(move.autoSelect, d.move.autoSelect),
      showTransformControls: bool(move.showTransformControls, d.move.showTransformControls),
    },
  }
}

function sanitizeColorList(value: unknown, fallback: readonly Rgb8[], max: number): Rgb8[] {
  if (!Array.isArray(value)) return [...fallback]
  const out: Rgb8[] = []
  for (const entry of value) {
    const color = sanitizeRgb(entry, { r: -1, g: -1, b: -1 })
    if (color.r < 0) continue
    out.push(color)
    if (out.length >= max) break
  }
  return out
}

function sanitizePanels(value: unknown): PanelVisibility {
  const source = record(value)
  return {
    layers: bool(source.layers, DEFAULT_PANELS.layers),
    properties: bool(source.properties, DEFAULT_PANELS.properties),
    history: bool(source.history, DEFAULT_PANELS.history),
    color: bool(source.color, DEFAULT_PANELS.color),
  }
}

function sanitizeView(value: unknown, fallback: ViewTransform): ViewTransform {
  const source = record(value)
  const zoom = num(source.zoom, fallback.zoom, 1e-4, 1e4)
  return { zoom: zoom > 0 ? zoom : fallback.zoom, offsetX: num(source.offsetX, fallback.offsetX, -1e9, 1e9), offsetY: num(source.offsetY, fallback.offsetY, -1e9, 1e9) }
}

export const MAX_SWATCHES = 60
export const MAX_RECENT_COLORS = 12

/** The colour first in the list, without duplicates, at most `max` long. */
export function rememberColor(list: readonly Rgb8[], color: Rgb8, max = MAX_RECENT_COLORS): Rgb8[] {
  const next = [color, ...list.filter((entry) => entry.r !== color.r || entry.g !== color.g || entry.b !== color.b)]
  return next.slice(0, Math.max(1, max))
}

export function sameColor(a: Rgb8 | null | undefined, b: Rgb8 | null | undefined): boolean {
  return Boolean(a && b && a.r === b.r && a.g === b.g && a.b === b.b)
}

// ---------------------------------------------------------------------------------------------
// The store
// ---------------------------------------------------------------------------------------------

export interface StorageLike {
  getItem(key: string): string | null
  setItem(key: string, value: string): void
}

export interface EditorStoreOptions {
  /** Where settings persist; default globalThis.localStorage when reachable, null = memory only. */
  readonly storage?: StorageLike | null
  /** Delay before writing after the last change (ms). Default 400. */
  readonly writeDelayMs?: number
  /** Starting values that override stored ones (tests). */
  readonly initial?: Partial<EditorState>
}

/** The shared EditorStore contract plus what the editor shell needs. */
export interface AdvancedEditorStore extends EditorStore {
  /** The tool a toolbox slot shows (the one last chosen in that slot); the tool itself for single slots. */
  slotTool(tool: ToolId): ToolId
  /** Writes pending settings to storage now. */
  flush(): void
  /** Flushes and stops listening; later updates still work in memory. */
  dispose(): void
}

interface Persisted {
  readonly version: 1
  readonly tool: ToolId
  readonly foreground: Rgb8
  readonly background: Rgb8
  readonly options: ToolOptions
  readonly swatches: readonly Rgb8[]
  readonly recentColors: readonly Rgb8[]
  readonly panels: PanelVisibility
  readonly slots: Readonly<Record<string, ToolId>>
}

function defaultStorage(): StorageLike | null {
  try {
    const storage = (globalThis as { localStorage?: StorageLike }).localStorage
    return storage && typeof storage.getItem === 'function' ? storage : null
  } catch {
    // Access can throw (blocked site data); settings then live in memory only.
    return null
  }
}

function readPersisted(storage: StorageLike | null, key: string): Partial<Persisted> {
  if (!storage) return {}
  try {
    const text = storage.getItem(key)
    if (!text) return {}
    const value = JSON.parse(text) as unknown
    return value && typeof value === 'object' ? value as Partial<Persisted> : {}
  } catch {
    return {}
  }
}

export function createEditorStore(storageKey: string = EDITOR_STORAGE_KEY, options: EditorStoreOptions = {}): AdvancedEditorStore {
  const storage = options.storage === undefined ? defaultStorage() : options.storage
  const delay = Number.isFinite(options.writeDelayMs) ? Math.max(0, Number(options.writeDelayMs)) : 400
  const saved = readPersisted(storage, storageKey)
  const slots = new Map<string, ToolId>()
  for (const slot of TOOL_SLOTS) {
    const stored = record(saved.slots)[slot[0]]
    slots.set(slot[0], isToolId(stored) && slot.includes(stored) ? stored : slot[0])
  }
  let state: EditorState = {
    ...DEFAULT_EDITOR_STATE,
    tool: isToolId(saved.tool) && saved.tool !== 'hand' && saved.tool !== 'zoom' ? saved.tool : DEFAULT_EDITOR_STATE.tool,
    foreground: sanitizeRgb(saved.foreground, DEFAULT_FOREGROUND),
    background: sanitizeRgb(saved.background, DEFAULT_BACKGROUND),
    options: sanitizeOptions(saved.options),
    swatches: 'swatches' in saved ? sanitizeColorList(saved.swatches, DEFAULT_EDITOR_STATE.swatches, MAX_SWATCHES) : [...DEFAULT_EDITOR_STATE.swatches],
    recentColors: sanitizeColorList(saved.recentColors, [], MAX_RECENT_COLORS),
    panels: sanitizePanels(saved.panels),
    ...(options.initial ?? {}),
  }
  const listeners = new Set<() => void>()
  let timer: ReturnType<typeof setTimeout> | null = null
  let dirty = false
  let disposed = false

  const persist = () => {
    timer = null
    if (!dirty || !storage) return
    dirty = false
    const payload: Persisted = {
      version: 1,
      // A spring-loaded tool (held Space / Alt) is not a choice; remember the tool it returns to.
      tool: state.springFrom ?? state.tool,
      foreground: state.foreground,
      background: state.background,
      options: state.options,
      swatches: state.swatches,
      recentColors: state.recentColors,
      panels: state.panels,
      slots: Object.fromEntries(slots),
    }
    try {
      storage.setItem(storageKey, JSON.stringify(payload))
    } catch {
      // Quota or blocked storage: settings stay for this session only.
    }
  }

  const schedule = () => {
    dirty = true
    if (!storage || disposed) return
    if (timer !== null) clearTimeout(timer)
    timer = setTimeout(persist, delay)
  }

  const emit = () => {
    for (const listener of [...listeners]) {
      try {
        listener()
      } catch (error) {
        console.error(error)
      }
    }
  }

  const store: AdvancedEditorStore = {
    getState: () => state,
    subscribe(listener: () => void): () => void {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    update(patch: Partial<Omit<EditorState, 'options'>>): void {
      if (!patch || typeof patch !== 'object') return
      const next: Record<string, unknown> = { ...state }
      let changed = false
      let persistent = false
      for (const [key, value] of Object.entries(patch)) {
        if (value === undefined || key === 'options') continue
        let clean: unknown = value
        switch (key) {
          case 'tool':
            if (!isToolId(value)) continue
            break
          case 'springFrom':
            if (value !== null && !isToolId(value)) continue
            break
          case 'foreground':
          case 'background':
            clean = sanitizeRgb(value, state[key])
            if (sameColor(clean as Rgb8, state[key])) continue
            break
          case 'swatches':
            clean = sanitizeColorList(value, state.swatches, MAX_SWATCHES)
            break
          case 'recentColors':
            clean = sanitizeColorList(value, state.recentColors, MAX_RECENT_COLORS)
            break
          case 'panels':
            clean = sanitizePanels({ ...state.panels, ...record(value) })
            break
          case 'view':
            clean = sanitizeView(value, state.view)
            break
          case 'lastFilter':
            if (value !== null && (typeof value !== 'object' || typeof (value as { type?: unknown }).type !== 'string')) continue
            break
          default:
            continue
        }
        if ((next as Record<string, unknown>)[key] === clean) continue
        next[key] = clean
        changed = true
        if (key !== 'view' && key !== 'lastFilter') persistent = true
      }
      if (!changed) return
      const nextState = next as unknown as EditorState
      if (nextState.tool !== state.tool) {
        const slot = slotOf(nextState.tool)
        if (slot && !nextState.springFrom) slots.set(slot[0], nextState.tool)
      }
      state = nextState
      if (persistent) schedule()
      emit()
    },
    updateOptions<K extends keyof ToolOptions>(tool: K, patch: Partial<ToolOptions[K]>): void {
      if (!patch || typeof patch !== 'object' || !(tool in state.options)) return
      const current = state.options[tool]
      const merged = { ...current, ...patch }
      const clean = sanitizeOptions({ ...state.options, [tool]: merged })[tool]
      const before = JSON.stringify(current)
      if (JSON.stringify(clean) === before) return
      state = { ...state, options: { ...state.options, [tool]: clean } }
      schedule()
      emit()
    },
    slotTool(tool: ToolId): ToolId {
      const slot = slotOf(tool)
      if (!slot) return tool
      const remembered = slots.get(slot[0])
      return remembered && slot.includes(remembered) ? remembered : slot[0]
    },
    flush(): void {
      if (timer !== null) {
        clearTimeout(timer)
        timer = null
      }
      persist()
    },
    dispose(): void {
      if (disposed) return
      store.flush()
      disposed = true
      listeners.clear()
    },
  }
  return store
}

// ---------------------------------------------------------------------------------------------
// Option helpers shared by commands, the options bar and the panels
// ---------------------------------------------------------------------------------------------

export type BrushOptionKey = 'brush' | 'eraser' | 'clone'
export type SizeOptionKey = BrushOptionKey | 'heal'
export type OpacityOptionKey = BrushOptionKey | 'bucket' | 'gradient'

/** The option set holding a tool's brush size ([ and ]), or null when the tool has no brush. */
export function sizeOptionKey(tool: ToolId): SizeOptionKey | null {
  switch (tool) {
    case 'brush': return 'brush'
    case 'eraser': return 'eraser'
    case 'clone-stamp': return 'clone'
    case 'spot-healing': return 'heal'
    default: return null
  }
}

/** The option set whose opacity the digit keys set, or null (digits then set the layer opacity). */
export function opacityOptionKey(tool: ToolId): OpacityOptionKey | null {
  switch (tool) {
    case 'brush': return 'brush'
    case 'eraser': return 'eraser'
    case 'clone-stamp': return 'clone'
    case 'paint-bucket': return 'bucket'
    case 'gradient': return 'gradient'
    default: return null
  }
}

/** The option set whose flow Shift+digit sets, or null. */
export function flowOptionKey(tool: ToolId): BrushOptionKey | null {
  switch (tool) {
    case 'brush': return 'brush'
    case 'eraser': return 'eraser'
    case 'clone-stamp': return 'clone'
    default: return null
  }
}

export const MIN_BRUSH_SIZE = 1
export const MAX_BRUSH_SIZE = 5000

/**
 * Photoshop's [ / ] ladder: 1 px steps below 10, 10 px up to 100, 25 px up to 200, 50 px up to 300, then
 * 100 px; sizes off the ladder snap to the next rung in the direction of travel.
 */
export function stepBrushSize(size: number, direction: 1 | -1): number {
  const current = Math.min(MAX_BRUSH_SIZE, Math.max(MIN_BRUSH_SIZE, Number.isFinite(size) ? size : MIN_BRUSH_SIZE))
  if (direction > 0) {
    const step = current < 10 ? 1 : current < 100 ? 10 : current < 200 ? 25 : current < 300 ? 50 : 100
    return Math.min(MAX_BRUSH_SIZE, Math.floor(current / step + 1e-9) * step + step)
  }
  const step = current <= 10 ? 1 : current <= 100 ? 10 : current <= 200 ? 25 : current <= 300 ? 50 : 100
  return Math.max(MIN_BRUSH_SIZE, Math.ceil(current / step - 1e-9) * step - step)
}

/** Shift+[ / Shift+]: hardness in 25% steps (0, 25, 50, 75, 100). */
export function stepHardness(hardness: number, direction: 1 | -1): number {
  const current = Math.min(1, Math.max(0, Number.isFinite(hardness) ? hardness : 1))
  const quarters = current * 4
  const next = direction > 0 ? Math.floor(quarters + 1e-9) + 1 : Math.ceil(quarters - 1e-9) - 1
  return Math.min(1, Math.max(0, next / 4))
}

/** Two digits typed within this many ms set an exact percentage ("4" then "5" = 45%). */
export const DIGIT_CHAIN_MS = 700

export interface DigitMemory {
  readonly digit: number
  readonly time: number
  readonly target: string
}

/**
 * Photoshop digit keys: 1 = 10% ... 9 = 90%, 0 = 100%; a second digit typed quickly after the first makes
 * an exact two-digit value ("0" "5" = 5%, "0" "0" = 0%). Returns the opacity (0..1) and what to remember.
 */
export function opacityFromDigit(digit: number, time: number, target: string, previous: DigitMemory | null): { readonly value: number; readonly memory: DigitMemory | null } {
  const d = Math.min(9, Math.max(0, Math.round(digit)))
  if (previous && previous.target === target && time - previous.time >= 0 && time - previous.time <= DIGIT_CHAIN_MS) {
    return { value: (previous.digit * 10 + d) / 100, memory: null }
  }
  return { value: d === 0 ? 1 : d / 10, memory: { digit: d, time, target } }
}
