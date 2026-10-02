// src/advanced/psd.ts (WP7)
// Layered Photoshop documents through ag-psd 31.0.2 (MIT), design 6.2.
//   importPsd(): readPsd with useRawData, so channel data stays compressed until each layer is decoded and
//     peak memory is the file plus one decoded layer. Layers (bottom to top) become raster, text and
//     adjustment layers with their masks; vector masks are rasterized; pass-through groups dissolve
//     ("Group / Layer" names) and other groups flatten through the pure compositor; everything Simple
//     cannot keep is reported as a PsdFidelityIssue. 16/32-bit channels are reduced to 8-bit.
//   exportPsd(): the document as an 8-bit RGB PSD with the merged image embedded (so every reader shows the
//     exact look), raster layers with tight bounds, masks, text layers that Photoshop re-renders on open,
//     adjustment layers as adjustment layers; shapes are written as pixels.
// ag-psd itself is loaded with a dynamic import on first use, so importing this module is cheap and the
// library becomes its own lazy chunk. Pure apart from the canvas hooks configured for ag-psd.
import type { IntRect, MaskBuffer, OpOptions, PixelBuffer } from '../imaging/types.ts'
import type {
  AdjustmentLayer,
  DocumentState,
  ExportPsd,
  ImportPsd,
  Layer,
  LayerMask,
  PsdExportOptions,
  PsdFidelityIssue,
  PsdImportOptions,
  PsdIssueCode,
  RasterCache,
  RasterLayer,
  TextSpec,
} from './types.ts'
import { LIMITS } from './types.ts'
import type {
  Layer as PsdLayer,
  LayerMaskData,
  PixelData,
  Psd,
} from 'ag-psd'
import {
  adjustmentFromPsd,
  adjustmentToPsd,
  advancedBlendingNotes,
  blendFromPsd,
  blendToPsd,
  FILL_OPACITY_SENSITIVE_MODES,
  forceOpaque,
  hasEnabledEffects,
  iccProfileDescription,
  isSrgbProfile,
  locksFromPsd,
  locksToPsd,
  maskChannel,
  multiplyMasks,
  ppiFromResolution,
  PSD_DAMAGED_MESSAGE,
  psdAdjustmentLabel,
  psdColorToRgb,
  psdHeaderProblem,
  quantizeOpacity,
  rasterizeVectorMask,
  readPsdStructure,
  removeWhiteMatte,
  resolutionFromPpi,
  sniffPsd,
  textFromPsd,
  textToPsd,
  toRgba8,
} from './psdMapping.ts'
import type { FontResolver, PsdSampleArray } from './psdMapping.ts'
import { loadLocalFonts, postScriptNameFor, resolvePostScriptFont } from './fonts.ts'
import { createAdjustmentLayer, createLayerMask, createRasterLayer, createTextLayer, layerContentBounds, rasterCacheFrom } from './document.ts'
import { maskFromBuffer, surfaceFromBuffer } from './tiles.ts'
import { compositeRect, flattenDocument } from './composite.ts'
import { estimateRasterBytes } from './memory.ts'
import { createMaskBuffer, pasteMask } from '../imaging/mask.ts'
import { measureText, rasterizeText, specKey } from '../shared/vector.ts'
import { yieldToEventLoop } from '../imaging/buffer.ts'

export { adjustmentFromPsd, adjustmentToPsd, blendFromPsd, blendToPsd, sniffPsd } from './psdMapping.ts'

type AgPsd = typeof import('ag-psd')
type CanvasMode = 'dom' | 'node-test'

// ---------------------------------------------------------------------------------------------
// Loading ag-psd and its canvas hooks
// ---------------------------------------------------------------------------------------------

let agPsdPromise: Promise<AgPsd> | null = null
let requestedCanvas: CanvasMode | null = null
let appliedCanvas: CanvasMode | null = null

function applyCanvas(module: AgPsd, mode: CanvasMode): void {
  if (mode === 'dom') {
    // Pixel data never goes through a canvas (useRawData + typed arrays); these hooks only replace the
    // temporary canvas ag-psd would otherwise keep for createImageData.
    module.initializeCanvas(
      (width, height) => {
        const canvas = document.createElement('canvas')
        canvas.width = width
        canvas.height = height
        return canvas
      },
      (width, height) => {
        try {
          return new ImageData(width, height)
        } catch {
          return { width, height, data: new Uint8ClampedArray(width * height * 4), colorSpace: 'srgb' } as ImageData
        }
      },
    )
  } else {
    module.initializeCanvas(
      () => { throw new Error('no canvas in tests') },
      (width, height) => ({ width, height, data: new Uint8ClampedArray(width * height * 4) }) as unknown as ImageData,
    )
  }
  appliedCanvas = mode
}

/** Chooses ag-psd's canvas hooks: 'dom' in the renderer, 'node-test' in Node (no canvas at all). */
export function configurePsdCanvas(mode: CanvasMode): void {
  requestedCanvas = mode
}

async function loadAgPsd(): Promise<AgPsd> {
  if (!agPsdPromise) {
    agPsdPromise = import('ag-psd')
    agPsdPromise.catch(() => { agPsdPromise = null })
  }
  let module: AgPsd
  try {
    module = await agPsdPromise
  } catch (error) {
    throw new Error(`Simple could not load its Photoshop document support (${error instanceof Error ? error.message : String(error)}). Restart Simple and try again.`)
  }
  const mode = requestedCanvas ?? (typeof document !== 'undefined' ? 'dom' : 'node-test')
  if (appliedCanvas !== mode) applyCanvas(module, mode)
  return module
}

// ---------------------------------------------------------------------------------------------
// Errors, memory and issues
// ---------------------------------------------------------------------------------------------

export const PSD_MEMORY_ERROR_CODE = 'psd-memory-limit'

function formatGigabytes(bytes: number): string {
  return `${(Math.max(0, bytes) / 1024 ** 3).toFixed(1)} GB`
}

/** Friendly error with a machine-readable code (the host can offer "Open flattened" on memory errors). */
function codedError(message: string, code: string): Error {
  const error = new Error(message) as Error & { code: string }
  error.code = code
  return error
}

function memoryError(needed: number, limit: number): Error {
  return codedError(
    `This Photoshop document is too large to open with its layers (it needs about ${formatGigabytes(needed)}; Simple allows ${formatGigabytes(limit)}). `
      + 'Merge or delete layers, or reduce the image size in Photoshop, then try again.',
    PSD_MEMORY_ERROR_CODE,
  )
}

/** ag-psd and allocation failures as plain language; the technical reason stays in parentheses. */
function friendlyReadError(error: unknown, limit: number): Error {
  if (error instanceof Error && (error as Error & { code?: string }).code) return error
  const message = error instanceof Error ? error.message : String(error)
  if (/Exceeded memory limit/i.test(message) || error instanceof RangeError || /allocation failed|out of memory/i.test(message)) {
    return memoryError(limit + 1, limit)
  }
  if (/Color mode not supported/i.test(message)) {
    return new Error('This PSD uses a color mode Simple can\'t edit. In Photoshop choose Image > Mode > RGB Color, then save a copy.')
  }
  if (/Invalid channel count/i.test(message)) {
    return new Error('This Photoshop document has more than 16 channels. Delete extra alpha channels in Photoshop (Channels panel), then save a copy.')
  }
  return new Error(`${PSD_DAMAGED_MESSAGE} (${message})`)
}

class MemoryBudget {
  used = 0
  readonly limit: number

  constructor(limit: number) {
    this.limit = limit
  }

  /** Throws a friendly error when `transient` more bytes would not fit right now. */
  ensure(transient: number): void {
    if (this.used + transient > this.limit) throw memoryError(this.used + transient, this.limit)
  }

  add(bytes: number): void {
    this.used += Math.max(0, bytes)
    if (this.used > this.limit) throw memoryError(this.used, this.limit)
  }
}

class IssueList {
  readonly items: PsdFidelityIssue[] = []

  add(code: PsdIssueCode, layerName: string | null, detail: string): void {
    this.items.push(Object.freeze({ code, layerName, detail }))
  }
}

// ---------------------------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------------------------

interface ImportContext {
  readonly ag: AgPsd
  readonly width: number
  readonly height: number
  readonly options: PsdImportOptions
  readonly budget: MemoryBudget
  readonly issues: IssueList
  readonly resolveFont: FontResolver
}

interface DecodedPixels {
  readonly buffer: PixelBuffer
  /** Document position of buffer (0, 0). */
  readonly left: number
  readonly top: number
}

interface DecodedMask {
  /** Null when the mask is uniform (defaultValue everywhere). */
  readonly buffer: MaskBuffer | null
  readonly left: number
  readonly top: number
  readonly defaultValue: 0 | 255
  readonly disabled: boolean
}

function sampleBytes(depth: number): number {
  return depth === 16 ? 2 : depth === 32 ? 4 : 1
}

function boxOf(record: { left?: number; top?: number; right?: number; bottom?: number }): IntRect {
  const left = Math.round(record.left ?? 0)
  const top = Math.round(record.top ?? 0)
  return { x: left, y: top, width: Math.max(0, Math.round(record.right ?? 0) - left), height: Math.max(0, Math.round(record.bottom ?? 0) - top) }
}

function intersect(a: IntRect, b: IntRect): IntRect | null {
  const x = Math.max(a.x, b.x)
  const y = Math.max(a.y, b.y)
  const right = Math.min(a.x + a.width, b.x + b.width)
  const bottom = Math.min(a.y + a.height, b.y + b.height)
  return right > x && bottom > y ? { x, y, width: right - x, height: bottom - y } : null
}

function union(a: IntRect | null, b: IntRect | null): IntRect | null {
  if (!a) return b
  if (!b) return a
  const x = Math.min(a.x, b.x)
  const y = Math.min(a.y, b.y)
  return { x, y, width: Math.max(a.x + a.width, b.x + b.width) - x, height: Math.max(a.y + a.height, b.y + b.height) - y }
}

function cropPixels(source: PixelBuffer, sourceRect: IntRect, rect: IntRect): PixelBuffer {
  const out = new Uint8ClampedArray(rect.width * rect.height * 4)
  for (let row = 0; row < rect.height; row += 1) {
    const from = ((rect.y - sourceRect.y + row) * source.width + (rect.x - sourceRect.x)) * 4
    out.set(source.data.subarray(from, from + rect.width * 4), row * rect.width * 4)
  }
  return { width: rect.width, height: rect.height, data: out }
}

/** A layer's pixels as straight 8-bit RGBA at their document position; null when the layer has none. */
function decodePixels(context: ImportContext, node: PsdLayer): DecodedPixels | null {
  const box = boxOf(node)
  if (!box.width || !box.height || !node.rawData) return null
  const depth = node.rawData.bitsPerChannel || 8
  const pixels = box.width * box.height
  context.budget.ensure(pixels * 4 * sampleBytes(depth) + (depth === 8 ? 0 : pixels * 4) + estimateRasterBytes(box.width, box.height))
  let data: PixelData | undefined
  try {
    data = context.ag.getLayerImageData(node)
  } catch (error) {
    throw friendlyReadError(error, context.budget.limit)
  }
  if (!data || data.width !== box.width || data.height !== box.height) return null
  let buffer: PixelBuffer = { width: box.width, height: box.height, data: toRgba8(data.data as PsdSampleArray, pixels) }
  let rect = box
  // Photoshop keeps pixels outside the canvas. They are invisible; a huge off-canvas layer is cropped
  // to the canvas so it cannot exhaust memory.
  if (pixels > context.options.maxPixels) {
    const visible = intersect(box, { x: 0, y: 0, width: context.width, height: context.height })
    if (!visible) return null
    buffer = cropPixels(buffer, box, visible)
    rect = visible
  }
  return { buffer, left: rect.x, top: rect.y }
}

function decodeMaskRecord(context: ImportContext, record: LayerMaskData, read: () => PixelData | undefined): DecodedMask {
  const box = boxOf(record)
  const defaultValue: 0 | 255 = (record.defaultColor ?? 0) >= 128 ? 255 : 0
  let buffer: MaskBuffer | null = null
  if (box.width && box.height) {
    context.budget.ensure(box.width * box.height * 8 + estimateRasterBytes(box.width, box.height) / 4)
    let data: PixelData | undefined
    try {
      data = read()
    } catch (error) {
      throw friendlyReadError(error, context.budget.limit)
    }
    if (data && data.width === box.width && data.height === box.height) buffer = maskChannel(data.data as PsdSampleArray, box.width, box.height)
  }
  return { buffer, left: box.x, top: box.y, defaultValue, disabled: Boolean(record.disabled) }
}

/** A decoded mask as a document-size coverage buffer (outside its rectangle: its default value). */
function documentMask(context: ImportContext, mask: DecodedMask): MaskBuffer {
  const out = createMaskBuffer(context.width, context.height, mask.defaultValue)
  if (mask.buffer) pasteMask(out, mask.buffer, mask.left, mask.top)
  return out
}

function layerMaskFromDecoded(context: ImportContext, mask: DecodedMask): LayerMask {
  const surface = mask.buffer
    ? maskFromBuffer(mask.buffer, mask.defaultValue, mask.left, mask.top)
    : maskFromBuffer({ width: 0, height: 0, data: new Uint8Array(0) }, mask.defaultValue)
  context.budget.add(surface.byteSize)
  return createLayerMask({ surface, enabled: !mask.disabled, linked: true })
}

/**
 * The layer's mask: its pixel mask, combined with its vector mask when it has one (Photoshop applies
 * both). A vector mask comes from the file's rendering of it when present, else from its paths.
 */
function decodeLayerMask(context: ImportContext, node: PsdLayer, name: string): LayerMask | null {
  const ag = context.ag
  let user: DecodedMask | null = null
  let vectorRecord: DecodedMask | null = null
  if (node.realMask) {
    // Both a pixel mask and a vector mask: the "real" mask is the pixel mask, the other the rendered vector.
    user = decodeMaskRecord(context, node.realMask, () => ag.getLayerRealMaskImageData(node))
    if (node.mask) vectorRecord = decodeMaskRecord(context, node.mask, () => ag.getLayerMaskImageData(node))
  } else if (node.mask) {
    const decoded = decodeMaskRecord(context, node.mask, () => ag.getLayerMaskImageData(node))
    if (node.mask.fromVectorData) vectorRecord = decoded
    else user = decoded
  }
  let vector: MaskBuffer | null = null
  let vectorDefault: 0 | 255 = 0
  const vectorMask = node.vectorMask
  if (vectorRecord && !(vectorMask && vectorMask.disable)) {
    vector = documentMask(context, vectorRecord)
    vectorDefault = vectorRecord.defaultValue
  } else if (vectorMask && !vectorMask.disable && vectorMask.paths?.length) {
    context.budget.ensure(context.width * context.height * 2)
    vector = rasterizeVectorMask(vectorMask, context.width, context.height)
    vectorDefault = Boolean(vectorMask.invert) !== Boolean(vectorMask.fillStartsWithAllPixels) ? 255 : 0
  }
  if (vectorMask || vectorRecord) {
    context.issues.add('vector-mask', name, vector ? 'The vector mask was converted to a pixel mask.' : 'The disabled vector mask was left out.')
  }
  if (!vector) return user ? layerMaskFromDecoded(context, user) : null
  let outside: 0 | 255 = vectorDefault
  if (user && !user.disabled) {
    multiplyMasks(vector, documentMask(context, user))
    outside = vectorDefault === 255 && user.defaultValue === 255 ? 255 : 0
  }
  const surface = maskFromBuffer(vector, outside)
  context.budget.add(surface.byteSize)
  return createLayerMask({ surface, enabled: true, linked: true })
}

interface LayerProps {
  readonly name: string
  readonly visible: boolean
  readonly opacity: number
  readonly blendMode: AdjustmentLayer['blendMode']
  readonly clipped: boolean
  readonly locks: AdjustmentLayer['locks']
}

function layerProps(context: ImportContext, node: PsdLayer, name: string, clippable: boolean): LayerProps {
  const blend = blendFromPsd(node.blendMode)
  if (!blend.supported) context.issues.add('unsupported-blend-mode', name, `The "${String(node.blendMode)}" blend mode is shown as Normal.`)
  const effects = hasEnabledEffects(node)
  if (effects) context.issues.add('layer-effects', name, 'Layer styles (shadows, strokes, glows, ...) were left out.')
  let opacity = Math.max(0, Math.min(1, node.opacity ?? 1))
  const fill = node.fillOpacity
  if (typeof fill === 'number' && Number.isFinite(fill) && fill < 1) {
    opacity *= Math.max(0, fill)
    if (effects || FILL_OPACITY_SENSITIVE_MODES.has(blend.mode)) {
      context.issues.add('fill-opacity', name, 'Fill opacity was combined with the layer opacity.')
    }
  }
  const advanced = advancedBlendingNotes(node)
  if (advanced.length) context.issues.add('knockout-or-advanced-blending', name, `Advanced blending (${advanced.join(', ')}) was left out.`)
  return {
    name,
    visible: !node.hidden,
    opacity: quantizeOpacity(opacity),
    blendMode: blend.mode,
    clipped: clippable && Boolean(node.clipping),
    locks: locksFromPsd(node),
  }
}

function withProps<T extends Layer>(layer: T, props: LayerProps, mask: LayerMask | null): T {
  return { ...layer, name: props.name, visible: props.visible, opacity: props.opacity, blendMode: props.blendMode, clipped: props.clipped, locks: props.locks, mask } as T
}

function rasterFromPixels(context: ImportContext, pixels: DecodedPixels, init: Partial<Parameters<typeof createRasterLayer>[0]> & { name: string }): RasterLayer {
  const surface = surfaceFromBuffer(pixels.buffer, pixels.left, pixels.top)
  context.budget.add(surface.byteSize)
  return createRasterLayer({ ...init, surface })
}

function emptyPixels(left: number, top: number): DecodedPixels {
  return { buffer: { width: 0, height: 0, data: new Uint8ClampedArray(0) }, left, top }
}

async function textRasterCache(context: ImportContext, spec: TextSpec, pixels: DecodedPixels | null): Promise<RasterCache> {
  const key = specKey(spec)
  if (pixels) {
    const cache = rasterCacheFrom(pixels.buffer, pixels.left, pixels.top, key)
    context.budget.add(cache.surface.byteSize)
    return cache
  }
  try {
    const rendered = await rasterizeText(spec)
    return rasterCacheFrom(rendered.pixels, rendered.offsetX, rendered.offsetY, key)
  } catch {
    // No canvas here (tests) or nothing to draw: an empty cache with the matching key.
    return rasterCacheFrom(emptyPixels(0, 0).buffer, 0, 0, key)
  }
}

/** One non-group PSD layer; null when it is left out (unsupported adjustment or fill). */
async function convertLayer(context: ImportContext, node: PsdLayer, rootBottom: boolean, clippable: boolean): Promise<Layer | null> {
  const name = String(node.name ?? '').trim() || 'Layer'
  try {
    const props = layerProps(context, node, name, clippable)
    if (node.placedLayer) context.issues.add('smart-object', name, 'The smart object was imported as pixels.')
    if (node.adjustment) {
      const spec = adjustmentFromPsd(node.adjustment)
      if (!spec) {
        context.issues.add('unsupported-adjustment', name, `${psdAdjustmentLabel(node.adjustment)} adjustment layers are not supported; the layer was left out.`)
        return null
      }
      const mask = decodeLayerMask(context, node, name)
      return withProps(createAdjustmentLayer(name, spec, mask), props, mask)
    }
    const hasTransparency = Boolean(node.rawData?.channels.some((channel) => channel.id === -1))
    const mask = decodeLayerMask(context, node, name)
    let pixels = decodePixels(context, node)
    if (node.text) {
      const { spec, notes } = textFromPsd(node.text, context.resolveFont, measureText)
      context.issues.add('text-rerender', name, notes.length
        ? `Simple redraws the text when it is edited (${notes.join('; ')}).`
        : 'Simple redraws the text when it is edited.')
      const raster = await textRasterCache(context, spec, pixels)
      return withProps(await createTextLayer(name, spec, raster), props, mask)
    }
    const fill = node.vectorFill
    if (!pixels && fill) {
      if (fill.type !== 'color') {
        context.issues.add('unsupported-adjustment', name, 'Gradient and pattern fill layers are not supported; the layer was left out.')
        return null
      }
      const color = psdColorToRgb(fill.color, 'descriptor') ?? { r: 0, g: 0, b: 0 }
      context.budget.ensure(context.width * context.height * 8)
      const data = new Uint8ClampedArray(context.width * context.height * 4)
      for (let i = 0; i < data.length; i += 4) {
        data[i] = color.r
        data[i + 1] = color.g
        data[i + 2] = color.b
        data[i + 3] = 255
      }
      pixels = { buffer: { width: context.width, height: context.height, data }, left: 0, top: 0 }
    }
    const coversCanvas = Boolean(pixels && pixels.left <= 0 && pixels.top <= 0
      && pixels.left + pixels.buffer.width >= context.width && pixels.top + pixels.buffer.height >= context.height)
    if (rootBottom && !hasTransparency && coversCanvas && !mask && props.opacity === 1 && props.blendMode === 'normal' && !props.clipped && pixels) {
      // Photoshop's Background: the bottom layer stored without a transparency channel.
      const visible = intersect({ x: pixels.left, y: pixels.top, width: pixels.buffer.width, height: pixels.buffer.height }, { x: 0, y: 0, width: context.width, height: context.height })
      const canvasPixels = visible && (visible.width !== pixels.buffer.width || visible.height !== pixels.buffer.height)
        ? { buffer: cropPixels(pixels.buffer, { x: pixels.left, y: pixels.top, width: pixels.buffer.width, height: pixels.buffer.height }, visible), left: 0, top: 0 }
        : pixels
      return rasterFromPixels(context, canvasPixels, { name, isBackground: true, visible: props.visible })
    }
    const layer = rasterFromPixels(context, pixels ?? emptyPixels(0, 0), { name })
    return withProps(layer, props, mask)
  } finally {
    delete node.rawData
  }
}

/** Layers whose isolation does not matter: normal pixel layers (a 'normal' group of them can dissolve). */
function isolationIrrelevant(layers: readonly Layer[]): boolean {
  return layers.every((layer) => layer.kind !== 'adjustment' && layer.blendMode === 'normal' && !layer.clipped)
}

function canvasRect(context: Pick<ImportContext, 'width' | 'height'>): IntRect {
  return { x: 0, y: 0, width: context.width, height: context.height }
}

/** Composites `layers` (as if visible as a group) into one raster surface; null when nothing shows. */
function compositeLayers(context: ImportContext, layers: readonly Layer[]): DecodedPixels | null {
  let bounds: IntRect | null = null
  for (const layer of layers) {
    if (layer.kind === 'adjustment' || !layer.visible) continue
    bounds = union(bounds, layerContentBounds(layer))
  }
  const rect = bounds ? intersect(bounds, canvasRect(context)) : null
  if (!rect) return null
  context.budget.ensure(rect.width * rect.height * 4 + estimateRasterBytes(rect.width, rect.height))
  const buffer = compositeRect({ width: context.width, height: context.height, layers }, rect)
  return { buffer, left: rect.x, top: rect.y }
}

function releaseLayers(context: ImportContext, layers: readonly Layer[]): void {
  for (const layer of layers) {
    if (layer.kind === 'raster') context.budget.used -= layer.surface.byteSize
    else if (layer.kind === 'text' || layer.kind === 'shape') context.budget.used -= layer.raster.surface.byteSize
    if (layer.mask) context.budget.used -= layer.mask.surface.byteSize
  }
  if (context.budget.used < 0) context.budget.used = 0
}

async function convertGroup(context: ImportContext, node: PsdLayer, clippedAbove: boolean, clippable: boolean): Promise<Layer[]> {
  const name = String(node.name ?? '').trim() || 'Group'
  const children = await convertChildren(context, node.children ?? [], false)
  if (!children.length) return []
  const passThrough = node.blendMode === undefined || node.blendMode === 'pass through'
  const opaque = (node.opacity ?? 1) >= 1 - 1e-9 && (node.fillOpacity ?? 1) >= 1 - 1e-9
  const masked = Boolean((node.mask && !node.mask.disabled) || node.realMask || (node.vectorMask && !node.vectorMask.disable))
  const dissolvable = opaque && !masked && !hasEnabledEffects(node) && !clippedAbove && !node.clipping
    && (passThrough || (node.blendMode === 'normal' && isolationIrrelevant(children)))
  if (dissolvable) {
    return children.map((layer, index) => ({
      ...layer,
      name: `${name} / ${layer.name}`,
      visible: layer.visible && !node.hidden,
      // Clipping never reaches out of a group: its bottom layer is not clipped.
      clipped: index === 0 ? false : layer.clipped,
    }))
  }
  context.issues.add('group-flattened', name, 'The layer group was merged into one layer.')
  const props = layerProps(context, node, name, clippable)
  const mask = decodeLayerMask(context, node, name)
  const pixels = compositeLayers(context, children)
  releaseLayers(context, children)
  const layer = rasterFromPixels(context, pixels ?? emptyPixels(0, 0), { name })
  return [withProps(layer, props, mask)]
}

/** PSD siblings (bottom to top) as Advanced layers (bottom to top). */
async function convertChildren(context: ImportContext, nodes: readonly PsdLayer[], root: boolean): Promise<Layer[]> {
  const out: Layer[] = []
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index]
    const clippable = out.length > 0
    if (node.children) {
      out.push(...await convertGroup(context, node, Boolean(nodes[index + 1]?.clipping), clippable))
    } else {
      const layer = await convertLayer(context, node, root && index === 0, clippable)
      if (layer) out.push(layer)
    }
    await yieldToEventLoop()
  }
  return out
}

/**
 * Keeps the layer count within LIMITS.maxLayers by merging the bottom layers into one. Exact: the
 * composite of the bottom layers is the backdrop every layer above sees.
 */
function enforceLayerLimit(context: ImportContext, layers: Layer[]): Layer[] {
  if (layers.length <= LIMITS.maxLayers) return layers
  let count = layers.length - LIMITS.maxLayers + 1
  while (count < layers.length && layers[count].clipped) count += 1
  const bottom = layers.slice(0, count)
  const pixels = compositeLayers(context, bottom)
  releaseLayers(context, bottom)
  context.issues.add('group-flattened', null, `The bottom ${count} layers were merged into one (Simple edits up to ${LIMITS.maxLayers} layers).`)
  const merged = rasterFromPixels(context, pixels ?? emptyPixels(0, 0), { name: `Merged layers (${count})` })
  return [merged, ...layers.slice(count)]
}

function containsText(nodes: readonly PsdLayer[]): boolean {
  return nodes.some((node) => Boolean(node.text) || (node.children ? containsText(node.children) : false))
}

function hasAnyTransparency(buffer: PixelBuffer): boolean {
  const data = buffer.data
  for (let i = 3; i < data.length; i += 4) if (data[i] < 255) return true
  return false
}

/** The merged image as the only layer: Background, or "Layer 0" when it has transparency. */
function layerFromComposite(context: ImportContext, composite: PixelBuffer): RasterLayer {
  const transparent = hasAnyTransparency(composite)
  return rasterFromPixels(context, { buffer: composite, left: 0, top: 0 }, transparent ? { name: 'Layer 0' } : { name: 'Background', isBackground: true })
}

function decodeComposite(context: ImportContext, psd: Psd, mergedAlpha: boolean, depth: number): PixelBuffer | null {
  if (!psd.rawCompositeData || !psd.rawCompositeData.length) return null
  if (psd.imageResources?.versionInfo?.hasRealMergedData === false) return null
  const pixels = context.width * context.height
  context.budget.ensure(pixels * 4 * sampleBytes(depth) + pixels * 4)
  let data: PixelData | undefined
  try {
    data = context.ag.getCompositeImageData(psd)
  } catch {
    return null
  }
  if (!data || data.width !== context.width || data.height !== context.height) return null
  // 8-bit data comes back as ag-psd's own array (made for this call), so fixing it in place is safe.
  const rgba = toRgba8(data.data as PsdSampleArray, pixels)
  if (mergedAlpha) removeWhiteMatte(rgba)
  else forceOpaque(rgba)
  context.budget.add(rgba.byteLength)
  return { width: context.width, height: context.height, data: rgba }
}

const DEFAULT_IMPORT_OPTIONS: PsdImportOptions = Object.freeze({
  maxPixels: LIMITS.maxPixels,
  maxDimension: LIMITS.maxDimension,
  memoryLimitBytes: LIMITS.psdMemoryLimitBytes,
  mode: 'layers',
})

function toUint8Array(bytes: Uint8Array | ArrayBuffer): Uint8Array {
  if (bytes instanceof Uint8Array) return bytes
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes)
  if (ArrayBuffer.isView(bytes)) {
    const view = bytes as ArrayBufferView
    return new Uint8Array(view.buffer, view.byteOffset, view.byteLength)
  }
  throw new TypeError('A Photoshop document must be given as bytes.')
}

/** Parses a PSD into an ImportedDocument (see the module comment). Throws friendly errors. */
export const importPsd: ImportPsd = async (bytes, options = {}) => {
  const opts: PsdImportOptions = { ...DEFAULT_IMPORT_OPTIONS, ...options }
  const view = toUint8Array(bytes)
  const header = sniffPsd(view)
  if (!header) throw new Error('This file is not a Photoshop document (.psd).')
  const problem = psdHeaderProblem(header, opts)
  if (problem) throw new Error(problem)
  const budget = new MemoryBudget(Math.max(1, opts.memoryLimitBytes))
  const issues = new IssueList()
  const structure = readPsdStructure(view)
  const ag = await loadAgPsd()
  // useRawData keeps views into the file, and ag-psd byte-swaps raw (uncompressed) 16- and 32-bit channel
  // data in place while decoding: work on a copy then, so the caller's bytes (kept for byte-exact unchanged
  // saves) never change. 8-bit data is only read.
  const source = header.depth > 8 ? view.slice() : view
  let psd: Psd
  try {
    psd = ag.readPsd(source, {
      useRawData: true,
      skipThumbnail: true,
      skipLinkedFilesData: true,
      totalMemoryLimit: opts.memoryLimitBytes,
    })
  } catch (error) {
    throw friendlyReadError(error, budget.limit)
  }
  const width = psd.width
  const height = psd.height
  const context: ImportContext = {
    ag,
    width,
    height,
    options: opts,
    budget,
    issues,
    resolveFont: opts.resolveFont ?? resolvePostScriptFont,
  }
  if (header.depth === 16 || header.depth === 32) {
    issues.add('bit-depth-reduced', null, `The document uses ${header.depth}-bit channels; Simple edits it with 8 bits per channel.`)
  }
  if (header.colorMode === 3 && structure.iccProfile && !isSrgbProfile(structure.iccProfile)) {
    const description = iccProfileDescription(structure.iccProfile) ?? 'A non-sRGB color profile'
    issues.add('color-profile', null, `${description}: colors are shown and saved as sRGB.`)
  }
  const ppi = ppiFromResolution(psd.imageResources?.resolutionInfo)
  try {
    const composite = decodeComposite(context, psd, structure.mergedAlpha, header.depth)
    let layers: Layer[]
    if (opts.mode === 'flattened') {
      if (composite) {
        layers = [layerFromComposite(context, composite)]
      } else {
        const all = await convertChildren(context, psd.children ?? [], true)
        const pixels = compositeLayers(context, all)
        releaseLayers(context, all)
        layers = [rasterFromPixels(context, pixels ?? emptyPixels(0, 0), { name: 'Layer 0' })]
      }
    } else {
      if (!opts.resolveFont && containsText(psd.children ?? [])) await loadLocalFonts(1500)
      layers = psd.children?.length ? await convertChildren(context, psd.children, true) : []
      if (!layers.length) {
        layers = composite
          ? [layerFromComposite(context, composite)]
          : [rasterFromPixels(context, emptyPixels(0, 0), { name: 'Layer 0' })]
      }
      layers = enforceLayerLimit(context, layers)
    }
    return { width, height, ppi, layers, composite, issues: issues.items }
  } catch (error) {
    throw friendlyReadError(error, budget.limit)
  }
}

// ---------------------------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------------------------

/** Photoshop's per-layer limit for PSD (PSB allows more). */
const PSD_MAX_SIDE = 30_000

interface ExportedPixels {
  readonly left: number
  readonly top: number
  readonly imageData: PixelData
}

function pixelsAt(surface: RasterLayer['surface'], local: IntRect, offsetX: number, offsetY: number, canvas: IntRect): ExportedPixels | null {
  let rect: IntRect | null = { x: local.x + offsetX, y: local.y + offsetY, width: local.width, height: local.height }
  if (rect.width > PSD_MAX_SIDE || rect.height > PSD_MAX_SIDE) rect = intersect(rect, canvas)
  if (!rect || !rect.width || !rect.height) return null
  const buffer = surface.read({ x: rect.x - offsetX, y: rect.y - offsetY, width: rect.width, height: rect.height })
  return { left: rect.x, top: rect.y, imageData: { width: buffer.width, height: buffer.height, data: buffer.data } }
}

function maskToPsd(mask: LayerMask, canvas: IntRect): LayerMaskData {
  const defaultColor = mask.surface.defaultValue
  let bounds = mask.surface.contentBounds()
  let rect: IntRect | null = bounds ? { x: bounds.x + mask.offsetX, y: bounds.y + mask.offsetY, width: bounds.width, height: bounds.height } : null
  if (rect && (rect.width > PSD_MAX_SIDE || rect.height > PSD_MAX_SIDE)) rect = intersect(rect, canvas)
  if (!rect) return { left: 0, top: 0, right: 0, bottom: 0, defaultColor, disabled: !mask.enabled }
  bounds = { x: rect.x - mask.offsetX, y: rect.y - mask.offsetY, width: rect.width, height: rect.height }
  const coverage = mask.surface.read(bounds)
  const data = new Uint8ClampedArray(rect.width * rect.height * 4)
  for (let p = 0, i = 0; p < coverage.data.length; p += 1, i += 4) {
    const value = coverage.data[p]
    data[i] = value
    data[i + 1] = value
    data[i + 2] = value
    data[i + 3] = 255
  }
  return {
    left: rect.x,
    top: rect.y,
    right: rect.x + rect.width,
    bottom: rect.y + rect.height,
    defaultColor,
    disabled: !mask.enabled,
    imageData: { width: rect.width, height: rect.height, data },
  }
}

function commonToPsd(layer: Layer, clippable: boolean): PsdLayer {
  return {
    name: layer.name,
    hidden: !layer.visible,
    opacity: quantizeOpacity(layer.opacity),
    blendMode: blendToPsd(layer.blendMode) as PsdLayer['blendMode'],
    clipping: clippable && layer.clipped,
    ...locksToPsd(layer.locks),
  }
}

function withPixels(target: PsdLayer, pixels: ExportedPixels | null): PsdLayer {
  if (!pixels) return target
  return {
    ...target,
    left: pixels.left,
    top: pixels.top,
    right: pixels.left + pixels.imageData.width,
    bottom: pixels.top + pixels.imageData.height,
    imageData: pixels.imageData,
  }
}

type ExportDocument = Pick<DocumentState, 'width' | 'height' | 'ppi' | 'layers'>

function layerToPsd(doc: ExportDocument, index: number, canvas: IntRect): PsdLayer {
  const layer = doc.layers[index]
  const base = commonToPsd(layer, index > 0)
  if (layer.mask) base.mask = maskToPsd(layer.mask, canvas)
  if (layer.kind === 'raster') {
    if (layer.isBackground && index === 0) {
      // Background: the whole canvas, opaque, so Photoshop opens it as its Background layer.
      return withPixels(base, pixelsAt(layer.surface, { x: -layer.offsetX, y: -layer.offsetY, width: doc.width, height: doc.height }, layer.offsetX, layer.offsetY, canvas))
    }
    const bounds = layer.surface.contentBounds()
    return withPixels(base, bounds ? pixelsAt(layer.surface, bounds, layer.offsetX, layer.offsetY, canvas) : null)
  }
  if (layer.kind === 'text' || layer.kind === 'shape') {
    const bounds = layer.raster.surface.contentBounds()
    const target = withPixels(base, bounds ? pixelsAt(layer.raster.surface, bounds, layer.raster.offsetX, layer.raster.offsetY, canvas) : null)
    if (layer.kind === 'text') target.text = textToPsd(layer.text, postScriptNameFor(layer.text.style), measureText)
    return target
  }
  const adjustment = adjustmentToPsd(layer.adjustment)
  if (adjustment) return { ...base, adjustment: adjustment as PsdLayer['adjustment'] }
  // No Photoshop equivalent (Simple's quick adjust): keep the exact look as a pixel layer holding the
  // composite up to and including this layer. The layers below stay in the file.
  if (!layer.visible) return { ...base, name: layer.name }
  const stamp = compositeRect({ width: doc.width, height: doc.height, layers: doc.layers.slice(0, index + 1) }, canvas)
  const stamped: PsdLayer = { ...base, name: `${layer.name} (merged)`, opacity: 1, blendMode: 'normal' }
  delete stamped.mask
  return withPixels(stamped, { left: 0, top: 0, imageData: { width: stamp.width, height: stamp.height, data: stamp.data } })
}

function tightCopy(bytes: Uint8Array): Uint8Array {
  return bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength ? bytes : bytes.slice()
}

/** Writes the document as a layered 8-bit RGB PSD with `composite` (document size) as the merged image. */
export const exportPsd: ExportPsd = async (doc, composite, options = {}) => {
  const width = Math.round(doc.width)
  const height = Math.round(doc.height)
  if (!(width > 0 && height > 0) || width > LIMITS.maxDimension || height > LIMITS.maxDimension || width * height > LIMITS.maxPixels) {
    throw new Error('This document is too large to save as a Photoshop document.')
  }
  if (!composite || composite.width !== width || composite.height !== height || composite.data.length !== width * height * 4) {
    throw new Error('The merged image does not match the document size.')
  }
  const exportOptions: PsdExportOptions = { invalidateText: true, ...options }
  const ag = await loadAgPsd()
  const canvas: IntRect = { x: 0, y: 0, width, height }
  const children: PsdLayer[] = []
  for (let index = 0; index < doc.layers.length; index += 1) {
    children.push(layerToPsd(doc, index, canvas))
    if (index % 8 === 7) await yieldToEventLoop()
  }
  const first = doc.layers[0]
  const hasBackground = Boolean(first && first.kind === 'raster' && first.isBackground)
  const psd: Psd = {
    width,
    height,
    imageData: { width, height, data: composite.data },
    imageResources: {
      resolutionInfo: resolutionFromPpi(doc.ppi),
      versionInfo: { hasRealMergedData: true, writerName: 'Simple Image', readerName: 'Simple Image', fileVersion: 1 },
    },
    children,
  }
  let bytes: Uint8Array
  try {
    bytes = ag.writePsdUint8Array(psd, { invalidateTextLayers: exportOptions.invalidateText !== false, noBackground: !hasBackground })
  } catch (error) {
    if (error instanceof RangeError) throw new Error('There isn\'t enough memory to save this Photoshop document. Merge or delete some layers, then try again.')
    throw new Error(`The Photoshop document could not be written (${error instanceof Error ? error.message : String(error)}).`)
  }
  return tightCopy(bytes)
}

/** The merged image exportPsd needs, for callers without a display compositor (tests, QA, workers). */
export function flattenForExport(doc: Pick<DocumentState, 'width' | 'height' | 'layers'>, options?: OpOptions): PixelBuffer {
  return flattenDocument(doc, options)
}
