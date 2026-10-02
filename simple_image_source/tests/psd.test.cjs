'use strict'
// WP7 layered PSD import/export (src/advanced/psd.ts) in Node: Photoshop documents are generated with
// ag-psd's writer (and by hand for 16/32-bit, which ag-psd cannot write), opened with importPsd, saved
// with exportPsd and read back both by ag-psd itself and by importPsd. Layer count, names, order,
// opacity, blend modes, visibility, clipping, masks (default colour and pixels), adjustment parameters
// and text (string, font, transform) must survive; groups dissolve or flatten; memory limits, refused
// colour modes and damaged files give friendly errors; the caller's bytes are never modified.
const test = require('node:test')
const assert = require('node:assert/strict')
const path = require('node:path')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping (not disabled by --no-experimental-strip-types), '
  + 'or run node with --experimental-strip-types.'

function load(relative) {
  if (!(process.features && process.features.typescript)) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

const ag = require('ag-psd')
const psd = load('advanced/psd.ts')
const documentModule = load('advanced/document.ts')
const tiles = load('advanced/tiles.ts')
const { compositeRect, flattenDocument } = load('advanced/composite.ts')
const { LIMITS } = load('advanced/types.ts')

psd.configurePsdCanvas('node-test')
// The fixtures below are written by ag-psd in this process too; it needs the same canvas-free hooks.
ag.initializeCanvas(() => { throw new Error('no canvas in tests') }, (width, height) => ({ width, height, data: new Uint8ClampedArray(width * height * 4) }))

// #region fixtures

function solid(width, height, rgba) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < data.length; i += 4) data.set(rgba, i)
  return { width, height, data }
}

function pattern(width, height, seed) {
  const data = new Uint8ClampedArray(width * height * 4)
  let s = seed >>> 0
  for (let i = 0; i < data.length; i += 4) {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0
    data[i] = s & 255
    data[i + 1] = (s >>> 8) & 255
    data[i + 2] = (s >>> 16) & 255
    data[i + 3] = 40 + ((s >>> 24) % 216)
  }
  return { width, height, data }
}

function grayMask(width, height, value) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let p = 0; p < width * height; p += 1) {
    const v = typeof value === 'function' ? value(p % width, Math.floor(p / width)) : value
    data.set([v, v, v, 255], p * 4)
  }
  return { width, height, data }
}

const W = 96
const H = 64

/** The acceptance fixture: Background, masked multiply layer, clipped, hidden, text, adjustments, groups. */
function richPsd() {
  return ag.writePsdUint8Array({
    width: W,
    height: H,
    imageData: solid(W, H, [30, 60, 90, 255]),
    imageResources: { resolutionInfo: { horizontalResolution: 150, horizontalResolutionUnit: 'PPI', widthUnit: 'Inches', verticalResolution: 150, verticalResolutionUnit: 'PPI', heightUnit: 'Inches' } },
    children: [
      { name: 'Background', left: 0, top: 0, imageData: solid(W, H, [30, 60, 90, 255]) },
      {
        name: 'Red half', left: 8, top: 4, blendMode: 'multiply', opacity: 0.5,
        imageData: pattern(40, 30, 7),
        mask: { left: 8, top: 4, right: 48, bottom: 34, defaultColor: 0, imageData: grayMask(40, 30, (x, y) => (x + y) * 3 % 256) },
      },
      { name: 'Clipped', left: 10, top: 6, clipping: true, imageData: solid(12, 10, [0, 0, 255, 255]) },
      { name: 'Hidden', left: 50, top: 20, hidden: true, imageData: solid(9, 7, [0, 255, 0, 255]) },
      {
        name: 'Title', left: 2, top: 40, imageData: solid(30, 12, [200, 10, 20, 255]),
        text: { text: 'Hello', transform: [1, 0, 0, 1, 2, 50], style: { font: { name: 'ArialMT' }, fontSize: 12, fillColor: { r: 200, g: 10, b: 20 } } },
      },
      { name: 'Levels 1', opacity: 0.8, adjustment: { type: 'levels', rgb: { shadowInput: 10, highlightInput: 240, shadowOutput: 0, highlightOutput: 255, midtoneInput: 1.2 } } },
      {
        name: 'Curves 1', mask: { left: 0, top: 0, right: W, bottom: 32, defaultColor: 255, imageData: grayMask(W, 32, 0) },
        adjustment: { type: 'curves', rgb: [{ input: 0, output: 0 }, { input: 128, output: 160 }, { input: 255, output: 255 }] },
      },
      { name: 'Hue/Saturation 1', hidden: true, adjustment: { type: 'hue/saturation', master: { a: 0, b: 0, c: 0, d: 0, hue: 20, saturation: -30, lightness: 5 } } },
      { name: 'Group', opened: true, blendMode: 'pass through', children: [{ name: 'Inner', left: 60, top: 40, imageData: solid(8, 8, [0, 255, 0, 255]) }] },
      {
        name: 'Mul group', opened: true, blendMode: 'multiply', opacity: 0.75,
        children: [
          { name: 'A', left: 70, top: 10, imageData: solid(10, 10, [0, 255, 0, 255]) },
          { name: 'B', left: 75, top: 15, blendMode: 'screen', imageData: solid(10, 10, [255, 0, 0, 200]) },
        ],
      },
    ],
  }, { invalidateTextLayers: true })
}

function be(bytes, write) {
  const buffer = Buffer.alloc(bytes)
  write(buffer)
  return buffer
}
const u16 = (value) => be(2, (b) => b.writeUInt16BE(value, 0))
const i16 = (value) => be(2, (b) => b.writeInt16BE(value, 0))
const u32 = (value) => be(4, (b) => b.writeUInt32BE(value >>> 0, 0))
const i32 = (value) => be(4, (b) => b.writeInt32BE(value, 0))

function pascal4(text) {
  const raw = Buffer.from(text, 'latin1')
  const out = Buffer.alloc(Math.ceil((raw.length + 1) / 4) * 4)
  out[0] = raw.length
  raw.copy(out, 1)
  return out
}

/**
 * A raw-compressed RGB PSD at 8, 16 or 32 bits written by hand. layers: { name, left, top, width,
 * height, channels: { '-1'?: number[], 0: number[], 1: number[], 2: number[] } }; composite: planes.
 */
function deepPsd({ width, height, depth, layers = [], composite, mergedAlpha = false, colorMode = 3 }) {
  const size = depth / 8
  const plane = (values) => {
    const buffer = Buffer.alloc(values.length * size)
    values.forEach((value, index) => {
      if (depth === 16) buffer.writeUInt16BE(value, index * 2)
      else if (depth === 32) buffer.writeFloatBE(value, index * 4)
      else buffer[index] = value
    })
    return buffer
  }
  const records = []
  const images = []
  for (const layer of layers) {
    const ids = Object.keys(layer.channels).map(Number).sort((a, b) => a - b)
    const data = ids.map((id) => Buffer.concat([u16(0), plane(layer.channels[id])]))
    const extra = Buffer.concat([u32(0), u32(0), pascal4(layer.name)])
    records.push(Buffer.concat([
      i32(layer.top), i32(layer.left), i32(layer.top + layer.height), i32(layer.left + layer.width),
      u16(ids.length), ...ids.map((id, index) => Buffer.concat([i16(id), u32(data[index].length)])),
      Buffer.from('8BIMnorm', 'latin1'), Buffer.from([255, 0, 8, 0]), u32(extra.length), extra,
    ]))
    images.push(...data)
  }
  let layerInfo = Buffer.concat([i16(mergedAlpha ? -layers.length : layers.length), ...records, ...images])
  if (layerInfo.length % 2) layerInfo = Buffer.concat([layerInfo, Buffer.alloc(1)])
  const layerAndMask = Buffer.concat([u32(layerInfo.length), layerInfo, u32(0)])
  const head = Buffer.concat([Buffer.from('8BPS', 'latin1'), u16(1), Buffer.alloc(6), u16(composite.length), u32(height), u32(width), u16(depth), u16(colorMode)])
  return new Uint8Array(Buffer.concat([head, u32(0), u32(0), u32(layerAndMask.length), layerAndMask, u16(0), ...composite.map(plane)]))
}

// #endregion

// #region helpers

function surfacePixels(layer) {
  const surface = layer.kind === 'raster' ? layer.surface : layer.raster.surface
  const bounds = surface.contentBounds()
  if (!bounds) return null
  const offsetX = layer.kind === 'raster' ? layer.offsetX : layer.raster.offsetX
  const offsetY = layer.kind === 'raster' ? layer.offsetY : layer.raster.offsetY
  return { x: bounds.x + offsetX, y: bounds.y + offsetY, pixels: Buffer.from(surface.read(bounds).data) }
}

function maskPixels(mask) {
  const bounds = mask.surface.contentBounds()
  return {
    defaultValue: mask.surface.defaultValue,
    enabled: mask.enabled,
    rect: bounds ? { x: bounds.x + mask.offsetX, y: bounds.y + mask.offsetY, width: bounds.width, height: bounds.height } : null,
    data: bounds ? Buffer.from(mask.surface.read(bounds).data) : null,
  }
}

function summary(layer) {
  return {
    kind: layer.kind,
    name: layer.name,
    visible: layer.visible,
    opacity: layer.opacity,
    blendMode: layer.blendMode,
    clipped: layer.clipped,
    isBackground: layer.kind === 'raster' ? layer.isBackground : undefined,
    adjustment: layer.kind === 'adjustment' ? layer.adjustment : undefined,
    text: layer.kind === 'text' ? { text: layer.text.text, family: layer.text.style.fontFamily, size: layer.text.style.fontSize, color: layer.text.style.color, transform: layer.text.transform } : undefined,
  }
}

function assertSameDocument(actual, expected) {
  assert.equal(actual.width, expected.width)
  assert.equal(actual.height, expected.height)
  assert.equal(actual.ppi, expected.ppi)
  assert.deepEqual(actual.layers.map((layer) => layer.name), expected.layers.map((layer) => layer.name), 'names and order')
  for (let index = 0; index < expected.layers.length; index += 1) {
    const a = actual.layers[index]
    const e = expected.layers[index]
    const sa = summary(a)
    const se = summary(e)
    assert.ok(Math.abs(sa.opacity - se.opacity) <= 1 / 255, `${e.name} opacity`)
    assert.deepEqual({ ...sa, opacity: 0 }, { ...se, opacity: 0 }, `${e.name} properties`)
    if (e.kind !== 'adjustment') assert.deepEqual(surfacePixels(a), surfacePixels(e), `${e.name} pixels`)
    assert.equal(Boolean(a.mask), Boolean(e.mask), `${e.name} mask presence`)
    if (e.mask) assert.deepEqual(maskPixels(a.mask), maskPixels(e.mask), `${e.name} mask`)
  }
}

function hashBuffer(buffer) {
  return require('node:crypto').createHash('sha256').update(buffer.data).digest('hex')
}

// #endregion

test('a layered Photoshop document imports with every supported property', async () => {
  const bytes = richPsd()
  const before = Buffer.from(bytes)
  const doc = await psd.importPsd(bytes)
  assert.equal(Buffer.from(bytes).equals(before), true, 'the caller\'s bytes are untouched')
  assert.equal(doc.width, W)
  assert.equal(doc.height, H)
  assert.equal(doc.ppi, 150)
  assert.ok(doc.composite && doc.composite.width === W && doc.composite.height === H)
  assert.deepEqual(Array.from(doc.composite.data.subarray(0, 4)), [30, 60, 90, 255])
  assert.deepEqual(doc.layers.map((layer) => [layer.kind, layer.name]), [
    ['raster', 'Background'],
    ['raster', 'Red half'],
    ['raster', 'Clipped'],
    ['raster', 'Hidden'],
    ['text', 'Title'],
    ['adjustment', 'Levels 1'],
    ['adjustment', 'Curves 1'],
    ['adjustment', 'Hue/Saturation 1'],
    ['raster', 'Group / Inner'],
    ['raster', 'Mul group'],
  ])
  const [background, red, clipped, hidden, title, levels, curves, hue, inner, flattened] = doc.layers
  assert.equal(background.isBackground, true)
  assert.deepEqual(background.locks, { pixels: false, position: true, transparency: true })
  assert.equal(red.blendMode, 'multiply')
  assert.equal(red.opacity, 128 / 255)
  assert.deepEqual(surfacePixels(red), { x: 8, y: 4, pixels: Buffer.from(pattern(40, 30, 7).data) })
  assert.equal(red.mask.surface.defaultValue, 0)
  assert.equal(red.mask.enabled, true)
  assert.equal(red.mask.surface.read({ x: 8 + 5, y: 4 + 7, width: 1, height: 1 }).data[0], (5 + 7) * 3)
  assert.equal(red.mask.surface.read({ x: 0, y: 0, width: 1, height: 1 }).data[0], 0, 'outside the mask rectangle: default colour')
  assert.equal(clipped.clipped, true)
  assert.equal(hidden.visible, false)
  assert.equal(title.text.text, 'Hello')
  assert.equal(title.text.style.fontFamily, 'Arial')
  assert.equal(title.text.style.fontSize, 12)
  assert.deepEqual(title.text.style.color, { r: 200, g: 10, b: 20 })
  assert.equal(title.raster.specKey !== '', true)
  assert.deepEqual(surfacePixels(title), { x: 2, y: 40, pixels: Buffer.from(solid(30, 12, [200, 10, 20, 255]).data) }, 'Photoshop\'s own text pixels are kept')
  assert.equal(levels.adjustment.type, 'levels')
  assert.deepEqual(levels.adjustment.rgb, { inBlack: 10, inWhite: 240, gamma: 1.2, outBlack: 0, outWhite: 255 })
  assert.equal(levels.opacity, 204 / 255)
  assert.deepEqual(curves.adjustment.rgb, [{ x: 0, y: 0 }, { x: 128, y: 160 }, { x: 255, y: 255 }])
  assert.equal(curves.mask.surface.defaultValue, 255)
  assert.deepEqual(curves.mask.surface.contentBounds(), { x: 0, y: 0, width: W, height: 32 })
  assert.equal(hue.visible, false)
  assert.deepEqual(hue.adjustment.master, { hue: 20, saturation: -30, lightness: 5 })
  assert.equal(inner.blendMode, 'normal')
  assert.equal(flattened.blendMode, 'multiply')
  assert.equal(flattened.opacity, 191 / 255)
  // The flattened group holds exactly what its children composite to.
  const children = await psd.importPsd(ag.writePsdUint8Array({
    width: W, height: H, children: [
      { name: 'A', left: 70, top: 10, imageData: solid(10, 10, [0, 255, 0, 255]) },
      { name: 'B', left: 75, top: 15, blendMode: 'screen', imageData: solid(10, 10, [255, 0, 0, 200]) },
    ],
  }, { noBackground: true }))
  const expected = compositeRect({ width: W, height: H, layers: children.layers }, { x: 70, y: 10, width: 15, height: 15 })
  assert.deepEqual(surfacePixels(flattened), { x: 70, y: 10, pixels: Buffer.from(expected.data) })
  const codes = doc.issues.map((issue) => `${issue.code}:${issue.layerName}`)
  assert.deepEqual(codes, ['text-rerender:Title', 'group-flattened:Mul group'])
  // The editor's document store accepts the imported layers as they are.
  const store = documentModule.createDocumentStore({ width: doc.width, height: doc.height, ppi: doc.ppi, layers: doc.layers, host: { nextRevision: () => 1, setRevision: () => {} }, baseLabel: 'Open' })
  assert.equal(store.getState().layers.length, doc.layers.length)
  assert.equal(store.getState().activeLayerId, doc.layers[doc.layers.length - 1].id)
  store.dispose()
})

test('PSD round trip: export then import keeps layers, masks, adjustments and text', async () => {
  const doc = await psd.importPsd(richPsd())
  const composite = psd.flattenForExport(doc)
  const bytes = await psd.exportPsd(doc, composite)
  assert.ok(bytes instanceof Uint8Array)
  assert.equal(bytes.byteLength, bytes.buffer.byteLength, 'a tight buffer (IPC sends the whole ArrayBuffer)')
  assert.equal(Buffer.from(bytes.subarray(0, 6)).toString('latin1'), '8BPS\u0000\u0001')

  // Independent reader: what Photoshop (or any PSD reader) sees.
  const raw = ag.readPsd(bytes, { useImageData: true, skipThumbnail: true })
  assert.equal(raw.width, W)
  assert.equal(raw.bitsPerChannel, 8)
  assert.equal(raw.imageResources.resolutionInfo.horizontalResolution, 150)
  assert.equal(raw.imageResources.versionInfo.hasRealMergedData, true)
  assert.deepEqual(Buffer.from(raw.imageData.data), Buffer.from(composite.data), 'the merged image is the given composite')
  assert.deepEqual(raw.children.map((layer) => layer.name), doc.layers.map((layer) => layer.name))
  const byName = Object.fromEntries(raw.children.map((layer) => [layer.name, layer]))
  assert.equal(byName['Red half'].blendMode, 'multiply')
  assert.equal(byName['Red half'].opacity, 128 / 255)
  assert.equal(byName['Red half'].mask.defaultColor, 0)
  assert.equal(byName.Clipped.clipping, true)
  assert.equal(byName.Hidden.hidden, true)
  assert.equal(byName.Title.text.text, 'Hello')
  assert.equal(byName.Title.text.style.font.name, 'ArialMT')
  assert.deepEqual(byName.Title.text.transform, [1, 0, 0, 1, 2, 50], 'point text anchored at its baseline again')
  assert.equal(byName['Levels 1'].adjustment.type, 'levels')
  assert.equal(byName['Curves 1'].mask.defaultColor, 255)
  assert.equal(byName['Hue/Saturation 1'].adjustment.type, 'hue/saturation')
  assert.equal(byName.Background.left, 0)
  assert.equal(byName.Background.right, W)

  // Our own reader: the document comes back the same.
  const back = await psd.importPsd(bytes)
  assertSameDocument(back, doc)
  assert.deepEqual(back.issues.map((issue) => issue.code), ['text-rerender'], 'nothing new is lost')
  assert.equal(hashBuffer(flattenDocument(back)), hashBuffer(flattenDocument(doc)), 'the composite is identical')

  // Saving the re-opened document again writes the same bytes (the writer is deterministic).
  const again = await psd.exportPsd(back, psd.flattenForExport(back))
  assert.equal(Buffer.from(again).equals(Buffer.from(bytes)), true)
})

test('documents built in the editor export with layer properties Photoshop understands', async () => {
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(solid(40, 30, [200, 200, 200, 255])) })
  const maskBuffer = { width: 10, height: 10, data: new Uint8Array(100).fill(128) }
  const red = documentModule.createRasterLayer({
    name: 'Red', surface: tiles.surfaceFromBuffer(solid(10, 10, [255, 0, 0, 255])), offsetX: 5, offsetY: 6, opacity: 0.4, blendMode: 'screen',
    locks: { pixels: true, position: false, transparency: true },
    mask: documentModule.createLayerMask({ surface: tiles.maskFromBuffer(maskBuffer, 255, 0, 0), offsetX: 5, offsetY: 6, enabled: false }),
  })
  const offCanvas = documentModule.createRasterLayer({ name: 'Off canvas', surface: tiles.surfaceFromBuffer(solid(20, 4, [0, 0, 0, 255]), -10, 28) })
  const empty = documentModule.createRasterLayer({ name: 'Empty' })
  const curves = documentModule.createAdjustmentLayer('Curves 1', { type: 'curves', rgb: [{ x: 0, y: 0 }, { x: 128, y: 160 }, { x: 255, y: 255 }], red: [{ x: 0, y: 0 }, { x: 255, y: 255 }], green: [{ x: 0, y: 0 }, { x: 255, y: 255 }], blue: [{ x: 0, y: 0 }, { x: 255, y: 255 }] })
  const state = { width: 40, height: 30, ppi: 300, layers: [background, red, offCanvas, empty, curves] }
  const bytes = await psd.exportPsd(state, psd.flattenForExport(state))
  const raw = ag.readPsd(bytes, { useImageData: true, skipThumbnail: true })
  assert.equal(raw.imageResources.resolutionInfo.horizontalResolution, 300)
  const [rawBackground, rawRed, rawOff, rawEmpty, rawCurves] = raw.children
  assert.deepEqual([rawBackground.left, rawBackground.top, rawBackground.right, rawBackground.bottom], [0, 0, 40, 30])
  assert.deepEqual([rawRed.left, rawRed.top, rawRed.right, rawRed.bottom], [5, 6, 15, 16], 'tight bounds at the layer offset')
  assert.equal(rawRed.blendMode, 'screen')
  assert.equal(rawRed.opacity, 102 / 255)
  assert.deepEqual(rawRed.protected, { transparency: true, composite: true, position: false })
  assert.equal(rawRed.mask.disabled, true)
  assert.equal(rawRed.mask.defaultColor, 255)
  assert.deepEqual([rawRed.mask.left, rawRed.mask.top], [5, 6])
  assert.deepEqual([rawOff.left, rawOff.top, rawOff.right, rawOff.bottom], [-10, 28, 10, 32], 'pixels outside the canvas are kept')
  assert.equal(rawEmpty.imageData, undefined)
  assert.equal(rawCurves.adjustment.type, 'curves')
  const back = await psd.importPsd(bytes)
  assert.equal(back.layers[0].isBackground, true)
  assert.deepEqual(back.layers[1].locks, { pixels: true, position: false, transparency: true })
  assert.equal(back.layers[1].mask.enabled, false)
  assert.deepEqual(surfacePixels(back.layers[2]), surfacePixels(offCanvas))
  assert.equal(back.layers[3].surface.contentBounds(), null)

  // Without a Background the bottom layer stays a normal (transparent-capable) layer.
  const floating = documentModule.createRasterLayer({ name: 'Layer 0', surface: tiles.surfaceFromBuffer(solid(40, 30, [9, 9, 9, 255])) })
  const plain = { width: 40, height: 30, ppi: 72, layers: [floating] }
  const reopened = await psd.importPsd(await psd.exportPsd(plain, psd.flattenForExport(plain)))
  assert.equal(reopened.layers[0].isBackground, false)
  assert.equal(reopened.layers[0].name, 'Layer 0')
})

test('an adjustment Photoshop cannot store is saved as its exact look, and text keeps its font', async () => {
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(solid(20, 10, [100, 100, 100, 255])) })
  const quick = documentModule.createAdjustmentLayer('Adjust 1', { type: 'quick', exposure: 30, brightness: 0, contrast: 0, highlights: 0, shadows: 0, saturation: 0, warmth: 0, auto: null })
  const spec = { text: 'Bold', style: { fontFamily: 'Times New Roman', fontSize: 20, fontWeight: 700, italic: true, underline: false, color: { r: 1, g: 2, b: 3 }, align: 'center', lineHeight: 1.2, letterSpacing: 0 }, boxWidth: null, transform: [1, 0, 0, 1, 3, 4] }
  const text = await documentModule.createTextLayer('Bold', spec, documentModule.rasterCacheFrom(solid(4, 4, [1, 2, 3, 255]), 3, 4, 'k'))
  const state = { width: 20, height: 10, ppi: 72, layers: [background, quick, text] }
  const composite = psd.flattenForExport(state)
  const bytes = await psd.exportPsd(state, composite)
  const raw = ag.readPsd(bytes, { useImageData: true, skipThumbnail: true })
  assert.equal(raw.children[1].name, 'Adjust 1 (merged)')
  assert.equal(raw.children[1].adjustment, undefined)
  const stamp = compositeRect({ width: 20, height: 10, layers: [background, quick] }, { x: 0, y: 0, width: 20, height: 10 })
  assert.deepEqual(Buffer.from(raw.children[1].imageData.data), Buffer.from(stamp.data))
  assert.equal(raw.children[2].text.style.font.name, 'TimesNewRomanPS-BoldItalicMT')
  const back = await psd.importPsd(bytes)
  assert.equal(back.layers[2].text.style.fontFamily, 'Times New Roman')
  assert.equal(back.layers[2].text.style.fontWeight, 700)
  assert.equal(back.layers[2].text.style.italic, true)
  assert.equal(back.layers[2].text.style.align, 'center')
  for (let i = 0; i < 6; i += 1) assert.ok(Math.abs(back.layers[2].text.transform[i] - spec.transform[i]) < 1e-6)
  assert.equal(hashBuffer(flattenDocument(back)), hashBuffer(composite), 'the look survives')
})

test('16-bit documents are reduced to 8-bit exactly and say so, and the original bytes stay unchanged', async () => {
  const values = [65535, 257 * 128, 0, 65535, 257 * 10, 257 * 20, 257 * 30, 32768]
  const bytes = deepPsd({
    width: 2,
    height: 1,
    depth: 16,
    layers: [{ name: 'Deep', left: 0, top: 0, width: 2, height: 1, channels: { '-1': [values[3], values[7]], 0: [values[0], values[4]], 1: [values[1], values[5]], 2: [values[2], values[6]] } }],
    composite: [[65535, 257 * 10], [257 * 128, 257 * 20], [0, 257 * 30]],
  })
  assert.equal(require('../electron/image-files.cjs').inspectImageBytes(Buffer.from(bytes), 'deep.psd').format, 'psd')
  const before = Buffer.from(bytes)
  const doc = await psd.importPsd(bytes)
  assert.equal(Buffer.from(bytes).equals(before), true, 'ag-psd byte-swaps raw 16-bit data in place; the import works on a copy')
  assert.deepEqual(Array.from(doc.layers[0].surface.read({ x: 0, y: 0, width: 2, height: 1 }).data), [255, 128, 0, 255, 10, 20, 30, 128])
  assert.deepEqual(Array.from(doc.composite.data), [255, 128, 0, 255, 10, 20, 30, 255])
  assert.deepEqual(doc.issues.map((issue) => issue.code), ['bit-depth-reduced'])
  // Saving writes an 8-bit document.
  const raw = ag.readPsd(await psd.exportPsd(doc, doc.composite), { useImageData: true, skipThumbnail: true })
  assert.equal(raw.bitsPerChannel, 8)
})

test('32-bit documents are converted from linear light, and the original bytes stay unchanged', async () => {
  const bytes = deepPsd({
    width: 2,
    height: 1,
    depth: 32,
    layers: [{ name: 'HDR', left: 0, top: 0, width: 2, height: 1, channels: { '-1': [1, 0.5], 0: [1, 2], 1: [0.21586, 0], 2: [0, 0.0031308] } }],
    composite: [[1, 2], [0.21586, 0], [0, 0.0031308]],
  })
  const before = Buffer.from(bytes)
  const doc = await psd.importPsd(bytes)
  assert.equal(Buffer.from(bytes).equals(before), true, 'ag-psd byte-swaps raw 32-bit data in place; the import works on a copy')
  assert.deepEqual(Array.from(doc.layers[0].surface.read({ x: 0, y: 0, width: 2, height: 1 }).data), [255, 128, 0, 255, 255, 0, 10, 128])
  assert.deepEqual(Array.from(doc.composite.data), [255, 128, 0, 255, 255, 0, 10, 255])
  assert.deepEqual(doc.issues.map((issue) => issue.code), ['bit-depth-reduced'])
})

test('grayscale documents open in color with their gray values', async () => {
  const flat = await psd.importPsd(deepPsd({ width: 3, height: 1, depth: 8, colorMode: 1, composite: [[0, 128, 255]] }))
  assert.equal(flat.layers.length, 1)
  assert.equal(flat.layers[0].isBackground, true)
  assert.deepEqual(Array.from(flat.layers[0].surface.read({ x: 0, y: 0, width: 3, height: 1 }).data), [0, 0, 0, 255, 128, 128, 128, 255, 255, 255, 255, 255])
  const layered = await psd.importPsd(deepPsd({
    width: 2, height: 1, depth: 8, colorMode: 1, mergedAlpha: true,
    layers: [{ name: 'Ink', left: 0, top: 0, width: 2, height: 1, channels: { '-1': [255, 64], 0: [30, 200] } }],
    composite: [[30, 200], [255, 64]],
  }))
  assert.deepEqual(layered.layers.map((layer) => layer.name), ['Ink'])
  assert.deepEqual(Array.from(layered.layers[0].surface.read({ x: 0, y: 0, width: 2, height: 1 }).data), [30, 30, 30, 255, 200, 200, 200, 64])
  assert.deepEqual(layered.issues, [])
})

test('groups: pass-through groups dissolve (hidden ones hide their layers); others flatten with an issue', async () => {
  const bytes = ag.writePsdUint8Array({
    width: 40, height: 40,
    children: [
      { name: 'Base', left: 0, top: 0, imageData: solid(40, 40, [255, 255, 255, 255]) },
      {
        name: 'Outer', opened: true, blendMode: 'pass through',
        children: [
          { name: 'One', left: 0, top: 0, clipping: false, imageData: solid(5, 5, [1, 1, 1, 255]) },
          { name: 'Inner', opened: false, blendMode: 'pass through', hidden: true, children: [{ name: 'Two', left: 5, top: 5, imageData: solid(5, 5, [2, 2, 2, 255]) }] },
        ],
      },
      { name: 'Normal set', opened: true, blendMode: 'normal', children: [{ name: 'Plain', left: 10, top: 10, imageData: solid(5, 5, [3, 3, 3, 255]) }] },
      { name: 'Faded', opened: true, blendMode: 'pass through', opacity: 0.5, children: [{ name: 'Half', left: 20, top: 20, imageData: solid(5, 5, [4, 4, 4, 255]) }] },
      { name: 'Clip base', opened: true, blendMode: 'pass through', children: [{ name: 'Shape', left: 30, top: 30, imageData: solid(5, 5, [5, 5, 5, 255]) }] },
      { name: 'Clipped onto group', left: 30, top: 30, clipping: true, imageData: solid(10, 10, [6, 6, 6, 255]) },
      { name: 'Empty group', opened: true, children: [] },
    ],
  }, { noBackground: true })
  const doc = await psd.importPsd(bytes)
  assert.deepEqual(doc.layers.map((layer) => [layer.name, layer.visible, layer.clipped]), [
    ['Base', true, false],
    ['Outer / One', true, false],
    ['Outer / Inner / Two', false, false],
    ['Normal set / Plain', true, false],
    ['Faded', true, false],
    ['Clip base', true, false],
    ['Clipped onto group', true, true],
  ])
  assert.equal(doc.layers[4].opacity, 128 / 255)
  assert.deepEqual(surfacePixels(doc.layers[4]), { x: 20, y: 20, pixels: Buffer.from(solid(5, 5, [4, 4, 4, 255]).data) })
  assert.deepEqual(doc.issues.map((issue) => `${issue.code}:${issue.layerName}`), ['group-flattened:Faded', 'group-flattened:Clip base'])
})

test('memory limits, refused colour modes and damaged files give friendly errors', async () => {
  const bytes = richPsd()
  await assert.rejects(psd.importPsd(bytes, { memoryLimitBytes: 64 * 1024 }), (error) => {
    assert.match(error.message, /too large to open with its layers/)
    assert.match(error.message, /Merge or delete layers/)
    assert.doesNotMatch(error.message, /Exceeded memory limit|RangeError/)
    assert.equal(error.code, psd.PSD_MEMORY_ERROR_CODE)
    return true
  })
  const cmyk = Buffer.from(bytes)
  cmyk.writeUInt16BE(4, 24)
  await assert.rejects(psd.importPsd(new Uint8Array(cmyk)), /This PSD uses CMYK color\. In Photoshop choose Image > Mode > RGB Color, then save a copy\./)
  const lab = Buffer.from(bytes)
  lab.writeUInt16BE(9, 24)
  await assert.rejects(psd.importPsd(new Uint8Array(lab)), /uses Lab color/)
  const psb = Buffer.from(bytes)
  psb.writeUInt16BE(2, 4)
  await assert.rejects(psd.importPsd(new Uint8Array(psb)), /\.psb\) are not supported/)
  await assert.rejects(psd.importPsd(new Uint8Array(Buffer.from('not a photoshop document at all'))), /not a Photoshop document/)
  const huge = Buffer.from(bytes)
  huge.writeUInt32BE(25_000, 18)
  await assert.rejects(psd.importPsd(new Uint8Array(huge)), /too large to edit safely/)
  const truncated = bytes.slice(0, Math.floor(bytes.length / 3))
  await assert.rejects(psd.importPsd(truncated), (error) => {
    assert.match(error.message, /damaged|can't read/)
    return true
  })
  // Export refuses a merged image of the wrong size instead of writing a broken file.
  const doc = await psd.importPsd(bytes)
  await assert.rejects(psd.exportPsd(doc, solid(10, 10, [0, 0, 0, 255])), /does not match the document size/)
})

test('"Open flattened" gives one layer with the exact merged image', async () => {
  const doc = await psd.importPsd(richPsd(), { mode: 'flattened' })
  assert.equal(doc.layers.length, 1)
  assert.equal(doc.layers[0].name, 'Background')
  assert.equal(doc.layers[0].isBackground, true)
  assert.deepEqual(Buffer.from(doc.layers[0].surface.read({ x: 0, y: 0, width: W, height: H }).data), Buffer.from(doc.composite.data))
  // A transparent merged image opens as "Layer 0", with Photoshop's white matte removed.
  const transparent = ag.writePsdUint8Array({
    width: 4, height: 1,
    imageData: { width: 4, height: 1, data: new Uint8ClampedArray([255, 0, 0, 128, 0, 0, 255, 255, 0, 0, 0, 0, 0, 255, 0, 64]) },
    children: [{ name: 'Layer 1', left: 0, top: 0, imageData: { width: 4, height: 1, data: new Uint8ClampedArray([255, 0, 0, 128, 0, 0, 255, 255, 0, 0, 0, 0, 0, 255, 0, 64]) } }],
  })
  const flat = await psd.importPsd(transparent, { mode: 'flattened' })
  assert.equal(flat.layers[0].name, 'Layer 0')
  const merged = Array.from(flat.composite.data)
  const expected = [255, 0, 0, 128, 0, 0, 255, 255, 0, 0, 0, 0, 0, 255, 0, 64]
  for (let i = 0; i < expected.length; i += 1) assert.ok(Math.abs(merged[i] - expected[i]) <= 2, `merged[${i}] ${merged[i]} vs ${expected[i]}`)
})

test('Photoshop features Simple cannot edit are reported, and the rest still opens', async () => {
  const square = { open: false, operation: 'combine', fillRule: 'non-zero', knots: [[4, 4], [20, 4], [20, 20], [4, 20]].map(([x, y]) => ({ linked: true, points: [x, y, x, y, x, y] })) }
  const bytes = ag.writePsdUint8Array({
    width: 32, height: 32,
    children: [
      { name: 'Background', left: 0, top: 0, imageData: solid(32, 32, [255, 255, 255, 255]) },
      { name: 'Styled', left: 0, top: 0, imageData: solid(8, 8, [9, 9, 9, 255]), effects: { dropShadow: [{ enabled: true, opacity: 0.5, size: { units: 'Pixels', value: 4 } }] } },
      { name: 'Half fill', left: 0, top: 0, fillOpacity: 0.5, imageData: solid(8, 8, [9, 9, 9, 255]) },
      { name: 'Mixer', adjustment: { type: 'channel mixer', monochrome: false, red: { red: 100, green: 0, blue: 0, constant: 0 }, green: { red: 0, green: 100, blue: 0, constant: 0 }, blue: { red: 0, green: 0, blue: 100, constant: 0 } } },
      { name: 'Masked shape', left: 0, top: 0, imageData: solid(32, 32, [0, 0, 255, 255]), vectorMask: { paths: [square] } },
      { name: 'Blend if', left: 0, top: 0, imageData: solid(4, 4, [1, 2, 3, 255]), blendingRanges: { compositeGrayBlendSource: [40, 80, 255, 255], compositeGraphBlendDestinationRange: [0, 0, 255, 255], ranges: [] } },
    ],
  })
  const doc = await psd.importPsd(bytes)
  assert.deepEqual(doc.layers.map((layer) => layer.name), ['Background', 'Styled', 'Half fill', 'Masked shape', 'Blend if'])
  assert.ok(Math.abs(doc.layers[2].opacity - 0.5) <= 1 / 255, 'fill opacity folds into opacity without layer styles')
  const shapeMask = doc.layers[3].mask
  assert.ok(shapeMask, 'the vector mask became a pixel mask')
  assert.equal(shapeMask.surface.read({ x: 10, y: 10, width: 1, height: 1 }).data[0], 255)
  assert.equal(shapeMask.surface.read({ x: 25, y: 25, width: 1, height: 1 }).data[0], 0)
  const codes = doc.issues.map((issue) => `${issue.code}:${issue.layerName}`)
  assert.ok(codes.includes('layer-effects:Styled'))
  assert.ok(codes.includes('unsupported-adjustment:Mixer'))
  assert.ok(codes.includes('vector-mask:Masked shape'))
  assert.ok(codes.includes('knockout-or-advanced-blending:Blend if'))
  assert.ok(!codes.some((code) => code.startsWith('fill-opacity')))
  assert.match(doc.issues.find((issue) => issue.code === 'unsupported-adjustment').detail, /Channel Mixer adjustment layers are not supported/)
})

test('more layers than the editor allows: the bottom ones merge into one, exactly', async () => {
  const count = LIMITS.maxLayers + 5
  const children = [{ name: 'Background', left: 0, top: 0, imageData: solid(16, 16, [255, 255, 255, 255]) }]
  for (let index = 1; index < count; index += 1) {
    children.push({ name: `Dot ${index}`, left: index % 16, top: Math.floor(index / 16) % 16, blendMode: index % 3 ? 'normal' : 'multiply', opacity: 0.7, imageData: solid(1, 1, [index % 256, 40, 200, 255]) })
  }
  const bytes = ag.writePsdUint8Array({ width: 16, height: 16, children })
  const doc = await psd.importPsd(bytes)
  assert.equal(doc.layers.length, LIMITS.maxLayers)
  assert.equal(doc.layers[0].name, 'Merged layers (6)')
  assert.equal(doc.layers[1].name, 'Dot 6')
  assert.ok(doc.issues.some((issue) => issue.code === 'group-flattened' && /bottom 6 layers were merged/.test(issue.detail)))
  // The merge is exact: compositing the stored layers gives the same image as before merging.
  const reference = { width: 16, height: 16, layers: [] }
  for (const child of children) {
    reference.layers.push(documentModule.createRasterLayer({
      name: child.name, surface: tiles.surfaceFromBuffer(child.imageData, child.left, child.top), opacity: Math.round((child.opacity ?? 1) * 255) / 255,
      blendMode: child.blendMode === 'multiply' ? 'multiply' : 'normal',
    }))
  }
  assert.equal(hashBuffer(flattenDocument(doc)), hashBuffer(flattenDocument(reference)))
})

test('ag-psd stays a lazily loaded dependency and the module exposes the contract', () => {
  for (const name of ['importPsd', 'exportPsd', 'configurePsdCanvas', 'sniffPsd', 'blendFromPsd', 'blendToPsd', 'adjustmentFromPsd', 'adjustmentToPsd']) {
    assert.equal(typeof psd[name], 'function', name)
  }
  const source = require('node:fs').readFileSync(path.join(root, 'src', 'advanced', 'psd.ts'), 'utf8')
  assert.doesNotMatch(source, /^import [^\n]*from 'ag-psd'/m, 'only type imports of ag-psd at module level')
  assert.match(source, /import\('ag-psd'\)/)
  assert.doesNotMatch(source, /initialize-canvas/, 'never the node-canvas entry point')
})
