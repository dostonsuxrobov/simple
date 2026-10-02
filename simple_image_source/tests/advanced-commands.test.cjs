'use strict'
// WP6 commands and editor state (src/advanced/commands.ts, src/advanced/editorState.ts).
//   Part 1, editor state: defaults, persistence (round trip, damaged storage, out-of-range values), toolbox
//   slot memory, Photoshop's [ / ] size ladder, hardness steps and digit-key opacity.
//   Part 2, commands on the real document store (inline imaging client, fake view and services): each
//   command makes exactly one history step with Photoshop's label (none for selection-only and active-layer
//   changes, which also keep the document unmodified), refusals change nothing, pixels come out right and
//   undo restores the exact bytes.
const test = require('node:test')
const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const path = require('node:path')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping (not disabled by --no-experimental-strip-types), '
  + 'or run node with --experimental-strip-types.'

function load(relative) {
  if (!(process.features && process.features.typescript)) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

const editorState = load('advanced/editorState.ts')
const commands = load('advanced/commands.ts')
const documentModule = load('advanced/document.ts')
const tiles = load('advanced/tiles.ts')
const composite = load('advanced/composite.ts')
const selectionModule = load('advanced/selection.ts')
const mask = load('imaging/mask.ts')
const workerOps = load('shared/worker-ops/index.ts')
const workerClient = load('shared/workerClient.ts')

// =============================================================================================
// Part 1: editor state
// =============================================================================================

function memoryStorage(initial = {}) {
  const data = new Map(Object.entries(initial))
  return {
    data,
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => { data.set(key, String(value)) },
  }
}

test('editor store: Photoshop defaults and a working store without storage', () => {
  const store = editorState.createEditorStore('k', { storage: null })
  const state = store.getState()
  assert.equal(state.tool, 'move')
  assert.equal(state.springFrom, null)
  assert.deepEqual(state.foreground, { r: 0, g: 0, b: 0 })
  assert.deepEqual(state.background, { r: 255, g: 255, b: 255 })
  assert.equal(state.options.brush.spacing, 0.25, 'Photoshop spacing 25%')
  assert.equal(state.options.wand.tolerance, 32)
  assert.equal(state.options.crop.deleteCroppedPixels, true)
  assert.deepEqual(state.panels, { layers: true, properties: true, history: true, color: true })
  let calls = 0
  const off = store.subscribe(() => { calls += 1 })
  store.update({ tool: 'brush' })
  store.update({ tool: 'brush' })
  assert.equal(calls, 1, 'an unchanged value does not notify')
  store.update({ tool: 'not-a-tool', foreground: { r: 300, g: -4, b: 12.6 } })
  assert.equal(store.getState().tool, 'brush', 'unknown tools are ignored')
  assert.deepEqual(store.getState().foreground, { r: 255, g: 0, b: 13 }, 'colours are clamped and rounded')
  store.updateOptions('brush', { size: 99999, hardness: -1, blendMode: 'nonsense' })
  assert.equal(store.getState().options.brush.size, 5000)
  assert.equal(store.getState().options.brush.hardness, 0)
  assert.equal(store.getState().options.brush.blendMode, 'normal')
  off()
})

test('editor store: settings persist, and damaged storage falls back field by field', () => {
  const storage = memoryStorage()
  const first = editorState.createEditorStore('settings', { storage, writeDelayMs: 0 })
  first.update({ tool: 'brush', foreground: { r: 10, g: 20, b: 30 }, panels: { history: false } })
  first.updateOptions('brush', { size: 77, opacity: 0.5 })
  first.update({ tool: 'marquee-ellipse' })
  first.update({ view: { zoom: 3, offsetX: 1, offsetY: 2 } })
  first.dispose()
  const saved = JSON.parse(storage.getItem('settings'))
  assert.equal(saved.version, 1)
  assert.equal(saved.view, undefined, 'the view is not a setting')
  const second = editorState.createEditorStore('settings', { storage })
  const state = second.getState()
  assert.equal(state.tool, 'marquee-ellipse')
  assert.deepEqual(state.foreground, { r: 10, g: 20, b: 30 })
  assert.equal(state.options.brush.size, 77)
  assert.equal(state.options.brush.opacity, 0.5)
  assert.equal(state.panels.history, false)
  assert.equal(state.panels.layers, true)
  assert.equal(second.slotTool('marquee-rect'), 'marquee-ellipse', 'the slot remembers its last tool')
  second.dispose()

  const broken = memoryStorage({ bad: '{"tool":"lasso","options":{"brush":{"size":"huge","hardness":0.3},"wand":{"tolerance":9000}},"foreground":"red","swatches":[{"r":1,"g":2,"b":3},"x",{"r":1}]}' })
  const recovered = editorState.createEditorStore('bad', { storage: broken }).getState()
  assert.equal(recovered.tool, 'lasso')
  assert.equal(recovered.options.brush.size, editorState.DEFAULT_TOOL_OPTIONS.brush.size, 'a mistyped size takes the default')
  assert.equal(recovered.options.brush.hardness, 0.3, 'valid neighbours are kept')
  assert.equal(recovered.options.wand.tolerance, 255, 'out of range is clamped')
  assert.deepEqual(recovered.foreground, { r: 0, g: 0, b: 0 })
  assert.deepEqual(recovered.swatches, [{ r: 1, g: 2, b: 3 }], 'bad swatches are dropped')
  assert.equal(editorState.createEditorStore('x', { storage: memoryStorage({ x: 'not json' }) }).getState().tool, 'move')
  const throwing = { getItem() { throw new Error('blocked') }, setItem() { throw new Error('blocked') } }
  const safe = editorState.createEditorStore('x', { storage: throwing, writeDelayMs: 0 })
  safe.update({ tool: 'eraser' })
  safe.flush()
  assert.equal(safe.getState().tool, 'eraser', 'blocked storage keeps working in memory')
})

test('editor store: a spring-loaded tool is not remembered as the chosen tool', () => {
  const storage = memoryStorage()
  const store = editorState.createEditorStore('s', { storage, writeDelayMs: 0 })
  store.update({ tool: 'gradient' })
  store.update({ springFrom: 'gradient', tool: 'hand' })
  store.flush()
  assert.equal(JSON.parse(storage.getItem('s')).tool, 'gradient')
  store.update({ tool: 'paint-bucket', springFrom: null })
  assert.equal(store.slotTool('gradient'), 'paint-bucket')
  assert.equal(store.slotTool('brush'), 'brush', 'single slots answer themselves')
})

test('[ and ] follow the Photoshop size ladder; hardness steps by 25%', () => {
  const up = [[1, 2], [9, 10], [10, 20], [15, 20], [95, 100], [100, 125], [125, 150], [200, 250], [250, 300], [300, 400], [400, 500], [4950, 5000], [5000, 5000]]
  for (const [from, to] of up) assert.equal(editorState.stepBrushSize(from, 1), to, `${from} -> ${to}`)
  const down = [[1, 1], [2, 1], [10, 9], [15, 10], [20, 10], [100, 90], [125, 100], [200, 175], [250, 200], [300, 250], [400, 300]]
  for (const [from, to] of down) assert.equal(editorState.stepBrushSize(from, -1), to, `${from} -> ${to}`)
  assert.equal(editorState.stepHardness(1, -1), 0.75)
  assert.equal(editorState.stepHardness(0.8, -1), 0.75)
  assert.equal(editorState.stepHardness(0.8, 1), 1)
  assert.equal(editorState.stepHardness(0, -1), 0)
  assert.equal(editorState.stepHardness(0.5, 1), 0.75)
})

test('digit keys: 1 = 10%, 0 = 100%, two quick digits are exact', () => {
  const first = editorState.opacityFromDigit(4, 1000, 'brush', null)
  assert.equal(first.value, 0.4)
  const second = editorState.opacityFromDigit(5, 1300, 'brush', first.memory)
  assert.equal(second.value, 0.45)
  assert.equal(second.memory, null, 'a third digit starts again')
  assert.equal(editorState.opacityFromDigit(0, 0, 'brush', null).value, 1)
  const zero = editorState.opacityFromDigit(0, 0, 'brush', null)
  assert.equal(editorState.opacityFromDigit(5, 200, 'brush', zero.memory).value, 0.05)
  assert.equal(editorState.opacityFromDigit(0, 200, 'brush', zero.memory).value, 0)
  assert.equal(editorState.opacityFromDigit(5, 5000, 'brush', first.memory).value, 0.5, 'too slow: a new single digit')
  assert.equal(editorState.opacityFromDigit(5, 1100, 'eraser', first.memory).value, 0.5, 'another target starts again')
})

// =============================================================================================
// Part 2: commands on the real store
// =============================================================================================

const W = 64
const H = 48
const RED = { r: 255, g: 0, b: 0 }
const BLUE = { r: 0, g: 0, b: 255 }
const WHITE = { r: 255, g: 255, b: 255 }
const BLACK = { r: 0, g: 0, b: 0 }

function solid(width, height, color, alpha = 255) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let p = 0; p < data.length; p += 4) {
    data[p] = color.r
    data[p + 1] = color.g
    data[p + 2] = color.b
    data[p + 3] = alpha
  }
  return { width, height, data }
}

/** Opaque ramp: every pixel different (catches any wrong mapping). */
function ramp(width = W, height = H) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const p = (y * width + x) * 4
      data[p] = (x * 3) & 255
      data[p + 1] = (y * 5) & 255
      data[p + 2] = (x * 7 + y * 11) & 255
      data[p + 3] = 255
    }
  }
  return { width, height, data }
}

function rectSelection(state, rect) {
  const target = mask.createMaskBuffer(state.width, state.height)
  mask.rasterizeRect(target, rect, false)
  return selectionModule.selectionFromMask(target, 1000 + Math.floor(Math.random() * 1000))
}

function makeHarness(setup = {}) {
  let counter = 0
  const host = {
    revision: 0,
    nextRevision: () => (counter += 1),
    setRevision(revision) { host.revision = revision },
    currentRevision: () => host.revision,
  }
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(setup.backgroundPixels ?? ramp()) })
  const layer = documentModule.createRasterLayer({ name: 'Layer 1', surface: tiles.surfaceFromBuffer(solid(16, 16, RED)), offsetX: 8, offsetY: 8 })
  const layers = setup.layers ?? [background, layer]
  const store = documentModule.createDocumentStore({
    width: setup.width ?? W, height: setup.height ?? H, ppi: 72, layers, host, baseLabel: 'Open', activeLayerId: setup.active ?? layers[layers.length - 1].id,
  })
  const editor = editorState.createEditorStore('test', { storage: null, initial: { tool: setup.tool ?? 'move', foreground: RED, background: WHITE } })
  const notes = []
  const calls = { zoom: [], fit: 0, actual: 0, transform: 0, dialogs: [], clipboard: [], overlays: 0, commit: 0, cancel: 0 }
  let view = { zoom: 1, offsetX: 0, offsetY: 0 }
  const viewController = {
    getView: () => view,
    getViewportSize: () => ({ width: 640, height: 480, dpr: 1 }),
    docToScreen: (p) => ({ x: p.x * view.zoom + view.offsetX, y: p.y * view.zoom + view.offsetY }),
    screenToDoc: (p) => ({ x: (p.x - view.offsetX) / view.zoom, y: (p.y - view.offsetY) / view.zoom }),
    zoomAt(zoom, anchor) { calls.zoom.push({ zoom, anchor }); view = { ...view, zoom } },
    panBy(dx, dy) { view = { ...view, offsetX: view.offsetX + dx, offsetY: view.offsetY + dy } },
    fit() { calls.fit += 1 },
    actualPixels() { calls.actual += 1 },
    requestOverlay() { calls.overlays += 1 },
  }
  const compositor = {
    level: 0,
    attach() {},
    setView() {},
    invalidate() {},
    setPreview() {},
    settle: async () => {},
    flatten: async () => composite.flattenDocument(store.getState()),
    renderToCanvas: async () => { throw new Error('no canvas in Node') },
    sample: (x, y, size, source) => composite.sampleDocument(store.getState(), x, y, size, source),
    dispose() {},
  }
  const imaging = workerClient.createImagingClient({ inline: workerOps.HANDLERS })
  const session = setup.session ?? null
  const services = setup.services === null ? undefined : {
    activeTool: () => session,
    transformSession: () => null,
    startTransform: () => { calls.transform += 1 },
    openDialog: (request) => { calls.dialogs.push(request); return true },
    canOpenDialog: () => true,
    clipboard: {
      async write(pixels) { calls.clipboard.push(pixels) },
      async read() { return setup.clipboard ?? null },
    },
    busy: () => Boolean(setup.busy),
    ...(setup.services ?? {}),
  }
  const ctx = {
    store,
    editor,
    view: viewController,
    compositor,
    imaging,
    host: {
      ...host,
      notify(message, tone) { notes.push({ message, tone }) },
      isSuspended: () => false,
      requestExit() {},
      save: async () => true,
      openExportMenu() {},
      print() {},
      copyPng: async () => {},
      readClipboardImage: async () => null,
    },
    setCursor() {},
    setHint() {},
    setOverlayElement() {},
    services,
  }
  commands.attachCommands(store)
  const run = (id) => commands.runCommandAsync(id, ctx)
  return {
    store, ctx, editor, host, notes, calls, background, layer, run,
    state: () => store.getState(),
    steps: () => store.history.getState().entries.length - 1,
    labels: () => store.history.getState().entries.slice(1).map((entry) => entry.label),
    active: () => store.getState().layers.find((entry) => entry.id === store.getState().activeLayerId),
    find: (name) => store.getState().layers.find((entry) => entry.name === name),
    names: () => store.getState().layers.map((entry) => entry.name),
    pixel(layerOrName, x, y) {
      const state = store.getState()
      const found = typeof layerOrName === 'string' ? state.layers.find((entry) => entry.name === layerOrName) : state.layers.find((entry) => entry.id === layerOrName.id)
      const source = found.kind === 'raster' ? { surface: found.surface, ox: found.offsetX, oy: found.offsetY } : { surface: found.raster.surface, ox: found.raster.offsetX, oy: found.raster.offsetY }
      return Array.from(source.surface.read({ x: x - source.ox, y: y - source.oy, width: 1, height: 1 }).data)
    },
    composite(x, y) {
      return Array.from(composite.compositeRect(store.getState(), { x, y, width: 1, height: 1 }).data)
    },
    hash() {
      const state = store.getState()
      return crypto.createHash('sha1').update(composite.flattenDocument(state).data).digest('hex')
    },
    select(rect) {
      const selection = rectSelection(store.getState(), rect)
      store.transact('Select', 'selection', (tx) => tx.setSelection(selection), { affectsOutput: false })
      return selection
    },
  }
}

test('registry: every command has a label and an enabled state; unknown ids are refused', async () => {
  const h = makeHarness()
  assert.ok(commands.COMMAND_IDS.length > 140)
  assert.equal(new Set(commands.COMMAND_IDS).size, commands.COMMAND_IDS.length)
  for (const id of commands.COMMAND_IDS) {
    const label = commands.commandLabel(id, h.ctx)
    assert.ok(label && label !== id, `${id} has a label`)
    assert.equal(typeof commands.isCommandEnabled(id, h.ctx), 'boolean', id)
  }
  assert.equal(commands.isEditorCommand('adjust.quick'), false, 'Simple-only adjustment')
  assert.equal(commands.isCommandEnabled('nope.nothing', h.ctx), false)
  await h.run('nope.nothing')
  assert.match(h.notes.at(-1).message, /not a command/)
  assert.equal(h.steps(), 0)
})

test('layers: new, duplicate and delete are one step each with Photoshop names', async () => {
  const h = makeHarness()
  await h.run('layer.new')
  assert.deepEqual(h.names(), ['Background', 'Layer 1', 'Layer 2'])
  assert.equal(h.active().name, 'Layer 2', 'the new layer is active, above the old active one')
  assert.equal(h.host.revision > 0, true, 'a new layer marks the document modified')
  await h.run('layer.select-below')
  assert.equal(h.active().name, 'Layer 1')
  await h.run('layer.duplicate')
  assert.deepEqual(h.names(), ['Background', 'Layer 1', 'Layer 1 copy', 'Layer 2'])
  assert.deepEqual(h.pixel('Layer 1 copy', 10, 10), [255, 0, 0, 255])
  await h.run('layer.delete')
  assert.deepEqual(h.names(), ['Background', 'Layer 1', 'Layer 2'])
  assert.deepEqual(h.labels(), ['New Layer', 'Duplicate Layer', 'Delete Layer'], 'selecting a layer is not a history step')
  const single = makeHarness({ layers: [documentModule.createRasterLayer({ name: 'Only', surface: tiles.surfaceFromBuffer(solid(4, 4, RED)) })] })
  assert.equal(commands.isCommandEnabled('layer.delete', single.ctx), false)
  await single.run('layer.delete')
  assert.equal(single.state().layers.length, 1, 'the last layer cannot be deleted')
  assert.match(single.notes.at(-1).message, /at least one layer/)
})

test('selection commands never mark the document modified; Reselect brings back the last selection', async () => {
  const h = makeHarness()
  const revision = h.host.revision
  await h.run('select.all')
  assert.ok(selectionModule.isSelectAll(h.state().selection))
  await h.run('select.deselect')
  assert.equal(h.state().selection, null)
  assert.equal(commands.isCommandEnabled('select.reselect', h.ctx), true)
  await h.run('select.reselect')
  assert.ok(selectionModule.isSelectAll(h.state().selection), 'Reselect restores the select-all')
  h.select({ x: 10, y: 10, width: 20, height: 10 })
  await h.run('select.inverse')
  assert.equal(h.state().selection.mask.data[15 * W + 15], 0, 'inside is now outside')
  assert.equal(h.state().selection.mask.data[2 * W + 2], 255)
  // A deselect by anyone (not only the command) is remembered.
  h.store.transact('Deselect', 'selection', (tx) => tx.setSelection(null), { affectsOutput: false })
  await h.run('select.reselect')
  assert.equal(h.state().selection.mask.data[2 * W + 2], 255)
  await h.run('select.load-layer-alpha')
  assert.deepEqual(h.state().selection.bounds, { x: 8, y: 8, width: 16, height: 16 }, 'Layer 1 alpha')
  assert.equal(h.host.revision, revision, 'no content revision for selection changes')
  await h.run('select.feather')
  assert.deepEqual(h.calls.dialogs.at(-1), { kind: 'modify-selection', operation: 'feather' })
})

test('Layer via Copy / Cut take the selected pixels; without a selection Ctrl+J duplicates', async () => {
  const h = makeHarness({ active: undefined })
  h.select({ x: 4, y: 4, width: 8, height: 8 })
  await h.run('layer.select-below')
  assert.equal(h.active().name, 'Background')
  await h.run('layer.via-copy')
  const copy = h.active()
  assert.equal(copy.name, 'Layer 2')
  assert.deepEqual(h.pixel(copy, 6, 6), Array.from(ramp().data.slice((6 * W + 6) * 4, (6 * W + 6) * 4 + 4)))
  assert.deepEqual(h.pixel(copy, 20, 20), [0, 0, 0, 0], 'outside the selection stays empty')
  assert.ok(h.state().selection, 'the selection stays (Photoshop)')
  await h.run('layer.select-below')
  await h.run('layer.select-below')
  assert.equal(h.active().name, 'Background')
  await h.run('layer.via-cut')
  assert.deepEqual(h.pixel('Background', 6, 6), [255, 255, 255, 255], 'cutting from the Background leaves the background colour')
  await h.run('select.deselect')
  const before = h.state().layers.length
  await h.run('layer.via-copy')
  assert.equal(h.state().layers.length, before + 1)
  assert.match(h.active().name, /copy/)
  assert.deepEqual(h.labels(), ['Select', 'Layer Via Copy', 'Layer Via Cut', 'Deselect', 'Layer Via Copy'])
  await h.run('layer.via-cut')
  assert.match(h.notes.at(-1).message, /needs a selection/)
})

test('fills: foreground / background / preserve transparency, inside the selection, one step each', async () => {
  const h = makeHarness()
  h.select({ x: 10, y: 10, width: 4, height: 4 })
  await h.run('edit.fill-foreground')
  assert.deepEqual(h.pixel('Layer 1', 11, 11), [255, 0, 0, 255])
  h.editor.update({ foreground: BLUE })
  await h.run('edit.fill-foreground')
  assert.deepEqual(h.pixel('Layer 1', 11, 11), [0, 0, 255, 255])
  assert.deepEqual(h.pixel('Layer 1', 20, 20), [255, 0, 0, 255], 'outside the selection unchanged')
  await h.run('select.deselect')
  // Preserve transparency: the empty part of Layer 1 stays empty.
  await h.run('edit.fill-foreground-preserve')
  assert.deepEqual(h.pixel('Layer 1', 20, 20), [0, 0, 255, 255])
  assert.deepEqual(h.pixel('Layer 1', 40, 40), [0, 0, 0, 0], 'transparent pixels stay transparent')
  await h.run('edit.fill-background')
  assert.deepEqual(h.pixel('Layer 1', 40, 40), [255, 255, 255, 255], 'a normal fill covers the whole canvas')
  assert.deepEqual(h.labels(), ['Select', 'Fill', 'Fill', 'Deselect', 'Fill', 'Fill'])
  h.store.history.undo()
  assert.deepEqual(h.pixel('Layer 1', 40, 40), [0, 0, 0, 0])
})

test('Delete clears the selection (the Background takes the background colour); without one it deletes the layer', async () => {
  const h = makeHarness()
  h.select({ x: 8, y: 8, width: 4, height: 4 })
  await h.run('edit.clear')
  assert.equal(h.pixel('Layer 1', 9, 9)[3], 0, 'cleared to transparency')
  assert.deepEqual(h.pixel('Layer 1', 14, 14), [255, 0, 0, 255])
  await h.run('layer.select-below')
  await h.run('edit.clear')
  assert.deepEqual(h.pixel('Background', 9, 9), [255, 255, 255, 255])
  await h.run('select.deselect')
  await h.run('layer.select-above')
  await h.run('edit.clear')
  assert.deepEqual(h.names(), ['Background'], 'Delete without a selection removes the layer')
  await h.run('edit.clear')
  assert.deepEqual(h.names(), ['Background'], 'but never the last one')
})

test('Merge Down keeps the look and the lower layer\'s name; Merge Visible and Stamp Visible keep the composite', async () => {
  const h = makeHarness()
  h.store.transact('Props', 'layer', (tx) => tx.updateLayer(h.layer.id, { opacity: 0.5, blendMode: 'multiply' }))
  const look = h.hash()
  await h.run('layer.merge-down')
  assert.deepEqual(h.names(), ['Background'])
  assert.equal(h.active().isBackground, true)
  assert.equal(h.hash(), look, 'merging changes nothing visible')
  h.store.history.undo()
  assert.deepEqual(h.names(), ['Background', 'Layer 1'])
  assert.equal(commands.isCommandEnabled('layer.merge-down', h.ctx), true)

  await h.run('layer.new')
  h.select({ x: 30, y: 30, width: 6, height: 6 })
  await h.run('edit.fill-foreground')
  await h.run('select.deselect')
  h.store.transact('Hide', 'layer', (tx) => tx.updateLayer(h.layer.id, { visible: false }))
  const look2 = h.hash()
  await h.run('layer.stamp-visible')
  const stamp = h.active()
  assert.equal(stamp.name, 'Layer 3')
  h.store.transact('Hide stamp', 'layer', (tx) => tx.updateLayer(stamp.id, { visible: false }))
  assert.equal(h.hash(), look2)
  h.store.history.undo()
  h.store.history.undo()
  await h.run('layer.merge-visible')
  assert.deepEqual(h.names(), ['Background', 'Layer 1'], 'the hidden layer stays, the visible ones become the Background')
  assert.equal(h.find('Layer 1').visible, false)
  assert.equal(h.hash(), look2)
})

test('Flatten Image composites onto white and discards hidden layers', async () => {
  const transparent = documentModule.createRasterLayer({ name: 'Layer 0', surface: tiles.surfaceFromBuffer(solid(10, 10, BLUE, 128)) })
  const hidden = documentModule.createRasterLayer({ name: 'Hidden', visible: false, surface: tiles.surfaceFromBuffer(solid(20, 20, RED)) })
  const h = makeHarness({ layers: [transparent, hidden] })
  await h.run('layer.flatten')
  assert.deepEqual(h.names(), ['Background'])
  assert.equal(h.active().isBackground, true)
  const px = h.pixel('Background', 2, 2)
  assert.ok(Math.abs(px[0] - 127) <= 1 && Math.abs(px[2] - 255) <= 1 && px[3] === 255, `half-transparent blue over white: ${px}`)
  assert.deepEqual(h.pixel('Background', 30, 30), [255, 255, 255, 255])
  assert.deepEqual(h.labels(), ['Flatten Image'])
})

test('masks: add from the selection (Background converts), apply, toggle, delete', async () => {
  const h = makeHarness({ active: undefined })
  await h.run('layer.select-below')
  h.select({ x: 0, y: 0, width: 32, height: 48 })
  await h.run('layer.add-mask')
  const bg = h.state().layers[0]
  assert.equal(bg.name, 'Layer 0', 'the Background becomes Layer 0 to take a mask')
  assert.equal(bg.isBackground, false)
  assert.ok(bg.mask)
  assert.equal(h.state().editTarget, 'mask', 'the new mask is targeted')
  assert.equal(h.state().selection, null, 'the selection became the mask')
  assert.equal(h.composite(40, 2)[3], 0, 'outside the selection is hidden')
  await h.run('layer.toggle-mask')
  assert.equal(h.state().layers[0].mask.enabled, false)
  assert.equal(h.composite(40, 2)[3], 255)
  assert.equal(commands.commandLabel('layer.toggle-mask', h.ctx), 'Enable Layer Mask')
  await h.run('layer.toggle-mask')
  await h.run('layer.apply-mask')
  assert.equal(h.state().layers[0].mask, null)
  assert.equal(h.pixel('Layer 0', 40, 2)[3], 0, 'applying wrote the mask into the pixels')
  assert.equal(h.pixel('Layer 0', 2, 2)[3], 255)
  await h.run('layer.add-mask')
  await h.run('layer.delete-mask')
  assert.equal(h.state().layers[0].mask, null)
  assert.deepEqual(h.labels(), ['Select', 'Add Layer Mask', 'Disable Layer Mask', 'Enable Layer Mask', 'Apply Layer Mask', 'Add Layer Mask', 'Delete Layer Mask'])
})

test('clipping, arrange and Layer from Background', async () => {
  const h = makeHarness()
  await h.run('layer.toggle-clipping')
  assert.equal(h.active().clipped, true)
  assert.equal(commands.commandLabel('layer.toggle-clipping', h.ctx), 'Release Clipping Mask')
  await h.run('layer.toggle-clipping')
  await h.run('layer.new')
  await h.run('layer.to-back')
  assert.deepEqual(h.names(), ['Background', 'Layer 2', 'Layer 1'], 'nothing goes below the Background')
  await h.run('layer.raise')
  assert.deepEqual(h.names(), ['Background', 'Layer 1', 'Layer 2'])
  await h.run('layer.lower')
  await h.run('layer.to-front')
  assert.deepEqual(h.names(), ['Background', 'Layer 1', 'Layer 2'])
  await h.run('layer.select-below')
  await h.run('layer.select-below')
  assert.equal(h.active().name, 'Background')
  assert.equal(commands.isCommandEnabled('layer.raise', h.ctx), false, 'the Background is locked at the bottom')
  await h.run('layer.toggle-clipping')
  assert.match(h.notes.at(-1).message, /needs a layer below/)
  await h.run('layer.from-background')
  assert.equal(h.active().name, 'Layer 0')
  assert.equal(h.active().isBackground, false)
  assert.deepEqual(h.active().locks, { pixels: false, position: false, transparency: false })
})

test('Image Rotation and flips are exact permutations, carry the selection and undo exactly', async () => {
  const h = makeHarness()
  h.select({ x: 0, y: 0, width: 10, height: 5 })
  const original = h.hash()
  const corner = h.composite(0, 0)
  await h.run('image.rotate-cw')
  assert.equal(h.state().width, H)
  assert.equal(h.state().height, W)
  assert.deepEqual(h.composite(H - 1, 0), corner, 'top-left goes to top-right')
  const layer1 = h.find('Layer 1')
  assert.deepEqual({ x: layer1.offsetX, y: layer1.offsetY }, { x: H - 24, y: 8 })
  assert.deepEqual(h.state().selection.bounds, { x: H - 5, y: 0, width: 5, height: 10 }, 'the selection turns too')
  await h.run('image.rotate-ccw')
  assert.equal(h.hash(), original, 'clockwise then counter-clockwise is identity')
  await h.run('image.rotate-180')
  assert.deepEqual(h.composite(W - 1, H - 1), corner)
  await h.run('image.rotate-180')
  assert.equal(h.hash(), original)
  await h.run('image.flip-horizontal')
  assert.deepEqual(h.composite(W - 1, 0), corner)
  await h.run('image.flip-vertical')
  assert.deepEqual(h.composite(W - 1, H - 1), corner)
  while (h.store.history.canUndo()) h.store.history.undo()
  assert.equal(h.hash(), original, 'undo restores the exact bytes')
  assert.equal(h.host.revision, 0, 'and the original revision')
})

test('crop to selection and trim', async () => {
  const h = makeHarness()
  h.select({ x: 4, y: 6, width: 20, height: 10 })
  const sample = h.composite(10, 10)
  await h.run('image.crop-to-selection')
  assert.equal(h.state().width, 20)
  assert.equal(h.state().height, 10)
  assert.deepEqual(h.composite(6, 4), sample, 'content shifted by the crop origin')
  assert.equal(h.state().selection, null)
  const layer1 = h.find('Layer 1')
  assert.deepEqual(layer1.surface.contentBounds(), { x: 0, y: 0, width: 16, height: 8 }, 'pixels outside the crop are deleted')

  const t = makeHarness({ layers: [documentModule.createRasterLayer({ name: 'Layer 0', surface: tiles.surfaceFromBuffer(solid(10, 6, RED)), offsetX: 20, offsetY: 15 })] })
  assert.equal(commands.isCommandEnabled('image.trim', t.ctx), true)
  await t.run('image.trim')
  assert.equal(t.state().width, 10)
  assert.equal(t.state().height, 6)
  assert.deepEqual(t.composite(0, 0), [255, 0, 0, 255])
  assert.deepEqual(t.labels(), ['Trim'])
  await t.run('image.trim')
  assert.match(t.notes.at(-1).message, /no transparent edges/)
})

test('Canvas Size anchors the content and extends the Background with the background colour', () => {
  const h = makeHarness()
  const topLeft = h.composite(0, 0)
  assert.equal(commands.resizeCanvas(h.ctx, W + 20, H + 10, 'bottom-right'), true)
  assert.equal(h.state().width, W + 20)
  assert.deepEqual(h.composite(20, 10), topLeft)
  assert.deepEqual(h.composite(0, 0), [255, 255, 255, 255], 'new area takes the background colour')
  const layer1 = h.find('Layer 1')
  assert.deepEqual({ x: layer1.offsetX, y: layer1.offsetY }, { x: 28, y: 18 })
  assert.equal(commands.resizeCanvas(h.ctx, 10, 10, 'top-left'), true)
  assert.equal(h.state().width, 10)
  assert.deepEqual(h.find('Layer 1').surface.contentBounds(), { x: 0, y: 0, width: 16, height: 16 }, 'other layers keep their pixels outside the canvas')
  assert.equal(commands.resizeCanvas(h.ctx, 0, 10), false)
  assert.equal(commands.resizeCanvas(h.ctx, 30000, 10), false)
  assert.deepEqual(h.labels(), ['Canvas Size', 'Canvas Size'])
})

test('Image Size resamples every layer and scales offsets in one step', async () => {
  const h = makeHarness()
  assert.equal(await commands.resizeImage(h.ctx, W / 2, H / 2, { method: 'area', ppi: 144 }), true)
  assert.equal(h.state().width, W / 2)
  assert.equal(h.state().height, H / 2)
  assert.equal(h.state().ppi, 144)
  const layer1 = h.find('Layer 1')
  assert.deepEqual({ x: layer1.offsetX, y: layer1.offsetY }, { x: 4, y: 4 })
  assert.deepEqual(layer1.surface.contentBounds(), { x: 0, y: 0, width: 8, height: 8 })
  assert.deepEqual(h.pixel(layer1, 6, 6), [255, 0, 0, 255], 'a solid layer stays solid')
  assert.deepEqual(h.labels(), ['Image Size'])
  h.store.history.undo()
  assert.equal(h.state().width, W)
  assert.equal(await commands.resizeImage(h.ctx, 0, 5), false)
})

test('arbitrary rotation: 90 degrees is exact, other angles grow the canvas and fill the Background corners', async () => {
  const h = makeHarness()
  const corner = h.composite(0, 0)
  assert.equal(await commands.rotateCanvas(h.ctx, 90), true)
  assert.deepEqual(h.composite(H - 1, 0), corner)
  h.store.history.undo()
  assert.equal(await commands.rotateCanvas(h.ctx, 45), true)
  const size = Math.round((W + H) * Math.SQRT1_2)
  assert.ok(Math.abs(h.state().width - size) <= 1 && Math.abs(h.state().height - size) <= 1, `${h.state().width} x ${h.state().height}`)
  assert.deepEqual(h.composite(0, 0), [255, 255, 255, 255], 'Background corners take the background colour')
  assert.deepEqual(h.labels(), ['Rotate Canvas'])
})

test('adjustments: destructive Invert inside the selection; adjustment layers take the selection as mask', async () => {
  const h = makeHarness()
  h.select({ x: 8, y: 8, width: 4, height: 4 })
  await h.run('adjust.invert')
  assert.deepEqual(h.pixel('Layer 1', 9, 9), [0, 255, 255, 255])
  assert.deepEqual(h.pixel('Layer 1', 14, 14), [255, 0, 0, 255])
  await h.run('adjustment-layer.invert')
  const adjustment = h.active()
  assert.equal(adjustment.kind, 'adjustment')
  assert.equal(adjustment.name, 'Invert 1')
  assert.ok(adjustment.mask, 'the selection became its mask')
  assert.equal(h.state().selection, null)
  assert.deepEqual(h.composite(9, 9), [255, 0, 0, 255], 'inverted twice inside the mask')
  assert.deepEqual(h.composite(14, 14), [255, 0, 0, 255], 'outside its mask the adjustment layer does nothing')
  await h.run('adjust.levels')
  assert.deepEqual(h.calls.dialogs.at(-1), { kind: 'adjustment', type: 'levels' })
  assert.deepEqual(h.labels(), ['Select', 'Invert', 'Invert Layer'])
  const noDialogs = makeHarness({ services: { openDialog: undefined, canOpenDialog: undefined } })
  assert.equal(commands.isCommandEnabled('adjust.levels', noDialogs.ctx), false, 'no dialog, no menu item')
  assert.equal(commands.isCommandEnabled('adjust.invert', noDialogs.ctx), true, 'Invert needs no dialog')
})

test('Auto Tone / Desaturate / filters apply to the active layer; Last Filter repeats', async () => {
  const h = makeHarness()
  await h.run('layer.select-below')
  await h.run('image.desaturate')
  const gray = h.pixel('Background', 30, 30)
  assert.equal(gray[0], gray[1])
  assert.equal(gray[1], gray[2])
  await h.run('image.auto-contrast')
  assert.equal(commands.isCommandEnabled('filter.repeat', h.ctx), false, 'no filter used yet')
  assert.equal(await commands.applyFilterToLayer(h.ctx, { type: 'gaussian-blur', radius: 1 }), true)
  assert.deepEqual(h.editor.getState().lastFilter, { type: 'gaussian-blur', radius: 1 })
  assert.equal(commands.commandLabel('filter.repeat', h.ctx), 'Last Filter (Gaussian Blur)')
  await h.run('filter.repeat')
  await h.run('filter.find-edges')
  await h.run('filter.sharpen-more')
  assert.deepEqual(h.labels(), ['Desaturate', 'Auto Contrast', 'Gaussian Blur', 'Gaussian Blur', 'Find Edges', 'Sharpen More'])
  // A blur of a solid layer leaves its inside unchanged and spills nothing outside the selection.
  const s = makeHarness()
  s.select({ x: 14, y: 14, width: 4, height: 4 })
  await commands.applyFilterToLayer(s.ctx, { type: 'gaussian-blur', radius: 2 })
  assert.deepEqual(s.pixel('Layer 1', 15, 15), [255, 0, 0, 255], 'blurring solid colour away from the layer edge changes nothing')
  assert.deepEqual(s.pixel('Layer 1', 9, 9), [255, 0, 0, 255], 'outside the selection unchanged')
  const edge = makeHarness()
  edge.select({ x: 8, y: 8, width: 4, height: 4 })
  await commands.applyFilterToLayer(edge.ctx, { type: 'gaussian-blur', radius: 2 })
  assert.ok(edge.pixel('Layer 1', 8, 8)[3] < 255, 'at the layer edge the blur mixes in the transparent surroundings (Photoshop)')
  assert.deepEqual(edge.pixel('Layer 1', 7, 7), [0, 0, 0, 0], 'but nothing spills outside the selection')
})

test('a command refuses while another job runs, and never writes over later edits', async () => {
  const busy = makeHarness({ busy: true })
  assert.equal(commands.isCommandEnabled('layer.new', busy.ctx), false)
  assert.equal(commands.isCommandEnabled('view.zoom-in', busy.ctx), true, 'viewing still works')
  await busy.run('layer.new')
  assert.equal(busy.state().layers.length, 2)
  assert.match(busy.notes.at(-1).message, /Wait/)

  const h = makeHarness()
  const pending = commands.applyFilterToLayer(h.ctx, { type: 'gaussian-blur', radius: 1 })
  assert.equal(commands.commandsBusy(h.store), true)
  // Paint while the filter runs: the filter result must not replace the newer pixels.
  h.store.transact('Paint', 'brush', (tx) => {
    const editor = tx.editPixels(h.layer.id, 'pixels')
    editor.writePixels(0, 0, solid(2, 2, BLUE))
  })
  assert.equal(await pending, false)
  assert.match(h.notes.at(-1).message, /changed while/)
  assert.deepEqual(h.pixel('Layer 1', 8, 8), [0, 0, 255, 255])
  await commands.whenCommandsIdle(h.store)
  assert.equal(commands.commandsBusy(h.store), false)
})

test('brush size, hardness and digits act on the active tool; digits set the layer opacity for other tools', async () => {
  const h = makeHarness({ tool: 'brush' })
  await h.run('brush.larger')
  assert.equal(h.editor.getState().options.brush.size, 30)
  await h.run('brush.smaller')
  assert.equal(h.editor.getState().options.brush.size, 20)
  await h.run('brush.softer')
  assert.equal(h.editor.getState().options.brush.hardness, 0.75)
  await h.run('tool.opacity-3')
  assert.equal(h.editor.getState().options.brush.opacity, 0.3)
  await h.run('tool.opacity-5')
  assert.equal(h.editor.getState().options.brush.opacity, 0.35, 'two quick digits')
  await h.run('tool.flow-5')
  assert.equal(h.editor.getState().options.brush.flow, 0.5)
  h.editor.update({ tool: 'eraser' })
  await h.run('brush.larger')
  assert.equal(h.editor.getState().options.eraser.size, 30, 'the eraser has its own size')
  h.editor.update({ springFrom: 'eraser', tool: 'hand' })
  await h.run('brush.larger')
  assert.equal(h.editor.getState().options.eraser.size, 40, 'while Space is held the brush keys still size the eraser')
  h.editor.update({ tool: 'move', springFrom: null })
  assert.equal(commands.isCommandEnabled('brush.larger', h.ctx), false)
  const steps = h.steps()
  await h.run('tool.opacity-5')
  assert.equal(h.active().opacity, 0.5)
  await h.run('tool.opacity-0')
  assert.equal(h.active().opacity, 0.5, '"5" then "0" quickly is exactly 50%')
  assert.equal(h.steps(), steps + 1, 'quick opacity changes coalesce into one step')
})

test('tools: letters select tools, Shift cycles slots and shapes', async () => {
  const h = makeHarness()
  await h.run('tool.brush')
  assert.equal(h.editor.getState().tool, 'brush')
  await h.run('tool.cycle-marquee')
  assert.equal(h.editor.getState().tool, 'marquee-rect', 'from another tool the slot shows its tool first')
  await h.run('tool.cycle-marquee')
  assert.equal(h.editor.getState().tool, 'marquee-ellipse')
  await h.run('tool.cycle-marquee')
  assert.equal(h.editor.getState().tool, 'marquee-rect')
  await h.run('tool.cycle-gradient')
  await h.run('tool.cycle-gradient')
  assert.equal(h.editor.getState().tool, 'paint-bucket')
  await h.run('tool.shape')
  await h.run('tool.cycle-shape')
  assert.equal(h.editor.getState().options.shape.kind, 'ellipse')
  await h.run('edit.swap-colors')
  assert.deepEqual(h.editor.getState().foreground, WHITE)
  await h.run('edit.default-colors')
  assert.deepEqual(h.editor.getState().foreground, BLACK)
  assert.deepEqual(h.editor.getState().background, WHITE)
  assert.equal(h.steps(), 0, 'tools and colours are not history steps')
})

test('history, view and session commands', async () => {
  const session = {
    open: true,
    hasSession() { return this.open },
    commitSession() { calls.commit += 1; this.open = false },
    cancelSession() { calls.cancel += 1; this.open = false },
  }
  const h = makeHarness({ session })
  const calls = h.calls
  await h.run('session.commit')
  assert.equal(calls.commit, 1)
  session.open = true
  await h.run('edit.undo')
  assert.equal(calls.cancel, 1, 'Ctrl+Z inside a session cancels it first')
  session.id = 'text'
  session.open = true
  await h.run('edit.undo')
  assert.equal(calls.cancel, 1, 'but typed text is never thrown away by Undo (the text field undoes its own typing)')
  session.open = false
  delete session.id
  await h.run('layer.new')
  await h.run('layer.new')
  await h.run('edit.undo')
  assert.equal(h.state().layers.length, 3)
  await h.run('edit.redo')
  assert.equal(h.state().layers.length, 4)
  await h.run('edit.toggle-last')
  assert.equal(h.state().layers.length, 3)
  await h.run('edit.toggle-last')
  assert.equal(h.state().layers.length, 4)
  await h.run('view.zoom-in')
  assert.ok(h.calls.zoom.at(-1).zoom > 1)
  await h.run('view.fit')
  await h.run('view.actual-pixels')
  assert.equal(h.calls.fit, 1)
  assert.equal(h.calls.actual, 1)
  await h.run('edit.free-transform')
  assert.equal(h.calls.transform, 1)
  await h.run('view.toggle-panels')
  assert.deepEqual(h.editor.getState().panels, { layers: false, properties: false, history: false, color: false })
  await h.run('view.toggle-panels')
  assert.equal(h.editor.getState().panels.layers, true)
  await h.run('view.panel-color')
  assert.equal(h.editor.getState().panels.color, false)
})

test('copy and paste: Ctrl+C copies the selected pixels, Paste in Place puts them back where they were', async () => {
  const h = makeHarness()
  h.select({ x: 10, y: 10, width: 4, height: 4 })
  await h.run('edit.copy')
  assert.equal(h.calls.clipboard.length, 1)
  assert.equal(h.calls.clipboard[0].width, 4)
  const entry = commands.internalClipboard(h.store)
  assert.deepEqual(entry.rect, { x: 10, y: 10, width: 4, height: 4 })
  await h.run('edit.paste-in-place')
  const pasted = h.active()
  assert.deepEqual({ x: pasted.offsetX, y: pasted.offsetY }, { x: 10, y: 10 })
  assert.deepEqual(h.pixel(pasted, 11, 11), [255, 0, 0, 255])
  await h.run('edit.paste')
  const centred = h.active()
  assert.deepEqual({ x: centred.offsetX, y: centred.offsetY }, { x: 318, y: 238 }, 'Paste centres in the view')
  await h.run('edit.copy-merged')
  assert.equal(h.calls.clipboard.length, 2)
  await h.run('layer.select-below')
  await h.run('layer.select-below')
  assert.equal(h.active().name, 'Layer 1')
  await h.run('edit.cut')
  assert.equal(h.pixel('Layer 1', 11, 11)[3], 0, 'cut leaves transparency')
  assert.deepEqual(h.labels(), ['Select', 'Paste in Place', 'Paste', 'Cut'])
  const empty = makeHarness({ services: { clipboard: { write: async () => {}, read: async () => null } } })
  await empty.run('edit.paste')
  assert.match(empty.notes.at(-1).message, /no image/)
  const external = makeHarness({ clipboard: solid(3, 2, BLUE) })
  await external.run('edit.paste')
  assert.deepEqual(external.pixel(external.active(), external.active().offsetX, external.active().offsetY), [0, 0, 255, 255])
})

test('modifySelection expands and contracts in the worker without marking the document modified', async () => {
  const h = makeHarness()
  h.select({ x: 20, y: 20, width: 10, height: 10 })
  assert.equal(await commands.modifySelection(h.ctx, 'expand', 2), true)
  assert.deepEqual(h.state().selection.bounds, { x: 18, y: 18, width: 14, height: 14 })
  assert.equal(await commands.modifySelection(h.ctx, 'contract', 4), true)
  assert.deepEqual(h.state().selection.bounds, { x: 22, y: 22, width: 6, height: 6 })
  assert.equal(h.host.revision, 0)
  assert.deepEqual(h.labels(), ['Select', 'Expand', 'Contract'])
  await h.run('select.deselect')
  assert.equal(await commands.modifySelection(h.ctx, 'feather', 2), false)
})

test('placePixelsAsLayer adds an image as a new named layer above the active one', () => {
  const h = makeHarness()
  const id = commands.placePixelsAsLayer(h.ctx, solid(5, 5, BLUE), 'Pasted', { x: 3, y: 4 })
  const layer = h.state().layers.find((entry) => entry.id === id)
  assert.equal(layer.name, 'Pasted 1')
  assert.deepEqual({ x: layer.offsetX, y: layer.offsetY }, { x: 3, y: 4 })
  assert.equal(h.state().activeLayerId, id)
  assert.equal(commands.placePixelsAsLayer(h.ctx, { width: 0, height: 0, data: new Uint8ClampedArray(0) }, 'x'), null)
})

test('turnRect and turnMask agree with the pixel rotations', () => {
  const m = { width: 3, height: 2, data: new Uint8Array([1, 2, 3, 4, 5, 6]) }
  assert.deepEqual(Array.from(commands.turnMask(m, 'cw').data), [4, 1, 5, 2, 6, 3])
  assert.deepEqual(Array.from(commands.turnMask(m, 'ccw').data), [3, 6, 2, 5, 1, 4])
  assert.deepEqual(Array.from(commands.turnMask(m, '180').data), [6, 5, 4, 3, 2, 1])
  assert.deepEqual(Array.from(commands.turnMask(m, 'flip-h').data), [3, 2, 1, 6, 5, 4])
  assert.deepEqual(Array.from(commands.turnMask(m, 'flip-v').data), [4, 5, 6, 1, 2, 3])
  assert.deepEqual(commands.turnRect({ x: 1, y: 2, width: 3, height: 4 }, 'cw', 10, 8), { x: 2, y: 1, width: 4, height: 3 })
  assert.deepEqual(commands.turnRect({ x: 1, y: 2, width: 3, height: 4 }, 'ccw', 10, 8), { x: 2, y: 6, width: 4, height: 3 })
  const affine = commands.turnAffine('cw', 10, 8)
  const apply = (p) => ({ x: affine[0] * p.x + affine[2] * p.y + affine[4], y: affine[1] * p.x + affine[3] * p.y + affine[5] })
  assert.deepEqual(apply({ x: 0, y: 0 }), { x: 8, y: 0 })
  assert.deepEqual(apply({ x: 10, y: 8 }), { x: 0, y: 10 })
})
