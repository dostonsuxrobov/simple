'use strict'
// WP3 document model and store (src/advanced/document.ts, memory.ts): every transaction primitive and stroke
// session undoes and redoes to byte-identical document hashes; rollback and exceptions revert everything;
// coalescing within 1,200 ms; selection-only and active-layer changes never change the revision; strokes
// mark the document modified from the first pixel and survive Ctrl+Z / transactions mid-stroke; change
// notifications carry document-space dirty areas; layer factories, naming, flat-equivalence and the memory
// refusal policy.
const test = require('node:test')
const assert = require('node:assert/strict')
const path = require('node:path')
const crypto = require('node:crypto')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping (not disabled by --no-experimental-strip-types), '
  + 'or run node with --experimental-strip-types.'

function load(relative) {
  if (!(process.features && process.features.typescript)) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

const documentModule = load('advanced/document.ts')
const tiles = load('advanced/tiles.ts')
const memory = load('advanced/memory.ts')
const selection = load('advanced/selection.ts')
const mask = load('imaging/mask.ts')
const composite = load('advanced/composite.ts')
const { COALESCE_MS, LIMITS, TILE_SIZE, DEFAULT_LOCKS, BACKGROUND_LOCKS } = load('advanced/types.ts')

const W = 600
const H = 400

function noise(width, height, seed = 1, alpha = null) {
  const data = new Uint8ClampedArray(width * height * 4)
  let s = seed >>> 0
  for (let i = 0; i < data.length; i += 4) {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0
    data[i] = s & 255
    data[i + 1] = (s >>> 8) & 255
    data[i + 2] = (s >>> 16) & 255
    data[i + 3] = alpha === null ? s >>> 24 : alpha
  }
  return { width, height, data }
}

function makeHost(start = 0) {
  let counter = start
  const host = {
    revision: start,
    nextRevision: () => {
      counter += 1
      return counter
    },
    setRevision: (revision) => {
      host.revision = revision
    },
    currentRevision: () => host.revision,
  }
  return host
}

function makeStore(extra = {}) {
  const host = extra.host ?? makeHost()
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(noise(W, H, 1, 255)) })
  const layer = documentModule.createRasterLayer({ name: 'Layer 1', surface: tiles.surfaceFromBuffer(noise(300, 200, 2), -40, -30), offsetX: 100, offsetY: 120 })
  let time = 0
  const store = documentModule.createDocumentStore({
    width: W, height: H, ppi: 72, layers: [background, layer], host, baseLabel: 'Open', now: () => time, ...extra,
  })
  return { store, host, background, layer, tick: (ms) => { time += ms } }
}

function hashGrid(hash, grid) {
  if (!grid) {
    hash.update('none')
    return
  }
  const entries = []
  grid.forEachTile((tx, ty, data) => entries.push([tx, ty, data]))
  entries.sort((a, b) => a[1] - b[1] || a[0] - b[0])
  hash.update(`grid:${grid.channels}:${grid.defaultValue ?? ''}:${entries.length};`)
  for (const [tx, ty, data] of entries) {
    hash.update(`${tx},${ty};`)
    hash.update(Buffer.from(data.buffer, data.byteOffset, data.byteLength))
  }
}

/** Hash of everything a user could observe: metadata, every tile byte, masks, selection, canvas. */
function documentHash(store) {
  const s = store.getState()
  const hash = crypto.createHash('sha256')
  hash.update(JSON.stringify({ width: s.width, height: s.height, ppi: s.ppi, active: s.activeLayerId, target: s.editTarget }))
  for (const layer of s.layers) {
    const meta = {
      id: layer.id, kind: layer.kind, name: layer.name, visible: layer.visible, opacity: layer.opacity, blendMode: layer.blendMode,
      locks: layer.locks, clipped: layer.clipped, isBackground: layer.isBackground, offsetX: layer.offsetX, offsetY: layer.offsetY,
      adjustment: layer.adjustment, shape: layer.shape, text: layer.text,
      raster: layer.raster ? { x: layer.raster.offsetX, y: layer.raster.offsetY, key: layer.raster.specKey } : null,
      mask: layer.mask ? { x: layer.mask.offsetX, y: layer.mask.offsetY, enabled: layer.mask.enabled, linked: layer.mask.linked } : null,
    }
    hash.update(JSON.stringify(meta))
    hashGrid(hash, layer.kind === 'raster' ? layer.surface : layer.raster ? layer.raster.surface : null)
    hashGrid(hash, layer.mask ? layer.mask.surface : null)
  }
  if (s.selection) {
    hash.update(JSON.stringify(s.selection.bounds))
    hash.update(Buffer.from(s.selection.mask.data))
  } else {
    hash.update('no selection')
  }
  return hash.digest('hex')
}

/** Runs `change` as one transaction and checks undo/redo restore byte-identical documents twice over. */
function assertUndoRedo(store, label, change, options) {
  const before = documentHash(store)
  const entries = store.history.getState().entries.length
  store.transact(label, 'layer', change, options)
  const after = documentHash(store)
  assert.notEqual(after, before, `${label} changed the document`)
  assert.equal(store.history.getState().entries.length, entries + 1, `${label} recorded one step`)
  for (let round = 0; round < 2; round += 1) {
    store.history.undo()
    assert.equal(documentHash(store), before, `${label}: undo restores the document`)
    store.history.redo()
    assert.equal(documentHash(store), after, `${label}: redo restores the change`)
  }
  return after
}

function rectSelection(x, y, width, height, version = 1) {
  const m = mask.createMaskBuffer(W, H)
  mask.rasterizeRect(m, { x, y, width, height }, false)
  return selection.selectionFromMask(m, version)
}

test('store creation validates the document', () => {
  const host = makeHost()
  const layer = documentModule.createRasterLayer({ name: 'A' })
  assert.throws(() => documentModule.createDocumentStore({ width: 0, height: 10, ppi: 72, layers: [], host, baseLabel: 'Open' }), RangeError)
  assert.throws(() => documentModule.createDocumentStore({ width: 20001, height: 10, ppi: 72, layers: [], host, baseLabel: 'Open' }), /20,000/)
  assert.throws(() => documentModule.createDocumentStore({ width: 10000, height: 6000, ppi: 72, layers: [], host, baseLabel: 'Open' }), /megapixels/)
  assert.throws(() => documentModule.createDocumentStore({ width: 10, height: 10, ppi: 72, layers: [], host: {}, baseLabel: 'Open' }), TypeError)
  assert.throws(() => documentModule.createDocumentStore({ width: 10, height: 10, ppi: 72, layers: [layer, layer], host, baseLabel: 'Open' }), /share the id/)
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true })
  assert.throws(() => documentModule.createDocumentStore({ width: 10, height: 10, ppi: 72, layers: [layer, background], host, baseLabel: 'Open' }), /bottom/)
  const store = documentModule.createDocumentStore({ width: 10, height: 10, ppi: 0, layers: [background, layer], host, baseLabel: 'Advanced editor', revision: 5 })
  const state = store.getState()
  assert.equal(state.ppi, 72, 'unknown resolution falls back to 72 ppi')
  assert.equal(state.activeLayerId, layer.id, 'the top layer is active by default')
  assert.equal(state.revision, 5)
  assert.equal(store.history.getState().entries[0].label, 'Advanced editor')
  assert.ok(Object.isFrozen(LIMITS))
})

test('layer structure primitives undo and redo to byte-identical documents', () => {
  const { store, background, layer } = makeStore()
  const added = documentModule.createRasterLayer({ name: 'Layer 2', surface: tiles.surfaceFromBuffer(noise(50, 50, 3), 500, 350) })
  assertUndoRedo(store, 'New Layer', (tx) => {
    tx.insertLayer(added)
    tx.setActiveLayer(added.id)
  })
  assert.equal(store.getState().activeLayerId, added.id)
  assert.deepEqual(store.getState().layers.map((item) => item.name), ['Background', 'Layer 1', 'Layer 2'])
  // Nothing goes below the Background.
  const low = documentModule.createRasterLayer({ name: 'Low' })
  assertUndoRedo(store, 'Insert low', (tx) => tx.insertLayer(low, 0))
  assert.deepEqual(store.getState().layers.map((item) => item.name), ['Background', 'Low', 'Layer 1', 'Layer 2'])
  assertUndoRedo(store, 'Move', (tx) => tx.moveLayer(added.id, 1))
  assert.deepEqual(store.getState().layers.map((item) => item.name), ['Background', 'Layer 2', 'Low', 'Layer 1'])
  store.transact('noop', 'layer', (tx) => tx.moveLayer(background.id, 3))
  assert.equal(store.getState().layers[0].id, background.id, 'the Background never moves')
  // Removing the active layer selects the one below (Photoshop).
  store.transact('select', 'layer', (tx) => tx.setActiveLayer(layer.id))
  assertUndoRedo(store, 'Delete Layer', (tx) => tx.removeLayer(layer.id))
  assert.equal(store.getState().activeLayerId, low.id)
  store.history.undo()
  assert.equal(store.getState().activeLayerId, layer.id, 'undo restores the active layer')
  assert.throws(() => store.transact('x', 'layer', (tx) => tx.removeLayer('missing')), /does not exist/)
  assert.throws(() => store.transact('x', 'layer', (tx) => tx.insertLayer(added)), /already exists/)
})

test('layer property patches undo and redo, including linked masks and Background conversion', () => {
  const { store, background, layer } = makeStore()
  assertUndoRedo(store, 'Rename', (tx) => tx.updateLayer(layer.id, { name: '  Sky  ' }))
  assert.equal(store.getState().layers[1].name, 'Sky')
  assertUndoRedo(store, 'Opacity', (tx) => tx.updateLayer(layer.id, { opacity: 0.35, blendMode: 'multiply', visible: false }))
  assertUndoRedo(store, 'Locks', (tx) => tx.updateLayer(layer.id, { locks: { pixels: true, position: false, transparency: true } }))
  assertUndoRedo(store, 'Clip', (tx) => tx.updateLayer(layer.id, { clipped: true }))
  const layerMask = documentModule.createLayerMask({ offsetX: 7, offsetY: 9 })
  layerMask.surface.write(0, 0, { width: 2, height: 1, data: new Uint8Array([0, 100]) })
  assertUndoRedo(store, 'Add Mask', (tx) => tx.setMask(layer.id, layerMask))
  assertUndoRedo(store, 'Move', (tx) => tx.updateLayer(layer.id, { offset: { x: 130, y: 90 } }))
  const moved = store.getState().layers[1]
  assert.deepEqual([moved.offsetX, moved.offsetY, moved.mask.offsetX, moved.mask.offsetY], [130, 90, 37, -21], 'a linked mask moves with its layer')
  assertUndoRedo(store, 'Unlink', (tx) => tx.updateLayer(layer.id, { maskLinked: false, maskEnabled: false }))
  store.transact('Move', 'move', (tx) => tx.updateLayer(layer.id, { offset: { x: 0, y: 0 } }))
  assert.deepEqual([store.getState().layers[1].mask.offsetX, store.getState().layers[1].mask.offsetY], [37, -21], 'an unlinked mask stays')
  assertUndoRedo(store, 'Delete Mask', (tx) => tx.setMask(layer.id, null))
  assert.throws(() => store.transact('x', 'layer', (tx) => tx.setMask(background.id, documentModule.createLayerMask())), /Background/)
  assert.throws(() => store.transact('x', 'layer', (tx) => tx.updateLayer(layer.id, { opacity: Number.NaN })), RangeError)
  assert.throws(() => store.transact('x', 'layer', (tx) => tx.updateLayer(layer.id, { blendMode: 'plus' })), RangeError)
  assert.throws(() => store.transact('x', 'layer', (tx) => tx.updateLayer(layer.id, { isBackground: true })), /bottom/)
  assertUndoRedo(store, 'Layer from Background', (tx) => tx.updateLayer(background.id, { isBackground: false, locks: DEFAULT_LOCKS, name: 'Layer 0' }))
  assert.equal(store.getState().layers[0].isBackground, false)
  const entries = store.history.getState().entries.length
  store.transact('Same', 'layer', (tx) => tx.updateLayer(layer.id, { opacity: store.getState().layers[1].opacity }))
  assert.equal(store.history.getState().entries.length, entries, 'a patch that changes nothing records nothing')
})

test('pixel and mask edits are copy-on-write and undo/redo exactly, across tiles and negative coordinates', () => {
  const { store, layer } = makeStore()
  const before = noise(1, 1)
  assertUndoRedo(store, 'Paint', (tx) => {
    const editor = tx.editPixels(layer.id, 'pixels')
    editor.writePixels(-300, -100, noise(700, 260, 9))
    // A second editor on the same surface in the same transaction.
    const second = tx.editPixels(layer.id, 'pixels')
    second.writePixels(200, 50, noise(100, 300, 10))
    // readBefore sees the pre-edit pixels of untouched and touched tiles alike.
    const pre = second.readBefore({ x: 200, y: 50, width: 4, height: 1 })
    assert.equal(pre.width, 4)
  })
  // Clearing to transparency frees tiles; undo brings them back.
  assertUndoRedo(store, 'Clear', (tx) => {
    tx.editPixels(layer.id, 'pixels').writePixels(-1000, -1000, noise(2000, 2000, 1, 0))
  })
  assert.equal(store.getState().layers[1].surface.tileCount, 0)
  // Whole-tile replacement through setTile and through the guarded editor.surface view.
  assertUndoRedo(store, 'Tiles', (tx) => {
    const editor = tx.editPixels(layer.id, 'pixels')
    const tile = new Uint8ClampedArray(TILE_SIZE * TILE_SIZE * 4).fill(77)
    editor.surface.setTile(-1, 2, tile)
    editor.surface.write(5, 5, noise(10, 10, 4))
  })
  store.transact('Add Mask', 'layer', (tx) => tx.setMask(layer.id, documentModule.createLayerMask()))
  assertUndoRedo(store, 'Mask paint', (tx) => {
    const editor = tx.editPixels(layer.id, 'mask')
    editor.writeMask(-10, -10, { width: 300, height: 20, data: new Uint8Array(6000).fill(30) })
    assert.throws(() => editor.writePixels(0, 0, before), TypeError)
  })
  assert.throws(() => store.transact('x', 'brush', (tx) => tx.editPixels(store.getState().layers[0].id, 'mask')), /no layer mask/)
  const adjustment = documentModule.createAdjustmentLayer('Invert 1', { type: 'invert' })
  store.transact('New', 'adjustment', (tx) => tx.insertLayer(adjustment))
  assert.throws(() => store.transact('x', 'brush', (tx) => tx.editPixels(adjustment.id, 'pixels')), /adjustment layer/)
})

test('surface swaps, canvas changes and selections undo and redo exactly', () => {
  const { store, layer } = makeStore()
  const replacement = tiles.surfaceFromBuffer(noise(80, 90, 12))
  assertUndoRedo(store, 'Free Transform', (tx) => tx.replaceSurface(layer.id, 'pixels', replacement, { x: -20, y: 300 }))
  assert.equal(store.getState().layers[1].surface, replacement, 'swapped by reference, no copy')
  store.transact('Add Mask', 'layer', (tx) => tx.setMask(layer.id, documentModule.createLayerMask()))
  assertUndoRedo(store, 'Mask transform', (tx) => tx.replaceSurface(layer.id, 'mask', tiles.maskFromBuffer({ width: 2, height: 2, data: new Uint8Array([0, 50, 100, 150]) }, 255), { x: 3, y: 4 }))
  assertUndoRedo(store, 'Rectangular Marquee', (tx) => tx.setSelection(rectSelection(10, 20, 30, 40)))
  assertUndoRedo(store, 'Select All', (tx) => tx.setSelection(selection.selectAll(W, H, 7)))
  assertUndoRedo(store, 'Deselect', (tx) => tx.setSelection(null))
  store.transact('Select', 'selection', (tx) => tx.setSelection(rectSelection(1, 1, 5, 5)))
  assertUndoRedo(store, 'Canvas Size', (tx) => tx.setCanvas(800, 300, 300))
  const state = store.getState()
  assert.deepEqual([state.width, state.height, state.ppi, state.selection], [800, 300, 300, null], 'a size change drops the selection')
  store.history.undo()
  assert.deepEqual(store.getState().selection.bounds, { x: 1, y: 1, width: 5, height: 5 })
  assert.throws(() => store.transact('x', 'image', (tx) => tx.setCanvas(0, 5)), RangeError)
  assert.throws(() => store.transact('x', 'selection', (tx) => tx.setSelection({ mask: mask.createMaskBuffer(2, 2), bounds: { x: 0, y: 0, width: 1, height: 1 }, version: 1 })), RangeError)
})

test('shape and adjustment layers update their specs undoably', () => {
  const { store } = makeStore()
  const spec = { kind: 'ellipse', x1: 0, y1: 0, x2: 40, y2: 30, fill: { r: 255, g: 0, b: 0, a: 255 }, stroke: null, strokeWidth: 0, cornerRadius: 0, arrowHeads: 'none', transform: [1, 0, 0, 1, 10, 10] }
  const shape = documentModule.createShapeLayer('Ellipse 1', spec)
  store.transact('Ellipse Tool', 'shape', (tx) => tx.insertLayer(shape))
  const moved = assertUndoRedo(store, 'Move', (tx) => tx.updateLayer(shape.id, { shape: { ...spec, transform: [1, 0, 0, 1, 15, 12] } }))
  const after = store.getState().layers[2]
  assert.equal(after.raster.surface, shape.raster.surface, 'a whole-pixel move only shifts the raster cache')
  assert.deepEqual([after.raster.offsetX - shape.raster.offsetX, after.raster.offsetY - shape.raster.offsetY], [5, 2])
  assert.ok(moved)
  assertUndoRedo(store, 'Resize', (tx) => tx.updateLayer(shape.id, { shape: { ...spec, x2: 80 } }))
  assert.notEqual(store.getState().layers[2].raster.surface, shape.raster.surface, 'a resize re-rasterizes')
  const levels = documentModule.createAdjustmentLayer('Levels 1', { type: 'posterize', levels: 4 })
  store.transact('New', 'adjustment', (tx) => tx.insertLayer(levels))
  assertUndoRedo(store, 'Posterize', (tx) => tx.updateLayer(levels.id, { adjustment: { type: 'posterize', levels: 8 } }))
  assert.throws(() => store.transact('x', 'layer', (tx) => tx.updateLayer(levels.id, { shape: spec })), /not a shape layer/)
})

test('rollback and exceptions revert everything, pixels included, and record nothing', () => {
  const { store, host, layer } = makeStore()
  const before = documentHash(store)
  const revision = host.revision
  const result = store.transact('Doomed', 'brush', (tx) => {
    tx.editPixels(layer.id, 'pixels').writePixels(0, 0, noise(400, 300, 5))
    tx.updateLayer(layer.id, { opacity: 0.1 })
    tx.insertLayer(documentModule.createRasterLayer({ name: 'Temp' }))
    tx.setSelection(rectSelection(0, 0, 10, 10))
    tx.rollback()
  })
  assert.equal(result, undefined)
  assert.equal(documentHash(store), before)
  assert.equal(store.history.getState().entries.length, 1)
  assert.equal(host.revision, revision)
  assert.throws(() => store.transact('Broken', 'brush', (tx) => {
    tx.editPixels(layer.id, 'pixels').writePixels(-50, -50, noise(100, 100, 6))
    throw new Error('tool failed')
  }), /tool failed/)
  assert.equal(documentHash(store), before)
  assert.equal(store.history.getState().entries.length, 1)
  // An editor from a finished transaction is inert.
  let leaked = null
  store.transact('Paint', 'brush', (tx) => {
    leaked = tx.editPixels(layer.id, 'pixels')
    leaked.writePixels(0, 0, noise(2, 2, 7))
  })
  const painted = documentHash(store)
  leaked.writePixels(0, 0, noise(50, 50, 8))
  leaked.surface.write(0, 0, noise(50, 50, 8))
  assert.equal(documentHash(store), painted)
  let leakedTx = null
  store.transact('Nothing', 'layer', (tx) => { leakedTx = tx })
  assert.throws(() => leakedTx.insertLayer(documentModule.createRasterLayer({ name: 'Late' })), /already finished/)
  assert.equal(store.history.getState().entries.length, 2, 'an empty transaction records nothing')
})

test('nested transactions join the outer one', () => {
  const { store, layer } = makeStore()
  store.transact('Outer', 'layer', (tx) => {
    tx.updateLayer(layer.id, { opacity: 0.5 })
    store.transact('Inner', 'layer', (inner) => inner.updateLayer(layer.id, { name: 'Joined' }))
  })
  const entries = store.history.getState().entries
  assert.deepEqual(entries.map((entry) => entry.label), ['Open', 'Outer'])
  assert.equal(store.getState().layers[1].name, 'Joined')
})

test('transactions with the same coalescing key merge within 1,200 ms', () => {
  const { store, layer, tick, host } = makeStore()
  const start = documentHash(store)
  store.transact('Opacity', 'layer', (tx) => tx.updateLayer(layer.id, { opacity: 0.9 }), { coalesceKey: 'opacity' })
  tick(500)
  store.transact('Opacity', 'layer', (tx) => tx.updateLayer(layer.id, { opacity: 0.8 }), { coalesceKey: 'opacity' })
  tick(COALESCE_MS)
  store.transact('Opacity', 'layer', (tx) => tx.updateLayer(layer.id, { opacity: 0.7 }), { coalesceKey: 'opacity' })
  let entries = store.history.getState().entries
  assert.equal(entries.length, 2, 'all three merged (each within 1,200 ms of the previous)')
  assert.equal(entries[1].label, 'Opacity')
  tick(COALESCE_MS + 1)
  store.transact('Opacity', 'layer', (tx) => tx.updateLayer(layer.id, { opacity: 0.6 }), { coalesceKey: 'opacity' })
  tick(10)
  store.transact('Nudge', 'move', (tx) => tx.updateLayer(layer.id, { offset: { x: 101, y: 120 } }), { coalesceKey: 'nudge' })
  tick(10)
  store.transact('Nudge', 'move', (tx) => tx.updateLayer(layer.id, { offset: { x: 102, y: 120 } }), { coalesceKey: 'nudge' })
  entries = store.history.getState().entries
  assert.deepEqual(entries.map((entry) => entry.label), ['Open', 'Opacity', 'Opacity', 'Nudge'])
  store.history.undo()
  assert.equal(store.getState().layers[1].offsetX, 100, 'one undo reverts both nudges')
  store.history.undo()
  assert.equal(store.getState().layers[1].opacity, 0.7)
  store.history.undo()
  assert.equal(documentHash(store), start)
  assert.equal(host.revision, 0)
})

test('selection-only and active-layer changes never change the revision', () => {
  const { store, host, background, layer } = makeStore()
  store.transact('Paint', 'brush', (tx) => tx.editPixels(layer.id, 'pixels').writePixels(0, 0, noise(4, 4, 3)))
  const saved = host.revision
  assert.ok(saved > 0)
  store.transact('Rectangular Marquee', 'marquee-rect', (tx) => tx.setSelection(rectSelection(5, 5, 50, 50)), { affectsOutput: false })
  assert.equal(host.revision, saved)
  assert.equal(store.getState().revision, saved)
  // Even when the caller forgets affectsOutput: false.
  store.transact('Deselect', 'selection', (tx) => tx.setSelection(null))
  assert.equal(host.revision, saved)
  const entries = store.history.getState().entries.length
  store.transact('Select layer', 'layer', (tx) => tx.setActiveLayer(background.id))
  assert.equal(host.revision, saved)
  assert.equal(store.history.getState().entries.length, entries, 'choosing a layer is not a history step')
  store.history.undo()
  store.history.undo()
  assert.equal(host.revision, saved, 'undoing selection steps keeps the saved revision')
  // A mislabelled real edit still gets a revision.
  store.transact('Opacity', 'layer', (tx) => tx.updateLayer(layer.id, { opacity: 0.2 }), { affectsOutput: false })
  assert.notEqual(host.revision, saved)
})

test('revisions follow undo and redo so the Modified badge is always right', () => {
  const { store, host, layer } = makeStore()
  const opened = host.revision
  store.transact('A', 'brush', (tx) => tx.editPixels(layer.id, 'pixels').writePixels(0, 0, noise(4, 4, 3)))
  const afterA = host.revision
  store.transact('B', 'layer', (tx) => tx.updateLayer(layer.id, { name: 'B' }))
  const afterB = host.revision
  assert.ok(opened < afterA && afterA < afterB)
  store.history.undo()
  assert.equal(host.revision, afterA)
  store.history.undo()
  assert.equal(host.revision, opened)
  store.history.redo()
  assert.equal(host.revision, afterA)
  assert.equal(store.getState().revision, afterA)
  // A new edit after undo gets a fresh revision (never reuses the redone one).
  store.transact('C', 'layer', (tx) => tx.updateLayer(layer.id, { name: 'C' }))
  assert.ok(host.revision > afterB)
})

test('stroke sessions: one step, modified from the first pixel, exact undo/redo, cancel', () => {
  const { store, host, layer } = makeStore()
  const before = documentHash(store)
  const opened = host.revision
  const area = { x: -60, y: -40, width: 20, height: 20 }
  const original = layer.surface.read(area)
  const stroke = store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush')
  assert.ok(store.hasOpenStroke())
  assert.equal(host.revision, opened, 'starting a stroke is not an edit')
  stroke.editor.writePixels(-60, -40, noise(20, 20, 1))
  const provisional = host.revision
  assert.notEqual(provisional, opened, 'modified from the first pixel')
  assert.notDeepEqual(layer.surface.read(area).data, original.data)
  assert.deepEqual(stroke.editor.readBefore(area).data, original.data, 'readBefore serves the pre-stroke pixels')
  for (let i = 0; i < 50; i += 1) stroke.editor.writePixels(-60 + i * 8, -40 + i * 3, noise(20, 20, i + 2))
  stroke.commit()
  assert.ok(!store.hasOpenStroke())
  assert.ok(host.revision > provisional, 'the finished stroke has its own revision')
  const after = documentHash(store)
  assert.deepEqual(store.history.getState().entries.map((entry) => entry.label), ['Open', 'Brush Tool'])
  store.history.undo()
  assert.equal(documentHash(store), before)
  assert.equal(host.revision, opened)
  store.history.redo()
  assert.equal(documentHash(store), after)
  // Writes after commit are ignored.
  stroke.editor.writePixels(0, 0, noise(30, 30, 99))
  assert.equal(documentHash(store), after)
  // Cancel restores pixels and the revision and records nothing.
  const revision = host.revision
  const cancelled = store.beginStroke(layer.id, 'pixels', 'Eraser', 'eraser')
  cancelled.editor.writePixels(0, 0, noise(300, 300, 5, 0))
  assert.notEqual(documentHash(store), after)
  cancelled.cancel()
  assert.equal(documentHash(store), after)
  assert.equal(host.revision, revision)
  assert.equal(store.history.getState().entries.length, 2)
  // A stroke without pixels records nothing.
  store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush').commit()
  assert.equal(store.history.getState().entries.length, 2)
})

test('Ctrl+Z, a transaction or another stroke during a stroke first commits it; nothing is lost', () => {
  const { store, host, layer, background } = makeStore()
  const before = documentHash(store)
  const stroke = store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush')
  stroke.editor.writePixels(0, 0, noise(50, 50, 1))
  const painted = documentHash(store)
  store.history.undo()
  assert.equal(documentHash(store), before, 'Ctrl+Z mid-stroke undoes exactly that stroke')
  store.history.redo()
  assert.equal(documentHash(store), painted)
  stroke.editor.writePixels(100, 100, noise(5, 5, 2))
  assert.equal(documentHash(store), painted, 'the committed stroke cannot be changed afterwards')
  const second = store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush')
  second.editor.writePixels(10, 10, noise(5, 5, 3))
  store.transact('Rename', 'layer', (tx) => tx.updateLayer(background.id, { name: 'Base' }))
  assert.deepEqual(store.history.getState().entries.map((entry) => entry.label), ['Open', 'Brush Tool', 'Brush Tool', 'Rename'])
  const third = store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush')
  third.editor.writePixels(20, 20, noise(5, 5, 4))
  store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush')
  assert.equal(store.history.getState().entries.length, 5, 'starting a stroke commits the open one')
  store.commitStroke()
  assert.ok(!store.hasOpenStroke())
  assert.ok(host.revision > 0)
  assert.throws(() => store.transact('x', 'brush', () => store.beginStroke(layer.id, 'pixels', 'B', 'brush')), /inside a transaction/)
})

test('change notifications carry document-space dirty areas and flags', () => {
  const { store, layer } = makeStore()
  const changes = []
  const unsubscribe = store.subscribe((change) => changes.push(change))
  store.transact('Paint', 'brush', (tx) => tx.editPixels(layer.id, 'pixels').writePixels(10, 20, noise(5, 5, 1)))
  assert.deepEqual(changes[0].dirty, [{ x: 110, y: 140, width: 5, height: 5 }], 'layer-local rect shifted by the layer offset')
  assert.equal(changes[0].history, true)
  const stroke = store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush')
  stroke.editor.writePixels(0, 0, noise(3, 3, 2))
  assert.deepEqual(changes[1], { structure: false, selection: false, history: false, dirty: [{ x: 100, y: 120, width: 3, height: 3 }] })
  stroke.commit()
  assert.equal(changes[2].history, true)
  store.transact('Invert', 'adjustment', (tx) => tx.insertLayer(documentModule.createAdjustmentLayer('Invert 1', { type: 'invert' })))
  assert.equal(changes[3].dirty, 'all')
  assert.equal(changes[3].structure, true)
  store.transact('Marquee', 'marquee-rect', (tx) => tx.setSelection(rectSelection(0, 0, 4, 4)))
  assert.equal(changes[4].selection, true)
  assert.deepEqual(changes[4].dirty, [])
  store.history.undo()
  assert.equal(changes[5].history, true)
  assert.equal(changes[5].selection, true)
  const pixelVersion = store.getState().pixelVersion
  store.history.undo()
  assert.ok(store.getState().pixelVersion > pixelVersion)
  unsubscribe()
  store.transact('Rename', 'layer', (tx) => tx.updateLayer(layer.id, { name: 'X' }))
  assert.equal(changes.length, 7)
  // A throwing subscriber does not break the store.
  const quiet = console.error
  console.error = () => {}
  try {
    store.subscribe(() => { throw new Error('bad listener') })
    store.transact('Rename', 'layer', (tx) => tx.updateLayer(layer.id, { name: 'Y' }))
  } finally {
    console.error = quiet
  }
  assert.equal(store.getState().layers[1].name, 'Y')
})

test('layer factories, names and helpers', () => {
  const background = documentModule.createRasterLayer({ name: '', isBackground: true, opacity: 0.2, blendMode: 'multiply' })
  assert.equal(background.name, 'Background')
  assert.deepEqual(background.locks, BACKGROUND_LOCKS)
  assert.deepEqual([background.opacity, background.blendMode, background.mask, background.clipped], [1, 'normal', null, false])
  const layer = documentModule.createRasterLayer({ name: 'Layer 1', offsetX: 3.6, opacity: 2 })
  assert.deepEqual([layer.offsetX, layer.opacity, layer.blendMode], [4, 1, 'normal'])
  assert.notEqual(documentModule.newLayerId(), documentModule.newLayerId())
  const state = { layers: [background, layer, { ...layer, id: 'b', name: 'Layer 7' }, { ...layer, id: 'c', name: 'Levels 1' }, { ...layer, id: 'd', name: 'Layer 1 copy' }] }
  assert.equal(documentModule.nextLayerName(state, 'Layer'), 'Layer 8')
  assert.equal(documentModule.nextLayerName(state, 'Levels'), 'Levels 2')
  assert.equal(documentModule.nextLayerName(state, 'Curves'), 'Curves 1')
  assert.equal(documentModule.nextLayerName(state, 'Layer 1 copy'), 'Layer 1 copy 2')
  assert.equal(documentModule.nextLayerName(state, 'Layer 7 copy'), 'Layer 7 copy')
  // Duplicates are deep copies; a copy of the Background is a normal layer.
  const filled = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(noise(10, 10, 1, 255)) })
  const copy = documentModule.duplicateLayer(filled)
  assert.equal(copy.name, 'Background copy')
  assert.equal(copy.isBackground, false)
  assert.deepEqual(copy.locks, DEFAULT_LOCKS)
  assert.notEqual(copy.surface, filled.surface)
  assert.deepEqual(copy.surface.read({ x: 0, y: 0, width: 10, height: 10 }).data, filled.surface.read({ x: 0, y: 0, width: 10, height: 10 }).data)
  // Flat equivalence (design 7.3).
  const doc = { width: 10, height: 10, layers: [filled] }
  assert.ok(documentModule.isFlatEquivalent(doc))
  assert.ok(!documentModule.isFlatEquivalent({ ...doc, layers: [filled, copy] }))
  assert.ok(!documentModule.isFlatEquivalent({ ...doc, layers: [{ ...filled, opacity: 0.99 }] }))
  assert.ok(!documentModule.isFlatEquivalent({ ...doc, layers: [{ ...filled, mask: documentModule.createLayerMask({ enabled: false }) }] }), 'even a disabled mask')
  assert.ok(!documentModule.isFlatEquivalent({ ...doc, layers: [{ ...filled, offsetX: 1 }] }), 'pixels outside the canvas')
  assert.ok(!documentModule.isFlatEquivalent({ ...doc, layers: [documentModule.createAdjustmentLayer('Invert', { type: 'invert' })] }))
  assert.ok(documentModule.isFlatEquivalent({ ...doc, layers: [documentModule.createRasterLayer({ name: 'Empty' })] }))
  assert.deepEqual(documentModule.layerContentBounds({ ...filled, offsetX: 5, offsetY: -2 }), { x: 5, y: -2, width: 10, height: 10 })
  // Masks from selections hide everything outside the selection.
  const fromSelection = documentModule.maskFromSelection(rectSelection(10, 20, 30, 40))
  assert.equal(fromSelection.surface.defaultValue, 0)
  assert.deepEqual(fromSelection.surface.contentBounds(), { x: 10, y: 20, width: 30, height: 40 })
  assert.equal(documentModule.maskFromSelection(null).surface.defaultValue, 255)
  const adjustment = documentModule.createAdjustmentLayer('Levels 1', { type: 'invert' }, fromSelection)
  assert.equal(adjustment.mask, fromSelection)
  // Edit blockers.
  assert.match(documentModule.pixelEditBlocker(null), /Select a layer/)
  assert.match(documentModule.pixelEditBlocker(adjustment), /adjustment layer/)
  assert.equal(documentModule.pixelEditBlocker(adjustment, 'mask'), null)
  assert.match(documentModule.pixelEditBlocker({ ...layer, locks: { ...DEFAULT_LOCKS, pixels: true } }), /locked/)
  assert.match(documentModule.pixelEditBlocker({ ...layer, visible: false }), /hidden/)
  assert.equal(documentModule.pixelEditBlocker(layer), null)
  // Shapes rasterize without a canvas (pure rasterizer); rasterizing makes a pixel layer.
  const shape = documentModule.createShapeLayer('Rectangle 1', { kind: 'rectangle', x1: 0, y1: 0, x2: 10, y2: 4, fill: { r: 0, g: 0, b: 0, a: 255 }, stroke: null, strokeWidth: 0, cornerRadius: 0, arrowHeads: 'none', transform: [1, 0, 0, 1, 2, 3] })
  assert.ok(shape.raster.surface.contentBounds())
  const rasterized = documentModule.rasterizeVectorLayer(shape)
  assert.equal(rasterized.kind, 'raster')
  assert.deepEqual(documentModule.layerContentBounds(rasterized), documentModule.layerContentBounds(shape))
})

test('text layers accept a ready raster cache (PSD import)', async () => {
  const raster = documentModule.rasterCacheFrom(noise(4, 2, 1, 255), 10, 20, 'psd-key')
  const text = await documentModule.createTextLayer('Title', {
    text: 'Hello', style: { fontFamily: 'Arial', fontSize: 12, fontWeight: 400, italic: false, underline: false, color: { r: 0, g: 0, b: 0 }, align: 'left', lineHeight: 1.2, letterSpacing: 0 },
    boxWidth: null, transform: [1, 0, 0, 1, 10, 20],
  }, raster)
  assert.equal(text.kind, 'text')
  assert.equal(text.raster, raster)
  assert.deepEqual(documentModule.layerContentBounds(text), { x: 10, y: 20, width: 4, height: 2 })
})

test('memory usage counts layers, masks and history; the refusal policy trims history before refusing', () => {
  const host = makeHost()
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(noise(512, 256, 1, 255)) })
  const store = documentModule.createDocumentStore({ width: 512, height: 256, ppi: 72, layers: [background], host, baseLabel: 'Open' })
  const tile = TILE_SIZE * TILE_SIZE * 4
  let usage = store.memoryUsage()
  assert.equal(usage.layers, 2 * tile)
  assert.equal(usage.masks, 0)
  assert.equal(usage.history, 0)
  assert.equal(usage.budget, LIMITS.documentBudgetBytes)
  for (let i = 0; i < 3; i += 1) {
    const stroke = store.beginStroke(background.id, 'pixels', 'Brush Tool', 'brush')
    stroke.editor.writePixels(0, 0, noise(512, 10, i + 2, 255))
    stroke.commit()
  }
  usage = store.memoryUsage()
  assert.ok(usage.history >= 6 * tile)
  assert.equal(memory.estimateRasterBytes(257, 1), 2 * tile)
  assert.equal(memory.estimateMaskBytes(257, 1), 2 * tile / 4)
  assert.equal(memory.formatBytes(1536 * 1024 * 1024), '1.5 GB')
  assert.equal(memory.formatBytes(412 * 1024 * 1024), '412 MB')
  assert.deepEqual(memory.ensureMemory(store, 1024), { ok: true, freedCaches: 0, freedHistory: 0 })
  // Pretend the Simple canvas holds almost everything: history is trimmed first (newest step kept).
  memory.setExternalBytes(LIMITS.documentBudgetBytes - usage.layers - 2 * tile - 1000)
  try {
    // An operation that would not fit even without history is refused without losing undo steps.
    const hopeless = memory.ensureMemory(store, 100 * tile, 'paste this image')
    assert.equal(hopeless.ok, false)
    assert.match(hopeless.message, /paste this image/)
    assert.equal(store.history.getState().entries.length, 4, 'no undo step was trimmed for a refusal')
    const check = memory.ensureMemory(store, 0, 'add a layer')
    assert.equal(check.ok, true)
    assert.ok(check.freedHistory > 0)
    assert.equal(store.history.getState().entries.length, 2, 'only the newest step is left')
    assert.ok(store.history.canUndo())
    const refused = memory.ensureMemory(store, 10 * tile, 'duplicate this layer')
    assert.equal(refused.ok, false)
    assert.match(refused.message, /enough memory to duplicate this layer/)
    assert.ok(store.history.canUndo(), 'the last step stays undoable even when refusing')
  } finally {
    memory.setExternalBytes(0)
  }
})

test('a disposed store refuses changes', () => {
  const { store, layer } = makeStore()
  const stroke = store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush')
  stroke.editor.writePixels(0, 0, noise(2, 2, 1))
  store.dispose()
  assert.ok(!store.hasOpenStroke(), 'disposing commits the open stroke first')
  assert.throws(() => store.transact('x', 'layer', () => {}), /closed/)
  assert.throws(() => store.beginStroke(layer.id, 'pixels', 'B', 'brush'), /closed/)
  store.dispose()
  assert.ok(composite.compositeRect(store.getState(), { x: 0, y: 0, width: 2, height: 2 }))
})
