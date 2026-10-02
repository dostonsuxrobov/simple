'use strict'
// WP3 history (src/advanced/history.ts): navigation order, revisions handed to the host, redo-tail drop,
// Ctrl+Alt+Z toggling, count and byte budgets (base and newest step always kept, exact byte accounting),
// coalescing windows, reset and dispose; plus copy-on-write byte accounting through the document store.
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

const history = load('advanced/history.ts')
const tiles = load('advanced/tiles.ts')
const documentModule = load('advanced/document.ts')
const { COALESCE_MS, LIMITS, TILE_SIZE } = load('advanced/types.ts')

const TILE_BYTES = TILE_SIZE * TILE_SIZE * 4

function makeHost() {
  let counter = 0
  const host = {
    revision: 0,
    log: [],
    nextRevision: () => {
      counter += 1
      return counter
    },
    setRevision: (revision) => {
      host.revision = revision
      host.log.push(revision)
    },
    currentRevision: () => host.revision,
  }
  return host
}

/** A history whose ops are plain markers; `applied` records (entry label, direction). */
function makeHistory(options = {}) {
  const host = makeHost()
  const applied = []
  const events = []
  const controller = history.createHistory({
    baseLabel: 'Open',
    baseRevision: 0,
    host,
    apply: (entry, direction) => applied.push(`${direction}:${entry.label}`),
    beforeNavigate: () => events.push('before'),
    afterNavigate: (revision) => events.push(`after:${revision}`),
    ...options,
  })
  return { controller, host, applied, events }
}

/** An op that keeps `tileCount` RGBA tiles alive (exact byte size known). */
function tilesOp(tileCount) {
  const stored = new Map()
  for (let i = 0; i < tileCount; i += 1) stored.set(i, new Uint8ClampedArray(TILE_BYTES))
  return { kind: 'tiles', layerId: 'L', target: 'pixels', surface: tiles.createSurface(), stored }
}

function canvasOp() {
  return { kind: 'canvas', before: { width: 1, height: 1, ppi: 72 }, after: { width: 2, height: 2, ppi: 72 } }
}

let clock = 0
function record(controller, label, ops = [canvasOp()], revision = null, coalesceKey = null) {
  clock += 10_000
  const state = controller.getState()
  const before = controller.currentRevision()
  return controller.record({ label, icon: 'brush', ops, revisionBefore: before, revisionAfter: revision ?? before + 1, coalesceKey, time: clock })
}

test('the base entry, undo, redo and jumps apply entries in order and set host revisions', () => {
  const { controller, host, applied, events } = makeHistory()
  let state = controller.getState()
  assert.equal(state.entries.length, 1)
  assert.equal(state.entries[0].label, 'Open')
  assert.equal(state.cursor, 0)
  assert.ok(!controller.canUndo() && !controller.canRedo())
  record(controller, 'A', [canvasOp()], 1)
  record(controller, 'B', [canvasOp()], 2)
  record(controller, 'C', [canvasOp()], 3)
  state = controller.getState()
  assert.deepEqual(state.entries.map((entry) => entry.label), ['Open', 'A', 'B', 'C'])
  assert.equal(state.cursor, 3)
  assert.equal(controller.getState(), state, 'state snapshots are cached until something changes')
  controller.undo()
  assert.deepEqual(applied, ['undo:C'])
  assert.equal(host.revision, 2)
  controller.jumpTo(0)
  assert.deepEqual(applied, ['undo:C', 'undo:B', 'undo:A'])
  assert.equal(host.revision, 0)
  assert.ok(!controller.canUndo() && controller.canRedo())
  controller.jumpTo(2)
  assert.deepEqual(applied.slice(3), ['redo:A', 'redo:B'])
  assert.equal(host.revision, 2)
  controller.redo()
  assert.equal(host.revision, 3)
  assert.ok(events.includes('before') && events.includes('after:3'))
  controller.jumpTo(99)
  assert.equal(controller.getState().cursor, 3, 'jumps clamp to the last entry')
})

test('a new step drops the redo tail and its bytes', () => {
  const { controller } = makeHistory()
  record(controller, 'A', [tilesOp(1)])
  record(controller, 'B', [tilesOp(2)])
  controller.undo()
  const undone = controller.getState().totalBytes
  record(controller, 'C', [tilesOp(1)])
  const state = controller.getState()
  assert.deepEqual(state.entries.map((entry) => entry.label), ['Open', 'A', 'C'])
  assert.ok(!controller.canRedo())
  assert.equal(state.totalBytes, state.entries.reduce((sum, entry) => sum + entry.bytes, 0))
  assert.ok(state.totalBytes < undone)
})

test('Ctrl+Alt+Z toggles between the current and the previous state', () => {
  const { controller, host } = makeHistory()
  record(controller, 'A', [canvasOp()], 1)
  record(controller, 'B', [canvasOp()], 2)
  controller.toggleLast()
  assert.equal(controller.getState().cursor, 1)
  controller.toggleLast()
  assert.equal(controller.getState().cursor, 2)
  controller.toggleLast()
  assert.equal(controller.getState().cursor, 1)
  assert.equal(host.revision, 1)
  // From the base state it steps forward.
  controller.jumpTo(0)
  controller.toggleLast()
  assert.equal(controller.getState().cursor, 1)
})

test('the step budget evicts the oldest steps, keeps the base and the newest, and relabels the base', () => {
  const { controller, host } = makeHistory({ maxEntries: 3 })
  for (const [label, revision] of [['A', 1], ['B', 2], ['C', 3], ['D', 4], ['E', 5]]) record(controller, label, [canvasOp()], revision)
  const state = controller.getState()
  assert.deepEqual(state.entries.map((entry) => entry.label), [history.TRIMMED_LABEL, 'C', 'D', 'E'])
  assert.ok(state.trimmed)
  assert.equal(state.maxEntries, 3)
  // The base now stands for the state after the last evicted step (B).
  controller.jumpTo(0)
  assert.equal(host.revision, 2)
  // The Photoshop default: 50 steps.
  const fresh = makeHistory().controller
  for (let i = 0; i < 60; i += 1) record(fresh, `S${i}`)
  assert.equal(fresh.getState().entries.length, LIMITS.historyMaxEntries + 1)
  assert.equal(fresh.getState().entries[LIMITS.historyMaxEntries].label, 'S59')
})

test('the byte budget evicts oldest steps first, keeps an oversized newest step, and accounting is exact', () => {
  const entryBytes = (tileCount) => history.ENTRY_OVERHEAD_BYTES + history.SMALL_OP_BYTES + tileCount * TILE_BYTES
  const budget = entryBytes(2) + entryBytes(3)
  const { controller } = makeHistory({ budgetBytes: budget })
  record(controller, 'A', [tilesOp(2)])
  record(controller, 'B', [tilesOp(3)])
  let state = controller.getState()
  assert.equal(state.totalBytes, budget)
  assert.deepEqual(state.entries.map((entry) => entry.bytes), [0, entryBytes(2), entryBytes(3)])
  record(controller, 'C', [tilesOp(1)])
  state = controller.getState()
  assert.deepEqual(state.entries.map((entry) => entry.label), [history.TRIMMED_LABEL, 'B', 'C'])
  assert.equal(state.totalBytes, entryBytes(3) + entryBytes(1))
  // A single step larger than the whole budget stays undoable.
  record(controller, 'Huge', [tilesOp(20)])
  state = controller.getState()
  assert.deepEqual(state.entries.map((entry) => entry.label), [history.TRIMMED_LABEL, 'Huge'])
  assert.equal(state.totalBytes, entryBytes(20))
  assert.ok(controller.canUndo())
  // Lowering the budget evicts immediately; raising it does not bring steps back.
  record(controller, 'Small', [tilesOp(1)])
  controller.setBudget(entryBytes(1), 50)
  state = controller.getState()
  assert.deepEqual(state.entries.map((entry) => entry.label), [history.TRIMMED_LABEL, 'Small'])
  controller.setBudget(LIMITS.historyBudgetBytes, LIMITS.historyMaxEntries)
  assert.equal(controller.getState().entries.length, 2)
})

test('with the cursor low, redo steps are dropped from the end before anything else', () => {
  const entryBytes = (tileCount) => history.ENTRY_OVERHEAD_BYTES + history.SMALL_OP_BYTES + tileCount * TILE_BYTES
  const { controller } = makeHistory()
  record(controller, 'A', [tilesOp(1)])
  record(controller, 'B', [tilesOp(1)])
  record(controller, 'C', [tilesOp(1)])
  controller.jumpTo(1)
  controller.setBudget(entryBytes(1) * 2, 50)
  const state = controller.getState()
  assert.deepEqual(state.entries.map((entry) => entry.label), ['Open', 'A', 'B'])
  assert.equal(state.cursor, 1)
  assert.ok(!state.trimmed, 'nothing before the cursor was evicted')
})

test('coalescing targets only the newest step with the same key inside the window', () => {
  const { controller, host } = makeHistory()
  const entry = controller.record({ label: 'Opacity', icon: 'layer', ops: [canvasOp()], revisionBefore: 0, revisionAfter: 1, coalesceKey: 'opacity', time: 1000 })
  assert.equal(controller.coalesceTarget('opacity', 1000 + COALESCE_MS), entry)
  assert.equal(controller.coalesceTarget('opacity', 1001 + COALESCE_MS), null)
  assert.equal(controller.coalesceTarget('nudge', 1100), null)
  assert.equal(controller.coalesceTarget(null, 1100), null)
  assert.equal(controller.coalesceTarget('opacity', 999), null, 'clock going backwards never merges')
  controller.extend(entry, [tilesOp(1)], 2, 1500)
  assert.equal(entry.revisionAfter, 2)
  assert.equal(entry.ops.length, 2)
  assert.equal(controller.getState().totalBytes, entry.bytes)
  assert.equal(controller.coalesceTarget('opacity', 1500 + COALESCE_MS), entry, 'the window restarts at the last merge')
  controller.undo()
  assert.equal(host.revision, 0)
  assert.equal(controller.coalesceTarget('opacity', 1600), null, 'not after an undo')
  assert.throws(() => controller.extend(entry, [], 3, 1700), /newest/)
})

test('reset keeps the current revision; dispose makes the history inert', () => {
  const { controller, host, events } = makeHistory()
  record(controller, 'A', [canvasOp()], 7)
  controller.reset('Advanced editor')
  const state = controller.getState()
  assert.deepEqual(state.entries.map((entry) => entry.label), ['Advanced editor'])
  assert.equal(controller.currentRevision(), 7)
  assert.ok(events.includes('before'))
  controller.dispose()
  assert.ok(!controller.canUndo())
  assert.throws(() => controller.undo(), /closed/)
  assert.throws(() => record(controller, 'B'), /closed/)
  assert.equal(host.revision, 0, 'reset and dispose do not change the host revision')
})

test('a failing apply leaves history usable', () => {
  let fail = true
  const { controller } = makeHistory({
    apply: () => {
      if (fail) throw new Error('boom')
    },
  })
  record(controller, 'A')
  assert.throws(() => controller.undo(), /boom/)
  fail = false
  controller.undo()
  assert.equal(controller.getState().cursor, 0)
})

test('opBytes counts what history keeps alive for every operation kind', () => {
  const layer = documentModule.createRasterLayer({ name: 'L', surface: tiles.surfaceFromBuffer({ width: 1, height: 1, data: new Uint8ClampedArray([1, 2, 3, 4]) }) })
  const layerBytes = TILE_BYTES + history.SMALL_OP_BYTES
  assert.equal(history.opBytes({ kind: 'insert', layer, index: 0 }, true), history.SMALL_OP_BYTES)
  assert.equal(history.opBytes({ kind: 'insert', layer, index: 0 }, false), layerBytes)
  assert.equal(history.opBytes({ kind: 'remove', layer, index: 0 }, true), layerBytes)
  assert.equal(history.opBytes({ kind: 'remove', layer, index: 0 }, false), history.SMALL_OP_BYTES)
  const other = tiles.surfaceFromBuffer({ width: 300, height: 1, data: new Uint8ClampedArray(1200).fill(9) })
  assert.equal(history.opBytes({ kind: 'surface', layerId: 'L', target: 'pixels', other, otherOffset: { x: 0, y: 0 } }, true), history.SMALL_OP_BYTES + 2 * TILE_BYTES)
  const moved = { ...layer, offsetX: 5 }
  assert.equal(history.opBytes({ kind: 'layer', before: layer, after: moved }, true), history.SMALL_OP_BYTES, 'shared surfaces cost nothing')
  const repainted = { ...layer, surface: other }
  assert.equal(history.opBytes({ kind: 'layer', before: layer, after: repainted }, true), history.SMALL_OP_BYTES + TILE_BYTES)
  assert.equal(history.opBytes({ kind: 'selection', before: { kind: 'none' }, after: { kind: 'all' } }, true), 32)
  assert.ok(!history.opAffectsOutput({ kind: 'selection' }))
  assert.ok(!history.opAffectsOutput({ kind: 'active' }))
  assert.ok(history.opAffectsOutput({ kind: 'layer' }))
})

test('copy-on-write strokes cost one tile copy per touched tile, matching the history total', () => {
  const host = makeHost()
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true,
    surface: tiles.surfaceFromBuffer({ width: 600, height: 300, data: new Uint8ClampedArray(600 * 300 * 4).fill(200) }) })
  const store = documentModule.createDocumentStore({ width: 600, height: 300, ppi: 72, layers: [background], host, baseLabel: 'Open' })
  const stroke = store.beginStroke(background.id, 'pixels', 'Brush Tool', 'brush')
  // A line across three tiles, painted in many small dabs.
  for (let x = 10; x < 590; x += 4) stroke.editor.writePixels(x, 100, { width: 4, height: 4, data: new Uint8ClampedArray(64).fill(10) })
  stroke.commit()
  const state = store.history.getState()
  assert.equal(state.entries.length, 2)
  const expected = history.ENTRY_OVERHEAD_BYTES + history.SMALL_OP_BYTES + 3 * TILE_BYTES
  assert.equal(state.entries[1].bytes, expected)
  assert.equal(state.totalBytes, expected)
  store.history.undo()
  assert.equal(store.history.getState().totalBytes, expected, 'an undone stroke keeps the same number of tiles')
  assert.equal(store.memoryUsage().history, expected)
})

// ---------------------------------------------------------------------------------------------
// Randomised consistency: long seeded sequences of edits, strokes, rollbacks, coalesced changes and
// undo / redo / jumps / Ctrl+Alt+Z. Every history state must come back byte-identical, and the host
// revision must always match the state it lands on.
// ---------------------------------------------------------------------------------------------

const crypto = require('node:crypto')
const selectionModule = load('advanced/selection.ts')
const maskModule = load('imaging/mask.ts')

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

/** Everything history restores (the active layer is deliberately not a history step). */
function historyHash(store) {
  const s = store.getState()
  const hash = crypto.createHash('sha256')
  hash.update(JSON.stringify({ width: s.width, height: s.height, ppi: s.ppi }))
  for (const layer of s.layers) {
    hash.update(JSON.stringify({
      id: layer.id, name: layer.name, visible: layer.visible, opacity: layer.opacity, blendMode: layer.blendMode, locks: layer.locks,
      clipped: layer.clipped, isBackground: layer.isBackground, offsetX: layer.offsetX, offsetY: layer.offsetY,
      mask: layer.mask ? [layer.mask.offsetX, layer.mask.offsetY, layer.mask.enabled, layer.mask.linked] : null,
    }))
    hashGrid(hash, layer.kind === 'raster' ? layer.surface : null)
    hashGrid(hash, layer.mask ? layer.mask.surface : null)
  }
  hash.update(s.selection ? JSON.stringify(s.selection.bounds) + Buffer.from(s.selection.mask.data).toString('base64') : 'none')
  return hash.digest('hex')
}

function fuzzNoise(width, height, random, transparentShare = 0.2) {
  const data = new Uint8ClampedArray(width * height * 4)
  const seed = Math.floor(random() * 0xffffffff)
  let s = seed >>> 0
  const clear = random() < transparentShare
  for (let i = 0; i < data.length; i += 4) {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0
    data[i] = s & 255
    data[i + 1] = (s >>> 8) & 255
    data[i + 2] = (s >>> 16) & 255
    data[i + 3] = clear ? 0 : s >>> 24
  }
  return { width, height, data }
}

for (const seed of [1, 2, 3]) {
  test(`random edit and undo sequences restore every state exactly (seed ${seed})`, () => {
    let state = seed * 0x9e3779b1
    const random = () => {
      state = (state + 0x6d2b79f5) >>> 0
      let t = state
      t = Math.imul(t ^ (t >>> 15), t | 1)
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296
    }
    const pick = (list) => list[Math.floor(random() * list.length)]
    const int = (lo, hi) => lo + Math.floor(random() * (hi - lo + 1))
    let counter = 0
    const host = { revision: 0, nextRevision: () => (counter += 1), setRevision: (r) => { host.revision = r }, currentRevision: () => host.revision }
    let clock = 0
    const W = 520
    const H = 300
    const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(fuzzNoise(W, H, random, 0)) })
    const store = documentModule.createDocumentStore({ width: W, height: H, ppi: 72, layers: [background], host, baseLabel: 'Open', now: () => clock, historyMaxEntries: 1000 })
    // hashes[i] / revisions[i]: the expected document and host revision of history state i.
    const hashes = [historyHash(store)]
    const revisions = [host.revision]
    // Hash of the live document while a stroke is open; that state becomes a history step when the stroke
    // is committed, explicitly or implicitly by the next transaction, stroke or navigation.
    let openStroke = null
    const ids = () => store.history.getState().entries.map((entry) => entry.id)
    /** Updates the expectations after an action, from which history entries are new. */
    const track = (before, action) => {
      const after = ids()
      let common = 0
      while (common < before.length && common < after.length && before[common] === after[common]) common += 1
      const { cursor } = store.history.getState()
      const entries = store.history.entriesForDebug()
      hashes.length = Math.min(hashes.length, common)
      revisions.length = hashes.length
      let next = common
      if (openStroke && next < after.length && !(store.hasOpenStroke() && openStroke.stillOpen)) {
        // The first new entry is the stroke that was open before this action.
        hashes[next] = openStroke.hash
        revisions[next] = entries[next].revisionAfter
        next += 1
      }
      for (; next < after.length; next += 1) {
        assert.equal(next, cursor, `${action}: a new entry other than the committed stroke is the current state`)
        hashes[next] = historyHash(store)
        revisions[next] = host.revision
      }
      if (after.length === before.length && common === after.length && action !== 'undo' && action !== 'redo'
        && action !== 'jump' && action !== 'toggle' && !store.hasOpenStroke()) {
        if (action === 'coalesced') {
          // Merged into the current step.
          hashes[cursor] = historyHash(store)
          revisions[cursor] = host.revision
        } else {
          // Nothing recorded means nothing changed (no-ops, cancelled strokes, rollbacks, choosing a layer).
          assert.equal(historyHash(store), hashes[cursor], `${action}: no history step, no change`)
          assert.equal(host.revision, revisions[cursor], `${action}: no history step, same revision`)
        }
      }
      openStroke = store.hasOpenStroke() ? { hash: historyHash(store), stillOpen: true } : null
    }
    const verify = (label) => {
      const { cursor } = store.history.getState()
      assert.equal(historyHash(store), hashes[cursor], `${label}: state ${cursor} restored exactly`)
      assert.equal(host.revision, revisions[cursor], `${label}: revision of state ${cursor}`)
      assert.equal(store.getState().revision, host.revision)
    }
    const pixelLayers = () => store.getState().layers.filter((layer) => layer.kind === 'raster')
    const movable = () => store.getState().layers.filter((layer) => !layer.isBackground)
    const operations = {
      paint() {
        const layer = pick(pixelLayers())
        store.transact('Paint', 'brush', (tx) => {
          const editor = tx.editPixels(layer.id, 'pixels')
          for (let i = int(1, 3); i > 0; i -= 1) editor.writePixels(int(-300, W), int(-200, H), fuzzNoise(int(1, 300), int(1, 200), random))
        })
      },
      stroke() {
        const layer = pick(pixelLayers())
        const stroke = store.beginStroke(layer.id, 'pixels', 'Brush Tool', 'brush')
        for (let i = int(1, 12); i > 0; i -= 1) stroke.editor.writePixels(int(-50, W), int(-50, H), fuzzNoise(int(1, 40), int(1, 40), random))
        if (random() < 0.2) stroke.cancel()
        else if (random() < 0.3) return // left open: the next transaction or navigation must commit it
        else stroke.commit()
      },
      insert() {
        if (store.getState().layers.length > 8) return
        const layer = documentModule.createRasterLayer({ name: 'L', surface: tiles.surfaceFromBuffer(fuzzNoise(int(1, 300), int(1, 300), random)), offsetX: int(-100, W), offsetY: int(-100, H) })
        store.transact('New Layer', 'layer', (tx) => { tx.insertLayer(layer, int(0, 9)); tx.setActiveLayer(layer.id) })
      },
      remove() {
        const candidates = movable()
        if (!candidates.length) return
        store.transact('Delete Layer', 'layer', (tx) => tx.removeLayer(pick(candidates).id))
      },
      move() {
        const candidates = movable()
        if (!candidates.length) return
        store.transact('Arrange', 'layer', (tx) => tx.moveLayer(pick(candidates).id, int(0, 9)))
      },
      props() {
        const layer = pick(store.getState().layers)
        const patch = pick([{ opacity: random() }, { blendMode: pick(['multiply', 'screen', 'overlay', 'normal']) }, { visible: random() < 0.5 },
          { name: `N${int(0, 99)}` }, { clipped: random() < 0.5 }, { offset: { x: int(-200, 200), y: int(-200, 200) } }])
        store.transact('Properties', 'layer', (tx) => tx.updateLayer(layer.id, patch))
      },
      coalesced() {
        const layer = pick(store.getState().layers)
        clock += int(0, 1500)
        store.transact('Opacity', 'layer', (tx) => tx.updateLayer(layer.id, { opacity: random() }), { coalesceKey: `opacity:${layer.id}` })
      },
      mask() {
        const layer = pick(movable())
        if (!layer) return
        if (!layer.mask) {
          store.transact('Add Mask', 'layer', (tx) => tx.setMask(layer.id, documentModule.createLayerMask({ offsetX: int(-50, 50), offsetY: int(-50, 50) })))
        } else if (random() < 0.3) {
          store.transact('Delete Mask', 'layer', (tx) => tx.setMask(layer.id, null))
        } else {
          store.transact('Mask', 'brush', (tx) => tx.editPixels(layer.id, 'mask').writeMask(int(-100, W), int(-100, H), { width: 64, height: 32, data: new Uint8Array(64 * 32).fill(int(0, 255)) }))
        }
      },
      select() {
        const s = store.getState()
        let next = null
        if (random() < 0.7) {
          const m = maskModule.createMaskBuffer(s.width, s.height)
          maskModule.rasterizeEllipse(m, { x: int(-20, s.width), y: int(-20, s.height), width: int(1, 200), height: int(1, 200) }, true)
          next = selectionModule.selectionFromMask(m, int(1, 1000))
        }
        store.transact('Marquee', 'marquee-rect', (tx) => tx.setSelection(next))
      },
      swap() {
        const layer = pick(pixelLayers())
        store.transact('Free Transform', 'transform', (tx) => tx.replaceSurface(layer.id, 'pixels', tiles.surfaceFromBuffer(fuzzNoise(int(1, 200), int(1, 200), random)), { x: int(-100, 300), y: int(-100, 200) }))
      },
      canvas() {
        store.transact('Canvas Size', 'image', (tx) => tx.setCanvas(int(200, 700), int(150, 400), pick([72, 150, 300])))
      },
      rollback() {
        const layer = pick(pixelLayers())
        const before = historyHash(store)
        store.transact('Doomed', 'brush', (tx) => {
          tx.editPixels(layer.id, 'pixels').writePixels(0, 0, fuzzNoise(100, 100, random))
          tx.updateLayer(layer.id, { opacity: 0.5 })
          tx.rollback()
        })
        assert.equal(historyHash(store), before, 'rollback leaves no trace')
      },
      select_layer() {
        store.transact('Select', 'layer', (tx) => tx.setActiveLayer(pick(store.getState().layers).id))
      },
    }
    const names = Object.keys(operations)
    for (let step = 0; step < 260; step += 1) {
      const roll = random()
      const before = ids()
      if (roll < 0.18 && (store.history.canUndo() || store.hasOpenStroke())) {
        store.history.undo()
        track(before, 'undo')
        verify(`step ${step} undo`)
      } else if (roll < 0.26 && store.history.canRedo()) {
        store.history.redo()
        track(before, 'redo')
        verify(`step ${step} redo`)
      } else if (roll < 0.3) {
        store.history.jumpTo(int(0, store.history.getState().entries.length - 1))
        track(before, 'jump')
        verify(`step ${step} jump`)
      } else if (roll < 0.33) {
        store.history.toggleLast()
        track(before, 'toggle')
        verify(`step ${step} toggle`)
      } else {
        const name = pick(names)
        if (openStroke) openStroke.stillOpen = false
        operations[name]()
        if (store.hasOpenStroke() && random() < 0.4) {
          // Save and Close commit an open stroke explicitly.
          track(before, name)
          const committed = ids()
          store.commitStroke()
          track(committed, 'commit')
        } else {
          track(before, name)
        }
        if (!store.hasOpenStroke()) verify(`step ${step} ${name}`)
      }
    }
    // Walk the whole history both ways.
    const before = ids()
    store.commitStroke()
    track(before, 'commit')
    const last = store.history.getState().entries.length - 1
    for (let index = last; index >= 0; index -= 1) {
      store.history.jumpTo(index)
      verify(`walk back to ${index}`)
    }
    for (let index = 0; index <= last; index += 1) {
      store.history.jumpTo(index)
      verify(`walk forward to ${index}`)
    }
  })
}
