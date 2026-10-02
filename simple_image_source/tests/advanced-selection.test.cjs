'use strict'
// WP4 Advanced selection model (src/advanced/selection.ts): selection ops on the current selection,
// immutability, history snapshots (none / all / region) restoring exactly, translation, inverse,
// coverage reads and layer-alpha selections.
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

const selection = load('advanced/selection.ts')
const mask = load('imaging/mask.ts')

const W = 64
const H = 48

function shape(draw) {
  const m = mask.createMaskBuffer(W, H)
  draw(m)
  return m
}

const rect = (x, y, width, height, antiAlias = false) => shape((m) => mask.rasterizeRect(m, { x, y, width, height }, antiAlias))
const value = (sel, x, y) => (sel ? sel.mask.data[y * W + x] : 0)

test('selectionFromMask wraps a mask with tight bounds; empty masks are no selection', () => {
  const m = rect(10, 5, 8, 6)
  const sel = selection.selectionFromMask(m, 3)
  assert.equal(sel.mask, m, 'the mask is taken over, not copied')
  assert.deepEqual(sel.bounds, { x: 10, y: 5, width: 8, height: 6 })
  assert.equal(sel.version, 3)
  assert.equal(selection.selectionFromMask(mask.createMaskBuffer(W, H), 4), null)
  assert.throws(() => selection.selectionFromMask({ width: 2, height: 2, data: new Uint8Array(3) }, 1), RangeError)
})

test('selectAll covers the document and is recognised as such', () => {
  const all = selection.selectAll(W, H, 1)
  assert.deepEqual(all.bounds, { x: 0, y: 0, width: W, height: H })
  assert.ok(all.mask.data.every((v) => v === 255))
  assert.ok(selection.isSelectAll(all))
  assert.ok(!selection.isSelectAll(null))
  assert.ok(!selection.isSelectAll(selection.selectionFromMask(rect(0, 0, W, H - 1), 2)))
  const almost = selection.selectAll(W, H, 1)
  almost.mask.data[W * H - 1] = 254
  assert.ok(!selection.isSelectAll(almost))
  assert.throws(() => selection.selectAll(0, 5, 1), RangeError)
})

test('selection operations follow their truth tables', () => {
  // current = left square, shape = right square overlapping it.
  const current = selection.selectionFromMask(rect(10, 10, 20, 20), 1)
  const other = rect(20, 10, 20, 20)
  const probes = { onlyA: [12, 15], both: [25, 15], onlyB: [35, 15], neither: [50, 40] }
  const check = (sel, expected) => {
    for (const [name, [x, y]] of Object.entries(probes)) assert.equal(value(sel, x, y), expected[name] ? 255 : 0, name)
  }
  const replaced = selection.applySelectionOp(current, other, 'replace', 2)
  check(replaced, { onlyA: false, both: true, onlyB: true, neither: false })
  assert.deepEqual(replaced.bounds, { x: 20, y: 10, width: 20, height: 20 })
  const added = selection.applySelectionOp(current, other, 'add', 3)
  check(added, { onlyA: true, both: true, onlyB: true, neither: false })
  assert.deepEqual(added.bounds, { x: 10, y: 10, width: 30, height: 20 })
  const subtracted = selection.applySelectionOp(current, other, 'subtract', 4)
  check(subtracted, { onlyA: true, both: false, onlyB: false, neither: false })
  assert.deepEqual(subtracted.bounds, { x: 10, y: 10, width: 10, height: 20 })
  const intersected = selection.applySelectionOp(current, other, 'intersect', 5)
  check(intersected, { onlyA: false, both: true, onlyB: false, neither: false })
  assert.deepEqual(intersected.bounds, { x: 20, y: 10, width: 10, height: 20 })
  assert.deepEqual([replaced.version, added.version, subtracted.version, intersected.version], [2, 3, 4, 5])
})

test('partial coverage combines with exact rounding', () => {
  const a = mask.createMaskBuffer(W, H, 0)
  a.data[0] = 200
  a.data[1] = 100
  const b = mask.createMaskBuffer(W, H, 0)
  b.data[0] = 100
  b.data[1] = 200
  const current = selection.selectionFromMask(a, 1)
  assert.deepEqual([...selection.applySelectionOp(current, b, 'add', 2).mask.data.slice(0, 2)], [200, 200])
  assert.deepEqual([...selection.applySelectionOp(current, b, 'subtract', 2).mask.data.slice(0, 2)], [Math.round(200 * 155 / 255), Math.round(100 * 55 / 255)])
  assert.deepEqual([...selection.applySelectionOp(current, b, 'intersect', 2).mask.data.slice(0, 2)], [Math.round(200 * 100 / 255), Math.round(100 * 200 / 255)])
})

test('without a current selection: add = replace, subtract and intersect select nothing', () => {
  const s = rect(5, 5, 4, 4)
  assert.deepEqual(selection.applySelectionOp(null, s, 'add', 1).bounds, { x: 5, y: 5, width: 4, height: 4 })
  assert.equal(selection.applySelectionOp(null, s, 'subtract', 1), null)
  assert.equal(selection.applySelectionOp(null, s, 'intersect', 1), null)
  // Empty results become no selection; an empty shape keeps the selection for add / subtract.
  const current = selection.selectionFromMask(rect(5, 5, 4, 4), 1)
  assert.equal(selection.applySelectionOp(current, rect(5, 5, 4, 4), 'subtract', 2), null)
  assert.equal(selection.applySelectionOp(current, rect(40, 30, 4, 4), 'intersect', 2), null)
  const empty = mask.createMaskBuffer(W, H)
  assert.equal(selection.applySelectionOp(current, empty, 'add', 2), current)
  assert.equal(selection.applySelectionOp(current, empty, 'subtract', 2), current)
  assert.equal(selection.applySelectionOp(current, empty, 'replace', 2), null)
  assert.equal(selection.applySelectionOp(current, empty, 'intersect', 2), null)
  assert.throws(() => selection.applySelectionOp(current, mask.createMaskBuffer(W, H + 1), 'add', 2), RangeError)
  assert.throws(() => selection.applySelectionOp(current, s, 'xor', 2), RangeError)
})

test('operations never write into the selection or shape they were given', () => {
  const currentMask = rect(10, 10, 20, 20, true)
  const shapeMask = rect(15.5, 12.25, 30, 10, true)
  const current = selection.selectionFromMask(currentMask, 1)
  const before = Uint8Array.from(currentMask.data)
  const shapeBefore = Uint8Array.from(shapeMask.data)
  for (const op of ['replace', 'add', 'subtract', 'intersect']) {
    const result = selection.applySelectionOp(current, shapeMask, op, 2)
    assert.notEqual(result.mask, currentMask)
    assert.notEqual(result.mask, shapeMask)
    result.mask.data.fill(7)
    assert.deepEqual(currentMask.data, before, op)
    assert.deepEqual(shapeMask.data, shapeBefore, op)
  }
})

test('snapshots of none, all and a region restore exactly', () => {
  assert.deepEqual(selection.snapshotSelection(null), { kind: 'none' })
  assert.equal(selection.restoreSelection({ kind: 'none' }, W, H, 9), null)
  const all = selection.selectAll(W, H, 1)
  assert.deepEqual(selection.snapshotSelection(all), { kind: 'all' })
  const restoredAll = selection.restoreSelection({ kind: 'all' }, W, H, 9)
  assert.deepEqual(restoredAll.mask.data, all.mask.data)
  assert.deepEqual(restoredAll.bounds, all.bounds)
  assert.equal(restoredAll.version, 9)
  // An anti-aliased ellipse minus a lasso: arbitrary partial values.
  const m = shape((target) => mask.rasterizeEllipse(target, { x: 7.3, y: 4.6, width: 40.2, height: 30.9 }, true))
  const cut = shape((target) => mask.rasterizePolygon(target, [{ x: 20, y: 10 }, { x: 35.5, y: 22 }, { x: 18, y: 30 }], true))
  const region = selection.applySelectionOp(selection.selectionFromMask(m, 1), cut, 'subtract', 2)
  const snapshot = selection.snapshotSelection(region)
  assert.equal(snapshot.kind, 'region')
  assert.deepEqual(snapshot.rect, region.bounds)
  assert.equal(snapshot.data.length, region.bounds.width * region.bounds.height, 'only the bounds region is stored')
  assert.equal(selection.snapshotBytes(snapshot), snapshot.data.byteLength + 64)
  const restored = selection.restoreSelection(snapshot, W, H, 10)
  assert.deepEqual(restored.mask.data, region.mask.data)
  assert.deepEqual(restored.bounds, region.bounds)
  assert.equal(restored.version, 10)
  // The snapshot is a copy: later edits to the live mask do not leak into history.
  region.mask.data.fill(0)
  assert.deepEqual(selection.restoreSelection(snapshot, W, H, 11).mask.data, restored.mask.data)
})

test('restoring into a smaller document clips; damaged snapshots are rejected', () => {
  const sel = selection.selectionFromMask(rect(40, 30, 20, 10), 1)
  const snapshot = selection.snapshotSelection(sel)
  const clipped = selection.restoreSelection(snapshot, 50, 35, 2)
  assert.deepEqual(clipped.bounds, { x: 40, y: 30, width: 10, height: 5 })
  assert.equal(clipped.mask.width, 50)
  assert.equal(selection.restoreSelection(snapshot, 30, 30, 3), null)
  assert.throws(() => selection.restoreSelection({ kind: 'region', rect: { x: 0, y: 0, width: 4, height: 4 }, data: new Uint8Array(15) }, W, H, 1), RangeError)
})

test('translateSelection moves by whole pixels and clips at the canvas edge', () => {
  const sel = selection.selectionFromMask(rect(10, 10, 6, 4), 1)
  const moved = selection.translateSelection(sel, 5.4, -3.6, 2)
  assert.deepEqual(moved.bounds, { x: 15, y: 6, width: 6, height: 4 })
  assert.equal(value(moved, 15, 6), 255)
  assert.equal(value(moved, 10, 10), 0)
  assert.equal(moved.version, 2)
  const clipped = selection.translateSelection(sel, 50, 0, 3)
  assert.deepEqual(clipped.bounds, { x: 60, y: 10, width: 4, height: 4 })
  assert.equal(selection.translateSelection(sel, 100, 0, 4), null)
  const still = selection.translateSelection(sel, 0, 0, 5)
  assert.equal(still.mask, sel.mask)
  assert.equal(still.version, 5)
})

test('inverse, coverage reads and selections from layer alpha', () => {
  const sel = selection.selectionFromMask(rect(0, 0, 32, H), 1)
  const inverse = selection.invertSelection(sel, 2)
  assert.deepEqual(inverse.bounds, { x: 32, y: 0, width: 32, height: H })
  assert.equal(value(sel, 40, 5), 0, 'the original is unchanged')
  assert.equal(selection.invertSelection(selection.selectAll(W, H, 1), 3), null)
  // Coverage over a rect that leaves the canvas: 0 outside with a selection, 255 everywhere without.
  const cover = selection.selectionCoverage(sel, { x: 30, y: -2, width: 4, height: 3 })
  assert.deepEqual([...cover.data], [0, 0, 0, 0, 0, 0, 0, 0, 255, 255, 0, 0])
  assert.ok(selection.selectionCoverage(null, { x: -5, y: -5, width: 3, height: 2 }).data.every((v) => v === 255))
  // Layer alpha at an offset (Ctrl+click on a layer thumbnail).
  const pixels = { width: 3, height: 2, data: Uint8ClampedArray.from([9, 9, 9, 255, 9, 9, 9, 0, 9, 9, 9, 128, 9, 9, 9, 64, 9, 9, 9, 0, 9, 9, 9, 255]) }
  const fromAlpha = selection.selectionFromAlpha(pixels, 62, 46, W, H, 7)
  // Clipped to columns 62..63; column 63 holds alpha 0 in both rows, so the tight bounds are one wide.
  assert.deepEqual(fromAlpha.bounds, { x: 62, y: 46, width: 1, height: 2 })
  assert.deepEqual([value(fromAlpha, 62, 46), value(fromAlpha, 63, 46), value(fromAlpha, 62, 47), value(fromAlpha, 63, 47)], [255, 0, 64, 0])
  assert.equal(selection.selectionFromAlpha(pixels, 100, 0, W, H, 8), null)
  assert.equal(selection.selectionFromAlpha({ width: 1, height: 1, data: new Uint8ClampedArray(4) }, 0, 0, W, H, 9), null)
})
