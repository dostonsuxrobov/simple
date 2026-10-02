'use strict'
// WP8 Simple mode: the byte-budgeted undo store. Swap semantics (undo/redo restore exact pixels and the
// state on each side), tile patches captured once per tile, zero-byte inverse geometry steps, byte
// accounting, trimming that always keeps the newest step, and the "50 MP keeps 10+ strokes" property
// that the old count-based store (1 step at 24 MP) failed.
const test = require('node:test')
const assert = require('node:assert/strict')
const path = require('node:path')
const crypto = require('node:crypto')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping.'

function load(relative) {
  if (!(process.features && process.features.typescript)) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

const H = load('simple/simpleHistory.ts')

/** An in-memory HistorySurface with exact geometry ops (what the canvas adapter does with drawImage). */
function memorySurface(width, height, seed = 1) {
  let w = width
  let h = height
  let data = new Uint8ClampedArray(w * h * 4)
  let state = seed
  for (let i = 0; i < data.length; i += 1) {
    state = (state * 1103515245 + 12345) >>> 0
    data[i] = state >>> 24
  }
  const reads = { count: 0 }
  const surface = {
    get width() { return w },
    get height() { return h },
    reads,
    read(rect) {
      reads.count += 1
      const out = new Uint8ClampedArray(rect.width * rect.height * 4)
      for (let y = 0; y < rect.height; y += 1) {
        const from = ((rect.y + y) * w + rect.x) * 4
        out.set(data.subarray(from, from + rect.width * 4), y * rect.width * 4)
      }
      return { width: rect.width, height: rect.height, data: out }
    },
    write(x, y, pixels) {
      for (let row = 0; row < pixels.height; row += 1) {
        data.set(pixels.data.subarray(row * pixels.width * 4, (row + 1) * pixels.width * 4), ((y + row) * w + x) * 4)
      }
    },
    replace(pixels) {
      w = pixels.width
      h = pixels.height
      data = new Uint8ClampedArray(pixels.data)
    },
    transform(op) {
      const out = new Uint8ClampedArray(data.length)
      const rotate = op === 'rotate-cw' || op === 'rotate-ccw'
      const nw = rotate ? h : w
      const nh = rotate ? w : h
      for (let y = 0; y < h; y += 1) {
        for (let x = 0; x < w; x += 1) {
          let tx
          let ty
          if (op === 'rotate-cw') { tx = h - 1 - y; ty = x }
          else if (op === 'rotate-ccw') { tx = y; ty = w - 1 - x }
          else if (op === 'flip-h') { tx = w - 1 - x; ty = y }
          else { tx = x; ty = h - 1 - y }
          out.set(data.subarray((y * w + x) * 4, (y * w + x) * 4 + 4), (ty * nw + tx) * 4)
        }
      }
      w = nw
      h = nh
      data = out
    },
    paint(rect, value) {
      for (let y = rect.y; y < rect.y + rect.height; y += 1) {
        for (let x = rect.x; x < rect.x + rect.width; x += 1) data.fill(value, (y * w + x) * 4, (y * w + x) * 4 + 4)
      }
    },
    hash() { return `${w}x${h}:${crypto.createHash('sha256').update(data).digest('hex')}` },
  }
  return surface
}

const state = (revision, extra = {}) => ({ revision, hasAlpha: false, pristine: revision === 0, ...extra })

function stroke(stacks, surface, rect, revision, value = 7) {
  const recorder = H.createPatchRecorder(surface)
  recorder.touch(rect)
  surface.paint(rect, value)
  const entry = recorder.finish('Brush', state(revision - 1), state(revision))
  H.pushEntry(stacks, entry)
  return entry
}

test('a patch entry captures each touched tile once and restores exact pixels on undo and redo', () => {
  const surface = memorySurface(600, 300)
  const stacks = { undo: [], redo: [] }
  const original = surface.hash()
  const recorder = H.createPatchRecorder(surface)
  recorder.touch({ x: 250, y: 10, width: 20, height: 20 })
  recorder.touch({ x: 255, y: 12, width: 4, height: 4 })
  assert.equal(recorder.tileCount, 2, 'the rectangle spans the first two tile columns')
  surface.paint({ x: 250, y: 10, width: 20, height: 20 }, 0)
  recorder.touch({ x: 590, y: 290, width: 50, height: 50 })
  assert.equal(recorder.tileCount, 3, 'edge tiles are clipped to the image')
  surface.paint({ x: 590, y: 290, width: 10, height: 10 }, 255)
  const entry = recorder.finish('Brush', state(0), state(1))
  assert.equal(entry.kind, 'patch')
  assert.equal(H.entryBytes(entry), (256 * 256 + 256 * 256 + 88 * 44) * 4)
  H.pushEntry(stacks, entry)
  const edited = surface.hash()
  assert.notEqual(edited, original)
  const undone = H.undoEntry(stacks, surface)
  assert.equal(undone, entry)
  assert.equal(surface.hash(), original)
  assert.deepEqual(undone.before, state(0))
  assert.equal(stacks.redo.length, 1)
  const redone = H.redoEntry(stacks, surface)
  assert.equal(surface.hash(), edited)
  assert.deepEqual(redone.after, state(1))
  assert.equal(H.undoEntry(stacks, surface), entry)
  assert.equal(surface.hash(), original)
  assert.equal(H.undoEntry(stacks, surface), null, 'nothing left to undo')
})

test('an empty recorder makes no entry', () => {
  const surface = memorySurface(10, 10)
  const recorder = H.createPatchRecorder(surface)
  recorder.touch({ x: 20, y: 20, width: 5, height: 5 })
  assert.equal(recorder.finish('Brush', state(0), state(1)), null)
})

test('snapshot entries swap whole images of different sizes', () => {
  const surface = memorySurface(64, 32)
  const stacks = { undo: [], redo: [] }
  const original = surface.hash()
  const before = surface.read(H.fullRect(surface))
  surface.replace({ width: 16, height: 8, data: new Uint8ClampedArray(16 * 8 * 4).fill(9) })
  const cropped = surface.hash()
  H.pushEntry(stacks, H.snapshotEntry('Crop', before, state(0, { hasAlpha: true }), state(1)))
  assert.equal(H.historyBytes(stacks), 64 * 32 * 4)
  const undone = H.undoEntry(stacks, surface)
  assert.equal(surface.hash(), original)
  assert.equal(undone.before.hasAlpha, true, 'the state of the restored side travels with the entry')
  assert.equal(H.historyBytes(stacks), 16 * 8 * 4, 'the entry now holds the cropped side')
  H.redoEntry(stacks, surface)
  assert.equal(surface.hash(), cropped)
})

test('quarter turns and flips cost no bytes and undo exactly through their inverse', () => {
  const surface = memorySurface(5, 3)
  const stacks = { undo: [], redo: [] }
  const original = surface.hash()
  for (const op of ['rotate-cw', 'flip-h', 'rotate-ccw', 'flip-v', 'rotate-cw']) {
    surface.transform(op)
    H.pushEntry(stacks, H.geometryEntry(op, op, state(stacks.undo.length), state(stacks.undo.length + 1)))
  }
  assert.equal(H.historyBytes(stacks), 0)
  assert.deepEqual([surface.width, surface.height], [3, 5])
  while (H.undoEntry(stacks, surface));
  assert.equal(surface.hash(), original)
  assert.equal(H.inverseGeometry('rotate-cw'), 'rotate-ccw')
  assert.equal(H.inverseGeometry('flip-v'), 'flip-v')
  assert.deepEqual(H.geometrySize(5, 3, 'rotate-ccw'), { width: 3, height: 5 })
})

test('a new step drops the redo branch', () => {
  const surface = memorySurface(300, 300)
  const stacks = { undo: [], redo: [] }
  stroke(stacks, surface, { x: 0, y: 0, width: 10, height: 10 }, 1)
  stroke(stacks, surface, { x: 20, y: 0, width: 10, height: 10 }, 2)
  H.undoEntry(stacks, surface)
  assert.equal(stacks.redo.length, 1)
  stroke(stacks, surface, { x: 40, y: 0, width: 10, height: 10 }, 3)
  assert.equal(stacks.redo.length, 0)
  assert.equal(stacks.undo.length, 2)
})

test('the byte budget trims the oldest steps but always keeps the newest', () => {
  const surface = memorySurface(512, 512)
  const stacks = { undo: [], redo: [] }
  const tile = 256 * 256 * 4
  const budget = tile * 3
  for (let index = 0; index < 6; index += 1) {
    const recorder = H.createPatchRecorder(surface)
    recorder.touch({ x: (index % 2) * 256, y: 0, width: 1, height: 1 })
    H.pushEntry(stacks, recorder.finish('Brush', state(index), state(index + 1)), budget)
  }
  assert.equal(stacks.undo.length, 3)
  assert.ok(H.historyBytes(stacks) <= budget)
  assert.deepEqual(stacks.undo.map((entry) => entry.after.revision), [4, 5, 6], 'the newest steps survive')
  const huge = H.snapshotEntry('Resize', surface.read(H.fullRect(surface)), state(6), state(7))
  H.pushEntry(stacks, huge, budget)
  assert.deepEqual(stacks.undo, [huge], 'one step larger than the budget is still kept')
})

test('a 50 MP image keeps far more than 10 undoable brush strokes', () => {
  // The old store allowed floor(160 MiB / (50 MP * 4)) = 0 -> 1 step. Patches cost only the tiles touched.
  const fake = { width: 10_000, height: 5_000 }
  const tileBytes = 256 * 256 * 4
  const surface = {
    get width() { return fake.width },
    get height() { return fake.height },
    read: (rect) => ({ width: rect.width, height: rect.height, data: new Uint8ClampedArray(rect.width * rect.height * 4) }),
    write() {},
    replace() {},
    transform() {},
  }
  const stacks = { undo: [], redo: [] }
  for (let index = 0; index < 40; index += 1) {
    const recorder = H.createPatchRecorder(surface)
    // A long diagonal stroke across ~12 tiles.
    for (let step = 0; step < 12; step += 1) recorder.touch({ x: (300 * index) % 7000 + step * 200, y: step * 200, width: 40, height: 40 })
    H.pushEntry(stacks, recorder.finish('Brush', state(index), state(index + 1)))
  }
  assert.equal(stacks.undo.length, 40)
  assert.ok(H.historyBytes(stacks) <= H.SIMPLE_HISTORY_BUDGET_BYTES)
  assert.ok(H.historyBytes(stacks) / 40 <= 24 * tileBytes, 'each stroke costs its tiles, not the image')
})

test('patches refuse to apply to an image of another size, and leave the stacks intact', () => {
  const surface = memorySurface(300, 300)
  const stacks = { undo: [], redo: [] }
  stroke(stacks, surface, { x: 0, y: 0, width: 10, height: 10 }, 1)
  surface.replace({ width: 10, height: 10, data: new Uint8ClampedArray(400) })
  assert.throws(() => H.undoEntry(stacks, surface), /no longer matches/)
  assert.equal(stacks.undo.length, 1)
  assert.equal(stacks.redo.length, 0)
})

test('clipRect clips to the image and rejects empty areas', () => {
  assert.deepEqual(H.clipRect({ x: -5.5, y: 3.2, width: 20, height: 4 }, 10, 10), { x: 0, y: 3, width: 10, height: 5 })
  assert.equal(H.clipRect({ x: 20, y: 0, width: 5, height: 5 }, 10, 10), null)
})
