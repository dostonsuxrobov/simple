'use strict'
// WP8 Simple mode: crop-box constraints (ratio lock on corner and edge handles, Alt about the centre,
// bounds and minimum size, Shift-style locked ratio), aspect presets, nudges, integer rounding, and the
// straighten helpers (angle folding, level-by-line, rotated frame, auto-crop inscribed box, drags kept
// inside the turned image).
const test = require('node:test')
const assert = require('node:assert/strict')
const path = require('node:path')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping.'

function load(relative) {
  if (!(process.features && process.features.typescript)) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

const C = load('simple/cropMath.ts')

const bounds = { width: 1000, height: 600 }
const free = { ratio: null, fromCenter: false, bounds }

function near(actual, expected, epsilon = 1e-6, message) {
  assert.ok(Math.abs(actual - expected) <= epsilon, `${message || ''} expected ${expected}, got ${actual}`)
}

function inside(rect, b = bounds) {
  return rect.x >= -1e-9 && rect.y >= -1e-9 && rect.x + rect.width <= b.width + 1e-9 && rect.y + rect.height <= b.height + 1e-9
}

test('free corner and edge drags move only the dragged edges', () => {
  const start = { x: 100, y: 100, width: 200, height: 100 }
  assert.deepEqual(C.constrainDrag(start, 'se', 50, 20, free), { x: 100, y: 100, width: 250, height: 120 })
  assert.deepEqual(C.constrainDrag(start, 'nw', -50, -20, free), { x: 50, y: 80, width: 250, height: 120 })
  assert.deepEqual(C.constrainDrag(start, 'e', 30, 99, free), { x: 100, y: 100, width: 230, height: 100 })
  assert.deepEqual(C.constrainDrag(start, 'n', 99, -40, free), { x: 100, y: 60, width: 200, height: 140 })
  assert.deepEqual(C.constrainDrag(start, 'move', 30, -20, free), { x: 130, y: 80, width: 200, height: 100 })
})

test('drags are clamped to the bounds and to the minimum size', () => {
  const start = { x: 100, y: 100, width: 200, height: 100 }
  const grown = C.constrainDrag(start, 'se', 5000, 5000, free)
  assert.deepEqual(grown, { x: 100, y: 100, width: 900, height: 500 })
  const crushed = C.constrainDrag(start, 'se', -1000, -1000, free)
  assert.equal(crushed.width, C.MIN_CROP_SIZE)
  assert.equal(crushed.height, C.MIN_CROP_SIZE)
  assert.equal(crushed.x, 100)
  const moved = C.constrainDrag(start, 'move', -500, 900, free)
  assert.deepEqual(moved, { x: 0, y: 500, width: 200, height: 100 })
})

test('a locked ratio holds for corner drags (the larger side wins) and edge drags', () => {
  const options = { ratio: 16 / 9, fromCenter: false, bounds }
  const start = { x: 100, y: 100, width: 320, height: 180 }
  const corner = C.constrainDrag(start, 'se', 160, 10, options)
  near(corner.width / corner.height, 16 / 9, 1e-9, 'corner ratio')
  near(corner.width, 480, 1e-9)
  assert.equal(corner.x, 100)
  assert.equal(corner.y, 100)
  const tall = C.constrainDrag(start, 'se', 0, 100, options)
  near(tall.height, 280, 1e-9)
  near(tall.width / tall.height, 16 / 9, 1e-9)
  const edge = C.constrainDrag(start, 'e', 160, 0, options)
  near(edge.width / edge.height, 16 / 9, 1e-9, 'edge ratio')
  near(edge.y + edge.height / 2, start.y + start.height / 2, 1e-9, 'edge drag keeps the vertical centre')
  const north = C.constrainDrag(start, 'n', 0, -90, options)
  near(north.width / north.height, 16 / 9, 1e-9)
  near(north.height, 270, 1e-9)
})

test('a locked ratio never leaves the bounds (the box shrinks to fit, ratio kept)', () => {
  const options = { ratio: 1, fromCenter: false, bounds }
  const start = { x: 700, y: 300, width: 100, height: 100 }
  const result = C.constrainDrag(start, 'se', 2000, 2000, options)
  assert.ok(inside(result))
  near(result.width, result.height, 1e-9)
  near(result.width, 300, 1e-9, 'limited by the right edge (1000 - 700)')
  const edgeAtBorder = C.constrainDrag({ x: 0, y: 0, width: 100, height: 100 }, 's', 0, 300, options)
  assert.ok(inside(edgeAtBorder))
  near(edgeAtBorder.width, edgeAtBorder.height, 1e-9)
  near(edgeAtBorder.height, 400, 1e-9, 'the derived side may shift instead of being blocked by the near edge')
})

test('Alt resizes about the centre', () => {
  const start = { x: 400, y: 200, width: 200, height: 100 }
  const result = C.constrainDrag(start, 'e', 50, 0, { ratio: null, fromCenter: true, bounds })
  assert.deepEqual(result, { x: 350, y: 200, width: 300, height: 100 })
  const corner = C.constrainDrag(start, 'nw', -20, -10, { ratio: null, fromCenter: true, bounds })
  assert.deepEqual(corner, { x: 380, y: 190, width: 240, height: 120 })
  const ratioCorner = C.constrainDrag(start, 'se', 100, 0, { ratio: 2, fromCenter: true, bounds })
  near(ratioCorner.x + ratioCorner.width / 2, 500, 1e-9)
  near(ratioCorner.y + ratioCorner.height / 2, 250, 1e-9)
  near(ratioCorner.width / ratioCorner.height, 2, 1e-9)
  const limited = C.constrainDrag(start, 'e', 5000, 0, { ratio: null, fromCenter: true, bounds })
  assert.ok(inside(limited))
  near(limited.width, 1000, 1e-9)
})

test('a new box from two points honours ratio, centre and bounds', () => {
  const box = C.rectFromPoints({ x: 100, y: 100 }, { x: 40, y: 20 }, free)
  assert.deepEqual(box, { x: 40, y: 20, width: 60, height: 80 })
  const square = C.rectFromPoints({ x: 100, y: 100 }, { x: 300, y: 150 }, { ratio: 1, fromCenter: false, bounds })
  assert.deepEqual(square, { x: 100, y: 100, width: 200, height: 200 })
  const centred = C.rectFromPoints({ x: 500, y: 300 }, { x: 520, y: 310 }, { ratio: null, fromCenter: true, bounds })
  assert.deepEqual(centred, { x: 480, y: 290, width: 40, height: 20 })
  const click = C.rectFromPoints({ x: 10, y: 10 }, { x: 10, y: 10 }, free)
  assert.equal(click.width, C.MIN_CROP_SIZE)
})

test('aspect presets give exact ratios, follow the image orientation and can be turned', () => {
  const landscape = { width: 4000, height: 3000 }
  const portrait = { width: 3000, height: 4000 }
  assert.equal(C.presetRatio('free', landscape, false), null)
  near(C.presetRatio('original', landscape, false), 4 / 3)
  near(C.presetRatio('original', landscape, true), 3 / 4)
  near(C.presetRatio('original', portrait, true), 3 / 4)
  near(C.presetRatio('16:9', landscape, false), 16 / 9)
  near(C.presetRatio('16:9', landscape, true), 9 / 16)
  near(C.presetRatio('3:2', portrait, C.defaultPortrait(portrait)), 2 / 3)
  assert.equal(C.presetRatio('1:1', landscape, true), 1)
  assert.equal(C.presetCanSwap('1:1'), false)
  assert.equal(C.presetCanSwap('4:3'), true)
  for (const preset of ['1:1', '4:3', '3:2', '16:9']) {
    for (const turned of [false, true]) {
      const ratio = C.presetRatio(preset, landscape, turned)
      const fitted = C.fitAspect({ x: 1000, y: 1000, width: 10, height: 10 }, ratio, landscape)
      near(fitted.width / fitted.height, ratio, 1e-9, `${preset} turned=${turned}`)
      assert.ok(inside(fitted, landscape))
      assert.ok(Math.abs(fitted.width - landscape.width) < 1e-9 || Math.abs(fitted.height - landscape.height) < 1e-9, 'largest box of the ratio')
    }
  }
})

test('nudges move whole pixels inside the bounds; integer rounding stays inside', () => {
  const rect = { x: 10, y: 10, width: 100, height: 50 }
  assert.deepEqual(C.nudgeRect(rect, 1, 0, bounds), { x: 11, y: 10, width: 100, height: 50 })
  assert.deepEqual(C.nudgeRect(rect, -10, -10, bounds), { x: 0, y: 0, width: 100, height: 50 })
  assert.deepEqual(C.nudgeRect(rect, 5000, 0, bounds), { x: 900, y: 10, width: 100, height: 50 })
  assert.deepEqual(C.toIntegerRect({ x: 9.6, y: -3, width: 2000.4, height: 49.5 }, bounds), { x: 10, y: 0, width: 990, height: 50 })
  assert.deepEqual(C.defaultCropRect({ width: 4096, height: 3072 }), { x: 410, y: 307, width: 3277, height: 2458 })
})

test('straighten angles fold into (-45, 45] and level-by-line evens out the drawn line', () => {
  assert.equal(C.foldAngle(10), 10)
  assert.equal(C.foldAngle(-44.5), -44.5)
  assert.equal(C.foldAngle(88), -2)
  assert.equal(C.foldAngle(-91), -1)
  assert.equal(C.foldAngle(180), 0)
  assert.equal(C.clampStraighten(57), 45)
  near(C.clampStraighten(2.44), 2.4, 1e-9)
  // A horizon falling 5 degrees to the right is levelled by turning 5 degrees anticlockwise.
  const t = (5 * Math.PI) / 180
  near(C.levelAngle({ x: 0, y: 0 }, { x: Math.cos(t) * 100, y: Math.sin(t) * 100 }), -5, 1e-6)
  // A nearly vertical line (88 degrees) is stood upright with +2 degrees.
  const v = (88 * Math.PI) / 180
  near(C.levelAngle({ x: 0, y: 0 }, { x: Math.cos(v) * 100, y: Math.sin(v) * 100 }), 2, 1e-6)
  // Drawn on a preview already turned by 3 degrees: the total angle accounts for it.
  near(C.levelAngle({ x: 0, y: 0 }, { x: Math.cos(t) * 100, y: Math.sin(t) * 100 }, 3), -2, 1e-6)
})

test('the frame matches the worker expand size and maps back to the image', () => {
  const image = { width: 400, height: 300 }
  assert.deepEqual(C.cropFrame(image, 0), image)
  const frame = C.cropFrame(image, 10)
  const t = (10 * Math.PI) / 180
  assert.equal(frame.width, Math.ceil(400 * Math.cos(t) + 300 * Math.sin(t) - 1e-6))
  assert.equal(frame.height, Math.ceil(400 * Math.sin(t) + 300 * Math.cos(t) - 1e-6))
  const centre = C.frameToImage({ x: frame.width / 2, y: frame.height / 2 }, image, 10)
  near(centre.x, 200, 1e-9)
  near(centre.y, 150, 1e-9)
})

test('auto-crop is the largest centred box of the ratio inside the turned image', () => {
  const image = { width: 400, height: 300 }
  for (const degrees of [-30, -10, -1.5, 3, 10, 45]) {
    const box = C.inscribedCrop(image, degrees)
    assert.ok(C.insideRotated(box, image, degrees, 1.5), `inside at ${degrees}`)
    const t = (Math.abs(degrees) * Math.PI) / 180
    const scale = Math.min(396 / (396 * Math.cos(t) + 296 * Math.sin(t)), 296 / (396 * Math.sin(t) + 296 * Math.cos(t)))
    assert.ok(box.width >= Math.floor(396 * scale) - 2 && box.width <= Math.ceil(396 * scale), `width at ${degrees}: ${box.width} vs ${396 * scale}`)
    near(box.width / box.height, 4 / 3, 0.02, `ratio at ${degrees}`)
    // Growing the box by 2 px in each direction leaves the turned image.
    assert.equal(C.insideRotated({ x: box.x - 3, y: box.y - 3, width: box.width + 6, height: box.height + 6 }, image, degrees, 0), false)
  }
  const square = C.inscribedCrop(image, 10, 1)
  near(square.width, square.height, 1)
  assert.deepEqual(C.inscribedCrop(image, 0, null), { x: 0, y: 0, width: 400, height: 300 })
})

test('drags while straightened stop at the turned image edge and slide along it', () => {
  const image = { width: 400, height: 300 }
  const degrees = 10
  const frame = C.cropFrame(image, degrees)
  const start = C.inscribedCrop(image, degrees)
  const options = { ratio: null, fromCenter: false, bounds: frame }
  const keep = (rect) => C.insideRotated(rect, image, degrees)
  const grown = C.constrainDragInside(start, 'se', 400, 400, options, keep)
  assert.ok(keep(grown))
  assert.ok(grown.width >= start.width && grown.height >= start.height)
  const moved = C.constrainDragInside(start, 'move', 0, -500, options, keep)
  assert.ok(keep(moved))
  assert.ok(moved.y < start.y, 'moved up as far as the image allows')
  const ratioOptions = { ratio: 4 / 3, fromCenter: false, bounds: frame }
  const locked = C.constrainDragInside(start, 'se', 300, 0, ratioOptions, keep)
  assert.ok(keep(locked))
  near(locked.width / locked.height, 4 / 3, 1e-6)
  const fresh = C.rectFromPointsInside({ x: frame.width / 2, y: frame.height / 2 }, { x: frame.width, y: frame.height }, options, keep)
  assert.ok(fresh && keep(fresh))
  assert.equal(C.rectFromPointsInside({ x: 0, y: 0 }, { x: 50, y: 50 }, options, keep), null, 'a box cannot start in an empty corner')
})
