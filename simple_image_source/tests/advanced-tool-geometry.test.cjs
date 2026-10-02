'use strict'
// WP5 Advanced tools (src/advanced/toolGeometry.ts, src/advanced/tools/*).
//   Part 1, pure geometry: free-transform handle math (proportional corners by default, Shift toggling,
//   Alt about the reference point, rotation snapping, skew, distort and perspective quads), crop
//   constraints (presets, handles, rotation, straighten, the crop transform), polygon closing and the
//   marquee modifier rules.
//   Part 2, tool behaviour on the real document store with a fake editor context: every tool commits
//   exactly one history step per gesture (none for the eyedropper, hand and zoom), selection tools never
//   mark the document modified, refusals record nothing, and the pixels / selections come out right.
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

const geo = load('advanced/toolGeometry.ts')
const tools = load('advanced/tools/index.ts')
const shared = load('advanced/tools/shared.ts')
const documentModule = load('advanced/document.ts')
const tiles = load('advanced/tiles.ts')
const composite = load('advanced/composite.ts')
const selectionModule = load('advanced/selection.ts')
const mask = load('imaging/mask.ts')
const transform = load('imaging/transform.ts')
const workerOps = load('shared/worker-ops/index.ts')
const workerClient = load('shared/workerClient.ts')

const near = (actual, expected, tolerance = 1e-6, message) => {
  assert.ok(Math.abs(actual - expected) <= tolerance, `${message ?? 'value'}: expected ${expected}, got ${actual}`)
}
const nearPoint = (actual, expected, tolerance = 1e-6, message = 'point') => {
  near(actual.x, expected.x, tolerance, `${message}.x`)
  near(actual.y, expected.y, tolerance, `${message}.y`)
}
const nearQuad = (actual, expected, tolerance = 1e-6) => {
  for (let i = 0; i < 4; i += 1) nearPoint(actual[i], expected[i], tolerance, `corner ${i}`)
}

// =============================================================================================
// Part 1: geometry
// =============================================================================================

const SOURCE = { x: 10, y: 20, width: 100, height: 50 }

test('a fresh frame is the source rectangle with the reference point at its centre', () => {
  const frame = geo.createTransformFrame(SOURCE)
  nearQuad(frame.quad, [{ x: 10, y: 20 }, { x: 110, y: 20 }, { x: 110, y: 70 }, { x: 10, y: 70 }])
  nearPoint(frame.pivot, { x: 60, y: 45 })
  assert.equal(geo.isIdentityFrame(frame), true)
  assert.deepEqual(geo.frameIntegerTranslation(frame), { x: 0, y: 0 })
})

test('corner scaling is proportional by default and follows the pointer projected on the diagonal', () => {
  const frame = geo.createTransformFrame(SOURCE)
  // Pointer off the diagonal: the scale is the projection, so the aspect ratio stays 2:1.
  const quad = geo.scaleFrame(frame, 'se', { x: 210, y: 80 }, { proportional: true, fromPivot: false })
  const width = quad[1].x - quad[0].x
  const height = quad[3].y - quad[0].y
  near(width / height, 2, 1e-9, 'aspect')
  nearPoint(quad[0], { x: 10, y: 20 }, 1e-9, 'fixed opposite corner')
  // Exactly on the diagonal: exact doubling.
  const doubled = geo.scaleFrame(frame, 'se', { x: 210, y: 120 }, { proportional: true, fromPivot: false })
  nearQuad(doubled, [{ x: 10, y: 20 }, { x: 210, y: 20 }, { x: 210, y: 120 }, { x: 10, y: 120 }], 1e-9)
})

test('Shift toggles corner scaling to free (each axis follows the pointer)', () => {
  const frame = geo.createTransformFrame(SOURCE)
  const quad = geo.scaleFrame(frame, 'se', { x: 210, y: 95 }, { proportional: false, fromPivot: false })
  nearQuad(quad, [{ x: 10, y: 20 }, { x: 210, y: 20 }, { x: 210, y: 95 }, { x: 10, y: 95 }], 1e-9)
  // The nw corner scales about the se corner.
  const nw = geo.scaleFrame(frame, 'nw', { x: 60, y: 45 }, { proportional: false, fromPivot: false })
  nearQuad(nw, [{ x: 60, y: 45 }, { x: 110, y: 45 }, { x: 110, y: 70 }, { x: 60, y: 70 }], 1e-9)
})

test('Alt scales about the reference point', () => {
  const frame = geo.createTransformFrame(SOURCE)
  const quad = geo.scaleFrame(frame, 'se', { x: 160, y: 95 }, { proportional: false, fromPivot: true })
  // Pivot (60, 45): se moved from (110, 70) to (160, 95) = double distance, so the frame doubles about it.
  nearQuad(quad, [{ x: -40, y: -5 }, { x: 160, y: -5 }, { x: 160, y: 95 }, { x: -40, y: 95 }], 1e-9)
  nearPoint(geo.quadCenter(quad), { x: 60, y: 45 }, 1e-9, 'centre stays')
})

test('edge handles scale one axis; Shift scales both by the same factor', () => {
  const frame = geo.createTransformFrame(SOURCE)
  const wide = geo.scaleFrame(frame, 'e', { x: 160, y: 999 }, { proportional: false, fromPivot: false })
  nearQuad(wide, [{ x: 10, y: 20 }, { x: 160, y: 20 }, { x: 160, y: 70 }, { x: 10, y: 70 }], 1e-9)
  const both = geo.scaleFrame(frame, 'e', { x: 210, y: 999 }, { proportional: true, fromPivot: false })
  near(both[1].x - both[0].x, 200, 1e-9, 'width')
  near(both[3].y - both[0].y, 100, 1e-9, 'height')
  // The opposite edge's centre line stays put.
  near((both[0].y + both[3].y) / 2, 45, 1e-9, 'vertical centre')
  const flipped = geo.scaleFrame(frame, 'e', { x: -90, y: 0 }, { proportional: false, fromPivot: false })
  assert.ok(flipped[1].x < flipped[0].x, 'dragging past the opposite edge flips the frame')
  assert.equal(geo.isValidQuad(flipped), true, 'a mirrored frame is still valid')
})

test('rotated frames scale along their own axes', () => {
  const frame = geo.createTransformFrame({ x: 0, y: 0, width: 100, height: 50 })
  const rotated = { ...frame, quad: geo.rotateFrame(frame, { x: 100, y: 25 }, { x: 50, y: 75 }) }
  near(geo.quadAngle(rotated.quad), 90, 1e-9, 'quarter turn')
  // The 'e' handle now points down; dragging it 50 px further down doubles the frame's width.
  const handle = geo.transformHandlePoint(rotated, 'e')
  const scaled = geo.scaleFrame(rotated, 'e', { x: handle.x, y: handle.y + 50 }, { proportional: false, fromPivot: false })
  const info = geo.frameInfo({ ...rotated, quad: scaled })
  near(info.width, 150, 1e-6, 'own width')
  near(info.height, 50, 1e-6, 'own height unchanged')
})

test('rotation turns about the reference point and Shift snaps the total angle to 15 degrees', () => {
  const frame = geo.createTransformFrame(SOURCE)
  const pivot = frame.pivot
  const quad = geo.rotateFrame(frame, { x: pivot.x + 100, y: pivot.y }, { x: pivot.x + 100, y: pivot.y + 40 })
  for (let i = 0; i < 4; i += 1) near(geo.distance(quad[i], pivot), geo.distance(frame.quad[i], pivot), 1e-9, 'radius kept')
  const free = geo.quadAngle(quad)
  near(free, Math.atan2(40, 100) * 180 / Math.PI, 1e-9, 'free angle')
  const snapped = geo.rotateFrame(frame, { x: pivot.x + 100, y: pivot.y }, { x: pivot.x + 100, y: pivot.y + 40 }, 15)
  near(geo.quadAngle(snapped), 15, 1e-9, 'snapped to 15')
  // Snapping applies to the total angle of an already turned frame.
  const turned = { ...frame, quad: geo.rotateFrame(frame, { x: 1, y: 0 }, { x: 1, y: 0.1 }) }
  const again = geo.rotateFrame(turned, { x: turned.pivot.x + 10, y: turned.pivot.y }, { x: turned.pivot.x + 10, y: turned.pivot.y + 8 }, 15)
  near(Math.round(geo.quadAngle(again) * 1e6) / 1e6 % 15, 0, 1e-6, 'multiple of 15')
})

test('Ctrl+edge skews: the edge slides along itself, Alt mirrors the opposite edge', () => {
  const frame = geo.createTransformFrame(SOURCE)
  const quad = geo.skewFrame(frame, 'n', { x: 60, y: 20 }, { x: 80, y: 5 }, false)
  nearPoint(quad[0], { x: 30, y: 20 }, 1e-9, 'top-left slid by the horizontal part only')
  nearPoint(quad[1], { x: 130, y: 20 }, 1e-9)
  nearPoint(quad[2], { x: 110, y: 70 }, 1e-9, 'bottom unchanged')
  assert.equal(geo.isAffineQuad(quad), true, 'a skew stays affine')
  const symmetric = geo.skewFrame(frame, 'n', { x: 60, y: 20 }, { x: 80, y: 20 }, true)
  nearPoint(symmetric[3], { x: -10, y: 70 }, 1e-9, 'bottom slides the other way')
  nearPoint(symmetric[2], { x: 90, y: 70 }, 1e-9)
})

test('Ctrl+corner distorts one corner; the result is a true homography of the source', () => {
  const frame = geo.createTransformFrame(SOURCE)
  const quad = geo.distortFrame(frame, 'se', { x: 110, y: 70 }, { x: 140, y: 90 })
  nearPoint(quad[2], { x: 140, y: 90 }, 1e-9)
  nearPoint(quad[0], frame.quad[0], 1e-9)
  assert.equal(geo.isAffineQuad(quad), false)
  assert.equal(geo.frameAffine({ source: SOURCE, quad }), null)
  const h = geo.frameHomography({ source: SOURCE, quad })
  const corners = geo.rectToQuad(SOURCE)
  for (let i = 0; i < 4; i += 1) nearPoint(transform.applyHomography(h, corners[i]), quad[i], 1e-6, `mapped corner ${i}`)
})

test('Ctrl+Alt+Shift+corner gives a symmetric perspective trapezoid', () => {
  const frame = geo.createTransformFrame(SOURCE)
  const horizontal = geo.perspectiveFrame(frame, 'nw', { x: 10, y: 20 }, { x: 30, y: 22 })
  nearPoint(horizontal[0], { x: 30, y: 20 }, 1e-9, 'corner slides along the top edge')
  nearPoint(horizontal[1], { x: 90, y: 20 }, 1e-9, 'its neighbour slides the other way')
  nearPoint(horizontal[2], { x: 110, y: 70 }, 1e-9)
  nearPoint(horizontal[3], { x: 10, y: 70 }, 1e-9)
  assert.equal(geo.isValidQuad(horizontal), true)
  const vertical = geo.perspectiveFrame(frame, 'ne', { x: 110, y: 20 }, { x: 111, y: 30 })
  nearPoint(vertical[1], { x: 110, y: 30 }, 1e-9)
  nearPoint(vertical[2], { x: 110, y: 60 }, 1e-9)
})

test('quad validity, affine extraction, integer translations and numeric fields', () => {
  assert.equal(geo.isValidQuad([{ x: 0, y: 0 }, { x: 10, y: 10 }, { x: 10, y: 0 }, { x: 0, y: 10 }]), false, 'a bow-tie is invalid')
  const frame = geo.createTransformFrame(SOURCE)
  const moved = { ...frame, quad: geo.translateQuad(frame.quad, 7, -3) }
  assert.deepEqual(geo.frameIntegerTranslation(moved), { x: 7, y: -3 })
  assert.equal(geo.frameIntegerTranslation({ ...frame, quad: geo.translateQuad(frame.quad, 0.5, 0) }), null)
  const affine = geo.frameAffine(moved)
  assert.deepEqual(affine.map((v) => Math.round(v * 1e9) / 1e9), [1, 0, 0, 1, 7, -3])
  const rebuilt = geo.frameFromInfo(frame, { width: 200, height: 100, angle: 30 })
  const info = geo.frameInfo(rebuilt)
  near(info.width, 200, 1e-9)
  near(info.height, 100, 1e-9)
  near(info.angle, 30, 1e-9)
  nearPoint({ x: info.x, y: info.y }, frame.pivot, 1e-9, 'reference point kept')
  nearPoint(geo.quadCenter(rebuilt.quad), frame.pivot, 1e-9, 'centred reference point stays the centre')
})

test('hit testing finds the reference point, handles, inside and outside', () => {
  const frame = geo.createTransformFrame(SOURCE)
  assert.deepEqual(geo.hitTestTransform(frame, { x: 60, y: 45 }, 4), { kind: 'pivot' })
  assert.deepEqual(geo.hitTestTransform(frame, { x: 111, y: 71 }, 4), { kind: 'handle', handle: 'se' })
  assert.deepEqual(geo.hitTestTransform(frame, { x: 60, y: 21 }, 4), { kind: 'handle', handle: 'n' })
  assert.deepEqual(geo.hitTestTransform(frame, { x: 30, y: 40 }, 4), { kind: 'inside' })
  assert.deepEqual(geo.hitTestTransform(frame, { x: 200, y: 200 }, 4), { kind: 'outside' })
})

test('warpInverse maps destination pixels back into the source buffer', () => {
  const frame = geo.createTransformFrame(SOURCE)
  const quad = geo.scaleFrame(frame, 'se', { x: 210, y: 120 }, { proportional: true, fromPivot: false })
  const forward = geo.frameHomography({ source: SOURCE, quad })
  const inverse = geo.warpInverse(forward, SOURCE.x, SOURCE.y)
  nearPoint(transform.applyHomography(inverse, { x: 210, y: 120 }), { x: 100, y: 50 }, 1e-9, 'far corner')
  nearPoint(transform.applyHomography(inverse, { x: 10, y: 20 }), { x: 0, y: 0 }, 1e-9, 'origin')
  // A quarter turn keeps exact integer entries (pixel-exact warps).
  const turned = geo.rotateFrame(geo.createTransformFrame({ x: 0, y: 0, width: 4, height: 2 }), { x: 4, y: 1 }, { x: 2, y: 3 })
  const exact = geo.warpInverse(geo.frameHomography({ source: { x: 0, y: 0, width: 4, height: 2 }, quad: turned }), 0, 0)
  for (const value of exact.slice(0, 6)) assert.equal(value, Math.round(value * 1e6) / 1e6)
})

test('crop presets give exact ratios, portrait swaps them, Original follows the document', () => {
  const doc = { width: 3000, height: 2000 }
  assert.equal(geo.cropAspectRatio('free', false, doc), null)
  near(geo.cropAspectRatio('original', false, doc), 1.5)
  near(geo.cropAspectRatio('original', true, doc), 2 / 3)
  near(geo.cropAspectRatio('16:9', false, doc), 16 / 9)
  near(geo.cropAspectRatio('16:9', true, doc), 9 / 16)
  near(geo.cropAspectRatio('4:3', false, doc), 4 / 3)
  near(geo.cropAspectRatio('1:1', true, doc), 1)
  for (const preset of ['1:1', '4:3', '3:2', '16:9', '5:4', '7:5']) {
    const box = geo.dragCrop({ x: 10, y: 10 }, { x: 410, y: 90 }, { ratio: geo.cropAspectRatio(preset, false, doc) })
    near(box.width / box.height, geo.cropAspectRatio(preset, false, doc), 1e-9, preset)
  }
})

test('crop handles keep the opposite side, presets and Shift keep ratios, Alt resizes about the centre', () => {
  const box = geo.cropBoxFromRect({ x: 0, y: 0, width: 400, height: 300 })
  const free = geo.resizeCrop(box, 'se', { x: 500, y: 350 }, { ratio: null })
  assert.deepEqual(geo.cropBoxRect(free), { x: 0, y: 0, width: 500, height: 350 })
  const ratio = geo.resizeCrop(box, 'se', { x: 500, y: 320 }, { ratio: 16 / 9 })
  near(ratio.width / ratio.height, 16 / 9, 1e-9, 'preset ratio')
  near(ratio.cx - ratio.width / 2, 0, 1e-9, 'left edge fixed')
  near(ratio.cy - ratio.height / 2, 0, 1e-9, 'top edge fixed')
  const kept = geo.resizeCrop(box, 'se', { x: 800, y: 320 }, { ratio: null, keepRatio: true })
  near(kept.width / kept.height, 4 / 3, 1e-9, 'Shift keeps the current ratio')
  const edge = geo.resizeCrop(box, 'e', { x: 600, y: 999 }, { ratio: 1 })
  near(edge.width, 600, 1e-9)
  near(edge.height, 600, 1e-9)
  near(edge.cy, 150, 1e-9, 'an edge with a ratio grows about the centre line')
  const centred = geo.resizeCrop(box, 'e', { x: 300, y: 0 }, { ratio: null, fromCenter: true })
  near(centred.cx, 200, 1e-9)
  near(centred.width, 200, 1e-9)
  const tiny = geo.resizeCrop(box, 'e', { x: 0, y: 0 }, { ratio: null })
  assert.ok(tiny.width >= 1, 'never narrower than one pixel')
  const flipped = geo.resizeCrop(box, 'e', { x: -100, y: 0 }, { ratio: null })
  assert.deepEqual(geo.cropBoxRect(flipped), { x: -100, y: 0, width: 100, height: 300 })
})

test('a turned crop box resizes in its own frame and its transform maps it onto the new canvas', () => {
  const box = { ...geo.cropBoxFromRect({ x: 100, y: 100, width: 200, height: 100 }), angle: 30 }
  const west = geo.cropHandlePoint(box, 'w')
  const east = geo.cropHandlePoint(box, 'e')
  const along = { x: (east.x - west.x) / 200, y: (east.y - west.y) / 200 }
  const pointer = { x: east.x + along.x * 50, y: east.y + along.y * 50 }
  const resized = geo.resizeCrop(box, 'e', pointer, { ratio: null })
  near(resized.width, 250, 1e-9)
  near(resized.height, 100, 1e-9)
  nearPoint(geo.cropHandlePoint(resized, 'w'), west, 1e-9, 'west edge fixed in the document')
  const { matrix, width, height } = geo.cropTransform(box)
  assert.equal(width, 200)
  assert.equal(height, 100)
  const map = (p) => ({ x: matrix[0] * p.x + matrix[2] * p.y + matrix[4], y: matrix[1] * p.x + matrix[3] * p.y + matrix[5] })
  const corners = geo.cropBoxCorners(box)
  nearPoint(map(corners[0]), { x: 0, y: 0 }, 1e-9, 'top-left')
  nearPoint(map(corners[1]), { x: 200, y: 0 }, 1e-9, 'top-right')
  nearPoint(map(corners[2]), { x: 200, y: 100 }, 1e-9, 'bottom-right')
  nearPoint(map(corners[3]), { x: 0, y: 100 }, 1e-9, 'bottom-left')
})

test('crop rotation snaps, straighten levels the drawn line, presets fit inside the box', () => {
  const box = geo.cropBoxFromRect({ x: 0, y: 0, width: 400, height: 300 })
  const turned = geo.rotateCrop(box, { x: 400, y: 150 }, { x: 400, y: 190 }, 15)
  near(turned.angle % 15, 0, 1e-9, 'snapped')
  near(geo.rotateCrop(box, { x: 400, y: 150 }, { x: 400, y: 190 }).angle, Math.atan2(40, 200) * 180 / Math.PI, 1e-9)
  near(geo.straightenAngle({ x: 0, y: 0 }, { x: 100, y: Math.tan(5 * Math.PI / 180) * 100 }), 5, 1e-9, 'near-horizontal')
  near(geo.straightenAngle({ x: 0, y: 0 }, { x: Math.tan(-4 * Math.PI / 180) * 100, y: 100 }), 4, 1e-9, 'near-vertical')
  const fitted = geo.fitCropRatio(box, 1)
  assert.equal(fitted.width, 300)
  assert.equal(fitted.height, 300)
  assert.deepEqual(geo.cropCommitRect({ cx: 50.4, cy: 40.6, width: 20.2, height: 10.4, angle: 0 }), { x: 40, y: 35, width: 21, height: 11 })
  assert.equal(geo.hitTestCrop(box, { x: 401, y: 299 }, 5), 'se')
  assert.equal(geo.hitTestCrop(box, { x: 200, y: 150 }, 5), 'inside')
  assert.equal(geo.hitTestCrop(box, { x: 600, y: 150 }, 5), 'outside')
})

test('polygon closing, degenerate polygons, path thinning and simplification', () => {
  const points = [{ x: 0, y: 0 }, { x: 10, y: 0 }, { x: 10, y: 10 }]
  assert.equal(geo.closesPolygon(points, { x: 1, y: 1 }, 2), true)
  assert.equal(geo.closesPolygon(points, { x: 5, y: 5 }, 2), false)
  assert.equal(geo.closesPolygon(points.slice(0, 2), { x: 0, y: 0 }, 2), false, 'two vertices cannot close')
  assert.equal(geo.isDegeneratePolygon([{ x: 0, y: 0 }, { x: 5, y: 5 }, { x: 10, y: 10 }]), true, 'collinear')
  assert.equal(geo.isDegeneratePolygon(points), false)
  near(geo.polygonArea(points), 50, 1e-9)
  const path = []
  assert.equal(geo.appendPathPoint(path, { x: 0, y: 0 }, 1), true)
  assert.equal(geo.appendPathPoint(path, { x: 0.5, y: 0 }, 1), false, 'too close')
  assert.equal(geo.appendPathPoint(path, { x: 1, y: 0 }, 1), true)
  const line = Array.from({ length: 11 }, (_, i) => ({ x: i, y: i % 2 ? 0.05 : 0 }))
  const simple = geo.simplifyPath(line, 0.2)
  assert.deepEqual(simple, [{ x: 0, y: 0 }, { x: 10, y: 0 }])
  assert.equal(geo.pointInPolygon(points, { x: 8, y: 2 }), true)
  assert.equal(geo.pointInPolygon(points, { x: 2, y: 8 }), false)
})

test('marquee modifiers: Shift/Alt pick add/subtract/intersect only with a selection, then constrain after release', () => {
  assert.deepEqual(geo.selectionOpFromModifiers(true, false, false, 'replace'), { op: 'replace', shiftLatched: false, altLatched: false })
  assert.deepEqual(geo.selectionOpFromModifiers(false, false, false, 'subtract'), { op: 'replace', shiftLatched: false, altLatched: false }, 'no selection: new')
  assert.deepEqual(geo.selectionOpFromModifiers(true, false, true, 'replace'), { op: 'add', shiftLatched: true, altLatched: false })
  assert.deepEqual(geo.selectionOpFromModifiers(false, true, true, 'replace'), { op: 'subtract', shiftLatched: false, altLatched: true })
  assert.deepEqual(geo.selectionOpFromModifiers(true, true, true, 'replace'), { op: 'intersect', shiftLatched: true, altLatched: true })
  assert.deepEqual(geo.selectionOpFromModifiers(false, false, true, 'add'), { op: 'add', shiftLatched: false, altLatched: false })
  let state = geo.modifierState(true, true)
  assert.deepEqual(state, { latched: true, active: false }, 'held since pointer-down: no constraint')
  state = geo.modifierState(state.latched, false)
  assert.deepEqual(state, { latched: false, active: false })
  state = geo.modifierState(state.latched, true)
  assert.deepEqual(state, { latched: false, active: true }, 'pressed again: constrains')
  const square = geo.marqueeRect({ x: 10, y: 10 }, { x: 40, y: 20 }, { square: true, fromCenter: false })
  assert.deepEqual(square, { x: 10, y: 10, width: 30, height: 30 })
  const upLeft = geo.marqueeRect({ x: 40, y: 40 }, { x: 30, y: 10 }, { square: true, fromCenter: false })
  assert.deepEqual(upLeft, { x: 10, y: 10, width: 30, height: 30 }, 'square towards the top-left')
  const centred = geo.marqueeRect({ x: 50, y: 50 }, { x: 60, y: 55 }, { square: false, fromCenter: true })
  assert.deepEqual(centred, { x: 40, y: 45, width: 20, height: 10 })
  const ratio = geo.marqueeRect({ x: 0, y: 0 }, { x: 40, y: 5 }, { square: false, fromCenter: false }, { style: 'fixed-ratio', ratio: { width: 2, height: 1 }, fixedSize: { width: 1, height: 1 } })
  assert.deepEqual(ratio, { x: 0, y: 0, width: 40, height: 20 })
  const fixed = geo.marqueeRect({ x: 0, y: 0 }, { x: 7, y: 9 }, { square: false, fromCenter: false }, { style: 'fixed-size', ratio: { width: 1, height: 1 }, fixedSize: { width: 64, height: 32 } })
  assert.deepEqual(fixed, { x: 7, y: 9, width: 64, height: 32 })
  assert.deepEqual(geo.snapRectToPixels({ x: 1.4, y: 2.6, width: 3.3, height: 2 }), { x: 1, y: 3, width: 4, height: 2 })
})

test('angle snapping, line drags, nudges, double clicks, scrubby zoom and affine composition', () => {
  nearPoint(geo.snapToAngle({ x: 0, y: 0 }, { x: 10, y: 1 }, 45), { x: Math.hypot(10, 1), y: 0 }, 1e-9, 'horizontal')
  const diagonal = geo.snapToAngle({ x: 0, y: 0 }, { x: 10, y: 9 }, 45)
  near(diagonal.x, diagonal.y, 1e-9, '45 degrees')
  const line = geo.lineFromDrag({ x: 10, y: 10 }, { x: 20, y: 12 }, { snap: true, fromCenter: true })
  near(line.start.y, 10, 1e-9)
  near(line.end.y, 10, 1e-9)
  near(line.start.x + line.end.x, 20, 1e-9, 'mirrored about the anchor')
  assert.deepEqual(geo.nudgeDelta('ArrowLeft', false), { x: -1, y: 0 })
  assert.deepEqual(geo.nudgeDelta('ArrowDown', true), { x: 0, y: 10 })
  assert.equal(geo.nudgeDelta('a', false), null)
  assert.equal(geo.isDoubleClick({ time: 0, x: 0, y: 0 }, { time: 300, x: 2, y: 2 }), true)
  assert.equal(geo.isDoubleClick({ time: 0, x: 0, y: 0 }, { time: 900, x: 0, y: 0 }), false)
  near(geo.scrubbyZoom(1, 150, 1), 2, 1e-9)
  near(geo.scrubbyZoom(1, -150, 1), 0.5, 1e-9)
  near(geo.scrubbyZoom(10, 10000, 1), 32, 1e-9, 'clamped to 3200%')
  assert.deepEqual(geo.composeAffine([2, 0, 0, 2, 5, 5], [1, 0, 0, 1, 3, 4]), [2, 0, 0, 2, 11, 13])
  const t = geo.anchoredTranslation([1, 0, 0, 1], { x: 100, y: 50 }, { x: 40, y: 20 })
  assert.deepEqual(t, { x: 60, y: 30 })
})

// =============================================================================================
// Part 2: tools on the real document store
// =============================================================================================

const W = 64
const H = 48
const RED = { r: 255, g: 0, b: 0 }
const WHITE = { r: 255, g: 255, b: 255 }

const OPTIONS = Object.freeze({
  brush: { size: 6, hardness: 1, opacity: 1, flow: 1, spacing: 0.25, smoothing: 0, pressureSize: false, pressureOpacity: false, blendMode: 'normal' },
  eraser: { size: 6, hardness: 1, opacity: 1, flow: 1, spacing: 0.25, smoothing: 0, pressureSize: false, pressureOpacity: false, blendMode: 'normal' },
  clone: { size: 6, hardness: 1, opacity: 1, flow: 1, spacing: 0.25, smoothing: 0, pressureSize: false, pressureOpacity: false, blendMode: 'normal', aligned: true, sample: 'current' },
  heal: { size: 6, hardness: 1, sample: 'current' },
  marquee: { op: 'replace', feather: 0, antiAlias: true, style: 'normal', ratio: { width: 1, height: 1 }, fixedSize: { width: 10, height: 10 } },
  lasso: { op: 'replace', feather: 0, antiAlias: false },
  wand: { op: 'replace', tolerance: 16, contiguous: true, antiAlias: false, sample: 'current' },
  bucket: { tolerance: 16, contiguous: true, antiAlias: false, sample: 'current', opacity: 1, blendMode: 'normal' },
  gradient: { kind: 'linear', preset: 'foreground-background', reverse: false, dither: false, opacity: 1, blendMode: 'normal' },
  crop: { aspect: 'free', portrait: false, deleteCroppedPixels: true, overlay: 'thirds' },
  text: { fontFamily: 'Segoe UI', fontSize: 24, fontWeight: 400, italic: false, underline: false, color: { r: 0, g: 0, b: 0 }, align: 'left', lineHeight: 1.2, letterSpacing: 0 },
  shape: { kind: 'rectangle', fill: { r: 0, g: 0, b: 255, a: 255 }, stroke: null, strokeWidth: 2, cornerRadius: 0 },
  eyedropper: { size: 1, sample: 'all' },
  move: { autoSelect: false, showTransformControls: false },
})

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

/** Background: a horizontal grey ramp (distinct, opaque). */
function ramp() {
  const data = new Uint8ClampedArray(W * H * 4)
  for (let y = 0; y < H; y += 1) {
    for (let x = 0; x < W; x += 1) {
      const p = (y * W + x) * 4
      data[p] = 40 + x * 2
      data[p + 1] = 60 + x * 2
      data[p + 2] = 80 + y
      data[p + 3] = 255
    }
  }
  return { width: W, height: H, data }
}

function makeHarness(setup = {}) {
  let counter = 0
  let time = 1000
  const host = {
    revision: 0,
    nextRevision: () => (counter += 1),
    setRevision: (revision) => { host.revision = revision },
    currentRevision: () => host.revision,
  }
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(ramp()) })
  const layer = documentModule.createRasterLayer({ name: 'Layer 1', surface: tiles.surfaceFromBuffer(solid(16, 16, RED)), offsetX: 8, offsetY: 8 })
  const layers = setup.layers ?? [background, layer]
  const store = documentModule.createDocumentStore({
    width: W, height: H, ppi: 72, layers, host, baseLabel: 'Open', now: () => time, activeLayerId: setup.active ?? layer.id,
  })
  let editorState = {
    tool: 'move',
    springFrom: null,
    foreground: RED,
    background: WHITE,
    options: JSON.parse(JSON.stringify(OPTIONS)),
    swatches: [],
    recentColors: [],
    view: { zoom: 1, offsetX: 0, offsetY: 0 },
    panels: { layers: true, properties: true, history: true, color: true },
    lastFilter: null,
  }
  const editorListeners = new Set()
  const editor = {
    getState: () => editorState,
    subscribe(listener) { editorListeners.add(listener); return () => editorListeners.delete(listener) },
    update(patch) { editorState = { ...editorState, ...patch }; for (const l of [...editorListeners]) l() },
    updateOptions(tool, patch) {
      editorState = { ...editorState, options: { ...editorState.options, [tool]: { ...editorState.options[tool], ...patch } } }
      for (const l of [...editorListeners]) l()
    },
  }
  const calls = { pan: [], zoom: [], fit: 0, actual: 0, previews: [], notes: [], cursors: [], overlays: [] }
  let view = { zoom: 1, offsetX: 0, offsetY: 0 }
  const viewController = {
    getView: () => view,
    getViewportSize: () => ({ width: 640, height: 480, dpr: 1 }),
    docToScreen: (p) => ({ x: p.x * view.zoom + view.offsetX, y: p.y * view.zoom + view.offsetY }),
    screenToDoc: (p) => ({ x: (p.x - view.offsetX) / view.zoom, y: (p.y - view.offsetY) / view.zoom }),
    zoomAt(zoom, anchor) { calls.zoom.push({ zoom, anchor }); view = { ...view, zoom } },
    panBy(dx, dy) { calls.pan.push({ dx, dy }); view = { ...view, offsetX: view.offsetX + dx, offsetY: view.offsetY + dy } },
    fit() { calls.fit += 1 },
    actualPixels() { calls.actual += 1 },
    requestOverlay() {},
  }
  const compositor = {
    level: 0,
    attach() {},
    setView() {},
    invalidate() {},
    setPreview(preview) { calls.previews.push(preview) },
    settle: async () => {},
    flatten: async () => composite.flattenDocument(store.getState()),
    renderToCanvas: async () => { throw new Error('no canvas in Node') },
    sample: (x, y, size, source) => composite.sampleDocument(store.getState(), x, y, size, source),
    dispose() {},
  }
  const imaging = workerClient.createImagingClient({ inline: workerOps.HANDLERS })
  const ctx = {
    store,
    editor,
    view: viewController,
    compositor,
    imaging,
    host: {
      ...host,
      notify(message, tone) { calls.notes.push({ message, tone }) },
      isSuspended: () => false,
      requestExit() {},
      save: async () => true,
      openExportMenu() {},
      print() {},
      copyPng: async () => {},
      readClipboardImage: async () => null,
    },
    setCursor(cursor) { calls.cursors.push(cursor) },
    setHint() {},
    setOverlayElement(element) { calls.overlays.push(element) },
  }
  return {
    store, ctx, editor, calls, host, background, layer,
    tick(ms) { time += ms },
    steps: () => store.history.getState().entries.length - 1,
    labels: () => store.history.getState().entries.slice(1).map((entry) => entry.label),
    pixel(layerOrId, x, y) {
      const state = store.getState()
      const found = typeof layerOrId === 'string' ? state.layers.find((item) => item.id === layerOrId) : layerOrId
      const live = state.layers.find((item) => item.id === found.id)
      return Array.from(shared.readLayerRect(live, { x, y, width: 1, height: 1 }).data)
    },
    composite(x, y) {
      return Array.from(composite.compositeRect(store.getState(), { x, y, width: 1, height: 1 }).data)
    },
    find(id) { return store.getState().layers.find((item) => item.id === id) },
  }
}

let clock = 0
function pe(x, y, extra = {}) {
  clock += 16
  return {
    doc: { x, y },
    screen: { x, y },
    pressure: 1,
    button: 0,
    buttons: 1,
    shift: false,
    alt: false,
    ctrl: false,
    pointerType: 'mouse',
    time: clock,
    coalesced: [],
    ...extra,
  }
}

function key(name, extra = {}) {
  return { key: name, code: name, shiftKey: false, ctrlKey: false, altKey: false, metaKey: false, isComposing: false, repeat: false, ...extra }
}

/** A complete drag from a to b through intermediate points. */
function drag(tool, from, to, extra = {}, steps = 6) {
  tool.pointerDown(pe(from.x, from.y, extra))
  for (let i = 1; i <= steps; i += 1) {
    const t = i / steps
    tool.pointerMove(pe(from.x + (to.x - from.x) * t, from.y + (to.y - from.y) * t, extra))
  }
  tool.pointerUp(pe(to.x, to.y, { ...extra, buttons: 0 }))
}

function activate(h, id) {
  const tool = tools.TOOL_FACTORIES[id]()
  tool.activate(h.ctx)
  return tool
}

async function idle(tool) {
  if (tool.whenIdle) await tool.whenIdle()
  await new Promise((resolve) => setTimeout(resolve, 0))
  if (tool.whenIdle) await tool.whenIdle()
}

test('registry: one factory per tool id, matching ids, groups cover every tool once', () => {
  const ids = Object.keys(tools.TOOL_META)
  assert.equal(ids.length, 18)
  for (const id of ids) {
    const controller = tools.TOOL_FACTORIES[id]()
    assert.equal(controller.id, id)
    for (const method of ['activate', 'deactivate', 'pointerDown', 'pointerMove', 'pointerUp', 'pointerCancel', 'keyDown', 'keyUp', 'drawOverlay', 'hasSession', 'commitSession', 'cancelSession']) {
      assert.equal(typeof controller[method], 'function', `${id}.${method}`)
    }
  }
  assert.deepEqual(tools.TOOL_GROUPS.flat().sort(), ids.slice().sort())
})

test('brush: one stroke is one step, paints the foreground on the active layer and respects locks', () => {
  const h = makeHarness()
  const brush = activate(h, 'brush')
  drag(brush, { x: 2, y: 30 }, { x: 40, y: 30 })
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Brush Tool'])
  assert.deepEqual(h.pixel(h.layer, 20, 30), [255, 0, 0, 255], 'painted (the layer grows past its old bounds)')
  assert.notEqual(h.host.revision, 0, 'marks the document modified')
  h.store.transact('Lock', 'layer', (tx) => tx.updateLayer(h.layer.id, { locks: { pixels: true, position: false, transparency: false } }))
  const before = h.steps()
  drag(brush, { x: 2, y: 40 }, { x: 30, y: 40 })
  assert.equal(h.steps(), before, 'a locked layer records nothing')
  assert.ok(h.calls.notes.some((note) => /locked/i.test(note.message)))
  brush.deactivate()
})

test('brush: Shift+click draws a straight line from the previous stroke, Alt+click picks the colour', () => {
  const h = makeHarness()
  const brush = activate(h, 'brush')
  drag(brush, { x: 2, y: 2 }, { x: 3, y: 2 }, {}, 1)
  brush.pointerDown(pe(40, 2, { shift: true }))
  brush.pointerUp(pe(40, 2, { shift: true, buttons: 0 }))
  assert.equal(h.steps(), 2)
  assert.deepEqual(h.pixel(h.layer, 20, 2), [255, 0, 0, 255], 'the line between both points is painted')
  brush.pointerDown(pe(60, 40, { alt: true }))
  brush.pointerUp(pe(60, 40, { alt: true, buttons: 0 }))
  assert.equal(h.steps(), 2, 'picking a colour records nothing')
  assert.notDeepEqual(h.editor.getState().foreground, RED)
})

test('eraser: lowers alpha on a layer and paints the background colour on the Background', () => {
  const h = makeHarness()
  const eraser = activate(h, 'eraser')
  drag(eraser, { x: 9, y: 16 }, { x: 22, y: 16 })
  assert.equal(h.steps(), 1)
  assert.equal(h.pixel(h.layer, 16, 16)[3], 0, 'erased to transparency')
  h.store.transact('Select', 'layer', (tx) => tx.setActiveLayer(h.background.id))
  drag(eraser, { x: 40, y: 40 }, { x: 50, y: 40 })
  assert.equal(h.steps(), 2)
  assert.deepEqual(h.pixel(h.background, 45, 40), [255, 255, 255, 255], 'Background erases to the background colour')
  assert.deepEqual(h.labels(), ['Eraser', 'Eraser'])
})

test('clone stamp: Alt+click sets the source (no step), a stroke copies source pixels (one step)', () => {
  const h = makeHarness()
  h.store.transact('Select', 'layer', (tx) => tx.setActiveLayer(h.background.id))
  const clone = activate(h, 'clone-stamp')
  clone.pointerDown(pe(10, 20, { alt: true }))
  clone.pointerUp(pe(10, 20, { alt: true, buttons: 0 }))
  assert.equal(h.steps(), 0)
  drag(clone, { x: 40, y: 20 }, { x: 44, y: 20 })
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Clone Stamp'])
  const sourcePixel = Array.from(ramp().data.slice((20 * W + 12) * 4, (20 * W + 12) * 4 + 4))
  assert.deepEqual(h.pixel(h.background, 42, 20), sourcePixel, 'target shows the source 30 px to the left')
})

test('spot healing: one step after the worker job, the spot is replaced by its surroundings', async () => {
  const flat = solid(W, H, { r: 120, g: 120, b: 120 })
  for (let y = 20; y < 24; y += 1) for (let x = 30; x < 34; x += 1) flat.data.set([0, 0, 0, 255], (y * W + x) * 4)
  const background = documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(flat) })
  const h = makeHarness({ layers: [background], active: background.id })
  const heal = activate(h, 'spot-healing')
  drag(heal, { x: 30, y: 22 }, { x: 34, y: 22 })
  await idle(heal)
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Spot Healing Brush'])
  const healed = h.pixel(background, 32, 22)
  assert.ok(Math.abs(healed[0] - 120) <= 6, `healed to the flat grey, got ${healed}`)
})

test('gradient: one step, foreground at the start and background at the end, inside the selection', async () => {
  const h = makeHarness()
  h.store.transact('Select', 'layer', (tx) => tx.setActiveLayer(h.background.id))
  const gradient = activate(h, 'gradient')
  drag(gradient, { x: 0, y: 10 }, { x: 64, y: 10 })
  await idle(gradient)
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Gradient'])
  const left = h.pixel(h.background, 0, 30)
  const right = h.pixel(h.background, 63, 30)
  assert.ok(left[0] > 250 && left[1] < 10, `left is red, got ${left}`)
  assert.ok(right[0] > 245 && right[1] > 245 && right[2] > 245, `right is white, got ${right}`)
  // With a selection only the selected area changes.
  const shape = mask.createMaskBuffer(W, H)
  mask.rasterizeRect(shape, { x: 0, y: 0, width: 10, height: 10 }, false)
  shared.commitSelectionShape(h.ctx, 'Rectangular Marquee', 'marquee-rect', shape, 'replace')
  const outside = h.pixel(h.background, 30, 30)
  drag(gradient, { x: 64, y: 0 }, { x: 0, y: 0 })
  await idle(gradient)
  assert.deepEqual(h.pixel(h.background, 30, 30), outside, 'outside the selection unchanged')
})

test('paint bucket: one step, fills the contiguous region with the foreground', async () => {
  const h = makeHarness()
  h.editor.update({ foreground: { r: 0, g: 200, b: 0 } })
  const bucket = activate(h, 'paint-bucket')
  bucket.pointerDown(pe(12, 12))
  bucket.pointerUp(pe(12, 12, { buttons: 0 }))
  await idle(bucket)
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Paint Bucket'])
  assert.deepEqual(h.pixel(h.layer, 20, 20), [0, 200, 0, 255])
})

test('rectangular and elliptical marquee: one step each, exact pixels, no content revision', () => {
  const h = makeHarness()
  const rect = activate(h, 'marquee-rect')
  drag(rect, { x: 10.2, y: 5.4 }, { x: 30.4, y: 20.6 })
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.store.getState().selection.bounds, { x: 10, y: 5, width: 20, height: 16 })
  assert.equal(h.host.revision, 0, 'selections never mark the document modified')
  const ellipse = activate(h, 'marquee-ellipse')
  drag(ellipse, { x: 40, y: 10 }, { x: 60, y: 30 }, { shift: true })
  assert.equal(h.steps(), 2)
  assert.deepEqual(h.labels(), ['Rectangular Marquee', 'Elliptical Marquee'])
  const selection = h.store.getState().selection
  assert.ok(selection.bounds.x <= 10 && selection.bounds.x + selection.bounds.width >= 60, 'Shift added the ellipse')
  // A click deselects (one step); a click without a selection records nothing.
  rect.pointerDown(pe(5, 40))
  rect.pointerUp(pe(5, 40, { buttons: 0 }))
  assert.equal(h.store.getState().selection, null)
  assert.equal(h.steps(), 3)
  rect.pointerDown(pe(5, 40))
  rect.pointerUp(pe(5, 40, { buttons: 0 }))
  assert.equal(h.steps(), 3)
})

test('marquee: dragging inside the selection moves its outline; arrow nudges coalesce into one step', () => {
  const h = makeHarness()
  const rect = activate(h, 'marquee-rect')
  drag(rect, { x: 10, y: 10 }, { x: 20, y: 20 })
  drag(rect, { x: 15, y: 15 }, { x: 20, y: 18 })
  assert.deepEqual(h.store.getState().selection.bounds, { x: 15, y: 13, width: 10, height: 10 })
  assert.equal(h.steps(), 2)
  for (let i = 0; i < 4; i += 1) {
    assert.equal(rect.keyDown(key('ArrowRight')), true)
    h.tick(100)
  }
  assert.deepEqual(h.store.getState().selection.bounds, { x: 19, y: 13, width: 10, height: 10 })
  assert.equal(h.steps(), 3, 'four nudges are one step')
})

test('lasso and polygonal lasso: one step per closed outline; Backspace, Esc and Enter', () => {
  const h = makeHarness()
  const lasso = activate(h, 'lasso')
  lasso.pointerDown(pe(10, 10))
  for (const [x, y] of [[30, 10], [30, 30], [10, 30]]) lasso.pointerMove(pe(x, y))
  lasso.pointerUp(pe(10, 30, { buttons: 0 }))
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.store.getState().selection.bounds, { x: 10, y: 10, width: 20, height: 20 })
  const polygon = activate(h, 'lasso-polygon')
  for (const [x, y] of [[40, 5], [60, 5], [60, 25], [50, 40]]) {
    polygon.pointerDown(pe(x, y))
    polygon.pointerUp(pe(x, y, { buttons: 0 }))
    h.tick(1000)
    clock += 1000
  }
  assert.equal(polygon.hasSession(), true)
  assert.equal(polygon.keyDown(key('Backspace')), true, 'removes the last vertex')
  assert.equal(polygon.keyDown(key('Enter')), true)
  assert.equal(polygon.hasSession(), false)
  assert.equal(h.steps(), 2)
  const bounds = h.store.getState().selection.bounds
  assert.equal(bounds.y + bounds.height <= 26, true, 'the removed vertex is not part of the outline')
  // Esc cancels an open polygon without a step.
  polygon.pointerDown(pe(5, 5))
  polygon.pointerUp(pe(5, 5, { buttons: 0 }))
  assert.equal(polygon.keyDown(key('Escape')), true)
  assert.equal(h.steps(), 2)
  // Clicking the first vertex closes the polygon.
  for (const [x, y] of [[5, 35], [25, 35], [25, 45], [5.5, 35.5]]) {
    clock += 1000
    polygon.pointerDown(pe(x, y))
    polygon.pointerUp(pe(x, y, { buttons: 0 }))
  }
  assert.equal(polygon.hasSession(), false)
  assert.equal(h.steps(), 3)
  assert.deepEqual(h.labels(), ['Lasso', 'Polygonal Lasso', 'Polygonal Lasso'])
})

test('magic wand: one step after the flood fill, selects the clicked colour region', async () => {
  const h = makeHarness()
  const wand = activate(h, 'magic-wand')
  wand.pointerDown(pe(12, 12))
  wand.pointerUp(pe(12, 12, { buttons: 0 }))
  await idle(wand)
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Magic Wand'])
  assert.deepEqual(h.store.getState().selection.bounds, { x: 8, y: 8, width: 16, height: 16 })
})

test('crop: Enter crops to the drawn box in one step; cropped pixels are deleted or kept', () => {
  const h = makeHarness()
  const crop = activate(h, 'crop')
  assert.equal(crop.hasSession(), false, 'the untouched full-canvas box is not a pending edit')
  drag(crop, { x: 12, y: 12 }, { x: 44, y: 36 })
  assert.equal(crop.hasSession(), true)
  assert.equal(crop.keyDown(key('Enter')), true)
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Crop'])
  const state = h.store.getState()
  assert.equal(state.width, 32)
  assert.equal(state.height, 24)
  assert.deepEqual(h.composite(10, 10), [255, 0, 0, 255], 'the red square moved with the canvas origin')
  assert.equal(h.find(h.layer.id).offsetX, 0, 'trimmed layer starts at the new origin')
  assert.equal(crop.hasSession(), false)
  // Undo restores the canvas.
  h.store.history.undo()
  assert.equal(h.store.getState().width, W)
  // Keep cropped pixels: the layer only shifts.
  h.editor.updateOptions('crop', { deleteCroppedPixels: false })
  drag(crop, { x: 12, y: 12 }, { x: 40, y: 40 })
  assert.equal(crop.hasSession(), true)
  crop.commitSession()
  assert.equal(h.store.getState().width, 28)
  assert.equal(h.find(h.layer.id).offsetX, -4)
  assert.equal(h.find(h.layer.id).surface.contentBounds().width, 16, 'pixels outside the canvas are kept')
})

test('crop: a turned box rotates the layers and crops in one step', async () => {
  const h = makeHarness()
  const crop = activate(h, 'crop')
  drag(crop, { x: 16, y: 12 }, { x: 48, y: 36 })
  // Drag outside the box (away from its handles) to turn it.
  drag(crop, { x: 60, y: 40 }, { x: 55, y: 46 })
  assert.equal(crop.hasSession(), true)
  crop.keyDown(key('Enter'))
  await idle(crop)
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Crop'])
  assert.equal(h.store.getState().width, 32)
  assert.equal(h.store.getState().height, 24)
  assert.equal(h.composite(1, 1)[3], 255, 'the Background stays opaque')
})

test('eyedropper, hand and zoom never record history', () => {
  const h = makeHarness()
  const eyedropper = activate(h, 'eyedropper')
  eyedropper.pointerDown(pe(50, 40))
  eyedropper.pointerUp(pe(50, 40, { buttons: 0 }))
  const expected = h.composite(50, 40)
  assert.deepEqual(h.editor.getState().foreground, { r: expected[0], g: expected[1], b: expected[2] })
  eyedropper.pointerDown(pe(12, 12, { alt: true }))
  eyedropper.pointerUp(pe(12, 12, { alt: true, buttons: 0 }))
  assert.deepEqual(h.editor.getState().background, RED)
  const hand = activate(h, 'hand')
  drag(hand, { x: 10, y: 10 }, { x: 30, y: 25 })
  const total = h.calls.pan.reduce((sum, call) => ({ x: sum.x + call.dx, y: sum.y + call.dy }), { x: 0, y: 0 })
  assert.deepEqual(total, { x: 20, y: 15 })
  const zoom = activate(h, 'zoom')
  zoom.pointerDown(pe(10, 10))
  zoom.pointerUp(pe(10, 10, { buttons: 0 }))
  assert.equal(h.calls.zoom.length, 1)
  assert.ok(h.calls.zoom[0].zoom > 1, 'click zooms in')
  assert.equal(h.steps(), 0)
  assert.equal(h.host.revision, 0)
})

test('move: dragging a layer is one step and only changes its offset; the Background is refused', () => {
  const h = makeHarness()
  const move = activate(h, 'move')
  drag(move, { x: 12, y: 12 }, { x: 22, y: 17 })
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Move'])
  const moved = h.find(h.layer.id)
  assert.deepEqual([moved.offsetX, moved.offsetY], [18, 13])
  assert.equal(moved.surface, h.layer.surface, 'pixels untouched (O(1) move)')
  assert.ok(h.calls.previews.some((preview) => preview && preview.kind === 'layer-pixels'), 'the drag was previewed at the display level')
  assert.equal(h.calls.previews[h.calls.previews.length - 1], null, 'the preview is cleared')
  h.store.transact('Select', 'layer', (tx) => tx.setActiveLayer(h.background.id))
  drag(move, { x: 40, y: 40 }, { x: 45, y: 40 })
  assert.equal(h.steps(), 1, 'the Background does not move')
  assert.ok(h.calls.notes.some((note) => /locked/i.test(note.message)))
})

test('move: Alt-drag copies the layer in one step; Ctrl-click picks the layer under the pointer', () => {
  const h = makeHarness()
  h.store.transact('Select', 'layer', (tx) => tx.setActiveLayer(h.background.id))
  const move = activate(h, 'move')
  drag(move, { x: 12, y: 12 }, { x: 32, y: 12 }, { ctrl: true, alt: true })
  assert.equal(h.steps(), 1)
  const state = h.store.getState()
  assert.equal(state.layers.length, 3)
  const copy = state.layers[2]
  assert.equal(copy.name, 'Layer 1 copy')
  assert.equal(state.activeLayerId, copy.id)
  assert.equal(copy.offsetX, 28)
  assert.equal(h.find(h.layer.id).offsetX, 8, 'the original stays')
})

test('move with a selection floats the pixels: one step per drag, nothing underneath is lost', () => {
  const h = makeHarness()
  h.store.transact('Select', 'layer', (tx) => tx.setActiveLayer(h.background.id))
  const original = Array.from(composite.flattenDocument(h.store.getState()).data)
  const shape = mask.createMaskBuffer(W, H)
  mask.rasterizeRect(shape, { x: 40, y: 30, width: 8, height: 8 }, false)
  shared.commitSelectionShape(h.ctx, 'Rectangular Marquee', 'marquee-rect', shape, 'replace')
  const move = activate(h, 'move')
  const lifted = h.pixel(h.background, 42, 32)
  drag(move, { x: 42, y: 32 }, { x: 22, y: 32 })
  assert.equal(h.steps(), 2)
  assert.deepEqual(h.pixel(h.background, 22, 32), lifted, 'pixels arrived')
  assert.deepEqual(h.pixel(h.background, 42, 32), [255, 255, 255, 255], 'the Background hole takes the background colour')
  assert.deepEqual(h.store.getState().selection.bounds, { x: 20, y: 30, width: 8, height: 8 }, 'the selection moved along')
  // Moving back restores everything exactly: the pixels it passed over were kept.
  drag(move, { x: 22, y: 32 }, { x: 42, y: 32 })
  assert.equal(h.steps(), 3)
  assert.deepEqual(Array.from(composite.flattenDocument(h.store.getState()).data), original)
  // Nudges coalesce.
  for (let i = 0; i < 3; i += 1) {
    move.keyDown(key('ArrowLeft'))
    h.tick(50)
  }
  assert.equal(h.steps(), 4)
  assert.deepEqual(h.store.getState().selection.bounds, { x: 37, y: 30, width: 8, height: 8 })
  assert.deepEqual(h.labels(), ['Rectangular Marquee', 'Move', 'Move', 'Nudge'])
})

test('shape: one step creates an editable shape layer above the active layer', () => {
  const h = makeHarness()
  const shape = activate(h, 'shape')
  drag(shape, { x: 30, y: 20 }, { x: 50, y: 30 })
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Rectangle Tool'])
  const state = h.store.getState()
  const created = state.layers[2]
  assert.equal(created.kind, 'shape')
  assert.equal(created.name, 'Rectangle 1')
  assert.equal(state.activeLayerId, created.id)
  assert.deepEqual(h.composite(40, 25), [0, 0, 255, 255])
  // A click without a drag draws nothing.
  shape.pointerDown(pe(5, 5))
  shape.pointerUp(pe(5, 5, { buttons: 0 }))
  assert.equal(h.steps(), 1)
})

test('free transform: scaling a pixel layer applies in one step; cancelling records nothing', async () => {
  const h = makeHarness()
  const ended = []
  const ft = tools.createFreeTransform({ onEnd: (committed) => ended.push(committed) })
  ft.activate(h.ctx)
  assert.equal(ft.hasSession(), true)
  assert.ok(h.calls.previews.some((p) => p && p.kind === 'layer-props' && p.visible === false), 'the layer is hidden while transforming')
  ft.keyDown(key('Escape'))
  assert.equal(ft.hasSession(), false)
  assert.deepEqual(ended, [false])
  assert.equal(h.steps(), 0)
  ft.activate(h.ctx)
  // Drag the bottom-right handle (24, 24) to (40, 40): proportional doubling about (8, 8).
  drag(ft, { x: 24, y: 24 }, { x: 40, y: 40 })
  ft.keyDown(key('Enter'))
  await idle(ft)
  assert.equal(h.steps(), 1)
  assert.deepEqual(h.labels(), ['Free Transform'])
  assert.deepEqual(ended, [false, true])
  const bounds = shared.layerPixelBounds(h.find(h.layer.id))
  assert.ok(Math.abs(bounds.x - 8) <= 1 && Math.abs(bounds.width - 32) <= 2, `doubled: ${JSON.stringify(bounds)}`)
  assert.deepEqual(h.pixel(h.layer, 30, 30), [255, 0, 0, 255])
})

test('free transform: a pure move only shifts the layer; a selection transforms the floating pixels', async () => {
  const h = makeHarness()
  // At 400% the 7 px handle and reference-point targets are under 2 document pixels.
  h.ctx.view.zoomAt(4, { x: 0, y: 0 })
  const ft = tools.createFreeTransform()
  ft.activate(h.ctx)
  drag(ft, { x: 12, y: 12 }, { x: 22, y: 16 })
  ft.commitSession()
  await idle(ft)
  const layer = h.find(h.layer.id)
  assert.deepEqual([layer.offsetX, layer.offsetY], [18, 12])
  assert.equal(layer.surface, h.layer.surface, 'no resampling for whole-pixel moves')
  const shape = mask.createMaskBuffer(W, H)
  mask.rasterizeRect(shape, { x: 18, y: 12, width: 8, height: 16 }, false)
  shared.commitSelectionShape(h.ctx, 'Rectangular Marquee', 'marquee-rect', shape, 'replace')
  const floating = tools.createFreeTransform()
  floating.activate(h.ctx)
  assert.equal(floating.hasSession(), true)
  // Rotate 90 degrees with Shift (snaps): drag outside from the right of the reference point to below it.
  const info = floating.info()
  floating.setInfo({ angle: 90 })
  assert.equal(Math.round(floating.info().angle), 90)
  floating.keyDown(key('Enter'))
  await idle(floating)
  assert.equal(h.steps(), 3)
  assert.deepEqual(h.labels(), ['Free Transform', 'Rectangular Marquee', 'Free Transform'])
  const selection = h.store.getState().selection
  assert.ok(Math.abs(selection.bounds.width - 16) <= 2 && Math.abs(selection.bounds.height - 8) <= 2, `the selection turned with the pixels: ${JSON.stringify(selection.bounds)}`)
  assert.ok(info.width === 8 && info.height === 16)
})

test('free transform: a shape layer stays a shape (its transform changes)', async () => {
  const h = makeHarness()
  const shapeTool = activate(h, 'shape')
  drag(shapeTool, { x: 30, y: 20 }, { x: 50, y: 30 })
  const ft = tools.createFreeTransform()
  ft.activate(h.ctx)
  assert.equal(ft.canDistort(), false)
  ft.setInfo({ angle: 45 })
  ft.keyDown(key('Enter'))
  await idle(ft)
  assert.equal(h.steps(), 2)
  const layer = h.store.getState().layers[2]
  assert.equal(layer.kind, 'shape')
  assert.ok(Math.abs(Math.atan2(layer.shape.transform[1], layer.shape.transform[0]) * 180 / Math.PI - 45) < 1e-6)
})

test('session tools: deactivating commits a changed crop or polygon instead of dropping it', () => {
  const h = makeHarness()
  const crop = activate(h, 'crop')
  drag(crop, { x: 12, y: 12 }, { x: 44, y: 36 })
  crop.deactivate()
  assert.equal(h.store.getState().width, 32, 'the crop was applied')
  const untouched = activate(h, 'crop')
  untouched.deactivate()
  assert.equal(h.steps(), 1, 'an untouched crop box applies nothing')
  const polygon = activate(h, 'lasso-polygon')
  for (const [x, y] of [[2, 2], [20, 2], [20, 20]]) {
    clock += 1000
    polygon.pointerDown(pe(x, y))
    polygon.pointerUp(pe(x, y, { buttons: 0 }))
  }
  polygon.deactivate()
  assert.equal(h.steps(), 2)
  assert.ok(h.store.getState().selection)
})

test('nextSelectionVersion stays above every version the store hands out', () => {
  const h = makeHarness()
  const state = h.store.getState()
  const a = shared.nextSelectionVersion(state)
  const b = shared.nextSelectionVersion({ selection: { version: a + 5 } })
  assert.ok(b > a + 5)
  assert.ok(selectionModule.selectAll(W, H, b).version === b)
})

test('brush and gradient paint a targeted layer mask in grey', async () => {
  const h = makeHarness()
  h.store.transact('Add Mask', 'layer', (tx) => {
    tx.setMask(h.layer.id, documentModule.createLayerMask())
    tx.setActiveLayer(h.layer.id, 'mask')
  })
  h.editor.update({ foreground: { r: 0, g: 0, b: 0 }, background: WHITE })
  const brush = activate(h, 'brush')
  drag(brush, { x: 10, y: 12 }, { x: 20, y: 12 })
  const maskOf = () => h.find(h.layer.id).mask
  const value = (x, y) => maskOf().surface.read({ x: x - maskOf().offsetX, y: y - maskOf().offsetY, width: 1, height: 1 }).data[0]
  assert.equal(value(15, 12), 0, 'black hides')
  assert.equal(value(15, 20), 255, 'elsewhere still reveals')
  assert.deepEqual(h.pixel(h.layer, 15, 12), [255, 0, 0, 255], 'the pixels themselves are untouched')
  assert.deepEqual(h.composite(15, 12), Array.from(ramp().data.slice((12 * W + 15) * 4, (12 * W + 15) * 4 + 4)), 'the masked area shows the Background')
  const gradient = activate(h, 'gradient')
  h.editor.update({ foreground: { r: 255, g: 255, b: 255 }, background: { r: 0, g: 0, b: 0 } })
  drag(gradient, { x: 0, y: 0 }, { x: 64, y: 0 })
  await idle(gradient)
  assert.ok(value(1, 30) > 245 && value(62, 30) < 10, `white to black: ${value(1, 30)} ${value(62, 30)}`)
  assert.deepEqual(h.labels(), ['Add Mask', 'Brush Tool', 'Gradient'])
})

test('display-level previews equal the committed result (move and floating pixels)', () => {
  const h = makeHarness()
  // A layer move previewed at level 0 shows exactly the moved pixels.
  const state = h.store.getState()
  const source = shared.pixelSourceOf(h.layer)
  const area = { x: 0, y: 0, width: W, height: H }
  const preview = shared.movedLayerPreview(h.layer.id, source, 6, -3, 0, area, false)
  const moved = shared.readLayerRect({ ...h.layer, offsetX: h.layer.offsetX + 6, offsetY: h.layer.offsetY - 3 }, area)
  assert.deepEqual(Array.from(preview.pixels.data), Array.from(moved.data))
  // Floating pixels: the preview of a placement equals the layer after merging there, at level 0 and level 1.
  const shape = mask.createMaskBuffer(W, H)
  mask.rasterizeEllipse(shape, { x: 10, y: 10, width: 12, height: 10 }, true)
  const selection = selectionModule.selectionFromMask(shape, shared.nextSelectionVersion(state))
  const floating = shared.liftFloating(state, h.layer, selection, { cut: true, fill: null })
  const previewer = shared.createFloatingPreviewer(floating)
  const placed = { rect: shared.translateRect(floating.rect, 20, 10), pixels: floating.lifted, key: floating.lifted }
  const level0 = previewer.preview(h.ctx, state, 0, placed, false)
  const level1 = previewer.preview(h.ctx, state, 1, placed, false)
  const result = shared.composeFloating(floating, shared.placedAt(floating, 20, 10))
  h.store.transact('Move', 'move', (tx) => {
    tx.editPixels(floating.layerId, 'pixels').writePixels(result.rect.x - floating.origin.x, result.rect.y - floating.origin.y, result.pixels)
  })
  const after = h.find(h.layer.id)
  assert.deepEqual(Array.from(level0.pixels.data), Array.from(shared.readLayerRect(after, level0.rect).data), 'level 0 is exact')
  const pyramid = load('advanced/pyramid.ts')
  const expected = pyramid.readSurfaceLevel(after.surface, 1, {
    x: level1.rect.x - pyramid.levelOffset(after.offsetX, 1), y: level1.rect.y - pyramid.levelOffset(after.offsetY, 1), width: level1.rect.width, height: level1.rect.height,
  })
  let worst = 0
  for (let i = 0; i < expected.data.length; i += 1) worst = Math.max(worst, Math.abs(expected.data[i] - level1.pixels.data[i]))
  assert.ok(worst <= 2, `level 1 matches the pyramid of the result within 2 levels (worst ${worst})`)
})

test('session tools pan with Space instead of losing their session', async () => {
  const h = makeHarness()
  const crop = activate(h, 'crop')
  drag(crop, { x: 12, y: 12 }, { x: 44, y: 36 })
  assert.equal(crop.keyDown(key(' ', { code: 'Space' })), true, 'Space is consumed by the crop session')
  drag(crop, { x: 30, y: 30 }, { x: 40, y: 34 })
  assert.equal(crop.keyUp(key(' ', { code: 'Space' })), true)
  assert.ok(h.calls.pan.length > 0, 'the view panned')
  assert.equal(crop.hasSession(), true, 'the crop box is still pending')
  crop.keyDown(key('Enter'))
  assert.equal(h.store.getState().width, 32, 'the box was not moved by the pan')
  const ft = tools.createFreeTransform()
  ft.activate(h.ctx)
  assert.equal(ft.keyDown(key(' ', { code: 'Space' })), true)
  const before = h.calls.pan.length
  drag(ft, { x: 5, y: 5 }, { x: 15, y: 5 })
  ft.keyUp(key(' ', { code: 'Space' }))
  assert.ok(h.calls.pan.length > before)
  assert.equal(ft.hasSession(), true)
  ft.keyDown(key('Escape'))
  await idle(ft)
  assert.equal(h.steps(), 1, 'only the crop recorded a step')
})
