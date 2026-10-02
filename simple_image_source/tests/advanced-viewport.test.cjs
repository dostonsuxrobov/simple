'use strict'
// WP3 view math (src/advanced/viewport.ts): pyramid level choice (never a magnified proxy), Photoshop zoom
// ladder and labels, fit / 100% / zoom-at-anchor, coordinate mapping, visible and dirty tile ranges,
// pan constraints and render order.
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

const viewport = load('advanced/viewport.ts')

function close(actual, expected, epsilon = 1e-9, message = '') {
  assert.ok(Math.abs(actual - expected) <= epsilon, `${message} ${actual} vs ${expected}`)
}

test('chooseLevel picks the coarsest level that is never magnified', () => {
  assert.equal(viewport.chooseLevel(1, 1, 6), 0)
  assert.equal(viewport.chooseLevel(4, 1, 6), 0)
  assert.equal(viewport.chooseLevel(0.99, 1, 6), 0)
  assert.equal(viewport.chooseLevel(0.5, 1, 6), 1)
  assert.equal(viewport.chooseLevel(0.4999, 1, 6), 1)
  assert.equal(viewport.chooseLevel(0.25, 1, 6), 2)
  assert.equal(viewport.chooseLevel(1 / 3, 1, 6), 1)
  assert.equal(viewport.chooseLevel(0.01, 1, 6), 6, 'clamped to the deepest level')
  // Device pixels count: at 200% Windows scaling a CSS zoom of 0.5 is 100%.
  assert.equal(viewport.chooseLevel(0.5, 2, 6), 0)
  assert.equal(viewport.chooseLevel(0.25, 2, 6), 1)
  assert.equal(viewport.chooseLevel(0.4, 1.25, 6), 1)
  assert.equal(viewport.chooseLevel(0, 1, 3), 3)
  for (let scale = 0.011; scale < 3; scale *= 1.07) {
    const level = viewport.chooseLevel(scale, 1, 10)
    if (scale >= 1) {
      assert.equal(level, 0, 'magnified views use full resolution')
      continue
    }
    assert.ok(viewport.levelScale(level) >= scale - 1e-12, `level ${level} at ${scale} is not magnified`)
    if (level > 0) assert.ok(viewport.levelScale(level - 1) > scale, 'and is the coarsest such level')
  }
})

test('the zoom ladder, clamping and labels follow Photoshop', () => {
  assert.equal(viewport.ZOOM_STEPS[0], 0.01)
  assert.equal(viewport.ZOOM_STEPS[viewport.ZOOM_STEPS.length - 1], 32)
  assert.ok(Object.isFrozen(viewport.ZOOM_STEPS))
  for (let i = 1; i < viewport.ZOOM_STEPS.length; i += 1) assert.ok(viewport.ZOOM_STEPS[i] > viewport.ZOOM_STEPS[i - 1])
  assert.equal(viewport.stepZoom(1, 1, 1), 2)
  assert.equal(viewport.stepZoom(1, 1, -1), 2 / 3)
  assert.equal(viewport.stepZoom(0.7, 1, -1), 2 / 3)
  assert.equal(viewport.stepZoom(0.7, 1, 1), 1)
  assert.equal(viewport.stepZoom(1 / 3, 1, 1), 0.5, 'an exact step moves to the next one')
  assert.equal(viewport.stepZoom(32, 1, 1), 32)
  assert.equal(viewport.stepZoom(0.01, 1, -1), 0.01)
  // With dpr 2, the steps are device-pixel percentages.
  assert.equal(viewport.stepZoom(0.5, 2, 1), 1)
  assert.equal(viewport.clampZoom(100, 1), 32)
  assert.equal(viewport.clampZoom(0.0001, 2), 0.005)
  assert.equal(viewport.zoomLabel(1, 1), '100%')
  assert.equal(viewport.zoomLabel(0.5, 2), '100%')
  assert.equal(viewport.zoomLabel(1 / 3, 1), '33.33%')
  assert.equal(viewport.zoomLabel(0.125, 1), '12.5%')
  assert.equal(viewport.zoomLabel(16, 1), '1600%')
  assert.equal(viewport.zoomLabel(2 / 3, 1), '66.67%')
  assert.equal(viewport.nearestZoomStep(0.3), 1 / 3)
  assert.equal(viewport.nearestZoomStep(1.4), 1)
  assert.equal(viewport.nearestZoomStep(100), 32)
})

test('zooming around an anchor keeps the document point under it fixed', () => {
  const start = { zoom: 0.37, offsetX: 123.4, offsetY: -56.7 }
  const anchor = { x: 640.5, y: 380.25 }
  const docPoint = viewport.screenToDoc(start, anchor)
  for (const zoom of [0.01, 0.5, 1, 3.7, 32]) {
    const next = viewport.zoomViewAt(start, zoom, anchor)
    assert.equal(next.zoom, zoom)
    const screen = viewport.docToScreen(next, docPoint)
    close(screen.x, anchor.x, 1e-9)
    close(screen.y, anchor.y, 1e-9)
  }
  assert.equal(viewport.zoomViewAt(start, Number.NaN, anchor).zoom, start.zoom)
  const panned = viewport.panView(start, 10, -5)
  assert.deepEqual(panned, { zoom: 0.37, offsetX: 133.4, offsetY: -61.7 })
  // docToScreen and screenToDoc are inverses.
  const p = { x: 17.25, y: -3.5 }
  const back = viewport.screenToDoc(start, viewport.docToScreen(start, p))
  close(back.x, p.x, 1e-9)
  close(back.y, p.y, 1e-9)
})

test('fit, 100% and centring', () => {
  const size = { width: 1000, height: 600, dpr: 1 }
  const fit = viewport.fitView({ width: 4000, height: 3000 }, size, 20)
  close(fit.zoom, (600 - 40) / 3000)
  close(fit.offsetX, (1000 - 4000 * fit.zoom) / 2)
  close(fit.offsetY, 20)
  const small = viewport.fitView({ width: 10, height: 10 }, size, 0)
  assert.equal(small.zoom, 32, 'small images enlarge up to 3200%')
  const hiDpi = viewport.fitView({ width: 100, height: 100 }, { width: 400, height: 400, dpr: 2 }, 0)
  assert.equal(hiDpi.zoom, 4)
  const actual = viewport.actualPixelsView(fit, { width: 1000, height: 600, dpr: 1.5 })
  close(actual.zoom * 1.5, 1)
  const centre = viewport.screenToDoc(fit, { x: 500, y: 300 })
  const after = viewport.docToScreen(actual, centre)
  close(after.x, 500, 1e-9)
  close(after.y, 300, 1e-9)
  const centred = viewport.centerView({ width: 100, height: 50 }, { width: 300, height: 300 }, 2)
  assert.deepEqual(centred, { zoom: 2, offsetX: 50, offsetY: 100 })
  const snapped = viewport.snapView({ zoom: 1, offsetX: 10.3, offsetY: 7.74 }, 2)
  assert.deepEqual(snapped, { zoom: 1, offsetX: 10.5, offsetY: 7.5 })
  assert.ok(viewport.viewsEqual(snapped, { ...snapped }))
  assert.ok(!viewport.viewsEqual(snapped, fit))
})

test('visible tile ranges cover exactly the tiles on screen, per level', () => {
  const doc = { width: 6000, height: 4000 }
  const size = { width: 1000, height: 600 }
  // 100%: the view shows document x 300..1300, y 100..700.
  const view = { zoom: 1, offsetX: -300, offsetY: -100 }
  assert.deepEqual(viewport.visibleTileRange(view, size, doc, 0), { tx0: 1, ty0: 0, tx1: 6, ty1: 3 })
  assert.deepEqual(viewport.visibleTileRange(view, size, doc, 0, 1), { tx0: 0, ty0: 0, tx1: 7, ty1: 4 })
  assert.deepEqual(viewport.visibleTileRange(view, size, doc, 1), { tx0: 0, ty0: 0, tx1: 3, ty1: 2 })
  // A viewport edge exactly on a tile boundary does not pull in the next tile.
  assert.deepEqual(viewport.visibleTileRange({ zoom: 1, offsetX: 0, offsetY: 0 }, { width: 512, height: 256 }, doc, 0), { tx0: 0, ty0: 0, tx1: 2, ty1: 1 })
  // Clipped to the document; empty when it is off screen.
  assert.deepEqual(viewport.visibleTileRange({ zoom: 0.1, offsetX: 0, offsetY: 0 }, size, doc, 3), { tx0: 0, ty0: 0, tx1: 3, ty1: 2 })
  assert.ok(viewport.isEmptyTileRange(viewport.visibleTileRange({ zoom: 1, offsetX: 7000, offsetY: 0 }, size, doc, 0)))
  assert.ok(viewport.isEmptyTileRange(viewport.visibleTileRange({ zoom: 1, offsetX: 0, offsetY: 0 }, { width: 0, height: 10 }, doc, 0)))
  assert.deepEqual(viewport.levelTileCounts(doc, 0), { columns: 24, rows: 16 })
  assert.deepEqual(viewport.levelTileCounts(doc, 2), { columns: 6, rows: 4 })
  assert.deepEqual(viewport.levelTileRect(doc, 2, 5, 3), { x: 1280, y: 768, width: 220, height: 232 })
  assert.deepEqual(viewport.levelTileRect({ width: 10, height: 10 }, 0, 1, 0), { x: 256, y: 0, width: 0, height: 10 })
  const visible = viewport.visibleDocRect({ zoom: 2, offsetX: -100, offsetY: 50 }, { width: 400, height: 300 })
  assert.deepEqual(visible, { x: 50, y: -25, width: 200, height: 150 })
  assert.deepEqual(viewport.documentScreenRect({ zoom: 2, offsetX: -100, offsetY: 50 }, { width: 10, height: 20 }), { x: -100, y: 50, width: 20, height: 40 })
})

test('dirty tile ranges include the proxy pixel a floored offset can move into', () => {
  assert.equal(viewport.dirtyTileRange({ x: 0, y: 0, width: 0, height: 5 }, 0), null)
  assert.deepEqual(viewport.dirtyTileRange({ x: 256, y: 0, width: 1, height: 1 }, 0), { tx0: 1, ty0: 0, tx1: 2, ty1: 1 })
  // At level 1 a change at document x 512 can land on proxy pixel 255 (floor(offset/2) + floor(local/2)).
  assert.deepEqual(viewport.dirtyTileRange({ x: 512, y: 512, width: 2, height: 2 }, 1), { tx0: 0, ty0: 0, tx1: 2, ty1: 2 })
  assert.deepEqual(viewport.dirtyTileRange({ x: -300, y: 10, width: 600, height: 1 }, 0), { tx0: -2, ty0: 0, tx1: 2, ty1: 1 })
  assert.deepEqual(viewport.dirtyTileRange({ x: 1000, y: 1000, width: 3000, height: 10 }, 3), { tx0: 0, ty0: 0, tx1: 2, ty1: 1 })
})

test('constrainView keeps part of the document on screen', () => {
  const doc = { width: 1000, height: 800 }
  const size = { width: 500, height: 400 }
  const lost = viewport.constrainView({ zoom: 1, offsetX: 5000, offsetY: -5000 }, doc, size, 64)
  assert.deepEqual(lost, { zoom: 1, offsetX: 436, offsetY: -736 })
  const fine = { zoom: 1, offsetX: -100, offsetY: -100 }
  assert.equal(viewport.constrainView(fine, doc, size, 64), fine)
  // A tiny document must stay fully reachable.
  const tiny = viewport.constrainView({ zoom: 1, offsetX: -50, offsetY: 900 }, { width: 10, height: 10 }, size, 64)
  assert.deepEqual(tiny, { zoom: 1, offsetX: 0, offsetY: 390 })
})

test('tiles render nearest the viewport centre first', () => {
  const order = viewport.tilesByDistance({ tx0: 0, ty0: 0, tx1: 3, ty1: 3 }, { x: 1.5, y: 1.5 })
  assert.deepEqual(order[0], { tx: 1, ty: 1 })
  assert.deepEqual(order.slice(1, 5), [{ tx: 1, ty: 0 }, { tx: 0, ty: 1 }, { tx: 2, ty: 1 }, { tx: 1, ty: 2 }])
  assert.equal(order.length, 9)
  assert.deepEqual(viewport.tilesByDistance({ tx0: 0, ty0: 0, tx1: 0, ty1: 3 }, { x: 0, y: 0 }), [])
})
