'use strict'
// WP8 Simple mode markup: creation with Shift constraints, hit-testing of strokes, fills, text and
// pictures, handles, move/resize (pictures keep their aspect ratio), restyling, the session undo stack
// and the conversion to the shared vector specs used for baking.
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

const M = load('simple/markupModel.ts')
const V = load('shared/vector.ts')

/** Deterministic text metrics: every character is half the font size wide. */
const measure = (text, style) => ({ width: Array.from(text).length * style.fontSize * 0.5, ascent: style.fontSize * 0.8, descent: style.fontSize * 0.2 })
const style = { ...M.DEFAULT_MARKUP_STYLE, color: '#112233', width: 4, fill: null, fontSize: 20, bold: false }

function near(actual, expected, epsilon = 1e-9, message = '') {
  assert.ok(Math.abs(actual - expected) <= epsilon, `${message} expected ${expected}, got ${actual}`)
}

test('Shift snaps lines and arrows to 45 degrees and makes boxes square', () => {
  const line = M.createShapeItem('a', 'line', { x: 10, y: 10 }, { x: 110, y: 22 }, style, true)
  assert.equal(line.y2, 10, 'nearly horizontal snaps to horizontal')
  near(line.x2, 10 + Math.hypot(100, 12))
  const diagonal = M.createShapeItem('b', 'arrow', { x: 0, y: 0 }, { x: 100, y: 90 }, style, true)
  near(diagonal.x2, diagonal.y2, 1e-9, '45 degrees')
  const box = M.createShapeItem('c', 'rectangle', { x: 50, y: 50 }, { x: 10, y: 80 }, style, true)
  assert.deepEqual([box.x2, box.y2], [10, 90], 'square keeps the drag direction')
  const free = M.createShapeItem('d', 'ellipse', { x: 0, y: 0 }, { x: 30, y: 10 }, style, false)
  assert.deepEqual([free.x2, free.y2], [30, 10])
  assert.equal(M.createShapeItem('e', 'line', { x: 0, y: 0 }, { x: 1, y: 1 }, style, false).fill, null)
  assert.equal(M.isMeaningfulShape(M.createShapeItem('f', 'line', { x: 0, y: 0 }, { x: 1, y: 1 }, style, false)), false)
  assert.equal(M.isMeaningfulShape(free), true)
})

test('strokes are hit within half their width plus the tolerance; fills only when filled', () => {
  const line = M.createShapeItem('l', 'line', { x: 0, y: 50 }, { x: 100, y: 50 }, style, false)
  assert.equal(M.hitItem(line, { x: 50, y: 53 }, 2), true)
  assert.equal(M.hitItem(line, { x: 50, y: 55 }, 2), false)
  assert.equal(M.hitItem(line, { x: 105, y: 50 }, 2), false)
  const arrow = M.createShapeItem('a', 'arrow', { x: 0, y: 0 }, { x: 100, y: 0 }, style, false)
  assert.equal(M.hitItem(arrow, { x: 95, y: 8 }, 1), true, 'the head is part of the arrow')
  const outline = M.createShapeItem('r', 'rectangle', { x: 10, y: 10 }, { x: 110, y: 60 }, style, false)
  assert.equal(M.hitItem(outline, { x: 10, y: 30 }, 1), true, 'on the edge')
  assert.equal(M.hitItem(outline, { x: 60, y: 35 }, 1), false, 'an unfilled interior is not hit')
  const filled = M.createShapeItem('r2', 'rectangle', { x: 10, y: 10 }, { x: 110, y: 60 }, { ...style, fill: '#ff0000' }, false)
  assert.equal(M.hitItem(filled, { x: 60, y: 35 }, 1), true)
  const ring = M.createShapeItem('e', 'ellipse', { x: 0, y: 0 }, { x: 200, y: 100 }, style, false)
  assert.equal(M.hitItem(ring, { x: 200, y: 50 }, 1), true)
  assert.equal(M.hitItem(ring, { x: 100, y: 0 }, 1), true)
  assert.equal(M.hitItem(ring, { x: 100, y: 50 }, 1), false)
  assert.equal(M.hitItem(ring, { x: 5, y: 5 }, 1), false, 'the box corner is outside an ellipse')
})

test('text and pictures are hit inside their boxes; hitTest returns the topmost item', () => {
  const text = M.createTextItem('t', { x: 100, y: 100 }, style, 'Hello')
  const box = M.textBox(text, measure).rect
  assert.deepEqual(box, { x: 100, y: 100, width: 50, height: 24 })
  assert.equal(M.hitItem(text, { x: 140, y: 110 }, 0, measure), true)
  assert.equal(M.hitItem(text, { x: 160, y: 110 }, 0, measure), false)
  const picture = M.createImageItem('p', 'blob:x', 400, 200, { width: 1000, height: 1000 })
  assert.equal(M.hitItem(picture, { x: 500, y: 500 }, 0), true)
  const covering = M.createShapeItem('cover', 'rectangle', { x: 0, y: 0 }, { x: 1000, y: 1000 }, { ...style, fill: '#000000' }, false)
  assert.equal(M.hitTest([picture, covering], { x: 500, y: 500 }, 0, measure).id, 'cover')
  assert.equal(M.hitTest([covering, picture], { x: 500, y: 500 }, 0, measure).id, 'p')
  assert.equal(M.hitTest([], { x: 1, y: 1 }, 0, measure), null)
})

test('handles: line ends, eight box handles, picture corners, none for text', () => {
  const line = M.createShapeItem('l', 'arrow', { x: 0, y: 0 }, { x: 10, y: 10 }, style, false)
  assert.deepEqual(M.handlesOf(line).map((h) => h.handle), ['start', 'end'])
  const box = M.createShapeItem('b', 'rectangle', { x: 0, y: 0 }, { x: 100, y: 50 }, style, false)
  assert.equal(M.handlesOf(box).length, 8)
  assert.equal(M.hitHandle(box, { x: 101, y: 49 }, 4), 'se')
  assert.equal(M.hitHandle(box, { x: 50, y: 25 }, 4), null)
  const picture = M.createImageItem('p', 'blob:x', 100, 100, { width: 1000, height: 1000 })
  assert.deepEqual(M.handlesOf(picture).map((h) => h.handle), ['nw', 'ne', 'se', 'sw'])
  assert.deepEqual(M.handlesOf(M.createTextItem('t', { x: 0, y: 0 }, style, 'x')), [])
})

test('move and resize: Shift squares boxes, pictures keep their aspect ratio', () => {
  const box = M.createShapeItem('b', 'rectangle', { x: 0, y: 0 }, { x: 100, y: 50 }, style, false)
  const moved = M.moveItem(box, 10, -5)
  assert.deepEqual([moved.x1, moved.y1, moved.x2, moved.y2], [10, -5, 110, 45])
  const wider = M.resizeItem(box, 'e', { x: 150, y: 999 }, false)
  assert.deepEqual([wider.x1, wider.y1, wider.x2, wider.y2], [0, 0, 150, 50])
  const square = M.resizeItem(box, 'se', { x: 120, y: 60 }, true)
  near(square.x2 - square.x1, square.y2 - square.y1, 1e-9, 'square')
  const flipped = M.createShapeItem('f', 'ellipse', { x: 100, y: 50 }, { x: 0, y: 0 }, style, false)
  const flippedResized = M.resizeItem(flipped, 'nw', { x: -20, y: -10 }, false)
  assert.deepEqual([flippedResized.x1, flippedResized.y1, flippedResized.x2, flippedResized.y2], [100, 50, -20, -10], 'drawing direction kept')
  const picture = M.createImageItem('p', 'blob:x', 400, 200, { width: 1000, height: 1000 })
  const grown = M.resizeItem(picture, 'se', { x: picture.x + 800, y: picture.y + 10 }, false)
  near(grown.width / grown.height, 2, 1e-9, 'picture aspect')
  assert.equal(grown.x, picture.x)
  const line = M.createShapeItem('l', 'line', { x: 0, y: 0 }, { x: 100, y: 0 }, style, false)
  const end = M.resizeItem(line, 'end', { x: 70, y: 68 }, true)
  near(end.x2, end.y2, 1e-9, 'Shift snaps the moved end to 45 degrees')
  const text = M.createTextItem('t', { x: 5, y: 6 }, style, 'x')
  assert.deepEqual(M.moveItem(text, 1, 2), { ...text, x: 6, y: 8 })
})

test('restyle touches only properties the item has; replace and remove by id', () => {
  const line = M.createShapeItem('l', 'line', { x: 0, y: 0 }, { x: 10, y: 0 }, style, false)
  const restyled = M.restyleItem(line, { color: '#00ff00', width: 500, fill: '#ffffff', fontSize: 99 })
  assert.equal(restyled.color, '#00ff00')
  assert.equal(restyled.width, M.MAX_STROKE_WIDTH)
  assert.equal(restyled.fill, null, 'lines have no fill')
  const text = M.restyleItem(M.createTextItem('t', { x: 0, y: 0 }, style, 'x'), { fontSize: 2, bold: true, width: 9 })
  assert.equal(text.fontSize, M.MIN_FONT_SIZE)
  assert.equal(text.bold, true)
  const items = [line, text]
  assert.equal(M.replaceItem(items, restyled)[0], restyled)
  assert.deepEqual(M.removeItem(items, 'l'), [text])
})

test('bounds include stroke, arrow heads and glyph overhang', () => {
  const arrow = M.createShapeItem('a', 'arrow', { x: 100, y: 100 }, { x: 200, y: 100 }, { ...style, width: 4 }, false)
  const bounds = M.itemBounds(arrow, measure)
  const head = V.arrowHeadLength(4)
  assert.ok(bounds.x <= 100 - head && bounds.x + bounds.width >= 200 + head)
  const text = M.createTextItem('t', { x: 10, y: 10 }, style, 'Hi')
  const textBounds = M.itemBounds(text, measure)
  assert.ok(textBounds.x < 10 && textBounds.y < 10 && textBounds.width > 20)
  const union = M.markupBounds([arrow, text], measure)
  assert.ok(union.x <= textBounds.x && union.x + union.width >= bounds.x + bounds.width)
  assert.equal(M.markupBounds([], measure), null)
})

test('the session stack undoes and redoes item lists', () => {
  let history = M.emptyMarkupHistory()
  const a = M.createTextItem('a', { x: 0, y: 0 }, style, 'a')
  const b = M.createTextItem('b', { x: 0, y: 0 }, style, 'b')
  history = M.commitMarkup(history, [a])
  history = M.commitMarkup(history, [a, b])
  assert.equal(history.past.length, 2)
  history = M.undoMarkup(history)
  assert.deepEqual(history.present, [a])
  history = M.undoMarkup(history)
  assert.deepEqual(history.present, [])
  assert.equal(M.undoMarkup(history), history, 'nothing to undo')
  history = M.redoMarkup(history)
  assert.deepEqual(history.present, [a])
  history = M.commitMarkup(history, [b])
  assert.equal(history.future.length, 0, 'a new change drops the redo branch')
  assert.equal(M.commitMarkup(history, history.present), history)
})

test('items convert to the shared vector specs used for baking', () => {
  const rect = M.createShapeItem('r', 'rectangle', { x: 1, y: 2 }, { x: 30, y: 40 }, { ...style, fill: '#ffffff' }, false)
  const spec = M.toShapeSpec(rect)
  assert.deepEqual(spec.stroke, { r: 0x11, g: 0x22, b: 0x33, a: 255 })
  assert.deepEqual(spec.fill, { r: 255, g: 255, b: 255, a: 255 })
  assert.equal(spec.strokeWidth, 4)
  assert.deepEqual(spec.transform, [1, 0, 0, 1, 0, 0])
  assert.equal(M.toShapeSpec(M.createShapeItem('a', 'arrow', { x: 0, y: 0 }, { x: 9, y: 9 }, style, false)).arrowHeads, 'end')
  assert.equal(M.toShapeSpec(M.createShapeItem('l', 'line', { x: 0, y: 0 }, { x: 9, y: 9 }, style, false)).arrowHeads, 'none')
  const text = M.toTextSpec(M.createTextItem('t', { x: 7, y: 8 }, { ...style, bold: true }, 'Hi'))
  assert.equal(text.style.fontWeight, 700)
  assert.deepEqual(text.transform, [1, 0, 0, 1, 7, 8])
  assert.equal(text.boxWidth, null)
  // The pure rasterizer paints a horizontal 4 px line where the model says it is.
  const raster = V.rasterizeShapePure(M.toShapeSpec(M.createShapeItem('l', 'line', { x: 10, y: 10 }, { x: 50, y: 10 }, style, false)))
  const at = (x, y) => raster.pixels.data[((y - raster.offsetY) * raster.pixels.width + (x - raster.offsetX)) * 4 + 3]
  assert.ok(at(30, 10) > 200 && at(30, 9) > 200, 'inside the stroke')
  assert.equal(at(30, 12), 0, 'outside the stroke')
  assert.deepEqual(M.paintableItems([M.createTextItem('e', { x: 0, y: 0 }, style, '  '), rect]), [rect])
  assert.deepEqual(M.hexToRgb('#abc'), { r: 0xaa, g: 0xbb, b: 0xcc })
})
