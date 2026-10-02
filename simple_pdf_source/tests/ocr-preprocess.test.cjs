const test = require('node:test')
const assert = require('node:assert/strict')
const fixture = require('./fixtures/scan-fixture.cjs')

const load = () => import('../electron/ocr-preprocess.mjs')

function gray(width, height, fill = 255) {
  return { width, height, data: new Uint8Array(width * height).fill(fill) }
}

function paint(image, x0, y0, width, height, value) {
  for (let y = y0; y < y0 + height; y += 1) for (let x = x0; x < x0 + width; x += 1) image.data[y * image.width + x] = value
}

function standardDeviation(values) {
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length
  return Math.sqrt(values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / values.length)
}

let raster150
async function textRaster() {
  raster150 ??= (async () => {
    const { bytes } = await fixture.vectorPage()
    return fixture.rasterize(bytes, 150)
  })()
  return raster150
}

test('toGray uses BT.601 luma and composites transparency over white paper', async () => {
  const { toGray } = await load()
  const rgba = Uint8ClampedArray.of(255, 255, 255, 255, 0, 0, 0, 255, 255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255, 0, 0, 0, 0, 0, 0, 0, 128, 90, 140, 200, 255)
  const out = toGray(rgba, 8, 1)
  const expected = [255, 0, 76, 150, 29, 255, 127, Math.round(0.299 * 90 + 0.587 * 140 + 0.114 * 200)]
  out.data.forEach((value, index) => assert.ok(Math.abs(value - expected[index]) <= 1, `pixel ${index}: ${value} vs ${expected[index]}`))
  assert.deepEqual([...toGray(Uint8Array.of(10, 20, 30), 1, 1, 3).data], [Math.round(0.299 * 10 + 0.587 * 20 + 0.114 * 30)])
  assert.deepEqual([...toGray(Uint8Array.of(7, 8), 2, 1, 1).data], [7, 8])
  assert.throws(() => toGray(new Uint8Array(3), 1, 1), /truncated/)
  assert.throws(() => toGray(new Uint8Array(4), 0, 1), /positive integers/)
})

test('normalizeBackground flattens uneven light and stains to under 3 grey levels', async () => {
  const { normalizeBackground } = await load()
  const width = 1200
  const height = 1600
  const image = gray(width, height)
  const paperAt = (x, y) => 230 - 60 * (x / width) - 30 * (y / height) - 35 * Math.exp(-(((x - 600) / 250) ** 2 + ((y - 500) / 150) ** 2))
  const ink = new Uint8Array(width * height)
  for (let y = 0; y < height; y += 1) for (let x = 0; x < width; x += 1) image.data[y * width + x] = Math.round(paperAt(x, y))
  for (let row = 100; row < height - 100; row += 60) {
    for (let x = 80; x < width - 120; x += 55) {
      for (let y = row; y < row + 10; y += 1) for (let xx = x; xx < x + 38; xx += 1) {
        image.data[y * width + xx] = Math.round(paperAt(xx, y) * 0.25)
        ink[y * width + xx] = 1
      }
    }
  }
  const before = []
  const after = []
  const inkAfter = []
  const normalized = normalizeBackground(image)
  for (let y = 4; y < height - 4; y += 3) {
    for (let x = 4; x < width - 4; x += 3) {
      const index = y * width + x
      if (ink[index]) { inkAfter.push(normalized.data[index]); continue }
      // Keep away from ink so only the paper is measured.
      if (ink[index - 3 * width] || ink[index + 3 * width] || ink[index - 3] || ink[index + 3]) continue
      before.push(image.data[index])
      after.push(normalized.data[index])
    }
  }
  assert.ok(standardDeviation(before) > 15, 'fixture background should be uneven')
  assert.ok(standardDeviation(after) < 3, `background sigma after normalisation: ${standardDeviation(after).toFixed(2)}`)
  assert.ok(Math.max(...inkAfter) < 100, 'ink must stay dark')
})

test('stretchContrast maps ink to black and paper to white but never blackens a sparse page', async () => {
  const { stretchContrast, measureInkLevels } = await load()
  const faded = gray(400, 300, 200)
  for (let row = 20; row < 280; row += 20) paint(faded, 20, row, 360, 4, 90)
  const stretched = stretchContrast(faded)
  assert.equal(stretched.data[0], 255)
  assert.ok(stretched.data[22 * 400 + 40] <= 5)

  const sparse = gray(800, 1000)
  paint(sparse, 300, 400, 40, 12, 60)
  assert.equal(measureInkLevels(sparse).ink, 60)
  const sparseOut = stretchContrast(sparse)
  assert.equal(sparseOut.data[0], 255, 'paper must stay white')
  assert.ok(sparseOut.data[405 * 800 + 310] <= 60)

  const blank = gray(300, 300)
  assert.equal(measureInkLevels(blank).ink, null)
  assert.equal(stretchContrast(blank), blank, 'nothing to stretch on a blank page')
})

test('median3 matches a sort-based median, removes specks and keeps 2 px strokes', async () => {
  const { median3 } = await load()
  const random = fixture.createRandom(3)
  const noisy = gray(61, 47)
  for (let i = 0; i < noisy.data.length; i += 1) noisy.data[i] = Math.floor(random() * 256)
  const fast = median3(noisy)
  for (let y = 0; y < noisy.height; y += 1) {
    for (let x = 0; x < noisy.width; x += 1) {
      const index = y * noisy.width + x
      if (x === 0 || y === 0 || x === noisy.width - 1 || y === noisy.height - 1) {
        assert.equal(fast.data[index], noisy.data[index], 'borders are copied')
        continue
      }
      const window = []
      for (let dy = -1; dy <= 1; dy += 1) for (let dx = -1; dx <= 1; dx += 1) window.push(noisy.data[index + dy * noisy.width + dx])
      window.sort((a, b) => a - b)
      assert.equal(fast.data[index], window[4], `pixel ${x},${y}`)
    }
  }

  const page = gray(80, 60)
  for (const [x, y] of [[10, 10], [40, 20], [70, 50], [25, 45]]) page.data[y * 80 + x] = 0
  paint(page, 50, 5, 2, 40, 0)
  paint(page, 5, 30, 30, 2, 0)
  const cleaned = median3(page)
  for (const [x, y] of [[10, 10], [40, 20], [70, 50], [25, 45]]) assert.equal(cleaned.data[y * 80 + x], 255, `speck at ${x},${y}`)
  for (let y = 6; y < 44; y += 1) assert.deepEqual([cleaned.data[y * 80 + 50], cleaned.data[y * 80 + 51]], [0, 0], `vertical stroke row ${y}`)
  for (let x = 6; x < 34; x += 1) assert.deepEqual([cleaned.data[30 * 80 + x], cleaned.data[31 * 80 + x]], [0, 0], `horizontal stroke column ${x}`)
})

test('estimateSkew is within 0.1 degree from -4 to 4 degrees and rotateGray levels the lines', async () => {
  const { estimateSkew, rotateGray } = await load()
  const raster = await textRaster()
  const image = { width: raster.width, height: raster.height, data: raster.data }
  assert.equal(estimateSkew(image), 0)
  for (const degrees of [-4, -2.5, -1.2, 0.7, 2, 3.3, 4]) {
    const skewed = fixture.rotateImage(raster, degrees)
    const estimate = estimateSkew({ width: skewed.width, height: skewed.height, data: skewed.data })
    assert.ok(Math.abs(estimate - degrees) <= 0.1, `skew ${degrees}: estimated ${estimate}`)
    if (Math.abs(degrees) >= 2) {
      const levelled = rotateGray({ width: skewed.width, height: skewed.height, data: skewed.data }, -estimate)
      assert.ok(Math.abs(estimateSkew(levelled)) <= 0.1, `after deskewing ${degrees}`)
    }
  }
  assert.equal(estimateSkew(gray(200, 200)), 0, 'a blank page has no skew')
})

test('deskewPoint maps deskewed OCR points back onto the rendered page', async () => {
  const { rotateGray, deskewPoint } = await load()
  const deskew = { degrees: 3.5, cx: 300, cy: 200 }
  const angle = (-deskew.degrees * Math.PI) / 180
  for (const [x, y] of [[10, 10], [300, 200], [580, 390], [123.4, 321.9]]) {
    const dx = x - deskew.cx
    const dy = y - deskew.cy
    const forward = { x: Math.cos(angle) * dx - Math.sin(angle) * dy + deskew.cx, y: Math.sin(angle) * dx + Math.cos(angle) * dy + deskew.cy }
    const back = deskewPoint(forward.x, forward.y, deskew)
    assert.ok(Math.abs(back.x - x) < 1e-9 && Math.abs(back.y - y) < 1e-9)
  }
  assert.deepEqual(deskewPoint(5, 6, null), { x: 5, y: 6 })

  // Pixel level: a dot found in the rotated image maps back to where it was drawn.
  const image = gray(600, 400)
  paint(image, 450, 90, 4, 4, 0)
  const rotated = rotateGray(image, -deskew.degrees)
  let sumX = 0
  let sumY = 0
  let weight = 0
  for (let y = 0; y < 400; y += 1) {
    for (let x = 0; x < 600; x += 1) {
      const ink = 255 - rotated.data[y * 600 + x]
      sumX += (x + 0.5) * ink
      sumY += (y + 0.5) * ink
      weight += ink
    }
  }
  const mapped = deskewPoint(sumX / weight, sumY / weight, deskew)
  assert.ok(Math.hypot(mapped.x - 452, mapped.y - 92) < 0.25, `dot mapped to ${mapped.x.toFixed(2)},${mapped.y.toFixed(2)}`)
})

test('sauvola keeps thin grey strokes on uneven paper and the ink floor stops thick strokes hollowing', async () => {
  const { sauvola } = await load()
  const width = 300
  const height = 200
  const image = gray(width, height)
  for (let y = 0; y < height; y += 1) for (let x = 0; x < width; x += 1) image.data[y * width + x] = 250 - Math.round(70 * x / width)
  // Thin strokes well short of black, at a level a global threshold would split across the page.
  for (let x = 20; x < 280; x += 1) for (let y = 40; y < 43; y += 1) image.data[y * width + x] = (250 - Math.round(70 * x / width)) - 110
  const binary = sauvola(image)
  for (let x = 25; x < 275; x += 5) assert.equal(binary.data[41 * width + x], 0, `grey stroke at ${x}`)
  let paperBlack = 0
  for (let y = 80; y < 200; y += 1) for (let x = 0; x < width; x += 1) paperBlack += binary.data[y * width + x] === 0 ? 1 : 0
  assert.equal(paperBlack, 0, 'clean paper stays white')

  const block = gray(200, 200)
  paint(block, 40, 40, 120, 120, 30)
  assert.equal(sauvola(block).data[100 * 200 + 100], 255, 'without a floor the middle of a wide stroke turns white')
  assert.equal(sauvola(block, { inkFloor: 64 }).data[100 * 200 + 100], 0)
})

test('isBlank ignores isolated specks but keeps a page with one short word', async () => {
  const { isBlank, countInkPixels, minimumInkPixels } = await load()
  const page = gray(1275, 1650)
  const random = fixture.createRandom(11)
  for (let i = 0; i < 4000; i += 1) page.data[Math.floor(random() * page.data.length)] = 0
  assert.ok(countInkPixels(page) < minimumInkPixels(1275, 1650))
  assert.equal(isBlank(page), true)
  // About two 10 pt characters at 150 DPI.
  paint(page, 600, 800, 3, 20, 0)
  paint(page, 610, 800, 3, 20, 0)
  paint(page, 610, 800, 10, 3, 0)
  assert.equal(isBlank(page), false)
})

test('encodePgm writes an 8-bit P5 image that decodes back', async () => {
  const { encodePgm, decodePgm } = await load()
  const image = { width: 3, height: 2, data: Uint8Array.of(0, 1, 2, 253, 254, 255) }
  const pgm = encodePgm(image)
  const header = 'P5\n3 2\n255\n'
  assert.equal(String.fromCharCode(...pgm.subarray(0, header.length)), header)
  assert.equal(pgm.length, header.length + 6)
  assert.deepEqual(decodePgm(pgm), image)
  assert.throws(() => decodePgm(Uint8Array.of(80, 54, 10)), /PGM/)
})

test('preparePage deskews strong skew only, denoises grain, binarises faint scans and skips blank pages', async () => {
  const { preparePage, estimateSkew } = await load()
  const raster = await textRaster()
  const run = (image, options = {}) => preparePage(image.data, image.width, image.height, { channels: 1, dpi: 150, ...options })

  const strong = run(fixture.rotateImage(raster, 3))
  assert.ok(strong.deskew && Math.abs(strong.deskew.degrees - 3) <= 0.1, JSON.stringify(strong.deskew))
  assert.equal(strong.deskew.cx, raster.width / 2)
  assert.ok(Math.abs(estimateSkew(strong.image)) <= 0.1)
  assert.equal(strong.blank, false)

  const mild = run(fixture.rotateImage(raster, 1.2))
  assert.equal(mild.deskew, null, 'below 1.5 degrees the page is not rotated')
  assert.ok(Math.abs(mild.stats.skewDegrees - 1.2) <= 0.1)

  const grainy = run(fixture.degrade(raster, { noise: 14, specks: 0.0005, seed: 5 }))
  assert.equal(grainy.stats.denoised, true)
  assert.equal(grainy.stats.binarized, false)
  const clean = run(raster)
  assert.equal(clean.stats.denoised, false)
  assert.equal(clean.stats.binarized, false)

  const faded = run(fixture.hardScan(raster, { seed: 2 }))
  assert.ok(faded.stats.contrast < 160, `contrast ${faded.stats.contrast}`)
  assert.equal(faded.stats.binarized, true)
  assert.ok(faded.image.data.every((value) => value === 0 || value === 255))

  const blank = run({ width: 600, height: 800, data: new Uint8Array(600 * 800).fill(240) })
  assert.equal(blank.blank, true)

  const rgba = new Uint8ClampedArray(raster.width * raster.height * 4)
  raster.data.forEach((value, index) => rgba.set([value, value, value, 255], index * 4))
  const fromCanvas = preparePage(rgba, raster.width, raster.height, { dpi: 150 })
  assert.deepEqual(fromCanvas.image.data, clean.image.data, 'RGBA and grey input give the same OCR image')
})

test('engine signature, Tesseract parameters and the content hash are stable', async () => {
  const { engineSignature, tesseractParameters, sha256Hex, PREP_VERSION } = await load()
  assert.equal(engineSignature(), `tjs7|core-simd-lstm|eng-best_int|psm3|${PREP_VERSION}`)
  assert.equal(engineSignature({ language: 'eng+deu', psm: '6' }), `tjs7|core-simd-lstm|eng+deu-best_int|psm6|${PREP_VERSION}`)
  assert.deepEqual(tesseractParameters({ dpi: 299.6 }), { user_defined_dpi: '300', tessedit_pageseg_mode: '3', preserve_interword_spaces: '0' })
  assert.equal(Object.keys(tesseractParameters({ dpi: 300, psm: '7' })).includes('thresholding_method'), false)
  assert.throws(() => tesseractParameters({ dpi: Number.NaN }), /resolution/)
  assert.throws(() => tesseractParameters({ dpi: 300, psm: '1' }), /segmentation/)
  assert.equal(await sha256Hex(new TextEncoder().encode('abc')), 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
})
