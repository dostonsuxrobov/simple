'use strict'
// WP7 PSD mapping (src/advanced/psdMapping.ts) and font names (src/advanced/fonts.ts), pure and DOM-free:
// header refusals, merged-alpha and ICC detection, all 27 blend modes, the 13 adjustment types both ways,
// colour scales, text anchoring for point and paragraph text, bit-depth conversion, the white matte,
// vector-mask rasterization, locks, and PostScript <-> CSS font names.
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

const mapping = load('advanced/psdMapping.ts')
const fonts = load('advanced/fonts.ts')
const { BLEND_MODE_MENU, LIMITS } = load('advanced/types.ts')
const { defaultAdjustment, ADJUSTMENT_TYPES } = load('imaging/adjustments.ts')

const LIMIT = { maxPixels: LIMITS.maxPixels, maxDimension: LIMITS.maxDimension }

// #region helpers

function header({ version = 1, channels = 3, width = 40, height = 30, depth = 8, colorMode = 3 } = {}) {
  const bytes = Buffer.alloc(26)
  bytes.write('8BPS', 0, 'latin1')
  bytes.writeUInt16BE(version, 4)
  bytes.writeUInt16BE(channels, 12)
  bytes.writeUInt32BE(height, 14)
  bytes.writeUInt32BE(width, 18)
  bytes.writeUInt16BE(depth, 22)
  bytes.writeUInt16BE(colorMode, 24)
  return bytes
}

function u32(value) {
  const bytes = Buffer.alloc(4)
  bytes.writeUInt32BE(value >>> 0, 0)
  return bytes
}

function iccV2(description) {
  const text = Buffer.from(`${description}\0`, 'latin1')
  const tag = Buffer.concat([Buffer.from('desc'), Buffer.alloc(4), u32(text.length), text])
  const tagOffset = 128 + 4 + 12
  const table = Buffer.concat([u32(1), Buffer.from('desc'), u32(tagOffset), u32(tag.length)])
  const profile = Buffer.concat([Buffer.alloc(128), table, tag])
  profile.writeUInt32BE(profile.length, 0)
  return profile
}

function iccV4(description) {
  const text = Buffer.from(description, 'utf16le').swap16()
  const tag = Buffer.concat([Buffer.from('mluc'), Buffer.alloc(4), u32(1), u32(12), Buffer.from('enUS'), u32(text.length), u32(28), text])
  const tagOffset = 128 + 4 + 12
  const table = Buffer.concat([u32(1), Buffer.from('desc'), u32(tagOffset), u32(tag.length)])
  return Buffer.concat([Buffer.alloc(128), table, tag])
}

/** A PSD with one image resource (ICC profile) and a layer section whose layer count is `layerCount`. */
function structured({ icc = null, layerCount = 0 } = {}) {
  const resources = []
  if (icc) {
    const data = icc.length % 2 ? Buffer.concat([icc, Buffer.alloc(1)]) : icc
    resources.push(Buffer.concat([Buffer.from('8BIM'), Buffer.from([0x04, 0x0f]), Buffer.from([0, 0]), u32(icc.length), data]))
  }
  const resourceBlock = Buffer.concat(resources)
  const count = Buffer.alloc(2)
  count.writeInt16BE(layerCount, 0)
  const layerInfo = Buffer.concat([u32(2), count])
  return Buffer.concat([header(), u32(0), u32(resourceBlock.length), resourceBlock, u32(layerInfo.length), layerInfo])
}

/** Deterministic text measurement: every character 0.5 em wide, ascent 0.8 em, descent 0.2 em. */
function measure(text, style) {
  return { width: Array.from(text).length * style.fontSize * 0.5 + Array.from(text).length * (style.letterSpacing || 0), ascent: style.fontSize * 0.8, descent: style.fontSize * 0.2 }
}

function textStyle(overrides = {}) {
  return {
    fontFamily: 'Arial', fontSize: 40, fontWeight: 400, italic: false, underline: false,
    color: { r: 12, g: 200, b: 99 }, align: 'left', lineHeight: 1.2, letterSpacing: 0, ...overrides,
  }
}

function approxAffine(actual, expected, tolerance = 1e-6) {
  assert.equal(actual.length, 6)
  for (let i = 0; i < 6; i += 1) assert.ok(Math.abs(actual[i] - expected[i]) <= tolerance, `affine[${i}] ${actual[i]} vs ${expected[i]}`)
}

// #endregion

test('sniffPsd reads the 26-byte header; other bytes are not a PSD', () => {
  assert.deepEqual(mapping.sniffPsd(header({ width: 4000, height: 3000, depth: 16, channels: 4 })), {
    version: 1, channels: 4, width: 4000, height: 3000, depth: 16, colorMode: 3,
  })
  assert.equal(mapping.sniffPsd(new Uint8Array(10)), null)
  assert.equal(mapping.sniffPsd(Buffer.from('8BPX' + ' '.repeat(30))), null)
  const offset = new Uint8Array(40)
  offset.set(header({ width: 7, height: 9 }), 8)
  assert.deepEqual(mapping.sniffPsd(offset.subarray(8)).width, 7, 'subarray views are read at their own offset')
})

test('psdHeaderProblem refuses PSB, CMYK/Lab/Duotone/Multichannel, odd depths, >16 channels and oversize documents', () => {
  assert.equal(mapping.psdHeaderProblem(mapping.sniffPsd(header()), LIMIT), null)
  assert.match(mapping.psdHeaderProblem(mapping.sniffPsd(header({ version: 2 })), LIMIT), /\.psb\) are not supported/)
  for (const [mode, name] of [[4, 'CMYK'], [9, 'Lab'], [8, 'Duotone'], [7, 'Multichannel']]) {
    assert.equal(
      mapping.psdHeaderProblem(mapping.sniffPsd(header({ colorMode: mode })), LIMIT),
      `This PSD uses ${name} color. In Photoshop choose Image > Mode > RGB Color, then save a copy.`,
    )
  }
  assert.match(mapping.psdHeaderProblem(mapping.sniffPsd(header({ depth: 12 })), LIMIT), /unsupported bit depth/)
  assert.match(mapping.psdHeaderProblem(mapping.sniffPsd(header({ colorMode: 2, depth: 16 })), LIMIT), /unsupported bit depth/)
  assert.equal(mapping.psdHeaderProblem(mapping.sniffPsd(header({ colorMode: 0, depth: 1, channels: 1 })), LIMIT), null)
  assert.match(mapping.psdHeaderProblem(mapping.sniffPsd(header({ channels: 20 })), LIMIT), /more than 16 channels/)
  assert.match(mapping.psdHeaderProblem(mapping.sniffPsd(header({ width: 20_001, height: 10 })), LIMIT), /too large to edit safely/)
  assert.match(mapping.psdHeaderProblem(mapping.sniffPsd(header({ width: 8000, height: 8000 })), LIMIT), /50 megapixels/)
  assert.match(mapping.psdHeaderProblem(mapping.sniffPsd(header({ width: 0 })), LIMIT), /damaged/)
})

test('readPsdStructure finds the merged-alpha flag and the ICC profile; ICC descriptions decide sRGB', () => {
  assert.deepEqual(mapping.readPsdStructure(structured({ layerCount: -3 })), { mergedAlpha: true, iccProfile: null })
  assert.equal(mapping.readPsdStructure(structured({ layerCount: 3 })).mergedAlpha, false)
  const adobe = iccV2('Adobe RGB (1998)')
  const found = mapping.readPsdStructure(structured({ icc: adobe, layerCount: 1 }))
  assert.equal(Buffer.from(found.iccProfile).equals(adobe), true)
  assert.equal(mapping.iccProfileDescription(found.iccProfile), 'Adobe RGB (1998)')
  assert.equal(mapping.isSrgbProfile(found.iccProfile), false)
  assert.equal(mapping.isSrgbProfile(iccV2('sRGB IEC61966-2.1')), true)
  assert.equal(mapping.iccProfileDescription(iccV4('Display P3')), 'Display P3')
  assert.equal(mapping.isSrgbProfile(iccV4('sRGB v4 ICC preference')), true)
  assert.equal(mapping.isSrgbProfile(null), true)
  // Truncated or garbage data never throws.
  assert.deepEqual(mapping.readPsdStructure(header()), { mergedAlpha: false, iccProfile: null })
  assert.equal(mapping.iccProfileDescription(new Uint8Array(140)), null)
})

test('resolution: Photoshop stores pixels per inch whatever unit it displays', () => {
  assert.equal(mapping.ppiFromResolution({ horizontalResolution: 300, horizontalResolutionUnit: 'PPCM' }), 300)
  assert.equal(mapping.ppiFromResolution(undefined), 72)
  assert.equal(mapping.ppiFromResolution({ horizontalResolution: 0 }), 72)
  assert.deepEqual(mapping.resolutionFromPpi(150), {
    horizontalResolution: 150, horizontalResolutionUnit: 'PPI', widthUnit: 'Inches',
    verticalResolution: 150, verticalResolutionUnit: 'PPI', heightUnit: 'Inches',
  })
})

test('all 27 blend modes round-trip through ag-psd names; unknown modes fall back to Normal', () => {
  const modes = BLEND_MODE_MENU.filter((mode) => mode !== '-')
  assert.equal(modes.length, 27)
  const names = new Set()
  for (const mode of modes) {
    const psd = mapping.blendToPsd(mode)
    assert.equal(typeof psd, 'string')
    names.add(psd)
    assert.deepEqual(mapping.blendFromPsd(psd), { mode, supported: true })
  }
  assert.equal(names.size, 27)
  assert.equal(mapping.blendToPsd('color-burn'), 'color burn')
  assert.equal(mapping.blendToPsd('linear-dodge'), 'linear dodge')
  assert.deepEqual(mapping.blendFromPsd('pass through'), { mode: 'normal', supported: true })
  assert.deepEqual(mapping.blendFromPsd(undefined), { mode: 'normal', supported: true })
  assert.deepEqual(mapping.blendFromPsd('linear height'), { mode: 'normal', supported: false })
})

test('opacity is quantised like Photoshop stores it', () => {
  assert.equal(mapping.quantizeOpacity(0.5), 128 / 255)
  assert.equal(mapping.quantizeOpacity(1), 1)
  assert.equal(mapping.quantizeOpacity(-1), 0)
  assert.equal(mapping.quantizeOpacity(Number.NaN), 1)
})

test('every Photoshop adjustment Simple supports round-trips through ag-psd structures', () => {
  const specs = [
    { type: 'brightness-contrast', brightness: -40, contrast: 55, legacy: false },
    { type: 'brightness-contrast', brightness: 120, contrast: -50, legacy: true },
    {
      type: 'levels',
      rgb: { inBlack: 10, inWhite: 240, gamma: 1.2, outBlack: 5, outWhite: 250 },
      red: { inBlack: 0, inWhite: 255, gamma: 0.8, outBlack: 0, outWhite: 255 },
      green: { inBlack: 3, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 },
      blue: { inBlack: 0, inWhite: 200, gamma: 1, outBlack: 20, outWhite: 255 },
    },
    {
      type: 'curves',
      rgb: [{ x: 0, y: 0 }, { x: 64, y: 50 }, { x: 128, y: 160 }, { x: 255, y: 255 }],
      red: [{ x: 0, y: 20 }, { x: 255, y: 255 }],
      green: [{ x: 0, y: 0 }, { x: 255, y: 230 }],
      blue: [{ x: 0, y: 0 }, { x: 255, y: 255 }],
    },
    { type: 'exposure', exposure: 1.5, offset: -0.0625, gamma: 0.75 },
    { type: 'vibrance', vibrance: 40, saturation: -10 },
    { type: 'hue-saturation', colorize: false, master: { hue: 20, saturation: -30, lightness: 5 }, ranges: { reds: { hue: -15, saturation: 10, lightness: 0 }, blues: { hue: 0, saturation: 0, lightness: -20 } } },
    { type: 'hue-saturation', colorize: true, master: { hue: 210, saturation: 40, lightness: -10 }, ranges: {} },
    {
      type: 'color-balance',
      shadows: { cyanRed: -10, magentaGreen: 5, yellowBlue: 0 },
      midtones: { cyanRed: 20, magentaGreen: 0, yellowBlue: -30 },
      highlights: { cyanRed: 0, magentaGreen: 0, yellowBlue: 15 },
      preserveLuminosity: false,
    },
    { type: 'black-white', reds: 40, yellows: 60, greens: 40, cyans: 60, blues: 20, magentas: 80, tint: null },
    { type: 'black-white', reds: -50, yellows: 120, greens: 0, cyans: 300, blues: -200, magentas: 10, tint: { r: 225, g: 211, b: 179 } },
    { type: 'photo-filter', color: { r: 236, g: 138, b: 0 }, density: 60, preserveLuminosity: true },
    { type: 'invert' },
    { type: 'posterize', levels: 6 },
    { type: 'threshold', level: 140 },
    {
      type: 'gradient-map',
      stops: [
        { position: 0, color: { r: 20, g: 0, b: 80 }, midpoint: 0.3 },
        { position: 0.4, color: { r: 200, g: 50, b: 0 }, midpoint: 0.7 },
        { position: 1, color: { r: 255, g: 240, b: 200 } },
      ],
      reverse: true,
      dither: true,
    },
  ]
  for (const spec of specs) {
    const psd = mapping.adjustmentToPsd(spec)
    assert.ok(psd && typeof psd.type === 'string', spec.type)
    assert.deepEqual(mapping.adjustmentFromPsd(psd), spec, `${spec.type} round trip`)
  }
  // Every adjustment type except Simple's quick adjust has a Photoshop form; defaults round-trip too.
  for (const type of ADJUSTMENT_TYPES) {
    const spec = defaultAdjustment(type)
    assert.deepEqual(mapping.adjustmentFromPsd(mapping.adjustmentToPsd(spec)), spec, `default ${type}`)
  }
  assert.equal(mapping.adjustmentToPsd(defaultAdjustment('quick')), null)
})

test('adjustments read from Photoshop files: colorize flag, hue range windows, missing channels, unsupported types', () => {
  const colorized = mapping.adjustmentToPsd({ type: 'hue-saturation', colorize: true, master: { hue: 30, saturation: 25, lightness: 0 }, ranges: {} })
  assert.equal(colorized.master.a, 0x0100, 'colorize flag in the high byte of the first word')
  assert.deepEqual([colorized.reds.a, colorized.reds.b, colorized.reds.c, colorized.reds.d], [315, 345, 15, 45])
  assert.deepEqual(mapping.adjustmentFromPsd({ type: 'levels', rgb: { shadowInput: 5, highlightInput: 250, shadowOutput: 0, highlightOutput: 255, midtoneInput: 1.5 } }).red,
    { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 })
  assert.deepEqual(mapping.adjustmentFromPsd({ type: 'curves', rgb: [{ input: 128, output: 100 }, { input: 0, output: 0 }, { input: 255, output: 255 }] }).rgb,
    [{ x: 0, y: 0 }, { x: 128, y: 100 }, { x: 255, y: 255 }])
  assert.equal(mapping.adjustmentFromPsd({ type: 'channel mixer' }), null)
  assert.equal(mapping.adjustmentFromPsd({ type: 'selective color' }), null)
  assert.equal(mapping.adjustmentFromPsd({ type: 'color lookup' }), null)
  assert.equal(mapping.adjustmentFromPsd({ type: 'gradient map', gradientType: 'noise' }), null)
  assert.equal(mapping.adjustmentFromPsd(null), null)
  assert.equal(mapping.psdAdjustmentLabel({ type: 'channel mixer' }), 'Channel Mixer')
  // Photo filter colours written by Photoshop as Lab (normalised binary record).
  const lab = mapping.adjustmentFromPsd({ type: 'photo filter', color: { l: 1, a: 0, b: 0 }, density: 25, preserveLuminosity: true })
  assert.deepEqual(lab.color, { r: 255, g: 255, b: 255 })
  // Photoshop stores a gradient midpoint on the stop that ends its segment.
  const map = mapping.adjustmentFromPsd({
    type: 'gradient map', gradientType: 'solid',
    colorStops: [{ color: { r: 0, g: 0, b: 0 }, location: 0, midpoint: 0.5 }, { color: { r: 255, g: 255, b: 255 }, location: 1, midpoint: 0.25 }],
  })
  assert.equal(map.stops[0].midpoint, 0.25)
})

test('colours convert from every ag-psd scale to 8-bit sRGB', () => {
  assert.deepEqual(mapping.psdColorToRgb({ r: 254.6, g: 0.2, b: 128 }, 'binary'), { r: 255, g: 0, b: 128 })
  assert.deepEqual(mapping.psdColorToRgb({ fr: 1, fg: 0.5, fb: 0 }, 'descriptor'), { r: 255, g: 128, b: 0 })
  assert.deepEqual(mapping.psdColorToRgb({ l: 100, a: 0, b: 0 }, 'descriptor'), { r: 255, g: 255, b: 255 })
  assert.deepEqual(mapping.psdColorToRgb({ l: 0, a: 0, b: 0 }, 'binary'), { r: 0, g: 0, b: 0 })
  assert.deepEqual(mapping.psdColorToRgb({ h: 0, s: 1, b: 1 }, 'binary'), { r: 255, g: 0, b: 0 })
  assert.deepEqual(mapping.psdColorToRgb({ h: 120, s: 100, b: 100 }, 'descriptor'), { r: 0, g: 255, b: 0 })
  assert.deepEqual(mapping.psdColorToRgb({ c: 0, m: 255, y: 255, k: 0 }, 'binary'), { r: 255, g: 0, b: 0 })
  assert.deepEqual(mapping.psdColorToRgb({ k: 128 }, 'binary'), { r: 128, g: 128, b: 128 })
  assert.deepEqual(mapping.psdColorToRgb({ k: 100 }, 'descriptor'), { r: 0, g: 0, b: 0 })
  assert.equal(mapping.psdColorToRgb(undefined, 'binary'), null)
  // Lab mid-gray (L* 53.39 is sRGB 128) lands within one level.
  const gray = mapping.psdColorToRgb({ l: 53.389, a: 0, b: 0 }, 'descriptor')
  assert.ok(Math.abs(gray.r - 128) <= 1 && Math.abs(gray.g - 128) <= 1 && Math.abs(gray.b - 128) <= 1, JSON.stringify(gray))
})

test('point text is anchored at its first baseline by justification; paragraph text at its box', () => {
  const resolve = (name) => ({ ArialMT: { family: 'Arial', weight: 400, italic: false } })[name] ?? { family: 'Fallback', weight: 400, italic: false }
  for (const align of ['left', 'center', 'right']) {
    const spec = { text: 'Hello\nWorld!', style: textStyle({ align, letterSpacing: 2 }), boxWidth: null, transform: [1, 0, 0, 1, 100.5, 40.25] }
    const psd = mapping.textToPsd(spec, { postScriptName: 'ArialMT', fauxBold: false, fauxItalic: false }, measure)
    assert.equal(psd.shapeType, 'point')
    assert.equal(psd.paragraphStyle.justification, align)
    // Baseline of the first line: half leading (40 * 0.1) + ascent (32) = 36; widest line 6 chars.
    const width = 6 * 20 + 6 * 2
    const anchor = align === 'left' ? 0 : align === 'center' ? width / 2 : width
    approxAffine(psd.transform, [1, 0, 0, 1, 100.5 + anchor, 40.25 + 36])
    assert.equal(psd.style.font.name, 'ArialMT')
    assert.equal(psd.style.fontSize, 40)
    assert.equal(psd.style.tracking, 50)
    assert.deepEqual(psd.style.fillColor, { r: 12, g: 200, b: 99 })
    const back = mapping.textFromPsd(psd, resolve, measure)
    assert.deepEqual(back.notes, [])
    assert.equal(back.spec.text, spec.text)
    assert.equal(back.spec.boxWidth, null)
    assert.deepEqual(back.spec.style, spec.style)
    approxAffine(back.spec.transform, spec.transform)
  }
  const rotated = { text: 'Tilt', style: textStyle({ lineHeight: 1.5 }), boxWidth: null, transform: [0.8, 0.6, -0.6, 0.8, 30, 70] }
  const rotatedBack = mapping.textFromPsd(mapping.textToPsd(rotated, { postScriptName: 'ArialMT', fauxBold: false, fauxItalic: false }, measure), resolve, measure)
  approxAffine(rotatedBack.spec.transform, rotated.transform)
  assert.equal(rotatedBack.spec.style.lineHeight, 1.5)

  const paragraph = { text: 'A paragraph that wraps', style: textStyle({ align: 'center' }), boxWidth: 180, transform: [1, 0, 0, 1, 10, 20] }
  const boxed = mapping.textToPsd(paragraph, { postScriptName: 'ArialMT', fauxBold: false, fauxItalic: false }, measure)
  assert.equal(boxed.shapeType, 'box')
  assert.deepEqual(boxed.boxBounds.slice(0, 3), [0, 0, 180])
  approxAffine(boxed.transform, [1, 0, 0, 1, 10, 20])
  const boxedBack = mapping.textFromPsd(boxed, resolve, measure)
  assert.equal(boxedBack.spec.boxWidth, 180)
  approxAffine(boxedBack.spec.transform, paragraph.transform)
})

test('Photoshop text details: faux styles, leading, scales, box offsets and what Simple cannot draw', () => {
  const resolve = () => ({ family: 'Segoe UI', weight: 400, italic: false })
  const { spec, notes } = mapping.textFromPsd({
    text: 'Hi\rthere',
    transform: [2, 0, 0, 2, 50, 60],
    shapeType: 'box',
    boxBounds: [5, 7, 105, 57],
    style: { font: { name: 'SegoeUI' }, fontSize: 24, fauxBold: true, fauxItalic: true, autoLeading: false, leading: 36, horizontalScale: 1, verticalScale: 1, tracking: -100, underline: true, fillColor: { r: 1, g: 2, b: 3 } },
    styleRuns: [{ length: 2, style: { fillColor: { r: 255, g: 0, b: 0 } } }, { length: 6, style: { fillColor: { r: 0, g: 0, b: 255 } } }],
    paragraphStyle: { justification: 'justify-right' },
    warp: { style: 'arc', value: 50 },
  }, resolve, measure)
  assert.equal(spec.text, 'Hi\nthere')
  assert.equal(spec.style.fontWeight, 700)
  assert.equal(spec.style.italic, true)
  assert.equal(spec.style.underline, true)
  assert.equal(spec.style.lineHeight, 1.5)
  assert.equal(spec.style.letterSpacing, -2.4)
  assert.equal(spec.style.align, 'right')
  assert.deepEqual(spec.style.color, { r: 255, g: 0, b: 0 }, 'the first style run wins')
  assert.equal(spec.boxWidth, 100)
  approxAffine(spec.transform, [2, 0, 0, 2, 60, 74])
  assert.ok(notes.some((note) => /mixed character styles/.test(note)))
  assert.ok(notes.some((note) => /warped/.test(note)))
  const faux = mapping.textToPsd({ ...spec, style: { ...spec.style, fontWeight: 700 } }, { postScriptName: 'Impact', fauxBold: true, fauxItalic: false }, measure)
  assert.equal(faux.style.fauxBold, true)
  assert.equal(faux.style.autoLeading, false)
  assert.equal(faux.style.leading, 36)
})

test('bit depths: 16-bit rounds exactly, 32-bit is sRGB-encoded, masks take the gray channel', () => {
  const sixteen = new Uint16Array([0, 257 * 128, 65535, 32768, 128, 65535 - 128, 257, 65535])
  assert.deepEqual(Array.from(mapping.toRgba8(sixteen, 2)), [0, 128, 255, 128, 0, 255, 1, 255])
  const eight = new Uint8ClampedArray([1, 2, 3, 4])
  assert.equal(mapping.toRgba8(eight, 1), eight, '8-bit data is used as is')
  const thirtyTwo = new Float32Array([0, 0.21586, 1, 0.5, 2, -1, 0.0031308, 1])
  assert.deepEqual(Array.from(mapping.toRgba8(thirtyTwo, 2)), [0, 128, 255, 128, 255, 0, 10, 255])
  const mask16 = new Uint16Array([65535, 0, 0, 65535, 257 * 64, 0, 0, 65535])
  assert.deepEqual(Array.from(mapping.maskChannel(mask16, 2, 1).data), [255, 64])
  const mask32 = new Float32Array([0.5, 0, 0, 1, 1, 0, 0, 1])
  assert.deepEqual(Array.from(mapping.maskChannel(mask32, 2, 1).data), [128, 255])
})

test('the merged image white matte is removed, and saved selections never become transparency', () => {
  // Photoshop stores a 50% red pixel of a transparent document as red over white: (255, 127, 127).
  const matted = new Uint8ClampedArray([255, 127, 127, 128, 10, 20, 30, 255, 0, 0, 0, 0])
  mapping.removeWhiteMatte(matted)
  assert.deepEqual(Array.from(matted), [255, 0, 0, 128, 10, 20, 30, 255, 0, 0, 0, 0])
  const selection = new Uint8ClampedArray([1, 2, 3, 0, 4, 5, 6, 99])
  mapping.forceOpaque(selection)
  assert.deepEqual(Array.from(selection), [1, 2, 3, 255, 4, 5, 6, 255])
})

test('vector masks rasterize Bezier paths with combine, subtract, exclude and invert', () => {
  const square = (x0, y0, x1, y1, operation) => ({
    open: false,
    operation,
    fillRule: 'non-zero',
    knots: [[x0, y0], [x1, y0], [x1, y1], [x0, y1]].map(([x, y]) => ({ linked: true, points: [x, y, x, y, x, y] })),
  })
  const area = (mask) => mask.data.reduce((sum, value) => sum + value, 0) / 255
  const one = mapping.rasterizeVectorMask({ paths: [square(10, 10, 30, 20, 'combine')] }, 40, 40)
  assert.ok(Math.abs(area(one) - 200) < 0.5, `area ${area(one)}`)
  assert.equal(one.data[15 * 40 + 20], 255)
  assert.equal(one.data[5 * 40 + 5], 0)
  const cut = mapping.rasterizeVectorMask({ paths: [square(0, 0, 40, 40, 'combine'), square(10, 10, 30, 30, 'subtract')] }, 40, 40)
  assert.ok(Math.abs(area(cut) - (1600 - 400)) < 0.5)
  const xor = mapping.rasterizeVectorMask({ paths: [square(0, 0, 20, 20, 'combine'), square(10, 10, 30, 30, 'exclude')] }, 40, 40)
  assert.ok(Math.abs(area(xor) - (400 + 400 - 200)) < 0.5)
  const inverted = mapping.rasterizeVectorMask({ invert: true, paths: [square(10, 10, 30, 20, 'combine')] }, 40, 40)
  assert.ok(Math.abs(area(inverted) - (1600 - 200)) < 0.5)
  assert.equal(mapping.rasterizeVectorMask({ disable: true, paths: [square(0, 0, 4, 4)] }, 40, 40), null)
  assert.equal(mapping.rasterizeVectorMask({ paths: [] }, 40, 40), null)
  // Curved knots: a circle-ish path covers about pi r^2.
  const k = 0.5523 * 10
  const circle = {
    open: false, fillRule: 'non-zero', operation: 'combine',
    knots: [
      { linked: true, points: [20 - k, 10, 20, 10, 20 + k, 10] },
      { linked: true, points: [30, 20 - k, 30, 20, 30, 20 + k] },
      { linked: true, points: [20 + k, 30, 20, 30, 20 - k, 30] },
      { linked: true, points: [10, 20 + k, 10, 20, 10, 20 - k] },
    ],
  }
  const round = mapping.rasterizeVectorMask({ paths: [circle] }, 40, 40)
  assert.ok(Math.abs(area(round) - Math.PI * 100) / (Math.PI * 100) < 0.02, `circle area ${area(round)}`)
})

test('locks and advanced blending notes', () => {
  assert.deepEqual(mapping.locksFromPsd({ transparencyProtected: true }), { pixels: false, position: false, transparency: true })
  assert.deepEqual(mapping.locksFromPsd({ protected: { composite: true, position: true } }), { pixels: true, position: true, transparency: false })
  assert.deepEqual(mapping.locksFromPsd(mapping.locksToPsd({ pixels: true, position: false, transparency: true })), { pixels: true, position: false, transparency: true })
  assert.deepEqual(mapping.advancedBlendingNotes({}), [])
  const defaults = { compositeGrayBlendSource: [0, 0, 255, 255], compositeGraphBlendDestinationRange: [0, 0, 255, 255], ranges: [{ sourceRange: [0, 0, 255, 255], destRange: [0, 0, 255, 255] }] }
  assert.deepEqual(mapping.advancedBlendingNotes({ blendingRanges: defaults, blendClippendElements: true, transparencyShapesLayer: true }), [])
  assert.deepEqual(mapping.advancedBlendingNotes({ blendingRanges: { ...defaults, compositeGrayBlendSource: [30, 60, 255, 255] }, knockout: true }), ['Blend If', 'knockout'])
  assert.equal(mapping.hasEnabledEffects({ effects: { dropShadow: [{ enabled: true }] } }), true)
  assert.equal(mapping.hasEnabledEffects({ effects: { disabled: true, dropShadow: [{ enabled: true }] } }), false)
  assert.equal(mapping.hasEnabledEffects({ effects: { stroke: [{ enabled: false }] } }), false)
})

test('fonts: PostScript names resolve through the table, then installed fonts, then a heuristic', () => {
  fonts.clearLocalFonts()
  assert.deepEqual(fonts.resolvePostScriptFont('ArialMT'), { family: 'Arial', weight: 400, italic: false })
  assert.deepEqual(fonts.resolvePostScriptFont('TimesNewRomanPS-BoldItalicMT'), { family: 'Times New Roman', weight: 700, italic: true })
  assert.deepEqual(fonts.resolvePostScriptFont('segoeui-bold'), { family: 'Segoe UI', weight: 700, italic: false })
  assert.deepEqual(fonts.resolvePostScriptFont('SegoeUI-Semibold'), { family: 'Segoe UI', weight: 700, italic: false })
  assert.deepEqual(fonts.resolvePostScriptFont('MyriadPro-Regular'), { family: 'Myriad Pro', weight: 400, italic: false })
  assert.deepEqual(fonts.resolvePostScriptFont('Montserrat-SemiBoldItalic'), { family: 'Montserrat', weight: 700, italic: true })
  assert.deepEqual(fonts.resolvePostScriptFont('Roboto-LightItalic'), { family: 'Roboto', weight: 400, italic: true })
  assert.deepEqual(fonts.resolvePostScriptFont('HelveticaNeue'), { family: 'Helvetica Neue', weight: 400, italic: false })
  assert.deepEqual(fonts.resolvePostScriptFont('SourceSansPro-BoldIt'), { family: 'Source Sans Pro', weight: 700, italic: true })
  assert.deepEqual(fonts.resolvePostScriptFont(''), { family: 'Arial', weight: 400, italic: false })
  fonts.setLocalFonts([
    { postscriptName: 'Inter-Regular', fullName: 'Inter Regular', family: 'Inter', style: 'Regular' },
    { postscriptName: 'Inter-SemiBold', fullName: 'Inter SemiBold', family: 'Inter', style: 'SemiBold' },
    { postscriptName: 'ArialMT', fullName: 'Arial', family: 'Arial', style: 'Regular' },
  ])
  assert.equal(fonts.localFontsLoaded(), true)
  assert.deepEqual(fonts.resolvePostScriptFont('Inter-SemiBold'), { family: 'Inter', weight: 700, italic: false })
  assert.ok(fonts.availableFontFamilies().includes('Inter'))
  assert.ok(fonts.availableFontFamilies().includes('Segoe UI'))
  assert.deepEqual(fonts.postScriptNameFor({ fontFamily: 'Inter', fontWeight: 700, italic: false }), { postScriptName: 'Inter-SemiBold', fauxBold: false, fauxItalic: false })
  assert.deepEqual(fonts.postScriptNameFor({ fontFamily: 'Inter', fontWeight: 400, italic: true }), { postScriptName: 'Inter-Regular', fauxBold: false, fauxItalic: true })
  fonts.clearLocalFonts()
})

test('fonts: CSS family + weight + italic -> the PostScript name Photoshop looks up', () => {
  fonts.clearLocalFonts()
  assert.deepEqual(fonts.postScriptNameFor({ fontFamily: 'Arial', fontWeight: 700, italic: false }), { postScriptName: 'Arial-BoldMT', fauxBold: false, fauxItalic: false })
  assert.deepEqual(fonts.postScriptNameFor({ fontFamily: 'Segoe UI', fontWeight: 400, italic: false }), { postScriptName: 'SegoeUI', fauxBold: false, fauxItalic: false })
  assert.deepEqual(fonts.postScriptNameFor({ fontFamily: 'Times New Roman', fontWeight: 400, italic: true }), { postScriptName: 'TimesNewRomanPS-ItalicMT', fauxBold: false, fauxItalic: false })
  assert.deepEqual(fonts.postScriptNameFor({ fontFamily: 'Impact', fontWeight: 700, italic: true }), { postScriptName: 'Impact', fauxBold: true, fauxItalic: true })
  assert.deepEqual(fonts.postScriptNameFor({ fontFamily: 'Fancy Font', fontWeight: 700, italic: true }), { postScriptName: 'FancyFont-BoldItalic', fauxBold: false, fauxItalic: false })
  // Every curated family resolves back to itself.
  for (const family of ['Segoe UI', 'Arial', 'Calibri', 'Cambria', 'Candara', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Georgia', 'Impact', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Bahnschrift', 'Segoe Print', 'Segoe Script', 'Sitka Text']) {
    for (const [weight, italic] of [[400, false], [700, false], [400, true], [700, true]]) {
      const choice = fonts.postScriptNameFor({ fontFamily: family, fontWeight: weight, italic })
      const resolved = fonts.resolvePostScriptFont(choice.postScriptName)
      assert.equal(resolved.family, family, `${family} ${weight} ${italic}`)
      assert.equal(choice.fauxBold ? 700 : resolved.weight, weight, `${family} ${weight} ${italic} weight`)
      assert.equal(choice.fauxItalic || resolved.italic, italic, `${family} ${weight} ${italic} italic`)
    }
  }
})

test('loadLocalFonts resolves false without the API and never throws', async () => {
  fonts.clearLocalFonts()
  assert.equal(await fonts.loadLocalFonts(10), false)
  const previous = globalThis.queryLocalFonts
  try {
    globalThis.queryLocalFonts = async () => { throw new Error('SecurityError: needs a user gesture') }
    assert.equal(await fonts.loadLocalFonts(10), false)
    globalThis.queryLocalFonts = async () => [{ postscriptName: 'Foo-Bold', fullName: 'Foo Bold', family: 'Foo', style: 'Bold' }]
    assert.equal(await fonts.loadLocalFonts(100), true)
    assert.deepEqual(fonts.resolvePostScriptFont('Foo-Bold'), { family: 'Foo', weight: 700, italic: false })
  } finally {
    if (previous === undefined) delete globalThis.queryLocalFonts
    else globalThis.queryLocalFonts = previous
    fonts.clearLocalFonts()
  }
})
