const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const Module = require('node:module')
const { pathToFileURL } = require('node:url')
const {
  EDITABLE_EXTENSIONS,
  MIME_BY_EXTENSION,
  SUPPORTED_EXTENSIONS,
  atomicWrite,
  decodeSvgText,
  detectAnimation,
  ensureOutputExtension,
  imageDimensions,
  inspectImageBytes,
  isSupportedExtension,
  matchesSignature,
  outputExtension,
  prepareSvg,
  psdHeader,
  sniffFormat,
  validateImageBytes,
  validateImageDimensions,
} = require('../electron/image-files.cjs')
const { crc32 } = require('../electron/image-metadata.cjs')

// #region fixture builders

function pngChunk(type, data = Buffer.alloc(0)) {
  const header = Buffer.alloc(8)
  header.writeUInt32BE(data.length, 0)
  header.write(type, 4, 'latin1')
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(header.subarray(4), data), 0)
  return Buffer.concat([header, data, crc])
}

function makePng({ width = 4, height = 3, bitDepth = 8, before = [], after = [] } = {}) {
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = bitDepth
  ihdr[9] = 6
  const rowBytes = width * 4 * (bitDepth / 8)
  const raw = Buffer.alloc((rowBytes + 1) * height, 0x7f)
  for (let row = 0; row < height; row += 1) raw[row * (rowBytes + 1)] = 0
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    ...before,
    pngChunk('IDAT', zlib.deflateSync(raw)),
    ...after,
    pngChunk('IEND'),
  ])
}

function makeApng(frames = 2) {
  const actl = Buffer.alloc(8)
  actl.writeUInt32BE(frames, 0)
  const fctl = Buffer.alloc(26)
  const fdat = Buffer.concat([Buffer.alloc(4), zlib.deflateSync(Buffer.alloc(13))])
  return makePng({ before: [pngChunk('acTL', actl), pngChunk('fcTL', fctl)], after: [pngChunk('fcTL', fctl), pngChunk('fdAT', fdat)] })
}

function makeJpeg({ width = 64, height = 48, segments = [] } = {}) {
  const segment = (marker, payload) => {
    const header = Buffer.from([0xff, marker, 0, 0])
    header.writeUInt16BE(payload.length + 2, 2)
    return Buffer.concat([header, payload])
  }
  const sof = Buffer.from([8, 0, 0, 0, 0, 3, 1, 0x11, 0, 2, 0x11, 1, 3, 0x11, 1])
  sof.writeUInt16BE(height, 1)
  sof.writeUInt16BE(width, 3)
  return Buffer.concat([
    Buffer.from([0xff, 0xd8]),
    segment(0xe0, Buffer.from([0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0])),
    ...segments,
    segment(0xdb, Buffer.alloc(65)),
    segment(0xc0, sof),
    segment(0xc4, Buffer.alloc(20)),
    segment(0xda, Buffer.from([1, 1, 0, 0, 0x3f, 0])),
    Buffer.from([0x12, 0x34, 0x56, 0xff, 0xd9]),
  ])
}

function riffChunk(type, data) {
  const header = Buffer.alloc(8)
  header.write(type, 0, 'latin1')
  header.writeUInt32LE(data.length, 4)
  return Buffer.concat(data.length & 1 ? [header, data, Buffer.alloc(1)] : [header, data])
}

function riff(chunks) {
  const payload = Buffer.concat([Buffer.from('WEBP'), ...chunks])
  const header = Buffer.alloc(8)
  header.write('RIFF', 0, 'latin1')
  header.writeUInt32LE(payload.length, 4)
  return Buffer.concat([header, payload])
}

function vp8l(width, height, alpha = false) {
  const data = Buffer.alloc(16)
  data[0] = 0x2f
  data.writeUInt32LE(((width - 1) | ((height - 1) << 14) | (alpha ? 1 << 28 : 0)) >>> 0, 1)
  return data
}

function makeWebp(width = 10, height = 6) {
  return riff([riffChunk('VP8L', vp8l(width, height))])
}

function makeAnimatedWebp(frames = 3, width = 10, height = 6) {
  const vp8x = Buffer.alloc(10)
  vp8x[0] = 0x02 | 0x10
  vp8x.writeUIntLE(width - 1, 4, 3)
  vp8x.writeUIntLE(height - 1, 7, 3)
  const anmf = () => Buffer.concat([Buffer.alloc(16), riffChunk('VP8L', vp8l(width, height))])
  return riff([riffChunk('VP8X', vp8x), riffChunk('ANIM', Buffer.alloc(6)), ...Array.from({ length: frames }, () => riffChunk('ANMF', anmf()))])
}

function makeGif(frames = 1, width = 5, height = 4) {
  const header = Buffer.concat([Buffer.from('GIF89a'), Buffer.from([width, 0, height, 0, 0x80, 0, 0]), Buffer.alloc(6)])
  const frame = Buffer.concat([
    Buffer.from([0x21, 0xf9, 4, 0, 10, 0, 0, 0]),
    Buffer.from([0x2c, 0, 0, 0, 0, width, 0, height, 0, 0]),
    Buffer.from([2, 2, 0x4c, 0x01, 0]),
  ])
  return Buffer.concat([header, Buffer.from([0x21, 0xff, 11]), Buffer.from('NETSCAPE2.0'), Buffer.from([3, 1, 0, 0, 0]), ...Array.from({ length: frames }, () => frame), Buffer.from([0x3b])])
}

function makeIco(sizes = [16, 32, 0]) {
  const header = Buffer.from([0, 0, 1, 0, sizes.length, 0])
  const entries = sizes.map((size) => {
    const entry = Buffer.alloc(16)
    entry[0] = size
    entry[1] = size
    entry.writeUInt32LE(40, 8)
    entry.writeUInt32LE(6 + sizes.length * 16, 12)
    return entry
  })
  return Buffer.concat([header, ...entries, Buffer.alloc(40)])
}

const BMP = (() => {
  const bytes = Buffer.alloc(58)
  bytes.write('BM', 0, 'latin1')
  bytes.writeUInt32LE(40, 14)
  bytes.writeInt32LE(7, 18)
  bytes.writeInt32LE(9, 22)
  return bytes
})()
const AVIF = (() => {
  const ftyp = Buffer.from([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70, 0x61, 0x76, 0x69, 0x66, 0, 0, 0, 0, 0x6d, 0x69, 0x66, 0x31, 0x61, 0x76, 0x69, 0x66])
  const ispe = Buffer.alloc(20)
  ispe.writeUInt32BE(20, 0)
  ispe.write('ispe', 4, 'latin1')
  ispe.writeUInt32BE(320, 12)
  ispe.writeUInt32BE(200, 16)
  return Buffer.concat([ftyp, ispe])
})()
const PSD = (() => {
  const bytes = Buffer.alloc(40)
  bytes.write('8BPS', 0, 'latin1')
  bytes.writeUInt16BE(1, 4)
  bytes.writeUInt16BE(3, 12)
  bytes.writeUInt32BE(30, 14)
  bytes.writeUInt32BE(40, 18)
  bytes.writeUInt16BE(8, 22)
  bytes.writeUInt16BE(3, 24)
  return bytes
})()

const ILLUSTRATOR_SVG = `<?xml version="1.0" encoding="utf-8"?>
<!-- Generator: Adobe Illustrator 27.0.0, SVG Export Plug-In . SVG Version: 6.00 Build 0)  -->
<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd" [
	<!ENTITY ns_extend "http://ns.adobe.com/Extensibility/1.0/">
	<!ENTITY ns_ai "http://ns.adobe.com/AdobeIllustrator/10.0/">
	<!ENTITY ns_graphs "http://ns.adobe.com/Graphs/1.0/">
	<!ENTITY ns_svg "http://www.w3.org/2000/svg">
	<!ENTITY ns_xlink "http://www.w3.org/1999/xlink">
	<!ENTITY st0 "fill:#FF0000;">
]>
<svg version="1.1" id="Layer_1" xmlns="&ns_svg;" xmlns:xlink="&ns_xlink;" xmlns:i="&ns_ai;" x="0px" y="0px"
	 viewBox="0 0 800 400" style="enable-background:new 0 0 800 400;" xml:space="preserve">
<rect style="&st0;" width="800" height="400"/>
</svg>
`

// #endregion fixture builders

const fixtures = {
  '.png': makePng(),
  '.jpg': makeJpeg(),
  '.jpeg': makeJpeg(),
  '.jfif': makeJpeg(),
  '.jpe': makeJpeg(),
  '.jif': makeJpeg(),
  '.apng': makeApng(),
  '.webp': makeWebp(),
  '.gif': makeGif(),
  '.bmp': BMP,
  '.svg': Buffer.from('<?xml version="1.0"?>\n<!-- safe -->\n<svg viewBox="0 0 1 1"></svg>', 'utf8'),
  '.svgz': zlib.gzipSync(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"></svg>')),
  '.avif': AVIF,
  '.ico': makeIco(),
  '.psd': PSD,
}

test('the extension contract is explicit and maps every format to a MIME type', () => {
  assert.deepEqual(SUPPORTED_EXTENSIONS, [
    '.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg', '.avif',
    '.jfif', '.jpe', '.jif', '.apng', '.ico', '.svgz', '.psd',
  ])
  for (const extension of SUPPORTED_EXTENSIONS) {
    assert.equal(isSupportedExtension(`C:\\Pictures\\sample${extension.toUpperCase()}`), true)
    assert.match(MIME_BY_EXTENSION[extension], /^image\//)
    assert.ok(fixtures[extension], `fixture for ${extension}`)
  }
  assert.equal(MIME_BY_EXTENSION['.jfif'], 'image/jpeg')
  assert.equal(MIME_BY_EXTENSION['.apng'], 'image/png')
  assert.equal(MIME_BY_EXTENSION['.svgz'], 'image/svg+xml')
  assert.ok(EDITABLE_EXTENSIONS.includes('.jfif') && !EDITABLE_EXTENSIONS.includes('.gif') && !EDITABLE_EXTENSIONS.includes('.svgz'))
  assert.equal(isSupportedExtension('notes.txt'), false)
})

test('each supported format must match its file signature', () => {
  for (const [extension, bytes] of Object.entries(fixtures)) {
    assert.doesNotThrow(() => validateImageBytes(bytes, extension), extension)
  }
  assert.throws(() => validateImageBytes(fixtures['.png'], '.jpg'), /does not match/)
  assert.throws(() => validateImageBytes(Buffer.alloc(0), '.png'), /empty/)
  assert.throws(() => validateImageBytes(Buffer.from('hello'), '.tiff'), /not supported/)
  assert.throws(() => validateImageBytes(Buffer.concat([Buffer.from('BM'), Buffer.alloc(24)]), '.bmp'), /does not match/)
  const cmyk = Buffer.from(PSD)
  cmyk.writeUInt16BE(4, 24)
  assert.throws(() => validateImageBytes(cmyk, '.psd'), /CMYK color/)
})

test('sniffFormat identifies content by magic bytes alone', () => {
  const expected = {
    '.png': 'png', '.jpg': 'jpeg', '.jfif': 'jpeg', '.apng': 'png', '.webp': 'webp', '.gif': 'gif', '.bmp': 'bmp',
    '.svg': 'svg', '.svgz': 'gzip', '.avif': 'avif', '.ico': 'ico', '.psd': 'psd',
  }
  for (const [extension, format] of Object.entries(expected)) assert.equal(sniffFormat(fixtures[extension]), format, extension)
  assert.equal(sniffFormat(Buffer.from('II*\0\x08\0\0\0', 'latin1')), 'tiff')
  assert.equal(sniffFormat(Buffer.from([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70, 0x68, 0x65, 0x69, 0x63, 0, 0, 0, 0, 0x6d, 0x69, 0x66, 0x31, 0x68, 0x65, 0x69, 0x63])), 'heic')
  assert.equal(sniffFormat(Buffer.from('<html><body>not an image</body></html>')), null)
  assert.equal(sniffFormat(Buffer.from('hello world')), null)
})

test('mislabeled files open by content and are never saved in place', () => {
  const cases = [
    ['webp_named.jpg', makeWebp(), 'webp', 'image/webp', 'webp_named.webp'],
    ['png_named.jpg', makePng(), 'png', 'image/png', 'png_named.png'],
    ['jpeg_named.png', makeJpeg(), 'jpg', 'image/jpeg', 'jpeg_named.jpg'],
    ['gif_named.webp', makeGif(), 'gif', 'image/gif', 'gif_named.png'],
  ]
  for (const [name, bytes, format, mime, suggestion] of cases) {
    const inspected = inspectImageBytes(bytes, name)
    assert.equal(inspected.payloadFormat, format, name)
    assert.equal(inspected.mime, mime, name)
    assert.equal(inspected.mismatch, true, name)
    assert.equal(inspected.canSaveInPlace, false, name)
    assert.equal(inspected.suggestedName, suggestion, name)
    assert.match(inspected.notices[0], /This file is actually an? .* image\. Save will create/, name)
  }
  assert.match(inspectImageBytes(makeWebp(), 'x.jpg').notices[0], /actually a WebP image/)
  assert.throws(() => inspectImageBytes(Buffer.from('plain text'), 'notes.jpg'), /does not match the JPG format, and is not another supported image/)
  assert.throws(() => inspectImageBytes(Buffer.from('II*\0\x08\0\0\0', 'latin1'), 'scan.jpg'), /TIFF images are not supported yet/)
})

test('new extensions open with the right format and save policy', () => {
  for (const name of ['download.jfif', 'camera.jpe', 'photo.jif']) {
    const inspected = inspectImageBytes(makeJpeg(), name)
    assert.equal(inspected.payloadFormat, 'jpg')
    assert.equal(inspected.mismatch, false)
    assert.equal(inspected.canSaveInPlace, true, name)
  }
  assert.equal(inspectImageBytes(makeJpeg(), 'a.jpeg').payloadFormat, 'jpeg')
  const ico = inspectImageBytes(makeIco([16, 48, 0]), 'favicon.ico')
  assert.deepEqual([ico.payloadFormat, ico.width, ico.height, ico.canSaveInPlace, ico.suggestedName], ['ico', 256, 256, false, 'favicon.png'])
  const apng = inspectImageBytes(makeApng(2), 'anim.apng')
  assert.deepEqual([apng.payloadFormat, apng.animated, apng.frameCount, apng.canSaveInPlace], ['png', true, 2, false])
  assert.equal(apng.suggestedName, 'anim (frame 1).png')
  const svgz = inspectImageBytes(fixtures['.svgz'], 'compressed.svgz')
  assert.equal(svgz.payloadFormat, 'svg')
  assert.equal(svgz.canSaveInPlace, false)
  assert.match(svgz.bytes.toString('utf8'), /^<svg /)
  assert.deepEqual([svgz.width, svgz.height], [2048, 1024])
  const psd = inspectImageBytes(PSD, 'poster.psd')
  assert.deepEqual([psd.payloadFormat, psd.width, psd.height, psd.canSaveInPlace], ['psd', 40, 30, true])
})

test('a gzip bomb is rejected quickly', () => {
  const bomb = zlib.gzipSync(Buffer.alloc(256 * 1024 * 1024))
  const started = Date.now()
  assert.throws(() => inspectImageBytes(bomb, 'bomb.svgz'), /32 MB safety limit/)
  assert.ok(Date.now() - started < 5000, 'decompression stops at the cap')
  assert.throws(() => inspectImageBytes(zlib.gzipSync(Buffer.from('not svg')), 'x.png'), /compressed and is not an image/)
})

test('SVG: Illustrator DOCTYPE subsets, UTF-16 text and viewBox-only sizes', () => {
  const illustrator = Buffer.from(ILLUSTRATOR_SVG, 'utf8')
  assert.equal(matchesSignature(illustrator, '.svg'), true)
  assert.doesNotThrow(() => validateImageBytes(illustrator, '.svg'))
  assert.deepEqual(imageDimensions(illustrator, '.svg'), { width: 800, height: 400 })
  const prepared = prepareSvg(illustrator)
  const text = prepared.bytes.toString('utf8')
  assert.doesNotMatch(text, /<!DOCTYPE|<!ENTITY|&ns_svg;|&st0;/)
  assert.match(text, /xmlns="http:\/\/www\.w3\.org\/2000\/svg"/)
  assert.match(text, /style="fill:#FF0000;"/)
  assert.match(text, /<svg [^>]*width="2048" height="1024">/)
  assert.deepEqual(prepared.intrinsic, { width: 800, height: 400 })

  for (const encoding of ['utf16le', 'utf16be']) {
    const source = '<?xml version="1.0" encoding="UTF-16"?>\n<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 60"><text>héllo</text></svg>'
    let bytes = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(source, 'utf16le')])
    if (encoding === 'utf16be') bytes = Buffer.from(bytes).swap16()
    assert.equal(decodeSvgText(bytes), source, encoding)
    assert.equal(sniffFormat(bytes), 'svg', encoding)
    assert.doesNotThrow(() => validateImageBytes(bytes, '.svg'))
    assert.deepEqual(imageDimensions(bytes, '.svg'), { width: 120, height: 60 })
    const out = inspectImageBytes(bytes, 'utf16.svg')
    const decoded = out.bytes.toString('utf8')
    assert.doesNotMatch(decoded, /encoding=/)
    assert.match(decoded, /héllo/)
    assert.deepEqual([out.width, out.height], [2048, 1024])
  }
  const bomless = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>', 'utf16le')
  assert.equal(sniffFormat(bomless), 'svg')

  const icon = inspectImageBytes(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 0h24v24z"/></svg>'), 'icon_viewbox.svg')
  assert.deepEqual([icon.width, icon.height], [2048, 2048])
  assert.deepEqual(icon.intrinsic, { width: 24, height: 24 })
  assert.match(icon.notices.join(' '), /SVG rasterized at 2048 × 2048/)
  const percent = inspectImageBytes(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 800 400"/>'), 'pct.svg')
  assert.deepEqual([percent.width, percent.height], [2048, 1024])
  assert.match(percent.bytes.toString('utf8'), /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg"\s+viewBox="0 0 800 400" width="2048" height="1024"\/>$/)
  const unsized = inspectImageBytes(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><rect width="300" height="150"/></svg>'), 'plain.svg')
  assert.deepEqual([unsized.width, unsized.height], [2048, 1024])
  assert.match(unsized.bytes.toString('utf8'), /viewBox="0 0 300 150"/)
  const huge = inspectImageBytes(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="100000" height="50000"></svg>'), 'poster.svg')
  assert.deepEqual([huge.width, huge.height], [8192, 4096])
})

test('SVG entity handling stays bounded and never resolves external entities', () => {
  const laughs = ['<?xml version="1.0"?>', '<!DOCTYPE svg [', '<!ENTITY a "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa">']
  for (let level = 1; level < 10; level += 1) {
    const previous = String.fromCharCode(96 + level)
    laughs.push(`<!ENTITY ${String.fromCharCode(97 + level)} "${`&${previous};`.repeat(10)}">`)
  }
  laughs.push(']>', '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><text>&j;</text></svg>')
  const started = Date.now()
  let output = null
  try { output = inspectImageBytes(Buffer.from(laughs.join('\n')), 'laughs.svg').bytes } catch (error) { assert.match(error.message, /safety limit/) }
  if (output) assert.ok(output.length < 16 * 1024 * 1024)
  assert.ok(Date.now() - started < 5000)

  const external = '<?xml version="1.0"?><!DOCTYPE svg [<!ENTITY xxe SYSTEM "file:///C:/Windows/win.ini">]><svg xmlns="http://www.w3.org/2000/svg" width="5" height="5"><text>&xxe;</text></svg>'
  const prepared = inspectImageBytes(Buffer.from(external), 'xxe.svg').bytes.toString('utf8')
  assert.doesNotMatch(prepared, /win\.ini|<!ENTITY/)
  assert.match(prepared, /&xxe;/, 'the unresolved reference is left for the parser to reject')
})

test('animation detection: APNG acTL, WebP ANIM/ANMF and multi-frame GIF', () => {
  assert.deepEqual(detectAnimation(makeApng(2)), { animated: true, frameCount: 2 })
  assert.deepEqual(detectAnimation(makePng()), { animated: false, frameCount: 1 })
  assert.deepEqual(detectAnimation(makeAnimatedWebp(3)), { animated: true, frameCount: 3 })
  assert.deepEqual(detectAnimation(makeWebp()), { animated: false, frameCount: 1 })
  assert.deepEqual(detectAnimation(makeGif(3)), { animated: true, frameCount: 3 })
  assert.deepEqual(detectAnimation(makeGif(1)), { animated: false, frameCount: 1 })
  const animatedPng = inspectImageBytes(makeApng(4), 'anim_as_png.png')
  assert.equal(animatedPng.animated, true)
  assert.equal(animatedPng.canSaveInPlace, false)
  assert.match(animatedPng.notices.join(' '), /Animated image \(4 frames\): only the first frame is editable/)
  const webp = inspectImageBytes(makeAnimatedWebp(3), 'anim.webp')
  assert.deepEqual([webp.animated, webp.frameCount, webp.width, webp.height, webp.canSaveInPlace], [true, 3, 10, 6, false])
})

test('dimensions are inspected before renderer decoding and oversized images are rejected', () => {
  const png = Buffer.alloc(24)
  fixtures['.png'].subarray(0, 8).copy(png)
  png.writeUInt32BE(6000, 16)
  png.writeUInt32BE(4000, 20)
  assert.deepEqual(imageDimensions(png, '.png'), { width: 6000, height: 4000 })
  assert.deepEqual(validateImageDimensions(png, '.png'), { width: 6000, height: 4000 })
  assert.deepEqual(imageDimensions(png, '.apng'), { width: 6000, height: 4000 })

  const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0, 17, 8, 0x0f, 0xa0, 0x17, 0x70, 3, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0])
  assert.deepEqual(imageDimensions(jpeg, '.jpg'), { width: 6000, height: 4000 })
  assert.deepEqual(imageDimensions(jpeg, '.jfif'), { width: 6000, height: 4000 })
  assert.deepEqual(imageDimensions(fixtures['.ico'], '.ico'), { width: 256, height: 256 })
  assert.deepEqual(imageDimensions(AVIF, '.avif'), { width: 320, height: 200 })

  const huge = Buffer.from(png)
  huge.writeUInt32BE(30_000, 16)
  assert.throws(() => validateImageDimensions(huge, '.png'), /too large to edit safely/)
  assert.throws(() => inspectImageBytes(huge, 'huge.png'), /too large to edit safely/)
  // A vector image has no fixed pixel size: a huge drawing is rasterized within the limits instead.
  assert.deepEqual(validateImageDimensions(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="100000" height="100000"></svg>'), '.svg'), { width: 7071, height: 7071 })
  assert.throws(() => validateImageDimensions(Buffer.concat([Buffer.from('GIF89a'), Buffer.alloc(4)]), '.gif'), /dimensions could not be read safely/)
})

test('save names retain compatible extensions and add the chosen format when needed', () => {
  assert.equal(ensureOutputExtension('photo.jpeg', 'jpeg'), 'photo.jpeg')
  assert.equal(ensureOutputExtension('photo.jfif', 'jpeg'), 'photo.jfif')
  assert.equal(ensureOutputExtension('anim.apng', 'png'), 'anim.apng')
  assert.equal(ensureOutputExtension('photo', 'jpeg'), 'photo.jpg')
  assert.equal(ensureOutputExtension('drawing', 'png'), 'drawing.png')
  assert.equal(ensureOutputExtension('drawing.jpg', 'png'), 'drawing.jpg.png')
  assert.equal(ensureOutputExtension('poster', 'psd'), 'poster.psd')
  assert.throws(() => ensureOutputExtension('drawing', 'gif'), /PNG, JPEG, WebP, or PSD/)
})

test('atomic writes replace a target without leaving temporary or backup files', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-test-'))
  const target = path.join(directory, 'image.png')
  try {
    await fs.writeFile(target, 'old')
    await atomicWrite(target, Buffer.from('new image bytes'))
    assert.equal(await fs.readFile(target, 'utf8'), 'new image bytes')
    assert.deepEqual(await fs.readdir(directory), ['image.png'])
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
})

// #region electron/main.cjs open handlers with a stubbed electron module

async function loadMainWithStubbedElectron(temporaryRoot) {
  const handlers = new Map()
  const stubWindow = () => new Proxy(function stub() {}, {
    get: (_target, key) => (key === 'then' ? undefined : key === 'isDestroyed' ? () => false : stubWindow()),
    apply: () => stubWindow(),
    construct: () => stubWindow(),
  })
  const electron = {
    app: {
      isPackaged: true,
      requestSingleInstanceLock: () => true,
      on: () => {},
      quit: () => {},
      getVersion: () => '0.0.0-test',
      whenReady: () => Promise.resolve(),
    },
    BrowserWindow: Object.assign(function BrowserWindow() { return stubWindow() }, {
      fromWebContents: () => ({ isDestroyed: () => false }),
      getAllWindows: () => [],
      getFocusedWindow: () => null,
    }),
    clipboard: {
      readImage: () => ({ isEmpty: () => true, toPNG: () => Buffer.alloc(0) }),
      availableFormats: () => [],
    },
    dialog: { showOpenDialog: async () => ({ canceled: true }), showSaveDialog: async () => ({ canceled: true }) },
    ipcMain: { handle: (channel, handler) => handlers.set(channel, handler), on: () => {} },
    nativeImage: {},
    shell: {},
  }
  const previousTemp = [process.env.TEMP, process.env.TMP]
  // main.cjs sweeps stale print folders in the temp directory at startup; keep that inside this test's folder.
  process.env.TEMP = temporaryRoot
  process.env.TMP = temporaryRoot
  const originalLoad = Module._load
  Module._load = function load(request, parent, isMain) {
    if (request === 'electron') return electron
    return originalLoad.call(this, request, parent, isMain)
  }
  try {
    const mainPath = require.resolve('../electron/main.cjs')
    delete require.cache[mainPath]
    require(mainPath)
    await new Promise((resolve) => setTimeout(resolve, 50))
  } finally {
    Module._load = originalLoad
    process.env.TEMP = previousTemp[0]
    process.env.TMP = previousTemp[1]
  }
  const rendererUrl = pathToFileURL(path.join(__dirname, '..', 'dist', 'index.html')).href
  const sender = { id: 1, getURL: () => rendererUrl, once: () => {} }
  const invoke = (channel, ...args) => handlers.get(channel)({ sender, senderFrame: { url: rendererUrl } }, ...args)
  const untrusted = { id: 2, getURL: () => 'https://example.com/', once: () => {} }
  const invokeUntrusted = async (channel, ...args) => handlers.get(channel)({ sender: untrusted, senderFrame: { url: 'https://example.com/' } }, ...args)
  return { handlers, invoke, invokeUntrusted, electron }
}

test('file:open-path and file:open-bytes trust content and report animation', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-open-'))
  try {
    const { invoke } = await loadMainWithStubbedElectron(directory)
    const files = {
      'webp_named.jpg': [makeWebp(), 'webp', false],
      'png_named.jpg': [makePng(), 'png', false],
      'jpeg_named.png': [makeJpeg(), 'jpg', false],
      'download.jfif': [makeJpeg(), 'jpg', true],
      'camera.jpe': [makeJpeg(), 'jpg', true],
      'favicon.ico': [makeIco(), 'ico', false],
      'anim.apng': [makeApng(3), 'png', false],
      'anim_as_png.png': [makeApng(3), 'png', false],
      'compressed.svgz': [fixtures['.svgz'], 'svg', false],
      'still.png': [makePng(), 'png', true],
    }
    for (const [name, [bytes, format, directSave]] of Object.entries(files)) {
      const filePath = path.join(directory, name)
      await fs.writeFile(filePath, bytes)
      const payload = await invoke('file:open-path', filePath)
      assert.equal(payload.format, format, name)
      assert.equal(payload.directSave, directSave, name)
      assert.equal(payload.path, filePath)
      assert.equal(payload.size, bytes.length)
      assert.equal(typeof payload.sourceToken, 'string')
      assert.ok(payload.data instanceof Uint8Array)
    }
    const animated = await invoke('file:open-path', path.join(directory, 'anim_as_png.png'))
    assert.deepEqual([animated.animated, animated.frameCount, animated.suggestedName], [true, 3, 'anim_as_png (frame 1).png'])
    const mismatch = await invoke('file:open-path', path.join(directory, 'webp_named.jpg'))
    assert.deepEqual([mismatch.formatMismatch, mismatch.extension, mismatch.mime, mismatch.suggestedName], [true, 'jpg', 'image/webp', 'webp_named.webp'])
    const dropped = await invoke('file:open-bytes', { name: 'Pasted image.png', data: new Uint8Array(makeAnimatedWebp(3)) })
    assert.deepEqual([dropped.format, dropped.path, dropped.directSave, dropped.animated, dropped.frameCount], ['webp', null, false, true, 3])
    const svg = await invoke('file:open-bytes', { name: 'icon.svg', data: new Uint8Array(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 12"/>')) })
    assert.deepEqual([svg.width, svg.height, svg.intrinsicSize], [2048, 1024, { width: 24, height: 12 }])
    await assert.rejects(invoke('file:open-bytes', { name: 'notes.png', data: new Uint8Array(Buffer.from('hello')) }), /does not match the PNG format/)
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
})

// #endregion

// #region Photoshop documents and clipboard paste (WP7)

function makePsd({ version = 1, channels = 3, width = 40, height = 30, depth = 8, colorMode = 3 } = {}) {
  const bytes = Buffer.alloc(40)
  bytes.write('8BPS', 0, 'latin1')
  bytes.writeUInt16BE(version, 4)
  bytes.writeUInt16BE(channels, 12)
  bytes.writeUInt32BE(height, 14)
  bytes.writeUInt32BE(width, 18)
  bytes.writeUInt16BE(depth, 22)
  bytes.writeUInt16BE(colorMode, 24)
  return bytes
}

/** A real layered PSD (ag-psd writer) so the open and save handlers see a complete file. */
function layeredPsd(color = [10, 20, 30, 255]) {
  const ag = require('ag-psd')
  const solid = (width, height) => {
    const data = new Uint8ClampedArray(width * height * 4)
    for (let i = 0; i < data.length; i += 4) data.set(color, i)
    return { width, height, data }
  }
  return Buffer.from(ag.writePsdUint8Array({
    width: 24,
    height: 16,
    imageData: solid(24, 16),
    children: [{ name: 'Background', left: 0, top: 0, imageData: solid(24, 16) }, { name: 'Layer 1', left: 2, top: 3, opacity: 0.5, imageData: solid(5, 4) }],
  }))
}

test('PSD headers: 8BPS version 1 opens; PSB, CMYK, Lab and other modes are refused with the Photoshop steps', () => {
  const psd = makePsd({ width: 4000, height: 3000 })
  assert.equal(sniffFormat(psd), 'psd')
  assert.doesNotThrow(() => validateImageBytes(psd, '.psd'))
  // Width is read from bytes 18..21 and height from 14..17.
  assert.deepEqual(psdHeader(psd), { version: 1, channels: 3, height: 3000, width: 4000, depth: 8, colorMode: 3 })
  assert.deepEqual(imageDimensions(psd, '.psd'), { width: 4000, height: 3000 })
  assert.deepEqual(validateImageDimensions(psd, '.psd'), { width: 4000, height: 3000 })
  const opened = inspectImageBytes(psd, 'poster.PSD')
  assert.deepEqual([opened.format, opened.payloadFormat, opened.mime, opened.width, opened.height, opened.canSaveInPlace, opened.mismatch],
    ['psd', 'psd', 'image/vnd.adobe.photoshop', 4000, 3000, true, false])
  assert.equal(MIME_BY_EXTENSION['.psd'], 'image/vnd.adobe.photoshop')
  assert.ok(EDITABLE_EXTENSIONS.includes('.psd'))
  for (const depth of [1, 16, 32]) {
    const colorMode = depth === 1 ? 0 : 3
    assert.doesNotThrow(() => inspectImageBytes(makePsd({ depth, colorMode, channels: depth === 1 ? 1 : 3 }), 'deep.psd'), `depth ${depth}`)
  }
  assert.doesNotThrow(() => inspectImageBytes(makePsd({ colorMode: 1, channels: 1 }), 'gray.psd'))
  assert.doesNotThrow(() => inspectImageBytes(makePsd({ colorMode: 2, channels: 1 }), 'indexed.psd'))

  for (const [colorMode, name] of [[4, 'CMYK'], [9, 'Lab'], [8, 'Duotone'], [7, 'Multichannel']]) {
    const message = `This PSD uses ${name} color. In Photoshop choose Image > Mode > RGB Color, then save a copy.`
    assert.throws(() => inspectImageBytes(makePsd({ colorMode }), 'print.psd'), { message })
    assert.throws(() => validateImageBytes(makePsd({ colorMode }), '.psd'), { message })
  }
  assert.throws(() => inspectImageBytes(makePsd({ version: 2 }), 'large.psb'), /Large Photoshop documents \(\.psb\) are not supported/)
  assert.throws(() => inspectImageBytes(makePsd({ version: 2 }), 'renamed.psd'), /\.psb\) are not supported/)
  assert.throws(() => validateImageBytes(makePsd({ version: 2 }), '.psd'), /does not match the PSD format/)
  assert.throws(() => inspectImageBytes(makePsd({ depth: 12 }), 'odd.psd'), /unsupported bit depth/)
  assert.throws(() => inspectImageBytes(makePsd({ depth: 8, colorMode: 0 }), 'bitmap8.psd'), /unsupported bit depth/)
  assert.throws(() => inspectImageBytes(makePsd({ depth: 16, colorMode: 2 }), 'indexed16.psd'), /unsupported bit depth/)
  assert.throws(() => inspectImageBytes(makePsd({ channels: 20 }), 'alphas.psd'), /more than 16 channels/)
  assert.throws(() => inspectImageBytes(makePsd({ channels: 0 }), 'broken.psd'), /damaged/)
  assert.throws(() => inspectImageBytes(makePsd({ width: 20_001, height: 10 }), 'wide.psd'), /too large to edit safely/)
  assert.throws(() => inspectImageBytes(makePsd({ width: 8000, height: 7000 }), 'big.psd'), /50 megapixels and 20,000 pixels per side/)
  assert.throws(() => validateImageDimensions(makePsd({ width: 0 }), '.psd'), /dimensions could not be read safely/)
  assert.throws(() => inspectImageBytes(Buffer.from('8BPS'), 'tiny.psd'), /does not match the PSD format/)
})

test('PSD save names: ensureOutputExtension adds .psd and keeps an existing one', () => {
  assert.equal(outputExtension('psd'), '.psd')
  assert.equal(outputExtension('PSD'), '.psd')
  assert.equal(ensureOutputExtension('art', 'psd'), 'art.psd')
  assert.equal(ensureOutputExtension('art.psd', 'psd'), 'art.psd')
  assert.equal(ensureOutputExtension('C:\\Work\\art.PSD', 'psd'), 'C:\\Work\\art.PSD')
  assert.equal(ensureOutputExtension('art.png', 'psd'), 'art.png.psd')
  assert.equal(ensureOutputExtension('art.psd', 'png'), 'art.psd.png')
})

test('file:open-path opens a PSD, and saving it unchanged writes the original bytes exactly', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-psd-'))
  try {
    const { invoke, electron } = await loadMainWithStubbedElectron(directory)
    const original = layeredPsd()
    const filePath = path.join(directory, 'poster.psd')
    await fs.writeFile(filePath, original)
    const payload = await invoke('file:open-path', filePath)
    assert.deepEqual([payload.format, payload.mime, payload.directSave, payload.width, payload.height, payload.bitDepth, payload.extension],
      ['psd', 'image/vnd.adobe.photoshop', true, 24, 16, 8, 'psd'])
    assert.equal(Buffer.from(payload.data).equals(original), true)

    // Unchanged Save: the renderer sends the original bytes back, and they are written as they are.
    const saved = await invoke('file:save', { data: payload.data, path: payload.path, name: payload.name, format: 'psd', forceDialog: false, sourceToken: payload.sourceToken })
    assert.deepEqual([saved.path, saved.format, saved.size], [filePath, 'psd', original.length])
    assert.deepEqual(saved.metadata, { kept: [], dropped: [] })
    assert.deepEqual(saved.warnings, [])
    assert.equal((await fs.readFile(filePath)).equals(original), true, 'byte-exact')

    // An edited layered document replaces the file in place.
    const edited = layeredPsd([200, 100, 50, 255])
    await invoke('file:save', { data: new Uint8Array(edited), path: filePath, name: 'poster.psd', format: 'psd', forceDialog: false, sourceToken: payload.sourceToken })
    assert.equal((await fs.readFile(filePath)).equals(edited), true)
    assert.deepEqual(await fs.readdir(directory), ['poster.psd'], 'no temporary or backup files are left behind')

    // PSD data is validated like any image; PNG bytes cannot be saved as PSD.
    await assert.rejects(invoke('file:save', { data: new Uint8Array(makePng()), path: filePath, name: 'poster.psd', format: 'psd', forceDialog: false }), /does not match the PSD format/)
    await assert.rejects(invoke('file:save', { data: new Uint8Array(makePsd({ colorMode: 4 })), path: filePath, name: 'poster.psd', format: 'psd', forceDialog: false }), /CMYK color/)
    assert.equal((await fs.readFile(filePath)).equals(edited), true, 'a refused save leaves the file alone')

    // Save As offers a Photoshop filter and a .psd name in the source folder.
    let dialogOptions = null
    electron.dialog.showSaveDialog = async (_window, options) => {
      dialogOptions = options
      return { canceled: false, filePath: path.join(directory, 'copy') }
    }
    const copy = await invoke('file:save', { data: new Uint8Array(edited), path: filePath, name: 'poster.psd', format: 'psd', forceDialog: true })
    assert.deepEqual(dialogOptions.filters, [{ name: 'Photoshop document', extensions: ['psd'] }])
    assert.equal(dialogOptions.defaultPath, path.join(directory, 'poster.psd'))
    assert.equal(copy.path, path.join(directory, 'copy.psd'))
    assert.equal((await fs.readFile(copy.path)).equals(edited), true)

    // A CMYK document is refused with the steps to convert it, before any pixels reach the renderer.
    const cmykPath = path.join(directory, 'print.psd')
    const cmyk = Buffer.from(original)
    cmyk.writeUInt16BE(4, 24)
    await fs.writeFile(cmykPath, cmyk)
    await assert.rejects(invoke('file:open-path', cmykPath), /This PSD uses CMYK color\. In Photoshop choose Image > Mode > RGB Color, then save a copy\./)
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
})

test('clipboard:read-image returns validated PNG bytes; clipboard:has-image reports image formats', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-clipboard-'))
  try {
    const { invoke, invokeUntrusted, electron } = await loadMainWithStubbedElectron(directory)
    assert.equal(await invoke('clipboard:read-image'), null, 'an empty clipboard reads as null')
    assert.equal(await invoke('clipboard:has-image'), false)

    const png = makePng({ width: 7, height: 5 })
    electron.clipboard.readImage = () => ({ isEmpty: () => false, toPNG: () => Buffer.from(png) })
    electron.clipboard.availableFormats = () => ['text/plain', 'image/png']
    const pasted = await invoke('clipboard:read-image')
    assert.ok(pasted instanceof Uint8Array)
    assert.equal(Buffer.from(pasted).equals(png), true)
    assert.equal(pasted.byteLength, pasted.buffer.byteLength, 'a copy of exactly the PNG bytes')
    assert.equal(await invoke('clipboard:has-image'), true)

    electron.clipboard.availableFormats = () => ['text/plain', 'text/html']
    assert.equal(await invoke('clipboard:has-image'), false)
    electron.clipboard.availableFormats = () => { throw new Error('clipboard busy') }
    assert.equal(await invoke('clipboard:has-image'), false)

    // Huge clipboard images are refused like huge files.
    const huge = Buffer.from(png)
    huge.writeUInt32BE(30_000, 16)
    electron.clipboard.readImage = () => ({ isEmpty: () => false, toPNG: () => huge })
    await assert.rejects(async () => invoke('clipboard:read-image'), /too large to edit safely/)
    electron.clipboard.readImage = () => ({ isEmpty: () => false, toPNG: () => Buffer.alloc(0) })
    assert.equal(await invoke('clipboard:read-image'), null)

    // Only the image workspace's own renderer may read the clipboard.
    await assert.rejects(invokeUntrusted('clipboard:read-image'), /did not come from the image workspace/)
    await assert.rejects(invokeUntrusted('clipboard:has-image'), /did not come from the image workspace/)
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
})

test('the preload exposes the clipboard read bridge', async () => {
  const source = await fs.readFile(path.join(__dirname, '..', 'electron', 'preload.cjs'), 'utf8')
  assert.match(source, /readClipboardImage: \(\) => ipcRenderer\.invoke\('clipboard:read-image'\)/)
  assert.match(source, /clipboardHasImage: \(\) => ipcRenderer\.invoke\('clipboard:has-image'\)/)
})

// #endregion
