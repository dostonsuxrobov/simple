const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const {
  MIME_BY_EXTENSION,
  SUPPORTED_EXTENSIONS,
  atomicWrite,
  ensureOutputExtension,
  imageDimensions,
  isSupportedExtension,
  validateImageBytes,
  validateImageDimensions,
} = require('../electron/image-files.cjs')

const fixtures = {
  '.png': Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0]),
  '.jpg': Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 0, 0xff, 0xd9]),
  '.jpeg': Buffer.from([0xff, 0xd8, 0xff, 0xe1, 0, 0, 0xff, 0xd9]),
  '.webp': Buffer.from('RIFF0000WEBPVP8 ', 'ascii'),
  '.gif': Buffer.from('GIF89a0000', 'ascii'),
  '.bmp': Buffer.concat([Buffer.from('BM', 'ascii'), Buffer.alloc(24)]),
  '.svg': Buffer.from('<?xml version="1.0"?>\n<!-- safe -->\n<svg viewBox="0 0 1 1"></svg>', 'utf8'),
  '.avif': Buffer.from([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70, 0x61, 0x76, 0x69, 0x66, 0, 0, 0, 0]),
}

test('the extension contract is explicit and maps every format to a MIME type', () => {
  assert.deepEqual(SUPPORTED_EXTENSIONS, ['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg', '.avif'])
  for (const extension of SUPPORTED_EXTENSIONS) {
    assert.equal(isSupportedExtension(`C:\\Pictures\\sample${extension}`), true)
    assert.match(MIME_BY_EXTENSION[extension], /^image\//)
  }
  assert.equal(isSupportedExtension('notes.txt'), false)
})

test('each supported format must match its file signature', () => {
  for (const [extension, bytes] of Object.entries(fixtures)) {
    assert.doesNotThrow(() => validateImageBytes(bytes, extension), extension)
  }
  assert.throws(() => validateImageBytes(fixtures['.png'], '.jpg'), /does not match/)
  assert.throws(() => validateImageBytes(Buffer.alloc(0), '.png'), /empty/)
  assert.throws(() => validateImageBytes(Buffer.from('hello'), '.tiff'), /not supported/)
})

test('dimensions are inspected before renderer decoding and oversized images are rejected', () => {
  const png = Buffer.alloc(24)
  fixtures['.png'].subarray(0, 8).copy(png)
  png.writeUInt32BE(6000, 16)
  png.writeUInt32BE(4000, 20)
  assert.deepEqual(imageDimensions(png, '.png'), { width: 6000, height: 4000 })
  assert.deepEqual(validateImageDimensions(png, '.png'), { width: 6000, height: 4000 })

  const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0, 17, 8, 0x0f, 0xa0, 0x17, 0x70, 3, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0])
  assert.deepEqual(imageDimensions(jpeg, '.jpg'), { width: 6000, height: 4000 })

  const oversizedSvg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="100000" height="100000"></svg>')
  assert.throws(() => validateImageDimensions(oversizedSvg, '.svg'), /too large to edit safely/)
  assert.throws(() => validateImageDimensions(Buffer.concat([Buffer.from('GIF89a'), Buffer.alloc(4)]), '.gif'), /dimensions could not be read safely/)
})

test('save names retain compatible extensions and add the chosen format when needed', () => {
  assert.equal(ensureOutputExtension('photo.jpeg', 'jpeg'), 'photo.jpeg')
  assert.equal(ensureOutputExtension('photo', 'jpeg'), 'photo.jpg')
  assert.equal(ensureOutputExtension('drawing', 'png'), 'drawing.png')
  assert.equal(ensureOutputExtension('drawing.jpg', 'png'), 'drawing.jpg.png')
  assert.throws(() => ensureOutputExtension('drawing', 'gif'), /PNG, JPEG, or WebP/)
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
