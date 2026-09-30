'use strict'

const assert = require('node:assert/strict')
const path = require('node:path')
const test = require('node:test')
const {
  assertCopyTarget,
  assertFrameTarget,
  frameTimestamp,
  suggestedCopyName,
  suggestedFrameName,
  validateFrameBytes,
} = require('../electron/video-export.cjs')

test('export names retain the real format and include a deterministic frame timestamp', () => {
  assert.equal(suggestedCopyName('C:\\Media\\My clip.MP4'), 'My clip copy.MP4')
  assert.equal(frameTimestamp(3723.9), '01-02-03.900')
  assert.equal(frameTimestamp(-5), '00-00-00')
  assert.equal(suggestedFrameName('C:\\Media\\My clip.MP4', 'png', 62), 'My clip frame 00-01-02.png')
  assert.equal(suggestedFrameName('/media/clip.webm', 'jpeg', 0), 'clip frame 00-00-00.jpg')
  assert.notEqual(suggestedFrameName('clip.webm', 'png', 1.25), suggestedFrameName('clip.webm', 'png', 1.75))
  assert.equal(frameTimestamp(59.9998), '00-01-00')
})

test('a copied video cannot masquerade as another format or overwrite its source', () => {
  const source = path.resolve('C:\\Media\\clip.mp4')
  assert.doesNotThrow(() => assertCopyTarget(source, path.resolve('C:\\Media\\clip copy.MP4')))
  assert.throws(() => assertCopyTarget(source, path.resolve('C:\\Media\\clip.webm')), /must keep its original .MP4 extension/)
  assert.throws(() => assertCopyTarget(source, source.toUpperCase()), /Choose a new filename/)
})

test('still-image targets use only their actual encoded extension', () => {
  assert.doesNotThrow(() => assertFrameTarget('frame.PNG', 'png'))
  assert.doesNotThrow(() => assertFrameTarget('frame.jpeg', 'jpeg'))
  assert.throws(() => assertFrameTarget('frame.jpg', 'png'), /must end in .png/)
  assert.throws(() => assertFrameTarget('frame.png', 'jpeg'), /must end in .jpg or .jpeg/)
})

test('encoded frame payloads are checked before being written', () => {
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  const jpeg = Uint8Array.from([0xff, 0xd8, 0xff, 0xd9])
  assert.deepEqual(validateFrameBytes(png, 'png'), Buffer.from(png))
  assert.deepEqual(validateFrameBytes(jpeg.buffer, 'jpeg'), Buffer.from(jpeg))
  assert.throws(() => validateFrameBytes(jpeg, 'png'), /not a valid PNG/)
  assert.throws(() => validateFrameBytes(new Uint8Array(), 'jpeg'), /empty/)
  assert.throws(() => validateFrameBytes(png, 'gif'), /Choose PNG or JPEG/)
})
