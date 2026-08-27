'use strict'

const assert = require('node:assert/strict')
const path = require('node:path')
const test = require('node:test')
const { SUPPORTED_EXTENSIONS, isSupportedVideoPath, supportedPaths } = require('../electron/routing.cjs')

test('the routing contract contains only the unified video extensions', () => {
  assert.deepEqual(SUPPORTED_EXTENSIONS, ['.mp4', '.m4v', '.webm', '.ogv', '.mov', '.mkv'])
})

test('extension checks are case insensitive and reject ambiguous audio containers', () => {
  assert.equal(isSupportedVideoPath('C:\\Media\\Clip.MP4'), true)
  assert.equal(isSupportedVideoPath('/media/clip.ogv'), true)
  assert.equal(isSupportedVideoPath('/media/audio.ogg'), false)
  assert.equal(isSupportedVideoPath('/media/renamed.txt'), false)
  assert.equal(isSupportedVideoPath(''), false)
})

test('command-line routing preserves order, resolves relative paths, and de-duplicates on Windows', () => {
  const cwd = path.resolve('C:\\Videos')
  const result = supportedPaths([
    '--flag',
    'one.mp4',
    'notes.txt',
    'ONE.MP4',
    'nested/two.webm',
  ], cwd, 'win32')
  assert.deepEqual(result, [
    path.resolve(cwd, 'one.mp4'),
    path.resolve(cwd, 'nested/two.webm'),
  ])
})

test('command-line routing is deterministic for empty and invalid inputs', () => {
  assert.deepEqual(supportedPaths(null), [])
  assert.deepEqual(supportedPaths(['--inspect=9222', '.', 'movie.avi']), [])
})
