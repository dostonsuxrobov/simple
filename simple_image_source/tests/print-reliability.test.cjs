'use strict'
const test = require('node:test')
const assert = require('node:assert/strict')
const { EventEmitter } = require('node:events')
const { preparePrintableImage } = require('../electron/print-image.cjs')
const { submitPrintJob } = require('../electron/default-printer.cjs')

const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0, 11, 8, 0, 40, 0, 60, 1, 1, 0x11, 0, 0xff, 0xd9])
test('untouched JPEG print retains exact compressed bytes and permits browser-oriented dimensions', () => {
  const result = preparePrintableImage({ data: jpeg, width: 60, height: 40 })
  assert.equal(result.extension, '.jpg')
  assert.deepEqual(result.bytes, jpeg)
  assert.deepEqual(preparePrintableImage({ data: jpeg, width: 40, height: 60 }).dimensions, { width: 40, height: 60 })
  assert.throws(() => preparePrintableImage({ data: jpeg, width: 20, height: 60 }), /dimensions/)
  assert.throws(() => preparePrintableImage({ data: Buffer.from('wrong'), width: 1, height: 1 }), /content/)
})

test('a print callback that never arrives stops waiting without automatically retrying', async () => {
  const renderer = new EventEmitter()
  let calls = 0
  let callback
  renderer.print = (_options, cb) => { calls += 1; callback = cb }
  await assert.rejects(submitPrintJob(renderer, {}, { timeoutMs: 5 }), /Check the print queue.*duplicate/)
  callback(true)
  assert.equal(calls, 1)
  assert.equal(renderer.listenerCount('destroyed'), 0)
  assert.equal(renderer.listenerCount('render-process-gone'), 0)
})

test('print renderer crashes and driver exceptions cleanly release the operation', async () => {
  const renderer = new EventEmitter()
  renderer.print = () => queueMicrotask(() => renderer.emit('render-process-gone'))
  await assert.rejects(submitPrintJob(renderer, {}), /renderer stopped/)
  renderer.print = () => { throw new Error('driver failed') }
  await assert.rejects(submitPrintJob(renderer, {}), /driver failed/)
  await assert.rejects(submitPrintJob(null, {}), /renderer stopped/)
  assert.equal(renderer.listenerCount('destroyed'), 0)
})

test('successful print submission settles once and removes all lifecycle listeners', async () => {
  const renderer = new EventEmitter()
  renderer.print = (_options, callback) => { callback(true); callback(false, 'late') }
  assert.equal(await submitPrintJob(renderer, {}), true)
  assert.equal(renderer.listenerCount('destroyed'), 0)
  assert.equal(renderer.listenerCount('render-process-gone'), 0)
})
