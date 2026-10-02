'use strict'
// WP1 imaging worker client: inline routing, transfer of owning buffers and compact copies of
// non-owning views, progress, abort, striping across the pool, start-failure fallback and crash
// recovery. A fake Worker runs the real worker runtime (worker-ops/index.ts) and moves messages with
// structuredClone + transfer, so detaching behaves like a real worker.
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

const { createImagingClient, planStripes, prepareTransfer } = load('shared/workerClient.ts')
const { HANDLERS, HANDLER_TABLES, attachImagingWorker } = load('shared/worker-ops/index.ts')

const SPEC = { type: 'find-edges' }
const tick = () => new Promise((resolve) => setImmediate(resolve))
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

async function waitFor(predicate, label, timeout = 2000) {
  const end = Date.now() + timeout
  while (!predicate()) {
    if (Date.now() > end) throw new Error(`Timed out: ${label}`)
    await sleep(2)
  }
}

function buffer(width, height, seed = 1) {
  const data = new Uint8ClampedArray(width * height * 4)
  let state = seed >>> 0
  for (let index = 0; index < data.length; index += 1) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    data[index] = state >>> 24
  }
  return { width, height, data }
}

function copy(src) {
  return { width: src.width, height: src.height, data: new Uint8ClampedArray(src.data) }
}

function invert({ src }) {
  const data = new Uint8ClampedArray(src.data.length)
  for (let index = 0; index < data.length; index += 1) data[index] = index % 4 === 3 ? src.data[index] : 255 - src.data[index]
  return { width: src.width, height: src.height, data }
}

/** Position-independent 3x3 mean with clamped edges: exact under striping with a 1-row margin. */
function boxBlur3({ src }) {
  const { width, height } = src
  const data = new Uint8ClampedArray(src.data.length)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      for (let channel = 0; channel < 4; channel += 1) {
        let sum = 0
        for (let dy = -1; dy <= 1; dy += 1) {
          for (let dx = -1; dx <= 1; dx += 1) {
            const sx = Math.min(width - 1, Math.max(0, x + dx))
            const sy = Math.min(height - 1, Math.max(0, y + dy))
            sum += src.data[(sy * width + sx) * 4 + channel]
          }
        }
        data[(y * width + x) * 4 + channel] = Math.round(sum / 9)
      }
    }
  }
  return { width, height, data }
}

/** Point operation that only touches pixels where the mask is set. */
function maskedInvert({ src, mask }) {
  const out = invert({ src })
  for (let pixel = 0; pixel < src.width * src.height; pixel += 1) {
    if (mask && mask.data[pixel] === 0) out.data.set(src.data.subarray(pixel * 4, pixel * 4 + 4), pixel * 4)
  }
  return out
}

class FakeWorker {
  constructor(handlers, options = {}) {
    this.listeners = new Map()
    this.posted = []
    this.terminated = false
    this.options = options
    const inbound = []
    this.scope = {
      postMessage: (message, transfer = []) => {
        if (this.terminated) return
        const cloned = structuredClone(message, { transfer })
        setImmediate(() => { if (!this.terminated) this.dispatch('message', { data: cloned }) })
      },
      addEventListener: (type, listener) => { if (type === 'message') inbound.push(listener) },
    }
    this.deliver = (message) => { for (const listener of inbound) listener({ data: message }) }
    setImmediate(() => {
      if (this.terminated) return
      if (options.failStart) this.dispatch('error', { message: 'Failed to load module script', preventDefault() {} })
      else attachImagingWorker(this.scope, handlers)
    })
  }

  addEventListener(type, listener) {
    if (!this.listeners.has(type)) this.listeners.set(type, [])
    this.listeners.get(type).push(listener)
  }

  removeEventListener(type, listener) {
    this.listeners.set(type, (this.listeners.get(type) || []).filter((entry) => entry !== listener))
  }

  dispatch(type, event) {
    for (const listener of this.listeners.get(type) || []) listener(event)
  }

  postMessage(message, transfer = []) {
    if (this.terminated) throw new Error('The worker was terminated.')
    this.posted.push({ message, transferSizes: transfer.map((entry) => entry.byteLength) })
    const cloned = structuredClone(message, { transfer })
    if (this.options.crashOnRun && cloned.type === 'run') {
      setImmediate(() => this.dispatch('error', { message: 'Out of memory', preventDefault() {} }))
      return
    }
    setImmediate(() => { if (!this.terminated) this.deliver(cloned) })
  }

  terminate() {
    this.terminated = true
  }
}

test('worker-op tables are disjoint and merged into one routing table', () => {
  const color = Object.keys(HANDLER_TABLES.color)
  const mask = Object.keys(HANDLER_TABLES.mask)
  assert.deepEqual(color.filter((op) => mask.includes(op)), [], 'an op must have exactly one owner')
  assert.deepEqual(Object.keys(HANDLERS).sort(), [...color, ...mask].sort())
  assert.ok(Object.isFrozen(HANDLERS))
})

test('the worker entry is inert when imported outside a worker', () => {
  assert.doesNotThrow(() => load('shared/imaging.worker.ts'))
})

test('inline mode routes ops to handlers and reports unknown ops clearly', async () => {
  const client = createImagingClient({ inline: { filter: invert } })
  assert.equal(client.mode, 'inline')
  const src = { width: 2, height: 1, data: Uint8ClampedArray.from([10, 20, 30, 255, 0, 0, 0, 0]) }
  const output = await client.run('filter', { src, spec: SPEC, mask: null })
  assert.deepEqual([...output.data], [245, 235, 225, 255, 255, 255, 255, 0])
  await assert.rejects(client.run('heal', { src, hole: { width: 2, height: 1, data: new Uint8Array(2) }, options: {} }),
    (error) => error.name === 'NotSupportedError' && /"heal" is not available/.test(error.message))
  client.dispose()
  assert.equal(client.mode, 'disposed')
  await assert.rejects(client.run('filter', { src, spec: SPEC, mask: null }), { name: 'AbortError' })
})

test('without Worker the default client runs the merged worker-op tables inline', async () => {
  assert.equal(typeof globalThis.Worker, 'undefined')
  const client = createImagingClient()
  assert.equal(client.mode, 'inline')
  const ops = ['adjust', 'filter', 'look', 'resample', 'rotate', 'warp', 'histogram', 'autoEnhance', 'flood', 'feather', 'expand', 'contract', 'heal']
  const missing = ops.find((op) => typeof HANDLERS[op] !== 'function')
  if (missing) {
    await assert.rejects(client.run(missing, {}), { name: 'NotSupportedError' })
  }
  client.dispose()
})

test('prepareTransfer transfers owning buffers, compacts non-owning views and normalises pixel objects', () => {
  const owned = new Uint8ClampedArray(8)
  const backing = new Uint8Array(32).map((_, index) => index)
  const view = backing.subarray(4, 12)
  const buffers = new Set()
  const prepared = prepareTransfer({ a: owned, b: view, list: [owned], n: 3 }, true, buffers)
  assert.equal(prepared.a, owned)
  assert.notEqual(prepared.b, view)
  assert.equal(prepared.b.byteLength, 8)
  assert.equal(prepared.b.buffer.byteLength, 8)
  assert.deepEqual([...prepared.b], [4, 5, 6, 7, 8, 9, 10, 11])
  assert.deepEqual([...buffers].map((entry) => entry.byteLength).sort(), [8, 8], 'deduplicated: owned once, copy once')
  class PixelsLike { constructor() { this.width = 1; this.height = 1; this.data = new Uint8ClampedArray(4) } }
  const normalised = prepareTransfer({ src: new PixelsLike() }, false, new Set())
  assert.equal(Object.getPrototypeOf(normalised.src), Object.prototype)
  const keep = new Set()
  prepareTransfer({ data: new Uint8ClampedArray(4) }, false, keep)
  assert.equal(keep.size, 0, 'transfer: false transfers nothing it does not own')
})

test('owning buffers are transferred, non-owning views are copied, transfer: false keeps the caller buffer', async () => {
  const workers = []
  const client = createImagingClient({
    workers: 1,
    createWorker: () => {
      const worker = new FakeWorker({ filter: invert })
      workers.push(worker)
      return worker
    },
  })
  assert.equal(client.mode, 'workers')
  const runs = () => workers.flatMap((worker) => worker.posted).filter((entry) => entry.message.type === 'run')

  const owned = { width: 2, height: 2, data: new Uint8ClampedArray(16).fill(7) }
  const first = await client.run('filter', { src: owned, spec: SPEC, mask: null })
  assert.equal(owned.data.byteLength, 0, 'an owning view is transferred, so the caller copy detaches')
  assert.equal(first.data[0], 248)
  assert.deepEqual(runs()[0].transferSizes, [16])

  const backing = new Uint8ClampedArray(64).fill(9)
  const second = await client.run('filter', { src: { width: 2, height: 2, data: backing.subarray(16, 32) }, spec: SPEC, mask: null })
  assert.equal(backing.byteLength, 64, 'the backing buffer of a non-owning view is never detached')
  assert.deepEqual(runs()[1].transferSizes, [16], 'only a compact copy of the 16 viewed bytes travels')
  assert.equal(second.data[0], 246)

  const kept = { width: 2, height: 2, data: new Uint8ClampedArray(16).fill(1) }
  await client.run('filter', { src: kept, spec: SPEC, mask: null }, { transfer: false })
  assert.equal(kept.data.byteLength, 16)
  assert.deepEqual(runs()[2].transferSizes, [])
  assert.equal(workers.length, 1, 'jobs reuse the pooled worker')
  client.dispose()
  assert.ok(workers.every((worker) => worker.terminated))
})

test('progress is forwarded and abort reaches the running handler', async () => {
  let workerSawAbort = false
  const slow = async ({ src }, { signal, onProgress }) => {
    for (let step = 1; step <= 12; step += 1) {
      await sleep(1)
      if (signal.aborted) {
        workerSawAbort = true
        const error = new Error('stopped')
        error.name = 'AbortError'
        throw error
      }
      onProgress(step / 12)
    }
    return { width: src.width, height: src.height, data: new Uint8ClampedArray(src.data) }
  }
  const client = createImagingClient({ workers: 1, createWorker: () => new FakeWorker({ filter: slow }) })

  const fractions = []
  await client.run('filter', { src: buffer(4, 4), spec: SPEC, mask: null }, { onProgress: (fraction) => fractions.push(fraction) })
  assert.ok(fractions.length >= 2)
  assert.equal(fractions.at(-1), 1)
  assert.ok(fractions.every((fraction, index) => index === 0 || fraction >= fractions[index - 1]), 'progress never goes backwards')

  const controller = new AbortController()
  const pending = client.run('filter', { src: buffer(4, 4), spec: SPEC, mask: null }, {
    signal: controller.signal,
    onProgress: (fraction) => { if (fraction >= 0.1) controller.abort() },
  })
  await assert.rejects(pending, { name: 'AbortError' })
  await waitFor(() => workerSawAbort, 'the worker-side signal flips')
  assert.ok(await client.run('filter', { src: buffer(4, 4), spec: SPEC, mask: null }), 'the worker is reusable after an abort')

  const early = new AbortController()
  early.abort()
  await assert.rejects(client.run('filter', { src: buffer(2, 2), spec: SPEC, mask: null }, { signal: early.signal }), { name: 'AbortError' })

  const queuedController = new AbortController()
  const blocker = client.run('filter', { src: buffer(4, 4), spec: SPEC, mask: null })
  const queued = client.run('filter', { src: buffer(4, 4), spec: SPEC, mask: null }, { signal: queuedController.signal })
  queuedController.abort()
  await assert.rejects(queued, { name: 'AbortError' })
  await blocker
  client.dispose()
})

test('large jobs are striped across the pool and match the whole-image result exactly', async () => {
  const workers = []
  const client = createImagingClient({
    workers: 3,
    stripeThresholdPixels: 100,
    stripeMargins: { filter: () => 1, adjust: () => 0 },
    createWorker: () => {
      const worker = new FakeWorker({ filter: boxBlur3, adjust: maskedInvert })
      workers.push(worker)
      return worker
    },
  })
  const src = buffer(37, 61, 7)
  const expected = boxBlur3({ src: copy(src) })
  const fractions = []
  const output = await client.run('filter', { src, spec: SPEC, mask: null }, { onProgress: (fraction) => fractions.push(fraction) })
  assert.deepEqual(Buffer.from(output.data), Buffer.from(expected.data))
  assert.equal(src.data.byteLength, 37 * 61 * 4, 'stripes are copies; the caller buffer stays intact')
  assert.ok(workers.length > 1, 'several workers share the job')
  const runs = workers.flatMap((worker) => worker.posted).filter((entry) => entry.message.type === 'run')
  assert.ok(runs.length >= 3)
  assert.equal(fractions.at(-1), 1)

  const mask = { width: 37, height: 61, data: new Uint8Array(37 * 61).map((_, index) => (index % 3 ? 255 : 0)) }
  const masked = await client.run('adjust', { src: copy(src), specs: [], mask, opacity: 1 })
  assert.deepEqual(Buffer.from(masked.data), Buffer.from(maskedInvert({ src: copy(src), mask }).data), 'masks are striped with their pixels')
  client.dispose()
})

test('planStripes covers every row once and refuses splits that do not pay', () => {
  const stripes = planStripes(1000, 997, 2, 4)
  assert.ok(stripes.length >= 4)
  assert.equal(stripes[0].top, 0)
  assert.equal(stripes.at(-1).bottom, 997)
  for (let index = 1; index < stripes.length; index += 1) assert.equal(stripes[index].top, stripes[index - 1].bottom)
  assert.equal(planStripes(1000, 30, 10, 4), null, 'a margin that dwarfs the stripe height is not worth splitting')
  assert.equal(planStripes(10, 1, 0, 4), null)
})

test('a pool that cannot start falls back to inline handlers without losing input', async () => {
  const client = createImagingClient({
    workers: 2,
    inline: { filter: invert },
    createWorker: () => new FakeWorker({ filter: invert }, { failStart: true }),
  })
  const src = { width: 1, height: 1, data: Uint8ClampedArray.from([1, 2, 3, 255]) }
  const output = await client.run('filter', { src, spec: SPEC, mask: null })
  assert.deepEqual([...output.data], [254, 253, 252, 255])
  assert.equal(client.mode, 'inline')

  const throwing = createImagingClient({ inline: { filter: invert }, createWorker: () => { throw new Error('blocked by CSP') } })
  const again = await throwing.run('filter', { src: { width: 1, height: 1, data: Uint8ClampedArray.from([0, 0, 0, 9]) }, spec: SPEC, mask: null })
  assert.deepEqual([...again.data], [255, 255, 255, 9])
  assert.equal(throwing.mode, 'inline')
  client.dispose()
  throwing.dispose()
})

test('a worker that dies mid-job rejects that job and the pool recovers', async () => {
  let created = 0
  const client = createImagingClient({
    workers: 1,
    createWorker: () => {
      created += 1
      return new FakeWorker({ filter: invert }, { crashOnRun: created === 1 })
    },
  })
  await assert.rejects(client.run('filter', { src: buffer(2, 2), spec: SPEC, mask: null }), /stopped unexpectedly/)
  const output = await client.run('filter', { src: { width: 1, height: 1, data: Uint8ClampedArray.from([5, 5, 5, 5]) }, spec: SPEC, mask: null })
  assert.deepEqual([...output.data], [250, 250, 250, 5])
  assert.equal(created, 2)
  client.dispose()
})

test('dispose rejects queued and running jobs', async () => {
  const never = () => new Promise(() => {})
  const client = createImagingClient({ workers: 1, createWorker: () => new FakeWorker({ filter: never }) })
  const running = client.run('filter', { src: buffer(2, 2), spec: SPEC, mask: null })
  const queued = client.run('filter', { src: buffer(2, 2), spec: SPEC, mask: null })
  await waitFor(() => client.pending === 2, 'both jobs registered')
  await tick()
  client.dispose()
  await assert.rejects(running, { name: 'AbortError' })
  await assert.rejects(queued, { name: 'AbortError' })
  assert.equal(client.pending, 0)
})

test('the worker runtime announces readiness, transfers results and names its errors', async () => {
  const posted = []
  let deliver = null
  const scope = {
    postMessage: (message, transfer) => posted.push({ message, transfer }),
    addEventListener: (type, listener) => { if (type === 'message') deliver = listener },
  }
  const rangeError = () => { throw new RangeError('Array buffer allocation failed') }
  attachImagingWorker(scope, { filter: invert, heal: rangeError })
  assert.deepEqual(posted[0].message, { type: 'ready' })
  const src = { width: 1, height: 1, data: Uint8ClampedArray.from([0, 10, 20, 255]) }
  deliver({ data: { type: 'run', id: 1, op: 'filter', input: { src, spec: SPEC, mask: null } } })
  deliver({ data: { type: 'run', id: 2, op: 'heal', input: {} } })
  deliver({ data: { type: 'run', id: 3, op: 'warp', input: {} } })
  deliver({ data: 'not a request' })
  deliver({ data: { type: 'abort', id: 99 } })
  await waitFor(() => posted.length >= 4, 'three answers')
  const byId = new Map(posted.slice(1).map((entry) => [entry.message.id, entry]))
  assert.equal(byId.get(1).message.type, 'result')
  assert.deepEqual([...byId.get(1).message.output.data], [255, 245, 235, 255])
  assert.equal(byId.get(1).transfer[0], byId.get(1).message.output.data.buffer, 'outputs are transferred, not copied')
  assert.deepEqual(byId.get(2).message, { type: 'error', id: 2, name: 'RangeError', message: 'Array buffer allocation failed' })
  assert.equal(byId.get(3).message.name, 'NotSupportedError')
  assert.equal(posted.length, 4)
})
