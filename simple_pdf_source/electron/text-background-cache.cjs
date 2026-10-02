'use strict'

const crypto = require('node:crypto')

/**
 * The editor asks for a page with some native text removed every time a text
 * item is clicked or an edited page scrolls back into view. The whole PDF
 * arrives each time; parsing it with pdf-lib for every request took seconds
 * and hundreds of MB on large files. Results are cached per document revision
 * (a hash of its bytes): the parsed document (one, released when idle), each
 * extracted single page, and each finished preview.
 */
function asUint8Array(value) {
  if (value instanceof Uint8Array) return value
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return new Uint8Array(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Uint8Array.from(value.data)
  return new Uint8Array(value || [])
}

function revisionKey(bytes) {
  const view = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return `${view.length}:${crypto.createHash('sha1').update(view).digest('hex')}`
}

function editsKey(edits) {
  return JSON.stringify((Array.isArray(edits) ? edits : []).map((edit) => [
    edit?.type, Boolean(edit?.cover), edit?.originalText ?? '', Number(edit?.angle) || 0,
    ['x', 'y', 'width', 'height'].map((key) => Number(edit?.originalRect?.[key])),
  ]))
}

class LruCache {
  constructor(maxEntries, maxBytes = Number.POSITIVE_INFINITY) {
    this.maxEntries = maxEntries
    this.maxBytes = maxBytes
    this.entries = new Map()
    this.bytes = 0
  }

  get(key) {
    if (!this.entries.has(key)) return undefined
    const entry = this.entries.get(key)
    this.entries.delete(key)
    this.entries.set(key, entry)
    return entry.value
  }

  set(key, value, size = 0) {
    this.delete(key)
    this.entries.set(key, { value, size })
    this.bytes += size
    while (this.entries.size > this.maxEntries || (this.bytes > this.maxBytes && this.entries.size > 1)) {
      this.delete(this.entries.keys().next().value)
    }
  }

  resize(key, size) {
    const entry = this.entries.get(key)
    if (!entry) return
    this.delete(key)
    this.set(key, entry.value, size)
  }

  delete(key) {
    const entry = this.entries.get(key)
    if (!entry) return
    this.bytes -= entry.size
    this.entries.delete(key)
  }

  clear() {
    this.entries.clear()
    this.bytes = 0
  }
}

/**
 * @param {{
 *   parse: (bytes: Uint8Array) => Promise<any>,
 *   extractPage: (doc: any, pageIndex: number) => Promise<Uint8Array>,
 *   removeText: (pageBytes: Uint8Array, edits: object[]) => Promise<Uint8Array>,
 *   idleMs?: number, maxPages?: number, maxResults?: number, maxResultBytes?: number,
 * }} options
 */
function createTextBackgroundCache({
  parse,
  extractPage,
  removeText,
  idleMs = 60_000,
  maxPages = 24,
  maxResults = 48,
  maxResultBytes = 128 * 1024 * 1024,
}) {
  let parsed = null
  const pages = new LruCache(maxPages, maxResultBytes)
  const results = new LruCache(maxResults, maxResultBytes)

  function parsedDocument(key, bytes) {
    if (parsed?.key !== key) {
      if (parsed?.timer) clearTimeout(parsed.timer)
      parsed = { key, promise: parse(bytes), timer: null }
      parsed.promise.catch(() => { if (parsed?.key === key) parsed = null })
    }
    if (parsed.timer) clearTimeout(parsed.timer)
    const current = parsed
    current.timer = setTimeout(() => { if (parsed === current) parsed = null }, idleMs)
    current.timer.unref?.()
    return current.promise
  }

  function memo(cache, key, compute) {
    const cached = cache.get(key)
    if (cached) return cached
    const promise = compute()
    cache.set(key, promise)
    promise.then((value) => cache.resize(key, value?.byteLength || 0), () => cache.delete(key))
    return promise
  }

  async function render(data, pageIndex, edits) {
    const bytes = asUint8Array(data)
    const key = revisionKey(bytes)
    const index = Number(pageIndex)
    if (!Number.isInteger(index) || index < 0) throw new RangeError('The page to preview does not exist.')
    return memo(results, `${key}#${index}#${editsKey(edits)}`, async () => {
      const page = await memo(pages, `${key}#${index}`, async () => extractPage(await parsedDocument(key, bytes), index))
      return removeText(page, (Array.isArray(edits) ? edits : []).map((edit) => ({ ...edit, pageIndex: 0 })))
    })
  }

  function clear() {
    if (parsed?.timer) clearTimeout(parsed.timer)
    parsed = null
    pages.clear()
    results.clear()
  }

  return { render, clear }
}

module.exports = { createTextBackgroundCache, revisionKey }
