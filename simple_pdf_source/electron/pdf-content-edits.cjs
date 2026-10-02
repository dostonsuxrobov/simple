const {
  decodePDFRawStream,
  PDFArray,
  PDFDict,
  PDFName,
  PDFRawStream,
} = require('pdf-lib')
const { codedError } = require('./pdf-problems.cjs')

const IDENTITY = [1, 0, 0, 1, 0, 0]
const WHITESPACE = new Set([' ', '\t', '\r', '\n', '\f', '\0'])
const DELIMITERS = new Set(['(', ')', '<', '>', '[', ']', '{', '}', '/', '%'])
const SPACE = 0x20

function multiplyMatrix(left, right) {
  return [
    left[0] * right[0] + left[2] * right[1],
    left[1] * right[0] + left[3] * right[1],
    left[0] * right[2] + left[2] * right[3],
    left[1] * right[2] + left[3] * right[3],
    left[0] * right[4] + left[2] * right[5] + left[4],
    left[1] * right[4] + left[3] * right[5] + left[5],
  ]
}

function transformedUnitRect(matrix) {
  const points = [[0, 0], [1, 0], [1, 1], [0, 1]].map(([x, y]) => ([
    matrix[0] * x + matrix[2] * y + matrix[4],
    matrix[1] * x + matrix[3] * y + matrix[5],
  ]))
  const xs = points.map(([x]) => x)
  const ys = points.map(([, y]) => y)
  const x = Math.min(...xs)
  const y = Math.min(...ys)
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y }
}

function rectsMatch(left, right, tolerance) {
  return Math.abs(left.x - right.x) <= tolerance
    && Math.abs(left.y - right.y) <= tolerance
    && Math.abs(left.width - right.width) <= tolerance
    && Math.abs(left.height - right.height) <= tolerance
}

function readImageResourceNames(page) {
  const resources = page.node.Resources()
  const xobjects = resources?.lookupMaybe(PDFName.of('XObject'), PDFDict)
  const names = new Set()
  if (!xobjects) return names
  for (const [key] of xobjects.entries()) {
    const value = xobjects.lookup(key)
    const dictionary = value instanceof PDFRawStream ? value.dict : value instanceof PDFDict ? value : null
    if (dictionary?.get(PDFName.of('Subtype'))?.toString() === '/Image') {
      names.add(key.toString().slice(1))
    }
  }
  return names
}

function contentStreamBytes(stream) {
  if (stream instanceof PDFRawStream) return decodePDFRawStream(stream).decode()
  // Streams created by pdf-lib in this session (drawn overlays) are not raw.
  if (typeof stream?.getUnencodedContents === 'function') return stream.getUnencodedContents()
  return null
}

/**
 * The page's content as the exact decoded bytes, streams joined by a newline.
 * Returns null when a stream cannot be decoded, so callers leave the page alone
 * instead of rewriting it without that stream.
 */
function readPageContentBytes(page) {
  const contents = page.node.get(PDFName.of('Contents'))
  if (!contents) return null
  const resolved = page.doc.context.lookup(contents)
  const streams = resolved instanceof PDFArray
    ? Array.from({ length: resolved.size() }, (_, index) => page.doc.context.lookup(resolved.get(index)))
    : [resolved]
  const parts = []
  for (const stream of streams) {
    if (!stream) continue
    let bytes
    try { bytes = contentStreamBytes(stream) } catch { return null }
    if (!bytes) return null
    parts.push(bytes)
  }
  if (!parts.length) return null
  const length = parts.reduce((total, part) => total + part.length, 0) + parts.length - 1
  const joined = new Uint8Array(length)
  let offset = 0
  parts.forEach((part, index) => {
    if (index) joined[offset++] = 0x0a
    joined.set(part, offset)
    offset += part.length
  })
  return joined
}

/**
 * A one-character-per-byte view of content bytes. Node's 'latin1' is true
 * ISO-8859-1; WHATWG TextDecoder('latin1') is windows-1252 and would remap
 * 27 byte values between 0x80 and 0x9F, so string offsets equal byte offsets
 * only with this conversion.
 */
function binaryString(bytes) {
  return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString('latin1')
}

function readPageContent(page) {
  const bytes = readPageContentBytes(page)
  return bytes ? binaryString(bytes) : ''
}

function skipLiteralString(content, index) {
  let depth = 0
  while (index < content.length) {
    const character = content[index]
    if (character === '\\') {
      index += 2
      continue
    }
    if (character === '(') depth += 1
    if (character === ')' && --depth === 0) return index + 1
    index += 1
  }
  return index
}

function skipInlineImage(content, index) {
  const dataStart = content.indexOf('ID', index)
  if (dataStart < 0) return content.length
  for (let cursor = dataStart + 2; cursor + 1 < content.length; cursor += 1) {
    if (content[cursor] === 'E' && content[cursor + 1] === 'I'
      && WHITESPACE.has(content[cursor - 1])
      && (cursor + 2 >= content.length || WHITESPACE.has(content[cursor + 2]))) return cursor + 2
  }
  return content.length
}

function isNumberToken(value) {
  return /^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(value)
}

function walkImageDraws(content, imageNames) {
  const hits = []
  const operands = []
  const stack = []
  let matrix = [...IDENTITY]
  let index = 0

  while (index < content.length) {
    const character = content[index]
    if (WHITESPACE.has(character)) {
      index += 1
      continue
    }
    if (character === '%') {
      while (index < content.length && content[index] !== '\n' && content[index] !== '\r') index += 1
      continue
    }
    if (character === '(') {
      const start = index
      index = skipLiteralString(content, index)
      operands.push({ value: '()', start, end: index })
      continue
    }
    if (character === '<' && content[index + 1] !== '<') {
      const start = index++
      while (index < content.length && content[index] !== '>') index += 1
      index += 1
      operands.push({ value: '<>', start, end: index })
      continue
    }
    if (character === '/') {
      const start = index++
      while (index < content.length && !WHITESPACE.has(content[index]) && !DELIMITERS.has(content[index])) index += 1
      operands.push({ value: content.slice(start, index), start, end: index })
      continue
    }
    if (DELIMITERS.has(character)) {
      operands.push({ value: character, start: index, end: index + 1 })
      index += 1
      continue
    }

    const start = index
    while (index < content.length && !WHITESPACE.has(content[index]) && !DELIMITERS.has(content[index])) index += 1
    const value = content.slice(start, index)
    if (isNumberToken(value)) {
      operands.push({ value, start, end: index })
      continue
    }

    if (value === 'q') stack.push([...matrix])
    else if (value === 'Q') matrix = stack.pop() || [...IDENTITY]
    else if (value === 'cm' && operands.length >= 6) {
      const next = operands.slice(-6).map((operand) => Number(operand.value))
      if (next.every(Number.isFinite)) matrix = multiplyMatrix(matrix, next)
    } else if (value === 'Do') {
      const nameOperand = operands.at(-1)
      const name = nameOperand?.value.startsWith('/') ? nameOperand.value.slice(1) : ''
      if (nameOperand && imageNames.has(name)) {
        hits.push({
          name,
          rect: transformedUnitRect(matrix),
          start: nameOperand.start,
          end: index,
        })
      }
    } else if (value === 'BI') {
      index = skipInlineImage(content, index)
    }
    operands.length = 0
  }
  return hits
}

function listPageImageDraws(page) {
  return walkImageDraws(readPageContent(page), readImageResourceNames(page))
    .map(({ name, rect }) => ({ name, rect }))
}

function sameBytes(left, right) {
  if (left.length !== right.length) return false
  return Buffer.from(left.buffer, left.byteOffset, left.byteLength)
    .equals(Buffer.from(right.buffer, right.byteOffset, right.byteLength))
}

/**
 * Give the page its own /Resources and /XObject dictionaries before deleting
 * an entry, so pages and forms sharing the original dictionaries keep it.
 */
function privateXObjectDictionary(page) {
  const { context } = page.doc
  const resources = page.node.Resources()
  const xobjects = resources?.lookupMaybe(PDFName.of('XObject'), PDFDict)
  if (!resources || !xobjects) return null
  const ownResources = resources.clone(context)
  const ownXObjects = xobjects.clone(context)
  ownResources.set(PDFName.of('XObject'), ownXObjects)
  page.node.set(PDFName.of('Resources'), ownResources)
  return ownXObjects
}

/**
 * Forget image resources whose last drawing was removed, so the pixels are not
 * kept in the saved file. A form XObject without its own /Resources may draw
 * through the page's resources, so the entries stay in that (rare) case.
 */
function dropUnusedImageResources(page, removedNames, remainingContent) {
  if (!removedNames.size) return 0
  const resources = page.node.Resources()
  const xobjects = resources?.lookupMaybe(PDFName.of('XObject'), PDFDict)
  if (!xobjects) return 0
  for (const [key] of xobjects.entries()) {
    const value = xobjects.lookup(key)
    const dictionary = value instanceof PDFRawStream ? value.dict : value instanceof PDFDict ? value : null
    if (dictionary?.get(PDFName.of('Subtype'))?.toString() === '/Form' && !dictionary.has(PDFName.of('Resources'))) return 0
  }
  const stillDrawn = new Set(walkImageDraws(remainingContent, removedNames).map((hit) => hit.name))
  const unused = [...removedNames].filter((name) => !stillDrawn.has(name))
  if (!unused.length) return 0
  const ownXObjects = privateXObjectDictionary(page)
  if (!ownXObjects) return 0
  for (const name of unused) ownXObjects.delete(PDFName.of(name))
  return unused.length
}

/**
 * Remove only image-XObject invocations whose painted rectangle matches a
 * selected native image. Shared image resources remain available to every
 * other invocation on the page. All other content bytes are kept exactly:
 * text in single-byte encodings, CID strings and inline image data included.
 */
function removePageImageDraws(page, targets, tolerance = 2) {
  if (!Array.isArray(targets) || !targets.length) return 0
  const bytes = readPageContentBytes(page)
  if (!bytes) return 0
  const content = binaryString(bytes)
  const imageNames = readImageResourceNames(page)
  const hits = walkImageDraws(content, imageNames)
    .filter((hit) => targets.some((target) => rectsMatch(hit.rect, target, tolerance)))
  if (!hits.length) return 0

  const edited = Uint8Array.from(bytes)
  for (const hit of hits) edited.fill(SPACE, hit.start, hit.end)
  // Bytes outside the blanked operators must survive compression unchanged.
  for (let index = 0, hit = 0; index < bytes.length; index += 1) {
    while (hit < hits.length && hits[hit].end <= index) hit += 1
    const blanked = hit < hits.length && index >= hits[hit].start && index < hits[hit].end
    if (!blanked && edited[index] !== bytes[index]) {
      throw codedError('CONTENT_REWRITE_MISMATCH', 'An image could not be removed without changing other page content.')
    }
  }
  const stream = page.doc.context.flateStream(edited)
  if (!sameBytes(decodePDFRawStream(stream).decode(), edited)) {
    throw codedError('CONTENT_REWRITE_MISMATCH', 'An image could not be removed without changing other page content.')
  }
  page.node.set(PDFName.of('Contents'), page.doc.context.register(stream))
  dropUnusedImageResources(page, new Set(hits.map((hit) => hit.name)), binaryString(edited))
  return hits.length
}

module.exports = {
  binaryString,
  listPageImageDraws,
  readPageContentBytes,
  removePageImageDraws,
  transformedUnitRect,
}
