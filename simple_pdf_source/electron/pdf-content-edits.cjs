const {
  decodePDFRawStream,
  PDFArray,
  PDFDict,
  PDFName,
  PDFNumber,
  PDFRawStream,
} = require('pdf-lib')

const LATIN1 = new TextDecoder('latin1')
const IDENTITY = [1, 0, 0, 1, 0, 0]
const WHITESPACE = new Set([' ', '\t', '\r', '\n', '\f', '\0'])
const DELIMITERS = new Set(['(', ')', '<', '>', '[', ']', '{', '}', '/', '%'])

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

function readPageContent(page) {
  const contents = page.node.get(PDFName.of('Contents'))
  if (!contents) return ''
  const resolved = page.doc.context.lookup(contents)
  const streams = resolved instanceof PDFArray
    ? Array.from({ length: resolved.size() }, (_, index) => page.doc.context.lookup(resolved.get(index)))
    : [resolved]
  return streams
    .filter((stream) => stream instanceof PDFRawStream)
    .map((stream) => LATIN1.decode(decodePDFRawStream(stream).decode()))
    .join('\n')
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

/**
 * Remove only image-XObject invocations whose painted rectangle matches a
 * selected native image. Shared image resources remain available to every
 * other invocation on the page.
 */
function removePageImageDraws(page, targets, tolerance = 2) {
  if (!Array.isArray(targets) || !targets.length) return 0
  const content = readPageContent(page)
  if (!content) return 0
  const hits = walkImageDraws(content, readImageResourceNames(page))
    .filter((hit) => targets.some((target) => rectsMatch(hit.rect, target, tolerance)))
  if (!hits.length) return 0

  const characters = content.split('')
  for (const hit of hits) {
    for (let index = hit.start; index < hit.end; index += 1) characters[index] = ' '
  }
  const stream = page.doc.context.flateStream(characters.join(''))
  page.node.set(PDFName.of('Contents'), page.doc.context.register(stream))
  return hits.length
}

module.exports = {
  listPageImageDraws,
  removePageImageDraws,
  transformedUnitRect,
}
