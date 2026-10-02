'use strict'
// Minimal PDF text extraction for tests (no PDF.js): per page, the strings shown by
// Tj/TJ/'/" operators, decoded through each font's ToUnicode CMap (pdfkit writes
// embedded fonts as 2-byte codes with a ToUnicode map), each with its position in
// PDF points (origin at the bottom left of the page, the graphics state's cm
// transforms applied), so tests can tell header and footer text from the body.
const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRawStream, PDFRef, decodePDFRawStream } = require('pdf-lib')

function streamBytes(stream) {
  if (stream instanceof PDFRawStream) return Buffer.from(decodePDFRawStream(stream).decode())
  return Buffer.from(stream.getContents())
}

function hexToBytes(hex) {
  const clean = hex.replace(/[^0-9a-f]/gi, '')
  return Buffer.from(clean.length % 2 ? `${clean}0` : clean, 'hex')
}

function utf16(bytes) {
  let text = ''
  for (let index = 0; index + 1 < bytes.length; index += 2) text += String.fromCharCode(bytes.readUInt16BE(index))
  return text
}

/** Parses bfchar/bfrange entries of a ToUnicode CMap into code → text, plus the code width in bytes. */
function parseCMap(source) {
  const map = new Map()
  let width = 1
  const codespace = /begincodespacerange\s*<([0-9a-f]+)>/i.exec(source)
  if (codespace) width = Math.max(1, codespace[1].length / 2)
  for (const block of source.matchAll(/beginbfchar([\s\S]*?)endbfchar/gi)) {
    for (const entry of block[1].matchAll(/<([0-9a-f]+)>\s*<([0-9a-f\s]*)>/gi)) map.set(parseInt(entry[1], 16), utf16(hexToBytes(entry[2])))
  }
  for (const block of source.matchAll(/beginbfrange([\s\S]*?)endbfrange/gi)) {
    for (const entry of block[1].matchAll(/<([0-9a-f]+)>\s*<([0-9a-f]+)>\s*(\[[^\]]*\]|<[0-9a-f\s]*>)/gi)) {
      const start = parseInt(entry[1], 16)
      const end = parseInt(entry[2], 16)
      if (entry[3].startsWith('[')) {
        // A value can hold several UTF-16 units (<0074 0069> is the "ti" ligature).
        const values = [...entry[3].matchAll(/<([0-9a-f\s]*)>/gi)].map((value) => utf16(hexToBytes(value[1])))
        values.forEach((value, offset) => map.set(start + offset, value))
      } else {
        const base = hexToBytes(entry[3].slice(1, -1))
        for (let code = start; code <= end && code - start < 65536; code += 1) {
          const bytes = Buffer.from(base)
          bytes.writeUInt16BE((bytes.readUInt16BE(bytes.length - 2) + code - start) & 0xffff, bytes.length - 2)
          map.set(code, utf16(bytes))
        }
      }
    }
  }
  return { map, width }
}

function decodeString(bytes, font) {
  if (!font) return bytes.toString('latin1')
  let text = ''
  for (let index = 0; index + font.width <= bytes.length; index += font.width) {
    const code = font.width === 2 ? bytes.readUInt16BE(index) : bytes[index]
    text += font.map.get(code) ?? ''
  }
  return text
}

function literalBytes(literal) {
  const out = []
  for (let index = 0; index < literal.length; index += 1) {
    const char = literal[index]
    if (char !== '\\') { out.push(char.charCodeAt(0) & 0xff); continue }
    const next = literal[++index]
    const escapes = { n: 10, r: 13, t: 9, b: 8, f: 12, '(': 40, ')': 41, '\\': 92 }
    if (next in escapes) out.push(escapes[next])
    else if (/[0-7]/.test(next ?? '')) {
      let octal = next
      while (octal.length < 3 && /[0-7]/.test(literal[index + 1] ?? '')) octal += literal[++index]
      out.push(parseInt(octal, 8) & 0xff)
    }
  }
  return Buffer.from(out)
}

const DELIMITER = /[\s/<>[\]()%{}]/

/** Content-stream tokens: numbers, names, strings (bytes), array brackets and operators. */
function* tokenize(source) {
  let index = 0
  while (index < source.length) {
    const char = source[index]
    if (/\s/.test(char)) { index += 1; continue }
    if (char === '%') { while (index < source.length && source[index] !== '\n') index += 1; continue }
    if (char === '(') {
      let depth = 1
      let at = index + 1
      let text = ''
      while (at < source.length) {
        const current = source[at]
        if (current === '\\') { text += current + (source[at + 1] ?? ''); at += 2; continue }
        if (current === '(') depth += 1
        else if (current === ')' && --depth === 0) break
        text += current
        at += 1
      }
      yield { type: 'string', bytes: literalBytes(text) }
      index = at + 1
      continue
    }
    if (source.startsWith('<<', index) || source.startsWith('>>', index)) { yield { type: 'op', value: source.slice(index, index + 2) }; index += 2; continue }
    if (char === '<') {
      const end = source.indexOf('>', index)
      yield { type: 'string', bytes: hexToBytes(source.slice(index + 1, end)) }
      index = end + 1
      continue
    }
    if (char === '[' || char === ']') { yield { type: char }; index += 1; continue }
    let end = index + 1
    while (end < source.length && !DELIMITER.test(source[end])) end += 1
    const word = source.slice(index, end)
    index = end
    if (word.startsWith('/')) yield { type: 'name', value: word.slice(1) }
    else if (/^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(word)) yield { type: 'number', value: Number(word) }
    else yield { type: 'op', value: word }
  }
}

const IDENTITY = [1, 0, 0, 1, 0, 0]
/** m × n for PDF matrices [a b c d e f]. */
function multiply(m, n) {
  return [
    m[0] * n[0] + m[1] * n[2],
    m[0] * n[1] + m[1] * n[3],
    m[2] * n[0] + m[3] * n[2],
    m[2] * n[1] + m[3] * n[3],
    m[4] * n[0] + m[5] * n[2] + n[4],
    m[4] * n[1] + m[5] * n[3] + n[5],
  ]
}

function resolve(context, value) {
  return value instanceof PDFRef ? context.lookup(value) : value
}

/** Text runs of every page: { height, text, runs: [{ text, x, y }] } with x/y in points from the bottom left. */
async function pdfPageTexts(bytes) {
  const pdf = await PDFDocument.load(bytes, { updateMetadata: false })
  const pages = []
  for (const page of pdf.getPages()) {
    const context = pdf.context
    const fonts = new Map()
    const resources = resolve(context, page.node.Resources())
    const fontDict = resources ? resolve(context, resources.get(PDFName.of('Font'))) : null
    if (fontDict instanceof PDFDict) {
      for (const [name, ref] of fontDict.entries()) {
        const font = resolve(context, ref)
        const toUnicode = font instanceof PDFDict ? resolve(context, font.get(PDFName.of('ToUnicode'))) : null
        if (toUnicode) fonts.set(name.asString().replace(/^\//, ''), parseCMap(streamBytes(toUnicode).toString('latin1')))
      }
    }
    const contents = resolve(context, page.node.get(PDFName.of('Contents')))
    const streams = contents instanceof PDFArray ? contents.asArray().map((ref) => resolve(context, ref)) : [contents]
    const program = streams.filter(Boolean).map((stream) => streamBytes(stream).toString('latin1')).join('\n')
    const runs = []
    const stack = []
    let ctm = IDENTITY
    let tm = IDENTITY
    let font = null
    let operands = []
    let array = null
    const show = (text) => {
      const point = multiply(tm, ctm)
      if (text) runs.push({ text, x: point[4], y: point[5] })
    }
    for (const token of tokenize(program)) {
      if (token.type === '[') { array = []; continue }
      if (token.type === ']') { operands.push({ type: 'array', items: array ?? [] }); array = null; continue }
      if (array) { array.push(token); continue }
      if (token.type !== 'op') { operands.push(token); continue }
      const numbers = operands.filter((operand) => operand.type === 'number').map((operand) => operand.value)
      switch (token.value) {
        case 'q': stack.push(ctm); break
        case 'Q': ctm = stack.pop() ?? IDENTITY; break
        case 'cm': if (numbers.length >= 6) ctm = multiply(numbers.slice(-6), ctm); break
        case 'BT': tm = IDENTITY; break
        case 'Tm': if (numbers.length >= 6) tm = numbers.slice(-6); break
        case 'Td': case 'TD': if (numbers.length >= 2) tm = multiply([1, 0, 0, 1, numbers.at(-2), numbers.at(-1)], tm); break
        case 'Tf': { const name = operands.find((operand) => operand.type === 'name'); font = name ? fonts.get(name.value) ?? null : null; break }
        case 'Tj': case "'": case '"': { const string = operands.filter((operand) => operand.type === 'string').at(-1); if (string) show(decodeString(string.bytes, font)); break }
        case 'TJ': {
          const list = operands.find((operand) => operand.type === 'array')
          if (list) show(list.items.filter((item) => item.type === 'string').map((item) => decodeString(item.bytes, font)).join(''))
          break
        }
        default: break
      }
      operands = []
    }
    const { height } = page.getSize()
    pages.push({ height, runs, text: runs.map((run) => run.text).join(' ') })
  }
  return pages
}

/** The text of a page's top or bottom band (within `points` of the edge), in reading order. */
function bandText(page, where, points = 72) {
  const inBand = page.runs.filter((run) => (where === 'bottom' ? run.y < points : run.y > page.height - points))
  return inBand.sort((a, b) => b.y - a.y || a.x - b.x).map((run) => run.text).join(' ')
}

module.exports = { pdfPageTexts, parseCMap, bandText }
