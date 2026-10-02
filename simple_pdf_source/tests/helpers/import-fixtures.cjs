'use strict'

// Synthetic files for the import tests: text in several encodings, images in
// every format Simple accepts, RTF, OpenDocument text, and formats that only
// an office engine converts. Images are split left red / right blue so a
// rendered page proves the pixels arrived the right way round.
const fs = require('node:fs')
const path = require('node:path')
const JSZip = require('jszip')

const RED = [220, 20, 30]
const BLUE = [20, 40, 210]

function pixelAt(x, width) {
  return x < width / 2 ? RED : BLUE
}

function bmp(width = 40, height = 20) {
  const rowBytes = Math.ceil(width * 3 / 4) * 4
  const data = Buffer.alloc(54 + rowBytes * height)
  data.write('BM', 0, 'latin1')
  data.writeUInt32LE(data.length, 2)
  data.writeUInt32LE(54, 10)
  data.writeUInt32LE(40, 14)
  data.writeInt32LE(width, 18)
  data.writeInt32LE(height, 22)
  data.writeUInt16LE(1, 26)
  data.writeUInt16LE(24, 28)
  data.writeUInt32LE(rowBytes * height, 34)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const [r, g, b] = pixelAt(x, width)
      const at = 54 + y * rowBytes + x * 3
      data[at] = b
      data[at + 1] = g
      data[at + 2] = r
    }
  }
  return data
}

/** A GIF whose LZW stream restarts before every pixel (valid and tiny to write). */
function gif(width = 40, height = 20) {
  const header = Buffer.from('GIF89a', 'latin1')
  const screen = Buffer.alloc(7)
  screen.writeUInt16LE(width, 0)
  screen.writeUInt16LE(height, 2)
  screen[4] = 0x80 // global colour table, 2 entries
  const palette = Buffer.from([...RED, ...BLUE])
  const descriptor = Buffer.alloc(10)
  descriptor[0] = 0x2c
  descriptor.writeUInt16LE(width, 5)
  descriptor.writeUInt16LE(height, 7)
  const codes = []
  for (let y = 0; y < height; y += 1) for (let x = 0; x < width; x += 1) codes.push(4, x < width / 2 ? 0 : 1)
  codes.push(5)
  const packed = []
  let buffer = 0
  let bits = 0
  for (const code of codes) {
    buffer |= code << bits
    bits += 3
    while (bits >= 8) { packed.push(buffer & 0xff); buffer >>= 8; bits -= 8 }
  }
  if (bits) packed.push(buffer & 0xff)
  const blocks = []
  for (let index = 0; index < packed.length; index += 255) {
    const chunk = packed.slice(index, index + 255)
    blocks.push(Buffer.from([chunk.length, ...chunk]))
  }
  return Buffer.concat([header, screen, palette, descriptor, Buffer.from([2]), ...blocks, Buffer.from([0, 0x3b])])
}

/** An uncompressed RGB TIFF with one page per entry of `pages` (each a width). */
function tiff(pages = [40, 30, 20], height = 20) {
  const header = Buffer.from([0x49, 0x49, 42, 0, 0, 0, 0, 0])
  const chunks = [header]
  let offset = header.length
  const entries = pages.map((width) => {
    const pixels = Buffer.alloc(width * height * 3)
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) pixels.set(pixelAt(x, width), (y * width + x) * 3)
    }
    chunks.push(pixels)
    const entry = { width, length: pixels.length, pixelOffset: offset }
    offset += pixels.length
    return entry
  })
  for (const entry of entries) {
    const bits = Buffer.alloc(6)
    bits.writeUInt16LE(8, 0); bits.writeUInt16LE(8, 2); bits.writeUInt16LE(8, 4)
    chunks.push(bits)
    entry.bitsOffset = offset
    offset += bits.length
  }
  const tagCount = 9
  const directorySize = 2 + tagCount * 12 + 4
  entries.forEach((entry, index) => { entry.directoryOffset = offset + index * directorySize })
  for (const [index, entry] of entries.entries()) {
    const tags = [
      [256, 4, 1, entry.width], [257, 4, 1, height], [258, 3, 3, entry.bitsOffset], [259, 3, 1, 1], [262, 3, 1, 2],
      [273, 4, 1, entry.pixelOffset], [277, 3, 1, 3], [278, 4, 1, height], [279, 4, 1, entry.length],
    ]
    const directory = Buffer.alloc(directorySize)
    directory.writeUInt16LE(tagCount, 0)
    tags.forEach(([tag, type, count, value], position) => {
      const at = 2 + position * 12
      directory.writeUInt16LE(tag, at)
      directory.writeUInt16LE(type, at + 2)
      directory.writeUInt32LE(count, at + 4)
      if (type === 3 && count === 1) directory.writeUInt16LE(value, at + 8)
      else directory.writeUInt32LE(value, at + 8)
    })
    directory.writeUInt32LE(index + 1 < entries.length ? entries[index + 1].directoryOffset : 0, directorySize - 4)
    chunks.push(directory)
  }
  header.writeUInt32LE(entries[0].directoryOffset, 4)
  return Buffer.concat(chunks)
}

function svg() {
  return Buffer.from('<?xml version="1.0" encoding="UTF-8"?>\n<!-- vector fixture -->\n<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 200 100"><rect width="100" height="100" fill="rgb(220,20,30)"/><rect x="100" width="100" height="100" fill="rgb(20,40,210)"/><text x="10" y="60" font-family="Arial" font-size="20" fill="#fff">SVG_VECTOR_TEXT</text></svg>\n', 'utf8')
}

function encodeSingleByte(text, table) {
  return Buffer.from(Array.from(text, (character) => {
    const code = character.charCodeAt(0)
    if (code < 0x80) return code
    if (table[character] === undefined) throw new Error(`No byte for ${character}`)
    return table[character]
  }))
}

const CP1252 = { 'é': 0xe9, 'ï': 0xef, '€': 0x80, 'ü': 0xfc, 'ñ': 0xf1, '—': 0x97 }
function cp1251(text) {
  return Buffer.from(Array.from(text, (character) => {
    const code = character.charCodeAt(0)
    if (code < 0x80) return code
    if (code >= 0x410 && code <= 0x44f) return code - 0x410 + 0xc0
    if (character === 'Ё') return 0xa8
    if (character === 'ё') return 0xb8
    throw new Error(`No byte for ${character}`)
  }))
}

function rtf() {
  return Buffer.from([
    '{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman\\fcharset0 Times New Roman;}{\\f1\\fswiss\\fcharset204 Arial;}}',
    '{\\colortbl;\\red200\\green0\\blue0;}',
    '\\pard\\qc\\b\\fs36 RTF_TITLE\\b0\\par',
    '\\pard\\fs24 Plain {\\b RTF_BOLD} and {\\i italic} text with caf\\\'e9 and {\\f1 \\\'cf\\\'f0\\\'e8\\\'e2\\\'e5\\\'f2} and \\u8364? euro.\\par',
    '{\\cf1 Red words}\\par',
    '\\trowd\\cellx3000\\cellx6000 \\intbl RTF_CELL_A\\cell RTF_CELL_B\\cell\\row',
    '\\pard\\page After the page break.\\par}',
  ].join('\n'), 'latin1')
}

async function odt() {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>')
  zip.file('content.xml', '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2"><office:automatic-styles><style:style style:name="B" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style></office:automatic-styles><office:body><office:text><text:h text:outline-level="1">ODT_HEADING</text:h><text:p>First <text:span text:style-name="B">ODT_BOLD</text:span> paragraph &amp; more.</text:p><text:list><text:list-item><text:p>ODT_ITEM_ONE</text:p></text:list-item></text:list><table:table><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell><text:p>ODT_CELL_A</text:p></table:table-cell><table:table-cell><text:p>ODT_CELL_B</text:p></table:table-cell></table:table-row></table:table></office:text></office:body></office:document-content>')
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function pptx() {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>')
  zip.file('ppt/presentation.xml', '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** The first bytes of a Word 97-2003 compound file: enough to be recognised. */
function legacyDoc() {
  const data = Buffer.alloc(2048)
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(data, 0)
  Buffer.from('WordDocument', 'utf16le').copy(data, 1024)
  return data
}

function largeText(megabytes = 5) {
  const lines = []
  let size = 0
  for (let index = 0; size < megabytes * 1024 * 1024; index += 1) {
    const line = `Line ${index} The quick brown fox jumps over the lazy dog. Съешь же ещё этих мягких французских булок.`
    lines.push(line)
    size += Buffer.byteLength(line) + 1
  }
  return Buffer.from(lines.join('\n'), 'utf8')
}

/**
 * Writes the fixtures to `directory` and returns their paths by key.
 * @param {string} directory
 * @param {{webp?: Buffer, large?: boolean}} [options] webp bytes come from the caller's encoder
 */
async function writeImportFixtures(directory, options = {}) {
  const files = {
    'utf16-bom.txt': Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('UTF16 Notepad text Привет\r\nSecond line', 'utf16le')]),
    'utf16-nobom.txt': Buffer.from('UTF16 without mark: plain words here\r\nand Привет', 'utf16le'),
    'cp1252.txt': encodeSingleByte('Café naïve résumé € price\r\nSecond — line', CP1252),
    'cp1251.txt': cp1251('Привет мир\r\nВторая строка'),
    'multi.txt': Buffer.from('Scripts: 你好 ✅ مرحبا Привет\n\fSecond page text', 'utf8'),
    'picture.bmp': bmp(),
    'picture.gif': gif(),
    'picture.svg': svg(),
    'pages.tif': tiff(),
    'notes.rtf': rtf(),
    'notes.odt': await odt(),
    'slides.pptx': await pptx(),
    'legacy.doc': legacyDoc(),
    'bad.xyz': Buffer.from([0x00, 0x13, 0x37, 0x00, 0x99, 0xfe, 0x00, 0x01, 0x02]),
    'png_named_as.jpg': null,
  }
  if (options.webp) files['picture.webp'] = options.webp
  if (options.png) files['png_named_as.jpg'] = options.png
  if (options.large) files['large.txt'] = largeText()
  const paths = {}
  for (const [name, data] of Object.entries(files)) {
    if (!data) continue
    const target = path.join(directory, name)
    fs.writeFileSync(target, data)
    paths[name] = target
  }
  return paths
}

module.exports = { BLUE, RED, bmp, gif, legacyDoc, odt, pptx, rtf, svg, tiff, writeImportFixtures }
