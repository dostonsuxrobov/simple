const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { createFontCatalog, faceStyle, fontMetrics, fontNames, hasUnicodeCmap, installedDocumentFonts, scanInstalledFontFamilies } = require('../electron/document-fonts.cjs')
const { fontCoverage } = require('../electron/font-coverage.cjs')

// A small but complete installed face: name (Windows English + Mac), head, hhea, OS/2 and cmap tables.
function installedFace({ family, subfamily = 'Regular', selection = 0x40, weight = 400, fsType = 0, cmap = [[3, 1]], macFamily } = {}) {
  const utf16 = (text) => Buffer.from(text, 'utf16le').swap16()
  const names = [[3, 1, 0x0409, 1, utf16(family)], [3, 1, 0x0409, 2, utf16(subfamily)], [3, 1, 0x0419, 1, utf16(`${family} RU`)]]
  if (macFamily) names.unshift([1, 0, 0, 1, Buffer.from(macFamily, 'latin1')])
  const storage = Buffer.concat(names.map((entry) => entry[4]))
  const name = Buffer.alloc(6 + names.length * 12 + storage.length)
  name.writeUInt16BE(0, 0)
  name.writeUInt16BE(names.length, 2)
  name.writeUInt16BE(6 + names.length * 12, 4)
  let offset = 0
  names.forEach(([platform, encoding, language, id, bytes], index) => {
    const at = 6 + index * 12
    name.writeUInt16BE(platform, at)
    name.writeUInt16BE(encoding, at + 2)
    name.writeUInt16BE(language, at + 4)
    name.writeUInt16BE(id, at + 6)
    name.writeUInt16BE(bytes.length, at + 8)
    name.writeUInt16BE(offset, at + 10)
    offset += bytes.length
  })
  storage.copy(name, 6 + names.length * 12)
  const head = Buffer.alloc(54)
  head.writeUInt16BE(2048, 18)
  head.writeUInt16BE((selection & 0x20 ? 1 : 0) | (selection & 0x01 ? 2 : 0), 44)
  const hhea = Buffer.alloc(36)
  hhea.writeInt16BE(1854, 4)
  hhea.writeInt16BE(-434, 6)
  hhea.writeInt16BE(67, 8)
  const os2 = Buffer.alloc(96)
  os2.writeUInt16BE(weight, 4)
  os2.writeUInt16BE(fsType, 8)
  os2.writeUInt16BE(selection, 62)
  os2.writeUInt16BE(1854, 74)
  os2.writeUInt16BE(434, 76)
  const cmapTable = Buffer.alloc(4 + cmap.length * 8)
  cmapTable.writeUInt16BE(cmap.length, 2)
  cmap.forEach(([platform, encoding], index) => {
    cmapTable.writeUInt16BE(platform, 4 + index * 8)
    cmapTable.writeUInt16BE(encoding, 6 + index * 8)
  })
  const tables = [['OS/2', os2], ['cmap', cmapTable], ['head', head], ['hhea', hhea], ['name', name]]
  const directory = 12 + tables.length * 16
  const total = directory + tables.reduce((sum, [, data]) => sum + data.length, 0)
  const bytes = Buffer.alloc(total)
  bytes.writeUInt32BE(0x00010000, 0)
  bytes.writeUInt16BE(tables.length, 4)
  let at = directory
  tables.forEach(([tag, data], index) => {
    bytes.write(tag, 12 + index * 16, 'latin1')
    bytes.writeUInt32BE(at, 20 + index * 16)
    bytes.writeUInt32BE(data.length, 24 + index * 16)
    data.copy(bytes, at)
    at += data.length
  })
  return bytes
}

function fixtureFont(flags = 0) {
  const bytes = Buffer.alloc(124)
  bytes.writeUInt32BE(0x00010000, 0)
  bytes.writeUInt16BE(3, 4)
  for (const [index, name, offset, length] of [[0, 'head', 60, 20], [1, 'hhea', 80, 10], [2, 'OS/2', 100, 24]]) {
    bytes.write(name, 12 + index * 16)
    bytes.writeUInt32BE(offset, 20 + index * 16)
    bytes.writeUInt32BE(length, 24 + index * 16)
  }
  bytes.writeUInt16BE(2048, 78)
  bytes.writeInt16BE(1900, 84)
  bytes.writeInt16BE(-500, 86)
  bytes.writeUInt16BE(flags, 108)
  return bytes
}

// A minimal SFNT with head, hhea and (optionally) a full-size OS/2 table.
function metricsFont({ ascender, descender, lineGap, winAscent, winDescent, os2 = true }) {
  const tables = [['head', 20], ['hhea', 10], ...(os2 ? [['OS/2', 96]] : [])]
  const directory = 12 + tables.length * 16
  const bytes = Buffer.alloc(directory + 20 + 10 + (os2 ? 96 : 0))
  bytes.writeUInt32BE(0x00010000, 0)
  bytes.writeUInt16BE(tables.length, 4)
  let offset = directory
  const at = {}
  tables.forEach(([name, length], index) => {
    bytes.write(name, 12 + index * 16)
    bytes.writeUInt32BE(offset, 20 + index * 16)
    bytes.writeUInt32BE(length, 24 + index * 16)
    at[name] = offset
    offset += length
  })
  bytes.writeUInt16BE(2048, at.head + 18)
  bytes.writeInt16BE(ascender, at.hhea + 4)
  bytes.writeInt16BE(descender, at.hhea + 6)
  bytes.writeInt16BE(lineGap, at.hhea + 8)
  if (os2) {
    bytes.writeUInt16BE(winAscent, at['OS/2'] + 74)
    bytes.writeUInt16BE(winDescent, at['OS/2'] + 76)
  }
  return bytes
}

const lineEm = (metrics) => Number((metrics.ascent + metrics.descent).toFixed(4))

test('installed font line boxes match Word/GDI single spacing, including the hhea line gap', () => {
  // Calibri: hhea 1536/-512 gap 452, usWin 1950/550. Word's single line is 2500/2048 em.
  const calibri = fontMetrics(metricsFont({ ascender: 1536, descender: -512, lineGap: 452, winAscent: 1950, winDescent: 550 }))
  assert.deepEqual(calibri, { ascent: 1950 / 2048, descent: 550 / 2048 })
  assert.equal(lineEm(calibri), 1.2207)
  // Consolas: the gap is entirely inside the Windows ascent/descent.
  assert.equal(lineEm(fontMetrics(metricsFont({ ascender: 1521, descender: -527, lineGap: 350, winAscent: 1884, winDescent: 514 }))), 1.1709)
  // Times New Roman: usWin equals hhea, so the 87-unit gap is external leading (13.8 pt at 12 pt).
  const times = fontMetrics(metricsFont({ ascender: 1825, descender: -443, lineGap: 87, winAscent: 1825, winDescent: 443 }))
  assert.deepEqual(times, { ascent: 1825 / 2048, descent: (443 + 87) / 2048 })
  assert.equal(Number(((times.ascent + times.descent) * 12).toFixed(2)), 13.8)
  // Without a usable OS/2 table the hhea ascender, descender and line gap are used.
  assert.deepEqual(fontMetrics(metricsFont({ ascender: 1500, descender: -500, lineGap: 200, os2: false })), { ascent: 1500 / 2048, descent: 700 / 2048 })
  assert.deepEqual(fontMetrics(metricsFont({ ascender: 1500, descender: -500, lineGap: 200, winAscent: 0, winDescent: 0 })), { ascent: 1500 / 2048, descent: 700 / 2048 })
  // A negative line gap never shrinks the line.
  assert.deepEqual(fontMetrics(metricsFont({ ascender: 1500, descender: -500, lineGap: -300, os2: false })), { ascent: 1500 / 2048, descent: 500 / 2048 })
})

test('installed fonts use real ascent/descent and reject malformed or restricted faces', () => {
  assert.deepEqual(fontMetrics(fixtureFont()), { ascent: 1900 / 2048, descent: 500 / 2048 })
  assert.equal(fontMetrics(fixtureFont(2)), null)
  assert.equal(fontMetrics(fixtureFont(512)), null)
  assert.equal(fontMetrics(Buffer.from('not a font')), null)
  assert.equal(fontMetrics(fixtureFont().subarray(0, 80)), null)
})

test('font manifest serves only existing allowlisted faces, leaving missing families to fallbacks', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-doc-font-test-'))
  try {
    await fs.writeFile(path.join(directory, 'arial.ttf'), fixtureFont())
    await fs.writeFile(path.join(directory, 'arialbd.ttf'), fixtureFont())
    await fs.writeFile(path.join(directory, 'arbitrary.ttf'), fixtureFont())
    const { fonts, files } = await installedDocumentFonts({ roots: [directory] })
    assert.equal(fonts.length, 1)
    assert.equal(fonts[0].family, 'Arial')
    assert.equal(Object.keys(fonts[0].faces).length, 2)
    assert.ok(fonts[0].faces.regular.startsWith('simple-font://installed/'))
    assert.equal(files.size, 2)
  } finally { await fs.rm(directory, { recursive: true, force: true }) }
})

test('glyph coverage reads mapped Unicode ranges and rejects truncated tables', async () => {
  const bytes = Buffer.alloc(80)
  bytes.writeUInt32BE(0x10000, 0)
  bytes.writeUInt16BE(1, 4)
  bytes.write('cmap', 12)
  bytes.writeUInt32BE(28, 20)
  bytes.writeUInt32BE(52, 24)
  bytes.writeUInt16BE(1, 30)
  bytes.writeUInt16BE(3, 32)
  bytes.writeUInt16BE(10, 34)
  bytes.writeUInt32BE(12, 36)
  bytes.writeUInt16BE(12, 40)
  bytes.writeUInt32BE(40, 44)
  bytes.writeUInt32BE(2, 52)
  bytes.writeUInt32BE(0x4e00, 56)
  bytes.writeUInt32BE(0x4e02, 60)
  bytes.writeUInt32BE(0, 64)
  bytes.writeUInt32BE(0xac00, 68)
  bytes.writeUInt32BE(0xac01, 72)
  bytes.writeUInt32BE(15, 76)
  assert.deepEqual(fontCoverage(bytes), [[0x4e01, 0x4e02], [0xac00, 0xac01]])
  assert.deepEqual(fontCoverage(bytes.subarray(0, 60)), [])
  assert.deepEqual(fontCoverage(Buffer.from('not a font')), [])
  const bundled = fontCoverage(await fs.readFile(path.join(__dirname, '../node_modules/@forevka/wordcanvas/dist-node/fonts/NotoSansSC-Regular.ttf')))
  const has = (char) => bundled.some(([start, end]) => char.codePointAt(0) >= start && char.codePointAt(0) <= end)
  assert.equal(has('测'), true)
  assert.equal(has('語'), false)
  assert.equal(has('한'), false)
})

test('installed faces report their Windows English family, style and Unicode support', () => {
  const face = installedFace({ family: 'Century Gothic', subfamily: 'Bold Italic', selection: 0x21, weight: 700, macFamily: 'CenturyGothicMac' })
  const tables = new Map()
  for (let index = 0; index < face.readUInt16BE(4); index += 1) {
    const at = 12 + index * 16
    tables.set(face.toString('latin1', at, at + 4), face.subarray(face.readUInt32BE(at + 8), face.readUInt32BE(at + 8) + face.readUInt32BE(at + 12)))
  }
  assert.deepEqual(fontNames(tables.get('name')), { family: 'Century Gothic', subfamily: 'Bold Italic' })
  assert.equal(faceStyle(tables.get('head'), tables.get('OS/2')), 'boldItalic')
  assert.equal(faceStyle(tables.get('head'), null), 'boldItalic', 'head macStyle is the fallback')
  assert.equal(hasUnicodeCmap(tables.get('cmap')), true)
  assert.equal(fontNames(Buffer.alloc(2)).family, null)
  assert.equal(fontMetrics(face) !== null, true)
})

test('every installed TTF/OTF family is listed with its styles; symbol, restricted and collection files are left out', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-doc-font-scan-'))
  const user = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-doc-font-user-'))
  try {
    await fs.writeFile(path.join(directory, 'gothic.ttf'), installedFace({ family: 'Century Gothic' }))
    await fs.writeFile(path.join(directory, 'gothicb.ttf'), installedFace({ family: 'Century Gothic', subfamily: 'Bold', selection: 0x20, weight: 700 }))
    await fs.writeFile(path.join(directory, 'gothici.TTF'), installedFace({ family: 'Century Gothic', subfamily: 'Italic', selection: 0x01 }))
    await fs.writeFile(path.join(directory, 'gothicbi.otf'), installedFace({ family: 'Century Gothic', subfamily: 'Bold Italic', selection: 0x21, weight: 700 }))
    // A second "regular" face that is further from 400 does not replace the first.
    await fs.writeFile(path.join(directory, 'gothic-light.ttf'), installedFace({ family: 'Century Gothic', weight: 300 }))
    await fs.writeFile(path.join(directory, 'black.ttf'), installedFace({ family: 'Arial Black', subfamily: 'Bold', selection: 0x20, weight: 900 }))
    await fs.writeFile(path.join(directory, 'wingding.ttf'), installedFace({ family: 'Wingdings', cmap: [[3, 0]] }))
    await fs.writeFile(path.join(directory, 'secret.ttf'), installedFace({ family: 'Licensed Only', fsType: 0x0002 }))
    await fs.writeFile(path.join(directory, 'cambria.ttc'), Buffer.from('ttcf'))
    await fs.writeFile(path.join(directory, 'broken.ttf'), Buffer.from('not a font'))
    await fs.writeFile(path.join(directory, 'notes.txt'), 'hello')
    await fs.writeFile(path.join(user, 'bahnschrift.ttf'), installedFace({ family: 'Bahnschrift' }))
    const families = await scanInstalledFontFamilies({ roots: [directory, user, path.join(directory, 'missing')] })
    assert.deepEqual(families.map((entry) => entry.family), ['Arial Black', 'Bahnschrift', 'Century Gothic'])
    const gothic = families.find((entry) => entry.family === 'Century Gothic')
    assert.deepEqual(Object.keys(gothic.faces).sort(), ['bold', 'boldItalic', 'italic', 'regular'])
    assert.equal(path.basename(gothic.faces.regular.file), 'gothic.ttf')
    assert.equal(path.basename(gothic.faces.bold.file), 'gothicb.ttf')
    const black = families.find((entry) => entry.family === 'Arial Black')
    assert.deepEqual(Object.keys(black.faces), ['regular'], 'a single bold-only face serves as the regular face')

    const catalog = createFontCatalog({ roots: [directory, user] })
    assert.deepEqual(await catalog.familyList(), [{ family: 'Arial Black', styles: 1 }, { family: 'Bahnschrift', styles: 1 }, { family: 'Century Gothic', styles: 4 }])
    const fonts = await catalog.documentFonts(['century gothic', 'Bahnschrift', 'Wingdings', 'Not Installed', 'Century  Gothic'])
    assert.deepEqual(fonts.map((font) => font.family), ['Century Gothic', 'Bahnschrift'], 'unknown and duplicate requests are skipped')
    const font = fonts[0]
    assert.deepEqual(Object.keys(font.faces).sort(), ['bold', 'boldItalic', 'italic', 'regular'])
    assert.ok(font.sizing.ascent > 0.9 && font.sizing.descent > 0.2)
    const token = font.faces.bold.split('/').pop()
    assert.equal(path.basename(await catalog.fileForToken(token)), 'gothicb.ttf')
    assert.equal(await catalog.fileForToken('999'), null)
    // The same file keeps its token across requests.
    assert.equal((await catalog.documentFonts(['Century Gothic']))[0].faces.bold, font.faces.bold)
    assert.deepEqual(await catalog.documentFonts(), [], 'no curated fonts exist in the fixture folders')
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
    await fs.rm(user, { recursive: true, force: true })
  }
})

test('the catalog serves the curated set first and never more than 32 extra families', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-doc-font-limit-'))
  try {
    await fs.writeFile(path.join(directory, 'arial.ttf'), installedFace({ family: 'Arial' }))
    const wanted = []
    for (let index = 0; index < 40; index += 1) {
      const family = `Family ${String(index).padStart(2, '0')}`
      wanted.push(family)
      await fs.writeFile(path.join(directory, `family${index}.ttf`), installedFace({ family }))
    }
    const catalog = createFontCatalog({ roots: [directory] })
    const fonts = await catalog.documentFonts(['Arial', ...wanted])
    assert.equal(fonts[0].family, 'Arial')
    assert.equal(fonts.length, 1 + 32)
    assert.equal(path.basename(await catalog.fileForToken(fonts[0].faces.regular.split('/').pop())), 'arial.ttf')
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
})
