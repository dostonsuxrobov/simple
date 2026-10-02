const test = require('node:test')
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { detectImage, repairDocxPackage } = require('../electron/docx-repair.cjs')
const { validateDocxBytes } = require('../electron/docx-files.cjs')

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')
const WEBP = Buffer.concat([Buffer.from('RIFF'), Buffer.from([30, 0, 0, 0]), Buffer.from('WEBPVP8L'), Buffer.alloc(22)])
const BMP = Buffer.concat([Buffer.from('BM'), Buffer.alloc(60)])
const SVG = Buffer.from('<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg" width="2" height="2"/>')

async function fixtureDocx({ media, documentXml }) {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
  zip.file('word/document.xml', documentXml)
  const relationships = Object.keys(media).map((name, index) => `<Relationship Id="rId${index + 10}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`).join('')
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`)
  zip.file('word/_rels/header1.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${Object.keys(media)[0]}"/></Relationships>`)
  for (const [name, bytes] of Object.entries(media)) zip.file(`word/media/${name}`, bytes)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const BODY = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p/><w:sectPr/></w:body></w:document>'

test('image signatures are recognized from the bytes, not the name', () => {
  assert.equal(detectImage(PNG).extension, 'png')
  assert.equal(detectImage(WEBP).extension, 'webp')
  assert.equal(detectImage(BMP).extension, 'bmp')
  assert.equal(detectImage(SVG).contentType, 'image/svg+xml')
  assert.equal(detectImage(Buffer.from('II*\u0000xxxx')).extension, 'tiff')
  assert.equal(detectImage(Buffer.from('plain text')), null)
})

test('DOCX media keep a truthful name and content type (DOC-SIE-17)', async () => {
  const input = await fixtureDocx({ media: { 'image1.png': PNG, 'image2.png': WEBP, 'image3.png': BMP, 'image4.png': SVG }, documentXml: BODY })
  const result = await repairDocxPackage(input)
  assert.equal(result.changed, true)
  assert.match(result.repairs.join(' '), /3 pictures were relabeled/)
  assert.deepEqual(result.warnings, [])
  validateDocxBytes(result.bytes)
  const zip = await JSZip.loadAsync(result.bytes)
  const media = Object.keys(zip.files).filter((name) => name.startsWith('word/media/') && !name.endsWith('/')).sort()
  assert.deepEqual(media, ['word/media/image1.png', 'word/media/image2.webp', 'word/media/image3.bmp', 'word/media/image4.svg'])
  for (const name of media) {
    const bytes = await zip.file(name).async('nodebuffer')
    assert.equal(detectImage(bytes).extension, name.split('.').pop(), `${name} magic matches its extension`)
  }
  const types = await zip.file('[Content_Types].xml').async('string')
  assert.match(types, /<Default Extension="webp" ContentType="image\/webp"\/>/)
  assert.match(types, /<Default Extension="bmp" ContentType="image\/bmp"\/>/)
  assert.match(types, /<Default Extension="svg" ContentType="image\/svg\+xml"\/>/)
  const relationships = await zip.file('word/_rels/document.xml.rels').async('string')
  assert.match(relationships, /Target="media\/image2\.webp"/)
  assert.doesNotMatch(relationships, /image2\.png/)
  // A correctly labeled package is returned untouched, byte for byte.
  const clean = await fixtureDocx({ media: { 'image1.png': PNG }, documentXml: BODY })
  const untouched = await repairDocxPackage(clean)
  assert.equal(untouched.changed, false)
  assert.equal(untouched.bytes, clean)
})

test('mid-document sections never keep empty header/footer relationship ids (DOC-SIE-19 safety net)', async () => {
  const documentXml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id=""/><w:footerReference w:type="first" r:id=""/><w:pgSz w:w="12240" w:h="15840"/><w:titlePg/></w:sectPr></w:pPr></w:p><w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId9"/></w:sectPr></w:body></w:document>'
  const result = await repairDocxPackage(await fixtureDocx({ media: { 'image1.png': PNG }, documentXml }))
  assert.equal(result.changed, true)
  assert.match(result.warnings.join(' '), /section 1/)
  const xml = await (await JSZip.loadAsync(result.bytes)).file('word/document.xml').async('string')
  assert.doesNotMatch(xml, /r:id=""/)
  assert.match(xml, /<w:headerReference w:type="default" r:id="rId9"\/>/)
  assert.match(xml, /<w:pgSz w:w="12240" w:h="15840"\/>/)
})

test('locates DOC-SIE-19: WordCanvas 0.12.0 writes the first section\'s bands with empty ids and no parts', async () => {
  const { runExport } = await import('@forevka/wordcanvas/export')
  const { installMeasureHost } = await import('@forevka/wordcanvas/export/measure')
  await installMeasureHost()
  const style = { fontFamily: 'Calibri', fontSizePx: 14.667, bold: false, italic: false, underline: false, strikethrough: false, color: '#111111' }
  const paragraph = (id, text, extra = {}) => ({ kind: 'paragraph', id, revision: 0, runs: [{ text, style }], style: { align: 'left', lineHeight: 1.15, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0, ...extra } })
  const margins = { top: 96, right: 96, bottom: 96, left: 96 }
  const document = {
    section: { pageWidthPx: 1056, pageHeightPx: 816, marginPx: margins, header: [paragraph('h2', 'SECTION_TWO_HEADER')] },
    blocks: [
      paragraph('p1', 'Section one', { sectionBreak: { type: 'nextPage', props: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: margins, header: [paragraph('h1', 'SECTION_ONE_HEADER')], footer: [paragraph('f1', 'SECTION_ONE_FOOTER')] } } }),
      paragraph('p2', 'Section two'),
    ],
  }
  const { bytes } = await runExport(document, 'docx')
  const exported = await JSZip.loadAsync(bytes)
  const xml = await exported.file('word/document.xml').async('string')
  const parts = await Promise.all(Object.keys(exported.files).filter((name) => /^word\/(header|footer)\d+\.xml$/.test(name)).map((name) => exported.file(name).async('string')))
  const engineFixed = !/r:id=""/.test(xml)
  if (engineFixed) {
    // Once scripts/patch-wordcanvas.cjs carries the band fix, every section keeps its parts.
    assert.ok(parts.some((part) => part.includes('SECTION_ONE_HEADER')))
    return
  }
  // Unpatched engine: the paragraph-level sectPr writer passes `() => ""` as its
  // band allocator, so section 1's header/footer XML is never written.
  assert.ok(!parts.some((part) => part.includes('SECTION_ONE_HEADER')), 'section 1 header content is absent from the package')
  const repaired = await repairDocxPackage(bytes)
  assert.match(repaired.warnings.join(' '), /section 1/)
  assert.doesNotMatch(await (await JSZip.loadAsync(repaired.bytes)).file('word/document.xml').async('string'), /r:id=""/)
})
