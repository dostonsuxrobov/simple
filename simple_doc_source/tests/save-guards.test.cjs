const test = require('node:test')
const assert = require('node:assert/strict')
const path = require('node:path')
const fs = require('node:fs')
const { pathToFileURL } = require('node:url')
const JSZip = require('jszip')

const guards = () => import('../src/ui-guards.js')

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
const IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c6360f8cf000000030101005d2bb1a40000000049454e44ae426082', 'hex')

const p = (inner) => `<w:p>${inner}</w:p>`
const r = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const field = (instr, result, nested = '') => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>${nested}<w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
const drawing = (graphicData) => `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="d"/><a:graphic>${graphicData}</a:graphic></wp:inline></w:drawing></w:r>`
const picture = (attributes) => drawing(`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip ${attributes}/></pic:blipFill></pic:pic></a:graphicData>`)
const relationship = (id, type, target, external = false) => `<Relationship Id="${id}" Type="${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`

async function buildDocx({ body, rels = [], parts = {}, mainPart = 'word/document.xml', compression = 'DEFLATE' }) {
  const zip = new JSZip()
  const options = { compression }
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/${mainPart}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`, options)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', mainPart)}</Relationships>`, options)
  zip.file(mainPart, `<?xml version="1.0" encoding="UTF-8"?><w:document ${W_NS}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`, options)
  const slash = mainPart.lastIndexOf('/')
  zip.file(`${mainPart.slice(0, slash + 1)}_rels/${mainPart.slice(slash + 1)}.rels`, `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`, options)
  for (const [name, content] of Object.entries(parts)) zip.file(name, content, options)
  return zip.generateAsync({ type: 'uint8array', compression })
}

function richFixture(options = {}) {
  const footer = `<?xml version="1.0" encoding="UTF-8"?><w:ftr ${W_NS}>${p(`${r('Page ')}<w:fldSimple w:instr=" PAGE \\* MERGEFORMAT ">${r('1')}</w:fldSimple>${r(' of ')}${field('NUMPAGES', '3')}${r(' ')}${field('FILENAME', 'report.docx')}`)}</w:ftr>`
  const footnotes = `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1">${p('<w:r><w:separator/></w:r>')}</w:footnote><w:footnote w:id="1">${p(`${r('See ')}${field('REF Target \\h', 'TARGET')}`)}</w:footnote></w:footnotes>`
  const comments = `<?xml version="1.0" encoding="UTF-8"?><w:comments ${W_NS}><w:comment w:id="1" w:author="A">${p(r('First'))}</w:comment><w:comment w:id="2" w:author="B">${p(r('Second'))}</w:comment></w:comments>`
  const body = [
    p(field('TOC \\o "1-3" \\h', 'Heading', field('PAGEREF _Toc1 \\h', '1'))),
    p(`${r('Merge ')}${field('MERGEFIELD Name', '«Name»')}${r(' on ')}${field('DATE \\@ "d MMMM yyyy"', '1 May 2026')}${r(' page ')}${field('PAGE', '1')}`),
    p(`<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Ann</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" PAGE ">${r('1')}</w:fldSimple>`),
    p(`<w:commentRangeStart w:id="1"/>${r('commented')}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>`),
    p(`<w:ins w:id="10" w:author="A">${r('inserted')}</w:ins><w:del w:id="11" w:author="A"><w:r><w:delText>gone</w:delText></w:r>${field('MERGEFIELD Deleted', 'x')}</w:del>`),
    p(`<w:r><w:rPr><w:b/><w:rPrChange w:id="12" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t>bold</w:t></w:r>`),
    p(drawing('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData>')),
    p(drawing('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdDm"/></a:graphicData>')),
    p(picture('r:embed="rIdEmf"')),
    p(picture('r:embed="rIdPng"')),
    p(picture('r:link="rIdLinked"')),
    p(picture('r:link="rIdWeb"')),
    p(`<w:r><w:object><v:shape id="ole" style="width:72pt;height:72pt"><v:imagedata r:id="rIdEmf2"/></v:shape><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rIdOle"/></w:object></w:r>`),
    '<w:altChunk r:id="rIdAlt"/>',
    p(`<w:r><w:pict><v:shape id="wm" type="#_x0000_t136"><v:textpath style="font-family:Calibri" string="DRAFT"/></v:shape></w:pict></w:r>`),
    p(`<w:r><mc:AlternateContent><mc:Choice Requires="wps">${drawing('<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>' + p(r('box')) + '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData>').replace(/^<w:r>|<\/w:r>$/g, '')}</mc:Choice><mc:Fallback><w:pict><v:shape><v:imagedata r:id="rIdEmf3"/></v:shape>${field('MERGEFIELD Fallback', 'y')}</w:pict></mc:Fallback></mc:AlternateContent></w:r>`),
    p(`<w:r><w:pict><v:shape id="logo" style="width:72pt;height:72pt"><v:imagedata r:id="rIdPng"/></v:shape></w:pict></w:r>`),
  ].join('')
  return buildDocx({
    ...options,
    body,
    rels: [
      relationship('rIdFooter', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', 'footer1.xml'),
      relationship('rIdNotes', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', 'footnotes.xml'),
      relationship('rIdComments', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comments.xml'),
      relationship('rIdEmf', IMAGE, 'media/image1.emf'),
      relationship('rIdEmf2', IMAGE, 'media/image2.emf'),
      relationship('rIdEmf3', IMAGE, 'media/image3.emf'),
      relationship('rIdPng', IMAGE, 'media/image4.png'),
      relationship('rIdLinked', IMAGE, 'file:///C:/Pictures/logo.png', true),
      relationship('rIdWeb', IMAGE, 'https://example.com/logo.png', true),
      relationship('rIdOle', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package', 'embeddings/sheet.xlsx'),
      relationship('rIdAlt', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk', 'chunk.html'),
    ],
    parts: {
      'word/footer1.xml': footer,
      'word/footnotes.xml': footnotes,
      'word/comments.xml': comments,
      'word/media/image4.png': PNG,
    },
  })
}

test('review keys ignore rebase bookkeeping and change with real review edits', async () => {
  const { reviewContentKey, reviewSummary, describeReviewItems } = await guards()
  const empty = { docId: 'local', baseVersion: 0, suggestions: [], threads: [] }
  assert.equal(reviewContentKey(empty), reviewContentKey({ ...empty, baseVersion: 9 }))
  assert.equal(reviewContentKey(null), reviewContentKey(empty))
  const thread = { id: 't1', anchor: { start: { blockId: 'p', offset: 0 }, end: { blockId: 'p', offset: 3 } }, status: 'open', comments: [] }
  const withThread = { ...empty, threads: [thread] }
  assert.notEqual(reviewContentKey(withThread), reviewContentKey(empty))
  assert.notEqual(reviewContentKey({ ...withThread, threads: [{ ...thread, status: 'resolved' }] }), reviewContentKey(withThread))
  assert.deepEqual(reviewSummary({ suggestions: [{}, {}], threads: [thread] }), { suggestions: 2, comments: 1, total: 3 })
  assert.deepEqual(reviewSummary(undefined), { suggestions: 0, comments: 0, total: 0 })
  assert.equal(describeReviewItems({ suggestions: 2, comments: 1 }), '2 suggestions and 1 comment')
  assert.equal(describeReviewItems({ suggestions: 1, comments: 0 }), '1 suggestion')
  assert.equal(describeReviewItems({ suggestions: 0, comments: 3 }), '3 comments')
  assert.equal(describeReviewItems({ suggestions: 0, comments: 0 }), '')
})

test('document content keys ignore layout revisions and run boundaries, nothing else', async () => {
  const { documentContentKey } = await guards()
  const style = { fontFamily: 'Calibri', bold: false }
  const doc = (blocks) => ({ section: { pageWidthPx: 816 }, blocks })
  const original = doc([{ kind: 'paragraph', id: 'p1', revision: 0, runs: [{ text: 'Hello world', style }], style: { align: 'left' } }])
  const undone = doc([{ kind: 'paragraph', id: 'p1', revision: 7, runs: [{ text: 'Hello', style }, { text: '', style }, { text: ' world', style }], style: { align: 'left' } }])
  assert.equal(documentContentKey(undone), documentContentKey(original), 'undo leaves equal content in new objects')
  assert.notEqual(documentContentKey(doc([{ ...original.blocks[0], runs: [{ text: 'Hello World', style }] }])), documentContentKey(original), 'text changes count')
  assert.notEqual(documentContentKey(doc([{ ...original.blocks[0], runs: [{ text: 'Hello', style: { ...style, bold: true } }, { text: ' world', style }] }])), documentContentKey(original), 'formatting changes count')
  assert.notEqual(documentContentKey(doc([{ ...original.blocks[0], style: { align: 'center' } }])), documentContentKey(original), 'paragraph changes count')
  assert.notEqual(documentContentKey({ ...original, section: { pageWidthPx: 1056 } }), documentContentKey(original), 'page setup counts')
  const table = (text, revision) => doc([{ kind: 'table', id: 't', revision, rows: [{ cells: [{ id: 'c', blocks: [{ kind: 'paragraph', id: 'cp', revision, runs: [{ text, style }], style: {} }] }] }] }])
  assert.equal(documentContentKey(table('A', 1)), documentContentKey(table('A', 4)))
  assert.notEqual(documentContentKey(table('A', 1)), documentContentKey(table('B', 1)))
  const emptyParagraph = (runs) => doc([{ kind: 'paragraph', id: 'e', revision: 0, runs, style: {} }])
  assert.equal(documentContentKey(emptyParagraph([{ text: '', style }])), documentContentKey(emptyParagraph([{ text: '', style }, { text: '', style }])))
})

test('open errors are plain sentences; only a missing file counts as gone', async () => {
  const { cleanErrorMessage, isMissingFileError, openErrorMessage } = await guards()
  const ipc = (message) => new Error(`Error invoking remote method 'file:open-path': Error: ${message}`)
  assert.equal(cleanErrorMessage(ipc('Password-protected Word documents are not supported.')), 'Password-protected Word documents are not supported.')
  assert.equal(cleanErrorMessage(new Error('   '), 'Fallback'), 'Fallback')
  assert.equal(cleanErrorMessage('plain'), 'plain')
  const missing = ipc("ENOENT: no such file or directory, stat 'C:\\Docs\\gone.docx'")
  assert.equal(isMissingFileError(missing), true)
  assert.equal(isMissingFileError(ipc('This DOCX archive is incomplete.')), false)
  assert.equal(isMissingFileError(ipc("EBUSY: resource busy or locked, open 'C:\\Docs\\locked.docx'")), false)
  assert.match(openErrorMessage(missing), /no longer available/)
  assert.match(openErrorMessage(ipc("EPERM: operation not permitted, open 'C:\\x.docx'")), /another program may be using it/i)
  assert.equal(openErrorMessage(ipc('Password-protected Word documents are not supported. Remove the password in Word, then try the .docx file again.')), 'Password-protected Word documents are not supported. Remove the password in Word, then try the .docx file again.')
})

test('recovered work opens as a clearly named untitled copy, never name.doc.docx', async () => {
  const { recoveredDocumentName } = await guards()
  assert.equal(recoveredDocumentName('report.docx'), 'report (recovered)')
  assert.equal(recoveredDocumentName('report (recovered).docx'), 'report (recovered)')
  assert.equal(recoveredDocumentName('Q3 plan v2.1.docx'), 'Q3 plan v2.1 (recovered)')
  assert.equal(recoveredDocumentName('legacy-report.doc'), 'legacy-report (recovered)')
  assert.equal(recoveredDocumentName(''), 'Untitled document (recovered)')
  assert.doesNotMatch(recoveredDocumentName('legacy-report.doc.docx'), /\.doc/)
})

test('toasts stay long enough to read', async () => {
  const { notificationDuration } = await guards()
  assert.equal(notificationDuration('Saved.'), 3200)
  assert.ok(notificationDuration('x'.repeat(200)) > 3200)
  assert.ok(notificationDuration('Short error', 'error') >= 6000)
  assert.ok(notificationDuration('x'.repeat(5000)) <= 12000)
})

test('the fidelity scan lists exactly the content the editor drops', async () => {
  const { scanDocxFidelity, describeFidelityItems } = await guards()
  const report = await scanDocxFidelity(await richFixture())
  const counts = Object.fromEntries(report.items.map((item) => [item.kind, item.count]))
  assert.deepEqual(Object.keys(counts), ['comments', 'trackedChanges', 'charts', 'smartArt', 'unsupportedPictures', 'linkedPictures', 'embeddedObjects', 'wordArt', 'fields'])
  assert.equal(counts.comments, 2)
  assert.ok(counts.trackedChanges >= 3, 'insertions, deletions and formatting changes')
  assert.equal(counts.charts, 1)
  assert.equal(counts.smartArt, 1)
  assert.equal(counts.unsupportedPictures, 1, 'the OLE preview and the mc:Fallback picture are not counted as pictures')
  assert.equal(counts.linkedPictures, 1, 'web-linked pictures still load; only local links are lost')
  assert.equal(counts.embeddedObjects, 2, 'one OLE object and one altChunk')
  assert.equal(counts.wordArt, 1)
  // TOC (with its nested PAGEREF counted once), MERGEFIELD, simple AUTHOR and
  // body PAGE fields, the footer FILENAME and the footnote REF. Kept: DATE and
  // PAGE in the body, PAGE and NUMPAGES in the footer; fields inside a deletion
  // or an mc:Fallback copy are never imported at all.
  assert.equal(counts.fields, 6)
  const text = describeFidelityItems(report.items)
  assert.match(text, /^2 comments, tracked changes \(shown as accepted\), 1 chart, 1 SmartArt graphic, 1 picture in EMF format, 1 linked picture, 2 embedded objects, 1 WordArt object and 6 fields \(kept as plain text\)$/)
})

test('the fidelity scan finds nothing in documents the editor keeps whole', async () => {
  const { scanDocxFidelity } = await guards()
  const plain = await buildDocx({ body: p(r('Plain text')) + p(picture('r:embed="rIdPng"')) + p(field('DATE', '1 May 2026')), rels: [relationship('rIdPng', IMAGE, 'media/image1.png')], parts: { 'word/media/image1.png': PNG } })
  assert.deepEqual((await scanDocxFidelity(plain)).items, [])
  const qaFixture = fs.readFileSync(path.join(__dirname, '..', 'qa', 'fixtures', 'simple-docs-roundtrip-fixture.docx'))
  assert.deepEqual((await scanDocxFidelity(new Uint8Array(qaFixture))).items, [], 'the QA round-trip fixture saves without a copy prompt')
})

test('the fidelity scan follows package relationships and reads stored entries', async () => {
  const { scanDocxFidelity, listZipEntries } = await guards()
  const stored = await richFixture({ mainPart: 'word/main.xml', compression: 'STORE' })
  assert.ok([...listZipEntries(stored).values()].every((entry) => entry.method === 0))
  const report = await scanDocxFidelity(stored)
  assert.equal(report.items.find((item) => item.kind === 'comments')?.count, 2)
  assert.equal(report.items.find((item) => item.kind === 'fields')?.count, 6)
  await assert.rejects(scanDocxFidelity(new TextEncoder().encode('not a zip')), /not a Word package/)
})

test('fidelity claims match a real WordCanvas import and export', { timeout: 120000 }, async () => {
  const { scanDocxFidelity } = await guards()
  const engine = (file) => import(pathToFileURL(path.join(__dirname, '..', 'node_modules', '@forevka', 'wordcanvas', 'dist-node', file)).href)
  const { runImport } = await engine('import.js')
  const { runExport } = await engine('export.js')
  const { installMeasureHost } = await engine('measure.js')
  await installMeasureHost()
  const roundTrip = async (bytes) => {
    const imported = runImport(bytes, undefined, { collectMediaBytes: true })
    const images = Object.fromEntries((imported.media || []).map((media) => [media.src, media.bytes]))
    const exported = await runExport(imported.doc, 'docx', images)
    const zip = await JSZip.loadAsync(exported.bytes)
    return zip.file('word/document.xml').async('string')
  }
  const cases = [
    { body: p(field('DATE \\@ "M/d/yyyy"', '1/1/2026')), marker: /\bDATE\b/, kept: true },
    { body: p(field('PAGE', '1')), marker: /\bPAGE\b/, kept: true },
    { body: p(field('MERGEFIELD Name', 'NAME-RESULT')), marker: /MERGEFIELD/, kept: false },
    { body: p(field('TOC \\o "1-3"', 'Heading')), marker: /\bTOC\b/, kept: false },
    { body: p(`<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Ann</w:t></w:r></w:fldSimple>`), marker: /AUTHOR/, kept: false },
    { body: p(`<w:ins w:id="1" w:author="A">${r('INS')}</w:ins><w:del w:id="2" w:author="A"><w:r><w:delText>DELETED-TEXT</w:delText></w:r></w:del>`), marker: /DELETED-TEXT/, kept: false },
    { body: p(drawing('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData>')), marker: /<w:drawing/, kept: false },
    { body: p(picture('r:embed="rIdImg"')), rels: [relationship('rIdImg', IMAGE, 'media/image1.emf')], parts: { 'word/media/image1.emf': Buffer.from('01000000', 'hex') }, marker: /<w:drawing/, kept: false },
    { body: p(picture('r:embed="rIdImg"')), rels: [relationship('rIdImg', IMAGE, 'media/image1.png')], parts: { 'word/media/image1.png': PNG }, marker: /<w:drawing/, kept: true },
    { body: p(`<w:r><w:pict><v:shape id="tb" style="width:157pt;height:47pt" type="#_x0000_t202"><v:textbox><w:txbxContent>${p(r('TEXTBOX-TEXT'))}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`), marker: /TEXTBOX-TEXT/, kept: true },
    { body: p(`<w:r><w:pict><v:shape id="wm" type="#_x0000_t136" style="width:400pt;height:100pt"><v:textpath style="font-family:Calibri" string="WORDART-TEXT"/></v:shape></w:pict></w:r>`), marker: /WORDART-TEXT/, kept: false },
  ]
  for (const { marker, kept, ...fixture } of cases) {
    const bytes = await buildDocx(fixture)
    const flagged = (await scanDocxFidelity(bytes)).items.length > 0
    const survived = marker.test(await roundTrip(bytes))
    assert.equal(survived, kept, `engine round trip of ${marker}`)
    assert.equal(flagged, !kept, `scan flags ${marker} exactly when the engine drops it`)
  }
})
