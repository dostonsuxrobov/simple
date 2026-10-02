const test = require('node:test')
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { OLE_COMPOUND_FILE_SIGNATURE, convertLegacyDocToDocx, sniffLegacyDocContent } = require('../electron/legacy-doc.cjs')
const { parseRtf } = require('../electron/rtf-import.cjs')
const { htmlToFlow, mhtToFlow, word2003ToFlow } = require('../electron/html-import.cjs')
const { buildFlowDocx, flowText } = require('../electron/simple-docx.cjs')
const { validateDocxPackage } = require('../electron/docx-files.cjs')

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')
// No local Office engine: the native readers must do the whole job.
const NO_ENGINE = { convertOffice: async () => { throw new Error('No local Office engine.') } }

async function importDocx(bytes) {
  validateDocxPackage(bytes)
  const { runImport } = await import('@forevka/wordcanvas/import')
  return runImport(new Uint8Array(bytes), undefined, { collectMediaBytes: true }).doc
}

const text = (block) => (block.runs || []).map((run) => run.text).join('')
const paragraphs = (doc) => doc.blocks.filter((block) => block.kind === 'paragraph')

const RTF = String.raw`{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fcharset0 Arial;}{\f1\froman\fcharset204 Times New Roman Cyr;}{\f3\fnil\fcharset2 Symbol;}}
{\colortbl;\red0\green0\blue0;\red192\green0\blue0;\red255\green255\blue0;}
{\stylesheet{\s0 Normal;}{\s1\b\fs32 heading 1;}}
{\info{\title Quarterly RTF}{\author Someone}}
{\header \pard\qc RTF_HEADER\par}
\pard\s1 RTF Heading\par
\pard Plain {\b bold} {\i italic} {\ul under} {\cf2 red} {\f1 \'cf\'f0\'e8\'e2\'e5\'f2} caf\'e9 \u8364? {\v hidden}\uc0\u8212 end.\par
{\listtext\pard\plain\f3 \'b7\tab}\pard\ls1\ilvl0 Bullet item\par
{\listtext\pard\plain 1.\tab}\pard\ls2\ilvl0 Numbered item\par
\pard {\field{\*\fldinst {HYPERLINK "https://example.com/"}}{\fldrslt {\ul Link text}}}\par
\trowd\clcbpat3\cellx3000\cellx6000\pard\intbl A1\cell B1\cell\row
\trowd\cellx3000\cellx6000\pard\intbl A2\cell B2\cell\row
\pard After{\super\chftn}{\footnote\pard\plain{\super\chftn} RTF_FOOTNOTE}\par
{\*\shppict{\pict\pngblip\picwgoal600\pichgoal600 ${PNG.toString('hex')}}}{\nonshppict{\pict\wmetafile8 0102}}\par
{\pict\emfblip 01000000}\par
Line one\line Line two\par
}`

test('.doc content is identified by its bytes, not its name (DOC-SIE-14)', () => {
  const ole = Buffer.alloc(512)
  OLE_COMPOUND_FILE_SIGNATURE.copy(ole)
  assert.equal(sniffLegacyDocContent(ole), 'ole')
  assert.equal(sniffLegacyDocContent(Buffer.from('PK\u0003\u0004rest')), 'docx')
  assert.equal(sniffLegacyDocContent(Buffer.from('{\\rtf1\\ansi x}')), 'rtf')
  assert.equal(sniffLegacyDocContent(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('\r\n  {\\rtf1 x}')])), 'rtf')
  assert.equal(sniffLegacyDocContent(Buffer.from('<!DOCTYPE html><html><body>x</body></html>')), 'html')
  assert.equal(sniffLegacyDocContent(Buffer.from('<html xmlns:o="urn:schemas-microsoft-com:office:office"><body>x')), 'html')
  assert.equal(sniffLegacyDocContent(Buffer.from('\ufeff<html><body>x')), 'html')
  assert.equal(sniffLegacyDocContent(Buffer.from('<?xml version="1.0"?><?mso-application progid="Word.Document"?><w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">')), 'word2003')
  assert.equal(sniffLegacyDocContent(Buffer.from('MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary="x"\r\n')), 'mht')
  assert.equal(sniffLegacyDocContent(Buffer.from('plain text that is not a document')), null)
})

test('an RTF document saved as .doc opens natively with its structure (DOC-SIE-14)', async () => {
  const result = await convertLegacyDocToDocx(Buffer.from(RTF, 'latin1'), { title: 'report.doc', ...NO_ENGINE })
  assert.equal(result.conversionMethod, 'rtf')
  assert.match(result.warnings[0], /contains a Rich Text \(RTF\) document.*leaves the original unchanged/)
  assert.match(result.warnings.join(' '), /WMF or EMF/)
  const doc = await importDocx(result.data)
  const all = paragraphs(doc)
  const heading = all.find((block) => text(block) === 'RTF Heading')
  assert.ok(heading, 'heading paragraph kept')
  assert.equal(heading.style.namedStyle, 'Heading1')
  const body = all.find((block) => text(block).startsWith('Plain '))
  // The space after a control word is its delimiter, as in Word.
  assert.equal(text(body), 'Plain bold italic under red Привет café € \u2014end.')
  assert.ok(body.runs.some((run) => run.text === 'bold' && run.style.bold))
  assert.ok(body.runs.some((run) => run.text === 'italic' && run.style.italic))
  assert.ok(body.runs.some((run) => run.text === 'red' && run.style.color === '#c00000'))
  assert.ok(!JSON.stringify(doc).includes('hidden'), 'hidden text stays hidden')
  const bullet = all.find((block) => text(block) === 'Bullet item')
  const numbered = all.find((block) => text(block) === 'Numbered item')
  assert.equal(doc.lists[bullet.style.list.listId].levels[0].format, 'bullet')
  assert.equal(doc.lists[numbered.style.list.listId].levels[0].format, 'decimal')
  assert.ok(all.some((block) => block.runs.some((run) => run.text === 'Link text' && run.style.link === 'https://example.com/')))
  const table = doc.blocks.find((block) => block.kind === 'table')
  assert.deepEqual(table.rows.map((row) => row.cells.map((cell) => text(cell.blocks[0]))), [['A1', 'B1'], ['A2', 'B2']])
  assert.equal(table.rows[0].cells[0].shading, '#ffff00')
  assert.equal(doc.blocks.filter((block) => block.kind === 'image').length, 1)
  assert.ok(JSON.stringify(doc).includes('RTF_FOOTNOTE'))
  assert.equal(text(doc.section.header[0]), 'RTF_HEADER')
  assert.ok(all.some((block) => text(block) === 'Line one\vLine two'))
})

test('a local Office engine is used for RTF only when one is installed', async () => {
  const calls = []
  const converted = await buildFlowDocx({ blocks: [{ type: 'paragraph', runs: [{ text: 'From the Office engine' }] }] })
  const result = await convertLegacyDocToDocx(Buffer.from(RTF, 'latin1'), {
    convertOffice: async (input) => { calls.push(input); return converted },
  })
  assert.equal(calls[0].inputExtension, 'rtf')
  assert.equal(calls[0].outputExtension, 'docx')
  assert.equal(result.data, converted)
})

test('an HTML page saved as .doc (including Word "Web Page" output) opens natively', async () => {
  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="windows-1252"><title>Web &amp; title</title><style>p{color:red}</style><script>alert(1)</script></head>
<body lang=EN-US><h1>Heading <em>one</em></h1><p class=MsoNormal>Para with <b>bold</b>, <i>italic</i>,<br>a break and <a href="https://example.com">a link</a> &amp;&nbsp;more<o:p></o:p></p>
<ul><li>Item one<li>Item two<ol><li>Nested</ol></ul>
<p class=MsoListParagraph style='mso-list:l0 level1 lfo1'><![if !supportLists]><span>1.<span>&nbsp;&nbsp;</span></span><![endif]>Word list item</p>
<table border=1><tr><th>H1<th>H2<tr><td style="background-color:#ffff00">c1<td>c2</table>
<p>Image <img src="data:image/png;base64,${PNG.toString('base64')}" width=20 height=10> and <img src="https://example.com/x.png" alt="remote"></p>
<pre>code line 1
  code line 2</pre><p onclick="evil()">Safe end</p></body></html>`
  const result = await convertLegacyDocToDocx(Buffer.from(html, 'latin1'), { title: 'web.doc' })
  assert.equal(result.conversionMethod, 'html')
  assert.match(result.warnings.join(' '), /not downloaded\. Simple never fetches content from the internet/)
  const doc = await importDocx(result.data)
  const all = paragraphs(doc)
  assert.equal(all[0].style.namedStyle, 'Heading1')
  assert.equal(text(all[0]), 'Heading one')
  assert.ok(all.some((block) => text(block) === 'Para with bold, italic,\va break and a link &\u00a0more'))
  assert.ok(all.some((block) => block.runs.some((run) => run.text === 'a link' && run.style.link === 'https://example.com/')))
  const item = all.find((block) => text(block) === 'Item one')
  const nested = all.find((block) => text(block) === 'Nested')
  assert.equal(doc.lists[item.style.list.listId].levels[0].format, 'bullet')
  assert.equal(nested.style.list.level, 1)
  const wordItem = all.find((block) => text(block) === 'Word list item')
  assert.equal(doc.lists[wordItem.style.list.listId].levels[0].format, 'decimal')
  const table = doc.blocks.find((block) => block.kind === 'table')
  assert.equal(table.rows.length, 2)
  assert.equal(table.rows[1].cells[0].shading, '#ffff00')
  assert.equal(doc.blocks.filter((block) => block.kind === 'image').length, 1)
  assert.ok(all.some((block) => text(block) === 'code line 1\v  code line 2'))
  const serialized = JSON.stringify(doc)
  assert.doesNotMatch(serialized, /alert\(1\)|color:red|evil/)
  assert.match(serialized, /Safe end/)
})

test('single-file web archives and Word 2003 XML saved as .doc open natively', async () => {
  const boundary = '----=_NextPart_01D'
  const mht = [
    'MIME-Version: 1.0',
    `Content-Type: multipart/related; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    'Content-Location: file:///C:/doc.htm',
    'Content-Transfer-Encoding: quoted-printable',
    'Content-Type: text/html; charset="utf-8"',
    '',
    '<html><body><p>MHT =D0=9F=D1=80=D0=B8=D0=B2=D0=B5=D1=82 paragraph</p><img src=3D"doc_files/image001.png"></body></html>',
    `--${boundary}`,
    'Content-Location: file:///C:/doc_files/image001.png',
    'Content-Transfer-Encoding: base64',
    'Content-Type: image/png',
    '',
    PNG.toString('base64'),
    `--${boundary}--`,
  ].join('\r\n')
  const archived = await convertLegacyDocToDocx(Buffer.from(mht, 'latin1'))
  assert.equal(archived.conversionMethod, 'mht')
  const archivedDoc = await importDocx(archived.data)
  assert.ok(paragraphs(archivedDoc).some((block) => text(block) === 'MHT Привет paragraph'))
  assert.equal(archivedDoc.blocks.filter((block) => block.kind === 'image').length, 1)

  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><?mso-application progid="Word.Document"?>
<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" xmlns:o="urn:schemas-microsoft-com:office:office">
<o:DocumentProperties><o:Title>W2003 title</o:Title></o:DocumentProperties>
<w:body><wx:sect><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Old XML heading</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:hlink w:dest="https://example.com/"><w:r><w:t>link</w:t></w:r></w:hlink></w:p>
<w:p><w:pPr><w:listPr><w:ilvl w:val="0"/><w:ilfo w:val="1"/><wx:t wx:val="1."/></w:listPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p><w:r><w:t>X1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Y1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
</wx:sect></w:body></w:wordDocument>`
  const legacyXml = await convertLegacyDocToDocx(Buffer.from(xml))
  assert.equal(legacyXml.conversionMethod, 'word2003')
  const xmlDoc = await importDocx(legacyXml.data)
  const all = paragraphs(xmlDoc)
  assert.equal(all[0].style.namedStyle, 'Heading1')
  assert.ok(all.some((block) => block.runs.some((run) => run.text === 'Bold' && run.style.bold)))
  assert.ok(all.some((block) => block.runs.some((run) => run.text === 'link' && run.style.link === 'https://example.com/')))
  const first = all.find((block) => text(block) === 'First')
  assert.equal(xmlDoc.lists[first.style.list.listId].levels[0].format, 'decimal')
  assert.ok(xmlDoc.blocks.some((block) => block.kind === 'table'))
})

test('a real binary .doc without an Office engine keeps the clear text-recovery warning', async () => {
  const ole = Buffer.alloc(512)
  OLE_COMPOUND_FILE_SIGNATURE.copy(ole)
  const result = await convertLegacyDocToDocx(ole, {
    Extractor: class { async extract() { return { getBody: () => 'Recovered body' } } },
  })
  assert.equal(result.conversionMethod, 'text')
  assert.match(result.warnings[0], /^Text-only import: .*only its text was recovered\. Saving creates a separate \.docx file and leaves the original \.doc unchanged\./)
  await assert.rejects(convertLegacyDocToDocx(Buffer.from('plain text that is not a document, long enough to be checked as a legacy file '.repeat(10))), /not a valid legacy \.doc/)
})

test('the RTF reader handles Unicode, code pages and malformed input safely', () => {
  const flow = parseRtf(String.raw`{\rtf1\ansi\ansicpg1251{\fonttbl{\f0 Arial;}}\f0 \'cf\'f0\'e8 \uc1\u-4064?\u1488?\'3f {\*\unknown skipped} \{braces\} \\back\par}`)
  // \uN's fallback character is skipped; a skipped destination leaves both spaces.
  assert.equal(flowText(flow.blocks), 'При \uf020א?  {braces} \\back')
  assert.throws(() => parseRtf('not rtf'), /not a Rich Text Format/)
  // Unbalanced and deeply nested input never throws past the depth guard.
  assert.doesNotThrow(() => parseRtf(`{\\rtf1 ${'{'.repeat(100)}x`))
  assert.throws(() => parseRtf(`{\\rtf1 ${'{'.repeat(600)}x`), /nested too deeply/)
})

test('flow documents become valid, readable DOCX packages', async () => {
  const flow = htmlToFlow('<p>Alpha</p><table><tr><td colspan=2>Wide</td></tr><tr><td>a</td><td><table><tr><td>inner</td></tr></table></td></tr></table>')
  const bytes = await buildFlowDocx(flow, { title: 'Flow <&> test' })
  const zip = await JSZip.loadAsync(bytes)
  assert.match(await zip.file('docProps/core.xml').async('string'), /Flow &lt;&amp;&gt; test/)
  const doc = await importDocx(bytes)
  const table = doc.blocks.find((block) => block.kind === 'table')
  assert.equal(table.rows[0].cells[0].colSpan, 2)
  assert.ok(JSON.stringify(doc).includes('inner'))
  assert.equal(mhtToFlow.length, 1)
  assert.equal(word2003ToFlow.length, 1)
})
