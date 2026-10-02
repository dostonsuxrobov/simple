const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const { pathToFileURL } = require('node:url')

// Clipboard paste (DOC-008, DOC-012, DOC-SIE-25). The TypeScript module loads through
// Node's type stripping; its transactions are applied with WordCanvas 0.12.0's own
// operation applier, so these checks see what the editor would do.
const load = () => import('../src/importers/paste.ts')

const DIST = path.join(__dirname, '..', 'node_modules', '@forevka', 'wordcanvas', 'dist-lib')
let applier = null
async function engineApply() {
  if (applier) return applier
  const file = fs.readdirSync(DIST).find((name) => /^paintStyle-.*\.js$/.test(name))
  const module = await import(pathToFileURL(path.join(DIST, file)).href)
  assert.equal(typeof module.j, 'function', 'WordCanvas 0.12.0 exports its operation applier from the paintStyle chunk')
  applier = (doc, ops) => ops.reduce((current, op) => module.j(current, op).doc, doc)
  return applier
}

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const SECTION = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } }
const CHAR = { fontFamily: 'Cambria', fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: '#000000' }
const paragraph = (id, text, style = {}, char = CHAR) => ({ kind: 'paragraph', id, revision: 0, runs: [{ text, style: { ...char } }], style: { align: 'left', lineHeight: 1.15, spaceBeforePx: 0, spaceAfterPx: 8, indentFirstLinePx: 0, indentLeftPx: 0, ...style } })
const sheet = () => ({ styles: [{ id: 'Normal', name: 'Normal', type: 'paragraph', char: { fontFamily: 'Cambria', fontSizePx: 16 }, para: {} }], defaultStyleId: 'Normal' })
const documentWith = (blocks, extra = {}) => ({ section: { ...SECTION }, blocks, stylesheet: sheet(), ...extra })
const caret = (blockId, offset) => ({ anchor: { blockId, offset }, focus: { blockId, offset } })
const range = (blockId, start, endBlockId, end) => ({ anchor: { blockId, offset: start }, focus: { blockId: endBlockId, offset: end } })
const text = (block) => (block.runs || []).map((run) => run.text).join('')
const outline = (doc) => doc.blocks.map((block) => block.kind === 'paragraph' ? `${block.style.list ? `[${doc.lists?.[block.style.list.listId]?.levels[0].format ?? 'missing'}] ` : ''}${block.style.namedStyle ? `{${block.style.namedStyle}} ` : ''}${text(block)}` : `<${block.kind}>`)
let counter = 0
const ids = () => `n${counter++}`

const HTML = `<html><head><style>.lead{color:#336699}</style></head><body><!--StartFragment-->
<h1>Pasted heading</h1><p class="lead">Lead <b>bold</b> <a href="https://example.com/a">link</a></p>
<ul><li>Bullet A</li><li>Bullet B<ul><li>Nested</li></ul></li></ul><ol><li>One</li><li>Two</li></ol>
<table><tr><th>H1</th><th>H2</th></tr><tr><td>C1</td><td>C2</td></tr></table>
<p><img src="data:image/png;base64,${PNG}" width="40" height="20"><img src="https://example.com/remote.png"></p>
<!--EndFragment--></body></html>`

test('a web address is recognized only on its own', async () => {
  const { pastedLink, looksLikeAddress } = await load()
  assert.equal(pastedLink(' https://example.com/a?b=1#c '), 'https://example.com/a?b=1#c')
  assert.equal(pastedLink('http://intranet/page'), 'http://intranet/page')
  assert.equal(pastedLink('www.example.org/x'), 'https://www.example.org/x')
  assert.equal(pastedLink('mailto:ana@example.com'), 'mailto:ana@example.com')
  for (const value of ['example.com', 'see https://example.com', 'https://example.com\nmore', 'javascript:alert(1)', 'file:///C:/x', 'ftp://host/x', '']) assert.equal(pastedLink(value), null, value)
  assert.equal(looksLikeAddress('https://old.example.com'), true)
  assert.equal(looksLikeAddress('ana@example.com'), true)
  assert.equal(looksLikeAddress(String.raw`C:\Reports\q3.docx`), true)
  assert.equal(looksLikeAddress('our pricing page'), false)
})

test('clipboard text splits into lines like Word', async () => {
  const { textLines } = await load()
  assert.deepEqual(textLines('a\r\nb\rc\nd'), ['a', 'b', 'c', 'd'])
  // Whole lines keep their final break, so the text after the caret stays on its own line.
  assert.deepEqual(textLines('one\r\n'), ['one', ''])
  assert.deepEqual(textLines('\n\n'), ['', '', ''])
  assert.deepEqual(textLines('tab\there\vsoft\u0007bell'), ['tab\there\vsoftbell'])
  assert.deepEqual(textLines('x\u2029y'), ['x', 'y'])
})

test('a browser\'s Copy image is recognized as a picture, not as text', async () => {
  const { imageOnlyHtml } = await load()
  assert.equal(imageOnlyHtml('<html><body><!--StartFragment--><img src="https://example.com/cat.png" alt="Cat"><!--EndFragment--></body></html>'), true)
  assert.equal(imageOnlyHtml('<meta charset="utf-8"><img src="data:image/png;base64,AAAA">'), true)
  assert.equal(imageOnlyHtml('<p>Caption <img src="x.png"></p>'), false)
  assert.equal(imageOnlyHtml('<p>No picture</p>'), false)
  assert.equal(imageOnlyHtml(''), false)
})

test('JPEG orientation is read from EXIF (both byte orders)', async () => {
  const { jpegOrientation } = await load()
  const exif = (little, value) => {
    const entry = little ? [0x12, 0x01, 3, 0, 1, 0, 0, 0, value, 0, 0, 0] : [0x01, 0x12, 0, 3, 0, 0, 0, 1, 0, value, 0, 0]
    const tiff = little ? [0x49, 0x49, 0x2a, 0, 8, 0, 0, 0, 1, 0, ...entry, 0, 0, 0, 0] : [0x4d, 0x4d, 0, 0x2a, 0, 0, 0, 8, 0, 1, ...entry, 0, 0, 0, 0]
    const payload = [0x45, 0x78, 0x69, 0x66, 0, 0, ...tiff]
    return new Uint8Array([0xff, 0xd8, 0xff, 0xe1, (payload.length + 2) >> 8, (payload.length + 2) & 0xff, ...payload, 0xff, 0xd9])
  }
  assert.equal(jpegOrientation(exif(true, 6)), 6)
  assert.equal(jpegOrientation(exif(false, 3)), 3)
  assert.equal(jpegOrientation(new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 0, 4, 0, 0, 0xff, 0xd9])), 1)
  assert.equal(jpegOrientation(Buffer.from(PNG, 'base64')), 1)
})

test('pasted HTML maps to real lists, a table, a heading, links and the embedded picture; nothing is fetched', async () => {
  const { fragmentFromClipboard, isStructured } = await load()
  const fragment = fragmentFromClipboard({ html: HTML, text: 'ignored' }, { base: { fontFamily: 'Cambria', fontSizePx: 16, color: '#222222' }, idPrefix: 'clip-', maxImageWidthPx: 500 })
  const kinds = fragment.blocks.map((block) => block.kind === 'paragraph' ? text(block) : `<${block.kind}>`)
  assert.deepEqual(kinds, ['Pasted heading', 'Lead bold link', 'Bullet A', 'Bullet B', 'Nested', 'One', 'Two', '<table>', '<image>'])
  assert.equal(isStructured(fragment.blocks), true)
  assert.equal(fragment.blocks[0].style.namedStyle, 'Heading1')
  assert.deepEqual(fragment.styles.map((style) => style.id), ['Heading1'])
  const lists = fragment.blocks.filter((block) => block.kind === 'paragraph' && block.style.list)
  assert.deepEqual(lists.map((block) => [text(block), block.style.list.level]), [['Bullet A', 0], ['Bullet B', 0], ['Nested', 1], ['One', 0], ['Two', 0]])
  for (const block of lists) assert.ok(fragment.lists[block.style.list.listId], 'every list the paste uses comes with it')
  assert.equal(fragment.lists[lists[0].style.list.listId].levels[0].format, 'bullet')
  assert.equal(fragment.lists[lists[3].style.list.listId].levels[0].format, 'decimal')
  // Unformatted text takes the destination's font; the source's own styling stays.
  const lead = fragment.blocks[1]
  assert.deepEqual([lead.runs[0].style.fontFamily, lead.runs[0].style.fontSizePx, lead.runs[0].style.color], ['Cambria', 16, '#336699'])
  assert.equal(lead.runs.find((run) => run.text === 'bold').style.bold, true)
  assert.equal(lead.runs.find((run) => run.text === 'link').style.link, 'https://example.com/a')
  assert.equal(fragment.blocks[0].runs[0].style.fontFamily, 'Cambria', 'headings follow the destination font')
  const table = fragment.blocks[7]
  assert.deepEqual(table.rows.map((row) => row.cells.map((cell) => text(cell.blocks[0]))), [['H1', 'H2'], ['C1', 'C2']])
  assert.match(fragment.blocks[8].src, /^data:image\/png;base64,/)
  // The remote picture is left out (never downloaded), with a note instead of placeholder text.
  assert.ok(!kinds.some((value) => /\[Picture/.test(value)))
  assert.match(fragment.warnings.join(' '), /stored on the web was not downloaded and was left out/)
  // Word puts its pictures in temporary files: they are left out, with a paste-worded note.
  const word = fragmentFromClipboard({ html: '<p class=MsoNormal>Logo <img width=40 height=20 src="file:///C:/Users/x/AppData/Local/Temp/msohtmlclip1/01/clip_image001.png"></p>' })
  assert.deepEqual(word.blocks.map(text), ['Logo'])
  assert.deepEqual(word.warnings, ['A picture in the copied content couldn’t be read and was left out.'])
})

test('clipboard markup without content gives no fragment; RTF is read when there is no HTML', async () => {
  const { fragmentFromClipboard } = await load()
  assert.equal(fragmentFromClipboard({ html: '<html><body><!--StartFragment--><p> </p><p></p><!--EndFragment--></body></html>' }), null)
  assert.equal(fragmentFromClipboard({ text: 'only text' }), null)
  const rtf = fragmentFromClipboard({ rtf: String.raw`{\rtf1\ansi{\fonttbl{\f0 Arial;}}\pard\b Bold\b0  plain\par\pard Second\par}` })
  assert.deepEqual(rtf.blocks.map(text), ['Bold plain', 'Second'])
  assert.equal(rtf.blocks[0].runs[0].style.bold, true)
})

test('Merge formatting keeps emphasis and structure and takes the destination font', async () => {
  const { fragmentFromClipboard, mergeFragment } = await load()
  const fragment = fragmentFromClipboard({ html: '<h2>Title</h2><p><span style="font-family:Georgia;font-size:24px;color:#c00000"><b>Big</b> <i>red</i> <a href="https://x.org">go</a></span></p>' })
  const merged = mergeFragment(fragment, { ...CHAR, highlightColor: '#ffff00', sdtPath: ['c1'] })
  const runs = merged.blocks[1].runs
  for (const run of runs) assert.deepEqual([run.style.fontFamily, run.style.fontSizePx, run.style.color, run.style.sdtPath], ['Cambria', 16, '#000000', undefined])
  assert.equal(runs.find((run) => run.text === 'Big').style.bold, true)
  assert.equal(runs.find((run) => run.text === 'red').style.italic, true)
  assert.equal(runs.find((run) => run.text === 'go').style.link, 'https://x.org')
  assert.equal(merged.blocks[0].style.namedStyle, 'Heading2', 'structure stays')
  assert.ok(merged.blocks[0].runs[0].style.fontSizePx > 16, 'a heading keeps its size')
})

test('structured content pastes as blocks, with its list and style definitions, in one transaction', async () => {
  const { fragmentFromClipboard, buildPasteTransaction } = await load()
  const apply = await engineApply()
  const fragment = fragmentFromClipboard({ html: HTML }, { idPrefix: 's1-' })
  const doc = documentWith([paragraph('p1', 'Hello world'), paragraph('p2', 'Second')])
  // In the middle of a line: the line splits around the blocks.
  const middle = buildPasteTransaction(doc, caret('p1', 5), { kind: 'fragment', fragment, mode: 'keep' }, ids)
  const result = apply(doc, middle.ops)
  assert.deepEqual(outline(result), ['Hello', '{Heading1} Pasted heading', 'Lead bold link', '[bullet] Bullet A', '[bullet] Bullet B', '[bullet] Nested', '[decimal] One', '[decimal] Two', '<table>', '<image>', ' world', 'Second'])
  assert.ok(result.stylesheet.styles.some((style) => style.id === 'Heading1'), 'the missing heading style is added')
  assert.equal(middle.selectionAfter.focus.blockId, result.blocks[10].id, 'the caret ends after the pasted blocks')
  // At the start of a line: the blocks go before it, no empty paragraph is left behind.
  const start = buildPasteTransaction(doc, caret('p2', 0), { kind: 'fragment', fragment, mode: 'keep' }, ids)
  const atStart = apply(doc, start.ops)
  assert.deepEqual(outline(atStart).slice(0, 2), ['Hello world', '{Heading1} Pasted heading'])
  assert.equal(outline(atStart).at(-1), 'Second')
  assert.deepEqual(start.selectionAfter.focus, { blockId: 'p2', offset: 0 })
  // Over a selection across paragraphs: the selection goes, then the blocks come.
  const over = buildPasteTransaction(doc, range('p1', 6, 'p2', 3), { kind: 'fragment', fragment, mode: 'keep' }, ids)
  const replaced = apply(doc, over.ops)
  assert.equal(outline(replaced)[0], 'Hello ')
  assert.equal(outline(replaced).at(-1), 'ond')
  assert.equal(replaced.blocks.filter((block) => block.kind === 'table').length, 1)
})

test('a document without a stylesheet keeps the engine\'s built-in headings and drops the others', async () => {
  const { fragmentFromClipboard, buildPasteTransaction } = await load()
  const apply = await engineApply()
  const fragment = fragmentFromClipboard({ html: '<h1>One</h1><h3>Three</h3><p>Body</p>' })
  const doc = { section: { ...SECTION }, blocks: [paragraph('p1', '')] }
  const transaction = buildPasteTransaction(doc, caret('p1', 0), { kind: 'fragment', fragment, mode: 'keep' }, ids)
  assert.ok(!transaction.ops.some((op) => op.type === 'setStylesheet'))
  const result = apply(doc, transaction.ops)
  assert.deepEqual(outline(result), ['{Heading1} One', 'Three', 'Body', ''])
  assert.equal(result.blocks[1].style.outlineLevel, 2, 'the third-level heading keeps its outline level')
})

test('one paragraph merges into the line; several merge at both ends like the engine\'s own paste', async () => {
  const { fragmentFromClipboard, buildPasteTransaction } = await load()
  const apply = await engineApply()
  const doc = documentWith([paragraph('p1', 'Hello world', { align: 'center' })])
  const one = fragmentFromClipboard({ html: '<p><b>bold</b> words</p>' })
  const inline = buildPasteTransaction(doc, caret('p1', 6), { kind: 'fragment', fragment: one, mode: 'keep' }, ids)
  assert.deepEqual(inline.ops.map((op) => op.type), ['insertRuns'])
  assert.equal(text(apply(doc, inline.ops).blocks[0]), 'Hello bold wordsworld')
  assert.deepEqual(inline.selectionAfter.focus, { blockId: 'p1', offset: 16 })
  const several = fragmentFromClipboard({ html: '<p>A</p><p style="text-align:right">B</p><p>C</p>' })
  const merge = buildPasteTransaction(doc, caret('p1', 6), { kind: 'fragment', fragment: several, mode: 'keep' }, ids)
  const merged = apply(doc, merge.ops)
  assert.deepEqual(merged.blocks.map(text), ['Hello A', 'B', 'Cworld'])
  assert.equal(merged.blocks[1].style.align, 'right', 'keep source formatting gives the middle paragraph its own style')
  assert.equal(merged.blocks[2].style.align, 'center', 'the last paragraph joins the destination line')
  assert.deepEqual(merge.selectionAfter.focus, { blockId: merged.blocks[2].id, offset: 1 })
  // Merge formatting: middle paragraphs keep the destination paragraph style.
  const mergeMode = apply(doc, buildPasteTransaction(doc, caret('p1', 6), { kind: 'fragment', fragment: several, mode: 'merge' }, ids).ops)
  assert.equal(mergeMode.blocks[1].style.align, 'center')
})

test('one list item or heading pasted on an empty line becomes that kind of paragraph', async () => {
  const { fragmentFromClipboard, buildPasteTransaction } = await load()
  const apply = await engineApply()
  const doc = documentWith([paragraph('p1', 'Intro'), paragraph('p2', '')])
  const item = fragmentFromClipboard({ html: '<ol><li>Only item</li></ol>' }, { idPrefix: 'li-' })
  const transaction = buildPasteTransaction(doc, caret('p2', 0), { kind: 'fragment', fragment: item, mode: 'keep' }, ids)
  const result = apply(doc, transaction.ops)
  assert.deepEqual(outline(result), ['Intro', '[decimal] Only item'])
  assert.deepEqual(transaction.selectionAfter.focus, { blockId: 'p2', offset: 9 })
  // In the middle of text it is just text.
  const inline = apply(doc, buildPasteTransaction(doc, caret('p1', 5), { kind: 'fragment', fragment: item, mode: 'keep' }, ids).ops)
  assert.deepEqual(outline(inline), ['IntroOnly item', ''])
})

test('plain text pastes in one transaction, line by line, formatted like the caret', async () => {
  const { buildPasteTransaction } = await load()
  const apply = await engineApply()
  const bold = { ...CHAR, bold: true }
  const doc = documentWith([paragraph('p1', 'Hello world', {}, bold), paragraph('p2', 'Second')])
  const transaction = buildPasteTransaction(doc, caret('p1', 6), { kind: 'text', text: 'one\ntwo\r\nthree' }, ids)
  const result = apply(doc, transaction.ops)
  assert.deepEqual(result.blocks.map(text), ['Hello one', 'two', 'threeworld', 'Second'])
  assert.ok(result.blocks.slice(0, 3).every((block) => block.runs.every((run) => run.style.bold)), 'the caret formatting continues')
  assert.deepEqual(transaction.selectionAfter.focus, { blockId: result.blocks[2].id, offset: 5 })
  // Over a selection within one paragraph.
  const replaced = apply(doc, buildPasteTransaction(doc, range('p1', 0, 'p1', 5), { kind: 'text', text: 'Howdy' }, ids).ops)
  assert.equal(text(replaced.blocks[0]), 'Howdy world')
  // Line breaks become paragraphs (two copied empty lines, two empty paragraphs); an
  // empty clipboard over a selection just deletes it.
  assert.deepEqual(apply(doc, buildPasteTransaction(doc, caret('p2', 0), { kind: 'text', text: '\n\n' }, ids).ops).blocks.map(text), ['Hello world', '', '', 'Second'])
  assert.deepEqual(apply(doc, buildPasteTransaction(doc, caret('p2', 3), { kind: 'text', text: 'whole line\r\n' }, ids).ops).blocks.map(text), ['Hello world', 'Secwhole line', 'ond'])
  const deleteOnly = buildPasteTransaction(doc, range('p1', 5, 'p2', 0), { kind: 'text', text: '' }, ids)
  assert.deepEqual(apply(doc, deleteOnly.ops).blocks.map(text), ['HelloSecond'])
  assert.equal(buildPasteTransaction(doc, caret('p1', 0), { kind: 'text', text: '' }, ids), null)
  // A web address pasted at a caret is a link with the caret formatting.
  const linked = buildPasteTransaction(doc, caret('p2', 6), { kind: 'text', text: 'https://example.com', link: 'https://example.com' }, ids)
  const run = apply(doc, linked.ops).blocks[1].runs.find((item) => item.text === 'https://example.com')
  assert.deepEqual([run.style.link, run.style.fontFamily], ['https://example.com', 'Cambria'])
})

test('a web address over selected text links the text, also across paragraphs', async () => {
  const { buildPasteTransaction } = await load()
  const apply = await engineApply()
  const doc = documentWith([paragraph('p1', 'Read our pricing page'), paragraph('p2', 'and the FAQ')])
  const one = apply(doc, buildPasteTransaction(doc, range('p1', 9, 'p1', 21), { kind: 'link', url: 'https://example.com/pricing' }, ids).ops)
  assert.deepEqual(one.blocks[0].runs.map((run) => [run.text, run.style.link ?? null]), [['Read our ', null], ['pricing page', 'https://example.com/pricing']])
  const both = apply(doc, buildPasteTransaction(doc, range('p1', 9, 'p2', 3), { kind: 'link', url: 'https://example.com' }, ids).ops)
  assert.equal(both.blocks[1].runs[0].style.link, 'https://example.com')
  assert.equal(both.blocks[1].runs[0].text, 'and')
  assert.equal(buildPasteTransaction(doc, caret('p1', 2), { kind: 'link', url: 'https://example.com' }, ids), null, 'nothing selected: nothing to link')
  assert.equal(buildPasteTransaction(doc, range('p1', 0, 'p1', 4), { kind: 'link', url: 'javascript:alert(1)' }, ids), null)
})

test('a table cell gets paragraphs only: lists as text, table rows as tab-separated lines, pictures noted', async () => {
  const { fragmentFromClipboard, buildPasteTransaction } = await load()
  const apply = await engineApply()
  const cell = (id, value) => ({ id: `c-${id}`, blocks: [paragraph(id, value)] })
  const doc = documentWith([{ kind: 'table', id: 't1', revision: 0, rows: [{ cells: [cell('a', 'Cell text'), cell('b', '')] }] }, paragraph('p9', '')])
  const fragment = fragmentFromClipboard({ html: HTML }, { idPrefix: 'cell-' })
  const transaction = buildPasteTransaction(doc, caret('a', 4), { kind: 'fragment', fragment, mode: 'keep' }, ids)
  assert.ok(!transaction.ops.some((op) => op.type === 'insertBlock' || op.type === 'setListDefinition'))
  const result = apply(doc, transaction.ops)
  const texts = result.blocks[0].rows[0].cells[0].blocks.map(text)
  assert.deepEqual(texts, ['CellPasted heading', 'Lead bold link', '•\tBullet A', '•\tBullet B', '◦\tNested', '1.\tOne', '2.\tTwo', 'H1\tH2', 'C1\tC2 text'])
  assert.match(transaction.warnings.join(' '), /picture wasn't pasted/)
})

test('places the module cannot handle are left to the engine', async () => {
  const { buildPasteTransaction, fragmentFromClipboard } = await load()
  const fragment = fragmentFromClipboard({ html: '<p>x</p>' })
  const control = paragraph('p1', 'Locked')
  control.sdtPath = ['sdt1']
  const doc = documentWith([control, paragraph('p2', 'Free')], { section: { ...SECTION, header: [paragraph('h1', 'Header text')] } })
  assert.equal(buildPasteTransaction(doc, caret('p1', 2), { kind: 'fragment', fragment, mode: 'keep' }, ids), null, 'content controls keep the engine\'s locks')
  assert.equal(buildPasteTransaction(doc, range('p2', 0, 'h1', 2), { kind: 'text', text: 'x' }, ids), null, 'a selection across stories')
  assert.equal(buildPasteTransaction(doc, null, { kind: 'text', text: 'x' }, ids), null)
  assert.equal(buildPasteTransaction(doc, caret('missing', 0), { kind: 'text', text: 'x' }, ids), null)
  // A header takes blocks too.
  const apply = await engineApply()
  const blocks = fragmentFromClipboard({ html: '<ul><li>a</li><li>b</li></ul>' }, { idPrefix: 'hd-' })
  const header = apply(doc, buildPasteTransaction(doc, caret('h1', 0), { kind: 'fragment', fragment: blocks, mode: 'keep' }, ids).ops)
  assert.deepEqual(header.section.header.map(text), ['a', 'b', 'Header text'])
})
