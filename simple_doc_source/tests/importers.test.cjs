const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const JSZip = require('jszip')

// Native importers (DOC-017, DOC-SIE-16): .txt, .md, .html, .rtf and .odt open
// without LibreOffice. The TypeScript modules load through Node's type stripping.
const load = () => import('../src/importers/index.ts')

const EXAMPLES = path.join(__dirname, '..', '..', 'Simple test examples')
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')

const text = (block) => (block.runs || []).map((run) => run.text).join('')
const paragraphs = (blocks) => blocks.filter((block) => block.kind === 'paragraph')
const find = (blocks, pattern) => paragraphs(blocks).find((block) => pattern.test(text(block)))
const headings = (doc) => paragraphs(doc.blocks).filter((block) => /^Heading\d$/.test(block.style.namedStyle || ''))
const tables = (doc) => doc.blocks.filter((block) => block.kind === 'table')
const images = (doc) => doc.blocks.filter((block) => block.kind === 'image')
const listParagraphs = (doc) => paragraphs(doc.blocks).filter((block) => block.style.list)
const cellText = (cell) => cell.blocks.map(text).join('\n')
const runWith = (blocks, pattern) => paragraphs(blocks).flatMap((block) => block.runs).find((run) => pattern.test(run.text))

/** Reads a copy of a shared example file, so the originals are never touched. */
function exampleCopy(name, t) {
  const source = path.join(EXAMPLES, name)
  if (!fs.existsSync(source)) {
    t.skip(`${name} is not available`)
    return null
  }
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-importers-'))
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }))
  const copy = path.join(directory, name)
  fs.copyFileSync(source, copy)
  return fs.readFileSync(copy)
}

/** Every list reference and footnote reference resolves inside the document. */
function assertSelfContained(doc) {
  const visit = (blocks) => {
    for (const block of blocks) {
      if (block.kind === 'paragraph') {
        assert.ok(block.runs.length > 0, 'a paragraph always has a run')
        if (block.style.list) assert.ok(doc.lists?.[block.style.list.listId], `list ${block.style.list.listId} is defined`)
        for (const run of block.runs) if (run.style.footnoteRef) assert.ok(doc.footnotes?.[run.style.footnoteRef], 'the footnote body exists')
      }
      if (block.kind === 'table') for (const row of block.rows) for (const cell of row.cells) {
        assert.ok(cell.blocks.length > 0, 'a cell always has content')
        visit(cell.blocks)
      }
    }
  }
  visit(doc.blocks)
  const ids = new Set()
  const collect = (blocks) => {
    for (const block of blocks) {
      assert.ok(!ids.has(block.id), `block id ${block.id} is unique`)
      ids.add(block.id)
      if (block.kind === 'table') for (const row of block.rows) for (const cell of row.cells) collect(cell.blocks)
    }
  }
  collect(doc.blocks)
  assert.equal(doc.blocks[doc.blocks.length - 1].kind, 'paragraph', 'the document ends with a paragraph')
}

/** The engine's own DOCX writer and reader accept the imported model. */
async function engineRoundTrip(doc) {
  const { runExport } = await import('@forevka/wordcanvas/export')
  const { runImport } = await import('@forevka/wordcanvas/import')
  const pictures = {}
  const collect = (blocks) => {
    for (const block of blocks) {
      if (block.kind === 'image' && block.src.startsWith('data:')) pictures[block.src] = Buffer.from(block.src.split(',')[1], 'base64')
      if (block.kind === 'table') for (const row of block.rows) for (const cell of row.cells) collect(cell.blocks)
    }
  }
  collect(doc.blocks)
  const { bytes } = await runExport(doc, 'docx', pictures)
  return runImport(bytes, undefined, { collectMediaBytes: true }).doc
}

function utf16be(value) {
  const bytes = Buffer.from(value, 'utf16le')
  for (let i = 0; i < bytes.length; i += 2) [bytes[i], bytes[i + 1]] = [bytes[i + 1], bytes[i]]
  return bytes
}

function cp1252(value) {
  const special = { '€': 0x80, '‚': 0x82, '„': 0x84, '…': 0x85, '‘': 0x91, '’': 0x92, '“': 0x93, '”': 0x94, '–': 0x96, '—': 0x97, '™': 0x99 }
  return Buffer.from(Array.from(value, (ch) => special[ch] ?? ch.charCodeAt(0)))
}

function cp1251(value) {
  return Buffer.from(Array.from(value, (ch) => {
    const code = ch.charCodeAt(0)
    if (code >= 0x410 && code <= 0x44f) return code - 0x410 + 0xc0
    if (ch === 'Ё') return 0xa8
    if (ch === 'ё') return 0xb8
    return code
  }))
}

// ---- Markdown -----------------------------------------------------------------------

test('Complex document.md imports with headings, tables, lists, picture and line breaks', async (t) => {
  const bytes = exampleCopy('Complex document.md', t)
  if (!bytes) return
  const { importDocument } = await load()
  const result = await importDocument(bytes, { fileName: 'Complex document.md', idPrefix: 'md-' })
  const doc = result.document
  assert.equal(result.format, 'md')
  assert.equal(result.encoding, 'utf-8')
  assert.deepEqual(headings(doc).map((block) => [block.style.namedStyle, text(block)]), [
    ['Heading1', 'created'],
    ['Heading2', 'Section 2 first-page header'],
    ['Heading2', 'Section 2 first-page footer'],
    ['Heading2', 'Footnote 1'],
  ])
  assert.equal(headings(doc)[0].style.outlineLevel, 0)
  const [first, second] = tables(doc)
  assert.equal(tables(doc).length, 2)
  assert.deepEqual([first.rows.length, first.rows[0].cells.length, second.rows.length, second.rows[0].cells.length], [3, 3, 2, 2])
  assert.equal(cellText(first.rows[0].cells[0]), 'TABLE_ONE_1 value')
  assert.ok(first.rows[0].cells[0].blocks[0].runs[0].style.bold, 'header cells are bold')
  assert.match(cellText(first.rows[1].cells[1]), /^TABLE_ONE_5 Long wrapped cell content\./)
  assert.equal(cellText(second.rows[1].cells[1]), 'TABLE_TWO_4')
  const bullets = listParagraphs(doc)
  assert.deepEqual(bullets.map(text), ['BULLET_ONE', 'BULLET_TWO'])
  assert.equal(bullets[0].style.list.listId, bullets[1].style.list.listId)
  assert.equal(doc.lists[bullets[0].style.list.listId].levels[0].format, 'bullet')
  const pictures = images(doc)
  assert.equal(pictures.length, 1)
  assert.match(pictures[0].src, /^data:image\/png;base64,/)
  assert.deepEqual([pictures[0].widthPx, pictures[0].heightPx], [320, 140])
  const bold = runWith(doc.blocks, /STRESS_START/)
  assert.ok(bold.style.bold)
  assert.ok(runWith(doc.blocks, /italic accent café naïve Ångström/).style.italic)
  assert.equal(text(find(doc.blocks, /^STRESS_END/)), 'STRESS_END — final retained sentence.\vSOFT_BREAK_LINE — same paragraph.1')
  assert.ok(find(doc.blocks, /^RTL العربية 123 نهاية — עברית שלום 456 סוף$/))
  assert.ok(find(doc.blocks, /^CJK 中文测试 日本語の文章 한국어 문장$/))
  assert.ok(find(doc.blocks, /^FOOTNOTE_CONTENT — survives each export\.$/))
  assert.equal(paragraphs(doc.blocks).filter((block) => /^PARA_\d\d /.test(text(block))).length, 16)
  assert.deepEqual(result.warnings, [])
  assertSelfContained(doc)

  const back = await engineRoundTrip(doc)
  assert.equal(back.blocks.filter((block) => block.kind === 'table').length, 2)
  assert.equal(back.blocks.filter((block) => block.kind === 'image').length, 1, 'the picture survives a DOCX save')
  assert.equal(back.blocks.filter((block) => block.style?.list).length, 2)
})

test('Markdown: nested and continued lists, code, quotes, task lists, footnotes, links and front matter', async () => {
  const { importMarkdown } = await load()
  const source = [
    '---', 'title: "Plan"', 'author: me', '---',
    '# One', '', '## Two', '', '###### Six', '',
    '1. first', '2. second', '   - inner *em*', '     1. deep', '3. third', '',
    'Interruption.', '',
    '4. fourth', '',
    '- [ ] open task', '- [x] done task', '',
    '```js', 'const a = 1;', '', 'call();', '```', '',
    '> quoted **text**', '',
    'Text with a note[^n] and `code` and ~~gone~~ and <https://example.com/x> and [bad](javascript:alert(1)).', '',
    '| Left | Right |', '|:-----|------:|', '| a | b |', '',
    '![remote](https://example.com/a.png) ![local](pictures/b.png)', '',
    '[^n]: The **note** body.',
    '    Second line of the note.',
  ].join('\n')
  const result = importMarkdown(Buffer.from(source), { idPrefix: 'm-' })
  const doc = result.document
  assert.equal(result.title, 'Plan')
  assert.deepEqual(headings(doc).map((block) => block.style.namedStyle), ['Heading1', 'Heading2', 'Heading6'])
  const items = listParagraphs(doc)
  const ordered = items.filter((block) => /first|second|inner|deep|third|fourth/.test(text(block)))
  assert.deepEqual(ordered.map((block) => [text(block), block.style.list.level]), [['first', 0], ['second', 0], ['inner em', 1], ['deep', 2], ['third', 0], ['fourth', 0]])
  const listId = ordered[0].style.list.listId
  assert.ok(ordered.every((block) => block.style.list.listId === listId), '"4." after a paragraph continues the same numbered list')
  const levels = doc.lists[listId].levels
  // Like Word and Google Docs, a third numbered level counts i, ii, iii.
  assert.deepEqual([levels[0].format, levels[1].format, levels[2].format], ['decimal', 'bullet', 'lowerRoman'])
  assert.equal(levels[0].text, '%1.')
  const tasks = items.filter((block) => /task/.test(text(block)))
  assert.deepEqual(tasks.map(text), ['☐ open task', '☒ done task'])
  assert.notEqual(tasks[0].style.list.listId, listId)
  const code = ['const a = 1;', '', 'call();'].map((line) => paragraphs(doc.blocks).find((block) => text(block) === line && block.runs[0].style.fontFamily === 'Consolas'))
  assert.ok(code.every(Boolean), 'each code line is a monospace paragraph, blank lines kept')
  const quote = find(doc.blocks, /^quoted text$/)
  assert.ok(quote.style.indentLeftPx >= 48)
  const sentence = find(doc.blocks, /^Text with a note/)
  const note = sentence.runs.find((run) => run.style.footnoteRef)
  assert.equal(note.text, '1')
  assert.equal(note.style.verticalAlign, 'super')
  assert.equal(doc.footnotes[note.style.footnoteRef].map(text).join(' '), 'The note body. Second line of the note.')
  assert.equal(sentence.runs.find((run) => run.text === 'code').style.fontFamily, 'Consolas')
  assert.ok(sentence.runs.find((run) => run.text === 'gone').style.strikethrough)
  assert.equal(sentence.runs.find((run) => run.text === 'https://example.com/x').style.link, 'https://example.com/x')
  assert.ok(!sentence.runs.some((run) => run.style.link && /javascript/i.test(run.style.link)), 'script links are never kept')
  const [table] = tables(doc)
  assert.equal(table.rows[1].cells[1].blocks[0].style.align, 'right')
  assert.equal(text(find(doc.blocks, /\[Picture: remote\]/)), '[Picture: remote] [Picture: local]')
  assert.equal(result.warnings.length, 2)
  assert.match(result.warnings.join(' '), /stored on the web was not downloaded/)
  assert.match(result.warnings.join(' '), /stored outside the file was not found/)
  assertSelfContained(doc)
  const back = await engineRoundTrip(doc)
  assert.deepEqual(Object.values(back.footnotes || {}).map((note) => note.map(text).join(' ')), ['The note body. Second line of the note.'])
})

test('Markdown pictures next to the file load through the caller, never the network', async () => {
  const { importMarkdown } = await load()
  const requested = []
  const result = importMarkdown('![logo](images/logo.png)\n', { resolveImage: (source) => { requested.push(source); return source === 'images/logo.png' ? { bytes: PNG } : null } })
  assert.deepEqual(requested, ['images/logo.png'])
  const [picture] = images(result.document)
  assert.deepEqual([picture.widthPx, picture.heightPx], [1, 1])
  assert.deepEqual(result.warnings, [])
})

// ---- HTML ----------------------------------------------------------------------------

test('Complex document.html imports with headings, tables, lists, picture and its page break', async (t) => {
  const bytes = exampleCopy('Complex document.html', t)
  if (!bytes) return
  const { importDocument } = await load()
  const result = await importDocument(bytes, { fileName: 'Complex document.html', idPrefix: 'html-' })
  const doc = result.document
  assert.equal(result.format, 'html')
  assert.equal(result.title, 'created')
  assert.deepEqual(headings(doc).map(text), ['Section 2 first-page header', 'Section 2 first-page footer', 'Footnote 1'])
  assert.equal(tables(doc).length, 2)
  assert.deepEqual(tables(doc).map((table) => [table.rows.length, table.rows[0].cells.length]), [[3, 3], [2, 2]])
  // Simple's own HTML writes list markers as text; they become a real list again.
  const bullets = listParagraphs(doc)
  assert.deepEqual(bullets.map(text), ['BULLET_ONE', 'BULLET_TWO'])
  assert.equal(doc.lists[bullets[0].style.list.listId].levels[0].bulletChar, '•')
  const pictures = images(doc)
  assert.equal(pictures.length, 1)
  assert.deepEqual([pictures[0].widthPx, pictures[0].heightPx], [320, 140])
  const landscape = find(doc.blocks, /^LANDSCAPE_MARKER$/)
  assert.equal(landscape.style.pageBreakBefore, true, 'the second section starts on a new page')
  const mixed = find(doc.blocks, /^Mixed Arial blue/)
  assert.deepEqual([mixed.runs[0].style.fontFamily, mixed.runs[0].style.color, mixed.runs[0].style.fontSizePx], ['Arial', '#003388', 18])
  assert.equal(runWith(doc.blocks, /^bold text$/).style.bold, true)
  assert.equal(runWith(doc.blocks, /^STRESS_START/).style.fontSizePx, 32)
  assert.match(text(find(doc.blocks, /^STRESS_END/)), /^STRESS_END — final retained sentence\.\vSOFT_BREAK_LINE — same paragraph\.1$/)
  assert.ok(find(doc.blocks, /^CJK 中文测试/))
  assertSelfContained(doc)
  const back = await engineRoundTrip(doc)
  assert.equal(back.blocks.filter((block) => block.kind === 'table').length, 2)
})

test('HTML import is inert: no scripts, frames, handlers, script links or remote fetches', async () => {
  const { importHtml } = await load()
  const html = `<!doctype html><html><head><title>Unsafe page</title><script>alert("x")</script>
    <style>.red{color:red} p{margin:40px} .gap{margin-bottom:20px}</style><base href="https://example.com/"></head>
    <body onload="steal()"><p class="red gap" onclick="steal()">Red <a href="javascript:alert(1)">bad link</a> <a href=" JaVaScRiPt:alert(2)">bad two</a> <a href="https://ok.example/path?a=1&amp;b=2">good</a></p>
    <img src="https://example.com/track.png" alt="Tracker"><img src="file:///C:/secret.png"><img src="//cdn.example/x.png">
    <iframe src="https://evil.example/"></iframe><object data="x.swf">fallback</object><noscript>no script text</noscript>
    <p style="display:none">hidden words</p><p hidden>also hidden</p><template><p>template words</p></template>
    <p>Data <img src="data:image/png;base64,${PNG.toString('base64')}" width="20"> <img src="data:text/html;base64,PHNjcmlwdD4="></p>
    <svg><script>alert(3)</script><text>svg words</text></svg></body></html>`
  const result = importHtml(Buffer.from(html), { idPrefix: 'h-' })
  const doc = result.document
  const all = JSON.stringify(doc)
  for (const word of ['alert', 'steal', 'evil', 'hidden words', 'also hidden', 'template words', 'svg words', 'no script text', 'fallback', 'secret']) assert.ok(!all.includes(word), `${word} is not imported`)
  assert.equal(result.title, 'Unsafe page')
  const first = find(doc.blocks, /^Red/)
  assert.equal(first.runs[0].style.color, '#ff0000', 'class rules apply')
  assert.equal(first.style.spaceAfterPx, 20, 'class margins apply')
  assert.equal(first.style.spaceBeforePx, 0, 'bare tag rules do not set spacing')
  assert.equal(text(first), 'Red bad link bad two good', 'script links keep their text')
  assert.deepEqual(first.runs.filter((run) => run.style.link).map((run) => [run.text, run.style.link]), [['good', 'https://ok.example/path?a=1&b=2']])
  assert.equal(text(find(doc.blocks, /Tracker/)), '[Picture: Tracker] [Picture] [Picture]')
  const [picture] = images(doc)
  assert.deepEqual([picture.widthPx, picture.heightPx], [20, 20])
  assert.match(result.warnings.join('\n'), /2 pictures stored on the web were not downloaded/)
  assert.match(result.warnings.join('\n'), /stored outside the file was not found/)
  assert.match(result.warnings.join('\n'), /format Simple cannot show/)
  assert.match(result.warnings.join('\n'), /Embedded videos, frames and other active content were left out/)
})

test('HTML: lists, merged table cells, headings, inline formatting and page breaks', async () => {
  const { importHtml } = await load()
  const html = `<meta charset="windows-1252"><h1>Main</h1><h3 style="color:#00f">Third</h3>
    <p>A <b>b</b> <i>i</i> <u>u</u> <s>s</s> <mark>m</mark> <span style="background-color:#0f0;font-size:14pt;font-family:'Times New Roman',serif">styled</span> H<sub>2</sub>O x<sup>2</sup> <font color="red" face="Courier New">font</font> caf\u00e9</p>
    <ol type="a" start="3"><li>c item<ul><li>bullet</li></ul></li><li><p>para one</p><p>para two</p></li></ol>
    <table border="1"><colgroup><col width="100"><col width="300"></colgroup>
      <tr><th colspan="2" bgcolor="#eeeeee">Head</th></tr><tr><td rowspan="2">Tall</td><td>R1</td></tr><tr><td style="text-align:center">R2</td></tr></table>
    <div style="page-break-before:always">After break</div><p style="text-align:justify;text-indent:2em;margin-left:1in;line-height:150%">Body</p><hr><pre>  code
  more</pre>`
  const result = importHtml(cp1252(html), { idPrefix: 'x-' })
  const doc = result.document
  assert.equal(result.encoding, 'windows-1252')
  const [h1, h3] = headings(doc)
  assert.deepEqual([h1.style.namedStyle, h3.style.namedStyle, h3.style.outlineLevel, h3.runs[0].style.color], ['Heading1', 'Heading3', 2, '#0000ff'])
  const line = find(doc.blocks, /^A b i u s m styled H2O x2 font café$/)
  const run = (value) => line.runs.find((item) => item.text === value).style
  assert.ok(run('b').bold && run('i').italic && run('u').underline && run('s').strikethrough)
  assert.equal(run('m').highlightColor, '#ffff00')
  assert.deepEqual([run('styled').highlightColor, run('styled').fontSizePx, run('styled').fontFamily], ['#00ff00', 18.667, 'Times New Roman'])
  assert.equal(run('2').verticalAlign, 'sub')
  assert.deepEqual([run('font').color, run('font').fontFamily], ['#ff0000', 'Courier New'])
  const items = listParagraphs(doc)
  assert.deepEqual(items.map((block) => [text(block), block.style.list.level]), [['c item', 0], ['bullet', 1], ['para one', 0]])
  const definition = doc.lists[items[0].style.list.listId]
  assert.deepEqual([definition.levels[0].format, definition.levels[0].start, definition.levels[1].format], ['lowerLetter', 3, 'bullet'])
  const second = find(doc.blocks, /^para two$/)
  assert.equal(second.style.list, undefined, 'a second paragraph in an item has no marker')
  assert.equal(second.style.indentLeftPx, 48, 'it lines up with the item text')
  const [table] = tables(doc)
  assert.deepEqual(table.rows.map((row) => row.cells.map((cell) => [cellText(cell), cell.colSpan ?? 1, cell.rowSpan ?? 1])), [[['Head', 2, 1]], [['Tall', 1, 2], ['R1', 1, 1]], [['R2', 1, 1]]])
  assert.equal(table.rows[0].cells[0].shading, '#eeeeee')
  assert.deepEqual(table.colFractions, [0.25, 0.75])
  assert.equal(table.rows[2].cells[0].blocks[0].style.align, 'center')
  assert.equal(find(doc.blocks, /^After break$/).style.pageBreakBefore, true)
  const body = find(doc.blocks, /^Body$/)
  assert.deepEqual([body.style.align, body.style.indentFirstLinePx, body.style.indentLeftPx, body.style.lineHeight], ['justify', 29.33, 96, 1.5])
  assert.ok(paragraphs(doc.blocks).some((block) => block.style.borders?.bottom), 'a horizontal rule is a bottom border')
  assert.deepEqual(['  code', '  more'].map((value) => !!paragraphs(doc.blocks).find((block) => text(block) === value)), [true, true])
  assertSelfContained(doc)
  await engineRoundTrip(doc)
})

test('HTML written by Simple itself and by Word comes back with real lists and footnotes', async () => {
  const { importHtml, importMarkdown } = await load()
  const { documentToHtml } = await import('../src/document-export.js')
  const original = importMarkdown('# Report\n\n1. one\n2. two\n\nGap.\n\n- dot\n\nText[^a].\n\n[^a]: Note text.\n', { idPrefix: 's-' }).document
  const again = importHtml(documentToHtml(original, 'Report'), { idPrefix: 'r-' }).document
  assert.deepEqual(headings(again).map(text), ['Report'])
  assert.deepEqual(listParagraphs(again).map((block) => text(block)), ['one', 'two', 'dot'])
  const [one, two, dot] = listParagraphs(again)
  assert.equal(one.style.list.listId, two.style.list.listId)
  assert.notEqual(one.style.list.listId, dot.style.list.listId)
  assert.equal(again.lists[one.style.list.listId].levels[0].format, 'decimal')
  const note = find(again.blocks, /^Text/).runs.find((run) => run.style.footnoteRef)
  assert.ok(note, 'the note reference is a real footnote again')
  assert.equal(again.footnotes[note.style.footnoteRef].map(text).join(''), 'Note text.')
  assert.ok(!find(again.blocks, /Note text/), 'the note body is not repeated in the text')

  const word = `<html xmlns:o="urn:schemas-microsoft-com:office:office"><head><style>
    p.MsoNormal{margin:0in;font-size:11.0pt;font-family:"Calibri",sans-serif}
    p.MsoListParagraph{margin-left:.5in}</style></head><body>
    <p class=MsoNormal>Intro<a style='mso-footnote-id:ftn1' href="#_ftn1" name="_ftnref1" title=""><span class=MsoFootnoteReference>[1]</span></a></p>
    <p class=MsoListParagraph style='text-indent:-.25in;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp; </span></span></span><![endif]>Word bullet</p>
    <p class=MsoListParagraph style='margin-left:1.0in;text-indent:-.25in;mso-list:l0 level2 lfo1'><span style='mso-list:Ignore'>o<span>&nbsp;</span></span>Word sub</p>
    <p class=MsoNormal><![if !vml]><img width=2 height=2 src="data:image/png;base64,${PNG.toString('base64')}"><![endif]></p>
    <p class=MsoNormal>Outro<o:p></o:p></p>
    <div style='mso-element:footnote-list'><div style='mso-element:footnote' id=ftn1><p class=MsoFootnoteText><a style='mso-footnote-id:ftn1' href="#_ftnref1" name="_ftn1" title=""><span class=MsoFootnoteReference>[1]</span></a> Word note.</p></div></div></body></html>`
  const fromWord = importHtml(word, { idPrefix: 'w-' }).document
  const [bullet, sub] = listParagraphs(fromWord)
  assert.deepEqual([text(bullet), bullet.style.list.level, text(sub), sub.style.list.level], ['Word bullet', 0, 'Word sub', 1])
  assert.equal(bullet.style.list.listId, sub.style.list.listId)
  assert.equal(bullet.style.indentLeftPx, 0, 'Word list indents come from the list, not the paragraph')
  assert.equal(fromWord.lists[bullet.style.list.listId].levels[1].bulletChar, '◦')
  const intro = find(fromWord.blocks, /^Intro/)
  assert.equal(text(intro), 'Intro1')
  assert.equal(fromWord.footnotes[intro.runs[1].style.footnoteRef].map(text).join(''), 'Word note.')
  assert.equal(intro.runs[0].style.fontSizePx, 14.667)
  assert.ok(find(fromWord.blocks, /^Outro$/))
  assert.equal(images(fromWord).length, 1, 'the picture Word writes for browsers without VML is kept')
  assertSelfContained(fromWord)
})

// ---- Plain text ----------------------------------------------------------------------

test('plain text: BOMs, UTF-16 without a BOM, UTF-8, Windows-1252 and Windows-1251', async () => {
  const { importText, importDocument } = await load()
  const lines = (result) => result.document.blocks.map(text)
  const sample = 'Hello café\r\nSecond line\r\n'
  const le = importText(Buffer.from(sample, 'utf16le'))
  assert.equal(le.encoding, 'utf-16le')
  assert.deepEqual(lines(le), ['Hello café', 'Second line'])
  const be = importText(Buffer.concat([Buffer.from([0xfe, 0xff]), utf16be(sample)]))
  assert.equal(be.encoding, 'utf-16be')
  assert.deepEqual(lines(be), ['Hello café', 'Second line'])
  const beNoBom = importText(utf16be('Plain ASCII text in UTF-16 big endian.\n'))
  assert.equal(beNoBom.encoding, 'utf-16be')
  const bom8 = importText(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('Ünïcode ✓\n')]))
  assert.deepEqual([bom8.encoding, ...lines(bom8)], ['utf-8', 'Ünïcode ✓'])
  const western = importText(cp1252('Café naïve – “quoted” €5 Ångström\r\nLine 2'))
  assert.equal(western.encoding, 'windows-1252')
  assert.deepEqual(lines(western), ['Café naïve – “quoted” €5 Ångström', 'Line 2'])
  const cyrillic = importText(cp1251('Привет, мир! Съешь ещё этих мягких булок.\nВторая строка'))
  assert.equal(cyrillic.encoding, 'windows-1251')
  assert.deepEqual(lines(cyrillic), ['Привет, мир! Съешь ещё этих мягких булок.', 'Вторая строка'])
  const forced = importText(cp1251('Да'), { encoding: 'windows-1252' })
  assert.equal(lines(forced)[0], 'Äà')
  const paged = await importDocument(Buffer.from('Page one\n\fPage two\twith tab\n\nAfter blank\u0007'), { fileName: 'notes.txt' })
  assert.equal(paged.format, 'txt')
  assert.deepEqual(lines(paged), ['Page one', 'Page two\twith tab', '', 'After blank'])
  assert.equal(paged.document.blocks[1].style.pageBreakBefore, true)
  assert.ok(paged.document.blocks.every((block) => block.style.spaceAfterPx === 0), 'lines are not spaced apart')
  const empty = importText(Buffer.alloc(0))
  assert.equal(empty.document.blocks.length, 1)
  assertSelfContained(paged.document)
})

// ---- RTF -----------------------------------------------------------------------------

const WORDPAD_RTF = String.raw`{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil\fcharset0 Calibri;}{\f1\fnil\fcharset2 Symbol;}{\f2\fnil\fcharset204 Arial Cyr;}}
{\colortbl ;\red255\green0\blue0;\red0\green77\blue187;\red255\green255\blue0;}
{\*\generator Riched20 10.0.22621}\viewkind4\uc1
\pard\sa200\sl276\slmult1\qc\b\f0\fs32\lang9 WordPad Title\b0\fs22\par
\pard\sa200\sl276\slmult1 Normal \i italic\i0  \ul underlined\ulnone  \strike struck\strike0  \cf1 red\cf0  \highlight3 marked\highlight0  {\f2\'cf\'f0\'e8\'e2\'e5\'f2} na\'efve \'93q\'94 \u8364?5\par
\pard{\pntext\f1\'B7\tab}{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\'B7}}\fi-360\li720\sa200\sl276\slmult1 First bullet\par
{\pntext\f1\'B7\tab}Second bullet\par
\pard{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}\fi-360\li720\sa200\sl276\slmult1 One\par
{\pntext\f0 2.\tab}Two\par
\pard\sa200\sl276\slmult1 Picture:\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal15\pichgoal15 0100090000}\par
{\*\shppict{\pict\pngblip\picw1\pich1\picwgoal300\pichgoal150 89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63fccff0500f000485018084a98c210000000049454e44ae426082}}{\nonshppict{\pict\wmetafile8 0100}}\par
Line one\line Line two\page Next page\par
}`

test('RTF written by WordPad: formatting, code pages, bullets, numbering, pictures and breaks', async () => {
  const { importDocument } = await load()
  const result = await importDocument(Buffer.from(WORDPAD_RTF, 'latin1'), { fileName: 'letter.rtf', idPrefix: 'wp-' })
  const doc = result.document
  assert.equal(result.format, 'rtf')
  const title = find(doc.blocks, /^WordPad Title$/)
  assert.deepEqual([title.style.align, title.runs[0].style.bold, title.runs[0].style.fontSizePx, title.runs[0].style.fontFamily, title.style.lineHeight], ['center', true, 21.333, 'Calibri', 1.15])
  assert.equal(title.style.spaceAfterPx, 13.33)
  const normal = find(doc.blocks, /^Normal/)
  assert.equal(text(normal), 'Normal italic underlined struck red marked Привет naïve “q” €5')
  const style = (value) => normal.runs.find((run) => run.text === value).style
  assert.ok(style('italic').italic && style('underlined').underline && style('struck').strikethrough)
  assert.equal(style('red').color, '#ff0000')
  assert.equal(style('marked').highlightColor, '#ffff00')
  assert.equal(style('Привет').fontFamily, 'Arial Cyr')
  const items = listParagraphs(doc)
  assert.deepEqual(items.map(text), ['First bullet', 'Second bullet', 'One', 'Two'])
  assert.equal(items[0].style.list.listId, items[1].style.list.listId)
  assert.equal(items[2].style.list.listId, items[3].style.list.listId)
  assert.notEqual(items[0].style.list.listId, items[2].style.list.listId)
  assert.deepEqual([doc.lists[items[0].style.list.listId].levels[0].format, doc.lists[items[0].style.list.listId].levels[0].bulletChar], ['bullet', '•'])
  assert.equal(doc.lists[items[2].style.list.listId].levels[0].format, 'decimal')
  const [picture] = images(doc)
  assert.deepEqual([picture.widthPx, picture.heightPx], [20, 10])
  assert.match(picture.src, /^data:image\/png;base64,/)
  assert.ok(find(doc.blocks, /^\[Picture\]$/), 'a WMF-only picture is a visible placeholder')
  assert.match(result.warnings.join(' '), /WMF, EMF/)
  assert.equal(text(find(doc.blocks, /^Line one/)), 'Line one\vLine two')
  assert.equal(find(doc.blocks, /^Next page$/).style.pageBreakBefore, true)
  assertSelfContained(doc)
})

const WORD_RTF = String.raw`{\rtf1\ansi\ansicpg1252\uc1\deff0
{\fonttbl{\f0\froman\fcharset0 Times New Roman;}{\f1\fswiss\fcharset0 Arial;}{\f2\fnil\fcharset2 Symbol;}}
{\colortbl;\red0\green0\blue0;\red0\green112\blue192;\red217\green217\blue217;}
{\stylesheet{\ql \f0\fs24 \snext0 Normal;}{\s2\ql\b\f1\fs28\outlinelevel1 \sbasedon0 \snext0 heading 2;}{\*\cs10 \additive Default Paragraph Font;}}
{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc23\levelstartat1{\leveltext\'01\u-3913 ?;}{\levelnumbers;}\f2\fi-360\li720}{\listlevel\levelnfc4\levelstartat1{\leveltext\'02\'01);}{\levelnumbers\'01;}\fi-360\li1440}\listid100}
{\list\listtemplateid2{\listlevel\levelnfc0\levelstartat3{\leveltext\'02\'00.;}{\levelnumbers\'01;}\fi-360\li720}\listid200}}
{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls1}{\listoverride\listid200\listoverridecount0\ls2}}
{\info{\title Word RTF}}
\paperw11906\paperh16838\margl1134\margr1134\margt1134\margb1134
{\header \pard\plain\qc\f0\fs20 Page header\par}
{\footerf \pard\plain First footer only\par}
\pard\plain\s2\outlinelevel1\b\f1\fs28 Word Heading\par
\pard\plain\ql\fi720\li360\sb120\sa240\f0\fs24 Body with {\field{\*\fldinst{\rtlch HYPERLINK "https://example.com/docs" }{\rtlch {\*\datafield 00d0c9ea}}}{\fldrslt{\cf2\ul link}}} and a note{\super\chftn}{\footnote\pard\plain\f0\fs20{\super\chftn} Footnote {\b body}.}\par
{\listtext\pard\plain\f2 \'b7\tab}\pard\plain\ls1\ilvl0\f0\fs24 Bullet one\par
{\listtext\pard\plain a)\tab}\pard\plain\ls1\ilvl1\f0\fs24 Sub item\par
{\listtext\pard\plain 3.\tab}\pard\plain\ls2\ilvl0\f0\fs24 Third\par
\trowd\trgaph108\trleft0\clcbpat3\clmgf\cellx2000\clmrg\cellx4000\cellx6000
\pard\plain\intbl Merged\cell\cell Right\cell\row
\trowd\trgaph108\trleft0\clvmgf\cellx2000\cellx4000\cellx6000
\pard\plain\intbl Down\cell Mid\cell End\cell\row
\trowd\trgaph108\trleft0\clvmrg\cellx2000\cellx4000\cellx6000
\pard\plain\intbl\cell M2\cell E2\cell\row
\pard\plain After table\par
{\upr{\pard ansi version\par}{\*\ud{\pard Unicode \u1488?\par}}}
}`

test('RTF written by Word: headings, list table, merged cells, links, footnotes, header and page size', async () => {
  const { importRtf } = await load()
  const result = importRtf(Buffer.from(WORD_RTF, 'latin1'), { idPrefix: 'wd-' })
  const doc = result.document
  assert.equal(result.title, 'Word RTF')
  const [heading] = headings(doc)
  assert.deepEqual([heading.style.namedStyle, text(heading), heading.runs[0].style.bold, heading.runs[0].style.fontFamily, heading.runs[0].style.fontSizePx], ['Heading2', 'Word Heading', true, 'Arial', 18.667])
  const body = find(doc.blocks, /^Body with/)
  assert.equal(text(body), 'Body with link and a note1')
  assert.deepEqual([body.style.indentLeftPx, body.style.indentFirstLinePx, body.style.spaceBeforePx, body.style.spaceAfterPx], [24, 48, 8, 16])
  const link = body.runs.find((run) => run.text === 'link').style
  assert.deepEqual([link.link, link.color, link.underline], ['https://example.com/docs', '#0070c0', true])
  const note = body.runs.find((run) => run.style.footnoteRef)
  assert.equal(note.style.verticalAlign, 'super')
  assert.equal(doc.footnotes[note.style.footnoteRef].map(text).join(''), 'Footnote body.')
  const items = listParagraphs(doc)
  assert.deepEqual(items.map((block) => [text(block), block.style.list.level]), [['Bullet one', 0], ['Sub item', 1], ['Third', 0]])
  const bullets = doc.lists[items[0].style.list.listId]
  assert.deepEqual([bullets.levels[0].format, bullets.levels[0].bulletChar, bullets.levels[1].format, bullets.levels[1].text], ['bullet', '•', 'lowerLetter', '%2)'])
  const numbers = doc.lists[items[2].style.list.listId]
  assert.deepEqual([numbers.levels[0].format, numbers.levels[0].start, numbers.levels[0].text], ['decimal', 3, '%1.'])
  const [table] = tables(doc)
  assert.deepEqual(table.rows.map((row) => row.cells.map((cell) => [cellText(cell), cell.colSpan ?? 1, cell.rowSpan ?? 1])), [
    [['Merged', 2, 1], ['Right', 1, 1]],
    [['Down', 1, 2], ['Mid', 1, 1], ['End', 1, 1]],
    [['M2', 1, 1], ['E2', 1, 1]],
  ])
  assert.equal(table.rows[0].cells[0].shading, '#d9d9d9')
  assert.deepEqual(table.colFractions, [0.3333, 0.3333, 0.3333])
  assert.ok(find(doc.blocks, /^After table$/))
  assert.ok(find(doc.blocks, /^Unicode א$/), 'the Unicode alternative of \\upr is used')
  assert.ok(!find(doc.blocks, /ansi version/))
  assert.deepEqual(doc.section.header.map(text), ['Page header'])
  assert.equal(doc.section.footer, undefined, 'first-page footers are not the default footer')
  assert.deepEqual([doc.section.pageWidthPx, doc.section.pageHeightPx, doc.section.marginPx.left], [793.73, 1122.53, 75.6])
  assertSelfContained(doc)
  const back = await engineRoundTrip(doc)
  assert.equal(Object.keys(back.footnotes || {}).length, 1)
  assert.equal(back.blocks.filter((block) => block.kind === 'table').length, 1)
})

test('RTF is refused when the bytes are not RTF; deep nesting and odd input do not crash', async () => {
  const { importRtf, sniffImportFormat } = await load()
  assert.throws(() => importRtf(Buffer.from('plain text')), /not a Rich Text/)
  const deep = `{\\rtf1 ${'{'.repeat(5000)}deep${'}'.repeat(5000)}\\par}`
  assert.ok(find(importRtf(deep).document.blocks, /deep/))
  const truncated = importRtf('{\\rtf1\\ansi {\\b bold text without end')
  assert.ok(find(truncated.document.blocks, /bold text without end/))
  assert.equal(sniffImportFormat(Buffer.from('{\\rtf1 x}'), 'report.doc'), 'rtf')
})

// ---- ODT -----------------------------------------------------------------------------

const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"'

async function libreOfficeStyleOdt() {
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${ODF_NS} office:version="1.3">
 <office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'"/></office:font-face-decls>
 <office:automatic-styles>
  <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:text-align="center" fo:break-before="page"/></style:style>
  <style:style style:name="P2" style:family="paragraph" style:parent-style-name="Text_20_body"><style:paragraph-properties fo:text-align="justify" fo:margin-left="0.5in" fo:text-indent="0.25in"/></style:style>
  <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:color="#c9211e"/></style:style>
  <style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic" style:text-underline-style="solid" fo:background-color="#ffff00" style:font-name="Liberation Sans" fo:font-size="14pt"/></style:style>
  <style:style style:name="T3" style:family="text"><style:text-properties style:text-position="super 58%"/></style:style>
  <style:style style:name="T4" style:family="text"><style:text-properties style:text-line-through-style="solid"/></style:style>
  <style:style style:name="Table1.A" style:family="table-column"><style:table-column-properties style:column-width="2in"/></style:style>
  <style:style style:name="Table1.B" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style>
  <style:style style:name="Table1.A1" style:family="table-cell"><style:table-cell-properties fo:background-color="#dddddd"/></style:style>
  <text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"/><text:list-level-style-bullet text:level="2" text:bullet-char="◦"/></text:list-style>
  <text:list-style style:name="L2"><text:list-level-style-number text:level="1" style:num-prefix="(" style:num-suffix=")" style:num-format="i" text:start-value="2"/></text:list-style>
 </office:automatic-styles>
 <office:body><office:text>
  <text:sequence-decls><text:sequence-decl text:display-outline-level="0" text:name="Illustration"/></text:sequence-decls>
  <text:h text:style-name="Heading_20_1" text:outline-level="1">ODT Heading</text:h>
  <text:p text:style-name="Text_20_body">Plain <text:span text:style-name="T1">bold red</text:span> and <text:span text:style-name="T2">marked</text:span> x<text:s text:c="3"/>y<text:tab/>z E=mc<text:span text:style-name="T3">2</text:span> <text:span text:style-name="T4">gone</text:span> <text:a xlink:type="simple" xlink:href="https://example.org/">a link</text:a>.<text:note text:id="ftn1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p text:style-name="Footnote">ODT note text.</text:p></text:note-body></text:note><office:annotation><text:p>comment text</text:p></office:annotation></text:p>
  <text:list text:style-name="L1"><text:list-item><text:p>First</text:p><text:list><text:list-item><text:p>Inner</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>Second</text:p><text:p>Second, more</text:p></text:list-item></text:list>
  <text:p text:style-name="P2">Between lists</text:p>
  <text:list text:style-name="L1" text:continue-numbering="true"><text:list-item><text:p>Third</text:p></text:list-item></text:list>
  <text:list text:style-name="L2"><text:list-item><text:p>Roman</text:p></text:list-item></text:list>
  <table:table table:name="Table1"><table:table-column table:style-name="Table1.A"/><table:table-column table:style-name="Table1.B" table:number-columns-repeated="2"/>
   <table:table-header-rows><table:table-row><table:table-cell table:style-name="Table1.A1" table:number-columns-spanned="2" office:value-type="string"><text:p>Wide</text:p></table:table-cell><table:covered-table-cell/><table:table-cell><text:p>C1</text:p></table:table-cell></table:table-row></table:table-header-rows>
   <table:table-row><table:table-cell table:number-rows-spanned="2"><text:p>Tall</text:p></table:table-cell><table:table-cell><text:p>B2</text:p></table:table-cell><table:table-cell><text:p>C2</text:p></table:table-cell></table:table-row>
   <table:table-row><table:covered-table-cell/><table:table-cell><text:p>B3</text:p></table:table-cell><table:table-cell><text:p>C3</text:p></table:table-cell></table:table-row>
  </table:table>
  <text:p text:style-name="P1"><draw:frame draw:name="Image1" text:anchor-type="as-char" svg:width="1in" svg:height="0.5in"><draw:image xlink:href="Pictures/pic.png" xlink:type="simple"/></draw:frame></text:p>
  <text:p>Line<text:line-break/>break <draw:frame svg:width="1in" svg:height="1in"><draw:image xlink:href="https://example.com/remote.png"/><svg:title>Remote logo</svg:title></draw:frame></text:p>
 </office:text></office:body>
</office:document-content>`
  const styles = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${ODF_NS} office:version="1.3">
 <office:font-face-decls><style:font-face style:name="Liberation Serif" svg:font-family="'Liberation Serif'"/></office:font-face-decls>
 <office:styles>
  <style:default-style style:family="paragraph"><style:text-properties style:font-name="Liberation Serif" fo:font-size="12pt"/></style:default-style>
  <style:style style:name="Standard" style:family="paragraph" style:class="text"/>
  <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-top="0.1665in" fo:margin-bottom="0.0835in" fo:keep-with-next="always"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="14pt"/></style:style>
  <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:default-outline-level="1"><style:text-properties fo:font-size="130%" fo:font-weight="bold"/></style:style>
  <style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-top="0in" fo:margin-bottom="0.0972in" fo:line-height="115%"/></style:style>
 </office:styles>
 <office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2.5cm" fo:margin-right="2.5cm"/></style:page-layout></office:automatic-styles>
 <office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>ODT header</text:p></style:header><style:footer><text:p>Footer <text:page-number text:select-page="current">1</text:page-number></text:p></style:footer></style:master-page></office:master-styles>
</office:document-styles>`
  const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${ODF_NS}><office:meta><dc:title>ODT Fixture</dc:title></office:meta></office:document-meta>`
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' })
  zip.file('content.xml', content)
  zip.file('styles.xml', styles)
  zip.file('meta.xml', meta)
  zip.file('Pictures/pic.png', PNG)
  zip.file('META-INF/manifest.xml', '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>')
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

test('ODT written by LibreOffice: styles, lists, tables with spans, notes, pictures, header and page', async () => {
  const { importDocument } = await load()
  const bytes = await libreOfficeStyleOdt()
  const result = await importDocument(bytes, { fileName: 'fixture.odt', idPrefix: 'od-' })
  const doc = result.document
  assert.equal(result.format, 'odt')
  assert.equal(result.title, 'ODT Fixture')
  const [heading] = headings(doc)
  assert.deepEqual([heading.style.namedStyle, text(heading), heading.runs[0].style.bold, heading.runs[0].style.fontFamily, heading.runs[0].style.fontSizePx], ['Heading1', 'ODT Heading', true, 'Liberation Sans', 24.267])
  const body = find(doc.blocks, /^Plain/)
  assert.equal(text(body), 'Plain bold red and marked x   y\tz E=mc2 gone a link.1')
  assert.deepEqual([body.style.lineHeight, body.style.spaceAfterPx, body.runs[0].style.fontFamily, body.runs[0].style.fontSizePx], [1.15, 9.33, 'Liberation Serif', 16])
  const run = (value) => body.runs.find((item) => item.text === value).style
  assert.deepEqual([run('bold red').bold, run('bold red').color], [true, '#c9211e'])
  assert.deepEqual([run('marked').italic, run('marked').underline, run('marked').highlightColor, run('marked').fontFamily, run('marked').fontSizePx], [true, true, '#ffff00', 'Liberation Sans', 18.667])
  assert.equal(run('2').verticalAlign, 'super')
  assert.equal(run('gone').strikethrough, true)
  assert.equal(run('a link').link, 'https://example.org/')
  const note = body.runs.find((item) => item.style.footnoteRef)
  assert.equal(doc.footnotes[note.style.footnoteRef].map(text).join(''), 'ODT note text.')
  assert.ok(!JSON.stringify(doc).includes('comment text'), 'comments are not body text')
  const items = listParagraphs(doc)
  assert.deepEqual(items.map((block) => [text(block), block.style.list.level]), [['First', 0], ['Inner', 1], ['Second', 0], ['Third', 0], ['Roman', 0]])
  const [first, inner, second, third, roman] = items
  assert.ok([inner, second, third].every((block) => block.style.list.listId === first.style.list.listId), 'continue-numbering keeps the list')
  assert.notEqual(roman.style.list.listId, first.style.list.listId)
  const levels = doc.lists[first.style.list.listId].levels
  assert.deepEqual([levels[0].format, levels[0].text, levels[1].format, levels[1].bulletChar], ['decimal', '%1.', 'bullet', '◦'])
  assert.deepEqual([doc.lists[roman.style.list.listId].levels[0].format, doc.lists[roman.style.list.listId].levels[0].text, doc.lists[roman.style.list.listId].levels[0].start], ['lowerRoman', '(%1)', 2])
  const more = find(doc.blocks, /^Second, more$/)
  assert.deepEqual([more.style.list, more.style.indentLeftPx], [undefined, 48])
  const between = find(doc.blocks, /^Between lists$/)
  assert.deepEqual([between.style.align, between.style.indentLeftPx, between.style.indentFirstLinePx], ['justify', 48, 24])
  const [table] = tables(doc)
  assert.deepEqual(table.rows.map((row) => row.cells.map((cell) => [cellText(cell), cell.colSpan ?? 1, cell.rowSpan ?? 1])), [
    [['Wide', 2, 1], ['C1', 1, 1]],
    [['Tall', 1, 2], ['B2', 1, 1], ['C2', 1, 1]],
    [['B3', 1, 1], ['C3', 1, 1]],
  ])
  assert.equal(table.rows[0].cells[0].shading, '#dddddd')
  assert.deepEqual(table.colFractions, [0.5, 0.25, 0.25])
  const [picture] = images(doc)
  assert.deepEqual([picture.widthPx, picture.heightPx, picture.align], [96, 48, 'center'])
  const pictureIndex = doc.blocks.indexOf(picture)
  assert.equal(doc.blocks[pictureIndex - 1].style.pageBreakBefore, true, 'the page break before the picture is kept')
  assert.equal(text(find(doc.blocks, /^Line/)), 'Line\vbreak [Picture: Remote logo]')
  assert.match(result.warnings.join(' '), /could not be read/)
  assert.deepEqual(doc.section.header.map(text), ['ODT header'])
  assert.deepEqual(doc.section.footer.map(text), ['Footer 1'])
  assert.deepEqual([doc.section.pageWidthPx, doc.section.pageHeightPx, doc.section.marginPx.left, doc.section.marginPx.top], [793.74, 1122.52, 94.49, 75.59])
  assertSelfContained(doc)
  const back = await engineRoundTrip(doc)
  assert.equal(back.blocks.filter((block) => block.kind === 'image').length, 1)
  assert.equal(Object.keys(back.footnotes || {}).length, 1)
})

test('ODT written by Simple itself comes back with the same structure', async () => {
  const { importMarkdown, importOdt } = await load()
  const { documentToOdt } = await import('../src/exporters/odt.js')
  const source = `# Title\n\nSome **bold** and [link](https://example.com/).[^1]\n\n## Sub\n\n1. one\n2. two\n   - nested\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n![pic](data:image/png;base64,${PNG.toString('base64')})\n\n[^1]: The note.\n`
  const original = importMarkdown(source, { idPrefix: 'a-' }).document
  const odt = await documentToOdt(original, { title: 'Round trip' })
  const result = await importOdt(odt, { idPrefix: 'b-' })
  const doc = result.document
  assert.equal(result.title, 'Round trip')
  assert.deepEqual(headings(doc).map((block) => [block.style.namedStyle, text(block)]), [['Heading1', 'Title'], ['Heading2', 'Sub']])
  assert.deepEqual(listParagraphs(doc).map((block) => [text(block), block.style.list.level]), [['one', 0], ['two', 0], ['nested', 1]])
  assert.equal(tables(doc).length, 1)
  assert.equal(images(doc).length, 1)
  assert.equal(runWith(doc.blocks, /^link$/).style.link, 'https://example.com/')
  assert.equal(Object.values(doc.footnotes).map((note) => note.map(text).join('')).join(), 'The note.')
  assertSelfContained(doc)
})

test('ODT: wrong packages are refused with a plain message; flat .fodt opens', async () => {
  const { importOdt, importDocument, sniffImportFormat } = await load()
  await assert.rejects(importOdt(Buffer.from('not a zip')), /not an OpenDocument text document/)
  const spreadsheet = new JSZip()
  spreadsheet.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet')
  spreadsheet.file('content.xml', '<office:document-content/>')
  await assert.rejects(importOdt(await spreadsheet.generateAsync({ type: 'nodebuffer' })), /not a text document/)
  const flat = `<?xml version="1.0" encoding="UTF-8"?><office:document ${ODF_NS} office:mimetype="application/vnd.oasis.opendocument.text"><office:meta><dc:title>Flat</dc:title></office:meta><office:body><office:text><text:h text:outline-level="2">Flat heading</text:h><text:p>Flat body</text:p></office:text></office:body></office:document>`
  assert.equal(sniffImportFormat(Buffer.from(flat), 'x.fodt'), 'odt')
  const result = await importDocument(Buffer.from(flat), { fileName: 'x.fodt' })
  assert.equal(result.title, 'Flat')
  assert.deepEqual(headings(result.document).map((block) => [block.style.namedStyle, text(block)]), [['Heading2', 'Flat heading']])
  assert.ok(find(result.document.blocks, /^Flat body$/))
})

// ---- insertion into an open document -------------------------------------------------

test('blocks prepared for insertBlocks only reference lists and notes the target has', async () => {
  const { importMarkdown, blocksForInsert, formatListNumber } = await load()
  const { document } = importMarkdown('1. one\n2. two\n   1. inner\n3. three\n\n- dot\n\nSee[^x].\n\n[^x]: Note.\n', { idPrefix: 'i-' })
  const blocks = blocksForInsert(document, { lists: {}, footnotes: {} })
  const lines = paragraphs(blocks).map(text)
  assert.deepEqual(lines, ['1.\tone', '2.\ttwo', 'a.\tinner', '3.\tthree', '•\tdot', 'See1.', '1 Note.'])
  assert.ok(paragraphs(blocks).every((block) => !block.style.list && block.runs.every((run) => !run.style.footnoteRef)))
  assert.deepEqual([blocks[2].style.indentLeftPx, blocks[2].style.indentFirstLinePx], [96, -24])
  assert.ok(document.blocks[0].style.list, 'the imported document itself is not changed')
  const kept = blocksForInsert(document, { lists: document.lists, footnotes: document.footnotes })
  assert.equal(kept.filter((block) => block.style?.list).length, 5)
  assert.deepEqual([formatListNumber(4, 'upperRoman'), formatListNumber(28, 'lowerLetter'), formatListNumber(7, 'decimal')], ['IV', 'ab', '7'])
})

// ---- dispatch ------------------------------------------------------------------------

test('the importer is chosen from the content first, then the file name', async () => {
  const { sniffImportFormat, importFormatForName, importDocument, IMPORT_EXTENSIONS } = await load()
  assert.deepEqual(Object.keys(IMPORT_EXTENSIONS).filter((ext) => ['txt', 'md', 'html', 'htm', 'rtf', 'odt'].includes(ext)).sort(), ['htm', 'html', 'md', 'odt', 'rtf', 'txt'])
  assert.equal(importFormatForName('C:\\Docs\\Read Me.MD'), 'md')
  assert.equal(importFormatForName('page.HTM'), 'html')
  assert.equal(importFormatForName('archive.zip'), null)
  assert.equal(sniffImportFormat(Buffer.from('{\\rtf1 hi}'), 'notes.txt'), 'rtf')
  assert.equal(sniffImportFormat(await libreOfficeStyleOdt(), 'fixture.zip'), 'odt')
  assert.equal(sniffImportFormat(Buffer.from('<!DOCTYPE html><p>x</p>'), 'saved'), 'html')
  assert.equal(sniffImportFormat(Buffer.from('<p>x</p>'), 'notes.txt'), 'txt', 'a .txt with markup stays text')
  assert.equal(sniffImportFormat(Buffer.from('# Title'), 'readme.md'), 'md')
  await assert.rejects(importDocument(Buffer.from('x'), { fileName: 'image.png' }), /cannot open this kind of file/)
})
