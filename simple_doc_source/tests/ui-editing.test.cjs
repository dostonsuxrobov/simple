const test = require('node:test')
const assert = require('node:assert/strict')

const editing = () => import('../src/ui/editing.ts')

const style = (extra = {}) => ({ fontFamily: 'Calibri', fontSizePx: 14.667, bold: false, italic: false, underline: false, strikethrough: false, color: '#111111', ...extra })
const paragraph = (id, ...runs) => ({ kind: 'paragraph', id, revision: 0, style: { align: 'left' }, runs: runs.map(([text, extra]) => ({ text, style: style(extra) })) })
const table = (id, ...cells) => ({ kind: 'table', id, rows: [{ cells: cells.map((blocks) => ({ blocks })) }] })

test('paragraphs are found in the body, table cells, header bands and notes', async () => {
  const { findParagraph, firstParagraphId, paragraphStories } = await editing()
  const doc = {
    blocks: [table('t', [paragraph('c1', ['cell one'])], [paragraph('c2', ['cell two'])]), paragraph('p1', ['body'])],
    section: { header: [paragraph('h1', ['head'])], pageWidthPx: 816 },
    footnotes: { 1: [paragraph('n1', ['note'])] },
  }
  assert.equal(firstParagraphId(doc), 'c1', 'a document starting with a table gets its caret in the first cell')
  assert.equal(findParagraph(doc, 'h1')?.runs[0].text, 'head')
  assert.equal(findParagraph(doc, 'n1')?.runs[0].text, 'note')
  assert.equal(findParagraph(doc, 'missing'), null)
  assert.deepEqual(paragraphStories(doc).map((story) => story.map((block) => block.id)), [['c1', 'c2', 'p1'], ['h1'], ['n1']])
  assert.equal(firstParagraphId({ blocks: [] }), null)
})

test('selection ranges cover each paragraph in reading order, backwards selections included', async () => {
  const { selectionRanges, isCollapsed } = await editing()
  const doc = { blocks: [paragraph('a', ['Hello world']), paragraph('b', ['Second']), paragraph('c', ['Third line'])] }
  const forward = selectionRanges(doc, { anchor: { blockId: 'a', offset: 6 }, focus: { blockId: 'c', offset: 5 } })
  assert.deepEqual(forward.map((range) => [range.block.id, range.start, range.end]), [['a', 6, 11], ['b', 0, 6], ['c', 0, 5]])
  const backward = selectionRanges(doc, { anchor: { blockId: 'c', offset: 5 }, focus: { blockId: 'a', offset: 6 } })
  assert.deepEqual(backward.map((range) => [range.block.id, range.start, range.end]), [['a', 6, 11], ['b', 0, 6], ['c', 0, 5]])
  const inside = selectionRanges(doc, { anchor: { blockId: 'a', offset: 9 }, focus: { blockId: 'a', offset: 2 } })
  assert.deepEqual(inside.map((range) => [range.start, range.end]), [[2, 9]])
  assert.equal(isCollapsed({ anchor: { blockId: 'a', offset: 2 }, focus: { blockId: 'a', offset: 2 } }), true)
  assert.equal(isCollapsed({ anchor: { blockId: 'a', offset: 2 }, focus: { blockId: 'a', offset: 3 } }), false)
})

test('the word at the caret includes a caret touching its start or end', async () => {
  const { wordRangeAt } = await editing()
  assert.deepEqual(wordRangeAt('Hello brave world', 8), { start: 6, end: 11 })
  assert.deepEqual(wordRangeAt('Hello brave world', 11), { start: 6, end: 11 })
  assert.deepEqual(wordRangeAt('Hello brave world', 6), { start: 6, end: 11 })
  assert.deepEqual(wordRangeAt('Привет мир', 3), { start: 0, end: 6 })
  assert.equal(wordRangeAt('Hello  world', 6), null)
  assert.equal(wordRangeAt('', 0), null)
})

test("Shift+F3 cycles lowercase, UPPERCASE and Capitalize Each Word like Word", async () => {
  const { nextCaseMode } = await editing()
  assert.equal(nextCaseMode('hello world'), 'upper')
  assert.equal(nextCaseMode('HELLO WORLD'), 'title')
  assert.equal(nextCaseMode('Hello World'), 'lower')
  assert.equal(nextCaseMode('hELLO wORLD'), 'lower')
  assert.equal(nextCaseMode('привет'), 'upper')
  assert.equal(nextCaseMode('2024 — 15%'), null)
})

test('link addresses are completed like Word and Docs, and unsafe ones refused', async () => {
  const { normalizeLinkAddress, linkAddressProblem } = await editing()
  assert.equal(normalizeLinkAddress(' example.com '), 'https://example.com')
  assert.equal(normalizeLinkAddress('www.example.org/a?b=1'), 'https://www.example.org/a?b=1')
  assert.equal(normalizeLinkAddress('name@example.com'), 'mailto:name@example.com')
  assert.equal(normalizeLinkAddress('http://example.com'), 'http://example.com')
  assert.equal(normalizeLinkAddress('#Chapter_one'), '#Chapter_one')
  assert.equal(normalizeLinkAddress('C:\\Reports\\Q1 plan.docx'), 'C:\\Reports\\Q1 plan.docx')
  assert.equal(normalizeLinkAddress('пример.рф'), 'https://пример.рф')
  assert.equal(linkAddressProblem(''), 'Type or paste an address.')
  assert.match(linkAddressProblem('javascript:alert(1)'), /web, email or document/)
  assert.match(linkAddressProblem('data:text/html,hi'), /web, email or document/)
  assert.match(linkAddressProblem('https://exa mple.com'), /spaces/)
  assert.equal(linkAddressProblem('example.com'), null)
  assert.equal(linkAddressProblem('C:\\Reports\\Q1 plan.docx'), null, 'file paths may contain spaces')
})

test("bookmark names follow Word's rules", async () => {
  const { bookmarkNameProblem } = await editing()
  assert.equal(bookmarkNameProblem('Chapter_1'), null)
  assert.equal(bookmarkNameProblem('Глава_2'), null)
  assert.match(bookmarkNameProblem('Chapter one'), /spaces/)
  assert.match(bookmarkNameProblem('1st'), /start with a letter/)
  assert.match(bookmarkNameProblem('a'.repeat(41)), /40/)
  assert.match(bookmarkNameProblem(' '), /Type a bookmark name/)
  assert.match(bookmarkNameProblem('Intro', ['Intro']), /already exists/)
  assert.equal(bookmarkNameProblem('Intro', ['Intro'], 'Intro'), null, 'renaming to the same name is not a clash')
})

test('Ctrl+Space resets manual character formatting to the paragraph style and keeps links and notes', async () => {
  const { clearCharacterRuns, resolveStyleChar } = await editing()
  const sheet = {
    defaultStyleId: 'Normal',
    styles: [
      { id: 'Normal', name: 'Normal', char: { fontFamily: 'Calibri', fontSizePx: 14.667, color: '#111111' }, para: {} },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', char: { fontSizePx: 24, bold: true }, para: {} },
    ],
  }
  assert.deepEqual(resolveStyleChar(sheet, 'Heading1'), { fontFamily: 'Calibri', fontSizePx: 24, color: '#111111', bold: true })
  assert.deepEqual(resolveStyleChar(sheet, undefined), { fontFamily: 'Calibri', fontSizePx: 14.667, color: '#111111' })
  const runs = [
    { text: 'Plain ', style: style() },
    { text: 'bold red link', style: style({ bold: true, color: '#ff0000', fontFamily: 'Arial', highlightColor: '#ffff00', link: 'https://example.com', charStyleId: 'Strong' }) },
    { text: ' note', style: style({ verticalAlign: 'super', footnoteRef: '1', italic: true }) },
  ]
  const cleared = clearCharacterRuns(runs, 6, 24, resolveStyleChar(sheet, 'Normal'))
  assert.deepEqual(cleared.map((run) => run.text), ['Plain ', 'bold red link', ' note'])
  const link = cleared[1].style
  assert.equal(link.bold, false)
  assert.equal(link.color, '#111111')
  assert.equal(link.fontFamily, 'Calibri')
  assert.equal(link.link, 'https://example.com', 'the hyperlink stays')
  assert.equal('highlightColor' in link, false)
  assert.equal('charStyleId' in link, false)
  const note = cleared[2].style
  assert.equal(note.footnoteRef, '1', 'the note reference stays')
  assert.equal(note.italic, false)
  assert.equal('verticalAlign' in note, false)
  assert.equal(cleared[0], runs[0], 'text outside the selection is untouched')
  const partial = clearCharacterRuns([{ text: 'abcdef', style: style({ bold: true }) }], 2, 4, {})
  assert.deepEqual(partial.map((run) => [run.text, run.style.bold]), [['ab', true], ['cd', false], ['ef', true]])
  const heading = clearCharacterRuns([{ text: 'Title', style: style({ italic: true, fontSizePx: 30 }) }], 0, 5, resolveStyleChar(sheet, 'Heading1'))
  assert.deepEqual([heading[0].style.bold, heading[0].style.italic, heading[0].style.fontSizePx], [true, false, 24], 'a heading keeps its style formatting')
})

test('link ranges, run styles and missing heading styles', async () => {
  const { linkRangeAt, runStyleAt, headingStyleDefinition } = await editing()
  const block = paragraph('p', ['See '], ['the ', { link: 'https://a.example' }], ['site', { link: 'https://a.example', bold: true }], [' now'])
  assert.deepEqual(linkRangeAt(block, 6), { start: 4, end: 12, link: 'https://a.example' })
  assert.deepEqual(linkRangeAt(block, 12), { start: 4, end: 12, link: 'https://a.example' }, 'a caret right after the link edits it')
  assert.equal(linkRangeAt(block, 2), null)
  assert.equal(runStyleAt(block, 10).bold, true)
  assert.equal(runStyleAt(block, 0).bold, false)
  const builderSheet = { defaultStyleId: 'Normal', styles: [{ id: 'Normal', name: 'Normal', char: {}, para: {} }, { id: 'Heading2', name: 'Heading 2', char: {}, para: {} }] }
  const h3 = headingStyleDefinition(3, builderSheet)
  assert.equal(h3.id, 'Heading3')
  assert.equal(h3.name, 'Heading 3')
  assert.equal(h3.basedOn, 'Heading2', 'builds on the document’s own Heading 2')
  assert.equal(h3.para.outlineLevel, 2)
  const wordDoc = headingStyleDefinition(1, { defaultStyleId: 'Normal', styles: [{ id: 'Normal', name: 'Normal', char: {}, para: {} }] })
  assert.equal(wordDoc.basedOn, 'Normal')
  assert.equal(wordDoc.para.outlineLevel, 0)
  assert.ok(wordDoc.char.fontSizePx > 20, "Word's Heading 1 size")
})
