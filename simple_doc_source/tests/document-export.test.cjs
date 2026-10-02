const test = require('node:test')
const assert = require('node:assert/strict')

const style = (overrides = {}) => ({
  fontFamily: 'Calibri',
  fontSizePx: 14,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: '#111111',
  ...overrides,
})

const paragraph = (id, text, overrides = {}) => ({
  kind: 'paragraph',
  id,
  revision: 1,
  runs: [{ text, style: style() }],
  style: {
    align: 'left',
    lineHeight: 1.15,
    spaceBeforePx: 0,
    spaceAfterPx: 6,
    indentFirstLinePx: 0,
    indentLeftPx: 0,
    ...overrides,
  },
})

function fixture() {
  const heading = paragraph('heading', 'Intro <script>alert(1)</script>', { outlineLevel: 1 })
  heading.runs.push({ text: ' safe', style: style({ bold: true, link: 'https://example.com' }) })
  heading.runs.push({ text: ' unsafe', style: style({ link: 'javascript:alert(1)' }) })
  heading.runs.push({ text: ' secret', style: style({ hidden: true }) })
  return {
    section: {
      pageWidthPx: 816,
      pageHeightPx: 1056,
      marginPx: { top: 72, right: 72, bottom: 72, left: 72 },
    },
    lists: {
      numbered: {
        id: 'numbered',
        levels: [{ format: 'decimal', text: '%1.', indentLeftPx: 24, hangingPx: 12, start: 1 }],
      },
    },
    blocks: [
      heading,
      paragraph('list-1', 'First item', { list: { listId: 'numbered', level: 0 } }),
      paragraph('list-2', 'Second item', { list: { listId: 'numbered', level: 0 } }),
      {
        kind: 'table', id: 'table', revision: 1, rows: [{ cells: [
          { id: 'cell-a', blocks: [paragraph('cell-a-p', 'Name')] },
          { id: 'cell-b', blocks: [paragraph('cell-b-p', 'Value')] },
        ] }, { cells: [
          { id: 'cell-c', blocks: [paragraph('cell-c-p', 'Alpha')] },
          { id: 'cell-d', blocks: [paragraph('cell-d-p', '42')] },
        ] }],
      },
      { kind: 'image', id: 'image', revision: 1, src: 'data:image/png;base64,iVBORw0KGgo=', widthPx: 80, heightPx: 40, align: 'center' },
      { kind: 'custom', id: 'custom', revision: 1, customType: 'example', data: {} },
    ],
  }
}

test('HTML export is standalone, formatted, and sanitizes active content', async () => {
  const { documentToHtml } = await import('../src/document-export.js')
  const output = documentToHtml(fixture(), 'Quarterly <Review>')
  assert.match(output, /^<!doctype html>/)
  assert.match(output, /<title>Quarterly &lt;Review&gt;<\/title>/)
  assert.match(output, /Intro &lt;script&gt;alert\(1\)&lt;\/script&gt;/)
  assert.match(output, /font-weight:700/)
  assert.match(output, /href="https:\/\/example\.com\/"/)
  assert.doesNotMatch(output, /javascript:/)
  assert.doesNotMatch(output, /secret/)
  assert.match(output, /<table>/)
  assert.match(output, /data:image\/png;base64,iVBORw0KGgo=/)
  assert.match(output, /\[Unsupported document object\]/)
})

test('Markdown export preserves headings, numbering, tables, emphasis, and images', async () => {
  const { documentToMarkdown } = await import('../src/document-export.js')
  const output = documentToMarkdown(fixture(), 'Quarterly Review')
  // No injected file-name title (DOC-SIE-23); the heading keeps its own level
  // and the link's leading space stays outside the brackets.
  assert.doesNotMatch(output, /^# Quarterly Review/m)
  assert.match(output, /^## Intro &lt;script&gt;alert\(1\)&lt;\/script&gt; \[safe\]\(https:\/\/example\.com\/\) unsafe$/m)
  assert.match(output, /^1\. First item$/m)
  assert.match(output, /^2\. Second item$/m)
  assert.match(output, /^\| Name \| Value \|$/m)
  assert.match(output, /^!\[Document image\]\(data:image\/png;base64,iVBORw0KGgo=\)$/m)
  assert.doesNotMatch(output, /secret/)
  assert.doesNotMatch(output, /javascript:/)
})

test('plain-text export remains readable without leaking hidden runs', async () => {
  const { documentToText, serializeDocument } = await import('../src/document-export.js')
  const output = documentToText(fixture())
  assert.match(output, /Intro <script>alert\(1\)<\/script> safe unsafe/)
  assert.match(output, /1\. First item\n2\. Second item/)
  assert.match(output, /Name\tValue\nAlpha\t42/)
  assert.match(output, /\[Image\]/)
  assert.doesNotMatch(output, /secret/)
  assert.equal(serializeDocument(fixture(), 'txt'), output)
  assert.throws(() => serializeDocument(fixture(), 'csv'), /Unsupported structured document export/)
})

test('unsafe and unsupported image sources become explicit placeholders', async () => {
  const { documentToHtml, documentToMarkdown } = await import('../src/document-export.js')
  const document = fixture()
  document.blocks[4].src = 'https://example.com/tracker.png'
  assert.match(documentToHtml(document), /\[Image not available in this export\]/)
  assert.match(documentToMarkdown(document), /\[Image not available in this export\]/)
  assert.doesNotMatch(documentToHtml(document), /tracker\.png/)
})

test('object-URL images are embedded once without mutating the live document or fetching remote sources', async () => {
  const { prepareDocumentForExport, documentToHtml, documentToMarkdown } = await import('../src/document-export.js')
  const document = fixture()
  document.blocks[4].src = 'blob:local-image'
  document.section.header = [structuredClone(document.blocks[4])]
  const calls = []
  const prepared = await prepareDocumentForExport(document, async (source) => {
    calls.push(source)
    return 'data:image/png;base64,iVBORw0KGgo='
  })
  assert.deepEqual(calls, ['blob:local-image'])
  assert.equal(document.blocks[4].src, 'blob:local-image')
  assert.match(documentToHtml(prepared), /data:image\/png;base64/)
  assert.match(documentToMarkdown(prepared), /data:image\/png;base64/)
  // A source that cannot be embedded degrades to a placeholder instead of aborting.
  const { exportWarnings } = await import('../src/document-export.js')
  const degraded = await prepareDocumentForExport(document, async () => 'https://example.com')
  assert.match(documentToHtml(degraded), /\[Image not available in this export\]/)
  assert.doesNotMatch(documentToHtml(degraded), /example\.com"|blob:/)
  assert.equal(exportWarnings(degraded).length, 1)
})

test('HTML and Markdown exports skip what they cannot embed and report it (DOC-SIE-18)', async () => {
  const { prepareDocumentForExport, serializeDocumentWithWarnings, exportWarnings } = await import('../src/document-export.js')
  const document = fixture()
  document.blocks[4].src = 'blob:svg-image'
  document.blocks.push({ ...structuredClone(document.blocks[4]), id: 'huge', src: 'blob:huge-image' })
  document.blocks.push({ ...structuredClone(document.blocks[4]), id: 'tiff', src: 'blob:tiff-image' })
  const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><rect width="4" height="4"/></svg>')
  const blob = (type, bytes) => ({ type, size: bytes.length, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) })
  const read = []
  const prepared = await prepareDocumentForExport(document, async (source) => {
    read.push(source)
    if (source === 'blob:svg-image') return blob('image/svg+xml', svg)
    if (source === 'blob:huge-image') return { type: 'image/png', size: 70 * 1024 * 1024, arrayBuffer: async () => { throw new Error('a 70 MB image must not be read') } }
    return blob('image/tiff', Buffer.from('II*\u0000'))
  })
  assert.deepEqual(read.sort(), ['blob:huge-image', 'blob:svg-image', 'blob:tiff-image'])
  const warnings = exportWarnings(prepared)
  assert.equal(warnings.length, 2, warnings.join(' | '))
  assert.match(warnings.join(' '), /larger than 64 MB/)
  assert.match(warnings.join(' '), /format this export cannot embed/)
  for (const format of ['html', 'md']) {
    const result = serializeDocumentWithWarnings(prepared, format, 'Images')
    assert.match(result.text, /data:image\/svg\+xml;base64,/, `${format} embeds the SVG as a data URI`)
    assert.equal((result.text.match(/Image not available in this export/g) || []).length, 2, `${format} keeps two placeholders`)
    assert.deepEqual(result.warnings.slice(0, 2), warnings)
    assert.match(result.warnings.join(' '), /object that this format cannot represent/)
  }
  // A clean export reports nothing.
  const clean = fixture()
  clean.blocks.pop()
  assert.deepEqual(serializeDocumentWithWarnings(clean, 'md').warnings, [])
})

const run = (text, overrides = {}) => ({ text, style: style(overrides) })
const paragraphWithRuns = (id, runs, overrides = {}) => ({ ...paragraph(id, ''), runs, style: { ...paragraph(id, '').style, ...overrides } })

test('Markdown emphasis hugs non-space text and never emits **a****b** (doc-io-fidelity-9)', async () => {
  const { documentToMarkdown } = await import('../src/document-export.js')
  const document = { ...fixture(), blocks: [
    paragraphWithRuns('a', [run('Note: ', { bold: true }), run('check totals')]),
    paragraphWithRuns('b', [run('Total', { bold: true, color: '#ff0000' }), run(':', { bold: true }), run(' 42')]),
    paragraphWithRuns('c', [run('re'), run('write', { italic: true }), run(' and '), run(' safe', { bold: true, link: 'https://example.com' }), run(' done')]),
    paragraphWithRuns('d', [run('(note)', { bold: true }), run('after')]),
    paragraphWithRuns('e', [run('  ', { bold: true }), run('plain '), run('struck', { strikethrough: true }), run(' x', { italic: true, bold: true })]),
  ] }
  const output = documentToMarkdown(document)
  assert.match(output, /^\*\*Note:\*\* check totals$/m)
  assert.match(output, /^\*\*Total:\*\* 42$/m)
  assert.doesNotMatch(output, /\*\*\*\*/)
  assert.match(output, /^re\*write\* and  \[\*\*safe\*\*\]\(https:\/\/example\.com\/\) done$/m)
  // A delimiter next to punctuation and a letter is not flanking; inline HTML keeps it bold.
  assert.match(output, /^<strong>\(note\)<\/strong>after$/m)
  assert.match(output, /^plain ~~struck~~ \*\*\*x\*\*\*$/m)
  // No marker ever touches whitespace on its inner side: an opener is never
  // followed by a space and a closer never follows one.
  assert.doesNotMatch(output, /(^|\s)(\*\*|~~|\*)\s/m)
  assert.doesNotMatch(output, /\s(\*\*|~~)(\s|$)/m)
})

test('Markdown is clean GFM: no title line, selective escaping, footnotes, tight lists (DOC-SIE-23)', async () => {
  const { documentToMarkdown } = await import('../src/document-export.js')
  const document = fixture()
  document.footnotes = { fn7: [paragraph('fn', 'Footnote text.')] }
  document.blocks = [
    paragraphWithRuns('h', [run('Results', { bold: true })], { outlineLevel: 0 }),
    paragraphWithRuns('p', [run('Version 1.2 (final) costs $5 - see the notes.'), { text: '1', style: style({ footnoteRef: 'fn7' }) }]),
    paragraph('n1', '# not a heading'),
    paragraph('n2', '2. not a list'),
    paragraph('l1', 'One', { list: { listId: 'numbered', level: 0 } }),
    paragraph('l2', 'Two', { list: { listId: 'numbered', level: 0 } }),
    paragraph('after', 'Done.'),
    paragraph('empty', ''),
  ]
  document.section.header = [paragraph('hd', '')]
  document.section.footer = [paragraph('ft', 'Footer words')]
  const output = documentToMarkdown(document, 'Complex document')
  assert.doesNotMatch(output, /^# Complex document/m)
  assert.match(output, /^# Results$/m, 'headings carry no ** markers')
  assert.doesNotMatch(output, /# \*\*/)
  assert.match(output, /^Version 1\.2 \(final\) costs \$5 - see the notes\.\[\^1\]$/m)
  assert.ok((output.match(/\\\./g) || []).length < 5)
  assert.match(output, /^\[\^1\]: Footnote text\.$/m)
  assert.doesNotMatch(output, /## Footnote/)
  assert.match(output, /^\\# not a heading$/m)
  assert.match(output, /^2\\\. not a list$/m)
  assert.match(output, /^1\. One\n2\. Two\n\nDone\.$/m, 'tight list followed by a paragraph')
  assert.doesNotMatch(output, /\*\*Header\*\*/, 'empty header stories are skipped')
  assert.match(output, /\*\*Footer\*\*\n\nFooter words/)
})

test('HTML internal links have bookmark targets and table header rows use th (DOC-SIE-23)', async () => {
  const { documentToHtml } = await import('../src/document-export.js')
  const document = fixture()
  document.bookmarks = {
    _Toc123: { start: { blockId: 'list-2', offset: 0 }, end: { blockId: 'list-2', offset: 6 } },
    middle: { start: { blockId: 'heading', offset: 6 }, end: { blockId: 'heading', offset: 9 } },
  }
  document.blocks.push(paragraphWithRuns('links', [run('Go to '), run('item', { link: '#_Toc123' }), run(' or '), run('middle', { link: '#middle' })]))
  document.blocks[3].condOverrides = { firstRow: true }
  document.footnotes = { f1: [paragraph('fn', 'Note body')] }
  document.blocks.push(paragraphWithRuns('ref', [run('See'), { text: '1', style: style({ footnoteRef: 'f1' }) }]))
  const output = documentToHtml(document, 'Links', { lang: 'uz-Cyrl', author: 'A. Author' })
  for (const target of output.matchAll(/href="#([^"]+)"/g)) assert.match(output, new RegExp(`id="${target[1]}"`), `#${target[1]} has a target`)
  assert.match(output, /<h2[^>]*><span[^>]*>Intro <\/span>/)
  assert.match(output, /Intro <\/span>|Intro /)
  assert.match(output, /<a id="middle"><\/a>/)
  assert.match(output, /<thead><tr><th>/)
  assert.match(output, /<html lang="uz-Cyrl">/)
  assert.match(output, /<meta name="author" content="A\. Author">/)
  assert.match(output, /<li id="note-1">/)
})

test('all section header variants, notes, and soft breaks survive structured exports', async () => {
  const { serializeDocument } = await import('../src/document-export.js')
  const document = fixture()
  document.blocks[0].style.sectionBreak = { type: 'nextPage', props: {
    pageWidthPx: 816, pageHeightPx: 1056,
    header: [paragraph('h1', 'FIRST_SECTION_HEADER')],
    footerEven: [paragraph('fe', 'EVEN_FOOTER')],
  } }
  document.section.headerFirst = [paragraph('hf', 'FIRST_PAGE_HEADER')]
  document.section.pageWidthPx = 1056
  document.section.pageHeightPx = 816
  document.footnotes = { footnote1: [paragraph('fn', 'FOOTNOTE_CONTENT')] }
  document.endnotes = { endnote1: [paragraph('en', 'ENDNOTE_CONTENT')] }
  document.blocks.push(paragraph('soft', 'LINE_ONE\vLINE_TWO'))
  for (const format of ['html', 'md', 'txt']) {
    const output = serializeDocument(document, format)
    for (const token of ['FIRST_SECTION_HEADER', 'EVEN_FOOTER', 'FIRST_PAGE_HEADER', 'FOOTNOTE_CONTENT', 'ENDNOTE_CONTENT']) {
      assert.ok(output.replace(/\\/g, '').includes(token), `${format} retains ${token}`)
    }
    assert.doesNotMatch(output, /\v/)
  }
  assert.match(serializeDocument(document, 'html'), /@page section1\{size:1056px 816px\}/)
  assert.match(serializeDocument(document, 'html'), /LINE_ONE<br>LINE_TWO/)
})
