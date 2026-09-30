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
  assert.match(output, /^# Quarterly Review/m)
  assert.match(output, /^## Intro &lt;script&gt;alert\\\(1\\\)&lt;\/script&gt;\[\*\* safe\*\*\]\(https:\/\/example\.com\/\) unsafe$/m)
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
  await assert.rejects(prepareDocumentForExport(document, async () => 'https://example.com'), /could not be included/)
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
