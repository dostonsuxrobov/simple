const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const JSZip = require('jszip')
const { parseRtf } = require('../electron/rtf-import.cjs')
const { flowText } = require('../electron/simple-docx.cjs')
const { ODT_MIME_TYPE, validateExportContent } = require('../electron/export-files.cjs')

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')
const GIF = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64')
const COMPLEX_DOCUMENT = path.join(__dirname, '..', '..', 'Simple test examples', 'Complex document.docx')

const style = (overrides = {}) => ({ fontFamily: 'Calibri', fontSizePx: 14.667, bold: false, italic: false, underline: false, strikethrough: false, color: '#111111', ...overrides })
const run = (text, overrides = {}) => ({ text, style: style(overrides) })
const paragraph = (id, runs, overrides = {}) => ({ kind: 'paragraph', id, revision: 1, runs: typeof runs === 'string' ? [run(runs)] : runs, style: { align: 'left', lineHeight: 1.15, spaceBeforePx: 0, spaceAfterPx: 6, indentFirstLinePx: 0, indentLeftPx: 0, ...overrides } })
const margins = { top: 96, right: 96, bottom: 96, left: 96 }

function fixture() {
  return {
    section: { pageWidthPx: 1056, pageHeightPx: 816, marginPx: margins, header: [paragraph('h2', 'SECOND_HEADER')], footer: [paragraph('f2', [run('Page {page} of {pages}')], { align: 'center' })] },
    lists: {
      bullets: { id: 'bullets', levels: [{ format: 'bullet', bulletChar: '•', text: '', indentLeftPx: 24, hangingPx: 18, start: 1 }, { format: 'bullet', bulletChar: '◦', text: '', indentLeftPx: 48, hangingPx: 18, start: 1 }] },
      numbers: { id: 'numbers', levels: [{ format: 'decimal', text: '%1.', indentLeftPx: 24, hangingPx: 18, start: 1 }, { format: 'lowerLetter', text: '%1.%2)', indentLeftPx: 48, hangingPx: 18, start: 1 }] },
    },
    footnotes: { fn1: [paragraph('fn', 'FOOTNOTE_TEXT')] },
    endnotes: { en1: [paragraph('en', 'ENDNOTE_TEXT')] },
    bookmarks: { target: { start: { blockId: 'heading', offset: 0 }, end: { blockId: 'heading', offset: 7 } } },
    blocks: [
      paragraph('heading', [run('Heading One', { bold: true, fontSizePx: 24, fontFamily: 'Georgia, serif' })], { outlineLevel: 0 }),
      paragraph('format', [
        run('Plain '), run('bold', { bold: true }), run(' '), run('italic', { italic: true }), run(' '), run('under', { underline: true }),
        run(' '), run('strike', { strikethrough: true }), run(' '), run('red', { color: '#c00000' }), run(' '), run('mark', { highlightColor: '#ffff00' }),
        run(' x'), run('2', { verticalAlign: 'super' }), run(' H'), run('2', { verticalAlign: 'sub' }), run('O '), run('caps', { caps: true }),
        run(' Привет 中文 😀 tab\there soft\vbreak {braces} back\\slash'),
        run('1', { footnoteRef: 'fn1' }), run('i', { endnoteRef: 'en1' }),
      ], { align: 'justify', spaceBeforePx: 8 }),
      paragraph('rtl', [run('مرحبا بالعالم', { rtl: true })], { direction: 'rtl', align: 'right' }),
      paragraph('b1', 'BULLET_ONE', { list: { listId: 'bullets', level: 0 } }),
      paragraph('b2', 'BULLET_NESTED', { list: { listId: 'bullets', level: 1 } }),
      paragraph('n1', 'NUMBER_ONE', { list: { listId: 'numbers', level: 0 } }),
      paragraph('n2', 'NUMBER_SUB', { list: { listId: 'numbers', level: 1 } }),
      paragraph('n3', 'NUMBER_TWO', { list: { listId: 'numbers', level: 0 } }),
      paragraph('links', [run('Visit '), run('the site', { link: 'https://example.com/a b' }), run(' or '), run('the heading', { link: '#target' })]),
      {
        kind: 'table', id: 'table', revision: 1, colFractions: [0.5, 0.25, 0.25],
        rows: [
          { props: { repeatHeader: true }, cells: [{ id: 'c1', blocks: [paragraph('c1p', 'HEAD_WIDE')], colSpan: 2, shading: '#d9e2f3' }, { id: 'c2', blocks: [paragraph('c2p', 'HEAD_3')] }] },
          { cells: [{ id: 'c3', blocks: [paragraph('c3p', 'TALL_CELL')], rowSpan: 2, borders: { top: { color: '#ff0000', widthPx: 2, style: 'double' } } }, { id: 'c4', blocks: [paragraph('c4p', 'R2C2')] }, { id: 'c5', blocks: [paragraph('c5p', 'R2C3')] }] },
          { cells: [{ id: 'c6', blocks: [paragraph('c6p', 'R3C2')] }, { id: 'c7', blocks: [paragraph('c7p', 'R3C3'), paragraph('c7q', 'SECOND_PARAGRAPH_IN_CELL')] }] },
        ],
      },
      { kind: 'image', id: 'png', revision: 1, src: `data:image/png;base64,${PNG.toString('base64')}`, widthPx: 120, heightPx: 60, align: 'center' },
      { kind: 'image', id: 'gif', revision: 1, src: `data:image/gif;base64,${GIF.toString('base64')}`, widthPx: 20, heightPx: 20, align: 'left' },
      { kind: 'equation', id: 'eq', revision: 1, equation: { display: true, root: { type: 'row', children: [{ type: 'ident', text: 'E' }, { type: 'op', text: '=' }, { type: 'ident', text: 'mc' }] } } },
      { kind: 'shape', id: 'shape', revision: 1, geometry: { preset: 'rect' }, widthPx: 100, heightPx: 50, align: 'left', text: { blocks: [paragraph('sp', 'SHAPE_TEXT')] } },
      { kind: 'custom', id: 'custom', revision: 1, customType: 'chart', data: {} },
      paragraph('end1', 'FIRST_SECTION_END', {
        sectionBreak: { type: 'nextPage', props: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: margins, header: [paragraph('h1', 'FIRST_HEADER')], headerFirst: [paragraph('hf', 'TITLE_PAGE_HEADER')] } },
      }),
      paragraph('landscape', 'LANDSCAPE_SECTION'),
    ],
  }
}

// Strict enough to catch what ODF consumers reject: nesting, attributes,
// entities, stray markup and undeclared namespace prefixes.
function assertWellFormedXml(xml, label) {
  const source = xml.replace(/^﻿/, '')
  const stack = []
  let roots = 0
  const declared = new Set(['xml'])
  const token = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)((?:\s+[A-Za-z_][\w.:-]*\s*=\s*(?:"[^"<]*"|'[^'<]*'))*)\s*(\/?)>|<|&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/g
  for (const match of source.matchAll(token)) {
    const text = match[0]
    assert.notEqual(text, '<', `${label}: stray "<" at ${match.index}: ${source.slice(match.index, match.index + 60)}`)
    assert.ok(!text.startsWith('&'), `${label}: invalid entity at ${match.index}: ${source.slice(match.index, match.index + 30)}`)
    if (/^<(?:!--|\?|!\[CDATA\[)/.test(text)) continue
    const [, closing, name, attributes, selfClosing] = match
    const attributeNames = [...attributes.matchAll(/([A-Za-z_][\w.:-]*)\s*=/g)].map((item) => item[1])
    assert.equal(new Set(attributeNames).size, attributeNames.length, `${label}: duplicate attribute on <${name}>`)
    for (const item of attributeNames) if (item.startsWith('xmlns:')) declared.add(item.slice(6))
    for (const qualified of [name, ...attributeNames.filter((item) => !item.startsWith('xmlns'))]) {
      if (qualified.includes(':')) assert.ok(declared.has(qualified.split(':')[0]), `${label}: undeclared prefix in ${qualified}`)
    }
    if (closing) assert.equal(stack.pop(), name, `${label}: mismatched </${name}> at ${match.index}`)
    else if (!selfClosing) {
      if (!stack.length) roots += 1
      stack.push(name)
    } else if (!stack.length) roots += 1
  }
  assert.deepEqual(stack, [], `${label}: unclosed elements`)
  assert.equal(roots, 1, `${label}: exactly one root element`)
}

function rtfDepthBalanced(text) {
  let depth = 0
  for (let index = 0; index < text.length; index += 1) {
    if (text[index] === '\\') { index += 1; continue }
    if (text[index] === '{') depth += 1
    if (text[index] === '}') depth -= 1
    if (depth < 0) return false
  }
  return depth === 0
}

test('RTF export writes structure that RTF readers rebuild (DOC-SIE-24)', async () => {
  const { documentToRtf } = await import('../src/exporters/rtf.js')
  const warnings = []
  const rtf = documentToRtf(fixture(), { title: 'Fixture <title>', warnings })
  assert.match(rtf, /^\{\\rtf1\\ansi/)
  assert.ok(rtfDepthBalanced(rtf), 'braces are balanced')
  assert.ok(/^[\x09\x0a\x0d\x20-\x7e]*$/.test(rtf), 'RTF stays 7-bit ASCII with \\u escapes')
  assert.doesNotThrow(() => validateExportContent('rtf', Buffer.from(rtf, 'latin1')))
  assert.match(rtf, /\{\\title Fixture <title>\}/)
  assert.match(rtf, /\\sect\b/)
  assert.match(rtf, /\\lndscpsxn/, 'the landscape section keeps its orientation')
  assert.match(rtf, /\{\\headerf /)
  assert.match(rtf, /\\titlepg/)
  assert.match(rtf, /\{\\field\{\\\*\\fldinst\{[^}]*PAGE\}\}/, 'footer page numbers are live fields')
  assert.match(rtf, /HYPERLINK \\\\l "target"/)
  assert.match(rtf, /\{\\\*\\bkmkstart target\}/)
  assert.match(rtf, /\\clvmgf/)
  assert.match(rtf, /\\clvmrg/)
  assert.match(rtf, /\\trhdr/)
  assert.match(rtf, /\\pngblip/)
  assert.match(rtf, /\\ftnalt/, 'endnotes are marked as endnotes')
  assert.match(rtf, /\{\\\*\\listtable/)
  assert.match(warnings.join(' '), /1 picture is not PNG or JPEG/)
  assert.match(warnings.join(' '), /1 object that RTF cannot represent/)

  const flow = parseRtf(Buffer.from(rtf, 'latin1'))
  const text = flowText(flow.blocks)
  for (const marker of ['Heading One', 'Plain bold italic under strike red mark', 'Привет 中文 😀', 'tab\there soft\nbreak {braces} back\\slash', 'مرحبا بالعالم', 'BULLET_NESTED', 'NUMBER_SUB', 'the site', 'HEAD_WIDE', 'TALL_CELL', 'SECOND_PARAGRAPH_IN_CELL', 'E=mc', 'SHAPE_TEXT', 'FIRST_SECTION_END', 'LANDSCAPE_SECTION', 'FOOTNOTE_TEXT', 'ENDNOTE_TEXT']) {
    assert.ok(text.includes(marker), `RTF keeps ${JSON.stringify(marker)}`)
  }
  const heading = flow.blocks.find((block) => flowText([block]) === 'Heading One')
  assert.equal(heading.heading, 1)
  const formatted = flow.blocks.find((block) => flowText([block]).startsWith('Plain '))
  assert.ok(formatted.runs.some((item) => item.text === 'bold' && item.bold))
  assert.ok(formatted.runs.some((item) => item.text === 'red' && item.color === '#c00000'))
  assert.ok(flow.blocks.some((block) => block.runs?.some((item) => item.text === 'the site' && item.link === 'https://example.com/a%20b')))
  assert.equal(flow.blocks.filter((block) => block.type === 'table').length, 1)
  assert.equal(flow.blocks.filter((block) => block.runs?.some((item) => item.image)).length, 1)
  const numbered = flow.blocks.find((block) => flowText([block]) === 'NUMBER_ONE')
  assert.equal(numbered.list.ordered, true)
  assert.equal(flow.blocks.find((block) => flowText([block]) === 'BULLET_ONE').list.ordered, false)
  // Each section keeps its own header; a reader shows the first one.
  assert.equal(flowText(flow.header), 'FIRST_HEADER')
  assert.match(rtf, /\{\\header [^]*?SECOND_HEADER/)
})

test('ODT export is a valid OpenDocument package with the document structure (DOC-SIE-24)', async () => {
  const { documentToOdt } = await import('../src/exporters/odt.js')
  const warnings = []
  const bytes = Buffer.from(await documentToOdt(fixture(), { title: 'Fixture & title', warnings }))
  assert.doesNotThrow(() => validateExportContent('odt', bytes))
  // The first local entry is the uncompressed mimetype.
  assert.equal(bytes.readUInt32LE(0), 0x04034b50)
  assert.equal(bytes.readUInt16LE(8), 0)
  assert.equal(bytes.subarray(30, 38).toString(), 'mimetype')
  const zip = await JSZip.loadAsync(bytes)
  assert.equal(await zip.file('mimetype').async('string'), ODT_MIME_TYPE)
  assert.ok(!Object.keys(zip.files).some((name) => name.endsWith('/')), 'no directory entries')
  const manifest = await zip.file('META-INF/manifest.xml').async('string')
  for (const name of Object.keys(zip.files).filter((entry) => entry !== 'mimetype' && entry !== 'META-INF/manifest.xml')) {
    assert.match(manifest, new RegExp(`manifest:full-path="${name.replace(/[.]/g, '\\.')}"`), `manifest lists ${name}`)
  }
  for (const name of ['content.xml', 'styles.xml', 'meta.xml', 'META-INF/manifest.xml']) assertWellFormedXml(await zip.file(name).async('string'), name)
  const content = await zip.file('content.xml').async('string')
  const styles = await zip.file('styles.xml').async('string')
  assert.match(await zip.file('meta.xml').async('string'), /<dc:title>Fixture &amp; title<\/dc:title>/)
  const plain = content.replace(/<text:note-citation>[^<]*<\/text:note-citation>/g, '').replace(/<text:tab\/>/g, '\t').replace(/<text:line-break\/>/g, '\n').replace(/<text:s text:c="(\d+)"\/>/g, (_m, count) => ' '.repeat(Number(count))).replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
  for (const marker of ['Heading One', 'Привет 中文 😀', 'tab\there soft\nbreak {braces} back\\slash', 'مرحبا بالعالم', 'BULLET_NESTED', 'NUMBER_SUB', 'HEAD_WIDE', 'TALL_CELL', 'SECOND_PARAGRAPH_IN_CELL', 'E=mc', 'SHAPE_TEXT', 'FIRST_SECTION_END', 'LANDSCAPE_SECTION', 'FOOTNOTE_TEXT', 'ENDNOTE_TEXT']) {
    assert.ok(plain.includes(marker), `ODT keeps ${JSON.stringify(marker)}`)
  }
  assert.match(content, /<text:h text:style-name="[^"]+" text:outline-level="1"><text:bookmark-start text:name="target"\/>/)
  const item = (value) => `<text:p[^>]*>(?:<text:span[^>]*>)?${value}(?:</text:span>)?</text:p>`
  assert.match(content, new RegExp(`<text:list text:style-name="L1"><text:list-item>${item('BULLET_ONE')}<text:list><text:list-item>${item('BULLET_NESTED')}</text:list-item></text:list></text:list-item></text:list>`))
  assert.match(content, /<text:a xlink:type="simple" xlink:href="#target">/)
  assert.match(content, /<text:a xlink:type="simple" xlink:href="https:\/\/example\.com\/a%20b">/)
  assert.match(content, /table:number-columns-spanned="2"/)
  assert.match(content, /table:number-rows-spanned="2"/)
  assert.equal((content.match(/<table:covered-table-cell\/>/g) || []).length, 2)
  assert.match(content, /<table:table-header-rows>/)
  assert.match(content, /fo:background-color="#d9e2f3"/)
  assert.match(content, /fo:border-top="1\.50pt double #ff0000"/)
  assert.equal((content.match(/<draw:image /g) || []).length, 2, 'PNG and GIF pictures are both embedded')
  assert.ok(Object.keys(zip.files).includes('Pictures/image1.png') && Object.keys(zip.files).includes('Pictures/image2.gif'))
  assert.match(content, /text:note-class="footnote"/)
  assert.match(content, /text:note-class="endnote"/)
  assert.match(content, /style:text-position="super 58%"/)
  assert.match(content, /fo:background-color="#ffff00"/)
  assert.match(content, /style:writing-mode="rl-tb"/)
  // One master page per section with its own size, orientation and header.
  assert.equal((styles.match(/<style:master-page /g) || []).length, 2)
  assert.match(styles, /fo:page-width="8\.5000in" fo:page-height="11\.0000in" style:print-orientation="portrait"/)
  assert.match(styles, /fo:page-width="11\.0000in" fo:page-height="8\.5000in" style:print-orientation="landscape"/)
  assert.match(styles, /<style:header-first>[\s\S]*TITLE_PAGE_HEADER/)
  assert.match(styles, /<text:page-number text:select-page="current">1<\/text:page-number> of <text:page-count>1<\/text:page-count>/)
  assert.match(content, /style:master-page-name="Section2"/)
  assert.match(warnings.join(' '), /1 object that ODT cannot represent/)
})

async function complexDocumentCopy() {
  try {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-doc-complex-'))
    const copy = path.join(directory, 'Complex document.docx')
    await fs.copyFile(COMPLEX_DOCUMENT, copy)
    return { directory, copy }
  } catch {
    return null
  }
}

test('RTF and ODT exports of a copy of Complex document.docx open as valid files', async (t) => {
  const sample = await complexDocumentCopy()
  if (!sample) { t.skip('Simple test examples/Complex document.docx is not available'); return }
  try {
    const { runImport } = await import('@forevka/wordcanvas/import')
    const { documentToRtf } = await import('../src/exporters/rtf.js')
    const { documentToOdt } = await import('../src/exporters/odt.js')
    const imported = runImport(new Uint8Array(await fs.readFile(sample.copy)), undefined, { collectMediaBytes: true })
    const document = structuredClone(imported.doc)
    // Pictures arrive as data: URIs, as prepareDocumentForExport provides them in the editor.
    const embed = (value) => {
      if (!value || typeof value !== 'object') return
      if (value.kind === 'image' && String(value.src).startsWith('cw-media:')) {
        const media = imported.media[value.src.slice('cw-media:'.length)]
        value.src = `data:${media.mime};base64,${Buffer.from(media.bytes).toString('base64')}`
      }
      for (const child of Object.values(value)) embed(child)
    }
    embed(document)
    const markers = ['STRESS_START', 'Mixed Arial blue', 'Ångström', 'TABLE_ONE_5', 'TABLE_ONE_9', 'TABLE_TWO_4', 'BULLET_TWO', 'PARA_16', 'العربية', 'עברית', '中文测试', '한국어', 'STRESS_SECTION_TWO', 'LANDSCAPE_MARKER', 'STRESS_END', 'SOFT_BREAK_LINE', 'FOOTNOTE_CONTENT']

    const rtfWarnings = []
    const rtf = documentToRtf(document, { title: 'Complex document', warnings: rtfWarnings })
    assert.deepEqual(rtfWarnings, [])
    assert.match(rtf, /^\{\\rtf1/)
    assert.ok(rtfDepthBalanced(rtf))
    assert.doesNotThrow(() => validateExportContent('rtf', Buffer.from(rtf, 'latin1')))
    const flow = parseRtf(Buffer.from(rtf, 'latin1'))
    const rtfText = flowText(flow.blocks)
    for (const marker of markers) assert.ok(rtfText.includes(marker), `RTF keeps ${marker}`)
    assert.equal(flow.blocks.filter((block) => block.type === 'table').length, 2)
    assert.equal(flow.blocks.filter((block) => block.runs?.some((item) => item.image)).length, 1)
    assert.equal(flow.footnoteCount, 1)
    assert.match(rtf, /\\lndscpsxn/)

    const odt = Buffer.from(await documentToOdt(document, { title: 'Complex document' }))
    assert.doesNotThrow(() => validateExportContent('odt', odt))
    const zip = await JSZip.loadAsync(odt)
    for (const name of ['content.xml', 'styles.xml', 'meta.xml', 'META-INF/manifest.xml']) assertWellFormedXml(await zip.file(name).async('string'), name)
    const content = await zip.file('content.xml').async('string')
    const plain = content.replace(/<text:line-break\/>/g, '\n').replace(/<[^>]+>/g, ' ')
    for (const marker of markers) assert.ok(plain.includes(marker), `ODT keeps ${marker}`)
    assert.equal((content.match(/<table:table /g) || []).length, 2)
    assert.equal((content.match(/<draw:image /g) || []).length, 1)
    assert.equal((content.match(/<text:note /g) || []).length, 1)
    assert.equal((await zip.file('styles.xml').async('string')).match(/<style:master-page /g).length, 2)
    // The supplied original is never modified.
    assert.deepEqual(await fs.readFile(COMPLEX_DOCUMENT), await fs.readFile(sample.copy))
  } finally {
    await fs.rm(sample.directory, { recursive: true, force: true })
  }
})
