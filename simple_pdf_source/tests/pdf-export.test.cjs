'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const test = require('node:test')
const JSZip = require('jszip')
const mammoth = require('mammoth')
const ts = require('typescript')
const { DOMParser } = require('@xmldom/xmldom')
const {
  buildTextExport,
  buildTextExportFiles,
  imageExportFileName,
  normalizeTextPages,
  safeExportBaseName,
} = require('../electron/pdf-export.cjs')

const SAMPLE_PDF = path.join(__dirname, '..', '..', 'Simple test examples', 'Complex document.pdf')

function rendererExportHelpers() {
  const sourcePath = path.join(__dirname, '..', 'src', 'lib', 'pdfExport.ts')
  const output = ts.transpileModule(fs.readFileSync(sourcePath, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
  }).outputText
  const loaded = { exports: {} }
  Function('module', 'exports', 'require', output)(loaded, loaded.exports, require)
  return loaded.exports
}

function assertWellFormed(xml, label) {
  const errors = []
  new DOMParser({ onError: (level, message) => { if (level !== 'warning') errors.push(message) } }).parseFromString(xml, 'application/xml')
  assert.deepEqual(errors, [], `${label} is not well-formed XML`)
}

async function docxParts(bytes) {
  const zip = await JSZip.loadAsync(bytes)
  const parts = {}
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir) continue
    if (/\.(xml|rels)$/.test(name)) {
      parts[name] = await zip.file(name).async('string')
      assertWellFormed(parts[name], name)
    } else parts[name] = await zip.file(name).async('nodebuffer')
  }
  return parts
}

const count = (text, pattern) => (text.match(pattern) || []).length

// A 2 x 2 PNG (red, blue / blue, red).
const TINY_PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP4z8DwHwyBNMN/BgYGBgYAN/8F+1Vb5ZwAAAAASUVORK5CYII=', 'base64')

const pages = [
  { pageNumber: 1, text: 'First page\nA < B & C' },
  { pageNumber: 3, text: 'Third page\nSecond line' },
]

/** Pages as the renderer's layout model sends them. */
function structuredPages() {
  const run = (text, extra = {}) => ({ text, size: 11, font: 'Georgia', ...extra })
  return [
    {
      pageNumber: 1, width: 612, height: 792, text: 'Report title',
      blocks: [
        { type: 'heading', level: 1, lines: [{ runs: [run('Report', { bold: true, size: 24 })] }, { runs: [run('title', { bold: true, size: 24 })] }] },
        { type: 'paragraph', lines: [
          { runs: [run('Plain '), run('bold', { bold: true }), run(' and '), run('italic', { italic: true }), run(' text that'), run('1', { superscript: true, size: 7 })] },
          { runs: [run('continues here.')], hardBreak: true },
          { runs: [run('A forced break & < markup.')] },
        ] },
        { type: 'paragraph', list: 'bullet', lines: [{ runs: [run('First bullet')] }] },
        { type: 'paragraph', list: 'bullet', lines: [{ runs: [run('Second bullet')] }] },
        { type: 'table', widths: [120, 200], rows: [
          [[{ runs: [run('Name', { bold: true })] }], [{ runs: [run('Amount', { bold: true })] }]],
          [[{ runs: [run('Tea | leaves')] }], [{ runs: [run('1,234.50')] }, { runs: [run('(estimated)')] }]],
        ] },
        { type: 'image', data: TINY_PNG, mime: 'image/png', width: 144, height: 144 },
        { type: 'paragraph', lines: [{ runs: [run('This sentence runs on')] }] },
      ],
    },
    {
      pageNumber: 2, width: 792, height: 612, text: 'onto the wide page.',
      blocks: [
        { type: 'paragraph', lines: [{ runs: [run('onto the wide page.')] }] },
        { type: 'heading', level: 2, lines: [{ runs: [run('Wide section', { bold: true, size: 18 })] }] },
      ],
    },
  ]
}

test('export names stay local to the chosen destination', () => {
  assert.equal(safeExportBaseName('Quarterly: Report?.pdf'), 'Quarterly Report')
  assert.equal(safeExportBaseName('Quarterly.v2'), 'Quarterly.v2')
  assert.equal(safeExportBaseName('Quarterly.V2.PDF'), 'Quarterly.V2')
  assert.equal(imageExportFileName('Quarterly Report', 3, 12, 'png'), 'Quarterly Report - page 003.png')
  assert.equal(imageExportFileName('Quarterly Report', 3, 12, 'jpeg'), 'Quarterly Report - page 003.jpg')
  assert.equal(imageExportFileName('Quarterly Report', 3, 12, 'webp'), 'Quarterly Report - page 003.webp')
})

test('text page validation rejects empty and unreasonable inputs and bounds the model', () => {
  assert.throws(() => normalizeTextPages([]), /between 1 and 10,000/i)
  assert.throws(() => normalizeTextPages([{ pageNumber: 0, text: '' }]), /invalid page number/i)
  assert.deepEqual(normalizeTextPages(pages), pages)
  const [page] = normalizeTextPages([{ pageNumber: 1, text: 'x', width: 612, height: 792, blocks: [
    { type: 'image', data: Buffer.from('GIF89a-not-allowed'), width: 10, height: 10 },
    { type: 'script', lines: [] },
    { type: 'paragraph', lines: [{ runs: [{ text: 'Safe\u0001 text', size: 1e9, font: 'Evil"<Font>', bold: 1 }] }] },
    { type: 'heading', level: 9, lines: [{ runs: [{ text: 'Top' }] }] },
  ] }])
  assert.deepEqual(page.blocks, [
    { type: 'paragraph', lines: [{ runs: [{ text: 'Safe text', bold: true, font: 'EvilFont' }] }] },
    { type: 'heading', level: 3, lines: [{ runs: [{ text: 'Top' }] }] },
  ])
})

test('plain text keeps every line, a blank line between paragraphs and a form feed between pages', async () => {
  const text = (await buildTextExport('txt', pages, 'Sample')).toString('utf8')
  assert.equal(text, 'First page\nA < B & C\n\f\nThird page\nSecond line\n')
  const structured = (await buildTextExport('txt', structuredPages(), 'Report')).toString('utf8')
  assert.equal(structured.split('\f')[0], [
    'Report title',
    '',
    'Plain bold and italic text that1',
    'continues here.',
    'A forced break & < markup.',
    '',
    '• First bullet',
    '',
    '• Second bullet',
    '',
    'Name\tAmount',
    'Tea | leaves\t1,234.50 (estimated)',
    '',
    'This sentence runs on',
    '',
  ].join('\n'))
})

test('web pages carry headings, emphasis, lists, tables, pictures and escaped text', async () => {
  const html = (await buildTextExport('html', pages, 'Sample')).toString('utf8')
  assert.match(html, /<section class="page" data-page="1">/)
  assert.match(html, /A &lt; B &amp; C/)
  assert.doesNotMatch(html, /A < B & C/)
  const structured = (await buildTextExport('html', structuredPages(), 'Report')).toString('utf8')
  assert.match(structured, /<h1>Report title<\/h1>/)
  assert.match(structured, /Plain <strong>bold<\/strong> and <em>italic<\/em> text that<sup>1<\/sup> continues here\.<br>A forced break &amp; &lt; markup\./)
  assert.match(structured, /<ul>\n<li>First bullet<\/li>\n<li>Second bullet<\/li>\n<\/ul>/)
  assert.match(structured, /<table><colgroup><col style="width:120pt"><col style="width:200pt"><\/colgroup><tr><td><strong>Name<\/strong><\/td>/)
  assert.match(structured, /<td>1,234\.50 \(estimated\)<\/td>/)
  assert.match(structured, /<img src="data:image\/png;base64,/)
  assert.match(structured, /<h2>Wide section<\/h2>/)
})

test('Markdown has headings, emphasis, lists, tables, picture files and flows across pages', async () => {
  const markdown = (await buildTextExport('md', pages, 'Sample')).toString('utf8')
  assert.equal(markdown, 'First page\\\nA \\< B & C\n\nThird page\\\nSecond line\n')
  const files = await buildTextExportFiles('md', structuredPages(), 'Report', { imageFolder: 'Report images' })
  const text = files.data.toString('utf8')
  assert.match(text, /^# Report title$/m)
  assert.match(text, /^Plain \*\*bold\*\* and \*italic\* text that<sup>1<\/sup>$/m)
  assert.match(text, /^continues here\.\\$/m)
  assert.match(text, /^- First bullet\n\n- Second bullet$/m)
  assert.match(text, /^\| \*\*Name\*\* \| \*\*Amount\*\* \|\n\| --- \| --- \|\n\| Tea \\\| leaves \| 1,234\.50 \(estimated\) \|$/m)
  assert.match(text, /^!\[\]\(<Report images\/image-001\.png>\)$/m)
  assert.deepEqual(files.assets.map((asset) => asset.name), ['image-001.png'])
  assert.deepEqual(files.assets[0].data, TINY_PNG)
  // The paragraph cut by the page break is one paragraph again.
  assert.match(text, /^This sentence runs on\nonto the wide page\.$/m)
  assert.match(text, /^## Wide section$/m)
  const embedded = (await buildTextExport('md', structuredPages(), 'Report')).toString('utf8')
  assert.match(embedded, /!\[\]\(data:image\/png;base64,/)
})

test('Word export is a valid package with styles, runs, lists, tables, pictures and sections', async () => {
  const bytes = await buildTextExport('docx', structuredPages(), 'Sample & Review')
  const parts = await docxParts(bytes)
  const documentXml = parts['word/document.xml']
  assert.match(parts['docProps/core.xml'], /<dc:title>Sample &amp; Review<\/dc:title>/)
  assert.match(documentXml, /<w:pStyle w:val="Heading1"\/><\/w:pPr><w:r><w:t xml:space="preserve">Report title<\/w:t><\/w:r>/)
  assert.match(documentXml, /<w:rPr><w:b\/><w:bCs\/><\/w:rPr><w:t xml:space="preserve">bold<\/w:t>/)
  assert.match(documentXml, /<w:rPr><w:i\/><w:iCs\/><\/w:rPr><w:t xml:space="preserve">italic<\/w:t>/)
  assert.match(documentXml, /<w:vertAlign w:val="superscript"\/><\/w:rPr><w:t xml:space="preserve">1<\/w:t>/)
  assert.match(documentXml, /continues here\.<\/w:t><\/w:r><w:r><w:br\/><\/w:r>/)
  assert.match(documentXml, /A forced break &amp; &lt; markup\./)
  assert.equal(count(documentXml, /<w:numId w:val="1"\/>/g), 2)
  assert.match(parts['word/numbering.xml'], /<w:numFmt w:val="bullet"\/>/)
  assert.equal(count(documentXml, /<w:tbl>/g), 1)
  assert.match(documentXml, /<w:tcW w:w="2400" w:type="dxa"\/>/)
  assert.match(documentXml, /Tea \| leaves/)
  assert.match(documentXml, /<a:blip r:embed="rIdImage1"\/>/)
  assert.match(parts['word/_rels/document.xml.rels'], /Target="media\/image1\.png"/)
  assert.deepEqual(parts['word/media/image1.png'], TINY_PNG)
  assert.match(parts['[Content_Types].xml'], /Extension="png"/)
  // The wide page is its own landscape section.
  assert.match(documentXml, /<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"\/>/)
  assert.match(documentXml, /<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>.*<\/w:sectPr><\/w:body>/)
  assert.match(parts['word/styles.xml'], /w:styleId="Heading1".*<w:sz w:val="48"\/>/)
  const html = (await mammoth.convertToHtml({ buffer: bytes })).value
  assert.match(html, /<h1>Report title<\/h1>/)
  assert.match(html, /<strong>bold<\/strong>/)
  assert.match(html, /<ul><li>First bullet<\/li><li>Second bullet<\/li><\/ul>/)
  assert.match(html, /<table>/)
  assert.match(html, /<img src="data:image\/png;base64,/)
  assert.match(html, /<h2>Wide section<\/h2>/)

  const plain = await buildTextExport('docx', [...pages, { pageNumber: 4, text: 'Safe\u0001 text\u000b remains' }], 'Plain')
  const plainXml = (await docxParts(plain))['word/document.xml']
  assert.equal(count(plainXml, /<w:br w:type="page"\/>/g), 2)
  assert.match(plainXml, /A &lt; B &amp; C/)
  assert.doesNotMatch(plainXml, /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/)
  const extracted = await mammoth.extractRawText({ buffer: plain })
  assert.match(extracted.value, /First page/)
  assert.match(extracted.value, /Safe text remains/)
})

test('plain-text extraction keeps pdf.js line ends, including empty end-of-line items', () => {
  const { textContentToPlainText } = rendererExportHelpers()
  assert.equal(textContentToPlainText([{ str: 'A', hasEOL: false }, { str: '', hasEOL: true }, { str: 'B' }]), 'A\nB')
  const at = (str, x, y, width = str.length * 6) => ({ str, transform: [12, 0, 0, 12, x, y], width, height: 12 })
  assert.equal(textContentToPlainText([at('First', 72, 700), at('line', 110, 700), at('Second', 72, 684)]), 'First line\nSecond')
  // Glyph runs split by kerning stay one word.
  assert.equal(textContentToPlainText([at('Ker', 72, 700, 18), at('ning', 90, 700)]), 'Kerning')
})

test('the layout model finds headings, emphasis, paragraphs, lists, tables and page numbers', () => {
  const { buildPageLayout, finishLayout } = rendererExportHelpers()
  const item = (str, x, y, { size = 12, font = 'body', width, eol = false } = {}) => ({
    str, transform: [size, 0, 0, size, x, y], width: width ?? str.length * size * 0.5, height: size, fontName: font, hasEOL: eol,
  })
  const eol = { str: '', hasEOL: true, transform: [12, 0, 0, 12, 0, 0], width: 0, height: 0 }
  const items = [
    item('Annual', 72, 720, { size: 24, font: 'bold' }), eol,
    item('Report', 72, 692, { size: 24, font: 'bold' }), eol,
    item('Intro with ', 72, 650), item('emphasis', 138, 650, { font: 'italic' }), item(' and', 186, 650), eol,
    item('a second line of the same paragraph that is long enough to fill.', 72, 636), eol,
    item('Next paragraph after a gap.', 72, 610), eol,
    item('• First point', 90, 584), eol,
    item('• Second point', 90, 570), eol,
    item('Name', 80, 540), item(' ', 104, 540, { width: 140 }), item('Total', 260, 540), item(' ', 290, 540, { width: 120 }), item('Note', 420, 540), eol,
    item('Tea', 80, 524), item(' ', 98, 524, { width: 150 }), item('12', 260, 524), item(' ', 272, 524, { width: 140 }), item('fresh', 420, 524), eol,
    item('leaves', 260, 510), eol,
    item('After the table.', 72, 480), item('2', 168, 485, { size: 7 }), eol,
    item('3', 300, 30),
  ]
  const draft = buildPageLayout({ pageNumber: 1, width: 612, height: 792, items, fonts: { body: 'ABCDEF+TimesNewRomanPSMT', bold: 'ABCDEF+Arial-BoldMT', italic: 'ABCDEF+TimesNewRomanPS-ItalicMT' } })
  const [page] = finishLayout([draft])
  const summary = page.blocks.map((block) => {
    if (block.type === 'table') return ['table', block.rows.map((row) => row.map((cell) => cell.map((line) => line.runs.map((run) => run.text).join('')).join(' ')))]
    return [block.type === 'heading' ? `h${block.level}` : block.list ? 'bullet' : 'p', block.lines.map((line) => line.runs.map((run) => run.text).join('')).join(' | ')]
  })
  assert.deepEqual(summary, [
    ['h1', 'Annual | Report'],
    ['p', 'Intro with emphasis and | a second line of the same paragraph that is long enough to fill.'],
    ['p', 'Next paragraph after a gap.'],
    ['bullet', 'First point'],
    ['bullet', 'Second point'],
    ['table', [['Name', 'Total', 'Note'], ['Tea', '12 leaves', 'fresh']]],
    ['p', 'After the table.2'],
  ])
  const heading = page.blocks[0]
  assert.equal(heading.lines[0].runs[0].bold, true)
  assert.equal(heading.lines[0].runs[0].font, 'Arial')
  const intro = page.blocks[1].lines[0].runs
  assert.deepEqual(intro.map((run) => [run.text, Boolean(run.italic), run.font]), [['Intro with ', false, 'Times New Roman'], ['emphasis', true, 'Times New Roman'], [' and', false, 'Times New Roman']])
  assert.equal(page.blocks.at(-1).lines[0].runs.at(-1).superscript, true)
  assert.match(page.text, /^Annual Report\n\nIntro with emphasis and\n/)
  assert.doesNotMatch(page.text, /\n3$/)
})

test('renderer text extraction releases a page after a failed text read', async () => {
  const { extractPdfText } = rendererExportHelpers()
  let cleanups = 0
  const pdf = {
    getPage: async () => ({
      getTextContent: async () => { throw new Error('broken text stream') },
      cleanup: () => { cleanups += 1 },
    }),
  }
  await assert.rejects(extractPdfText(pdf, [0]), /broken text stream/)
  assert.equal(cleanups, 1)
})

test('a sample PDF exports to Word, text and Markdown with its structure', { skip: !fs.existsSync(SAMPLE_PDF) && 'sample files are not available' }, async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-pdf-export-'))
  try {
    const copy = path.join(directory, 'Complex document.pdf')
    fs.copyFileSync(SAMPLE_PDF, copy)
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const document = await getDocument({ data: new Uint8Array(fs.readFileSync(copy)), isEvalSupported: false, disableFontFace: true }).promise
    let exported
    try {
      exported = await rendererExportHelpers().extractPdfText(document, [...Array(document.numPages).keys()])
    } finally {
      await document.destroy()
    }
    assert.equal(exported.length, 6)

    const text = (await buildTextExport('txt', exported, 'Complex document')).toString('utf8')
    const firstPage = text.split('\f')[0].trim().split('\n')
    assert.ok(firstPage.length >= 25, `page 1 has ${firstPage.length} lines`)
    assert.ok(Math.max(...firstPage.map((line) => line.length)) <= 300)
    assert.ok(firstPage.includes('Table one follows.'))
    assert.ok(firstPage.includes('PARA_01 Reliable documents preserve every sentence, table boundary and page break.'))
    assert.equal(count(text, /\f/g), 5)

    const word = await buildTextExport('docx', exported, 'Complex document')
    const parts = await docxParts(word)
    const xml = parts['word/document.xml']
    assert.ok(count(xml, /<w:drawing>/g) >= 1, 'the picture is kept')
    assert.ok(count(xml, /<w:pStyle w:val="Heading1"\/>/g) >= 1)
    assert.ok(count(xml, /<w:b\/>/g) >= 1)
    assert.ok(count(xml, /<w:i\/>/g) >= 1)
    // Five page boundaries: four page breaks and the landscape section break.
    assert.equal(count(xml, /<w:br w:type="page"\/>/g) + count(xml, /<w:p><w:pPr><w:sectPr>/g), 5)
    assert.match(xml, /w:orient="landscape"/)
    assert.ok(count(xml, /<w:tbl>/g) >= 2, 'both tables are tables')
    assert.ok(count(xml, /<w:p[ >]/g) >= 40, 'paragraphs are kept apart')
    const html = (await mammoth.convertToHtml({ buffer: word })).value
    assert.match(html, /<p>Table one follows\.<\/p>/)
    assert.match(html, /<h1>STRESS_START — Local editor creation<\/h1>/)
    assert.match(html, /<td><p>TABLE_ONE_5 Long wrapped cell content\. Long wrapped/)
    for (let number = 1; number <= 16; number += 1) assert.match(html, new RegExp(`<p>PARA_${String(number).padStart(2, '0')} Reliable`))

    const markdown = await buildTextExportFiles('md', exported, 'Complex document', { imageFolder: 'Complex document images' })
    const md = markdown.data.toString('utf8')
    assert.match(md, /^# STRESS\\_START — Local editor creation$/m)
    assert.match(md, /^## LANDSCAPE\\_MARKER$/m)
    assert.match(md, /^Mixed Arial blue \*\*bold text\*\* \*italic accent café naïve Ångström\*$/m)
    assert.match(md, /^\| TABLE\\_ONE\\_1 value \| TABLE\\_ONE\\_2 value \| TABLE\\_ONE\\_3 value \|$/m)
    assert.match(md, /^PARA\\_01 Reliable documents preserve every sentence, table boundary and page break\.\nReliable documents/m)
    assert.equal(markdown.assets.length, 1)
  } finally {
    fs.rmSync(directory, { recursive: true, force: true })
  }
})

test('renderer image conversion releases its page and canvas after an encoding setup failure', async () => {
  const { renderPdfPageImage } = rendererExportHelpers()
  let cleanups = 0
  const canvas = { width: 0, height: 0, getContext: () => null }
  const previousDocument = global.document
  global.document = { createElement: () => canvas }
  const pdf = {
    getPage: async () => ({
      rotate: 90,
      getViewport: ({ scale }) => ({ width: 100 * scale, height: 200 * scale }),
      cleanup: () => { cleanups += 1 },
    }),
  }
  try {
    await assert.rejects(renderPdfPageImage(pdf, 0, 'png', 2, 0.9), /canvas could not be created/i)
  } finally {
    global.document = previousDocument
  }
  assert.equal(cleanups, 1)
  assert.equal(canvas.width, 0)
  assert.equal(canvas.height, 0)
})

test('renderer image conversion applies intrinsic page rotation and requests the chosen encoding', async () => {
  const { renderPdfPageImage } = rendererExportHelpers()
  const viewportRequests = []
  const canvas = {
    width: 0,
    height: 0,
    getContext: () => ({}),
    toBlob: (callback, mimeType, quality) => {
      assert.equal(mimeType, 'image/webp')
      assert.equal(quality, 0.8)
      callback(new Blob([Uint8Array.of(1, 2, 3)], { type: mimeType }))
    },
  }
  let renderedViewport = null
  let cleanups = 0
  const previousDocument = global.document
  global.document = { createElement: () => canvas }
  const pdf = {
    getPage: async () => ({
      rotate: 270,
      getViewport: (options) => {
        viewportRequests.push(options)
        return { width: 120 * options.scale, height: 80 * options.scale }
      },
      render: ({ viewport }) => {
        renderedViewport = viewport
        return { promise: Promise.resolve() }
      },
      cleanup: () => { cleanups += 1 },
    }),
  }
  let result
  try {
    result = await renderPdfPageImage(pdf, 0, 'webp', 2, 0.8)
  } finally {
    global.document = previousDocument
  }
  assert.deepEqual([...result], [1, 2, 3])
  assert.deepEqual(viewportRequests.map((request) => request.rotation), [270, 270])
  assert.deepEqual(renderedViewport, { width: 240, height: 160 })
  assert.equal(cleanups, 1)
  assert.equal(canvas.width, 0)
  assert.equal(canvas.height, 0)
})

test('extremely large PDF pages respect image-export limits below 0.1x scale', async () => {
  const { renderPdfPageImage } = rendererExportHelpers()
  let rendered = false
  const canvas = {
    width: 0, height: 0, getContext: () => ({}),
    toBlob(callback) {
      assert.ok(this.width <= 8192 && this.height <= 8192)
      assert.ok(this.width * this.height <= 64_000_000)
      callback(new Blob([Uint8Array.of(1)], { type: 'image/png' }))
    },
  }
  const previousDocument = global.document
  global.document = { createElement: () => canvas }
  try {
    await renderPdfPageImage({ getPage: async () => ({
      rotate: 0,
      getViewport: ({ scale }) => ({ width: 200_000 * scale, height: 200_000 * scale }),
      render: () => { rendered = true; return { promise: Promise.resolve() } },
      cleanup() {},
    }) }, 0, 'png', 3, .9)
    assert.equal(rendered, true)
  } finally { global.document = previousDocument }
})
