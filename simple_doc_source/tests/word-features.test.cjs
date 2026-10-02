'use strict'
// Word features built on the WordCanvas model (DOC-011, DOC-014, DOC-018): page numbers,
// headers and footers, paragraph indents and spacing and the font list. The model
// operations are applied with the engine's own operation code (DocumentEditor.commit)
// and exported with its headless DOCX/PDF pipeline, so these tests check what Word and
// a PDF reader will see, not just the operations.
const test = require('node:test')
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { pdfPageTexts, bandText } = require('./pdf-text.cjs')

const headerFooter = () => import('../src/ui/header-footer.ts')
const paragraphFormat = () => import('../src/ui/paragraph-format.ts')
const fontList = () => import('../src/ui/font-list.ts')

const style = (extra = {}) => ({ fontFamily: 'Calibri', fontSizePx: 14.667, bold: false, italic: false, underline: false, strikethrough: false, color: '#111111', ...extra })
const paraStyle = (extra = {}) => ({ align: 'left', lineHeight: 1.15, spaceBeforePx: 0, spaceAfterPx: 10.667, indentFirstLinePx: 0, indentLeftPx: 0, ...extra })
const paragraph = (id, text, extra = {}, runStyle = {}) => ({ kind: 'paragraph', id, revision: 0, runs: [{ text, style: style(runStyle) }], style: paraStyle(extra) })
const margins = { top: 96, right: 96, bottom: 96, left: 96 }
const LONG = 'Reliable documents keep every sentence where it belongs, page after page. '.repeat(6)

/** A document of about `pages` Letter pages of body text. */
function document(pages = 3, section = {}) {
  const blocks = []
  for (let index = 0; index < pages * 9; index += 1) blocks.push(paragraph(`p${index}`, `PARA_${index} ${LONG}`))
  return { section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: margins, ...section }, blocks, stylesheet: { defaultStyleId: 'Normal', styles: [{ id: 'Normal', name: 'Normal', char: { fontFamily: 'Georgia, serif', fontSizePx: 16 }, para: {} }] } }
}

let sequence = 0
const ids = () => `t${sequence++}`

async function engine() {
  const { DocumentEditor } = await import('@forevka/wordcanvas/query')
  const { runExport } = await import('@forevka/wordcanvas/export')
  const { runImport } = await import('@forevka/wordcanvas/import')
  const { installMeasureHost } = await import('@forevka/wordcanvas/export/measure')
  await installMeasureHost()
  return {
    apply(doc, ops) {
      const editor = new DocumentEditor(doc)
      editor.commit(ops)
      return editor.doc
    },
    async pdf(doc) {
      return pdfPageTexts((await runExport(doc, 'pdf')).bytes)
    },
    async docx(doc) {
      return (await runExport(doc, 'docx')).bytes
    },
    reopen(bytes) {
      return runImport(new Uint8Array(bytes), undefined, { collectMediaBytes: true }).doc
    },
  }
}

const bandTexts = (doc, band) => (doc.section[band] ?? []).map((block) => block.runs.map((run) => run.text).join(''))

// ---------------------------------------------------------------------------------------
// Page numbers, headers and footers (DOC-011)

test('a page number goes into the footer as PAGE/NUMPAGES fields and shows on every page of the PDF', async () => {
  const { insertPageNumberOps, currentPageNumberOptions, documentHasPageNumbers } = await headerFooter()
  const wc = await engine()
  const source = document(3)
  const ops = insertPageNumberOps(source, { position: 'bottom', align: 'center', text: 'page-x-of-y', format: 'arabic', showOnFirstPage: true, startAt: null }, ids)
  assert.deepEqual(ops.map((op) => op.type), ['setField', 'setField', 'setSectionBand'])
  assert.equal(ops[0].def.instruction, ' PAGE \\* MERGEFORMAT ')
  assert.equal(ops[1].def.instruction, ' NUMPAGES \\* MERGEFORMAT ')
  const doc = wc.apply(source, ops)
  assert.deepEqual(bandTexts(doc, 'footer'), ['Page {page} of {pages}'])
  assert.equal(doc.section.footer[0].style.align, 'center')
  assert.equal(doc.section.footer[0].runs[1].style.fieldId, ops[0].id)
  assert.equal(doc.section.footer[0].runs[0].style.fontFamily, 'Calibri', 'the number uses the body text font, not the stylesheet default')
  assert.equal(documentHasPageNumbers(doc), true)
  assert.deepEqual(currentPageNumberOptions(doc), { position: 'bottom', align: 'center', text: 'page-x-of-y', format: 'arabic', showOnFirstPage: true, startAt: null })

  const pages = await wc.pdf(doc)
  assert.ok(pages.length >= 3, `three pages or more (${pages.length})`)
  pages.forEach((page, index) => {
    assert.equal(bandText(page, 'bottom'), `Page ${index + 1} of ${pages.length}`, `page ${index + 1} footer`)
    assert.equal(bandText(page, 'top'), '', 'nothing in the header')
  })

  // Word sees real fields in a footer part.
  const zip = await JSZip.loadAsync(await wc.docx(doc))
  const footerName = Object.keys(zip.files).find((name) => /^word\/footer\d+\.xml$/.test(name))
  const footerXml = await zip.file(footerName).async('string')
  assert.match(footerXml, /w:instr=" PAGE \\\* MERGEFORMAT "/)
  assert.match(footerXml, /w:instr=" NUMPAGES \\\* MERGEFORMAT "/)
  assert.match(await zip.file('word/document.xml').async('string'), /<w:footerReference w:type="default"/)
  // And Simple reads them back as page numbers.
  const reopened = wc.reopen(await wc.docx(doc))
  assert.equal(documentHasPageNumbers(reopened), true)
  assert.equal(currentPageNumberOptions(reopened).text, 'page-x-of-y')
})

test('changing the page number never adds a second one; moving it to the header leaves other footer text', async () => {
  const { insertPageNumberOps, removePageNumberOps } = await headerFooter()
  const wc = await engine()
  let doc = document(2, { footer: [paragraph('f1', 'Confidential draft', { spaceAfterPx: 0 })] })
  doc = wc.apply(doc, insertPageNumberOps(doc, { position: 'bottom', align: 'right', text: 'number', format: 'arabic', showOnFirstPage: true, startAt: null }, ids))
  assert.deepEqual(bandTexts(doc, 'footer'), ['Confidential draft', '{page}'], 'existing footer text stays; the number gets its own paragraph')
  doc = wc.apply(doc, insertPageNumberOps(doc, { position: 'bottom', align: 'left', text: 'page-x', format: 'roman', showOnFirstPage: true, startAt: null }, ids))
  assert.deepEqual(bandTexts(doc, 'footer'), ['Confidential draft', 'Page {page:roman}'], 'the number paragraph is replaced, not duplicated')
  assert.equal(doc.section.footer[1].style.align, 'left')
  doc = wc.apply(doc, insertPageNumberOps(doc, { position: 'top', align: 'right', text: 'number', format: 'Roman', showOnFirstPage: true, startAt: null }, ids))
  assert.deepEqual(bandTexts(doc, 'footer'), ['Confidential draft'])
  assert.deepEqual(bandTexts(doc, 'header'), ['{page:Roman}'])
  const pages = await wc.pdf(doc)
  assert.equal(bandText(pages[0], 'top'), 'I')
  assert.equal(bandText(pages[1], 'top'), 'II')
  assert.equal(bandText(pages[1], 'bottom'), 'Confidential draft')
  doc = wc.apply(doc, removePageNumberOps(doc))
  assert.deepEqual(bandTexts(doc, 'header'), [''], 'Remove page numbers leaves an empty header')
  assert.deepEqual(bandTexts(doc, 'footer'), ['Confidential draft'])
})

test('Start at and Show on first page: the first page hides only the number, and numbering starts at 5', async () => {
  const { insertPageNumberOps, currentPageNumberOptions, hasDifferentFirstPage } = await headerFooter()
  const wc = await engine()
  let doc = document(3, { header: [paragraph('h1', 'Annual report', { spaceAfterPx: 0 })] })
  const ops = insertPageNumberOps(doc, { position: 'bottom', align: 'right', text: 'number', format: 'arabic', showOnFirstPage: false, startAt: 5 }, ids)
  assert.ok(ops.some((op) => op.type === 'setSectionProps' && op.geometry.pageNumberStart === 5))
  doc = wc.apply(doc, ops)
  assert.equal(hasDifferentFirstPage(doc), true)
  assert.deepEqual(bandTexts(doc, 'headerFirst'), ['Annual report'], 'the first page keeps the header text')
  assert.deepEqual(bandTexts(doc, 'footerFirst'), [''], 'but not the number')
  assert.notEqual(doc.section.headerFirst[0].id, doc.section.header[0].id, 'copies get their own ids')
  assert.deepEqual(currentPageNumberOptions(doc), { position: 'bottom', align: 'right', text: 'number', format: 'arabic', showOnFirstPage: false, startAt: 5 })
  const pages = await wc.pdf(doc)
  assert.equal(bandText(pages[0], 'bottom'), '', 'no number on the first page')
  assert.equal(bandText(pages[0], 'top'), 'Annual report')
  assert.equal(bandText(pages[1], 'bottom'), '6')
  assert.equal(bandText(pages[2], 'bottom'), '7')
  // Showing it on the first page again puts the number there too.
  doc = wc.apply(doc, insertPageNumberOps(doc, { position: 'bottom', align: 'right', text: 'number', format: 'arabic', showOnFirstPage: true, startAt: null }, ids))
  assert.equal(doc.section.pageNumberStart, undefined, 'blank Start at numbers from 1 again')
  const again = await wc.pdf(doc)
  assert.equal(bandText(again[0], 'bottom'), '1')
  assert.equal(bandText(again[1], 'bottom'), '2')
  // Word gets the title-page switch with both first-page parts.
  const xml = await (await JSZip.loadAsync(await wc.docx(doc))).file('word/document.xml').async('string')
  assert.match(xml, /<w:titlePg\/>/)
  assert.match(xml, /<w:footerReference w:type="first"/)
  assert.match(xml, /<w:headerReference w:type="first"/)
})

test('an earlier section with its own footer gets the number too; others inherit the last section\'s', async () => {
  const { insertPageNumberOps } = await headerFooter()
  const wc = await engine()
  const doc = document(1)
  doc.blocks[3] = paragraph('break', 'End of section one', { sectionBreak: { type: 'nextPage', props: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: margins, footer: [paragraph('f-own', 'Section one footer')] } } })
  const ops = insertPageNumberOps(doc, { position: 'bottom', align: 'center', text: 'number', format: 'arabic', showOnFirstPage: true, startAt: null }, ids)
  const breakOp = ops.find((op) => op.type === 'setParaStyle')
  assert.equal(breakOp.blockId, 'break')
  assert.deepEqual(breakOp.patch.sectionBreak.props.footer.map((block) => block.runs.map((run) => run.text).join('')), ['Section one footer', '{page}'])
  const applied = wc.apply(doc, ops)
  assert.deepEqual(bandTexts(applied, 'footer'), ['{page}'])
  const pages = await wc.pdf(applied)
  assert.match(bandText(pages[0], 'bottom'), /Section one footer\s+1/)
  assert.equal(bandText(pages.at(-1), 'bottom'), String(pages.length))
})

test('Edit header creates an empty header to type in; Remove header and Different first page are single steps', async () => {
  const { ensureBandOps, removeBandOps, differentFirstPageOps, hasBand, hasDifferentFirstPage, bandCaretTarget, bandCharStyle } = await headerFooter()
  const wc = await engine()
  let doc = document(1)
  assert.equal(hasBand(doc, 'header'), false)
  const ops = ensureBandOps(doc, 'header', ids)
  assert.equal(ops.length, 1)
  doc = wc.apply(doc, ops)
  assert.deepEqual(bandTexts(doc, 'header'), [''])
  assert.equal(doc.section.header[0].runs[0].style.fontFamily, 'Calibri')
  assert.deepEqual(doc.section.header[0].style, { align: 'left', lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 })
  assert.deepEqual(ensureBandOps(doc, 'header', ids), [], 'an existing header is kept')
  assert.deepEqual(bandCaretTarget(doc, 'header'), { blockId: doc.section.header[0].id, offset: 0 })
  assert.equal(bandCharStyle({ blocks: [], section: {}, stylesheet: { defaultStyleId: 'Normal', styles: [{ id: 'Normal', char: { fontFamily: 'Aptos', fontSizePx: 16 } }] } }).fontFamily, 'Aptos', 'without body text the Normal style decides')
  doc = wc.apply(doc, differentFirstPageOps(doc, true, { newId: ids }))
  assert.equal(hasDifferentFirstPage(doc), true)
  assert.deepEqual(bandTexts(doc, 'headerFirst'), [''])
  assert.deepEqual(bandTexts(doc, 'footerFirst'), [''])
  assert.deepEqual(differentFirstPageOps(doc, true), [], 'already on')
  doc = wc.apply(doc, differentFirstPageOps(doc, false))
  assert.equal(hasDifferentFirstPage(doc), false)
  doc = wc.apply(doc, removeBandOps(doc, 'header'))
  assert.equal(hasBand(doc, 'header'), false)
  assert.deepEqual(removeBandOps(doc, 'header'), [], 'nothing left to remove')
})

// ---------------------------------------------------------------------------------------
// Paragraph indents and spacing (DOC-014)

test('paragraph dialog values: units, mixed selections and only the changed fields', async () => {
  const { unitForLocale, readParagraphValues, paragraphPatch, toUnit, fromUnit, toPoints, fromPoints, parseLength, spaceToggle, lineSpacingPatch } = await paragraphFormat()
  assert.equal(unitForLocale('en-US'), 'in')
  assert.equal(unitForLocale('en-GB'), 'cm')
  assert.equal(unitForLocale('uz-Latn-UZ'), 'cm')
  assert.equal(unitForLocale('ru'), 'cm')
  assert.equal(unitForLocale('en'), 'in')
  assert.equal(toUnit(48, 'in'), 0.5)
  assert.equal(toUnit(fromUnit(1.27, 'cm'), 'cm'), 1.27)
  assert.equal(toPoints(fromPoints(12)), 12)
  assert.equal(parseLength('2 cm', 'in'), 2 * 96 / 2.54)
  assert.equal(parseLength('0,5', 'in'), 48)
  assert.equal(parseLength('12pt', 'cm'), 16)
  assert.equal(parseLength('0.5"', 'cm'), 48)
  assert.equal(parseLength('abc', 'in'), null)

  const a = { indentLeftPx: 48, indentFirstLinePx: -24, spaceBeforePx: 16, spaceAfterPx: 8, lineHeight: 1.15, contextualSpacing: true }
  const b = { indentLeftPx: 48, indentFirstLinePx: 24, spaceBeforePx: 16, spaceAfterPx: 0, lineRule: 'exact', lineHeightPx: 24 }
  const one = readParagraphValues([a])
  assert.deepEqual(one, { left: 48, right: 0, special: 'hanging', by: 24, before: 16, after: 8, lineRule: '1.15', lineValue: 1.15, contextualSpacing: true })
  const mixed = readParagraphValues([a, b])
  assert.equal(mixed.left, 48)
  assert.equal(mixed.before, 16)
  assert.equal(mixed.special, null, 'hanging and first line differ')
  assert.equal(mixed.after, null)
  assert.equal(mixed.lineRule, null)
  assert.equal(mixed.lineValue, null)
  assert.equal(mixed.contextualSpacing, null)

  // Only what changed is applied; blank (null) fields keep each paragraph's own value.
  const edited = { ...mixed, left: 96 }
  assert.deepEqual(paragraphPatch(mixed, edited, a), { indentLeftPx: 96 })
  assert.deepEqual(paragraphPatch(mixed, edited, b), { indentLeftPx: 96 })
  assert.deepEqual(paragraphPatch(one, { ...one, special: 'firstLine' }, a), { indentFirstLinePx: 24 })
  assert.deepEqual(paragraphPatch(one, { ...one, special: 'none', by: 0 }, a), { indentFirstLinePx: 0 })
  assert.deepEqual(paragraphPatch(one, { ...one, lineRule: 'exactly', lineValue: fromPoints(18) }, a), { lineRule: 'exact', lineHeightPx: 24 })
  assert.deepEqual(paragraphPatch(one, { ...one, lineRule: 'double', lineValue: null }, a), { lineHeight: 2, lineRule: undefined, lineHeightPx: undefined })
  assert.deepEqual(paragraphPatch(one, { ...one, contextualSpacing: false }, a), { contextualSpacing: false })
  assert.deepEqual(paragraphPatch(one, { ...one }, a), {}, 'nothing changed, nothing applied')
  assert.deepEqual(lineSpacingPatch('multiple', 3), { lineHeight: 3, lineRule: undefined, lineHeightPx: undefined })
  assert.deepEqual(lineSpacingPatch('atLeast', null), { lineRule: 'atLeast', lineHeightPx: 16 })

  assert.deepEqual(spaceToggle([{ spaceBeforePx: 0 }], 'before'), { label: 'Add space before paragraph', add: true, px: 16 })
  assert.deepEqual(spaceToggle([{ spaceAfterPx: 10.667 }], 'after'), { label: 'Remove space after paragraph', add: false, px: 0 })
})

test('paragraph indents and spacing round-trip through DOCX save and reopen', async () => {
  const { readParagraphValues, paragraphPatch, fromUnit, fromPoints } = await paragraphFormat()
  const wc = await engine()
  const source = document(1)
  const original = readParagraphValues([source.blocks[0].style])
  const targets = [
    { left: fromUnit(1, 'in'), right: fromUnit(0.5, 'in'), special: 'hanging', by: fromUnit(0.25, 'in'), before: fromPoints(12), after: fromPoints(6), lineRule: 'exactly', lineValue: fromPoints(18), contextualSpacing: true },
    { left: fromUnit(2.5, 'cm'), right: 0, special: 'firstLine', by: fromUnit(1.27, 'cm'), before: fromPoints(24), after: 0, lineRule: 'multiple', lineValue: 3, contextualSpacing: false },
    { left: 0, right: fromUnit(1, 'cm'), special: 'none', by: 0, before: 0, after: fromPoints(18), lineRule: 'atLeast', lineValue: fromPoints(20), contextualSpacing: null },
    { left: 0, right: 0, special: 'none', by: 0, before: fromPoints(6), after: fromPoints(6), lineRule: '1.5', lineValue: 1.5, contextualSpacing: null },
  ]
  const ops = targets.map((edited, index) => ({ type: 'setParaStyle', blockId: source.blocks[index].id, patch: paragraphPatch(original, edited, source.blocks[index].style) }))
  const doc = wc.apply(source, ops)
  const reopened = wc.reopen(await wc.docx(doc))
  const close = (actual, expected, label) => assert.ok(Math.abs(actual - expected) < 0.2, `${label}: ${actual} vs ${expected}`)
  targets.forEach((edited, index) => {
    const values = readParagraphValues([reopened.blocks[index].style])
    close(values.left, edited.left, `paragraph ${index + 1} left`)
    close(values.right, edited.right, `paragraph ${index + 1} right`)
    assert.equal(values.special, edited.special, `paragraph ${index + 1} special`)
    close(values.by, edited.by, `paragraph ${index + 1} by`)
    close(values.before, edited.before, `paragraph ${index + 1} before`)
    close(values.after, edited.after, `paragraph ${index + 1} after`)
    assert.equal(values.lineRule, edited.lineRule, `paragraph ${index + 1} line rule`)
    close(values.lineValue, edited.lineValue, `paragraph ${index + 1} line value`)
    if (edited.contextualSpacing !== null) assert.equal(values.contextualSpacing, edited.contextualSpacing, `paragraph ${index + 1} contextual spacing`)
  })
})

// ---------------------------------------------------------------------------------------
// The font list (DOC-018)

function memoryStorage() {
  const data = new Map()
  return { getItem: (key) => (data.has(key) ? data.get(key) : null), setItem: (key, value) => data.set(key, String(value)), data }
}

test('recent and document fonts are remembered locally and requested at the next start', async () => {
  const { readRecentFonts, rememberRecentFont, rememberDocumentFonts, startupFontRequests, normalizeFamily, MAX_RECENT_FONTS } = await fontList()
  const storage = memoryStorage()
  assert.deepEqual(readRecentFonts(storage), [])
  rememberRecentFont(storage, 'Bahnschrift')
  rememberRecentFont(storage, 'Century  Gothic')
  rememberRecentFont(storage, 'bahnschrift')
  assert.deepEqual(readRecentFonts(storage), ['bahnschrift', 'Century Gothic'], 'most recent first, no duplicates, spaces collapsed')
  for (let index = 0; index < 12; index += 1) rememberRecentFont(storage, `Font ${index}`)
  assert.equal(readRecentFonts(storage).length, MAX_RECENT_FONTS)
  rememberDocumentFonts(storage, ['Segoe Print', 'Font 11'])
  assert.deepEqual(startupFontRequests(storage).slice(0, 3), ['Font 11', 'Font 10', 'Font 9'])
  assert.ok(startupFontRequests(storage).includes('Segoe Print'))
  assert.equal(startupFontRequests(storage).filter((family) => family === 'Font 11').length, 1)
  assert.equal(normalizeFamily('"Cambria", serif'), 'Cambria')
  // Broken or blocked storage never breaks fonts.
  assert.deepEqual(readRecentFonts({ getItem: () => '{not json', setItem() {} }), [])
  assert.deepEqual(rememberRecentFont({ getItem: () => null, setItem() { throw new Error('blocked') } }, 'Arial'), ['Arial'])
  assert.deepEqual(readRecentFonts(null), [])
})

test('the font list: type-to-search, recent and document sections, and what each font can do now', async () => {
  const { fontSections, matchRank, documentFontFamilies, allFonts } = await fontList()
  const view = {
    installed: ['Arial', 'Bahnschrift', 'Calibri', 'Century Gothic', 'Segoe UI'],
    ready: new Set(['arial', 'calibri', 'segoe ui']),
    builtin: [{ family: 'Cambria', label: 'Cambria (Caladea)' }, { family: 'Calibri', label: 'Calibri (Carlito)' }],
  }
  const all = allFonts(view)
  assert.deepEqual(all.map((entry) => entry.family), ['Arial', 'Bahnschrift', 'Calibri', 'Cambria', 'Century Gothic', 'Segoe UI'])
  assert.deepEqual(Object.fromEntries(all.map((entry) => [entry.family, entry.state])), { Arial: 'ready', Bahnschrift: 'pending', Calibri: 'ready', Cambria: 'ready', 'Century Gothic': 'pending', 'Segoe UI': 'ready' })
  assert.equal(all.find((entry) => entry.family === 'Cambria').label, 'Cambria (Caladea)', 'a built-in look-alike says what it is')

  const sections = fontSections(view, { recent: ['Bahnschrift', 'Not There'], document: ['Calibri', 'Bahnschrift', 'Wingdings 9'] })
  assert.deepEqual(sections.map((section) => section.title), ['Recently used', 'In this document', 'All fonts'])
  assert.deepEqual(sections[0].fonts.map((entry) => entry.family), ['Bahnschrift'], 'an uninstalled recent font is left out')
  assert.deepEqual(sections[1].fonts.map((entry) => [entry.family, entry.state]), [['Calibri', 'ready'], ['Wingdings 9', 'missing']])

  const found = fontSections(view, { query: 'goth' })
  assert.deepEqual(found.map((section) => section.title), ['Matching fonts'])
  assert.deepEqual(found[0].fonts.map((entry) => entry.family), ['Century Gothic'])
  assert.deepEqual(fontSections(view, { query: 'ca' })[0].fonts.map((entry) => entry.family), ['Calibri', 'Cambria'], 'prefix matches first')
  assert.deepEqual(fontSections(view, { query: 'wing', document: ['Wingdings 9'] })[0].fonts.map((entry) => entry.family), ['Wingdings 9'])
  assert.deepEqual(fontSections(view, { query: 'zzz' })[0].fonts, [])
  assert.equal(matchRank('Century Gothic', 'century gothic'), 0)
  assert.equal(matchRank('Century Gothic', 'cen'), 1)
  assert.equal(matchRank('Century Gothic', 'got'), 2)
  assert.equal(matchRank('Century Gothic', 'cgth'), 4)
  assert.equal(matchRank('Arial', 'xyz'), null)

  const doc = document(1)
  doc.blocks[1] = paragraph('p-b', 'Heading', {}, { fontFamily: 'Bahnschrift' })
  doc.section.footer = [paragraph('f', 'Footer', {}, { fontFamily: '"Segoe UI", sans-serif' })]
  doc.blocks.push({ kind: 'table', id: 't', revision: 0, rows: [{ cells: [{ id: 'c', blocks: [paragraph('c-p', 'Cell', {}, { fontFamily: 'Consolas' })] }] }] })
  assert.deepEqual(documentFontFamilies(doc), ['Calibri', 'Bahnschrift', 'Consolas', 'Segoe UI'], 'body, tables and footers; the stylesheet default is not text')
})

test('an installed font outside the curated list survives DOCX save and reopen', async () => {
  const wc = await engine()
  const doc = document(1)
  doc.blocks[0] = paragraph('p0', 'Bahnschrift text', {}, { fontFamily: 'Bahnschrift' })
  const { normalizeFamily } = await fontList()
  const reopened = wc.reopen(await wc.docx(doc))
  // The importer adds a generic fallback ("Bahnschrift, serif"); the engine and the font box use the first family.
  assert.equal(normalizeFamily(reopened.blocks[0].runs[0].style.fontFamily), 'Bahnschrift')
  const xml = await (await JSZip.loadAsync(await wc.docx(doc))).file('word/document.xml').async('string')
  assert.match(xml, /<w:rFonts w:ascii="Bahnschrift" w:hAnsi="Bahnschrift"/)
})

test('dialog numbers accept a decimal comma and report what is not a number', async () => {
  const { parseNumber, formatNumber } = await import('../src/ui/settings-dialog.ts')
  assert.equal(parseNumber('1,5'), 1.5)
  assert.equal(parseNumber(' 12 '), 12)
  assert.equal(parseNumber('-0.25'), -0.25)
  assert.equal(parseNumber(''), null, 'blank keeps the value as it is')
  assert.ok(Number.isNaN(parseNumber('2 cm')), 'units belong in the unit label, not the number')
  assert.ok(Number.isNaN(parseNumber('abc')))
  assert.equal(formatNumber(0.30000000000000004), '0.3')
  assert.equal(formatNumber(1.276), '1.28')
})
