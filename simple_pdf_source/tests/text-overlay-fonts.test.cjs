'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { PDFDocument } = require('pdf-lib')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { textItems } = require('./helpers/pdf-test-utils.cjs')

async function blankPage() {
  const doc = await PDFDocument.create()
  doc.addPage([500, 600])
  return doc.save()
}

function addedText(id, text, y, extra = {}) {
  return {
    id, type: 'text', pageIndex: 0, rect: { x: 40, y, width: 400, height: 30 }, text,
    fontSize: 16, fontFamily: 'Segoe UI', color: [0, 0, 0], align: 'left', cover: false, textFit: 'wrap', ...extra,
  }
}

const normalize = (value) => value.replace(/\s+/g, '')

test('CJK, kana, Hangul, emoji and Uzbek text is saved as real glyphs, not invisible .notdef', async () => {
  const { invoke } = loadMain()
  const samples = ['你好 世界 Hello', '日本語テキスト', 'Done ✅', '한국어 텍스트', 'Oʻzbekiston — Ўзбекистон']
  const overlays = samples.map((text, index) => addedText(`t${index}`, text, 520 - index * 60))
  const result = await invoke('pdf:flatten-overlays', await blankPage(), overlays, {}, { report: true })
  assert.equal(result.ok, true, JSON.stringify(result.failures))
  const [items] = await textItems(result.data)
  const text = items.map((item) => item.str).join('')
  assert.ok(!text.includes('\u0000'), 'no glyph 0 in the text layer')
  for (const sample of samples) assert.ok(normalize(text).includes(normalize(sample)), `“${sample}” is in ${JSON.stringify(text)}`)
  const fonts = new Set(items.map((item) => item.fontName))
  assert.ok(fonts.size >= 3, 'fallback fonts were used for the scripts the main font lacks')
})

test('one over-full text box is fitted and reported; it does not stop the other edits from saving', async () => {
  const { invoke } = loadMain()
  const longText = 'This sentence is far too long for the small box it was typed into, so it must be fitted'
  const overlays = [
    addedText('first', 'First box', 520),
    addedText('crowded', longText, 400, { rect: { x: 40, y: 400, width: 120, height: 40 } }),
    addedText('third', 'Third box', 300),
    addedText('fourth', 'Fourth box', 240),
  ]
  const result = await invoke('pdf:flatten-overlays', await blankPage(), overlays, {}, { report: true })
  assert.equal(result.ok, true)
  assert.equal(result.warnings[0].overlayId, 'crowded')
  assert.match(result.warnings[0].code, /^TEXT_(SHRUNK|OVERFLOW)$/)
  assert.ok(result.warnings[0].fontSize < 16)
  const [items] = await textItems(result.data)
  const text = normalize(items.map((item) => item.str).join(''))
  for (const expected of ['First box', 'Third box', 'Fourth box', longText]) assert.ok(text.includes(normalize(expected)), expected)

  // The same save without a report also succeeds: nothing was dropped.
  const bytes = await invoke('pdf:flatten-overlays', await blankPage(), overlays, {}, {})
  assert.ok(bytes instanceof Uint8Array)
})

test('a character no installed font can draw is reported instead of being saved as tofu', async () => {
  const { invoke } = loadMain()
  const overlays = [addedText('rare', 'Code \u{10FFFD} end', 500), addedText('fine', 'Fine text', 400)]
  const result = await invoke('pdf:flatten-overlays', await blankPage(), overlays, {}, { report: true })
  assert.equal(result.ok, false)
  assert.equal(result.failures[0].code, 'MISSING_GLYPHS')
  assert.equal(result.failures[0].overlayId, 'rare')
  assert.deepEqual(result.failures[0].chars, ['\u{10FFFD}'])
  const [items] = await textItems(result.data)
  const text = items.map((item) => item.str).join('')
  assert.ok(text.includes('Fine text'))
  assert.ok(text.includes('�'), 'the missing character is visibly marked')
  await assert.rejects(
    invoke('pdf:flatten-overlays', await blankPage(), overlays, {}, {}),
    /MISSING_GLYPHS: Nothing was saved: Text on page 1 contains a character/,
  )
})

test('an edit whose original text cannot be found is reported with its box; other edits still save', async () => {
  const { invoke } = loadMain()
  const overlays = [
    {
      id: 'ghost', type: 'text', pageIndex: 0, rect: { x: 40, y: 500, width: 200, height: 20 },
      originalRect: { x: 40, y: 500, width: 200, height: 20 }, originalText: 'Not on this page', text: 'Replacement',
      fontSize: 12, fontFamily: 'Segoe UI', color: [0, 0, 0], align: 'left', cover: true, textFit: 'fit',
    },
    addedText('kept', 'Kept text', 300),
  ]
  const result = await invoke('pdf:flatten-overlays', await blankPage(), overlays, {}, { report: true })
  assert.equal(result.failures[0].code, 'TEXT_SOURCE_NOT_FOUND')
  assert.equal(result.failures[0].overlayId, 'ghost')
  const [items] = await textItems(result.data)
  assert.deepEqual(items.map((item) => item.str).filter(Boolean), ['Kept text'])
  await assert.rejects(invoke('pdf:flatten-overlays', await blankPage(), overlays, {}, {}), /TEXT_SOURCE_NOT_FOUND: Nothing was saved: The original text “Not on this page” on page 1 could not be located/)
})
