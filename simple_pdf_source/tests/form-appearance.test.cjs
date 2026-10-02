'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { PDFDict, PDFDocument, PDFName, StandardFonts } = require('pdf-lib')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { formDocument, savedFormValues } = require('./helpers/form-fixtures.cjs')
const { decodedStreams } = require('./helpers/pdf-test-utils.cjs')

const HIGHLIGHT = {
  id: 'highlight', type: 'highlight', pageIndex: 0, rect: { x: 20, y: 20, width: 100, height: 12 }, color: [1, 0.9, 0.2], opacity: 0.3,
}

async function acroFormOf(bytes) {
  const doc = await PDFDocument.load(bytes)
  return { doc, acroForm: doc.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict) }
}

async function appearanceOf(bytes, name) {
  const doc = await PDFDocument.load(bytes)
  const field = doc.getForm().getFields().find((candidate) => candidate.getName() === name)
  const widget = field.acroField.getWidgets()[0]
  const normalRef = widget.dict.lookup(PDFName.of('AP'), PDFDict).get(PDFName.of('N'))
  const stream = (await decodedStreams(bytes)).find(({ ref }) => ref.tag === normalRef.tag)
  const fonts = doc.context.lookup(normalRef).dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict)
  const baseFonts = fonts.entries().map(([, ref]) => doc.context.lookup(ref).get(PDFName.of('BaseFont'))?.decodeText())
  return { text: stream.data.toString('latin1'), baseFonts, field, doc }
}

test('a highlight-only save never touches the form: a Cyrillic value without an appearance does not break it', async () => {
  const { invoke } = loadMain()
  const input = await formDocument({ unicodeWithoutAppearance: true })
  const output = await invoke('pdf:flatten-overlays', input, [HIGHLIGHT], {}, {})
  const values = await savedFormValues(output)
  assert.equal(values.fio, 'Иванов Иван')
  const { doc, acroForm } = await acroFormOf(output)
  assert.equal(acroForm.get(PDFName.of('NeedAppearances'))?.asBoolean?.(), true)
  const widget = doc.getForm().getTextField('fio').acroField.getWidgets()[0]
  assert.equal(widget.dict.has(PDFName.of('AP')), false, 'no appearance was generated for an untouched field')
})

test('a highlight-only save keeps the XFA layer of a hybrid form', async () => {
  const { invoke } = loadMain()
  const output = await invoke('pdf:flatten-overlays', await formDocument({ xfa: true }), [HIGHLIGHT], {}, {})
  const { acroForm } = await acroFormOf(output)
  assert.equal(acroForm.has(PDFName.of('XFA')), true)
})

test('a plain PDF does not gain an /AcroForm from a highlight save', async () => {
  const { invoke } = loadMain()
  const doc = await PDFDocument.create()
  doc.addPage([300, 300]).drawText('Plain', { x: 20, y: 250, size: 12, font: await doc.embedFont(StandardFonts.Helvetica) })
  const output = await invoke('pdf:flatten-overlays', await doc.save(), [HIGHLIGHT], {}, {})
  assert.equal((await acroFormOf(output)).acroForm, undefined)
})

test('editing a field of a hybrid form removes XFA (so the values show) and says so', async () => {
  const { invoke } = loadMain()
  const result = await invoke('pdf:flatten-overlays', await formDocument({ xfa: true }), [], { zip: '12345' }, { report: true })
  assert.ok(result.warnings.some((warning) => warning.code === 'XFA_REMOVED'))
  const { acroForm } = await acroFormOf(result.data)
  assert.equal(acroForm.has(PDFName.of('XFA')), false)
  assert.equal((await savedFormValues(result.data)).zip, '12345')
})

test('an auto-size multi-line field gets text of at most 12 pt and keeps its auto size', async () => {
  const { invoke } = loadMain()
  const output = await invoke('pdf:flatten-overlays', await formDocument(), [], { notes: 'line one' }, {})
  const { text, field } = await appearanceOf(output, 'notes')
  const sizes = [...text.matchAll(/([\d.]+)\s+Tf/g)].map((match) => Number(match[1]))
  assert.ok(sizes.length && sizes.every((size) => size > 0 && size <= 12), `appearance font sizes ${sizes}`)
  assert.equal(field.acroField.getDefaultAppearance(), '/Helv 0 Tf 0 g', 'the field keeps /DA (auto size, /Helv)')
})

test('Cyrillic, Uzbek and CJK values are saved with a font that can draw them', async () => {
  const { invoke } = loadMain()
  const latin = await invoke('pdf:flatten-overlays', await formDocument(), [], { zip: '12345' }, { report: true })
  assert.deepEqual((await appearanceOf(latin.data, 'zip')).baseFonts, ['Helvetica'], 'Latin text keeps the form’s Helvetica')

  const cyrillic = await invoke('pdf:flatten-overlays', await formDocument(), [], { Address: 'Иванов Иван\nТошкент, Ўзбекистон' }, { report: true })
  assert.equal(cyrillic.ok, true)
  assert.deepEqual(cyrillic.warnings, [])
  assert.equal((await savedFormValues(cyrillic.data)).Address, 'Иванов Иван\nТошкент, Ўзбекистон')
  const cyrillicFonts = (await appearanceOf(cyrillic.data, 'Address')).baseFonts
  assert.ok(cyrillicFonts.every((name) => !/^Helvetica/.test(name)), `a Unicode font is embedded: ${cyrillicFonts}`)

  const cjk = await invoke('pdf:flatten-overlays', await formDocument(), [], { Address: '你好 世界' }, { report: true })
  assert.equal(cjk.ok, true)
  assert.equal((await savedFormValues(cjk.data)).Address, '你好 世界')
  const cjkFonts = (await appearanceOf(cjk.data, 'Address')).baseFonts
  assert.ok(cjkFonts.some((name) => /YaHei|Gothic|SimSun|JhengHei|Malgun/i.test(name)), `a CJK font is embedded: ${cjkFonts}`)
})

test('a character no installed font has keeps the value and asks viewers to draw it', async () => {
  const { invoke } = loadMain()
  const result = await invoke('pdf:flatten-overlays', await formDocument(), [], { Address: 'Code \u{10FFFD}' }, { report: true })
  assert.equal(result.ok, true)
  assert.ok(result.warnings.some((warning) => warning.code === 'FORM_APPEARANCE_FALLBACK' && warning.field === 'Address'))
  assert.equal((await savedFormValues(result.data)).Address, 'Code \u{10FFFD}')
  const { acroForm } = await acroFormOf(result.data)
  assert.equal(acroForm.get(PDFName.of('NeedAppearances'))?.asBoolean?.(), true)
})
