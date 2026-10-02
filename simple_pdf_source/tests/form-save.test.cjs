'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { PDFDocument, PDFName } = require('pdf-lib')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { formDocument, savedFormValues } = require('./helpers/form-fixtures.cjs')

const REPORT = { report: true }

async function flatten(formValues, { report = true, input } = {}) {
  const { invoke } = loadMain()
  return invoke('pdf:flatten-overlays', input || await formDocument(), [], formValues, report ? REPORT : {})
}

test('values are written to fields under nameless parents and multi-line fields keep line breaks', async () => {
  const result = await flatten({ child: 'Nested value', Address: 'Line 1\nLine 2 (edited)' })
  assert.equal(result.ok, true, JSON.stringify(result.failures))
  assert.deepEqual(result.failures, [])
  const values = await savedFormValues(result.data)
  assert.equal(values['undefined.child'], 'Nested value')
  assert.equal(values.Address, 'Line 1\nLine 2 (edited)')
})

test('a value longer than /MaxLen is truncated only as far as needed, and reported', async () => {
  const result = await flatten({ zip: '100017' })
  assert.deepEqual(result.failures, [])
  const [warning] = result.warnings
  assert.equal(warning.code, 'FORM_VALUE_TRUNCATED')
  assert.equal(warning.field, 'zip')
  assert.equal(warning.savedValue, '10001')
  assert.equal(warning.dataLoss, true)
  assert.equal((await savedFormValues(result.data)).zip, '10001')
})

test('without a problem report, a truncated value fails the save instead of silently losing text', async () => {
  await assert.rejects(flatten({ zip: '100017' }, { report: false }), /FORM_VALUE_TRUNCATED: Nothing was saved: Form field “zip” allows at most 5 characters \(you typed 6\)/)
})

test('fields that cannot be written are reported by name; the other values are still written', async () => {
  const result = await flatten({ missing: 'value', zip: '12345', country: 'Mars' })
  assert.equal(result.ok, false)
  const byCode = Object.fromEntries(result.failures.map((failure) => [failure.code, failure]))
  assert.equal(byCode.FORM_FIELD_NOT_FOUND.field, 'missing')
  assert.equal(byCode.FORM_VALUE_INVALID.field, 'country')
  const values = await savedFormValues(result.data)
  assert.deepEqual(values.country, [])
  assert.equal(values.zip, '12345')
  await assert.rejects(flatten({ missing: 'value' }, { report: false }), /FORM_FIELD_NOT_FOUND: Nothing was saved: Form field “missing” was not found/)
})

test('a value typed into a read-only field is kept and reported, not dropped', async () => {
  const result = await flatten({ locked: 'changed' })
  assert.equal(result.ok, true)
  assert.deepEqual(result.warnings.map((warning) => [warning.code, warning.field]), [['FORM_FIELD_READ_ONLY', 'locked']])
  assert.equal((await savedFormValues(result.data)).locked, 'changed')
  const bytes = await flatten({ locked: 'changed' }, { report: false })
  assert.equal((await savedFormValues(bytes)).locked, 'changed')
})

test('an unchanged value is not a change: read-only and untouched fields stay as they are', async () => {
  const result = await flatten({ locked: 'fixed', zip: '00000' })
  assert.equal(result.ok, true)
  assert.deepEqual(result.warnings, [])
})

test('radio groups, dropdowns with export/display pairs and multi-select lists are saved', async () => {
  const result = await flatten({ color: 'green', country: 'United States', langs: ['uz', 'ru'] })
  assert.equal(result.ok, true, JSON.stringify(result.failures))
  const doc = await PDFDocument.load(result.data)
  const form = doc.getForm()
  const color = form.getRadioGroup('color')
  assert.equal(color.getSelected(), 'green')
  const states = color.acroField.getWidgets().map((widget) => widget.getAppearanceState()?.decodeText())
  assert.deepEqual(states, ['Off', '1', 'Off'])
  const country = form.getDropdown('country')
  assert.deepEqual(country.getSelected(), ['US'], 'the export value is stored')
  const langs = form.getOptionList('langs')
  assert.deepEqual([...langs.getSelected()].sort(), ['ru', 'uz'])
  assert.deepEqual(langs.acroField.dict.lookup(PDFName.of('I')).asArray().map((value) => value.asNumber()), [1, 2])

  // pdf.js reports radio on-state names ("1") rather than export values.
  const byState = await flatten({ color: '2', country: 'CA' })
  const reloaded = (await PDFDocument.load(byState.data)).getForm()
  assert.equal(reloaded.getRadioGroup('color').getSelected(), 'blue')
  assert.deepEqual(reloaded.getDropdown('country').getSelected(), ['CA'])
})

test('a dropdown appearance shows the option text, not the export value', async () => {
  const { decodedStreams } = require('./helpers/pdf-test-utils.cjs')
  const result = await flatten({ country: 'UZ' })
  const doc = await PDFDocument.load(result.data)
  const widget = doc.getForm().getDropdown('country').acroField.getWidgets()[0]
  const appearance = doc.context.lookup(widget.dict.lookup(PDFName.of('AP')).get(PDFName.of('N')))
  const streams = await decodedStreams(result.data)
  const stream = streams.find(({ dict }) => dict === appearance.dict) || streams.find(({ ref }) => ref.tag === widget.dict.lookup(PDFName.of('AP')).get(PDFName.of('N')).tag)
  const hex = Buffer.from('Uzbekistan').toString('hex').toUpperCase()
  assert.ok(stream.data.includes(hex) || stream.data.includes('Uzbekistan'), 'appearance draws “Uzbekistan”')
})

test('a field whose widget is missing from /Fields is still filled, as pdf.js shows it', async () => {
  const { PDFArray } = require('pdf-lib')
  const doc = await PDFDocument.load(await formDocument())
  const fields = doc.catalog.lookup(PDFName.of('AcroForm')).lookup(PDFName.of('Fields'), PDFArray)
  const zipRef = doc.getForm().getTextField('zip').ref
  for (let index = fields.size() - 1; index >= 0; index -= 1) if (fields.get(index) === zipRef) fields.remove(index)
  const input = await doc.save({ updateFieldAppearances: false })
  const result = await flatten({ zip: '77777' }, { input })
  assert.equal(result.ok, true, JSON.stringify(result.failures))
  const reloaded = await PDFDocument.load(result.data)
  const value = reloaded.context.lookup(zipRef).lookup(PDFName.of('V'))
  assert.equal(value.decodeText(), '77777')
})

test('a widget in a PDF without any /AcroForm is still filled', async () => {
  const doc = await PDFDocument.load(await formDocument())
  const zipRef = doc.getForm().getTextField('zip').ref
  doc.catalog.delete(PDFName.of('AcroForm'))
  const input = await doc.save({ updateFieldAppearances: false })
  const result = await flatten({ zip: '22222' }, { input })
  assert.equal(result.ok, true, JSON.stringify(result.failures))
  const reloaded = await PDFDocument.load(result.data)
  assert.equal(reloaded.context.lookup(zipRef).lookup(PDFName.of('V')).decodeText(), '22222')
})

test('fields can be addressed by widget id as well as by name', async () => {
  const input = await formDocument()
  const doc = await PDFDocument.load(input)
  // pdf.js annotation ids name the widget ("12R"), not the field.
  const widgetRef = doc.getForm().getTextField('zip').acroField.dict.lookup(PDFName.of('Kids')).get(0)
  const id = `${widgetRef.objectNumber}R`
  const { invoke } = loadMain()
  const result = await invoke('pdf:flatten-overlays', input, [], [{ ref: id, value: '54321' }], REPORT)
  assert.equal(result.ok, true)
  assert.equal((await savedFormValues(result.data)).zip, '54321')
})
