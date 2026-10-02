'use strict'

const { PDFArray, PDFDocument, PDFHexString, PDFName, PDFString } = require('pdf-lib')

/**
 * A form with the field kinds the editor fills: MaxLen text, a field under a
 * parent without /T (pdf.js calls it "child", pdf-lib "undefined.child"),
 * multi-line, read-only, radio group, dropdown with export/display pairs,
 * multi-select list box, an auto-size multi-line field, and optionally a
 * Cyrillic value without an appearance (/NeedAppearances) and an XFA layer.
 */
async function formDocument({ unicodeWithoutAppearance = false, xfa = false } = {}) {
  const doc = await PDFDocument.create()
  const page = doc.addPage([500, 760])
  const form = doc.getForm()
  const box = (y, height = 20, width = 180) => ({ x: 20, y, width, height })

  const zip = form.createTextField('zip')
  zip.setMaxLength(5)
  zip.setText('00000')
  zip.addToPage(page, box(700))

  const address = form.createTextField('Address')
  address.enableMultiline()
  address.setText('Line 1\nLine 2')
  address.addToPage(page, box(620, 60))

  const locked = form.createTextField('locked')
  locked.setText('fixed')
  locked.enableReadOnly()
  locked.addToPage(page, box(590))

  const child = form.createTextField('child')
  child.addToPage(page, box(560))
  // Move "child" under a parent field that has no name.
  const parentRef = doc.context.register(doc.context.obj({ Kids: [child.ref] }))
  child.acroField.dict.set(PDFName.of('Parent'), parentRef)
  const fields = form.acroForm.dict.lookup(PDFName.of('Fields'), PDFArray)
  for (let index = 0; index < fields.size(); index += 1) {
    if (fields.get(index) === child.ref) fields.set(index, parentRef)
  }

  const color = form.createRadioGroup('color')
  ;['red', 'green', 'blue'].forEach((option, index) => color.addOptionToPage(option, page, { x: 20 + index * 40, y: 520, width: 15, height: 15 }))

  const country = form.createDropdown('country')
  country.addOptions(['CA', 'US', 'UZ'])
  country.addToPage(page, box(480))
  country.acroField.dict.set(PDFName.of('Opt'), doc.context.obj([
    [PDFString.of('CA'), PDFString.of('Canada')],
    [PDFString.of('US'), PDFString.of('United States')],
    [PDFString.of('UZ'), PDFString.of('Uzbekistan')],
  ]))

  const langs = form.createOptionList('langs')
  langs.addOptions(['en', 'ru', 'uz'])
  langs.enableMultiselect()
  langs.addToPage(page, box(400, 60))

  const notes = form.createTextField('notes')
  notes.enableMultiline()
  notes.addToPage(page, box(250, 120, 300))
  notes.acroField.dict.set(PDFName.of('DA'), PDFString.of('/Helv 0 Tf 0 g'))

  if (unicodeWithoutAppearance) {
    const name = form.createTextField('fio')
    name.addToPage(page, box(200))
    name.acroField.setValue(PDFHexString.fromText('Иванов Иван'))
    name.acroField.getWidgets()[0].dict.delete(PDFName.of('AP'))
    form.acroForm.dict.set(PDFName.of('NeedAppearances'), doc.context.obj(true))
  }
  if (xfa) {
    const template = doc.context.register(doc.context.flateStream('<template xmlns="http://www.xfa.org/schema/xfa-template/3.3/"/>'))
    form.acroForm.dict.set(PDFName.of('XFA'), doc.context.obj([PDFString.of('template'), template]))
  }
  return doc.save({ updateFieldAppearances: false })
}

/** Field values read back with pdf-lib, keyed by pdf-lib's names. */
async function savedFormValues(bytes) {
  const doc = await PDFDocument.load(bytes)
  const values = {}
  for (const field of doc.getForm().getFields()) {
    const name = field.getName()
    if ('getText' in field) values[name] = field.getText() ?? ''
    else if ('getSelected' in field) values[name] = field.getSelected()
    else if ('isChecked' in field) values[name] = field.isChecked()
  }
  return values
}

module.exports = { formDocument, savedFormValues }
