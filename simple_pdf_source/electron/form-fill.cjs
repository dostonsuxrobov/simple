'use strict'

/**
 * Write typed form values into a PDF without losing any of them silently.
 *
 *  - Fields are found by the renderer's pdf.js name (nameless parents are
 *    skipped, as pdf.js does; pdf-lib would call them "undefined.child"), by
 *    pdf-lib's fully qualified name, or by a widget id such as "12R".
 *  - Nothing happens unless a value really changes: no AcroForm is created,
 *    XFA is kept, and no other field's appearance is regenerated.
 *  - A value longer than /MaxLen is truncated only as far as the limit forces,
 *    and reported. Every field that cannot be written is reported.
 *  - Appearances use a font that can draw the value (base-14 font from the
 *    field's /DA when possible, otherwise an installed Unicode font), with
 *    /NeedAppearances as the fallback when no installed font can.
 */
const {
  PDFAcroCheckBox,
  PDFAcroComboBox,
  PDFAcroListBox,
  PDFAcroPushButton,
  PDFAcroRadioButton,
  PDFAcroSignature,
  PDFAcroText,
  PDFArray,
  PDFBool,
  PDFButton,
  PDFCheckBox,
  PDFDict,
  PDFDocument,
  PDFDropdown,
  PDFForm,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFOptionList,
  PDFRadioGroup,
  PDFRef,
  PDFSignature,
  PDFString,
  PDFTextField,
  StandardFonts,
  TextAlignment,
  createPDFAcroField,
  defaultDropdownAppearanceProvider,
  defaultOptionListAppearanceProvider,
  defaultTextFieldAppearanceProvider,
  layoutMultilineText,
} = require('pdf-lib')
const { problem } = require('./pdf-problems.cjs')
const { FontLibrary } = require('./font-fallback.cjs')

const OFF = PDFName.of('Off')
const MULTILINE_AUTO_MAX_SIZE = 12

function stringOf(value) {
  if (value instanceof PDFString || value instanceof PDFHexString) return value.decodeText()
  if (value instanceof PDFName) return value.decodeText()
  return undefined
}

/**
 * Accepts the renderer's `{ [pdfjsFieldName]: value }` map, values shaped as
 * `{ value, ref }`, or an array of `{ name?, ref?, value }` entries.
 */
function normalizeFormValues(formValues) {
  const entries = []
  const push = (name, ref, value) => entries.push({
    name: typeof name === 'string' ? name : undefined,
    ref: typeof ref === 'string' && ref ? ref : undefined,
    value,
  })
  if (Array.isArray(formValues)) {
    for (const item of formValues) if (item && typeof item === 'object') push(item.name, item.ref ?? item.id, item.value)
  } else if (formValues && typeof formValues === 'object') {
    for (const [name, value] of Object.entries(formValues)) {
      if (value && typeof value === 'object' && !Array.isArray(value) && 'value' in value) push(name, value.ref ?? value.id, value.value)
      else push(name, undefined, value)
    }
  }
  return entries.filter((entry) => entry.name !== undefined || entry.ref !== undefined)
}

/** pdf.js joins /T up the /Parent chain and skips ancestors without /T. */
function pdfjsFieldName(context, dict) {
  const parts = []
  const visited = new Set()
  let current = dict
  while (current instanceof PDFDict && !visited.has(current) && visited.size < 256) {
    visited.add(current)
    const title = stringOf(current.lookup(PDFName.of('T')))
    if (title !== undefined) parts.unshift(title)
    const parent = current.get(PDFName.of('Parent'))
    current = parent === undefined ? undefined : context.lookup(parent)
  }
  return parts.join('.')
}

/** pdf.js annotation ids: "12R" for generation 0, "12R3" otherwise. */
function widgetId(ref) {
  return ref.generationNumber ? `${ref.objectNumber}R${ref.generationNumber}` : `${ref.objectNumber}R`
}

function widgetRefs(field) {
  const kids = field.acroField.dict.lookupMaybe(PDFName.of('Kids'), PDFArray)
  if (!kids) return [field.ref]
  const refs = []
  for (let index = 0; index < kids.size(); index += 1) {
    const value = kids.get(index)
    if (value instanceof PDFRef) refs.push(value)
  }
  return refs
}

function safeName(field) {
  try { return field.getName() } catch { return undefined }
}

function fieldFromDict(pdfDoc, dict, ref) {
  const acroField = createPDFAcroField(dict, ref)
  const types = [
    [PDFTextField, PDFAcroText],
    [PDFCheckBox, PDFAcroCheckBox],
    [PDFRadioGroup, PDFAcroRadioButton],
    [PDFDropdown, PDFAcroComboBox],
    [PDFOptionList, PDFAcroListBox],
    [PDFButton, PDFAcroPushButton],
    [PDFSignature, PDFAcroSignature],
  ]
  const match = types.find(([, AcroClass]) => acroField instanceof AcroClass)
  return match ? match[0].of(acroField, ref, pdfDoc) : undefined
}

function buildFieldIndex(pdfDoc, form) {
  const { context } = pdfDoc
  const byName = new Map()
  const byLibName = new Map()
  const byWidget = new Map()
  const fieldRefs = new Set()
  const add = (field) => {
    const info = {
      field,
      name: pdfjsFieldName(context, field.acroField.dict),
      libName: safeName(field),
      widgetIds: widgetRefs(field).map(widgetId),
    }
    fieldRefs.add(field.ref.tag)
    for (const [map, key] of [[byName, info.name], [byLibName, info.libName]]) {
      if (key === undefined) continue
      if (!map.has(key)) map.set(key, [])
      map.get(key).push(info)
    }
    for (const id of info.widgetIds) byWidget.set(id, info)
    if (!byWidget.has(widgetId(field.ref))) byWidget.set(widgetId(field.ref), info)
  }
  if (form) for (const field of form.getFields()) add(field)
  // Widgets that are missing from /Fields still show (and are filled) in
  // viewers that read page annotations, as pdf.js does.
  for (const page of pdfDoc.getPages()) {
    const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annots) continue
    for (let index = 0; index < annots.size(); index += 1) {
      const ref = annots.get(index)
      if (!(ref instanceof PDFRef) || byWidget.has(widgetId(ref))) continue
      const widget = context.lookup(ref)
      if (!(widget instanceof PDFDict) || stringOf(widget.get(PDFName.of('Subtype'))) !== 'Widget') continue
      const owner = widget.has(PDFName.of('T')) || !widget.has(PDFName.of('Parent')) ? { dict: widget, ref } : (() => {
        const parentRef = widget.get(PDFName.of('Parent'))
        const parent = context.lookup(parentRef)
        return parentRef instanceof PDFRef && parent instanceof PDFDict ? { dict: parent, ref: parentRef } : null
      })()
      if (!owner || fieldRefs.has(owner.ref.tag)) continue
      const field = (() => { try { return fieldFromDict(pdfDoc, owner.dict, owner.ref) } catch { return undefined } })()
      if (field) add(field)
    }
  }
  return { byName, byLibName, byWidget }
}

function resolveFields(index, entry) {
  if (entry.ref && index.byWidget.has(entry.ref)) return [index.byWidget.get(entry.ref)]
  if (entry.name !== undefined) {
    if (index.byName.has(entry.name)) return index.byName.get(entry.name)
    if (index.byLibName.has(entry.name)) return index.byLibName.get(entry.name)
  }
  return []
}

function label(info) {
  return info.name || info.libName || 'unnamed'
}

function inheritedNumber(context, dict, key) {
  const visited = new Set()
  let current = dict
  while (current instanceof PDFDict && !visited.has(current) && visited.size < 256) {
    visited.add(current)
    const value = context.lookup(current.get(PDFName.of(key)))
    if (value instanceof PDFNumber) return value.asNumber()
    const parent = current.get(PDFName.of('Parent'))
    current = parent === undefined ? undefined : context.lookup(parent)
  }
  return undefined
}

/** Longest prefix that fits `limit` UTF-16 units without splitting a character. */
function truncateText(text, limit) {
  let result = ''
  for (const character of text) {
    if (result.length + character.length > limit) break
    result += character
  }
  return result
}

class FieldProblem extends Error {
  constructor(code, message) {
    super(message)
    this.problemCode = code
  }
}

/**
 * The author locked this field, but the value was typed (the editor allowed
 * it). Dropping it would lose input silently, so it is written and reported.
 */
function noteReadOnly(info, report) {
  if (!info.field.isReadOnly()) return
  report.warnings.push(problem('FORM_FIELD_READ_ONLY',
    `Form field “${label(info)}” is read-only in this PDF; the value you typed was saved anyway.`,
    { field: info.name, ref: info.widgetIds[0] }))
}

function readText(field) {
  try { return field.getText() ?? '' } catch { return '' }
}

function writeText(info, raw, report) {
  const { field } = info
  const context = field.acroField.dict.context
  let text = raw === null || raw === undefined || raw === false ? '' : String(raw)
  const maxLength = inheritedNumber(context, field.acroField.dict, 'MaxLen')
  if (Number.isInteger(maxLength) && maxLength >= 0 && text.length > maxLength) {
    const kept = truncateText(text, maxLength)
    report.warnings.push(problem('FORM_VALUE_TRUNCATED',
      `Form field “${label(info)}” allows at most ${maxLength} characters, so only “${kept}” was saved.`,
      {
        field: info.name,
        ref: info.widgetIds[0],
        maxLength,
        value: text,
        savedValue: kept,
        dataLoss: true,
        blockingMessage: `Form field “${label(info)}” allows at most ${maxLength} characters (you typed ${text.length}). Shorten it and save again.`,
      }))
    text = kept
  }
  if (readText(field) === text) return null
  noteReadOnly(info, report)
  field.setText(text || undefined)
  return { kind: 'text', expected: text }
}

function buttonOnValues(field) {
  const values = []
  for (const widget of field.acroField.getWidgets()) {
    const onValue = widget.getOnValue()
    if (onValue && !values.includes(onValue)) values.push(onValue)
  }
  return values
}

function buttonExportValues(field) {
  try { return field.acroField.getExportValues()?.map(stringOf) } catch { return undefined }
}

/** The on-state for a value: an appearance state name or an /Opt export value. */
function buttonStateFor(field, raw) {
  const onValues = buttonOnValues(field)
  if (raw === false || raw === null || raw === undefined || raw === '' || raw === 'Off' || raw === 'false') return OFF
  const wanted = String(raw)
  if (raw !== true) {
    const byName = onValues.find((value) => value.decodeText() === wanted)
    if (byName) return byName
    const exportValues = buttonExportValues(field)
    if (exportValues) {
      const widgets = field.acroField.getWidgets()
      const index = exportValues.findIndex((value) => value === wanted)
      const onValue = index >= 0 ? widgets[index]?.getOnValue() : undefined
      if (onValue) return onValue
    }
  }
  // A plain "checked" from a check box control.
  if (raw === true || ['true', 'on', 'On', 'Yes', 'yes'].includes(wanted)) {
    if (field instanceof PDFCheckBox || onValues.length === 1) return onValues[0] || PDFName.of('Yes')
  }
  return undefined
}

function writeButton(info, raw, report) {
  const { field } = info
  const state = buttonStateFor(field, raw)
  if (!state) {
    throw new FieldProblem('FORM_VALUE_INVALID', `“${String(raw)}” is not one of the choices of form field “${label(info)}”, so it was not saved.`)
  }
  const current = field.acroField.dict.context.lookup(field.acroField.V?.() ?? field.acroField.dict.get(PDFName.of('V')))
  const currentName = current instanceof PDFName ? current : OFF
  if (currentName === state) return null
  noteReadOnly(info, report)
  field.acroField.dict.set(PDFName.of('V'), state)
  for (const widget of field.acroField.getWidgets()) {
    // A check box without appearances gets pdf-lib's /Yes appearance later.
    const onValue = widget.getOnValue() ?? (field instanceof PDFCheckBox ? PDFName.of('Yes') : undefined)
    widget.setAppearanceState(onValue === state ? state : OFF)
  }
  return { kind: 'button', expected: state.decodeText() }
}

function choiceOptions(field) {
  return field.acroField.getOptions().map((option) => ({
    exportValue: stringOf(option.value) ?? '',
    display: stringOf(option.display || option.value) ?? '',
  }))
}

function choiceValues(field) {
  return field.acroField.getValues().map((value) => stringOf(value) ?? '')
}

function writeChoice(info, raw, report) {
  const { field } = info
  const options = choiceOptions(field)
  const wanted = (Array.isArray(raw) ? raw : raw === null || raw === undefined || raw === '' || raw === false ? [] : [raw])
    .map((value) => String(value))
  const multiSelect = field instanceof PDFOptionList || field instanceof PDFDropdown
    ? field.isMultiselect()
    : false
  if (wanted.length > 1 && !multiSelect) {
    throw new FieldProblem('FORM_VALUE_INVALID', `Form field “${label(info)}” accepts one choice, so the ${wanted.length} selected choices were not saved.`)
  }
  const indices = []
  const values = wanted.map((value) => {
    let index = options.findIndex((option) => option.exportValue === value)
    if (index < 0) index = options.findIndex((option) => option.display === value)
    if (index >= 0) {
      indices.push(index)
      return options[index].exportValue
    }
    if (field instanceof PDFDropdown && field.isEditable()) return value
    throw new FieldProblem('FORM_VALUE_INVALID', `“${value}” is not one of the choices of form field “${label(info)}”, so it was not saved.`)
  })
  const current = choiceValues(field)
  const same = current.length === values.length && (field instanceof PDFOptionList
    ? [...current].sort().join('\u0000') === [...values].sort().join('\u0000')
    : current.every((value, index) => value === values[index]))
  if (same) return null
  noteReadOnly(info, report)
  const { dict } = field.acroField
  if (!values.length) dict.delete(PDFName.of('V'))
  else if (values.length === 1) dict.set(PDFName.of('V'), PDFHexString.fromText(values[0]))
  else dict.set(PDFName.of('V'), dict.context.obj(values.map((value) => PDFHexString.fromText(value))))
  // /I (selected option indices, ascending) disambiguates repeated export
  // values; a stale /I would make viewers show the old selection.
  if (indices.length && (field instanceof PDFOptionList || values.length > 1)) {
    dict.set(PDFName.of('I'), dict.context.obj([...new Set(indices)].sort((left, right) => left - right).map((value) => PDFNumber.of(value))))
  } else {
    dict.delete(PDFName.of('I'))
  }
  return { kind: 'choice', expected: values }
}

function writeValue(info, raw, report) {
  const { field } = info
  if (field instanceof PDFTextField) return writeText(info, raw, report)
  if (field instanceof PDFCheckBox || field instanceof PDFRadioGroup) return writeButton(info, raw, report)
  if (field instanceof PDFDropdown || field instanceof PDFOptionList) return writeChoice(info, raw, report)
  throw new FieldProblem('FORM_FIELD_UNSUPPORTED', `Form field “${label(info)}” is a ${field instanceof PDFSignature ? 'signature field' : 'button'} and cannot hold a typed value.`)
}

// ---------------------------------------------------------------- appearances

const FAMILY_FILES = {
  helvetica: { regular: 'arial.ttf', bold: 'arialbd.ttf', italic: 'ariali.ttf', boldItalic: 'arialbi.ttf' },
  times: { regular: 'times.ttf', bold: 'timesbd.ttf', italic: 'timesi.ttf', boldItalic: 'timesbi.ttf' },
  courier: { regular: 'cour.ttf', bold: 'courbd.ttf', italic: 'couri.ttf', boldItalic: 'courbi.ttf' },
  calibri: { regular: 'calibri.ttf', bold: 'calibrib.ttf', italic: 'calibrii.ttf', boldItalic: 'calibriz.ttf' },
  cambria: { regular: 'cambria.ttc', bold: 'cambriab.ttf', italic: 'cambriai.ttf', boldItalic: 'cambriaz.ttf' },
  georgia: { regular: 'georgia.ttf', bold: 'georgiab.ttf', italic: 'georgiai.ttf', boldItalic: 'georgiaz.ttf' },
  verdana: { regular: 'verdana.ttf', bold: 'verdanab.ttf', italic: 'verdanai.ttf', boldItalic: 'verdanaz.ttf' },
  tahoma: { regular: 'tahoma.ttf', bold: 'tahomabd.ttf' },
  segoe: { regular: 'segoeui.ttf', bold: 'segoeuib.ttf', italic: 'segoeuii.ttf', boldItalic: 'segoeuiz.ttf' },
}

const STANDARD_FAMILIES = {
  helvetica: [StandardFonts.Helvetica, StandardFonts.HelveticaBold, StandardFonts.HelveticaOblique, StandardFonts.HelveticaBoldOblique],
  times: [StandardFonts.TimesRoman, StandardFonts.TimesRomanBold, StandardFonts.TimesRomanItalic, StandardFonts.TimesRomanBoldItalic],
  courier: [StandardFonts.Courier, StandardFonts.CourierBold, StandardFonts.CourierOblique, StandardFonts.CourierBoldOblique],
}

function defaultAppearanceText(context, ...dicts) {
  for (const dict of dicts) {
    const value = dict ? stringOf(context.lookup(dict.get(PDFName.of('DA')))) : undefined
    if (value) return value
  }
  return undefined
}

function inheritedDefaultAppearance(context, dict) {
  const visited = new Set()
  let current = dict
  while (current instanceof PDFDict && !visited.has(current) && visited.size < 256) {
    visited.add(current)
    const value = stringOf(context.lookup(current.get(PDFName.of('DA'))))
    if (value) return value
    const parent = current.get(PDFName.of('Parent'))
    current = parent === undefined ? undefined : context.lookup(parent)
  }
  return undefined
}

const TF_PATTERN = /\/([^\s/<>[\]()%]+)\s+([+-]?(?:\d+\.?\d*|\.\d+))\s+Tf/g

function lastFontInAppearance(da) {
  let match
  let last = null
  TF_PATTERN.lastIndex = 0
  while ((match = TF_PATTERN.exec(String(da || '')))) last = { name: match[1], size: Number(match[2]) }
  return last
}

/** The family a field's /DA font stands for, e.g. /Helv → Helvetica. */
function fieldFontHint(pdfDoc, info, widget) {
  const { context } = pdfDoc
  const acroForm = pdfDoc.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  const da = defaultAppearanceText(context, widget?.dict) || inheritedDefaultAppearance(context, info.field.acroField.dict)
    || defaultAppearanceText(context, acroForm)
  const resource = lastFontInAppearance(da)?.name || 'Helv'
  const fonts = acroForm?.lookupMaybe(PDFName.of('DR'), PDFDict)?.lookupMaybe(PDFName.of('Font'), PDFDict)
  const fontDict = fonts?.lookupMaybe(PDFName.of(resource), PDFDict)
  const baseFont = stringOf(fontDict?.get(PDFName.of('BaseFont'))) || resource
  const value = `${resource} ${baseFont}`.toLowerCase()
  const bold = /bold|black|heavy|semibold|demi|,b\b|-bd\b/.test(value)
  const italic = /italic|oblique|,i\b|-it\b/.test(value)
  const family = /^(helv|arial)|helvetica|arial/.test(value) ? 'helvetica'
    : /tiro|times/.test(value) ? 'times'
      : /cour/.test(value) ? 'courier'
        : Object.keys(FAMILY_FILES).find((key) => value.includes(key)) || 'helvetica'
  return { family, bold, italic }
}

function variantIndex(style) {
  return style.bold && style.italic ? 3 : style.bold ? 1 : style.italic ? 2 : 0
}

async function fieldFont(pdfDoc, fonts, standardFonts, hint, text) {
  const preferred = []
  const standard = STANDARD_FAMILIES[hint.family]?.[variantIndex(hint)]
  if (standard) {
    if (!standardFonts.has(standard)) standardFonts.set(standard, pdfDoc.embedStandardFont(standard))
    preferred.push({ pdfFont: standardFonts.get(standard) })
  }
  const files = FAMILY_FILES[hint.family]
  if (files) preferred.push({ key: hint.family, files })
  return fonts.singleFontFor(text, { preferred, style: hint })
}

function displayedText(field) {
  if (field instanceof PDFTextField) return readText(field)
  if (field instanceof PDFOptionList) return choiceOptions(field).map((option) => option.display).join('\n')
  if (field instanceof PDFDropdown) {
    const options = choiceOptions(field)
    return choiceValues(field).map((value) => options.find((option) => option.exportValue === value)?.display ?? value).join('\n')
  }
  return ''
}

function setDefaultAppearanceSize(widget, da, size) {
  const base = da || '0 g'
  const font = lastFontInAppearance(base)
  const next = font
    ? base.replace(new RegExp(`\\/${font.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s+[+-]?(?:\\d+\\.?\\d*|\\.\\d+)\\s+Tf(?![\\s\\S]*Tf)`), `/${font.name} ${size} Tf`)
    : `${base} /Helv ${size} Tf`
  widget.setDefaultAppearance(next)
}

/**
 * pdf-lib sizes auto-size (0 Tf) multi-line text as large as the box allows,
 * so 'line one' becomes huge. Viewers use at most 12 pt for these fields.
 */
function multilineAutoSize(textField, widget, font) {
  const rectangle = widget.getRectangle()
  const rotation = ((widget.getAppearanceCharacteristics()?.getRotation() ?? 0) % 360 + 360) % 360
  const width = rotation === 90 || rotation === 270 ? rectangle.height : rectangle.width
  const height = rotation === 90 || rotation === 270 ? rectangle.width : rectangle.height
  const borderWidth = widget.getBorderStyle()?.getWidth() ?? 0
  const bounds = { x: borderWidth + 1, y: borderWidth + 1, width: width - (borderWidth + 1) * 2, height: height - (borderWidth + 1) * 2 }
  const text = readText(textField)
  if (!text.trim() || bounds.width <= 0 || bounds.height <= 0) return MULTILINE_AUTO_MAX_SIZE
  try {
    const { fontSize } = layoutMultilineText(text, { alignment: TextAlignment.Left, fontSize: 0, font, bounds })
    return Math.max(4, Math.min(MULTILINE_AUTO_MAX_SIZE, fontSize))
  } catch {
    return MULTILINE_AUTO_MAX_SIZE
  }
}

/**
 * Appearance provider that leaves the field's /DA untouched. pdf-lib rewrites
 * /DA to name the appearance font and a fixed size; that font exists only in
 * the appearance stream's resources, and the fixed size would replace "auto".
 */
function preservingDefaultAppearance(field, provider) {
  return (target, widget, font) => {
    const key = PDFName.of('DA')
    const saved = [[widget.dict, widget.dict.get(key)], [field.acroField.dict, field.acroField.dict.get(key)]]
    try {
      return provider(target, widget, font)
    } finally {
      for (const [dict, value] of saved.reverse()) {
        if (value === undefined) dict.delete(key)
        else dict.set(key, value)
      }
    }
  }
}

function appearanceProvider(pdfDoc, field) {
  const { context } = pdfDoc
  if (field instanceof PDFTextField) {
    return preservingDefaultAppearance(field, (textField, widget, font) => {
      if (textField.isMultiline()) {
        const da = defaultAppearanceText(context, widget.dict) || inheritedDefaultAppearance(context, textField.acroField.dict)
        const size = lastFontInAppearance(da)?.size
        if (!(size > 0)) setDefaultAppearanceSize(widget, da, multilineAutoSize(textField, widget, font))
      }
      return defaultTextFieldAppearanceProvider(textField, widget, font)
    })
  }
  // Choice appearances show the option's display text, not its export value.
  const options = choiceOptions(field)
  const view = Object.create(field)
  view.getSelected = () => choiceValues(field).map((value) => options.find((option) => option.exportValue === value)?.display ?? value)
  const provider = field instanceof PDFDropdown ? defaultDropdownAppearanceProvider : defaultOptionListAppearanceProvider
  return preservingDefaultAppearance(field, (_target, widget, font) => provider(view, widget, font))
}

function setNeedAppearances(pdfDoc) {
  const acroForm = pdfDoc.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  acroForm?.set(PDFName.of('NeedAppearances'), PDFBool.True)
}

async function refreshAppearance(pdfDoc, info, fonts, standardFonts, report) {
  const { field } = info
  if (field instanceof PDFCheckBox || field instanceof PDFRadioGroup) {
    // Keep the document's own check marks; only widgets without an
    // appearance for their state get one generated.
    const missing = field.acroField.getWidgets().some((widget) => {
      const normal = widget.getAppearances()?.normal
      const state = widget.getAppearanceState?.() ?? widget.dict.get(PDFName.of('AS'))
      return !(normal instanceof PDFDict) || (state instanceof PDFName && state !== OFF && !normal.has(state))
    })
    if (missing) field.defaultUpdateAppearances()
    return
  }
  const text = displayedText(field)
  const widget = field.acroField.getWidgets()[0]
  const choice = await fieldFont(pdfDoc, fonts, standardFonts, fieldFontHint(pdfDoc, info, widget), text)
  if (!choice?.complete) {
    setNeedAppearances(pdfDoc)
    report.warnings.push(problem('FORM_APPEARANCE_FALLBACK',
      `No installed font can draw every character of form field “${label(info)}”. The value is saved; PDF viewers will draw it with their own fonts.`,
      { field: info.name, ref: info.widgetIds[0] }))
    if (!choice) return
  }
  field.updateAppearances(choice.font, appearanceProvider(pdfDoc, field))
}

/**
 * Apply `formValues` to `pdfDoc`. Returns what changed plus problems; never
 * throws for a single field.
 */
async function applyFormValues(pdfDoc, formValues, { fonts = new FontLibrary(pdfDoc) } = {}) {
  const report = { changed: 0, written: [], warnings: [], failures: [] }
  const entries = normalizeFormValues(formValues)
  if (!entries.length) return report
  const acroFormDict = pdfDoc.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  // pdf-lib's setters reach doc.getForm(), which strips XFA on the spot, so
  // note it before any value is written.
  const hadXfa = Boolean(acroFormDict?.has(PDFName.of('XFA')))
  const acroForm = pdfDoc.catalog.getAcroForm()
  // PDFForm.of() neither creates an AcroForm nor deletes XFA (getForm() does).
  // Without an /AcroForm, widgets on the pages are still found (pdf.js shows them).
  const index = buildFieldIndex(pdfDoc, acroForm ? PDFForm.of(acroForm, pdfDoc) : null)
  const changed = new Map()
  for (const entry of entries) {
    const targets = resolveFields(index, entry)
    if (!targets.length) {
      report.failures.push(problem('FORM_FIELD_NOT_FOUND',
        `Form field “${entry.name ?? entry.ref}” was not found in this PDF, so its value was not saved.`,
        { field: entry.name, ref: entry.ref }))
      continue
    }
    for (const info of targets) {
      try {
        const written = writeValue(info, entry.value, report)
        if (!written) continue
        changed.set(info.field.ref.tag, info)
        report.written.push({ ref: info.field.ref, name: info.name, ...written })
      } catch (error) {
        report.failures.push(problem(error instanceof FieldProblem ? error.problemCode : 'FORM_FIELD_FAILED',
          error instanceof FieldProblem ? error.message : `Form field “${label(info)}” could not be saved: ${error?.message || error}`,
          { field: info.name, ref: info.widgetIds[0] }))
      }
    }
  }
  report.changed = changed.size
  if (!changed.size) return report
  if (hadXfa) {
    // Acrobat shows XFA data in preference to the AcroForm values just written.
    acroFormDict.delete(PDFName.of('XFA'))
    report.warnings.push(problem('XFA_REMOVED', 'This form also had an XFA layer, which was removed so the values you typed are the ones shown.'))
  }
  const standardFonts = new Map()
  for (const info of changed.values()) {
    try {
      await refreshAppearance(pdfDoc, info, fonts, standardFonts, report)
    } catch (error) {
      setNeedAppearances(pdfDoc)
      // The old appearance shows the old value; viewers that do not redraw
      // fields would keep showing it, so drop it rather than mislead.
      if (!(info.field instanceof PDFCheckBox || info.field instanceof PDFRadioGroup)) {
        for (const widget of info.field.acroField.getWidgets()) widget.dict.delete(PDFName.of('AP'))
      }
      report.warnings.push(problem('FORM_APPEARANCE_FALLBACK',
        `The appearance of form field “${label(info)}” could not be drawn (${error?.message || error}). The value is saved; PDF viewers will draw it.`,
        { field: info.name, ref: info.widgetIds[0] }))
    }
  }
  return report
}

function savedValue(field) {
  if (field instanceof PDFTextField) return readText(field)
  if (field instanceof PDFCheckBox || field instanceof PDFRadioGroup) {
    const value = field.acroField.dict.context.lookup(field.acroField.V?.() ?? field.acroField.dict.get(PDFName.of('V')))
    return value instanceof PDFName ? value.decodeText() : 'Off'
  }
  if (field instanceof PDFDropdown || field instanceof PDFOptionList) return choiceValues(field)
  return undefined
}

/** Reload the saved bytes and confirm every written value is really there. */
async function verifyFormValues(bytes, written) {
  if (!written?.length) return []
  const failures = []
  const doc = await PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false })
  const acroForm = doc.catalog.getAcroForm()
  const fields = new Map()
  if (acroForm) {
    for (const field of PDFForm.of(acroForm, doc).getFields()) fields.set(field.ref.tag, field)
  }
  for (const entry of written) {
    let field = fields.get(entry.ref.tag)
    if (!field) {
      const dict = doc.context.lookup(entry.ref)
      if (dict instanceof PDFDict) { try { field = fieldFromDict(doc, dict, entry.ref) } catch { field = undefined } }
    }
    const actual = field ? savedValue(field) : undefined
    const expected = entry.expected
    const same = Array.isArray(expected)
      ? Array.isArray(actual) && actual.length === expected.length && [...actual].sort().join('\u0000') === [...expected].sort().join('\u0000')
      : actual === expected
    if (!same) {
      failures.push(problem('FORM_VALUE_NOT_SAVED', `Form field “${entry.name || entry.ref.tag}” did not keep its new value in the saved file.`, { field: entry.name }))
    }
  }
  return failures
}

module.exports = {
  applyFormValues,
  normalizeFormValues,
  pdfjsFieldName,
  verifyFormValues,
  widgetId,
}
