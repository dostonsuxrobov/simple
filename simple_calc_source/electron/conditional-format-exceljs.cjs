'use strict'

// Conditional-formatting fidelity patches for ExcelJS 4.4.
//
// ExcelJS reads and writes <conditionalFormatting> lossily. Unpatched it:
//   - never writes duplicateValues, uniqueValues, notContainsText, beginsWith or endsWith rules
//     (they vanish on save) and writes containsBlanks/containsErrors with an invalid operator;
//   - drops stopIfTrue, text, equalAverage and stdDev on read and write;
//   - parses formula-valued thresholds (<cfvo type="formula" val="$A$1"/>, <xm:f>) as NaN and
//     drops icon-set gte="0" (">" instead of ">=");
//   - drops the main <dataBar showValue minLength maxLength> attributes, reads the x14 data-bar
//     defaults incorrectly, writes x14 children in schema-invalid order, and only writes the
//     x14 extension for solid bars (so gradient bars lose negative/axis/border settings);
//   - leaves Excel 2010 x14 rules other than dataBar/iconSet (cross-sheet expression/cellIs
//     rules) without formulae, which crashes the next save;
//   - turns a dxf <numFmt> into an object that is written back as formatCode="[object Object]".
//
// The patches below are prototype-level and idempotent. They keep ExcelJS's model shape so the
// renderer's conditional-format engine (src/lib/conditional-format.ts) and authored rules round
// trip exactly. Anything not recognised is left to ExcelJS's original behaviour.

const PATCHED = Symbol.for('simple_calc.conditionalFormattingPatched')

const TEXT_OPERATOR_TYPES = {
  containsText: 'containsText',
  notContains: 'notContainsText',
  notContainsText: 'notContainsText',
  beginsWith: 'beginsWith',
  endsWith: 'endsWith',
  containsBlanks: 'containsBlanks',
  notContainsBlanks: 'notContainsBlanks',
  containsErrors: 'containsErrors',
  notContainsErrors: 'notContainsErrors',
}

const TEXT_RULE_OPERATOR = {
  containsText: 'containsText',
  notContainsText: 'notContains',
  beginsWith: 'beginsWith',
  endsWith: 'endsWith',
}

const MAIN_RENDERABLE = new Set([
  'expression', 'cellIs', 'top10', 'aboveAverage', 'colorScale', 'dataBar', 'iconSet', 'timePeriod',
  'duplicateValues', 'uniqueValues', 'containsText', 'notContainsText', 'beginsWith', 'endsWith',
  'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors',
])

const EXT_ICON_SETS = new Set(['3Triangles', '3Stars', '5Boxes'])

function toBool(value) {
  if (value === undefined || value === null) return undefined
  const text = String(value).toLowerCase()
  return text === '1' || text === 'true'
}

function toInt(value) {
  if (value === undefined || value === null || value === '') return undefined
  const number = parseInt(value, 10)
  return Number.isFinite(number) ? number : undefined
}

function boolAttribute(value) {
  return value === undefined || value === null ? undefined : value ? '1' : '0'
}

/** A cfvo value: finite numbers stay numbers, anything else is a formula string. */
function cfvoValue(text) {
  if (text === undefined || text === null) return undefined
  const trimmed = String(text).trim()
  if (!trimmed) return undefined
  if (/^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/.test(trimmed)) return Number(trimmed)
  return trimmed.replace(/^=/, '')
}

/** Canonical Excel rule type (ExcelJS models blank/error rules as type containsText + operator). */
function canonicalType(rule) {
  if (!rule || typeof rule !== 'object') return undefined
  if (rule.type === 'containsText') return TEXT_OPERATOR_TYPES[rule.operator || 'containsText'] || 'containsText'
  return rule.type
}

function topLeftOf(ref) {
  const first = String(ref || 'A1').trim().split(/[\s,]+/)[0] || 'A1'
  const cell = first.split(':')[0].replace(/\$/g, '')
  const bang = cell.lastIndexOf('!')
  return (bang >= 0 ? cell.slice(bang + 1) : cell) || 'A1'
}

function quoteText(text) {
  return `"${String(text ?? '').replace(/"/g, '""')}"`
}

/** Formula Excel writes for text, blank, error and date rules (used when a model has none). */
function derivedFormula(type, rule) {
  const tl = topLeftOf(rule.ref)
  const text = quoteText(rule.text)
  switch (type) {
    case 'containsText': return `NOT(ISERROR(SEARCH(${text},${tl})))`
    case 'notContainsText': return `ISERROR(SEARCH(${text},${tl}))`
    case 'beginsWith': return `LEFT(${tl},LEN(${text}))=${text}`
    case 'endsWith': return `RIGHT(${tl},LEN(${text}))=${text}`
    case 'containsBlanks': return `LEN(TRIM(${tl}))=0`
    case 'notContainsBlanks': return `LEN(TRIM(${tl}))>0`
    case 'containsErrors': return `ISERROR(${tl})`
    case 'notContainsErrors': return `NOT(ISERROR(${tl}))`
    case 'timePeriod':
      switch (rule.timePeriod) {
        case 'today': return `FLOOR(${tl},1)=TODAY()`
        case 'yesterday': return `FLOOR(${tl},1)=TODAY()-1`
        case 'tomorrow': return `FLOOR(${tl},1)=TODAY()+1`
        case 'last7Days': return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`
        case 'thisWeek': return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`
        case 'lastWeek': return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`
        case 'nextWeek': return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`
        case 'thisMonth': return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`
        case 'lastMonth': return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`
        case 'nextMonth': return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`
        default: return undefined
      }
    default: return undefined
  }
}

function ruleFormulae(rule) {
  return Array.isArray(rule.formulae) ? rule.formulae.filter((formula) => typeof formula === 'string' || typeof formula === 'number').map(String) : []
}

/** Extra cfRule attributes ExcelJS does not model. */
function extraRuleAttributes(attributes) {
  const extra = {}
  const stopIfTrue = toBool(attributes.stopIfTrue)
  if (stopIfTrue) extra.stopIfTrue = true
  if (attributes.text !== undefined) extra.text = attributes.text
  const equalAverage = toBool(attributes.equalAverage)
  if (equalAverage) extra.equalAverage = true
  const stdDev = toInt(attributes.stdDev)
  if (stdDev) extra.stdDev = stdDev
  return extra
}

function iconSetIsExt(rule) {
  return Boolean(rule && rule.type === 'iconSet' && (rule.custom || (Array.isArray(rule.icons) && rule.icons.length) || EXT_ICON_SETS.has(rule.iconSet)))
}

function normalizeDxfStyle(style) {
  if (!style || typeof style !== 'object') return style
  const result = {}
  for (const [key, value] of Object.entries(style)) {
    if (value === undefined || value === null) continue
    if (key === 'numFmt' && typeof value === 'object') {
      if (typeof value.formatCode === 'string' && value.formatCode) result.numFmt = value.formatCode
      continue
    }
    result[key] = value
  }
  return result
}

function installConditionalFormattingPatches() {
  let CfRuleXform
  let CfvoXform
  let DatabarXform
  let IconSetXform
  let ConditionalFormattingsXform
  let CfRuleExtXform
  let CfvoExtXform
  let DatabarExtXform
  let IconSetExtXform
  let ColorXform
  let DxfXform
  let FExtXform
  try {
    CfRuleXform = require('exceljs/lib/xlsx/xform/sheet/cf/cf-rule-xform')
    CfvoXform = require('exceljs/lib/xlsx/xform/sheet/cf/cfvo-xform')
    DatabarXform = require('exceljs/lib/xlsx/xform/sheet/cf/databar-xform')
    IconSetXform = require('exceljs/lib/xlsx/xform/sheet/cf/icon-set-xform')
    ConditionalFormattingsXform = require('exceljs/lib/xlsx/xform/sheet/cf/conditional-formattings-xform')
    CfRuleExtXform = require('exceljs/lib/xlsx/xform/sheet/cf-ext/cf-rule-ext-xform')
    CfvoExtXform = require('exceljs/lib/xlsx/xform/sheet/cf-ext/cfvo-ext-xform')
    DatabarExtXform = require('exceljs/lib/xlsx/xform/sheet/cf-ext/databar-ext-xform')
    IconSetExtXform = require('exceljs/lib/xlsx/xform/sheet/cf-ext/icon-set-ext-xform')
    FExtXform = require('exceljs/lib/xlsx/xform/sheet/cf-ext/f-ext-xform')
    ColorXform = require('exceljs/lib/xlsx/xform/style/color-xform')
    DxfXform = require('exceljs/lib/xlsx/xform/style/dxf-xform')
  } catch {
    // A future ExcelJS layout: keep its own (lossy) conditional-format handling.
    return false
  }
  if (CfRuleXform[PATCHED]) return true

  // ---- <cfvo> (main part) -------------------------------------------------------------------
  CfvoXform.prototype.parseOpen = function parseOpen(node) {
    const attributes = node.attributes || {}
    this.model = { type: attributes.type }
    const value = cfvoValue(attributes.val)
    if (value !== undefined) this.model.value = value
    const gte = toBool(attributes.gte)
    if (gte === false) this.model.gte = false
  }
  CfvoXform.prototype.render = function render(xmlStream, model) {
    // The 2007 schema has no autoMin/autoMax: those live in the x14 extension.
    const type = model.type === 'autoMin' ? 'min' : model.type === 'autoMax' ? 'max' : model.type
    xmlStream.leafNode(this.tag, {
      type,
      val: model.value === undefined || model.value === null || type === 'min' || type === 'max' ? undefined : String(model.value),
      gte: model.gte === false ? '0' : undefined,
    })
  }

  // ---- <dataBar> (main part) ----------------------------------------------------------------
  DatabarXform.prototype.createNewModel = function createNewModel(node) {
    const attributes = (node && node.attributes) || {}
    const model = { cfvo: [] }
    const minLength = toInt(attributes.minLength)
    const maxLength = toInt(attributes.maxLength)
    const showValue = toBool(attributes.showValue)
    if (minLength !== undefined) model.minLength = minLength
    if (maxLength !== undefined) model.maxLength = maxLength
    if (showValue === false) model.showValue = false
    return model
  }
  DatabarXform.prototype.render = function render(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      minLength: model.minLength === undefined ? undefined : String(model.minLength),
      maxLength: model.maxLength === undefined ? undefined : String(model.maxLength),
      showValue: model.showValue === false ? '0' : undefined,
    })
    const cfvo = Array.isArray(model.cfvo) && model.cfvo.length >= 2 ? model.cfvo.slice(0, 2) : [{ type: 'min' }, { type: 'max' }]
    cfvo.forEach((item) => this.cfvoXform.render(xmlStream, item))
    this.colorXform.render(xmlStream, model.color && typeof model.color === 'object' && !Array.isArray(model.color) ? model.color : { argb: 'FF638EC6' })
    xmlStream.closeNode()
  }

  // ---- <iconSet> (main part) ----------------------------------------------------------------
  IconSetXform.prototype.createNewModel = function createNewModel({ attributes = {} }) {
    const model = { iconSet: attributes.iconSet || '3TrafficLights1', cfvo: [] }
    const reverse = toBool(attributes.reverse)
    const showValue = toBool(attributes.showValue)
    const percent = toBool(attributes.percent)
    if (reverse) model.reverse = true
    if (showValue === false) model.showValue = false
    if (percent === false) model.percent = false
    return model
  }
  IconSetXform.prototype.render = function render(xmlStream, model) {
    const iconSet = model.iconSet === '3TrafficLights' ? '3TrafficLights1' : model.iconSet
    xmlStream.openNode(this.tag, {
      iconSet: iconSet && iconSet !== '3TrafficLights1' ? iconSet : undefined,
      showValue: model.showValue === false ? '0' : undefined,
      percent: model.percent === false ? '0' : undefined,
      reverse: model.reverse ? '1' : undefined,
    })
    ;(Array.isArray(model.cfvo) ? model.cfvo : []).forEach((item) => this.cfvoXform.render(xmlStream, item))
    xmlStream.closeNode()
  }

  // ---- <cfRule> (main part) -----------------------------------------------------------------
  const originalCreateRule = CfRuleXform.prototype.createNewModel
  CfRuleXform.prototype.createNewModel = function createNewModel(node) {
    const model = originalCreateRule.call(this, node)
    return Object.assign(model, extraRuleAttributes((node && node.attributes) || {}))
  }

  CfRuleXform.isPrimitive = function isPrimitive(rule) {
    const type = canonicalType(rule)
    if (!MAIN_RENDERABLE.has(type)) return false
    if (type === 'iconSet') return !iconSetIsExt(rule)
    if (type === 'expression' || type === 'cellIs') return ruleFormulae(rule).length > 0
    if (TEXT_RULE_OPERATOR[type]) return ruleFormulae(rule).length > 0 || (rule.text !== undefined && rule.text !== null)
    return true
  }

  CfRuleXform.prototype.render = function render(xmlStream, model) {
    const type = canonicalType(model)
    if (!CfRuleXform.isPrimitive(model)) return
    const attributes = {
      type,
      dxfId: undefined,
      priority: model.priority,
      stopIfTrue: model.stopIfTrue ? '1' : undefined,
    }
    const styled = type !== 'colorScale' && type !== 'dataBar' && type !== 'iconSet'
    if (styled && model.dxfId !== undefined) attributes.dxfId = model.dxfId
    let formulae = []
    switch (type) {
      case 'expression':
        formulae = ruleFormulae(model).slice(0, 1)
        break
      case 'cellIs':
        attributes.operator = model.operator || 'equal'
        formulae = ruleFormulae(model).slice(0, model.operator === 'between' || model.operator === 'notBetween' ? 2 : 1)
        break
      case 'top10':
        attributes.percent = model.percent ? '1' : undefined
        attributes.bottom = model.bottom ? '1' : undefined
        attributes.rank = String(Number.isFinite(Number(model.rank)) && model.rank !== undefined ? Math.floor(Number(model.rank)) : 10)
        break
      case 'aboveAverage':
        attributes.aboveAverage = model.aboveAverage === false ? '0' : undefined
        attributes.equalAverage = model.equalAverage ? '1' : undefined
        attributes.stdDev = model.stdDev ? String(model.stdDev) : undefined
        break
      case 'containsText':
      case 'notContainsText':
      case 'beginsWith':
      case 'endsWith':
        attributes.operator = TEXT_RULE_OPERATOR[type]
        attributes.text = model.text === undefined || model.text === null ? undefined : String(model.text)
        formulae = ruleFormulae(model).slice(0, 1)
        if (!formulae.length && attributes.text !== undefined) formulae = [derivedFormula(type, model)]
        break
      case 'containsBlanks':
      case 'notContainsBlanks':
      case 'containsErrors':
      case 'notContainsErrors':
        formulae = ruleFormulae(model).slice(0, 1)
        if (!formulae.length) formulae = [derivedFormula(type, model)]
        break
      case 'timePeriod':
        attributes.timePeriod = model.timePeriod
        formulae = ruleFormulae(model).slice(0, 1)
        if (!formulae.length) formulae = [derivedFormula(type, model)].filter(Boolean)
        break
      default:
        break
    }
    xmlStream.openNode(this.tag, attributes)
    if (type === 'colorScale') this.colorScaleXform.render(xmlStream, { cfvo: model.cfvo || [], color: model.color || [] })
    else if (type === 'dataBar') {
      this.databarXform.render(xmlStream, model)
      this.extLstRefXform.render(xmlStream, model)
    } else if (type === 'iconSet') this.iconSetXform.render(xmlStream, model)
    formulae.forEach((formula) => this.formulaXform.render(xmlStream, formula))
    xmlStream.closeNode()
  }

  // ---- dxf styles, priorities and ext cfvo promotion ------------------------------------------
  const originalPrepare = ConditionalFormattingsXform.prototype.prepare
  ConditionalFormattingsXform.prototype.prepare = function prepare(model, options) {
    ;(Array.isArray(model) ? model : []).forEach((cf) => {
      if (!cf || !Array.isArray(cf.rules)) return
      cf.rules.forEach((rule) => {
        if (!rule || typeof rule !== 'object') return
        const type = canonicalType(rule)
        if (rule.style && (type === 'colorScale' || type === 'dataBar' || type === 'iconSet')) delete rule.style
        if (rule.style) rule.style = normalizeDxfStyle(rule.style)
      })
    })
    return originalPrepare.call(this, model, options)
  }

  const originalReconcile = ConditionalFormattingsXform.prototype.reconcile
  ConditionalFormattingsXform.prototype.reconcile = function reconcile(model, options) {
    originalReconcile.call(this, model, options)
    ;(Array.isArray(model) ? model : []).forEach((cf) => {
      if (!cf || !Array.isArray(cf.rules)) return
      cf.rules.forEach((rule) => {
        if (!rule || typeof rule !== 'object') return
        if (rule.style) rule.style = normalizeDxfStyle(rule.style)
        // The x14 thresholds are the precise ones (autoMin/autoMax, formula text).
        if (Array.isArray(rule.extCfvo)) {
          if (rule.extCfvo.length) rule.cfvo = rule.extCfvo
          delete rule.extCfvo
        }
        if (rule.extStyle) {
          if (!rule.style) rule.style = normalizeDxfStyle(rule.extStyle)
          delete rule.extStyle
        }
        if (rule.type !== 'dataBar' && !iconSetIsExt(rule)) delete rule.x14Id
      })
    })
  }

  // ---- x14 extension ------------------------------------------------------------------------
  DatabarExtXform.isExt = function isExt() {
    // Excel 2010+ always pairs a data bar with its x14 definition (axis, negative bars, borders).
    return true
  }
  const originalIsExt = CfRuleExtXform.isExt
  CfRuleExtXform.isExt = function isExt(rule) {
    if (rule && rule.type === 'dataBar') return true
    if (rule && rule.type === 'iconSet') return iconSetIsExt(rule)
    return originalIsExt.call(this, rule)
  }

  CfvoExtXform.prototype.createNewModel = function createNewModel(node) {
    const attributes = (node && node.attributes) || {}
    const model = { type: attributes.type }
    const gte = toBool(attributes.gte)
    if (gte === false) model.gte = false
    return model
  }
  CfvoExtXform.prototype.onParserClose = function onParserClose(name, parser) {
    if (name === 'xm:f') {
      const value = cfvoValue(parser.model)
      if (value !== undefined) this.model.value = value
    }
  }
  CfvoExtXform.prototype.render = function render(xmlStream, model) {
    xmlStream.openNode(this.tag, { type: model.type, gte: model.gte === false ? '0' : undefined })
    const bare = model.type === 'min' || model.type === 'max' || model.type === 'autoMin' || model.type === 'autoMax'
    if (!bare && model.value !== undefined && model.value !== null) this.fExtXform.render(xmlStream, String(model.value))
    xmlStream.closeNode()
  }

  const DATABAR_EXT_BOOLEANS = ['border', 'gradient', 'negativeBarColorSameAsPositive', 'negativeBarBorderColorSameAsPositive']
  const originalDatabarExtParseOpen = DatabarExtXform.prototype.parseOpen
  DatabarExtXform.prototype.parseOpen = function parseOpen(node) {
    if (!this.map['x14:fillColor']) this.map['x14:fillColor'] = new ColorXform('x14:fillColor')
    return originalDatabarExtParseOpen.call(this, node)
  }
  DatabarExtXform.prototype.createNewModel = function createNewModel({ attributes = {} }) {
    // Only attributes that are present: the x14 schema defaults differ from ExcelJS's.
    const model = { extCfvo: [] }
    const minLength = toInt(attributes.minLength)
    const maxLength = toInt(attributes.maxLength)
    if (minLength !== undefined) model.minLength = minLength
    if (maxLength !== undefined) model.maxLength = maxLength
    for (const key of DATABAR_EXT_BOOLEANS) {
      const value = toBool(attributes[key])
      if (value !== undefined) model[key] = value
    }
    if (attributes.axisPosition) model.axisPosition = attributes.axisPosition
    if (attributes.direction) model.direction = attributes.direction
    return model
  }
  DatabarExtXform.prototype.onParserClose = function onParserClose(name, parser) {
    const prop = name.split(':')[1]
    if (prop === 'cfvo') this.model.extCfvo.push(parser.model)
    else if (prop === 'fillColor') this.model.color = parser.model
    else if (parser.model) this.model[prop] = parser.model
  }
  DatabarExtXform.prototype.render = function render(xmlStream, model) {
    const attributes = {
      minLength: model.minLength === undefined ? undefined : String(model.minLength),
      maxLength: model.maxLength === undefined ? undefined : String(model.maxLength),
    }
    for (const key of DATABAR_EXT_BOOLEANS) attributes[key] = boolAttribute(model[key])
    attributes.axisPosition = model.axisPosition && model.axisPosition !== 'auto' ? model.axisPosition : undefined
    attributes.direction = model.direction || undefined
    xmlStream.openNode(this.tag, attributes)
    const cfvo = Array.isArray(model.cfvo) && model.cfvo.length >= 2 ? model.cfvo.slice(0, 2) : [{ type: 'autoMin' }, { type: 'autoMax' }]
    cfvo.forEach((item) => this.cfvoXform.render(xmlStream, item))
    // CT_DataBar (x14) child order: cfvo, cfvo, fillColor?, borderColor?, negativeFillColor?, negativeBorderColor?, axisColor?
    this.borderColorXform.render(xmlStream, model.borderColor)
    this.negativeFillColorXform.render(xmlStream, model.negativeFillColor)
    this.negativeBorderColorXform.render(xmlStream, model.negativeBorderColor)
    this.axisColorXform.render(xmlStream, model.axisColor)
    xmlStream.closeNode()
  }

  IconSetExtXform.prototype.createNewModel = function createNewModel({ attributes = {} }) {
    const model = { cfvo: [], iconSet: attributes.iconSet || '3TrafficLights1' }
    const reverse = toBool(attributes.reverse)
    const showValue = toBool(attributes.showValue)
    if (reverse) model.reverse = true
    if (showValue === false) model.showValue = false
    // custom="1" is implied by the parsed `icons` list (written back when icons exist).
    return model
  }
  IconSetExtXform.prototype.render = function render(xmlStream, model) {
    // ExcelJS overwrote each custom icon's iconId with its position; keep the real id.
    const icons = Array.isArray(model.icons) && model.icons.length ? model.icons : undefined
    const iconSet = model.iconSet === '3TrafficLights' ? '3TrafficLights1' : model.iconSet
    xmlStream.openNode(this.tag, {
      iconSet: iconSet && iconSet !== '3TrafficLights1' ? iconSet : undefined,
      showValue: model.showValue === false ? '0' : undefined,
      reverse: model.reverse ? '1' : undefined,
      custom: icons ? '1' : undefined,
    })
    ;(Array.isArray(model.cfvo) ? model.cfvo : []).forEach((item) => this.cfvoXform.render(xmlStream, item))
    if (icons) {
      icons.forEach((icon) => this.cfIconXform.render(xmlStream, {
        iconSet: icon && icon.iconSet ? icon.iconSet : 'NoIcons',
        iconId: Math.max(0, Math.floor(Number(icon && icon.iconId) || 0)),
      }))
    }
    xmlStream.closeNode()
  }

  // x14:cfRule of other types (Excel 2010 writes cross-sheet expression/cellIs rules here).
  class X14DxfXform extends DxfXform {
    get tag() {
      return 'x14:dxf'
    }
  }
  const originalExtRuleParseOpen = CfRuleExtXform.prototype.parseOpen
  CfRuleExtXform.prototype.parseOpen = function parseOpen(node) {
    if (!this.map['xm:f']) this.map['xm:f'] = new FExtXform()
    if (!this.map['x14:dxf']) this.map['x14:dxf'] = new X14DxfXform()
    return originalExtRuleParseOpen.call(this, node)
  }
  const BLANK_ERROR_TYPES = new Set(['containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'])
  CfRuleExtXform.prototype.createNewModel = function createNewModel({ attributes = {} }) {
    const type = attributes.type
    // Same model conventions as the main part (ExcelJS's opType).
    const model = BLANK_ERROR_TYPES.has(type)
      ? { type: 'containsText', operator: type }
      : type === 'containsText'
        ? { type, operator: 'containsText' }
        : { type, ...(attributes.operator ? { operator: attributes.operator } : {}) }
    model.x14Id = attributes.id
    model.priority = toInt(attributes.priority)
    if (attributes.timePeriod) model.timePeriod = attributes.timePeriod
    const rank = toInt(attributes.rank)
    if (rank !== undefined) model.rank = rank
    const percent = toBool(attributes.percent)
    if (percent !== undefined) model.percent = percent
    const bottom = toBool(attributes.bottom)
    if (bottom !== undefined) model.bottom = bottom
    const aboveAverage = toBool(attributes.aboveAverage)
    if (aboveAverage === false) model.aboveAverage = false
    return Object.assign(model, extraRuleAttributes(attributes))
  }
  CfRuleExtXform.prototype.onParserClose = function onParserClose(name, parser) {
    if (name === 'xm:f') {
      this.model.formulae = this.model.formulae || []
      this.model.formulae.push(parser.model)
    } else if (name === 'x14:dxf') {
      this.model.extStyle = parser.model
    } else {
      Object.assign(this.model, parser.model)
    }
  }

  CfRuleXform[PATCHED] = true
  return true
}

module.exports = { installConditionalFormattingPatches }
