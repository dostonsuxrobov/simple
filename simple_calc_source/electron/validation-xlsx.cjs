'use strict'

// Data-validation fidelity fixes for ExcelJS:
// - bounds of whole/decimal/textLength/date rules that are formulas or references (=TODAY(),
//   $B$1) were parsed with parseInt/parseFloat and read back as NaN;
// - date bounds were always written through `new Date(formula)`, so a formula or a serial
//   number became NaN or a 1970 date;
// - showDropDown (Excel's "hide the in-cell dropdown" flag) was dropped on read and write.

const DataValidationsXform = require('exceljs/lib/xlsx/xform/sheet/data-validations-xform')
const utils = require('exceljs/lib/utils/utils')

let patched = false
const NUMBER = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/

function dateBoundText(formula) {
  if (formula instanceof Date) return Number.isNaN(formula.getTime()) ? null : String(utils.dateToExcel(formula))
  if (formula && typeof formula === 'object' && formula.type === 'date' && typeof formula.value === 'string') {
    const date = new Date(formula.value)
    return Number.isNaN(date.getTime()) ? null : String(utils.dateToExcel(date))
  }
  if (typeof formula === 'number') return Number.isFinite(formula) ? String(formula) : null
  if (typeof formula === 'string') {
    const text = formula.trim().replace(/^=/, '')
    if (!text) return null
    if (NUMBER.test(text)) return text
    // ISO date text from a plain clone of a Date.
    if (/^\d{4}-\d{2}-\d{2}T/.test(text)) {
      const date = new Date(text)
      if (!Number.isNaN(date.getTime())) return String(utils.dateToExcel(date))
    }
    return text
  }
  return null
}

function boundText(validation, formula) {
  if (validation && validation.type === 'date') return dateBoundText(formula)
  if (typeof formula === 'string') return formula.trim().replace(/^=/, '')
  return formula == null ? null : String(formula)
}

function sourceValidation(model, sqref) {
  if (!model || !sqref) return null
  if (model[sqref]) return model[sqref]
  const first = String(sqref).split(/[\s:]/)[0]
  if (model[first]) return model[first]
  for (const [key, value] of Object.entries(model)) {
    if (key.split(/\s+/).includes(sqref) || key.split(/\s+/)[0].split(':')[0] === first) return value
  }
  return null
}

function installValidationPatches() {
  if (patched) return
  patched = true
  const proto = DataValidationsXform.prototype

  const originalParseOpen = proto.parseOpen
  proto.parseOpen = function parseOpen(node) {
    const handled = originalParseOpen.call(this, node)
    if (node && node.name === 'dataValidation' && this._dataValidation && node.attributes.showDropDown !== undefined) {
      this._dataValidation.showDropDown = utils.parseBoolean(node.attributes.showDropDown)
    }
    return handled
  }

  const originalParseClose = proto.parseClose
  proto.parseClose = function parseClose(name) {
    if ((name === 'formula1' || name === 'formula2') && this._formula && this._dataValidation) {
      const text = this._formula.join('')
      const type = this._dataValidation.type
      if (!NUMBER.test(text) && ['whole', 'decimal', 'textLength', 'date'].includes(type)) {
        this._dataValidation.formulae.push(text)
        this._formula = undefined
        return true
      }
    }
    return originalParseClose.call(this, name)
  }

  const originalRender = proto.render
  proto.render = function render(xmlStream, model) {
    // Let ExcelJS group cells into ranges, then replay its output with corrected bounds.
    const events = []
    const recorder = {
      openNode: (name, attributes) => events.push(['open', name, attributes]),
      addAttribute: (name, value) => events.push(['attr', name, value]),
      writeText: (text) => events.push(['text', text]),
      closeNode: () => events.push(['close']),
    }
    originalRender.call(this, recorder, model)
    let validation = null
    let formulaIndex = -1
    for (const [kind, first, second] of events) {
      if (kind === 'open') {
        if (first === 'dataValidation') validation = null
        formulaIndex = /^formula([12])$/.test(first) ? Number(first.slice(-1)) - 1 : -1
        xmlStream.openNode(first, second)
      } else if (kind === 'attr') {
        xmlStream.addAttribute(first, second)
        if (first === 'sqref') {
          validation = sourceValidation(model, second)
          if (validation && validation.showDropDown) xmlStream.addAttribute('showDropDown', '1')
        }
      } else if (kind === 'text') {
        if (formulaIndex >= 0 && validation && Array.isArray(validation.formulae)) {
          const text = boundText(validation, validation.formulae[formulaIndex])
          xmlStream.writeText(text == null ? first : text)
        } else xmlStream.writeText(first)
      } else {
        xmlStream.closeNode()
        formulaIndex = -1
      }
    }
  }
}

module.exports = { installValidationPatches }
