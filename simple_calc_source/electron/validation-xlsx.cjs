'use strict'

// Data-validation fidelity fixes for ExcelJS:
// - bounds of whole/decimal/textLength/date rules that are formulas or references (=TODAY(),
//   $B$1) were parsed with parseInt/parseFloat and read back as NaN;
// - date bounds were always written through `new Date(formula)`, so a formula or a serial
//   number became NaN or a 1970 date;
// - showDropDown (Excel's "hide the in-cell dropdown" flag) was dropped on read and write;
// - every sqref range was expanded into one model entry per cell, so a dropdown on a whole
//   column produced a million entries and the importer had to truncate the rules (edits to
//   such sheets could then not be saved). Ranges are kept as range keys ("D2:D20001"), which
//   the editor and the ExcelJS writer both understand.

const DataValidationsXform = require('exceljs/lib/xlsx/xform/sheet/data-validations-xform')
const DataValidations = require('exceljs/lib/doc/data-validations')
const utils = require('exceljs/lib/utils/utils')

let patched = false
const NUMBER = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/
const CELL = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/i
const MAX_ROW = 1_048_576
const MAX_COL = 16_384

function columnNumber(letters) {
  let value = 0
  for (const character of letters.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value
}

function columnLetters(number) {
  let value = number
  let result = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    value = Math.floor((value - 1) / 26)
  }
  return result
}

/** "$d$20:b2" -> "B2:D20"; a single cell stays "A1". Returns null for anything else. */
function normalizeValidationKey(value) {
  const text = String(value == null ? '' : value).trim()
  if (!text) return null
  const parts = text.split(':')
  if (parts.length > 2) return null
  const points = parts.map((part) => {
    const match = CELL.exec(part.trim())
    if (!match) return null
    const col = columnNumber(match[1])
    const row = Number(match[2])
    return col >= 1 && col <= MAX_COL && row >= 1 && row <= MAX_ROW ? { row, col } : null
  })
  if (points.some((point) => !point)) return null
  if (points.length === 1) return `${columnLetters(points[0].col)}${points[0].row}`
  const top = Math.min(points[0].row, points[1].row)
  const bottom = Math.max(points[0].row, points[1].row)
  const left = Math.min(points[0].col, points[1].col)
  const right = Math.max(points[0].col, points[1].col)
  return `${columnLetters(left)}${top}:${columnLetters(right)}${bottom}`
}

function keyBounds(key) {
  const normalized = normalizeValidationKey(key)
  if (!normalized) return null
  const [start, end = start] = normalized.split(':')
  const a = CELL.exec(start)
  const b = CELL.exec(end)
  return { top: Number(a[2]), left: columnNumber(a[1]), bottom: Number(b[2]), right: columnNumber(b[1]) }
}

/** Rebuild a validation model with normalized keys (single cells or tl:br ranges), dropping invalid ones. */
function normalizeValidationModel(model) {
  const result = {}
  for (const [rawKey, validation] of Object.entries(model || {})) {
    if (!validation || typeof validation !== 'object') continue
    for (const area of String(rawKey).split(/\s+/).filter(Boolean)) {
      const key = normalizeValidationKey(area)
      if (key && !result[key]) result[key] = validation
    }
  }
  return result
}

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
    if (name === 'dataValidation' && this._dataValidation && this.model) {
      if (!this._dataValidation.formulae || !this._dataValidation.formulae.length) {
        delete this._dataValidation.formulae
        delete this._dataValidation.operator
      }
      // One entry per sqref area (not per cell); multi-area sqrefs share the rule.
      for (const area of String(this._address || '').split(/\s+/).filter(Boolean)) {
        const key = normalizeValidationKey(area)
        if (key) this.model[key] = this._dataValidation
      }
      return true
    }
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

  // cell.dataValidation looks a cell up by its exact address; with range keys it also has to
  // find the range that contains the cell.
  const originalFind = DataValidations.prototype.find
  DataValidations.prototype.find = function find(address) {
    const direct = originalFind.call(this, address)
    if (direct) return direct
    const target = keyBounds(address)
    if (!target) return undefined
    for (const [key, validation] of Object.entries(this.model || {})) {
      if (!validation || !key.includes(':')) continue
      const bounds = keyBounds(key)
      if (bounds && target.top >= bounds.top && target.top <= bounds.bottom && target.left >= bounds.left && target.left <= bounds.right) return validation
    }
    return undefined
  }

  const originalRender = proto.render
  proto.render = function render(xmlStream, model) {
    // ExcelJS's range grouping indexes the model by its exact keys, so they must be normalized.
    model = normalizeValidationModel(model)
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

module.exports = { installValidationPatches, normalizeValidationKey, normalizeValidationModel }
