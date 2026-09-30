'use strict'

// Row/column outline fidelity for ExcelJS. ExcelJS derives `collapsed` from the outline level
// (`outlineLevel >= sheet.outlineLevelRow`) instead of reading the attribute, so every detail
// row at the deepest level is read and written as collapsed, and the real flag on summary rows
// is lost. Keep the attribute as stored, and write the sheet's outline depth that Excel uses
// to size the outline bar.

const Row = require('exceljs/lib/doc/row')
const Column = require('exceljs/lib/doc/column')

let patched = false

function installOutlinePatches() {
  if (patched) return
  patched = true
  for (const Klass of [Row, Column]) {
    Object.defineProperty(Klass.prototype, 'collapsed', {
      configurable: true,
      get() { return Boolean(this._storedCollapsed) },
      set(value) { this._storedCollapsed = Boolean(value) },
    })
  }

  const rowModel = Object.getOwnPropertyDescriptor(Row.prototype, 'model')
  Object.defineProperty(Row.prototype, 'model', {
    configurable: true,
    // ExcelJS only writes rows with a height or cells; hidden, grouped, collapsed or styled
    // empty rows (spacers, collapsed detail) were silently dropped.
    get() {
      const model = rowModel.get.call(this)
      if (model) return model
      const style = this.style || {}
      const styled = Boolean(style.font || style.numFmt || style.alignment || style.border || style.fill || style.protection)
      if (!this.hidden && !this.outlineLevel && !this.collapsed && !styled) return null
      return { cells: [], number: this.number, min: 0, max: 0, height: this.height, style: this.style, hidden: this.hidden, outlineLevel: this.outlineLevel, collapsed: this.collapsed }
    },
    set(value) {
      rowModel.set.call(this, value)
      this._storedCollapsed = Boolean(value && value.collapsed)
    },
  })

  const columnDefn = Object.getOwnPropertyDescriptor(Column.prototype, 'defn')
  Object.defineProperty(Column.prototype, 'defn', {
    configurable: true,
    get: columnDefn.get,
    set(value) {
      columnDefn.set.call(this, value)
      this._storedCollapsed = Boolean(value && value.collapsed)
    },
  })

  const isDefault = Object.getOwnPropertyDescriptor(Column.prototype, 'isDefault')
  Object.defineProperty(Column.prototype, 'isDefault', {
    configurable: true,
    get() { return isDefault.get.call(this) && !this.collapsed },
  })

  const equivalentTo = Column.prototype.equivalentTo
  Column.prototype.equivalentTo = function equivalent(other) {
    return equivalentTo.call(this, other) && Boolean(this.collapsed) === Boolean(other && other.collapsed)
  }
}

/** Write the model's outline levels, collapsed flags and depths onto an ExcelJS worksheet. */
function applyOutlineToWorksheet(worksheet, sheet) {
  let rowDepth = 0
  let columnDepth = 0
  for (const [key, properties] of Object.entries(sheet.rowProperties || {})) {
    const index = Number(key)
    if (!Number.isInteger(index) || index < 1 || !properties || typeof properties !== 'object') continue
    const level = Math.min(7, Math.max(0, Math.floor(Number(properties.outlineLevel) || 0)))
    rowDepth = Math.max(rowDepth, level)
    if (properties.collapsed === true || level) worksheet.getRow(index).collapsed = properties.collapsed === true
  }
  for (const [key, properties] of Object.entries(sheet.columnProperties || {})) {
    const index = Number(key)
    if (!Number.isInteger(index) || index < 1 || !properties || typeof properties !== 'object') continue
    const level = Math.min(7, Math.max(0, Math.floor(Number(properties.outlineLevel) || 0)))
    columnDepth = Math.max(columnDepth, level)
    if (properties.collapsed === true || level) worksheet.getColumn(index).collapsed = properties.collapsed === true
  }
  worksheet.properties = { ...(worksheet.properties || {}), outlineLevelRow: rowDepth, outlineLevelCol: columnDepth }
}

module.exports = { installOutlinePatches, applyOutlineToWorksheet }
