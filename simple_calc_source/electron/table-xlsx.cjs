'use strict'

// Excel tables (ListObjects): import into the editor model, write the model back through
// ExcelJS, and fix ExcelJS's reading of Excel-authored tables.

const Table = require('exceljs/lib/doc/table')
const TableXform = require('exceljs/lib/xlsx/xform/table/table-xform')
const AutoFilterXform = require('exceljs/lib/xlsx/xform/table/auto-filter-xform')

let patched = false

/**
 * Excel omits headerRowCount when a table has its default single header row; ExcelJS reads
 * only headerRowCount="1" as a header and would save such tables with headerRowCount="0",
 * turning their header labels into data. Treat an absent attribute as the default.
 */
function installTablePatches() {
  if (patched) return
  patched = true
  const originalParseOpen = TableXform.prototype.parseOpen
  TableXform.prototype.parseOpen = function parseOpen(node) {
    const handled = originalParseOpen.call(this, node)
    if (node && node.name === 'table' && this.model && !this.parser) {
      this.model.headerRow = node.attributes.headerRowCount !== '0'
      this.model.totalsRow = Number(node.attributes.totalsRowCount || 0) > 0
    }
    return handled
  }
  // A table whose filter buttons are turned off has no <autoFilter> element at all.
  const originalAutoFilterRender = AutoFilterXform.prototype.render
  AutoFilterXform.prototype.render = function render(xmlStream, model) {
    if (!model || !model.autoFilterRef) return false
    return originalAutoFilterRender.call(this, xmlStream, model)
  }
  const originalAutoFilterPrepare = AutoFilterXform.prototype.prepare
  AutoFilterXform.prototype.prepare = function prepare(model) {
    if (!model || !model.autoFilterRef) return
    return originalAutoFilterPrepare.call(this, model)
  }
}

const RANGE_RE = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i

function columnNumber(label) {
  let value = 0
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value
}

function columnLabel(number) {
  let value = number
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function decodeRange(range) {
  const match = RANGE_RE.exec(String(range || '').trim())
  if (!match) return null
  const startColumn = columnNumber(match[1])
  const startRow = Number(match[2])
  const endColumn = match[3] ? columnNumber(match[3]) : startColumn
  const endRow = match[4] ? Number(match[4]) : startRow
  return {
    top: Math.min(startRow, endRow),
    bottom: Math.max(startRow, endRow),
    left: Math.min(startColumn, endColumn),
    right: Math.max(startColumn, endColumn),
  }
}

function encodeRange(bounds) {
  return `${columnLabel(bounds.left)}${bounds.top}:${columnLabel(bounds.right)}${bounds.bottom}`
}

function tableModelOf(table) {
  return (table && (table.model || table.table)) || null
}

/** Editor-model tables for a loaded ExcelJS worksheet. */
function tablesFromWorksheet(worksheet) {
  const tables = typeof worksheet.getTables === 'function' ? worksheet.getTables() : Object.values(worksheet.tables || {})
  const output = []
  for (const table of tables) {
    const model = tableModelOf(table)
    if (!model || !model.name) continue
    const ref = model.tableRef || model.ref
    if (!decodeRange(ref)) continue
    output.push({
      id: `table-${model.name}`,
      name: String(model.name),
      displayName: model.displayName ? String(model.displayName) : undefined,
      ref: String(ref).replace(/\$/g, '').toUpperCase(),
      headerRow: model.headerRow !== false,
      totalsRow: Boolean(model.totalsRow),
      ...(model.headerRow !== false && !model.autoFilterRef ? { showFilterButton: false } : {}),
      columns: (model.columns || []).map((column) => ({
        name: String(column.name == null ? '' : column.name),
        ...(column.totalsRowFunction ? { totalsRowFunction: String(column.totalsRowFunction) } : {}),
        ...(column.totalsRowLabel != null ? { totalsRowLabel: String(column.totalsRowLabel) } : {}),
        ...(column.totalsRowFormula ? { totalsRowFormula: String(column.totalsRowFormula) } : {}),
      })),
      style: model.style ? {
        theme: model.style.theme || undefined,
        showRowStripes: Boolean(model.style.showRowStripes),
        showColumnStripes: Boolean(model.style.showColumnStripes),
        showFirstColumn: Boolean(model.style.showFirstColumn),
        showLastColumn: Boolean(model.style.showLastColumn),
      } : undefined,
      imported: true,
    })
  }
  return output
}

/**
 * Make the ExcelJS worksheet's tables match the editor model. Tables loaded from the source
 * package keep their extra column attributes (dxf styles, calculated column formulas).
 */
function applyTablesToWorksheet(worksheet, modelTables) {
  if (!Array.isArray(modelTables)) return
  const existing = new Map()
  const current = typeof worksheet.getTables === 'function' ? worksheet.getTables() : Object.values(worksheet.tables || {})
  for (const table of current) {
    const model = tableModelOf(table)
    if (model && model.name) existing.set(String(model.name).toLowerCase(), model)
  }
  const next = {}
  for (const table of modelTables) {
    if (!table || typeof table.name !== 'string' || !table.name) continue
    const bounds = decodeRange(table.ref)
    if (!bounds) continue
    const width = bounds.right - bounds.left + 1
    const previous = existing.get(table.name.toLowerCase())
    const previousColumns = previous && Array.isArray(previous.columns) ? previous.columns : []
    const columns = Array.from({ length: width }, (_, index) => {
      const source = table.columns && table.columns[index] ? table.columns[index] : {}
      const kept = previousColumns[index] && previousColumns[index].name === source.name ? { ...previousColumns[index] } : {}
      const column = {
        ...kept,
        name: source.name || `Column${index + 1}`,
        filterButton: table.headerRow !== false && table.showFilterButton !== false,
      }
      if (source.totalsRowFunction) column.totalsRowFunction = source.totalsRowFunction
      else delete column.totalsRowFunction
      if (source.totalsRowLabel != null) column.totalsRowLabel = source.totalsRowLabel
      else delete column.totalsRowLabel
      if (source.totalsRowFormula) column.totalsRowFormula = source.totalsRowFormula
      else if (!source.totalsRowFunction || source.totalsRowFunction !== 'custom') delete column.totalsRowFormula
      return column
    })
    const filterBottom = table.totalsRow ? bounds.bottom - 1 : bounds.bottom
    const style = table.style || {}
    const model = {
      ...(previous || {}),
      name: table.name,
      displayName: table.displayName || table.name,
      tableRef: encodeRange(bounds),
      ref: encodeRange(bounds),
      headerRow: table.headerRow !== false,
      totalsRow: Boolean(table.totalsRow),
      autoFilterRef: table.headerRow !== false && table.showFilterButton !== false ? encodeRange({ ...bounds, bottom: Math.max(bounds.top, filterBottom) }) : undefined,
      columns,
      style: {
        ...((previous && previous.style) || {}),
        theme: style.theme || (previous && previous.style && previous.style.theme) || 'TableStyleMedium2',
        showFirstColumn: Boolean(style.showFirstColumn),
        showLastColumn: Boolean(style.showLastColumn),
        showRowStripes: style.showRowStripes !== false,
        showColumnStripes: Boolean(style.showColumnStripes),
      },
    }
    const instance = new Table()
    instance.worksheet = worksheet
    instance.model = model
    next[table.name] = instance
  }
  worksheet.tables = next
}

module.exports = { installTablePatches, tablesFromWorksheet, applyTablesToWorksheet, decodeRange, encodeRange }
