'use strict'

const fs = require('fs')
const path = require('path')
// ExcelJS's builtin number-format table has no entry for the locale-reserved ids Excel
// reserves for its Currency / Comma / Accounting cell styles, and those ids are never
// written into <numFmts>.  Without them getStyleModel leaves cell.numFmt undefined and the
// format is lost, so a workbook formatted with the toolbar's comma or accounting button
// renders as raw digits.  Seeding the table before exceljs loads also keeps its reverse
// lookup intact so the ids survive a save.
const RESERVED_BUILTIN_NUMBER_FORMATS = Object.freeze({
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
})

try {
  const defaultNumberFormats = require('exceljs/lib/xlsx/defaultnumformats')
  for (const [id, format] of Object.entries(RESERVED_BUILTIN_NUMBER_FORMATS)) {
    if (!defaultNumberFormats[id] || !defaultNumberFormats[id].f) defaultNumberFormats[id] = { f: format }
  }
} catch {
  // A future ExcelJS layout can hide the table; the reader still works without the seed.
}

const ExcelJS = require('exceljs')
// Conditional formatting: ExcelJS 4.4 drops several rule types and attributes on read/write
// (duplicate/unique/beginsWith/endsWith rules, stopIfTrue, text, formula thresholds, x14 data
// bars...). The idempotent prototype patches in this module make Excel files round-trip.
require('./conditional-format-exceljs.cjs').installConditionalFormattingPatches()
const { installTablePatches, tablesFromWorksheet, applyTablesToWorksheet } = require('./table-xlsx.cjs')
const { installValidationPatches, normalizeValidationModel } = require('./validation-xlsx.cjs')
const { installOutlinePatches, applyOutlineToWorksheet } = require('./outline-xlsx.cjs')
const { imagesFromWorksheet, applyImagesToWorksheet, compactWorkbookMedia } = require('./image-xlsx.cjs')
const { installProtectionPatches } = require('./protection.cjs')
installTablePatches()
installValidationPatches()
installOutlinePatches()
installProtectionPatches()
const JSZip = require('jszip')
const XLSX = require('xlsx')
const { extractLegacyBiffStyles } = require('./legacy-biff-styles.cjs')
const { importChartsIntoSheets, writeChartsToPackage, workbookSheetParts, readRelationships, relsPathFor, renderRelationships } = require('./chart-xlsx.cjs')
const { importPivotDefinitions, writePivotPackage } = require('./pivot-xlsx.cjs')
const { importSparklineGroups, writeSparklinesToPackage } = require('./sparkline-xlsx.cjs')
const { readDelimited, writeDelimited } = require('./delimited-text.cjs')
const { officeEngineAvailable } = require('./office-converter.cjs')
installHyperlinkPatches()

// Formula cells that also carry a hyperlink: ExcelJS 4.4 turns them into plain hyperlink
// cells on read (the formula is lost). Keep the formula; the link is read separately from the
// sheet's <hyperlinks> (see captureWorksheetHyperlinks).
function installHyperlinkPatches() {
  const CellXform = require('exceljs/lib/xlsx/xform/sheet/cell-xform')
  if (CellXform.prototype.__simpleCalcHyperlinkPatch) return
  const originalReconcile = CellXform.prototype.reconcile
  CellXform.prototype.reconcile = function reconcile(model, options) {
    if (model && model.type === ExcelJS.ValueType.Formula && options && options.hyperlinkMap && options.hyperlinkMap[model.address]) {
      const hyperlinkMap = options.hyperlinkMap
      const target = hyperlinkMap[model.address]
      delete hyperlinkMap[model.address]
      try {
        return originalReconcile.call(this, model, options)
      } finally {
        hyperlinkMap[model.address] = target
      }
    }
    return originalReconcile.call(this, model, options)
  }
  CellXform.prototype.__simpleCalcHyperlinkPatch = true
}

// SheetJS is deliberately used for the long tail of spreadsheet formats.  The
// model itself is format-neutral and JSON serializable, so it can cross an
// Electron IPC boundary without retaining library-specific objects.
const SUPPORTED_EXTENSIONS = new Set([
  '.xlsx',
  '.xlsm',
  '.xlsb',
  '.xls',
  '.xltx',
  '.xltm',
  '.xlt',
  '.xlam',
  '.xla',
  '.xml',
  '.ods',
  '.fods',
  '.csv',
  '.tsv',
  '.tab',
  '.txt',
  '.numbers',
  '.slk',
  '.sylk',
  '.dif',
  '.dbf',
  '.prn',
  '.wk1',
  '.wk2',
  '.wk3',
  '.wk4',
  '.wks',
  '.wq1',
  '.wq2',
  '.wb1',
  '.wb2',
  '.wb3',
  '.123',
  '.qpw',
  '.html',
  '.htm',
])

const MODERN_FORMAT = 'xlsx'
const RICH_OOXML_FORMATS = new Set(['xlsx', 'xlsm', 'xltx', 'xltm', 'xlam'])
const MAX_SOURCE_BYTES = 512 * 1024 * 1024
const MAX_WORKBOOK_XML_BYTES = 16 * 1024 * 1024
const MAX_ADVANCED_METADATA_ENTRIES = 10_000
const MAX_VALIDATION_AREAS = 200_000
const MAX_METADATA_ROWS = 1_048_576
const MAX_METADATA_COLS = 16_384
const MAX_SHEET_NAME_LENGTH = 31
const CELL_ADDRESS_RE = /^\$?([A-Z]{1,7})\$?([1-9]\d*)$/i
const RANGE_RE = /^\$?([A-Z]{1,7})\$?([1-9]\d*):\$?([A-Z]{1,7})\$?([1-9]\d*)$/i

const ERROR_CODE_TO_TEXT = Object.freeze({
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
  43: '#GETTING_DATA',
  45: '#SPILL!',
  46: '#CONNECT!',
  47: '#BLOCKED!',
  48: '#UNKNOWN!',
  49: '#FIELD!',
  50: '#CALC!',
})
const ERROR_TEXT_TO_CODE = Object.freeze(
  Object.fromEntries(Object.entries(ERROR_CODE_TO_TEXT).map(([code, text]) => [text, Number(code)])),
)

function extensionFromName(name) {
  if (typeof name !== 'string' || !name.trim()) {
    throw new TypeError('A spreadsheet file name is required.')
  }
  return path.extname(name.trim()).toLowerCase()
}

function ensureSupportedExtension(name) {
  const extension = extensionFromName(name)
  if (!SUPPORTED_EXTENSIONS.has(extension)) {
    throw new Error(`Unsupported spreadsheet format: ${extension || '(no extension)'}`)
  }
  return extension
}

function bytesToBuffer(data) {
  let buffer
  if (Buffer.isBuffer(data)) {
    buffer = Buffer.from(data)
  } else if (data instanceof ArrayBuffer) {
    buffer = Buffer.from(new Uint8Array(data))
  } else if (ArrayBuffer.isView(data)) {
    buffer = Buffer.from(data.buffer, data.byteOffset, data.byteLength)
  } else if (Array.isArray(data)) {
    buffer = Buffer.from(data)
  } else if (data && data.type === 'Buffer' && Array.isArray(data.data)) {
    buffer = Buffer.from(data.data)
  } else {
    throw new TypeError('Spreadsheet data must be a Buffer, ArrayBuffer, or byte array.')
  }

  if (buffer.length > MAX_SOURCE_BYTES) {
    throw new RangeError(`Spreadsheet exceeds the ${MAX_SOURCE_BYTES / 1024 / 1024} MB import limit.`)
  }
  return buffer
}

function plainClone(value, depth = 0) {
  if (value == null || typeof value === 'string' || typeof value === 'boolean') return value
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value === 'bigint') return value.toString()
  if (value instanceof Date) return { type: 'date', value: value.toISOString() }
  if (Buffer.isBuffer(value) || value instanceof Uint8Array) {
    return { type: 'binary', value: Buffer.from(value).toString('base64') }
  }
  if (depth >= 8) return String(value)
  if (Array.isArray(value)) return value.slice(0, 10_000).map((item) => plainClone(item, depth + 1))
  if (typeof value === 'object') {
    const result = {}
    for (const key of Object.keys(value).slice(0, 1_000)) {
      const item = value[key]
      if (typeof item !== 'function' && typeof item !== 'symbol' && item !== undefined) {
        result[key] = plainClone(item, depth + 1)
      }
    }
    return result
  }
  return String(value)
}

function hydratePlainClone(value, depth = 0) {
  if (value == null || typeof value !== 'object' || depth >= 10) return value
  if (Array.isArray(value)) return value.map((item) => hydratePlainClone(item, depth + 1))
  if (value.type === 'date' && typeof value.value === 'string') {
    const date = new Date(value.value)
    return Number.isNaN(date.getTime()) ? value.value : date
  }
  if (value.type === 'binary' && typeof value.value === 'string') {
    try {
      return Buffer.from(value.value, 'base64')
    } catch {
      return Buffer.alloc(0)
    }
  }
  const result = {}
  for (const [key, item] of Object.entries(value)) result[key] = hydratePlainClone(item, depth + 1)
  return result
}

function cloneAdvancedRecord(value, limit = MAX_ADVANCED_METADATA_ENTRIES) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return { value: {}, truncated: false }
  const entries = Object.entries(value)
  return {
    value: Object.fromEntries(entries.slice(0, limit).map(([key, item]) => [key, plainClone(item)])),
    truncated: entries.length > limit,
  }
}

function normalizeColor(color) {
  if (!color || typeof color !== 'object') return undefined
  const result = {}
  const rawArgb = color.argb || color.rgb
  if (typeof rawArgb === 'string' && /^[0-9a-f]{6,8}$/i.test(rawArgb)) {
    result.argb = rawArgb.length === 6 ? `FF${rawArgb.toUpperCase()}` : rawArgb.toUpperCase()
  }
  if (Number.isInteger(color.indexed)) result.indexed = color.indexed
  if (Number.isInteger(color.theme)) result.theme = color.theme
  if (Number.isFinite(color.tint)) result.tint = color.tint
  if (color.auto === true) result.auto = true
  return Object.keys(result).length ? result : undefined
}

function normalizeFont(font) {
  if (!font || typeof font !== 'object') return undefined
  const result = {}
  const directKeys = [
    'name',
    'family',
    'scheme',
    'charset',
    'size',
    'bold',
    'italic',
    'underline',
    'strike',
    'outline',
    'shadow',
    'vertAlign',
    'condense',
    'extend',
  ]
  for (const key of directKeys) {
    let value = font[key]
    if (key === 'size' && value == null) value = font.sz
    if (value !== undefined && value !== null && value !== false) result[key] = plainClone(value)
  }
  const color = normalizeColor(font.color)
  if (color) result.color = color
  return Object.keys(result).length ? result : undefined
}

function normalizeFill(fill) {
  if (!fill || typeof fill !== 'object') return undefined
  const pattern = fill.pattern || fill.patternType
  const type = fill.type || (pattern ? 'pattern' : undefined)
  if (!type && !fill.fgColor && !fill.bgColor) return undefined
  if ((pattern === 'none' || pattern === 'gray125') && !fill.fgColor && !fill.bgColor) return undefined

  const result = {}
  if (type) result.type = type
  if (pattern) result.pattern = pattern
  if (typeof fill.gradient === 'string' && fill.gradient) result.gradient = fill.gradient
  const fgColor = normalizeColor(fill.fgColor)
  const bgColor = normalizeColor(fill.bgColor)
  if (fgColor) result.fgColor = fgColor
  if (bgColor) result.bgColor = bgColor
  if (Number.isFinite(fill.degree)) result.degree = fill.degree
  if (fill.center && typeof fill.center === 'object') result.center = plainClone(fill.center)
  if (Array.isArray(fill.stops)) result.stops = plainClone(fill.stops)
  // A bare <patternFill/> — what Google Sheets writes for the unused fills[0] — and a
  // pattern whose colour elements carry no usable value both survive the guards above but
  // paint nothing.  Keeping them would mark ordinary cells as filled, which costs them
  // their gridlines and inflates the used range.
  const paintsNothing = !result.fgColor && !result.bgColor && !(Array.isArray(result.stops) && result.stops.length)
  if (paintsNothing && (!result.pattern || result.pattern === 'none' || result.pattern === 'gray125')) return undefined
  return Object.keys(result).length ? result : undefined
}

function normalizeBorderSide(side) {
  if (!side || typeof side !== 'object') return undefined
  const result = {}
  if (typeof side.style === 'string' && side.style) result.style = side.style
  const color = normalizeColor(side.color)
  if (color) result.color = color
  return Object.keys(result).length ? result : undefined
}

function normalizeBorder(border) {
  if (!border || typeof border !== 'object') return undefined
  const result = {}
  for (const key of ['top', 'left', 'bottom', 'right', 'diagonal', 'vertical', 'horizontal']) {
    const side = normalizeBorderSide(border[key])
    if (side) result[key] = side
  }
  if (border.diagonalUp === true) result.diagonalUp = true
  if (border.diagonalDown === true) result.diagonalDown = true
  if (border.outline === false) result.outline = false
  return Object.keys(result).length ? result : undefined
}

function normalizeAlignment(alignment) {
  if (!alignment || typeof alignment !== 'object') return undefined
  const result = {}
  const keys = [
    'horizontal',
    'vertical',
    'wrapText',
    'shrinkToFit',
    'indent',
    'textRotation',
    'readingOrder',
    'justifyLastLine',
    'relativeIndent',
  ]
  for (const key of keys) {
    const value = alignment[key]
    if (value !== undefined && value !== null && value !== false && value !== 0) {
      result[key] = plainClone(value)
    }
  }
  return Object.keys(result).length ? result : undefined
}

// ExcelJS hands back the raw OOXML spellings for the locale-dependent builtin number
// formats.  Excel and Sheets render builtin 14 as the system short date and builtin 22 as
// short date + time, so the literal spellings would show "03-15-23" and a stray "h" here.
const BUILTIN_NUMBER_FORMAT_ALIASES = Object.freeze({
  'mm-dd-yy': 'm/d/yyyy',
  'm/d/yy "h":mm': 'm/d/yyyy h:mm',
})

function normalizeNumberFormat(numFmt) {
  if (typeof numFmt !== 'string') return undefined
  const format = numFmt.trim()
  if (!format || format === 'General') return undefined
  return BUILTIN_NUMBER_FORMAT_ALIASES[format] || format
}

function normalizeStyle(style) {
  if (!style || typeof style !== 'object') return undefined
  const result = {}
  const font = normalizeFont(style.font)
  const fill = normalizeFill(style.fill || (style.patternType ? style : undefined))
  const border = normalizeBorder(style.border)
  const alignment = normalizeAlignment(style.alignment)
  if (font) result.font = font
  if (fill) result.fill = fill
  if (border) result.border = border
  if (alignment) result.alignment = alignment
  if (style.protection && typeof style.protection === 'object') {
    const protection = {}
    if (style.protection.locked !== undefined) protection.locked = Boolean(style.protection.locked)
    if (style.protection.hidden !== undefined) protection.hidden = Boolean(style.protection.hidden)
    if (Object.keys(protection).length) result.protection = protection
  }
  if (Number.isInteger(style.quotePrefix)) result.quotePrefix = style.quotePrefix
  return Object.keys(result).length ? result : undefined
}

function normalizeHyperlink(value) {
  if (!value) return undefined
  if (typeof value === 'string') return { target: value }
  if (typeof value !== 'object') return undefined
  const target = value.target || value.hyperlink || value.Target
  if (typeof target !== 'string' || !target) return undefined
  const result = { target }
  const tooltip = value.tooltip || value.Tooltip
  if (typeof tooltip === 'string' && tooltip) result.tooltip = tooltip
  return result
}

function normalizeNote(note) {
  if (note == null) return undefined
  if (typeof note === 'string') return note
  return plainClone(note)
}

function dateToExcelSerial(value, date1904 = false) {
  const date = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(date.getTime())) return null
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
  return (date.getTime() - epoch) / 86_400_000
}

function serializeCellValue(value, date1904 = false) {
  if (value == null) return value
  if (value instanceof Date) return dateToExcelSerial(value, date1904)
  if (Buffer.isBuffer(value) || value instanceof Uint8Array) {
    return Buffer.from(value).toString('base64')
  }
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value === 'string' || typeof value === 'boolean') return value
  if (typeof value === 'bigint') return value.toString()
  if (value && Array.isArray(value.richText)) {
    return {
      type: 'richText',
      runs: value.richText.map((run) => {
        const result = { text: String(run && run.text != null ? run.text : '') }
        const font = normalizeFont(run && run.font)
        if (font) result.font = font
        return result
      }),
    }
  }
  if (value && value.error != null) return String(value.error)
  return String(value)
}

function deserializeCellValue(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return value
  if (value.type === 'date') {
    const date = new Date(value.value)
    return Number.isNaN(date.getTime()) ? String(value.value || '') : date
  }
  if (value.type === 'binary') {
    try {
      return Buffer.from(String(value.value || ''), 'base64')
    } catch {
      return Buffer.alloc(0)
    }
  }
  if (value.type === 'error') return { error: String(value.value || '#VALUE!') }
  if (value.type === 'richText') {
    return {
      richText: (Array.isArray(value.runs) ? value.runs : []).map((run) => {
        const result = { text: String(run && run.text != null ? run.text : '') }
        const font = normalizeFont(run && run.font)
        if (font) result.font = font
        return result
      }),
    }
  }
  return plainClone(value)
}

function richTextToPlainText(value) {
  if (value && value.type === 'richText' && Array.isArray(value.runs)) {
    return value.runs.map((run) => String(run && run.text != null ? run.text : '')).join('')
  }
  if (value && value.type === 'date') return String(value.value || '')
  if (value && value.type === 'error') return String(value.value || '#VALUE!')
  if (value && value.type === 'binary') return '[binary]'
  return value
}

function normalizeFormula(formula) {
  if (formula == null) return undefined
  const text = String(formula).trim()
  if (!text) return undefined
  return text.startsWith('=') ? text.slice(1) : text
}

/**
 * Translate a shared (filled-down) formula from its master cell to another cell of the group.
 * ExcelJS's own `slideFormula` shifts every letters-plus-digits token, including text inside
 * string literals ("Q1"), quoted sheet names and structured references (Sales[Q1]). Here the
 * formula is split into literal segments, which are copied unchanged, and code segments, which
 * are slid; a token directly followed by `[` is a table name and is never shifted.
 */
function translateSharedFormula(masterFormula, masterAddress, address) {
  const { slideFormula } = require('exceljs/lib/utils/shared-formula')
  const text = String(masterFormula || '')
  let output = ''
  let code = ''
  const flushCode = (beforeBracket) => {
    if (!code) return
    if (beforeBracket) {
      // Keep the table name in front of a structured reference as it is.
      const match = /([A-Za-z_\\][A-Za-z0-9_.]*)$/.exec(code)
      const head = match ? code.slice(0, match.index) : code
      output += slideFormula(head, masterAddress, address) + (match ? match[1] : '')
    } else {
      output += slideFormula(code, masterAddress, address)
    }
    code = ''
  }
  let position = 0
  while (position < text.length) {
    const char = text[position]
    if (char === '"' || char === "'") {
      let end = position + 1
      while (end < text.length) {
        if (text[end] === char) {
          if (text[end + 1] === char) { end += 2; continue }
          break
        }
        end += 1
      }
      flushCode(false)
      output += text.slice(position, end + 1)
      position = end + 1
      continue
    }
    if (char === '[') {
      let depth = 0
      let end = position
      while (end < text.length) {
        if (text[end] === "'" && text[end + 1] && (text[end + 1] === '[' || text[end + 1] === ']' || text[end + 1] === '#' || text[end + 1] === "'")) { end += 2; continue }
        if (text[end] === '[') depth += 1
        else if (text[end] === ']') {
          depth -= 1
          if (depth === 0) break
        }
        end += 1
      }
      flushCode(true)
      output += text.slice(position, end + 1)
      position = end + 1
      continue
    }
    code += char
    position += 1
  }
  flushCode(false)
  return output
}

function normalizeAddress(address) {
  const match = CELL_ADDRESS_RE.exec(String(address || '').trim())
  if (!match) return undefined
  return `${match[1].toUpperCase()}${Number(match[2])}`
}

function addressPosition(address) {
  const normalized = normalizeAddress(address)
  if (!normalized) return undefined
  try {
    const decoded = XLSX.utils.decode_cell(normalized)
    return { address: normalized, row: decoded.r + 1, col: decoded.c + 1 }
  } catch {
    return undefined
  }
}

function normalizeRange(range) {
  if (typeof range !== 'string') return undefined
  const text = range.replace(/\s+/g, '').toUpperCase()
  if (CELL_ADDRESS_RE.test(text)) return `${normalizeAddress(text)}:${normalizeAddress(text)}`
  const match = RANGE_RE.exec(text)
  if (!match) return undefined
  const start = addressPosition(`${match[1]}${match[2]}`)
  const end = addressPosition(`${match[3]}${match[4]}`)
  if (!start || !end) return undefined
  const top = Math.min(start.row, end.row)
  const bottom = Math.max(start.row, end.row)
  const left = Math.min(start.col, end.col)
  const right = Math.max(start.col, end.col)
  return XLSX.utils.encode_range({ s: { r: top - 1, c: left - 1 }, e: { r: bottom - 1, c: right - 1 } })
}

function autoFilterForExcelJS(value) {
  if (typeof value === 'string') return normalizeRange(value) || value
  if (!value || typeof value !== 'object') return undefined
  if (typeof value.ref === 'string') return normalizeRange(value.ref) || value.ref
  if (typeof value.Ref === 'string') return normalizeRange(value.Ref) || value.Ref
  if (value.from && value.to) return hydratePlainClone({ from: value.from, to: value.to })
  return undefined
}

function autoFilterForSheetJS(value) {
  if (typeof value === 'string') return { ref: normalizeRange(value) || value }
  if (!value || typeof value !== 'object') return undefined
  if (typeof value.ref === 'string') return { ...hydratePlainClone(value), ref: normalizeRange(value.ref) || value.ref }
  if (typeof value.Ref === 'string') return { ref: normalizeRange(value.Ref) || value.Ref }
  const address = (endpoint) => {
    if (typeof endpoint === 'string') return normalizeAddress(endpoint)
    const row = endpoint && Number(endpoint.row)
    const column = endpoint && Number(endpoint.column)
    if (!Number.isInteger(row) || !Number.isInteger(column) || row < 1 || column < 1) return undefined
    return XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })
  }
  const from = address(value.from)
  const to = address(value.to)
  return from && to ? { ref: normalizeRange(`${from}:${to}`) } : undefined
}

function initialStats(sourceBytes) {
  return {
    sourceBytes,
    sheets: 0,
    cells: 0,
    formulas: 0,
    sheetCount: 0,
    cellCount: 0,
    formulaCount: 0,
    styledCellCount: 0,
    mergeCount: 0,
    commentCount: 0,
    hyperlinkCount: 0,
    maxRow: 0,
    maxCol: 0,
  }
}

function updateStatsForCell(stats, cell, row, col) {
  stats.cellCount += 1
  stats.cells += 1
  if (cell.formula) stats.formulaCount += 1
  if (cell.formula) stats.formulas += 1
  if (cell.style || cell.numFmt) stats.styledCellCount += 1
  if (cell.note) stats.commentCount += 1
  if (cell.hyperlink) stats.hyperlinkCount += 1
  stats.maxRow = Math.max(stats.maxRow, row)
  stats.maxCol = Math.max(stats.maxCol, col)
}

function workbookPropertiesFromExcelJS(workbook) {
  const result = {}
  const stringKeys = [
    'creator',
    'lastModifiedBy',
    'title',
    'subject',
    'description',
    'keywords',
    'category',
    'company',
    'manager',
    'language',
    'contentStatus',
  ]
  for (const key of stringKeys) {
    if (typeof workbook[key] === 'string' && workbook[key]) result[key] = workbook[key]
  }
  for (const key of ['created', 'modified', 'lastPrinted']) {
    if (workbook[key] instanceof Date && !Number.isNaN(workbook[key].getTime())) {
      result[key] = workbook[key].toISOString()
    }
  }
  if (workbook.revision !== undefined && workbook.revision !== null) result.revision = plainClone(workbook.revision)
  return result
}

function workbookPropertiesFromSheetJS(workbook) {
  const props = workbook.Props || {}
  const custom = workbook.Custprops || {}
  const result = {}
  const aliases = {
    Author: 'creator',
    LastAuthor: 'lastModifiedBy',
    Title: 'title',
    Subject: 'subject',
    Comments: 'description',
    Keywords: 'keywords',
    Category: 'category',
    Company: 'company',
    Manager: 'manager',
    CreatedDate: 'created',
    ModifiedDate: 'modified',
    LastPrinted: 'lastPrinted',
    Language: 'language',
    ContentStatus: 'contentStatus',
    Revision: 'revision',
  }
  for (const [source, target] of Object.entries(aliases)) {
    const value = props[source]
    if (value instanceof Date && !Number.isNaN(value.getTime())) result[target] = value.toISOString()
    else if (value !== undefined && value !== null && String(value)) result[target] = String(value)
  }
  if (Object.keys(custom).length) result.custom = plainClone(custom)
  return result
}

function customPropertiesFromOoxml(buffer) {
  try {
    const propertiesOnly = XLSX.read(buffer, { type: 'buffer', bookProps: true })
    return propertiesOnly && propertiesOnly.Custprops && typeof propertiesOnly.Custprops === 'object'
      ? plainClone(propertiesOnly.Custprops)
      : undefined
  } catch {
    return undefined
  }
}

function decodeXml(value) {
  return String(value || '')
    .replace(/&#x([0-9a-f]+);/gi, (_match, code) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/&#(\d+);/g, (_match, code) => String.fromCodePoint(Number(code)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

function encodeXmlText(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function encodeXmlAttribute(value) {
  return encodeXmlText(value).replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

function parseXmlAttributes(source) {
  const result = {}
  const attributePattern = /([A-Za-z_:][A-Za-z0-9_.:-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
  let match
  while ((match = attributePattern.exec(String(source || '')))) result[match[1]] = decodeXml(match[2] ?? match[3] ?? '')
  return result
}

function xmlScalar(value) {
  if (value === 'true' || value === '1') return true
  if (value === 'false' || value === '0') return false
  if (/^-?\d+(?:\.\d+)?$/.test(String(value))) {
    const number = Number(value)
    if (Number.isFinite(number)) return number
  }
  return value
}

// Style records address theme colours by index, so the palette has to come from the
// workbook's own clrScheme.  The index order is the OOXML one — lt1 and dk1 are swapped
// relative to the document order in the XML.
const THEME_COLOR_ORDER = Object.freeze(['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'])
const SYSTEM_COLOR_FALLBACKS = Object.freeze({ window: 'FFFFFF', windowtext: '000000' })

function themeColorsFromXml(xml) {
  if (typeof xml !== 'string' || !xml) return undefined
  const scheme = /<a:clrScheme[ >][^]*?<[/]a:clrScheme[^>]*>/i.exec(xml)
  if (!scheme) return undefined
  const colors = THEME_COLOR_ORDER.map((slot) => {
    const entry = new RegExp('<a:' + slot + '(?=[ />])[^>]*>([^]*?)<[/]a:' + slot + '[^>]*>', 'i').exec(scheme[0])
    if (!entry) return ''
    const srgb = /<a:srgbClr[^>]* val="([0-9A-Fa-f]{6})"/.exec(entry[1])
    if (srgb) return srgb[1].toUpperCase()
    const system = /<a:sysClr[^>]* lastClr="([0-9A-Fa-f]{6})"/.exec(entry[1])
    if (system) return system[1].toUpperCase()
    const named = /<a:sysClr[^>]* val="([A-Za-z]+)"/.exec(entry[1])
    return named ? SYSTEM_COLOR_FALLBACKS[named[1].toLowerCase()] || '' : ''
  })
  return colors.some((hex) => hex) ? colors : undefined
}

async function workbookXmlFromOoxml(buffer) {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file('xl/workbook.xml')
  if (!entry) return undefined
  const expectedSize = entry._data && Number(entry._data.uncompressedSize)
  if (Number.isFinite(expectedSize) && expectedSize > MAX_WORKBOOK_XML_BYTES) {
    throw new RangeError('The OOXML workbook metadata is unexpectedly large.')
  }
  const data = await entry.async('nodebuffer')
  if (data.length > MAX_WORKBOOK_XML_BYTES) throw new RangeError('The OOXML workbook metadata is unexpectedly large.')
  const partNames = Object.keys(zip.files)
  const themePart = partNames.find((partName) => /^xl\/theme\/theme\d*\.xml$/i.test(partName))
  let themeColors
  if (themePart) {
    const themeEntry = zip.file(themePart)
    const themeSize = themeEntry && themeEntry._data && Number(themeEntry._data.uncompressedSize)
    if (themeEntry && (!Number.isFinite(themeSize) || themeSize <= MAX_WORKBOOK_XML_BYTES)) {
      themeColors = themeColorsFromXml((await themeEntry.async('nodebuffer')).toString('utf8'))
    }
  }
  let normalFont, fontSizes
  const stylesEntry = zip.file('xl/styles.xml')
  if (stylesEntry && Number(stylesEntry._data?.uncompressedSize || 0) <= MAX_WORKBOOK_XML_BYTES) {
    const styles = await stylesEntry.async('string')
    if (Buffer.byteLength(styles) <= MAX_WORKBOOK_XML_BYTES) {
      normalFont = normalFontFromStylesXml(styles)
      const fonts = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i.exec(styles)?.[1] || ''
      fontSizes = [...fonts.matchAll(/<font\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/font>)/gi)].map(match => {
        const attributes = /<sz\b([^>]*)>/i.exec(match[0])?.[1]
        const size = Number(attributes ? parseXmlAttributes(attributes).val : NaN)
        return Number.isFinite(size) && size > 0 && size <= 409 ? size : undefined
      })
    }
  }
  return { xml: data.toString('utf8'), partNames, themeColors, normalFont, fontSizes }
}

function normalFontFromStylesXml(xml) {
  // Column widths use the built-in Normal style's font, not the first cell font.
  const section = (name) => new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)<\\/${name}>`, 'i').exec(xml)?.[1] || ''
  const styles = [...section('cellStyles').matchAll(/<cellStyle\b([^>]*)\/?\s*>/gi)]
  const normal = styles.map(match => parseXmlAttributes(match[1])).find(style => style.builtinId === '0')
  if (!normal || !/^\d+$/.test(normal.xfId || '')) return undefined
  const xfs = [...section('cellStyleXfs').matchAll(/<xf\b([^>]*)>/gi)]
  const xf = xfs[Number(normal.xfId)]
  if (!xf) return undefined
  const fontIndex = Number(parseXmlAttributes(xf[1]).fontId)
  const fonts = [...section('fonts').matchAll(/<font\b[^>]*>([\s\S]*?)<\/font>/gi)]
  const font = fonts[fontIndex]?.[1]
  if (!font) return undefined
  const value = (tag) => { const attributes = new RegExp(`<${tag}\\b([^>]*)>`, 'i').exec(font)?.[1]; return attributes ? parseXmlAttributes(attributes).val : undefined }
  const name = value('name'), size = Number(value('sz'))
  if (!name || !Number.isFinite(size) || size <= 0 || size > 409) return undefined
  const flag = tag => new RegExp(`<${tag}(?:\\s|/|>)`, 'i').test(font) && !['0', 'false'].includes(value(tag))
  return { name, size, ...(flag('b') ? { bold: true } : {}), ...(flag('i') ? { italic: true } : {}) }
}

async function extractOoxmlWorkbookMetadata(buffer, warnings) {
  try {
    const packageMetadata = await workbookXmlFromOoxml(buffer)
    if (!packageMetadata) return {}
    const { xml, partNames, themeColors, normalFont, fontSizes } = packageMetadata
    const advancedPackageParts = []
    // Charts on worksheets are modelled and round-tripped by chart-xlsx.cjs (charts
    // it cannot draw are warned about there and kept byte-for-byte), so only
    // chart sheets remain an unmodelled part.
    const advancedPartMatchers = [
      [/^xl\/chartsheets\//i, 'chart sheets'],
      [/^xl\/pivot/i, 'pivot tables or caches'],
      [/^xl\/slicer/i, 'slicers'],
      [/^xl\/externalLinks\//i, 'external links'],
      [/^xl\/(connections|queryTables)\b/i, 'data connections or queries'],
      [/^xl\/model\//i, 'data models'],
      [/^xl\/(ctrlProps|activeX|embeddings)\//i, 'embedded controls or objects'],
      [/^xl\/threadedComments\//i, 'threaded comments'],
      [/^_xmlsignatures\//i, 'digital signatures'],
    ]
    for (const [pattern, label] of advancedPartMatchers) {
      if (partNames.some((partName) => pattern.test(partName))) advancedPackageParts.push(label)
    }
    if (advancedPackageParts.length) {
      warnings.push(
        `This workbook contains ${advancedPackageParts.join(', ')} that are not fully modeled; even source-backed saving may simplify or remove those parts.`,
      )
    }
    const calcMatch = /<calcPr\b([^>]*?)(?:\/?\s*>)/i.exec(xml)
    const calcProperties = {}
    if (calcMatch) {
      for (const [key, value] of Object.entries(parseXmlAttributes(calcMatch[1]))) calcProperties[key] = xmlScalar(value)
    }
    const workbookPropertiesMatch = /<workbookPr\b([^>]*?)(?:\/?\s*>)/i.exec(xml)
    const workbookProperties = {}
    if (workbookPropertiesMatch) {
      for (const [key, value] of Object.entries(parseXmlAttributes(workbookPropertiesMatch[1]))) {
        workbookProperties[key] = xmlScalar(value)
      }
    }

    const workbookViews = []
    const viewPattern = /<workbookView\b([^>]*?)(?:\/?\s*>)/gi
    let viewMatch
    while ((viewMatch = viewPattern.exec(xml)) && workbookViews.length < 100) {
      const view = {}
      for (const [key, value] of Object.entries(parseXmlAttributes(viewMatch[1]))) view[key] = xmlScalar(value)
      workbookViews.push(view)
    }

    const definedNames = []
    const namePattern = /<definedName\b([^>]*)>([\s\S]*?)<\/definedName\s*>/gi
    let nameMatch
    let namesTruncated = false
    while ((nameMatch = namePattern.exec(xml))) {
      if (definedNames.length >= MAX_ADVANCED_METADATA_ENTRIES) {
        namesTruncated = true
        break
      }
      const attributes = parseXmlAttributes(nameMatch[1])
      const name = attributes.name
      const ref = decodeXml(nameMatch[2].replace(/<[^>]*>/g, '')).trim()
      if (!name || !ref) continue
      const item = { name: String(name), ref, ranges: [ref] }
      if (attributes.localSheetId != null && Number.isInteger(Number(attributes.localSheetId))) {
        item.localSheetIndex = Number(attributes.localSheetId)
      }
      if (attributes.hidden != null) item.hidden = attributes.hidden === '1' || attributes.hidden === 'true'
      if (attributes.comment) item.comment = attributes.comment
      const auxiliary = { ...attributes }
      delete auxiliary.name
      delete auxiliary.localSheetId
      delete auxiliary.hidden
      delete auxiliary.comment
      if (Object.keys(auxiliary).length) item.attributes = plainClone(auxiliary)
      definedNames.push(item)
    }
    if (namesTruncated) warnings.push('Some defined names exceeded the metadata safety limit and require the original package for preservation.')
    const metadata = { calcProperties, workbookProperties, workbookViews, definedNames, advancedPackageParts }
    for (const tag of ['workbookProtection', 'fileSharing']) {
      const match = new RegExp(`<${tag}\\b([^>]*?)\\/?\\s*>`, 'i').exec(xml)
      if (match) {
        const attributes = parseXmlAttributes(match[1])
        if (Object.keys(attributes).length) metadata[tag] = attributes
      }
    }
    // Charts: only open drawing parts again when the package has any (chart-xlsx.cjs).
    if (partNames.some((partName) => /^xl\/drawings\/[^/]+\.xml$/i.test(partName))) metadata.hasDrawingParts = true
    // Pivot definitions: only scan custom XML parts when the package has any (pivot-xlsx.cjs).
    if (partNames.some((partName) => /^customXml\/item\d+\.xml$/i.test(partName))) metadata.hasCustomXmlParts = true
    if (themeColors) metadata.themeColors = themeColors
    if (normalFont) metadata.normalFont = normalFont
    if (fontSizes) metadata.fontSizes = fontSizes
    return metadata
  } catch (error) {
    warnings.push(`Some OOXML workbook-level metadata could not be inspected (${error.message || 'invalid metadata'}).`)
    return {}
  }
}

function extractFrozenView(views) {
  const view = Array.isArray(views) ? views.find((candidate) => candidate && candidate.state === 'frozen') : null
  if (!view) return null
  const frozen = {
    rows: Math.max(0, Number(view.ySplit) || 0),
    columns: Math.max(0, Number(view.xSplit) || 0),
  }
  if (typeof view.topLeftCell === 'string') frozen.topLeftCell = view.topLeftCell
  if (typeof view.activeCell === 'string') frozen.activeCell = view.activeCell
  return frozen
}

function extractExcelJSNote(note) {
  if (note == null) return undefined
  if (typeof note === 'string') return note
  return plainClone(note)
}

function excelJSCellToModel(cell, date1904 = false) {
  let formula
  let result
  const rawValue = cell.value
  const isSharedClone = Boolean(rawValue && typeof rawValue === 'object' && typeof rawValue.sharedFormula === 'string')
  if (cell.type === ExcelJS.ValueType.Formula || (rawValue && typeof rawValue === 'object' && ('formula' in rawValue || 'sharedFormula' in rawValue))) {
    // A shared-formula clone's own formula is the master's, translated to this cell. If the
    // master cannot be found there is no formula to keep (never the master's address). The
    // translation skips string literals and structured references (translateSharedFormula),
    // unlike ExcelJS's `cell.formula`, which would turn "Q1" into "Q2" in the filled cells.
    try {
      if (isSharedClone) {
        const worksheet = cell.worksheet
        const master = worksheet && typeof worksheet.findCell === 'function' ? worksheet.findCell(rawValue.sharedFormula) : null
        const masterValue = master && master.value
        const masterFormula = masterValue && typeof masterValue === 'object' && typeof masterValue.formula === 'string' ? masterValue.formula : null
        formula = masterFormula ? normalizeFormula(translateSharedFormula(masterFormula, master.address, cell.address)) : undefined
      } else {
        formula = normalizeFormula(cell.formula || rawValue.formula)
      }
    } catch {
      formula = normalizeFormula(rawValue.formula)
    }
    const rawResult = cell.result !== undefined ? cell.result : rawValue.result
    if (rawResult !== undefined) result = serializeCellValue(rawResult, date1904)
  }

  const modelCell = {}
  if (formula) {
    modelCell.formula = formula
    if (result !== undefined) modelCell.result = result
    if (rawValue && rawValue.result instanceof Date) modelCell.resultType = 'date'
    if (rawValue && rawValue.result && rawValue.result.error != null) modelCell.resultType = 'error'
    // Shared (filled-down) formulas are kept per cell: each cell owns its translated formula, so
    // editing the first cell of a fill can never rewrite or invalidate the others. Only array
    // formulas keep their group and range.
    if (rawValue && rawValue.shareType === 'array') {
      modelCell.formulaType = 'array'
      if (typeof rawValue.ref === 'string' && rawValue.ref) modelCell.formulaRange = rawValue.ref
    } else {
      // A plain <f> in a file is what Excel calculates with implicit intersection (a formula
      // entered in Excel 365 that needs arrays is saved as an array formula). The calculation
      // engine treats it the same way until the cell is edited.
      modelCell.implicitIntersection = true
    }
  } else if (isSharedClone) {
    const value = serializeCellValue(rawValue.result, date1904)
    modelCell.value = value && typeof value === 'object' ? richTextToPlainText(value) : value === undefined ? null : value
    if (rawValue.result instanceof Date) modelCell.type = 'date'
    else if (rawValue.result && rawValue.result.error != null) modelCell.type = 'error'
  } else if (rawValue && typeof rawValue === 'object' && rawValue.hyperlink) {
    modelCell.value = serializeCellValue(rawValue.text != null ? rawValue.text : cell.text, date1904)
  } else {
    const serialized = serializeCellValue(rawValue, date1904)
    if (serialized && typeof serialized === 'object' && serialized.type === 'richText') {
      // The model's value is a scalar by contract, so rich text lands as its plain text and
      // the run formatting travels beside it; the writer re-attaches the runs only while the
      // text still matches, so an edited cell keeps the user's words instead of stale runs.
      modelCell.value = richTextToPlainText(serialized)
      modelCell.richText = plainClone(serialized.runs)
      modelCell.type = 'richText'
    } else {
      modelCell.value = serialized
      if (rawValue instanceof Date) modelCell.type = 'date'
      else if (rawValue && typeof rawValue === 'object' && rawValue.error != null) modelCell.type = 'error'
    }
  }

  const cellText = cell.text == null ? '' : String(cell.text)
  const plainResult = richTextToPlainText(result)
  modelCell.display = cellText === '[object Object]' && plainResult != null ? String(plainResult) : cellText
  const style = normalizeStyle(cell.style)
  if (style) modelCell.style = style
  const numFmt = normalizeNumberFormat(cell.numFmt)
  if (numFmt) modelCell.numFmt = numFmt
  const hyperlink = normalizeHyperlink(rawValue && rawValue.hyperlink ? rawValue : cell.hyperlink)
  if (hyperlink) {
    modelCell.hyperlink = hyperlink.target
    if (hyperlink.tooltip) modelCell.hyperlinkTooltip = hyperlink.tooltip
  }
  const note = extractExcelJSNote(cell.note)
  if (note !== undefined) modelCell.note = note
  return modelCell
}

function excelJSRows(worksheet) {
  if (Array.isArray(worksheet._rows)) return worksheet._rows.filter(Boolean)
  const rows = []
  worksheet.eachRow({ includeEmpty: false }, (row) => rows.push(row))
  return rows
}

function excelJSCells(row) {
  if (Array.isArray(row._cells)) return row._cells.filter(Boolean)
  const cells = []
  row.eachCell({ includeEmpty: false }, (cell) => cells.push(cell))
  return cells
}

/** Every A1 address of a small range ("A1", "A1:B3"); null for ranges over the limit. */
function rangeAddresses(ref, limit = 10_000) {
  const range = normalizeRange(String(ref || ''))
  if (!range) return null
  const decoded = XLSX.utils.decode_range(range)
  const count = (decoded.e.r - decoded.s.r + 1) * (decoded.e.c - decoded.s.c + 1)
  if (count > limit) return null
  const addresses = []
  for (let row = decoded.s.r; row <= decoded.e.r; row += 1) {
    for (let col = decoded.s.c; col <= decoded.e.c; col += 1) addresses.push(XLSX.utils.encode_cell({ r: row, c: col }))
  }
  return addresses
}

/** Attach hyperlinks ExcelJS could not attach: in-workbook places, formula cells, range refs. */
function applyCapturedHyperlinks(cells, links, stats) {
  for (const link of Array.isArray(links) ? links : []) {
    const single = normalizeAddress(link.ref)
    const addresses = single ? [single] : rangeAddresses(link.ref) || []
    for (const address of addresses) {
      let cell = cells[address]
      if (!cell) {
        if (!single) continue
        cell = cells[address] = { value: null, display: '' }
      }
      if (!cell.hyperlink) stats.hyperlinkCount += 1
      cell.hyperlink = link.target
      if (link.tooltip) cell.hyperlinkTooltip = link.tooltip
    }
  }
}

function isCheckboxRule(rule) {
  if (!rule || rule.type !== 'list' || !Array.isArray(rule.formulae)) return false
  return /^"\s*TRUE\s*,\s*FALSE\s*"$/i.test(String(rule.formulae[0] || '').trim())
}

/**
 * Checkboxes are saved as TRUE/FALSE cells with a "TRUE,FALSE" list rule (the convention the
 * editor writes). Mark such cells as checkboxes on open so they are not shown as dropdowns.
 */
function markCheckboxCells(cells, validations) {
  const ranges = []
  for (const [key, rule] of Object.entries(validations || {})) {
    if (!isCheckboxRule(rule)) continue
    const range = normalizeRange(key)
    if (range) ranges.push(XLSX.utils.decode_range(range))
  }
  if (!ranges.length) return
  for (const [address, cell] of Object.entries(cells)) {
    if (!cell || typeof cell.value !== 'boolean' || cell.formula || cell.type) continue
    const position = XLSX.utils.decode_cell(address)
    if (ranges.some((range) => position.r >= range.s.r && position.r <= range.e.r && position.c >= range.s.c && position.c <= range.e.c)) cell.type = 'checkbox'
  }
}

function excelJSSheetToModel(worksheet, index, stats, warnings, date1904 = false, hyperlinks) {
  const cells = {}
  let actualMaxRow = 0
  let actualMaxCol = 0

  for (const row of excelJSRows(worksheet)) {
    for (const cell of excelJSCells(row)) {
      if (cell.type === ExcelJS.ValueType.Merge) continue
      const position = addressPosition(cell.address)
      if (!position) continue
      const hasStoredStyle = Boolean(normalizeStyle(cell.style)) || (typeof cell.numFmt === 'string' && cell.numFmt !== 'General')
      const hasContent = cell.value != null || hasStoredStyle || cell.note != null || cell.hyperlink
      if (!hasContent) continue
      const modelCell = excelJSCellToModel(cell, date1904)
      cells[position.address] = modelCell
      actualMaxRow = Math.max(actualMaxRow, position.row)
      actualMaxCol = Math.max(actualMaxCol, position.col)
      updateStatsForCell(stats, modelCell, position.row, position.col)
    }
  }

  const merges = []
  const rawMerges = worksheet.model && Array.isArray(worksheet.model.merges) ? worksheet.model.merges : []
  for (const rawMerge of rawMerges) {
    const merge = normalizeRange(rawMerge)
    if (merge && !merges.includes(merge)) merges.push(merge)
  }
  stats.mergeCount += merges.length

  const colWidths = {}
  const hiddenCols = []
  const columnProperties = {}
  let columnMetadataMax = 0
  const columns = Array.isArray(worksheet._columns) ? worksheet._columns : worksheet.columns || []
  columns.forEach((column, columnIndex) => {
    if (column && Number.isFinite(column.width) && column.width > 0) {
      colWidths[String(columnIndex + 1)] = column.width
      columnMetadataMax = Math.max(columnMetadataMax, columnIndex + 1)
    }
    if (column && column.hidden === true) {
      hiddenCols.push(columnIndex + 1)
      columnMetadataMax = Math.max(columnMetadataMax, columnIndex + 1)
    }
    if (column) {
      const properties = {}
      if (Number.isFinite(column.outlineLevel) && column.outlineLevel > 0) properties.outlineLevel = column.outlineLevel
      if (column.collapsed === true) properties.collapsed = true
      const style = normalizeStyle(column.style)
      if (style) properties.style = style
      const columnNumFmt = column.style && normalizeNumberFormat(column.style.numFmt)
      if (columnNumFmt) properties.numFmt = columnNumFmt
      if (Object.keys(properties).length) columnProperties[String(columnIndex + 1)] = properties
    }
  })

  const rowHeights = {}
  const hiddenRows = []
  const rowProperties = {}
  let rowMetadataMax = 0
  for (const row of excelJSRows(worksheet)) {
    if (Number.isFinite(row.height) && row.height > 0) {
      rowHeights[String(row.number)] = row.height
      rowMetadataMax = Math.max(rowMetadataMax, row.number)
    }
    if (row.hidden === true) {
      hiddenRows.push(row.number)
      rowMetadataMax = Math.max(rowMetadataMax, row.number)
    }
    const properties = {}
    if (Number.isFinite(row.outlineLevel) && row.outlineLevel > 0) properties.outlineLevel = row.outlineLevel
    if (row.collapsed === true) properties.collapsed = true
    const style = normalizeStyle(row.style)
    if (style) properties.style = style
    const rowNumFmt = row.style && normalizeNumberFormat(row.style.numFmt)
    if (rowNumFmt) properties.numFmt = rowNumFmt
    if (Object.keys(properties).length) rowProperties[String(row.number)] = properties
  }

  const dimensions = worksheet.dimensions
  const dimensionBottom = dimensions && Number.isFinite(dimensions.bottom) ? dimensions.bottom : 0
  const dimensionRight = dimensions && Number.isFinite(dimensions.right) ? dimensions.right : 0
  const rawRowCount = Math.max(actualMaxRow, dimensionBottom, rowMetadataMax)
  const rawColCount = Math.max(actualMaxCol, dimensionRight, columnMetadataMax)
  if (rawRowCount > MAX_METADATA_ROWS || rawColCount > MAX_METADATA_COLS) {
    warnings.push(`Sheet "${worksheet.name}" has dimensions beyond modern Excel limits; sparse cells were retained.`)
  }

  const model = {
    id: `sheet-${index + 1}`,
    name: worksheet.name || `Sheet${index + 1}`,
    sourceWorksheetId: Number.isInteger(worksheet.id) ? worksheet.id : undefined,
    sourceSheetName: worksheet.name || `Sheet${index + 1}`,
    sourceSheetIndex: index,
    state: worksheet.state === 'hidden' || worksheet.state === 'veryHidden' ? worksheet.state : 'visible',
    rowCount: Math.max(1, Math.min(rawRowCount || 1, MAX_METADATA_ROWS)),
    colCount: Math.max(1, Math.min(rawColCount || 1, MAX_METADATA_COLS)),
    cells,
    merges,
    colWidths,
    rowHeights,
    hiddenCols,
    hiddenRows,
    columnProperties,
    rowProperties,
    frozen: extractFrozenView(worksheet.views),
    views: plainClone(worksheet.views || []),
    properties: plainClone(worksheet.properties || {}),
    pageSetup: { showGridLines: false, showRowColHeaders: false, ...plainClone(worksheet.pageSetup || {}) },
    headerFooter: plainClone(worksheet.headerFooter || {}),
    rowBreaks: plainClone(worksheet.rowBreaks || []),
  }
  if (worksheet.autoFilter) model.autoFilter = plainClone(worksheet.autoFilter)

  if (worksheet.sheetProtection) model.sheetProtection = plainClone(worksheet.sheetProtection)
  // Validation rules are keyed by their sqref ranges (validation-xlsx.cjs), so a dropdown on a
  // whole column is one entry, not a million; only a sheet with an extreme number of separate
  // rule areas is truncated.
  const validations = cloneAdvancedRecord(worksheet.dataValidations && worksheet.dataValidations.model, MAX_VALIDATION_AREAS)
  if (Object.keys(validations.value).length) model.dataValidations = validations.value
  if (validations.truncated) {
    model.dataValidationsTruncated = true
    warnings.push(`Sheet "${worksheet.name}" has more than ${MAX_VALIDATION_AREAS.toLocaleString('en-US')} separate data-validation ranges; its original rules are kept when saving, so validation changes on that sheet are not saved.`)
  }
  applyCapturedHyperlinks(cells, hyperlinks, stats)
  markCheckboxCells(cells, model.dataValidations)
  const conditionalFormattings = Array.isArray(worksheet.conditionalFormattings) ? worksheet.conditionalFormattings : []
  if (conditionalFormattings.length) {
    model.conditionalFormattings = conditionalFormattings.slice(0, MAX_ADVANCED_METADATA_ENTRIES).map((item) => plainClone(item))
  }
  if (conditionalFormattings.length > MAX_ADVANCED_METADATA_ENTRIES) {
    model.conditionalFormattingsTruncated = true
    warnings.push(`Sheet "${worksheet.name}" has extensive conditional formatting; the original OOXML package is required to retain every rule.`)
  }

  // Data validation and conditional formatting are modelled, edited, and written back, so
  // they no longer need a compatibility note; only features the editor cannot author do.
  const advanced = []
  // Tables are modelled (structured references, styles, totals) and written back.
  model.tables = tablesFromWorksheet(worksheet)
  // Pictures are modelled (shown, moved, resized, inserted, deleted) unless the workbook's
  // pictures are too large to hold in the editor; then the package keeps them untouched.
  const images = imagesFromWorksheet(worksheet, model)
  if (images && images.length) model.images = images
  else if (images === null) advanced.push('images')
  if (advanced.length) {
    model.requiresSourcePackage = advanced.slice()
    warnings.push(
      `Sheet "${worksheet.name}" contains ${advanced.join(', ')}. Source-backed OOXML saving retains supported structures, but conversion or a fresh rebuild may simplify them.`,
    )
  }
  return model
}

function definedNamesFromExcelJS(workbook) {
  try {
    return (workbook.definedNames.model || []).map((item) => ({
      name: String(item.name),
      ranges: Array.isArray(item.ranges) ? item.ranges.map(String) : [],
    }))
  } catch {
    return []
  }
}

function attachDefinedNameSheetIds(definedNames, sheets) {
  for (const item of Array.isArray(definedNames) ? definedNames : []) {
    if (!item || !Number.isInteger(item.localSheetIndex)) continue
    const sheet = sheets[item.localSheetIndex]
    if (sheet && typeof sheet.id === 'string') item.localSheetRefId = sheet.id
  }
  return definedNames
}

function activeSheetIndexFromExcelJS(workbook) {
  const view = Array.isArray(workbook.views) ? workbook.views[0] : undefined
  const index = view && Number.isInteger(view.activeTab) ? view.activeTab : 0
  return Math.max(0, Math.min(index, Math.max(0, workbook.worksheets.length - 1)))
}

/**
 * Hyperlinks of every worksheet model as ExcelJS parsed them, before its reconcile drops the
 * ones it cannot attach to a cell: links to a place in the workbook (location only, no
 * relationship), links on formula cells and links whose ref is a range.
 */
function captureWorksheetHyperlinks(model) {
  const captured = new Map()
  for (const worksheet of Array.isArray(model.worksheets) ? model.worksheets : []) {
    const links = Array.isArray(worksheet && worksheet.hyperlinks) ? worksheet.hyperlinks : []
    if (!links.length) continue
    const rels = (Array.isArray(model.worksheetRels) && model.worksheetRels[worksheet.sheetNo]) || []
    const byId = new Map(rels.map((rel) => [rel.Id, rel]))
    const list = []
    for (const link of links.slice(0, MAX_ADVANCED_METADATA_ENTRIES)) {
      if (!link || typeof link.address !== 'string') continue
      const rel = link.rId ? byId.get(link.rId) : null
      const external = rel && typeof rel.Target === 'string' ? rel.Target : ''
      // ExcelJS keeps the `location` attribute as `target`.
      const location = typeof link.target === 'string' ? link.target : ''
      let target = ''
      if (external) target = location ? `${external}#${location}` : external
      else if (location) target = `#${location}`
      if (!target) continue
      list.push({ ref: link.address, target, ...(link.tooltip ? { tooltip: String(link.tooltip) } : {}) })
    }
    if (list.length) captured.set(worksheet, list)
  }
  return captured
}

async function loadExcelJSWithExactFonts(workbook, buffer, metadata, hyperlinkSink) {
  const sizes = metadata?.fontSizes
  const loader = workbook.xlsx
  const fixFonts = sizes?.some(size => Number.isFinite(size) && !Number.isInteger(size))
  if (!fixFonts && !hyperlinkSink) return loader.load(buffer)
  // ExcelJS4.4 parses <sz> with IntegerXform. Restore the bounded OOXML font
  // table by ID before reconciliation shares those font objects with cells,
  // rows and columns. This adapter affects only this workbook/load, never the
  // dependency's global parser or another concurrent import.
  const reconcile = loader.reconcile
  loader.reconcile = function (model, ...args) {
    if (fixFonts) {
      const fonts = model.styles?.model?.fonts
      if (!Array.isArray(fonts) || fonts.length !== sizes.length) throw new Error('The workbook font table could not be matched safely.')
      sizes.forEach((size, index) => { if (Number.isFinite(size) && fonts[index]) fonts[index].size = size })
    }
    const captured = hyperlinkSink ? captureWorksheetHyperlinks(model) : null
    const result = reconcile.call(this, model, ...args)
    // Worksheet ids are assigned during reconcile.
    if (captured) for (const [worksheet, links] of captured) if (Number.isInteger(worksheet.id)) hyperlinkSink.set(worksheet.id, links)
    return result
  }
  try { return await loader.load(buffer) }
  finally { loader.reconcile = reconcile }
}

async function importWithExcelJS(buffer, sourceName, warnings) {
  const ooxmlMetadata = await extractOoxmlWorkbookMetadata(buffer, warnings)
  const excelWorkbook = new ExcelJS.Workbook()
  const hyperlinks = new Map()
  await loadExcelJSWithExactFonts(excelWorkbook, buffer, ooxmlMetadata, hyperlinks)
  const stats = initialStats(buffer.length)
  const sourceDate1904 = Boolean(excelWorkbook.properties && excelWorkbook.properties.date1904)
  // The renderer and output model use one canonical 1900-based serial system.
  // ExcelJS has already materialized formatted dates as real Date objects, so
  // converting those objects to the canonical epoch avoids four-year shifts.
  const sheets = excelWorkbook.worksheets.map((worksheet, index) => excelJSSheetToModel(worksheet, index, stats, warnings, false, hyperlinks.get(worksheet.id)))
  // Never present a non-empty package as an empty workbook: saving it would destroy the file.
  if (!sheets.length) throw new Error(`No sheets could be read from ${sourceName}.`)
  // Charts: ExcelJS ignores DrawingML charts; read them into sheet.charts.
  if (ooxmlMetadata.hasDrawingParts) {
    try {
      await importChartsIntoSheets(buffer, sheets, warnings, ooxmlMetadata.themeColors)
    } catch (error) {
      warnings.push(`Charts in this workbook could not be read (${error.message || 'invalid drawing parts'}); they will not be kept when saving.`)
    }
  }
  // Excel sparkline groups live in the worksheet extLst, which ExcelJS does not read.
  try {
    await importSparklineGroups(buffer, sheets, ooxmlMetadata.themeColors)
  } catch {
    // Sparklines are optional; the cells still open.
  }
  if (ooxmlMetadata.hasCustomXmlParts) {
    try {
      await importPivotDefinitions(buffer, sheets, warnings)
    } catch {
      // Pivot definitions are optional; their last results stay as plain cells.
    }
  }
  stats.sheetCount = sheets.length
  stats.sheets = sheets.length
  const activeIndex = activeSheetIndexFromExcelJS(excelWorkbook)
  const properties = workbookPropertiesFromExcelJS(excelWorkbook)
  const customProperties = customPropertiesFromOoxml(buffer)
  if (customProperties && Object.keys(customProperties).length) properties.custom = customProperties
  const definedNames = attachDefinedNameSheetIds(Array.isArray(ooxmlMetadata.definedNames) && ooxmlMetadata.definedNames.length
    ? ooxmlMetadata.definedNames
    : definedNamesFromExcelJS(excelWorkbook), sheets)
  // ExcelJS reads only the first range of a multi-area Print_Area name and turns whole-column
  // or whole-row areas ($A:$F, $1:$20) into "ANaN:FNaN". The defined name is authoritative.
  applyPrintAreasFromDefinedNames(sheets, definedNames, warnings)
  const workbookViews = Array.isArray(ooxmlMetadata.workbookViews) && ooxmlMetadata.workbookViews.length
    ? ooxmlMetadata.workbookViews
    : plainClone(excelWorkbook.views || [])
  return {
    model: {
      version: 1,
      name: path.parse(sourceName).name || 'Workbook',
      activeSheetId: sheets[activeIndex] ? sheets[activeIndex].id : sheets[0].id,
      sheets,
      definedNames,
      metadata: {
        importedWith: 'exceljs',
        ...properties,
        properties,
        date1904: false,
        sourceDate1904,
        workbookProperties: plainClone(ooxmlMetadata.workbookProperties || {}),
        calcProperties: plainClone(ooxmlMetadata.calcProperties || excelWorkbook.calcProperties || {}),
        workbookViews,
        advancedPackageParts: plainClone(ooxmlMetadata.advancedPackageParts || []),
        // The workbook's own clrScheme, in OOXML theme-index order, so styles that address a
        // colour by theme index resolve against this file's palette rather than a default one.
        ...(ooxmlMetadata.themeColors ? { themeColors: plainClone(ooxmlMetadata.themeColors) } : {}),
        ...(ooxmlMetadata.normalFont ? { normalFont: plainClone(ooxmlMetadata.normalFont) } : {}),
        // "Protect Workbook" (structure lock) and the file-sharing write password are not
        // modelled by ExcelJS; they are carried as opaque attributes and written back.
        ...(ooxmlMetadata.workbookProtection ? { workbookProtection: plainClone(ooxmlMetadata.workbookProtection) } : {}),
        ...(ooxmlMetadata.fileSharing ? { fileSharing: plainClone(ooxmlMetadata.fileSharing) } : {}),
        definedNames: definedNames.map((item) => ({
          name: item.name,
          ranges: (item.ranges || []).join(','),
          ...(Number.isInteger(item.localSheetIndex) ? { localSheetId: item.localSheetIndex } : {}),
        })),
      },
    },
    stats,
  }
}

function columnLabel(index) {
  return XLSX.utils.encode_col(Math.max(0, index - 1))
}

/** Split a defined-name formula at top-level commas (sheet names may be quoted). */
function splitAreas(text) {
  const areas = []
  let current = ''
  let quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === "'") {
      if (quoted && text[index + 1] === "'") { current += "''"; index += 1; continue }
      quoted = !quoted
    }
    if (character === ',' && !quoted) { areas.push(current); current = '' } else current += character
  }
  areas.push(current)
  return areas.map((area) => area.trim()).filter(Boolean)
}

/**
 * One print area reference: an explicit range, or a whole-column ($A:$F) / whole-row ($1:$20)
 * area, which is bounded to the sheet's used size for printing and remembered in its original
 * form so an unchanged area is saved exactly as Excel wrote it.
 */
function parsePrintArea(reference, sheet) {
  const local = String(reference || '').replace(/^.*!/, '').trim()
  const range = normalizeRange(local)
  if (range) return { range }
  const columns = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/i.exec(local)
  if (columns) {
    const left = XLSX.utils.decode_col(columns[1].toUpperCase())
    const right = XLSX.utils.decode_col(columns[2].toUpperCase())
    if (left < 0 || right < 0 || right >= MAX_METADATA_COLS) return null
    const lastRow = Math.max(1, Math.min(MAX_METADATA_ROWS, Number(sheet && sheet.rowCount) || 1))
    return {
      range: normalizeRange(`${XLSX.utils.encode_col(Math.min(left, right))}1:${XLSX.utils.encode_col(Math.max(left, right))}${lastRow}`),
      whole: `$${XLSX.utils.encode_col(Math.min(left, right))}:$${XLSX.utils.encode_col(Math.max(left, right))}`,
    }
  }
  const rows = /^\$?(\d{1,7}):\$?(\d{1,7})$/.exec(local)
  if (rows) {
    const top = Math.min(Number(rows[1]), Number(rows[2]))
    const bottom = Math.max(Number(rows[1]), Number(rows[2]))
    if (top < 1 || bottom > MAX_METADATA_ROWS) return null
    const lastCol = Math.max(1, Math.min(MAX_METADATA_COLS, Number(sheet && sheet.colCount) || 1))
    return { range: normalizeRange(`A${top}:${columnLabel(lastCol)}${bottom}`), whole: `$${top}:$${bottom}` }
  }
  return null
}

function applyPrintAreasFromDefinedNames(sheets, definedNames, warnings) {
  const bySheet = new Map()
  for (const item of definedNames || []) {
    if (!item || item.name !== '_xlnm.Print_Area' || !Number.isInteger(item.localSheetIndex)) continue
    const sheet = sheets[item.localSheetIndex]
    if (!sheet) continue
    const entry = bySheet.get(sheet) || { areas: [], whole: {}, failed: false }
    for (const reference of splitAreas(String(item.ref || (item.ranges || []).join(',')))) {
      const parsed = parsePrintArea(reference, sheet)
      if (!parsed || !parsed.range) { entry.failed = true; continue }
      if (!entry.areas.includes(parsed.range)) entry.areas.push(parsed.range)
      if (parsed.whole) entry.whole[parsed.range] = parsed.whole
    }
    bySheet.set(sheet, entry)
  }
  for (const sheet of sheets) {
    const entry = bySheet.get(sheet)
    const pageSetup = { ...(sheet.pageSetup || {}) }
    const existing = typeof pageSetup.printArea === 'string' ? pageSetup.printArea.split('&&').filter(Boolean) : []
    // Without a defined name (metadata unreadable) keep the reader's value only if it is valid.
    if (!entry && existing.length && existing.every((area) => normalizeRange(area))) continue
    delete pageSetup.printArea
    delete pageSetup.printAreaWhole
    if (!entry && existing.length) warnings.push(`The print area of sheet "${sheet.name}" could not be read; the sheet prints its used range instead.`)
    if (entry && entry.areas.length) {
      pageSetup.printArea = entry.areas.join('&&')
      if (Object.keys(entry.whole).length) pageSetup.printAreaWhole = entry.whole
    }
    if (entry && entry.failed) warnings.push(`The print area of sheet "${sheet.name}" uses a reference simple_calc cannot read; ${entry.areas.length ? 'only the readable part is kept' : 'the sheet prints its used range instead'}.`)
    if (sheet.pageSetup) sheet.pageSetup = pageSetup
  }
}

function sheetJSErrorValue(cell) {
  const text = cell.w || ERROR_CODE_TO_TEXT[cell.v] || (typeof cell.v === 'string' ? cell.v : '#VALUE!')
  return String(text)
}

function sheetJSCellValue(cell, sourceDate1904 = false) {
  if (!cell || cell.t === 'z') return null
  if (cell.t === 'e') return sheetJSErrorValue(cell)
  if (cell.t === 'd') {
    const date = cell.v instanceof Date ? cell.v : new Date(cell.v)
    return Number.isNaN(date.getTime()) ? String(cell.v == null ? '' : cell.v) : dateToExcelSerial(date)
  }
  if (cell.t === 'b') return Boolean(cell.v)
  if (cell.t === 'n') {
    if (sourceDate1904 && Number.isFinite(cell.v) && cell.z && XLSX.SSF && typeof XLSX.SSF.is_date === 'function' && XLSX.SSF.is_date(cell.z)) {
      return cell.v + 1462
    }
    return Number.isFinite(cell.v) ? cell.v : null
  }
  if (cell.v == null) return null
  return typeof cell.v === 'string' ? cell.v : String(cell.v)
}

function sheetJSNote(cell) {
  if (!cell || !Array.isArray(cell.c) || !cell.c.length) return undefined
  return {
    comments: cell.c.map((comment) => ({
      author: String((comment && (comment.a || comment.author)) || ''),
      text: String((comment && (comment.t || comment.text)) || ''),
      ...(comment && comment.hidden != null ? { hidden: Boolean(comment.hidden) } : {}),
    })),
    ...(cell.c.hidden != null ? { hidden: Boolean(cell.c.hidden) } : {}),
  }
}

function mergeImportedStyles(sheetJsStyle, legacyStyle) {
  const fallback = normalizeStyle(sheetJsStyle)
  const rich = normalizeStyle(legacyStyle)
  if (!fallback) return rich
  if (!rich) return fallback
  return {
    ...fallback,
    ...rich,
    ...(rich.font || fallback.font ? { font: rich.font || fallback.font } : {}),
    ...(rich.fill || fallback.fill ? { fill: rich.fill || fallback.fill } : {}),
    ...(rich.border || fallback.border ? { border: rich.border || fallback.border } : {}),
    ...(rich.alignment || fallback.alignment ? { alignment: rich.alignment || fallback.alignment } : {}),
    ...(rich.protection || fallback.protection ? { protection: rich.protection || fallback.protection } : {}),
  }
}

function sheetJSCellToModel(cell, sourceDate1904 = false, legacyCell) {
  const formula = normalizeFormula(cell.f)
  const value = sheetJSCellValue(cell, sourceDate1904)
  const modelCell = {}
  if (formula) {
    modelCell.formula = formula
    if (cell.v !== undefined) modelCell.result = value
    if (cell.t === 'e') modelCell.resultType = 'error'
    if (cell.t === 'd') modelCell.resultType = 'date'
    if (typeof cell.F === 'string' && cell.F) modelCell.formulaRange = cell.F
    if (cell.D === true) modelCell.dynamicFormula = true
    // A plain (non-array) formula from a file calculates with implicit intersection, as in Excel.
    if (!modelCell.formulaRange && !modelCell.dynamicFormula) modelCell.implicitIntersection = true
  } else {
    modelCell.value = value
    if (cell.t === 'e') modelCell.type = 'error'
    if (cell.t === 'd') modelCell.type = 'date'
  }
  modelCell.display = cell.w != null ? String(cell.w) : value == null ? '' : String(richTextToPlainText(value))
  const style = mergeImportedStyles(cell.s, legacyCell && legacyCell.style)
  if (style) modelCell.style = style
  const numFmt = legacyCell && typeof legacyCell.numFmt === 'string'
    ? legacyCell.numFmt
    : typeof cell.z === 'string' ? cell.z : undefined
  if (numFmt && numFmt !== 'General') modelCell.numFmt = numFmt
  const hyperlink = normalizeHyperlink(cell.l)
  if (hyperlink) {
    modelCell.hyperlink = hyperlink.target
    if (hyperlink.tooltip) modelCell.hyperlinkTooltip = hyperlink.tooltip
  }
  const note = sheetJSNote(cell)
  if (note) modelCell.note = note
  return modelCell
}

function sheetVisibilityFromSheetJS(workbook, index) {
  const sheetInfo = workbook.Workbook && Array.isArray(workbook.Workbook.Sheets) ? workbook.Workbook.Sheets[index] : undefined
  if (!sheetInfo) return 'visible'
  if (sheetInfo.Hidden === 2) return 'veryHidden'
  if (sheetInfo.Hidden === 1) return 'hidden'
  return 'visible'
}

function safeSheetJSRange(sheet) {
  if (!sheet || typeof sheet['!ref'] !== 'string') return undefined
  try {
    return XLSX.utils.decode_range(sheet['!ref'])
  } catch {
    return undefined
  }
}

function sheetJSSheetToModel(workbook, sheetName, index, stats, warnings, sourceDate1904 = false, legacySheet) {
  const sheet = workbook.Sheets[sheetName] || {}
  const cells = {}
  let actualMaxRow = 0
  let actualMaxCol = 0

  for (const key of Object.keys(sheet)) {
    if (key.startsWith('!')) continue
    const position = addressPosition(key)
    if (!position) continue
    const rawCell = sheet[key]
    if (!rawCell || typeof rawCell !== 'object') continue
    const hasContent = rawCell.v != null || rawCell.f || rawCell.s || rawCell.z || rawCell.l || rawCell.c
    if (!hasContent) continue
    const modelCell = sheetJSCellToModel(rawCell, sourceDate1904, legacySheet && legacySheet.cells && legacySheet.cells[position.address])
    cells[position.address] = modelCell
    actualMaxRow = Math.max(actualMaxRow, position.row)
    actualMaxCol = Math.max(actualMaxCol, position.col)
    updateStatsForCell(stats, modelCell, position.row, position.col)
  }

  if (legacySheet && legacySheet.cells && typeof legacySheet.cells === 'object') {
    for (const [address, legacyCell] of Object.entries(legacySheet.cells)) {
      if (cells[address] || !legacyCell || typeof legacyCell !== 'object') continue
      const position = addressPosition(address)
      if (!position) continue
      const modelCell = { value: null, display: '' }
      const style = normalizeStyle(legacyCell.style)
      if (style) modelCell.style = style
      if (typeof legacyCell.numFmt === 'string' && legacyCell.numFmt && legacyCell.numFmt !== 'General') {
        modelCell.numFmt = legacyCell.numFmt
      }
      cells[position.address] = modelCell
      actualMaxRow = Math.max(actualMaxRow, position.row)
      actualMaxCol = Math.max(actualMaxCol, position.col)
      updateStatsForCell(stats, modelCell, position.row, position.col)
    }
  }

  const merges = []
  for (const rawMerge of Array.isArray(sheet['!merges']) ? sheet['!merges'] : []) {
    try {
      const merge = normalizeRange(XLSX.utils.encode_range(rawMerge))
      if (merge && !merges.includes(merge)) merges.push(merge)
    } catch {
      // Ignore malformed merge metadata without losing the sheet's cells.
    }
  }
  stats.mergeCount += merges.length

  const colWidths = {}
  const hiddenCols = []
  const columnProperties = {}
  let columnMetadataMax = 0
  ;(Array.isArray(sheet['!cols']) ? sheet['!cols'] : []).forEach((column, columnIndex) => {
    if (!column) return
    const width = Number.isFinite(column.width) ? column.width : Number.isFinite(column.wch) ? column.wch : undefined
    if (width > 0) {
      colWidths[String(columnIndex + 1)] = width
      columnMetadataMax = Math.max(columnMetadataMax, columnIndex + 1)
    }
    if (column.hidden === true) {
      hiddenCols.push(columnIndex + 1)
      columnMetadataMax = Math.max(columnMetadataMax, columnIndex + 1)
    }
    const properties = {}
    const outlineLevel = Number.isFinite(column.outlineLevel) ? column.outlineLevel : Number.isFinite(column.level) ? column.level : undefined
    if (outlineLevel > 0) properties.outlineLevel = outlineLevel
    if (column.collapsed === true) properties.collapsed = true
    if (column.bestFit === true) properties.bestFit = true
    const style = normalizeStyle(column.style || column.s)
    if (style) properties.style = style
    if (column.style && typeof column.style.numFmt === 'string' && column.style.numFmt !== 'General') {
      properties.numFmt = column.style.numFmt
    }
    if (Object.keys(properties).length) columnProperties[String(columnIndex + 1)] = properties
  })

  // BIFF templates commonly hide every column after the printable form all
  // the way through the legacy IV limit. Treat that contiguous, metadata-only
  // tail as a page-layout hint instead of collapsing the editor from F to IX.
  const meaningfulColumnMax = Math.max(
    actualMaxCol,
    0,
    ...Object.keys(colWidths).map(Number),
    ...Object.keys(columnProperties).map(Number),
  )
  const hiddenTailStart = hiddenCols[0]
  const isLegacyHiddenTail = Number.isInteger(hiddenTailStart) && hiddenTailStart > meaningfulColumnMax &&
    hiddenCols.length >= 200 && hiddenCols.every((column, position) => column === hiddenTailStart + position)
  if (isLegacyHiddenTail) {
    hiddenCols.length = 0
    columnMetadataMax = meaningfulColumnMax
  }

  const rowHeights = {}
  const hiddenRows = []
  const rowProperties = {}
  let rowMetadataMax = 0
  for (const [row, height] of Object.entries(legacySheet?.rowHeights || {})) {
    if (Number(row) >= 1 && Number(row) <= MAX_METADATA_ROWS && Number.isFinite(height) && height > 0) {
      rowHeights[row] = height
      rowMetadataMax = Math.max(rowMetadataMax, Number(row))
    }
  }
  ;(Array.isArray(sheet['!rows']) ? sheet['!rows'] : []).forEach((row, rowIndex) => {
    if (!row) return
    const height = Number.isFinite(row.hpt) ? row.hpt : Number.isFinite(row.hpx) ? row.hpx * 0.75 : undefined
    if (height > 0) {
      rowHeights[String(rowIndex + 1)] = height
      rowMetadataMax = Math.max(rowMetadataMax, rowIndex + 1)
    }
    if (row.hidden === true) {
      hiddenRows.push(rowIndex + 1)
      rowMetadataMax = Math.max(rowMetadataMax, rowIndex + 1)
    }
    const properties = {}
    const outlineLevel = Number.isFinite(row.outlineLevel) ? row.outlineLevel : Number.isFinite(row.level) ? row.level : undefined
    if (outlineLevel > 0) properties.outlineLevel = outlineLevel
    if (row.collapsed === true) properties.collapsed = true
    const style = normalizeStyle(row.style || row.s)
    if (style) properties.style = style
    if (row.style && typeof row.style.numFmt === 'string' && row.style.numFmt !== 'General') properties.numFmt = row.style.numFmt
    if (Object.keys(properties).length) rowProperties[String(rowIndex + 1)] = properties
  })

  const ref = safeSheetJSRange(sheet)
  const refMaxRow = ref ? ref.e.r + 1 : 0
  const refMaxCol = ref ? ref.e.c + 1 : 0
  const rawRowCount = Math.max(actualMaxRow, refMaxRow, rowMetadataMax)
  const rawColCount = Math.max(actualMaxCol, refMaxCol, columnMetadataMax)
  if (rawRowCount > MAX_METADATA_ROWS || rawColCount > MAX_METADATA_COLS) {
    warnings.push(`Sheet "${sheetName}" has oversized range metadata; its sparse cells were retained.`)
  }

  const model = {
    id: `sheet-${index + 1}`,
    name: sheetName || `Sheet${index + 1}`,
    sourceSheetName: sheetName || `Sheet${index + 1}`,
    sourceSheetIndex: index,
    state: sheetVisibilityFromSheetJS(workbook, index),
    rowCount: Math.max(1, Math.min(rawRowCount || 1, MAX_METADATA_ROWS)),
    colCount: Math.max(1, Math.min(rawColCount || 1, MAX_METADATA_COLS)),
    cells,
    merges,
    colWidths,
    rowHeights,
    hiddenCols,
    hiddenRows,
    columnProperties,
    rowProperties,
    frozen: null,
    ...(legacySheet?.properties ? { properties: plainClone(legacySheet.properties) } : {}),
    ...(isLegacyHiddenTail ? { properties: { ...plainClone(legacySheet?.properties || {}), legacyHiddenColumnTail: { start: hiddenTailStart, end: 256 } } } : {}),
  }
  if (sheet['!autofilter']) model.autoFilter = plainClone(sheet['!autofilter'])
  if (legacySheet?.pageSetup) model.pageSetup = plainClone(legacySheet.pageSetup)
  if (sheet['!margins']) model.pageSetup = { ...model.pageSetup, margins: plainClone(sheet['!margins']) }
  if (legacySheet?.headerFooter) model.headerFooter = plainClone(legacySheet.headerFooter)
  const printName = (workbook.Workbook?.Names || []).find((name) => name.Name === '_xlnm.Print_Area' && name.Sheet === index)
  if (printName?.Ref) {
    const parsed = splitAreas(decodeXml(String(printName.Ref))).map((reference) => parsePrintArea(reference, model)).filter((area) => area && area.range)
    if (parsed.length) {
      const whole = Object.fromEntries(parsed.filter((area) => area.whole).map((area) => [area.range, area.whole]))
      model.pageSetup = { ...model.pageSetup, printArea: [...new Set(parsed.map((area) => area.range))].join('&&'), ...(Object.keys(whole).length ? { printAreaWhole: whole } : {}) }
    }
  }
  if (sheet['!protect']) model.sheetProtection = plainClone(sheet['!protect'])
  if (sheet['!outline']) model.outline = plainClone(sheet['!outline'])
  return model
}

function definedNamesFromSheetJS(workbook) {
  const names = workbook.Workbook && Array.isArray(workbook.Workbook.Names) ? workbook.Workbook.Names : []
  return names
    .filter((item) => item && item.Name && item.Ref)
    .map((item) => {
      const result = { name: String(item.Name), ranges: [String(item.Ref)] }
      if (Number.isInteger(item.Sheet)) result.localSheetIndex = item.Sheet
      if (item.Hidden != null) result.hidden = Boolean(item.Hidden)
      if (item.Comment) result.comment = String(item.Comment)
      return result
    })
}

function activeSheetIndexFromSheetJS(workbook, sheetCount) {
  const views = workbook.Workbook && (workbook.Workbook.WBView || workbook.Workbook.Views)
  const view = Array.isArray(views) ? views[0] : undefined
  const index = view && Number.isInteger(view.activeTab) ? view.activeTab : 0
  return Math.max(0, Math.min(index, Math.max(0, sheetCount - 1)))
}

function blankSheetModel() {
  return {
    id: 'sheet-1',
    name: 'Sheet1',
    state: 'visible',
    rowCount: 1,
    colCount: 1,
    cells: {},
    merges: [],
    colWidths: {},
    rowHeights: {},
    hiddenCols: [],
    hiddenRows: [],
    frozen: null,
  }
}

function compatibilityWarnings(format, options = {}) {
  const warnings = []
  const officeEngine = options.officeEngine !== false
  if (format === 'xls' || format === 'xlt' || format === 'xla') {
    warnings.push(officeEngine
      ? 'Save keeps .xls and creates an original backup. Macros and some advanced Excel features may change after editing.'
      : 'Saving your edits creates an .xlsx copy next to the original; the .xls file itself is not changed. Macros are not kept in the copy.')
  } else if (format === 'ods') {
    warnings.push(officeEngine
      ? 'Save keeps .ods and creates an original backup. Some advanced spreadsheet features may change after editing.'
      : 'Opened without the document engine; some formatting was simplified. Saving your edits creates an .xlsx copy next to the original; the .ods file itself is not changed.')
  } else if (format === 'fods') {
    warnings.push('Opened through the compatibility reader; some formatting was simplified. Save as XLSX to keep your edits.')
  } else if (format !== MODERN_FORMAT && RICH_OOXML_FORMATS.has(format)) {
    warnings.push(`Opened ${format.toUpperCase()} as an OOXML workbook. Save as XLSX for modern interchange.`)
  } else if (format !== MODERN_FORMAT) {
    warnings.push(`Opened ${format.toUpperCase()} through the compatibility importer. Save as XLSX for the best interchange fidelity.`)
  }
  if (['xlsm', 'xltm', 'xlam', 'xla'].includes(format)) {
    warnings.push('VBA macros are not executed or represented in the editable model and will not be included in an XLSX save.')
  }
  if (['csv', 'tsv', 'tab', 'txt', 'prn', 'dif', 'dbf'].includes(format)) {
    warnings.push('This source format cannot store every workbook feature; formulas, styles, comments, or extra sheets may not exist in the source.')
  }
  return warnings
}

const DELIMITED_TEXT_FORMATS = new Set(['csv', 'tsv', 'tab', 'txt', 'prn'])

/**
 * Delimited text through delimited-text.cjs: per-column number and date inference, Excel's
 * quote rules and the source dialect (delimiter, decimal comma, encoding, BOM, line endings),
 * which is kept in the metadata so an in-format save writes the file back the same way.
 */
function importDelimitedText(buffer, sourceName, sourceFormat, warnings, options = {}) {
  if (!buffer.length) return undefined
  const stats = initialStats(buffer.length)
  const read = readDelimited(buffer, { sourceFormat, locale: options.locale, onCell: (cell, row, col) => updateStatsForCell(stats, cell, row, col) })
  if (!read) return undefined
  if (read.truncated) warnings.push('Some rows or columns beyond modern spreadsheet limits were dropped from this delimited text file.')
  for (const note of read.notes) warnings.push(note)
  const sheet = {
    ...blankSheetModel(),
    sourceSheetName: 'Sheet1',
    sourceSheetIndex: 0,
    rowCount: Math.max(1, read.maxRow),
    colCount: Math.max(1, read.maxCol),
    cells: read.cells,
    columnProperties: {},
    rowProperties: {},
  }
  stats.sheetCount = 1
  stats.sheets = 1
  return {
    model: {
      version: 1,
      name: path.parse(sourceName).name || 'Workbook',
      activeSheetId: sheet.id,
      sheets: [sheet],
      definedNames: [],
      metadata: {
        importedWith: 'delimited-text',
        delimiter: read.dialect.delimiter,
        decimalComma: read.dialect.decimalComma,
        dialect: { ...read.dialect, format: sourceFormat },
        properties: {},
        date1904: false,
        sourceDate1904: false,
        workbookProperties: {},
        workbookViews: [],
        calcProperties: {},
        definedNames: [],
      },
    },
    stats,
  }
}

/**
 * SheetJS reports OpenDocument number styles it cannot map on console.error. The calls are
 * synchronous, so the notes can be filtered for exactly their duration.
 */
function withoutSheetJSFormatNotes(run) {
  const original = console.error
  console.error = (...args) => {
    if (typeof args[0] === 'string' && /^(ODS number format may be incorrect|unrecognized character .* in ODF format)/.test(args[0])) return
    original.apply(console, args)
  }
  try {
    return run()
  } finally {
    console.error = original
  }
}

function importWithSheetJS(buffer, sourceName, sourceFormat, warnings) {
  let sheetWorkbook
  if (!buffer.length && ['csv', 'tsv', 'tab', 'txt', 'prn'].includes(sourceFormat)) {
    sheetWorkbook = { SheetNames: ['Sheet1'], Sheets: { Sheet1: { '!ref': 'A1:A1' } } }
  } else {
    sheetWorkbook = withoutSheetJSFormatNotes(() => XLSX.read(buffer, {
      type: 'buffer',
      // Infer ordinary CSV/text numbers while keeping dates as serial values
      // (cellDates:false) so opening a file never shifts a date by timezone.
      raw: false,
      dense: false,
      cellFormula: true,
      cellNF: true,
      cellStyles: true,
      cellText: true,
      cellDates: false,
      sheetStubs: true,
      xlfn: true,
      bookVBA: true,
      bookDeps: false,
      WTF: false,
    }))
  }

  const stats = initialStats(buffer.length)
  if (sheetWorkbook.vbaraw) warnings.push('This workbook contains VBA macros. Macros are not represented in the editor and will be omitted from an edited save.')
  const rawDate1904 = sheetWorkbook.Workbook && sheetWorkbook.Workbook.WBProps && sheetWorkbook.Workbook.WBProps.date1904
  const sourceDate1904 = rawDate1904 === true || rawDate1904 === 1 || rawDate1904 === '1' || rawDate1904 === 'true'
  if (buffer.length && !(Array.isArray(sheetWorkbook.SheetNames) && sheetWorkbook.SheetNames.length)) {
    throw new Error('No sheets could be read from this file.')
  }
  const sheetNames = Array.isArray(sheetWorkbook.SheetNames) && sheetWorkbook.SheetNames.length ? sheetWorkbook.SheetNames : ['Sheet1']
  if (!sheetWorkbook.Sheets) sheetWorkbook.Sheets = { Sheet1: { '!ref': 'A1:A1' } }
  let legacyStyles = null
  if (['xls', 'xlt', 'xla'].includes(sourceFormat)) {
    try {
      legacyStyles = extractLegacyBiffStyles(buffer, sheetWorkbook.SSF)
    } catch {
      warnings.push('The legacy workbook opened, but some BIFF cell formatting could not be decoded.')
    }
  }
  const sheets = sheetNames.map((sheetName, index) => sheetJSSheetToModel(
    sheetWorkbook,
    sheetName,
    index,
    stats,
    warnings,
    sourceDate1904,
    legacyStyles && legacyStyles.sheets[index],
  ))
  stats.sheetCount = sheets.length
  stats.sheets = sheets.length
  const activeIndex = activeSheetIndexFromSheetJS(sheetWorkbook, sheets.length)
  if (sheets.some(sheet => Object.values(sheet.headerFooter || {}).some(value => typeof value === 'string' && /&G/i.test(value)))) {
    warnings.push('This file has header images. They may not appear in print previews or edited copies. An unchanged save preserves the original file.')
  }
  const properties = workbookPropertiesFromSheetJS(sheetWorkbook)
  const definedNames = attachDefinedNameSheetIds(definedNamesFromSheetJS(sheetWorkbook), sheets)
  return {
    model: {
      version: 1,
      name: path.parse(sourceName).name || 'Workbook',
      activeSheetId: sheets[activeIndex] ? sheets[activeIndex].id : sheets[0].id,
      sheets,
      definedNames,
      metadata: {
        importedWith: 'sheetjs',
        ...(legacyStyles?.normalFont ? { normalFont: normalizeFont(legacyStyles.normalFont) } : {}),
        ...(legacyStyles ? { legacyStyles: { fonts: legacyStyles.fontCount, xfs: legacyStyles.xfCount } } : {}),
        ...properties,
        properties,
        date1904: false,
        sourceDate1904,
        workbookProperties: plainClone((sheetWorkbook.Workbook && sheetWorkbook.Workbook.WBProps) || {}),
        workbookViews: plainClone((sheetWorkbook.Workbook && (sheetWorkbook.Workbook.WBView || sheetWorkbook.Workbook.Views)) || []),
        calcProperties: plainClone((sheetWorkbook.Workbook && sheetWorkbook.Workbook.CalcPr) || {}),
        definedNames: definedNames.map((item) => ({
          name: item.name,
          ranges: item.ranges.join(','),
          ...(Number.isInteger(item.localSheetIndex) ? { localSheetId: item.localSheetIndex } : {}),
        })),
      },
    },
    stats,
  }
}

// ---------------------------------------------------------------------------
// OpenDocument without the document engine (SheetJS reader plus post-processing)
// ---------------------------------------------------------------------------

function quoteSheetNameForFormula(name) {
  const text = String(name)
  return /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(text) && !/^[A-Za-z]{1,3}\d+$/.test(text) && !/^R\d*C\d*$/i.test(text)
    ? text
    : `'${text.replace(/'/g, "''")}'`
}

/** One OpenFormula reference body ("Sheet.A1:.B2", ".$C$3", "'My sheet'.A1") as A1 text. */
function openFormulaReference(body) {
  const parts = []
  let current = ''
  let quoted = false
  for (let index = 0; index < body.length; index += 1) {
    const character = body[index]
    if (character === "'") {
      if (quoted && body[index + 1] === "'") { current += "''"; index += 1; continue }
      quoted = !quoted
    }
    if (character === ':' && !quoted) { parts.push(current); current = '' } else current += character
  }
  parts.push(current)
  if (parts.length > 2) return null
  const decoded = parts.map((part) => {
    const match = /^\$?(?:'((?:[^']|'')*)'|([^.']*))\.(.+)$/.exec(part.trim())
    if (!match) return null
    const sheet = match[1] != null ? decodeXml(match[1].replace(/''/g, "'")) : decodeXml(match[2] || '')
    return { sheet, cell: match[3].trim() }
  })
  if (decoded.some((part) => !part || !part.cell)) return null
  const [first, second] = decoded
  const prefix = first.sheet ? `${quoteSheetNameForFormula(first.sheet)}${second && second.sheet && second.sheet !== first.sheet ? `:${quoteSheetNameForFormula(second.sheet)}` : ''}!` : ''
  return `${prefix}${first.cell}${second ? `:${second.cell}` : ''}`
}

/** Translate the bracketed OpenFormula references SheetJS leaves in ODS formulas into A1 syntax. */
function openFormulaToA1(formula) {
  const text = String(formula || '')
  if (!text.includes('[')) return text
  let output = ''
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === '"') {
      const end = text.indexOf('"', index + 1)
      let stop = end < 0 ? text.length : end + 1
      while (end >= 0 && text[stop] === '"') {
        const next = text.indexOf('"', stop + 1)
        stop = next < 0 ? text.length : next + 1
        if (next < 0) break
      }
      output += text.slice(index, stop)
      index = stop - 1
      continue
    }
    if (character === '[') {
      const end = text.indexOf(']', index + 1)
      if (end > index) {
        const translated = openFormulaReference(text.slice(index + 1, end))
        if (translated) {
          output += translated
          index = end
          continue
        }
      }
    }
    output += character
  }
  return output
}

function odfLengthToPoints(value) {
  const match = /^([0-9.]+)\s*(cm|mm|in|pt|pc|px)?$/i.exec(String(value || '').trim())
  if (!match) return null
  const number = Number(match[1])
  if (!Number.isFinite(number)) return null
  const unit = (match[2] || 'pt').toLowerCase()
  return number * ({ cm: 72 / 2.54, mm: 72 / 25.4, in: 72, pt: 1, pc: 12, px: 0.75 }[unit] || 1)
}

/** Table and column styles of an OpenDocument content.xml (hidden sheets, column widths). */
function openDocumentLayout(xml) {
  const styles = new Map()
  for (const match of xml.matchAll(/<style:style\b([^>]*?)(?:\/>|>([\s\S]*?)<\/style:style>)/g)) {
    const attributes = parseXmlAttributes(match[1])
    const name = attributes['style:name']
    if (!name) continue
    const body = match[2] || ''
    const table = /<style:table-properties\b([^>]*)>/.exec(body)
    const column = /<style:table-column-properties\b([^>]*)>/.exec(body)
    styles.set(name, {
      hidden: Boolean(table && parseXmlAttributes(table[1])['table:display'] === 'false'),
      width: column ? odfLengthToPoints(parseXmlAttributes(column[1])['style:column-width']) : null,
    })
  }
  const tables = []
  const tablePattern = /<table:table(?=[\s>])([^>]*)>/g
  let match
  while ((match = tablePattern.exec(xml))) {
    const attributes = parseXmlAttributes(match[1])
    const start = match.index + match[0].length
    const rowAt = xml.slice(start).search(/<table:table-row\b/)
    const head = xml.slice(start, rowAt < 0 ? Math.min(xml.length, start + 200_000) : start + rowAt)
    const columns = []
    for (const column of head.matchAll(/<table:table-column(?=[\s/>])([^>]*?)\/?>/g)) {
      const columnAttributes = parseXmlAttributes(column[1])
      const repeat = Math.max(1, Math.min(MAX_METADATA_COLS, Number(columnAttributes['table:number-columns-repeated']) || 1))
      const style = styles.get(columnAttributes['table:style-name'])
      columns.push({ repeat, width: style ? style.width : null, hidden: columnAttributes['table:visibility'] === 'collapse' })
    }
    const style = styles.get(attributes['table:style-name'])
    tables.push({ name: attributes['table:name'] || '', hidden: Boolean(style && style.hidden), columns })
  }
  return tables
}

async function openDocumentContentXml(buffer, format) {
  if (format === 'fods') return buffer.toString('utf8')
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file('content.xml')
  if (!entry) return ''
  const size = Number(entry._data && entry._data.uncompressedSize)
  if (Number.isFinite(size) && size > MAX_SOURCE_BYTES) return ''
  return entry.async('string')
}

/** Number format that reproduces a cell's displayed text, for formats SheetJS misreads from ODS. */
function numberFormatFromDisplay(value, display) {
  const { parseNumber, numberFormatFor } = require('./delimited-text.cjs')
  for (const convention of ['dot', 'comma']) {
    const parsed = parseNumber(String(display || ''), convention)
    if (parsed && Math.abs(parsed.value - value) <= Math.max(1e-9, Math.abs(value) * 1e-9)) return numberFormatFor(parsed) || null
  }
  return null
}

/**
 * SheetJS reads OpenDocument values and formulas but leaves XML entities in sheet names,
 * OpenFormula references in formulas and drops hidden-sheet state and column widths. Restore
 * them from content.xml so a file opened without the document engine is faithful.
 */
async function postProcessOpenDocument(imported, buffer, format) {
  const model = imported.model
  for (const sheet of model.sheets) {
    const decoded = decodeXml(sheet.name)
    sheet.name = decoded
    sheet.sourceSheetName = decoded
  }
  for (const sheet of model.sheets) {
    for (const cell of Object.values(sheet.cells || {})) {
      if (cell && typeof cell.formula === 'string' && cell.formula) cell.formula = openFormulaToA1(decodeXml(cell.formula))
      if (cell && typeof cell.value === 'number' && typeof cell.numFmt === 'string' && cell.numFmt.includes('@')) {
        const numFmt = numberFormatFromDisplay(cell.value, cell.display)
        if (numFmt) cell.numFmt = numFmt
        else delete cell.numFmt
      }
      if (cell && cell.formula && typeof cell.result === 'number' && typeof cell.numFmt === 'string' && cell.numFmt.includes('@')) {
        const numFmt = numberFormatFromDisplay(cell.result, cell.display)
        if (numFmt) cell.numFmt = numFmt
        else delete cell.numFmt
      }
    }
  }
  const decodeNames = (items) => (Array.isArray(items) ? items : []).forEach((item) => {
    if (!item) return
    if (Array.isArray(item.ranges)) item.ranges = item.ranges.map((range) => decodeXml(range))
    else if (typeof item.ranges === 'string') item.ranges = decodeXml(item.ranges)
    if (typeof item.ref === 'string') item.ref = decodeXml(item.ref)
  })
  decodeNames(model.definedNames)
  decodeNames(model.metadata && model.metadata.definedNames)
  let layout = []
  try {
    layout = openDocumentLayout(await openDocumentContentXml(buffer, format))
  } catch {
    layout = []
  }
  model.sheets.forEach((sheet, index) => {
    const table = layout[index]
    if (!table) return
    if (table.hidden) sheet.state = 'hidden'
    // Only the used columns: ODS files repeat a default column style to the sheet's edge.
    let column = 1
    const usedColumns = Math.max(1, Number(sheet.colCount) || 1)
    for (const entry of table.columns) {
      for (let repeat = 0; repeat < entry.repeat && column <= usedColumns; repeat += 1, column += 1) {
        if (Number.isFinite(entry.width) && entry.width > 0) {
          const pixels = entry.width * (4 / 3)
          sheet.colWidths = { ...(sheet.colWidths || {}), [String(column)]: Math.max(0.5, Math.round(((pixels - 5) / 7) * 100) / 100) }
        }
        if (entry.hidden) sheet.hiddenCols = [...new Set([...(sheet.hiddenCols || []), column])]
      }
      if (column > usedColumns) break
    }
  })
  if (!model.sheets.some((sheet) => sheet.state !== 'hidden' && sheet.state !== 'veryHidden')) model.sheets[0].state = 'visible'
  const active = model.sheets.find((sheet) => sheet.id === model.activeSheetId)
  if (!active || active.state !== 'visible') model.activeSheetId = (model.sheets.find((sheet) => sheet.state === 'visible') || model.sheets[0]).id
}

async function importOpenDocumentNatively(buffer, sourceName, format, warnings) {
  const imported = importWithSheetJS(buffer, sourceName, format, warnings)
  await postProcessOpenDocument(imported, buffer, format)
  imported.model.metadata.importedWith = format === 'fods' ? 'sheetjs-fods' : 'sheetjs-ods'
  return imported
}

function finalizePayload(sourceName, sourceFormat, imported, warnings, options = {}) {
  const safeName = path.basename(sourceName) || `Workbook.${sourceFormat}`
  const requiresSaveAs = Boolean(options.requiresSaveAs) || !['xlsx', 'xls', 'ods', 'csv', 'tsv'].includes(sourceFormat)
  const uniqueWarnings = [...new Set(warnings.filter(Boolean))]
  imported.model.metadata = {
    ...(imported.model.metadata || {}),
    sourceName: safeName,
    sourceFormat,
    requiresSaveAs,
    warnings: uniqueWarnings.slice(),
    stats: { ...imported.stats },
  }
  return {
    name: safeName,
    sourceFormat,
    requiresSaveAs,
    ...(options.saveAsFormat ? { saveAsFormat: options.saveAsFormat } : {}),
    ...(typeof options.officeEngine === 'boolean' ? { officeEngine: options.officeEngine } : {}),
    warnings: uniqueWarnings,
    stats: imported.stats,
    workbook: imported.model,
  }
}

const FORMAT_DESCRIPTIONS = Object.freeze({
  xlsx: 'an Excel workbook (.xlsx)',
  xlsm: 'a macro-enabled Excel workbook (.xlsm)',
  xlsb: 'an Excel binary workbook (.xlsb)',
  xls: 'an Excel 97-2003 workbook (.xls)',
  ods: 'an OpenDocument spreadsheet (.ods)',
  fods: 'a flat OpenDocument spreadsheet (.fods)',
  csv: 'comma-separated text (CSV)',
  tsv: 'tab-separated text',
  html: 'a web page (HTML table)',
  xml: 'an Excel 2003 XML spreadsheet',
  numbers: 'an Apple Numbers document',
})
const ZIP_SIGNATURES = [[0x50, 0x4b, 0x03, 0x04], [0x50, 0x4b, 0x05, 0x06], [0x50, 0x4b, 0x07, 0x08]]
const BINARY_EXTENSIONS = new Set(['dbf', 'wk1', 'wk2', 'wk3', 'wk4', 'wks', 'wq1', 'wq2', 'wb1', 'wb2', 'wb3', '123', 'qpw'])
const TEXT_EXTENSIONS = new Set(['csv', 'tsv', 'tab', 'txt', 'prn', 'slk', 'sylk', 'dif'])

function hasZipSignature(buffer) {
  return buffer.length >= 4 && ZIP_SIGNATURES.some((signature) => signature.every((byte, index) => buffer[index] === byte))
}

/** What the bytes are, whatever the extension says. */
async function sniffContainer(buffer) {
  if (!buffer.length) return 'empty'
  if (hasZipSignature(buffer)) {
    if (buffer.includes('xl/workbook.xml')) return 'xlsx'
    if (buffer.includes('xl/workbook.bin')) return 'xlsb'
    if (buffer.subarray(0, 512).includes('application/vnd.oasis.opendocument.spreadsheet')) return 'ods'
    if (buffer.includes('word/document.xml')) return 'docx'
    if (buffer.includes('ppt/presentation.xml')) return 'pptx'
    if (buffer.includes('Index/Document.iwa') || buffer.includes('Index.zip')) return 'numbers'
    if (buffer.includes('content.xml')) {
      try {
        const zip = await JSZip.loadAsync(buffer)
        const mime = zip.file('mimetype') ? await zip.file('mimetype').async('string') : ''
        if (/opendocument\.spreadsheet/.test(mime)) return 'ods'
        if (/opendocument\./.test(mime)) return 'odf-other'
      } catch {}
    }
    return 'zip'
  }
  if (buffer.length >= 8 && buffer.subarray(0, 8).equals(CFB_SIGNATURE)) return isEncryptedWorkbook(buffer) ? 'encrypted' : 'cfb'
  const head = buffer.subarray(0, 8192)
  const text = head.toString('utf8').replace(/^[\s﻿]+/, '')
  if (text.startsWith('<')) {
    if (/urn:schemas-microsoft-com:office:spreadsheet|progid="Excel\.Sheet"/i.test(text)) return 'xml'
    if (/<office:document\b/.test(text) || /opendocument\.spreadsheet/.test(text)) return 'fods'
    if (/<(!doctype\s+html|html|table|body|head|meta)\b/i.test(text)) return 'html'
    return 'xml-other'
  }
  // UTF-16 text has NULs in every other byte; anything else with NULs is binary.
  let zeros = 0
  for (const byte of head) if (byte === 0) zeros += 1
  if (zeros) {
    const utf16 = (head[0] === 0xff && head[1] === 0xfe) || (head[0] === 0xfe && head[1] === 0xff) || zeros >= head.length * 0.3
    if (!utf16) return 'binary'
  }
  return 'text'
}

/** The format the content really is, keeping the extension when it agrees with the bytes. */
function realSourceFormat(container, extensionFormat, sourceName, buffer) {
  switch (container) {
    case 'empty':
      if (RICH_OOXML_FORMATS.has(extensionFormat) || ['xls', 'xlsb', 'ods'].includes(extensionFormat)) throw new Error(`Unable to open ${sourceName}: the file is empty.`)
      return extensionFormat || 'csv'
    case 'xlsx':
      return RICH_OOXML_FORMATS.has(extensionFormat) ? extensionFormat : 'xlsx'
    case 'xlsb':
      return 'xlsb'
    case 'ods':
      return 'ods'
    case 'numbers':
      return 'numbers'
    case 'encrypted':
      throw new Error(ENCRYPTED_WORKBOOK_MESSAGE)
    case 'cfb':
      return ['xls', 'xlt', 'xla'].includes(extensionFormat) ? extensionFormat : 'xls'
    case 'docx':
    case 'pptx':
    case 'odf-other':
      throw new Error(`Unable to open ${sourceName}: it is a ${container === 'docx' ? 'Word document' : container === 'pptx' ? 'PowerPoint presentation' : 'text or presentation document'}, not a spreadsheet.`)
    case 'zip':
      if (RICH_OOXML_FORMATS.has(extensionFormat) || extensionFormat === 'xlsb') throw new Error(`Unable to open ${sourceName}: the file is not a valid ${extensionFormat.toUpperCase()} package.`)
      return extensionFormat
    case 'html':
      return ['html', 'htm'].includes(extensionFormat) ? extensionFormat : 'html'
    case 'xml':
      return 'xml'
    case 'fods':
      return 'fods'
    case 'xml-other':
      return extensionFormat === 'xml' ? 'xml' : extensionFormat === 'html' || extensionFormat === 'htm' ? extensionFormat : 'xml'
    case 'binary':
      if (BINARY_EXTENSIONS.has(extensionFormat) || extensionFormat === 'numbers') return extensionFormat
      if (RICH_OOXML_FORMATS.has(extensionFormat) || ['xls', 'xlsb', 'ods'].includes(extensionFormat)) throw new Error(`Unable to open ${sourceName}: the file is damaged or is not a ${extensionFormat.toUpperCase()} workbook.`)
      return extensionFormat || 'csv'
    case 'text':
    default: {
      if (TEXT_EXTENSIONS.has(extensionFormat)) return extensionFormat
      // A delimited text file under a workbook name (or no extension at all).
      const { sniffDelimiter, decodeText } = require('./delimited-text.cjs')
      const sample = decodeText(buffer.subarray(0, 65_536)).text
      return sniffDelimiter(sample, ',') === '\t' ? 'tsv' : 'csv'
    }
  }
}

function sameFormatFamily(extensionFormat, sourceFormat) {
  if (extensionFormat === sourceFormat) return true
  const family = (format) => (RICH_OOXML_FORMATS.has(format) ? 'ooxml' : ['csv', 'tsv', 'tab', 'txt', 'prn'].includes(format) ? 'text' : ['xls', 'xlt', 'xla'].includes(format) ? 'biff' : ['html', 'htm'].includes(format) ? 'html' : format)
  return family(extensionFormat) === family(sourceFormat)
}

const ODS_ENGINE_FAILED = 'The document engine could not open this file, so it was opened directly; some formatting was simplified.'

/**
 * Read a workbook from bytes. Returns the renderer payload plus the merge base: the OOXML package
 * the model's source identities (worksheet ids, chart parts, picture ids) point into. The main
 * process keeps that package for the life of the document, so later saves never mix the model
 * with a different package (for example the file written by the previous save).
 */
async function importWorkbookBytes(name, data, options = {}) {
  const extension = extensionFromName(name)
  if (extension && !SUPPORTED_EXTENSIONS.has(extension)) throw new Error(`Unsupported spreadsheet format: ${extension}`)
  const extensionFormat = extension.slice(1)
  const sourceName = path.basename(name)
  const buffer = bytesToBuffer(data)
  if (isEncryptedWorkbook(buffer)) throw new Error(ENCRYPTED_WORKBOOK_MESSAGE)
  const container = await sniffContainer(buffer)
  const sourceFormat = realSourceFormat(container, extensionFormat, sourceName, buffer)
  const officeEngine = typeof options.officeEngine === 'boolean' ? options.officeEngine : await officeEngineAvailable()
  const mismatch = Boolean(extensionFormat) && !sameFormatFamily(extensionFormat, sourceFormat)
  const warnings = compatibilityWarnings(sourceFormat, { officeEngine })
  if (mismatch) {
    warnings.unshift(`This file is actually ${FORMAT_DESCRIPTIONS[sourceFormat] || `a ${sourceFormat.toUpperCase()} file`}, although its name ends in .${extensionFormat}. It opened normally; saving creates a correctly named copy and does not overwrite it.`)
  }

  let imported
  let mergeBase = null
  try {
    if (sourceFormat === 'ods') {
      let converted = null
      if (officeEngine) {
        try {
          const { convertOfficeBytes } = require('./office-converter.cjs')
          converted = await convertOfficeBytes({ bytes: buffer, inputExtension: 'ods', outputExtension: 'xlsx', filter: 'Calc MS Excel 2007 XML' })
        } catch {
          warnings.push(ODS_ENGINE_FAILED)
        }
      }
      if (converted) {
        imported = await importWithExcelJS(converted, sourceName, warnings)
        imported.model.metadata.importedWith = 'office-ods'
        mergeBase = converted
      } else {
        imported = await importOpenDocumentNatively(buffer, sourceName, 'ods', warnings)
      }
    } else if (sourceFormat === 'fods') {
      imported = await importOpenDocumentNatively(buffer, sourceName, 'fods', warnings)
    } else if (RICH_OOXML_FORMATS.has(sourceFormat)) {
      try {
        imported = await importWithExcelJS(buffer, sourceName, warnings)
        mergeBase = buffer
      } catch (excelError) {
        warnings.push('The rich OOXML reader could not open this file; the compatibility reader was used instead, so some formatting may be simplified.')
        try {
          imported = importWithSheetJS(buffer, sourceName, sourceFormat, warnings)
        } catch (sheetError) {
          const error = new Error(`Unable to open ${sourceName}: ${sheetError.message || excelError.message}`)
          error.cause = sheetError
          throw error
        }
      }
    } else {
      if (DELIMITED_TEXT_FORMATS.has(sourceFormat)) {
        try {
          imported = importDelimitedText(buffer, sourceName, sourceFormat, warnings, options)
        } catch {
          warnings.push('The delimited-text reader could not parse this file; the compatibility reader was used instead.')
        }
      }
      if (!imported) imported = importWithSheetJS(buffer, sourceName, sourceFormat, warnings)
    }
  } catch (error) {
    if (/^Unable to open /.test(String(error && error.message))) throw error
    if (/password|encrypt/i.test(String(error && error.message))) {
      const wrapped = new Error(ENCRYPTED_WORKBOOK_MESSAGE)
      wrapped.cause = error
      throw wrapped
    }
    const wrapped = new Error(`Unable to open ${sourceName}: ${error.message || 'invalid or damaged spreadsheet'}`)
    wrapped.cause = error
    throw wrapped
  }
  const legacyWithoutEngine = ['xls', 'xlt', 'xla', 'ods'].includes(sourceFormat) && !officeEngine
  const payload = finalizePayload(sourceName, sourceFormat, imported, warnings, {
    requiresSaveAs: mismatch || legacyWithoutEngine,
    saveAsFormat: mismatch || legacyWithoutEngine ? (['csv', 'tsv'].includes(sourceFormat) ? sourceFormat : 'xlsx') : undefined,
    officeEngine,
  })
  return { payload, mergeBase }
}

async function workbookPayloadFromBytes(name, data, options = {}) {
  return (await importWorkbookBytes(name, data, options)).payload
}

/** Kept for callers that only need the extension check (SUPPORTED_EXTENSIONS or no extension). */
function supportedOrExtensionless(name) {
  const extension = extensionFromName(name)
  return !extension || SUPPORTED_EXTENSIONS.has(extension)
}

const CFB_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
const ENCRYPTED_STREAM_MARKERS = [Buffer.from('EncryptedPackage', 'utf16le'), Buffer.from('EncryptionInfo', 'utf16le')]
const ENCRYPTED_WORKBOOK_MESSAGE = 'This workbook is password-protected; encrypted files are not supported.'

function isEncryptedWorkbook(buffer) {
  if (buffer.length < 8 || !buffer.subarray(0, 8).equals(CFB_SIGNATURE)) return false
  // Standard and agile OOXML encryption both store these CFB stream names.
  return ENCRYPTED_STREAM_MARKERS.some((marker) => buffer.includes(marker))
}

async function workbookPayloadFromPath(filePath, options = {}) {
  return (await importWorkbookPath(filePath, options)).payload
}

async function importWorkbookPath(filePath, options = {}) {
  if (typeof filePath !== 'string' || !filePath.trim()) throw new TypeError('A spreadsheet file path is required.')
  const resolvedPath = path.resolve(filePath)
  if (!supportedOrExtensionless(resolvedPath)) ensureSupportedExtension(resolvedPath)
  let stat
  try {
    stat = await fs.promises.stat(resolvedPath)
  } catch (error) {
    const wrapped = new Error(`Unable to access spreadsheet: ${error.message}`)
    wrapped.cause = error
    throw wrapped
  }
  if (!stat.isFile()) throw new Error('The selected spreadsheet path is not a file.')
  if (stat.size > MAX_SOURCE_BYTES) throw new RangeError(`Spreadsheet exceeds the ${MAX_SOURCE_BYTES / 1024 / 1024} MB import limit.`)
  const data = await fs.promises.readFile(resolvedPath)
  const imported = await importWorkbookBytes(path.basename(resolvedPath), data, options)
  return { ...imported, bytes: data, stat }
}

function unwrapWorkbookModel(model) {
  if (model && model.workbook && Array.isArray(model.workbook.sheets)) return model.workbook
  return model
}

function validateWorkbookModel(input) {
  const model = unwrapWorkbookModel(input)
  if (!model || typeof model !== 'object') throw new TypeError('A workbook model is required.')
  if (!Array.isArray(model.sheets) || !model.sheets.length) throw new Error('The workbook must contain at least one sheet.')
  if (model.sheets.length > 10_000) throw new RangeError('The workbook contains too many sheets.')
  const names = new Set()
  for (const [index, sheet] of model.sheets.entries()) {
    if (!sheet || typeof sheet !== 'object') throw new TypeError(`Sheet ${index + 1} is invalid.`)
    const name = typeof sheet.name === 'string' ? sheet.name.trim() : ''
    if (!name) throw new Error(`Sheet ${index + 1} needs a name.`)
    if (name.length > MAX_SHEET_NAME_LENGTH || /[\\/?*\[\]:]/.test(name)) throw new Error(`Sheet name "${name}" is not valid in XLSX.`)
    const folded = name.toLocaleLowerCase()
    if (names.has(folded)) throw new Error(`Duplicate sheet name: ${name}`)
    names.add(folded)
    if (sheet.cells != null && (typeof sheet.cells !== 'object' || Array.isArray(sheet.cells))) {
      throw new TypeError(`Cells for sheet "${name}" must be an A1-keyed object.`)
    }
  }
  if (!model.sheets.some((sheet) => sheet.state !== 'hidden' && sheet.state !== 'veryHidden')) {
    throw new Error('At least one workbook sheet must remain visible.')
  }
  return model
}

function applyWorkbookProperties(excelWorkbook, properties) {
  if (!properties || typeof properties !== 'object') return
  for (const key of [
    'creator',
    'lastModifiedBy',
    'title',
    'subject',
    'description',
    'keywords',
    'category',
    'company',
    'manager',
    'language',
    'contentStatus',
  ]) {
    if (typeof properties[key] === 'string') excelWorkbook[key] = properties[key]
  }
  if (properties.revision !== undefined && properties.revision !== null) excelWorkbook.revision = plainClone(properties.revision)
  for (const key of ['created', 'modified', 'lastPrinted']) {
    if (!properties[key]) continue
    const date = new Date(properties[key])
    if (!Number.isNaN(date.getTime())) excelWorkbook[key] = date
  }
}

function modelNoteToExcelJS(note) {
  if (typeof note === 'string') return note
  if (!note || typeof note !== 'object') return String(note == null ? '' : note)
  if (Array.isArray(note.comments)) {
    return {
      texts: note.comments.map((comment, index) => ({
        text: `${index ? '\n' : ''}${comment.author ? `${comment.author}: ` : ''}${comment.text || ''}`,
      })),
    }
  }
  if (Array.isArray(note.texts)) return plainClone(note)
  if (typeof note.text === 'string') return note.text
  return JSON.stringify(note)
}

function formulaResultForExcelJS(result, resultType) {
  if (resultType === 'error' || (typeof result === 'string' && ERROR_TEXT_TO_CODE[result] != null)) {
    return { error: String(result || '#VALUE!') }
  }
  const value = deserializeCellValue(result)
  if (value && typeof value === 'object' && value.richText) {
    return value.richText.map((run) => run.text || '').join('')
  }
  if (Buffer.isBuffer(value)) return undefined
  return value
}

/**
 * Write one model cell. Hyperlinks ExcelJS cannot write (links to a place in the workbook and
 * links on formula cells) are collected in `pendingLinks` and added to the package afterwards.
 */
function setExcelJSCell(cell, modelCell, pendingLinks) {
  const formula = normalizeFormula(modelCell && modelCell.formula)
  const hyperlink = normalizeHyperlink(modelCell && modelCell.hyperlink)
  const tooltip = hyperlink && (hyperlink.tooltip || (modelCell && modelCell.hyperlinkTooltip))
  if (formula) {
    // Every formula is written as the cell's own formula. Shared (filled-down) groups are
    // never re-emitted: editing the first cell of a fill must not change or break the others.
    const formulaValue = { formula }
    if (modelCell.result !== undefined) {
      const result = formulaResultForExcelJS(modelCell.result, modelCell.resultType)
      if (result !== undefined) formulaValue.result = result
    }
    const arrayRange = typeof modelCell.formulaRange === 'string' ? normalizeRange(modelCell.formulaRange) : undefined
    const isArray = modelCell.formulaType === 'array' || (!modelCell.formulaType && Boolean(arrayRange))
    if (isArray && arrayRange && arrayRange.split(':')[0] === normalizeAddress(cell.address)) {
      formulaValue.shareType = 'array'
      formulaValue.ref = arrayRange
    }
    cell.value = formulaValue
    if (hyperlink && Array.isArray(pendingLinks)) pendingLinks.push({ address: cell.address, target: hyperlink.target, ...(tooltip ? { tooltip } : {}) })
  } else {
    let value = deserializeCellValue(modelCell ? modelCell.value : null)
    const runs = modelCell && Array.isArray(modelCell.richText) ? modelCell.richText : null
    if (runs && typeof value === 'string' && runs.map((run) => String(run && run.text != null ? run.text : '')).join('') === value) {
      value = deserializeCellValue({ type: 'richText', runs })
    }
    if (modelCell && modelCell.type === 'error') value = { error: String(modelCell.value || '#VALUE!') }
    if (hyperlink && hyperlink.target.startsWith('#')) {
      // A place in this workbook: written as <hyperlink location> without a relationship.
      if (Array.isArray(pendingLinks)) pendingLinks.push({ address: cell.address, target: hyperlink.target, ...(tooltip ? { tooltip } : {}) })
    } else if (hyperlink) {
      const textValue = richTextToPlainText(modelCell.value)
      value = {
        text: textValue == null || typeof textValue === 'object' ? String(modelCell.display || hyperlink.target) : String(textValue),
        hyperlink: hyperlink.target,
      }
      if (tooltip) value.tooltip = tooltip
    }
    cell.value = value
  }

  const style = normalizeStyle(modelCell && modelCell.style)
  if (style) cell.style = style
  if (modelCell && typeof modelCell.numFmt === 'string' && modelCell.numFmt) cell.numFmt = modelCell.numFmt
  if (modelCell && modelCell.note !== undefined) {
    cell.note = modelNoteToExcelJS(modelCell.note)
    // ExcelJS omits a completely empty, unstyled cell from the worksheet XML,
    // which also drops its note. A neutral style keeps the cell record (and
    // note) without turning it into a value or changing formula semantics.
    if (cell.value == null && !cell.hasStyle) cell.alignment = {}
  }
}

function applyExcelJSSheetDimensions(worksheet, sheet) {
  for (const [key, rawProperties] of Object.entries(sheet.columnProperties || {})) {
    const index = Number(key)
    if (!Number.isInteger(index) || index < 1 || index > MAX_METADATA_COLS || !rawProperties || typeof rawProperties !== 'object') continue
    const column = worksheet.getColumn(index)
    const outlineLevel = Math.max(0, Math.floor(Number(rawProperties.outlineLevel) || 0))
    if (outlineLevel) column.outlineLevel = outlineLevel
    const style = normalizeStyle(rawProperties.style)
    if (style) column.style = style
    if (typeof rawProperties.numFmt === 'string' && rawProperties.numFmt) column.numFmt = rawProperties.numFmt
  }
  for (const [key, rawProperties] of Object.entries(sheet.rowProperties || {})) {
    const index = Number(key)
    if (!Number.isInteger(index) || index < 1 || index > MAX_METADATA_ROWS || !rawProperties || typeof rawProperties !== 'object') continue
    const row = worksheet.getRow(index)
    const outlineLevel = Math.max(0, Math.floor(Number(rawProperties.outlineLevel) || 0))
    if (outlineLevel) row.outlineLevel = outlineLevel
    const style = normalizeStyle(rawProperties.style)
    if (style) row.style = style
    if (typeof rawProperties.numFmt === 'string' && rawProperties.numFmt) row.numFmt = rawProperties.numFmt
  }
  for (const [key, value] of Object.entries(sheet.colWidths || {})) {
    const index = Number(key)
    const width = Number(value && typeof value === 'object' ? value.width : value)
    if (Number.isInteger(index) && index >= 1 && index <= MAX_METADATA_COLS && Number.isFinite(width) && width > 0) {
      worksheet.getColumn(index).width = Math.min(width, 255)
    }
  }
  for (const [key, value] of Object.entries(sheet.rowHeights || {})) {
    const index = Number(key)
    const height = Number(value && typeof value === 'object' ? value.height : value)
    if (Number.isInteger(index) && index >= 1 && index <= MAX_METADATA_ROWS && Number.isFinite(height) && height > 0) {
      worksheet.getRow(index).height = Math.min(height, 409)
    }
  }
  for (const rawIndex of Array.isArray(sheet.hiddenCols) ? sheet.hiddenCols : []) {
    const index = Number(rawIndex)
    if (Number.isInteger(index) && index >= 1 && index <= MAX_METADATA_COLS) worksheet.getColumn(index).hidden = true
  }
  for (const rawIndex of Array.isArray(sheet.hiddenRows) ? sheet.hiddenRows : []) {
    const index = Number(rawIndex)
    if (Number.isInteger(index) && index >= 1 && index <= MAX_METADATA_ROWS) {
      const row = worksheet.getRow(index)
      row.hidden = true
      // ExcelJS only emits an otherwise-empty row when it has a height.  A
      // default height marker keeps an empty hidden row from disappearing.
      if (!row.height) row.height = worksheet.properties.defaultRowHeight || 15
    }
  }
}

function applyExcelJSSheetViews(worksheet, sheet) {
  let views = Array.isArray(sheet.views)
    ? hydratePlainClone(sheet.views).filter((view) => view && typeof view === 'object')
    : plainClone(worksheet.views || [])
  if (!Object.prototype.hasOwnProperty.call(sheet, 'frozen')) {
    worksheet.views = views
    return
  }

  const frozen = sheet.frozen
  const rows = frozen && typeof frozen === 'object' ? Math.max(0, Math.floor(Number(frozen.rows) || 0)) : 0
  const columns = frozen && typeof frozen === 'object' ? Math.max(0, Math.floor(Number(frozen.columns) || 0)) : 0
  const frozenIndex = views.findIndex((view) => view.state === 'frozen')
  if (!rows && !columns) {
    views = views.filter((view) => view.state !== 'frozen')
    worksheet.views = views
    return
  }

  const targetIndex = frozenIndex >= 0 ? frozenIndex : 0
  const view = { ...(views[targetIndex] || {}), state: 'frozen', xSplit: columns, ySplit: rows }
  if (normalizeAddress(frozen.topLeftCell)) view.topLeftCell = normalizeAddress(frozen.topLeftCell)
  else delete view.topLeftCell
  if (normalizeAddress(frozen.activeCell)) view.activeCell = normalizeAddress(frozen.activeCell)
  else delete view.activeCell
  if (frozenIndex >= 0) views[frozenIndex] = view
  else views.unshift(view)
  worksheet.views = views
}

const DEFAULT_HEADER_FOOTER = Object.freeze({
  differentFirst: false,
  differentOddEven: false,
  oddHeader: null,
  oddFooter: null,
  evenHeader: null,
  evenFooter: null,
  firstHeader: null,
  firstFooter: null,
})

/**
 * Apply the model's sheet-level settings. `matched` is true when the worksheet was reused from
 * the source package: then the model is authoritative and a setting the model no longer has
 * (a removed filter, protection, header, tab colour, validation or conditional format) is
 * cleared instead of surviving from the source.
 */
function applyExcelJSSheetMetadata(worksheet, sheet, preserveBase, matched = false) {
  const has = (key) => Object.prototype.hasOwnProperty.call(sheet, key)
  if (sheet.properties && typeof sheet.properties === 'object') {
    const properties = hydratePlainClone(sheet.properties)
    worksheet.properties = { ...(worksheet.properties || {}), ...properties }
    // The model is authoritative for the tab colour: a removed colour must not survive
    // from the source package.
    if (properties.tabColor == null) delete worksheet.properties.tabColor
    if (properties.outlineProperties) {
      worksheet.properties.outlineProperties = {
        ...((worksheet.properties && worksheet.properties.outlineProperties) || {}),
        ...properties.outlineProperties,
      }
    }
  } else if (matched && worksheet.properties) {
    delete worksheet.properties.tabColor
  }
  if (sheet.outline && typeof sheet.outline === 'object') {
    worksheet.properties = {
      ...(worksheet.properties || {}),
      outlineProperties: {
        ...((worksheet.properties && worksheet.properties.outlineProperties) || {}),
        summaryBelow: !sheet.outline.above,
        summaryRight: !sheet.outline.left,
      },
    }
  }
  if (sheet.pageSetup && typeof sheet.pageSetup === 'object') {
    const pageSetup = hydratePlainClone(sheet.pageSetup)
    worksheet.pageSetup = { ...(worksheet.pageSetup || {}), ...pageSetup }
    // ExcelJS treats any firstPageNumber value as an instruction to restart.
    // Its importer supplies 1 even when the source flag is false, so remove
    // that dormant default before serializing an automatically numbered sheet.
    if (pageSetup.useFirstPageNumber === false) delete worksheet.pageSetup.firstPageNumber
    if (pageSetup.margins) {
      worksheet.pageSetup.margins = { ...((worksheet.pageSetup && worksheet.pageSetup.margins) || {}), ...pageSetup.margins }
    }
    // Print areas and titles come from the model only (they are rewritten as defined names).
    for (const key of ['printArea', 'printTitlesRow', 'printTitlesColumn']) if (!(key in pageSetup)) delete worksheet.pageSetup[key]
    delete worksheet.pageSetup.printAreaWhole
  } else if (matched && worksheet.pageSetup) {
    for (const key of ['printArea', 'printTitlesRow', 'printTitlesColumn']) delete worksheet.pageSetup[key]
  }
  if (sheet.headerFooter && typeof sheet.headerFooter === 'object') {
    worksheet.headerFooter = matched
      ? { ...DEFAULT_HEADER_FOOTER, ...hydratePlainClone(sheet.headerFooter) }
      : { ...(worksheet.headerFooter || {}), ...hydratePlainClone(sheet.headerFooter) }
  } else if (matched) {
    worksheet.headerFooter = { ...DEFAULT_HEADER_FOOTER }
  }
  if (Array.isArray(sheet.rowBreaks)) worksheet.rowBreaks = hydratePlainClone(sheet.rowBreaks)
  if (has('sheetProtection') || matched) {
    worksheet.sheetProtection = sheet.sheetProtection ? hydratePlainClone(sheet.sheetProtection) : null
  }
  // Validation rules are range-keyed; ExcelJS's writer needs normalized keys.
  if (sheet.dataValidations && (!preserveBase || !sheet.dataValidationsTruncated)) {
    worksheet.dataValidations.model = normalizeValidationModel(hydratePlainClone(sheet.dataValidations))
  } else if (matched && !sheet.dataValidationsTruncated) {
    worksheet.dataValidations.model = {}
  }
  if (Array.isArray(sheet.conditionalFormattings) && (!preserveBase || !sheet.conditionalFormattingsTruncated)) {
    worksheet.conditionalFormattings = hydratePlainClone(sheet.conditionalFormattings)
  } else if (matched && !sheet.conditionalFormattingsTruncated) {
    worksheet.conditionalFormattings = []
  }
  if (has('autoFilter') || matched) {
    worksheet.autoFilter = sheet.autoFilter ? autoFilterForExcelJS(sheet.autoFilter) : undefined
  }
  applyExcelJSSheetViews(worksheet, sheet)
}

function activeSheetIndex(model) {
  let index = Math.max(0, model.sheets.findIndex((sheet) => sheet.id === model.activeSheetId))
  if (model.sheets[index].state === 'hidden' || model.sheets[index].state === 'veryHidden') {
    index = model.sheets.findIndex((sheet) => sheet.state !== 'hidden' && sheet.state !== 'veryHidden')
  }
  return Math.max(0, index)
}

function applyExcelJSWorkbookViews(excelWorkbook, model) {
  const metadata = model.metadata || {}
  const sourceViews = Array.isArray(metadata.workbookViews) ? metadata.workbookViews : []
  const views = sourceViews.map((source) => {
    const view = hydratePlainClone(source)
    return {
      x: Number.isFinite(Number(view.x)) ? Number(view.x) : Number.isFinite(Number(view.xWindow)) ? Number(view.xWindow) : 0,
      y: Number.isFinite(Number(view.y)) ? Number(view.y) : Number.isFinite(Number(view.yWindow)) ? Number(view.yWindow) : 0,
      width: Number.isFinite(Number(view.width)) ? Number(view.width) : Number.isFinite(Number(view.windowWidth)) ? Number(view.windowWidth) : 12000,
      height: Number.isFinite(Number(view.height)) ? Number(view.height) : Number.isFinite(Number(view.windowHeight)) ? Number(view.windowHeight) : 24000,
      firstSheet: Number.isInteger(Number(view.firstSheet)) ? Number(view.firstSheet) : 0,
      activeTab: Number.isInteger(Number(view.activeTab)) ? Number(view.activeTab) : 0,
      visibility: typeof view.visibility === 'string' ? view.visibility : 'visible',
    }
  })
  if (!views.length) views.push({ activeTab: 0, firstSheet: 0, visibility: 'visible' })
  const activeTab = activeSheetIndex(model)
  views[0].activeTab = activeTab
  views[0].firstSheet = Math.min(Math.max(0, Number(views[0].firstSheet) || 0), Math.max(0, model.sheets.length - 1))
  excelWorkbook.views = views
}

function applyExcelJSCalcProperties(excelWorkbook, model) {
  const source = model.metadata && model.metadata.calcProperties
  const calcProperties = source && typeof source === 'object' ? hydratePlainClone(source) : {}
  excelWorkbook.calcProperties = { ...(excelWorkbook.calcProperties || {}), ...calcProperties, fullCalcOnLoad: true, forceFullCalc: true }
  if (!excelWorkbook.calcProperties.calcMode) excelWorkbook.calcProperties.calcMode = 'auto'
}

function clearExcelJSWorksheetCells(worksheet) {
  const merges = worksheet.model && Array.isArray(worksheet.model.merges) ? worksheet.model.merges.slice() : []
  for (const merge of merges) {
    try {
      worksheet.unMergeCells(merge)
    } catch {
      // A malformed source merge must not block saving the editable cell grid.
    }
  }
  worksheet._rows = []
  worksheet._columns = null
  worksheet._keys = {}
  worksheet._merges = {}
}

function matchExcelJSWorksheets(excelWorkbook, model, preserveBase) {
  if (!preserveBase) {
    return model.sheets.map((sheet) =>
      excelWorkbook.addWorksheet(sheet.name, {
        state: sheet.state === 'hidden' || sheet.state === 'veryHidden' ? sheet.state : 'visible',
      }),
    )
  }

  // A model sheet reuses a source worksheet only through the identity recorded when the
  // workbook was opened. A sheet added in the editor has none and always gets a fresh
  // worksheet: matching it by name or position would hand it the pictures, protection,
  // dropdowns, conditional formats, tables and header of a sheet the user deleted.
  const existing = excelWorkbook.worksheets.slice()
  const used = new Set()
  const identityOf = (sheet) => {
    if (Number.isInteger(sheet.sourceWorksheetId)) {
      return (candidate) => candidate.id === sheet.sourceWorksheetId
    }
    if (typeof sheet.sourceSheetName === 'string' && sheet.sourceSheetName) {
      return (candidate) => candidate.name === sheet.sourceSheetName
    }
    if (Number.isInteger(sheet.sourceSheetIndex)) {
      const candidate = existing[sheet.sourceSheetIndex]
      return (item) => item === candidate
    }
    return null
  }
  const matches = new Array(model.sheets.length).fill(undefined)
  // A duplicated sheet carries its original's identity. The sheet that still has the source
  // name keeps the source worksheet; the copy then gets a fresh one, whatever the tab order.
  const passes = [
    (sheet) => sheet.name === sheet.sourceSheetName,
    () => true,
  ]
  for (const accept of passes) {
    model.sheets.forEach((sheet, index) => {
      if (matches[index] || !accept(sheet)) return
      const matcher = identityOf(sheet)
      if (!matcher) return
      const worksheet = existing.find((candidate) => !used.has(candidate) && matcher(candidate))
      if (!worksheet) return
      used.add(worksheet)
      matches[index] = worksheet
    })
  }

  existing.forEach((worksheet, index) => {
    worksheet.name = `__sc_source_${index + 1}_${worksheet.id}`.slice(0, MAX_SHEET_NAME_LENGTH)
  })
  for (const worksheet of existing) {
    if (!used.has(worksheet)) excelWorkbook.removeWorksheet(worksheet.id)
  }
  return matches.map((worksheet, index) => {
    const modelSheet = model.sheets[index]
    const target = worksheet || excelWorkbook.addWorksheet(modelSheet.name)
    target.name = modelSheet.name
    target.state = modelSheet.state === 'hidden' || modelSheet.state === 'veryHidden' ? modelSheet.state : 'visible'
    target.orderNo = index
    if (worksheet) target.__simpleCalcMatched = true
    return target
  })
}

function xmlSafeString(value) {
  return String(value == null ? '' : value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
}

function renderXmlAttributes(attributes, allowedKeys) {
  const allowed = new Set(allowedKeys)
  const result = []
  for (const [key, rawValue] of Object.entries(attributes || {})) {
    if (!allowed.has(key) || rawValue == null || rawValue === '') continue
    const value = typeof rawValue === 'boolean' ? (rawValue ? '1' : '0') : rawValue
    result.push(`${key}="${encodeXmlAttribute(xmlSafeString(value))}"`)
  }
  return result.length ? ` ${result.join(' ')}` : ''
}

const WORKBOOK_PROPERTY_KEYS = [
  'date1904',
  'showObjects',
  'showBorderUnselectedTables',
  'filterPrivacy',
  'promptedSolutions',
  'showInkAnnotation',
  'backupFile',
  'saveExternalLinkValues',
  'updateLinks',
  'codeName',
  'hidePivotFieldList',
  'showPivotChartFilter',
  'allowRefreshQuery',
  'publishItems',
  'checkCompatibility',
  'autoCompressPictures',
  'refreshAllConnections',
  'defaultThemeVersion',
]

const CALC_PROPERTY_KEYS = [
  'calcId',
  'calcMode',
  'fullCalcOnLoad',
  'forceFullCalc',
  'calcCompleted',
  'calcOnSave',
  'concurrentCalc',
  'concurrentManualCount',
  'fullPrecision',
  'iterate',
  'iterateCount',
  'iterateDelta',
  'refMode',
]

const WORKBOOK_VIEW_KEYS = [
  'visibility',
  'minimized',
  'showHorizontalScroll',
  'showVerticalScroll',
  'showSheetTabs',
  'xWindow',
  'yWindow',
  'windowWidth',
  'windowHeight',
  'tabRatio',
  'firstSheet',
  'activeTab',
  'autoFilterDateGrouping',
]

const DEFINED_NAME_ATTRIBUTE_KEYS = [
  'name',
  'comment',
  'customMenu',
  'description',
  'help',
  'statusBar',
  'localSheetId',
  'hidden',
  'function',
  'functionGroupId',
  'shortcutKey',
  'publishToServer',
  'workbookParameter',
  'vbProcedure',
  'xlm',
]

/**
 * Absolute references for a sheet's print area. An area still equal to the bounded form of a
 * whole-column/row area that was opened ($A:$F, $1:$20) is written back in that form; an area
 * that cannot be read is skipped (with a note) rather than blocking the save.
 */
function printAreaReferences(sheet, prefix, warnings) {
  const setup = sheet.pageSetup || {}
  const whole = setup.printAreaWhole && typeof setup.printAreaWhole === 'object' ? setup.printAreaWhole : {}
  const references = []
  for (const raw of splitAreas(String(setup.printArea || '').replace(/&&/g, ','))) {
    const local = raw.replace(/^.*!/, '').trim()
    const range = normalizeRange(local)
    if (range) {
      const original = typeof whole[range] === 'string' && /^\$[A-Z]{1,3}:\$[A-Z]{1,3}$|^\$\d{1,7}:\$\d{1,7}$/.test(whole[range]) ? whole[range] : null
      references.push(prefix + (original || range.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')))
      continue
    }
    const columns = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/i.exec(local)
    const rows = /^\$?(\d{1,7}):\$?(\d{1,7})$/.exec(local)
    if (columns) references.push(`${prefix}$${columns[1].toUpperCase()}:$${columns[2].toUpperCase()}`)
    else if (rows && Number(rows[1]) >= 1 && Number(rows[2]) >= 1) references.push(`${prefix}$${rows[1]}:$${rows[2]}`)
    else if (Array.isArray(warnings)) warnings.push(`The print area "${raw}" of sheet "${sheet.name}" is not a cell range, so it was not saved.`)
  }
  return references
}

function renderDefinedNamesBlock(model, warnings) {
  const source = Array.isArray(model.definedNames)
    ? model.definedNames
    : model.metadata && Array.isArray(model.metadata.definedNames)
      ? model.metadata.definedNames
      : undefined
  const items = []
  for (const item of (source || []).slice(0, MAX_ADVANCED_METADATA_ENTRIES)) {
    if (!item || typeof item.name !== 'string' || !item.name) continue
    // Page setup is authoritative for these built-ins. Keeping imported names
    // here would silently restore old print areas after a layout change.
    if (item.name === '_xlnm.Print_Area' || item.name === '_xlnm.Print_Titles') continue
    const ranges = Array.isArray(item.ranges) ? item.ranges : typeof item.ranges === 'string' ? [item.ranges] : []
    const ref = typeof item.ref === 'string' && item.ref ? item.ref : ranges.filter((range) => typeof range === 'string').join(',')
    if (!ref || ref.length > 65_536) continue
    const attributes = { ...(item.attributes && typeof item.attributes === 'object' ? item.attributes : {}), name: item.name }
    const referencedSheetIndex =
      typeof item.localSheetRefId === 'string' ? model.sheets.findIndex((sheet) => sheet.id === item.localSheetRefId) : -1
    if (typeof item.localSheetRefId === 'string' && referencedSheetIndex < 0) continue
    const localSheetIndex = referencedSheetIndex >= 0
      ? referencedSheetIndex
      : Number.isInteger(item.localSheetIndex)
        ? item.localSheetIndex
        : item.localSheetId
    if (Number.isInteger(localSheetIndex) && localSheetIndex >= 0 && localSheetIndex < model.sheets.length) {
      attributes.localSheetId = localSheetIndex
    }
    if (item.hidden != null) attributes.hidden = Boolean(item.hidden)
    if (typeof item.comment === 'string' && item.comment) attributes.comment = item.comment
    items.push(`<definedName${renderXmlAttributes(attributes, DEFINED_NAME_ATTRIBUTE_KEYS)}>${encodeXmlText(xmlSafeString(ref))}</definedName>`)
  }
  model.sheets.forEach((sheet, localSheetId) => {
    const setup = sheet.pageSetup || {}
    const prefix = `'${String(sheet.name).replace(/'/g, "''")}'!`
    const add = (name, references) => {
      if (references.length) items.push(`<definedName name="${name}" localSheetId="${localSheetId}">${encodeXmlText(references.join(','))}</definedName>`)
    }
    // ExcelJS emits $A1:$B2: some native readers consequently ignore the area.
    // A print area is an absolute range, including both row and column anchors.
    add('_xlnm.Print_Area', printAreaReferences(sheet, prefix, warnings))
    const titles = []
    const rows = /^\$?(\d+):\$?(\d+)$/.exec(String(setup.printTitlesRow || ''))
    const cols = /^\$?([A-Z]+):\$?([A-Z]+)$/i.exec(String(setup.printTitlesColumn || ''))
    if (rows) titles.push(`${prefix}$${rows[1]}:$${rows[2]}`)
    if (cols) titles.push(`${prefix}$${cols[1].toUpperCase()}:$${cols[2].toUpperCase()}`)
    add('_xlnm.Print_Titles', titles)
  })
  return items.length ? `<definedNames>${items.join('')}</definedNames>` : ''
}

function replaceOrInsertWorkbookTag(xml, tagName, renderedTag) {
  const pattern = new RegExp(`<${tagName}\\b[^>]*(?:\\/\\s*>|>[\\s\\S]*?<\\/${tagName}\\s*>)`, 'i')
  if (pattern.test(xml)) return xml.replace(pattern, renderedTag)
  return xml.replace(/<\/workbook\s*>/i, `${renderedTag}</workbook>`)
}

function renderCustomPropertiesXml(properties) {
  if (!properties || typeof properties !== 'object' || Array.isArray(properties)) return undefined
  const items = []
  let propertyId = 2
  for (const [rawName, rawValue] of Object.entries(properties).slice(0, 1_000)) {
    if (rawValue === undefined || typeof rawValue === 'function' || typeof rawValue === 'symbol') continue
    const name = xmlSafeString(rawName)
    if (!name) continue
    let tag = 'vt:lpwstr'
    let value = rawValue
    if (typeof rawValue === 'boolean') tag = 'vt:bool'
    else if (typeof rawValue === 'number' && Number.isFinite(rawValue)) {
      tag = Number.isInteger(rawValue) && rawValue >= -2_147_483_648 && rawValue <= 2_147_483_647 ? 'vt:i4' : 'vt:r8'
    } else if (rawValue instanceof Date || (rawValue && rawValue.type === 'date')) {
      const date = rawValue instanceof Date ? rawValue : new Date(rawValue.value)
      if (!Number.isNaN(date.getTime())) {
        tag = 'vt:filetime'
        value = date.toISOString()
      }
    } else if (rawValue && typeof rawValue === 'object') {
      value = JSON.stringify(rawValue)
    }
    items.push(
      `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${propertyId}" name="${encodeXmlAttribute(name)}"><${tag}>${encodeXmlText(xmlSafeString(value))}</${tag}></property>`,
    )
    propertyId += 1
  }
  if (!items.length) return undefined
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${items.join('')}</Properties>`
}

async function customPropertiesXmlFromBase(baseBuffer) {
  if (!baseBuffer) return undefined
  const baseZip = await JSZip.loadAsync(baseBuffer)
  const entry = baseZip.file('docProps/custom.xml')
  if (!entry) return undefined
  const expectedSize = entry._data && Number(entry._data.uncompressedSize)
  if (Number.isFinite(expectedSize) && expectedSize > MAX_WORKBOOK_XML_BYTES) return undefined
  const data = await entry.async('nodebuffer')
  return data.length <= MAX_WORKBOOK_XML_BYTES ? data : undefined
}

async function addCustomPropertiesPart(zip, customXml) {
  if (!customXml) return
  zip.file('docProps/custom.xml', customXml)

  const contentTypesEntry = zip.file('[Content_Types].xml')
  if (contentTypesEntry) {
    let contentTypes = await contentTypesEntry.async('string')
    if (!/PartName=["']\/docProps\/custom\.xml["']/i.test(contentTypes)) {
      contentTypes = contentTypes.replace(
        /<\/Types\s*>/i,
        '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>',
      )
      zip.file('[Content_Types].xml', contentTypes)
    }
  }

  const relationshipsEntry = zip.file('_rels/.rels')
  if (relationshipsEntry) {
    let relationships = await relationshipsEntry.async('string')
    if (!/relationships\/custom-properties["']/i.test(relationships)) {
      const usedIds = new Set([...relationships.matchAll(/\bId=["']([^"']+)["']/gi)].map((match) => match[1]))
      let suffix = 1
      while (usedIds.has(`rIdSimpleCalcCustom${suffix}`)) suffix += 1
      relationships = relationships.replace(
        /<\/Relationships\s*>/i,
        `<Relationship Id="rIdSimpleCalcCustom${suffix}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>`,
      )
      zip.file('_rels/.rels', relationships)
    }
  }
}

const HYPERLINK_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const AFTER_HYPERLINKS = new Set(['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'])

/** Insert `elements` into the worksheet's <hyperlinks>, creating it at its schema position. */
function insertHyperlinkElements(sheetXml, elements) {
  const prefixMatch = /<([A-Za-z0-9_]+:)?worksheet\b/.exec(sheetXml)
  const prefix = (prefixMatch && prefixMatch[1]) || ''
  const existing = new RegExp(`<${prefix}hyperlinks\\b[^>]*>([\\s\\S]*?)</${prefix}hyperlinks>`).exec(sheetXml)
  if (existing) {
    const close = existing.index + existing[0].lastIndexOf(`</${prefix}hyperlinks>`)
    return sheetXml.slice(0, close) + elements.join('') + sheetXml.slice(close)
  }
  const block = `<${prefix}hyperlinks>${elements.join('')}</${prefix}hyperlinks>`
  let scanFrom = sheetXml.lastIndexOf(`</${prefix}sheetData>`)
  if (scanFrom < 0) {
    const empty = sheetXml.search(new RegExp(`<${prefix}sheetData\\s*/>`))
    scanFrom = empty >= 0 ? empty : 0
  }
  const tagRe = /<(\/?)([A-Za-z0-9_]+:)?([A-Za-z0-9_]+)\b[^>]*?(\/?)>/g
  tagRe.lastIndex = scanFrom
  let depth = 0
  let match
  while ((match = tagRe.exec(sheetXml))) {
    const [, closing, , local, selfClosing] = match
    if (closing) {
      if (depth === 0) {
        if (local === 'worksheet') return sheetXml.slice(0, match.index) + block + sheetXml.slice(match.index)
        continue
      }
      depth -= 1
      continue
    }
    if (depth === 0 && AFTER_HYPERLINKS.has(local)) return sheetXml.slice(0, match.index) + block + sheetXml.slice(match.index)
    if (!selfClosing) depth += 1
  }
  return sheetXml.replace(new RegExp(`</${prefix}worksheet>\\s*$`), `${block}</${prefix}worksheet>`)
}

/**
 * Hyperlinks ExcelJS cannot write: links to a place in the workbook (written as `location`,
 * without a relationship) and links on formula cells.
 */
async function writeHyperlinksToPackage(zip, linksBySheet) {
  if (!linksBySheet || !linksBySheet.size) return
  const parts = await workbookSheetParts(zip)
  for (const [sheetName, links] of linksBySheet) {
    const part = parts.find((item) => item.name === sheetName)
    const entry = part && zip.file(part.part)
    if (!entry || !links.length) continue
    let sheetXml = await entry.async('string')
    const relationships = await readRelationships(zip, part.part)
    const used = new Set(relationships.map((rel) => rel.id))
    let counter = 1
    const elements = []
    for (const link of links) {
      const tooltip = link.tooltip ? ` tooltip="${encodeXmlAttribute(xmlSafeString(link.tooltip))}"` : ''
      if (link.target.startsWith('#')) {
        const location = link.target.slice(1)
        if (!location) continue
        elements.push(`<hyperlink ref="${link.address}" location="${encodeXmlAttribute(xmlSafeString(location))}"${tooltip}/>`)
      } else {
        while (used.has(`rIdLink${counter}`)) counter += 1
        const id = `rIdLink${counter}`
        used.add(id)
        relationships.push({ id, type: HYPERLINK_RELATIONSHIP, target: xmlSafeString(link.target), external: true })
        elements.push(`<hyperlink ref="${link.address}" r:id="${id}"${tooltip}/>`)
      }
    }
    if (!elements.length) continue
    if (!/\sxmlns:r=/.test(/<([A-Za-z0-9_]+:)?worksheet\b[^>]*>/.exec(sheetXml)?.[0] || '')) {
      sheetXml = sheetXml.replace(/<([A-Za-z0-9_]+:)?worksheet\b/, (tag) => `${tag} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
    }
    zip.file(part.part, insertHyperlinkElements(sheetXml, elements))
    zip.file(relsPathFor(part.part), renderRelationships(relationships))
  }
}

/** Re-insert <fileSharing> and <workbookProtection> in CT_Workbook order. */
function insertWorkbookProtection(xml, metadata) {
  let output = xml.replace(/<fileSharing\b[^>]*?(?:\/>|>[\s\S]*?<\/fileSharing>)/gi, '').replace(/<workbookProtection\b[^>]*?(?:\/>|>[\s\S]*?<\/workbookProtection>)/gi, '')
  const render = (tag, attributes) => {
    const parts = Object.entries(attributes || {})
      .filter(([key, value]) => /^[A-Za-z_:][A-Za-z0-9_.:-]*$/.test(key) && value != null && value !== '')
      .map(([key, value]) => `${key}="${encodeXmlAttribute(xmlSafeString(typeof value === 'boolean' ? (value ? '1' : '0') : value))}"`)
    return parts.length ? `<${tag} ${parts.join(' ')}/>` : ''
  }
  const fileSharing = metadata.fileSharing && typeof metadata.fileSharing === 'object' ? render('fileSharing', metadata.fileSharing) : ''
  const protection = metadata.workbookProtection && typeof metadata.workbookProtection === 'object' ? render('workbookProtection', metadata.workbookProtection) : ''
  if (fileSharing) {
    // fileVersion?, fileSharing?, workbookPr?, ...
    if (/<fileVersion\b[^>]*?(?:\/>|>[\s\S]*?<\/fileVersion>)/i.test(output)) output = output.replace(/(<fileVersion\b[^>]*?(?:\/>|>[\s\S]*?<\/fileVersion>))/i, `$1${fileSharing}`)
    else if (/<workbookPr\b/i.test(output)) output = output.replace(/<workbookPr\b/i, `${fileSharing}<workbookPr`)
    else output = output.replace(/<bookViews\b/i, `${fileSharing}<bookViews`)
  }
  if (protection) {
    // ..., workbookPr?, workbookProtection?, bookViews?, sheets
    if (/<bookViews\b/i.test(output)) output = output.replace(/<bookViews\b/i, `${protection}<bookViews`)
    else output = output.replace(/<sheets\b/i, `${protection}<sheets`)
  }
  return output
}

async function patchXlsxWorkbookMetadata(buffer, model, baseBuffer, context = {}) {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file('xl/workbook.xml')
  if (!entry) return buffer
  const expectedSize = entry._data && Number(entry._data.uncompressedSize)
  if (Number.isFinite(expectedSize) && expectedSize > MAX_WORKBOOK_XML_BYTES) return buffer
  const xmlBuffer = await entry.async('nodebuffer')
  if (xmlBuffer.length > MAX_WORKBOOK_XML_BYTES) return buffer
  let xml = xmlBuffer.toString('utf8')
  const metadata = model.metadata || {}

  const currentWorkbookPropertiesMatch = /<workbookPr\b([^>]*?)(?:\/?\s*>)/i.exec(xml)
  const workbookProperties = {
    ...(metadata.workbookProperties && typeof metadata.workbookProperties === 'object' ? metadata.workbookProperties : {}),
    ...(currentWorkbookPropertiesMatch ? parseXmlAttributes(currentWorkbookPropertiesMatch[1]) : {}),
    date1904: Boolean(metadata.date1904),
  }
  xml = replaceOrInsertWorkbookTag(xml, 'workbookPr', `<workbookPr${renderXmlAttributes(workbookProperties, WORKBOOK_PROPERTY_KEYS)}/>`)

  const currentCalcMatch = /<calcPr\b([^>]*?)(?:\/?\s*>)/i.exec(xml)
  const calcProperties = {
    ...(currentCalcMatch ? parseXmlAttributes(currentCalcMatch[1]) : {}),
    ...(metadata.calcProperties && typeof metadata.calcProperties === 'object' ? metadata.calcProperties : {}),
    fullCalcOnLoad: true,
    forceFullCalc: true,
  }
  if (!calcProperties.calcMode) calcProperties.calcMode = 'auto'
  xml = replaceOrInsertWorkbookTag(xml, 'calcPr', `<calcPr${renderXmlAttributes(calcProperties, CALC_PROPERTY_KEYS)}/>`)

  const sourceViews = Array.isArray(metadata.workbookViews) ? metadata.workbookViews : []
  let viewIndex = 0
  xml = xml.replace(/<workbookView\b([^>]*?)\/\s*>/gi, (_tag, rawAttributes) => {
    const current = parseXmlAttributes(rawAttributes)
    const source = sourceViews[viewIndex] && typeof sourceViews[viewIndex] === 'object' ? sourceViews[viewIndex] : {}
    viewIndex += 1
    return `<workbookView${renderXmlAttributes({ ...source, ...current }, WORKBOOK_VIEW_KEYS)}/>`
  })

  const namesBlock = renderDefinedNamesBlock(model, context.warnings)
  if (namesBlock !== undefined) {
    const namesPattern = /<definedNames\b[^>]*>[\s\S]*?<\/definedNames\s*>/i
    if (namesPattern.test(xml)) xml = xml.replace(namesPattern, namesBlock)
    else if (namesBlock) xml = xml.replace(/<calcPr\b/i, `${namesBlock}<calcPr`)
  }
  xml = insertWorkbookProtection(xml, metadata)

  zip.file('xl/workbook.xml', xml)
  await writeHyperlinksToPackage(zip, context.links)
  // Column widths are defined by the Normal style font, independently of each
  // cell's font. Keep it separate from font0/cellXfs so explicitly styled
  // Calibri cells do not change when a legacy template uses Arial as Normal.
  const normalFont = normalizeFont(model.metadata?.normalFont)
  const stylesEntry = zip.file('xl/styles.xml')
  if (normalFont?.name && normalFont.size > 0 && stylesEntry) {
    let styles = await stylesEntry.async('string')
    const fontsMatch = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(styles)
    const normalStyle = /<cellStyle\b[^>]*\bbuiltinId="0"[^>]*\/>/.exec(styles)
    const normalIndex = normalStyle ? Number(parseXmlAttributes(normalStyle[0]).xfId || 0) : 0
    if (fontsMatch && /<cellStyleXfs\b/.test(styles)) {
      const fontCount = (fontsMatch[1].match(/<font\b/g) || []).length
      const fontXml = `<font><name val="${encodeXmlAttribute(normalFont.name)}"/><sz val="${Number(normalFont.size)}"/>${normalFont.bold ? '<b/>' : ''}${normalFont.italic ? '<i/>' : ''}${normalFont.family ? `<family val="${Number(normalFont.family)}"/>` : ''}</font>`
      styles = styles.replace(fontsMatch[0], `<fonts count="${fontCount + 1}">${fontsMatch[1]}${fontXml}</fonts>`)
      styles = styles.replace(/(<cellStyleXfs\b[^>]*>)([\s\S]*?)(<\/cellStyleXfs>)/, (_all, open, content, close) => {
        let index = 0
        const patched = content.replace(/<xf\b[^>]*>/g, (tag) => {
          if (index++ !== normalIndex) return tag
          return /\bfontId="[^"]*"/.test(tag) ? tag.replace(/\bfontId="[^"]*"/, `fontId="${fontCount}"`) : tag.replace(/\/?\s*>$/, ` fontId="${fontCount}"/>`)
        })
        return open + patched + close
      })
      zip.file('xl/styles.xml', styles)
    }
  }
  const customProperties = model.metadata && model.metadata.properties && model.metadata.properties.custom
  const customXml = (await customPropertiesXmlFromBase(baseBuffer)) || renderCustomPropertiesXml(customProperties)
  await addCustomPropertiesPart(zip, customXml)
  // Charts: ExcelJS drops them, so write every model chart (drawing, chart parts,
  // rels, content types) into the package it produced; unmodified imported charts
  // are copied byte-for-byte from the source package.
  await writeChartsToPackage(zip, model, baseBuffer, context.warnings)
  await writePivotPackage(zip, model)
  await writeSparklinesToPackage(zip, model)
  return Buffer.from(
    await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    }),
  )
}

function applyDefinedNamesToExcelJS(excelWorkbook, definedNames) {
  if (!Array.isArray(definedNames)) return
  const safeNames = []
  for (const item of definedNames) {
    if (!item || typeof item.name !== 'string' || !item.name) continue
    const sourceRanges = Array.isArray(item.ranges) ? item.ranges : typeof item.ranges === 'string' ? [item.ranges] : []
    const ranges = sourceRanges.filter((range) => typeof range === 'string' && range.includes('!') && range.length <= 512)
    if (ranges.length) safeNames.push({ name: item.name, ranges })
  }
  try {
    excelWorkbook.definedNames.model = safeNames
  } catch {
    // Defined names are auxiliary; an invalid imported name must not prevent
    // users from saving the actual workbook cells and formulas.
  }
}

function mergedOuterBorder(sheet, decoded) {
  const addressAt = (row, col) => XLSX.utils.encode_cell({ r: row, c: col })
  const masterAddress = addressAt(decoded.s.r, decoded.s.c)
  const masterBorder = normalizeBorder(sheet.cells && sheet.cells[masterAddress] && sheet.cells[masterAddress].style && sheet.cells[masterAddress].style.border) || {}
  const border = { ...masterBorder }
  const firstSide = (side, positions) => {
    for (const [row, col] of positions) {
      const address = addressAt(row, col)
      const candidate = normalizeBorderSide(
        sheet.cells && sheet.cells[address] && sheet.cells[address].style && sheet.cells[address].style.border && sheet.cells[address].style.border[side],
      )
      if (candidate) return candidate
    }
    return undefined
  }
  const columns = Array.from({ length: decoded.e.c - decoded.s.c + 1 }, (_, offset) => decoded.s.c + offset)
  const rows = Array.from({ length: decoded.e.r - decoded.s.r + 1 }, (_, offset) => decoded.s.r + offset)
  const top = firstSide('top', columns.map((col) => [decoded.s.r, col]))
  const bottom = firstSide('bottom', columns.map((col) => [decoded.e.r, col]))
  const left = firstSide('left', rows.map((row) => [row, decoded.s.c]))
  const right = firstSide('right', rows.map((row) => [row, decoded.e.c]))
  if (top) border.top = top
  if (bottom) border.bottom = bottom
  if (left) border.left = left
  if (right) border.right = right
  return Object.keys(border).length ? border : undefined
}

function writeModelSheetToExcelJS(worksheet, sheet, preserveBase, pendingLinks) {
  const matched = Boolean(preserveBase && worksheet.__simpleCalcMatched)
  const packageOwned = Array.isArray(sheet.requiresSourcePackage) ? sheet.requiresSourcePackage : []
  clearExcelJSWorksheetCells(worksheet)
  worksheet.state = sheet.state === 'hidden' || sheet.state === 'veryHidden' ? sheet.state : 'visible'
  applyExcelJSSheetMetadata(worksheet, sheet, preserveBase, matched)
  applyExcelJSSheetDimensions(worksheet, sheet)
  applyOutlineToWorksheet(worksheet, sheet)
  // On a reused source worksheet the model is authoritative: no list means none left.
  applyTablesToWorksheet(worksheet, Array.isArray(sheet.tables) ? sheet.tables : matched ? [] : undefined)
  applyImagesToWorksheet(worksheet, sheet, preserveBase, { clearMissing: matched && !packageOwned.includes('images') })

  const entries = Object.entries(sheet.cells || {}).sort((left, right) => {
    const a = addressPosition(left[0])
    const b = addressPosition(right[0])
    if (!a || !b) return left[0].localeCompare(right[0])
    return a.row - b.row || a.col - b.col
  })
  for (const [rawAddress, modelCell] of entries) {
    const position = addressPosition(rawAddress)
    if (!position) throw new Error(`Invalid cell address in sheet "${sheet.name}": ${rawAddress}`)
    if (position.row > MAX_METADATA_ROWS || position.col > MAX_METADATA_COLS) {
      throw new RangeError(`Cell ${position.address} in sheet "${sheet.name}" is outside XLSX limits.`)
    }
    if (!modelCell || typeof modelCell !== 'object') continue
    setExcelJSCell(worksheet.getCell(position.address), modelCell, pendingLinks)
  }

  for (const rawMerge of Array.isArray(sheet.merges) ? sheet.merges : []) {
    const merge = normalizeRange(rawMerge)
    if (!merge) throw new Error(`Invalid merged range in sheet "${sheet.name}": ${rawMerge}`)
    const decoded = XLSX.utils.decode_range(merge)
    if (decoded.e.r + 1 > MAX_METADATA_ROWS || decoded.e.c + 1 > MAX_METADATA_COLS) {
      throw new RangeError(`Merged range ${merge} in sheet "${sheet.name}" is outside XLSX limits.`)
    }
    const border = mergedOuterBorder(sheet, decoded)
    if (border) worksheet.getCell(decoded.s.r + 1, decoded.s.c + 1).border = border
    worksheet.mergeCells(merge)
  }
}

async function excelJSWorkbookForSerialization(options) {
  if (!options || options.baseBytes == null) return { excelWorkbook: new ExcelJS.Workbook(), preserveBase: false, baseBuffer: undefined }
  const baseBuffer = bytesToBuffer(options.baseBytes)
  const excelWorkbook = new ExcelJS.Workbook()
  try {
    await loadExcelJSWithExactFonts(excelWorkbook, baseBuffer, await workbookXmlFromOoxml(baseBuffer))
  } catch (error) {
    const wrapped = new Error(`Unable to use the source workbook for fidelity-preserving save: ${error.message || 'invalid OOXML package'}`)
    wrapped.code = 'BASE_WORKBOOK_LOAD_FAILED'
    wrapped.cause = error
    throw wrapped
  }
  // A package without worksheets (or not a workbook at all) cannot be a merge base.
  if (!excelWorkbook.worksheets.length || !excelWorkbook.properties) {
    return { excelWorkbook: new ExcelJS.Workbook(), preserveBase: false, baseBuffer: undefined }
  }
  return { excelWorkbook, preserveBase: true, baseBuffer }
}

async function serializeXlsx(model, options = {}) {
  const { excelWorkbook, preserveBase, baseBuffer } = await excelJSWorkbookForSerialization(options)
  applyWorkbookProperties(excelWorkbook, model.metadata && (model.metadata.properties || model.metadata))
  excelWorkbook.properties.date1904 = Boolean(model.metadata && model.metadata.date1904)
  applyExcelJSCalcProperties(excelWorkbook, model)

  const worksheets = matchExcelJSWorksheets(excelWorkbook, model, preserveBase)
  const links = new Map()
  model.sheets.forEach((sheet, index) => {
    const pending = []
    writeModelSheetToExcelJS(worksheets[index], sheet, preserveBase, pending)
    if (pending.length) links.set(sheet.name, pending)
  })
  // Media the source package carried but no sheet shows any more (deleted pictures) are
  // dropped instead of being written back invisibly.
  compactWorkbookMedia(excelWorkbook)

  applyExcelJSWorkbookViews(excelWorkbook, model)
  applyDefinedNamesToExcelJS(excelWorkbook, model.definedNames || (model.metadata && model.metadata.definedNames))
  const result = await excelWorkbook.xlsx.writeBuffer({ useStyles: true, useSharedStrings: true })
  return patchXlsxWorkbookMetadata(Buffer.from(result), model, baseBuffer, { links, warnings: options.warnings })
}

function modelValueToSheetJS(value, valueType) {
  if (valueType === 'error') {
    const text = String(value || '#VALUE!')
    return { t: 'e', v: ERROR_TEXT_TO_CODE[text] == null ? 15 : ERROR_TEXT_TO_CODE[text], w: text }
  }
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    if (value.type === 'date') {
      const date = new Date(value.value)
      return Number.isNaN(date.getTime()) ? { t: 's', v: String(value.value || '') } : { t: 'd', v: date }
    }
    if (value.type === 'error') {
      const text = String(value.value || '#VALUE!')
      return { t: 'e', v: ERROR_TEXT_TO_CODE[text] == null ? 15 : ERROR_TEXT_TO_CODE[text], w: text }
    }
    if (value.type === 'richText') return { t: 's', v: String(richTextToPlainText(value)) }
    if (value.type === 'binary') return { t: 's', v: '[binary]' }
  }
  if (typeof value === 'number' && Number.isFinite(value)) return { t: 'n', v: value }
  if (typeof value === 'boolean') return { t: 'b', v: value }
  if (value == null) return { t: 'z', v: undefined }
  return { t: 's', v: String(value) }
}

function styleToSheetJS(style) {
  const normalized = normalizeStyle(style)
  if (!normalized) return undefined
  const result = plainClone(normalized)
  if (result.font && result.font.size != null) {
    result.font.sz = result.font.size
    delete result.font.size
  }
  const convertColor = (color) => {
    if (!color || !color.argb) return color
    const converted = { ...color, rgb: color.argb.length === 8 ? color.argb.slice(2) : color.argb }
    delete converted.argb
    return converted
  }
  if (result.font && result.font.color) result.font.color = convertColor(result.font.color)
  if (result.fill) {
    if (result.fill.pattern) result.fill.patternType = result.fill.pattern
    if (result.fill.fgColor) result.fill.fgColor = convertColor(result.fill.fgColor)
    if (result.fill.bgColor) result.fill.bgColor = convertColor(result.fill.bgColor)
  }
  return result
}

function noteToSheetJS(note) {
  if (typeof note === 'string') return [{ a: 'simple_calc', t: note }]
  if (!note || typeof note !== 'object') return undefined
  if (Array.isArray(note.comments)) {
    const comments = note.comments.map((comment) => ({
      a: String(comment.author || 'simple_calc'),
      t: String(comment.text || ''),
    }))
    if (note.hidden === true || note.comments.some((comment) => comment && comment.hidden === true)) comments.hidden = true
    return comments
  }
  if (Array.isArray(note.texts)) return [{ a: 'simple_calc', t: note.texts.map((item) => item.text || '').join('') }]
  if (typeof note.text === 'string') return [{ a: 'simple_calc', t: note.text }]
  return [{ a: 'simple_calc', t: JSON.stringify(note) }]
}

function modelCellToSheetJS(modelCell, targetFormat) {
  const formula = normalizeFormula(modelCell && modelCell.formula)
  const base = modelValueToSheetJS(
    formula ? modelCell.result : modelCell && modelCell.value,
    formula ? modelCell.resultType : modelCell && modelCell.type,
  )
  const cell = { t: base.t }
  if (base.v !== undefined) cell.v = base.v
  if (base.w !== undefined) cell.w = base.w
  // The SheetJS ODS writer omits error-typed cells entirely.  Emitting their
  // visible error token as a string keeps the cell (and any formula) intact.
  if (targetFormat === 'ods' && cell.t === 'e') {
    cell.t = 's'
    cell.v = String(base.w || (modelCell && (modelCell.result || modelCell.value)) || '#VALUE!')
    delete cell.w
  }
  if (formula) cell.f = formula
  if (formula && typeof modelCell.formulaRange === 'string' && normalizeRange(modelCell.formulaRange)) {
    cell.F = normalizeRange(modelCell.formulaRange)
  }
  if (formula && modelCell.dynamicFormula === true) cell.D = true
  if (modelCell && typeof modelCell.numFmt === 'string' && modelCell.numFmt) cell.z = modelCell.numFmt
  const style = styleToSheetJS(modelCell && modelCell.style)
  if (style) cell.s = style
  const hyperlink = normalizeHyperlink(modelCell && modelCell.hyperlink)
  if (hyperlink) {
    cell.l = { Target: hyperlink.target }
    const tooltip = hyperlink.tooltip || (modelCell && modelCell.hyperlinkTooltip)
    if (tooltip) cell.l.Tooltip = tooltip
  }
  const comments = noteToSheetJS(modelCell && modelCell.note)
  if (comments) cell.c = comments
  // The SheetJS ODS writer emits an empty cell for the blank type, taking any formula,
  // note or link attached to it along.  Workbooks written without cached results — and
  // every formula the live engine evaluates to blank — land here, so give those cells an
  // empty string body that the writer can carry.
  if (targetFormat === 'ods' && cell.t === 'z' && (cell.f || cell.c || cell.l)) {
    cell.t = 's'
    cell.v = String((modelCell && modelCell.display) || '')
  }
  return cell
}

function sheetToSheetJS(sheet, targetFormat) {
  const result = {}
  let maxRow = 1
  let maxCol = 1
  for (const [rawAddress, modelCell] of Object.entries(sheet.cells || {})) {
    const position = addressPosition(rawAddress)
    if (!position) throw new Error(`Invalid cell address in sheet "${sheet.name}": ${rawAddress}`)
    if (position.row > MAX_METADATA_ROWS || position.col > MAX_METADATA_COLS) {
      throw new RangeError(`Cell ${position.address} in sheet "${sheet.name}" is outside modern spreadsheet limits.`)
    }
    result[position.address] = modelCellToSheetJS(modelCell, targetFormat)
    maxRow = Math.max(maxRow, position.row)
    maxCol = Math.max(maxCol, position.col)
  }

  const merges = []
  for (const rawMerge of Array.isArray(sheet.merges) ? sheet.merges : []) {
    const merge = normalizeRange(rawMerge)
    if (!merge) continue
    const decoded = XLSX.utils.decode_range(merge)
    merges.push(decoded)
    maxRow = Math.max(maxRow, decoded.e.r + 1)
    maxCol = Math.max(maxCol, decoded.e.c + 1)
  }
  if (merges.length) result['!merges'] = merges

  const cols = []
  for (const [key, rawWidth] of Object.entries(sheet.colWidths || {})) {
    const index = Number(key)
    const width = Number(rawWidth && typeof rawWidth === 'object' ? rawWidth.width : rawWidth)
    if (Number.isInteger(index) && index >= 1 && index <= MAX_METADATA_COLS && Number.isFinite(width) && width > 0) {
      cols[index - 1] = { wch: Math.min(width, 255) }
    }
  }
  for (const rawIndex of Array.isArray(sheet.hiddenCols) ? sheet.hiddenCols : []) {
    const index = Number(rawIndex)
    if (!Number.isInteger(index) || index < 1 || index > MAX_METADATA_COLS) continue
    cols[index - 1] = { ...(cols[index - 1] || {}), hidden: true }
  }
  for (const [key, rawProperties] of Object.entries(sheet.columnProperties || {})) {
    const index = Number(key)
    if (!Number.isInteger(index) || index < 1 || index > MAX_METADATA_COLS || !rawProperties || typeof rawProperties !== 'object') continue
    const column = { ...(cols[index - 1] || {}) }
    const outlineLevel = Math.max(0, Math.floor(Number(rawProperties.outlineLevel) || 0))
    if (outlineLevel) column.level = outlineLevel
    if (rawProperties.collapsed === true) column.collapsed = true
    if (rawProperties.bestFit === true) column.bestFit = true
    const style = styleToSheetJS(rawProperties.style)
    if (style) column.s = style
    cols[index - 1] = column
  }
  if (cols.some(Boolean)) result['!cols'] = cols

  const rows = []
  for (const [key, rawHeight] of Object.entries(sheet.rowHeights || {})) {
    const index = Number(key)
    const height = Number(rawHeight && typeof rawHeight === 'object' ? rawHeight.height : rawHeight)
    if (Number.isInteger(index) && index >= 1 && index <= MAX_METADATA_ROWS && Number.isFinite(height) && height > 0) {
      rows[index - 1] = { hpt: Math.min(height, 409) }
    }
  }
  for (const rawIndex of Array.isArray(sheet.hiddenRows) ? sheet.hiddenRows : []) {
    const index = Number(rawIndex)
    if (!Number.isInteger(index) || index < 1 || index > MAX_METADATA_ROWS) continue
    rows[index - 1] = { ...(rows[index - 1] || {}), hidden: true }
  }
  for (const [key, rawProperties] of Object.entries(sheet.rowProperties || {})) {
    const index = Number(key)
    if (!Number.isInteger(index) || index < 1 || index > MAX_METADATA_ROWS || !rawProperties || typeof rawProperties !== 'object') continue
    const row = { ...(rows[index - 1] || {}) }
    const outlineLevel = Math.max(0, Math.floor(Number(rawProperties.outlineLevel) || 0))
    if (outlineLevel) row.level = outlineLevel
    if (rawProperties.collapsed === true) row.collapsed = true
    const style = styleToSheetJS(rawProperties.style)
    if (style) row.s = style
    rows[index - 1] = row
  }
  if (rows.some(Boolean)) result['!rows'] = rows
  if (sheet.autoFilter) {
    const autoFilter = autoFilterForSheetJS(sheet.autoFilter)
    if (autoFilter) result['!autofilter'] = autoFilter
  }
  if (sheet.pageSetup && sheet.pageSetup.margins) result['!margins'] = hydratePlainClone(sheet.pageSetup.margins)
  if (sheet.sheetProtection) result['!protect'] = hydratePlainClone(sheet.sheetProtection)
  if (sheet.outline) result['!outline'] = hydratePlainClone(sheet.outline)
  else if (sheet.properties && sheet.properties.outlineProperties) {
    result['!outline'] = {
      above: sheet.properties.outlineProperties.summaryBelow === false,
      left: sheet.properties.outlineProperties.summaryRight === false,
    }
  }
  result['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow - 1, c: maxCol - 1 } })
  return result
}

function workbookToSheetJS(model, targetFormat) {
  const workbook = { SheetNames: [], Sheets: {}, Workbook: { Sheets: [], Names: [] } }
  for (const sheet of model.sheets) {
    workbook.SheetNames.push(sheet.name)
    workbook.Sheets[sheet.name] = sheetToSheetJS(sheet, targetFormat)
    workbook.Workbook.Sheets.push({
      name: sheet.name,
      Hidden: sheet.state === 'veryHidden' ? 2 : sheet.state === 'hidden' ? 1 : 0,
    })
  }
  const modelNames = model.definedNames || (model.metadata && model.metadata.definedNames)
  if (Array.isArray(modelNames)) {
    for (const item of modelNames) {
      if (!item || !item.name) continue
      const ranges = Array.isArray(item.ranges) ? item.ranges : typeof item.ranges === 'string' ? [item.ranges] : []
      for (const range of ranges) {
        if (typeof range !== 'string' || !range) continue
        const name = { Name: String(item.name), Ref: range }
        const referencedSheetIndex =
          typeof item.localSheetRefId === 'string' ? model.sheets.findIndex((sheet) => sheet.id === item.localSheetRefId) : -1
        if (typeof item.localSheetRefId === 'string' && referencedSheetIndex < 0) continue
        const localSheetIndex = referencedSheetIndex >= 0
          ? referencedSheetIndex
          : Number.isInteger(item.localSheetIndex)
            ? item.localSheetIndex
            : item.localSheetId
        if (Number.isInteger(localSheetIndex)) name.Sheet = localSheetIndex
        if (item.hidden != null) name.Hidden = Boolean(item.hidden)
        if (item.comment) name.Comment = String(item.comment)
        workbook.Workbook.Names.push(name)
      }
    }
  }
  const properties = model.metadata && (model.metadata.properties || model.metadata)
  if (properties && typeof properties === 'object') {
    workbook.Props = {
      Author: properties.creator,
      LastAuthor: properties.lastModifiedBy,
      Title: properties.title,
      Subject: properties.subject,
      Comments: properties.description,
      Keywords: properties.keywords,
      Category: properties.category,
      Company: properties.company,
      Manager: properties.manager,
      Language: properties.language,
      ContentStatus: properties.contentStatus,
      Revision: properties.revision,
    }
    for (const [source, target] of [
      ['created', 'CreatedDate'],
      ['modified', 'ModifiedDate'],
      ['lastPrinted', 'LastPrinted'],
    ]) {
      if (!properties[source]) continue
      const date = new Date(properties[source])
      if (!Number.isNaN(date.getTime())) workbook.Props[target] = date
    }
    if (properties.custom && typeof properties.custom === 'object') workbook.Custprops = hydratePlainClone(properties.custom)
  }
  const metadata = model.metadata || {}
  workbook.Workbook.WBProps = {
    ...(metadata.workbookProperties && typeof metadata.workbookProperties === 'object' ? hydratePlainClone(metadata.workbookProperties) : {}),
    date1904: Boolean(metadata.date1904),
  }
  if (metadata.calcProperties && typeof metadata.calcProperties === 'object') {
    workbook.Workbook.CalcPr = hydratePlainClone(metadata.calcProperties)
  }
  const workbookViews = Array.isArray(metadata.workbookViews) ? hydratePlainClone(metadata.workbookViews) : []
  if (workbookViews.length) workbook.Workbook.WBView = workbookViews
  const activeTab = activeSheetIndex(model)
  if (!workbook.Workbook.WBView) workbook.Workbook.WBView = [{}]
  workbook.Workbook.WBView[0].activeTab = activeTab
  workbook.Workbook.WBView[0].firstSheet = Math.min(
    Math.max(0, Number(workbook.Workbook.WBView[0].firstSheet) || 0),
    Math.max(0, model.sheets.length - 1),
  )
  return workbook
}

// ---------------------------------------------------------------------------
// XLS and ODS without the document engine (SheetJS writers)
// ---------------------------------------------------------------------------

function countWhere(model, predicate) {
  let count = 0
  for (const sheet of model.sheets) for (const cell of Object.values(sheet.cells || {})) if (cell && predicate(cell)) count += 1
  return count
}

function anySheet(model, predicate) {
  return model.sheets.some((sheet) => predicate(sheet))
}

/**
 * What a workbook loses when written by the basic XLS (values only) or ODS writer. Only the
 * features this workbook actually uses are listed, in plain words.
 */
function nativeExportLosses(input, format) {
  const model = unwrapWorkbookModel(input)
  if (!model || !Array.isArray(model.sheets)) return []
  const losses = []
  const formulas = countWhere(model, (cell) => Boolean(cell.formula))
  if (format === 'xls' && formulas) losses.push(`${formulas.toLocaleString('en-US')} ${formulas === 1 ? 'formula' : 'formulas'} (their current results are kept as values)`)
  if (countWhere(model, (cell) => Boolean(cell.style && (cell.style.font || cell.style.fill || cell.style.border || cell.style.alignment)))) losses.push('fonts, fills, borders and alignment')
  if (countWhere(model, (cell) => Array.isArray(cell.richText) && cell.richText.length > 1)) losses.push('mixed text formatting inside cells')
  if (format === 'xls' && anySheet(model, (sheet) => Object.keys(sheet.rowHeights || {}).length || (sheet.hiddenRows || []).length)) losses.push('row heights and hidden rows')
  if (anySheet(model, (sheet) => sheet.frozen && (Number(sheet.frozen.rows) > 0 || Number(sheet.frozen.columns) > 0))) losses.push('frozen panes')
  if (format === 'xls' && (Array.isArray(model.definedNames) ? model.definedNames : []).some((item) => item && item.name && !/^_xlnm\./.test(item.name))) losses.push('named ranges')
  if (anySheet(model, (sheet) => sheet.dataValidations && Object.keys(sheet.dataValidations).length)) losses.push('dropdowns and data validation')
  if (anySheet(model, (sheet) => Array.isArray(sheet.conditionalFormattings) && sheet.conditionalFormattings.length)) losses.push('conditional formatting')
  if (anySheet(model, (sheet) => Array.isArray(sheet.tables) && sheet.tables.length)) losses.push('tables (their cells are kept)')
  if (anySheet(model, (sheet) => Array.isArray(sheet.charts) && sheet.charts.length)) losses.push('charts')
  if (anySheet(model, (sheet) => (Array.isArray(sheet.images) && sheet.images.length) || (sheet.requiresSourcePackage || []).includes('images'))) losses.push('pictures')
  if (anySheet(model, (sheet) => Array.isArray(sheet.sparklineGroups) && sheet.sparklineGroups.length)) losses.push('sparklines')
  if (anySheet(model, (sheet) => Array.isArray(sheet.pivots) && sheet.pivots.length)) losses.push('pivot table settings (their cells are kept)')
  if (anySheet(model, (sheet) => Boolean(sheet.sheetProtection))) losses.push('sheet protection')
  if (anySheet(model, (sheet) => sheet.pageSetup && (sheet.pageSetup.printArea || sheet.pageSetup.printTitlesRow || sheet.pageSetup.orientation === 'landscape'))) losses.push('print settings')
  return losses
}

function lossyExportError(format, losses) {
  const label = format === 'xls' ? 'Excel 97-2003 (.xls)' : 'OpenDocument (.ods)'
  const error = new Error(`${label} without the document engine keeps values and basic formatting only. This workbook would lose: ${losses.join('; ')}. Confirm to export anyway, or choose XLSX to keep everything.`)
  error.code = 'LOSSY_CONFIRM_REQUIRED'
  error.losses = losses
  return error
}

/** Quote the literal characters of a number format that the SheetJS ODS writer rejects. */
function odsSafeNumberFormat(format) {
  if (typeof format !== 'string' || !format || format === 'General') return format
  let isDate = false
  try { isDate = Boolean(XLSX.SSF.is_date(format)) } catch { isDate = false }
  if (!isDate) return format
  let output = ''
  for (let index = 0; index < format.length; index += 1) {
    const character = format[index]
    if (character === '"') {
      const end = format.indexOf('"', index + 1)
      const stop = end < 0 ? format.length : end + 1
      output += format.slice(index, stop)
      index = stop - 1
    } else if (character === '\\') {
      if (index + 1 < format.length) output += `"${format[index + 1] === '"' ? '' : format[index + 1]}"`
      index += 1
    } else if (character === '[') {
      const end = format.indexOf(']', index + 1)
      const stop = end < 0 ? format.length : end + 1
      output += format.slice(index, stop)
      index = stop - 1
    } else if (/[A-Za-z0-9]/.test(character)) {
      output += character
    } else {
      output += `"${character}"`
    }
  }
  return output.replace(/""/g, '')
}

function assertXlsLimits(model) {
  for (const sheet of model.sheets) {
    for (const address of Object.keys(sheet.cells || {})) {
      const position = addressPosition(address)
      if (position && (position.row > 65_536 || position.col > 256)) throw new Error(`XLS supports at most 65,536 rows and 256 columns. Cell ${address} in "${sheet.name}" is outside those limits. Save as XLSX to keep all cells.`)
    }
    for (const merge of sheet.merges || []) {
      const range = XLSX.utils.decode_range(merge)
      if (range.e.r >= 65_536 || range.e.c >= 256) throw new Error('A merged range exceeds the XLS limits. Save as XLSX to keep the entire worksheet.')
    }
  }
}

/** Excel 97-2003 values-only workbook through the SheetJS BIFF8 writer. */
function serializeXlsNatively(model, options = {}) {
  const losses = nativeExportLosses(model, 'xls')
  if (options.acceptLoss !== true && options.valuesOnly !== true) throw lossyExportError('xls', losses)
  const workbook = workbookToSheetJS(model, 'xls')
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name]
    for (const [key, cell] of Object.entries(sheet)) {
      if (key.startsWith('!') || !cell || typeof cell !== 'object') continue
      delete cell.f
      delete cell.F
      delete cell.D
      if (cell.t === 'z' && cell.v === undefined && !cell.c && !cell.l) delete sheet[key]
    }
    delete sheet['!protect']
  }
  const output = XLSX.write(workbook, { type: 'buffer', bookType: 'biff8' })
  const bytes = Buffer.isBuffer(output) ? output : Buffer.from(output)
  if (!bytes.subarray(0, 8).equals(CFB_SIGNATURE)) throw new Error('The XLS writer did not produce a valid workbook. Save as XLSX instead.')
  if (Array.isArray(options.warnings) && losses.length) options.warnings.push(`Saved as Excel 97-2003 values only; not kept: ${losses.join('; ')}.`)
  return bytes
}

/** OpenDocument spreadsheet with basic formatting through the SheetJS ODS writer. */
async function serializeOdsNatively(model, options = {}) {
  const losses = nativeExportLosses(model, 'ods')
  const workbook = workbookToSheetJS(model, 'ods')
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name]
    for (const [key, cell] of Object.entries(sheet)) {
      if (key.startsWith('!') || !cell || typeof cell !== 'object') continue
      if (typeof cell.z === 'string') cell.z = odsSafeNumberFormat(cell.z)
    }
  }
  const output = withoutSheetJSFormatNotes(() => XLSX.write(workbook, { type: 'buffer', bookType: 'ods' }))
  const zip = await JSZip.loadAsync(Buffer.isBuffer(output) ? output : Buffer.from(output))
  const contentEntry = zip.file('content.xml')
  if (!contentEntry) throw new Error('The ODS writer did not produce a valid workbook. Save as XLSX instead.')
  let content = await contentEntry.async('string')
  // SheetJS spells out the default textual="false", which its own reader takes for month names.
  content = content.replace(/ number:textual="false"/g, '')
  // Hidden sheets stay hidden (the writer drops the state; showing them could leak data).
  const hidden = new Set(model.sheets.filter((sheet) => sheet.state === 'hidden' || sheet.state === 'veryHidden').map((sheet) => sheet.name))
  if (hidden.size) {
    const style = '<style:style style:name="taSimpleHidden" style:family="table"><style:table-properties table:display="false" style:writing-mode="lr-tb"/></style:style>'
    content = content.replace(/<office:automatic-styles>/, `<office:automatic-styles>${style}`)
    content = content.replace(/<table:table(?=[\s>])([^>]*)>/g, (tag, attributes) => {
      const name = decodeXml(parseXmlAttributes(attributes)['table:name'] || '')
      if (!hidden.has(name)) return tag
      return /table:style-name="/.test(tag) ? tag.replace(/table:style-name="[^"]*"/, 'table:style-name="taSimpleHidden"') : tag.replace(/>$/, ' table:style-name="taSimpleHidden">')
    })
  }
  zip.file('content.xml', content)
  // ODF wants the uncompressed mimetype entry first.
  const packaged = new JSZip()
  packaged.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' })
  for (const [name, entry] of Object.entries(zip.files)) {
    if (name === 'mimetype' || entry.dir) continue
    packaged.file(name, await entry.async('nodebuffer'))
  }
  const bytes = Buffer.from(await packaged.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
  if (Array.isArray(options.warnings) && losses.length) options.warnings.push(`Saved as OpenDocument with basic formatting; not kept: ${losses.join('; ')}.`)
  return bytes
}

function normalizeOutputFormat(format) {
  if (typeof format !== 'string' || !format.trim()) throw new TypeError('An output format is required.')
  const normalized = format.trim().toLowerCase().replace(/^\./, '')
  if (normalized === 'tab') return 'tsv'
  if (normalized === 'fods') return 'ods'
  if (!['xlsx', 'xls', 'csv', 'tsv', 'txt', 'ods'].includes(normalized)) {
    throw new Error(`Cannot save ${normalized.toUpperCase()} directly. Save as XLSX, XLS, ODS, CSV, or TSV.`)
  }
  return normalized
}

async function serializeWorkbook(input, format = MODERN_FORMAT, options = {}) {
  const model = validateWorkbookModel(input)
  const normalizedFormat = normalizeOutputFormat(format)
  if (options == null) options = {}
  if (typeof options !== 'object' || Array.isArray(options)) throw new TypeError('Workbook serialization options must be an object.')
  if (normalizedFormat === 'xlsx') return serializeXlsx(model, options)
  const officeEngine = ['ods', 'xls'].includes(normalizedFormat)
    ? (typeof options.officeEngine === 'boolean' ? options.officeEngine : await officeEngineAvailable())
    : false
  if (normalizedFormat === 'ods' && !officeEngine) return serializeOdsNatively(model, options)
  if (normalizedFormat === 'xls' && !officeEngine) {
    assertXlsLimits(model)
    return serializeXlsNatively(model, options)
  }
  if (normalizedFormat === 'ods') {
    // SheetJS CE's ODS writer omits rich styles and hidden-sheet state. A real
    // ODF conversion keeps them and quoted sheet names attached to formulas.
    const { convertOfficeBytes } = require('./office-converter.cjs')
    const intermediate = await serializeXlsx(model, { ...options, baseBytes: null })
    const output = await convertOfficeBytes({ bytes: intermediate, inputExtension: 'xlsx', outputExtension: 'ods', filter: 'calc8' })
    const archive = await JSZip.loadAsync(output)
    if (await archive.file('mimetype')?.async('string') !== 'application/vnd.oasis.opendocument.spreadsheet') throw new Error('The conversion did not produce a valid ODS workbook. The original file is unchanged.')
    const content = archive.file('content.xml')
    if (!content || Number(content._data?.uncompressedSize || 0) > MAX_SOURCE_BYTES) throw new Error('The ODS workbook content is missing or too large.')
    const names = [...(await content.async('string')).matchAll(/<table:table(?=\s|>)([^>]*)>/g)].map(match => parseXmlAttributes(match[1])['table:name'])
    if (names.length !== model.sheets.length || names.some((name,index) => name !== model.sheets[index].name)) throw new Error('The ODS conversion changed worksheet names or order. Save as XLSX to retain the workbook.')
    return output
  }
  if (normalizedFormat === 'xls') {
    assertXlsLimits(model)
    const { convertOfficeBytes } = require('./office-converter.cjs')
    const intermediate = await serializeXlsx(model, { ...options, baseBytes: null })
    const convertedOutput = await convertOfficeBytes({ bytes: intermediate, inputExtension: 'xlsx', outputExtension: 'xls', filter: 'MS Excel 97' })
    const output = require('./legacy-layout-patch.cjs').retainLegacyLayout(convertedOutput, model)
    if (!output.subarray(0, 8).equals(CFB_SIGNATURE)) throw new Error('The Office engine did not produce a valid XLS workbook. The original file has not been changed.')
    const reopened = XLSX.read(output, { type: 'buffer', cellFormula: true })
    if (reopened.SheetNames.length !== model.sheets.length) throw new Error('The XLS conversion changed the number of worksheets. Save as XLSX to keep the entire workbook.')
    for (const sheet of model.sheets) {
      const convertedSheet = reopened.Sheets[sheet.name]
      if (!convertedSheet) throw new Error(`The XLS conversion omitted worksheet "${sheet.name}". Save as XLSX to keep it.`)
      for (const [address, cell] of Object.entries(sheet.cells || {})) {
        const converted = convertedSheet[address]
        const value = cell.value
        const invalidFormula = cell.formula && !converted?.f
        const invalidScalar = !cell.formula && value != null && value !== '' &&
          (cell.type === 'error' ? converted?.t !== 'e' || sheetJSErrorValue(converted) !== value
            : typeof value === 'number' ? typeof converted?.v !== 'number' || Math.abs(converted.v - value) > Math.max(1, Math.abs(value)) * 1e-12 : converted?.v !== value)
        if (invalidFormula || invalidScalar) throw new Error(`The XLS conversion could not retain cell ${address} in "${sheet.name}". Save as XLSX to keep all values and formulas.`)
      }
    }
    return output
  }

  // Delimited text: the active sheet, quoted per RFC 4180 (line breaks inside cells included),
  // values as displayed with four-digit years, in the source file's dialect when the save keeps
  // its format (options.dialect), else Excel-friendly UTF-8 with BOM.
  const activeSheet = model.sheets.find((sheet) => sheet.id === model.activeSheetId) || model.sheets[0]
  return writeDelimited(activeSheet, { format: normalizedFormat, dialect: options.dialect, warnings: options.warnings })
}

/**
 * The dialect to keep when a delimited document is saved in its own format: the one recorded
 * at open (model metadata), else the one main kept on the document record.
 */
function delimitedDialectFor(input, format, fallback) {
  const model = unwrapWorkbookModel(input)
  const recorded = model && model.metadata && model.metadata.dialect
  const dialect = recorded && typeof recorded === 'object' ? recorded : fallback
  if (!dialect || typeof dialect !== 'object') return undefined
  const source = String(dialect.format || '').toLowerCase()
  const family = (value) => (value === 'tsv' || value === 'tab' ? 'tsv' : value)
  if (source && family(source) !== family(String(format || '').toLowerCase())) return undefined
  return dialect
}

module.exports = {
  SUPPORTED_EXTENSIONS,
  workbookPayloadFromPath,
  workbookPayloadFromBytes,
  importWorkbookBytes,
  importWorkbookPath,
  serializeWorkbook,
  nativeExportLosses,
  delimitedDialectFor,
  supportedOrExtensionless,
  // exposed for QA
  _internal: { openFormulaToA1, parsePrintArea, sniffContainer, odsSafeNumberFormat },
}
