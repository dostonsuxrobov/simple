'use strict'

const fs = require('fs')
const path = require('path')
const ExcelJS = require('exceljs')
const JSZip = require('jszip')
const XLSX = require('xlsx')
const { extractLegacyBiffStyles } = require('./legacy-biff-styles.cjs')

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
  return { xml: data.toString('utf8'), partNames: Object.keys(zip.files) }
}

async function extractOoxmlWorkbookMetadata(buffer, warnings) {
  try {
    const packageMetadata = await workbookXmlFromOoxml(buffer)
    if (!packageMetadata) return {}
    const { xml, partNames } = packageMetadata
    const advancedPackageParts = []
    const advancedPartMatchers = [
      [/^xl\/charts\//i, 'charts'],
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
    return { calcProperties, workbookProperties, workbookViews, definedNames, advancedPackageParts }
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
  if (cell.type === ExcelJS.ValueType.Formula || (rawValue && typeof rawValue === 'object' && ('formula' in rawValue || 'sharedFormula' in rawValue))) {
    try {
      formula = normalizeFormula(cell.formula || rawValue.formula || rawValue.sharedFormula)
    } catch {
      formula = normalizeFormula(rawValue.formula || rawValue.sharedFormula)
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
    if (rawValue && (rawValue.shareType === 'array' || rawValue.shareType === 'shared')) {
      modelCell.formulaType = rawValue.shareType
    }
    if (rawValue && typeof rawValue.sharedFormula === 'string' && normalizeAddress(rawValue.sharedFormula)) {
      modelCell.formulaType = 'shared'
      modelCell.sharedFormulaMaster = normalizeAddress(rawValue.sharedFormula)
    }
    if (rawValue && typeof rawValue.ref === 'string' && rawValue.ref) modelCell.formulaRange = rawValue.ref
  } else if (rawValue && typeof rawValue === 'object' && rawValue.hyperlink) {
    modelCell.value = serializeCellValue(rawValue.text != null ? rawValue.text : cell.text, date1904)
  } else {
    modelCell.value = serializeCellValue(rawValue, date1904)
    if (rawValue instanceof Date) modelCell.type = 'date'
    else if (rawValue && typeof rawValue === 'object' && rawValue.error != null) modelCell.type = 'error'
    else if (rawValue && typeof rawValue === 'object' && Array.isArray(rawValue.richText)) modelCell.type = 'richText'
  }

  const cellText = cell.text == null ? '' : String(cell.text)
  const plainResult = richTextToPlainText(result)
  modelCell.display = cellText === '[object Object]' && plainResult != null ? String(plainResult) : cellText
  const style = normalizeStyle(cell.style)
  if (style) modelCell.style = style
  if (typeof cell.numFmt === 'string' && cell.numFmt && cell.numFmt !== 'General') modelCell.numFmt = cell.numFmt
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

function excelJSSheetToModel(worksheet, index, stats, warnings, date1904 = false) {
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
      if (column.style && typeof column.style.numFmt === 'string' && column.style.numFmt !== 'General') {
        properties.numFmt = column.style.numFmt
      }
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
    if (row.style && typeof row.style.numFmt === 'string' && row.style.numFmt !== 'General') properties.numFmt = row.style.numFmt
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
    pageSetup: plainClone(worksheet.pageSetup || {}),
    headerFooter: plainClone(worksheet.headerFooter || {}),
    rowBreaks: plainClone(worksheet.rowBreaks || []),
  }
  if (worksheet.autoFilter) model.autoFilter = plainClone(worksheet.autoFilter)

  if (worksheet.sheetProtection) model.sheetProtection = plainClone(worksheet.sheetProtection)
  const validations = cloneAdvancedRecord(worksheet.dataValidations && worksheet.dataValidations.model)
  if (Object.keys(validations.value).length) model.dataValidations = validations.value
  if (validations.truncated) {
    model.dataValidationsTruncated = true
    warnings.push(`Sheet "${worksheet.name}" has extensive data validation; the original OOXML package is required to retain every rule.`)
  }
  const conditionalFormattings = Array.isArray(worksheet.conditionalFormattings) ? worksheet.conditionalFormattings : []
  if (conditionalFormattings.length) {
    model.conditionalFormattings = conditionalFormattings.slice(0, MAX_ADVANCED_METADATA_ENTRIES).map((item) => plainClone(item))
  }
  if (conditionalFormattings.length > MAX_ADVANCED_METADATA_ENTRIES) {
    model.conditionalFormattingsTruncated = true
    warnings.push(`Sheet "${worksheet.name}" has extensive conditional formatting; the original OOXML package is required to retain every rule.`)
  }

  const advanced = []
  if (worksheet.dataValidations && worksheet.dataValidations.model && Object.keys(worksheet.dataValidations.model).length) {
    advanced.push('data validation')
  }
  if (conditionalFormattings.length) advanced.push('conditional formatting')
  if (worksheet.tables && Object.keys(worksheet.tables).length) advanced.push('tables')
  if (typeof worksheet.getImages === 'function' && worksheet.getImages().length) advanced.push('images')
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

async function importWithExcelJS(buffer, sourceName, warnings) {
  const ooxmlMetadata = await extractOoxmlWorkbookMetadata(buffer, warnings)
  const excelWorkbook = new ExcelJS.Workbook()
  await excelWorkbook.xlsx.load(buffer)
  const stats = initialStats(buffer.length)
  const sourceDate1904 = Boolean(excelWorkbook.properties && excelWorkbook.properties.date1904)
  // The renderer and output model use one canonical 1900-based serial system.
  // ExcelJS has already materialized formatted dates as real Date objects, so
  // converting those objects to the canonical epoch avoids four-year shifts.
  const sheets = excelWorkbook.worksheets.map((worksheet, index) => excelJSSheetToModel(worksheet, index, stats, warnings, false))
  if (!sheets.length) sheets.push(blankSheetModel())
  stats.sheetCount = sheets.length
  stats.sheets = sheets.length
  const activeIndex = activeSheetIndexFromExcelJS(excelWorkbook)
  const properties = workbookPropertiesFromExcelJS(excelWorkbook)
  const customProperties = customPropertiesFromOoxml(buffer)
  if (customProperties && Object.keys(customProperties).length) properties.custom = customProperties
  const definedNames = attachDefinedNameSheetIds(Array.isArray(ooxmlMetadata.definedNames) && ooxmlMetadata.definedNames.length
    ? ooxmlMetadata.definedNames
    : definedNamesFromExcelJS(excelWorkbook), sheets)
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
  }
  if (sheet['!autofilter']) model.autoFilter = plainClone(sheet['!autofilter'])
  if (sheet['!margins']) model.pageSetup = { margins: plainClone(sheet['!margins']) }
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

function compatibilityWarnings(format) {
  const warnings = []
  if (format !== MODERN_FORMAT && RICH_OOXML_FORMATS.has(format)) {
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
const DELIMITER_CANDIDATES = [',', ';', '\t', '|']
const ISO_DATE_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
const YMD_SLASH_RE = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/
const SLASH_DATE_RE = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/
const EU_DECIMAL_RE = /^-?\d+,\d+$/
const EU_GROUPED_RE = /^-?\d{1,3}(?:\.\d{3})+(?:,\d+)?$/
const US_DECIMAL_RE = /^-?\d+\.\d+$/
const US_GROUPED_RE = /^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/

function decodeDelimitedText(buffer) {
  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
    return buffer.subarray(3).toString('utf8')
  }
  if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
    return buffer.subarray(2).toString('utf16le')
  }
  if (buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff) {
    return Buffer.from(buffer.subarray(2)).swap16().toString('utf16le')
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buffer)
  } catch {
    // BOM-less non-UTF-8 text falls back to the classic Windows codepage.
  }
  try {
    return new TextDecoder('windows-1252').decode(buffer)
  } catch {
    return buffer.toString('latin1')
  }
}

function countDelimitersOutsideQuotes(line, delimiter) {
  let count = 0
  let quoted = false
  for (const character of line) {
    if (character === '"') quoted = !quoted
    else if (!quoted && character === delimiter) count += 1
  }
  return count
}

function sniffDelimiter(text, extensionDefault) {
  const lines = text.split(/\r\n|[\r\n]/).filter((line) => line.trim() !== '').slice(0, 50)
  if (!lines.length) return extensionDefault
  let best
  for (const delimiter of DELIMITER_CANDIDATES) {
    const counts = lines.map((line) => countDelimitersOutsideQuotes(line, delimiter))
    const tally = new Map()
    for (const count of counts) {
      if (count > 0) tally.set(count, (tally.get(count) || 0) + 1)
    }
    let mode = 0
    let modeLines = 0
    for (const [count, occurrences] of tally) {
      if (occurrences > modeLines || (occurrences === modeLines && count > mode)) {
        mode = count
        modeLines = occurrences
      }
    }
    if (!mode) continue
    const consistency = modeLines / counts.length
    if (consistency < 0.5) continue
    const score = consistency * 1_000_000 + (delimiter === extensionDefault ? 1_000 : 0) + Math.min(mode, 999)
    if (!best || score > best.score) best = { delimiter, score }
  }
  return best ? best.delimiter : extensionDefault
}

function parseDelimitedRecords(text, delimiter) {
  const records = []
  let record = []
  let field = ''
  let quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 1
        } else {
          quoted = false
        }
      } else {
        field += character
      }
    } else if (character === '"') {
      quoted = true
    } else if (character === delimiter) {
      record.push(field)
      field = ''
    } else if (character === '\n' || character === '\r') {
      if (character === '\r' && text[index + 1] === '\n') index += 1
      record.push(field)
      records.push(record)
      field = ''
      record = []
    } else {
      field += character
    }
  }
  if (field !== '' || record.length) {
    record.push(field)
    records.push(record)
  }
  return records
}

function detectDecimalComma(records, delimiter) {
  if (delimiter === ';') return true
  let euSignals = 0
  let usSignals = 0
  let scanned = 0
  for (const record of records) {
    for (const rawField of record) {
      if (scanned >= 50_000) return euSignals > 0 && euSignals > usSignals
      scanned += 1
      const field = rawField.trim().replace(/%$/, '')
      if (!field) continue
      const eu = EU_DECIMAL_RE.test(field) || EU_GROUPED_RE.test(field)
      const us = US_DECIMAL_RE.test(field) || US_GROUPED_RE.test(field)
      if (eu && !us) euSignals += 1
      else if (us && !eu) usSignals += 1
    }
  }
  return euSignals > 0 && euSignals > usSignals
}

function parseDelimitedNumber(rawText, decimalComma) {
  let text = rawText
  let sign = 1
  const parenthesized = /^\((.+)\)$/.exec(text)
  if (parenthesized) {
    text = parenthesized[1]
    sign = -1
  }
  let percent = false
  if (text.endsWith('%')) {
    percent = true
    text = text.slice(0, -1)
  }
  const currency = text.includes('$')
  if (currency) text = text.replace(/\$/g, '')
  text = text.trim()
  if (!text) return undefined
  let normalized = text
  let grouped = false
  let euStyled = false
  if (decimalComma) {
    if (EU_GROUPED_RE.test(text)) {
      grouped = true
      euStyled = true
      normalized = text.replace(/\./g, '').replace(',', '.')
    } else if (EU_DECIMAL_RE.test(text)) {
      euStyled = true
      normalized = text.replace(',', '.')
    }
  } else if (US_GROUPED_RE.test(text)) {
    grouped = true
    normalized = text.replace(/,/g, '')
  }
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(normalized)) return undefined
  const value = Number(normalized)
  if (!Number.isFinite(value)) return undefined
  const result = { value: sign * (percent ? value / 100 : value) }
  if (euStyled && !percent && !currency) {
    const decimalMatch = /\.(\d+)$/.exec(normalized)
    const decimals = decimalMatch ? Math.min(decimalMatch[1].length, 10) : 0
    if (grouped) result.numFmt = `#,##0${decimals ? `.${'0'.repeat(decimals)}` : ''}`
    else if (decimals) result.numFmt = `0.${'0'.repeat(decimals)}`
  }
  return result
}

function delimitedDateSerial(year, month, day, hours = 0, minutes = 0, seconds = 0) {
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return null
  if (hours > 23 || minutes > 59 || seconds > 59) return null
  const utc = Date.UTC(year, month - 1, day, hours, minutes, seconds)
  const check = new Date(utc)
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) return null
  return (utc - Date.UTC(1899, 11, 30)) / 86_400_000
}

function delimitedFieldToModelCell(field, decimalComma) {
  if (field === '') return undefined
  const trimmed = field.trim()
  if (!trimmed) return { value: field, display: field }
  if (trimmed.startsWith('=') && trimmed.length > 1) {
    return { formula: trimmed.slice(1), result: field, display: field }
  }
  if (trimmed === 'TRUE' || trimmed === 'FALSE') return { value: trimmed === 'TRUE', display: field }
  if (ERROR_TEXT_TO_CODE[trimmed] != null) return { value: trimmed, type: 'error', display: field }
  // Leading-zero identifiers and digit runs beyond double precision must stay
  // text; coercing them to numbers silently corrupts the source data.
  if (/^0\d+$/.test(trimmed) || /^-?\d{16,}$/.test(trimmed)) {
    return { value: trimmed, display: field, numFmt: '@' }
  }
  const isoMatch = ISO_DATE_RE.exec(trimmed)
  if (isoMatch) {
    const serial = delimitedDateSerial(
      Number(isoMatch[1]),
      Number(isoMatch[2]),
      Number(isoMatch[3]),
      Number(isoMatch[4] || 0),
      Number(isoMatch[5] || 0),
      Number(isoMatch[6] || 0),
    )
    if (serial != null) {
      const numFmt = isoMatch[6] != null ? 'yyyy-mm-dd hh:mm:ss' : isoMatch[4] != null ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd'
      return { value: serial, display: field, numFmt }
    }
  }
  const ymdMatch = YMD_SLASH_RE.exec(trimmed)
  if (ymdMatch) {
    const serial = delimitedDateSerial(Number(ymdMatch[1]), Number(ymdMatch[2]), Number(ymdMatch[3]))
    if (serial != null) return { value: serial, display: field, numFmt: 'yyyy-mm-dd' }
  }
  const slashMatch = SLASH_DATE_RE.exec(trimmed)
  if (slashMatch) {
    const first = Number(slashMatch[1])
    const second = Number(slashMatch[2])
    let year = Number(slashMatch[3])
    if (slashMatch[3].length <= 2) year = year < 50 ? 2000 + year : 1900 + year
    let month
    let day
    let numFmt = 'm/d/yy'
    if (decimalComma && second <= 12) {
      day = first
      month = second
      numFmt = 'd/m/yy'
    } else if (first <= 12) {
      month = first
      day = second
    }
    if (month != null) {
      const serial = delimitedDateSerial(year, month, day)
      if (serial != null) return { value: serial, display: field, numFmt }
    }
  }
  const parsed = parseDelimitedNumber(trimmed, decimalComma)
  if (parsed) {
    const modelCell = { value: parsed.value, display: field }
    if (parsed.numFmt) modelCell.numFmt = parsed.numFmt
    return modelCell
  }
  return { value: field, display: field }
}

function importDelimitedText(buffer, sourceName, sourceFormat, warnings) {
  if (!buffer.length) return undefined
  const text = decodeDelimitedText(buffer)
  const head = text.slice(0, 2_048).replace(/^[\s\uFEFF]+/, '')
  // XML/HTML tables, SYLK, and DIF sources keep using the SheetJS readers.
  if (/^</.test(head) || /^ID;P/i.test(head) || /^TABLE\r?\n/.test(head)) return undefined
  const extensionDefault = sourceFormat === 'csv' ? ',' : sourceFormat === 'prn' ? undefined : '\t'
  const delimiter = sniffDelimiter(text, extensionDefault)
  if (!delimiter) return undefined
  const records = parseDelimitedRecords(text, delimiter)
  const decimalComma = detectDecimalComma(records, delimiter)
  const stats = initialStats(buffer.length)
  const cells = {}
  let maxRow = 0
  let maxCol = 0
  let truncated = false
  records.forEach((record, rowIndex) => {
    if (rowIndex >= MAX_METADATA_ROWS) {
      truncated = true
      return
    }
    record.forEach((field, colIndex) => {
      if (colIndex >= MAX_METADATA_COLS) {
        truncated = true
        return
      }
      const modelCell = delimitedFieldToModelCell(field, decimalComma)
      if (!modelCell) return
      cells[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })] = modelCell
      maxRow = Math.max(maxRow, rowIndex + 1)
      maxCol = Math.max(maxCol, colIndex + 1)
      updateStatsForCell(stats, modelCell, rowIndex + 1, colIndex + 1)
    })
  })
  if (truncated) warnings.push('Some rows or columns beyond modern spreadsheet limits were dropped from this delimited text file.')
  const sheet = {
    ...blankSheetModel(),
    sourceSheetName: 'Sheet1',
    sourceSheetIndex: 0,
    rowCount: Math.max(1, maxRow),
    colCount: Math.max(1, maxCol),
    cells,
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
        delimiter,
        decimalComma,
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

function importWithSheetJS(buffer, sourceName, sourceFormat, warnings) {
  let sheetWorkbook
  if (!buffer.length && ['csv', 'tsv', 'tab', 'txt', 'prn'].includes(sourceFormat)) {
    sheetWorkbook = { SheetNames: ['Sheet1'], Sheets: { Sheet1: { '!ref': 'A1:A1' } } }
  } else {
    sheetWorkbook = XLSX.read(buffer, {
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
    })
  }

  const stats = initialStats(buffer.length)
  const rawDate1904 = sheetWorkbook.Workbook && sheetWorkbook.Workbook.WBProps && sheetWorkbook.Workbook.WBProps.date1904
  const sourceDate1904 = rawDate1904 === true || rawDate1904 === 1 || rawDate1904 === '1' || rawDate1904 === 'true'
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

function finalizePayload(sourceName, sourceFormat, imported, warnings) {
  const safeName = path.basename(sourceName) || `Workbook.${sourceFormat}`
  const requiresSaveAs = sourceFormat !== MODERN_FORMAT
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
    warnings: uniqueWarnings,
    stats: imported.stats,
    workbook: imported.model,
  }
}

const CFB_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
const ENCRYPTED_STREAM_MARKERS = [Buffer.from('EncryptedPackage', 'utf16le'), Buffer.from('EncryptionInfo', 'utf16le')]
const ENCRYPTED_WORKBOOK_MESSAGE = 'This workbook is password-protected; encrypted files are not supported.'

function isEncryptedWorkbook(buffer) {
  if (buffer.length < 8 || !buffer.subarray(0, 8).equals(CFB_SIGNATURE)) return false
  // Standard and agile OOXML encryption both store these CFB stream names.
  return ENCRYPTED_STREAM_MARKERS.some((marker) => buffer.includes(marker))
}

async function workbookPayloadFromBytes(name, data) {
  const extension = ensureSupportedExtension(name)
  const sourceFormat = extension.slice(1)
  const sourceName = path.basename(name)
  const buffer = bytesToBuffer(data)
  if (isEncryptedWorkbook(buffer)) throw new Error(ENCRYPTED_WORKBOOK_MESSAGE)
  const warnings = compatibilityWarnings(sourceFormat)

  let imported
  const hasZipSignature =
    buffer.length >= 4 &&
    buffer[0] === 0x50 &&
    buffer[1] === 0x4b &&
    ((buffer[2] === 0x03 && buffer[3] === 0x04) ||
      (buffer[2] === 0x05 && buffer[3] === 0x06) ||
      (buffer[2] === 0x07 && buffer[3] === 0x08))
  if (RICH_OOXML_FORMATS.has(sourceFormat) && hasZipSignature) {
    try {
      imported = await importWithExcelJS(buffer, sourceName, warnings)
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
  } else if (sourceFormat === MODERN_FORMAT) {
    if (!hasZipSignature) {
      const hasCompoundFileSignature = buffer.length >= 8 && buffer.subarray(0, 8).equals(CFB_SIGNATURE)
      if (hasCompoundFileSignature) {
        throw new Error(`Unable to open ${sourceName}: password-protected modern Excel workbooks are not supported.`)
      }
      throw new Error(`Unable to open ${sourceName}: the file is not a valid XLSX package.`)
    }
  } else {
    try {
      if (DELIMITED_TEXT_FORMATS.has(sourceFormat)) {
        try {
          imported = importDelimitedText(buffer, sourceName, sourceFormat, warnings)
        } catch {
          warnings.push('The delimited-text reader could not parse this file; the compatibility reader was used instead.')
        }
      }
      if (!imported) imported = importWithSheetJS(buffer, sourceName, sourceFormat, warnings)
    } catch (error) {
      if (/password|encrypt/i.test(String(error && error.message))) {
        const wrapped = new Error(ENCRYPTED_WORKBOOK_MESSAGE)
        wrapped.cause = error
        throw wrapped
      }
      const wrapped = new Error(`Unable to open ${sourceName}: ${error.message || 'invalid or damaged spreadsheet'}`)
      wrapped.cause = error
      throw wrapped
    }
  }
  return finalizePayload(sourceName, sourceFormat, imported, warnings)
}

async function workbookPayloadFromPath(filePath) {
  if (typeof filePath !== 'string' || !filePath.trim()) throw new TypeError('A spreadsheet file path is required.')
  const resolvedPath = path.resolve(filePath)
  ensureSupportedExtension(resolvedPath)
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
  return workbookPayloadFromBytes(path.basename(resolvedPath), data)
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

function setExcelJSCell(cell, modelCell) {
  const formula = normalizeFormula(modelCell && modelCell.formula)
  if (formula) {
    const sharedFormulaMaster =
      modelCell && modelCell.formulaType === 'shared' && normalizeAddress(modelCell.sharedFormulaMaster)
        ? normalizeAddress(modelCell.sharedFormulaMaster)
        : undefined
    const formulaValue = sharedFormulaMaster ? { sharedFormula: sharedFormulaMaster } : { formula }
    if (modelCell.result !== undefined) {
      const result = formulaResultForExcelJS(modelCell.result, modelCell.resultType)
      if (result !== undefined) formulaValue.result = result
    }
    if (!sharedFormulaMaster && (modelCell.formulaType === 'array' || modelCell.formulaType === 'shared')) {
      formulaValue.shareType = modelCell.formulaType
    }
    if (typeof modelCell.formulaRange === 'string' && normalizeRange(modelCell.formulaRange)) {
      formulaValue.ref = normalizeRange(modelCell.formulaRange)
    }
    cell.value = formulaValue
  } else {
    let value = deserializeCellValue(modelCell ? modelCell.value : null)
    const hyperlink = normalizeHyperlink(modelCell && modelCell.hyperlink)
    if (modelCell && modelCell.type === 'error') value = { error: String(modelCell.value || '#VALUE!') }
    if (hyperlink) {
      const textValue = richTextToPlainText(modelCell.value)
      value = {
        text: textValue == null || typeof textValue === 'object' ? String(modelCell.display || hyperlink.target) : String(textValue),
        hyperlink: hyperlink.target,
      }
      const tooltip = hyperlink.tooltip || (modelCell && modelCell.hyperlinkTooltip)
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

function applyExcelJSSheetMetadata(worksheet, sheet, preserveBase) {
  if (sheet.properties && typeof sheet.properties === 'object') {
    const properties = hydratePlainClone(sheet.properties)
    worksheet.properties = { ...(worksheet.properties || {}), ...properties }
    if (properties.outlineProperties) {
      worksheet.properties.outlineProperties = {
        ...((worksheet.properties && worksheet.properties.outlineProperties) || {}),
        ...properties.outlineProperties,
      }
    }
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
    if (pageSetup.margins) {
      worksheet.pageSetup.margins = { ...((worksheet.pageSetup && worksheet.pageSetup.margins) || {}), ...pageSetup.margins }
    }
  }
  if (sheet.headerFooter && typeof sheet.headerFooter === 'object') {
    worksheet.headerFooter = { ...(worksheet.headerFooter || {}), ...hydratePlainClone(sheet.headerFooter) }
  }
  if (Array.isArray(sheet.rowBreaks)) worksheet.rowBreaks = hydratePlainClone(sheet.rowBreaks)
  if (Object.prototype.hasOwnProperty.call(sheet, 'sheetProtection')) {
    worksheet.sheetProtection = sheet.sheetProtection ? hydratePlainClone(sheet.sheetProtection) : null
  }
  if (sheet.dataValidations && (!preserveBase || !sheet.dataValidationsTruncated)) {
    worksheet.dataValidations.model = hydratePlainClone(sheet.dataValidations)
  }
  if (Array.isArray(sheet.conditionalFormattings) && (!preserveBase || !sheet.conditionalFormattingsTruncated)) {
    worksheet.conditionalFormattings = hydratePlainClone(sheet.conditionalFormattings)
  }
  if (Object.prototype.hasOwnProperty.call(sheet, 'autoFilter')) {
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

  const existing = excelWorkbook.worksheets.slice()
  const used = new Set()
  const matches = model.sheets.map((sheet, index) => {
    let worksheet
    if (Number.isInteger(sheet.sourceWorksheetId)) worksheet = existing.find((candidate) => candidate.id === sheet.sourceWorksheetId && !used.has(candidate))
    if (!worksheet && typeof sheet.sourceSheetName === 'string') {
      worksheet = existing.find((candidate) => candidate.name === sheet.sourceSheetName && !used.has(candidate))
    }
    if (!worksheet && Number.isInteger(sheet.sourceSheetIndex)) {
      const candidate = existing[sheet.sourceSheetIndex]
      if (candidate && !used.has(candidate)) worksheet = candidate
    }
    if (!worksheet) worksheet = existing.find((candidate) => candidate.name === sheet.name && !used.has(candidate))
    if (!worksheet && !Object.prototype.hasOwnProperty.call(sheet, 'sourceSheetIndex')) {
      const candidate = existing[index]
      if (candidate && !used.has(candidate)) worksheet = candidate
    }
    if (worksheet) used.add(worksheet)
    return worksheet
  })

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

function renderDefinedNamesBlock(model) {
  const source = Array.isArray(model.definedNames)
    ? model.definedNames
    : model.metadata && Array.isArray(model.metadata.definedNames)
      ? model.metadata.definedNames
      : undefined
  if (!source) return undefined
  const items = []
  for (const item of source.slice(0, MAX_ADVANCED_METADATA_ENTRIES)) {
    if (!item || typeof item.name !== 'string' || !item.name) continue
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

async function patchXlsxWorkbookMetadata(buffer, model, baseBuffer) {
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

  const namesBlock = renderDefinedNamesBlock(model)
  if (namesBlock !== undefined) {
    const namesPattern = /<definedNames\b[^>]*>[\s\S]*?<\/definedNames\s*>/i
    if (namesPattern.test(xml)) xml = xml.replace(namesPattern, namesBlock)
    else if (namesBlock) xml = xml.replace(/<calcPr\b/i, `${namesBlock}<calcPr`)
  }

  zip.file('xl/workbook.xml', xml)
  const customProperties = model.metadata && model.metadata.properties && model.metadata.properties.custom
  const customXml = (await customPropertiesXmlFromBase(baseBuffer)) || renderCustomPropertiesXml(customProperties)
  await addCustomPropertiesPart(zip, customXml)
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

function writeModelSheetToExcelJS(worksheet, sheet, preserveBase) {
  clearExcelJSWorksheetCells(worksheet)
  worksheet.state = sheet.state === 'hidden' || sheet.state === 'veryHidden' ? sheet.state : 'visible'
  applyExcelJSSheetMetadata(worksheet, sheet, preserveBase)
  applyExcelJSSheetDimensions(worksheet, sheet)

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
    setExcelJSCell(worksheet.getCell(position.address), modelCell)
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
    await excelWorkbook.xlsx.load(baseBuffer)
  } catch (error) {
    const wrapped = new Error(`Unable to use the source workbook for fidelity-preserving save: ${error.message || 'invalid OOXML package'}`)
    wrapped.code = 'BASE_WORKBOOK_LOAD_FAILED'
    wrapped.cause = error
    throw wrapped
  }
  return { excelWorkbook, preserveBase: true, baseBuffer }
}

async function serializeXlsx(model, options = {}) {
  const { excelWorkbook, preserveBase, baseBuffer } = await excelJSWorkbookForSerialization(options)
  applyWorkbookProperties(excelWorkbook, model.metadata && (model.metadata.properties || model.metadata))
  excelWorkbook.properties.date1904 = Boolean(model.metadata && model.metadata.date1904)
  applyExcelJSCalcProperties(excelWorkbook, model)

  const worksheets = matchExcelJSWorksheets(excelWorkbook, model, preserveBase)
  model.sheets.forEach((sheet, index) => writeModelSheetToExcelJS(worksheets[index], sheet, preserveBase))

  applyExcelJSWorkbookViews(excelWorkbook, model)
  applyDefinedNamesToExcelJS(excelWorkbook, model.definedNames || (model.metadata && model.metadata.definedNames))
  const result = await excelWorkbook.xlsx.writeBuffer({ useStyles: true, useSharedStrings: true })
  return patchXlsxWorkbookMetadata(Buffer.from(result), model, baseBuffer)
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

function normalizeOutputFormat(format) {
  if (typeof format !== 'string' || !format.trim()) throw new TypeError('An output format is required.')
  const normalized = format.trim().toLowerCase().replace(/^\./, '')
  if (normalized === 'tab') return 'tsv'
  if (normalized === 'fods') return 'ods'
  if (!['xlsx', 'csv', 'tsv', 'txt', 'ods'].includes(normalized)) {
    throw new Error(`Cannot save ${normalized.toUpperCase()} directly. Save as XLSX, ODS, CSV, or TSV.`)
  }
  return normalized
}

async function serializeWorkbook(input, format = MODERN_FORMAT, options = {}) {
  const model = validateWorkbookModel(input)
  const normalizedFormat = normalizeOutputFormat(format)
  if (options == null) options = {}
  if (typeof options !== 'object' || Array.isArray(options)) throw new TypeError('Workbook serialization options must be an object.')
  if (normalizedFormat === 'xlsx') return serializeXlsx(model, options)

  const sheetWorkbook = workbookToSheetJS(model, normalizedFormat)
  let result
  if (normalizedFormat === 'ods') {
    result = XLSX.write(sheetWorkbook, {
      type: 'buffer',
      bookType: 'ods',
      cellStyles: true,
      bookSST: true,
      compression: true,
    })
  } else {
    const activeSheet = model.sheets.find((sheet) => sheet.id === model.activeSheetId) || model.sheets[0]
    result = XLSX.write(sheetWorkbook, {
      type: 'buffer',
      bookType: normalizedFormat === 'txt' ? 'txt' : 'csv',
      sheet: activeSheet.name,
      FS: normalizedFormat === 'tsv' || normalizedFormat === 'txt' ? '\t' : ',',
      RS: '\r\n',
    })
  }
  return Buffer.isBuffer(result) ? result : Buffer.from(result)
}

module.exports = {
  SUPPORTED_EXTENSIONS,
  workbookPayloadFromPath,
  workbookPayloadFromBytes,
  serializeWorkbook,
}
