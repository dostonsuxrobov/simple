'use strict'

// DrawingML chart import/export for XLSX packages.
//
// Import: xl/workbook.xml -> worksheet rels -> drawing -> anchors -> chart parts,
// parsed into the SheetChart model (src/spreadsheet-types.ts).
// Export: ExcelJS drops charts, so after it writes the package we add a drawing
// per sheet (merged with any picture drawing ExcelJS wrote), chart parts,
// relationships and content types. Unmodified imported charts are copied
// byte-for-byte from the source package (with their style/colour/embedding
// parts); new or edited charts are generated as strict DrawingML.

const crypto = require('crypto')
const path = require('path')

const MAX_PART_BYTES = 32 * 1024 * 1024
const MAX_CHARTS_PER_WORKBOOK = 2_000
const MAX_CACHE_POINTS = 20_000
const EMU_PER_PIXEL = 9525
const MAX_ROWS = 1_048_576
const MAX_COLS = 16_384

const NS = {
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}
const CANONICAL_PREFIX = new Map(Object.entries(NS).map(([prefix, uri]) => [uri, prefix]))
// Strict OOXML namespaces map onto the transitional prefixes.
CANONICAL_PREFIX.set('http://purl.oclc.org/ooxml/drawingml/chart', 'c')
CANONICAL_PREFIX.set('http://purl.oclc.org/ooxml/drawingml/main', 'a')
CANONICAL_PREFIX.set('http://purl.oclc.org/ooxml/officeDocument/relationships', 'r')
CANONICAL_PREFIX.set('http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing', 'xdr')
CANONICAL_PREFIX.set('http://purl.oclc.org/ooxml/spreadsheetml/main', 'x')

const REL = {
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  chartEx: 'http://schemas.microsoft.com/office/2014/relationships/chartEx',
}
const CT = {
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  chartEx: 'application/vnd.ms-office.chartex+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chartStyle: 'application/vnd.ms-office.chartstyle+xml',
  chartColors: 'application/vnd.ms-office.chartcolorstyle+xml',
  chartShapes: 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml',
  themeOverride: 'application/vnd.openxmlformats-officedocument.themeOverride+xml',
  relationships: 'application/vnd.openxmlformats-package.relationships+xml',
}
const DEFAULT_EXTENSION_TYPES = {
  png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff',
  emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', bin: 'application/vnd.openxmlformats-officedocument.oleObject', xml: 'application/xml',
  rels: CT.relationships,
}

const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']
const OFFICE_ACCENTS = OFFICE_THEME.slice(4, 10)

// ---------------------------------------------------------------------------
// Minimal namespace-aware XML parser (with source offsets for byte-exact slicing)
// ---------------------------------------------------------------------------

function decodeEntities(text) {
  if (!text.includes('&')) return text
  return text.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi, (all, entity) => {
    const lower = entity.toLowerCase()
    if (lower === 'lt') return '<'
    if (lower === 'gt') return '>'
    if (lower === 'amp') return '&'
    if (lower === 'quot') return '"'
    if (lower === 'apos') return "'"
    const code = lower.startsWith('#x') ? parseInt(lower.slice(2), 16) : parseInt(lower.slice(1), 10)
    return Number.isFinite(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : all
  })
}

const ATTRIBUTE_RE = /([^\s=/]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

function normalizeName(qualified, namespaces, isAttribute) {
  const colon = qualified.indexOf(':')
  if (colon < 0) {
    if (isAttribute) return qualified
    const uri = namespaces.get('')
    const prefix = uri && CANONICAL_PREFIX.get(uri)
    return prefix ? `${prefix}:${qualified}` : qualified
  }
  const prefix = qualified.slice(0, colon)
  if (prefix === 'xmlns' || prefix === 'xml') return qualified
  const uri = namespaces.get(prefix)
  const canonical = uri && CANONICAL_PREFIX.get(uri)
  return canonical ? `${canonical}:${qualified.slice(colon + 1)}` : qualified
}

function parseXml(xml) {
  if (typeof xml !== 'string') throw new TypeError('XML text expected')
  const root = { name: '#document', attrs: {}, children: [], text: '', start: 0, end: xml.length }
  const stack = [root]
  const nsStack = [new Map([['xml', 'http://www.w3.org/XML/1998/namespace']])]
  let index = 0
  const length = xml.length
  while (index < length) {
    const lt = xml.indexOf('<', index)
    const top = stack[stack.length - 1]
    if (lt < 0) { top.text += decodeEntities(xml.slice(index)); break }
    if (lt > index) top.text += decodeEntities(xml.slice(index, lt))
    if (xml.startsWith('<!--', lt)) { const end = xml.indexOf('-->', lt + 4); index = end < 0 ? length : end + 3; continue }
    if (xml.startsWith('<![CDATA[', lt)) { const end = xml.indexOf(']]>', lt + 9); top.text += xml.slice(lt + 9, end < 0 ? length : end); index = end < 0 ? length : end + 3; continue }
    if (xml.startsWith('<?', lt)) { const end = xml.indexOf('?>', lt + 2); index = end < 0 ? length : end + 2; continue }
    if (xml.startsWith('<!', lt)) throw new Error('Document type declarations are not allowed in package parts.')
    if (xml[lt + 1] === '/') {
      const gt = xml.indexOf('>', lt)
      if (gt < 0) throw new Error('Unterminated closing tag.')
      if (stack.length > 1) { const node = stack.pop(); node.end = gt + 1; nsStack.pop() }
      index = gt + 1
      continue
    }
    let cursor = lt + 1
    let quote = ''
    for (; cursor < length; cursor += 1) {
      const character = xml[cursor]
      if (quote) { if (character === quote) quote = '' } else if (character === '"' || character === "'") quote = character
      else if (character === '>') break
    }
    if (cursor >= length) throw new Error('Unterminated tag.')
    const selfClosing = xml[cursor - 1] === '/'
    const inner = xml.slice(lt + 1, selfClosing ? cursor - 1 : cursor)
    const nameMatch = /^[^\s/>]+/.exec(inner)
    if (!nameMatch) throw new Error('Malformed tag.')
    const rawName = nameMatch[0]
    const rawAttributes = []
    let namespaces = nsStack[nsStack.length - 1]
    ATTRIBUTE_RE.lastIndex = rawName.length
    let match
    while ((match = ATTRIBUTE_RE.exec(inner))) {
      const key = match[1]
      const value = decodeEntities(match[2] !== undefined ? match[2] : match[3] || '')
      if (key === 'xmlns' || key.startsWith('xmlns:')) {
        if (namespaces === nsStack[nsStack.length - 1]) namespaces = new Map(namespaces)
        namespaces.set(key === 'xmlns' ? '' : key.slice(6), value)
      }
      rawAttributes.push([key, value])
    }
    const attrs = {}
    for (const [key, value] of rawAttributes) attrs[normalizeName(key, namespaces, true)] = value
    const node = { name: normalizeName(rawName, namespaces, false), attrs, children: [], text: '', start: lt, openEnd: cursor + 1, end: selfClosing ? cursor + 1 : -1 }
    top.children.push(node)
    if (!selfClosing) { stack.push(node); nsStack.push(namespaces) }
    index = cursor + 1
    if (stack.length > 400) throw new Error('XML nesting is too deep.')
  }
  return root
}

function documentElement(doc) {
  return doc.children.find((node) => node.name !== '#text') || null
}

function child(node, name) {
  if (!node) return null
  for (const item of node.children) if (item.name === name) return item
  return null
}

function childrenOf(node, name) {
  return node ? node.children.filter((item) => item.name === name) : []
}

function path1(node, ...names) {
  let current = node
  for (const name of names) { current = child(current, name); if (!current) return null }
  return current
}

function valOf(node, name, fallback) {
  const item = child(node, name)
  return item && item.attrs.val !== undefined ? item.attrs.val : fallback
}

function boolVal(node, name, fallback) {
  const item = child(node, name)
  if (!item) return fallback
  const value = item.attrs.val
  if (value === undefined) return true // CT_Boolean defaults to true
  return value === '1' || value === 'true'
}

function numVal(node, name) {
  const value = Number(valOf(node, name))
  return Number.isFinite(value) ? value : undefined
}

function textContent(node) {
  if (!node) return ''
  let text = node.text || ''
  for (const item of node.children) text += textContent(item)
  return text
}

function findDescendant(node, name, depth = 0) {
  if (!node || depth > 12) return null
  for (const item of node.children) {
    if (item.name === name) return item
    const found = findDescendant(item, name, depth + 1)
    if (found) return found
  }
  return null
}

// ---------------------------------------------------------------------------
// XML output helpers
// ---------------------------------------------------------------------------

function escapeXml(value) {
  return String(value == null ? '' : value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

// ---------------------------------------------------------------------------
// Package helpers
// ---------------------------------------------------------------------------

function partExists(zip, name) {
  return Boolean(zip.file(name))
}

async function readPart(zip, name) {
  const entry = zip.file(name)
  if (!entry) return null
  const size = Number(entry._data && entry._data.uncompressedSize)
  if (Number.isFinite(size) && size > MAX_PART_BYTES) throw new RangeError(`${name} is too large.`)
  const text = await entry.async('string')
  if (text.length > MAX_PART_BYTES) throw new RangeError(`${name} is too large.`)
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text
}

function relsPathFor(partName) {
  const directory = path.posix.dirname(partName)
  return `${directory === '.' ? '' : `${directory}/`}_rels/${path.posix.basename(partName)}.rels`
}

function resolveTarget(sourcePart, target) {
  if (!target) return null
  if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return null
  const clean = target.split('#')[0]
  if (clean.startsWith('/')) return path.posix.normalize(clean.slice(1))
  const base = path.posix.dirname(sourcePart)
  const resolved = path.posix.normalize(path.posix.join(base === '.' ? '' : base, clean))
  return resolved.startsWith('..') ? null : resolved
}

function relativeTarget(fromPart, toPart) {
  const relative = path.posix.relative(path.posix.dirname(fromPart), toPart)
  return relative || path.posix.basename(toPart)
}

async function readRelationships(zip, partName) {
  const xml = await readPart(zip, relsPathFor(partName))
  if (!xml) return []
  const doc = documentElement(parseXml(xml))
  return childrenOf(doc, 'pr:Relationship').map((node) => ({
    id: node.attrs.Id,
    type: node.attrs.Type,
    target: node.attrs.Target,
    external: node.attrs.TargetMode === 'External',
    resolved: node.attrs.TargetMode === 'External' ? null : resolveTarget(partName, node.attrs.Target),
  }))
}

function renderRelationships(relationships) {
  return `${XML_DECLARATION}<Relationships xmlns="${NS.pr}">${relationships.map((rel) => `<Relationship Id="${escapeXml(rel.id)}" Type="${escapeXml(rel.type)}" Target="${escapeXml(rel.target)}"${rel.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`
}

async function contentTypesOf(zip) {
  const xml = await readPart(zip, '[Content_Types].xml')
  const overrides = new Map()
  const defaults = new Map()
  if (xml) {
    const doc = documentElement(parseXml(xml))
    for (const node of childrenOf(doc, 'ct:Override')) overrides.set(String(node.attrs.PartName || '').replace(/^\//, '').toLowerCase(), node.attrs.ContentType)
    for (const node of childrenOf(doc, 'ct:Default')) defaults.set(String(node.attrs.Extension || '').toLowerCase(), node.attrs.ContentType)
  }
  return { xml, overrides, defaults }
}

function contentTypeFor(types, partName) {
  const override = types.overrides.get(partName.toLowerCase())
  if (override) return override
  const extension = path.posix.extname(partName).slice(1).toLowerCase()
  return types.defaults.get(extension) || DEFAULT_EXTENSION_TYPES[extension]
}

/** Sheets in workbook order: { name, sheetId, part } (worksheets only). */
async function workbookSheetParts(zip) {
  const xml = await readPart(zip, 'xl/workbook.xml')
  if (!xml) return []
  const doc = documentElement(parseXml(xml))
  const rels = await readRelationships(zip, 'xl/workbook.xml')
  const byId = new Map(rels.map((rel) => [rel.id, rel]))
  const sheets = childrenOf(child(doc, 'x:sheets'), 'x:sheet')
  return sheets.map((node, index) => {
    const rel = byId.get(node.attrs['r:id'])
    return { name: node.attrs.name, sheetId: Number(node.attrs.sheetId), index, part: rel && rel.resolved, type: rel && rel.type }
  }).filter((item) => item.part && /\/worksheet$/.test(String(item.type || '')))
}

// ---------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------

function hexToRgb(hex) {
  const value = parseInt(hex, 16)
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255]
}

function rgbToHex(rgb) {
  return rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase()
}

function rgbToHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255
  const max = Math.max(r, g, b), min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
  return [h / 6, s, l]
}

function hslToRgb([h, s, l]) {
  if (s === 0) return [l * 255, l * 255, l * 255]
  const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255]
}

const PRESET_COLORS = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080', darkBlue: '00008B', darkRed: '8B0000', darkGreen: '006400', ltGray: 'D3D3D3', dkGray: 'A9A9A9' }
const SCHEME_INDEX = { lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11 }

function colorNodeToHex(node, theme) {
  if (!node) return undefined
  let hex
  if (node.name === 'a:srgbClr') hex = /^[0-9a-f]{6}$/i.test(node.attrs.val || '') ? node.attrs.val.toUpperCase() : undefined
  else if (node.name === 'a:schemeClr') {
    const index = SCHEME_INDEX[node.attrs.val]
    hex = index === undefined ? undefined : (theme[index] || OFFICE_THEME[index])
  } else if (node.name === 'a:sysClr') hex = node.attrs.lastClr || (node.attrs.val === 'window' ? 'FFFFFF' : '000000')
  else if (node.name === 'a:prstClr') hex = PRESET_COLORS[node.attrs.val]
  else if (node.name === 'a:scrgbClr') hex = rgbToHex(['r', 'g', 'b'].map((key) => (Number(node.attrs[key]) || 0) / 100000 * 255))
  else if (node.name === 'a:hslClr') hex = rgbToHex(hslToRgb([(Number(node.attrs.hue) || 0) / 21600000, (Number(node.attrs.sat) || 0) / 100000, (Number(node.attrs.lum) || 0) / 100000]))
  if (!hex || !/^[0-9A-F]{6}$/i.test(hex)) return undefined
  let rgb = hexToRgb(hex)
  for (const mod of node.children) {
    const value = Number(mod.attrs.val) / 100000
    if (!Number.isFinite(value)) continue
    if (mod.name === 'a:lumMod' || mod.name === 'a:lumOff') {
      const hsl = rgbToHsl(rgb)
      hsl[2] = mod.name === 'a:lumMod' ? hsl[2] * value : hsl[2] + value
      hsl[2] = Math.max(0, Math.min(1, hsl[2]))
      rgb = hslToRgb(hsl)
    } else if (mod.name === 'a:satMod') {
      const hsl = rgbToHsl(rgb); hsl[1] = Math.max(0, Math.min(1, hsl[1] * value)); rgb = hslToRgb(hsl)
    } else if (mod.name === 'a:tint') rgb = rgb.map((c) => c * value + 255 * (1 - value))
    else if (mod.name === 'a:shade') rgb = rgb.map((c) => c * value)
  }
  return `#${rgbToHex(rgb)}`
}

const COLOR_ELEMENTS = new Set(['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr', 'a:scrgbClr', 'a:hslClr'])

/** Fill colour of an element holding a:solidFill/a:gradFill/a:noFill: '#RRGGBB', 'none' or undefined. */
function fillOf(node, theme) {
  if (!node) return undefined
  if (child(node, 'a:noFill')) return 'none'
  const solid = child(node, 'a:solidFill')
  if (solid) return colorNodeToHex(solid.children.find((item) => COLOR_ELEMENTS.has(item.name)), theme)
  const gradient = child(node, 'a:gradFill')
  if (gradient) {
    const stop = findDescendant(gradient, 'a:gs')
    return stop ? colorNodeToHex(stop.children.find((item) => COLOR_ELEMENTS.has(item.name)), theme) : undefined
  }
  const pattern = child(node, 'a:pattFill')
  if (pattern) return colorNodeToHex((child(pattern, 'a:fgClr') || { children: [] }).children.find((item) => COLOR_ELEMENTS.has(item.name)), theme)
  return undefined
}

// ---------------------------------------------------------------------------
// Chart XML -> model
// ---------------------------------------------------------------------------

function richText(rich) {
  if (!rich) return ''
  return childrenOf(rich, 'a:p').map((paragraph) => paragraph.children
    .filter((item) => item.name === 'a:r' || item.name === 'a:fld' || item.name === 'a:br')
    .map((item) => (item.name === 'a:br' ? '\n' : textContent(child(item, 'a:t'))))
    .join('')).join('\n')
}

function strCacheValues(cache) {
  if (!cache) return []
  const count = Math.min(MAX_CACHE_POINTS, Math.max(0, Number(valOf(cache, 'c:ptCount', 0)) || 0))
  const values = new Array(count).fill('')
  for (const point of childrenOf(cache, 'c:pt')) {
    const index = Number(point.attrs.idx)
    if (Number.isInteger(index) && index >= 0 && index < MAX_CACHE_POINTS) {
      if (index >= values.length) values.length = index + 1
      values[index] = textContent(child(point, 'c:v'))
    }
  }
  for (let index = 0; index < values.length; index += 1) if (values[index] === undefined) values[index] = ''
  return values
}

function numCacheValues(cache) {
  if (!cache) return { values: [], formatCode: undefined }
  const count = Math.min(MAX_CACHE_POINTS, Math.max(0, Number(valOf(cache, 'c:ptCount', 0)) || 0))
  const values = new Array(count).fill(null)
  for (const point of childrenOf(cache, 'c:pt')) {
    const index = Number(point.attrs.idx)
    const value = Number(textContent(child(point, 'c:v')))
    if (Number.isInteger(index) && index >= 0 && index < MAX_CACHE_POINTS) {
      while (values.length <= index) values.push(null)
      values[index] = Number.isFinite(value) ? value : null
    }
  }
  const formatCode = textContent(child(cache, 'c:formatCode')) || undefined
  return { values, formatCode }
}

/** c:tx / c:cat / c:val / c:xVal / c:yVal data source. */
function dataSource(node) {
  if (!node) return null
  const strRef = child(node, 'c:strRef')
  if (strRef) return { ref: textContent(child(strRef, 'c:f')).trim() || undefined, texts: strCacheValues(child(strRef, 'c:strCache')) }
  const numRef = child(node, 'c:numRef')
  if (numRef) {
    const cache = numCacheValues(child(numRef, 'c:numCache'))
    return { ref: textContent(child(numRef, 'c:f')).trim() || undefined, numbers: cache.values, formatCode: cache.formatCode }
  }
  const multi = child(node, 'c:multiLvlStrRef')
  if (multi) {
    const cache = child(multi, 'c:multiLvlStrCache')
    const levels = childrenOf(cache, 'c:lvl')
    return { ref: textContent(child(multi, 'c:f')).trim() || undefined, texts: levels.length ? strCacheValues(levels[0]) : [] }
  }
  const strLit = child(node, 'c:strLit')
  if (strLit) return { texts: strCacheValues(strLit) }
  const numLit = child(node, 'c:numLit')
  if (numLit) { const cache = numCacheValues(numLit); return { numbers: cache.values, formatCode: cache.formatCode } }
  const literal = child(node, 'c:v')
  if (literal) return { texts: [textContent(literal)] }
  return null
}

const LABEL_POSITIONS = { outEnd: 'outEnd', inEnd: 'inEnd', ctr: 'center', inBase: 'inBase', t: 'above', b: 'below', l: 'left', r: 'right', bestFit: 'bestFit' }

function parseDataLabels(node) {
  if (!node) return undefined
  if (boolVal(node, 'c:delete', false)) return null
  const labels = {}
  if (boolVal(node, 'c:showVal', false)) labels.showValue = true
  if (boolVal(node, 'c:showCatName', false)) labels.showCategory = true
  if (boolVal(node, 'c:showSerName', false)) labels.showSeriesName = true
  if (boolVal(node, 'c:showPercent', false)) labels.showPercent = true
  const position = LABEL_POSITIONS[valOf(node, 'c:dLblPos')]
  if (position) labels.position = position
  const numFmt = child(node, 'c:numFmt')
  if (numFmt && numFmt.attrs.sourceLinked !== '1' && numFmt.attrs.formatCode) labels.numFmt = numFmt.attrs.formatCode
  return Object.keys(labels).length ? labels : null
}

const MARKERS = new Set(['circle', 'dash', 'diamond', 'dot', 'none', 'picture', 'plus', 'square', 'star', 'triangle', 'x', 'auto'])

function parseTitle(titleNode, theme) {
  if (!titleNode) return null
  const tx = child(titleNode, 'c:tx')
  const result = { text: undefined, ref: undefined, size: undefined }
  if (tx) {
    const rich = child(tx, 'c:rich')
    if (rich) {
      result.text = richText(rich)
      const run = findDescendant(rich, 'a:rPr') || findDescendant(rich, 'a:defRPr')
      const size = run && Number(run.attrs.sz)
      if (Number.isFinite(size) && size > 0) result.size = size / 100
    }
    const strRef = child(tx, 'c:strRef')
    if (strRef) {
      result.ref = textContent(child(strRef, 'c:f')).trim() || undefined
      result.text = strCacheValues(child(strRef, 'c:strCache')).join(' ')
    }
  } else {
    const defaults = findDescendant(child(titleNode, 'c:txPr'), 'a:defRPr')
    const size = defaults && Number(defaults.attrs.sz)
    if (Number.isFinite(size) && size > 0) result.size = size / 100
  }
  void theme
  return result
}

function parseAxis(node, theme) {
  if (!node) return undefined
  const axis = {}
  const title = parseTitle(child(node, 'c:title'), theme)
  if (title && title.text) axis.title = title.text
  const numFmt = child(node, 'c:numFmt')
  if (numFmt && numFmt.attrs.sourceLinked !== '1' && numFmt.attrs.formatCode && numFmt.attrs.formatCode !== 'General') axis.numFmt = numFmt.attrs.formatCode
  const scaling = child(node, 'c:scaling')
  const min = numVal(scaling, 'c:min'), max = numVal(scaling, 'c:max'), logBase = numVal(scaling, 'c:logBase')
  if (min !== undefined) axis.min = min
  if (max !== undefined) axis.max = max
  if (logBase !== undefined && logBase >= 2) axis.logBase = logBase
  if (valOf(scaling, 'c:orientation') === 'maxMin') axis.reverse = true
  const majorUnit = numVal(node, 'c:majorUnit')
  if (majorUnit !== undefined && majorUnit > 0) axis.majorUnit = majorUnit
  axis.gridlines = Boolean(child(node, 'c:majorGridlines'))
  if (child(node, 'c:minorGridlines')) axis.minorGridlines = true
  if (boolVal(node, 'c:delete', false) || valOf(node, 'c:tickLblPos') === 'none') axis.visible = false
  const bodyPr = findDescendant(child(node, 'c:txPr'), 'a:bodyPr')
  const rotation = bodyPr && Number(bodyPr.attrs.rot)
  if (Number.isFinite(rotation) && rotation !== 0 && Math.abs(rotation) <= 5_400_000) axis.labelRotation = Math.round(rotation / 60000)
  return axis
}

const GROUP_KINDS = {
  'c:barChart': 'bar', 'c:bar3DChart': 'bar', 'c:lineChart': 'line', 'c:line3DChart': 'line', 'c:areaChart': 'area', 'c:area3DChart': 'area',
  'c:pieChart': 'pie', 'c:pie3DChart': 'pie', 'c:ofPieChart': 'pie', 'c:doughnutChart': 'doughnut', 'c:scatterChart': 'scatter', 'c:radarChart': 'radar',
  'c:bubbleChart': 'bubble', 'c:stockChart': 'stock', 'c:surfaceChart': 'surface', 'c:surface3DChart': 'surface',
}

function parseSeries(ser, group, theme, index) {
  const series = { id: `series-${index + 1}` }
  const name = dataSource(child(ser, 'c:tx'))
  if (name) {
    if (name.ref) series.nameRef = name.ref
    const text = (name.texts || []).join(' ') || (name.numbers || []).filter((v) => v !== null).join(' ')
    if (text) series.name = text
  }
  const spPr = child(ser, 'c:spPr')
  const lineKind = group.kind === 'line' || group.kind === 'scatter' || group.kind === 'radar'
  const line = child(spPr, 'a:ln')
  const fill = fillOf(spPr, theme)
  const lineFill = line ? fillOf(line, theme) : undefined
  if (lineKind) {
    if (lineFill && lineFill !== 'none') series.color = lineFill
    else if (fill && fill !== 'none') series.color = fill
    if (lineFill === 'none') series.showLine = false
    const width = line && Number(line.attrs.w)
    if (Number.isFinite(width) && width > 0) series.lineWidth = Math.round((width / 12700) * 100) / 100
  } else if (fill && fill !== 'none') series.color = fill
  else if (fill === 'none' && lineFill && lineFill !== 'none') series.color = lineFill
  const marker = child(ser, 'c:marker')
  if (marker) {
    const symbol = valOf(marker, 'c:symbol')
    if (symbol && MARKERS.has(symbol)) series.marker = symbol === 'picture' ? 'auto' : symbol
    else if (lineKind) series.marker = 'auto'
    const size = numVal(marker, 'c:size')
    if (size !== undefined && size >= 2 && size <= 72) series.markerSize = size
    if (!series.color) { const markerFill = fillOf(child(marker, 'c:spPr'), theme); if (markerFill && markerFill !== 'none') series.color = markerFill }
  } else if (group.kind === 'line' || group.kind === 'radar') {
    series.marker = group.marker === false ? 'none' : 'auto'
  } else if (group.kind === 'scatter') {
    series.marker = group.scatterStyle === 'line' || group.scatterStyle === 'smooth' ? 'none' : 'auto'
  }
  if (group.kind === 'scatter') {
    if (group.scatterStyle === 'marker' || group.scatterStyle === 'none') series.showLine = false
    else if (series.showLine !== false) series.showLine = true
  }
  const pointColors = {}
  for (const point of childrenOf(ser, 'c:dPt')) {
    const pointIndex = Number(valOf(point, 'c:idx'))
    const color = fillOf(child(point, 'c:spPr'), theme)
    if (Number.isInteger(pointIndex) && pointIndex >= 0 && color && color !== 'none') pointColors[String(pointIndex)] = color
  }
  if (Object.keys(pointColors).length) series.pointColors = pointColors
  const labels = parseDataLabels(child(ser, 'c:dLbls'))
  if (labels) series.dataLabels = labels
  else if (labels === undefined && group.dataLabels) series.dataLabels = group.dataLabels
  const categories = dataSource(child(ser, group.kind === 'scatter' ? 'c:xVal' : 'c:cat'))
  const values = dataSource(child(ser, group.kind === 'scatter' ? 'c:yVal' : 'c:val'))
  if (group.kind === 'scatter') {
    if (categories) {
      if (categories.ref) series.xValuesRef = categories.ref
      if (categories.numbers) series.xValuesCache = categories.numbers
      else if (categories.texts) series.categoriesCache = categories.texts
    }
  } else if (categories) {
    if (categories.ref) series.categoriesRef = categories.ref
    const texts = categories.texts || (categories.numbers || []).map((value) => (value === null ? '' : String(value)))
    if (texts.length) series.categoriesCache = texts
  }
  if (values) {
    if (values.ref) series.valuesRef = values.ref
    if (values.numbers) series.valuesCache = values.numbers
    else if (values.texts) series.valuesCache = values.texts.map((text) => (Number.isFinite(Number(text)) && text !== '' ? Number(text) : null))
    if (values.formatCode && values.formatCode !== 'General') series.valuesNumFmt = values.formatCode
  }
  if (boolVal(ser, 'c:smooth', false)) series.smooth = true
  if (boolVal(ser, 'c:invertIfNegative', false)) series.invertIfNegative = true
  const order = Number(valOf(ser, 'c:order'))
  return { series, order: Number.isFinite(order) ? order : index }
}

function parseChartXml(xml, theme) {
  const doc = documentElement(parseXml(xml))
  if (!doc || doc.name !== 'c:chartSpace') throw new Error('Not a DrawingML chart part.')
  const chartNode = child(doc, 'c:chart')
  const plotArea = child(chartNode, 'c:plotArea')
  const chart = { series: [] }
  const title = parseTitle(child(chartNode, 'c:title'), theme)
  if (title) {
    if (title.text !== undefined && title.text !== '') chart.title = title.text
    if (title.ref) chart.titleRef = title.ref
    if (title.size && Math.abs(title.size - 14) > 0.01) chart.style = { ...(chart.style || {}), titleFontSize: title.size }
  }
  if (boolVal(chartNode, 'c:autoTitleDeleted', false) && !title) chart.autoTitleDeleted = true
  const axisNodes = new Map()
  for (const node of plotArea ? plotArea.children : []) {
    if (['c:catAx', 'c:valAx', 'c:dateAx', 'c:serAx'].includes(node.name)) axisNodes.set(valOf(node, 'c:axId'), node)
  }
  const groups = []
  for (const node of plotArea ? plotArea.children : []) {
    const kind = GROUP_KINDS[node.name]
    if (!kind) continue
    const group = {
      node,
      kind,
      barDir: valOf(node, 'c:barDir', 'col'),
      grouping: valOf(node, 'c:grouping', 'clustered'),
      varyColors: boolVal(node, 'c:varyColors', false),
      axIds: childrenOf(node, 'c:axId').map((item) => item.attrs.val),
      threeD: /3D/.test(node.name),
      marker: child(node, 'c:marker') ? boolVal(node, 'c:marker', true) : true,
      scatterStyle: valOf(node, 'c:scatterStyle', 'lineMarker'),
      dataLabels: parseDataLabels(child(node, 'c:dLbls')) || undefined,
    }
    groups.push(group)
  }
  if (!groups.length) return { ...chart, type: 'unsupported', unsupportedKind: 'empty' }
  const unsupported = groups.find((group) => ['bubble', 'stock', 'surface'].includes(group.kind))
  if (unsupported) return { ...chart, type: 'unsupported', unsupportedKind: unsupported.kind }
  const primary = groups[0]
  const primaryValueAxis = primary.axIds[1]
  const kindOf = (group) => (group.kind === 'bar' ? (group.barDir === 'bar' ? 'bar' : 'column') : group.kind)
  const kinds = [...new Set(groups.map(kindOf))]
  const secondaryGroups = groups.filter((group) => group !== primary && group.axIds.length >= 2 && group.axIds[1] !== primaryValueAxis && !['pie', 'doughnut'].includes(group.kind))
  const combo = kinds.length > 1 || secondaryGroups.length > 0
  const allSeries = []
  let seriesCounter = 0
  for (const group of groups) {
    for (const ser of childrenOf(group.node, 'c:ser')) {
      const parsed = parseSeries(ser, group, theme, seriesCounter)
      seriesCounter += 1
      if (combo) {
        const kind = kindOf(group)
        parsed.series.type = kind === 'bar' ? 'column' : ['column', 'line', 'area', 'scatter'].includes(kind) ? kind : 'line'
        if (secondaryGroups.includes(group)) parsed.series.secondaryAxis = true
      }
      allSeries.push(parsed)
    }
    if (allSeries.length > 1000) break
  }
  allSeries.sort((a, b) => a.order - b.order)
  chart.series = allSeries.map((item, index) => ({ ...item.series, id: `series-${index + 1}` }))
  chart.type = combo ? 'combo' : kindOf(primary)
  if (['column', 'bar', 'line', 'area'].includes(chart.type) || combo) {
    const grouping = primary.grouping === 'stacked' ? 'stacked' : primary.grouping === 'percentStacked' ? 'percentStacked' : 'clustered'
    chart.grouping = grouping
  }
  if (primary.varyColors && ['column', 'bar', 'line'].includes(chart.type) && chart.series.length === 1) chart.varyColors = true
  if (primary.threeD) chart.threeD = true
  const barGroup = groups.find((group) => group.kind === 'bar')
  if (barGroup) {
    const gapWidth = numVal(barGroup.node, 'c:gapWidth')
    const overlap = numVal(barGroup.node, 'c:overlap')
    if (gapWidth !== undefined) chart.gapWidth = gapWidth
    if (overlap !== undefined) chart.overlap = overlap
  }
  const pieGroup = groups.find((group) => group.kind === 'pie' || group.kind === 'doughnut')
  if (pieGroup) {
    const angle = numVal(pieGroup.node, 'c:firstSliceAng')
    if (angle) chart.firstSliceAngle = angle
    const hole = numVal(pieGroup.node, 'c:holeSize')
    if (hole !== undefined) chart.holeSize = hole
    if (chart.type === 'pie') chart.series = chart.series.slice(0, 1)
  }
  // Axes: x = category (or scatter X) axis, y = primary value axis, y2 = secondary value axis.
  if (!['pie', 'doughnut'].includes(chart.type)) {
    const axes = {}
    let xNode = axisNodes.get(primary.axIds[0])
    let yNode = axisNodes.get(primary.axIds[1])
    if (primary.kind === 'scatter' && xNode && yNode) {
      const xPos = valOf(xNode, 'c:axPos'), yPos = valOf(yNode, 'c:axPos')
      if ((xPos === 'l' || xPos === 'r') && (yPos === 'b' || yPos === 't')) [xNode, yNode] = [yNode, xNode]
    }
    if (xNode) axes.x = parseAxis(xNode, theme)
    if (yNode) axes.y = parseAxis(yNode, theme)
    if (secondaryGroups.length) {
      const y2Node = axisNodes.get(secondaryGroups[0].axIds[1])
      if (y2Node) axes.y2 = parseAxis(y2Node, theme)
    }
    chart.axes = axes
  }
  // Legend
  const legend = child(chartNode, 'c:legend')
  const position = legend ? valOf(legend, 'c:legendPos', 'r') : null
  chart.legend = !legend ? 'none' : ({ r: 'right', l: 'left', t: 'top', b: 'bottom', tr: 'right' })[position] || 'right'
  chart.plotVisibleOnly = boolVal(chartNode, 'c:plotVisOnly', true)
  const blanks = valOf(chartNode, 'c:dispBlanksAs')
  if (blanks === 'zero' || blanks === 'span' || blanks === 'gap') chart.displayBlanksAs = blanks
  // Chart area and text defaults
  const style = { ...(chart.style || {}) }
  const areaFill = fillOf(child(doc, 'c:spPr'), theme)
  if (areaFill === 'none') style.background = 'transparent'
  else if (areaFill && areaFill !== '#FFFFFF') style.background = areaFill
  const areaLine = child(child(doc, 'c:spPr'), 'a:ln')
  const border = areaLine ? fillOf(areaLine, theme) : undefined
  if (border === 'none') style.border = null
  else if (border && border !== '#D9D9D9') style.border = border
  const defaults = findDescendant(child(doc, 'c:txPr'), 'a:defRPr')
  if (defaults) {
    const size = Number(defaults.attrs.sz)
    if (Number.isFinite(size) && size > 0 && size !== 900 && size !== 1000) style.fontSize = size / 100
    const color = fillOf(defaults, theme)
    if (color && color !== 'none' && color !== '#595959') style.textColor = color
    const latin = child(defaults, 'a:latin')
    if (latin && latin.attrs.typeface && !latin.attrs.typeface.startsWith('+')) style.fontFamily = latin.attrs.typeface
  }
  const rounded = child(doc, 'c:roundedCorners')
  style.roundedCorners = rounded ? boolVal(doc, 'c:roundedCorners', true) : true
  if (!style.roundedCorners) delete style.roundedCorners
  if (Object.keys(style).length) chart.style = style
  return chart
}

function parseChartExXml(xml) {
  const doc = documentElement(parseXml(xml))
  const chart = { type: 'unsupported', series: [], legend: 'none' }
  const chartNode = child(doc, 'cx:chart')
  const title = child(chartNode, 'cx:title')
  if (title) {
    const text = textContent(findDescendant(title, 'cx:v')) || richText(findDescendant(title, 'cx:rich'))
    if (text) chart.title = text
  }
  const series = findDescendant(child(chartNode, 'cx:plotArea'), 'cx:series')
  chart.unsupportedKind = (series && series.attrs.layoutId) || 'chartex'
  return chart
}

// ---------------------------------------------------------------------------
// Fingerprint (detects edits made without `modified: true`)
// ---------------------------------------------------------------------------

// Fields that never change the chart part (placement, bookkeeping, editor state, caches).
const FINGERPRINT_IGNORED_CHART = new Set(['id', 'anchor', 'name', 'description', 'sourcePart', 'sourceInfo', 'modified', 'dataRange', 'seriesIn', 'firstRowHeaders', 'firstColumnLabels'])
const FINGERPRINT_IGNORED_SERIES = new Set(['id', 'categoriesCache', 'valuesCache', 'xValuesCache'])

function canonical(value, depth = 0) {
  if (depth > 12) return null
  if (Array.isArray(value)) return value.map((item) => canonical(item, depth + 1))
  if (value && typeof value === 'object') {
    const output = {}
    for (const key of Object.keys(value).sort()) if (value[key] !== undefined) output[key] = canonical(value[key], depth + 1)
    return output
  }
  return value
}

function chartFingerprint(chart) {
  const projection = {}
  for (const key of Object.keys(chart || {})) if (!FINGERPRINT_IGNORED_CHART.has(key)) projection[key] = chart[key]
  if (Array.isArray(projection.series)) {
    projection.series = projection.series.map((item) => {
      const output = {}
      for (const key of Object.keys(item || {})) if (!FINGERPRINT_IGNORED_SERIES.has(key)) output[key] = item[key]
      return output
    })
  }
  return crypto.createHash('sha1').update(JSON.stringify(canonical(projection))).digest('hex').slice(0, 20)
}

// ---------------------------------------------------------------------------
// Anchors
// ---------------------------------------------------------------------------

function excelColumnPixels(width) {
  return Math.trunc(((256 * width + Math.trunc(128 / 7)) / 256) * 7)
}

function sheetMetrics(sheet) {
  const defaultWidth = Number(sheet && sheet.properties && sheet.properties.defaultColWidth) || 8.43
  const defaultHeight = Number(sheet && sheet.properties && sheet.properties.defaultRowHeight) || 15
  const hiddenCols = new Set((sheet && sheet.hiddenCols) || [])
  const hiddenRows = new Set((sheet && sheet.hiddenRows) || [])
  return {
    col(index) { if (hiddenCols.has(index + 1)) return 0; const width = Number(sheet && sheet.colWidths && sheet.colWidths[String(index + 1)]); return excelColumnPixels(Number.isFinite(width) && width > 0 ? width : defaultWidth) * EMU_PER_PIXEL },
    row(index) { if (hiddenRows.has(index + 1)) return 0; const height = Number(sheet && sheet.rowHeights && sheet.rowHeights[String(index + 1)]); return Math.round((Number.isFinite(height) && height > 0 ? height : defaultHeight) * 12700) },
  }
}

function locate(offsetEmu, sizeOf, max) {
  let index = 0
  let remaining = Math.max(0, offsetEmu)
  while (index < max - 1) {
    const size = sizeOf(index)
    if (remaining < size || index > 200_000) break
    remaining -= size
    index += 1
  }
  return { index, offset: Math.round(remaining) }
}

function positionOf(node) {
  const number = (name) => Math.max(0, Math.trunc(Number(textContent(child(node, name))) || 0))
  return { col: Math.min(MAX_COLS - 1, number('xdr:col')), colOffsetEmu: number('xdr:colOff'), row: Math.min(MAX_ROWS - 1, number('xdr:row')), rowOffsetEmu: number('xdr:rowOff') }
}

function anchorFromNode(anchorNode, sheet) {
  const metrics = sheetMetrics(sheet)
  const extentOf = (node) => { const ext = child(node, 'xdr:ext'); return { cx: Math.max(0, Number(ext && ext.attrs.cx) || 0), cy: Math.max(0, Number(ext && ext.attrs.cy) || 0) } }
  const endFrom = (from, cx, cy) => {
    let startX = 0
    for (let col = 0; col < from.col && col < 20_000; col += 1) startX += metrics.col(col)
    let startY = 0
    for (let row = 0; row < from.row && row < 200_000; row += 1) startY += metrics.row(row)
    const x = locate(startX + from.colOffsetEmu + cx, metrics.col, MAX_COLS)
    const y = locate(startY + from.rowOffsetEmu + cy, metrics.row, MAX_ROWS)
    return { col: x.index, colOffsetEmu: x.offset, row: y.index, rowOffsetEmu: y.offset }
  }
  if (anchorNode.name === 'xdr:twoCellAnchor') {
    const from = positionOf(child(anchorNode, 'xdr:from'))
    const to = positionOf(child(anchorNode, 'xdr:to'))
    const editAs = anchorNode.attrs.editAs
    const anchor = { from, to }
    if (editAs === 'oneCell' || editAs === 'absolute') anchor.editAs = editAs
    return anchor
  }
  if (anchorNode.name === 'xdr:oneCellAnchor') {
    const from = positionOf(child(anchorNode, 'xdr:from'))
    const { cx, cy } = extentOf(anchorNode)
    return { from, to: endFrom(from, cx, cy), editAs: 'oneCell' }
  }
  if (anchorNode.name === 'xdr:absoluteAnchor') {
    const pos = child(anchorNode, 'xdr:pos')
    const x = Math.max(0, Number(pos && pos.attrs.x) || 0), y = Math.max(0, Number(pos && pos.attrs.y) || 0)
    const { cx, cy } = extentOf(anchorNode)
    const start = { ...(() => { const a = locate(x, metrics.col, MAX_COLS), b = locate(y, metrics.row, MAX_ROWS); return { col: a.index, colOffsetEmu: a.offset, row: b.index, rowOffsetEmu: b.offset } })() }
    return { from: start, to: endFrom(start, cx, cy), editAs: 'absolute' }
  }
  return null
}

const ANCHOR_NAMES = new Set(['xdr:twoCellAnchor', 'xdr:oneCellAnchor', 'xdr:absoluteAnchor'])

/** Top-level drawing objects: { index, anchorNode, content } where content holds the graphic frame. */
function drawingObjects(wsDr) {
  const objects = []
  let index = 0
  for (const node of wsDr ? wsDr.children : []) {
    let anchorNode = node
    if (node.name === 'mc:AlternateContent') {
      const choice = child(node, 'mc:Choice') || child(node, 'mc:Fallback')
      anchorNode = choice && choice.children.find((item) => ANCHOR_NAMES.has(item.name))
    }
    if (!anchorNode || !ANCHOR_NAMES.has(anchorNode.name)) { if (ANCHOR_NAMES.has(node.name) || node.name === 'mc:AlternateContent') index += 1; continue }
    objects.push({ index, wrapper: node, anchorNode })
    index += 1
  }
  return objects
}

/** Chart reference inside an anchor: { kind, rId, content } where content is the element to copy verbatim. */
function chartReferenceOf(anchorNode) {
  for (const content of anchorNode.children) {
    if (content.name === 'xdr:graphicFrame') {
      const chartRef = findDescendant(content, 'c:chart')
      if (chartRef && chartRef.attrs['r:id']) return { kind: 'chart', rId: chartRef.attrs['r:id'], content, frame: content }
      const chartEx = findDescendant(content, 'cx:chart')
      if (chartEx && chartEx.attrs['r:id']) return { kind: 'chartex', rId: chartEx.attrs['r:id'], content, frame: content }
    }
    if (content.name === 'mc:AlternateContent') {
      for (const branch of content.children) {
        const frame = branch.children.find((item) => item.name === 'xdr:graphicFrame')
        if (!frame) continue
        const chartRef = findDescendant(frame, 'c:chart')
        if (chartRef && chartRef.attrs['r:id']) return { kind: 'chart', rId: chartRef.attrs['r:id'], content, frame }
        const chartEx = findDescendant(frame, 'cx:chart')
        if (chartEx && chartEx.attrs['r:id']) return { kind: 'chartex', rId: chartEx.attrs['r:id'], content, frame }
      }
    }
  }
  return null
}

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

/**
 * Parse every chart of an OOXML package and attach them to the model sheets
 * (matched by sheetId, then name). Returns the number of charts imported.
 */
async function importChartsIntoSheets(zipOrBuffer, sheets, warnings, themeColors) {
  const JSZip = require('jszip')
  const zip = zipOrBuffer && typeof zipOrBuffer.file === 'function' ? zipOrBuffer : await JSZip.loadAsync(zipOrBuffer)
  const theme = Array.isArray(themeColors) && themeColors.length >= 10 ? themeColors.map((value, index) => (/^[0-9a-f]{6}$/i.test(String(value || '')) ? String(value).toUpperCase() : OFFICE_THEME[index])) : OFFICE_THEME
  const parts = await workbookSheetParts(zip)
  let imported = 0
  for (const sheetPart of parts) {
    const sheet = sheets.find((item) => Number.isInteger(item.sourceWorksheetId) && item.sourceWorksheetId === sheetPart.sheetId)
      || sheets.find((item) => item.sourceSheetName === sheetPart.name || item.name === sheetPart.name)
    if (!sheet) continue
    let rels
    try { rels = await readRelationships(zip, sheetPart.part) } catch { continue }
    const drawingRels = rels.filter((rel) => rel.type === REL.drawing && rel.resolved && partExists(zip, rel.resolved))
    const charts = []
    const unsupportedKinds = []
    let failures = 0
    for (const drawingRel of drawingRels) {
      let drawingXml, drawingDoc, drawingRelationships
      try {
        drawingXml = await readPart(zip, drawingRel.resolved)
        drawingDoc = documentElement(parseXml(drawingXml))
        drawingRelationships = await readRelationships(zip, drawingRel.resolved)
      } catch { failures += 1; continue }
      const relById = new Map(drawingRelationships.map((rel) => [rel.id, rel]))
      for (const object of drawingObjects(drawingDoc)) {
        const reference = chartReferenceOf(object.anchorNode)
        if (!reference) continue
        if (imported + charts.length >= MAX_CHARTS_PER_WORKBOOK) break
        const rel = relById.get(reference.rId)
        if (!rel || !rel.resolved || !partExists(zip, rel.resolved)) { failures += 1; continue }
        try {
          const chartXml = await readPart(zip, rel.resolved)
          const parsed = reference.kind === 'chartex' ? parseChartExXml(chartXml) : parseChartXml(chartXml, theme)
          const anchor = anchorFromNode(object.anchorNode, sheet)
          if (!anchor) { failures += 1; continue }
          const properties = findDescendant(reference.frame, 'xdr:cNvPr')
          const chart = {
            id: `chart-${sheet.id}-${charts.length + 1}`,
            ...parsed,
            anchor,
            sourcePart: rel.resolved,
          }
          if (properties && properties.attrs.name) chart.name = properties.attrs.name
          if (properties && properties.attrs.descr) chart.description = properties.attrs.descr
          if (chart.type === 'unsupported') unsupportedKinds.push(chart.unsupportedKind || 'unknown')
          chart.sourceInfo = { drawingPart: drawingRel.resolved, anchorIndex: object.index, kind: reference.kind, fingerprint: '' }
          chart.sourceInfo.fingerprint = chartFingerprint(chart)
          charts.push(chart)
        } catch {
          failures += 1
        }
      }
    }
    if (charts.length) {
      sheet.charts = [...(Array.isArray(sheet.charts) ? sheet.charts : []), ...charts]
      imported += charts.length
    }
    if (unsupportedKinds.length) {
      const kinds = [...new Set(unsupportedKinds)].join(', ')
      warnings.push(`Sheet "${sheet.name}" has ${unsupportedKinds.length === 1 ? 'a chart' : `${unsupportedKinds.length} charts`} (${kinds}) that simple_calc shows as a placeholder; ${unsupportedKinds.length === 1 ? 'it is' : 'they are'} kept when saving as .xlsx.`)
    }
    if (failures) warnings.push(`Sheet "${sheet.name}" has ${failures === 1 ? 'a chart' : `${failures} charts`} that could not be read and will not be kept when saving.`)
  }
  return imported
}

// ---------------------------------------------------------------------------
// Reading model data for chart caches
// ---------------------------------------------------------------------------

function columnIndex(label) {
  let value = 0
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value - 1
}

function columnLabel(index) {
  let value = index + 1
  let label = ''
  while (value > 0) { const remainder = (value - 1) % 26; label = String.fromCharCode(65 + remainder) + label; value = Math.floor((value - 1) / 26) }
  return label
}

function splitTopLevel(text) {
  const parts = []
  let current = '', quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === "'") { if (quoted && text[index + 1] === "'") { current += "''"; index += 1; continue } quoted = !quoted }
    if (!quoted && character === ',') { parts.push(current); current = '' } else current += character
  }
  parts.push(current)
  return parts
}

function parseRefAreas(ref, defaultSheet) {
  if (typeof ref !== 'string') return null
  let text = ref.trim().replace(/^=/, '')
  if (!text || /#REF!/i.test(text)) return null
  while (text.startsWith('(') && text.endsWith(')')) text = text.slice(1, -1)
  const areas = []
  let lastSheet = defaultSheet
  for (const piece of splitTopLevel(text)) {
    const bang = piece.lastIndexOf('!')
    let sheet = lastSheet, body = piece.trim()
    if (bang >= 0) {
      sheet = piece.slice(0, bang).trim()
      if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'")
      if (sheet.includes('[')) return null
      body = piece.slice(bang + 1).trim()
      lastSheet = sheet
    }
    const match = /^\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?$/i.exec(body)
    if (!match) return null
    const r1 = Number(match[2]) - 1, c1 = columnIndex(match[1])
    const r2 = match[4] ? Number(match[4]) - 1 : r1, c2 = match[3] ? columnIndex(match[3]) : c1
    areas.push({ sheet, top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: Math.min(c1, c2), right: Math.max(c1, c2) })
  }
  return areas.length ? areas : null
}

function quoteSheet(name) {
  return /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${String(name).replace(/'/g, "''")}'`
}

/** Normalise a reference to absolute A1 with a sheet name, as Excel writes it in c:f. */
function absoluteRef(ref, defaultSheet) {
  const areas = parseRefAreas(ref, defaultSheet)
  if (!areas) return typeof ref === 'string' ? ref.replace(/^=/, '') : ref
  const parts = areas.map((area) => {
    const start = `$${columnLabel(area.left)}$${area.top + 1}`
    const end = `$${columnLabel(area.right)}$${area.bottom + 1}`
    const body = area.top === area.bottom && area.left === area.right ? start : `${start}:${end}`
    return area.sheet ? `${quoteSheet(area.sheet)}!${body}` : body
  })
  return parts.length > 1 ? `(${parts.join(',')})` : parts[0]
}

function scalarOf(cell) {
  if (!cell || typeof cell !== 'object') return undefined
  let value = cell.formula ? (cell.result !== undefined ? cell.result : cell.value) : cell.value
  if (value && typeof value === 'object') {
    if (value.type === 'date') { const date = new Date(value.value); value = Number.isNaN(date.getTime()) ? undefined : date.getTime() / 86_400_000 + 25_569 } else if (value.type === 'richText') value = cell.display || ''
    else value = value.value
  }
  return value === null ? undefined : value
}

function modelReader(model) {
  const byName = new Map(model.sheets.map((sheet) => [String(sheet.name).toLowerCase(), sheet]))
  return function read(ref, defaultSheet) {
    const areas = parseRefAreas(ref, defaultSheet)
    if (!areas) return null
    const values = [], texts = [], formats = []
    for (const area of areas) {
      const sheet = byName.get(String(area.sheet || '').toLowerCase())
      if (!sheet) return null
      // Caches cover every cell of the reference (Excel applies plotVisOnly when drawing).
      const byColumn = area.bottom - area.top >= area.right - area.left
      const count = byColumn ? area.bottom - area.top + 1 : area.right - area.left + 1
      for (let offset = 0; offset < count && values.length < MAX_CACHE_POINTS; offset += 1) {
        const row = byColumn ? area.top + offset : area.bottom
        const col = byColumn ? area.right : area.left + offset
        const cell = sheet.cells && sheet.cells[`${columnLabel(col)}${row + 1}`]
        const value = scalarOf(cell)
        values.push(value)
        texts.push(cell && typeof cell.display === 'string' ? cell.display : value === undefined ? '' : String(value))
        formats.push(cell && (cell.numFmt || (cell.style && cell.style.numFmt)))
      }
    }
    return { values, texts, formats }
  }
}

// ---------------------------------------------------------------------------
// Model -> chart XML
// ---------------------------------------------------------------------------

function hex6(color, fallback) {
  const match = /^#?([0-9a-f]{6})/i.exec(String(color || ''))
  return match ? match[1].toUpperCase() : fallback
}

function paletteColor(index, palette) {
  const colors = palette && palette.length ? palette : OFFICE_ACCENTS
  const base = hex6(colors[index % colors.length], '4472C4')
  const cycle = Math.floor(index / colors.length) % 7
  const variants = [[1, 0], [0.6, 0], [0.6, 0.4], [0.8, 0], [0.8, 0.2], [0.5, 0], [0.4, 0.6]]
  if (!cycle) return base
  const hsl = rgbToHsl(hexToRgb(base))
  hsl[2] = Math.max(0, Math.min(1, hsl[2] * variants[cycle][0] + variants[cycle][1]))
  return rgbToHex(hslToRgb(hsl))
}

function solidFill(hex) { return `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>` }
const NO_FILL_SP = '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'

function textProperties(sizePt, color, options = {}) {
  const rot = options.rot ? ` rot="${options.rot}" vert="horz"` : ''
  const bold = options.bold ? ' b="1"' : ' b="0"'
  const latin = options.font ? `<a:latin typeface="${escapeXml(options.font)}"/>` : ''
  return `<c:txPr><a:bodyPr${rot}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round(sizePt * 100)}"${bold}>${solidFill(color)}${latin}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`
}

function richTitle(text, sizePt, color, options = {}) {
  const rot = options.rot ? ` rot="${options.rot}" vert="horz"` : ''
  const paragraphs = String(text).split(/\r?\n/).map((line) => `<a:p><a:pPr><a:defRPr sz="${Math.round(sizePt * 100)}" b="0">${solidFill(color)}</a:defRPr></a:pPr><a:r><a:rPr lang="en-US" sz="${Math.round(sizePt * 100)}" b="0">${solidFill(color)}</a:rPr><a:t>${escapeXml(line)}</a:t></a:r></a:p>`).join('')
  return `<c:title><c:tx><c:rich><a:bodyPr${rot}/><a:lstStyle/>${paragraphs}</c:rich></c:tx><c:overlay val="0"/>${NO_FILL_SP}</c:title>`
}

function refTitle(ref, text, sizePt, color) {
  return `<c:title><c:tx><c:strRef><c:f>${escapeXml(ref)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(text || '')}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:overlay val="0"/>${NO_FILL_SP}${textProperties(sizePt, color)}</c:title>`
}

function numberText(value) {
  if (!Number.isFinite(value)) return '0'
  return String(value)
}

function numCacheXml(values, formatCode, tag) {
  const points = []
  values.forEach((value, index) => { if (typeof value === 'number' && Number.isFinite(value)) points.push(`<c:pt idx="${index}"><c:v>${numberText(value)}</c:v></c:pt>`) })
  return `<c:${tag}><c:formatCode>${escapeXml(formatCode || 'General')}</c:formatCode><c:ptCount val="${values.length}"/>${points.join('')}</c:${tag}>`
}

function strCacheXml(texts, tag) {
  return `<c:${tag}><c:ptCount val="${texts.length}"/>${texts.map((text, index) => `<c:pt idx="${index}"><c:v>${escapeXml(text)}</c:v></c:pt>`).join('')}</c:${tag}>`
}

/** Build the c:cat/c:val-style content for a reference with a cache from the model (or the stored cache). */
function dataXml(tag, ref, read, defaultSheet, cache, numeric, formatHint) {
  const data = ref ? read(ref, defaultSheet) : null
  const f = ref ? `<c:f>${escapeXml(absoluteRef(ref, defaultSheet))}</c:f>` : ''
  if (numeric) {
    const values = data ? data.values.map((value) => (typeof value === 'number' && Number.isFinite(value) ? value : null)) : (cache || []).map((value) => (typeof value === 'number' ? value : null))
    const formatCode = formatHint || (data && data.formats.find(Boolean)) || 'General'
    if (!ref) return `<c:${tag}>${numCacheXml(values, formatCode, 'numLit')}</c:${tag}>`
    return `<c:${tag}><c:numRef>${f}${numCacheXml(values, formatCode, 'numCache')}</c:numRef></c:${tag}>`
  }
  // Categories: numeric when every non-empty cell is a number (dates stay numbers with their format).
  if (data && data.values.length && data.values.every((value) => value === undefined || typeof value === 'number') && data.values.some((value) => typeof value === 'number')) {
    return `<c:${tag}><c:numRef>${f}${numCacheXml(data.values.map((value) => (typeof value === 'number' ? value : null)), data.formats.find(Boolean) || 'General', 'numCache')}</c:numRef></c:${tag}>`
  }
  const texts = data ? data.texts : (cache || []).map((value) => String(value == null ? '' : value))
  if (!ref) return `<c:${tag}>${strCacheXml(texts, 'strLit')}</c:${tag}>`
  return `<c:${tag}><c:strRef>${f}${strCacheXml(texts, 'strCache')}</c:strRef></c:${tag}>`
}

function seriesNameXml(series, read, defaultSheet, fallback) {
  if (series.nameRef) {
    const data = read(series.nameRef, defaultSheet)
    const text = data ? data.texts.filter(Boolean).join(' ') : series.name || ''
    return `<c:tx><c:strRef><c:f>${escapeXml(absoluteRef(series.nameRef, defaultSheet))}</c:f>${strCacheXml([text || series.name || fallback], 'strCache')}</c:strRef></c:tx>`
  }
  return `<c:tx><c:v>${escapeXml(series.name || fallback)}</c:v></c:tx>`
}

const DLBL_POSITION = { outEnd: 'outEnd', inEnd: 'inEnd', center: 'ctr', inBase: 'inBase', above: 't', below: 'b', left: 'l', right: 'r', bestFit: 'bestFit' }

function allowedLabelPositions(kind, stacked) {
  if (kind === 'column' || kind === 'bar') return stacked ? ['ctr', 'inEnd', 'inBase'] : ['outEnd', 'inEnd', 'ctr', 'inBase']
  if (kind === 'line' || kind === 'scatter') return ['t', 'b', 'l', 'r', 'ctr']
  if (kind === 'pie') return ['bestFit', 'outEnd', 'inEnd', 'ctr']
  return []
}

function dataLabelsXml(labels, kind, stacked) {
  if (!labels || !(labels.showValue || labels.showCategory || labels.showSeriesName || labels.showPercent)) return ''
  const position = DLBL_POSITION[labels.position]
  const allowed = allowedLabelPositions(kind, stacked)
  const positionXml = position && allowed.includes(position) ? `<c:dLblPos val="${position}"/>` : ''
  const numFmt = labels.numFmt ? `<c:numFmt formatCode="${escapeXml(labels.numFmt)}" sourceLinked="0"/>` : ''
  const flag = (value) => (value ? '1' : '0')
  const percent = kind === 'pie' || kind === 'doughnut' ? labels.showPercent : false
  return `<c:dLbls>${numFmt}${NO_FILL_SP}${textProperties(9, '404040')}${positionXml}<c:showLegendKey val="0"/><c:showVal val="${flag(labels.showValue)}"/><c:showCatName val="${flag(labels.showCategory)}"/><c:showSerName val="${flag(labels.showSeriesName)}"/><c:showPercent val="${flag(percent)}"/><c:showBubbleSize val="0"/>${kind === 'pie' || kind === 'doughnut' ? '<c:showLeaderLines val="1"/>' : ''}</c:dLbls>`
}

const MARKER_SYMBOLS = new Set(['circle', 'dash', 'diamond', 'dot', 'none', 'plus', 'square', 'star', 'triangle', 'x'])

function markerXml(series, color) {
  const symbol = series.marker === 'auto' ? 'circle' : series.marker
  if (!symbol || symbol === 'none' || !MARKER_SYMBOLS.has(symbol)) return '<c:marker><c:symbol val="none"/></c:marker>'
  const size = Math.max(2, Math.min(72, Math.round(Number(series.markerSize) || 5)))
  return `<c:marker><c:symbol val="${symbol}"/><c:size val="${size}"/><c:spPr>${solidFill(color)}<a:ln w="9525">${solidFill(color)}</a:ln></c:spPr></c:marker>`
}

function lineSpPr(series, color) {
  const width = Math.round(Math.max(0.25, Math.min(20, Number(series.lineWidth) || 2.25)) * 12700)
  if (series.showLine === false) return `<c:spPr><a:ln w="${width}" cap="rnd"><a:noFill/><a:round/></a:ln></c:spPr>`
  return `<c:spPr><a:ln w="${width}" cap="rnd">${solidFill(color)}<a:round/></a:ln></c:spPr>`
}

function pointColorXml(series, kind, count, palette, vary) {
  const overrides = series.pointColors || {}
  const indices = vary ? Array.from({ length: count }, (_, index) => index) : Object.keys(overrides).map(Number).filter((index) => Number.isInteger(index) && index >= 0).sort((a, b) => a - b)
  return indices.map((index) => {
    const color = hex6(overrides[String(index)], vary ? paletteColor(index, palette) : undefined)
    if (!color) return ''
    if (kind === 'pie' || kind === 'doughnut') return `<c:dPt><c:idx val="${index}"/><c:bubble3D val="0"/><c:spPr>${solidFill(color)}<a:ln w="19050">${solidFill('FFFFFF')}</a:ln></c:spPr></c:dPt>`
    if (kind === 'line' || kind === 'scatter' || kind === 'radar') return `<c:dPt><c:idx val="${index}"/><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${solidFill(color)}</c:spPr></c:marker><c:bubble3D val="0"/></c:dPt>`
    return `<c:dPt><c:idx val="${index}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr>${solidFill(color)}<a:ln><a:noFill/></a:ln></c:spPr></c:dPt>`
  }).join('')
}

function seriesXml(series, index, kind, context, stacked) {
  const { read, defaultSheet, palette, chart } = context
  const color = hex6(series.color, paletteColor(index, palette))
  const name = seriesNameXml(series, read, defaultSheet, `Series${index + 1}`)
  const head = `<c:idx val="${index}"/><c:order val="${index}"/>${name}`
  const labels = dataLabelsXml(series.dataLabels, kind, stacked)
  const vary = (kind === 'pie' || kind === 'doughnut') || (chart.varyColors === true && chart.series.length === 1)
  const valueData = series.valuesRef ? read(series.valuesRef, defaultSheet) : null
  const count = valueData ? valueData.values.length : (series.valuesCache || []).length
  const points = pointColorXml(series, kind, count, palette, vary)
  const categories = kind === 'scatter'
    ? dataXml('xVal', series.xValuesRef || series.categoriesRef, read, defaultSheet, series.xValuesCache || series.categoriesCache, Boolean(series.xValuesRef ? true : series.xValuesCache), undefined)
    : (series.categoriesRef || (series.categoriesCache && series.categoriesCache.length) ? dataXml('cat', series.categoriesRef, read, defaultSheet, series.categoriesCache, false) : '')
  const values = dataXml(kind === 'scatter' ? 'yVal' : 'val', series.valuesRef, read, defaultSheet, series.valuesCache, true, series.valuesNumFmt)
  switch (kind) {
    case 'column':
    case 'bar':
      return `<c:ser>${head}<c:spPr>${solidFill(color)}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="${series.invertIfNegative ? 1 : 0}"/>${points}${labels}${categories}${values}</c:ser>`
    case 'line':
      return `<c:ser>${head}${lineSpPr(series, color)}${markerXml(series, color)}${points}${labels}${categories}${values}<c:smooth val="${series.smooth ? 1 : 0}"/></c:ser>`
    case 'area':
      return `<c:ser>${head}<c:spPr>${solidFill(color)}<a:ln><a:noFill/></a:ln></c:spPr>${points}${labels}${categories}${values}</c:ser>`
    case 'pie':
    case 'doughnut':
      return `<c:ser>${head}<c:spPr>${solidFill(color)}<a:ln w="19050">${solidFill('FFFFFF')}</a:ln></c:spPr>${points}${labels}${categories}${values}</c:ser>`
    case 'scatter':
      return `<c:ser>${head}${lineSpPr(series, color)}${markerXml({ ...series, marker: series.marker || 'circle' }, color)}${points}${labels}${categories}${values}<c:smooth val="${series.smooth ? 1 : 0}"/></c:ser>`
    case 'radar':
      return `<c:ser>${head}${lineSpPr(series, color)}${markerXml(series, color)}${points}${labels}${categories}${values}</c:ser>`
    default:
      return ''
  }
}

function gridlinesXml() {
  return `<c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${solidFill('D9D9D9')}<a:round/></a:ln></c:spPr></c:majorGridlines>`
}

function axisTitleXml(axis, vertical) {
  if (!axis || !axis.title) return ''
  return richTitle(axis.title, 10, '595959', vertical ? { rot: -5400000 } : {})
}

function catAxXml(id, crossId, axis, options) {
  axis = axis || {}
  const visible = axis.visible !== false && !options.deleted
  const rotation = Number.isFinite(axis.labelRotation) && axis.labelRotation ? Math.round(axis.labelRotation * 60000) : 0
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="${axis.reverse ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="${visible ? 0 : 1}"/><c:axPos val="${options.position}"/>${axis.gridlines ? gridlinesXml() : ''}${axisTitleXml(axis, options.position === 'l' || options.position === 'r')}<c:numFmt formatCode="${escapeXml(axis.numFmt || 'General')}" sourceLinked="${axis.numFmt ? 0 : 1}"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${solidFill('D9D9D9')}<a:round/></a:ln></c:spPr>${textProperties(9, '595959', { rot: rotation || undefined })}<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
}

function valAxXml(id, crossId, axis, options) {
  axis = axis || {}
  const visible = axis.visible !== false && !options.deleted
  const scaling = `<c:scaling>${axis.logBase >= 2 ? `<c:logBase val="${axis.logBase}"/>` : ''}<c:orientation val="${axis.reverse ? 'maxMin' : 'minMax'}"/>${Number.isFinite(axis.max) ? `<c:max val="${axis.max}"/>` : ''}${Number.isFinite(axis.min) ? `<c:min val="${axis.min}"/>` : ''}</c:scaling>`
  const numFmt = options.percent ? '<c:numFmt formatCode="0%" sourceLinked="1"/>' : `<c:numFmt formatCode="${escapeXml(axis.numFmt || 'General')}" sourceLinked="${axis.numFmt ? 0 : 1}"/>`
  const line = options.line ? `<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${solidFill('D9D9D9')}<a:round/></a:ln></c:spPr>` : NO_FILL_SP
  const gridlines = axis.gridlines === undefined ? options.defaultGridlines : axis.gridlines
  return `<c:valAx><c:axId val="${id}"/>${scaling}<c:delete val="${visible ? 0 : 1}"/><c:axPos val="${options.position}"/>${gridlines ? gridlinesXml() : ''}${axisTitleXml(axis, options.position === 'l' || options.position === 'r')}${numFmt}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${line}${textProperties(9, '595959')}<c:crossAx val="${crossId}"/><c:crosses val="${options.crosses || 'autoZero'}"/><c:crossBetween val="${options.crossBetween || 'between'}"/>${axis.majorUnit > 0 ? `<c:majorUnit val="${axis.majorUnit}"/>` : ''}</c:valAx>`
}

function groupingValue(kind, grouping) {
  if (grouping === 'stacked' || grouping === 'percentStacked') return grouping
  return kind === 'column' || kind === 'bar' ? 'clustered' : 'standard'
}

const AX = { cat: 500000001, val: 500000002, cat2: 500000003, val2: 500000004 }

/** Generate a complete c:chartSpace part for a model chart. */
function renderChartXml(chart, context) {
  const series = Array.isArray(chart.series) ? chart.series : []
  const type = chart.type
  const grouping = chart.grouping || 'clustered'
  const stacked = grouping !== 'clustered'
  const indexed = series.map((item, index) => ({ item, index }))
  const style = chart.style || {}
  const textColor = hex6(style.textColor, '595959')
  const groups = []
  const kindOfSeries = (item) => {
    if (type !== 'combo') return type
    const kind = item.type || 'column'
    return kind === 'bar' ? 'column' : kind === 'scatter' ? 'line' : kind
  }
  if (type === 'combo') {
    const order = ['area', 'column', 'line']
    for (const secondary of [false, true]) {
      for (const kind of order) {
        const members = indexed.filter(({ item }) => kindOfSeries(item) === kind && Boolean(item.secondaryAxis) === secondary)
        if (members.length) groups.push({ kind, secondary, members })
      }
    }
  } else groups.push({ kind: type, secondary: false, members: type === 'pie' ? indexed.slice(0, 1) : indexed })
  const hasSecondary = groups.some((group) => group.secondary)
  if (hasSecondary && !groups.some((group) => !group.secondary)) groups.forEach((group) => { group.secondary = false })
  const secondaryUsed = groups.some((group) => group.secondary)
  const horizontal = type === 'bar'
  const groupXml = groups.map((group) => {
    const kind = group.kind
    const members = group.members.map(({ item, index }) => seriesXml(item, index, kind, context, stacked && (type !== 'combo' || kind === 'column'))).join('')
    const axIds = group.secondary ? `<c:axId val="${AX.cat2}"/><c:axId val="${AX.val2}"/>` : `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/>`
    const vary = chart.varyColors === true && series.length === 1 ? 1 : 0
    switch (kind) {
      case 'column':
      case 'bar': {
        const barGrouping = type === 'combo' ? groupingValue('column', grouping) : groupingValue(kind, grouping)
        const isStacked = barGrouping !== 'clustered'
        const gap = Math.max(0, Math.min(500, Math.round(Number.isFinite(chart.gapWidth) ? chart.gapWidth : isStacked ? 150 : 219)))
        const overlap = isStacked ? 100 : Math.max(-100, Math.min(100, Math.round(Number.isFinite(chart.overlap) ? chart.overlap : -27)))
        return `<c:barChart><c:barDir val="${kind === 'bar' || horizontal ? 'bar' : 'col'}"/><c:grouping val="${barGrouping}"/><c:varyColors val="${vary}"/>${members}<c:gapWidth val="${gap}"/><c:overlap val="${overlap}"/>${axIds}</c:barChart>`
      }
      case 'line': {
        const lineGrouping = type === 'combo' ? 'standard' : groupingValue('line', grouping)
        return `<c:lineChart><c:grouping val="${lineGrouping}"/><c:varyColors val="0"/>${members}<c:marker val="1"/>${axIds}</c:lineChart>`
      }
      case 'area': {
        const areaGrouping = type === 'combo' ? 'standard' : groupingValue('area', grouping)
        return `<c:areaChart><c:grouping val="${areaGrouping}"/><c:varyColors val="0"/>${members}${axIds}</c:areaChart>`
      }
      case 'pie':
        return `<c:pieChart><c:varyColors val="1"/>${members}<c:firstSliceAng val="${Math.max(0, Math.min(360, Math.round(chart.firstSliceAngle || 0)))}"/></c:pieChart>`
      case 'doughnut':
        return `<c:doughnutChart><c:varyColors val="1"/>${members}<c:firstSliceAng val="${Math.max(0, Math.min(360, Math.round(chart.firstSliceAngle || 0)))}"/><c:holeSize val="${Math.max(10, Math.min(90, Math.round(Number.isFinite(chart.holeSize) ? chart.holeSize : 75)))}"/></c:doughnutChart>`
      case 'scatter':
        return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${members}${axIds}</c:scatterChart>`
      case 'radar':
        return `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${members}${axIds}</c:radarChart>`
      default:
        return ''
    }
  }).join('')
  const axes = chart.axes || {}
  let axesXml = ''
  const percent = grouping === 'percentStacked' && type !== 'combo'
  if (type === 'scatter') {
    axesXml = valAxXml(AX.cat, AX.val, axes.x, { position: 'b', crossBetween: 'midCat', defaultGridlines: false, line: true }) +
      valAxXml(AX.val, AX.cat, axes.y, { position: 'l', crossBetween: 'midCat', defaultGridlines: true })
  } else if (type === 'radar') {
    axesXml = catAxXml(AX.cat, AX.val, axes.x, { position: 'b' }) + valAxXml(AX.val, AX.cat, axes.y, { position: 'l', defaultGridlines: true, crossBetween: 'between' })
  } else if (type !== 'pie' && type !== 'doughnut') {
    const allArea = groups.every((group) => group.kind === 'area')
    axesXml = catAxXml(AX.cat, AX.val, axes.x, { position: horizontal ? 'l' : 'b' }) +
      valAxXml(AX.val, AX.cat, axes.y, { position: horizontal ? 'b' : 'l', crossBetween: allArea ? 'midCat' : 'between', defaultGridlines: true, percent })
    if (secondaryUsed) {
      axesXml += catAxXml(AX.cat2, AX.val2, {}, { position: 'b', deleted: true }) +
        valAxXml(AX.val2, AX.cat2, axes.y2, { position: 'r', crosses: 'max', crossBetween: 'between', defaultGridlines: false })
    }
  }
  // Title: explicit text, cell reference, automatic (single series) or none.
  const titleSize = Number.isFinite(style.titleFontSize) && style.titleFontSize > 0 ? style.titleFontSize : 14
  let titleXml = ''
  let autoTitleDeleted = 1
  if (chart.titleRef) {
    const data = context.read(chart.titleRef, context.defaultSheet)
    titleXml = refTitle(absoluteRef(chart.titleRef, context.defaultSheet), data ? data.texts.join(' ') : chart.title, titleSize, textColor)
    autoTitleDeleted = 0
  } else if (typeof chart.title === 'string' && chart.title.trim()) {
    titleXml = richTitle(chart.title, titleSize, textColor)
    autoTitleDeleted = 0
  } else if (chart.title === undefined && !chart.autoTitleDeleted && series.length === 1) {
    titleXml = `<c:title><c:overlay val="0"/>${NO_FILL_SP}${textProperties(titleSize, textColor)}</c:title>`
    autoTitleDeleted = 0
  }
  const legendPosition = { right: 'r', left: 'l', top: 't', bottom: 'b' }[chart.legend || (series.length > 1 || type === 'pie' || type === 'doughnut' ? 'bottom' : 'none')]
  const legendXml = legendPosition ? `<c:legend><c:legendPos val="${legendPosition}"/><c:overlay val="0"/>${NO_FILL_SP}${textProperties(9, textColor)}</c:legend>` : ''
  const background = style.background === 'transparent' ? '<a:noFill/>' : solidFill(hex6(style.background, 'FFFFFF'))
  const border = style.border === null ? '<a:ln><a:noFill/></a:ln>' : `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${solidFill(hex6(style.border, 'D9D9D9'))}<a:round/></a:ln>`
  const blanks = ['gap', 'zero', 'span'].includes(chart.displayBlanksAs) ? chart.displayBlanksAs : 'gap'
  const fontSize = Number.isFinite(style.fontSize) && style.fontSize > 0 ? style.fontSize : 9
  return `${XML_DECLARATION}<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="${style.roundedCorners ? 1 : 0}"/>` +
    `<c:chart>${titleXml}<c:autoTitleDeleted val="${autoTitleDeleted}"/><c:plotArea><c:layout/>${groupXml}${axesXml}${NO_FILL_SP}</c:plotArea>${legendXml}` +
    `<c:plotVisOnly val="${chart.plotVisibleOnly === false ? 0 : 1}"/><c:dispBlanksAs val="${blanks}"/></c:chart>` +
    `<c:spPr>${background}${border}</c:spPr>${textProperties(fontSize, textColor, { font: style.fontFamily })}` +
    `<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>` +
    '</c:chartSpace>'
}

// ---------------------------------------------------------------------------
// Export: write charts into the package ExcelJS produced
// ---------------------------------------------------------------------------

function anchorPointXml(tag, point) {
  const clampInt = (value, max) => Math.max(0, Math.min(max, Math.trunc(Number(value) || 0)))
  return `<xdr:${tag}><xdr:col>${clampInt(point && point.col, MAX_COLS - 1)}</xdr:col><xdr:colOff>${clampInt(point && point.colOffsetEmu, 2_000_000_000)}</xdr:colOff><xdr:row>${clampInt(point && point.row, MAX_ROWS - 1)}</xdr:row><xdr:rowOff>${clampInt(point && point.rowOffsetEmu, 2_000_000_000)}</xdr:rowOff></xdr:${tag}>`
}

function wrapAnchor(anchor, contentXml) {
  const editAs = anchor && (anchor.editAs === 'oneCell' || anchor.editAs === 'absolute') ? ` editAs="${anchor.editAs}"` : ''
  const from = (anchor && anchor.from) || { row: 0, col: 0 }
  let to = (anchor && anchor.to) || { row: from.row + 15, col: from.col + 8 }
  if (to.row < from.row || to.col < from.col) to = { ...to, row: Math.max(to.row, from.row), col: Math.max(to.col, from.col) }
  return `<xdr:twoCellAnchor${editAs}>${anchorPointXml('from', from)}${anchorPointXml('to', to)}${contentXml}<xdr:clientData/></xdr:twoCellAnchor>`
}

function graphicFrameXml(objectId, name, description, rId) {
  return `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${objectId}" name="${escapeXml(name)}"${description ? ` descr="${escapeXml(description)}"` : ''}/><xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rId}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
}

class PartAllocator {
  constructor(zip) {
    this.used = new Set(Object.keys(zip.files).map((name) => name.toLowerCase()))
  }
  take(name) { this.used.add(name.toLowerCase()); return name }
  isFree(name) { return !this.used.has(name.toLowerCase()) }
  /** Next free name following `dir/stem{n}.ext`. */
  next(directory, stem, extension) {
    for (let index = 1; index < 1_000_000; index += 1) {
      const candidate = `${directory}/${stem}${index}${extension}`
      if (this.isFree(candidate)) return this.take(candidate)
    }
    throw new Error('Could not allocate a package part name.')
  }
  like(sourcePart) {
    const directory = path.posix.dirname(sourcePart)
    const extension = path.posix.extname(sourcePart)
    const stem = path.posix.basename(sourcePart, extension).replace(/\d+$/, '') || 'part'
    return this.next(directory, stem, extension)
  }
}

class ContentTypes {
  constructor(xml) {
    this.xml = xml || `${XML_DECLARATION}<Types xmlns="${NS.ct}"></Types>`
    this.overrides = new Set()
    this.defaults = new Set()
    for (const match of this.xml.matchAll(/<Override\b[^>]*PartName="([^"]+)"/g)) this.overrides.add(match[1].toLowerCase())
    for (const match of this.xml.matchAll(/<Default\b[^>]*Extension="([^"]+)"/g)) this.defaults.add(match[1].toLowerCase())
    this.additions = []
    this.defaultAdditions = []
  }
  override(partName, contentType) {
    const key = `/${partName}`.toLowerCase()
    if (!contentType || this.overrides.has(key)) return
    this.overrides.add(key)
    this.additions.push(`<Override PartName="/${escapeXml(partName)}" ContentType="${escapeXml(contentType)}"/>`)
  }
  ensureDefault(extension, contentType) {
    const key = String(extension || '').toLowerCase()
    if (!key || !contentType || this.defaults.has(key)) return
    this.defaults.add(key)
    this.defaultAdditions.push(`<Default Extension="${escapeXml(key)}" ContentType="${escapeXml(contentType)}"/>`)
  }
  render() {
    if (!this.additions.length && !this.defaultAdditions.length) return this.xml
    let xml = this.xml
    if (this.defaultAdditions.length) {
      const firstOverride = xml.search(/<Override\b/)
      xml = firstOverride >= 0 ? `${xml.slice(0, firstOverride)}${this.defaultAdditions.join('')}${xml.slice(firstOverride)}` : xml.replace(/<\/Types>\s*$/, `${this.defaultAdditions.join('')}</Types>`)
    }
    return xml.replace(/<\/Types>\s*$/, `${this.additions.join('')}</Types>`)
  }
}

/** Copy a part and everything it references (recursively) from the base package. Returns the new part name. */
async function copyPartGraph(baseZip, baseTypes, zip, allocator, types, sourcePart, copied, depth = 0) {
  if (copied.has(sourcePart)) return copied.get(sourcePart)
  if (depth > 6) throw new Error('Chart part references are nested too deeply.')
  const entry = baseZip.file(sourcePart)
  if (!entry) throw new Error(`Missing source part ${sourcePart}`)
  const targetPart = allocator.like(sourcePart)
  copied.set(sourcePart, targetPart)
  zip.file(targetPart, await entry.async('nodebuffer'))
  const contentType = contentTypeFor(baseTypes, sourcePart)
  const extension = path.posix.extname(sourcePart).slice(1).toLowerCase()
  if (baseTypes.overrides.has(sourcePart.toLowerCase()) || ['xml'].includes(extension)) types.override(targetPart, contentType)
  else types.ensureDefault(extension, contentType)
  const relationships = await readRelationships(baseZip, sourcePart)
  if (relationships.length) {
    const rewritten = []
    for (const rel of relationships) {
      if (rel.external || !rel.resolved) { rewritten.push(rel); continue }
      if (!baseZip.file(rel.resolved)) continue
      const copiedTarget = await copyPartGraph(baseZip, baseTypes, zip, allocator, types, rel.resolved, copied, depth + 1)
      rewritten.push({ ...rel, target: relativeTarget(targetPart, copiedTarget) })
    }
    zip.file(relsPathFor(targetPart), renderRelationships(rewritten))
  }
  return targetPart
}

function rootStartTag(xml) {
  const doc = /<(?!\?|!)[^>]+>/.exec(xml)
  return doc ? { tag: doc[0], index: doc.index } : null
}

/** Add xmlns declarations (and mc:Ignorable) from `sourceRootTag` missing on the XML root. */
function mergeRootNamespaces(xml, sourceRootTag) {
  const start = rootStartTag(xml)
  if (!start || !sourceRootTag) return xml
  let tag = start.tag
  for (const match of sourceRootTag.matchAll(/\s(xmlns:[A-Za-z0-9_.-]+)="([^"]*)"/g)) {
    if (!new RegExp(`\\s${match[1].replace('.', '\\.')}=`).test(tag)) tag = tag.replace(/\s*(\/?)>$/, ` ${match[1]}="${match[2]}"$1>`)
  }
  const ignorable = /\smc:Ignorable="([^"]*)"/.exec(sourceRootTag)
  if (ignorable && !/\smc:Ignorable=/.test(tag)) tag = tag.replace(/\s*(\/?)>$/, ` mc:Ignorable="${ignorable[1]}"$1>`)
  return xml.slice(0, start.index) + tag + xml.slice(start.index + start.tag.length)
}

function ensureRootNamespace(xml, prefix, uri) {
  const start = rootStartTag(xml)
  if (!start || new RegExp(`\\sxmlns:${prefix}=`).test(start.tag)) return xml
  const tag = start.tag.replace(/\s*(\/?)>$/, ` xmlns:${prefix}="${uri}"$1>`)
  return xml.slice(0, start.index) + tag + xml.slice(start.index + start.tag.length)
}

const AFTER_DRAWING = new Set(['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'])

/** Insert <drawing r:id> as a direct worksheet child at its schema position. */
function insertDrawingElement(sheetXml, rId) {
  const prefixMatch = /<([A-Za-z0-9_]+:)?worksheet\b/.exec(sheetXml)
  const prefix = (prefixMatch && prefixMatch[1]) || ''
  const element = `<${prefix}drawing r:id="${rId}"/>`
  let scanFrom = sheetXml.lastIndexOf(`</${prefix}sheetData>`)
  if (scanFrom < 0) { const empty = sheetXml.search(new RegExp(`<${prefix}sheetData\\s*/>`)); scanFrom = empty >= 0 ? empty : 0 }
  const tagRe = /<(\/?)([A-Za-z0-9_]+:)?([A-Za-z0-9_]+)\b[^>]*?(\/?)>/g
  tagRe.lastIndex = scanFrom
  let depth = 0
  let match
  while ((match = tagRe.exec(sheetXml))) {
    const [, closing, , local, selfClosing] = match
    if (closing) {
      if (depth === 0) {
        if (local === 'worksheet') return sheetXml.slice(0, match.index) + element + sheetXml.slice(match.index)
        continue // the closing sheetData tag
      }
      depth -= 1
      continue
    }
    if (depth === 0 && AFTER_DRAWING.has(local)) return sheetXml.slice(0, match.index) + element + sheetXml.slice(match.index)
    if (!selfClosing) depth += 1
  }
  return sheetXml.replace(new RegExp(`</${prefix}worksheet>\\s*$`), `${element}</${prefix}worksheet>`)
}

function nextRelationshipId(relationships, stem) {
  const used = new Set(relationships.map((rel) => rel.id))
  for (let index = 1; ; index += 1) if (!used.has(`${stem}${index}`)) return `${stem}${index}`
}

function isChartExportable(chart) {
  return chart && typeof chart === 'object' && chart.anchor && chart.anchor.from && chart.anchor.to && Array.isArray(chart.series)
}

/**
 * Write every model chart into the output package (JSZip instance, mutated in place).
 * `baseBuffer` is the original package used for byte-exact copies of unmodified charts.
 */
async function writeChartsToPackage(zip, model, baseBuffer, warnings) {
  const sheetsWithCharts = (model && Array.isArray(model.sheets) ? model.sheets : []).filter((sheet) => Array.isArray(sheet.charts) && sheet.charts.some(isChartExportable))
  if (!sheetsWithCharts.length) return 0
  const JSZip = require('jszip')
  const allocator = new PartAllocator(zip)
  const types = new ContentTypes(await readPart(zip, '[Content_Types].xml'))
  const outputSheets = await workbookSheetParts(zip)
  const read = modelReader(model)
  const themeColors = model.metadata && model.metadata.themeColors
  const theme = Array.isArray(themeColors) && themeColors.length >= 10 ? themeColors.map((value, index) => (/^[0-9a-f]{6}$/i.test(String(value || '')) ? String(value).toUpperCase() : OFFICE_THEME[index])) : OFFICE_THEME
  // The base part must still be the chart that was imported: a stale part name (for example
  // after the package was rewritten) would otherwise copy a different chart over this one.
  const basePartMatches = async (chart, info) => {
    try {
      const xml = await readPart(baseZip, chart.sourcePart)
      if (!xml) return false
      const parsed = info.kind === 'chartex' ? parseChartExXml(xml) : parseChartXml(xml, theme)
      return chartFingerprint(parsed) === info.fingerprint
    } catch {
      return false
    }
  }
  let baseZip = null, baseTypes = null
  const baseDrawings = new Map()
  const loadBase = async () => {
    if (baseZip === null) {
      try { baseZip = baseBuffer ? await JSZip.loadAsync(baseBuffer) : false } catch { baseZip = false }
      if (baseZip) baseTypes = await contentTypesOf(baseZip)
    }
    return baseZip || null
  }
  const baseDrawing = async (part) => {
    if (baseDrawings.has(part)) return baseDrawings.get(part)
    let value = null
    try {
      const xml = await readPart(baseZip, part)
      if (xml) {
        const doc = documentElement(parseXml(xml))
        value = { xml, doc, objects: drawingObjects(doc), rels: await readRelationships(baseZip, part), rootTag: (rootStartTag(xml) || {}).tag }
      }
    } catch { value = null }
    baseDrawings.set(part, value)
    return value
  }
  let written = 0
  for (const sheet of sheetsWithCharts) {
    const target = outputSheets.find((item) => item.name === sheet.name) || outputSheets[model.sheets.indexOf(sheet)]
    if (!target) continue
    let sheetXml = await readPart(zip, target.part)
    if (!sheetXml) continue
    const sheetRels = await readRelationships(zip, target.part)
    let drawingRel = sheetRels.find((rel) => rel.type === REL.drawing && rel.resolved && zip.file(rel.resolved))
    let drawingPart, drawingXml, drawingRels
    if (drawingRel) {
      drawingPart = drawingRel.resolved
      drawingXml = await readPart(zip, drawingPart)
      drawingRels = await readRelationships(zip, drawingPart)
    } else {
      drawingPart = allocator.next('xl/drawings', 'drawing', '.xml')
      drawingXml = `${XML_DECLARATION}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`
      drawingRels = []
    }
    let objectId = Math.max(1, ...[...drawingXml.matchAll(/<(?:xdr:)?cNvPr\b[^>]*\bid="(\d+)"/g)].map((match) => Number(match[1]) || 0)) + 1
    const anchors = []
    const sourceRootTags = []
    const sheetNames = new Set(model.sheets.map((item) => String(item.name).toLowerCase()))
    for (const chart of sheet.charts.filter(isChartExportable)) {
      let anchorXml = null
      // 1) Byte-for-byte copy of unmodified imported charts.
      const info = chart.sourceInfo
      const unchanged = chart.sourcePart && info && !chart.modified && (chart.type === 'unsupported' || chartFingerprint(chart) === info.fingerprint)
      if (chart.sourcePart && info && (unchanged || chart.type === 'unsupported') && await loadBase()) {
        try {
          const drawing = await baseDrawing(info.drawingPart)
          const relFor = (object) => { const reference = object && chartReferenceOf(object.anchorNode); const rel = reference && drawing.rels.find((item) => item.id === reference.rId); return { reference, rel } }
          let object = drawing && drawing.objects.find((item) => item.index === info.anchorIndex)
          let { reference, rel } = relFor(object)
          if (!rel || rel.resolved !== chart.sourcePart) {
            object = drawing && drawing.objects.find((item) => relFor(item).rel?.resolved === chart.sourcePart)
            ;({ reference, rel } = relFor(object))
          }
          if (object && reference && rel && baseZip.file(chart.sourcePart) && await basePartMatches(chart, info)) {
            const copied = new Map()
            const chartPart = await copyPartGraph(baseZip, baseTypes, zip, allocator, types, chart.sourcePart, copied)
            const rId = nextRelationshipId(drawingRels, 'rIdChart')
            drawingRels.push({ id: rId, type: rel.type, target: relativeTarget(drawingPart, chartPart) })
            let fragment = drawing.xml.slice(reference.content.start, reference.content.end)
            // Other relationships used by the fragment (e.g. fallback pictures) are copied too.
            const replacements = new Map([[reference.rId, rId]])
            for (const match of fragment.matchAll(/\sr:(?:id|embed|link|pict)="([^"]+)"/g)) {
              const otherId = match[1]
              if (replacements.has(otherId)) continue
              const other = drawing.rels.find((item) => item.id === otherId)
              if (!other) continue
              const newId = nextRelationshipId(drawingRels, 'rIdChartObj')
              if (other.external || !other.resolved) drawingRels.push({ id: newId, type: other.type, target: other.target, external: other.external })
              else if (baseZip.file(other.resolved)) drawingRels.push({ id: newId, type: other.type, target: relativeTarget(drawingPart, await copyPartGraph(baseZip, baseTypes, zip, allocator, types, other.resolved, copied)) })
              else continue
              replacements.set(otherId, newId)
            }
            fragment = fragment.replace(/(\sr:(?:id|embed|link|pict)=")([^"]+)(")/g, (all, open, value, close) => (replacements.has(value) ? `${open}${replacements.get(value)}${close}` : all))
            fragment = fragment.replace(/(<(?:xdr:)?cNvPr\b[^>]*?\bid=")(\d+)(")/g, (_all, open, _value, close) => `${open}${objectId++}${close}`)
            anchorXml = wrapAnchor(chart.anchor, fragment)
            if (drawing.rootTag) sourceRootTags.push(drawing.rootTag)
          }
        } catch {
          anchorXml = null
        }
      }
      // 2) Generated DrawingML for new/edited charts.
      if (!anchorXml && chart.type !== 'unsupported' && chart.series.length) {
        const defaultSheet = sheet.name
        const palette = Array.isArray(chart.style && chart.style.palette) && chart.style.palette.length
          ? chart.style.palette
          : Array.isArray(model.metadata && model.metadata.themeColors) && model.metadata.themeColors.length >= 10 ? model.metadata.themeColors.slice(4, 10) : OFFICE_ACCENTS
        const chartXml = renderChartXml(chart, { read, defaultSheet, palette, chart, sheetNames })
        const chartPart = allocator.next('xl/charts', 'chart', '.xml')
        zip.file(chartPart, chartXml)
        types.override(chartPart, CT.chart)
        const rId = nextRelationshipId(drawingRels, 'rIdChart')
        drawingRels.push({ id: rId, type: REL.chart, target: relativeTarget(drawingPart, chartPart) })
        const id = objectId++
        anchorXml = wrapAnchor(chart.anchor, graphicFrameXml(id, chart.name || `Chart ${id - 1}`, chart.description, rId))
      }
      if (!anchorXml && chart.type === 'unsupported' && Array.isArray(warnings)) {
        const label = chart.title || chart.name || chart.unsupportedKind || 'chart'
        warnings.push(`The ${String(chart.unsupportedKind || 'special')} chart "${label}" on sheet "${sheet.name}" could not be kept because its original chart data is unavailable.`)
      }
      if (anchorXml) { anchors.push(anchorXml); written += 1 }
    }
    if (!anchors.length) continue
    for (const tag of sourceRootTags) drawingXml = mergeRootNamespaces(drawingXml, tag)
    drawingXml = ensureRootNamespace(drawingXml, 'a', NS.a)
    const close = drawingXml.lastIndexOf('</xdr:wsDr>')
    if (close < 0) continue
    drawingXml = drawingXml.slice(0, close) + anchors.join('') + drawingXml.slice(close)
    zip.file(drawingPart, drawingXml)
    zip.file(relsPathFor(drawingPart), renderRelationships(drawingRels))
    types.override(drawingPart, CT.drawing)
    if (!drawingRel) {
      const rId = nextRelationshipId(sheetRels, 'rIdDr')
      sheetRels.push({ id: rId, type: REL.drawing, target: relativeTarget(target.part, drawingPart) })
      zip.file(relsPathFor(target.part), renderRelationships(sheetRels))
      sheetXml = ensureRootNamespace(sheetXml, 'r', NS.r)
      sheetXml = insertDrawingElement(sheetXml, rId)
      zip.file(target.part, sheetXml)
    }
  }
  zip.file('[Content_Types].xml', types.render())
  return written
}

module.exports = {
  importChartsIntoSheets,
  writeChartsToPackage,
  parseChartXml,
  renderChartXml,
  chartFingerprint,
  parseXml,
  workbookSheetParts,
  readRelationships,
  relsPathFor,
  renderRelationships,
  // exposed for QA
  _internal: { insertDrawingElement, absoluteRef, parseRefAreas, drawingObjects, documentElement },
}
