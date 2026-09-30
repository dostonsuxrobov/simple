'use strict'

// Excel sparkline groups (x14:sparklineGroups in the worksheet extLst). ExcelJS ignores them, so
// they are read from and written back into the package here. Unedited groups are copied from
// the source XML byte-for-byte.

const { parseXml } = require('./chart-xlsx.cjs')

const EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']
const COLOR_KEYS = [['colorSeries', 'series'], ['colorNegative', 'negative'], ['colorAxis', 'axis'], ['colorMarkers', 'markers'], ['colorFirst', 'first'], ['colorLast', 'last'], ['colorHigh', 'high'], ['colorLow', 'low']]
const FLAGS = ['markers', 'high', 'low', 'first', 'last', 'negative', 'displayXAxis', 'displayHidden', 'rightToLeft', 'dateAxis']

const localName = (name) => String(name || '').replace(/^.*:/, '')
const childrenNamed = (node, name) => (node ? node.children.filter((item) => localName(item.name) === name) : [])
const textOf = (node) => (node ? (node.text || '') + node.children.map(textOf).join('') : '')

function escapeXml(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function tint(hex, amount) {
  if (!amount) return hex
  const channels = [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16) / 255)
  const max = Math.max(...channels), min = Math.min(...channels)
  let h = 0, s = 0
  let l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    const [r, g, b] = channels
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
    h /= 6
  }
  l = amount < 0 ? l * (1 + amount) : l * (1 - amount) + amount
  const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const rgb = s === 0 ? [l, l, l] : [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)]
  return rgb.map((value) => Math.round(Math.max(0, Math.min(1, value)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase()
}

function colorOf(node, theme) {
  if (!node) return undefined
  const attrs = node.attrs || {}
  let hex = null
  if (attrs.rgb && /^[0-9a-f]{6,8}$/i.test(attrs.rgb)) hex = attrs.rgb.slice(-6).toUpperCase()
  else if (attrs.theme !== undefined && Number.isInteger(Number(attrs.theme))) hex = (theme[Number(attrs.theme)] || OFFICE_THEME[Number(attrs.theme)] || '4472C4').toUpperCase()
  if (!hex) return undefined
  return `#${tint(hex, Number(attrs.tint) || 0)}`
}

function flag(value) {
  return value === '1' || value === 'true'
}

function groupSignature(group) {
  const copy = { ...group }
  delete copy.sourceXml
  delete copy.signature
  return JSON.stringify(copy)
}

function parseGroup(node, theme) {
  const attrs = node.attrs || {}
  const group = {
    type: attrs.type === 'column' ? 'column' : attrs.type === 'stacked' ? 'stacked' : 'line',
    colors: {},
    sparklines: [],
  }
  for (const key of FLAGS) if (flag(attrs[key])) group[key] = true
  if (attrs.displayEmptyCellsAs) group.displayEmptyCellsAs = attrs.displayEmptyCellsAs
  for (const key of ['minAxisType', 'maxAxisType']) if (attrs[key]) group[key] = attrs[key]
  for (const key of ['manualMin', 'manualMax', 'lineWeight']) if (attrs[key] !== undefined && Number.isFinite(Number(attrs[key]))) group[key] = Number(attrs[key])
  for (const [element, key] of COLOR_KEYS) {
    const color = colorOf(childrenNamed(node, element)[0], theme)
    if (color) group.colors[key] = color
  }
  for (const list of childrenNamed(node, 'sparklines')) {
    for (const item of childrenNamed(list, 'sparkline')) {
      const source = textOf(childrenNamed(item, 'f')[0]).trim()
      const cell = textOf(childrenNamed(item, 'sqref')[0]).trim().toUpperCase()
      if (cell) group.sparklines.push({ source, cell })
    }
  }
  return group
}

async function sheetParts(zip) {
  const workbook = zip.file('xl/workbook.xml') ? await zip.file('xl/workbook.xml').async('string') : ''
  const rels = zip.file('xl/_rels/workbook.xml.rels') ? await zip.file('xl/_rels/workbook.xml.rels').async('string') : ''
  const targets = new Map()
  for (const match of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = /\bId="([^"]+)"/.exec(match[1])
    const target = /\bTarget="([^"]+)"/.exec(match[1])
    if (id && target) targets.set(id[1], target[1].replace(/^\/?xl\//, '').replace(/^\//, ''))
  }
  const parts = []
  for (const match of workbook.matchAll(/<(?:[A-Za-z0-9]+:)?sheet\b([^>]*)\/?>/g)) {
    const name = /\bname="([^"]*)"/.exec(match[1])
    const id = /\br:id="([^"]+)"/.exec(match[1]) || /\b[A-Za-z0-9]+:id="([^"]+)"/.exec(match[1])
    const target = id && targets.get(id[1])
    if (name && target) parts.push({ name: name[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'"), path: `xl/${target}` })
  }
  return parts
}

/** Read every sheet's sparkline groups into the model (sheet.sparklineGroups). */
async function importSparklineGroups(zipOrBuffer, sheets, themeColors) {
  const JSZip = require('jszip')
  const zip = zipOrBuffer && typeof zipOrBuffer.file === 'function' ? zipOrBuffer : await JSZip.loadAsync(zipOrBuffer)
  const theme = Array.isArray(themeColors) && themeColors.length >= 10 ? themeColors.map((value, index) => (/^[0-9a-f]{6}$/i.test(String(value || '')) ? String(value).toUpperCase() : OFFICE_THEME[index])) : OFFICE_THEME
  for (const part of await sheetParts(zip)) {
    const file = zip.file(part.path)
    if (!file) continue
    const xml = await file.async('string')
    if (!xml.includes('sparklineGroup')) continue
    const sheet = sheets.find((item) => item.sourceSheetName === part.name || item.name === part.name)
    if (!sheet) continue
    const block = /<(x14:)?sparklineGroups\b[\s\S]*?<\/(?:x14:)?sparklineGroups>/.exec(xml)
    if (!block) continue
    let doc
    try { doc = parseXml(block[0]) } catch { continue }
    const root = doc.children.find((item) => localName(item.name) === 'sparklineGroups')
    const groups = []
    for (const node of childrenNamed(root, 'sparklineGroup')) {
      const group = parseGroup(node, theme)
      if (!group.sparklines.length) continue
      const raw = block[0].slice(node.start, node.end)
      // Byte-exact reuse needs the conventional x14/xm prefixes the writer declares.
      if (/^<x14:sparklineGroup\b/.test(raw)) group.sourceXml = raw
      group.signature = groupSignature(group)
      groups.push(group)
    }
    if (groups.length) sheet.sparklineGroups = groups
  }
}

function renderColor(name, value) {
  if (!value || !/^#[0-9a-f]{6}$/i.test(value)) return ''
  return `<x14:${name} rgb="FF${value.slice(1).toUpperCase()}"/>`
}

function renderGroup(group) {
  if (group.sourceXml && group.signature === groupSignature(group)) return group.sourceXml
  const attributes = []
  if (group.manualMax !== undefined) attributes.push(`manualMax="${Number(group.manualMax)}"`)
  if (group.manualMin !== undefined) attributes.push(`manualMin="${Number(group.manualMin)}"`)
  if (group.lineWeight !== undefined) attributes.push(`lineWeight="${Number(group.lineWeight)}"`)
  if (group.type === 'column' || group.type === 'stacked') attributes.push(`type="${group.type}"`)
  if (group.dateAxis) attributes.push('dateAxis="1"')
  attributes.push(`displayEmptyCellsAs="${['gap', 'zero', 'span'].includes(group.displayEmptyCellsAs) ? group.displayEmptyCellsAs : 'gap'}"`)
  for (const key of ['markers', 'high', 'low', 'first', 'last', 'negative', 'displayXAxis', 'displayHidden']) if (group[key]) attributes.push(`${key}="1"`)
  if (group.minAxisType && group.minAxisType !== 'individual') attributes.push(`minAxisType="${escapeXml(group.minAxisType)}"`)
  if (group.maxAxisType && group.maxAxisType !== 'individual') attributes.push(`maxAxisType="${escapeXml(group.maxAxisType)}"`)
  if (group.rightToLeft) attributes.push('rightToLeft="1"')
  const colors = group.colors || {}
  const colorXml = COLOR_KEYS.map(([element, key]) => renderColor(element, colors[key] || (key === 'series' ? '#376092' : key === 'negative' ? '#D00000' : key === 'axis' ? '#000000' : key === 'markers' ? '#D00000' : key === 'first' || key === 'last' ? '#D00000' : key === 'high' || key === 'low' ? '#D00000' : undefined))).join('')
  const sparklines = (group.sparklines || []).filter((item) => item && item.cell)
    .map((item) => `<x14:sparkline><xm:f>${escapeXml(item.source || '')}</xm:f><xm:sqref>${escapeXml(item.cell)}</xm:sqref></x14:sparkline>`).join('')
  return `<x14:sparklineGroup ${attributes.join(' ')}>${colorXml}<x14:sparklines>${sparklines}</x14:sparklines></x14:sparklineGroup>`
}

/** Write the model's sparkline groups into each worksheet part of a package built by ExcelJS. */
async function writeSparklinesToPackage(zip, model) {
  const sheets = (model.sheets || []).filter((sheet) => Array.isArray(sheet.sparklineGroups))
  if (!sheets.length) return
  const parts = await sheetParts(zip)
  for (const sheet of sheets) {
    const part = parts.find((item) => item.name === sheet.name)
    const file = part && zip.file(part.path)
    if (!file) continue
    let xml = await file.async('string')
    xml = xml.replace(new RegExp(`<ext\\b[^>]*uri="${EXT_URI.replace(/[{}]/g, '\\$&')}"[\\s\\S]*?</ext>`, 'g'), '')
    xml = xml.replace(/<extLst>\s*<\/extLst>/, '')
    const groups = sheet.sparklineGroups.filter((group) => group && Array.isArray(group.sparklines) && group.sparklines.length)
    if (groups.length) {
      const ext = `<ext uri="${EXT_URI}" xmlns:x14="${NS_X14}"><x14:sparklineGroups xmlns:xm="${NS_XM}">${groups.map(renderGroup).join('')}</x14:sparklineGroups></ext>`
      const close = xml.lastIndexOf('</worksheet>')
      const lastExtList = xml.lastIndexOf('</extLst>')
      if (lastExtList > 0 && !xml.slice(lastExtList + 9, close).trim()) xml = `${xml.slice(0, lastExtList)}${ext}${xml.slice(lastExtList)}`
      else xml = `${xml.slice(0, close)}<extLst>${ext}</extLst>${xml.slice(close)}`
    }
    zip.file(part.path, xml)
  }
}

module.exports = { importSparklineGroups, writeSparklinesToPackage }
