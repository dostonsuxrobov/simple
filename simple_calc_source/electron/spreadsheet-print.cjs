const MAX_PRINT_CELLS = 400_000
const MAX_PRINT_SHEETS = 100
const MAX_PRINT_COPIES = 999
// Collapsed table borders add their outside stroke to the sum of column/row
// dimensions. Reserve the maximum supported 3px stroke plus rounding room so
// fit-to-page never places that stroke beyond the clipped content rectangle.
const TABLE_EDGE_ALLOWANCE = 4
// SPARKLINE() results are a private-use marker string (src/lib/formula-lib-sparkline.ts).
// The marker is never printed as text; the picture comes from input.visuals.
const SPARKLINE_MARKER = '\uE000sparkline:'
// 'fit-height' fits every row on one page; 'custom' is a fixed percentage; 'fit-pages'
// fits the content to a number of pages wide by tall (0 = automatic), as Excel's Page Setup.
const SCALING_MODES = new Set(['actual', 'fit-width', 'fit-height', 'fit-sheet', 'custom', 'fit-pages'])

const PAPER_SIZES = {
  letter: { css: 'Letter', label: 'Letter', width: 8.5, height: 11 },
  a4: { css: 'A4', label: 'A4', width: 210 / 25.4, height: 297 / 25.4 },
  legal: { css: 'Legal', label: 'Legal', width: 8.5, height: 14 },
}

const MARGIN_PRESETS = {
  normal: { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 },
  narrow: { top: 0.25, right: 0.25, bottom: 0.25, left: 0.25 },
  wide: { top: 0.75, right: 0.75, bottom: 0.75, left: 0.75 },
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function columnName(index) {
  let value = Math.max(0, Math.trunc(index)) + 1
  let result = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    value = Math.floor((value - 1) / 26)
  }
  return result
}

function addressOf(row, col) {
  return `${columnName(col)}${row + 1}`
}

function parseAddress(value) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$/i.exec(String(value || '').trim())
  if (!match) return null
  let col = 0
  for (const character of match[1].toUpperCase()) col = col * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2]) - 1
  col -= 1
  if (row < 0 || row > 1_048_575 || col < 0 || col > 16_383) return null
  return { row, col }
}

function parseRange(value) {
  const unqualified = String(value || '').split('!').pop().replace(/'/g, '')
  const pieces = unqualified.split(':')
  if (pieces.length > 2) return null
  const [startText, endText = startText] = pieces
  const start = parseAddress(startText)
  const end = parseAddress(endText)
  if (!start || !end) return null
  return {
    top: Math.min(start.row, end.row),
    bottom: Math.max(start.row, end.row),
    left: Math.min(start.col, end.col),
    right: Math.max(start.col, end.col),
  }
}

const WHOLE_COLUMNS = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/i
const WHOLE_ROWS = /^\$?([1-9][0-9]{0,6}):\$?([1-9][0-9]{0,6})$/

function columnIndexOf(label) {
  let col = 0
  for (const character of label.toUpperCase()) col = col * 26 + character.charCodeAt(0) - 64
  return col - 1
}

/**
 * One print area. Whole-column ($A:$F) and whole-row ($1:$20) areas print to the sheet's
 * used extent, as Excel does; so does a bounded area the importer expanded from that form
 * (pageSetup.printAreaWhole maps it back to the original reference).
 */
function printAreaRange(piece, whole, extents) {
  const unqualified = String(piece || '').split('!').pop().replace(/'/g, '').trim()
  const range = parseRange(unqualified)
  const original = range && typeof whole[unqualified] === 'string' ? whole[unqualified] : ''
  const columns = WHOLE_COLUMNS.exec(range ? original : unqualified)
  const rows = !columns && WHOLE_ROWS.exec(range ? original : unqualified)
  // The bounded copy only records where the sheet ended when it was opened: rows or columns
  // added since are part of a whole-column / whole-row area, so the area follows the used extent.
  if (columns) {
    const left = columnIndexOf(columns[1]), right = columnIndexOf(columns[2])
    if (left > 16_383 || right > 16_383) return null
    return { top: 0, bottom: Math.max(0, extents.bottom), left: Math.min(left, right), right: Math.max(left, right) }
  }
  if (rows) {
    const first = Number(rows[1]) - 1, last = Number(rows[2]) - 1
    if (first > 1_048_575 || last > 1_048_575) return null
    return { top: Math.min(first, last), bottom: Math.max(first, last), left: 0, right: Math.max(0, extents.right) }
  }
  return range
}

function printAreas(setup, extents = { top: 0, bottom: 0, left: 0, right: 0 }) {
  const value = setup && typeof setup === 'object' ? setup.printArea : setup
  if (!value) return []
  const whole = setup && typeof setup === 'object' && setup.printAreaWhole && typeof setup.printAreaWhole === 'object' ? setup.printAreaWhole : {}
  // Excel separates areas with commas (ExcelJS also accepts &&). A quoted sheet
  // name can itself contain commas or escaped apostrophes.
  const pieces = []
  let current = ''
  let quoted = false
  const text = String(value)
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === "'") {
      if (quoted && text[index + 1] === "'") { current += "''"; index += 1; continue }
      quoted = !quoted
    }
    if (!quoted && (character === ',' || text.slice(index, index + 2) === '&&')) {
      pieces.push(current)
      current = ''
      if (character === '&') index += 1
    } else current += character
  }
  pieces.push(current)
  const ranges = pieces.map((piece) => printAreaRange(piece, whole, extents))
  if (quoted || pieces.length > 100 || ranges.some((range) => !range)) {
    throw new Error('The saved print area is not a supported cell range. Select the cells to print and choose Selection.')
  }
  return ranges
}

function normalizedBounds(value) {
  if (!value || typeof value !== 'object') return null
  const top = Math.max(0, Math.min(1_048_575, Math.trunc(Number(value.top))))
  const bottom = Math.max(0, Math.min(1_048_575, Math.trunc(Number(value.bottom))))
  const left = Math.max(0, Math.min(16_383, Math.trunc(Number(value.left))))
  const right = Math.max(0, Math.min(16_383, Math.trunc(Number(value.right))))
  if (![top, bottom, left, right].every(Number.isFinite)) return null
  return {
    top: Math.min(top, bottom),
    bottom: Math.max(top, bottom),
    left: Math.min(left, right),
    right: Math.max(left, right),
  }
}

function usedBounds(sheet) {
  let top = Infinity
  let bottom = -1
  let left = Infinity
  let right = -1
  const include = (row, col) => {
    top = Math.min(top, row)
    bottom = Math.max(bottom, row)
    left = Math.min(left, col)
    right = Math.max(right, col)
  }
  for (const address of Object.keys(sheet?.cells || {})) {
    const coord = parseAddress(address)
    if (coord) include(coord.row, coord.col)
  }
  for (const merge of sheet?.merges || []) {
    const range = parseRange(merge)
    if (!range) continue
    include(range.top, range.left)
    include(range.bottom, range.right)
  }
  return bottom >= 0 ? { top, bottom, left, right } : { top: 0, bottom: 0, left: 0, right: 0 }
}

function clampInteger(value, min, max, fallback) {
  if (value === null || value === undefined || value === '') return fallback
  const number = Math.trunc(Number(value))
  return Number.isFinite(number) ? Math.max(min, Math.min(max, number)) : fallback
}

/** Custom margins in inches (0-10 each); null when any side is missing or invalid. */
function customMarginsOf(value) {
  if (!value || typeof value !== 'object') return null
  const margins = {}
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const number = Number(value[side])
    if (value[side] === null || value[side] === '' || !Number.isFinite(number) || number < 0 || number > 10) return null
    margins[side] = number
  }
  for (const side of ['header', 'footer']) {
    const number = Number(value[side])
    if (value[side] !== null && value[side] !== undefined && value[side] !== '' && Number.isFinite(number) && number >= 0 && number <= 10) margins[side] = number
  }
  return margins
}

/** Pages to print, 1-based and inclusive; `to` may be omitted for "to the end". */
function pageRangeOf(value) {
  if (!value || typeof value !== 'object') return null
  const from = clampInteger(value.from, 1, 1_000_000, 1)
  const to = clampInteger(value.to, 1, 1_000_000, null)
  if (from === 1 && to === null) return null
  return { from, to }
}

/**
 * Printer settings for the native job (never used for layout): an explicitly chosen printer
 * queue, copies and collation. Without deviceName the job goes to the Windows default.
 */
function printJobOf(value) {
  const input = value && typeof value === 'object' ? value : {}
  const job = { copies: clampInteger(input.copies, 1, MAX_PRINT_COPIES, 1), collate: input.collate !== false }
  const deviceName = typeof input.deviceName === 'string' ? input.deviceName.trim() : ''
  if (deviceName && deviceName.length <= 256 && !/[\u0000-\u001f\u007f]/.test(deviceName)) job.deviceName = deviceName
  return job
}

function safeOptions(value) {
  const input = value && typeof value === 'object' ? value : {}
  const scaling = SCALING_MODES.has(input.scaling) ? input.scaling : 'fit-width'
  const customMargins = input.margins === 'custom' ? customMarginsOf(input.customMargins) : null
  const options = {
    scope: ['active-sheet', 'selection', 'workbook'].includes(input.scope) ? input.scope : 'active-sheet',
    orientation: input.orientation === 'landscape' ? 'landscape' : 'portrait',
    scaling,
    paperSize: Object.hasOwn(PAPER_SIZES, input.paperSize) ? input.paperSize : 'letter',
    margins: Object.hasOwn(MARGIN_PRESETS, input.margins) || customMargins ? input.margins : 'normal',
    gridlines: input.gridlines !== false,
    headings: input.headings === true,
    useSavedLayout: input.useSavedLayout === true,
  }
  if (scaling === 'custom') options.scalePercent = clampInteger(input.scalePercent, 10, 400, 100)
  if (scaling === 'fit-pages') {
    options.fitWidth = clampInteger(input.fitWidth, 0, 100, 1)
    options.fitHeight = clampInteger(input.fitHeight, 0, 100, 0)
    // One page by one page is Fit Sheet on One Page; nothing constrained is actual size.
    if (options.fitWidth === 1 && options.fitHeight === 1) options.fitSheet = true
  }
  if (customMargins) options.customMargins = customMargins
  const pageRange = pageRangeOf(input.pageRange)
  if (pageRange) options.pageRange = pageRange
  return options
}

/** The layout rule a page model follows (separate from the reported `scaling` choice). */
function layoutMode(options) {
  if (options.scaling === 'fit-sheet' || options.fitSheet) return 'single'
  if (options.savedScale || options.scaling === 'custom') return 'grid'
  if (options.fitWidth || options.fitHeight) return 'grid'
  if (options.scaling === 'fit-width') return 'one-wide'
  if (options.scaling === 'fit-height') return 'one-tall'
  return 'grid'
}

// Manual page breaks apply on an axis whose page count is not fixed by a fit rule, which
// keeps Fit All Columns on One Page (the default) honouring row breaks, as Google Sheets does.
function honoursRowBreaks(options) {
  const mode = layoutMode(options)
  return mode !== 'single' && mode !== 'one-tall' && !options.fitHeight
}

function honoursColumnBreaks(options) {
  const mode = layoutMode(options)
  return mode !== 'single' && mode !== 'one-wide' && !options.fitWidth
}

function sheetOptions(sheet, options, warnings) {
  const result = { ...options }
  const setup = sheet.pageSetup || {}
  if (!options.useSavedLayout) return result
  const paperSize = { 1: 'letter', 5: 'legal', 9: 'a4' }[Number(setup.paperSize)]
  if (paperSize) result.paperSize = paperSize
  else if (setup.paperSize) warnings.add(`${sheet.name}: the saved paper size is not supported. Choose Letter, A4 or Legal in the print settings.`)
  if (['portrait', 'landscape'].includes(setup.orientation)) result.orientation = setup.orientation
  if (typeof setup.showGridLines === 'boolean') result.gridlines = setup.showGridLines
  if (typeof setup.showRowColHeaders === 'boolean') result.headings = setup.showRowColHeaders
  result.savedMargins = setup.margins
  result.horizontalCentered = setup.horizontalCentered === true
  result.verticalCentered = setup.verticalCentered === true
  result.pageOrder = setup.pageOrder
  // The saved layout replaces the manual scaling choice entirely.
  delete result.scalePercent
  delete result.fitWidth
  delete result.fitHeight
  delete result.fitSheet
  delete result.customMargins
  if (!Object.hasOwn(MARGIN_PRESETS, result.margins)) result.margins = 'normal'
  // The dormant fitToWidth/fitToHeight fields are often 1 even when Excel's
  // active mode is percentage scale. Only fitToPage enables those fields.
  if (setup.fitToPage === true) {
    result.fitWidth = Math.max(0, Math.min(100, Math.trunc(Number(setup.fitToWidth) || 0)))
    result.fitHeight = Math.max(0, Math.min(100, Math.trunc(Number(setup.fitToHeight) || 0)))
    result.scaling = result.fitWidth === 1 && result.fitHeight === 1 ? 'fit-sheet' : 'actual'
  } else if (Number.isFinite(Number(setup.scale)) && Number(setup.scale) >= 10 && Number(setup.scale) <= 400) {
    result.savedScale = Number(setup.scale) / 100
    result.scaling = 'actual'
  } else if (!['actual', 'fit-width', 'fit-sheet'].includes(result.scaling)) {
    // A file without a saved scale uses the dialog's default, never a hidden manual choice.
    result.scaling = 'fit-width'
  }
  return result
}

// Microsoft documents the text fields and formatting switches here:
// https://learn.microsoft.com/en-us/office/vba/excel/concepts/workbooks-and-worksheets/formatting-and-vba-codes-for-headers-and-footers
// Header/footer codes are a declarative text format. Never evaluate workbook
// formulas, scripts, markup, paths or external content while expanding fields.
function renderHeaderFooter(value, context, scale, warnings) {
  const sections = { left: [], center: [], right: [] }
  const states = Object.fromEntries(Object.keys(sections).map((key) => [key, { size: 11, family: 'Arial' }]))
  let section = 'center'
  let pending = ''
  const flush = () => {
    if (!pending) return
    const state = states[section]
    const css = [`font-size:${(state.size * scale).toFixed(2)}pt`, `font-family:"${state.family}",sans-serif`]
    if (state.bold) css.push('font-weight:700')
    if (state.italic) css.push('font-style:italic')
    if (state.color) css.push(`color:${state.color}`)
    const decorations = [state.underline ? 'underline' : '', state.strike ? 'line-through' : ''].filter(Boolean)
    if (decorations.length) css.push(`text-decoration:${decorations.join(' ')}`)
    if (state.underline === 'double') css.push('text-decoration-style:double')
    if (state.script) css.push(`vertical-align:${state.script};font-size:${(state.size * scale * 0.75).toFixed(2)}pt`)
    sections[section].push(`<span style="${escapeHtml(css.join(';'))}">${escapeHtml(pending)}</span>`)
    pending = ''
  }
  const text = String(value || '').slice(0, 32_768).replace(/\r\n?/g, '\n')
  for (let index = 0; index < text.length;) {
    if (text[index] !== '&' || index + 1 >= text.length) { pending += text[index++]; continue }
    const code = text[index + 1]
    if (code === '&') { pending += '&'; index += 2; continue }
    if (['L', 'C', 'R'].includes(code)) { flush(); section = { L: 'left', C: 'center', R: 'right' }[code]; index += 2; continue }
    if (code === '"') {
      const end = text.indexOf('"', index + 2)
      if (end < 0) { pending += '&'; index += 1; continue }
      flush()
      const [family, style = ''] = text.slice(index + 2, end).split(',')
      const safeFamily = family.replace(/["'\\;{}<>\r\n]/g, '').slice(0, 80).trim()
      if (safeFamily && !['+', '-'].includes(safeFamily)) states[section].family = safeFamily
      if (style) { states[section].bold = /bold/i.test(style); states[section].italic = /italic/i.test(style) }
      index = end + 1
      continue
    }
    const fontSize = /^&(\d{1,3})/.exec(text.slice(index))
    if (fontSize) { flush(); states[section].size = Math.max(1, Math.min(72, Number(fontSize[1]))); index += fontSize[0].length; continue }
    const color = /^&K([a-f0-9]{6})/i.exec(text.slice(index))
    if (color) { flush(); states[section].color = `#${color[1]}`; index += color[0].length; continue }
    if (['B', 'I', 'U', 'E', 'S', 'X', 'Y'].includes(code)) {
      flush()
      const state = states[section]
      if (code === 'B') state.bold = !state.bold
      if (code === 'I') state.italic = !state.italic
      if (code === 'S') state.strike = !state.strike
      if (code === 'U' || code === 'E') state.underline = state.underline ? false : code === 'E' ? 'double' : 'single'
      if (code === 'X' || code === 'Y') state.script = state.script ? '' : code === 'X' ? 'super' : 'sub'
      index += 2
      continue
    }
    if (code === 'G') { warnings.add(`${context.sheetName}: a header or footer picture cannot be rendered by this print view.`); index += 2; continue }
    const fields = { P: context.page, N: context.pages, F: context.filename, A: context.sheetName, Z: context.directory, D: context.date, T: context.time }
    if (Object.hasOwn(fields, code)) {
      let value = fields[code] ?? ''
      const offset = code === 'P' && /^([+-]\d+)/.exec(text.slice(index + 2))
      if (offset) value = Number(value) + Number(offset[1])
      pending += String(value)
      index += 2 + (offset ? offset[1].length : 0)
      continue
    }
    // Preserve unfamiliar text rather than silently discarding it.
    warnings.add(`${context.sheetName}: an unsupported header or footer code is shown as text.`)
    pending += `&${code}`
    index += 2
  }
  flush()
  if (!Object.values(sections).some((runs) => runs.length)) return ''
  return Object.entries(sections).map(([key, runs]) => `<div class="header-footer-${key}">${runs.join('')}</div>`).join('')
}

function colorToCss(value, themeColors = []) {
  if (!value) return ''
  if (typeof value === 'string') {
    const normalized = value.replace(/^#/, '')
    if (/^[0-9a-f]{8}$/i.test(normalized)) return `#${normalized.slice(2)}`
    if (/^[0-9a-f]{6}$/i.test(normalized)) return `#${normalized}`
    return ''
  }
  if (typeof value !== 'object') return ''
  if (typeof value.rgb === 'string' || typeof value.argb === 'string') return colorToCss(value.rgb || value.argb, themeColors)
  const themeIndex = Math.trunc(Number(value.theme))
  if (Number.isFinite(themeIndex) && themeColors[themeIndex]) return colorToCss(themeColors[themeIndex], themeColors)
  return ''
}

function borderDeclaration(side, themeColors) {
  if (!side || typeof side !== 'object' || !side.style) return ''
  const styles = {
    hair: '1px solid', thin: '1px solid', medium: '2px solid', thick: '3px solid',
    dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed', dashDotDot: '1px dashed',
    mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed',
    double: '3px double', slantDashDot: '2px dashed',
  }
  const declaration = styles[side.style] || '1px solid'
  return `${declaration} ${colorToCss(side.color, themeColors) || '#5f6361'}`
}

function cellStyle(cell, themeColors, valueType) {
  const style = cell?.style || {}
  const font = style.font || {}
  const fill = style.fill || {}
  const alignment = style.alignment || {}
  const border = style.border || {}
  const css = []
  const fontColor = colorToCss(font.color, themeColors)
  if (fontColor) css.push(`color:${fontColor}`)
  const background = colorToCss(fill.fgColor || fill.color, themeColors)
  if (background && fill.pattern !== 'none') css.push(`background-color:${background}`)
  if (font.bold) css.push('font-weight:700')
  if (font.italic) css.push('font-style:italic')
  const decorations = []
  if (font.underline) decorations.push('underline')
  if (font.strike) decorations.push('line-through')
  if (decorations.length) css.push(`text-decoration:${decorations.join(' ')}`)
  const rawFontSize = Number(font.size)
  if (Number.isFinite(rawFontSize) && rawFontSize > 0) {
    css.push(`font-size:${Math.max(6, Math.min(72, rawFontSize))}pt`)
  }
  if (typeof font.name === 'string' && font.name.trim()) {
    const family = font.name.replace(/["'\\;{}]/g, '').trim().slice(0, 80)
    if (family) css.push(`font-family:"${family}",sans-serif`)
  }
  const horizontal = String(alignment.horizontal || '').toLowerCase()
  if (valueType === 'number' && String(cell.numFmt || style.numFmt || '').includes('*')) css.push('text-align:right')
  else if (['left', 'center', 'right', 'justify'].includes(horizontal)) css.push(`text-align:${horizontal}`)
  else if (valueType === 'number') css.push('text-align:right')
  else if (valueType === 'boolean') css.push('text-align:center')
  if (horizontal === 'centercontinuous') css.push('text-align:center')
  const vertical = String(alignment.vertical || '').toLowerCase()
  if (vertical === 'top') css.push('vertical-align:top')
  else if (vertical === 'bottom') css.push('vertical-align:bottom')
  else if (vertical) css.push('vertical-align:middle')
  if (alignment.wrapText) css.push('white-space:pre-wrap;overflow-wrap:anywhere')
  else if (alignment.shrinkToFit) css.push('white-space:nowrap;font-size:8pt')
  else css.push('white-space:pre')
  const indent = Math.max(0, Math.min(20, Math.trunc(Number(alignment.indent) || 0)))
  if (indent) css.push(`padding-left:${4 + indent * 8}px`)
  for (const sideName of ['top', 'right', 'bottom', 'left']) {
    const declaration = borderDeclaration(border[sideName], themeColors)
    if (declaration) css.push(`border-${sideName}:${declaration}`)
  }
  return css.join(';')
}

function columnWidth(sheet, col) {
  const value = Number(sheet?.colWidths?.[String(col + 1)] ?? sheet?.properties?.defaultColWidth)
  if (!Number.isFinite(value) || value <= 0) return 64
  // Excel column units use digits in the workbook's Normal font, independent
  // of a cell's font. The print request can carry its locally measured digit
  // width; a fixed screen grid metric otherwise makes imported forms paginate
  // wider than Excel. This metric is ephemeral and never changes saved widths.
  const measuredWidth = Number(sheet?.properties?.printDigitWidth)
  const digitWidth = Number.isFinite(measuredWidth) && measuredWidth >= 1 && measuredWidth <= 40 ? measuredWidth : 8
  return Math.max(2, Math.min(1_000, value * digitWidth))
}

function rowHeight(sheet, row) {
  const value = Number(sheet?.rowHeights?.[String(row + 1)] ?? sheet?.properties?.defaultRowHeight)
  if (!Number.isFinite(value) || value <= 0) return 20
  return Math.max(2, Math.min(640, value * 4 / 3))
}

function cellContentHeight(cell, height, gridlines) {
  const width = (side) => {
    const style = cell?.style?.border?.[side]?.style
    if (!style) return gridlines ? 1 : 0
    return style === 'thick' || style === 'double' ? 3 : String(style).startsWith('medium') ? 2 : 1
  }
  // Collapsed table borders occupy half their stroke inside either adjacent
  // row. Bound content, including wrapped text, so its intrinsic line box
  // cannot expand a saved row height after pagination has already happened.
  return Math.max(0, height - 2 - (width('top') + width('bottom')) / 2)
}

function rangeIntersection(first, second) {
  const top = Math.max(first.top, second.top)
  const bottom = Math.min(first.bottom, second.bottom)
  const left = Math.max(first.left, second.left)
  const right = Math.min(first.right, second.right)
  return top <= bottom && left <= right ? { top, bottom, left, right } : null
}

function visibleIndices(start, end, hiddenValues) {
  // Workbook dimensions are stored with Excel's 1-based row/column numbering;
  // print iteration uses zero-based coordinates.
  const hidden = new Set((hiddenValues || [])
    .map((value) => Math.trunc(Number(value)) - 1)
    .filter((value) => Number.isFinite(value) && value >= 0))
  const values = []
  for (let index = start; index <= end; index += 1) if (!hidden.has(index)) values.push(index)
  return values
}

function mergeLayout(sheet, bounds, rows, cols) {
  const rowPosition = new Map(rows.map((value, index) => [value, index]))
  const colPosition = new Map(cols.map((value, index) => [value, index]))
  const masters = new Map()
  const covered = new Set()
  for (const mergeText of sheet.merges || []) {
    const original = parseRange(mergeText)
    const clipped = original && rangeIntersection(original, bounds)
    if (!clipped) continue
    const visibleRows = rows.filter((row) => row >= clipped.top && row <= clipped.bottom)
    const visibleCols = cols.filter((col) => col >= clipped.left && col <= clipped.right)
    if (!visibleRows.length || !visibleCols.length) continue
    const firstRow = visibleRows[0]
    const firstCol = visibleCols[0]
    const masterKey = `${firstRow}:${firstCol}`
    masters.set(masterKey, {
      rowspan: visibleRows.length,
      colspan: visibleCols.length,
      sourceAddress: addressOf(original.top, original.left),
      original,
    })
    for (const row of visibleRows) {
      for (const col of visibleCols) {
        if (row === firstRow && col === firstCol) continue
        if (rowPosition.has(row) && colPosition.has(col)) covered.add(`${row}:${col}`)
      }
    }
  }
  return { masters, covered }
}

function mergedRenderCell(sheet, merge, source) {
  if (!merge) return source
  const range = merge.original
  const border = { ...(source.style?.border || {}) }
  for (const side of ['top', 'right', 'bottom', 'left']) delete border[side]
  const weight = (side) => !side?.style ? 0 : /thick|double/.test(side.style) ? 3 : /medium/.test(side.style) ? 2 : 1
  const include = (address, side) => {
    const candidate = sheet.cells?.[address]?.style?.border?.[side]
    if (weight(candidate) > weight(border[side])) border[side] = candidate
  }
  // Imported merged cells often store the closing edge on the last covered
  // cell, not the master. Gather all four outside edges before painting.
  for (const address of Object.keys(sheet.cells || {})) {
    const position = parseAddress(address)
    if (!position || position.row < range.top || position.row > range.bottom || position.col < range.left || position.col > range.right) continue
    if (position.row === range.top) include(address, 'top')
    if (position.row === range.bottom) include(address, 'bottom')
    if (position.col === range.left) include(address, 'left')
    if (position.col === range.right) include(address, 'right')
  }
  // Simple's own merged cells may store the complete rectangle on the master
  // alone. Prefer explicit outer cells, then use the same master fallback as
  // the editor so newly created boxes do not lose their bottom/right edges.
  for (const side of ['top', 'right', 'bottom', 'left']) if (!border[side]) border[side] = source.style?.border?.[side]
  return { ...source, style: { ...(source.style || {}), border } }
}

function printDisplayParts(cell, display, supplied) {
  const raw = cell.formula ? cell.result : cell.value
  const type = ['number', 'boolean', 'text'].includes(supplied?.type) ? supplied.type : typeof raw === 'number' ? 'number' : typeof raw === 'boolean' ? 'boolean' : 'text'
  if (type !== 'number') return { type }
  if (typeof supplied?.accounting?.symbol === 'string' && typeof supplied?.accounting?.amount === 'string') return { type, accounting: supplied.accounting }
  // Direct backend callers can use the same accounting rule as the grid.
  // Numeric type is determined above from data, never a currency-looking label.
  const format = String(cell.numFmt || cell.style?.numFmt || '')
  const symbol = format.includes('*') && format.match(/[$€£¥₹₩₽]/)?.[0]
  const text = String(display || '').trim()
  const index = symbol ? text.indexOf(symbol) : -1
  return { type, ...(index >= 0 ? { accounting: { symbol, amount: (text.slice(0, index) + text.slice(index + symbol.length)).trim() } } : {}) }
}

function overflowWidth(sheet, row, colIndex, cols, cell, merge, layout, displayValues, parts) {
  const first = cols[colIndex]
  let width = columnWidth(sheet, first)
  const alignment = cell.style?.alignment || {}
  if (merge || parts.type !== 'text' || alignment.wrapText || alignment.clipText || alignment.shrinkToFit || alignment.textRotation || ['right', 'center', 'centercontinuous', 'justify', 'distributed', 'fill'].includes(String(alignment.horizontal || '').toLowerCase())) return width
  if (cell.style?.border?.right?.style) return width
  for (let index = colIndex + 1; index < cols.length; index += 1) {
    const col = cols[index]
    const key = `${row}:${col}`
    if (layout.masters.has(key) || layout.covered.has(key)) break
    const address = addressOf(row, col)
    const next = sheet.cells?.[address] || {}
    const display = displayValues?.[address] ?? next.display ?? next.result ?? next.value ?? ''
    if (next.formula || display !== '' || (next.value !== undefined && next.value !== null && next.value !== '') || next.style?.border?.left?.style) break
    width += columnWidth(sheet, col)
    if (next.style?.border?.right?.style) break
  }
  return width
}

function pageMetrics(options) {
  const paper = PAPER_SIZES[options.paperSize]
  const custom = options.margins === 'custom' && options.customMargins
  const margins = custom
    ? { header: Math.min(0.2, custom.top), footer: Math.min(0.2, custom.bottom), ...custom }
    : { ...MARGIN_PRESETS[options.margins] || MARGIN_PRESETS.normal, header: 0.2, footer: 0.2 }
  for (const side of ['top', 'right', 'bottom', 'left', 'header', 'footer']) {
    const value = Number(options.savedMargins?.[side])
    if (Number.isFinite(value) && value >= 0 && value <= 10) margins[side] = value
  }
  const widthInches = options.orientation === 'landscape' ? paper.height : paper.width
  const heightInches = options.orientation === 'landscape' ? paper.width : paper.height
  if (margins.left + margins.right >= widthInches || margins.top + margins.bottom >= heightInches) {
    throw new Error(custom
      ? 'These margins leave no room for cells. Choose smaller margins.'
      : 'The saved page margins leave no room for cells. Turn off Use saved page layout and choose smaller margins.')
  }
  return {
    paper,
    margins,
    widthInches,
    heightInches,
    width: widthInches * 96,
    height: heightInches * 96,
    contentWidth: Math.max(1, (widthInches - margins.left - margins.right) * 96),
    contentHeight: Math.max(1, (heightInches - margins.top - margins.bottom) * 96),
  }
}

const HEADING_WIDTH = 38
const HEADING_HEIGHT = 24

/** Title rows/columns that are not already part of the printed indices, in order. */
function prependedTitles(indices, titles) {
  if (!titles.length) return []
  const present = new Set(indices)
  return titles.filter((index) => !present.has(index))
}

function scaleFor(sheet, rows, cols, options, metrics) {
  if (options.savedScale) return options.savedScale
  if (options.scaling === 'custom') return Math.max(0.1, Math.min(4, (Number(options.scalePercent) || 100) / 100))
  const headX = options.headings ? HEADING_WIDTH : 0
  const headY = options.headings ? HEADING_HEIGHT : 0
  if ((options.fitWidth || options.fitHeight) && !options.fitSheet) {
    const width = cols.reduce((sum, col) => sum + columnWidth(sheet, col), headX)
    const height = rows.reduce((sum, row) => sum + rowHeight(sheet, row), headY)
    let scale = Math.min(1,
      options.fitWidth ? metrics.contentWidth * options.fitWidth / Math.max(1, width + TABLE_EDGE_ALLOWANCE) : Infinity,
      options.fitHeight ? metrics.contentHeight * options.fitHeight / Math.max(1, height + TABLE_EDGE_ALLOWANCE) : Infinity)
    // Account for whole rows/columns at page boundaries, rather than assuming a
    // cell may split over two pages when fitting to a saved page count.
    const titleRows = repeatedRows(sheet, options)
    const titleCols = repeatedColumns(sheet, options)
    for (let attempt = 0; attempt < 100; attempt += 1) {
      const across = options.fitWidth ? chunkColumns(sheet, cols, metrics.contentWidth / scale - TABLE_EDGE_ALLOWANCE - headX, titleCols).length : 0
      const down = options.fitHeight ? chunkRows(sheet, rows, metrics.contentHeight / scale - TABLE_EDGE_ALLOWANCE - headY, titleRows).length : 0
      if ((!options.fitWidth || across <= options.fitWidth) && (!options.fitHeight || down <= options.fitHeight)) break
      scale *= 0.99
    }
    return Math.max(0.01, scale)
  }
  // Fit to automatic x automatic pages is actual size; 1 x 1 is Fit Sheet on One Page below.
  if (options.scaling === 'actual' || (options.scaling === 'fit-pages' && !options.fitSheet)) return 1
  // Fit width/height/sheet: repeated titles outside the printed range are part of every page.
  const width = [...prependedTitles(cols, repeatedColumns(sheet, options)), ...cols].reduce((sum, col) => sum + columnWidth(sheet, col), headX)
  const height = [...prependedTitles(rows, repeatedRows(sheet, options)), ...rows].reduce((sum, row) => sum + rowHeight(sheet, row), headY)
  const widthScale = metrics.contentWidth / Math.max(1, width + TABLE_EDGE_ALLOWANCE)
  const heightScale = metrics.contentHeight / Math.max(1, height + TABLE_EDGE_ALLOWANCE)
  const value = layoutMode(options) === 'single' ? Math.min(widthScale, heightScale) : options.scaling === 'fit-height' ? heightScale : widthScale
  return Math.max(0.01, Math.min(1, value))
}

/**
 * Split printed rows or columns into pages. `titles` repeat at the start of every
 * continuation page (and lead the first page when they lie outside the printed range);
 * `breaks` holds the indices that must start a new page (manual page breaks).
 */
function chunkAxis(indices, sizeFor, capacity, titles = [], breaks = null, fullMessage = '') {
  if (!indices.length) return [[]]
  const titleSize = titles.reduce((sum, index) => sum + Math.max(1, sizeFor(index)), 0)
  if (titles.length && titleSize >= capacity) throw new Error(fullMessage)
  const chunks = []
  const lead = prependedTitles(indices, titles)
  let current = lead
  let consumed = lead.reduce((sum, index) => sum + Math.max(1, sizeFor(index)), 0)
  let body = 0
  for (const index of indices) {
    if (body && current.includes(index)) continue
    const size = Math.max(1, sizeFor(index))
    const forced = Boolean(breaks && body && breaks.has(index))
    if (body && (forced || consumed + size > capacity)) {
      chunks.push(current)
      current = [...titles]
      consumed = titleSize
      body = 0
      if (current.includes(index)) continue
    }
    current.push(index)
    consumed += size
    body += 1
  }
  if (current.length) chunks.push(current)
  return chunks
}

function chunkIndices(indices, sizeFor, capacity) {
  return chunkAxis(indices, sizeFor, capacity)
}

function repeatedRows(sheet, options) {
  if (options.scope === 'selection') return []
  const match = /^\$?(\d+):\$?(\d+)$/.exec(String(sheet.pageSetup?.printTitlesRow || ''))
  if (!match) return []
  const first = Number(match[1]), last = Number(match[2])
  if (first < 1 || last < first || last > 1048576 || last - first > 4095) throw new Error('The saved rows to repeat are too large. Choose a smaller print title range.')
  return visibleIndices(first - 1, last - 1, sheet.hiddenRows)
}

function repeatedColumns(sheet, options) {
  if (options.scope === 'selection') return []
  const match = /^\$?([A-Z]{1,3})(?::\$?([A-Z]{1,3}))?$/i.exec(String(sheet.pageSetup?.printTitlesColumn || '').trim())
  if (!match) return []
  const first = columnIndexOf(match[1]), last = columnIndexOf(match[2] || match[1])
  if (first < 0 || last < first || last > 16_383 || last - first > 1023) throw new Error('The saved columns to repeat are too large. Choose a smaller print title range.')
  // Hidden columns are stored 1-based, like hidden rows.
  return visibleIndices(first, last, sheet.hiddenCols)
}

function chunkRows(sheet, rows, capacity, titles, breaks = null) {
  return chunkAxis(rows, (row) => rowHeight(sheet, row), capacity, titles, breaks, 'The repeated heading rows fill the page. Use a smaller scale or fewer heading rows.')
}

function chunkColumns(sheet, cols, capacity, titles, breaks = null) {
  return chunkAxis(cols, (col) => columnWidth(sheet, col), capacity, titles, breaks, 'The repeated heading columns fill the page. Use a smaller scale or fewer heading columns.')
}

/**
 * Manual page breaks as the 0-based indices that start a new page. The model keeps
 * ExcelJS's shape ({ id, max, man } with id = the 1-based row above the break, which is
 * the 0-based index of the first row after it); plain numbers mean the same.
 */
function manualBreaks(values) {
  if (!Array.isArray(values) || !values.length) return null
  const breaks = new Set()
  for (const value of values.slice(0, 1026)) {
    const id = Math.trunc(Number(value && typeof value === 'object' ? value.id : value))
    if (Number.isFinite(id) && id >= 1 && id <= 1_048_575) breaks.add(id)
  }
  return breaks.size ? breaks : null
}

function paginateSheet(sheet, bounds, options, metrics) {
  const rows = visibleIndices(bounds.top, bounds.bottom, sheet.hiddenRows)
  const cols = visibleIndices(bounds.left, bounds.right, sheet.hiddenCols)
  const scale = scaleFor(sheet, rows, cols, options, metrics)
  const mode = layoutMode(options)
  if (mode === 'single') return [{ rows, cols, scale }]

  const unscaledWidth = metrics.contentWidth / scale - TABLE_EDGE_ALLOWANCE - (options.headings ? HEADING_WIDTH : 0)
  const unscaledHeight = metrics.contentHeight / scale - TABLE_EDGE_ALLOWANCE - (options.headings ? HEADING_HEIGHT : 0)
  const titleRows = repeatedRows(sheet, options)
  const titleCols = repeatedColumns(sheet, options)
  const columnChunks = mode === 'one-wide'
    ? [[...prependedTitles(cols, titleCols), ...cols]]
    : chunkColumns(sheet, cols, Math.max(1, unscaledWidth), titleCols, honoursColumnBreaks(options) ? manualBreaks(sheet.colBreaks) : null)
  const rowChunks = mode === 'one-tall'
    ? [[...prependedTitles(rows, titleRows), ...rows]]
    : chunkRows(sheet, rows, Math.max(1, unscaledHeight), titleRows, honoursRowBreaks(options) ? manualBreaks(sheet.rowBreaks) : null)
  const pages = []
  if (options.pageOrder === 'downThenOver') {
    for (const columnChunk of columnChunks) for (const rowChunk of rowChunks) pages.push({ rows: rowChunk, cols: columnChunk, scale })
  } else for (const rowChunk of rowChunks) for (const columnChunk of columnChunks) pages.push({ rows: rowChunk, cols: columnChunk, scale })
  return pages.length ? pages : [{ rows: [], cols: [], scale }]
}

// ---------------------------------------------------------------------------
// Charts on printed pages.
// Payload contract: input.charts = { [sheetId]: [{ id?, svg, from, to }] } (or an
// array of { sheetId, svg, from, to }), where `svg` is renderChartSvg() output
// (src/lib/chart-render.ts) and from/to are chart anchors: 0-based
// { row, col, rowOffsetEmu?, colOffsetEmu? } (1px = 9525 EMU). Each SVG is
// scaled into the anchored rectangle and clipped to the page's cell area.
// ---------------------------------------------------------------------------
const MAX_PRINT_CHARTS_PER_SHEET = 200
const MAX_PRINT_CHART_SVG = 4 * 1024 * 1024
const EMU_PER_PIXEL = 9525

// Chart SVG is an allowlist of drawing elements: no scripts, event handlers,
// links, styles or external references (text content is already escaped).
const PRINT_SVG_ELEMENTS = new Set(['svg', 'g', 'defs', 'clippath', 'rect', 'path', 'text', 'tspan', 'title', 'desc', 'circle', 'ellipse', 'line', 'polyline', 'polygon'])

function safeChartSvg(svg) {
  if (typeof svg !== 'string') return null
  const text = svg.trim()
  if (!text.startsWith('<svg') || text.length > MAX_PRINT_CHART_SVG) return null
  if (/<!|<\?/.test(text)) return null
  for (const tag of text.matchAll(/<\s*\/?\s*([A-Za-z][\w:.-]*)([^>]*)>/g)) {
    if (!PRINT_SVG_ELEMENTS.has(tag[1].toLowerCase())) return null
    for (const attribute of tag[2].matchAll(/([^\s=/]+)\s*=\s*("[^"]*"|'[^']*')/g)) {
      const name = attribute[1].toLowerCase()
      if (name.startsWith('on') || name.includes('href') || name === 'style') return null
      if (/javascript:|url\(\s*['"]?(?!#)/i.test(attribute[2])) return null
    }
  }
  return text
}

function anchorPoint(value) {
  if (!value || typeof value !== 'object') return null
  const row = Math.trunc(Number(value.row)), col = Math.trunc(Number(value.col))
  if (!Number.isFinite(row) || !Number.isFinite(col) || row < 0 || col < 0 || row > 1_048_575 || col > 16_383) return null
  return { row, col, rowOffsetEmu: Math.max(0, Number(value.rowOffsetEmu) || 0), colOffsetEmu: Math.max(0, Number(value.colOffsetEmu) || 0) }
}

// Pictures: input.images = { [sheetId]: [{ src, from, to }] } where src is a base64 PNG, JPEG
// or GIF data URL (the charset check keeps markup out of the attribute).
const PRINT_IMAGE_SOURCE = /^data:image\/(?:png|jpe?g|gif);base64,[A-Za-z0-9+/=]+$/
const MAX_PRINT_IMAGE = 24 * 1024 * 1024

function printImagesFor(input, sheet) {
  const source = input && input.images
  const list = source && typeof source === 'object' && Array.isArray(source[sheet.id]) ? source[sheet.id] : []
  const images = []
  for (const item of list.slice(0, MAX_PRINT_CHARTS_PER_SHEET)) {
    const src = item && typeof item.src === 'string' && item.src.length <= MAX_PRINT_IMAGE && PRINT_IMAGE_SOURCE.test(item.src) ? item.src : null
    const from = anchorPoint(item && item.from), to = anchorPoint(item && item.to)
    if (!src || !from || !to || to.row < from.row || to.col < from.col) continue
    images.push({ svg: `<img class="print-image" alt="" src="${src}">`, from, to })
  }
  return images
}

function printChartsFor(input, sheet) {
  const source = input && input.charts
  let list = []
  if (Array.isArray(source)) list = source.filter((item) => item && item.sheetId === sheet.id)
  else if (source && typeof source === 'object' && Array.isArray(source[sheet.id])) list = source[sheet.id]
  const charts = []
  for (const item of list.slice(0, MAX_PRINT_CHARTS_PER_SHEET)) {
    const svg = safeChartSvg(item && item.svg)
    const from = anchorPoint(item && item.from), to = anchorPoint(item && item.to)
    if (!svg || !from || !to || to.row < from.row || to.col < from.col) continue
    charts.push({ svg, from, to })
  }
  return charts
}

/** Excel's default print area includes drawing objects, not only cells. */
function boundsIncludingCharts(bounds, charts) {
  if (!charts.length) return bounds
  const result = { ...bounds }
  for (const chart of charts) {
    result.top = Math.min(result.top, chart.from.row)
    result.left = Math.min(result.left, chart.from.col)
    result.bottom = Math.max(result.bottom, chart.to.row)
    result.right = Math.max(result.right, chart.to.col)
  }
  return result
}

function axisPositions(count, sizeOf, hiddenValues) {
  const hidden = new Set((hiddenValues || []).map((value) => Math.trunc(Number(value)) - 1))
  const positions = new Float64Array(count + 1)
  for (let index = 0; index < count; index += 1) positions[index + 1] = positions[index] + (hidden.has(index) ? 0 : sizeOf(index))
  return { at: (index) => positions[Math.max(0, Math.min(count, index))], size: (index) => (hidden.has(index) ? 0 : sizeOf(index)) }
}

/**
 * Leading repeated titles on a continuation page (they sit before the page's body). On
 * the first page the titles are simply the first printed rows/columns: no prefix.
 */
function leadingTitleCount(indices, titles) {
  if (!titles.size) return 0
  let prefix = 0
  while (prefix < indices.length && titles.has(indices[prefix])) prefix += 1
  if (prefix >= indices.length || (prefix > 0 && indices[prefix] === indices[prefix - 1] + 1)) return 0
  return prefix
}

function chartOverlayHtml(sheet, rows, cols, options, charts) {
  if (!charts.length || !rows.length || !cols.length) return ''
  const headX = options.headings ? HEADING_WIDTH : 0
  const headY = options.headings ? HEADING_HEIGHT : 0
  let maxCol = cols.reduce((max, col) => Math.max(max, col), 0), maxRow = rows.reduce((max, row) => Math.max(max, row), 0)
  for (const chart of charts) { maxCol = Math.max(maxCol, chart.to.col); maxRow = Math.max(maxRow, chart.to.row) }
  const xs = axisPositions(maxCol + 2, (col) => columnWidth(sheet, col), sheet.hiddenCols)
  const ys = axisPositions(maxRow + 2, (row) => rowHeight(sheet, row), sheet.hiddenRows)
  // Repeated title rows/columns prepended to a later page sit above/left of its body.
  const rowPrefix = leadingTitleCount(rows, new Set(repeatedRows(sheet, options)))
  const colPrefix = leadingTitleCount(cols, new Set(repeatedColumns(sheet, options)))
  const bodyTop = headY + rows.slice(0, rowPrefix).reduce((sum, row) => sum + rowHeight(sheet, row), 0)
  const bodyLeft = headX + cols.slice(0, colPrefix).reduce((sum, col) => sum + columnWidth(sheet, col), 0)
  const width = cols.slice(colPrefix).reduce((sum, col) => sum + columnWidth(sheet, col), 0)
  const height = rows.slice(rowPrefix).reduce((sum, row) => sum + rowHeight(sheet, row), 0)
  const originX = xs.at(cols[colPrefix])
  const originY = ys.at(rows[rowPrefix])
  const place = (point) => ({
    x: xs.at(point.col) + Math.min(xs.size(point.col), point.colOffsetEmu / EMU_PER_PIXEL) - originX,
    y: ys.at(point.row) + Math.min(ys.size(point.row), point.rowOffsetEmu / EMU_PER_PIXEL) - originY,
  })
  const items = []
  for (const chart of charts) {
    const a = place(chart.from), b = place(chart.to)
    const w = b.x - a.x, h = b.y - a.y
    if (w < 1 || h < 1 || a.x >= width || a.y >= height || b.x <= 0 || b.y <= 0) continue
    items.push(`<div class="print-chart" style="left:${a.x.toFixed(2)}px;top:${a.y.toFixed(2)}px;width:${w.toFixed(2)}px;height:${h.toFixed(2)}px">${chart.svg}</div>`)
  }
  if (!items.length) return ''
  return `<div class="print-charts" aria-hidden="false" style="left:${bodyLeft.toFixed(2)}px;top:${bodyTop.toFixed(2)}px;width:${width.toFixed(2)}px;height:${height.toFixed(2)}px">${items.join('')}</div>`
}

// ---------------------------------------------------------------------------
// Resolved cell visuals: conditional formatting (fills, colour scales, data bars, icon
// sets, fonts, borders), table styles and banding, number-format colours, checkboxes and
// sparklines, resolved in the renderer by src/lib/print-visuals.ts with the same rules the
// grid draws (src/lib/visual-style.ts).
// Payload contract: input.visuals = {
//   icons: { 'set:index': '<svg…>' },                 // serialized ConditionalIcon glyphs
//   sheets: { [sheetId]: {
//     styles: [{ fill?, color?, bold?, italic?, decoration?, borders?, bar?, icon?, hide?, checkbox? }],
//     cells: { [address]: styleIndex },
//     text?: { [address]: string },                  // the value as a CF number format shows it
//     sparklines?: { [address]: '<svg…>' },          // drawn at the printed cell size
//   } } }
// Everything is validated here: colours are #RRGGBB, borders and data-bar backgrounds follow
// a strict CSS grammar (no url() or other external references), SVG passes the chart allowlist.
// ---------------------------------------------------------------------------
const MAX_VISUAL_STYLES = 50_000
const MAX_VISUAL_ICONS = 400
const MAX_VISUAL_ICON_SVG = 16 * 1024
const MAX_VISUAL_SVG = 256 * 1024
const MAX_VISUAL_SVG_TOTAL = 32 * 1024 * 1024
const MAX_VISUAL_TEXT = 32_767
const HEX_COLOR = /^#[0-9a-f]{6}$/i
const CSS_BORDER = /^(?:[0-9]|10)(?:\.[0-9]{1,3})?px (?:solid|dashed|dotted|double) #[0-9a-f]{6}$/i
const ICON_KEY = /^[0-9A-Za-z]{1,32}:[0-9]{1,2}$/
const TEXT_DECORATIONS = new Set(['underline', 'line-through', 'underline line-through'])
const BACKGROUND_FUNCTIONS = new Set(['linear-gradient', 'repeating-linear-gradient', 'calc'])

/** A data-bar background value: gradients of hex colours, lengths and keywords only. */
function safeBackgroundValue(value) {
  if (typeof value !== 'string' || !value || value.length > 2_000) return ''
  if (!/^[#a-z0-9%.,()\s+-]+$/i.test(value)) return ''
  const functions = [...value.matchAll(/([a-z-]+)\s*\(/gi)]
  if (functions.some((match) => !BACKGROUND_FUNCTIONS.has(match[1].toLowerCase()))) return ''
  // Every "(" opens an allowed function: no bare groups.
  if ((value.match(/\(/g) || []).length !== functions.length) return ''
  return value
}

function safeIconTable(source) {
  const icons = new Map()
  if (!source || typeof source !== 'object') return icons
  for (const [key, svg] of Object.entries(source).slice(0, MAX_VISUAL_ICONS)) {
    if (!ICON_KEY.test(key) || typeof svg !== 'string' || svg.length > MAX_VISUAL_ICON_SVG) continue
    const safe = safeChartSvg(svg)
    if (safe) icons.set(key, safe)
  }
  return icons
}

function safeVisualStyle(value, icons) {
  if (!value || typeof value !== 'object') return null
  const css = []
  if (typeof value.fill === 'string' && HEX_COLOR.test(value.fill)) css.push(`background-color:${value.fill}`)
  if (typeof value.color === 'string' && HEX_COLOR.test(value.color)) css.push(`color:${value.color}`)
  if (typeof value.bold === 'boolean') css.push(`font-weight:${value.bold ? 700 : 400}`)
  if (typeof value.italic === 'boolean') css.push(`font-style:${value.italic ? 'italic' : 'normal'}`)
  if (TEXT_DECORATIONS.has(value.decoration)) css.push(`text-decoration:${value.decoration}`)
  if (value.borders && typeof value.borders === 'object') {
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const border = value.borders[side]
      if (typeof border === 'string' && CSS_BORDER.test(border)) css.push(`border-${side}:${border}`)
    }
  }
  if (value.bar && typeof value.bar === 'object') {
    const image = safeBackgroundValue(value.bar.image)
    const size = safeBackgroundValue(value.bar.size)
    const position = safeBackgroundValue(value.bar.position)
    const repeat = safeBackgroundValue(value.bar.repeat)
    if (image && size && position && repeat) css.push(`background-image:${image}`, `background-size:${size}`, `background-position:${position}`, `background-repeat:${repeat}`)
  }
  const style = {
    css: css.join(';'),
    icon: typeof value.icon === 'string' && icons.has(value.icon) ? icons.get(value.icon) : '',
    hide: value.hide === true,
    checkbox: typeof value.checkbox === 'boolean' ? value.checkbox : null,
  }
  return style.css || style.icon || style.hide || style.checkbox !== null ? style : null
}

/** Validated visuals for one sheet, or null when the renderer sent none. */
function sheetVisualsFor(input, sheet, icons, budget) {
  const sheets = input && input.visuals && typeof input.visuals === 'object' ? input.visuals.sheets : null
  const source = sheets && typeof sheets === 'object' && Object.hasOwn(sheets, sheet.id) ? sheets[sheet.id] : null
  if (!source || typeof source !== 'object') return null
  const styles = Array.isArray(source.styles) ? source.styles.slice(0, MAX_VISUAL_STYLES).map((style) => safeVisualStyle(style, icons)) : []
  const visuals = { cells: new Map(), text: new Map(), sparklines: new Map() }
  const entries = (value) => (value && typeof value === 'object' ? Object.entries(value).slice(0, MAX_PRINT_CELLS) : [])
  const key = (address) => {
    const position = parseAddress(address)
    return position ? addressOf(position.row, position.col) : ''
  }
  for (const [address, index] of entries(source.cells)) {
    const style = Number.isInteger(index) ? styles[index] : null
    const target = key(address)
    if (style && target) visuals.cells.set(target, style)
  }
  for (const [address, text] of entries(source.text)) {
    const target = key(address)
    if (target && typeof text === 'string' && text.length <= MAX_VISUAL_TEXT) visuals.text.set(target, text)
  }
  for (const [address, svg] of entries(source.sparklines)) {
    const target = key(address)
    if (!target || typeof svg !== 'string' || svg.length > MAX_VISUAL_SVG || budget.svg + svg.length > MAX_VISUAL_SVG_TOTAL) continue
    const safe = safeChartSvg(svg)
    if (!safe) continue
    budget.svg += safe.length
    visuals.sparklines.set(target, safe)
  }
  return visuals
}

/** A printed checkbox, drawn like the grid's checkbox control. */
function checkboxSvg(checked) {
  return checked
    ? '<svg viewBox="0 0 14 14" width="13" height="13" role="img" aria-label="Checked"><rect x="0.75" y="0.75" width="12.5" height="12.5" rx="2.5" fill="#476b57" stroke="#476b57" stroke-width="1.5"/><path d="M3.6 7.3 L6 9.6 L10.5 4.7" fill="none" stroke="#ffffff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>'
    : '<svg viewBox="0 0 14 14" width="13" height="13" role="img" aria-label="Not checked"><rect x="0.75" y="0.75" width="12.5" height="12.5" rx="2.5" fill="#ffffff" stroke="#6b756f" stroke-width="1.2"/></svg>'
}

function checkedValue(value) {
  return value === true || String(value).toUpperCase() === 'TRUE'
}

function renderPage(sheet, bounds, rows, cols, displayValues, options, themeColors, metrics, pageNumber, sheetPage, scale, pageName, headerContext, warnings, displayParts, charts = [], visuals = null) {
  // Repeated title rows/columns may lie outside the printed area; their merges still apply.
  const layout = mergeLayout(sheet, {
    top: rows.reduce((min, row) => Math.min(min, row), bounds.top),
    bottom: rows.reduce((max, row) => Math.max(max, row), bounds.bottom),
    left: cols.reduce((min, col) => Math.min(min, col), bounds.left),
    right: cols.reduce((max, col) => Math.max(max, col), bounds.right),
  }, rows, cols)
  const { masters, covered } = layout
  const columns = []
  if (options.headings) columns.push('<col class="row-heading-column" style="width:38px">')
  for (const col of cols) columns.push(`<col style="width:${columnWidth(sheet, col).toFixed(2)}px">`)
  const head = options.headings
    ? `<thead><tr><th class="corner-heading" aria-hidden="true"></th>${cols.map((col) => `<th class="column-heading" scope="col">${columnName(col)}</th>`).join('')}</tr></thead>`
    : ''
  const body = rows.map((row, rowIndex) => {
    const cells = []
    if (options.headings) cells.push(`<th class="row-heading" scope="row"><span style="max-height:${Math.max(0, rowHeight(sheet, row) - 3).toFixed(2)}px">${row + 1}</span></th>`)
    for (const [colIndex, col] of cols.entries()) {
      const key = `${row}:${col}`
      if (covered.has(key)) continue
      const merge = masters.get(key)
      const address = merge?.sourceAddress || addressOf(row, col)
      const cell = mergedRenderCell(sheet, merge, sheet.cells?.[address] || {})
      const visual = visuals?.cells.get(address) || null
      const sparkline = visuals?.sparklines.get(address) || ''
      // A checkbox cell control (cell.type) prints as a box even without renderer visuals.
      const checkbox = visual && visual.checkbox !== null
        ? visual.checkbox
        : cell.type === 'checkbox' ? checkedValue(cell.formula ? cell.result : cell.value) : null
      let value = displayValues?.[address] ?? cell.display ?? cell.result ?? cell.value ?? ''
      if (typeof value === 'string' && value.startsWith(SPARKLINE_MARKER)) value = ''
      const override = visuals?.text.get(address)
      if (override !== undefined) value = override
      if (visual?.hide || checkbox !== null) value = ''
      const parts = printDisplayParts(cell, value, override === undefined ? displayParts?.[address] : undefined)
      const accounting = value !== '' ? parts.accounting : undefined
      const height = rows.slice(rowIndex, rowIndex + (merge?.rowspan || 1)).reduce((sum, item) => sum + rowHeight(sheet, item), 0)
      const contentHeight = cellContentHeight(cell, height, options.gridlines)
      const availableWidth = value !== '' ? overflowWidth(sheet, row, colIndex, cols, cell, merge, layout, displayValues, parts) : columnWidth(sheet, col)
      const overflowing = !merge && availableWidth > columnWidth(sheet, col)
      const classes = [overflowing ? 'print-overflow' : '', sparkline ? 'has-sparkline' : ''].filter(Boolean).join(' ')
      const css = cellStyle(cell, themeColors, parts.type) + (visual?.css ? `;${visual.css}` : '')
      const attributes = [
        `data-address="${escapeHtml(addressOf(row, col))}"`,
        merge?.rowspan > 1 ? `rowspan="${merge.rowspan}"` : '',
        merge?.colspan > 1 ? `colspan="${merge.colspan}"` : '',
        `style="${escapeHtml(css)}"`,
        classes ? `class="${classes}"` : '',
      ].filter(Boolean).join(' ')
      const contentStyle = `max-height:${contentHeight.toFixed(2)}px${overflowing ? `;width:${Math.max(0, availableWidth - 8).toFixed(2)}px` : ''}`
      const text = accounting ? `<span class="accounting-symbol">${escapeHtml(accounting.symbol)}</span><span class="accounting-amount">${escapeHtml(accounting.amount)}</span>` : escapeHtml(value)
      let content
      if (checkbox !== null) content = `<span class="print-checkbox" style="${contentStyle}">${checkboxSvg(checkbox)}</span>`
      else if (visual?.icon && !cell.hyperlink) content = `<span class="print-cf" style="${contentStyle}"><span class="print-cf-icon" aria-hidden="true">${visual.icon}</span><span class="print-cf-text${accounting ? ' print-accounting' : ''}">${text}</span></span>`
      else content = `<span${accounting ? ' class="print-accounting"' : ''} style="${contentStyle}">${text}</span>`
      if (sparkline) content = `<span class="print-sparkline" aria-hidden="true">${sparkline}</span>${content}`
      cells.push(`<td ${attributes}>${content}</td>`)
    }
    return `<tr style="height:${rowHeight(sheet, row).toFixed(2)}px">${cells.join('')}</tr>`
  }).join('')
  const firstRow = rows[0]
  const lastRow = rows[rows.length - 1]
  const firstCol = cols[0]
  const lastCol = cols[cols.length - 1]
  const rowRange = Number.isFinite(firstRow) ? `${firstRow + 1}:${lastRow + 1}` : ''
  const columnRange = Number.isFinite(firstCol) ? `${columnName(firstCol)}:${columnName(lastCol)}` : ''
  // The page's own rows and columns, after repeated titles: where its page breaks fall.
  const bodyRows = rows.slice(leadingTitleCount(rows, new Set(repeatedRows(sheet, options))))
  const bodyCols = cols.slice(leadingTitleCount(cols, new Set(repeatedColumns(sheet, options))))
  const bodyRange = bodyRows.length && bodyCols.length
    ? ` data-body-rows="${bodyRows[0] + 1}:${bodyRows[bodyRows.length - 1] + 1}" data-body-columns="${bodyCols[0] + 1}:${bodyCols[bodyCols.length - 1] + 1}"`
    : ''
  const style = [
    `--page-width:${metrics.width.toFixed(2)}px`,
    `--page-height:${metrics.height.toFixed(2)}px`,
    `--margin-top:${(metrics.margins.top * 96).toFixed(2)}px`,
    `--margin-right:${(metrics.margins.right * 96).toFixed(2)}px`,
    `--margin-bottom:${(metrics.margins.bottom * 96).toFixed(2)}px`,
    `--margin-left:${(metrics.margins.left * 96).toFixed(2)}px`,
    `--header-top:${(metrics.margins.header * 96).toFixed(2)}px`,
    `--footer-bottom:${(metrics.margins.footer * 96).toFixed(2)}px`,
    `page:${pageName}`,
  ].join(';')
  const headerFooter = sheet.headerFooter || {}
  const prefix = headerFooter.differentFirst && sheetPage === 1 ? 'first' : headerFooter.differentOddEven && pageNumber % 2 === 0 ? 'even' : 'odd'
  const headerScale = headerFooter.scaleWithSheet === false ? 1 : scale
  const header = renderHeaderFooter(headerFooter[`${prefix}Header`], headerContext, headerScale, warnings)
  const footer = renderHeaderFooter(headerFooter[`${prefix}Footer`], headerContext, headerScale, warnings)
  const alignment = [options.horizontalCentered ? 'justify-content:center' : '', options.verticalCentered ? 'align-items:center' : ''].filter(Boolean).join(';')
  const tableWidth = cols.reduce((sum, col) => sum + columnWidth(sheet, col), options.headings ? 38 : 0)
  const tableHtml = `<table class="sheet-table${options.gridlines ? ' show-gridlines' : ''}${options.headings ? ' show-headings' : ''}" style="width:${tableWidth.toFixed(2)}px"><colgroup>${columns.join('')}</colgroup>${head}<tbody>${body}</tbody></table>`
  // Charts: positioned over the table inside the scaled sheet so they scale with it.
  const chartOverlay = chartOverlayHtml(sheet, rows, cols, options, charts)
  const sheetHtml = chartOverlay ? `<div class="print-sheet-body">${tableHtml}${chartOverlay}</div>` : tableHtml
  return `<section class="print-page print-sheet" aria-label="Page ${pageNumber}, ${escapeHtml(sheet.name)}" data-page-number="${pageNumber}" data-sheet-page="${sheetPage}" data-sheet-id="${escapeHtml(sheet.id)}" data-sheet-name="${escapeHtml(sheet.name)}" data-row-range="${rowRange}" data-column-range="${columnRange}"${bodyRange} data-scale="${scale.toFixed(4)}" style="${style}">${header ? `<header class="page-header">${header}</header>` : ''}<div class="page-content" style="${alignment}"><div class="sheet-scale" style="--sheet-scale:${scale.toFixed(4)}">${sheetHtml}</div></div>${footer ? `<footer class="page-footer">${footer}</footer>` : ''}</section>`
}

function selectedSheets(workbook, options) {
  const sheets = Array.isArray(workbook?.sheets) ? workbook.sheets : []
  const active = sheets.find((sheet) => sheet?.id === workbook.activeSheetId) || sheets[0]
  if (!active) throw new Error('There is no worksheet to print.')
  if (options.scope === 'workbook') {
    const visible = sheets.filter((sheet) => sheet && sheet.state !== 'hidden' && sheet.state !== 'veryHidden')
    if (visible.length > MAX_PRINT_SHEETS) throw new Error(`This workbook has more than ${MAX_PRINT_SHEETS} visible sheets. Print a smaller workbook or one sheet at a time.`)
    return visible.length ? visible : [active]
  }
  return [active]
}

function createSpreadsheetPrintDocument(input) {
  if (!input || typeof input !== 'object' || !input.workbook || typeof input.workbook !== 'object') throw new Error('Invalid print request.')
  const workbook = input.workbook
  const options = safeOptions(input.options)
  const sheets = selectedSheets(workbook, options)
  const selection = normalizedBounds(input.selection)
  if (options.scope === 'selection' && !selection) throw new Error('Select a valid range before printing the selection.')
  let printedCells = 0
  const warnings = new Set()
  const chartsBySheet = new Map(sheets.map((sheet) => [sheet.id, [...printImagesFor(input, sheet), ...printChartsFor(input, sheet)]]))
  const renderable = sheets.flatMap((sheet, sheetIndex) => {
    const extents = boundsIncludingCharts(usedBounds(sheet), chartsBySheet.get(sheet.id) || [])
    const areas = options.scope === 'selection' ? [selection] : printAreas(sheet.pageSetup, extents)
    if (!areas.length) areas.push(extents)
    const resolvedOptions = sheetOptions(sheet, options, warnings)
    const metrics = pageMetrics(resolvedOptions)
    return areas.map((bounds) => {
      const visibleRows = visibleIndices(bounds.top, bounds.bottom, sheet.hiddenRows).length
      const visibleCols = visibleIndices(bounds.left, bounds.right, sheet.hiddenCols).length
      printedCells += visibleRows * visibleCols
      if (printedCells > MAX_PRINT_CELLS) throw new Error('This print job is too large to preview at once. Print a smaller selection or fewer sheets.')
      return { sheet, bounds, options: resolvedOptions, metrics, pageName: `sheet-${sheetIndex}` }
    })
  })
  const themeColors = Array.isArray(workbook.metadata?.themeColors) ? workbook.metadata.themeColors : []
  const pages = []
  let minimumScale = 1
  let oversizedDimensions = 0
  const sheetPageCounts = new Map()
  let renderedCells = 0
  for (const { sheet, bounds, options: resolvedOptions, metrics, pageName } of renderable) {
    const rows = visibleIndices(bounds.top, bounds.bottom, sheet.hiddenRows)
    const cols = visibleIndices(bounds.left, bounds.right, sheet.hiddenCols)
    const scale = scaleFor(sheet, rows, cols, resolvedOptions, metrics)
    if (resolvedOptions.scaling === 'actual' || resolvedOptions.scaling === 'custom') {
      const availableWidth = Math.max(1, metrics.contentWidth / scale - TABLE_EDGE_ALLOWANCE - (resolvedOptions.headings ? 38 : 0))
      const availableHeight = Math.max(1, metrics.contentHeight / scale - TABLE_EDGE_ALLOWANCE - (resolvedOptions.headings ? 24 : 0))
      oversizedDimensions += visibleIndices(bounds.left, bounds.right, sheet.hiddenCols)
        .filter((col) => columnWidth(sheet, col) > availableWidth).length
      oversizedDimensions += visibleIndices(bounds.top, bounds.bottom, sheet.hiddenRows)
        .filter((row) => rowHeight(sheet, row) > availableHeight).length
    }
    const sheetPages = paginateSheet(sheet, bounds, resolvedOptions, metrics)
    sheetPages.forEach((page) => {
      renderedCells += page.rows.length * page.cols.length
      if (renderedCells > MAX_PRINT_CELLS) throw new Error('The repeated print headings make this job too large. Print fewer rows or sheets.')
      minimumScale = Math.min(minimumScale, page.scale)
      const sheetPage = (sheetPageCounts.get(sheet.id) || 0) + 1
      sheetPageCounts.set(sheet.id, sheetPage)
      pages.push({ sheet, bounds, ...page, options: resolvedOptions, metrics, pageName, sheetPage })
    })
  }
  const metrics = renderable[0].metrics
  const primaryOptions = renderable[0].options
  const mixedPaperSizes = renderable.some((entry) => entry.metrics.width !== metrics.width || entry.metrics.height !== metrics.height)
  if (mixedPaperSizes) warnings.add('These sheets use different paper sizes or orientations. PDF preserves each page; print sheets separately if your printer uses one paper setting for the whole job.')
  const title = String(input.name || workbook.name || 'Spreadsheet').slice(0, 180)
  const filename = title.split(/[\\/]/).pop()
  const sourcePath = String(input.path || workbook.path || '')
  const directory = sourcePath.slice(0, Math.max(sourcePath.lastIndexOf('/'), sourcePath.lastIndexOf('\\')) + 1)
  const requestedDate = new Date(input.printedAt || Date.now())
  const timestamp = Number.isNaN(requestedDate.getTime()) ? new Date() : requestedDate
  // A page range keeps each page's number and the total (Page 3 of 10) from the whole printout.
  let printed = pages.map((page, index) => ({ page, number: index + 1 }))
  if (options.pageRange) {
    const { from, to } = options.pageRange
    if (to !== null && to < from) throw new Error('The last page to print comes before the first. Check the page range.')
    if (from > pages.length) throw new Error(`This printout has ${pages.length} ${pages.length === 1 ? 'page' : 'pages'}. Choose pages from 1 to ${pages.length}.`)
    printed = printed.slice(from - 1, to === null ? pages.length : Math.min(pages.length, to))
  }
  const icons = safeIconTable(input.visuals && typeof input.visuals === 'object' ? input.visuals.icons : null)
  const budget = { svg: 0 }
  const visualsBySheet = new Map(sheets.map((sheet) => [sheet.id, sheetVisualsFor(input, sheet, icons, budget)]))
  const sections = printed.map(({ page, number }) => renderPage(
    page.sheet,
    page.bounds,
    page.rows,
    page.cols,
    input.displayValues?.[page.sheet.id] || {},
    page.options,
    themeColors,
    page.metrics,
    number,
    page.sheetPage,
    page.scale,
    page.pageName,
    {
      page: page.sheet.pageSetup?.useFirstPageNumber === true && Number(page.sheet.pageSetup.firstPageNumber) > 0 ? Number(page.sheet.pageSetup.firstPageNumber) + page.sheetPage - 1 : number,
      pages: pages.length,
      filename,
      directory,
      sheetName: String(page.sheet.name || ''),
      date: timestamp.toLocaleDateString(),
      time: timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
    },
    warnings,
    input.displayParts?.[page.sheet.id] || {},
    chartsBySheet.get(page.sheet.id) || [],
    visualsBySheet.get(page.sheet.id) || null,
  )).join('')
  const paper = metrics.paper
  const pageStyles = Array.from(new Map(renderable.map((entry) => [entry.pageName, `@page ${entry.pageName} { size: ${entry.metrics.paper.css} ${entry.options.orientation}; margin: 0; }`])).values()).join('\n')
  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(title)}</title>
<style>
@page { size: ${paper.css} ${primaryOptions.orientation}; margin: 0; }
${pageStyles}
* { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
html, body { margin: 0; min-height: 100%; padding: 0; color: #171a18; font-family: Aptos, Calibri, Arial, sans-serif; font-size: 10pt; }
:root { --preview-zoom:1; }
.print-page { position: relative; width: var(--page-width); height: var(--page-height); padding: var(--margin-top) var(--margin-right) var(--margin-bottom) var(--margin-left); overflow: hidden; background: #fff; break-after: page; page-break-after: always; }
.print-page:last-child { break-after: auto; page-break-after: auto; }
.page-content { display: flex; align-items: flex-start; width: 100%; height: 100%; overflow: hidden; }
.sheet-scale { flex: none; zoom: var(--sheet-scale); transform-origin: top left; }
.print-sheet-body { position: relative; width: max-content; }
.print-charts { position: absolute; overflow: hidden; pointer-events: none; }
.print-chart { position: absolute; }
.print-chart > svg { display: block; width: 100%; height: 100%; }
.print-chart > .print-image { display: block; width: 100%; height: 100%; }
.page-header, .page-footer { position: absolute; left: var(--margin-left); right: var(--margin-right); display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); color: #000; font: 11pt Arial, sans-serif; line-height: 1.15; white-space: pre; }
.page-header { top: var(--header-top); }
.page-footer { bottom: var(--footer-bottom); }
.header-footer-left { text-align: left; }
.header-footer-center { text-align: center; }
.header-footer-right { text-align: right; }
.sheet-table { border-collapse: collapse; table-layout: fixed; width: max-content; }
.sheet-table th, .sheet-table td { min-width: 2px; height: inherit; padding: 1px 4px; overflow: hidden; vertical-align: bottom; line-height: 1.05; }
.sheet-table td > span, .row-heading > span { display: block; overflow: hidden; text-overflow: clip; }
.sheet-table td.print-overflow { overflow: visible; }
.print-overflow > span { position: relative; z-index: 1; }
.sheet-table td > .print-accounting { display: flex; justify-content: space-between; gap: 4px; }
.accounting-symbol { flex: none; }
.accounting-amount { min-width: 0; overflow: hidden; text-align: right; }
.sheet-table.show-gridlines td { border: 1px solid #c8ceca; }
.sheet-table td.has-sparkline { position: relative; }
.sheet-table td > .print-sparkline { position: absolute; inset: 2px 3px; display: block; overflow: hidden; }
.print-sparkline > svg { display: block; width: 100%; height: 100%; }
.sheet-table td > .print-cf { display: flex; align-items: center; gap: 3px; }
.print-cf-icon { flex: none; display: block; width: 13px; height: 13px; }
.print-cf-icon > svg { display: block; width: 13px; height: 13px; }
.print-cf-text { display: block; flex: 1 1 auto; min-width: 0; overflow: hidden; }
.print-cf-text.print-accounting { display: flex; justify-content: space-between; gap: 4px; }
.sheet-table td > .print-checkbox { display: flex; align-items: center; justify-content: center; }
.print-checkbox > svg { display: block; flex: none; }
.column-heading, .row-heading, .corner-heading { border: 1px solid #adb5b0; color: #47514b; background: #eef1ef; font-weight: 600; text-align: center; vertical-align: middle; }
.column-heading { height: 24px; }
.row-heading { width: 38px; padding: 1px 2px; }
thead { display: table-header-group; }
tr, td, th { break-inside: avoid; page-break-inside: avoid; }
@media screen {
  html { background: #dfe3e0; }
  body { width: max-content; min-width: 100%; display: flex; flex-direction: column; align-items: center; gap: 24px; padding: 24px; background: #dfe3e0; zoom: var(--preview-zoom); }
  .print-page { flex: 0 0 auto; box-shadow: 0 2px 14px rgba(31, 38, 34, 0.22); }
}
@media print {
  html, body { background: #fff; }
  body { display: block; }
  .print-page { margin: 0; box-shadow: none; }
}
</style></head><body data-print-scope="${options.scope}" data-orientation="${primaryOptions.orientation}" data-scaling="${primaryOptions.scaling}" data-paper-size="${primaryOptions.paperSize}" data-gridlines="${primaryOptions.gridlines}" data-headings="${primaryOptions.headings}" data-total-pages="${pages.length}" data-first-page="${printed.length ? printed[0].number : 1}">${sections}</body></html>`
  return {
    html,
    title,
    options: primaryOptions,
    warnings: [...warnings],
    mixedPaperSizes,
    sheetCount: sheets.length,
    printedCells,
    // Pages in this document (a page range prints part of the printout) and in the whole printout.
    pageCount: printed.length,
    totalPages: pages.length,
    pageBreaks: Math.max(0, pages.length - sheets.length),
    minimumScale,
    oversizedDimensions,
    paper: {
      label: paper.label,
      widthInches: metrics.widthInches,
      heightInches: metrics.heightInches,
    },
    // Native job settings for webContents.print (see default-printer.cjs directPrintOptions).
    printJob: printJobOf(input.options && typeof input.options === 'object' ? input.options.printer : null),
  }
}

module.exports = {
  MAX_PRINT_CELLS,
  MAX_PRINT_COPIES,
  MARGIN_PRESETS,
  PAPER_SIZES,
  createSpreadsheetPrintDocument,
  parseAddress,
  parseRange,
  usedBounds,
}
