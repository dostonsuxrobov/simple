const MAX_PRINT_CELLS = 400_000
const MAX_PRINT_SHEETS = 100
// Collapsed table borders add their outside stroke to the sum of column/row
// dimensions. Reserve the maximum supported 3px stroke plus rounding room so
// fit-to-page never places that stroke beyond the clipped content rectangle.
const TABLE_EDGE_ALLOWANCE = 4

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

function printAreas(value) {
  if (!value) return []
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
  const ranges = pieces.map(parseRange)
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

function safeOptions(value) {
  const input = value && typeof value === 'object' ? value : {}
  return {
    scope: ['active-sheet', 'selection', 'workbook'].includes(input.scope) ? input.scope : 'active-sheet',
    orientation: input.orientation === 'landscape' ? 'landscape' : 'portrait',
    scaling: ['actual', 'fit-width', 'fit-sheet'].includes(input.scaling) ? input.scaling : 'fit-width',
    paperSize: Object.hasOwn(PAPER_SIZES, input.paperSize) ? input.paperSize : 'letter',
    margins: Object.hasOwn(MARGIN_PRESETS, input.margins) ? input.margins : 'normal',
    gridlines: input.gridlines !== false,
    headings: input.headings === true,
    useSavedLayout: input.useSavedLayout === true,
  }
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
  // The dormant fitToWidth/fitToHeight fields are often 1 even when Excel's
  // active mode is percentage scale. Only fitToPage enables those fields.
  if (setup.fitToPage === true) {
    result.fitWidth = Math.max(0, Math.min(100, Math.trunc(Number(setup.fitToWidth) || 0)))
    result.fitHeight = Math.max(0, Math.min(100, Math.trunc(Number(setup.fitToHeight) || 0)))
    result.scaling = result.fitWidth === 1 && result.fitHeight === 1 ? 'fit-sheet' : 'actual'
  } else if (Number.isFinite(Number(setup.scale)) && Number(setup.scale) >= 10 && Number(setup.scale) <= 400) {
    result.savedScale = Number(setup.scale) / 100
    result.scaling = 'actual'
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
  const margins = { ...MARGIN_PRESETS[options.margins], header: 0.2, footer: 0.2 }
  for (const side of ['top', 'right', 'bottom', 'left', 'header', 'footer']) {
    const value = Number(options.savedMargins?.[side])
    if (Number.isFinite(value) && value >= 0 && value <= 10) margins[side] = value
  }
  const widthInches = options.orientation === 'landscape' ? paper.height : paper.width
  const heightInches = options.orientation === 'landscape' ? paper.width : paper.height
  if (margins.left + margins.right >= widthInches || margins.top + margins.bottom >= heightInches) {
    throw new Error('The saved page margins leave no room for cells. Turn off Use saved page layout and choose smaller margins.')
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

function scaleFor(sheet, rows, cols, options, metrics) {
  if (options.savedScale) return options.savedScale
  if (options.fitWidth || options.fitHeight) {
    const width = cols.reduce((sum, col) => sum + columnWidth(sheet, col), options.headings ? 38 : 0)
    const height = rows.reduce((sum, row) => sum + rowHeight(sheet, row), options.headings ? 24 : 0)
    let scale = Math.min(1,
      options.fitWidth ? metrics.contentWidth * options.fitWidth / Math.max(1, width + TABLE_EDGE_ALLOWANCE) : Infinity,
      options.fitHeight ? metrics.contentHeight * options.fitHeight / Math.max(1, height + TABLE_EDGE_ALLOWANCE) : Infinity)
    // Account for whole rows/columns at page boundaries, rather than assuming a
    // cell may split over two pages when fitting to a saved page count.
    for (let attempt = 0; attempt < 100; attempt += 1) {
      const across = chunkIndices(cols, (col) => columnWidth(sheet, col), metrics.contentWidth / scale - TABLE_EDGE_ALLOWANCE - (options.headings ? 38 : 0)).length
      const down = chunkRows(sheet, rows, metrics.contentHeight / scale - TABLE_EDGE_ALLOWANCE - (options.headings ? 24 : 0), repeatedRows(sheet, options)).length
      if ((!options.fitWidth || across <= options.fitWidth) && (!options.fitHeight || down <= options.fitHeight)) break
      scale *= 0.99
    }
    return Math.max(0.01, scale)
  }
  if (options.scaling === 'actual') return 1
  const width = cols.reduce((sum, col) => sum + columnWidth(sheet, col), options.headings ? 38 : 0)
  const height = rows.reduce((sum, row) => sum + rowHeight(sheet, row), options.headings ? 24 : 0)
  const widthScale = metrics.contentWidth / Math.max(1, width + TABLE_EDGE_ALLOWANCE)
  const heightScale = metrics.contentHeight / Math.max(1, height + TABLE_EDGE_ALLOWANCE)
  const value = options.scaling === 'fit-sheet' ? Math.min(widthScale, heightScale) : widthScale
  return Math.max(0.01, Math.min(1, value))
}

function chunkIndices(indices, sizeFor, capacity) {
  if (!indices.length) return [[]]
  const chunks = []
  let current = []
  let consumed = 0
  for (const index of indices) {
    const size = Math.max(1, sizeFor(index))
    if (current.length && consumed + size > capacity) {
      chunks.push(current)
      current = []
      consumed = 0
    }
    current.push(index)
    consumed += size
  }
  if (current.length) chunks.push(current)
  return chunks
}

function repeatedRows(sheet, options) {
  if (options.scope === 'selection') return []
  const match = /^\$?(\d+):\$?(\d+)$/.exec(String(sheet.pageSetup?.printTitlesRow || ''))
  if (!match) return []
  const first = Number(match[1]), last = Number(match[2])
  if (first < 1 || last < first || last > 1048576 || last - first > 4095) throw new Error('The saved rows to repeat are too large. Choose a smaller print title range.')
  return visibleIndices(first - 1, last - 1, sheet.hiddenRows)
}

function chunkRows(sheet, rows, capacity, titles) {
  if (!titles.length) return chunkIndices(rows, row => rowHeight(sheet, row), capacity)
  const titleHeight = titles.reduce((sum,row) => sum + rowHeight(sheet,row), 0)
  if (titleHeight >= capacity) throw new Error('The repeated heading rows fill the page. Use a smaller scale or fewer heading rows.')
  const chunks = [], present = new Set(rows)
  let current = titles.filter(row => !present.has(row))
  let consumed = current.reduce((sum,row) => sum + rowHeight(sheet,row),0)
  for (const row of rows) {
    if (current.includes(row)) continue
    const size = rowHeight(sheet,row)
    if (current.length && consumed + size > capacity) {
      chunks.push(current)
      current = [...titles]
      consumed = titleHeight
    }
    if (!current.includes(row)) { current.push(row); consumed += size }
  }
  if (current.length) chunks.push(current)
  return chunks
}

function paginateSheet(sheet, bounds, options, metrics) {
  const rows = visibleIndices(bounds.top, bounds.bottom, sheet.hiddenRows)
  const cols = visibleIndices(bounds.left, bounds.right, sheet.hiddenCols)
  const scale = scaleFor(sheet, rows, cols, options, metrics)
  if (options.scaling === 'fit-sheet') return [{ rows, cols, scale }]

  const unscaledWidth = metrics.contentWidth / scale - TABLE_EDGE_ALLOWANCE - (options.headings ? 38 : 0)
  const unscaledHeight = metrics.contentHeight / scale - TABLE_EDGE_ALLOWANCE - (options.headings ? 24 : 0)
  const columnChunks = options.scaling === 'fit-width'
    ? [cols]
    : chunkIndices(cols, (col) => columnWidth(sheet, col), Math.max(1, unscaledWidth))
  const rowChunks = chunkRows(sheet, rows, Math.max(1, unscaledHeight), repeatedRows(sheet, options))
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

function chartOverlayHtml(sheet, rows, cols, options, charts) {
  if (!charts.length || !rows.length || !cols.length) return ''
  const headX = options.headings ? 38 : 0
  const headY = options.headings ? 24 : 0
  let maxCol = cols[cols.length - 1], maxRow = rows.reduce((max, row) => Math.max(max, row), 0)
  for (const chart of charts) { maxCol = Math.max(maxCol, chart.to.col); maxRow = Math.max(maxRow, chart.to.row) }
  const xs = axisPositions(maxCol + 2, (col) => columnWidth(sheet, col), sheet.hiddenCols)
  const ys = axisPositions(maxRow + 2, (row) => rowHeight(sheet, row), sheet.hiddenRows)
  // Repeated title rows prepended to a later page sit above the page's body rows.
  const titles = new Set(repeatedRows(sheet, options))
  let prefix = 0
  while (prefix < rows.length && titles.has(rows[prefix])) prefix += 1
  if (prefix >= rows.length || (prefix > 0 && rows[prefix] === rows[prefix - 1] + 1)) prefix = 0
  const bodyTop = headY + rows.slice(0, prefix).reduce((sum, row) => sum + rowHeight(sheet, row), 0)
  const width = cols.reduce((sum, col) => sum + columnWidth(sheet, col), 0)
  const height = rows.slice(prefix).reduce((sum, row) => sum + rowHeight(sheet, row), 0)
  const originX = xs.at(cols[0])
  const originY = ys.at(rows[prefix])
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
  return `<div class="print-charts" aria-hidden="false" style="left:${headX}px;top:${bodyTop.toFixed(2)}px;width:${width.toFixed(2)}px;height:${height.toFixed(2)}px">${items.join('')}</div>`
}

function renderPage(sheet, bounds, rows, cols, displayValues, options, themeColors, metrics, pageNumber, sheetPage, scale, pageName, headerContext, warnings, displayParts, charts = []) {
  const layout = mergeLayout(sheet, { ...bounds, top: rows.reduce((min,row) => Math.min(min,row), bounds.top), bottom: rows.reduce((max,row) => Math.max(max,row), bounds.bottom) }, rows, cols)
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
      const value = displayValues?.[address] ?? cell.display ?? cell.result ?? cell.value ?? ''
      const parts = printDisplayParts(cell, value, displayParts?.[address])
      const height = rows.slice(rowIndex, rowIndex + (merge?.rowspan || 1)).reduce((sum, item) => sum + rowHeight(sheet, item), 0)
      const contentHeight = cellContentHeight(cell, height, options.gridlines)
      const availableWidth = value !== '' ? overflowWidth(sheet, row, colIndex, cols, cell, merge, layout, displayValues, parts) : columnWidth(sheet, col)
      const overflowing = !merge && availableWidth > columnWidth(sheet, col)
      const attributes = [
        `data-address="${escapeHtml(addressOf(row, col))}"`,
        merge?.rowspan > 1 ? `rowspan="${merge.rowspan}"` : '',
        merge?.colspan > 1 ? `colspan="${merge.colspan}"` : '',
        `style="${escapeHtml(cellStyle(cell, themeColors, parts.type))}"`,
        overflowing ? 'class="print-overflow"' : '',
      ].filter(Boolean).join(' ')
      const contentStyle = `max-height:${contentHeight.toFixed(2)}px${overflowing ? `;width:${Math.max(0, availableWidth - 8).toFixed(2)}px` : ''}`
      const content = parts.accounting ? `<span class="accounting-symbol">${escapeHtml(parts.accounting.symbol)}</span><span class="accounting-amount">${escapeHtml(parts.accounting.amount)}</span>` : escapeHtml(value)
      cells.push(`<td ${attributes}><span${parts.accounting ? ' class="print-accounting"' : ''} style="${contentStyle}">${content}</span></td>`)
    }
    return `<tr style="height:${rowHeight(sheet, row).toFixed(2)}px">${cells.join('')}</tr>`
  }).join('')
  const firstRow = rows[0]
  const lastRow = rows[rows.length - 1]
  const firstCol = cols[0]
  const lastCol = cols[cols.length - 1]
  const rowRange = Number.isFinite(firstRow) ? `${firstRow + 1}:${lastRow + 1}` : ''
  const columnRange = Number.isFinite(firstCol) ? `${columnName(firstCol)}:${columnName(lastCol)}` : ''
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
  return `<section class="print-page print-sheet" aria-label="Page ${pageNumber}, ${escapeHtml(sheet.name)}" data-page-number="${pageNumber}" data-sheet-page="${sheetPage}" data-sheet-id="${escapeHtml(sheet.id)}" data-sheet-name="${escapeHtml(sheet.name)}" data-row-range="${rowRange}" data-column-range="${columnRange}" data-scale="${scale.toFixed(4)}" style="${style}">${header ? `<header class="page-header">${header}</header>` : ''}<div class="page-content" style="${alignment}"><div class="sheet-scale" style="--sheet-scale:${scale.toFixed(4)}">${sheetHtml}</div></div>${footer ? `<footer class="page-footer">${footer}</footer>` : ''}</section>`
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
  for (const sheet of sheets) if (sheet.pageSetup?.printTitlesColumn && options.scope !== 'selection') warnings.add(`${sheet.name}: repeated heading columns are not yet included in print previews.`)
  const chartsBySheet = new Map(sheets.map((sheet) => [sheet.id, [...printImagesFor(input, sheet), ...printChartsFor(input, sheet)]]))
  const renderable = sheets.flatMap((sheet, sheetIndex) => {
    const areas = options.scope === 'selection' ? [selection] : printAreas(sheet.pageSetup?.printArea)
    if (!areas.length) areas.push(boundsIncludingCharts(usedBounds(sheet), chartsBySheet.get(sheet.id) || []))
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
    if (resolvedOptions.scaling === 'actual') {
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
  const sections = pages.map((page, index) => renderPage(
    page.sheet,
    page.bounds,
    page.rows,
    page.cols,
    input.displayValues?.[page.sheet.id] || {},
    page.options,
    themeColors,
    page.metrics,
    index + 1,
    page.sheetPage,
    page.scale,
    page.pageName,
    {
      page: page.sheet.pageSetup?.useFirstPageNumber === true && Number(page.sheet.pageSetup.firstPageNumber) > 0 ? Number(page.sheet.pageSetup.firstPageNumber) + page.sheetPage - 1 : index + 1,
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
  )).join('')
  const paper = metrics.paper
  const pageStyles = Array.from(new Map(renderable.map((entry) => [entry.pageName, `@page ${entry.pageName} { size: ${entry.metrics.paper.css} ${entry.options.orientation}; margin: 0; }`])).values()).join('\n')
  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(title)}</title>
<style>
@page { size: ${paper.css} ${primaryOptions.orientation}; margin: 0; }
${pageStyles}
* { box-sizing: border-box; }
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
</style></head><body data-print-scope="${options.scope}" data-orientation="${primaryOptions.orientation}" data-scaling="${primaryOptions.scaling}" data-paper-size="${primaryOptions.paperSize}" data-gridlines="${primaryOptions.gridlines}" data-headings="${primaryOptions.headings}">${sections}</body></html>`
  return {
    html,
    title,
    options: primaryOptions,
    warnings: [...warnings],
    mixedPaperSizes,
    sheetCount: sheets.length,
    printedCells,
    pageCount: pages.length,
    pageBreaks: Math.max(0, pages.length - sheets.length),
    minimumScale,
    oversizedDimensions,
    paper: {
      label: paper.label,
      widthInches: metrics.widthInches,
      heightInches: metrics.heightInches,
    },
  }
}

module.exports = {
  MAX_PRINT_CELLS,
  MARGIN_PRESETS,
  PAPER_SIZES,
  createSpreadsheetPrintDocument,
  parseAddress,
  parseRange,
  usedBounds,
}
