'use strict'

const XLSX = require('xlsx')

const COMPOUND_FILE_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])

const RECORD = Object.freeze({
  BOF: 0x0809,
  BOUND_SHEET: 0x0085,
  FONT: 0x0031,
  XF: 0x00e0,
  PALETTE: 0x0092,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOL_ERR: 0x0205,
  FORMULA: 0x0006,
  FORMULA_BIFF3: 0x0206,
  FORMULA_BIFF4: 0x0406,
  BLANK: 0x0201,
  LABEL_SST: 0x00fd,
  RK: 0x027e,
  RSTRING: 0x00d6,
  MUL_RK: 0x00bd,
  MUL_BLANK: 0x00be,
  DEFAULT_ROW_HEIGHT: 0x0225,
  DEFAULT_COL_WIDTH: 0x0055,
  ROW: 0x0208,
  HEADER: 0x0014,
  FOOTER: 0x0015,
  SETUP: 0x00a1,
  WS_BOOL: 0x0081,
  PRINT_GRID: 0x002b,
  PRINT_ROW_COL: 0x002a,
  HCENTER: 0x0083,
  VCENTER: 0x0084,
  STYLE: 0x0293,
  EOF: 0x000a,
})

const FILL_PATTERNS = [
  null,
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
]

const BORDER_STYLES = [
  null,
  'thin',
  'medium',
  'dashed',
  'dotted',
  'thick',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
]

const HORIZONTAL_ALIGNMENTS = [
  undefined,
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
]

const VERTICAL_ALIGNMENTS = ['top', 'middle', 'bottom', 'justify', 'distributed']

function isCompoundFile(buffer) {
  return Buffer.isBuffer(buffer) && buffer.length >= 8 && buffer.subarray(0, 8).equals(COMPOUND_FILE_SIGNATURE)
}

function workbookStream(buffer) {
  if (!isCompoundFile(buffer)) return null
  const cfb = XLSX.CFB.read(buffer, { type: 'buffer' })
  const entry = (Array.isArray(cfb.FileIndex) ? cfb.FileIndex : []).find((item) => (
    item && item.type === 2 && /^(?:Workbook|Book)$/i.test(String(item.name || '')) && item.content
  ))
  if (!entry) return null
  const content = Buffer.from(entry.content)
  return Number.isFinite(entry.size) && entry.size >= 0 ? content.subarray(0, Math.min(content.length, entry.size)) : content
}

function records(stream, start = 0, end = stream.length) {
  const result = []
  let offset = Math.max(0, start)
  const limit = Math.min(stream.length, Math.max(offset, end))
  while (offset + 4 <= limit) {
    const type = stream.readUInt16LE(offset)
    const length = stream.readUInt16LE(offset + 2)
    const payloadStart = offset + 4
    const payloadEnd = payloadStart + length
    if (payloadEnd > limit) break
    result.push({ type, offset, data: stream.subarray(payloadStart, payloadEnd) })
    offset = payloadEnd
  }
  return result
}

function parseBoundSheet(record) {
  if (!record || record.data.length < 4) return null
  return { position: record.data.readUInt32LE(0), recordOffset: record.offset }
}

function parseFont(data) {
  if (!data || data.length < 16) return undefined
  const height = data.readUInt16LE(0)
  const flags = data.readUInt16LE(2)
  const colorIndex = data.readUInt16LE(4)
  const weight = data.readUInt16LE(6)
  const escapement = data.readUInt16LE(8)
  const underline = data[10]
  const family = data[11]
  const charset = data[12]
  const characterCount = data[14]
  const unicode = Boolean(data[15] & 0x01)
  const byteCount = characterCount * (unicode ? 2 : 1)
  const end = Math.min(data.length, 16 + byteCount)
  const name = unicode
    ? data.toString('utf16le', 16, end)
    : data.toString('latin1', 16, end)

  const font = {}
  if (name) font.name = name
  if (height > 0) font.size = height / 20
  if (weight >= 700) font.bold = true
  if (flags & 0x0002) font.italic = true
  if (flags & 0x0008) font.strike = true
  if (flags & 0x0010) font.outline = true
  if (flags & 0x0020) font.shadow = true
  if (flags & 0x0040) font.condense = true
  if (flags & 0x0080) font.extend = true
  if (underline === 1) font.underline = 'single'
  else if (underline === 2) font.underline = 'double'
  else if (underline === 0x21) font.underline = 'singleAccounting'
  else if (underline === 0x22) font.underline = 'doubleAccounting'
  if (escapement === 1) font.vertAlign = 'superscript'
  else if (escapement === 2) font.vertAlign = 'subscript'
  if (family) font.family = family
  if (charset) font.charset = charset
  if (colorIndex !== 0x7fff) font.colorIndex = colorIndex
  return font
}

function parseXf(data) {
  if (!data || data.length < 20) return undefined
  const flags = data.readUInt16LE(4)
  const alignment = data.readUInt32LE(6)
  const border1 = data.readUInt32LE(10)
  const border2 = data.readUInt32LE(14)
  const fill = data.readUInt16LE(18)
  return {
    fontIndex: data.readUInt16LE(0),
    numFmtId: data.readUInt16LE(2),
    locked: Boolean(flags & 0x0001),
    formulaHidden: Boolean(flags & 0x0002),
    styleXf: Boolean(flags & 0x0004),
    parentXf: (flags >>> 4) & 0x0fff,
    horizontal: alignment & 0x07,
    wrapText: Boolean((alignment >>> 3) & 0x01),
    vertical: (alignment >>> 4) & 0x07,
    justifyLastLine: Boolean((alignment >>> 7) & 0x01),
    textRotation: (alignment >>> 8) & 0xff,
    indent: (alignment >>> 16) & 0x0f,
    shrinkToFit: Boolean((alignment >>> 20) & 0x01),
    readingOrder: (alignment >>> 22) & 0x03,
    leftStyle: border1 & 0x0f,
    rightStyle: (border1 >>> 4) & 0x0f,
    topStyle: (border1 >>> 8) & 0x0f,
    bottomStyle: (border1 >>> 12) & 0x0f,
    leftColor: (border1 >>> 16) & 0x7f,
    rightColor: (border1 >>> 23) & 0x7f,
    diagonalDown: Boolean((border1 >>> 30) & 0x01),
    diagonalUp: Boolean((border1 >>> 31) & 0x01),
    topColor: border2 & 0x7f,
    bottomColor: (border2 >>> 7) & 0x7f,
    diagonalColor: (border2 >>> 14) & 0x7f,
    diagonalStyle: (border2 >>> 21) & 0x0f,
    patternIndex: (border2 >>> 26) & 0x3f,
    foregroundColor: fill & 0x7f,
    backgroundColor: (fill >>> 7) & 0x7f,
  }
}

function parsePalette(data) {
  if (!data || data.length < 2) return []
  const count = Math.min(data.readUInt16LE(0), Math.floor((data.length - 2) / 4), 56)
  const palette = []
  for (let index = 0; index < count; index += 1) {
    const offset = 2 + index * 4
    palette[index] = [data[offset], data[offset + 1], data[offset + 2]]
      .map((channel) => channel.toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()
  }
  return palette
}

function colorForIndex(index, palette, automaticColor) {
  if (index === 0x7fff || index === 0x40 || index === 0x41) {
    return automaticColor ? { argb: `FF${automaticColor}` } : { auto: true }
  }
  if (index >= 8 && index <= 63 && palette[index - 8]) return { argb: `FF${palette[index - 8]}` }
  if (index >= 0 && index <= 63) return { indexed: index }
  return automaticColor ? { argb: `FF${automaticColor}` } : { auto: true }
}

function fontForXf(xf, fonts, palette) {
  const source = fonts[xf.fontIndex]
  if (!source) return undefined
  const font = { ...source }
  const colorIndex = font.colorIndex
  delete font.colorIndex
  if (Number.isInteger(colorIndex)) font.color = colorForIndex(colorIndex, palette)
  return Object.keys(font).length ? font : undefined
}

function borderSide(styleIndex, colorIndex, palette) {
  const style = BORDER_STYLES[styleIndex]
  if (!style) return undefined
  return { style, color: colorForIndex(colorIndex, palette, '000000') }
}

function borderForXf(xf, palette) {
  const border = {}
  const left = borderSide(xf.leftStyle, xf.leftColor, palette)
  const right = borderSide(xf.rightStyle, xf.rightColor, palette)
  const top = borderSide(xf.topStyle, xf.topColor, palette)
  const bottom = borderSide(xf.bottomStyle, xf.bottomColor, palette)
  const diagonal = borderSide(xf.diagonalStyle, xf.diagonalColor, palette)
  if (left) border.left = left
  if (right) border.right = right
  if (top) border.top = top
  if (bottom) border.bottom = bottom
  if (diagonal) border.diagonal = diagonal
  if (xf.diagonalUp) border.diagonalUp = true
  if (xf.diagonalDown) border.diagonalDown = true
  return Object.keys(border).length ? border : undefined
}

function alignmentForXf(xf) {
  const alignment = {}
  const horizontal = HORIZONTAL_ALIGNMENTS[xf.horizontal]
  const vertical = VERTICAL_ALIGNMENTS[xf.vertical]
  if (horizontal) alignment.horizontal = horizontal
  if (vertical) alignment.vertical = vertical
  if (xf.wrapText) alignment.wrapText = true
  if (xf.justifyLastLine) alignment.justifyLastLine = true
  if (xf.textRotation === 255) alignment.textRotation = 'vertical'
  else if (xf.textRotation >= 1 && xf.textRotation <= 90) alignment.textRotation = xf.textRotation
  else if (xf.textRotation >= 91 && xf.textRotation <= 180) alignment.textRotation = 90 - xf.textRotation
  if (xf.indent) alignment.indent = xf.indent
  if (xf.shrinkToFit) alignment.shrinkToFit = true
  if (xf.readingOrder === 1) alignment.readingOrder = 'ltr'
  else if (xf.readingOrder === 2) alignment.readingOrder = 'rtl'
  return Object.keys(alignment).length ? alignment : undefined
}

function fillForXf(xf, palette) {
  const pattern = FILL_PATTERNS[xf.patternIndex]
  if (!pattern) return undefined
  return {
    type: 'pattern',
    pattern,
    fgColor: colorForIndex(xf.foregroundColor, palette, '000000'),
    bgColor: colorForIndex(xf.backgroundColor, palette, 'FFFFFF'),
  }
}

function clone(value) {
  if (!value || typeof value !== 'object') return value
  return JSON.parse(JSON.stringify(value))
}

function styleDescriptor(xf, fonts, palette, numberFormats) {
  if (!xf) return undefined
  const style = {}
  const font = fontForXf(xf, fonts, palette)
  const fill = fillForXf(xf, palette)
  const border = borderForXf(xf, palette)
  const alignment = alignmentForXf(xf)
  if (font) style.font = font
  if (fill) style.fill = fill
  if (border) style.border = border
  if (alignment) style.alignment = alignment
  const numFmt = numberFormats[xf.numFmtId]
  return {
    ...(Object.keys(style).length ? { style } : {}),
    ...(typeof numFmt === 'string' && numFmt && numFmt !== 'General' ? { numFmt } : {}),
  }
}

function addCellStyle(cells, row, column, xfIndex, descriptors) {
  if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || row > 65_535 || column < 0 || column > 255) return
  const descriptor = descriptors[xfIndex]
  if (!descriptor || (!descriptor.style && !descriptor.numFmt)) return
  cells[XLSX.utils.encode_cell({ r: row, c: column })] = clone(descriptor)
}

function parseSheetStyles(stream, start, end, descriptors) {
  const cells = {}
  for (const record of records(stream, start, end)) {
    const data = record.data
    switch (record.type) {
      case RECORD.NUMBER:
      case RECORD.LABEL:
      case RECORD.BOOL_ERR:
      case RECORD.FORMULA:
      case RECORD.FORMULA_BIFF3:
      case RECORD.FORMULA_BIFF4:
      case RECORD.BLANK:
      case RECORD.LABEL_SST:
      case RECORD.RK:
      case RECORD.RSTRING:
        if (data.length >= 6) addCellStyle(cells, data.readUInt16LE(0), data.readUInt16LE(2), data.readUInt16LE(4), descriptors)
        break
      case RECORD.MUL_RK: {
        if (data.length < 10) break
        const row = data.readUInt16LE(0)
        const firstColumn = data.readUInt16LE(2)
        const lastColumn = data.readUInt16LE(data.length - 2)
        const count = Math.min(lastColumn - firstColumn + 1, Math.floor((data.length - 6) / 6))
        for (let index = 0; index < count; index += 1) {
          addCellStyle(cells, row, firstColumn + index, data.readUInt16LE(4 + index * 6), descriptors)
        }
        break
      }
      case RECORD.MUL_BLANK: {
        if (data.length < 8) break
        const row = data.readUInt16LE(0)
        const firstColumn = data.readUInt16LE(2)
        const lastColumn = data.readUInt16LE(data.length - 2)
        const count = Math.min(lastColumn - firstColumn + 1, Math.floor((data.length - 6) / 2))
        for (let index = 0; index < count; index += 1) {
          addCellStyle(cells, row, firstColumn + index, data.readUInt16LE(4 + index * 2), descriptors)
        }
        break
      }
      case RECORD.EOF:
        return cells
      default:
        break
    }
  }
  return cells
}

function extractLegacyBiffStyles(buffer, numberFormats = {}) {
  const stream = workbookStream(buffer)
  if (!stream) return null
  const allRecords = records(stream)
  const globalsBof = allRecords.find((record) => record.type === RECORD.BOF)
  if (
    !globalsBof ||
    globalsBof.data.length < 4 ||
    globalsBof.data.readUInt16LE(0) !== 0x0600 ||
    globalsBof.data.readUInt16LE(2) !== 0x0005
  ) return null
  const boundSheets = allRecords
    .filter((record) => record.type === RECORD.BOUND_SHEET)
    .map(parseBoundSheet)
    .filter((item) => item && item.position >= 0 && item.position < stream.length)
  if (!boundSheets.length) return null

  const globalEnd = Math.min(...boundSheets.map((item) => item.position))
  const fonts = []
  const xfs = []
  let palette = []
  let fontRecordIndex = 0
  let normalXfIndex = 0
  for (const record of allRecords) {
    if (record.offset >= globalEnd) break
    if (record.type === RECORD.FONT) {
      const fontIndex = fontRecordIndex < 4 ? fontRecordIndex : fontRecordIndex + 1
      fonts[fontIndex] = parseFont(record.data)
      fontRecordIndex += 1
    } else if (record.type === RECORD.XF) {
      xfs.push(parseXf(record.data))
    } else if (record.type === RECORD.PALETTE) {
      palette = parsePalette(record.data)
    } else if (record.type === RECORD.STYLE && record.data.length >= 4) {
      const flags = record.data.readUInt16LE(0)
      if ((flags & 0x8000) && record.data[2] === 0) normalXfIndex = flags & 0x0fff
    }
  }

  const workbookNumberFormats = numberFormats && typeof numberFormats === 'object' ? numberFormats : {}
  const descriptors = xfs.map((xf) => styleDescriptor(xf, fonts, palette, workbookNumberFormats))
  const sortedPositions = [...new Set(boundSheets.map((item) => item.position))].sort((a, b) => a - b)
  const sheets = boundSheets.map((sheet) => {
    const positionIndex = sortedPositions.indexOf(sheet.position)
    const end = positionIndex >= 0 && sortedPositions[positionIndex + 1] != null
      ? sortedPositions[positionIndex + 1]
      : stream.length
    const properties = {}
    const rowHeights = {}
    const pageSetup = {}
    const headerFooter = {}
    for (const record of records(stream, sheet.position, end)) {
      if (record.type === RECORD.EOF) break
      if (record.type === RECORD.DEFAULT_ROW_HEIGHT && record.data.length >= 4 && !(record.data.readUInt16LE(0) & 2)) {
        const twips = record.data.readInt16LE(2)
        if (twips > 0 && twips <= 8179) properties.defaultRowHeight = twips / 20
      } else if (record.type === RECORD.DEFAULT_COL_WIDTH && record.data.length >= 2) {
        const characters = record.data.readUInt16LE(0)
        if (characters > 0 && characters <= 255) properties.defaultColWidth = characters
      } else if (record.type === RECORD.ROW && record.data.length >= 16) {
        // ROW.miyRw records the displayed height even when Excel calculated it
        // automatically. SheetJS only exposes heights with fUnsynced set, so
        // common 16pt automatic rows otherwise shrink to our 15pt fallback.
        const row = record.data.readUInt16LE(0) + 1
        const twips = record.data.readUInt16LE(6) & 0x7fff
        if (twips > 0 && twips <= 8179) rowHeights[String(row)] = twips / 20
      } else if ([RECORD.HEADER, RECORD.FOOTER].includes(record.type) && record.data.length >= 3) {
        const characters = record.data.readUInt16LE(0)
        const wide = Boolean(record.data[2] & 1)
        const endOfText = 3 + characters * (wide ? 2 : 1)
        if (endOfText <= record.data.length) {
          headerFooter[record.type === RECORD.HEADER ? 'oddHeader' : 'oddFooter'] = record.data.subarray(3, endOfText).toString(wide ? 'utf16le' : 'latin1')
        }
      } else if (record.type === RECORD.SETUP && record.data.length >= 34) {
        const flags = record.data.readUInt16LE(10)
        pageSetup.fitToWidth = record.data.readUInt16LE(6)
        pageSetup.fitToHeight = record.data.readUInt16LE(8)
        pageSetup.pageOrder = flags & 1 ? 'overThenDown' : 'downThenOver'
        pageSetup.blackAndWhite = Boolean(flags & 8)
        pageSetup.draft = Boolean(flags & 16)
        if (!(flags & 4)) {
          pageSetup.paperSize = record.data.readUInt16LE(0)
          const scale = record.data.readUInt16LE(2)
          if (scale >= 10 && scale <= 400) pageSetup.scale = scale
          if (!(flags & 64)) pageSetup.orientation = flags & 2 ? 'portrait' : 'landscape'
          pageSetup.horizontalDpi = record.data.readUInt16LE(12)
          pageSetup.verticalDpi = record.data.readUInt16LE(14)
        }
        pageSetup.useFirstPageNumber = Boolean(flags & 128)
        if (flags & 128) pageSetup.firstPageNumber = record.data.readInt16LE(4)
      } else if (record.type === RECORD.WS_BOOL && record.data.length >= 2) {
        pageSetup.fitToPage = Boolean(record.data.readUInt16LE(0) & 0x100)
      } else if ([RECORD.PRINT_GRID, RECORD.PRINT_ROW_COL, RECORD.HCENTER, RECORD.VCENTER].includes(record.type) && record.data.length >= 2) {
        const key = { [RECORD.PRINT_GRID]: 'showGridLines', [RECORD.PRINT_ROW_COL]: 'showRowColHeaders', [RECORD.HCENTER]: 'horizontalCentered', [RECORD.VCENTER]: 'verticalCentered' }[record.type]
        pageSetup[key] = Boolean(record.data.readUInt16LE(0))
      }
    }
    return { cells: parseSheetStyles(stream, sheet.position, end, descriptors), properties, rowHeights, pageSetup, headerFooter }
  })

  const normalFont = xfs[normalXfIndex] ? fontForXf(xfs[normalXfIndex], fonts, palette) : undefined
  return { sheets, normalFont, fontCount: fonts.filter(Boolean).length, xfCount: xfs.filter(Boolean).length }
}

module.exports = {
  extractLegacyBiffStyles,
}
