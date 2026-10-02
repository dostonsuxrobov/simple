'use strict'

// A small, dependency-free WordprocessingML writer for documents that Simple
// imports natively (RTF, HTML or Word 2003 XML saved with a .doc name). It
// writes a deliberately plain package that the editor's DOCX importer reads:
// paragraphs, headings, lists, character formatting, links, tables and
// PNG/JPEG/GIF pictures. Anything else stays text.

const JSZip = require('jszip')
const { validateDocxBytes } = require('./docx-files.cjs')

const ZIP_ENTRY_DATE = new Date('1980-01-01T00:00:00.000Z')
const EMU_PER_PIXEL = 9525
const TWIPS_PER_PIXEL = 15
const MAX_IMAGE_WIDTH_PX = 624 // 6.5 in, the content width of Letter with 1 in margins.
const IMAGE_TYPES = Object.freeze({
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
})

function isXmlCharacter(codePoint) {
  return codePoint === 0x9 || codePoint === 0xa || codePoint === 0xd
    || (codePoint >= 0x20 && codePoint <= 0xd7ff)
    || (codePoint >= 0xe000 && codePoint <= 0xfffd)
    || (codePoint >= 0x10000 && codePoint <= 0x10ffff)
}

function xmlText(value) {
  let safe = ''
  for (const character of String(value ?? '')) {
    if (isXmlCharacter(character.codePointAt(0))) safe += character
  }
  return safe.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

function hexColor(value) {
  const match = /^#?([0-9a-f]{6})$/i.exec(String(value || '').trim())
  return match ? match[1].toUpperCase() : null
}

function sniffImage(bytes) {
  if (!bytes || bytes.length < 12) return null
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'image/png'
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg'
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x38) return 'image/gif'
  return null
}

/** Intrinsic pixel size of a PNG, GIF or baseline/progressive JPEG. */
function imagePixelSize(bytes) {
  try {
    const type = sniffImage(bytes)
    if (type === 'image/png' && bytes.length >= 24) return { width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) }
    if (type === 'image/gif') return { width: bytes.readUInt16LE(6), height: bytes.readUInt16LE(8) }
    if (type === 'image/jpeg') {
      let offset = 2
      while (offset + 9 < bytes.length) {
        if (bytes[offset] !== 0xff) { offset += 1; continue }
        const marker = bytes[offset + 1]
        if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { offset += 2; continue }
        const length = bytes.readUInt16BE(offset + 2)
        if ((marker >= 0xc0 && marker <= 0xcf) && ![0xc4, 0xc8, 0xcc].includes(marker)) {
          return { width: bytes.readUInt16BE(offset + 7), height: bytes.readUInt16BE(offset + 5) }
        }
        offset += 2 + length
      }
    }
  } catch { /* Fall through to the caller's default. */ }
  return null
}

function fittedImageSize(image) {
  const intrinsic = imagePixelSize(image.data) || { width: 320, height: 240 }
  let width = Number(image.widthPx) > 0 ? Number(image.widthPx) : intrinsic.width
  let height = Number(image.heightPx) > 0 ? Number(image.heightPx) : intrinsic.height
  if (!(Number(image.heightPx) > 0) && Number(image.widthPx) > 0 && intrinsic.width) height = intrinsic.height * width / intrinsic.width
  if (!(Number(image.widthPx) > 0) && Number(image.heightPx) > 0 && intrinsic.height) width = intrinsic.width * height / intrinsic.height
  if (width > MAX_IMAGE_WIDTH_PX) { height = height * MAX_IMAGE_WIDTH_PX / width; width = MAX_IMAGE_WIDTH_PX }
  return { width: Math.max(1, Math.round(width)), height: Math.max(1, Math.round(height)) }
}

function createPackageState() {
  return { relationships: [], media: [], nextRelationship: 1, nextDrawing: 1, numbering: [], usesNumbering: false }
}

function addRelationship(state, type, target, external = false) {
  const id = `rId${state.nextRelationship++}`
  state.relationships.push({ id, type, target, external })
  return id
}

function runPropertiesXml(run) {
  const parts = []
  if (run.link) parts.push('<w:rStyle w:val="Hyperlink"/>')
  if (run.font) {
    const font = xmlText(String(run.font).split(',')[0].replace(/["']/g, '').trim().slice(0, 64))
    if (font) parts.push(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}" w:eastAsia="${font}"/>`)
  }
  if (run.bold) parts.push('<w:b/>')
  if (run.italic) parts.push('<w:i/>')
  if (run.caps) parts.push('<w:caps/>')
  if (run.smallCaps) parts.push('<w:smallCaps/>')
  if (run.strike) parts.push('<w:strike/>')
  const color = hexColor(run.color)
  if (color) parts.push(`<w:color w:val="${color}"/>`)
  const size = Math.round(Number(run.sizeHalfPoints))
  if (size >= 2 && size <= 3276) parts.push(`<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`)
  const highlight = hexColor(run.highlight)
  if (highlight) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${highlight}"/>`)
  if (run.underline) parts.push('<w:u w:val="single"/>')
  if (run.vertAlign === 'superscript' || run.vertAlign === 'subscript') parts.push(`<w:vertAlign w:val="${run.vertAlign}"/>`)
  return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : ''
}

function imageRunXml(image, state) {
  const type = sniffImage(image.data)
  if (!type) return null
  const extension = IMAGE_TYPES[type]
  const index = state.media.length + 1
  const name = `media/image${index}.${extension}`
  state.media.push({ name, data: image.data, extension, type })
  const relationship = addRelationship(state, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', name)
  const { width, height } = fittedImageSize(image)
  const cx = width * EMU_PER_PIXEL
  const cy = height * EMU_PER_PIXEL
  const id = state.nextDrawing++
  const description = image.alt ? ` descr="${xmlText(String(image.alt).slice(0, 1000))}"` : ''
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Picture ${id}"${description}/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="image${index}.${extension}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationship}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
}

function textRunXml(run) {
  const properties = runPropertiesXml(run)
  if (run.break === 'line') return `<w:r>${properties}<w:br/></w:r>`
  if (run.break === 'page') return '<w:r><w:br w:type="page"/></w:r>'
  if (run.break === 'tab') return `<w:r>${properties}<w:tab/></w:r>`
  const text = String(run.text ?? '')
  if (!text) return ''
  const pieces = text.split(/(\t|\n)/)
  return pieces.map((piece) => {
    if (piece === '\t') return `<w:r>${properties}<w:tab/></w:r>`
    if (piece === '\n') return `<w:r>${properties}<w:br/></w:r>`
    return piece ? `<w:r>${properties}<w:t xml:space="preserve">${xmlText(piece)}</w:t></w:r>` : ''
  }).join('')
}

function safeLinkTarget(value) {
  const link = String(value || '').trim()
  if (/^#[\w.-]{1,120}$/.test(link)) return { anchor: link.slice(1) }
  try {
    const parsed = new URL(link)
    if (['http:', 'https:', 'mailto:'].includes(parsed.protocol)) return { url: parsed.toString() }
  } catch {}
  return null
}

function runsXml(runs, state) {
  const output = []
  for (let index = 0; index < runs.length;) {
    const run = runs[index]
    if (run?.image) {
      output.push(imageRunXml(run.image, state) || '')
      index += 1
      continue
    }
    const target = run?.link ? safeLinkTarget(run.link) : null
    if (!target) {
      output.push(textRunXml({ ...run, link: undefined }))
      index += 1
      continue
    }
    // Consecutive runs with one target become a single hyperlink.
    const group = []
    while (index < runs.length && runs[index]?.link === run.link && !runs[index]?.image) group.push(runs[index++])
    const inner = group.map((item) => textRunXml(item)).join('')
    if (!inner) continue
    if (target.anchor) output.push(`<w:hyperlink w:anchor="${xmlText(target.anchor)}">${inner}</w:hyperlink>`)
    else {
      const relationship = addRelationship(state, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', target.url, true)
      output.push(`<w:hyperlink r:id="${relationship}">${inner}</w:hyperlink>`)
    }
  }
  return output.join('')
}

function twips(value) {
  const number = Math.round(Number(value))
  return Number.isFinite(number) ? number : 0
}

function listNumberId(block, state) {
  state.usesNumbering = true
  if (!block.list.ordered) return 1
  const key = block.list.id ?? 'default'
  let entry = state.numbering.find((item) => item.key === key)
  if (!entry) {
    entry = { key, id: state.numbering.length + 2 }
    state.numbering.push(entry)
  }
  return entry.id
}

function paragraphXml(block, state) {
  const properties = []
  if (block.heading >= 1 && block.heading <= 6) properties.push(`<w:pStyle w:val="Heading${Math.floor(block.heading)}"/>`)
  else if (block.list) properties.push('<w:pStyle w:val="ListParagraph"/>')
  else if (block.style === 'Quote' || block.style === 'Code') properties.push(`<w:pStyle w:val="${block.style}"/>`)
  if (block.pageBreakBefore) properties.push('<w:pageBreakBefore/>')
  if (block.list) {
    const level = Math.max(0, Math.min(8, Math.floor(Number(block.list.level) || 0)))
    properties.push(`<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${listNumberId(block, state)}"/></w:numPr>`)
  }
  const spacing = []
  if (Number.isFinite(block.spaceBeforeTw)) spacing.push(`w:before="${Math.max(0, twips(block.spaceBeforeTw))}"`)
  if (Number.isFinite(block.spaceAfterTw)) spacing.push(`w:after="${Math.max(0, twips(block.spaceAfterTw))}"`)
  if (spacing.length) properties.push(`<w:spacing ${spacing.join(' ')}/>`)
  const indents = []
  if (Number.isFinite(block.indentLeftTw) && !block.list) indents.push(`w:left="${twips(block.indentLeftTw)}"`)
  if (Number.isFinite(block.indentRightTw)) indents.push(`w:right="${twips(block.indentRightTw)}"`)
  if (Number.isFinite(block.indentFirstTw) && !block.list) {
    const first = twips(block.indentFirstTw)
    indents.push(first < 0 ? `w:hanging="${-first}"` : `w:firstLine="${first}"`)
  }
  if (indents.length) properties.push(`<w:ind ${indents.join(' ')}/>`)
  if (['center', 'right', 'justify'].includes(block.align)) properties.push(`<w:jc w:val="${block.align === 'justify' ? 'both' : block.align}"/>`)
  if (block.rtl) properties.push('<w:bidi/>')
  const content = runsXml(Array.isArray(block.runs) ? block.runs : [], state)
  return `<w:p>${properties.length ? `<w:pPr>${properties.join('')}</w:pPr>` : ''}${content}</w:p>`
}

function tableXml(block, state) {
  const rows = (block.rows || []).filter((row) => row?.cells?.length)
  if (!rows.length) return ''
  const columns = Math.max(1, ...rows.map((row) => row.cells.reduce((sum, cell) => sum + Math.max(1, Math.floor(cell.colSpan || 1)), 0)))
  const contentWidth = 9360
  const widths = Array.isArray(block.widthsTw) && block.widthsTw.length === columns && block.widthsTw.every((value) => value > 0)
    ? block.widthsTw.map(twips)
    : Array.from({ length: columns }, () => Math.floor(contentWidth / columns))
  const grid = widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')
  const body = rows.map((row, rowIndex) => {
    let column = 0
    const cells = row.cells.map((cell) => {
      const span = Math.max(1, Math.min(columns - column, Math.floor(cell.colSpan || 1)))
      const width = widths.slice(column, column + span).reduce((sum, value) => sum + value, 0)
      column += span
      const properties = [`<w:tcW w:w="${width}" w:type="dxa"/>`]
      if (span > 1) properties.push(`<w:gridSpan w:val="${span}"/>`)
      const fill = hexColor(cell.shading)
      if (fill) properties.push(`<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`)
      const content = blocksXml(cell.blocks?.length ? cell.blocks : [{ type: 'paragraph', runs: [] }], state, true)
      return `<w:tc><w:tcPr>${properties.join('')}</w:tcPr>${content}</w:tc>`
    }).join('')
    const header = rowIndex === 0 && block.headerRow ? '<w:trPr><w:tblHeader/></w:trPr>' : ''
    return `<w:tr>${header}${cells}</w:tr>`
  }).join('')
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`
}

function blocksXml(blocks, state, insideCell = false) {
  const output = []
  for (const block of blocks || []) {
    if (!block) continue
    if (block.type === 'table') {
      const table = tableXml(block, state)
      if (!table) continue
      output.push(table)
      // Word requires a paragraph after a table inside a cell and between tables.
      if (insideCell) output.push('<w:p/>')
    } else if (block.type === 'image') output.push(paragraphXml({ align: block.align, runs: [{ image: block }] }, state))
    else output.push(paragraphXml(block, state))
  }
  if (insideCell && (!output.length || /<\/w:tbl>$/.test(output.at(-1)))) output.push('<w:p/>')
  return output.join('')
}

function levelXml(level, ordered) {
  const indent = 720 * (level + 1)
  if (!ordered) {
    const bullets = ['•', '◦', '▪']
    return `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bullets[level % 3]}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr></w:lvl>`
  }
  const formats = ['decimal', 'lowerLetter', 'lowerRoman']
  return `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="${formats[level % 3]}"/><w:lvlText w:val="%${level + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr></w:lvl>`
}

function numberingXml(state) {
  const levels = (ordered) => Array.from({ length: 9 }, (_value, level) => levelXml(level, ordered)).join('')
  const ordered = state.numbering.map((entry) => `<w:num w:numId="${entry.id}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${levels(false)}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${levels(true)}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>${ordered}</w:numbering>`
}

function stylesXml() {
  const headingSizes = [32, 26, 24, 22, 22, 22]
  const headings = headingSizes.map((size, index) => `<w:style w:type="paragraph" w:styleId="Heading${index + 1}"><w:name w:val="heading ${index + 1}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${index ? 160 : 240}" w:after="80"/><w:outlineLvl w:val="${index}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${headings}<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/><w:contextualSpacing/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`
}

function sectionXml(page = {}, bandReferences = '') {
  const width = twips(page.widthTw) > 1440 ? twips(page.widthTw) : 12240
  const height = twips(page.heightTw) > 1440 ? twips(page.heightTw) : 15840
  const margin = (key) => {
    const value = twips(page.margins?.[key])
    return value >= 0 && value <= 4320 && page.margins?.[key] !== undefined ? value : 1440
  }
  const orient = width > height ? ' w:orient="landscape"' : ''
  return `<w:sectPr>${bandReferences}<w:pgSz w:w="${width}" w:h="${height}"${orient}/><w:pgMar w:top="${margin('top')}" w:right="${margin('right')}" w:bottom="${margin('bottom')}" w:left="${margin('left')}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
}

const NAMESPACES = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

/**
 * Build a validated DOCX from a flow document:
 * { title, blocks, header?, footer?, page? } where blocks are
 * { type: 'paragraph', runs, heading?, list?: {ordered, level, id}, align? },
 * { type: 'table', rows: [{ cells: [{ blocks, colSpan?, shading? }] }] } and
 * { type: 'image', data, widthPx?, heightPx?, align? }.
 */
async function buildFlowDocx(flow, options = {}) {
  const state = createPackageState()
  const blocks = Array.isArray(flow?.blocks) && flow.blocks.length ? flow.blocks : [{ type: 'paragraph', runs: [] }]
  const bodyXml = blocksXml(blocks, state)
  const parts = {}
  const overrides = []
  let bandReferences = ''
  for (const kind of ['header', 'footer']) {
    const bandBlocks = flow?.[kind]
    if (!Array.isArray(bandBlocks) || !bandBlocks.length) continue
    // Bands get their own relationship scope; pictures stay in the body only.
    const bandState = createPackageState()
    const xml = blocksXml(bandBlocks.map((block) => block.type === 'paragraph' ? { ...block, runs: (block.runs || []).filter((run) => !run.image) } : block).filter((block) => block.type !== 'image'), bandState)
    const root = kind === 'header' ? 'w:hdr' : 'w:ftr'
    parts[`word/${kind}1.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${root} ${NAMESPACES}>${xml || '<w:p/>'}</${root}>`
    const bandRelationships = bandState.relationships.filter((item) => item.external)
    if (bandRelationships.length) parts[`word/_rels/${kind}1.xml.rels`] = relationshipsXml(bandRelationships)
    const id = addRelationship(state, `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`, `${kind}1.xml`)
    bandReferences += `<w:${kind}Reference w:type="default" r:id="${id}"/>`
    overrides.push(`<Override PartName="/word/${kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`)
  }
  addRelationship(state, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'styles.xml')
  if (state.usesNumbering) {
    addRelationship(state, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', 'numbering.xml')
    parts['word/numbering.xml'] = numberingXml(state)
    overrides.push('<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>')
  }
  parts['word/document.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NAMESPACES}><w:body>${bodyXml}${sectionXml(flow?.page, bandReferences)}</w:body></w:document>`
  parts['word/styles.xml'] = stylesXml()
  parts['word/_rels/document.xml.rels'] = relationshipsXml(state.relationships)
  for (const item of state.media) parts[`word/${item.name}`] = item.data
  const title = xmlText(String(options.title ?? flow?.title ?? '').slice(0, 240))
  parts['docProps/core.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${title}</dc:title><cp:lastModifiedBy>Simple Docs</cp:lastModifiedBy></cp:coreProperties>`
  parts['docProps/app.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Simple Docs</Application></Properties>`
  const mediaDefaults = [...new Set(state.media.map((item) => item.extension))]
    .map((extension) => `<Default Extension="${extension}" ContentType="image/${extension}"/>`).join('')
  parts['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${mediaDefaults}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${overrides.join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`
  parts['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`

  const zip = new JSZip()
  // [Content_Types].xml first, as Word itself writes it.
  for (const name of ['[Content_Types].xml', ...Object.keys(parts).filter((name) => name !== '[Content_Types].xml')]) {
    zip.file(name, parts[name], { date: ZIP_ENTRY_DATE })
  }
  const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, platform: 'DOS' })
  return validateDocxBytes(bytes)
}

function relationshipsXml(relationships) {
  const items = relationships.map((item) => `<Relationship Id="${item.id}" Type="${item.type}" Target="${xmlText(item.target)}"${item.external ? ' TargetMode="External"' : ''}/>`).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`
}

/** Plain text of a flow document, for warnings and tests. */
function flowText(blocks) {
  const output = []
  for (const block of blocks || []) {
    if (block?.type === 'paragraph') output.push((block.runs || []).map((run) => run.break === 'tab' ? '\t' : run.break ? '\n' : run.text || '').join(''))
    else if (block?.type === 'table') for (const row of block.rows || []) for (const cell of row.cells || []) output.push(flowText(cell.blocks))
  }
  return output.join('\n')
}

module.exports = {
  EMU_PER_PIXEL,
  TWIPS_PER_PIXEL,
  buildFlowDocx,
  flowText,
  hexColor,
  imagePixelSize,
  sniffImage,
  xmlText,
}
