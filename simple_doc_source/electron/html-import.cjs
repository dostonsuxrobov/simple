'use strict'

// Native readers for the markup formats that are often saved with a .doc name:
// HTML (including Word's "Web Page" output), single-file web archives (MHT)
// and Word 2003 XML. No DOM, network or script is involved: remote pictures are
// never fetched, and only embedded PNG/JPEG/GIF data is kept.

const { sniffImage } = require('./simple-docx.cjs')

const MAX_MARKUP_BYTES = 128 * 1024 * 1024
const MAX_DEPTH = 256
const VOID_ELEMENTS = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr', 'basefont', 'frame', 'keygen'])
const RAW_TEXT_ELEMENTS = new Set(['script', 'style', 'textarea', 'title', 'xmp', 'noscript', 'template', 'svg', 'math', 'object', 'iframe', 'xml'])
const BLOCK_ELEMENTS = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'table', 'tr', 'td', 'th', 'thead', 'tbody', 'tfoot', 'caption', 'blockquote', 'pre', 'hr', 'address', 'section', 'article', 'header', 'footer', 'main', 'nav', 'aside', 'figure', 'figcaption', 'dl', 'dt', 'dd', 'center', 'form', 'fieldset', 'body', 'html', 'details', 'summary'])
const CLOSES_PARAGRAPH = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'table', 'blockquote', 'pre', 'hr', 'address', 'section', 'article', 'header', 'footer', 'main', 'nav', 'aside', 'figure', 'dl', 'center', 'form', 'fieldset', 'details'])
const NAMED_ENTITIES = Object.freeze({
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', copy: '©', reg: '®', trade: '™', hellip: '…',
  mdash: '—', ndash: '–', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', bull: '•', middot: '·',
  laquo: '«', raquo: '»', euro: '€', pound: '£', yen: '¥', cent: '¢', sect: '§', para: '¶',
  deg: '°', plusmn: '±', times: '×', divide: '÷', frac12: '½', frac14: '¼', frac34: '¾',
  sup2: '²', sup3: '³', shy: '­', ensp: ' ', emsp: ' ', thinsp: ' ', zwj: '‍', zwnj: '‌',
  iexcl: '¡', iquest: '¿', szlig: 'ß', aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó',
  uacute: 'ú', agrave: 'à', egrave: 'è', ograve: 'ò', auml: 'ä', euml: 'ë', ouml: 'ö', uuml: 'ü',
  Auml: 'Ä', Ouml: 'Ö', Uuml: 'Ü', ccedil: 'ç', ntilde: 'ñ', aring: 'å', oslash: 'ø', aelig: 'æ',
  rarr: '→', larr: '←', uarr: '↑', darr: '↓', check: '✓',
})
const NAMED_COLORS = Object.freeze({
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', gray: '#808080', grey: '#808080',
  silver: '#c0c0c0', maroon: '#800000', navy: '#000080', purple: '#800080', teal: '#008080', olive: '#808000', lime: '#00ff00',
  aqua: '#00ffff', fuchsia: '#ff00ff', orange: '#ffa500', darkblue: '#00008b', darkred: '#8b0000', darkgreen: '#006400',
})

function decodeEntities(value) {
  return String(value).replace(/&(#x[0-9a-f]{1,6}|#\d{1,7}|[a-z][a-z0-9]{1,31});?/gi, (entity, name) => {
    if (name[0] === '#') {
      const code = name[1] === 'x' || name[1] === 'X' ? Number.parseInt(name.slice(2), 16) : Number.parseInt(name.slice(1), 10)
      if (code === 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return '�'
      // Windows-1252 code points used as numeric references by old HTML.
      if (code >= 0x80 && code <= 0x9f) return new TextDecoder('windows-1252').decode(Uint8Array.of(code))
      return String.fromCodePoint(code)
    }
    return Object.hasOwn(NAMED_ENTITIES, name) ? NAMED_ENTITIES[name] : Object.hasOwn(NAMED_ENTITIES, name.toLowerCase()) ? NAMED_ENTITIES[name.toLowerCase()] : entity
  })
}

function parseAttributes(source) {
  const attributes = {}
  for (const match of String(source).matchAll(/([^\s"'<>/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g)) {
    const name = match[1].toLowerCase()
    if (!Object.hasOwn(attributes, name)) attributes[name] = decodeEntities(match[2] ?? match[3] ?? match[4] ?? '')
  }
  return attributes
}

/** Tolerant tree builder for HTML (and, with xml:true, Word 2003 XML). */
function parseMarkup(text, options = {}) {
  const xml = Boolean(options.xml)
  const root = { name: '#root', attributes: {}, children: [] }
  const stack = [root]
  let title = ''
  const top = () => stack[stack.length - 1]
  const closeTo = (name) => {
    for (let index = stack.length - 1; index > 0; index -= 1) {
      if (stack[index].name === name) { stack.length = index; return true }
    }
    return false
  }
  const openInScope = (name, boundaries) => {
    for (let index = stack.length - 1; index > 0; index -= 1) {
      if (stack[index].name === name) return index
      if (boundaries.has(stack[index].name)) return -1
    }
    return -1
  }
  let index = 0
  const length = text.length
  while (index < length) {
    const open = text.indexOf('<', index)
    const end = open < 0 ? length : open
    if (end > index) top().children.push(decodeEntities(text.slice(index, end)))
    if (open < 0) break
    if (text.startsWith('<!--', open)) {
      const close = text.indexOf('-->', open + 4)
      index = close < 0 ? length : close + 3
      continue
    }
    if (text.startsWith('<![CDATA[', open)) {
      const close = text.indexOf(']]>', open)
      top().children.push(text.slice(open + 9, close < 0 ? length : close))
      index = close < 0 ? length : close + 3
      continue
    }
    if (text.startsWith('<![', open)) {
      // Word's downlevel-revealed conditionals: <![if !supportLists]> … <![endif]>.
      const close = text.indexOf(']>', open)
      const marker = text.slice(open + 3, close < 0 ? length : close).trim().toLowerCase()
      if (marker.startsWith('if') && marker.includes('supportlists')) {
        const finish = text.indexOf('<![endif]>', close)
        top().children.push({ name: '#list-marker', attributes: {}, children: [stripTags(text.slice(close + 2, finish < 0 ? length : finish))] })
        index = finish < 0 ? length : finish + 10
        continue
      }
      index = close < 0 ? length : close + 2
      continue
    }
    if (text[open + 1] === '!' || text[open + 1] === '?') {
      const close = text.indexOf('>', open)
      index = close < 0 ? length : close + 1
      continue
    }
    const tag = /^<\/?([a-zA-Z][\w:.-]*)((?:\s+[^\s"'<>/=]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?)*)\s*(\/?)>/.exec(text.slice(open, open + 65536))
    if (!tag) {
      top().children.push('<')
      index = open + 1
      continue
    }
    index = open + tag[0].length
    const rawName = tag[1]
    const name = xml ? rawName : rawName.toLowerCase()
    if (tag[0][1] === '/') {
      if (!xml && name === 'p' && openInScope('p', new Set(['td', 'th', 'li', 'blockquote', 'div'])) < 0) continue
      closeTo(name)
      continue
    }
    const attributes = parseAttributes(tag[2])
    const selfClosing = tag[3] === '/' || (!xml && VOID_ELEMENTS.has(name))
    if (!xml) {
      if (CLOSES_PARAGRAPH.has(name) || name === 'li' || name === 'tr' || name === 'td' || name === 'th') {
        const paragraph = openInScope('p', new Set(['td', 'th', 'li', 'blockquote', 'table', 'div']))
        if (paragraph > 0) stack.length = paragraph
      }
      if (name === 'li') { const item = openInScope('li', new Set(['ul', 'ol'])); if (item > 0) stack.length = item }
      if (name === 'td' || name === 'th') { const cell = Math.max(openInScope('td', new Set(['tr', 'table'])), openInScope('th', new Set(['tr', 'table']))); if (cell > 0) stack.length = cell }
      if (name === 'tr') { const row = openInScope('tr', new Set(['table'])); if (row > 0) stack.length = row }
      if (name === 'dt' || name === 'dd') { const item = Math.max(openInScope('dt', new Set(['dl'])), openInScope('dd', new Set(['dl']))); if (item > 0) stack.length = item }
      if (RAW_TEXT_ELEMENTS.has(name) && !selfClosing) {
        const closing = new RegExp(`</${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*>`, 'ig')
        closing.lastIndex = index
        const close = closing.exec(text)?.index ?? -1
        const content = text.slice(index, close < 0 ? length : close)
        if (name === 'title' && !title) title = decodeEntities(content).replace(/\s+/g, ' ').trim()
        const closeEnd = close < 0 ? length : text.indexOf('>', close)
        index = closeEnd < 0 ? length : closeEnd + 1
        continue
      }
    }
    const element = { name, attributes, children: [] }
    top().children.push(element)
    if (!selfClosing && stack.length < MAX_DEPTH) stack.push(element)
  }
  return { root, title }
}

function stripTags(value) {
  return decodeEntities(String(value).replace(/<!--[\s\S]*?-->/g, '').replace(/<[^>]*>/g, '')).replace(/\s+/g, ' ').trim()
}

function parseStyle(value) {
  const style = {}
  for (const declaration of String(value || '').split(';')) {
    const colon = declaration.indexOf(':')
    if (colon < 1) continue
    style[declaration.slice(0, colon).trim().toLowerCase()] = declaration.slice(colon + 1).trim()
  }
  return style
}

function cssColor(value) {
  const color = String(value || '').trim().toLowerCase()
  if (!color || color === 'transparent' || color === 'inherit' || color === 'windowtext' || color === 'auto') return null
  let match = /^#([0-9a-f]{6})$/.exec(color)
  if (match) return `#${match[1]}`
  match = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(color)
  if (match) return `#${match[1]}${match[1]}${match[2]}${match[2]}${match[3]}${match[3]}`
  match = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/.exec(color)
  if (match) return `#${match.slice(1, 4).map((part) => Math.min(255, Number(part)).toString(16).padStart(2, '0')).join('')}`
  return NAMED_COLORS[color] || null
}

function cssHalfPoints(value) {
  const match = /^([\d.]+)\s*(pt|px|em|rem|%)?$/.exec(String(value || '').trim().toLowerCase())
  if (!match) {
    const keywords = { 'xx-small': 14, 'x-small': 15, small: 20, medium: 24, large: 27, 'x-large': 36, 'xx-large': 48 }
    return keywords[String(value || '').trim().toLowerCase()] || null
  }
  const number = Number(match[1])
  const unit = match[2] || 'px'
  const points = unit === 'pt' ? number : unit === 'px' ? number * 0.75 : unit === '%' ? 12 * number / 100 : 12 * number
  return points > 0 && points < 1638 ? Math.round(points * 2) : null
}

function cssLengthTwips(value) {
  const match = /^(-?[\d.]+)\s*(pt|px|in|cm|mm|em)?$/.exec(String(value || '').trim().toLowerCase())
  if (!match) return null
  const number = Number(match[1])
  const factor = { pt: 20, px: 15, in: 1440, cm: 567, mm: 56.7, em: 240 }[match[2] || 'px']
  return Math.round(number * factor)
}

function decodeDataImage(source) {
  const match = /^data:(image\/[a-z0-9.+-]+)?(;[^,]*)?,(.*)$/is.exec(String(source || '').trim())
  if (!match || !/;base64/i.test(match[2] || '')) return null
  try {
    const bytes = Buffer.from(match[3].replace(/\s+/g, ''), 'base64')
    return sniffImage(bytes) ? bytes : null
  } catch { return null }
}

const HEADING = /^h([1-6])$/

/** Convert a parsed HTML tree to a flow document. */
function htmlTreeToFlow(root, options = {}) {
  const resources = options.resources || new Map()
  const warnings = { remoteImages: 0, unsupportedImages: 0 }
  const blocks = []
  let listCounter = 0
  const FORMAT_KEYS = ['bold', 'italic', 'underline', 'strike', 'vertAlign', 'font', 'highlight', 'sizeHalfPoints', 'color', 'link', 'caps', 'smallCaps']

  function characterFrom(element, inherited) {
    const character = { ...inherited }
    const name = element.name
    const style = parseStyle(element.attributes.style)
    if (name === 'b' || name === 'strong' || name === 'th') character.bold = true
    if (name === 'i' || name === 'em' || name === 'cite' || name === 'dfn' || name === 'var') character.italic = true
    if (name === 'u' || name === 'ins') character.underline = true
    if (name === 's' || name === 'strike' || name === 'del') character.strike = true
    if (name === 'sup') character.vertAlign = 'superscript'
    if (name === 'sub') character.vertAlign = 'subscript'
    if (name === 'code' || name === 'kbd' || name === 'tt' || name === 'samp') character.font = 'Consolas'
    if (name === 'mark') character.highlight = '#ffff00'
    if (name === 'small') character.sizeHalfPoints = Math.max(2, Math.round((character.sizeHalfPoints || 22) * 0.83))
    if (name === 'big') character.sizeHalfPoints = Math.round((character.sizeHalfPoints || 22) * 1.2)
    if (name === 'font') {
      if (element.attributes.color) character.color = cssColor(element.attributes.color) || character.color
      if (element.attributes.face) character.font = element.attributes.face.split(',')[0].replace(/["']/g, '').trim() || character.font
      const size = Number(element.attributes.size)
      if (size >= 1 && size <= 7) character.sizeHalfPoints = [16, 20, 24, 27, 36, 48, 72][size - 1]
    }
    if (name === 'a' && element.attributes.href) {
      const href = element.attributes.href.trim()
      if (/^(https?:|mailto:)/i.test(href) || /^#[\w.-]+$/.test(href)) character.link = href
    }
    if (style['font-weight']) character.bold = /bold|bolder|[6-9]00/.test(style['font-weight'])
    if (style['font-style']) character.italic = /italic|oblique/.test(style['font-style'])
    const decoration = style['text-decoration'] || style['text-decoration-line']
    if (decoration) {
      if (/none/.test(decoration)) { character.underline = false; character.strike = false }
      if (/underline/.test(decoration)) character.underline = true
      if (/line-through/.test(decoration)) character.strike = true
    }
    if (style.color) character.color = cssColor(style.color) || character.color
    const background = style['background-color'] || style.background || style['mso-highlight']
    if (background && !['td', 'th', 'tr', 'table'].includes(name)) character.highlight = cssColor(background.split(/\s+/)[0]) || character.highlight
    if (style['font-size']) character.sizeHalfPoints = cssHalfPoints(style['font-size']) || character.sizeHalfPoints
    if (style['font-family']) character.font = style['font-family'].split(',')[0].replace(/["']/g, '').trim() || character.font
    if (/super/.test(style['vertical-align'] || '')) character.vertAlign = 'superscript'
    if (/sub/.test(style['vertical-align'] || '')) character.vertAlign = 'subscript'
    if (/uppercase/.test(style['text-transform'] || '')) character.caps = true
    if (/small-caps/.test(style['font-variant'] || '')) character.smallCaps = true
    if (/none/.test(style.display || '') || /hidden/.test(style.visibility || '') || /^all$/i.test(style['mso-hide'] || '')) character.hidden = true
    return character
  }

  function createBuilder(output) {
    return { output, paragraph: null }
  }

  function ensure(builder, context) {
    if (!builder.paragraph) {
      const paragraph = { type: 'paragraph', runs: [] }
      if (context.align) paragraph.align = context.align
      if (context.heading) paragraph.heading = context.heading
      if (context.style) paragraph.style = context.style
      if (context.indentLeftTw) paragraph.indentLeftTw = context.indentLeftTw
      if (context.rtl) paragraph.rtl = true
      if (context.list) {
        paragraph.list = { ...context.list }
        if (context.marker?.text !== undefined) paragraph.listMarkerText = context.marker.text
      }
      builder.paragraph = paragraph
      builder.output.push(paragraph)
    }
    return builder.paragraph
  }

  function endParagraph(builder) {
    const paragraph = builder.paragraph
    builder.paragraph = null
    if (!paragraph) return
    const visible = paragraph.runs.some((run) => run.image || run.break === 'page' || (typeof run.text === 'string' && run.text.replace(/ /g, ' ').trim()))
    if (!visible) {
      const index = builder.output.lastIndexOf(paragraph)
      if (index >= 0) builder.output.splice(index, 1)
      return
    }
    // Trim layout whitespace at the paragraph's end.
    const last = paragraph.runs.at(-1)
    if (typeof last?.text === 'string') last.text = last.text.replace(/[ \t]+$/, '')
  }

  function appendText(builder, context, text) {
    const character = context.character
    if (!text || character.hidden) return
    const paragraph = ensure(builder, context)
    const run = { text }
    for (const key of FORMAT_KEYS) if (character[key]) run[key] = character[key]
    const previous = paragraph.runs.at(-1)
    if (previous && typeof previous.text === 'string' && FORMAT_KEYS.every((key) => previous[key] === run[key])) previous.text += text
    else paragraph.runs.push(run)
  }

  function imageRun(element, character) {
    const source = element.attributes.src || ''
    let bytes = decodeDataImage(source)
    if (!bytes && resources.has(source)) bytes = resources.get(source)
    if (!bytes && /^cid:/i.test(source) && resources.has(source.slice(4))) bytes = resources.get(source.slice(4))
    if (!bytes) {
      if (/^(https?:)?\/\//i.test(source)) warnings.remoteImages += 1
      else if (source) warnings.unsupportedImages += 1
      const alt = (element.attributes.alt || '').trim()
      return alt ? { text: `[${alt}]` } : null
    }
    if (character.hidden) return null
    const style = parseStyle(element.attributes.style)
    const width = Number.parseFloat(element.attributes.width) || (style.width ? cssLengthTwips(style.width) / 15 : 0)
    const height = Number.parseFloat(element.attributes.height) || (style.height ? cssLengthTwips(style.height) / 15 : 0)
    return { image: { data: bytes, ...(width > 0 ? { widthPx: width } : {}), ...(height > 0 ? { heightPx: height } : {}), ...(element.attributes.alt ? { alt: element.attributes.alt } : {}) } }
  }

  function blockContext(node, context) {
    const style = parseStyle(node.attributes.style)
    const child = { ...context, character: characterFrom(node, context.character), marker: { text: undefined } }
    const heading = HEADING.exec(node.name)
    if (heading) child.heading = Number(heading[1])
    const align = (style['text-align'] || node.attributes.align || (node.name === 'center' ? 'center' : '')).toLowerCase()
    if (['center', 'right', 'justify'].includes(align)) child.align = align
    else if (align === 'left') child.align = undefined
    if (node.name === 'pre') { child.pre = true; child.style = 'Code' }
    if (node.name === 'blockquote') child.style = 'Quote'
    if ((node.attributes.dir || style.direction || '').toLowerCase() === 'rtl') child.rtl = true
    const margin = cssLengthTwips(style['margin-left'] || '')
    if (margin > 0 && node.name !== 'li' && !style['mso-list']) child.indentLeftTw = Math.min(8640, (context.indentLeftTw || 0) + margin)
    const msoList = /level(\d)/i.exec(style['mso-list'] || '')
    if (msoList && !/ignore/i.test(style['mso-list'])) {
      const id = /^(l\d+)/i.exec(style['mso-list'].trim())?.[1] || 'mso'
      child.list = { id: `mso-${id}`, level: Math.max(0, Math.min(8, Number(msoList[1]) - 1)), ordered: null }
    }
    return child
  }

  function walk(nodes, context, builder) {
    for (const node of nodes) {
      if (typeof node === 'string') {
        if (context.pre) {
          node.replace(/\r\n?/g, '\n').split('\n').forEach((line, index) => {
            if (index) ensure(builder, context).runs.push({ break: 'line' })
            appendText(builder, context, line)
          })
          continue
        }
        let text = node.replace(/[\t\r\n ]+/g, ' ')
        const current = builder.paragraph?.runs.at(-1)
        const atStart = !builder.paragraph || !builder.paragraph.runs.length || current?.break === 'line'
        const afterSpace = typeof current?.text === 'string' && / $/.test(current.text)
        if ((atStart || afterSpace) && text.startsWith(' ')) text = text.slice(1)
        if (text) appendText(builder, context, text)
        continue
      }
      const name = node.name
      if (name === '#list-marker') {
        if (context.marker) context.marker.text = node.children.join('')
        continue
      }
      if (name === 'br') { ensure(builder, context).runs.push({ break: 'line' }); continue }
      if (name === 'img') {
        const run = imageRun(node, context.character)
        if (run?.image) ensure(builder, context).runs.push(run)
        else if (run) appendText(builder, { ...context, character: { ...context.character, italic: true } }, run.text)
        continue
      }
      if (name === 'hr') { endParagraph(builder); continue }
      if (['head', 'title', 'meta', 'link', 'style', 'script'].includes(name)) continue
      if (name.includes(':')) {
        // Office namespace tags: <o:p> is an empty paragraph mark; VML is skipped.
        if (name === 'o:p') walk(node.children, context, builder)
        continue
      }
      if (name === 'table') {
        endParagraph(builder)
        const table = tableFrom(node, context)
        if (table) builder.output.push(table)
        continue
      }
      if (name === 'ul' || name === 'ol') {
        endParagraph(builder)
        const level = context.listContainer ? Math.min(8, context.listContainer.level + 1) : 0
        const list = { ordered: name === 'ol', level, id: `html-list-${++listCounter}` }
        walk(node.children, { ...context, listContainer: list, list: null }, builder)
        endParagraph(builder)
        continue
      }
      if (name === 'li') {
        endParagraph(builder)
        const container = context.listContainer || { ordered: false, level: 0, id: `html-list-${++listCounter}` }
        const child = blockContext(node, context)
        child.list = { ordered: container.ordered, level: container.level, id: container.id }
        walk(node.children, child, builder)
        endParagraph(builder)
        continue
      }
      if (BLOCK_ELEMENTS.has(name)) {
        endParagraph(builder)
        walk(node.children, blockContext(node, context), builder)
        endParagraph(builder)
        continue
      }
      if (name === 'q') appendText(builder, context, '“')
      walk(node.children, { ...context, character: characterFrom(node, context.character) }, builder)
      if (name === 'q') appendText(builder, context, '”')
    }
  }

  function tableFrom(node, context) {
    const rows = []
    const collectRows = (element, inHead) => {
      for (const child of element.children) {
        if (typeof child === 'string') continue
        if (child.name === 'tr') rows.push({ element: child, head: inHead })
        else if (['thead', 'tbody', 'tfoot'].includes(child.name)) collectRows(child, child.name === 'thead')
      }
    }
    collectRows(node, false)
    const result = []
    let headerRow = false
    for (const { element, head } of rows) {
      const cells = []
      for (const cell of element.children) {
        if (typeof cell === 'string' || (cell.name !== 'td' && cell.name !== 'th')) continue
        const style = parseStyle(cell.attributes.style)
        const cellBlocks = []
        const builder = createBuilder(cellBlocks)
        walk(cell.children, blockContext(cell, { ...context, list: null, listContainer: null, heading: undefined, indentLeftTw: undefined }), builder)
        endParagraph(builder)
        const shading = cssColor(style['background-color'] || (style.background || '').split(/\s+/)[0] || cell.attributes.bgcolor)
        const colSpan = Math.max(1, Math.min(63, Number.parseInt(cell.attributes.colspan, 10) || 1))
        cells.push({ blocks: cellBlocks.length ? cellBlocks : [{ type: 'paragraph', runs: [] }], ...(colSpan > 1 ? { colSpan } : {}), ...(shading ? { shading } : {}) })
      }
      if (!cells.length) continue
      if (!result.length && (head || element.children.every((cell) => typeof cell === 'string' || cell.name === 'th'))) headerRow = true
      result.push({ cells })
    }
    return result.length ? { type: 'table', rows: result, ...(headerRow ? { headerRow: true } : {}) } : null
  }

  const body = findElement(root, 'body') || root
  const builder = createBuilder(blocks)
  walk(body.children, { character: {}, marker: { text: undefined } }, builder)
  endParagraph(builder)
  finalizeLists(blocks)
  const messages = []
  if (warnings.remoteImages) messages.push(`${warnings.remoteImages} linked web ${warnings.remoteImages === 1 ? 'picture was' : 'pictures were'} not downloaded. Simple never fetches content from the internet.`)
  if (warnings.unsupportedImages) messages.push(`${warnings.unsupportedImages} ${warnings.unsupportedImages === 1 ? 'picture' : 'pictures'} stored outside this file or in an unsupported format could not be included.`)
  return { blocks, warnings: messages }
}

// Lists from Word HTML carry their marker text; it decides numbered vs bulleted.
function finalizeLists(blocks) {
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]
    if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) finalizeLists(cell.blocks)
      continue
    }
    if (block.type !== 'paragraph' || !block.list) continue
    if (block.list.ordered === null) {
      const marker = String(block.listMarkerText || '').trim()
      block.list = { ...block.list, ordered: /^[(\[]?(?:\d+|[a-z]{1,3}|[ivxlcdm]{1,6})[.)\]]/i.test(marker) }
    }
    delete block.listMarkerText
  }
}

function findElement(node, name) {
  for (const child of node.children || []) {
    if (typeof child === 'string') continue
    if (child.name === name) return child
    const found = findElement(child, name)
    if (found) return found
  }
  return null
}

function isUtf8(bytes) {
  try { new TextDecoder('utf-8', { fatal: true }).decode(bytes); return true } catch { return false }
}

function decodeMarkupBytes(input) {
  const bytes = Buffer.isBuffer(input) ? input : Buffer.from(input)
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return bytes.subarray(3).toString('utf8')
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return bytes.subarray(2).toString('utf16le')
  if (bytes[0] === 0xfe && bytes[1] === 0xff) { const swapped = Buffer.from(bytes.subarray(2)); swapped.swap16(); return swapped.toString('utf16le') }
  if (bytes.length >= 4 && bytes[0] === 0x3c && bytes[1] === 0 && bytes[2] !== 0) return bytes.toString('utf16le')
  if (isUtf8(bytes)) return bytes.toString('utf8')
  const head = bytes.subarray(0, 4096).toString('latin1')
  const declared = /charset\s*=\s*["']?([\w-]+)/i.exec(head)?.[1] || /encoding\s*=\s*["']([\w-]+)/i.exec(head)?.[1]
  try { return new TextDecoder(declared || 'windows-1252').decode(bytes) } catch { return new TextDecoder('windows-1252').decode(bytes) }
}

function htmlToFlow(input, options = {}) {
  const text = typeof input === 'string' ? input : decodeMarkupBytes(input)
  if (text.length > MAX_MARKUP_BYTES) throw new Error('This web document is too large to open safely.')
  const { root, title } = parseMarkup(text)
  const flow = htmlTreeToFlow(root, options)
  return { title, blocks: flow.blocks, warnings: flow.warnings }
}

function decodeQuotedPrintable(value) {
  const bytes = []
  const source = value.replace(/=\r?\n/g, '')
  for (let index = 0; index < source.length; index += 1) {
    if (source[index] === '=' && /^[0-9a-f]{2}$/i.test(source.slice(index + 1, index + 3))) {
      bytes.push(Number.parseInt(source.slice(index + 1, index + 3), 16))
      index += 2
    } else bytes.push(source.charCodeAt(index) & 0xff)
  }
  return Buffer.from(bytes)
}

/** Single-file web archive (MIME HTML, .mht/.mhtml), as written by Word and browsers. */
function mhtToFlow(input) {
  const text = Buffer.isBuffer(input) ? input.toString('latin1') : String(input)
  const boundary = /boundary\s*=\s*"?([^"\r\n;]+)"?/i.exec(text)?.[1]
  if (!boundary) throw new Error('This web archive has no readable parts.')
  const parts = text.split(`--${boundary}`).slice(1)
  const resources = new Map()
  let html = null
  for (const part of parts) {
    const split = part.search(/\r?\n\r?\n/)
    if (split < 0) continue
    const headers = part.slice(0, split)
    let content = part.slice(split).replace(/^\r?\n\r?\n/, '').replace(/\r?\n$/, '')
    const type = /content-type:\s*([^;\r\n]+)/i.exec(headers)?.[1]?.trim().toLowerCase() || ''
    const encoding = /content-transfer-encoding:\s*([^\s;]+)/i.exec(headers)?.[1]?.toLowerCase() || '7bit'
    const location = /content-location:\s*([^\r\n]+)/i.exec(headers)?.[1]?.trim()
    const contentId = /content-id:\s*<?([^>\r\n]+)>?/i.exec(headers)?.[1]?.trim()
    let bytes
    if (encoding === 'base64') bytes = Buffer.from(content.replace(/\s+/g, ''), 'base64')
    else if (encoding === 'quoted-printable') bytes = decodeQuotedPrintable(content)
    else bytes = Buffer.from(content, 'latin1')
    if (type === 'text/html' && !html) {
      const charset = /charset\s*=\s*"?([\w-]+)/i.exec(headers)?.[1]
      try { html = new TextDecoder(charset || (isUtf8(bytes) ? 'utf-8' : 'windows-1252')).decode(bytes) } catch { html = bytes.toString('utf8') }
    } else if (type.startsWith('image/') && sniffImage(bytes)) {
      if (location) {
        resources.set(location, bytes)
        resources.set(location.split(/[\\/]/).pop(), bytes)
      }
      if (contentId) resources.set(contentId, bytes)
    }
  }
  if (html === null) throw new Error('This web archive does not contain a web page.')
  // Relative references inside the page resolve against the part names.
  const relativeResources = new Map(resources)
  for (const [key, value] of resources) relativeResources.set(key.replace(/^.*[\\/]/, ''), value)
  const { root, title } = parseMarkup(html)
  const flow = htmlTreeToFlow(root, { resources: { has: (key) => relativeResources.has(key) || relativeResources.has(String(key).replace(/^.*[\\/]/, '')), get: (key) => relativeResources.get(key) || relativeResources.get(String(key).replace(/^.*[\\/]/, '')) } })
  return { title, blocks: flow.blocks, warnings: flow.warnings }
}

/** Word 2003 XML ("WordprocessingML 2003", <w:wordDocument>). */
function word2003ToFlow(input) {
  const text = typeof input === 'string' ? input : decodeMarkupBytes(input)
  const { root } = parseMarkup(text, { xml: true })
  const document = findElement(root, 'w:wordDocument')
  if (!document) throw new Error('This is not a Word 2003 XML document.')
  const body = findElement(document, 'w:body')
  const blocks = []
  let listCounter = 0
  const attribute = (element, name) => element?.attributes?.[name.toLowerCase()]
  const child = (element, name) => (element?.children || []).find((item) => typeof item !== 'string' && item.name === name)
  const on = (element) => element && !/^(?:off|false|0)$/i.test(attribute(element, 'w:val') || '')
  const textOf = (element) => (element.children || []).map((item) => typeof item === 'string' ? item : textOf(item)).join('')

  function runFrom(run, link, output) {
    const properties = child(run, 'w:rPr')
    const format = {}
    if (on(child(properties, 'w:b'))) format.bold = true
    if (on(child(properties, 'w:i'))) format.italic = true
    const underline = child(properties, 'w:u')
    if (underline && attribute(underline, 'w:val') !== 'none') format.underline = true
    if (on(child(properties, 'w:strike'))) format.strike = true
    const color = attribute(child(properties, 'w:color'), 'w:val')
    if (color && /^[0-9a-f]{6}$/i.test(color)) format.color = `#${color}`
    const size = Number(attribute(child(properties, 'w:sz'), 'w:val'))
    if (size > 0) format.sizeHalfPoints = size
    const vertical = attribute(child(properties, 'w:vertAlign'), 'w:val')
    if (vertical === 'superscript' || vertical === 'subscript') format.vertAlign = vertical
    if (on(child(properties, 'w:vanish'))) return
    if (link) format.link = link
    for (const item of run.children || []) {
      if (typeof item === 'string') continue
      if (item.name === 'w:t') output.push({ text: textOf(item), ...format })
      else if (item.name === 'w:tab') output.push({ break: 'tab' })
      else if (item.name === 'w:br' || item.name === 'w:cr') output.push({ break: attribute(item, 'w:type') === 'page' ? 'page' : 'line' })
      else if (item.name === 'w:pict') {
        const data = findElement(item, 'w:binData')
        if (data) {
          const bytes = Buffer.from(textOf(data).replace(/\s+/g, ''), 'base64')
          if (sniffImage(bytes)) output.push({ image: { data: bytes } })
        }
      } else if (item.name === 'w:sym') {
        const code = Number.parseInt(attribute(item, 'w:char') || '', 16)
        if (Number.isFinite(code)) output.push({ text: code >= 0xf000 && code <= 0xf0ff ? '•' : String.fromCodePoint(code), ...format })
      }
    }
  }

  function paragraphFrom(element) {
    const properties = child(element, 'w:pPr')
    const block = { type: 'paragraph', runs: [] }
    const style = attribute(child(properties, 'w:pStyle'), 'w:val') || ''
    const heading = /^heading\s*([1-6])$/i.exec(style) || /^Heading([1-6])$/.exec(style)
    if (heading) block.heading = Number(heading[1])
    const align = attribute(child(properties, 'w:jc'), 'w:val')
    if (align === 'center' || align === 'right') block.align = align
    if (align === 'both' || align === 'distribute') block.align = 'justify'
    const list = child(properties, 'w:listPr')
    if (list) {
      const level = Number(attribute(child(list, 'w:ilvl'), 'w:val')) || 0
      const marker = attribute(child(list, 'wx:t'), 'wx:val') || ''
      block.list = { id: `w2003-${attribute(child(list, 'w:ilfo'), 'w:val') || ++listCounter}`, level: Math.min(8, level), ordered: /\d|^[a-z]{1,3}[.)]/i.test(marker) }
    }
    const visit = (nodes, link) => {
      for (const item of nodes || []) {
        if (typeof item === 'string') continue
        if (item.name === 'w:r') runFrom(item, link, block.runs)
        else if (item.name === 'w:hlink') {
          const destination = attribute(item, 'w:dest') || (attribute(item, 'w:bookmark') ? `#${attribute(item, 'w:bookmark')}` : null)
          visit(item.children, destination && /^(https?:|mailto:|#)/i.test(destination) ? destination : link)
        } else if (['w:fldSimple', 'w:smartTag', 'w:customXml', 'w:ins', 'aml:annotation', 'aml:content', 'w:sdt', 'w:sdtContent'].includes(item.name)) visit(item.children, link)
      }
    }
    visit(element.children, null)
    return block
  }

  function tableFrom(element) {
    const rows = []
    for (const row of element.children || []) {
      if (typeof row === 'string' || row.name !== 'w:tr') continue
      const cells = []
      for (const cell of row.children || []) {
        if (typeof cell === 'string' || cell.name !== 'w:tc') continue
        const cellBlocks = []
        visitBlocks(cell.children, cellBlocks)
        const span = Number(attribute(child(child(cell, 'w:tcPr'), 'w:gridSpan'), 'w:val')) || 1
        cells.push({ blocks: cellBlocks.length ? cellBlocks : [{ type: 'paragraph', runs: [] }], ...(span > 1 ? { colSpan: span } : {}) })
      }
      if (cells.length) rows.push({ cells })
    }
    return rows.length ? { type: 'table', rows } : null
  }

  function visitBlocks(nodes, output) {
    for (const item of nodes || []) {
      if (typeof item === 'string') continue
      if (item.name === 'w:p') output.push(paragraphFrom(item))
      else if (item.name === 'w:tbl') { const table = tableFrom(item); if (table) output.push(table) }
      else if (item.name === 'wx:sect' || item.name === 'wx:sub-section' || item.name === 'w:sdt' || item.name === 'w:sdtContent' || item.name === 'aml:annotation' || item.name === 'aml:content') visitBlocks(item.children, output)
    }
  }
  visitBlocks(body?.children, blocks)
  const titleElement = findElement(findElement(document, 'o:DocumentProperties') || { children: [] }, 'o:Title')
  return { title: titleElement ? textOf(titleElement).trim() : '', blocks, warnings: [] }
}

/** Identify markup content regardless of the file extension. */
function sniffMarkup(input) {
  const bytes = Buffer.isBuffer(input) ? input : Buffer.from(input || [])
  const head = decodeMarkupBytes(bytes.subarray(0, 8192)).replace(/^﻿/, '').trimStart()
  const lower = head.slice(0, 4096).toLowerCase()
  if (/^mime-version:/i.test(head) || (/^(?:from|subject|date|message-id|content-type):/im.test(head.slice(0, 1024)) && /multipart\/related/i.test(head))) return 'mht'
  if (lower.startsWith('<?xml') && /<w:worddocument\b/.test(lower)) return 'word2003'
  if (lower.startsWith('<?xml') && /<html\b/.test(lower)) return 'html'
  if (/^<!doctype\s+html/.test(lower) || /^<html\b/.test(lower) || /^(?:<!--[\s\S]*?-->\s*)*<(?:html|head|body|meta|title|p|div|table|h[1-6])\b/.test(lower)) return 'html'
  return null
}

module.exports = { decodeMarkupBytes, htmlToFlow, mhtToFlow, parseMarkup, sniffMarkup, word2003ToFlow }
