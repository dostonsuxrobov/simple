'use strict'

// Native, dependency-free RTF reader. It turns RTF (including Word's own RTF
// and RTF files that carry a .doc name) into the flow document consumed by
// simple-docx.cjs: paragraphs, headings, lists, character formatting, links,
// tables, PNG/JPEG pictures, the default header/footer and footnote text.

const MAX_RTF_BYTES = 256 * 1024 * 1024
const MAX_GROUP_DEPTH = 512

const CHARSET_CODEPAGES = Object.freeze({
  0: 1252, 77: 10000, 128: 932, 129: 949, 130: 1361, 134: 936, 136: 950,
  161: 1253, 162: 1254, 163: 1258, 177: 1255, 178: 1256, 186: 1257, 204: 1251, 222: 874, 238: 1250,
})
const CODEPAGE_ENCODINGS = Object.freeze({
  437: 'ibm437', 850: 'ibm850', 866: 'ibm866', 874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5',
  1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254',
  1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257', 1258: 'windows-1258', 10000: 'macintosh', 65001: 'utf-8',
})
// Symbol and Wingdings bullets used by Word's list text.
const SYMBOL_FONT = /^(?:symbol|wingdings\s*\d?|webdings)$/i
const SPECIAL_CHARACTERS = Object.freeze({
  emdash: '—', endash: '–', bullet: '•', lquote: '‘', rquote: '’', ldblquote: '“', rdblquote: '”',
  emspace: ' ', enspace: ' ', qmspace: ' ', zwj: '‍', zwnj: '‌', zwbo: '​', ltrmark: '‎', rtlmark: '‏',
})
// Destinations whose text never belongs to the document body.
const SKIPPED_DESTINATIONS = new Set([
  'filetbl', 'revtbl', 'rsidtbl', 'listtable', 'listoverridetable', 'generator', 'xmlnstbl', 'themedata', 'colorschememapping',
  'latentstyles', 'datastore', 'defchp', 'defpap', 'pgdsctbl', 'nonshppict', 'headerl', 'headerf', 'footerl', 'footerf',
  'bkmkstart', 'bkmkend', 'objdata', 'xe', 'tc', 'pn', 'atnid', 'atnauthor', 'annotation', 'atrfstart', 'atrfend', 'protusertbl',
  'docvar', 'userprops', 'mmathPr', 'pgptbl', 'nesttableprops', 'template', 'fchars', 'lchars', 'ftnsep', 'ftnsepc', 'ftncn',
  'aftnsep', 'aftnsepc', 'aftncn', 'wgrffmtfilter', 'passwordhash', 'background', 'author', 'operator', 'company', 'subject',
  'keywords', 'doccomm', 'comment', 'creatim', 'revtim', 'printim', 'buptim', 'category', 'manager', 'hlinkbase',
])
// "\*" destinations that Simple understands instead of skipping.
const KNOWN_STARRED = new Set(['shppict', 'fldinst', 'listtext', 'pntext'])
const TEXT_FREE_DESTINATIONS = new Set(['info', 'pict'])

function decodeCodepage(bytes, codepage) {
  const label = CODEPAGE_ENCODINGS[codepage] || 'windows-1252'
  try { return new TextDecoder(label).decode(Uint8Array.from(bytes)) } catch { return new TextDecoder('windows-1252').decode(Uint8Array.from(bytes)) }
}

function symbolCharacter(byte) {
  // Word lists commonly encode bullets as Symbol/Wingdings code points.
  if ([0xb7, 0xa7, 0x6c, 0xfc, 0xd8, 0x6e, 0x71, 0xa8].includes(byte)) return '•'
  if (byte === 0x6f) return '◦'
  return String.fromCharCode(byte)
}

function isOrderedMarker(marker) {
  return /^[(\[]?(?:\d+(?:\.\d+)*|[a-z]{1,3}|[ivxlcdm]{1,6})[.)\]]?$/i.test(marker) && /[\d.)\]]/.test(marker)
}

function initialCharacter(font = null) {
  return { bold: false, italic: false, underline: false, strike: false, sizeHalfPoints: 24, color: null, highlight: null, font, vertAlign: null, caps: false, smallCaps: false, hidden: false }
}

function initialParagraph() {
  return { align: 'left', indentLeftTw: 0, indentRightTw: 0, indentFirstTw: 0, spaceBeforeTw: undefined, spaceAfterTw: undefined, style: 0, outline: null, inTable: false, list: null, level: 0, pageBreakBefore: false, rtl: false }
}

function copyParagraph(paragraph) {
  return { ...paragraph, list: paragraph.list ? { ...paragraph.list } : null }
}

function createSink() {
  return { blocks: [], runs: [], paragraph: null, rows: [], cells: [], cellBlocks: [], cellWidths: [], shading: [], pendingShading: null, listMarkerText: undefined }
}

function sameFormatting(left, right) {
  for (const key of ['bold', 'italic', 'underline', 'strike', 'sizeHalfPoints', 'color', 'highlight', 'font', 'vertAlign', 'caps', 'smallCaps', 'link']) {
    if ((left[key] ?? null) !== (right[key] ?? null)) return false
  }
  return true
}

/**
 * Parse RTF bytes or text. Returns { title, blocks, header, footer, page,
 * warnings, footnoteCount } for buildFlowDocx().
 */
function parseRtf(input) {
  const source = Buffer.isBuffer(input) ? input.toString('latin1') : input instanceof Uint8Array ? Buffer.from(input).toString('latin1') : String(input || '')
  if (source.length > MAX_RTF_BYTES) throw new Error('This RTF document is too large to open safely.')
  const start = source.search(/\{\s*\\rtf/)
  if (start < 0) throw new Error('This file is not a Rich Text Format document.')

  const fonts = new Map()
  const colors = []
  const styles = new Map()
  const body = createSink()
  const notes = createSink()
  const bands = { header: null, footer: null }
  const page = { margins: {} }
  let defaultFont = null
  let documentCodepage = 1252
  let title = ''
  let footnoteCount = 0
  let unsupportedPictures = 0

  let state = {
    destination: 'body', skip: false, sink: body, character: initialCharacter(), paragraph: initialParagraph(),
    unicodeSkip: 1, link: null, field: null, picture: null, fontEntry: null, styleEntry: null, color: null, listText: null,
  }
  const stack = []
  let pendingBytes = []
  let pendingSkip = 0

  const fontName = (index) => (fonts.get(index ?? defaultFont)?.name || '').replace(/;.*$/s, '').trim() || null
  const fontCodepage = (index) => fonts.get(index ?? defaultFont)?.codepage || documentCodepage

  function colorAt(index) {
    const color = colors[index]
    if (!color) return null
    return `#${[color.red, color.green, color.blue].map((value) => Math.max(0, Math.min(255, value)).toString(16).padStart(2, '0')).join('')}`
  }

  function emitText(text) {
    if (!text || state.skip) return
    switch (state.destination) {
      case 'fonttbl': if (state.fontEntry) state.fontEntry.name += text; return
      case 'stylesheet': if (state.styleEntry) state.styleEntry.name += text; return
      case 'colortbl':
        for (const character of text) {
          if (character !== ';') continue
          colors.push(state.color?.set ? state.color : null)
          state.color = { red: 0, green: 0, blue: 0, set: false }
        }
        return
      case 'fldinst': if (state.field) state.field.instruction += text; return
      case 'title': title += text; return
      case 'listtext': if (state.listText) state.listText.value += text; return
      default: if (TEXT_FREE_DESTINATIONS.has(state.destination)) return
    }
    if (state.character.hidden) return
    const sink = state.sink
    if (!sink.paragraph) sink.paragraph = copyParagraph(state.paragraph)
    const character = state.character
    const run = {
      text,
      bold: character.bold || undefined,
      italic: character.italic || undefined,
      underline: character.underline || undefined,
      strike: character.strike || undefined,
      // RTF's default is 12 pt; the DOCX default is 11 pt, so always write it.
      sizeHalfPoints: character.sizeHalfPoints,
      color: character.color || undefined,
      highlight: character.highlight || undefined,
      font: fontName(character.font) || undefined,
      vertAlign: character.vertAlign || undefined,
      caps: character.caps || undefined,
      smallCaps: character.smallCaps || undefined,
      link: state.link || undefined,
    }
    const previous = sink.runs.at(-1)
    if (previous && typeof previous.text === 'string' && !previous.break && !previous.image && sameFormatting(previous, run)) previous.text += text
    else sink.runs.push(run)
  }

  function flushBytes() {
    if (!pendingBytes.length) return
    const bytes = pendingBytes
    pendingBytes = []
    if (SYMBOL_FONT.test(fontName(state.character.font) || '')) emitText(bytes.map(symbolCharacter).join(''))
    else emitText(decodeCodepage(bytes, fontCodepage(state.character.font)))
  }

  function pushRun(run) {
    flushBytes()
    if (state.skip || state.destination !== 'body' || state.character.hidden) return
    const sink = state.sink
    if (!sink.paragraph) sink.paragraph = copyParagraph(state.paragraph)
    sink.runs.push(run)
  }

  function finishParagraph(sink = state.sink, options = {}) {
    flushBytes()
    const properties = sink.paragraph || copyParagraph(state.paragraph)
    const runs = sink.runs
    const marker = sink.listMarkerText
    sink.runs = []
    sink.paragraph = null
    sink.listMarkerText = undefined
    const block = { type: 'paragraph', runs }
    if (properties.align !== 'left') block.align = properties.align
    if (properties.rtl) block.rtl = true
    const heading = /^heading\s*([1-6])$/i.exec((styles.get(properties.style) || '').trim())
    if (heading) block.heading = Number(heading[1])
    else if (Number.isInteger(properties.outline) && properties.outline >= 0 && properties.outline <= 5) block.heading = properties.outline + 1
    if (properties.list || marker) {
      block.list = {
        id: properties.list?.id || 'pn',
        level: properties.list ? properties.level || 0 : 0,
        ordered: marker ? isOrderedMarker(marker) : false,
      }
    } else {
      if (properties.indentLeftTw) block.indentLeftTw = properties.indentLeftTw
      if (properties.indentFirstTw) block.indentFirstTw = properties.indentFirstTw
    }
    if (properties.indentRightTw) block.indentRightTw = properties.indentRightTw
    if (properties.spaceBeforeTw !== undefined) block.spaceBeforeTw = properties.spaceBeforeTw
    if (properties.spaceAfterTw !== undefined) block.spaceAfterTw = properties.spaceAfterTw
    if (properties.pageBreakBefore) block.pageBreakBefore = true
    if (properties.inTable || options.inTable) {
      sink.cellBlocks.push(block)
      return
    }
    flushTable(sink)
    sink.blocks.push(block)
  }

  function finishCell(sink = state.sink) {
    flushBytes()
    if (sink.runs.length || sink.paragraph || !sink.cellBlocks.length) finishParagraph(sink, { inTable: true })
    const blocks = sink.cellBlocks
    sink.cellBlocks = []
    sink.cells.push({ blocks })
  }

  function finishRow(sink = state.sink) {
    flushBytes()
    if (sink.runs.length || sink.cellBlocks.length) finishCell(sink)
    if (!sink.cells.length) return
    // Word may define a row's cells and shading before or after its content.
    sink.cells.forEach((cell, index) => { if (sink.shading[index]) cell.shading = sink.shading[index] })
    sink.rows.push({ cells: sink.cells, widths: sink.cellWidths.slice() })
    sink.cells = []
  }

  function flushTable(sink) {
    if (sink.cellBlocks.length || sink.cells.length) finishRow(sink)
    if (!sink.rows.length) return
    const rows = sink.rows
    sink.rows = []
    const first = rows[0]
    let widthsTw
    if (first.widths.length && first.widths.length === first.cells.length && rows.every((row) => row.cells.length === first.cells.length)) {
      widthsTw = first.widths.map((right, index) => right - (index ? first.widths[index - 1] : 0))
      if (widthsTw.some((width) => width <= 0)) widthsTw = undefined
    }
    sink.blocks.push({ type: 'table', rows: rows.map((row) => ({ cells: row.cells })), ...(widthsTw ? { widthsTw } : {}) })
  }

  function finishPicture(picture) {
    if (!picture || state.skip) return
    let bytes = picture.binary
    if (!bytes && picture.hex.length) {
      const hex = picture.hex.join('')
      bytes = Buffer.from(hex.length % 2 ? hex.slice(0, -1) : hex, 'hex')
    }
    if (!bytes?.length) return
    const isPng = bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e
    const isJpeg = bytes[0] === 0xff && bytes[1] === 0xd8
    if (!isPng && !isJpeg) {
      unsupportedPictures += 1
      return
    }
    const pixels = (twips, scale) => twips > 0 ? Math.round(twips / 15 * (scale || 100) / 100) : 0
    const widthPx = pixels(picture.widthGoal, picture.scaleX)
    const heightPx = pixels(picture.heightGoal, picture.scaleY)
    if (state.destination !== 'body') return
    const sink = state.sink
    if (!sink.paragraph) sink.paragraph = copyParagraph(state.paragraph)
    sink.runs.push({ image: { data: bytes, ...(widthPx ? { widthPx } : {}), ...(heightPx ? { heightPx } : {}) } })
  }

  function enterBand(kind) {
    if (bands[kind]) { state.skip = true; return }
    bands[kind] = createSink()
    state.sink = bands[kind]
    state.paragraph = initialParagraph()
  }

  function controlWord(word, parameter) {
    if (SKIPPED_DESTINATIONS.has(word)) { state.skip = true; return }
    switch (word) {
      case 'ansicpg': documentCodepage = parameter || 1252; return
      case 'mac': documentCodepage = 10000; return
      case 'pc': documentCodepage = 437; return
      case 'pca': documentCodepage = 850; return
      case 'deff': defaultFont = parameter ?? 0; state.character.font ??= defaultFont; return
      case 'fonttbl': state.destination = 'fonttbl'; return
      case 'colortbl': state.destination = 'colortbl'; state.color = { red: 0, green: 0, blue: 0, set: false }; return
      case 'stylesheet': state.destination = 'stylesheet'; return
      case 'info': state.destination = 'info'; return
      case 'title': if (state.destination === 'info') state.destination = 'title'; else state.skip = true; return
      case 'shppict': return
      case 'object': return
      case 'pict':
        state.destination = 'pict'
        state.picture = { hex: [], binary: null, widthGoal: 0, heightGoal: 0, scaleX: 100, scaleY: 100 }
        return
      case 'picwgoal': if (state.picture) state.picture.widthGoal = parameter || 0; return
      case 'pichgoal': if (state.picture) state.picture.heightGoal = parameter || 0; return
      case 'picscalex': if (state.picture) state.picture.scaleX = parameter || 100; return
      case 'picscaley': if (state.picture) state.picture.scaleY = parameter || 100; return
      case 'field': state.field = { instruction: '' }; return
      case 'fldinst': state.destination = 'fldinst'; return
      case 'fldrslt': {
        const instruction = state.field?.instruction || ''
        const hyperlink = /HYPERLINK\s+(\\l\s+)?"([^"]+)"/i.exec(instruction) || /HYPERLINK\s+(\\l\s+)?(\S+)/i.exec(instruction)
        if (hyperlink) state.link = hyperlink[1] ? `#${hyperlink[2]}` : hyperlink[2]
        return
      }
      case 'listtext': case 'pntext':
        state.destination = 'listtext'
        state.listText = { value: '' }
        return
      case 'header': case 'headerr': enterBand('header'); return
      case 'footer': case 'footerr': enterBand('footer'); return
      case 'footnote':
        state.sink = notes
        state.paragraph = initialParagraph()
        notes.pendingNumber = ++footnoteCount
        return
      case 'chftn':
        pushRun({ text: String(state.sink === notes ? notes.pendingNumber : footnoteCount + 1), vertAlign: 'superscript' })
        return
      default: break
    }

    if (state.destination === 'fonttbl') {
      if (word === 'f') { state.fontEntry = { name: '', codepage: null }; fonts.set(parameter ?? 0, state.fontEntry) }
      else if (word === 'fcharset' && state.fontEntry) state.fontEntry.codepage = CHARSET_CODEPAGES[parameter] ?? state.fontEntry.codepage
      else if (word === 'cpg' && state.fontEntry) state.fontEntry.codepage = parameter
      return
    }
    if (state.destination === 'colortbl') {
      if (word === 'red' || word === 'green' || word === 'blue') { state.color[word] = parameter || 0; state.color.set = true }
      return
    }
    if (state.destination === 'stylesheet') {
      if (word === 's') state.styleEntry = { index: parameter ?? 0, name: '' }
      else if (word === 'cs' || word === 'ds' || word === 'ts') state.styleEntry = null
      return
    }
    if (state.destination !== 'body' && state.destination !== 'listtext') return

    const character = state.character
    const paragraph = state.paragraph
    const on = parameter === null || parameter !== 0
    switch (word) {
      case 'par': finishParagraph(); return
      case 'line': pushRun({ break: 'line' }); return
      case 'tab': pushRun({ break: 'tab' }); return
      case 'page': pushRun({ break: 'page' }); return
      case 'sect': if (state.sink.runs.length) finishParagraph(); return
      case 'cell': finishCell(); return
      case 'row': finishRow(); return
      case 'nestcell': case 'nestrow': finishParagraph(); return
      case 'u':
        if (parameter === null) return
        emitText(String.fromCharCode(parameter < 0 ? parameter + 65536 : parameter))
        pendingSkip = state.unicodeSkip
        return
      case 'uc': state.unicodeSkip = Math.max(0, parameter ?? 1); return
      case 'pard': state.paragraph = initialParagraph(); return
      case 'ql': paragraph.align = 'left'; return
      case 'qc': paragraph.align = 'center'; return
      case 'qr': paragraph.align = 'right'; return
      case 'qj': case 'qd': paragraph.align = 'justify'; return
      case 'li': paragraph.indentLeftTw = parameter || 0; return
      case 'ri': paragraph.indentRightTw = parameter || 0; return
      case 'fi': paragraph.indentFirstTw = parameter || 0; return
      case 'sb': paragraph.spaceBeforeTw = Math.max(0, parameter || 0); return
      case 'sa': paragraph.spaceAfterTw = Math.max(0, parameter || 0); return
      case 's': paragraph.style = parameter ?? 0; return
      case 'outlinelevel': paragraph.outline = parameter; return
      case 'intbl': paragraph.inTable = true; return
      case 'itap': paragraph.inTable = (parameter ?? 1) > 0; return
      case 'pagebb': paragraph.pageBreakBefore = true; return
      case 'rtlpar': paragraph.rtl = true; return
      case 'ltrpar': paragraph.rtl = false; return
      case 'ls': paragraph.list = { id: `ls${parameter ?? 0}` }; return
      case 'ilvl': paragraph.level = Math.max(0, Math.min(8, parameter || 0)); return
      case 'trowd': state.sink.cellWidths = []; state.sink.shading = []; state.sink.pendingShading = null; return
      case 'cellx': state.sink.cellWidths.push(parameter || 0); state.sink.shading.push(state.sink.pendingShading); state.sink.pendingShading = null; return
      case 'clcbpat': state.sink.pendingShading = colorAt(parameter); return
      case 'plain': state.character = initialCharacter(defaultFont); return
      case 'b': character.bold = on; return
      case 'i': character.italic = on; return
      case 'ul': case 'uld': case 'uldash': case 'uldashd': case 'uldashdd': case 'uldb': case 'ulth': case 'ulw': case 'ulwave': case 'ulhwave': case 'ululdbwave':
        character.underline = on; return
      case 'ulnone': character.underline = false; return
      case 'strike': case 'striked': character.strike = on; return
      case 'caps': character.caps = on; return
      case 'scaps': character.smallCaps = on; return
      case 'v': character.hidden = on; return
      case 'fs': character.sizeHalfPoints = parameter > 0 ? parameter : 24; return
      case 'f': character.font = parameter ?? 0; return
      case 'cf': character.color = colorAt(parameter); return
      case 'highlight': case 'cb': case 'chcbpat': character.highlight = colorAt(parameter); return
      case 'super': character.vertAlign = 'superscript'; return
      case 'sub': character.vertAlign = 'subscript'; return
      case 'nosupersub': character.vertAlign = null; return
      case 'paperw': case 'pgwsxn': page.widthTw ??= parameter; return
      case 'paperh': case 'pghsxn': page.heightTw ??= parameter; return
      case 'margl': case 'marglsxn': page.margins.left ??= parameter; return
      case 'margr': case 'margrsxn': page.margins.right ??= parameter; return
      case 'margt': case 'margtsxn': page.margins.top ??= parameter; return
      case 'margb': case 'margbsxn': page.margins.bottom ??= parameter; return
      default: break
    }
    if (Object.hasOwn(SPECIAL_CHARACTERS, word)) emitText(SPECIAL_CHARACTERS[word])
  }

  for (let index = start; index < source.length; index += 1) {
    const character = source[index]
    if (character === '{') {
      flushBytes()
      if (stack.length >= MAX_GROUP_DEPTH) throw new Error('This RTF document is nested too deeply to open safely.')
      stack.push(state)
      state = { ...state, character: { ...state.character }, paragraph: copyParagraph(state.paragraph) }
      pendingSkip = 0
      continue
    }
    if (character === '}') {
      flushBytes()
      const closing = state
      if (!stack.length) break
      state = stack.pop()
      pendingSkip = 0
      if (closing.skip) continue
      if (closing.destination === 'pict' && closing.picture && state.destination !== 'pict') finishPicture(closing.picture)
      if (closing.destination === 'stylesheet' && closing.styleEntry) styles.set(closing.styleEntry.index, closing.styleEntry.name.replace(/;.*$/s, '').trim())
      if (closing.destination === 'listtext' && closing.listText && state.destination !== 'listtext') {
        state.sink.listMarkerText = closing.listText.value.replace(/\s+/g, '')
      }
      if (closing.sink !== state.sink && closing.sink !== body) {
        // Leaving a header, footer or footnote: close what it left open.
        if (closing.sink.runs.length || closing.sink.paragraph) finishParagraph(closing.sink)
        flushTable(closing.sink)
      }
      continue
    }
    if (character === '\\') {
      const next = source[index + 1]
      if (next === undefined) break
      if (next === '\\' || next === '{' || next === '}') {
        index += 1
        if (pendingSkip > 0) { pendingSkip -= 1; continue }
        flushBytes()
        emitText(next)
        continue
      }
      if (next === "'") {
        const hex = source.slice(index + 2, index + 4)
        index += 3
        if (pendingSkip > 0) { pendingSkip -= 1; continue }
        if (/^[0-9a-f]{2}$/i.test(hex) && !state.skip && state.destination !== 'pict') pendingBytes.push(Number.parseInt(hex, 16))
        continue
      }
      if (next === '*') {
        index += 1
        const destination = /^\s*\\([a-zA-Z]{1,32})/.exec(source.slice(index + 1, index + 48))
        if (!destination || !KNOWN_STARRED.has(destination[1])) state.skip = true
        continue
      }
      if (next === '~') { index += 1; flushBytes(); emitText(' '); continue }
      if (next === '_') { index += 1; flushBytes(); emitText('‑'); continue }
      if (next === '\r' || next === '\n') { index += 1; if (!state.skip && state.destination === 'body') finishParagraph(); continue }
      const match = /^([a-zA-Z]{1,32})(-?\d{1,10})? ?/.exec(source.slice(index + 1, index + 48))
      if (!match) { index += 1; continue }
      index += match[0].length
      const word = match[1]
      const parameter = match[2] === undefined ? null : Number(match[2])
      if (word === 'bin') {
        const length = Math.max(0, Math.min(parameter || 0, source.length - index - 1))
        if (!state.skip && state.picture && state.destination === 'pict') state.picture.binary = Buffer.from(source.slice(index + 1, index + 1 + length), 'latin1')
        index += length
        continue
      }
      if (word !== 'u') pendingSkip = 0
      flushBytes()
      if (!state.skip) controlWord(word, parameter)
      continue
    }
    if (character === '\r' || character === '\n') continue
    if (state.skip) continue
    if (state.destination === 'pict') {
      const run = /^[0-9a-fA-F\s]+/.exec(source.slice(index, index + 65536))
      if (run && state.picture) {
        state.picture.hex.push(run[0].replace(/\s+/g, ''))
        index += run[0].length - 1
      }
      continue
    }
    if (pendingSkip > 0) { pendingSkip -= 1; continue }
    const code = character.charCodeAt(0)
    if (code >= 0x80) { pendingBytes.push(code); continue }
    if (code < 0x20 && code !== 0x09) continue
    if (code === 0x09) { pushRun({ break: 'tab' }); continue }
    flushBytes()
    // Ordinary printable ASCII runs are taken in one step.
    const plain = /^[\x20-\x5b\x5d-\x7a\x7c\x7e\x7f]+/.exec(source.slice(index, index + 65536))
    emitText(plain[0])
    index += plain[0].length - 1
  }
  flushBytes()
  for (const sink of [body, bands.header, bands.footer, notes]) {
    if (!sink) continue
    if (sink.runs.length || sink.paragraph) finishParagraph(sink)
    flushTable(sink)
  }
  if (notes.blocks.length) {
    body.blocks.push({ type: 'paragraph', runs: [{ text: 'Notes', bold: true }] })
    body.blocks.push(...notes.blocks)
  }
  const warnings = []
  if (unsupportedPictures) {
    warnings.push(unsupportedPictures === 1
      ? 'One picture uses a format (such as WMF or EMF) that Simple cannot show, so it was left out.'
      : `${unsupportedPictures} pictures use formats (such as WMF or EMF) that Simple cannot show, so they were left out.`)
  }
  const hasPage = page.widthTw || page.heightTw || Object.keys(page.margins).length
  return {
    title: title.trim(),
    blocks: body.blocks,
    ...(bands.header?.blocks?.length ? { header: bands.header.blocks } : {}),
    ...(bands.footer?.blocks?.length ? { footer: bands.footer.blocks } : {}),
    ...(hasPage ? { page } : {}),
    footnoteCount,
    warnings,
  }
}

module.exports = { parseRtf }
