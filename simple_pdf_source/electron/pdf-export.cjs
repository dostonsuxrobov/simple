'use strict'

// Writers for Export As > Word, web page, Markdown and plain text. Pages come
// from the renderer (src/lib/pdfExport.ts): each has its plain `text` and,
// from the layout model, `blocks` (headings, paragraphs of styled runs,
// bullets, simple tables and pictures). Pages without blocks fall back to
// their text, keeping its line breaks.

const JSZip = require('jszip')

const TEXT_EXPORT_FORMATS = new Set(['docx', 'txt', 'md', 'html'])
const IMAGE_EXPORT_FORMATS = new Set(['png', 'jpeg', 'webp'])
const MAX_TEXT_LENGTH = 50_000_000
const MAX_IMAGE_BYTES = 200 * 1024 * 1024
const TWIPS_PER_POINT = 20
const EMU_PER_POINT = 12700

function safeExportBaseName(value) {
  return String(value || 'Untitled')
    // Callers normally pass the already extensionless PDF name. Only remove an
    // actual PDF suffix here so meaningful dotted suffixes such as
    // "Quarterly.v2" are not mistaken for a file extension a second time.
    .replace(/\.pdf$/i, '')
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim() || 'Untitled'
}

function stripInvalidXmlCharacters(value) {
  // XML 1.0 permits tabs and line breaks, but not the remaining C0 controls or
  // the two BMP noncharacters. PDF text streams occasionally expose those
  // bytes; removing them here keeps DOCX and HTML outputs parseable.
  return String(value ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
}

function escapeXml(value) {
  return stripInvalidXmlCharacters(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

// ---------------------------------------------------------------------------
// Validation: the renderer's model is checked and bounded before writing.
// ---------------------------------------------------------------------------

function toBuffer(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data)
  return null
}

function imageMime(bytes) {
  if (bytes.length > 8 && bytes.readUInt32BE(0) === 0x89504e47 && bytes.readUInt32BE(4) === 0x0d0a1a0a) return 'image/png'
  if (bytes.length > 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg'
  return null
}

function finiteBetween(value, minimum, maximum) {
  const number = Number(value)
  return Number.isFinite(number) && number >= minimum && number <= maximum ? number : undefined
}

function normalizeRun(raw, budget) {
  const text = stripInvalidXmlCharacters(raw?.text).replace(/\r\n?|\n/g, ' ')
  budget.text += text.length
  if (budget.text > MAX_TEXT_LENGTH) throw new Error('The extracted text is too large to export safely in one file.')
  const run = { text }
  if (raw?.bold) run.bold = true
  if (raw?.italic) run.italic = true
  const size = finiteBetween(raw?.size, 1, 1000)
  if (size) run.size = size
  const font = String(raw?.font || '').replace(/[^\p{L}\p{N} ._-]/gu, '').trim().slice(0, 64)
  if (font) run.font = font
  if (raw?.superscript) run.superscript = true
  return run
}

function normalizeLines(raw, budget, limit) {
  if (!Array.isArray(raw)) return []
  return raw.slice(0, limit).map((line) => {
    const runs = (Array.isArray(line?.runs) ? line.runs : []).slice(0, 2000).map((run) => normalizeRun(run, budget)).filter((run) => run.text)
    return line?.hardBreak ? { runs, hardBreak: true } : { runs }
  }).filter((line) => line.runs.length)
}

function normalizeBlock(raw, budget) {
  const align = raw?.align === 'center' || raw?.align === 'right' ? { align: raw.align } : {}
  switch (raw?.type) {
    case 'heading': {
      const lines = normalizeLines(raw.lines, budget, 50)
      return lines.length ? { type: 'heading', level: Math.max(1, Math.min(3, Math.trunc(Number(raw.level)) || 1)), lines, ...align } : null
    }
    case 'paragraph': {
      const lines = normalizeLines(raw.lines, budget, 20_000)
      return lines.length ? { type: 'paragraph', lines, ...align, ...(raw.list === 'bullet' ? { list: 'bullet' } : {}) } : null
    }
    case 'table': {
      const rows = (Array.isArray(raw.rows) ? raw.rows : []).slice(0, 10_000)
        .map((row) => (Array.isArray(row) ? row : []).slice(0, 64).map((cell) => normalizeLines(cell, budget, 2000)))
        .filter((row) => row.length)
      if (!rows.length) return null
      const columns = Math.max(...rows.map((row) => row.length))
      const widths = Array.from({ length: columns }, (_, index) => finiteBetween(raw.widths?.[index], 1, 14_400) || 72)
      return { type: 'table', rows: rows.map((row) => [...row, ...Array.from({ length: columns - row.length }, () => [])]), widths }
    }
    case 'image': {
      const data = toBuffer(raw.data)
      const mime = data && imageMime(data)
      if (!mime) return null
      budget.images += data.length
      if (budget.images > MAX_IMAGE_BYTES) return null
      return { type: 'image', data, mime, width: finiteBetween(raw.width, 1, 14_400) || 144, height: finiteBetween(raw.height, 1, 14_400) || 144 }
    }
    default:
      return null
  }
}

function normalizeTextPages(input) {
  if (!Array.isArray(input) || input.length < 1 || input.length > 10_000) {
    throw new Error('Choose between 1 and 10,000 pages to export.')
  }
  const budget = { text: 0, images: 0 }
  return input.map((page, index) => {
    const pageNumber = Number(page?.pageNumber)
    const text = stripInvalidXmlCharacters(page?.text).replace(/\r\n?/g, '\n')
    if (!Number.isInteger(pageNumber) || pageNumber < 1 || pageNumber > 1_000_000) {
      throw new Error(`Page ${index + 1} has an invalid page number.`)
    }
    budget.text += text.length
    if (budget.text > MAX_TEXT_LENGTH) throw new Error('The extracted text is too large to export safely in one file.')
    const result = { pageNumber, text }
    const width = finiteBetween(page?.width, 1, 14_400)
    const height = finiteBetween(page?.height, 1, 14_400)
    if (width && height) {
      result.width = width
      result.height = height
    }
    if (Array.isArray(page?.blocks)) result.blocks = page.blocks.slice(0, 100_000).map((block) => normalizeBlock(block, budget)).filter(Boolean)
    return result
  })
}

/** Blocks of a page; a page without a model becomes one paragraph per text paragraph. */
function pageBlocks(page) {
  if (page.blocks) return page.blocks
  return page.text.split(/\n{2,}/).map((paragraph) => paragraph.replace(/^\n+|\n+$/g, '')).filter(Boolean)
    .map((paragraph) => ({ type: 'paragraph', lines: paragraph.split('\n').map((line) => ({ runs: [{ text: line }], hardBreak: true })) }))
}

function lineString(line) {
  return line.runs.map((run) => run.text).join('')
}

/** Lines of a paragraph joined into flowing text: a space, or nothing after a hyphenated break. */
function joinsWithoutSpace(previous, next) {
  return /[\p{L}]-$/u.test(previous) && /^\p{Ll}/u.test(next)
}

function sameFormat(left, right) {
  return Boolean(left.bold) === Boolean(right.bold) && Boolean(left.italic) === Boolean(right.italic)
    && left.size === right.size && left.font === right.font && Boolean(left.superscript) === Boolean(right.superscript)
}

/**
 * The runs of a paragraph's lines as one sequence: lines that flow are joined
 * by a space (folded into the run before it), forced breaks become
 * `{break: true}`, and neighbouring runs with the same format are merged.
 * @param {{runs: object[], hardBreak?: boolean}[]} lines
 * @param {{breaks?: boolean, plain?: (run: object) => object}} [options]
 */
function flatRuns(lines, { breaks = true, plain = (run) => run } = {}) {
  const items = []
  // Spaces join a neighbouring run, but never a superscript one.
  const add = (run) => {
    const last = items[items.length - 1]
    const open = last && !last.break
    if (open && (sameFormat(last, run) || (!run.text.trim() && !last.superscript))) last.text += run.text
    else if (open && !last.text.trim() && !last.superscript && !run.superscript) items[items.length - 1] = { ...run, text: last.text + run.text }
    else items.push({ ...run })
  }
  lines.forEach((line, index) => {
    for (const run of line.runs) add(plain(run))
    const next = lines[index + 1]
    if (!next) return
    if (breaks && line.hardBreak) items.push({ break: true })
    else if (!/\s$/.test(lineString(line)) && !joinsWithoutSpace(lineString(line), lineString(next))) {
      add({ ...plain(line.runs[line.runs.length - 1]), superscript: false, text: ' ' })
    }
  })
  return items
}

function bodyStyle(pages) {
  const sizes = new Map()
  const fonts = new Map()
  for (const page of pages) {
    for (const block of pageBlocks(page)) {
      if (block.type !== 'paragraph') continue
      for (const line of block.lines) {
        for (const run of line.runs) {
          const weight = run.text.trim().length
          if (run.size) sizes.set(run.size, (sizes.get(run.size) || 0) + weight)
          if (run.font) fonts.set(run.font, (fonts.get(run.font) || 0) + weight)
        }
      }
    }
  }
  const most = (map, fallback) => [...map].sort((left, right) => right[1] - left[1])[0]?.[0] ?? fallback
  return { size: most(sizes, 11), font: most(fonts, '') }
}

// ---------------------------------------------------------------------------
// Plain text: the PDF's own lines, a blank line between paragraphs, a form
// feed between pages.
// ---------------------------------------------------------------------------

function blockText(block) {
  if (block.type === 'image') return ''
  if (block.type === 'heading') return block.lines.map(lineString).join(' ').replace(/ {2,}/g, ' ')
  if (block.type === 'table') return block.rows.map((row) => row.map((cell) => cell.map(lineString).join(' ')).join('\t')).join('\n')
  const lines = block.lines.map(lineString)
  if (block.list === 'bullet') return lines.map((line, index) => `${index ? '  ' : '• '}${line}`).join('\n')
  return lines.join('\n')
}

function plainTextDocument(pages) {
  return `${pages.map((page) => (page.blocks ? page.blocks.map(blockText).filter(Boolean).join('\n\n') : page.text.trim())).join('\n\f\n')}\n`
}

// ---------------------------------------------------------------------------
// Markdown
// ---------------------------------------------------------------------------

function escapeMarkdown(text) {
  return text.replace(/[\\`*_[\]<>|]/g, '\\$&')
}

function markdownRuns(runs) {
  const merged = []
  for (const run of runs) {
    const last = merged[merged.length - 1]
    if (last && ((Boolean(last.bold) === Boolean(run.bold) && Boolean(last.italic) === Boolean(run.italic) && Boolean(last.superscript) === Boolean(run.superscript)) || !run.text.trim())) last.text += run.text
    else if (last && !last.text.trim()) merged[merged.length - 1] = { ...run, text: last.text + run.text }
    else merged.push({ ...run })
  }
  return merged.map((run) => {
    const match = /^(\s*)([\s\S]*?)(\s*)$/.exec(run.text.replace(/\t/g, ' '))
    let body = escapeMarkdown(match[2])
    if (!body) return run.text.replace(/\t/g, ' ')
    if (run.superscript) body = `<sup>${body}</sup>`
    const marker = run.bold && run.italic ? '***' : run.bold ? '**' : run.italic ? '*' : ''
    return `${match[1]}${marker}${body}${marker}${match[3]}`
  }).join('')
}

function markdownLineStart(text) {
  return text.replace(/^(\s*)(#{1,6}\s|[-+*]\s|\d{1,9}[.)]\s|>|=+\s*$|-{3,}\s*$)/, (_match, space, marker) => `${space}\\${marker}`)
}

/** A paragraph that runs on from the previous page joins it (Markdown has no pages). */
function flowAcrossPages(pages) {
  const blocks = []
  pages.forEach((page, pageIndex) => {
    pageBlocks(page).forEach((block, blockIndex) => {
      const previous = blocks[blocks.length - 1]
      const lastLine = previous?.type === 'paragraph' && !previous.list ? lineString(previous.lines[previous.lines.length - 1]) : ''
      const firstLine = block.type === 'paragraph' && !block.list ? lineString(block.lines[0]) : ''
      if (pageIndex && !blockIndex && lastLine && firstLine && !/[.!?:;"”’)\]]$/.test(lastLine.trim()) && /^[\p{Ll}\p{N}(,;]/u.test(firstLine.trim())) {
        blocks[blocks.length - 1] = { ...previous, lines: [...previous.lines, ...block.lines] }
      } else blocks.push(block)
    })
  })
  return blocks
}

function markdownTable(block) {
  const cell = (lines) => markdownRuns(flatRuns(lines, { breaks: false })).replace(/\n/g, ' ') || ' '
  const rows = block.rows.map((row) => `| ${row.map(cell).join(' | ')} |`)
  const separator = `| ${block.rows[0].map(() => '---').join(' | ')} |`
  return [rows[0], separator, ...rows.slice(1)].join('\n')
}

/**
 * @param {object[]} pages
 * @param {{imageFolder?: string}} options pictures are written to imageFolder
 *   (relative to the Markdown file) when given, else embedded as data URLs
 * @returns {{text: string, assets: {name: string, data: Buffer}[]}}
 */
function markdownDocument(pages, options = {}) {
  const assets = []
  const parts = flowAcrossPages(pages).map((block) => {
    if (block.type === 'image') {
      const extension = block.mime === 'image/png' ? 'png' : 'jpg'
      const name = `image-${String(assets.length + 1).padStart(3, '0')}.${extension}`
      if (options.imageFolder) {
        assets.push({ name, data: block.data })
        return `![](<${options.imageFolder}/${name}>)`
      }
      return `![](data:${block.mime};base64,${block.data.toString('base64')})`
    }
    if (block.type === 'table') return markdownTable(block)
    if (block.type === 'heading') return `${'#'.repeat(block.level)} ${markdownRuns(flatRuns(block.lines, { breaks: false, plain: (run) => ({ ...run, bold: false }) }))}`
    const lines = block.lines.map((line, index) => {
      const text = markdownLineStart(markdownRuns(line.runs))
      const end = line.hardBreak && index < block.lines.length - 1 ? '\\' : ''
      return `${block.list === 'bullet' ? (index ? '  ' : '- ') : ''}${text}${end}`
    })
    return lines.join('\n')
  })
  return { text: `${parts.filter(Boolean).join('\n\n')}\n`, assets }
}

// ---------------------------------------------------------------------------
// Web page
// ---------------------------------------------------------------------------

function cssFont(font) {
  return `"${String(font).replace(/["\\]/g, '')}"`
}

function htmlRuns(runs, body) {
  return runs.map((run) => {
    let html = escapeXml(run.text).replace(/\t/g, '&emsp;')
    const style = []
    if (run.size && Math.abs(run.size - body.size) >= 0.5 && !run.superscript) style.push(`font-size:${run.size}pt`)
    if (run.font && run.font !== body.font) style.push(`font-family:${escapeXml(cssFont(run.font))},inherit`)
    if (style.length) html = `<span style="${style.join(';')}">${html}</span>`
    if (run.superscript) html = `<sup>${html}</sup>`
    if (run.italic) html = `<em>${html}</em>`
    if (run.bold) html = `<strong>${html}</strong>`
    return html
  }).join('')
}

function htmlLines(lines, body, options) {
  return flatRuns(lines, options).map((item) => (item.break ? '<br>' : htmlRuns([item], body))).join('')
}

function htmlBlocks(blocks, body) {
  const output = []
  let list = false
  for (const block of blocks) {
    if (list && !(block.type === 'paragraph' && block.list)) { output.push('</ul>'); list = false }
    const align = block.align ? ` style="text-align:${block.align}"` : ''
    if (block.type === 'image') {
      output.push(`<p class="picture"><img src="data:${block.mime};base64,${block.data.toString('base64')}" style="width:${Math.round(block.width * 100) / 100}pt" alt=""></p>`)
    } else if (block.type === 'table') {
      const columns = block.widths.map((width) => `<col style="width:${Math.round(width)}pt">`).join('')
      const rows = block.rows.map((row) => `<tr>${row.map((cell) => `<td>${htmlLines(cell, body, { breaks: false }) || '&nbsp;'}</td>`).join('')}</tr>`).join('')
      output.push(`<table><colgroup>${columns}</colgroup>${rows}</table>`)
    } else if (block.type === 'heading') {
      // A heading's level sets its size and weight.
      output.push(`<h${block.level}${align}>${htmlLines(block.lines, body, { breaks: false, plain: (run) => ({ ...run, bold: false, size: undefined }) })}</h${block.level}>`)
    } else if (block.list === 'bullet') {
      if (!list) { output.push('<ul>'); list = true }
      output.push(`<li>${htmlLines(block.lines, body)}</li>`)
    } else output.push(`<p${align}>${htmlLines(block.lines, body)}</p>`)
  }
  if (list) output.push('</ul>')
  return output.join('\n')
}

function htmlDocument(pages, title) {
  const documentTitle = escapeXml(String(title || 'Exported PDF'))
  const body = bodyStyle(pages)
  const sections = pages.map((page) => {
    const width = page.width ? ` style="max-width:${Math.round(page.width)}pt"` : ''
    return `<section class="page" data-page="${page.pageNumber}"${width}>\n${htmlBlocks(pageBlocks(page), body)}\n</section>`
  }).join('\n')
  // Inside <style> text is not entity-decoded; font names are already reduced to letters, digits and spaces.
  const fontFamily = `${body.font ? `${cssFont(body.font)}, ` : ''}"Segoe UI", Arial, sans-serif`
  return `<!doctype html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${documentTitle}</title><style>
body{margin:0;padding:32px 16px;background:#eee;color:#18181b;font-family:${fontFamily};font-size:${body.size}pt;line-height:1.45}
.page{box-sizing:border-box;margin:0 auto 24px;padding:56px 64px;background:#fff;box-shadow:0 2px 14px #0002;overflow-wrap:anywhere}
.page:last-child{margin-bottom:0}
h1,h2,h3{line-height:1.25;margin:0.6em 0 0.35em}p{margin:0 0 0.7em}ul{margin:0 0 0.7em;padding-left:1.6em}
table{border-collapse:collapse;margin:0 0 0.9em;max-width:100%}td{border:1px solid #a1a1aa;padding:3px 6px;vertical-align:top}
img{max-width:100%;height:auto}p.picture{margin:0.4em 0 0.8em}
@media print{body{padding:0;background:#fff}.page{max-width:none!important;margin:0;padding:0;box-shadow:none;break-after:page}.page:last-child{break-after:auto}}
</style></head><body>
${sections}
</body></html>`
}

// ---------------------------------------------------------------------------
// Word
// ---------------------------------------------------------------------------

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

function halfPoints(size) {
  return Math.max(2, Math.round(size * 2))
}

function fontXml(font) {
  const name = escapeXml(font)
  return `<w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:eastAsia="${name}" w:cs="${name}"/>`
}

/** A run; `base` is the paragraph style's font, size and boldness. */
function runXml(run, base, text = run.text) {
  const properties = []
  if (run.font && run.font !== base.font) properties.push(fontXml(run.font))
  if (Boolean(run.bold) !== Boolean(base.bold)) properties.push(run.bold ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/><w:bCs w:val="0"/>')
  if (run.italic) properties.push('<w:i/><w:iCs/>')
  if (run.size && Math.abs(run.size - base.size) >= 0.5 && !run.superscript) properties.push(`<w:sz w:val="${halfPoints(run.size)}"/><w:szCs w:val="${halfPoints(run.size)}"/>`)
  if (run.superscript) properties.push('<w:vertAlign w:val="superscript"/>')
  const content = text.split('\t').map((part) => (part ? `<w:t xml:space="preserve">${escapeXml(part)}</w:t>` : '')).join('<w:tab/>')
  return `<w:r>${properties.length ? `<w:rPr>${properties.join('')}</w:rPr>` : ''}${content}</w:r>`
}

function linesXml(lines, base, options) {
  return flatRuns(lines, options).map((item) => (item.break ? '<w:r><w:br/></w:r>' : runXml(item, base))).join('')
}

function paragraphProperties(parts) {
  const xml = parts.filter(Boolean).join('')
  return xml ? `<w:pPr>${xml}</w:pPr>` : ''
}

function drawingXml(image, index, maxWidth, maxHeight) {
  const scale = Math.min(1, maxWidth / image.width, maxHeight / image.height)
  const cx = Math.max(1, Math.round(image.width * scale * EMU_PER_POINT))
  const cy = Math.max(1, Math.round(image.height * scale * EMU_PER_POINT))
  const id = index + 1
  return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
}

function tableXml(block, base, maxWidth) {
  const total = block.widths.reduce((sum, width) => sum + width, 0)
  const scale = total > maxWidth ? maxWidth / total : 1
  const widths = block.widths.map((width) => Math.max(240, Math.round(width * scale * TWIPS_PER_POINT)))
  const grid = widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')
  const rows = block.rows.map((row) => `<w:tr>${row.map((cell, column) => {
    const paragraphs = `<w:p>${linesXml(cell, base, { breaks: false })}</w:p>`
    return `<w:tc><w:tcPr><w:tcW w:w="${widths[column]}" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc>`
  }).join('')}</w:tr>`).join('')
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`
}

function sectionXml(page) {
  const width = Math.round((page.width || 612) * TWIPS_PER_POINT)
  const height = Math.round((page.height || 792) * TWIPS_PER_POINT)
  const margin = Math.min(width, height) < 9000 ? 720 : 1440
  return `<w:sectPr><w:pgSz w:w="${width}" w:h="${height}"${width > height ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${margin}" w:right="${margin}" w:bottom="${margin}" w:left="${margin}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
}

function headingSizes(pages, body) {
  const sizes = { 1: new Map(), 2: new Map(), 3: new Map() }
  for (const page of pages) {
    for (const block of pageBlocks(page)) {
      if (block.type !== 'heading') continue
      for (const line of block.lines) for (const run of line.runs) if (run.size) sizes[block.level].set(run.size, (sizes[block.level].get(run.size) || 0) + run.text.length)
    }
  }
  const fallback = { 1: Math.max(body.size * 1.8, 18), 2: Math.max(body.size * 1.45, 15), 3: Math.max(body.size * 1.2, 13) }
  return Object.fromEntries([1, 2, 3].map((level) => [level, [...sizes[level]].sort((left, right) => right[1] - left[1])[0]?.[0] || fallback[level]]))
}

function stylesXml(body, headings) {
  const font = body.font || 'Calibri'
  const heading = (level) => `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${level === 1 ? 240 : 160}" w:after="80"/><w:outlineLvl w:val="${level - 1}"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="${halfPoints(headings[level])}"/><w:szCs w:val="${halfPoints(headings[level])}"/></w:rPr></w:style>`
  const border = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>${fontXml(font)}<w:sz w:val="${halfPoints(body.size)}"/><w:szCs w:val="${halfPoints(body.size)}"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${heading(1)}${heading(2)}${heading(3)}<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders></w:tblPr></w:style></w:styles>`
}

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`

async function docxDocument(pages, title) {
  const body = bodyStyle(pages)
  const headings = headingSizes(pages, body)
  const images = []
  let bullets = false
  const parts = []
  pages.forEach((page, pageIndex) => {
    const next = pages[pageIndex + 1]
    const sameSize = (left, right) => Math.round(left.width || 612) === Math.round(right.width || 612) && Math.round(left.height || 792) === Math.round(right.height || 792)
    const margin = Math.min(page.width || 612, page.height || 792) < 450 ? 36 : 72
    const maxWidth = Math.max(72, (page.width || 612) - margin * 2)
    const maxHeight = Math.max(72, (page.height || 792) - margin * 2)
    const blocks = pageBlocks(page)
    blocks.forEach((block, blockIndex) => {
      const align = block.align ? `<w:jc w:val="${block.align}"/>` : ''
      if (block.type === 'image') {
        images.push(block)
        parts.push(drawingXml(block, images.length - 1, maxWidth, maxHeight))
      } else if (block.type === 'table') {
        parts.push(tableXml(block, { font: body.font, size: body.size, bold: false }, maxWidth))
        // Word needs a paragraph between two tables and after a page's last table.
        const following = blocks[blockIndex + 1]
        if (!following || following.type === 'table') parts.push('<w:p/>')
      } else if (block.type === 'heading') {
        const base = { font: body.font, size: headings[block.level], bold: true }
        parts.push(`<w:p>${paragraphProperties([`<w:pStyle w:val="Heading${block.level}"/>`, align])}${linesXml(block.lines, base, { breaks: false })}</w:p>`)
      } else {
        const base = { font: body.font, size: body.size, bold: false }
        if (block.list === 'bullet') bullets = true
        const style = block.list === 'bullet' ? '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : ''
        parts.push(`<w:p>${paragraphProperties([style, align])}${linesXml(block.lines, base)}</w:p>`)
      }
    })
    if (!next) return
    // A page of a new size starts a new section; otherwise a page break.
    if (sameSize(page, next)) parts.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    else parts.push(`<w:p><w:pPr>${sectionXml(page)}</w:pPr></w:p>`)
  })
  parts.push(sectionXml(pages[pages.length - 1]))

  const zip = new JSZip()
  const hasPng = images.some((image) => image.mime === 'image/png')
  const hasJpeg = images.some((image) => image.mime === 'image/jpeg')
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${hasPng ? '<Default Extension="png" ContentType="image/png"/>' : ''}${hasJpeg ? '<Default Extension="jpeg" ContentType="image/jpeg"/>' : ''}
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${bullets ? '\n<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ''}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`)
  const relationships = ['<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>']
  if (bullets) relationships.push('<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>')
  images.forEach((image, index) => {
    const name = `image${index + 1}.${image.mime === 'image/png' ? 'png' : 'jpeg'}`
    zip.file(`word/media/${name}`, image.data, { compression: 'STORE' })
    relationships.push(`<Relationship Id="rIdImage${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`)
  })
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join('')}</Relationships>`)
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${parts.join('')}</w:body></w:document>`)
  zip.file('word/styles.xml', stylesXml(body, headings))
  if (bullets) zip.file('word/numbering.xml', NUMBERING_XML)
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${escapeXml(title || 'Exported PDF')}</dc:title><dc:creator>simple</dc:creator></cp:coreProperties>`)
  zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>simple</Application></Properties>`)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } })
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/**
 * The exported file and, for Markdown with pictures and an `imageFolder`,
 * the picture files to write into that folder next to it.
 * @param {'docx'|'txt'|'md'|'html'} format
 * @param {object[]} rawPages pages from the renderer
 * @param {string} title
 * @param {{imageFolder?: string}} [options]
 * @returns {Promise<{data: Buffer, assets: {name: string, data: Buffer}[]}>}
 */
async function buildTextExportFiles(format, rawPages, title, options = {}) {
  if (!TEXT_EXPORT_FORMATS.has(format)) throw new Error('This text export format is not supported.')
  const pages = normalizeTextPages(rawPages)
  if (format === 'docx') return { data: await docxDocument(pages, String(title || '')), assets: [] }
  if (format === 'md') {
    const markdown = markdownDocument(pages, options)
    return { data: Buffer.from(markdown.text, 'utf8'), assets: markdown.assets }
  }
  const value = format === 'html' ? htmlDocument(pages, title) : plainTextDocument(pages)
  return { data: Buffer.from(value, 'utf8'), assets: [] }
}

/** The exported file as one buffer (Markdown pictures are embedded). */
async function buildTextExport(format, rawPages, title) {
  return (await buildTextExportFiles(format, rawPages, title)).data
}

function imageExportFileName(baseName, pageNumber, pageCount, format) {
  if (!IMAGE_EXPORT_FORMATS.has(format)) throw new Error('This image export format is not supported.')
  const width = Math.max(3, String(Math.max(1, pageCount)).length)
  const extension = format === 'jpeg' ? 'jpg' : format
  return `${safeExportBaseName(baseName)} - page ${String(pageNumber).padStart(width, '0')}.${extension}`
}

module.exports = {
  IMAGE_EXPORT_FORMATS,
  TEXT_EXPORT_FORMATS,
  buildTextExport,
  buildTextExportFiles,
  imageExportFileName,
  normalizeTextPages,
  safeExportBaseName,
}
