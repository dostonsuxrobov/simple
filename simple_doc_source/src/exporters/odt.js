// Native OpenDocument Text (.odt) writer for the WordCanvas document model.
// It needs no Office engine: paragraphs and headings, character formatting,
// real lists, tables with spans/shading/borders, PNG/JPEG/GIF/SVG pictures,
// links, bookmarks, footnotes/endnotes, and one master page per section with
// its page size, margins and header/footer. Pass a document prepared by
// prepareDocumentForExport() so pictures are embedded data: URIs.

import JSZip from "jszip";
import { createListState, headingLevel, mathText, noteReference, safeLink, sectionsOf, storyHasContent } from "../document-export.js";

export const ODT_MIME_TYPE = "application/vnd.oasis.opendocument.text";
const PX_PER_IN = 96;
const DEFAULT_BORDER = "0.5pt solid #bfbfbf";
const IMAGE_TYPES = { "image/png": "png", "image/jpeg": "jpg", "image/jpg": "jpg", "image/gif": "gif", "image/svg+xml": "svg" };
const NUMBER_FORMATS = { decimal: "1", lowerLetter: "a", upperLetter: "A", lowerRoman: "i", upperRoman: "I" };

const inches = (px) => `${(Math.max(0, Number(px) || 0) / PX_PER_IN).toFixed(4)}in`;
const points = (px) => `${((Number(px) || 0) * 0.75).toFixed(2)}pt`;

function isXmlCharacter(code) {
  return code === 0x9 || code === 0xa || code === 0xd || (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd) || (code >= 0x10000 && code <= 0x10ffff);
}

export function xmlEscape(value) {
  let safe = "";
  for (const character of String(value ?? "")) if (isXmlCharacter(character.codePointAt(0))) safe += character;
  return safe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function hexColor(value) {
  const color = String(value || "").trim().toLowerCase();
  let match = /^#([0-9a-f]{6})$/.exec(color);
  if (match) return `#${match[1]}`;
  match = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(color);
  if (match) return `#${match[1]}${match[1]}${match[2]}${match[2]}${match[3]}${match[3]}`;
  match = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/.exec(color);
  if (match) return `#${match.slice(1, 4).map((part) => Math.min(255, Number(part)).toString(16).padStart(2, "0")).join("")}`;
  return null;
}

function familyName(value) {
  const family = String(value || "").split(",")[0].replace(/["']/g, "").trim();
  return family && family.length <= 64 ? family : null;
}

function decodeDataImage(source) {
  const match = /^data:(image\/[a-z0-9.+-]+);base64,([a-z0-9+/=\s]+)$/i.exec(String(source || ""));
  if (!match) return null;
  const type = match[1].toLowerCase();
  if (!IMAGE_TYPES[type]) return null;
  const binary = atob(match[2].replace(/\s/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return { bytes, type: type === "image/jpg" ? "image/jpeg" : type, extension: IMAGE_TYPES[type] };
}

/** Deduplicating automatic-style registry. */
function styleRegistry(prefix) {
  const byKey = new Map();
  return {
    name(key, build) {
      if (!byKey.has(key)) byKey.set(key, { name: `${prefix}${byKey.size + 1}`, xml: null, build });
      return byKey.get(key).name;
    },
    xml() {
      return [...byKey.values()].map((entry) => entry.build(entry.name)).join("");
    },
  };
}

function textProperties(style) {
  const properties = [];
  const family = familyName(style?.fontFamily);
  if (family) properties.push(`style:font-name="${xmlEscape(family)}" fo:font-family="${xmlEscape(family)}"`);
  const size = Number(style?.fontSizePx);
  if (Number.isFinite(size) && size > 0) properties.push(`fo:font-size="${points(size)}" style:font-size-asian="${points(size)}" style:font-size-complex="${points(size)}"`);
  if (style?.bold) properties.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
  if (style?.italic) properties.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
  if (style?.underline) properties.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
  if (style?.strikethrough) properties.push('style:text-line-through-style="solid" style:text-line-through-type="single"');
  const color = hexColor(style?.color);
  if (color) properties.push(`fo:color="${color}"`);
  const highlight = hexColor(style?.highlightColor);
  if (highlight) properties.push(`fo:background-color="${highlight}"`);
  if (style?.verticalAlign === "super") properties.push('style:text-position="super 58%"');
  if (style?.verticalAlign === "sub") properties.push('style:text-position="sub 58%"');
  if (style?.smallCaps) properties.push('fo:font-variant="small-caps"');
  else if (style?.caps) properties.push('fo:text-transform="uppercase"');
  if (style?.hidden) properties.push('text:display="none"');
  const spacing = Number(style?.letterSpacingPx);
  if (Number.isFinite(spacing) && spacing) properties.push(`fo:letter-spacing="${points(spacing)}"`);
  return properties.join(" ");
}

function createContext(document, options) {
  const fonts = new Set(["Calibri"]);
  const media = [];
  const mediaBySource = new Map();
  const context = {
    document,
    warnings: options.warnings || [],
    counts: { images: 0, objects: 0 },
    lists: createListState(document),
    textStyles: styleRegistry("T"),
    paragraphStyles: styleRegistry("P"),
    tableStyles: styleRegistry("Tbl"),
    cellStyles: styleRegistry("Cell"),
    columnStyles: styleRegistry("Col"),
    graphicStyles: styleRegistry("fr"),
    fonts,
    media,
    noteCount: 0,
    listSeen: new Set(),
    drawCount: 0,
    tableCount: 0,
    addMedia(source) {
      if (mediaBySource.has(source)) return mediaBySource.get(source);
      const decoded = decodeDataImage(source);
      if (!decoded) return null;
      const entry = { path: `Pictures/image${media.length + 1}.${decoded.extension}`, ...decoded };
      media.push(entry);
      mediaBySource.set(source, entry);
      return entry;
    },
  };
  return context;
}

function textStyleName(style, context) {
  const properties = textProperties(style);
  if (!properties) return null;
  const family = familyName(style?.fontFamily);
  if (family) context.fonts.add(family);
  return context.textStyles.name(properties, (name) => `<style:style style:name="${name}" style:family="text"><style:text-properties ${properties}/></style:style>`);
}

function spanXml(text, style, context, options = {}) {
  if (!text) return "";
  const name = textStyleName(style, context);
  let content = "";
  // Spaces, tabs and line breaks need explicit ODF elements.
  for (const piece of String(text).split(/(\t|\n|\v| {2,})/)) {
    if (!piece) continue;
    if (piece === "\t") content += "<text:tab/>";
    else if (piece === "\n" || piece === "\v") content += "<text:line-break/>";
    else if (/^ {2,}$/.test(piece)) content += ` <text:s text:c="${piece.length - 1}"/>`;
    else if (options.inBand) {
      content += piece.split(/(\{page\}|\{pages\})/i).map((part) => {
        if (/^\{page\}$/i.test(part)) return '<text:page-number text:select-page="current">1</text:page-number>';
        if (/^\{pages\}$/i.test(part)) return "<text:page-count>1</text:page-count>";
        return xmlEscape(part);
      }).join("");
    } else content += xmlEscape(piece);
  }
  return name ? `<text:span text:style-name="${name}">${content}</text:span>` : content;
}

function noteXml(reference, context) {
  const blocks = context.document[reference.kind === "endnote" ? "endnotes" : "footnotes"]?.[reference.id] || [];
  const number = ++context.noteCount;
  const body = blocks.map((block) => blockXml(block, context, { inNote: true })).join("") || '<text:p text:style-name="Footnote"/>';
  return `<text:note text:id="note${number}" text:note-class="${reference.kind === "endnote" ? "endnote" : "footnote"}"><text:note-citation>${number}</text:note-citation><text:note-body>${body}</text:note-body></text:note>`;
}

function bookmarksAt(context, blockId) {
  const marks = [];
  for (const [name, range] of Object.entries(context.document.bookmarks || {})) {
    if (!/^[\w.-]{1,200}$/.test(name)) continue;
    if (range?.start?.blockId === blockId) marks.push({ offset: Math.max(0, Number(range.start.offset) || 0), xml: `<text:bookmark-start text:name="${xmlEscape(name)}"/>`, start: true });
    if (range?.end?.blockId === blockId) marks.push({ offset: Math.max(0, Number(range.end.offset) || 0), xml: `<text:bookmark-end text:name="${xmlEscape(name)}"/>` });
  }
  return marks.sort((left, right) => left.offset - right.offset || Number(right.start || false) - Number(left.start || false));
}

function inlineXml(paragraph, context, options = {}) {
  const marks = bookmarksAt(context, paragraph.id);
  const pieces = [];
  let offset = 0;
  let mark = 0;
  for (const run of paragraph.runs || []) {
    const text = String(run?.text || "");
    while (mark < marks.length && marks[mark].offset <= offset) pieces.push({ raw: marks[mark++].xml });
    const reference = noteReference(run);
    if (reference) {
      // Notes inside notes are not allowed in ODF; keep the marker text instead.
      pieces.push(options.inNote ? { run, text } : { raw: noteXml(reference, context) });
    } else if (run?.style?.equation) {
      pieces.push({ run, text: mathText(run.style.equation.root) });
    } else {
      let start = 0;
      while (mark < marks.length && marks[mark].offset < offset + text.length) {
        const at = marks[mark].offset - offset;
        if (at > start) pieces.push({ run, text: text.slice(start, at) });
        pieces.push({ raw: marks[mark++].xml });
        start = Math.max(start, at);
      }
      if (start < text.length) pieces.push({ run, text: text.slice(start) });
    }
    offset += text.length;
  }
  while (mark < marks.length) pieces.push({ raw: marks[mark++].xml });
  let output = "";
  for (let index = 0; index < pieces.length;) {
    const piece = pieces[index];
    const link = piece.run ? safeLink(piece.run.style?.link) : null;
    if (!link) {
      output += piece.raw ?? spanXml(piece.text, piece.run.style, context, options);
      index += 1;
      continue;
    }
    let inner = "";
    while (index < pieces.length && pieces[index].run && safeLink(pieces[index].run.style?.link) === link) {
      inner += spanXml(pieces[index].text, pieces[index].run.style, context, options);
      index += 1;
    }
    output += `<text:a xlink:type="simple" xlink:href="${xmlEscape(link)}">${inner}</text:a>`;
  }
  return output;
}

function paragraphStyleName(paragraph, context, options = {}) {
  const style = paragraph.style || {};
  const properties = [];
  const align = { center: "center", right: "end", justify: "justify", left: "start" }[style.align];
  if (align) properties.push(`fo:text-align="${align}"`);
  if (style.direction === "rtl") properties.push('style:writing-mode="rl-tb"');
  if (!options.inList) {
    if (style.indentLeftPx) properties.push(`fo:margin-left="${inches(style.indentLeftPx)}"`);
    if (style.indentFirstLinePx) properties.push(`fo:text-indent="${(Number(style.indentFirstLinePx) / PX_PER_IN).toFixed(4)}in"`);
  }
  if (style.indentRightPx) properties.push(`fo:margin-right="${inches(style.indentRightPx)}"`);
  properties.push(`fo:margin-top="${inches(style.spaceBeforePx || 0)}" fo:margin-bottom="${inches(style.spaceAfterPx || 0)}"`);
  if (style.lineRule === "exact" && Number(style.lineHeightPx) > 0) properties.push(`fo:line-height="${points(style.lineHeightPx)}"`);
  else if (style.lineRule === "atLeast" && Number(style.lineHeightPx) > 0) properties.push(`style:line-height-at-least="${points(style.lineHeightPx)}"`);
  else if (Number(style.lineHeight) > 0) properties.push(`fo:line-height="${Math.round(Number(style.lineHeight) * 100)}%"`);
  if (style.keepWithNext) properties.push('fo:keep-with-next="always"');
  if (style.keepLinesTogether) properties.push('fo:keep-together="always"');
  if (style.pageBreakBefore && !options.masterPage) properties.push('fo:break-before="page"');
  const shading = hexColor(style.shading);
  if (shading) properties.push(`fo:background-color="${shading}"`);
  const tabs = (style.tabStops || []).map((stop) => `<style:tab-stop style:position="${inches(stop.posPx)}"${stop.align && stop.align !== "left" ? ` style:type="${stop.align === "decimal" ? "char" : stop.align}"${stop.align === "decimal" ? ' style:char="."' : ""}` : ""}${stop.leader && stop.leader !== "none" ? ` style:leader-style="${{ dot: "dotted", dash: "dash", underscore: "solid" }[stop.leader] || "dotted"}" style:leader-text="${{ dot: ".", dash: "-", underscore: "_" }[stop.leader] || "."}"` : ""}/>`).join("");
  const level = headingLevel(paragraph);
  const parent = options.parent || (level ? `Heading_20_${level}` : "Standard");
  const master = options.masterPage ? ` style:master-page-name="${options.masterPage}"` : "";
  const key = `${parent}|${master}|${properties.join(" ")}|${tabs}`;
  return context.paragraphStyles.name(key, (name) => `<style:style style:name="${name}" style:family="paragraph" style:parent-style-name="${parent}"${master}><style:paragraph-properties ${properties.join(" ")}>${tabs ? `<style:tab-stops>${tabs}</style:tab-stops>` : ""}</style:paragraph-properties></style:style>`);
}

function paragraphXml(paragraph, context, options = {}) {
  const level = headingLevel(paragraph);
  const name = paragraphStyleName(paragraph, context, options);
  const content = inlineXml(paragraph, context, options);
  if (level && !options.inNote) return `<text:h text:style-name="${name}" text:outline-level="${level}">${content}</text:h>`;
  return `<text:p text:style-name="${name}">${content}</text:p>`;
}

function imageXml(block, context, options = {}) {
  const media = context.addMedia(block.src);
  const alignStyle = { center: "center", right: "end" }[block.align] || "start";
  const paragraphStyle = paragraphStyleName({ style: { align: { center: "center", right: "right" }[block.align] || "left" } }, context, options);
  if (!media) {
    context.counts.images += 1;
    return `<text:p text:style-name="${paragraphStyle}"><text:span>${xmlEscape("[Picture not available in this export]")}</text:span></text:p>`;
  }
  const graphic = context.graphicStyles.name(`inline-${alignStyle}`, (name) => `<style:style style:name="${name}" style:family="graphic" style:parent-style-name="Graphics"><style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline" style:horizontal-pos="center" style:horizontal-rel="paragraph" fo:border="none" style:wrap="none"/></style:style>`);
  const index = ++context.drawCount;
  const width = inches(Math.max(1, Number(block.widthPx) || 320));
  const height = inches(Math.max(1, Number(block.heightPx) || 240));
  return `<text:p text:style-name="${paragraphStyle}"><draw:frame draw:style-name="${graphic}" draw:name="Image${index}" text:anchor-type="as-char" svg:width="${width}" svg:height="${height}" draw:z-index="${index}"><draw:image xlink:href="${media.path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="${media.type}"/><svg:desc>Document image</svg:desc></draw:frame></text:p>`;
}

function borderValue(border) {
  if (!border || border.style === "none") return "none";
  const color = hexColor(border.color) || "#000000";
  const width = `${Math.max(0.25, (Number(border.widthPx) || 1) * 0.75).toFixed(2)}pt`;
  const style = { double: "double", dashed: "dashed", dotted: "dotted" }[border.style] || "solid";
  return `${width} ${style} ${color}`;
}

function tableXml(table, context, options = {}) {
  const rows = table.rows || [];
  const columns = Math.max(1, ...rows.map((row) => (row.cells || []).reduce((sum, cell) => sum + Math.max(1, Number(cell.colSpan) || 1), 0)));
  const widthPx = Math.max(96, options.contentWidthPx || 624);
  const fractions = Array.isArray(table.colFractions) && table.colFractions.length === columns ? table.colFractions : Array.from({ length: columns }, () => 1 / columns);
  const total = fractions.reduce((sum, value) => sum + Math.max(0.01, Number(value) || 0), 0);
  const tableName = `Table${++context.tableCount}`;
  const master = options.masterPage ? ` style:master-page-name="${options.masterPage}"` : "";
  const tableStyle = context.tableStyles.name(`table-${widthPx}-${master}`, (name) => `<style:style style:name="${name}" style:family="table"${master}><style:table-properties style:width="${inches(widthPx)}" table:align="margins"/></style:style>`);
  const columnXml = fractions.map((fraction) => {
    const width = inches(widthPx * Math.max(0.01, Number(fraction) || 0) / total);
    const name = context.columnStyles.name(`column-${width}`, (style) => `<style:style style:name="${style}" style:family="table-column"><style:table-column-properties style:column-width="${width}"/></style:style>`);
    return `<table:table-column table:style-name="${name}"/>`;
  }).join("");
  const pending = new Array(columns).fill(0);
  const rowXml = [];
  const headerRows = [];
  rows.forEach((row, rowIndex) => {
    const cells = [];
    let column = 0;
    const emitCovered = () => {
      while (column < columns && pending[column] > 0) {
        pending[column] -= 1;
        cells.push("<table:covered-table-cell/>");
        column += 1;
      }
    };
    for (const cell of row.cells || []) {
      emitCovered();
      if (column >= columns) break;
      const span = Math.max(1, Math.min(columns - column, Number(cell.colSpan) || 1));
      const rowSpan = Math.max(1, Math.min(rows.length - rowIndex, Number(cell.rowSpan) || 1));
      const borders = ["top", "right", "bottom", "left"].map((side) => {
        const value = cell.borders?.[side] || table.defaultBorders?.[side];
        if (value) return `fo:border-${side}="${borderValue(value)}"`;
        return cell.borders || table.defaultBorders ? `fo:border-${side}="none"` : `fo:border-${side}="${DEFAULT_BORDER}"`;
      }).join(" ");
      const shading = hexColor(cell.shading);
      const margin = cell.margin ? ` fo:padding-top="${inches(cell.margin.top)}" fo:padding-right="${inches(cell.margin.right)}" fo:padding-bottom="${inches(cell.margin.bottom)}" fo:padding-left="${inches(cell.margin.left)}"` : ' fo:padding="0.0382in"';
      const cellStyle = context.cellStyles.name(`${borders}|${shading}|${margin}`, (name) => `<style:style style:name="${name}" style:family="table-cell"><style:table-cell-properties ${borders}${shading ? ` fo:background-color="${shading}"` : ""}${margin}/></style:style>`);
      const blocks = cell.blocks?.length ? cell.blocks : [{ kind: "paragraph", id: "", runs: [], style: {} }];
      const content = blocks.map((block) => blockXml(block, context, { ...options, masterPage: undefined, inTable: true, contentWidthPx: widthPx * span / columns })).join("");
      cells.push(`<table:table-cell table:style-name="${cellStyle}" office:value-type="string"${span > 1 ? ` table:number-columns-spanned="${span}"` : ""}${rowSpan > 1 ? ` table:number-rows-spanned="${rowSpan}"` : ""}>${content}</table:table-cell>`);
      for (let offset = 1; offset < span; offset += 1) cells.push("<table:covered-table-cell/>");
      if (rowSpan > 1) for (let offset = 0; offset < span; offset += 1) pending[column + offset] = rowSpan - 1;
      column += span;
    }
    emitCovered();
    while (column < columns) { cells.push('<table:table-cell office:value-type="string"><text:p/></table:table-cell>'); column += 1; }
    const xml = `<table:table-row>${cells.join("")}</table:table-row>`;
    if (row.props?.repeatHeader && rowXml.length === 0) headerRows.push(xml);
    else rowXml.push(xml);
  });
  const header = headerRows.length ? `<table:table-header-rows>${headerRows.join("")}</table:table-header-rows>` : "";
  return `<table:table table:name="${tableName}" table:style-name="${tableStyle}">${columnXml}${header}${rowXml.join("")}</table:table>`;
}

function blockXml(block, context, options = {}) {
  if (block.kind === "paragraph") return paragraphXml(block, context, options);
  if (block.kind === "image") return imageXml(block, context, options);
  if (block.kind === "table") return tableXml(block, context, options);
  if (block.kind === "equation") {
    return paragraphXml({ id: block.id, style: { align: block.align || "center" }, runs: [{ text: mathText(block.equation?.root) || "[Equation]", style: { italic: true, fontFamily: "Cambria Math" } }] }, context, options);
  }
  if (block.kind === "shape") {
    const blocks = block.text?.blocks?.length ? block.text.blocks : [{ kind: "paragraph", id: block.id, style: {}, runs: [{ text: "[Shape]", style: { italic: true } }] }];
    return blocks.map((child) => blockXml(child, context, options)).join("");
  }
  context.counts.objects += 1;
  return paragraphXml({ id: block.id, style: {}, runs: [{ text: "[Unsupported document object]", style: { italic: true } }] }, context, options);
}

/** Body blocks with consecutive list paragraphs nested as ODF lists. */
function flowXml(blocks, context, options = {}) {
  const output = [];
  let index = 0;
  while (index < blocks.length) {
    const block = blocks[index];
    const list = block.kind === "paragraph" ? block.style?.list : null;
    if (!list || !context.document.lists?.[list.listId]) {
      output.push(blockXml(block, context, { ...options, masterPage: index === 0 ? options.masterPage : undefined }));
      index += 1;
      continue;
    }
    const run = [];
    while (index < blocks.length && blocks[index].kind === "paragraph" && blocks[index].style?.list?.listId === list.listId) run.push(blocks[index++]);
    output.push(listXml(run, list.listId, context, { ...options, masterPage: output.length === 0 ? options.masterPage : undefined }));
  }
  return output.join("");
}

function listXml(paragraphs, listId, context, options) {
  const styleName = `L${Object.keys(context.document.lists).indexOf(listId) + 1}`;
  const continued = context.listSeen.has(listId);
  context.listSeen.add(listId);
  let xml = `<text:list text:style-name="${styleName}"${continued ? ' text:continue-numbering="true"' : ""}>`;
  // depth = nesting of the innermost open <text:list>; itemOpen = whether a
  // <text:list-item> is open inside it.
  let depth = 0;
  let itemOpen = false;
  paragraphs.forEach((paragraph, index) => {
    const level = Math.max(0, Math.min(8, Number(paragraph.style.list.level) || 0));
    if (level > depth) {
      // Deeper levels nest inside the current item (or an empty container item).
      if (!itemOpen) xml += "<text:list-item>";
      while (depth < level) {
        xml += "<text:list>";
        depth += 1;
        if (depth < level) xml += "<text:list-item>";
      }
      itemOpen = false;
    } else {
      while (depth > level) {
        if (itemOpen) xml += "</text:list-item>";
        xml += "</text:list>";
        depth -= 1;
        itemOpen = true;
      }
      if (itemOpen) xml += "</text:list-item>";
    }
    xml += `<text:list-item>${paragraphXml(paragraph, context, { ...options, inList: true, masterPage: index === 0 ? options.masterPage : undefined })}`;
    itemOpen = true;
  });
  while (depth > 0) {
    if (itemOpen) xml += "</text:list-item>";
    xml += "</text:list>";
    depth -= 1;
    itemOpen = true;
  }
  if (itemOpen) xml += "</text:list-item>";
  return `${xml}</text:list>`;
}

function listStylesXml(document) {
  return Object.values(document.lists || {}).map((list, index) => {
    const levels = Array.from({ length: 10 }, (_value, levelIndex) => {
      const level = list.levels?.[Math.min(levelIndex, 8)] || { format: "bullet", bulletChar: "•", indentLeftPx: 24 * (levelIndex + 1), hangingPx: 18 };
      const left = inches(level.indentLeftPx ?? 24 * (levelIndex + 1));
      const hanging = (Number(level.hangingPx ?? 18) / PX_PER_IN).toFixed(4);
      const alignment = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${left}" fo:text-indent="-${hanging}in" fo:margin-left="${left}"/></style:list-level-properties>`;
      if (level.format === "bullet") return `<text:list-level-style-bullet text:level="${levelIndex + 1}" text:bullet-char="${xmlEscape(level.bulletChar || "•")}">${alignment}</text:list-level-style-bullet>`;
      const pattern = level.text || `%${levelIndex + 1}.`;
      const own = `%${levelIndex + 1}`;
      const at = pattern.lastIndexOf(own);
      const prefix = at >= 0 ? pattern.slice(0, at).replace(/%[1-9]/g, "").replace(/[.]$/, "") : "";
      const suffix = at >= 0 ? pattern.slice(at + own.length) : ".";
      const display = Math.max(1, (pattern.match(/%[1-9]/g) || []).length);
      return `<text:list-level-style-number text:level="${levelIndex + 1}" style:num-format="${NUMBER_FORMATS[level.format] || "1"}"${prefix && display === 1 ? ` style:num-prefix="${xmlEscape(prefix)}"` : ""} style:num-suffix="${xmlEscape(suffix)}" text:display-levels="${Math.min(display, levelIndex + 1)}" text:start-value="${Math.max(1, Math.floor(Number(level.start) || 1))}">${alignment}</text:list-level-style-number>`;
    }).join("");
    return `<text:list-style style:name="L${index + 1}">${levels}</text:list-style>`;
  }).join("");
}

function pageLayoutXml(name, props, hasHeader, hasFooter) {
  const margin = props.marginPx || { top: 96, right: 96, bottom: 96, left: 96 };
  const width = Number(props.pageWidthPx) || 816;
  const height = Number(props.pageHeightPx) || 1056;
  const headerDistance = Number(props.headerDistancePx ?? 48);
  const footerDistance = Number(props.footerDistancePx ?? 48);
  // ODF measures the header from the page margin; Word measures from the edge.
  const top = hasHeader ? Math.max(0, Math.min(margin.top, headerDistance)) : margin.top;
  const bottom = hasFooter ? Math.max(0, Math.min(margin.bottom, footerDistance)) : margin.bottom;
  const columns = props.columns?.count > 1 ? `<style:columns fo:column-count="${Math.min(45, Math.floor(props.columns.count))}" fo:column-gap="${inches(props.columns.gapPx || 48)}"/>` : "";
  const header = hasHeader ? `<style:header-style><style:header-footer-properties fo:min-height="0in" fo:margin-bottom="${inches(Math.max(0, margin.top - top - 24))}"/></style:header-style>` : "<style:header-style/>";
  const footer = hasFooter ? `<style:footer-style><style:header-footer-properties fo:min-height="0in" fo:margin-top="${inches(Math.max(0, margin.bottom - bottom - 24))}"/></style:footer-style>` : "<style:footer-style/>";
  return `<style:page-layout style:name="${name}"><style:page-layout-properties fo:page-width="${inches(width)}" fo:page-height="${inches(height)}" style:print-orientation="${width > height ? "landscape" : "portrait"}" fo:margin-top="${inches(top)}" fo:margin-bottom="${inches(bottom)}" fo:margin-left="${inches(margin.left)}" fo:margin-right="${inches(margin.right)}">${columns}</style:page-layout-properties>${header}${footer}</style:page-layout>`;
}

function bandXml(element, blocks, context, contentWidthPx) {
  if (!storyHasContent(blocks)) return "";
  return `<style:${element}>${flowXml(blocks, context, { inBand: true, contentWidthPx })}</style:${element}>`;
}

const NAMESPACES = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';

/** Build the ODT parts (strings and picture bytes) for a document model. */
export function documentToOdtParts(document, options = {}) {
  if (!document || !Array.isArray(document.blocks)) throw new Error("The open document is not ready to export.");
  const context = createContext(document, options);
  const sections = sectionsOf(document);
  const masters = [];
  const layouts = [];
  let body = "";
  sections.forEach(({ blocks, props }, index) => {
    const masterName = `Section${index + 1}`;
    const layoutName = `pm${index + 1}`;
    const margin = props.marginPx || { left: 96, right: 96 };
    const contentWidthPx = (Number(props.pageWidthPx) || 816) - (margin.left || 0) - (margin.right || 0);
    const header = bandXml("header", props.header, context, contentWidthPx);
    const footer = bandXml("footer", props.footer, context, contentWidthPx);
    const headerFirst = bandXml("header-first", props.headerFirst, context, contentWidthPx);
    const footerFirst = bandXml("footer-first", props.footerFirst, context, contentWidthPx);
    const headerEven = bandXml("header-left", props.headerEven, context, contentWidthPx);
    const footerEven = bandXml("footer-left", props.footerEven, context, contentWidthPx);
    layouts.push(pageLayoutXml(layoutName, props, Boolean(header || headerFirst || headerEven), Boolean(footer || footerFirst || footerEven)));
    masters.push(`<style:master-page style:name="${masterName}" style:page-layout-name="${layoutName}">${header}${headerEven}${headerFirst}${footer}${footerEven}${footerFirst}</style:master-page>`);
    const breakType = index ? (props.breakType ?? sections[index].breakType) : undefined;
    // A continuous section keeps the previous page; others start a new master page.
    const masterPage = index === 0 || breakType !== "continuous" ? masterName : undefined;
    const sectionBlocks = blocks.length ? blocks : [{ kind: "paragraph", id: "", runs: [], style: {} }];
    body += flowXml(sectionBlocks, context, { masterPage, contentWidthPx });
  });
  const fontDeclarations = [...context.fonts].map((family) => `<style:font-face style:name="${xmlEscape(family)}" svg:font-family="${xmlEscape(/\s/.test(family) ? `'${family}'` : family)}"/>`).join("");
  const automatic = `${context.paragraphStyles.xml()}${context.textStyles.xml()}${context.tableStyles.xml()}${context.columnStyles.xml()}${context.cellStyles.xml()}${context.graphicStyles.xml()}${listStylesXml(document)}`;
  const content = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NAMESPACES} office:version="1.3"><office:font-face-decls>${fontDeclarations}</office:font-face-decls><office:automatic-styles>${automatic}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`;
  const headingSizes = [16, 13, 12, 11, 11, 11];
  const headingStyles = headingSizes.map((size, index) => `<style:style style:name="Heading_20_${index + 1}" style:display-name="Heading ${index + 1}" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Standard" style:default-outline-level="${index + 1}" style:class="text"><style:text-properties fo:font-size="${size}pt" fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>`).join("");
  const styles = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NAMESPACES} office:version="1.3"><office:font-face-decls>${fontDeclarations}</office:font-face-decls><office:styles><style:default-style style:family="paragraph"><style:paragraph-properties fo:orphans="2" fo:widows="2"/><style:text-properties style:font-name="Calibri" fo:font-family="Calibri" fo:font-size="11pt" fo:language="en" fo:country="US"/></style:default-style><style:style style:name="Standard" style:family="paragraph" style:class="text"/><style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:class="text"><style:paragraph-properties fo:margin-top="0.1665in" fo:margin-bottom="0.0835in" fo:keep-with-next="always"/></style:style>${headingStyles}<style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:text-properties fo:font-size="10pt"/></style:style><style:style style:name="Graphics" style:family="graphic"/><text:outline-style style:name="Outline">${Array.from({ length: 10 }, (_value, index) => `<text:outline-level-style text:level="${index + 1}" style:num-format=""/>`).join("")}</text:outline-style><text:notes-configuration text:note-class="footnote" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/><text:notes-configuration text:note-class="endnote" style:num-format="i" text:start-value="0"/></office:styles><office:automatic-styles>${layouts.join("")}</office:automatic-styles><office:master-styles>${masters.join("")}</office:master-styles></office:document-styles>`;
  const title = String(options.title || "").trim();
  const meta = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta ${NAMESPACES} office:version="1.3"><office:meta><meta:generator>Simple Docs</meta:generator>${title ? `<dc:title>${xmlEscape(title.slice(0, 255))}</dc:title>` : ""}</office:meta></office:document-meta>`;
  const manifestEntries = [
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${ODT_MIME_TYPE}"/>`,
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>',
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>',
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>',
    ...context.media.map((item) => `<manifest:file-entry manifest:full-path="${item.path}" manifest:media-type="${item.type}"/>`),
  ];
  const manifest = `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">${manifestEntries.join("")}</manifest:manifest>`;
  if (context.counts.images) context.warnings.push(`${context.counts.images} ${context.counts.images === 1 ? "picture" : "pictures"} could not be embedded and ${context.counts.images === 1 ? "is" : "are"} shown as a placeholder in the ODT file.`);
  if (context.counts.objects) context.warnings.push(`${context.counts.objects} ${context.counts.objects === 1 ? "object" : "objects"} that ODT cannot represent ${context.counts.objects === 1 ? "is" : "are"} shown as a placeholder.`);
  return { "content.xml": content, "styles.xml": styles, "meta.xml": meta, "META-INF/manifest.xml": manifest, media: context.media };
}

/**
 * Serialize to ODT bytes. The uncompressed "mimetype" entry comes first, as
 * the OpenDocument packaging rules require.
 */
export async function documentToOdt(document, options = {}) {
  const parts = documentToOdtParts(document, options);
  const zip = new JSZip();
  const date = new Date("1980-01-01T00:00:00Z");
  zip.file("mimetype", ODT_MIME_TYPE, { compression: "STORE", date, createFolders: false });
  for (const name of ["content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml"]) zip.file(name, parts[name], { compression: "DEFLATE", date, createFolders: false });
  for (const item of parts.media) zip.file(item.path, item.bytes, { compression: item.extension === "svg" ? "DEFLATE" : "STORE", date, createFolders: false });
  return zip.generateAsync({ type: "uint8array", mimeType: ODT_MIME_TYPE, compressionOptions: { level: 6 }, platform: "DOS" });
}
