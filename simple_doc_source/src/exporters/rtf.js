// Native Rich Text Format writer for the WordCanvas document model. It runs
// locally without an Office engine and keeps paragraphs, character formatting,
// headings, real lists, tables (spans, shading, borders), PNG/JPEG pictures,
// links, bookmarks, footnotes/endnotes, sections with page setup, and section
// headers/footers. Pass a document prepared by prepareDocumentForExport() so
// pictures are embedded data: URIs.

import { createListState, headingLevel, mathText, noteReference, safeLink, sectionsOf, storyHasContent } from "../document-export.js";

const TWIPS_PER_PX = 15;
const DEFAULT_FONT = "Calibri";
const DEFAULT_BORDER = { color: "#bfbfbf", widthPx: 1, style: "single" };
const LEVEL_FORMATS = { decimal: 0, upperRoman: 1, lowerRoman: 2, upperLetter: 3, lowerLetter: 4, bullet: 23 };
const BORDER_STYLES = { single: "\\brdrs", double: "\\brdrdb", dashed: "\\brdrdash", dotted: "\\brdrdot" };

const twips = (px) => Math.round((Number(px) || 0) * TWIPS_PER_PX);

/** Escape text for RTF: control characters, Unicode as \uN with a ? fallback. */
export function rtfText(value) {
  let output = "";
  for (const character of String(value ?? "")) {
    const code = character.codePointAt(0);
    if (character === "\\" || character === "{" || character === "}") output += `\\${character}`;
    else if (character === "\t") output += "\\tab ";
    else if (character === "\n" || character === "\v" || character === "\r") output += "\\line ";
    else if (code === 0xa0) output += "\\~";
    else if (code === 0xad) output += "\\-";
    else if (code < 0x20) continue;
    else if (code < 0x80) output += character;
    else if (code <= 0xffff) output += `\\u${code > 32767 ? code - 65536 : code}?`;
    else {
      const high = Math.floor((code - 0x10000) / 0x400) + 0xd800;
      const low = ((code - 0x10000) % 0x400) + 0xdc00;
      output += `\\u${high - 65536}?\\u${low - 65536}?`;
    }
  }
  return output;
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
  const match = /^data:(image\/(?:png|jpe?g));base64,([a-z0-9+/=\s]+)$/i.exec(String(source || ""));
  if (!match) return null;
  const binary = atob(match[2].replace(/\s/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return { bytes, type: /png/i.test(match[1]) ? "png" : "jpeg" };
}

function pixelSize(bytes, type) {
  try {
    if (type === "png" && bytes.length >= 24) {
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      return { width: view.getUint32(16), height: view.getUint32(20) };
    }
    if (type === "jpeg") {
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      let offset = 2;
      while (offset + 9 < bytes.length) {
        if (bytes[offset] !== 0xff) { offset += 1; continue; }
        const marker = bytes[offset + 1];
        if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { offset += 2; continue; }
        if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) return { width: view.getUint16(offset + 7), height: view.getUint16(offset + 5) };
        offset += 2 + view.getUint16(offset + 2);
      }
    }
  } catch { /* Fall back to the display size. */ }
  return null;
}

function hex(bytes) {
  const lines = [];
  let line = "";
  for (let index = 0; index < bytes.length; index += 1) {
    line += bytes[index].toString(16).padStart(2, "0");
    if (line.length >= 128) { lines.push(line); line = ""; }
  }
  if (line) lines.push(line);
  return lines.join("\n");
}

function createContext(document, options) {
  const fonts = new Map([[DEFAULT_FONT, 0]]);
  const colors = new Map();
  return {
    document,
    warnings: options.warnings || [],
    counts: { images: 0, objects: 0 },
    lists: createListState(document),
    listIndex: new Map(),
    font(value) {
      const family = familyName(value) || DEFAULT_FONT;
      if (!fonts.has(family)) fonts.set(family, fonts.size);
      return fonts.get(family);
    },
    color(value) {
      const color = hexColor(value);
      if (!color) return 0;
      if (!colors.has(color)) colors.set(color, colors.size + 1);
      return colors.get(color);
    },
    fonts,
    colors,
  };
}

function characterCodes(style, context) {
  const codes = [`\\f${context.font(style?.fontFamily)}`];
  const size = Number(style?.fontSizePx);
  codes.push(`\\fs${Number.isFinite(size) && size > 0 ? Math.max(2, Math.round(size * 1.5)) : 22}`);
  if (style?.bold) codes.push("\\b");
  if (style?.italic) codes.push("\\i");
  if (style?.underline) codes.push("\\ul");
  if (style?.strikethrough) codes.push("\\strike");
  if (style?.caps) codes.push("\\caps");
  if (style?.smallCaps) codes.push("\\scaps");
  if (style?.hidden) codes.push("\\v");
  if (style?.verticalAlign === "super") codes.push("\\super");
  if (style?.verticalAlign === "sub") codes.push("\\sub");
  const color = context.color(style?.color);
  if (color) codes.push(`\\cf${color}`);
  const highlight = context.color(style?.highlightColor);
  if (highlight) codes.push(`\\chcbpat${highlight}\\highlight${highlight}`);
  if (style?.rtl) codes.push("\\rtlch");
  const spacing = Number(style?.letterSpacingPx);
  if (Number.isFinite(spacing) && spacing) codes.push(`\\expndtw${twips(spacing)}`);
  return codes.join("");
}

function fieldRtf(instruction, result, codes) {
  return `{\\field{\\*\\fldinst{${codes} ${instruction}}}{\\fldrslt{${codes} ${result}}}}`;
}

function textWithFields(text, codes, inBand) {
  if (!inBand) return rtfText(text);
  // Header/footer {page} and {pages} tokens become live page fields.
  return String(text).split(/(\{page\}|\{pages\})/i).map((piece) => {
    if (/^\{page\}$/i.test(piece)) return `}${fieldRtf("PAGE", "1", codes)}{${codes} `;
    if (/^\{pages\}$/i.test(piece)) return `}${fieldRtf("NUMPAGES", "1", codes)}{${codes} `;
    return rtfText(piece);
  }).join("");
}

function noteRtf(reference, run, context) {
  const blocks = context.document[reference.kind === "endnote" ? "endnotes" : "footnotes"]?.[reference.id] || [];
  const codes = characterCodes(run.style, context);
  const body = blocks.map((block, index) => paragraphRtf(block, context, { inNote: true, first: index === 0, last: index === blocks.length - 1 })).join("");
  return `{${codes}\\super\\chftn}{\\footnote${reference.kind === "endnote" ? "\\ftnalt" : ""}\\pard\\plain{${codes}\\super\\chftn} ${body || "\\pard\\plain "}}`;
}

function bookmarkMarkers(context, blockId) {
  const markers = [];
  for (const [name, range] of Object.entries(context.document.bookmarks || {})) {
    if (!/^[\w.-]{1,40}$/.test(name)) continue;
    if (range?.start?.blockId === blockId) markers.push({ offset: Math.max(0, Number(range.start.offset) || 0), rtf: `{\\*\\bkmkstart ${name}}` });
    if (range?.end?.blockId === blockId) markers.push({ offset: Math.max(0, Number(range.end.offset) || 0), rtf: `{\\*\\bkmkend ${name}}`, end: true });
  }
  return markers.sort((left, right) => left.offset - right.offset || Number(left.end) - Number(right.end));
}

function runsRtf(paragraph, context, options = {}) {
  const runs = paragraph.runs || [];
  const markers = bookmarkMarkers(context, paragraph.id);
  const pieces = [];
  let offset = 0;
  let marker = 0;
  const flushMarkers = (limit) => {
    while (marker < markers.length && markers[marker].offset <= limit) pieces.push({ raw: markers[marker++].rtf });
  };
  for (const run of runs) {
    const text = String(run?.text || "");
    const reference = noteReference(run);
    flushMarkers(offset);
    if (reference) {
      pieces.push({ raw: noteRtf(reference, run, context), link: null });
    } else if (run?.style?.equation) {
      const codes = characterCodes(run.style, context);
      pieces.push({ raw: `{${codes} ${rtfText(mathText(run.style.equation.root))}}`, link: safeLink(run.style.link) });
    } else {
      // Split text at bookmark offsets inside this run.
      let start = 0;
      while (marker < markers.length && markers[marker].offset < offset + text.length) {
        const at = markers[marker].offset - offset;
        if (at > start) pieces.push({ run, text: text.slice(start, at) });
        pieces.push({ raw: markers[marker++].rtf });
        start = Math.max(start, at);
      }
      if (start < text.length) pieces.push({ run, text: text.slice(start) });
    }
    offset += text.length;
  }
  flushMarkers(Number.MAX_SAFE_INTEGER);
  let output = "";
  for (let index = 0; index < pieces.length;) {
    const piece = pieces[index];
    const link = piece.run ? safeLink(piece.run.style?.link) : piece.link || null;
    if (!link) {
      output += piece.raw ?? runRtf(piece.run, piece.text, context, options);
      index += 1;
      continue;
    }
    let inner = "";
    while (index < pieces.length) {
      const next = pieces[index];
      const nextLink = next.run ? safeLink(next.run.style?.link) : next.link || null;
      if (nextLink !== link) break;
      inner += next.raw ?? runRtf(next.run, next.text, context, options);
      index += 1;
    }
    const instruction = link.startsWith("#") ? `HYPERLINK \\\\l "${rtfText(link.slice(1))}"` : `HYPERLINK "${rtfText(link)}"`;
    output += `{\\field{\\*\\fldinst{${instruction}}}{\\fldrslt{${inner}}}}`;
  }
  return output;
}

function runRtf(run, text, context, options) {
  if (!text) return "";
  const codes = characterCodes(run.style, context);
  return `{${codes} ${textWithFields(text, codes, options.inBand)}}`;
}

function listDefinitions(context) {
  const lists = Object.values(context.document.lists || {});
  if (!lists.length) return "";
  const definitions = [];
  const overrides = [];
  lists.forEach((list, index) => {
    const listId = 1000 + index;
    const override = index + 1;
    context.listIndex.set(list.id, override);
    const levels = Array.from({ length: 9 }, (_value, levelIndex) => {
      const level = list.levels?.[levelIndex] || list.levels?.at(-1) || { format: "bullet", bulletChar: "•", indentLeftPx: 24 * (levelIndex + 1), hangingPx: 18, start: 1 };
      const format = LEVEL_FORMATS[level.format] ?? 0;
      let template;
      let numbers = "";
      if (level.format === "bullet") template = `\\'01${rtfText(level.bulletChar || "•")}`;
      else {
        const pattern = level.text || `%${levelIndex + 1}.`;
        const parts = [];
        let position = 0;
        for (const token of pattern.split(/(%[1-9])/)) {
          if (!token) continue;
          if (/^%[1-9]$/.test(token)) {
            parts.push(`\\'${(Number(token[1]) - 1).toString(16).padStart(2, "0")}`);
            position += 1;
            numbers += `\\'${position.toString(16).padStart(2, "0")}`;
          } else {
            for (const character of token) {
              parts.push(rtfText(character));
              position += 1;
            }
          }
        }
        template = `\\'${Math.min(255, position).toString(16).padStart(2, "0")}${parts.join("")}`;
      }
      const left = twips(level.indentLeftPx ?? 24 * (levelIndex + 1));
      const hanging = twips(level.hangingPx ?? 18);
      return `{\\listlevel\\levelnfc${format}\\levelnfcn${format}\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat${Math.max(0, Math.floor(Number(level.start) || 1))}\\levelspace0\\levelindent0{\\leveltext${template};}{\\levelnumbers${numbers};}\\fi-${hanging}\\li${left}\\lin${left}}`;
    }).join("");
    definitions.push(`{\\list\\listtemplateid${2000 + index}\\listhybrid${levels}{\\listname ;}\\listid${listId}}`);
    overrides.push(`{\\listoverride\\listid${listId}\\listoverridecount0\\ls${override}}`);
  });
  return `{\\*\\listtable${definitions.join("")}}{\\*\\listoverridetable${overrides.join("")}}`;
}

function paragraphCodes(paragraph, context, options = {}) {
  const style = paragraph.style || {};
  const codes = ["\\pard\\plain"];
  const level = headingLevel(paragraph);
  if (level) codes.push(`\\s${level}\\outlinelevel${level - 1}`);
  if (style.direction === "rtl") codes.push("\\rtlpar");
  codes.push({ center: "\\qc", right: "\\qr", justify: "\\qj" }[style.align] || "\\ql");
  const marker = options.marker;
  if (marker) {
    const definition = context.document.lists?.[style.list.listId]?.levels?.[marker.level];
    const left = twips(definition?.indentLeftPx ?? 24 * (marker.level + 1));
    const hanging = twips(definition?.hangingPx ?? 18);
    codes.push(`\\fi-${hanging}\\li${left}`);
  } else {
    if (style.indentLeftPx) codes.push(`\\li${twips(style.indentLeftPx)}`);
    if (style.indentFirstLinePx) codes.push(`\\fi${twips(style.indentFirstLinePx)}`);
  }
  if (style.indentRightPx) codes.push(`\\ri${twips(style.indentRightPx)}`);
  if (Number(style.spaceBeforePx) > 0) codes.push(`\\sb${twips(style.spaceBeforePx)}`);
  if (Number(style.spaceAfterPx) >= 0) codes.push(`\\sa${twips(style.spaceAfterPx)}`);
  if (style.lineRule === "exact" && Number(style.lineHeightPx) > 0) codes.push(`\\sl-${twips(style.lineHeightPx)}\\slmult0`);
  else if (style.lineRule === "atLeast" && Number(style.lineHeightPx) > 0) codes.push(`\\sl${twips(style.lineHeightPx)}\\slmult0`);
  else if (Number(style.lineHeight) > 0 && Math.abs(Number(style.lineHeight) - 1) > 0.001) codes.push(`\\sl${Math.round(240 * Number(style.lineHeight))}\\slmult1`);
  if (style.keepWithNext) codes.push("\\keepn");
  if (style.keepLinesTogether) codes.push("\\keep");
  if (style.pageBreakBefore) codes.push("\\pagebb");
  const shading = context.color(style.shading);
  if (shading) codes.push(`\\cbpat${shading}`);
  for (const stop of style.tabStops || []) {
    const align = { center: "\\tqc", right: "\\tqr", decimal: "\\tqdec" }[stop.align] || "";
    const leader = { dot: "\\tldot", dash: "\\tlhyph", underscore: "\\tlul" }[stop.leader] || "";
    codes.push(`${align}${leader}\\tx${twips(stop.posPx)}`);
  }
  if (options.inTable) codes.push("\\intbl");
  if (marker) codes.push(`\\ls${context.listIndex.get(style.list.listId) || 1}\\ilvl${marker.level}`);
  return codes.join("");
}

function paragraphRtf(paragraph, context, options = {}) {
  const marker = paragraph.style?.list && context.listIndex.has(paragraph.style.list.listId) ? context.lists.marker(paragraph) : null;
  const codes = paragraphCodes(paragraph, context, { ...options, marker });
  const listText = marker ? `{\\listtext\\pard\\plain${characterCodes(paragraph.runs?.[0]?.style, context)} ${rtfText(marker.text)}\\tab}` : "";
  const ending = options.inNote ? (options.last ? "" : "\\par") : options.cellEnd ? "\\cell" : "\\par";
  return `${listText}${codes} ${runsRtf(paragraph, context, options)}${ending}\n`;
}

function imageRtf(block, context, options = {}) {
  const decoded = decodeDataImage(block.src);
  const ending = options.cellEnd ? "\\cell" : "\\par";
  const align = { center: "\\qc", right: "\\qr" }[block.align] || "\\ql";
  const pard = `\\pard\\plain${align}${options.inTable ? "\\intbl" : ""}`;
  if (!decoded) {
    context.counts.images += 1;
    return `${pard} {\\i ${rtfText("[Picture not available in this export]")}}${ending}\n`;
  }
  const width = Math.max(1, Number(block.widthPx) || 320);
  const height = Math.max(1, Number(block.heightPx) || 240);
  const intrinsic = pixelSize(decoded.bytes, decoded.type) || { width, height };
  return `${pard} {\\pict\\${decoded.type === "png" ? "pngblip" : "jpegblip"}\\picw${intrinsic.width}\\pich${intrinsic.height}\\picwgoal${twips(width)}\\pichgoal${twips(height)}\n${hex(decoded.bytes)}\n}${ending}\n`;
}

function tableRtf(table, context, options = {}) {
  const rows = table.rows || [];
  const columns = Math.max(1, ...rows.map((row) => (row.cells || []).reduce((sum, cell) => sum + Math.max(1, Number(cell.colSpan) || 1), 0)));
  const contentWidth = Math.max(1440, options.contentWidthTw || 9360);
  const fractions = Array.isArray(table.colFractions) && table.colFractions.length === columns ? table.colFractions : Array.from({ length: columns }, () => 1 / columns);
  const edges = [];
  let total = 0;
  for (const fraction of fractions) { total += Math.max(0.01, Number(fraction) || 0); edges.push(total); }
  const right = (column) => Math.round(contentWidth * edges[Math.min(columns, column) - 1] / total);
  const pending = new Array(columns).fill(null);
  const border = (side, cell) => {
    const value = cell?.borders?.[side] || table.defaultBorders?.[side] || (cell?.borders || table.defaultBorders ? null : DEFAULT_BORDER);
    if (!value || value.style === "none") return "";
    const color = context.color(value.color);
    return `\\clbrdr${side[0]}${BORDER_STYLES[value.style] || "\\brdrs"}\\brdrw${Math.max(2, twips(value.widthPx || 1))}${color ? `\\brdrcf${color}` : ""}`;
  };
  let output = "";
  rows.forEach((row) => {
    const definitions = [];
    const contents = [];
    let column = 0;
    const emitContinues = () => {
      while (column < columns && pending[column]?.remaining > 0) {
        const merge = pending[column];
        merge.remaining -= 1;
        column += merge.span;
        definitions.push(`\\clvmrg${merge.borders}\\cellx${right(column)}`);
        contents.push(`\\pard\\plain\\intbl \\cell\n`);
        if (merge.remaining <= 0) pending[column - merge.span] = null;
      }
    };
    for (const cell of row.cells || []) {
      emitContinues();
      const span = Math.max(1, Math.min(columns - column, Number(cell.colSpan) || 1));
      const borders = ["top", "left", "bottom", "right"].map((side) => border(side, cell)).join("");
      const shading = context.color(cell.shading);
      const rowSpan = Math.max(1, Number(cell.rowSpan) || 1);
      if (rowSpan > 1) pending[column] = { remaining: rowSpan - 1, span, borders };
      column += span;
      definitions.push(`${rowSpan > 1 ? "\\clvmgf" : ""}\\clvertalt${shading ? `\\clcbpat${shading}` : ""}${borders}\\cellx${right(column)}`);
      const blocks = cell.blocks?.length ? cell.blocks : [{ kind: "paragraph", id: "", runs: [], style: {} }];
      contents.push(blocks.map((block, index) => blockRtf(block, context, { inTable: true, cellEnd: index === blocks.length - 1, contentWidthTw: right(column) - right(column - span) })).join(""));
    }
    emitContinues();
    const header = row.props?.repeatHeader ? "\\trhdr" : "";
    const rowDefinition = `\\trowd\\trgaph108\\trleft0${header}${definitions.join("")}`;
    output += `${rowDefinition}\n${contents.join("")}${rowDefinition}\\row\n`;
  });
  return `${output}\\pard\\plain\n`;
}

function blockRtf(block, context, options = {}) {
  if (block.kind === "paragraph") return paragraphRtf(block, context, options);
  if (block.kind === "image") return imageRtf(block, context, options);
  if (block.kind === "table") {
    if (options.inTable) {
      // Nested tables are flattened to tab-separated rows inside the cell.
      const rows = (block.rows || []).map((row) => ({ kind: "paragraph", id: "", style: {}, runs: [{ text: (row.cells || []).map((cell) => (cell.blocks || []).map((child) => (child.runs || []).map((run) => run.text || "").join("")).join(" ")).join("\t"), style: {} }] }));
      return rows.map((row, index) => paragraphRtf(row, context, { ...options, cellEnd: options.cellEnd && index === rows.length - 1 })).join("");
    }
    return tableRtf(block, context, options);
  }
  if (block.kind === "equation") {
    const equation = { kind: "paragraph", id: block.id, style: { align: block.align || "center" }, runs: [{ text: mathText(block.equation?.root) || "[Equation]", style: { italic: true, fontFamily: "Cambria Math" } }] };
    return paragraphRtf(equation, context, options);
  }
  if (block.kind === "shape") {
    const blocks = block.text?.blocks?.length ? block.text.blocks : [{ kind: "paragraph", id: block.id, style: {}, runs: [{ text: "[Shape]", style: { italic: true } }] }];
    return blocks.map((child, index) => blockRtf(child, context, { ...options, cellEnd: options.cellEnd && index === blocks.length - 1 })).join("");
  }
  context.counts.objects += 1;
  return paragraphRtf({ kind: "paragraph", id: block.id, style: {}, runs: [{ text: "[Unsupported document object]", style: { italic: true } }] }, context, options);
}

function bandRtf(destination, blocks, context, contentWidthTw) {
  if (!storyHasContent(blocks)) return "";
  const body = blocks.map((block) => blockRtf(block, context, { inBand: true, contentWidthTw })).join("");
  return `{\\${destination} ${body}}\n`;
}

function sectionCodes(props, breakType, index) {
  const width = twips(props.pageWidthPx || 816);
  const height = twips(props.pageHeightPx || 1056);
  const margin = props.marginPx || { top: 96, right: 96, bottom: 96, left: 96 };
  const codes = [`\\sectd${index ? { continuous: "\\sbknone", evenPage: "\\sbkeven", oddPage: "\\sbkodd" }[breakType] || "\\sbkpage" : ""}`];
  codes.push(`\\pgwsxn${width}\\pghsxn${height}`);
  if (width > height) codes.push("\\lndscpsxn");
  codes.push(`\\marglsxn${twips(margin.left)}\\margrsxn${twips(margin.right)}\\margtsxn${twips(margin.top)}\\margbsxn${twips(margin.bottom)}`);
  if (props.headerDistancePx !== undefined) codes.push(`\\headery${twips(props.headerDistancePx)}`);
  if (props.footerDistancePx !== undefined) codes.push(`\\footery${twips(props.footerDistancePx)}`);
  if (props.columns?.count > 1) codes.push(`\\cols${Math.min(45, Math.floor(props.columns.count))}\\colsx${twips(props.columns.gapPx || 48)}`);
  if (Number.isInteger(props.pageNumberStart)) codes.push(`\\pgnstarts${props.pageNumberStart}\\pgnrestart`);
  if (storyHasContent(props.headerFirst) || storyHasContent(props.footerFirst)) codes.push("\\titlepg");
  return codes.join("");
}

/**
 * Serialize a document model to RTF text. Unsupported pictures and objects
 * become visible placeholders and are reported through options.warnings.
 */
export function documentToRtf(document, options = {}) {
  if (!document || !Array.isArray(document.blocks)) throw new Error("The open document is not ready to export.");
  const context = createContext(document, options);
  const listTable = listDefinitions(context);
  const sections = sectionsOf(document);
  const facing = sections.some(({ props }) => storyHasContent(props.headerEven) || storyHasContent(props.footerEven));
  const first = sections[0].props || {};
  let body = "";
  sections.forEach(({ blocks, props }, index) => {
    const breakType = index ? (props.breakType ?? sections[index].breakType) : undefined;
    const margin = props.marginPx || { left: 96, right: 96 };
    const contentWidthTw = twips((props.pageWidthPx || 816) - (margin.left || 0) - (margin.right || 0));
    body += `${index ? "\\sect\n" : ""}${sectionCodes(props, breakType, index)}\n`;
    body += bandRtf(facing ? "headerr" : "header", props.header, context, contentWidthTw);
    body += bandRtf("footer" + (facing ? "r" : ""), props.footer, context, contentWidthTw);
    if (facing) {
      body += bandRtf("headerl", props.headerEven || props.header, context, contentWidthTw);
      body += bandRtf("footerl", props.footerEven || props.footer, context, contentWidthTw);
    }
    body += bandRtf("headerf", props.headerFirst, context, contentWidthTw);
    body += bandRtf("footerf", props.footerFirst, context, contentWidthTw);
    for (const block of blocks) body += blockRtf(block, context, { contentWidthTw });
  });
  const fontTable = [...context.fonts].map(([family, index]) => `{\\f${index}\\fnil\\fcharset0 ${rtfText(family)};}`).join("");
  const colorTable = [...context.colors.keys()].map((color) => `\\red${parseInt(color.slice(1, 3), 16)}\\green${parseInt(color.slice(3, 5), 16)}\\blue${parseInt(color.slice(5, 7), 16)};`).join("");
  const headingSizes = [32, 26, 24, 22, 22, 22];
  const styles = `{\\stylesheet{\\s0\\snext0 Normal;}${headingSizes.map((size, index) => `{\\s${index + 1}\\sbasedon0\\snext0\\outlinelevel${index}\\keepn\\b\\fs${size} heading ${index + 1};}`).join("")}}`;
  const margin = first.marginPx || { top: 96, right: 96, bottom: 96, left: 96 };
  const title = options.title ? `{\\title ${rtfText(String(options.title).slice(0, 255))}}` : "";
  const header = [
    `{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1`,
    `{\\fonttbl${fontTable}}`,
    `{\\colortbl;${colorTable}}`,
    styles,
    listTable,
    "{\\*\\generator Simple Docs;}",
    `{\\info${title}}`,
    `\\paperw${twips(first.pageWidthPx || 816)}\\paperh${twips(first.pageHeightPx || 1056)}\\margl${twips(margin.left)}\\margr${twips(margin.right)}\\margt${twips(margin.top)}\\margb${twips(margin.bottom)}${facing ? "\\facingp" : ""}\\widowctrl\\ftnbj\\aenddoc\\ftnnar\\aftnnrlc\\viewkind1`,
  ].filter(Boolean).join("\n");
  if (context.counts.images) context.warnings.push(`${context.counts.images} ${context.counts.images === 1 ? "picture is" : "pictures are"} not PNG or JPEG and ${context.counts.images === 1 ? "is" : "are"} shown as a placeholder in the RTF file.`);
  if (context.counts.objects) context.warnings.push(`${context.counts.objects} ${context.counts.objects === 1 ? "object" : "objects"} that RTF cannot represent ${context.counts.objects === 1 ? "is" : "are"} shown as a placeholder.`);
  return `${header}\n${body}}\n`;
}

/** RTF bytes (7-bit ASCII) for saving. */
export function documentToRtfBytes(document, options = {}) {
  return new TextEncoder().encode(documentToRtf(document, options));
}
