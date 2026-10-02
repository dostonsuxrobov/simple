const HTML_IMAGE_SOURCE = /^data:image\/(?:png|jpe?g|gif|webp|bmp|svg\+xml);base64,[a-z0-9+/=\s]+$/i;
const EMBEDDABLE_IMAGE_TYPE = /^image\/(?:png|jpeg|gif|webp|bmp|svg\+xml)$/i;
const MAX_EMBEDDED_IMAGE_BYTES = 64 * 1024 * 1024;
const exportWarningsByDocument = new WeakMap();

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  })[character]);
}

function safeNumber(value, fallback = 0, min = 0, max = 10_000) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.min(max, Math.max(min, number)) : fallback;
}

function safeColor(value) {
  const color = String(value || "").trim();
  return /^(?:#[0-9a-f]{3,8}|rgba?\([\d\s.,%]+\)|hsla?\([\d\s.,%deg]+\))$/i.test(color) ? color : null;
}

function safeFontFamily(value) {
  const family = String(value || "").trim();
  return /^[\p{L}\p{N} _,-]{1,80}$/u.test(family) ? family : null;
}

function safeLink(value) {
  const link = String(value || "").trim();
  if (link.startsWith("#") && /^#[\w.-]+$/.test(link)) return link;
  try {
    const parsed = new URL(link);
    return ["http:", "https:", "mailto:"].includes(parsed.protocol) ? parsed.toString() : null;
  } catch {
    return null;
  }
}

function safeImageSource(value) {
  const source = String(value || "");
  return HTML_IMAGE_SOURCE.test(source) ? source.replace(/\s/g, "") : null;
}

function safeAnchorId(value) {
  const id = String(value || "").trim();
  return /^[\w.-]{1,200}$/.test(id) ? id : null;
}

function bytesToBase64(bytes) {
  const parts = [];
  for (let offset = 0; offset < bytes.length; offset += 32768) parts.push(String.fromCharCode(...bytes.subarray(offset, offset + 32768)));
  return btoa(parts.join(""));
}

function isBlobLike(value) {
  return Boolean(value) && typeof value === "object" && typeof value.size === "number" && typeof value.type === "string" && typeof value.arrayBuffer === "function";
}

async function defaultImageResolver(source) {
  const response = await fetch(source);
  if (!response.ok) throw new Error("A document image could not be read.");
  return response.blob();
}

// Editor-created and imported images live in local object URLs. Resolve them
// into the exported copy while that editor session still owns the bytes. An
// image that cannot be embedded becomes a visible placeholder and a warning;
// it never aborts the whole export.
async function prepareDocumentForExport(document, resolveImage = defaultImageResolver) {
  const copy = structuredClone(document);
  const images = new Map();
  const omitted = [];
  async function embed(source) {
    let resolved;
    try {
      resolved = await resolveImage(source);
    } catch {
      return { reason: "unreadable" };
    }
    if (typeof resolved === "string") return safeImageSource(resolved) ? { source: safeImageSource(resolved) } : { reason: "unsupported" };
    if (!isBlobLike(resolved)) return { reason: "unreadable" };
    if (resolved.size > MAX_EMBEDDED_IMAGE_BYTES) return { reason: "too-large" };
    const type = String(resolved.type || "").toLowerCase().replace("image/jpg", "image/jpeg");
    if (!EMBEDDABLE_IMAGE_TYPE.test(type)) return { reason: "unsupported" };
    try {
      const bytes = new Uint8Array(await resolved.arrayBuffer());
      return { source: `data:${type};base64,${bytesToBase64(bytes)}` };
    } catch {
      return { reason: "unreadable" };
    }
  }
  async function visit(value) {
    if (!value || typeof value !== "object") return;
    if (value.kind === "image" && typeof value.src === "string" && value.src.startsWith("blob:")) {
      if (!images.has(value.src)) images.set(value.src, embed(value.src));
      const result = await images.get(value.src);
      if (result.source) value.src = result.source;
      else {
        omitted.push({ kind: "image", reason: result.reason });
        value.src = "";
      }
    }
    for (const child of Object.values(value)) if (child && typeof child === "object") await visit(child);
  }
  await visit(copy);
  exportWarningsByDocument.set(copy, describeOmissions(omitted));
  return copy;
}

function describeOmissions(omitted) {
  const counts = new Map();
  for (const item of omitted) counts.set(item.reason, (counts.get(item.reason) || 0) + 1);
  const messages = [];
  const pictures = (count) => `${count} ${count === 1 ? "picture" : "pictures"}`;
  if (counts.get("too-large")) messages.push(`${pictures(counts.get("too-large"))} larger than 64 MB could not be embedded and ${counts.get("too-large") === 1 ? "is" : "are"} shown as a placeholder.`);
  if (counts.get("unsupported")) messages.push(`${pictures(counts.get("unsupported"))} in a format this export cannot embed ${counts.get("unsupported") === 1 ? "is" : "are"} shown as a placeholder.`);
  if (counts.get("unreadable")) messages.push(`${pictures(counts.get("unreadable"))} could not be read and ${counts.get("unreadable") === 1 ? "is" : "are"} shown as a placeholder.`);
  return messages;
}

/** Warnings recorded while a document was prepared for export. */
function exportWarnings(document) {
  return [...(exportWarningsByDocument.get(document) || [])];
}

function roman(value) {
  if (!Number.isInteger(value) || value < 1 || value > 3999) return String(value);
  const symbols = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let remaining = value;
  let result = "";
  for (const [amount, symbol] of symbols) {
    while (remaining >= amount) {
      result += symbol;
      remaining -= amount;
    }
  }
  return result;
}

function alpha(value) {
  if (!Number.isInteger(value) || value < 1) return String(value);
  let remaining = value;
  let result = "";
  while (remaining > 0) {
    remaining -= 1;
    result = String.fromCharCode(97 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26);
  }
  return result;
}

function formatCounter(value, format) {
  if (format === "lowerLetter") return alpha(value);
  if (format === "upperLetter") return alpha(value).toUpperCase();
  if (format === "lowerRoman") return roman(value).toLowerCase();
  if (format === "upperRoman") return roman(value);
  return String(value);
}

function createListState(document) {
  const counters = new Map();
  return {
    marker(paragraph) {
      const list = paragraph?.style?.list;
      if (!list) return null;
      const definition = document.lists?.[list.listId];
      const levelIndex = Math.max(0, Math.min(8, Number(list.level) || 0));
      const level = definition?.levels?.[levelIndex];
      if (!level || level.format === "bullet") return { text: level?.bulletChar || "•", level: levelIndex, ordered: false, value: 1 };
      let values = counters.get(list.listId);
      if (!values) {
        values = [];
        counters.set(list.listId, values);
      }
      for (let index = levelIndex + 1; index < values.length; index += 1) values[index] = undefined;
      values[levelIndex] = (values[levelIndex] ?? Math.max(1, Number(level.start) || 1) - 1) + 1;
      const pattern = level.text || `%${levelIndex + 1}.`;
      const text = pattern.replace(/%(\d)/g, (_match, rawLevel) => {
        const referencedIndex = Number(rawLevel) - 1;
        const referencedLevel = definition?.levels?.[referencedIndex];
        const value = values[referencedIndex] ?? Math.max(1, Number(referencedLevel?.start) || 1);
        return formatCounter(value, referencedLevel?.format);
      });
      return { text, level: levelIndex, ordered: true, value: values[levelIndex] };
    },
  };
}

function headingLevel(paragraph) {
  const outline = Number(paragraph?.style?.outlineLevel);
  if (Number.isInteger(outline) && outline >= 0 && outline <= 8) return Math.min(6, outline + 1);
  const match = String(paragraph?.style?.namedStyle || "").match(/^heading\s*([1-9])$/i);
  return match ? Math.min(6, Number(match[1])) : null;
}

function mathText(node) {
  if (!node || typeof node !== "object") return "";
  if (["ident", "number", "op", "text"].includes(node.type)) return String(node.text || "");
  if (node.type === "space") return " ";
  if (node.type === "row") return (node.children || []).map(mathText).join("");
  if (node.type === "frac") return `(${mathText(node.num)})/(${mathText(node.den)})`;
  if (node.type === "script") return `${mathText(node.base)}${node.sub ? `_${mathText(node.sub)}` : ""}${node.sup ? `^${mathText(node.sup)}` : ""}`;
  if (node.type === "radical") return `${node.index ? `${mathText(node.index)}` : ""}√(${mathText(node.radicand)})`;
  if (node.type === "fenced") return `${node.open || "("}${mathText(node.child)}${node.close || ")"}`;
  if (node.type === "limit") return `${mathText(node.base)}${node.under ? `_${mathText(node.under)}` : ""}${node.over ? `^${mathText(node.over)}` : ""}`;
  if (node.type === "nary") return `${node.hideOp ? "" : node.op || ""}${node.sub ? `_${mathText(node.sub)}` : ""}${node.sup ? `^${mathText(node.sup)}` : ""}${mathText(node.body)}`;
  if (node.type === "matrix") return `[${(node.rows || []).map((row) => row.map(mathText).join(", ")).join("; ")}]`;
  if (node.type === "phantom") return mathText(node.child);
  return node.mathml ? "[Equation]" : "";
}

function runPlainText(run) {
  if (run?.style?.hidden) return "";
  if (run?.style?.equation) return mathText(run.style.equation.root);
  return String(run?.text || "").replace(/\v/g, "\n");
}

function runsPlainText(runs) {
  return (runs || []).map(runPlainText).join("");
}

// Notes are numbered by first reference in reading order, as Word displays them.
function createNoteState(document) {
  const numbers = new Map();
  const order = [];
  return {
    number(kind, id) {
      const key = `${kind}:${id}`;
      if (!numbers.has(key)) {
        numbers.set(key, numbers.size + 1);
        order.push({ kind, id, number: numbers.size });
      }
      return numbers.get(key);
    },
    referenced() {
      return order.map((entry) => ({ ...entry, blocks: document[entry.kind === "endnote" ? "endnotes" : "footnotes"]?.[entry.id] || [] }));
    },
    unreferenced() {
      const result = [];
      for (const [kind, collection] of [["footnote", document.footnotes], ["endnote", document.endnotes]]) {
        for (const [id, blocks] of Object.entries(collection || {})) {
          if (!numbers.has(`${kind}:${id}`)) result.push({ kind, id, blocks });
        }
      }
      return result;
    },
  };
}

function noteReference(run) {
  if (run?.style?.footnoteRef) return { kind: "footnote", id: run.style.footnoteRef };
  if (run?.style?.endnoteRef) return { kind: "endnote", id: run.style.endnoteRef };
  return null;
}

// Bookmark starts by block, with character offsets, for HTML/RTF anchors.
function bookmarkStarts(document) {
  const starts = new Map();
  for (const [name, range] of Object.entries(document?.bookmarks || {})) {
    const id = safeAnchorId(name);
    const blockId = range?.start?.blockId;
    if (!id || !blockId) continue;
    if (!starts.has(blockId)) starts.set(blockId, []);
    starts.get(blockId).push({ name: id, offset: Math.max(0, Number(range.start.offset) || 0) });
  }
  for (const list of starts.values()) list.sort((left, right) => left.offset - right.offset);
  return starts;
}

function runHtml(run, context) {
  if (run?.style?.hidden) return "";
  const note = noteReference(run);
  if (note && context?.notes) {
    const number = context.notes.number(note.kind, note.id);
    return `<sup class="note-ref"><a href="#note-${number}" id="note-ref-${number}">${number}</a></sup>`;
  }
  const rawText = run?.style?.equation ? mathText(run.style.equation.root) : String(run?.text || "");
  let content = escapeHtml(rawText).replace(/[\n\v]/g, "<br>");
  const styles = [];
  const family = safeFontFamily(run?.style?.fontFamily);
  const color = safeColor(run?.style?.color);
  const highlight = safeColor(run?.style?.highlightColor);
  if (family) styles.push(`font-family:${family}`);
  if (Number.isFinite(Number(run?.style?.fontSizePx))) styles.push(`font-size:${safeNumber(run.style.fontSizePx, 14, 6, 144)}px`);
  if (color) styles.push(`color:${color}`);
  if (highlight) styles.push(`background:${highlight}`);
  if (run?.style?.bold) styles.push("font-weight:700");
  if (run?.style?.italic) styles.push("font-style:italic");
  const decorations = [run?.style?.underline && "underline", run?.style?.strikethrough && "line-through"].filter(Boolean);
  if (decorations.length) styles.push(`text-decoration:${decorations.join(" ")}`);
  if (run?.style?.verticalAlign === "sub") styles.push("vertical-align:sub;font-size:.8em");
  if (run?.style?.verticalAlign === "super") styles.push("vertical-align:super;font-size:.8em");
  if (run?.style?.caps) styles.push("text-transform:uppercase");
  if (run?.style?.smallCaps) styles.push("font-variant:small-caps");
  if (styles.length) content = `<span style="${styles.join(";")}">${content}</span>`;
  const link = safeLink(run?.style?.link);
  if (link) content = `<a href="${escapeHtml(link)}">${content}</a>`;
  return content;
}

// Split runs so bookmark anchors land at their exact character offsets.
function runsWithAnchors(runs, anchors) {
  if (!anchors?.length) return (runs || []).map((run) => ({ run }));
  const output = [];
  let offset = 0;
  let next = 0;
  for (const run of runs || []) {
    const text = String(run?.text || "");
    const end = offset + text.length;
    let cursor = 0;
    while (next < anchors.length && anchors[next].offset < end) {
      const at = Math.max(0, anchors[next].offset - offset);
      if (at > cursor) output.push({ run: { ...run, text: text.slice(cursor, at) } });
      output.push({ anchor: anchors[next].name });
      cursor = Math.max(cursor, at);
      next += 1;
    }
    if (cursor === 0) output.push({ run });
    else if (cursor < text.length) output.push({ run: { ...run, text: text.slice(cursor) } });
    offset = end;
  }
  while (next < anchors.length) output.push({ anchor: anchors[next++].name });
  return output;
}

function paragraphHtml(paragraph, context) {
  const level = headingLevel(paragraph);
  const tag = level ? `h${level}` : "p";
  const marker = context.lists.marker(paragraph);
  const styles = [];
  if (["left", "center", "right", "justify"].includes(paragraph?.style?.align)) styles.push(`text-align:${paragraph.style.align}`);
  if (paragraph?.style?.direction === "rtl") styles.push("direction:rtl");
  if (paragraph?.style?.pageBreakBefore) styles.push("break-before:page");
  const before = safeNumber(paragraph?.style?.spaceBeforePx, 0, 0, 500);
  const after = safeNumber(paragraph?.style?.spaceAfterPx, 0, 0, 500);
  if (before) styles.push(`margin-top:${before}px`);
  if (after) styles.push(`margin-bottom:${after}px`);
  if (marker?.level) styles.push(`margin-left:${marker.level * 1.5}em`);
  const markerHtml = marker ? `<span class="list-marker">${escapeHtml(marker.text)}</span>` : "";
  const pieces = runsWithAnchors(paragraph.runs, context.anchors?.get(paragraph.id));
  const content = pieces.map((piece) => piece.anchor ? `<a id="${escapeHtml(piece.anchor)}"></a>` : runHtml(piece.run, context)).join("") || "<br>";
  return `<${tag}${styles.length ? ` style="${styles.join(";")}"` : ""}>${markerHtml}${content}</${tag}>`;
}

function imagePlaceholder(context) {
  if (context?.warnings) context.warnings.missingImages += 1;
  return '<p class="unsupported">[Image not available in this export]</p>';
}

function blocksHtml(blocks, context) {
  return (blocks || []).map((block) => {
    if (block.kind === "paragraph") return paragraphHtml(block, context);
    if (block.kind === "image") {
      const source = safeImageSource(block.src);
      if (!source) return imagePlaceholder(context);
      const width = safeNumber(block.widthPx, 320, 1, 10_000);
      const height = safeNumber(block.heightPx, 240, 1, 10_000);
      return `<figure style="text-align:${["left", "center", "right"].includes(block.align) ? block.align : "center"}"><img src="${source}" alt="Document image" width="${width}" height="${height}"></figure>`;
    }
    if (block.kind === "table") {
      const headerRow = block.condOverrides?.firstRow === true;
      const rows = (block.rows || []).map((row, rowIndex) => `<tr>${(row.cells || []).map((cell) => {
        const cellTag = headerRow && rowIndex === 0 ? "th" : "td";
        const spans = `${cell.colSpan > 1 ? ` colspan="${Math.floor(safeNumber(cell.colSpan, 1, 1, 100))}"` : ""}${cell.rowSpan > 1 ? ` rowspan="${Math.floor(safeNumber(cell.rowSpan, 1, 1, 100))}"` : ""}`;
        const shading = safeColor(cell.shading);
        return `<${cellTag}${spans}${shading ? ` style="background:${shading}"` : ""}>${blocksHtml(cell.blocks, context)}</${cellTag}>`;
      }).join("")}</tr>`);
      const head = headerRow && rows.length ? `<thead>${rows[0]}</thead>` : "";
      const body = (headerRow ? rows.slice(1) : rows).join("");
      return `<table>${head}${body ? `<tbody>${body}</tbody>` : ""}</table>`;
    }
    if (block.kind === "equation") return `<p class="equation">${escapeHtml(mathText(block.equation?.root) || "[Equation]")}</p>`;
    if (block.kind === "shape") {
      const content = block.text?.blocks?.length ? blocksHtml(block.text.blocks, context) : "[Shape]";
      return `<aside class="shape">${content}</aside>`;
    }
    if (context?.warnings) context.warnings.unsupportedObjects += 1;
    return '<p class="unsupported">[Unsupported document object]</p>';
  }).join("\n");
}

function sectionsOf(document) {
  const sections = [];
  let blocks = [];
  for (const block of document.blocks || []) {
    blocks.push(block);
    if (block.kind === 'paragraph' && block.style?.sectionBreak) {
      sections.push({ blocks, props: block.style.sectionBreak.props || {}, breakType: block.style.sectionBreak.type });
      blocks = [];
    }
  }
  sections.push({ blocks, props: document.section || {}, breakType: document.section?.breakType });
  return sections;
}

function storyHasContent(blocks) {
  return (blocks || []).some((block) => block.kind !== "paragraph"
    || (block.runs || []).some((run) => !run?.style?.hidden && (run?.style?.equation || String(run?.text || "").replace(/\{(?:page|pages)\}/gi, "").trim())));
}

function supplementaryStories(document, excludeRegularBands = false) {
  const stories = [];
  const seen = new Set();
  const bands = { header: 'Header', headerFirst: 'First-page header', headerEven: 'Even-page header', footer: 'Footer', footerFirst: 'First-page footer', footerEven: 'Even-page footer' };
  const sections = sectionsOf(document);
  for (const [index, section] of sections.entries()) {
    for (const [key, label] of Object.entries(bands)) {
      if (excludeRegularBands && (key === 'header' || key === 'footer')) continue;
      const blocks = section.props[key];
      if (!storyHasContent(blocks)) continue;
      const identity = JSON.stringify(blocks);
      if (seen.has(identity)) continue;
      seen.add(identity);
      stories.push({ label: sections.length > 1 ? `Section ${index + 1} ${label.toLowerCase()}` : label, blocks });
    }
  }
  return stories;
}

function noteStories(document) {
  const stories = [];
  for (const [kind, label] of [['footnotes', 'Footnote'], ['endnotes', 'Endnote']]) {
    let number = 0;
    for (const blocks of Object.values(document[kind] || {})) stories.push({ label: `${label} ${++number}`, blocks });
  }
  return stories;
}

function createWarningCounter() {
  return { missingImages: 0, unsupportedObjects: 0 };
}

function counterWarnings(counter) {
  const messages = [];
  if (counter.missingImages) messages.push(`${counter.missingImages} ${counter.missingImages === 1 ? "picture is" : "pictures are"} shown as a placeholder.`);
  if (counter.unsupportedObjects) messages.push(`${counter.unsupportedObjects} ${counter.unsupportedObjects === 1 ? "object" : "objects"} that this format cannot represent ${counter.unsupportedObjects === 1 ? "is" : "are"} shown as a placeholder.`);
  return messages;
}

function documentToHtml(document, title = "Untitled document", options = {}) {
  const notes = createNoteState(document);
  const context = { lists: createListState(document), notes, anchors: bookmarkStarts(document), warnings: options.warningCounter || createWarningCounter() };
  const sections = sectionsOf(document);
  const pageRules = sections.map(({props}, index) => `@page section${index}{size:${safeNumber(props.pageWidthPx,816,100,5000)}px ${safeNumber(props.pageHeightPx,1056,100,5000)}px}`).join('');
  const body = sections.map(({blocks, props}, index) => `<section style="page:section${index}${index && sections[index].breakType !== 'continuous' ? ';break-before:page' : ''}">${storyHasContent(props.header) ? `<header>${blocksHtml(props.header, context)}</header>` : ''}${blocksHtml(blocks, context)}${storyHasContent(props.footer) ? `<footer>${blocksHtml(props.footer, context)}</footer>` : ''}</section>`).join('\n');
  const supplements = supplementaryStories(document, true).map(({label, blocks}) => `<section class="document-note"><h2>${escapeHtml(label)}</h2>${blocksHtml(blocks, context)}</section>`).join('\n');
  const referenced = notes.referenced();
  const noteItems = referenced.map(({ number, blocks }) => `<li id="note-${number}">${blocksHtml(blocks, { ...context, notes: null })} <a href="#note-ref-${number}" class="note-back" aria-label="Back to reference">↩</a></li>`).join('\n');
  const unreferenced = notes.unreferenced().map(({ kind, blocks }, index) => `<section class="document-note"><h2>${kind === 'endnote' ? 'Endnote' : 'Footnote'} ${referenced.length + index + 1}</h2>${blocksHtml(blocks, { ...context, notes: null })}</section>`).join('\n');
  const footnotes = noteItems ? `<section class="footnotes" aria-label="Notes"><ol>${noteItems}</ol></section>` : '';
  const lang = /^[a-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$/.test(String(options.lang || '')) ? options.lang : 'en';
  const author = String(options.author || '').trim();
  return `<!doctype html>
<html lang="${escapeHtml(lang)}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="Simple Docs">${author ? `\n<meta name="author" content="${escapeHtml(author.slice(0, 200))}">` : ''}
<title>${escapeHtml(title)}</title>
<style>
  :root{font-family:Calibri,"Segoe UI",sans-serif;color:#111;background:#fff}body{max-width:8.5in;margin:0 auto;padding:.75in;line-height:1.35}p{margin:.25em 0}.list-marker{display:inline-block;min-width:1.7em;margin-right:.25em}table{width:100%;border-collapse:collapse;margin:1em 0}td,th{border:1px solid #888;padding:.35em;vertical-align:top;text-align:inherit}th{font-weight:700}td>p:first-child,th>p:first-child{margin-top:0}td>p:last-child,th>p:last-child{margin-bottom:0}img{max-width:100%;height:auto}figure{margin:1em 0}.equation{text-align:center;font-family:"Cambria Math",serif}.shape{border:1px solid #aaa;padding:.5em;margin:.75em 0}.unsupported{color:#666;font-style:italic}header,footer{color:#555;border-color:#ccc}header{border-bottom:1px solid;margin-bottom:1.5em}footer{border-top:1px solid;margin-top:1.5em}.footnotes{border-top:1px solid #aaa;margin-top:2em;padding-top:1em;font-size:.9em}.footnotes li>p{display:inline}.note-ref a,.note-back{text-decoration:none}@media print{body{max-width:none;padding:0}@page{margin:.75in}}
</style>
<style>${pageRules}.document-note{border-top:1px solid #aaa;margin-top:1.5em;padding-top:.5em}.document-note h2{font-size:1em}</style>
</head>
<body>
<main>${body}</main>${supplements}${footnotes}${unreferenced}
</body>
</html>
`;
}

// ---------------------------------------------------------------------------
// Markdown (GitHub Flavored Markdown)

const MONOSPACE_FONT = /\b(?:consolas|courier|mono|menlo|monaco|lucida console)\b/i;
const ASCII_PUNCTUATION = /[!-/:-@[-`{-~]/;

function isWhitespaceCharacter(character) {
  return !character || /\s/u.test(character);
}

function isPunctuationCharacter(character) {
  return Boolean(character) && (ASCII_PUNCTUATION.test(character) || /\p{P}|\p{S}/u.test(character));
}

// Escape only what Markdown would otherwise interpret inside a line.
function escapeMarkdownInline(value, options = {}) {
  let text = String(value)
    .replace(/\\/g, "\\\\")
    .replace(/`/g, "\\`")
    .replace(/\*/g, "\\*")
    .replace(/~/g, "\\~")
    .replace(/[[\]]/g, (character) => `\\${character}`)
    .replace(/&(?=#?\w+;)/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/(^|[^\p{L}\p{N}])_|_(?=$|[^\p{L}\p{N}])/gu, (match) => match.replace("_", "\\_"));
  if (options.table) text = text.replace(/\|/g, "\\|");
  return text;
}

// Line-start syntax: headings, quotes, lists, rules and indented code.
function escapeMarkdownLineStart(line) {
  const indent = line.match(/^[ \t]*/)[0];
  const text = line.slice(indent.length)
    .replace(/^(#{1,6})(?=\s|$)/, (match) => `\\${match}`)
    .replace(/^([-+=])(?=\s|$|\1)/, "\\$1")
    .replace(/^(\d{1,9})([.)])(?=\s|$)/, "$1\\$2");
  return `${indent}${text}`;
}

function codeSpan(text) {
  const longest = Math.max(0, ...[...String(text).matchAll(/`+/g)].map((match) => match[0].length));
  const fence = "`".repeat(longest + 1);
  const padded = /^`|`$/.test(text) || /^ .* $/.test(text) ? ` ${text} ` : text;
  return `${fence}${padded}${fence}`;
}

function markdownSegments(runs, context, options = {}) {
  const segments = [];
  for (const run of runs || []) {
    if (run?.style?.hidden) continue;
    const note = noteReference(run);
    if (note && context.notes) {
      segments.push({ raw: `[^${context.notes.number(note.kind, note.id)}]` });
      continue;
    }
    const text = run?.style?.equation ? mathText(run.style.equation.root) : String(run?.text || "");
    if (!text) continue;
    const style = run?.style || {};
    segments.push({
      text,
      bold: Boolean(style.bold) && !options.plainBold,
      italic: Boolean(style.italic),
      strike: Boolean(style.strikethrough),
      code: MONOSPACE_FONT.test(String(style.fontFamily || "")) && !/[\n\v]/.test(text),
      link: safeLink(style.link) || null,
    });
  }
  // Merge neighbours whose Markdown-relevant formatting is identical.
  const merged = [];
  for (const segment of segments) {
    const previous = merged.at(-1);
    if (previous && segment.text !== undefined && previous.text !== undefined
      && ["bold", "italic", "strike", "code", "link"].every((key) => previous[key] === segment[key])) previous.text += segment.text;
    else merged.push({ ...segment });
  }
  return merged;
}

function wrapEmphasis(core, segment, before, after, options) {
  let content = segment.code ? codeSpan(core) : escapeMarkdownInline(core, options).replace(/[\n\v]/g, options.table ? "<br>" : "\\\n");
  const markers = [];
  if (segment.strike) markers.push(["~~", "<del>", "</del>"]);
  if (segment.bold) markers.push(["**", "<strong>", "</strong>"]);
  if (segment.italic) markers.push(["*", "<em>", "</em>"]);
  if (!markers.length) return content;
  // CommonMark only treats delimiters as emphasis when they are flanking; fall
  // back to inline HTML where punctuation next to letters would break them.
  const first = Array.from(core)[0];
  const last = Array.from(core).at(-1);
  const opens = !isPunctuationCharacter(first) || isWhitespaceCharacter(before) || isPunctuationCharacter(before);
  const closes = !isPunctuationCharacter(last) || isWhitespaceCharacter(after) || isPunctuationCharacter(after);
  const html = segment.code || !opens || !closes;
  for (const [delimiter, open, close] of markers.reverse()) {
    content = html ? `${open}${content}${close}` : `${delimiter}${content}${delimiter}`;
  }
  return content;
}

function runsMarkdown(runs, context, options = {}) {
  const segments = markdownSegments(runs, context, options);
  const full = segments.map((segment) => segment.raw ?? segment.text).join("");
  const lineBreak = options.table ? "<br>" : "\\\n";
  const whitespace = (value) => value.replace(/[\n\v]/g, lineBreak);
  let position = 0;
  const rendered = segments.map((segment) => {
    const start = position;
    const text = segment.raw ?? segment.text;
    position += text.length;
    if (segment.raw) return { link: null, text: segment.raw };
    // Markers hug the non-space text; surrounding spaces stay outside them.
    const lead = text.match(/^\s*/)[0];
    const core = text.slice(lead.length).replace(/\s+$/, "");
    const trail = text.slice(lead.length + core.length);
    const before = full[start + lead.length - 1];
    const after = full[start + lead.length + core.length];
    return { link: segment.link, text: `${whitespace(lead)}${core ? wrapEmphasis(core, segment, before, after, options) : ""}${whitespace(trail)}` };
  });
  let output = "";
  for (let index = 0; index < rendered.length;) {
    const link = rendered[index].link;
    let inner = "";
    while (index < rendered.length && rendered[index].link === link) inner += rendered[index++].text;
    if (!link) {
      output += inner;
      continue;
    }
    // Whitespace around a link stays outside its brackets.
    const match = /^(\s*)([\s\S]*?)(\s*)$/.exec(inner);
    const target = link.replace(/[()\s<>]/g, (value) => encodeURIComponent(value));
    output += match[2] ? `${match[1]}[${match[2]}](${target})${match[3]}` : inner;
  }
  return output;
}

function blockPlainText(block, context) {
  if (block.kind === "paragraph") {
    const marker = context.lists.marker(block);
    return `${marker ? `${"  ".repeat(marker.level)}${marker.text} ` : ""}${runsPlainText(block.runs)}`;
  }
  if (block.kind === "image") return "[Image]";
  if (block.kind === "equation") return mathText(block.equation?.root) || "[Equation]";
  if (block.kind === "shape") return block.text?.blocks?.map((child) => blockPlainText(child, context)).join("\n") || "[Shape]";
  if (block.kind === "table") return (block.rows || []).map((row) => (row.cells || []).map((cell) => (cell.blocks || []).map((child) => blockPlainText(child, context)).join(" ")).join("\t")).join("\n");
  return "[Unsupported document object]";
}

function documentToText(document) {
  const context = { lists: createListState(document) };
  const sections = [];
  sections.push((document.blocks || []).map((block) => blockPlainText(block, context)).join("\n"));
  for (const {label, blocks} of [...supplementaryStories(document), ...noteStories(document)]) sections.push(`${label}\n${blocks.map((block) => blockPlainText(block, context)).join('\n')}`);
  return `${sections.filter(Boolean).join("\n\n")}\n`;
}

function tableCellMarkdown(cell, context) {
  const parts = [];
  for (const block of cell.blocks || []) {
    if (block.kind === "paragraph") {
      const marker = context.lists.marker(block);
      const text = runsMarkdown(block.runs, context, { table: true });
      if (text.trim() || marker) parts.push(`${marker ? `${escapeMarkdownInline(marker.text, { table: true })} ` : ""}${text.trim()}`);
    } else if (block.kind === "image") {
      const source = safeImageSource(block.src);
      if (source) parts.push(`![Document image](${source})`);
      else {
        context.warnings.missingImages += 1;
        parts.push("*[Image not available in this export]*");
      }
    } else parts.push(escapeMarkdownInline(blockPlainText(block, context).replace(/\s+/g, " ").trim(), { table: true }));
  }
  return parts.join("<br>").replace(/\n/g, " ");
}

function tableMarkdown(table, context) {
  const rows = (table.rows || []).map((row) => (row.cells || []).flatMap((cell) => {
    const content = tableCellMarkdown(cell, context);
    const span = Math.max(1, Math.min(63, Math.floor(Number(cell.colSpan) || 1)));
    return [content, ...Array(span - 1).fill("")];
  }));
  if (!rows.length) return "";
  const columns = Math.max(1, ...rows.map((row) => row.length));
  const normalize = (row) => [...row, ...Array(Math.max(0, columns - row.length)).fill("")];
  const [first, ...rest] = rows.map(normalize);
  const line = (cells) => `| ${cells.map((cell) => cell || " ").join(" | ")} |`.replace(/ {2,}\|/g, "  |");
  return [line(first), `| ${Array(columns).fill("---").join(" | ")} |`, ...rest.map(line)].join("\n");
}

function markdownBlocks(blocks, context, options = {}) {
  const output = [];
  let previousWasList = false;
  const push = (text, isList = false) => {
    if (output.length && !(isList && previousWasList)) output.push("");
    output.push(text);
    previousWasList = isList;
  };
  for (const block of blocks || []) {
    if (block.kind === "paragraph") {
      const level = headingLevel(block);
      const marker = context.lists.marker(block);
      const content = runsMarkdown(block.runs, context, { plainBold: Boolean(level) }).trim();
      if (level) {
        if (content) push(`${"#".repeat(level)} ${content.replace(/\\\n/g, " ")}`);
        continue;
      }
      if (marker) {
        const indent = "    ".repeat(marker.level);
        const prefix = marker.ordered ? `${marker.value}. ` : "- ";
        const continuation = `${indent}${" ".repeat(prefix.length)}`;
        const lines = content.split("\n").map((line, index) => index ? `${continuation}${escapeMarkdownLineStart(line.trimStart())}` : escapeMarkdownLineStart(line));
        push(`${indent}${prefix}${lines.join("\n")}`, true);
        continue;
      }
      if (!content) continue;
      push(content.split("\n").map((line) => escapeMarkdownLineStart(line.trimStart())).join("\n"));
    } else if (block.kind === "table") {
      const table = tableMarkdown(block, context);
      if (table) push(table);
    } else if (block.kind === "image") {
      const source = safeImageSource(block.src);
      if (source) push(`![Document image](${source})`);
      else {
        context.warnings.missingImages += 1;
        push("*[Image not available in this export]*");
      }
    } else if (block.kind === "equation") {
      push(codeSpan(mathText(block.equation?.root) || "Equation"));
    } else if (block.kind === "shape") {
      const inner = markdownBlocks(block.text?.blocks || [], context, options);
      push(inner || "*[Shape]*");
    } else {
      context.warnings.unsupportedObjects += 1;
      push("*[Unsupported document object]*");
    }
  }
  return output.join("\n");
}

function documentToMarkdown(document, title = "", options = {}) {
  const context = { lists: createListState(document), notes: createNoteState(document), warnings: options.warningCounter || createWarningCounter() };
  const parts = [markdownBlocks(document.blocks || [], context)];
  for (const { label, blocks } of supplementaryStories(document)) {
    const story = markdownBlocks(blocks, { ...context, lists: createListState(document) });
    if (story.trim()) parts.push(`**${label}**\n\n${story}`);
  }
  // GFM footnotes; definitions may contain several paragraphs.
  const definitions = [];
  const definition = (number, blocks) => {
    const body = markdownBlocks(blocks, { ...context, lists: createListState(document) }) || " ";
    return `[^${number}]: ${body.split("\n").map((line, lineIndex) => lineIndex && line ? `    ${line}` : line).join("\n")}`;
  };
  for (const { number, blocks } of context.notes.referenced()) definitions.push(definition(number, blocks));
  // Notes without a reference in the text are still kept, after the others.
  for (const { kind, id, blocks } of context.notes.unreferenced()) definitions.push(definition(context.notes.number(kind, id), blocks));
  if (definitions.length) parts.push(definitions.join("\n\n"));
  return `${parts.filter((part) => part.trim()).join("\n\n").replace(/\n{3,}/g, "\n\n").trimEnd()}\n`;
}

function serializeDocument(document, format, title) {
  return serializeDocumentWithWarnings(document, format, title).text;
}

/** Serialize and report what the format had to leave out or replace. */
function serializeDocumentWithWarnings(document, format, title, options = {}) {
  if (!document || !Array.isArray(document.blocks)) throw new Error("The open document is not ready to export.");
  const warningCounter = createWarningCounter();
  let text;
  if (format === "html") text = documentToHtml(document, title, { ...options, warningCounter });
  else if (format === "md") text = documentToMarkdown(document, title, { ...options, warningCounter });
  else if (format === "txt") text = documentToText(document);
  else throw new Error(`Unsupported structured document export: ${format}`);
  const prepared = exportWarnings(document);
  // Placeholders for pictures already reported by prepareDocumentForExport are
  // not counted twice.
  const omittedPictures = prepared.length ? 0 : warningCounter.missingImages;
  return { text, warnings: [...prepared, ...counterWarnings({ ...warningCounter, missingImages: omittedPictures })] };
}

export {
  bookmarkStarts,
  createListState,
  createNoteState,
  documentToHtml,
  documentToMarkdown,
  documentToText,
  exportWarnings,
  headingLevel,
  mathText,
  noteReference,
  prepareDocumentForExport,
  runPlainText,
  safeImageSource,
  safeLink,
  sectionsOf,
  serializeDocument,
  serializeDocumentWithWarnings,
  storyHasContent,
};
