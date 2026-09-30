const HTML_IMAGE_SOURCE = /^data:image\/(?:png|jpe?g|gif|webp|bmp);base64,[a-z0-9+/=\s]+$/i;

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

// Editor-created and imported images live in local object URLs. Resolve them
// into the exported copy while that editor session still owns the bytes.
async function prepareDocumentForExport(document, resolveImage = async (source) => {
  const response = await fetch(source);
  if (!response.ok) throw new Error('A document image could not be read. Please reopen the document and try again.');
  const blob = await response.blob();
  if (blob.size > 64 * 1024 * 1024 || !/^image\/(png|jpeg|gif|webp|bmp)$/i.test(blob.type)) {
    throw new Error('A document image cannot be included in this export. Save as DOCX to preserve it.');
  }
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const parts = [];
  for (let offset = 0; offset < bytes.length; offset += 32768) parts.push(String.fromCharCode(...bytes.subarray(offset, offset + 32768)));
  return `data:${blob.type};base64,${btoa(parts.join(''))}`;
}) {
  const copy = structuredClone(document);
  const images = new Map();
  async function visit(value) {
    if (!value || typeof value !== 'object') return;
    if (value.kind === 'image' && typeof value.src === 'string' && value.src.startsWith('blob:')) {
      if (!images.has(value.src)) images.set(value.src, Promise.resolve(resolveImage(value.src)));
      const source = await images.get(value.src);
      if (!safeImageSource(source)) throw new Error('A document image could not be included in this export.');
      value.src = source;
    }
    for (const child of Object.values(value)) if (child && typeof child === 'object') await visit(child);
  }
  await visit(copy);
  return copy;
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
      if (!level || level.format === "bullet") return { text: level?.bulletChar || "•", level: levelIndex, ordered: false };
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
      return { text, level: levelIndex, ordered: true };
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

function runHtml(run) {
  if (run?.style?.hidden) return "";
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
  const content = (paragraph.runs || []).map(runHtml).join("") || "<br>";
  return `<${tag}${styles.length ? ` style="${styles.join(";")}"` : ""}>${markerHtml}${content}</${tag}>`;
}

function blocksHtml(blocks, context) {
  return (blocks || []).map((block) => {
    if (block.kind === "paragraph") return paragraphHtml(block, context);
    if (block.kind === "image") {
      const source = safeImageSource(block.src);
      if (!source) return '<p class="unsupported">[Image not available in this export]</p>';
      const width = safeNumber(block.widthPx, 320, 1, 10_000);
      const height = safeNumber(block.heightPx, 240, 1, 10_000);
      return `<figure style="text-align:${["left", "center", "right"].includes(block.align) ? block.align : "center"}"><img src="${source}" alt="Document image" width="${width}" height="${height}"></figure>`;
    }
    if (block.kind === "table") {
      const rows = (block.rows || []).map((row) => `<tr>${(row.cells || []).map((cell) => {
        const spans = `${cell.colSpan > 1 ? ` colspan="${Math.floor(safeNumber(cell.colSpan, 1, 1, 100))}"` : ""}${cell.rowSpan > 1 ? ` rowspan="${Math.floor(safeNumber(cell.rowSpan, 1, 1, 100))}"` : ""}`;
        const shading = safeColor(cell.shading);
        return `<td${spans}${shading ? ` style="background:${shading}"` : ""}>${blocksHtml(cell.blocks, context)}</td>`;
      }).join("")}</tr>`).join("");
      return `<table>${rows}</table>`;
    }
    if (block.kind === "equation") return `<p class="equation">${escapeHtml(mathText(block.equation?.root) || "[Equation]")}</p>`;
    if (block.kind === "shape") {
      const content = block.text?.blocks?.length ? blocksHtml(block.text.blocks, context) : "[Shape]";
      return `<aside class="shape">${content}</aside>`;
    }
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
  sections.push({ blocks, props: document.section || {} });
  return sections;
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
      if (!blocks?.length) continue;
      const identity = JSON.stringify(blocks);
      if (seen.has(identity)) continue;
      seen.add(identity);
      stories.push({ label: sections.length > 1 ? `Section ${index + 1} ${label.toLowerCase()}` : label, blocks });
    }
  }
  for (const [kind, label] of [['footnotes', 'Footnote'], ['endnotes', 'Endnote']]) {
    let number = 0;
    for (const blocks of Object.values(document[kind] || {})) stories.push({ label: `${label} ${++number}`, blocks });
  }
  return stories;
}

function documentToHtml(document, title = "Untitled document") {
  const context = { lists: createListState(document) };
  const sections = sectionsOf(document);
  const pageRules = sections.map(({props}, index) => `@page section${index}{size:${safeNumber(props.pageWidthPx,816,100,5000)}px ${safeNumber(props.pageHeightPx,1056,100,5000)}px}`).join('');
  const body = sections.map(({blocks, props}, index) => `<section style="page:section${index}${index && sections[index - 1].breakType !== 'continuous' ? ';break-before:page' : ''}">${props.header?.length ? `<header>${blocksHtml(props.header, context)}</header>` : ''}${blocksHtml(blocks, context)}${props.footer?.length ? `<footer>${blocksHtml(props.footer, context)}</footer>` : ''}</section>`).join('\n');
  const supplements = supplementaryStories(document, true).map(({label, blocks}) => `<section class="document-note"><h2>${escapeHtml(label)}</h2>${blocksHtml(blocks, context)}</section>`).join('\n');
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
  :root{font-family:Calibri,"Segoe UI",sans-serif;color:#111;background:#fff}body{max-width:8.5in;margin:0 auto;padding:.75in;line-height:1.35}p{margin:.25em 0}.list-marker{display:inline-block;min-width:1.7em;margin-right:.25em}table{width:100%;border-collapse:collapse;margin:1em 0}td{border:1px solid #888;padding:.35em;vertical-align:top}td>p:first-child{margin-top:0}td>p:last-child{margin-bottom:0}img{max-width:100%;height:auto}figure{margin:1em 0}.equation{text-align:center;font-family:"Cambria Math",serif}.shape{border:1px solid #aaa;padding:.5em;margin:.75em 0}.unsupported{color:#666;font-style:italic}header,footer{color:#555;border-color:#ccc}header{border-bottom:1px solid;margin-bottom:1.5em}footer{border-top:1px solid;margin-top:1.5em}.footnotes{border-top:1px solid #aaa;margin-top:2em;padding-top:1em}@media print{body{max-width:none;padding:0}@page{margin:.75in}}
</style>
<style>${pageRules}.document-note{border-top:1px solid #aaa;margin-top:1.5em;padding-top:.5em}.document-note h2{font-size:1em}</style>
</head>
<body>
<main>${body}</main>${supplements}
</body>
</html>
`;
}

function escapeMarkdown(value) {
  return String(value).replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\\/g, "\\\\").replace(/([`*_{}\[\]()#+.!|>-])/g, "\\$1");
}

function runMarkdown(run) {
  if (run?.style?.hidden) return "";
  let text = escapeMarkdown(run?.style?.equation ? mathText(run.style.equation.root) : String(run?.text || "")).replace(/\v/g, '  \n');
  if (!text) return "";
  if (run?.style?.strikethrough) text = `~~${text}~~`;
  if (run?.style?.bold) text = `**${text}**`;
  if (run?.style?.italic) text = `_${text}_`;
  const link = safeLink(run?.style?.link);
  return link ? `[${text}](${link.replace(/[()\s]/g, (value) => encodeURIComponent(value))})` : text;
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
  for (const {label, blocks} of supplementaryStories(document)) sections.push(`${label}\n${blocks.map((block) => blockPlainText(block, context)).join('\n')}`);
  return `${sections.filter(Boolean).join("\n\n")}\n`;
}

function tableMarkdown(table, context) {
  const rows = (table.rows || []).map((row) => (row.cells || []).map((cell) => escapeMarkdown((cell.blocks || []).map((block) => blockPlainText(block, context)).join(" ").replace(/\s+/g, " ").trim())));
  if (!rows.length) return "";
  const columns = Math.max(1, ...rows.map((row) => row.length));
  const normalize = (row) => [...row, ...Array(Math.max(0, columns - row.length)).fill("")];
  const [first, ...rest] = rows.map(normalize);
  return [`| ${first.join(" | ")} |`, `| ${Array(columns).fill("---").join(" | ")} |`, ...rest.map((row) => `| ${row.join(" | ")} |`)].join("\n");
}

function documentToMarkdown(document, title = "") {
  const context = { lists: createListState(document) };
  const lines = [];
  if (title) lines.push(`# ${escapeMarkdown(title)}`, "");
  for (const block of document.blocks || []) {
    if (block.kind === "paragraph") {
      const level = headingLevel(block);
      const marker = context.lists.marker(block);
      const content = (block.runs || []).map(runMarkdown).join("");
      if (level) lines.push(`${"#".repeat(level)} ${content}`.trimEnd());
      else if (marker) lines.push(`${"  ".repeat(marker.level)}${marker.ordered ? `${marker.text} ` : "- "}${content}`.trimEnd());
      else lines.push(content);
    } else if (block.kind === "table") lines.push(tableMarkdown(block, context));
    else if (block.kind === "image") {
      const source = safeImageSource(block.src);
      lines.push(source ? `![Document image](${source})` : "_[Image not available in this export]_");
    } else if (block.kind === "equation") lines.push(`$${escapeMarkdown(mathText(block.equation?.root) || "Equation")}$`);
    else if (block.kind === "shape") lines.push(block.text?.blocks?.map((child) => blockPlainText(child, context)).join("\n") || "_[Shape]_");
    else lines.push("_[Unsupported document object]_");
    lines.push("");
  }
  for (const {label, blocks} of supplementaryStories(document)) {
    lines.push(`## ${label}`, '', documentToMarkdown({ ...document, blocks, section: {}, footnotes: {}, endnotes: {} }).trimEnd(), '');
  }
  return `${lines.join("\n").replace(/\n{3,}/g, "\n\n").trimEnd()}\n`;
}

function serializeDocument(document, format, title) {
  if (!document || !Array.isArray(document.blocks)) throw new Error("The open document is not ready to export.");
  if (format === "html") return documentToHtml(document, title);
  if (format === "md") return documentToMarkdown(document, title);
  if (format === "txt") return documentToText(document);
  throw new Error(`Unsupported structured document export: ${format}`);
}

export { documentToHtml, documentToMarkdown, documentToText, serializeDocument, prepareDocumentForExport };
