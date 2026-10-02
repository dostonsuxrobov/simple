/**
 * Clipboard paste for Simple Docs (DOC-008, DOC-012, DOC-SIE-25).
 *
 * Pasted HTML (and RTF) goes through the same DOM-to-model mapper as file imports, so
 * lists keep real numbering, tables stay tables, headings keep their style and embedded
 * pictures come along, instead of the engine's own paste parser flattening them all to
 * paragraphs in its default font. The result is applied as ONE ordinary engine
 * transaction (one undo step, tracked as suggestions in Suggesting mode) built from the
 * operations WordCanvas 0.12.0's own paste, Enter and list commands use: insertText,
 * insertRuns, splitParagraph, setParaStyle, insertBlock, setListDefinition,
 * setStylesheet, setRuns, deleteRange, removeBlock and mergeParagraphs.
 *
 * Three ways to paste, like Word's Paste Options:
 *   keep  - Keep source formatting (Ctrl+V). Text the source leaves unformatted takes
 *           the destination's body font, so a web page never pastes in a stray font.
 *   merge - Merge formatting: structure and emphasis (bold, italic, underline, links,
 *           lists, tables, headings) with the destination's font, size and colour.
 *   text  - Keep text only (Ctrl+Shift+V): the plain text, formatted like the caret.
 * A single web address pasted over selected text links that text instead.
 *
 * No DOM or runtime imports: Node tests load this module directly.
 */
import type { DocPosition, DocSelection } from "@forevka/wordcanvas/query";
import { htmlToDocument } from "./html.ts";
import { blocksForInsert } from "./insert.ts";
import { parseMarkup, textContent, findElements } from "./markup.ts";
import type { Block, CharStyle, Document, ImportOptions, ListDefinition, NamedStyle, Paragraph, ParaStyle, Run } from "./model.ts";
import { importRtf } from "./rtf.ts";

export type PasteMode = "keep" | "merge" | "text";

/** What a clipboard (or a drop) offers, as text. */
export interface ClipboardContent {
  html?: string | null;
  text?: string | null;
  rtf?: string | null;
}

export interface PasteFragment {
  /** Blocks to paste. List paragraphs reference `lists`; styled paragraphs `styles`. */
  blocks: Block[];
  lists: Record<string, ListDefinition>;
  /** Named paragraph styles the blocks use (from the importer's stylesheet). */
  styles: NamedStyle[];
  /** Plain-language notes about content that was left out. */
  warnings: string[];
}

export interface FragmentOptions {
  /** Font, size and colour for text the source leaves unformatted (the destination's body text). */
  base?: Partial<Pick<CharStyle, "fontFamily" | "fontSizePx" | "color">>;
  /** Widest picture in px (the destination's text width). */
  maxImageWidthPx?: number;
  /** Pictures that are not data: URLs, resolved from bytes the caller already holds. Never fetch. */
  resolveImage?: ImportOptions["resolveImage"];
  /** Prefix for new block ids; must not occur in the destination document. */
  idPrefix?: string;
  /** The destination's footnotes: pasted notes it lacks become superscript numbers and text. */
  footnotes?: Document["footnotes"];
}

/** The parts of a WordCanvas document the paste reads. */
export interface PasteDocument {
  blocks: Block[];
  section?: Document["section"] | null;
  lists?: Document["lists"] | null;
  footnotes?: Document["footnotes"] | null;
  endnotes?: Record<string, Paragraph[]> | null;
  stylesheet?: Document["stylesheet"] | null;
}

export type PasteContent =
  /** Plain text formatted like the caret; with `link`, one line pasted as a link. */
  | { kind: "text"; text: string; link?: string | null }
  | { kind: "fragment"; fragment: PasteFragment; mode: "keep" | "merge" }
  /** Turns the selected text into a link. */
  | { kind: "link"; url: string };

/** One engine transaction: pass `{ ...transaction, origin }` to the engine editor's dispatch. */
export interface PasteTransaction {
  ops: PasteOp[];
  selectionAfter: DocSelection;
  /** Content the destination cannot hold (for example pictures inside a table cell). */
  warnings: string[];
}

export type PasteOp = { type: string; [key: string]: unknown };

const BANDS = ["header", "footer", "headerFirst", "headerEven", "footerFirst", "footerEven"] as const;
type BandName = (typeof BANDS)[number];

// ---- clipboard content -------------------------------------------------------------

/**
 * A single web or mail address pasted as text, ready to use as a link; null for anything
 * else (several words, several lines, plain words or file paths).
 */
export function pastedLink(text: string | null | undefined): string | null {
  const value = String(text ?? "").trim();
  if (!value || value.length > 2048 || /\s/.test(value)) return null;
  if (/^https?:\/\/[^\s/?#.][^\s]*$/i.test(value)) return value;
  if (/^mailto:[^\s@]+@[^\s@]+\.[^\s@]+$/i.test(value)) return value;
  if (/^www\.[^\s./]+\.[^\s]+$/i.test(value)) return `https://${value}`;
  return null;
}

/** True for text that is itself an address (selected text like this is replaced, not linked). */
export function looksLikeAddress(text: string): boolean {
  const value = text.trim();
  return pastedLink(value) !== null || /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value) || /^(?:[a-z]:\\|\\\\|\/)/i.test(value);
}

/** True when clipboard HTML holds pictures and no text (a "Copy image" from a browser). */
export function imageOnlyHtml(html: string | null | undefined): boolean {
  if (!html || !/<img\b/i.test(html)) return false;
  const root = parseMarkup(html);
  for (const element of findElements(root, (candidate) => ["script", "style", "title", "head", "template"].includes(candidate.name))) {
    element.children = [];
  }
  return textContent(root).replace(/[\s\u00a0\u200b]+/g, "") === "";
}

const paragraphText = (block: { runs?: Run[] } | null | undefined) => (block?.runs ?? []).map((run) => run.text).join("");
const isParagraph = (block: Block | undefined | null): block is Paragraph => !!block && block.kind === "paragraph";

function hasContent(block: Block): boolean {
  if (block.kind !== "paragraph") return true;
  return paragraphText(block).replace(/[\s\u00a0\u200b]+/g, "") !== "" || block.runs.some((run) => run.style.footnoteRef);
}

function collectStyleUse(blocks: Block[], lists: Set<string>, styles: Set<string>) {
  for (const block of blocks) {
    if (block.kind === "paragraph") {
      if (block.style.list) lists.add(block.style.list.listId);
      if (block.style.namedStyle) styles.add(block.style.namedStyle);
    } else if (block.kind === "table") {
      for (const row of block.rows) for (const cell of row.cells) collectStyleUse(cell.blocks, lists, styles);
    }
  }
}

/** The importer's notes, worded for a paste (a clipboard has no "file"). */
function pasteWarnings(warnings: string[]): string[] {
  return warnings.map((warning) => warning
    .replace(/^A picture stored outside the file was not found/, "A picture in the copied content couldn’t be read")
    .replace(/^(\d+) pictures stored outside the file were not found/, "$1 pictures in the copied content couldn’t be read"));
}

/** Shapes an imported document into a paste: notes the target lacks become text; empty edges go. */
function fragmentFromDocument(document: Document, warnings: string[], options: FragmentOptions): PasteFragment | null {
  const blocks = blocksForInsert(document, { lists: document.lists, footnotes: options.footnotes ?? {} });
  // Clipboard markup often wraps the selection in empty paragraphs.
  let start = 0;
  let end = blocks.length;
  while (start < end && !hasContent(blocks[start])) start += 1;
  while (end > start && !hasContent(blocks[end - 1])) end -= 1;
  const kept = blocks.slice(start, end);
  if (!kept.length) return null;
  const listIds = new Set<string>();
  const styleIds = new Set<string>();
  collectStyleUse(kept, listIds, styleIds);
  const lists: Record<string, ListDefinition> = {};
  for (const id of listIds) if (document.lists?.[id]) lists[id] = document.lists[id];
  const styles = (document.stylesheet?.styles ?? []).filter((style) => styleIds.has(style.id));
  return { blocks: kept, lists, styles, warnings: pasteWarnings(warnings) };
}

/**
 * The blocks a clipboard's HTML (else RTF) maps to, or null when it holds no rich
 * content. Never fetches anything: pictures come only from data: URLs and `resolveImage`.
 */
export function fragmentFromClipboard(content: ClipboardContent, options: FragmentOptions = {}): PasteFragment | null {
  const importOptions: ImportOptions = {
    idPrefix: options.idPrefix,
    maxImageWidthPx: options.maxImageWidthPx,
    resolveImage: options.resolveImage,
  };
  const html = content.html?.trim() ? content.html : null;
  if (html) {
    const result = htmlToDocument(html, importOptions, { baseCharStyle: options.base, imagePlaceholders: false });
    return fragmentFromDocument(result.document, result.warnings, options);
  }
  const rtf = content.rtf?.trim() ? content.rtf : null;
  if (rtf && /^\s*\{\\rtf/.test(rtf)) {
    try {
      const result = importRtf(rtf, importOptions);
      return fragmentFromDocument(result.document, result.warnings, options);
    } catch {
      return null;
    }
  }
  return null;
}

/** The plain text of a fragment: paragraphs on their own lines, table cells separated by tabs. */
export function fragmentText(fragment: PasteFragment): string {
  return flattenBlocks(fragment.blocks, null).map(paragraphText).join("\n");
}

const EMPHASIS_KEYS = ["bold", "italic", "underline", "underlineStyle", "strikethrough", "verticalAlign", "link", "footnoteRef", "endnoteRef"] as const;

/**
 * Merge formatting: keeps the structure and the emphasis (bold, italic, underline,
 * strikethrough, super/subscript, links) and takes the font, size and colour from the
 * destination. Headings keep their size.
 */
export function mergeFragment(fragment: PasteFragment, destination: Partial<CharStyle>): PasteFragment {
  const restyle = (run: Run, heading: boolean): Run => {
    const style: Record<string, unknown> = { ...destination };
    const source = run.style as unknown as Record<string, unknown>;
    delete style.sdtPath;
    for (const key of EMPHASIS_KEYS) {
      if (source[key] !== undefined) style[key] = source[key];
      else delete style[key];
    }
    if (heading && run.style.fontSizePx) style.fontSizePx = run.style.fontSizePx;
    return { ...run, style: style as unknown as CharStyle };
  };
  const visit = (blocks: Block[]): Block[] => blocks.map((block) => {
    if (block.kind === "paragraph") {
      const heading = /^(?:Heading\d|Title)$/.test(block.style.namedStyle ?? "");
      return { ...block, runs: block.runs.map((run) => restyle(run, heading)) };
    }
    if (block.kind === "table") {
      return { ...block, rows: block.rows.map((row) => ({ ...row, cells: row.cells.map((cell) => ({ ...cell, blocks: visit(cell.blocks) })) })) };
    }
    return block;
  });
  return { ...fragment, blocks: visit(fragment.blocks) };
}

// ---- pictures ------------------------------------------------------------------------

/** EXIF orientation of a JPEG (1-8; 1 when absent or unreadable). */
export function jpegOrientation(bytes: Uint8Array): number {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return 1;
  let offset = 2;
  while (offset + 4 <= bytes.length) {
    if (bytes[offset] !== 0xff) return 1;
    const marker = bytes[offset + 1];
    if (marker === 0xd9 || marker === 0xda) return 1;
    const length = (bytes[offset + 2] << 8) | bytes[offset + 3];
    if (length < 2) return 1;
    if (marker === 0xe1 && offset + 10 <= bytes.length && String.fromCharCode(...bytes.subarray(offset + 4, offset + 10)) === "Exif\u0000\u0000") {
      const tiff = offset + 10;
      const little = bytes[tiff] === 0x49 && bytes[tiff + 1] === 0x49;
      const read16 = (at: number) => little ? bytes[at] | (bytes[at + 1] << 8) : (bytes[at] << 8) | bytes[at + 1];
      const read32 = (at: number) => little
        ? (bytes[at] | (bytes[at + 1] << 8) | (bytes[at + 2] << 16) | (bytes[at + 3] << 24)) >>> 0
        : ((bytes[at] << 24) | (bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3]) >>> 0;
      if (tiff + 8 > bytes.length) return 1;
      const ifd = tiff + read32(tiff + 4);
      if (ifd + 2 > bytes.length) return 1;
      const count = read16(ifd);
      for (let index = 0; index < count; index += 1) {
        const entry = ifd + 2 + index * 12;
        if (entry + 12 > bytes.length) return 1;
        if (read16(entry) === 0x0112) {
          const value = read16(entry + 8);
          return value >= 1 && value <= 8 ? value : 1;
        }
      }
      return 1;
    }
    offset += 2 + length;
  }
  return 1;
}

// ---- the transaction -------------------------------------------------------------------

interface Location {
  paragraph: Paragraph;
  /** "body", a header/footer band, or "nested" (table cell, note, text box). */
  container: "body" | BandName | "nested";
  /** Sibling blocks for body and bands. */
  siblings: Block[] | null;
  index: number;
}

function locate(doc: PasteDocument, blockId: string): Location | null {
  const top = doc.blocks.findIndex((block) => block.id === blockId);
  if (top >= 0) {
    const block = doc.blocks[top];
    return isParagraph(block) ? { paragraph: block, container: "body", siblings: doc.blocks, index: top } : null;
  }
  const section = (doc.section ?? {}) as Partial<Record<BandName, Block[]>>;
  for (const band of BANDS) {
    const list = section[band];
    if (!Array.isArray(list)) continue;
    const index = list.findIndex((block) => block.id === blockId);
    if (index >= 0) return isParagraph(list[index]) ? { paragraph: list[index] as Paragraph, container: band, siblings: list, index } : null;
  }
  const nested = findNested(doc, blockId);
  return nested ? { paragraph: nested, container: "nested", siblings: null, index: -1 } : null;
}

function findIn(blocks: Block[] | undefined, blockId: string): Paragraph | null {
  for (const block of blocks ?? []) {
    if (block.kind === "paragraph") {
      if (block.id === blockId) return block;
    } else if (block.kind === "table") {
      for (const row of block.rows) for (const cell of row.cells) {
        const found = findIn(cell.blocks, blockId);
        if (found) return found;
      }
    } else if (block.kind === "shape") {
      const found = findIn((block as { text?: { blocks?: Block[] } }).text?.blocks, blockId);
      if (found) return found;
    }
  }
  return null;
}

function findNested(doc: PasteDocument, blockId: string): Paragraph | null {
  const found = findIn(doc.blocks, blockId);
  if (found) return found;
  const section = (doc.section ?? {}) as Partial<Record<BandName, Block[]>>;
  for (const band of BANDS) {
    const inBand = findIn(section[band], blockId);
    if (inBand) return inBand;
  }
  for (const notes of [doc.footnotes, doc.endnotes]) {
    for (const list of Object.values(notes ?? {})) {
      const inNote = findIn(list as Block[], blockId);
      if (inNote) return inNote;
    }
  }
  return null;
}

/** A paragraph that is (or a caret that sits) inside a content control: the engine's own paste handles locks there. */
export function inContentControl(paragraph: Paragraph, offset: number): boolean {
  if ((paragraph as { sdtPath?: unknown[] }).sdtPath?.length) return true;
  let at = 0;
  for (const run of paragraph.runs) {
    const end = at + run.text.length;
    if ((run.style as { sdtPath?: unknown[] }).sdtPath?.length && offset >= at && offset <= end) return true;
    at = end;
  }
  return false;
}

/** Text that is fully hidden (the engine keeps such paragraphs when a selection is deleted). */
const isHiddenParagraph = (block: Block) => block.kind === "paragraph" && block.runs.some((run) => run.text.length > 0) && block.runs.every((run) => run.text.length === 0 || run.style.hidden === true);

const CONTROL_CHARACTERS = /[\u0000-\u0008\u000c\u000e-\u001f\u007f]/g;

/**
 * Clipboard text as lines (CRLF, CR and Unicode separators split lines; tabs and soft
 * breaks stay). A final line break is kept, as in Word: pasting whole lines leaves the
 * text after the caret on its own line.
 */
export function textLines(text: string): string[] {
  return text.replace(/\r\n?|\u2028|\u2029|\u0085/g, "\n").replace(CONTROL_CHARACTERS, "").split("\n");
}

function patchRuns(runs: Run[], start: number, end: number, patch: Partial<CharStyle>): Run[] {
  const out: Run[] = [];
  let at = 0;
  for (const run of runs) {
    const from = at;
    const to = at + run.text.length;
    at = to;
    if (to <= start || from >= end || run.text.length === 0) {
      out.push(run);
      continue;
    }
    const a = Math.max(start, from) - from;
    const b = Math.min(end, to) - from;
    if (a > 0) out.push({ text: run.text.slice(0, a), style: run.style });
    out.push({ text: run.text.slice(a, b), style: { ...run.style, ...patch } });
    if (b < run.text.length) out.push({ text: run.text.slice(b), style: run.style });
  }
  return out;
}

function ordered(doc: PasteDocument, selection: DocSelection): { start: DocPosition; end: DocPosition } | null {
  const { anchor, focus } = selection;
  if (anchor.blockId === focus.blockId) return anchor.offset <= focus.offset ? { start: anchor, end: focus } : { start: focus, end: anchor };
  const a = locate(doc, anchor.blockId);
  const f = locate(doc, focus.blockId);
  if (!a || !f || a.container === "nested" || a.container !== f.container) return null;
  return a.index < f.index ? { start: anchor, end: focus } : { start: focus, end: anchor };
}

/** Plain paragraphs for a destination that cannot hold blocks: lists as marker text, table rows as tab-separated lines. */
function flattenBlocks(blocks: Block[], lists: Record<string, ListDefinition> | null, dropped?: { pictures: number }): Paragraph[] {
  const source: Document = { section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } }, blocks, ...(lists ? { lists } : {}) } as Document;
  const marked = lists ? blocksForInsert(source, { lists: {}, footnotes: {} }) : blocks;
  const out: Paragraph[] = [];
  for (const block of marked) {
    if (block.kind === "paragraph") {
      const style = { ...block.style };
      delete style.list;
      delete style.namedStyle;
      delete style.outlineLevel;
      out.push({ ...block, style });
    } else if (block.kind === "table") {
      for (const row of block.rows) {
        const runs: Run[] = [];
        row.cells.forEach((cell, index) => {
          const cellParagraphs = flattenBlocks(cell.blocks, lists, dropped);
          if (index > 0) runs.push({ text: "\t", style: runs[runs.length - 1]?.style ?? cellParagraphs[0]?.runs[0]?.style ?? ({} as CharStyle) });
          cellParagraphs.forEach((paragraph, at) => {
            if (at > 0) runs.push({ text: " ", style: paragraph.runs[0]?.style ?? ({} as CharStyle) });
            runs.push(...paragraph.runs.filter((run) => run.text.length > 0));
          });
        });
        const first = row.cells[0]?.blocks.find(isParagraph);
        out.push({ kind: "paragraph", id: `${block.id}-r${out.length}`, revision: 0, runs: runs.length ? runs : [{ text: "", style: first?.runs[0]?.style ?? ({} as CharStyle) }], style: first ? { ...first.style } : ({} as ParaStyle) });
      }
    } else if (block.kind === "image" && dropped) {
      dropped.pictures += 1;
    }
  }
  return out;
}

const STRUCTURE_STYLE = /^(?:Heading\d|Title|Subtitle|Quote|Code)$/;
/**
 * Styles WordCanvas 0.12.0 supplies itself to a document without a stylesheet (its
 * default stylesheet); other named styles cannot be added to such a document.
 */
const ENGINE_DEFAULT_STYLES = new Set(["Normal", "Title", "Subtitle", "Heading1", "Heading2", "Quote", "Code"]);

const hasStructureStyle = (paragraph: Paragraph) => Boolean(paragraph.style.list) || STRUCTURE_STYLE.test(paragraph.style.namedStyle ?? "") || paragraph.style.outlineLevel !== undefined;

/** Several blocks with lists, tables, pictures or headings paste as blocks; anything else merges into the line. */
export function isStructured(blocks: Block[]): boolean {
  if (blocks.length === 1 && isParagraph(blocks[0])) return false;
  return blocks.some((block) => block.kind !== "paragraph" || hasStructureStyle(block));
}

/**
 * The definitions a fragment needs in the destination (its lists, its missing named
 * styles) as operations, and its blocks without references the destination cannot hold.
 */
function definitionsFor(doc: PasteDocument, fragment: PasteFragment): { definitions: PasteOp[]; blocks: Block[] } {
  const definitions: PasteOp[] = [];
  for (const [id, definition] of Object.entries(fragment.lists)) {
    if (!doc.lists?.[id]) definitions.push({ type: "setListDefinition", listId: id, def: definition });
  }
  const sheet = doc.stylesheet;
  let blocks = fragment.blocks;
  if (sheet) {
    const missing = fragment.styles.filter((style) => !sheet.styles.some((existing) => existing.id === style.id));
    if (missing.length) definitions.push({ type: "setStylesheet", stylesheet: { ...sheet, styles: [...sheet.styles, ...missing] } });
  } else {
    const unknown = new Set(fragment.styles.map((style) => style.id).filter((id) => !ENGINE_DEFAULT_STYLES.has(id)));
    if (unknown.size) blocks = withoutStyles(blocks, unknown);
  }
  return { definitions, blocks };
}

/** The paragraph-style patch that gives a split-off paragraph exactly `style`. */
function stylePatch(current: ParaStyle, style: ParaStyle): Record<string, unknown> {
  const patch: Record<string, unknown> = { ...style };
  for (const key of Object.keys(current)) if (!(key in patch)) patch[key] = undefined;
  return patch;
}

const runsLength = (runs: Run[]) => runs.reduce((total, run) => total + run.text.length, 0);
const nonEmptyRuns = (runs: Run[]) => runs.filter((run) => run.text.length > 0);

/**
 * Builds the paste as one transaction, or returns null when this module cannot place
 * it (a selection across stories, a caret in a content control, or nothing selected);
 * the caller then leaves the paste to the engine.
 */
export function buildPasteTransaction(doc: PasteDocument, selection: DocSelection | null, content: PasteContent, newId: () => string): PasteTransaction | null {
  if (!selection) return null;
  const range = ordered(doc, selection);
  if (!range) return null;
  const { start, end } = range;
  const startAt = locate(doc, start.blockId);
  const endAt = locate(doc, end.blockId);
  if (!startAt || !endAt) return null;
  const startText = paragraphText(startAt.paragraph);
  const endText = paragraphText(endAt.paragraph);
  if (start.offset < 0 || start.offset > startText.length || end.offset < 0 || end.offset > endText.length) return null;
  if (inContentControl(startAt.paragraph, start.offset) || inContentControl(endAt.paragraph, end.offset)) return null;
  const collapsed = start.blockId === end.blockId && start.offset === end.offset;
  const warnings: string[] = [];

  if (content.kind === "link") {
    if (collapsed) return null;
    const url = pastedLink(content.url);
    if (!url) return null;
    const ops: PasteOp[] = [];
    if (start.blockId === end.blockId) {
      ops.push({ type: "setRuns", blockId: start.blockId, runs: patchRuns(startAt.paragraph.runs, start.offset, end.offset, { link: url }) });
    } else {
      const siblings = startAt.siblings!;
      for (let index = startAt.index; index <= endAt.index; index += 1) {
        const block = siblings[index];
        if (!isParagraph(block)) continue;
        const length = paragraphText(block).length;
        const from = index === startAt.index ? start.offset : 0;
        const to = index === endAt.index ? end.offset : length;
        if (to > from) ops.push({ type: "setRuns", blockId: block.id, runs: patchRuns(block.runs, from, to, { link: url }) });
      }
    }
    return ops.length ? { ops, selectionAfter: { anchor: end, focus: end }, warnings } : null;
  }

  // 1. Remove the selection, the way the engine's own paste does.
  const ops: PasteOp[] = [];
  const before = startText.slice(0, start.offset);
  if (!collapsed) {
    if (start.blockId === end.blockId) {
      ops.push({ type: "deleteRange", blockId: start.blockId, start: start.offset, end: end.offset });
    } else {
      const siblings = startAt.siblings!;
      if (start.offset < startText.length) ops.push({ type: "deleteRange", blockId: start.blockId, start: start.offset, end: startText.length });
      if (end.offset > 0) ops.push({ type: "deleteRange", blockId: end.blockId, start: 0, end: end.offset });
      for (let index = startAt.index + 1; index < endAt.index; index += 1) {
        // A fully hidden paragraph survives, and the engine then keeps both ends apart.
        if (isHiddenParagraph(siblings[index])) return null;
        ops.push({ type: "removeBlock", blockId: siblings[index].id });
      }
      ops.push({ type: "mergeParagraphs", firstBlockId: start.blockId });
    }
  }
  const at: DocPosition = { blockId: start.blockId, offset: start.offset };
  const sdt = (startAt.paragraph as { sdtPath?: string[] }).sdtPath;
  const split = (position: DocPosition, newBlockId: string, extra: Record<string, unknown> = {}): PasteOp => ({ type: "splitParagraph", at: position, newBlockId, ...(sdt?.length ? { newSdtPath: sdt } : {}), ...extra });

  // 2a. Plain text: one paragraph per line, formatted like the caret.
  if (content.kind === "text") {
    const lines = textLines(content.text);
    if (lines.length === 1 && !lines[0]) return ops.length ? { ops, selectionAfter: { anchor: at, focus: at }, warnings } : null;
    let position = at;
    lines.forEach((line, index) => {
      if (index > 0) {
        const id = newId();
        ops.push(split(position, id));
        position = { blockId: id, offset: 0 };
      }
      if (line.length) {
        const op: PasteOp = { type: "insertText", at: position, text: line };
        if (content.link && lines.length === 1) op.style = { ...(runStyleAt(startAt.paragraph, start.offset) ?? {}), link: content.link };
        ops.push(op);
        position = { blockId: position.blockId, offset: position.offset + line.length };
      }
    });
    return { ops, selectionAfter: { anchor: position, focus: position }, warnings };
  }

  // 2b. A fragment.
  const after = collapsed ? startText.slice(start.offset) : start.blockId === end.blockId ? startText.slice(end.offset) : endText.slice(end.offset);
  const { definitions, blocks: defined } = definitionsFor(doc, content.fragment);
  let blocks = defined;
  const canHoldBlocks = startAt.container !== "nested";
  const single = blocks.length === 1 && isParagraph(blocks[0]) ? blocks[0] : null;
  if (single && !before && !after && hasStructureStyle(single)) {
    // One list item or heading pasted on an empty line becomes that kind of paragraph.
    const runs = nonEmptyRuns(single.runs);
    if (runs.length) ops.push({ type: "insertRuns", at, runs });
    ops.push({ type: "setParaStyle", blockId: at.blockId, patch: stylePatch(startAt.paragraph.style, single.style) });
    const caret = { blockId: at.blockId, offset: runsLength(runs) };
    return { ops: [...definitions, ...ops], selectionAfter: { anchor: caret, focus: caret }, warnings };
  }
  if (isStructured(blocks) && canHoldBlocks) {
    const where = startAt.container === "body" ? {} : { where: startAt.container };
    const index = startAt.index;
    const inserts: PasteOp[] = [];
    let selectionAfter: DocPosition;
    if (before.length === 0) {
      // At the start of a line: the blocks go before it and the caret stays at its start.
      blocks.forEach((block, offset) => inserts.push({ type: "insertBlock", index: index + offset, block, ...where }));
      selectionAfter = at;
    } else {
      const tail = newId();
      inserts.push(split(at, tail));
      blocks.forEach((block, offset) => inserts.push({ type: "insertBlock", index: index + 1 + offset, block, ...where }));
      selectionAfter = { blockId: tail, offset: 0 };
    }
    return { ops: [...definitions, ...ops, ...inserts], selectionAfter: { anchor: selectionAfter, focus: selectionAfter }, warnings };
  }
  if (isStructured(blocks)) {
    // Table cells, notes and text boxes hold paragraphs only here.
    const dropped = { pictures: 0 };
    blocks = flattenBlocks(blocks, content.fragment.lists, dropped);
    if (dropped.pictures) warnings.push(dropped.pictures === 1 ? "The picture wasn't pasted: pictures can't be pasted into a table cell, note or text box yet." : `${dropped.pictures} pictures weren't pasted: pictures can't be pasted into a table cell, note or text box yet.`);
    if (!blocks.length) return null;
  }
  const paragraphs = blocks.filter(isParagraph);
  if (!paragraphs.length) return null;
  const first = paragraphs[0];
  if (paragraphs.length === 1) {
    const runs = nonEmptyRuns(first.runs);
    if (!runs.length) return ops.length ? { ops, selectionAfter: { anchor: at, focus: at }, warnings } : null;
    ops.push({ type: "insertRuns", at, runs });
    const caret = { blockId: at.blockId, offset: at.offset + runsLength(runs) };
    return { ops, selectionAfter: { anchor: caret, focus: caret }, warnings };
  }
  // Several paragraphs: the first joins the text before the caret, the last the text
  // after it, and the ones between become paragraphs of their own (as the engine does).
  const keepStyles = content.mode === "keep";
  let tail = newId();
  ops.push(split(at, tail));
  const firstRuns = nonEmptyRuns(first.runs);
  if (firstRuns.length) ops.push({ type: "insertRuns", at, runs: firstRuns });
  for (let index = 1; index < paragraphs.length; index += 1) {
    const paragraph = paragraphs[index];
    const runs = nonEmptyRuns(paragraph.runs);
    if (index === paragraphs.length - 1) {
      if (runs.length) ops.push({ type: "insertRuns", at: { blockId: tail, offset: 0 }, runs });
      const caret = { blockId: tail, offset: runsLength(runs) };
      return { ops, selectionAfter: { anchor: caret, focus: caret }, warnings };
    }
    const next = newId();
    ops.push(split({ blockId: tail, offset: 0 }, next));
    if (runs.length) ops.push({ type: "insertRuns", at: { blockId: tail, offset: 0 }, runs });
    if (keepStyles) ops.push({ type: "setParaStyle", blockId: tail, patch: stylePatch(startAt.paragraph.style, paragraph.style) });
    tail = next;
  }
  return null;
}

/** The character style at an offset (the run before it, like typing does). */
function runStyleAt(paragraph: Paragraph, offset: number): CharStyle | null {
  let at = 0;
  let found: CharStyle | null = paragraph.runs[0]?.style ?? null;
  for (const run of paragraph.runs) {
    if (offset > at || (offset === 0 && at === 0)) found = run.style;
    at += run.text.length;
    if (at >= offset) break;
  }
  return found ? { ...found } : null;
}

function withoutStyles(blocks: Block[], ids: Set<string>): Block[] {
  return blocks.map((block) => {
    if (block.kind === "paragraph" && block.style.namedStyle && ids.has(block.style.namedStyle)) {
      const style = { ...block.style };
      delete style.namedStyle;
      return { ...block, style };
    }
    if (block.kind === "table") return { ...block, rows: block.rows.map((row) => ({ ...row, cells: row.cells.map((cell) => ({ ...cell, blocks: withoutStyles(cell.blocks, ids) })) })) };
    return block;
  });
}
