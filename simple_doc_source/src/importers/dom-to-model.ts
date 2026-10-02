/**
 * The one DOM-to-model mapper behind every importer. HTML, Markdown (through HTML),
 * RTF and ODT all arrive here as a plain node tree (see markup.ts) and leave as a
 * WordCanvas document: paragraphs and runs with bold/italic/underline/strike, colour,
 * highlight, size and font; headings 1-6; bulleted and numbered lists with nesting;
 * tables with merged cells; pictures; links; footnotes; page breaks.
 *
 * Structure hints that HTML cannot express are passed by the trusted front-ends (RTF,
 * ODT, Markdown) as data-simple-* attributes; they are ignored in untrusted HTML:
 *   data-simple-list="key" data-simple-level="n"  list membership (definitions in options.lists)
 *   data-simple-footnote="key"                     footnote reference (the element's text is replaced)
 *   data-simple-footnote-body="key"                footnote body (moved out of the text flow)
 *   data-simple-band="header" | "footer"           page header/footer story
 */
import {
  assembleDocument, createIdSource, defaultCharStyle, defaultParaStyle, DEFAULT_SPACE_AFTER_PX, headingCharStyle, headingParaStyle,
  LINK_COLOR, listDefinition, listLevel, MONOSPACE_FONT, sameCharStyle,
} from "./model.ts";
import type { Block, CharStyle, Document, ImageBlockModel, ImportOptions, ListDefinition, ListNumberFormat, Paragraph, ParaStyle, Run, TableBlock, TableCell } from "./model.ts";
import { boxValues, computeDeclarations, firstFontFamily, parseColor, parseFontSize, parseLength, parseStyleSheet } from "./css.ts";
import type { Declarations, StyleSheet } from "./css.ts";
import { findElements, hasClass, textContent } from "./markup.ts";
import type { DomElement, DomNode } from "./markup.ts";
import { bytesToDataUrl, decodeImageDataUrl, imageNaturalSize, MAX_IMAGE_BYTES, sniffImageType } from "./images.ts";
import type { ImageBytes } from "./images.ts";

export interface ExplicitListLevel {
  format: ListNumberFormat;
  start?: number;
  bulletChar?: string;
  /** Number pattern such as "%1." or "(%1)". */
  text?: string;
}

export interface MapOptions extends ImportOptions {
  /** Honour data-simple-* structure attributes (front-ends that build the tree). */
  trusted?: boolean;
  /** Level definitions for data-simple-list keys. */
  lists?: Record<string, ExplicitListLevel[]>;
  /** Page geometry for the document. */
  section?: Partial<Document["section"]>;
  /** Read <style> rules (tag, .class, tag.class). Default true. */
  authorStyles?: boolean;
  /**
   * Font, size and colour of text the source leaves unformatted, headings included
   * (default: Simple's body text). Pasting passes the destination document's.
   */
  baseCharStyle?: Partial<Pick<CharStyle, "fontFamily" | "fontSizePx" | "color">>;
  /** Show a "[Picture]" placeholder for a picture that cannot be read (default true); otherwise it is left out and only counted in the warnings. */
  imagePlaceholders?: boolean;
}

export interface MapResult {
  document: Document;
  warnings: string[];
}

const MAX_DEPTH = 200;
const pt = (value: number) => value * 96 / 72;
const PLACEHOLDER_COLOR = "#6b6b6b";

const SKIPPED = new Set([
  "head", "script", "style", "title", "meta", "link", "base", "noscript", "template", "iframe", "frame", "frameset", "object", "embed", "applet",
  "param", "svg", "math", "canvas", "video", "audio", "source", "track", "map", "area", "select", "option", "optgroup", "textarea", "button",
  "datalist", "output", "progress", "meter", "xml", "picture-source", "noembed", "noframes", "dialog",
]);
const BLOCKS = new Set([
  "p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "address", "article", "aside", "section", "header", "footer", "main",
  "nav", "figure", "figcaption", "center", "dl", "dt", "dd", "details", "summary", "fieldset", "legend", "form", "hgroup", "listing", "xmp",
  "plaintext", "body", "html", "caption",
]);
// Blocks that stand for one paragraph even when empty.
const PARAGRAPH_BLOCKS = new Set(["p", "h1", "h2", "h3", "h4", "h5", "h6", "dt", "dd"]);
const BULLET_MARKERS = /^[•·◦▪■□○●‣⁃∙\-–—*o§Øü✓✔➢➤►▸→]$/;

interface ListItemState {
  listId: string;
  level: number;
  used: boolean;
}

interface Ctx {
  char: CharStyle;
  para: ParaStyle;
  /** Spacing a new block starts from (body text vs. table cell vs. note). */
  spacing: { before: number; after: number };
  pre: boolean;
  item: ListItemState | null;
  listId: string | null;
  listLevel: number;
  listMarkers: boolean;
  depth: number;
}

interface OpenParagraph {
  style: ParaStyle;
  runs: Run[];
  char: CharStyle;
  spacePending: boolean;
  pre: boolean;
}

class Story {
  blocks: Block[] = [];
  open: OpenParagraph | null = null;
  produced = 0;
}

const clampInt = (value: string | undefined, min: number, max: number): number | null => {
  const number = Number.parseInt(value ?? "", 10);
  return Number.isFinite(number) ? Math.min(max, Math.max(min, number)) : null;
};

/** Link targets that may open from a document (never javascript:, data:, file: or relative paths). */
export function safeHref(value: string | undefined): string | null {
  const href = (value ?? "").replace(/[\u0000-\u001f\u007f]/g, "").trim();
  if (/^(?:https?|mailto|tel|ftp):/i.test(href) && href.length <= 4096) return href;
  return null;
}

/** List format from a marker such as "1.", "b)", "iv." or "•". */
export function markerFormat(marker: string): { format: ListNumberFormat; start: number; bulletChar?: string } {
  const text = marker.replace(/[\s\u00a0]+/g, "");
  if (!text || BULLET_MARKERS.test(text)) return { format: "bullet", start: 1, bulletChar: text === "o" ? "◦" : text === "§" ? "▪" : !text || text === "·" || text === "∙" ? "•" : text };
  let match = /^\(?(\d{1,6})[.)]?$/.exec(text);
  if (match) return { format: "decimal", start: Number(match[1]) };
  match = /^\(?([ivxlcdm]{2,8}|i)[.)]?$/.exec(text);
  if (match) return { format: "lowerRoman", start: romanValue(match[1]) };
  match = /^\(?([IVXLCDM]{2,8}|I)[.)]?$/.exec(text);
  if (match) return { format: "upperRoman", start: romanValue(match[1]) };
  match = /^\(?([a-z])[.)]?$/.exec(text);
  if (match) return { format: "lowerLetter", start: match[1].charCodeAt(0) - 96 };
  match = /^\(?([A-Z])[.)]?$/.exec(text);
  if (match) return { format: "upperLetter", start: match[1].charCodeAt(0) - 64 };
  return { format: "bullet", start: 1, bulletChar: "•" };
}

function romanValue(text: string): number {
  const values: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  let total = 0;
  const lower = text.toLowerCase();
  for (let i = 0; i < lower.length; i += 1) {
    const value = values[lower[i]] ?? 0;
    const next = values[lower[i + 1]] ?? 0;
    total += value < next ? -value : value;
  }
  return Math.max(1, total);
}

function htmlListFormat(element: DomElement, css: Declarations, ordered: boolean): { format: ListNumberFormat; bulletChar?: string } | null {
  const type = (css["list-style-type"] ?? css["list-style"] ?? "").toLowerCase();
  if (/\bnone\b/.test(type)) return null;
  if (/decimal/.test(type)) return { format: "decimal" };
  if (/lower-(?:alpha|latin)/.test(type)) return { format: "lowerLetter" };
  if (/upper-(?:alpha|latin)/.test(type)) return { format: "upperLetter" };
  if (/lower-roman/.test(type)) return { format: "lowerRoman" };
  if (/upper-roman/.test(type)) return { format: "upperRoman" };
  if (/\bdisc\b/.test(type)) return { format: "bullet", bulletChar: "•" };
  if (/\bcircle\b/.test(type)) return { format: "bullet", bulletChar: "◦" };
  if (/\bsquare\b/.test(type)) return { format: "bullet", bulletChar: "▪" };
  const attr = element.attrs.type;
  if (ordered && attr) {
    if (attr === "a") return { format: "lowerLetter" };
    if (attr === "A") return { format: "upperLetter" };
    if (attr === "i") return { format: "lowerRoman" };
    if (attr === "I") return { format: "upperRoman" };
    return { format: "decimal" };
  }
  if (!ordered && attr) {
    const bullet = attr.toLowerCase();
    if (bullet === "circle") return { format: "bullet", bulletChar: "◦" };
    if (bullet === "square") return { format: "bullet", bulletChar: "▪" };
  }
  return null;
}

function breaksBefore(css: Declarations): boolean {
  return /^(?:page|always|left|right|recto|verso)$/i.test(css["break-before"] ?? css["page-break-before"] ?? "");
}

function breaksAfter(css: Declarations): boolean {
  return /^(?:page|always|left|right|recto|verso)$/i.test(css["break-after"] ?? css["page-break-after"] ?? "");
}

function backgroundColor(css: Declarations): string | null {
  const direct = css["background-color"];
  if (direct) return parseColor(direct);
  const shorthand = css.background;
  if (!shorthand || /url\(/i.test(shorthand)) return null;
  for (const part of shorthand.split(/\s+(?![^(]*\))/)) {
    const color = parseColor(part);
    if (color) return color;
  }
  return null;
}

class Mapper {
  readonly warnings: string[] = [];
  private readonly ids: (kind?: string) => string;
  private readonly trusted: boolean;
  private readonly sheet: StyleSheet | null;
  private readonly declarationCache = new Map<DomElement, Declarations>();
  private readonly stories: Story[] = [new Story()];
  private readonly lists: Record<string, ListDefinition> = {};
  private readonly touchedLevels = new Map<string, Set<number>>();
  private readonly explicitLists = new Map<string, string>();
  private readonly markerLists = new Map<string, string>();
  private markerGroup: { listId: string; story: Story; blockCount: number; ordered: boolean; next: number } | null = null;
  private lastOrderedList: { listId: string; next: number; story: Story } | null = null;
  private readonly topLevelCounts = new Map<string, number>();
  private readonly skipNodes = new Set<DomElement>();
  private pendingBreak = false;
  private readonly maxImageWidth: number;
  // Footnotes.
  private readonly noteBodies = new Map<string, DomElement>();
  private readonly noteBodyKey = new Map<DomElement, string>();
  private readonly noteContainers = new Set<DomElement>();
  private readonly noteRefs = new Map<DomElement, string>();
  private readonly referencedNotes = new Set<string>();
  private readonly refAnchorIds = new Set<string>();
  private readonly noteIds = new Map<string, string>();
  private readonly pendingNotes: Array<{ key: string; id: string }> = [];
  private readonly footnotes: Record<string, Paragraph[]> = {};
  private inNote = false;
  private noteCount = 0;
  private readonly bands: Partial<Record<"header" | "footer", Block[]>> = {};
  private counts = { remoteImages: 0, localImages: 0, badImages: 0, nestedTooDeep: 0, droppedNoteObjects: 0 };

  private readonly root: DomElement;
  private readonly options: MapOptions;
  private readonly noteNumbers = new Map<string, number>();

  constructor(root: DomElement, options: MapOptions) {
    this.root = root;
    this.options = options;
    this.ids = createIdSource(options.idPrefix);
    this.trusted = options.trusted === true;
    this.maxImageWidth = Math.max(16, options.maxImageWidthPx ?? 624);
    if (options.authorStyles !== false) {
      const css = findElements(root, (element) => element.name === "style").map((element) => textContent(element)).join("\n");
      this.sheet = css.trim() ? parseStyleSheet(css) : null;
    } else this.sheet = null;
    this.collectNotes();
  }

  private get story(): Story {
    return this.stories[this.stories.length - 1];
  }

  private css(element: DomElement): Declarations {
    let cached = this.declarationCache.get(element);
    if (!cached) {
      cached = computeDeclarations(element, this.sheet);
      this.declarationCache.set(element, cached);
    }
    return cached;
  }

  // ---- footnotes -------------------------------------------------------------

  private collectNotes() {
    const addBody = (key: string, element: DomElement) => {
      if (!key || this.noteBodies.has(key)) return;
      this.noteBodies.set(key, element);
      this.noteBodyKey.set(element, key);
    };
    if (this.trusted) {
      for (const element of findElements(this.root, (candidate) => candidate.attrs["data-simple-footnote-body"] !== undefined)) addBody(element.attrs["data-simple-footnote-body"], element);
    }
    // Footnote sections written by Simple, markdown-it-footnote, pandoc and most generators.
    const containers = findElements(this.root, (element) => hasClass(element, "footnotes") || /^doc-(?:end|foot)notes$/.test(element.attrs.role ?? "") || element.attrs.id === "footnotes");
    for (const container of containers) {
      this.noteContainers.add(container);
      for (const item of findElements(container, (element) => element.name === "li" && !!element.attrs.id)) addBody(item.attrs.id, item);
    }
    // Word's "Web Page" footnotes: <div style="mso-element:footnote" id="ftn1"><a name="_ftn1" ...>
    for (const element of findElements(this.root, (candidate) => /footnote|endnote/i.test(this.css(candidate)["mso-element"] ?? ""))) {
      if (element.attrs.id) addBody(element.attrs.id, element);
      for (const anchor of findElements(element, (candidate) => candidate.name === "a")) {
        if (anchor.attrs.name) addBody(anchor.attrs.name, element);
        if (anchor.attrs.id) addBody(anchor.attrs.id, element);
      }
    }
    const insideBody = (element: DomElement) => {
      for (let node: DomElement | null = element; node; node = node.parent) if (this.noteBodyKey.has(node)) return true;
      return false;
    };
    if (this.trusted) {
      for (const element of findElements(this.root, (candidate) => candidate.attrs["data-simple-footnote"] !== undefined)) {
        const key = element.attrs["data-simple-footnote"];
        if (!this.noteBodies.has(key)) continue;
        this.noteRefs.set(element, key);
        this.referencedNotes.add(key);
      }
    }
    if (!this.noteBodies.size) return;
    for (const anchor of findElements(this.root, (candidate) => candidate.name === "a" && (candidate.attrs.href ?? "").startsWith("#"))) {
      const key = decodeURIComponent(anchor.attrs.href.slice(1));
      if (!this.noteBodies.has(key) || insideBody(anchor) || this.referencedNotes.has(key)) continue;
      this.noteRefs.set(anchor, key);
      this.referencedNotes.add(key);
      if (anchor.attrs.id) this.refAnchorIds.add(anchor.attrs.id);
      if (anchor.attrs.name) this.refAnchorIds.add(anchor.attrs.name);
    }
  }

  private isReferencedBody(element: DomElement): boolean {
    const key = this.noteBodyKey.get(element);
    if (key !== undefined && this.referencedNotes.has(key) && this.noteBodies.get(key) === element) return true;
    if (this.noteContainers.has(element)) {
      const bodies = findElements(element, (candidate) => this.noteBodyKey.has(candidate));
      return bodies.length > 0 && bodies.every((body) => this.referencedNotes.has(this.noteBodyKey.get(body)!));
    }
    return false;
  }

  private isBackLink(element: DomElement): boolean {
    if (element.name !== "a") return false;
    if (hasClass(element, "note-back") || hasClass(element, "footnote-backref") || element.attrs.role === "doc-backlink") return true;
    const href = element.attrs.href ?? "";
    return href.startsWith("#") && this.refAnchorIds.has(decodeURIComponent(href.slice(1)));
  }

  private footnoteRef(key: string, ctx: Ctx) {
    const known = this.noteIds.get(key);
    const open = this.ensure(ctx);
    const style: CharStyle = { ...ctx.char, verticalAlign: "super" };
    delete style.link;
    if (known) {
      // A second reference to the same note: Word cannot point twice at one note.
      this.appendRun(open, String(this.noteNumbers.get(known) ?? ""), style);
      return;
    }
    const id = this.ids("fn");
    this.noteCount += 1;
    this.noteIds.set(key, id);
    this.noteNumbers.set(id, this.noteCount);
    this.pendingNotes.push({ key, id });
    this.appendRun(open, String(this.noteCount), { ...style, footnoteRef: id });
    open.spacePending = false;
  }

  private convertNotes() {
    for (let index = 0; index < this.pendingNotes.length; index += 1) {
      const { key, id } = this.pendingNotes[index];
      const body = this.noteBodies.get(key);
      const story = new Story();
      this.stories.push(story);
      this.inNote = true;
      const ctx = this.baseContext({ before: 0, after: 0 });
      ctx.char.fontSizePx = Math.round(pt(10) * 1000) / 1000;
      ctx.para.lineHeight = 1;
      ctx.para.spaceAfterPx = 0;
      if (body) this.walkChildren(body, ctx);
      this.close();
      this.inNote = false;
      this.stories.pop();
      const paragraphs = story.blocks.filter((block): block is Paragraph => block.kind === "paragraph");
      if (paragraphs.length !== story.blocks.length) this.counts.droppedNoteObjects += story.blocks.length - paragraphs.length;
      // Drop leading space and a "[1]" marker a note body may start with.
      const first = paragraphs[0]?.runs[0];
      if (first) first.text = first.text.replace(/^\s*(?:\[\d+\]\s*)?/, "");
      this.footnotes[id] = paragraphs.length ? paragraphs : [this.emptyParagraph(ctx)];
    }
  }

  // ---- stories and paragraphs ----------------------------------------------------

  baseContext(spacing = { before: 0, after: DEFAULT_SPACE_AFTER_PX }): Ctx {
    const para = defaultParaStyle();
    para.spaceBeforePx = spacing.before;
    para.spaceAfterPx = spacing.after;
    return { char: { ...defaultCharStyle(), ...this.baseChar() }, para, spacing, pre: false, item: null, listId: null, listLevel: -1, listMarkers: true, depth: 0 };
  }

  /** The caller's base font, size and colour (only those three). */
  private baseChar(): Partial<CharStyle> {
    const base = this.options.baseCharStyle;
    const char: Partial<CharStyle> = {};
    if (typeof base?.fontFamily === "string" && base.fontFamily.trim()) char.fontFamily = base.fontFamily.trim();
    if (typeof base?.fontSizePx === "number" && base.fontSizePx > 0 && base.fontSizePx < 400) char.fontSizePx = base.fontSizePx;
    if (typeof base?.color === "string" && base.color) char.color = base.color;
    return char;
  }

  private ensure(ctx: Ctx): OpenParagraph {
    const story = this.story;
    if (story.open) return story.open;
    const style: ParaStyle = { ...ctx.para };
    if (ctx.para.borders) style.borders = { ...ctx.para.borders };
    if (this.pendingBreak && story === this.stories[0]) {
      style.pageBreakBefore = true;
      this.pendingBreak = false;
    }
    if (ctx.item) {
      if (!ctx.item.used) {
        style.list = { listId: ctx.item.listId, level: ctx.item.level };
        style.contextualSpacing = true;
        ctx.item.used = true;
        if (ctx.item.level === 0) this.topLevelCounts.set(ctx.item.listId, (this.topLevelCounts.get(ctx.item.listId) ?? 0) + 1);
      } else {
        // A further paragraph of the same item lines up with the item's text.
        style.indentLeftPx = (style.indentLeftPx ?? 0) + 48 * (ctx.item.level + 1);
        style.contextualSpacing = true;
      }
    }
    story.open = { style, runs: [], char: { ...ctx.char }, spacePending: true, pre: ctx.pre };
    story.produced += 1;
    return story.open;
  }

  private appendRun(open: OpenParagraph, text: string, style: CharStyle) {
    if (!text) return;
    const last = open.runs[open.runs.length - 1];
    if (last && sameCharStyle(last.style, style)) last.text += text;
    else open.runs.push({ text, style: { ...style } });
  }

  private close() {
    const story = this.story;
    const open = story.open;
    if (!open) return;
    story.open = null;
    const runs = open.runs;
    // A trailing <br> ends a line, it does not add one.
    const last = runs[runs.length - 1];
    if (last && last.text.endsWith("\v")) last.text = last.text.slice(0, -1);
    if (!open.pre) {
      for (let i = runs.length - 1; i >= 0; i -= 1) {
        runs[i].text = runs[i].text.replace(/ +$/, "");
        if (runs[i].text) break;
      }
    }
    const kept = runs.filter((run) => run.text.length > 0);
    story.blocks.push({ kind: "paragraph", id: this.ids("p"), revision: 0, runs: kept.length ? kept : [{ text: "", style: open.char }], style: open.style });
  }

  /** Closes the open paragraph for an object (picture, table); drops it when it is empty. */
  private closeForObject() {
    const open = this.story.open;
    if (open && !open.style.list && !open.style.pageBreakBefore && open.runs.every((run) => !run.text.replace(/[ \v]/g, ""))) {
      this.story.open = null;
      return;
    }
    this.close();
  }

  private pushBlock(block: Block, ctx: Ctx) {
    this.closeForObject();
    const story = this.story;
    if (this.pendingBreak && story === this.stories[0]) {
      const breakParagraph = this.emptyParagraph(ctx);
      breakParagraph.style.pageBreakBefore = true;
      breakParagraph.style.spaceAfterPx = 0;
      story.blocks.push(breakParagraph);
      this.pendingBreak = false;
    }
    story.blocks.push(block);
    story.produced += 1;
  }

  private emptyParagraph(ctx: Ctx): Paragraph {
    return { kind: "paragraph", id: this.ids("p"), revision: 0, runs: [{ text: "", style: { ...ctx.char } }], style: { ...ctx.para } };
  }

  // ---- walking -------------------------------------------------------------------

  walkChildren(element: DomElement, ctx: Ctx) {
    for (const child of element.children) this.walk(child, ctx);
  }

  private walk(node: DomNode, ctx: Ctx) {
    if (node.type === "text") this.text(node.text, ctx, node.pre === true);
    else this.element(node, ctx);
  }

  private text(raw: string, ctx: Ctx, preserve: boolean) {
    if (!raw) return;
    if (ctx.pre) {
      const lines = raw.replace(/\r\n?/g, "\n").split("\n");
      lines.forEach((line, index) => {
        if (index > 0) {
          this.ensure(ctx);
          this.close();
        }
        if (line) this.appendRun(this.ensure(ctx), line, ctx.char);
      });
      return;
    }
    let text = preserve ? raw.replace(/[\r\n]+/g, " ") : raw.replace(/[ \t\n\r\f]+/g, " ");
    if (text === " " && !this.story.open) return;
    const open = this.ensure(ctx);
    if (open.spacePending && !preserve) text = text.replace(/^ /, "");
    if (!text) return;
    this.appendRun(open, text, ctx.char);
    open.spacePending = !preserve && text.endsWith(" ");
  }

  private element(element: DomElement, ctx: Ctx) {
    const name = element.name;
    if (SKIPPED.has(name) || this.skipNodes.has(element)) return;
    if (ctx.depth > MAX_DEPTH) {
      this.counts.nestedTooDeep += 1;
      this.text(textContent(element), ctx, false);
      return;
    }
    const css = this.css(element);
    if (/^none$/i.test(css.display ?? "") || /^hidden$/i.test(css.visibility ?? "") || /^all$/i.test(css["mso-hide"] ?? "") || element.attrs.hidden !== undefined) return;
    if (/^ignore$/i.test((css["mso-list"] ?? "").trim())) return;
    if (this.isReferencedBody(element)) return;
    if (hasClass(element, "footnotes-sep")) return;
    if (this.inNote && this.isBackLink(element)) return;
    const noteKey = this.noteRefs.get(element);
    if (noteKey !== undefined && !this.inNote) {
      this.footnoteRef(noteKey, ctx);
      return;
    }
    const next: Ctx = { ...ctx, depth: ctx.depth + 1 };
    if (this.trusted && element.attrs["data-simple-band"]) {
      this.band(element, next);
      return;
    }
    switch (name) {
      case "br":
        if (breaksBefore(css) || breaksAfter(css)) {
          this.close();
          this.pendingBreak = true;
        } else {
          const open = this.ensure(ctx);
          this.appendRun(open, "\v", ctx.char);
          open.spacePending = true;
        }
        return;
      case "img":
        this.image(element, next, css);
        return;
      case "table":
        this.table(element, next, css);
        return;
      case "hr":
        this.rule(element, next, css);
        return;
      case "ul": case "ol": case "menu": case "dir":
        this.list(element, next, css);
        return;
      case "li":
        this.listItem(element, next, css);
        return;
      case "input":
        if ((element.attrs.type ?? "").toLowerCase() === "checkbox") {
          const open = this.ensure(ctx);
          this.appendRun(open, element.attrs.checked !== undefined ? "☒ " : "☐ ", ctx.char);
          open.spacePending = true;
        }
        return;
      case "tr": case "thead": case "tbody": case "tfoot": case "td": case "th":
        // Table parts outside a table: keep their text as paragraphs.
        this.block(element, next, css);
        return;
      default:
        if (BLOCKS.has(name) || element === this.root || name === "#document") this.block(element, next, css);
        else this.inline(element, next, css);
    }
  }

  private band(element: DomElement, ctx: Ctx) {
    const kind = element.attrs["data-simple-band"] === "footer" ? "footer" : "header";
    const story = new Story();
    this.stories.push(story);
    const bandCtx = this.baseContext({ before: 0, after: 0 });
    bandCtx.depth = ctx.depth;
    this.walkChildren(element, bandCtx);
    this.close();
    this.stories.pop();
    if (story.blocks.length) this.bands[kind] = [...(this.bands[kind] ?? []), ...story.blocks];
  }

  // ---- styles --------------------------------------------------------------------

  private applyCharCss(char: CharStyle, css: Declarations, parentSize: number, inline: boolean) {
    const family = firstFontFamily(css["font-family"]);
    if (family) char.fontFamily = family;
    const size = parseFontSize(css["font-size"], parentSize);
    if (size) char.fontSizePx = Math.round(size * 1000) / 1000;
    const weight = (css["font-weight"] ?? "").toLowerCase();
    if (weight) {
      if (/^(?:bold|bolder|[6-9]00)$/.test(weight)) char.bold = true;
      else if (/^(?:normal|lighter|[1-5]00)$/.test(weight)) char.bold = false;
    }
    const style = (css["font-style"] ?? "").toLowerCase();
    if (/^(?:italic|oblique)/.test(style)) char.italic = true;
    else if (style === "normal") char.italic = false;
    const decoration = (css["text-decoration-line"] ?? css["text-decoration"] ?? "").toLowerCase();
    if (decoration) {
      if (/\bnone\b/.test(decoration)) {
        char.underline = false;
        char.strikethrough = false;
      }
      if (/\bunderline\b/.test(decoration)) char.underline = true;
      if (/\bline-through\b/.test(decoration)) char.strikethrough = true;
    }
    const color = parseColor(css.color);
    if (color) char.color = color;
    if (inline) {
      const highlight = backgroundColor(css) ?? parseColor(css["mso-highlight"]);
      if (highlight && highlight !== "#ffffff") char.highlightColor = highlight;
    }
    const vertical = (css["vertical-align"] ?? "").toLowerCase();
    if (vertical === "super" || vertical === "sub") char.verticalAlign = vertical;
    else if (vertical === "baseline") delete char.verticalAlign;
    const transform = (css["text-transform"] ?? "").toLowerCase();
    if (transform === "uppercase") char.caps = true;
    else if (transform === "none") delete char.caps;
    const variant = (css["font-variant-caps"] ?? css["font-variant"] ?? "").toLowerCase();
    if (variant.includes("small-caps")) char.smallCaps = true;
    else if (variant === "normal") delete char.smallCaps;
  }

  private applyParaCss(para: ParaStyle, css: Declarations, fontSize: number, element: DomElement) {
    const align = (css["text-align"] ?? element.attrs.align ?? "").toLowerCase();
    if (align === "left" || align === "start") para.align = "left";
    else if (align === "right" || align === "end") para.align = "right";
    else if (align === "center" || align === "middle") para.align = "center";
    else if (align === "justify") para.align = "justify";
    const box = boxValues(css.margin);
    const length = (value: string | undefined) => {
      const px = parseLength(value, fontSize);
      return px === null ? null : Math.round(px * 100) / 100;
    };
    const top = length(css["margin-top"] ?? box?.[0]);
    const bottom = length(css["margin-bottom"] ?? box?.[2]);
    const left = length(css["margin-left"] ?? box?.[3]);
    const right = length(css["margin-right"] ?? box?.[1]);
    if (top !== null) para.spaceBeforePx = Math.min(Math.max(top, 0), 600);
    if (bottom !== null) para.spaceAfterPx = Math.min(Math.max(bottom, 0), 600);
    if (left !== null && !(css["mso-list"])) para.indentLeftPx = Math.min(Math.max((para.indentLeftPx ?? 0) + left, 0), 1000);
    if (right !== null && right > 0) para.indentRightPx = Math.min(right, 1000);
    const indent = length(css["text-indent"]);
    if (indent !== null && !(css["mso-list"])) para.indentFirstLinePx = Math.max(Math.min(indent, 1000), -(para.indentLeftPx ?? 0));
    const lineHeight = (css["line-height"] ?? "").trim().toLowerCase();
    if (/^\d*\.?\d+$/.test(lineHeight)) para.lineHeight = Math.min(Math.max(Number(lineHeight), 0.5), 6);
    else if (/^\d*\.?\d+%$/.test(lineHeight)) para.lineHeight = Math.min(Math.max(Number.parseFloat(lineHeight) / 100, 0.5), 6);
    else if (lineHeight && lineHeight !== "normal") {
      const px = parseLength(lineHeight, fontSize);
      if (px !== null && px > 0) {
        para.lineRule = "atLeast";
        para.lineHeightPx = Math.min(px, 1000);
      }
    }
    const direction = (css.direction ?? element.attrs.dir ?? "").toLowerCase();
    if (direction === "rtl") para.direction = "rtl";
    else if (direction === "ltr") delete para.direction;
    const shading = backgroundColor(css) ?? parseColor(element.attrs.bgcolor);
    if (shading && shading !== "#ffffff") para.shading = shading;
    if (/^avoid$/i.test(css["page-break-inside"] ?? css["break-inside"] ?? "")) para.keepLinesTogether = true;
    if (/^avoid$/i.test(css["page-break-after"] ?? css["break-after"] ?? "")) para.keepWithNext = true;
  }

  private blockContext(element: DomElement, ctx: Ctx, css: Declarations): Ctx {
    const name = element.name;
    const para: ParaStyle = { ...ctx.para };
    delete para.list;
    delete para.pageBreakBefore;
    delete para.shading;
    delete para.namedStyle;
    delete para.outlineLevel;
    delete para.keepWithNext;
    delete para.borders;
    para.spaceBeforePx = ctx.spacing.before;
    para.spaceAfterPx = ctx.spacing.after;
    let char: CharStyle = { ...ctx.char };
    const next: Ctx = { ...ctx, para, char };
    const heading = /^h([1-6])$/.exec(name);
    if (heading) {
      const level = Number(heading[1]);
      next.char = char = headingCharStyle(level, char);
      // Headings follow the caller's font and colour (the size stays the heading's).
      const base = this.baseChar();
      if (base.fontFamily) char.fontFamily = base.fontFamily;
      if (base.color) char.color = base.color;
      if (ctx.char.link) char.link = ctx.char.link;
      next.para = headingParaStyle(level, para);
      if (ctx.spacing.before === 0 && ctx.spacing.after === 0 && ctx.para.indentLeftPx) next.para.indentLeftPx = ctx.para.indentLeftPx;
    } else if (name === "blockquote") {
      para.indentLeftPx = (para.indentLeftPx ?? 0) + 48;
    } else if (name === "pre" || name === "listing" || name === "xmp" || name === "plaintext") {
      char.fontFamily = MONOSPACE_FONT;
      char.fontSizePx = Math.round(pt(10) * 1000) / 1000;
      para.lineHeight = 1;
      para.spaceAfterPx = 0;
      next.pre = true;
    } else if (name === "dt") {
      char.bold = true;
      para.spaceAfterPx = 0;
    } else if (name === "dd") {
      para.indentLeftPx = (para.indentLeftPx ?? 0) + 48;
    } else if (name === "center") {
      para.align = "center";
    } else if (name === "address") {
      char.italic = true;
    } else if (name === "caption" || name === "figcaption") {
      para.align = "center";
      char.italic = true;
    }
    this.applyCharCss(next.char, css, ctx.char.fontSizePx, false);
    this.applyParaCss(next.para, css, next.char.fontSizePx, element);
    if (/^(?:pre|pre-wrap|break-spaces)$/.test((css["white-space"] ?? "").toLowerCase())) next.pre = true;
    return next;
  }

  // ---- blocks --------------------------------------------------------------------

  private block(element: DomElement, ctx: Ctx, css: Declarations) {
    this.close();
    if (breaksBefore(css)) this.pendingBreak = true;
    const next = this.blockContext(element, ctx, css);
    this.listMembership(element, css, ctx, next);
    // HTML drops one newline right after <pre>.
    const leading = element.children[0];
    if (next.pre && leading?.type === "text") leading.text = leading.text.replace(/^\r?\n/, "");
    const story = this.story;
    const before = story.produced;
    this.walkChildren(element, next);
    if (PARAGRAPH_BLOCKS.has(element.name) && story.produced === before && !story.open) this.ensure(next);
    this.close();
    if (next.item && next.item.used && this.markerGroup && this.markerGroup.listId === next.item.listId) this.markerGroup.blockCount = story.blocks.length;
    if (breaksAfter(css)) this.pendingBreak = true;
  }

  /** List membership a paragraph declares itself (front-end hints, Word HTML, Simple's own HTML). */
  private listMembership(element: DomElement, css: Declarations, ctx: Ctx, next: Ctx) {
    if (this.trusted && element.attrs["data-simple-list"] !== undefined) {
      const key = element.attrs["data-simple-list"];
      const level = clampInt(element.attrs["data-simple-level"], 0, 8) ?? 0;
      let listId = this.explicitLists.get(key);
      if (!listId) {
        listId = this.newList(false);
        this.explicitLists.set(key, listId);
        const levels = this.options.lists?.[key];
        if (levels?.length) {
          this.lists[listId].levels = this.lists[listId].levels.map((fallback, index) => {
            const level = levels[index] ?? levels[levels.length - 1];
            return level ? listLevel(index, level.format, { start: level.start, bulletChar: level.bulletChar, text: level.text }) : fallback;
          });
        }
      }
      next.item = { listId, level, used: false };
      next.para.indentLeftPx = ctx.para.indentLeftPx;
      next.para.indentFirstLinePx = 0;
      return;
    }
    if (!ctx.listMarkers || ctx.item) return;
    // Word "Web Page": style="mso-list:l0 level2 lfo1" plus a marker span with mso-list:Ignore.
    const mso = /\bl(\d+)\s+level(\d+)/i.exec(css["mso-list"] ?? "");
    let markerText: string | null = null;
    let level = 0;
    let key: string | null = null;
    if (mso) {
      const marker = findElements(element, (candidate) => /^ignore$/i.test((this.css(candidate)["mso-list"] ?? "").trim()), 1)[0];
      markerText = marker ? textContent(marker).replace(/[\s\u00a0]+/g, " ").trim() : "";
      level = Math.min(8, Math.max(0, Number(mso[2]) - 1));
      key = `mso:${mso[1]}`;
    } else {
      // Simple's HTML export: <p><span class="list-marker">1.</span>…</p>
      const first = element.children.find((child) => child.type === "element" || child.text.trim());
      if (first && first.type === "element" && first.name === "span" && hasClass(first, "list-marker")) {
        markerText = textContent(first).trim();
        this.skipNodes.add(first);
        const margin = /^([\d.]+)em$/.exec((css["margin-left"] ?? "").trim());
        level = margin ? Math.min(8, Math.round(Number(margin[1]) / 1.5)) : 0;
      }
    }
    if (markerText === null) return;
    const marker = markerFormat(markerText);
    const ordered = marker.format !== "bullet";
    let listId: string;
    if (key) {
      listId = this.markerLists.get(key) ?? this.newList(ordered);
      this.markerLists.set(key, listId);
    } else {
      const group = this.markerGroup;
      const consecutive = group && group.story === this.story && group.blockCount === this.story.blocks.length;
      if (consecutive && (level > 0 || group.ordered === ordered)) listId = group.listId;
      else if (group && ordered && level === 0 && group.ordered && group.story === this.story && marker.start === group.next) listId = group.listId;
      else listId = this.newList(ordered);
      this.markerGroup = { listId, story: this.story, blockCount: -1, ordered: level === 0 ? ordered : group?.ordered ?? ordered, next: level === 0 && ordered ? marker.start + 1 : group?.next ?? 1 };
    }
    this.defineLevel(listId, level, marker.format, { start: marker.start, bulletChar: marker.bulletChar });
    next.item = { listId, level, used: false };
    next.para.indentLeftPx = ctx.para.indentLeftPx;
    next.para.indentFirstLinePx = 0;
  }

  private newList(ordered: boolean): string {
    const id = this.ids("list");
    this.lists[id] = listDefinition(id, ordered);
    this.touchedLevels.set(id, new Set());
    return id;
  }

  private defineLevel(listId: string, level: number, format: ListNumberFormat, options: { start?: number; bulletChar?: string; text?: string }) {
    const touched = this.touchedLevels.get(listId)!;
    if (touched.has(level)) return;
    touched.add(level);
    this.lists[listId].levels[level] = listLevel(level, format, options);
  }

  private list(element: DomElement, ctx: Ctx, css: Declarations) {
    this.close();
    if (breaksBefore(css)) this.pendingBreak = true;
    const ordered = element.name === "ol";
    const style = htmlListFormat(element, css, ordered);
    const level = Math.min(8, ctx.listLevel + 1);
    const start = clampInt(element.attrs.start, 0, 1_000_000) ?? 1;
    let listId = ctx.listId;
    if (!listId || ctx.listLevel < 0) {
      const previous = this.lastOrderedList;
      if (ordered && start > 1 && previous && previous.story === this.story && previous.next === start) listId = previous.listId;
      else listId = this.newList(ordered);
    }
    if (style) this.defineLevel(listId, level, style.format, { start, bulletChar: style.bulletChar });
    else if (!ordered) this.defineLevel(listId, level, "bullet", {});
    else this.defineLevel(listId, level, (["decimal", "lowerLetter", "lowerRoman"] as const)[level % 3], { start });
    const next: Ctx = { ...ctx, listId, listLevel: level, item: null };
    if (!style && /\bnone\b/.test((css["list-style-type"] ?? css["list-style"] ?? "").toLowerCase())) next.listMarkers = false;
    next.para = { ...ctx.para };
    this.applyCharCss(next.char = { ...ctx.char }, css, ctx.char.fontSizePx, false);
    this.walkChildren(element, next);
    this.close();
    if (level === 0 && ordered) this.lastOrderedList = { listId, next: this.lists[listId].levels[0].start + (this.topLevelCounts.get(listId) ?? 0), story: this.story };
    if (breaksAfter(css)) this.pendingBreak = true;
  }

  private listItem(element: DomElement, ctx: Ctx, css: Declarations) {
    this.close();
    let base = ctx;
    if (!ctx.listId) {
      const listId = this.newList(false);
      this.defineLevel(listId, 0, "bullet", {});
      base = { ...ctx, listId, listLevel: 0 };
    }
    const next = this.blockContext(element, base, css);
    next.para.indentLeftPx = base.para.indentLeftPx;
    next.para.indentFirstLinePx = 0;
    next.item = base.listMarkers ? { listId: base.listId!, level: Math.max(0, base.listLevel), used: false } : null;
    if (!base.listMarkers) next.para.indentLeftPx = (next.para.indentLeftPx ?? 0) + 48 * (Math.max(0, base.listLevel) + 1);
    const story = this.story;
    const before = story.produced;
    this.walkChildren(element, next);
    if (story.produced === before && !story.open) this.ensure(next);
    this.close();
  }

  private rule(element: DomElement, ctx: Ctx, css: Declarations) {
    this.close();
    if (breaksBefore(css) || breaksAfter(css) || hasClass(element, "page-break") || hasClass(element, "pagebreak")) {
      this.pendingBreak = true;
      return;
    }
    const next = this.blockContext(element, ctx, {});
    next.para.borders = { bottom: { color: parseColor(css["border-color"] ?? css.color) ?? "#bfbfbf", widthPx: 1, style: "single" } };
    next.para.spaceAfterPx = Math.max(next.para.spaceAfterPx, 6);
    this.ensure(next);
    this.close();
  }

  // ---- inline ------------------------------------------------------------------

  private inline(element: DomElement, ctx: Ctx, css: Declarations) {
    const char: CharStyle = { ...ctx.char };
    const next: Ctx = { ...ctx, char };
    switch (element.name) {
      case "b": case "strong": char.bold = true; break;
      case "i": case "em": case "cite": case "dfn": case "var": char.italic = true; break;
      case "u": case "ins": char.underline = true; break;
      case "s": case "strike": case "del": char.strikethrough = true; break;
      case "sup": char.verticalAlign = "super"; break;
      case "sub": char.verticalAlign = "sub"; break;
      case "code": case "kbd": case "samp": case "tt": char.fontFamily = MONOSPACE_FONT; break;
      case "mark": char.highlightColor = "#ffff00"; break;
      case "small": char.fontSizePx = Math.round(ctx.char.fontSizePx * 0.833 * 1000) / 1000; break;
      case "big": char.fontSizePx = Math.round(ctx.char.fontSizePx * 1.2 * 1000) / 1000; break;
      case "font": {
        const color = parseColor(element.attrs.color);
        if (color) char.color = color;
        const face = firstFontFamily(element.attrs.face);
        if (face) char.fontFamily = face;
        const size = element.attrs.size?.trim();
        if (size) {
          const sizes = [10, 13, 16, 18, 24, 32, 48];
          const relative = /^[+-]/.test(size);
          const index = (relative ? 3 + Number.parseInt(size, 10) : Number.parseInt(size, 10)) - 1;
          if (Number.isFinite(index)) char.fontSizePx = sizes[Math.min(6, Math.max(0, index))];
        }
        break;
      }
      case "a": {
        const href = safeHref(element.attrs.href);
        if (href) {
          char.link = href;
          char.color = LINK_COLOR;
          char.underline = true;
        }
        break;
      }
      case "q": {
        this.applyCharCss(char, css, ctx.char.fontSizePx, true);
        this.text("“", next, true);
        this.walkChildren(element, next);
        this.text("”", next, true);
        return;
      }
      default: break;
    }
    this.applyCharCss(char, css, ctx.char.fontSizePx, true);
    this.walkChildren(element, next);
  }

  // ---- pictures ----------------------------------------------------------------

  private resolveImage(source: string): ImageBytes | null {
    const src = source.trim();
    if (!src) return null;
    if (/^data:/i.test(src)) {
      const decoded = decodeImageDataUrl(src);
      if (!decoded) this.counts.badImages += 1;
      return decoded;
    }
    const resolved = this.options.resolveImage?.(src);
    if (resolved && resolved.bytes?.length && resolved.bytes.length <= MAX_IMAGE_BYTES) {
      const type = sniffImageType(resolved.bytes);
      if (type) return { bytes: resolved.bytes, type };
      this.counts.badImages += 1;
      return null;
    }
    if (/^(?:https?|ftp):/i.test(src) || src.startsWith("//")) this.counts.remoteImages += 1;
    else this.counts.localImages += 1;
    return null;
  }

  private image(element: DomElement, ctx: Ctx, css: Declarations) {
    const picture = this.resolveImage(element.attrs.src ?? element.attrs["data-src"] ?? "");
    if (!picture) {
      if (this.options.imagePlaceholders === false) return;
      const alt = (element.attrs.alt ?? element.attrs.title ?? "").replace(/\s+/g, " ").trim();
      const open = this.ensure(ctx);
      if (!open.spacePending && open.runs.length) this.appendRun(open, " ", ctx.char);
      this.appendRun(open, alt ? `[Picture: ${alt}]` : "[Picture]", { ...ctx.char, italic: true, color: PLACEHOLDER_COLOR, link: ctx.char.link });
      open.spacePending = false;
      return;
    }
    const natural = imageNaturalSize(picture.bytes, picture.type);
    const dimension = (cssValue: string | undefined, attr: string | undefined) => {
      const value = cssValue && cssValue !== "auto" ? cssValue : attr;
      if (!value) return null;
      if (/%\s*$/.test(value)) return parseLength(value, this.maxImageWidth);
      const px = parseLength(/^\s*[\d.]+\s*$/.test(value) ? `${value}px` : value, ctx.char.fontSizePx);
      return px !== null && px > 0 ? px : null;
    };
    let width = dimension(css.width, element.attrs.width);
    let height = dimension(css.height, element.attrs.height);
    const ratio = natural ? natural.height / natural.width : 0.75;
    if (width && !height) height = width * ratio;
    else if (height && !width) width = height / ratio;
    else if (!width || !height) {
      width = natural?.width ?? 320;
      height = natural?.height ?? 240;
    }
    if (width! > this.maxImageWidth) {
      height = height! * this.maxImageWidth / width!;
      width = this.maxImageWidth;
    }
    const float = (css.float ?? element.attrs.align ?? "").toLowerCase();
    const align: ImageBlockModel["align"] = float === "right" ? "right" : float === "left" ? "left" : ctx.para.align === "center" || ctx.para.align === "right" ? ctx.para.align : "left";
    const block: ImageBlockModel = {
      kind: "image",
      id: this.ids("img"),
      revision: 0,
      src: bytesToDataUrl(picture.bytes, picture.type),
      widthPx: Math.max(1, Math.round(width!)),
      heightPx: Math.max(1, Math.round(height!)),
      align,
    };
    if (float === "left" || float === "right") block.wrap = "square";
    this.pushBlock(block, ctx);
  }

  // ---- tables ------------------------------------------------------------------

  private table(element: DomElement, ctx: Ctx, css: Declarations) {
    this.close();
    if (breaksBefore(css)) this.pendingBreak = true;
    const rowElements: DomElement[] = [];
    const columns: DomElement[] = [];
    let caption: DomElement | null = null;
    const loose: DomElement[] = [];
    for (const child of element.children) {
      if (child.type !== "element") continue;
      if (child.name === "tr") rowElements.push(child);
      else if (child.name === "thead" || child.name === "tbody" || child.name === "tfoot") {
        for (const row of child.children) if (row.type === "element" && row.name === "tr") rowElements.push(row);
      } else if (child.name === "caption") caption = child;
      else if (child.name === "colgroup") columns.push(...child.children.filter((col): col is DomElement => col.type === "element" && col.name === "col"));
      else if (child.name === "col") columns.push(child);
      else if (child.name === "td" || child.name === "th") loose.push(child);
    }
    if (loose.length && !rowElements.length) {
      const row: DomElement = { type: "element", name: "tr", attrs: {}, children: loose, parent: element };
      rowElements.push(row);
    }
    if (caption) this.block(caption, ctx, this.css(caption));
    const rows: TableBlock["rows"] = [];
    const firstRowWidths: Array<number | null> = [];
    let firstRowUnit: "px" | "%" | null = null;
    rowElements.forEach((rowElement, rowIndex) => {
      const rowCss = this.css(rowElement);
      if (/^none$/i.test(rowCss.display ?? "")) return;
      const cells: TableCell[] = [];
      for (const cellElement of rowElement.children) {
        if (cellElement.type !== "element" || (cellElement.name !== "td" && cellElement.name !== "th")) continue;
        const cellCss = this.css(cellElement);
        const cellCtx = this.baseContext({ before: 0, after: 0 });
        cellCtx.char = { ...ctx.char };
        cellCtx.depth = ctx.depth + 1;
        cellCtx.para.align = ctx.para.align === "justify" ? "left" : ctx.para.align;
        cellCtx.para.direction = ctx.para.direction;
        if (cellCtx.para.direction === undefined) delete cellCtx.para.direction;
        if (cellElement.name === "th") cellCtx.char.bold = true;
        this.applyCharCss(cellCtx.char, rowCss, ctx.char.fontSizePx, false);
        this.applyCharCss(cellCtx.char, cellCss, ctx.char.fontSizePx, false);
        const align = (cellCss["text-align"] ?? cellElement.attrs.align ?? rowCss["text-align"] ?? rowElement.attrs.align ?? "").toLowerCase();
        if (align === "center" || align === "right" || align === "justify") cellCtx.para.align = align;
        else if (align === "left" || align === "start") cellCtx.para.align = "left";
        else if (align === "end") cellCtx.para.align = "right";
        const story = new Story();
        this.stories.push(story);
        this.walkChildren(cellElement, cellCtx);
        this.close();
        this.stories.pop();
        const cell: TableCell = { id: this.ids("c"), blocks: story.blocks.length ? story.blocks : [this.emptyParagraph(cellCtx)] };
        const colSpan = clampInt(cellElement.attrs.colspan, 1, 1000) ?? 1;
        const rowSpan = clampInt(cellElement.attrs.rowspan, 1, 10_000) ?? 1;
        if (colSpan > 1) cell.colSpan = colSpan;
        if (rowSpan > 1) cell.rowSpan = rowSpan;
        const shading = backgroundColor(cellCss) ?? parseColor(cellElement.attrs.bgcolor) ?? backgroundColor(rowCss) ?? parseColor(rowElement.attrs.bgcolor);
        if (shading) cell.shading = shading;
        if (rowIndex === 0) {
          const raw = (cellCss.width ?? cellElement.attrs.width ?? "").trim();
          const unit = raw.endsWith("%") ? "%" : raw ? "px" : null;
          const value = raw ? parseLength(/^[\d.]+$/.test(raw) ? `${raw}px` : raw, 100) : null;
          if (unit && value && (firstRowUnit === null || firstRowUnit === unit)) {
            firstRowUnit = unit;
            for (let i = 0; i < colSpan; i += 1) firstRowWidths.push(value / colSpan);
          } else for (let i = 0; i < colSpan; i += 1) firstRowWidths.push(null);
        }
        cells.push(cell);
      }
      if (cells.length) rows.push({ cells });
    });
    if (!rows.length) {
      if (breaksAfter(css)) this.pendingBreak = true;
      return;
    }
    const table: TableBlock = { kind: "table", id: this.ids("t"), revision: 0, rows };
    const gridWidth = tableGridWidth(rows);
    const columnWidths = columns.flatMap((col) => {
      const span = clampInt(col.attrs.span, 1, 1000) ?? 1;
      const raw = (this.css(col).width ?? col.attrs.width ?? "").trim();
      const value = raw ? parseLength(/^[\d.]+$/.test(raw) ? `${raw}px` : raw, 100) : null;
      return Array.from({ length: span }, () => value ? value / span : null);
    });
    const widths = columnWidths.length === gridWidth && columnWidths.every((value) => value) ? columnWidths : firstRowWidths.length === gridWidth && firstRowWidths.every((value) => value) ? firstRowWidths : null;
    if (widths && gridWidth > 1) {
      const total = widths.reduce((sum, value) => sum! + value!, 0)!;
      if (total > 0) table.colFractions = widths.map((value) => Math.round((value! / total) * 10000) / 10000);
    }
    this.pushBlock(table, ctx);
    if (breaksAfter(css)) this.pendingBreak = true;
  }

  // ---- result --------------------------------------------------------------------

  finish(): MapResult {
    this.close();
    this.convertNotes();
    const body = this.stories[0];
    const baseCtx = this.baseContext();
    const section: Partial<Document["section"]> = { ...this.options.section };
    if (this.bands.header?.length) section.header = this.bands.header;
    if (this.bands.footer?.length) section.footer = this.bands.footer;
    const document = assembleDocument({ blocks: body.blocks, lists: this.usedLists(body.blocks), footnotes: this.footnotes, section, emptyParagraph: () => this.emptyParagraph(baseCtx) });
    const { remoteImages, localImages, badImages, nestedTooDeep, droppedNoteObjects } = this.counts;
    const plural = (count: number, one: string, many: string) => count === 1 ? one : many.replace("#", String(count));
    const shown = this.options.imagePlaceholders !== false;
    const placeholder = (count: number) => shown ? `${plural(count, "is", "are")} shown as a placeholder` : `${plural(count, "was", "were")} left out`;
    if (remoteImages) this.warnings.push(`${plural(remoteImages, "A picture stored on the web was", "# pictures stored on the web were")} not downloaded and ${placeholder(remoteImages)}.`);
    if (localImages) this.warnings.push(`${plural(localImages, "A picture stored outside the file was", "# pictures stored outside the file were")} not found and ${placeholder(localImages)}.`);
    if (badImages) this.warnings.push(`${plural(badImages, "A picture is", "# pictures are")} in a format Simple cannot show and ${shown ? `${plural(badImages, "was", "were")} replaced with a placeholder` : placeholder(badImages)}.`);
    if (nestedTooDeep) this.warnings.push("Very deeply nested content was kept as plain text.");
    if (droppedNoteObjects) this.warnings.push("Tables or pictures inside footnotes were left out.");
    return { document, warnings: this.warnings };
  }

  private usedLists(blocks: Block[]): Record<string, ListDefinition> {
    const used = new Set<string>();
    const visit = (items: Block[]) => {
      for (const block of items) {
        if (block.kind === "paragraph" && block.style.list) used.add(block.style.list.listId);
        if (block.kind === "table") for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    };
    visit(blocks);
    for (const band of Object.values(this.bands)) if (band) visit(band);
    const result: Record<string, ListDefinition> = {};
    for (const id of used) if (this.lists[id]) result[id] = this.lists[id];
    return result;
  }
}

/** Number of grid columns of a table, counting row spans from rows above. */
export function tableGridWidth(rows: TableBlock["rows"]): number {
  const carried: number[] = [];
  let width = 0;
  for (const row of rows) {
    let column = 0;
    const next: number[] = [];
    const occupied = (index: number) => (carried[index] ?? 0) > 0;
    for (const cell of row.cells) {
      while (occupied(column)) column += 1;
      const span = cell.colSpan ?? 1;
      for (let i = 0; i < span; i += 1) next[column + i] = (cell.rowSpan ?? 1) - 1;
      column += span;
    }
    const rowWidth = Math.max(column, carried.length);
    width = Math.max(width, rowWidth);
    for (let i = 0; i < rowWidth; i += 1) carried[i] = Math.max((carried[i] ?? 0) - 1, next[i] ?? 0, 0);
  }
  return width;
}

/** Maps a parsed tree to a WordCanvas document. */
export function domToDocument(root: DomElement, options: MapOptions = {}): MapResult {
  const mapper = new Mapper(root, options);
  mapper.walkChildren(root, mapper.baseContext());
  return mapper.finish();
}
