/**
 * Rich Text Format (.rtf) import for the RTF that WordPad, Word and most editors write:
 * fonts and code pages, colours, bold/italic/underline/strike, sizes, highlight,
 * super/subscript, paragraphs and alignment, headings (from the style sheet), Word and
 * WordPad lists, tables with merged cells, hyperlinks, footnotes, the default header and
 * footer, page breaks and embedded PNG/JPEG pictures. A small tokenizer turns the file
 * into a group tree; the reader turns that into the shared node tree for the mapper.
 */
import { domToDocument, markerFormat } from "./dom-to-model.ts";
import type { ExplicitListLevel } from "./dom-to-model.ts";
import { bytesToDataUrl, hexToBytes, imageNaturalSize, sniffImageType } from "./images.ts";
import { appendChild, createElement } from "./markup.ts";
import type { DomElement, DomNode } from "./markup.ts";
import type { Document, ImportOptions, ImportResult, ListNumberFormat } from "./model.ts";
import { toBytes } from "./text-decoding.ts";

export type RtfItem = RtfGroup | RtfWord | RtfText | RtfHex | RtfBin;
export interface RtfGroup { t: "g"; items: RtfItem[] }
export interface RtfWord { t: "w"; word: string; param: number | null }
export interface RtfText { t: "x"; text: string }
export interface RtfHex { t: "h"; byte: number }
export interface RtfBin { t: "b"; data: string }

const MAX_DEPTH = 1000;

/** Tokenizes RTF (one char per byte) into a group tree. */
export function parseRtfGroups(source: string): RtfGroup {
  const root: RtfGroup = { t: "g", items: [] };
  const stack: RtfGroup[] = [root];
  let depthOverflow = 0;
  let text = "";
  const current = () => stack[stack.length - 1];
  const flushText = () => {
    if (text) current().items.push({ t: "x", text });
    text = "";
  };
  const length = source.length;
  let i = 0;
  while (i < length) {
    const ch = source[i];
    if (ch === "{") {
      flushText();
      if (stack.length > MAX_DEPTH) depthOverflow += 1;
      else {
        const group: RtfGroup = { t: "g", items: [] };
        current().items.push(group);
        stack.push(group);
      }
      i += 1;
    } else if (ch === "}") {
      flushText();
      if (depthOverflow) depthOverflow -= 1;
      else if (stack.length > 1) stack.pop();
      i += 1;
    } else if (ch === "\\") {
      const next = source[i + 1];
      if (next === undefined) break;
      if (/[A-Za-z]/.test(next)) {
        flushText();
        let end = i + 1;
        while (end < length && end - i <= 32 && /[A-Za-z]/.test(source[end])) end += 1;
        const word = source.slice(i + 1, end);
        let param: number | null = null;
        const number = /^-?\d{1,10}/.exec(source.slice(end, end + 11));
        if (number) {
          param = Number(number[0]);
          end += number[0].length;
        }
        if (source[end] === " ") end += 1;
        if (word === "bin" && param !== null && param > 0) {
          current().items.push({ t: "b", data: source.slice(end, end + param) });
          end += param;
        } else current().items.push({ t: "w", word, param });
        i = end;
      } else if (next === "'") {
        flushText();
        const byte = parseInt(source.slice(i + 2, i + 4), 16);
        if (Number.isFinite(byte)) current().items.push({ t: "h", byte });
        i += 4;
      } else if (next === "{" || next === "}" || next === "\\") {
        text += next;
        i += 2;
      } else if (next === "\r" || next === "\n") {
        flushText();
        current().items.push({ t: "w", word: "par", param: null });
        i += next === "\r" && source[i + 2] === "\n" ? 3 : 2;
      } else {
        flushText();
        current().items.push({ t: "w", word: next, param: null });
        i += 2;
      }
    } else if (ch === "\r" || ch === "\n") {
      i += 1;
    } else {
      text += ch;
      i += 1;
    }
  }
  flushText();
  return root;
}

const CHARSET_CODEPAGES: Record<number, number> = {
  0: 1252, 77: 10000, 128: 932, 129: 949, 130: 1361, 134: 936, 136: 950, 161: 1253, 162: 1254, 163: 1258, 177: 1255, 178: 1256,
  186: 1257, 204: 1251, 222: 874, 238: 1250, 255: 437,
};
const CODEPAGE_LABELS: Record<number, string> = {
  437: "ibm866", 850: "ibm866", 866: "ibm866", 874: "windows-874", 932: "shift_jis", 936: "gbk", 949: "euc-kr", 950: "big5",
  1250: "windows-1250", 1251: "windows-1251", 1252: "windows-1252", 1253: "windows-1253", 1254: "windows-1254", 1255: "windows-1255",
  1256: "windows-1256", 1257: "windows-1257", 1258: "windows-1258", 10000: "macintosh", 65001: "utf-8",
};
const SYMBOL_LETTERS = "αβχδεφγηιϕκλμνοπθρστυϖωξψζ";
const SYMBOL_CAPITALS = "ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ";
const SYMBOL_BULLETS: Record<number, string> = { 0xb7: "•", 0xa7: "▪", 0x6f: "◦", 0xd8: "➢", 0xfc: "✓", 0x76: "❖", 0x71: "❑", 0xa8: "•", 0x6e: "■", 0x6c: "●", 0x2d: "–" };
const SPECIAL_WORDS: Record<string, string> = {
  tab: "\t", emdash: "—", endash: "–", emspace: " ", enspace: " ", qmspace: " ", bullet: "•", lquote: "‘", rquote: "’",
  ldblquote: "“", rdblquote: "”", zwj: "\u200d", zwnj: "\u200c", ltrmark: "\u200e", rtlmark: "\u200f", line: "\v", softline: "\v",
  "~": "\u00a0", _: "‑",
};
// Destinations whose content is never document text.
const SKIPPED_DESTINATIONS = new Set([
  "filetbl", "revtbl", "rsidtbl", "generator", "xmlnstbl", "themedata", "colorschememapping", "latentstyles", "datastore", "defchp",
  "defpap", "pgdsctbl", "nonshppict", "headerl", "headerf", "footerl", "footerf", "bkmkstart", "bkmkend", "objdata", "xe", "tc", "tcn",
  "atnid", "atnauthor", "annotation", "atnref", "atrfstart", "atrfend", "atndate", "atnparent", "protusertbl", "docvar", "userprops",
  "mmathPr", "pgptbl", "nesttableprops", "template", "fchars", "lchars", "ftnsep", "ftnsepc", "ftncn", "aftnsep", "aftnsepc", "aftncn",
  "wgrffmtfilter", "passwordhash", "background", "fldinst", "pntext", "listtext", "blipuid", "picprop", "shpinst", "sp", "sn", "sv",
  "do", "dptxbxtext", "private", "keywords", "doccomm", "comment", "author", "operator", "company", "subject", "category", "manager",
  "hlinkbase", "creatim", "revtim", "printim", "buptim", "listpicture", "falt", "panose", "fname", "oldcprops", "oldpprops",
  "oldsprops", "oldtprops", "ud" /* handled through upr */, "formfield", "datafield", "ffdeftext", "ffname", "ffstattext", "ffhelptext",
  "ffentrymcr", "ffexitmcr", "ffformat", "fftext", "ebcstart", "ebcend", "factoidname", "smarttagtype", "mhtmltag", "htmltag",
]);
const LEVEL_FORMATS: Record<number, ListNumberFormat> = { 0: "decimal", 1: "upperRoman", 2: "lowerRoman", 3: "upperLetter", 4: "lowerLetter", 5: "decimal", 22: "decimal", 23: "bullet" };

interface CharProps {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  font: number;
  size: number;
  color: number;
  highlight: number;
  background: number;
  vert: "super" | "sub" | null;
  caps: boolean;
  smallCaps: boolean;
  hidden: boolean;
}

interface PnProps {
  format: ListNumberFormat;
  start: number;
  level: number;
  bullet?: string;
}

interface ParaProps {
  align: "left" | "center" | "right" | "justify";
  li: number;
  ri: number;
  fi: number;
  sb: number;
  sa: number;
  sl: number;
  slmult: boolean;
  keepNext: boolean;
  style: number;
  outline: number | null;
  inTable: boolean;
  ls: number | null;
  ilvl: number;
  pn: PnProps | null;
  rtl: boolean;
  shading: number;
}

interface State {
  char: CharProps;
  para: ParaProps;
  uc: number;
  link: string | null;
}

interface CellDef {
  right: number;
  hMerge: "first" | "cont" | null;
  vMerge: "first" | "cont" | null;
  shading: number;
}

interface RowDef {
  left: number;
  cells: CellDef[];
}

interface CellContent {
  blocks: DomElement[];
}

interface Sink {
  blocks: DomElement[];
  inline: DomNode[];
  inlineKey: string | null;
  rows: Array<{ cells: CellContent[]; def: RowDef }>;
  cells: CellContent[];
  cellBlocks: DomElement[];
}

const newSink = (): Sink => ({ blocks: [], inline: [], inlineKey: null, rows: [], cells: [], cellBlocks: [] });

function defaultPara(): ParaProps {
  return { align: "left", li: 0, ri: 0, fi: 0, sb: 0, sa: 0, sl: 0, slmult: false, keepNext: false, style: 0, outline: null, inTable: false, ls: null, ilvl: 0, pn: null, rtl: false, shading: 0 };
}

const twipsToPx = (twips: number) => Math.round(twips / 15 * 100) / 100;

class RtfReader {
  readonly fonts = new Map<number, { name: string; codepage: number; symbol: boolean }>();
  colors: Array<string | null> = [];
  readonly styles = new Map<number, { name: string; outline: number | null }>();
  readonly listLevels = new Map<number, ExplicitListLevel[]>();
  readonly overrides = new Map<number, number>();
  readonly explicitLists: Record<string, ExplicitListLevel[]> = {};
  codepage = 1252;
  defaultFont = 0;
  sink: Sink = newSink();
  readonly body: Sink = this.sink;
  readonly extra: DomElement[] = [];
  title: string | undefined;
  section: Partial<Document["section"]> | undefined;
  private page: { width?: number; height?: number; left?: number; right?: number; top?: number; bottom?: number } = {};
  private bytes: number[] = [];
  private bytesState: State | null = null;
  private skip = 0;
  private pendingBreak = false;
  private noteCount = 0;
  private rowDef: RowDef = { left: 0, cells: [] };
  private cellDraft: CellDef = { right: 0, hMerge: null, vMerge: null, shading: 0 };
  private pnSerial = 0;
  private listText = "";
  private lastPnKey: string | null = null;
  private lastWasPn = false;
  private hasHeader = false;
  private hasFooter = false;
  counts = { unsupportedPictures: 0, nestedTables: 0 };

  initialState(): State {
    return { char: this.defaultChar(), para: defaultPara(), uc: 1, link: null };
  }

  private defaultChar(): CharProps {
    return { bold: false, italic: false, underline: false, strike: false, font: this.defaultFont, size: 24, color: 0, highlight: 0, background: 0, vert: null, caps: false, smallCaps: false, hidden: false };
  }

  private clone(state: State): State {
    return { char: { ...state.char }, para: { ...state.para, pn: state.para.pn ? { ...state.para.pn } : null }, uc: state.uc, link: state.link };
  }

  // ---- text ------------------------------------------------------------------------

  private codepageFor(state: State): { codepage: number; symbol: boolean } {
    const font = this.fonts.get(state.char.font);
    return { codepage: font?.codepage ?? this.codepage, symbol: font?.symbol ?? false };
  }

  private decode(bytes: number[], state: State): string {
    const { codepage, symbol } = this.codepageFor(state);
    if (symbol) return bytes.map((byte) => symbolCharacter(byte)).join("");
    const label = CODEPAGE_LABELS[codepage] ?? "windows-1252";
    try {
      return new TextDecoder(label).decode(Uint8Array.from(bytes));
    } catch {
      return new TextDecoder("windows-1252").decode(Uint8Array.from(bytes));
    }
  }

  private pushBytes(values: number[], state: State) {
    if (this.bytesState && this.bytesState !== state) this.flush();
    this.bytesState = state;
    for (const value of values) this.bytes.push(value);
  }

  /** Text token chars are bytes; chars above U+00FF only occur in already-decoded input. */
  private textBytes(text: string, state: State) {
    let bytes: number[] = [];
    for (const ch of text) {
      const code = ch.codePointAt(0)!;
      if (code <= 0xff) {
        bytes.push(code);
        continue;
      }
      if (bytes.length) this.pushBytes(bytes, state);
      bytes = [];
      this.flush();
      this.emitText(ch, state);
    }
    if (bytes.length) this.pushBytes(bytes, state);
  }

  private flush() {
    if (!this.bytes.length || !this.bytesState) return;
    const state = this.bytesState;
    const text = this.decode(this.bytes, state);
    this.bytes = [];
    this.bytesState = null;
    this.emitText(text, state);
  }

  private emitText(text: string, state: State) {
    if (!text || state.char.hidden) return;
    const sink = this.sink;
    const css = this.charCss(state);
    const key = JSON.stringify(css) + (state.link ?? "");
    const last = sink.inline[sink.inline.length - 1];
    if (last && last.type === "element" && sink.inlineKey === key) {
      const textNode = last.children[0];
      if (textNode && textNode.type === "text") {
        textNode.text += text;
        return;
      }
    }
    const node = createElement(state.link ? "a" : "span", state.link ? { href: state.link } : {}, [{ type: "text", text, pre: true }], css);
    sink.inline.push(node);
    sink.inlineKey = key;
  }

  private pushInline(node: DomNode) {
    this.flush();
    this.sink.inline.push(node);
    this.sink.inlineKey = null;
  }

  private charCss(state: State): Record<string, string> {
    const char = state.char;
    const css: Record<string, string> = {};
    const font = this.fonts.get(char.font);
    if (font?.name && !font.symbol) css["font-family"] = font.name;
    css["font-size"] = `${Math.max(1, char.size) / 2}pt`;
    css["font-weight"] = char.bold ? "bold" : "normal";
    css["font-style"] = char.italic ? "italic" : "normal";
    css["text-decoration"] = [char.underline ? "underline" : "", char.strike ? "line-through" : ""].filter(Boolean).join(" ") || "none";
    const color = this.colors[char.color];
    if (char.color > 0 && color) css.color = color;
    const highlight = (char.highlight > 0 && this.colors[char.highlight]) || (char.background > 0 && this.colors[char.background]) || null;
    if (highlight) css["background-color"] = highlight;
    if (char.vert) css["vertical-align"] = char.vert;
    if (char.caps) css["text-transform"] = "uppercase";
    if (char.smallCaps) css["font-variant"] = "small-caps";
    return css;
  }

  // ---- paragraphs and tables -----------------------------------------------------------

  private headingLevel(para: ParaProps): number | null {
    const style = this.styles.get(para.style);
    const match = style ? /^heading\s*([1-9])$/i.exec(style.name.trim()) : null;
    if (match) return Math.min(6, Number(match[1]));
    const outline = para.outline ?? style?.outline ?? null;
    return outline !== null && outline >= 0 && outline < 6 ? outline + 1 : null;
  }

  private listKey(para: ParaProps): { key: string; level: number } | null {
    if (para.ls !== null && para.ls > 0) {
      const key = `ls${para.ls}`;
      if (!this.explicitLists[key]) {
        const listId = this.overrides.get(para.ls);
        const known = listId !== undefined ? this.listLevels.get(listId) : undefined;
        if (known) this.explicitLists[key] = known;
        else {
          // No list table: the marker text Word wrote in listtext tells the format.
          const marker = markerFormat(this.listText);
          this.explicitLists[key] = [{ format: marker.format, start: marker.start, bulletChar: marker.bulletChar }];
        }
      }
      return { key, level: Math.min(8, Math.max(0, para.ilvl)) };
    }
    if (para.pn) {
      const signature = `${para.pn.format}:${para.pn.bullet ?? ""}`;
      if (!this.lastWasPn || this.lastPnKey === null || !this.lastPnKey.endsWith(`|${signature}`)) {
        this.pnSerial += 1;
        this.lastPnKey = `pn${this.pnSerial}|${signature}`;
        const levels: ExplicitListLevel[] = [];
        for (let level = 0; level <= para.pn.level; level += 1) levels.push({ format: para.pn.format, start: para.pn.start, bulletChar: para.pn.bullet });
        this.explicitLists[this.lastPnKey] = levels;
      }
      return { key: this.lastPnKey, level: para.pn.level };
    }
    return null;
  }

  private paragraphElement(state: State): DomElement {
    const para = state.para;
    const heading = this.headingLevel(para);
    const css: Record<string, string> = {
      "text-align": para.align,
      "margin-top": `${twipsToPx(Math.max(0, para.sb))}px`,
      "margin-bottom": `${twipsToPx(Math.max(0, para.sa))}px`,
    };
    const list = this.listKey(para);
    this.lastWasPn = para.pn !== null;
    this.listText = "";
    if (!list) {
      if (para.li) css["margin-left"] = `${twipsToPx(para.li)}px`;
      if (para.fi) css["text-indent"] = `${twipsToPx(para.fi)}px`;
    }
    if (para.ri > 0) css["margin-right"] = `${twipsToPx(para.ri)}px`;
    if (para.sl > 0 && para.slmult) css["line-height"] = String(Math.round(para.sl / 240 * 100) / 100);
    else if (para.sl > 0) css["line-height"] = `${twipsToPx(para.sl)}px`;
    else if (para.sl < 0) css["line-height"] = `${twipsToPx(-para.sl)}px`;
    else css["line-height"] = "1";
    if (para.rtl) css.direction = "rtl";
    if (para.keepNext) css["page-break-after"] = "avoid";
    const shading = para.shading > 0 ? this.colors[para.shading] : null;
    if (shading) css["background-color"] = shading;
    if (this.pendingBreak && this.sink === this.body) {
      css["break-before"] = "page";
      this.pendingBreak = false;
    }
    const attrs: Record<string, string> = {};
    if (list) {
      attrs["data-simple-list"] = list.key;
      attrs["data-simple-level"] = String(list.level);
    }
    const element = createElement(heading ? `h${heading}` : "p", attrs, [], css);
    for (const node of this.sink.inline) appendChild(element, node);
    this.sink.inline = [];
    this.sink.inlineKey = null;
    return element;
  }

  private endParagraph(state: State, toCell = false) {
    this.flush();
    const element = this.paragraphElement(state);
    if (state.para.inTable || toCell) this.sink.cellBlocks.push(element);
    else {
      this.flushTable();
      this.sink.blocks.push(element);
    }
  }

  private endCell(state: State) {
    this.flush();
    if (this.sink.inline.length || !this.sink.cellBlocks.length) this.endParagraph(state, true);
    this.sink.cells.push({ blocks: this.sink.cellBlocks });
    this.sink.cellBlocks = [];
  }

  private endRow() {
    this.flush();
    const sink = this.sink;
    if (sink.inline.length || sink.cellBlocks.length) {
      // Text after the last \cell belongs to a cell of its own.
      sink.cells.push({ blocks: sink.cellBlocks });
      sink.cellBlocks = [];
    }
    if (sink.cells.length) sink.rows.push({ cells: sink.cells, def: { left: this.rowDef.left, cells: this.rowDef.cells.map((cell) => ({ ...cell })) } });
    sink.cells = [];
  }

  flushTable() {
    const sink = this.sink;
    if (sink.cells.length || sink.cellBlocks.length) this.endRow();
    if (!sink.rows.length) return;
    const rows = sink.rows;
    sink.rows = [];
    sink.blocks.push(buildTable(rows, this.colors));
  }

  finishSink(state: State) {
    this.flush();
    if (this.sink.inline.length) this.endParagraph({ ...state, para: { ...state.para, inTable: false } });
    this.flushTable();
  }

  // ---- groups ------------------------------------------------------------------------

  walkGroup(group: RtfGroup, parent: State) {
    this.flush();
    this.skip = 0;
    const items = group.items;
    let index = 0;
    let starred = false;
    if (items[0]?.t === "w" && items[0].word === "*") {
      starred = true;
      index = 1;
    }
    const first = items[index];
    if (first?.t === "w") {
      if (this.destination(first.word, group, parent, index)) {
        this.skip = 0;
        return;
      }
      if (starred || SKIPPED_DESTINATIONS.has(first.word)) return;
    }
    const state = this.clone(parent);
    for (; index < items.length; index += 1) {
      const item = items[index];
      if (item.t === "g") {
        this.walkGroup(item, state);
        continue;
      }
      if (this.skip > 0) {
        if (item.t === "x") {
          const take = Math.min(this.skip, item.text.length);
          this.skip -= take;
          if (take < item.text.length) this.textBytes(item.text.slice(take), state);
          continue;
        }
        if (item.t === "h" || item.t === "b" || (item.t === "w" && item.word !== "u")) {
          this.skip -= 1;
          continue;
        }
      }
      if (item.t === "x") this.textBytes(item.text, state);
      else if (item.t === "h") this.pushBytes([item.byte], state);
      else if (item.t === "w") this.word(item.word, item.param, state);
    }
    this.flush();
    this.skip = 0;
  }

  /** Handles destination groups; returns true when the group was consumed. */
  private destination(word: string, group: RtfGroup, parent: State, index: number): boolean {
    switch (word) {
      case "fonttbl": this.fontTable(group); return true;
      case "colortbl": this.colorTable(group); return true;
      case "stylesheet": this.styleSheet(group); return true;
      case "listtable": this.listTable(group); return true;
      case "listoverridetable": this.overrideTable(group); return true;
      case "info": this.info(group); return true;
      case "pict": this.picture(group, parent); return true;
      case "pn": this.pn(group, parent); return true;
      case "shppict": this.walkItems(group.items.slice(index + 1), parent); return true;
      case "listtext": case "pntext": this.listText = this.groupText({ t: "g", items: group.items.slice(index + 1) }).replace(/[\s\u00a0]+/g, "").trim(); return true;
      case "field": this.field(group, parent); return true;
      case "footnote": this.footnote(group, parent); return true;
      case "header": case "headerr": case "footer": case "footerr": this.band(group, parent, word.startsWith("header") ? "header" : "footer"); return true;
      case "upr": {
        const unicode = group.items.find((item): item is RtfGroup => item.t === "g" && item.items.some((inner) => inner.t === "w" && inner.word === "ud"));
        if (unicode) this.walkItems(unicode.items.slice(unicode.items.findIndex((inner) => inner.t === "w" && inner.word === "ud") + 1), parent);
        else this.walkItems(group.items.slice(index + 1), parent);
        return true;
      }
      case "shp": {
        const picture = findPicture(group, true);
        if (picture) this.picture(picture, parent);
        else {
          const result = group.items.find((item): item is RtfGroup => item.t === "g" && firstWord(item) === "shprslt");
          if (result) this.walkItems(result.items.slice(1), parent);
        }
        return true;
      }
      case "object": {
        const result = group.items.find((item): item is RtfGroup => item.t === "g" && firstWord(item) === "result");
        if (result) this.walkItems(result.items.slice(1), parent);
        return true;
      }
      default: return false;
    }
  }

  private walkItems(items: RtfItem[], parent: State) {
    this.walkGroup({ t: "g", items: [{ t: "w", word: "rtlch", param: null }, ...items] }, parent);
  }

  private word(word: string, param: number | null, state: State) {
    const on = param === null || param !== 0;
    const char = state.char;
    const para = state.para;
    if (word !== "u") this.flush();
    switch (word) {
      // Character formatting.
      case "plain": state.char = this.defaultChar(); return;
      case "b": char.bold = on; return;
      case "i": char.italic = on; return;
      case "ul": case "uld": case "uldash": case "uldashd": case "uldashdd": case "uldb": case "ulth": case "ulw": case "ulwave": case "ulhwave":
      case "ulldash": case "ulthd": case "ulthdash": case "ulthdashd": case "ulthdashdd": case "ulthldash": case "ululdbwave":
        char.underline = on; return;
      case "ulnone": char.underline = false; return;
      case "strike": case "striked": char.strike = on; return;
      case "f": char.font = param ?? 0; return;
      case "fs": if (param && param > 0) char.size = param; return;
      case "cf": char.color = param ?? 0; return;
      case "highlight": char.highlight = param ?? 0; return;
      case "chcbpat": case "cb": char.background = param ?? 0; return;
      case "super": char.vert = "super"; return;
      case "sub": char.vert = "sub"; return;
      case "nosupersub": char.vert = null; return;
      case "up": char.vert = on && param !== 0 ? "super" : null; return;
      case "dn": char.vert = on && param !== 0 ? "sub" : null; return;
      case "caps": char.caps = on; return;
      case "scaps": char.smallCaps = on; return;
      case "v": char.hidden = on; return;
      case "uc": state.uc = Math.max(0, param ?? 1); return;
      case "u": {
        this.flush();
        const code = param === null ? 0x3f : param < 0 ? param + 65536 : param;
        this.emitText(String.fromCharCode(code), state);
        this.skip = state.uc;
        return;
      }
      // Paragraph formatting.
      case "pard": {
        state.para = defaultPara();
        return;
      }
      case "ql": para.align = "left"; return;
      case "qc": para.align = "center"; return;
      case "qr": para.align = "right"; return;
      case "qj": case "qd": para.align = "justify"; return;
      case "li": case "lin": para.li = param ?? 0; return;
      case "ri": case "rin": para.ri = param ?? 0; return;
      case "fi": para.fi = param ?? 0; return;
      case "sb": para.sb = param ?? 0; return;
      case "sa": para.sa = param ?? 0; return;
      case "sl": para.sl = param ?? 0; return;
      case "slmult": para.slmult = on; return;
      case "keepn": para.keepNext = on; return;
      case "pagebb": if (on) this.pendingBreak = true; return;
      case "s": para.style = param ?? 0; return;
      case "outlinelevel": para.outline = param; return;
      case "intbl": para.inTable = on; return;
      case "itap": para.inTable = (param ?? 1) > 0; if ((param ?? 0) > 1) this.counts.nestedTables += 1; return;
      case "ls": para.ls = param; return;
      case "ilvl": para.ilvl = param ?? 0; return;
      case "rtlpar": para.rtl = true; return;
      case "ltrpar": para.rtl = false; return;
      case "cbpat": para.shading = param ?? 0; return;
      // Breaks.
      case "par": this.endParagraph(state); return;
      case "sect": this.endParagraph(state); this.pendingBreak = true; return;
      case "page":
        this.flush();
        if (this.sink.inline.length) this.endParagraph(state);
        this.pendingBreak = true;
        return;
      case "column": this.emitText("\v", state); return;
      case "cell": this.endCell(state); return;
      case "nestcell": this.endParagraph(state, true); return;
      case "row": this.endRow(); return;
      case "nestrow": return;
      case "lastrow": return;
      // Table definitions.
      case "trowd": this.rowDef = { left: 0, cells: [] }; this.cellDraft = { right: 0, hMerge: null, vMerge: null, shading: 0 }; return;
      case "trleft": this.rowDef.left = param ?? 0; return;
      case "clmgf": this.cellDraft.hMerge = "first"; return;
      case "clmrg": this.cellDraft.hMerge = "cont"; return;
      case "clvmgf": this.cellDraft.vMerge = "first"; return;
      case "clvmrg": this.cellDraft.vMerge = "cont"; return;
      case "clcbpat": this.cellDraft.shading = param ?? 0; return;
      case "cellx":
        this.rowDef.cells.push({ ...this.cellDraft, right: param ?? 0 });
        this.cellDraft = { right: 0, hMerge: null, vMerge: null, shading: 0 };
        return;
      // Document.
      case "ansicpg": if (param && CODEPAGE_LABELS[param]) this.codepage = param; return;
      case "mac": this.codepage = 10000; return;
      case "pc": case "pca": this.codepage = 866; return;
      case "deff": this.defaultFont = param ?? 0; state.char.font = this.defaultFont; return;
      case "paperw": this.page.width = param ?? undefined; return;
      case "paperh": this.page.height = param ?? undefined; return;
      case "margl": this.page.left = param ?? undefined; return;
      case "margr": this.page.right = param ?? undefined; return;
      case "margt": this.page.top = param ?? undefined; return;
      case "margb": this.page.bottom = param ?? undefined; return;
      case "chftn": case "chpgn": case "chdate": case "chtime": case "chatn": return;
      default: {
        const special = SPECIAL_WORDS[word];
        if (special) this.emitText(special, state);
      }
    }
  }

  // ---- destinations -------------------------------------------------------------------

  private groupText(group: RtfGroup, codepage = this.codepage): string {
    let bytes: number[] = [];
    let text = "";
    let skip = 0;
    const flush = () => {
      if (!bytes.length) return;
      text += new TextDecoder(CODEPAGE_LABELS[codepage] ?? "windows-1252").decode(Uint8Array.from(bytes));
      bytes = [];
    };
    const visit = (items: RtfItem[]) => {
      for (const item of items) {
        if (item.t === "g") {
          if (item.items[0]?.t === "w" && item.items[0].word === "*") continue;
          visit(item.items);
        } else if (item.t === "x") {
          const chars = item.text.slice(skip);
          skip = Math.max(0, skip - item.text.length);
          for (const ch of chars) bytes.push(ch.charCodeAt(0) & 0xff);
        } else if (item.t === "h") {
          if (skip > 0) skip -= 1;
          else bytes.push(item.byte);
        } else if (item.t === "w" && item.word === "u" && item.param !== null) {
          flush();
          text += String.fromCharCode(item.param < 0 ? item.param + 65536 : item.param);
          skip = 1;
        } else if (item.t === "w" && SPECIAL_WORDS[item.word]) {
          flush();
          text += SPECIAL_WORDS[item.word];
        }
      }
    };
    visit(group.items);
    flush();
    return text;
  }

  private fontTable(group: RtfGroup) {
    const entries = group.items.some((item) => item.t === "g") ? group.items.filter((item): item is RtfGroup => item.t === "g") : [group];
    for (const entry of entries) {
      let index: number | null = null;
      let charset: number | null = null;
      let cpg: number | null = null;
      for (const item of entry.items) {
        if (item.t !== "w") continue;
        if (item.word === "f" && index === null) index = item.param;
        else if (item.word === "fcharset") charset = item.param;
        else if (item.word === "cpg") cpg = item.param;
      }
      if (index === null) continue;
      const codepage = cpg ?? (charset !== null ? CHARSET_CODEPAGES[charset] : undefined) ?? this.codepage;
      const plain: RtfGroup = { t: "g", items: entry.items.filter((item) => item.t !== "g") };
      const name = this.groupText(plain, codepage).replace(/;[\s\S]*$/, "").trim();
      const symbol = charset === 2 || /^(?:symbol|wingdings\s*\d?|webdings)$/i.test(name);
      this.fonts.set(index, { name, codepage, symbol });
    }
  }

  private colorTable(group: RtfGroup) {
    const colors: Array<string | null> = [];
    let red = -1;
    let green = -1;
    let blue = -1;
    const hex = (value: number) => Math.max(0, Math.min(255, value)).toString(16).padStart(2, "0");
    for (const item of group.items) {
      if (item.t === "w") {
        if (item.word === "red") red = item.param ?? 0;
        else if (item.word === "green") green = item.param ?? 0;
        else if (item.word === "blue") blue = item.param ?? 0;
      } else if (item.t === "x") {
        for (const ch of item.text) {
          if (ch !== ";") continue;
          colors.push(red < 0 && green < 0 && blue < 0 ? null : `#${hex(Math.max(red, 0))}${hex(Math.max(green, 0))}${hex(Math.max(blue, 0))}`);
          red = green = blue = -1;
        }
      }
    }
    this.colors = colors;
  }

  private styleSheet(group: RtfGroup) {
    for (const entry of group.items) {
      if (entry.t !== "g") continue;
      let index = 0;
      let outline: number | null = null;
      let paragraphStyle = true;
      for (const item of entry.items) {
        if (item.t !== "w") continue;
        if (item.word === "s") index = item.param ?? 0;
        else if (item.word === "cs" || item.word === "ds" || item.word === "ts") paragraphStyle = false;
        else if (item.word === "outlinelevel") outline = item.param;
      }
      if (!paragraphStyle) continue;
      const plain: RtfGroup = { t: "g", items: entry.items.filter((item) => item.t !== "g") };
      const name = this.groupText(plain).replace(/;[\s\S]*$/, "").trim();
      this.styles.set(index, { name, outline });
    }
  }

  private listTable(group: RtfGroup) {
    for (const list of group.items) {
      if (list.t !== "g" || firstWord(list) !== "list") continue;
      let id: number | null = null;
      const levels: ExplicitListLevel[] = [];
      for (const item of list.items) {
        if (item.t === "w" && item.word === "listid") id = item.param;
        if (item.t !== "g" || firstWord(item) !== "listlevel") continue;
        let format: ListNumberFormat = "decimal";
        let start = 1;
        let text: string | undefined;
        let bullet: string | undefined;
        for (const part of item.items) {
          if (part.t === "w" && (part.word === "levelnfc" || part.word === "levelnfcn")) format = LEVEL_FORMATS[part.param ?? 0] ?? (part.param === 255 ? "bullet" : "decimal");
          else if (part.t === "w" && part.word === "levelstartat") start = part.param ?? 1;
          else if (part.t === "g" && firstWord(part) === "leveltext") {
            const codes = levelTextCodes(part);
            const count = codes[0] ?? 0;
            const body = codes.slice(1, 1 + count);
            text = body.map((code) => (code < 9 ? `%${code + 1}` : String.fromCharCode(code))).join("");
            const symbolic = body.find((code) => code >= 9);
            if (symbolic !== undefined) bullet = bulletFromCode(symbolic);
          }
        }
        levels.push(format === "bullet" ? { format, start, bulletChar: bullet ?? "•" } : { format, start, text: text && /%\d/.test(text) ? text : undefined });
      }
      if (id !== null && levels.length) this.listLevels.set(id, levels);
    }
  }

  private overrideTable(group: RtfGroup) {
    for (const override of group.items) {
      if (override.t !== "g" || firstWord(override) !== "listoverride") continue;
      let id: number | null = null;
      let ls: number | null = null;
      for (const item of override.items) {
        if (item.t === "w" && item.word === "listid") id = item.param;
        if (item.t === "w" && item.word === "ls") ls = item.param;
      }
      if (id !== null && ls !== null) this.overrides.set(ls, id);
    }
  }

  private info(group: RtfGroup) {
    const title = group.items.find((item): item is RtfGroup => item.t === "g" && firstWord(item) === "title");
    if (title) {
      const text = this.groupText({ t: "g", items: title.items.slice(1) }).trim();
      if (text) this.title = text;
    }
  }

  private pn(group: RtfGroup, parent: State) {
    let format: ListNumberFormat | null = null;
    let level = 0;
    let start = 1;
    let bullet: string | undefined;
    let symbolFont = false;
    for (const item of group.items) {
      if (item.t === "w") {
        switch (item.word) {
          case "pnlvlblt": format = "bullet"; break;
          case "pnlvlbody": format = format ?? "decimal"; break;
          case "pnlvl": level = Math.min(8, Math.max(0, (item.param ?? 1) - 1)); break;
          case "pndec": format = "decimal"; break;
          case "pnucltr": format = "upperLetter"; break;
          case "pnlcltr": format = "lowerLetter"; break;
          case "pnucrm": format = "upperRoman"; break;
          case "pnlcrm": format = "lowerRoman"; break;
          case "pnstart": start = item.param ?? 1; break;
          case "pnf": symbolFont = this.fonts.get(item.param ?? -1)?.symbol ?? false; break;
          default: break;
        }
      } else if (item.t === "g" && firstWord(item) === "pntxtb") {
        const codes = levelTextCodes({ t: "g", items: item.items.slice(1) }, false);
        if (codes.length) bullet = symbolFont ? bulletFromCode(codes[0]) : bulletFromCode(codes[0]);
      }
    }
    parent.para.pn = { format: format ?? "bullet", start, level, bullet: (format ?? "bullet") === "bullet" ? bullet ?? "•" : undefined };
  }

  private field(group: RtfGroup, parent: State) {
    const instruction = group.items.find((item): item is RtfGroup => item.t === "g" && fieldPart(item) === "fldinst");
    const result = group.items.find((item): item is RtfGroup => item.t === "g" && fieldPart(item) === "fldrslt");
    const instructionText = instruction ? this.groupText({ t: "g", items: instruction.items }) : "";
    const hyperlink = /^\s*HYPERLINK\s+(?:\\[a-z]\s+(?:"[^"]*"\s+)?)*"?([^"\s]+)"?/i.exec(instructionText);
    if (!result) return;
    const state = this.clone(parent);
    if (hyperlink && /^(?:https?|mailto|tel|ftp):/i.test(hyperlink[1])) state.link = hyperlink[1];
    // A picture inserted through INCLUDEPICTURE lives in the result; everything else too.
    this.walkGroup(result, state);
  }

  private footnote(group: RtfGroup, parent: State) {
    this.flush();
    this.noteCount += 1;
    const key = `rtf-note-${this.noteCount}`;
    this.pushInline(createElement("sup", { "data-simple-footnote": key }, [{ type: "text", text: String(this.noteCount) }], this.charCss(parent)));
    const saved = this.sink;
    const savedBreak = this.pendingBreak;
    this.sink = newSink();
    const state = this.clone(parent);
    state.para = defaultPara();
    this.walkGroup({ t: "g", items: group.items.slice(1) }, state);
    this.finishSink(state);
    const body = createElement("aside", { "data-simple-footnote-body": key }, this.sink.blocks);
    this.sink = saved;
    this.pendingBreak = savedBreak;
    this.extra.push(body);
  }

  private band(group: RtfGroup, parent: State, kind: "header" | "footer") {
    if ((kind === "header" && this.hasHeader) || (kind === "footer" && this.hasFooter)) return;
    this.flush();
    const saved = this.sink;
    const savedBreak = this.pendingBreak;
    this.sink = newSink();
    const state = this.clone(parent);
    state.para = defaultPara();
    this.walkGroup({ t: "g", items: group.items.slice(1) }, state);
    this.finishSink(state);
    const blocks = this.sink.blocks;
    this.sink = saved;
    this.pendingBreak = savedBreak;
    if (!blocks.some((block) => blockHasText(block))) return;
    if (kind === "header") this.hasHeader = true;
    else this.hasFooter = true;
    this.extra.push(createElement("div", { "data-simple-band": kind }, blocks));
  }

  private picture(group: RtfGroup, parent: State) {
    let type: string | null = null;
    let widthGoal = 0;
    let heightGoal = 0;
    let scaleX = 100;
    let scaleY = 100;
    let hex = "";
    let binary: string | null = null;
    for (const item of group.items) {
      if (item.t === "w") {
        switch (item.word) {
          case "pngblip": type = "image/png"; break;
          case "jpegblip": type = "image/jpeg"; break;
          case "emfblip": case "wmetafile": case "macpict": case "pmmetafile": case "dibitmap": case "wbitmap": type = type ?? "unsupported"; break;
          case "picwgoal": widthGoal = item.param ?? 0; break;
          case "pichgoal": heightGoal = item.param ?? 0; break;
          case "picscalex": scaleX = item.param ?? 100; break;
          case "picscaley": scaleY = item.param ?? 100; break;
          default: break;
        }
      } else if (item.t === "x") hex += item.text;
      else if (item.t === "b") binary = item.data;
    }
    const bytes = binary !== null ? Uint8Array.from(binary, (ch) => ch.charCodeAt(0) & 0xff) : hexToBytes(hex);
    const sniffed = bytes.length ? sniffImageType(bytes) : null;
    if (!sniffed || (type === "unsupported" && !sniffed)) {
      this.counts.unsupportedPictures += 1;
      this.emitText("[Picture]", { ...parent, char: { ...parent.char, italic: true, hidden: false } });
      return;
    }
    const natural = imageNaturalSize(bytes, sniffed);
    let width = widthGoal > 0 ? twipsToPx(widthGoal) * scaleX / 100 : natural?.width ?? 320;
    let height = heightGoal > 0 ? twipsToPx(heightGoal) * scaleY / 100 : natural ? natural.height * (width / natural.width) : 240;
    if (!(width > 0) || !(height > 0)) {
      width = natural?.width ?? 320;
      height = natural?.height ?? 240;
    }
    this.pushInline(createElement("img", { src: bytesToDataUrl(bytes, sniffed), width: String(Math.round(width)), height: String(Math.round(height)) }));
  }

  finish(): Partial<Document["section"]> | undefined {
    const page = this.page;
    if (!page.width && !page.height && page.left === undefined && page.top === undefined) return undefined;
    const width = page.width ?? 12240;
    const height = page.height ?? 15840;
    if (width < 1440 || height < 1440 || width > 200_000 || height > 200_000) return undefined;
    return {
      pageWidthPx: twipsToPx(width),
      pageHeightPx: twipsToPx(height),
      marginPx: { top: twipsToPx(page.top ?? 1440), right: twipsToPx(page.right ?? 1800), bottom: twipsToPx(page.bottom ?? 1440), left: twipsToPx(page.left ?? 1800) },
    };
  }
}

function firstWord(group: RtfGroup): string | null {
  for (const item of group.items) {
    if (item.t === "w") return item.word === "*" ? firstWordAfterStar(group) : item.word;
    if (item.t === "x" && item.text.trim()) return null;
  }
  return null;
}

function firstWordAfterStar(group: RtfGroup): string | null {
  const index = group.items.findIndex((item) => item.t === "w" && item.word === "*");
  const next = group.items[index + 1];
  return next?.t === "w" ? next.word : "*";
}

function fieldPart(group: RtfGroup): string | null {
  return firstWord(group);
}

/** A PNG/JPEG picture inside a Word shape (shpinst > sp > sv > pict). */
function findPicture(group: RtfGroup, preferInstructions: boolean): RtfGroup | null {
  const pending: RtfGroup[] = [group];
  while (pending.length) {
    const current = pending.shift()!;
    for (const item of current.items) {
      if (item.t !== "g") continue;
      const word = firstWord(item);
      if (word === "pict" && item.items.some((inner) => inner.t === "w" && (inner.word === "pngblip" || inner.word === "jpegblip"))) return item;
      if (!preferInstructions || word !== "shprslt") pending.push(item);
    }
  }
  return null;
}

function levelTextCodes(group: RtfGroup, skipLeadingWord = true): number[] {
  const codes: number[] = [];
  let skip = 0;
  const items = skipLeadingWord ? group.items.slice(1) : group.items;
  for (const item of items) {
    if (item.t === "h") {
      if (skip > 0) skip -= 1;
      else codes.push(item.byte);
    } else if (item.t === "x") {
      for (const ch of item.text) {
        if (skip > 0) { skip -= 1; continue; }
        if (ch === ";") return codes;
        codes.push(ch.charCodeAt(0));
      }
    } else if (item.t === "w" && item.word === "u" && item.param !== null) {
      codes.push(item.param < 0 ? item.param + 65536 : item.param);
      skip = 1;
    }
  }
  return codes;
}

function symbolCharacter(byte: number): string {
  if (SYMBOL_BULLETS[byte]) return SYMBOL_BULLETS[byte];
  if (byte >= 0x61 && byte <= 0x7a) return SYMBOL_LETTERS[byte - 0x61];
  if (byte >= 0x41 && byte <= 0x5a) return SYMBOL_CAPITALS[byte - 0x41];
  return String.fromCharCode(byte);
}

function bulletFromCode(code: number): string {
  const byte = code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code;
  if (SYMBOL_BULLETS[byte]) return SYMBOL_BULLETS[byte];
  if (byte < 0x20 || (code >= 0xe000 && code <= 0xf8ff)) return "•";
  return String.fromCharCode(byte);
}

function blockHasText(element: DomElement): boolean {
  const stack: DomNode[] = [element];
  while (stack.length) {
    const node = stack.pop()!;
    if (node.type === "text") {
      if (node.text.trim()) return true;
    } else {
      if (node.name === "img") return true;
      stack.push(...node.children);
    }
  }
  return false;
}

/** Builds an HTML-like table from RTF rows, resolving \cellx boundaries and merges. */
function buildTable(rows: Array<{ cells: CellContent[]; def: RowDef }>, colors: Array<string | null>): DomElement {
  // Grid boundaries from every row's cell edges.
  const edges = new Set<number>();
  const rowEdges = rows.map(({ cells, def }) => {
    const rights = def.cells.map((cell) => cell.right);
    // Rows without (enough) \cellx get equal columns.
    if (rights.length < cells.length) {
      const last = rights[rights.length - 1] ?? 0;
      const step = 1440 * 1.5;
      for (let i = rights.length; i < cells.length; i += 1) rights.push(last + step * (i - rights.length + 1));
    }
    const left = def.left;
    edges.add(left);
    for (const right of rights) edges.add(right);
    return { left, rights };
  });
  const grid = [...edges].sort((a, b) => a - b);
  const columnOf = (position: number) => {
    let best = 0;
    for (let i = 0; i < grid.length; i += 1) if (Math.abs(grid[i] - position) < Math.abs(grid[best] - position)) best = i;
    return best;
  };
  interface Placed { element: DomElement; column: number; span: number; vMerge: "first" | "cont" | null }
  const placedRows: Placed[][] = [];
  rows.forEach(({ cells, def }, rowIndex) => {
    const { left, rights } = rowEdges[rowIndex];
    const placed: Placed[] = [];
    let start = columnOf(left);
    cells.forEach((cell, cellIndex) => {
      const cellDef = def.cells[cellIndex] ?? { right: rights[cellIndex], hMerge: null, vMerge: null, shading: 0 };
      const end = Math.max(start + 1, columnOf(rights[cellIndex]));
      if (cellDef.hMerge === "cont" && placed.length) {
        const previous = placed[placed.length - 1];
        previous.span += end - start;
        start = end;
        return;
      }
      const css: Record<string, string> = {};
      const shading = cellDef.shading > 0 ? colors[cellDef.shading] : null;
      if (shading) css["background-color"] = shading;
      const element = createElement("td", {}, cell.blocks, css);
      placed.push({ element, column: start, span: end - start, vMerge: cellDef.vMerge });
      start = end;
    });
    placedRows.push(placed);
  });
  // Vertical merges: a "cont" cell extends the cell above that starts in the same column.
  const table = createElement("table");
  const colgroup = createElement("colgroup");
  for (let i = 1; i < grid.length; i += 1) appendChild(colgroup, createElement("col", { width: String(Math.max(1, twipsToPx(grid[i] - grid[i - 1]))) }));
  appendChild(table, colgroup);
  const owners = new Map<number, Placed>();
  for (const placed of placedRows) {
    const row = createElement("tr");
    for (const cell of placed) {
      if (cell.vMerge === "cont" && owners.has(cell.column)) {
        const owner = owners.get(cell.column)!;
        owner.element.attrs.rowspan = String(Number(owner.element.attrs.rowspan ?? "1") + 1);
        continue;
      }
      if (cell.span > 1) cell.element.attrs.colspan = String(cell.span);
      if (cell.vMerge === "first") owners.set(cell.column, cell);
      else owners.delete(cell.column);
      appendChild(row, cell.element);
    }
    if (row.children.length) appendChild(table, row);
  }
  return table;
}

/** Converts RTF text (one char per byte, or already-decoded text) into a node tree. */
export function rtfToTree(source: string): { root: DomElement; lists: Record<string, ExplicitListLevel[]>; title?: string; section?: Partial<Document["section"]>; warnings: string[] } {
  const reader = new RtfReader();
  const groups = parseRtfGroups(source);
  const top = groups.items.find((item): item is RtfGroup => item.t === "g") ?? groups;
  const state = reader.initialState();
  // The document group's leading words (\rtf1\ansi\deff0 ...) are ordinary words.
  reader.walkGroup({ t: "g", items: top.items.filter((item, index) => !(index === 0 && item.t === "w" && item.word === "rtf")) }, state);
  reader.finishSink(state);
  const root = createElement("#document");
  for (const block of reader.body.blocks) appendChild(root, block);
  for (const extra of reader.extra) appendChild(root, extra);
  const warnings: string[] = [];
  if (reader.counts.unsupportedPictures) warnings.push(reader.counts.unsupportedPictures === 1 ? "A picture in WMF, EMF or another older format is shown as a placeholder." : `${reader.counts.unsupportedPictures} pictures in WMF, EMF or other older formats are shown as placeholders.`);
  if (reader.counts.nestedTables) warnings.push("Tables inside table cells were flattened into the outer table.");
  return { root, lists: reader.explicitLists, title: reader.title, section: reader.finish(), warnings };
}

/** One char per byte (TextDecoder("latin1") is windows-1252 and would remap 0x80-0x9F). */
function binaryString(bytes: Uint8Array): string {
  let text = "";
  for (let i = 0; i < bytes.length; i += 8192) text += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return text;
}

export function looksLikeRtf(bytes: Uint8Array): boolean {
  let offset = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (offset < bytes.length && offset < 64 && (bytes[offset] === 0x20 || bytes[offset] === 0x0d || bytes[offset] === 0x0a || bytes[offset] === 0x09)) offset += 1;
  return bytes[offset] === 0x7b && bytes[offset + 1] === 0x5c && bytes[offset + 2] === 0x72 && bytes[offset + 3] === 0x74 && bytes[offset + 4] === 0x66;
}

export function importRtf(input: Uint8Array | ArrayBuffer | string, options: ImportOptions = {}): ImportResult {
  let source: string;
  if (typeof input === "string") source = input;
  else {
    const bytes = toBytes(input);
    // One char per byte: \'hh, \bin and 8-bit text are decoded by the reader per font code page.
    source = binaryString(bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? bytes.subarray(3) : bytes);
  }
  if (!/^\s*\{\\rtf/.test(source)) throw new Error("This file is not a Rich Text (RTF) document.");
  const tree = rtfToTree(source);
  const mapped = domToDocument(tree.root, { ...options, trusted: true, authorStyles: false, lists: tree.lists, section: tree.section });
  const result: ImportResult = { document: mapped.document, format: "rtf", warnings: [...tree.warnings, ...mapped.warnings] };
  if (tree.title) result.title = tree.title;
  return result;
}
