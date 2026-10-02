/**
 * OpenDocument Text (.odt, .ott and flat .fodt) import without LibreOffice: JSZip reads
 * the package, content.xml/styles.xml are parsed with the inert XML parser, styles are
 * resolved along their parent chains into CSS, and the shared mapper builds the model.
 * Covers headings, paragraphs and spans, lists with their list styles, tables with
 * spans, pictures from Pictures/, links, footnotes/endnotes, page size and the default
 * header and footer.
 */
import JSZip from "jszip";
import { domToDocument } from "./dom-to-model.ts";
import type { ExplicitListLevel } from "./dom-to-model.ts";
import { parseLength } from "./css.ts";
import { bytesToDataUrl, MAX_IMAGE_BYTES, sniffImageType } from "./images.ts";
import { appendChild, childElements, createElement, findElements, parseMarkup, textContent } from "./markup.ts";
import type { DomElement, DomNode } from "./markup.ts";
import type { Document, ImportOptions, ImportResult, ListNumberFormat } from "./model.ts";
import { decodeTextBytes, toBytes } from "./text-decoding.ts";

const ODT_TYPES = new Set(["application/vnd.oasis.opendocument.text", "application/vnd.oasis.opendocument.text-template", "application/vnd.oasis.opendocument.text-master"]);
const MAX_PICTURE_TOTAL = 200 * 1024 * 1024;
const MAX_XML_BYTES = 256 * 1024 * 1024;
const MAX_REPEAT = 200;

interface StyleEntry {
  family: string;
  parent: string | null;
  text: Record<string, string>;
  paragraph: Record<string, string>;
  cell: Record<string, string>;
  column: Record<string, string>;
  masterPage: string | null;
  outlineLevel: number | null;
}

interface ResolvedStyle {
  text: Record<string, string>;
  paragraph: Record<string, string>;
  cell: Record<string, string>;
  column: Record<string, string>;
  outlineLevel: number | null;
  fontSizePt: number;
  name: string;
}

const NUMBER_FORMATS: Record<string, ListNumberFormat> = { "1": "decimal", a: "lowerLetter", A: "upperLetter", i: "lowerRoman", I: "upperRoman" };

function attributesOf(element: DomElement | undefined | null): Record<string, string> {
  return element ? { ...element.attrs } : {};
}

function odfLengthPx(value: string | undefined): number | null {
  return value ? parseLength(value, 16) : null;
}

class OdtReader {
  private readonly styles = new Map<string, StyleEntry>();
  private readonly defaults = new Map<string, StyleEntry>();
  private readonly fontFaces = new Map<string, string>();
  private readonly listStyles = new Map<string, ExplicitListLevel[]>();
  private readonly resolved = new Map<string, ResolvedStyle>();
  readonly lists: Record<string, ExplicitListLevel[]> = {};
  private readonly lastListKey = new Map<string, string>();
  private listSerial = 0;
  private noteCount = 0;
  readonly notes: DomElement[] = [];
  private readonly pictures: Map<string, Uint8Array>;
  missingPictures = 0;
  firstMasterPage: string | null = null;

  constructor(pictures: Map<string, Uint8Array>) {
    this.pictures = pictures;
  }

  // ---- styles ------------------------------------------------------------------------

  readStyles(root: DomElement) {
    for (const face of findElements(root, (element) => element.name === "style:font-face")) {
      const name = face.attrs["style:name"];
      const family = (face.attrs["svg:font-family"] ?? name ?? "").replace(/^['"]|['"]$/g, "").trim();
      if (name && family) this.fontFaces.set(name, family);
    }
    for (const element of findElements(root, (candidate) => candidate.name === "style:style" || candidate.name === "style:default-style")) {
      const family = element.attrs["style:family"] ?? "paragraph";
      const props = (name: string) => attributesOf(childElements(element, name)[0]);
      const entry: StyleEntry = {
        family,
        parent: element.attrs["style:parent-style-name"] ?? null,
        text: props("style:text-properties"),
        paragraph: props("style:paragraph-properties"),
        cell: props("style:table-cell-properties"),
        column: props("style:table-column-properties"),
        masterPage: element.attrs["style:master-page-name"] ?? null,
        outlineLevel: element.attrs["style:default-outline-level"] ? Number(element.attrs["style:default-outline-level"]) : null,
      };
      if (element.name === "style:default-style") this.defaults.set(family, entry);
      else if (element.attrs["style:name"]) this.styles.set(`${family}:${element.attrs["style:name"]}`, entry);
    }
    for (const style of findElements(root, (element) => element.name === "text:list-style")) {
      const name = style.attrs["style:name"];
      if (!name) continue;
      const levels: ExplicitListLevel[] = [];
      for (const level of style.children) {
        if (level.type !== "element") continue;
        const index = Math.min(9, Math.max(1, Number(level.attrs["text:level"] ?? 1))) - 1;
        if (level.name === "text:list-level-style-bullet" || level.name === "text:list-level-style-image") {
          levels[index] = { format: "bullet", bulletChar: level.attrs["text:bullet-char"] || "•" };
        } else if (level.name === "text:list-level-style-number") {
          const numFormat = level.attrs["style:num-format"] ?? "1";
          const format = NUMBER_FORMATS[numFormat];
          if (!format) {
            levels[index] = { format: "bullet", bulletChar: "•" };
            continue;
          }
          const display = Math.min(index + 1, Math.max(1, Number(level.attrs["text:display-levels"] ?? 1)));
          const numbers = Array.from({ length: display }, (_value, offset) => `%${index + 2 - display + offset}`).join(".");
          levels[index] = { format, start: Math.max(0, Number(level.attrs["text:start-value"] ?? 1)), text: `${level.attrs["style:num-prefix"] ?? ""}${numbers}${level.attrs["style:num-suffix"] ?? ""}` };
        }
      }
      for (let i = 0; i < levels.length; i += 1) if (!levels[i]) levels[i] = levels[i - 1] ?? { format: "bullet", bulletChar: "•" };
      if (levels.length) this.listStyles.set(name, levels);
    }
  }

  private resolve(family: string, name: string | undefined): ResolvedStyle {
    const key = `${family}:${name ?? ""}`;
    const cached = this.resolved.get(key);
    if (cached) return cached;
    const chain: StyleEntry[] = [];
    const seen = new Set<string>();
    let current = name;
    while (current && !seen.has(current) && chain.length < 32) {
      seen.add(current);
      const entry = this.styles.get(`${family}:${current}`);
      if (!entry) break;
      chain.unshift(entry);
      current = entry.parent ?? undefined;
    }
    const base = this.defaults.get(family);
    if (base) chain.unshift(base);
    const result: ResolvedStyle = { text: {}, paragraph: {}, cell: {}, column: {}, outlineLevel: null, fontSizePt: 12, name: name ?? "" };
    for (const entry of chain) {
      const size = entry.text["fo:font-size"];
      if (size) {
        const percent = /^([\d.]+)%$/.exec(size);
        const px = percent ? null : parseLength(size, result.fontSizePt * 96 / 72);
        if (percent) result.fontSizePt = result.fontSizePt * Number(percent[1]) / 100;
        else if (px) result.fontSizePt = px * 72 / 96;
      }
      Object.assign(result.text, entry.text);
      Object.assign(result.paragraph, entry.paragraph);
      Object.assign(result.cell, entry.cell);
      Object.assign(result.column, entry.column);
      if (entry.outlineLevel !== null) result.outlineLevel = entry.outlineLevel;
    }
    this.resolved.set(key, result);
    return result;
  }

  /** CSS for text properties; `explicit` lists only the properties a style sets itself. */
  private textCss(text: Record<string, string>, fontSizePt: number | null): Record<string, string> {
    const css: Record<string, string> = {};
    const face = text["style:font-name"] ? this.fontFaces.get(text["style:font-name"]) ?? text["style:font-name"] : text["fo:font-family"];
    if (face) css["font-family"] = face.replace(/^['"]|['"]$/g, "");
    if (fontSizePt !== null) css["font-size"] = `${Math.round(fontSizePt * 100) / 100}pt`;
    const weight = text["fo:font-weight"];
    if (weight) css["font-weight"] = weight === "bold" || Number(weight) >= 600 ? "bold" : "normal";
    const style = text["fo:font-style"];
    if (style) css["font-style"] = style === "italic" || style === "oblique" ? "italic" : "normal";
    const decorations: string[] = [];
    const underline = text["style:text-underline-style"];
    const strike = text["style:text-line-through-style"];
    if (underline && underline !== "none") decorations.push("underline");
    if (strike && strike !== "none") decorations.push("line-through");
    if (decorations.length) css["text-decoration"] = decorations.join(" ");
    else if (underline === "none" || strike === "none") css["text-decoration"] = "none";
    if (text["fo:color"]) css.color = text["fo:color"];
    const background = text["fo:background-color"] ?? text["style:text-background-color"];
    if (background && background !== "transparent") css["background-color"] = background;
    const position = (text["style:text-position"] ?? "").trim();
    if (/^super/.test(position) || /^[1-9]\d*%/.test(position)) css["vertical-align"] = "super";
    else if (/^sub/.test(position) || /^-\d+%/.test(position)) css["vertical-align"] = "sub";
    else if (/^0%?/.test(position)) css["vertical-align"] = "baseline";
    if (text["fo:text-transform"] === "uppercase") css["text-transform"] = "uppercase";
    if (text["fo:font-variant"] === "small-caps") css["font-variant"] = "small-caps";
    if (text["text:display"] === "none") css.display = "none";
    return css;
  }

  private paragraphCss(style: ResolvedStyle): Record<string, string> {
    const p = style.paragraph;
    const css: Record<string, string> = this.textCss(style.text, style.fontSizePt);
    const align = p["fo:text-align"];
    if (align) css["text-align"] = align === "start" ? "left" : align === "end" ? "right" : align;
    css["margin-top"] = p["fo:margin-top"] ?? "0";
    css["margin-bottom"] = p["fo:margin-bottom"] ?? "0";
    if (p["fo:margin-left"]) css["margin-left"] = p["fo:margin-left"];
    if (p["fo:margin-right"]) css["margin-right"] = p["fo:margin-right"];
    if (p["fo:text-indent"]) css["text-indent"] = p["fo:text-indent"];
    if (p["fo:line-height"] && p["fo:line-height"] !== "normal") {
      const percent = /^([\d.]+)%$/.exec(p["fo:line-height"]);
      css["line-height"] = percent ? String(Number(percent[1]) / 100) : p["fo:line-height"];
    } else if (p["style:line-height-at-least"]) css["line-height"] = p["style:line-height-at-least"];
    if (p["fo:break-before"] === "page") css["break-before"] = "page";
    if (p["fo:break-after"] === "page") css["break-after"] = "page";
    if (p["fo:keep-with-next"] === "always") css["page-break-after"] = "avoid";
    if (p["fo:background-color"] && p["fo:background-color"] !== "transparent") css["background-color"] = p["fo:background-color"];
    if (/^rl/.test(p["style:writing-mode"] ?? "")) css.direction = "rtl";
    return css;
  }

  // ---- body --------------------------------------------------------------------------

  convertChildren(element: DomElement, target: DomElement, list: { key: string; level: number; marker: boolean } | null = null) {
    for (const child of element.children) {
      const converted = this.convertBlock(child, list);
      for (const node of converted) {
        appendChild(target, node);
        if (list && node.type === "element" && node.attrs["data-simple-list"]) list = { ...list, marker: false };
      }
    }
  }

  private convertBlock(node: DomNode, list: { key: string; level: number; marker: boolean } | null): DomNode[] {
    if (node.type === "text") return node.text.trim() ? [{ type: "text", text: node.text }] : [];
    switch (node.name) {
      case "text:p": case "text:h": return [this.paragraph(node, list)];
      case "text:list": return this.list(node, list);
      case "table:table": return [this.table(node)];
      case "text:section": case "text:index-body": case "text:table-of-content": case "text:alphabetical-index": case "text:illustration-index":
      case "text:table-index": case "text:object-index": case "text:user-index": case "text:bibliography": case "text:index-title":
      case "office:text": case "text:list-item": case "text:list-header": case "draw:text-box": case "table:table-cell": {
        const container = createElement("div");
        this.convertChildren(node, container, list);
        return container.children;
      }
      case "text:soft-page-break":
        return [];
      case "text:sequence-decls": case "text:variable-decls": case "text:user-field-decls": case "text:tracked-changes": case "office:forms":
      case "table:named-expressions": case "office:annotation": case "text:dde-connection-decls":
        return [];
      default:
        return this.inline(node);
    }
  }

  private paragraph(element: DomElement, list: { key: string; level: number; marker: boolean } | null): DomElement {
    const style = this.resolve("paragraph", element.attrs["text:style-name"]);
    if (this.firstMasterPage === null) {
      const entry = this.styles.get(`paragraph:${element.attrs["text:style-name"]}`);
      this.firstMasterPage = entry?.masterPage ?? "";
    }
    let level: number | null = null;
    if (element.name === "text:h") level = Number(element.attrs["text:outline-level"] ?? style.outlineLevel ?? 1);
    else if (style.outlineLevel && /heading/i.test(style.name)) level = style.outlineLevel;
    const tag = level && level >= 1 && level <= 6 ? `h${level}` : "p";
    const css = this.paragraphCss(style);
    const attrs: Record<string, string> = {};
    if (list) {
      if (list.marker) {
        attrs["data-simple-list"] = list.key;
        attrs["data-simple-level"] = String(list.level);
      } else css["margin-left"] = `${48 * (list.level + 1)}px`;
      delete css["text-indent"];
    }
    const paragraph = createElement(tag, attrs, [], css);
    for (const child of element.children) for (const node of this.inline(child)) appendChild(paragraph, node);
    return paragraph;
  }

  private list(element: DomElement, parent: { key: string; level: number; marker: boolean } | null): DomNode[] {
    const level = parent ? Math.min(8, parent.level + 1) : 0;
    let key = parent?.key;
    if (!key) {
      const styleName = element.attrs["text:style-name"] ?? "";
      const continued = element.attrs["text:continue-numbering"] === "true" || element.attrs["text:continue-list"] !== undefined;
      key = continued && this.lastListKey.has(styleName) ? this.lastListKey.get(styleName)! : `${styleName}#${++this.listSerial}`;
      this.lastListKey.set(styleName, key);
      if (!this.lists[key]) this.lists[key] = this.listStyles.get(styleName) ?? [{ format: "bullet", bulletChar: "•" }];
    }
    const nodes: DomNode[] = [];
    for (const item of element.children) {
      if (item.type !== "element") continue;
      if (item.name !== "text:list-item" && item.name !== "text:list-header") continue;
      const container = createElement("div");
      this.convertChildren(item, container, { key, level, marker: item.name === "text:list-item" });
      nodes.push(...container.children);
    }
    return nodes;
  }

  private table(element: DomElement): DomElement {
    const table = createElement("table");
    const colgroup = createElement("colgroup");
    const rows: DomElement[] = [];
    const visit = (parent: DomElement) => {
      for (const child of parent.children) {
        if (child.type !== "element") continue;
        if (child.name === "table:table-column") {
          const repeat = Math.min(MAX_REPEAT, Math.max(1, Number(child.attrs["table:number-columns-repeated"] ?? 1)));
          const width = this.resolve("table-column", child.attrs["table:style-name"]).column["style:column-width"];
          const px = odfLengthPx(width);
          for (let i = 0; i < repeat; i += 1) appendChild(colgroup, createElement("col", px ? { width: String(Math.round(px)) } : {}));
        } else if (child.name === "table:table-row") {
          const repeat = Math.min(MAX_REPEAT, Math.max(1, Number(child.attrs["table:number-rows-repeated"] ?? 1)));
          const row = this.row(child);
          rows.push(row);
          // Repeated rows are usually empty filler; keep at most a few copies.
          for (let i = 1; i < Math.min(repeat, 3); i += 1) rows.push(this.row(child));
        } else if (/^table:table-(?:header-rows|rows|row-group|columns|column-group|header-columns)$/.test(child.name)) visit(child);
      }
    };
    visit(element);
    if (colgroup.children.length) appendChild(table, colgroup);
    for (const row of rows) appendChild(table, row);
    return table;
  }

  private row(element: DomElement): DomElement {
    const row = createElement("tr");
    for (const cell of element.children) {
      if (cell.type !== "element" || cell.name !== "table:table-cell") continue;
      const repeat = Math.min(MAX_REPEAT, Math.max(1, Number(cell.attrs["table:number-columns-repeated"] ?? 1)));
      for (let i = 0; i < repeat; i += 1) {
        const style = this.resolve("table-cell", cell.attrs["table:style-name"]);
        const css: Record<string, string> = {};
        const background = style.cell["fo:background-color"];
        if (background && background !== "transparent") css["background-color"] = background;
        const attrs: Record<string, string> = {};
        if (cell.attrs["table:number-columns-spanned"] && cell.attrs["table:number-columns-spanned"] !== "1") attrs.colspan = cell.attrs["table:number-columns-spanned"];
        if (cell.attrs["table:number-rows-spanned"] && cell.attrs["table:number-rows-spanned"] !== "1") attrs.rowspan = cell.attrs["table:number-rows-spanned"];
        const td = createElement("td", attrs, [], css);
        this.convertChildren(cell, td);
        appendChild(row, td);
      }
    }
    return row;
  }

  private inline(node: DomNode): DomNode[] {
    if (node.type === "text") return [{ type: "text", text: node.text }];
    switch (node.name) {
      case "text:span": {
        const style = this.resolve("text", node.attrs["text:style-name"]);
        const explicit = this.styles.get(`text:${node.attrs["text:style-name"]}`);
        const size = explicit && (explicit.text["fo:font-size"] || explicit.parent) ? style.fontSizePt : null;
        const span = createElement("span", {}, [], this.textCss(style.text, size));
        for (const child of node.children) for (const inner of this.inline(child)) appendChild(span, inner);
        return [span];
      }
      case "text:a": {
        const link = createElement("a", { href: node.attrs["xlink:href"] ?? "" });
        for (const child of node.children) for (const inner of this.inline(child)) appendChild(link, inner);
        return [link];
      }
      case "text:s": return [{ type: "text", text: " ".repeat(Math.min(1000, Math.max(1, Number(node.attrs["text:c"] ?? 1)))), pre: true }];
      case "text:tab": return [{ type: "text", text: "\t", pre: true }];
      case "text:line-break": return [createElement("br")];
      case "text:soft-page-break": case "text:bookmark": case "text:bookmark-start": case "text:bookmark-end": case "text:reference-mark":
      case "text:reference-mark-start": case "text:reference-mark-end": case "office:annotation": case "office:annotation-end": case "text:change":
      case "text:change-start": case "text:change-end": case "text:alphabetical-index-mark": case "text:toc-mark": case "text:user-index-mark":
        return [];
      case "text:note": return this.note(node);
      case "draw:frame": return this.frame(node);
      case "draw:a": {
        const link = createElement("a", { href: node.attrs["xlink:href"] ?? "" });
        for (const child of node.children) for (const inner of this.inline(child)) appendChild(link, inner);
        return [link];
      }
      case "text:p": case "text:h": case "text:list": case "table:table":
        // Block content inside a frame or note.
        return this.convertBlock(node, null);
      case "text:ruby": {
        const base = childElements(node, "text:ruby-base")[0];
        return base ? base.children.flatMap((child) => this.inline(child)) : [];
      }
      default: {
        // Fields (page numbers, dates, references, ...) and unknown wrappers keep their text.
        const nodes: DomNode[] = [];
        for (const child of node.children) nodes.push(...this.inline(child));
        return nodes;
      }
    }
  }

  private note(element: DomElement): DomNode[] {
    const body = childElements(element, "text:note-body")[0];
    if (!body) return [];
    this.noteCount += 1;
    const key = `odt-note-${this.noteCount}`;
    const aside = createElement("aside", { "data-simple-footnote-body": key });
    this.convertChildren(body, aside);
    this.notes.push(aside);
    return [createElement("sup", { "data-simple-footnote": key }, [{ type: "text", text: textContent(childElements(element, "text:note-citation")[0] ?? createElement("x")) }])];
  }

  private frame(element: DomElement): DomNode[] {
    const width = odfLengthPx(element.attrs["svg:width"]);
    const height = odfLengthPx(element.attrs["svg:height"]);
    for (const image of childElements(element, "draw:image")) {
      const href = (image.attrs["xlink:href"] ?? "").replace(/^\.\//, "");
      const bytes = this.pictures.get(href);
      let src: string | null = null;
      if (bytes) {
        const type = sniffImageType(bytes);
        if (type) src = bytesToDataUrl(bytes, type);
      } else {
        const data = childElements(image, "office:binary-data")[0];
        if (data) src = `data:application/octet-stream;base64,${textContent(data).replace(/\s+/g, "")}`;
      }
      if (!src) continue;
      const attrs: Record<string, string> = { src };
      if (width) attrs.width = String(Math.round(width));
      if (height) attrs.height = String(Math.round(height));
      return [createElement("img", attrs)];
    }
    const box = childElements(element, "draw:text-box")[0];
    if (box) {
      const container = createElement("div");
      this.convertChildren(box, container);
      return [container];
    }
    if (childElements(element, "draw:image").length || childElements(element, "draw:object").length || childElements(element, "draw:object-ole").length) {
      this.missingPictures += 1;
      const alt = textContent(childElements(element, "svg:title")[0] ?? childElements(element, "svg:desc")[0] ?? createElement("x")).trim();
      return [createElement("span", {}, [{ type: "text", text: alt ? `[Picture: ${alt}]` : "[Picture]" }], { "font-style": "italic", color: "#6b6b6b" })];
    }
    return [];
  }

  band(root: DomElement, masterName: string | null, kind: "header" | "footer"): DomElement | null {
    const masters = findElements(root, (element) => element.name === "style:master-page");
    const master = masters.find((element) => element.attrs["style:name"] === masterName) ?? masters.find((element) => element.attrs["style:name"] === "Standard") ?? masters[0];
    const band = master ? childElements(master, `style:${kind}`)[0] : undefined;
    if (!band || band.attrs["style:display"] === "false") return null;
    const container = createElement("div", { "data-simple-band": kind });
    this.convertChildren(band, container);
    return textContent(band).trim() || findElements(band, (element) => element.name === "draw:frame", 1).length ? container : null;
  }

  pageSection(root: DomElement, masterName: string | null): Partial<Document["section"]> | undefined {
    const masters = findElements(root, (element) => element.name === "style:master-page");
    const master = masters.find((element) => element.attrs["style:name"] === masterName) ?? masters.find((element) => element.attrs["style:name"] === "Standard") ?? masters[0];
    const layoutName = master?.attrs["style:page-layout-name"];
    const layout = findElements(root, (element) => element.name === "style:page-layout" && element.attrs["style:name"] === layoutName, 1)[0];
    const props = layout ? childElements(layout, "style:page-layout-properties")[0] : undefined;
    if (!props) return undefined;
    const width = odfLengthPx(props.attrs["fo:page-width"]);
    const height = odfLengthPx(props.attrs["fo:page-height"]);
    if (!width || !height || width < 96 || height < 96 || width > 20000 || height > 20000) return undefined;
    const margin = (name: string) => {
      const value = odfLengthPx(props.attrs[`fo:margin-${name}`] ?? props.attrs["fo:margin"]);
      return value !== null && value >= 0 && value < Math.min(width, height) / 2 ? Math.round(value * 100) / 100 : 96;
    };
    return { pageWidthPx: Math.round(width * 100) / 100, pageHeightPx: Math.round(height * 100) / 100, marginPx: { top: margin("top"), right: margin("right"), bottom: margin("bottom"), left: margin("left") } };
  }
}

export function looksLikeOdt(bytes: Uint8Array): boolean {
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return false;
  const head = new TextDecoder("latin1").decode(bytes.subarray(0, Math.min(bytes.length, 200)));
  return head.includes("mimetypeapplication/vnd.oasis.opendocument.text");
}

function isFlatOdt(bytes: Uint8Array): boolean {
  const head = new TextDecoder("utf-8").decode(bytes.subarray(0, Math.min(bytes.length, 2048)));
  return /<office:document[\s>]/.test(head) && /opendocument\.text/.test(head);
}

/** Imports .odt/.ott packages and flat .fodt XML. */
export async function importOdt(input: Uint8Array | ArrayBuffer, options: ImportOptions = {}): Promise<ImportResult> {
  const bytes = toBytes(input);
  let content: string;
  let styles = "";
  let meta = "";
  const pictures = new Map<string, Uint8Array>();
  if (isFlatOdt(bytes)) {
    content = decodeTextBytes(bytes).text;
    styles = content;
    meta = content;
  } else {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(bytes);
    } catch {
      throw new Error("This file is not an OpenDocument text document, or it is damaged.");
    }
    const mimetype = (await zip.file("mimetype")?.async("string"))?.trim();
    if (mimetype && !ODT_TYPES.has(mimetype)) throw new Error("This OpenDocument file is not a text document.");
    const contentFile = zip.file("content.xml");
    if (!contentFile) throw new Error("This OpenDocument file has no document content.");
    // A tiny package can unpack to gigabytes; refuse before inflating.
    for (const name of ["content.xml", "styles.xml", "meta.xml"]) {
      const size = (zip.file(name) as unknown as { _data?: { uncompressedSize?: number } } | null)?._data?.uncompressedSize ?? 0;
      if (size > MAX_XML_BYTES) throw new Error("This OpenDocument file is too large to open as a document.");
    }
    content = await contentFile.async("string");
    styles = (await zip.file("styles.xml")?.async("string")) ?? "";
    meta = (await zip.file("meta.xml")?.async("string")) ?? "";
    // Only pictures the document references are read.
    const referenced = new Set<string>();
    for (const match of content.matchAll(/xlink:href="([^"]+)"/g)) referenced.add(match[1].replace(/^\.\//, "").replace(/&amp;/g, "&"));
    for (const match of styles.matchAll(/xlink:href="([^"]+)"/g)) referenced.add(match[1].replace(/^\.\//, "").replace(/&amp;/g, "&"));
    let total = 0;
    for (const path of referenced) {
      const file = zip.file(path);
      if (!file || /^(?:[a-z]+:|\/)/i.test(path)) continue;
      const declared = (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;
      if (declared > MAX_IMAGE_BYTES || total + declared > MAX_PICTURE_TOTAL) continue;
      const data = await file.async("uint8array");
      total += data.length;
      if (data.length > MAX_IMAGE_BYTES || total > MAX_PICTURE_TOTAL) continue;
      pictures.set(path, data);
    }
  }
  const contentRoot = parseMarkup(content, { xml: true });
  const stylesRoot = styles === content ? contentRoot : parseMarkup(styles, { xml: true });
  const reader = new OdtReader(pictures);
  if (stylesRoot !== contentRoot) reader.readStyles(stylesRoot);
  reader.readStyles(contentRoot);
  const text = findElements(contentRoot, (element) => element.name === "office:text", 1)[0];
  if (!text) throw new Error("This OpenDocument file has no text body.");
  const root = createElement("#document");
  reader.convertChildren(text, root);
  const masterName = reader.firstMasterPage || null;
  for (const kind of ["header", "footer"] as const) {
    const band = reader.band(stylesRoot, masterName, kind);
    if (band) appendChild(root, band);
  }
  for (const note of reader.notes) appendChild(root, note);
  const mapped = domToDocument(root, { ...options, trusted: true, authorStyles: false, lists: reader.lists, section: reader.pageSection(stylesRoot, masterName) });
  const warnings = [...mapped.warnings];
  if (reader.missingPictures) warnings.push(reader.missingPictures === 1 ? "A picture or embedded object could not be read and is shown as a placeholder." : `${reader.missingPictures} pictures or embedded objects could not be read and are shown as placeholders.`);
  const result: ImportResult = { document: mapped.document, format: "odt", warnings };
  const metaRoot = meta === content ? contentRoot : meta ? parseMarkup(meta, { xml: true }) : null;
  const title = metaRoot ? findElements(metaRoot, (element) => element.name === "dc:title", 1)[0] : undefined;
  const titleText = title ? textContent(title).trim() : "";
  if (titleText) result.title = titleText;
  return result;
}
