/**
 * HTML (.html/.htm) import. The file is parsed into an inert tree (markup.ts): scripts,
 * frames, embedded objects and forms are removed, event handlers and script URLs are
 * dropped, and nothing is ever fetched. Pictures embedded as data: URLs are kept; any
 * other picture becomes a placeholder.
 */
import { domToDocument } from "./dom-to-model.ts";
import type { MapOptions } from "./dom-to-model.ts";
import { findElements, parseMarkup, textContent } from "./markup.ts";
import type { DomElement } from "./markup.ts";
import type { ImportOptions, ImportResult } from "./model.ts";
import { decodeTextBytes, toBytes } from "./text-decoding.ts";

export interface HtmlImportOptions extends ImportOptions {
  /** Force a text encoding label. */
  encoding?: string;
}

const REMOVED = new Set(["script", "iframe", "frame", "frameset", "object", "embed", "applet", "noscript", "template", "form-associated", "portal", "fencedframe"]);
const EMBEDDED = new Set(["iframe", "frame", "frameset", "object", "embed", "applet", "video", "audio", "canvas", "portal", "fencedframe"]);
const URL_ATTRIBUTES = new Set(["href", "src", "action", "formaction", "poster", "background", "srcset", "xlink:href", "data", "codebase", "cite", "longdesc", "lowsrc", "dynsrc", "ping"]);

/** The charset an HTML file declares in its first bytes (meta charset / http-equiv). */
export function declaredCharset(bytes: Uint8Array): string | null {
  const head = new TextDecoder("latin1").decode(bytes.subarray(0, Math.min(bytes.length, 4096)));
  const match = /<meta[^>]+charset\s*=\s*["']?\s*([A-Za-z0-9_.:-]+)/i.exec(head) ?? /^<\?xml[^>]+encoding\s*=\s*["']([A-Za-z0-9_.:-]+)/i.exec(head);
  return match ? match[1].toLowerCase() : null;
}

/**
 * Removes active content from a parsed HTML tree in place and reports what was removed.
 * The mapper never renders or loads anything, so this is defence in depth for callers
 * that keep the tree.
 */
export function sanitizeHtmlTree(root: DomElement, options: { keepHints?: boolean } = {}): { scripts: number; embedded: number } {
  const counts = { scripts: 0, embedded: 0 };
  const visit = (element: DomElement) => {
    element.children = element.children.filter((child) => {
      if (child.type !== "element") return true;
      if (REMOVED.has(child.name) || EMBEDDED.has(child.name)) {
        if (child.name === "script" || child.name === "noscript") counts.scripts += 1;
        else if (EMBEDDED.has(child.name)) counts.embedded += 1;
        return false;
      }
      return true;
    });
    for (const name of Object.keys(element.attrs)) {
      const value = element.attrs[name];
      if (name.startsWith("on") || (!options.keepHints && name.startsWith("data-simple-")) || name === "srcdoc") delete element.attrs[name];
      else if (URL_ATTRIBUTES.has(name) && /^\s*(?:javascript|vbscript|livescript):/i.test(value.replace(/[\u0000-\u001f\s]+/g, ""))) delete element.attrs[name];
    }
    if (element.attrs.style && /expression\s*\(|javascript:|behavior\s*:|-moz-binding/i.test(element.attrs.style)) delete element.attrs.style;
    for (const child of element.children) if (child.type === "element") visit(child);
  };
  visit(root);
  return counts;
}

export function htmlTitle(root: DomElement): string | undefined {
  const title = findElements(root, (element) => element.name === "title", 1)[0];
  const text = title ? textContent(title).replace(/\s+/g, " ").trim() : "";
  return text || undefined;
}

/** Imports an HTML string. `mapOptions` lets other front-ends reuse the HTML path. */
export function htmlToDocument(html: string, options: ImportOptions = {}, mapOptions: Partial<MapOptions> = {}): ImportResult {
  const root = parseMarkup(html);
  const removed = sanitizeHtmlTree(root, { keepHints: mapOptions.trusted === true });
  const title = htmlTitle(root);
  const mapped = domToDocument(root, { ...options, ...mapOptions });
  const warnings = [...mapped.warnings];
  if (removed.embedded) warnings.push("Embedded videos, frames and other active content were left out.");
  const result: ImportResult = { document: mapped.document, format: "html", warnings };
  if (title) result.title = title;
  return result;
}

export function importHtml(input: Uint8Array | ArrayBuffer | string, options: HtmlImportOptions = {}): ImportResult {
  if (typeof input === "string") return htmlToDocument(input, options);
  const bytes = toBytes(input);
  const decoded = decodeTextBytes(bytes, { encoding: options.encoding, declared: declaredCharset(bytes) });
  const result = htmlToDocument(decoded.text, options);
  result.encoding = decoded.encoding;
  return result;
}
