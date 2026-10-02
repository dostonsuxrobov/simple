/**
 * Native document importers (DOC-017, DOC-SIE-16): .txt, .md, .html/.htm, .rtf and
 * .odt open without LibreOffice or any network access. Every importer returns a complete
 * WordCanvas document for `handle.setDocument()` plus plain-language warnings; the
 * caller opens it untitled (Save As .docx), since these formats cannot hold everything
 * the editor can.
 *
 *   const result = await importDocument(bytes, { fileName: "notes.md" });
 *   handle.setDocument(result.document);
 *
 * To insert an import into an open document instead, pass
 * `blocksForInsert(result.document, handle.getDocument())` to engine-bridge
 * `insertBlocks()`: inserted blocks may only reference lists and notes that exist in the
 * target, so missing ones become text markers and appended note text.
 */
import { importHtml } from "./html.ts";
import type { HtmlImportOptions } from "./html.ts";
import { importMarkdown } from "./markdown.ts";
import { importOdt, looksLikeOdt } from "./odt.ts";
import { importRtf, looksLikeRtf } from "./rtf.ts";
import { importText } from "./text.ts";
import type { ImportFormat, ImportOptions, ImportResult } from "./model.ts";
import { toBytes } from "./text-decoding.ts";

export type { ImportFormat, ImportOptions, ImportResult } from "./model.ts";
export type { MapOptions, MapResult, ExplicitListLevel } from "./dom-to-model.ts";
export type { DomElement, DomNode, DomText } from "./markup.ts";
export { domToDocument, markerFormat, safeHref } from "./dom-to-model.ts";
export { parseMarkup, decodeEntities } from "./markup.ts";
export { decodeTextBytes, guessSingleByteEncoding } from "./text-decoding.ts";
export { importHtml, htmlToDocument, sanitizeHtmlTree } from "./html.ts";
export { importMarkdown, markdownToHtml } from "./markdown.ts";
export { importOdt } from "./odt.ts";
export { importRtf, parseRtfGroups } from "./rtf.ts";
export { importText, textToDocument } from "./text.ts";
export { blocksForInsert, formatListNumber } from "./insert.ts";

/** Largest file the importers accept. */
export const MAX_IMPORT_BYTES = 100 * 1024 * 1024;

/** File extensions (lower case, no dot) and the importer that reads them. */
export const IMPORT_EXTENSIONS: Readonly<Record<string, ImportFormat>> = Object.freeze({
  txt: "txt", text: "txt", log: "txt",
  md: "md", markdown: "md", mdown: "md", mkd: "md",
  html: "html", htm: "html", xhtml: "html",
  rtf: "rtf",
  odt: "odt", ott: "odt", fodt: "odt",
});

export function importFormatForName(fileName: string | undefined | null): ImportFormat | null {
  const match = /\.([A-Za-z0-9]+)$/.exec(fileName ?? "");
  return match ? IMPORT_EXTENSIONS[match[1].toLowerCase()] ?? null : null;
}

/**
 * The importer for a file, from its content first (an RTF or OpenDocument file with the
 * wrong extension still opens correctly), then from its name.
 */
export function sniffImportFormat(input: Uint8Array | ArrayBuffer, fileName?: string): ImportFormat | null {
  const bytes = toBytes(input);
  if (looksLikeRtf(bytes)) return "rtf";
  if (looksLikeOdt(bytes)) return "odt";
  const byName = importFormatForName(fileName);
  const head = new TextDecoder("utf-8").decode(bytes.subarray(0, Math.min(bytes.length, 1024))).replace(/^\ufeff/, "").trimStart().toLowerCase();
  if (/^<\?xml[^>]*>\s*<office:document[\s>]/.test(head) || (head.startsWith("<office:document") && head.includes("opendocument.text"))) return "odt";
  if (byName === "md" || byName === "txt") return byName;
  if (/^(?:<!doctype html|<html[\s>]|<head[\s>]|<body[\s>]|<meta[\s>])/.test(head)) return "html";
  return byName;
}

export interface ImportDocumentOptions extends ImportOptions, HtmlImportOptions {
  /** File name, used to pick the importer when the content does not say. */
  fileName?: string;
  /** Force an importer. */
  format?: ImportFormat;
}

/** Imports any supported file. Rejects with a plain-language message when it cannot. */
export async function importDocument(input: Uint8Array | ArrayBuffer, options: ImportDocumentOptions = {}): Promise<ImportResult> {
  const bytes = toBytes(input);
  if (bytes.length > MAX_IMPORT_BYTES) throw new Error("This file is too large to open as a document.");
  const format = options.format ?? sniffImportFormat(bytes, options.fileName);
  switch (format) {
    case "txt": return importText(bytes, options);
    case "md": return importMarkdown(bytes, options);
    case "html": return importHtml(bytes, options);
    case "rtf": return importRtf(bytes, options);
    case "odt": return importOdt(bytes, options);
    default: throw new Error("Simple cannot open this kind of file as a document.");
  }
}
