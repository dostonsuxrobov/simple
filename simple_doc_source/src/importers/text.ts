/**
 * Plain text (.txt) import. Like Word, every line becomes a paragraph (no spacing
 * between lines), tabs are kept and a form feed starts a new page.
 */
import { assembleDocument, createIdSource, defaultCharStyle, defaultParaStyle } from "./model.ts";
import type { Block, ImportOptions, ImportResult, Paragraph } from "./model.ts";
import { decodeTextBytes, toBytes } from "./text-decoding.ts";

export interface TextImportOptions extends ImportOptions {
  /** Force a text encoding label (for example "windows-1251"). */
  encoding?: string;
}

/** Characters that never belong in document text (C0 controls except tab/newline/form feed, DEL). */
const CONTROL_CHARACTERS = /[\u0000-\u0008\u000b\u000e-\u001f\u007f]/g;

export function textToDocument(text: string, options: ImportOptions = {}): ImportResult {
  const ids = createIdSource(options.idPrefix);
  const char = defaultCharStyle();
  const blocks: Block[] = [];
  let pageBreak = false;
  const normalized = text.replace(/^\ufeff/, "").replace(/\r\n?|\u2028|\u2029|\u0085/g, "\n").replace(CONTROL_CHARACTERS, "");
  const lines = normalized.split("\n");
  // A final newline ends the last line; it does not add an empty paragraph.
  if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
  for (const line of lines) {
    const pages = line.split("\f");
    pages.forEach((part, index) => {
      if (index > 0) pageBreak = true;
      // "\f" on its own line, or at either end of a line, only starts a page.
      if (!part && pages.length > 1) return;
      const style = { ...defaultParaStyle(), spaceAfterPx: 0 };
      if (pageBreak) style.pageBreakBefore = true;
      pageBreak = false;
      const paragraph: Paragraph = { kind: "paragraph", id: ids("p"), revision: 0, runs: [{ text: part, style: { ...char } }], style };
      blocks.push(paragraph);
    });
  }
  if (pageBreak) {
    blocks.push({ kind: "paragraph", id: ids("p"), revision: 0, runs: [{ text: "", style: { ...char } }], style: { ...defaultParaStyle(), spaceAfterPx: 0, pageBreakBefore: true } });
  }
  const emptyParagraph = (): Paragraph => ({ kind: "paragraph", id: ids("p"), revision: 0, runs: [{ text: "", style: { ...char } }], style: { ...defaultParaStyle(), spaceAfterPx: 0 } });
  return { document: assembleDocument({ blocks, emptyParagraph }), format: "txt", warnings: [] };
}

export function importText(input: Uint8Array | ArrayBuffer | string, options: TextImportOptions = {}): ImportResult {
  if (typeof input === "string") return textToDocument(input, options);
  const decoded = decodeTextBytes(toBytes(input), { encoding: options.encoding });
  const result = textToDocument(decoded.text, options);
  result.encoding = decoded.encoding;
  return result;
}
