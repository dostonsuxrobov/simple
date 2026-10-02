/**
 * WordCanvas model helpers shared by every importer: Simple's default page, text and
 * heading styles, list definitions, ids, and the import result type.
 */
import type { Block, CharStyle, Document, NamedStyle, Paragraph, ParaStyle, Run, TableBlock } from "@forevka/wordcanvas/query";

export type { Block, CharStyle, Document, NamedStyle, Paragraph, ParaStyle, Run, TableBlock };
export type ListDefinition = NonNullable<Document["lists"]>[string];
export type ListLevel = ListDefinition["levels"][number];
export type ListNumberFormat = ListLevel["format"];
export type TableCell = TableBlock["rows"][number]["cells"][number];
export type ImageBlockModel = Extract<Block, { kind: "image" }>;

export type ImportFormat = "txt" | "md" | "html" | "rtf" | "odt";

export interface ImportResult {
  /** A complete document for `handle.setDocument()`: its lists, footnotes and styles exist. */
  document: Document;
  format: ImportFormat;
  /** Plain-language notes about content that was simplified or left out. */
  warnings: string[];
  /** Document title from the file's metadata, when it has one. */
  title?: string;
  /** Text encoding used to read the file (text-based formats). */
  encoding?: string;
}

export interface ImportOptions {
  /** Prefix for block ids (deterministic ids in tests). Default: a random prefix. */
  idPrefix?: string;
  /** Widest picture in px (default: the text width of a Letter page, 624 px). */
  maxImageWidthPx?: number;
  /**
   * Resolves a picture that is not a data: URL (for example a file next to a Markdown
   * document) to bytes the caller already holds locally. Never fetch from the network.
   */
  resolveImage?: (source: string) => { bytes: Uint8Array; type?: string } | null | undefined;
}

/** Simple's default body text (matches the editor's overrideDefaultStyles). */
export const DEFAULT_FONT = "Calibri";
export const DEFAULT_FONT_SIZE_PX = 14.667;
export const DEFAULT_COLOR = "#111111";
export const DEFAULT_LINE_HEIGHT = 1.15;
export const DEFAULT_SPACE_AFTER_PX = 10.667;
export const MONOSPACE_FONT = "Consolas";
/** Word's Hyperlink character style colour. */
export const LINK_COLOR = "#0563c1";

export const DEFAULT_SECTION: Document["section"] = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

export function defaultCharStyle(): CharStyle {
  return { fontFamily: DEFAULT_FONT, fontSizePx: DEFAULT_FONT_SIZE_PX, bold: false, italic: false, underline: false, strikethrough: false, color: DEFAULT_COLOR };
}

export function defaultParaStyle(): ParaStyle {
  return { align: "left", lineHeight: DEFAULT_LINE_HEIGHT, spaceBeforePx: 0, spaceAfterPx: DEFAULT_SPACE_AFTER_PX, indentFirstLinePx: 0, indentLeftPx: 0 };
}

const pt = (value: number) => Math.round(value * 96 / 72 * 1000) / 1000;

/** Heading looks: Google Docs sizes, Word spacing, Simple's ink colour. */
export const HEADING_STYLES: ReadonlyArray<{ char: Partial<CharStyle>; para: Partial<ParaStyle> }> = [
  { char: { fontSizePx: pt(20), bold: true }, para: { spaceBeforePx: pt(20), spaceAfterPx: pt(6) } },
  { char: { fontSizePx: pt(16), bold: true }, para: { spaceBeforePx: pt(18), spaceAfterPx: pt(6) } },
  { char: { fontSizePx: pt(14), bold: true }, para: { spaceBeforePx: pt(16), spaceAfterPx: pt(4) } },
  { char: { fontSizePx: pt(12), bold: true }, para: { spaceBeforePx: pt(14), spaceAfterPx: pt(4) } },
  { char: { fontSizePx: pt(11), bold: true }, para: { spaceBeforePx: pt(12), spaceAfterPx: pt(4) } },
  { char: { fontSizePx: pt(11), bold: true, italic: true }, para: { spaceBeforePx: pt(12), spaceAfterPx: pt(4) } },
];

export function headingCharStyle(level: number, base: CharStyle = defaultCharStyle()): CharStyle {
  const heading = HEADING_STYLES[Math.min(Math.max(level, 1), 6) - 1];
  return { ...base, bold: false, italic: false, underline: false, strikethrough: false, ...heading.char, fontFamily: DEFAULT_FONT, color: DEFAULT_COLOR };
}

export function headingParaStyle(level: number, base: ParaStyle = defaultParaStyle()): ParaStyle {
  const clamped = Math.min(Math.max(level, 1), 6);
  return { ...base, ...HEADING_STYLES[clamped - 1].para, lineHeight: DEFAULT_LINE_HEIGHT, keepWithNext: true, namedStyle: `Heading${clamped}`, outlineLevel: clamped - 1 };
}

/** The stylesheet every imported document carries (Normal, Title, Heading 1-6, Quote, Code). */
export function defaultStylesheet(): NonNullable<Document["stylesheet"]> {
  const normal = defaultCharStyle();
  const styles: NamedStyle[] = [
    { id: "Normal", name: "Normal", type: "paragraph", char: { ...normal }, para: { ...defaultParaStyle() } },
    { id: "Title", name: "Title", type: "paragraph", basedOn: "Normal", char: { fontSizePx: pt(26) }, para: { spaceAfterPx: pt(3) } },
    { id: "Subtitle", name: "Subtitle", type: "paragraph", basedOn: "Normal", char: { fontSizePx: pt(15), color: "#555555" }, para: { spaceAfterPx: pt(16) } },
  ];
  HEADING_STYLES.forEach((heading, index) => {
    styles.push({ id: `Heading${index + 1}`, name: `Heading ${index + 1}`, type: "paragraph", basedOn: "Normal", char: { ...heading.char }, para: { ...heading.para, keepWithNext: true, outlineLevel: index } });
  });
  styles.push(
    { id: "Quote", name: "Quote", type: "paragraph", basedOn: "Normal", char: { italic: true, color: "#404040" }, para: { indentLeftPx: 48, indentRightPx: 48 } },
    { id: "Code", name: "Code", type: "paragraph", basedOn: "Normal", char: { fontFamily: MONOSPACE_FONT, fontSizePx: pt(10) }, para: { lineHeight: 1, spaceAfterPx: 0 } },
  );
  return { styles, defaultStyleId: "Normal" };
}

const BULLETS = ["•", "◦", "▪"];
const NUMBER_CYCLE: ListNumberFormat[] = ["decimal", "lowerLetter", "lowerRoman"];

/** Word-like level geometry: 0.5in steps with a 0.25in hanging indent. */
export function listLevel(level: number, format: ListNumberFormat, options: { start?: number; bulletChar?: string; text?: string } = {}): ListLevel {
  const base: ListLevel = {
    format,
    text: format === "bullet" ? "" : options.text ?? `%${level + 1}.`,
    indentLeftPx: 48 * (level + 1),
    hangingPx: 24,
    start: Number.isInteger(options.start) && options.start! >= 0 ? options.start! : 1,
  };
  if (format === "bullet") base.bulletChar = options.bulletChar || BULLETS[level % 3];
  return base;
}

/** A nine-level list; `ordered` picks decimal/letter/roman cycling, otherwise bullets. */
export function listDefinition(id: string, ordered: boolean): ListDefinition {
  const levels: ListLevel[] = [];
  for (let level = 0; level < 9; level += 1) levels.push(listLevel(level, ordered ? NUMBER_CYCLE[level % 3] : "bullet"));
  return { id, levels };
}

/** Allocates ids that are unique inside one import. */
export function createIdSource(prefix?: string) {
  const base = prefix ?? `imp${Math.random().toString(36).slice(2, 8)}-`;
  let counter = 0;
  return (kind = "b") => `${base}${kind}${(counter++).toString(36)}`;
}

export function sameCharStyle(a: CharStyle, b: CharStyle): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]) as Set<keyof CharStyle>;
  for (const key of keys) if (a[key] !== b[key]) return false;
  return true;
}

/** Assembles the document object from body blocks and the shared parts. */
export function assembleDocument(parts: {
  blocks: Block[];
  lists?: Record<string, ListDefinition>;
  footnotes?: Record<string, Paragraph[]>;
  section?: Partial<Document["section"]>;
  emptyParagraph: () => Paragraph;
}): Document {
  const blocks = parts.blocks.length ? parts.blocks : [parts.emptyParagraph()];
  // The editor needs a paragraph to place the caret after a trailing table or picture.
  if (blocks[blocks.length - 1].kind !== "paragraph") blocks.push(parts.emptyParagraph());
  const document: Document = {
    section: { ...DEFAULT_SECTION, marginPx: { ...DEFAULT_SECTION.marginPx }, ...parts.section },
    blocks,
    stylesheet: defaultStylesheet(),
  };
  if (parts.lists && Object.keys(parts.lists).length) document.lists = parts.lists;
  if (parts.footnotes && Object.keys(parts.footnotes).length) document.footnotes = parts.footnotes;
  return document;
}
