/**
 * Headers, footers and page numbers (Word's Insert > Header / Footer / Page Number and
 * Docs' Insert > Headers & footers / Page numbers) as WordCanvas model operations.
 *
 * Every function returns the engine operations of ONE transaction (one Ctrl+Z):
 * `setSectionBand` for the document's own (last) section, `setParaStyle` on a section
 * break for an earlier section that has bands of its own, `setField` for PAGE and
 * NUMPAGES fields and `setSectionProps` for the starting page number. Earlier sections
 * without their own bands inherit the last section's (Word's "Link to previous"), so a
 * page number added there numbers every page.
 *
 * Page numbers are the engine's own field runs: the text "{page}" (or "{page:roman}")
 * and "{pages}" carrying a `fieldId`, which layout replaces on every page and which
 * DOCX export writes as PAGE and NUMPAGES fields.
 *
 * No DOM or runtime imports: Node tests load this module directly.
 */
import type { BandContainer, Block, CharStyle, ParaStyle, Paragraph, Run, SectionProps } from "@forevka/wordcanvas/query";

export type BandKind = "header" | "footer";
export type PageNumberAlign = "left" | "center" | "right";
/** "number": 1; "page-x": Page 1; "page-x-of-y": Page 1 of 3. */
export type PageNumberText = "number" | "page-x" | "page-x-of-y";
/** OOXML number formats Word offers for page numbers (1, i, I, a, A). */
export type PageNumberFormat = "arabic" | "roman" | "Roman" | "alpha" | "Alpha";

export interface PageNumberOptions {
  position: "top" | "bottom";
  align: PageNumberAlign;
  text: PageNumberText;
  format: PageNumberFormat;
  /** Show the number on the first page; false turns on "Different first page" without it. */
  showOnFirstPage: boolean;
  /** First page number, or null to continue from the previous section (1 for the first). */
  startAt: number | null;
}

/** Engine operation shapes used here (the engine's ops are untyped in its public types). */
export type ModelOp =
  | { type: "setSectionBand"; band: BandContainer; blocks: Block[] | null }
  | { type: "setParaStyle"; blockId: string; patch: Partial<ParaStyle> & { sectionBreak?: SectionBreak } }
  | { type: "setField"; id: string; def: FieldDefinition | null }
  | { type: "setSectionProps"; geometry: SectionGeometry };

export interface FieldDefinition {
  id: string;
  instruction: string;
  name: string;
  kind: "builtin";
  spec: { type: "PAGE" | "NUMPAGES"; numFmt?: Exclude<PageNumberFormat, "arabic"> };
}

/** The document parts these functions read. */
export interface HeaderFooterDocument {
  blocks: readonly Block[];
  section: SectionProps & Record<string, unknown>;
  stylesheet?: { styles: Array<{ id: string; char?: Partial<CharStyle>; para?: Partial<ParaStyle>; basedOn?: string }>; defaultStyleId: string } | null;
  fields?: Record<string, { name?: string; spec?: { type?: string } } | undefined>;
}

interface SectionBreak {
  type: string;
  props: Partial<SectionProps> & Record<string, unknown>;
}

/** The geometry object the engine's setSectionProps op expects (every field present). */
export interface SectionGeometry {
  pageWidthPx: number;
  pageHeightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
  columns: unknown;
  pageNumberStart: number | null;
  headerDistancePx: number | null;
  footerDistancePx: number | null;
  pageColorHex: string | null;
  pageBorders: unknown;
  breakType: string | null;
  lineNumbering: unknown;
}

export const BAND_CONTAINERS: readonly BandContainer[] = ["header", "footer", "headerFirst", "headerEven", "footerFirst", "footerEven"];
const PAGE_TOKEN = /\{page(?::[A-Za-z]+)?\}/;
const PAGES_TOKEN = /\{pages\}/;

export type IdFactory = () => string;

let fallbackSerial = 0;
/** A block or field id that does not collide with the engine's ids. */
export const newModelId: IdFactory = () => `simple-hf-${Date.now().toString(36)}-${(fallbackSerial++).toString(36)}`;

function plainText(block: Paragraph): string {
  return block.runs.map((run) => run.text).join("");
}

function paragraphs(blocks: readonly Block[] | null | undefined): Paragraph[] {
  return (blocks ?? []).filter((block): block is Paragraph => block?.kind === "paragraph");
}

/** True when a run is a PAGE or NUMPAGES field result (or the bare layout token). */
export function isPageNumberRun(run: Run, fields?: HeaderFooterDocument["fields"]): boolean {
  const fieldId = (run.style as { fieldId?: string } | undefined)?.fieldId;
  const field = fieldId ? fields?.[fieldId] : undefined;
  const type = field?.spec?.type ?? field?.name;
  if (type === "PAGE" || type === "NUMPAGES") return true;
  return PAGE_TOKEN.test(run.text) || PAGES_TOKEN.test(run.text);
}

/** A paragraph holding a page number (a PAGE field). */
export function hasPageNumber(block: Block | null | undefined, fields?: HeaderFooterDocument["fields"]): boolean {
  return block?.kind === "paragraph" && block.runs.some((run) => {
    const fieldId = (run.style as { fieldId?: string } | undefined)?.fieldId;
    const field = fieldId ? fields?.[fieldId] : undefined;
    return (field?.spec?.type ?? field?.name) === "PAGE" || PAGE_TOKEN.test(run.text);
  });
}

/** Whether the document numbers its pages somewhere in a header or footer. */
export function documentHasPageNumbers(doc: HeaderFooterDocument): boolean {
  return bandSources(doc).some((source) => BAND_CONTAINERS.some((band) => paragraphs(source.bands[band]).some((block) => hasPageNumber(block, doc.fields))));
}

/** The character style a header or footer paragraph starts with: the document's Normal text. */
export function bandCharStyle(doc: HeaderFooterDocument): CharStyle {
  const sheet = doc.stylesheet;
  const byId = new Map((sheet?.styles ?? []).map((style) => [style.id, style]));
  const chain: Array<Partial<CharStyle>> = [];
  const seen = new Set<string>();
  let current = sheet ? byId.get(sheet.defaultStyleId) : undefined;
  while (current && !seen.has(current.id)) {
    seen.add(current.id);
    chain.unshift(current.char ?? {});
    current = current.basedOn ? byId.get(current.basedOn) : undefined;
  }
  const resolved = Object.assign({}, ...chain) as Partial<CharStyle>;
  // What body text really looks like wins: runs carry their own font, size and color, and
  // a document can differ from its stylesheet (Simple's blank document does). Body text
  // is the first Normal-style paragraph with text, else the first paragraph.
  const all = paragraphs(doc.blocks);
  const normal = (block: Paragraph) => !block.style.namedStyle || block.style.namedStyle === sheet?.defaultStyleId;
  const sample = all.find((block) => normal(block) && block.runs.some((run) => run.text.trim()))?.runs.find((run) => run.text.trim())?.style
    ?? all[0]?.runs[0]?.style;
  return {
    fontFamily: sample?.fontFamily ?? resolved.fontFamily ?? "Calibri",
    fontSizePx: sample?.fontSizePx ?? resolved.fontSizePx ?? 14.667,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: sample?.color ?? resolved.color ?? "#000000",
  };
}

/** Word's Header and Footer styles: Normal text, single spacing, nothing before or after. */
export function bandParagraphStyle(align: PageNumberAlign = "left"): ParaStyle {
  return { align, lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
}

export function emptyBandParagraph(doc: HeaderFooterDocument, newId: IdFactory = newModelId): Paragraph {
  return { kind: "paragraph", id: newId(), revision: 0, runs: [{ text: "", style: bandCharStyle(doc) }], style: bandParagraphStyle() };
}

/** The field instruction Word writes for a page number (` PAGE \* roman \* MERGEFORMAT `). */
export function fieldDefinition(id: string, type: "PAGE" | "NUMPAGES", format: PageNumberFormat = "arabic"): FieldDefinition {
  const switches: Record<PageNumberFormat, string> = { arabic: "", roman: " \\* roman", Roman: " \\* ROMAN", alpha: " \\* alphabetic", Alpha: " \\* ALPHABETIC" };
  const numbered = type === "PAGE" && format !== "arabic";
  return {
    id,
    instruction: ` ${type}${type === "PAGE" ? switches[format] : ""} \\* MERGEFORMAT `,
    name: type,
    kind: "builtin",
    spec: numbered ? { type, numFmt: format as Exclude<PageNumberFormat, "arabic"> } : { type },
  };
}

/** The runs of a page number paragraph, and the fields they need. */
export function pageNumberRuns(text: PageNumberText, format: PageNumberFormat, style: CharStyle, newId: IdFactory = newModelId): { runs: Run[]; fields: FieldDefinition[] } {
  const page = fieldDefinition(newId(), "PAGE", format);
  const token = format === "arabic" ? "{page}" : `{page:${format}}`;
  const pageRun: Run = { text: token, style: { ...style, fieldId: page.id } as CharStyle };
  if (text === "number") return { runs: [pageRun], fields: [page] };
  if (text === "page-x") return { runs: [{ text: "Page ", style }, pageRun], fields: [page] };
  const pages = fieldDefinition(newId(), "NUMPAGES");
  return {
    runs: [{ text: "Page ", style }, pageRun, { text: " of ", style }, { text: "{pages}", style: { ...style, fieldId: pages.id } as CharStyle }],
    fields: [page, pages],
  };
}

/** Where bands are stored: the last section (on the document) and every earlier section that has bands of its own. */
interface BandSource {
  kind: "document" | "break";
  /** The section-break paragraph (for "break"). */
  blockId?: string;
  sectionBreak?: SectionBreak;
  bands: Partial<Record<BandContainer, Block[]>>;
}

function bandSources(doc: HeaderFooterDocument): BandSource[] {
  const sources: BandSource[] = [{ kind: "document", bands: doc.section as Partial<Record<BandContainer, Block[]>> }];
  for (const block of doc.blocks) {
    const sectionBreak = block.kind === "paragraph" ? (block.style as { sectionBreak?: SectionBreak }).sectionBreak : undefined;
    if (!sectionBreak?.props) continue;
    const props = sectionBreak.props as Partial<Record<BandContainer, Block[]>>;
    if (BAND_CONTAINERS.some((band) => Array.isArray(props[band]))) sources.push({ kind: "break", blockId: block.id, sectionBreak, bands: props });
  }
  return sources;
}

/** Collects band changes per source and turns them into ops. */
function bandWriter() {
  const changes = new Map<BandSource, Partial<Record<BandContainer, Block[] | null>>>();
  return {
    set(source: BandSource, band: BandContainer, blocks: Block[] | null) {
      const entry = changes.get(source) ?? {};
      entry[band] = blocks;
      changes.set(source, entry);
    },
    current(source: BandSource, band: BandContainer): Block[] | undefined {
      const entry = changes.get(source);
      if (entry && band in entry) return entry[band] ?? undefined;
      return source.bands[band];
    },
    ops(): ModelOp[] {
      const ops: ModelOp[] = [];
      for (const [source, bands] of changes) {
        if (source.kind === "document") {
          for (const band of BAND_CONTAINERS) if (band in bands) ops.push({ type: "setSectionBand", band, blocks: bands[band] ?? null });
          continue;
        }
        const props: Record<string, unknown> = { ...source.sectionBreak!.props };
        for (const band of BAND_CONTAINERS) {
          if (!(band in bands)) continue;
          if (bands[band]) props[band] = bands[band];
          else delete props[band];
        }
        ops.push({ type: "setParaStyle", blockId: source.blockId!, patch: { sectionBreak: { ...source.sectionBreak!, props } } });
      }
      return ops;
    },
  };
}

const firstBandOf = (kind: BandKind): BandContainer => (kind === "header" ? "headerFirst" : "footerFirst");

/** True when the document's last section has a different first page header/footer. */
export function hasDifferentFirstPage(doc: HeaderFooterDocument): boolean {
  return Array.isArray(doc.section.headerFirst) || Array.isArray(doc.section.footerFirst);
}

/** A band's content with fresh ids (a first-page copy must not share ids with the original). */
function copyBlocks(blocks: readonly Block[] | undefined, newId: IdFactory): Block[] {
  const copy = (block: Block): Block => {
    if (block.kind === "table") {
      return { ...block, id: newId(), revision: 0, rows: block.rows.map((row) => ({ ...row, cells: row.cells.map((cell) => ({ ...cell, id: newId(), blocks: cell.blocks.map(copy) })) })) };
    }
    return { ...block, id: newId(), revision: 0 } as Block;
  };
  return (blocks ?? []).map(copy);
}

function withoutPageNumbers(blocks: readonly Block[], fields: HeaderFooterDocument["fields"]): Block[] {
  const kept: Block[] = [];
  for (const block of blocks) {
    if (block.kind !== "paragraph" || !block.runs.some((run) => isPageNumberRun(run, fields))) {
      kept.push(block);
      continue;
    }
    const runs = block.runs.filter((run) => !isPageNumberRun(run, fields));
    const text = runs.map((run) => run.text).join("");
    // "Page 3 of 9" leaves "Page  of " behind: a paragraph that held only the number goes.
    if (!text.replace(/\bpage\b|\bof\b|[\s\-–—|/:.]/gi, "")) continue;
    kept.push({ ...block, runs });
  }
  return kept;
}

/**
 * Turns "Different first page" on or off for the last section. On creates an empty
 * first-page header and footer (Word and Docs), or with `copyDefault` copies of the
 * normal ones without page numbers; off removes them.
 */
export function differentFirstPageOps(doc: HeaderFooterDocument, enabled: boolean, options: { copyDefault?: boolean; newId?: IdFactory } = {}): ModelOp[] {
  const newId = options.newId ?? newModelId;
  if (enabled === hasDifferentFirstPage(doc)) return [];
  if (!enabled) return [{ type: "setSectionBand", band: "headerFirst", blocks: null }, { type: "setSectionBand", band: "footerFirst", blocks: null }];
  const ops: ModelOp[] = [];
  for (const kind of ["header", "footer"] as const) {
    const source = options.copyDefault ? withoutPageNumbers(copyBlocks(doc.section[kind] as Block[] | undefined, newId), doc.fields) : [];
    ops.push({ type: "setSectionBand", band: firstBandOf(kind), blocks: source.length ? source : [emptyBandParagraph(doc, newId)] });
  }
  return ops;
}

/** Whether any section has a header (or footer), first-page and even-page ones included. */
export function hasBand(doc: HeaderFooterDocument, kind: BandKind): boolean {
  return bandSources(doc).some((source) => BAND_CONTAINERS.some((band) => band.startsWith(kind) && Array.isArray(source.bands[band]) && source.bands[band]!.length > 0));
}

/** Ops that give the last section a header (or footer) to type in, when it has none. */
export function ensureBandOps(doc: HeaderFooterDocument, kind: BandKind, newId: IdFactory = newModelId): ModelOp[] {
  const blocks = doc.section[kind] as Block[] | undefined;
  return Array.isArray(blocks) && blocks.length ? [] : [{ type: "setSectionBand", band: kind, blocks: [emptyBandParagraph(doc, newId)] }];
}

/** Word's Remove Header / Remove Footer: every section's header (or footer), first-page and even-page ones included. */
export function removeBandOps(doc: HeaderFooterDocument, kind: BandKind): ModelOp[] {
  const writer = bandWriter();
  const containers = BAND_CONTAINERS.filter((band) => band.startsWith(kind));
  for (const source of bandSources(doc)) {
    for (const band of containers) if (Array.isArray(source.bands[band])) writer.set(source, band, null);
  }
  return writer.ops();
}

/** Removes every page number (PAGE and NUMPAGES fields) from all headers and footers. */
export function removePageNumberOps(doc: HeaderFooterDocument): ModelOp[] {
  const writer = bandWriter();
  for (const source of bandSources(doc)) {
    for (const band of BAND_CONTAINERS) {
      const blocks = source.bands[band];
      if (!Array.isArray(blocks) || !blocks.some((block) => block.kind === "paragraph" && block.runs.some((run) => isPageNumberRun(run, doc.fields)))) continue;
      const kept = withoutPageNumbers(blocks, doc.fields);
      writer.set(source, band, kept.length ? kept : [emptyBandParagraph(doc)]);
    }
  }
  return writer.ops();
}

/** The section geometry the engine's setSectionProps op needs, from the last section. */
export function sectionGeometry(section: SectionProps & Record<string, unknown>, changes: Partial<SectionGeometry> = {}): SectionGeometry {
  const value = <T>(key: string, fallback: T): T => (section[key] === undefined ? fallback : section[key] as T);
  return {
    pageWidthPx: section.pageWidthPx,
    pageHeightPx: section.pageHeightPx,
    marginPx: { ...section.marginPx },
    columns: value("columns", null),
    pageNumberStart: value("pageNumberStart", null),
    headerDistancePx: value("headerDistancePx", null),
    footerDistancePx: value("footerDistancePx", null),
    pageColorHex: value("pageColorHex", null),
    pageBorders: value("pageBorders", null),
    breakType: value("breakType", null),
    lineNumbering: value("lineNumbering", null),
    ...changes,
  };
}

/** The page number options a dialog starts with, read from the document. */
export function currentPageNumberOptions(doc: HeaderFooterDocument): PageNumberOptions {
  const options: PageNumberOptions = { position: "bottom", align: "right", text: "number", format: "arabic", showOnFirstPage: true, startAt: null };
  const start = Number(doc.section.pageNumberStart);
  if (Number.isFinite(start) && start > 0) options.startAt = Math.round(start);
  for (const kind of ["footer", "header"] as const) {
    const block = paragraphs(doc.section[kind] as Block[] | undefined).find((candidate) => hasPageNumber(candidate, doc.fields));
    if (!block) continue;
    options.position = kind === "header" ? "top" : "bottom";
    options.align = block.style.align === "center" || block.style.align === "left" ? block.style.align : "right";
    const text = plainText(block);
    const token = /\{page:([A-Za-z]+)\}/.exec(text)?.[1] as PageNumberFormat | undefined;
    if (token && ["roman", "Roman", "alpha", "Alpha"].includes(token)) options.format = token;
    options.text = PAGES_TOKEN.test(text) ? "page-x-of-y" : /\bpage\b/i.test(text.replace(PAGE_TOKEN, "")) ? "page-x" : "number";
    const first = paragraphs(doc.section[firstBandOf(kind)] as Block[] | undefined);
    options.showOnFirstPage = !hasDifferentFirstPage(doc) || first.some((candidate) => hasPageNumber(candidate, doc.fields));
    break;
  }
  return options;
}

/**
 * Ops that put a page number in every section's header or footer (Word's Insert >
 * Page Number). A paragraph that already holds a page number is replaced, so
 * changing the position, alignment or format never adds a second number; other
 * header/footer text is kept. "Show on first page" off gives the first page copies
 * of the header and footer without the number; Start at sets the first number.
 */
export function insertPageNumberOps(doc: HeaderFooterDocument, options: PageNumberOptions, newId: IdFactory = newModelId): ModelOp[] {
  const kind: BandKind = options.position === "top" ? "header" : "footer";
  const other: BandKind = kind === "header" ? "footer" : "header";
  const style = bandCharStyle(doc);
  const writer = bandWriter();
  const fieldOps: ModelOp[] = [];
  const numberParagraph = (existing?: Paragraph): Paragraph => {
    const runStyle = existing?.runs.find((run) => isPageNumberRun(run, doc.fields))?.style ?? existing?.runs[0]?.style ?? style;
    const clean = { ...runStyle } as CharStyle & { fieldId?: string };
    delete clean.fieldId;
    const made = pageNumberRuns(options.text, options.format, clean, newId);
    for (const field of made.fields) fieldOps.push({ type: "setField", id: field.id, def: field });
    return {
      kind: "paragraph",
      id: existing?.id ?? newId(),
      revision: (existing?.revision ?? 0) + 1,
      runs: made.runs,
      style: { ...(existing?.style ?? bandParagraphStyle()), align: options.align },
    };
  };
  const place = (source: BandSource, band: BandContainer) => {
    const blocks = writer.current(source, band) ?? [];
    const index = blocks.findIndex((block) => hasPageNumber(block, doc.fields));
    if (index >= 0) {
      const next = blocks.slice();
      next[index] = numberParagraph(blocks[index] as Paragraph);
      writer.set(source, band, next);
      return;
    }
    const empty = blocks.length === 0 || (blocks.length === 1 && blocks[0].kind === "paragraph" && plainText(blocks[0] as Paragraph) === "");
    if (empty) writer.set(source, band, [numberParagraph(blocks[0]?.kind === "paragraph" ? { ...(blocks[0] as Paragraph), runs: [] } : undefined)]);
    // Existing header or footer text stays; the number gets its own paragraph after it.
    else writer.set(source, band, [...blocks, numberParagraph()]);
  };
  const removeFrom = (source: BandSource, band: BandContainer) => {
    const blocks = writer.current(source, band);
    if (!Array.isArray(blocks) || !blocks.some((block) => hasPageNumber(block, doc.fields))) return;
    const kept = withoutPageNumbers(blocks, doc.fields);
    writer.set(source, band, kept.length ? kept : [emptyBandParagraph(doc, newId)]);
  };

  const sources = bandSources(doc);
  for (const source of sources) {
    // Moving the number from the footer to the header (or back) takes it out of the other band.
    for (const band of BAND_CONTAINERS.filter((name) => name.startsWith(other))) removeFrom(source, band);
    if (source.kind === "document" || Array.isArray(source.bands[kind])) place(source, kind);
    if (source.kind === "break" && Array.isArray(source.bands[firstBandOf(kind)])) {
      if (options.showOnFirstPage) place(source, firstBandOf(kind));
      else removeFrom(source, firstBandOf(kind));
    }
  }

  const ops: ModelOp[] = [];
  const documentSource = sources[0];
  if (options.showOnFirstPage) {
    if (hasDifferentFirstPage(doc)) place(documentSource, firstBandOf(kind));
  } else if (hasDifferentFirstPage(doc)) {
    removeFrom(documentSource, firstBandOf(kind));
  } else {
    // Hiding the number on the first page turns on "Different first page" with
    // copies of the header and footer, so only the number disappears there.
    for (const band of ["header", "footer"] as const) {
      const copy = withoutPageNumbers(copyBlocks(writer.current(documentSource, band), newId), doc.fields);
      writer.set(documentSource, firstBandOf(band), copy.length ? copy : [emptyBandParagraph(doc, newId)]);
    }
  }
  ops.push(...fieldOps, ...writer.ops());
  const currentStart = Number.isFinite(Number(doc.section.pageNumberStart)) ? Number(doc.section.pageNumberStart) : null;
  const wantedStart = options.startAt !== null && Number.isFinite(options.startAt) ? Math.max(0, Math.round(options.startAt)) : null;
  if (wantedStart !== currentStart) ops.push({ type: "setSectionProps", geometry: sectionGeometry(doc.section, { pageNumberStart: wantedStart }) });
  return ops;
}

/** The last paragraph of a band (where Edit header/footer puts the caret). */
export function bandCaretTarget(doc: HeaderFooterDocument, band: BandContainer): { blockId: string; offset: number } | null {
  const last = paragraphs(doc.section[band] as Block[] | undefined).at(-1);
  return last ? { blockId: last.id, offset: plainText(last).length } : null;
}
