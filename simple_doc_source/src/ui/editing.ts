/**
 * Pure editing helpers for Simple Docs commands: paragraph lookup and selection
 * ranges over the WordCanvas model, word boundaries, Word's Shift+F3 case cycle,
 * link and bookmark validation, character-formatting reset and missing heading
 * styles. They read the plain document model only (no engine internals).
 *
 * No DOM or runtime imports: Node tests load this module directly.
 */
import type { CharStyle, DocPosition, DocSelection, NamedStyle, Run } from "@forevka/wordcanvas/query";

/** The parts of a paragraph block these helpers read. */
export interface ParagraphLike {
  kind: "paragraph";
  id: string;
  runs: Run[];
  style: { namedStyle?: string; [key: string]: unknown };
}

export interface StylesheetLike {
  styles: NamedStyle[];
  defaultStyleId: string;
}

/** The parts of a document these helpers read. */
export interface DocumentLike {
  blocks: unknown[];
  section?: object | null;
  footnotes?: Record<string, unknown[]> | null;
  endnotes?: Record<string, unknown[]> | null;
  stylesheet?: StylesheetLike | null;
  bookmarks?: Record<string, unknown> | null;
}

export interface ParagraphRange {
  block: ParagraphLike;
  start: number;
  end: number;
}

const BANDS = ["header", "footer", "headerFirst", "headerEven", "footerFirst", "footerEven"];

function isParagraph(block: unknown): block is ParagraphLike {
  const candidate = block as Partial<ParagraphLike> | null;
  return !!candidate && candidate.kind === "paragraph" && typeof candidate.id === "string" && Array.isArray(candidate.runs);
}

/** Paragraphs of a block list in reading order, including table cells and text boxes. */
function collectParagraphs(blocks: unknown, into: ParagraphLike[]) {
  if (!Array.isArray(blocks)) return into;
  for (const block of blocks) {
    if (isParagraph(block)) {
      into.push(block);
      continue;
    }
    const container = block as { kind?: string; rows?: Array<{ cells?: Array<{ blocks?: unknown[] }> }>; text?: { blocks?: unknown[] } } | null;
    if (container?.kind === "table") {
      for (const row of container.rows ?? []) for (const cell of row.cells ?? []) collectParagraphs(cell.blocks, into);
    } else if (container?.kind === "shape") {
      collectParagraphs(container.text?.blocks, into);
    }
  }
  return into;
}

/** Every story of the document as its own paragraph sequence: the body, each header/footer band, each note. */
export function paragraphStories(doc: DocumentLike | null | undefined): ParagraphLike[][] {
  if (!doc) return [];
  const stories = [collectParagraphs(doc.blocks, [])];
  const section = doc.section as Record<string, unknown> | null | undefined;
  for (const band of BANDS) {
    const list = section?.[band];
    if (Array.isArray(list)) stories.push(collectParagraphs(list, []));
  }
  for (const notes of [doc.footnotes, doc.endnotes]) {
    for (const list of Object.values(notes ?? {})) stories.push(collectParagraphs(list, []));
  }
  return stories.filter((story) => story.length > 0);
}

export function findParagraph(doc: DocumentLike | null | undefined, blockId: string): ParagraphLike | null {
  for (const story of paragraphStories(doc)) {
    const found = story.find((block) => block.id === blockId);
    if (found) return found;
  }
  return null;
}

/** The first paragraph of the body (a table's first cell when the document starts with a table). */
export function firstParagraphId(doc: DocumentLike | null | undefined): string | null {
  return doc ? collectParagraphs(doc.blocks, [])[0]?.id ?? null : null;
}

export function paragraphText(block: { runs?: Run[] } | null | undefined): string {
  return (block?.runs ?? []).map((run) => run.text).join("");
}

/** The selected part of each paragraph, in reading order (a collapsed selection gives one empty range). */
export function selectionRanges(doc: DocumentLike | null | undefined, selection: DocSelection | null | undefined): ParagraphRange[] {
  if (!doc || !selection) return [];
  const { anchor, focus } = selection;
  for (const story of paragraphStories(doc)) {
    const a = story.findIndex((block) => block.id === anchor.blockId);
    const f = story.findIndex((block) => block.id === focus.blockId);
    if (a < 0 || f < 0) continue;
    let [from, to]: DocPosition[] = [anchor, focus];
    let [first, last] = [a, f];
    if (f < a || (f === a && focus.offset < anchor.offset)) {
      [from, to] = [focus, anchor];
      [first, last] = [f, a];
    }
    const ranges: ParagraphRange[] = [];
    for (let index = first; index <= last; index += 1) {
      const block = story[index];
      const length = paragraphText(block).length;
      const start = index === first ? Math.max(0, Math.min(length, from.offset)) : 0;
      const end = index === last ? Math.max(start, Math.min(length, to.offset)) : length;
      ranges.push({ block, start, end });
    }
    return ranges;
  }
  const block = findParagraph(doc, focus.blockId);
  return block ? [{ block, start: focus.offset, end: focus.offset }] : [];
}

export function isCollapsed(selection: DocSelection | null | undefined): boolean {
  return !!selection && selection.anchor.blockId === selection.focus.blockId && selection.anchor.offset === selection.focus.offset;
}

const wordSegmenter = typeof Intl !== "undefined" && "Segmenter" in Intl ? new Intl.Segmenter(undefined, { granularity: "word" }) : null;

/** The word at a caret offset (inside it, or touching its start or end), or null between words. */
export function wordRangeAt(text: string, offset: number): { start: number; end: number } | null {
  const value = String(text ?? "");
  const at = Math.max(0, Math.min(value.length, Math.trunc(offset)));
  const words: Array<{ start: number; end: number }> = [];
  if (wordSegmenter) {
    for (const segment of wordSegmenter.segment(value)) {
      if (segment.isWordLike) words.push({ start: segment.index, end: segment.index + segment.segment.length });
    }
  } else {
    for (const match of value.matchAll(/[\p{L}\p{N}_'’-]+/gu)) words.push({ start: match.index ?? 0, end: (match.index ?? 0) + match[0].length });
  }
  return words.find((word) => word.start <= at && at < word.end)
    ?? words.find((word) => word.end === at)
    ?? words.find((word) => word.start === at)
    ?? null;
}

export type CaseMode = "upper" | "lower" | "title";

/**
 * Word's Shift+F3 cycle: lowercase → UPPERCASE → Capitalize Each Word →
 * lowercase. Text without letters stays as it is (null).
 */
export function nextCaseMode(text: string): CaseMode | null {
  const value = String(text ?? "");
  if (!/\p{L}/u.test(value)) return null;
  const upper = value.toLocaleUpperCase();
  const lower = value.toLocaleLowerCase();
  if (value === lower) return "upper";
  if (value === upper) return "title";
  return "lower";
}

const UNSAFE_LINK = /^(?:javascript|vbscript|data|blob|about):/i;

/** Word and Docs complete bare addresses: example.com → https://example.com, name@example.com → mailto:. */
export function normalizeLinkAddress(raw: string): string {
  const value = String(raw ?? "").trim();
  if (!value) return "";
  if (/^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith("#")) return value;
  if (/^[^\s@/]+@[^\s@/]+\.[^\s@/]+$/.test(value)) return `mailto:${value}`;
  if (/^\\\\/.test(value) || /^[a-z]:[\\/]/i.test(value)) return value;
  if (/^(?:www\.)?[\p{L}\p{N}-]+(?:\.[\p{L}\p{N}-]+)+(?::\d+)?(?:[/?#].*)?$/u.test(value)) return `https://${value}`;
  return value;
}

/** Why an address cannot become a link, or null when it can. */
export function linkAddressProblem(raw: string): string | null {
  const value = String(raw ?? "").trim();
  if (!value) return "Type or paste an address.";
  const address = normalizeLinkAddress(value);
  if (UNSAFE_LINK.test(address)) return "Use a web, email or document address.";
  if (/^(?:https?|mailto|ftp):/i.test(address) && /\s/.test(address)) return "Web and email addresses can't contain spaces.";
  return null;
}

/**
 * Word's bookmark rule: a letter first, then letters, digits or underscores,
 * at most 40 characters. `existing` names are refused except `current` (rename).
 */
export function bookmarkNameProblem(name: string, existing: readonly string[] = [], current?: string): string | null {
  const value = String(name ?? "").trim();
  if (!value) return "Type a bookmark name.";
  if (value.length > 40) return "Bookmark names can have up to 40 characters.";
  if (!/^\p{L}[\p{L}\p{N}_]*$/u.test(value)) return /\s/.test(value)
    ? "Bookmark names can't contain spaces. Use an underscore instead, like Chapter_one."
    : "Bookmark names start with a letter and use only letters, numbers and underscores.";
  if (value !== current && existing.includes(value)) return "A bookmark with this name already exists.";
  return null;
}

/** Character properties a named paragraph style gives its text (its basedOn chain resolved, root first). */
export function resolveStyleChar(stylesheet: StylesheetLike | null | undefined, styleId: string | undefined): Partial<CharStyle> {
  if (!stylesheet) return {};
  const byId = new Map(stylesheet.styles.map((style) => [style.id, style]));
  const chain: NamedStyle[] = [];
  const seen = new Set<string>();
  let current = byId.get(styleId ?? "") ?? byId.get(stylesheet.defaultStyleId);
  while (current && !seen.has(current.id)) {
    seen.add(current.id);
    chain.unshift(current);
    current = current.basedOn ? byId.get(current.basedOn) : undefined;
  }
  const defaults = byId.get(stylesheet.defaultStyleId);
  if (defaults && !seen.has(defaults.id)) chain.unshift(defaults);
  return Object.assign({}, ...chain.map((style) => style.char ?? {}));
}

/** Manual character formatting that Ctrl+Space removes; links, notes, fields and controls stay. */
const MANUAL_CHARACTER_FORMATTING = [
  "bold", "italic", "underline", "strikethrough", "doubleStrikethrough", "underlineStyle", "underlineColor",
  "highlightColor", "verticalAlign", "letterSpacingPx", "caps", "smallCaps", "positionPx", "widthScalePct",
  "kerningMinPx", "emphasisMark", "charStyleId", "fontFamily", "fontSizePx", "color",
];

/** Runs with the manual character formatting in [start, end) replaced by the paragraph style's (`base`). */
export function clearCharacterRuns(runs: readonly Run[], start: number, end: number, base: Partial<CharStyle>): Run[] {
  const out: Run[] = [];
  let at = 0;
  const reset = (style: CharStyle): CharStyle => {
    const next: Record<string, unknown> = { ...style };
    for (const key of MANUAL_CHARACTER_FORMATTING) delete next[key];
    Object.assign(next, { bold: false, italic: false, underline: false, strikethrough: false }, base);
    if (next.fontFamily === undefined) next.fontFamily = style.fontFamily;
    if (next.fontSizePx === undefined) next.fontSizePx = style.fontSizePx;
    if (next.color === undefined) next.color = style.color;
    return next as unknown as CharStyle;
  };
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
    out.push({ text: run.text.slice(a, b), style: reset(run.style) });
    if (b < run.text.length) out.push({ text: run.text.slice(b), style: run.style });
  }
  return out;
}

/** The run style at a caret offset (the character before the caret, like typing would continue it). */
export function runStyleAt(block: { runs?: Run[] } | null | undefined, offset: number): CharStyle | null {
  const runs = block?.runs ?? [];
  let at = 0;
  for (const run of runs) {
    const end = at + run.text.length;
    if (offset > at && offset <= end) return run.style;
    at = end;
  }
  return runs[0]?.style ?? null;
}

/** The whole hyperlink around an offset: the touching runs that share one address. */
export function linkRangeAt(block: { runs?: Run[] } | null | undefined, offset: number): { start: number; end: number; link: string } | null {
  const runs = block?.runs ?? [];
  const spans: Array<{ start: number; end: number; link?: string }> = [];
  let at = 0;
  for (const run of runs) {
    spans.push({ start: at, end: at + run.text.length, link: run.style?.link });
    at += run.text.length;
  }
  let index = spans.findIndex((span) => span.link && span.start <= offset && offset < span.end);
  if (index < 0) index = spans.findIndex((span) => span.link && span.end === offset && offset > span.start);
  if (index < 0) return null;
  const link = spans[index].link!;
  let first = index;
  let last = index;
  while (first > 0 && spans[first - 1].link === link) first -= 1;
  while (last < spans.length - 1 && spans[last + 1].link === link) last += 1;
  return { start: spans[first].start, end: spans[last].end, link };
}

const WORD_HEADINGS: Record<number, { sizePx: number; color: string; beforePx: number }> = {
  1: { sizePx: 21.333, color: "#2f5496", beforePx: 16 },
  2: { sizePx: 17.333, color: "#2f5496", beforePx: 2.667 },
  3: { sizePx: 16, color: "#1f3763", beforePx: 2.667 },
};

/**
 * A heading style for a document that has none at this level (Word documents
 * only carry the styles they use). It builds on the previous heading when the
 * document has one, so it matches the document's own look; otherwise it uses
 * Word's default heading look on the document's body style.
 */
export function headingStyleDefinition(level: 1 | 2 | 3, stylesheet: StylesheetLike | null | undefined): NamedStyle {
  const id = `Heading${level}`;
  const previous = level > 1 ? stylesheet?.styles.find((style) => style.id === `Heading${level - 1}`) : undefined;
  if (previous) {
    const sizes: Record<number, number> = { 2: 19, 3: 16 };
    return { id, name: `Heading ${level}`, basedOn: previous.id, char: { fontSizePx: sizes[level] }, para: { spaceBeforePx: level === 2 ? 14 : 12, spaceAfterPx: level === 2 ? 6 : 4, keepWithNext: true, outlineLevel: level - 1 } };
  }
  const look = WORD_HEADINGS[level];
  return {
    id,
    name: `Heading ${level}`,
    ...(stylesheet?.defaultStyleId ? { basedOn: stylesheet.defaultStyleId } : {}),
    char: { fontSizePx: look.sizePx, color: look.color },
    para: { spaceBeforePx: look.beforePx, spaceAfterPx: 0, keepWithNext: true, keepLinesTogether: true, outlineLevel: level - 1 },
  };
}
