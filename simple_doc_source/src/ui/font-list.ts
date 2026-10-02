/**
 * The font list behind Simple Docs' font box (Word's font menu, Docs' "More fonts"):
 * every font installed on this computer plus the editor's built-in look-alikes, with
 * recently used fonts and the document's own fonts on top, and type-to-search.
 *
 * Fonts are enumerated locally by the main process; nothing is downloaded. The editor
 * loads a font's files when the window starts, so a window loads the curated set plus
 * the recent fonts and the fonts of recently opened documents (remembered here, in this
 * computer's local storage). A font chosen for the first time is applied at once and
 * kept in the document; when this editor cannot load fonts while running, it is shown
 * with a similar font until Simple Docs starts again.
 *
 * No DOM or runtime imports: Node tests load this module directly.
 */

export const RECENT_FONTS_KEY = "simple-docs:recent-fonts";
export const DOCUMENT_FONTS_KEY = "simple-docs:document-fonts";
export const MAX_RECENT_FONTS = 8;
export const MAX_DOCUMENT_FONTS = 16;

type StorageLike = Pick<Storage, "getItem" | "setItem">;

export type FontState =
  /** The editor shows and exports it now. */
  | "ready"
  /** Installed, but this window started without it: shown with a similar font until restart. */
  | "pending"
  /** Not installed on this computer (only the document uses it): shown with a similar font. */
  | "missing";

export interface FontEntry {
  family: string;
  /** Label in the list (a built-in look-alike shows its source: "Cambria (Caladea)"). */
  label: string;
  state: FontState;
}

export interface FontSection {
  title: string;
  fonts: FontEntry[];
}

export interface FontCatalogView {
  /** Installed families (from the main process). */
  installed: readonly string[];
  /** Families the editor has loaded (lower case). */
  ready: ReadonlySet<string>;
  /** The engine's built-in look-alikes: family and label. */
  builtin: ReadonlyArray<{ family: string; label: string }>;
}

/** Collapses spaces; a CSS stack ("Calibri, sans-serif") becomes its first family. */
export function normalizeFamily(family: string | null | undefined): string {
  return String(family ?? "").split(",")[0].replace(/^\s*["']|["']\s*$/g, "").trim().replace(/\s+/g, " ");
}

export const familyKey = (family: string | null | undefined) => normalizeFamily(family).toLowerCase();

function readList(storage: StorageLike | null | undefined, key: string): string[] {
  try {
    const value = JSON.parse(storage?.getItem(key) ?? "[]");
    return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string").map(normalizeFamily).filter(Boolean) : [];
  } catch {
    return [];
  }
}

function writeList(storage: StorageLike | null | undefined, key: string, list: readonly string[]) {
  try {
    storage?.setItem(key, JSON.stringify(list));
  } catch {
    // Local storage is a convenience: fonts still work without it.
  }
}

/** `first` then `rest`, without duplicates (ignoring case and extra spaces), at most `max`. */
function mergeLists(first: readonly string[], rest: readonly string[], max: number): string[] {
  const unique: string[] = [];
  for (const family of [...first, ...rest].map(normalizeFamily).filter(Boolean)) {
    if (!unique.some((item) => familyKey(item) === familyKey(family))) unique.push(family);
  }
  return unique.slice(0, max);
}

export function readRecentFonts(storage: StorageLike | null | undefined): string[] {
  return readList(storage, RECENT_FONTS_KEY).slice(0, MAX_RECENT_FONTS);
}

/** Puts a font first in the recent list (most recent first). Returns the new list. */
export function rememberRecentFont(storage: StorageLike | null | undefined, family: string): string[] {
  const list = mergeLists([family], readRecentFonts(storage), MAX_RECENT_FONTS);
  writeList(storage, RECENT_FONTS_KEY, list);
  return list;
}

export function readDocumentFonts(storage: StorageLike | null | undefined): string[] {
  return readList(storage, DOCUMENT_FONTS_KEY).slice(0, MAX_DOCUMENT_FONTS);
}

/** Remembers installed fonts a document used that the editor had not loaded, for the next start. */
export function rememberDocumentFonts(storage: StorageLike | null | undefined, families: readonly string[]): string[] {
  if (!families.length) return readDocumentFonts(storage);
  const list = mergeLists(families, readDocumentFonts(storage), MAX_DOCUMENT_FONTS);
  writeList(storage, DOCUMENT_FONTS_KEY, list);
  return list;
}

/** The installed families a starting window asks the main process for: recent fonts, then document fonts. */
export function startupFontRequests(storage: StorageLike | null | undefined): string[] {
  return mergeLists(readRecentFonts(storage), readDocumentFonts(storage), MAX_RECENT_FONTS + MAX_DOCUMENT_FONTS);
}

interface RunLike { style?: { fontFamily?: string } | null }
interface BlockLike { kind?: string; runs?: RunLike[]; rows?: Array<{ cells?: Array<{ blocks?: BlockLike[] }> }>; text?: { blocks?: BlockLike[] } }

/** Font families a document's text uses (body, tables, text boxes, headers, footers, notes); every run names its font. */
export function documentFontFamilies(doc: {
  blocks?: readonly unknown[];
  section?: object | null;
  footnotes?: object | null;
  endnotes?: object | null;
} | null | undefined): string[] {
  const found = new Map<string, string>();
  const add = (family: string | undefined) => {
    const name = normalizeFamily(family);
    if (name && !found.has(name.toLowerCase())) found.set(name.toLowerCase(), name);
  };
  const visit = (blocks: readonly unknown[] | undefined) => {
    for (const raw of blocks ?? []) {
      const block = raw as BlockLike;
      if (block?.kind === "paragraph") for (const run of block.runs ?? []) add(run.style?.fontFamily);
      else if (block?.kind === "table") for (const row of block.rows ?? []) for (const cell of row.cells ?? []) visit(cell.blocks);
      else if (block?.kind === "shape") visit(block.text?.blocks);
    }
  };
  visit(doc?.blocks);
  const section = (doc?.section ?? {}) as Record<string, unknown>;
  for (const band of ["header", "footer", "headerFirst", "headerEven", "footerFirst", "footerEven"]) {
    const blocks = section[band];
    if (Array.isArray(blocks)) visit(blocks);
  }
  for (const notes of [doc?.footnotes, doc?.endnotes]) for (const blocks of Object.values((notes ?? {}) as Record<string, unknown>)) if (Array.isArray(blocks)) visit(blocks);
  return [...found.values()];
}

/** How well a family matches a query: lower is better, null for no match. */
export function matchRank(family: string, query: string): number | null {
  const name = family.toLowerCase();
  const wanted = query.trim().toLowerCase().replace(/\s+/g, " ");
  if (!wanted) return 3;
  if (name === wanted) return 0;
  if (name.startsWith(wanted)) return 1;
  if (name.split(/[\s-]+/).some((word) => word.startsWith(wanted)) || name.includes(` ${wanted}`)) return 2;
  if (name.includes(wanted)) return 3;
  // "centgoth" finds Century Gothic: every query letter in order.
  let at = 0;
  for (const char of name) if (char === wanted[at]) at += 1;
  return at === wanted.length && wanted.length >= 3 ? 4 : null;
}

function entryFor(family: string, view: FontCatalogView): FontEntry {
  const key = familyKey(family);
  const builtin = view.builtin.find((item) => familyKey(item.family) === key);
  const installed = view.installed.some((item) => familyKey(item) === key);
  const state: FontState = view.ready.has(key) || (builtin && !installed) ? "ready" : installed ? "pending" : "missing";
  return { family: builtin && !installed && !view.ready.has(key) ? builtin.family : normalizeFamily(family), label: builtin && !installed && !view.ready.has(key) ? builtin.label : normalizeFamily(family), state };
}

/** Every font the list offers (installed and built-in), sorted by name. */
export function allFonts(view: FontCatalogView): FontEntry[] {
  const families = new Map<string, string>();
  for (const family of view.installed) families.set(familyKey(family), normalizeFamily(family));
  for (const item of view.builtin) if (!families.has(familyKey(item.family))) families.set(familyKey(item.family), normalizeFamily(item.family));
  return [...families.values()].sort((a, b) => a.localeCompare(b, "en", { sensitivity: "base" })).map((family) => entryFor(family, view));
}

/**
 * The sections the font list shows. Without a query: recently used fonts, the
 * document's fonts, then all fonts. With a query: the matching fonts, best first.
 */
export function fontSections(view: FontCatalogView, options: { query?: string; recent?: readonly string[]; document?: readonly string[] } = {}): FontSection[] {
  const fonts = allFonts(view);
  const query = String(options.query ?? "").trim();
  if (query) {
    const ranked = fonts
      .map((entry) => ({ entry, rank: matchRank(entry.family, query) }))
      .filter((item): item is { entry: FontEntry; rank: number } => item.rank !== null)
      .sort((a, b) => a.rank - b.rank || a.entry.family.localeCompare(b.entry.family, "en", { sensitivity: "base" }))
      .map((item) => item.entry);
    // A missing document font is still findable by name.
    for (const family of options.document ?? []) {
      const rank = matchRank(normalizeFamily(family), query);
      if (rank !== null && !ranked.some((entry) => familyKey(entry.family) === familyKey(family))) ranked.push(entryFor(family, view));
    }
    return [{ title: "Matching fonts", fonts: ranked }];
  }
  const sections: FontSection[] = [];
  const known = (family: string) => fonts.find((entry) => familyKey(entry.family) === familyKey(family));
  const recent = (options.recent ?? []).map((family) => known(family)).filter((entry): entry is FontEntry => Boolean(entry)).slice(0, MAX_RECENT_FONTS);
  if (recent.length) sections.push({ title: "Recently used", fonts: recent });
  const inDocument = (options.document ?? [])
    .filter((family) => !recent.some((entry) => familyKey(entry.family) === familyKey(family)))
    .map((family) => known(family) ?? entryFor(family, view));
  if (inDocument.length) sections.push({ title: "In this document", fonts: inDocument });
  sections.push({ title: "All fonts", fonts });
  return sections;
}
