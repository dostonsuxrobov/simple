/**
 * AutoFormat as you type, Word-style, for the engine's insert-text hook.
 *
 * The engine calls the hook synchronously after it inserted typed text (or after Enter)
 * and applies the returned edits as one separate transaction, so a single Ctrl+Z puts
 * back exactly what was typed and the rule does not fire again on that text: each rule
 * only looks at the character that was just typed.
 *
 * Rules (each can be turned off):
 *  - smart quotes, chosen by locale (“ ” ‘ ’, „ “ ‚ ‘, « » ...; Uzbek oʻ/gʻ and ʼ);
 *  - "--" to an en dash and "---" to an em dash once the next character is typed, and
 *    " - " between words to an en dash;
 *  - (c) (r) (tm) to © ® ™, "..." to …, and -> --> <-- ==> <== <-> <=> to arrows;
 *  - 1/2, 1/4, 3/4 to ½ ¼ ¾ and English ordinals (1st) to superscript;
 *  - "1. ", "1) ", "a) ", "- ", "* ", "• " at the start of a paragraph start a list;
 *  - URLs, www. addresses and e-mail addresses become links on space, Tab or Enter;
 *  - the first letter of a sentence is capitalized (and "i" in English), and
 *    TWo INitial CApitals are corrected.
 *
 * This module has no runtime imports, so Node tests can load it directly.
 */
import type { EngineBridge, InsertTextContext, InsertTextEdit, InsertTextHook } from "./engine-bridge";

export type AutoFormatRule = "quotes" | "dashes" | "symbols" | "fractions" | "ordinals" | "lists" | "links" | "capitalize";

export interface AutoFormatOptions {
  /** The single "AutoFormat as you type" switch: false turns every rule off. Default true. */
  enabled?: boolean;
  /** BCP 47 tag that picks the quote style and English-only rules; default "en-US". */
  locale?: string;
  smartQuotes?: boolean;
  dashes?: boolean;
  symbols?: boolean;
  fractions?: boolean;
  ordinals?: boolean;
  lists?: boolean;
  links?: boolean;
  capitalizeSentences?: boolean;
  /** THe -> The. Default true. */
  twoInitialCaps?: boolean;
  /** Called once per applied change (for an "AutoFormat options" affordance or telemetry-free UI hints). */
  onChange?: (change: AutoFormatChange) => void;
}

/** A list edit can carry the numbering the user typed ("a)" is lower-letter). */
export type AutoFormatEdit =
  | Extract<InsertTextEdit, { type: "replace" | "format" }>
  | (Extract<InsertTextEdit, { type: "list" }> & { numberFormat?: "decimal" | "lowerLetter" | "upperLetter"; marker?: string });

export interface AutoFormatChange {
  rule: AutoFormatRule;
  /** Paragraph the change applies to. */
  blockId: string;
  edits: AutoFormatEdit[];
  /** What the user typed in the changed range (restored by Ctrl+Z). */
  typed: string;
  /** What replaced it ("" for a list marker that became a list). */
  result: string;
}

export interface QuoteStyle {
  open: string;
  close: string;
  openSingle: string;
  closeSingle: string;
  apostrophe: string;
}

// ---------------------------------------------------------------------------------------
// Quotes

const QUOTES: Record<string, QuoteStyle> = {
  en: { open: "“", close: "”", openSingle: "‘", closeSingle: "’", apostrophe: "’" },
  de: { open: "„", close: "“", openSingle: "‚", closeSingle: "‘", apostrophe: "’" },
  ru: { open: "«", close: "»", openSingle: "„", closeSingle: "“", apostrophe: "’" },
  fr: { open: "«", close: "»", openSingle: "“", closeSingle: "”", apostrophe: "’" },
  pl: { open: "„", close: "”", openSingle: "«", closeSingle: "»", apostrophe: "’" },
  hu: { open: "„", close: "”", openSingle: "»", closeSingle: "«", apostrophe: "’" },
  sv: { open: "”", close: "”", openSingle: "’", closeSingle: "’", apostrophe: "’" },
  ja: { open: "「", close: "」", openSingle: "『", closeSingle: "』", apostrophe: "’" },
  uz: { open: "“", close: "”", openSingle: "‘", closeSingle: "’", apostrophe: "ʼ" },
};
const QUOTE_FAMILY: Record<string, keyof typeof QUOTES> = {
  de: "de", cs: "de", sk: "de", sl: "de", lt: "de", et: "de", is: "de", bg: "de", ka: "de", lb: "de",
  ru: "ru", uk: "ru", be: "ru", kk: "ru", ky: "ru", tg: "ru", az: "ru",
  fr: "fr", es: "fr", it: "fr", ca: "fr", el: "fr", nb: "fr", no: "fr", nn: "fr", pt: "fr", hy: "fr", fa: "fr", ar: "fr",
  pl: "pl", ro: "pl", hr: "pl", nl: "en",
  hu: "hu", sv: "sv", fi: "sv",
  ja: "ja",
  uz: "uz",
};

/** Quote characters for a locale (CLDR conventions; English for unknown locales). */
export function quoteStyle(locale = "en-US"): QuoteStyle {
  const parts = String(locale || "en").toLowerCase().split(/[-_]/);
  const base = parts[0];
  if (base === "uz" && parts.includes("cyrl")) return { ...QUOTES.ru, apostrophe: "ʼ" };
  if (base === "pt" && parts.includes("br")) return QUOTES.en;
  if ((base === "zh" && (parts.includes("tw") || parts.includes("hk") || parts.includes("hant")))) return QUOTES.ja;
  return QUOTES[QUOTE_FAMILY[base] ?? "en"];
}

const OPENING_CONTEXT = /[\s([{<—–\-/«„“‘‚「『\u0000]/u;

function isOpeningPosition(previous: string | undefined, openers: string): boolean {
  return previous === undefined || previous === "" || OPENING_CONTEXT.test(previous) || openers.includes(previous);
}

// ---------------------------------------------------------------------------------------
// Helpers

const LETTER = /\p{L}/u;
const LETTER_OR_DIGIT = /[\p{L}\p{N}]/u;
const WHITESPACE = /^[ \t\v ]$/;
const TRAILING_PUNCTUATION = /[.,;:!?'"’”»)\]}>]+$/u;
const LEADING_PUNCTUATION = /^[(\[{<"'“‘«„‚]+/u;
const URL = /^(?:(?:https?|ftp):\/\/[^\s/$.?#][^\s]*|www\.[\p{L}\p{N}-]+(?:\.[\p{L}\p{N}-]+)+(?:[/?#][^\s]*)?)$/iu;
const EMAIL = /^(?:mailto:)?[\p{L}\p{N}._%+-]+@[\p{L}\p{N}-]+(?:\.[\p{L}\p{N}-]+)*\.\p{L}{2,}$/iu;
const ABBREVIATIONS = new Set([
  "e.g", "i.e", "etc", "vs", "cf", "approx", "appt", "apt", "dept", "est", "fig", "figs", "incl", "no", "nos", "p", "pp", "vol", "vols",
  "mr", "mrs", "ms", "dr", "prof", "sr", "jr", "st", "mt", "ft", "inc", "ltd", "co", "corp", "jan", "feb", "mar", "apr", "jun", "jul",
  "aug", "sep", "sept", "oct", "nov", "dec", "a.m", "p.m", "ca", "al", "ed", "eds", "ch", "sec", "min", "max", "ref", "tel", "viz",
]);
const CODE_STYLE = /code|preformatted|source|verbatim|macro/i;

/** The whitespace-delimited token that ends at `end` in `text`. */
function tokenBefore(text: string, end: number): { start: number; token: string } {
  let start = end;
  while (start > 0 && !/\s/.test(text[start - 1])) start--;
  return { start, token: text.slice(start, end) };
}

function upperFirst(word: string, locale: string): string | null {
  const first = String.fromCodePoint(word.codePointAt(0)!);
  let upper: string;
  try {
    upper = first.toLocaleUpperCase(locale);
  } catch {
    upper = first.toUpperCase();
  }
  if (upper === first || [...upper].length !== 1) return null;
  return upper;
}

/** Applies replace edits to a paragraph's text (right to left, like the engine). Test and preview helper. */
export function applyTextEdits(text: string, edits: readonly InsertTextEdit[]): string {
  const replaces = edits.filter((edit): edit is Extract<InsertTextEdit, { type: "replace" }> => edit.type === "replace").sort((a, b) => b.start - a.start);
  let result = text;
  for (const edit of replaces) {
    const start = Math.max(0, Math.min(result.length, edit.start));
    const end = Math.max(start, Math.min(result.length, edit.end));
    result = result.slice(0, start) + edit.text + result.slice(end);
  }
  return result;
}

/**
 * Edits that undo replace edits on the text they were applied to (`before` is the text
 * the edits were computed for). Format and list edits are left to the engine's undo.
 */
export function invertTextEdits(before: string, edits: readonly InsertTextEdit[]): InsertTextEdit[] {
  const replaces = edits.filter((edit): edit is Extract<InsertTextEdit, { type: "replace" }> => edit.type === "replace").sort((a, b) => a.start - b.start);
  const inverse: InsertTextEdit[] = [];
  let shift = 0;
  for (const edit of replaces) {
    const start = edit.start + shift;
    inverse.push({ type: "replace", ...(edit.blockId ? { blockId: edit.blockId } : {}), start, end: start + edit.text.length, text: before.slice(edit.start, edit.end) });
    shift += edit.text.length - (edit.end - edit.start);
  }
  return inverse;
}

// ---------------------------------------------------------------------------------------
// Rules

interface Resolved {
  locale: string;
  english: boolean;
  quotes: QuoteStyle;
  smartQuotes: boolean;
  dashes: boolean;
  symbols: boolean;
  fractions: boolean;
  ordinals: boolean;
  lists: boolean;
  links: boolean;
  capitalizeSentences: boolean;
  twoInitialCaps: boolean;
}

function resolve(options: AutoFormatOptions): Resolved {
  const locale = options.locale || "en-US";
  return {
    locale,
    english: /^en\b/i.test(locale),
    quotes: quoteStyle(locale),
    smartQuotes: options.smartQuotes ?? true,
    dashes: options.dashes ?? true,
    symbols: options.symbols ?? true,
    fractions: options.fractions ?? true,
    ordinals: options.ordinals ?? true,
    lists: options.lists ?? true,
    links: options.links ?? true,
    capitalizeSentences: options.capitalizeSentences ?? true,
    twoInitialCaps: options.twoInitialCaps ?? true,
  };
}

function change(rule: AutoFormatRule, blockId: string, edits: AutoFormatEdit[], typed: string, result: string): AutoFormatChange {
  return { rule, blockId, edits, typed, result };
}

function replace(blockId: string, start: number, end: number, text: string, before: string, rule: AutoFormatRule): AutoFormatChange {
  return change(rule, blockId, [{ type: "replace", blockId, start, end, text }], before.slice(start, end), text);
}

function quoteRule(text: string, start: number, typed: string, blockId: string, options: Resolved): AutoFormatChange | null {
  if (!options.smartQuotes || (typed !== '"' && typed !== "'")) return null;
  const previous = start > 0 ? text[start - 1] : undefined;
  const quotes = options.quotes;
  let next: string;
  if (typed === '"') {
    next = isOpeningPosition(previous, quotes.open + quotes.openSingle) ? quotes.open : quotes.close;
  } else if (previous !== undefined && LETTER_OR_DIGIT.test(previous)) {
    // Inside or right after a word: an apostrophe (Uzbek oʻ and gʻ use the turned comma).
    // Where the closing single quote differs from the apostrophe (German ‚…‘), an open
    // inner quote is more likely being closed.
    const uzbekLatin = /^uz\b/i.test(options.locale) && !/cyrl/i.test(options.locale);
    const before = text.slice(0, start);
    const unclosed = quotes.closeSingle !== quotes.apostrophe && before.split(quotes.openSingle).length > before.split(quotes.closeSingle).length;
    next = uzbekLatin && /[oOgG]/.test(previous) ? "ʻ" : unclosed ? quotes.closeSingle : quotes.apostrophe;
  } else {
    next = isOpeningPosition(previous, quotes.open + quotes.openSingle) ? quotes.openSingle : quotes.closeSingle;
  }
  return next === typed ? null : replace(blockId, start, start + 1, next, text, "quotes");
}

const ARROWS: ReadonlyArray<[string, string]> = [
  ["<->", "↔"], ["<=>", "⇔"], ["-->", "→"], ["==>", "⇒"], ["->", "→"],
];
const LEFT_ARROWS: ReadonlyArray<[string, string]> = [["<--", "←"], ["<==", "⇐"]];
const SYMBOLS: ReadonlyArray<[RegExp, string]> = [[/\((?:c|C)\)$/, "©"], [/\((?:r|R)\)$/, "®"], [/\((?:tm|TM|Tm)\)$/, "™"]];

function symbolRule(text: string, offset: number, typed: string, blockId: string, options: Resolved): AutoFormatChange | null {
  if (!options.symbols) return null;
  const before = text.slice(0, offset);
  if (typed === ")") {
    for (const [pattern, symbol] of SYMBOLS) {
      const match = pattern.exec(before);
      if (match) return replace(blockId, offset - match[0].length, offset, symbol, text, "symbols");
    }
  }
  if (typed === ">") {
    for (const [arrow, symbol] of ARROWS) if (before.endsWith(arrow) && before[offset - arrow.length - 1] !== "-" && before[offset - arrow.length - 1] !== "=") {
      return replace(blockId, offset - arrow.length, offset, symbol, text, "symbols");
    }
  }
  if (typed === "-" || typed === "=") {
    for (const [arrow, symbol] of LEFT_ARROWS) if (before.endsWith(arrow) && text[offset] !== "-" && text[offset] !== "=" && text[offset] !== ">") {
      return replace(blockId, offset - arrow.length, offset, symbol, text, "symbols");
    }
  }
  if (typed === "." && before.endsWith("...") && before[offset - 4] !== ".") {
    return replace(blockId, offset - 3, offset, "…", text, "symbols");
  }
  return null;
}

function dashRule(text: string, start: number, typed: string, blockId: string, options: Resolved): AutoFormatChange | null {
  if (!options.dashes || typed === "-" || typed === ">" || typed === "=") return null;
  let runStart = start;
  while (runStart > 0 && text[runStart - 1] === "-") runStart--;
  const run = start - runStart;
  if (run === 2 || run === 3) {
    const before = runStart > 0 ? text[runStart - 1] : undefined;
    // Needs text before it; leaves "<--", "<!--" and anything inside a URL alone.
    if (before === undefined || before === "<" || before === "!") return null;
    const { token } = tokenBefore(text, runStart);
    if (/:\/\/|^www\./i.test(token)) return null;
    return replace(blockId, runStart, start, run === 2 ? "–" : "—", text, "dashes");
  }
  // "word - word": a spaced hyphen becomes an en dash when the space after it is typed.
  if (/^[  ]$/.test(typed) && start >= 3 && text[start - 1] === "-" && /[  ]/.test(text[start - 2])) {
    const previous = text[start - 3];
    if (previous && /[\p{L}\p{N})\]"'”’»%]/u.test(previous)) return replace(blockId, start - 1, start, "–", text, "dashes");
  }
  return null;
}

function fractionRule(text: string, start: number, typed: string, blockId: string, options: Resolved): AutoFormatChange | null {
  if (!options.fractions || LETTER_OR_DIGIT.test(typed) || typed === "/") return null;
  const fractions: Record<string, string> = { "1/2": "½", "1/4": "¼", "3/4": "¾" };
  const candidate = text.slice(Math.max(0, start - 3), start);
  const symbol = fractions[candidate];
  if (!symbol) return null;
  const previous = start > 3 ? text[start - 4] : undefined;
  if (previous !== undefined && !/[\s(\[]/.test(previous)) return null;
  return replace(blockId, start - 3, start, symbol, text, "fractions");
}

function ordinalRule(text: string, start: number, typed: string, blockId: string, options: Resolved): AutoFormatChange | null {
  if (!options.ordinals || !options.english || LETTER_OR_DIGIT.test(typed)) return null;
  const { start: tokenStart, token } = tokenBefore(text, start);
  const match = /^(\d+)(st|nd|rd|th)$/i.exec(token);
  if (!match) return null;
  const number = Number(match[1]);
  const lastTwo = number % 100;
  const last = number % 10;
  const expected = lastTwo >= 11 && lastTwo <= 13 ? "th" : last === 1 ? "st" : last === 2 ? "nd" : last === 3 ? "rd" : "th";
  if (match[2].toLowerCase() !== expected) return null;
  const suffixStart = tokenStart + match[1].length;
  return change("ordinals", blockId, [{ type: "format", blockId, start: suffixStart, end: start, style: { verticalAlign: "super" } }], match[2], match[2]);
}

function listRule(context: InsertTextContext, start: number, typed: string, options: Resolved): AutoFormatChange | null {
  if (!options.lists || (typed !== " " && typed !== "\t") || context.paragraphStyle?.list) return null;
  const marker = context.paragraphText.slice(0, start);
  const blockId = context.blockId;
  const edits = (kind: "bullet" | "number", numberFormat?: "decimal" | "lowerLetter" | "upperLetter"): AutoFormatEdit[] => [
    { type: "replace", blockId, start: 0, end: context.offset, text: "" },
    numberFormat ? { type: "list", blockId, kind, numberFormat, marker } : { type: "list", blockId, kind, marker },
  ];
  const typedText = marker + typed;
  if (/^[-*•]$/.test(marker)) return change("lists", blockId, edits("bullet"), typedText, "");
  if (/^1[.)]$/.test(marker)) return change("lists", blockId, edits("number", "decimal"), typedText, "");
  if (/^a[.)]$/.test(marker)) return change("lists", blockId, edits("number", "lowerLetter"), typedText, "");
  if (/^A[.)]$/.test(marker)) return change("lists", blockId, edits("number", "upperLetter"), typedText, "");
  return null;
}

/** The link target for a typed token, or null when it is not an address. */
export function linkTarget(token: string): string | null {
  if (URL.test(token)) return /^www\./i.test(token) ? `https://${token}` : token;
  if (EMAIL.test(token)) return /^mailto:/i.test(token) ? token : `mailto:${token}`;
  return null;
}

function trimAddress(text: string, start: number, end: number): { start: number; end: number } {
  let token = text.slice(start, end);
  const leading = LEADING_PUNCTUATION.exec(token);
  if (leading) {
    start += leading[0].length;
    token = token.slice(leading[0].length);
  }
  const trailing = TRAILING_PUNCTUATION.exec(token);
  if (trailing) {
    let cut = trailing[0];
    // Keep a closing parenthesis that belongs to the address (Wikipedia-style links).
    if (cut.startsWith(")") && token.slice(0, -cut.length).includes("(")) cut = cut.slice(1);
    end -= cut.length;
  }
  return { start, end };
}

function linkRule(text: string, end: number, blockId: string, options: Resolved): AutoFormatChange | null {
  if (!options.links) return null;
  const token = tokenBefore(text, end);
  const range = trimAddress(text, token.start, end);
  if (range.end <= range.start) return null;
  const address = text.slice(range.start, range.end);
  const href = linkTarget(address);
  if (!href) return null;
  return change("links", blockId, [{ type: "format", blockId, start: range.start, end: range.end, style: { link: href } }], address, address);
}

/** True when the word ending before `wordStart` closes a sentence (so the next word starts one). */
function startsSentence(text: string, wordStart: number): boolean {
  let index = wordStart;
  while (index > 0 && LEADING_PUNCTUATION.test(text[index - 1])) index--;
  let gapEnd = index;
  while (index > 0 && /\s/.test(text[index - 1])) index--;
  if (index === 0) return true;
  if (index === gapEnd) return false; // no whitespace between the words
  gapEnd = index;
  while (index > 0 && /["'’”»)\]]/.test(text[index - 1])) index--;
  const terminator = text[index - 1];
  if (terminator === "!" || terminator === "?") return true;
  if (terminator !== ".") return false;
  if (text.slice(Math.max(0, index - 3), index) === "...") return false;
  const { token } = tokenBefore(text, index - 1);
  const word = token.replace(LEADING_PUNCTUATION, "").toLowerCase();
  if (!word || ABBREVIATIONS.has(word)) return false;
  // Initials and abbreviations with inner dots (J. Smith, U.S. economy).
  if (word.includes(".") || /^\p{L}$/u.test(word)) return false;
  return true;
}

function capitalizeRule(text: string, end: number, blockId: string, options: Resolved): AutoFormatChange | null {
  const { start: tokenStart, token } = tokenBefore(text, end);
  const lead = LEADING_PUNCTUATION.exec(token)?.[0].length ?? 0;
  const wordStart = tokenStart + lead;
  const word = text.slice(wordStart, end).replace(/[.,;:!?…"’”»)\]]+$/u, "");
  if (!word || !LETTER.test(word[0])) return null;
  if (!/^[\p{L}\p{M}'’ʻʼ-]+$/u.test(word)) return null;
  if (options.twoInitialCaps && /^\p{Lu}\p{Lu}\p{Ll}+$/u.test(word) && !/^\p{Lu}{2,}s$/u.test(word)) {
    const second = word[1];
    return replace(blockId, wordStart + 1, wordStart + 2, second.toLocaleLowerCase(options.locale), text, "capitalize");
  }
  if (!options.capitalizeSentences) return null;
  const pronoun = options.english && /^i(?:['’](?:m|ll|ve|d))?$/.test(word);
  if (!pronoun && /\p{Lu}/u.test(word.slice(1))) return null;
  if (!pronoun && !startsSentence(text, tokenStart)) return null;
  const upper = upperFirst(word, options.locale);
  if (!upper) return null;
  return replace(blockId, wordStart, wordStart + word[0].length, upper, text, "capitalize");
}

function merge(changes: ReadonlyArray<AutoFormatChange | null>): AutoFormatChange[] {
  const accepted: AutoFormatChange[] = [];
  const spans: Array<[string, number, number]> = [];
  for (const candidate of changes) {
    if (!candidate) continue;
    const ranges = candidate.edits.filter((edit) => edit.type !== "list").map((edit) => [edit.blockId ?? candidate.blockId, (edit as { start: number }).start, (edit as { end: number }).end] as [string, number, number]);
    if (ranges.some(([block, start, end]) => spans.some(([other, s, e]) => other === block && start < e && s < end))) continue;
    spans.push(...ranges);
    accepted.push(candidate);
  }
  return accepted;
}

/**
 * The AutoFormat changes for one insert-text hook call, in the order they apply. Each
 * change says which rule fired, what was typed and what replaced it.
 */
export function autoFormatChanges(context: InsertTextContext, options: AutoFormatOptions = {}): AutoFormatChange[] {
  if (options.enabled === false) return [];
  const resolved = resolve(options);
  if (context.mode === "view" || CODE_STYLE.test(context.paragraphStyle?.namedStyle ?? "")) return [];
  if (context.kind === "paragraph") {
    const previousId = context.previousBlockId;
    const previous = context.previousText;
    if (!previousId || typeof previous !== "string" || previous.length === 0) return [];
    if (!/\S$/.test(previous)) return [];
    return merge([
      linkRule(previous, previous.length, previousId, resolved),
      ordinalRule(previous, previous.length, "\n", previousId, resolved),
      capitalizeRule(previous, previous.length, previousId, resolved),
    ]);
  }
  const text = context.paragraphText;
  const typed = context.text;
  const offset = context.offset;
  const start = offset - typed.length;
  // One typed character, and the model really holds it where the engine says.
  if (typed.length !== 1 || start < 0 || text.slice(start, offset) !== typed) return [];
  const blockId = context.blockId;

  const list = listRule(context, start, typed, resolved);
  if (list) return [list];

  const changes: Array<AutoFormatChange | null> = [
    quoteRule(text, start, typed, blockId, resolved),
    symbolRule(text, offset, typed, blockId, resolved),
    dashRule(text, start, typed, blockId, resolved),
  ];
  if (WHITESPACE.test(typed)) {
    changes.push(linkRule(text, start, blockId, resolved));
    changes.push(fractionRule(text, start, typed, blockId, resolved));
    changes.push(ordinalRule(text, start, typed, blockId, resolved));
    changes.push(capitalizeRule(text, start, blockId, resolved));
  } else if (/[,;:!?.)]/.test(typed)) {
    changes.push(fractionRule(text, start, typed, blockId, resolved));
  }
  return merge(changes);
}

/** The edits for one insert-text hook call (all changes together become one undo step). */
export function autoFormat(context: InsertTextContext, options: AutoFormatOptions = {}): AutoFormatEdit[] {
  return autoFormatChanges(context, options).flatMap((item) => item.edits);
}

/** An insert-text hook running the AutoFormat rules. */
export function createAutoFormatHook(options: AutoFormatOptions | (() => AutoFormatOptions) = {}): InsertTextHook {
  return (context) => {
    const current = typeof options === "function" ? options() : options;
    const changes = autoFormatChanges(context, current);
    if (changes.length === 0) return null;
    if (current.onChange) for (const item of changes) {
      try { current.onChange(item); } catch { /* a listener never blocks typing */ }
    }
    return changes.flatMap((item) => item.edits);
  };
}

export interface AutoFormatInstallation {
  /** Replaces the options (for example after the AutoFormat toggle or a locale change). */
  setOptions(options: AutoFormatOptions): void;
  /** Removes the hook and turns the engine's own typing conversions back on. */
  dispose(): void;
  /** False when the engine has no insert-text hook (AutoFormat is then off). */
  readonly active: boolean;
}

/**
 * Installs AutoFormat on the engine bridge and turns the engine's built-in conversions
 * off so the two never both rewrite a keystroke.
 */
export function installAutoFormat(bridge: Pick<EngineBridge, "setInsertTextHook" | "setBuiltinAutoCorrect">, options: AutoFormatOptions = {}): AutoFormatInstallation {
  let current = options;
  const active = bridge.setInsertTextHook(createAutoFormatHook(() => current));
  if (active) bridge.setBuiltinAutoCorrect(false);
  let disposed = false;
  return {
    active,
    setOptions(next) {
      current = next;
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      if (!active) return;
      bridge.setInsertTextHook(null);
      bridge.setBuiltinAutoCorrect(true);
    },
  };
}
