/**
 * Spell checking for Simple Docs, fully local.
 *
 * The dictionary is the Windows Spell Checking API (the same offline checker Windows
 * uses everywhere), reached through the preload `simpleDocs.spell` bridge; nothing is
 * downloaded. This module is pure: it tokenizes paragraph text, batches and caches
 * dictionary lookups, applies the user dictionary and "Ignore all", and turns the
 * results into screen-only decorations for the engine (never printed or exported) plus
 * the data a context menu needs (suggestions, Add to dictionary, Ignore all).
 *
 * Tokenizing follows Word's defaults: Internet and file addresses, words with numbers,
 * words in UPPERCASE and single letters are not checked; hyphenated compounds are
 * checked part by part; apostrophes inside a word are part of it (don't, oʻzbek);
 * words in another script than the checking language are left alone, and a word that
 * mixes Latin, Cyrillic or Greek letters (a typical look-alike typo) is flagged.
 * A repeated word ("the the") is flagged on its second occurrence.
 *
 * This module has no runtime imports, so Node tests can load it directly.
 */
import type { Decoration, DecorationClick } from "./engine-bridge";
import type { Block, DocPosition, Paragraph } from "@forevka/wordcanvas/query";

/** A word (or hyphen-separated part) to look up, with UTF-16 offsets into the paragraph. */
export interface SpellToken {
  word: string;
  start: number;
  end: number;
  /** Set when the word mixes alphabets (Latin + Cyrillic, ...); flagged without a lookup. */
  mixedScript?: boolean;
  /** For a part of a hyphenated compound: the whole compound (a user-dictionary entry accepts all parts). */
  compound?: string;
}

export interface SpellWord {
  word: string;
  start: number;
  end: number;
}

export type SpellIssueKind = "spelling" | "repeated" | "mixed-script";

export interface SpellIssue {
  kind: SpellIssueKind;
  blockId: string;
  start: number;
  end: number;
  word: string;
  /** For a repeated word: the range to delete (the word and the whitespace before it). */
  deleteStart?: number;
  deleteEnd?: number;
}

export interface SpellTokenizeOptions {
  /** BCP 47 tag of the checking language; decides which script is checked. Default "en-US". */
  language?: string;
  /** Skip words in UPPERCASE (NASA, NASA's). Default true. */
  ignoreUppercase?: boolean;
  /** Skip words that contain digits or underscores (A4, 2nd, v8_x). Default true. */
  ignoreWordsWithNumbers?: boolean;
  /** Skip URLs, e-mail addresses, file paths, domains, file names, #tags and @mentions. Default true. */
  ignoreInternetAndFileAddresses?: boolean;
  /** Skip words with an inner capital (iPhone, JavaScript), usually names or code. Default true. */
  ignoreMixedCase?: boolean;
  /** Flag words that mix Latin, Cyrillic or Greek letters. Default true. */
  flagMixedScripts?: boolean;
}

export interface SpellcheckOptions extends SpellTokenizeOptions {
  /** Flag the second word of "the the". Default true. */
  flagRepeatedWords?: boolean;
  /** Words per dictionary request. Default 250. */
  batchSize?: number;
  /** Cached lookups kept (least recently used are dropped). Default 20000. */
  cacheSize?: number;
  /** Underline colour. Default "#d93025". */
  color?: string;
  /** Underline thickness in px. Default 1.5. */
  thickness?: number;
  /** Suggestions offered in the context menu. Default 5. */
  maxSuggestions?: number;
}

/** Result of a dictionary lookup through the preload bridge. */
export interface SpellCheckResult {
  /** "windows": Windows Spell Checking API; "chromium": Electron's own checker; null: no dictionary. */
  source: "windows" | "chromium" | null;
  /** The dictionary language that was used, when known. */
  language?: string | null;
  /** One entry per requested word; true = misspelled. */
  misspelled: boolean[];
}

/** The dictionary behind the checker (the preload bridge in the app, a fake in tests). */
export interface SpellBackend {
  checkWords(words: readonly string[], language?: string): Promise<SpellCheckResult>;
  getSuggestions(word: string, language?: string): Promise<string[]>;
}

export interface UserDictionary {
  /** "Add to dictionary": accepted everywhere, persisted in Simple's user data. */
  words?: readonly string[];
  /** "Ignore all": accepted as written. */
  ignored?: readonly string[];
}

/** The `simpleDocs.spell` object exposed by electron/preload.cjs. */
export interface PreloadSpellApi {
  isWordMisspelled(word: string, language?: string): Promise<boolean>;
  getWordSuggestions(word: string, language?: string): Promise<string[]>;
  checkWords(words: readonly string[], language?: string): Promise<SpellCheckResult>;
  getLanguages(): Promise<{ available: boolean; languages: string[]; preferred: string | null }>;
  getUserDictionary(): Promise<Required<UserDictionary>>;
  addWord(word: string): Promise<Required<UserDictionary>>;
  removeWord(word: string): Promise<Required<UserDictionary>>;
  ignoreWord(word: string): Promise<Required<UserDictionary>>;
  unignoreWord(word: string): Promise<Required<UserDictionary>>;
  onDictionaryChanged(callback: (dictionary: Required<UserDictionary>) => void): () => void;
}

export interface SpellSuggestion {
  /** Menu label, e.g. "receive" or "Delete repeated word". */
  label: string;
  /** Replacement text for [start, end) of `blockId` ("" deletes). */
  replacement: string;
  blockId: string;
  start: number;
  end: number;
}

/** Everything a spelling context menu shows for one issue. */
export interface SpellMenu {
  issue: SpellIssue;
  suggestions: SpellSuggestion[];
  /** Offer "Add to dictionary" (spelling issues only). */
  canAddToDictionary: boolean;
  /** Offer "Ignore all" (spelling and mixed-script issues). */
  canIgnore: boolean;
}

/** A paragraph's id and plain text (UTF-16 offsets match engine positions). */
export interface SpellParagraph {
  id: string;
  text: string;
  /** Skip the whole paragraph (code styles). */
  skip?: boolean;
}

export interface DocumentLike {
  blocks: readonly Block[];
}

// ---------------------------------------------------------------------------------------
// Scripts and languages

type Script = "Latin" | "Cyrillic" | "Greek" | "Arabic" | "Hebrew" | "Armenian" | "Georgian" | "Other";

const SCRIPT_TESTS: ReadonlyArray<[Script, RegExp]> = [
  ["Latin", /\p{Script=Latin}/u],
  ["Cyrillic", /\p{Script=Cyrillic}/u],
  ["Greek", /\p{Script=Greek}/u],
  ["Arabic", /\p{Script=Arabic}/u],
  ["Hebrew", /\p{Script=Hebrew}/u],
  ["Armenian", /\p{Script=Armenian}/u],
  ["Georgian", /\p{Script=Georgian}/u],
];
/** Scripts whose words are separated by spaces and have a spell checker worth asking. */
const CASED_LOOKALIKE: ReadonlySet<Script> = new Set(["Latin", "Cyrillic", "Greek"]);
const CYRILLIC_LANGUAGES = new Set(["ru", "uk", "be", "bg", "sr", "mk", "kk", "ky", "tg", "mn", "tt", "ba", "cv", "sah"]);

function scriptOf(char: string): Script | null {
  if (!/\p{L}/u.test(char)) return null;
  for (const [script, test] of SCRIPT_TESTS) if (test.test(char)) return script;
  // Modifier letters such as ʻ (U+02BB) and ʼ (U+02BC) are Script=Common: no script.
  return /\p{Script=Common}|\p{Script=Inherited}/u.test(char) ? null : "Other";
}

/** The alphabet a language is written in (uz-Cyrl and sr-Cyrl are Cyrillic, sr-Latn Latin). */
export function scriptForLanguage(language: string | undefined | null): Script {
  const parts = String(language || "en").toLowerCase().split(/[-_]/);
  const base = parts[0];
  if (parts.includes("cyrl")) return "Cyrillic";
  if (parts.includes("latn")) return "Latin";
  if (CYRILLIC_LANGUAGES.has(base)) return "Cyrillic";
  if (base === "el") return "Greek";
  if (base === "ar" || base === "fa" || base === "ur" || base === "ps") return "Arabic";
  if (base === "he" || base === "yi") return "Hebrew";
  if (base === "hy") return "Armenian";
  if (base === "ka") return "Georgian";
  return "Latin";
}

// ---------------------------------------------------------------------------------------
// Tokenizing

const WORD = /[\p{L}\p{M}\p{N}_]+(?:['’‑-][\p{L}\p{M}\p{N}_]+)*/gu;
const HYPHENS = /[‑-]/;

const ADDRESS_PATTERNS: readonly RegExp[] = [
  // Scheme URLs and www. addresses.
  /\b(?:[a-z][a-z0-9+.-]{1,15}:\/\/|www\.|mailto:)[^\s<>"“”«»]+/giu,
  // E-mail addresses.
  /[\p{L}\p{N}._%+-]+@[\p{L}\p{N}-]+(?:\.[\p{L}\p{N}-]+)+/gu,
  // Windows drive paths and UNC paths.
  /\b[a-z]:\\[^\s<>"|?*]*/giu,
  /\\\\[^\s\\<>"|?*]+\\[^\s<>"|?*]*/gu,
  // Bare domains and file names: name.ext, sub.example.org (last part lower-case, 2-6 letters).
  /(?<![\p{L}\p{N}_@.])[\p{L}\p{N}_-]+(?:\.[\p{L}\p{N}_-]+)*\.[a-z][a-z0-9]{1,5}(?![\p{L}\p{N}_])(?:\/[^\s<>"]*)?/gu,
  // #tags and @mentions.
  /(?<![\p{L}\p{N}_])[#@][\p{L}\p{N}_]+/gu,
];

/** Ranges that are Internet or file addresses (never spell-checked). */
export function addressRanges(text: string): Array<{ start: number; end: number }> {
  const ranges: Array<{ start: number; end: number }> = [];
  for (const pattern of ADDRESS_PATTERNS) {
    pattern.lastIndex = 0;
    for (const match of text.matchAll(pattern)) {
      let end = match.index + match[0].length;
      // Sentence punctuation after an address is not part of it.
      while (end > match.index && /[.,;:!?)\]}'"’”»]/.test(text[end - 1])) end--;
      if (end > match.index) ranges.push({ start: match.index, end });
    }
  }
  return ranges.sort((a, b) => a.start - b.start);
}

function overlaps(ranges: ReadonlyArray<{ start: number; end: number }>, start: number, end: number): boolean {
  for (const range of ranges) {
    if (range.start >= end) break;
    if (range.end > start) return true;
  }
  return false;
}

function letters(word: string): string {
  return word.replace(/[^\p{L}]/gu, "");
}

function isUppercaseWord(word: string): boolean {
  const core = word.replace(/['’]s$/u, "");
  const only = letters(core);
  return only.length >= 2 && only === only.toUpperCase() && only !== only.toLowerCase();
}

/** Plain words of a text with offsets (letters, digits, inner apostrophes and hyphens). */
export function wordsOf(text: string): SpellWord[] {
  const words: SpellWord[] = [];
  WORD.lastIndex = 0;
  for (const match of text.matchAll(WORD)) words.push({ word: match[0], start: match.index, end: match.index + match[0].length });
  return words;
}

function classifyPart(part: string, options: Required<SpellTokenizeOptions>, languageScript: Script): "check" | "skip" | "mixed" {
  if (options.ignoreWordsWithNumbers && /[\p{N}_]/u.test(part)) return "skip";
  const onlyLetters = letters(part);
  if (onlyLetters.length < 2) return "skip";
  if (options.ignoreUppercase && isUppercaseWord(part)) return "skip";
  if (options.ignoreMixedCase && /\p{Ll}\p{Lu}/u.test(part)) return "skip";
  const scripts = new Set<Script>();
  for (const char of onlyLetters) {
    const script = scriptOf(char);
    if (script) scripts.add(script);
  }
  if (scripts.has("Other")) return "skip";
  if (scripts.size > 1) {
    const lookalikes = [...scripts].every((script) => CASED_LOOKALIKE.has(script));
    return lookalikes && options.flagMixedScripts && scripts.has(languageScript) ? "mixed" : "skip";
  }
  if (scripts.size === 1 && !scripts.has(languageScript)) return "skip";
  return "check";
}

function tokenizeOptions(options: SpellTokenizeOptions = {}): Required<SpellTokenizeOptions> {
  return {
    language: options.language || "en-US",
    ignoreUppercase: options.ignoreUppercase ?? true,
    ignoreWordsWithNumbers: options.ignoreWordsWithNumbers ?? true,
    ignoreInternetAndFileAddresses: options.ignoreInternetAndFileAddresses ?? true,
    ignoreMixedCase: options.ignoreMixedCase ?? true,
    flagMixedScripts: options.flagMixedScripts ?? true,
  };
}

/**
 * The words of one paragraph that need a dictionary lookup, with UTF-16 offsets.
 * Hyphenated compounds yield one token per part ("well-knwon" -> "well", "knwon").
 */
export function tokenizeSpelling(text: string, options: SpellTokenizeOptions = {}): SpellToken[] {
  const resolved = tokenizeOptions(options);
  const languageScript = scriptForLanguage(resolved.language);
  const skip = resolved.ignoreInternetAndFileAddresses ? addressRanges(text) : [];
  const tokens: SpellToken[] = [];
  for (const { word, start, end } of wordsOf(text)) {
    if (overlaps(skip, start, end)) continue;
    let partStart = start;
    const parts = word.split(HYPHENS);
    for (const part of parts) {
      const kind = classifyPart(part, resolved, languageScript);
      if (kind !== "skip") {
        const token: SpellToken = { word: part, start: partStart, end: partStart + part.length };
        if (kind === "mixed") token.mixedScript = true;
        if (parts.length > 1) token.compound = word;
        tokens.push(token);
      }
      partStart += part.length + 1;
    }
  }
  return tokens;
}

/** Second words of "the the" pairs (separated by spaces or tabs only, same letters ignoring case). */
export function repeatedWords(text: string, options: SpellTokenizeOptions = {}): SpellIssue[] {
  const resolved = tokenizeOptions(options);
  const skip = resolved.ignoreInternetAndFileAddresses ? addressRanges(text) : [];
  const words = wordsOf(text);
  const issues: SpellIssue[] = [];
  for (let index = 1; index < words.length; index++) {
    const previous = words[index - 1];
    const current = words[index];
    if (!/^[ \t ]+$/.test(text.slice(previous.end, current.start))) continue;
    if (/[\p{N}_]/u.test(current.word) || letters(current.word).length === 0) continue;
    if (previous.word.toLocaleLowerCase() !== current.word.toLocaleLowerCase()) continue;
    if (overlaps(skip, previous.start, current.end)) continue;
    issues.push({ kind: "repeated", blockId: "", start: current.start, end: current.end, word: current.word, deleteStart: previous.end, deleteEnd: current.end });
  }
  return issues;
}

/** Key used for dictionary lookups: typographic apostrophes are looked up as ASCII ones. */
export function lookupKey(word: string): string {
  return word.normalize("NFC").replace(/’/g, "'");
}

// ---------------------------------------------------------------------------------------
// Document text

const CODE_STYLE = /code|preformatted|source|verbatim|macro/i;

/**
 * Plain text of a paragraph with offsets that match engine positions. Hidden text and
 * footnote markers are replaced by separators of the same length, so their letters
 * never join the words around them.
 */
export function paragraphSpellText(paragraph: Paragraph): string {
  let text = "";
  for (const run of paragraph.runs) {
    const value = run.text ?? "";
    text += run.style?.hidden || run.style?.footnoteRef ? "\u0000".repeat(value.length) : value;
  }
  return text;
}

/** Every paragraph of the body, including paragraphs inside (nested) table cells. */
export function collectSpellParagraphs(document: DocumentLike): SpellParagraph[] {
  const paragraphs: SpellParagraph[] = [];
  const visit = (blocks: readonly Block[]) => {
    for (const block of blocks) {
      if (block.kind === "paragraph") {
        const skip = CODE_STYLE.test(block.style?.namedStyle ?? "");
        paragraphs.push(skip ? { id: block.id, text: paragraphSpellText(block), skip } : { id: block.id, text: paragraphSpellText(block) });
      } else if (block.kind === "table") {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  visit(document.blocks ?? []);
  return paragraphs;
}

// ---------------------------------------------------------------------------------------
// Decorations

/** Engine decorations (screen-only underlines) for spelling issues. */
export function spellDecorations(
  issues: readonly SpellIssue[],
  options: { color?: string; thickness?: number; onClick?: (issue: SpellIssue, event: { clientX: number; clientY: number }) => void } = {},
): Decoration[] {
  const color = options.color ?? "#d93025";
  const thickness = options.thickness ?? 1.5;
  return issues.map((issue) => {
    const decoration: Decoration = {
      type: "underline",
      range: { anchor: { blockId: issue.blockId, offset: issue.start }, focus: { blockId: issue.blockId, offset: issue.end } },
      color,
      thickness,
    };
    if (options.onClick) {
      const click = options.onClick;
      const handler: DecorationClick = (event) => click(issue, event);
      decoration.onClick = handler;
    }
    return decoration;
  });
}

/** Matches the capitalization of `word` (Word -> Suggestion). */
export function matchCase(word: string, suggestion: string): string {
  if (!suggestion) return suggestion;
  const first = word.charAt(0);
  if (isUppercaseWord(word) && letters(word).length > 1) return suggestion.toUpperCase();
  if (first && first === first.toUpperCase() && first !== first.toLowerCase()) {
    const head = suggestion.charAt(0);
    if (head === head.toLowerCase()) return head.toUpperCase() + suggestion.slice(1);
  }
  return suggestion;
}

// ---------------------------------------------------------------------------------------
// The checker

interface ParagraphCacheEntry {
  text: string;
  tokens: SpellToken[];
  repeated: SpellIssue[];
}

function sameDictionaryCase(word: string, entry: string): boolean {
  if (word === entry) return true;
  // A lower-case entry also accepts Capitalized and UPPERCASE forms (like Word).
  if (entry === entry.toLowerCase()) {
    const lower = word.toLowerCase();
    if (lower !== entry) return false;
    return word === word.toUpperCase() || word === entry.charAt(0).toUpperCase() + entry.slice(1);
  }
  // A Capitalized entry also accepts UPPERCASE.
  return word === entry.toUpperCase();
}

/**
 * Stateful checker for one editor: caches lookups per language, re-tokenizes only
 * paragraphs whose text changed, and batches unknown words into few dictionary calls.
 */
export class SpellChecker {
  readonly backend: SpellBackend;
  private options: Required<SpellcheckOptions>;
  private cache = new Map<string, boolean>();
  private inflight = new Map<string, Promise<boolean>>();
  private paragraphs = new Map<string, ParagraphCacheEntry>();
  private suggestionCache = new Map<string, string[]>();
  private userWords = new Map<string, string[]>();
  private ignored = new Set<string>();
  private sessionIgnored = new Set<string>();
  private lastIssues: SpellIssue[] = [];
  private unavailable = false;

  constructor(backend: SpellBackend, options: SpellcheckOptions = {}) {
    this.backend = backend;
    this.options = SpellChecker.resolve(options);
  }

  private static resolve(options: SpellcheckOptions): Required<SpellcheckOptions> {
    return {
      ...tokenizeOptions(options),
      flagRepeatedWords: options.flagRepeatedWords ?? true,
      batchSize: Math.max(1, Math.floor(options.batchSize ?? 250)),
      cacheSize: Math.max(100, Math.floor(options.cacheSize ?? 20000)),
      color: options.color ?? "#d93025",
      thickness: options.thickness ?? 1.5,
      maxSuggestions: Math.max(1, Math.floor(options.maxSuggestions ?? 5)),
    };
  }

  get language(): string {
    return this.options.language;
  }

  /** False after the dictionary reported it has no checker (then nothing is flagged). */
  get available(): boolean {
    return !this.unavailable;
  }

  /** The issues of the last completed check. */
  get issues(): readonly SpellIssue[] {
    return this.lastIssues;
  }

  setOptions(options: SpellcheckOptions): void {
    const next = SpellChecker.resolve({ ...this.options, ...options });
    const retokenize = (["language", "ignoreUppercase", "ignoreWordsWithNumbers", "ignoreInternetAndFileAddresses", "ignoreMixedCase", "flagMixedScripts"] as const)
      .some((key) => next[key] !== this.options[key]);
    if (next.language !== this.options.language) {
      this.cache.clear();
      this.suggestionCache.clear();
      this.unavailable = false;
    }
    if (retokenize) this.paragraphs.clear();
    this.options = next;
  }

  /** Replaces the user dictionary (from `getUserDictionary()` or a change notification). */
  setUserDictionary(dictionary: UserDictionary): void {
    this.userWords.clear();
    for (const entry of dictionary.words ?? []) this.addUserWord(entry);
    this.ignored = new Set((dictionary.ignored ?? []).map((word) => word.normalize("NFC")));
  }

  private addUserWord(entry: string): void {
    const word = entry.normalize("NFC");
    const key = word.toLowerCase();
    const list = this.userWords.get(key) ?? [];
    if (!list.includes(word)) list.push(word);
    this.userWords.set(key, list);
  }

  /** Accepts a word locally at once (the caller persists it through the bridge). */
  addToDictionary(word: string): void {
    this.addUserWord(word);
  }

  /** "Ignore all": accepts this exact word; `persist: false` keeps it for this session only. */
  ignoreAll(word: string, options: { persist?: boolean } = {}): void {
    (options.persist === false ? this.sessionIgnored : this.ignored).add(word.normalize("NFC"));
  }

  /** True when the user dictionary or an Ignore all accepts the word. */
  isAccepted(word: string): boolean {
    const normalized = word.normalize("NFC");
    if (this.ignored.has(normalized) || this.sessionIgnored.has(normalized)) return true;
    const variants = [normalized, normalized.replace(/’/g, "'"), normalized.replace(/'/g, "’")];
    for (const variant of variants) {
      for (const entry of this.userWords.get(variant.toLowerCase()) ?? []) if (sameDictionaryCase(variant, entry)) return true;
    }
    return false;
  }

  private acceptsToken(token: SpellToken): boolean {
    return this.isAccepted(token.word) || (token.compound !== undefined && this.isAccepted(token.compound));
  }

  private cacheKey(word: string): string {
    return `${this.options.language}\u0000${lookupKey(word)}`;
  }

  private remember(key: string, misspelled: boolean): void {
    if (this.cache.has(key)) this.cache.delete(key);
    this.cache.set(key, misspelled);
    if (this.cache.size > this.options.cacheSize) {
      const oldest = this.cache.keys().next().value;
      if (oldest !== undefined) this.cache.delete(oldest);
    }
  }

  private cached(key: string): boolean | undefined {
    const value = this.cache.get(key);
    if (value !== undefined) {
      this.cache.delete(key);
      this.cache.set(key, value);
    }
    return value;
  }

  /** Looks up words (deduplicated, cached, batched). Resolves true per misspelled word. */
  async lookup(words: readonly string[]): Promise<boolean[]> {
    const keys = words.map((word) => this.cacheKey(word));
    const missing = new Map<string, string>();
    for (let index = 0; index < words.length; index++) {
      const key = keys[index];
      if (this.cache.has(key) || this.inflight.has(key) || missing.has(key)) continue;
      missing.set(key, lookupKey(words[index]));
    }
    const pending = [...missing.entries()];
    for (let offset = 0; offset < pending.length; offset += this.options.batchSize) {
      const batch = pending.slice(offset, offset + this.options.batchSize);
      const request = this.unavailable
        ? Promise.resolve<SpellCheckResult>({ source: null, misspelled: [] })
        : this.backend.checkWords(batch.map(([, word]) => word), this.options.language).catch((): SpellCheckResult => ({ source: null, misspelled: [] }));
      batch.forEach(([key], index) => {
        this.inflight.set(key, request.then((result) => {
          this.inflight.delete(key);
          if (result.source === null) {
            this.unavailable = true;
            return false;
          }
          const misspelled = result.misspelled[index] === true;
          this.remember(key, misspelled);
          return misspelled;
        }));
      });
    }
    return Promise.all(keys.map((key) => {
      const value = this.cached(key);
      return value !== undefined ? value : this.inflight.get(key) ?? false;
    }));
  }

  private paragraphEntry(paragraph: SpellParagraph): ParagraphCacheEntry {
    const known = this.paragraphs.get(paragraph.id);
    if (known && known.text === paragraph.text) return known;
    const entry: ParagraphCacheEntry = {
      text: paragraph.text,
      tokens: tokenizeSpelling(paragraph.text, this.options),
      repeated: repeatedWords(paragraph.text, this.options),
    };
    this.paragraphs.set(paragraph.id, entry);
    return entry;
  }

  /**
   * Checks paragraphs and returns their issues in document order. Pass `caret` to leave
   * the word being typed alone, as Word does until the caret leaves it.
   */
  async checkParagraphs(paragraphs: readonly SpellParagraph[], options: { caret?: DocPosition | null } = {}): Promise<SpellIssue[]> {
    const live = new Set<string>();
    const words: string[] = [];
    const entries: Array<[SpellParagraph, ParagraphCacheEntry]> = [];
    for (const paragraph of paragraphs) {
      live.add(paragraph.id);
      if (paragraph.skip) continue;
      const entry = this.paragraphEntry(paragraph);
      entries.push([paragraph, entry]);
      for (const token of entry.tokens) if (!token.mixedScript && !this.acceptsToken(token)) words.push(token.word);
    }
    for (const id of [...this.paragraphs.keys()]) if (!live.has(id)) this.paragraphs.delete(id);
    const results = await this.lookup(words);
    const verdict = new Map<string, boolean>();
    words.forEach((word, index) => verdict.set(word, results[index]));
    const caret = options.caret ?? null;
    const issues: SpellIssue[] = [];
    for (const [paragraph, entry] of entries) {
      const found: SpellIssue[] = [];
      for (const token of entry.tokens) {
        if (this.acceptsToken(token)) continue;
        if (token.mixedScript) found.push({ kind: "mixed-script", blockId: paragraph.id, start: token.start, end: token.end, word: token.word });
        else if (verdict.get(token.word)) found.push({ kind: "spelling", blockId: paragraph.id, start: token.start, end: token.end, word: token.word });
      }
      if (this.options.flagRepeatedWords) {
        for (const repeated of entry.repeated) {
          if (!found.some((issue) => issue.start < repeated.end && repeated.start < issue.end)) found.push({ ...repeated, blockId: paragraph.id });
        }
      }
      found.sort((a, b) => a.start - b.start);
      for (const issue of found) {
        if (caret && caret.blockId === issue.blockId && caret.offset >= issue.start && caret.offset <= issue.end) continue;
        issues.push(issue);
      }
    }
    this.lastIssues = issues;
    return issues;
  }

  /** Checks every body and table paragraph of a document. */
  checkDocument(document: DocumentLike, options: { caret?: DocPosition | null } = {}): Promise<SpellIssue[]> {
    return this.checkParagraphs(collectSpellParagraphs(document), options);
  }

  /** Decorations for the last check (or for the given issues). */
  decorations(issues: readonly SpellIssue[] = this.lastIssues, onClick?: (issue: SpellIssue, event: { clientX: number; clientY: number }) => void): Decoration[] {
    return spellDecorations(issues, { color: this.options.color, thickness: this.options.thickness, onClick });
  }

  /** The issue under a document position (for a right-click), from the last check. */
  issueAt(position: DocPosition | null | undefined): SpellIssue | null {
    if (!position) return null;
    return this.lastIssues.find((issue) => issue.blockId === position.blockId && position.offset >= issue.start && position.offset <= issue.end) ?? null;
  }

  /** Context-menu entries for an issue: suggestions (case matched), Add to dictionary, Ignore all. */
  async menuFor(issue: SpellIssue): Promise<SpellMenu> {
    if (issue.kind === "repeated") {
      const start = issue.deleteStart ?? issue.start;
      const end = issue.deleteEnd ?? issue.end;
      return {
        issue,
        suggestions: [{ label: "Delete repeated word", replacement: "", blockId: issue.blockId, start, end }],
        canAddToDictionary: false,
        canIgnore: false,
      };
    }
    const words = await this.suggestionsFor(issue.word);
    return {
      issue,
      suggestions: words.map((word) => ({ label: word, replacement: word, blockId: issue.blockId, start: issue.start, end: issue.end })),
      canAddToDictionary: issue.kind === "spelling",
      canIgnore: true,
    };
  }

  /** Suggestions for a word (cached, deduplicated, case matched, at most `maxSuggestions`). */
  async suggestionsFor(word: string): Promise<string[]> {
    const key = this.cacheKey(word);
    let raw = this.suggestionCache.get(key);
    if (!raw) {
      raw = this.unavailable ? [] : await this.backend.getSuggestions(lookupKey(word), this.options.language).catch(() => []);
      this.suggestionCache.set(key, raw);
      if (this.suggestionCache.size > 500) this.suggestionCache.delete(this.suggestionCache.keys().next().value!);
    }
    const seen = new Set<string>();
    const result: string[] = [];
    for (const suggestion of raw) {
      if (typeof suggestion !== "string") continue;
      let cased = matchCase(word, suggestion.trim());
      if (word.includes("’")) cased = cased.replace(/'/g, "’");
      if (!cased || cased === word || seen.has(cased)) continue;
      seen.add(cased);
      result.push(cased);
      if (result.length >= this.options.maxSuggestions) break;
    }
    return result;
  }

  /** Forgets cached lookups (for example after the Windows dictionary changed). */
  clearCache(): void {
    this.cache.clear();
    this.suggestionCache.clear();
    this.unavailable = false;
  }
}

/** Adapts the preload `simpleDocs.spell` object; null when it is missing or incomplete. */
export function preloadSpellBackend(api: unknown): SpellBackend | null {
  const spell = api as Partial<PreloadSpellApi> | null | undefined;
  if (!spell || typeof spell.checkWords !== "function" || typeof spell.getWordSuggestions !== "function") return null;
  return {
    checkWords: async (words, language) => {
      const result = await spell.checkWords!(words, language);
      if (!result || !Array.isArray(result.misspelled)) return { source: null, misspelled: [] };
      return result;
    },
    getSuggestions: async (word, language) => {
      const suggestions = await spell.getWordSuggestions!(word, language);
      return Array.isArray(suggestions) ? suggestions : [];
    },
  };
}
