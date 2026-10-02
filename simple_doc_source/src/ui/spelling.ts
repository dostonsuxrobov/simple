/**
 * Spelling as you type for Simple Docs: red underlines through the engine's screen-only
 * decorations (never printed or exported), and Word's right-click choices on top of the
 * editor's own context menu: suggestions, Ignore, Ignore all and Add to dictionary.
 *
 * The dictionary is the Windows Spell Checking API reached through the preload bridge
 * (fully offline); src/spellcheck.ts does the tokenizing, caching and batching. Checks
 * run shortly after each committed edit and only look up words not seen before, so
 * typing stays fast. The word being typed is left alone until the caret leaves it.
 * Review > Spelling turns the underlines off and on; the choice is remembered on this
 * computer.
 */
import type { EditorHandle } from "@forevka/wordcanvas";
import type { EngineBridge } from "../engine-bridge.ts";
import { preloadSpellBackend, SpellChecker, type PreloadSpellApi, type SpellIssue, type SpellMenu, type UserDictionary } from "../spellcheck.ts";
import { findParagraph, isCollapsed, paragraphText, type DocumentLike } from "./editing.ts";

export const SPELLING_KEY = "simple-docs:spelling";
/** Underlines drawn at most (Word also stops marking a document with very many errors). */
const MAX_MARKS = 3000;
const CHECK_DELAY_MS = 250;
const MAX_CHECK_DELAY_MS = 1200;

type StorageLike = Pick<Storage, "getItem" | "setItem">;

export interface SpellingOptions {
  bridge: EngineBridge;
  handle: EditorHandle;
  /** The element holding the editor (#editor): right-clicks on the page reach it. */
  host: HTMLElement;
  /** window.simpleDocs.spell, when the preload offers it. */
  api: Partial<PreloadSpellApi> | null | undefined;
  storage: StorageLike | null;
  notify(message: string, tone?: "normal" | "error"): void;
  focusEditor(): boolean;
  /** Replaces [start, end) of a paragraph as one undoable edit; false when it could not. */
  replaceText(blockId: string, start: number, end: number, text: string): boolean;
  /** Fallback language when Windows reports none (BCP 47). */
  language?: string;
  /** Called when the on/off state or availability changes (to refresh the ribbon toggle). */
  onStateChange?(): void;
}

export interface Spelling {
  readonly enabled: boolean;
  /** False when this computer has no spell checker for the language. */
  readonly available: boolean;
  readonly language: string;
  setEnabled(enabled: boolean): void;
  /** Re-checks soon (after an edit). */
  schedule(delay?: number): void;
  /** The caret moved: a word left behind gets checked. */
  selectionChanged(): void;
  /** A document was opened or replaced: everything is checked again. */
  documentReplaced(): void;
  /** The issues of the last check (for tests and command search). */
  issues(): readonly SpellIssue[];
  /** Checks now and resolves when the underlines are drawn. */
  checkNow(): Promise<readonly SpellIssue[]>;
  dispose(): void;
}

function readEnabled(storage: StorageLike | null): boolean {
  try {
    return storage?.getItem(SPELLING_KEY) !== "off";
  } catch {
    return true;
  }
}

function writeEnabled(storage: StorageLike | null, enabled: boolean) {
  try {
    storage?.setItem(SPELLING_KEY, enabled ? "on" : "off");
  } catch {
    // The choice still applies to this window.
  }
}

const occurrenceBase = (issue: SpellIssue) => `${issue.blockId}\u0000${issue.kind}\u0000${issue.word}`;

/** Which occurrence of the same word in its paragraph an issue is (for "Ignore" of one occurrence). */
export function occurrence(issue: SpellIssue, all: readonly SpellIssue[]): string {
  const base = occurrenceBase(issue);
  const index = all.filter((other) => other.start < issue.start && occurrenceBase(other) === base).length;
  return `${base}\u0000${index}`;
}

/** The issues not ignored one by one, in one pass (issues come in document order). */
export function withoutIgnored(all: readonly SpellIssue[], ignored: ReadonlySet<string>): readonly SpellIssue[] {
  if (!ignored.size) return all;
  const counts = new Map<string, number>();
  return all.filter((issue) => {
    const base = occurrenceBase(issue);
    const index = counts.get(base) ?? 0;
    counts.set(base, index + 1);
    return !ignored.has(`${base}\u0000${index}`);
  });
}

export function createSpelling(options: SpellingOptions): Spelling {
  const { bridge, handle } = options;
  const backend = preloadSpellBackend(options.api);
  let language = options.language || "en-US";
  const checker = backend ? new SpellChecker(backend, { language }) : null;
  let enabled = readEnabled(options.storage);
  let available = Boolean(checker);
  let timer: number | null = null;
  let firstScheduled = 0;
  let generation = 0;
  let disposed = false;
  let lastCaret = "";
  let shown: SpellIssue[] = [];
  const ignoredOnce = new Set<string>();
  const cleanups: Array<() => void> = [];

  const doc = () => handle.getDocument() as unknown as DocumentLike;
  const caretPosition = () => {
    const selection = bridge.getSelection();
    return selection && isCollapsed(selection) ? selection.focus : null;
  };
  const caretKey = () => {
    const caret = caretPosition();
    return caret ? `${caret.blockId}:${caret.offset}` : "";
  };

  const draw = () => {
    if (!checker || !enabled || handle.getMode() === "view") {
      bridge.clearDecorations();
      return;
    }
    shown = withoutIgnored(checker.issues, ignoredOnce).slice(0, MAX_MARKS);
    bridge.setDecorations(checker.decorations(shown));
  };

  const run = async () => {
    timer = null;
    firstScheduled = 0;
    if (!checker || !enabled || disposed) return shown;
    const current = generation += 1;
    const model = handle.getDocument();
    lastCaret = caretKey();
    try {
      await checker.checkDocument(model, { caret: caretPosition() });
    } catch (error) {
      console.warn("Spelling check failed", error);
      return shown;
    }
    if (disposed || current !== generation) return shown;
    if (!checker.available && available) {
      available = false;
      options.onStateChange?.();
    }
    // An edit landed while words were looked up: check again (cached words are instant).
    if (handle.getDocument() !== model) {
      schedule(60);
      return shown;
    }
    draw();
    return shown;
  };

  const schedule = (delay = CHECK_DELAY_MS) => {
    if (!checker || !enabled || disposed) return;
    const now = Date.now();
    if (!firstScheduled) firstScheduled = now;
    if (timer !== null) window.clearTimeout(timer);
    // Continuous typing still gets checked about once a second.
    const wait = Math.max(0, Math.min(delay, firstScheduled + MAX_CHECK_DELAY_MS - now));
    timer = window.setTimeout(() => void run(), wait);
  };

  const applyDictionary = (dictionary: Partial<UserDictionary> | null | undefined) => {
    if (!checker || !dictionary) return;
    checker.setUserDictionary(dictionary);
    schedule(0);
  };

  if (checker) {
    void Promise.resolve(options.api?.getUserDictionary?.()).then(applyDictionary).catch(() => {});
    const unsubscribe = options.api?.onDictionaryChanged?.((dictionary) => applyDictionary(dictionary));
    if (typeof unsubscribe === "function") cleanups.push(unsubscribe);
    void Promise.resolve(options.api?.getLanguages?.()).then((languages) => {
      if (!languages) return;
      if (languages.available === false) {
        available = false;
        options.onStateChange?.();
        return;
      }
      const preferred = languages.preferred || languages.languages?.[0];
      if (preferred && preferred !== language) {
        language = preferred;
        checker.setOptions({ language });
        schedule(0);
      }
    }).catch(() => {});
  }

  // ---- The context menu ---------------------------------------------------------------

  let menusBefore = new Set<Element>();
  const onContextMenuCapture = () => {
    menusBefore = new Set(document.querySelectorAll(".cw-menu"));
  };

  const closeMenus = () => {
    // A press outside closes the editor's menu through its own handler (which also
    // removes its window listeners); then any leftover element goes.
    document.body.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    document.querySelectorAll(".cw-menu").forEach((menu) => menu.remove());
  };

  const menuItem = (label: string, run: (() => void) | null, extra = "") => {
    const item = document.createElement("div");
    item.className = `cw-menu-item simple-spell-item${extra ? ` ${extra}` : ""}${run ? "" : " cw-disabled"}`;
    item.setAttribute("role", "menuitem");
    const icon = document.createElement("span");
    icon.className = "cw-menu-ico";
    const text = document.createElement("span");
    text.className = "cw-menu-lbl";
    text.textContent = label;
    item.append(icon, text);
    if (run) {
      item.addEventListener("mouseup", (event) => {
        if (event.button !== 0) return;
        event.preventDefault();
        event.stopPropagation();
        closeMenus();
        run();
      });
    }
    return item;
  };

  /** Moves a menu that grew back inside the window. */
  const keepInWindow = (menu: HTMLElement) => {
    const rect = menu.getBoundingClientRect();
    if (rect.bottom > window.innerHeight - 6) menu.style.top = `${Math.max(6, window.innerHeight - rect.height - 6)}px`;
    if (rect.right > window.innerWidth - 6) menu.style.left = `${Math.max(6, window.innerWidth - rect.width - 6)}px`;
  };

  const stillThere = (issue: SpellIssue) => {
    const block = findParagraph(doc(), issue.blockId);
    return Boolean(block) && paragraphText(block).slice(issue.start, issue.end) === issue.word;
  };

  const replaceIssue = (blockId: string, start: number, end: number, text: string) => {
    if (handle.getMode() === "view") {
      options.notify("Switch from Viewing to Editing to correct spelling.");
      return;
    }
    if (!options.replaceText(blockId, start, end, text)) options.notify("That word can't be changed here.", "error");
    options.focusEditor();
    schedule(0);
  };

  const fillMenu = (menu: HTMLElement, issue: SpellIssue, spellMenu: SpellMenu | null) => {
    const section = document.createElement("div");
    section.className = "simple-spell-section";
    const editable = handle.getMode() !== "view";
    if (!spellMenu) {
      section.append(menuItem("Looking for suggestions…", null, "simple-spell-pending"));
    } else if (spellMenu.suggestions.length) {
      for (const suggestion of spellMenu.suggestions) {
        section.append(menuItem(suggestion.label, editable ? () => replaceIssue(suggestion.blockId, suggestion.start, suggestion.end, suggestion.replacement) : null, "simple-spell-suggestion"));
      }
    } else {
      section.append(menuItem("No suggestions", null, "simple-spell-none"));
    }
    section.append(menuItem("Ignore", () => {
      ignoredOnce.add(occurrence(issue, checker!.issues));
      draw();
      options.focusEditor();
    }));
    if (issue.kind !== "repeated") {
      section.append(menuItem("Ignore all", () => {
        checker!.ignoreAll(issue.word);
        void Promise.resolve(options.api?.ignoreWord?.(issue.word)).catch(() => {});
        void run();
        options.focusEditor();
      }));
    }
    if (issue.kind === "spelling") {
      section.append(menuItem("Add to dictionary", () => {
        checker!.addToDictionary(issue.word);
        void Promise.resolve(options.api?.addWord?.(issue.word)).catch((error) => {
          console.warn("The word could not be added to the dictionary", error);
          options.notify("Simple Docs couldn't save that word to your dictionary.", "error");
        });
        void run();
        options.focusEditor();
      }));
    }
    const separator = document.createElement("div");
    separator.className = "cw-menu-sep";
    section.append(separator);
    menu.querySelector(":scope > .simple-spell-section")?.remove();
    menu.prepend(section);
    keepInWindow(menu);
  };

  const onContextMenu = (event: MouseEvent) => {
    if (!checker || !enabled || disposed || handle.getMode() === "view") return;
    const position = bridge.positionFromPoint(event.clientX, event.clientY);
    const issue = position ? shown.find((candidate) => candidate.blockId === position.blockId && position.offset >= candidate.start && position.offset <= candidate.end) ?? null : null;
    if (!issue || !stillThere(issue)) return;
    const menu = [...document.querySelectorAll<HTMLElement>(".cw-menu")].filter((element) => !menusBefore.has(element)).at(-1);
    if (!menu) return;
    fillMenu(menu, issue, null);
    void checker.menuFor(issue).then((spellMenu) => {
      if (menu.isConnected) fillMenu(menu, issue, spellMenu);
    }).catch(() => {
      if (menu.isConnected) fillMenu(menu, issue, { issue, suggestions: [], canAddToDictionary: issue.kind === "spelling", canIgnore: true });
    });
  };

  options.host.addEventListener("contextmenu", onContextMenuCapture, true);
  options.host.addEventListener("contextmenu", onContextMenu);
  cleanups.push(() => {
    options.host.removeEventListener("contextmenu", onContextMenuCapture, true);
    options.host.removeEventListener("contextmenu", onContextMenu);
  });

  return {
    get enabled() {
      return enabled;
    },
    get available() {
      return available;
    },
    get language() {
      return language;
    },
    setEnabled(next) {
      if (next === enabled) return;
      enabled = next;
      writeEnabled(options.storage, enabled);
      if (enabled) schedule(0);
      else {
        if (timer !== null) window.clearTimeout(timer);
        timer = null;
        generation += 1;
        shown = [];
        bridge.clearDecorations();
      }
      options.onStateChange?.();
    },
    schedule,
    selectionChanged() {
      if (!enabled || !checker) return;
      if (caretKey() !== lastCaret) schedule(120);
    },
    documentReplaced() {
      ignoredOnce.clear();
      shown = [];
      if (!enabled) {
        bridge.clearDecorations();
        return;
      }
      schedule(0);
    },
    issues: () => shown,
    async checkNow() {
      if (timer !== null) window.clearTimeout(timer);
      timer = null;
      return run();
    },
    dispose() {
      disposed = true;
      if (timer !== null) window.clearTimeout(timer);
      for (const cleanup of cleanups.splice(0)) cleanup();
    },
  };
}
