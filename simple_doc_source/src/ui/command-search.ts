/**
 * Command search, like Word's "Search" box (Alt+Q) and Docs' "Search the
 * menus" (Alt+/): type part of a command's name, press Enter to run it.
 * Recently used commands come first. The ranking and ribbon-title parsing are
 * pure and tested in tests/ui-command-search.test.cjs; createCommandSearch()
 * wires them to an input and a listbox.
 */

export interface SearchCommand {
  /** Stable id, e.g. "ribbon:home.paragraph.center" or "simple:save-as". */
  id: string;
  label: string;
  /** Where the command lives, e.g. "Home" or "File". */
  group?: string;
  keywords?: readonly string[];
  shortcut?: string;
  /** Why it is unavailable, or extra detail. */
  hint?: string;
  enabled: boolean;
  run(): void | Promise<unknown>;
}

export const RECENT_COMMANDS_KEY = "simple-docs:recent-commands";
const MAX_RECENT = 8;

/** Lower case, no accents, words separated by single spaces. */
export function normalizeSearchText(text: string): string {
  return String(text ?? "")
    .normalize("NFD")
    .replace(/\p{M}+/gu, "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim();
}

/**
 * Splits a ribbon tooltip into a command name, its shortcut and extra detail:
 * "Wrap text around image (square) — select an image first" →
 * { label: "Wrap text around image (square)", detail: "select an image first" };
 * "Bold (Ctrl+B)" → { label: "Bold", shortcut: "Ctrl+B" }.
 */
export function describeRibbonTitle(title: string): { label: string; shortcut?: string; detail?: string } {
  let text = String(title ?? "").replace(/\s+/g, " ").trim();
  let shortcut: string | undefined;
  text = text.replace(/\s*\(((?:Ctrl|Alt|Shift)\+[^)]*)\)/, (_match, value: string) => {
    shortcut = value.trim();
    return "";
  }).trim();
  const dash = text.indexOf(" — ");
  const label = (dash >= 0 ? text.slice(0, dash) : text).trim();
  const detail = dash >= 0 ? text.slice(dash + 3).trim() : "";
  return { label, ...(shortcut ? { shortcut } : {}), ...(detail ? { detail } : {}) };
}

function isSubsequence(needle: string, haystack: string): number {
  let at = 0;
  let gaps = 0;
  for (const character of needle) {
    const found = haystack.indexOf(character, at);
    if (found < 0) return -1;
    gaps += found - at;
    at = found + 1;
  }
  return gaps;
}

function textScore(query: string, text: string): number {
  if (!text) return -1;
  if (text === query) return 1000;
  if (text.startsWith(query)) return 800 - Math.min(100, text.length - query.length);
  const words = text.split(" ");
  const queryWords = query.split(" ");
  if (queryWords.length > 1 && queryWords.every((part) => words.some((word) => word.startsWith(part)))) return 650;
  if (words.some((word) => word.startsWith(query))) return 600;
  if (text.includes(query)) return 400;
  const compact = query.replace(/ /g, "");
  if (compact.length >= 2) {
    const gaps = isSubsequence(compact, text.replace(/ /g, ""));
    if (gaps >= 0 && gaps <= compact.length * 3) return 150 - gaps;
  }
  return -1;
}

/** How well a command matches a normalized query; negative when it does not. */
export function scoreCommand(query: string, command: SearchCommand): number {
  const label = textScore(query, normalizeSearchText(command.label));
  const keywords = textScore(query, normalizeSearchText((command.keywords ?? []).join(" ")));
  const group = textScore(query, normalizeSearchText(`${command.group ?? ""} ${command.label}`));
  return Math.max(label, keywords >= 0 ? keywords - 220 : -1, group >= 0 ? group - 300 : -1);
}

/**
 * Commands for a query, best first: matches by name, then by keyword, recently
 * used ones ahead of equal matches. An empty query lists recent commands, then
 * the suggested ones.
 */
export function rankCommands(query: string, commands: readonly SearchCommand[], recentIds: readonly string[] = [], options: { limit?: number; suggestedIds?: readonly string[] } = {}): SearchCommand[] {
  const limit = options.limit ?? 8;
  const byId = new Map(commands.map((command) => [command.id, command]));
  const recency = (id: string) => {
    const index = recentIds.indexOf(id);
    return index < 0 ? 0 : MAX_RECENT - index;
  };
  const normalized = normalizeSearchText(query);
  if (!normalized) {
    const picked: SearchCommand[] = [];
    for (const id of [...recentIds, ...(options.suggestedIds ?? [])]) {
      const command = byId.get(id);
      if (command && !picked.includes(command)) picked.push(command);
      if (picked.length >= limit) break;
    }
    return picked;
  }
  return commands
    .map((command, index) => ({ command, index, score: scoreCommand(normalized, command) }))
    .filter((entry) => entry.score >= 0)
    .map((entry) => ({ ...entry, score: entry.score + recency(entry.command.id) * 6 - (entry.command.enabled ? 0 : 40) }))
    .sort((left, right) => right.score - left.score || left.command.label.length - right.command.label.length || left.index - right.index)
    .slice(0, limit)
    .map((entry) => entry.command);
}

/** The recent list after running a command: it moves to the front, the list stays short. */
export function rememberRecent(recentIds: readonly string[], id: string): string[] {
  return [id, ...recentIds.filter((existing) => existing !== id)].slice(0, MAX_RECENT);
}

export interface CommandSearchOptions {
  input: HTMLInputElement;
  list: HTMLElement;
  /** Focus leaving this element closes the results. */
  container: HTMLElement;
  getCommands(): SearchCommand[];
  suggestedIds?: readonly string[];
  storage?: Pick<Storage, "getItem" | "setItem"> | null;
  /** Called before a command runs (the search box closes first), e.g. to return focus to the document. */
  beforeRun?(command: SearchCommand): void;
  /** Escape closed the search: put focus back where the person was working. */
  onEscape?(): void;
  /** A disabled command was chosen. */
  onUnavailable?(command: SearchCommand): void;
}

export interface CommandSearch {
  focus(): void;
  close(): void;
  isOpen(): boolean;
}

export function createCommandSearch(options: CommandSearchOptions): CommandSearch {
  const { input, list, container } = options;
  let results: SearchCommand[] = [];
  let active = -1;
  let open = false;

  const readRecent = (): string[] => {
    try {
      const value = JSON.parse(options.storage?.getItem(RECENT_COMMANDS_KEY) ?? "[]");
      return Array.isArray(value) ? value.filter((id): id is string => typeof id === "string").slice(0, MAX_RECENT) : [];
    } catch {
      return [];
    }
  };
  const writeRecent = (id: string) => {
    try {
      options.storage?.setItem(RECENT_COMMANDS_KEY, JSON.stringify(rememberRecent(readRecent(), id)));
    } catch {
      // A per-person convenience only; searching works without it.
    }
  };

  const setActive = (index: number) => {
    active = results.length ? (index + results.length) % results.length : -1;
    list.querySelectorAll<HTMLElement>("[role=option]").forEach((option, position) => {
      const selected = position === active;
      option.setAttribute("aria-selected", String(selected));
      if (selected) {
        input.setAttribute("aria-activedescendant", option.id);
        option.scrollIntoView({ block: "nearest" });
      }
    });
    if (active < 0) input.removeAttribute("aria-activedescendant");
  };

  const render = () => {
    const recent = readRecent();
    let commands: SearchCommand[] = [];
    try {
      commands = options.getCommands();
    } catch (error) {
      console.warn("Command search could not list commands", error);
    }
    results = rankCommands(input.value, commands, recent, { suggestedIds: options.suggestedIds });
    list.replaceChildren();
    if (!results.length) {
      const empty = document.createElement("div");
      empty.className = "command-search-empty";
      empty.textContent = input.value.trim() ? `No command matches “${input.value.trim()}”.` : "Type a command, like “link” or “heading”.";
      list.append(empty);
    }
    results.forEach((command, index) => {
      const option = document.createElement("div");
      option.id = `command-search-option-${index}`;
      option.className = "command-search-option";
      option.setAttribute("role", "option");
      option.setAttribute("aria-selected", "false");
      if (!command.enabled) option.setAttribute("aria-disabled", "true");
      const text = document.createElement("span");
      text.className = "command-search-text";
      const label = document.createElement("strong");
      label.textContent = command.label;
      text.append(label);
      const detail = [command.group, !command.enabled ? command.hint : ""].filter(Boolean).join(" · ");
      if (detail) {
        const small = document.createElement("small");
        small.textContent = detail;
        text.append(small);
      }
      option.append(text);
      if (command.shortcut) {
        const key = document.createElement("kbd");
        key.textContent = command.shortcut;
        option.append(key);
      }
      if (!input.value.trim() && recent.includes(command.id)) option.dataset.recent = "true";
      option.addEventListener("mousedown", (event) => event.preventDefault());
      option.addEventListener("mousemove", () => { if (active !== index) setActive(index); });
      option.addEventListener("click", () => void run(index));
      list.append(option);
    });
    setActive(results.length ? 0 : -1);
  };

  const show = () => {
    if (!open) {
      open = true;
      list.hidden = false;
      input.setAttribute("aria-expanded", "true");
      container.classList.add("is-open");
    }
    render();
  };

  const close = () => {
    if (!open) return;
    open = false;
    list.hidden = true;
    list.replaceChildren();
    results = [];
    active = -1;
    input.setAttribute("aria-expanded", "false");
    input.removeAttribute("aria-activedescendant");
    container.classList.remove("is-open");
  };

  const run = async (index: number) => {
    const command = results[index];
    if (!command) return;
    if (!command.enabled) {
      options.onUnavailable?.(command);
      return;
    }
    writeRecent(command.id);
    close();
    input.value = "";
    input.blur();
    options.beforeRun?.(command);
    try {
      await command.run();
    } catch (error) {
      console.error(`Command "${command.label}" failed`, error);
    }
  };

  input.addEventListener("focus", show);
  input.addEventListener("click", () => { if (!open) show(); });
  input.addEventListener("input", show);
  input.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      if (!open) show();
      else setActive(active + (event.key === "ArrowDown" ? 1 : -1));
    } else if (event.key === "Enter") {
      event.preventDefault();
      if (!open) show();
      else if (active >= 0) void run(active);
    } else if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      input.value = "";
      close();
      input.blur();
      options.onEscape?.();
    } else if (event.key === "Tab") {
      close();
    }
  });
  container.addEventListener("focusout", (event) => {
    const next = event.relatedTarget as Node | null;
    if (!next || !container.contains(next)) close();
  });

  return {
    focus() {
      input.focus();
      input.select();
      show();
    },
    close,
    isOpen: () => open,
  };
}
