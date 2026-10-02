/**
 * Word and Google Docs keyboard shortcuts that WordCanvas 0.12.0 does not
 * handle itself. Commands are run by src/ui/editor-commands.ts, mostly by
 * clicking the engine's own ribbon controls (data-ribbon-item ids, verified in
 * tests/ui-shortcuts.test.cjs), so a shortcut behaves exactly like its button.
 *
 * No DOM or runtime imports: Node tests load this module directly.
 */
import { comboMatches, isComposing, parseCombo, usesAltGraph, usesLatinKey, type Combo, type KeyEventLike } from "./keys.ts";

export type ShortcutCommand =
  // Editing commands (the document has focus)
  | "find"
  | "replace"
  | "link"
  | "align-left"
  | "align-center"
  | "align-right"
  | "justify"
  | "grow-font"
  | "shrink-font"
  | "subscript"
  | "superscript"
  | "line-spacing-1"
  | "line-spacing-1.5"
  | "line-spacing-2"
  | "heading-1"
  | "heading-2"
  | "heading-3"
  | "normal-style"
  | "clear-character-formatting"
  | "comment"
  | "footnote"
  | "endnote"
  | "change-case"
  | "formatting-marks"
  // Engine shortcuts that only follow `key`: repeated here for non-Latin layouts
  | "undo"
  | "redo"
  | "bold"
  | "italic"
  | "underline"
  | "select-all"
  // Window commands (any focus, no dialog open)
  | "save-as"
  | "close"
  | "command-search";

export interface ShortcutBinding {
  command: ShortcutCommand;
  /** As written in tooltips, e.g. "Ctrl+Shift+=". */
  combo: string;
  /** "editor": the document has focus; "app": anywhere in the window. */
  scope: "editor" | "app";
  /** Only when the layout does not type Latin letters (the engine handles Latin itself). */
  nonLatinOnly?: boolean;
}

export const SHORTCUTS: readonly ShortcutBinding[] = [
  { command: "find", combo: "Ctrl+F", scope: "editor" },
  { command: "replace", combo: "Ctrl+H", scope: "editor" },
  { command: "link", combo: "Ctrl+K", scope: "editor" },
  { command: "align-left", combo: "Ctrl+L", scope: "editor" },
  { command: "align-center", combo: "Ctrl+E", scope: "editor" },
  { command: "align-right", combo: "Ctrl+R", scope: "editor" },
  { command: "justify", combo: "Ctrl+J", scope: "editor" },
  { command: "grow-font", combo: "Ctrl+]", scope: "editor" },
  { command: "shrink-font", combo: "Ctrl+[", scope: "editor" },
  { command: "grow-font", combo: "Ctrl+Shift+.", scope: "editor" },
  { command: "shrink-font", combo: "Ctrl+Shift+,", scope: "editor" },
  { command: "subscript", combo: "Ctrl+=", scope: "editor" },
  { command: "superscript", combo: "Ctrl+Shift+=", scope: "editor" },
  { command: "line-spacing-1", combo: "Ctrl+1", scope: "editor" },
  { command: "line-spacing-2", combo: "Ctrl+2", scope: "editor" },
  { command: "line-spacing-1.5", combo: "Ctrl+5", scope: "editor" },
  { command: "heading-1", combo: "Ctrl+Alt+1", scope: "editor" },
  { command: "heading-2", combo: "Ctrl+Alt+2", scope: "editor" },
  { command: "heading-3", combo: "Ctrl+Alt+3", scope: "editor" },
  { command: "normal-style", combo: "Ctrl+Shift+N", scope: "editor" },
  { command: "clear-character-formatting", combo: "Ctrl+Space", scope: "editor" },
  { command: "comment", combo: "Ctrl+Alt+M", scope: "editor" },
  { command: "footnote", combo: "Ctrl+Alt+F", scope: "editor" },
  { command: "endnote", combo: "Ctrl+Alt+D", scope: "editor" },
  { command: "change-case", combo: "Shift+F3", scope: "editor" },
  { command: "formatting-marks", combo: "Ctrl+Shift+8", scope: "editor" },
  { command: "undo", combo: "Ctrl+Z", scope: "editor", nonLatinOnly: true },
  { command: "redo", combo: "Ctrl+Y", scope: "editor", nonLatinOnly: true },
  { command: "redo", combo: "Ctrl+Shift+Z", scope: "editor", nonLatinOnly: true },
  { command: "bold", combo: "Ctrl+B", scope: "editor", nonLatinOnly: true },
  { command: "italic", combo: "Ctrl+I", scope: "editor", nonLatinOnly: true },
  { command: "underline", combo: "Ctrl+U", scope: "editor", nonLatinOnly: true },
  { command: "select-all", combo: "Ctrl+A", scope: "editor", nonLatinOnly: true },
  { command: "save-as", combo: "F12", scope: "app" },
  { command: "close", combo: "Ctrl+W", scope: "app" },
  { command: "command-search", combo: "Alt+Q", scope: "app" },
  { command: "command-search", combo: "Alt+/", scope: "app" },
];

/**
 * Ribbon items the commands click (WordCanvas 0.12.0 ids derive from button
 * labels; tests/ui-shortcuts.test.cjs pins them against the verified inventory).
 */
export const SHORTCUT_RIBBON_ITEMS: Partial<Record<ShortcutCommand, string>> = {
  find: "home.editing.find-replace",
  replace: "home.editing.replace",
  link: "insert.links.insert-remove-hyperlink",
  "align-left": "home.paragraph.align-left",
  "align-center": "home.paragraph.center",
  "align-right": "home.paragraph.align-right",
  justify: "home.paragraph.justify",
  "grow-font": "home.font.grow-font",
  "shrink-font": "home.font.shrink-font",
  subscript: "home.font.subscript",
  superscript: "home.font.superscript",
  "line-spacing-1": "home.paragraph.line-spacing",
  "line-spacing-1.5": "home.paragraph.line-spacing",
  "line-spacing-2": "home.paragraph.line-spacing",
  "clear-character-formatting": "home.font.clear-all-formatting",
  footnote: "insert.references.insert-footnote",
  endnote: "insert.references.insert-endnote",
  "change-case": "home.font.change-case",
  "formatting-marks": "home.paragraph.show-hide-formatting-marks",
  bold: "home.font.bold",
  italic: "home.font.italic",
  underline: "home.font.underline",
  "select-all": "home.editing.select-all",
};

const parsed = new Map<string, Combo>();
const comboOf = (text: string) => {
  let combo = parsed.get(text);
  if (!combo) parsed.set(text, combo = parseCombo(text));
  return combo;
};

/**
 * The binding this key event triggers, or null. AltGr characters, IME
 * composition and Alt shortcuts that would type on the layout never match.
 */
export function findShortcut(event: KeyEventLike, bindings: readonly ShortcutBinding[] = SHORTCUTS): ShortcutBinding | null {
  if (!event || isComposing(event) || usesAltGraph(event)) return null;
  for (const binding of bindings) {
    if (binding.nonLatinOnly && usesLatinKey(event)) continue;
    if (comboMatches(comboOf(binding.combo), event)) return binding;
  }
  return null;
}

/** The first (tooltip) combo of a command. */
export function shortcutFor(command: ShortcutCommand, bindings: readonly ShortcutBinding[] = SHORTCUTS): string | undefined {
  return bindings.find((binding) => binding.command === command && !binding.nonLatinOnly)?.combo;
}

/**
 * The shortcut of the command that clicks this ribbon control. None for the
 * line spacing menu (three shortcuts) and for "Clear all formatting", which
 * also resets the paragraph style while Ctrl+Space keeps it.
 */
export function ribbonItemShortcut(ribbonItemId: string): string | undefined {
  const commands = (Object.entries(SHORTCUT_RIBBON_ITEMS) as Array<[ShortcutCommand, string]>)
    .filter(([command, id]) => id === ribbonItemId && command !== "clear-character-formatting")
    .map(([command]) => command);
  return commands.length === 1 ? shortcutFor(commands[0]) : undefined;
}
