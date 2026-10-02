/**
 * Layout-independent keyboard shortcuts.
 *
 * A layout that types ASCII letters keeps them (Ctrl+K on AZERTY or Dvorak is
 * the key labelled K, as in Word). Cyrillic, Greek, Hebrew, Arabic and other
 * layouts report their own character as `key`, so the physical key `code`
 * decides instead (Ctrl+Л on a Russian keyboard is Ctrl+K). Digits and
 * punctuation always follow the physical key, because their characters change
 * with Shift and with the layout. Matches electron/shortcuts.cjs.
 *
 * No DOM or runtime imports: Node tests load this module directly.
 */

export interface KeyEventLike {
  key: string;
  code?: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  altKey?: boolean;
  shiftKey?: boolean;
  isComposing?: boolean;
  keyCode?: number;
  getModifierState?(key: string): boolean;
}

export interface Combo {
  ctrl: boolean;
  alt: boolean;
  shift: boolean;
  /** Normalized key token: "a"…"z", "0"…"9", "=", "[", "]", ",", ".", "/", "\\", "Space", "F1"…"F12", "Tab"… */
  key: string;
}

const CODE_TOKENS: Record<string, string> = {
  BracketLeft: "[",
  BracketRight: "]",
  Equal: "=",
  Minus: "-",
  Period: ".",
  Comma: ",",
  Slash: "/",
  Backslash: "\\",
  Semicolon: ";",
  Quote: "'",
  Backquote: "`",
  Space: "Space",
};

const NAMED_KEYS = new Set(["Tab", "Enter", "Escape", "Backspace", "Delete", "Insert", "Home", "End", "PageUp", "PageDown", "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"]);

/** True when the key types an ASCII letter in the current layout. */
export function usesLatinKey(event: Pick<KeyEventLike, "key">): boolean {
  return /^[a-z]$/i.test(String(event?.key ?? ""));
}

/** The Latin letter of a shortcut, independent of the keyboard layout. */
export function shortcutLetter(event: Pick<KeyEventLike, "key" | "code">): string | null {
  const key = String(event?.key ?? "");
  if (/^[a-z]$/i.test(key)) return key.toLowerCase();
  const code = /^Key([A-Z])$/.exec(String(event?.code ?? ""));
  return code ? code[1].toLowerCase() : null;
}

/** The normalized key token of an event, or null for keys shortcuts never use. */
export function keyToken(event: Pick<KeyEventLike, "key" | "code">): string | null {
  const code = String(event?.code ?? "");
  const key = String(event?.key ?? "");
  const digit = /^Digit([0-9])$/.exec(code);
  if (digit) return digit[1];
  if (code in CODE_TOKENS) return CODE_TOKENS[code];
  const letter = shortcutLetter(event);
  if (letter) return letter;
  if (/^F(?:[1-9]|1[0-2])$/.test(key)) return key;
  if (NAMED_KEYS.has(key)) return key;
  if (key === " " || key === "Spacebar") return "Space";
  // Synthetic events without a code: fall back to the printable character.
  if (!code && key.length === 1 && /[0-9=\[\],./\\;'`-]/.test(key)) return key;
  return null;
}

/** AltGr (Ctrl+Alt on Windows) types characters on many layouts; never treat it as a shortcut. */
export function usesAltGraph(event: KeyEventLike): boolean {
  try {
    return Boolean(event?.getModifierState?.("AltGraph"));
  } catch {
    return false;
  }
}

/** An IME is composing (or the browser reports the composition key code 229). */
export function isComposing(event: KeyEventLike): boolean {
  return Boolean(event?.isComposing) || event?.keyCode === 229;
}

/** Parses "Ctrl+Shift+=", "Shift+F3", "Alt+Q" or "Ctrl+Space". */
export function parseCombo(text: string): Combo {
  const parts = String(text).split("+");
  // "Ctrl++" style combos are written with "=" instead; a trailing empty part means a literal "+".
  const keyPart = parts.pop() ?? "";
  const modifiers = new Set(parts.map((part) => part.trim().toLowerCase()));
  let key = keyPart.trim();
  if (/^[a-z]$/i.test(key)) key = key.toLowerCase();
  else if (/^space$/i.test(key)) key = "Space";
  else if (/^f(?:[1-9]|1[0-2])$/i.test(key)) key = key.toUpperCase();
  return { ctrl: modifiers.has("ctrl"), alt: modifiers.has("alt"), shift: modifiers.has("shift"), key };
}

/** True when the event is exactly this combo (extra modifiers never match). */
export function comboMatches(combo: Combo, event: KeyEventLike): boolean {
  const ctrl = Boolean(event.ctrlKey || event.metaKey);
  if (ctrl !== combo.ctrl || Boolean(event.altKey) !== combo.alt || Boolean(event.shiftKey) !== combo.shift) return false;
  return keyToken(event) === combo.key;
}

/** Display text for a combo, e.g. "Ctrl+Shift+=". */
export function formatCombo(text: string): string {
  const combo = parseCombo(text);
  const key = combo.key.length === 1 ? combo.key.toUpperCase() : combo.key;
  return [combo.ctrl && "Ctrl", combo.alt && "Alt", combo.shift && "Shift", key].filter(Boolean).join("+");
}

/** Appends "(Ctrl+K)" to a tooltip, replacing an older shortcut note in parentheses. */
export function withShortcutHint(title: string, combo: string): string {
  const hint = formatCombo(combo);
  const base = String(title ?? "").trim();
  if (!base) return hint;
  if (base.includes(`(${hint})`)) return base;
  const replaced = base.replace(/\s*\((?:Ctrl|Alt|Shift)\+[^)]*\)\s*$/, "");
  return `${replaced} (${hint})`;
}
