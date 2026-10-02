/**
 * The CSS subset the importers understand: inline declarations, colours, lengths and
 * simple author rules (tag, .class, tag.class) from <style> blocks. Nothing here loads
 * a resource; url(...) values are ignored.
 */
import type { DomElement } from "./markup.ts";

export type Declarations = Record<string, string>;

const PX_PER_PT = 96 / 72;

/** Parses "a: b; c: d" into lower-case property names; later declarations win. */
export function parseDeclarations(text: string | undefined): Declarations {
  const result: Declarations = {};
  if (!text) return result;
  let depth = 0;
  let quote = "";
  let start = 0;
  const flush = (end: number) => {
    const part = text.slice(start, end);
    const colon = part.indexOf(":");
    if (colon > 0) {
      const name = part.slice(0, colon).trim().toLowerCase();
      const value = part.slice(colon + 1).replace(/!\s*important\s*$/i, "").trim();
      if (name && value) result[name] = value;
    }
    start = end + 1;
  };
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (quote) {
      if (ch === quote) quote = "";
    } else if (ch === "\"" || ch === "'") quote = ch;
    else if (ch === "(") depth += 1;
    else if (ch === ")") depth = Math.max(0, depth - 1);
    else if (ch === ";" && depth === 0) flush(i);
  }
  flush(text.length);
  return result;
}

const NAMED_COLORS: Record<string, string> = {
  black: "#000000", white: "#ffffff", red: "#ff0000", green: "#008000", blue: "#0000ff", yellow: "#ffff00", cyan: "#00ffff", aqua: "#00ffff",
  magenta: "#ff00ff", fuchsia: "#ff00ff", gray: "#808080", grey: "#808080", silver: "#c0c0c0", maroon: "#800000", olive: "#808000",
  lime: "#00ff00", navy: "#000080", purple: "#800080", teal: "#008080", orange: "#ffa500", pink: "#ffc0cb", brown: "#a52a2a", gold: "#ffd700",
  darkred: "#8b0000", darkblue: "#00008b", darkgreen: "#006400", darkcyan: "#008b8b", darkmagenta: "#8b008b", darkgray: "#a9a9a9",
  darkgrey: "#a9a9a9", lightgray: "#d3d3d3", lightgrey: "#d3d3d3", lightyellow: "#ffffe0", lightblue: "#add8e6", lightgreen: "#90ee90",
  lightpink: "#ffb6c1", darkorange: "#ff8c00", indigo: "#4b0082", violet: "#ee82ee", crimson: "#dc143c", coral: "#ff7f50", salmon: "#fa8072",
  khaki: "#f0e68c", beige: "#f5f5dc", ivory: "#fffff0", tan: "#d2b48c", turquoise: "#40e0d0", skyblue: "#87ceeb", steelblue: "#4682b4",
  royalblue: "#4169e1", dodgerblue: "#1e90ff", slategray: "#708090", dimgray: "#696969", whitesmoke: "#f5f5f5", gainsboro: "#dcdcdc",
  windowtext: "#000000", buttontext: "#000000", canvastext: "#000000", window: "#ffffff", canvas: "#ffffff",
};

const hex2 = (value: number) => Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0");

/** A CSS colour as "#rrggbb", or null for transparent, "auto", inherit or garbage. */
export function parseColor(value: string | undefined | null): string | null {
  if (!value) return null;
  const color = value.trim().toLowerCase();
  if (!color || color === "transparent" || color === "auto" || color === "inherit" || color === "initial" || color === "currentcolor" || color === "none") return null;
  if (NAMED_COLORS[color]) return NAMED_COLORS[color];
  let match = /^#([0-9a-f]{3,8})$/.exec(color);
  if (match) {
    const digits = match[1];
    if (digits.length === 3 || digits.length === 4) {
      if (digits.length === 4 && digits[3] === "0") return null;
      return `#${digits[0]}${digits[0]}${digits[1]}${digits[1]}${digits[2]}${digits[2]}`;
    }
    if (digits.length === 6) return `#${digits}`;
    if (digits.length === 8) return digits.slice(6) === "00" ? null : `#${digits.slice(0, 6)}`;
    return null;
  }
  match = /^rgba?\(\s*([\d.]+)(%?)[\s,]+([\d.]+)(%?)[\s,]+([\d.]+)(%?)\s*(?:[,/]\s*([\d.]+)(%?))?\s*\)$/.exec(color);
  if (match) {
    const channel = (number: string, percent: string) => percent ? Number(number) * 2.55 : Number(number);
    if (match[7] !== undefined && Number(match[7]) === 0) return null;
    return `#${hex2(channel(match[1], match[2]))}${hex2(channel(match[3], match[4]))}${hex2(channel(match[5], match[6]))}`;
  }
  return null;
}

/**
 * A CSS/ODF length in px. `basePx` resolves em and % (font sizes). Returns null for
 * keywords and garbage. Unitless numbers are px (HTML attributes), except 0.
 */
export function parseLength(value: string | undefined | null, basePx = 16): number | null {
  if (value === undefined || value === null) return null;
  const match = /^\s*(-?(?:\d+\.?\d*|\.\d+))\s*(px|pt|pc|in|cm|mm|em|rem|ex|ch|%|q)?\s*$/i.exec(value);
  if (!match) return null;
  const number = Number(match[1]);
  if (!Number.isFinite(number)) return null;
  switch ((match[2] ?? "px").toLowerCase()) {
    case "px": return number;
    case "pt": return number * PX_PER_PT;
    case "pc": return number * 16;
    case "in": return number * 96;
    case "cm": return number * 96 / 2.54;
    case "mm": return number * 96 / 25.4;
    case "q": return number * 96 / 101.6;
    case "em": case "rem": return number * basePx;
    case "ex": case "ch": return number * basePx / 2;
    case "%": return number * basePx / 100;
    default: return null;
  }
}

const FONT_SIZE_KEYWORDS: Record<string, number> = {
  "xx-small": 9, "x-small": 10, small: 13, medium: 16, large: 18, "x-large": 24, "xx-large": 32, "xxx-large": 48,
};

/** font-size in px relative to the parent size. */
export function parseFontSize(value: string | undefined, parentPx: number): number | null {
  if (!value) return null;
  const keyword = value.trim().toLowerCase();
  if (FONT_SIZE_KEYWORDS[keyword]) return FONT_SIZE_KEYWORDS[keyword];
  if (keyword === "smaller") return parentPx / 1.2;
  if (keyword === "larger") return parentPx * 1.2;
  const px = parseLength(value, parentPx);
  return px !== null && px > 0 ? Math.min(px, 1600) : null;
}

/** First family of a font-family list, unquoted ("Georgia, serif" -> "Georgia"). */
export function firstFontFamily(value: string | undefined): string | null {
  if (!value) return null;
  const GENERIC: Record<string, string> = { serif: "Times New Roman", "sans-serif": "Arial", monospace: "Consolas", cursive: "Comic Sans MS", "system-ui": "Segoe UI" };
  for (const part of value.split(",")) {
    const name = part.trim().replace(/^["']|["']$/g, "").trim();
    if (!name || /^(?:inherit|initial|unset|-apple-system|blinkmacsystemfont)$/i.test(name)) continue;
    return GENERIC[name.toLowerCase()] ?? name;
  }
  return null;
}

/** The four values of a margin/padding shorthand: [top, right, bottom, left]. */
export function boxValues(value: string | undefined): [string, string, string, string] | null {
  if (!value) return null;
  const parts = value.trim().split(/\s+/);
  if (!parts.length || parts.length > 4) return null;
  const [top, right = top, bottom = top, left = right] = parts;
  return [top, right, bottom, left];
}

interface StyleRule {
  tag: string | null;
  className: string | null;
  declarations: Declarations;
  order: number;
  specificity: number;
}

export interface StyleSheet {
  rules: StyleRule[];
}

/** Collects simple rules (tag, .class, tag.class) from CSS text; everything else is ignored. */
export function parseStyleSheet(cssText: string): StyleSheet {
  const rules: StyleRule[] = [];
  const text = cssText.replace(/\/\*[\s\S]*?\*\//g, "").replace(/<!--|-->/g, "");
  let index = 0;
  let order = 0;
  while (index < text.length) {
    const open = text.indexOf("{", index);
    if (open < 0) break;
    const selector = text.slice(index, open).trim();
    // Find the matching close brace (at-rules may nest).
    let depth = 1;
    let cursor = open + 1;
    while (cursor < text.length && depth > 0) {
      if (text[cursor] === "{") depth += 1;
      else if (text[cursor] === "}") depth -= 1;
      cursor += 1;
    }
    const body = text.slice(open + 1, cursor - 1);
    index = cursor;
    if (selector.startsWith("@")) continue;
    const declarations = parseDeclarations(body);
    for (const raw of selector.split(",")) {
      const match = /^([a-z][a-z0-9]*)?(?:\.([A-Za-z_][\w-]*))?$/i.exec(raw.trim());
      if (!match || (!match[1] && !match[2])) continue;
      const tag = match[1]?.toLowerCase() ?? null;
      const className = match[2] ?? null;
      rules.push({ tag, className, declarations, order: order++, specificity: (tag ? 1 : 0) + (className ? 10 : 0) });
    }
  }
  return { rules };
}

// Properties a bare tag rule ("p { margin: .25em 0 }") may set. Page-oriented exports use
// tag rules for screen layout, so box spacing only comes from class rules and inline styles.
const TAG_RULE_PROPERTIES = new Set([
  "font-family", "font-size", "font-weight", "font-style", "color", "text-decoration", "text-decoration-line", "text-align",
  "vertical-align", "text-transform", "font-variant", "background-color", "direction",
]);
const SKIPPED_TAGS = new Set(["html", "body", "table", "td", "th", "tr", "tbody", "thead", "tfoot", "div", "section", "main", "article", "img", "figure", "a"]);

/** Declarations for an element: matching author rules, then its style attribute / css. */
export function computeDeclarations(element: DomElement, sheet: StyleSheet | null): Declarations {
  const inline = element.css ?? parseDeclarations(element.attrs.style);
  if (!sheet || !sheet.rules.length) return inline;
  const classes = (element.attrs.class ?? "").split(/\s+/).filter(Boolean).map((name) => name.toLowerCase());
  const matching = sheet.rules.filter((rule) => {
    if (rule.tag && rule.tag !== element.name) return false;
    if (rule.className && !classes.includes(rule.className.toLowerCase())) return false;
    if (!rule.className && SKIPPED_TAGS.has(rule.tag ?? "")) return false;
    return true;
  });
  if (!matching.length) return inline;
  matching.sort((a, b) => a.specificity - b.specificity || a.order - b.order);
  const result: Declarations = {};
  for (const rule of matching) {
    for (const [name, value] of Object.entries(rule.declarations)) {
      if (!rule.className && !TAG_RULE_PROPERTIES.has(name)) continue;
      result[name] = value;
    }
  }
  return Object.assign(result, inline);
}
