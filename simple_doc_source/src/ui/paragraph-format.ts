/**
 * Paragraph indents and spacing (Word's Paragraph dialog > Indents and Spacing, and
 * the line-spacing menu's Add/Remove space before/after paragraph), as pure functions
 * over the WordCanvas paragraph style.
 *
 * Lengths are CSS px in the model (96 per inch). The dialog shows indents in inches
 * (US) or centimetres and spacing in points, like Word. A multi-paragraph selection
 * with different values shows the field blank (null), and only the fields the person
 * changed are applied, so mixed formatting is kept.
 *
 * No DOM or runtime imports: Node tests load this module directly.
 */
import type { ParaStyle } from "@forevka/wordcanvas/query";

export type LengthUnit = "in" | "cm";
export type SpecialIndent = "none" | "firstLine" | "hanging";
export type LineSpacingRule = "single" | "1.15" | "1.5" | "double" | "multiple" | "atLeast" | "exactly";

/** What the Paragraph dialog shows; null means "the selected paragraphs differ". */
export interface ParagraphValues {
  /** Left indent, px. */
  left: number | null;
  /** Right indent, px. */
  right: number | null;
  special: SpecialIndent | null;
  /** First-line or hanging amount, px (0 with "none"). */
  by: number | null;
  /** Space before, px. */
  before: number | null;
  /** Space after, px. */
  after: number | null;
  lineRule: LineSpacingRule | null;
  /** The multiple for "multiple" (and the presets), or px for "atLeast"/"exactly". */
  lineValue: number | null;
  /** "Don't add space between paragraphs of the same style". */
  contextualSpacing: boolean | null;
}

export const PX_PER_INCH = 96;
export const PX_PER_CM = 96 / 2.54;
export const PX_PER_PT = 96 / 72;
/** Word's Add Space Before/After Paragraph amount. */
export const SPACE_STEP_PT = 12;
const EPSILON = 0.01;

/** Inches where people measure in inches (United States, Liberia, Myanmar), else centimetres. */
export function unitForLocale(locale: string | null | undefined): LengthUnit {
  const parts = String(locale || "en-US").split(/[-_]/);
  const region = parts.slice(1).find((part) => /^[A-Za-z]{2}$/.test(part))?.toUpperCase();
  if (region) return region === "US" || region === "LR" || region === "MM" ? "in" : "cm";
  return parts[0].toLowerCase() === "en" ? "in" : "cm";
}

export function pxPerUnit(unit: LengthUnit): number {
  return unit === "in" ? PX_PER_INCH : PX_PER_CM;
}

/** A px length in the unit, rounded like Word shows it (0.01 in, 0.01 cm). */
export function toUnit(px: number, unit: LengthUnit): number {
  return Math.round((px / pxPerUnit(unit)) * 100) / 100;
}

export function fromUnit(value: number, unit: LengthUnit): number {
  return value * pxPerUnit(unit);
}

/** Points, rounded to 0.1 pt. */
export function toPoints(px: number): number {
  return Math.round((px / PX_PER_PT) * 10) / 10;
}

export function fromPoints(points: number): number {
  return points * PX_PER_PT;
}

const same = (a: number, b: number) => Math.abs(a - b) < EPSILON;

function lineSpacingOf(style: Partial<ParaStyle>): { rule: LineSpacingRule; value: number } {
  if (style.lineRule === "exact" && Number.isFinite(style.lineHeightPx)) return { rule: "exactly", value: style.lineHeightPx! };
  if (style.lineRule === "atLeast" && Number.isFinite(style.lineHeightPx)) return { rule: "atLeast", value: style.lineHeightPx! };
  const multiple = Number.isFinite(style.lineHeight) && style.lineHeight! > 0 ? style.lineHeight! : 1;
  if (same(multiple, 1)) return { rule: "single", value: 1 };
  if (same(multiple, 1.15)) return { rule: "1.15", value: 1.15 };
  if (same(multiple, 1.5)) return { rule: "1.5", value: 1.5 };
  if (same(multiple, 2)) return { rule: "double", value: 2 };
  return { rule: "multiple", value: multiple };
}

function valuesOf(style: Partial<ParaStyle>): ParagraphValues {
  const first = Number(style.indentFirstLinePx) || 0;
  const spacing = lineSpacingOf(style);
  return {
    left: Number(style.indentLeftPx) || 0,
    right: Number(style.indentRightPx) || 0,
    special: first > EPSILON ? "firstLine" : first < -EPSILON ? "hanging" : "none",
    by: Math.abs(first),
    before: Number(style.spaceBeforePx) || 0,
    after: Number(style.spaceAfterPx) || 0,
    lineRule: spacing.rule,
    lineValue: spacing.value,
    contextualSpacing: style.contextualSpacing === true,
  };
}

/** The dialog values for the selected paragraphs (null where they differ). */
export function readParagraphValues(styles: ReadonlyArray<Partial<ParaStyle>>): ParagraphValues {
  const all = styles.map(valuesOf);
  if (all.length === 0) return valuesOf({});
  const agree = <K extends keyof ParagraphValues>(key: K): ParagraphValues[K] => {
    const value = all[0][key];
    for (const other of all) {
      const candidate = other[key];
      if (typeof value === "number" && typeof candidate === "number" ? !same(value, candidate) : candidate !== value) return null as ParagraphValues[K];
    }
    return value;
  };
  const values: ParagraphValues = {
    left: agree("left"),
    right: agree("right"),
    special: agree("special"),
    by: agree("by"),
    before: agree("before"),
    after: agree("after"),
    lineRule: agree("lineRule"),
    lineValue: agree("lineValue"),
    contextualSpacing: agree("contextualSpacing"),
  };
  if (values.lineRule === null) values.lineValue = null;
  return values;
}

/** The line-spacing part of a paragraph patch. */
export function lineSpacingPatch(rule: LineSpacingRule, value: number | null): Partial<ParaStyle> {
  const presets: Partial<Record<LineSpacingRule, number>> = { single: 1, "1.15": 1.15, "1.5": 1.5, double: 2 };
  const preset = presets[rule];
  if (preset !== undefined) return { lineHeight: preset, lineRule: undefined, lineHeightPx: undefined };
  if (rule === "multiple") {
    const multiple = Number.isFinite(value) && value! > 0 ? Math.min(132, Math.max(0.06, value!)) : 1;
    return { lineHeight: multiple, lineRule: undefined, lineHeightPx: undefined };
  }
  const px = Number.isFinite(value) && value! > 0 ? Math.min(fromPoints(1584), Math.max(fromPoints(0.7), value!)) : fromPoints(12);
  return { lineRule: rule === "exactly" ? "exact" : "atLeast", lineHeightPx: px };
}

/**
 * The paragraph style patch for what changed between `original` and `edited`. A field
 * left blank (null) or unchanged is not part of the patch. `base` is one paragraph's
 * current style, so a changed "Special" keeps that paragraph's amount when "By" was
 * left blank.
 */
export function paragraphPatch(original: ParagraphValues, edited: ParagraphValues, base: Partial<ParaStyle> = {}): Partial<ParaStyle> {
  const patch: Partial<ParaStyle> = {};
  const changed = (key: keyof ParagraphValues) => {
    const a = original[key];
    const b = edited[key];
    if (b === null) return false;
    if (typeof a === "number" && typeof b === "number") return !same(a, b);
    return a !== b;
  };
  const clamp = (px: number, min: number, max: number) => Math.min(max, Math.max(min, px));
  const LIMIT = fromPoints(1584);
  if (changed("left")) patch.indentLeftPx = clamp(edited.left!, -LIMIT, LIMIT);
  if (changed("right")) patch.indentRightPx = clamp(edited.right!, -LIMIT, LIMIT);
  if (changed("special") || changed("by")) {
    const special = edited.special ?? valuesOf(base).special ?? "none";
    const by = edited.by ?? valuesOf(base).by ?? 0;
    patch.indentFirstLinePx = special === "none" ? 0 : special === "firstLine" ? clamp(by, 0, LIMIT) : -clamp(by, 0, LIMIT);
  }
  if (changed("before")) patch.spaceBeforePx = clamp(edited.before!, 0, LIMIT);
  if (changed("after")) patch.spaceAfterPx = clamp(edited.after!, 0, LIMIT);
  if (changed("lineRule") || changed("lineValue")) {
    const rule = edited.lineRule ?? valuesOf(base).lineRule ?? "single";
    const value = edited.lineValue ?? (edited.lineRule === original.lineRule ? valuesOf(base).lineValue : null);
    Object.assign(patch, lineSpacingPatch(rule, value));
  }
  if (changed("contextualSpacing")) patch.contextualSpacing = edited.contextualSpacing!;
  return patch;
}

/** Word's line-spacing menu toggle: the label and the new value for space before or after. */
export function spaceToggle(styles: ReadonlyArray<Partial<ParaStyle>>, which: "before" | "after"): { label: string; add: boolean; px: number } {
  const key = which === "before" ? "spaceBeforePx" : "spaceAfterPx";
  // Like Word, the first selected paragraph decides between Add and Remove.
  const current = Number(styles[0]?.[key]) || 0;
  const add = current <= EPSILON;
  return { label: `${add ? "Add" : "Remove"} space ${which} paragraph`, add, px: add ? fromPoints(SPACE_STEP_PT) : 0 };
}

/**
 * Reads a length typed in a field: a plain number is in `unit`; "2 cm", "0.5in",
 * '0.5"', "10 mm", "12 pt" and "16px" are converted. Null when it is not a length.
 */
export function parseLength(text: string, unit: LengthUnit | "pt"): number | null {
  const match = /^\s*(-?\d+(?:[.,]\d+)?|-?[.,]\d+)\s*(cm|mm|in|inch|inches|"|″|pt|px)?\s*$/i.exec(String(text ?? ""));
  if (!match) return null;
  const value = Number(match[1].replace(",", "."));
  if (!Number.isFinite(value)) return null;
  const suffix = (match[2] ?? "").toLowerCase();
  const as = suffix === "" ? unit : suffix === '"' || suffix === "″" || suffix.startsWith("in") ? "in" : suffix;
  switch (as) {
    case "in": return value * PX_PER_INCH;
    case "cm": return value * PX_PER_CM;
    case "mm": return value * PX_PER_CM / 10;
    case "pt": return value * PX_PER_PT;
    case "px": return value;
    default: return null;
  }
}
