/**
 * Word's Paragraph dialog (Indents and Spacing): left and right indents, a first-line
 * or hanging indent, space before and after, line spacing and "Don't add space
 * between paragraphs of the same style". Indents show in inches or centimetres
 * (by locale), spacing in points. Blank fields mean the selected paragraphs differ
 * and stay as they are.
 */
import type { ParaStyle } from "@forevka/wordcanvas/query";
import {
  fromPoints,
  fromUnit,
  readParagraphValues,
  toPoints,
  toUnit,
  type LengthUnit,
  type LineSpacingRule,
  type ParagraphValues,
  type SpecialIndent,
} from "./paragraph-format.ts";
import { openSettingsDialog, type SettingsDialogApi, type SettingsValues } from "./settings-dialog.ts";

const LINE_RULES: ReadonlyArray<{ value: LineSpacingRule; label: string }> = [
  { value: "single", label: "Single" },
  { value: "1.15", label: "1.15" },
  { value: "1.5", label: "1.5 lines" },
  { value: "double", label: "Double" },
  { value: "multiple", label: "Multiple" },
  { value: "atLeast", label: "At least" },
  { value: "exactly", label: "Exactly" },
];

const UNIT_LABEL: Record<LengthUnit, string> = { in: "in", cm: "cm" };
const NEEDS_AT = new Set<LineSpacingRule>(["multiple", "atLeast", "exactly"]);

function atUnit(rule: LineSpacingRule | null): string {
  return rule === "atLeast" || rule === "exactly" ? "pt" : "lines";
}

/** The "At" field value for a rule (multiple → lines, at least/exactly → points). */
function atValue(rule: LineSpacingRule | null, value: number | null): number | null {
  if (value === null || rule === null) return null;
  return rule === "atLeast" || rule === "exactly" ? toPoints(value) : Math.round(value * 100) / 100;
}

export interface ParagraphDialogResult {
  original: ParagraphValues;
  edited: ParagraphValues;
}

/** Opens the dialog for the selected paragraphs' styles; resolves the edited values, or null on Cancel. */
export async function openParagraphDialog(options: {
  styles: ReadonlyArray<Partial<ParaStyle>>;
  unit: LengthUnit;
  icon?: string;
  onClosed?: () => void;
}): Promise<ParagraphDialogResult | null> {
  const original = readParagraphValues(options.styles);
  const { unit } = options;
  const length = (px: number | null) => (px === null ? null : toUnit(px, unit));
  const points = (px: number | null) => (px === null ? null : toPoints(px));
  const step = unit === "in" ? 0.1 : 0.25;
  const many = options.styles.length > 1;
  const result = await openSettingsDialog({
    title: "Indents and spacing",
    icon: options.icon,
    message: many ? "Blank fields differ between the selected paragraphs and stay as they are." : undefined,
    submitLabel: "Apply",
    sections: [
      {
        title: "Indentation",
        controls: [
          { kind: "number", id: "left", label: "Left", value: length(original.left), unit: UNIT_LABEL[unit], step, min: -22, max: 22 },
          { kind: "number", id: "right", label: "Right", value: length(original.right), unit: UNIT_LABEL[unit], step, min: -22, max: 22 },
          {
            kind: "select",
            id: "special",
            label: "Special",
            value: original.special,
            options: [{ value: "none", label: "(none)" }, { value: "firstLine", label: "First line" }, { value: "hanging", label: "Hanging" }],
          },
          { kind: "number", id: "by", label: "By", value: original.special === "none" ? null : length(original.by), unit: UNIT_LABEL[unit], step, min: 0, max: 22, hidden: original.special === "none" },
        ],
      },
      {
        title: "Spacing",
        controls: [
          { kind: "number", id: "before", label: "Before", value: points(original.before), unit: "pt", step: 6, min: 0, max: 1584 },
          { kind: "number", id: "after", label: "After", value: points(original.after), unit: "pt", step: 6, min: 0, max: 1584 },
          { kind: "select", id: "line", label: "Line spacing", value: original.lineRule, options: LINE_RULES },
          {
            kind: "number",
            id: "at",
            label: "At",
            value: atValue(original.lineRule, original.lineValue),
            unit: atUnit(original.lineRule),
            step: original.lineRule === "multiple" ? 0.25 : 1,
            min: 0,
            max: 1584,
            hidden: !original.lineRule || !NEEDS_AT.has(original.lineRule),
          },
          { kind: "checkbox", id: "contextual", label: "Don't add space between paragraphs of the same style", value: original.contextualSpacing },
        ],
      },
    ],
    onChange(changed: string, dialog: SettingsDialogApi) {
      const values = dialog.values();
      if (changed === "special") {
        const special = values.special as SpecialIndent | null;
        dialog.setHidden("by", special === "none" || special === null);
        // Word's default amount for a new first-line or hanging indent: 0.5 in / 1.27 cm.
        if ((special === "firstLine" || special === "hanging") && !(Number(values.by) > 0)) dialog.set("by", toUnit(fromUnit(unit === "in" ? 0.5 : 1.27, unit), unit));
      }
      if (changed === "line") {
        const rule = values.line as LineSpacingRule | null;
        const needsAt = rule !== null && NEEDS_AT.has(rule);
        dialog.setHidden("at", !needsAt);
        dialog.setUnit("at", atUnit(rule));
        dialog.setNumberLimits("at", { step: rule === "multiple" ? 0.25 : 1, min: rule === "multiple" ? 0.06 : 0.7 });
        if (rule === "multiple") dialog.set("at", original.lineRule === "multiple" && original.lineValue ? original.lineValue : 3);
        else if (rule === "atLeast" || rule === "exactly") dialog.set("at", original.lineRule === rule && original.lineValue ? toPoints(original.lineValue) : 12);
      }
    },
    validate(values: SettingsValues) {
      const rule = values.line as LineSpacingRule | null;
      if (rule && NEEDS_AT.has(rule) && !(Number(values.at) > 0)) return { field: "at", message: rule === "multiple" ? "Type how many lines, for example 3." : "Type the line height in points, for example 12." };
      return null;
    },
    onClosed: options.onClosed,
  });
  if (!result) return null;
  const values = result.values;
  const number = (key: string) => (typeof values[key] === "number" && Number.isFinite(values[key]) ? values[key] as number : null);
  const rule = (values.line as LineSpacingRule | null) ?? null;
  const special = (values.special as SpecialIndent | null) ?? null;
  const at = number("at");
  const edited: ParagraphValues = {
    left: number("left") === null ? null : fromUnit(number("left")!, unit),
    right: number("right") === null ? null : fromUnit(number("right")!, unit),
    special,
    by: special === "none" ? 0 : number("by") === null ? null : fromUnit(number("by")!, unit),
    before: number("before") === null ? null : fromPoints(number("before")!),
    after: number("after") === null ? null : fromPoints(number("after")!),
    lineRule: rule,
    lineValue: rule === null ? null : rule === "multiple" ? at : rule === "atLeast" || rule === "exactly" ? (at === null ? null : fromPoints(at)) : original.lineRule === rule ? original.lineValue : null,
    contextualSpacing: typeof values.contextual === "boolean" ? values.contextual : null,
  };
  // Shown values are rounded (0.01 in, 0.1 pt): an untouched field must not count as a change.
  const keepIfShownSame = (key: "left" | "right" | "by", shown: (px: number | null) => number | null) => {
    if (edited[key] !== null && original[key] !== null && shown(edited[key]) === shown(original[key])) edited[key] = original[key];
  };
  keepIfShownSame("left", length);
  keepIfShownSame("right", length);
  keepIfShownSame("by", length);
  for (const key of ["before", "after"] as const) if (edited[key] !== null && original[key] !== null && points(edited[key]) === points(original[key])) edited[key] = original[key];
  if (edited.lineRule === original.lineRule && edited.lineValue !== null && original.lineValue !== null && atValue(edited.lineRule, edited.lineValue) === atValue(original.lineRule, original.lineValue)) edited.lineValue = original.lineValue;
  return { original, edited };
}
