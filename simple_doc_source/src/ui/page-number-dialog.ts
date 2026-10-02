/**
 * Insert > Page number: where the number goes (top or bottom, left, center or right),
 * how it reads (1, Page 1, Page 1 of 3), its format (1, i, I, a, A), the first number
 * and whether the first page shows it, like Word's Page Number gallery plus Format
 * Page Numbers and Docs' Page numbers dialog in one small dialog.
 */
import type { PageNumberAlign, PageNumberFormat, PageNumberOptions, PageNumberText } from "./header-footer.ts";
import { openSettingsDialog } from "./settings-dialog.ts";

export type PageNumberDialogResult = { action: "insert"; options: PageNumberOptions } | { action: "remove" };

export async function openPageNumberDialog(options: {
  current: PageNumberOptions;
  hasPageNumbers: boolean;
  icon?: string;
  onClosed?: () => void;
}): Promise<PageNumberDialogResult | null> {
  const { current } = options;
  const result = await openSettingsDialog({
    title: options.hasPageNumbers ? "Page numbers" : "Insert page numbers",
    icon: options.icon,
    submitLabel: options.hasPageNumbers ? "Apply" : "Insert",
    ...(options.hasPageNumbers ? { extraAction: { id: "remove", label: "Remove page numbers" } } : {}),
    sections: [
      {
        controls: [
          { kind: "select", id: "position", label: "Position", value: current.position, options: [{ value: "top", label: "Top of page (header)" }, { value: "bottom", label: "Bottom of page (footer)" }] },
          { kind: "select", id: "align", label: "Alignment", value: current.align, options: [{ value: "left", label: "Left" }, { value: "center", label: "Center" }, { value: "right", label: "Right" }] },
          { kind: "select", id: "text", label: "Show", value: current.text, options: [{ value: "number", label: "1" }, { value: "page-x", label: "Page 1" }, { value: "page-x-of-y", label: "Page 1 of 3" }] },
          {
            kind: "select",
            id: "format",
            label: "Number format",
            value: current.format,
            options: [{ value: "arabic", label: "1, 2, 3" }, { value: "roman", label: "i, ii, iii" }, { value: "Roman", label: "I, II, III" }, { value: "alpha", label: "a, b, c" }, { value: "Alpha", label: "A, B, C" }],
          },
          { kind: "number", id: "start", label: "Start at", value: current.startAt, min: 0, max: 32767, step: 1, hint: "Leave blank to number the pages from 1." },
          { kind: "checkbox", id: "first", label: "Show on first page", value: current.showOnFirstPage },
        ],
      },
    ],
    validate(values) {
      const start = values.start;
      if (start !== null && (typeof start !== "number" || !Number.isInteger(start) || start < 0 || start > 32767)) return { field: "start", message: "Type a whole number from 0 to 32767, or leave Start at blank." };
      return null;
    },
    onClosed: options.onClosed,
  });
  if (!result) return null;
  if (result.action === "remove") return { action: "remove" };
  const values = result.values;
  return {
    action: "insert",
    options: {
      position: values.position === "top" ? "top" : "bottom",
      align: (values.align as PageNumberAlign | null) ?? current.align,
      text: (values.text as PageNumberText | null) ?? current.text,
      format: (values.format as PageNumberFormat | null) ?? current.format,
      startAt: typeof values.start === "number" ? values.start : null,
      showOnFirstPage: values.first !== false,
    },
  };
}
