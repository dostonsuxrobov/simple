/**
 * Prepares imported blocks for engine-bridge `insertBlocks()` into an open document.
 * Inserted blocks may only reference lists and footnotes that already exist in the
 * target, so list paragraphs whose list is missing get their marker as text with a
 * hanging indent, and footnote references whose note is missing become a superscript
 * number with the note text appended after the inserted content.
 */
import type { Block, Document, ListDefinition, ListNumberFormat, Paragraph } from "./model.ts";

function roman(value: number): string {
  const numerals: Array<[number, string]> = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let rest = Math.max(1, Math.min(3999, value));
  let text = "";
  for (const [amount, numeral] of numerals) while (rest >= amount) { text += numeral; rest -= amount; }
  return text;
}

function letters(value: number): string {
  let rest = Math.max(1, value);
  let text = "";
  while (rest > 0) {
    rest -= 1;
    text = String.fromCharCode(97 + (rest % 26)) + text;
    rest = Math.floor(rest / 26);
  }
  return text;
}

export function formatListNumber(value: number, format: ListNumberFormat): string {
  switch (format) {
    case "lowerLetter": return letters(value);
    case "upperLetter": return letters(value).toUpperCase();
    case "lowerRoman": return roman(value);
    case "upperRoman": return roman(value).toUpperCase();
    default: return String(value);
  }
}

/** Marker text for each list paragraph, numbered the way the editor numbers them. */
function listMarkers(blocks: Block[], lists: Record<string, ListDefinition>) {
  const counters = new Map<string, number[]>();
  const markers = new Map<Paragraph, { text: string; indent: number; hanging: number }>();
  const visit = (items: Block[]) => {
    for (const block of items) {
      if (block.kind === "table") for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      if (block.kind !== "paragraph" || !block.style.list) continue;
      const definition = lists[block.style.list.listId];
      if (!definition) continue;
      const level = Math.min(Math.max(block.style.list.level, 0), definition.levels.length - 1);
      const levels = definition.levels;
      const counts = counters.get(definition.id) ?? [];
      counts[level] = (counts[level] ?? levels[level].start - 1) + 1;
      counts.length = level + 1;
      counters.set(definition.id, counts);
      const spec = levels[level];
      const text = spec.format === "bullet"
        ? spec.bulletChar || "•"
        : spec.text.replace(/%([1-9])/g, (_all, digit: string) => {
          const index = Number(digit) - 1;
          return formatListNumber(counts[index] ?? levels[index]?.start ?? 1, levels[index]?.format ?? "decimal");
        });
      markers.set(block, { text, indent: spec.indentLeftPx, hanging: spec.hangingPx });
    }
  };
  visit(blocks);
  return markers;
}

/**
 * Copies `source.blocks` so that every list and footnote they use exists in `target`.
 * Lists and notes the target already has (same id) are kept as they are.
 */
export function blocksForInsert(source: Document, target: Pick<Document, "lists" | "footnotes"> | null = null): Block[] {
  const blocks = structuredClone(source.blocks) as Block[];
  const targetLists = target?.lists ?? {};
  const targetNotes = target?.footnotes ?? {};
  const missingLists: Record<string, ListDefinition> = {};
  for (const [id, definition] of Object.entries(source.lists ?? {})) if (!targetLists[id]) missingLists[id] = definition;
  const markers = listMarkers(blocks, missingLists);
  for (const [paragraph, marker] of markers) {
    delete paragraph.style.list;
    paragraph.style.indentLeftPx = (paragraph.style.indentLeftPx ?? 0) + marker.indent;
    paragraph.style.indentFirstLinePx = -marker.hanging;
    const style = paragraph.runs[0]?.style ? { ...paragraph.runs[0].style } : undefined;
    if (style) {
      delete style.link;
      delete style.footnoteRef;
      paragraph.runs.unshift({ text: `${marker.text}\t`, style });
    }
  }
  const notes: Paragraph[] = [];
  const visit = (items: Block[]) => {
    for (const block of items) {
      if (block.kind === "table") for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      if (block.kind !== "paragraph") continue;
      for (const run of block.runs) {
        const id = run.style.footnoteRef;
        if (!id || targetNotes[id]) continue;
        delete run.style.footnoteRef;
        run.style.verticalAlign = "super";
        const body = structuredClone(source.footnotes?.[id] ?? []) as Paragraph[];
        body.forEach((paragraph, index) => {
          paragraph.id = `${paragraph.id}-note`;
          if (index === 0) paragraph.runs.unshift({ text: `${run.text} `, style: { ...(paragraph.runs[0]?.style ?? run.style), verticalAlign: "super" } });
          notes.push(paragraph);
        });
      }
    }
  };
  visit(blocks);
  return notes.length ? [...blocks, ...notes] : blocks;
}
