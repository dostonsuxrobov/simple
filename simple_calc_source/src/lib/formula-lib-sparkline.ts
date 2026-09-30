// Google Sheets' SPARKLINE(data, [options]): a miniature chart drawn inside the cell. The
// function returns a private-use marker string carrying the data and options; the grid draws
// it (sparkline-render.ts) and saving writes an empty cached result.

import { isEvaluationError } from "./formulas";
import { rectArg, spec, valueError } from "./formula-lib-shared";
import type { Specs } from "./formula-lib-shared";

/** Marker prefix of a sparkline result (U+E000 is private use, never typed by a user). */
export const SPARKLINE_PREFIX = "sparkline:";
const MAX_POINTS = 2000;

export interface SparklineSpec {
  data: Array<number | null>;
  options: Record<string, string | number | boolean>;
}

export function isSparklineValue(value: unknown): value is string {
  return typeof value === "string" && value.startsWith(SPARKLINE_PREFIX);
}

export function parseSparkline(value: unknown): SparklineSpec | null {
  if (!isSparklineValue(value)) return null;
  try {
    const parsed = JSON.parse(value.slice(SPARKLINE_PREFIX.length)) as { d?: unknown; o?: unknown };
    const data = Array.isArray(parsed.d) ? parsed.d.map((item) => (typeof item === "number" && Number.isFinite(item) ? item : null)) : [];
    const options = parsed.o && typeof parsed.o === "object" ? (parsed.o as SparklineSpec["options"]) : {};
    return { data, options };
  } catch {
    return null;
  }
}

export const SPARKLINE_FUNCTIONS: Specs = {
  SPARKLINE: spec(1, 2, (values) => {
    const data = rectArg(values[0]);
    if (isEvaluationError(data)) return data;
    // A single row or column of data (a 2-D range reads row by row, as in Google Sheets).
    const points: Array<number | null> = [];
    for (const item of data.values.slice(0, MAX_POINTS)) {
      if (isEvaluationError(item)) points.push(null);
      else if (typeof item === "number" && Number.isFinite(item)) points.push(item);
      else if (typeof item === "boolean") points.push(item ? 1 : 0);
      else points.push(null);
    }
    if (!points.length) return valueError();
    const options: SparklineSpec["options"] = {};
    if (values[1] !== undefined) {
      const table = rectArg(values[1]);
      if (isEvaluationError(table)) return table;
      if (table.columnCount < 2) return valueError();
      for (let row = 0; row < table.rowCount; row += 1) {
        const key = table.values[row * table.columnCount];
        const setting = table.values[row * table.columnCount + 1];
        if (typeof key !== "string" || isEvaluationError(setting) || setting === null || setting === undefined) continue;
        if (typeof setting === "string" || typeof setting === "number" || typeof setting === "boolean") options[key.trim().toLowerCase()] = setting;
      }
    }
    return `${SPARKLINE_PREFIX}${JSON.stringify({ d: points, o: options })}`;
  }),
};
