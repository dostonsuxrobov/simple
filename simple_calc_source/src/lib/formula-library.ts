// Extended worksheet function library. Importing this module registers several hundred
// additional Excel 365 (and common Google Sheets) functions with the formula engine.
// Each domain lives in its own formula-lib-*.ts module; this file only merges and registers
// them, never replacing a function the core engine already defines unless it is listed in
// LIBRARY_OVERRIDES.

import { hasFormulaFunction, registerFormulaFunctions } from "./formulas";
import type { FunctionSpec } from "./formulas";
import { ARRAY_FUNCTIONS } from "./formula-lib-array";
import { DATE_FUNCTIONS } from "./formula-lib-date";
import { ENGINEERING_FUNCTIONS } from "./formula-lib-engineering";
import { FINANCIAL_FUNCTIONS } from "./formula-lib-financial";
import { INFO_FUNCTIONS } from "./formula-lib-info";
import { MATH_FUNCTIONS } from "./formula-lib-math";
import { STATS_FUNCTIONS } from "./formula-lib-stats";
import { TEXT_FUNCTIONS } from "./formula-lib-text";
import { SPARKLINE_FUNCTIONS } from "./formula-lib-sparkline";

/** Domain modules in registration order (exported for the QA script's duplicate check). */
export const LIBRARY_MODULES: Record<string, Record<string, FunctionSpec>> = {
  array: ARRAY_FUNCTIONS,
  date: DATE_FUNCTIONS,
  engineering: ENGINEERING_FUNCTIONS,
  financial: FINANCIAL_FUNCTIONS,
  info: INFO_FUNCTIONS,
  math: MATH_FUNCTIONS,
  stats: STATS_FUNCTIONS,
  text: TEXT_FUNCTIONS,
  sparkline: SPARKLINE_FUNCTIONS,
};

/**
 * Core functions the library intentionally replaces with fuller implementations:
 * TEXTBEFORE/TEXTAFTER gain Excel's match_mode, match_end, if_not_found and delimiter arrays.
 */
const LIBRARY_OVERRIDES = new Set<string>(["TEXTBEFORE", "TEXTAFTER"]);

const additions: Record<string, FunctionSpec> = {};
for (const specs of Object.values(LIBRARY_MODULES)) {
  for (const [name, spec] of Object.entries(specs)) {
    const key = name.toUpperCase();
    if (hasFormulaFunction(key) && !LIBRARY_OVERRIDES.has(key)) continue;
    additions[key] = spec;
  }
}
registerFormulaFunctions(additions);

/** Names this module added to the engine. */
export const LIBRARY_FUNCTION_NAMES: readonly string[] = Object.keys(additions).sort();
