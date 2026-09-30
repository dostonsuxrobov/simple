// Text, regular-expression, and Google Sheets text functions for the extended library
// (see formula-library.ts).

import {
  FUNCTION_REGISTRY,
  isEvaluationError,
  isEvaluationRange,
  isLambdaValue,
  MAX_TEXT_RESULT_LENGTH,
  numberToText,
  safeTextResult,
  scalarArgument,
  toNumber,
  toText,
} from "./formulas";
import type { EvaluationError, FunctionSpec } from "./formulas";
import {
  arrayResult,
  booleanArg,
  naError,
  optionalBooleanArg,
  optionalIntegerArg,
  rectArg,
  spec,
  textArg,
  valueError,
} from "./formula-lib-shared";
import type { Scalar, Specs, Value } from "./formula-lib-shared";

// ---- Regular expressions ----------------------------------------------------------------

const MAX_PATTERN_LENGTH = 10_000;
const MAX_REGEX_INPUT = 1_000_000;
const MAX_MATCHES = 100_000;
const REGEX_CACHE = new Map<string, RegExp>();
const REGEX_CACHE_LIMIT = 500;

/**
 * Compile an Excel (PCRE2-flavoured) pattern to a JavaScript RegExp. A leading inline flag group
 * such as `(?i)` becomes a RegExp flag, and \A / \z anchors map to ^ / $. Patterns are tried in
 * Unicode mode first (for \p{L} and astral characters) and fall back to legacy mode, which also
 * accepts PCRE's permissive escapes like `\-`. Invalid patterns are #VALUE!.
 */
export function compileRegex(pattern: string, caseInsensitive: boolean, global = false): RegExp | EvaluationError {
  if (pattern.length > MAX_PATTERN_LENGTH) return valueError();
  const flags = new Set<string>();
  if (global) flags.add("g");
  if (caseInsensitive) flags.add("i");
  let source = pattern;
  const inline = /^\(\?([imsxn]+)\)/.exec(source);
  if (inline) {
    source = source.slice(inline[0].length);
    for (const flag of inline[1]) if (flag === "i" || flag === "m" || flag === "s") flags.add(flag);
  }
  source = source.replace(/(^|[^\\])\\A/g, "$1^").replace(/(^|[^\\])\\[zZ]/g, "$1$");
  const flagText = [...flags].sort().join("");
  const key = `${flagText}\u0000${source}`;
  const cached = REGEX_CACHE.get(key);
  if (cached) {
    cached.lastIndex = 0;
    return cached;
  }
  let regex: RegExp;
  try {
    regex = new RegExp(source, `${flagText}u`);
  } catch {
    try {
      regex = new RegExp(source, flagText);
    } catch {
      return valueError();
    }
  }
  if (REGEX_CACHE.size >= REGEX_CACHE_LIMIT) REGEX_CACHE.clear();
  REGEX_CACHE.set(key, regex);
  return regex;
}

/** case_sensitivity: 0 (default) = case-sensitive, 1 = case-insensitive. */
function caseInsensitiveArg(value: Value | undefined): boolean | EvaluationError {
  const mode = optionalIntegerArg(value, 0);
  if (isEvaluationError(mode)) return mode;
  if (mode !== 0 && mode !== 1) return valueError();
  return mode === 1;
}

function allMatches(regex: RegExp, text: string): RegExpExecArray[] | EvaluationError {
  const global = new RegExp(regex.source, regex.flags.includes("g") ? regex.flags : `${regex.flags}g`);
  const matches: RegExpExecArray[] = [];
  let match: RegExpExecArray | null;
  while ((match = global.exec(text)) !== null) {
    matches.push(match);
    if (matches.length > MAX_MATCHES) return valueError();
    if (match[0] === "") {
      // Step past empty matches (by code point in Unicode mode).
      const code = text.codePointAt(global.lastIndex);
      global.lastIndex += global.unicode && code !== undefined && code > 0xffff ? 2 : 1;
      if (global.lastIndex > text.length) break;
    }
  }
  return matches;
}

/** Expand a PCRE2-style replacement template: $n, ${n}, $0, ${name}, $<name>, $$. */
function expandReplacement(template: string, match: RegExpExecArray): string {
  return template.replace(/\$(?:\$|(\d+)|\{(\w+)\}|<(\w+)>)/g, (token, digits: string, braced: string, angled: string) => {
    if (token === "$$") return "$";
    const name = digits ?? braced ?? angled;
    if (/^\d+$/.test(name)) {
      const index = Number(name);
      if (index < match.length) return match[index] ?? "";
      return token;
    }
    return match.groups?.[name] ?? "";
  });
}

function textValue(value: Value | undefined): string | EvaluationError {
  const text = textArg(value);
  if (isEvaluationError(text)) return text;
  return text.length > MAX_REGEX_INPUT ? valueError() : text;
}

// ---- Delimited splitting ----------------------------------------------------------------

function delimiterList(value: Value | undefined): string[] | EvaluationError {
  if (value === undefined || value === null) return [];
  const range = rectArg(value);
  if (isEvaluationError(range)) return range;
  const delimiters: string[] = [];
  for (const entry of range.values) {
    const text = toText(entry);
    if (isEvaluationError(text)) return text;
    if (text !== "") delimiters.push(text);
  }
  return delimiters.sort((left, right) => right.length - left.length);
}

/** Split `text` at any of `delimiters` (longest first); matching may ignore case. */
function splitAt(text: string, delimiters: string[], ignoreCase: boolean): string[] {
  if (delimiters.length === 0) return [text];
  const haystack = ignoreCase ? text.toLocaleLowerCase() : text;
  const needles = ignoreCase ? delimiters.map((delimiter) => delimiter.toLocaleLowerCase()) : delimiters;
  const parts: string[] = [];
  let start = 0;
  let position = 0;
  while (position < text.length) {
    let matched = -1;
    for (let index = 0; index < needles.length; index += 1) {
      if (haystack.startsWith(needles[index], position)) {
        matched = index;
        break;
      }
    }
    if (matched >= 0) {
      parts.push(text.slice(start, position));
      position += needles[matched].length;
      start = position;
    } else {
      position += 1;
    }
  }
  parts.push(text.slice(start));
  return parts;
}

// ---- Value-to-text rendering ------------------------------------------------------------

function renderValue(value: Scalar, strict: boolean): string {
  if (value === null) return "";
  if (isEvaluationError(value)) return value.code;
  if (typeof value === "number") return numberToText(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return strict ? `"${value.replace(/"/g, '""')}"` : value;
}

function formatArg(value: Value | undefined): boolean | EvaluationError {
  const format = optionalIntegerArg(value, 0);
  if (isEvaluationError(format)) return format;
  if (format !== 0 && format !== 1) return valueError();
  return format === 1;
}

// ---- Width conversion (ASC / JIS) -------------------------------------------------------

function toHalfWidth(text: string): string {
  return text.replace(/[！-～　]/g, (character) =>
    character === "　" ? " " : String.fromCharCode(character.charCodeAt(0) - 0xfee0),
  );
}

function toFullWidth(text: string): string {
  return text.replace(/[!-~ ]/g, (character) =>
    character === " " ? "　" : String.fromCharCode(character.charCodeAt(0) + 0xfee0),
  );
}

function alias(name: string): FunctionSpec {
  const target = FUNCTION_REGISTRY[name];
  return { ...target };
}

/** Numeric value of Sheets TO_* conversion inputs; text that is not numeric passes through. */
function pureNumber(value: Scalar): Value {
  if (isEvaluationError(value)) return value;
  if (typeof value === "string") {
    const number = toNumber(value);
    return isEvaluationError(number) ? value : number;
  }
  if (typeof value === "boolean") return value ? 1 : 0;
  return value ?? 0;
}

/**
 * TEXTBEFORE/TEXTAFTER with Excel's full signature: (text, delimiter, [instance_num],
 * [match_mode], [match_end], [if_not_found]). `delimiter` may be an array of alternatives;
 * negative instances count from the end; match_end treats the end (or, searching backwards,
 * the start) of the text as a delimiter.
 */
function textBeforeAfter(name: "TEXTBEFORE" | "TEXTAFTER"): FunctionSpec {
  return spec(
    2,
    6,
    (values) => {
      const text = textArg(values[0]);
      if (isEvaluationError(text)) return text;
      const delimiterRange = rectArg(values[1]);
      if (isEvaluationError(delimiterRange)) return delimiterRange;
      const delimiters: string[] = [];
      for (const entry of delimiterRange.values) {
        const delimiter = toText(entry);
        if (isEvaluationError(delimiter)) return delimiter;
        delimiters.push(delimiter);
      }
      const instance = optionalIntegerArg(values[2], 1);
      if (isEvaluationError(instance)) return instance;
      const ignoreCase = caseInsensitiveArg(values[3]);
      if (isEvaluationError(ignoreCase)) return ignoreCase;
      const matchEnd = optionalBooleanArg(values[4], false);
      if (isEvaluationError(matchEnd)) return matchEnd;
      if (instance === 0 || Math.abs(instance) > text.length + 1) return valueError();
      const notFound = (): Value => (values[5] === undefined ? naError() : scalarArgument(values[5]));

      const nonEmpty = delimiters.filter((delimiter) => delimiter !== "").sort((left, right) => right.length - left.length);
      const positions: Array<{ start: number; end: number }> = [];
      if (nonEmpty.length === 0) {
        // An empty delimiter matches immediately at the search origin.
        positions.push(instance > 0 ? { start: 0, end: 0 } : { start: text.length, end: text.length });
      } else {
        const haystack = ignoreCase ? text.toLocaleLowerCase() : text;
        const needles = ignoreCase ? nonEmpty.map((delimiter) => delimiter.toLocaleLowerCase()) : nonEmpty;
        let position = 0;
        while (position < text.length) {
          const needle = needles.find((candidate) => haystack.startsWith(candidate, position));
          if (needle === undefined) {
            position += 1;
            continue;
          }
          positions.push({ start: position, end: position + needle.length });
          position += needle.length;
        }
        if (matchEnd) {
          if (instance > 0) positions.push({ start: text.length, end: text.length });
          else positions.unshift({ start: 0, end: 0 });
        }
      }
      const index = instance > 0 ? instance - 1 : positions.length + instance;
      if (index < 0 || index >= positions.length) return notFound();
      const found = positions[index];
      return name === "TEXTBEFORE" ? text.slice(0, found.start) : text.slice(found.end);
    },
    { liftArgs: [0] },
  );
}

// ---- Registry ---------------------------------------------------------------------------

export const TEXT_FUNCTIONS: Specs = {
  TEXTBEFORE: textBeforeAfter("TEXTBEFORE"),
  TEXTAFTER: textBeforeAfter("TEXTAFTER"),
  TEXTSPLIT: spec(
    2,
    6,
    (values) => {
      const text = textArg(values[0]);
      if (isEvaluationError(text)) return text;
      const columnDelimiters = delimiterList(values[1]);
      if (isEvaluationError(columnDelimiters)) return columnDelimiters;
      const rowDelimiters = delimiterList(values[2]);
      if (isEvaluationError(rowDelimiters)) return rowDelimiters;
      if (columnDelimiters.length === 0 && rowDelimiters.length === 0) return valueError();
      const ignoreEmpty = optionalBooleanArg(values[3], false);
      if (isEvaluationError(ignoreEmpty)) return ignoreEmpty;
      const ignoreCase = caseInsensitiveArg(values[4]);
      if (isEvaluationError(ignoreCase)) return ignoreCase;
      const pad: Scalar = values[5] === undefined ? naError() : scalarArgument(values[5]);
      let lines = splitAt(text, rowDelimiters, ignoreCase);
      if (ignoreEmpty) lines = lines.filter((line) => line !== "");
      const rows = lines.map((line) => {
        const parts = splitAt(line, columnDelimiters, ignoreCase);
        return ignoreEmpty ? parts.filter((part) => part !== "") : parts;
      });
      const kept = ignoreEmpty ? rows.filter((row) => row.length > 0) : rows;
      if (kept.length === 0) return { kind: "evaluationError", code: "#CALC!" };
      const width = Math.max(...kept.map((row) => row.length));
      if (kept.length * width > 100_000) return valueError();
      const output: Scalar[] = [];
      for (const row of kept) for (let column = 0; column < width; column += 1) output.push(column < row.length ? row[column] : pad);
      return arrayResult(output, kept.length, width);
    },
    { returnsArray: true, liftArgs: [0] },
  ),

  VALUETOTEXT: spec(
    1,
    2,
    (values) => {
      const value = scalarArgument(values[0]);
      const strict = formatArg(values[1]);
      if (isEvaluationError(strict)) return strict;
      return renderValue(value, strict);
    },
    { liftArgs: "all" },
  ),

  ARRAYTOTEXT: spec(1, 2, (values) => {
    const strict = formatArg(values[1]);
    if (isEvaluationError(strict)) return strict;
    const value = values[0];
    if (isLambdaValue(value)) return valueError();
    const range = isEvaluationRange(value) ? rectArg(value) : { values: [value], rowCount: 1, columnCount: 1 };
    if (isEvaluationError(range)) return range;
    if (!strict) return safeTextResult(range.values.map((entry) => renderValue(entry, false)).join(", "));
    const rows: string[] = [];
    for (let row = 0; row < range.rowCount; row += 1) {
      const cells = range.values.slice(row * range.columnCount, (row + 1) * range.columnCount);
      rows.push(cells.map((entry) => renderValue(entry, true)).join(","));
    }
    return safeTextResult(`{${rows.join(";")}}`);
  }),

  NUMBERVALUE: spec(
    1,
    3,
    (values) => {
      const raw = scalarArgument(values[0]);
      if (isEvaluationError(raw)) return raw;
      if (typeof raw === "number") return raw;
      const decimal = values[1] === undefined ? "." : textArg(values[1]);
      if (isEvaluationError(decimal)) return decimal;
      const group = values[2] === undefined ? "," : textArg(values[2]);
      if (isEvaluationError(group)) return group;
      if (decimal === "") return valueError();
      const decimalSeparator = Array.from(decimal)[0];
      const groupSeparator = group === "" ? "" : Array.from(group)[0];
      if (decimalSeparator === groupSeparator) return valueError();
      const text = toText(raw);
      if (isEvaluationError(text)) return text;
      let source = text.replace(/\s+/g, "");
      if (source === "") return 0;
      let percent = 0;
      while (source.endsWith("%")) {
        source = source.slice(0, -1);
        percent += 1;
      }
      const decimalAt = source.indexOf(decimalSeparator);
      if (decimalAt >= 0 && source.indexOf(decimalSeparator, decimalAt + decimalSeparator.length) >= 0) return valueError();
      let integerPart = decimalAt >= 0 ? source.slice(0, decimalAt) : source;
      let fractionPart = decimalAt >= 0 ? source.slice(decimalAt + decimalSeparator.length) : "";
      if (groupSeparator) {
        if (fractionPart.includes(groupSeparator)) return valueError();
        integerPart = integerPart.split(groupSeparator).join("");
      }
      const exponentMatch = /^(\d*)([eE][+-]?\d+)$/.exec(fractionPart);
      let exponent = "";
      if (exponentMatch) {
        fractionPart = exponentMatch[1];
        exponent = exponentMatch[2];
      } else if (decimalAt < 0) {
        const integerExponent = /^([+-]?\d*)([eE][+-]?\d+)$/.exec(integerPart);
        if (integerExponent) {
          integerPart = integerExponent[1];
          exponent = integerExponent[2];
        }
      }
      if (!/^[+-]?\d*$/.test(integerPart) || !/^\d*$/.test(fractionPart)) return valueError();
      if (!/\d/.test(integerPart + fractionPart)) return valueError();
      const number = Number(`${integerPart}.${fractionPart}${exponent}`.replace(/^([+-]?)\./, "$10."));
      if (!Number.isFinite(number)) return valueError();
      return number / 100 ** percent;
    },
    { liftArgs: "all" },
  ),

  UNICHAR: spec(
    1,
    1,
    (values) => {
      const number = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(number)) return number;
      const code = Math.trunc(number);
      if (code < 1 || code > 0x10ffff) return valueError();
      if (code >= 0xd800 && code <= 0xdfff) return naError();
      return String.fromCodePoint(code);
    },
    { liftArgs: "all" },
  ),

  UNICODE: spec(
    1,
    1,
    (values) => {
      const text = textArg(values[0]);
      if (isEvaluationError(text)) return text;
      const code = text.codePointAt(0);
      return code === undefined ? valueError() : code;
    },
    { liftArgs: "all" },
  ),

  T: spec(
    1,
    1,
    (values) => {
      const value = scalarArgument(values[0]);
      if (isEvaluationError(value)) return value;
      return typeof value === "string" ? value : "";
    },
    { liftArgs: "all" },
  ),

  // Double-byte ("B") variants: character based, as Excel behaves without a DBCS default language.
  LEFTB: alias("LEFT"),
  RIGHTB: alias("RIGHT"),
  MIDB: alias("MID"),
  LENB: alias("LEN"),
  FINDB: alias("FIND"),
  SEARCHB: alias("SEARCH"),
  REPLACEB: alias("REPLACE"),

  ASC: spec(
    1,
    1,
    (values) => {
      const text = textArg(values[0]);
      return isEvaluationError(text) ? text : toHalfWidth(text);
    },
    { liftArgs: "all" },
  ),
  JIS: spec(
    1,
    1,
    (values) => {
      const text = textArg(values[0]);
      return isEvaluationError(text) ? text : toFullWidth(text);
    },
    { liftArgs: "all" },
  ),
  DBCS: spec(
    1,
    1,
    (values) => {
      const text = textArg(values[0]);
      return isEvaluationError(text) ? text : toFullWidth(text);
    },
    { liftArgs: "all" },
  ),
  PHONETIC: spec(1, 1, (values) => {
    const value = values[0];
    const first = isEvaluationRange(value) ? (value.values[0] ?? null) : value;
    const text = textArg(isLambdaValue(first) ? valueError() : first);
    return text;
  }),

  REGEXTEST: spec(
    2,
    3,
    (values) => {
      const text = textValue(values[0]);
      if (isEvaluationError(text)) return text;
      const pattern = textArg(values[1]);
      if (isEvaluationError(pattern)) return pattern;
      const ignoreCase = caseInsensitiveArg(values[2]);
      if (isEvaluationError(ignoreCase)) return ignoreCase;
      const regex = compileRegex(pattern, ignoreCase);
      if (isEvaluationError(regex)) return regex;
      return regex.test(text);
    },
    { liftArgs: "all" },
  ),

  REGEXEXTRACT: spec(
    2,
    4,
    (values) => {
      const text = textValue(values[0]);
      if (isEvaluationError(text)) return text;
      const pattern = textArg(values[1]);
      if (isEvaluationError(pattern)) return pattern;
      const mode = optionalIntegerArg(values[2], 0);
      if (isEvaluationError(mode)) return mode;
      if (mode < 0 || mode > 2) return valueError();
      const ignoreCase = caseInsensitiveArg(values[3]);
      if (isEvaluationError(ignoreCase)) return ignoreCase;
      const regex = compileRegex(pattern, ignoreCase);
      if (isEvaluationError(regex)) return regex;
      if (mode === 1) {
        const matches = allMatches(regex, text);
        if (isEvaluationError(matches)) return matches;
        if (matches.length === 0) return naError();
        return arrayResult(matches.map((match) => match[0]), 1, matches.length);
      }
      const match = regex.exec(text);
      if (!match) return naError();
      if (mode === 0) return match[0];
      const groups = match.slice(1).map((group) => group ?? "");
      if (groups.length === 0) return naError();
      return arrayResult(groups, 1, groups.length);
    },
    { liftArgs: [0], returnsArray: true },
  ),

  REGEXREPLACE: spec(
    3,
    5,
    (values) => {
      const text = textValue(values[0]);
      if (isEvaluationError(text)) return text;
      const pattern = textArg(values[1]);
      if (isEvaluationError(pattern)) return pattern;
      const replacement = textArg(values[2]);
      if (isEvaluationError(replacement)) return replacement;
      const occurrence = optionalIntegerArg(values[3], 0);
      if (isEvaluationError(occurrence)) return occurrence;
      const ignoreCase = caseInsensitiveArg(values[4]);
      if (isEvaluationError(ignoreCase)) return ignoreCase;
      const regex = compileRegex(pattern, ignoreCase);
      if (isEvaluationError(regex)) return regex;
      const matches = allMatches(regex, text);
      if (isEvaluationError(matches)) return matches;
      let targets: RegExpExecArray[];
      if (occurrence === 0) targets = matches;
      else {
        const index = occurrence > 0 ? occurrence - 1 : matches.length + occurrence;
        targets = index >= 0 && index < matches.length ? [matches[index]] : [];
      }
      let output = "";
      let cursor = 0;
      for (const match of targets) {
        output += text.slice(cursor, match.index) + expandReplacement(replacement, match);
        cursor = match.index + match[0].length;
        if (output.length > MAX_TEXT_RESULT_LENGTH) return valueError();
      }
      return safeTextResult(output + text.slice(cursor));
    },
    { liftArgs: "all" },
  ),

  REGEXMATCH: spec(
    2,
    2,
    (values) => {
      const text = textValue(values[0]);
      if (isEvaluationError(text)) return text;
      const pattern = textArg(values[1]);
      if (isEvaluationError(pattern)) return pattern;
      const regex = compileRegex(pattern, false);
      if (isEvaluationError(regex)) return regex;
      return regex.test(text);
    },
    { liftArgs: "all" },
  ),

  SPLIT: spec(
    2,
    4,
    (values) => {
      const text = textArg(values[0]);
      if (isEvaluationError(text)) return text;
      const delimiter = textArg(values[1]);
      if (isEvaluationError(delimiter)) return delimiter;
      const splitByEach = booleanArg(values[2], true);
      if (isEvaluationError(splitByEach)) return splitByEach;
      const removeEmpty = booleanArg(values[3], true);
      if (isEvaluationError(removeEmpty)) return removeEmpty;
      if (delimiter === "") return valueError();
      const delimiters = splitByEach ? Array.from(delimiter) : [delimiter];
      let parts = splitAt(text, delimiters, false);
      if (removeEmpty) parts = parts.filter((part) => part !== "");
      if (parts.length === 0) return "";
      const output: Scalar[] = parts.map((part) => {
        const trimmed = part.trim();
        const number = trimmed === "" ? NaN : Number(trimmed);
        return Number.isFinite(number) ? number : part;
      });
      return arrayResult(output, 1, output.length);
    },
    { returnsArray: true },
  ),

  ENCODEURL: spec(
    1,
    1,
    (values) => {
      const text = textArg(values[0]);
      if (isEvaluationError(text)) return text;
      try {
        return safeTextResult(encodeURIComponent(text));
      } catch {
        return valueError();
      }
    },
    { liftArgs: "all" },
  ),

  TO_TEXT: spec(
    1,
    1,
    (values) => {
      const value = scalarArgument(values[0]);
      return isEvaluationError(value) ? value : renderValue(value, false);
    },
    { liftArgs: "all" },
  ),
  TO_PURE_NUMBER: spec(1, 1, (values) => pureNumber(scalarArgument(values[0])), { liftArgs: "all" }),
  TO_PERCENT: spec(1, 1, (values) => pureNumber(scalarArgument(values[0])), { liftArgs: "all" }),
  TO_DOLLARS: spec(1, 1, (values) => pureNumber(scalarArgument(values[0])), { liftArgs: "all" }),
};
