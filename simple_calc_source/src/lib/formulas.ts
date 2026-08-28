import { SSF } from "xlsx";

/** Values a cell resolver may return to the formula engine. */
export type FormulaPrimitive = number | string | boolean | Date | null | undefined;

/** Spreadsheet errors returned by the evaluator. */
export type FormulaError =
  | "#NULL!"
  | "#DIV/0!"
  | "#VALUE!"
  | "#REF!"
  | "#NAME?"
  | "#NUM!"
  | "#N/A"
  | "#CIRC!"
  | "#PARSE!"
  | "#SPILL!"
  | "#CALC!";

/** A resolved cell can provide a literal, a formula, or an explicit error. */
export interface ResolvedFormulaCell {
  value?: FormulaPrimitive;
  formula?: string;
  error?: FormulaError;
}

export type FormulaResolverResult = FormulaPrimitive | ResolvedFormulaCell;

/** Resolve an A1 address (without `$` anchors) on a sheet. */
export type FormulaResolver = (
  sheetId: string,
  address: string,
) => FormulaResolverResult;

/** Inclusive one-based bounds of a rectangular sheet region. */
export interface FormulaRangeBounds {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

/**
 * Optional host integration for the evaluator. `getUsedRange` clamps
 * whole-column/row references (`A:A`, `1:1`) to the sheet's used range;
 * without it they are capped at 10,000 rows by 256 columns.
 * `forEachCellInRange` visits only populated cells so aggregates over large
 * sparse ranges avoid materializing rectangles; without it ranges beyond
 * 100,000 cells evaluate to #VALUE!. `resolveDefinedName` maps a workbook
 * name to a reference/range string (interpreted as A1) or a literal value;
 * without it unknown identifiers evaluate to #NAME?. `currentCell` backs
 * zero-argument ROW()/COLUMN(). `isFormulaCell` reports whether a cell holds
 * a formula, for hosts whose resolver returns computed values rather than
 * formula text; without it ISFORMULA falls back to inspecting the raw
 * resolver result.
 */
export interface FormulaEvaluationHooks {
  getUsedRange?: (
    sheetId: string,
  ) => { maxRow: number; maxCol: number } | null | undefined;
  forEachCellInRange?: (
    sheetId: string,
    bounds: FormulaRangeBounds,
    visit: (row: number, column: number) => void,
  ) => void;
  resolveDefinedName?: (name: string) => FormulaPrimitive;
  currentCell?: { row: number; column: number };
  isFormulaCell?: (sheetId: string, address: string) => boolean;
}

/** Public scalar result. Errors are returned as spreadsheet-style strings. */
export type FormulaResult = number | string | boolean;

/** Parsed, one-based A1 coordinates. */
export interface A1Address {
  row: number;
  column: number;
  rowAbsolute: boolean;
  columnAbsolute: boolean;
}

const MAX_EXCEL_ROW = 1_048_576;
const MAX_EXCEL_COLUMN = 16_384;
const MAX_RANGE_CELLS = 100_000;
const MAX_FORMULA_LENGTH = 100_000;
const MAX_CALCULATION_DEPTH = 256;
const MAX_TEXT_RESULT_LENGTH = 1_000_000;
const MAX_WILDCARD_STEPS = 5_000_000;
const MILLISECONDS_PER_DAY = 86_400_000;
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const MAX_DATE_SERIAL = 2_958_465;
const FALLBACK_USED_RANGE_ROWS = 10_000;
const FALLBACK_USED_RANGE_COLUMNS = 256;

const FORMULA_ERRORS = new Set<FormulaError>([
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#CIRC!",
  "#PARSE!",
  "#SPILL!",
  "#CALC!",
]);

/** Return whether a value is one of the engine's spreadsheet errors. */
export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === "string" && FORMULA_ERRORS.has(value as FormulaError);
}

/** Convert an Excel column label (`A`, `XFD`) to a one-based index. */
export function columnLabelToNumber(label: string): number | null {
  if (!/^[A-Za-z]{1,3}$/.test(label)) return null;

  let column = 0;
  for (const character of label.toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return column <= MAX_EXCEL_COLUMN ? column : null;
}

/** Convert a one-based column index to its Excel column label. */
export function columnNumberToLabel(column: number): string | null {
  if (!Number.isInteger(column) || column < 1 || column > MAX_EXCEL_COLUMN) {
    return null;
  }

  let value = column;
  let label = "";
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

/** Parse an A1 address, including optional absolute anchors. */
export function parseA1Address(address: string): A1Address | null {
  const match = /^\s*(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)\s*$/.exec(address);
  if (!match) return null;

  const column = columnLabelToNumber(match[2]);
  const row = Number(match[4]);
  if (column === null || !Number.isSafeInteger(row) || row > MAX_EXCEL_ROW) {
    return null;
  }

  return {
    row,
    column,
    columnAbsolute: match[1] === "$",
    rowAbsolute: match[3] === "$",
  };
}

/** Format one-based coordinates as an A1 address. */
export function formatA1Address(address: A1Address): string | null {
  if (
    !Number.isInteger(address.row) ||
    address.row < 1 ||
    address.row > MAX_EXCEL_ROW
  ) {
    return null;
  }

  const column = columnNumberToLabel(address.column);
  if (column === null) return null;
  return `${address.columnAbsolute ? "$" : ""}${column}${
    address.rowAbsolute ? "$" : ""
  }${address.row}`;
}

type TokenKind =
  | "number"
  | "string"
  | "identifier"
  | "cell"
  | "columnRange"
  | "rowRange"
  | "sheet"
  | "error"
  | "operator"
  | "leftParen"
  | "rightParen"
  | "leftBrace"
  | "rightBrace"
  | "comma"
  | "semicolon"
  | "colon"
  | "bang"
  | "eof";

interface Token {
  kind: TokenKind;
  text: string;
  value?: string | number;
  position: number;
}

class FormulaParseError extends Error {
  constructor(
    readonly formulaError: FormulaError = "#PARSE!",
    message = "Invalid formula",
  ) {
    super(message);
  }
}

function nextNonWhitespace(source: string, start: number): string {
  let position = start;
  while (/\s/.test(source[position] ?? "")) position += 1;
  return source[position] ?? "";
}

function tokenize(source: string): Token[] {
  if (source.length > MAX_FORMULA_LENGTH) throw new FormulaParseError();

  const tokens: Token[] = [];
  let position = 0;

  while (position < source.length) {
    const character = source[position];
    if (/\s/.test(character)) {
      position += 1;
      continue;
    }

    if (character === '"') {
      const start = position;
      position += 1;
      let value = "";
      let closed = false;
      while (position < source.length) {
        if (source[position] !== '"') {
          value += source[position];
          position += 1;
        } else if (source[position + 1] === '"') {
          value += '"';
          position += 2;
        } else {
          position += 1;
          closed = true;
          break;
        }
      }
      if (!closed) throw new FormulaParseError();
      tokens.push({ kind: "string", text: source.slice(start, position), value, position: start });
      continue;
    }

    if (character === "'") {
      const start = position;
      position += 1;
      let value = "";
      let closed = false;
      while (position < source.length) {
        if (source[position] !== "'") {
          value += source[position];
          position += 1;
        } else if (source[position + 1] === "'") {
          value += "'";
          position += 2;
        } else {
          position += 1;
          closed = true;
          break;
        }
      }
      if (!closed) throw new FormulaParseError();
      tokens.push({ kind: "sheet", text: source.slice(start, position), value, position: start });
      continue;
    }

    const rowRangeMatch = /^\$?[1-9]\d*:\$?[1-9]\d*/.exec(source.slice(position));
    if (rowRangeMatch) {
      const end = position + rowRangeMatch[0].length;
      if (!/[A-Za-z0-9_.$]/.test(source[end] ?? "")) {
        tokens.push({ kind: "rowRange", text: rowRangeMatch[0], value: rowRangeMatch[0], position });
        position = end;
        continue;
      }
    }

    const numberMatch = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(
      source.slice(position),
    );
    if (numberMatch) {
      const value = Number(numberMatch[0]);
      if (!Number.isFinite(value)) throw new FormulaParseError("#NUM!");
      tokens.push({ kind: "number", text: numberMatch[0], value, position });
      position += numberMatch[0].length;
      continue;
    }

    if (character === "#") {
      const errorMatch = /^#[A-Za-z0-9/]+[!?]?/.exec(source.slice(position));
      const text = errorMatch?.[0].toUpperCase();
      if (!text || !isFormulaError(text)) throw new FormulaParseError();
      tokens.push({ kind: "error", text, value: text, position });
      position += errorMatch![0].length;
      continue;
    }

    const columnRangeMatch = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/.exec(
      source.slice(position),
    );
    if (columnRangeMatch) {
      const end = position + columnRangeMatch[0].length;
      if (!/[A-Za-z0-9_.$]/.test(source[end] ?? "")) {
        tokens.push({ kind: "columnRange", text: columnRangeMatch[0], value: columnRangeMatch[0], position });
        position = end;
        continue;
      }
    }

    const cellMatch = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*/.exec(source.slice(position));
    if (cellMatch) {
      const end = position + cellMatch[0].length;
      const following = source[end] ?? "";
      const isComplete = !/[A-Za-z0-9_.]/.test(following);
      const looksLikeFunction =
        !cellMatch[0].includes("$") && nextNonWhitespace(source, end) === "(";
      if (isComplete && !looksLikeFunction) {
        tokens.push({ kind: "cell", text: cellMatch[0], value: cellMatch[0], position });
        position = end;
        continue;
      }
    }

    const identifierMatch = /^[A-Za-z_\\][A-Za-z0-9_.]*/.exec(source.slice(position));
    if (identifierMatch) {
      tokens.push({
        kind: "identifier",
        text: identifierMatch[0],
        value: identifierMatch[0],
        position,
      });
      position += identifierMatch[0].length;
      continue;
    }

    const twoCharacterOperator = source.slice(position, position + 2);
    if (["<=", ">=", "<>"].includes(twoCharacterOperator)) {
      tokens.push({ kind: "operator", text: twoCharacterOperator, position });
      position += 2;
      continue;
    }
    if (["+", "-", "*", "/", "^", "&", "=", "<", ">", "%"].includes(character)) {
      tokens.push({ kind: "operator", text: character, position });
      position += 1;
      continue;
    }

    const punctuation: Partial<Record<string, TokenKind>> = {
      "(": "leftParen",
      ")": "rightParen",
      "{": "leftBrace",
      "}": "rightBrace",
      ",": "comma",
      ";": "semicolon",
      ":": "colon",
      "!": "bang",
    };
    const kind = punctuation[character];
    if (kind) {
      tokens.push({ kind, text: character, position });
      position += 1;
      continue;
    }

    throw new FormulaParseError();
  }

  tokens.push({ kind: "eof", text: "", position: source.length });
  return tokens;
}

interface LiteralNode {
  kind: "literal";
  value: number | string | boolean;
}

interface ErrorNode {
  kind: "error";
  value: FormulaError;
}

interface ReferenceNode {
  kind: "reference";
  sheet?: string;
  address: A1Address;
}

interface RangeNode {
  kind: "range";
  start: ReferenceNode;
  end: ReferenceNode;
}

interface WholeRangeNode {
  kind: "wholeRange";
  sheet?: string;
  axis: "column" | "row";
  start: number;
  end: number;
}

interface ArrayNode {
  kind: "array";
  values: EvaluationScalar[];
  rowCount: number;
  columnCount: number;
}

interface NameNode {
  kind: "name";
  name: string;
}

interface UnaryNode {
  kind: "unary";
  operator: "+" | "-" | "%";
  operand: FormulaNode;
}

interface BinaryNode {
  kind: "binary";
  operator: string;
  left: FormulaNode;
  right: FormulaNode;
}

interface CallNode {
  kind: "call";
  name: string;
  arguments: FormulaNode[];
}

type FormulaNode =
  | LiteralNode
  | ErrorNode
  | ReferenceNode
  | RangeNode
  | WholeRangeNode
  | ArrayNode
  | NameNode
  | UnaryNode
  | BinaryNode
  | CallNode;

class FormulaParser {
  private position = 0;

  constructor(private readonly tokens: Token[]) {}

  parse(): FormulaNode {
    const expression = this.parseComparison();
    if (this.current.kind !== "eof") throw new FormulaParseError();
    return expression;
  }

  private get current(): Token {
    return this.tokens[this.position];
  }

  private get next(): Token {
    return this.tokens[this.position + 1] ?? this.tokens[this.tokens.length - 1];
  }

  private advance(): Token {
    const token = this.current;
    this.position += 1;
    return token;
  }

  private match(kind: TokenKind, text?: string): boolean {
    if (this.current.kind !== kind || (text !== undefined && this.current.text !== text)) {
      return false;
    }
    this.position += 1;
    return true;
  }

  private parseComparison(): FormulaNode {
    let node = this.parseConcatenation();
    while (
      this.current.kind === "operator" &&
      ["=", "<>", "<", ">", "<=", ">="].includes(this.current.text)
    ) {
      const operator = this.advance().text;
      node = { kind: "binary", operator, left: node, right: this.parseConcatenation() };
    }
    return node;
  }

  private parseConcatenation(): FormulaNode {
    let node = this.parseAdditive();
    while (this.match("operator", "&")) {
      node = { kind: "binary", operator: "&", left: node, right: this.parseAdditive() };
    }
    return node;
  }

  private parseAdditive(): FormulaNode {
    let node = this.parseMultiplicative();
    while (this.current.kind === "operator" && ["+", "-"].includes(this.current.text)) {
      const operator = this.advance().text;
      node = { kind: "binary", operator, left: node, right: this.parseMultiplicative() };
    }
    return node;
  }

  private parseMultiplicative(): FormulaNode {
    let node = this.parseUnary();
    while (this.current.kind === "operator" && ["*", "/"].includes(this.current.text)) {
      const operator = this.advance().text;
      node = { kind: "binary", operator, left: node, right: this.parseUnary() };
    }
    return node;
  }

  private parseUnary(): FormulaNode {
    if (this.current.kind === "operator" && ["+", "-"].includes(this.current.text)) {
      const operator = this.advance().text as "+" | "-";
      return { kind: "unary", operator, operand: this.parseUnary() };
    }
    return this.parsePower();
  }

  private parsePower(): FormulaNode {
    let node = this.parsePostfix();
    if (this.match("operator", "^")) {
      node = { kind: "binary", operator: "^", left: node, right: this.parseUnary() };
    }
    return node;
  }

  private parsePostfix(): FormulaNode {
    let node = this.parsePrimary();
    while (this.match("operator", "%")) {
      node = { kind: "unary", operator: "%", operand: node };
    }
    return node;
  }

  private parsePrimary(): FormulaNode {
    if (this.current.kind === "number") {
      return { kind: "literal", value: this.advance().value as number };
    }
    if (this.current.kind === "string") {
      return { kind: "literal", value: this.advance().value as string };
    }
    if (this.current.kind === "error") {
      return { kind: "error", value: this.advance().value as FormulaError };
    }

    if (this.current.kind === "identifier" && this.next.kind === "leftParen") {
      return this.parseCall();
    }

    if (this.current.kind === "leftBrace") {
      return this.parseArray();
    }

    const reference = this.parseReference();
    if (reference) {
      if (reference.kind === "wholeRange") return reference;
      if (this.match("colon")) {
        const end = this.parseReference();
        if (!end || end.kind !== "reference") throw new FormulaParseError("#REF!");
        return { kind: "range", start: reference, end };
      }
      return reference;
    }

    if (this.current.kind === "identifier") {
      const text = String(this.advance().value);
      const name = text.toUpperCase();
      if (name === "TRUE" || name === "FALSE") {
        return { kind: "literal", value: name === "TRUE" };
      }
      return { kind: "name", name: text };
    }

    if (this.match("leftParen")) {
      const expression = this.parseComparison();
      if (!this.match("rightParen")) throw new FormulaParseError();
      return expression;
    }

    throw new FormulaParseError();
  }

  private parseCall(): CallNode {
    const name = String(this.advance().value);
    if (!this.match("leftParen")) throw new FormulaParseError();
    const args: FormulaNode[] = [];
    if (!this.match("rightParen")) {
      do {
        args.push(this.parseComparison());
      } while (this.match("comma") || this.match("semicolon"));
      if (!this.match("rightParen")) throw new FormulaParseError();
    }
    return { kind: "call", name, arguments: args };
  }

  private parseArray(): ArrayNode {
    this.advance();
    const rows: EvaluationScalar[][] = [[this.parseArrayElement()]];
    while (!this.match("rightBrace")) {
      if (this.match("comma")) {
        rows[rows.length - 1].push(this.parseArrayElement());
      } else if (this.match("semicolon")) {
        rows.push([this.parseArrayElement()]);
      } else {
        throw new FormulaParseError();
      }
    }
    const columnCount = rows[0].length;
    if (rows.some((row) => row.length !== columnCount)) {
      throw new FormulaParseError("#VALUE!");
    }
    const values: EvaluationScalar[] = [];
    for (const row of rows) values.push(...row);
    return { kind: "array", values, rowCount: rows.length, columnCount };
  }

  private parseArrayElement(): EvaluationScalar {
    let negative = false;
    while (this.current.kind === "operator" && ["+", "-"].includes(this.current.text)) {
      if (this.advance().text === "-") negative = !negative;
    }
    if (this.current.kind === "number") {
      const value = this.advance().value as number;
      return negative ? -value : value;
    }
    if (negative) throw new FormulaParseError();
    if (this.current.kind === "string") {
      return this.advance().value as string;
    }
    if (this.current.kind === "error") {
      return evaluationError(this.advance().value as FormulaError);
    }
    if (this.current.kind === "identifier") {
      const name = String(this.current.value).toUpperCase();
      if (name === "TRUE" || name === "FALSE") {
        this.advance();
        return name === "TRUE";
      }
    }
    throw new FormulaParseError();
  }

  private parseReference(): ReferenceNode | WholeRangeNode | null {
    let sheet: string | undefined;
    if (
      ["sheet", "identifier", "cell"].includes(this.current.kind) &&
      this.next.kind === "bang"
    ) {
      sheet = String(this.advance().value);
      this.advance();
    }

    if (this.current.kind === "columnRange" || this.current.kind === "rowRange") {
      return this.parseWholeRange(sheet);
    }

    if (this.current.kind !== "cell") {
      if (sheet !== undefined) throw new FormulaParseError("#REF!");
      return null;
    }

    const address = parseA1Address(String(this.advance().value));
    if (!address) throw new FormulaParseError("#REF!");
    return { kind: "reference", sheet, address };
  }

  private parseWholeRange(sheet: string | undefined): WholeRangeNode {
    const token = this.advance();
    const [firstText, secondText] = token.text.split(":");
    if (token.kind === "columnRange") {
      const first = columnLabelToNumber(firstText.replace("$", ""));
      const second = columnLabelToNumber(secondText.replace("$", ""));
      if (first === null || second === null) throw new FormulaParseError("#REF!");
      return {
        kind: "wholeRange",
        sheet,
        axis: "column",
        start: Math.min(first, second),
        end: Math.max(first, second),
      };
    }
    const first = Number(firstText.replace("$", ""));
    const second = Number(secondText.replace("$", ""));
    if (
      !Number.isSafeInteger(first) ||
      !Number.isSafeInteger(second) ||
      first > MAX_EXCEL_ROW ||
      second > MAX_EXCEL_ROW
    ) {
      throw new FormulaParseError("#REF!");
    }
    return {
      kind: "wholeRange",
      sheet,
      axis: "row",
      start: Math.min(first, second),
      end: Math.max(first, second),
    };
  }
}

interface EvaluationError {
  kind: "evaluationError";
  code: FormulaError;
}

type EvaluationScalar = number | string | boolean | null | EvaluationError;

interface EvaluationRange {
  kind: "evaluationRange";
  values: EvaluationScalar[];
  rowCount: number;
  columnCount: number;
  /** Sparse ranges hold only populated cells; positional lookups are unavailable. */
  sparse?: boolean;
}

type EvaluationValue = EvaluationScalar | EvaluationRange;

interface EvaluationContext {
  resolver: FormulaResolver;
  hooks: FormulaEvaluationHooks;
  visiting: Set<string>;
  memo: Map<string, EvaluationScalar>;
  now: Date;
  calculationDepth: number;
}

function evaluationError(code: FormulaError): EvaluationError {
  return { kind: "evaluationError", code };
}

function isEvaluationError(value: unknown): value is EvaluationError {
  return (
    typeof value === "object" &&
    value !== null &&
    "kind" in value &&
    value.kind === "evaluationError"
  );
}

function isEvaluationRange(value: EvaluationValue): value is EvaluationRange {
  return typeof value === "object" && value !== null && value.kind === "evaluationRange";
}

function localDateToExcelSerial(date: Date): number {
  const day =
    (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - EXCEL_EPOCH_UTC) /
    MILLISECONDS_PER_DAY;
  const fraction =
    (date.getHours() * 3_600_000 +
      date.getMinutes() * 60_000 +
      date.getSeconds() * 1_000 +
      date.getMilliseconds()) /
    MILLISECONDS_PER_DAY;
  return day + fraction;
}

function normalizePrimitive(value: FormulaPrimitive): EvaluationScalar {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) {
    const serial = localDateToExcelSerial(value);
    return Number.isFinite(serial) ? serial : evaluationError("#VALUE!");
  }
  if (typeof value === "number" && !Number.isFinite(value)) {
    return evaluationError("#NUM!");
  }
  return typeof value === "number" || typeof value === "string" || typeof value === "boolean"
    ? value
    : evaluationError("#VALUE!");
}

function toNumber(value: EvaluationValue): number | EvaluationError {
  if (isEvaluationError(value)) return value;
  if (isEvaluationRange(value)) return evaluationError("#VALUE!");
  if (value === null) return 0;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  const trimmed = value.trim();
  if (trimmed === "") return 0;
  const number = Number(trimmed);
  return Number.isFinite(number) ? number : evaluationError("#VALUE!");
}

function toBoolean(value: EvaluationValue): boolean | EvaluationError {
  if (isEvaluationError(value)) return value;
  if (isEvaluationRange(value)) return evaluationError("#VALUE!");
  if (value === null) return false;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (value.trim().toUpperCase() === "TRUE") return true;
  if (value.trim().toUpperCase() === "FALSE") return false;
  const numeric = Number(value);
  return value.trim() !== "" && Number.isFinite(numeric)
    ? numeric !== 0
    : evaluationError("#VALUE!");
}

function toText(value: EvaluationValue): string | EvaluationError {
  if (isEvaluationError(value)) return value;
  if (isEvaluationRange(value)) return evaluationError("#VALUE!");
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function compareValues(left: EvaluationScalar, right: EvaluationScalar): number | EvaluationError {
  if (isEvaluationError(left)) return left;
  if (isEvaluationError(right)) return right;

  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  if (!isEvaluationError(leftNumber) && !isEvaluationError(rightNumber)) {
    return leftNumber === rightNumber ? 0 : leftNumber < rightNumber ? -1 : 1;
  }

  const leftText = toText(left);
  const rightText = toText(right);
  if (isEvaluationError(leftText)) return leftText;
  if (isEvaluationError(rightText)) return rightText;
  const normalizedLeft = leftText.toLowerCase();
  const normalizedRight = rightText.toLowerCase();
  return normalizedLeft === normalizedRight ? 0 : normalizedLeft < normalizedRight ? -1 : 1;
}

function applyBinaryOperator(
  operator: string,
  left: EvaluationValue,
  right: EvaluationValue,
): EvaluationValue {
  if (isEvaluationError(left)) return left;
  if (isEvaluationError(right)) return right;
  if (isEvaluationRange(left) || isEvaluationRange(right)) return evaluationError("#VALUE!");

  if (operator === "&") {
    const leftText = toText(left);
    const rightText = toText(right);
    if (isEvaluationError(leftText)) return leftText;
    if (isEvaluationError(rightText)) return rightText;
    return leftText + rightText;
  }

  if (["=", "<>", "<", ">", "<=", ">="].includes(operator)) {
    const comparison = compareValues(left, right);
    if (isEvaluationError(comparison)) return comparison;
    if (operator === "=") return comparison === 0;
    if (operator === "<>") return comparison !== 0;
    if (operator === "<") return comparison < 0;
    if (operator === ">") return comparison > 0;
    if (operator === "<=") return comparison <= 0;
    return comparison >= 0;
  }

  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  if (isEvaluationError(leftNumber)) return leftNumber;
  if (isEvaluationError(rightNumber)) return rightNumber;

  let result: number;
  if (operator === "+") result = leftNumber + rightNumber;
  else if (operator === "-") result = leftNumber - rightNumber;
  else if (operator === "*") result = leftNumber * rightNumber;
  else if (operator === "/") {
    if (rightNumber === 0) return evaluationError("#DIV/0!");
    result = leftNumber / rightNumber;
  } else if (operator === "^") {
    if (leftNumber === 0 && rightNumber < 0) return evaluationError("#DIV/0!");
    result = leftNumber ** rightNumber;
  } else {
    return evaluationError("#PARSE!");
  }

  return Number.isFinite(result) ? result : evaluationError("#NUM!");
}

function addressWithoutAnchors(address: A1Address): string {
  return `${columnNumberToLabel(address.column) ?? "#REF!"}${address.row}`;
}

function resolveReference(
  reference: ReferenceNode,
  currentSheetId: string,
  context: EvaluationContext,
): EvaluationScalar {
  const sheetId = reference.sheet ?? currentSheetId;
  const address = addressWithoutAnchors(reference.address);
  if (address.startsWith("#")) return evaluationError("#REF!");
  const key = `${sheetId.toLowerCase()}\u0000${address}`;

  const cached = context.memo.get(key);
  if (cached !== undefined) return cached;
  if (context.visiting.has(key)) return evaluationError("#CIRC!");

  context.visiting.add(key);
  let value: EvaluationScalar;
  try {
    const resolved = context.resolver(sheetId, address);
    if (resolved instanceof Date || resolved === null || typeof resolved !== "object") {
      if (typeof resolved === "string" && resolved.startsWith("=")) {
        value = evaluateFormulaText(resolved, sheetId, context);
      } else if (isFormulaError(resolved)) {
        value = evaluationError(resolved);
      } else {
        value = normalizePrimitive(resolved);
      }
    } else if (resolved.error) {
      value = evaluationError(resolved.error);
    } else if (typeof resolved.formula === "string") {
      value = evaluateFormulaText(resolved.formula, sheetId, context);
    } else {
      value = normalizePrimitive(resolved.value);
    }
  } catch {
    value = evaluationError("#REF!");
  } finally {
    context.visiting.delete(key);
  }

  context.memo.set(key, value);
  return value;
}

interface RectangleBounds {
  sheetId: string;
  firstRow: number;
  lastRow: number;
  firstColumn: number;
  lastColumn: number;
}

function wholeRangeBounds(
  range: WholeRangeNode,
  currentSheetId: string,
  context: EvaluationContext,
): RectangleBounds {
  const sheetId = range.sheet ?? currentSheetId;
  let used: { maxRow: number; maxCol: number } | null | undefined;
  try {
    used = context.hooks.getUsedRange?.(sheetId);
  } catch {
    used = undefined;
  }
  const maxRow = Math.min(
    Math.max(1, Math.trunc(used?.maxRow ?? FALLBACK_USED_RANGE_ROWS)),
    MAX_EXCEL_ROW,
  );
  const maxColumn = Math.min(
    Math.max(1, Math.trunc(used?.maxCol ?? FALLBACK_USED_RANGE_COLUMNS)),
    MAX_EXCEL_COLUMN,
  );
  return range.axis === "column"
    ? { sheetId, firstRow: 1, lastRow: maxRow, firstColumn: range.start, lastColumn: range.end }
    : { sheetId, firstRow: range.start, lastRow: range.end, firstColumn: 1, lastColumn: maxColumn };
}

function resolveRectangle(
  bounds: RectangleBounds,
  currentSheetId: string,
  context: EvaluationContext,
): EvaluationValue {
  const { sheetId, firstRow, lastRow, firstColumn, lastColumn } = bounds;
  const rowCount = lastRow - firstRow + 1;
  const columnCount = lastColumn - firstColumn + 1;
  if (rowCount > MAX_RANGE_CELLS / columnCount) {
    return resolveSparseRectangle(bounds, currentSheetId, context);
  }

  const values: EvaluationScalar[] = [];
  for (let row = firstRow; row <= lastRow; row += 1) {
    for (let column = firstColumn; column <= lastColumn; column += 1) {
      values.push(
        resolveReference(
          {
            kind: "reference",
            sheet: sheetId,
            address: { row, column, rowAbsolute: false, columnAbsolute: false },
          },
          currentSheetId,
          context,
        ),
      );
    }
  }
  return { kind: "evaluationRange", values, rowCount, columnCount };
}

function resolveSparseRectangle(
  bounds: RectangleBounds,
  currentSheetId: string,
  context: EvaluationContext,
): EvaluationValue {
  const forEachCellInRange = context.hooks.forEachCellInRange;
  if (!forEachCellInRange) return evaluationError("#VALUE!");

  const { sheetId, firstRow, lastRow, firstColumn, lastColumn } = bounds;
  // Hosts may visit populated cells in insertion order; collect coordinates
  // first and sort them row-major so order-sensitive consumers (TEXTJOIN,
  // NPV, ...) see cells in sheet order.
  const coordinates: Array<{ row: number; column: number }> = [];
  const values: EvaluationScalar[] = [];
  let overflowed = false;
  try {
    forEachCellInRange(
      sheetId,
      { startRow: firstRow, endRow: lastRow, startColumn: firstColumn, endColumn: lastColumn },
      (row, column) => {
        if (overflowed) return;
        if (row < firstRow || row > lastRow || column < firstColumn || column > lastColumn) {
          return;
        }
        if (coordinates.length >= MAX_RANGE_CELLS) {
          overflowed = true;
          return;
        }
        coordinates.push({ row, column });
      },
    );
    if (overflowed) return evaluationError("#VALUE!");
    coordinates.sort((a, b) => a.row - b.row || a.column - b.column);
    for (const { row, column } of coordinates) {
      values.push(
        resolveReference(
          {
            kind: "reference",
            sheet: sheetId,
            address: { row, column, rowAbsolute: false, columnAbsolute: false },
          },
          currentSheetId,
          context,
        ),
      );
    }
  } catch {
    return evaluationError("#REF!");
  }
  return {
    kind: "evaluationRange",
    values,
    rowCount: lastRow - firstRow + 1,
    columnCount: lastColumn - firstColumn + 1,
    sparse: true,
  };
}

function resolveRange(
  range: RangeNode | WholeRangeNode,
  currentSheetId: string,
  context: EvaluationContext,
): EvaluationValue {
  if (range.kind === "wholeRange") {
    return resolveRectangle(wholeRangeBounds(range, currentSheetId, context), currentSheetId, context);
  }

  const startSheet = range.start.sheet ?? currentSheetId;
  const endSheet = range.end.sheet ?? startSheet;
  if (startSheet !== endSheet) return evaluationError("#REF!");

  return resolveRectangle(
    {
      sheetId: startSheet,
      firstRow: Math.min(range.start.address.row, range.end.address.row),
      lastRow: Math.max(range.start.address.row, range.end.address.row),
      firstColumn: Math.min(range.start.address.column, range.end.address.column),
      lastColumn: Math.max(range.start.address.column, range.end.address.column),
    },
    currentSheetId,
    context,
  );
}

interface CollectedValue {
  value: EvaluationScalar;
  fromRange: boolean;
}

function collectValues(values: EvaluationValue[]): CollectedValue[] {
  const collected: CollectedValue[] = [];
  for (const value of values) {
    if (isEvaluationRange(value)) {
      for (const entry of value.values) collected.push({ value: entry, fromRange: true });
    } else {
      collected.push({ value, fromRange: false });
    }
  }
  return collected;
}

function collectNumbers(
  values: EvaluationValue[],
): { values: number[]; error?: EvaluationError } {
  const numbers: number[] = [];
  for (const entry of collectValues(values)) {
    if (isEvaluationError(entry.value)) return { values: numbers, error: entry.value };
    if (typeof entry.value === "number") {
      numbers.push(entry.value);
    } else if (!entry.fromRange && typeof entry.value === "boolean") {
      numbers.push(entry.value ? 1 : 0);
    } else if (!entry.fromRange && typeof entry.value === "string") {
      const converted = Number(entry.value.trim());
      if (entry.value.trim() !== "" && Number.isFinite(converted)) numbers.push(converted);
    }
  }
  return { values: numbers };
}

interface RectangularValues {
  values: EvaluationScalar[];
  rowCount: number;
  columnCount: number;
}

function asRectangularValues(
  value: EvaluationValue,
): RectangularValues | EvaluationError {
  if (isEvaluationError(value)) return value;
  if (isEvaluationRange(value)) {
    return value.sparse ? evaluationError("#VALUE!") : value;
  }
  return { values: [value], rowCount: 1, columnCount: 1 };
}

function sameRangeShape(left: RectangularValues, right: RectangularValues): boolean {
  return left.rowCount === right.rowCount && left.columnCount === right.columnCount;
}

function scalarArgument(value: EvaluationValue): EvaluationScalar | EvaluationError {
  return isEvaluationRange(value) ? evaluationError("#VALUE!") : value;
}

function numericTextValue(value: string): number | null {
  const trimmed = value.trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(trimmed)) {
    return null;
  }
  const number = Number(trimmed);
  return Number.isFinite(number) ? number : null;
}

interface WildcardToken {
  kind: "literal" | "single" | "many";
  value?: string;
}

function wildcardTokens(pattern: string, caseInsensitive: boolean): WildcardToken[] {
  const characters = Array.from(pattern);
  const tokens: WildcardToken[] = [];
  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index];
    const next = characters[index + 1];
    if (character === "~" && next && ["*", "?", "~"].includes(next)) {
      tokens.push({
        kind: "literal",
        value: caseInsensitive ? next.toLocaleLowerCase() : next,
      });
      index += 1;
    } else if (character === "*") {
      if (tokens[tokens.length - 1]?.kind !== "many") tokens.push({ kind: "many" });
    } else if (character === "?") {
      tokens.push({ kind: "single" });
    } else {
      tokens.push({
        kind: "literal",
        value: caseInsensitive ? character.toLocaleLowerCase() : character,
      });
    }
  }
  return tokens;
}

function wildcardPrefixMatches(
  characters: string[],
  start: number,
  tokens: WildcardToken[],
  work: { steps: number },
): boolean | null {
  let characterIndex = start;
  let tokenIndex = 0;
  let starIndex = -1;
  let starCharacterIndex = start;

  while (characterIndex < characters.length) {
    work.steps += 1;
    if (work.steps > MAX_WILDCARD_STEPS) return null;
    const token = tokens[tokenIndex];
    if (
      token &&
      (token.kind === "single" ||
        (token.kind === "literal" && token.value === characters[characterIndex]))
    ) {
      tokenIndex += 1;
      characterIndex += 1;
    } else if (token?.kind === "many") {
      starIndex = tokenIndex;
      starCharacterIndex = characterIndex;
      tokenIndex += 1;
    } else if (starIndex >= 0) {
      tokenIndex = starIndex + 1;
      starCharacterIndex += 1;
      characterIndex = starCharacterIndex;
    } else {
      return false;
    }

    if (tokenIndex === tokens.length) return true;
  }

  while (tokens[tokenIndex]?.kind === "many") tokenIndex += 1;
  return tokenIndex === tokens.length;
}

function wildcardMatches(
  value: string,
  tokens: WildcardToken[],
  caseInsensitive: boolean,
  work: { steps: number } = { steps: 0 },
): boolean | null {
  const characters = Array.from(
    caseInsensitive ? value.toLocaleLowerCase() : value,
  );
  const matched = wildcardPrefixMatches(characters, 0, tokens, work);
  if (!matched) return matched;

  // A full criteria match must consume the whole value. Re-run with an explicit
  // end sentinel so a prefix-only match cannot be accepted.
  const sentinel = "\u0000";
  const fullTokens = [...tokens, { kind: "literal", value: sentinel } as WildcardToken];
  return wildcardPrefixMatches([...characters, sentinel], 0, fullTokens, work);
}

function wildcardSearchPosition(value: string, pattern: string): number | EvaluationError {
  const characters = Array.from(value.toLocaleLowerCase());
  const tokens = wildcardTokens(pattern, true);
  const work = { steps: 0 };
  for (let start = 0; start <= characters.length; start += 1) {
    const matched = wildcardPrefixMatches(characters, start, tokens, work);
    if (matched === null) return evaluationError("#VALUE!");
    if (matched) return start;
  }
  return evaluationError("#VALUE!");
}

function compareCriterionValue(
  value: EvaluationScalar,
  operand: EvaluationScalar,
): number | null | EvaluationError {
  if (isEvaluationError(value)) return value;
  if (isEvaluationError(operand)) return operand;

  if (typeof operand === "number") {
    const candidate =
      typeof value === "number"
        ? value
        : typeof value === "string"
          ? numericTextValue(value)
          : null;
    if (candidate === null) return null;
    return candidate === operand ? 0 : candidate < operand ? -1 : 1;
  }

  if (typeof operand === "boolean") {
    if (typeof value !== "boolean") return null;
    return value === operand ? 0 : value ? 1 : -1;
  }

  const leftText = value === null ? "" : typeof value === "boolean" ? (value ? "TRUE" : "FALSE") : String(value);
  const rightText = operand === null ? "" : String(operand);
  const left = leftText.toLocaleLowerCase();
  const right = rightText.toLocaleLowerCase();
  return left === right ? 0 : left < right ? -1 : 1;
}

type CriterionTest = (
  value: EvaluationScalar,
) => boolean | EvaluationError;

function createCriterionTest(
  criterion: EvaluationScalar,
): CriterionTest | EvaluationError {
  if (isEvaluationError(criterion)) return criterion;

  let operator = "=";
  let operand: EvaluationScalar = criterion;
  let wildcard: WildcardToken[] | null = null;
  const wildcardWork = { steps: 0 };
  if (typeof criterion === "string") {
    const match = /^(<=|>=|<>|=|<|>)([\s\S]*)$/.exec(criterion);
    if (match) {
      operator = match[1];
      operand = match[2];
    }

    const operandText = String(operand);
    const numericOperand = numericTextValue(operandText);
    if (numericOperand !== null) operand = numericOperand;
    else if (operator === "=" || operator === "<>") {
      wildcard = wildcardTokens(operandText, true);
    }
  }

  return (value: EvaluationScalar): boolean | EvaluationError => {
    if (isEvaluationError(value)) return value;
    if (wildcard) {
      const text = value === null ? "" : typeof value === "boolean" ? (value ? "TRUE" : "FALSE") : String(value);
      const matches = wildcardMatches(text, wildcard, true, wildcardWork);
      if (matches === null) return evaluationError("#VALUE!");
      return operator === "<>" ? !matches : matches;
    }

    const comparison = compareCriterionValue(value, operand);
    if (isEvaluationError(comparison)) return comparison;
    if (comparison === null) return operator === "<>";
    if (operator === "=") return comparison === 0;
    if (operator === "<>") return comparison !== 0;
    if (operator === "<") return comparison < 0;
    if (operator === ">") return comparison > 0;
    if (operator === "<=") return comparison <= 0;
    return comparison >= 0;
  };
}

function conditionalAggregate(
  name:
    | "SUMIF"
    | "SUMIFS"
    | "COUNTIF"
    | "COUNTIFS"
    | "AVERAGEIF"
    | "AVERAGEIFS"
    | "MINIFS"
    | "MAXIFS",
  args: EvaluationValue[],
): EvaluationValue {
  const isPlural = name.endsWith("S");
  const isCount = name.startsWith("COUNT");
  const isAverage = name.startsWith("AVERAGE");
  const isExtreme = name === "MINIFS" || name === "MAXIFS";

  let resultRange: RectangularValues | undefined;
  let pairStart = 0;
  if (name === "SUMIFS" || name === "AVERAGEIFS" || isExtreme) {
    if (args.length < 3 || args.length % 2 === 0) return evaluationError("#VALUE!");
    const range = asRectangularValues(args[0]);
    if (isEvaluationError(range)) return range;
    resultRange = range;
    pairStart = 1;
  } else if (isPlural) {
    if (args.length < 2 || args.length % 2 !== 0) return evaluationError("#VALUE!");
  } else {
    const expectedMinimum = 2;
    const expectedMaximum = isCount ? 2 : 3;
    if (args.length < expectedMinimum || args.length > expectedMaximum) {
      return evaluationError("#VALUE!");
    }
  }

  const criteriaRanges: RectangularValues[] = [];
  const criteriaTests: CriterionTest[] = [];

  if (!isPlural) {
    const criteriaRange = asRectangularValues(args[0]);
    if (isEvaluationError(criteriaRange)) return criteriaRange;
    const criterion = scalarArgument(args[1]);
    if (isEvaluationError(criterion)) return criterion;
    const test = createCriterionTest(criterion);
    if (isEvaluationError(test)) return test;
    criteriaRanges.push(criteriaRange);
    criteriaTests.push(test);

    if (!isCount) {
      const aggregateRange = args[2]
        ? asRectangularValues(args[2])
        : criteriaRange;
      if (isEvaluationError(aggregateRange)) return aggregateRange;
      resultRange = aggregateRange;
    }
  } else {
    for (let index = pairStart; index < args.length; index += 2) {
      const criteriaRange = asRectangularValues(args[index]);
      if (isEvaluationError(criteriaRange)) return criteriaRange;
      const criterion = scalarArgument(args[index + 1]);
      if (isEvaluationError(criterion)) return criterion;
      const test = createCriterionTest(criterion);
      if (isEvaluationError(test)) return test;
      criteriaRanges.push(criteriaRange);
      criteriaTests.push(test);
    }
  }

  const shape = resultRange ?? criteriaRanges[0];
  if (!shape) return evaluationError("#VALUE!");
  if (
    criteriaRanges.some((range) => !sameRangeShape(range, shape)) ||
    (resultRange && !sameRangeShape(resultRange, shape))
  ) {
    return evaluationError("#VALUE!");
  }

  let count = 0;
  let sum = 0;
  let extreme: number | null = null;
  for (let index = 0; index < shape.values.length; index += 1) {
    let matches = true;
    for (let criterionIndex = 0; criterionIndex < criteriaRanges.length; criterionIndex += 1) {
      const result = criteriaTests[criterionIndex](
        criteriaRanges[criterionIndex].values[index],
      );
      if (isEvaluationError(result)) return result;
      if (!result) {
        matches = false;
        break;
      }
    }
    if (!matches) continue;

    if (isCount) {
      count += 1;
      continue;
    }

    const aggregate = resultRange!.values[index];
    if (isEvaluationError(aggregate)) return aggregate;
    if (typeof aggregate === "number") {
      if (isExtreme) {
        extreme =
          extreme === null
            ? aggregate
            : name === "MINIFS"
              ? Math.min(extreme, aggregate)
              : Math.max(extreme, aggregate);
        continue;
      }
      sum += aggregate;
      count += 1;
      if (!Number.isFinite(sum)) return evaluationError("#NUM!");
    }
  }

  if (isCount) return count;
  if (isExtreme) return extreme ?? 0;
  if (isAverage) return count ? sum / count : evaluationError("#DIV/0!");
  return sum;
}

function decimalShift(value: number, places: number): number {
  if (value === 0) return value;
  const [coefficient, exponent = "0"] = String(value).split(/[Ee]/);
  return Number(`${coefficient}e${Number(exponent) + places}`);
}

function directedRound(
  value: number,
  digits: number,
  direction: "up" | "down",
): number | EvaluationError {
  if (Math.abs(digits) > 308) return evaluationError("#NUM!");
  const shifted = decimalShift(Math.abs(value), digits);
  if (!Number.isFinite(shifted)) return evaluationError("#NUM!");
  const integer = direction === "up" ? Math.ceil(shifted) : Math.floor(shifted);
  const result = Math.sign(value) * decimalShift(integer, -digits);
  return Number.isFinite(result) ? (Object.is(result, -0) ? 0 : result) : evaluationError("#NUM!");
}

function safeTextResult(value: string): EvaluationValue {
  return value.length <= MAX_TEXT_RESULT_LENGTH
    ? value
    : evaluationError("#VALUE!");
}

function properCase(value: string): string {
  return value
    .toLocaleLowerCase()
    .replace(/(^|[^\p{L}])(\p{L})/gu, (_match, prefix: string, letter: string) =>
      prefix + letter.toLocaleUpperCase(),
    );
}

function parseInvariantValue(value: string): number | null {
  let source = value.trim();
  if (!source) return null;

  let sign = 1;
  if (source.startsWith("(") && source.endsWith(")")) {
    sign = -1;
    source = source.slice(1, -1).trim();
  }

  source = source.replace(/^[\$£€¥₹]\s*/, "");
  let divisor = 1;
  if (source.endsWith("%")) {
    divisor = 100;
    source = source.slice(0, -1).trim();
  }

  if (source.includes(",")) {
    if (!/^[+-]?\d{1,3}(?:,\d{3})*(?:\.\d*)?(?:[Ee][+-]?\d+)?$/.test(source)) {
      return null;
    }
    source = source.replace(/,/g, "");
  }
  const parsed = numericTextValue(source);
  if (parsed === null) return null;
  const result = (sign * parsed) / divisor;
  return Number.isFinite(result) ? result : null;
}

function utcDateFromSerial(serial: number): Date | null {
  if (!Number.isFinite(serial)) return null;
  const milliseconds = EXCEL_EPOCH_UTC + Math.floor(serial) * MILLISECONDS_PER_DAY;
  const date = new Date(milliseconds);
  return Number.isFinite(date.getTime()) ? date : null;
}

function serialFromUtcDate(date: Date): number | EvaluationError {
  const serial = (date.getTime() - EXCEL_EPOCH_UTC) / MILLISECONDS_PER_DAY;
  return Number.isFinite(serial) ? serial : evaluationError("#NUM!");
}

function createUtcDate(year: number, monthIndex: number, day: number): Date | null {
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, monthIndex, day);
  return Number.isFinite(date.getTime()) ? date : null;
}

function daysInUtcMonth(year: number, monthIndex: number): number | null {
  const date = createUtcDate(year, monthIndex + 1, 0);
  return date ? date.getUTCDate() : null;
}

function lookupComparison(
  candidate: EvaluationScalar,
  searchKey: EvaluationScalar,
): number | null | EvaluationError {
  if (isEvaluationError(candidate)) return candidate;
  if (isEvaluationError(searchKey)) return searchKey;

  if (typeof searchKey === "number") {
    if (typeof candidate !== "number") return null;
    return candidate === searchKey ? 0 : candidate < searchKey ? -1 : 1;
  }
  if (typeof searchKey === "boolean") {
    if (typeof candidate !== "boolean") return null;
    return candidate === searchKey ? 0 : candidate ? 1 : -1;
  }

  if (candidate !== null && typeof candidate !== "string") return null;
  const left = (candidate ?? "").toLocaleLowerCase();
  const right = (searchKey ?? "").toLocaleLowerCase();
  return left === right ? 0 : left < right ? -1 : 1;
}

function findLookupIndex(
  candidates: EvaluationScalar[],
  searchKey: EvaluationScalar,
  searchType: -1 | 0 | 1,
): number | EvaluationError {
  if (isEvaluationError(searchKey)) return searchKey;

  let result = -1;
  for (let index = 0; index < candidates.length; index += 1) {
    const comparison = lookupComparison(candidates[index], searchKey);
    if (isEvaluationError(comparison)) {
      if (searchType === 0) continue;
      return comparison;
    }
    if (comparison === null) continue;
    if (searchType === 0) {
      if (comparison === 0) return index;
      continue;
    }
    if (searchType === 1) {
      if (comparison <= 0) result = index;
      else break;
    } else if (comparison >= 0) {
      result = index;
    } else {
      break;
    }
  }
  return result >= 0 ? result : evaluationError("#N/A");
}

function rangeSlice(
  range: RectangularValues,
  row: number,
  column: number,
): EvaluationValue {
  if (row > 0 && column > 0) {
    return range.values[(row - 1) * range.columnCount + column - 1];
  }

  if (row === 0 && column === 0) {
    if (range.values.length === 1) return range.values[0];
    return { kind: "evaluationRange", ...range };
  }

  if (row === 0) {
    const values: EvaluationScalar[] = [];
    for (let currentRow = 0; currentRow < range.rowCount; currentRow += 1) {
      values.push(range.values[currentRow * range.columnCount + column - 1]);
    }
    return values.length === 1
      ? values[0]
      : { kind: "evaluationRange", values, rowCount: range.rowCount, columnCount: 1 };
  }

  const start = (row - 1) * range.columnCount;
  const values = range.values.slice(start, start + range.columnCount);
  return values.length === 1
    ? values[0]
    : { kind: "evaluationRange", values, rowCount: 1, columnCount: range.columnCount };
}

function evaluateLookupFunction(
  name: "INDEX" | "MATCH" | "VLOOKUP" | "HLOOKUP",
  args: EvaluationValue[],
): EvaluationValue {
  if (name === "INDEX") {
    if (args.length < 1 || args.length > 3) return evaluationError("#VALUE!");
    const range = asRectangularValues(args[0]);
    if (isEvaluationError(range)) return range;
    let row = 0;
    let column = 0;
    if (args[1] !== undefined) {
      const value = toNumber(scalarArgument(args[1]));
      if (isEvaluationError(value) || !Number.isInteger(value) || value < 0) {
        return isEvaluationError(value) ? value : evaluationError("#VALUE!");
      }
      row = value;
    }
    if (args[2] !== undefined) {
      const value = toNumber(scalarArgument(args[2]));
      if (isEvaluationError(value) || !Number.isInteger(value) || value < 0) {
        return isEvaluationError(value) ? value : evaluationError("#VALUE!");
      }
      column = value;
    }
    if (row > range.rowCount || column > range.columnCount) {
      return evaluationError("#REF!");
    }
    return rangeSlice(range, row, column);
  }

  if (name === "MATCH") {
    if (args.length < 2 || args.length > 3) return evaluationError("#VALUE!");
    const searchKey = scalarArgument(args[0]);
    if (isEvaluationError(searchKey)) return searchKey;
    const range = asRectangularValues(args[1]);
    if (isEvaluationError(range)) return range;
    if (range.rowCount > 1 && range.columnCount > 1) return evaluationError("#N/A");
    let searchType: -1 | 0 | 1 = 1;
    if (args[2] !== undefined) {
      const rawType = toNumber(scalarArgument(args[2]));
      if (isEvaluationError(rawType)) return rawType;
      if (rawType !== -1 && rawType !== 0 && rawType !== 1) {
        return evaluationError("#VALUE!");
      }
      searchType = rawType;
    }
    const index = findLookupIndex(range.values, searchKey, searchType);
    return isEvaluationError(index) ? index : index + 1;
  }

  if (args.length < 3 || args.length > 4) return evaluationError("#VALUE!");
  const searchKey = scalarArgument(args[0]);
  if (isEvaluationError(searchKey)) return searchKey;
  const range = asRectangularValues(args[1]);
  if (isEvaluationError(range)) return range;
  const rawIndex = toNumber(scalarArgument(args[2]));
  if (isEvaluationError(rawIndex)) return rawIndex;
  if (!Number.isInteger(rawIndex) || rawIndex < 1) return evaluationError("#VALUE!");
  const maximumIndex = name === "VLOOKUP" ? range.columnCount : range.rowCount;
  if (rawIndex > maximumIndex) return evaluationError("#REF!");

  let isSorted = true;
  if (args[3] !== undefined) {
    const rawSorted = toBoolean(scalarArgument(args[3]));
    if (isEvaluationError(rawSorted)) return rawSorted;
    isSorted = rawSorted;
  }

  const candidates: EvaluationScalar[] = [];
  if (name === "VLOOKUP") {
    for (let row = 0; row < range.rowCount; row += 1) {
      candidates.push(range.values[row * range.columnCount]);
    }
  } else {
    candidates.push(...range.values.slice(0, range.columnCount));
  }
  const matched = findLookupIndex(candidates, searchKey, isSorted ? 1 : 0);
  if (isEvaluationError(matched)) return matched;
  return name === "VLOOKUP"
    ? range.values[matched * range.columnCount + rawIndex - 1]
    : range.values[(rawIndex - 1) * range.columnCount + matched];
}

function evaluateLogicalArguments(
  name: "AND" | "OR" | "XOR",
  args: FormulaNode[],
  currentSheetId: string,
  context: EvaluationContext,
  depth: number,
): EvaluationValue {
  if (args.length === 0) return evaluationError("#VALUE!");
  let sawValue = false;
  let trueCount = 0;
  for (const argument of args) {
    const evaluated = evaluateNode(argument, currentSheetId, context, depth + 1);
    const entries = collectValues([evaluated]);
    for (const entry of entries) {
      if (isEvaluationError(entry.value)) return entry.value;
      if (entry.value === null || (entry.fromRange && typeof entry.value === "string")) {
        continue;
      }
      sawValue = true;
      const boolean = toBoolean(entry.value);
      if (isEvaluationError(boolean)) return boolean;
      if (name === "AND" && !boolean) return false;
      if (name === "OR" && boolean) return true;
      if (boolean) trueCount += 1;
    }
  }
  if (!sawValue) return evaluationError("#VALUE!");
  if (name === "XOR") return trueCount % 2 === 1;
  return name === "AND";
}

const MONTH_NAMES = [
  "JANUARY",
  "FEBRUARY",
  "MARCH",
  "APRIL",
  "MAY",
  "JUNE",
  "JULY",
  "AUGUST",
  "SEPTEMBER",
  "OCTOBER",
  "NOVEMBER",
  "DECEMBER",
];

function monthIndexFromName(token: string): number | null {
  const upper = token.toUpperCase();
  if (upper.length < 3) return null;
  const index = MONTH_NAMES.findIndex((month) => month.startsWith(upper));
  return index >= 0 ? index : null;
}

function dateSerialFromParts(year: number, month: number, day: number): number | null {
  const fullYear = year < 100 ? (year < 30 ? 2000 + year : 1900 + year) : year;
  const date = createUtcDate(fullYear, month - 1, day);
  if (
    !date ||
    date.getUTCFullYear() !== fullYear ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  const serial = serialFromUtcDate(date);
  return isEvaluationError(serial) || serial < 0 || serial > MAX_DATE_SERIAL
    ? null
    : serial;
}

function parseDatePrefix(text: string): { serial: number; rest: string } | null {
  const source = text.trim();
  let match = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/.exec(source);
  if (match) {
    const serial = dateSerialFromParts(Number(match[1]), Number(match[2]), Number(match[3]));
    if (serial === null) return null;
    return { serial, rest: source.slice(match[0].length).replace(/^[T ]/, "").trim() };
  }
  match = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/.exec(source);
  if (match) {
    const serial = dateSerialFromParts(Number(match[3]), Number(match[1]), Number(match[2]));
    if (serial === null) return null;
    return { serial, rest: source.slice(match[0].length).trim() };
  }
  match = /^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})/.exec(source);
  if (match) {
    const month = monthIndexFromName(match[1]);
    if (month === null) return null;
    const serial = dateSerialFromParts(Number(match[3]), month + 1, Number(match[2]));
    if (serial === null) return null;
    return { serial, rest: source.slice(match[0].length).trim() };
  }
  match = /^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{2,4})/.exec(source);
  if (match) {
    const month = monthIndexFromName(match[2]);
    if (month === null) return null;
    const serial = dateSerialFromParts(Number(match[3]), month + 1, Number(match[1]));
    if (serial === null) return null;
    return { serial, rest: source.slice(match[0].length).trim() };
  }
  return null;
}

function parseTimeOfDay(text: string): number | null {
  const match = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?(?:\s*([AaPp])\.?[Mm]\.?)?$/.exec(
    text.trim(),
  );
  if (!match) return null;
  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = match[3] === undefined ? 0 : Number(match[3]);
  if (minutes > 59 || seconds >= 60) return null;
  const meridiem = match[4]?.toUpperCase();
  if (meridiem) {
    if (hours < 1 || hours > 12) return null;
    if (meridiem === "P" && hours !== 12) hours += 12;
    if (meridiem === "A" && hours === 12) hours = 0;
  } else if (hours > 23) {
    return null;
  }
  return (hours * 3600 + minutes * 60 + seconds) / 86_400;
}

function dowFromSerial(serial: number): number {
  return ((((serial % 7) + 7) % 7) + 6) % 7;
}

function isoWeekNumber(date: Date): number {
  const target = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  target.setUTCDate(target.getUTCDate() - ((target.getUTCDay() + 6) % 7) + 3);
  const thursday = target.getTime();
  target.setUTCMonth(0, 1);
  if (target.getUTCDay() !== 4) {
    target.setUTCMonth(0, 1 + ((4 - target.getUTCDay()) + 7) % 7);
  }
  return 1 + Math.round((thursday - target.getTime()) / (7 * MILLISECONDS_PER_DAY));
}

function collectHolidaySerials(
  value: EvaluationValue | undefined,
): Set<number> | EvaluationError {
  const holidays = new Set<number>();
  if (value === undefined) return holidays;
  const collected = collectNumbers([value]);
  if (collected.error) return collected.error;
  for (const holiday of collected.values) holidays.add(Math.trunc(holiday));
  return holidays;
}

function finiteResult(value: number): number | EvaluationError {
  if (!Number.isFinite(value)) return evaluationError("#NUM!");
  return Object.is(value, -0) ? 0 : value;
}

function significanceRound(
  value: number,
  significance: number,
  direction: "up" | "down",
): number | EvaluationError {
  const quotient = value / significance;
  const nearest = Math.round(quotient);
  const steps =
    Math.abs(quotient - nearest) < 1e-9
      ? nearest
      : direction === "up"
        ? Math.ceil(quotient)
        : Math.floor(quotient);
  return finiteResult(steps * significance);
}

function roundToDigits(value: number, digits: number): number | EvaluationError {
  if (Math.abs(digits) > 308) return evaluationError("#NUM!");
  const shifted = decimalShift(Math.abs(value), digits);
  if (!Number.isFinite(shifted)) return evaluationError("#NUM!");
  return finiteResult(Math.sign(value) * decimalShift(Math.round(shifted), -digits));
}

function groupThousands(integerDigits: string): string {
  return integerDigits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function fixedNumberText(
  value: number,
  decimals: number,
  withCommas: boolean,
): string | EvaluationError {
  if (decimals > 100 || decimals < -100) return evaluationError("#VALUE!");
  const rounded = roundToDigits(value, decimals);
  if (isEvaluationError(rounded)) return rounded;
  const text = Math.abs(rounded).toFixed(Math.max(decimals, 0));
  const [integerPart, fractionPart] = text.split(".");
  const grouped = withCommas ? groupThousands(integerPart) : integerPart;
  const magnitude = fractionPart ? `${grouped}.${fractionPart}` : grouped;
  return rounded < 0 ? `-${magnitude}` : magnitude;
}

function occurrenceIndices(text: string, search: string): number[] {
  const indices: number[] = [];
  let from = 0;
  while (from <= text.length - search.length) {
    const index = text.indexOf(search, from);
    if (index < 0) break;
    indices.push(index);
    from = index + 1;
  }
  return indices;
}

function percentileInc(numbers: number[], k: number): number | EvaluationError {
  if (numbers.length === 0 || k < 0 || k > 1) return evaluationError("#NUM!");
  const sorted = [...numbers].sort((left, right) => left - right);
  const position = k * (sorted.length - 1);
  const lower = Math.floor(position);
  const fraction = position - lower;
  const base = sorted[lower];
  const next = sorted[Math.min(lower + 1, sorted.length - 1)];
  return finiteResult(base + fraction * (next - base));
}

function annuityFutureValue(
  rate: number,
  nper: number,
  pmt: number,
  pv: number,
  type: number,
): number {
  if (rate === 0) return -(pv + pmt * nper);
  const growth = (1 + rate) ** nper;
  return -(pv * growth + (pmt * (1 + rate * type) * (growth - 1)) / rate);
}

function annuityPayment(
  rate: number,
  nper: number,
  pv: number,
  fv: number,
  type: number,
): number | EvaluationError {
  if (nper === 0) return evaluationError("#NUM!");
  if (rate === 0) return finiteResult(-(pv + fv) / nper);
  const growth = (1 + rate) ** nper;
  const denominator = (1 + rate * type) * (growth - 1);
  if (!Number.isFinite(growth) || denominator === 0) return evaluationError("#NUM!");
  return finiteResult((-(pv * growth + fv) * rate) / denominator);
}

function financialArguments(
  values: EvaluationValue[],
  totalCount: number,
  optionalDefaults: number[],
): number[] | EvaluationError {
  const numbers: number[] = [];
  for (const value of values) {
    const number = toNumber(scalarArgument(value));
    if (isEvaluationError(number)) return number;
    numbers.push(number);
  }
  while (numbers.length < totalCount) {
    numbers.push(optionalDefaults[optionalDefaults.length - (totalCount - numbers.length)]);
  }
  return numbers;
}

function periodicPaymentPart(
  name: "IPMT" | "PPMT",
  values: EvaluationValue[],
): EvaluationValue {
  const parts = financialArguments(values, 6, [0, 0]);
  if (isEvaluationError(parts)) return parts;
  const [rate, rawPeriod, nper, pv, fv, rawType] = parts;
  const type = rawType ? 1 : 0;
  const period = Math.trunc(rawPeriod);
  if (period < 1 || period > nper) return evaluationError("#NUM!");
  const payment = annuityPayment(rate, nper, pv, fv, type);
  if (isEvaluationError(payment)) return payment;
  let interest: number;
  if (period === 1) {
    interest = type === 1 ? 0 : -pv * rate;
  } else if (type === 1) {
    interest = (annuityFutureValue(rate, period - 2, payment, pv, 1) - payment) * rate;
  } else {
    interest = annuityFutureValue(rate, period - 1, payment, pv, 0) * rate;
  }
  if (name === "IPMT") return finiteResult(interest);
  return finiteResult(payment - interest);
}

function rectangleRows(range: RectangularValues): EvaluationScalar[][] {
  const rows: EvaluationScalar[][] = [];
  for (let row = 0; row < range.rowCount; row += 1) {
    rows.push(range.values.slice(row * range.columnCount, (row + 1) * range.columnCount));
  }
  return rows;
}

function transposedRows(rows: EvaluationScalar[][]): EvaluationScalar[][] {
  const result: EvaluationScalar[][] = [];
  for (let column = 0; column < (rows[0]?.length ?? 0); column += 1) {
    result.push(rows.map((row) => row[column]));
  }
  return result;
}

function rangeFromRows(rows: EvaluationScalar[][]): EvaluationValue {
  if (rows.length === 0 || rows[0].length === 0) return evaluationError("#CALC!");
  const values: EvaluationScalar[] = [];
  for (const row of rows) values.push(...row);
  return { kind: "evaluationRange", values, rowCount: rows.length, columnCount: rows[0].length };
}

function arrayEntryKey(value: EvaluationScalar): string {
  if (isEvaluationError(value)) return `e:${value.code}`;
  if (value === null) return "n";
  if (typeof value === "boolean") return `b:${value}`;
  if (typeof value === "number") return `#:${value}`;
  return `s:${JSON.stringify(value.toLocaleLowerCase())}`;
}

function sortScalarCompare(left: EvaluationScalar, right: EvaluationScalar): number {
  const rank = (value: EvaluationScalar): number => {
    if (value === null) return 4;
    if (isEvaluationError(value)) return 3;
    if (typeof value === "boolean") return 2;
    if (typeof value === "string") return 1;
    return 0;
  };
  const leftRank = rank(left);
  const rightRank = rank(right);
  if (leftRank !== rightRank) return leftRank - rightRank;
  if (typeof left === "number" && typeof right === "number") {
    return left === right ? 0 : left < right ? -1 : 1;
  }
  if (typeof left === "string" && typeof right === "string") {
    const normalizedLeft = left.toLocaleLowerCase();
    const normalizedRight = right.toLocaleLowerCase();
    return normalizedLeft === normalizedRight ? 0 : normalizedLeft < normalizedRight ? -1 : 1;
  }
  if (typeof left === "boolean" && typeof right === "boolean") {
    return left === right ? 0 : left ? 1 : -1;
  }
  return 0;
}

function referenceNodeBounds(
  node: FormulaNode,
  currentSheetId: string,
  context: EvaluationContext,
): RectangleBounds | null {
  if (node.kind === "reference") {
    return {
      sheetId: node.sheet ?? currentSheetId,
      firstRow: node.address.row,
      lastRow: node.address.row,
      firstColumn: node.address.column,
      lastColumn: node.address.column,
    };
  }
  if (node.kind === "range") {
    const startSheet = node.start.sheet ?? currentSheetId;
    const endSheet = node.end.sheet ?? startSheet;
    if (startSheet !== endSheet) return null;
    return {
      sheetId: startSheet,
      firstRow: Math.min(node.start.address.row, node.end.address.row),
      lastRow: Math.max(node.start.address.row, node.end.address.row),
      firstColumn: Math.min(node.start.address.column, node.end.address.column),
      lastColumn: Math.max(node.start.address.column, node.end.address.column),
    };
  }
  if (node.kind === "wholeRange") {
    return wholeRangeBounds(node, currentSheetId, context);
  }
  return null;
}

interface FunctionEvaluation {
  currentSheetId: string;
  context: EvaluationContext;
  depth: number;
  argumentNodes: FormulaNode[];
  evaluate: (node: FormulaNode) => EvaluationValue;
}

function evaluateXlookup(call: FunctionEvaluation): EvaluationValue {
  const nodes = call.argumentNodes;
  const searchKey = scalarArgument(call.evaluate(nodes[0]));
  if (isEvaluationError(searchKey)) return searchKey;
  const lookupRange = asRectangularValues(call.evaluate(nodes[1]));
  if (isEvaluationError(lookupRange)) return lookupRange;
  const returnRange = asRectangularValues(call.evaluate(nodes[2]));
  if (isEvaluationError(returnRange)) return returnRange;
  if (lookupRange.rowCount > 1 && lookupRange.columnCount > 1) {
    return evaluationError("#VALUE!");
  }
  let matchMode = 0;
  if (nodes[4]) {
    const rawMode = toNumber(scalarArgument(call.evaluate(nodes[4])));
    if (isEvaluationError(rawMode)) return rawMode;
    if (rawMode !== 0 && rawMode !== -1 && rawMode !== 1 && rawMode !== 2) {
      return evaluationError("#VALUE!");
    }
    matchMode = rawMode;
  }
  let searchMode = 1;
  if (nodes[5]) {
    const rawMode = toNumber(scalarArgument(call.evaluate(nodes[5])));
    if (isEvaluationError(rawMode)) return rawMode;
    if (rawMode !== 1 && rawMode !== -1 && rawMode !== 2 && rawMode !== -2) {
      return evaluationError("#VALUE!");
    }
    searchMode = rawMode > 0 ? 1 : -1;
  }
  const candidates = lookupRange.values;
  const byRow = lookupRange.columnCount === 1;
  if (
    byRow
      ? returnRange.rowCount !== candidates.length
      : returnRange.columnCount !== candidates.length
  ) {
    return evaluationError("#VALUE!");
  }
  const wildcard =
    matchMode === 2 && typeof searchKey === "string"
      ? wildcardTokens(searchKey, true)
      : null;
  const wildcardWork = { steps: 0 };
  let exact = -1;
  let approximate = -1;
  const start = searchMode < 0 ? candidates.length - 1 : 0;
  for (let offset = 0; offset < candidates.length; offset += 1) {
    const index = start + offset * searchMode;
    const candidate = candidates[index];
    if (wildcard) {
      if (typeof candidate !== "string") continue;
      const matched = wildcardMatches(candidate, wildcard, true, wildcardWork);
      if (matched === null) return evaluationError("#VALUE!");
      if (matched) {
        exact = index;
        break;
      }
      continue;
    }
    const comparison = lookupComparison(candidate, searchKey);
    if (isEvaluationError(comparison) || comparison === null) continue;
    if (comparison === 0) {
      exact = index;
      break;
    }
    if (matchMode !== 0 && (matchMode === 1 ? comparison > 0 : comparison < 0)) {
      if (approximate < 0) {
        approximate = index;
      } else {
        const relative = lookupComparison(candidate, candidates[approximate]);
        if (
          typeof relative === "number" &&
          (matchMode === 1 ? relative < 0 : relative > 0)
        ) {
          approximate = index;
        }
      }
    }
  }
  const matched = exact >= 0 ? exact : approximate;
  if (matched < 0) {
    return nodes[3] ? call.evaluate(nodes[3]) : evaluationError("#N/A");
  }
  return byRow
    ? rangeSlice(returnRange, matched + 1, 0)
    : rangeSlice(returnRange, 0, matched + 1);
}

function evaluateOffset(call: FunctionEvaluation): EvaluationValue {
  const bounds = referenceNodeBounds(call.argumentNodes[0], call.currentSheetId, call.context);
  if (!bounds) return evaluationError("#VALUE!");
  const rawRows = toNumber(scalarArgument(call.evaluate(call.argumentNodes[1])));
  if (isEvaluationError(rawRows)) return rawRows;
  const rawColumns = toNumber(scalarArgument(call.evaluate(call.argumentNodes[2])));
  if (isEvaluationError(rawColumns)) return rawColumns;
  let height = bounds.lastRow - bounds.firstRow + 1;
  if (call.argumentNodes[3]) {
    const rawHeight = toNumber(scalarArgument(call.evaluate(call.argumentNodes[3])));
    if (isEvaluationError(rawHeight)) return rawHeight;
    height = Math.trunc(rawHeight);
  }
  let width = bounds.lastColumn - bounds.firstColumn + 1;
  if (call.argumentNodes[4]) {
    const rawWidth = toNumber(scalarArgument(call.evaluate(call.argumentNodes[4])));
    if (isEvaluationError(rawWidth)) return rawWidth;
    width = Math.trunc(rawWidth);
  }
  if (height < 1 || width < 1) return evaluationError("#REF!");
  const firstRow = bounds.firstRow + Math.trunc(rawRows);
  const firstColumn = bounds.firstColumn + Math.trunc(rawColumns);
  const lastRow = firstRow + height - 1;
  const lastColumn = firstColumn + width - 1;
  if (
    firstRow < 1 ||
    firstColumn < 1 ||
    lastRow > MAX_EXCEL_ROW ||
    lastColumn > MAX_EXCEL_COLUMN
  ) {
    return evaluationError("#REF!");
  }
  const resolved = resolveRectangle(
    { sheetId: bounds.sheetId, firstRow, lastRow, firstColumn, lastColumn },
    call.currentSheetId,
    call.context,
  );
  if (isEvaluationError(resolved)) return resolved;
  if (isEvaluationRange(resolved) && resolved.values.length === 1 && !resolved.sparse) {
    return resolved.values[0];
  }
  return resolved;
}

function conditionalSpec(
  name:
    | "SUMIF"
    | "SUMIFS"
    | "COUNTIF"
    | "COUNTIFS"
    | "AVERAGEIF"
    | "AVERAGEIFS"
    | "MINIFS"
    | "MAXIFS",
): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => conditionalAggregate(name, values),
  };
}

function lookupSpec(name: "INDEX" | "MATCH" | "VLOOKUP" | "HLOOKUP"): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => evaluateLookupFunction(name, values),
  };
}

function scalarTestSpec(test: (value: EvaluationScalar) => boolean): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const value = values[0];
      if (isEvaluationRange(value)) return evaluationError("#VALUE!");
      return test(value);
    },
  };
}

function rowColumnSpec(name: "ROW" | "COLUMN"): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: 1,
    lazy: true,
    impl: (_values, call) => {
      const node = call.argumentNodes[0];
      if (!node) {
        const cell = call.context.hooks.currentCell;
        if (!cell) return evaluationError("#VALUE!");
        return name === "ROW" ? cell.row : cell.column;
      }
      if (node.kind === "reference") {
        return name === "ROW" ? node.address.row : node.address.column;
      }
      if (node.kind === "range") {
        return name === "ROW"
          ? Math.min(node.start.address.row, node.end.address.row)
          : Math.min(node.start.address.column, node.end.address.column);
      }
      if (node.kind === "wholeRange") {
        return node.axis === (name === "ROW" ? "row" : "column") ? node.start : 1;
      }
      return evaluationError("#VALUE!");
    },
  };
}

function rowsColumnsSpec(name: "ROWS" | "COLUMNS"): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: 1,
    lazy: true,
    impl: (_values, call) => {
      const node = call.argumentNodes[0];
      if (node.kind === "reference") return 1;
      if (node.kind === "range") {
        return name === "ROWS"
          ? Math.abs(node.start.address.row - node.end.address.row) + 1
          : Math.abs(node.start.address.column - node.end.address.column) + 1;
      }
      if (node.kind === "wholeRange") {
        if (node.axis === (name === "ROWS" ? "row" : "column")) {
          return node.end - node.start + 1;
        }
        return name === "ROWS" ? MAX_EXCEL_ROW : MAX_EXCEL_COLUMN;
      }
      const value = call.evaluate(node);
      if (isEvaluationError(value)) return value;
      if (isEvaluationRange(value)) {
        return name === "ROWS" ? value.rowCount : value.columnCount;
      }
      return 1;
    },
  };
}

function numericAggregateSpec(name: "SUM" | "AVERAGE" | "MIN" | "MAX"): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => {
      const collected = collectNumbers(values);
      if (collected.error) return collected.error;
      if (name === "SUM" || name === "AVERAGE") {
        const sum = collected.values.reduce((result, value) => result + value, 0);
        if (!Number.isFinite(sum)) return evaluationError("#NUM!");
        if (name === "SUM") return sum;
        return collected.values.length
          ? sum / collected.values.length
          : evaluationError("#DIV/0!");
      }
      if (collected.values.length === 0) return 0;
      return collected.values.reduce((result, value) =>
        name === "MIN" ? Math.min(result, value) : Math.max(result, value),
      );
    },
  };
}

function productMedianSpec(name: "PRODUCT" | "MEDIAN"): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => {
      const collected = collectNumbers(values);
      if (collected.error) return collected.error;
      if (name === "PRODUCT") {
        if (collected.values.length === 0) return 0;
        let product = 1;
        for (const value of collected.values) {
          product *= value;
          if (!Number.isFinite(product)) return evaluationError("#NUM!");
        }
        return product;
      }

      if (collected.values.length === 0) return evaluationError("#NUM!");
      const sorted = [...collected.values].sort((left, right) => left - right);
      const middle = Math.floor(sorted.length / 2);
      const median =
        sorted.length % 2 === 1
          ? sorted[middle]
          : (sorted[middle - 1] + sorted[middle]) / 2;
      return Number.isFinite(median) ? median : evaluationError("#NUM!");
    },
  };
}

function largeSmallSpec(name: "LARGE" | "SMALL"): FunctionSpec {
  return {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const collected = collectNumbers([values[0]]);
      if (collected.error) return collected.error;
      const rawRank = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(rawRank)) return rawRank;
      if (!Number.isInteger(rawRank) || rawRank < 1 || rawRank > collected.values.length) {
        return evaluationError("#NUM!");
      }
      const sorted = [...collected.values].sort((left, right) =>
        name === "LARGE" ? right - left : left - right,
      );
      return sorted[rawRank - 1];
    },
  };
}

function rankSpec(): FunctionSpec {
  return {
    minArgs: 2,
    maxArgs: 3,
    impl: (values) => {
      const number = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(number)) return number;
      const collected = collectNumbers([values[1]]);
      if (collected.error) return collected.error;
      if (collected.values.length === 0) return evaluationError("#N/A");
      let ascending = false;
      if (values[2] !== undefined) {
        const order = toBoolean(scalarArgument(values[2]));
        if (isEvaluationError(order)) return order;
        ascending = order;
      }
      return (
        1 +
        collected.values.filter((value) =>
          ascending ? value < number : value > number,
        ).length
      );
    },
  };
}

function sampleStatisticSpec(name: "STDEV" | "STDEV.S" | "VAR" | "VAR.S"): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => {
      const collected = collectNumbers(values);
      if (collected.error) return collected.error;
      if (collected.values.length < 2) return evaluationError("#DIV/0!");
      const mean =
        collected.values.reduce((sum, value) => sum + value, 0) /
        collected.values.length;
      const variance =
        collected.values.reduce((sum, value) => sum + (value - mean) ** 2, 0) /
        (collected.values.length - 1);
      if (!Number.isFinite(variance)) return evaluationError("#NUM!");
      return name.startsWith("STDEV") ? Math.sqrt(variance) : variance;
    },
  };
}

function populationStatisticSpec(
  name: "STDEVP" | "STDEV.P" | "VARP" | "VAR.P",
): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => {
      const collected = collectNumbers(values);
      if (collected.error) return collected.error;
      if (collected.values.length === 0) return evaluationError("#DIV/0!");
      const mean =
        collected.values.reduce((sum, value) => sum + value, 0) /
        collected.values.length;
      const variance =
        collected.values.reduce((sum, value) => sum + (value - mean) ** 2, 0) /
        collected.values.length;
      if (!Number.isFinite(variance)) return evaluationError("#NUM!");
      return name.startsWith("STDEV") ? Math.sqrt(variance) : variance;
    },
  };
}

function percentileSpec(): FunctionSpec {
  return {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const collected = collectNumbers([values[0]]);
      if (collected.error) return collected.error;
      const k = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(k)) return k;
      return percentileInc(collected.values, k);
    },
  };
}

function quartileSpec(): FunctionSpec {
  return {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const collected = collectNumbers([values[0]]);
      if (collected.error) return collected.error;
      const rawQuarter = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(rawQuarter)) return rawQuarter;
      const quarter = Math.trunc(rawQuarter);
      if (quarter < 0 || quarter > 4) return evaluationError("#NUM!");
      return percentileInc(collected.values, quarter / 4);
    },
  };
}

function modeSpec(): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: Infinity,
    impl: (values) => {
      const collected = collectNumbers(values);
      if (collected.error) return collected.error;
      const counts = new Map<number, number>();
      let best: number | null = null;
      let bestCount = 1;
      for (const value of collected.values) {
        const count = (counts.get(value) ?? 0) + 1;
        counts.set(value, count);
        if (count > bestCount) {
          best = value;
          bestCount = count;
        }
      }
      return best === null ? evaluationError("#N/A") : best;
    },
  };
}

function numberSpec(compute: (value: number) => EvaluationValue): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      return isEvaluationError(value) ? value : compute(value);
    },
  };
}

function textSpec(compute: (text: string) => EvaluationValue): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const text = toText(scalarArgument(values[0]));
      return isEvaluationError(text) ? text : compute(text);
    },
  };
}

function directedRoundSpec(direction: "up" | "down"): FunctionSpec {
  return {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      const digitsValue = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(value)) return value;
      if (isEvaluationError(digitsValue)) return digitsValue;
      return directedRound(value, Math.trunc(digitsValue), direction);
    },
  };
}

function ceilingFloorMathSpec(direction: "up" | "down"): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: 3,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(value)) return value;
      let significance = 1;
      if (values[1] !== undefined) {
        const rawSignificance = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawSignificance)) return rawSignificance;
        significance = Math.abs(rawSignificance);
      }
      if (significance === 0) return 0;
      let mode = 0;
      if (values[2] !== undefined) {
        const rawMode = toNumber(scalarArgument(values[2]));
        if (isEvaluationError(rawMode)) return rawMode;
        mode = rawMode;
      }
      if (mode !== 0 && value < 0) {
        return significanceRound(value, significance, direction === "up" ? "down" : "up");
      }
      return significanceRound(value, significance, direction);
    },
  };
}

function leftRightSpec(name: "LEFT" | "RIGHT"): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: 2,
    impl: (values) => {
      const text = toText(scalarArgument(values[0]));
      if (isEvaluationError(text)) return text;
      let count = 1;
      if (values[1] !== undefined) {
        const rawCount = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawCount)) return rawCount;
        count = Math.trunc(rawCount);
      }
      if (count < 0) return evaluationError("#VALUE!");
      const characters = Array.from(text);
      return name === "LEFT"
        ? characters.slice(0, count).join("")
        : characters.slice(Math.max(characters.length - count, 0)).join("");
    },
  };
}

function concatSpec(): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => {
      let result = "";
      for (const entry of collectValues(values)) {
        const text = toText(entry.value);
        if (isEvaluationError(text)) return text;
        result += text;
        if (result.length > MAX_TEXT_RESULT_LENGTH) return evaluationError("#VALUE!");
      }
      return result;
    },
  };
}

function joinSpec(name: "JOIN" | "TEXTJOIN"): FunctionSpec {
  return {
    minArgs: name === "JOIN" ? 2 : 3,
    maxArgs: Infinity,
    impl: (values) => {
      const delimiter = toText(scalarArgument(values[0]));
      if (isEvaluationError(delimiter)) return delimiter;
      let ignoreEmpty = false;
      let textStart = 1;
      if (name === "TEXTJOIN") {
        const rawIgnoreEmpty = toBoolean(scalarArgument(values[1]));
        if (isEvaluationError(rawIgnoreEmpty)) return rawIgnoreEmpty;
        ignoreEmpty = rawIgnoreEmpty;
        textStart = 2;
      }
      const parts: string[] = [];
      for (const entry of collectValues(values.slice(textStart))) {
        if (isEvaluationError(entry.value)) return entry.value;
        const text = toText(entry.value);
        if (isEvaluationError(text)) return text;
        if (ignoreEmpty && text === "") continue;
        parts.push(text);
      }
      return safeTextResult(parts.join(delimiter));
    },
  };
}

function findSearchSpec(name: "FIND" | "SEARCH"): FunctionSpec {
  return {
    minArgs: 2,
    maxArgs: 3,
    impl: (values) => {
      const needle = toText(scalarArgument(values[0]));
      const haystack = toText(scalarArgument(values[1]));
      if (isEvaluationError(needle)) return needle;
      if (isEvaluationError(haystack)) return haystack;
      let startingAt = 1;
      if (values[2] !== undefined) {
        const rawStart = toNumber(scalarArgument(values[2]));
        if (isEvaluationError(rawStart)) return rawStart;
        startingAt = Math.trunc(rawStart);
      }
      const characters = Array.from(haystack);
      if (startingAt < 1 || startingAt > characters.length + 1) {
        return evaluationError("#VALUE!");
      }
      const prefix = characters.slice(0, startingAt - 1).join("");
      const source = characters.slice(startingAt - 1).join("");
      let position: number;
      if (name === "FIND") {
        position = source.indexOf(needle);
      } else {
        const matched = wildcardSearchPosition(source, needle);
        if (isEvaluationError(matched)) return matched;
        return startingAt + matched;
      }
      if (position < 0) return evaluationError("#VALUE!");
      return Array.from(prefix + source.slice(0, position)).length + 1;
    },
  };
}

function textBeforeAfterSpec(name: "TEXTBEFORE" | "TEXTAFTER"): FunctionSpec {
  return {
    minArgs: 2,
    maxArgs: 3,
    impl: (values) => {
      const text = toText(scalarArgument(values[0]));
      if (isEvaluationError(text)) return text;
      const delimiter = toText(scalarArgument(values[1]));
      if (isEvaluationError(delimiter)) return delimiter;
      if (delimiter === "") return evaluationError("#VALUE!");
      let instance = 1;
      if (values[2] !== undefined) {
        const rawInstance = toNumber(scalarArgument(values[2]));
        if (isEvaluationError(rawInstance)) return rawInstance;
        instance = Math.trunc(rawInstance);
      }
      if (instance === 0) return evaluationError("#VALUE!");
      const indices = occurrenceIndices(text, delimiter);
      if (Math.abs(instance) > indices.length) return evaluationError("#N/A");
      const index = instance > 0 ? indices[instance - 1] : indices[indices.length + instance];
      return name === "TEXTBEFORE"
        ? text.slice(0, index)
        : text.slice(index + delimiter.length);
    },
  };
}

function datePartSpec(name: "YEAR" | "MONTH" | "DAY"): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const serial = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(serial)) return serial;
      const date = utcDateFromSerial(serial);
      if (!date) return evaluationError("#NUM!");
      if (name === "YEAR") return date.getUTCFullYear();
      if (name === "MONTH") return date.getUTCMonth() + 1;
      return date.getUTCDate();
    },
  };
}

function edateSpec(name: "EDATE" | "EOMONTH"): FunctionSpec {
  return {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const serial = toNumber(scalarArgument(values[0]));
      const rawMonths = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(serial)) return serial;
      if (isEvaluationError(rawMonths)) return rawMonths;
      const start = utcDateFromSerial(serial);
      if (!start) return evaluationError("#NUM!");
      const months = Math.trunc(rawMonths);
      const targetFirst = createUtcDate(
        start.getUTCFullYear(),
        start.getUTCMonth() + months,
        1,
      );
      if (!targetFirst) return evaluationError("#NUM!");
      const targetYear = targetFirst.getUTCFullYear();
      const targetMonth = targetFirst.getUTCMonth();
      if (targetYear < 0 || targetYear > 10_000) return evaluationError("#NUM!");
      const lastDay = daysInUtcMonth(targetYear, targetMonth);
      if (lastDay === null) return evaluationError("#NUM!");
      const day = name === "EOMONTH" ? lastDay : Math.min(start.getUTCDate(), lastDay);
      const result = createUtcDate(targetYear, targetMonth, day);
      return result ? serialFromUtcDate(result) : evaluationError("#NUM!");
    },
  };
}

function timePartSpec(name: "HOUR" | "MINUTE" | "SECOND"): FunctionSpec {
  return {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const serial = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(serial)) return serial;
      if (serial < 0) return evaluationError("#NUM!");
      const totalSeconds = Math.round((serial - Math.floor(serial)) * 86_400) % 86_400;
      if (name === "HOUR") return Math.floor(totalSeconds / 3600);
      if (name === "MINUTE") return Math.floor(totalSeconds / 60) % 60;
      return totalSeconds % 60;
    },
  };
}

interface FunctionSpec {
  minArgs: number;
  maxArgs: number;
  lazy?: boolean;
  impl: (values: EvaluationValue[], call: FunctionEvaluation) => EvaluationValue;
}

function logicalSpec(name: "AND" | "OR" | "XOR"): FunctionSpec {
  return {
    minArgs: 0,
    maxArgs: Infinity,
    lazy: true,
    impl: (_values, call) =>
      evaluateLogicalArguments(name, call.argumentNodes, call.currentSheetId, call.context, call.depth),
  };
}

const FUNCTION_REGISTRY: Record<string, FunctionSpec> = {
  IF: {
    minArgs: 2,
    maxArgs: 3,
    lazy: true,
    impl: (_values, call) => {
      const condition = toBoolean(call.evaluate(call.argumentNodes[0]));
      if (isEvaluationError(condition)) return condition;
      if (condition) return call.evaluate(call.argumentNodes[1]);
      return call.argumentNodes[2] ? call.evaluate(call.argumentNodes[2]) : false;
    },
  },
  IFERROR: {
    minArgs: 1,
    maxArgs: 2,
    lazy: true,
    impl: (_values, call) => {
      const value = call.evaluate(call.argumentNodes[0]);
      if (!isEvaluationError(value)) return value;
      return call.argumentNodes[1] ? call.evaluate(call.argumentNodes[1]) : "";
    },
  },
  IFNA: {
    minArgs: 2,
    maxArgs: 2,
    lazy: true,
    impl: (_values, call) => {
      const value = call.evaluate(call.argumentNodes[0]);
      if (isEvaluationError(value) && value.code === "#N/A") {
        return call.evaluate(call.argumentNodes[1]);
      }
      return value;
    },
  },
  IFS: {
    minArgs: 2,
    maxArgs: Infinity,
    lazy: true,
    impl: (_values, call) => {
      if (call.argumentNodes.length % 2 !== 0) return evaluationError("#VALUE!");
      for (let index = 0; index < call.argumentNodes.length; index += 2) {
        const condition = toBoolean(call.evaluate(call.argumentNodes[index]));
        if (isEvaluationError(condition)) return condition;
        if (condition) return call.evaluate(call.argumentNodes[index + 1]);
      }
      return evaluationError("#N/A");
    },
  },
  AND: logicalSpec("AND"),
  OR: logicalSpec("OR"),
  XOR: logicalSpec("XOR"),
  NOT: {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const value = toBoolean(values[0]);
      return isEvaluationError(value) ? value : !value;
    },
  },
  CHOOSE: {
    minArgs: 2,
    maxArgs: Infinity,
    lazy: true,
    impl: (_values, call) => {
      const index = toNumber(scalarArgument(call.evaluate(call.argumentNodes[0])));
      if (isEvaluationError(index)) return index;
      const chosen = Math.trunc(index);
      if (chosen < 1 || chosen >= call.argumentNodes.length) {
        return evaluationError("#VALUE!");
      }
      return call.evaluate(call.argumentNodes[chosen]);
    },
  },
  TODAY: {
    minArgs: 0,
    maxArgs: 0,
    impl: (_values, call) => Math.floor(localDateToExcelSerial(call.context.now)),
  },
  NOW: {
    minArgs: 0,
    maxArgs: 0,
    impl: (_values, call) => localDateToExcelSerial(call.context.now),
  },
  NA: { minArgs: 0, maxArgs: 0, impl: () => evaluationError("#N/A") },
  SUMIF: conditionalSpec("SUMIF"),
  SUMIFS: conditionalSpec("SUMIFS"),
  COUNTIF: conditionalSpec("COUNTIF"),
  COUNTIFS: conditionalSpec("COUNTIFS"),
  AVERAGEIF: conditionalSpec("AVERAGEIF"),
  AVERAGEIFS: conditionalSpec("AVERAGEIFS"),
  MINIFS: conditionalSpec("MINIFS"),
  MAXIFS: conditionalSpec("MAXIFS"),
  INDEX: lookupSpec("INDEX"),
  MATCH: lookupSpec("MATCH"),
  VLOOKUP: lookupSpec("VLOOKUP"),
  HLOOKUP: lookupSpec("HLOOKUP"),
  XLOOKUP: { minArgs: 3, maxArgs: 6, lazy: true, impl: (_values, call) => evaluateXlookup(call) },
  OFFSET: { minArgs: 3, maxArgs: 5, lazy: true, impl: (_values, call) => evaluateOffset(call) },
  INDIRECT: {
    minArgs: 1,
    maxArgs: 2,
    impl: (values, call) => {
      const text = toText(scalarArgument(values[0]));
      if (isEvaluationError(text)) return text;
      if (values[1] !== undefined) {
        const a1 = toBoolean(scalarArgument(values[1]));
        if (isEvaluationError(a1)) return a1;
        if (!a1) return evaluationError("#REF!");
      }
      let node: FormulaNode;
      try {
        node = new FormulaParser(tokenize(text)).parse();
      } catch {
        return evaluationError("#REF!");
      }
      if (node.kind === "reference") {
        return resolveReference(node, call.currentSheetId, call.context);
      }
      if (node.kind === "range" || node.kind === "wholeRange") {
        return resolveRange(node, call.currentSheetId, call.context);
      }
      return evaluationError("#REF!");
    },
  },
  ROW: rowColumnSpec("ROW"),
  COLUMN: rowColumnSpec("COLUMN"),
  ROWS: rowsColumnsSpec("ROWS"),
  COLUMNS: rowsColumnsSpec("COLUMNS"),
  ISBLANK: scalarTestSpec((value) => value === null),
  ISNUMBER: scalarTestSpec((value) => typeof value === "number"),
  ISTEXT: scalarTestSpec((value) => typeof value === "string"),
  ISLOGICAL: scalarTestSpec((value) => typeof value === "boolean"),
  ISERROR: scalarTestSpec((value) => isEvaluationError(value)),
  ISERR: scalarTestSpec((value) => isEvaluationError(value) && value.code !== "#N/A"),
  ISNA: scalarTestSpec((value) => isEvaluationError(value) && value.code === "#N/A"),
  ISFORMULA: {
    minArgs: 1,
    maxArgs: 1,
    lazy: true,
    impl: (_values, call) => {
      const node = call.argumentNodes[0];
      const target =
        node.kind === "reference" ? node : node.kind === "range" ? node.start : null;
      if (!target) return evaluationError("#VALUE!");
      const sheetId = target.sheet ?? call.currentSheetId;
      const address = addressWithoutAnchors(target.address);
      if (address.startsWith("#")) return evaluationError("#REF!");
      const hook = call.context.hooks.isFormulaCell;
      if (hook) {
        try {
          return Boolean(hook(sheetId, address));
        } catch {
          return evaluationError("#REF!");
        }
      }
      try {
        const resolved = call.context.resolver(sheetId, address);
        if (typeof resolved === "string") return resolved.startsWith("=");
        return (
          typeof resolved === "object" &&
          resolved !== null &&
          !(resolved instanceof Date) &&
          typeof resolved.formula === "string"
        );
      } catch {
        return evaluationError("#REF!");
      }
    },
  },
  SUM: numericAggregateSpec("SUM"),
  AVERAGE: numericAggregateSpec("AVERAGE"),
  MIN: numericAggregateSpec("MIN"),
  MAX: numericAggregateSpec("MAX"),

  PRODUCT: productMedianSpec("PRODUCT"),
  MEDIAN: productMedianSpec("MEDIAN"),
  LARGE: largeSmallSpec("LARGE"),
  SMALL: largeSmallSpec("SMALL"),
  RANK: rankSpec(),
  "RANK.EQ": rankSpec(),
  STDEV: sampleStatisticSpec("STDEV"),
  "STDEV.S": sampleStatisticSpec("STDEV.S"),
  VAR: sampleStatisticSpec("VAR"),
  "VAR.S": sampleStatisticSpec("VAR.S"),
  STDEVP: populationStatisticSpec("STDEVP"),
  "STDEV.P": populationStatisticSpec("STDEV.P"),
  VARP: populationStatisticSpec("VARP"),
  "VAR.P": populationStatisticSpec("VAR.P"),
  COUNT: {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => {
      let count = 0;
      for (const entry of collectValues(values)) {
        if (isEvaluationError(entry.value)) {
          if (!entry.fromRange) return entry.value;
        } else if (typeof entry.value === "number") {
          count += 1;
        } else if (!entry.fromRange && typeof entry.value === "boolean") {
          count += 1;
        } else if (!entry.fromRange && typeof entry.value === "string") {
          const converted = Number(entry.value.trim());
          if (entry.value.trim() !== "" && Number.isFinite(converted)) count += 1;
        }
      }
      return count;
    },
  },
  COUNTA: {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => {
      let count = 0;
      for (const entry of collectValues(values)) {
        if (isEvaluationError(entry.value) && !entry.fromRange) return entry.value;
        if (entry.value !== null) count += 1;
      }
      return count;
    },
  },
  COUNTBLANK: {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const value = values[0];
      if (isEvaluationError(value)) return value;
      if (!isEvaluationRange(value)) {
        return value === null || value === "" ? 1 : 0;
      }
      let filled = 0;
      for (const entry of value.values) {
        if (isEvaluationError(entry)) {
          filled += 1;
        } else if (entry !== null && entry !== "") {
          filled += 1;
        }
      }
      return value.rowCount * value.columnCount - filled;
    },
  },
  AVERAGEA: {
    minArgs: 0,
    maxArgs: Infinity,
    impl: (values) => {
      let sum = 0;
      let count = 0;
      for (const entry of collectValues(values)) {
        if (isEvaluationError(entry.value)) return entry.value;
        if (entry.value === null) continue;
        if (typeof entry.value === "number") {
          sum += entry.value;
        } else if (typeof entry.value === "boolean") {
          sum += entry.value ? 1 : 0;
        } else if (!entry.fromRange) {
          const converted = toNumber(entry.value);
          if (isEvaluationError(converted)) return converted;
          sum += converted;
        }
        count += 1;
        if (!Number.isFinite(sum)) return evaluationError("#NUM!");
      }
      return count ? sum / count : evaluationError("#DIV/0!");
    },
  },
  SUMPRODUCT: {
    minArgs: 1,
    maxArgs: Infinity,
    impl: (values) => {
      const ranges: RectangularValues[] = [];
      for (const value of values) {
        const range = asRectangularValues(value);
        if (isEvaluationError(range)) return range;
        ranges.push(range);
      }
      const shape = ranges[0];
      if (ranges.some((range) => !sameRangeShape(range, shape))) {
        return evaluationError("#VALUE!");
      }
      let sum = 0;
      for (let index = 0; index < shape.values.length; index += 1) {
        let product = 1;
        for (const range of ranges) {
          const value = range.values[index];
          if (isEvaluationError(value)) return value;
          product *= typeof value === "number" ? value : 0;
        }
        sum += product;
      }
      return finiteResult(sum);
    },
  },
  PERCENTILE: percentileSpec(),
  "PERCENTILE.INC": percentileSpec(),
  QUARTILE: quartileSpec(),
  "QUARTILE.INC": quartileSpec(),
  MODE: modeSpec(),
  "MODE.SNGL": modeSpec(),

  ABS: numberSpec((value) => Math.abs(value)),
  INT: numberSpec((value) => Math.floor(value)),
  SQRT: numberSpec((value) => (value < 0 ? evaluationError("#NUM!") : Math.sqrt(value))),
  MOD: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const left = toNumber(scalarArgument(values[0]));
      const right = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(left)) return left;
      if (isEvaluationError(right)) return right;
      if (right === 0) return evaluationError("#DIV/0!");
      const result = left - right * Math.floor(left / right);
      return Number.isFinite(result)
        ? (Object.is(result, -0) ? 0 : result)
        : evaluationError("#NUM!");
    },
  },
  POWER: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const left = toNumber(scalarArgument(values[0]));
      const right = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(left)) return left;
      if (isEvaluationError(right)) return right;
      if (left === 0 && right < 0) return evaluationError("#DIV/0!");
      const result = left ** right;
      return Number.isFinite(result) ? result : evaluationError("#NUM!");
    },
  },
  ROUND: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      const digitsValue = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(value)) return value;
      if (isEvaluationError(digitsValue)) return digitsValue;
      const digits = Math.trunc(digitsValue);
      if (Math.abs(digits) > 308) return evaluationError("#NUM!");
      const shifted = decimalShift(Math.abs(value), digits);
      if (!Number.isFinite(shifted)) return evaluationError("#NUM!");
      const rounded = Math.sign(value) * decimalShift(Math.round(shifted), -digits);
      return Number.isFinite(rounded) ? rounded : evaluationError("#NUM!");
    },
  },
  ROUNDUP: directedRoundSpec("up"),
  ROUNDDOWN: directedRoundSpec("down"),
  TRUNC: {
    minArgs: 1,
    maxArgs: 2,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(value)) return value;
      let digits = 0;
      if (values[1] !== undefined) {
        const rawDigits = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawDigits)) return rawDigits;
        digits = Math.trunc(rawDigits);
      }
      return directedRound(value, digits, "down");
    },
  },
  CEILING: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      const significance = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(value)) return value;
      if (isEvaluationError(significance)) return significance;
      if (significance === 0) return 0;
      if (value > 0 && significance < 0) return evaluationError("#NUM!");
      return significanceRound(value, significance, "up");
    },
  },
  "CEILING.MATH": ceilingFloorMathSpec("up"),
  FLOOR: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      const significance = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(value)) return value;
      if (isEvaluationError(significance)) return significance;
      if (significance === 0) return evaluationError("#DIV/0!");
      if (value > 0 && significance < 0) return evaluationError("#NUM!");
      return significanceRound(value, significance, "down");
    },
  },
  "FLOOR.MATH": ceilingFloorMathSpec("down"),
  MROUND: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      const multiple = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(value)) return value;
      if (isEvaluationError(multiple)) return multiple;
      if (multiple === 0) return 0;
      if (value * multiple < 0) return evaluationError("#NUM!");
      const quotient = value / multiple;
      const rounded =
        Math.abs(quotient - Math.trunc(quotient) - 0.5) < 1e-9
          ? Math.trunc(quotient) + 1
          : Math.round(quotient);
      return finiteResult(rounded * multiple);
    },
  },
  EVEN: numberSpec((value) =>
    value === 0 ? 0 : finiteResult(Math.sign(value) * Math.ceil(Math.abs(value) / 2) * 2),
  ),
  ODD: numberSpec((value) => {
    const magnitude = Math.ceil((Math.abs(value) + 1) / 2) * 2 - 1;
    return finiteResult(value < 0 ? -magnitude : magnitude);
  }),
  SIGN: numberSpec((value) => Math.sign(value)),
  EXP: numberSpec((value) => finiteResult(Math.exp(value))),
  LN: numberSpec((value) => (value <= 0 ? evaluationError("#NUM!") : Math.log(value))),
  LOG: {
    minArgs: 1,
    maxArgs: 2,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(value)) return value;
      let base = 10;
      if (values[1] !== undefined) {
        const rawBase = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawBase)) return rawBase;
        base = rawBase;
      }
      if (value <= 0 || base <= 0) return evaluationError("#NUM!");
      if (base === 1) return evaluationError("#DIV/0!");
      return finiteResult(Math.log(value) / Math.log(base));
    },
  },
  LOG10: numberSpec((value) => (value <= 0 ? evaluationError("#NUM!") : Math.log10(value))),
  PI: { minArgs: 0, maxArgs: 0, impl: () => Math.PI },
  SIN: numberSpec((value) => finiteResult(Math.sin(value))),
  COS: numberSpec((value) => finiteResult(Math.cos(value))),
  TAN: numberSpec((value) => finiteResult(Math.tan(value))),
  ASIN: numberSpec((value) =>
    Math.abs(value) > 1 ? evaluationError("#NUM!") : Math.asin(value),
  ),
  ACOS: numberSpec((value) =>
    Math.abs(value) > 1 ? evaluationError("#NUM!") : Math.acos(value),
  ),
  ATAN: numberSpec((value) => Math.atan(value)),
  ATAN2: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const x = toNumber(scalarArgument(values[0]));
      const y = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(x)) return x;
      if (isEvaluationError(y)) return y;
      if (x === 0 && y === 0) return evaluationError("#DIV/0!");
      return Math.atan2(y, x);
    },
  },
  DEGREES: numberSpec((value) => finiteResult((value * 180) / Math.PI)),
  RADIANS: numberSpec((value) => finiteResult((value * Math.PI) / 180)),
  SQRTPI: numberSpec((value) =>
    value < 0 ? evaluationError("#NUM!") : Math.sqrt(value * Math.PI),
  ),
  RAND: { minArgs: 0, maxArgs: 0, impl: () => Math.random() },
  RANDBETWEEN: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const low = toNumber(scalarArgument(values[0]));
      const high = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(low)) return low;
      if (isEvaluationError(high)) return high;
      const bottom = Math.ceil(low);
      const top = Math.floor(high);
      if (bottom > top) return evaluationError("#NUM!");
      return bottom + Math.floor(Math.random() * (top - bottom + 1));
    },
  },

  LEN: textSpec((text) => Array.from(text).length),
  TRIM: textSpec((text) => text.replace(/^ +| +$/g, "").replace(/ +/g, " ")),
  UPPER: textSpec((text) => safeTextResult(text.toLocaleUpperCase())),
  LOWER: textSpec((text) => safeTextResult(text.toLocaleLowerCase())),
  PROPER: textSpec((text) => safeTextResult(properCase(text))),
  CLEAN: textSpec((text) => safeTextResult(text.replace(/[\u0000-\u001F]/g, ""))),
  LEFT: leftRightSpec("LEFT"),
  RIGHT: leftRightSpec("RIGHT"),
  MID: {
    minArgs: 3,
    maxArgs: 3,
    impl: (values) => {
      const text = toText(scalarArgument(values[0]));
      const rawStart = toNumber(scalarArgument(values[1]));
      const rawCount = toNumber(scalarArgument(values[2]));
      if (isEvaluationError(text)) return text;
      if (isEvaluationError(rawStart)) return rawStart;
      if (isEvaluationError(rawCount)) return rawCount;
      const start = Math.trunc(rawStart);
      const count = Math.trunc(rawCount);
      if (start < 1 || count < 0) return evaluationError("#VALUE!");
      return Array.from(text).slice(start - 1, start - 1 + count).join("");
    },
  },
  CONCAT: concatSpec(),
  CONCATENATE: concatSpec(),
  JOIN: joinSpec("JOIN"),
  TEXTJOIN: joinSpec("TEXTJOIN"),
  SUBSTITUTE: {
    minArgs: 3,
    maxArgs: 4,
    impl: (values) => {
      const source = toText(scalarArgument(values[0]));
      const searchFor = toText(scalarArgument(values[1]));
      const replacement = toText(scalarArgument(values[2]));
      if (isEvaluationError(source)) return source;
      if (isEvaluationError(searchFor)) return searchFor;
      if (isEvaluationError(replacement)) return replacement;
      if (searchFor === "") return safeTextResult(source);
      if (values[3] === undefined) {
        return safeTextResult(source.split(searchFor).join(replacement));
      }
      const rawOccurrence = toNumber(scalarArgument(values[3]));
      if (isEvaluationError(rawOccurrence)) return rawOccurrence;
      if (!Number.isInteger(rawOccurrence) || rawOccurrence < 1) {
        return evaluationError("#VALUE!");
      }
      let position = -1;
      let from = 0;
      for (let occurrence = 0; occurrence < rawOccurrence; occurrence += 1) {
        position = source.indexOf(searchFor, from);
        if (position < 0) return safeTextResult(source);
        from = position + searchFor.length;
      }
      return safeTextResult(
        source.slice(0, position) + replacement + source.slice(position + searchFor.length),
      );
    },
  },
  FIND: findSearchSpec("FIND"),
  SEARCH: findSearchSpec("SEARCH"),
  EXACT: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const left = toText(scalarArgument(values[0]));
      const right = toText(scalarArgument(values[1]));
      if (isEvaluationError(left)) return left;
      if (isEvaluationError(right)) return right;
      return left === right;
    },
  },
  VALUE: {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const text = toText(scalarArgument(values[0]));
      if (isEvaluationError(text)) return text;
      const parsed = parseInvariantValue(text);
      return parsed === null ? evaluationError("#VALUE!") : parsed;
    },
  },
  TEXT: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const value = scalarArgument(values[0]);
      if (isEvaluationError(value)) return value;
      const format = toText(scalarArgument(values[1]));
      if (isEvaluationError(format)) return format;
      let source: number | string;
      if (value === null) source = 0;
      else if (typeof value === "boolean") source = value ? "TRUE" : "FALSE";
      else if (typeof value === "string") source = numericTextValue(value) ?? value;
      else source = value;
      try {
        const formatted = SSF.format(format, source);
        return typeof formatted === "string"
          ? safeTextResult(formatted)
          : evaluationError("#VALUE!");
      } catch {
        return evaluationError("#VALUE!");
      }
    },
  },
  REPT: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const text = toText(scalarArgument(values[0]));
      if (isEvaluationError(text)) return text;
      const rawCount = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(rawCount)) return rawCount;
      const count = Math.trunc(rawCount);
      if (count < 0) return evaluationError("#VALUE!");
      if (count === 0 || text === "") return "";
      if (text.length * count > MAX_TEXT_RESULT_LENGTH) return evaluationError("#VALUE!");
      return text.repeat(count);
    },
  },
  CHAR: numberSpec((value) => {
    const code = Math.trunc(value);
    if (code < 1 || code > 255) return evaluationError("#VALUE!");
    return String.fromCharCode(code);
  }),
  CODE: textSpec((text) => {
    const code = text.codePointAt(0);
    return code === undefined ? evaluationError("#VALUE!") : code;
  }),
  REPLACE: {
    minArgs: 4,
    maxArgs: 4,
    impl: (values) => {
      const source = toText(scalarArgument(values[0]));
      const rawStart = toNumber(scalarArgument(values[1]));
      const rawCount = toNumber(scalarArgument(values[2]));
      const replacement = toText(scalarArgument(values[3]));
      if (isEvaluationError(source)) return source;
      if (isEvaluationError(rawStart)) return rawStart;
      if (isEvaluationError(rawCount)) return rawCount;
      if (isEvaluationError(replacement)) return replacement;
      const start = Math.trunc(rawStart);
      const count = Math.trunc(rawCount);
      if (start < 1 || count < 0) return evaluationError("#VALUE!");
      const characters = Array.from(source);
      return safeTextResult(
        characters.slice(0, start - 1).join("") +
          replacement +
          characters.slice(start - 1 + count).join(""),
      );
    },
  },
  FIXED: {
    minArgs: 1,
    maxArgs: 3,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(value)) return value;
      let decimals = 2;
      if (values[1] !== undefined) {
        const rawDecimals = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawDecimals)) return rawDecimals;
        decimals = Math.trunc(rawDecimals);
      }
      let noCommas = false;
      if (values[2] !== undefined) {
        const rawNoCommas = toBoolean(scalarArgument(values[2]));
        if (isEvaluationError(rawNoCommas)) return rawNoCommas;
        noCommas = rawNoCommas;
      }
      return fixedNumberText(value, decimals, !noCommas);
    },
  },
  DOLLAR: {
    minArgs: 1,
    maxArgs: 2,
    impl: (values) => {
      const value = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(value)) return value;
      let decimals = 2;
      if (values[1] !== undefined) {
        const rawDecimals = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawDecimals)) return rawDecimals;
        decimals = Math.trunc(rawDecimals);
      }
      const magnitude = fixedNumberText(Math.abs(value), decimals, true);
      if (isEvaluationError(magnitude)) return magnitude;
      const rounded = roundToDigits(value, decimals);
      if (isEvaluationError(rounded)) return rounded;
      return rounded < 0 ? `($${magnitude})` : `$${magnitude}`;
    },
  },
  TEXTBEFORE: textBeforeAfterSpec("TEXTBEFORE"),
  TEXTAFTER: textBeforeAfterSpec("TEXTAFTER"),
  DATE: {
    minArgs: 3,
    maxArgs: 3,
    impl: (values) => {
      const parts: number[] = [];
      for (const value of values) {
        const scalar = scalarArgument(value);
        if (isEvaluationError(scalar)) return scalar;
        if (typeof scalar !== "number") return evaluationError("#VALUE!");
        parts.push(Math.trunc(scalar));
      }
      let [year, month, day] = parts;
      if (year < 0 || year > 10_000) return evaluationError("#NUM!");
      if (year <= 1899) year += 1900;
      const date = createUtcDate(year, month - 1, day);
      if (!date || date.getUTCFullYear() < 0 || date.getUTCFullYear() > 10_000) {
        return evaluationError("#NUM!");
      }
      return serialFromUtcDate(date);
    },
  },
  YEAR: datePartSpec("YEAR"),
  MONTH: datePartSpec("MONTH"),
  DAY: datePartSpec("DAY"),
  EDATE: edateSpec("EDATE"),
  EOMONTH: edateSpec("EOMONTH"),
  DATEVALUE: textSpec((text) => {
    const parsed = parseDatePrefix(text);
    if (!parsed) return evaluationError("#VALUE!");
    if (parsed.rest !== "" && parseTimeOfDay(parsed.rest) === null) {
      return evaluationError("#VALUE!");
    }
    return parsed.serial;
  }),
  TIMEVALUE: textSpec((text) => {
    const direct = parseTimeOfDay(text);
    if (direct !== null) return direct;
    const parsed = parseDatePrefix(text);
    if (parsed && parsed.rest !== "") {
      const fraction = parseTimeOfDay(parsed.rest);
      if (fraction !== null) return fraction;
    }
    return evaluationError("#VALUE!");
  }),
  TIME: {
    minArgs: 3,
    maxArgs: 3,
    impl: (values) => {
      const parts: number[] = [];
      for (const value of values) {
        const part = toNumber(scalarArgument(value));
        if (isEvaluationError(part)) return part;
        parts.push(Math.trunc(part));
      }
      const total = parts[0] * 3600 + parts[1] * 60 + parts[2];
      if (total < 0) return evaluationError("#NUM!");
      return (total % 86_400) / 86_400;
    },
  },
  HOUR: timePartSpec("HOUR"),
  MINUTE: timePartSpec("MINUTE"),
  SECOND: timePartSpec("SECOND"),
  WEEKDAY: {
    minArgs: 1,
    maxArgs: 2,
    impl: (values) => {
      const serial = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(serial)) return serial;
      if (serial < 0) return evaluationError("#NUM!");
      let type = 1;
      if (values[1] !== undefined) {
        const rawType = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawType)) return rawType;
        type = Math.trunc(rawType);
      }
      const dow = dowFromSerial(Math.floor(serial));
      if (type === 1) return dow + 1;
      if (type === 2) return ((dow + 6) % 7) + 1;
      if (type === 3) return (dow + 6) % 7;
      if (type >= 11 && type <= 17) return ((dow + 6 - (type - 11) + 7) % 7) + 1;
      return evaluationError("#NUM!");
    },
  },
  WEEKNUM: {
    minArgs: 1,
    maxArgs: 2,
    impl: (values) => {
      const serial = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(serial)) return serial;
      if (serial < 0) return evaluationError("#NUM!");
      let type = 1;
      if (values[1] !== undefined) {
        const rawType = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawType)) return rawType;
        type = Math.trunc(rawType);
      }
      const date = utcDateFromSerial(serial);
      if (!date) return evaluationError("#NUM!");
      if (type === 21) return isoWeekNumber(date);
      let weekStart: number;
      if (type === 1 || type === 17) weekStart = 0;
      else if (type === 2) weekStart = 1;
      else if (type >= 11 && type <= 16) weekStart = type - 10;
      else return evaluationError("#NUM!");
      const jan1 = createUtcDate(date.getUTCFullYear(), 0, 1);
      if (!jan1) return evaluationError("#NUM!");
      const jan1Serial = serialFromUtcDate(jan1);
      if (isEvaluationError(jan1Serial)) return jan1Serial;
      const offset = (dowFromSerial(jan1Serial) - weekStart + 7) % 7;
      const dayOfYear = Math.floor(serial) - jan1Serial + 1;
      return Math.floor((dayOfYear - 1 + offset) / 7) + 1;
    },
  },
  DATEDIF: {
    minArgs: 3,
    maxArgs: 3,
    impl: (values) => {
      const rawStart = toNumber(scalarArgument(values[0]));
      const rawEnd = toNumber(scalarArgument(values[1]));
      const unit = toText(scalarArgument(values[2]));
      if (isEvaluationError(rawStart)) return rawStart;
      if (isEvaluationError(rawEnd)) return rawEnd;
      if (isEvaluationError(unit)) return unit;
      const startSerial = Math.trunc(rawStart);
      const endSerial = Math.trunc(rawEnd);
      if (startSerial < 0 || startSerial > endSerial) return evaluationError("#NUM!");
      const start = utcDateFromSerial(startSerial);
      const end = utcDateFromSerial(endSerial);
      if (!start || !end) return evaluationError("#NUM!");
      const kind = unit.trim().toUpperCase();
      if (kind === "D") return endSerial - startSerial;
      const beforeAnniversary =
        end.getUTCMonth() < start.getUTCMonth() ||
        (end.getUTCMonth() === start.getUTCMonth() &&
          end.getUTCDate() < start.getUTCDate());
      if (kind === "Y") {
        return end.getUTCFullYear() - start.getUTCFullYear() - (beforeAnniversary ? 1 : 0);
      }
      const wholeMonths =
        (end.getUTCFullYear() - start.getUTCFullYear()) * 12 +
        end.getUTCMonth() -
        start.getUTCMonth() -
        (end.getUTCDate() < start.getUTCDate() ? 1 : 0);
      if (kind === "M") return wholeMonths;
      if (kind === "YM") return ((wholeMonths % 12) + 12) % 12;
      if (kind === "MD") {
        if (end.getUTCDate() >= start.getUTCDate()) {
          return end.getUTCDate() - start.getUTCDate();
        }
        const previousMonthDays = daysInUtcMonth(end.getUTCFullYear(), end.getUTCMonth() - 1);
        if (previousMonthDays === null) return evaluationError("#NUM!");
        return end.getUTCDate() + previousMonthDays - start.getUTCDate();
      }
      if (kind === "YD") {
        const anchorYear = end.getUTCFullYear() - (beforeAnniversary ? 1 : 0);
        const anchor = createUtcDate(anchorYear, start.getUTCMonth(), start.getUTCDate());
        if (!anchor) return evaluationError("#NUM!");
        const anchorSerial = serialFromUtcDate(anchor);
        if (isEvaluationError(anchorSerial)) return anchorSerial;
        return endSerial - anchorSerial;
      }
      return evaluationError("#NUM!");
    },
  },
  NETWORKDAYS: {
    minArgs: 2,
    maxArgs: 3,
    impl: (values) => {
      const rawStart = toNumber(scalarArgument(values[0]));
      const rawEnd = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(rawStart)) return rawStart;
      if (isEvaluationError(rawEnd)) return rawEnd;
      let first = Math.trunc(rawStart);
      let last = Math.trunc(rawEnd);
      const sign = first <= last ? 1 : -1;
      if (sign < 0) [first, last] = [last, first];
      if (first < 0) return evaluationError("#NUM!");
      const holidays = collectHolidaySerials(values[2]);
      if (isEvaluationError(holidays)) return holidays;
      const fullWeeks = Math.floor((last - first + 1) / 7);
      let count = fullWeeks * 5;
      for (let day = first + fullWeeks * 7; day <= last; day += 1) {
        const dow = dowFromSerial(day);
        if (dow !== 0 && dow !== 6) count += 1;
      }
      for (const holiday of holidays) {
        if (holiday < first || holiday > last) continue;
        const dow = dowFromSerial(holiday);
        if (dow !== 0 && dow !== 6) count -= 1;
      }
      return sign * count;
    },
  },
  WORKDAY: {
    minArgs: 2,
    maxArgs: 3,
    impl: (values) => {
      const rawStart = toNumber(scalarArgument(values[0]));
      const rawDays = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(rawStart)) return rawStart;
      if (isEvaluationError(rawDays)) return rawDays;
      let current = Math.trunc(rawStart);
      if (current < 0 || current > MAX_DATE_SERIAL) return evaluationError("#NUM!");
      const holidays = collectHolidaySerials(values[2]);
      if (isEvaluationError(holidays)) return holidays;
      let remaining = Math.trunc(rawDays);
      const step = remaining >= 0 ? 1 : -1;
      remaining = Math.abs(remaining);
      while (remaining > 0) {
        current += step;
        if (current < 1 || current > MAX_DATE_SERIAL) return evaluationError("#NUM!");
        const dow = dowFromSerial(current);
        if (dow !== 0 && dow !== 6 && !holidays.has(current)) remaining -= 1;
      }
      return current;
    },
  },
  DAYS: {
    minArgs: 2,
    maxArgs: 2,
    impl: (values) => {
      const end = toNumber(scalarArgument(values[0]));
      const start = toNumber(scalarArgument(values[1]));
      if (isEvaluationError(end)) return end;
      if (isEvaluationError(start)) return start;
      return Math.trunc(end) - Math.trunc(start);
    },
  },
  PMT: {
    minArgs: 3,
    maxArgs: 5,
    impl: (values) => {
      const parts = financialArguments(values, 5, [0, 0]);
      if (isEvaluationError(parts)) return parts;
      const [rate, nper, pv, fv, type] = parts;
      return annuityPayment(rate, nper, pv, fv, type ? 1 : 0);
    },
  },
  FV: {
    minArgs: 3,
    maxArgs: 5,
    impl: (values) => {
      const parts = financialArguments(values, 5, [0, 0]);
      if (isEvaluationError(parts)) return parts;
      const [rate, nper, pmt, pv, type] = parts;
      return finiteResult(annuityFutureValue(rate, nper, pmt, pv, type ? 1 : 0));
    },
  },
  PV: {
    minArgs: 3,
    maxArgs: 5,
    impl: (values) => {
      const parts = financialArguments(values, 5, [0, 0]);
      if (isEvaluationError(parts)) return parts;
      const [rate, nper, pmt, fv, rawType] = parts;
      const type = rawType ? 1 : 0;
      if (rate === 0) return finiteResult(-(fv + pmt * nper));
      const growth = (1 + rate) ** nper;
      if (!Number.isFinite(growth) || growth === 0) return evaluationError("#NUM!");
      return finiteResult(-(fv + (pmt * (1 + rate * type) * (growth - 1)) / rate) / growth);
    },
  },
  NPER: {
    minArgs: 3,
    maxArgs: 5,
    impl: (values) => {
      const parts = financialArguments(values, 5, [0, 0]);
      if (isEvaluationError(parts)) return parts;
      const [rate, pmt, pv, fv, rawType] = parts;
      const type = rawType ? 1 : 0;
      if (rate === 0) {
        if (pmt === 0) return evaluationError("#NUM!");
        return finiteResult(-(pv + fv) / pmt);
      }
      if (rate <= -1) return evaluationError("#NUM!");
      const payment = (pmt * (1 + rate * type)) / rate;
      const denominator = pv + payment;
      if (denominator === 0) return evaluationError("#NUM!");
      const ratio = (payment - fv) / denominator;
      if (ratio <= 0) return evaluationError("#NUM!");
      return finiteResult(Math.log(ratio) / Math.log(1 + rate));
    },
  },
  IPMT: {
    minArgs: 4,
    maxArgs: 6,
    impl: (values) => periodicPaymentPart("IPMT", values),
  },
  PPMT: {
    minArgs: 4,
    maxArgs: 6,
    impl: (values) => periodicPaymentPart("PPMT", values),
  },
  NPV: {
    minArgs: 2,
    maxArgs: Infinity,
    impl: (values) => {
      const rate = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(rate)) return rate;
      if (rate === -1) return evaluationError("#DIV/0!");
      const collected = collectNumbers(values.slice(1));
      if (collected.error) return collected.error;
      let sum = 0;
      for (let index = 0; index < collected.values.length; index += 1) {
        sum += collected.values[index] / (1 + rate) ** (index + 1);
      }
      return finiteResult(sum);
    },
  },
  IRR: {
    minArgs: 1,
    maxArgs: 2,
    impl: (values) => {
      const collected = collectNumbers([values[0]]);
      if (collected.error) return collected.error;
      const cashflows = collected.values;
      if (
        cashflows.length < 2 ||
        !cashflows.some((value) => value > 0) ||
        !cashflows.some((value) => value < 0)
      ) {
        return evaluationError("#NUM!");
      }
      let guess = 0.1;
      if (values[1] !== undefined) {
        const rawGuess = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawGuess)) return rawGuess;
        guess = rawGuess;
      }
      const presentValue = (rate: number): number => {
        let sum = 0;
        for (let index = 0; index < cashflows.length; index += 1) {
          sum += cashflows[index] / (1 + rate) ** index;
        }
        return sum;
      };
      let rate = guess <= -1 ? 0.1 : guess;
      for (let iteration = 0; iteration < 100; iteration += 1) {
        const value = presentValue(rate);
        if (Math.abs(value) < 1e-9) return finiteResult(rate);
        let derivative = 0;
        for (let index = 1; index < cashflows.length; index += 1) {
          derivative -= (index * cashflows[index]) / (1 + rate) ** (index + 1);
        }
        if (!Number.isFinite(derivative) || derivative === 0) break;
        let next = rate - value / derivative;
        if (next <= -1) next = (rate - 1) / 2;
        if (Math.abs(next - rate) < 1e-12) {
          rate = next;
          break;
        }
        rate = next;
      }
      return Math.abs(presentValue(rate)) < 1e-6 ? finiteResult(rate) : evaluationError("#NUM!");
    },
  },
  RATE: {
    minArgs: 3,
    maxArgs: 6,
    impl: (values) => {
      const parts = financialArguments(values, 6, [0, 0, 0.1]);
      if (isEvaluationError(parts)) return parts;
      const [nper, pmt, pv, fv, rawType, guess] = parts;
      const type = rawType ? 1 : 0;
      if (nper <= 0) return evaluationError("#NUM!");
      const balance = (rate: number): number =>
        rate === 0
          ? pv + pmt * nper + fv
          : pv * (1 + rate) ** nper +
            (pmt * (1 + rate * type) * ((1 + rate) ** nper - 1)) / rate +
            fv;
      let rate = guess <= -1 ? 0.1 : guess;
      const step = 1e-7;
      for (let iteration = 0; iteration < 100; iteration += 1) {
        const value = balance(rate);
        if (Math.abs(value) < 1e-9) return finiteResult(rate);
        const derivative = (balance(rate + step) - balance(rate - step)) / (2 * step);
        if (!Number.isFinite(derivative) || derivative === 0) break;
        let next = rate - value / derivative;
        if (next <= -1) next = (rate - 1) / 2;
        if (Math.abs(next - rate) < 1e-12) {
          rate = next;
          break;
        }
        rate = next;
      }
      return Math.abs(balance(rate)) < 1e-6 ? finiteResult(rate) : evaluationError("#NUM!");
    },
  },
  UNIQUE: {
    minArgs: 1,
    maxArgs: 3,
    impl: (values) => {
      const range = asRectangularValues(values[0]);
      if (isEvaluationError(range)) return range;
      const byColumn = toBoolean(values[1] === undefined ? false : scalarArgument(values[1]));
      if (isEvaluationError(byColumn)) return byColumn;
      const exactlyOnce = toBoolean(values[2] === undefined ? false : scalarArgument(values[2]));
      if (isEvaluationError(exactlyOnce)) return exactlyOnce;
      let rows = rectangleRows(range);
      if (byColumn) rows = transposedRows(rows);
      const counts = new Map<string, number>();
      for (const row of rows) {
        const key = row.map(arrayEntryKey).join(",");
        counts.set(key, (counts.get(key) ?? 0) + 1);
      }
      const seen = new Set<string>();
      const kept: EvaluationScalar[][] = [];
      for (const row of rows) {
        const key = row.map(arrayEntryKey).join(",");
        if (seen.has(key)) continue;
        seen.add(key);
        if (exactlyOnce && counts.get(key) !== 1) continue;
        kept.push(row);
      }
      return rangeFromRows(byColumn ? transposedRows(kept) : kept);
    },
  },
  SORT: {
    minArgs: 1,
    maxArgs: 4,
    impl: (values) => {
      const range = asRectangularValues(values[0]);
      if (isEvaluationError(range)) return range;
      let sortIndex = 1;
      if (values[1] !== undefined) {
        const rawIndex = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawIndex)) return rawIndex;
        sortIndex = Math.trunc(rawIndex);
      }
      let order = 1;
      if (values[2] !== undefined) {
        const rawOrder = toNumber(scalarArgument(values[2]));
        if (isEvaluationError(rawOrder)) return rawOrder;
        if (rawOrder !== 1 && rawOrder !== -1) return evaluationError("#VALUE!");
        order = rawOrder;
      }
      const byColumn = toBoolean(values[3] === undefined ? false : scalarArgument(values[3]));
      if (isEvaluationError(byColumn)) return byColumn;
      let rows = rectangleRows(range);
      if (byColumn) rows = transposedRows(rows);
      if (sortIndex < 1 || sortIndex > (rows[0]?.length ?? 0)) {
        return evaluationError("#VALUE!");
      }
      const sorted = [...rows].sort((left, right) => {
        const leftValue = left[sortIndex - 1];
        const rightValue = right[sortIndex - 1];
        const leftBlank = leftValue === null;
        const rightBlank = rightValue === null;
        if (leftBlank || rightBlank) {
          return leftBlank === rightBlank ? 0 : leftBlank ? 1 : -1;
        }
        return sortScalarCompare(leftValue, rightValue) * order;
      });
      return rangeFromRows(byColumn ? transposedRows(sorted) : sorted);
    },
  },
  FILTER: {
    minArgs: 2,
    maxArgs: 3,
    lazy: true,
    impl: (_values, call) => {
      const range = asRectangularValues(call.evaluate(call.argumentNodes[0]));
      if (isEvaluationError(range)) return range;
      const include = asRectangularValues(call.evaluate(call.argumentNodes[1]));
      if (isEvaluationError(include)) return include;
      const rows = rectangleRows(range);
      const byRow = include.columnCount === 1 && include.rowCount === range.rowCount;
      const byColumn = !byRow && include.rowCount === 1 && include.columnCount === range.columnCount;
      if (!byRow && !byColumn) return evaluationError("#VALUE!");
      const flags: boolean[] = [];
      for (const value of include.values) {
        const flag = toBoolean(value);
        if (isEvaluationError(flag)) return flag;
        flags.push(flag);
      }
      const kept = byRow
        ? rows.filter((_row, index) => flags[index])
        : transposedRows(transposedRows(rows).filter((_column, index) => flags[index]));
      if (kept.length === 0 || kept[0].length === 0) {
        return call.argumentNodes[2]
          ? call.evaluate(call.argumentNodes[2])
          : evaluationError("#CALC!");
      }
      return rangeFromRows(kept);
    },
  },
  SEQUENCE: {
    minArgs: 1,
    maxArgs: 4,
    impl: (values) => {
      const rawRows = toNumber(scalarArgument(values[0]));
      if (isEvaluationError(rawRows)) return rawRows;
      const rows = Math.trunc(rawRows);
      let columns = 1;
      if (values[1] !== undefined) {
        const rawColumns = toNumber(scalarArgument(values[1]));
        if (isEvaluationError(rawColumns)) return rawColumns;
        columns = Math.trunc(rawColumns);
      }
      let start = 1;
      if (values[2] !== undefined) {
        const rawStart = toNumber(scalarArgument(values[2]));
        if (isEvaluationError(rawStart)) return rawStart;
        start = rawStart;
      }
      let step = 1;
      if (values[3] !== undefined) {
        const rawStep = toNumber(scalarArgument(values[3]));
        if (isEvaluationError(rawStep)) return rawStep;
        step = rawStep;
      }
      if (rows < 1 || columns < 1) return evaluationError("#VALUE!");
      if (rows > MAX_RANGE_CELLS / columns) return evaluationError("#VALUE!");
      const resultValues: EvaluationScalar[] = [];
      for (let index = 0; index < rows * columns; index += 1) {
        const value = start + index * step;
        if (!Number.isFinite(value)) return evaluationError("#NUM!");
        resultValues.push(value);
      }
      return { kind: "evaluationRange", values: resultValues, rowCount: rows, columnCount: columns };
    },
  },
  TRANSPOSE: {
    minArgs: 1,
    maxArgs: 1,
    impl: (values) => {
      const range = asRectangularValues(values[0]);
      if (isEvaluationError(range)) return range;
      return rangeFromRows(transposedRows(rectangleRows(range)));
    },
  },
};

function evaluateCall(
  call: CallNode,
  currentSheetId: string,
  context: EvaluationContext,
  depth: number,
): EvaluationValue {
  const name = call.name.replace(/^_xlfn\./i, "").toUpperCase();
  const spec = Object.prototype.hasOwnProperty.call(FUNCTION_REGISTRY, name)
    ? FUNCTION_REGISTRY[name]
    : undefined;
  if (!spec) return evaluationError("#NAME?");
  if (call.arguments.length < spec.minArgs || call.arguments.length > spec.maxArgs) {
    return evaluationError("#VALUE!");
  }
  const evaluation: FunctionEvaluation = {
    currentSheetId,
    context,
    depth,
    argumentNodes: call.arguments,
    evaluate: (node) => evaluateNode(node, currentSheetId, context, depth + 1),
  };
  const values = spec.lazy
    ? []
    : call.arguments.map((argument) =>
        evaluateNode(argument, currentSheetId, context, depth + 1),
      );
  return spec.impl(values, evaluation);
}

function resolveName(
  node: NameNode,
  currentSheetId: string,
  context: EvaluationContext,
  depth: number,
): EvaluationValue {
  const resolveDefinedName = context.hooks.resolveDefinedName;
  if (!resolveDefinedName) return evaluationError("#NAME?");
  let resolved: FormulaPrimitive;
  try {
    resolved = resolveDefinedName(node.name);
  } catch {
    return evaluationError("#NAME?");
  }
  if (resolved === undefined || resolved === null) return evaluationError("#NAME?");
  if (typeof resolved === "string") {
    let parsed: FormulaNode | null = null;
    try {
      parsed = new FormulaParser(tokenize(resolved)).parse();
    } catch {
      parsed = null;
    }
    if (
      parsed &&
      (parsed.kind === "reference" || parsed.kind === "range" || parsed.kind === "wholeRange")
    ) {
      return evaluateNode(parsed, currentSheetId, context, depth + 1);
    }
    return resolved;
  }
  return normalizePrimitive(resolved);
}

function evaluateNode(
  node: FormulaNode,
  currentSheetId: string,
  context: EvaluationContext,
  depth = 0,
): EvaluationValue {
  if (depth > MAX_CALCULATION_DEPTH) return evaluationError("#CALC!");

  if (node.kind === "literal") return node.value;
  if (node.kind === "error") return evaluationError(node.value);
  if (node.kind === "reference") return resolveReference(node, currentSheetId, context);
  if (node.kind === "range" || node.kind === "wholeRange") {
    return resolveRange(node, currentSheetId, context);
  }
  if (node.kind === "array") {
    return {
      kind: "evaluationRange",
      values: node.values,
      rowCount: node.rowCount,
      columnCount: node.columnCount,
    };
  }
  if (node.kind === "name") return resolveName(node, currentSheetId, context, depth);
  if (node.kind === "call") return evaluateCall(node, currentSheetId, context, depth);

  if (node.kind === "unary") {
    const value = toNumber(evaluateNode(node.operand, currentSheetId, context, depth + 1));
    if (isEvaluationError(value)) return value;
    if (node.operator === "-") return -value;
    if (node.operator === "%") return value / 100;
    return value;
  }

  const left = evaluateNode(node.left, currentSheetId, context, depth + 1);
  const right = evaluateNode(node.right, currentSheetId, context, depth + 1);
  return applyBinaryOperator(node.operator, left, right);
}

function evaluateFormulaText(
  formula: string,
  currentSheetId: string,
  context: EvaluationContext,
): EvaluationScalar {
  if (context.calculationDepth >= MAX_CALCULATION_DEPTH) {
    return evaluationError("#CALC!");
  }

  context.calculationDepth += 1;
  try {
    let source = formula.trim();
    if (source.startsWith("=")) source = source.slice(1).trim();
    if (!source) return evaluationError("#PARSE!");
    const node = new FormulaParser(tokenize(source)).parse();
    const value = evaluateNode(node, currentSheetId, context);
    if (isEvaluationRange(value)) {
      return value.sparse ? evaluationError("#VALUE!") : (value.values[0] ?? null);
    }
    return value;
  } catch (error) {
    return evaluationError(
      error instanceof FormulaParseError ? error.formulaError : "#PARSE!",
    );
  } finally {
    context.calculationDepth -= 1;
  }
}

/**
 * Safely evaluate a common Excel-style formula without executing JavaScript.
 * `TODAY` and `NOW` return Excel serial date numbers. A formula evaluating to
 * a rectangular array (dynamic-array functions, `A1:B2`, `{1,2;3,4}`) returns
 * its top-left value; spilling is not supported.
 */
export function evaluateFormula(
  formula: string,
  currentSheetId: string,
  resolver: FormulaResolver,
  hooks?: FormulaEvaluationHooks,
): FormulaResult {
  if (typeof formula !== "string" || typeof resolver !== "function") return "#PARSE!";
  const context: EvaluationContext = {
    resolver,
    hooks: hooks ?? {},
    visiting: new Set(),
    memo: new Map(),
    now: new Date(),
    calculationDepth: 0,
  };
  const result = evaluateFormulaText(formula, currentSheetId, context);
  if (isEvaluationError(result)) return result.code;
  return result === null ? 0 : result;
}

function hasReferenceBoundaryBefore(source: string, position: number): boolean {
  if (position === 0) return true;
  return !/[A-Za-z0-9_.$]/.test(source[position - 1]);
}

function hasReferenceBoundaryAfter(source: string, position: number): boolean {
  return !/[A-Za-z0-9_.]/.test(source[position] ?? "");
}

function shiftedReference(
  original: string,
  rowDelta: number,
  columnDelta: number,
): string {
  const address = parseA1Address(original);
  if (!address) return original;

  const row = address.rowAbsolute ? address.row : address.row + rowDelta;
  const column = address.columnAbsolute
    ? address.column
    : address.column + columnDelta;
  if (row < 1 || row > MAX_EXCEL_ROW || column < 1 || column > MAX_EXCEL_COLUMN) {
    return "#REF!";
  }

  const formatted = formatA1Address({ ...address, row, column });
  if (!formatted) return "#REF!";
  const originalColumn = /^\$?([A-Za-z]+)/.exec(original)?.[1] ?? "A";
  return originalColumn === originalColumn.toLowerCase()
    ? formatted.replace(/[A-Z]+/, (label) => label.toLowerCase())
    : formatted;
}

function shiftedColumnPart(
  anchor: string,
  label: string,
  columnDelta: number,
): string | null {
  const column = columnLabelToNumber(label);
  if (column === null) return null;
  if (anchor === "$") return `$${label}`;
  const shifted = column + columnDelta;
  if (shifted < 1 || shifted > MAX_EXCEL_COLUMN) return "#REF!";
  const formatted = columnNumberToLabel(shifted);
  if (formatted === null) return "#REF!";
  return label === label.toLowerCase() ? formatted.toLowerCase() : formatted;
}

function shiftedRowPart(anchor: string, digits: string, rowDelta: number): string | null {
  const row = Number(digits);
  if (!Number.isSafeInteger(row) || row > MAX_EXCEL_ROW) return null;
  if (anchor === "$") return `$${digits}`;
  const shifted = row + rowDelta;
  return shifted >= 1 && shifted <= MAX_EXCEL_ROW ? String(shifted) : "#REF!";
}

function wholeRangeShift(
  source: string,
  rowDelta: number,
  columnDelta: number,
): { length: number; text: string } | null {
  const columnRange = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.$])/.exec(source);
  if (columnRange) {
    const first = shiftedColumnPart(columnRange[1], columnRange[2], columnDelta);
    const second = shiftedColumnPart(columnRange[3], columnRange[4], columnDelta);
    if (first === null || second === null) return null;
    return { length: columnRange[0].length, text: `${first}:${second}` };
  }
  const rowRange = /^(\$?)([1-9]\d*):(\$?)([1-9]\d*)(?![A-Za-z0-9_.$])/.exec(source);
  if (rowRange) {
    const first = shiftedRowPart(rowRange[1], rowRange[2], rowDelta);
    const second = shiftedRowPart(rowRange[3], rowRange[4], rowDelta);
    if (first === null || second === null) return null;
    return { length: rowRange[0].length, text: `${first}:${second}` };
  }
  return null;
}

function copyDoubleQuotedString(source: string, start: number): number {
  let position = start + 1;
  while (position < source.length) {
    if (source[position] !== '"') position += 1;
    else if (source[position + 1] === '"') position += 2;
    else return position + 1;
  }
  return source.length;
}

function quotedSheetPrefixEnd(source: string, start: number): number | null {
  let position = start + 1;
  while (position < source.length) {
    if (source[position] !== "'") position += 1;
    else if (source[position + 1] === "'") position += 2;
    else return source[position + 1] === "!" ? position + 2 : null;
  }
  return null;
}

/**
 * Shift relative A1 references for copy/fill operations. Absolute `$` anchors,
 * sheet prefixes, string literals, and structured-reference brackets are preserved.
 */
export function shiftFormulaReferences(
  formula: string,
  rowDelta: number,
  columnDelta: number,
): string {
  const rowShift = Number.isFinite(rowDelta) ? Math.trunc(rowDelta) : 0;
  const columnShift = Number.isFinite(columnDelta) ? Math.trunc(columnDelta) : 0;
  let output = "";
  let position = 0;

  while (position < formula.length) {
    if (formula[position] === '"') {
      const end = copyDoubleQuotedString(formula, position);
      output += formula.slice(position, end);
      position = end;
      continue;
    }

    if (formula[position] === "[") {
      const end = formula.indexOf("]", position + 1);
      if (end >= 0) {
        output += formula.slice(position, end + 1);
        position = end + 1;
        continue;
      }
    }

    if (formula[position] === "'" && hasReferenceBoundaryBefore(formula, position)) {
      const prefixEnd = quotedSheetPrefixEnd(formula, position);
      if (prefixEnd !== null) {
        const match = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*/.exec(formula.slice(prefixEnd));
        if (match && hasReferenceBoundaryAfter(formula, prefixEnd + match[0].length)) {
          output += formula.slice(position, prefixEnd);
          output += shiftedReference(match[0], rowShift, columnShift);
          position = prefixEnd + match[0].length;
          continue;
        }
        const whole = wholeRangeShift(formula.slice(prefixEnd), rowShift, columnShift);
        if (whole) {
          output += formula.slice(position, prefixEnd) + whole.text;
          position = prefixEnd + whole.length;
          continue;
        }
      }
    }

    if (hasReferenceBoundaryBefore(formula, position)) {
      const qualified = /^([A-Za-z_\\][A-Za-z0-9_.]*!)(\$?[A-Za-z]{1,3}\$?[1-9]\d*)/.exec(
        formula.slice(position),
      );
      if (qualified && hasReferenceBoundaryAfter(formula, position + qualified[0].length)) {
        output += qualified[1] + shiftedReference(qualified[2], rowShift, columnShift);
        position += qualified[0].length;
        continue;
      }

      const qualifiedPrefix = /^[A-Za-z_\\][A-Za-z0-9_.]*!/.exec(formula.slice(position));
      if (qualifiedPrefix) {
        const whole = wholeRangeShift(
          formula.slice(position + qualifiedPrefix[0].length),
          rowShift,
          columnShift,
        );
        if (whole) {
          output += qualifiedPrefix[0] + whole.text;
          position += qualifiedPrefix[0].length + whole.length;
          continue;
        }
      }

      const reference = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*/.exec(formula.slice(position));
      if (reference && hasReferenceBoundaryAfter(formula, position + reference[0].length)) {
        const after = position + reference[0].length;
        if (nextNonWhitespace(formula, after) !== "(") {
          output += shiftedReference(reference[0], rowShift, columnShift);
          position = after;
          continue;
        }
      }

      const whole = wholeRangeShift(formula.slice(position), rowShift, columnShift);
      if (whole) {
        output += whole.text;
        position += whole.length;
        continue;
      }
    }

    output += formula[position];
    position += 1;
  }

  return output;
}
