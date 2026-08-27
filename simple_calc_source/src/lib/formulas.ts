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
  | "sheet"
  | "error"
  | "operator"
  | "leftParen"
  | "rightParen"
  | "comma"
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
      ",": "comma",
      ";": "comma",
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

    const reference = this.parseReference();
    if (reference) {
      if (this.match("colon")) {
        const end = this.parseReference();
        if (!end) throw new FormulaParseError("#REF!");
        return { kind: "range", start: reference, end };
      }
      return reference;
    }

    if (this.current.kind === "identifier") {
      const name = String(this.advance().value).toUpperCase();
      if (name === "TRUE" || name === "FALSE") {
        return { kind: "literal", value: name === "TRUE" };
      }
      return { kind: "error", value: "#NAME?" };
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
      } while (this.match("comma"));
      if (!this.match("rightParen")) throw new FormulaParseError();
    }
    return { kind: "call", name, arguments: args };
  }

  private parseReference(): ReferenceNode | null {
    let sheet: string | undefined;
    if (
      ["sheet", "identifier", "cell"].includes(this.current.kind) &&
      this.next.kind === "bang"
    ) {
      sheet = String(this.advance().value);
      this.advance();
    }

    if (this.current.kind !== "cell") {
      if (sheet !== undefined) throw new FormulaParseError("#REF!");
      return null;
    }

    const address = parseA1Address(String(this.advance().value));
    if (!address) throw new FormulaParseError("#REF!");
    return { kind: "reference", sheet, address };
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
}

type EvaluationValue = EvaluationScalar | EvaluationRange;

interface EvaluationContext {
  resolver: FormulaResolver;
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

function resolveRange(
  range: RangeNode,
  currentSheetId: string,
  context: EvaluationContext,
): EvaluationValue {
  const startSheet = range.start.sheet ?? currentSheetId;
  const endSheet = range.end.sheet ?? startSheet;
  if (startSheet !== endSheet) return evaluationError("#REF!");

  const firstRow = Math.min(range.start.address.row, range.end.address.row);
  const lastRow = Math.max(range.start.address.row, range.end.address.row);
  const firstColumn = Math.min(range.start.address.column, range.end.address.column);
  const lastColumn = Math.max(range.start.address.column, range.end.address.column);
  const rowCount = lastRow - firstRow + 1;
  const columnCount = lastColumn - firstColumn + 1;
  if (rowCount > MAX_RANGE_CELLS / columnCount) return evaluationError("#VALUE!");

  const values: EvaluationScalar[] = [];
  for (let row = firstRow; row <= lastRow; row += 1) {
    for (let column = firstColumn; column <= lastColumn; column += 1) {
      values.push(
        resolveReference(
          {
            kind: "reference",
            sheet: startSheet,
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
  if (isEvaluationRange(value)) return value;
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
  name: "SUMIF" | "SUMIFS" | "COUNTIF" | "COUNTIFS" | "AVERAGEIF" | "AVERAGEIFS",
  args: EvaluationValue[],
): EvaluationValue {
  const isPlural = name.endsWith("S");
  const isCount = name.startsWith("COUNT");
  const isAverage = name.startsWith("AVERAGE");

  let resultRange: RectangularValues | undefined;
  let pairStart = 0;
  if (name === "SUMIFS" || name === "AVERAGEIFS") {
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
      sum += aggregate;
      count += 1;
      if (!Number.isFinite(sum)) return evaluationError("#NUM!");
    }
  }

  if (isCount) return count;
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

function evaluateCall(
  call: CallNode,
  currentSheetId: string,
  context: EvaluationContext,
  depth: number,
): EvaluationValue {
  const name = call.name.replace(/^_xlfn\./i, "").toUpperCase();

  if (name === "IF") {
    if (call.arguments.length < 2 || call.arguments.length > 3) {
      return evaluationError("#VALUE!");
    }
    const condition = toBoolean(
      evaluateNode(call.arguments[0], currentSheetId, context, depth + 1),
    );
    if (isEvaluationError(condition)) return condition;
    if (condition) return evaluateNode(call.arguments[1], currentSheetId, context, depth + 1);
    return call.arguments[2]
      ? evaluateNode(call.arguments[2], currentSheetId, context, depth + 1)
      : false;
  }

  if (name === "IFERROR") {
    if (call.arguments.length < 1 || call.arguments.length > 2) {
      return evaluationError("#VALUE!");
    }
    const value = evaluateNode(call.arguments[0], currentSheetId, context, depth + 1);
    if (!isEvaluationError(value)) return value;
    return call.arguments[1]
      ? evaluateNode(call.arguments[1], currentSheetId, context, depth + 1)
      : "";
  }

  if (name === "IFS") {
    if (call.arguments.length < 2 || call.arguments.length % 2 !== 0) {
      return evaluationError("#VALUE!");
    }
    for (let index = 0; index < call.arguments.length; index += 2) {
      const condition = toBoolean(
        evaluateNode(call.arguments[index], currentSheetId, context, depth + 1),
      );
      if (isEvaluationError(condition)) return condition;
      if (condition) {
        return evaluateNode(call.arguments[index + 1], currentSheetId, context, depth + 1);
      }
    }
    return evaluationError("#N/A");
  }

  if (name === "AND" || name === "OR" || name === "XOR") {
    return evaluateLogicalArguments(name, call.arguments, currentSheetId, context, depth);
  }

  if (name === "NOT") {
    if (call.arguments.length !== 1) return evaluationError("#VALUE!");
    const value = toBoolean(
      evaluateNode(call.arguments[0], currentSheetId, context, depth + 1),
    );
    return isEvaluationError(value) ? value : !value;
  }

  if (name === "TODAY" || name === "NOW") {
    if (call.arguments.length !== 0) return evaluationError("#VALUE!");
    const serial = localDateToExcelSerial(context.now);
    return name === "TODAY" ? Math.floor(serial) : serial;
  }

  const values = call.arguments.map((argument) =>
    evaluateNode(argument, currentSheetId, context, depth + 1),
  );

  if (
    name === "SUMIF" ||
    name === "SUMIFS" ||
    name === "COUNTIF" ||
    name === "COUNTIFS" ||
    name === "AVERAGEIF" ||
    name === "AVERAGEIFS"
  ) {
    return conditionalAggregate(name, values);
  }

  if (name === "INDEX" || name === "MATCH" || name === "VLOOKUP" || name === "HLOOKUP") {
    return evaluateLookupFunction(name, values);
  }

  if (["SUM", "AVERAGE", "MIN", "MAX"].includes(name)) {
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
  }

  if (name === "PRODUCT" || name === "MEDIAN") {
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
  }

  if (name === "LARGE" || name === "SMALL") {
    if (values.length !== 2) return evaluationError("#VALUE!");
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
  }

  if (name === "RANK" || name === "RANK.EQ") {
    if (values.length < 2 || values.length > 3) return evaluationError("#VALUE!");
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
  }

  if (["STDEV", "STDEV.S", "VAR", "VAR.S"].includes(name)) {
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
  }

  if (name === "COUNT") {
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
  }

  if (name === "COUNTA") {
    let count = 0;
    for (const entry of collectValues(values)) {
      if (isEvaluationError(entry.value) && !entry.fromRange) return entry.value;
      if (entry.value !== null) count += 1;
    }
    return count;
  }

  if (name === "ABS") {
    if (values.length !== 1) return evaluationError("#VALUE!");
    const value = toNumber(scalarArgument(values[0]));
    return isEvaluationError(value) ? value : Math.abs(value);
  }

  if (name === "INT" || name === "SQRT") {
    if (values.length !== 1) return evaluationError("#VALUE!");
    const value = toNumber(scalarArgument(values[0]));
    if (isEvaluationError(value)) return value;
    if (name === "SQRT") {
      return value < 0 ? evaluationError("#NUM!") : Math.sqrt(value);
    }
    return Math.floor(value);
  }

  if (name === "MOD" || name === "POWER") {
    if (values.length !== 2) return evaluationError("#VALUE!");
    const left = toNumber(scalarArgument(values[0]));
    const right = toNumber(scalarArgument(values[1]));
    if (isEvaluationError(left)) return left;
    if (isEvaluationError(right)) return right;
    if (name === "MOD") {
      if (right === 0) return evaluationError("#DIV/0!");
      const result = left - right * Math.floor(left / right);
      return Number.isFinite(result)
        ? (Object.is(result, -0) ? 0 : result)
        : evaluationError("#NUM!");
    }
    if (left === 0 && right < 0) return evaluationError("#DIV/0!");
    const result = left ** right;
    return Number.isFinite(result) ? result : evaluationError("#NUM!");
  }

  if (name === "ROUND" || name === "ROUNDUP" || name === "ROUNDDOWN") {
    if (values.length !== 2) return evaluationError("#VALUE!");
    const value = toNumber(scalarArgument(values[0]));
    const digitsValue = toNumber(scalarArgument(values[1]));
    if (isEvaluationError(value)) return value;
    if (isEvaluationError(digitsValue)) return digitsValue;
    const digits = Math.trunc(digitsValue);
    if (name === "ROUNDUP") return directedRound(value, digits, "up");
    if (name === "ROUNDDOWN") return directedRound(value, digits, "down");
    if (Math.abs(digits) > 308) return evaluationError("#NUM!");
    const shifted = decimalShift(Math.abs(value), digits);
    if (!Number.isFinite(shifted)) return evaluationError("#NUM!");
    const rounded = Math.sign(value) * decimalShift(Math.round(shifted), -digits);
    return Number.isFinite(rounded) ? rounded : evaluationError("#NUM!");
  }

  if (name === "LEN" || name === "TRIM" || name === "UPPER" || name === "LOWER" || name === "PROPER") {
    if (values.length !== 1) return evaluationError("#VALUE!");
    const text = toText(scalarArgument(values[0]));
    if (isEvaluationError(text)) return text;
    if (name === "LEN") return Array.from(text).length;
    if (name === "TRIM") return text.replace(/^ +| +$/g, "").replace(/ +/g, " ");
    if (name === "UPPER") return safeTextResult(text.toLocaleUpperCase());
    if (name === "LOWER") return safeTextResult(text.toLocaleLowerCase());
    return safeTextResult(properCase(text));
  }

  if (name === "LEFT" || name === "RIGHT") {
    if (values.length < 1 || values.length > 2) return evaluationError("#VALUE!");
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
  }

  if (name === "MID") {
    if (values.length !== 3) return evaluationError("#VALUE!");
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
  }

  if (name === "CONCAT" || name === "CONCATENATE") {
    let result = "";
    for (const entry of collectValues(values)) {
      const text = toText(entry.value);
      if (isEvaluationError(text)) return text;
      result += text;
      if (result.length > MAX_TEXT_RESULT_LENGTH) return evaluationError("#VALUE!");
    }
    return result;
  }

  if (name === "JOIN" || name === "TEXTJOIN") {
    const minimum = name === "JOIN" ? 2 : 3;
    if (values.length < minimum) return evaluationError("#VALUE!");
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
  }

  if (name === "SUBSTITUTE") {
    if (values.length < 3 || values.length > 4) return evaluationError("#VALUE!");
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
  }

  if (name === "FIND" || name === "SEARCH") {
    if (values.length < 2 || values.length > 3) return evaluationError("#VALUE!");
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
  }

  if (name === "EXACT") {
    if (values.length !== 2) return evaluationError("#VALUE!");
    const left = toText(scalarArgument(values[0]));
    const right = toText(scalarArgument(values[1]));
    if (isEvaluationError(left)) return left;
    if (isEvaluationError(right)) return right;
    return left === right;
  }

  if (name === "VALUE") {
    if (values.length !== 1) return evaluationError("#VALUE!");
    const text = toText(scalarArgument(values[0]));
    if (isEvaluationError(text)) return text;
    const parsed = parseInvariantValue(text);
    return parsed === null ? evaluationError("#VALUE!") : parsed;
  }

  if (name === "DATE") {
    if (values.length !== 3) return evaluationError("#VALUE!");
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
  }

  if (name === "YEAR" || name === "MONTH" || name === "DAY") {
    if (values.length !== 1) return evaluationError("#VALUE!");
    const serial = toNumber(scalarArgument(values[0]));
    if (isEvaluationError(serial)) return serial;
    const date = utcDateFromSerial(serial);
    if (!date) return evaluationError("#NUM!");
    if (name === "YEAR") return date.getUTCFullYear();
    if (name === "MONTH") return date.getUTCMonth() + 1;
    return date.getUTCDate();
  }

  if (name === "EDATE" || name === "EOMONTH") {
    if (values.length !== 2) return evaluationError("#VALUE!");
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
  }

  return evaluationError("#NAME?");
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
  if (node.kind === "range") return resolveRange(node, currentSheetId, context);
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
    return isEvaluationRange(value) ? evaluationError("#VALUE!") : value;
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
 * `TODAY` and `NOW` return Excel serial date numbers.
 */
export function evaluateFormula(
  formula: string,
  currentSheetId: string,
  resolver: FormulaResolver,
): FormulaResult {
  if (typeof formula !== "string" || typeof resolver !== "function") return "#PARSE!";
  const context: EvaluationContext = {
    resolver,
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

      const reference = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*/.exec(formula.slice(position));
      if (reference && hasReferenceBoundaryAfter(formula, position + reference[0].length)) {
        const after = position + reference[0].length;
        if (nextNonWhitespace(formula, after) !== "(") {
          output += shiftedReference(reference[0], rowShift, columnShift);
          position = after;
          continue;
        }
      }
    }

    output += formula[position];
    position += 1;
  }

  return output;
}
