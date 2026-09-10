import type { FormulaFunctionContext, FormulaFunctionRegistry, FormulaFunctionArgument, FormulaRangeArgument } from "./function-registry.js";

export type FormulaErrorCode = "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#N/A" | "#CYCLE!";

export interface FormulaError {
    readonly kind: "error";
    readonly code: FormulaErrorCode;
    readonly message?: string;
}

export type FormulaValue = string | number | boolean | null | FormulaError;

export function formulaError(code: FormulaErrorCode, message?: string): FormulaError {
    return { kind: "error", code, message };
}

export function isFormulaError(value: unknown): value is FormulaError {
    return typeof value === "object" && value !== null && "kind" in value && value.kind === "error";
}

export function columnIndexToName(index: number): string {
    if (!Number.isInteger(index) || index < 0) throw new Error("Column index must be a non-negative integer");
    let result = "";
    for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
        result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
    }
    return result;
}

export function columnNameToIndex(name: string): number | undefined {
    const normalized = name.replaceAll("$", "").toUpperCase();
    if (!/^[A-Z]+$/.test(normalized)) return undefined;
    let result = 0;
    for (const char of normalized) result = result * 26 + char.charCodeAt(0) - 64;
    return result - 1;
}

export interface CellReference {
    readonly kind: "cell";
    readonly col: number;
    readonly row: number;
    readonly absoluteCol: boolean;
    readonly absoluteRow: boolean;
}

export interface CurrentRowReference {
    readonly kind: "current-row";
    readonly columnName: string;
}

/**
 * A caption-based cell reference. Rows are one-based in the formula text, so
 * `[Quantity]3` addresses the Quantity column on the third data row. Unlike
 * `[@Quantity]`, this reference is deliberately tied to the explicit row and
 * is useful when a formula is authored by clicking a cell on another row.
 */
export interface CaptionCellReference {
    readonly kind: "caption-cell";
    readonly columnName: string;
    /** Zero-based row used internally by the evaluator. */
    readonly row: number;
}

export interface ColumnReference {
    readonly kind: "column";
    readonly columnName: string;
}

type ReferenceNode = CellReference | CurrentRowReference | CaptionCellReference | ColumnReference;
type RangeCellReference = CellReference | CaptionCellReference;
type AstNode =
    | { readonly kind: "literal"; readonly value: FormulaValue }
    | ReferenceNode
    | { readonly kind: "range"; readonly from: RangeCellReference; readonly to: RangeCellReference }
    | { readonly kind: "unary"; readonly operator: "+" | "-"; readonly value: AstNode }
    | { readonly kind: "binary"; readonly operator: string; readonly left: AstNode; readonly right: AstNode }
    | { readonly kind: "call"; readonly name: string; readonly args: readonly AstNode[] };

type TokenKind = "number" | "string" | "error" | "identifier" | "structured" | "structured-cell" | "operator" | "left" | "right" | "comma" | "colon" | "eof";
interface Token {
    readonly kind: TokenKind;
    readonly value: string;
    /** One-based row suffix for a structured-cell token, when present. */
    readonly row?: number;
}

function tokenize(input: string): Token[] {
    const result: Token[] = [];
    let index = 0;
    while (index < input.length) {
        const char = input[index];
        if (/\s/.test(char)) {
            index++;
            continue;
        }
        if (char === '"') {
            let value = "";
            index++;
            let closed = false;
            while (index < input.length) {
                if (input[index] === '"' && input[index + 1] === '"') {
                    value += '"';
                    index += 2;
                } else if (input[index] === '"') {
                    index++;
                    closed = true;
                    break;
                } else {
                    value += input[index++];
                }
            }
            if (!closed) throw new Error("Missing closing double quote");
            result.push({ kind: "string", value });
            continue;
        }
        if (char === "[") {
            const end = input.indexOf("]", index + 1);
            if (end === -1) throw new Error("Missing closing ] in structured reference");
            const value = input.slice(index + 1, end).trim();
            const rowMatch = input.slice(end + 1).match(/^\$?(\d+)/);
            if (rowMatch !== null) {
                const row = Number(rowMatch[1]);
                if (row < 1) throw new Error("A structured cell reference row must be at least 1");
                result.push({ kind: "structured-cell", value, row: row - 1 });
                index = end + 1 + rowMatch[0].length;
            } else {
                result.push({ kind: "structured", value });
                index = end + 1;
            }
            continue;
        }
        if (char === "#") {
            const errorCode = ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#CYCLE!", "#N/A"].find(code => input.startsWith(code, index));
            if (errorCode === undefined) throw new Error(`Unknown formula error '${input.slice(index)}'`);
            result.push({ kind: "error", value: errorCode });
            index += errorCode.length;
            continue;
        }
        const numberMatch = input.slice(index).match(/^(?:\d+(?:\.\d*)?|\.\d+)/);
        if (numberMatch !== null) {
            result.push({ kind: "number", value: numberMatch[0] });
            index += numberMatch[0].length;
            continue;
        }
        const identifierMatch = input.slice(index).match(/^(?:\$?[A-Za-z]+\$?\d+|[A-Za-z_][A-Za-z0-9_.]*)/);
        if (identifierMatch !== null) {
            result.push({ kind: "identifier", value: identifierMatch[0] });
            index += identifierMatch[0].length;
            continue;
        }
        const pair = input.slice(index, index + 2);
        if (["<=", ">=", "<>", "=="].includes(pair)) {
            result.push({ kind: "operator", value: pair });
            index += 2;
            continue;
        }
        const punctuation: Record<string, TokenKind> = {
            "(": "left",
            ")": "right",
            ",": "comma",
            ":": "colon",
        };
        if (punctuation[char] !== undefined) {
            result.push({ kind: punctuation[char], value: char });
            index++;
            continue;
        }
        if ("+-*/^&=<>".includes(char)) {
            result.push({ kind: "operator", value: char });
            index++;
            continue;
        }
        throw new Error(`Unexpected character '${char}'`);
    }
    result.push({ kind: "eof", value: "" });
    return result;
}

function parseCellReference(value: string): CellReference | undefined {
    const match = value.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
    if (match === null) return undefined;
    const col = columnNameToIndex(match[2]);
    if (col === undefined || Number(match[4]) < 1) return undefined;
    return {
        kind: "cell",
        col,
        row: Number(match[4]) - 1,
        absoluteCol: match[1] === "$",
        absoluteRow: match[3] === "$",
    };
}

class Parser {
    private index = 0;
    public constructor(private readonly tokens: readonly Token[]) {}

    public parse(): AstNode {
        const result = this.parseComparison();
        if (this.peek().kind !== "eof") throw new Error(`Unexpected token '${this.peek().value}'`);
        return result;
    }

    private peek(): Token {
        return this.tokens[this.index];
    }

    private take(): Token {
        return this.tokens[this.index++];
    }

    private parseComparison(): AstNode {
        let node = this.parseConcat();
        while (this.peek().kind === "operator" && ["=", "==", "<>", "<", ">", "<=", ">="].includes(this.peek().value)) {
            const operator = this.take().value;
            node = { kind: "binary", operator, left: node, right: this.parseConcat() };
        }
        return node;
    }

    private parseConcat(): AstNode {
        let node = this.parseAdditive();
        while (this.peek().kind === "operator" && this.peek().value === "&") {
            node = { kind: "binary", operator: this.take().value, left: node, right: this.parseAdditive() };
        }
        return node;
    }

    private parseAdditive(): AstNode {
        let node = this.parseMultiplicative();
        while (this.peek().kind === "operator" && ["+", "-"].includes(this.peek().value)) {
            node = { kind: "binary", operator: this.take().value, left: node, right: this.parseMultiplicative() };
        }
        return node;
    }

    private parseMultiplicative(): AstNode {
        let node = this.parseUnary();
        while (this.peek().kind === "operator" && ["*", "/"].includes(this.peek().value)) {
            node = { kind: "binary", operator: this.take().value, left: node, right: this.parseUnary() };
        }
        return node;
    }

    private parsePower(): AstNode {
        let node = this.parsePrimary();
        if (this.peek().kind === "operator" && this.peek().value === "^") {
            // The right side is unary so expressions such as 2^-2 remain valid,
            // while unary minus has lower precedence than exponentiation (-2^2).
            node = { kind: "binary", operator: this.take().value, left: node, right: this.parseUnary() };
        }
        return node;
    }

    private parseUnary(): AstNode {
        if (this.peek().kind === "operator" && ["+", "-"].includes(this.peek().value)) {
            return { kind: "unary", operator: this.take().value as "+" | "-", value: this.parseUnary() };
        }
        return this.parsePower();
    }

    private parsePrimary(): AstNode {
        const token = this.take();
        if (token.kind === "number") return { kind: "literal", value: Number(token.value) };
        if (token.kind === "string") return { kind: "literal", value: token.value };
        if (token.kind === "error") return { kind: "literal", value: formulaError(token.value as FormulaErrorCode) };
        if (token.kind === "structured") {
            const name = token.value.trim();
            if (name.length === 0 || (name.startsWith("@") && name.slice(1).trim().length === 0)) {
                throw new Error("Structured reference must include a column name");
            }
            return name.startsWith("@")
                ? { kind: "current-row", columnName: name.slice(1).trim() }
                : { kind: "column", columnName: name };
        }
        if (token.kind === "structured-cell") {
            const reference: CaptionCellReference = {
                kind: "caption-cell",
                columnName: token.value,
                row: token.row as number,
            };
            if (this.peek().kind === "colon") {
                this.take();
                const to = this.parseRangeReference(this.take());
                if (to === undefined) throw new Error("A range must end with a cell reference");
                return { kind: "range", from: reference, to };
            }
            return reference;
        }
        if (token.kind === "left") {
            const node = this.parseComparison();
            if (this.take().kind !== "right") throw new Error("Missing closing parenthesis");
            return node;
        }
        if (token.kind !== "identifier") throw new Error(`Unexpected token '${token.value}'`);

        const upper = token.value.toUpperCase();
        if (upper === "TRUE" || upper === "FALSE") return { kind: "literal", value: upper === "TRUE" };
        if (this.peek().kind === "left") {
            this.take();
            const args: AstNode[] = [];
            if (this.peek().kind !== "right") {
                do {
                    args.push(this.parseComparison());
                    if (this.peek().kind !== "comma") break;
                    this.take();
                } while (true);
            }
            if (this.take().kind !== "right") throw new Error("Missing closing parenthesis");
            return { kind: "call", name: upper, args };
        }

        const reference = parseCellReference(token.value);
        if (reference === undefined) throw new Error(`Unknown name '${token.value}'`);
        if (this.peek().kind === "colon") {
            this.take();
            const to = this.parseRangeReference(this.take());
            if (to === undefined) throw new Error("A range must end with a cell reference");
            return { kind: "range", from: reference, to };
        }
        return reference;
    }

    private parseRangeReference(token: Token): RangeCellReference | undefined {
        if (token.kind === "identifier") return parseCellReference(token.value);
        if (token.kind === "structured-cell" && token.row !== undefined) {
            return { kind: "caption-cell", columnName: token.value, row: token.row };
        }
        return undefined;
    }
}

export interface FormulaDependencies {
    readonly cells: ReadonlySet<string>;
    readonly columns: ReadonlySet<number>;
}

export interface FormulaContext {
    readonly currentRow: number;
    readonly rowCount: number;
    readonly resolveColumn: (name: string) => number | undefined;
    readonly getCellValue: (col: number, row: number) => FormulaValue;
    readonly functionRegistry?: FormulaFunctionRegistry;
}

export interface CompileFormulaOptions {
    readonly functionRegistry?: FormulaFunctionRegistry;
}

interface RangeResult {
    readonly kind: "range-result";
    readonly values: readonly FormulaValue[];
    readonly rows: number;
    readonly columns: number;
}
type EvaluationValue = FormulaValue | RangeResult;

export interface CompiledFormula {
    readonly source: string;
    readonly evaluate: (context: FormulaContext) => FormulaValue;
    readonly dependencies: (context: Pick<FormulaContext, "currentRow" | "rowCount" | "resolveColumn">) => FormulaDependencies;
}

function asNumber(value: FormulaValue): number | FormulaError {
    if (isFormulaError(value)) return value;
    if (typeof value === "number") return value;
    if (typeof value === "boolean") return value ? 1 : 0;
    if (value === null || value === "") return 0;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : formulaError("#VALUE!", `Cannot convert '${value}' to a number`);
}

function scalar(value: EvaluationValue): FormulaValue {
    return typeof value === "object" && value !== null && "kind" in value && value.kind === "range-result"
        ? formulaError("#VALUE!", "A range cannot be used as a scalar")
        : value;
}

function flatten(values: readonly EvaluationValue[]): FormulaValue[] {
    return values.flatMap(value =>
        typeof value === "object" && value !== null && "kind" in value && value.kind === "range-result"
            ? value.values
            : [value]
    );
}

function asText(value: FormulaValue): string | FormulaError {
    if (isFormulaError(value)) return value;
    if (value === null) return "";
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    return String(value);
}

function scalarArgument(args: readonly EvaluationValue[], index: number): FormulaValue {
    return scalar(args[index] ?? formulaError("#VALUE!", "Missing argument"));
}

function argumentNumber(args: readonly EvaluationValue[], index: number): number | FormulaError {
    return asNumber(scalarArgument(args, index));
}

function argumentCountError(name: string, expected: string): FormulaError {
    return formulaError("#VALUE!", `${name} expects ${expected}`);
}

// Dates use an Excel-compatible serial representation: integer day 0 is
// 1899-12-30, so date values remain ordinary FormulaValue numbers and can be
// compared or used in arithmetic without adding a Date value to the public API.
const DATE_EPOCH_MS = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

function dateToSerial(date: Date): number {
    return Math.round((date.getTime() - DATE_EPOCH_MS) / DAY_MS);
}

function serialToDate(value: FormulaValue): Date | FormulaError {
    const serial = asNumber(value);
    if (typeof serial !== "number" || !Number.isFinite(serial)) return typeof serial === "number" ? formulaError("#VALUE!", "Invalid date") : serial;
    const date = new Date(DATE_EPOCH_MS + serial * DAY_MS);
    return Number.isNaN(date.getTime()) ? formulaError("#VALUE!", "Invalid date") : date;
}

function evaluateTextCall(name: string, args: readonly EvaluationValue[]): FormulaValue | undefined {
    if (name === "CONCAT") {
        const values = flatten(args);
        const textParts: string[] = [];
        for (const value of values) {
            const text = asText(value);
            if (isFormulaError(text)) return text;
            textParts.push(text);
        }
        return textParts.join("");
    }

    if (!["LEN", "LOWER", "UPPER", "TRIM", "LEFT", "RIGHT", "MID"].includes(name)) return undefined;

    const value = scalarArgument(args, 0);
    const text = asText(value);
    if (isFormulaError(text)) return text;
    switch (name) {
        case "LEN":
            return args.length === 1 ? text.length : argumentCountError(name, "one argument");
        case "LOWER":
            return args.length === 1 ? text.toLowerCase() : argumentCountError(name, "one argument");
        case "UPPER":
            return args.length === 1 ? text.toUpperCase() : argumentCountError(name, "one argument");
        case "TRIM":
            return args.length === 1 ? text.trim().replace(/\s+/g, " ") : argumentCountError(name, "one argument");
        case "LEFT":
        case "RIGHT": {
            if (args.length < 1 || args.length > 2) return argumentCountError(name, "one or two arguments");
            const count = args.length === 1 ? 1 : argumentNumber(args, 1);
            if (isFormulaError(count)) return count;
            if (!Number.isFinite(count) || count < 0) return formulaError("#VALUE!", "Character count must be non-negative");
            const length = Math.trunc(count);
            return name === "LEFT" ? text.slice(0, length) : length === 0 ? "" : text.slice(-length);
        }
        case "MID": {
            if (args.length !== 3) return argumentCountError(name, "three arguments");
            const start = argumentNumber(args, 1);
            const count = argumentNumber(args, 2);
            if (isFormulaError(start)) return start;
            if (isFormulaError(count)) return count;
            if (!Number.isFinite(start) || !Number.isFinite(count) || start < 1 || count < 0) return formulaError("#VALUE!", "Invalid MID arguments");
            return text.slice(Math.trunc(start) - 1, Math.trunc(start) - 1 + Math.trunc(count));
        }
        default:
            return undefined;
    }
}

function evaluateDateCall(name: string, args: readonly EvaluationValue[]): FormulaValue | undefined {
    if (name === "TODAY") {
        if (args.length !== 0) return argumentCountError(name, "no arguments");
        const now = new Date();
        // Use local calendar components, but construct the serial in UTC so
        // the result is an integer independent of the machine's time zone.
        return dateToSerial(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())));
    }
    if (name === "DATE") {
        if (args.length !== 3) return argumentCountError(name, "three arguments");
        const year = argumentNumber(args, 0);
        const month = argumentNumber(args, 1);
        const day = argumentNumber(args, 2);
        if (isFormulaError(year)) return year;
        if (isFormulaError(month)) return month;
        if (isFormulaError(day)) return day;
        if (![year, month, day].every(Number.isInteger) || year < 0 || year > 9999) return formulaError("#VALUE!", "Invalid DATE arguments");
        const date = new Date(DATE_EPOCH_MS);
        date.setUTCFullYear(year, month - 1, day);
        return Number.isNaN(date.getTime()) ? formulaError("#VALUE!", "Invalid date") : dateToSerial(date);
    }
    if (["YEAR", "MONTH", "DAY"].includes(name)) {
        if (args.length !== 1) return argumentCountError(name, "one argument");
        const date = serialToDate(scalarArgument(args, 0));
        if (!(date instanceof Date)) return date;
        if (name === "YEAR") return date.getUTCFullYear();
        if (name === "MONTH") return date.getUTCMonth() + 1;
        return date.getUTCDate();
    }
    return undefined;
}

function isValidFormulaValue(value: unknown): value is FormulaValue {
    if (value === null || typeof value === "string" || typeof value === "boolean") return true;
    if (typeof value === "number") return Number.isFinite(value);
    if (!isFormulaError(value)) return false;
    return ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#CYCLE!"].includes(value.code);
}

function customFunctionArgument(value: EvaluationValue): FormulaFunctionArgument {
    if (typeof value === "object" && value !== null && "kind" in value && value.kind === "range-result") {
        const range: FormulaRangeArgument = { kind: "range", values: value.values, rows: value.rows, columns: value.columns };
        return range;
    }
    return value;
}

function rangeArgument(args: readonly EvaluationValue[], index: number, name: string): RangeResult | FormulaError {
    const value = args[index];
    if (value === undefined) return formulaError("#VALUE!", `${name} expects a range argument`);
    if (isFormulaError(value)) return value;
    if (value === null || typeof value !== "object" || value.kind !== "range-result") return formulaError("#VALUE!", `${name} expects a range argument`);
    return value;
}

function indexArgument(args: readonly EvaluationValue[], index: number, name: string): number | FormulaError {
    const value = args[index];
    if (value === undefined) return formulaError("#VALUE!", `${name} expects an index argument`);
    const number = asNumber(scalar(value));
    if (isFormulaError(number) || !Number.isInteger(number) || number < 1) return formulaError("#VALUE!", `${name} indexes must be positive integers`);
    return number;
}

function evaluateLookupCall(name: string, args: readonly EvaluationValue[]): FormulaValue | undefined {
    if (name === "INDEX") {
        if (args.length < 2 || args.length > 3) return argumentCountError(name, "two or three arguments");
        const range = rangeArgument(args, 0, name);
        if (isFormulaError(range)) return range;
        const row = indexArgument(args, 1, name);
        if (isFormulaError(row)) return row;
        const column = args.length === 3 ? indexArgument(args, 2, name) : 1;
        if (isFormulaError(column)) return column;
        if (row > range.rows || column > range.columns) return formulaError("#REF!", `${name} index is outside the range`);
        const offset = (row - 1) * range.columns + column - 1;
        return offset >= range.values.length ? formulaError("#REF!", `${name} index is outside the range`) : range.values[offset];
    }

    if (name === "MATCH") {
        if (args.length < 2 || args.length > 3) return argumentCountError(name, "two or three arguments");
        const range = rangeArgument(args, 1, name);
        if (isFormulaError(range)) return range;
        if (range.rows !== 1 && range.columns !== 1) return formulaError("#VALUE!", "MATCH requires a one-dimensional range");
        if (args.length === 3) {
            const matchType = asNumber(scalar(args[2] as EvaluationValue));
            if (isFormulaError(matchType)) return matchType;
            if (matchType !== 0) return formulaError("#N/A", "MATCH only supports exact matchType 0");
        }
        const lookup = scalar(args[0] as EvaluationValue);
        if (isFormulaError(lookup)) return lookup;
        for (let index = 0; index < range.values.length; index++) {
            const value = range.values[index];
            if (!isFormulaError(value) && value === lookup) return index + 1;
        }
        return formulaError("#N/A", "MATCH did not find the lookup value");
    }

    if (name === "XLOOKUP") {
        if (args.length < 3 || args.length > 4) return argumentCountError(name, "three or four arguments");
        const lookupRange = rangeArgument(args, 1, name);
        const returnRange = rangeArgument(args, 2, name);
        if (isFormulaError(lookupRange)) return lookupRange;
        if (isFormulaError(returnRange)) return returnRange;
        if (lookupRange.rows !== 1 && lookupRange.columns !== 1) return formulaError("#VALUE!", "XLOOKUP requires a one-dimensional lookup range");
        if (returnRange.rows !== 1 && returnRange.columns !== 1) return formulaError("#VALUE!", "XLOOKUP requires a one-dimensional return range");
        if (lookupRange.rows !== returnRange.rows || lookupRange.columns !== returnRange.columns) {
            return formulaError("#VALUE!", "XLOOKUP ranges must have matching dimensions");
        }
        const lookup = scalar(args[0] as EvaluationValue);
        if (isFormulaError(lookup)) return lookup;
        for (let index = 0; index < lookupRange.values.length; index++) {
            const value = lookupRange.values[index];
            if (!isFormulaError(value) && value === lookup) return index >= returnRange.values.length ? formulaError("#REF!") : returnRange.values[index];
        }
        if (args.length === 4) return scalar(args[3] as EvaluationValue);
        return formulaError("#N/A", "XLOOKUP did not find the lookup value");
    }
    return undefined;
}

function consumePromiseLike(value: object): void {
    // Promise.resolve adopts native promises and thenables alike. Attaching a
    // rejection handler is important even though evaluation remains sync:
    // Promise.reject() must not become an unhandled rejection later.
    try {
        void Promise.resolve(value as PromiseLike<unknown>).catch(() => undefined);
    } catch {
        // A hostile thenable may throw while being assimilated; it is still an
        // invalid synchronous formula result and needs no further propagation.
    }
}

function evaluateCustomCall(name: string, args: readonly EvaluationValue[], context: FormulaContext): FormulaValue | undefined {
    const fn = context.functionRegistry?.get(name);
    if (fn === undefined) return undefined;
    try {
        const result = fn(args.map(customFunctionArgument), context as FormulaFunctionContext);
        if (typeof result === "object" && result !== null && "then" in result && typeof result.then === "function") {
            consumePromiseLike(result);
            return formulaError("#VALUE!", `Custom function '${name}' returned a Promise`);
        }
        return isValidFormulaValue(result) ? result : formulaError("#VALUE!", `Custom function '${name}' returned an invalid value`);
    } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return formulaError("#VALUE!", `Custom function '${name}' failed: ${message}`);
    }
}

function evaluateCall(name: string, args: readonly EvaluationValue[], context: FormulaContext): FormulaValue {
    const customResult = evaluateCustomCall(name, args, context);
    if (customResult !== undefined) return customResult;
    const lookupResult = evaluateLookupCall(name, args);
    if (lookupResult !== undefined) return lookupResult;
    const textResult = evaluateTextCall(name, args);
    if (textResult !== undefined) return textResult;
    const dateResult = evaluateDateCall(name, args);
    if (dateResult !== undefined) return dateResult;
    const values = flatten(args);
    const firstError = values.find(isFormulaError);
    if (firstError !== undefined) return firstError;
    const numbers = values.map(asNumber).filter((value): value is number => typeof value === "number");
    switch (name) {
        case "SUM": return numbers.reduce((sum, value) => sum + value, 0);
        case "AVERAGE": return numbers.length === 0 ? formulaError("#DIV/0!", "AVERAGE has no numeric values to average; divide by zero") : numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
        case "MIN": return numbers.length === 0 ? 0 : Math.min(...numbers);
        case "MAX": return numbers.length === 0 ? 0 : Math.max(...numbers);
        case "COUNT": return numbers.length;
        case "COUNTA": return values.filter(value => value !== null && value !== "").length;
        case "ABS": return numbers.length === 1 ? Math.abs(numbers[0]) : formulaError("#VALUE!", "ABS expects one argument");
        case "ROUND": return numbers.length >= 1 ? Number(numbers[0].toFixed(numbers[1] ?? 0)) : formulaError("#VALUE!");
        case "AND": return values.every(Boolean);
        case "OR": return values.some(Boolean);
        case "NOT": return values.length === 1 ? !Boolean(values[0]) : formulaError("#VALUE!");
        default: return formulaError("#NAME?", `Unknown function '${name}'`);
    }
}

function evaluateNode(node: AstNode, context: FormulaContext): EvaluationValue {
    switch (node.kind) {
        case "literal": return node.value;
        case "cell": return context.getCellValue(node.col, node.row);
        case "caption-cell": {
            const col = context.resolveColumn(node.columnName);
            return col === undefined
                ? formulaError("#REF!", `Unknown column '${node.columnName}'`)
                : context.getCellValue(col, node.row);
        }
        case "current-row": {
            const col = context.resolveColumn(node.columnName);
            return col === undefined ? formulaError("#REF!", `Unknown column '${node.columnName}'`) : context.getCellValue(col, context.currentRow);
        }
        case "column": {
            const col = context.resolveColumn(node.columnName);
            if (col === undefined) return formulaError("#REF!", `Unknown column '${node.columnName}'`);
            return {
                kind: "range-result",
                values: Array.from({ length: context.rowCount }, (_, row) => context.getCellValue(col, row)),
                rows: context.rowCount,
                columns: 1,
            };
        }
        case "range": {
            const from = resolveRangeCellReference(node.from, context);
            const to = resolveRangeCellReference(node.to, context);
            if (isFormulaError(from)) return from;
            if (isFormulaError(to)) return to;
            const values: FormulaValue[] = [];
            for (let row = Math.min(from.row, to.row); row <= Math.max(from.row, to.row); row++) {
                for (let col = Math.min(from.col, to.col); col <= Math.max(from.col, to.col); col++) {
                    values.push(context.getCellValue(col, row));
                }
            }
            return {
                kind: "range-result",
                values,
                rows: Math.abs(from.row - to.row) + 1,
                columns: Math.abs(from.col - to.col) + 1,
            };
        }
        case "unary": {
            const value = asNumber(scalar(evaluateNode(node.value, context)));
            return typeof value === "number" ? (node.operator === "-" ? -value : value) : value;
        }
        case "call": {
            if (node.name === "IF" && context.functionRegistry?.has("IF") !== true) {
                const condition = scalar(evaluateNode(node.args[0] ?? { kind: "literal", value: false }, context));
                if (isFormulaError(condition)) return condition;
                return scalar(evaluateNode(Boolean(condition) ? node.args[1] ?? { kind: "literal", value: true } : node.args[2] ?? { kind: "literal", value: false }, context));
            }
            return evaluateCall(node.name, node.args.map(arg => evaluateNode(arg, context)), context);
        }
        case "binary": {
            const left = scalar(evaluateNode(node.left, context));
            const right = scalar(evaluateNode(node.right, context));
            if (isFormulaError(left)) return left;
            if (isFormulaError(right)) return right;
            if (node.operator === "&") return `${left ?? ""}${right ?? ""}`;
            if (["=", "==", "<>", "<", ">", "<=", ">="].includes(node.operator)) {
                if (node.operator === "=" || node.operator === "==") return left === right;
                if (node.operator === "<>") return left !== right;
                const comparableLeft = left ?? 0;
                const comparableRight = right ?? 0;
                if (node.operator === "<") return comparableLeft < comparableRight;
                if (node.operator === ">") return comparableLeft > comparableRight;
                if (node.operator === "<=") return comparableLeft <= comparableRight;
                return comparableLeft >= comparableRight;
            }
            const leftNumber = asNumber(left);
            const rightNumber = asNumber(right);
            if (isFormulaError(leftNumber)) return leftNumber;
            if (isFormulaError(rightNumber)) return rightNumber;
            if (node.operator === "+") return leftNumber + rightNumber;
            if (node.operator === "-") return leftNumber - rightNumber;
            if (node.operator === "*") return leftNumber * rightNumber;
            if (node.operator === "/") return rightNumber === 0 ? formulaError("#DIV/0!", "Cannot divide by zero") : leftNumber / rightNumber;
            if (node.operator === "^") return leftNumber ** rightNumber;
            return formulaError("#VALUE!");
        }
    }
}

function resolveRangeCellReference(reference: RangeCellReference, context: Pick<FormulaContext, "resolveColumn">): { readonly col: number; readonly row: number } | FormulaError {
    if (reference.kind === "cell") return reference;
    const col = context.resolveColumn(reference.columnName);
    return col === undefined
        ? formulaError("#REF!", `Unknown column '${reference.columnName}'`)
        : { col, row: reference.row };
}

function collectDependencies(node: AstNode, context: Pick<FormulaContext, "currentRow" | "rowCount" | "resolveColumn">, cells: Set<string>, columns: Set<number>): void {
    switch (node.kind) {
        case "cell": cells.add(`${node.col}:${node.row}`); break;
        case "caption-cell": {
            const col = context.resolveColumn(node.columnName);
            if (col !== undefined) cells.add(`${col}:${node.row}`);
            break;
        }
        case "current-row": {
            const col = context.resolveColumn(node.columnName);
            if (col !== undefined) cells.add(`${col}:${context.currentRow}`);
            break;
        }
        case "column": {
            const col = context.resolveColumn(node.columnName);
            if (col !== undefined) columns.add(col);
            break;
        }
        case "range":
            const from = resolveRangeCellReference(node.from, context);
            const to = resolveRangeCellReference(node.to, context);
            if (isFormulaError(from) || isFormulaError(to)) break;
            for (let row = Math.min(from.row, to.row); row <= Math.max(from.row, to.row); row++) {
                for (let col = Math.min(from.col, to.col); col <= Math.max(from.col, to.col); col++) cells.add(`${col}:${row}`);
            }
            break;
        case "unary": collectDependencies(node.value, context, cells, columns); break;
        case "binary":
            collectDependencies(node.left, context, cells, columns);
            collectDependencies(node.right, context, cells, columns);
            break;
        case "call": node.args.forEach(arg => collectDependencies(arg, context, cells, columns)); break;
        case "literal": break;
    }
}

export function compileFormula(source: string, options: CompileFormulaOptions = {}): CompiledFormula {
    const expression = source.startsWith("=") ? source.slice(1) : source;
    const ast = new Parser(tokenize(expression)).parse();
    return {
        source,
        evaluate: context => scalar(evaluateNode(ast, context.functionRegistry === undefined && options.functionRegistry !== undefined
            ? { ...context, functionRegistry: options.functionRegistry }
            : context)),
        dependencies: context => {
            const cells = new Set<string>();
            const columns = new Set<number>();
            collectDependencies(ast, context, cells, columns);
            return { cells, columns };
        },
    };
}

export function translateFormula(source: string, columnDelta: number, rowDelta: number): string {
    const referencePattern = /(\$?)([A-Za-z]+)(\$?)(\d+)/y;
    let result = "";
    let index = 0;
    while (index < source.length) {
        const char = source[index];
        if (char === '"') {
            // Copy string literals verbatim. A doubled quote is an escaped
            // quote, not the end of the literal, so references inside either
            // form are never translated.
            const start = index++;
            while (index < source.length) {
                if (source[index] === '"' && source[index + 1] === '"') {
                    index += 2;
                } else if (source[index] === '"') {
                    index++;
                    break;
                } else {
                    index++;
                }
            }
            result += source.slice(start, index);
            continue;
        }
        if (char === "[") {
            // Structured references are names, not A1 references. Keep the
            // complete bracketed token unchanged.
            const end = source.indexOf("]", index + 1);
            if (end === -1) {
                result += source.slice(index);
                break;
            }
            result += source.slice(index, end + 1);
            index = end + 1;
            continue;
        }

        referencePattern.lastIndex = index;
        const match = referencePattern.exec(source);
        const previous = index === 0 ? "" : source[index - 1];
        const end = match === null ? index : index + match[0].length;
        const next = end < source.length ? source[end] : "";
        const hasTokenBoundaries = !/[A-Za-z0-9_.]/.test(previous) && !/[A-Za-z0-9_.]/.test(next) && next !== "(";
        if (match !== null && hasTokenBoundaries) {
            const [, colAbsolute, colName, rowAbsolute, rowText] = match;
            const col = columnNameToIndex(colName);
            if (col !== undefined) {
                const translatedCol = colAbsolute === "$" ? col : col + columnDelta;
                const translatedRow = rowAbsolute === "$" ? Number(rowText) : Number(rowText) + rowDelta;
                result += translatedCol < 0 || translatedRow < 1
                    ? "#REF!"
                    : `${colAbsolute}${columnIndexToName(translatedCol)}${rowAbsolute}${translatedRow}`;
                index = end;
                continue;
            }
        }
        result += char;
        index++;
    }
    return result;
}
