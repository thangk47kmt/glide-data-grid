import { isFormulaError } from "./formula.js";
import type { CellInput } from "./model.js";

export type CsvEmptyValue = "null" | "string";
export type CsvFormulaParsePolicy = "text" | "reject";

export interface CsvParseOptions {
    readonly delimiter?: string;
    readonly inferTypes?: boolean;
    /** Empty fields become null only when explicitly requested. */
    readonly emptyValue?: CsvEmptyValue;
    /** Parsing never compiles formulas; text preserves a formula-looking string, reject refuses it. */
    readonly formulaPolicy?: CsvFormulaParsePolicy;
    readonly maxRows?: number;
    readonly maxColumns?: number;
    readonly maxCellLength?: number;
}

export type CsvFormulaStringifyPolicy = "raw" | "values";

export interface CsvStringifyOptions {
    readonly delimiter?: string;
    /** RFC4180 CRLF is the default; LF is also available for Unix-oriented consumers. */
    readonly lineEnding?: "\r\n" | "\n";
    readonly quoteAll?: boolean;
    readonly formulaPolicy?: CsvFormulaStringifyPolicy;
    /** Required for formula cells when formulaPolicy is values. */
    readonly formulaValues?: readonly (readonly CellInput[])[];
    /** Prefixes string values beginning with =, +, -, or @. */
    readonly safeForExcel?: boolean;
}

export type CsvParseErrorCode =
    | "invalid-delimiter"
    | "unexpected-quote"
    | "unterminated-quote"
    | "invalid-after-quote"
    | "formula-not-allowed"
    | "max-rows"
    | "max-columns"
    | "max-cell-length";

/** A malformed CSV or configured input limit, with one-based row/column location. */
export class CsvParseError extends Error {
    public readonly code: CsvParseErrorCode;
    public readonly row: number;
    public readonly column: number;

    public constructor(code: CsvParseErrorCode, message: string, row: number, column: number) {
        super(message);
        this.name = "CsvParseError";
        this.code = code;
        this.row = row;
        this.column = column;
    }
}

const NUMBER_LITERAL = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

function delimiterOf(delimiter: string | undefined): string {
    const value = delimiter ?? ",";
    if (value.length !== 1 || value === '"' || value === "\r" || value === "\n") {
        throw new TypeError("CSV delimiter must be one non-quote, non-line-ending character");
    }
    return value;
}

function parseCell(raw: string, options: CsvParseOptions, row: number, column: number): CellInput {
    if (raw.startsWith("=") && options.formulaPolicy === "reject") {
        throw new CsvParseError("formula-not-allowed", "Formula-looking CSV values are not allowed", row, column);
    }
    if (raw === "" && options.emptyValue === "null") return null;
    if (options.inferTypes !== true) return raw;
    if (raw === "TRUE") return true;
    if (raw === "FALSE") return false;
    if (NUMBER_LITERAL.test(raw)) {
        const number = Number(raw);
        if (Number.isFinite(number)) return number;
    }
    return raw;
}

function assertCellLength(length: number, options: CsvParseOptions, row: number, column: number): void {
    if (options.maxCellLength !== undefined && (!Number.isInteger(options.maxCellLength) || options.maxCellLength < 0)) {
        throw new TypeError("maxCellLength must be a non-negative integer");
    }
    if (options.maxCellLength !== undefined && length > options.maxCellLength) {
        throw new CsvParseError("max-cell-length", "CSV cell exceeds maxCellLength", row, column);
    }
}

function assertRowLimit(rows: number, options: CsvParseOptions, row: number, column: number): void {
    if (options.maxRows !== undefined && (!Number.isInteger(options.maxRows) || options.maxRows < 0)) throw new TypeError("maxRows must be a non-negative integer");
    if (options.maxRows !== undefined && rows > options.maxRows) throw new CsvParseError("max-rows", "CSV exceeds maxRows", row, column);
}

/**
 * Parses RFC4180-style CSV. Rows are returned without an extra row for a
 * terminal line ending; trailing empty fields are preserved. Empty input is [].
 * Parse errors and limits throw CsvParseError with one-based row and column.
 */
export function parseCsv(text: string, options: CsvParseOptions = {}): CellInput[][] {
    const delimiter = delimiterOf(options.delimiter);
    const rows: CellInput[][] = [];
    let row: CellInput[] = [];
    let field = "";
    let inQuotes = false;
    let quoteClosed = false;
    let fieldAtStart = true;
    let recordStarted = false;
    let rowNumber = 1;
    let columnNumber = 1;

    const pushField = (): void => {
        if (options.maxColumns !== undefined && (!Number.isInteger(options.maxColumns) || options.maxColumns < 0)) throw new TypeError("maxColumns must be a non-negative integer");
        if (options.maxColumns !== undefined && row.length + 1 > options.maxColumns) throw new CsvParseError("max-columns", "CSV exceeds maxColumns", rowNumber, columnNumber);
        assertCellLength(field.length, options, rowNumber, columnNumber);
        row.push(parseCell(field, options, rowNumber, columnNumber));
        field = "";
        fieldAtStart = true;
        quoteClosed = false;
        columnNumber++;
    };

    const pushRow = (): void => {
        pushField();
        assertRowLimit(rows.length + 1, options, rowNumber, columnNumber - 1);
        rows.push(row);
        row = [];
        rowNumber++;
        columnNumber = 1;
        recordStarted = false;
    };

    for (let index = 0; index < text.length;) {
        const char = text[index];
        if (inQuotes) {
            if (char === '"') {
                if (text[index + 1] === '"') {
                    field += '"';
                    index += 2;
                    assertCellLength(field.length, options, rowNumber, columnNumber);
                } else {
                    inQuotes = false;
                    quoteClosed = true;
                    index++;
                }
            } else {
                field += char;
                index++;
                assertCellLength(field.length, options, rowNumber, columnNumber);
            }
            continue;
        }

        if (quoteClosed) {
            if (char === delimiter) {
                pushField();
                recordStarted = true;
                index++;
                continue;
            }
            if (char === "\r" || char === "\n") {
                if (char === "\r" && text[index + 1] === "\n") index++;
                pushRow();
                index++;
                continue;
            }
            throw new CsvParseError("invalid-after-quote", "Only delimiter or line ending may follow a closing quote", rowNumber, columnNumber);
        }

        if (fieldAtStart && char === '"') {
            inQuotes = true;
            recordStarted = true;
            index++;
            continue;
        }
        if (char === '"') throw new CsvParseError("unexpected-quote", "Quotes must start a CSV field", rowNumber, columnNumber);
        if (char === delimiter) {
            pushField();
            recordStarted = true;
            index++;
            continue;
        }
        if (char === "\r" || char === "\n") {
            if (char === "\r" && text[index + 1] === "\n") index++;
            pushRow();
            index++;
            continue;
        }
        field += char;
        fieldAtStart = false;
        recordStarted = true;
        index++;
        assertCellLength(field.length, options, rowNumber, columnNumber);
    }

    if (inQuotes) throw new CsvParseError("unterminated-quote", "CSV ended inside a quoted field", rowNumber, columnNumber);
    if (recordStarted || row.length > 0 || field.length > 0 || quoteClosed) pushRow();
    return rows;
}

function outputValue(value: CellInput): string {
    if (isFormulaError(value)) return value.code;
    return value === null ? "" : String(value);
}

function quoteCsv(value: string, delimiter: string, quoteAll: boolean): string {
    if (quoteAll || value.includes(delimiter) || value.includes('"') || value.includes("\r") || value.includes("\n")) {
        return `"${value.replaceAll('"', '""')}"`;
    }
    return value;
}

/**
 * Serializes rows with RFC4180 escaping and no final line ending. Formula
 * inputs are preserved by default; formulaPolicy=values uses caller-provided
 * formulaValues. safeForExcel prefixes dangerous leading formula characters.
 */
export function stringifyCsv(rows: readonly (readonly CellInput[])[], options: CsvStringifyOptions = {}): string {
    const delimiter = delimiterOf(options.delimiter);
    const lineEnding = options.lineEnding ?? "\r\n";
    const formulaPolicy = options.formulaPolicy ?? "raw";
    if (formulaPolicy === "values" && options.formulaValues === undefined) throw new TypeError("formulaValues are required when formulaPolicy is values");
    return rows.map((row, rowIndex) => row.map((input, columnIndex) => {
        let value = input;
        if (formulaPolicy === "values" && typeof input === "string" && input.startsWith("=")) {
            const provided = options.formulaValues?.[rowIndex]?.[columnIndex];
            if (provided === undefined) throw new TypeError(`Missing formula value at row ${rowIndex + 1}, column ${columnIndex + 1}`);
            value = provided;
        }
        let text = outputValue(value);
        if (options.safeForExcel === true && typeof value === "string" && /^[=+\-@]/.test(text)) text = `'${text}`;
        return quoteCsv(text, delimiter, options.quoteAll === true);
    }).join(delimiter)).join(lineEnding);
}
