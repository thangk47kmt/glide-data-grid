import type { FormulaErrorCode } from "./formula.js";
import type { CellInput, SpreadsheetColumn } from "./model.js";

export type SnapshotPrimitive = string | number | boolean | null;
export type SnapshotMetadata = Readonly<Record<string, SnapshotPrimitive>>;

export interface SnapshotCell {
    readonly col: number;
    readonly row: number;
    readonly value: CellInput;
}

export interface SpreadsheetSnapshotV1 {
    readonly version: 1;
    readonly columns: readonly SpreadsheetColumn[];
    readonly rowCount: number;
    /** Sparse raw inputs; null and empty-string cells are normally omitted by snapshotFromRows. */
    readonly cells: readonly SnapshotCell[];
    readonly formats?: SnapshotMetadata;
    readonly validation?: SnapshotMetadata;
    readonly extensions?: SnapshotMetadata;
}

export interface SnapshotLimits {
    readonly maxRows?: number;
    readonly maxColumns?: number;
    readonly maxCells?: number;
    readonly maxJsonLength?: number;
    readonly maxMetadataEntries?: number;
    readonly maxMetadataKeyLength?: number;
}

export type SnapshotErrorCode =
    | "invalid-json"
    | "invalid-shape"
    | "unsupported-version"
    | "duplicate-column-id"
    | "duplicate-cell"
    | "out-of-bounds-cell"
    | "row-count-mismatch"
    | "non-finite-number"
    | "invalid-formula-error"
    | "invalid-metadata"
    | "limit-exceeded";

/** Structured persistence error. `path` uses JSONPath-like, one-based arrays are not used. */
export class SnapshotError extends Error {
    public readonly code: SnapshotErrorCode;
    public readonly path: string;

    public constructor(code: SnapshotErrorCode, message: string, path: string) {
        super(message);
        this.name = "SnapshotError";
        this.code = code;
        this.path = path;
    }
}

const FORMULA_ERROR_CODES: readonly FormulaErrorCode[] = ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#CYCLE!"];
const DEFAULT_MAX_METADATA_ENTRIES = 1000;
const DEFAULT_MAX_METADATA_KEY_LENGTH = 256;

function isObject(value: unknown): value is Record<string, unknown> {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}

function fail(code: SnapshotErrorCode, message: string, path: string): never {
    throw new SnapshotError(code, message, path);
}

function validateLimit(value: number | undefined, name: string): void {
    if (value !== undefined && (!Number.isInteger(value) || value < 0)) throw new TypeError(`${name} must be a non-negative integer`);
}

function validateLimits(limits: SnapshotLimits): void {
    validateLimit(limits.maxRows, "maxRows");
    validateLimit(limits.maxColumns, "maxColumns");
    validateLimit(limits.maxCells, "maxCells");
    validateLimit(limits.maxJsonLength, "maxJsonLength");
    validateLimit(limits.maxMetadataEntries, "maxMetadataEntries");
    validateLimit(limits.maxMetadataKeyLength, "maxMetadataKeyLength");
}

function ownKeys(value: Record<string, unknown>): string[] {
    return Object.keys(value);
}

function validateFormulaError(value: Record<string, unknown>, path: string): void {
    if (value.kind !== "error" || typeof value.code !== "string" || !FORMULA_ERROR_CODES.includes(value.code as FormulaErrorCode)) {
        fail("invalid-formula-error", "Invalid FormulaError value", path);
    }
    if (value.message !== undefined && typeof value.message !== "string") fail("invalid-formula-error", "FormulaError message must be a string", `${path}.message`);
    for (const key of ownKeys(value)) if (key !== "kind" && key !== "code" && key !== "message") fail("invalid-formula-error", "Unknown FormulaError property", `${path}.${key}`);
}

function validateCellInput(value: unknown, path: string): asserts value is CellInput {
    if (value === null || typeof value === "string" || typeof value === "boolean") return;
    if (typeof value === "number") {
        if (!Number.isFinite(value)) fail("non-finite-number", "Cell number must be finite", path);
        return;
    }
    if (isObject(value)) {
        validateFormulaError(value, path);
        return;
    }
    fail("invalid-shape", "Cell value must be a string, finite number, boolean, null, or FormulaError", path);
}

function validateMetadata(value: unknown, path: string, limits: SnapshotLimits): SnapshotMetadata {
    if (!isObject(value)) fail("invalid-metadata", "Metadata must be an object", path);
    const maxEntries = limits.maxMetadataEntries ?? DEFAULT_MAX_METADATA_ENTRIES;
    const maxKeyLength = limits.maxMetadataKeyLength ?? DEFAULT_MAX_METADATA_KEY_LENGTH;
    const keys = ownKeys(value);
    if (keys.length > maxEntries) fail("limit-exceeded", "Metadata exceeds maxMetadataEntries", path);
    const result: Record<string, SnapshotPrimitive> = {};
    for (const key of keys) {
        if (key.length > maxKeyLength) fail("limit-exceeded", "Metadata key exceeds maxMetadataKeyLength", `${path}.${key}`);
        const entry = value[key];
        if (entry === null || typeof entry === "string" || typeof entry === "boolean") result[key] = entry;
        else if (typeof entry === "number" && Number.isFinite(entry)) result[key] = entry;
        else if (typeof entry === "number") fail("non-finite-number", "Metadata number must be finite", `${path}.${key}`);
        else fail("invalid-metadata", "Metadata values must be primitive and JSON serializable", `${path}.${key}`);
    }
    return result;
}

function validateColumns(value: unknown, limits: SnapshotLimits): SpreadsheetColumn[] {
    if (!Array.isArray(value)) fail("invalid-shape", "columns must be an array", "$.columns");
    if (limits.maxColumns !== undefined && value.length > limits.maxColumns) fail("limit-exceeded", "Snapshot exceeds maxColumns", "$.columns");
    const columns: SpreadsheetColumn[] = [];
    const ids = new Set<string>();
    const normalizedIds = new Set<string>();
    value.forEach((entry, index) => {
        const path = `$.columns[${index}]`;
        if (!isObject(entry) || typeof entry.id !== "string" || entry.id.length === 0 || typeof entry.title !== "string") fail("invalid-shape", "Column requires string id and title", path);
        if (ids.has(entry.id) || normalizedIds.has(entry.id.toLowerCase())) fail("duplicate-column-id", "Column ids must be unique", `${path}.id`);
        ids.add(entry.id);
        normalizedIds.add(entry.id.toLowerCase());
        if (entry.width !== undefined && (typeof entry.width !== "number" || !Number.isFinite(entry.width) || entry.width < 0)) fail("invalid-shape", "Column width must be a finite non-negative number", `${path}.width`);
        if (entry.type !== undefined && entry.type !== "text" && entry.type !== "number" && entry.type !== "boolean") fail("invalid-shape", "Column type is invalid", `${path}.type`);
        for (const key of ownKeys(entry)) if (!["id", "title", "width", "type"].includes(key)) fail("invalid-shape", "Unknown column property", `${path}.${key}`);
        columns.push({ id: entry.id, title: entry.title, ...(entry.width === undefined ? {} : { width: entry.width }), ...(entry.type === undefined ? {} : { type: entry.type }) });
    });
    return columns;
}

function validateSnapshot(input: unknown, limits: SnapshotLimits = {}): SpreadsheetSnapshotV1 {
    validateLimits(limits);
    if (!isObject(input)) fail("invalid-shape", "Snapshot must be an object", "$");
    if (input.version !== 1) {
        if (typeof input.version === "number" && input.version > 1) fail("unsupported-version", "Snapshot version is newer than supported v1", "$.version");
        fail("unsupported-version", "Only snapshot version 1 is supported", "$.version");
    }
    if (!Number.isInteger(input.rowCount) || (input.rowCount as number) < 0) fail("row-count-mismatch", "rowCount must be a non-negative integer", "$.rowCount");
    const rowCount = input.rowCount as number;
    if (limits.maxRows !== undefined && rowCount > limits.maxRows) fail("limit-exceeded", "Snapshot exceeds maxRows", "$.rowCount");
    const columns = validateColumns(input.columns, limits);
    if (!Array.isArray(input.cells)) fail("invalid-shape", "cells must be an array", "$.cells");
    if (limits.maxCells !== undefined && input.cells.length > limits.maxCells) fail("limit-exceeded", "Snapshot exceeds maxCells", "$.cells");
    const cells: SnapshotCell[] = [];
    const coordinates = new Set<string>();
    input.cells.forEach((entry, index) => {
        const path = `$.cells[${index}]`;
        if (!isObject(entry) || !Number.isInteger(entry.col) || !Number.isInteger(entry.row)) fail("invalid-shape", "Cell requires integer col and row", path);
        const col = entry.col as number;
        const row = entry.row as number;
        if (col < 0 || col >= columns.length || row < 0 || row >= rowCount) fail("out-of-bounds-cell", "Sparse cell is outside snapshot bounds", path);
        const coordinate = `${col}:${row}`;
        if (coordinates.has(coordinate)) fail("duplicate-cell", "Sparse cell coordinates must be unique", path);
        coordinates.add(coordinate);
        validateCellInput(entry.value, `${path}.value`);
        for (const key of ownKeys(entry)) if (!["col", "row", "value"].includes(key)) fail("invalid-shape", "Unknown cell property", `${path}.${key}`);
        cells.push({ col, row, value: entry.value });
    });
    const formats = input.formats === undefined ? undefined : validateMetadata(input.formats, "$.formats", limits);
    const validation = input.validation === undefined ? undefined : validateMetadata(input.validation, "$.validation", limits);
    const extensions = input.extensions === undefined ? undefined : validateMetadata(input.extensions, "$.extensions", limits);
    for (const key of ownKeys(input)) if (!["version", "columns", "rowCount", "cells", "formats", "validation", "extensions"].includes(key)) fail("invalid-shape", "Unknown snapshot property", `$.${key}`);
    cells.sort((left, right) => left.row - right.row || left.col - right.col);
    return { version: 1, columns, rowCount, cells, ...(formats === undefined ? {} : { formats }), ...(validation === undefined ? {} : { validation }), ...(extensions === undefined ? {} : { extensions }) };
}

function sortedMetadata(value: SnapshotMetadata | undefined): SnapshotMetadata | undefined {
    if (value === undefined) return undefined;
    const result: Record<string, SnapshotPrimitive> = {};
    Object.keys(value).sort().forEach(key => { result[key] = value[key]; });
    return result;
}

function canonicalCellValue(value: CellInput): CellInput {
    if (isObject(value) && value.kind === "error") {
        return {
            kind: "error",
            code: value.code as FormulaErrorCode,
            ...(value.message === undefined ? {} : { message: value.message as string }),
        };
    }
    return value;
}

function canonicalSnapshot(snapshot: SpreadsheetSnapshotV1): SpreadsheetSnapshotV1 {
    return {
        version: 1,
        columns: snapshot.columns.map(column => ({ id: column.id, title: column.title, ...(column.width === undefined ? {} : { width: column.width }), ...(column.type === undefined ? {} : { type: column.type }) })),
        rowCount: snapshot.rowCount,
        cells: snapshot.cells.map(cell => ({
            col: cell.col,
            row: cell.row,
            value: canonicalCellValue(cell.value),
        })).sort((left, right) => left.row - right.row || left.col - right.col),
        ...(snapshot.formats === undefined ? {} : { formats: sortedMetadata(snapshot.formats) }),
        ...(snapshot.validation === undefined ? {} : { validation: sortedMetadata(snapshot.validation) }),
        ...(snapshot.extensions === undefined ? {} : { extensions: sortedMetadata(snapshot.extensions) }),
    };
}

/** Validates and serializes a v1 snapshot with stable cell and metadata ordering. */
export function serializeSnapshot(snapshot: SpreadsheetSnapshotV1): string {
    return JSON.stringify(canonicalSnapshot(validateSnapshot(snapshot)));
}

/** Parses JSON and validates its untrusted runtime shape and configured limits. */
export function parseSnapshot(json: string, limits: SnapshotLimits = {}): SpreadsheetSnapshotV1 {
    validateLimits(limits);
    if (limits.maxJsonLength !== undefined && json.length > limits.maxJsonLength) throw new SnapshotError("limit-exceeded", "Snapshot JSON exceeds maxJsonLength", "$<json>");
    let input: unknown;
    try {
        input = JSON.parse(json) as unknown;
    } catch {
        throw new SnapshotError("invalid-json", "Snapshot is not valid JSON", "$<json>");
    }
    return migrateSnapshot(input, limits);
}

/** Migration entry point; v1 is normalized and future versions are rejected explicitly. */
export function migrateSnapshot(input: unknown, limits: SnapshotLimits = {}): SpreadsheetSnapshotV1 {
    return validateSnapshot(input, limits);
}

/** Builds a sparse snapshot, omitting null and empty-string cells. */
export function snapshotFromRows(columns: readonly SpreadsheetColumn[], rows: readonly (readonly CellInput[])[]): SpreadsheetSnapshotV1 {
    if (rows.some(row => row.length > columns.length)) throw new SnapshotError("row-count-mismatch", "A row has more cells than columns", "$.rows");
    const cells: SnapshotCell[] = [];
    rows.forEach((row, rowIndex) => row.forEach((value, colIndex) => {
        if (value !== null && value !== "") cells.push({ col: colIndex, row: rowIndex, value });
    }));
    return validateSnapshot({ version: 1, columns, rowCount: rows.length, cells });
}

/** Expands sparse cells into rows, filling omitted cells with null. */
export function rowsFromSnapshot(snapshot: SpreadsheetSnapshotV1): CellInput[][] {
    const validated = validateSnapshot(snapshot);
    const rows = Array.from({ length: validated.rowCount }, () => Array<CellInput>(validated.columns.length).fill(null));
    validated.cells.forEach(cell => { rows[cell.row][cell.col] = cell.value; });
    return rows;
}
