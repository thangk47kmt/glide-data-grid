import type { CellEdit, CellLocation } from "./history.js";
import type { CellInput, SpreadsheetColumn } from "./model.js";
import type { WorkbookPayload } from "./workbook-io.js";

export type StructuralWorkbookPayload = WorkbookPayload;

export interface StructuralEditLimits {
    readonly maxRows?: number;
    readonly maxColumns?: number;
    readonly maxCells?: number;
}

export interface StructuralEditResult extends WorkbookPayload {
    /** Normalized workbook before the operation, suitable for an atomic undo. */
    readonly beforePayload: WorkbookPayload;
    /** Normalized workbook after the operation. Also exposed directly as `rows`, `columns`, and `rowCount`. */
    readonly afterPayload: WorkbookPayload;
    /** The after payload is also exposed directly as `rows`, `columns`, and `rowCount`. */
    readonly payload: WorkbookPayload;
    /** Deterministic row-major value changes for audit/display; this does not encode schema changes. */
    readonly edits: readonly CellEdit<CellInput>[];
}

export type StructuralAxis = "row" | "column";

export type StructuralOperation =
    | { readonly type: "insert-rows"; readonly index: number; readonly count: number }
    | { readonly type: "delete-rows"; readonly index: number; readonly count: number }
    | { readonly type: "insert-columns"; readonly index: number; readonly columns: readonly SpreadsheetColumn[] }
    | { readonly type: "delete-columns"; readonly index: number; readonly count: number };

interface ReferenceToken {
    readonly start: number;
    readonly end: number;
    readonly col: number;
    readonly row: number;
    readonly colAbsolute: boolean;
    readonly rowAbsolute: boolean;
}

function columnIndexToName(index: number): string {
    let result = "";
    for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
    return result;
}

function columnNameToIndex(name: string): number {
    let result = 0;
    for (const char of name) result = result * 26 + char.toUpperCase().charCodeAt(0) - 64;
    return result - 1;
}

function referenceToken(start: number, end: number, match: RegExpMatchArray): ReferenceToken | undefined {
    const col = columnNameToIndex(match[2]);
    const row = Number(match[4]) - 1;
    if (!Number.isInteger(col) || col < 0 || !Number.isInteger(row) || row < 0) return undefined;
    return { start, end, col, row, colAbsolute: match[1] === "$", rowAbsolute: match[3] === "$" };
}

function scanReferences(source: string): ReferenceToken[] {
    const references: ReferenceToken[] = [];
    const pattern = /^(\$?)([A-Za-z]+)(\$?)(\d+)/;
    let index = 0;
    while (index < source.length) {
        if (source[index] === '"') {
            index++;
            while (index < source.length) {
                if (source[index] === '"' && source[index + 1] === '"') index += 2;
                else if (source[index++] === '"') break;
            }
            continue;
        }
        if (source[index] === "[") {
            const end = source.indexOf("]", index + 1);
            index = end === -1 ? source.length : end + 1;
            continue;
        }
        const match = source.slice(index).match(pattern);
        if (match === null) {
            index++;
            continue;
        }
        const end = index + match[0].length;
        const previous = index === 0 ? "" : source[index - 1];
        const next = source[end] ?? "";
        if (!/[A-Za-z0-9_.]/.test(previous) && !/[A-Za-z0-9_.]/.test(next) && next !== "(") {
            const token = referenceToken(index, end, match);
            if (token !== undefined) references.push(token);
        }
        index = end;
    }
    return references;
}

function mapPoint(value: number, index: number, count: number, operation: "insert" | "delete"): number | undefined {
    if (operation === "insert") return value >= index ? value + count : value;
    if (value < index) return value;
    if (value >= index + count) return value - count;
    return undefined;
}

function mapRange(start: number, end: number, index: number, count: number, operation: "insert" | "delete"): [number, number] | undefined {
    if (operation === "insert") return [mapPoint(start, index, count, operation) as number, mapPoint(end, index, count, operation) as number];
    const low = Math.min(start, end);
    const high = Math.max(start, end);
    const deletionEnd = index + count;
    const hasLeft = low < index;
    const hasRight = high >= deletionEnd;
    if (!hasLeft && !hasRight) return undefined;
    // Both surviving portions are contiguous after deletion. Computing their
    // endpoints directly avoids iterating over potentially billion-cell ranges.
    const mappedLow = hasLeft ? low : Math.max(low, deletionEnd) - count;
    const mappedHigh = hasRight ? high - count : index - 1;
    return start <= end ? [mappedLow, mappedHigh] : [mappedHigh, mappedLow];
}

function formatReference(reference: ReferenceToken, col: number, row: number): string {
    return `${reference.colAbsolute ? "$" : ""}${columnIndexToName(col)}${reference.rowAbsolute ? "$" : ""}${row + 1}`;
}

function rewriteFormula(source: string, axis: StructuralAxis, index: number, count: number, operation: "insert" | "delete"): string {
    const references = scanReferences(source);
    if (references.length === 0) return source;
    let result = "";
    let cursor = 0;
    let referenceIndex = 0;
    while (referenceIndex < references.length) {
        const first = references[referenceIndex];
        result += source.slice(cursor, first.start);
        const second = references[referenceIndex + 1];
        const between = second === undefined ? "" : source.slice(first.end, second.start);
        const isRange = second !== undefined && /^\s*:\s*$/.test(between);
        if (isRange) {
            const mappedRows = axis === "row" ? mapRange(first.row, second.row, index, count, operation) : [first.row, second.row] as [number, number];
            const mappedColumns = axis === "column" ? mapRange(first.col, second.col, index, count, operation) : [first.col, second.col] as [number, number];
            if (mappedRows === undefined || mappedColumns === undefined) result += "#REF!";
            else {
                result += formatReference(first, mappedColumns[0], mappedRows[0]) + between + formatReference(second, mappedColumns[1], mappedRows[1]);
            }
            cursor = second.end;
            referenceIndex += 2;
            continue;
        }
        const mappedCol = axis === "column" ? mapPoint(first.col, index, count, operation) : first.col;
        const mappedRow = axis === "row" ? mapPoint(first.row, index, count, operation) : first.row;
        result += mappedCol === undefined || mappedRow === undefined ? "#REF!" : formatReference(first, mappedCol, mappedRow);
        cursor = first.end;
        referenceIndex++;
    }
    return result + source.slice(cursor);
}

function validateLimit(value: number | undefined, name: string): void {
    if (value !== undefined && (!Number.isInteger(value) || value < 0)) throw new RangeError(`${name} must be a non-negative integer`);
}

function validatePayload(payload: WorkbookPayload, limits: StructuralEditLimits): WorkbookPayload {
    if (payload === null || typeof payload !== "object" || !Array.isArray(payload.columns) || !Array.isArray(payload.rows)) throw new TypeError("A workbook payload with columns and rows is required");
    if (!Number.isInteger(payload.rowCount) || payload.rowCount < 0 || payload.rowCount !== payload.rows.length) throw new RangeError("rowCount must equal rows.length");
    validateLimit(limits.maxRows, "maxRows");
    validateLimit(limits.maxColumns, "maxColumns");
    validateLimit(limits.maxCells, "maxCells");
    if (limits.maxRows !== undefined && payload.rowCount > limits.maxRows) throw new RangeError("Workbook exceeds maxRows");
    if (limits.maxColumns !== undefined && payload.columns.length > limits.maxColumns) throw new RangeError("Workbook exceeds maxColumns");
    if (limits.maxCells !== undefined && payload.rowCount * payload.columns.length > limits.maxCells) throw new RangeError("Workbook exceeds maxCells");
    const ids = new Set<string>();
    const columns = payload.columns.map(column => {
        if (column === null || typeof column !== "object" || typeof column.id !== "string" || column.id.length === 0 || typeof column.title !== "string") throw new TypeError("Columns require non-empty string ids and titles");
        const id = column.id.toLocaleLowerCase();
        if (ids.has(id)) throw new Error(`Duplicate column id '${column.id}'`);
        ids.add(id);
        return { ...column };
    });
    const rows = payload.rows.map(row => {
        if (!Array.isArray(row) || row.length > columns.length) throw new RangeError("Rows cannot exceed the column count");
        return Array.from({ length: columns.length }, (_, col) => row[col] ?? null);
    });
    return { columns, rows, rowCount: payload.rowCount };
}

function result(before: WorkbookPayload, after: WorkbookPayload): StructuralEditResult {
    const edits: CellEdit<CellInput>[] = [];
    const rowCount = Math.max(before.rowCount, after.rowCount);
    const columnCount = Math.max(before.columns.length, after.columns.length);
    for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < columnCount; col++) {
            const beforeValue = before.rows[row]?.[col] ?? null;
            const afterValue = after.rows[row]?.[col] ?? null;
            if (!Object.is(beforeValue, afterValue)) edits.push({ location: [col, row] as CellLocation, before: beforeValue, after: afterValue });
        }
    }
    return { ...after, beforePayload: before, afterPayload: after, payload: after, edits };
}

function transform(payload: WorkbookPayload, operation: StructuralOperation, limits: StructuralEditLimits = {}): StructuralEditResult {
    const before = validatePayload(payload, limits);
    let after: WorkbookPayload;
    if (operation.type === "insert-rows") {
        if (!Number.isInteger(operation.index) || operation.index < 0 || operation.index > before.rowCount) throw new RangeError("Row insertion index is outside the workbook");
        if (!Number.isInteger(operation.count) || operation.count < 1) throw new RangeError("Row count must be a positive integer");
        const rows = [...before.rows.slice(0, operation.index), ...Array.from({ length: operation.count }, () => Array<CellInput>(before.columns.length).fill(null)), ...before.rows.slice(operation.index)];
        after = { columns: before.columns.map(column => ({ ...column })), rows, rowCount: before.rowCount + operation.count };
        if (limits.maxRows !== undefined && after.rowCount > limits.maxRows) throw new RangeError("Workbook exceeds maxRows");
    } else if (operation.type === "delete-rows") {
        if (!Number.isInteger(operation.index) || operation.index < 0 || operation.index > before.rowCount) throw new RangeError("Row deletion index is outside the workbook");
        if (!Number.isInteger(operation.count) || operation.count < 1 || operation.index + operation.count > before.rowCount) throw new RangeError("Row deletion count is outside the workbook");
        const rows = [...before.rows.slice(0, operation.index), ...before.rows.slice(operation.index + operation.count)];
        after = { columns: before.columns.map(column => ({ ...column })), rows, rowCount: before.rowCount - operation.count };
    } else if (operation.type === "insert-columns") {
        if (!Number.isInteger(operation.index) || operation.index < 0 || operation.index > before.columns.length) throw new RangeError("Column insertion index is outside the workbook");
        if (operation.columns.length < 1) throw new RangeError("At least one column is required");
        const inserted = operation.columns.map(column => ({ ...column }));
        const ids = new Set(before.columns.map(column => column.id.toLocaleLowerCase()));
        inserted.forEach(column => {
            if (typeof column.id !== "string" || column.id.length === 0 || typeof column.title !== "string") throw new TypeError("Columns require non-empty string ids and titles");
            const id = column.id.toLocaleLowerCase();
            if (ids.has(id)) throw new Error(`Duplicate column id '${column.id}'`);
            ids.add(id);
        });
        const columns = [...before.columns.slice(0, operation.index), ...inserted, ...before.columns.slice(operation.index)];
        after = {
            columns,
            rows: before.rows.map(row => [...row.slice(0, operation.index), ...Array.from({ length: inserted.length }, () => null), ...row.slice(operation.index)]),
            rowCount: before.rowCount,
        };
        if (limits.maxColumns !== undefined && after.columns.length > limits.maxColumns) throw new RangeError("Workbook exceeds maxColumns");
    } else {
        if (!Number.isInteger(operation.index) || operation.index < 0 || operation.index > before.columns.length) throw new RangeError("Column deletion index is outside the workbook");
        if (!Number.isInteger(operation.count) || operation.count < 1 || operation.index + operation.count > before.columns.length) throw new RangeError("Column deletion count is outside the workbook");
        const columns = [...before.columns.slice(0, operation.index), ...before.columns.slice(operation.index + operation.count)];
        after = { columns, rows: before.rows.map(row => [...row.slice(0, operation.index), ...row.slice(operation.index + operation.count)]), rowCount: before.rowCount };
    }
    const axis: StructuralAxis = operation.type.includes("columns") ? "column" : "row";
    const operationKind = operation.type.startsWith("insert") ? "insert" : "delete";
    const index = operation.index;
    const count = operation.type === "insert-columns" ? operation.columns.length : operation.count;
    const rewrittenRows = after.rows.map(row => row.map(value => typeof value === "string" && value.startsWith("=") ? rewriteFormula(value, axis, index, count, operationKind) : value));
    return result(before, { ...after, rows: rewrittenRows });
}

export function insertRows(payload: WorkbookPayload, index: number, count: number, limits: StructuralEditLimits = {}): StructuralEditResult {
    return transform(payload, { type: "insert-rows", index, count }, limits);
}

export function deleteRows(payload: WorkbookPayload, index: number, count: number, limits: StructuralEditLimits = {}): StructuralEditResult {
    return transform(payload, { type: "delete-rows", index, count }, limits);
}

export function insertColumns(payload: WorkbookPayload, index: number, columns: readonly SpreadsheetColumn[], limits?: StructuralEditLimits): StructuralEditResult;
export function insertColumns(payload: WorkbookPayload, index: number, count: number, columns?: readonly SpreadsheetColumn[], limits?: StructuralEditLimits): StructuralEditResult;
export function insertColumns(payload: WorkbookPayload, index: number, columnsOrCount: readonly SpreadsheetColumn[] | number, columnsOrLimits?: readonly SpreadsheetColumn[] | StructuralEditLimits, limits: StructuralEditLimits = {}): StructuralEditResult {
    const columns = typeof columnsOrCount === "number"
        ? Array.isArray(columnsOrLimits)
            ? columnsOrLimits
            : Array.from({ length: columnsOrCount }, (_, offset) => ({ id: `inserted-column-${index + offset + 1}`, title: `Column ${index + offset + 1}` }))
        : columnsOrCount;
    const selectedLimits: StructuralEditLimits = typeof columnsOrCount === "number" && !Array.isArray(columnsOrLimits)
        ? columnsOrLimits as StructuralEditLimits
        : limits;
    if (typeof columnsOrCount === "number" && Array.isArray(columnsOrLimits) && columnsOrLimits.length !== columnsOrCount) throw new RangeError("Inserted column count does not match column definitions");
    return transform(payload, { type: "insert-columns", index, columns }, selectedLimits);
}

export function deleteColumns(payload: WorkbookPayload, index: number, count: number, limits: StructuralEditLimits = {}): StructuralEditResult {
    return transform(payload, { type: "delete-columns", index, count }, limits);
}

export function applyStructuralEdit(payload: WorkbookPayload, operation: StructuralOperation, limits: StructuralEditLimits = {}): StructuralEditResult {
    return transform(payload, operation, limits);
}
