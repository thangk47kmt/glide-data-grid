import { parseCsv, stringifyCsv, type CsvParseOptions, type CsvStringifyOptions } from "./csv.js";
import { parseSnapshot, rowsFromSnapshot, serializeSnapshot, snapshotFromRows, type SnapshotLimits, type SnapshotMetadata } from "./persistence.js";
import type { CellInput, SpreadsheetColumn, SpreadsheetModel, SpreadsheetModelOptions } from "./model.js";
import { SpreadsheetModel as Model } from "./model.js";

export interface WorkbookCsvExportOptions extends Omit<CsvStringifyOptions, "formulaValues"> {
    readonly includeHeaders?: boolean;
    /** Inputs preserve formulas; values exports the model's computed results. */
    readonly valueMode?: "inputs" | "values";
}

export interface WorkbookCsvImportOptions extends CsvParseOptions {
    readonly hasHeaders?: boolean;
    /** Supply metadata to preserve ids/titles/types; otherwise deterministic metadata is generated. */
    readonly columns?: readonly SpreadsheetColumn[];
}

export interface WorkbookPayload {
    readonly columns: readonly SpreadsheetColumn[];
    readonly rows: readonly (readonly CellInput[])[];
    readonly rowCount: number;
}

export interface WorkbookSnapshotOptions {
    readonly formats?: SnapshotMetadata;
    readonly validation?: SnapshotMetadata;
    readonly extensions?: SnapshotMetadata;
}

export type WorkbookIoErrorCode = "invalid-columns" | "column-count-mismatch" | "row-too-wide";

export class WorkbookIoError extends Error {
    public readonly code: WorkbookIoErrorCode;
    public readonly path: string;

    public constructor(code: WorkbookIoErrorCode, message: string, path: string) {
        super(message);
        this.name = "WorkbookIoError";
        this.code = code;
        this.path = path;
    }
}

function modelRows(model: SpreadsheetModel, computed: boolean): CellInput[][] {
    return Array.from({ length: model.rowCount }, (_, row) =>
        model.columns.map((_, col) => computed ? model.getValue(col, row) : model.getInput(col, row))
    );
}

/** Materializes raw inputs for structural transforms or external persistence. */
export function workbookPayloadFromModel(model: SpreadsheetModel): WorkbookPayload {
    return { columns: model.columns, rows: modelRows(model, false), rowCount: model.rowCount };
}

/** Reconstructs a model from a normalized workbook payload. */
export function modelFromWorkbookPayload(payload: WorkbookPayload, options?: SpreadsheetModelOptions): SpreadsheetModel {
    return new Model(payload.columns, payload.rowCount, payload.rows, options);
}

/** Exports model inputs or computed values, optionally preceded by column-title headers. */
export function exportModelToCsv(model: SpreadsheetModel, options: WorkbookCsvExportOptions = {}): string {
    const includeHeaders = options.includeHeaders ?? true;
    const valueMode = options.valueMode ?? "inputs";
    const inputRows = modelRows(model, false);
    const needsComputedRows = valueMode === "values" || options.formulaPolicy === "values";
    const computedRows = needsComputedRows ? modelRows(model, true) : undefined;
    const headers: CellInput[] = model.columns.map(column => column.title);
    const rows = valueMode === "values" ? computedRows as CellInput[][] : inputRows;
    const outputRows = includeHeaders ? [headers, ...rows] : rows;
    const formulaPolicy = valueMode === "values" ? "raw" : options.formulaPolicy ?? "raw";
    const formulaValues = formulaPolicy === "values"
        ? (includeHeaders ? [headers, ...computedRows as CellInput[][]] : computedRows as CellInput[][])
        : undefined;
    const { includeHeaders: _includeHeaders, valueMode: _valueMode, ...csvOptions } = options;
    return stringifyCsv(outputRows, {
        ...csvOptions,
        formulaPolicy,
        ...(formulaValues === undefined ? {} : { formulaValues }),
    });
}

function generatedColumns(count: number, headers: readonly CellInput[] | undefined): SpreadsheetColumn[] {
    return Array.from({ length: count }, (_, index) => ({
        id: `column-${index + 1}`,
        title: headers?.[index] === undefined || typeof headers[index] !== "string" ? `Column ${index + 1}` : headers[index] as string,
    }));
}

function validateProvidedColumns(columns: readonly SpreadsheetColumn[]): SpreadsheetColumn[] {
    if (!Array.isArray(columns)) throw new WorkbookIoError("invalid-columns", "columns must be an array", "$.columns");
    if (columns.length === 0) return [];
    const ids = new Set<string>();
    const normalizedIds = new Set<string>();
    return columns.map((column, index) => {
        if (typeof column.id !== "string" || column.id.length === 0 || typeof column.title !== "string") throw new WorkbookIoError("invalid-columns", "Column requires non-empty id and title", `$.columns[${index}]`);
        if (ids.has(column.id) || normalizedIds.has(column.id.toLowerCase())) throw new WorkbookIoError("invalid-columns", "Column ids must be unique", `$.columns[${index}].id`);
        ids.add(column.id);
        normalizedIds.add(column.id.toLowerCase());
        if (column.width !== undefined && (typeof column.width !== "number" || !Number.isFinite(column.width) || column.width < 0)) throw new WorkbookIoError("invalid-columns", "Column width must be a finite non-negative number", `$.columns[${index}].width`);
        if (column.type !== undefined && column.type !== "text" && column.type !== "number" && column.type !== "boolean") throw new WorkbookIoError("invalid-columns", "Column type is invalid", `$.columns[${index}].type`);
        return { id: column.id, title: column.title, ...(column.width === undefined ? {} : { width: column.width }), ...(column.type === undefined ? {} : { type: column.type }) };
    });
}

/** Imports CSV into normalized rows and column metadata; formulas remain raw strings unless rejected by policy. */
export function importCsvToWorkbook(text: string, options: WorkbookCsvImportOptions = {}): WorkbookPayload {
    const hasHeaders = options.hasHeaders ?? true;
    const parsed = parseCsv(text, options);
    const header = hasHeaders && parsed.length > 0 ? parsed[0] : undefined;
    const dataRows = hasHeaders && parsed.length > 0 ? parsed.slice(1) : parsed;
    const columns = options.columns === undefined ? generatedColumns(parsed.reduce((width, row) => Math.max(width, row.length), 0), header) : validateProvidedColumns(options.columns);
    const width = columns.length;
    if (options.columns !== undefined && header !== undefined && header.length !== width) throw new WorkbookIoError("column-count-mismatch", "CSV header column count does not match supplied columns", "$.header");
    if (options.columns !== undefined && header === undefined && dataRows.some(row => row.length > width)) throw new WorkbookIoError("row-too-wide", "CSV row has more fields than supplied columns", "$.rows");
    if (options.columns === undefined && width === 0) return { columns, rows: [], rowCount: 0 };
    const rows = dataRows.map((row, rowIndex) => {
        if (row.length > width) throw new WorkbookIoError("row-too-wide", "CSV row has more fields than columns", `$.rows[${rowIndex}]`);
        return Array.from({ length: width }, (_, columnIndex) => row[columnIndex] ?? null);
    });
    return { columns, rows, rowCount: rows.length };
}

/** Serializes only model raw inputs into a sparse, versioned snapshot. */
export function serializeModelSnapshot(model: SpreadsheetModel, options: WorkbookSnapshotOptions = {}): string {
    const snapshot = snapshotFromRows(model.columns, modelRows(model, false));
    return serializeSnapshot({ ...snapshot, ...(options.formats === undefined ? {} : { formats: options.formats }), ...(options.validation === undefined ? {} : { validation: options.validation }), ...(options.extensions === undefined ? {} : { extensions: options.extensions }) });
}

/** Restores a new SpreadsheetModel from a validated v1 snapshot; computed caches are rebuilt lazily. */
export function restoreModelSnapshot(json: string, limits: SnapshotLimits = {}): SpreadsheetModel {
    const snapshot = parseSnapshot(json, limits);
    return new Model(snapshot.columns, snapshot.rowCount, rowsFromSnapshot(snapshot));
}
