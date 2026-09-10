import type { FilterOperator, SpreadsheetColumnSearch, SpreadsheetFilter, SpreadsheetSort, SpreadsheetViewOptions } from "./view.js";

/** The persisted schema version for {@link ColumnViewState}. */
export const COLUMN_VIEW_STATE_VERSION = 1 as const;

export interface ColumnViewSearch {
    readonly column: string;
    readonly value: string;
}

export interface ColumnViewFilter {
    readonly column: string;
    readonly operator: FilterOperator;
    readonly value?: string | number | boolean;
}

export interface ColumnViewSort {
    readonly column: string;
    readonly direction: "asc" | "desc";
}

/** Immutable, serializable state for the spreadsheet header controls. */
export interface ColumnViewState {
    readonly globalSearch: string;
    readonly columnSearches: readonly ColumnViewSearch[];
    readonly filters: readonly ColumnViewFilter[];
    /** Sort priority is array order: the first entry is the primary sort. */
    readonly sorts: readonly ColumnViewSort[];
}

export interface ColumnViewStateInput {
    readonly globalSearch?: string;
    readonly columnSearches?: readonly ColumnViewSearch[];
    readonly filters?: readonly ColumnViewFilter[];
    readonly sorts?: readonly ColumnViewSort[];
}

export interface SortToggleOptions {
    /** Keep other sort keys when true; otherwise this becomes the only sort. */
    readonly additive?: boolean;
}

export interface ColumnViewCompileOptions {
    /** A resolver may return a canonical model column id, or undefined for unknown columns. */
    readonly resolveColumn?: (column: string) => string | undefined;
    /** Convenience resolver built from column ids and optional titles. */
    readonly columns?: readonly { readonly id: string; readonly title?: string }[];
}

export interface ColumnViewStateLimits {
    /** Defaults are 1 MiB JSON, 10,000 searches, 10,000 filters, and 100 sorts. */
    readonly maxJsonLength?: number;
    readonly maxSearchLength?: number;
    readonly maxColumns?: number;
    readonly maxFilters?: number;
    readonly maxSorts?: number;
}

export type ColumnViewStateErrorCode =
    | "invalid-json"
    | "invalid-shape"
    | "unsupported-version"
    | "limit-exceeded"
    | "duplicate-column";

export class ColumnViewStateError extends Error {
    public readonly code: ColumnViewStateErrorCode;

    public constructor(code: ColumnViewStateErrorCode, message: string) {
        super(message);
        this.name = "ColumnViewStateError";
        this.code = code;
    }
}

const FILTER_OPERATORS: readonly FilterOperator[] = ["contains", "equals", "not-equals", "gt", "gte", "lt", "lte", "empty", "not-empty"];
const DEFAULT_LIMITS: Required<ColumnViewStateLimits> = {
    maxJsonLength: 1_048_576,
    maxSearchLength: 10_000,
    maxColumns: 10_000,
    maxFilters: 10_000,
    maxSorts: 100,
};
const EMPTY_STATE: ColumnViewState = Object.freeze({ globalSearch: "", columnSearches: Object.freeze([]), filters: Object.freeze([]), sorts: Object.freeze([]) });

function identity(column: string): string {
    return column.trim().toLocaleLowerCase("en-US");
}

function requireColumn(column: unknown): string {
    if (typeof column !== "string" || identity(column) === "") throw new TypeError("Column name must be a non-empty string");
    return column.trim();
}

function requireSearch(value: unknown, name: string): string {
    if (typeof value !== "string") throw new TypeError(`${name} must be a string`);
    return value;
}

function requireFilterOperator(operator: unknown): FilterOperator {
    if (typeof operator !== "string" || !FILTER_OPERATORS.includes(operator as FilterOperator)) throw new TypeError("Unknown filter operator");
    return operator as FilterOperator;
}

function requireFilterValue(value: unknown): string | number | boolean | undefined {
    if (value === undefined || typeof value === "string" || typeof value === "boolean") return value;
    if (typeof value === "number" && Number.isFinite(value)) return value;
    throw new TypeError("Filter value must be a finite number, string, boolean, or undefined");
}

function freezeState(state: ColumnViewState): ColumnViewState {
    const columnSearches = Object.freeze(state.columnSearches.map(entry => Object.freeze({ column: entry.column, value: entry.value })));
    const filters = Object.freeze(state.filters.map(entry => Object.freeze({ column: entry.column, operator: entry.operator, ...(entry.value === undefined ? {} : { value: entry.value }) })));
    const sorts = Object.freeze(state.sorts.map(entry => Object.freeze({ column: entry.column, direction: entry.direction })));
    return Object.freeze({ globalSearch: state.globalSearch, columnSearches, filters, sorts });
}

function replaceByColumn<T extends { readonly column: string }>(entries: readonly T[], entry: T): T[] {
    const wanted = identity(entry.column);
    const result: T[] = [];
    let replaced = false;
    for (const current of entries) {
        if (identity(current.column) !== wanted) result.push(current);
        else if (!replaced) {
            result.push(entry);
            replaced = true;
        }
    }
    if (!replaced) result.push(entry);
    return result;
}

function removeByColumn<T extends { readonly column: string }>(entries: readonly T[], column: string): T[] {
    const wanted = identity(column);
    return entries.filter(entry => identity(entry.column) !== wanted);
}

/** Creates a normalized, frozen state. Duplicate entries are coalesced by column. */
export function createColumnViewState(input: ColumnViewStateInput = {}): ColumnViewState {
    const globalSearch = input.globalSearch === undefined ? "" : requireSearch(input.globalSearch, "globalSearch");
    const columnSearches: ColumnViewSearch[] = [];
    for (const entry of input.columnSearches ?? []) {
        if (entry === null || typeof entry !== "object") throw new TypeError("Invalid column search");
        const column = requireColumn(entry.column);
        const value = requireSearch(entry.value, "column search value");
        if (value !== "") columnSearches.splice(0, columnSearches.length, ...replaceByColumn(columnSearches, { column, value }));
    }
    const filters: ColumnViewFilter[] = [];
    for (const entry of input.filters ?? []) {
        if (entry === null || typeof entry !== "object") throw new TypeError("Invalid column filter");
        const column = requireColumn(entry.column);
        const filter: ColumnViewFilter = { column, operator: requireFilterOperator(entry.operator), ...(entry.value === undefined ? {} : { value: requireFilterValue(entry.value) }) };
        const next = replaceByColumn(filters, filter);
        filters.splice(0, filters.length, ...next);
    }
    const sorts: ColumnViewSort[] = [];
    for (const entry of input.sorts ?? []) {
        if (entry === null || typeof entry !== "object") throw new TypeError("Invalid column sort");
        const column = requireColumn(entry.column);
        if (entry.direction !== "asc" && entry.direction !== "desc") throw new TypeError("Sort direction must be asc or desc");
        const next = replaceByColumn(sorts, { column, direction: entry.direction });
        sorts.splice(0, sorts.length, ...next);
    }
    return freezeState({ globalSearch, columnSearches, filters, sorts });
}

/** An empty frozen state suitable as the initial value for a header. */
export const emptyColumnViewState: ColumnViewState = EMPTY_STATE;

export function setGlobalSearch(state: ColumnViewState, value: string): ColumnViewState {
    return freezeState({ ...state, globalSearch: requireSearch(value, "globalSearch") });
}

export function setColumnSearch(state: ColumnViewState, column: string, value: string): ColumnViewState {
    const name = requireColumn(column);
    const searches = value === "" ? removeByColumn(state.columnSearches, name) : replaceByColumn(state.columnSearches, { column: name, value: requireSearch(value, "column search value") });
    return freezeState({ ...state, columnSearches: searches });
}

export function clearColumnSearch(state: ColumnViewState, column: string): ColumnViewState {
    return freezeState({ ...state, columnSearches: removeByColumn(state.columnSearches, requireColumn(column)) });
}

export function setColumnFilter(state: ColumnViewState, column: string, operator: FilterOperator, value?: string | number | boolean): ColumnViewState {
    const name = requireColumn(column);
    const filter: ColumnViewFilter = { column: name, operator: requireFilterOperator(operator), ...(value === undefined ? {} : { value: requireFilterValue(value) }) };
    return freezeState({ ...state, filters: replaceByColumn(state.filters, filter) });
}

export function clearColumnFilter(state: ColumnViewState, column: string): ColumnViewState {
    return freezeState({ ...state, filters: removeByColumn(state.filters, requireColumn(column)) });
}

/** Clears search, filter, and sorting state for one column. */
export function clearColumn(state: ColumnViewState, column: string): ColumnViewState {
    const name = requireColumn(column);
    return freezeState({ ...state, columnSearches: removeByColumn(state.columnSearches, name), filters: removeByColumn(state.filters, name), sorts: removeByColumn(state.sorts, name) });
}

export function clearAllColumnViewState(): ColumnViewState {
    return EMPTY_STATE;
}

/** Alias for callers that keep view state alongside other clearable state. */
export const clearAll = clearAllColumnViewState;

export function clearSort(state: ColumnViewState, column?: string): ColumnViewState {
    return column === undefined ? freezeState({ ...state, sorts: [] }) : freezeState({ ...state, sorts: removeByColumn(state.sorts, requireColumn(column)) });
}

export function setSort(state: ColumnViewState, column: string, direction: "asc" | "desc", options: SortToggleOptions = {}): ColumnViewState {
    const name = requireColumn(column);
    if (direction !== "asc" && direction !== "desc") throw new TypeError("Sort direction must be asc or desc");
    const next = replaceByColumn(state.sorts, { column: name, direction });
    return freezeState({ ...state, sorts: options.additive ? next : [{ column: name, direction }] });
}

/** Cycles asc -> desc -> none. Additive toggles preserve other sort priorities. */
export function toggleSort(state: ColumnViewState, column: string, options: SortToggleOptions = {}): ColumnViewState {
    const name = requireColumn(column);
    const current = state.sorts.find(sort => identity(sort.column) === identity(name));
    const nextDirection = current?.direction === undefined ? "asc" : current.direction === "asc" ? "desc" : undefined;
    if (nextDirection === undefined) return clearSort(options.additive ? state : freezeState({ ...state, sorts: [] }), name);
    return setSort(state, name, nextDirection, options);
}

function makeResolver(options: ColumnViewCompileOptions): (column: string) => string | undefined {
    if (options.resolveColumn !== undefined) return column => {
        try {
            const resolved = options.resolveColumn?.(column);
            return resolved === undefined ? undefined : requireColumn(resolved);
        } catch {
            return undefined;
        }
    };
    if (options.columns === undefined) return column => column;
    const lookup = new Map<string, string>();
    const ambiguous = new Set<string>();
    const register = (name: string, id: string): void => {
        const normalized = identity(name);
        if (normalized === "" || ambiguous.has(normalized)) return;
        const existing = lookup.get(normalized);
        if (existing === undefined) lookup.set(normalized, id);
        else if (existing !== id) {
            lookup.delete(normalized);
            ambiguous.add(normalized);
        }
    };
    for (const column of options.columns) {
        if (column === null || typeof column !== "object" || typeof column.id !== "string" || identity(column.id) === "") continue;
        register(column.id, column.id.trim());
        if (typeof column.title === "string") register(column.title, column.id.trim());
    }
    return column => {
        const normalized = identity(column);
        return ambiguous.has(normalized) ? undefined : lookup.get(normalized);
    };
}

/** Compiles header state into deterministic view options. With a resolver/schema, unknown column entries are omitted; without one names pass through. */
export function compileColumnViewState(state: ColumnViewState, options: ColumnViewCompileOptions = {}): SpreadsheetViewOptions {
    const resolve = makeResolver(options);
    const columnSearches: SpreadsheetColumnSearch[] = [];
    for (const search of state.columnSearches) {
        const column = resolve(search.column);
        if (column !== undefined) columnSearches.push({ column, value: search.value });
    }
    const filters: SpreadsheetFilter[] = [];
    for (const filter of state.filters) {
        const column = resolve(filter.column);
        if (column !== undefined) filters.push({ column, operator: filter.operator, ...(filter.value === undefined ? {} : { value: filter.value }) });
    }
    const sorts: SpreadsheetSort[] = [];
    for (const sort of state.sorts) {
        const column = resolve(sort.column);
        if (column !== undefined) sorts.push({ column, direction: sort.direction });
    }
    return { search: state.globalSearch, columnSearches, filters, sorts };
}

function validateLimit(value: number | undefined, name: string): void {
    if (value !== undefined && (!Number.isInteger(value) || value < 0)) throw new RangeError(`${name} must be a non-negative integer`);
}

function validateLimits(limits: ColumnViewStateLimits): void {
    validateLimit(limits.maxJsonLength, "maxJsonLength");
    validateLimit(limits.maxSearchLength, "maxSearchLength");
    validateLimit(limits.maxColumns, "maxColumns");
    validateLimit(limits.maxFilters, "maxFilters");
    validateLimit(limits.maxSorts, "maxSorts");
}

function effectiveLimits(limits: ColumnViewStateLimits): Required<ColumnViewStateLimits> {
    return { ...DEFAULT_LIMITS, ...limits };
}

function assertNoDuplicate(entries: readonly { readonly column: string }[], kind: string): void {
    const seen = new Set<string>();
    for (const entry of entries) {
        const name = identity(entry.column);
        if (seen.has(name)) throw new ColumnViewStateError("duplicate-column", `Duplicate ${kind} column`);
        seen.add(name);
    }
}

function assertKeys(value: Record<string, unknown>, allowed: readonly string[], label: string): void {
    for (const key of Object.keys(value)) if (!allowed.includes(key)) throw new ColumnViewStateError("invalid-shape", `Unknown ${label} property: ${key}`);
}

function assertPersistedShape(input: unknown, limits: ColumnViewStateLimits): ColumnViewState {
    validateLimits(limits);
    const bounded = effectiveLimits(limits);
    if (input === null || typeof input !== "object" || Array.isArray(input)) throw new ColumnViewStateError("invalid-shape", "Column view state must be an object");
    const value = input as Record<string, unknown>;
    assertKeys(value, ["version", "globalSearch", "columnSearches", "filters", "sorts"], "state");
    if (value.version !== COLUMN_VIEW_STATE_VERSION) throw new ColumnViewStateError("unsupported-version", "Only column view state version 1 is supported");
    if (typeof value.globalSearch !== "string") throw new ColumnViewStateError("invalid-shape", "globalSearch must be a string");
    const maxSearch = bounded.maxSearchLength;
    if (value.globalSearch.length > maxSearch) throw new ColumnViewStateError("limit-exceeded", "Search exceeds maxSearchLength");
    if (!Array.isArray(value.columnSearches) || !Array.isArray(value.filters) || !Array.isArray(value.sorts)) throw new ColumnViewStateError("invalid-shape", "Searches, filters, and sorts must be arrays");
    if (value.columnSearches.length > bounded.maxColumns) throw new ColumnViewStateError("limit-exceeded", "Column searches exceed maxColumns");
    if (value.filters.length > bounded.maxFilters) throw new ColumnViewStateError("limit-exceeded", "Filters exceed maxFilters");
    if (value.sorts.length > bounded.maxSorts) throw new ColumnViewStateError("limit-exceeded", "Sorts exceed maxSorts");
    const columnSearches: ColumnViewSearch[] = value.columnSearches.map(entry => {
        if (entry === null || typeof entry !== "object" || Array.isArray(entry)) throw new ColumnViewStateError("invalid-shape", "Invalid column search");
        const item = entry as Record<string, unknown>;
        assertKeys(item, ["column", "value"], "column search");
        if (typeof item.column !== "string" || typeof item.value !== "string") throw new ColumnViewStateError("invalid-shape", "Column search requires string column and value");
        if (item.value.length > maxSearch) throw new ColumnViewStateError("limit-exceeded", "Column search exceeds maxSearchLength");
        return { column: requireColumn(item.column), value: item.value };
    });
    const filters: ColumnViewFilter[] = value.filters.map(entry => {
        if (entry === null || typeof entry !== "object" || Array.isArray(entry)) throw new ColumnViewStateError("invalid-shape", "Invalid column filter");
        const item = entry as Record<string, unknown>;
        assertKeys(item, ["column", "operator", "value"], "column filter");
        const column = requireColumn(item.column);
        const operator = requireFilterOperator(item.operator);
        const filter: ColumnViewFilter = { column, operator, ...(item.value === undefined ? {} : { value: requireFilterValue(item.value) }) };
        return filter;
    });
    const sorts: ColumnViewSort[] = value.sorts.map(entry => {
        if (entry === null || typeof entry !== "object" || Array.isArray(entry)) throw new ColumnViewStateError("invalid-shape", "Invalid column sort");
        const item = entry as Record<string, unknown>;
        assertKeys(item, ["column", "direction"], "column sort");
        const column = requireColumn(item.column);
        if (item.direction !== "asc" && item.direction !== "desc") throw new ColumnViewStateError("invalid-shape", "Invalid sort direction");
        return { column, direction: item.direction };
    });
    assertNoDuplicate(columnSearches, "search");
    assertNoDuplicate(filters, "filter");
    assertNoDuplicate(sorts, "sort");
    return freezeState({ globalSearch: value.globalSearch, columnSearches, filters, sorts });
}

/** Serializes a versioned, stable JSON representation. */
export function serializeColumnViewState(state: ColumnViewState): string {
    const normalized = createColumnViewState(state);
    return JSON.stringify({
        version: COLUMN_VIEW_STATE_VERSION,
        globalSearch: normalized.globalSearch,
        // Search/filter order is normalized by the immutable state operations;
        // preserving it makes persistence round-trips lossless. Sort order is
        // intentionally preserved because it carries priority.
        columnSearches: normalized.columnSearches.map(entry => ({ column: entry.column, value: entry.value })),
        filters: normalized.filters.map(entry => ({ column: entry.column, operator: entry.operator, ...(entry.value === undefined ? {} : { value: entry.value }) })),
        sorts: normalized.sorts.map(entry => ({ column: entry.column, direction: entry.direction })),
    });
}

/** Parses and validates persisted state, including version and configurable bounds. */
export function deserializeColumnViewState(json: string, limits: ColumnViewStateLimits = {}): ColumnViewState {
    validateLimits(limits);
    if (typeof json !== "string") throw new ColumnViewStateError("invalid-json", "Persisted state must be a string");
    const bounded = effectiveLimits(limits);
    if (json.length > bounded.maxJsonLength) throw new ColumnViewStateError("limit-exceeded", "Persisted state exceeds maxJsonLength");
    let parsed: unknown;
    try {
        parsed = JSON.parse(json) as unknown;
    } catch {
        throw new ColumnViewStateError("invalid-json", "Persisted state is not valid JSON");
    }
    try {
        return assertPersistedShape(parsed, limits);
    } catch (error) {
        if (error instanceof ColumnViewStateError) throw error;
        throw new ColumnViewStateError("invalid-shape", error instanceof Error ? error.message : String(error));
    }
}
