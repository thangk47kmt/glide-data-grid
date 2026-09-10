import { compareForSort, matchesFilter, normalizeSearch, normalizeSearchValue } from "./view-helpers.js";
import type { FilterOperator, SpreadsheetColumnSearch, SpreadsheetFilter, SpreadsheetSort } from "./view.js";
import type { LoadingCell, PagedCellValue, PagedDataPage, PagedDataSource } from "./paged-data-source.js";

/** A column description used to resolve ids and human-readable titles. */
export interface PagedQueryColumn {
    readonly id: string;
    readonly title?: string;
}

/** The subset of a paged source needed by the query engine. */
export interface PagedQuerySource {
    readonly rowCount: number;
    readonly columnCount: number;
    readonly pageSize?: number;
    getCell(row: number, column: number): PagedCellValue | LoadingCell;
    getPage?(pageIndex: number): Promise<PagedDataPage>;
    /** Optional direct reader. It is useful for deterministic generated rows: no page is materialized. */
    readCell?(row: number, column: number): PagedCellValue;
    /** Optional pre-flattened row text used to accelerate global search. */
    readSearchText?(row: number): string;
}

export interface PagedQueryOptions {
    readonly search?: string;
    readonly columnSearches?: readonly SpreadsheetColumnSearch[];
    readonly filters?: readonly SpreadsheetFilter[];
    readonly sorts?: readonly SpreadsheetSort[];
    /** Zero-based result offset. `page` is a convenient alternative. */
    readonly offset?: number;
    /** Number of row ids to return. Defaults to 100. */
    readonly limit?: number;
    readonly page?: number;
    readonly pageSize?: number;
}

export interface PagedQueryResult {
    readonly rows: readonly number[];
    /** Number of rows matching search and filters, before the result window. */
    readonly totalCount: number;
    readonly matchCount: number;
    readonly totalRows: number;
    readonly offset: number;
    readonly limit: number;
    readonly hasMore: boolean;
}

export interface PagedFindOptions {
    readonly query: string;
    readonly columns?: readonly string[];
    readonly caseSensitive?: boolean;
    readonly wholeCell?: boolean;
    /** Interpret `query` as a JavaScript regular expression. */
    readonly regexp?: boolean;
    /** Defaults to 1,000 matches. A cap keeps accidental broad searches bounded. */
    readonly maxResults?: number;
}

export interface PagedFindMatch {
    readonly row: number;
    readonly column: number;
    readonly value: PagedCellValue;
}

export interface PagedFindResult {
    readonly matches: readonly PagedFindMatch[];
    readonly totalMatches: number;
    readonly truncated: boolean;
}

type SortRecord = { readonly row: number; readonly keys: readonly PagedCellValue[] };

function columnIdentity(value: string): string {
    return value.trim().normalize("NFKC").toLocaleLowerCase("en-US");
}

function validateNonNegative(value: number, name: string): number {
    if (!Number.isInteger(value) || value < 0) throw new RangeError(`${name} must be a non-negative integer`);
    return value;
}

function validatePositive(value: number, name: string): number {
    if (!Number.isInteger(value) || value <= 0) throw new RangeError(`${name} must be a positive integer`);
    return value;
}

function isLoading(value: PagedCellValue | LoadingCell): value is LoadingCell {
    return typeof value === "object" && value !== null && "kind" in value && value.kind === "loading";
}

function compareRecords(left: SortRecord, right: SortRecord, sorts: readonly SpreadsheetSort[]): number {
    for (let index = 0; index < sorts.length; index += 1) {
        const comparison = compareForSort(left.keys[index]!, right.keys[index]!);
        if (comparison !== 0) return sorts[index]!.direction === "desc" ? -comparison : comparison;
    }
    return left.row - right.row;
}

/**
 * Queries a virtual data source while retaining only row ids and sort keys.
 * It never builds a row-by-column result matrix. Sources with `readCell` can
 * scan deterministic generated data without loading any pages at all; other
 * sources are loaded one page at a time and rely on their bounded page cache.
 */
export class PagedDataQuery {
    private readonly lookup = new Map<string, number>();
    /** Duplicate ids/captions are rejected at resolution time instead of choosing one by order. */
    private readonly ambiguousColumns = new Set<string>();
    private readonly directReader: ((row: number, column: number) => PagedCellValue) | undefined;
    private readonly searchReader: ((row: number) => string) | undefined;
    private currentPage: PagedDataPage | undefined;

    public constructor(public readonly source: PagedQuerySource, columns: readonly PagedQueryColumn[] = []) {
        if (!Number.isInteger(source.rowCount) || source.rowCount < 0) throw new RangeError("rowCount must be a non-negative integer");
        if (!Number.isInteger(source.columnCount) || source.columnCount < 0) throw new RangeError("columnCount must be a non-negative integer");
        columns.forEach((column, index) => {
            if (typeof column.id !== "string" || columnIdentity(column.id) === "") throw new TypeError("Column id must be a non-empty string");
            if (index >= source.columnCount) throw new RangeError("Column definitions exceed source columnCount");
            this.registerColumnName(column.id, index);
            if (column.title !== undefined) this.registerColumnName(column.title, index);
        });
        this.directReader = source.readCell?.bind(source);
        this.searchReader = source.readSearchText?.bind(source);
    }

    /** Counts matching rows without retaining the matching row ids. */
    public async count(options: Omit<PagedQueryOptions, "offset" | "limit" | "page" | "pageSize"> = {}): Promise<number> {
        return (await this.query({ ...options, offset: 0, limit: 1 })).totalCount;
    }

    /** Returns one bounded result window and the full matching-row count. */
    public async query(options: PagedQueryOptions = {}): Promise<PagedQueryResult> {
        const window = this.resolveWindow(options);
        const search = normalizeSearch(options.search ?? "");
        const columnSearches = (options.columnSearches ?? []).map(entry => ({
            col: this.resolveColumn(entry.column),
            search: normalizeSearchValue(entry.value),
        }));
        const filters = (options.filters ?? []).map(filter => ({ col: this.resolveColumn(filter.column), filter }));
        const sorts = (options.sorts ?? []).map(sort => ({ col: this.resolveColumn(sort.column), sort })).filter((entry): entry is { col: number; sort: SpreadsheetSort } => entry.col !== undefined);
        const activeSorts = sorts.map(entry => entry.sort);

        // This is both a useful fast path and an important property for the
        // default 500k-row story: paging an unfiltered source does no scan.
        if (search === "" && columnSearches.every(entry => entry.search === "") && filters.length === 0 && activeSorts.length === 0) {
            const rows = Array.from({ length: Math.min(window.limit, Math.max(0, this.source.rowCount - window.offset)) }, (_, index) => window.offset + index);
            return this.result(rows, this.source.rowCount, window);
        }

        const records: SortRecord[] = [];
        let totalCount = 0;
        for (let pageIndex = 0; pageIndex < this.pageCount; pageIndex += 1) {
            if (pageIndex > 0 && pageIndex % 16 === 0) await new Promise<void>(resolve => setTimeout(resolve, 0));
            await this.ensurePage(pageIndex);
            const rowStart = pageIndex * this.pageSize;
            const rowEnd = Math.min(this.source.rowCount, rowStart + this.pageSize);
            for (let row = rowStart; row < rowEnd; row += 1) {
                if (!this.matches(row, search, columnSearches, filters)) continue;
                totalCount += 1;
                if (activeSorts.length > 0) {
                    records.push({ row, keys: sorts.map(entry => this.read(row, entry.col)) });
                } else if (totalCount > window.offset && totalCount <= window.offset + window.limit) {
                    records.push({ row, keys: [] });
                }
            }
        }
        if (activeSorts.length > 0) records.sort((left, right) => compareRecords(left, right, activeSorts));
        const rows = activeSorts.length > 0 ? records.slice(window.offset, window.offset + window.limit).map(record => record.row) : records.map(record => record.row);
        return this.result(rows, totalCount, window);
    }

    /** Alias that makes page-window use explicit at call sites. */
    public async queryPage(page: number, pageSize: number, options: Omit<PagedQueryOptions, "page" | "pageSize" | "offset" | "limit"> = {}): Promise<PagedQueryResult> {
        return this.query({ ...options, page, pageSize });
    }

    /** Bounded row-major find. It reports the complete match count even when results are capped. */
    public async find(options: PagedFindOptions): Promise<PagedFindResult> {
        if (typeof options.query !== "string" || options.query.length === 0) throw new TypeError("query must not be empty");
        const maxResults = options.maxResults === undefined ? 1_000 : validateNonNegative(options.maxResults, "maxResults");
        const wanted = options.caseSensitive === true ? options.query : options.query.toLocaleLowerCase("en-US");
        const expression = options.regexp === true
            ? new RegExp(options.wholeCell === true ? `^(?:${options.query})$` : options.query, options.caseSensitive === true ? "" : "i")
            : undefined;
        const columns = options.columns === undefined ? Array.from({ length: this.source.columnCount }, (_, index) => index) : options.columns.map(column => this.resolveColumn(column));
        const canPrefilterRows = options.columns === undefined && expression === undefined && this.searchReader !== undefined;
        if (columns.some(column => column === undefined)) return { matches: [], totalMatches: 0, truncated: false };
        const matches: PagedFindMatch[] = [];
        let totalMatches = 0;
        for (let pageIndex = 0; pageIndex < this.pageCount; pageIndex += 1) {
            if (pageIndex > 0 && pageIndex % 16 === 0) await new Promise<void>(resolve => setTimeout(resolve, 0));
            await this.ensurePage(pageIndex);
            const rowStart = pageIndex * this.pageSize;
            const rowEnd = Math.min(this.source.rowCount, rowStart + this.pageSize);
            for (let row = rowStart; row < rowEnd; row += 1) {
                if (canPrefilterRows && !normalizeSearchValue(this.searchReader!(row)).includes(normalizeSearchValue(options.query))) continue;
                for (const column of columns as number[]) {
                    const value = this.read(row, column);
                    const text = String(value === null ? "" : value);
                    const comparable = options.caseSensitive === true ? text : text.toLocaleLowerCase("en-US");
                    const found = expression !== undefined ? expression.test(text) : options.wholeCell === true ? comparable === wanted : comparable.includes(wanted);
                    if (!found) continue;
                    totalMatches += 1;
                    if (matches.length < maxResults) matches.push({ row, column, value });
                }
            }
        }
        return { matches, totalMatches, truncated: totalMatches > matches.length };
    }

    private get pageSize(): number {
        return this.source.pageSize ?? 100;
    }

    private get pageCount(): number {
        return Math.ceil(this.source.rowCount / this.pageSize);
    }

    private resolveWindow(options: PagedQueryOptions): { readonly offset: number; readonly limit: number } {
        const limit = options.pageSize ?? options.limit ?? 100;
        validatePositive(limit, "limit");
        const offset = options.page !== undefined ? validateNonNegative(options.page, "page") * limit : validateNonNegative(options.offset ?? 0, "offset");
        return { offset, limit };
    }

    private resolveColumn(column: string): number | undefined {
        if (typeof column !== "string") return undefined;
        const normalized = columnIdentity(column);
        // An ambiguous metadata name must not fall through to the legacy
        // numeric-index syntax (for example two captions both named "0").
        if (this.ambiguousColumns.has(normalized)) return undefined;
        const named = this.lookup.get(normalized);
        if (named !== undefined) return named;
        if (/^\d+$/.test(column.trim())) {
            const numeric = Number(column.trim());
            return Number.isSafeInteger(numeric) && numeric < this.source.columnCount ? numeric : undefined;
        }
        return undefined;
    }

    private registerColumnName(name: string, index: number): void {
        const normalized = columnIdentity(name);
        if (normalized === "" || this.ambiguousColumns.has(normalized)) return;
        const existing = this.lookup.get(normalized);
        if (existing === undefined) this.lookup.set(normalized, index);
        else if (existing !== index) {
            this.lookup.delete(normalized);
            this.ambiguousColumns.add(normalized);
        }
    }

    private async ensurePage(pageIndex: number): Promise<void> {
        if (this.directReader !== undefined) return;
        if (this.source.getPage === undefined) throw new Error("Paged query source must implement getPage or readCell");
        this.currentPage = await this.source.getPage(pageIndex);
    }

    private read(row: number, column: number): PagedCellValue {
        if (this.directReader !== undefined) return this.directReader(row, column);
        const value = this.source.getCell(row, column);
        if (isLoading(value)) {
            const page = this.currentPage;
            if (page !== undefined && row >= page.rowStart && row < page.rowStart + page.rowCount) return page.values[row - page.rowStart]?.[column] ?? null;
            throw new Error("Paged query attempted to read an unloaded page");
        }
        return value;
    }

    private matches(row: number, search: string, columnSearches: readonly { readonly col: number | undefined; readonly search: string }[], filters: readonly { readonly col: number | undefined; readonly filter: SpreadsheetFilter }[]): boolean {
        if (search !== "") {
            if (this.searchReader !== undefined) {
                if (!normalizeSearchValue(this.searchReader(row)).includes(search)) return false;
            } else {
                let found = false;
                for (let column = 0; column < this.source.columnCount; column += 1) {
                    if (normalizeSearchValue(String(this.read(row, column) ?? "")).includes(search)) {
                        found = true;
                        break;
                    }
                }
                if (!found) return false;
            }
        }
        if (columnSearches.some(entry => entry.col === undefined || entry.search !== "" && !normalizeSearchValue(String(this.read(row, entry.col) ?? "")).includes(entry.search))) return false;
        return filters.every(entry => entry.col !== undefined && matchesFilter(this.read(row, entry.col), entry.filter));
    }

    private result(rows: readonly number[], totalCount: number, window: { readonly offset: number; readonly limit: number }): PagedQueryResult {
        return { rows, totalCount, matchCount: totalCount, totalRows: this.source.rowCount, offset: window.offset, limit: window.limit, hasMore: window.offset + rows.length < totalCount };
    }
}

export function createPagedDataQuery(source: PagedQuerySource, columns: readonly PagedQueryColumn[] = []): PagedDataQuery {
    return new PagedDataQuery(source, columns);
}

/** A convenient source type for callers that already use PagedDataSource. */
export type VirtualQuerySource = PagedDataSource;

export type PagedQueryFilterOperator = FilterOperator;
