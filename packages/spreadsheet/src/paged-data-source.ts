/** A cell value produced by the virtual data source. */
export type PagedCellValue = string | number | boolean | null;

export interface LoadingCell {
    readonly kind: "loading";
}

/** Stable sentinel returned by {@link PagedDataSource.getCell} before a page is loaded. */
export const LOADING_CELL: LoadingCell = Object.freeze({ kind: "loading" });

export type PagedCellGenerator = (row: number, column: number, seed: number) => PagedCellValue;
export type PagedRowGenerator = (row: number, columnCount: number, seed: number) => readonly PagedCellValue[];

export interface PagedDataSourceOptions {
    /** Defaults to 500,000 rows. No rows are generated until requested. */
    readonly rowCount?: number;
    /** Defaults to 20 columns. */
    readonly columnCount?: number;
    /** Number of rows in one page; defaults to 100. */
    readonly pageSize?: number;
    /** Maximum resident pages; zero disables caching. Defaults to 8. */
    readonly maxCachedPages?: number;
    /** A uint32 seed used by the default generator. */
    readonly seed?: number;
    /** Optional deterministic cell generator. Ignored when generateRow is supplied. */
    readonly generateCell?: PagedCellGenerator;
    /** Optional deterministic row generator. It must return exactly columnCount values. */
    readonly generateRow?: PagedRowGenerator;
}

export interface PagedDataPage {
    readonly pageIndex: number;
    readonly rowStart: number;
    readonly rowCount: number;
    readonly columnCount: number;
    readonly values: readonly (readonly PagedCellValue[])[];
}

export interface PagedDataSourceStats {
    readonly rowCount: number;
    readonly columnCount: number;
    readonly pageSize: number;
    readonly pageCount: number;
    readonly maxCachedPages: number;
    readonly cachedPages: number;
    readonly inflightPages: number;
    readonly overlayCells: number;
    readonly generatedCells: number;
    readonly cacheHits: number;
    readonly cacheMisses: number;
    readonly evictions: number;
}

function validateNonNegativeInteger(value: number, name: string): number {
    if (!Number.isInteger(value) || value < 0) throw new RangeError(`${name} must be a non-negative integer`);
    return value;
}

function validatePositiveInteger(value: number, name: string): number {
    if (!Number.isInteger(value) || value <= 0) throw new RangeError(`${name} must be a positive integer`);
    return value;
}

function cellKey(row: number, column: number): string {
    return `${row}:${column}`;
}

function pageOf(row: number, pageSize: number): number {
    return Math.floor(row / pageSize);
}

function hashCell(row: number, column: number, seed: number): number {
    let value = (seed ^ Math.imul(row + 1, 0x9e3779b1) ^ Math.imul(column + 1, 0x85ebca6b)) >>> 0;
    value ^= value >>> 16;
    value = Math.imul(value, 0x7feb352d) >>> 0;
    value ^= value >>> 15;
    value = Math.imul(value, 0x846ca68b) >>> 0;
    return (value ^ (value >>> 16)) >>> 0;
}

/** Deterministic default value used when no generator is supplied. */
export function defaultPagedCell(row: number, column: number, seed: number): number {
    return hashCell(row, column, seed) % 1_000_000;
}

function freezePage(page: PagedDataPage): PagedDataPage {
    const values = Object.freeze(page.values.map(row => Object.freeze([...row])));
    return Object.freeze({ ...page, values });
}

/**
 * A model-independent virtual grid. Rows are generated only while their page
 * is requested; the source never materializes the complete row-by-column grid.
 * Cache entries and generated page data are immutable. Cell overlays are kept
 * separately and survive page invalidation.
 */
export class PagedDataSource {
    public readonly rowCount: number;
    public readonly columnCount: number;
    public readonly pageSize: number;
    public readonly pageCount: number;
    public readonly maxCachedPages: number;
    public readonly seed: number;

    private readonly generateCell: PagedCellGenerator;
    private readonly generateRow: PagedRowGenerator | undefined;
    private readonly cache = new Map<number, PagedDataPage>();
    private readonly inflight = new Map<number, Promise<PagedDataPage>>();
    private readonly overlays = new Map<string, PagedCellValue>();
    private readonly overlayColumnsByRow = new Map<number, Set<number>>();
    private readonly pageEpoch = new Map<number, number>();
    private directGeneratedRow: { readonly row: number; readonly values: readonly PagedCellValue[] } | undefined;
    private globalEpoch = 0;
    private generatedCells = 0;
    private cacheHits = 0;
    private cacheMisses = 0;
    private evictions = 0;

    public constructor(options: PagedDataSourceOptions = {}) {
        this.rowCount = validateNonNegativeInteger(options.rowCount ?? 500_000, "rowCount");
        this.columnCount = validateNonNegativeInteger(options.columnCount ?? 20, "columnCount");
        this.pageSize = validatePositiveInteger(options.pageSize ?? 100, "pageSize");
        this.maxCachedPages = validateNonNegativeInteger(options.maxCachedPages ?? 8, "maxCachedPages");
        const seed = options.seed ?? 0;
        if (!Number.isInteger(seed) || !Number.isFinite(seed)) throw new RangeError("seed must be a finite integer");
        this.seed = seed >>> 0;
        this.generateCell = options.generateCell ?? defaultPagedCell;
        this.generateRow = options.generateRow;
        if (typeof this.generateCell !== "function") throw new TypeError("generateCell must be a function");
        if (this.generateRow !== undefined && typeof this.generateRow !== "function") throw new TypeError("generateRow must be a function");
        this.pageCount = Math.ceil(this.rowCount / this.pageSize);
    }

    /** Returns a cell immediately, or the stable loading sentinel if its page is not resident. */
    public getCell(row: number, column: number): PagedCellValue | LoadingCell {
        this.validateCellCoordinate(row, column);
        const overlay = this.overlays.get(cellKey(row, column));
        if (overlay !== undefined || this.overlays.has(cellKey(row, column))) return overlay as PagedCellValue;
        const page = this.cache.get(pageOf(row, this.pageSize));
        if (page === undefined) return LOADING_CELL;
        this.promote(page.pageIndex, page);
        return page.values[row - page.rowStart]?.[column] ?? null;
    }

    /**
     * Reads a generated cell directly without loading or retaining its page.
     * Sparse overlays take precedence, which keeps this reader safe for edits.
     * This is primarily intended for bounded-memory scans such as filtering
     * and sorting a deterministic virtual dataset.
     */
    public readCell(row: number, column: number): PagedCellValue {
        this.validateCellCoordinate(row, column);
        const key = cellKey(row, column);
        if (this.overlays.has(key)) return this.overlays.get(key) as PagedCellValue;
        if (this.generateRow !== undefined) {
            if (this.directGeneratedRow?.row !== row) {
                const generated = this.generateRow(row, this.columnCount, this.seed);
                if (!Array.isArray(generated) || generated.length !== this.columnCount) throw new RangeError("generateRow must return exactly columnCount values");
                generated.forEach(value => this.validateValue(value));
                this.directGeneratedRow = { row, values: [...generated] };
            }
            return this.directGeneratedRow.values[column] ?? null;
        }
        const value = this.generateCell(row, column, this.seed);
        this.validateValue(value);
        return value;
    }

    /**
     * Loads one page asynchronously. Concurrent calls for the same page share
     * one promise. A page invalidated while loading is not inserted into cache.
     */
    public getPage(pageIndex: number): Promise<PagedDataPage> {
        this.validatePageIndex(pageIndex);
        const cached = this.cache.get(pageIndex);
        if (cached !== undefined) {
            this.cacheHits += 1;
            this.promote(pageIndex, cached);
            return Promise.resolve(cached);
        }
        const existing = this.inflight.get(pageIndex);
        if (existing !== undefined) return existing;
        this.cacheMisses += 1;
        const globalEpoch = this.globalEpoch;
        const pageEpoch = this.pageEpoch.get(pageIndex) ?? 0;
        const promise = Promise.resolve().then(() => this.buildPage(pageIndex)).then(page => {
            if (globalEpoch === this.globalEpoch && pageEpoch === (this.pageEpoch.get(pageIndex) ?? 0)) this.store(page);
            return page;
        }).finally(() => {
            if (this.inflight.get(pageIndex) === promise) this.inflight.delete(pageIndex);
        });
        this.inflight.set(pageIndex, promise);
        return promise;
    }

    /** Updates a sparse overlay without generating or materializing other cells. */
    public updateCell(row: number, column: number, value: PagedCellValue): void {
        this.validateCellCoordinate(row, column);
        this.validateValue(value);
        const key = cellKey(row, column);
        this.overlays.set(key, value);
        const columns = this.overlayColumnsByRow.get(row) ?? new Set<number>();
        columns.add(column);
        this.overlayColumnsByRow.set(row, columns);
        const index = pageOf(row, this.pageSize);
        const cached = this.cache.get(index);
        if (cached !== undefined) {
            const values = cached.values.map(pageRow => [...pageRow]);
            values[row - cached.rowStart]![column] = value;
            this.store(freezePage({ ...cached, values }));
        }
        if (this.inflight.has(index)) this.invalidatePage(index);
    }

    /** Removes cached page data while preserving cell overlays. */
    public invalidate(pageIndex?: number): void {
        if (pageIndex === undefined) {
            this.globalEpoch += 1;
            this.cache.clear();
            this.inflight.clear();
            return;
        }
        this.invalidatePage(pageIndex);
    }

    public clearOverlay(row: number, column: number): void {
        this.validateCellCoordinate(row, column);
        this.overlays.delete(cellKey(row, column));
        const columns = this.overlayColumnsByRow.get(row);
        columns?.delete(column);
        if (columns?.size === 0) this.overlayColumnsByRow.delete(row);
        this.invalidate(pageOf(row, this.pageSize));
    }

    public clearOverlays(): void {
        this.overlays.clear();
        this.overlayColumnsByRow.clear();
        this.invalidate();
    }

    /** Sparse overlay columns for one row, used by formula/search adapters. */
    public getOverlayColumns(row: number): readonly number[] {
        if (!Number.isInteger(row) || row < 0 || row >= this.rowCount) throw new RangeError("row is outside the data source");
        return [...(this.overlayColumnsByRow.get(row) ?? [])];
    }

    public getStats(): PagedDataSourceStats {
        return {
            rowCount: this.rowCount,
            columnCount: this.columnCount,
            pageSize: this.pageSize,
            pageCount: this.pageCount,
            maxCachedPages: this.maxCachedPages,
            cachedPages: this.cache.size,
            inflightPages: this.inflight.size,
            overlayCells: this.overlays.size,
            generatedCells: this.generatedCells,
            cacheHits: this.cacheHits,
            cacheMisses: this.cacheMisses,
            evictions: this.evictions,
        };
    }

    private buildPage(pageIndex: number): PagedDataPage {
        const rowStart = pageIndex * this.pageSize;
        const rowCount = Math.min(this.pageSize, this.rowCount - rowStart);
        const values: PagedCellValue[][] = [];
        for (let offset = 0; offset < rowCount; offset += 1) {
            const row = rowStart + offset;
            const generated = this.generateRow?.(row, this.columnCount, this.seed);
            if (generated !== undefined) {
                if (!Array.isArray(generated) || generated.length !== this.columnCount) throw new RangeError("generateRow must return exactly columnCount values");
                generated.forEach(value => this.validateValue(value));
                values.push([...generated]);
                this.generatedCells += this.columnCount;
            } else {
                const generatedRow: PagedCellValue[] = [];
                for (let column = 0; column < this.columnCount; column += 1) {
                    const value = this.generateCell(row, column, this.seed);
                    this.validateValue(value);
                    generatedRow.push(value);
                }
                values.push(generatedRow);
                this.generatedCells += this.columnCount;
            }
        }
        for (let offset = 0; offset < values.length; offset += 1) {
            for (let column = 0; column < this.columnCount; column += 1) {
                const overlay = this.overlays.get(cellKey(rowStart + offset, column));
                if (overlay !== undefined || this.overlays.has(cellKey(rowStart + offset, column))) values[offset]![column] = overlay as PagedCellValue;
            }
        }
        return freezePage({ pageIndex, rowStart, rowCount, columnCount: this.columnCount, values });
    }

    private store(page: PagedDataPage): void {
        if (this.maxCachedPages === 0) return;
        this.cache.delete(page.pageIndex);
        this.cache.set(page.pageIndex, page);
        while (this.cache.size > this.maxCachedPages) {
            const oldest = this.cache.keys().next().value as number;
            this.cache.delete(oldest);
            this.evictions += 1;
        }
    }

    private promote(pageIndex: number, page: PagedDataPage): void {
        this.cache.delete(pageIndex);
        this.cache.set(pageIndex, page);
    }

    private invalidatePage(pageIndex: number): void {
        this.validatePageIndex(pageIndex);
        this.pageEpoch.set(pageIndex, (this.pageEpoch.get(pageIndex) ?? 0) + 1);
        this.cache.delete(pageIndex);
        this.inflight.delete(pageIndex);
    }

    private validatePageIndex(pageIndex: number): void {
        if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= this.pageCount) throw new RangeError("pageIndex is outside the data source");
    }

    private validateCellCoordinate(row: number, column: number): void {
        if (!Number.isInteger(row) || row < 0 || row >= this.rowCount) throw new RangeError("row is outside the data source");
        if (!Number.isInteger(column) || column < 0 || column >= this.columnCount) throw new RangeError("column is outside the data source");
    }

    private validateValue(value: PagedCellValue): void {
        if (value === null || typeof value === "string" || typeof value === "boolean") return;
        if (typeof value === "number" && Number.isFinite(value)) return;
        throw new TypeError("Generated cells must be null, strings, booleans, or finite numbers");
    }
}

/** Descriptive alias for consumers that prefer the spreadsheet terminology. */
export const VirtualSpreadsheetDataSource = PagedDataSource;
