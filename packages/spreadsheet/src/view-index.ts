import { matchesFilter, compareForSort, normalizeSearch, normalizeSearchValue } from "./view-helpers.js";
import type { SpreadsheetModel } from "./model.js";
import type { SpreadsheetViewOptions } from "./view.js";

function containsWithinCell(text: string, boundaries: Uint32Array, rowOffset: number, columnCount: number, search: string, targetColumn?: number): boolean {
    for (let match = text.indexOf(search); match !== -1; match = text.indexOf(search, match + 1)) {
        const matchEnd = match + search.length;
        // Find the first cumulative end at or after the match. Empty cells
        // produce repeated boundaries, which is intentional and still gives
        // exact cell containment without storing per-cell start arrays.
        let low = targetColumn ?? 0;
        let high = targetColumn ?? columnCount - 1;
        while (low < high) {
            const middle = Math.floor((low + high) / 2);
            if (boundaries[rowOffset + middle] < matchEnd) low = middle + 1;
            else high = middle;
        }
        const cellEnd = boundaries[rowOffset + low];
        const previousEnd = low === 0 ? 0 : boundaries[rowOffset + low - 1];
        if (previousEnd <= match && matchEnd <= cellEnd && (targetColumn === undefined || low === targetColumn)) return true;
    }
    return false;
}

function validateRow(row: number, rowCount: number): void {
    if (!Number.isInteger(row) || row < 0 || row >= rowCount) throw new RangeError(`Row must be an integer in the range 0..${Math.max(0, rowCount - 1)}`);
}

/**
 * Search-accelerated view over a model. The model is not observed: call
 * invalidateRow(s) after edits or refreshAll after a batch. Filters and sorts
 * always read current model values, while cached search text follows that
 * explicit invalidation contract.
 */
export class SpreadsheetViewIndex {
    private readonly searchTextByRow: string[];
    private readonly columnCount: number;
    private boundaries: Uint32Array;

    public constructor(public readonly model: SpreadsheetModel) {
        this.columnCount = model.columns.length;
        this.searchTextByRow = [];
        this.boundaries = new Uint32Array(0);
        this.refreshAll();
    }

    public refreshAll(): void {
        this.assertDimensions();
        this.searchTextByRow.length = this.model.rowCount;
        this.boundaries = new Uint32Array(this.model.rowCount * this.columnCount);
        for (let row = 0; row < this.model.rowCount; row++) this.searchTextByRow[row] = this.buildSearchText(row);
    }

    public invalidateRow(row: number): void {
        this.assertDimensions();
        validateRow(row, this.model.rowCount);
        this.searchTextByRow[row] = this.buildSearchText(row);
    }

    public invalidateRows(rows: readonly number[]): void {
        this.assertDimensions();
        const uniqueRows = [...new Set(rows)];
        uniqueRows.forEach(row => validateRow(row, this.model.rowCount));
        uniqueRows.forEach(row => { this.searchTextByRow[row] = this.buildSearchText(row); });
    }

    /** Returns compact cache accounting useful for diagnostics/benchmarks. */
    public getSearchCacheStats(): { readonly rowCount: number; readonly normalizedCharacters: number; readonly cachedCellCount: number; readonly boundaryBytes: number } {
        return {
            rowCount: this.searchTextByRow.length,
            normalizedCharacters: this.searchTextByRow.reduce((total, row) => total + row.length, 0),
            cachedCellCount: this.model.rowCount * this.columnCount,
            boundaryBytes: this.boundaries.byteLength,
        };
    }

    public query(options: SpreadsheetViewOptions = {}): number[] {
        const search = normalizeSearch(options.search ?? "");
        const columnSearches = (options.columnSearches ?? []).map(entry => ({
            search: normalizeSearchValue(entry.value),
            col: this.model.resolveColumn(entry.column),
        }));
        const filters = (options.filters ?? []).map(filter => ({ filter, col: this.model.resolveColumn(filter.column) }));
        const sorts = (options.sorts ?? []).map(sort => ({ sort, col: this.model.resolveColumn(sort.column) }));
        const result = Array.from({ length: this.model.rowCount }, (_, row) => row).filter(row => {
            if (search !== "" && !containsWithinCell(this.searchTextByRow[row], this.boundaries, row * this.columnCount, this.columnCount, search)) return false;
            if (columnSearches.some(({ search: columnSearch, col }) => col === undefined || columnSearch !== "" && !containsWithinCell(this.searchTextByRow[row], this.boundaries, row * this.columnCount, this.columnCount, columnSearch, col))) return false;
            return filters.every(({ filter, col }) => col !== undefined && matchesFilter(this.model.getValue(col, row), filter));
        });
        result.sort((leftRow, rightRow) => {
            for (const { sort, col } of sorts) {
                if (col === undefined) continue;
                const direction = sort.direction === "desc" ? -1 : 1;
                const comparison = compareForSort(this.model.getValue(col, leftRow), this.model.getValue(col, rightRow));
                if (comparison !== 0) return comparison * direction;
            }
            return leftRow - rightRow;
        });
        return result;
    }

    private buildSearchText(row: number): string {
        let text = "";
        for (let col = 0; col < this.model.columns.length; col++) {
            const normalized = normalizeSearchValue(this.model.getDisplayValue(col, row));
            text += normalized;
            const boundary = text.length;
            if (!Number.isSafeInteger(boundary) || boundary > 0xffffffff) {
                throw new RangeError("Normalized search text exceeds Uint32 boundary capacity");
            }
            this.boundaries[row * this.columnCount + col] = boundary;
        }
        return text;
    }

    private assertDimensions(): void {
        if (this.model.columns.length !== this.columnCount) throw new RangeError("SpreadsheetViewIndex does not support changing column count");
        if (this.model.rowCount * this.columnCount > 0xffffffff) throw new RangeError("SpreadsheetViewIndex exceeds Uint32 boundary capacity");
    }
}

/** Factory alias for applications that prefer functional construction. */
export function createSpreadsheetViewIndex(model: SpreadsheetModel): SpreadsheetViewIndex {
    return new SpreadsheetViewIndex(model);
}
