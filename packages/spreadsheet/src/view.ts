import type { SpreadsheetModel } from "./model.js";
import { compareForSort, matchesFilter, normalizeSearch, normalizeSearchValue } from "./view-helpers.js";

export type FilterOperator = "contains" | "equals" | "not-equals" | "gt" | "gte" | "lt" | "lte" | "empty" | "not-empty";
export interface SpreadsheetFilter {
    readonly column: string;
    readonly operator: FilterOperator;
    readonly value?: string | number | boolean;
}
export interface SpreadsheetColumnSearch {
    readonly column: string;
    readonly value: string;
}
export interface SpreadsheetSort {
    readonly column: string;
    readonly direction: "asc" | "desc";
}
export interface SpreadsheetViewOptions {
    readonly search?: string;
    readonly columnSearches?: readonly SpreadsheetColumnSearch[];
    readonly filters?: readonly SpreadsheetFilter[];
    readonly sorts?: readonly SpreadsheetSort[];
}

export function createSpreadsheetView(model: SpreadsheetModel, options: SpreadsheetViewOptions = {}): number[] {
    const search = normalizeSearch(options.search ?? "");
    const columnSearches = (options.columnSearches ?? []).map(entry => ({
        search: normalizeSearchValue(entry.value),
        col: model.resolveColumn(entry.column),
    }));
    // Unknown filter columns intentionally match no rows. Unknown sort
    // columns are ignored, preserving source-row order as the stable fallback.
    const filters = (options.filters ?? []).map(filter => ({ filter, col: model.resolveColumn(filter.column) }));
    const sorts = (options.sorts ?? []).map(sort => ({ sort, col: model.resolveColumn(sort.column) }));
    const result = Array.from({ length: model.rowCount }, (_, row) => row).filter(row => {
        if (search !== "" && !model.columns.some((_, col) => normalizeSearchValue(model.getDisplayValue(col, row)).includes(search))) return false;
        if (columnSearches.some(({ search: columnSearch, col }) => col === undefined || columnSearch !== "" && !normalizeSearchValue(model.getDisplayValue(col, row)).includes(columnSearch))) return false;
        return filters.every(({ filter, col }) => col !== undefined && matchesFilter(model.getValue(col, row), filter));
    });
    result.sort((leftRow, rightRow) => {
        for (const { sort, col } of sorts) {
            if (col === undefined) continue;
            const direction = sort.direction === "desc" ? -1 : 1;
            const comparison = compareForSort(model.getValue(col, leftRow), model.getValue(col, rightRow));
            if (comparison !== 0) return comparison * direction;
        }
        return leftRow - rightRow;
    });
    return result;
}
