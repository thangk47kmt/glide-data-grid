import { performance } from "node:perf_hooks";
import { LOADING_CELL, PagedDataSource, SpreadsheetModel, SpreadsheetViewIndex, createSpreadsheetView } from "../dist/index.js";

function positiveInteger(name, fallback) {
    const raw = process.env[name];
    if (raw === undefined) return fallback;
    const value = Number(raw);
    if (!Number.isInteger(value) || value <= 0) throw new Error(`${name} must be a positive integer`);
    return value;
}

function measure(name, action) {
    const started = performance.now();
    const result = action();
    return { name, milliseconds: Number((performance.now() - started).toFixed(2)), result };
}

async function measureAsync(name, action) {
    const started = performance.now();
    const result = await action();
    return { name, milliseconds: Number((performance.now() - started).toFixed(2)), result };
}

const rowCount = positiveInteger("SPREADSHEET_BENCH_ROWS", 100_000);
const columnCount = positiveInteger("SPREADSHEET_BENCH_COLUMNS", 10);
const chainLength = positiveInteger("SPREADSHEET_BENCH_CHAIN", 500);
const columns = Array.from({ length: columnCount }, (_, index) => ({
    id: `column-${index + 1}`,
    title: `Column ${index + 1}`,
    type: "number",
}));

const memoryBefore = process.memoryUsage().heapUsed;
const rowMaterialization = measure("materialize-cells", () =>
    Array.from({ length: rowCount }, (_, row) => Array.from({ length: columnCount }, (_, col) => row + col))
);
const modelConstruction = measure("construct-model", () =>
    new SpreadsheetModel(columns, rowCount, rowMaterialization.result)
);
const model = modelConstruction.result;
const filterAndSort = measure("filter-sort-100k", () => createSpreadsheetView(model, {
    filters: [{ column: columns[columnCount - 1].id, operator: "gte", value: rowCount - 1 }],
    sorts: [{ column: columns[0].id, direction: "desc" }],
}));
const searchTerm = String(rowCount - 1);
const searchTerms = [searchTerm, String(Math.floor(rowCount / 2)), "42", "999", "not-present"];
const search = measure("search-100k", () => createSpreadsheetView(model, { search: searchTerm }));
const repeatedSearch = measure("search-100k-five-queries", () => searchTerms.map(term => createSpreadsheetView(model, { search: term })));
const indexMemoryBefore = process.memoryUsage().heapUsed;
const viewIndexConstruction = measure("construct-view-index", () => new SpreadsheetViewIndex(model));
const indexMemoryAfter = process.memoryUsage().heapUsed;
const indexedSearch = measure("indexed-search-100k", () => viewIndexConstruction.result.query({ search: searchTerm }));
const repeatedIndexedSearch = measure("indexed-search-100k-five-queries", () => searchTerms.map(term => viewIndexConstruction.result.query({ search: term })));

const chainRows = Array.from({ length: chainLength }, (_, row) => [row === 0 ? 1 : `=A${row}+1`]);
const chainModel = new SpreadsheetModel([{ id: "value", title: "Value", type: "number" }], chainLength, chainRows);
const dependencyChain = measure("dependency-chain", () => chainModel.getValue(0, chainLength - 1));
const memoryAfter = process.memoryUsage().heapUsed;

// Virtual 500k x 20 benchmark: only requested pages are generated. Keep this
// separate from the existing model benchmark, whose materialized rows are
// intentionally retained as a comparison point.
const virtualRowCount = 500_000;
const virtualColumnCount = 20;
const virtualPageSize = positiveInteger("SPREADSHEET_VIRTUAL_PAGE_SIZE", 512);
const virtualMaxCachedPages = positiveInteger("SPREADSHEET_VIRTUAL_MAX_CACHED_PAGES", 8);
const virtualPageRequests = positiveInteger("SPREADSHEET_VIRTUAL_PAGE_REQUESTS", 32);
const virtual = new PagedDataSource({
    rowCount: virtualRowCount,
    columnCount: virtualColumnCount,
    pageSize: virtualPageSize,
    maxCachedPages: virtualMaxCachedPages,
    seed: 0x5eed,
});
const virtualLoadingBefore = virtual.getCell(virtualRowCount - 1, virtualColumnCount - 1) === LOADING_CELL;
const virtualCold = await measureAsync("virtual-cold-page", () => virtual.getPage(0));
const virtualCached = await measureAsync("virtual-cached-page-access-1000", async () => {
    const pages = await Promise.all(Array.from({ length: 1000 }, () => virtual.getPage(0)));
    return pages.length;
});

function deterministicPageRequests(count, pageCount, seed) {
    const pages = [];
    let value = seed >>> 0;
    for (let index = 0; index < count; index += 1) {
        value = (Math.imul(value, 1664525) + 1013904223) >>> 0;
        pages.push(value % pageCount);
    }
    return pages;
}

virtual.invalidate();
const randomPages = deterministicPageRequests(virtualPageRequests, virtual.pageCount, 0x1234abcd);
const virtualRandom = await measureAsync("virtual-random-page-traversal", async () => {
    for (const pageIndex of randomPages) await virtual.getPage(pageIndex);
    return randomPages.length;
});
const virtualStats = virtual.getStats();
const virtualResidentCellsUpperBound = virtualStats.cachedPages * virtualPageSize * virtualColumnCount;
if (virtualStats.cachedPages > virtualMaxCachedPages) throw new Error("Virtual page cache exceeded maxCachedPages");
if (virtualResidentCellsUpperBound > virtualMaxCachedPages * virtualPageSize * virtualColumnCount) throw new Error("Virtual resident-cell bound exceeded");
if (virtualStats.generatedCells >= virtualRowCount * virtualColumnCount) throw new Error("Virtual benchmark generated the complete dataset");

const report = {
    configuration: { rowCount, columnCount, populatedCells: rowCount * columnCount, chainLength },
    timingsMs: {
        materializeCells: rowMaterialization.milliseconds,
        constructModel: modelConstruction.milliseconds,
        filterAndSort: filterAndSort.milliseconds,
        search: search.milliseconds,
        repeatedSearch: repeatedSearch.milliseconds,
        constructViewIndex: viewIndexConstruction.milliseconds,
        indexedSearch: indexedSearch.milliseconds,
        repeatedIndexedSearch: repeatedIndexedSearch.milliseconds,
        dependencyChain: dependencyChain.milliseconds,
        virtualColdPage: virtualCold.milliseconds,
        virtualCachedPageAccess1000: virtualCached.milliseconds,
        virtualRandomPageTraversal: virtualRandom.milliseconds,
    },
    resultChecks: {
        filteredRows: filterAndSort.result.length,
        searchMatches: search.result.length,
        indexedSearchMatches: indexedSearch.result.length,
        repeatedSearchMatches: repeatedSearch.result.map(rows => rows.length),
        repeatedIndexedSearchMatches: repeatedIndexedSearch.result.map(rows => rows.length),
        chainValue: dependencyChain.result,
        virtualLoadingBeforeFirstPage: virtualLoadingBefore,
        virtualColdPageRows: virtualCold.result.rowCount,
        virtualCachedAccesses: virtualCached.result,
        virtualRandomPageRequests: virtualRandom.result,
    },
    searchIndex: {
        ...viewIndexConstruction.result.getSearchCacheStats(),
        approximateHeapDeltaMb: Number(((indexMemoryAfter - indexMemoryBefore) / 1024 / 1024).toFixed(2)),
    },
    approximateHeapDeltaMb: Number(((memoryAfter - memoryBefore) / 1024 / 1024).toFixed(2)),
    virtualDataSource: {
        configuration: {
            rowCount: virtualRowCount,
            columnCount: virtualColumnCount,
            pageSize: virtualPageSize,
            maxCachedPages: virtualMaxCachedPages,
            pageRequests: virtualPageRequests,
            populatedCells: virtualRowCount * virtualColumnCount,
        },
        stats: virtualStats,
        residentCellsUpperBound: virtualResidentCellsUpperBound,
        residentCellsBound: virtualMaxCachedPages * virtualPageSize * virtualColumnCount,
    },
};

console.log(JSON.stringify(report, null, 2));
