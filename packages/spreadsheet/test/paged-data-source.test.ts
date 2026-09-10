import { describe, expect, test } from "vitest";
import { LOADING_CELL, PagedDataSource, defaultPagedCell } from "../src/paged-data-source.js";

describe("PagedDataSource", () => {
    test("defaults to a large virtual grid without materializing cells", () => {
        const source = new PagedDataSource({ pageSize: 128, maxCachedPages: 2, seed: 7 });
        expect(source.rowCount).toBe(500_000);
        expect(source.columnCount).toBe(20);
        expect(source.pageCount).toBe(Math.ceil(500_000 / 128));
        expect(source.getStats().cachedPages).toBe(0);
        expect(source.getCell(499_999, 19)).toBe(LOADING_CELL);
        expect(source.getStats().generatedCells).toBe(0);
    });

    test("generates deterministic pages and handles the final page boundary", async () => {
        const options = { rowCount: 5, columnCount: 2, pageSize: 2, maxCachedPages: 2, seed: 123 };
        const first = new PagedDataSource(options);
        const second = new PagedDataSource(options);
        const firstPage = await first.getPage(0);
        const samePage = await second.getPage(0);
        expect(firstPage).toEqual(samePage);
        expect(first.getCell(0, 0)).toBe(defaultPagedCell(0, 0, 123));
        expect(first.getCell(1, 1)).toBe(defaultPagedCell(1, 1, 123));
        expect(first.getCell(2, 0)).toBe(LOADING_CELL);
        const last = await first.getPage(2);
        expect(last).toMatchObject({ pageIndex: 2, rowStart: 4, rowCount: 1, columnCount: 2 });
        expect(last.values).toHaveLength(1);
        expect(() => first.getPage(3)).toThrow(RangeError);
    });

    test("deduplicates concurrent page requests", async () => {
        let generated = 0;
        const source = new PagedDataSource({ rowCount: 10, columnCount: 2, pageSize: 5, maxCachedPages: 2, generateCell: (row, column, seed) => {
            generated += 1;
            return row * 10 + column + seed;
        } });
        const one = source.getPage(1);
        const two = source.getPage(1);
        expect(one).toBe(two);
        await Promise.all([one, two]);
        expect(generated).toBe(10);
        expect(source.getStats()).toMatchObject({ cacheMisses: 1, inflightPages: 0, cachedPages: 1 });
    });

    test("uses bounded LRU cache and reports hits/evictions", async () => {
        const source = new PagedDataSource({ rowCount: 12, columnCount: 1, pageSize: 2, maxCachedPages: 2 });
        await source.getPage(0);
        await source.getPage(1);
        await source.getPage(0); // promote page 0
        await source.getPage(2); // evicts page 1, not page 0
        expect(source.getStats()).toMatchObject({ cachedPages: 2, cacheHits: 1, cacheMisses: 3, evictions: 1 });
        expect(source.getCell(0, 0)).not.toBe(LOADING_CELL);
        expect(source.getCell(2, 0)).toBe(LOADING_CELL);
        expect(source.getCell(4, 0)).not.toBe(LOADING_CELL);
    });

    test("updates sparse overlays without generating unrelated rows and survives invalidation", async () => {
        const source = new PagedDataSource({ rowCount: 100, columnCount: 20, pageSize: 10, maxCachedPages: 2, seed: 1 });
        source.updateCell(42, 3, 0);
        source.updateCell(42, 4, false);
        expect(source.getOverlayColumns(42)).toEqual([3, 4]);
        expect(source.getStats()).toMatchObject({ overlayCells: 2, generatedCells: 0 });
        expect(source.getCell(42, 3)).toBe(0);
        expect(source.getCell(42, 4)).toBe(false);
        await source.getPage(4);
        expect(source.getCell(42, 3)).toBe(0);
        expect(source.getCell(42, 4)).toBe(false);
        source.invalidate(4);
        expect(source.getCell(42, 3)).toBe(0);
        expect(source.getCell(42, 5)).toBe(LOADING_CELL);
        await source.getPage(4);
        source.clearOverlay(42, 3);
        expect(source.getOverlayColumns(42)).toEqual([4]);
        expect(source.getCell(42, 3)).toBe(LOADING_CELL);
        await source.getPage(4);
        expect(source.getCell(42, 3)).toBe(defaultPagedCell(42, 3, 1));
        source.clearOverlays();
        expect(source.getStats().overlayCells).toBe(0);
        expect(source.getOverlayColumns(42)).toEqual([]);
    });

    test("reuses a directly generated row while scanning its columns", () => {
        let calls = 0;
        const source = new PagedDataSource({ rowCount: 2, columnCount: 3, generateRow: row => { calls += 1; return [row, row + 1, row + 2]; } });
        expect([source.readCell(0, 0), source.readCell(0, 1), source.readCell(0, 2)]).toEqual([0, 1, 2]);
        expect(calls).toBe(1);
        expect(source.readCell(1, 2)).toBe(3);
        expect(calls).toBe(2);
    });

    test("invalidates an in-flight page without allowing stale data into cache", async () => {
        const source = new PagedDataSource({ rowCount: 2, columnCount: 1, pageSize: 1, maxCachedPages: 1 });
        const request = source.getPage(0);
        source.invalidate(0);
        await request;
        expect(source.getCell(0, 0)).toBe(LOADING_CELL);
        expect(source.getStats().cachedPages).toBe(0);
    });

    test("validates dimensions, coordinates, generators, and values", async () => {
        expect(() => new PagedDataSource({ pageSize: 0 })).toThrow(RangeError);
        expect(() => new PagedDataSource({ rowCount: -1 })).toThrow(RangeError);
        const source = new PagedDataSource({ rowCount: 2, columnCount: 2, pageSize: 1 });
        expect(() => source.getCell(2, 0)).toThrow(RangeError);
        expect(() => source.getCell(0, 2)).toThrow(RangeError);
        expect(() => source.getPage(-1)).toThrow(RangeError);
        const invalidRow = new PagedDataSource({ rowCount: 1, columnCount: 2, generateRow: () => [1] });
        await expect(invalidRow.getPage(0)).rejects.toThrow(RangeError);
        const invalidValue = new PagedDataSource({ rowCount: 1, columnCount: 1, generateCell: () => Number.NaN });
        await expect(invalidValue.getPage(0)).rejects.toThrow(TypeError);
    });
});
