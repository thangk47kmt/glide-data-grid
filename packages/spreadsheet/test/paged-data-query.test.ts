import { describe, expect, test } from "vitest";
import { PagedDataQuery } from "../src/paged-data-query.js";
import { PagedDataSource } from "../src/paged-data-source.js";

describe("PagedDataQuery", () => {
    test("scans a generated 500k-row source without loading pages", async () => {
        const source = new PagedDataSource({
            rowCount: 500_000,
            columnCount: 20,
            pageSize: 500,
            maxCachedPages: 2,
            generateCell: (row, column) => column === 0 ? `Record ${row}` : column === 1 ? row : row % 2 === 0,
        });
        const query = new PagedDataQuery(source, [{ id: "name", title: "Name" }, { id: "amount", title: "Amount" }]);
        const result = await query.query({ filters: [{ column: "amount", operator: "gte", value: 499_990 }], sorts: [{ column: "amount", direction: "desc" }], offset: 2, limit: 3 });
        expect(result.rows).toEqual([499_997, 499_996, 499_995]);
        expect(result.totalCount).toBe(10);
        expect(source.getStats()).toMatchObject({ cachedPages: 0, generatedCells: 0 });
    });

    test("supports global and column searches, typed filters, and stable multi-sort", async () => {
        const source = new PagedDataSource({
            rowCount: 6,
            columnCount: 4,
            pageSize: 2,
            maxCachedPages: 1,
            generateCell: (row, column) => [["Alpha", 2, true, "x"], ["beta", 1, false, "needle"], ["Alpha", 1, true, "x"], ["gamma", 2, false, "x"], ["beta", 1, true, "needle"], ["delta", 3, false, "x"]][row]![column] as string | number | boolean,
        });
        const query = new PagedDataQuery(source, [{ id: "label" }, { id: "score" }, { id: "enabled" }, { id: "notes" }]);
        const result = await query.query({ search: "NEEDLE", filters: [{ column: "enabled", operator: "equals", value: false }], sorts: [{ column: "score", direction: "asc" }, { column: "label", direction: "asc" }], limit: 10 });
        expect(result.rows).toEqual([1]);
        expect(result.totalCount).toBe(1);
        const exact = await query.query({ columnSearches: [{ column: "label", value: "alpha" }], filters: [{ column: "score", operator: "equals", value: 1 }], limit: 10 });
        expect(exact.rows).toEqual([2]);
        const columnOnly = await query.query({ columnSearches: [{ column: "label", value: "beta" }], limit: 10 });
        expect(columnOnly.rows).toEqual([1, 4]);
        expect(columnOnly.totalCount).toBe(2);
    });

    test("resolves display captions while keeping persisted ids as query output metadata", async () => {
        const source = new PagedDataSource({ rowCount: 2, columnCount: 2, generateCell: (row, column) => column === 0 ? row + 1 : (row + 1) * 10 });
        const query = new PagedDataQuery(source, [{ id: "so_luong", title: "Số lượng" }, { id: "don_gia", title: "Đơn giá" }]);
        expect((await query.query({ filters: [{ column: "Đơn giá", operator: "gte", value: 20 }], limit: 10 })).rows).toEqual([1]);
    });

    test("does not silently bind duplicate captions", async () => {
        const source = new PagedDataSource({ rowCount: 1, columnCount: 2, generateCell: (_row, column) => column + 1 });
        const query = new PagedDataQuery(source, [{ id: "left", title: "Giá" }, { id: "right", title: "Giá" }]);
        expect((await query.query({ filters: [{ column: "Giá", operator: "equals", value: 1 }], limit: 10 })).rows).toEqual([]);
    });

    test("does not bypass an ambiguous numeric caption using the index fallback", async () => {
        const source = new PagedDataSource({ rowCount: 1, columnCount: 2, generateCell: (_row, column) => column + 1 });
        const query = new PagedDataQuery(source, [{ id: "left", title: "0" }, { id: "right", title: "0" }]);
        expect((await query.query({ filters: [{ column: "0", operator: "equals", value: 1 }], limit: 10 })).rows).toEqual([]);
    });

    test("uses an optional flattened row reader for global search", async () => {
        let cellReads = 0;
        const query = new PagedDataQuery({
            rowCount: 3,
            columnCount: 2,
            readCell: () => { cellReads += 1; return "not used"; },
            getCell: () => "not used",
            readSearchText: row => ["alpha", "target", "omega"][row]!,
        });
        expect((await query.query({ search: "TARGET", limit: 10 })).rows).toEqual([1]);
        expect(cellReads).toBe(0);
        const found = await query.find({ query: "target" });
        expect(found.matches).toHaveLength(0);
        expect(cellReads).toBe(2);
    });

    test("returns page windows and a bounded find result", async () => {
        const source = new PagedDataSource({ rowCount: 20, columnCount: 2, pageSize: 4, maxCachedPages: 1, generateCell: (row, column) => column === 0 ? `Row ${row}` : row % 2 });
        const query = new PagedDataQuery(source, [{ id: "name" }, { id: "parity" }]);
        expect(await query.queryPage(2, 4)).toMatchObject({ rows: [8, 9, 10, 11], totalCount: 20, offset: 8, limit: 4, hasMore: true });
        const found = await query.find({ query: "row", columns: ["name"], maxResults: 3 });
        expect(found.matches.map(match => match.row)).toEqual([0, 1, 2]);
        expect(found.totalMatches).toBe(20);
        expect(found.truncated).toBe(true);
    });

    test("supports case-sensitive, whole-cell and regular-expression find", async () => {
        const source = new PagedDataSource({
            rowCount: 4,
            columnCount: 1,
            pageSize: 2,
            generateCell: row => ["Alpha 10", "alpha 20", "Alpha", "Beta 10"][row]!,
        });
        const query = new PagedDataQuery(source, [{ id: "name" }]);
        expect((await query.find({ query: "Alpha", caseSensitive: true })).matches.map(match => match.row)).toEqual([0, 2]);
        expect((await query.find({ query: "Alpha", wholeCell: true })).matches.map(match => match.row)).toEqual([2]);
        expect((await query.find({ query: "^alpha \\d+$", regexp: true })).matches.map(match => match.row)).toEqual([0, 1]);
        await expect(query.find({ query: "[", regexp: true })).rejects.toThrow(SyntaxError);
    });

    test("falls back to one-page-at-a-time sources", async () => {
        const source = new PagedDataSource({ rowCount: 10, columnCount: 1, pageSize: 2, maxCachedPages: 1, generateCell: row => row });
        const query = new PagedDataQuery({ rowCount: source.rowCount, columnCount: source.columnCount, pageSize: source.pageSize, getCell: source.getCell.bind(source), getPage: source.getPage.bind(source) }, [{ id: "value" }]);
        expect((await query.query({ filters: [{ column: "value", operator: "gt", value: 7 }], limit: 10 })).rows).toEqual([8, 9]);
        expect(source.getStats().cachedPages).toBe(1);
    });

    test("uses the returned page when a fallback source disables caching", async () => {
        const source = new PagedDataSource({ rowCount: 6, columnCount: 1, pageSize: 2, maxCachedPages: 0, generateCell: row => row });
        const query = new PagedDataQuery({ rowCount: source.rowCount, columnCount: source.columnCount, pageSize: source.pageSize, getCell: source.getCell.bind(source), getPage: source.getPage.bind(source) }, [{ id: "value" }]);
        expect((await query.query({ filters: [{ column: "value", operator: "gte", value: 4 }], limit: 10 })).rows).toEqual([4, 5]);
        expect((await query.query({ filters: [{ column: "value", operator: "lt", value: 2 }], limit: 10 })).rows).toEqual([0, 1]);
    });
});
