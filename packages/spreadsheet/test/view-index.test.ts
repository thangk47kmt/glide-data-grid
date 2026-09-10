import { describe, expect, test, vi } from "vitest";
import { SpreadsheetModel } from "../src/model.js";
import { createSpreadsheetView } from "../src/view.js";
import { SpreadsheetViewIndex, createSpreadsheetViewIndex } from "../src/view-index.js";

const columns = [
    { id: "item", title: "Item", type: "text" as const },
    { id: "quantity", title: "Quantity", type: "number" as const },
    { id: "active", title: "Active", type: "boolean" as const },
];

const model = () => new SpreadsheetModel(columns, 5, [
    ["Alpha 20", 20, true],
    ["Beta 5", 5, false],
    ["Gamma", null, true],
    ["alpha 10", 10, false],
    ["Delta", 2, true],
]);

describe("SpreadsheetViewIndex", () => {
    test("matches baseline for search, filters and stable multi-sort", () => {
        const spreadsheet = model();
        const index = new SpreadsheetViewIndex(spreadsheet);
        const queries = [
            { search: "alpha" },
            { filters: [{ column: "quantity", operator: "gte" as const, value: "10" }] },
            { filters: [{ column: "active", operator: "equals" as const, value: true }], sorts: [{ column: "quantity", direction: "desc" as const }] },
            { search: "a", filters: [{ column: "missing", operator: "equals" as const, value: "x" }] },
            { sorts: [{ column: "missing", direction: "asc" as const }, { column: "item", direction: "asc" as const }] },
        ];
        queries.forEach(options => expect(index.query(options)).toEqual(createSpreadsheetView(spreadsheet, options)));
    });

    test("matches baseline for per-column computed display searches and null filters", () => {
        const spreadsheet = new SpreadsheetModel([
            { id: "computed", title: "Computed", type: "number" },
            { id: "error", title: "Error", type: "text" },
            { id: "nullable", title: "Nullable", type: "text" },
        ], 3, [
            ["=10+10", "=1/0", null],
            ["=0", "=5", "present"],
            [null, null, ""],
        ]);
        const index = createSpreadsheetViewIndex(spreadsheet);
        const queries = [
            { columnSearches: [{ column: "computed", value: "20" }] },
            { columnSearches: [{ column: "computed", value: "0" }] },
            { columnSearches: [{ column: "error", value: "#DIV/0!" }] },
            { filters: [{ column: "nullable", operator: "empty" as const }] },
            { columnSearches: [{ column: "missing", value: "x" }] },
        ];
        queries.forEach(options => expect(index.query(options)).toEqual(createSpreadsheetView(spreadsheet, options)));
    });

    test("caches display search text across repeated queries", () => {
        const spreadsheet = model();
        const getDisplayValue = vi.spyOn(spreadsheet, "getDisplayValue");
        const index = createSpreadsheetViewIndex(spreadsheet);
        const callsAfterBuild = getDisplayValue.mock.calls.length;
        expect(callsAfterBuild).toBe(columns.length * spreadsheet.rowCount);
        index.query({ search: "a" });
        index.query({ search: "beta" });
        expect(getDisplayValue).toHaveBeenCalledTimes(callsAfterBuild);
    });

    test("does not match search text across column boundaries", () => {
        const spreadsheet = new SpreadsheetModel([
            { id: "left", title: "Left" },
            { id: "right", title: "Right" },
        ], 1, [["ab", "cd"]]);
        const index = createSpreadsheetViewIndex(spreadsheet);
        expect(index.query({ search: "bc" })).toEqual(createSpreadsheetView(spreadsheet, { search: "bc" }));
    });

    test("handles control characters without allowing cross-cell matches", () => {
        const spreadsheet = new SpreadsheetModel([
            { id: "left", title: "Left" },
            { id: "right", title: "Right" },
        ], 2, [["a\u0000b", "cd"], ["line\n", "tab\tvalue"]]);
        const index = createSpreadsheetViewIndex(spreadsheet);
        expect(index.query({ search: "\u0000b" })).toEqual([0]);
        expect(index.query({ search: "bcd" })).toEqual(createSpreadsheetView(spreadsheet, { search: "bcd" }));
        expect(index.query({ search: "\ntab" })).toEqual(createSpreadsheetView(spreadsheet, { search: "\ntab" }));
    });

    test("stores one normalized search entry per wide row", () => {
        const wideColumns = Array.from({ length: 10 }, (_, index) => ({ id: `c${index}`, title: `C${index}` }));
        const spreadsheet = new SpreadsheetModel(wideColumns, 1, [wideColumns.map((_, index) => `value-${index}`)]);
        const getDisplayValue = vi.spyOn(spreadsheet, "getDisplayValue");
        const index = createSpreadsheetViewIndex(spreadsheet);
        expect(getDisplayValue).toHaveBeenCalledTimes(10);
        expect(index.getSearchCacheStats()).toEqual({ rowCount: 1, normalizedCharacters: 70, cachedCellCount: 10, boundaryBytes: 40 });
        index.query({ search: "value-9" });
        index.query({ search: "value-0" });
        expect(getDisplayValue).toHaveBeenCalledTimes(10);
    });

    test("requires explicit invalidation after model edits", () => {
        const spreadsheet = model();
        const index = createSpreadsheetViewIndex(spreadsheet);
        spreadsheet.setCell(0, 1, "Alpha moved");
        expect(index.query({ search: "moved" })).toEqual([]);
        index.invalidateRow(1);
        expect(index.query({ search: "moved" })).toEqual([1]);
        spreadsheet.setCell(0, 2, "Beta moved");
        spreadsheet.setCell(0, 3, "Beta also");
        index.invalidateRows([2, 3, 3]);
        expect(index.query({ search: "beta" })).toEqual([2, 3]);
        spreadsheet.setCell(0, 4, "Delta changed");
        index.refreshAll();
        expect(index.query({ search: "changed" })).toEqual([4]);
    });

    test("keeps formula-error and null/filter semantics and validates rows", () => {
        const spreadsheet = new SpreadsheetModel([{ id: "value", title: "Value" }], 2, [["=1/0"], [null]]);
        const index = createSpreadsheetViewIndex(spreadsheet);
        expect(index.query({ search: "#div/0!" })).toEqual(createSpreadsheetView(spreadsheet, { search: "#div/0!" }));
        expect(index.query({ filters: [{ column: "value", operator: "empty" }] })).toEqual([1]);
        expect(() => index.invalidateRow(-1)).toThrow(RangeError);
        expect(() => index.invalidateRow(2)).toThrow(RangeError);
        expect(() => index.invalidateRows([0, 2])).toThrow(RangeError);
    });
});
