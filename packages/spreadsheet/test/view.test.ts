import { describe, expect, test } from "vitest";
import { SpreadsheetModel } from "../src/model.js";
import { createSpreadsheetView } from "../src/view.js";

const columns = [
    { id: "item", title: "Item", type: "text" as const },
    { id: "quantity", title: "Quantity", type: "number" as const },
    { id: "price", title: "Price", type: "number" as const },
    { id: "active", title: "Active", type: "boolean" as const },
    { id: "total", title: "Total", type: "number" as const },
];

function makeModel(): SpreadsheetModel {
    return new SpreadsheetModel(columns, 5, [
        ["Alpha", 10, 2, true, "=B1*C1"],
        ["beta", 2, 10, false, "=B2/C2"],
        ["Gamma", null, 5, true, "=1/0"],
        ["item10", 10, 1, false, null],
        ["item2", 2, 1, false, null],
    ]);
}

describe("spreadsheet view", () => {
    test("searches computed display values and formula errors", () => {
        const model = makeModel();
        expect(createSpreadsheetView(model, { search: "20" })).toEqual([0]);
        expect(createSpreadsheetView(model, { search: "#div/0!" })).toEqual([2]);
        expect(createSpreadsheetView(model, { search: "alpha" })).toEqual([0]);
    });

    test("searches one column using computed display text and rejects unknown columns", () => {
        const model = makeModel();
        expect(createSpreadsheetView(model, { columnSearches: [{ column: "total", value: "20" }] })).toEqual([0]);
        expect(createSpreadsheetView(model, { columnSearches: [{ column: "total", value: "#DIV/0!" }] })).toEqual([2]);
        expect(createSpreadsheetView(model, { columnSearches: [{ column: "quantity", value: "10" }] })).toEqual([0, 3]);
        expect(createSpreadsheetView(model, { columnSearches: [{ column: "missing", value: "x" }] })).toEqual([]);
        expect(createSpreadsheetView(model, { columnSearches: [{ column: "total", value: "" }] })).toEqual([0, 1, 2, 3, 4]);
    });

    test("applies every filter operator with typed comparisons", () => {
        const model = makeModel();
        expect(createSpreadsheetView(model, { filters: [{ column: "item", operator: "contains", value: "ITEM" }] })).toEqual([3, 4]);
        expect(createSpreadsheetView(model, { filters: [{ column: "item", operator: "equals", value: "alpha" }] })).toEqual([0]);
        expect(createSpreadsheetView(model, { filters: [{ column: "item", operator: "not-equals", value: "alpha" }] })).toEqual([1, 2, 3, 4]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "gt", value: 2 }] })).toEqual([0, 3]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "gte", value: "10" }] })).toEqual([0, 3]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "lt", value: 10 }] })).toEqual([1, 4]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "lte", value: 2 }] })).toEqual([1, 4]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "empty" }] })).toEqual([2]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "not-empty" }] })).toEqual([0, 1, 3, 4]);
        expect(createSpreadsheetView(model, { filters: [{ column: "active", operator: "equals", value: true }] })).toEqual([0, 2]);
        expect(createSpreadsheetView(model, { filters: [{ column: "total", operator: "equals", value: "#DIV/0!" }] })).toEqual([2]);
    });

    test("does not compare numbers lexicographically and handles null/errors explicitly", () => {
        const model = makeModel();
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "gt", value: "2" }] })).toEqual([0, 3]);
        expect(createSpreadsheetView(model, { filters: [{ column: "total", operator: "empty" }] })).toEqual([3, 4]);
        expect(createSpreadsheetView(model, { filters: [{ column: "total", operator: "not-empty" }] })).toEqual([0, 1, 2]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "equals", value: "" }] })).toEqual([]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "contains", value: "" }] })).toEqual([0, 1, 3, 4]);
        expect(createSpreadsheetView(model, { filters: [{ column: "quantity", operator: "not-equals", value: 999 }] })).toEqual([0, 1, 3, 4]);
        expect(createSpreadsheetView(model, { filters: [{ column: "missing", operator: "equals", value: "x" }] })).toEqual([]);
    });

    test("sorts by multiple columns stably and uses deterministic string ordering", () => {
        const model = makeModel();
        expect(createSpreadsheetView(model, { sorts: [{ column: "quantity", direction: "asc" }] })).toEqual([2, 1, 4, 0, 3]);
        expect(createSpreadsheetView(model, { sorts: [
            { column: "quantity", direction: "asc" },
            { column: "price", direction: "desc" },
        ] })).toEqual([2, 1, 4, 0, 3]);
        expect(createSpreadsheetView(model, { sorts: [{ column: "item", direction: "asc" }] })).toEqual([0, 1, 2, 4, 3]);
        // Unknown sort columns are ignored and preserve source-row order.
        expect(createSpreadsheetView(model, { sorts: [{ column: "missing", direction: "desc" }] })).toEqual([0, 1, 2, 3, 4]);
    });
});
