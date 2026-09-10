import { describe, expect, test } from "vitest";
import { SpreadsheetModel } from "../src/model.js";
import { createSpreadsheetView } from "../src/view.js";
import {
    ColumnViewStateError,
    clearAllColumnViewState,
    clearColumn,
    clearColumnFilter,
    clearColumnSearch,
    clearSort,
    compileColumnViewState,
    createColumnViewState,
    deserializeColumnViewState,
    emptyColumnViewState,
    serializeColumnViewState,
    setColumnFilter,
    setColumnSearch,
    setGlobalSearch,
    setSort,
    toggleSort,
} from "../src/column-view-state.js";

describe("column view state", () => {
    test("keeps immutable state and cycles a sort asc, desc, none", () => {
        const asc = toggleSort(emptyColumnViewState, "quantity");
        const desc = toggleSort(asc, "quantity");
        const none = toggleSort(desc, "quantity");

        expect(emptyColumnViewState.sorts).toEqual([]);
        expect(asc.sorts).toEqual([{ column: "quantity", direction: "asc" }]);
        expect(desc.sorts).toEqual([{ column: "quantity", direction: "desc" }]);
        expect(none.sorts).toEqual([]);
        expect(Object.isFrozen(asc)).toBe(true);
        expect(Object.isFrozen(asc.sorts)).toBe(true);
    });

    test("supports ordered additive multi-sort and coalesces duplicate sort keys", () => {
        const state = createColumnViewState({
            sorts: [
                { column: "A", direction: "asc" },
                { column: "B", direction: "desc" },
                { column: "a", direction: "desc" },
            ],
        });
        expect(state.sorts).toEqual([
            { column: "a", direction: "desc" },
            { column: "B", direction: "desc" },
        ]);

        const added = toggleSort(state, "C", { additive: true });
        expect(added.sorts).toEqual([
            { column: "a", direction: "desc" },
            { column: "B", direction: "desc" },
            { column: "C", direction: "asc" },
        ]);
        expect(toggleSort(added, "a", { additive: true }).sorts).toEqual([
            { column: "B", direction: "desc" },
            { column: "C", direction: "asc" },
        ]);
        expect(toggleSort(toggleSort(added, "a", { additive: true }), "a", { additive: true }).sorts).toEqual([
            { column: "B", direction: "desc" },
            { column: "C", direction: "asc" },
            { column: "a", direction: "asc" },
        ]);
    });

    test("compiles global/column search and preserves typed zero and false filters", () => {
        let state = setGlobalSearch(emptyColumnViewState, "  total ");
        state = setColumnSearch(state, "item", "mouse");
        state = setColumnFilter(state, "quantity", "equals", 0);
        state = setColumnFilter(state, "active", "equals", false);
        const options = compileColumnViewState(state, { columns: [
            { id: "item", title: "Item" },
            { id: "quantity", title: "Quantity" },
            { id: "active", title: "Active" },
        ] });

        expect(options).toEqual({
            search: "  total ",
            columnSearches: [
                { column: "item", value: "mouse" },
            ],
            filters: [
                { column: "quantity", operator: "equals", value: 0 },
                { column: "active", operator: "equals", value: false },
            ],
            sorts: [],
        });
    });

    test("compiles searches against computed formulas and formula errors, plus null filters", () => {
        const model = new SpreadsheetModel([
            { id: "name", title: "Name", type: "text" },
            { id: "computed", title: "Computed", type: "number" },
            { id: "error", title: "Error", type: "text" },
            { id: "nullable", title: "Nullable", type: "text" },
        ], 3, [
            ["first", "=10+10", "=1/0", null],
            ["second", "=5", "=5", "present"],
            ["third", null, null, ""],
        ]);

        let state = setColumnSearch(emptyColumnViewState, "computed", "20");
        state = setColumnSearch(state, "error", "#DIV/0!");
        expect(createSpreadsheetView(model, compileColumnViewState(state))).toEqual([0]);

        state = clearColumnSearch(state, "computed");
        state = clearColumnSearch(state, "error");
        state = setColumnFilter(state, "nullable", "empty");
        expect(createSpreadsheetView(model, compileColumnViewState(state))).toEqual([0, 2]);
    });

    test("preserves numeric zero and boolean false filters when applied to a model", () => {
        const model = new SpreadsheetModel([
            { id: "quantity", title: "Quantity", type: "number" },
            { id: "active", title: "Active", type: "boolean" },
        ], 3, [[0, false], [0, true], [1, false]]);
        let state = setColumnFilter(emptyColumnViewState, "quantity", "equals", 0);
        state = setColumnFilter(state, "active", "equals", false);
        expect(createSpreadsheetView(model, compileColumnViewState(state))).toEqual([0]);
    });

    test("clears one column without disturbing other columns, then clears all", () => {
        let state = setColumnSearch(emptyColumnViewState, "a", "one");
        state = setColumnFilter(state, "a", "contains", "x");
        state = setSort(state, "a", "asc");
        state = setColumnSearch(state, "b", "two");
        state = setColumnFilter(state, "b", "equals", 0);
        state = setSort(state, "b", "desc", { additive: true });

        const cleared = clearColumn(state, "A");
        expect(cleared.columnSearches).toEqual([{ column: "b", value: "two" }]);
        expect(cleared.filters).toEqual([{ column: "b", operator: "equals", value: 0 }]);
        expect(cleared.sorts).toEqual([{ column: "b", direction: "desc" }]);
        expect(clearColumnSearch(cleared, "b").columnSearches).toEqual([]);
        expect(clearColumnFilter(cleared, "b").filters).toEqual([]);
        expect(clearSort(cleared, "b").sorts).toEqual([]);
        expect(clearAllColumnViewState()).toBe(emptyColumnViewState);
    });

    test("resolves known ids/titles and safely handles unknown columns", () => {
        const state = createColumnViewState({
            columnSearches: [{ column: "Unknown", value: "x" }],
            filters: [{ column: "Quantity", operator: "equals", value: 0 }],
            sorts: [{ column: "Unknown", direction: "asc" }, { column: "Item", direction: "desc" }],
        });
        expect(compileColumnViewState(state, { columns: [
            { id: "item", title: "Item" },
            { id: "quantity", title: "Quantity" },
        ] })).toEqual({
            search: "",
            columnSearches: [],
            filters: [
                { column: "quantity", operator: "equals", value: 0 },
            ],
            sorts: [{ column: "item", direction: "desc" }],
        });

        expect(compileColumnViewState(state, { resolveColumn: () => { throw new Error("resolver failure"); } }).sorts).toEqual([]);
    });

    test("preserves sort priority when compiled into the stable view", () => {
        const model = new SpreadsheetModel([
            { id: "group", title: "Group", type: "text" },
            { id: "score", title: "Score", type: "number" },
        ], 3, [["A", 1], ["A", 2], ["B", 9]]);
        let state = setSort(emptyColumnViewState, "group", "asc");
        state = setSort(state, "score", "desc", { additive: true });
        expect(createSpreadsheetView(model, compileColumnViewState(state))).toEqual([1, 0, 2]);

        const scoreOnly = clearColumn(state, "group");
        expect(scoreOnly.sorts).toEqual([{ column: "score", direction: "desc" }]);
        expect(createSpreadsheetView(model, compileColumnViewState(scoreOnly))).toEqual([2, 1, 0]);

        const toggled = toggleSort(state, "group", { additive: true });
        expect(toggled.sorts).toEqual([{ column: "group", direction: "desc" }, { column: "score", direction: "desc" }]);
        expect(toggleSort(toggled, "group", { additive: true }).sorts).toEqual([{ column: "score", direction: "desc" }]);
    });

    test("serializes stable versioned JSON and validates malformed state and bounds", () => {
        let state = setColumnSearch(emptyColumnViewState, "B", "x");
        state = setColumnSearch(state, "A", "y");
        state = setSort(state, "B", "asc");
        const json = serializeColumnViewState(state);
        expect(json).toBe('{"version":1,"globalSearch":"","columnSearches":[{"column":"B","value":"x"},{"column":"A","value":"y"}],"filters":[],"sorts":[{"column":"B","direction":"asc"}]}');
        const restored = deserializeColumnViewState(json);
        expect(restored).toEqual(state);
        expect(() => deserializeColumnViewState("not json")).toThrow(ColumnViewStateError);
        expect(() => deserializeColumnViewState('{"version":2,"globalSearch":"","columnSearches":[],"filters":[],"sorts":[]}')).toThrow(/version/);
        expect(() => deserializeColumnViewState('{"version":1,"globalSearch":"","columnSearches":[],"filters":[],"sorts":[{"column":"A","direction":"asc"},{"column":"a","direction":"desc"}]}')).toThrow(/Duplicate/);
        expect(() => deserializeColumnViewState(json, { maxJsonLength: 4 })).toThrow(/maxJsonLength/);
        expect(() => deserializeColumnViewState(json, { maxSorts: 0 })).toThrow(/maxSorts/);
        expect(() => deserializeColumnViewState('{"version":1,"globalSearch":"","columnSearches":[],"filters":[{"column":"A","operator":"equals","value":null}],"sorts":[]}')).toThrow(ColumnViewStateError);
        const defaultBoundSorts = Array.from({ length: 101 }, (_, index) => ({ column: `column-${index}`, direction: "asc" as const }));
        expect(() => deserializeColumnViewState(JSON.stringify({ version: 1, globalSearch: "", columnSearches: [], filters: [], sorts: defaultBoundSorts }))).toThrow(/maxSorts/);
    });
});
