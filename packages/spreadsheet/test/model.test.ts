import { describe, expect, test } from "vitest";
import { SpreadsheetModel } from "../src/model.js";
import { createSpreadsheetView } from "../src/view.js";
import { FormulaFunctionRegistry } from "../src/function-registry.js";

const columns = [
    { id: "item", title: "Item", type: "text" as const },
    { id: "quantity", title: "Quantity", type: "number" as const },
    { id: "price", title: "Unit Price", type: "number" as const },
    { id: "total", title: "Total", type: "number" as const },
];

describe("SpreadsheetModel", () => {
    test("recalculates dependent formulas after an edit", () => {
        const model = new SpreadsheetModel(columns, 2, [
            ["Keyboard", 2, 10, "=[@Quantity]*[@Unit Price]"],
            ["Mouse", 3, 5, "=[@Quantity]*[@Unit Price]"],
        ]);
        expect(model.getValue(3, 0)).toBe(20);
        model.setCell(1, 0, 4);
        expect(model.getValue(3, 0)).toBe(40);
        expect(model.getValue(3, 1)).toBe(15);
    });

    test("invalidates a transitive dependency chain without rebuilding the model", () => {
        const chainColumns = [
            { id: "source", title: "Source", type: "number" as const },
            { id: "double", title: "Double", type: "number" as const },
            { id: "result", title: "Result", type: "number" as const },
        ];
        const model = new SpreadsheetModel(chainColumns, 2, [
            [2, "=A1*2", "=B1*3"],
            [100, "=A2+1", "=B2+1"],
        ]);
        const sameModel = model;

        // Prime both branches of the calculation cache before changing one
        // source cell. Only its dependent chain needs a new value.
        expect(model.getValue(2, 0)).toBe(12);
        expect(model.getValue(2, 1)).toBe(102);
        model.setCell(0, 0, 5);

        expect(model).toBe(sameModel);
        expect(model.getValue(1, 0)).toBe(10);
        expect(model.getValue(2, 0)).toBe(30);
        expect(model.getValue(2, 1)).toBe(102);
    });

    test("resolves persisted ids and user-facing captions, including Vietnamese captions", () => {
        const localizedColumns = [
            { id: "so_luong", title: "Số lượng", type: "number" as const },
            { id: "don_gia", title: "Đơn giá", type: "number" as const },
            { id: "thanh_tien", title: "Thành tiền", type: "number" as const },
        ];
        const model = new SpreadsheetModel(localizedColumns, 1, [[2, 25, "=[@Số lượng]*[@Đơn giá]"]]);
        expect(model.resolveColumn("so_luong")).toBe(0);
        expect(model.resolveColumn("ĐƠN GIÁ")).toBe(1);
        expect(model.getValue(2, 0)).toBe(50);
        model.setCell(1, 0, 30);
        expect(model.getValue(2, 0)).toBe(60);
    });

    test("does not silently choose a column when a caption is ambiguous", () => {
        const ambiguous = new SpreadsheetModel([
            { id: "left", title: "Giá", type: "number" },
            { id: "right", title: "Giá", type: "number" },
            { id: "total", title: "Tổng", type: "number" },
        ], 1, [[2, 3, "=[@Giá]"]]);
        expect(ambiguous.resolveColumn("Giá")).toBeUndefined();
        expect(ambiguous.resolveColumn("left")).toBe(0);
        expect(ambiguous.getDisplayValue(2, 0)).toBe("#REF!");
    });

    test("reports formula syntax and circular reference errors", () => {
        const model = new SpreadsheetModel(columns, 1);
        model.setCell(0, 0, "=1+");
        expect(model.getDisplayValue(0, 0)).toBe("#VALUE!");
        model.setCell(0, 0, "=B1");
        model.setCell(1, 0, "=A1");
        expect(model.getDisplayValue(0, 0)).toBe("#CYCLE!");
        model.setCell(0, 0, "=#REF!");
        expect(model.getDisplayValue(0, 0)).toBe("#REF!");
    });

    test("searches, filters and performs stable multi-column sorting", () => {
        const model = new SpreadsheetModel(columns, 3, [
            ["Keyboard", 1, 50, "=[@Quantity]*[@Unit Price]"],
            ["Mouse", 3, 10, "=[@Quantity]*[@Unit Price]"],
            ["Keyboard Pro", 2, 40, "=[@Quantity]*[@Unit Price]"],
        ]);
        expect(createSpreadsheetView(model, { search: "keyboard" })).toEqual([0, 2]);
        expect(createSpreadsheetView(model, { filters: [{ column: "total", operator: "gte", value: 40 }] })).toEqual([0, 2]);
        expect(createSpreadsheetView(model, { sorts: [{ column: "total", direction: "desc" }] })).toEqual([2, 0, 1]);
    });

    test("evaluates custom scalar, range, and structured-reference functions", () => {
        const registry = new FormulaFunctionRegistry();
        registry.register("double", args => {
            const value = args[0];
            return typeof value === "number" ? value * 2 : 0;
        });
        registry.register("range_total", args => {
            const value = args[0];
            if (value === null || typeof value !== "object" || !("values" in value)) return 0;
            return value.values.reduce<number>((sum, item) => sum + (typeof item === "number" ? item : 0), 0);
        });
        const model = new SpreadsheetModel(
            columns,
            2,
            [
                ["Keyboard", 2, 10, "=DOUBLE([@Quantity])"],
                ["Mouse", 3, 5, "=RANGE_TOTAL(B1:B2)"],
            ],
            { functionRegistry: registry }
        );
        expect(model.getValue(3, 0)).toBe(4);
        expect(model.getValue(3, 1)).toBe(5);
        model.setCell(1, 0, 4);
        expect(model.getValue(3, 0)).toBe(8);
        expect(model.getValue(3, 1)).toBe(7);
    });

    test("supports explicit built-in override and deterministic registry recalculation", () => {
        const registry = new FormulaFunctionRegistry();
        registry.register("SUM", () => 99, { overrideBuiltIn: true });
        let multiplier = 2;
        registry.register("scale", args => (typeof args[0] === "number" ? args[0] * multiplier : 0));
        const model = new SpreadsheetModel(columns, 1, [["Keyboard", 3, 10, "=SCALE(B1)"]], { functionRegistry: registry });
        expect(model.getValue(3, 0)).toBe(6);
        model.setCell(3, 0, "=SUM(1, 2)");
        expect(model.getValue(3, 0)).toBe(99);
        model.setCell(3, 0, "=SCALE(B1)");
        expect(model.getValue(3, 0)).toBe(6);
        multiplier = 4;
        expect(model.getValue(3, 0)).toBe(6);
        model.recalculateAll();
        expect(model.getValue(3, 0)).toBe(12);
        registry.unregister("scale");
        expect(model.getValue(3, 0)).toBe(12);
        model.invalidateAll();
        expect(model.getDisplayValue(3, 0)).toBe("#NAME?");
    });
});
