import { describe, expect, test } from "vitest";
import { SpreadsheetModel } from "../src/model.js";
import { FormulaFunctionRegistry } from "../src/function-registry.js";
import {
    exportModelToCsv,
    importCsvToWorkbook,
    restoreModelSnapshot,
    serializeModelSnapshot,
    WorkbookIoError,
} from "../src/workbook-io.js";

const columns = [
    { id: "item", title: "Item", type: "text" as const },
    { id: "quantity", title: "Quantity", type: "number" as const },
    { id: "total", title: "Total", type: "number" as const },
];

describe("workbook I/O", () => {
    test("exports headers and preserves raw formula inputs by default", () => {
        const model = new SpreadsheetModel(columns, 1, [["Keyboard", 2, "=B1*10"]]);
        expect(exportModelToCsv(model)).toBe("Item,Quantity,Total\r\nKeyboard,2,=B1*10");
        expect(exportModelToCsv(model, { includeHeaders: false, lineEnding: "\n" })).toBe("Keyboard,2,=B1*10");
    });

    test("exports computed formula values when requested", () => {
        const model = new SpreadsheetModel(columns, 1, [["Keyboard", 2, "=B1*10"]]);
        expect(exportModelToCsv(model, { valueMode: "values", lineEnding: "\n" })).toBe("Item,Quantity,Total\nKeyboard,2,20");
        expect(exportModelToCsv(model, { formulaPolicy: "values", lineEnding: "\n" })).toBe("Item,Quantity,Total\nKeyboard,2,20");
    });

    test("does not evaluate formulas during default raw export", () => {
        const registry = new FormulaFunctionRegistry();
        let calls = 0;
        registry.register("COUNTME", () => {
            calls++;
            return 7;
        });
        const model = new SpreadsheetModel(columns, 1, [["Keyboard", 2, "=COUNTME()"]], { functionRegistry: registry });
        expect(exportModelToCsv(model, { lineEnding: "\n" })).toBe("Item,Quantity,Total\nKeyboard,2,=COUNTME()");
        expect(calls).toBe(0);
        expect(exportModelToCsv(model, { valueMode: "values", lineEnding: "\n" })).toBe("Item,Quantity,Total\nKeyboard,2,7");
        expect(calls).toBe(1);
    });

    test("imports normalized rows, preserves supplied metadata and keeps formulas raw", () => {
        const payload = importCsvToWorkbook("Item,Quantity,Total\nKeyboard,2,=B2*10\nMouse,3", { columns });
        expect(payload.columns).toEqual(columns);
        expect(payload.rowCount).toBe(2);
        expect(payload.rows).toEqual([["Keyboard", "2", "=B2*10"], ["Mouse", "3", null]]);
        expect(importCsvToWorkbook("a,b\n1,TRUE", { hasHeaders: false, inferTypes: true }).columns).toEqual([
            { id: "column-1", title: "Column 1" },
            { id: "column-2", title: "Column 2" },
        ]);
    });

    test("honors explicit formula rejection and row/column limits", () => {
        expect(() => importCsvToWorkbook("Item,Quantity\na,=B2", { formulaPolicy: "reject" })).toThrow(/Formula-looking/);
        expect(() => importCsvToWorkbook("a,b\n1,2\n3,4", { maxRows: 2 })).toThrow();
        expect(() => importCsvToWorkbook("a,b\n1,2", { columns: [{ id: "a", title: "A" }] })).toThrowError(WorkbookIoError);
        try {
            importCsvToWorkbook("a,b\n1,2", { columns: [{ id: "a", title: "A" }] });
        } catch (caught) {
            const error = caught as WorkbookIoError;
            expect(error.code).toBe("column-count-mismatch");
            expect(error.path).toBe("$.header");
        }
    });

    test("serializes and restores model raw inputs through persistence helpers", () => {
        const model = new SpreadsheetModel(columns, 2, [["Keyboard", 2, "=B1*10"], [null, null, null]]);
        const json = serializeModelSnapshot(model, { extensions: { source: "test" } });
        const restored = restoreModelSnapshot(json);
        expect(restored.columns).toEqual(columns);
        expect(restored.rowCount).toBe(2);
        expect(restored.getInput(2, 0)).toBe("=B1*10");
        expect(restored.getValue(2, 0)).toBe(20);
        expect(restored.getInput(0, 1)).toBe(null);
    });

    test("failed snapshot restore leaves the current model untouched", () => {
        const model = new SpreadsheetModel(columns, 1, [["Keyboard", 2, "=B1*10"]]);
        expect(() => restoreModelSnapshot('{"version":2}')).toThrow();
        expect(model.getInput(0, 0)).toBe("Keyboard");
        expect(model.getInput(2, 0)).toBe("=B1*10");
    });
});
