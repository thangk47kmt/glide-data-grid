import { describe, expect, test } from "vitest";
import { History } from "../src/history.js";
import { SpreadsheetModel } from "../src/model.js";
import type { CellInput } from "../src/model.js";

describe("spreadsheet story history integration", () => {
    test("applies before on undo and after on redo, including formula recalculation", () => {
        const model = new SpreadsheetModel([
            { id: "value", title: "Value", type: "number" },
            { id: "total", title: "Total", type: "number" },
        ], 1, [[10, "=A1*2"]]);
        const history = new History<CellInput>({
            apply: (edit, direction) => model.setCell(edit.location[0], edit.location[1], direction === "undo" ? edit.before : edit.after),
        });
        history.execute({ id: "edit", edits: [{ location: [0, 0], before: 10, after: 15 }] });
        expect(model.getValue(0, 0)).toBe(15);
        expect(model.getValue(1, 0)).toBe(30);
        expect(model.getFormulaCount()).toBe(1);
        history.undo();
        expect(model.getValue(0, 0)).toBe(10);
        expect(model.getValue(1, 0)).toBe(20);
        expect(model.getFormulaCount()).toBe(1);
        history.redo();
        expect(model.getValue(0, 0)).toBe(15);
        expect(model.getValue(1, 0)).toBe(30);
        expect(model.getFormulaCount()).toBe(1);
    });
});
