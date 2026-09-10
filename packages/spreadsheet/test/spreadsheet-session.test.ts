import { describe, expect, test } from "vitest";
import { SpreadsheetModel } from "../src/model.js";
import { SpreadsheetSession } from "../src/spreadsheet-session.js";

describe("SpreadsheetSession", () => {
    test("orders cell and structural edits in one undo/redo history", () => {
        const session = new SpreadsheetSession(new SpreadsheetModel([
            { id: "value", title: "Value", type: "number" },
            { id: "double", title: "Double", type: "number" },
        ], 2, [[1, "=A1*2"], [2, "=A2*2"]]), { selection: "A1" });

        session.executeCells({ id: "cell-1", edits: [{ location: [0, 0], before: 1, after: 3 }], selectionBefore: "A1", selectionAfter: "A1" });
        session.insertRows(1, 1, undefined, "A2");
        session.executeCells({ id: "cell-2", edits: [{ location: [0, 1], before: null, after: 5 }], selectionBefore: "A2", selectionAfter: "A2" });
        expect(session.model.rowCount).toBe(3);
        expect(session.model.getValue(1, 2)).toBe(4);

        session.history.undo();
        expect(session.model.getInput(0, 1)).toBeNull();
        session.history.undo();
        expect(session.model.rowCount).toBe(2);
        expect(session.model.getValue(1, 1)).toBe(4);
        session.history.undo();
        expect(session.model.getValue(1, 0)).toBe(2);

        session.history.redo();
        session.history.redo();
        session.history.redo();
        expect(session.model.rowCount).toBe(3);
        expect(session.model.getInput(0, 1)).toBe(5);
        expect(session.selection).toBe("A2");
    });

    test("rewrites formulas and supports column deletion undo", () => {
        const session = new SpreadsheetSession(new SpreadsheetModel([
            { id: "a", title: "A" }, { id: "b", title: "B" }, { id: "c", title: "C" },
        ], 1, [[1, 2, "=A1+B1"]]));
        session.deleteColumns(0, 1);
        expect(session.model.columns.map(column => column.id)).toEqual(["b", "c"]);
        expect(session.model.getInput(1, 0)).toContain("#REF!");
        session.history.undo();
        expect(session.model.columns.map(column => column.id)).toEqual(["a", "b", "c"]);
        expect(session.model.getInput(2, 0)).toBe("=A1+B1");
    });
});
