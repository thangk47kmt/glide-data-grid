import { describe, expect, test } from "vitest";
import {
    applyStructuralEdit,
    deleteColumns,
    deleteRows,
    insertColumns,
    insertRows,
    type StructuralWorkbookPayload,
} from "../src/structural.js";

const payload: StructuralWorkbookPayload = {
    columns: [
        { id: "a", title: "A" },
        { id: "b", title: "B" },
        { id: "c", title: "C" },
    ],
    rowCount: 3,
    rows: [
        ["r1", "=A1:A3", '="A1"&A1&[A1]'],
        ["r2", "=$A$2+$B1", "=A2"],
        ["r3", "=A3", null],
    ],
};

describe("structural workbook edits", () => {
    test("inserts rows, shifts formula coordinates, expands ranges, and preserves literals", () => {
        const original = structuredClone(payload);
        const result = insertRows(payload, 1, 1);
        expect(result.rowCount).toBe(4);
        expect(result.rows).toEqual([
            ["r1", "=A1:A4", '="A1"&A1&[A1]'],
            [null, null, null],
            ["r2", "=$A$3+$B1", "=A3"],
            ["r3", "=A4", null],
        ]);
        expect(result.payload).toEqual({ columns: result.columns, rows: result.rows, rowCount: result.rowCount });
        expect(result.afterPayload).toBe(result.payload);
        expect(result.beforePayload).toEqual(original);
        expect(payload).toEqual(original);
        expect(result.edits.length).toBeGreaterThan(0);
    });

    test("inserts columns and shifts absolute references as structural coordinates", () => {
        const result = insertColumns(payload, 1, [{ id: "new", title: "New" }]);
        expect(result.columns.map(column => column.id)).toEqual(["a", "new", "b", "c"]);
        expect(result.rows[0]).toEqual(["r1", null, "=A1:A3", '="A1"&A1&[A1]']);
        expect(result.rows[1]?.[2]).toBe("=$A$2+$C1");
    });

    test("deletes intersecting ranges by shrinking them and deleted references become REF", () => {
        const result = deleteRows(payload, 1, 1);
        expect(result.rows[0]?.[1]).toBe("=A1:A2");
        expect(result.rows[1]?.[1]).toBe("=A2");
        expect(result.rows[0]?.[2]).toBe('="A1"&A1&[A1]');

        const allReference = applyStructuralEdit(
            { columns: payload.columns, rowCount: 2, rows: [["=A2", null, null], [null, null, null]] },
            { type: "delete-rows", index: 1, count: 1 }
        );
        expect(allReference.rows[0]?.[0]).toBe("=#REF!");
        expect(allReference.beforePayload).toEqual({ columns: payload.columns, rowCount: 2, rows: [["=A2", null, null], [null, null, null]] });
    });

    test("rewrites billion-cell ranges in constant time and supports atomic undo payloads", () => {
        const huge = { columns: [{ id: "a", title: "A" }], rowCount: 2, rows: [["=A1:A1000000000"], [null]] };
        const result = deleteRows(huge, 1, 1);
        expect(result.rows[0]?.[0]).toBe("=A1:A999999999");
        expect(result.afterPayload).toEqual(result.payload);
        expect(result.beforePayload).toEqual(huge);
    });

    test("deletes columns, shifts formulas, and supports generated column definitions", () => {
        const result = deleteColumns(payload, 1, 1);
        expect(result.columns.map(column => column.id)).toEqual(["a", "c"]);
        expect(result.rows[0]?.[1]).toBe('="A1"&A1&[A1]');
        const inserted = insertColumns(payload, 0, 2);
        expect(inserted.columns.slice(0, 2).map(column => column.id)).toEqual(["inserted-column-1", "inserted-column-2"]);
    });

    test("validates dimensions, limits, and duplicate column ids", () => {
        expect(() => insertRows(payload, 4, 1)).toThrow();
        expect(() => deleteRows(payload, 2, 2)).toThrow();
        expect(() => insertRows(payload, 0, 1, { maxRows: 3 })).toThrow();
        expect(() => insertColumns(payload, 0, [{ id: "A", title: "duplicate" }])).toThrow(/Duplicate column id/);
        expect(() => insertColumns(payload, 0, 2, [{ id: "x", title: "X" }])).toThrow(/count does not match/);
    });
});
