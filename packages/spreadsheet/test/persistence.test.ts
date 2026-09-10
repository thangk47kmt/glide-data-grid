import { describe, expect, test } from "vitest";
import {
    migrateSnapshot,
    parseSnapshot,
    rowsFromSnapshot,
    serializeSnapshot,
    snapshotFromRows,
    SnapshotError,
    type SpreadsheetSnapshotV1,
} from "../src/persistence.js";

const columns = [
    { id: "name", title: "Name", type: "text" as const },
    { id: "amount", title: "Amount", type: "number" as const },
];

describe("spreadsheet persistence", () => {
    test("serializes deterministically and round-trips sparse raw inputs", () => {
        const first: SpreadsheetSnapshotV1 = {
            version: 1,
            columns,
            rowCount: 3,
            cells: [
                { col: 1, row: 2, value: 9 },
                { col: 0, row: 0, value: "日本語" },
                { col: 1, row: 0, value: "=A1+1" },
            ],
            extensions: { z: true, a: "first" },
        };
        const second = { ...first, cells: [...first.cells].reverse() };
        const json = serializeSnapshot(first);
        expect(json).toBe(serializeSnapshot(second));
        expect(parseSnapshot(json)).toEqual({
            ...first,
            cells: [first.cells[1], first.cells[2], first.cells[0]],
            extensions: { a: "first", z: true },
        });
    });

    test("preserves FormulaError, booleans, numbers and null values", () => {
        const snapshot: SpreadsheetSnapshotV1 = {
            version: 1,
            columns: [{ id: "a", title: "A" }],
            rowCount: 1,
            cells: [
                { col: 0, row: 0, value: { kind: "error", code: "#VALUE!", message: "bad" } },
            ],
        };
        expect(parseSnapshot(serializeSnapshot(snapshot))).toEqual(snapshot);
        expect(snapshotFromRows(columns, [["x", 2], [null, ""]])).toEqual({
            version: 1,
            columns,
            rowCount: 2,
            cells: [{ col: 0, row: 0, value: "x" }, { col: 1, row: 0, value: 2 }],
        });
        expect(rowsFromSnapshot(snapshotFromRows(columns, [["x", 2], [null, ""]]))).toEqual([["x", 2], [null, null]]);
    });

    test("rejects duplicate ids/cells and out-of-bounds coordinates", () => {
        const base = { version: 1, columns, rowCount: 1, cells: [] };
        expect(() => migrateSnapshot({ ...base, columns: [{ id: "name", title: "N" }, { id: "NAME", title: "N2" }] })).toThrowError(SnapshotError);
        expect(() => migrateSnapshot({ ...base, cells: [{ col: 0, row: 0, value: 1 }, { col: 0, row: 0, value: 2 }] })).toThrowError(SnapshotError);
        expect(() => migrateSnapshot({ ...base, cells: [{ col: 2, row: 0, value: 1 }] })).toThrowError(SnapshotError);
        try {
            migrateSnapshot({ ...base, cells: [{ col: 0, row: 2, value: 1 }] });
        } catch (caught) {
            const error = caught as SnapshotError;
            expect(error.code).toBe("out-of-bounds-cell");
            expect(error.path).toBe("$.cells[0]");
        }
    });

    test("rejects untrusted shapes, non-finite values and invalid metadata", () => {
        expect(() => parseSnapshot("not json")).toThrowError(/valid JSON/);
        expect(() => migrateSnapshot({ version: 1, columns: [], rowCount: 0, cells: [{ col: 0, row: 0, value: NaN }] })).toThrowError(SnapshotError);
        expect(() => migrateSnapshot({ version: 1, columns: [], rowCount: 0, cells: [], formats: { value: [] } })).toThrowError(SnapshotError);
        expect(() => migrateSnapshot({ version: 1, columns: [{ id: "a", title: "A", width: Infinity }], rowCount: 0, cells: [] })).toThrowError(SnapshotError);
        expect(() => snapshotFromRows(columns, [["too", "many", "cells"]])).toThrowError(SnapshotError);
    });

    test("enforces limits with structured code/path and rejects future versions", () => {
        const json = serializeSnapshot(snapshotFromRows(columns, [["x", 1], ["y", 2]]));
        expect(() => parseSnapshot(json, { maxRows: 1 })).toThrowError(SnapshotError);
        expect(() => parseSnapshot(json, { maxColumns: 1 })).toThrowError(SnapshotError);
        expect(() => parseSnapshot(json, { maxCells: 1 })).toThrowError(SnapshotError);
        expect(() => parseSnapshot(json, { maxJsonLength: 2 })).toThrowError(SnapshotError);
        try {
            parseSnapshot(json, { maxCells: 1 });
        } catch (caught) {
            const error = caught as SnapshotError;
            expect(error.code).toBe("limit-exceeded");
            expect(error.path).toBe("$.cells");
        }
        try {
            migrateSnapshot({ version: 2, columns: [], rowCount: 0, cells: [] });
        } catch (caught) {
            const error = caught as SnapshotError;
            expect(error.code).toBe("unsupported-version");
            expect(error.path).toBe("$.version");
        }
    });
});
