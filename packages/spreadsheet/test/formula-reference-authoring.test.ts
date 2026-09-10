import { describe, expect, test } from "vitest";
import {
    formatFormulaReference,
    insertFormulaReference,
    insertFormulaReferenceTarget,
} from "../src/formula-reference-authoring.js";

describe("formula reference authoring", () => {
    test("inserts a cell reference at the caret", () => {
        expect(insertFormulaReference("=SUM()", "C5", { start: 5, end: 5 })).toEqual({
            formula: "=SUM(C5)",
            cursor: 7,
            referenceRange: { start: 5, end: 7 },
        });
    });

    test("replaces a live reference while a range is being dragged", () => {
        const first = insertFormulaReference("=SUM()", "C5", { start: 5, end: 5 });
        expect(insertFormulaReference(first.formula, "C5:E9", { start: first.cursor, end: first.cursor }, first.referenceRange)).toEqual({
            formula: "=SUM(C5:E9)",
            cursor: 10,
            referenceRange: { start: 5, end: 10 },
        });
    });

    test("formats a single cell and rectangular range as A1 references", () => {
        expect(formatFormulaReference({ kind: "cell", col: 0, row: 0 })).toBe("A1");
        expect(formatFormulaReference({ kind: "cell", col: 27, row: 3 })).toBe("AB4");
        expect(formatFormulaReference({ kind: "range", from: { col: 2, row: 3 }, to: { col: 0, row: 0 } })).toBe("A1:C4");
        expect(insertFormulaReferenceTarget("=SUM()", { kind: "range", from: { col: 0, row: 0 }, to: { col: 2, row: 3 } }, { start: 5, end: 5 }).formula).toBe("=SUM(A1:C4)");
    });

    test("formats explicit caption-plus-row cell references", () => {
        expect(formatFormulaReference({ kind: "caption-cell", columnName: "Số lượng", row: 2 })).toBe("[Số lượng]3");
        expect(insertFormulaReferenceTarget("=A1+", { kind: "caption-cell", columnName: "Số lượng", row: 0 }, { start: 4, end: 4 }).formula).toBe("=A1+[Số lượng]1");
    });

    test("formats a whole column using a structured caption", () => {
        expect(formatFormulaReference({ kind: "column", columnName: "Unit Price" })).toBe("[Unit Price]");
        expect(formatFormulaReference({ kind: "column", columnName: "Số lượng", currentRow: true })).toBe("[@Số lượng]");
        expect(insertFormulaReferenceTarget("=SUM(A1:A4)", { kind: "column", columnName: "Unit Price" }, { start: 5, end: 10 }).formula).toBe("=SUM([Unit Price])");
    });

    test("replaces the selected formula text with a structured column reference", () => {
        expect(insertFormulaReference("=old*2", "[@Số lượng]", { start: 1, end: 4 }).formula).toBe("=[@Số lượng]*2");
    });

    test("uses the caret or selected range and clamps out-of-bounds text positions", () => {
        expect(insertFormulaReferenceTarget("=A1+", { kind: "cell", col: 1, row: 1 }, { start: 4, end: 4 })).toMatchObject({ formula: "=A1+B2", cursor: 6 });
        expect(insertFormulaReferenceTarget("=A1+old", { kind: "cell", col: 1, row: 1 }, { start: 4, end: 7 })).toMatchObject({ formula: "=A1+B2", cursor: 6 });
        expect(insertFormulaReferenceTarget("=A1", { kind: "cell", col: 1, row: 1 }, { start: -5, end: 50 })).toMatchObject({ formula: "B2", cursor: 2 });
    });

    test("rejects invalid grid coordinates and structured captions", () => {
        expect(() => formatFormulaReference({ kind: "cell", col: -1, row: 0 })).toThrow(RangeError);
        expect(() => formatFormulaReference({ kind: "range", from: { col: 0, row: 0 }, to: { col: 1, row: -1 } })).toThrow(RangeError);
        expect(() => formatFormulaReference({ kind: "column", columnName: "  " })).toThrow(RangeError);
        expect(() => formatFormulaReference({ kind: "column", columnName: "Bad]Caption" })).toThrow(RangeError);
    });
});
