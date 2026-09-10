import { describe, expect, test } from "vitest";
import { fillPattern, smartFill, type FillEdit } from "../src/fill.js";
import { PagedDataSource } from "../src/paged-data-source.js";
import { PagedFormulaAdapter, type PagedFormulaInput } from "../src/paged-formula-adapter.js";
import type { SpreadsheetColumn } from "../src/model.js";
import { History } from "../src/history.js";

const rect = (x: number, y: number, width: number, height: number) => ({ x, y, width, height });
const values = (edits: readonly FillEdit[]) => edits.map(([, , value]) => value);

describe("smart fill", () => {
    test("translates relative, mixed and absolute formula references", () => {
        const edits = smartFill([["=A1+$B1+C$2+$D$4"]], rect(2, 2, 1, 1), rect(3, 3, 2, 1));
        expect(edits).toEqual([
            [3, 3, "=B2+$B2+D$2+$D$4"],
            [4, 3, "=C2+$B2+E$2+$D$4"],
        ]);
    });

    test("fills horizontal and vertical number series", () => {
        expect(values(smartFill([[1, 3]], rect(0, 0, 2, 1), rect(2, 0, 4, 1)))).toEqual([5, 7, 9, 11]);
        expect(values(smartFill([[2], [5]], rect(0, 0, 1, 2), rect(0, 2, 1, 4)))).toEqual([8, 11, 14, 17]);
    });

    test("fills ISO dates without local-timezone arithmetic", () => {
        expect(values(smartFill([["2024-02-28"], ["2024-03-01"]], rect(0, 0, 1, 2), rect(0, 2, 1, 3)))).toEqual([
            "2024-03-03",
            "2024-03-05",
            "2024-03-07",
        ]);
    });

    test("repeats text, booleans, null and unsupported strings", () => {
        expect(values(smartFill([["hello", true], [null, "2024-02-30"]], rect(0, 0, 2, 2), rect(2, 0, 2, 2)))).toEqual([
            "hello",
            true,
            null,
            "2024-02-30",
        ]);
    });

    test("repeats a two-dimensional pattern and excludes source cells", () => {
        const edits = smartFill([["a", "b"], ["c", "d"]], rect(1, 1, 2, 2), rect(0, 0, 4, 4));
        expect(edits).toHaveLength(12);
        expect(edits).not.toContainEqual([1, 1, "a"]);
        expect(edits).toContainEqual([0, 0, "a"]);
        expect(edits).toContainEqual([3, 3, "d"]);
    });

    test("supports filling up and left", () => {
        expect(smartFill([[1, 2]], rect(2, 2, 2, 1), rect(0, 2, 2, 1))).toEqual([
            [0, 2, -1],
            [1, 2, 0],
        ]);
        expect(smartFill([[1], [2]], rect(2, 2, 1, 2), rect(2, 0, 1, 2))).toEqual([
            [2, 0, -1],
            [2, 1, 0],
        ]);
    });

    test("does not mutate the pattern and exposes the alias", () => {
        const pattern: (string | number)[][] = [[1, 2]];
        const before = pattern.map(row => [...row]);
        expect(fillPattern(pattern, rect(0, 0, 2, 1), rect(0, 1, 2, 1))).toEqual([
            [0, 1, 1],
            [1, 1, 2],
        ]);
        expect(pattern).toEqual(before);
    });

    test("rejects invalid and ragged rectangles/patterns", () => {
        expect(() => smartFill([], rect(0, 0, 1, 1), rect(0, 1, 1, 1))).toThrow(RangeError);
        expect(() => smartFill([[1]], rect(0, 0, 2, 1), rect(0, 1, 2, 1))).toThrow(RangeError);
        expect(() => smartFill([[1], [2, 3]], rect(0, 0, 1, 2), rect(0, 2, 1, 1))).toThrow(RangeError);
        expect(() => smartFill([[1]], rect(-1, 0, 1, 1), rect(0, 1, 1, 1))).toThrow(RangeError);
    });

    test("fills one formula into many selected cells and evaluates each shifted reference", () => {
        const columns: readonly SpreadsheetColumn[] = [
            { id: "source", title: "Source", type: "number" },
            { id: "result", title: "Result", type: "number" },
        ];
        const source = new PagedDataSource({
            rowCount: 4,
            columnCount: 2,
            generateCell: (row, col) => col === 0 ? row + 1 : null,
        });
        const formulas = new PagedFormulaAdapter(source, columns);
        const edits = smartFill([["=A1*2"]], rect(1, 0, 1, 1), rect(1, 1, 1, 3));
        expect(edits).toEqual([
            [1, 1, "=A2*2"],
            [1, 2, "=A3*2"],
            [1, 3, "=A4*2"],
        ]);
        const history = new History<PagedFormulaInput>({
            apply: (edit, direction) => formulas.setCell(edit.location[0], edit.location[1], direction === "undo" ? edit.before : edit.after),
        });
        const transaction = history.execute({
            id: "paste-single-formula",
            edits: edits.map(([col, row, value]) => ({ location: [col, row] as const, before: null, after: value as string })),
        });
        expect(transaction?.edits).toHaveLength(3);
        expect(history.undoCount).toBe(1);
        expect(edits.map(([col, row]) => formulas.getValue(col, row))).toEqual([4, 6, 8]);
        history.undo();
        expect(edits.map(([col, row]) => formulas.getValue(col, row))).toEqual([null, null, null]);
        history.redo();
        expect(edits.map(([col, row]) => formulas.getValue(col, row))).toEqual([4, 6, 8]);
    });

    test("pastes a one-by-N formula source into a shifted matching region with relative references", () => {
        const pattern = [["=A1", "=B1"]];
        const edits = smartFill(pattern, rect(0, 0, 2, 1), rect(2, 1, 2, 1));
        expect(edits).toEqual([
            [2, 1, "=C2"],
            [3, 1, "=D2"],
        ]);
    });

    test("pastes a one-by-N source into a wider region and shifts each repeated tile", () => {
        const edits = smartFill([["=A1", "=B1"]], rect(2, 0, 2, 1), rect(2, 1, 4, 1));
        expect(edits).toEqual([
            [2, 1, "=A2"],
            [3, 1, "=B2"],
            [4, 1, "=C2"],
            [5, 1, "=D2"],
        ]);
    });

    test("pastes an N-by-M source into a larger region, repeating and shifting every tile", () => {
        const pattern = [
            ["=A1", "=B1"],
            ["=A2", "=B2"],
        ];
        const edits = smartFill(pattern, rect(2, 0, 2, 2), rect(2, 2, 4, 4));
        expect(edits).toHaveLength(16);
        expect(edits).toContainEqual([2, 2, "=A3"]);
        expect(edits).toContainEqual([3, 2, "=B3"]);
        expect(edits).toContainEqual([4, 2, "=C3"]);
        expect(edits).toContainEqual([5, 3, "=D4"]);
        expect(edits).not.toContainEqual([2, 0, "=A1"]);
    });

    test("rejects a multi-cell paste atomically when one destination has the wrong type", () => {
        const columns: readonly SpreadsheetColumn[] = [
            { id: "text", title: "Text", type: "text" },
            { id: "number", title: "Number", type: "number" },
            { id: "resultA", title: "Result A", type: "number" },
            { id: "resultB", title: "Result B", type: "number" },
        ];
        const source = new PagedDataSource({
            rowCount: 2,
            columnCount: 4,
            generateCell: (_row, col) => col === 0 ? "text value" : col === 1 ? 10 : null,
        });
        const formulas = new PagedFormulaAdapter(source, columns);
        const history = new History<PagedFormulaInput>({
            apply: edit => formulas.setCell(edit.location[0], edit.location[1], edit.after),
        });
        const edits = smartFill([["=A1", "=B1"]], rect(2, 0, 2, 1), rect(2, 1, 2, 1));
        const validations = edits.map(([col, row, value]) => formulas.validateFormula(col, row, value as string));
        expect(validations[0]).toMatchObject({ valid: false, error: { code: "#VALUE!" } });
        expect(validations[0]?.error?.message).toContain("source and target types must match");
        expect(validations[1]).toMatchObject({ valid: true, value: 10 });

        // Paste is atomic: validate every destination before applying any
        // edit, so the valid second cell cannot partially overwrite the grid.
        if (validations.every(validation => validation.valid)) {
            history.execute({
                id: "paste-invalid-type",
                edits: edits.map(([col, row, value]) => ({ location: [col, row] as const, before: null, after: value as string })),
            });
        }
        expect(formulas.hasFormula(2, 1)).toBe(false);
        expect(formulas.hasFormula(3, 1)).toBe(false);
        expect(formulas.getValue(2, 1)).toBeNull();
        expect(formulas.getValue(3, 1)).toBeNull();
        expect(formulas.getFormulaCount()).toBe(0);
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(0);
    });
});
