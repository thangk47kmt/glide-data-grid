import { describe, expect, test } from "vitest";
import { formulaError } from "../src/formula.js";
import {
    dropdownGridCell,
    planGridPaste,
    shouldOpenInlineFormulaEditorOnDoubleClick,
} from "../src/large-dataset-contracts.js";

const base = {
    pageRows: [10, 11, 12, 13],
    columnCount: 6,
    getBefore: (col: number, row: number) => `before:${col}:${row}`,
    parse: (value: string) => value,
    validateFormula: () => undefined,
};

describe("large dataset interaction contracts", () => {
    test("expands a clipboard matrix from one selected cell and rebases remembered formulas", () => {
        const values = [["=A1", "=B1"], ["=A2", "=B2"]];
        const result = planGridPaste({
            ...base,
            target: [2, 2],
            values,
            rememberedCopy: { x: 0, y: 0, width: 2, height: 2, sourceRows: [0, 1], values },
        });
        expect(result).toEqual({
            valid: true,
            edits: [
                { col: 2, row: 12, before: "before:2:12", after: "=C13" },
                { col: 3, row: 12, before: "before:3:12", after: "=D13" },
                { col: 2, row: 13, before: "before:2:13", after: "=C14" },
                { col: 3, row: 13, before: "before:3:13", after: "=D14" },
            ],
        });
    });

    test("tiles a copied formula over a larger explicit selection", () => {
        const values = [["=A1"]];
        const result = planGridPaste({
            ...base,
            target: [2, 0],
            selection: { x: 2, y: 0, width: 2, height: 2 },
            values,
            rememberedCopy: { x: 0, y: 0, width: 1, height: 1, sourceRows: [0], values },
        });
        expect(result.valid).toBe(true);
        if (result.valid) expect(result.edits.map(edit => [edit.col, edit.row, edit.after])).toEqual([
            [2, 10, "=C11"], [3, 10, "=D11"],
            [2, 11, "=C12"], [3, 11, "=D12"],
        ]);
    });

    test("does not rebase an external clipboard formula with no matching remembered copy", () => {
        const result = planGridPaste({ ...base, target: [2, 0], values: [["=A1"]] });
        expect(result.valid).toBe(true);
        if (result.valid) expect(result.edits[0]?.after).toBe("=A1");
    });

    test("rejects the whole plan with a specific validation message", () => {
        const result = planGridPaste({
            ...base,
            target: [1, 0],
            values: [["=A1", "=B1"]],
            validateFormula: col => col === 2 ? formulaError("#VALUE!", "Text cannot be stored in Number") : undefined,
        });
        expect(result).toEqual({
            valid: false,
            edits: [],
            col: 2,
            row: 10,
            value: "=B1",
            message: "#VALUE!: Text cannot be stored in Number",
        });
    });

    test("skips protected/loading columns in the planned transaction", () => {
        const result = planGridPaste({
            ...base,
            target: [0, 0],
            values: [["one", "protected", "loading", "four"]],
            canWrite: col => col !== 1 && col !== 2,
        });
        expect(result.valid).toBe(true);
        if (result.valid) expect(result.edits.map(edit => [edit.col, edit.after])).toEqual([[0, "one"], [3, "four"]]);
    });

    test("dropdown cells use the grid's default second-click activation", () => {
        const cell = dropdownGridCell("Approved", ["Approved", "Blocked"], "Approved");
        expect(cell).toMatchObject({
            kind: "custom",
            allowOverlay: true,
            copyData: "Approved",
            data: { kind: "dropdown-cell", value: "Approved", allowedValues: ["Approved", "Blocked"] },
        });
        expect("activationBehaviorOverride" in cell).toBe(false);
        expect(shouldOpenInlineFormulaEditorOnDoubleClick(false, "dropdown")).toBe(false);
        expect(shouldOpenInlineFormulaEditorOnDoubleClick(true, "dropdown")).toBe(false);
        expect(shouldOpenInlineFormulaEditorOnDoubleClick(true, "drilldown")).toBe(false);
        expect(shouldOpenInlineFormulaEditorOnDoubleClick(true, "multi-select")).toBe(false);
        expect(shouldOpenInlineFormulaEditorOnDoubleClick(true, "number")).toBe(true);
    });
});
