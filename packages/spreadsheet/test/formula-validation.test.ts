import { describe, expect, test } from "vitest";
import { formulaColumnSemanticType, formulaValueType, validateFormulaReferenceType, validateFormulaResultType } from "../src/formula-validation.js";
import { PagedDataSource } from "../src/paged-data-source.js";
import { PagedFormulaAdapter, type PagedFormulaInput } from "../src/paged-formula-adapter.js";
import type { SpreadsheetColumn } from "../src/model.js";
import type { FormulaValue } from "../src/formula.js";
import { validateCellInput, type ValidationRule } from "../src/validation.js";
import { History } from "../src/history.js";

const columns: readonly SpreadsheetColumn[] = [
    { id: "text", title: "Text", type: "text" },
    { id: "number", title: "Number", type: "number" },
    { id: "boolean", title: "Boolean", type: "boolean" },
    { id: "result", title: "Result", type: "number" },
    { id: "textResult", title: "Text result", type: "text" },
    { id: "numberResult", title: "Number result", type: "number" },
    { id: "booleanResult", title: "Boolean result", type: "boolean" },
];

function adapter() {
    const source = new PagedDataSource({
        rowCount: 2,
        columnCount: columns.length,
        pageSize: 2,
        generateCell: (_row, col) => ["00123", 12, true, null, null, null, null][col] ?? null,
    });
    return new PagedFormulaAdapter(source, columns);
}

function errorOf(value: FormulaValue | { readonly kind: "loading" }) {
    expect(value).toMatchObject({ kind: "error" });
    return value as Extract<FormulaValue, { readonly kind: "error" }>;
}

describe("formula type validation", () => {
    test("recognizes scalar and blank values without coercion", () => {
        expect(formulaValueType("00123")).toBe("text");
        expect(formulaValueType(123)).toBe("number");
        expect(formulaValueType(true)).toBe("boolean");
        expect(formulaValueType(null)).toBe("null");
    });

    test("rejects a direct reference when source and destination declarations differ", () => {
        const error = validateFormulaReferenceType("00123", columns[0]!, columns[1]!);
        expect(error).toMatchObject({ code: "#VALUE!" });
        expect(error?.message).toContain("source and destination types must match");
        expect(error?.message).toContain("Text");
        expect(error?.message).toContain("Number");
    });

    test("rejects a formula whose computed result type differs from the destination", () => {
        const formulas = adapter();
        const result = formulas.validateFormula(3, 0, "=TRUE");
        expect(result.valid).toBe(false);
        expect(result.error?.code).toBe("#VALUE!");
        expect(result.error?.message).toContain("Formula result type 'boolean'");
        expect(result.error?.message).toContain("does not match column 'Result' type 'number'");
    });

    test("accepts matching text, number, and boolean results", () => {
        expect(validateFormulaResultType("Invoice", columns[0]!)).toBeUndefined();
        expect(validateFormulaResultType(12, columns[1]!)).toBeUndefined();
        expect(validateFormulaResultType(true, columns[2]!)).toBeUndefined();
    });

    test("rejects semantic source/destination mismatches even when both values are strings", () => {
        const uri: SpreadsheetColumn = { id: "uri", title: "URI", type: "text", dataType: "uri" };
        const text: SpreadsheetColumn = { id: "text", title: "Text", type: "text", dataType: "text" };
        const error = validateFormulaReferenceType("https://example.com", uri, text);
        expect(error?.code).toBe("#VALUE!");
        expect(error?.message).toContain("source and destination types must match");
        expect(error?.message).toContain("uri");
        expect(error?.message).toContain("text");
    });

    test("adapter dry-run rejects a URI reference into a plain text column", () => {
        const semanticColumns: readonly SpreadsheetColumn[] = [
            { id: "uri", title: "URI", type: "text", dataType: "uri" },
            { id: "text", title: "Text", type: "text", dataType: "text" },
        ];
        const source = new PagedDataSource({ rowCount: 1, columnCount: 2, generateCell: (_row, col) => col === 0 ? "https://example.com" : null });
        const formulas = new PagedFormulaAdapter(source, semanticColumns);
        const result = formulas.validateFormula(1, 0, "=A1");
        expect(result.valid).toBe(false);
        expect(result.error?.code).toBe("#VALUE!");
        expect(result.error?.message).toContain("source and target types must match");
        expect(formulas.hasFormula(1, 0)).toBe(false);
    });

    test("adapter rejects plain text into a URI column despite the shared string runtime type", () => {
        const semanticColumns: readonly SpreadsheetColumn[] = [
            { id: "text", title: "Text", type: "text", dataType: "text" },
            { id: "uri", title: "URI", type: "text", dataType: "uri" },
        ];
        const source = new PagedDataSource({ rowCount: 1, columnCount: 2, generateCell: (_row, col) => col === 0 ? "https://example.com" : null });
        const formulas = new PagedFormulaAdapter(source, semanticColumns);
        const result = formulas.validateFormula(1, 0, "=A1");
        expect(result.valid).toBe(false);
        expect(result.error?.code).toBe("#VALUE!");
        expect(result.error?.message).toContain("Text");
        expect(result.error?.message).toContain("URI");
        expect(result.error?.message).toContain("types must match");
    });

    test.each([
        ["URI", { id: "uri", title: "URI", type: "text", dataType: "uri" }, "not-a-uri", "not a valid URI"],
        ["date", { id: "date", title: "Date", type: "text", dataType: "date" }, "2024-02-30", "not a valid ISO date"],
        ["dropdown", { id: "status", title: "Status", type: "text", dataType: "dropdown", allowedValues: ["Queued", "Approved"] }, "Unknown", "not an allowed value"],
        ["range", { id: "range", title: "Range", type: "number", dataType: "range", min: 0, max: 100 }, 101, "outside range"],
        ["stars", { id: "stars", title: "Stars", type: "number", dataType: "stars", min: 0, max: 5 }, 5.5, "must be an integer"],
    ] as const)("reports invalid %s semantic result", (_name, destination, value, message) => {
        const error = validateFormulaResultType(value, destination);
        expect(error?.code).toBe("#VALUE!");
        expect(error?.message).toContain(message);
    });

    test("allows blank results by default and reports a strict null policy", () => {
        expect(validateFormulaResultType(null, columns[1]!)).toBeUndefined();
        const error = validateFormulaResultType(null, columns[1]!, { allowNull: false });
        expect(error).toMatchObject({ code: "#VALUE!" });
        expect(error?.message).toContain("does not allow blank values");
    });

    test("preserves formula errors instead of replacing their specific message", () => {
        const sourceError = { kind: "error", code: "#DIV/0!", message: "Cannot divide by zero" } as const;
        expect(validateFormulaResultType(sourceError, columns[1]!)).toBe(sourceError);
    });

    test("supports semantic metadata and validates URI/date/dropdown/range/stars domains", () => {
        const uri = { id: "uri", title: "URI", type: "text" as const, dataType: "uri" as const };
        const date = { id: "date", title: "Date", type: "text" as const, semanticType: "date" as const };
        const dropdown = { id: "status", title: "Status", type: "text" as const, dataType: "dropdown" as const, allowedValues: ["Open", "Closed"] };
        const range = { id: "range", title: "Range", type: "number" as const, dataType: "range" as const, min: 0, max: 100 };
        const stars = { id: "stars", title: "Stars", type: "number" as const, dataType: "stars" as const, min: 0, max: 5 };
        expect(formulaColumnSemanticType(uri)).toBe("uri");
        expect(formulaColumnSemanticType(date)).toBe("date");
        expect(validateFormulaResultType("https://example.com", uri)).toBeUndefined();
        expect(validateFormulaResultType("not a uri", uri)?.message).toContain("valid URI");
        expect(validateFormulaResultType("2026-02-28", date)).toBeUndefined();
        expect(validateFormulaResultType("2026-02-30", date)?.message).toContain("valid ISO date");
        expect(validateFormulaResultType("Open", dropdown)).toBeUndefined();
        expect(validateFormulaResultType("Pending", dropdown)?.message).toContain("allowed value");
        expect(validateFormulaResultType(50, range)).toBeUndefined();
        expect(validateFormulaResultType(101, range)?.message).toContain("outside range");
        expect(validateFormulaResultType(5, stars)).toBeUndefined();
        expect(validateFormulaResultType(5.5, stars)?.message).toContain("integer");
        expect(validateFormulaReferenceType("https://example.com", uri, date)?.message).toContain("types must match");
    });
});

describe("formula evaluation error messages", () => {
    test.each([
        ["syntax error", "=1+", "#VALUE!", "Unexpected token"],
        ["unknown caption", "=[Missing]1", "#REF!", "Unknown column 'Missing'"],
        ["unknown A1 reference", "=ZZ1", "#REF!", "outside the data source"],
        ["out-of-range row", "=A999", "#REF!", "outside the data source"],
        ["divide by zero", "=1/0", "#DIV/0!", "divide by zero"],
    ])("reports a specific %s", (_name, formula, code, message) => {
        const formulas = adapter();
        const result = formulas.validateFormula(3, 0, formula);
        const error = errorOf(result.error ?? formulas.getValue(3, 0));
        expect(error.code).toBe(code);
        expect(error.message?.toLowerCase()).toContain(message.toLowerCase());
    });

    test("reports a circular reference with the cycle code and coordinates", () => {
        const formulas = adapter();
        formulas.setFormula(3, 0, "=D1");
        const error = errorOf(formulas.getValue(3, 0));
        expect(error.code).toBe("#CYCLE!");
        expect(error.message).toContain("Circular reference at 3:0");
    });

    test("evaluates successful text, number, and boolean direct references", () => {
        const formulas = adapter();
        formulas.setFormula(4, 0, "=A1");
        formulas.setFormula(5, 0, "=B1");
        formulas.setFormula(6, 0, "=C1");
        expect(formulas.getValue(4, 0)).toBe("00123");
        expect(formulas.getValue(5, 0)).toBe(12);
        expect(formulas.getValue(6, 0)).toBe(true);
    });

    test("keeps an invalid dry-run out of formula state", () => {
        const formulas = adapter();
        formulas.setFormula(3, 0, "=B1");
        const beforeStats = formulas.getStats();
        const rejected = formulas.validateFormula(3, 0, "=A1");

        expect(rejected.valid).toBe(false);
        expect(rejected.error?.message).toContain("source and target types must match");
        expect(formulas.getFormula(3, 0)).toBe("=B1");
        expect(formulas.getValue(3, 0)).toBe(12);
        expect(formulas.getFormulaCount()).toBe(beforeStats.formulaCells);
        expect(formulas.getStats().formulaCells).toBe(beforeStats.formulaCells);
    });

    test("does not record or apply an invalid formula commit", () => {
        const formulas = adapter();
        formulas.setFormula(3, 0, "=B1");
        const history = new History<PagedFormulaInput>({
            apply: edit => formulas.setCell(edit.location[0], edit.location[1], edit.after),
        });
        const beforeFormula = formulas.getFormula(3, 0);
        const beforeValue = formulas.getValue(3, 0);
        const validation = formulas.validateFormula(3, 0, "=A1");
        if (validation.valid) {
            history.execute({
                id: "formula-edit",
                edits: [{ location: [3, 0], before: beforeFormula ?? null, after: "=A1" }],
            });
        }

        expect(validation.valid).toBe(false);
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(0);
        expect(formulas.getFormula(3, 0)).toBe(beforeFormula);
        expect(formulas.getValue(3, 0)).toBe(beforeValue);
    });

    test("regresses E1 number <- A1 text: Enter/commit surfaces an error and leaves state untouched", () => {
        const e1Columns: readonly SpreadsheetColumn[] = [
            { id: "a", title: "A text", type: "text" },
            { id: "b", title: "B", type: "number" },
            { id: "c", title: "C", type: "number" },
            { id: "d", title: "D", type: "number" },
            { id: "e", title: "E total", type: "number" },
        ];
        const source = new PagedDataSource({
            rowCount: 1,
            columnCount: e1Columns.length,
            generateCell: (_row, col) => ["Record 000001", 10, 2, null, null][col] ?? null,
        });
        const formulas = new PagedFormulaAdapter(source, e1Columns);
        const history = new History<PagedFormulaInput>({
            apply: edit => formulas.setCell(edit.location[0], edit.location[1], edit.after),
        });
        const beforeFormula = formulas.getFormula(4, 0);
        const beforeValue = formulas.getValue(4, 0);
        let status = "";

        // This is the same guard used by the Enter/✓ commit path: validation
        // happens before the formula is written or added to undo history.
        const commitFromEditor = (draft: string): boolean => {
            const validation = formulas.validateFormula(4, 0, draft);
            if (!validation.valid) {
                const error = validation.error;
                status = error === undefined ? "Formula is invalid" : `${error.code}: ${error.message ?? "Formula is invalid"}`;
                return false;
            }
            history.execute({ id: "formula-edit", edits: [{ location: [4, 0], before: beforeFormula ?? null, after: draft }] });
            return true;
        };

        expect(commitFromEditor("=A1")).toBe(false);
        expect(status).toContain("#VALUE!");
        expect(status).toContain("source and target types must match");
        expect(formulas.getFormula(4, 0)).toBe(beforeFormula);
        expect(formulas.getValue(4, 0)).toBe(beforeValue);
        expect(formulas.getFormulaCount()).toBe(0);
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(0);
    });
});

describe("semantic value validation used by formula commits", () => {
    const invalidSemanticCases: readonly [string, ValidationRule, FormulaValue, string][] = [
        ["URI", { kind: "regex", id: "uri", pattern: /^https?:\/\// }, "not-a-uri", "Text does not match"],
        ["date", { kind: "date-iso-range", id: "date", min: "2020-01-01", max: "2030-12-31" }, "2024-02-30", "valid ISO date"],
        ["dropdown", { kind: "one-of", id: "status", values: ["Queued", "Approved"] }, "Unknown", "allowed option"],
        ["range", { kind: "number-range", id: "range", min: 0, max: 100 }, 101, "number in the allowed range"],
        ["stars", { kind: "custom", id: "stars", message: "Stars must be a whole number from 0 to 5", validate: value => typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= 5 }, 5.5, "Stars must be a whole number"],
    ];

    test.each(invalidSemanticCases)("reports a specific invalid %s value", (_name, rule, computedValue, message) => {
        const result = validateCellInput("=A1", [rule], { formulaPolicy: "computed", computedValue });
        expect(result.valid).toBe(false);
        expect(result.errors[0]?.message).toContain(message);
    });
});
