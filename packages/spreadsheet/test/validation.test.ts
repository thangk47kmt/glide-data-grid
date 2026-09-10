import { describe, expect, test, vi } from "vitest";
import { validateCellInput, type ValidationRule } from "../src/validation.js";

describe("cell validation", () => {
    test("checks required values and reports structured errors", () => {
        const result = validateCellInput("", [{ kind: "required", id: "required", message: "Name is required" }]);
        expect(result).toEqual({
            valid: false,
            errors: [{ ruleId: "required", message: "Name is required", severity: "error" }],
            warnings: [],
        });
        expect(validateCellInput(null, [{ kind: "required", id: "r" }]).valid).toBe(false);
        expect(validateCellInput("ok", [{ kind: "required", id: "r" }]).valid).toBe(true);
    });

    test("validates number boundaries and rejects non-finite values", () => {
        const rule: ValidationRule = { kind: "number-range", id: "n", min: 1, max: 3 };
        expect(validateCellInput(1, [rule]).valid).toBe(true);
        expect(validateCellInput(3, [rule]).valid).toBe(true);
        expect(validateCellInput(0, [rule]).valid).toBe(false);
        expect(validateCellInput(Number.NaN, [rule]).valid).toBe(false);
        expect(validateCellInput(Number.POSITIVE_INFINITY, [rule]).valid).toBe(false);
    });

    test("reports invalid number and text range configuration distinctly", () => {
        const numberResult = validateCellInput(2, [{ kind: "number-range", id: "number-order", min: 4, max: 1 }]);
        expect(numberResult.errors).toEqual([{ ruleId: "number-order", message: "Validation rule configuration is invalid", severity: "error" }]);
        const textResult = validateCellInput("hello", [{ kind: "text-length", id: "text-order", min: 5, max: 2 }]);
        expect(textResult.errors).toEqual([{ ruleId: "text-order", message: "Validation rule configuration is invalid", severity: "error" }]);
    });

    test("validates text length, options, and case sensitivity", () => {
        expect(validateCellInput("abc", [{ kind: "text-length", id: "len", min: 3, max: 3 }]).valid).toBe(true);
        expect(validateCellInput("ab", [{ kind: "text-length", id: "len", min: 3 }]).valid).toBe(false);
        expect(validateCellInput("YES", [{ kind: "one-of", id: "list", values: ["yes"], caseSensitive: false }]).valid).toBe(true);
        expect(validateCellInput("YES", [{ kind: "one-of", id: "list", values: ["yes"], caseSensitive: true }]).valid).toBe(false);
    });

    test("uses a supplied RegExp without mutating its state", () => {
        const pattern = /^item-\d+$/gy;
        pattern.lastIndex = 4;
        expect(validateCellInput("item-12", [{ kind: "regex", id: "pattern", pattern }]).valid).toBe(true);
        expect(pattern.lastIndex).toBe(4);
        expect(validateCellInput("other", [{ kind: "regex", id: "pattern", pattern }]).valid).toBe(false);
    });

    test("validates strict ISO dates, leap day and inclusive range", () => {
        const rule: ValidationRule = { kind: "date-iso-range", id: "date", min: "2024-02-29", max: "2024-03-02" };
        expect(validateCellInput("2024-02-29", [rule]).valid).toBe(true);
        expect(validateCellInput("2024-02-30", [rule]).valid).toBe(false);
        expect(validateCellInput("2023-02-29", [rule]).valid).toBe(false);
        expect(validateCellInput("2024-03-03", [rule]).valid).toBe(false);
    });

    test("reports reversed ISO date range as invalid configuration", () => {
        const result = validateCellInput("2024-02-29", [{ kind: "date-iso-range", id: "date-order", min: "2024-03-02", max: "2024-03-01" }]);
        expect(result).toEqual({
            valid: false,
            errors: [{ ruleId: "date-order", message: "Validation rule configuration is invalid", severity: "error" }],
            warnings: [],
        });
    });

    test("has an explicit formula raw/computed/defer policy", () => {
        const rule: ValidationRule = { kind: "number-range", id: "n", min: 10, max: 20 };
        expect(validateCellInput("=A1", [rule])).toEqual({ valid: true, errors: [], warnings: [] });
        expect(validateCellInput("=A1", [rule], { formulaPolicy: "raw" }).valid).toBe(false);
        expect(validateCellInput("=A1", [rule], { computedValue: 15 }).valid).toBe(true);
        expect(validateCellInput("=A1", [rule], { formulaPolicy: "computed", computedValue: 25 }).valid).toBe(false);
    });

    test("catches custom predicate exceptions and preserves severity", () => {
        const predicate = vi.fn(() => {
            throw new Error("boom");
        });
        const result = validateCellInput("x", [{ kind: "custom", id: "custom", severity: "warning", message: "Not recommended", validate: predicate }]);
        expect(result).toEqual({
            valid: true,
            errors: [],
            warnings: [{ ruleId: "custom", message: "Not recommended", severity: "warning" }],
        });
        expect(predicate).toHaveBeenCalledOnce();
    });

    test("collects all failures by default and supports stop-on-first", () => {
        const rules: ValidationRule[] = [
            { kind: "required", id: "required" },
            { kind: "text-length", id: "length", min: 4 },
        ];
        expect(validateCellInput("", rules).errors).toHaveLength(2);
        expect(validateCellInput("", rules, { stopOnFirst: true }).errors).toHaveLength(1);
        expect(validateCellInput("", rules, { collectAll: false }).errors).toHaveLength(1);
    });
});
