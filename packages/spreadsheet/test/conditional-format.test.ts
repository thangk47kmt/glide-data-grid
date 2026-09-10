import { describe, expect, test, vi } from "vitest";
import {
    compileConditionalFormatting,
    evaluateConditionalFormatting,
    ConditionalFormatError,
    type ConditionalFormatRule,
} from "../src/conditional-format.js";

const range = { x: 0, y: 0, width: 3, height: 2 };

describe("conditional formatting", () => {
    test("matches numeric comparisons and inclusive between rules", () => {
        const rules: ConditionalFormatRule[] = [
            { id: "high", kind: "number-compare", range, style: { bg: "red" }, operator: "gt", value: 10 },
            { id: "middle", kind: "number-between", range, style: { text: "amber" }, min: 5, max: 10 },
        ];
        expect(evaluateConditionalFormatting(0, 0, 12, rules)).toEqual({ style: { bg: "red" }, matchedRuleIds: ["high"], diagnostics: [] });
        expect(evaluateConditionalFormatting(0, 0, 10, rules)).toEqual({ style: { text: "amber" }, matchedRuleIds: ["middle"], diagnostics: [] });
    });

    test("matches text, empty and not-empty deterministically", () => {
        const rules: ConditionalFormatRule[] = [
            { id: "contains", kind: "text", range, style: { text: "blue" }, operator: "contains", value: "ABC" },
            { id: "empty", kind: "empty", range, style: { bg: "gray" } },
        ];
        expect(evaluateConditionalFormatting(0, 0, "xxabcxx", rules).matchedRuleIds).toEqual(["contains"]);
        expect(evaluateConditionalFormatting(0, 0, null, rules).matchedRuleIds).toEqual(["empty"]);
        expect(evaluateConditionalFormatting(0, 0, { kind: "error", code: "#REF!" }, [{ id: "empty", kind: "empty", range, style: { bg: "gray" } }]).matchedRuleIds).toEqual([]);
        expect(evaluateConditionalFormatting(0, 0, "", [{ id: "nonempty", kind: "not-empty", range, style: { bg: "gray" } }]).matchedRuleIds).toEqual([]);
    });

    test("indexes duplicate and unique ranges once", () => {
        const values = new Map([["0:0", "a"], ["1:0", "a"], ["2:0", "b"], ["0:1", "c"], ["1:1", "d"], ["2:1", "e"]]);
        const getCellValue = vi.fn((col: number, row: number) => values.get(`${col}:${row}`) ?? null);
        const rules: ConditionalFormatRule[] = [
            { id: "dupe", kind: "duplicate", range, style: { bg: "red" } },
            { id: "unique", kind: "unique", range, style: { text: "green" } },
        ];
        const engine = compileConditionalFormatting(rules, getCellValue);
        expect(getCellValue).toHaveBeenCalledTimes(6);
        expect(engine.evaluate(0, 0, "a").matchedRuleIds).toEqual(["dupe"]);
        expect(engine.evaluate(2, 0, "b").matchedRuleIds).toEqual(["unique"]);
        expect(getCellValue).toHaveBeenCalledTimes(6);

        const lazy = compileConditionalFormatting(rules);
        expect(lazy.evaluate(0, 0, "a", { getCellValue }).matchedRuleIds).toEqual(["dupe"]);
        expect(getCellValue).toHaveBeenCalledTimes(12);
        expect(lazy.evaluate(1, 0, "a", { getCellValue }).matchedRuleIds).toEqual(["dupe"]);
        expect(getCellValue).toHaveBeenCalledTimes(12);

        const valuesB = new Map([["0:0", "a"], ["1:0", "b"], ["2:0", "c"], ["0:1", "d"], ["1:1", "e"], ["2:1", "f"]]);
        const getCellValueB = vi.fn((col: number, row: number) => valuesB.get(`${col}:${row}`) ?? null);
        expect(lazy.evaluate(0, 0, "a", { getCellValue: getCellValueB }).matchedRuleIds).toEqual(["unique"]);
        expect(getCellValueB).toHaveBeenCalledTimes(6);
        expect(lazy.evaluate(1, 0, "b", { getCellValue: getCellValueB }).matchedRuleIds).toEqual(["unique"]);
        expect(getCellValueB).toHaveBeenCalledTimes(6);
    });

    test("merges overlapping styles by order and honors stopIfTrue", () => {
        const rules: ConditionalFormatRule[] = [
            { id: "first", kind: "number-compare", range, style: { bg: "red", text: "dark" }, operator: "gte", value: 1, stopIfTrue: true },
            { id: "second", kind: "number-compare", range, style: { bg: "blue", font: "bold" }, operator: "gt", value: 2 },
        ];
        expect(evaluateConditionalFormatting(0, 0, 3, rules)).toEqual({ style: { bg: "red", text: "dark" }, matchedRuleIds: ["first"], diagnostics: [] });
        expect(evaluateConditionalFormatting(0, 0, 0, rules)).toEqual({ style: {}, matchedRuleIds: [], diagnostics: [] });
    });

    test("matches formula errors by optional code", () => {
        const rules: ConditionalFormatRule[] = [{ id: "errors", kind: "formula-error", range, code: "#REF!", style: { text: "red", numberFormat: "error" } }];
        expect(evaluateConditionalFormatting(0, 0, { kind: "error", code: "#REF!" }, rules).matchedRuleIds).toEqual(["errors"]);
        expect(evaluateConditionalFormatting(0, 0, { kind: "error", code: "#VALUE!" }, rules).matchedRuleIds).toEqual([]);
    });

    test("turns custom exceptions into diagnostics without throwing", () => {
        const result = evaluateConditionalFormatting(0, 0, 1, [{ id: "custom", kind: "custom", range, style: { bg: "red" }, predicate: () => { throw new Error("bad"); } }]);
        expect(result).toEqual({ style: {}, matchedRuleIds: [], diagnostics: [{ ruleId: "custom", code: "custom-error", message: "Custom conditional-format predicate threw an exception" }] });
    });

    test("validates ranges/rules and does not mutate input rules", () => {
        const rule: ConditionalFormatRule = { id: "one", kind: "number-compare", range, style: { bg: "red" }, operator: "eq", value: 1 };
        const before = JSON.stringify(rule);
        const engine = compileConditionalFormatting([rule]);
        expect(JSON.stringify(rule)).toBe(before);
        expect(engine.rules).not.toBe(([] as unknown[]));
        expect(() => compileConditionalFormatting([{ ...rule, range: { x: -1, y: 0, width: 1, height: 1 } }])).toThrowError(ConditionalFormatError);
        expect(() => compileConditionalFormatting([{ ...rule, kind: "number-between", min: 5, max: 1 } as ConditionalFormatRule])).toThrowError(ConditionalFormatError);
        expect(() => compileConditionalFormatting([{ ...rule, id: "one" }, { ...rule, id: "one" }])).toThrowError(ConditionalFormatError);
    });
});
