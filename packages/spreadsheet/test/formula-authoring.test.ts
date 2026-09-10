import { describe, expect, test } from "vitest";
import { FormulaFunctionRegistry } from "../src/function-registry.js";
import {
    BUILT_IN_FUNCTION_CATALOG,
    diagnoseFormula,
    getFormulaCompletions,
} from "../src/formula-authoring.js";

describe("formula authoring", () => {
    test("diagnoses missing equals, parser errors and invalid references without throwing", () => {
        expect(diagnoseFormula("SUM(").some(item => item.code === "MISSING_EQUALS")).toBe(true);
        expect(diagnoseFormula('="unterminated').some(item => item.code === "PARSER_ERROR")).toBe(true);
        expect(diagnoseFormula("=A0 + #REF!").some(item => item.code === "INVALID_REFERENCE")).toBe(true);
        expect(() => diagnoseFormula(null as unknown as string)).not.toThrow();
        expect(diagnoseFormula(null as unknown as string)[0]?.code).toBe("INVALID_SOURCE");
    });

    test("finds unknown functions and structured columns while accepting known custom functions", () => {
        const registry = new FormulaFunctionRegistry();
        registry.register("discount", () => 0);
        const context = {
            functionRegistry: registry,
            resolveColumn: (name: string) => name.toLowerCase() === "quantity" ? 0 : undefined,
        };
        expect(diagnoseFormula("=MYSTERY(1)", context).some(item => item.code === "UNKNOWN_FUNCTION")).toBe(true);
        expect(diagnoseFormula("=DISCOUNT(1)", context)).toEqual([]);
        expect(diagnoseFormula("=[@Missing]", context).some(item => item.code === "UNKNOWN_COLUMN")).toBe(true);
        expect(diagnoseFormula("=[@Quantity]", context)).toEqual([]);
    });

    test("exports metadata for implemented built-ins", () => {
        const names = BUILT_IN_FUNCTION_CATALOG.map(item => item.name);
        expect(names).toContain("SUM");
        expect(names).toContain("XLOOKUP");
        expect(new Set(names).size).toBe(names.length);
        expect(BUILT_IN_FUNCTION_CATALOG.every(item => item.signature.length > 0 && item.description.length > 0)).toBe(true);
    });

    test("ranks and deduplicates function, custom, cell and structured completions", () => {
        const registry = new FormulaFunctionRegistry();
        registry.register("customTotal", () => 0);
        registry.register("SUM", () => 0, { overrideBuiltIn: true });
        const functions = getFormulaCompletions("=su", 3, { functionRegistry: registry });
        expect(functions[0]?.label).toBe("SUM");
        expect(functions.filter(item => item.label.toUpperCase() === "SUM")).toHaveLength(1);
        expect(getFormulaCompletions("=custom", 7, { functionRegistry: registry }).some(item => item.label === "CUSTOMTOTAL")).toBe(true);
        expect(getFormulaCompletions("=A", 2).some(item => item.label === "A1")).toBe(true);
        const columns = getFormulaCompletions("=[@Qu", 5, { structuredColumns: ["Price", "Quantity"] });
        expect(columns.map(item => item.label)).toEqual(["Quantity"]);
        expect(columns[0]?.insertText).toBe("[@Quantity]");
    });

    test("provides replacement ranges that produce valid authoring text", () => {
        const apply = (source: string, completion: ReturnType<typeof getFormulaCompletions>[number]): string =>
            source.slice(0, completion.replaceStart) + completion.insertText + source.slice(completion.replaceEnd);
        const rowColumn = getFormulaCompletions("=[@Qu", 5, { structuredColumns: ["Quantity"] })[0];
        const wholeColumn = getFormulaCompletions("=[Pri", 5, { structuredColumns: ["Price"] })[0];
        const functionCompletion = getFormulaCompletions("=SU", 3)[0];
        expect(rowColumn).toBeDefined();
        expect(wholeColumn).toBeDefined();
        expect(functionCompletion).toBeDefined();
        expect(apply("=[@Qu", rowColumn!)).toBe("=[@Quantity]");
        expect(apply("=[Pri", wholeColumn!)).toBe("=[Price]");
        expect(apply("=SU", functionCompletion!)).toBe("=SUM(");
    });

    test("completion remains safe for malformed source and hostile registry", () => {
        const hostile = { list: () => { throw new Error("broken"); } } as unknown as FormulaFunctionRegistry;
        expect(() => getFormulaCompletions("=", 999, { functionRegistry: hostile })).not.toThrow();
        expect(getFormulaCompletions("=", 999, { functionRegistry: hostile }).length).toBeGreaterThan(0);
    });

    test("completes captions containing spaces and Vietnamese characters", () => {
        const resolveColumn = (name: string) => name.normalize("NFKC").toLocaleLowerCase("en-US") === "đơn giá" ? 1 : undefined;
        const completions = getFormulaCompletions("=[@Đơn g", 8, {
            resolveColumn,
            structuredColumns: ["Số lượng", "Đơn giá"],
        });
        expect(completions.map(item => item.label)).toEqual(["Đơn giá"]);
        expect(completions[0]?.insertText).toBe("[@Đơn giá]");
        expect(completions[0]?.replaceStart).toBe(1);
        expect(completions[0]?.replaceEnd).toBe(8);
    });

    test("does not offer an ambiguous caption", () => {
        const completions = getFormulaCompletions("=[@Gi", 5, {
            resolveColumn: () => undefined,
            structuredColumns: ["Giá", "Giá"],
        });
        expect(completions).toEqual([]);
    });

    test("normalizes compatibility-equivalent caption text while completing", () => {
        const fullWidth = "Ｄｏｎ";
        expect(getFormulaCompletions(`=[@${fullWidth}`, 3 + fullWidth.length, {
            structuredColumns: ["Don gia"],
        })[0]?.label).toBe("Don gia");
    });
});
