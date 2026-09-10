import { describe, expect, test } from "vitest";
import { compileFormula, columnIndexToName, columnNameToIndex, translateFormula } from "../src/formula.js";
import { FormulaFunctionRegistry } from "../src/function-registry.js";

const values = new Map<string, string | number | boolean | null>([
    ["0:0", 2],
    ["0:1", 3],
    ["1:0", 10],
    ["1:1", 20],
]);

const context = {
    currentRow: 1,
    rowCount: 2,
    resolveColumn: (name: string) => ({ quantity: 0, price: 1 }[name.toLocaleLowerCase()]),
    getCellValue: (col: number, row: number) => values.get(`${col}:${row}`) ?? null,
};

describe("formula compiler", () => {
    test("converts spreadsheet column names", () => {
        expect(columnNameToIndex("A")).toBe(0);
        expect(columnNameToIndex("AA")).toBe(26);
        expect(columnIndexToName(701)).toBe("ZZ");
    });

    test("evaluates A1 references, ranges and functions", () => {
        expect(compileFormula("=A1*B2").evaluate(context)).toBe(40);
        expect(compileFormula("=SUM(A1:B2)").evaluate(context)).toBe(35);
        expect(compileFormula("=IF(A2>2, \"yes\", \"no\")").evaluate(context)).toBe("yes");
    });

    test("covers every Phase 0 built-in function", () => {
        expect(compileFormula("=SUM(A1:B2)").evaluate(context)).toBe(35);
        expect(compileFormula("=AVERAGE(A1:B2)").evaluate(context)).toBe(8.75);
        expect(compileFormula("=MIN(A1:B2)").evaluate(context)).toBe(2);
        expect(compileFormula("=MAX(A1:B2)").evaluate(context)).toBe(20);
        expect(compileFormula("=COUNT(A1:B2, \"ignored\")").evaluate(context)).toBe(4);
        expect(compileFormula("=COUNTA(A1:B2, \"present\", \"\")").evaluate(context)).toBe(5);
        expect(compileFormula("=IF(AND(A1=2, B2=20), \"yes\", \"no\")").evaluate(context)).toBe("yes");
        expect(compileFormula("=OR(FALSE, A2=3)").evaluate(context)).toBe(true);
        expect(compileFormula("=NOT(A1=3)").evaluate(context)).toBe(true);
        expect(compileFormula("=ABS(-2.5)").evaluate(context)).toBe(2.5);
        expect(compileFormula("=ROUND(2.345, 2)").evaluate(context)).toBe(2.35);
    });

    test("evaluates text functions and escaped quotes", () => {
        expect(compileFormula('=CONCAT("Hello", " ", "world")').evaluate(context)).toBe("Hello world");
        expect(compileFormula('=CONCAT("A", A1:B1)').evaluate(context)).toBe("A210");
        expect(compileFormula('=LEN("hello")').evaluate(context)).toBe(5);
        expect(compileFormula('=LOWER("HeLLo")').evaluate(context)).toBe("hello");
        expect(compileFormula('=UPPER("HeLLo")').evaluate(context)).toBe("HELLO");
        expect(compileFormula('=TRIM("  too   many   spaces  ")').evaluate(context)).toBe("too many spaces");
        expect(compileFormula('=LEFT("abcdef")').evaluate(context)).toBe("a");
        expect(compileFormula('=LEFT("abcdef", 3)').evaluate(context)).toBe("abc");
        expect(compileFormula('=RIGHT("abcdef", 2)').evaluate(context)).toBe("ef");
        expect(compileFormula('=MID("abcdef", 2, 3)').evaluate(context)).toBe("bcd");
        expect(compileFormula('="say ""hi"""').evaluate(context)).toBe('say "hi"');
    });

    test("uses numeric date serials for date functions", () => {
        // Serial 0 is 1899-12-30, matching the representation documented in formula.ts.
        expect(compileFormula("=DATE(2020, 1, 15)").evaluate(context)).toBe(43845);
        expect(compileFormula("=YEAR(DATE(2020, 1, 15))").evaluate(context)).toBe(2020);
        expect(compileFormula("=MONTH(DATE(2020, 1, 15))").evaluate(context)).toBe(1);
        expect(compileFormula("=DAY(DATE(2020, 1, 15))").evaluate(context)).toBe(15);

        const today = compileFormula("=TODAY()").evaluate(context);
        expect(typeof today).toBe("number");
        expect(today).toBeGreaterThan(45_000);
        expect(today).toBeLessThan(60_000);
        expect(Number.isInteger(today)).toBe(true);
    });

    test("evaluates current-row and whole-column named references", () => {
        expect(compileFormula("=[@Quantity]*[@Price]").evaluate(context)).toBe(60);
        expect(compileFormula("=SUM([Quantity])").evaluate(context)).toBe(5);
    });

    test("evaluates explicit caption-plus-row references and caption ranges", () => {
        const localizedContext = {
            ...context,
            resolveColumn: (name: string) => ({ "số lượng": 0, "đơn giá": 1 }[name.toLocaleLowerCase()]),
        };
        expect(compileFormula("=[Số lượng]1*[Đơn giá]2").evaluate(localizedContext)).toBe(40);
        expect(compileFormula("=SUM([Số lượng]1:[Số lượng]2)").evaluate(localizedContext)).toBe(5);
        expect(compileFormula("=[@Số lượng]+[Số lượng]1").evaluate({ ...localizedContext, currentRow: 1 })).toBe(5);
    });

    test("tracks exact dependencies for caption-plus-row references", () => {
        const compiled = compileFormula("=[Số lượng]2+[Đơn giá]1");
        expect([...compiled.dependencies({ ...context, resolveColumn: (name: string) => ({ "số lượng": 0, "đơn giá": 1 }[name.toLocaleLowerCase()]) }).cells].sort()).toEqual(["0:1", "1:0"]);
    });

    test("translates only relative parts of A1 references", () => {
        expect(translateFormula("=A1+$B1+C$2+$D$4", 2, 3)).toBe("=C4+$B4+E$2+$D$4");
    });

    test("does not translate A1-looking text or structured names", () => {
        expect(translateFormula('="A1"&A1&"A1""B2"&$B1', 1, 1)).toBe('="A1"&B2&"A1""B2"&$B2');
        expect(translateFormula("=[A1]&A1", 1, 1)).toBe("=[A1]&B2");
        expect(translateFormula("=[Số lượng]12&[@Số lượng]", 1, 1)).toBe("=[Số lượng]12&[@Số lượng]");
        expect(translateFormula("=LOG10(A1)", 1, 1)).toBe("=LOG10(B2)");
    });

    test("follows arithmetic precedence and right-associative powers", () => {
        expect(compileFormula("=1+2*3^2").evaluate(context)).toBe(19);
        expect(compileFormula("=-2^2").evaluate(context)).toBe(-4);
        expect(compileFormula("=2^3^2").evaluate(context)).toBe(512);
    });

    test("supports absolute references", () => {
        expect(compileFormula("=$A$1+$B2+A$1").evaluate(context)).toBe(24);
    });

    test("reports malformed formulas and evaluation errors", () => {
        expect(() => compileFormula("=SUM(1,")).toThrow();
        expect(() => compileFormula('="unterminated')).toThrow();
        expect(() => compileFormula("=A0")).toThrow();
        expect(compileFormula("=NOT_A_FUNCTION(1)").evaluate(context)).toMatchObject({ code: "#NAME?" });
        expect(compileFormula("=1/0").evaluate(context)).toMatchObject({ code: "#DIV/0!" });
        expect(compileFormula("=#REF!+1").evaluate(context)).toMatchObject({ code: "#REF!" });
    });

    test("evaluates case-insensitive custom scalar and range functions", () => {
        const registry = new FormulaFunctionRegistry();
        registry.register("double", args => (args[0] as number) * 2);
        registry.register("range_total", args => {
            const range = args[0];
            if (range === null || typeof range !== "object" || !("values" in range)) return 0;
            const values = (range as { readonly values: readonly unknown[] }).values;
            return values.reduce((sum: number, value) => sum + (typeof value === "number" ? value : 0), 0);
        });
        expect(compileFormula("=DOUBLE(A1)", { functionRegistry: registry }).evaluate(context)).toBe(4);
        expect(compileFormula("=range_total(A1:B2)", { functionRegistry: registry }).evaluate(context)).toBe(35);
        expect(registry.list()).toEqual(["DOUBLE", "RANGE_TOTAL"]);
    });

    test("rejects invalid names, duplicate functions, and built-in collisions", () => {
        const registry = new FormulaFunctionRegistry();
        expect(() => registry.register("A1", () => 1)).toThrow();
        expect(() => registry.register("[Column]", () => 1)).toThrow();
        registry.register("custom", () => 1);
        expect(() => registry.register("CUSTOM", () => 2)).toThrow(/already registered/);
        expect(() => registry.register("SUM", () => 1)).toThrow(/built in/);
        registry.register("SUM", () => 99, { overrideBuiltIn: true });
        expect(compileFormula("=sum(1, 2)", { functionRegistry: registry }).evaluate(context)).toBe(99);
        expect(registry.unregister("CuStOm")).toBe(true);
        expect(registry.has("custom")).toBe(false);
    });

    test("converts custom failures and asynchronous/invalid values to VALUE errors", () => {
        const registry = new FormulaFunctionRegistry();
        registry.register("throws", () => {
            throw new Error("boom");
        });
        registry.register("async_value", () => Promise.resolve(1) as unknown as number);
        registry.register("rejected_value", () => Promise.reject(new Error("async boom")) as unknown as number);
        registry.register("bad_value", () => ({ nope: true }) as unknown as number);
        expect(compileFormula("=THROWS()", { functionRegistry: registry }).evaluate(context)).toEqual({
            kind: "error",
            code: "#VALUE!",
            message: "Custom function 'THROWS' failed: boom",
        });
        expect(compileFormula("=ASYNC_VALUE()", { functionRegistry: registry }).evaluate(context)).toMatchObject({
            code: "#VALUE!",
            message: "Custom function 'ASYNC_VALUE' returned a Promise",
        });
        expect(compileFormula("=REJECTED_VALUE()", { functionRegistry: registry }).evaluate(context)).toMatchObject({
            code: "#VALUE!",
            message: "Custom function 'REJECTED_VALUE' returned a Promise",
        });
        expect(compileFormula("=BAD_VALUE()", { functionRegistry: registry }).evaluate(context)).toMatchObject({
            code: "#VALUE!",
            message: "Custom function 'BAD_VALUE' returned an invalid value",
        });
    });

    test("keeps dependencies from custom function arguments and supports unregister", () => {
        const registry = new FormulaFunctionRegistry();
        registry.register("identity", args => args[0] as number);
        const compiled = compileFormula("=IDENTITY(A1:B2)", { functionRegistry: registry });
        expect([...compiled.dependencies(context).cells].sort()).toEqual(["0:0", "0:1", "1:0", "1:1"]);
        registry.unregister("identity");
        expect(compiled.evaluate(context)).toMatchObject({ code: "#NAME?" });
    });

    test("evaluates INDEX, MATCH, and XLOOKUP with horizontal and vertical ranges", () => {
        expect(compileFormula("=INDEX(A1:B2, 2, 1)").evaluate(context)).toBe(3);
        expect(compileFormula("=INDEX(A1:B2, 1, 2)").evaluate(context)).toBe(10);
        expect(compileFormula("=INDEX(A1:A2, 2)").evaluate(context)).toBe(3);
        expect(compileFormula("=MATCH(3, A1:A2)").evaluate(context)).toBe(2);
        expect(compileFormula("=MATCH(20, B1:B2, 0)").evaluate(context)).toBe(2);
        expect(compileFormula('=XLOOKUP(3, A1:A2, B1:B2)').evaluate(context)).toBe(20);
        expect(compileFormula('=XLOOKUP(20, B1:B2, A1:A2, "missing")').evaluate(context)).toBe(3);
    });

    test("preserves empty cells returned by INDEX and XLOOKUP", () => {
        const emptyContext = {
            ...context,
            getCellValue: (col: number, row: number) => (col === 1 && row === 0 ? null : values.get(`${col}:${row}`) ?? null),
        };
        expect(compileFormula("=INDEX(B1:B2, 1)").evaluate(emptyContext)).toBeNull();
        expect(compileFormula("=XLOOKUP(2, A1:A2, B1:B2)").evaluate(emptyContext)).toBeNull();
    });

    test("reports lookup misses, unsupported modes, and invalid dimensions", () => {
        expect(compileFormula("=MATCH(999, A1:A2)").evaluate(context)).toMatchObject({ code: "#N/A" });
        expect(compileFormula("=MATCH(2, A1:A2, 1)").evaluate(context)).toMatchObject({ code: "#N/A" });
        expect(compileFormula("=MATCH(2, A1:B2)").evaluate(context)).toMatchObject({ code: "#VALUE!" });
        expect(compileFormula("=INDEX(A1:B2, 3, 1)").evaluate(context)).toMatchObject({ code: "#REF!" });
        expect(compileFormula("=XLOOKUP(2, A1:A2, B1:B1)").evaluate(context)).toMatchObject({ code: "#VALUE!" });
        expect(compileFormula('=XLOOKUP(999, A1:A2, B1:B2, "missing")').evaluate(context)).toBe("missing");
        expect(compileFormula("=XLOOKUP(999, A1:A2, B1:B2)").evaluate(context)).toMatchObject({ code: "#N/A" });
    });

    test("passes range shape metadata to custom functions", () => {
        const registry = new FormulaFunctionRegistry();
        registry.register("shape", args => {
            const range = args[0];
            return range !== null && typeof range === "object" && "rows" in range && "columns" in range
                ? `${range.rows}x${range.columns}`
                : "scalar";
        });
        expect(compileFormula("=SHAPE(A1:B2)", { functionRegistry: registry }).evaluate(context)).toBe("2x2");
        expect(compileFormula("=SHAPE(A1:A2)", { functionRegistry: registry }).evaluate(context)).toBe("2x1");
    });
});
