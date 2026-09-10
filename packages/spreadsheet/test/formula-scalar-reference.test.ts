import { describe, expect, test } from "vitest";
import { PagedDataSource } from "../src/paged-data-source.js";
import { PagedFormulaAdapter } from "../src/paged-formula-adapter.js";
import { SpreadsheetModel } from "../src/model.js";
import type { FormulaValue } from "../src/formula.js";

const columns = [
    { id: "text", title: "Text", type: "text" as const },
    { id: "number", title: "Number", type: "number" as const },
    { id: "enabled", title: "Enabled", type: "boolean" as const },
    { id: "empty", title: "Empty", type: "text" as const },
    { id: "url", title: "Web URL", type: "text" as const },
    { id: "date", title: "Date value", type: "text" as const },
    { id: "a1Result", title: "A1 result", type: "text" as const },
    { id: "captionResult", title: "Caption result", type: "text" as const },
];

const sourceValues: readonly [string, number, boolean, null, string, string] = [
    "00123",
    42,
    true,
    null,
    "https://example.test/orders/42",
    "2024-01-15",
];

const referenceCases: readonly {
    readonly sourceColumn: string;
    readonly caption: string;
    readonly expected: FormulaValue;
}[] = [
    { sourceColumn: "A", caption: "Text", expected: "00123" },
    { sourceColumn: "B", caption: "Number", expected: 42 },
    { sourceColumn: "C", caption: "Enabled", expected: true },
    { sourceColumn: "D", caption: "Empty", expected: null },
    { sourceColumn: "E", caption: "Web URL", expected: "https://example.test/orders/42" },
    { sourceColumn: "F", caption: "Date value", expected: "2024-01-15" },
];

function modelRows(): readonly (readonly FormulaValue[])[] {
    const rows: FormulaValue[][] = [
        [...sourceValues, null, null],
        ...referenceCases.slice(1).map(() => [null, null, null, null, null, null, null, null]),
    ];
    referenceCases.forEach(({ sourceColumn, caption }, row) => {
        rows[row]![6] = `=${sourceColumn}1`;
        rows[row]![7] = `=[${caption}]1`;
    });
    return rows;
}

describe("formula scalar references", () => {
    test("SpreadsheetModel preserves the source type for A1 and explicit caption-row references", () => {
        const model = new SpreadsheetModel(columns, referenceCases.length, modelRows());

        referenceCases.forEach(({ expected }, row) => {
            expect(model.getValue(6, row)).toStrictEqual(expected);
            expect(model.getValue(7, row)).toStrictEqual(expected);
        });

        // A numeric-looking text value must remain text, rather than being
        // coerced to a number while resolving a direct reference.
        expect(model.getValue(6, 0)).toBe("00123");
        expect(typeof model.getValue(6, 0)).toBe("string");
        expect(typeof model.getValue(6, 1)).toBe("number");
        expect(typeof model.getValue(6, 2)).toBe("boolean");
        expect(model.getValue(6, 3)).toBeNull();
        expect(model.getDisplayValue(6, 4)).toBe("https://example.test/orders/42");
        expect(model.getDisplayValue(6, 5)).toBe("2024-01-15");
    });

    test("PagedFormulaAdapter preserves scalar types without loading a page", () => {
        const source = new PagedDataSource({
            rowCount: referenceCases.length,
            columnCount: columns.length,
            pageSize: 2,
            generateRow: row => row === 0
                ? [...sourceValues, null, null]
                : [null, null, null, null, null, null, null, null],
        });
        const adapter = new PagedFormulaAdapter(source, columns);

        referenceCases.forEach(({ sourceColumn, caption, expected }, row) => {
            adapter.setFormula(6, row, `=${sourceColumn}1`);
            expect(adapter.getValue(6, row)).toStrictEqual(expected);

            adapter.setFormula(7, row, `=[${caption}]1`);
            expect(adapter.getValue(7, row)).toStrictEqual(expected);
        });

        // Reading direct references uses the source reader and does not turn
        // text/boolean/null values into numbers or require a resident page.
        expect(adapter.getValue(6, 0)).toBe("00123");
        expect(typeof adapter.getValue(6, 0)).toBe("string");
        expect(typeof adapter.getValue(6, 1)).toBe("number");
        expect(typeof adapter.getValue(6, 2)).toBe("boolean");
        expect(adapter.getValue(6, 3)).toBeNull();
        expect(adapter.getDisplayValue(6, 4)).toBe("https://example.test/orders/42");
        expect(adapter.getDisplayValue(6, 5)).toBe("2024-01-15");
        expect(source.getStats().cachedPages).toBe(0);
    });
});
