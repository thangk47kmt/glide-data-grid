import { describe, expect, test } from "vitest";
import { PagedDataSource } from "../src/paged-data-source.js";
import { PagedFormulaAdapter } from "../src/paged-formula-adapter.js";
import { PagedDataQuery } from "../src/paged-data-query.js";

const columns = [
    { id: "quantity", title: "Quantity", type: "number" as const },
    { id: "price", title: "Price", type: "number" as const },
    { id: "total", title: "Total", type: "number" as const },
];

function setup() {
    const source = new PagedDataSource({
        rowCount: 500_000,
        columnCount: 3,
        pageSize: 100,
        generateCell: (row, col) => col === 0 ? row + 1 : col === 1 ? 10 : null,
    });
    return { source, formulas: new PagedFormulaAdapter(source, columns) };
}

describe("PagedFormulaAdapter", () => {
    test("copies scalar references without coercing text, booleans, dates, URIs, or empty cells", () => {
        const source = new PagedDataSource({
            rowCount: 2,
            columnCount: 7,
            generateCell: (_row, col) => ["Invoice", 42, true, "2026-09-01", "https://example.com/1", "Result", null][col] ?? null,
        });
        const columns = [
            { id: "text", title: "Text", type: "text" as const },
            { id: "number", title: "Number", type: "number" as const },
            { id: "boolean", title: "Boolean", type: "boolean" as const },
            { id: "date", title: "Date", type: "text" as const },
            { id: "uri", title: "URI", type: "text" as const },
            { id: "empty", title: "Empty", type: "text" as const },
            { id: "blank", title: "Blank", type: "text" as const },
        ];
        const formulas = new PagedFormulaAdapter(source, columns);

        formulas.setFormula(5, 0, "=A1");
        expect(formulas.getValue(5, 0)).toBe("Invoice");
        formulas.setFormula(5, 1, "=[Boolean]1");
        expect(formulas.getValue(5, 1)).toBe(true);
        formulas.setFormula(5, 0, "=D1");
        expect(formulas.getValue(5, 0)).toBe("2026-09-01");
        formulas.setFormula(5, 1, "=[URI]1");
        expect(formulas.getValue(5, 1)).toBe("https://example.com/1");
        formulas.setFormula(5, 1, "=G1");
        expect(formulas.getValue(5, 1)).toBeNull();
    });

    test("evaluates A1 references without loading or materializing source pages", () => {
        const { source, formulas } = setup();
        formulas.setFormula(2, 499_999, "=A1*B1");
        expect(formulas.getValue(2, 499_999)).toBe(10);
        expect(source.getStats()).toMatchObject({ cachedPages: 0, generatedCells: 0 });
        expect(formulas.getStats().formulaCells).toBe(1);
    });

    test("supports current-row structured column names and invalidates dependencies", () => {
        const { formulas } = setup();
        formulas.setFormula(2, 4, "=[@Quantity]*[@Price]");
        expect(formulas.getValue(2, 4)).toBe(50);
        formulas.setCell(0, 4, 7);
        expect(formulas.getValue(2, 4)).toBe(70);
    });

    test("resolves persisted ids and user-facing Vietnamese captions", () => {
        const source = new PagedDataSource({ rowCount: 5, columnCount: 3, pageSize: 5, generateCell: (row, col) => col === 0 ? row + 1 : col === 1 ? 10 : null });
        const localized = new PagedFormulaAdapter(source, [
            { id: "so_luong", title: "Số lượng", type: "number" },
            { id: "don_gia", title: "Đơn giá", type: "number" },
            { id: "thanh_tien", title: "Thành tiền", type: "number" },
        ]);
        localized.setFormula(2, 1, "=[@Số lượng]*[@Đơn giá]");
        expect(localized.resolveColumn("don_gia")).toBe(1);
        expect(localized.getValue(2, 1)).toBe(20);
        localized.setCell(0, 1, 4);
        expect(localized.getValue(2, 1)).toBe(40);
    });

    test("resolves caption-plus-row references to exact rows and invalidates only those cells", () => {
        const source = new PagedDataSource({
            rowCount: 4,
            columnCount: 3,
            pageSize: 4,
            generateCell: (row, col) => col === 0 ? row + 1 : col === 1 ? (row + 1) * 10 : null,
        });
        const localized = new PagedFormulaAdapter(source, [
            { id: "qty_id", title: "Số lượng", type: "number" },
            { id: "price_id", title: "Đơn giá", type: "number" },
            { id: "total_id", title: "Thành tiền", type: "number" },
        ]);
        localized.setFormula(2, 0, "=[Số lượng]3+[Đơn giá]2");
        expect(localized.getValue(2, 0)).toBe(23);
        localized.setCell(0, 2, 20);
        expect(localized.getValue(2, 0)).toBe(40);
        localized.setCell(0, 1, 99);
        expect(localized.getValue(2, 0)).toBe(40);
    });

    test("supports caption-plus-row ranges with spaces and Unicode", () => {
        const source = new PagedDataSource({ rowCount: 3, columnCount: 2, generateCell: (row, col) => col === 0 ? row + 1 : null });
        const localized = new PagedFormulaAdapter(source, [
            { id: "so_luong_id", title: "Số lượng", type: "number" },
            { id: "thanh_tien_id", title: "Thành tiền", type: "number" },
        ]);
        localized.setFormula(1, 0, "=SUM([Số lượng]1:[Số lượng]3)");
        expect(localized.getValue(1, 0)).toBe(6);
    });

    test("does not silently choose an ambiguous caption", () => {
        const source = new PagedDataSource({ rowCount: 1, columnCount: 3, generateCell: (_row, col) => col < 2 ? col + 2 : null });
        const localized = new PagedFormulaAdapter(source, [
            { id: "left", title: "Giá", type: "number" },
            { id: "right", title: "Giá", type: "number" },
            { id: "total", title: "Tổng", type: "number" },
        ]);
        expect(localized.resolveColumn("Giá")).toBeUndefined();
        localized.setFormula(2, 0, "=[@Giá]");
        expect(localized.getDisplayValue(2, 0)).toBe("#REF!");
    });

    test("reports cycles and keeps formula storage sparse", () => {
        const { formulas } = setup();
        formulas.setFormula(0, 0, "=C1");
        formulas.setFormula(2, 0, "=A1");
        expect(formulas.getDisplayValue(0, 0)).toBe("#CYCLE!");
        expect(formulas.getStats().formulaCells).toBe(2);
    });

    test("clearFormula reveals the literal value underneath and maintains the row index", () => {
        const { formulas } = setup();
        formulas.setCell(0, 0, 7);
        formulas.setFormula(0, 0, "=1+1");
        expect(formulas.getFormulaColumns(0)).toEqual([0]);
        formulas.clearFormula(0, 0);
        expect(formulas.getValue(0, 0)).toBe(7);
        expect(formulas.getFormulaColumns(0)).toEqual([]);
    });

    test("out-of-bounds coordinates evaluate as REF errors", () => {
        const { formulas } = setup();
        formulas.setFormula(2, 0, "=ZZ1+A999999");
        expect(formulas.getDisplayValue(2, 0)).toBe("#REF!");
    });

    test("validates formula types and leaves rejected drafts uncommitted", () => {
        const source = new PagedDataSource({
            rowCount: 2,
            columnCount: 4,
            generateCell: (row, col) => col === 0 ? `Record ${row + 1}` : col === 1 ? row + 1 : col === 2 ? row % 2 === 0 : null,
        });
        const typedColumns = [
            { id: "text", title: "Text", type: "text" as const },
            { id: "number", title: "Number", type: "number" as const },
            { id: "boolean", title: "Boolean", type: "boolean" as const },
            { id: "result", title: "Result", type: "number" as const },
        ];
        const formulas = new PagedFormulaAdapter(source, typedColumns);

        expect(formulas.validateFormula(3, 0, "=B1")).toMatchObject({ valid: true, value: 1, actualType: "number" });
        expect(formulas.validateFormula(3, 0, "=A1")).toMatchObject({ valid: false, error: { code: "#VALUE!" } });
        expect(formulas.validateFormula(3, 0, "=TRUE")).toMatchObject({ valid: false, error: { code: "#VALUE!" } });
        expect(formulas.validateFormula(3, 0, "=B1/0")).toMatchObject({ valid: false, error: { code: "#DIV/0!" } });
        expect(formulas.validateFormula(3, 0, "=Unknown1")).toMatchObject({ valid: false, error: { code: "#REF!" } });
        expect(formulas.validateFormula(3, 0, "=B999")).toMatchObject({ valid: false, error: { code: "#REF!" } });
        expect(formulas.validateFormula(3, 0, "=B1+")).toMatchObject({ valid: false, error: { code: "#VALUE!" } });

        // Validation is a dry run: no formula or value is changed until the
        // editor explicitly calls setCell after validation succeeds.
        expect(formulas.hasFormula(3, 0)).toBe(false);
        expect(formulas.getValue(3, 0)).toBeNull();
        formulas.setFormula(3, 0, "=B1");
        expect(formulas.getValue(3, 0)).toBe(1);
        expect(formulas.validateFormula(1, 1, "=D2")).toMatchObject({ valid: true, value: null });
    });

    test("reports a cycle while validating a replacement formula", () => {
        const { formulas } = setup();
        formulas.setFormula(0, 0, "=C1");
        formulas.setFormula(2, 0, "=A1");
        expect(formulas.validateFormula(0, 0, "=C1")).toMatchObject({ valid: false, error: { code: "#CYCLE!" } });
    });

    test("feeds computed formula values into search, filter and sort", async () => {
        const source = new PagedDataSource({ rowCount: 3, columnCount: 3, generateCell: (row, col) => col === 0 ? row + 1 : col === 1 ? 10 : null });
        const formulas = new PagedFormulaAdapter(source, columns);
        formulas.setFormula(2, 0, "=[@Quantity]*[@Price]");
        formulas.setFormula(2, 1, "=[@Quantity]*[@Price]");
        const read = (row: number, col: number) => formulas.hasFormula(col, row) ? formulas.getValue(col, row) as number : source.readCell(row, col);
        const query = new PagedDataQuery({
            rowCount: 3,
            columnCount: 3,
            getCell: read,
            readCell: read,
            readSearchText: row => [read(row, 0), read(row, 1), read(row, 2)].join("\u0000"),
        }, columns);
        expect((await query.query({ search: "20", limit: 10 })).rows).toEqual([1]);
        expect((await query.query({ filters: [{ column: "total", operator: "gte", value: 10 }], sorts: [{ column: "total", direction: "desc" }], limit: 10 })).rows).toEqual([1, 0]);
    });
});
