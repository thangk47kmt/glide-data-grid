import { describe, expect, test } from "vitest";
import { clearFormatCache, formatCellValue, getFormatCacheStats } from "../src/format.js";

describe("spreadsheet value formatting", () => {
    test("formats general, text, null, booleans and formula errors deterministically", () => {
        expect(formatCellValue(null, { kind: "general" })).toBe("");
        expect(formatCellValue(true, { kind: "general" })).toBe("TRUE");
        expect(formatCellValue(false, { kind: "text" })).toBe("FALSE");
        expect(formatCellValue(1234.5, { kind: "general" })).toBe("1234.5");
        expect(formatCellValue({ kind: "error", code: "#DIV/0!" }, { kind: "currency", currency: "USD" })).toBe("#DIV/0!");
    });

    test("formats numbers, percentages and currency for supported locales", () => {
        expect(formatCellValue(1234.5, { kind: "number", maximumFractionDigits: 2 }, "en-US")).toBe("1,234.5");
        expect(formatCellValue(1234.5, { kind: "number", maximumFractionDigits: 2 }, "vi-VN")).toBe("1.234,5");
        expect(formatCellValue(0.125, { kind: "percent", maximumFractionDigits: 1 }, "en-US")).toBe("12.5%");
        expect(formatCellValue(1234.5, { kind: "currency", currency: "USD" }, "en-US")).toBe("$1,234.50");
        expect(formatCellValue(1234.5, { kind: "currency", currency: "VND" }, "vi-VN")).toBe("1.235 ₫");
    });

    test("formats formula date serials with locale presets", () => {
        // Numeric date serial 43845 is 2020-01-15 (epoch 1899-12-30).
        expect(formatCellValue(43845, { kind: "date", preset: "short" }, "en-US")).toBe("1/15/20");
        expect(formatCellValue(43845, { kind: "date", preset: "medium" }, "vi-VN")).toBe("15 thg 1, 2020");
        expect(formatCellValue(43845, { kind: "date", preset: "long" }, "en-US")).toBe("January 15, 2020");
        expect(formatCellValue("2020-01-15", { kind: "date" }, "en-US")).toBe("2020-01-15");
    });

    test("falls back without throwing for invalid configuration or values", () => {
        expect(formatCellValue(12.345, { kind: "number", minimumFractionDigits: 3, maximumFractionDigits: 2 }, "en-US")).toBe("12.345");
        expect(formatCellValue(12.345, { kind: "currency", currency: "US" }, "en-US")).toBe("12.345");
        expect(formatCellValue(Number.NaN, { kind: "number" }, "en-US")).toBe("NaN");
        expect(formatCellValue(1, { kind: "date" }, "!!!")).toBe("1");
    });

    test("reuses formatters and bounds cache growth", () => {
        clearFormatCache();
        formatCellValue(1.2, { kind: "number", maximumFractionDigits: 2 }, "en-US");
        const first = getFormatCacheStats();
        formatCellValue(3.4, { kind: "number", maximumFractionDigits: 2 }, "en-US");
        const second = getFormatCacheStats();
        expect(first.number).toBe(1);
        expect(second.number).toBe(1);
        for (let index = 0; index < 150; index++) {
            formatCellValue(index / 10, { kind: "number", maximumFractionDigits: index % 20 }, `en-US-u-nu-${index}`);
        }
        expect(getFormatCacheStats().number).toBeLessThanOrEqual(100);
    });
});
