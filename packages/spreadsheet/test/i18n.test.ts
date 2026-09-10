import { describe, expect, test } from "vitest";
import {
    createSpreadsheetI18n,
    normalizeSpreadsheetLocale,
    spreadsheetMessages,
} from "../src/i18n.js";

describe("spreadsheet i18n contract", () => {
    test("normalizes supported language tags and safely falls back to English", () => {
        expect(normalizeSpreadsheetLocale("vi")).toBe("vi");
        expect(normalizeSpreadsheetLocale("vi-VN")).toBe("vi");
        expect(normalizeSpreadsheetLocale("VI_vn")).toBe("vi");
        expect(normalizeSpreadsheetLocale("en-US")).toBe("en");
        expect(normalizeSpreadsheetLocale("fr-FR")).toBe("en");
        expect(normalizeSpreadsheetLocale(undefined)).toBe("en");
        expect(normalizeSpreadsheetLocale(null)).toBe("en");
    });

    test("provides Vietnamese labels and filter operator translations", () => {
        const i18n = createSpreadsheetI18n("vi-VN");
        expect(i18n.locale).toBe("vi");
        expect(i18n.t("pageSize")).toBe("Kích thước trang");
        expect(i18n.t("firstPage")).toBe("Đầu");
        expect(i18n.filterOperator("contains")).toBe("Chứa");
        expect(i18n.filterOperator("not-empty")).toBe("Không trống");
    });

    test("merges partial custom messages over the selected locale", () => {
        const i18n = createSpreadsheetI18n({ locale: "vi", messages: { pageSize: "Rows per page", clearFilter: "Reset filter" } });
        expect(i18n.t("pageSize")).toBe("Rows per page");
        expect(i18n.t("clearFilter")).toBe("Reset filter");
        expect(i18n.t("firstPage")).toBe(spreadsheetMessages.vi.firstPage);
        expect(i18n.t("sortAscending")).toBe(spreadsheetMessages.vi.sortAscending);
    });

    test("supports locale plus partial messages in the positional form", () => {
        const i18n = createSpreadsheetI18n("en-US", { loadedPage: "Loaded in {ms}ms" });
        expect(i18n.t("loadedPage", { ms: 12.5 })).toBe("Loaded in 12.5ms");
        expect(i18n.t("rowsColumns", { rows: 500_000, columns: 23 })).toBe("500,000 rows × 23 columns");
    });

    test("does not throw for incomplete interpolation data or invalid runtime overrides", () => {
        const i18n = createSpreadsheetI18n({
            locale: "unsupported-locale",
            messages: { matchingRows: "{count} records", pageSize: 42 as unknown as string },
        });
        expect(i18n.locale).toBe("en");
        expect(i18n.t("matchingRows")).toBe("{count} records");
        expect(i18n.t("rowsColumns", { rows: 1 })).toBe("1 rows × {columns} columns");
        expect(i18n.t("pageSize")).toBe(spreadsheetMessages.en.pageSize);
    });

    test("formats dynamic counts and accessibility labels with the active locale", () => {
        const i18n = createSpreadsheetI18n("vi-VN");
        expect(i18n.t("pastedCells", { count: 1_234 })).toBe("Đã dán 1.234 ô");
        expect(i18n.t("columnMenu", { column: "Số lượng" })).toBe("Menu cột Số lượng");
        expect(i18n.filterOperator("gte")).toBe("Lớn hơn hoặc bằng");
    });
});
