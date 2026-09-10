import { describe, expect, test } from "vitest";
import { CsvParseError, parseCsv, stringifyCsv } from "../src/csv.js";

describe("CSV", () => {
    test("round-trips quoting, delimiters, embedded newlines and Unicode", () => {
        const rows = [["plain", "a,b", "say \"hi\"", "line 1\nline 2", "日本語"]];
        const csv = stringifyCsv(rows);
        expect(csv).toBe("plain,\"a,b\",\"say \"\"hi\"\"\",\"line 1\nline 2\",日本語");
        expect(parseCsv(csv)).toEqual(rows);
    });

    test("supports configurable delimiter, CRLF/LF and preserves trailing empty fields", () => {
        expect(parseCsv("a;b;;\r\nc;d;\r\n", { delimiter: ";" })).toEqual([["a", "b", "", ""], ["c", "d", ""]]);
        expect(parseCsv("a,b\n\n")).toEqual([["a", "b"], [""]]);
        expect(stringifyCsv([["a", "b"], ["c", "d"]], { lineEnding: "\n" })).toBe("a,b\nc,d");
        expect(parseCsv("")).toEqual([]);
    });

    test("opts into finite number and TRUE/FALSE inference and empty nulls", () => {
        expect(parseCsv("1,1e2,TRUE,FALSE,,=A1", { inferTypes: true, emptyValue: "null" })).toEqual([[1, 100, true, false, null, "=A1"]]);
        expect(parseCsv("NaN,Infinity,-2", { inferTypes: true })).toEqual([["NaN", "Infinity", -2]]);
        expect(parseCsv(",", { emptyValue: "string" })).toEqual([["", ""]]);
    });

    test("keeps formula-looking text unless explicitly rejected and exports safe for Excel", () => {
        expect(parseCsv("=SUM(A1:A2),+cmd,-cmd,@cmd")).toEqual([["=SUM(A1:A2)", "+cmd", "-cmd", "@cmd"]]);
        expect(() => parseCsv("x,=SUM(A1)", { formulaPolicy: "reject" })).toThrowError(CsvParseError);
        const error = (() => {
            try {
                parseCsv("x,=SUM(A1)", { formulaPolicy: "reject" });
            } catch (caught) {
                return caught as CsvParseError;
            }
            return undefined;
        })();
        expect(error?.row).toBe(1);
        expect(error?.column).toBe(2);
        expect(stringifyCsv([["=SUM(A1)", "+cmd", "-cmd", "@cmd", -2]], { safeForExcel: true })).toBe("'=SUM(A1),'+cmd,'-cmd,'@cmd,-2");
        expect(stringifyCsv([["=A1", 10]], { formulaPolicy: "values", formulaValues: [[42, 10]] })).toBe("42,10");
    });

    test("rejects malformed quotes with a structured location", () => {
        expect(() => parseCsv('ok,"unterminated')).toThrowError(CsvParseError);
        try {
            parseCsv('ok,"unterminated');
        } catch (caught) {
            const error = caught as CsvParseError;
            expect(error.code).toBe("unterminated-quote");
            expect(error.row).toBe(1);
            expect(error.column).toBe(2);
        }
        expect(() => parseCsv('a"b')).toThrowError(/Quotes must start/);
        expect(() => parseCsv('"a"x')).toThrowError(/closing quote/);
    });

    test("enforces row, column and cell-length limits", () => {
        expect(() => parseCsv("a\nb", { maxRows: 1 })).toThrowError(CsvParseError);
        expect(() => parseCsv("a,b", { maxColumns: 1 })).toThrowError(CsvParseError);
        expect(() => parseCsv("abcd", { maxCellLength: 3 })).toThrowError(CsvParseError);
        try {
            parseCsv("a,b", { maxColumns: 1 });
        } catch (caught) {
            const error = caught as CsvParseError;
            expect(error.code).toBe("max-columns");
            expect(error.row).toBe(1);
            expect(error.column).toBe(2);
        }
        expect(() => parseCsv("a", { delimiter: "||" })).toThrowError(TypeError);
        expect(() => parseCsv("a", { delimiter: '"' })).toThrowError(TypeError);
        expect(() => parseCsv("a", { delimiter: "\r" })).toThrowError(TypeError);
        expect(() => parseCsv("a", { delimiter: "\n" })).toThrowError(TypeError);
    });
});
