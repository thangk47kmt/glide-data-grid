import { describe, expect, test } from "vitest";
import { History } from "../src/history.js";
import { findMatches, replaceAll, replaceOne } from "../src/find-replace.js";
import { SpreadsheetModel, type CellInput } from "../src/model.js";

const columns = [
    { id: "item", title: "Item", type: "text" as const },
    { id: "quantity", title: "Quantity", type: "number" as const },
    { id: "total", title: "Total", type: "number" as const },
    { id: "note", title: "Note", type: "text" as const },
];

function makeModel(): SpreadsheetModel {
    return new SpreadsheetModel(columns, 4, [
        ["Alpha", 10, "=B1*2", "foo"],
        ["alpha", 20, "=B2*2", "Foo"],
        ["Beta", null, "=1/0", "bar"],
        ["item10", 2, null, "$1"],
    ]);
}

describe("find and replace", () => {
    test("searches raw formulas separately from computed display values", () => {
        const model = makeModel();
        expect(findMatches(model, { query: "=B1*2" }).matches.map(match => [match.col, match.row])).toEqual([[2, 0]]);
        expect(findMatches(model, { query: "20", scope: "computed" }).matches.map(match => [match.col, match.row])).toEqual([[2, 0], [1, 1]]);
        expect(findMatches(model, { query: "#DIV/0!", scope: "computed" }).matches.map(match => [match.col, match.row])).toEqual([[2, 2]]);
    });

    test("supports case, whole-cell, regex, column and row bounds", () => {
        const model = makeModel();
        expect(findMatches(model, { query: "alpha" }).matches.map(match => match.row)).toEqual([0, 1]);
        expect(findMatches(model, { query: "alpha", caseSensitive: true }).matches.map(match => match.row)).toEqual([1]);
        expect(findMatches(model, { query: "Alpha", wholeCell: true, caseSensitive: true }).matches.map(match => match.row)).toEqual([0]);
        expect(findMatches(model, { query: "^a.*", regexp: true }).matches.map(match => match.row)).toEqual([0, 1]);
        expect(findMatches(model, { query: "alpha", columns: [0], rowRange: { start: 1, end: 1 } }).matches.map(match => match.row)).toEqual([1]);
    });

    test("reports invalid regex and skips protected cells without throwing", () => {
        const model = makeModel();
        const invalid = findMatches(model, { query: "[", regexp: true });
        expect(invalid.matches).toEqual([]);
        expect(invalid.diagnostic?.code).toBe("INVALID_REGEX");
        expect(findMatches(model, { query: "foo", isProtected: (col, row) => col === 3 && row === 0 }).matches.map(match => match.row)).toEqual([1]);
        expect(findMatches(model, { query: "foo", maxResults: 1 }).truncated).toBe(true);
        expect(findMatches(model, { query: "o", columns: [3], maxResults: 1 }).matches).toHaveLength(1);
        expect(findMatches(model, { query: "o", columns: [3], maxResults: 1 }).truncated).toBe(true);
    });

    test("rejects empty queries and invalid runtime scopes deterministically", () => {
        const model = makeModel();
        expect(findMatches(model, { query: "" })).toMatchObject({
            matches: [],
            truncated: false,
            diagnostic: { code: "INVALID_OPTIONS" },
        });
        expect(findMatches(model, { query: "alpha", scope: "invalid" as never }).diagnostic?.code).toBe("INVALID_OPTIONS");
        expect(replaceAll(model, { query: "", replacement: "x" })).toMatchObject({
            matches: [],
            truncated: false,
            replacedCount: 0,
            skippedCount: 0,
            diagnostic: { code: "INVALID_OPTIONS" },
        });
        expect(replaceAll(model, { query: "alpha", replacement: "x", scope: "invalid" as never }).diagnostic?.code).toBe("INVALID_OPTIONS");
    });

    test("replace is raw-input-only, returns normalized transactions and does not mutate", () => {
        const model = makeModel();
        const formula = replaceOne(model, { query: "B1", replacement: "C1" });
        expect(formula.replacedCount).toBe(1);
        expect(formula.transaction?.edits).toEqual([{ location: [2, 0], before: "=B1*2", after: "=C1*2" }]);
        expect(model.getInput(2, 0)).toBe("=B1*2");
        expect(replaceAll(model, { query: "foo", replacement: "$1" }).transaction?.edits[0]?.after).toBe("$1");
        expect(replaceAll(model, { query: "(foo)", replacement: "$1!", regexp: true, caseSensitive: false }).replacedCount).toBe(2);
        expect(replaceAll(model, { query: "20", replacement: "99", scope: "computed" }).diagnostic?.code).toBe("COMPUTED_SCOPE_READ_ONLY");
        const noMatch = replaceOne(model, { query: "missing", replacement: "x" });
        expect(noMatch.transaction).toBeUndefined();
        expect(noMatch.truncated).toBe(false);
        expect(noMatch.replacedCount).toBe(0);
        expect(noMatch.skippedCount).toBe(0);
    });

    test("reports capped replace-all results instead of silently applying a partial scan", () => {
        const model = makeModel();
        const result = replaceAll(model, { query: "a", replacement: "x", maxResults: 1 });
        expect(result.matches).toHaveLength(1);
        expect(result.truncated).toBe(true);
        expect(result.replacedCount).toBe(1);
        expect(result.skippedCount).toBe(0);
        expect(result.transaction?.edits).toHaveLength(1);
    });

    test("replaces non-string raw values as text by default and supports typed coercion", () => {
        const model = makeModel();
        const defaultResult = replaceOne(model, { query: "10", replacement: "11" });
        expect(defaultResult.transaction?.edits).toEqual([{ location: [1, 0], before: 10, after: "11" }]);
        expect(defaultResult.skippedCount).toBe(0);

        const typedResult = replaceOne(model, {
            query: "10",
            replacement: "11",
            coerceReplacement: text => Number(text),
        });
        expect(typedResult.transaction?.edits).toEqual([{ location: [1, 0], before: 10, after: 11 }]);
        expect(typedResult.skippedCount).toBe(0);

        const skipped = replaceOne(model, {
            query: "10",
            replacement: "11",
            coerceReplacement: () => {
                throw new Error("cannot parse");
            },
        });
        expect(skipped.transaction).toBeUndefined();
        expect(skipped.replacedCount).toBe(0);
        expect(skipped.skippedCount).toBe(1);
        expect(skipped.skippedReasons).toEqual(["cannot parse"]);
    });

    test("transactions can be executed by caller and preserve formula behavior", () => {
        const model = makeModel();
        const result = replaceOne(model, { query: "B1", replacement: "C1", transactionId: "replace-test" });
        const history = new History<CellInput>({ apply: edit => model.setCell(edit.location[0], edit.location[1], edit.after) });
        if (result.transaction !== undefined) history.execute(result.transaction);
        expect(model.getInput(2, 0)).toBe("=C1*2");
    });
});
