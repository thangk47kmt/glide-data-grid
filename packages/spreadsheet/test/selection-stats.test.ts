import { describe, expect, test } from "vitest";
import { aggregateSelectionRange } from "../src/selection-stats.js";

describe("aggregateSelectionRange", () => {
    test("maps display rows to source rows and aggregates the rectangle", () => {
        const result = aggregateSelectionRange(
            { x: 0, y: 0, width: 2, height: 2 },
            [7, 3],
            2,
            (column, sourceRow) => sourceRow * 10 + column,
        );

        expect(result).toMatchObject({ kind: "ready", cellCount: 4 });
        if (result.kind !== "ready") throw new Error("expected a ready result");
        expect(result.aggregate).toMatchObject({
            countAll: 4,
            countNumbers: 4,
            sum: 202,
            average: 50.5,
            min: 30,
            max: 71,
        });
    });

    test("does not evaluate a selection above the safety cap", () => {
        let reads = 0;
        const result = aggregateSelectionRange(
            { x: 0, y: 0, width: 3, height: 2 },
            [0, 1],
            3,
            () => {
                reads++;
                return 1;
            },
            5,
        );

        expect(result).toEqual({ kind: "too-large", cellCount: 6, maxCells: 5 });
        expect(reads).toBe(0);
    });

    test("returns an empty result for an absent or clipped range", () => {
        expect(aggregateSelectionRange(undefined, [0], 2, () => 1)).toEqual({ kind: "empty", cellCount: 0 });
        expect(aggregateSelectionRange({ x: 2, y: 0, width: 1, height: 1 }, [0], 2, () => 1)).toEqual({ kind: "empty", cellCount: 0 });
    });
});
