import { describe, expect, test } from "vitest";
import { aggregateValues, aggregateVisibleValues, createAggregateAccumulator } from "../src/aggregate.js";

describe("aggregate engine", () => {
    test("aggregates mixed values while ignoring null, empty text and non-numeric types", () => {
        expect(aggregateValues([1, "2", "", null, true, { kind: "error", code: "#VALUE!" }, Number.POSITIVE_INFINITY])).toEqual({
            countAll: 5,
            countNumbers: 1,
            sum: 1,
            average: 1,
            min: 1,
            max: 1,
            errors: 1,
        });
    });

    test("supports explicit numeric string and boolean policies", () => {
        expect(aggregateValues([1, "2.5", true, false, "not a number"], {
            numericStringPolicy: "coerce",
            booleanPolicy: "numeric",
        })).toEqual({
            countAll: 5,
            countNumbers: 4,
            sum: 4.5,
            average: 1.125,
            min: 0,
            max: 2.5,
            errors: 0,
        });
    });

    test("returns null metrics for empty input and can propagate the first error", () => {
        expect(aggregateValues([null, ""])).toEqual({ countAll: 0, countNumbers: 0, sum: null, average: null, min: null, max: null, errors: 0 });
        const result = aggregateValues([2, { kind: "error", code: "#REF!" }, 4, { kind: "error", code: "#VALUE!" }], { errorPolicy: "include-first-error" });
        expect(result.countAll).toBe(4);
        expect(result.errors).toBe(2);
        expect(result.sum).toMatchObject({ kind: "error", code: "#REF!" });
        expect(result.average).toMatchObject({ kind: "error", code: "#REF!" });
        expect(result.min).toMatchObject({ kind: "error", code: "#REF!" });
    });

    test("merges streaming chunks with the same aggregate result", () => {
        const values = [1e16, 1, -1e16, 0.1, 0.2, 0.3];
        const onePass = createAggregateAccumulator();
        values.forEach(value => onePass.add(value));
        const chunks = createAggregateAccumulator();
        const left = createAggregateAccumulator();
        const right = createAggregateAccumulator();
        values.slice(0, 3).forEach(value => left.add(value));
        values.slice(3).forEach(value => right.add(value));
        chunks.merge(left).merge(right);
        expect(chunks.finalize()).toEqual(onePass.finalize());
        expect(chunks.finalize().sum).toBeCloseTo(1.6, 12);
    });

    test("rejects merging accumulators with incompatible policies", () => {
        const strict = createAggregateAccumulator();
        const coerce = createAggregateAccumulator({ numericStringPolicy: "coerce" });
        expect(() => strict.merge(coerce)).toThrowError(TypeError);
        const ignoreErrors = createAggregateAccumulator();
        const includeErrors = createAggregateAccumulator({ errorPolicy: "include-first-error" });
        expect(() => ignoreErrors.merge(includeErrors)).toThrowError(TypeError);
        strict.add(1);
        expect(strict.finalize().sum).toBe(1);
    });

    test("uses compensated summation for a large stream", () => {
        const accumulator = createAggregateAccumulator();
        for (let index = 0; index < 100_000; index++) accumulator.add(0.1);
        const result = accumulator.finalize();
        expect(result.countNumbers).toBe(100_000);
        expect(result.sum).toBeCloseTo(10_000, 12);
        expect(result.average).toBeCloseTo(0.1, 14);
    });

    test("supports visibility predicates through the array adapter", () => {
        expect(aggregateVisibleValues([1, 2, 3, 4], (_value, index) => index % 2 === 0)).toMatchObject({
            countAll: 2,
            countNumbers: 2,
            sum: 4,
            average: 2,
            min: 1,
            max: 3,
        });
        expect(aggregateValues([1, 2, 3], { isVisible: (_value, index) => index > 0 }).sum).toBe(5);
    });
});
