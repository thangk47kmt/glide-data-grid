import { isFormulaError, type FormulaValue } from "./formula.js";
import type { SpreadsheetFilter } from "./view.js";

export const SORT_LOCALE = "en-US";

export function normalizeSearch(value: string): string {
    return value.trim().toLocaleLowerCase(SORT_LOCALE);
}

export function normalizeSearchValue(value: string): string {
    return value.toLocaleLowerCase(SORT_LOCALE);
}

export function displayComparable(value: FormulaValue): string | number | boolean {
    if (isFormulaError(value)) return value.code;
    if (value === null) return "";
    return value;
}

export function numericExpected(value: SpreadsheetFilter["value"]): number | undefined {
    if (typeof value === "number") return Number.isFinite(value) ? value : undefined;
    if (typeof value !== "string" || value.trim() === "") return undefined;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
}

export function compareForFilter(actual: FormulaValue, expected: SpreadsheetFilter["value"]): number | undefined {
    if (isFormulaError(actual) || actual === null || expected === undefined || expected === null) return undefined;
    if (typeof actual === "number") {
        const expectedNumber = numericExpected(expected);
        return expectedNumber === undefined || !Number.isFinite(actual) ? undefined : actual - expectedNumber;
    }
    if (typeof actual === "boolean") return typeof expected === "boolean" ? Number(actual) - Number(expected) : undefined;
    if (typeof expected !== "string") return undefined;
    return actual.localeCompare(expected, SORT_LOCALE, { numeric: true, sensitivity: "base" });
}

export function equalsForFilter(actual: FormulaValue, expected: SpreadsheetFilter["value"]): boolean {
    if (isFormulaError(actual)) return typeof expected === "string" && actual.code.toLowerCase() === expected.toLowerCase();
    if (actual === null) return false;
    if (typeof actual === "number") {
        const expectedNumber = numericExpected(expected);
        return expectedNumber !== undefined && Number.isFinite(actual) && actual === expectedNumber;
    }
    if (typeof actual === "boolean") return typeof expected === "boolean" ? actual === expected : String(actual) === String(expected).toLowerCase();
    return typeof expected === "string" && actual.toLocaleLowerCase(SORT_LOCALE) === expected.toLocaleLowerCase(SORT_LOCALE);
}

/** Unknown filter columns intentionally match no rows. */
export function matchesFilter(value: FormulaValue, filter: SpreadsheetFilter): boolean {
    const expected = filter.value;
    switch (filter.operator) {
        case "contains": {
            if (value === null) return false;
            const actual = isFormulaError(value) ? value.code : String(value);
            return actual.toLocaleLowerCase(SORT_LOCALE).includes(String(expected ?? "").toLocaleLowerCase(SORT_LOCALE));
        }
        case "equals": return equalsForFilter(value, expected);
        case "not-equals": return value === null ? false : !matchesFilter(value, { ...filter, operator: "equals" });
        case "gt": return (compareForFilter(value, expected) ?? Number.NaN) > 0;
        case "gte": return (compareForFilter(value, expected) ?? Number.NaN) >= 0;
        case "lt": return (compareForFilter(value, expected) ?? Number.NaN) < 0;
        case "lte": return (compareForFilter(value, expected) ?? Number.NaN) <= 0;
        case "empty": return value === null || value === "";
        case "not-empty": return value !== null && value !== "";
    }
}

export function compareForSort(left: FormulaValue, right: FormulaValue): number {
    const leftComparable = displayComparable(left);
    const rightComparable = displayComparable(right);
    if (typeof leftComparable === "number" && typeof rightComparable === "number") {
        if (!Number.isFinite(leftComparable) || !Number.isFinite(rightComparable)) return String(leftComparable).localeCompare(String(rightComparable), SORT_LOCALE);
        return leftComparable - rightComparable;
    }
    if (typeof leftComparable === "boolean" && typeof rightComparable === "boolean") return Number(leftComparable) - Number(rightComparable);
    return String(leftComparable).localeCompare(String(rightComparable), SORT_LOCALE, { numeric: true, sensitivity: "base" });
}
