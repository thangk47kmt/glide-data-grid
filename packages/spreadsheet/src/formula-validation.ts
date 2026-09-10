import { formulaError, isFormulaError, type FormulaError, type FormulaValue } from "./formula.js";
import type { SpreadsheetColumn, SpreadsheetColumnDataType } from "./model.js";

/** Runtime scalar types supported by a typed spreadsheet column. */
export type FormulaScalarType = "text" | "number" | "boolean";
export type FormulaSemanticType = SpreadsheetColumnDataType;

export interface FormulaTypeValidationOptions {
    /** Blank references are allowed by default, matching spreadsheet blanks. */
    readonly allowNull?: boolean;
}

function columnLabel(column: SpreadsheetColumn): string {
    return column.title.trim() || column.id;
}

/** Resolves the most specific column contract, keeping `type` as fallback. */
export function formulaColumnSemanticType(column: SpreadsheetColumn): FormulaSemanticType | undefined {
    return column.dataType ?? column.semanticType ?? column.type;
}

function primitiveType(type: FormulaSemanticType): FormulaScalarType {
    return type === "number" || type === "range" || type === "stars" || type === "sparkline" ? "number" : type === "boolean" ? "boolean" : "text";
}

function isValidUri(value: string): boolean {
    try {
        void new URL(value);
        return true;
    } catch {
        return false;
    }
}

function semanticTypeError(value: FormulaValue, destination: SpreadsheetColumn, expected: FormulaSemanticType): FormulaError | undefined {
    if (value === null || isFormulaError(value)) return undefined;
    const actual = formulaValueType(value);
    if (expected === "uri" && typeof value === "string") {
        if (!isValidUri(value)) {
            return formulaError("#VALUE!", `Formula result is not a valid URI for column '${columnLabel(destination)}'`);
        }
    }
    if (expected === "image" && typeof value === "string" && !value.startsWith("data:image/") && !isValidUri(value)) {
        return formulaError("#VALUE!", `Formula result is not a valid image URL for column '${columnLabel(destination)}'`);
    }
    if (expected === "date" && typeof value === "string") {
        const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
        if (match === null) {
            return formulaError("#VALUE!", `Formula result is not a valid ISO date for column '${columnLabel(destination)}'`);
        }
        const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
        if (Number.isNaN(date.getTime()) || date.getUTCFullYear() !== Number(match[1]) || date.getUTCMonth() !== Number(match[2]) - 1 || date.getUTCDate() !== Number(match[3])) {
            return formulaError("#VALUE!", `Formula result is not a valid ISO date for column '${columnLabel(destination)}'`);
        }
    }
    if (expected === "dropdown" && typeof value === "string" && destination.allowedValues !== undefined && !destination.allowedValues.includes(value)) {
        return formulaError("#VALUE!", `Formula result '${value}' is not an allowed value for dropdown '${columnLabel(destination)}'`);
    }
    if ((expected === "tags" || expected === "user-profile") && typeof value === "string" && destination.allowedValues !== undefined && !destination.allowedValues.includes(value)) {
        return formulaError("#VALUE!", `Formula result '${value}' is not an allowed value for column '${columnLabel(destination)}'`);
    }
    if (expected === "multi-select" && typeof value === "string" && destination.allowedValues !== undefined) {
        const invalid = value.split(",").map(item => item.trim()).filter(item => item !== "" && !destination.allowedValues!.includes(item));
        if (invalid.length > 0) return formulaError("#VALUE!", `Formula result contains invalid option '${invalid[0]}' for column '${columnLabel(destination)}'`);
    }
    if (expected === "range" && typeof value === "number") {
        const min = destination.min ?? 0;
        const max = destination.max ?? 100;
        if (!Number.isFinite(value) || value < min || value > max) return formulaError("#VALUE!", `Formula result ${value} is outside range ${min}-${max} for column '${columnLabel(destination)}'`);
    }
    if (expected === "stars" && typeof value === "number") {
        const min = destination.min ?? 0;
        const max = destination.max ?? 5;
        if (!Number.isInteger(value) || value < min || value > max) return formulaError("#VALUE!", `Formula result ${value} must be an integer between ${min} and ${max} for stars column '${columnLabel(destination)}'`);
    }
    void actual;
    return undefined;
}

/** Gets the scalar type of an evaluated formula value. Formula errors are not scalar values. */
export function formulaValueType(value: FormulaValue): FormulaScalarType | "null" | "error" {
    if (isFormulaError(value)) return "error";
    if (value === null) return "null";
    if (typeof value === "string") return "text";
    if (typeof value === "number") return "number";
    return "boolean";
}

/**
 * Validates that an evaluated formula result can be stored in its destination
 * column. An untyped column intentionally accepts every scalar value.
 * Existing formula errors are returned unchanged so their specific code and
 * message remain visible to the user.
 */
export function validateFormulaResultType(
    value: FormulaValue,
    destination: SpreadsheetColumn,
    options: FormulaTypeValidationOptions = {}
): FormulaError | undefined {
    if (isFormulaError(value)) return value;
    const actual = formulaValueType(value);
    if (actual === "null") {
        if (options.allowNull !== false) return undefined;
        return formulaError("#VALUE!", `Formula result is blank, but column '${columnLabel(destination)}' does not allow blank values`);
    }

    const expected = formulaColumnSemanticType(destination);
    if (expected === undefined) return undefined;
    if (actual !== primitiveType(expected)) {
        return formulaError(
            "#VALUE!",
            `Formula result type '${actual}' does not match column '${columnLabel(destination)}' type '${expected}'`
        );
    }
    const domainError = semanticTypeError(value, destination, expected);
    if (domainError !== undefined) return domainError;
    return undefined;
}

/**
 * Validates a direct cell/column reference when both source and destination
 * columns are known. This deliberately checks the declared source type first:
 * `=Text1` in a Number column is rejected even when the text happens to look
 * numeric, preserving the column's data contract.
 */
export function validateFormulaReferenceType(
    value: FormulaValue,
    source: SpreadsheetColumn,
    destination: SpreadsheetColumn,
    options: FormulaTypeValidationOptions = {}
): FormulaError | undefined {
    if (isFormulaError(value)) return value;
    const sourceType = formulaColumnSemanticType(source);
    const destinationType = formulaColumnSemanticType(destination);
    if (sourceType !== undefined && destinationType !== undefined && sourceType !== destinationType) {
        return formulaError(
            "#VALUE!",
            `Cannot reference '${columnLabel(source)}' (${sourceType}) from '${columnLabel(destination)}' (${destinationType}): source and destination types must match`
        );
    }
    return validateFormulaResultType(value, destination, options);
}
