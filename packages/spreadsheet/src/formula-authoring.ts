import type { FormulaFunctionRegistry } from "./function-registry.js";
import { columnNameToIndex, compileFormula } from "./formula.js";

export type FormulaDiagnosticSeverity = "error" | "warning" | "info";

export interface FormulaDiagnostic {
    readonly code: string;
    readonly message: string;
    readonly severity: FormulaDiagnosticSeverity;
    readonly start?: number;
    readonly end?: number;
}

export interface FormulaAuthoringContext {
    /** Resolves a structured column name. Omit when column validation is unavailable. */
    readonly resolveColumn?: (name: string) => number | undefined;
    readonly functionRegistry?: FormulaFunctionRegistry;
    /** Optional ordered column names used for structured-reference completion. */
    readonly structuredColumns?: readonly string[];
}

export interface FormulaFunctionMetadata {
    readonly name: string;
    readonly signature: string;
    readonly description: string;
}

/** Metadata for every built-in function implemented by the spreadsheet engine. */
export const BUILT_IN_FUNCTION_CATALOG: readonly FormulaFunctionMetadata[] = [
    { name: "ABS", signature: "ABS(number)", description: "Returns the absolute value of a number." },
    { name: "AND", signature: "AND(value1, value2, ...)", description: "Returns TRUE when all values are truthy." },
    { name: "AVERAGE", signature: "AVERAGE(value1, value2, ...)", description: "Returns the arithmetic mean of numeric values." },
    { name: "CONCAT", signature: "CONCAT(value1, value2, ...)", description: "Joins values into one text string." },
    { name: "COUNTA", signature: "COUNTA(value1, value2, ...)", description: "Counts non-empty values." },
    { name: "COUNT", signature: "COUNT(value1, value2, ...)", description: "Counts numeric values." },
    { name: "DATE", signature: "DATE(year, month, day)", description: "Creates a numeric date serial." },
    { name: "DAY", signature: "DAY(serial)", description: "Returns the day of month from a date serial." },
    { name: "IF", signature: "IF(condition, whenTrue, whenFalse)", description: "Chooses a value based on a condition." },
    { name: "INDEX", signature: "INDEX(range, row, column)", description: "Returns a value at a range position." },
    { name: "LEFT", signature: "LEFT(text, count)", description: "Returns characters from the left of text." },
    { name: "LEN", signature: "LEN(text)", description: "Returns the number of characters in text." },
    { name: "LOWER", signature: "LOWER(text)", description: "Converts text to lowercase." },
    { name: "MATCH", signature: "MATCH(value, range, matchType)", description: "Returns the position of an exact value in a range." },
    { name: "MAX", signature: "MAX(value1, value2, ...)", description: "Returns the largest numeric value." },
    { name: "MID", signature: "MID(text, start, count)", description: "Returns characters from the middle of text." },
    { name: "MIN", signature: "MIN(value1, value2, ...)", description: "Returns the smallest numeric value." },
    { name: "MONTH", signature: "MONTH(serial)", description: "Returns the month from a date serial." },
    { name: "NOT", signature: "NOT(value)", description: "Reverses a truthy value." },
    { name: "OR", signature: "OR(value1, value2, ...)", description: "Returns TRUE when any value is truthy." },
    { name: "RIGHT", signature: "RIGHT(text, count)", description: "Returns characters from the right of text." },
    { name: "ROUND", signature: "ROUND(number, digits)", description: "Rounds a number to a number of digits." },
    { name: "SUM", signature: "SUM(value1, value2, ...)", description: "Returns the sum of numeric values." },
    { name: "TODAY", signature: "TODAY()", description: "Returns today's date as a numeric serial." },
    { name: "TRIM", signature: "TRIM(text)", description: "Removes leading, trailing and repeated spaces." },
    { name: "UPPER", signature: "UPPER(text)", description: "Converts text to uppercase." },
    { name: "XLOOKUP", signature: "XLOOKUP(value, lookupRange, returnRange, notFound)", description: "Looks up a value in one range and returns a corresponding value." },
    { name: "YEAR", signature: "YEAR(serial)", description: "Returns the year from a date serial." },
];

const builtIns = new Map(BUILT_IN_FUNCTION_CATALOG.map(metadata => [metadata.name, metadata]));
const a1Helpers = Array.from({ length: 26 }, (_, index) => String.fromCharCode(65 + index) + "1");

function diagnostic(code: string, message: string, start?: number, end?: number): FormulaDiagnostic {
    return { code, message, severity: "error", ...(start === undefined ? {} : { start }), ...(end === undefined ? {} : { end }) };
}

function pushDiagnostic(result: FormulaDiagnostic[], item: FormulaDiagnostic): void {
    if (!result.some(existing => existing.code === item.code && existing.start === item.start && existing.message === item.message)) result.push(item);
}

function customNames(context: FormulaAuthoringContext | undefined): Set<string> {
    try {
        return new Set((context?.functionRegistry?.list() ?? []).map(name => name.toUpperCase()));
    } catch {
        return new Set();
    }
}

function isA1Token(value: string): boolean {
    return /^\$?[A-Za-z]+\$?\d+$/.test(value);
}

function isInvalidA1Token(value: string): boolean {
    return /^\$?[A-Za-z]+\$?0+$/.test(value) || /^\$?[A-Za-z]+\$?\d+$/.test(value) && columnNameToIndex(value.replaceAll("$", "").replace(/\d+$/, "")) === undefined;
}

function scanFormula(source: string, context: FormulaAuthoringContext | undefined, result: FormulaDiagnostic[]): void {
    const expressionOffset = source.startsWith("=") ? 1 : 0;
    const expression = source.slice(expressionOffset);
    const knownCustom = customNames(context);
    let index = 0;
    while (index < expression.length) {
        const char = expression[index];
        if (char === '"') {
            index++;
            while (index < expression.length) {
                if (expression[index] === '"' && expression[index + 1] === '"') index += 2;
                else if (expression[index++] === '"') break;
            }
            continue;
        }
        if (char === "[") {
            const end = expression.indexOf("]", index + 1);
            if (end < 0) break;
            const rawName = expression.slice(index + 1, end).trim();
            const name = rawName.startsWith("@") ? rawName.slice(1).trim() : rawName;
            if (name.length === 0) pushDiagnostic(result, diagnostic("INVALID_REFERENCE", "Structured reference must include a column name", expressionOffset + index, expressionOffset + end + 1));
            else if (context?.resolveColumn !== undefined) {
                try {
                    if (context.resolveColumn(name) === undefined) pushDiagnostic(result, diagnostic("UNKNOWN_COLUMN", `Unknown structured column '${name}'`, expressionOffset + index, expressionOffset + end + 1));
                } catch {
                    pushDiagnostic(result, diagnostic("UNKNOWN_COLUMN", `Unable to resolve structured column '${name}'`, expressionOffset + index, expressionOffset + end + 1));
                }
            }
            index = end + 1;
            continue;
        }
        if (char === "#" && expression.slice(index).match(/^#REF!/i) !== null) {
            pushDiagnostic(result, diagnostic("INVALID_REFERENCE", "#REF! is not a valid input reference", expressionOffset + index, expressionOffset + index + 5));
            index += 5;
            continue;
        }
        const tokenMatch = expression.slice(index).match(/^\$?[A-Za-z]+\$?\d+|^[A-Za-z_][A-Za-z0-9_.]*/);
        if (tokenMatch === null) {
            index++;
            continue;
        }
        const token = tokenMatch[0];
        const tokenStart = expressionOffset + index;
        const tokenEnd = tokenStart + token.length;
        const afterToken = expression.slice(index + token.length).match(/^\s*\(/) !== null;
        if (afterToken) {
            const upper = token.toUpperCase();
            if (!builtIns.has(upper) && !knownCustom.has(upper)) pushDiagnostic(result, diagnostic("UNKNOWN_FUNCTION", `Unknown function '${token}'`, tokenStart, tokenEnd));
        } else if (isA1Token(token) && isInvalidA1Token(token)) {
            pushDiagnostic(result, diagnostic("INVALID_REFERENCE", `Invalid cell reference '${token}'`, tokenStart, tokenEnd));
        }
        index += token.length;
    }
}

/** Diagnoses authoring input without ever throwing, including malformed parser input. */
export function diagnoseFormula(source: string, context?: FormulaAuthoringContext): readonly FormulaDiagnostic[] {
    const result: FormulaDiagnostic[] = [];
    try {
        if (typeof source !== "string") return [diagnostic("INVALID_SOURCE", "Formula source must be a string")];
        if (!source.startsWith("=")) pushDiagnostic(result, diagnostic("MISSING_EQUALS", "Formula must start with '='", 0, 0));
        scanFormula(source, context, result);
        try {
            compileFormula(source.startsWith("=") ? source : `=${source}`);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            const invalidReference = /Unknown name '([^']+)'/.exec(message);
            if (!result.some(item => item.code === "INVALID_REFERENCE")) {
                pushDiagnostic(result, invalidReference !== null && isInvalidA1Token(invalidReference[1])
                    ? diagnostic("INVALID_REFERENCE", message)
                    : diagnostic(/Unexpected character/.test(message) ? "INVALID_TOKEN" : "PARSER_ERROR", message));
            }
        }
    } catch (error) {
        pushDiagnostic(result, diagnostic("AUTHORING_ERROR", error instanceof Error ? error.message : String(error)));
    }
    return result;
}

export type FormulaCompletionKind = "function" | "column" | "reference";

export interface FormulaCompletion {
    readonly label: string;
    readonly insertText: string;
    readonly kind: FormulaCompletionKind;
    readonly detail: string;
    readonly score: number;
    /** Source range to replace when applying insertText (end is exclusive). */
    readonly replaceStart: number;
    readonly replaceEnd: number;
}

function activePrefix(source: string, cursor: number): { prefix: string; structured: boolean; replaceStart: number; replaceEnd: number; currentRow: boolean } {
    const position = Math.max(0, Math.min(cursor, source.length));
    const before = source.slice(0, position);
    const openBracket = before.lastIndexOf("[");
    const closeBracket = before.lastIndexOf("]");
    const structured = openBracket > closeBracket;
    // Structured references are user-facing captions, so they may contain
    // spaces, punctuation and non-ASCII letters (for example `Đơn giá`).
    // Keep the whole text after `[` as the prefix and let the resolver/ranker
    // decide whether it is a valid column name.
    const match = before.match(/[A-Za-z0-9_.$]*$/);
    const structuredText = structured ? before.slice(openBracket + 1) : undefined;
    return {
        prefix: (structuredText ?? match?.[0] ?? "").replace(/^@/, "").trim(),
        structured,
        replaceStart: structured ? openBracket : position - (match?.[0]?.length ?? 0),
        replaceEnd: position,
        currentRow: structured && before.slice(openBracket, position).startsWith("[@"),
    };
}

function completionScore(label: string, prefix: string): number {
    const normalized = label.trim().normalize("NFKC").toLocaleUpperCase("en-US");
    const wanted = prefix.trim().normalize("NFKC").toLocaleUpperCase("en-US");
    if (wanted === "") return 1;
    if (normalized === wanted) return 0;
    if (normalized.startsWith(wanted)) return 1;
    return normalized.includes(wanted) ? 2 : Number.POSITIVE_INFINITY;
}

/** Returns deterministic, case-insensitive formula completions ranked by prefix match. */
export function getFormulaCompletions(source: string, cursor: number, context?: FormulaAuthoringContext): readonly FormulaCompletion[] {
    try {
        if (typeof source !== "string") return [];
        const { prefix, structured, replaceStart, replaceEnd, currentRow } = activePrefix(source, cursor);
        const normalizedPrefix = prefix.toUpperCase();
        const result = new Map<string, FormulaCompletion>();
        if (structured) {
            for (const name of context?.structuredColumns ?? []) {
                if (context?.resolveColumn !== undefined) {
                    try {
                        // Do not offer an ambiguous caption: inserting it
                        // would create a formula that cannot be resolved.
                        if (context.resolveColumn(name) === undefined) continue;
                    } catch {
                        continue;
                    }
                }
                const score = completionScore(name, prefix);
                const key = name.trim().normalize("NFKC").toLocaleLowerCase("en-US");
                if (score !== Number.POSITIVE_INFINITY && key !== "" && !result.has(key)) result.set(key, { label: name, insertText: currentRow ? `[@${name}]` : `[${name}]`, kind: "column", detail: "Structured column", score, replaceStart, replaceEnd });
            }
        } else {
            for (const metadata of BUILT_IN_FUNCTION_CATALOG) {
                const score = completionScore(metadata.name, prefix);
                if (score !== Number.POSITIVE_INFINITY) result.set(metadata.name, { label: metadata.name, insertText: `${metadata.name}(`, kind: "function", detail: `${metadata.signature} — ${metadata.description}`, score, replaceStart, replaceEnd });
            }
            try {
                for (const name of context?.functionRegistry?.list() ?? []) {
                    const upper = name.toUpperCase();
                    const score = completionScore(upper, prefix);
                    if (score !== Number.POSITIVE_INFINITY && !result.has(upper)) result.set(upper, { label: upper, insertText: `${upper}(`, kind: "function", detail: "Custom formula function", score, replaceStart, replaceEnd });
                }
            } catch {
                // A completion UI should remain usable if a supplied registry is malformed.
            }
            if (/^\$?[A-Za-z]{0,2}\d*$/.test(normalizedPrefix)) {
                for (const name of a1Helpers) {
                    const score = completionScore(name, prefix);
                    if (score !== Number.POSITIVE_INFINITY && !result.has(name)) result.set(name, { label: name, insertText: name, kind: "reference", detail: "Cell reference helper", score, replaceStart, replaceEnd });
                }
            }
        }
        return [...result.values()].sort((left, right) => left.score - right.score || left.label.localeCompare(right.label, "en-US", { sensitivity: "base" })).slice(0, 50);
    } catch {
        return [];
    }
}
