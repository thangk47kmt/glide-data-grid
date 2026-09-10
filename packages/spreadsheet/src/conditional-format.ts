import type { FormulaErrorCode, FormulaValue } from "./formula.js";

export interface ConditionalFormatRange {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
}

/** A minimal renderer-independent style patch. Later matching rules override fields. */
export interface ConditionalFormatStyle {
    readonly bg?: string;
    readonly text?: string;
    readonly font?: string;
    readonly numberFormat?: string;
}

interface ConditionalFormatRuleBase {
    readonly id: string;
    readonly range: ConditionalFormatRange;
    readonly style: ConditionalFormatStyle;
    readonly stopIfTrue?: boolean;
}

export interface NumberCompareRule extends ConditionalFormatRuleBase {
    readonly kind: "number-compare";
    readonly operator: "eq" | "neq" | "gt" | "gte" | "lt" | "lte";
    readonly value: number;
}

export interface NumberBetweenRule extends ConditionalFormatRuleBase {
    readonly kind: "number-between";
    readonly min: number;
    readonly max: number;
}

export interface TextRule extends ConditionalFormatRuleBase {
    readonly kind: "text";
    readonly operator: "contains" | "equals";
    readonly value: string;
    readonly caseSensitive?: boolean;
}

export interface EmptyRule extends ConditionalFormatRuleBase {
    readonly kind: "empty" | "not-empty";
}

export interface DuplicateRule extends ConditionalFormatRuleBase {
    readonly kind: "duplicate" | "unique";
}

export interface FormulaErrorRule extends ConditionalFormatRuleBase {
    readonly kind: "formula-error";
    readonly code?: FormulaErrorCode;
}

export interface ConditionalFormatContext {
    readonly col: number;
    readonly row: number;
    readonly value: FormulaValue;
    readonly getCellValue?: CellValueAccessor;
}

export type ConditionalPredicate = (context: ConditionalFormatContext) => boolean;

export interface CustomConditionalRule extends ConditionalFormatRuleBase {
    readonly kind: "custom";
    readonly predicate: ConditionalPredicate;
}

export type ConditionalFormatRule = NumberCompareRule | NumberBetweenRule | TextRule | EmptyRule | DuplicateRule | FormulaErrorRule | CustomConditionalRule;
export type CellValueAccessor = (col: number, row: number) => FormulaValue;

export interface ConditionalFormatDiagnostic {
    readonly ruleId: string;
    readonly code: "custom-error" | "missing-value-accessor";
    readonly message: string;
}

export interface ConditionalFormatResult {
    readonly style: ConditionalFormatStyle;
    readonly matchedRuleIds: readonly string[];
    readonly diagnostics: readonly ConditionalFormatDiagnostic[];
}

export interface ConditionalFormatCompileOptions {
    readonly getCellValue?: CellValueAccessor;
}

export type ConditionalFormatErrorCode = "invalid-range" | "invalid-rule" | "invalid-number" | "invalid-between" | "duplicate-rule-id";

export class ConditionalFormatError extends Error {
    public readonly code: ConditionalFormatErrorCode;
    public readonly path: string;

    public constructor(code: ConditionalFormatErrorCode, message: string, path: string) {
        super(message);
        this.name = "ConditionalFormatError";
        this.code = code;
        this.path = path;
    }
}

const FORMULA_ERROR_CODES: readonly FormulaErrorCode[] = ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#CYCLE!"];

function fail(code: ConditionalFormatErrorCode, message: string, path: string): never {
    throw new ConditionalFormatError(code, message, path);
}

function isObject(value: unknown): value is Record<string, unknown> {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}

function validateRange(range: ConditionalFormatRange, path: string): ConditionalFormatRange {
    if (!isObject(range) || !Number.isInteger(range.x) || !Number.isInteger(range.y) || !Number.isInteger(range.width) || !Number.isInteger(range.height) || range.x < 0 || range.y < 0 || range.width <= 0 || range.height <= 0) {
        fail("invalid-range", "Conditional-format range must use non-negative coordinates and positive integer dimensions", path);
    }
    return { x: range.x, y: range.y, width: range.width, height: range.height };
}

function validateStyle(style: ConditionalFormatStyle, path: string): ConditionalFormatStyle {
    if (!isObject(style)) fail("invalid-rule", "Conditional-format style must be an object", path);
    const result: ConditionalFormatStyle = {};
    for (const key of ["bg", "text", "font", "numberFormat"] as const) {
        const value = style[key];
        if (value !== undefined && typeof value !== "string") fail("invalid-rule", "Style fields must be strings", `${path}.${key}`);
        if (value !== undefined) (result as Record<string, string>)[key] = value;
    }
    for (const key of Object.keys(style)) if (!["bg", "text", "font", "numberFormat"].includes(key)) fail("invalid-rule", "Unknown style property", `${path}.${key}`);
    return result;
}

function rangeKey(range: ConditionalFormatRange): string {
    return `${range.x}:${range.y}:${range.width}:${range.height}`;
}

function valueKey(value: FormulaValue): string {
    if (typeof value === "object" && value !== null && "kind" in value && value.kind === "error") return `error:${value.code}`;
    if (value === null) return "null";
    return `${typeof value}:${String(value)}`;
}

function inRange(range: ConditionalFormatRange, col: number, row: number): boolean {
    return col >= range.x && col < range.x + range.width && row >= range.y && row < range.y + range.height;
}

function validateRule(rule: ConditionalFormatRule, index: number): ConditionalFormatRule {
    const path = `$.rules[${index}]`;
    if (!isObject(rule) || typeof rule.id !== "string" || rule.id.length === 0) fail("invalid-rule", "Rule requires a non-empty id", `${path}.id`);
    const range = validateRange(rule.range, `${path}.range`);
    const style = validateStyle(rule.style, `${path}.style`);
    if (rule.stopIfTrue !== undefined && typeof rule.stopIfTrue !== "boolean") fail("invalid-rule", "stopIfTrue must be boolean", `${path}.stopIfTrue`);
    switch (rule.kind) {
        case "number-compare":
            if (!["eq", "neq", "gt", "gte", "lt", "lte"].includes(rule.operator) || typeof rule.value !== "number" || !Number.isFinite(rule.value)) fail("invalid-number", "Number comparison requires a finite value and valid operator", path);
            return { ...rule, range, style };
        case "number-between":
            if (typeof rule.min !== "number" || typeof rule.max !== "number" || !Number.isFinite(rule.min) || !Number.isFinite(rule.max)) fail("invalid-number", "Number range bounds must be finite", path);
            if (rule.min > rule.max) fail("invalid-between", "Number-between min must not exceed max", path);
            return { ...rule, range, style };
        case "text":
            if ((rule.operator !== "contains" && rule.operator !== "equals") || typeof rule.value !== "string") fail("invalid-rule", "Text rule requires a valid operator and string value", path);
            if (rule.caseSensitive !== undefined && typeof rule.caseSensitive !== "boolean") fail("invalid-rule", "caseSensitive must be boolean", path);
            return { ...rule, range, style };
        case "empty":
        case "not-empty":
        case "duplicate":
        case "unique":
            return { ...rule, range, style };
        case "formula-error":
            if (rule.code !== undefined && (typeof rule.code !== "string" || !FORMULA_ERROR_CODES.includes(rule.code as FormulaErrorCode))) fail("invalid-rule", "Formula-error code is invalid", path);
            return { ...rule, range, style };
        case "custom":
            if (typeof rule.predicate !== "function") fail("invalid-rule", "Custom rule requires a predicate", path);
            return { ...rule, range, style };
        default:
            fail("invalid-rule", "Unknown conditional-format rule kind", path);
    }
}

interface CompiledConditionalRule {
    readonly rule: ConditionalFormatRule;
    duplicateCounts?: ReadonlyMap<string, number>;
}

export interface ConditionalFormatEngine {
    evaluate: (col: number, row: number, value: FormulaValue, context?: Pick<ConditionalFormatContext, "getCellValue">) => ConditionalFormatResult;
    refreshIndexes: () => void;
    readonly rules: readonly ConditionalFormatRule[];
}

function buildDuplicateCounts(range: ConditionalFormatRange, getCellValue: CellValueAccessor): ReadonlyMap<string, number> {
    const counts = new Map<string, number>();
    for (let row = range.y; row < range.y + range.height; row++) {
        for (let col = range.x; col < range.x + range.width; col++) {
            const key = valueKey(getCellValue(col, row));
            counts.set(key, (counts.get(key) ?? 0) + 1);
        }
    }
    return counts;
}

function matches(compiled: CompiledConditionalRule, col: number, row: number, value: FormulaValue, context: ConditionalFormatContext, diagnostics: ConditionalFormatDiagnostic[]): boolean {
    const rule = compiled.rule;
    if (!inRange(rule.range, col, row)) return false;
    switch (rule.kind) {
        case "number-compare":
            if (typeof value !== "number" || !Number.isFinite(value)) return false;
            if (rule.operator === "eq") return value === rule.value;
            if (rule.operator === "neq") return value !== rule.value;
            if (rule.operator === "gt") return value > rule.value;
            if (rule.operator === "gte") return value >= rule.value;
            if (rule.operator === "lt") return value < rule.value;
            return value <= rule.value;
        case "number-between": return typeof value === "number" && Number.isFinite(value) && value >= rule.min && value <= rule.max;
        case "text": {
            if (typeof value !== "string") return false;
            const actual = rule.caseSensitive === true ? value : value.toLowerCase();
            const expected = rule.caseSensitive === true ? rule.value : rule.value.toLowerCase();
            return rule.operator === "equals" ? actual === expected : actual.includes(expected);
        }
        case "empty": return value === null || value === "";
        case "not-empty": return value !== null && value !== "";
        case "duplicate":
        case "unique": {
            const counts = compiled.duplicateCounts;
            if (counts === undefined) {
                diagnostics.push({ ruleId: rule.id, code: "missing-value-accessor", message: "Duplicate/unique rule requires a cell-value accessor" });
                return false;
            }
            const count = counts.get(valueKey(value)) ?? 0;
            return rule.kind === "duplicate" ? count > 1 : count === 1;
        }
        case "formula-error":
            return typeof value === "object" && value !== null && "kind" in value && value.kind === "error" && (rule.code === undefined || value.code === rule.code);
        case "custom":
            try {
                return rule.predicate(context);
            } catch {
                diagnostics.push({ ruleId: rule.id, code: "custom-error", message: "Custom conditional-format predicate threw an exception" });
                return false;
            }
    }
}

/**
 * Compiles rules once; duplicate/unique ranges are indexed once and reused by
 * evaluations. The rules array order is priority order: later matching rules
 * override earlier style fields, unless an earlier rule has stopIfTrue.
 * Per-call accessors may override the compile accessor; changing accessor
 * identity rebuilds duplicate indexes so results never use stale values.
 */
export function compileConditionalFormatting(rules: readonly ConditionalFormatRule[], options: ConditionalFormatCompileOptions | CellValueAccessor = {}): ConditionalFormatEngine {
    if (!Array.isArray(rules)) fail("invalid-rule", "Rules must be an array", "$.rules");
    const accessor = typeof options === "function" ? options : options.getCellValue;
    const seenIds = new Set<string>();
    const normalized = rules.map((rule, index) => {
        const copy = validateRule(rule, index);
        if (seenIds.has(copy.id)) fail("duplicate-rule-id", "Rule ids must be unique", `$.rules[${index}].id`);
        seenIds.add(copy.id);
        return copy;
    });
    const indexCache = new Map<string, ReadonlyMap<string, number> | undefined>();
    const compiled: CompiledConditionalRule[] = normalized.map(rule => {
        if (rule.kind !== "duplicate" && rule.kind !== "unique") return { rule };
        const key = rangeKey(rule.range);
        if (!indexCache.has(key)) indexCache.set(key, accessor === undefined ? undefined : buildDuplicateCounts(rule.range, accessor));
        return { rule, duplicateCounts: indexCache.get(key) };
    });
    let indexedAccessor = accessor;
    const ensureIndexes = (valueAccessor: CellValueAccessor): void => {
        if (indexedAccessor !== valueAccessor) {
            indexCache.clear();
            for (const entry of compiled) {
                if (entry.rule.kind === "duplicate" || entry.rule.kind === "unique") entry.duplicateCounts = undefined;
            }
            indexedAccessor = valueAccessor;
        }
        for (const entry of compiled) {
            if (entry.rule.kind !== "duplicate" && entry.rule.kind !== "unique") continue;
            const key = rangeKey(entry.rule.range);
            if (!indexCache.has(key) || indexCache.get(key) === undefined) indexCache.set(key, buildDuplicateCounts(entry.rule.range, valueAccessor));
            entry.duplicateCounts = indexCache.get(key);
        }
    };
    const refreshIndexes = () => {
        indexCache.clear();
        for (const entry of compiled) {
            if (entry.rule.kind === "duplicate" || entry.rule.kind === "unique") entry.duplicateCounts = undefined;
        }
        if (accessor !== undefined) ensureIndexes(accessor);
    };
    return {
        rules: normalized.map(rule => ({ ...rule, range: { ...rule.range }, style: { ...rule.style } })),
        refreshIndexes,
        evaluate: (col, row, value, context = {}) => {
            const style: ConditionalFormatStyle = {};
            const matchedRuleIds: string[] = [];
            const diagnostics: ConditionalFormatDiagnostic[] = [];
            // A fixed compile accessor is used when no per-call accessor is
            // supplied. For a lazy engine, retain the most recently supplied
            // accessor so repeated evaluations can reuse its index.
            const getCellValue = context.getCellValue ?? accessor ?? indexedAccessor;
            if (getCellValue !== undefined) ensureIndexes(getCellValue);
            for (const entry of compiled) {
                const ruleContext: ConditionalFormatContext = { col, row, value, ...(getCellValue === undefined ? {} : { getCellValue }) };
                if (!matches(entry, col, row, value, ruleContext, diagnostics)) continue;
                matchedRuleIds.push(entry.rule.id);
                Object.assign(style, entry.rule.style);
                if (entry.rule.stopIfTrue === true) break;
            }
            return { style, matchedRuleIds, diagnostics };
        },
    };
}

/** Convenience one-shot evaluation; compileConditionalFormatting is preferred for many cells. */
export function evaluateConditionalFormatting(col: number, row: number, value: FormulaValue, rules: readonly ConditionalFormatRule[], options: ConditionalFormatCompileOptions | CellValueAccessor = {}): ConditionalFormatResult {
    return compileConditionalFormatting(rules, options).evaluate(col, row, value);
}
