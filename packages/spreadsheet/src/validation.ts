import type { CellInput } from "./model.js";

export type ValidationSeverity = "error" | "warning";

export interface ValidationIssue {
    readonly ruleId: string;
    readonly message: string;
    readonly severity: ValidationSeverity;
}

export interface ValidationResult {
    readonly valid: boolean;
    readonly errors: readonly ValidationIssue[];
    readonly warnings: readonly ValidationIssue[];
}

export interface ValidationContext {
    /** Formula handling: validate its input, defer it, or validate computedValue. */
    readonly formulaPolicy?: "raw" | "skip" | "computed";
    readonly computedValue?: CellInput;
    /** Defaults to collect-all. `collectAll` is an explicit inverse alias. */
    readonly stopOnFirst?: boolean;
    readonly collectAll?: boolean;
    readonly [key: string]: unknown;
}

interface ValidationRuleBase {
    readonly id: string;
    readonly message?: string;
    readonly severity?: ValidationSeverity;
}

export interface RequiredRule extends ValidationRuleBase {
    readonly kind: "required";
}

export interface NumberRangeRule extends ValidationRuleBase {
    readonly kind: "number-range";
    readonly min?: number;
    readonly max?: number;
}

export interface TextLengthRule extends ValidationRuleBase {
    readonly kind: "text-length";
    readonly min?: number;
    readonly max?: number;
}

export interface OneOfRule extends ValidationRuleBase {
    readonly kind: "one-of";
    readonly values: readonly CellInput[];
    readonly caseSensitive?: boolean;
}

export interface RegexRule extends ValidationRuleBase {
    readonly kind: "regex";
    /** A precompiled expression; arbitrary expression strings are not accepted. */
    readonly pattern: RegExp;
}

export interface DateIsoRangeRule extends ValidationRuleBase {
    readonly kind: "date-iso-range";
    readonly min?: string;
    readonly max?: string;
}

export type CustomValidationPredicate = (value: CellInput, context: ValidationContext) => boolean | string;

export interface CustomRule extends ValidationRuleBase {
    readonly kind: "custom";
    readonly validate: CustomValidationPredicate;
}

export type ValidationRule = RequiredRule | NumberRangeRule | TextLengthRule | OneOfRule | RegexRule | DateIsoRangeRule | CustomRule;

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;
const DAY_MS = 86_400_000;

function isoDay(value: string): number | undefined {
    const match = value.match(ISO_DATE);
    if (match === null) return undefined;
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    // setUTCFullYear avoids Date.UTC's special 1900 offset for years 0..99.
    const date = new Date(0);
    date.setUTCHours(0, 0, 0, 0);
    date.setUTCFullYear(year, month - 1, day);
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return undefined;
    return Math.floor(date.getTime() / DAY_MS);
}

function sameValue(left: CellInput, right: CellInput, caseSensitive: boolean): boolean {
    if (typeof left === "string" && typeof right === "string") {
        return caseSensitive ? left === right : left.toLowerCase() === right.toLowerCase();
    }
    return Object.is(left, right);
}

function defaultMessage(rule: ValidationRule): string {
    switch (rule.kind) {
        case "required": return "A value is required";
        case "number-range": return "Value must be a number in the allowed range";
        case "text-length": return "Text length is outside the allowed range";
        case "one-of": return "Value is not an allowed option";
        case "regex": return "Text does not match the required pattern";
        case "date-iso-range": return "Date must be a valid ISO date in the allowed range";
        case "custom": return "Value failed custom validation";
    }
}

function isValidRule(rule: ValidationRule): boolean {
    switch (rule.kind) {
        case "number-range":
            return (rule.min === undefined || Number.isFinite(rule.min)) &&
                (rule.max === undefined || Number.isFinite(rule.max)) &&
                (rule.min === undefined || rule.max === undefined || rule.min <= rule.max);
        case "text-length":
            return (rule.min === undefined || (Number.isInteger(rule.min) && rule.min >= 0)) &&
                (rule.max === undefined || (Number.isInteger(rule.max) && rule.max >= 0)) &&
                (rule.min === undefined || rule.max === undefined || rule.min <= rule.max);
        case "date-iso-range": {
            const min = rule.min === undefined ? undefined : isoDay(rule.min);
            const max = rule.max === undefined ? undefined : isoDay(rule.max);
            return (rule.min === undefined || min !== undefined) &&
                (rule.max === undefined || max !== undefined) &&
                (min === undefined || max === undefined || min <= max);
        }
        default: return true;
    }
}

function passes(rule: ValidationRule, value: CellInput, context: ValidationContext): { readonly pass: boolean; readonly message?: string } {
    switch (rule.kind) {
        case "required":
            return { pass: value !== null && value !== "" };
        case "number-range":
            return {
                pass: typeof value === "number" && Number.isFinite(value) && (rule.min === undefined || value >= rule.min) && (rule.max === undefined || value <= rule.max),
            };
        case "text-length":
            return {
                pass: typeof value === "string" && (rule.min === undefined || value.length >= rule.min) && (rule.max === undefined || value.length <= rule.max),
            };
        case "one-of":
            return { pass: rule.values.some(option => sameValue(value, option, rule.caseSensitive ?? true)) };
        case "regex": {
            if (typeof value !== "string") return { pass: false };
            try {
                // Clone so testing a global/sticky rule does not mutate the supplied rule's lastIndex.
                return { pass: new RegExp(rule.pattern.source, rule.pattern.flags).test(value) };
            } catch {
                return { pass: false };
            }
        }
        case "date-iso-range": {
            if (typeof value !== "string") return { pass: false };
            const day = isoDay(value);
            const min = rule.min === undefined ? undefined : isoDay(rule.min);
            const max = rule.max === undefined ? undefined : isoDay(rule.max);
            return { pass: day !== undefined && (min === undefined || day >= min) && (max === undefined || day <= max) };
        }
        case "custom": {
            try {
                const result = rule.validate(value, context);
                return typeof result === "string" ? { pass: result.length === 0, message: result || undefined } : { pass: result === true };
            } catch {
                return { pass: false, message: "Custom validation failed" };
            }
        }
    }
}

function issueFor(rule: ValidationRule, result: { readonly message?: string }, invalidRule = false): ValidationIssue {
    return {
        ruleId: rule.id,
        message: invalidRule ? "Validation rule configuration is invalid" : rule.message ?? result.message ?? defaultMessage(rule),
        severity: rule.severity === "warning" ? "warning" : "error",
    };
}

/** Validates a cell input without throwing for malformed user data or predicate failures. */
export function validateCellInput(input: CellInput, rules: readonly ValidationRule[], context: ValidationContext = {}): ValidationResult {
    const hasComputedValue = Object.prototype.hasOwnProperty.call(context, "computedValue");
    const formulaPolicy = context.formulaPolicy ?? (hasComputedValue ? "computed" : "skip");
    const isFormula = typeof input === "string" && input.startsWith("=");
    if (isFormula && formulaPolicy === "skip") return { valid: true, errors: [], warnings: [] };
    if (isFormula && formulaPolicy === "computed" && !hasComputedValue) return { valid: true, errors: [], warnings: [] };
    const value = isFormula && formulaPolicy === "computed" ? context.computedValue as CellInput : input;
    const errors: ValidationIssue[] = [];
    const warnings: ValidationIssue[] = [];
    const stopOnFirst = context.stopOnFirst === true || context.collectAll === false;
    for (const rule of rules) {
        const validConfiguration = isValidRule(rule);
        const outcome = validConfiguration ? passes(rule, value, context) : { pass: false };
        if (outcome.pass) continue;
        const issue = issueFor(rule, outcome, !validConfiguration);
        (issue.severity === "warning" ? warnings : errors).push(issue);
        if (stopOnFirst) break;
    }
    return { valid: errors.length === 0, errors, warnings };
}
