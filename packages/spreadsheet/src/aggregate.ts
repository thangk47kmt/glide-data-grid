import { isFormulaError, type FormulaError, type FormulaValue } from "./formula.js";

export type AggregateErrorPolicy = "ignore" | "include-first-error";
export type NumericStringPolicy = "strict" | "coerce";
export type BooleanPolicy = "ignore" | "numeric";

export interface AggregateOptions {
    /** Formula errors are counted either way; include-first-error propagates the first one to metrics. */
    readonly errorPolicy?: AggregateErrorPolicy;
    /** Strict (the default) ignores numeric-looking strings in numeric metrics. */
    readonly numericStringPolicy?: NumericStringPolicy;
    /** Ignore booleans by default, or treat TRUE/FALSE as 1/0. */
    readonly booleanPolicy?: BooleanPolicy;
    /** Optional row/value visibility predicate used by aggregateValues. */
    readonly isVisible?: (value: FormulaValue, index: number) => boolean;
}

export type AggregateMetric = number | null | FormulaError;

export interface AggregateResult {
    readonly countAll: number;
    readonly countNumbers: number;
    readonly sum: AggregateMetric;
    readonly average: AggregateMetric;
    readonly min: AggregateMetric;
    readonly max: AggregateMetric;
    readonly errors: number;
}

function isEmpty(value: FormulaValue): boolean {
    return value === null || value === "";
}

function numericValue(value: FormulaValue, options: Required<Pick<AggregateOptions, "numericStringPolicy" | "booleanPolicy">>): number | undefined {
    if (typeof value === "number") return Number.isFinite(value) ? value : undefined;
    if (typeof value === "string") {
        if (options.numericStringPolicy !== "coerce" || value.trim() === "") return undefined;
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : undefined;
    }
    if (typeof value === "boolean" && options.booleanPolicy === "numeric") return value ? 1 : 0;
    return undefined;
}

/**
 * A small streaming accumulator. Numeric totals use Neumaier compensation,
 * retaining the precision benefit of Kahan summation while making merge cheap.
 */
export class AggregateAccumulator {
    private readonly errorPolicy: AggregateErrorPolicy;
    private readonly numericStringPolicy: NumericStringPolicy;
    private readonly booleanPolicy: BooleanPolicy;
    private countAllValue = 0;
    private countNumbersValue = 0;
    private sumValue = 0;
    private compensation = 0;
    private minValue: number | undefined;
    private maxValue: number | undefined;
    private errorsValue = 0;
    private firstError: FormulaError | undefined;

    public constructor(options: AggregateOptions = {}) {
        this.errorPolicy = options.errorPolicy ?? "ignore";
        this.numericStringPolicy = options.numericStringPolicy ?? "strict";
        this.booleanPolicy = options.booleanPolicy ?? "ignore";
    }

    public add(value: FormulaValue): this {
        if (!isEmpty(value)) this.countAllValue++;
        if (isFormulaError(value)) {
            this.errorsValue++;
            if (this.firstError === undefined) this.firstError = value;
            return this;
        }
        const number = numericValue(value, {
            numericStringPolicy: this.numericStringPolicy,
            booleanPolicy: this.booleanPolicy,
        });
        if (number === undefined) return this;
        this.countNumbersValue++;
        this.addNumber(number);
        this.minValue = this.minValue === undefined ? number : Math.min(this.minValue, number);
        this.maxValue = this.maxValue === undefined ? number : Math.max(this.maxValue, number);
        return this;
    }

    public merge(other: AggregateAccumulator): this {
        if (this.errorPolicy !== other.errorPolicy ||
            this.numericStringPolicy !== other.numericStringPolicy ||
            this.booleanPolicy !== other.booleanPolicy) {
            throw new TypeError("Cannot merge aggregate accumulators with incompatible policies");
        }
        this.countAllValue += other.countAllValue;
        this.countNumbersValue += other.countNumbersValue;
        this.errorsValue += other.errorsValue;
        if (this.firstError === undefined) this.firstError = other.firstError;
        if (other.countNumbersValue > 0) this.addNumber(other.total());
        if (other.minValue !== undefined) this.minValue = this.minValue === undefined ? other.minValue : Math.min(this.minValue, other.minValue);
        if (other.maxValue !== undefined) this.maxValue = this.maxValue === undefined ? other.maxValue : Math.max(this.maxValue, other.maxValue);
        return this;
    }

    public finalize(): AggregateResult {
        const error = this.errorPolicy === "include-first-error" ? this.firstError : undefined;
        const total = this.total();
        return {
            countAll: this.countAllValue,
            countNumbers: this.countNumbersValue,
            sum: error ?? (this.countNumbersValue === 0 ? null : total),
            average: error ?? (this.countNumbersValue === 0 ? null : total / this.countNumbersValue),
            min: error ?? (this.minValue ?? null),
            max: error ?? (this.maxValue ?? null),
            errors: this.errorsValue,
        };
    }

    private addNumber(value: number): void {
        const next = this.sumValue + value;
        this.compensation += Math.abs(this.sumValue) >= Math.abs(value)
            ? (this.sumValue - next) + value
            : (value - next) + this.sumValue;
        this.sumValue = next;
    }

    private total(): number {
        return this.sumValue + this.compensation;
    }
}

export function createAggregateAccumulator(options: AggregateOptions = {}): AggregateAccumulator {
    return new AggregateAccumulator(options);
}

export function aggregateValues(values: readonly FormulaValue[], options: AggregateOptions = {}): AggregateResult {
    const accumulator = new AggregateAccumulator(options);
    const isVisible = options.isVisible ?? (() => true);
    values.forEach((value, index) => {
        if (isVisible(value, index)) accumulator.add(value);
    });
    return accumulator.finalize();
}

/** Convenience adapter for callers that keep visibility separate from formatting/value arrays. */
export function aggregateVisibleValues(
    values: readonly FormulaValue[],
    isVisible: (value: FormulaValue, index: number) => boolean,
    options: Omit<AggregateOptions, "isVisible"> = {},
): AggregateResult {
    return aggregateValues(values, { ...options, isVisible });
}
