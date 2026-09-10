import { isFormulaError, type FormulaValue } from "./formula.js";

export interface FractionDigitOptions {
    readonly minimumFractionDigits?: number;
    readonly maximumFractionDigits?: number;
    readonly useGrouping?: boolean;
}

export type CellFormat =
    | { readonly kind: "general" }
    | ({ readonly kind: "number" } & FractionDigitOptions)
    | ({ readonly kind: "percent" } & FractionDigitOptions)
    | ({ readonly kind: "currency"; readonly currency: string; readonly currencyDisplay?: "symbol" | "code" | "name" } & FractionDigitOptions)
    | { readonly kind: "date"; readonly preset?: "short" | "medium" | "long" }
    | { readonly kind: "text" };

const DATE_EPOCH_MS = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;
const FORMAT_CACHE_LIMIT = 100;

const numberFormatterCache = new Map<string, Intl.NumberFormat>();
const dateFormatterCache = new Map<string, Intl.DateTimeFormat>();

export interface FormatCacheStats {
    readonly number: number;
    readonly date: number;
}

/** Clears cached Intl formatters, primarily useful for tests or long-lived hosts. */
export function clearFormatCache(): void {
    numberFormatterCache.clear();
    dateFormatterCache.clear();
}

/** Returns cache sizes so hosts can verify formatter reuse and bounded growth. */
export function getFormatCacheStats(): FormatCacheStats {
    return { number: numberFormatterCache.size, date: dateFormatterCache.size };
}

function getCached<T>(cache: Map<string, T>, key: string, create: () => T): T | undefined {
    const existing = cache.get(key);
    if (existing !== undefined) return existing;
    let created: T;
    try {
        created = create();
    } catch {
        return undefined;
    }
    if (cache.size >= FORMAT_CACHE_LIMIT) {
        const oldest = cache.keys().next().value;
        if (typeof oldest === "string") cache.delete(oldest);
    }
    cache.set(key, created);
    return created;
}

function primitiveText(value: FormulaValue): string {
    if (value === null) return "";
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    return String(value);
}

function validFractionDigits(value: number | undefined): boolean {
    return value === undefined || (Number.isInteger(value) && value >= 0 && value <= 20);
}

function validFractionOptions(format: FractionDigitOptions): boolean {
    return validFractionDigits(format.minimumFractionDigits) &&
        validFractionDigits(format.maximumFractionDigits) &&
        (format.minimumFractionDigits === undefined || format.maximumFractionDigits === undefined || format.minimumFractionDigits <= format.maximumFractionDigits);
}

function numberFormatOptions(format: Extract<CellFormat, { readonly kind: "number" | "percent" | "currency" }>): Intl.NumberFormatOptions | undefined {
    if (!validFractionOptions(format)) return undefined;
    const options: Intl.NumberFormatOptions = {
        style: format.kind === "percent" ? "percent" : format.kind === "currency" ? "currency" : "decimal",
        minimumFractionDigits: format.minimumFractionDigits,
        maximumFractionDigits: format.maximumFractionDigits,
        useGrouping: format.useGrouping,
    };
    if (format.kind === "currency") {
        if (!/^[A-Za-z]{3}$/.test(format.currency)) return undefined;
        options.currency = format.currency.toUpperCase();
        options.currencyDisplay = format.currencyDisplay;
    }
    return options;
}

function formatNumber(value: number, format: Extract<CellFormat, { readonly kind: "number" | "percent" | "currency" }>, locale: string): string {
    if (!Number.isFinite(value)) return String(value);
    const options = numberFormatOptions(format);
    if (options === undefined) return String(value);
    const key = JSON.stringify([locale, options]);
    const formatter = getCached(numberFormatterCache, key, () => new Intl.NumberFormat(locale, options));
    if (formatter === undefined) return String(value);
    try {
        return formatter.format(value);
    } catch {
        return String(value);
    }
}

function formatDate(value: number, format: Extract<CellFormat, { readonly kind: "date" }>, locale: string): string {
    if (!Number.isFinite(value)) return String(value);
    const date = new Date(DATE_EPOCH_MS + value * DAY_MS);
    if (Number.isNaN(date.getTime())) return String(value);
    const preset = format.preset ?? "short";
    const key = JSON.stringify([locale, preset]);
    const formatter = getCached(dateFormatterCache, key, () => new Intl.DateTimeFormat(locale, { dateStyle: preset, timeZone: "UTC" }));
    if (formatter === undefined) return String(value);
    try {
        return formatter.format(date);
    } catch {
        return String(value);
    }
}

/** Formats one spreadsheet value without throwing from a rendering hot path. */
export function formatCellValue(value: FormulaValue, format: CellFormat = { kind: "general" }, locale = "en-US"): string {
    if (isFormulaError(value)) return value.code;
    if (value === null) return "";
    if (format.kind === "general" || format.kind === "text") return primitiveText(value);
    if (format.kind === "number" || format.kind === "percent" || format.kind === "currency") {
        return typeof value === "number" ? formatNumber(value, format, locale) : primitiveText(value);
    }
    if (format.kind === "date") return typeof value === "number" ? formatDate(value, format, locale) : primitiveText(value);
    return primitiveText(value);
}
