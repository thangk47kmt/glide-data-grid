import type { CellInput, SpreadsheetModel } from "./model.js";
import type { HistoryTransaction } from "./history.js";

/** The cell representation searched by a find operation. */
export type FindScope = "raw" | "computed";

/** Inclusive zero-based row bounds for a find operation. */
export interface FindRowRange {
    readonly start: number;
    readonly end: number;
}

/** Search options shared by find and replacement operations. */
export interface FindOptions {
    readonly query: string;
    readonly scope?: FindScope;
    readonly caseSensitive?: boolean;
    readonly wholeCell?: boolean;
    /** Enables regular expressions. Plain text is the default. */
    readonly regexp?: boolean;
    /** Alias accepted for callers that prefer the expanded spelling. */
    readonly useRegExp?: boolean;
    /** Optional selected zero-based columns; omitted means all columns. Invalid columns are ignored. */
    readonly columns?: readonly number[];
    readonly rowRange?: FindRowRange;
    readonly maxResults?: number;
    readonly isProtected?: (col: number, row: number, raw: CellInput, display: string) => boolean;
}

/** A deterministic row-major search result, including both raw and display values. */
export interface FindMatch {
    readonly col: number;
    readonly row: number;
    readonly raw: CellInput;
    readonly display: string;
}

/** A non-throwing validation/read-only diagnostic. */
export interface FindDiagnostic {
    readonly code: "INVALID_REGEX" | "INVALID_OPTIONS" | "COMPUTED_SCOPE_READ_ONLY";
    readonly message: string;
}

/** Search results. `truncated` is true when the result cap stopped scanning. */
export interface FindResult {
    readonly matches: readonly FindMatch[];
    readonly truncated: boolean;
    readonly diagnostic?: FindDiagnostic;
}

/** Find options plus a raw-input replacement and optional type coercion hook. */
export interface ReplaceOptions extends FindOptions {
    readonly replacement: string;
    /**
     * Optional parser for typing replacement text back into a CellInput. The
     * default is the replacement string, including for numeric/boolean raw
     * values. A thrown parser error skips that cell and is reported.
     */
    readonly coerceReplacement?: (text: string, match: FindMatch) => CellInput;
    readonly transactionId?: string;
    readonly label?: string;
}

/** A non-mutating replacement plan suitable for `History.execute`/`record`. */
export interface ReplaceResult {
    readonly matches: readonly FindMatch[];
    readonly truncated: boolean;
    readonly replacedCount: number;
    readonly skippedCount: number;
    readonly skippedReasons?: readonly string[];
    readonly transaction?: HistoryTransaction<CellInput>;
    readonly diagnostic?: FindDiagnostic;
}

const DEFAULT_MAX_RESULTS = 1_000;
type MatcherOptions = FindOptions & { readonly replacement?: string };

function diagnostic(code: FindDiagnostic["code"], message: string): FindDiagnostic {
    return { code, message };
}

function rawText(value: CellInput): string {
    if (value === null) return "";
    if (typeof value === "object" && value !== null && "kind" in value && value.kind === "error") return value.code;
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    return String(value);
}

function resolveColumns(model: SpreadsheetModel, options: FindOptions): number[] {
    const selected = options.columns ?? Array.from({ length: model.columns.length }, (_, col) => col);
    return [...new Set(selected)].filter(col => Number.isInteger(col) && col >= 0 && col < model.columns.length).sort((left, right) => left - right);
}

function resolveRows(model: SpreadsheetModel, options: FindOptions): { start: number; end: number } | FindDiagnostic {
    if (options.rowRange === undefined) return { start: 0, end: model.rowCount - 1 };
    const { start, end } = options.rowRange;
    if (!Number.isInteger(start) || !Number.isInteger(end)) return diagnostic("INVALID_OPTIONS", "rowRange bounds must be integers");
    return { start: Math.max(0, start), end: Math.min(model.rowCount - 1, end) };
}

function maxResults(options: FindOptions): number | FindDiagnostic {
    const max = options.maxResults ?? DEFAULT_MAX_RESULTS;
    return Number.isInteger(max) && max >= 0 ? max : diagnostic("INVALID_OPTIONS", "maxResults must be a non-negative integer");
}

function useRegExp(options: FindOptions): boolean {
    return options.regexp === true || options.useRegExp === true;
}

function createMatcher(options: MatcherOptions): { test: (value: string) => boolean; replace: (value: string) => string } | FindDiagnostic {
    const flags = options.caseSensitive === true ? "g" : "gi";
    if (useRegExp(options)) {
        try {
            const expression = options.wholeCell === true ? `^(?:${options.query})$` : options.query;
            const matcher = new RegExp(expression, flags);
            return {
                test: value => {
                    matcher.lastIndex = 0;
                    return matcher.test(value);
                },
                // Regex replacements use JavaScript tokens: $& (whole match),
                // $1..$99 (captures), and $$ (literal dollar sign).
                replace: value => value.replace(matcher, options.replacement ?? ""),
            };
        } catch (error) {
            return diagnostic("INVALID_REGEX", error instanceof Error ? error.message : "Invalid regular expression");
        }
    }

    const wanted = options.caseSensitive === true ? options.query : options.query.toLocaleLowerCase("en-US");
    const equal = (value: string): boolean => (options.caseSensitive === true ? value : value.toLocaleLowerCase("en-US")) === wanted;
    if (options.wholeCell === true) return { test: equal, replace: () => options.replacement ?? "" };
    return {
        test: value => wanted === "" || (options.caseSensitive === true ? value : value.toLocaleLowerCase("en-US")).includes(wanted),
        // Plain-text replacement is literal: `$1`, `$&`, and `$$` remain typed.
        replace: value => {
            if (options.query === "") return value;
            const escaped = options.query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            try {
                return value.replace(new RegExp(escaped, options.caseSensitive === true ? "g" : "gi"), () => options.replacement ?? "");
            } catch {
                return value;
            }
        },
    };
}

function searchValue(model: SpreadsheetModel, col: number, row: number, scope: FindScope): { raw: CellInput; display: string; text: string } {
    const raw = model.getInput(col, row);
    const display = model.getDisplayValue(col, row);
    return { raw, display, text: scope === "computed" ? display : rawText(raw) };
}

/**
 * Finds raw or computed values in row-major order without materializing an
 * index. An empty query and malformed options return a diagnostic instead of
 * matching every cell.
 */
export function findMatches(model: SpreadsheetModel, options: FindOptions): FindResult {
    try {
        if (typeof options.query !== "string") return { matches: [], truncated: false, diagnostic: diagnostic("INVALID_OPTIONS", "query must be a string") };
        if (options.query.length === 0) return { matches: [], truncated: false, diagnostic: diagnostic("INVALID_OPTIONS", "query must not be empty") };
        if (options.scope !== undefined && options.scope !== "raw" && options.scope !== "computed") return { matches: [], truncated: false, diagnostic: diagnostic("INVALID_OPTIONS", "scope must be 'raw' or 'computed'") };
        const matcher = createMatcher(options);
        if ("code" in matcher) return { matches: [], truncated: false, diagnostic: matcher };
        const limit = maxResults(options);
        if (typeof limit !== "number") return { matches: [], truncated: false, diagnostic: limit };
        const rows = resolveRows(model, options);
        if ("code" in rows) return { matches: [], truncated: false, diagnostic: rows };
        const matches: FindMatch[] = [];
        if (rows.start > rows.end || limit === 0 && model.rowCount === 0) return { matches, truncated: false };
        const columns = resolveColumns(model, options);
        const scope = options.scope ?? "raw";
        for (let row = rows.start; row <= rows.end; row++) {
            for (const col of columns) {
                const value = searchValue(model, col, row, scope);
                if (options.isProtected?.(col, row, value.raw, value.display) === true || !matcher.test(value.text)) continue;
                if (matches.length >= limit) return { matches, truncated: true };
                matches.push({ col, row, raw: value.raw, display: value.display });
            }
        }
        return { matches, truncated: false };
    } catch (error) {
        return { matches: [], truncated: false, diagnostic: diagnostic("INVALID_OPTIONS", error instanceof Error ? error.message : String(error)) };
    }
}

function replace(model: SpreadsheetModel, options: ReplaceOptions, one: boolean): ReplaceResult {
    const emptyResult = (diagnosticValue?: FindDiagnostic): ReplaceResult => ({ matches: [], truncated: false, replacedCount: 0, skippedCount: 0, ...(diagnosticValue === undefined ? {} : { diagnostic: diagnosticValue }) });
    if (typeof options.query !== "string" || options.query.length === 0) return emptyResult(diagnostic("INVALID_OPTIONS", "query must not be empty"));
    if (typeof options.replacement !== "string") return emptyResult(diagnostic("INVALID_OPTIONS", "replacement must be a string"));
    if (options.scope !== undefined && options.scope !== "raw" && options.scope !== "computed") return emptyResult(diagnostic("INVALID_OPTIONS", "scope must be 'raw' or 'computed'"));
    if (options.coerceReplacement !== undefined && typeof options.coerceReplacement !== "function") return emptyResult(diagnostic("INVALID_OPTIONS", "coerceReplacement must be a function"));
    if ((options.scope ?? "raw") !== "raw") return emptyResult(diagnostic("COMPUTED_SCOPE_READ_ONLY", "Replace operates on raw inputs only"));
    const result = findMatches(model, { ...options, maxResults: one ? 1 : options.maxResults });
    if (result.diagnostic !== undefined) return emptyResult(result.diagnostic);
    const matcher = createMatcher(options);
    if ("code" in matcher) return emptyResult(matcher);
    const skippedReasons: string[] = [];
    const edits = result.matches.flatMap(match => {
        // The default intentionally returns replacement text for every raw
        // value, including numbers/booleans. Callers that need typed values
        // can provide coerceReplacement (for example, Number(text)).
        const afterText = matcher.replace(rawText(match.raw));
        let after: CellInput;
        try {
            after = options.coerceReplacement?.(afterText, match) ?? afterText;
        } catch (error) {
            skippedReasons.push(error instanceof Error ? error.message : "coerceReplacement failed");
            return [];
        }
        return Object.is(after, match.raw) ? [] : [{ location: [match.col, match.row] as const, before: match.raw, after }];
    });
    if (edits.length === 0) return { matches: result.matches, truncated: result.truncated, replacedCount: 0, skippedCount: skippedReasons.length, ...(skippedReasons.length === 0 ? {} : { skippedReasons }) };
    return {
        matches: result.matches,
        truncated: result.truncated,
        replacedCount: edits.length,
        skippedCount: skippedReasons.length,
        ...(skippedReasons.length === 0 ? {} : { skippedReasons }),
        transaction: {
            id: options.transactionId ?? (one ? "find-replace-one" : "find-replace-all"),
            label: options.label ?? (one ? "Replace one" : "Replace all"),
            edits,
        },
    };
}

/** Builds one raw-input replacement transaction; does not mutate the model. */
export function replaceOne(model: SpreadsheetModel, options: ReplaceOptions): ReplaceResult {
    return replace(model, options, true);
}

/**
 * Builds one replacement transaction for all matching raw-input cells; does
 * not mutate the model. The default cap is 1,000 matches; callers must check
 * `truncated` before treating the transaction as a complete workbook update.
 */
export function replaceAll(model: SpreadsheetModel, options: ReplaceOptions): ReplaceResult {
    return replace(model, options, false);
}
