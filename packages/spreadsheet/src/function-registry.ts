import type { FormulaValue } from "./formula.js";

/** A range passed to a custom function, kept distinct from a scalar value. */
export interface FormulaRangeArgument {
    readonly kind: "range";
    readonly values: readonly FormulaValue[];
    /** Number of rows in the original range (available for shape-aware functions). */
    readonly rows?: number;
    /** Number of columns in the original range (available for shape-aware functions). */
    readonly columns?: number;
}

export type FormulaFunctionArgument = FormulaValue | FormulaRangeArgument;

/** Read-only metadata and cell access supplied while evaluating a function. */
export interface FormulaFunctionContext {
    readonly currentRow: number;
    readonly rowCount: number;
    readonly resolveColumn: (name: string) => number | undefined;
    readonly getCellValue: (col: number, row: number) => FormulaValue;
}

export type FormulaFunction = (
    args: readonly FormulaFunctionArgument[],
    context: FormulaFunctionContext
) => FormulaValue;

export interface FormulaFunctionRegistrationOptions {
    /** Required when intentionally replacing a built-in function. */
    readonly overrideBuiltIn?: boolean;
}

// Kept here (rather than importing the evaluator) so the registry remains a
// standalone, model-independent API.
export const BUILT_IN_FUNCTION_NAMES = [
    "ABS",
    "AND",
    "AVERAGE",
    "CONCAT",
    "COUNTA",
    "COUNT",
    "DATE",
    "DAY",
    "IF",
    "INDEX",
    "LEFT",
    "LEN",
    "MAX",
    "MATCH",
    "MID",
    "MIN",
    "MONTH",
    "NOT",
    "OR",
    "RIGHT",
    "ROUND",
    "SUM",
    "TODAY",
    "TRIM",
    "UPPER",
    "LOWER",
    "YEAR",
    "XLOOKUP",
] as const;

const builtInNames = new Set<string>(BUILT_IN_FUNCTION_NAMES);

function normalizeName(name: string): string {
    if (typeof name !== "string" || !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) {
        throw new Error("Function names must match [A-Za-z_][A-Za-z0-9_.]*");
    }
    // A1 references are valid expression syntax, not function names.
    if (/^\$?[A-Za-z]+\$?\d+$/.test(name)) throw new Error(`Invalid function name '${name}'`);
    return name.toUpperCase();
}

/** Case-insensitive registry for synchronous custom formula functions. */
export class FormulaFunctionRegistry {
    private readonly functions = new Map<string, FormulaFunction>();

    public register(name: string, fn: FormulaFunction, options: FormulaFunctionRegistrationOptions = {}): void {
        const normalized = normalizeName(name);
        if (typeof fn !== "function") throw new TypeError("A formula function must be callable");
        if (this.functions.has(normalized)) throw new Error(`Function '${normalized}' is already registered`);
        if (builtInNames.has(normalized) && options.overrideBuiltIn !== true) {
            throw new Error(`Function '${normalized}' is built in; pass overrideBuiltIn to replace it`);
        }
        this.functions.set(normalized, fn);
    }

    public unregister(name: string): boolean {
        return this.functions.delete(normalizeName(name));
    }

    public has(name: string): boolean {
        return this.functions.has(normalizeName(name));
    }

    public get(name: string): FormulaFunction | undefined {
        return this.functions.get(normalizeName(name));
    }

    public list(): readonly string[] {
        return [...this.functions.keys()].sort();
    }
}
