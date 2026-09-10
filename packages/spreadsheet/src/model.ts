import {
    compileFormula,
    formulaError,
    isFormulaError,
    type CompiledFormula,
    type FormulaValue,
} from "./formula.js";
import type { FormulaFunctionRegistry } from "./function-registry.js";

export type CellInput = FormulaValue;

export interface SpreadsheetColumn {
    /** Stable database/API column name used for persistence and query state. */
    readonly id: string;
    /** User-facing caption shown in headers and accepted in structured formulas. */
    readonly title: string;
    readonly width?: number;
    readonly type?: "text" | "number" | "boolean";
    /**
     * Optional semantic type used by formula validation and custom renderers.
     * It is deliberately additive: callers that only provide `type` retain
     * the original primitive-column behaviour.
     */
    readonly dataType?: SpreadsheetColumnDataType;
    /** Alias for integrations that call this metadata semanticType. */
    readonly semanticType?: SpreadsheetColumnDataType;
    /** Allowed values for dropdown semantic columns. */
    readonly allowedValues?: readonly string[];
    /** Optional numeric bounds for range/stars semantic columns. */
    readonly min?: number;
    readonly max?: number;
}

export type SpreadsheetColumnDataType =
    | "text"
    | "number"
    | "boolean"
    | "uri"
    | "date"
    | "dropdown"
    | "range"
    | "stars"
    | "sparkline"
    | "image"
    | "markdown"
    | "tags"
    | "multi-select"
    | "row-id"
    | "bubbles"
    | "drilldown"
    | "links"
    | "user-profile";

interface FormulaRecord {
    readonly compiled: CompiledFormula;
    readonly cells: ReadonlySet<string>;
    readonly columns: ReadonlySet<number>;
}

export interface SpreadsheetModelOptions {
    /** Optional synchronous custom-function registry used by formula cells. */
    readonly functionRegistry?: FormulaFunctionRegistry;
}

function key(col: number, row: number): string {
    return `${col}:${row}`;
}

export class SpreadsheetModel {
    private readonly inputs = new Map<string, CellInput>();
    private readonly formulas = new Map<string, FormulaRecord>();
    private readonly syntaxErrors = new Map<string, FormulaValue>();
    private readonly cache = new Map<string, FormulaValue>();
    private readonly reverseCells = new Map<string, Set<string>>();
    private readonly reverseColumns = new Map<number, Set<string>>();
    private readonly columnLookup = new Map<string, number>();
    /** Names shared by more than one column are deliberately not resolvable. */
    private readonly ambiguousColumnNames = new Set<string>();
    public readonly functionRegistry: FormulaFunctionRegistry | undefined;

    public constructor(
        columns: readonly SpreadsheetColumn[],
        rowCount: number,
        initialRows?: readonly (readonly CellInput[])[],
        options?: SpreadsheetModelOptions
    );
    public constructor(columns: readonly SpreadsheetColumn[], rowCount: number, options?: SpreadsheetModelOptions);
    public constructor(
        public readonly columns: readonly SpreadsheetColumn[],
        public readonly rowCount: number,
        initialRowsOrOptions: readonly (readonly CellInput[])[] | SpreadsheetModelOptions = [],
        options: SpreadsheetModelOptions = {}
    ) {
        const hasInitialRows = Array.isArray(initialRowsOrOptions);
        const initialRows: readonly (readonly CellInput[])[] = hasInitialRows ? initialRowsOrOptions : [];
        const selectedOptions = hasInitialRows ? options : (initialRowsOrOptions as SpreadsheetModelOptions);
        this.functionRegistry = selectedOptions.functionRegistry;
        columns.forEach((column, index) => {
            this.registerColumnName(column.id, index);
            this.registerColumnName(column.title, index);
        });
        initialRows.forEach((row, rowIndex) => row.forEach((value, colIndex) => this.setCell(colIndex, rowIndex, value)));
    }

    public resolveColumn(name: string): number | undefined {
        const normalized = this.normalizeColumnName(name);
        return normalized === "" || this.ambiguousColumnNames.has(normalized) ? undefined : this.columnLookup.get(normalized);
    }

    public getInput(col: number, row: number): CellInput {
        return this.inputs.get(key(col, row)) ?? null;
    }

    public setCell(col: number, row: number, input: CellInput): void {
        if (col < 0 || col >= this.columns.length || row < 0 || row >= this.rowCount) return;
        const cellKey = key(col, row);
        this.removeFormula(cellKey);
        this.syntaxErrors.delete(cellKey);
        this.inputs.set(cellKey, input);
        if (typeof input === "string" && input.startsWith("=")) {
            try {
                const compiled = compileFormula(input, { functionRegistry: this.functionRegistry });
                const dependencies = compiled.dependencies({
                    currentRow: row,
                    rowCount: this.rowCount,
                    resolveColumn: name => this.resolveColumn(name),
                });
                const record: FormulaRecord = { compiled, ...dependencies };
                this.formulas.set(cellKey, record);
                for (const dependency of record.cells) this.addReverse(this.reverseCells, dependency, cellKey);
                for (const dependency of record.columns) this.addReverse(this.reverseColumns, dependency, cellKey);
            } catch (error) {
                this.syntaxErrors.set(cellKey, formulaError("#VALUE!", error instanceof Error ? error.message : String(error)));
            }
        }
        this.invalidate(cellKey, col);
    }

    public getValue(col: number, row: number): FormulaValue {
        return this.evaluate(key(col, row), new Set<string>());
    }

    public getDisplayValue(col: number, row: number): string {
        const value = this.getValue(col, row);
        if (isFormulaError(value)) return value.code;
        return value === null ? "" : String(value);
    }

    public getFormulaCount(): number {
        return this.formulas.size;
    }

    /**
     * Clears all cached formula results. Registry mutations are intentionally
     * not observed automatically; call this after registering/unregistering a
     * function to make the change take effect deterministically.
     */
    public recalculateAll(): void {
        this.cache.clear();
    }

    /** Alias for recalculateAll for integrations that use invalidate terminology. */
    public invalidateAll(): void {
        this.recalculateAll();
    }

    private evaluate(cellKey: string, stack: Set<string>): FormulaValue {
        const cached = this.cache.get(cellKey);
        if (cached !== undefined) return cached;
        const syntaxError = this.syntaxErrors.get(cellKey);
        if (syntaxError !== undefined) return syntaxError;
        const formula = this.formulas.get(cellKey);
        if (formula === undefined) return this.inputs.get(cellKey) ?? null;
        if (stack.has(cellKey)) return formulaError("#CYCLE!", `Circular reference at ${cellKey}`);
        stack.add(cellKey);
        const [colText, rowText] = cellKey.split(":");
        const value = formula.compiled.evaluate({
            currentRow: Number(rowText),
            rowCount: this.rowCount,
            resolveColumn: name => this.resolveColumn(name),
            getCellValue: (col, row) => {
                if (col < 0 || col >= this.columns.length || row < 0 || row >= this.rowCount) return formulaError("#REF!");
                return this.evaluate(key(col, row), stack);
            },
        });
        stack.delete(cellKey);
        this.cache.set(cellKey, value);
        void colText;
        return value;
    }

    private removeFormula(cellKey: string): void {
        const existing = this.formulas.get(cellKey);
        if (existing === undefined) return;
        for (const dependency of existing.cells) this.deleteReverse(this.reverseCells, dependency, cellKey);
        for (const dependency of existing.columns) this.deleteReverse(this.reverseColumns, dependency, cellKey);
        this.formulas.delete(cellKey);
    }

    private addReverse<TKey>(map: Map<TKey, Set<string>>, dependency: TKey, formulaKey: string): void {
        const formulas = map.get(dependency) ?? new Set<string>();
        formulas.add(formulaKey);
        map.set(dependency, formulas);
    }

    private deleteReverse<TKey>(map: Map<TKey, Set<string>>, dependency: TKey, formulaKey: string): void {
        const formulas = map.get(dependency);
        formulas?.delete(formulaKey);
        if (formulas?.size === 0) map.delete(dependency);
    }

    private normalizeColumnName(name: string): string {
        return name.trim().normalize("NFKC").toLocaleLowerCase("en-US");
    }

    private registerColumnName(name: string, index: number): void {
        const normalized = this.normalizeColumnName(name);
        if (normalized === "" || this.ambiguousColumnNames.has(normalized)) return;
        const existing = this.columnLookup.get(normalized);
        if (existing === undefined) {
            this.columnLookup.set(normalized, index);
        } else if (existing !== index) {
            this.columnLookup.delete(normalized);
            this.ambiguousColumnNames.add(normalized);
        }
    }

    private invalidate(changedKey: string, changedCol: number): void {
        const queue = [changedKey];
        const visited = new Set<string>();
        while (queue.length > 0) {
            const current = queue.shift() as string;
            if (visited.has(current)) continue;
            visited.add(current);
            this.cache.delete(current);
            const currentCol = Number(current.split(":")[0]);
            const dependents = new Set([...(this.reverseCells.get(current) ?? []), ...(this.reverseColumns.get(currentCol) ?? [])]);
            dependents.forEach(dependent => queue.push(dependent));
        }
        for (const dependent of this.reverseColumns.get(changedCol) ?? []) this.cache.delete(dependent);
    }
}
