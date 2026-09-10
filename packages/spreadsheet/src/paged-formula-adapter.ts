import {
    compileFormula,
    formulaError,
    isFormulaError,
    type CompiledFormula,
    type FormulaDependencies,
    type FormulaError,
    type FormulaValue,
} from "./formula.js";
import type { FormulaFunctionRegistry } from "./function-registry.js";
import {
    LOADING_CELL,
    PagedDataSource,
    type LoadingCell,
    type PagedCellValue,
} from "./paged-data-source.js";
import type { SpreadsheetColumn } from "./model.js";
import { formulaColumnSemanticType, validateFormulaResultType, type FormulaSemanticType } from "./formula-validation.js";

/** A raw value accepted by a paged formula cell. Formula strings start with `=`. */
export type PagedFormulaInput = PagedCellValue;

/** A value returned by the adapter. */
export type PagedFormulaValue = FormulaValue | LoadingCell;

export interface PagedFormulaAdapterOptions {
    /** Optional synchronous custom-function registry used by formula cells. */
    readonly functionRegistry?: FormulaFunctionRegistry;
    /**
     * Maximum row count for a whole-column structured reference (`[Amount]`).
     * Whole-column evaluation asks the shared formula engine for every row, so
     * the default deliberately keeps it away from a 500k-row data source.
     * Set to `Infinity` only when the caller can afford that work.
     */
    readonly maxWholeColumnRows?: number;
}

export interface PagedFormulaAdapterStats {
    readonly formulaCells: number;
    readonly syntaxErrors: number;
    readonly cachedValues: number;
    readonly formulaEvaluations: number;
}

/** The three scalar types supported by SpreadsheetColumn. */
export type SpreadsheetValueType = FormulaSemanticType;

/** Result returned by {@link PagedFormulaAdapter.validateFormula}. */
export interface FormulaValidationResult {
    readonly valid: boolean;
    readonly value?: FormulaValue;
    readonly error?: FormulaError;
    readonly expectedType?: SpreadsheetValueType;
    readonly actualType?: SpreadsheetValueType;
}

interface FormulaRecord {
    readonly compiled: CompiledFormula;
    readonly dependencies: FormulaDependencies;
}

const LOADING_SIGNAL = Symbol("paged-formula-loading");

type ColumnDefinition = SpreadsheetColumn | string;

function cellKey(col: number, row: number): string {
    return `${col}:${row}`;
}

function columnName(column: ColumnDefinition): string {
    return typeof column === "string" ? column : column.id;
}

function columnTitle(column: ColumnDefinition): string {
    return typeof column === "string" ? column : column.title;
}

function columnType(column: ColumnDefinition | undefined): SpreadsheetValueType | undefined {
    return column === undefined || typeof column === "string" ? undefined : formulaColumnSemanticType(column);
}

function valueType(value: FormulaValue): SpreadsheetValueType | undefined {
    if (value === null) return undefined;
    if (typeof value === "number") return "number";
    if (typeof value === "boolean") return "boolean";
    if (typeof value === "string") return "text";
    return undefined;
}

/**
 * Adapts the synchronous formula engine to a virtual/paged data source.
 *
 * Unlike `SpreadsheetModel`, this class never allocates a row-by-column input
 * matrix. Only cells containing formulas, their dependency indexes, and lazy
 * formula results are retained. Base values remain owned by `PagedDataSource`.
 */
export class PagedFormulaAdapter {
    private readonly formulas = new Map<string, FormulaRecord>();
    private readonly formulaSources = new Map<string, string>();
    private readonly syntaxErrors = new Map<string, FormulaValue>();
    private readonly cache = new Map<string, FormulaValue>();
    private readonly reverseCells = new Map<string, Set<string>>();
    private readonly reverseColumns = new Map<number, Set<string>>();
    private readonly formulaColumnsByRow = new Map<number, Set<number>>();
    private readonly columnLookup = new Map<string, number>();
    /** Captions shared by multiple columns must not silently bind formulas. */
    private readonly ambiguousColumnNames = new Set<string>();
    private formulaEvaluations = 0;
    private readonly maxWholeColumnRows: number;

    public readonly columns: readonly ColumnDefinition[];
    public readonly source: PagedDataSource;
    public readonly functionRegistry: FormulaFunctionRegistry | undefined;

    public constructor(
        source: PagedDataSource,
        columns: readonly ColumnDefinition[],
        options: PagedFormulaAdapterOptions = {}
    ) {
        if (columns.length !== source.columnCount) {
            throw new RangeError("columns must have the same length as the data source");
        }
        this.source = source;
        this.columns = columns;
        this.functionRegistry = options.functionRegistry;
        const maxRows = options.maxWholeColumnRows ?? 10_000;
        if (maxRows !== Infinity && (!Number.isInteger(maxRows) || maxRows < 0)) {
            throw new RangeError("maxWholeColumnRows must be a non-negative integer or Infinity");
        }
        this.maxWholeColumnRows = maxRows;
        columns.forEach((column, index) => {
            this.registerColumnName(columnName(column), index);
            this.registerColumnName(columnTitle(column), index);
        });
    }

    /** Resolves a column id/title for structured references such as `[@Price]`. */
    public resolveColumn(name: string): number | undefined {
        const normalized = this.normalizeColumnName(name);
        return normalized === "" || this.ambiguousColumnNames.has(normalized) ? undefined : this.columnLookup.get(normalized);
    }

    /** Returns the raw formula text, or the sparse source value for a literal cell. */
    public getInput(col: number, row: number): PagedCellValue | LoadingCell {
        this.validateCell(col, row);
        const key = cellKey(col, row);
        const formula = this.formulaSources.get(key);
        if (formula !== undefined) return formula;
        return this.source.readCell(row, col);
    }

    /** Returns a formula's raw text, including a formula that currently has a syntax error. */
    public getFormula(col: number, row: number): string | undefined {
        this.validateCell(col, row);
        return this.formulaSources.get(cellKey(col, row));
    }

    public hasFormula(col: number, row: number): boolean {
        return this.getFormula(col, row) !== undefined;
    }

    /** Sparse formula columns for one row; useful to keep virtual queries formula-aware. */
    public getFormulaColumns(row: number): readonly number[] {
        if (!Number.isInteger(row) || row < 0 || row >= this.source.rowCount) throw new RangeError("row is outside the data source");
        return [...(this.formulaColumnsByRow.get(row) ?? [])];
    }

    /**
     * Sets a sparse literal or formula. Formula strings are compiled and indexed
     * immediately, but their value is calculated only when `getValue` is called.
     */
    public setCell(col: number, row: number, input: PagedFormulaInput): void {
        this.validateCell(col, row);
        const key = cellKey(col, row);
        this.removeFormula(key);
        this.syntaxErrors.delete(key);
        if (typeof input === "string" && input.startsWith("=")) {
            this.formulaSources.set(key, input);
            const rowColumns = this.formulaColumnsByRow.get(row) ?? new Set<number>();
            rowColumns.add(col);
            this.formulaColumnsByRow.set(row, rowColumns);
            try {
                const compiled = compileFormula(input, { functionRegistry: this.functionRegistry });
                const dependencies = compiled.dependencies({
                    currentRow: row,
                    rowCount: this.source.rowCount,
                    resolveColumn: name => this.resolveColumn(name),
                });
                const record = { compiled, dependencies };
                this.formulas.set(key, record);
                this.addDependencies(key, dependencies);
            } catch (error) {
                this.syntaxErrors.set(key, formulaError("#VALUE!", error instanceof Error ? error.message : String(error)));
            }
        } else {
            this.formulaSources.delete(key);
            this.removeFormulaColumn(row, col);
            this.source.updateCell(row, col, input);
        }
        this.invalidate(key, col);
    }

    /** Convenience method for formula-bar integrations. */
    public setFormula(col: number, row: number, formula: string): void {
        if (typeof formula !== "string" || !formula.startsWith("=")) throw new TypeError("A formula must be a string starting with '='");
        this.setCell(col, row, formula);
    }

    /**
     * Validates a formula without changing the sparse source or formula
     * indexes. This is intended for editors that need to keep their draft
     * open when a formula is rejected.
     *
     * A typed column is deliberately strict: every referenced typed column
     * must have the same type as the destination, and the evaluated result
     * must also have that type. Empty results remain valid for nullable cells.
     * Formula errors are returned with their original Excel-style code and a
     * concrete message so callers can show an actionable status to the user.
     */
    public validateFormula(col: number, row: number, formula: string): FormulaValidationResult {
        this.validateCell(col, row);
        if (typeof formula !== "string" || !formula.startsWith("=")) {
            return { valid: false, error: formulaError("#VALUE!", "A formula must start with '='") };
        }

        let compiled: CompiledFormula;
        let dependencies: FormulaDependencies;
        try {
            compiled = compileFormula(formula, { functionRegistry: this.functionRegistry });
            dependencies = compiled.dependencies({
                currentRow: row,
                rowCount: this.source.rowCount,
                resolveColumn: name => this.resolveColumn(name),
            });
        } catch (error) {
            return {
                valid: false,
                error: formulaError("#VALUE!", error instanceof Error ? error.message : String(error)),
            };
        }

        // Resolve coordinate errors before checking declared column types. A
        // formula such as `=A999` must report the out-of-range reference, not
        // a misleading text/number mismatch from column A.
        for (const dependency of dependencies.cells) {
            const [dependencyCol, dependencyRow] = dependency.split(":").map(Number);
            if (!Number.isInteger(dependencyCol) || !Number.isInteger(dependencyRow) || dependencyCol < 0 || dependencyCol >= this.source.columnCount || dependencyRow < 0 || dependencyRow >= this.source.rowCount) {
                return {
                    valid: false,
                    error: formulaError("#REF!", `Cell reference ${dependency} is outside the data source`),
                };
            }
        }

        const expectedType = columnType(this.columns[col]!);
        if (expectedType !== undefined) {
            const dependencyColumns = new Set<number>();
            dependencies.columns.forEach(dependencyColumns.add, dependencyColumns);
            dependencies.cells.forEach(key => dependencyColumns.add(Number(key.split(":")[0])));
            for (const dependencyCol of dependencyColumns) {
                const sourceColumn = this.columns[dependencyCol];
                // Invalid A1 columns are reported by evaluation as #REF!, not
                // as a type mismatch while walking the dependency metadata.
                if (sourceColumn === undefined) continue;
                const sourceType = columnType(sourceColumn);
                if (sourceType !== undefined && sourceType !== expectedType) {
                    const sourceName = columnTitle(sourceColumn);
                    const targetName = columnTitle(this.columns[col]!);
                    const error = formulaError(
                        "#VALUE!",
                        `Cannot use ${sourceName} (${sourceType}) in ${targetName} (${expectedType}); source and target types must match`,
                    );
                    return { valid: false, error, expectedType, actualType: sourceType };
                }
            }
            if (dependencies.columns.size > 0 && this.source.rowCount > this.maxWholeColumnRows) {
                return {
                    valid: false,
                    expectedType,
                    error: formulaError("#VALUE!", "Whole-column references are disabled for this paged dataset"),
                };
            }
        }

        const targetKey = cellKey(col, row);
        const stack = new Set<string>();
        const evaluateCell = (key: string): FormulaValue => {
            if (stack.has(key)) return formulaError("#CYCLE!", `Circular reference at ${key}`);
            const [colText, rowText] = key.split(":");
            const dependencyCol = Number(colText);
            const dependencyRow = Number(rowText);
            if (dependencyCol < 0 || dependencyCol >= this.source.columnCount || dependencyRow < 0 || dependencyRow >= this.source.rowCount) {
                return formulaError("#REF!", `Cell reference ${key} is outside the data source`);
            }
            const record = key === targetKey ? { compiled, dependencies } : this.formulas.get(key);
            if (record === undefined) return this.source.readCell(dependencyRow, dependencyCol);
            if (record.dependencies.columns.size > 0 && this.source.rowCount > this.maxWholeColumnRows) {
                return formulaError("#VALUE!", "Whole-column references are disabled for this paged dataset");
            }
            stack.add(key);
            try {
                return record.compiled.evaluate({
                    currentRow: dependencyRow,
                    rowCount: this.source.rowCount,
                    resolveColumn: name => this.resolveColumn(name),
                    functionRegistry: this.functionRegistry,
                    getCellValue: (nextCol, nextRow) => {
                        if (nextCol < 0 || nextCol >= this.source.columnCount || nextRow < 0 || nextRow >= this.source.rowCount) {
                            return formulaError("#REF!", `Cell reference ${nextCol}:${nextRow} is outside the data source`);
                        }
                        return evaluateCell(cellKey(nextCol, nextRow));
                    },
                });
            } catch (error) {
                return formulaError("#VALUE!", error instanceof Error ? error.message : String(error));
            } finally {
                stack.delete(key);
            }
        };

        const value = evaluateCell(targetKey);
        if (isFormulaError(value)) return { valid: false, expectedType, error: value, value };
        const formulaValue = value as FormulaValue;
        const actualType = valueType(formulaValue);
        const resultError = typeof this.columns[col] === "string" ? undefined : validateFormulaResultType(formulaValue, this.columns[col]!);
        if (resultError !== undefined) return { valid: false, value: formulaValue, error: resultError, expectedType, actualType };
        return { valid: true, value: formulaValue, expectedType, actualType };
    }

    /** Removes a formula and reveals the source value that existed underneath it. */
    public clearFormula(col: number, row: number): void {
        this.validateCell(col, row);
        const key = cellKey(col, row);
        if (!this.formulaSources.has(key)) return;
        this.removeFormula(key);
        this.formulaSources.delete(key);
        this.syntaxErrors.delete(key);
        this.removeFormulaColumn(row, col);
        this.invalidate(key, col);
    }

    /** Gets a lazy computed value directly from the virtual source. */
    public getValue(col: number, row: number): PagedFormulaValue {
        this.validateCell(col, row);
        return this.evaluate(cellKey(col, row), new Set<string>());
    }

    public getDisplayValue(col: number, row: number): string {
        const value = this.getValue(col, row);
        if (value === LOADING_CELL) return "";
        if (isFormulaError(value)) return value.code;
        return value === null ? "" : String(value);
    }

    public getFormulaCount(): number {
        return this.formulas.size;
    }

    public getStats(): PagedFormulaAdapterStats {
        return {
            formulaCells: this.formulas.size,
            syntaxErrors: this.syntaxErrors.size,
            cachedValues: this.cache.size,
            formulaEvaluations: this.formulaEvaluations,
        };
    }

    /**
     * Call after mutating the source directly through `source.updateCell`.
     * Adapter-owned edits already invalidate automatically.
     */
    public notifyCellChanged(col: number, row: number): void {
        this.validateCell(col, row);
        this.invalidate(cellKey(col, row), col);
    }

    /** Clears computed values while retaining formulas and dependency indexes. */
    public recalculateAll(): void {
        this.cache.clear();
    }

    public invalidateAll(): void {
        this.recalculateAll();
    }

    private evaluate(key: string, stack: Set<string>): FormulaValue | LoadingCell {
        if (this.cache.has(key)) return this.cache.get(key) as FormulaValue;
        const syntaxError = this.syntaxErrors.get(key);
        if (syntaxError !== undefined) return syntaxError;
        const formula = this.formulas.get(key);
        if (formula === undefined) {
            const [colText, rowText] = key.split(":");
            return this.source.readCell(Number(rowText), Number(colText));
        }
        if (stack.has(key)) return formulaError("#CYCLE!", `Circular reference at ${key}`);
        stack.add(key);
        const [, rowText] = key.split(":");
        let value: FormulaValue;
        try {
            if (formula.dependencies.columns.size > 0 && this.source.rowCount > this.maxWholeColumnRows) {
                value = formulaError("#VALUE!", "Whole-column references are disabled for this paged dataset");
            } else {
                this.formulaEvaluations += 1;
                value = formula.compiled.evaluate({
                    currentRow: Number(rowText),
                    rowCount: this.source.rowCount,
                    resolveColumn: name => this.resolveColumn(name),
                    functionRegistry: this.functionRegistry,
                    getCellValue: (dependencyCol, dependencyRow) => {
                        if (dependencyCol < 0 || dependencyCol >= this.source.columnCount || dependencyRow < 0 || dependencyRow >= this.source.rowCount) {
                            return formulaError("#REF!", "Cell reference is outside the data source");
                        }
                        const dependency = this.evaluate(cellKey(dependencyCol, dependencyRow), stack);
                        if (dependency === LOADING_CELL) throw LOADING_SIGNAL;
                        return dependency as FormulaValue;
                    },
                });
            }
        } catch (error) {
            if (error === LOADING_SIGNAL) {
                stack.delete(key);
                return LOADING_CELL;
            }
            value = formulaError("#VALUE!", error instanceof Error ? error.message : String(error));
        }
        stack.delete(key);
        this.cache.set(key, value);
        return value;
    }

    private addDependencies(formulaKey: string, dependencies: FormulaDependencies): void {
        for (const dependency of dependencies.cells) this.addReverse(this.reverseCells, dependency, formulaKey);
        for (const dependency of dependencies.columns) this.addReverse(this.reverseColumns, dependency, formulaKey);
    }

    private removeFormula(formulaKey: string): void {
        const existing = this.formulas.get(formulaKey);
        if (existing !== undefined) {
            for (const dependency of existing.dependencies.cells) this.deleteReverse(this.reverseCells, dependency, formulaKey);
            for (const dependency of existing.dependencies.columns) this.deleteReverse(this.reverseColumns, dependency, formulaKey);
            this.formulas.delete(formulaKey);
        }
        this.cache.delete(formulaKey);
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
            const dependents = new Set([
                ...(this.reverseCells.get(current) ?? []),
                ...(this.reverseColumns.get(currentCol) ?? []),
            ]);
            dependents.forEach(dependent => queue.push(dependent));
        }
        for (const dependent of this.reverseColumns.get(changedCol) ?? []) this.cache.delete(dependent);
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

    private removeFormulaColumn(row: number, col: number): void {
        const columns = this.formulaColumnsByRow.get(row);
        columns?.delete(col);
        if (columns?.size === 0) this.formulaColumnsByRow.delete(row);
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

    private validateCell(col: number, row: number): void {
        if (!Number.isInteger(col) || col < 0 || col >= this.source.columnCount) throw new RangeError("column is outside the data source");
        if (!Number.isInteger(row) || row < 0 || row >= this.source.rowCount) throw new RangeError("row is outside the data source");
    }
}
