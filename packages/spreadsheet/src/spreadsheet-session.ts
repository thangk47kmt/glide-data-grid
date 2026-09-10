import type { HistoryTransaction } from "./history.js";
import type { CellInput, SpreadsheetColumn, SpreadsheetModel } from "./model.js";
import { SpreadsheetModel as Model } from "./model.js";
import { deleteColumns, deleteRows, insertColumns, insertRows, type StructuralEditLimits } from "./structural.js";
import { WorkbookHistory, cellTransactionToCommand, structuralResultToCommand, type WorkbookCommand } from "./workbook-history.js";
import { workbookPayloadFromModel } from "./workbook-io.js";

export interface SpreadsheetSessionOptions<S> {
    readonly maxHistory?: number;
    readonly selection?: S;
}

/**
 * Owns a replaceable spreadsheet model and one ordered history for both cell
 * and structural edits. Structural operations materialize one raw workbook
 * payload, then reuse their immutable before/after snapshots for undo/redo.
 */
export class SpreadsheetSession<S = object> {
    public readonly history: WorkbookHistory;
    private activeModel: SpreadsheetModel;
    private activeSelection: S | undefined;
    private revisionValue = 0;

    public constructor(model: SpreadsheetModel, options: SpreadsheetSessionOptions<S> = {}) {
        if (model === null || typeof model !== "object") throw new TypeError("A SpreadsheetModel is required");
        this.activeModel = model;
        this.activeSelection = options.selection;
        this.history = new WorkbookHistory({ maxHistory: options.maxHistory });
    }

    public get model(): SpreadsheetModel { return this.activeModel; }
    public get selection(): S | undefined { return this.activeSelection; }
    public get revision(): number { return this.revisionValue; }

    public executeCells(transaction: HistoryTransaction<CellInput, S>): Readonly<WorkbookCommand> | undefined {
        const base = cellTransactionToCommand(transaction, (edit, direction) => {
            this.activeModel.setCell(edit.location[0], edit.location[1], direction === "undo" ? edit.before : edit.after);
        });
        return this.history.execute(this.decorate(base, transaction.selectionBefore, transaction.selectionAfter));
    }

    public insertRows(index: number, count: number, limits?: StructuralEditLimits, selectionAfter?: S): Readonly<WorkbookCommand> | undefined {
        return this.executeStructural(insertRows(workbookPayloadFromModel(this.activeModel), index, count, limits), `Insert ${count} row(s)`, selectionAfter);
    }

    public deleteRows(index: number, count: number, limits?: StructuralEditLimits, selectionAfter?: S): Readonly<WorkbookCommand> | undefined {
        return this.executeStructural(deleteRows(workbookPayloadFromModel(this.activeModel), index, count, limits), `Delete ${count} row(s)`, selectionAfter);
    }

    public insertColumns(index: number, columns: readonly SpreadsheetColumn[], limits?: StructuralEditLimits, selectionAfter?: S): Readonly<WorkbookCommand> | undefined {
        return this.executeStructural(insertColumns(workbookPayloadFromModel(this.activeModel), index, columns, limits), `Insert ${columns.length} column(s)`, selectionAfter);
    }

    public deleteColumns(index: number, count: number, limits?: StructuralEditLimits, selectionAfter?: S): Readonly<WorkbookCommand> | undefined {
        return this.executeStructural(deleteColumns(workbookPayloadFromModel(this.activeModel), index, count, limits), `Delete ${count} column(s)`, selectionAfter);
    }

    private executeStructural(result: ReturnType<typeof insertRows>, label: string, selectionAfter?: S): Readonly<WorkbookCommand> | undefined {
        const selectionBefore = this.activeSelection;
        const base = structuralResultToCommand(result, {
            id: `structural-${this.revisionValue + 1}`,
            label,
            ...(selectionBefore === undefined ? {} : { selectionBefore }),
            ...(selectionAfter === undefined ? {} : { selectionAfter }),
        }, snapshot => {
            this.activeModel = new Model(snapshot.columns, snapshot.rowCount, snapshot.rows, { functionRegistry: this.activeModel.functionRegistry });
        });
        return this.history.execute(this.decorate(base, selectionBefore, selectionAfter));
    }

    private decorate(command: WorkbookCommand | undefined, before: S | undefined, after: S | undefined): WorkbookCommand | undefined {
        if (command === undefined) return undefined;
        return {
            id: command.id,
            ...(command.label === undefined ? {} : { label: command.label }),
            apply: () => {
                command.apply();
                this.activeSelection = after;
                this.revisionValue++;
            },
            revert: () => {
                command.revert();
                this.activeSelection = before;
                this.revisionValue++;
            },
        };
    }
}
