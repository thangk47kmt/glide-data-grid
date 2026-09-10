import type { CellEdit, CellLocation, HistoryApply, HistoryTransaction } from "./history.js";
import type { SnapshotApply, SnapshotTransaction } from "./snapshot-history.js";
import { snapshotTransactionFromStructural } from "./snapshot-history.js";
import type { StructuralEditResult } from "./structural.js";

export interface WorkbookCommand {
    readonly id: string;
    readonly label?: string;
    readonly apply: () => void;
    readonly revert: () => void;
}

export interface WorkbookCommandOptions {
    readonly maxHistory?: number;
}

function validateMaxHistory(value: number): number {
    if (!Number.isInteger(value) || value < 0) throw new RangeError("maxHistory must be a non-negative integer");
    return value;
}

function normalizeCommand(command: WorkbookCommand | undefined): WorkbookCommand | undefined {
    if (command === undefined) return undefined;
    if (command === null || typeof command !== "object") throw new TypeError("A workbook command is required");
    if (typeof command.id !== "string") throw new TypeError("A workbook command id must be a string");
    if (typeof command.apply !== "function" || typeof command.revert !== "function") throw new TypeError("A workbook command requires apply and revert callbacks");
    if (command.label !== undefined && typeof command.label !== "string") throw new TypeError("A workbook command label must be a string");
    // Keep a command record of our own so mutating the caller's object after
    // record cannot alter history. Callback closures may intentionally retain
    // snapshot references; those snapshots remain caller-owned and immutable.
    return {
        id: command.id,
        ...(command.label === undefined ? {} : { label: command.label }),
        apply: command.apply,
        revert: command.revert,
    };
}

/**
 * A single command stack shared by cell and workbook/schema changes. The stack
 * itself preserves state when a callback throws, but cannot undo arbitrary
 * external side effects. Cell adapters additionally compensate already-applied
 * edits best-effort; snapshot/generic commands are atomic only when their
 * consumer-owned callback is atomic.
 */
export class WorkbookHistory {
    private readonly undoStack: WorkbookCommand[] = [];
    private readonly redoStack: WorkbookCommand[] = [];
    private historyLimit: number;
    private applying = false;

    public constructor(options: WorkbookCommandOptions = {}) {
        this.historyLimit = validateMaxHistory(options.maxHistory ?? 100);
    }

    public get maxHistory(): number {
        return this.historyLimit;
    }

    public set maxHistory(value: number) {
        this.historyLimit = validateMaxHistory(value);
        this.trimStacks();
    }

    public canUndo(): boolean {
        return this.undoStack.length > 0;
    }

    public canRedo(): boolean {
        return this.redoStack.length > 0;
    }

    public get undoCount(): number {
        return this.undoStack.length;
    }

    public get redoCount(): number {
        return this.redoStack.length;
    }

    public execute(command: WorkbookCommand | undefined): Readonly<WorkbookCommand> | undefined {
        this.assertNotApplying();
        const normalized = normalizeCommand(command);
        if (normalized === undefined) return undefined;
        this.withApplying(normalized.apply);
        this.pushUndo(normalized);
        this.redoStack.length = 0;
        return normalized;
    }

    /** Records an already-applied command without invoking either callback. */
    public record(command: WorkbookCommand | undefined): Readonly<WorkbookCommand> | undefined {
        this.assertNotApplying();
        const normalized = normalizeCommand(command);
        if (normalized === undefined) return undefined;
        this.pushUndo(normalized);
        this.redoStack.length = 0;
        return normalized;
    }

    public push(command: WorkbookCommand | undefined): Readonly<WorkbookCommand> | undefined {
        return this.record(command);
    }

    public undo(): Readonly<WorkbookCommand> | undefined {
        this.assertNotApplying();
        const command = this.undoStack[this.undoStack.length - 1];
        if (command === undefined) return undefined;
        this.withApplying(command.revert);
        this.undoStack.pop();
        this.pushRedo(command);
        return command;
    }

    public redo(): Readonly<WorkbookCommand> | undefined {
        this.assertNotApplying();
        const command = this.redoStack[this.redoStack.length - 1];
        if (command === undefined) return undefined;
        this.withApplying(command.apply);
        this.redoStack.pop();
        this.pushUndo(command);
        return command;
    }

    public clear(): void {
        this.assertNotApplying();
        this.undoStack.length = 0;
        this.redoStack.length = 0;
    }

    private pushUndo(command: WorkbookCommand): void {
        if (this.historyLimit === 0) return;
        this.undoStack.push(command);
        if (this.undoStack.length > this.historyLimit) this.undoStack.shift();
    }

    private pushRedo(command: WorkbookCommand): void {
        if (this.historyLimit === 0) return;
        this.redoStack.push(command);
        if (this.redoStack.length > this.historyLimit) this.redoStack.shift();
    }

    private trimStacks(): void {
        while (this.undoStack.length > this.historyLimit) this.undoStack.shift();
        while (this.redoStack.length > this.historyLimit) this.redoStack.shift();
    }

    private withApplying(action: () => void): void {
        this.applying = true;
        try {
            action();
        } finally {
            this.applying = false;
        }
    }

    private assertNotApplying(): void {
        if (this.applying) throw new Error("Workbook history cannot be modified from a command callback");
    }
}

function normalizeCellTransaction<T, S>(transaction: HistoryTransaction<T, S>): HistoryTransaction<T, S> {
    if (transaction === null || typeof transaction !== "object") throw new TypeError("A cell transaction is required");
    if (typeof transaction.id !== "string") throw new TypeError("A cell transaction id must be a string");
    if (transaction.label !== undefined && typeof transaction.label !== "string") throw new TypeError("A cell transaction label must be a string");
    if (!Array.isArray(transaction.edits)) throw new TypeError("A cell transaction requires an edits array");
    const coalesced = new Map<string, CellEdit<T>>();
    transaction.edits.forEach(edit => {
        if (edit === null || typeof edit !== "object" || !Array.isArray(edit.location) || edit.location.length !== 2) throw new TypeError("Each cell edit requires a [col, row] location");
        const [col, row] = edit.location;
        if (!Number.isSafeInteger(col) || !Number.isSafeInteger(row) || col < 0 || row < 0) throw new RangeError("Cell locations must contain non-negative integer coordinates");
        const location: CellLocation = [col, row];
        const editKey = `${col}:${row}`;
        const previous = coalesced.get(editKey);
        coalesced.set(editKey, previous === undefined
            ? { location, before: edit.before, after: edit.after }
            : { location: previous.location, before: previous.before, after: edit.after });
    });
    const edits = [...coalesced.values()].filter(edit => !Object.is(edit.before, edit.after));
    return {
        id: transaction.id,
        ...(transaction.label === undefined ? {} : { label: transaction.label }),
        edits,
        ...(transaction.selectionBefore === undefined ? {} : { selectionBefore: transaction.selectionBefore }),
        ...(transaction.selectionAfter === undefined ? {} : { selectionAfter: transaction.selectionAfter }),
    };
}

/** Adapts an existing cell transaction to one unified workbook command. */
export function cellTransactionToCommand<T, S>(transaction: HistoryTransaction<T, S>, apply: HistoryApply<T, S>): WorkbookCommand | undefined {
    if (typeof apply !== "function") throw new TypeError("A cell command apply callback is required");
    const normalized = normalizeCellTransaction(transaction);
    if (normalized.edits.length === 0) return undefined;
    const run = (direction: "redo" | "undo"): void => {
        const ordered = direction === "undo" ? [...normalized.edits].reverse() : [...normalized.edits];
        const applied: CellEdit<T>[] = [];
        try {
            for (const edit of ordered) {
                apply(edit, direction, normalized);
                applied.push(edit);
            }
        } catch (error) {
            const compensationDirection = direction === "redo" ? "undo" : "redo";
            // Compensation always reverses the order in which edits actually
            // completed, including undo (whose completed order is already
            // reverse-original). This restores dependent edits safely.
            const compensation = [...applied].reverse();
            const compensationErrors: unknown[] = [];
            for (const edit of compensation) {
                try {
                    apply(edit, compensationDirection, normalized);
                } catch (compensationError) {
                    compensationErrors.push(compensationError);
                }
            }
            if (compensationErrors.length > 0) throw new AggregateError([error, ...compensationErrors], "Cell command compensation failed");
            throw error;
        }
    };
    return {
        id: normalized.id,
        ...(normalized.label === undefined ? {} : { label: normalized.label }),
        apply: () => run("redo"),
        revert: () => run("undo"),
    };
}

export const createCellEditCommand = cellTransactionToCommand;

function normalizeSnapshotTransaction<T, S>(transaction: SnapshotTransaction<T, S>): SnapshotTransaction<T, S> {
    if (transaction === null || typeof transaction !== "object") throw new TypeError("A snapshot transaction is required");
    if (typeof transaction.id !== "string") throw new TypeError("A snapshot transaction id must be a string");
    if (!("before" in transaction) || !("after" in transaction)) throw new TypeError("A snapshot transaction requires before and after snapshots");
    if (transaction.label !== undefined && typeof transaction.label !== "string") throw new TypeError("A snapshot transaction label must be a string");
    return {
        id: transaction.id,
        ...(transaction.label === undefined ? {} : { label: transaction.label }),
        before: transaction.before,
        after: transaction.after,
        ...(transaction.selectionBefore === undefined ? {} : { selectionBefore: transaction.selectionBefore }),
        ...(transaction.selectionAfter === undefined ? {} : { selectionAfter: transaction.selectionAfter }),
    };
}

/** Adapts an atomic snapshot transaction without cloning snapshot payloads. */
export function snapshotTransactionToCommand<T, S>(transaction: SnapshotTransaction<T, S>, apply: SnapshotApply<T, S>): WorkbookCommand | undefined {
    if (typeof apply !== "function") throw new TypeError("A snapshot command apply callback is required");
    const normalized = normalizeSnapshotTransaction(transaction);
    if (Object.is(normalized.before, normalized.after)) return undefined;
    return {
        id: normalized.id,
        ...(normalized.label === undefined ? {} : { label: normalized.label }),
        apply: () => apply(normalized.after, "redo", normalized),
        revert: () => apply(normalized.before, "undo", normalized),
    };
}

export const createSnapshotCommand = snapshotTransactionToCommand;

/** Converts a structural result into one atomic workbook replacement command. */
export function structuralResultToCommand<S = object>(
    result: StructuralEditResult,
    options: { readonly id: string; readonly label?: string; readonly selectionBefore?: S; readonly selectionAfter?: S },
    apply: SnapshotApply<StructuralEditResult["payload"], S>
): WorkbookCommand | undefined {
    const transaction = snapshotTransactionFromStructural(result, options);
    return snapshotTransactionToCommand(transaction, apply);
}

export const createStructuralCommand = structuralResultToCommand;
