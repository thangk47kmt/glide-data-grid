/** A column/row pair identifying a cell in a spreadsheet. */
export type CellLocation = readonly [col: number, row: number];

/**
 * Selection information is deliberately kept independent of the grid package.
 * Applications may use any selection object, or provide a more specific
 * selection type through the transaction/history generics.
 */
export type HistorySelection = object;

export interface CellEdit<T = unknown> {
    readonly location: CellLocation;
    readonly before: T;
    readonly after: T;
}

export interface HistoryTransaction<T = unknown, S = HistorySelection> {
    readonly id: string;
    readonly label?: string;
    readonly edits: readonly CellEdit<T>[];
    readonly selectionBefore?: S;
    readonly selectionAfter?: S;
}

export type HistoryDirection = "redo" | "undo";

/**
 * Applies one edit. For redo (including the initial execute), consumers should
 * apply `edit.after`; for undo, they should apply `edit.before`.
 */
export type HistoryApply<T = unknown, S = HistorySelection> = (
    edit: CellEdit<T>,
    direction: HistoryDirection,
    transaction: Readonly<HistoryTransaction<T, S>>
) => void;

export interface HistoryOptions<T = unknown, S = HistorySelection> {
    readonly apply: HistoryApply<T, S>;
    readonly maxHistory?: number;
}

type MutableHistoryTransaction<T, S> = {
    readonly id: string;
    readonly label?: string;
    readonly edits: readonly CellEdit<T>[];
    readonly selectionBefore?: S;
    readonly selectionAfter?: S;
};

function sameValue(left: unknown, right: unknown): boolean {
    return Object.is(left, right);
}

function validateMaxHistory(value: number): number {
    if (!Number.isInteger(value) || value < 0) {
        throw new RangeError("maxHistory must be a non-negative integer");
    }
    return value;
}

function normalizeTransaction<T, S>(transaction: HistoryTransaction<T, S>): MutableHistoryTransaction<T, S> | undefined {
    if (transaction === null || typeof transaction !== "object") {
        throw new TypeError("A history transaction is required");
    }
    if (typeof transaction.id !== "string") {
        throw new TypeError("A history transaction id must be a string");
    }

    // A map retains the first occurrence's position and before value while
    // updating the after value for subsequent edits to the same cell.
    const coalesced = new Map<string, CellEdit<T>>();
    for (const edit of transaction.edits) {
        if (edit === null || typeof edit !== "object" || !Array.isArray(edit.location) || edit.location.length !== 2) {
            throw new TypeError("Each history edit must have a [col, row] location");
        }
        const [col, row] = edit.location;
        if (!Number.isInteger(col) || !Number.isInteger(row) || col < 0 || row < 0) {
            throw new RangeError("Cell locations must contain non-negative integer coordinates");
        }
        const location: CellLocation = [col, row];
        const editKey = `${col}:${row}`;
        const previous = coalesced.get(editKey);
        if (previous === undefined) {
            coalesced.set(editKey, { location, before: edit.before, after: edit.after });
        } else {
            coalesced.set(editKey, { location: previous.location, before: previous.before, after: edit.after });
        }
    }

    const edits = [...coalesced.values()].filter(edit => !sameValue(edit.before, edit.after));
    if (edits.length === 0) return undefined;
    return {
        id: transaction.id,
        ...(transaction.label === undefined ? {} : { label: transaction.label }),
        edits,
        ...(transaction.selectionBefore === undefined ? {} : { selectionBefore: transaction.selectionBefore }),
        ...(transaction.selectionAfter === undefined ? {} : { selectionAfter: transaction.selectionAfter }),
    };
}

/**
 * A model-independent transaction history with bounded undo and redo stacks.
 * Transactions are immutable snapshots of their edit lists. Recording from an
 * apply callback is rejected to keep stack changes deterministic. If an apply
 * callback throws, the corresponding undo/redo stacks are left unchanged; the
 * engine cannot roll back external side effects already made by earlier
 * callback invocations, so consumers that require atomicity should provide a
 * transactional apply callback.
 */
export class History<T = unknown, S = HistorySelection> {
    private readonly applyEdit: HistoryApply<T, S>;
    private readonly undoStack: Array<MutableHistoryTransaction<T, S>> = [];
    private readonly redoStack: Array<MutableHistoryTransaction<T, S>> = [];
    private applying = false;
    private historyLimit: number;

    public constructor(options: HistoryOptions<T, S>);
    public constructor(apply: HistoryApply<T, S>, options?: { readonly maxHistory?: number } | number);
    public constructor(
        optionsOrApply: HistoryOptions<T, S> | HistoryApply<T, S>,
        options: { readonly maxHistory?: number } | number = {}
    ) {
        if (typeof optionsOrApply === "function") {
            this.applyEdit = optionsOrApply;
            this.historyLimit = validateMaxHistory(typeof options === "number" ? options : options.maxHistory ?? 100);
        } else {
            this.applyEdit = optionsOrApply.apply;
            this.historyLimit = validateMaxHistory(optionsOrApply.maxHistory ?? 100);
        }
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

    /** Applies a transaction as a new change and records it as one batch. */
    public execute(transaction: HistoryTransaction<T, S>): Readonly<HistoryTransaction<T, S>> | undefined {
        this.assertNotApplying();
        const normalized = normalizeTransaction(transaction);
        if (normalized === undefined) return undefined;
        this.withApplying(() => this.apply(normalized, "redo"));
        this.pushUndo(normalized);
        this.redoStack.length = 0;
        return normalized;
    }

    /** Records an already-applied transaction without invoking the callback. */
    public record(transaction: HistoryTransaction<T, S>): Readonly<HistoryTransaction<T, S>> | undefined {
        this.assertNotApplying();
        const normalized = normalizeTransaction(transaction);
        if (normalized === undefined) return undefined;
        this.pushUndo(normalized);
        this.redoStack.length = 0;
        return normalized;
    }

    /** Alias for record, useful when integrating with an editor's change-set API. */
    public push(transaction: HistoryTransaction<T, S>): Readonly<HistoryTransaction<T, S>> | undefined {
        return this.record(transaction);
    }

    /** Alias for execute for callers that model a batch as one operation. */
    public batch(transaction: HistoryTransaction<T, S>): Readonly<HistoryTransaction<T, S>> | undefined {
        return this.execute(transaction);
    }

    public undo(): Readonly<HistoryTransaction<T, S>> | undefined {
        this.assertNotApplying();
        const transaction = this.undoStack[this.undoStack.length - 1];
        if (transaction === undefined) return undefined;
        this.withApplying(() => this.apply(transaction, "undo"));
        this.undoStack.pop();
        this.pushRedo(transaction);
        return transaction;
    }

    public redo(): Readonly<HistoryTransaction<T, S>> | undefined {
        this.assertNotApplying();
        const transaction = this.redoStack[this.redoStack.length - 1];
        if (transaction === undefined) return undefined;
        this.withApplying(() => this.apply(transaction, "redo"));
        this.redoStack.pop();
        this.pushUndo(transaction);
        return transaction;
    }

    public clear(): void {
        this.assertNotApplying();
        this.undoStack.length = 0;
        this.redoStack.length = 0;
    }

    private apply(transaction: MutableHistoryTransaction<T, S>, direction: HistoryDirection): void {
        const edits = direction === "undo" ? [...transaction.edits].reverse() : transaction.edits;
        for (const edit of edits) this.applyEdit(edit, direction, transaction);
    }

    private pushUndo(transaction: MutableHistoryTransaction<T, S>): void {
        if (this.historyLimit === 0) return;
        this.undoStack.push(transaction);
        if (this.undoStack.length > this.historyLimit) this.undoStack.shift();
    }

    private pushRedo(transaction: MutableHistoryTransaction<T, S>): void {
        if (this.historyLimit === 0) return;
        this.redoStack.push(transaction);
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
        if (this.applying) throw new Error("History cannot be modified from its apply callback");
    }
}

/** Descriptive alias for consumers that prefer the longer class name. */
export { History as TransactionHistory };
