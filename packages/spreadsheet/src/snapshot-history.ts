import type { StructuralEditResult } from "./structural.js";

export type SnapshotDirection = "redo" | "undo";

export interface SnapshotTransaction<T, S = object> {
    readonly id: string;
    readonly label?: string;
    readonly before: T;
    readonly after: T;
    readonly selectionBefore?: S;
    readonly selectionAfter?: S;
}

export type SnapshotApply<T, S = object> = (
    snapshot: T,
    direction: SnapshotDirection,
    transaction: Readonly<SnapshotTransaction<T, S>>
) => void;

export interface SnapshotHistoryOptions<T, S = object> {
    readonly apply: SnapshotApply<T, S>;
    readonly maxHistory?: number;
}

function validateMaxHistory(value: number): number {
    if (!Number.isInteger(value) || value < 0) throw new RangeError("maxHistory must be a non-negative integer");
    return value;
}

function normalizeTransaction<T, S>(transaction: SnapshotTransaction<T, S>): SnapshotTransaction<T, S> {
    if (transaction === null || typeof transaction !== "object") throw new TypeError("A snapshot transaction is required");
    if (typeof transaction.id !== "string") throw new TypeError("A snapshot transaction id must be a string");
    if (!("before" in transaction) || !("after" in transaction)) throw new TypeError("A snapshot transaction requires before and after snapshots");
    if (transaction.label !== undefined && typeof transaction.label !== "string") throw new TypeError("A snapshot transaction label must be a string");
    // This shallow copy protects history metadata and snapshot selection from
    // mutation of the caller's transaction object. Snapshot values themselves
    // are intentionally not cloned; callers must keep them immutable.
    return {
        id: transaction.id,
        ...(transaction.label === undefined ? {} : { label: transaction.label }),
        before: transaction.before,
        after: transaction.after,
        ...(transaction.selectionBefore === undefined ? {} : { selectionBefore: transaction.selectionBefore }),
        ...(transaction.selectionAfter === undefined ? {} : { selectionAfter: transaction.selectionAfter }),
    };
}

/**
 * Converts a structural result into an atomic snapshot transaction. The
 * before/after payload references are reused directly; no large workbook is
 * copied by this helper or by SnapshotHistory. Callers must treat snapshots as
 * immutable after handing them to the history.
 */
export function snapshotTransactionFromStructural<S = object>(
    result: StructuralEditResult,
    options: { readonly id: string; readonly label?: string; readonly selectionBefore?: S; readonly selectionAfter?: S }
): SnapshotTransaction<StructuralEditResult["payload"], S> {
    return {
        id: options.id,
        ...(options.label === undefined ? {} : { label: options.label }),
        before: result.beforePayload,
        after: result.afterPayload,
        ...(options.selectionBefore === undefined ? {} : { selectionBefore: options.selectionBefore }),
        ...(options.selectionAfter === undefined ? {} : { selectionAfter: options.selectionAfter }),
    };
}

/** Short alias for callers using “structural result” terminology. */
export const structuralResultToSnapshotTransaction = snapshotTransactionFromStructural;

type StoredTransaction<T, S> = SnapshotTransaction<T, S>;

/**
 * Bounded, model-independent history for replacing immutable snapshots. The
 * apply callback owns the actual workbook state and receives the selected
 * snapshot by reference. Recording from an apply callback is rejected. If an
 * apply callback throws, stacks are unchanged; external side effects from any
 * earlier callback invocation cannot be rolled back by this class.
 */
export class SnapshotHistory<T, S = object> {
    private readonly applySnapshot: SnapshotApply<T, S>;
    private readonly undoStack: Array<StoredTransaction<T, S>> = [];
    private readonly redoStack: Array<StoredTransaction<T, S>> = [];
    private historyLimit: number;
    private applying = false;

    public constructor(options: SnapshotHistoryOptions<T, S>);
    public constructor(apply: SnapshotApply<T, S>, options?: { readonly maxHistory?: number } | number);
    public constructor(
        optionsOrApply: SnapshotHistoryOptions<T, S> | SnapshotApply<T, S>,
        options: { readonly maxHistory?: number } | number = {}
    ) {
        if (typeof optionsOrApply === "function") {
            this.applySnapshot = optionsOrApply;
            this.historyLimit = validateMaxHistory(typeof options === "number" ? options : options.maxHistory ?? 100);
        } else {
            this.applySnapshot = optionsOrApply.apply;
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

    public execute(transaction: SnapshotTransaction<T, S>): Readonly<SnapshotTransaction<T, S>> | undefined {
        this.assertNotApplying();
        const normalized = normalizeTransaction(transaction);
        if (Object.is(normalized.before, normalized.after)) return undefined;
        this.withApplying(() => this.applySnapshot(normalized.after, "redo", normalized));
        this.pushUndo(normalized);
        this.redoStack.length = 0;
        return normalized;
    }

    /** Records an already-applied snapshot replacement without invoking apply. */
    public record(transaction: SnapshotTransaction<T, S>): Readonly<SnapshotTransaction<T, S>> | undefined {
        this.assertNotApplying();
        const normalized = normalizeTransaction(transaction);
        if (Object.is(normalized.before, normalized.after)) return undefined;
        this.pushUndo(normalized);
        this.redoStack.length = 0;
        return normalized;
    }

    public push(transaction: SnapshotTransaction<T, S>): Readonly<SnapshotTransaction<T, S>> | undefined {
        return this.record(transaction);
    }

    public undo(): Readonly<SnapshotTransaction<T, S>> | undefined {
        this.assertNotApplying();
        const transaction = this.undoStack[this.undoStack.length - 1];
        if (transaction === undefined) return undefined;
        this.withApplying(() => this.applySnapshot(transaction.before, "undo", transaction));
        this.undoStack.pop();
        this.pushRedo(transaction);
        return transaction;
    }

    public redo(): Readonly<SnapshotTransaction<T, S>> | undefined {
        this.assertNotApplying();
        const transaction = this.redoStack[this.redoStack.length - 1];
        if (transaction === undefined) return undefined;
        this.withApplying(() => this.applySnapshot(transaction.after, "redo", transaction));
        this.redoStack.pop();
        this.pushUndo(transaction);
        return transaction;
    }

    public clear(): void {
        this.assertNotApplying();
        this.undoStack.length = 0;
        this.redoStack.length = 0;
    }

    private pushUndo(transaction: StoredTransaction<T, S>): void {
        if (this.historyLimit === 0) return;
        this.undoStack.push(transaction);
        if (this.undoStack.length > this.historyLimit) this.undoStack.shift();
    }

    private pushRedo(transaction: StoredTransaction<T, S>): void {
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
        if (this.applying) throw new Error("Snapshot history cannot be modified from its apply callback");
    }
}
