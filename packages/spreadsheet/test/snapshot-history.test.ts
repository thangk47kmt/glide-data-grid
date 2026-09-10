import { describe, expect, test, vi } from "vitest";
import { insertRows } from "../src/structural.js";
import {
    SnapshotHistory,
    snapshotTransactionFromStructural,
    type SnapshotTransaction,
} from "../src/snapshot-history.js";

const workbook = {
    columns: [{ id: "a", title: "A" }],
    rowCount: 2,
    rows: [[1], ["=A1"]],
} as const;

function change(id: string, before: number, after: number): SnapshotTransaction<number> {
    return { id, before, after };
}

describe("SnapshotHistory", () => {
    test("applies structural snapshots and restores the exact before payload", () => {
        const result = insertRows(workbook, 1, 1);
        let current = result.beforePayload;
        const history = new SnapshotHistory<typeof current>((snapshot, direction) => {
            current = snapshot;
            expect(direction).toBeDefined();
        });
        const transaction = snapshotTransactionFromStructural(result, { id: "insert-row", label: "Insert row" });
        expect(transaction.before).toBe(result.beforePayload);
        expect(transaction.after).toBe(result.afterPayload);
        history.execute(transaction);
        expect(current.rowCount).toBe(3);
        expect(current.rows[2]?.[0]).toBe("=A1");
        history.undo();
        expect(current).toBe(result.beforePayload);
        expect(current).toEqual(workbook);
        history.redo();
        expect(current).toBe(result.afterPayload);
        expect(current.columns).toEqual(result.columns);
    });

    test("supports no-op, clear, counts, and bounded history", () => {
        const history = new SnapshotHistory<number>({ apply: () => undefined, maxHistory: 2 });
        expect(history.execute(change("noop", 1, 1))).toBeUndefined();
        history.record(change("one", 0, 1));
        history.record(change("two", 1, 2));
        history.record(change("three", 2, 3));
        expect(history.undoCount).toBe(2);
        expect(history.canUndo()).toBe(true);
        history.undo();
        history.undo();
        expect(history.canUndo()).toBe(false);
        expect(history.redoCount).toBe(2);
        history.clear();
        expect(history.canUndo()).toBe(false);
        expect(history.canRedo()).toBe(false);
    });

    test("leaves stacks unchanged when apply throws", () => {
        let shouldThrow = false;
        const history = new SnapshotHistory<number>((snapshot, direction) => {
            if (shouldThrow) throw new Error(`${direction} failure`);
            void snapshot;
        });
        history.execute(change("one", 0, 1));
        history.undo();
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(1);
        shouldThrow = true;
        expect(() => history.redo()).toThrow("redo failure");
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(1);
    });

    test("rejects re-entrant mutations from the apply callback", () => {
        let history!: SnapshotHistory<number>;
        const error = vi.fn();
        history = new SnapshotHistory<number>((snapshot) => {
            try {
                history.record(change("nested", 1, 2));
            } catch (caught) {
                error(caught);
            }
            void snapshot;
        });
        history.execute(change("outer", 0, 1));
        expect(error).toHaveBeenCalledTimes(1);
        expect(history.undoCount).toBe(1);
    });

    test("normalizes transactions so later caller mutation cannot change undo or redo", () => {
        const applied: number[] = [];
        const history = new SnapshotHistory<number>(snapshot => applied.push(snapshot));
        const original = { id: "stable", before: 0, after: 1, label: "original" };
        const recorded = history.record(original);
        expect(recorded).not.toBe(original);
        original.id = "mutated";
        original.before = 10;
        original.after = 20;
        original.label = "mutated";
        history.undo();
        history.redo();
        expect(applied).toEqual([0, 1]);
        expect(recorded).toMatchObject({ id: "stable", before: 0, after: 1, label: "original" });
    });

    test("rejects malformed transactions at runtime", () => {
        const history = new SnapshotHistory<number>(() => undefined);
        expect(() => history.record(null as unknown as SnapshotTransaction<number>)).toThrow();
        expect(() => history.record({ id: 1, before: 0, after: 1 } as unknown as SnapshotTransaction<number>)).toThrow();
        expect(() => history.record({ id: "missing-before", after: 1 } as unknown as SnapshotTransaction<number>)).toThrow();
        expect(() => history.record({ id: "bad-label", before: 0, after: 1, label: 1 } as unknown as SnapshotTransaction<number>)).toThrow();
    });
});
