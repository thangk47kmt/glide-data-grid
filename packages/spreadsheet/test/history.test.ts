import { describe, expect, it, vi } from "vitest";
import { History, type CellEdit, type HistoryTransaction } from "../src/history.js";

function transaction(id: string, edits: readonly CellEdit<number>[], extra: Partial<HistoryTransaction<number>> = {}): HistoryTransaction<number> {
    return { id, edits, ...extra };
}

describe("History", () => {
    it("executes, undoes, and redoes a transaction", () => {
        const values = new Map<string, number>();
        const history = new History<number>({
            apply: (edit, direction) => values.set(`${edit.location[0]}:${edit.location[1]}`, direction === "undo" ? edit.before : edit.after),
        });
        const change = transaction("1", [{ location: [0, 0], before: 1, after: 2 }]);

        history.execute(change);
        expect(values.get("0:0")).toBe(2);
        expect(history.canUndo()).toBe(true);
        expect(history.canRedo()).toBe(false);
        history.undo();
        expect(values.get("0:0")).toBe(1);
        expect(history.canUndo()).toBe(false);
        expect(history.canRedo()).toBe(true);
        history.redo();
        expect(values.get("0:0")).toBe(2);
    });

    it("applies a multi-cell transaction as one batch and undoes in reverse order", () => {
        const applied: string[] = [];
        const history = new History<number>((edit, direction) => applied.push(`${direction}:${edit.location[0]}`));
        history.execute(
            transaction("batch", [
                { location: [0, 0], before: 0, after: 1 },
                { location: [1, 0], before: 0, after: 1 },
                { location: [2, 0], before: 0, after: 1 },
            ])
        );
        history.undo();
        history.redo();
        expect(applied).toEqual(["redo:0", "redo:1", "redo:2", "undo:2", "undo:1", "undo:0", "redo:0", "redo:1", "redo:2"]);
        expect(history.undoCount).toBe(1);
    });

    it("coalesces duplicate cells using the first before and last after", () => {
        const edits: CellEdit<number>[] = [];
        const history = new History<number>((edit, direction) => {
            if (direction === "redo") edits.push(edit);
        });
        const result = history.execute(
            transaction("duplicates", [
                { location: [1, 2], before: 3, after: 4 },
                { location: [0, 0], before: 8, after: 9 },
                { location: [1, 2], before: 4, after: 7 },
            ])
        );
        expect(result?.edits).toEqual([
            { location: [1, 2], before: 3, after: 7 },
            { location: [0, 0], before: 8, after: 9 },
        ]);
        expect(edits).toEqual(result?.edits);
    });

    it("does not record no-op edits", () => {
        const apply = vi.fn();
        const history = new History<number>(apply);
        expect(history.execute(transaction("noop", [{ location: [0, 0], before: 1, after: 1 }]))).toBeUndefined();
        expect(apply).not.toHaveBeenCalled();
        expect(history.canUndo()).toBe(false);
        history.record(transaction("mixed", [{ location: [0, 0], before: 1, after: 1 }, { location: [1, 0], before: 1, after: 2 }]));
        expect(history.undoCount).toBe(1);
    });

    it("clears redo after recording a new transaction", () => {
        const history = new History<number>(() => undefined);
        history.execute(transaction("one", [{ location: [0, 0], before: 0, after: 1 }]));
        history.undo();
        expect(history.canRedo()).toBe(true);
        history.record(transaction("two", [{ location: [0, 0], before: 0, after: 2 }]));
        expect(history.canRedo()).toBe(false);
    });

    it("bounds undo and redo history", () => {
        const history = new History<number>({ apply: () => undefined, maxHistory: 2 });
        for (let index = 0; index < 3; index++) {
            history.record(transaction(String(index), [{ location: [index, 0], before: 0, after: 1 }]));
        }
        expect(history.undoCount).toBe(2);
        history.undo();
        history.undo();
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(2);
        history.redo();
        history.redo();
        expect(history.redoCount).toBe(0);
    });

    it("retains selection metadata and rejects re-entrant recording", () => {
        const selectionBefore = { cell: [0, 0] };
        const selectionAfter = { cell: [1, 0] };
        let history!: History<number>;
        let reentrantError: unknown;
        history = new History<number>((edit, direction, tx) => {
            expect(tx.selectionBefore).toBe(selectionBefore);
            expect(tx.selectionAfter).toBe(selectionAfter);
            expect(edit.location).toEqual([0, 0]);
            expect(direction).toBe("redo");
            try {
                history.record(transaction("nested", [{ location: [1, 0], before: 0, after: 1 }]));
            } catch (error) {
                reentrantError = error;
            }
        });
        const result = history.execute(
            transaction("selection", [{ location: [0, 0], before: 0, after: 1 }], { selectionBefore, selectionAfter })
        );
        expect(result?.selectionBefore).toBe(selectionBefore);
        expect(result?.selectionAfter).toBe(selectionAfter);
        expect(reentrantError).toBeInstanceOf(Error);
        expect(history.undoCount).toBe(1);
    });

    it("can clear all entries", () => {
        const history = new History<number>(() => undefined);
        history.record(transaction("one", [{ location: [0, 0], before: 0, after: 1 }]));
        history.undo();
        history.clear();
        expect(history.canUndo()).toBe(false);
        expect(history.canRedo()).toBe(false);
    });

    it("does not record or clear redo when execute apply throws", () => {
        let fail = false;
        const history = new History<number>((edit, direction) => {
            if (fail) throw new Error("apply failed");
            void edit;
            void direction;
        });
        history.execute(transaction("initial", [{ location: [0, 0], before: 0, after: 1 }]));
        history.undo();
        expect(history.redoCount).toBe(1);
        fail = true;
        expect(() => history.execute(transaction("failed", [{ location: [1, 0], before: 0, after: 1 }]))).toThrow("apply failed");
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(1);
    });

    it("keeps undo transaction when undo apply throws", () => {
        let fail = false;
        const history = new History<number>((edit, direction) => {
            if (fail && direction === "undo") throw new Error("undo failed");
            void edit;
        });
        history.execute(transaction("undo-failure", [{ location: [0, 0], before: 0, after: 1 }]));
        fail = true;
        expect(() => history.undo()).toThrow("undo failed");
        expect(history.undoCount).toBe(1);
        expect(history.redoCount).toBe(0);
    });

    it("keeps redo transaction when redo apply throws", () => {
        let fail = false;
        const history = new History<number>((edit, direction) => {
            if (fail && direction === "redo") throw new Error("redo failed");
            void edit;
        });
        history.execute(transaction("redo-failure", [{ location: [0, 0], before: 0, after: 1 }]));
        history.undo();
        fail = true;
        expect(() => history.redo()).toThrow("redo failed");
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(1);
    });
});
