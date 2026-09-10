import { describe, expect, test, vi } from "vitest";
import { insertRows } from "../src/structural.js";
import { cellTransactionToCommand, snapshotTransactionToCommand, structuralResultToCommand, WorkbookHistory, type WorkbookCommand } from "../src/workbook-history.js";
import type { HistoryTransaction } from "../src/history.js";

const cells: HistoryTransaction<number> = {
    id: "cell-1",
    edits: [{ location: [0, 0], before: 0, after: 1 }],
};

describe("WorkbookHistory", () => {
    test("undoes and redoes interleaved cell and structural commands in one order", () => {
        const events: string[] = [];
        let value = 0;
        const cellCommand = cellTransactionToCommand(cells, (edit, direction) => {
            value = direction === "undo" ? edit.before : edit.after;
            events.push(`${direction}:cell`);
        });
        const before = { columns: [{ id: "a", title: "A" }], rowCount: 1, rows: [[value]] };
        const structural = insertRows(before, 1, 1);
        const structuralCommand = structuralResultToCommand(structural, { id: "row-1", label: "Insert row" }, snapshot => {
            events.push(`snapshot:${snapshot.rowCount}`);
        });
        const second = snapshotTransactionToCommand<number, object>({ id: "cell-2", before: 1, after: 2 }, snapshot => {
            value = snapshot;
            events.push(`snapshot-cell:${snapshot}`);
        });
        const history = new WorkbookHistory();
        history.execute(cellCommand as WorkbookCommand);
        history.execute(structuralCommand as WorkbookCommand);
        history.execute(second as WorkbookCommand);
        history.undo();
        history.undo();
        history.undo();
        history.redo();
        history.redo();
        history.redo();
        expect(events).toEqual([
            "redo:cell",
            "snapshot:2",
            "snapshot-cell:2",
            "snapshot-cell:1",
            "snapshot:1",
            "undo:cell",
            "redo:cell",
            "snapshot:2",
            "snapshot-cell:2",
        ]);
        expect(value).toBe(2);
    });

    test("record does not invoke callbacks and a new command clears redo", () => {
        const apply = vi.fn();
        const history = new WorkbookHistory();
        const command: WorkbookCommand = { id: "recorded", apply, revert: apply };
        history.record(command);
        expect(apply).not.toHaveBeenCalled();
        history.undo();
        expect(apply).toHaveBeenCalledTimes(1);
        history.record({ id: "new", apply, revert: apply });
        expect(history.canRedo()).toBe(false);
    });

    test("keeps stacks unchanged when command callback throws", () => {
        let fail = false;
        const history = new WorkbookHistory();
        history.execute({ id: "ok", apply: () => undefined, revert: () => { if (fail) throw new Error("revert failed"); } });
        fail = true;
        expect(() => history.undo()).toThrow("revert failed");
        expect(history.undoCount).toBe(1);
        expect(history.redoCount).toBe(0);
    });

    test("coalesces duplicate cell edits and removes no-op edits", () => {
        const edits: number[] = [];
        const command = cellTransactionToCommand({
            id: "coalesce",
            edits: [
                { location: [0, 0], before: 0, after: 1 },
                { location: [0, 0], before: 1, after: 3 },
                { location: [1, 0], before: 2, after: 2 },
            ],
        }, edit => edits.push(edit.after));
        const history = new WorkbookHistory();
        const normalized = history.execute(command);
        if (normalized === undefined) throw new Error("Expected command to be recorded");
        expect(normalized).toBeDefined();
        expect(edits).toEqual([3]);
    });

    test("compensates completed cell edits when a later edit fails", () => {
        const state = new Map<number, number>([[0, 0], [1, 0]]);
        const command = cellTransactionToCommand({
            id: "partial",
            edits: [
                { location: [0, 0], before: 0, after: 1 },
                { location: [1, 0], before: 0, after: 2 },
            ],
        }, (edit, direction) => {
            if (direction === "redo" && edit.location[0] === 1) throw new Error("second edit failed");
            state.set(edit.location[0], direction === "undo" ? edit.before : edit.after);
        });
        const history = new WorkbookHistory();
        expect(() => history.execute(command)).toThrow("second edit failed");
        expect(state).toEqual(new Map([[0, 0], [1, 0]]));
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(0);
    });

    test("compensates a failed undo in forward order", () => {
        const state = new Map<number, number>([[0, 1], [1, 1], [2, 1]]);
        const events: string[] = [];
        const command = cellTransactionToCommand({
            id: "undo-order",
            edits: [
                { location: [0, 0], before: 0, after: 1 },
                { location: [1, 0], before: 0, after: 1 },
                { location: [2, 0], before: 0, after: 1 },
            ],
        }, (edit, direction) => {
            const column = edit.location[0];
            events.push(`${direction}:${column}`);
            if (direction === "undo" && column === 0) throw new Error("first edit failed");
            state.set(column, direction === "undo" ? edit.before : edit.after);
        });
        const history = new WorkbookHistory();
        history.record(command);
        expect(() => history.undo()).toThrow("first edit failed");
        expect(events).toEqual(["undo:2", "undo:1", "undo:0", "redo:1", "redo:2"]);
        expect(state).toEqual(new Map([[0, 1], [1, 1], [2, 1]]));
        expect(history.undoCount).toBe(1);
        expect(history.redoCount).toBe(0);
    });

    test("reports original and compensation failures as AggregateError", () => {
        const command = cellTransactionToCommand({
            id: "compensation-failure",
            edits: [
                { location: [0, 0], before: 0, after: 1 },
                { location: [1, 0], before: 0, after: 1 },
            ],
        }, (edit, direction) => {
            if (direction === "redo" && edit.location[0] === 1) throw new Error("original failure");
            if (direction === "undo" && edit.location[0] === 0) throw new Error("compensation failure");
        });
        const history = new WorkbookHistory();
        let caught: unknown;
        try {
            history.execute(command);
        } catch (error) {
            caught = error;
        }
        expect(caught).toBeInstanceOf(AggregateError);
        expect((caught as AggregateError).errors).toHaveLength(2);
        expect(history.undoCount).toBe(0);
        expect(history.redoCount).toBe(0);
    });

    test("normalizes command metadata and bounds capacity", () => {
        const history = new WorkbookHistory({ maxHistory: 1 });
        const original = { id: "stable", label: "first", apply: () => undefined, revert: () => undefined };
        const normalized = history.record(original);
        original.id = "changed";
        original.label = "changed";
        expect(normalized?.id).toBe("stable");
        history.record({ id: "second", apply: () => undefined, revert: () => undefined });
        expect(history.undoCount).toBe(1);
        expect(history.undo()?.id).toBe("second");
        expect(() => history.record({ id: 1, apply: () => undefined, revert: () => undefined } as unknown as WorkbookCommand)).toThrow();
    });

    test("validates adapter callbacks and transaction labels", () => {
        expect(() => cellTransactionToCommand({ id: "bad", edits: [] }, null as never)).toThrow();
        expect(() => cellTransactionToCommand({ id: "bad", label: 1, edits: [] } as never, () => undefined)).toThrow();
        expect(() => snapshotTransactionToCommand({ id: "bad", label: 1, before: 0, after: 1 } as never, null as never)).toThrow();
    });
});
