import { createAggregateAccumulator, type AggregateResult } from "./aggregate.js";
import type { FormulaValue } from "./formula.js";

/** Maximum number of cells evaluated synchronously for a selection summary. */
export const DEFAULT_SELECTION_STATS_MAX_CELLS = 100_000;

export interface SelectionStatsRange {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
}

export type SelectionStatsResult =
    | {
          readonly kind: "empty";
          readonly cellCount: 0;
      }
    | {
          readonly kind: "too-large";
          readonly cellCount: number;
          readonly maxCells: number;
      }
    | {
          readonly kind: "ready";
          readonly cellCount: number;
          readonly aggregate: AggregateResult;
      };

/**
 * Aggregates the current rectangular grid range without assuming display rows
 * are source rows. The row map is therefore applied before reading each value.
 * Values are streamed only after the bounded-size check so a large selection
 * does not allocate a second array proportional to the selected area.
 */
export function aggregateSelectionRange(
    range: SelectionStatsRange | undefined,
    rowMap: readonly (number | undefined)[],
    columnCount: number,
    getValue: (column: number, sourceRow: number) => FormulaValue,
    maxCells = DEFAULT_SELECTION_STATS_MAX_CELLS,
): SelectionStatsResult {
    if (
        range === undefined ||
        !Number.isFinite(range.x) ||
        !Number.isFinite(range.y) ||
        !Number.isFinite(range.width) ||
        !Number.isFinite(range.height) ||
        !Number.isSafeInteger(columnCount) ||
        columnCount <= 0 ||
        !Number.isSafeInteger(maxCells) ||
        maxCells <= 0
    ) {
        return { kind: "empty", cellCount: 0 };
    }

    const startColumn = Math.max(0, Math.trunc(range.x));
    const startRow = Math.max(0, Math.trunc(range.y));
    const endColumn = Math.min(columnCount, Math.max(startColumn, Math.trunc(range.x + range.width)));
    const endRow = Math.min(rowMap.length, Math.max(startRow, Math.trunc(range.y + range.height)));
    const width = Math.max(0, endColumn - startColumn);
    const height = Math.max(0, endRow - startRow);
    const cellCount = width * height;

    if (cellCount === 0) return { kind: "empty", cellCount: 0 };
    if (cellCount > maxCells) return { kind: "too-large", cellCount, maxCells };

    const accumulator = createAggregateAccumulator();
    for (let row = startRow; row < endRow; row++) {
        const sourceRow = rowMap[row];
        if (sourceRow === undefined) continue;
        for (let column = startColumn; column < endColumn; column++) {
            accumulator.add(getValue(column, sourceRow));
        }
    }
    return { kind: "ready", cellCount, aggregate: accumulator.finalize() };
}
