export interface FormulaTextRange {
    readonly start: number;
    readonly end: number;
}

export interface FormulaReferenceInsertion {
    readonly formula: string;
    readonly cursor: number;
    readonly referenceRange: FormulaTextRange;
}

/** A zero-based grid coordinate used when formatting A1 references. */
export interface FormulaReferenceCell {
    readonly col: number;
    readonly row: number;
}

/** A rectangular grid selection. The two corners may be supplied in either order. */
export interface FormulaReferenceRange {
    readonly from: FormulaReferenceCell;
    readonly to: FormulaReferenceCell;
}

/** A whole column represented by its structured-reference caption. */
export interface FormulaReferenceColumn {
    readonly columnName: string;
    /** Use Excel table current-row syntax (`[@Caption]`) when true. */
    readonly currentRow?: boolean;
}

/** A caption-based cell reference with an explicit one-based row in formula text. */
export interface FormulaReferenceCaptionCell {
    readonly columnName: string;
    readonly row: number;
}

export type FormulaReferenceTarget =
    | ({ readonly kind: "cell" } & FormulaReferenceCell)
    | ({ readonly kind: "range" } & FormulaReferenceRange)
    | ({ readonly kind: "column" } & FormulaReferenceColumn)
    | ({ readonly kind: "caption-cell" } & FormulaReferenceCaptionCell);

function validCoordinate(value: FormulaReferenceCell): boolean {
    return Number.isSafeInteger(value.col) && value.col >= 0 && Number.isSafeInteger(value.row) && value.row >= 0;
}

function formatA1Cell(value: FormulaReferenceCell): string {
    // Importing the formula parser's formatter here would create a circular
    // dependency for consumers that only use authoring helpers. This local
    // conversion is intentionally kept small and mirrors columnIndexToName.
    let column = "";
    for (let index = value.col + 1; index > 0; index = Math.floor((index - 1) / 26)) {
        column = String.fromCharCode(65 + ((index - 1) % 26)) + column;
    }
    return `${column}${value.row + 1}`;
}

/** Formats a grid target as an Excel-compatible cell, range, or structured column reference. */
export function formatFormulaReference(target: FormulaReferenceTarget): string {
    if (target.kind === "caption-cell") {
        const columnName = target.columnName.trim();
        if (columnName.length === 0 || columnName.includes("]")) throw new RangeError("A structured reference column name is required");
        if (!Number.isSafeInteger(target.row) || target.row < 0) throw new RangeError("A caption cell reference row must be a non-negative safe integer");
        return `[${columnName}]${target.row + 1}`;
    }
    if (target.kind === "column") {
        const columnName = target.columnName.trim();
        if (columnName.length === 0 || columnName.includes("]")) throw new RangeError("A structured reference column name is required");
        return `[${target.currentRow === true ? "@" : ""}${columnName}]`;
    }
    if (target.kind === "cell") {
        if (!validCoordinate(target)) throw new RangeError("A cell reference coordinate must be a non-negative safe integer");
        return formatA1Cell(target);
    }
    if (!validCoordinate(target.from) || !validCoordinate(target.to)) throw new RangeError("A range reference coordinate must be a non-negative safe integer");
    const from = {
        col: Math.min(target.from.col, target.to.col),
        row: Math.min(target.from.row, target.to.row),
    };
    const to = {
        col: Math.max(target.from.col, target.to.col),
        row: Math.max(target.from.row, target.to.row),
    };
    return `${formatA1Cell(from)}:${formatA1Cell(to)}`;
}

/**
 * Inserts a cell/range/structured reference into a formula draft. While the
 * user is dragging a selection, pass the previous referenceRange so the live
 * preview is replaced instead of appending one reference per mouse move.
 */
export function insertFormulaReference(
    formula: string,
    reference: string,
    selection: FormulaTextRange,
    previousReference?: FormulaTextRange
): FormulaReferenceInsertion {
    const active = previousReference ?? selection;
    const start = Math.max(0, Math.min(formula.length, Math.min(active.start, active.end)));
    const end = Math.max(start, Math.min(formula.length, Math.max(active.start, active.end)));
    const next = `${formula.slice(0, start)}${reference}${formula.slice(end)}`;
    const cursor = start + reference.length;
    return { formula: next, cursor, referenceRange: { start, end: cursor } };
}

/**
 * Formats and inserts a grid target in one operation. Keeping target
 * formatting beside insertion makes direct cell editing and formula-bar
 * authoring use the same caret/selection semantics.
 */
export function insertFormulaReferenceTarget(
    formula: string,
    target: FormulaReferenceTarget,
    selection: FormulaTextRange,
    previousReference?: FormulaTextRange
): FormulaReferenceInsertion {
    return insertFormulaReference(formula, formatFormulaReference(target), selection, previousReference);
}
