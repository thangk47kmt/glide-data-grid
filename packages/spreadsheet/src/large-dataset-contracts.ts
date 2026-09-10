import { translateFormula, type FormulaError } from "./formula.js";

export interface PasteRectangle {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
}

export interface RememberedPasteRange extends PasteRectangle {
    readonly sourceRows: readonly number[];
    readonly values: readonly (readonly string[])[];
}

export interface PlannedPasteEdit<T> {
    readonly col: number;
    readonly row: number;
    readonly before: T;
    readonly after: T;
}

export type PastePlan<T> =
    | { readonly valid: true; readonly edits: readonly PlannedPasteEdit<T>[] }
    | { readonly valid: false; readonly edits: readonly []; readonly col: number; readonly row: number; readonly value: string; readonly message: string };

export interface PlanGridPasteOptions<T> {
    readonly target: readonly [number, number];
    readonly selection?: PasteRectangle;
    readonly values: readonly (readonly string[])[];
    readonly pageRows: readonly number[];
    readonly columnCount: number;
    readonly rememberedCopy?: RememberedPasteRange;
    readonly canWrite?: (col: number) => boolean;
    readonly getBefore: (col: number, row: number) => T;
    readonly parse: (value: string, col: number) => T;
    readonly validateFormula: (col: number, row: number, formula: string) => FormulaError | undefined;
}

export function formulaValidationMessage(error: FormulaError | undefined): string {
    return error === undefined ? "Formula is invalid" : `${error.code}: ${error.message ?? "Formula is invalid"}`;
}

/**
 * Plans the complete clipboard operation without mutating the data source.
 * Callers apply the returned edits only when `valid` is true, making a paste
 * atomic even when a later destination contains an invalid formula.
 */
export function planGridPaste<T>(options: PlanGridPasteOptions<T>): PastePlan<T> {
    const { target, selection, values, pageRows, columnCount, rememberedCopy } = options;
    if (values.length === 0) return { valid: true, edits: [] };
    const sourceHeight = values.length;
    const sourceWidth = Math.max(0, ...values.map(row => row.length));
    if (sourceWidth === 0) return { valid: true, edits: [] };

    const targetCol = selection?.x ?? target[0];
    const targetGridRow = selection?.y ?? target[1];
    const hasSelectedRange = (selection?.width ?? 1) > 1 || (selection?.height ?? 1) > 1;
    const targetWidth = hasSelectedRange ? Math.max(1, selection?.width ?? 1) : sourceWidth;
    const targetHeight = hasSelectedRange ? Math.max(1, selection?.height ?? 1) : sourceHeight;
    const tile = (targetWidth > sourceWidth || targetHeight > sourceHeight)
        && targetWidth % sourceWidth === 0
        && targetHeight % sourceHeight === 0;
    const copyOrigin = rememberedCopy !== undefined
        && rememberedCopy.height === sourceHeight
        && rememberedCopy.width === sourceWidth
        && rememberedCopy.values.every((row, rowIndex) => row.every((value, colIndex) => value === values[rowIndex]?.[colIndex]))
        ? rememberedCopy
        : undefined;
    const edits: PlannedPasteEdit<T>[] = [];
    const seen = new Set<string>();

    for (let rowOffset = 0; rowOffset < (tile ? targetHeight : Math.min(targetHeight, sourceHeight)); rowOffset++) {
        const sourceRowOffset = tile ? rowOffset % sourceHeight : rowOffset;
        const dataRow = values[sourceRowOffset] ?? [];
        const gridRow = targetGridRow + rowOffset;
        const row = pageRows[gridRow];
        if (row === undefined) break;
        for (let colOffset = 0; colOffset < (tile ? targetWidth : Math.min(targetWidth, sourceWidth)); colOffset++) {
            const sourceColOffset = tile ? colOffset % sourceWidth : colOffset;
            const value = dataRow[sourceColOffset];
            if (value === undefined) continue;
            const col = targetCol + colOffset;
            if (col < 0 || col >= columnCount || options.canWrite?.(col) === false) continue;
            const sourceGridCol = (copyOrigin?.x ?? targetCol) + sourceColOffset;
            const sourceGridRow = (copyOrigin?.y ?? targetGridRow) + sourceRowOffset;
            const sourceSourceRow = copyOrigin === undefined ? pageRows[sourceGridRow] : copyOrigin.sourceRows[sourceRowOffset];
            const translated = value.startsWith("=") && copyOrigin !== undefined && sourceSourceRow !== undefined
                ? translateFormula(value, col - sourceGridCol, row - sourceSourceRow)
                : value;
            const key = `${col}:${row}`;
            if (seen.has(key)) continue;
            seen.add(key);
            const after = options.parse(translated, col);
            if (typeof after === "string" && after.startsWith("=")) {
                const error = options.validateFormula(col, row, after);
                if (error !== undefined) {
                    return { valid: false, edits: [], col, row, value: after, message: formulaValidationMessage(error) };
                }
            }
            edits.push({ col, row, before: options.getBefore(col, row), after });
        }
    }
    return { valid: true, edits };
}

export function dropdownGridCell(
    value: string | null,
    allowedValues: readonly string[],
    copyData: string,
    themeOverride?: Readonly<Record<string, string>>,
) {
    return {
        kind: "custom" as const,
        allowOverlay: true,
        copyData,
        data: { kind: "dropdown-cell", value, allowedValues },
        themeOverride,
    };
}

/** Custom/native editors own double-click so their overlay can open. */
export function shouldOpenInlineFormulaEditorOnDoubleClick(isDoubleClick: boolean, dataType: string | undefined): boolean {
    return isDoubleClick && dataType !== "dropdown" && dataType !== "drilldown" && dataType !== "multi-select";
}
