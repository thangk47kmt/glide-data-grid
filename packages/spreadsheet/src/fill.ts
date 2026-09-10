import { translateFormula } from "./formula.js";
import type { CellInput } from "./model.js";

/** A rectangular region in zero-based grid coordinates. */
export interface FillRectangle {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
}

/** An input change represented as [column, row, value]. */
export type FillEdit = readonly [col: number, row: number, value: CellInput];

type Series = { readonly kind: "number"; readonly stepX: number; readonly stepY: number; readonly origin: number } | {
    readonly kind: "date";
    readonly stepX: number;
    readonly stepY: number;
    readonly origin: number;
};

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;

function isInteger(value: number): boolean {
    return Number.isInteger(value);
}

function validateRectangle(rectangle: FillRectangle, name: string): void {
    if (
        !isInteger(rectangle.x) ||
        !isInteger(rectangle.y) ||
        !isInteger(rectangle.width) ||
        !isInteger(rectangle.height) ||
        rectangle.x < 0 ||
        rectangle.y < 0 ||
        rectangle.width <= 0 ||
        rectangle.height <= 0
    ) {
        throw new RangeError(`${name} must have non-negative integer coordinates and positive integer dimensions`);
    }
}

function dateToDay(value: string): number | undefined {
    const match = value.match(ISO_DATE);
    if (match === null) return undefined;
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const timestamp = Date.UTC(year, month - 1, day);
    const date = new Date(timestamp);
    // Date.UTC normalizes invalid dates (for example, February 30), so reject
    // those rather than silently producing a surprising series.
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return undefined;
    return Math.floor(timestamp / 86_400_000);
}

function dayToDate(day: number): string {
    const date = new Date(day * 86_400_000);
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    const dateOfMonth = String(date.getUTCDate()).padStart(2, "0");
    return `${String(year).padStart(4, "0")}-${month}-${dateOfMonth}`;
}

function firstStep(values: readonly (number | undefined)[]): number {
    if (values.length < 2) return 0;
    const first = values[0];
    const second = values[1];
    return first !== undefined && second !== undefined ? second - first : 0;
}

function inferSeries(pattern: readonly (readonly CellInput[])[], width: number, height: number): Series | undefined {
    const values = pattern.flat();
    if (values.length < 2) return undefined;

    if (values.every(value => typeof value === "number" && Number.isFinite(value))) {
        const numeric = values as readonly number[];
        const stepX = width >= 2 ? firstStep(pattern[0] as readonly (number | undefined)[]) : 0;
        const firstColumn = pattern.map(row => row[0] as number | undefined);
        const stepY = height >= 2 ? firstStep(firstColumn) : 0;
        return { kind: "number", origin: numeric[0], stepX, stepY };
    }

    if (values.every(value => typeof value === "string" && dateToDay(value) !== undefined)) {
        const days = pattern.map(row => row.map(value => dateToDay(value as string) as number));
        const stepX = width >= 2 ? firstStep(days[0]) : 0;
        const stepY = height >= 2 ? firstStep(days.map(row => row[0])) : 0;
        return { kind: "date", origin: days[0][0], stepX, stepY };
    }

    return undefined;
}

function patternValue(pattern: readonly (readonly CellInput[])[], source: FillRectangle, series: Series | undefined, col: number, row: number, patternCol: number, patternRow: number): CellInput {
    const value = pattern[patternRow][patternCol];
    if (typeof value === "string" && value.startsWith("=")) {
        return translateFormula(value, col - (source.x + patternCol), row - (source.y + patternRow));
    }
    if (series === undefined) return value;

    // Series values are continuous across pattern boundaries. Using the
    // source origin here (rather than the repeated cell's origin) avoids
    // restarting a two-cell series at every pattern tile.
    const distanceX = col - source.x;
    const distanceY = row - source.y;
    const seriesValue = series.origin + series.stepX * distanceX + series.stepY * distanceY;
    return series.kind === "date" ? dayToDate(seriesValue) : seriesValue;
}

/**
 * Produces edits for filling destination from pattern. The pattern dimensions
 * must equal source dimensions. Destination is iterated from top-left and
 * repeats the pattern; formulas are translated from their corresponding source
 * cell. Cells in the source rectangle are omitted from the edits.
 *
 * Numeric and ISO-date patterns use the first two values in each available axis
 * as linear steps. For a two-dimensional pattern both steps are applied, with
 * horizontal step taking precedence only when a pattern has one row/column.
 */
export function smartFill(
    pattern: readonly (readonly CellInput[])[],
    source: FillRectangle,
    destination: FillRectangle
): FillEdit[] {
    validateRectangle(source, "source");
    validateRectangle(destination, "destination");
    if (pattern.length !== source.height || pattern.length === 0) {
        throw new RangeError("pattern height must equal source height");
    }
    if (pattern.some(row => row.length !== source.width)) {
        throw new RangeError("pattern must be rectangular and match source width");
    }

    const series = inferSeries(pattern, source.width, source.height);
    const edits: FillEdit[] = [];
    for (let row = destination.y; row < destination.y + destination.height; row++) {
        for (let col = destination.x; col < destination.x + destination.width; col++) {
            if (
                col >= source.x &&
                col < source.x + source.width &&
                row >= source.y &&
                row < source.y + source.height
            ) continue;
            const patternCol = (col - destination.x) % source.width;
            const patternRow = (row - destination.y) % source.height;
            edits.push([col, row, patternValue(pattern, source, series, col, row, patternCol, patternRow)]);
        }
    }
    return edits;
}

/** Alias emphasizing that the operation fills a destination rectangle. */
export const fillPattern = smartFill;
