import * as React from "react";
import {
    CompactSelection,
    DataEditor,
    GridCellKind,
    type GridSelection,
    type GridCell,
    type GridColumn,
    type Item,
    type EditableGridCell,
    type GridKeyEventArgs,
    type Rectangle,
} from "../../core/src/index.js";
import { allCells } from "../../cells/src/index.js";
import { LOADING_CELL, PagedDataSource, type PagedCellValue } from "./paged-data-source.js";
import { PagedDataQuery, type PagedFindMatch } from "./paged-data-query.js";
import { PagedFormulaAdapter } from "./paged-formula-adapter.js";
import { columnIndexToName, isFormulaError, type FormulaValue } from "./formula.js";
import { insertFormulaReference, type FormulaTextRange } from "./formula-reference-authoring.js";
import { aggregateSelectionRange } from "./selection-stats.js";
import {
    createSpreadsheetI18n,
    normalizeSpreadsheetLocale,
    type SpreadsheetLocale,
    type SpreadsheetMessageKey,
    type SpreadsheetMessageValues,
} from "./i18n.js";
import {
    dropdownGridCell,
    planGridPaste,
    shouldOpenInlineFormulaEditorOnDoubleClick,
} from "./large-dataset-contracts.js";
import type { SpreadsheetColumn } from "./model.js";
import type { FilterOperator, SpreadsheetColumnSearch, SpreadsheetFilter, SpreadsheetSort } from "./view.js";

export default {
    title: "Extra Packages/Spreadsheet/Large Dataset",
};

const ROW_COUNT = 500_000;
const COLUMN_COUNT = 23;
const PAGE_SIZES = [100, 500, 1_000] as const;
const IMAGE_URL =
    "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='32' height='32'%3E%3Crect width='32' height='32' fill='%234f46e5'/%3E%3Ccircle cx='16' cy='16' r='8' fill='%23fff'/%3E%3C/svg%3E";

const DROPDOWN_OPTIONS = Object.freeze([
    "Queued",
    "In progress",
    "In review",
    "Approved",
    "Blocked",
    "Needs information",
    "Ready for release",
    "Released",
    "Archived",
    "Customer follow-up",
    "On hold",
    "Escalated",
    "Scheduled",
    "Verified",
    "Backlog",
    "Duplicate",
]);
const TAG_OPTIONS = Object.freeze([
    { tag: "Finance", color: "#2563eb" },
    { tag: "Growth", color: "#059669" },
    { tag: "Operations", color: "#d97706" },
    { tag: "Security", color: "#dc2626" },
    { tag: "Support", color: "#7c3aed" },
    { tag: "Product", color: "#db2777" },
]);
const MULTI_SELECT_OPTIONS = Object.freeze([
    { value: "Design", label: "Design", color: "#7c3aed" },
    { value: "Engineering", label: "Engineering", color: "#2563eb" },
    { value: "Marketing", label: "Marketing", color: "#db2777" },
    { value: "Sales", label: "Sales", color: "#059669" },
    { value: "Support", label: "Support", color: "#d97706" },
]);
const MULTI_SELECT_VALUES = Object.freeze([["Design"], ["Engineering", "Support"], ["Marketing"], ["Sales", "Support"]]);
const LINKS = Object.freeze([
    { title: "Documentation", href: "https://example.com/docs" },
    { title: "Record", href: "https://example.com/records" },
]);
const SPARKLINES = Object.freeze(
    Array.from({ length: 16 }, (_, pattern) =>
        Object.freeze(Array.from({ length: 12 }, (_, point) => 20 + ((pattern * 17 + point * 11) % 71)))
    )
);
const DATES = Object.freeze(
    Array.from({ length: 28 }, (_, index) => new Date(Date.UTC(2021, index % 12, (index % 28) + 1)))
);
const USERS = Object.freeze(
    Array.from({ length: 16 }, (_, index) => ({
        image: IMAGE_URL,
        initial: String.fromCharCode(65 + index),
        tint: ["#2563eb", "#7c3aed", "#db2777", "#059669"][index % 4]!,
        name: `User ${String(index + 1).padStart(2, "0")}`,
    }))
);

const spreadsheetColumns: readonly SpreadsheetColumn[] = [
    { id: "text", title: "Text", width: 150, type: "text" },
    { id: "number", title: "Number", width: 100, type: "number" },
    // IDs are stable DB names; titles are UI captions and are what users type
    // inside structured formula references.
    { id: "quantity", title: "Số lượng", width: 110, type: "number" },
    { id: "unitPrice", title: "Đơn giá", width: 120, type: "number" },
    { id: "totalFormula", title: "Thành tiền / Công thức", width: 145, type: "number" },
    { id: "boolean", title: "Boolean", width: 95, type: "boolean" },
    { id: "image", title: "Image", width: 100, type: "text", dataType: "image" },
    { id: "uri", title: "URI", width: 190, type: "text", dataType: "uri" },
    { id: "markdown", title: "Markdown", width: 160, type: "text", dataType: "markdown" },
    { id: "bubbles", title: "Bubbles", width: 140, type: "text", dataType: "bubbles" },
    { id: "drilldown", title: "Drilldown", width: 155, type: "text", dataType: "drilldown" },
    { id: "protected", title: "Protected", width: 115, type: "text" },
    { id: "loading", title: "Loading", width: 115, type: "text" },
    { id: "rowId", title: "Row ID", width: 115, type: "text", dataType: "row-id" },
    { id: "status", title: "Dropdown (searchable)", width: 175, type: "text", dataType: "dropdown", allowedValues: DROPDOWN_OPTIONS },
    { id: "trend", title: "Sparkline", width: 145, type: "number", dataType: "sparkline" },
    { id: "stars", title: "Stars", width: 110, type: "number", dataType: "stars", min: 0, max: 5 },
    { id: "tags", title: "Tags", width: 145, type: "text", dataType: "tags", allowedValues: TAG_OPTIONS.map(option => option.tag) },
    { id: "date", title: "Date", width: 135, type: "text", dataType: "date" },
    { id: "links", title: "Links", width: 165, type: "text", dataType: "links" },
    { id: "range", title: "Range", width: 145, type: "number", dataType: "range", min: 0, max: 100 },
    { id: "teams", title: "Multi-select", width: 170, type: "text", dataType: "multi-select", allowedValues: MULTI_SELECT_OPTIONS.map(option => option.value) },
    { id: "owner", title: "User profile", width: 170, type: "text", dataType: "user-profile", allowedValues: USERS.map(user => user.name) },
];
const columns: readonly GridColumn[] = spreadsheetColumns.map(column => ({ id: column.id, title: column.title, width: column.width ?? 140 }));
const MIN_COLUMN_WIDTH = 72;
const MAX_COLUMN_WIDTH = 560;

type HeaderMenuState = { readonly col: number; readonly bounds: Rectangle; readonly additive: boolean };
type HistoryEdit = { readonly col: number; readonly row: number; readonly before: PagedCellValue; readonly after: PagedCellValue };
type FormulaEditTarget = { readonly col: number; readonly sourceRow: number; readonly gridRow: number };
type CopiedGridRange = {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
    readonly sourceRows: readonly number[];
    /** Clipboard text captured with the source so stale external clipboard data is never rebased. */
    readonly values: readonly (readonly string[])[];
};
type QueryConfig = {
    readonly search: string;
    readonly columnSearches: readonly SpreadsheetColumnSearch[];
    readonly filters: readonly SpreadsheetFilter[];
    readonly sorts: readonly SpreadsheetSort[];
};
type LocalizedStatus =
    | { readonly key: SpreadsheetMessageKey; readonly values?: SpreadsheetMessageValues }
    | { readonly raw: string };

const emptySelection: GridSelection = { columns: CompactSelection.empty(), rows: CompactSelection.empty() };

function singleCellSelection(cell: Item): GridSelection {
    return {
        current: { cell, range: { x: cell[0], y: cell[1], width: 1, height: 1 }, rangeStack: [] },
        columns: CompactSelection.empty(),
        rows: CompactSelection.empty(),
    };
}

function captionCellReference(column: number, sourceRow: number, currentSourceRow?: number): string | undefined {
    const caption = spreadsheetColumns[column]?.title;
    if (caption === undefined) return undefined;
    return sourceRow === currentSourceRow ? `[@${caption}]` : `[${caption}]${sourceRow + 1}`;
}

// Formula references can be authored in every editable column. Protected and
// loading cells are intentionally excluded because they do not represent a
// value that the user can edit or persist. Custom cells remain formula-capable
// while their ordinary (non-formula) activation/rendering is left unchanged.
function isFormulaCapableColumn(column: number): boolean {
    return column >= 0 && column < spreadsheetColumns.length && column !== 11 && column !== 12;
}

/**
 * Converts the grid's current selection into the A1 reference understood by
 * the formula engine. The visible row is translated through pageRows so a
 * formula continues to point at the source row after sorting/filtering.
 */
function selectionReference(
    selection: GridSelection,
    pageRows: readonly number[],
    rowCount: number,
    columnCount: number,
    currentSourceRow: number | undefined,
    requireContiguousRows = false,
): string | undefined {
    const selectedColumns = selection.columns.toArray();
    if (selectedColumns.length > 0) {
        if (selectedColumns.some((column, index) => index > 0 && column !== selectedColumns[index - 1]! + 1)) return undefined;
        const first = selectedColumns[0]!;
        const last = selectedColumns[selectedColumns.length - 1]!;
        const from = captionCellReference(first, 0);
        const to = captionCellReference(last, rowCount - 1);
        return from === undefined || to === undefined ? undefined : `${from}:${to}`;
    }

    const selectedRows = selection.rows.toArray();
    if (selectedRows.length > 0) {
        if (selectedRows.some((row, index) => index > 0 && row !== selectedRows[index - 1]! + 1)) return undefined;
        const sourceRows = selectedRows.map(row => pageRows[row]);
        if (sourceRows.some(row => row === undefined)) return undefined;
        if (requireContiguousRows && sourceRows.some((row, index) => index > 0 && row !== sourceRows[index - 1]! + 1)) return undefined;
        const from = captionCellReference(0, sourceRows[0]!);
        const to = captionCellReference(columnCount - 1, sourceRows[sourceRows.length - 1]!);
        return from === undefined || to === undefined ? undefined : `${from}:${to}`;
    }

    const range = selection.current?.range;
    if (range !== undefined && range.width > 0 && range.height > 0) {
        const sourceRows = Array.from({ length: range.height }, (_, offset) => pageRows[range.y + offset]);
        if (sourceRows.some(row => row === undefined)) return undefined;
        if (requireContiguousRows && sourceRows.some((row, index) => index > 0 && row !== sourceRows[index - 1]! + 1)) return undefined;
        const fromRow = sourceRows[0]!;
        const toRow = sourceRows[sourceRows.length - 1]!;
        const from = captionCellReference(range.x, fromRow, range.width === 1 && range.height === 1 ? currentSourceRow : undefined);
        const to = captionCellReference(range.x + range.width - 1, toRow, range.width === 1 && range.height === 1 ? currentSourceRow : undefined);
        if (from === undefined || to === undefined) return undefined;
        return from === to ? from : `${from}:${to}`;
    }

    return undefined;
}

function filterOperators(column: SpreadsheetColumn): readonly FilterOperator[] {
    return column.type === "number" ? ["equals", "not-equals", "gt", "gte", "lt", "lte"] : column.type === "boolean" ? ["equals", "not-equals"] : ["contains", "equals", "not-equals", "empty", "not-empty"];
}

type CustomData = {
    readonly kind: string;
    readonly [key: string]: unknown;
};

function customCell(data: CustomData, copyData: string, allowOverlay = true): GridCell {
    return {
        kind: GridCellKind.Custom,
        allowOverlay,
        copyData,
        data,
    };
}

function cellText(value: PagedCellValue): string {
    return value === null ? "" : String(value);
}

function generateCell(row: number, column: number): PagedCellValue {
    switch (column) {
        case 0:
            return `Record ${String(row + 1).padStart(6, "0")}`;
        case 1:
            return (row * 37 + 11) % 10_000;
        case 2:
            return row % 12 + 1;
        case 3:
            return (row % 90) + 10;
        case 4:
            return (row % 12 + 1) * ((row % 90) + 10);
        case 5:
            return row % 2 === 0;
        case 6:
            return IMAGE_URL;
        case 7:
            return `https://example.com/records/${row + 1}`;
        case 8:
            return `**Record ${row + 1}** - Stable sample`;
        case 9:
            return `Group ${row % 8 + 1}`;
        case 10:
            return `Detail ${row % 5 + 1}`;
        case 11:
            return "Protected value";
        case 12:
            return null;
        case 13:
            return `ROW-${String(row + 1).padStart(7, "0")}`;
        case 14:
            return DROPDOWN_OPTIONS[row % DROPDOWN_OPTIONS.length] ?? null;
        case 15:
            return row % 101;
        case 16:
            return row % 6;
        case 17:
            return TAG_OPTIONS[row % TAG_OPTIONS.length]?.tag ?? null;
        case 18:
            return DATES[row % DATES.length]?.toISOString().slice(0, 10) ?? null;
        case 19:
            return "Documentation";
        case 20: {
            return (row * 13) % 101;
        }
        case 21:
            return MULTI_SELECT_VALUES[row % MULTI_SELECT_VALUES.length]?.join(",") ?? null;
        case 22:
            return USERS[row % USERS.length]?.name ?? null;
        default:
            return null;
    }
}

/** Search representation avoids generating and normalizing 23 cells per row. */
function generateSearchText(row: number): string {
    return [
        `Record ${String(row + 1).padStart(6, "0")}`,
        (row * 37 + 11) % 10_000,
        row % 12 + 1,
        (row % 90) + 10,
        (row % 12 + 1) * ((row % 90) + 10),
        row % 2 === 0,
        `https://example.com/records/${row + 1}`,
        `**Record ${row + 1}** - Stable sample`,
        `Group ${row % 8 + 1}`,
        `Detail ${row % 5 + 1}`,
        "Protected value",
        `ROW-${String(row + 1).padStart(7, "0")}`,
        DROPDOWN_OPTIONS[row % DROPDOWN_OPTIONS.length],
        row % 101,
        row % 6,
        TAG_OPTIONS[row % TAG_OPTIONS.length]?.tag,
        DATES[row % DATES.length]?.toISOString().slice(0, 10),
        "Documentation",
        (row * 13) % 101,
        MULTI_SELECT_VALUES[row % MULTI_SELECT_VALUES.length]?.join(","),
        USERS[row % USERS.length]?.name,
    ].join("\u0000");
}

function loadingCell(): GridCell {
    return { kind: GridCellKind.Loading, allowOverlay: false, skeletonWidth: 82 };
}

function formulaValueText(value: FormulaValue): string {
    if (isFormulaError(value)) return value.code;
    return value === null ? "" : String(value);
}

function formulaCell(
    formula: string,
    computed: FormulaValue,
    column: number,
): GridCell {
    const themeOverride = isFormulaError(computed) ? { textDark: "#c62828", bgCell: "#fff4f4" } : { bgCell: "#f4f8ff" };
    const text = formulaValueText(computed);

    // Keep the native scalar renderer whenever the result has that scalar
    // type. This matters for direct references such as `=C1`: a boolean or
    // URI must not be forced through the number path just because the target
    // column happens to be numeric in the sample schema.
    if (typeof computed === "boolean") {
        return { kind: GridCellKind.Boolean, allowOverlay: true, data: computed, copyData: formula, themeOverride };
    }
    if (column === 6 && typeof computed === "string") {
        return { kind: GridCellKind.Image, allowOverlay: false, data: [computed], displayData: [computed], themeOverride };
    }
    if (column === 7 && typeof computed === "string") {
        return { kind: GridCellKind.Uri, allowOverlay: true, data: computed, displayData: computed, copyData: formula, themeOverride };
    }
    if (column === 8 && typeof computed === "string") {
        return { kind: GridCellKind.Markdown, allowOverlay: true, data: computed, copyData: formula, themeOverride };
    }
    if (column === 9 && typeof computed === "string") {
        return { kind: GridCellKind.Bubble, allowOverlay: false, data: [computed], themeOverride };
    }
    if (column === 10 && typeof computed === "string") {
        return {
            kind: GridCellKind.Drilldown,
            allowOverlay: true,
            activationBehaviorOverride: "double-click",
            data: [{ text: computed, img: IMAGE_URL }],
            themeOverride,
        };
    }
    if (column === 13 && typeof computed === "string") {
        return { kind: GridCellKind.RowID, allowOverlay: false, data: computed, themeOverride };
    }
    if (column === 14 && (typeof computed === "string" || computed === null)) {
        return dropdownGridCell(computed, DROPDOWN_OPTIONS, formula, themeOverride);
    }
    if (column === 15 && typeof computed === "number") {
        const seed = Math.abs(Math.trunc(computed));
        const values = Array.from({ length: 12 }, (_, point) => 20 + ((seed * 17 + point * 11) % 71));
        return { ...customCell({ kind: "sparkline-cell", graphKind: "area", values, yAxis: [0, 100], color: "#2563eb" }, formula, false), themeOverride };
    }
    if (column === 16 && typeof computed === "number") {
        return { ...customCell({ kind: "star-cell", rating: computed }, formula, false), themeOverride };
    }
    if (column === 17 && typeof computed === "string") {
        return { ...customCell({ kind: "tags-cell", tags: [computed], possibleTags: TAG_OPTIONS }, formula), themeOverride };
    }
    if (column === 18 && typeof computed === "string") {
        const date = new Date(`${computed}T00:00:00Z`);
        if (!Number.isNaN(date.getTime())) {
            return { ...customCell({ kind: "date-picker-cell", date, displayDate: computed, format: "date" }, formula), themeOverride };
        }
    }
    if (column === 19 && typeof computed === "string") {
        return { ...customCell({ kind: "links-cell", links: [{ title: computed, href: LINKS[0]!.href }] }, formula, false), themeOverride };
    }
    if (column === 20 && typeof computed === "number") {
        return { ...customCell({ kind: "range-cell", value: computed, min: 0, max: 100, step: 1 }, formula), themeOverride };
    }
    if (column === 21 && typeof computed === "string") {
        return {
            ...customCell({ kind: "multi-select-cell", values: computed === "" ? [] : computed.split(","), options: MULTI_SELECT_OPTIONS }, formula),
            activationBehaviorOverride: "double-click",
            themeOverride,
        };
    }
    if (column === 22 && typeof computed === "string") {
        const user = USERS.find(item => item.name === computed) ?? { ...USERS[0]!, name: computed };
        return { ...customCell({ kind: "user-profile-cell", ...user }, formula), themeOverride };
    }

    return {
        kind: GridCellKind.Text,
        allowOverlay: true,
        data: formula,
        displayData: text,
        // Formula results retain their scalar type in the adapter. Keep
        // numeric results aligned exactly like ordinary numeric cells,
        // while leaving text/date references left-aligned.
        contentAlign: typeof computed === "number" ? "right" : "left",
        copyData: formula,
        themeOverride,
    };
}

function contentFor(source: PagedDataSource, formulas: PagedFormulaAdapter, sourceRow: number, column: number): GridCell {
    const formula = formulas.getFormula(column, sourceRow);
    if (formula !== undefined) {
        const computed = formulas.getValue(column, sourceRow);
        if (computed === LOADING_CELL) return loadingCell();
        return formulaCell(formula, computed, column);
    }
    const raw = source.readCell(sourceRow, column);

    const text = cellText(raw);
    switch (column) {
        case 0:
            return { kind: GridCellKind.Text, allowOverlay: true, data: text, displayData: text };
        case 1: {
            // Keep numeric values as numbers in the source/formula model, but
            // use the text editor in the grid so a user can start an in-cell
            // formula with `=` (the built-in number editor only accepts digits).
            const value = typeof raw === "number" ? raw : 0;
            return { kind: GridCellKind.Text, allowOverlay: true, data: String(value), displayData: String(value), contentAlign: "right" };
        }
        case 2:
        case 3:
        case 4: {
                const value = typeof raw === "number" ? raw : 0;
                return { kind: GridCellKind.Text, allowOverlay: true, data: String(value), displayData: String(value), contentAlign: "right" };
            }
        case 5:
            return { kind: GridCellKind.Boolean, allowOverlay: false, data: raw === true };
        case 6:
            return { kind: GridCellKind.Image, allowOverlay: false, data: [IMAGE_URL], displayData: [IMAGE_URL] };
        case 7:
            return { kind: GridCellKind.Uri, allowOverlay: true, data: text, displayData: text };
        case 8:
            return { kind: GridCellKind.Markdown, allowOverlay: true, data: text };
        case 9:
            return { kind: GridCellKind.Bubble, allowOverlay: false, data: [text] };
        case 10:
            {
                const values = text.split(",").map(value => value.trim()).filter(value => value !== "");
                return {
                    kind: GridCellKind.Drilldown,
                    // Drilldown previews remain selectable/copyable on a
                    // single click; a double click opens the editable list.
                    allowOverlay: true,
                    activationBehaviorOverride: "double-click",
                    data: (values.length === 0 ? [""] : values).map(value => ({ text: value, img: IMAGE_URL })),
                };
            }
        case 11:
            return { kind: GridCellKind.Protected, allowOverlay: false };
        case 12:
            return loadingCell();
        case 13:
            return { kind: GridCellKind.RowID, allowOverlay: false, data: text };
        case 14:
            return dropdownGridCell(typeof raw === "string" ? raw : null, DROPDOWN_OPTIONS, text);
        case 15:
            return customCell({ kind: "sparkline-cell", graphKind: "area", values: SPARKLINES[sourceRow % SPARKLINES.length]!, yAxis: [0, 100], color: "#2563eb" }, text, false);
        case 16:
            return customCell({ kind: "star-cell", rating: typeof raw === "number" ? raw : 0 }, text, false);
        case 17:
            return customCell({ kind: "tags-cell", tags: [typeof raw === "string" ? raw : ""], possibleTags: TAG_OPTIONS }, text, false);
        case 18: {
            const date = DATES[sourceRow % DATES.length] ?? DATES[0]!;
            return customCell({ kind: "date-picker-cell", date, displayDate: date.toISOString().slice(0, 10), format: "date" }, text);
        }
        case 19:
            return customCell({ kind: "links-cell", links: LINKS }, text, false);
        case 20:
            return customCell({ kind: "range-cell", value: typeof raw === "number" ? raw : 0, min: 0, max: 100, step: 1 }, text, false);
        case 21:
            return {
                ...customCell(
                    {
                        kind: "multi-select-cell",
                        values: text
                            .split(",")
                            .map(value => value.trim())
                            .filter(value => value !== ""),
                        options: MULTI_SELECT_OPTIONS,
                    },
                    text
                ),
                activationBehaviorOverride: "double-click",
            };
        case 22:
            return customCell({ kind: "user-profile-cell", ...USERS[sourceRow % USERS.length]! }, text);
        default:
            return loadingCell();
    }
}

function valueFromEditedCell(cell: EditableGridCell): PagedCellValue | undefined {
    switch (cell.kind) {
        case GridCellKind.Text:
        case GridCellKind.Uri:
        case GridCellKind.Markdown:
            return cell.data;
        case GridCellKind.Number:
            return Number.isFinite(cell.data) ? cell.data : undefined;
        case GridCellKind.Boolean:
            return typeof cell.data === "boolean" ? cell.data : undefined;
        case GridCellKind.Drilldown:
            return cell.data.map(value => value.text).join(",");
        case GridCellKind.Custom: {
            const data = cell.data as {
                readonly kind?: unknown;
                readonly value?: unknown;
                readonly displayDate?: unknown;
                readonly values?: unknown;
            };
            if (data.kind === "dropdown-cell" && (typeof data.value === "string" || data.value === null)) return data.value;
            if (data.kind === "date-picker-cell" && typeof data.displayDate === "string") return data.displayDate;
            if (data.kind === "multi-select-cell" && Array.isArray(data.values)) {
                return data.values.filter((value): value is string => typeof value === "string").join(",");
            }
            return undefined;
        }
        default:
            return undefined;
    }
}

function BasicLargeDatasetStory(): React.ReactElement {
    const [pageSize, setPageSize] = React.useState<number>(500);
    const source = React.useMemo(
        () =>
            new PagedDataSource({
                rowCount: ROW_COUNT,
                columnCount: COLUMN_COUNT,
                pageSize,
                maxCachedPages: 8,
                seed: 0x5eed,
                generateCell,
            }),
        [pageSize]
    );
    const formulas = React.useMemo(() => new PagedFormulaAdapter(source, spreadsheetColumns), [source]);
    const [pageIndex, setPageIndex] = React.useState(0);
    const [pageInput, setPageInput] = React.useState("1");
    const [loadedPage, setLoadedPage] = React.useState<number | undefined>(undefined);
    const [loadMs, setLoadMs] = React.useState<number | undefined>(undefined);
    const [revision, setRevision] = React.useState(0);

    React.useEffect(() => {
        let cancelled = false;
        const started = performance.now();
        setLoadedPage(undefined);
        setLoadMs(undefined);
        void source.getPage(pageIndex).then(() => {
            if (!cancelled) {
                setLoadedPage(pageIndex);
                setLoadMs(performance.now() - started);
            }
        });
        return () => {
            cancelled = true;
        };
    }, [pageIndex, source]);

    const pageCount = source.pageCount;
    const pageRowCount = Math.min(pageSize, ROW_COUNT - pageIndex * pageSize);
    const goToPage = React.useCallback(
        (nextPage: number) => {
            const bounded = Math.max(0, Math.min(pageCount - 1, Math.trunc(nextPage)));
            setPageIndex(bounded);
            setPageInput(String(bounded + 1));
        },
        [pageCount]
    );
    const getCellContent = React.useCallback(
        ([column, row]: Item): GridCell => {
            if (loadedPage !== pageIndex || row < 0 || row >= pageRowCount) return loadingCell();
            return contentFor(source, formulas, pageIndex * pageSize + row, column);
        },
        [formulas, loadedPage, pageIndex, pageRowCount, pageSize, revision, source]
    );
    const onCellEdited = React.useCallback(
        ([column, row]: Item, cell: EditableGridCell) => {
            if (loadedPage !== pageIndex || row < 0 || row >= pageRowCount) return;
            const value = valueFromEditedCell(cell);
            if (value === undefined || column === 12) return;
            source.updateCell(pageIndex * pageSize + row, column, value);
            setRevision(value => value + 1);
        },
        [loadedPage, pageIndex, pageRowCount, pageSize, source]
    );
    const stats = source.getStats();

    return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8, height: 620, minWidth: 1_300 }}>
            <div style={{ display: "flex", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <strong>500,000 rows × 23 columns</strong>
                <label>
                    Page size{" "}
                    <select
                        aria-label="Large dataset page size"
                        value={pageSize}
                        onChange={event => {
                            setPageSize(Number(event.target.value));
                            setPageIndex(0);
                            setPageInput("1");
                        }}
                    >
                        {PAGE_SIZES.map(size => (
                            <option key={size} value={size}>
                                {size}
                            </option>
                        ))}
                    </select>
                </label>
                <button type="button" aria-label="First page" disabled={pageIndex === 0} onClick={() => goToPage(0)}>
                    First
                </button>
                <button type="button" aria-label="Previous page" disabled={pageIndex === 0} onClick={() => goToPage(pageIndex - 1)}>
                    Prev
                </button>
                <button type="button" aria-label="Next page" disabled={pageIndex >= pageCount - 1} onClick={() => goToPage(pageIndex + 1)}>
                    Next
                </button>
                <button type="button" aria-label="Last page" disabled={pageIndex >= pageCount - 1} onClick={() => goToPage(pageCount - 1)}>
                    Last
                </button>
                <label>
                    Page{" "}
                    <input
                        aria-label="Page number"
                        type="number"
                        min={1}
                        max={pageCount}
                        value={pageInput}
                        onChange={event => setPageInput(event.target.value)}
                        onKeyDown={event => {
                            if (event.key === "Enter") goToPage(Number(pageInput) - 1);
                        }}
                        style={{ width: 72 }}
                    />
                    {` / ${pageCount.toLocaleString()}`}
                </label>
            </div>
            <div aria-live="polite" style={{ color: "#4b5563", fontSize: 12 }}>
                {loadedPage === pageIndex ? `Loaded page in ${(loadMs ?? 0).toFixed(1)} ms` : "Loading page…"} · cache {stats.cachedPages}/{stats.maxCachedPages} · hits {stats.cacheHits} · misses {stats.cacheMisses} · generated {stats.generatedCells.toLocaleString()} cells
            </div>
            <div style={{ flex: 1, minHeight: 0 }}>
                <DataEditor
                    columns={columns}
                    rows={pageRowCount}
                    getCellContent={getCellContent}
                    onCellEdited={onCellEdited}
                    customRenderers={allCells}
                    rowMarkers="number"
                    freezeColumns={1}
                    smoothScrollX={true}
                    smoothScrollY={true}
                />
            </div>
        </div>
    );
}

function parseFormulaInput(value: string, column: SpreadsheetColumn): PagedCellValue {
    if (value.startsWith("=")) return value;
    if (column.type === "number") {
        const parsed = Number(value);
        return value.trim() === "" ? null : Number.isFinite(parsed) ? parsed : value;
    }
    if (column.type === "boolean") return value.toLocaleLowerCase() === "true";
    return value;
}

function formulaQueryValue(value: FormulaValue | typeof LOADING_CELL): PagedCellValue {
    if (value === LOADING_CELL) return null;
    if (isFormulaError(value)) return value.code;
    return value;
}

function selectionMetricText(value: FormulaValue, locale?: string): string {
    if (isFormulaError(value)) return value.code;
    if (value === null) return "—";
    return typeof value === "number" ? value.toLocaleString(locale, { maximumFractionDigits: 2 }) : String(value);
}

/** Integrated 500k-row spreadsheet: virtual paging, formulas, query, find and header menus. */
export function LargeDatasetStory(): React.ReactElement {
    const [locale, setLocale] = React.useState<SpreadsheetLocale>(() =>
        normalizeSpreadsheetLocale(typeof navigator === "undefined" ? "en" : navigator.language)
    );
    const i18n = React.useMemo(() => createSpreadsheetI18n(locale), [locale]);
    const { t } = i18n;
    const numberLocale = locale === "vi" ? "vi-VN" : "en-US";
    const builtInSearchLabels = React.useMemo(() => ({
        result: t("searchResult"),
        results: t("searchResults"),
        over1000: t("searchOver1000"),
        of: t("searchOf"),
        previous: t("previousResult"),
        next: t("nextResult"),
        close: t("closeSearch"),
        typeToSearch: t("typeToSearch"),
    }), [t]);
    const sourceRef = React.useRef<PagedDataSource | null>(null);
    if (sourceRef.current === null) sourceRef.current = new PagedDataSource({ rowCount: ROW_COUNT, columnCount: COLUMN_COUNT, pageSize: 512, maxCachedPages: 8, seed: 0x5eed, generateCell });
    const source = sourceRef.current;
    const formulaRef = React.useRef<PagedFormulaAdapter | null>(null);
    if (formulaRef.current === null) formulaRef.current = new PagedFormulaAdapter(source, spreadsheetColumns);
    const formulas = formulaRef.current;
    const queryRef = React.useRef<PagedDataQuery | null>(null);
    if (queryRef.current === null) queryRef.current = new PagedDataQuery({
        rowCount: ROW_COUNT,
        columnCount: COLUMN_COUNT,
        pageSize: 512,
        // Query the deterministic base source directly. Sparse formula values
        // stay in the formula adapter and do not turn an 11.5M-cell scan into
        // millions of dependency-map lookups.
        getCell: (row, col) => formulas.hasFormula(col, row) ? formulaQueryValue(formulas.getValue(col, row)) : source.readCell(row, col),
        readCell: (row, col) => formulas.hasFormula(col, row) ? formulaQueryValue(formulas.getValue(col, row)) : source.readCell(row, col),
        readSearchText: row => {
            const formulaColumns = formulas.getFormulaColumns(row);
            const overlayColumns = source.getOverlayColumns(row);
            if (formulaColumns.length === 0 && overlayColumns.length === 0) return generateSearchText(row);
            const formulaSet = new Set(formulaColumns);
            return Array.from({ length: COLUMN_COUNT }, (_, col) => formulaSet.has(col) ? formulas.getDisplayValue(col, row) : source.readCell(row, col) ?? "").join("\u0000");
        },
    }, spreadsheetColumns);
    const query = queryRef.current;
    const gridRef = React.useRef<React.ElementRef<typeof DataEditor>>(null);
    const headerMenuRef = React.useRef<HTMLDivElement>(null);
    const undoRef = React.useRef<HistoryEdit[][]>([]);
    const redoRef = React.useRef<HistoryEdit[][]>([]);
    const queryRequestRef = React.useRef(0);
    const [revision, setRevision] = React.useState(0);
    const [pageSize, setPageSize] = React.useState(500);
    const [pageIndex, setPageIndex] = React.useState(0);
    const [pageInput, setPageInput] = React.useState("1");
    const [selection, setSelection] = React.useState<GridSelection>(emptySelection);
    const [formulaDraft, setFormulaDraft] = React.useState("");
    const [searchDraft, setSearchDraft] = React.useState("");
    const [config, setConfig] = React.useState<QueryConfig>({ search: "", columnSearches: [], filters: [], sorts: [] });
    const [rowIndex, setRowIndex] = React.useState<readonly number[] | undefined>();
    const [totalCount, setTotalCount] = React.useState(ROW_COUNT);
    const [queryBusy, setQueryBusy] = React.useState(false);
    const [queryMs, setQueryMs] = React.useState(0);
    const [status, setStatus] = React.useState<LocalizedStatus>({ key: "ready" });
    const [formulaErrorMessage, setFormulaErrorMessage] = React.useState<string | undefined>();
    const [headerMenu, setHeaderMenu] = React.useState<HeaderMenuState | undefined>();
    const [filterOperator, setFilterOperator] = React.useState<FilterOperator>("contains");
    const [filterValue, setFilterValue] = React.useState("");
    const [columnSearchDraft, setColumnSearchDraft] = React.useState("");
    const [findOpen, setFindOpen] = React.useState(false);
    const [findQuery, setFindQuery] = React.useState("");
    const [replaceText, setReplaceText] = React.useState("");
    const [findCase, setFindCase] = React.useState(false);
    const [findWhole, setFindWhole] = React.useState(false);
    const [findRegex, setFindRegex] = React.useState(false);
    const [findBusy, setFindBusy] = React.useState(false);
    const [findMatches, setFindMatches] = React.useState<readonly PagedFindMatch[]>([]);
    const [findTotal, setFindTotal] = React.useState(0);
    const [findIndex, setFindIndex] = React.useState(-1);
    const statusText = "raw" in status ? status.raw : t(status.key, status.values);
    // Keep user-resized widths separate from the model columns. The model
    // columns are keyed by stable DB ids, while their titles can change as
    // captions and query indicators are updated during a render.
    const [columnWidths, setColumnWidths] = React.useState<Record<string, number>>(() =>
        Object.fromEntries(spreadsheetColumns.map(column => [column.id, column.width ?? 140]))
    );
    // Formula authoring is intentionally kept outside the 500k-row model. It
    // only tracks the input caret and the last live reference, so clicking a
    // cell/range while the formula bar is active can insert an A1 reference
    // without causing a full data scan.
    const formulaDraftRef = React.useRef(formulaDraft);
    const formulaCursorRef = React.useRef(0);
    const formulaEditingRef = React.useRef(false);
    const formulaInputRef = React.useRef<HTMLInputElement>(null);
    const inlineFormulaInputRef = React.useRef<HTMLInputElement>(null);
    const formulaEditTargetRef = React.useRef<FormulaEditTarget | undefined>(undefined);
    const inlineFormulaOpenRef = React.useRef(false);
    const formulaPickingRef = React.useRef(false);
    const formulaCellClickRef = React.useRef(false);
    // DataEditor advances the selection after a native overlay commits. When
    // that commit is a rejected formula, the following selection event must
    // not be mistaken for a reference pick and appended to the draft.
    const rejectedFormulaEditRef = React.useRef<Item | undefined>(undefined);
    const formulaReferenceRangeRef = React.useRef<FormulaTextRange | undefined>(undefined);
    const copiedGridRangeRef = React.useRef<CopiedGridRange | undefined>(undefined);
    const [inlineFormulaOpen, setInlineFormulaOpen] = React.useState(false);

    const setFormulaDraftValue = React.useCallback((value: string, cursor = value.length) => {
        formulaDraftRef.current = value;
        formulaCursorRef.current = Math.max(0, Math.min(value.length, cursor));
        // A new draft supersedes any previous validation failure.
        setFormulaErrorMessage(undefined);
        setFormulaDraft(value);
    }, []);

    const endFormulaSession = React.useCallback((clearDraft = false) => {
        formulaEditingRef.current = false;
        formulaEditTargetRef.current = undefined;
        formulaReferenceRangeRef.current = undefined;
        inlineFormulaOpenRef.current = false;
        setInlineFormulaOpen(false);
        setFormulaErrorMessage(undefined);
        if (clearDraft) setFormulaDraftValue("");
    }, [setFormulaDraftValue]);

    const executeQuery = React.useCallback(async (next: QueryConfig) => {
        const request = ++queryRequestRef.current;
        setConfig(next);
        setPageIndex(0);
        setPageInput("1");
        const active = next.search.trim() !== "" || next.columnSearches.some(item => item.value.trim() !== "") || next.filters.length > 0 || next.sorts.length > 0;
        if (!active) {
            setQueryBusy(false);
            setRowIndex(undefined);
            setTotalCount(ROW_COUNT);
            setQueryMs(0);
            setStatus({ key: "queryCleared" });
            return;
        }
        setQueryBusy(true);
        const started = performance.now();
        try {
            const result = await query.query({ ...next, offset: 0, limit: ROW_COUNT });
            if (request !== queryRequestRef.current) return;
            setRowIndex(result.rows);
            setTotalCount(result.totalCount);
            setQueryMs(performance.now() - started);
            setStatus({ key: "matchingRows", values: { count: result.totalCount } });
        } catch (error) {
            if (request !== queryRequestRef.current) return;
            setStatus({ raw: error instanceof Error ? error.message : String(error) });
        } finally {
            if (request === queryRequestRef.current) setQueryBusy(false);
        }
    }, [query]);

    const pageCount = Math.max(1, Math.ceil(totalCount / pageSize));
    const pageStart = pageIndex * pageSize;
    const pageRows = React.useMemo(() => Array.from({ length: Math.max(0, Math.min(pageSize, totalCount - pageStart)) }, (_, offset) => rowIndex?.[pageStart + offset] ?? pageStart + offset), [pageSize, pageStart, rowIndex, totalCount]);
    const insertReferenceForSelection = React.useCallback((nextSelection: GridSelection) => {
        const draft = formulaDraftRef.current;
        if (!formulaEditingRef.current || !draft.startsWith("=")) return false;
        const filteredOrSorted = rowIndex !== undefined || config.filters.length > 0 || config.sorts.length > 0;
        const reference = selectionReference(nextSelection, pageRows, ROW_COUNT, COLUMN_COUNT, formulaEditTargetRef.current?.sourceRow, filteredOrSorted);
        if (reference === undefined) {
            setStatus({ key: "nonContiguousReference" });
            formulaPickingRef.current = false;
            return false;
        }
        const cursor = formulaCursorRef.current;
        const inserted = insertFormulaReference(
            draft,
            reference,
            { start: cursor, end: cursor },
            formulaReferenceRangeRef.current,
        );
        formulaReferenceRangeRef.current = inserted.referenceRange;
        setFormulaDraftValue(inserted.formula, inserted.cursor);
        // Keep typing in the formula bar after a grid click. The timeout lets
        // the grid finish its pointer/selection event before focus is restored.
        setTimeout(() => {
            const input = inlineFormulaOpenRef.current ? inlineFormulaInputRef.current : formulaInputRef.current;
            input?.focus();
            input?.setSelectionRange(inserted.cursor, inserted.cursor);
            formulaPickingRef.current = false;
        }, 0);
        return true;
    }, [config.filters.length, config.sorts.length, pageRows, rowIndex, setFormulaDraftValue]);
    const beginFormulaSession = React.useCallback((target: Item, initialDraft = "=") => {
        const sourceRow = pageRows[target[1]];
        if (sourceRow === undefined || spreadsheetColumns[target[0]] === undefined || !isFormulaCapableColumn(target[0])) return;
        formulaEditTargetRef.current = { col: target[0], sourceRow, gridRow: target[1] };
        formulaEditingRef.current = true;
        formulaReferenceRangeRef.current = undefined;
        setFormulaDraftValue(initialDraft);
    }, [pageRows, setFormulaDraftValue]);
    const openInlineFormulaSession = React.useCallback((target: Item, initialDraft = "=", selectAll = false) => {
        beginFormulaSession(target, initialDraft);
        inlineFormulaOpenRef.current = true;
        setInlineFormulaOpen(true);
        setTimeout(() => {
            const input = inlineFormulaInputRef.current;
            input?.focus();
            if (selectAll) input?.select();
            else input?.setSelectionRange(initialDraft.length, initialDraft.length);
        }, 0);
    }, [beginFormulaSession]);
    const selectedCell = selection.current?.cell;
    const selectedSourceRow = selectedCell === undefined ? undefined : pageRows[selectedCell[1]];
    const selectionSummary = React.useMemo(() => {
        const result = aggregateSelectionRange(
            selection.current?.range,
            pageRows,
            spreadsheetColumns.length,
            (column, sourceRow) => {
                if (formulas.hasFormula(column, sourceRow)) {
                    const value = formulas.getValue(column, sourceRow);
                    return value === LOADING_CELL ? null : value;
                }
                return source.readCell(sourceRow, column);
            },
        );
        if (result.kind === "empty") return t("selectCellsToCalculate");
        if (result.kind === "too-large") return t("selectionStatsLimited", { max: result.maxCells.toLocaleString(numberLocale), count: result.cellCount.toLocaleString(numberLocale) });
        const aggregate = result.aggregate;
        return `${t("selected")} ${result.cellCount.toLocaleString(numberLocale)} · ${t("sum")}: ${selectionMetricText(aggregate.sum, numberLocale)} · ${t("average")}: ${selectionMetricText(aggregate.average, numberLocale)} · ${t("count")}: ${aggregate.countAll.toLocaleString(numberLocale)} · ${t("countNumbers")}: ${aggregate.countNumbers.toLocaleString(numberLocale)} · ${t("min")}: ${selectionMetricText(aggregate.min, numberLocale)} · ${t("max")}: ${selectionMetricText(aggregate.max, numberLocale)}`;
    }, [formulas, numberLocale, pageRows, revision, selection, source, t]);

    React.useEffect(() => {
        if (selectedCell === undefined || selectedSourceRow === undefined) {
            if (!formulaEditingRef.current) setFormulaDraftValue("");
            return;
        }
        if (formulaEditingRef.current) return;
        const input = formulas.getInput(selectedCell[0], selectedSourceRow);
        setFormulaDraftValue(input === LOADING_CELL || input === null ? "" : String(input));
    }, [formulas, revision, selectedCell, selectedSourceRow, setFormulaDraftValue]);

    React.useEffect(() => {
        if (headerMenu === undefined) return;
        const close = (event: PointerEvent) => {
            if (!(event.target instanceof Node) || !headerMenuRef.current?.contains(event.target)) setHeaderMenu(undefined);
        };
        document.addEventListener("pointerdown", close);
        return () => document.removeEventListener("pointerdown", close);
    }, [headerMenu]);

    const refreshGrid = () => {
        setRevision(value => value + 1);
        gridRef.current?.updateCells(pageRows.flatMap((_, row) => spreadsheetColumns.map((__, col) => ({ cell: [col, row] as Item }))));
    };
    const applyInput = (col: number, row: number, after: PagedCellValue, record = true, refresh = true) => {
        const rawBefore = formulas.getInput(col, row);
        const before = rawBefore === LOADING_CELL ? null : rawBefore;
        formulas.setCell(col, row, after);
        const edit = { col, row, before, after };
        if (record) {
            undoRef.current.push([edit]);
            redoRef.current = [];
        }
        if (refresh) refreshGrid();
        return edit;
    };
    const commitFormula = () => {
        const target = formulaEditTargetRef.current;
        if (target === undefined) return;
        const draft = formulaDraftRef.current;
        if (draft.startsWith("=")) {
            const validation = formulas.validateFormula(target.col, target.sourceRow, draft);
            if (!validation.valid) {
                const error = validation.error;
                const message = error === undefined ? t("formulaInvalid") : `${error.code}: ${error.message ?? t("formulaInvalid")}`;
                setFormulaErrorMessage(message);
                setStatus({ raw: message });
                refreshGrid();
                return;
            }
        }
        applyInput(target.col, target.sourceRow, parseFormulaInput(draft, spreadsheetColumns[target.col]!));
        endFormulaSession();
        setStatus({ key: "updatedCell", values: { cell: `${columnIndexToName(target.col)}${target.sourceRow + 1}` } });
    };
    /**
     * Own the clipboard matrix so a paste is one bounded transaction. The
     * DataEditor callback gives us visible grid coordinates; pageRows maps
     * each of those coordinates back to the source row even while a query is
     * filtered or sorted. Returning false prevents DataEditor from emitting
     * one edit at a time and lets us record one undo entry for the whole paste.
     */
    const onGridPaste = React.useCallback((target: Item, values: readonly (readonly string[])[]): boolean => {
        const plan = planGridPaste<PagedCellValue>({
            target,
            selection: selection.current?.range,
            values,
            pageRows,
            columnCount: spreadsheetColumns.length,
            rememberedCopy: copiedGridRangeRef.current,
            canWrite: col => col !== 11 && col !== 12,
            getBefore: (col, row) => {
                const before = formulas.getInput(col, row);
                return before === LOADING_CELL ? null : before;
            },
            parse: (value, col) => parseFormulaInput(value, spreadsheetColumns[col]!),
            validateFormula: (col, row, formula) => {
                const validation = formulas.validateFormula(col, row, formula);
                return validation.valid ? undefined : validation.error;
            },
        });

        if (!plan.valid) {
            // Keep the entire clipboard operation atomic: no cell is written
            // when any formula in the matrix fails validation.
            formulaEditTargetRef.current = { col: plan.col, sourceRow: plan.row, gridRow: pageRows.indexOf(plan.row) };
            formulaEditingRef.current = true;
            formulaReferenceRangeRef.current = undefined;
            setFormulaDraftValue(plan.value);
            setFormulaErrorMessage(plan.message);
            setStatus({ raw: plan.message });
            refreshGrid();
            return false;
        }

        if (plan.edits.length === 0) return false;
        plan.edits.forEach(edit => formulas.setCell(edit.col, edit.row, edit.after));
        undoRef.current.push([...plan.edits]);
        redoRef.current = [];
        refreshGrid();
        setFormulaErrorMessage(undefined);
        setStatus({ key: "pastedCells", values: { count: plan.edits.length } });
        return false;
    }, [formulas, pageRows, refreshGrid, selection, setFormulaDraftValue]);
    const undo = () => {
        const edits = undoRef.current.pop();
        if (edits === undefined) return;
        for (const edit of [...edits].reverse()) formulas.setCell(edit.col, edit.row, edit.before);
        redoRef.current.push(edits);
        refreshGrid();
    };
    const redo = () => {
        const edits = redoRef.current.pop();
        if (edits === undefined) return;
        for (const edit of edits) formulas.setCell(edit.col, edit.row, edit.after);
        undoRef.current.push(edits);
        refreshGrid();
    };
    const autoSum = () => {
        if (selectedCell === undefined || selectedSourceRow === undefined || selectedSourceRow === 0) return;
        if (selectedSourceRow > 100_000) {
            setStatus({ key: "autoSumLimit", values: { count: 100_000 } });
            return;
        }
        const name = columnIndexToName(selectedCell[0]);
        const formula = `=SUM(${name}1:${name}${selectedSourceRow})`;
        formulaEditTargetRef.current = { col: selectedCell[0], sourceRow: selectedSourceRow, gridRow: selectedCell[1] };
        formulaEditingRef.current = true;
        setFormulaDraftValue(formula);
    };
    const goToPage = (page: number) => {
        const safe = Number.isFinite(page) ? Math.max(0, Math.min(pageCount - 1, Math.trunc(page))) : 0;
        setPageIndex(safe);
        setPageInput(String(safe + 1));
    };
    const openHeader = (col: number, bounds: Rectangle, additive = false) => {
        const column = spreadsheetColumns[col];
        if (column === undefined) return;
        const filter = config.filters.find(item => item.column === column.id);
        setFilterOperator(filter?.operator ?? filterOperators(column)[0]!);
        setFilterValue(filter?.value === undefined ? "" : String(filter.value));
        setColumnSearchDraft(config.columnSearches.find(item => item.column === column.id)?.value ?? "");
        setHeaderMenu({ col, bounds, additive });
    };
    const applySort = (direction: SpreadsheetSort["direction"]) => {
        if (headerMenu === undefined) return;
        const id = spreadsheetColumns[headerMenu.col]!.id;
        const rest = config.sorts.filter(item => item.column !== id);
        const sorts = headerMenu.additive ? [...rest, { column: id, direction }] : [{ column: id, direction }];
        setHeaderMenu(undefined);
        void executeQuery({ ...config, sorts });
    };
    const applyFilter = () => {
        if (headerMenu === undefined) return;
        const column = spreadsheetColumns[headerMenu.col]!;
        const rest = config.filters.filter(item => item.column !== column.id);
        let value: string | number | boolean | undefined = filterValue;
        if (column.type === "number") value = Number(filterValue);
        if (column.type === "boolean") value = filterValue === "true";
        const filters = filterOperator === "empty" || filterOperator === "not-empty" ? [...rest, { column: column.id, operator: filterOperator }] : filterValue.trim() === "" ? rest : [...rest, { column: column.id, operator: filterOperator, value }];
        setHeaderMenu(undefined);
        void executeQuery({ ...config, filters });
    };
    const applyColumnSearch = () => {
        if (headerMenu === undefined) return;
        const id = spreadsheetColumns[headerMenu.col]!.id;
        const columnSearches = [...config.columnSearches.filter(item => item.column !== id), ...(columnSearchDraft.trim() === "" ? [] : [{ column: id, value: columnSearchDraft }])];
        setHeaderMenu(undefined);
        void executeQuery({ ...config, columnSearches });
    };
    const runFind = async () => {
        if (findQuery === "") return;
        setFindBusy(true);
        try {
            const result = await query.find({ query: findQuery, caseSensitive: findCase, wholeCell: findWhole, regexp: findRegex, maxResults: 1_000 });
            setFindMatches(result.matches);
            setFindTotal(result.totalMatches);
            setFindIndex(result.matches.length === 0 ? -1 : 0);
            setStatus(result.truncated
                ? { key: "findMatchesRetained", values: { count: result.totalMatches, retained: 1_000 } }
                : { key: "findMatches", values: { count: result.totalMatches } });
        } catch (error) {
            setStatus(error instanceof Error ? { key: "findError", values: { message: error.message } } : { raw: String(error) });
            setFindMatches([]);
            setFindTotal(0);
            setFindIndex(-1);
        } finally {
            setFindBusy(false);
        }
    };
    const revealMatch = (delta: number) => {
        if (findMatches.length === 0) return;
        const next = findIndex < 0 ? 0 : (findIndex + delta + findMatches.length) % findMatches.length;
        const match = findMatches[next]!;
        const displayRow = rowIndex === undefined ? match.row : rowIndex.indexOf(match.row);
        if (displayRow < 0) {
            setStatus({ key: "matchHidden" });
            return;
        }
        const targetPage = Math.floor(displayRow / pageSize);
        const localRow = displayRow % pageSize;
        goToPage(targetPage);
        setFindIndex(next);
        setSelection(singleCellSelection([match.column, localRow]));
        setTimeout(() => gridRef.current?.scrollTo(match.column, localRow, "both"), 0);
    };
    const replaceMatches = (all: boolean) => {
        const selected = all ? findMatches : findMatches[findIndex] === undefined ? [] : [findMatches[findIndex]!];
        if (all && findTotal > findMatches.length) {
            setStatus({ key: "replaceAllLimit", values: { count: 1_000 } });
            return;
        }
        let replacementExpression: RegExp | undefined;
        try {
            const pattern = findRegex ? findQuery : findQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            replacementExpression = new RegExp(findWhole ? `^(?:${pattern})$` : pattern, `${findCase ? "" : "i"}${all ? "g" : ""}`);
        } catch (error) {
            setStatus(error instanceof Error ? { key: "replaceError", values: { message: error.message } } : { raw: String(error) });
            return;
        }
        const edits: HistoryEdit[] = [];
        for (const match of selected) {
            const before = formulas.getInput(match.column, match.row);
            const rawBefore = before === LOADING_CELL ? null : before;
            const next = String(rawBefore ?? "").replace(replacementExpression, replaceText);
            edits.push(applyInput(match.column, match.row, parseFormulaInput(next, spreadsheetColumns[match.column]!), false, false));
        }
        if (edits.length > 0) {
            undoRef.current.push(edits);
            redoRef.current = [];
        }
        refreshGrid();
        setStatus({ key: "replacedCells", values: { count: selected.length } });
        void runFind();
    };

    const onColumnResize = React.useCallback((column: GridColumn, newSize: number) => {
        const columnId = column.id;
        if (columnId === undefined || !Number.isFinite(newSize)) return;
        const width = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(newSize)));
        setColumnWidths(current => current[columnId] === width ? current : { ...current, [columnId]: width });
    }, []);

    const onGridSelectionChange = React.useCallback((nextSelection: GridSelection) => {
        const rejectedTarget = rejectedFormulaEditRef.current;
        if (rejectedTarget !== undefined) {
            rejectedFormulaEditRef.current = undefined;
            setSelection(singleCellSelection(rejectedTarget));
            return;
        }
        // A formula-bar edit keeps its destination cell selected while the
        // user clicks/drag-selects references in the grid. The selected cell,
        // range, or entire column is inserted at the caret as A1 syntax.
        if (formulaEditingRef.current && formulaDraftRef.current.startsWith("=")) {
            // A cell click is handled by onCellClicked before DataEditor
            // dispatches this selection event. Keep the destination selected;
            // the clicked cell has already been inserted there.
            if (formulaCellClickRef.current) {
                formulaCellClickRef.current = false;
                return;
            }
            // Header clicks also emit a column selection before
            // onHeaderClicked. Defer that case so the header handler can
            // insert the stable caption form, e.g. [@Số lượng], instead of a
            // 500,000-row A1 range.
            const selectedColumns = nextSelection.columns.toArray();
            if (selectedColumns.length > 0) {
                if (selectedColumns.length === 1) {
                    const caption = spreadsheetColumns[selectedColumns[0]!]?.title;
                    if (caption !== undefined) {
                        formulaPickingRef.current = true;
                        const inserted = insertFormulaReference(
                            formulaDraftRef.current,
                            `[@${caption}]`,
                            { start: formulaCursorRef.current, end: formulaCursorRef.current },
                            formulaReferenceRangeRef.current,
                        );
                        formulaReferenceRangeRef.current = inserted.referenceRange;
                        setFormulaDraftValue(inserted.formula, inserted.cursor);
                        setTimeout(() => {
                            const input = inlineFormulaOpenRef.current ? inlineFormulaInputRef.current : formulaInputRef.current;
                            input?.focus();
                            input?.setSelectionRange(inserted.cursor, inserted.cursor);
                            formulaPickingRef.current = false;
                        }, 0);
                    }
                } else {
                    setStatus({ key: "singleColumnReference" });
                }
                setSelection(nextSelection);
                return;
            }
            formulaPickingRef.current = true;
            insertReferenceForSelection(nextSelection);
        } else if (formulaEditingRef.current) {
            // Clicking the grid after merely focusing/editing a non-formula
            // value is a normal selection change, not reference picking.
            endFormulaSession();
        }
        setSelection(nextSelection);
    }, [endFormulaSession, insertReferenceForSelection]);

    const onFormulaDraftChange = (value: string, cursor: number) => {
        formulaEditingRef.current = true;
        formulaReferenceRangeRef.current = undefined;
        setFormulaDraftValue(value, cursor);
    };

    const onFormulaInputBlur = () => {
        // Excel-style formula mode intentionally survives focus moving from
        // the editor to the grid: that blur is how the user starts picking a
        // reference. Enter/Apply/Escape are the explicit session boundaries.
        if (formulaEditingRef.current && formulaDraftRef.current.startsWith("=")) return;
        // Grid reference picking briefly blurs the input. Defer session exit
        // until the grid selection callback has had a chance to restore focus.
        setTimeout(() => {
            if (formulaPickingRef.current) return;
            if (document.activeElement !== formulaInputRef.current && document.activeElement !== inlineFormulaInputRef.current) {
                endFormulaSession();
            }
        }, 50);
    };

    const rememberCopiedGridRange = React.useCallback(() => {
        const range = selection.current?.range;
        if (range === undefined) return;
        const sourceRows = Array.from({ length: range.height }, (_, offset) => pageRows[range.y + offset]).filter((row): row is number => row !== undefined);
        if (sourceRows.length !== range.height) return;
        copiedGridRangeRef.current = {
            ...range,
            sourceRows,
            values: sourceRows.map(sourceRow => Array.from({ length: range.width }, (_, colOffset) => {
                const col = range.x + colOffset;
                const formula = formulas.getFormula(col, sourceRow);
                if (formula !== undefined) return formula;
                const input = formulas.getInput(col, sourceRow);
                return input === LOADING_CELL || input === null ? "" : String(input);
            })),
        };
        setStatus({ key: "copiedCells", values: { count: range.width * range.height } });
    }, [formulas, pageRows, selection]);

    React.useEffect(() => {
        const onCopy = (event: ClipboardEvent) => {
            const target = event.target;
            if (target instanceof HTMLElement && (target.tagName === "INPUT" || target.tagName === "TEXTAREA")) return;
            rememberCopiedGridRange();
        };
        // Capture before DataEditor serializes the cells so the source origin
        // and source-row mapping are retained for Excel-style formula rebasing.
        window.addEventListener("copy", onCopy, true);
        return () => window.removeEventListener("copy", onCopy, true);
    }, [rememberCopiedGridRange]);

    const onGridKeyDown = React.useCallback((event: GridKeyEventArgs) => {
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c" && selection.current?.range !== undefined) {
            rememberCopiedGridRange();
            return;
        }
        if ((event.key !== "=" && event.key !== "Enter") || formulaEditingRef.current || event.location === undefined) return;
        const [col, row] = event.location;
        const sourceRow = pageRows[row];
        if (sourceRow === undefined || spreadsheetColumns[col] === undefined || !isFormulaCapableColumn(col)) return;
        // Take ownership before DataEditor opens its native overlay. Enter
        // starts with the current value selected, so typing `=` immediately
        // replaces it just like Excel; `=` starts with an empty formula.
        event.cancel();
        const current = event.key === "Enter" ? formulas.getInput(col, sourceRow) : "=";
        const initialDraft = current === LOADING_CELL || current === null ? "" : String(current);
        openInlineFormulaSession([col, row], event.key === "Enter" ? initialDraft : "=", event.key === "Enter");
    }, [formulas, openInlineFormulaSession, pageRows, rememberCopiedGridRange, selection]);

    const displayColumns: readonly GridColumn[] = spreadsheetColumns.map((column, index) => {
        const sort = config.sorts.find(item => item.column === column.id)?.direction;
        const marked = config.filters.some(item => item.column === column.id) || config.columnSearches.some(item => item.column === column.id);
        // Keep the user-facing caption in the normal header row and expose
        // spreadsheet letters in a separate group-header tier.
        return { id: column.id, title: `${column.title}${sort === undefined ? "" : sort === "asc" ? " ↑" : " ↓"}${marked ? " •" : ""}`, group: columnIndexToName(index), width: columnWidths[column.id] ?? column.width ?? 140, hasMenu: true, menuIcon: "dots" };
    });
    const headerColumn = headerMenu === undefined ? undefined : spreadsheetColumns[headerMenu.col];
    const stats = source.getStats();
    const inlineFormulaTarget = formulaEditTargetRef.current;
    const inlineFormulaBounds = inlineFormulaOpen && inlineFormulaTarget === undefined
        ? undefined
        : inlineFormulaOpen && inlineFormulaTarget !== undefined
            ? gridRef.current?.getBounds(inlineFormulaTarget.col, inlineFormulaTarget.gridRow)
            : undefined;

    return <div style={{ height: "100vh", display: "flex", flexDirection: "column", fontFamily: "Inter, Arial, sans-serif", color: "#202124" }}>
        <div style={{ padding: 8, display: "flex", gap: 7, alignItems: "center", flexWrap: "wrap", borderBottom: "1px solid #dadce0", background: "#f8fafd" }}>
            <strong>500,000 × 23 Spreadsheet</strong>
            <label>{t("language")} <select aria-label={t("language")} value={locale} onChange={event => setLocale(event.target.value as SpreadsheetLocale)}><option value="en">{t("english")}</option><option value="vi">{t("vietnamese")}</option></select></label>
            <input aria-label={t("searchRows")} placeholder={t("searchPlaceholder")} value={searchDraft} onChange={event => setSearchDraft(event.target.value)} style={{ width: 210, padding: 6 }} />
            <button disabled={queryBusy} onClick={() => void executeQuery({ ...config, search: searchDraft })}>{queryBusy ? t("searching") : t("search")}</button>
            <button onClick={() => { setSearchDraft(""); void executeQuery({ search: "", columnSearches: [], filters: [], sorts: [] }); }}>{t("clearQuery")}</button>
            <button aria-label={t("undo")} disabled={undoRef.current.length === 0} onClick={undo}>{t("undo")}</button>
            <button aria-label={t("redo")} disabled={redoRef.current.length === 0} onClick={redo}>{t("redo")}</button>
            <button aria-expanded={findOpen} onClick={() => setFindOpen(value => !value)}>{t("findReplace")}</button>
            <span style={{ marginLeft: "auto", fontSize: 12, color: "#5f6368" }}>{totalCount.toLocaleString(numberLocale)}/{ROW_COUNT.toLocaleString(numberLocale)} {t("rows")} · {formulas.getFormulaCount().toLocaleString(numberLocale)} {t("formulas")} · {t("query")} {queryMs.toFixed(1)} ms</span>
        </div>
        <div style={{ padding: 7, display: "flex", gap: 7, alignItems: "center", borderBottom: "1px solid #dadce0" }}>
            <code style={{ width: 80 }}>{selectedCell === undefined || selectedSourceRow === undefined ? "—" : `${columnIndexToName(selectedCell[0])}${selectedSourceRow + 1}`}</code>
            <button onClick={autoSum}>Σ {t("autoSum")}</button><span>fx</span>
            <input
                ref={formulaInputRef}
                aria-label={t("formulaBar")}
                value={formulaDraft}
                onFocus={() => {
                    formulaEditingRef.current = true;
                    if (formulaEditTargetRef.current === undefined && selectedCell !== undefined && selectedSourceRow !== undefined) {
                        formulaEditTargetRef.current = { col: selectedCell[0], sourceRow: selectedSourceRow, gridRow: selectedCell[1] };
                    }
                    formulaCursorRef.current = formulaInputRef.current?.selectionStart ?? formulaDraft.length;
                }}
                onBlur={onFormulaInputBlur}
                onClick={event => {
                    formulaCursorRef.current = event.currentTarget.selectionStart ?? formulaDraft.length;
                    formulaReferenceRangeRef.current = undefined;
                }}
                onSelect={event => {
                    formulaCursorRef.current = event.currentTarget.selectionStart ?? formulaDraft.length;
                    if (!formulaPickingRef.current) formulaReferenceRangeRef.current = undefined;
                }}
                onChange={event => onFormulaDraftChange(event.target.value, event.target.selectionStart ?? event.target.value.length)}
                onKeyDown={event => {
                    if (event.key === "Enter") commitFormula();
                    if (event.key === "Escape") {
                        endFormulaSession(true);
                    }
                }}
                placeholder={t("formulaPlaceholder")}
                style={{ flex: 1, padding: 6, fontFamily: "monospace" }}
            />
            <button type="button" aria-label={t("confirm")} title={t("confirm")} onMouseDown={event => event.preventDefault()} onClick={commitFormula}>✓</button>
            <button type="button" aria-label={t("cancel")} title={t("cancel")} onMouseDown={event => event.preventDefault()} onClick={() => endFormulaSession(true)}>✕</button>
        </div>
        {formulaErrorMessage !== undefined && <div role="alert" aria-live="assertive" style={{ margin: "-1px 7px 5px", padding: "6px 8px", border: "1px solid #ef9a9a", borderRadius: 4, background: "#fff4f4", color: "#b91c1c", fontSize: 12 }}>{formulaErrorMessage}</div>}
        {findOpen && <div role="region" aria-label={t("findReplaceRegion")} style={{ padding: 7, display: "flex", gap: 7, alignItems: "center", flexWrap: "wrap", borderBottom: "1px solid #dadce0" }}>
            <input aria-label={t("findText")} placeholder={t("find")} value={findQuery} onChange={event => setFindQuery(event.target.value)} />
            <input aria-label={t("replacementText")} placeholder={t("replaceWith")} value={replaceText} onChange={event => setReplaceText(event.target.value)} />
            <label><input type="checkbox" checked={findCase} onChange={event => setFindCase(event.target.checked)} /> {t("caseSensitive")}</label>
            <label><input type="checkbox" checked={findWhole} onChange={event => setFindWhole(event.target.checked)} /> {t("wholeCell")}</label>
            <label><input type="checkbox" checked={findRegex} onChange={event => setFindRegex(event.target.checked)} /> {t("regex")}</label>
            <button disabled={findBusy || findQuery === ""} onClick={() => void runFind()}>{findBusy ? t("finding") : t("find")}</button>
            <button onClick={() => revealMatch(-1)}>{t("previous")}</button><button onClick={() => revealMatch(1)}>{t("next")}</button>
            <button onClick={() => replaceMatches(false)}>{t("replaceCurrent")}</button><button onClick={() => replaceMatches(true)}>{t("replaceAll")}</button>
            <span>{findTotal.toLocaleString(numberLocale)} {t("matches")}{findIndex >= 0 ? ` · ${findIndex + 1}/${findMatches.length}` : ""}</span>
        </div>}
        <div style={{ padding: 6, display: "flex", gap: 7, alignItems: "center", flexWrap: "wrap", borderBottom: "1px solid #dadce0", fontSize: 12 }}>
            <label>{t("pageSize")} <select aria-label={t("pageSize")} value={pageSize} onChange={event => { setPageSize(Number(event.target.value)); goToPage(0); }}>{PAGE_SIZES.map(size => <option key={size}>{size}</option>)}</select></label>
            <button aria-label={t("firstPage")} disabled={pageIndex === 0} onClick={() => goToPage(0)}>{t("firstPage")}</button><button aria-label={t("previousPage")} disabled={pageIndex === 0} onClick={() => goToPage(pageIndex - 1)}>{t("previousPage")}</button>
            <button aria-label={t("nextPage")} disabled={pageIndex >= pageCount - 1} onClick={() => goToPage(pageIndex + 1)}>{t("nextPage")}</button><button aria-label={t("lastPage")} disabled={pageIndex >= pageCount - 1} onClick={() => goToPage(pageCount - 1)}>{t("lastPage")}</button>
            <label>{t("page")} <input aria-label={t("pageNumber")} type="number" min={1} max={pageCount} value={pageInput} onChange={event => setPageInput(event.target.value)} onKeyDown={event => { if (event.key === "Enter") goToPage(Number(pageInput) - 1); }} style={{ width: 65 }} /> / {pageCount.toLocaleString(numberLocale)}</label>
            <span aria-live="polite">{statusText} · {t("cache")} {stats.cachedPages}/{stats.maxCachedPages} · {t("sparseEdits")} {stats.overlayCells.toLocaleString(numberLocale)}</span>
            <span aria-live="polite" style={{ marginLeft: "auto", color: "#374151" }}>{selectionSummary}</span>
        </div>
        <div style={{ flex: 1, minHeight: 0, position: "relative" }}><DataEditor
            ref={gridRef}
            columns={displayColumns}
            rows={pageRows.length}
            getCellContent={([col, row]) => pageRows[row] === undefined ? loadingCell() : contentFor(source, formulas, pageRows[row]!, col)}
            onCellEdited={([col, row], cell) => {
                const sourceRow = pageRows[row];
                const value = valueFromEditedCell(cell);
                if (sourceRow !== undefined && value !== undefined && col !== 12) {
                    // Text editors are used for formula-capable numeric cells,
                    // so parse ordinary numeric edits while preserving any
                    // value beginning with `=` as a formula source.
                    if (typeof value === "string" && value.startsWith("=")) {
                        const validation = formulas.validateFormula(col, sourceRow, value);
                        if (!validation.valid) {
                            // Native DataEditor overlays commit through this
                            // callback when the user presses Enter. Keep the
                            // rejected text in the formula bar so the user can
                            // fix it, and leave the adapter/history untouched.
                            formulaEditTargetRef.current = { col, sourceRow, gridRow: row };
                            rejectedFormulaEditRef.current = [col, row];
                            formulaEditingRef.current = true;
                            formulaReferenceRangeRef.current = undefined;
                            setFormulaDraftValue(value);
                            const error = validation.error;
                            const message = error === undefined ? t("formulaInvalid") : `${error.code}: ${error.message ?? t("formulaInvalid")}`;
                            setFormulaErrorMessage(message);
                            setStatus({ raw: message });
                            refreshGrid();
                            return;
                        }
                    }
                    applyInput(col, sourceRow, parseFormulaInput(value, spreadsheetColumns[col]!));
                }
            }}
            customRenderers={allCells}
            getCellsForSelection={true}
            onPaste={onGridPaste}
            rowMarkers="number"
            freezeColumns={1}
            gridSelection={selection}
            onGridSelectionChange={onGridSelectionChange}
            onKeyDown={onGridKeyDown}
            onCellClicked={([col, row], event) => {
                const sourceRow = pageRows[row];
                if (sourceRow === undefined) return;

                // A reference click must be handled before DataEditor selects
                // the clicked cell. Keeping the controlled selection on the
                // destination prevents the next edit from being committed to
                // the reference cell instead.
                if (formulaEditingRef.current && formulaDraftRef.current.startsWith("=")) {
                    event.preventDefault();
                    formulaCellClickRef.current = true;
                    formulaPickingRef.current = true;
                    const target = formulaEditTargetRef.current;
                    const caption = spreadsheetColumns[col]?.title;
                    // Same-row references are expressed using the visible
                    // caption, which is the table-style syntax users expect
                    // when the grid does not expose A/B/C headers. Keep A1
                    // syntax for cross-row references until a range is picked.
                    const reference = captionCellReference(col, sourceRow, target?.sourceRow);
                    if (reference === undefined) return;
                    const cursor = formulaCursorRef.current;
                    const inserted = insertFormulaReference(
                        formulaDraftRef.current,
                        reference,
                        { start: cursor, end: cursor },
                        formulaReferenceRangeRef.current,
                    );
                    formulaReferenceRangeRef.current = inserted.referenceRange;
                    setFormulaDraftValue(inserted.formula, inserted.cursor);
                    if (target !== undefined) setSelection(singleCellSelection([target.col, target.gridRow]));
                    setTimeout(() => {
                        const input = inlineFormulaOpenRef.current ? inlineFormulaInputRef.current : formulaInputRef.current;
                        input?.focus();
                        input?.setSelectionRange(inserted.cursor, inserted.cursor);
                        formulaPickingRef.current = false;
                    }, 0);
                    return;
                }

                if (!isFormulaCapableColumn(col)) return;

                // The default activation mode is second-click. Intercept the
                // second click so formula cells use the same lightweight
                // editor as direct `=` entry instead of the native overlay.
                const hasFormula = formulas.hasFormula(col, sourceRow);
                if (shouldOpenInlineFormulaEditorOnDoubleClick(event.isDoubleClick === true, spreadsheetColumns[col]?.dataType) ||
                    (hasFormula && event.isDoubleClick === true)) {
                    event.preventDefault();
                    const current = formulas.getInput(col, sourceRow);
                    const initialDraft = current === LOADING_CELL || current === null ? "" : String(current);
                    openInlineFormulaSession([col, row], initialDraft, true);
                }
            }}
            onHeaderMenuClick={(col, bounds) => openHeader(col, bounds)}
            onHeaderClicked={(col, event) => {
                if (formulaEditingRef.current && formulaDraftRef.current.startsWith("=")) {
                    event.preventDefault();
                    formulaPickingRef.current = true;
                    const caption = spreadsheetColumns[col]?.title;
                    if (caption !== undefined) {
                        const cursor = formulaCursorRef.current;
                        const inserted = insertFormulaReference(
                            formulaDraftRef.current,
                            `[@${caption}]`,
                            { start: cursor, end: cursor },
                            formulaReferenceRangeRef.current,
                        );
                        formulaReferenceRangeRef.current = inserted.referenceRange;
                        setFormulaDraftValue(inserted.formula, inserted.cursor);
                        setTimeout(() => {
                            const input = inlineFormulaOpenRef.current ? inlineFormulaInputRef.current : formulaInputRef.current;
                            input?.focus();
                            input?.setSelectionRange(inserted.cursor, inserted.cursor);
                            formulaPickingRef.current = false;
                        }, 0);
                    }
                    return;
                }
                event.preventDefault();
                openHeader(col, event.bounds, event.shiftKey);
            }}
            onGroupHeaderClicked={(col, event) => {
                if (!formulaEditingRef.current || !formulaDraftRef.current.startsWith("=")) return;
                event.preventDefault();
                const caption = spreadsheetColumns[col]?.title;
                if (caption === undefined) return;
                formulaPickingRef.current = true;
                const cursor = formulaCursorRef.current;
                const inserted = insertFormulaReference(
                    formulaDraftRef.current,
                    `[@${caption}]`,
                    { start: cursor, end: cursor },
                    formulaReferenceRangeRef.current,
                );
                formulaReferenceRangeRef.current = inserted.referenceRange;
                setFormulaDraftValue(inserted.formula, inserted.cursor);
                setTimeout(() => {
                    const input = inlineFormulaOpenRef.current ? inlineFormulaInputRef.current : formulaInputRef.current;
                    input?.focus();
                    input?.setSelectionRange(inserted.cursor, inserted.cursor);
                    formulaPickingRef.current = false;
                }, 0);
            }}
            onColumnResize={onColumnResize}
            minColumnWidth={MIN_COLUMN_WIDTH}
            maxColumnWidth={MAX_COLUMN_WIDTH}
            groupHeaderHeight={28}
            keybindings={{ search: true }}
            searchLabels={builtInSearchLabels}
        />
        {inlineFormulaOpen && inlineFormulaBounds !== undefined && <div
            role="group"
            aria-label={t("inCellFormulaEditorControls")}
            style={{ position: "fixed", zIndex: 25, left: inlineFormulaBounds.x, top: inlineFormulaBounds.y, height: inlineFormulaBounds.height, display: "flex", alignItems: "stretch" }}
        >
            <input
                ref={inlineFormulaInputRef}
                aria-label={t("inCellFormulaEditor")}
                value={formulaDraft}
                onFocus={() => { formulaEditingRef.current = true; }}
                onBlur={onFormulaInputBlur}
                onChange={event => onFormulaDraftChange(event.target.value, event.target.selectionStart ?? event.target.value.length)}
                onClick={event => {
                    formulaCursorRef.current = event.currentTarget.selectionStart ?? formulaDraft.length;
                    formulaReferenceRangeRef.current = undefined;
                }}
                onSelect={event => {
                    formulaCursorRef.current = event.currentTarget.selectionStart ?? formulaDraft.length;
                    if (!formulaPickingRef.current) formulaReferenceRangeRef.current = undefined;
                }}
                onKeyDown={event => {
                    if (event.key === "Enter" || event.key === "Tab") {
                        event.preventDefault();
                        commitFormula();
                    } else if (event.key === "Escape") {
                        event.preventDefault();
                        endFormulaSession(true);
                    }
                }}
                style={{ width: inlineFormulaBounds.width, height: inlineFormulaBounds.height, boxSizing: "border-box", border: "2px solid #1a73e8", outline: "none", padding: "0 6px", fontFamily: "monospace", background: "#fff" }}
            />
            <button type="button" aria-label={t("confirm")} title={t("confirm")} onMouseDown={event => event.preventDefault()} onClick={commitFormula} style={{ width: 25, padding: 0, border: "1px solid #9ca3af", background: "#ecfdf5", color: "#047857", fontSize: 16 }}>✓</button>
            <button type="button" aria-label={t("cancel")} title={t("cancel")} onMouseDown={event => event.preventDefault()} onClick={() => endFormulaSession(true)} style={{ width: 25, padding: 0, border: "1px solid #9ca3af", background: "#fef2f2", color: "#b91c1c", fontSize: 16 }}>✕</button>
        </div>}
        </div>
        {headerMenu !== undefined && headerColumn !== undefined && <div ref={headerMenuRef} role="dialog" aria-label={t("columnMenu", { column: headerColumn.title })} onPointerDown={event => event.stopPropagation()} style={{ position: "fixed", zIndex: 20, left: Math.max(8, Math.min(headerMenu.bounds.x, window.innerWidth - 250)), top: Math.max(8, Math.min(headerMenu.bounds.y + headerMenu.bounds.height + 4, window.innerHeight - 350)), width: 230, padding: 10, display: "flex", flexDirection: "column", gap: 7, border: "1px solid #c7cbd1", borderRadius: 6, background: "white", boxShadow: "0 4px 12px rgba(0,0,0,.2)" }}>
            <strong>{headerColumn.title}</strong><small>{headerMenu.additive ? t("shiftClickKeepsExistingSortKeys") : t("clickSortReplacesExistingKeys")}</small>
            <div style={{ display: "flex", gap: 5 }}><button onClick={() => applySort("asc")}>{t("sortAscending")}</button><button onClick={() => applySort("desc")}>{t("sortDescending")}</button></div>
            <button onClick={() => { const sorts = config.sorts.filter(item => item.column !== headerColumn.id); setHeaderMenu(undefined); void executeQuery({ ...config, sorts }); }}>{t("clearSort")}</button>
            <hr style={{ width: "100%" }} /><label>{t("searchThisColumn")}<input aria-label={t("columnSearch", { column: headerColumn.title })} value={columnSearchDraft} onChange={event => setColumnSearchDraft(event.target.value)} style={{ width: "100%", boxSizing: "border-box" }} /></label><button onClick={applyColumnSearch}>{t("applyColumnSearch")}</button>
            <hr style={{ width: "100%" }} /><label>{t("filter")}<select aria-label={t("filterOperator", { column: headerColumn.title })} value={filterOperator} onChange={event => setFilterOperator(event.target.value as FilterOperator)} style={{ width: "100%" }}>{filterOperators(headerColumn).map(operator => <option key={operator} value={operator}>{i18n.filterOperator(operator)}</option>)}</select></label>
            {headerColumn.type === "boolean" ? <select aria-label={t("filterValue", { column: headerColumn.title })} value={filterValue} onChange={event => setFilterValue(event.target.value)}><option value="">{t("any")}</option><option value="true">{t("trueValue")}</option><option value="false">{t("falseValue")}</option></select> : <input aria-label={t("filterValue", { column: headerColumn.title })} type={headerColumn.type === "number" ? "number" : "search"} value={filterValue} onChange={event => setFilterValue(event.target.value)} />}
            <button onClick={applyFilter}>{t("applyFilter")}</button><button onClick={() => { const filters = config.filters.filter(item => item.column !== headerColumn.id); setHeaderMenu(undefined); void executeQuery({ ...config, filters }); }}>{t("clearFilter")}</button>
        </div>}
    </div>;
}
