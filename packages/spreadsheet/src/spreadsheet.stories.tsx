import * as React from "react";
import {
    CompactSelection,
    DataEditor,
    GridCellKind,
    type EditableGridCell,
    type GridCell,
    type GridSelection,
    type Item,
    type FillPatternEventArgs,
    type Rectangle,
} from "../../core/src/index.js";
import { aggregateValues } from "./aggregate.js";
import { formatCellValue } from "./format.js";
import { smartFill, type FillEdit, type FillRectangle } from "./fill.js";
import { columnIndexToName, isFormulaError } from "./formula.js";
import { diagnoseFormula, getFormulaCompletions, type FormulaCompletion } from "./formula-authoring.js";
import { History, type CellEdit, type HistoryDirection } from "./history.js";
import { SpreadsheetModel, type CellInput, type SpreadsheetColumn } from "./model.js";
import { createSpreadsheetView, type FilterOperator, type SpreadsheetFilter, type SpreadsheetSort } from "./view.js";
import { validateCellInput } from "./validation.js";
import { exportModelToCsv, restoreModelSnapshot, serializeModelSnapshot } from "./workbook-io.js";
import { SnapshotError } from "./persistence.js";
import { compileConditionalFormatting, type ConditionalFormatRule } from "./conditional-format.js";
import { findMatches, replaceAll, replaceOne, type FindMatch, type FindScope } from "./find-replace.js";
import {
    clearColumnFilter,
    clearColumnSearch,
    clearSort,
    compileColumnViewState,
    createColumnViewState,
    setColumnFilter,
    setColumnSearch,
    setGlobalSearch,
    setSort as setColumnViewSort,
    type ColumnViewState,
} from "./column-view-state.js";

export default {
    title: "Extra Packages/Spreadsheet MVP",
};

const spreadsheetColumns: readonly SpreadsheetColumn[] = [
    { id: "item", title: "Item", width: 210, type: "text" },
    { id: "quantity", title: "Quantity", width: 120, type: "number" },
    { id: "unitPrice", title: "Unit Price", width: 140, type: "number" },
    { id: "total", title: "Total", width: 150, type: "number" },
    { id: "status", title: "Status", width: 150, type: "text" },
];

const rows: readonly (readonly CellInput[])[] = [
    ["Mechanical keyboard", 2, 89.5, "=[@Quantity]*[@Unit Price]", '=IF([@Total]>=150,"Large","Normal")'],
    ["Wireless mouse", 5, 32, "=[@Quantity]*[@Unit Price]", '=IF([@Total]>=150,"Large","Normal")'],
    ["USB-C dock", 1, 129, "=[@Quantity]*[@Unit Price]", '=IF([@Total]>=150,"Large","Normal")'],
    ["4K monitor", 3, 399, "=[@Quantity]*[@Unit Price]", '=IF([@Total]>=150,"Large","Normal")'],
    ["Laptop stand", 4, 45, "=[@Quantity]*[@Unit Price]", '=IF([@Total]>=150,"Large","Normal")'],
    ["Webcam", 2, 75, "=[@Quantity]*[@Unit Price]", '=IF([@Total]>=150,"Large","Normal")'],
    ["Headset", 6, 110, "=[@Quantity]*[@Unit Price]", '=IF([@Total]>=150,"Large","Normal")'],
    ["Desk mat", 8, 20, "=[@Quantity]*[@Unit Price]", '=IF([@Total]>=150,"Large","Normal")'],
];

const emptySelection: GridSelection = {
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
};

function parseInput(value: string, column: SpreadsheetColumn): CellInput {
    if (value.startsWith("=")) return value;
    if (column.type === "number") {
        const parsed = Number(value);
        return value.trim() === "" ? null : Number.isFinite(parsed) ? parsed : value;
    }
    if (column.type === "boolean") return value.toLocaleLowerCase() === "true";
    return value;
}

const nonNegativeNumberRules = [{ kind: "number-range" as const, id: "non-negative", min: 0, message: "Must be a number greater than or equal to 0" }];

type HeaderMenuState = { readonly col: number; readonly bounds: Rectangle; readonly additive: boolean };
type ColumnFilterDraft = { readonly operator: FilterOperator; readonly value: string };

const textFilterOperators: readonly FilterOperator[] = ["contains", "equals"];
const numberFilterOperators: readonly FilterOperator[] = ["equals", "gt", "gte", "lt", "lte"];
const booleanFilterOperators: readonly FilterOperator[] = ["equals"];
const findResultLimit = 1_000;

function singleCellSelection(cell: Item): GridSelection {
    return {
        current: { cell, range: { x: cell[0], y: cell[1], width: 1, height: 1 }, rangeStack: [] },
        columns: CompactSelection.empty(),
        rows: CompactSelection.empty(),
    };
}

function filterOperatorsForColumn(column: SpreadsheetColumn): readonly FilterOperator[] {
    if (column.type === "number") return numberFilterOperators;
    if (column.type === "boolean") return booleanFilterOperators;
    return textFilterOperators;
}

function filterDraftFor(column: SpreadsheetColumn, filter: SpreadsheetFilter | undefined): ColumnFilterDraft {
    const fallback = filterOperatorsForColumn(column)[0] ?? "contains";
    return {
        operator: filter !== undefined && filterOperatorsForColumn(column).includes(filter.operator) ? filter.operator : fallback,
        value: filter?.value === undefined ? "" : String(filter.value),
    };
}

function parseHeaderFilter(column: SpreadsheetColumn, draft: ColumnFilterDraft): SpreadsheetFilter | undefined {
    if (draft.value.trim() === "") return undefined;
    if (!filterOperatorsForColumn(column).includes(draft.operator)) return undefined;
    if (column.type === "number") {
        const value = Number(draft.value);
        return Number.isFinite(value) ? { column: column.id, operator: draft.operator, value } : undefined;
    }
    if (column.type === "boolean") {
        if (draft.value !== "true" && draft.value !== "false") return undefined;
        return { column: column.id, operator: "equals", value: draft.value === "true" };
    }
    return { column: column.id, operator: draft.operator, value: draft.value };
}

export function applySpreadsheetHistoryEdit(model: SpreadsheetModel, edit: CellEdit<CellInput>, direction: HistoryDirection): void {
    model.setCell(edit.location[0], edit.location[1], direction === "undo" ? edit.before : edit.after);
}

export function applyFormulaCompletion(source: string, completion: FormulaCompletion): { readonly value: string; readonly cursor: number } {
    const value = source.slice(0, completion.replaceStart) + completion.insertText + source.slice(completion.replaceEnd);
    return { value, cursor: completion.replaceStart + completion.insertText.length };
}

export const InteractiveSpreadsheet: React.FC = () => {
    const modelRef = React.useRef(new SpreadsheetModel(spreadsheetColumns, rows.length, rows));
    const gridRef = React.useRef<React.ElementRef<typeof DataEditor>>(null);
    const historyRef = React.useRef<History<CellInput> | null>(null);
    const transactionIdRef = React.useRef(0);
    const [revision, setRevision] = React.useState(0);
    const [columnViewState, setColumnViewState] = React.useState<ColumnViewState>(() => createColumnViewState());
    const [minimumTotal, setMinimumTotal] = React.useState("");
    const totalSort = columnViewState.sorts.find(item => item.column === "total");
    const sort: "none" | "total-asc" | "total-desc" = totalSort === undefined ? "none" : totalSort.direction === "asc" ? "total-asc" : "total-desc";
    const [headerMenu, setHeaderMenu] = React.useState<HeaderMenuState | undefined>();
    const [filterDraft, setFilterDraft] = React.useState<ColumnFilterDraft>({ operator: "contains", value: "" });
    const [columnSearchDraft, setColumnSearchDraft] = React.useState("");
    const headerMenuRef = React.useRef<HTMLDivElement>(null);
    const [selection, setSelection] = React.useState<GridSelection>(emptySelection);
    const [formulaDraft, setFormulaDraft] = React.useState("");
    const [formulaFocused, setFormulaFocused] = React.useState(false);
    const [formulaCursor, setFormulaCursor] = React.useState(0);
    const [completionIndex, setCompletionIndex] = React.useState(0);
    const formulaInputRef = React.useRef<HTMLInputElement>(null);
    const [statusMessage, setStatusMessage] = React.useState("");
    const [ioPanelOpen, setIoPanelOpen] = React.useState(false);
    const [ioText, setIoText] = React.useState("");
    const [ioError, setIoError] = React.useState("");
    const [findPanelOpen, setFindPanelOpen] = React.useState(false);
    const [findQuery, setFindQuery] = React.useState("");
    const [replacement, setReplacement] = React.useState("");
    const [findScope, setFindScope] = React.useState<FindScope>("raw");
    const [findCaseSensitive, setFindCaseSensitive] = React.useState(false);
    const [findWholeCell, setFindWholeCell] = React.useState(false);
    const [findRegExp, setFindRegExp] = React.useState(false);
    const [findMatchIndex, setFindMatchIndex] = React.useState(-1);
    const [conditionalFormattingEnabled, setConditionalFormattingEnabled] = React.useState(false);
    const [conditionalFormattingPreset, setConditionalFormattingPreset] = React.useState<"negative" | "duplicate" | "errors">("negative");
    const model = modelRef.current;
    if (historyRef.current === null) {
        historyRef.current = new History<CellInput>({
            apply: (edit, direction) => applySpreadsheetHistoryEdit(modelRef.current, edit, direction),
        });
    }
    const history = historyRef.current;

    const conditionalFormattingColumn = selection.current?.cell?.[0] ?? 0;
    const conditionalFormattingTargetColumn = conditionalFormattingEnabled && conditionalFormattingPreset === "duplicate"
        ? conditionalFormattingColumn
        : -1;
    const conditionalFormattingEngine = React.useMemo(() => {
        if (!conditionalFormattingEnabled) return undefined;
        const range = { x: 0, y: 0, width: model.columns.length, height: model.rowCount };
        const rules: ConditionalFormatRule[] = conditionalFormattingPreset === "negative"
            ? [{ id: "negative", kind: "number-compare", range, style: { bg: "#fff0f0", text: "#b71c1c" }, operator: "lt", value: 0 }]
            : conditionalFormattingPreset === "duplicate"
                ? [{ id: "duplicate", kind: "duplicate", range: { x: conditionalFormattingTargetColumn, y: 0, width: 1, height: model.rowCount }, style: { bg: "#fff8e1", text: "#8d6e00" } }]
                : [{ id: "formula-error", kind: "formula-error", range, style: { bg: "#fff0f0", text: "#c62828" } }];
        return compileConditionalFormatting(rules, (col, row) => model.getValue(col, row));
    }, [conditionalFormattingEnabled, conditionalFormattingPreset, conditionalFormattingTargetColumn, model, revision]);
    const formulaAuthoringContext = React.useMemo(() => ({
        resolveColumn: (name: string) => model.resolveColumn(name),
        functionRegistry: model.functionRegistry,
        // Formula authors see captions, while ids remain an internal/persisted
        // concern. The model still resolves legacy id references for imports.
        structuredColumns: model.columns.map(column => column.title),
    }), [model]);

    React.useEffect(() => {
        if (headerMenu === undefined) return;
        const column = model.columns[headerMenu.col];
        if (column === undefined) return;
        setFilterDraft(filterDraftFor(column, columnViewState.filters.find(filter => filter.column === column.id)));
        setColumnSearchDraft(columnViewState.columnSearches.find(searchEntry => searchEntry.column === column.id)?.value ?? "");
    }, [columnViewState, headerMenu, model]);

    React.useEffect(() => {
        if (headerMenu === undefined) return;
        const closeOnEscape = (event: KeyboardEvent) => {
            if (event.key === "Escape") setHeaderMenu(undefined);
        };
        const closeOnOutsidePointer = (event: PointerEvent) => {
            const target = event.target;
            if (!(target instanceof Node) || !headerMenuRef.current?.contains(target)) setHeaderMenu(undefined);
        };
        const focusTimer = window.setTimeout(() => {
            headerMenuRef.current?.querySelector<HTMLElement>("button, input, select")?.focus();
        }, 0);
        document.addEventListener("keydown", closeOnEscape);
        document.addEventListener("pointerdown", closeOnOutsidePointer);
        return () => {
            window.clearTimeout(focusTimer);
            document.removeEventListener("keydown", closeOnEscape);
            document.removeEventListener("pointerdown", closeOnOutsidePointer);
        };
    }, [headerMenu]);

    const compiledColumnView = React.useMemo(() => compileColumnViewState(columnViewState, {
        columns: model.columns.map(column => ({ id: column.id, title: column.title })),
    }), [columnViewState, model]);
    const viewFilters = React.useMemo<readonly SpreadsheetFilter[]>(() => [
        ...compiledColumnView.filters,
        ...(minimumTotal.trim() === "" ? [] : [{ column: "total", operator: "gte" as const, value: Number(minimumTotal) }]),
    ], [compiledColumnView.filters, minimumTotal]);
    const viewSorts = React.useMemo<readonly SpreadsheetSort[]>(() => [
        ...compiledColumnView.sorts,
    ], [compiledColumnView.sorts]);

    const visibleRows = React.useMemo(() => createSpreadsheetView(model, {
        search: compiledColumnView.search,
        columnSearches: compiledColumnView.columnSearches,
        filters: viewFilters,
        sorts: viewSorts,
    }), [compiledColumnView.columnSearches, compiledColumnView.search, model, revision, viewFilters, viewSorts]);

    const findResult = React.useMemo(() => findQuery === "" ? undefined : findMatches(model, {
        query: findQuery,
        scope: findScope,
        caseSensitive: findCaseSensitive,
        wholeCell: findWholeCell,
        regexp: findRegExp,
        maxResults: findResultLimit,
    }), [findCaseSensitive, findQuery, findRegExp, findScope, findWholeCell, model, revision]);

    // DataEditor may retain rendered cells between React renders. Explicitly
    // invalidate the visible grid when the conditional rule or its selected
    // duplicate column changes so styles never lag behind the control state.
    React.useEffect(() => {
        const cells = visibleRows.flatMap((_, row) => model.columns.map((_, col) => ({ cell: [col, row] as [number, number] })));
        gridRef.current?.updateCells(cells);
    }, [conditionalFormattingEngine, model, visibleRows]);

    const selectedCell = selection.current?.cell;
    React.useEffect(() => {
        if (selectedCell === undefined) {
            setFormulaDraft("");
            return;
        }
        const sourceRow = visibleRows[selectedCell[1]];
        setFormulaDraft(sourceRow === undefined ? "" : String(model.getInput(selectedCell[0], sourceRow) ?? ""));
    }, [model, revision, selectedCell, visibleRows]);

    const formulaDiagnostics = React.useMemo(() => {
        if (!formulaDraft.startsWith("=")) return [];
        return diagnoseFormula(formulaDraft, formulaAuthoringContext);
    }, [formulaAuthoringContext, formulaDraft]);
    const formulaCompletions = React.useMemo(() => {
        if (!formulaFocused || !formulaDraft.startsWith("=")) return [];
        return getFormulaCompletions(formulaDraft, formulaCursor, formulaAuthoringContext).slice(0, 6);
    }, [formulaAuthoringContext, formulaCursor, formulaDraft, formulaFocused]);

    const commit = React.useCallback((cell: Item, value: string): boolean => {
        const sourceRow = visibleRows[cell[1]];
        const column = model.columns[cell[0]];
        if (sourceRow === undefined || column === undefined) return false;
        const nextValue = parseInput(value, column);
        if (column.id === "quantity" || column.id === "unitPrice") {
            // Formula inputs are intentionally deferred; literal Quantity and
            // Unit Price values still have to satisfy the non-negative rule.
            const validation = validateCellInput(nextValue, nonNegativeNumberRules, { formulaPolicy: "skip" });
            if (!validation.valid) {
                setStatusMessage(validation.errors[0]?.message ?? "Quantity and Unit Price must be non-negative numbers");
                return false;
            }
        }
        const before = model.getInput(cell[0], sourceRow);
        const transaction = history.execute({
            id: `edit-${++transactionIdRef.current}`,
            label: "Edit cell",
            edits: [{ location: [cell[0], sourceRow], before, after: nextValue }],
        });
        if (transaction === undefined) return false;
        setStatusMessage("");
        setRevision(current => current + 1);
        gridRef.current?.updateCells(model.columns.map((_, col) => ({ cell: [col, cell[1]] })));
        return true;
    }, [history, model, visibleRows]);

    const getCellContent = React.useCallback((cell: Item): GridCell => {
        const sourceRow = visibleRows[cell[1]];
        if (sourceRow === undefined) return { kind: GridCellKind.Loading, allowOverlay: false };
        const value = model.getValue(cell[0], sourceRow);
        const input = model.getInput(cell[0], sourceRow);
        const conditionalStyle = conditionalFormattingEngine?.evaluate(cell[0], sourceRow, value).style;
        const displayData = isFormulaError(value)
            ? value.code
            : value === null
                ? ""
                : typeof value === "number"
                    ? formatCellValue(value, { kind: "number", maximumFractionDigits: 2 }, "vi-VN")
                    : String(value);
        const baseTheme = isFormulaError(value) ? { textDark: "#c62828", bgCell: "#fff4f4" } : input !== value ? { bgCell: "#f7fbff" } : undefined;
        const conditionalTheme = conditionalStyle === undefined ? undefined : {
            ...(conditionalStyle.bg === undefined ? {} : { bgCell: conditionalStyle.bg }),
            ...(conditionalStyle.text === undefined ? {} : { textDark: conditionalStyle.text }),
        };
        const themeOverride = baseTheme === undefined && conditionalTheme === undefined ? undefined : { ...(baseTheme ?? {}), ...(conditionalTheme ?? {}) };
        return {
            kind: GridCellKind.Text,
            allowOverlay: true,
            data: typeof input === "string" && input.startsWith("=") ? input : displayData,
            displayData,
            contentAlign: model.columns[cell[0]].type === "number" ? "right" : "left",
            themeOverride,
        };
    }, [conditionalFormattingEngine, model, revision, visibleRows]);

    const onCellEdited = React.useCallback((cell: Item, value: EditableGridCell) => {
        commit(cell, String(value.data ?? ""));
    }, [commit]);

    const applyFormulaBar = () => {
        if (selectedCell !== undefined) commit(selectedCell, formulaDraft);
    };

    const applyCompletion = (completion: FormulaCompletion) => {
        const applied = applyFormulaCompletion(formulaDraft, completion);
        setFormulaDraft(applied.value);
        setFormulaCursor(applied.cursor);
        setCompletionIndex(0);
        // Keep the editor focused after a mouse or keyboard completion.
        setTimeout(() => {
            formulaInputRef.current?.focus();
            formulaInputRef.current?.setSelectionRange(applied.cursor, applied.cursor);
        }, 0);
    };

    const onFormulaKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
        if (event.key === "ArrowDown" && formulaCompletions.length > 0) {
            event.preventDefault();
            setCompletionIndex(index => (index + 1) % formulaCompletions.length);
        } else if (event.key === "ArrowUp" && formulaCompletions.length > 0) {
            event.preventDefault();
            setCompletionIndex(index => (index - 1 + formulaCompletions.length) % formulaCompletions.length);
        } else if (event.key === "Tab" && formulaCompletions.length > 0) {
            event.preventDefault();
            applyCompletion(formulaCompletions[completionIndex] ?? formulaCompletions[0]);
        } else if (event.key === "Escape") {
            setFormulaFocused(false);
            setCompletionIndex(0);
        } else if (event.key === "Enter") {
            // Keep Enter's existing formula-bar Apply behavior.
            applyFormulaBar();
        }
    };

    const autoSum = () => {
        if (selectedCell === undefined || selectedCell[1] === 0) return;
        const sourceRow = visibleRows[selectedCell[1]];
        if (sourceRow === undefined) return;
        const column = columnIndexToName(selectedCell[0]);
        const formula = `=SUM(${column}1:${column}${sourceRow})`;
        setFormulaDraft(formula);
        commit(selectedCell, formula);
    };

    const onFillPattern = React.useCallback((event: FillPatternEventArgs) => {
        // A filtered or sorted view no longer has a safe one-to-one row mapping
        // between grid coordinates and model coordinates.
        if (columnViewState.globalSearch.trim() !== "" || minimumTotal.trim() !== "" || columnViewState.filters.length > 0 || columnViewState.sorts.length > 0) {
            event.preventDefault();
            setStatusMessage("Custom fill is disabled while rows are filtered or sorted.");
            return;
        }
        const toSourceRectangle = (rectangle: FillRectangle): FillRectangle | undefined => {
            if (rectangle.x < 0 || rectangle.y < 0 || rectangle.width <= 0 || rectangle.height <= 0 ||
                rectangle.x + rectangle.width > model.columns.length || rectangle.y + rectangle.height > visibleRows.length) return undefined;
            const mappedRows = visibleRows.slice(rectangle.y, rectangle.y + rectangle.height);
            if (mappedRows.length !== rectangle.height || mappedRows.some((row, index) => row !== mappedRows[0] + index)) return undefined;
            return { ...rectangle, y: mappedRows[0] };
        };
        const source = toSourceRectangle(event.patternSource);
        const destination = toSourceRectangle(event.fillDestination);
        if (source === undefined || destination === undefined) {
            event.preventDefault();
            setStatusMessage("Custom fill is unavailable for this row selection.");
            return;
        }
        const pattern = Array.from({ length: source.height }, (_, row) =>
            Array.from({ length: source.width }, (_, col) => model.getInput(source.x + col, source.y + row))
        );
        let fillEdits: FillEdit[];
        try {
            fillEdits = smartFill(pattern, source, destination);
        } catch {
            event.preventDefault();
            setStatusMessage("Unable to apply this fill pattern.");
            return;
        }
        const edits = fillEdits.map(([col, row, after]) => ({
            location: [col, row] as readonly [number, number],
            before: model.getInput(col, row),
            after,
        }));
        for (const edit of edits) {
            const column = model.columns[edit.location[0]];
            if ((column?.id === "quantity" || column?.id === "unitPrice") && !validateCellInput(edit.after, nonNegativeNumberRules, { formulaPolicy: "skip" }).valid) {
                event.preventDefault();
                setStatusMessage("Custom fill cannot create a negative Quantity or Unit Price.");
                return;
            }
        }
        event.preventDefault();
        const transaction = history.execute({ id: `fill-${++transactionIdRef.current}`, label: "Fill pattern", edits });
        if (transaction === undefined) return;
        setStatusMessage("");
        setRevision(current => current + 1);
        gridRef.current?.updateCells(edits.map(edit => ({ cell: [edit.location[0], visibleRows.indexOf(edit.location[1])] as Item })));
    }, [columnViewState, history, minimumTotal, model, visibleRows]);

    const selectionStats = React.useMemo(() => {
        const range = selection.current?.range;
        if (range === undefined) return "Select cells to calculate";
        const selectedValues: CellInput[] = [];
        for (let row = range.y; row < range.y + range.height; row++) {
            const sourceRow = visibleRows[row];
            if (sourceRow === undefined) continue;
            for (let col = range.x; col < range.x + range.width; col++) {
                selectedValues.push(model.getValue(col, sourceRow));
            }
        }
        const stats = aggregateValues(selectedValues);
        const sum = formatCellValue(stats.sum, { kind: "number", maximumFractionDigits: 2 }, "vi-VN");
        const average = formatCellValue(stats.average, { kind: "number", maximumFractionDigits: 2 }, "vi-VN");
        return `Count: ${stats.countNumbers}  ·  Sum: ${sum}  ·  Average: ${average}`;
    }, [model, revision, selection, visibleRows]);

    const undo = () => {
        if (history.undo() !== undefined) setRevision(current => current + 1);
    };

    const redo = () => {
        if (history.redo() !== undefined) setRevision(current => current + 1);
    };

    const onStoryKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
        if (!(event.ctrlKey || event.metaKey) || event.altKey || event.key.toLocaleLowerCase() !== "z") return;
        event.preventDefault();
        if (event.shiftKey) redo();
        else undo();
    };

    const openHeaderMenu = (col: number, bounds: Rectangle, additive = false) => {
        const column = model.columns[col];
        if (column === undefined) return;
        setFilterDraft(filterDraftFor(column, columnViewState.filters.find(filter => filter.column === column.id)));
        setColumnSearchDraft(columnViewState.columnSearches.find(searchEntry => searchEntry.column === column.id)?.value ?? "");
        setHeaderMenu({ col, bounds, additive });
    };

    const setHeaderSort = (direction: SpreadsheetSort["direction"]) => {
        if (headerMenu === undefined) return;
        const column = model.columns[headerMenu.col];
        if (column === undefined) return;
        setColumnViewState(current => setColumnViewSort(current, column.id, direction, { additive: headerMenu.additive }));
        setHeaderMenu(undefined);
    };

    const clearHeaderSort = () => {
        if (headerMenu === undefined) return;
        const column = model.columns[headerMenu.col];
        if (column === undefined) return;
        setColumnViewState(current => clearSort(current, column.id));
        setHeaderMenu(undefined);
    };

    const applyHeaderFilter = () => {
        if (headerMenu === undefined) return;
        const column = model.columns[headerMenu.col];
        if (column === undefined) return;
        if (filterDraft.value.trim() === "") {
            setColumnViewState(current => clearColumnFilter(current, column.id));
            setStatusMessage("");
            setHeaderMenu(undefined);
            return;
        }
        const filter = parseHeaderFilter(column, filterDraft);
        if (filter === undefined) {
            setStatusMessage(column.type === "number" ? "Enter a finite number for this filter." : "Choose a valid filter value.");
            return;
        }
        setColumnViewState(current => setColumnFilter(current, column.id, filter.operator, filter.value));
        setStatusMessage("");
        setHeaderMenu(undefined);
    };

    const clearHeaderFilter = () => {
        if (headerMenu === undefined) return;
        const column = model.columns[headerMenu.col];
        if (column === undefined) return;
        setColumnViewState(current => clearColumnFilter(current, column.id));
        setStatusMessage("");
        setHeaderMenu(undefined);
    };

    const applyHeaderColumnSearch = () => {
        if (headerMenu === undefined) return;
        const column = model.columns[headerMenu.col];
        if (column === undefined) return;
        setColumnViewState(current => setColumnSearch(current, column.id, columnSearchDraft));
        setStatusMessage("");
    };

    const clearHeaderColumnSearch = () => {
        if (headerMenu === undefined) return;
        const column = model.columns[headerMenu.col];
        if (column === undefined) return;
        setColumnViewState(current => clearColumnSearch(current, column.id));
        setColumnSearchDraft("");
        setStatusMessage("");
    };

    const findOptions = () => ({
        query: findQuery,
        scope: findScope,
        caseSensitive: findCaseSensitive,
        wholeCell: findWholeCell,
        regexp: findRegExp,
        maxResults: findResultLimit,
    });

    const revealFindMatch = (match: FindMatch): boolean => {
        const visibleRow = visibleRows.indexOf(match.row);
        if (visibleRow < 0) {
            setStatusMessage(`Match ${columnIndexToName(match.col)}${match.row + 1} is hidden by the current filters.`);
            return false;
        }
        const cell: Item = [match.col, visibleRow];
        setSelection(singleCellSelection(cell));
        gridRef.current?.scrollTo(match.col, visibleRow, "both", 8, 8);
        gridRef.current?.focus();
        setStatusMessage(`Match ${columnIndexToName(match.col)}${match.row + 1}`);
        return true;
    };

    const moveFindMatch = (delta: number) => {
        const matches = findResult?.matches ?? [];
        if (findResult?.diagnostic !== undefined) {
            setStatusMessage(findResult.diagnostic.message);
            return;
        }
        if (matches.length === 0) {
            setStatusMessage(findQuery === "" ? "Enter text to find." : "No matches found.");
            return;
        }
        const next = findMatchIndex < 0
            ? delta < 0 ? matches.length - 1 : 0
            : ((findMatchIndex + delta) % matches.length + matches.length) % matches.length;
        setFindMatchIndex(next);
        revealFindMatch(matches[next]);
    };

    const coerceReplacement = (text: string, match: FindMatch): CellInput => {
        const column = model.columns[match.col];
        return column === undefined ? text : parseInput(text, column);
    };

    const executeReplacement = (all: boolean) => {
        if (findScope !== "raw") {
            setStatusMessage("Replacement is available only when searching raw inputs/formulas.");
            return;
        }
        const currentMatch = findResult?.matches[findMatchIndex];
        const result = all
            ? replaceAll(model, { ...findOptions(), replacement, coerceReplacement, transactionId: `replace-all-${++transactionIdRef.current}` })
            : currentMatch === undefined
                ? undefined
                : replaceOne(model, {
                    ...findOptions(),
                    columns: [currentMatch.col],
                    rowRange: { start: currentMatch.row, end: currentMatch.row },
                    replacement,
                    coerceReplacement,
                    transactionId: `replace-one-${++transactionIdRef.current}`,
                });
        if (result === undefined) {
            setStatusMessage("Find a match before replacing.");
            return;
        }
        if (result.diagnostic !== undefined) {
            setStatusMessage(result.diagnostic.message);
            return;
        }
        if (result.truncated) {
            setStatusMessage(`Replace All stopped at ${findResultLimit} matches and was not applied. Narrow the search first.`);
            return;
        }
        if (result.transaction === undefined) {
            setStatusMessage(result.skippedCount > 0 ? `No cells replaced; ${result.skippedCount} replacement(s) could not be converted.` : "No cells changed.");
            return;
        }
        history.execute(result.transaction);
        setRevision(current => current + 1);
        setFindMatchIndex(-1);
        setStatusMessage(`${result.replacedCount} cell${result.replacedCount === 1 ? "" : "s"} replaced.`);
        gridRef.current?.updateCells(visibleRows.flatMap((_, row) => model.columns.map((__, col) => ({ cell: [col, row] as Item }))));
    };

    const generateRawCsv = () => {
        setIoText(exportModelToCsv(model, { valueMode: "inputs" }));
        setIoError("");
    };

    const generateComputedCsv = () => {
        setIoText(exportModelToCsv(model, { valueMode: "values" }));
        setIoError("");
    };

    const generateSnapshot = () => {
        setIoText(serializeModelSnapshot(model));
        setIoError("");
    };

    const restoreSnapshot = () => {
        try {
            // Parse and construct before changing refs, so bad JSON leaves the
            // current workbook and its history untouched.
            const restored = restoreModelSnapshot(ioText);
            const maxHistory = history.maxHistory;
            modelRef.current = restored;
            historyRef.current = new History<CellInput>({
                maxHistory,
                apply: (edit, direction) => applySpreadsheetHistoryEdit(modelRef.current, edit, direction),
            });
            setSelection(emptySelection);
            setColumnViewState(createColumnViewState());
            setHeaderMenu(undefined);
            setColumnSearchDraft("");
            setFormulaDraft("");
            setStatusMessage("");
            setIoError("");
            setRevision(current => current + 1);
        } catch (error) {
            const details = error instanceof SnapshotError
                ? `${error.code} at ${error.path}: ${error.message}`
                : error instanceof Error ? error.message : String(error);
            setIoError(`Restore failed: ${details}`);
        }
    };

    const headerMenuColumn = headerMenu === undefined ? undefined : model.columns[headerMenu.col];
    const headerMenuOperators = headerMenuColumn === undefined ? [] : filterOperatorsForColumn(headerMenuColumn);
    const headerMenuLeft = headerMenu === undefined || typeof window === "undefined"
        ? headerMenu?.bounds.x ?? 0
        : Math.max(8, Math.min(headerMenu.bounds.x, window.innerWidth - 254));
    const headerMenuTop = headerMenu === undefined || typeof window === "undefined"
        ? (headerMenu?.bounds.y ?? 0) + (headerMenu?.bounds.height ?? 0) + 4
        : Math.max(8, Math.min(headerMenu.bounds.y + headerMenu.bounds.height + 4, window.innerHeight - 380));

    return (
        <div onKeyDown={onStoryKeyDown} style={{ height: "100vh", display: "flex", flexDirection: "column", fontFamily: "Inter, Arial, sans-serif", color: "#202124" }}>
            <div style={{ padding: "10px 14px", display: "flex", gap: 10, alignItems: "center", borderBottom: "1px solid #dadce0", background: "#f8fafd" }}>
                <strong style={{ marginRight: 8 }}>Spreadsheet MVP</strong>
                <input aria-label="Search rows" placeholder="Search all columns…" value={columnViewState.globalSearch} onChange={event => setColumnViewState(current => setGlobalSearch(current, event.target.value))} style={{ width: 220, padding: 7 }} />
                <input aria-label="Minimum total" placeholder="Minimum total" type="number" value={minimumTotal} onChange={event => setMinimumTotal(event.target.value)} style={{ width: 130, padding: 7 }} />
                <select aria-label="Sort rows" value={sort} onChange={event => {
                    const next = event.target.value as typeof sort;
                    setColumnViewState(current => next === "none"
                        ? clearSort(current, "total")
                        : setColumnViewSort(current, "total", next === "total-asc" ? "asc" : "desc"));
                }} style={{ padding: 7 }}>
                    <option value="none">Original order</option>
                    <option value="total-asc">Total ↑</option>
                    <option value="total-desc">Total ↓</option>
                </select>
                <button aria-label="Undo" onClick={undo} disabled={!history.canUndo()}>Undo</button>
                <button aria-label="Redo" onClick={redo} disabled={!history.canRedo()}>Redo</button>
                <label style={{ display: "flex", alignItems: "center", gap: 4, marginLeft: 6 }}>
                    <input aria-label="Enable conditional formatting" type="checkbox" checked={conditionalFormattingEnabled} onChange={event => setConditionalFormattingEnabled(event.target.checked)} />
                    Conditional formatting
                </label>
                <select aria-label="Conditional formatting preset" value={conditionalFormattingPreset} onChange={event => setConditionalFormattingPreset(event.target.value as typeof conditionalFormattingPreset)} disabled={!conditionalFormattingEnabled} style={{ padding: 5 }}>
                    <option value="negative">Negative numbers</option>
                    <option value="duplicate">Duplicates in selected column</option>
                    <option value="errors">Formula errors</option>
                </select>
                <span aria-live="polite" style={{ color: "#5f6368", fontSize: 12 }}>
                    {conditionalFormattingEnabled ? `Formatting: ${conditionalFormattingPreset === "negative" ? "negative numbers" : conditionalFormattingPreset === "duplicate" ? "duplicates" : "formula errors"}` : "Formatting off"}
                </span>
                <span style={{ marginLeft: "auto", color: "#5f6368", fontSize: 13 }}>{visibleRows.length}/{model.rowCount} rows · {model.getFormulaCount()} formulas</span>
            </div>
            <div style={{ padding: "7px 14px", display: "flex", gap: 8, borderBottom: "1px solid #dadce0", alignItems: "center" }}>
                <code style={{ width: 54, color: "#5f6368" }}>{selectedCell === undefined ? "—" : `${columnIndexToName(selectedCell[0])}${(visibleRows[selectedCell[1]] ?? 0) + 1}`}</code>
                <button onClick={autoSum} title="Sum preceding cells in the selected column">Σ AutoSum</button>
                <span style={{ color: "#5f6368" }}>fx</span>
                <div style={{ flex: 1, position: "relative" }}>
                    <input
                        ref={formulaInputRef}
                        aria-label="Formula bar"
                        value={formulaDraft}
                        onFocus={event => { setFormulaFocused(true); setFormulaCursor(event.currentTarget.selectionStart ?? formulaDraft.length); }}
                        onBlur={() => setTimeout(() => setFormulaFocused(false), 0)}
                        onChange={event => { setFormulaDraft(event.target.value); setFormulaCursor(event.currentTarget.selectionStart ?? event.target.value.length); setCompletionIndex(0); }}
                        onSelect={event => setFormulaCursor(event.currentTarget.selectionStart ?? formulaDraft.length)}
                        onKeyDown={onFormulaKeyDown}
                        style={{ width: "100%", boxSizing: "border-box", padding: 7, fontFamily: "monospace" }}
                        placeholder='Try =A1*B1 or =[@Quantity]*[@Unit Price]'
                    />
                    {formulaFocused && formulaCompletions.length > 0 && <div role="listbox" aria-label="Formula completions" style={{ position: "absolute", zIndex: 2, left: 0, right: 0, top: "100%", maxHeight: 180, overflowY: "auto", border: "1px solid #dadce0", background: "#fff", boxShadow: "0 2px 6px rgba(0,0,0,.15)" }}>
                        {formulaCompletions.map((completion, index) => <button
                            key={`${completion.kind}-${completion.label}`}
                            type="button"
                            role="option"
                            aria-selected={index === completionIndex}
                            onMouseDown={event => { event.preventDefault(); applyCompletion(completion); }}
                            style={{ display: "block", width: "100%", padding: "5px 8px", border: 0, background: index === completionIndex ? "#eef3fd" : "#fff", textAlign: "left", cursor: "pointer" }}
                        >{completion.label}<span style={{ marginLeft: 8, color: "#5f6368", fontSize: 12 }}>{completion.detail}</span></button>)}
                    </div>}
                </div>
                <button onClick={applyFormulaBar}>Apply</button>
                {formulaDiagnostics[0] !== undefined && <span role="alert" aria-live="polite" style={{ color: "#c62828", fontSize: 12, maxWidth: 240 }}>{formulaDiagnostics[0].message}</span>}
            </div>
            <div style={{ padding: "6px 14px", borderBottom: "1px solid #dadce0", background: "#fbfcfe" }}>
                <div style={{ display: "flex", gap: 6 }}>
                    <button aria-expanded={findPanelOpen} aria-controls="spreadsheet-find-panel" onClick={() => setFindPanelOpen(open => !open)}>
                        {findPanelOpen ? "Hide Find / Replace" : "Find / Replace"}
                    </button>
                    <button aria-expanded={ioPanelOpen} aria-controls="spreadsheet-io-panel" onClick={() => setIoPanelOpen(open => !open)}>
                        {ioPanelOpen ? "Hide workbook I/O" : "Show workbook I/O"}
                    </button>
                </div>
                {findPanelOpen && <div id="spreadsheet-find-panel" role="region" aria-label="Find and replace" style={{ marginTop: 7, display: "flex", flexDirection: "column", gap: 6 }}>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                        <input aria-label="Find text" placeholder="Find" value={findQuery} onChange={event => { setFindQuery(event.target.value); setFindMatchIndex(-1); }} style={{ width: 180, padding: 6 }} />
                        <input aria-label="Replacement text" placeholder="Replace with" value={replacement} onChange={event => setReplacement(event.target.value)} style={{ width: 180, padding: 6 }} />
                        <select aria-label="Find scope" value={findScope} onChange={event => { setFindScope(event.target.value as FindScope); setFindMatchIndex(-1); }} style={{ padding: 6 }}>
                            <option value="raw">Raw values / formulas</option>
                            <option value="computed">Computed display values</option>
                        </select>
                        <label><input type="checkbox" checked={findCaseSensitive} onChange={event => { setFindCaseSensitive(event.target.checked); setFindMatchIndex(-1); }} /> Case sensitive</label>
                        <label><input type="checkbox" checked={findWholeCell} onChange={event => { setFindWholeCell(event.target.checked); setFindMatchIndex(-1); }} /> Whole cell</label>
                        <label><input type="checkbox" checked={findRegExp} onChange={event => { setFindRegExp(event.target.checked); setFindMatchIndex(-1); }} /> Regex</label>
                    </div>
                    <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                        <button type="button" onClick={() => moveFindMatch(-1)}>Previous</button>
                        <button type="button" onClick={() => moveFindMatch(1)}>Next</button>
                        <button type="button" onClick={() => executeReplacement(false)} disabled={findScope !== "raw"}>Replace current cell</button>
                        <button type="button" onClick={() => executeReplacement(true)} disabled={findScope !== "raw"}>Replace all</button>
                        <span aria-live="polite" style={{ color: findResult?.diagnostic === undefined ? "#5f6368" : "#c62828", fontSize: 12 }}>
                            {findQuery === ""
                                ? "Enter a query"
                                : findResult?.diagnostic?.message ?? `${findResult?.matches.length ?? 0}${findResult?.truncated ? "+" : ""} matches${findResult !== undefined && findResult.matches.length > 0 && findMatchIndex >= 0 ? ` · current ${Math.min(findMatchIndex + 1, findResult.matches.length)}` : ""}`}
                        </span>
                    </div>
                    <span style={{ color: "#5f6368", fontSize: 12 }}>Find scans the whole source sheet. Matches hidden by current filters are reported; Replace All includes hidden rows and is blocked when the 1,000-match safety cap is reached.</span>
                </div>}
                {ioPanelOpen && <div id="spreadsheet-io-panel" style={{ marginTop: 7, display: "flex", flexDirection: "column", gap: 6 }}>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        <button aria-label="Generate raw formula CSV" onClick={generateRawCsv}>Raw formula CSV</button>
                        <button aria-label="Generate computed value CSV" onClick={generateComputedCsv}>Computed CSV</button>
                        <button aria-label="Generate JSON snapshot" onClick={generateSnapshot}>JSON snapshot</button>
                        <button aria-label="Restore JSON snapshot" onClick={restoreSnapshot}>Restore</button>
                    </div>
                    <textarea
                        aria-label="Workbook CSV or JSON snapshot"
                        value={ioText}
                        onChange={event => setIoText(event.target.value)}
                        rows={3}
                        spellCheck={false}
                        placeholder="Generate CSV/snapshot, or paste a JSON snapshot here"
                        style={{ width: "100%", boxSizing: "border-box", padding: 7, fontFamily: "monospace", resize: "vertical" }}
                    />
                    {ioError !== "" && <span role="alert" style={{ color: "#c62828" }}>{ioError}</span>}
                </div>}
            </div>
            <div style={{ flex: 1, minHeight: 0 }}>
                <DataEditor
                    ref={gridRef}
                    columns={model.columns.map(column => {
                        const columnSort = columnViewState.sorts.find(item => item.column === column.id)?.direction;
                        const hasFilter = columnViewState.filters.some(item => item.column === column.id);
                        const hasColumnSearch = columnViewState.columnSearches.some(item => item.column === column.id);
                        const indicator = `${columnSort === undefined ? "" : ` ${columnSort === "asc" ? "↑" : "↓"}`}${hasFilter || hasColumnSearch ? " •" : ""}`;
                        return { id: column.id, title: `${column.title}${indicator}`, width: column.width ?? 140, hasMenu: true, menuIcon: "dots" };
                    })}
                    rows={visibleRows.length}
                    getCellContent={getCellContent}
                    getCellsForSelection={true}
                    onCellEdited={onCellEdited}
                    onHeaderMenuClick={(col, bounds) => openHeaderMenu(col, bounds)}
                    onHeaderClicked={(col, event) => {
                        event.preventDefault();
                        openHeaderMenu(col, event.bounds, event.shiftKey);
                    }}
                    onPaste={true}
                    fillHandle={true}
                    onFillPattern={onFillPattern}
                    rowMarkers="number"
                    freezeColumns={1}
                    gridSelection={selection}
                    onGridSelectionChange={setSelection}
                    keybindings={{ search: true, downFill: true, rightFill: true }}
                />
            </div>
            {headerMenu !== undefined && headerMenuColumn !== undefined && <div
                ref={headerMenuRef}
                role="dialog"
                aria-label={`${headerMenuColumn.title} column menu`}
                tabIndex={-1}
                onPointerDown={event => event.stopPropagation()}
                style={{
                    position: "fixed",
                    zIndex: 10,
                    left: headerMenuLeft,
                    top: headerMenuTop,
                    width: 238,
                    padding: 10,
                    display: "flex",
                    flexDirection: "column",
                    gap: 7,
                    border: "1px solid #c7cbd1",
                    borderRadius: 6,
                    background: "#fff",
                    boxShadow: "0 4px 12px rgba(32,33,36,.2)",
                    maxHeight: "calc(100vh - 16px)",
                    overflowY: "auto",
                }}
            >
                <strong>{headerMenuColumn.title}</strong>
                <span style={{ color: "#5f6368", fontSize: 12 }}>{headerMenu.additive ? "Shift-click: keep existing sort keys" : "Sort replaces existing keys (Shift-click to add)"}</span>
                <div style={{ display: "flex", gap: 5 }}>
                    <button type="button" onClick={() => setHeaderSort("asc")}>Sort ascending</button>
                    <button type="button" onClick={() => setHeaderSort("desc")}>Sort descending</button>
                </div>
                <button type="button" onClick={clearHeaderSort}>Clear sort</button>
                <div style={{ borderTop: "1px solid #e1e4e8", paddingTop: 7 }}>
                    <label htmlFor="column-search" style={{ display: "block", fontSize: 12, color: "#5f6368" }}>Search this column</label>
                    <input
                        id="column-search"
                        aria-label={`${headerMenuColumn.title} search`}
                        type="search"
                        value={columnSearchDraft}
                        placeholder={`Search ${headerMenuColumn.title}`}
                        onChange={event => setColumnSearchDraft(event.target.value)}
                        style={{ width: "100%", boxSizing: "border-box", marginTop: 3, padding: 5 }}
                    />
                    <div style={{ display: "flex", gap: 5, marginTop: 5 }}>
                        <button type="button" onClick={applyHeaderColumnSearch}>Apply search</button>
                        <button type="button" onClick={clearHeaderColumnSearch}>Clear search</button>
                    </div>
                </div>
                <div style={{ borderTop: "1px solid #e1e4e8", paddingTop: 7 }}>
                    <label htmlFor="column-filter-operator" style={{ display: "block", fontSize: 12, color: "#5f6368" }}>Filter</label>
                    <select
                        id="column-filter-operator"
                        aria-label={`${headerMenuColumn.title} filter operator`}
                        value={filterDraft.operator}
                        onChange={event => setFilterDraft(current => ({ ...current, operator: event.target.value as FilterOperator }))}
                        style={{ width: "100%", marginTop: 3, padding: 5 }}
                    >
                        {headerMenuOperators.map(operator => <option key={operator} value={operator}>{operator}</option>)}
                    </select>
                    {headerMenuColumn.type === "boolean"
                        ? <select
                            aria-label={`${headerMenuColumn.title} filter value`}
                            value={filterDraft.value}
                            onChange={event => setFilterDraft(current => ({ ...current, value: event.target.value }))}
                            style={{ width: "100%", marginTop: 5, padding: 5 }}
                        >
                            <option value="">Any value</option>
                            <option value="true">True</option>
                            <option value="false">False</option>
                        </select>
                        : <input
                            aria-label={`${headerMenuColumn.title} filter value`}
                            type={headerMenuColumn.type === "number" ? "number" : "search"}
                            value={filterDraft.value}
                            placeholder={headerMenuColumn.type === "number" ? "Number" : `Search ${headerMenuColumn.title}`}
                            onChange={event => setFilterDraft(current => ({ ...current, value: event.target.value }))}
                            style={{ width: "100%", boxSizing: "border-box", marginTop: 5, padding: 5 }}
                        />}
                    <div style={{ display: "flex", gap: 5, marginTop: 6 }}>
                        <button type="button" onClick={applyHeaderFilter}>Apply filter</button>
                        <button type="button" onClick={clearHeaderFilter}>Clear filter</button>
                    </div>
                </div>
            </div>}
            <div style={{ padding: "7px 14px", borderTop: "1px solid #dadce0", fontSize: 13, textAlign: "right", background: "#f8fafd" }}>
                {statusMessage === "" ? selectionStats : <span role="status" style={{ color: "#c62828" }}>{statusMessage}</span>}
            </div>
        </div>
    );
};
