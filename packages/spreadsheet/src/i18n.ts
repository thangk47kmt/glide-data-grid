/**
 * Small, framework-agnostic translations contract for spreadsheet controls.
 *
 * The package intentionally does not read from `navigator` or keep global
 * locale state. Hosts can pass their system locale at the boundary and may
 * override only the labels they own.
 */

export type SpreadsheetLocale = "en" | "vi";

export type SpreadsheetFilterOperator =
    | "contains"
    | "equals"
    | "not-equals"
    | "gt"
    | "gte"
    | "lt"
    | "lte"
    | "empty"
    | "not-empty";

/** All labels used by the spreadsheet demo controls and column menus. */
export interface SpreadsheetMessages {
    readonly language: string;
    readonly english: string;
    readonly vietnamese: string;
    readonly clearQuery: string;
    readonly undo: string;
    readonly redo: string;
    readonly findReplace: string;
    readonly search: string;
    readonly searching: string;
    readonly searchPlaceholder: string;
    readonly clearSearch: string;
    readonly find: string;
    readonly replace: string;
    readonly replaceWith: string;
    readonly caseSensitive: string;
    readonly wholeCell: string;
    readonly regex: string;
    readonly previous: string;
    readonly next: string;
    readonly replaceCurrent: string;
    readonly replaceAll: string;
    readonly matches: string;
    readonly pageSize: string;
    readonly firstPage: string;
    readonly previousPage: string;
    readonly nextPage: string;
    readonly lastPage: string;
    readonly page: string;
    readonly pageNumber: string;
    readonly rowsColumns: string;
    readonly loadedPage: string;
    readonly loadingPage: string;
    readonly cache: string;
    readonly hits: string;
    readonly misses: string;
    readonly generated: string;
    readonly sparseEdits: string;
    readonly query: string;
    readonly formulas: string;
    readonly ready: string;
    readonly queryCleared: string;
    readonly matchingRows: string;
    readonly autoSum: string;
    readonly formulaBar: string;
    readonly confirm: string;
    readonly cancel: string;
    readonly sortAscending: string;
    readonly sortDescending: string;
    readonly clearSort: string;
    readonly searchThisColumn: string;
    readonly applyColumnSearch: string;
    readonly filter: string;
    readonly any: string;
    readonly trueValue: string;
    readonly falseValue: string;
    readonly applyFilter: string;
    readonly clearFilter: string;
    readonly shiftClickKeepsExistingSortKeys: string;
    readonly clickSortReplacesExistingKeys: string;
    readonly operatorContains: string;
    readonly operatorEquals: string;
    readonly operatorNotEquals: string;
    readonly operatorGreaterThan: string;
    readonly operatorGreaterThanOrEqual: string;
    readonly operatorLessThan: string;
    readonly operatorLessThanOrEqual: string;
    readonly operatorEmpty: string;
    readonly operatorNotEmpty: string;
    readonly rows: string;
    readonly cells: string;
    readonly selected: string;
    readonly sum: string;
    readonly average: string;
    readonly count: string;
    readonly countNumbers: string;
    readonly min: string;
    readonly max: string;
    readonly selectCellsToCalculate: string;
    readonly selectionStatsLimited: string;
    readonly formulaPlaceholder: string;
    readonly finding: string;
    readonly findMatches: string;
    readonly retainedMatches: string;
    readonly findMatchesRetained: string;
    readonly updatedCell: string;
    readonly pastedCells: string;
    readonly copiedCells: string;
    readonly replacedCells: string;
    readonly formulaInvalid: string;
    readonly matchHidden: string;
    readonly replaceAllLimit: string;
    readonly autoSumLimit: string;
    readonly nonContiguousReference: string;
    readonly singleColumnReference: string;
    readonly findError: string;
    readonly replaceError: string;
    readonly columnMenu: string;
    readonly columnSearch: string;
    readonly filterOperator: string;
    readonly filterValue: string;
    readonly findReplaceRegion: string;
    readonly findText: string;
    readonly replacementText: string;
    readonly inCellFormulaEditor: string;
    readonly inCellFormulaEditorControls: string;
    readonly searchRows: string;
    readonly searchResult: string;
    readonly searchResults: string;
    readonly searchOver1000: string;
    readonly searchOf: string;
    readonly previousResult: string;
    readonly nextResult: string;
    readonly closeSearch: string;
    readonly typeToSearch: string;
}

export type SpreadsheetMessageKey = keyof SpreadsheetMessages;
export type SpreadsheetMessageValues = Readonly<Record<string, string | number>>;

export interface SpreadsheetI18nOptions {
    /** Accepts a BCP-47 locale (`vi-VN`, `en-US`, etc.). Unsupported values use English. */
    readonly locale?: string | null;
    /** Partial overrides are merged over the selected locale's messages. */
    readonly messages?: Partial<SpreadsheetMessages> | null;
}

export interface SpreadsheetI18n {
    readonly locale: SpreadsheetLocale;
    readonly messages: Readonly<SpreadsheetMessages>;
    /** Resolves a label and interpolates `{name}` placeholders without throwing. */
    readonly t: (key: SpreadsheetMessageKey, values?: SpreadsheetMessageValues) => string;
    readonly filterOperator: (operator: SpreadsheetFilterOperator) => string;
}

const EN_MESSAGES: SpreadsheetMessages = {
    language: "Language",
    english: "English",
    vietnamese: "Vietnamese",
    clearQuery: "Clear query",
    undo: "Undo",
    redo: "Redo",
    findReplace: "Find / Replace",
    search: "Search",
    searching: "Searching…",
    searchPlaceholder: "Search all rows…",
    clearSearch: "Clear search",
    find: "Find",
    replace: "Replace",
    replaceWith: "Replace with",
    caseSensitive: "Case sensitive",
    wholeCell: "Whole cell",
    regex: "Regex",
    previous: "Previous",
    next: "Next",
    replaceCurrent: "Replace current",
    replaceAll: "Replace all",
    matches: "matches",
    pageSize: "Page size",
    firstPage: "First",
    previousPage: "Prev",
    nextPage: "Next",
    lastPage: "Last",
    page: "Page",
    pageNumber: "Page number",
    rowsColumns: "{rows} rows × {columns} columns",
    loadedPage: "Loaded page in {ms} ms",
    loadingPage: "Loading page…",
    cache: "cache",
    hits: "hits",
    misses: "misses",
    generated: "generated",
    sparseEdits: "sparse edits",
    query: "query",
    formulas: "formulas",
    ready: "Ready",
    queryCleared: "Query cleared",
    matchingRows: "{count} matching rows",
    autoSum: "AutoSum",
    formulaBar: "Formula bar",
    confirm: "Confirm",
    cancel: "Cancel",
    sortAscending: "Sort ascending",
    sortDescending: "Sort descending",
    clearSort: "Clear sort",
    searchThisColumn: "Search this column",
    applyColumnSearch: "Apply column search",
    filter: "Filter",
    any: "Any",
    trueValue: "True",
    falseValue: "False",
    applyFilter: "Apply filter",
    clearFilter: "Clear filter",
    shiftClickKeepsExistingSortKeys: "Shift-click keeps existing sort keys",
    clickSortReplacesExistingKeys: "Click sort replaces existing keys",
    operatorContains: "Contains",
    operatorEquals: "Equals",
    operatorNotEquals: "Not equal to",
    operatorGreaterThan: "Greater than",
    operatorGreaterThanOrEqual: "Greater than or equal to",
    operatorLessThan: "Less than",
    operatorLessThanOrEqual: "Less than or equal to",
    operatorEmpty: "Is empty",
    operatorNotEmpty: "Is not empty",
    rows: "rows",
    cells: "cells",
    selected: "Selected",
    sum: "Sum",
    average: "Average",
    count: "Count",
    countNumbers: "Count Numbers",
    min: "Min",
    max: "Max",
    selectCellsToCalculate: "Select cells to calculate",
    selectionStatsLimited: "Selection stats limited to {max} cells (selected {count})",
    formulaPlaceholder: "=B1*2 or =[@Quantity]*[@Unit price]",
    finding: "Finding…",
    findMatches: "{count} find matches",
    retainedMatches: "first {count} retained",
    findMatchesRetained: "{count} find matches (first {retained} retained)",
    updatedCell: "Updated {cell}",
    pastedCells: "Pasted {count} cells",
    copiedCells: "Copied {count} cells",
    replacedCells: "Replaced {count} cells",
    formulaInvalid: "Formula is invalid",
    matchHidden: "Match is hidden by the active query",
    replaceAllLimit: "Replace All blocked: more than the {count}-match safety cap",
    autoSumLimit: "AutoSum is limited to {count} source rows in the local MVP to protect the UI",
    nonContiguousReference: "Cannot insert a multi-row reference while rows are filtered/sorted unless source rows are contiguous",
    singleColumnReference: "Select one column at a time when inserting a structured formula reference",
    findError: "Find error: {message}",
    replaceError: "Replace error: {message}",
    columnMenu: "{column} column menu",
    columnSearch: "{column} search",
    filterOperator: "{column} filter operator",
    filterValue: "{column} filter value",
    findReplaceRegion: "Large dataset find and replace",
    findText: "Find text",
    replacementText: "Replacement text",
    inCellFormulaEditor: "In-cell formula editor",
    inCellFormulaEditorControls: "In-cell formula editor controls",
    searchRows: "Search 500k rows",
    searchResult: "result",
    searchResults: "results",
    searchOver1000: "over 1000",
    searchOf: "of",
    previousResult: "Previous Result",
    nextResult: "Next Result",
    closeSearch: "Close Search",
    typeToSearch: "Type to search",
};

const VI_MESSAGES: SpreadsheetMessages = {
    language: "Ngôn ngữ",
    english: "Tiếng Anh",
    vietnamese: "Tiếng Việt",
    clearQuery: "Xóa truy vấn",
    undo: "Hoàn tác",
    redo: "Làm lại",
    findReplace: "Tìm / Thay thế",
    search: "Tìm kiếm",
    searching: "Đang tìm…",
    searchPlaceholder: "Tìm trong tất cả dòng…",
    clearSearch: "Xóa tìm kiếm",
    find: "Tìm",
    replace: "Thay thế",
    replaceWith: "Thay thế bằng",
    caseSensitive: "Phân biệt hoa thường",
    wholeCell: "Toàn bộ ô",
    regex: "Biểu thức chính quy",
    previous: "Trước",
    next: "Tiếp",
    replaceCurrent: "Thay thế kết quả hiện tại",
    replaceAll: "Thay thế tất cả",
    matches: "kết quả",
    pageSize: "Kích thước trang",
    firstPage: "Đầu",
    previousPage: "Trước",
    nextPage: "Tiếp",
    lastPage: "Cuối",
    page: "Trang",
    pageNumber: "Số trang",
    rowsColumns: "{rows} dòng × {columns} cột",
    loadedPage: "Đã tải trang trong {ms} ms",
    loadingPage: "Đang tải trang…",
    cache: "bộ nhớ đệm",
    hits: "lần trúng",
    misses: "lần trượt",
    generated: "đã tạo",
    sparseEdits: "chỉnh sửa thưa",
    query: "truy vấn",
    formulas: "công thức",
    ready: "Sẵn sàng",
    queryCleared: "Đã xóa truy vấn",
    matchingRows: "{count} dòng phù hợp",
    autoSum: "Tự động tính tổng",
    formulaBar: "Thanh công thức",
    confirm: "Xác nhận",
    cancel: "Hủy",
    sortAscending: "Sắp xếp tăng dần",
    sortDescending: "Sắp xếp giảm dần",
    clearSort: "Xóa sắp xếp",
    searchThisColumn: "Tìm trong cột này",
    applyColumnSearch: "Áp dụng tìm kiếm cột",
    filter: "Bộ lọc",
    any: "Bất kỳ",
    trueValue: "Đúng",
    falseValue: "Sai",
    applyFilter: "Áp dụng bộ lọc",
    clearFilter: "Xóa bộ lọc",
    shiftClickKeepsExistingSortKeys: "Shift-click giữ các khóa sắp xếp hiện tại",
    clickSortReplacesExistingKeys: "Click sắp xếp thay thế các khóa hiện tại",
    operatorContains: "Chứa",
    operatorEquals: "Bằng",
    operatorNotEquals: "Khác",
    operatorGreaterThan: "Lớn hơn",
    operatorGreaterThanOrEqual: "Lớn hơn hoặc bằng",
    operatorLessThan: "Nhỏ hơn",
    operatorLessThanOrEqual: "Nhỏ hơn hoặc bằng",
    operatorEmpty: "Trống",
    operatorNotEmpty: "Không trống",
    rows: "dòng",
    cells: "ô",
    selected: "Đã chọn",
    sum: "Tổng",
    average: "Trung bình",
    count: "Số ô có dữ liệu",
    countNumbers: "Số ô kiểu số",
    min: "Nhỏ nhất",
    max: "Lớn nhất",
    selectCellsToCalculate: "Chọn các ô để tính toán",
    selectionStatsLimited: "Thống kê vùng chọn giới hạn ở {max} ô (đã chọn {count})",
    formulaPlaceholder: "=B1*2 hoặc =[@Số lượng]*[@Đơn giá]",
    finding: "Đang tìm…",
    findMatches: "{count} kết quả tìm kiếm",
    retainedMatches: "giữ lại {count} kết quả đầu tiên",
    findMatchesRetained: "{count} kết quả tìm kiếm (giữ lại {retained} kết quả đầu tiên)",
    updatedCell: "Đã cập nhật {cell}",
    pastedCells: "Đã dán {count} ô",
    copiedCells: "Đã sao chép {count} ô",
    replacedCells: "Đã thay thế {count} ô",
    formulaInvalid: "Công thức không hợp lệ",
    matchHidden: "Kết quả đang bị ẩn bởi truy vấn hiện tại",
    replaceAllLimit: "Không thể thay thế tất cả: vượt giới hạn an toàn {count} kết quả",
    autoSumLimit: "Tự động tính tổng giới hạn ở {count} dòng nguồn trong bản local để bảo vệ hiệu năng UI",
    nonContiguousReference: "Không thể chèn tham chiếu nhiều dòng khi đang lọc/sắp xếp nếu các dòng nguồn không liền nhau",
    singleColumnReference: "Chỉ chọn một cột khi chèn tham chiếu công thức có cấu trúc",
    findError: "Lỗi tìm kiếm: {message}",
    replaceError: "Lỗi thay thế: {message}",
    columnMenu: "Menu cột {column}",
    columnSearch: "Tìm trong cột {column}",
    filterOperator: "Toán tử lọc cột {column}",
    filterValue: "Giá trị lọc cột {column}",
    findReplaceRegion: "Tìm và thay thế trên dữ liệu lớn",
    findText: "Nội dung cần tìm",
    replacementText: "Nội dung thay thế",
    inCellFormulaEditor: "Trình sửa công thức trong ô",
    inCellFormulaEditorControls: "Điều khiển trình sửa công thức trong ô",
    searchRows: "Tìm trong 500 nghìn dòng",
    searchResult: "kết quả",
    searchResults: "kết quả",
    searchOver1000: "hơn 1.000",
    searchOf: "trên",
    previousResult: "Kết quả trước",
    nextResult: "Kết quả tiếp theo",
    closeSearch: "Đóng tìm kiếm",
    typeToSearch: "Nhập để tìm kiếm",
};

export const spreadsheetMessages: Readonly<Record<SpreadsheetLocale, SpreadsheetMessages>> = Object.freeze({
    en: Object.freeze(EN_MESSAGES),
    vi: Object.freeze(VI_MESSAGES),
});

/** Converts a system/browser locale into one of the supported base locales. */
export function normalizeSpreadsheetLocale(locale?: string | null): SpreadsheetLocale {
    if (typeof locale !== "string") return "en";
    const language = locale.trim().toLowerCase().replace(/_/g, "-").split("-", 1)[0];
    return language === "vi" ? "vi" : "en";
}

function safeOverrides(overrides: Partial<SpreadsheetMessages> | null | undefined): Partial<SpreadsheetMessages> {
    if (overrides === null || typeof overrides !== "object") return {};
    const result: Record<string, string> = {};
    for (const key of Object.keys(EN_MESSAGES) as SpreadsheetMessageKey[]) {
        const value = overrides[key];
        if (typeof value === "string") result[key] = value;
    }
    return result;
}

function interpolate(template: string, values: SpreadsheetMessageValues | undefined, locale: SpreadsheetLocale): string {
    if (values === undefined) return template;
    return template.replace(/\{([A-Za-z0-9_]+)\}/g, (placeholder, key: string) => {
        const value = values[key];
        if (value === undefined) return placeholder;
        return typeof value === "number" ? value.toLocaleString(locale === "vi" ? "vi-VN" : "en-US") : String(value);
    });
}

const OPERATOR_KEYS: Readonly<Record<SpreadsheetFilterOperator, SpreadsheetMessageKey>> = Object.freeze({
    contains: "operatorContains",
    equals: "operatorEquals",
    "not-equals": "operatorNotEquals",
    gt: "operatorGreaterThan",
    gte: "operatorGreaterThanOrEqual",
    lt: "operatorLessThan",
    lte: "operatorLessThanOrEqual",
    empty: "operatorEmpty",
    "not-empty": "operatorNotEmpty",
});

/** Creates an isolated translator; it is safe to create one per rendered host/story. */
export function createSpreadsheetI18n(
    localeOrOptions: string | null | SpreadsheetI18nOptions = "en",
    messages?: Partial<SpreadsheetMessages> | null,
): SpreadsheetI18n {
    const options: SpreadsheetI18nOptions = typeof localeOrOptions === "string" || localeOrOptions === null
        ? { locale: localeOrOptions, messages }
        : localeOrOptions;
    const locale = normalizeSpreadsheetLocale(options.locale);
    const merged = Object.freeze({ ...spreadsheetMessages.en, ...spreadsheetMessages[locale], ...safeOverrides(options.messages) });
    const t = (key: SpreadsheetMessageKey, values?: SpreadsheetMessageValues): string => interpolate(merged[key], values, locale);
    return Object.freeze({
        locale,
        messages: merged,
        t,
        filterOperator: (operator: SpreadsheetFilterOperator): string => t(OPERATOR_KEYS[operator] ?? "operatorEquals"),
    });
}
