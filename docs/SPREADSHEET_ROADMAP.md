# Glide Data Grid Spreadsheet Roadmap

## Product goal

Add an optional spreadsheet layer with search, filter, sort, formulas, aggregates and Excel-like utilities without moving data ownership or calculation into `packages/core`. The canvas grid remains a small rendering and interaction engine. Spreadsheet capabilities live in `packages/spreadsheet` and applications can opt in.

## Principles

1. Keep `core` data-agnostic and avoid adding formula dependencies to its bundle.
2. Recalculate only dirty dependent cells; never scan the workbook on every render.
3. Keep raw input, computed value, display value and formatting separate.
4. Make UI operations transactional so paste, fill and undo are one history entry.
5. Run expensive calculation, filtering and sorting in a Web Worker once data size crosses a configurable threshold.
6. Ship features in compatibility tiers instead of promising complete Excel parity.

## Feature backlog

### Data exploration

- Global and per-column search with debounce, match count and next/previous navigation.
- Typed filters: text, number, boolean, date, empty, multi-value and custom predicates.
- Multi-column stable sort with null/error ordering and locale-aware text comparison.
- Filter/sort state serialization and server-side adapter hooks.
- Column menu UI, filter chips, clear-all and saved views.
- Selection status aggregates: count, count numbers, sum, average, min and max.

### Formula and calculation

- A1 references, absolute/mixed references and rectangular ranges.
- Structured references by column name: `[@Quantity]` for current row and `[Quantity]` for a column.
- Arithmetic, comparison, concatenation and parentheses.
- Function registry with initial math, aggregate, logical, text and date functions.
- Dependency graph, incremental recalculation, cycle detection and spreadsheet error values.
- Formula bar, reference highlighting, syntax diagnostics and function autocomplete.
- Relative-reference translation for copy, paste and fill.
- AutoSum and quick aggregate insertion.
- Named ranges, cross-sheet references and configurable locale separators in later phases.
- Async/custom functions behind explicit registration; never enable arbitrary JavaScript evaluation.

### Spreadsheet editing

- Smart fill for numbers, dates, weekdays and formulas.
- Transactional undo/redo for values, formulas, paste, fill, row/column operations and formatting.
- Insert/delete/move rows and columns while updating references.
- Find/replace in values or formulas.
- Data validation rules and dropdown sources.
- Cell protection and locked formula cells.

### Formatting and presentation

- Number, currency, percent and date formats independent from stored values.
- Per-cell font, fill, border and alignment model.
- Conditional formatting with range indexes.
- Freeze panes, merged regions and hidden rows/columns.
- Optional charts, comments and pivot features as separate packages.

### Interoperability

- Versioned JSON workbook format first.
- CSV import/export for values.
- XLSX import/export through an optional adapter package so the runtime bundle stays small.
- Persistence callbacks, optimistic changes and collaboration event protocol.

## Delivery plan

## Current implementation status

The optional spreadsheet package now includes the original runnable foundation plus the first production-MVP building blocks:

- Incremental single-sheet model with A1, absolute/mixed and structured references, cycle/error handling, and a synchronous custom-function registry.
- Aggregate, logical, math, text, date and exact lookup functions, including `INDEX`, `MATCH` and `XLOOKUP`.
- Typed global/per-column search, filter and ordered stable multi-sort over computed/display values, versioned column-view state, an optional compact search index, selection aggregates and locale-aware display formatting.
- Formula diagnostics/autocomplete, raw/computed Find, bounded Replace plans and safe formula-aware replacements; the Storybook panel supports navigation and undoable replacement without silently applying truncated results.
- Transactional edit/fill history, a unified cell/snapshot/structural command stack, a replaceable `SpreadsheetSession`, Undo/Redo, smart number/date/formula fill, and data-validation primitives.
- Immutable row/column structural transforms with A1 rewrites and deterministic `#REF!`, plus conditional-format rules with pre-indexed duplicate detection.
- Versioned sparse JSON snapshots, bounded CSV parsing/stringifying, and model-level raw/computed export helpers.
- A Storybook integration harness with formula bar suggestions/diagnostics, AutoSum, validation, conditional-format presets, Undo/Redo, global controls plus click/Shift-click column-header search/filter/sort, and a compact CSV/snapshot debug panel.

These capabilities remain outside `packages/core`, so applications pay for them only when they opt into `packages/spreadsheet`.

### Phase 0 — runnable foundation (included in this branch)

**Status: complete.** The functional checklist and exit criteria below are covered by the spreadsheet test suite and the Storybook acceptance checklist in [`SPREADSHEET_PHASE0_ACCEPTANCE.md`](./SPREADSHEET_PHASE0_ACCEPTANCE.md).

MVP outcome: developers can start Storybook, edit a small sheet, search/filter/sort it, enter formulas and see recalculation.

- New optional `packages/spreadsheet` workspace.
- Raw cell input and computed-value model.
- A1 and structured column references.
- Dependency-aware cache invalidation and cycle errors.
- `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `AND`, `OR`, `NOT`, `ABS`, `ROUND`.
- Search, typed filter primitives and stable multi-sort.
- Formula translation helper for relative and absolute A1 references.
- Storybook UI with formula bar, AutoSum, search, total filter, sort and selection aggregates.
- Unit tests for parser, calculation, dependency invalidation and data view.

Exit criteria:

- Formula edits update dependent values without rebuilding the model.
- Search/filter/sort use computed formula values.
- Syntax, division and circular errors render as spreadsheet errors.
- `npm run test-spreadsheet` passes and the Storybook story is interactive.

Verification command:

```sh
npm run verify:spreadsheet:phase0
```

### Phase 1 — production MVP

**Status: deferred.** Do not start Phase 1 work until it is explicitly requested.

Target: a dependable single-sheet component for internal business applications.

- Introduce an explicit transaction/change-set API and production-grade undo/redo.
- Connect `onFillPattern` to formula translation and numeric/date series inference.
- Add column-menu search/filter/sort UI and saved view state.
- Add text/date functions and configurable function registry.
- Add formula diagnostics with reference highlighting and autocomplete.
- Add number/date formats and data-validation rules.
- Add JSON/CSV persistence and import/export.
- Benchmark 100k rows, 1M populated cells and dependency chains; define performance budgets.
- Move calculation and large sort/filter operations to a worker adapter.

Exit criteria:

- Common edit, paste and fill operations are atomic and undoable.
- A 100k-row sheet stays interactive under the agreed hardware/browser profile.
- Formula and view engine have deterministic tests and no `eval`/`Function` execution.

### Phase 2 — workbook capability

**Status: deferred.** Do not start Phase 2 work until it is explicitly requested.

Target: Excel-like multi-sheet workflows without advanced analytics.

- Workbook and sheet model, sheet tabs and cross-sheet references.
- Named ranges and formula-safe row/column insert/delete.
- Conditional formatting and richer cell styles.
- Filtered-subtotal semantics and hidden row/column handling.
- Optional XLSX adapter with compatibility fixtures.
- Collaborative operation schema and conflict policy.

Exit criteria:

- Workbook round-trips through the versioned JSON format.
- Structural changes preserve valid references or produce `#REF!` deterministically.
- Cross-sheet recalculation remains incremental.

### Phase 3 — advanced optional packages

Target: capabilities that should not increase the default spreadsheet bundle.

- Dynamic arrays and spill ranges.
- Lookup expansion (`XLOOKUP`, `INDEX`, `MATCH`) and broader function compatibility.
- Pivot engine, charts, comments and audit tools.
- XLSX fidelity improvements.
- Async functions and remote calculation adapters.

Each feature ships as an opt-in module with its own bundle and performance budget.

## Local development

Use Node from `.nvmrc`, then install workspace dependencies:

```sh
npm install
```

Run the interactive spreadsheet story:

```sh
npm run debug:spreadsheet
```

This debug command compiles the core source directly through Storybook/Vite, so it also works on Windows machines without Bash. The existing release build scripts still require Bash 4 or newer.

Open Storybook at `http://localhost:9009` and select **Extra Packages / Spreadsheet MVP / Interactive Spreadsheet**.

Run only spreadsheet tests:

```sh
npm run test-spreadsheet
```

Run the repeatable local performance harness (defaults to 100,000 rows, 10 columns/1,000,000 populated cells and a 500-cell dependency chain):

```sh
npm run benchmark:spreadsheet
```

Override the workload with `SPREADSHEET_BENCH_ROWS`, `SPREADSHEET_BENCH_COLUMNS` and `SPREADSHEET_BENCH_CHAIN`. Results are machine-specific and are intended for regression comparison, not as universal latency guarantees.

Latest local baseline (2026-08-31, 100,000 rows × 10 columns): model construction 1.03 s; one uncached search 445 ms; five uncached searches 2.15 s; compact index construction 447 ms; one indexed search 31 ms; five indexed searches 76 ms. The index used about 33 MB additional heap in this run and a 4 MB flat boundary buffer. It is therefore opt-in and is most useful for repeated queries; one-off queries should keep the zero-index path.

Run all existing package tests plus spreadsheet tests:

```sh
npm run test-all
```

Build all workspaces:

```sh
npm run build
```

## Current MVP limitations

- One sheet only; no named ranges, cross-sheet references or collaborative operation protocol. Structural transforms exist as immutable payload operations but are not yet wired to Storybook row/column controls or in-place model resizing.
- Formula syntax uses English function names and comma separators. Diagnostics and autocomplete exist; reference highlighting and locale-aware separators do not.
- Smart fill is connected to the UI, but custom fill is deliberately disabled while rows are filtered or sorted because visible-row coordinates are non-contiguous.
- Whole-column structured references are tracked efficiently for invalidation but evaluation still visits all rows.
- Calculation, filter and sort still run on the main thread. The repeatable 100k-row/1M-cell harness and optional search index exist; an agreed cross-device budget and worker adapter remain future work.
- JSON and CSV are available; XLSX remains an optional future adapter.
- Formatting and validation engines plus conditional-format presets are integrated in the demo. The 500k story includes a searchable dropdown editor; rich workbook-level per-cell styles remain future work.
- Find/Replace is available in the Storybook harness. Replace All scans source rows (including rows hidden by view filters) and is deliberately blocked when the bounded result is `truncated`.
- `SpreadsheetSession` proves cell and structural commands can share one ordered history, but the Storybook row/column structural buttons are still pending migration from its existing cell-only history wiring.
- Column-header view state is available as a versioned immutable API. The Storybook header popover is a reference integration; applications remain responsible for saving/restoring that state with their own view persistence.
- The Storybook demo is an integration harness, not yet a public React spreadsheet component.

For the evidence-backed comparison with Excel/Google Sheets, see [EXCEL_COMPATIBILITY.md](./EXCEL_COMPATIBILITY.md).
