# Spreadsheet Phase 0 acceptance and regression contract

Phase 0 is complete. This checklist records the behavior that must remain intact while later work is added. Phase 1 and Phase 2 are intentionally deferred until explicitly requested.

## Automated gate

Run before accepting spreadsheet changes:

```sh
npm run verify:spreadsheet:phase0
```

The gate covers:

- raw inputs, computed values and display values;
- A1, mixed/absolute and caption/structured references;
- parser, calculation, dependency invalidation and cycle/error values;
- every promised Phase 0 function: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `AND`, `OR`, `NOT`, `ABS`, `ROUND`;
- computed-value search, typed filters and stable multi-sort;
- relative/absolute formula translation;
- selection aggregates, paged data, formula validation and history primitives.

## Required UI acceptance

Open the 500k story with `run-spreadsheet-test.bat` or:

```text
http://localhost:9009/iframe.html?id=extra-packages-spreadsheet-large-dataset--large-dataset-story&viewMode=story
```

Verify these behaviors after changes that touch the grid or Storybook integration:

- The grid reports 500,000 rows and can move between pages without materializing the whole dataset.
- The top letter tier shows A/B/C/... while the normal header shows the user-facing caption.
- Column widths can be dragged and numeric cells/results remain right-aligned.
- Clicking a column header opens search, filter and sort; Shift-click preserves multi-sort.
- Selecting multiple cells shows Count, Count Numbers, Sum, Average, Min and Max.
- A formula can be entered directly in a cell or in the formula bar; Enter and the check button commit, Escape and the cross button cancel.
- While editing a formula, clicking a cell/range/header inserts a reference. Both `B1` syntax and caption syntax remain accepted.
- Database column ids remain stable while captions are used for display and structured formula authoring.
- A direct reference between incompatible source/destination data types is rejected with a visible detailed alert; the old value/history remain unchanged and the draft stays editable.
- Direct references of matching types work for text, number, boolean and supported semantic/custom cell types.
- Copying one formula to a multi-cell selection updates every destination and rebases relative A1 references. Row, column and rectangular matrix paste remain one undo transaction.
- A failed formula in a pasted matrix rejects the transaction rather than partially writing it.
- Dropdown cells use one click for selection/copy and double-click for the searchable editor.
- Search, per-column search, filter, stable sort, Find/Replace, AutoSum, Undo and Redo remain operable.
- Image, URI, Markdown, Bubble, Drilldown, Protected, Loading, Row ID, Dropdown, Sparkline, Stars, Tags, Date, Links, Range, Multi-select and User profile examples still render.

## Change policy

- Do not remove an item above without an explicit product decision.
- A bug fix for one editing path must be checked against formula-bar, in-cell, copy/paste and undo paths.
- New Phase 1 or Phase 2 behavior must not be folded into Phase 0 implicitly; update the roadmap first when that work is requested.
- Story-only behavior must eventually receive a browser-level regression test. Until a repository browser test harness is added, the UI acceptance list above is mandatory for integration changes.
