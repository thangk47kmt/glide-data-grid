# Spreadsheet localization

Spreadsheet UI labels are resolved by the framework-agnostic helper exported from `@glideapps/glide-data-grid-spreadsheet`.
The helper does not read global browser state, so a production host remains the source of truth for the current system language.

```ts
import { createSpreadsheetI18n } from "@glideapps/glide-data-grid-spreadsheet";

const i18n = createSpreadsheetI18n({
    locale: systemLocale, // for example "vi-VN" or "en-US"
    messages: tenantOverrides, // optional Partial<SpreadsheetMessages>
});

i18n.t("pageSize");
i18n.t("pastedCells", { count: 12 });
i18n.filterOperator("gte");
```

The core grid's built-in `Ctrl/Cmd+F` search accepts the same host-driven approach:

```tsx
<DataEditor
    {...gridProps}
    searchLabels={{
        result: i18n.t("searchResult"),
        results: i18n.t("searchResults"),
        over1000: i18n.t("searchOver1000"),
        of: i18n.t("searchOf"),
        previous: i18n.t("previousResult"),
        next: i18n.t("nextResult"),
        close: i18n.t("closeSearch"),
        typeToSearch: i18n.t("typeToSearch"),
    }}
/>
```

Supported built-in languages are English and Vietnamese. Unknown locales fall back to English. Partial overrides fall back to the selected built-in language and never expose a raw message key.

The 500k Storybook story includes a language selector and updates toolbar, pagination, find/replace, formula controls, selection aggregates, status labels, column menus, filter operators and accessibility labels without resetting grid state.

## Stable data contract

- Column `id`, stored values, filter operators and formula syntax are never translated.
- A column caption may be localized by the host, but formula aliases must remain stable. Do not mutate the model column title after constructing the formula adapter.
- Dropdown/domain values should retain a stable stored `value`; localize only their displayed `label` in the host data source.
- Function names and spreadsheet error codes are invariant. Descriptive validation text can be localized separately.

This keeps saved filters, persisted formulas and database mappings independent from the current display language.
