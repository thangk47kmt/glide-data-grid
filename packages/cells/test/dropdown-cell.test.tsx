import * as React from "react";

import { fireEvent, render, cleanup } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { GridCellKind } from "@glideapps/glide-data-grid";
import renderer, {
    filterDropdownOptions,
    isDropdownValueAllowed,
    normalizeDropdownOptions,
    type DropdownCell,
    validateDropdownPaste,
} from "../src/cells/dropdown-cell.js";

const allowedValues = ["Alpha", { value: "beta", label: "Beta option" }, "Gamma"] as const;
const options = normalizeDropdownOptions(allowedValues);

function makeCell(value: string | undefined | null = "Alpha"): DropdownCell {
    return {
        kind: GridCellKind.Custom,
        allowOverlay: true,
        copyData: value ?? "",
        data: { kind: "dropdown-cell", value, allowedValues },
    };
}

describe("searchable dropdown helpers", () => {
    afterEach(cleanup);

    it("normalizes options and filters value or label case-insensitively", () => {
        expect(options).toEqual([
            { value: "Alpha", label: "Alpha" },
            { value: "beta", label: "Beta option" },
            { value: "Gamma", label: "Gamma" },
        ]);
        expect(filterDropdownOptions(options, "BETA")).toEqual([{ value: "beta", label: "Beta option" }]);
        expect(filterDropdownOptions(options, "amm")).toEqual([{ value: "Gamma", label: "Gamma" }]);
        expect(filterDropdownOptions(options, "")).toHaveLength(3);
    });

    it("validates controlled values and preserves invalid pasted values", () => {
        expect(isDropdownValueAllowed("Alpha", allowedValues)).toBe(true);
        expect(isDropdownValueAllowed("alpha", allowedValues)).toBe(false);
        expect(isDropdownValueAllowed("missing", allowedValues)).toBe(false);
        expect(validateDropdownPaste("beta", "Alpha", allowedValues)).toBe("beta");
        expect(validateDropdownPaste("BETA", "Alpha", allowedValues)).toBe("Alpha");
        expect(validateDropdownPaste("missing", null, allowedValues)).toBeNull();
    });

    it("uses the same allow-list validation for grid paste", () => {
        const cell = makeCell();
        const pasted = renderer.onPaste?.("Gamma", cell.data);
        const rejected = renderer.onPaste?.("Unknown", cell.data);
        expect(pasted?.value).toBe("Gamma");
        expect(rejected?.value).toBe("Alpha");
    });

    it("keeps keyboard navigation inside the searchable editor", () => {
        const Editor = renderer.provideEditor?.({ ...makeCell(), location: [0, 0] }).editor;
        if (Editor === undefined) throw new Error("Dropdown editor is unavailable");
        const onFinishedEditing = vi.fn();
        let bubbled = false;
        const cell = makeCell();
        const view = render(
            <div onKeyDown={() => { bubbled = true; }}>
                <Editor value={cell} isHighlighted={false} onChange={vi.fn()} onFinishedEditing={onFinishedEditing} />
            </div>
        );
        const input = view.getByRole("combobox");
        fireEvent.keyDown(input, { key: "ArrowDown" });
        fireEvent.keyDown(input, { key: "Escape" });
        expect(bubbled).toBe(false);
        expect(onFinishedEditing).not.toHaveBeenCalled();
    });
});
