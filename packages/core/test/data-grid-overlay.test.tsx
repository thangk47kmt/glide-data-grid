import * as React from "react";
import { render, cleanup, screen, fireEvent } from "@testing-library/react";
import BubblesOverlayEditor from "../src/internal/data-grid-overlay-editor/private/bubbles-overlay-editor.js";
import DrilldownOverlayEditor from "../src/internal/data-grid-overlay-editor/private/drilldown-overlay-editor.js";
import { GridCellKind, ImageOverlayEditor } from "../src/index.js";
import { MarkdownOverlayEditor } from "../src/internal/data-grid-overlay-editor/private/markdown-overlay-editor.js";
import NumberOverlayEditor from "../src/internal/data-grid-overlay-editor/private/number-overlay-editor.js";
import UriOverlayEditor from "../src/internal/data-grid-overlay-editor/private/uri-overlay-editor.js";
import { vi, describe, test, afterEach, expect } from "vitest";

describe("data-grid-overlay", () => {
    afterEach(() => {
        cleanup();
    });

    test("Smoke test bubbles", async () => {
        render(<BubblesOverlayEditor bubbles={["A", "B"]} />);
    });

    test("Drilldown overlay exposes its values to assistive technology", async () => {
        render(<DrilldownOverlayEditor drilldowns={[{ text: "A" }, { text: "B" }]} />);

        expect(screen.getByRole("list", { name: "Drilldown values" })).toBeTruthy();
        expect(screen.getAllByRole("listitem")).toHaveLength(2);
        expect(screen.getByRole("listitem", { name: "A" })).toBeTruthy();
    });

    test("Drilldown overlay edits values and commits with Enter", () => {
        const onChange = vi.fn();
        const onFinishedEditing = vi.fn();
        const value = {
            kind: GridCellKind.Drilldown,
            allowOverlay: true,
            data: [{ text: "A" }, { text: "B" }],
        } as const;

        render(
            <DrilldownOverlayEditor
                drilldowns={value.data}
                value={value}
                onChange={onChange}
                onFinishedEditing={onFinishedEditing}
            />
        );

        const input = screen.getByRole("textbox", { name: "Drilldown value 1" });
        fireEvent.change(input, { target: { value: "Edited" } });
        expect(onChange).toHaveBeenLastCalledWith(expect.objectContaining({ data: [{ text: "Edited" }, { text: "B" }] }));

        fireEvent.keyDown(input, { key: "Enter" });
        expect(onFinishedEditing).toHaveBeenCalledWith(
            expect.objectContaining({ data: [{ text: "Edited" }, { text: "B" }] }),
            [0, 1]
        );
    });

    test("Smoke test image overlay", async () => {
        const spy = vi.fn();

        render(<ImageOverlayEditor canWrite={false} onCancel={spy} onChange={spy} urls={["https://www.google.com"]} />);
    });

    test("Smoke test markdown overlay", async () => {
        const spy = vi.fn();

        render(
            <MarkdownOverlayEditor
                forceEditMode={false}
                value={{
                    kind: GridCellKind.Markdown,
                    allowOverlay: true,
                    data: "# Header",
                }}
                onChange={spy}
                onFinish={spy}
                targetRect={{ x: 0, y: 0, width: 200, height: 32 }}
            />
        );
    });

    test("Smoke test number overlay", async () => {
        const spy = vi.fn();

        render(<NumberOverlayEditor highlight={false} onChange={spy} value={35} />);
    });

    test("Smoke test uri overlay editor", async () => {
        const spy = vi.fn();

        render(
            <UriOverlayEditor
                forceEditMode={false}
                onChange={spy}
                readonly={false}
                uri={"https://google.com"}
                preview={"https://google.com"}
            />
        );
    });
});
