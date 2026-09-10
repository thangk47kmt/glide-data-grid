/* eslint-disable sonarjs/no-duplicate-string */
import * as React from "react";
import { act, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, test, vi } from "vitest";
import {
    basicProps,
    Context,
    EventedDataEditor,
    prep,
    standardAfterEach,
    standardBeforeEach,
} from "./test-utils.js";

describe("data-grid search labels", () => {
    beforeEach(() => {
        standardBeforeEach();
    });

    afterEach(() => {
        standardAfterEach();
        vi.useRealTimers();
    });

    test("uses custom labels for the search controls and result count", () => {
        vi.useFakeTimers();
        render(
            <EventedDataEditor
                {...basicProps}
                showSearch={true}
                onSearchClose={vi.fn()}
                searchLabels={{
                    result: "kết quả",
                    results: "kết quả",
                    of: "trên",
                    previous: "Kết quả trước",
                    next: "Kết quả sau",
                    close: "Đóng tìm kiếm",
                    typeToSearch: "Nhập để tìm kiếm",
                }}
            />,
            { wrapper: Context }
        );
        prep(false);

        expect(screen.getByText("Nhập để tìm kiếm")).toBeTruthy();
        expect(screen.getByRole("button", { name: "Kết quả trước" })).toBeTruthy();
        expect(screen.getByRole("button", { name: "Kết quả sau" })).toBeTruthy();
        expect(screen.getByRole("button", { name: "Đóng tìm kiếm" })).toBeTruthy();

        const input = screen.getByTestId("search-input");
        fireEvent.change(input, { target: { value: "1, 2" } });
        act(() => {
            vi.advanceTimersByTime(1000);
            vi.runAllTimers();
        });
        expect(screen.getByTestId("search-result-area").textContent).toBe("111 kết quả");
        fireEvent.keyDown(input, { key: "Enter" });
        expect(screen.getByTestId("search-result-area").textContent).toBe("1 trên 111 kết quả");
    });

    test("uses the custom type-to-search label before a query", () => {
        vi.useFakeTimers();
        render(
            <EventedDataEditor
                {...basicProps}
                showSearch={true}
                onSearchClose={vi.fn()}
                searchLabels={{ typeToSearch: "Nhập để tìm kiếm" }}
            />,
            { wrapper: Context }
        );
        prep(false);

        expect(screen.getByText("Nhập để tìm kiếm")).toBeTruthy();
        act(() => {
            vi.runAllTimers();
        });
    });
});
