import type { DrilldownCell, DrilldownCellData } from "../../data-grid/data-grid-types.js";
import * as React from "react";
import { styled } from "@linaria/react";

const DrilldownOverlayEditorStyle = styled.div`
    display: flex;
    flex-wrap: wrap;

    .doe-bubble {
        display: flex;
        justify-content: center;
        align-items: center;

        padding: 0 8px;
        height: 24px;

        background-color: var(--gdg-bg-cell);
        color: var(--gdg-text-dark);
        margin: 2px;

        border-radius: var(--gdg-rounding-radius, 6px);

        box-shadow:
            0 0 1px rgba(62, 65, 86, 0.4),
            0 1px 3px rgba(62, 65, 86, 0.4);

        img {
            height: 16px;
            object-fit: contain;

            margin-right: 4px;
        }
    }

    textarea {
        position: absolute;
        top: 0px;
        left: 0px;
        width: 0px;
        height: 0px;

        opacity: 0;
    }

    .doe-editor-row {
        display: flex;
        align-items: center;
        gap: 4px;
        margin: 2px;

        input {
            box-sizing: border-box;
            min-width: 96px;
            height: 24px;
            padding: 0 8px;
            border: 1px solid var(--gdg-border-color);
            border-radius: var(--gdg-rounding-radius, 6px);
            background: var(--gdg-bg-cell);
            color: var(--gdg-text-dark);
            font: inherit;
        }
    }
`;

interface Props {
    readonly drilldowns: readonly DrilldownCellData[];
    /** Editing props are supplied by the data-grid overlay. They remain optional
     * so the renderer can still be used as a read-only accessibility preview. */
    readonly value?: DrilldownCell;
    readonly onChange?: (value: DrilldownCell) => void;
    readonly onFinishedEditing?: (
        value?: DrilldownCell,
        movement?: readonly [-1 | 0 | 1, -1 | 0 | 1]
    ) => void;
}

const DrilldownOverlayEditor: React.FunctionComponent<Props> = p => {
    const { drilldowns, value, onChange, onFinishedEditing } = p;
    const editable = value !== undefined && onChange !== undefined && value.readonly !== true;
    const [drafts, setDrafts] = React.useState(() => drilldowns.map(d => d.text));
    const latestValue = React.useRef(value);
    const latestDrafts = React.useRef(drafts);
    latestValue.current = value;
    latestDrafts.current = drafts;

    React.useEffect(() => {
        const nextDrafts = drilldowns.map(d => d.text);
        setDrafts(nextDrafts);
        latestDrafts.current = nextDrafts;
    }, [drilldowns]);

    const updateDraft = React.useCallback((index: number, text: string) => {
        const nextDrafts = latestDrafts.current.map((draft, draftIndex) => draftIndex === index ? text : draft);
        latestDrafts.current = nextDrafts;
        setDrafts(nextDrafts);
        const current = latestValue.current;
        if (current !== undefined && onChange !== undefined) {
            onChange({
                ...current,
                data: current.data.map((entry, entryIndex) => entryIndex === index ? { ...entry, text } : entry),
            });
        }
    }, [onChange]);

    const finishEditing = React.useCallback((movement: readonly [-1 | 0 | 1, -1 | 0 | 1]) => {
        const current = latestValue.current;
        if (current !== undefined && onFinishedEditing !== undefined) {
            onFinishedEditing({
                ...current,
                data: current.data.map((entry, index) => ({ ...entry, text: latestDrafts.current[index] ?? entry.text })),
            }, movement);
        }
    }, [onFinishedEditing]);

    return (
        <DrilldownOverlayEditorStyle role="list" aria-label="Drilldown values">
            {drilldowns.map((d, i) => editable ? (
                <div key={i} className="doe-editor-row" role="listitem">
                    {d.img !== undefined && <img src={d.img} alt="" aria-hidden="true" />}
                    <input
                        aria-label={`Drilldown value ${i + 1}`}
                        value={drafts[i] ?? ""}
                        onChange={event => updateDraft(i, event.target.value)}
                        onKeyDown={event => {
                            if (event.key === "Enter") {
                                event.preventDefault();
                                finishEditing([0, 1]);
                            } else if (event.key === "Tab") {
                                event.preventDefault();
                                finishEditing([event.shiftKey ? -1 : 1, 0]);
                            }
                        }}
                    />
                </div>
            ) : (
                <div key={i} className="doe-bubble" role="listitem" aria-label={d.text}>
                    {d.img !== undefined && <img src={d.img} alt="" aria-hidden="true" />}
                    <div>{d.text}</div>
                </div>
            ))}
        </DrilldownOverlayEditorStyle>
    );
};
export default DrilldownOverlayEditor;
