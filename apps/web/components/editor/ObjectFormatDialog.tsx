import { useState } from "react";
import type {
  DrawingInline,
  EditCommand,
  Table,
  TableCell,
} from "@local/docx-ir-engine/core";
import { Dialog } from "./Controls";
export function ObjectFormatDialog({
  drawing,
  onClose,
  onApply,
  initialTab = "Size",
}: {
  drawing: DrawingInline;
  onClose: () => void;
  onApply: (d: DrawingInline) => Promise<boolean>;
  initialTab?: string;
}) {
  const [tab, setTab] = useState(initialTab),
    [draft, setDraft] = useState(() => structuredClone(drawing)),
    [lock, setLock] = useState(true),
    [error, setError] = useState("");
  const number = (
    label: string,
    value: number,
    change: (n: number) => void,
    min = 0,
    max = 15840,
  ) => (
    <label className="field">
      {label}
      <input
        type="number"
        min={min}
        max={max}
        step="0.1"
        value={Number(value.toFixed(2))}
        onChange={(e) => change(e.target.valueAsNumber || 0)}
      />
    </label>
  );
  const image = (patch: Partial<NonNullable<DrawingInline["image"]>>) =>
    setDraft((d) => ({ ...d, image: { ...d.image!, ...patch } }));
  return (
    <Dialog
      title={drawing.image ? "Format Picture" : "Size and Position"}
      onClose={onClose}
    >
      <form
        onSubmit={async (e) => {
          e.preventDefault();
          if (await onApply(draft)) onClose();
          else setError("The change could not be applied.");
        }}
      >
        <div className="format-tabs">
          {[
            "Size",
            "Text Wrapping",
            "Position",
            ...(draft.image ? ["Picture", "Crop"] : []),
          ].map((name) => (
            <button
              type="button"
              key={name}
              aria-pressed={tab === name}
              onClick={() => setTab(name)}
            >
              {name}
            </button>
          ))}
        </div>
        {tab === "Size" && (
          <>
            <div className="field-row">
              {number(
                "Width (pt)",
                draft.widthPt,
                (widthPt) =>
                  setDraft((d) => ({
                    ...d,
                    widthPt,
                    heightPt: lock
                      ? (d.heightPt * widthPt) / d.widthPt
                      : d.heightPt,
                  })),
                1,
              )}
              {number(
                "Height (pt)",
                draft.heightPt,
                (heightPt) =>
                  setDraft((d) => ({
                    ...d,
                    heightPt,
                    widthPt: lock
                      ? (d.widthPt * heightPt) / d.heightPt
                      : d.widthPt,
                  })),
                1,
              )}
            </div>
            <label className="check-field">
              <input
                type="checkbox"
                checked={lock}
                onChange={(e) => setLock(e.target.checked)}
              />
              Lock aspect ratio
            </label>
            {number(
              "Rotation (degrees)",
              draft.shape?.rotation ?? 0,
              (rotation) =>
                setDraft((d) => ({
                  ...d,
                  shape: {
                    ...d.shape,
                    kind: d.shape?.kind ?? "rect",
                    rotation,
                  },
                })),
              -36000,
              36000,
            )}
            <label className="field">
              Alt text
              <input
                value={draft.alt}
                onChange={(e) =>
                  setDraft((d) => ({ ...d, alt: e.target.value }))
                }
              />
            </label>
          </>
        )}
        {(tab === "Position" || tab === "Text Wrapping") && (
          <>
            <label className="field">
              Wrapping style
              <select
                value={
                  !draft.position
                    ? "inline"
                    : draft.position.behindText
                      ? "behind"
                      : draft.position.wrap
                }
                onChange={(e) => {
                  const wrap = e.target.value;
                  setDraft((d) => ({
                    ...d,
                    position:
                      wrap === "inline"
                        ? undefined
                        : {
                            ...d.position,
                            horizontal: d.position?.horizontal ?? {
                              relativeTo: "column",
                              offsetPt: 0,
                            },
                            vertical: d.position?.vertical ?? {
                              relativeTo: "paragraph",
                              offsetPt: 0,
                            },
                            behindText: wrap === "behind",
                            zIndex: d.position?.zIndex ?? 2,
                            wrap: wrap === "behind" ? "wrapNone" : wrap,
                          },
                  }));
                }}
              >
                {[
                  ["inline", "In line with text"],
                  ["wrapSquare", "Square"],
                  ["wrapTight", "Tight"],
                  ["wrapThrough", "Through"],
                  ["wrapTopAndBottom", "Top and bottom"],
                  ["behind", "Behind text"],
                  ["wrapNone", "In front of text"],
                ].map(([value, label]) => (
                  <option key={value} value={value}>
                    {label}
                  </option>
                ))}
              </select>
            </label>
            {draft.position && (
              <>
                {(["horizontal", "vertical"] as const).map((axis) => (
                  <div key={axis} className="field-row">
                    {number(
                      `${axis === "horizontal" ? "Horizontal" : "Vertical"} offset (pt)`,
                      draft.position![axis].offsetPt ?? 0,
                      (offsetPt) =>
                        setDraft((d) => ({
                          ...d,
                          position: {
                            ...d.position!,
                            [axis]: {
                              ...d.position![axis],
                              alignment: undefined,
                              offsetPt,
                            },
                          },
                        })),
                      -15840,
                    )}
                    <label className="field">
                      Relative to
                      <select
                        value={draft.position![axis].relativeTo}
                        onChange={(e) =>
                          setDraft((d) => ({
                            ...d,
                            position: {
                              ...d.position!,
                              [axis]: {
                                ...d.position![axis],
                                relativeTo: e.target.value,
                              },
                            },
                          }))
                        }
                      >
                        {(axis === "horizontal"
                          ? ["column", "margin", "page", "character"]
                          : ["paragraph", "margin", "page", "line"]
                        ).map((value) => (
                          <option key={value}>{value}</option>
                        ))}
                      </select>
                    </label>
                  </div>
                ))}
                <label className="check-field">
                  <input
                    type="checkbox"
                    checked={!!draft.position.locked}
                    onChange={(e) =>
                      setDraft((d) => ({
                        ...d,
                        position: { ...d.position!, locked: e.target.checked },
                      }))
                    }
                  />
                  Lock anchor
                </label>
              </>
            )}
          </>
        )}
        {tab === "Picture" && draft.image && (
          <>
            {number(
              "Brightness (%)",
              (draft.image.brightness ?? 0) * 100,
              (n) => image({ brightness: n / 100 }),
              -100,
              100,
            )}
            {number(
              "Contrast (%)",
              (draft.image.contrast ?? 0) * 100,
              (n) => image({ contrast: n / 100 }),
              -100,
              100,
            )}
            {number(
              "Transparency (%)",
              (1 - (draft.image.opacity ?? 1)) * 100,
              (n) => image({ opacity: 1 - n / 100 }),
              0,
              100,
            )}
            {number(
              "Blur (pt)",
              draft.image.blurPt ?? 0,
              (blurPt) => image({ blurPt }),
              0,
              30,
            )}
            <label className="check-field">
              <input
                type="checkbox"
                checked={!!draft.image.grayscale}
                onChange={(e) => image({ grayscale: e.target.checked })}
              />
              Grayscale
            </label>
          </>
        )}
        {tab === "Crop" && draft.image && (
          <div className="field-row">
            {(["left", "top", "right", "bottom"] as const).map((side) => (
              <div key={side}>
                {number(
                  `${side} (%)`,
                  (draft.image!.crop?.[side] ?? 0) * 100,
                  (n) =>
                    image({ crop: { ...draft.image!.crop, [side]: n / 100 } }),
                  0,
                  95,
                )}
              </div>
            ))}
          </div>
        )}
        {error && <p role="alert">{error}</p>}
        <div className="dialog-actions">
          <button type="button" onClick={onClose}>
            Cancel
          </button>
          <button className="primary-button">OK</button>
        </div>
      </form>
    </Dialog>
  );
}
export function TableFormatDialog({
  context,
  mode,
  onClose,
  onApply,
}: {
  context: { table: Table; row: number; cell: TableCell };
  mode: "properties" | "split" | "deleteCells";
  onClose: () => void;
  onApply: (
    patch: Omit<Extract<EditCommand, { type: "table" }>, "type" | "id">,
  ) => Promise<boolean>;
}) {
  const [error, setError] = useState("");
  return (
    <Dialog
      title={
        mode === "split"
          ? "Split Cells"
          : mode === "deleteCells"
            ? "Delete Cells"
            : "Table Properties"
      }
      onClose={onClose}
    >
      <form
        onSubmit={async (e) => {
          e.preventDefault();
          const data = new FormData(e.currentTarget);
          const command: Omit<
            Extract<EditCommand, { type: "table" }>,
            "type" | "id"
          > = mode === "split"
            ? {
                action: "split",
                rows: Number(data.get("rows")),
                columns: Number(data.get("columns")),
              }
            : mode === "deleteCells"
              ? {
                  action: String(data.get("delete")) as
                    | "deleteRow"
                    | "deleteColumn",
                }
              : {
                  format: {
                    cellMargins: Object.fromEntries(
                      ["top", "right", "bottom", "left"].map((side) => [
                        side,
                        Number(data.get(side)),
                      ]),
                    ) as {
                      top: number;
                      right: number;
                      bottom: number;
                      left: number;
                    },
                    rowHeightPt: Number(data.get("height")) || undefined,
                    columnWidthPt:
                      Math.abs(
                        Number(data.get("width")) -
                          (context.table.columnsPt[context.cell.column] ?? 72),
                      ) > 0.1
                        ? Number(data.get("width"))
                        : undefined,
                    borderColor:
                      data.get("borders") === "keep"
                        ? undefined
                        : String(data.get("color")),
                    borderWidthPt:
                      data.get("borders") === "keep"
                        ? undefined
                        : Number(data.get("pen")),
                    borders: (data.get("borders") === "keep"
                      ? undefined
                      : String(data.get("borders"))) as
                      | "all"
                      | "none"
                      | "outside"
                      | undefined,
                    alignment: String(data.get("alignment")) as
                      | "left"
                      | "right"
                      | "center",
                  },
                };
          if (await onApply(command)) onClose();
          else
            setError(
              "The table could not be changed. Check the document message for details.",
            );
        }}
      >
        {mode === "deleteCells" ? (
          <label className="field">
            Delete
            <select name="delete">
              <option value="deleteRow">Entire row</option>
              <option value="deleteColumn">Entire column</option>
            </select>
          </label>
        ) : mode === "split" ? (
          <div className="field-row">
            {["columns", "rows"].map((name) => (
              <label className="field" key={name}>
                {name}
                <input
                  name={name}
                  type="number"
                  min="1"
                  max="20"
                  defaultValue={
                    name === "columns"
                      ? context.cell.rowSpan > 1
                        ? context.cell.columnSpan
                        : Math.max(2, context.cell.columnSpan)
                      : context.cell.rowSpan
                  }
                />
              </label>
            ))}
          </div>
        ) : (
          <>
            <div className="field-row">
              {["top", "right", "bottom", "left"].map((side) => (
                <label className="field" key={side}>
                  {side} cell margin (pt)
                  <input
                    name={side}
                    type="number"
                    min="0"
                    max="144"
                    step="0.1"
                    defaultValue={
                      context.table.properties.cellMargins?.[side as "top"] ??
                      5.4
                    }
                  />
                </label>
              ))}
            </div>
            <div className="field-row">
              <label className="field">
                Row height (pt)
                <input
                  name="height"
                  type="number"
                  min="1"
                  max="1440"
                  defaultValue={context.table.rows[context.row].heightPt ?? ""}
                  placeholder="Automatic"
                />
              </label>
              <label className="field">
                Column width (pt)
                <input
                  name="width"
                  type="number"
                  min="6"
                  max="1440"
                  defaultValue={Number(
                    (
                      context.table.columnsPt[context.cell.column] ?? 72
                    ).toFixed(1),
                  )}
                />
              </label>
            </div>
            <div className="field-row">
              <label className="field">
                Table alignment
                <select
                  name="alignment"
                  defaultValue={context.table.properties.alignment ?? "left"}
                >
                  {["left", "center", "right"].map((a) => (
                    <option key={a}>{a}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                Borders
                <select name="borders">
                  <option value="keep">Keep existing</option>
                  <option value="all">All</option>
                  <option value="outside">Outside</option>
                  <option value="none">None</option>
                </select>
              </label>
              <label className="field">
                Pen color
                <input name="color" type="color" defaultValue="#000000" />
              </label>
              <label className="field">
                Pen width (pt)
                <input
                  name="pen"
                  type="number"
                  step=".25"
                  min=".25"
                  max="12"
                  defaultValue=".5"
                />
              </label>
            </div>
          </>
        )}
        {error && <p role="alert">{error}</p>}
        <div className="dialog-actions">
          <button type="button" onClick={onClose}>
            Cancel
          </button>
          <button className="primary-button">OK</button>
        </div>
      </form>
    </Dialog>
  );
}
