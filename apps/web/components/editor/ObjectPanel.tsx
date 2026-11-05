import type { Paragraph, DrawingInline } from "@local/docx-ir-engine/core";
export type ObjectRef = { id: string; index: number };
export function documentObjects(all: Paragraph[]) {
  return all.flatMap((p) =>
    p.children.flatMap((d, index) =>
      d.type === "drawing" ? [{ id: p.id!, index, drawing: d }] : [],
    ),
  );
}
export function ObjectPanel({
  all,
  selected,
  onSelect,
  onChange,
  onAction,
}: {
  all: Paragraph[];
  selected: ObjectRef[];
  onSelect: (ref: ObjectRef, multi: boolean) => void;
  onChange: (ref: ObjectRef, d: DrawingInline) => void;
  onAction: (action: string) => void;
}) {
  const objects = documentObjects(all),
    current = objects.find((o) =>
      selected.some((s) => s.id === o.id && s.index === o.index),
    );
  return (
    <>
      <p className="pane-description">
        Select an object on the page or here. Hold Ctrl to select multiple
        objects.
      </p>
      <div className="object-list">
        {objects.map((o, i) => (
          <button
            key={`${o.id}-${o.index}`}
            className={
              selected.some((s) => s.id === o.id && s.index === o.index)
                ? "selected"
                : ""
            }
            onClick={(e) => onSelect(o, e.ctrlKey || e.metaKey)}
          >
            {o.drawing.alt || `Object ${i + 1}`}
            {o.drawing.children ? " (group)" : ""}
          </button>
        ))}
      </div>
      {current && (
        <div className="object-properties">
          <label className="field">
            Name
            <input
              key={current.drawing.alt}
              defaultValue={current.drawing.alt}
              onBlur={(e) =>
                onChange(current, { ...current.drawing, alt: e.target.value })
              }
            />
          </label>
          {[
            ["Width (pt)", "widthPt", current.drawing.widthPt],
            ["Height (pt)", "heightPt", current.drawing.heightPt],
            [
              "Rotation (degrees)",
              "rotation",
              current.drawing.shape?.rotation ?? 0,
            ],
            [
              "Horizontal position (pt)",
              "x",
              current.drawing.position?.horizontal.offsetPt ?? 0,
            ],
            [
              "Vertical position (pt)",
              "y",
              current.drawing.position?.vertical.offsetPt ?? 0,
            ],
          ].map(([label, key, value]) => (
            <label className="field" key={key}>
              {label}
              <input
                type="number"
                key={String(value)}
                defaultValue={value}
                onBlur={(e) => {
                  const n = e.target.valueAsNumber;
                  if (!Number.isFinite(n)) return;
                  const d = structuredClone(current.drawing);
                  if (key === "rotation")
                    d.shape = {
                      ...d.shape,
                      kind: d.shape?.kind ?? "rect",
                      rotation: n,
                    };
                  else if (key === "x" || key === "y") {
                    d.position ??= {
                      horizontal: { relativeTo: "margin", offsetPt: 0 },
                      vertical: { relativeTo: "paragraph", offsetPt: 0 },
                      behindText: false,
                      zIndex: 2,
                      wrap: "wrapSquare",
                    };
                    if (key === "x") d.position.horizontal.offsetPt = n;
                    else d.position.vertical.offsetPt = n;
                  } else d[key as "widthPt" | "heightPt"] = n;
                  onChange(current, d);
                }}
              />
            </label>
          ))}
          {current.drawing.image && (
            <label className="field">
              Opacity
              <input
                type="range"
                min="0"
                max="1"
                step=".05"
                defaultValue={current.drawing.image.opacity ?? 1}
                onPointerUp={(e) =>
                  onChange(current, {
                    ...current.drawing,
                    image: {
                      ...current.drawing.image!,
                      opacity: Number(e.currentTarget.value),
                    },
                  })
                }
              />
            </label>
          )}
          {(current.drawing.chart ||
            current.drawing.equation ||
            current.drawing.spreadsheet) && (
            <button
              className="secondary-button"
              onClick={() => onAction("object:edit")}
            >
              Edit content…
            </button>
          )}
          <button
            className="secondary-button"
            onClick={() => onAction("object:delete")}
          >
            Delete selected object
          </button>
        </div>
      )}
    </>
  );
}
