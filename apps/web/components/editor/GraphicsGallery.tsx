import { useState } from "react";
import { iconLibrary, iconPaths, shapePaths } from "@local/docx-ir-engine/core";
import { Dialog } from "./Controls";
export function GraphicPreview({ kind }: { kind: string }) {
  const path = shapePaths[kind] ?? iconPaths[kind];
  return (
    <svg
      viewBox="0 0 100 100"
      aria-hidden="true"
      fill={kind.startsWith("icon_") ? "currentColor" : "#dce9f6"}
      stroke={kind.startsWith("icon_") ? "none" : "#4472c4"}
      strokeWidth="3"
    >
      {path ? (
        <path d={path} />
      ) : kind === "ellipse" ? (
        <ellipse cx="50" cy="50" rx="46" ry="40" />
      ) : (
        <rect
          x="5"
          y="10"
          width="90"
          height="80"
          rx={kind === "roundRect" ? 15 : 0}
        />
      )}
    </svg>
  );
}
const shapeGroups = [
  ["Lines", [["line", "Line"]]],
  [
    "Rectangles",
    [
      ["rect", "Rectangle"],
      ["roundRect", "Rounded Rectangle"],
    ],
  ],
  [
    "Basic Shapes",
    [
      ["ellipse", "Oval"],
      ["triangle", "Isosceles Triangle"],
      ["rtTriangle", "Right Triangle"],
      ["diamond", "Diamond"],
      ["pentagon", "Pentagon"],
      ["hexagon", "Hexagon"],
      ["plus", "Cross"],
      ["heart", "Heart"],
    ],
  ],
  [
    "Block Arrows",
    [
      ["rightArrow", "Right Arrow"],
      ["leftArrow", "Left Arrow"],
      ["upArrow", "Up Arrow"],
      ["downArrow", "Down Arrow"],
      ["chevron", "Chevron"],
    ],
  ],
  ["Stars", [["star5", "Five-Point Star"]]],
] as const;
export function ShapesGallery({ onPick }: { onPick: (kind: string) => void }) {
  return (
    <div className="shapes-gallery">
      {shapeGroups.map(([label, shapes]) => (
        <section key={label}>
          <p className="office-menu-heading">{label}</p>
          <div>
            {shapes.map(([kind, name]) => (
              <button
                type="button"
                key={kind}
                title={name}
                aria-label={name}
                onClick={() => onPick(kind)}
              >
                <GraphicPreview kind={kind} />
              </button>
            ))}
          </div>
        </section>
      ))}
    </div>
  );
}
export function IconsDialog({
  onClose,
  onInsert,
}: {
  onClose: () => void;
  onInsert: (ids: string[]) => Promise<boolean>;
}) {
  const [search, setSearch] = useState(""),
    [category, setCategory] = useState("All"),
    [selected, setSelected] = useState<string[]>([]),
    [busy, setBusy] = useState(false);
  const icons = iconLibrary.filter(
    (i) =>
      (category === "All" || i.category === category) &&
      `${i.name} ${i.category}`.toLowerCase().includes(search.toLowerCase()),
  );
  return (
    <Dialog title="Insert Icons" onClose={onClose}>
      <div className="icons-library">
        <input
          type="search"
          autoFocus
          placeholder="Search icons"
          aria-label="Search icons"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
        <div className="icons-library-body">
          <nav aria-label="Icon categories">
            {["All", ...new Set(iconLibrary.map((i) => i.category))].map(
              (name) => (
                <button
                  type="button"
                  key={name}
                  aria-pressed={name === category}
                  onClick={() => setCategory(name)}
                >
                  {name}
                </button>
              ),
            )}
          </nav>
          <div className="icons-library-grid">
            {icons.map((icon) => (
              <button
                type="button"
                key={icon.id}
                title={icon.name}
                aria-pressed={selected.includes(icon.id)}
                onClick={() =>
                  setSelected((s) =>
                    s.includes(icon.id)
                      ? s.filter((id) => id !== icon.id)
                      : [...s, icon.id],
                  )
                }
              >
                <GraphicPreview kind={`icon_${icon.id}`} />
                <span>{icon.name}</span>
              </button>
            ))}
            {!icons.length && <p>No icons match your search.</p>}
          </div>
        </div>
        <div className="dialog-actions">
          <span>{selected.length} selected</span>
          <button type="button" className="secondary-button" onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className="primary-button"
            disabled={!selected.length || busy}
            onClick={async () => {
              setBusy(true);
              try {
                if (await onInsert(selected)) onClose();
              } finally {
                setBusy(false);
              }
            }}
          >
            Insert
          </button>
        </div>
      </div>
    </Dialog>
  );
}
