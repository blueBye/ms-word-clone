import type { DocumentIR, DrawingInline } from "../model/document.js";
import { getParagraphs } from "../model/traversal.js";
import type { EditCommand } from "./types.js";
import { splitSections, pageForParagraph } from "./sections.js";
export function arrangeObjects(
  document: DocumentIR,
  command: Extract<EditCommand, { type: "arrange" }>,
): EditCommand[] {
  const paragraphs = [
    ...splitSections(document).flatMap((section) => [
      ...getParagraphs(section.blocks),
      ...Object.values(section.headers ?? {}).flatMap((b) =>
        getParagraphs(b ?? []),
      ),
      ...Object.values(section.footers ?? {}).flatMap((b) =>
        getParagraphs(b ?? []),
      ),
    ]),
    ...Object.values(document.headers ?? {}).flatMap((b) =>
      getParagraphs(b ?? []),
    ),
    ...Object.values(document.footers ?? {}).flatMap((b) =>
      getParagraphs(b ?? []),
    ),
  ];
  const objects = command.objects.map((ref) => {
    const drawing = paragraphs.find((p) => p.id === ref.id)?.children[
      ref.index
    ];
    if (drawing?.type !== "drawing") throw new Error("Select an object first.");
    return { ...ref, drawing: structuredClone(drawing) };
  });
  if (!objects.length) throw new Error("Select an object first.");
  const remove = () =>
    objects
      .slice()
      .sort((a, b) => a.id.localeCompare(b.id) || b.index - a.index)
      .map((o) => ({
        type: "drawing" as const,
        id: o.id,
        index: o.index,
        remove: true,
      }));
  if (command.action === "delete") return remove();
  if (command.action === "group") {
    if (objects.length < 2)
      throw new Error("Select at least two objects to group.");
    if (
      objects.some(
        (o) =>
          o.drawing.chart ||
          o.drawing.equation ||
          o.drawing.attachment ||
          o.drawing.spreadsheet ||
          o.drawing.model3d ||
          o.drawing.video,
      )
    )
      throw new Error(
        "Group pictures, shapes and text boxes. Charts, equations and embedded media keep their own native object types.",
      );
    const x = Math.min(
        ...objects.map((o) => o.drawing.position?.horizontal.offsetPt ?? 0),
      ),
      y = Math.min(
        ...objects.map((o) => o.drawing.position?.vertical.offsetPt ?? 0),
      );
    const children = objects.map(({ drawing: d }) => ({
      ...d,
      position: {
        horizontal: {
          relativeTo: "paragraph",
          offsetPt: (d.position?.horizontal.offsetPt ?? 0) - x,
        },
        vertical: {
          relativeTo: "paragraph",
          offsetPt: (d.position?.vertical.offsetPt ?? 0) - y,
        },
        behindText: false,
        zIndex: 2,
        wrap: "none",
      },
    }));
    const group: DrawingInline = {
      type: "drawing",
      alt: "Group",
      widthPt: Math.max(
        ...children.map(
          (d) => (d.position.horizontal.offsetPt ?? 0) + d.widthPt,
        ),
      ),
      heightPt: Math.max(
        ...children.map(
          (d) => (d.position.vertical.offsetPt ?? 0) + d.heightPt,
        ),
      ),
      children,
      ...(objects[0].drawing.position
        ? {
            position: {
              ...structuredClone(objects[0].drawing.position),
              horizontal: {
                ...objects[0].drawing.position.horizontal,
                offsetPt: x,
              },
              vertical: {
                ...objects[0].drawing.position.vertical,
                offsetPt: y,
              },
            },
          }
        : {}),
    };
    return [
      ...remove(),
      { type: "insertDrawing", id: objects[0].id, drawing: group },
    ];
  }
  if (command.action === "ungroup") {
    const group = objects[0];
    if (!group.drawing.children) throw new Error("Select a group to ungroup.");
    return [
      { type: "drawing", id: group.id, index: group.index, remove: true },
      ...group.drawing.children.map((d) => {
        const g = group.drawing,
          angle = ((g.shape?.rotation ?? 0) * Math.PI) / 180,
          flip = !!g.flipH !== !!g.flipV ? -1 : 1;
        const x =
            ((d.position?.horizontal.offsetPt ?? 0) +
              d.widthPt / 2 -
              g.widthPt / 2) *
            (g.flipH ? -1 : 1),
          y =
            ((d.position?.vertical.offsetPt ?? 0) +
              d.heightPt / 2 -
              g.heightPt / 2) *
            (g.flipV ? -1 : 1);
        return {
          type: "insertDrawing" as const,
          id: group.id,
          drawing: {
            ...d,
            shape: {
              ...d.shape,
              kind: d.shape?.kind ?? "rect",
              rotation:
                (g.shape?.rotation ?? 0) + (d.shape?.rotation ?? 0) * flip,
            },
            flipH: !!g.flipH !== !!d.flipH,
            flipV: !!g.flipV !== !!d.flipV,
            position: {
              horizontal: {
                ...g.position?.horizontal,
                relativeTo: g.position?.horizontal.relativeTo ?? "paragraph",
                offsetPt:
                  (g.position?.horizontal.offsetPt ?? 0) +
                  g.widthPt / 2 +
                  x * Math.cos(angle) -
                  y * Math.sin(angle) -
                  d.widthPt / 2,
              },
              vertical: {
                ...g.position?.vertical,
                relativeTo: g.position?.vertical.relativeTo ?? "paragraph",
                offsetPt:
                  (g.position?.vertical.offsetPt ?? 0) +
                  g.heightPt / 2 +
                  x * Math.sin(angle) +
                  y * Math.cos(angle) -
                  d.heightPt / 2,
              },
              behindText: g.position?.behindText ?? false,
              zIndex: g.position?.zIndex ?? 2,
              wrap: g.position?.wrap ?? "wrapSquare",
            },
          },
        };
      }),
    ];
  }
  const [action, value] = command.action.split(":");
  const page = pageForParagraph(document, objects[0].id);
  const width =
      (page.widthPt ?? 595.3) -
      (page.margins?.left ?? 72) -
      (page.margins?.right ?? 72),
    height =
      (page.heightPt ?? 841.9) -
      (page.margins?.top ?? 72) -
      (page.margins?.bottom ?? 72);
  const sorted = objects
    .slice()
    .sort(
      (a, b) =>
        (value === "vertical"
          ? (a.drawing.position?.vertical.offsetPt ?? 0)
          : (a.drawing.position?.horizontal.offsetPt ?? 0)) -
        (value === "vertical"
          ? (b.drawing.position?.vertical.offsetPt ?? 0)
          : (b.drawing.position?.horizontal.offsetPt ?? 0)),
    );
  return objects.map((o) => {
    const d = o.drawing;
    const position = () =>
      (d.position ??= {
        horizontal: { relativeTo: "margin", offsetPt: 0 },
        vertical: { relativeTo: "paragraph", offsetPt: 0 },
        behindText: false,
        zIndex: 2,
        wrap: "wrapSquare",
      });
    if (action === "inline") delete d.position;
    else if (action === "move") {
      const [x, y] = value.split(",").map(Number),
        p = position();
      if (!Number.isFinite(x) || !Number.isFinite(y))
        throw new Error("Invalid object movement.");
      p.horizontal.alignment = undefined;
      p.vertical.alignment = undefined;
      p.horizontal.offsetPt = (p.horizontal.offsetPt ?? 0) + x;
      p.vertical.offsetPt = (p.vertical.offsetPt ?? 0) + y;
    } else if (action === "flipH" || action === "flipV") d[action] = !d[action];
    else if (action === "rotate")
      d.shape = {
        ...d.shape,
        kind: d.shape?.kind ?? "rect",
        rotation: (d.shape?.rotation ?? 0) + Number(value),
      };
    else if (action === "wrap") position().wrap = value;
    else if (action === "behind" || action === "front") {
      position().behindText = action === "behind";
      position().wrap = "wrapNone";
    } else if (["forward", "backward", "toFront", "toBack"].includes(action))
      position().zIndex = Math.max(
        0,
        action === "toFront"
          ? 100000
          : action === "toBack"
            ? 0
            : position().zIndex + (action === "forward" ? 1 : -1),
      );
    else if (action === "position" || action === "align") {
      const [v, h] = action === "position" ? value.split("-") : [value, value];
      const p = position();
      p.horizontal.relativeTo = "margin";
      p.vertical.relativeTo = "margin";
      if (["left", "center", "right"].includes(h)) {
        p.horizontal.alignment = undefined;
        p.horizontal.offsetPt =
          h === "left"
            ? 0
            : h === "center"
              ? (width - d.widthPt) / 2
              : width - d.widthPt;
      }
      if (["top", "middle", "bottom"].includes(v)) {
        p.vertical.alignment = undefined;
        p.vertical.offsetPt =
          v === "top"
            ? 0
            : v === "middle"
              ? (height - d.heightPt) / 2
              : height - d.heightPt;
      }
    } else if (action === "distribute") {
      if (sorted.length < 3)
        throw new Error("Select at least three objects to distribute.");
      const axis = value === "vertical" ? "vertical" : "horizontal",
        first = sorted[0].drawing.position?.[axis].offsetPt ?? 0,
        last = sorted.at(-1)!.drawing.position?.[axis].offsetPt ?? 0;
      position()[axis].offsetPt =
        first + ((last - first) * sorted.indexOf(o)) / (sorted.length - 1);
    }
    return { type: "drawing", id: o.id, index: o.index, drawing: d };
  });
}
