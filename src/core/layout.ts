import type {
  DocumentIR,
  DrawingInline,
  Block,
  PageLayout,
} from "../model/document.js";
import type { EditCommand } from "./types.js";
import { pageForParagraph, sectionTarget } from "./sections.js";
import { getParagraphs } from "../model/traversal.js";
export function columnWidth(page: PageLayout = {}) {
  const count = page.columns?.count ?? 1;
  return Math.max(
    12,
    ((page.widthPt ?? 595.3) -
      (page.margins?.left ?? 72) -
      (page.margins?.right ?? 72) -
      (page.gutterPt ?? 0) -
      (count - 1) * (page.columns?.gapPt ?? 36)) /
      count,
  );
}
export function fitDrawing(d: DrawingInline, available: number): DrawingInline {
  available = Math.max(12, available);
  if (d.widthPt <= available) return d;
  const ratio = available / d.widthPt;
  const scale = (d: DrawingInline, child = false): DrawingInline => ({
    ...d,
    widthPt: d.widthPt * ratio,
    heightPt: Math.max(1, d.heightPt * ratio),
    ...(child && d.position
      ? {
          position: {
            ...d.position,
            horizontal: {
              ...d.position.horizontal,
              offsetPt: (d.position.horizontal.offsetPt ?? 0) * ratio,
            },
            vertical: {
              ...d.position.vertical,
              offsetPt: (d.position.vertical.offsetPt ?? 0) * ratio,
            },
          },
        }
      : {}),
    ...(d.children ? { children: d.children.map((c) => scale(c, true)) } : {}),
  });
  return scale(d);
}
/** A narrower text column constrains flow objects; explicit page/margin anchors retain their scope. */
export function fitSectionDrawings(
  document: DocumentIR,
  id?: string,
): EditCommand[] {
  const target = sectionTarget(document, id)?.id,
    commands: EditCommand[] = [];
  const visit = (blocks: Block[], width?: number) => {
    for (const block of blocks) {
      if (block.type === "table") {
        const first = getParagraphs([block])[0];
        const containerWidth =
          width ??
          columnWidth(
            pageForParagraph(
              document,
              first?.type === "paragraph" ? first.id : undefined,
            ),
          );
        const gridWidth = block.columnsPt.reduce((sum, n) => sum + n, 0),
          tableScale = gridWidth ? Math.min(1, containerWidth / gridWidth) : 1;
        for (const row of block.rows)
          for (const cell of row.cells) {
            const cellWidth = block.columnsPt
              .slice(cell.column, cell.column + cell.columnSpan)
              .reduce((n, v) => n + v, 0);
            visit(
              cell.blocks,
              cellWidth
                ? Math.min(containerWidth, cellWidth * tableScale) -
                    (cell.margins?.left ??
                      block.properties.cellMargins?.left ??
                      5.4) -
                    (cell.margins?.right ??
                      block.properties.cellMargins?.right ??
                      5.4)
                : containerWidth,
            );
          }
        continue;
      }
      if (sectionTarget(document, block.id)?.id !== target) continue;
      const available =
        width ?? columnWidth(pageForParagraph(document, block.id));
      block.children.forEach((d, index) => {
        if (d.type !== "drawing") return;
        const scoped =
          d.position &&
          ["page", "margin", "leftMargin", "rightMargin"].includes(
            d.position.horizontal.relativeTo,
          );
        const next = scoped ? d : fitDrawing(d, available);
        if (next !== d && block.id)
          commands.push({
            type: "drawing",
            id: block.id,
            index,
            drawing: next,
          });
      });
    }
  };
  visit(document.blocks);
  return commands;
}
