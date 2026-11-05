import type { Block, Paragraph, TextRun, DrawingInline } from "./document.js";

export function getParagraphs(blocks: readonly Block[]): Paragraph[] {
  const nested = (d: DrawingInline): Paragraph[] => [
    ...getParagraphs(d.blocks ?? []),
    ...(d.children ?? []).flatMap(nested),
  ];
  return blocks.flatMap((block) =>
    block.type === "paragraph"
      ? [
          block,
          ...block.children.flatMap((inline) =>
            inline.type === "drawing" ? nested(inline) : [],
          ),
        ]
      : block.rows.flatMap((row) =>
          row.cells.flatMap((cell) => getParagraphs(cell.blocks)),
        ),
  );
}

export function getTextRuns(blocks: readonly Block[]): TextRun[] {
  return getParagraphs(blocks).flatMap((p) =>
    p.children.filter((inline): inline is TextRun => inline.type === "text"),
  );
}
