import type {
  Block,
  DocumentIR,
  Table,
  TableCell,
  Paragraph,
  TableRow,
} from "../model/document.js";
import type { EditCommand } from "./types.js";
import { getParagraphs } from "../model/traversal.js";
import { finite } from "./text.js";
import { columnWidth } from "./layout.js";
import { pageForParagraph } from "./sections.js";
export function tableContext(
  document: DocumentIR,
  id: string,
): { table: Table; row: number; cell: TableCell } | undefined {
  const visit = (blocks: Block[]): ReturnType<typeof tableContext> => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        for (const d of block.children)
          if (d.type === "drawing" && d.blocks) {
            const nested = visit(d.blocks);
            if (nested) return nested;
          }
        continue;
      }
      for (let row = 0; row < block.rows.length; row++)
        for (const cell of block.rows[row].cells) {
          const nested = visit(cell.blocks);
          if (nested) return nested;
          if (getParagraphs(cell.blocks).some((p) => p.id === id))
            return { table: block, row, cell };
        }
    }
  };
  return (
    visit(document.blocks) ??
    [
      ...Object.values(document.headers ?? {}),
      ...Object.values(document.footers ?? {}),
      ...getParagraphs(document.blocks).flatMap((p) => [
        ...Object.values(p.section?.headers ?? {}),
        ...Object.values(p.section?.footers ?? {}),
      ]),
    ]
      .map((b) => visit(b ?? []))
      .find(Boolean)
  );
}
type Command = Extract<EditCommand, { type: "table" }>;
export function editTable(
  document: DocumentIR,
  command: Command,
  newId: () => string,
) {
  const context = tableContext(document, command.id);
  if (!context) throw new Error("Select a table cell first.");
  const original = context.table,
    table = structuredClone(original),
    count = Math.max(
      table.columnsPt.length,
      ...table.rows.flatMap((r) => r.cells.map((c) => c.column + c.columnSpan)),
    );
  if (!count || table.rows.length > 200 || count > 100)
    throw new Error("This table is outside the supported editing size.");
  const base = getParagraphs(context.cell.blocks)[0];
  const paragraph = (): Paragraph => ({
    type: "paragraph",
    id: newId(),
    properties: {},
    directProperties: {
      styleId: base?.properties.styleId,
      direction: base?.properties.direction,
      alignment: base?.properties.alignment,
    },
    directMarkProperties: base?.directMarkProperties,
    children: [
      {
        type: "text",
        text: "",
        properties: {},
        directProperties:
          base?.children.filter((r) => r.type === "text").at(-1)
            ?.directProperties ?? {},
      },
    ],
  });
  const cell = (column: number): TableCell => ({
    column,
    columnSpan: 1,
    rowSpan: 1,
    blocks: [paragraph()],
  });
  const range = command.range ?? {
    top: context.row,
    bottom: context.row,
    left: context.cell.column,
    right: context.cell.column + context.cell.columnSpan - 1,
  };
  for (const [key, value] of Object.entries(range))
    if (
      !Number.isInteger(value) ||
      value < 0 ||
      value >= (key === "top" || key === "bottom" ? table.rows.length : count)
    )
      throw new Error("Select cells in one table.");
  if (range.top > range.bottom || range.left > range.right)
    throw new Error("Invalid table selection.");
  let removed = false;
  const cells = () =>
    table.rows.flatMap((r, i) =>
      r.cells.filter(
        (c) =>
          !c.continuation &&
          i >= range.top &&
          i <= range.bottom &&
          c.column >= range.left &&
          c.column <= range.right,
      ),
    );
  const width = columnWidth(pageForParagraph(document, command.id));
  table.columnsPt = Array.from(
    { length: count },
    (_, i) => table.columnsPt[i] || width / count,
  );
  const f = command.format;
  if (f) {
    if (f.shading && !/^#[\da-f]{6}$/i.test(f.shading) && f.shading !== "none")
      throw new Error("Choose an RGB cell color.");
    if (f.borderColor && !/^#[\da-f]{6}$/i.test(f.borderColor))
      throw new Error("Choose an RGB border color.");
    if (f.borderWidthPt !== undefined) finite(f.borderWidthPt, 0.25, 12);
    if (f.cellMargins)
      Object.values(f.cellMargins).forEach((v) => finite(v, 0, 144));
    if (f.alignment) table.properties.alignment = f.alignment;
    if (f.cellMargins) table.properties.cellMargins = { ...f.cellMargins };
    if (f.style) {
      table.properties.layout = "fixed";
      table.properties.width = { unit: "percent", value: 100 };
      for (let r = 0; r < table.rows.length; r++)
        for (const c of table.rows[r].cells) {
          c.backgroundColor =
            f.style === "banded"
              ? r === 0
                ? "#B4C6E7"
                : r % 2
                  ? "#E8EFF8"
                  : "#FFFFFF"
              : undefined;
          if (f.style === "plain")
            c.borders = Object.fromEntries(
              ["left", "right", "top", "bottom"].map((k) => [
                k,
                { style: "nil", widthPt: 0, color: "#000000" },
              ]),
            );
          else
            c.borders = Object.fromEntries(
              ["left", "right", "top", "bottom"].map((k) => [
                k,
                { style: "single", widthPt: 0.5, color: "#B8C2CF" },
              ]),
            );
        }
      table.rows[0].header = f.style === "banded";
    }
    for (const c of cells()) {
      const row = table.rows.findIndex((row) => row.cells.includes(c));
      if (f.shading)
        c.backgroundColor = f.shading === "none" ? undefined : f.shading;
      if (f.verticalAlign) c.verticalAlign = f.verticalAlign;
      if (f.borders)
        c.borders = Object.fromEntries(
          ["top", "right", "bottom", "left"].map((k) => [
            k,
            {
              style:
                f.borders === "none" ||
                (f.borders === "bottom" && k !== "bottom") ||
                (f.borders === "outside" &&
                  !(
                    (k === "left" && c.column === range.left) ||
                    (k === "right" &&
                      c.column + c.columnSpan - 1 === range.right) ||
                    (k === "top" && row === range.top) ||
                    (k === "bottom" && row + c.rowSpan - 1 === range.bottom)
                  ))
                  ? "nil"
                  : "single",
              widthPt: f.borderWidthPt ?? 0.5,
              color: f.borderColor ?? "#000000",
            },
          ]),
        );
    }
    for (let r = range.top; r <= range.bottom; r++) {
      if (f.rowHeightPt !== undefined) {
        table.rows[r].heightPt = finite(f.rowHeightPt, 1, 1440);
        table.rows[r].heightRule = "atLeast";
      }
      if (f.repeatHeader !== undefined) table.rows[r].header = f.repeatHeader;
      if (f.cantSplit !== undefined) table.rows[r].cantSplit = f.cantSplit;
    }
    if (f.columnWidthPt !== undefined)
      for (let c = range.left; c <= range.right; c++)
        table.columnsPt[c] = finite(f.columnWidthPt, 6, 1440);
  }
  type Anchor = { r: number; c: TableCell };
  let anchors: Anchor[] = table.rows.flatMap((row, r) =>
    row.cells.filter((c) => !c.continuation).map((c) => ({ r, c })),
  );
  const rows = table.rows;
  const action = command.action;
  if (action === "deleteTable") removed = true;
  else if (action === "autofitContents") {
    table.properties.layout = "autofit";
    table.properties.width = { unit: "percent", value: 100 };
  } else if (action === "autofitWindow") {
    table.properties.layout = "fixed";
    table.properties.width = { unit: "percent", value: 100 };
    table.columnsPt = table.columnsPt.map(() => width / count);
  } else if (action === "fixedWidth") {
    table.properties.layout = "fixed";
    table.properties.width = {
      unit: "pt",
      value: table.columnsPt.reduce((a, b) => a + b, 0),
    };
  } else if (action === "distributeColumns") {
    const sum = table.columnsPt
      .slice(range.left, range.right + 1)
      .reduce((a, b) => a + b, 0);
    for (let c = range.left; c <= range.right; c++)
      table.columnsPt[c] = sum / (range.right - range.left + 1);
  } else if (action === "distributeRows") {
    const height = Math.max(
      18,
      ...rows.slice(range.top, range.bottom + 1).map((r) => r.heightPt ?? 18),
    );
    for (let r = range.top; r <= range.bottom; r++) {
      rows[r].heightPt = height;
      rows[r].heightRule = "atLeast";
    }
  } else if (action === "insertRowAbove" || action === "insertRowBelow") {
    const at = context.row + (action === "insertRowBelow" ? 1 : 0);
    for (const a of anchors)
      if (a.r >= at) a.r++;
      else if (a.r + a.c.rowSpan > at) a.c.rowSpan++;
    rows.splice(at, 0, { cells: [] });
    for (let c = 0; c < count; c++)
      if (
        !anchors.some(
          (a) =>
            a.r <= at &&
            a.r + a.c.rowSpan > at &&
            a.c.column <= c &&
            a.c.column + a.c.columnSpan > c,
        )
      )
        anchors.push({ r: at, c: cell(c) });
  } else if (action === "deleteRow") {
    const at = context.row;
    anchors = anchors.filter((a) => a.r !== at || a.c.rowSpan > 1);
    for (const a of anchors) {
      if (a.r > at) a.r--;
      else if (a.r + a.c.rowSpan > at) a.c.rowSpan--;
    }
    rows.splice(at, 1);
    if (!rows.length) removed = true;
  } else if (
    action === "insertColumnBefore" ||
    action === "insertColumnAfter"
  ) {
    const at =
      context.cell.column +
      (action === "insertColumnAfter" ? context.cell.columnSpan : 0);
    for (const a of anchors)
      if (a.c.column >= at) a.c.column++;
      else if (a.c.column + a.c.columnSpan > at) a.c.columnSpan++;
    table.columnsPt.splice(at, 0, table.columnsPt[Math.min(at, count - 1)]);
    for (let r = 0; r < rows.length; r++)
      if (
        !anchors.some(
          (a) =>
            a.r <= r &&
            a.r + a.c.rowSpan > r &&
            a.c.column <= at &&
            a.c.column + a.c.columnSpan > at,
        )
      )
        anchors.push({ r, c: cell(at) });
  } else if (action === "deleteColumn") {
    const at = context.cell.column;
    anchors = anchors.filter((a) => a.c.column !== at || a.c.columnSpan > 1);
    for (const a of anchors) {
      if (a.c.column > at) a.c.column--;
      else if (a.c.column + a.c.columnSpan > at) a.c.columnSpan--;
    }
    table.columnsPt.splice(at, 1);
    if (!table.columnsPt.length) removed = true;
  } else if (action === "merge") {
    const selected = anchors.filter(
      (a) =>
        a.r >= range.top &&
        a.r <= range.bottom &&
        a.c.column >= range.left &&
        a.c.column <= range.right,
    );
    if (selected.length < 2)
      throw new Error("Select two or more adjacent cells to merge.");
    if (
      anchors.some(
        (a) =>
          a.r <= range.bottom &&
          a.r + a.c.rowSpan > range.top &&
          a.c.column <= range.right &&
          a.c.column + a.c.columnSpan > range.left &&
          (a.r < range.top ||
            a.r + a.c.rowSpan - 1 > range.bottom ||
            a.c.column < range.left ||
            a.c.column + a.c.columnSpan - 1 > range.right),
      )
    )
      throw new Error("Select complete merged cells.");
    const first = selected[0];
    first.c = {
      ...first.c,
      blocks: selected.flatMap((a) => a.c.blocks),
      column: range.left,
      columnSpan: range.right - range.left + 1,
      rowSpan: range.bottom - range.top + 1,
    };
    first.r = range.top;
    anchors = anchors.filter((a) => a === first || !selected.includes(a));
  } else if (action === "split") {
    const a = anchors.find(
      (a) => a.r === context.row && a.c.column === context.cell.column,
    )!;
    const columns = Math.round(finite(command.columns ?? 2, 1, 20)),
      rowCount = Math.round(finite(command.rows ?? 1, 1, 20));
    if (a.c.rowSpan !== 1 || a.c.columnSpan !== 1) {
      if (columns !== a.c.columnSpan || rowCount !== a.c.rowSpan)
        throw new Error(
          "For a merged cell, split into its existing grid dimensions.",
        );
      anchors = anchors.filter((x) => x !== a);
      for (let r = 0; r < rowCount; r++)
        for (let c = 0; c < columns; c++)
          anchors.push({
            r: a.r + r,
            c: {
              ...cell(a.c.column + c),
              ...(r === 0 && c === 0 ? { blocks: a.c.blocks } : {}),
            },
          });
    } else {
      if (rowCount > 1) {
        for (const x of anchors)
          if (x !== a) {
            if (x.r > a.r) x.r += rowCount - 1;
            else if (x.r <= a.r && x.r + x.c.rowSpan > a.r)
              x.c.rowSpan += rowCount - 1;
          }
        const originalRow = rows[a.r];
        if (originalRow.heightPt) originalRow.heightPt /= rowCount;
        rows.splice(
          a.r + 1,
          0,
          ...Array.from({ length: rowCount - 1 }, () => ({
            ...originalRow,
            cells: [],
          })),
        );
      }
      const at = a.c.column;
      for (const x of anchors)
        if (x !== a) {
          if (x.c.column > at) x.c.column += columns - 1;
          else if (x.c.column <= at && x.c.column + x.c.columnSpan > at)
            x.c.columnSpan += columns - 1;
        }
      const w = table.columnsPt[at] / columns;
      table.columnsPt.splice(at, 1, ...Array(columns).fill(w));
      for (let r = 0; r < rowCount; r++)
        for (let c = 0; c < columns; c++)
          if (r || c) anchors.push({ r: a.r + r, c: cell(at + c) });
    }
  }
  if (
    action &&
    ![
      "deleteTable",
      "autofitContents",
      "autofitWindow",
      "fixedWidth",
      "distributeColumns",
      "distributeRows",
    ].includes(action) &&
    !removed
  ) {
    for (let r = 0; r < rows.length; r++) {
      rows[r].cells = anchors.filter((a) => a.r === r).map((a) => a.c);
      for (const a of anchors.filter((a) => a.r < r && a.r + a.c.rowSpan > r)) {
        const old = original.rows[r]?.cells.find(
          (c) => c.continuation && c.column === a.c.column,
        );
        rows[r].cells.push({
          ...a.c,
          continuation: true,
          rowSpan: 1,
          blocks: old?.blocks ?? [paragraph()],
        });
      }
      rows[r].cells.sort((a, b) => a.column - b.column);
    }
  }
  const replace = (blocks: Block[]): Block[] =>
    blocks.flatMap((b): Block[] =>
      b === original
        ? removed
          ? []
          : [table]
        : b.type === "table"
          ? [
              {
                ...b,
                rows: b.rows.map((r) => ({
                  ...r,
                  cells: r.cells.map((c) => ({
                    ...c,
                    blocks: replace(c.blocks),
                  })),
                })),
              },
            ]
          : [
              {
                ...b,
                children: b.children.map((d) =>
                  d.type === "drawing" && d.blocks
                    ? { ...d, blocks: replace(d.blocks) }
                    : d,
                ),
                ...(b.section
                  ? {
                      section: {
                        ...b.section,
                        ...Object.fromEntries(
                          (["headers", "footers"] as const)
                            .filter((k) => b.section?.[k])
                            .map((k) => [
                              k,
                              Object.fromEntries(
                                Object.entries(b.section![k]!).map(
                                  ([v, blocks]) => [v, replace(blocks!)],
                                ),
                              ),
                            ]),
                        ),
                      },
                    }
                  : {}),
              },
            ],
    );
  const next = { ...document, blocks: replace(document.blocks) };
  for (const kind of ["headers", "footers"] as const)
    if (next[kind])
      next[kind] = Object.fromEntries(
        Object.entries(next[kind]!).map(([v, b]) => [v, replace(b!)]),
      );
  const selected = getParagraphs(
    removed
      ? next.blocks
      : table.rows[Math.min(context.row, table.rows.length - 1)].cells[0]
          .blocks,
  )[0];
  return {
    document: next,
    table: removed ? undefined : table,
    original,
    selection: selected?.id ? { id: selected.id, start: 0, end: 0 } : undefined,
  };
}
