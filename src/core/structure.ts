import type { Block, DocumentIR, Paragraph } from "../model/document.js";
import type { EditCommand, ParagraphCapability } from "./types.js";
import {
  paragraphText,
  rewriteParagraph,
  validateRange,
  validateText,
  finite,
} from "./text.js";
import { ListStore } from "./lists.js";

export function editStructure(
  document: DocumentIR,
  command: EditCommand,
  lists: ListStore,
  newId: () => string,
  capability: (id: string) => ParagraphCapability,
) {
  let found = false,
    selection: { id: string; start: number; end: number } | undefined;
  const paragraph = (text = ""): Paragraph => ({
    type: "paragraph",
    id: newId(),
    properties: {},
    directProperties: {},
    children: [{ type: "text", text, properties: {}, directProperties: {} }],
  });
  const editable = (p: Paragraph) => {
    const allowed = capability(p.id!);
    if (!allowed.text) throw new Error(allowed.reason);
  };
  const transform = (blocks: Block[]): Block[] => {
    let result = blocks;
    const id =
      "id" in command
        ? command.id
        : command.type === "sort"
          ? command.ids[0]
          : undefined;
    const index = id
      ? blocks.findIndex((b) => b.type === "paragraph" && b.id === id)
      : -1;
    if (index >= 0 && !found) {
      found = true;
      result = blocks.slice();
      const p = blocks[index] as Paragraph;
      if (command.type === "splitParagraph") {
        editable(p);
        const text = paragraphText(p);
        validateRange(text, command.start, command.end);
        if (!text.length && p.properties.list)
          result[index] = lists.set(
            p,
            { type: "list", id: p.id!, kind: "none" },
            blocks.filter((b) => b.type === "paragraph"),
          );
        else {
          const first = rewriteParagraph(p, {
            type: "text",
            id: p.id!,
            start: command.start,
            end: text.length,
            text: "",
          });
          const second = {
            ...rewriteParagraph(p, {
              type: "text",
              id: p.id!,
              start: 0,
              end: command.end,
              text: "",
            }),
            id: newId(),
          };
          result.splice(index, 1, first, second);
          selection = { id: second.id, start: 0, end: 0 };
        }
      } else if (command.type === "joinParagraph") {
        editable(p);
        const other =
          blocks[index + (command.direction === "backward" ? -1 : 1)];
        if (other?.type !== "paragraph") return blocks;
        editable(other);
        const first = command.direction === "backward" ? other : p,
          second = command.direction === "backward" ? p : other,
          offset = paragraphText(first).length;
        result.splice(Math.min(index, blocks.indexOf(other)), 2, {
          ...first,
          children: [...first.children, ...second.children],
        });
        selection = { id: first.id!, start: offset, end: offset };
      } else if (command.type === "list") {
        editable(p);
        const neighbours: Paragraph[] = [];
        for (let i = index - 1; i >= 0 && blocks[i].type === "paragraph"; i--)
          neighbours.unshift(blocks[i] as Paragraph);
        neighbours.push(p);
        for (
          let i = index + 1;
          i < blocks.length && blocks[i].type === "paragraph";
          i++
        )
          neighbours.push(blocks[i] as Paragraph);
        result[index] = lists.set(p, command, neighbours);
        selection = { id: p.id!, start: 0, end: 0 };
      } else if (command.type === "sort") {
        const items = blocks.slice(index, index + command.ids.length);
        if (
          items.some(
            (b, i) => b.type !== "paragraph" || b.id !== command.ids[i],
          )
        )
          throw new Error(
            "Sort requires consecutive paragraphs in one container.",
          );
        items.forEach((b) => editable(b as Paragraph));
        const collator = new Intl.Collator(command.locale, {
          numeric: command.numeric,
          sensitivity: "base",
        });
        result.splice(
          index,
          items.length,
          ...items
            .slice()
            .sort(
              (a, b) =>
                collator.compare(
                  paragraphText(a as Paragraph),
                  paragraphText(b as Paragraph),
                ) * (command.descending ? -1 : 1),
            ),
        );
      } else if (
        command.type === "insertParagraph" ||
        command.type === "insertTable"
      )
        result.splice(index + 1, 0, makeBlock(command));
      else throw new Error(`Unsupported document command: ${command.type}`);
    }
    return result.map((b): Block => {
      if (b.type === "table")
        return {
          ...b,
          rows: b.rows.map((r) => ({
            ...r,
            cells: r.cells.map((c) => ({ ...c, blocks: transform(c.blocks) })),
          })),
        };
      return {
        ...b,
        children: b.children.map((r) =>
          r.type === "drawing" && r.blocks
            ? { ...r, blocks: transform(r.blocks) }
            : r,
        ),
      };
    });
  };
  const makeBlock = (
    insert: Extract<EditCommand, { type: "insertParagraph" | "insertTable" }>,
  ): Block => {
    if (insert.type === "insertParagraph") {
      validateText(insert.text);
      const p = paragraph(insert.text);
      if (insert.pageBreakBefore)
        p.directProperties = { pageBreakBefore: true };
      return p;
    }
    const rows = finite(insert.rows, 1, 50),
      columns = finite(insert.columns, 1, 20);
    if (!Number.isInteger(rows) || !Number.isInteger(columns))
      throw new Error("Table dimensions must be whole numbers.");
    return {
      type: "table",
      properties: { width: { unit: "percent", value: 100 } },
      columnsPt: [],
      rows: Array.from({ length: rows }, () => ({
        cells: Array.from({ length: columns }, (_, column) => ({
          column,
          columnSpan: 1,
          rowSpan: 1,
          blocks: [paragraph()],
        })),
      })),
    };
  };
  const result = { ...document, blocks: transform(document.blocks) };
  for (const kind of ["headers", "footers"] as const)
    if (result[kind]) {
      result[kind] = { ...result[kind] };
      for (const key of ["default", "first", "even"] as const)
        if (result[kind][key]) result[kind][key] = transform(result[kind][key]);
    }
  if (!found) {
    if (
      (command.type === "insertParagraph" || command.type === "insertTable") &&
      !command.id
    )
      result.blocks = [...result.blocks, makeBlock(command)];
    else throw new Error("Select a paragraph in the current document.");
  }
  return { document: result, selection };
}
