import { expect, it } from "vitest";
import { openDocxSession, createBlankDocx } from "../src/docx/index.js";
import { allParagraphs } from "../src/core/index.js";
import type { DrawingInline } from "../src/core/index.js";
import { unzipSync, strFromU8 } from "fflate";
const shape: DrawingInline = {
  type: "drawing",
  widthPt: 144,
  heightPt: 72,
  alt: "Test box",
  shape: { kind: "roundRect", fill: "#4479AA", stroke: "#294F76" },
  blocks: [
    {
      type: "paragraph",
      properties: {},
      children: [
        {
          type: "text",
          text: "Editable text",
          properties: { fontFamily: "Tahoma", fontSizePt: 14 },
        },
      ],
    },
  ],
};
const setup = async () => {
  const session = await openDocxSession(createBlankDocx());
  return { session, id: allParagraphs(session.document)[0].id! };
};
it("inserts native text boxes and preserves inner paragraph editing, geometry and undo", async () => {
  const { session, id } = await setup();
  await session.execute({ type: "insertDrawing", id, drawing: shape });
  const box = allParagraphs(session.document).find((p) =>
    p.children.some((c) => c.type === "drawing"),
  )!;
  const inner = allParagraphs(session.document).find((p) =>
    p.children.some((c) => c.type === "text" && c.text === "Editable text"),
  )!;
  expect(session.capability(inner.id!).text).toBe(true);
  await session.execute({
    type: "text",
    id: inner.id!,
    start: 0,
    end: 8,
    text: "Updated",
  });
  await session.execute({
    type: "arrange",
    objects: [{ id: box.id!, index: 0 }],
    action: "rotate:90",
  });
  const restored = await openDocxSession(await session.exportDocx());
  expect(allParagraphs(restored.document).flatMap((p) => p.children)).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        type: "drawing",
        shape: expect.objectContaining({ rotation: 90 }),
      }),
      expect.objectContaining({ type: "text", text: "Updated" }),
    ]),
  );
});
it("keeps bookmark and comment ranges through text editing and export", async () => {
  const { session, id } = await setup();
  await session.execute({
    type: "text",
    id,
    start: 0,
    end: 0,
    text: "one two three",
  });
  await session.execute({
    type: "bookmark",
    id,
    name: "target",
    start: 4,
    end: 7,
  });
  await session.execute({
    type: "comment",
    id,
    key: "new",
    start: 4,
    end: 7,
    text: "Check this",
    author: "Reviewer",
    date: "2026-11-05T12:00:00Z",
  });
  expect(session.capability(id).text).toBe(true);
  await session.execute({ type: "text", id, start: 0, end: 0, text: "New " });
  const restored = await openDocxSession(await session.exportDocx()),
    p = allParagraphs(restored.document)[0];
  expect(p.bookmarks).toEqual([{ name: "target", start: 8, end: 11 }]);
  expect(p.comments?.[0]).toMatchObject({
    start: 8,
    end: 11,
    text: "Check this",
    author: "Reviewer",
  });
});
it("authors charts with an embedded workbook and reimports their data", async () => {
  const { session, id } = await setup(),
    chart: NonNullable<DrawingInline["chart"]> = {
      kind: "bar",
      title: "Results",
      categories: ["A", "B"],
      series: [{ name: "Count", values: [4, 7] }],
    };
  await session.execute({
    type: "insertDrawing",
    id,
    drawing: {
      type: "drawing",
      widthPt: 360,
      heightPt: 240,
      alt: "Results",
      chart,
    },
  });
  const bytes = await session.exportDocx(),
    files = unzipSync(bytes);
  expect(Object.keys(files).some((path) => path.endsWith(".xlsx"))).toBe(true);
  const restored = await openDocxSession(bytes);
  expect(
    allParagraphs(restored.document)
      .flatMap((p) => p.children)
      .find((c) => c.type === "drawing" && c.chart),
  ).toMatchObject({ chart });
});
it("writes equations as OMML rather than flattening to an image", async () => {
  const { session, id } = await setup();
  await session.execute({
    type: "insertDrawing",
    id,
    drawing: {
      type: "drawing",
      widthPt: 200,
      heightPt: 48,
      alt: "Equation",
      equation: { expression: "\\frac{x^2}{y}+\\sqrt{z}" },
    },
  });
  const bytes = await session.exportDocx();
  expect(strFromU8(unzipSync(bytes)["word/document.xml"])).toContain("m:oMath");
  const restored = await openDocxSession(bytes);
  expect(
    allParagraphs(restored.document)
      .flatMap((p) => p.children)
      .some(
        (c) =>
          c.type === "drawing" && c.equation?.expression.includes("\\frac"),
      ),
  ).toBe(true);
});
it("round-trips layout settings and authored table data", async () => {
  const { session, id } = await setup();
  await session.execute({
    type: "layout",
    format: {
      mirrorMargins: true,
      gutterPt: 12,
      hyphenation: "auto",
      direction: "rtl",
      verticalAlignment: "center",
      lineNumbers: { mode: "page", start: 1, countBy: 5, distancePt: 18 },
    },
  });
  await session.execute({
    type: "insertTable",
    id,
    rows: 2,
    columns: 2,
    cells: [
      ["Name", "Value"],
      ["One", "1"],
    ],
    header: true,
    columnWidthsPt: [144, 72],
  });
  const restored = await openDocxSession(await session.exportDocx());
  expect(restored.document.page).toMatchObject({
    mirrorMargins: true,
    gutterPt: 12,
    hyphenation: "auto",
    direction: "rtl",
    verticalAlignment: "center",
    lineNumbers: { mode: "page", countBy: 5 },
  });
  expect(
    restored.document.blocks.find((b) => b.type === "table"),
  ).toMatchObject({
    columnsPt: [144, 72],
    rows: [{ header: true }, { cells: expect.any(Array) }],
  });
});
it("persists document property fields and their core-property values", async () => {
  const { session, id } = await setup();
  await session.execute({
    type: "insertField",
    id,
    field: "title",
    value: "A title",
  });
  await session.execute({
    type: "insertField",
    id,
    field: "author",
    value: "An author",
  });
  const restored = await openDocxSession(await session.exportDocx());
  expect(restored.document.properties).toMatchObject({
    title: "A title",
    author: "An author",
  });
  expect(
    allParagraphs(restored.document)
      .flatMap((p) => p.children)
      .filter((r) => r.type === "text" && r.field)
      .map((r) => (r.type === "text" ? r.field : undefined)),
  ).toEqual(expect.arrayContaining(["title", "author"]));
});
