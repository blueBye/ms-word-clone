import { expect, it } from "vitest";
import {
  createDocumentEngine,
  allParagraphs,
  similarFormatting,
  getParagraphs,
} from "../src/core/index.js";
import { createBlankDocx, openDocxSession, importDocx } from "../src/index.js";
import { rulerChange } from "../apps/web/lib/ruler.js";
import { groupColumns } from "../src/react/pagination.js";
import { textChange } from "../apps/web/lib/editor.js";

it("continues the caret's formatting through Enter and DOCX export", async () => {
  const session = await openDocxSession(createBlankDocx());
  const id = allParagraphs(session.document)[0].id!;
  await session.execute({
    type: "text",
    id,
    start: 0,
    end: 0,
    text: "red",
    format: { color: "#CC2200", fontFamily: "Tahoma", fontSizePt: 17 },
  });
  await session.execute({
    type: "splitParagraph",
    id,
    start: 3,
    end: 3,
    format: { bold: true },
  });
  const next = session.selection!.id;
  await session.execute({
    type: "text",
    id: next,
    start: 0,
    end: 0,
    text: "next",
  });
  const restored = await importDocx(await session.exportDocx());
  expect(
    getParagraphs(restored.blocks)[1].children.find(
      (r) => r.type === "text" && r.text === "next",
    ),
  ).toMatchObject({
    text: "next",
    properties: {
      color: "#CC2200",
      fontFamily: "Tahoma",
      fontSizePt: 17,
      bold: true,
    },
  });
});
it("inherits the preceding run when typing at a formatting boundary", async () => {
  const engine = createDocumentEngine();
  await engine.execute({
    type: "text",
    id: "p1",
    start: 0,
    end: 0,
    text: "redblue",
  });
  await engine.execute({
    type: "run",
    id: "p1",
    start: 0,
    end: 3,
    format: { color: "#CC2200" },
  });
  await engine.execute({
    type: "run",
    id: "p1",
    start: 3,
    end: 7,
    format: { color: "#0000FF" },
  });
  await engine.execute({ type: "text", id: "p1", start: 3, end: 3, text: "!" });
  expect(
    allParagraphs(engine.document)[0].children.find(
      (r) => r.type === "text" && r.text === "!",
    ),
  ).toMatchObject({ properties: { color: "#CC2200" } });
});
it("selects matching direct character formatting rather than every Normal paragraph", async () => {
  const engine = createDocumentEngine();
  await engine.execute({
    type: "text",
    id: "p1",
    start: 0,
    end: 0,
    text: "red plain red",
  });
  await engine.executeBatch([
    { type: "run", id: "p1", start: 0, end: 3, format: { color: "#CC2200" } },
    { type: "run", id: "p1", start: 10, end: 13, format: { color: "#CC2200" } },
  ]);
  expect(similarFormatting(allParagraphs(engine.document), "p1", 1)).toEqual([
    { id: "p1", start: 0, end: 3 },
    { id: "p1", start: 10, end: 13 },
  ]);
});
it("creates editable headers, page fields and removes story references with undo and round-trip", async () => {
  const session = await openDocxSession(createBlankDocx());
  await session.execute({ type: "story", kind: "header" });
  const p = session.document.headers!.default![0];
  if (p.type !== "paragraph") throw new Error("Expected a header paragraph");
  await session.execute({
    type: "text",
    id: p.id!,
    start: 0,
    end: 0,
    text: "My Header",
  });
  await session.execute({
    type: "story",
    kind: "footer",
    pageNumber: true,
    alignment: "center",
  });
  await session.execute({ type: "story", kind: "footer", appendText: true });
  const footer = allParagraphs({
    ...session.document,
    blocks: session.document.footers!.default!,
    headers: undefined,
    footers: undefined,
  });
  expect(footer).toHaveLength(2);
  expect(session.capability(footer[1].id!).text).toBe(true);
  const restored = await importDocx(await session.exportDocx());
  expect(getParagraphs(restored.headers!.default!)[0].children).toEqual(
    expect.arrayContaining([expect.objectContaining({ text: "My Header" })]),
  );
  expect(getParagraphs(restored.footers!.default!)[0].children).toEqual(
    expect.arrayContaining([expect.objectContaining({ field: "page" })]),
  );
  await session.execute({ type: "story", kind: "header", remove: true });
  expect(
    (await importDocx(await session.exportDocx())).headers?.default,
  ).toBeUndefined();
  await session.undo();
  expect(session.document.headers?.default).toBeDefined();
});
it("moves hanging indentation without moving the first line or list marker", () => {
  expect(rulerChange("hanging", 54, 36, -18, 450)).toEqual({
    indentStartPt: 54,
    firstLinePt: -36,
  });
  expect(rulerChange("start", 54, 36, -18, 450)).toEqual({ indentStartPt: 54 });
  expect(rulerChange("first", 24, 36, -18, 450)).toEqual({ firstLinePt: -12 });
});
it("keeps inserted hyperlinks editable and preserves the link through text edits and export", async () => {
  const session = await openDocxSession(createBlankDocx()),
    id = allParagraphs(session.document)[0].id!;
  await session.execute({
    type: "text",
    id,
    start: 0,
    end: 0,
    text: "Open site now",
  });
  await session.execute({
    type: "link",
    id,
    start: 5,
    end: 9,
    href: "https://example.com",
  });
  expect(session.capability(id).text).toBe(true);
  await session.execute({ type: "text", id, start: 9, end: 9, text: "!" });
  await session.execute({
    type: "run",
    id,
    start: 5,
    end: 10,
    format: { bold: true },
  });
  const restored = await openDocxSession(await session.exportDocx());
  const p = allParagraphs(restored.document)[0];
  expect(
    p.children
      .filter((r) => r.type === "text")
      .map((r) => r.text)
      .join(""),
  ).toBe("Open site! now");
  expect(p.children).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        text: "!",
        href: "https://example.com",
        properties: expect.objectContaining({ bold: true }),
      }),
    ]),
  );
  expect(restored.capability(p.id!).text).toBe(true);
});
it("embeds pictures in DOCX media parts and undoes their insertion", async () => {
  const session = await openDocxSession(createBlankDocx());
  const dataUrl =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a7E8AAAAASUVORK5CYII=";
  await session.execute({
    type: "insertImage",
    id: allParagraphs(session.document)[0].id!,
    dataUrl,
    widthPt: 72,
    heightPt: 72,
    alt: "Test picture",
  });
  const restored = await importDocx(await session.exportDocx());
  expect(getParagraphs(restored.blocks).flatMap((p) => p.children)).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        type: "drawing",
        widthPt: 72,
        heightPt: 72,
        alt: "Test picture",
        image: expect.objectContaining({ dataUrl }),
      }),
    ]),
  );
  expect(session.selection).toBeDefined();
  await session.undo();
  expect(
    allParagraphs(session.document)
      .flatMap((p) => p.children)
      .some((r) => r.type === "drawing"),
  ).toBe(false);
});
it("round-trips equal-width column layout and supports undo", async () => {
  const session = await openDocxSession(createBlankDocx());
  await session.execute({
    type: "columns",
    count: 2,
    gapPt: 24,
    separator: true,
  });
  expect((await importDocx(await session.exportDocx())).page?.columns).toEqual({
    count: 2,
    gapPt: 24,
    separator: true,
  });
  await session.undo();
  expect(session.document.page?.columns).toBeUndefined();
});
it("groups column slices without duplicating content and respects explicit page breaks", () => {
  const blocks = Array.from({ length: 5 }, (_, i) => ({
    type: "paragraph" as const,
    id: String(i),
    properties: { ...(i === 3 ? { pageBreakBefore: true } : {}) },
    children: [],
  }));
  const slices = blocks.map((b) => ({ blocks: [b], overflow: false }));
  const pages = groupColumns(slices, 2);
  expect(
    pages.map((p) =>
      p.columns?.map((column) =>
        column.map((p) => (p.type === "paragraph" ? p.id : "")),
      ),
    ),
  ).toEqual([[["0"], ["1"]], [["2"]], [["3"], ["4"]]]);
  expect(slices.map((s) => s.blocks.length)).toEqual([1, 1, 1, 1, 1]);
});
it("keeps repeated-character insertion at the caret instead of crossing a style boundary", () => {
  expect(textChange("aa", "aaa", 1)).toEqual({ start: 1, end: 1, text: "a" });
  expect(textChange("a💡a", "a💡💡a", 3)).toEqual({
    start: 3,
    end: 3,
    text: "💡",
  });
});
