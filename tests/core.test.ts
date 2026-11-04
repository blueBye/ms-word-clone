import { expect, it } from "vitest";
import {
  createDocumentEngine,
  allParagraphs,
  type EngineSnapshot,
} from "../src/core/index.js";

it("authors text and styles with no OOXML package, browser or React, and preserves direct overrides", async () => {
  const engine = createDocumentEngine();
  await engine.execute({
    type: "text",
    id: "p1",
    start: 0,
    end: 0,
    text: "Hello فارسی",
  });
  await engine.execute({
    type: "run",
    id: "p1",
    start: 6,
    end: 11,
    format: { color: "#CC2200" },
  });
  await engine.execute({
    type: "style",
    styleId: "Normal",
    run: { fontSizePt: 18 },
  });
  expect(allParagraphs(engine.document)[0].children.at(-1)).toMatchObject({
    text: "فارسی",
    properties: {
      color: "#CC2200",
      fontSizePt: 18,
      complexScript: { fontSizePt: 18 },
    },
  });
  expect(engine.statistics.words).toBe(2);
  await engine.undo();
  expect(allParagraphs(engine.document)[0].children.at(-1)).toMatchObject({
    properties: { fontSizePt: 12 },
  });
  await engine.redo();
  expect(allParagraphs(engine.document)[0].children.at(-1)).toMatchObject({
    properties: { fontSizePt: 18 },
  });
});

const manyParagraphs = (): EngineSnapshot => ({
  document: {
    type: "document",
    schemaVersion: 2,
    blocks: Array.from({ length: 1000 }, (_, i) => ({
      type: "paragraph",
      id: `p${i + 1}`,
      properties: { styleId: i < 10 ? "Heading" : "Normal" },
      directProperties: { styleId: i < 10 ? "Heading" : "Normal" },
      children: [
        {
          type: "text",
          text: `paragraph ${i}`,
          properties: {},
          directProperties: {},
        },
      ],
    })),
  },
  styles: {
    definitions: [
      {
        id: "Normal",
        name: "Normal",
        gallery: true,
        run: { fontSizePt: 12 },
        paragraph: {},
      },
      {
        id: "Heading",
        name: "Heading",
        basedOn: "Normal",
        gallery: true,
        run: { fontSizePt: 20 },
        paragraph: {},
      },
    ],
  },
});
it("uses d2ts dependency joins to invalidate only the affected style users and retains unrelated paragraph objects", async () => {
  const engine = createDocumentEngine(manyParagraphs());
  const unchanged = engine.getParagraph("p500");
  await engine.execute({
    type: "style",
    styleId: "Heading",
    run: { color: "#225588" },
  });
  expect(engine.statistics.resolvedParagraphs).toBe(10);
  expect(engine.getParagraph("p500")).toBe(unchanged);
  expect(engine.getParagraph("p1")?.children[0]).toMatchObject({
    properties: { color: "#225588" },
  });
  await engine.execute({
    type: "text",
    id: "p500",
    start: 0,
    end: 0,
    text: "new ",
  });
  expect(engine.statistics.resolvedParagraphs).toBe(1);
  expect(engine.statistics.words).toBe(2001);
  await engine.execute({
    type: "style",
    styleId: "Normal",
    run: { fontFamily: "Tahoma" },
  });
  expect(engine.statistics.resolvedParagraphs).toBe(1000);
});
it("rolls back an invalid batch atomically, rejects cycles and protects immutable document state", async () => {
  const engine = createDocumentEngine();
  const before = engine.document;
  await expect(
    engine.executeBatch([
      { type: "text", id: "p1", start: 0, end: 0, text: "first" },
      { type: "run", id: "p1", start: 0, end: 50, format: { bold: true } },
    ]),
  ).rejects.toThrow("Invalid text selection");
  expect(engine.document).toBe(before);
  expect(engine.canUndo).toBe(false);
  await engine.execute({
    type: "paragraph",
    id: "p1",
    format: { styleId: "Heading1" },
  });
  await expect(
    engine.execute({ type: "style", styleId: "Normal", basedOn: "Heading1" }),
  ).rejects.toThrow("cycle");
  expect(Object.isFrozen(engine.document.blocks)).toBe(true);
});

it("owns stable identities, structure, lists and page settings independently of Word", async () => {
  const engine = createDocumentEngine();
  await engine.execute({
    type: "text",
    id: "p1",
    start: 0,
    end: 0,
    text: "one",
  });
  await engine.execute({ type: "insertParagraph", id: "p1", text: "two" });
  const second = allParagraphs(engine.document)[1].id!;
  await engine.executeBatch([
    { type: "list", id: "p1", kind: "decimal" },
    { type: "list", id: second, kind: "decimal" },
  ]);
  expect(
    allParagraphs(engine.document).map((p) => p.properties.list?.marker),
  ).toEqual(["1.", "2."]);
  await engine.execute({
    type: "splitParagraph",
    id: second,
    start: 3,
    end: 3,
  });
  const empty = engine.selection!.id;
  expect(allParagraphs(engine.document)[2].properties.list?.marker).toBe("3.");
  await engine.execute({ type: "splitParagraph", id: empty, start: 0, end: 0 });
  expect(allParagraphs(engine.document)).toHaveLength(3);
  expect(engine.getParagraph(empty)?.properties.list).toBeUndefined();
  expect(engine.getParagraph(empty)?.properties.indentLeftPt).toBeUndefined();
  await engine.execute({ type: "insertTable", id: empty, rows: 2, columns: 2 });
  expect(engine.document.blocks.at(-1)?.type).toBe("table");
  expect(new Set(allParagraphs(engine.document).map((p) => p.id)).size).toBe(7);
  await engine.execute({ type: "page", widthPt: 612, heightPt: 792 });
  expect(engine.document.page).toMatchObject({ widthPt: 612, heightPt: 792 });
  const restored = createDocumentEngine(
    JSON.parse(JSON.stringify(engine.snapshot())),
  );
  expect(restored.document).toEqual(engine.document);
});

it("rewires inherited and character style dependencies after changing a base style", async () => {
  const snapshot = manyParagraphs();
  snapshot.styles.definitions.push({
    id: "Other",
    name: "Other",
    gallery: true,
    run: {},
    paragraph: {},
  });
  const engine = createDocumentEngine(snapshot);
  await engine.execute({ type: "style", styleId: "Heading", basedOn: "Other" });
  await engine.execute({
    type: "style",
    styleId: "Normal",
    run: { color: "#225588" },
  });
  expect(engine.statistics.resolvedParagraphs).toBe(990);
  expect(engine.getParagraph("p1")?.children[0]).not.toMatchObject({
    properties: { color: "#225588" },
  });
  const character = createDocumentEngine({
    document: {
      type: "document",
      schemaVersion: 2,
      blocks: [
        {
          type: "paragraph",
          id: "p1",
          properties: {},
          directProperties: {},
          children: [
            {
              type: "text",
              text: "styled",
              properties: {},
              directProperties: { styleId: "Emphasis" },
            },
          ],
        },
      ],
    },
    styles: {
      definitions: [
        {
          id: "Emphasis",
          kind: "character",
          name: "Emphasis",
          gallery: false,
          run: { italic: true },
          paragraph: {},
        },
      ],
    },
  });
  await character.execute({
    type: "style",
    styleId: "Emphasis",
    run: { color: "#CC2200" },
  });
  expect(character.getParagraph("p1")?.children[0]).toMatchObject({
    properties: { italic: true, color: "#CC2200" },
  });
});

it("retains cached search results across formatting edits and invalidates them on text edits", async () => {
  const engine = createDocumentEngine();
  await engine.execute({
    type: "text",
    id: "p1",
    start: 0,
    end: 0,
    text: "word word",
  });
  const matches = engine.find("word");
  expect(matches).toHaveLength(2);
  await engine.execute({
    type: "style",
    styleId: "Normal",
    run: { color: "#225588" },
  });
  expect(engine.find("word")).toBe(matches);
  await engine.execute({
    type: "text",
    id: "p1",
    start: 0,
    end: 4,
    text: "other",
  });
  expect(engine.find("word")).toHaveLength(1);
  const styles = await engine.styles();
  expect(Object.isFrozen(styles[0].run)).toBe(true);
});

it("keeps dirty state accurate when the oldest undo revisions have been pruned", async () => {
  const engine = createDocumentEngine(undefined, { historyLimit: 1 });
  await engine.execute({
    type: "text",
    id: "p1",
    start: 0,
    end: 0,
    text: "one",
  });
  await engine.execute({
    type: "text",
    id: "p1",
    start: 3,
    end: 3,
    text: " two",
  });
  await engine.undo();
  expect(engine.canUndo).toBe(false);
  expect(engine.dirty).toBe(true);
});
