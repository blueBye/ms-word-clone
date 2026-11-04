import { expect, it } from "vitest";
import {
  createBlankDocx,
  openDocxSession,
  importDocx,
  getParagraphs,
} from "../src/index.js";

it("defers OOXML work for text/style edits, then preserves export, undo, redo and edits after export", async () => {
  const session = await openDocxSession(createBlankDocx());
  const id = getParagraphs(session.document.blocks)[0].id!;
  await session.execute({ type: "text", id, start: 0, end: 0, text: "first" });
  await session.execute({
    type: "style",
    styleId: "Normal",
    run: { color: "#225588" },
  });
  expect(session.compatibilityStatistics).toEqual({
    packageParses: 1,
    partWrites: 0,
    queuedCommands: 2,
  });
  const exported = await session.exportDocx();
  expect(session.compatibilityStatistics).toEqual({
    packageParses: 1,
    partWrites: 2,
    queuedCommands: 0,
  });
  expect(
    getParagraphs((await importDocx(exported)).blocks)[0].children[1],
  ).toMatchObject({ text: "first", properties: { color: "#225588" } });
  await session.undo();
  expect(
    getParagraphs(session.document.blocks)[0].children.at(-1),
  ).not.toMatchObject({ properties: { color: "#225588" } });
  await session.redo();
  await session.execute({
    type: "text",
    id,
    start: 5,
    end: 5,
    text: " second",
  });
  const next = await importDocx(await session.exportDocx());
  expect(
    getParagraphs(next.blocks)[0]
      .children.filter((c) => c.type === "text")
      .map((c) => c.text)
      .join(""),
  ).toBe("first second");
});

it("reconciles structural compatibility operations without losing pending text, style or history", async () => {
  const s = await openDocxSession(createBlankDocx());
  const id = getParagraphs(s.document.blocks)[0].id!;
  await s.executeBatch([
    { type: "text", id, start: 0, end: 0, text: "hello" },
    { type: "run", id, start: 0, end: 5, format: { bold: true } },
  ]);
  await s.execute({ type: "splitParagraph", id, start: 2, end: 2 });
  expect(s.compatibilityStatistics.packageParses).toBe(2);
  expect(
    getParagraphs(s.document.blocks).map((p) =>
      p.children
        .filter((c) => c.type === "text")
        .map((c) => c.text)
        .join(""),
    ),
  ).toEqual(["he", "llo"]);
  await s.undo();
  expect(getParagraphs(s.document.blocks)).toHaveLength(1);
  await s.undo();
  expect(s.dirty).toBe(false);
  expect(await s.exportDocx()).toEqual(createBlankDocx());
});

it("keeps Automatic color and cleared highlighting consistent between core state and OOXML reimport", async () => {
  const s = await openDocxSession(createBlankDocx()),
    id = getParagraphs(s.document.blocks)[0].id!;
  await s.execute({ type: "text", id, start: 0, end: 0, text: "text" });
  await s.execute({
    type: "style",
    styleId: "Normal",
    run: { color: "#CC2200" },
  });
  await s.execute({ type: "run", id, format: { highlight: "yellow" } });
  await s.execute({
    type: "run",
    id,
    format: { color: "auto", highlight: "none" },
  });
  expect(getParagraphs(s.document.blocks)[0].children.at(-1)).toMatchObject({
    properties: { color: "#000000", backgroundColor: "transparent" },
  });
  expect(
    getParagraphs(
      (await importDocx(await s.exportDocx())).blocks,
    )[0].children.at(-1),
  ).toMatchObject({
    properties: { color: "#000000", backgroundColor: "transparent" },
  });
});
