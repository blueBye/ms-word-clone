import { it, expect } from "vitest";
import { unzipSync, strFromU8 } from "fflate";
import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { getParagraphs, openDocxSession, importDocx } from "../src/index.js";
import { DocumentRenderer } from "../src/react/index.js";
import { docx, documentXml } from "./helpers/docx.js";

const input = () =>
  docx(
    documentXml(
      ["one", "two", "three", "four"]
        .map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
        .join(""),
    ),
  );
it("changes bullet glyphs, retains list identity on indent, restarts and continues without altering the shared definition", async () => {
  const s = await openDocxSession(input()),
    ids = getParagraphs(s.document.blocks).map((p) => p.id!);
  await s.executeBatch(
    ids.slice(0, 3).map((id) => ({
      type: "list" as const,
      id,
      kind: "bullet" as const,
      bullet: "•",
    })),
  );
  let p = getParagraphs(s.document.blocks);
  expect(p.slice(0, 3).map((p) => p.properties.list?.marker)).toEqual([
    "•",
    "•",
    "•",
  ]);
  expect(
    new Set(p.slice(0, 3).map((p) => p.properties.numbering?.id)).size,
  ).toBe(1);
  await s.executeBatch(
    ids.slice(0, 3).map((id) => ({
      type: "list" as const,
      id,
      kind: "bullet" as const,
      bullet: "▪",
    })),
  );
  p = getParagraphs(s.document.blocks);
  expect(p.slice(0, 3).map((p) => p.properties.list?.marker)).toEqual([
    "▪",
    "▪",
    "▪",
  ]);
  const instance = p[1].properties.numbering?.id;
  await s.execute({
    type: "list",
    id: ids[1],
    kind: "bullet",
    level: 1,
    preserveDefinition: true,
  });
  p = getParagraphs(s.document.blocks);
  expect(p[1].properties.numbering).toEqual({ id: instance, level: 1 });
  expect(p[1].properties.list?.marker).toBe("▪");
  await s.executeBatch(
    ids.slice(0, 3).map((id) => ({
      type: "list" as const,
      id,
      kind: "decimal" as const,
      level: 0,
      pattern: "%1)",
    })),
  );
  p = getParagraphs(s.document.blocks);
  expect(p.slice(0, 3).map((p) => p.properties.list?.marker)).toEqual([
    "1)",
    "2)",
    "3)",
  ]);
  await s.execute({
    type: "list",
    id: ids[2],
    kind: "decimal",
    restart: true,
    start: 7,
  });
  p = getParagraphs(s.document.blocks);
  expect(p.slice(0, 3).map((p) => p.properties.list?.marker)).toEqual([
    "1)",
    "2)",
    "7)",
  ]);
  await s.execute({
    type: "list",
    id: ids[2],
    kind: "decimal",
    continue: true,
  });
  expect(getParagraphs(s.document.blocks)[2].properties.list?.marker).toBe(
    "3)",
  );
  const imported = await importDocx(await s.exportDocx());
  expect(
    getParagraphs(imported.blocks)
      .slice(0, 3)
      .map((p) => p.properties.list?.marker),
  ).toEqual(["1)", "2)", "3)"]);
});
it("keeps hierarchical numbering through level changes and ends an empty list item on Enter", async () => {
  const s = await openDocxSession(input()),
    ids = getParagraphs(s.document.blocks).map((p) => p.id!);
  await s.executeBatch(
    ids.slice(0, 3).map((id) => ({
      type: "list" as const,
      id,
      kind: "decimal" as const,
      multilevel: "decimal" as const,
    })),
  );
  await s.execute({
    type: "list",
    id: ids[1],
    kind: "decimal",
    level: 1,
    preserveDefinition: true,
  });
  const p = getParagraphs(s.document.blocks);
  expect(p[1].properties.list?.marker).toBe("1.1.");
  expect(p[2].properties.list?.marker).toBe("2.");
  await s.execute({ type: "splitParagraph", id: ids[2], start: 5, end: 5 });
  const empty = s.selection!.id;
  expect(
    getParagraphs(s.document.blocks).find((p) => p.id === empty)?.properties
      .list,
  ).toBeDefined();
  await s.execute({ type: "splitParagraph", id: empty, start: 0, end: 0 });
  expect(getParagraphs(s.document.blocks)).toHaveLength(5);
  expect(
    getParagraphs(s.document.blocks).find((p) => p.id === empty)?.properties
      .list,
  ).toBeUndefined();
});
it("updates a style definition for every instance, renames and hides it, preserves overrides and supports undo", async () => {
  const s = await openDocxSession(input()),
    ids = getParagraphs(s.document.blocks).map((p) => p.id!);
  await s.executeBatch(
    ids.slice(0, 2).map((id) => ({
      type: "paragraph" as const,
      id,
      format: { styleId: "Heading1" },
    })),
  );
  await s.execute({ type: "run", id: ids[1], format: { color: "#ff0000" } });
  await s.execute({
    type: "style",
    styleId: "Heading1",
    name: "Project heading",
    run: {
      fontFamily: "Tahoma",
      fontSizePt: 24,
      bold: false,
      color: "#225588",
    },
    paragraph: { spaceAfterPt: 14 },
  });
  let p = getParagraphs(s.document.blocks);
  expect(p[0].children[0]).toMatchObject({
    properties: { fontFamily: "Tahoma", fontSizePt: 24, color: "#225588" },
  });
  expect(
    p[0].children[0].type === "text" && p[0].children[0].properties.bold,
  ).toBeFalsy();
  expect(p[1].children[0]).toMatchObject({
    properties: { color: "#FF0000", fontSizePt: 24 },
  });
  expect(p[0].properties.spaceAfterPt).toBe(14);
  expect((await s.styles()).find((s) => s.id === "Heading1")).toMatchObject({
    name: "Project heading",
    run: { fontSizePt: 24 },
  });
  await s.execute({ type: "style", styleId: "Heading1", gallery: false });
  expect((await s.styles()).find((s) => s.id === "Heading1")?.gallery).toBe(
    false,
  );
  await s.undo();
  expect((await s.styles()).find((s) => s.id === "Heading1")?.gallery).toBe(
    true,
  );
  const reopened = await openDocxSession(await s.exportDocx());
  expect((await reopened.styles()).find((s) => s.id === "Heading1")?.name).toBe(
    "Project heading",
  );
  await expect(
    s.execute({ type: "style", styleId: "Normal", basedOn: "Heading1" }),
  ).rejects.toThrow("cycle");
});
it("exports qualified Office text effects, reimports typography, and clears them with undo support", async () => {
  const s = await openDocxSession(input()),
    id = getParagraphs(s.document.blocks)[0].id!;
  const format = {
    textOutline: { widthPt: 1, color: "#225588" },
    textShadow: {
      color: "#000000",
      blurPt: 3,
      distancePt: 2,
      angle: 45,
      opacity: 0.4,
    },
    glow: { color: "#00AAFF", radiusPt: 5, opacity: 0.6 },
    reflection: { distancePt: 2, opacity: 0.5, size: 0.6 },
    ligatures: "standardContextual",
    numberForm: "oldStyle" as const,
    numberSpacing: "tabular" as const,
    stylisticSets: [1, 3],
  };
  await s.execute({ type: "run", id, format });
  const output = await s.exportDocx(),
    p = getParagraphs((await importDocx(output)).blocks)[0];
  expect(p.children[0]).toMatchObject({ properties: format });
  const archive = unzipSync(output),
    xml = strFromU8(archive["custom/main.xml"]);
  expect(xml).toContain('mc:Ignorable="w14"');
  expect(xml).toContain('w14:blurRad="38100"');
  expect(xml).toContain('w14:val="standardContextual"');
  expect(xml).toContain("<w14:textFill>");
  expect(xml).not.toContain('w14:alpha w14:val="100000"');
  await s.execute({ type: "run", id, format: { fontSizePt: 20, bold: true } });
  const updatedXml = strFromU8(
    unzipSync(await s.exportDocx())["custom/main.xml"],
  );
  expect(updatedXml.indexOf("<w:sz ")).toBeLessThan(
    updatedXml.indexOf("<w14:glow"),
  );
  expect(updatedXml.indexOf("<w:b ")).toBeLessThan(
    updatedXml.indexOf("<w14:glow"),
  );
  const html = renderToStaticMarkup(
    React.createElement(DocumentRenderer, { document: s.document }),
  );
  expect(html).toContain("-webkit-text-stroke:1pt #225588");
  expect(html).toContain("text-shadow:");
  await s.execute({ type: "clearRun", id });
  expect(getParagraphs(s.document.blocks)[0].children[0]).not.toMatchObject({
    properties: { textOutline: format.textOutline },
  });
  await s.undo();
  expect(getParagraphs(s.document.blocks)[0].children[0]).toMatchObject({
    properties: format,
  });
});

it("retains each run color when adding Word effect fills across a colored selection", async () => {
  const s = await openDocxSession(
    docx(
      documentXml(
        '<w:p><w:r><w:rPr><w:color w:val="225588"/></w:rPr><w:t>blue</w:t></w:r><w:r><w:rPr><w:color w:val="CC2222"/></w:rPr><w:t>red</w:t></w:r></w:p>',
      ),
    ),
  );
  const id = getParagraphs(s.document.blocks)[0].id!;
  await s.execute({
    type: "run",
    id,
    format: { textOutline: { color: "#000000", widthPt: 0.5 } },
  });
  const p = getParagraphs(s.document.blocks)[0];
  expect(p.children[0]).toMatchObject({
    text: "blue",
    properties: { color: "#225588" },
  });
  expect(p.children[1]).toMatchObject({
    text: "red",
    properties: { color: "#CC2222" },
  });
});
it("renders strikethrough with a solid line independently of a wavy underline", async () => {
  const s = await openDocxSession(input()),
    id = getParagraphs(s.document.blocks)[0].id!;
  await s.execute({
    type: "run",
    id,
    format: { underline: "wave", strike: true },
  });
  const html = renderToStaticMarkup(
    React.createElement(DocumentRenderer, { document: s.document }),
  );
  expect(html).toContain("text-decoration-style:wavy");
  expect(html).toContain(
    '<span style="text-decoration:line-through solid">one</span>',
  );
  expect(html).not.toContain("underline line-through");
});

it("applying a paragraph style replaces uniform font overrides and retains distinct inline emphasis", async () => {
  const s = await openDocxSession(
    docx(
      documentXml(
        '<w:p><w:pPr><w:spacing w:after="400"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/><w:i/></w:rPr><w:t>emphasis</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr><w:t> plain</w:t></w:r></w:p>',
      ),
    ),
  );
  const id = getParagraphs(s.document.blocks)[0].id!;
  await s.execute({ type: "paragraph", id, format: { styleId: "NoSpacing" } });
  const p = getParagraphs(s.document.blocks)[0];
  expect(p.properties.spaceAfterPt).toBe(0);
  expect(p.children[0]).toMatchObject({
    properties: { fontSizePt: 12, italic: true },
  });
  expect(p.children[1]).toMatchObject({ properties: { fontSizePt: 12 } });
  expect(
    p.children[1].type === "text" && p.children[1].properties.italic,
  ).toBeFalsy();
});
