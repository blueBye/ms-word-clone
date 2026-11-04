import { expect, it } from "vitest";
import { unzipSync, strFromU8 } from "fflate";
import {
  createBlankDocx,
  openDocxSession,
  getParagraphs,
  importDocx,
} from "../src/index.js";
import { styleFormCommand } from "../apps/web/lib/editor.js";
import { docx, documentXml, W, OFFICE_REL, REL } from "./helpers/docx.js";

it("keeps new text and Enter-created paragraphs linked to Normal, with independent direct overrides", async () => {
  const s = await openDocxSession(createBlankDocx());
  let p = getParagraphs(s.document.blocks);
  const first = p[0].id!;
  expect(p[0].properties.styleId).toBe("Normal");
  await s.execute({
    type: "text",
    id: first,
    start: 0,
    end: 0,
    text: "English فارسی",
  });
  await s.execute({ type: "splitParagraph", id: first, start: 13, end: 13 });
  const second = s.selection!.id;
  await s.execute({
    type: "text",
    id: second,
    start: 0,
    end: 0,
    text: "second",
  });
  await s.execute({
    type: "run",
    id: second,
    start: 0,
    end: 6,
    format: { color: "#CC2200" },
  });
  await s.execute({
    type: "style",
    styleId: "Normal",
    run: { fontFamily: "Tahoma", fontSizePt: 18, color: "#225588" },
  });
  p = getParagraphs(s.document.blocks);
  expect(p.map((p) => p.properties.styleId)).toEqual(["Normal", "Normal"]);
  expect(p[0].children[0]).toMatchObject({
    properties: {
      fontFamily: "Tahoma",
      fontSizePt: 18,
      color: "#225588",
      complexScript: { fontFamily: "Tahoma", fontSizePt: 18 },
    },
  });
  expect(p[1].children[0]).toMatchObject({
    properties: { fontSizePt: 18, color: "#CC2200" },
  });
  await s.execute({
    type: "style",
    styleId: "Normal",
    run: { fontSizePt: 22 },
  });
  const reimported = getParagraphs(
    (await importDocx(await s.exportDocx())).blocks,
  );
  expect(
    reimported.every(
      (p) =>
        p.children[0].type === "text" &&
        p.children[0].properties.fontSizePt === 22,
    ),
  ).toBe(true);
  const archive = unzipSync(await s.exportDocx());
  expect(strFromU8(archive["word/document.xml"])).not.toContain("<w:sz");
  await s.undo();
  expect(getParagraphs(s.document.blocks)[0].children[0]).toMatchObject({
    properties: { fontSizePt: 18 },
  });
});

it("resolves implicit styles in tables and headers and respects a document's custom default", async () => {
  const bytes = docx(
    documentXml(
      "<w:p><w:r><w:t>body</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
    ),
    {
      "custom/_rels/main.xml.rels": `<Relationships xmlns="${REL}"><Relationship Id="s" Type="${OFFICE_REL}/styles" Target="styles.xml"/></Relationships>`,
      "custom/styles.xml": `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="BodyDefault"><w:name w:val="Body default"/><w:rPr><w:sz w:val="26"/></w:rPr></w:style></w:styles>`,
    },
  );
  const s = await openDocxSession(bytes);
  expect(
    getParagraphs(s.document.blocks).every(
      (p) => p.properties.styleId === "BodyDefault",
    ),
  ).toBe(true);
  expect(await s.exportDocx()).toEqual(bytes);
  await s.execute({
    type: "style",
    styleId: "BodyDefault",
    run: { fontSizePt: 20 },
  });
  expect(
    getParagraphs(s.document.blocks).every(
      (p) =>
        p.children[0].type === "text" &&
        p.children[0].properties.fontSizePt === 20,
    ),
  ).toBe(true);
});

it("treats Normal as the fallback in documents that omit its default flag", async () => {
  const s = await openDocxSession(
    docx(documentXml("<w:p><w:r><w:t>body</w:t></w:r></w:p>"), {
      "custom/_rels/main.xml.rels": `<Relationships xmlns="${REL}"><Relationship Id="s" Type="${OFFICE_REL}/styles" Target="styles.xml"/></Relationships>`,
      "custom/styles.xml": `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="30"/></w:rPr></w:style></w:styles>`,
    }),
  );
  expect(getParagraphs(s.document.blocks)[0].children[0]).toMatchObject({
    properties: { fontSizePt: 15 },
  });
  await s.execute({
    type: "style",
    styleId: "Normal",
    run: { fontSizePt: 19 },
  });
  expect(getParagraphs(s.document.blocks)[0].children[0]).toMatchObject({
    properties: { fontSizePt: 19 },
  });
});

it("modifies only edited style fields so based-on inheritance continues through later updates", async () => {
  const s = await openDocxSession(createBlankDocx()),
    id = getParagraphs(s.document.blocks)[0].id!;
  await s.execute({ type: "paragraph", id, format: { styleId: "NoSpacing" } });
  const style = (await s.styles()).find((s) => s.id === "NoSpacing")!;
  const data = new FormData();
  for (const [key, value] of Object.entries({
    name: style.name,
    basedOn: "Normal",
    gallery: "on",
    font: "Calibri",
    size: "12",
    color: "#225588",
    align: "left",
    before: "0",
    after: "0",
  }))
    data.set(key, value);
  const command = styleFormCommand(style, data);
  expect(command).toMatchObject({ run: { color: "#225588" } });
  expect(command).not.toHaveProperty("paragraph");
  await s.execute(command);
  await s.execute({
    type: "style",
    styleId: "Normal",
    run: { fontFamily: "Arial", fontSizePt: 21, italic: true },
  });
  expect(getParagraphs(s.document.blocks)[0].children[0]).toMatchObject({
    properties: {
      color: "#225588",
      fontFamily: "Arial",
      fontSizePt: 21,
      italic: true,
    },
  });
});

it("lets the source of Update to Match Selection follow subsequent changes too", async () => {
  const s = await openDocxSession(createBlankDocx()),
    id = getParagraphs(s.document.blocks)[0].id!;
  await s.execute({ type: "text", id, start: 0, end: 0, text: "source" });
  await s.execute({
    type: "run",
    id,
    format: { fontSizePt: 20, color: "#225588" },
  });
  await s.executeBatch([
    {
      type: "style",
      styleId: "Normal",
      run: { fontSizePt: 20, color: "#225588" },
    },
    { type: "paragraph", id, format: { styleId: "Normal" } },
  ]);
  await s.execute({
    type: "style",
    styleId: "Normal",
    run: { fontSizePt: 24, color: "#113366" },
  });
  expect(getParagraphs(s.document.blocks)[0].children[0]).toMatchObject({
    properties: { fontSizePt: 24, color: "#113366" },
  });
});
