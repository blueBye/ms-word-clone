import { expect, it } from "vitest";
import { readFile } from "node:fs/promises";
import { unzipSync, strFromU8 } from "fflate";
import {
  createBlankDocx,
  getParagraphs,
  importDocx,
  openDocxSession,
} from "../src/index.js";
import { docx, documentXml, W, REL, OFFICE_REL } from "./helpers/docx.js";
import { physicalAlignment } from "../src/core/format.js";

it("updates Title from a recolored substring at a run boundary and keeps later edits linked", async () => {
  const s = await openDocxSession(createBlankDocx());
  const first = getParagraphs(s.document.blocks)[0].id!;
  await s.execute({
    type: "text",
    id: first,
    start: 0,
    end: 0,
    text: "My title",
  });
  await s.execute({
    type: "paragraph",
    id: first,
    format: { styleId: "Title" },
  });
  await s.execute({ type: "insertParagraph", id: first, text: "Other title" });
  const other = getParagraphs(s.document.blocks)[1].id!;
  await s.execute({
    type: "paragraph",
    id: other,
    format: { styleId: "Title" },
  });
  await s.execute({
    type: "run",
    id: first,
    start: 3,
    end: 8,
    format: { color: "#CC2200" },
  });
  await s.execute({
    type: "updateStyleFromSelection",
    id: first,
    styleId: "Title",
    start: 3,
    end: 8,
  });
  let p = getParagraphs(s.document.blocks);
  expect(
    p
      .flatMap((p) => p.children)
      .every((c) => c.type === "text" && c.properties.color === "#CC2200"),
  ).toBe(true);
  expect(
    p[0].children.find((c) => c.type === "text" && c.text === "title"),
  ).toMatchObject({ directProperties: {} });
  expect((await s.styles()).find((s) => s.id === "Title")?.run.color).toBe(
    "#CC2200",
  );
  await s.execute({
    type: "style",
    styleId: "Title",
    run: { color: "#225588" },
  });
  p = getParagraphs((await importDocx(await s.exportDocx())).blocks);
  expect(
    p
      .flatMap((p) => p.children)
      .every((c) => c.type === "text" && c.properties.color === "#225588"),
  ).toBe(true);
  await s.undo();
  await s.undo();
  p = getParagraphs(s.document.blocks);
  expect(
    p[0].children.find((c) => c.type === "text" && c.text === "title"),
  ).toMatchObject({ properties: { color: "#CC2200" } });
  expect(
    p[1].children[0].type === "text" && p[1].children[0].properties.color,
  ).not.toBe("#CC2200");
});

const mixedDocument = () =>
  docx(
    documentXml(
      [
        '<w:p><w:pPr><w:bidi/><w:jc w:val="right"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr><w:r><w:t>متن فارسی</w:t></w:r></w:p>',
        "<w:p><w:r><w:t>English body</w:t></w:r></w:p>",
        '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>English title</w:t></w:r></w:p>',
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
      ].join(""),
    ),
    {
      "custom/_rels/main.xml.rels": `<Relationships xmlns="${REL}"><Relationship Id="s" Type="${OFFICE_REL}/styles" Target="styles.xml"/></Relationships>`,
      "custom/styles.xml": `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Tahoma"/><w:sz w:val="24"/><w:szCs w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style></w:styles>`,
    },
  );

it("changes Normal's color from Persian text without changing Latin/CS fonts, sizes, direction, spacing or tables", async () => {
  const s = await openDocxSession(mixedDocument());
  let p = getParagraphs(s.document.blocks);
  const before = p.map((p) => ({
    properties: p.properties,
    run: p.children[0].type === "text" ? p.children[0].properties : {},
  }));
  const id = p[0].id!;
  const originalStyles = strFromU8(
    unzipSync(await s.exportDocx())["custom/styles.xml"],
  );
  await s.execute({
    type: "run",
    id,
    start: 0,
    end: 9,
    format: { color: "#225588" },
  });
  await s.execute({
    type: "updateStyleFromSelection",
    id,
    styleId: "Normal",
    start: 0,
    end: 9,
  });
  p = getParagraphs(s.document.blocks);
  p.forEach((paragraph, index) => {
    expect(paragraph.properties).toEqual(before[index].properties);
    for (const run of paragraph.children) {
      expect(run).toMatchObject({
        properties: { ...before[index].run, color: "#225588" },
      });
    }
  });
  const styles = strFromU8(
    unzipSync(await s.exportDocx())["custom/styles.xml"],
  );
  // A color update should add only color, not snapshot all the Normal properties.
  expect(styles.replace('<w:color w:val="225588"/>', "")).toBe(originalStyles);
});

it("does not pick the first color or inline emphasis from a mixed selection", async () => {
  const s = await openDocxSession(mixedDocument());
  const id = getParagraphs(s.document.blocks)[1].id!;
  await s.execute({
    type: "run",
    id,
    start: 0,
    end: 7,
    format: { color: "#CC2200", italic: true },
  });
  await s.execute({
    type: "updateStyleFromSelection",
    id,
    styleId: "Normal",
    start: 0,
    end: 12,
  });
  const normal = (await s.styles()).find((s) => s.id === "Normal")!;
  expect(normal.run.color).toBeUndefined();
  expect(normal.run.italic).toBeFalsy();
  expect(getParagraphs(s.document.blocks)[1].children[0]).toMatchObject({
    properties: { color: "#CC2200", italic: true },
  });
});

it("keeps an unchanged CS font, size and boldness when updating Latin formatting", async () => {
  const s = await openDocxSession(mixedDocument());
  const id = getParagraphs(s.document.blocks)[1].id!;
  // Direct Latin-only formatting from an imported producer.
  await s.execute({
    type: "run",
    id,
    start: 0,
    end: 7,
    format: {
      fontSizePt: 18,
      complexScriptSizePt: 14,
      bold: true,
      complexScriptBold: false,
    },
  });
  await s.execute({
    type: "updateStyleFromSelection",
    id,
    styleId: "Normal",
    start: 0,
    end: 7,
  });
  const normal = (await s.styles()).find((s) => s.id === "Normal")!.run;
  expect(normal).toMatchObject({
    fontSizePt: 18,
    bold: true,
    complexScript: { fontFamily: "Tahoma", fontSizePt: 14 },
  });
  expect(normal.complexScript?.bold).toBeFalsy();
});

it("does not spread an RTL paragraph's layout when the caret is in recolored text", async () => {
  const s = await openDocxSession(mixedDocument());
  const id = getParagraphs(s.document.blocks)[0].id!;
  await s.execute({
    type: "run",
    id,
    start: 4,
    end: 9,
    format: { color: "#CC2200" },
  });
  await s.execute({
    type: "updateStyleFromSelection",
    id,
    styleId: "Normal",
    start: 6,
    end: 6,
  });
  const normal = (await s.styles()).find((s) => s.id === "Normal")!;
  expect(normal.paragraph.direction).toBeUndefined();
  expect(normal.paragraph.alignment).toBeUndefined();
  expect(normal.paragraph.lineSpacing).toBeUndefined();
  expect(normal.run).toMatchObject({
    fontFamily: "Arial",
    fontSizePt: 12,
    color: "#CC2200",
    complexScript: { fontFamily: "Tahoma", fontSizePt: 14 },
  });
  await s.execute({
    type: "style",
    styleId: "Normal",
    run: { color: "#225588" },
  });
  expect(
    getParagraphs(s.document.blocks)[0].children.every(
      (c) => c.type === "text" && c.properties.color === "#225588",
    ),
  ).toBe(true);
});

it("keeps the shipped sample's English paragraphs LTR when Normal is updated from Persian text, including repeated updates", async () => {
  const bytes = await readFile(
    new URL("../apps/web/public/sample.docx", import.meta.url),
  );
  const s = await openDocxSession(bytes);
  const original = getParagraphs(s.document.blocks);
  const originalProperties = original.map((p) => p.properties);
  const id = original[2].id!;
  expect(original[2].properties.direction).toBe("rtl");
  expect(original[1].properties.direction).not.toBe("rtl");

  // Merely placing the caret in the Persian paragraph must not redefine Normal as RTL.
  await s.execute({
    type: "updateStyleFromSelection",
    id,
    styleId: "Normal",
    start: 1,
    end: 1,
  });
  expect(await s.exportDocx()).toEqual(new Uint8Array(bytes));

  await s.execute({
    type: "run",
    id,
    start: 0,
    end: 3,
    format: { color: "#CC2200" },
  });
  await s.execute({
    type: "updateStyleFromSelection",
    id,
    styleId: "Normal",
    start: 0,
    end: 3,
  });
  const afterColor = await s.exportDocx();
  // Both a caret and a text selection remain no-ops once the color matches.
  await s.execute({
    type: "updateStyleFromSelection",
    id,
    styleId: "Normal",
    start: 1,
    end: 1,
  });
  await s.execute({
    type: "updateStyleFromSelection",
    id,
    styleId: "Normal",
    start: 0,
    end: 3,
  });
  expect(await s.exportDocx()).toEqual(afterColor);
  expect(getParagraphs(s.document.blocks).map((p) => p.properties)).toEqual(
    originalProperties,
  );
  expect(
    getParagraphs((await importDocx(afterColor)).blocks).map((p) =>
      physicalAlignment(p.properties),
    ),
  ).toEqual(originalProperties.map(physicalAlignment));
  const normal = (await s.styles()).find((style) => style.id === "Normal")!;
  expect(normal.run.color).toBe("#CC2200");
  expect(normal.paragraph.direction).toBeUndefined();
  expect(normal.paragraph.alignment).toBeUndefined();
});
