import { expect, it } from "vitest";
import { unzipSync, strFromU8 } from "fflate";
import {
  openDocxSession,
  createBlankDocx,
  importDocx,
  getParagraphs,
} from "../src/index.js";
import { docx, documentXml, W, W_STRICT } from "./helpers/docx.js";

it("keeps no-op bytes exact, preserves opaque parts/XML and exports bilingual range edits with undo/redo", async () => {
  const input = docx(
    documentXml(
      '<w:p><w:pPr><x:future xmlns:x="urn:future" value="retained"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>سلام English</w:t></w:r></w:p>',
    ),
    { "custom/opaque.bin": new Uint8Array([0, 128, 255]) },
  );
  const session = await openDocxSession(input),
    id = getParagraphs(session.document.blocks)[0].id!;
  expect(await session.exportDocx()).toEqual(input);
  expect(session.canSetPage).toBe(true);
  await session.execute({
    type: "run",
    id,
    start: 5,
    end: 12,
    format: { bold: true, fontSizePt: 16 },
  });
  await session.execute({ type: "text", id, start: 0, end: 4, text: "درود" });
  await session.execute({
    type: "paragraph",
    id,
    format: { direction: "rtl", spaceAfterPt: 12 },
  });
  const output = await session.exportDocx(),
    entries = unzipSync(output);
  expect(entries["custom/opaque.bin"]).toEqual(new Uint8Array([0, 128, 255]));
  expect(strFromU8(entries["custom/main.xml"])).toContain('value="retained"');
  expect(entries["_rels/.rels"]).toEqual(unzipSync(input)["_rels/.rels"]);
  const p = getParagraphs((await importDocx(output)).blocks)[0];
  expect(
    p.children
      .filter((c) => c.type === "text")
      .map((c) => c.text)
      .join(""),
  ).toBe("درود English");
  expect(p.properties.direction).toBe("rtl");
  expect(p.properties.spaceAfterPt).toBe(12);
  expect(
    p.children
      .filter((c) => c.type === "text")
      .find((c) => c.text === "English")?.properties.bold,
  ).toBe(true);
  await session.undo();
  expect(
    getParagraphs(session.document.blocks)[0].properties.direction,
  ).not.toBe("rtl");
  await session.redo();
  expect(getParagraphs(session.document.blocks)[0].properties.direction).toBe(
    "rtl",
  );
  await session.undo();
  await session.undo();
  await session.undo();
  expect(session.dirty).toBe(false);
  expect(await session.exportDocx()).toEqual(input);
});

it("locks structured text and controlled content, but permits an owned paragraph alignment without dropping a field", async () => {
  const session = await openDocxSession(
    docx(
      documentXml(
        '<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p><w:sdt><w:sdtContent><w:p><w:r><w:t>Locked</w:t></w:r></w:p></w:sdtContent></w:sdt>',
      ),
    ),
  );
  const [field, control] = getParagraphs(session.document.blocks);
  expect(session.capability(field.id!).text).toBe(false);
  await expect(
    session.execute({
      type: "text",
      id: field.id!,
      start: 0,
      end: 1,
      text: "x",
    }),
  ).rejects.toThrow("structured");
  await session.execute({
    type: "paragraph",
    id: field.id!,
    format: { alignment: "center" },
  });
  expect(
    strFromU8(unzipSync(await session.exportDocx())["custom/main.xml"]),
  ).toContain("w:fldSimple");
  await expect(
    session.execute({
      type: "paragraph",
      id: control.id!,
      format: { alignment: "right" },
    }),
  ).rejects.toThrow("read-only");
});

it("creates and edits a blank document, inserts a table and preserves Strict namespaces", async () => {
  const session = await openDocxSession(createBlankDocx());
  const id = getParagraphs(session.document.blocks)[0].id!;
  await session.execute({
    type: "text",
    id,
    start: 0,
    end: 0,
    text: "One\tدو\nThree",
  });
  await session.execute({ type: "insertTable", id, rows: 2, columns: 3 });
  expect(session.document.blocks[1].type).toBe("table");
  await session.execute({
    type: "page",
    widthPt: 612,
    heightPt: 792,
    margins: { top: 36, bottom: 36, left: 36, right: 36 },
  });
  expect((await importDocx(await session.exportDocx())).page?.widthPt).toBe(
    612,
  );
  const strict = await openDocxSession(
    docx(documentXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>", W_STRICT)),
  );
  await strict.execute({
    type: "run",
    id: getParagraphs(strict.document.blocks)[0].id!,
    format: { italic: true },
  });
  const xml = strFromU8(
    unzipSync(await strict.exportDocx())["custom/main.xml"],
  );
  expect(xml).toContain(W_STRICT);
  expect(xml).not.toContain(`xmlns:w="${W}"`);
});

it("rejects partial graphemes and invalid ranges without modifying history", async () => {
  const session = await openDocxSession(
    docx(documentXml("<w:p><w:r><w:t>é 👩‍💻</w:t></w:r></w:p>")),
  );
  const id = getParagraphs(session.document.blocks)[0].id!;
  await expect(
    session.execute({ type: "text", id, start: 1, end: 2, text: "x" }),
  ).rejects.toThrow("complete characters");
  await expect(
    session.execute({
      type: "run",
      id,
      start: 0,
      end: 100,
      format: { bold: true },
    }),
  ).rejects.toThrow("Invalid text selection");
  expect(session.canUndo).toBe(false);
});

it("keeps paragraph identities and formatting stable through splitting, joining and undo", async () => {
  const session = await openDocxSession(
    docx(
      documentXml(
        "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>hello سلام</w:t></w:r></w:p><w:p><w:r><w:t>untouched</w:t></w:r></w:p>",
      ),
    ),
  );
  const [first, last] = getParagraphs(session.document.blocks),
    id = first.id!,
    lastId = last.id!;
  await session.execute({ type: "splitParagraph", id, start: 6, end: 6 });
  const p = getParagraphs(session.document.blocks);
  expect(
    p.map((p) =>
      p.children
        .filter((c) => c.type === "text")
        .map((c) => c.text)
        .join(""),
    ),
  ).toEqual(["hello ", "سلام", "untouched"]);
  expect(p[0].id).toBe(id);
  expect(p[2].id).toBe(lastId);
  expect(session.selection?.id).toBe(p[1].id);
  await session.execute({
    type: "joinParagraph",
    id: p[1].id!,
    direction: "backward",
  });
  expect(getParagraphs(session.document.blocks).map((p) => p.id)).toEqual([
    id,
    lastId,
  ]);
  expect(
    getParagraphs(session.document.blocks)[0]
      .children.filter((c) => c.type === "text")
      .every((c) => c.properties.bold),
  ).toBe(true);
  await session.undo();
  expect(getParagraphs(session.document.blocks)).toHaveLength(3);
  await session.undo();
  expect(getParagraphs(session.document.blocks).map((p) => p.id)).toEqual([
    id,
    lastId,
  ]);
});

it("creates relationship-backed lists and styles, exports all new parts and preserves the original archive on undo", async () => {
  const original = createBlankDocx(),
    session = await openDocxSession(original),
    id = getParagraphs(session.document.blocks)[0].id!;
  await session.execute({
    type: "text",
    id,
    start: 0,
    end: 0,
    text: "متن فارسی",
  });
  await session.execute({
    type: "paragraph",
    id,
    format: { direction: "rtl", styleId: "Heading1" },
  });
  await session.execute({ type: "list", id, kind: "decimal", start: 4 });
  expect(
    getParagraphs(session.document.blocks)[0].properties.list?.marker,
  ).toContain("۴");
  expect((await session.styles()).some((s) => s.id === "Heading1")).toBe(true);
  const exported = unzipSync(await session.exportDocx());
  expect(exported["word/numbering.xml"]).toBeDefined();
  expect(exported["word/styles.xml"]).toBeDefined();
  expect(strFromU8(exported["word/_rels/document.xml.rels"])).toContain(
    "/numbering",
  );
  expect(strFromU8(exported["[Content_Types].xml"])).toContain(
    "/word/numbering.xml",
  );
  await session.undo();
  await session.undo();
  await session.undo();
  expect(await session.exportDocx()).toEqual(original);
});

it("applies pending typing marks only to new text and clears direct formatting without dropping unknown properties", async () => {
  const session = await openDocxSession(
      docx(
        documentXml(
          '<w:p><w:r><w:rPr><w:i/><x:future xmlns:x="urn:future"/></w:rPr><w:t>old</w:t></w:r></w:p>',
        ),
      ),
    ),
    id = getParagraphs(session.document.blocks)[0].id!;
  await session.execute({
    type: "text",
    id,
    start: 3,
    end: 3,
    text: "new",
    format: { bold: true, color: "#123456" },
  });
  let runs = getParagraphs(session.document.blocks)[0].children.filter(
    (c) => c.type === "text",
  );
  expect(runs[0].properties.bold).not.toBe(true);
  expect(runs[1].properties.bold).toBe(true);
  await session.execute({ type: "clearRun", id, start: 3, end: 6 });
  runs = getParagraphs(session.document.blocks)[0].children.filter(
    (c) => c.type === "text",
  );
  expect(runs.at(-1)?.properties.bold).not.toBe(true);
  expect(
    strFromU8(unzipSync(await session.exportDocx())["custom/main.xml"]),
  ).toContain("x:future");
});

it("sorts consecutive paragraphs without changing their identities and rolls back a failed batch atomically", async () => {
  const session = await openDocxSession(
    docx(
      documentXml(
        "<w:p><w:r><w:t>z</w:t></w:r></w:p><w:p><w:r><w:t>a</w:t></w:r></w:p>",
      ),
    ),
  );
  const [z, a] = getParagraphs(session.document.blocks);
  await session.execute({ type: "sort", ids: [z.id!, a.id!] });
  expect(getParagraphs(session.document.blocks).map((p) => p.id)).toEqual([
    a.id,
    z.id,
  ]);
  const before = await session.exportDocx();
  await expect(
    session.executeBatch([
      { type: "run", id: z.id!, format: { bold: true } },
      { type: "text", id: a.id!, start: 5, end: 9, text: "bad" },
    ]),
  ).rejects.toThrow();
  expect(await session.exportDocx()).toEqual(before);
  await session.executeBatch([
    { type: "run", id: z.id!, format: { bold: true } },
    { type: "run", id: a.id!, format: { bold: true } },
  ]);
  await session.undo();
  expect(await session.exportDocx()).toEqual(before);
});
