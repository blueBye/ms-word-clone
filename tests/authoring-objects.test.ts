import { expect, it } from "vitest";
import { openDocxSession, createBlankDocx } from "../src/docx/index.js";
import {
  allParagraphs,
  createDocumentEngine,
  pageForParagraph,
  splitSections,
  type DrawingInline,
} from "../src/core/index.js";
import { unzipSync, strFromU8 } from "fflate";
import { sectionPagePositions } from "../src/react/pagination.js";
import { embeddedFile, unpackFile } from "../src/docx/ole-package.js";
import { parseXml, descendants } from "../src/docx/parser/xml.js";
const text = (p: ReturnType<typeof allParagraphs>[number]) =>
  p.children
    .filter((r) => r.type === "text")
    .map((r) => r.text)
    .join("");
const objects = (s: { document: Parameters<typeof allParagraphs>[0] }) =>
  allParagraphs(s.document).flatMap((p) =>
    p.children.flatMap((d, index) =>
      d.type === "drawing" ? [{ id: p.id!, index, d }] : [],
    ),
  );
const setup = async () => {
  const session = await openDocxSession(createBlankDocx());
  return { session, id: allParagraphs(session.document)[0].id! };
};
const shape = (name: string, x = 0): DrawingInline => ({
  type: "drawing",
  widthPt: 72,
  heightPt: 48,
  alt: name,
  shape: { kind: "rect", fill: "#4479AA" },
  position: {
    horizontal: { relativeTo: "paragraph", offsetPt: x },
    vertical: { relativeTo: "paragraph", offsetPt: 0 },
    behindText: false,
    zIndex: 2,
    wrap: "wrapSquare",
  },
  blocks: [
    {
      type: "paragraph",
      properties: {},
      children: [{ type: "text", text: name, properties: { fontSizePt: 12 } }],
    },
  ],
});
it("groups and ungroups native shapes without losing editable inner text", async () => {
  const { session, id } = await setup();
  await session.execute({ type: "insertDrawing", id, drawing: shape("First") });
  await session.execute({
    type: "insertDrawing",
    id,
    drawing: shape("Second", 100),
  });
  await session.execute({
    type: "arrange",
    objects: objects(session).map((o) => ({ id: o.id, index: o.index })),
    action: "group",
  });
  const group = objects(session).find((o) => o.d.children)!;
  expect(group.d.children).toHaveLength(2);
  expect(group.d.widthPt).toBe(172);
  await session.execute({
    type: "arrange",
    objects: [{ id: group.id, index: group.index }],
    action: "rotate:90",
  });
  expect(objects(session).find((o) => o.d.children)?.d.shape?.rotation).toBe(
    90,
  );
  const inner = allParagraphs(session.document).find(
    (p) => text(p) === "First",
  )!;
  await session.execute({
    type: "splitParagraph",
    id: inner.id!,
    start: 2,
    end: 2,
  });
  expect(allParagraphs(session.document).map(text)).toEqual(
    expect.arrayContaining(["Fi", "rst", "Second"]),
  );
  await session.execute({
    type: "arrange",
    objects: [{ id: group.id, index: group.index }],
    action: "ungroup",
  });
  const restored = await openDocxSession(await session.exportDocx());
  expect(objects(restored).filter((o) => o.d.shape)).toHaveLength(2);
  expect(
    objects(restored)
      .filter((o) => o.d.shape)
      .every((o) => o.d.shape?.rotation === 90),
  ).toBe(true);
  expect(allParagraphs(restored.document).map(text)).toEqual(
    expect.arrayContaining(["Fi", "rst", "Second"]),
  );
});
it("preserves embedded worksheet data and package attachments through export", async () => {
  const { session, id } = await setup();
  const cells = [
    ["Name", "Value"],
    ["One", "=1+2"],
  ];
  await session.execute({
    type: "insertDrawing",
    id,
    drawing: {
      type: "drawing",
      widthPt: 300,
      heightPt: 150,
      alt: "Worksheet",
      spreadsheet: { cells },
    },
  });
  await session.execute({
    type: "insertDrawing",
    id,
    drawing: {
      type: "drawing",
      widthPt: 180,
      heightPt: 60,
      alt: "note.txt",
      attachment: {
        name: "note.txt",
        dataUrl: "data:text/plain;base64,SGVsbG8=",
        mime: "text/plain",
      },
    },
  });
  const bytes = await session.exportDocx(),
    restored = await openDocxSession(bytes);
  const root = parseXml(
    strFromU8(unzipSync(bytes)["word/document.xml"]),
  ).documentElement!;
  const shapeIds = descendants(root, "shape").map((e) => e.getAttribute("id"));
  expect(new Set(shapeIds).size).toBe(shapeIds.length);
  expect(
    objects(restored).find((o) => o.d.spreadsheet)?.d.spreadsheet?.cells,
  ).toEqual(cells);
  expect(
    objects(restored).find((o) => o.d.attachment && !o.d.spreadsheet)?.d
      .attachment,
  ).toMatchObject({ dataUrl: "data:text/plain;base64,SGVsbG8=" });
  const sheet = objects(restored).find((o) => o.d.spreadsheet)!;
  await restored.execute({
    type: "drawing",
    id: sheet.id,
    index: sheet.index,
    drawing: { ...sheet.d, spreadsheet: { cells: [["Updated", "9"]] } },
  });
  expect(
    objects(await openDocxSession(await restored.exportDocx())).find(
      (o) => o.d.spreadsheet,
    )?.d.spreadsheet?.cells,
  ).toEqual([["Updated", "9"]]);
});
it("round-trips self-contained GLB model relationships and video metadata", async () => {
  const json = new TextEncoder().encode(
      JSON.stringify({
        asset: { version: "2.0" },
        scenes: [{ nodes: [] }],
        scene: 0,
      }),
    ),
    length = Math.ceil(json.length / 4) * 4,
    bytes = new Uint8Array(20 + length),
    view = new DataView(bytes.buffer);
  view.setUint32(0, 0x46546c67, true);
  view.setUint32(4, 2, true);
  view.setUint32(8, bytes.length, true);
  view.setUint32(12, length, true);
  view.setUint32(16, 0x4e4f534a, true);
  bytes.fill(32, 20);
  bytes.set(json, 20);
  const dataUrl = `data:model/gltf-binary;base64,${Buffer.from(bytes).toString("base64")}`,
    { session, id } = await setup();
  await session.execute({
    type: "insertDrawing",
    id,
    drawing: {
      type: "drawing",
      widthPt: 180,
      heightPt: 144,
      alt: "Model",
      model3d: { dataUrl },
    },
  });
  const output = await session.exportDocx(),
    files = unzipSync(output);
  expect(strFromU8(files["word/_rels/document.xml.rels"])).toContain(
    "2017/06/relationships/model3d",
  );
  expect(
    objects(await openDocxSession(output)).find((o) => o.d.model3d)?.d.model3d
      ?.dataUrl,
  ).toBe(dataUrl);
});
it("creates and removes a cover while keeping the document body and undo", async () => {
  const { session, id } = await setup();
  await session.execute({
    type: "text",
    id,
    start: 0,
    end: 0,
    text: "Original body",
  });
  await session.execute({
    type: "cover",
    title: "A report",
    subtitle: "Details",
    author: "Author",
    date: "2026",
    design: "banded",
  });
  expect(allParagraphs(session.document).map(text)).toContain("A report");
  await session.execute({
    type: "cover",
    title: "",
    subtitle: "",
    author: "",
    date: "",
    design: "simple",
    remove: true,
  });
  expect(allParagraphs(session.document).map(text)).toEqual(["Original body"]);
  await session.undo();
  expect(allParagraphs(session.document).map(text)).toContain("A report");
});
it("scopes page geometry, columns and headers to the selected section in core and OOXML", async () => {
  for (const session of [createDocumentEngine(), (await setup()).session]) {
    const first = allParagraphs(session.document)[0].id!;
    await session.execute({
      type: "text",
      id: first,
      start: 0,
      end: 0,
      text: "BeforeAfter",
    });
    await session.execute({
      type: "sectionBreak",
      id: first,
      start: 6,
      end: 6,
      kind: "nextPage",
    });
    const second = session.selection!.id;
    await session.execute({
      type: "page",
      id: second,
      widthPt: 792,
      heightPt: 612,
    });
    await session.execute({ type: "columns", id: first, count: 2, gapPt: 18 });
    expect(
      pageForParagraph(session.document, first).widthPt ?? 595.3,
    ).toBeLessThan(700);
    expect(pageForParagraph(session.document, second).widthPt).toBe(792);
    expect(pageForParagraph(session.document, first).columns?.count).toBe(2);
    expect(pageForParagraph(session.document, second).columns?.count ?? 1).toBe(
      1,
    );
    await session.execute({ type: "story", id: first, kind: "header" });
    const header = allParagraphs(session.document).find(
      (p) => p.id !== first && p.id !== second,
    )!;
    expect(header).toBeDefined();
    await session.execute({
      type: "text",
      id: header.id!,
      start: 0,
      end: 0,
      text: "First section",
    });
    expect(splitSections(session.document)[0].headers?.default).toBeDefined();
    if ("exportDocx" in session) {
      const restored = await openDocxSession(await session.exportDocx());
      const sections = splitSections(restored.document);
      expect(sections).toHaveLength(2);
      expect(sections[0].page?.columns?.count).toBe(2);
      expect(sections[1].page?.widthPt).toBe(792);
      expect(
        sections[0].headers?.default &&
          allParagraphs({
            ...sections[0],
            blocks: sections[0].headers.default,
            headers: undefined,
            footers: undefined,
          }).map(text),
      ).toContain("First section");
    }
    await session.undo();
    await session.undo();
    await session.undo();
    expect(pageForParagraph(session.document, first).columns?.count ?? 1).toBe(
      1,
    );
  }
});
it("does not repeat inserted page and column breaks on ordinary Enter", async () => {
  const { session, id } = await setup();
  await session.execute({ type: "text", id, start: 0, end: 0, text: "abcdef" });
  await session.execute({
    type: "insertBreak",
    id,
    start: 3,
    end: 3,
    kind: "column",
  });
  const after = session.selection!.id;
  expect(
    allParagraphs(session.document).find((p) => p.id === after)?.properties
      .columnBreakBefore,
  ).toBe(true);
  await session.execute({
    type: "splitParagraph",
    id: after,
    start: 1,
    end: 1,
  });
  expect(
    allParagraphs(session.document).find((p) => p.id === session.selection!.id)
      ?.properties.columnBreakBefore,
  ).toBeUndefined();
});
it("keeps physical odd/even pages separate from restarted PAGE field numbering", () => {
  const document = createDocumentEngine().document;
  const result = sectionPagePositions(
    [
      { ...document, page: { pageNumberStart: 5 } },
      { ...document, page: { sectionStart: "oddPage" } },
      { ...document, page: { sectionStart: "evenPage", pageNumberStart: 1 } },
      { ...document, page: {} },
    ],
    { 0: 1, 1: 2, 2: 1, 3: 1 },
  );
  expect(result.total).toBe(7);
  expect(result.positions.map((p) => p.blank)).toEqual([
    false,
    true,
    true,
    false,
  ]);
  expect(result.positions.map((p) => p.fieldOffset)).toEqual([0, 6, 0, 1]);
});
it("keeps Unicode filenames and exact file bytes inside Windows OLE packages", () => {
  const file = embeddedFile({
    name: "یادداشت.txt",
    dataUrl: "data:text/plain;base64,SGVsbG8=",
  });
  expect(file.kind).toBe("oleObject");
  expect(unpackFile(file.bytes)).toEqual({
    name: "یادداشت.txt",
    bytes: new TextEncoder().encode("Hello"),
  });
});
it("retains an existing header when a section is split in the portable engine", async () => {
  const engine = createDocumentEngine();
  await engine.execute({ type: "story", kind: "header" });
  const header = engine.selection!.id;
  await engine.execute({
    type: "text",
    id: header,
    start: 0,
    end: 0,
    text: "Header",
  });
  await engine.execute({
    type: "text",
    id: "p1",
    start: 0,
    end: 0,
    text: "BeforeAfter",
  });
  await engine.execute({
    type: "sectionBreak",
    id: "p1",
    start: 6,
    end: 6,
    kind: "nextPage",
  });
  const sections = splitSections(engine.document);
  expect(sections).toHaveLength(2);
  expect(sections.every((s) => s.headers?.default?.length)).toBe(true);
});
it("updates and removes drop caps without joining unrelated paragraphs", async () => {
  for (const session of [createDocumentEngine(), (await setup()).session]) {
    const id = allParagraphs(session.document)[0].id!;
    await session.execute({
      type: "text",
      id,
      start: 0,
      end: 0,
      text: "Drop capital",
    });
    await session.execute({
      type: "insertParagraph",
      id,
      text: "Next paragraph",
    });
    await session.execute({
      type: "dropCap",
      id,
      lines: 3,
      distancePt: 4,
      kind: "drop",
      remove: true,
    });
    expect(allParagraphs(session.document).map(text)).toEqual([
      "Drop capital",
      "Next paragraph",
    ]);
    await session.execute({
      type: "dropCap",
      id,
      lines: 3,
      distancePt: 4,
      kind: "drop",
    });
    expect(allParagraphs(session.document).map(text)).toEqual([
      "D",
      "rop capital",
      "Next paragraph",
    ]);
    await session.execute({
      type: "dropCap",
      id,
      lines: 4,
      distancePt: 6,
      kind: "drop",
    });
    expect(allParagraphs(session.document)[0].properties.dropCap).toMatchObject(
      { lines: 4, distancePt: 6 },
    );
    await session.execute({
      type: "dropCap",
      id,
      lines: 3,
      distancePt: 4,
      kind: "drop",
      remove: true,
    });
    expect(allParagraphs(session.document).map(text)).toEqual([
      "Drop capital",
      "Next paragraph",
    ]);
  }
});
