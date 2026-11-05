import { expect, it } from "vitest";
import {
  createDocumentEngine,
  allParagraphs,
  tableContext,
  columnWidth,
  type Table,
  type DrawingInline,
} from "../src/core/index.js";
import {
  openDocxSession,
  createBlankDocx,
  importDocx,
} from "../src/docx/index.js";
import { getParagraphs } from "../src/model/traversal.js";
import { unzipSync, strFromU8 } from "fflate";
import {
  parseXml,
  descendants,
  wordAttribute,
} from "../src/docx/parser/xml.js";
import { OLE_PREVIEW } from "../src/docx/ole-preview.js";
import { docx, documentXml } from "./helpers/docx.js";
import { iconLibrary, vectorCommands } from "../src/model/graphics.js";
import { snapRuler } from "../apps/web/lib/ruler.js";
import { commentSelection } from "../apps/web/lib/comments.js";
const text = (p: ReturnType<typeof getParagraphs>[number]) =>
  p.children
    .filter((r) => r.type === "text")
    .map((r) => r.text)
    .join("");
it("normalizes the entire title from a substring, retaining bidi and later style linkage", async () => {
  for (const session of [
    createDocumentEngine(),
    await openDocxSession(createBlankDocx()),
  ]) {
    const id = allParagraphs(session.document)[0].id!;
    await session.execute({
      type: "text",
      id,
      start: 0,
      end: 0,
      text: "My mixed title",
    });
    await session.execute({
      type: "paragraph",
      id,
      format: { styleId: "Title", direction: "rtl", alignment: "right" },
    });
    await session.execute({
      type: "run",
      id,
      start: 0,
      end: 2,
      format: { color: "#CC2200", fontSizePt: 24 },
    });
    await session.execute({
      type: "run",
      id,
      start: 3,
      end: 8,
      format: { color: "#225588", italic: true, fontSizePt: 36 },
    });
    await session.execute({
      type: "updateStyleFromSelection",
      id,
      styleId: "Title",
      start: 0,
      end: 2,
      normalizeParagraph: true,
    });
    const p = allParagraphs(session.document)[0];
    expect(p.properties.direction).toBe("rtl");
    expect(
      p.children.every(
        (r) =>
          r.type === "text" &&
          r.properties.color === "#CC2200" &&
          r.properties.fontSizePt === 24 &&
          !r.properties.italic,
      ),
    ).toBe(true);
    await session.execute({
      type: "style",
      styleId: "Title",
      run: { color: "#119944" },
    });
    if ("exportDocx" in session) {
      const out = await importDocx(await session.exportDocx());
      expect(
        getParagraphs(out.blocks)[0].children.every(
          (r) => r.type === "text" && r.properties.color === "#119944",
        ),
      ).toBe(true);
    }
  }
});
it("exports RTL paragraph alignment logically and preserves direction after insertion", async () => {
  const s = await openDocxSession(createBlankDocx()),
    id = allParagraphs(s.document)[0].id!;
  await s.execute({
    type: "paragraph",
    id,
    format: { direction: "rtl", alignment: "right" },
  });
  await s.execute({
    type: "text",
    id,
    start: 0,
    end: 0,
    text: "متن English فارسی",
  });
  await s.execute({
    type: "insertImage",
    id,
    dataUrl: OLE_PREVIEW,
    widthPt: 120,
    heightPt: 50,
    alt: "Picture",
  });
  const after = s.selection!.id;
  expect(
    allParagraphs(s.document).find((p) => p.id === after)?.properties.direction,
  ).toBe("rtl");
  await s.execute({
    type: "text",
    id: after,
    start: 0,
    end: 0,
    text: "ادامه متن",
  });
  await s.execute({ type: "insertTable", id: after, rows: 2, columns: 2 });
  const out = await s.exportDocx(),
    root = parseXml(
      strFromU8(unzipSync(out)["word/document.xml"]),
    ).documentElement!;
  const rtl = descendants(root, "p").filter((p) =>
    descendants(p, "bidi").some((b) => wordAttribute(b, "val") !== "0"),
  );
  expect(rtl.length).toBeGreaterThan(4);
  expect(
    rtl.every((p) => wordAttribute(descendants(p, "jc")[0], "val") === "start"),
  ).toBe(true);
  const restored = await openDocxSession(out);
  expect(
    getParagraphs(restored.document.blocks)
      .filter((p) => text(p))
      .every((p) => p.properties.direction === "rtl"),
  ).toBe(true);
});
it("fits inline pictures and charts to columns in both the engine and Word file", async () => {
  for (const s of [
    createDocumentEngine(),
    await openDocxSession(createBlankDocx()),
  ]) {
    let id = allParagraphs(s.document)[0].id!;
    const drawing: DrawingInline = {
      type: "drawing",
      alt: "Chart",
      widthPt: 360,
      heightPt: 240,
      chart: {
        kind: "bar",
        title: "Results",
        categories: ["A", "B"],
        series: [{ name: "Count", values: [-4, 7] }],
      },
    };
    await s.execute({ type: "insertDrawing", id, drawing });
    await s.execute({
      type: "insertImage",
      id,
      dataUrl: OLE_PREVIEW,
      widthPt: 400,
      heightPt: 200,
      alt: "Picture",
    });
    await s.execute({ type: "columns", count: 2, gapPt: 36 });
    const max = columnWidth(s.document.page);
    for (const d of allParagraphs(s.document)
      .flatMap((p) => p.children)
      .filter((r) => r.type === "drawing"))
      expect(d.widthPt).toBeLessThanOrEqual(max + 0.001);
    if ("exportDocx" in s) {
      const files = unzipSync(await s.exportDocx()),
        chart = Object.entries(files).find(
          ([p]) => p.includes("charts/") && p.endsWith(".xml"),
        )![1],
        xml = parseXml(strFromU8(chart)).documentElement!;
      expect(
        descendants(xml, "tickLblPos").map((e) => e.getAttribute("val")),
      ).toEqual(["low", "low"]);
      expect(
        descendants(xml, "delete").map((e) => e.getAttribute("val")),
      ).toEqual(["0", "0"]);
      expect(
        descendants(xml, "overlay").map((e) => e.getAttribute("val")),
      ).toEqual(["0", "0"]);
      const restored = await importDocx(await s.exportDocx());
      expect(
        getParagraphs(restored.blocks)
          .flatMap((p) => p.children)
          .filter((r) => r.type === "drawing")
          .every((r) => r.widthPt <= max + 0.001),
      ).toBe(true);
    }
  }
});
it("edits table structure and cell properties with preserved content and native round-trip", async () => {
  for (const s of [
    createDocumentEngine(),
    await openDocxSession(createBlankDocx()),
  ]) {
    const id = allParagraphs(s.document)[0].id!;
    await s.execute({
      type: "insertTable",
      id,
      rows: 2,
      columns: 2,
      cells: [
        ["A", "B"],
        ["C", "D"],
      ],
    });
    let table = s.document.blocks.find((b): b is Table => b.type === "table")!,
      anchor = getParagraphs(table.rows[0].cells[0].blocks)[0].id!;
    await s.execute({ type: "table", id: anchor, action: "insertRowBelow" });
    table = tableContext(s.document, anchor)!.table;
    expect(table.rows).toHaveLength(3);
    expect(getParagraphs(table.rows[2].cells[1].blocks).map(text)).toEqual([
      "D",
    ]);
    await s.execute({ type: "table", id: anchor, action: "insertColumnAfter" });
    table = tableContext(s.document, anchor)!.table;
    expect(table.columnsPt).toHaveLength(3);
    expect(getParagraphs(table.rows[2].cells[2].blocks).map(text)).toEqual([
      "D",
    ]);
    await s.execute({
      type: "table",
      id: anchor,
      action: "merge",
      range: { top: 0, bottom: 0, left: 0, right: 2 },
    });
    table = tableContext(s.document, anchor)!.table;
    expect(table.rows[0].cells[0].columnSpan).toBe(3);
    expect(
      getParagraphs(table.rows[0].cells[0].blocks).map(text).join(""),
    ).toBe("AB");
    await s.execute({
      type: "table",
      id: anchor,
      format: {
        shading: "#EEDDCC",
        verticalAlign: "center",
        cellMargins: { top: 8, bottom: 8, left: 10, right: 10 },
        repeatHeader: true,
      },
    });
    if ("exportDocx" in s) {
      const restored = await openDocxSession(await s.exportDocx()),
        t = restored.document.blocks.find(
          (b): b is Table => b.type === "table",
        )!;
      expect(t.rows).toHaveLength(3);
      expect(t.rows[0].cells[0]).toMatchObject({
        columnSpan: 3,
        backgroundColor: "#EEDDCC",
        verticalAlign: "center",
      });
      expect(t.properties.cellMargins).toMatchObject({ top: 8, left: 10 });
      expect(getParagraphs([t]).map(text).join("")).toContain("CD");
    }
    await s.undo();
    await s.undo();
    table = tableContext(s.document, anchor)!.table;
    expect(table.rows[0].cells).toHaveLength(3);
  }
});
it("round-trips picture effects, crop and anchor locking", async () => {
  const s = await openDocxSession(createBlankDocx()),
    id = allParagraphs(s.document)[0].id!;
  await s.execute({
    type: "insertImage",
    id,
    dataUrl: OLE_PREVIEW,
    widthPt: 180,
    heightPt: 100,
    alt: "Picture",
  });
  const p = allParagraphs(s.document).find((p) =>
      p.children.some((r) => r.type === "drawing"),
    )!,
    index = p.children.findIndex((r) => r.type === "drawing"),
    d = p.children[index] as DrawingInline;
  await s.execute({
    type: "drawing",
    id: p.id!,
    index,
    drawing: {
      ...d,
      image: {
        ...d.image!,
        brightness: 0.2,
        contrast: 0.4,
        grayscale: true,
        blurPt: 2,
        opacity: 0.7,
        crop: { left: 0.1, right: 0.2 },
      },
      position: {
        horizontal: { relativeTo: "column", offsetPt: 0 },
        vertical: { relativeTo: "paragraph", offsetPt: 0 },
        wrap: "wrapSquare",
        behindText: false,
        zIndex: 2,
        locked: true,
      },
    },
  });
  const restored = await openDocxSession(await s.exportDocx());
  expect(
    allParagraphs(restored.document)
      .flatMap((p) => p.children)
      .find((r) => r.type === "drawing"),
  ).toMatchObject({
    image: {
      brightness: 0.2,
      contrast: 0.4,
      grayscale: true,
      blurPt: 2,
      opacity: 0.7,
      crop: { left: 0.1, right: 0.2 },
    },
    position: { locked: true },
  });
});
it("fits newly inserted objects in an existing multi-column section", async () => {
  for (const s of [
    createDocumentEngine(),
    await openDocxSession(createBlankDocx()),
  ]) {
    const id = allParagraphs(s.document)[0].id!;
    await s.execute({ type: "columns", count: 3, gapPt: 18 });
    await s.execute({
      type: "insertImage",
      id,
      dataUrl: OLE_PREVIEW,
      widthPt: 400,
      heightPt: 200,
      alt: "New picture",
    });
    await s.execute({
      type: "insertDrawing",
      id,
      drawing: {
        type: "drawing",
        widthPt: 400,
        heightPt: 200,
        alt: "New chart",
        chart: {
          kind: "line",
          title: "Test",
          categories: ["A"],
          series: [{ name: "Count", values: [3] }],
        },
      },
    });
    const available = columnWidth(s.document.page),
      objects = allParagraphs(s.document)
        .flatMap((p) => p.children)
        .filter((r) => r.type === "drawing");
    expect(objects).toHaveLength(2);
    expect(
      objects.every(
        (d) =>
          d.widthPt <= available + 0.001 &&
          Math.abs(d.widthPt / d.heightPt - 2) < 0.001,
      ),
    ).toBe(true);
  }
});
it("edits and removes Word display equations, preserving neighbouring text and undo", async () => {
  const bytes = docx(
    documentXml(
      '<w:p><w:r><w:t>Before</w:t></w:r><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara><w:r><w:t>After</w:t></w:r></w:p>',
    ),
  );
  const s = await openDocxSession(bytes),
    p = allParagraphs(s.document)[0],
    index = p.children.findIndex((r) => r.type === "drawing"),
    drawing = p.children[index] as DrawingInline;
  expect(index).toBeGreaterThan(-1);
  await s.execute({
    type: "drawing",
    id: p.id!,
    index,
    drawing: { ...drawing, equation: { expression: "\\frac{x+1}{2}" } },
  });
  let restored = await importDocx(await s.exportDocx());
  expect(
    getParagraphs(restored.blocks)[0].children.find(
      (r) => r.type === "drawing",
    ),
  ).toMatchObject({
    equation: { expression: expect.stringContaining("frac") },
  });
  await s.execute({ type: "drawing", id: p.id!, index, remove: true });
  restored = await importDocx(await s.exportDocx());
  expect(text(getParagraphs(restored.blocks)[0])).toBe("BeforeAfter");
  expect(
    getParagraphs(restored.blocks)[0].children.some(
      (r) => r.type === "drawing",
    ),
  ).toBe(false);
  await s.undo();
  expect(
    allParagraphs(s.document)[0].children.some((r) => r.type === "drawing"),
  ).toBe(true);
});
it("exports distinct icons as editable native vector geometry", async () => {
  const s = await openDocxSession(createBlankDocx()),
    id = allParagraphs(s.document)[0].id!;
  for (const icon of iconLibrary) {
    expect(vectorCommands(icon.path).length).toBeGreaterThan(3);
    await s.execute({
      type: "insertDrawing",
      id,
      drawing: {
        type: "drawing",
        widthPt: 24,
        heightPt: 24,
        alt: icon.name,
        shape: { kind: `icon_${icon.id}`, fill: "#123456", strokeWidthPt: 0 },
      },
    });
  }
  const bytes = await s.exportDocx(),
    restored = await importDocx(bytes),
    drawings = getParagraphs(restored.blocks)
      .flatMap((p) => p.children)
      .filter((r) => r.type === "drawing");
  expect(new Set(drawings.map((d) => d.shape?.kind)).size).toBe(
    iconLibrary.length,
  );
  expect(drawings.every((d) => d.shape?.fill === "#123456")).toBe(true);
  const root = parseXml(
    strFromU8(unzipSync(bytes)["word/document.xml"]),
  ).documentElement!;
  expect(descendants(root, "custGeom")).toHaveLength(iconLibrary.length);
});
it("splits cells in both directions and removes rows and columns without losing unrelated text", async () => {
  for (const s of [
    createDocumentEngine(),
    await openDocxSession(createBlankDocx()),
  ]) {
    await s.execute({
      type: "insertTable",
      rows: 2,
      columns: 2,
      cells: [
        ["Keep A", "Keep B"],
        ["Keep C", "Keep D"],
      ],
    });
    let t = s.document.blocks.find((b): b is Table => b.type === "table")!,
      id = getParagraphs(t.rows[0].cells[0].blocks)[0].id!;
    await s.execute({
      type: "table",
      id,
      action: "split",
      rows: 2,
      columns: 2,
    });
    t = tableContext(s.document, id)!.table;
    expect(t.rows).toHaveLength(3);
    expect(t.columnsPt).toHaveLength(3);
    expect(t.rows[0].cells.at(-1)?.rowSpan).toBe(2);
    expect(t.rows[2].cells[0].columnSpan).toBe(2);
    await s.execute({
      type: "table",
      id,
      action: "merge",
      range: { top: 0, bottom: 1, left: 0, right: 1 },
    });
    await s.execute({
      type: "table",
      id,
      action: "split",
      rows: 2,
      columns: 2,
    });
    await s.execute({ type: "table", id, action: "deleteColumn" });
    t = s.document.blocks.find((b): b is Table => b.type === "table")!;
    id = getParagraphs(t.rows[1].cells[0].blocks)[0].id!;
    await s.execute({ type: "table", id, action: "deleteRow" });
    t = (
      "exportDocx" in s ? await importDocx(await s.exportDocx()) : s.document
    ).blocks.find((b): b is Table => b.type === "table")!;
    expect(t.rows, s.constructor.name).toHaveLength(2);
    expect(t.columnsPt).toHaveLength(2);
    expect(getParagraphs([t]).map(text).join("")).toContain("Keep B");
    expect(getParagraphs([t]).map(text).join("")).toContain("Keep D");
  }
});
it("snaps ruler indicators to fine ticks and anchors point comments to text", () => {
  expect(snapRuler(13)).toBe(13.5);
  expect(snapRuler(-8)).toBe(-9);
  expect(commentSelection("a sample title", { start: 5, end: 5 })).toEqual({
    start: 2,
    end: 8,
  });
  expect(commentSelection("a sample title", { start: 2, end: 4 })).toEqual({
    start: 2,
    end: 4,
  });
});
it("keeps source cell metadata, pattern fills and bookmarks when changing table structure", async () => {
  const xml = documentXml(
    '<w:p/><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:val="pct20" w:color="FF0000" w:fill="FFFF00"/></w:tcPr><w:p w:rsidR="12345678"><w:bookmarkStart w:id="5" w:name="kept"/><w:r><w:rPr><w:i/></w:rPr><w:t>Keep this content</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p></w:tc></w:tr></w:tbl>',
  );
  const s = await openDocxSession(
      docx(xml, { "custom/preserved.bin": new Uint8Array([5, 8, 13]) }),
    ),
    id = allParagraphs(s.document).find(
      (p) => text(p) === "Keep this content",
    )!.id!;
  await s.execute({ type: "table", id, action: "insertRowBelow" });
  await s.execute({
    type: "table",
    id,
    format: { cellMargins: { top: 0, bottom: 0, left: 8, right: 8 } },
  });
  const files = unzipSync(await s.exportDocx()),
    root = parseXml(strFromU8(files["custom/main.xml"])).documentElement!;
  expect(files["custom/preserved.bin"]).toEqual(new Uint8Array([5, 8, 13]));
  expect(wordAttribute(descendants(root, "shd")[0], "val")).toBe("pct20");
  expect(descendants(root, "bookmarkStart")[0].getAttribute("w:name")).toBe(
    "kept",
  );
  expect(
    descendants(root, "p").some(
      (p) => p.getAttribute("w:rsidR") === "12345678",
    ),
  ).toBe(true);
  expect(descendants(root, "i")).toHaveLength(1);
});
it("updates native shape fill and picture outlines without replacing the source object", async () => {
  const s = await openDocxSession(createBlankDocx());
  await s.execute({
    type: "insertDrawing",
    drawing: {
      type: "drawing",
      alt: "Shape",
      widthPt: 72,
      heightPt: 72,
      shape: {
        kind: "heart",
        fill: "#4472C4",
        stroke: "#123456",
        strokeWidthPt: 1,
      },
    },
  });
  await s.execute({
    type: "insertImage",
    dataUrl: OLE_PREVIEW,
    widthPt: 72,
    heightPt: 48,
    alt: "Picture",
  });
  for (const p of allParagraphs(s.document))
    for (let index = 0; index < p.children.length; index++) {
      const d = p.children[index];
      if (d.type === "drawing")
        await s.execute({
          type: "drawing",
          id: p.id!,
          index,
          drawing: {
            ...d,
            shape: {
              ...d.shape,
              kind: d.shape?.kind ?? "rect",
              fill: d.image ? undefined : "#EE8800",
              stroke: "#CC2200",
              strokeWidthPt: 3,
            },
          },
        });
    }
  const restored = await importDocx(await s.exportDocx()),
    drawings = getParagraphs(restored.blocks)
      .flatMap((p) => p.children)
      .filter((d) => d.type === "drawing");
  expect(
    drawings.every(
      (d) => d.shape?.stroke === "#CC2200" && d.shape.strokeWidthPt === 3,
    ),
  ).toBe(true);
  expect(drawings.find((d) => !d.image)?.shape?.fill).toBe("#EE8800");
});
