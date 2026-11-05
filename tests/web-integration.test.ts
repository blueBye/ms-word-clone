import { readFile } from "node:fs/promises";
import { expect, it } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { importDocx } from "../src/index.js";
import { DocumentRenderer } from "../src/react/index.js";

it("imports the Next.js sample DOCX and renders both directions through the engine", async () => {
  const bytes = await readFile(
    new URL("../apps/web/public/sample.docx", import.meta.url),
  );
  const document = await importDocx(bytes);
  expect(document.blocks).toHaveLength(6);
  expect(document.blocks[2].properties.direction).toBe("rtl");
  expect(document.blocks[2].properties.alignment).toBe("start");
  const html = renderToStaticMarkup(DocumentRenderer({ document }));
  expect(html).toContain("A document in two directions");
  expect(html).toContain("این یک پاراگراف فارسی است");
  expect(html).toContain("هذه فقرة باللغة العربية");
  expect(html).toContain('dir="rtl"');
});
