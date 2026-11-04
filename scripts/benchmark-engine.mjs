import { createDocumentEngine } from "../dist/core/index.js";

const count = 10_000;
const snapshot = {
  document: {
    type: "document",
    schemaVersion: 2,
    blocks: Array.from({ length: count }, (_, i) => ({
      type: "paragraph",
      id: `p${i + 1}`,
      properties: { styleId: i < 100 ? "Heading" : "Normal" },
      directProperties: { styleId: i < 100 ? "Heading" : "Normal" },
      children: [
        {
          type: "text",
          text: `Paragraph ${i} فارسی`,
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
};
const start = performance.now(),
  engine = createDocumentEngine(snapshot);
const initializationMs = performance.now() - start;
const measure = async (command) => {
  const start = performance.now();
  await engine.execute(command);
  return {
    milliseconds: Number((performance.now() - start).toFixed(2)),
    resolvedParagraphs: engine.statistics.resolvedParagraphs,
  };
};
const textEdit = await measure({
  type: "text",
  id: "p500",
  start: 0,
  end: 0,
  text: "New ",
});
const headingStyle = await measure({
  type: "style",
  styleId: "Heading",
  run: { color: "#225588" },
});
const baseStyle = await measure({
  type: "style",
  styleId: "Normal",
  run: { fontFamily: "Tahoma" },
});
console.log(
  JSON.stringify(
    {
      paragraphs: count,
      initializationMs: Number(initializationMs.toFixed(2)),
      textEdit,
      headingStyle,
      baseStyle,
    },
    null,
    2,
  ),
);
