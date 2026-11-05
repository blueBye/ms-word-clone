import type { Element } from "@xmldom/xmldom";
import type { PageLayout } from "../model/document.js";
import {
  childWord,
  childElements,
  numericAttribute,
  onOff,
  wordAttribute,
} from "./parser/xml.js";
export function readSection(section: Element): PageLayout {
  const size = childWord(section, "pgSz"),
    mar = childWord(section, "pgMar"),
    columns = childWord(section, "cols"),
    type = childWord(section, "type"),
    number = childWord(section, "pgNumType"),
    page: PageLayout = {};
  if (size) {
    const width = numericAttribute(size, "w"),
      height = numericAttribute(size, "h");
    if (width) page.widthPt = width / 20;
    if (height) page.heightPt = height / 20;
  }
  if (mar) {
    page.margins = {};
    for (const key of [
      "top",
      "bottom",
      "left",
      "right",
      "header",
      "footer",
    ] as const) {
      const value = numericAttribute(mar, key);
      if (value !== undefined) page.margins[key] = value / 20;
    }
    const gutter = numericAttribute(mar, "gutter");
    if (gutter !== undefined) page.gutterPt = gutter / 20;
  }
  if (columns)
    page.columns = {
      count:
        numericAttribute(columns, "num") ??
        Math.max(
          1,
          childElements(columns).filter((c) => c.localName === "col").length,
        ),
      gapPt: (numericAttribute(columns, "space") ?? 720) / 20,
      separator: ["1", "true"].includes(wordAttribute(columns, "sep") ?? ""),
    };
  const start = type && wordAttribute(type, "val");
  if (
    start &&
    ["nextPage", "continuous", "evenPage", "oddPage"].includes(start)
  )
    page.sectionStart = start as PageLayout["sectionStart"];
  page.differentFirstPage = onOff(childWord(section, "titlePg"));
  const direction = onOff(childWord(section, "bidi"));
  if (direction !== undefined) page.direction = direction ? "rtl" : "ltr";
  const align = childWord(section, "vAlign"),
    value = align && wordAttribute(align, "val");
  if (value && ["top", "center", "bottom"].includes(value))
    page.verticalAlignment = value as PageLayout["verticalAlignment"];
  const pageStart = numericAttribute(number, "start");
  if (pageStart !== undefined) page.pageNumberStart = pageStart;
  const line = childWord(section, "lnNumType");
  if (line)
    page.lineNumbers = {
      start: (numericAttribute(line, "start") ?? 0) + 1,
      countBy: numericAttribute(line, "countBy") ?? 1,
      distancePt: (numericAttribute(line, "distance") ?? 360) / 20,
      mode:
        wordAttribute(line, "restart") === "newPage"
          ? "page"
          : wordAttribute(line, "restart") === "newSection"
            ? "section"
            : "continuous",
    };
  return page;
}
export function sectionFor(root: Element, paragraph?: Element) {
  const body = childWord(root, "body");
  if (!body) throw new Error("No document body.");
  let block = paragraph;
  while (block && block.parentNode !== body)
    block =
      block.parentNode?.nodeType === 1
        ? (block.parentNode as Element)
        : undefined;
  const children = childElements(body),
    index = block ? children.indexOf(block) : children.length - 1;
  for (const child of children.slice(Math.max(0, index))) {
    if (child.localName === "sectPr") return child;
    const section = childWord(childWord(child, "pPr") ?? child, "sectPr");
    if (section) return section;
  }
  let section = childWord(body, "sectPr");
  if (!section) {
    section = root.ownerDocument!.createElementNS(
      root.namespaceURI,
      "w:sectPr",
    );
    body.appendChild(section);
  }
  return section;
}
