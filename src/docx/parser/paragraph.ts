import type { Element } from "@xmldom/xmldom";
import type { Paragraph, Inline } from "../../model/document.js";
import {
  childWord,
  childElements,
  contentElements,
  isWord,
  wordAttribute,
} from "./xml.js";
import { localizeDigits } from "../../model/numerals.js";
import { parseRun } from "./run.js";
import {
  mergeParagraph,
  mergeRun,
  parseParagraphProperties,
  parseRunPropertyElement,
} from "./properties.js";
import { relationshipAttribute, safeHyperlink } from "../relationships.js";
import type { BlockParser, ParseContext } from "./context.js";
import { equationText } from "../equation.js";
import { readSection } from "../section.js";

export function parseParagraph(
  p: Element,
  context?: ParseContext,
  parseBlocks?: BlockParser,
): Paragraph {
  const pPr = childWord(p, "pPr");
  const direct = parseParagraphProperties(pPr, context?.theme);
  const resolved = context?.styles.paragraph(direct, context.tableLayers);
  let formattingContext = context?.styles.context(context.tableLayers);
  let properties = resolved?.properties ?? direct;
  const numbering = properties.numbering;
  let listMarkProperties;
  if (numbering?.id && numbering.id !== "0") {
    const level = context?.numbering.next(numbering.id, numbering.level ?? 0);
    if (level) {
      listMarkProperties = level.list.properties;
      properties = {
        ...mergeParagraph(level.paragraph, properties),
        list: level.list,
      };
      if (formattingContext)
        formattingContext = {
          ...formattingContext,
          paragraph: mergeParagraph(
            level.paragraph,
            formattingContext.paragraph,
          ),
        };
    } else
      context?.warn(
        "missing-numbering",
        "A list definition could not be resolved.",
      );
  }
  const children: Inline[] = [];
  const bookmarks: NonNullable<Paragraph["bookmarks"]> = [],
    comments: NonNullable<Paragraph["comments"]> = [];
  const bookmarkIds = new Map<string, number>(),
    commentIds = new Map<string, number>();
  let position = 0;
  const fields: { instruction: string; result: boolean }[] = [];
  const fieldKind = (instruction: string) =>
    /^\s*NUMPAGES\b/i.test(instruction)
      ? ("pages" as const)
      : /^\s*PAGE\b/i.test(instruction)
        ? ("page" as const)
        : /^\s*DATE\b/i.test(instruction)
          ? ("date" as const)
          : /^\s*(?:DOCPROPERTY\s+)?TITLE\b/i.test(instruction)
            ? ("title" as const)
            : /^\s*(?:DOCPROPERTY\s+)?AUTHOR\b/i.test(instruction)
              ? ("author" as const)
              : undefined;
  function visit(parent: Element, href?: string): void {
    for (const c of contentElements(parent)) {
      if (c.localName === "oMath" && c.namespaceURI?.includes("math")) {
        const expression = equationText(c);
        children.push({
          type: "drawing",
          alt: "Equation",
          widthPt: Math.min(432, Math.max(48, expression.length * 5)),
          heightPt: expression.includes("frac") ? 48 : 24,
          equation: { expression },
        });
        continue;
      }
      if (isWord(c, "bookmarkStart")) {
        bookmarkIds.set(wordAttribute(c, "id") ?? "", bookmarks.length);
        bookmarks.push({
          name: wordAttribute(c, "name") ?? "",
          start: position,
          end: position,
        });
        continue;
      }
      if (isWord(c, "bookmarkEnd")) {
        const index = bookmarkIds.get(wordAttribute(c, "id") ?? "");
        if (index !== undefined) bookmarks[index].end = position;
        continue;
      }
      if (isWord(c, "commentRangeStart")) {
        const key = wordAttribute(c, "id") ?? "",
          data = context?.comments?.get(key);
        commentIds.set(key, comments.length);
        comments.push({
          key,
          start: position,
          end: position,
          text: data?.text ?? "",
          author: data?.author ?? "",
          date: data?.date ?? "",
        });
        continue;
      }
      if (isWord(c, "commentRangeEnd")) {
        const index = commentIds.get(wordAttribute(c, "id") ?? "");
        if (index !== undefined) comments[index].end = position;
        continue;
      }
      if (isWord(c, "r")) {
        for (const node of childElements(c)) {
          if (
            isWord(node, "fldChar") &&
            wordAttribute(node, "fldCharType") === "begin"
          )
            fields.push({ instruction: "", result: false });
          if (isWord(node, "instrText") && fields.length)
            fields[fields.length - 1].instruction += node.textContent ?? "";
          if (
            isWord(node, "fldChar") &&
            wordAttribute(node, "fldCharType") === "separate" &&
            fields.length
          )
            fields[fields.length - 1].result = true;
        }
        const inlines = parseRun(c, resolved?.run, context, parseBlocks);
        for (const inline of inlines)
          if (inline.type === "text") {
            if (href) inline.href = href;
            const field = fields.at(-1);
            if (field?.result) inline.field = fieldKind(field.instruction);
          }
        children.push(...inlines);
        position += inlines.reduce(
          (n, r) => n + (r.type === "text" ? r.text.length : 0),
          0,
        );
        for (const node of childElements(c))
          if (
            isWord(node, "fldChar") &&
            wordAttribute(node, "fldCharType") === "end"
          )
            fields.pop();
      } else if (isWord(c, "sdt")) {
        const content = childWord(c, "sdtContent");
        if (content) visit(content, href);
      } else if (isWord(c, "hyperlink")) {
        const id = relationshipAttribute(c, "id");
        const rel = id && context?.relationships.get(id);
        const anchor = wordAttribute(c, "anchor");
        visit(
          c,
          anchor
            ? `#${encodeURIComponent(anchor)}`
            : rel && rel.type === "hyperlink"
              ? safeHyperlink(rel.target)
              : undefined,
        );
      } else if (isWord(c, "fldSimple")) {
        fields.push({
          instruction: wordAttribute(c, "instr") ?? "",
          result: true,
        });
        visit(c, href);
        fields.pop();
      } else if (
        ["sdtContent", "ins", "smartTag", "customXml", "moveTo"].some((name) =>
          isWord(c, name),
        )
      ) {
        visit(c, href);
      }
    }
  }
  visit(p);
  if (children[0]?.type === "break" && children[0].kind === "column") {
    children.shift();
    direct.columnBreakBefore = true;
    properties = { ...properties, columnBreakBefore: true };
  }
  const mark = pPr && childWord(pPr, "rPr");
  const markProperties = mergeRun(
    resolved?.run ?? {},
    parseRunPropertyElement(mark, context?.theme),
  );
  if (properties.list) {
    const textRuns = children.filter((inline) => inline.type === "text");
    const text = textRuns.map((run) => run.text).join("");
    const language = /[پچژگکی۰-۹]/u.test(text)
      ? "fa-IR"
      : markProperties.complexScriptLanguage?.startsWith("fa")
        ? markProperties.complexScriptLanguage
        : (context?.numberingLanguage ?? markProperties.complexScriptLanguage);
    const locale = properties.direction === "rtl" ? language : undefined;
    properties = {
      ...properties,
      list: {
        ...properties.list,
        properties: mergeRun(markProperties, properties.list.properties),
        locale,
        marker: properties.list.format?.startsWith("decimal")
          ? localizeDigits(properties.list.marker, locale)
          : properties.list.marker,
      },
    };
  }
  const paragraph: Paragraph = {
    ...(pPr && childWord(pPr, "sectPr")
      ? { section: readSection(childWord(pPr, "sectPr")!) }
      : {}),
    ...(bookmarks.length ? { bookmarks } : {}),
    ...(comments.length ? { comments } : {}),
    type: "paragraph",
    properties,
    children,
    ...(context
      ? {
          directProperties: direct,
          formattingContext,
          directMarkProperties: parseRunPropertyElement(mark, context.theme),
          listMarkProperties,
        }
      : {}),
    ...(Object.keys(markProperties).length
      ? {
          markProperties,
        }
      : {}),
  };
  context?.onParagraph?.(p, paragraph);
  return paragraph;
}
