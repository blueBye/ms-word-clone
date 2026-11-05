import type {
  DocumentDiagnostic,
  DocumentIR,
  PageLayout,
  StoryVariants,
} from "../model/document.js";
import {
  findMainDocumentPart,
  openDocxPackage,
  type ImportOptions,
} from "./package.js";
import {
  childElements,
  childWord,
  descendants,
  isWord,
  numericAttribute,
  onOff,
  parseXml,
  wordAttribute,
} from "./parser/xml.js";
import { parseBlocks, parseDocumentXml } from "./parser/document.js";
import { readChart } from "./chart.js";
import { spreadsheetRows } from "./spreadsheet.js";
import { readProperties } from "./metadata.js";
import { readSection } from "./section.js";
import { unpackFile } from "./ole-package.js";
import {
  readRelationships,
  relationshipAttribute,
  type Relationship,
} from "./relationships.js";
import { Styles, parseTheme } from "./styles.js";
import { Numbering } from "./numbering.js";
import { colorValue } from "./parser/properties.js";
import type { ParseContext } from "./parser/context.js";
import type { Element } from "@xmldom/xmldom";
import type { Paragraph } from "../model/document.js";
import type { DocxPackage } from "./package.js";
import type { StyleCatalog } from "../core/styles.js";
import type { ListDefinition } from "../core/lists.js";

/** Import to preview IR; source packages are not yet round-trippable. */
export async function importDocx(
  input: ArrayBuffer | Uint8Array,
  options: ImportOptions = {},
): Promise<DocumentIR> {
  const pkg = await openDocxPackage(input, options);
  return parseDocxPackage(pkg);
}

/** Shared parser for a source-retaining session; not an application-specific API. */
export async function parseDocxPackage(
  pkg: DocxPackage,
  onParagraph?: (part: string, source: Element, paragraph: Paragraph) => void,
  onStyles?: (catalog: StyleCatalog, lists: ListDefinition[]) => void,
): Promise<DocumentIR> {
  const path = await findMainDocumentPart(pkg);
  const relationships = await readRelationships(pkg, path);
  const relatedXml = async (type: string) => {
    const relation = [...relationships.values()].find(
      (rel) => rel.type === type && !rel.external,
    );
    if (!relation || !pkg.hasPart(relation.target)) return undefined;
    const root = parseXml(await pkg.readText(relation.target)).documentElement!;
    const valid =
      type === "theme"
        ? root.localName === "theme" &&
          [
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
          ].includes(root.namespaceURI ?? "")
        : isWord(root, type);
    if (!valid) throw new Error(`Invalid ${type} part root`);
    return root;
  };
  const [styleXml, themeXml, numberingXml, settingsXml] = await Promise.all([
    relatedXml("styles"),
    relatedXml("theme"),
    relatedXml("numbering"),
    relatedXml("settings"),
  ]);
  const theme = parseTheme(themeXml);
  const diagnostics: DocumentDiagnostic[] = [];
  const warningKeys = new Set<string>();
  const warning = (part: string) => (code: string, message: string) => {
    const key = `${part}:${code}`;
    if (!warningKeys.has(key)) {
      warningKeys.add(key);
      diagnostics.push({ code, message, part });
    }
  };
  const styles = new Styles(styleXml, theme, warning(path));
  onStyles?.(styles.catalog(), new Numbering(numberingXml, theme).snapshot());
  const xml = parseXml(await pkg.readText(path));
  const sourceLanguage = /[پچژگکی۰-۹]/u.test(
    childWord(xml.documentElement!, "body")?.textContent ?? "",
  )
    ? "fa-IR"
    : styles.defaults.run.complexScriptLanguage;
  const assets = new Map<string, string>();
  const attachmentNames = new Map<string, string>();
  const charts: NonNullable<ParseContext["charts"]> = new Map();
  const spreadsheets = new Map<string, string[][]>();
  const comments = new Map<
    string,
    { text: string; author: string; date: string }
  >();
  const commentsPart = [...relationships.values()].find(
    (r) => r.type === "comments" && !r.external,
  );
  if (commentsPart && pkg.hasPart(commentsPart.target))
    for (const c of childElements(
      parseXml(await pkg.readText(commentsPart.target)).documentElement!,
    ))
      if (isWord(c, "comment"))
        comments.set(wordAttribute(c, "id") ?? "", {
          text: descendants(c, "t")
            .map((t) => t.textContent ?? "")
            .join(""),
          author: wordAttribute(c, "author") ?? "",
          date: wordAttribute(c, "date") ?? "",
        });
  async function context(
    part: string,
    rels: Map<string, Relationship>,
  ): Promise<ParseContext> {
    const types = parseXml(
      await pkg.readText("[Content_Types].xml"),
    ).documentElement!;
    const contentType = (path: string) =>
      childElements(types)
        .find(
          (e) =>
            e.localName === "Override" &&
            e.getAttribute("PartName") === `/${path}`,
        )
        ?.getAttribute("ContentType") ??
      childElements(types)
        .find(
          (e) =>
            e.localName === "Default" &&
            e.getAttribute("Extension") === path.split(".").at(-1),
        )
        ?.getAttribute("ContentType");
    for (const rel of rels.values()) {
      if (
        !rel.external &&
        ["model3d", "package", "oleObject"].includes(rel.type) &&
        !assets.has(rel.target) &&
        pkg.hasPart(rel.target)
      ) {
        if (rel.target.endsWith(".xlsx")) {
          try {
            spreadsheets.set(
              rel.target,
              await spreadsheetRows(await pkg.readBinary(rel.target)),
            );
          } catch {
            /* Preserve unsupported workbooks as embedded files. */
          }
        }
        const original = await pkg.readBinary(rel.target),
          unpacked =
            rel.type === "oleObject" ? unpackFile(original) : undefined;
        if (unpacked) attachmentNames.set(rel.target, unpacked.name);
        const mime = unpacked
            ? ((
                {
                  txt: "text/plain",
                  pdf: "application/pdf",
                  png: "image/png",
                  jpg: "image/jpeg",
                  csv: "text/csv",
                } as Record<string, string>
              )[unpacked.name.split(".").at(-1)!.toLowerCase()] ??
              "application/octet-stream")
            : rel.type === "model3d"
              ? "model/gltf-binary"
              : rel.target.endsWith(".xlsx")
                ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                : (contentType(rel.target) ?? "application/octet-stream"),
          bytes = unpacked?.bytes ?? original;
        let raw = "";
        for (let i = 0; i < bytes.length; i += 8192)
          raw += String.fromCharCode(...bytes.subarray(i, i + 8192));
        assets.set(rel.target, `data:${mime};base64,${btoa(raw)}`);
      }
      if (
        !rel.external &&
        rel.type === "chart" &&
        !charts.has(rel.target) &&
        pkg.hasPart(rel.target)
      ) {
        const chart = readChart(
          parseXml(await pkg.readText(rel.target)).documentElement!,
        );
        if (chart) charts.set(rel.target, chart);
      }
      if (
        rel.external ||
        rel.type !== "image" ||
        assets.has(rel.target) ||
        !pkg.hasPart(rel.target)
      )
        continue;
      const extension = rel.target.split(".").pop()?.toLowerCase() ?? "";
      const mime: Record<string, string> = {
        png: "image/png",
        jpg: "image/jpeg",
        jpeg: "image/jpeg",
        gif: "image/gif",
        webp: "image/webp",
        bmp: "image/bmp",
      };
      if (!mime[extension]) continue; // Prefer DrawingML's raster fallback to active SVG/unsupported formats.
      const bytes = await pkg.readBinary(rel.target);
      let binary = "";
      for (let i = 0; i < bytes.length; i += 8192)
        binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
      assets.set(rel.target, `data:${mime[extension]};base64,${btoa(binary)}`);
    }
    return {
      styles,
      theme,
      numbering: new Numbering(numberingXml, theme),
      relationships: rels,
      assets,
      attachmentNames,
      charts,
      spreadsheets,
      comments,
      warn: warning(part),
      numberingLanguage: sourceLanguage,
      onParagraph:
        onParagraph &&
        ((source, paragraph) => onParagraph(part, source, paragraph)),
    };
  }
  const ctx = await context(path, relationships);
  const document = parseDocumentXml(xml, ctx);
  document.properties = await readProperties(pkg);
  if (sourceLanguage) document.language = sourceLanguage;
  const body = childWord(xml.documentElement!, "body")!;
  const section = childWord(body, "sectPr");
  if (section) {
    const size = childWord(section, "pgSz"),
      margins = childWord(section, "pgMar");
    const page: PageLayout = {
      widthPt: (numericAttribute(size, "w") ?? 11906) / 20,
      heightPt: (numericAttribute(size, "h") ?? 16838) / 20,
      margins: {},
      differentFirstPage: onOff(childWord(section, "titlePg")),
      evenAndOddHeaders: settingsXml
        ? onOff(childWord(settingsXml, "evenAndOddHeaders"))
        : false,
      pageNumberStart: numericAttribute(
        childWord(section, "pgNumType"),
        "start",
      ),
    };
    for (const side of [
      "top",
      "right",
      "bottom",
      "left",
      "header",
      "footer",
    ] as const) {
      const value = numericAttribute(margins, side);
      if (value !== undefined) page.margins![side] = value / 20;
    }
    document.page = page;
    page.mirrorMargins = onOff(
      settingsXml && childWord(settingsXml, "mirrorMargins"),
    );
    page.hyphenation = onOff(
      settingsXml && childWord(settingsXml, "autoHyphenation"),
    )
      ? "auto"
      : "none";
    const gutter = numericAttribute(margins, "gutter");
    if (gutter !== undefined) page.gutterPt = gutter / 20;
    page.direction = onOff(childWord(section, "bidi")) ? "rtl" : "ltr";
    const alignment = childWord(section, "vAlign"),
      vertical = alignment ? wordAttribute(alignment, "val") : undefined;
    if (vertical === "center" || vertical === "bottom" || vertical === "top")
      page.verticalAlignment = vertical;
    const line = childWord(section, "lnNumType");
    if (line)
      page.lineNumbers = {
        mode:
          wordAttribute(line, "restart") === "newPage"
            ? "page"
            : wordAttribute(line, "restart") === "newSection"
              ? "section"
              : "continuous",
        start: (numericAttribute(line, "start") ?? 0) + 1,
        countBy: numericAttribute(line, "countBy") ?? 1,
        distancePt: (numericAttribute(line, "distance") ?? 360) / 20,
      };
    const columns = childWord(section, "cols");
    const count =
      numericAttribute(columns, "num") ??
      (columns
        ? childElements(columns).filter((e) => e.localName === "col").length
        : 0);
    if (count > 1)
      page.columns = {
        count,
        gapPt: (numericAttribute(columns, "space") ?? 720) / 20,
        separator: ["1", "true", "on"].includes(
          wordAttribute(columns!, "sep") ?? "",
        ),
      };
    for (const kind of ["header", "footer"] as const) {
      const variants: StoryVariants = {};
      for (const ref of childElements(section).filter(
        (n) => n.localName === `${kind}Reference`,
      )) {
        const id = relationshipAttribute(ref, "id");
        const rel = id && relationships.get(id);
        if (
          !rel ||
          rel.external ||
          rel.type !== kind ||
          !pkg.hasPart(rel.target)
        ) {
          ctx.warn(
            "missing-story",
            "A header or footer part could not be resolved.",
          );
          continue;
        }
        const type = wordAttribute(ref, "type");
        if (!["first", "default", "even"].includes(type ?? "")) continue;
        const story = parseXml(await pkg.readText(rel.target)).documentElement!;
        variants[type as keyof StoryVariants] = parseBlocks(
          story,
          await context(rel.target, await readRelationships(pkg, rel.target)),
        );
      }
      if (Object.keys(variants).length)
        document[kind === "header" ? "headers" : "footers"] = variants;
    }
  }
  document.backgroundColor = colorValue(
    childWord(xml.documentElement!, "background"),
    "color",
    theme,
  );
  if (!document.backgroundColor) delete document.backgroundColor;
  const marked = document.blocks.filter(
    (b): b is import("../model/document.js").Paragraph =>
      b.type === "paragraph" && !!b.section,
  );
  if (marked.length) {
    const nodes = descendants(body, "sectPr"),
      entries = [
        ...marked.map((p) => ({ paragraph: p, page: p.section! })),
        { paragraph: undefined, page: document.page ?? {} },
      ];
    if (nodes.length === entries.length) {
      const cache = new Map<string, import("../model/document.js").Block[]>();
      let headers: StoryVariants = {},
        footers: StoryVariants = {},
        previous: PageLayout = {
          widthPt: 595.3,
          heightPt: 841.9,
          margins: {
            top: 72,
            bottom: 72,
            left: 72,
            right: 72,
            header: 36,
            footer: 36,
          },
          columns: { count: 1, gapPt: 36 },
        };
      for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i],
          own = readSection(node),
          page = {
            ...previous,
            ...own,
            margins: { ...previous.margins, ...own.margins },
            mirrorMargins: document.page?.mirrorMargins,
            hyphenation: document.page?.hyphenation,
            evenAndOddHeaders: document.page?.evenAndOddHeaders,
            pageNumberStart: own.pageNumberStart,
            lineNumbers: own.lineNumbers,
            sectionStart: own.sectionStart,
          };
        headers = { ...headers };
        footers = { ...footers };
        for (const kind of ["header", "footer"] as const)
          for (const ref of childElements(node).filter(
            (n) => n.localName === `${kind}Reference`,
          )) {
            const variant = wordAttribute(ref, "type"),
              id = relationshipAttribute(ref, "id"),
              rel = id && relationships.get(id);
            if (
              !variant ||
              !["first", "default", "even"].includes(variant) ||
              !rel ||
              rel.external ||
              !pkg.hasPart(rel.target)
            )
              continue;
            let blocks = cache.get(rel.target);
            if (!blocks) {
              blocks = parseBlocks(
                parseXml(await pkg.readText(rel.target)).documentElement!,
                await context(
                  rel.target,
                  await readRelationships(pkg, rel.target),
                ),
              );
              cache.set(rel.target, blocks);
            }
            (kind === "header" ? headers : footers)[
              variant as keyof StoryVariants
            ] = blocks;
          }
        if (entries[i].paragraph)
          entries[i].paragraph!.section = { ...page, headers, footers };
        else {
          document.page = page;
          document.headers = headers;
          document.footers = footers;
        }
        previous = page;
      }
    } else
      ctx.warn(
        "section-layout",
        "Some nested section boundaries cannot be previewed independently.",
      );
  }
  if (document.headers || document.footers)
    ctx.warn(
      "pagination-approximation",
      "Page boundaries and repeating headers/footers use browser layout; spacing and page counts can differ from Word.",
    );
  if (descendants(xml.documentElement!, "anchor").length)
    ctx.warn(
      "anchored-layout",
      "Positioned drawings use browser layout; Word wrapping and pagination can differ.",
    );
  if (diagnostics.length) document.diagnostics = diagnostics;
  return document;
}
