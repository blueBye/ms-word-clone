import { XMLSerializer, type Element } from "@xmldom/xmldom";
import type { DocxPackage } from "./package.js";
import type { EditCommand } from "./session.js";
import { readRelationships } from "./relationships.js";
import {
  childElements,
  childWord,
  isWord,
  parseXml,
  wordAttribute,
} from "./parser/xml.js";

const serialize = (root: Element) =>
  new XMLSerializer().serializeToString(root.ownerDocument!);
const relNS = "http://schemas.openxmlformats.org/package/2006/relationships";
const contentNS =
  "http://schemas.openxmlformats.org/package/2006/content-types";
function child(parent: Element, name: string) {
  const existing = childWord(parent, name);
  if (existing) return existing;
  const node = parent.ownerDocument!.createElementNS(
    parent.namespaceURI,
    `${parent.prefix || "w"}:${name}`,
  );
  parent.appendChild(node);
  return node;
}
function value(node: Element, key: string, text: string | number) {
  node.setAttributeNS(
    node.namespaceURI,
    `${node.prefix || "w"}:${key}`,
    String(text),
  );
}
function val(parent: Element, name: string, text: string | number) {
  const node = child(parent, name);
  value(node, "val", text);
  return node;
}

/** Add a relationship-backed XML part without rebuilding unrelated package parts. */
async function related(
  kind: "styles" | "numbering",
  changes: Map<string, string>,
  pkg: DocxPackage,
  main: string,
) {
  const relationships = await readRelationships(pkg, main);
  const found = [...relationships.values()].find(
    (r) => r.type === kind && !r.external,
  );
  if (found) {
    if (!pkg.hasPart(found.target)) throw new Error(`Missing ${kind} part.`);
    return {
      path: found.target,
      root: parseXml(await pkg.readText(found.target)).documentElement!,
    };
  }
  const folder = main.split("/").slice(0, -1).join("/"),
    basename = main.split("/").at(-1)!;
  const documentRoot = parseXml(await pkg.readText(main)).documentElement!,
    w = documentRoot.namespaceURI!;
  let filename = `${kind}.xml`,
    n = 1;
  while (pkg.hasPart(`${folder ? folder + "/" : ""}${filename}`))
    filename = `editor-${kind}-${n++}.xml`;
  const path = `${folder ? folder + "/" : ""}${filename}`,
    relsPath = `${folder ? folder + "/" : ""}_rels/${basename}.rels`;
  const relRoot = pkg.hasPart(relsPath)
    ? parseXml(await pkg.readText(relsPath)).documentElement!
    : parseXml(`<Relationships xmlns="${relNS}"/>`).documentElement!;
  const ids = new Set(childElements(relRoot).map((e) => e.getAttribute("Id")));
  let index = 1;
  while (ids.has(`rId${index}`)) index++;
  const relation = relRoot.ownerDocument!.createElementNS(
    relNS,
    "Relationship",
  );
  relation.setAttribute("Id", `rId${index}`);
  relation.setAttribute(
    "Type",
    `${w.includes("purl.oclc.org") ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}/${kind}`,
  );
  relation.setAttribute("Target", filename);
  relRoot.appendChild(relation);
  changes.set(relsPath, serialize(relRoot));
  const types = parseXml(
    await pkg.readText("[Content_Types].xml"),
  ).documentElement!;
  const type = types.ownerDocument!.createElementNS(contentNS, "Override");
  type.setAttribute("PartName", `/${path}`);
  type.setAttribute(
    "ContentType",
    `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`,
  );
  types.appendChild(type);
  changes.set("[Content_Types].xml", serialize(types));
  return {
    path,
    root: parseXml(`<w:${kind} xmlns:w="${w}"/>`).documentElement!,
  };
}

const builtins: Record<
  string,
  {
    name: string;
    size: number;
    bold?: boolean;
    after?: number;
    before?: number;
    outline?: number;
  }
> = {
  Normal: { name: "Normal", size: 12, after: 8 },
  NoSpacing: { name: "No Spacing", size: 12, after: 0 },
  Heading1: {
    name: "Heading 1",
    size: 20,
    bold: true,
    before: 12,
    after: 6,
    outline: 0,
  },
  Heading2: {
    name: "Heading 2",
    size: 16,
    bold: true,
    before: 10,
    after: 4,
    outline: 1,
  },
  Heading3: {
    name: "Heading 3",
    size: 14,
    bold: true,
    before: 8,
    after: 4,
    outline: 2,
  },
  Title: { name: "Title", size: 28, after: 12 },
  Subtitle: { name: "Subtitle", size: 16, after: 8 },
};
export async function readStyles(pkg: DocxPackage, main: string) {
  const rels = await readRelationships(pkg, main),
    found = [...rels.values()].find((r) => r.type === "styles" && !r.external);
  const result = new Map(
    Object.entries(builtins).map(([id, style]) => [
      id,
      { id, name: style.name },
    ]),
  );
  if (found && pkg.hasPart(found.target)) {
    const root = parseXml(await pkg.readText(found.target)).documentElement!;
    for (const style of childElements(root).filter(
      (e) => isWord(e, "style") && wordAttribute(e, "type") === "paragraph",
    )) {
      const id = wordAttribute(style, "styleId");
      if (id)
        result.set(id, {
          id,
          name: wordAttribute(childWord(style, "name") ?? style, "val") ?? id,
        });
    }
  }
  return [...result.values()];
}
export async function ensureStyle(
  id: string,
  changes: Map<string, string>,
  pkg: DocxPackage,
  main: string,
) {
  const { path, root } = await related("styles", changes, pkg, main);
  if (
    childElements(root).some(
      (e) => isWord(e, "style") && wordAttribute(e, "styleId") === id,
    )
  )
    return;
  const preset = builtins[id];
  if (!preset)
    throw new Error("The selected style does not exist in this document.");
  const style = root.ownerDocument!.createElementNS(
    root.namespaceURI,
    `${root.prefix || "w"}:style`,
  );
  value(style, "type", "paragraph");
  value(style, "styleId", id);
  root.appendChild(style);
  val(style, "name", preset.name);
  if (id !== "Normal") val(style, "basedOn", "Normal");
  else value(style, "default", "1");
  val(style, "qFormat", "1");
  const p = child(style, "pPr");
  if (preset.outline !== undefined) val(p, "keepNext", "1");
  const spacing = child(p, "spacing");
  value(spacing, "before", (preset.before ?? 0) * 20);
  value(spacing, "after", (preset.after ?? 0) * 20);
  if (preset.outline !== undefined) val(p, "outlineLvl", preset.outline);
  const r = child(style, "rPr");
  if (preset.bold) {
    val(r, "b", "1");
    val(r, "bCs", "1");
  }
  val(r, "sz", preset.size * 2);
  val(r, "szCs", preset.size * 2);
  changes.set(path, serialize(root));
  if (
    id !== "Normal" &&
    !childElements(root).some((e) => wordAttribute(e, "styleId") === "Normal")
  )
    await ensureStyle(
      "Normal",
      changes,
      {
        ...pkg,
        hasPart: (p) => changes.has(p) || pkg.hasPart(p),
        readText: async (p) => changes.get(p) ?? pkg.readText(p),
      },
      main,
    );
}

export async function setList(
  p: Element,
  command: Extract<EditCommand, { type: "list" }>,
  changes: Map<string, string>,
  pkg: DocxPackage,
  main: string,
  rtl: boolean,
) {
  const pPr = child(p, "pPr");
  if (pPr !== p.firstChild) p.insertBefore(pPr, p.firstChild);
  let numPr = childWord(pPr, "numPr");
  if (!numPr) {
    numPr = child(pPr, "numPr");
    const following = childElements(pPr).find((e) =>
      [
        "pBdr",
        "shd",
        "tabs",
        "bidi",
        "spacing",
        "ind",
        "jc",
        "rPr",
        "sectPr",
      ].some((n) => isWord(e, n)),
    );
    if (following) pPr.insertBefore(numPr, following);
  }
  if (command.kind === "none") {
    val(numPr, "numId", "0");
    return;
  }
  if (
    !["bullet", "decimal", "lowerLetter", "upperRoman"].includes(command.kind)
  )
    throw new Error("Unsupported list format.");
  const level = command.level ?? 0;
  if (!Number.isInteger(level) || level < 0 || level > 8)
    throw new Error("List level must be between 1 and 9.");
  const start = command.start ?? 1;
  if (!Number.isSafeInteger(start) || start < 1 || start > 1_000_000)
    throw new Error("List start must be a positive whole number.");
  const { path, root } = await related("numbering", changes, pkg, main);
  const nums = childElements(root).filter((e) => isWord(e, "num")),
    abstracts = childElements(root).filter((e) => isWord(e, "abstractNum"));
  // Continue the existing list for a level change; restart/format changes create a new instance.
  const siblings = childElements(p.parentNode as Element),
    previous = siblings[siblings.indexOf(p) - 1];
  const previousPr =
    previous && isWord(previous, "p") ? childWord(previous, "pPr") : undefined;
  const previousNum = previousPr && childWord(previousPr, "numPr");
  const existing =
    wordAttribute(childWord(numPr, "numId") ?? numPr, "val") ??
    (previousNum &&
      wordAttribute(childWord(previousNum, "numId") ?? previousNum, "val"));
  let id = existing;
  const definition = id && nums.find((e) => wordAttribute(e, "numId") === id);
  const abstractId =
    definition &&
    wordAttribute(childWord(definition, "abstractNumId") ?? definition, "val");
  const current = abstracts.find(
    (e) => wordAttribute(e, "abstractNumId") === abstractId,
  );
  const currentKind =
    current &&
    wordAttribute(
      childWord(
        childElements(current).find((e) => isWord(e, "lvl")) ?? current,
        "numFmt",
      ) ?? current,
      "val",
    );
  if (
    !definition ||
    command.restart ||
    command.start !== undefined ||
    currentKind !== command.kind
  ) {
    const next = (items: Element[], key: string) =>
      String(
        Math.max(0, ...items.map((e) => Number(wordAttribute(e, key)) || 0)) +
          1,
      );
    const a = root.ownerDocument!.createElementNS(
      root.namespaceURI,
      `${root.prefix || "w"}:abstractNum`,
    );
    const aId = next(abstracts, "abstractNumId");
    value(a, "abstractNumId", aId);
    const firstNum = nums[0];
    root.insertBefore(a, firstNum ?? null);
    val(a, "multiLevelType", "multilevel");
    for (let i = 0; i < 9; i++) {
      const lvl = root.ownerDocument!.createElementNS(
        root.namespaceURI,
        `${root.prefix || "w"}:lvl`,
      );
      value(lvl, "ilvl", i);
      a.appendChild(lvl);
      val(lvl, "start", i === level ? start : 1);
      val(lvl, "numFmt", command.kind);
      const bullet = command.bullet ?? "•";
      if (bullet.length > 8 || /[\u0000-\u001f]/u.test(bullet))
        throw new Error("Choose a valid bullet symbol.");
      val(lvl, "lvlText", command.kind === "bullet" ? bullet : `%${i + 1}.`);
      val(lvl, "lvlJc", rtl ? "right" : "left");
      const pr = child(lvl, "pPr"),
        tabs = child(pr, "tabs"),
        tab = child(tabs, "tab");
      value(tab, "val", "num");
      value(tab, "pos", 720 * (i + 1));
      const ind = child(pr, "ind");
      value(ind, rtl ? "right" : "left", 720 * (i + 1));
      value(ind, "hanging", 360);
    }
    const num = root.ownerDocument!.createElementNS(
      root.namespaceURI,
      `${root.prefix || "w"}:num`,
    );
    id = next(nums, "numId");
    value(num, "numId", id);
    val(num, "abstractNumId", aId);
    root.appendChild(num);
    changes.set(path, serialize(root));
  }
  val(numPr, "ilvl", level);
  val(numPr, "numId", id!);
  // ilvl precedes numId in CT_NumPr, including when editing an existing list.
  const ilvl = childWord(numPr, "ilvl")!;
  if (ilvl !== numPr.firstChild) numPr.insertBefore(ilvl, numPr.firstChild);
}
