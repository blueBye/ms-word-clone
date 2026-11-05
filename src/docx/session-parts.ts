import { builtins } from "../core/style-presets.js";
import { XMLSerializer, type Element } from "@xmldom/xmldom";
import type { DocxPackage } from "./package.js";
import type {
  EditCommand,
  DocxStyle,
  RunFormat,
  ParagraphFormat,
} from "./session.js";
import { Styles, parseTheme } from "./styles.js";
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
export async function related(
  kind: "styles" | "numbering" | "header" | "footer",
  changes: Map<string, string>,
  pkg: DocxPackage,
  main: string,
  reuse = true,
) {
  const relationships = await readRelationships(pkg, main);
  const entry = [...relationships].find(
    ([, r]) => reuse && r.type === kind && !r.external,
  );
  const found = entry?.[1];
  if (found) {
    if (!pkg.hasPart(found.target)) throw new Error(`Missing ${kind} part.`);
    return {
      id: entry![0],
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
    id: `rId${index}`,
    path,
    root: parseXml(
      `<w:${kind === "header" ? "hdr" : kind === "footer" ? "ftr" : kind} xmlns:w="${w}"/>`,
    ).documentElement!,
  };
}

export async function readStyles(pkg: DocxPackage, main: string) {
  const rels = await readRelationships(pkg, main),
    found = [...rels.values()].find((r) => r.type === "styles" && !r.external);
  const themePart = [...rels.values()].find(
    (r) => r.type === "theme" && !r.external,
  );
  const theme = parseTheme(
    themePart && pkg.hasPart(themePart.target)
      ? parseXml(await pkg.readText(themePart.target)).documentElement!
      : undefined,
  );
  const result = new Map<string, DocxStyle>(
    Object.entries(builtins).map(([id, style]) => [
      id,
      {
        id,
        name: style.name,
        gallery: true,
        run: { fontSizePt: style.size ?? 12, bold: !!style.bold },
        paragraph: {
          spaceBeforePt: style.before ?? 0,
          spaceAfterPt: style.after ?? 0,
        },
      },
    ]),
  );
  if (found && pkg.hasPart(found.target)) {
    const root = parseXml(await pkg.readText(found.target)).documentElement!;
    const resolver = new Styles(root, theme, () => {});
    const normal = resolver.paragraph({}).run;
    for (const [id, preset] of Object.entries(builtins)) {
      const style = result.get(id)!;
      style.run = {
        ...normal,
        ...(preset.size !== undefined ? { fontSizePt: preset.size } : {}),
        ...(preset.bold ? { bold: true } : {}),
        complexScript: {
          ...normal.complexScript,
          ...(preset.size !== undefined ? { fontSizePt: preset.size } : {}),
          ...(preset.bold ? { bold: true } : {}),
        },
      };
    }
    for (const style of childElements(root).filter(
      (e) => isWord(e, "style") && wordAttribute(e, "type") === "paragraph",
    )) {
      const id = wordAttribute(style, "styleId");
      if (id) {
        const resolved = resolver.paragraph({ styleId: id });
        result.set(id, {
          id,
          name: wordAttribute(childWord(style, "name") ?? style, "val") ?? id,
          basedOn: wordAttribute(childWord(style, "basedOn") ?? style, "val"),
          gallery:
            !childWord(style, "semiHidden") &&
            (!!childWord(style, "qFormat") ||
              !!childWord(style, "unhideWhenUsed") ||
              id === "Normal"),
          run: resolved.run,
          paragraph: resolved.properties,
        });
      }
    }
  }
  return [...result.values()];
}

/** Edit the existing style definition; unrelated metadata and unknown properties survive. */
export async function editStyle(
  command: Extract<EditCommand, { type: "style" }>,
  changes: Map<string, string>,
  pkg: DocxPackage,
  main: string,
  writeRun: (run: Element, format: RunFormat) => void,
  writeParagraph: (p: Element, format: ParagraphFormat) => void,
) {
  await ensureStyle(command.styleId, changes, pkg, main);
  const view = {
    ...pkg,
    hasPart: (p: string) => changes.has(p) || pkg.hasPart(p),
    readText: async (p: string) => changes.get(p) ?? pkg.readText(p),
  };
  const { path, root } = await related("styles", changes, view, main);
  const style = childElements(root).find(
    (e) =>
      isWord(e, "style") && wordAttribute(e, "styleId") === command.styleId,
  )!;
  if (command.name !== undefined) {
    if (!command.name.trim() || command.name.length > 200)
      throw new Error("Enter a style name between 1 and 200 characters.");
    val(style, "name", command.name.trim());
  }
  if (command.basedOn !== undefined) {
    if (command.basedOn === command.styleId)
      throw new Error("A style cannot be based on itself.");
    let next = command.basedOn;
    const seen = new Set([command.styleId]);
    while (next) {
      if (seen.has(next))
        throw new Error("Style inheritance would create a cycle.");
      seen.add(next);
      const parent = childElements(root).find(
        (e) => isWord(e, "style") && wordAttribute(e, "styleId") === next,
      );
      if (!parent) throw new Error("The base style does not exist.");
      next = wordAttribute(childWord(parent, "basedOn") ?? parent, "val") ?? "";
    }
    if (command.basedOn) val(style, "basedOn", command.basedOn);
    else {
      const base = childWord(style, "basedOn");
      if (base) style.removeChild(base);
    }
  }
  if (command.gallery !== undefined) {
    const existing = childWord(style, "qFormat");
    if (command.gallery) {
      val(style, "qFormat", "1");
      const hidden = childWord(style, "semiHidden");
      if (hidden) style.removeChild(hidden);
    } else {
      if (existing) style.removeChild(existing);
      val(style, "semiHidden", "1");
    }
  }
  if (command.run) {
    const holder = root.ownerDocument!.createElementNS(
      root.namespaceURI,
      `${root.prefix || "w"}:r`,
    );
    const original = childWord(style, "rPr");
    if (original) holder.appendChild(original.cloneNode(true));
    const format = { ...command.run };
    // In style definitions OOXML boolean properties toggle relative to the base style.
    const basedOn = wordAttribute(childWord(style, "basedOn") ?? style, "val");
    const resolver = new Styles(root, { colors: {}, fonts: {} }, () => {});
    const base = basedOn
      ? resolver.paragraph({ styleId: basedOn }).run
      : resolver.defaults.run;
    for (const key of [
      "bold",
      "italic",
      "strike",
      "caps",
      "smallCaps",
      "hidden",
    ] as const)
      if (format[key] !== undefined) format[key] = format[key] !== !!base[key];
    writeRun(holder, format);
    const replacement = childWord(holder, "rPr")!;
    for (const [key, tag] of [
      ["bold", "bCs"],
      ["italic", "iCs"],
    ] as const) {
      const value =
        command.run[
          key === "bold" ? "complexScriptBold" : "complexScriptItalic"
        ] ?? command.run[key];
      if (value !== undefined)
        val(
          replacement,
          tag,
          value !== !!base.complexScript?.[key] ? "1" : "0",
        );
    }
    if (original) style.replaceChild(replacement, original);
    else style.appendChild(replacement);
  }
  if (command.paragraph) {
    const holder = root.ownerDocument!.createElementNS(
      root.namespaceURI,
      `${root.prefix || "w"}:p`,
    );
    const original = childWord(style, "pPr");
    if (original) holder.appendChild(original.cloneNode(true));
    const { styleId: ignored, ...format } = command.paragraph;
    writeParagraph(holder, format);
    const replacement = childWord(holder, "pPr")!;
    if (original) style.replaceChild(replacement, original);
    else style.insertBefore(replacement, childWord(style, "rPr") ?? null);
  }
  // CT_Style metadata precedes paragraph/run properties, even when a menu adds it later.
  const styleOrder = [
    "name",
    "aliases",
    "basedOn",
    "next",
    "link",
    "autoRedefine",
    "hidden",
    "uiPriority",
    "semiHidden",
    "unhideWhenUsed",
    "qFormat",
    "locked",
    "personal",
    "personalCompose",
    "personalReply",
    "rsid",
    "pPr",
    "rPr",
    "tblPr",
    "trPr",
    "tcPr",
    "tblStylePr",
  ];
  const managed = childElements(style)
    .filter((e) => isWord(e, e.localName!) && styleOrder.includes(e.localName!))
    .sort(
      (a, b) =>
        styleOrder.indexOf(a.localName!) - styleOrder.indexOf(b.localName!),
    );
  for (const el of managed) style.appendChild(el);
  changes.set(path, serialize(root));
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
  else if (
    !childElements(root).some(
      (e) =>
        isWord(e, "style") &&
        wordAttribute(e, "type") === "paragraph" &&
        ["1", "true"].includes(wordAttribute(e, "default") ?? ""),
    )
  )
    value(style, "default", "1");
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
  if (preset.size !== undefined) {
    val(r, "sz", preset.size * 2);
    val(r, "szCs", preset.size * 2);
  }
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
  const numPr = child(pPr, "numPr");
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
  if (command.kind === "none") {
    val(numPr, "numId", "0");
    return;
  }
  const formats = [
    "bullet",
    "decimal",
    "decimalZero",
    "lowerLetter",
    "upperLetter",
    "lowerRoman",
    "upperRoman",
  ];
  if (!formats.includes(command.kind))
    throw new Error("Unsupported list format.");
  const level =
    command.level ??
    Number(wordAttribute(childWord(numPr, "ilvl") ?? numPr, "val") ?? 0);
  if (!Number.isInteger(level) || level < 0 || level > 8)
    throw new Error("List level must be between 1 and 9.");
  const start = command.start ?? 1;
  if (!Number.isSafeInteger(start) || start < 1 || start > 1_000_000)
    throw new Error("List start must be a positive whole number.");
  const bullet = command.bullet ?? "\u2022",
    pattern =
      command.pattern?.replace(/%[1-9]/g, `%${level + 1}`) ?? `%${level + 1}.`;
  if (bullet.length > 8 || /[\u0000-\u001f]/u.test(bullet))
    throw new Error("Choose a valid bullet symbol.");
  if (pattern.length > 100 || /[\u0000-\u001f]/u.test(pattern))
    throw new Error("Invalid number format.");
  const { path, root } = await related("numbering", changes, pkg, main);
  const nums = childElements(root).filter((e) => isWord(e, "num"));
  const abstracts = childElements(root).filter((e) => isWord(e, "abstractNum"));
  const next = (items: Element[], key: string) =>
    String(
      Math.max(0, ...items.map((e) => Number(wordAttribute(e, key)) || 0)) + 1,
    );
  const numId = (el: Element | undefined) => {
    const pr = el && childWord(el, "pPr"),
      num = pr && childWord(pr, "numPr");
    return num && wordAttribute(childWord(num, "numId") ?? num, "val");
  };
  const definition = (id: string | undefined) => {
    const num = nums.find((e) => wordAttribute(e, "numId") === id);
    const aId =
      num && wordAttribute(childWord(num, "abstractNumId") ?? num, "val");
    return {
      num,
      abstract: abstracts.find(
        (e) => wordAttribute(e, "abstractNumId") === aId,
      ),
    };
  };
  const own = wordAttribute(childWord(numPr, "numId") ?? numPr, "val");
  const siblings = childElements(p.parentNode as Element),
    index = siblings.indexOf(p);
  const matches = (id: string | undefined) => {
    const a = definition(id).abstract;
    const lvl =
      a &&
      childElements(a).find(
        (e) => isWord(e, "lvl") && wordAttribute(e, "ilvl") === String(level),
      );
    if (!lvl) return false;
    if (command.multilevel)
      return (
        wordAttribute(childWord(a!, "name") ?? a!, "val") ===
        `Docs-${command.multilevel}`
      );
    return (
      wordAttribute(childWord(lvl, "numFmt") ?? lvl, "val") === command.kind &&
      (((command.restart || command.start !== undefined || command.continue) &&
        command.pattern === undefined &&
        command.bullet === undefined) ||
        wordAttribute(childWord(lvl, "lvlText") ?? lvl, "val") ===
          (command.kind === "bullet" ? bullet : pattern))
    );
  };
  let id = command.continue ? undefined : own;
  if (command.continue) {
    for (let i = index - 1; i >= 0; i--) {
      if (!isWord(siblings[i], "p")) break;
      const candidate = numId(siblings[i]);
      if (candidate && matches(candidate)) {
        id = candidate;
        break;
      }
    }
    if (!id)
      throw new Error(
        "No previous list with this format was found in this container.",
      );
  }
  if (
    !command.preserveDefinition &&
    !command.continue &&
    (!matches(id) || command.restart || command.start !== undefined)
  ) {
    const adjacent = [siblings[index - 1], siblings[index + 1]].filter(
      (e) => e && isWord(e, "p"),
    );
    id =
      !command.restart && command.start === undefined
        ? adjacent
            .map(numId)
            .find((candidate) => candidate !== own && matches(candidate))
        : undefined;
    if (!id) {
      let a = definition(own).abstract;
      if (!a || !matches(own)) {
        a = root.ownerDocument!.createElementNS(
          root.namespaceURI,
          `${root.prefix || "w"}:abstractNum`,
        );
        value(a, "abstractNumId", next(abstracts, "abstractNumId"));
        root.insertBefore(a, nums[0] ?? null);
        val(
          a,
          "multiLevelType",
          command.multilevel ? "multilevel" : "hybridMultilevel",
        );
        if (command.multilevel) val(a, "name", `Docs-${command.multilevel}`);
        for (let i = 0; i < 9; i++) {
          const lvl = root.ownerDocument!.createElementNS(
            root.namespaceURI,
            `${root.prefix || "w"}:lvl`,
          );
          value(lvl, "ilvl", i);
          a.appendChild(lvl);
          const kind =
            command.multilevel === "mixed"
              ? ["decimal", "lowerLetter", "lowerRoman"][i % 3]
              : command.multilevel === "bullet"
                ? "bullet"
                : command.kind;
          val(lvl, "start", 1);
          val(lvl, "numFmt", kind);
          const marker =
            command.multilevel === "decimal"
              ? Array.from({ length: i + 1 }, (_, n) => `%${n + 1}`).join(".") +
                "."
              : kind === "bullet"
                ? command.multilevel === "bullet"
                  ? ["\u2022", "\u25cb", "\u25aa"][i % 3]
                  : bullet
                : command.pattern
                  ? command.pattern.replace(/%[1-9]/g, `%${i + 1}`)
                  : `%${i + 1}.`;
          val(lvl, "lvlText", marker);
          val(lvl, "lvlJc", rtl ? "right" : "left");
          val(lvl, "suff", "tab");
          const pr = child(lvl, "pPr"),
            tab = child(child(pr, "tabs"), "tab");
          value(tab, "val", "num");
          value(tab, "pos", 720 * (i + 1));
          const ind = child(pr, "ind");
          value(ind, rtl ? "right" : "left", 720 * (i + 1));
          value(ind, "hanging", 360);
        }
      }
      const num = root.ownerDocument!.createElementNS(
        root.namespaceURI,
        `${root.prefix || "w"}:num`,
      );
      id = next(nums, "numId");
      value(num, "numId", id);
      val(num, "abstractNumId", wordAttribute(a, "abstractNumId")!);
      if (command.restart || command.start !== undefined) {
        const override = child(num, "lvlOverride");
        value(override, "ilvl", level);
        val(override, "startOverride", start);
      }
      root.appendChild(num);
      changes.set(path, serialize(root));
    }
  }
  if (!id || (command.preserveDefinition && !definition(id).num))
    throw new Error("This paragraph does not have a list definition.");
  val(numPr, "ilvl", level);
  val(numPr, "numId", id);
  const ilvl = childWord(numPr, "ilvl")!;
  if (ilvl !== numPr.firstChild) numPr.insertBefore(ilvl, numPr.firstChild);
  const ind = childWord(pPr, "ind");
  if (ind) pPr.removeChild(ind);
}
