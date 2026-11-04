import { XMLSerializer, type Element } from "@xmldom/xmldom";
import { zipSync, strToU8 } from "fflate";
import type {
  DocumentIR,
  Paragraph,
  ParagraphProperties,
  RunProperties,
} from "../model/document.js";
import { parseDocxPackage } from "./import-docx.js";
import {
  openDocxPackage,
  findMainDocumentPart,
  type DocxPackage,
  type ImportOptions,
} from "./package.js";
import {
  childElements,
  childWord,
  descendants,
  isWord,
  parseXml,
  wordAttribute,
} from "./parser/xml.js";
import { readRelationships } from "./relationships.js";
import {
  setList,
  ensureStyle,
  readStyles,
  editStyle,
} from "./session-parts.js";
import { writeTextEffects, W14 } from "./text-effects.js";

export interface DocxStyle {
  id: string;
  name: string;
  basedOn?: string;
  gallery: boolean;
  run: RunProperties;
  paragraph: ParagraphProperties;
}

export interface RunFormat {
  textOutline?: RunProperties["textOutline"];
  textShadow?: RunProperties["textShadow"];
  glow?: RunProperties["glow"];
  reflection?: RunProperties["reflection"];
  ligatures?: RunProperties["ligatures"];
  numberForm?: RunProperties["numberForm"];
  numberSpacing?: RunProperties["numberSpacing"];
  stylisticSets?: RunProperties["stylisticSets"];
  clear?: boolean;
  bold?: boolean;
  italic?: boolean;
  underline?:
    | "single"
    | "double"
    | "thick"
    | "dotted"
    | "dash"
    | "dotDash"
    | "dotDotDash"
    | "wave"
    | "none";
  strike?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  highlight?: string;
  verticalAlign?: "baseline" | "superscript" | "subscript";
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
  letterSpacingPt?: number;
  positionPt?: number;
  complexScriptFont?: string;
  complexScriptSizePt?: number;
}
export interface ParagraphFormat {
  direction?: "rtl" | "ltr";
  alignment?: "left" | "center" | "right" | "justify";
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  indentStartPt?: number;
  indentEndPt?: number;
  lineSpacing?: number;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  styleId?: string;
  firstLinePt?: number;
  shading?: string;
  borders?: "none" | "all" | "outside" | "top" | "bottom" | "left" | "right";
}
export type EditCommand =
  | {
      type: "text";
      id: string;
      start: number;
      end: number;
      text: string;
      format?: RunFormat;
    }
  | { type: "run"; id: string; format: RunFormat; start?: number; end?: number }
  | { type: "clearRun"; id: string; start?: number; end?: number }
  | { type: "splitParagraph"; id: string; start: number; end: number }
  | { type: "joinParagraph"; id: string; direction: "backward" | "forward" }
  | {
      type: "list";
      id: string;
      kind:
        | "bullet"
        | "decimal"
        | "decimalZero"
        | "lowerLetter"
        | "upperLetter"
        | "upperRoman"
        | "lowerRoman"
        | "none";
      level?: number;
      start?: number;
      bullet?: string;
      restart?: boolean;
      continue?: boolean;
      preserveDefinition?: boolean;
      pattern?: string;
      multilevel?: "decimal" | "mixed" | "bullet";
    }
  | {
      type: "style";
      styleId: string;
      name?: string;
      basedOn?: string;
      gallery?: boolean;
      run?: RunFormat;
      paragraph?: ParagraphFormat;
    }
  | {
      type: "sort";
      ids: string[];
      descending?: boolean;
      numeric?: boolean;
      locale?: string;
    }
  | { type: "paragraph"; id: string; format: ParagraphFormat }
  | {
      type: "insertParagraph";
      id?: string;
      text: string;
      pageBreakBefore?: boolean;
    }
  | { type: "insertTable"; id?: string; rows: number; columns: number }
  | {
      type: "page";
      widthPt: number;
      heightPt: number;
      margins?: { top: number; bottom: number; left: number; right: number };
    };
export interface ParagraphCapability {
  text: boolean;
  format: boolean;
  reason?: string;
}
export interface DocxSession {
  readonly document: DocumentIR;
  readonly dirty: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly canSetPage: boolean;
  readonly lockedReason: string | undefined;
  readonly selection: { id: string; start: number; end: number } | undefined;
  styles(): Promise<DocxStyle[]>;
  capability(id: string): ParagraphCapability;
  execute(command: EditCommand): Promise<DocumentIR>;
  executeBatch(commands: EditCommand[]): Promise<DocumentIR>;
  undo(): Promise<DocumentIR>;
  redo(): Promise<DocumentIR>;
  exportDocx(): Promise<Uint8Array>;
  original(): Uint8Array;
}
type Source = {
  part: string;
  path: number[];
  element: Element;
  paragraph: Paragraph;
};
const serializer = new XMLSerializer();
const XML = "http://www.w3.org/XML/1998/namespace";
const highlights = new Set([
  "none",
  "yellow",
  "green",
  "cyan",
  "magenta",
  "blue",
  "red",
  "darkBlue",
  "darkCyan",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "darkGray",
  "lightGray",
  "black",
]);
const pOrder = [
  "pStyle",
  "keepNext",
  "keepLines",
  "pageBreakBefore",
  "framePr",
  "widowControl",
  "numPr",
  "suppressLineNumbers",
  "pBdr",
  "shd",
  "tabs",
  "suppressAutoHyphens",
  "kinsoku",
  "wordWrap",
  "overflowPunct",
  "topLinePunct",
  "autoSpaceDE",
  "autoSpaceDN",
  "bidi",
  "adjustRightInd",
  "snapToGrid",
  "spacing",
  "ind",
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
  "sectPr",
  "pPrChange",
];
const rOrder = [
  "rStyle",
  "rFonts",
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "noProof",
  "snapToGrid",
  "vanish",
  "webHidden",
  "color",
  "spacing",
  "w",
  "kern",
  "position",
  "sz",
  "szCs",
  "highlight",
  "u",
  "effect",
  "bdr",
  "shd",
  "fitText",
  "vertAlign",
  "rtl",
  "cs",
  "em",
  "lang",
  "eastAsianLayout",
  "specVanish",
  "oMath",
  "rPrChange",
];
function ensure(parent: Element, name: string, order: string[] = []): Element {
  const existing = childWord(parent, name);
  if (existing) return existing;
  const node = parent.ownerDocument!.createElementNS(
    parent.namespaceURI,
    `${parent.prefix || "w"}:${name}`,
  );
  const rank = order.indexOf(name);
  const before = childElements(parent).find(
    (e) =>
      rank >= 0 &&
      (order.indexOf(e.localName!) > rank || e.namespaceURI === W14),
  );
  parent.insertBefore(node, before ?? null);
  return node;
}
function attr(node: Element, key: string, value: string | number) {
  node.setAttributeNS(
    node.namespaceURI,
    `${node.prefix || "w"}:${key}`,
    String(value),
  );
}
function prop(
  parent: Element,
  name: string,
  value: string | number,
  order: string[],
) {
  attr(ensure(parent, name, order), "val", value);
}
function properties(parent: Element, name: "pPr" | "rPr") {
  const node = ensure(parent, name);
  if (node !== parent.firstChild) parent.insertBefore(node, parent.firstChild);
  return node;
}
function finite(value: number, min: number, max: number): number {
  if (!Number.isFinite(value) || value < min || value > max)
    throw new Error(`Value must be between ${min} and ${max}.`);
  return value;
}
function color(value: string) {
  if (!/^#[0-9a-f]{6}$/i.test(value))
    throw new Error("Choose a six-digit RGB color.");
  return value.slice(1).toUpperCase();
}
function sourcePath(element: Element): number[] {
  const path: number[] = [];
  for (
    let e = element;
    e.parentNode?.nodeType === 1;
    e = e.parentNode as Element
  )
    path.unshift(childElements(e.parentNode as Element).indexOf(e));
  return path;
}
function at(root: Element, path: number[]): Element {
  return path.reduce((e, index) => childElements(e)[index], root);
}
function textOf(run: Element): string {
  return childElements(run)
    .map((e) =>
      isWord(e, "t")
        ? (e.textContent ?? "")
        : isWord(e, "tab")
          ? "\t"
          : isWord(e, "br") || isWord(e, "cr")
            ? "\n"
            : "",
    )
    .join("");
}
function plain(p: Element): boolean {
  return childElements(p).every(
    (e) =>
      isWord(e, "pPr") ||
      (isWord(e, "r") &&
        childElements(e).every(
          (c) =>
            isWord(c, "rPr") ||
            isWord(c, "t") ||
            isWord(c, "tab") ||
            isWord(c, "cr") ||
            (isWord(c, "br") && !wordAttribute(c, "type")),
        )),
  );
}
function textContent(p: Element) {
  return childElements(p)
    .filter((e) => isWord(e, "r"))
    .map(textOf)
    .join("");
}
function validText(text: string) {
  if (
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/u.test(text) ||
    /[\uD800-\uDFFF]/u.test(text)
  )
    throw new Error("Text contains characters that cannot be saved in XML.");
  if (text.length > 1_000_000)
    throw new Error("Text is too large for a single edit.");
}
function range(text: string, start: number, end: number) {
  if (
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(end) ||
    start < 0 ||
    end < start ||
    end > text.length
  )
    throw new Error("Invalid text selection.");
  const boundaries = new Set([
    0,
    text.length,
    ...Array.from(
      new Intl.Segmenter(undefined, { granularity: "grapheme" }).segment(text),
      (s) => s.index,
    ),
  ]);
  if (!boundaries.has(start) || !boundaries.has(end))
    throw new Error(
      "Select complete characters, including their combining marks.",
    );
}
function runWithText(base: Element, text: string): Element {
  const run = base.cloneNode(true) as Element;
  for (const child of childElements(run))
    if (!isWord(child, "rPr")) run.removeChild(child);
  for (const piece of text.split(/([\n\t])/)) {
    const tag = piece === "\n" ? "br" : piece === "\t" ? "tab" : "t";
    const node = run.ownerDocument!.createElementNS(
      run.namespaceURI,
      `${run.prefix || "w"}:${tag}`,
    );
    if (tag === "t") {
      node.setAttributeNS(XML, "xml:space", "preserve");
      node.appendChild(run.ownerDocument!.createTextNode(piece));
    }
    run.appendChild(node);
  }
  return run;
}
function setRun(run: Element, format: RunFormat, inheritedColor?: string) {
  if (format.clear) clearRun(run);
  const pr = properties(run, "rPr");
  for (const [key, tags] of [
    ["bold", ["b", "bCs"]],
    ["italic", ["i", "iCs"]],
    ["strike", ["strike"]],
  ] as const)
    if (format[key] !== undefined)
      for (const tag of tags) prop(pr, tag, format[key] ? "1" : "0", rOrder);
  for (const [key, tag] of [
    ["caps", "caps"],
    ["smallCaps", "smallCaps"],
    ["hidden", "vanish"],
  ] as const)
    if (format[key] !== undefined)
      prop(pr, tag, format[key] ? "1" : "0", rOrder);
  if (format.fontFamily !== undefined) {
    if (!format.fontFamily.trim() || format.fontFamily.length > 200)
      throw new Error("Enter a valid font family name.");
    const fonts = ensure(pr, "rFonts", rOrder);
    for (const key of ["ascii", "hAnsi", "eastAsia", "cs"]) {
      attr(fonts, key, format.fontFamily.trim());
      fonts.removeAttributeNS(
        fonts.namespaceURI,
        key === "cs" ? "cstheme" : `${key}Theme`,
      );
    }
  }
  if (format.fontSizePt !== undefined)
    for (const tag of ["sz", "szCs"])
      prop(pr, tag, Math.round(finite(format.fontSizePt, 1, 400) * 2), rOrder);
  if (format.complexScriptFont !== undefined) {
    if (
      !format.complexScriptFont.trim() ||
      format.complexScriptFont.length > 200
    )
      throw new Error("Invalid complex-script font.");
    const fonts = ensure(pr, "rFonts", rOrder);
    attr(fonts, "cs", format.complexScriptFont.trim());
    fonts.removeAttributeNS(fonts.namespaceURI, "cstheme");
  }
  if (format.complexScriptSizePt !== undefined)
    prop(
      pr,
      "szCs",
      Math.round(finite(format.complexScriptSizePt, 1, 400) * 2),
      rOrder,
    );
  if (format.letterSpacingPt !== undefined)
    prop(
      pr,
      "spacing",
      Math.round(finite(format.letterSpacingPt, -20, 100) * 20),
      rOrder,
    );
  if (format.positionPt !== undefined)
    prop(
      pr,
      "position",
      Math.round(finite(format.positionPt, -100, 100) * 2),
      rOrder,
    );
  if (format.color !== undefined) {
    const e = ensure(pr, "color", rOrder);
    attr(e, "val", format.color === "auto" ? "auto" : color(format.color));
    for (const key of ["themeColor", "themeTint", "themeShade"])
      e.removeAttributeNS(e.namespaceURI, key);
  }
  if (format.highlight !== undefined) {
    if (!highlights.has(format.highlight))
      throw new Error("Unsupported highlight color.");
    prop(pr, "highlight", format.highlight, rOrder);
  }
  if (format.underline !== undefined) {
    if (
      ![
        "single",
        "double",
        "thick",
        "dotted",
        "dash",
        "dotDash",
        "dotDotDash",
        "wave",
        "none",
      ].includes(format.underline)
    )
      throw new Error("Unsupported underline.");
    prop(pr, "u", format.underline, rOrder);
  }
  if (format.verticalAlign !== undefined) {
    if (
      !["baseline", "superscript", "subscript"].includes(format.verticalAlign)
    )
      throw new Error("Unsupported script position.");
    prop(pr, "vertAlign", format.verticalAlign, rOrder);
  }
  writeTextEffects(pr, format, inheritedColor);
}
function setParagraph(p: Element, format: ParagraphFormat) {
  const pr = properties(p, "pPr");
  if (format.styleId !== undefined) {
    if (!format.styleId || format.styleId.length > 250)
      throw new Error("Invalid style.");
    prop(pr, "pStyle", format.styleId, pOrder);
    for (const node of childElements(pr))
      if (
        [
          "jc",
          "spacing",
          "ind",
          "keepNext",
          "keepLines",
          "pageBreakBefore",
          "outlineLvl",
          "contextualSpacing",
        ].some((n) => isWord(node, n))
      )
        pr.removeChild(node);
    // Whole-paragraph direct formatting yields to the style. Distinct inline emphasis survives.
    if (plain(p)) {
      const runs = childElements(p).filter((e) => isWord(e, "r"));
      for (const tag of [
        "rFonts",
        "sz",
        "szCs",
        "b",
        "bCs",
        "i",
        "iCs",
        "color",
        "u",
        "strike",
        "caps",
        "smallCaps",
      ]) {
        const props = runs.map((r) => {
          const rp = childWord(r, "rPr");
          return rp && childWord(rp, tag);
        });
        if (
          props.length &&
          props.every(
            (node) =>
              node &&
              serializer.serializeToString(node) ===
                serializer.serializeToString(props[0]!),
          )
        )
          for (const node of props) node!.parentNode!.removeChild(node!);
      }
    }
  }
  if (format.shading !== undefined) {
    const shd = ensure(pr, "shd", pOrder);
    attr(shd, "val", "clear");
    attr(
      shd,
      "fill",
      format.shading === "none" ? "auto" : color(format.shading),
    );
    for (const key of ["themeFill", "themeFillTint", "themeFillShade"])
      shd.removeAttributeNS(shd.namespaceURI, key);
  }
  if (format.firstLinePt !== undefined) {
    const ind = ensure(pr, "ind", pOrder),
      n = finite(format.firstLinePt, -720, 1440);
    ind.removeAttributeNS(ind.namespaceURI, n < 0 ? "firstLine" : "hanging");
    attr(ind, n < 0 ? "hanging" : "firstLine", Math.round(Math.abs(n) * 20));
  }
  if (format.borders !== undefined) {
    if (
      !["none", "all", "outside", "top", "bottom", "left", "right"].includes(
        format.borders,
      )
    )
      throw new Error("Invalid paragraph borders.");
    const borders = ensure(pr, "pBdr", pOrder);
    for (const side of ["top", "left", "bottom", "right"]) {
      const b = ensure(borders, side);
      attr(
        b,
        "val",
        format.borders === "all" ||
          format.borders === "outside" ||
          format.borders === side
          ? "single"
          : "nil",
      );
      attr(b, "sz", 4);
      attr(b, "space", 3);
      attr(b, "color", "000000");
    }
  }
  if (format.direction !== undefined) {
    if (!["ltr", "rtl"].includes(format.direction))
      throw new Error("Invalid paragraph direction.");
    prop(pr, "bidi", format.direction === "rtl" ? "1" : "0", pOrder);
  }
  if (format.alignment !== undefined) {
    if (!["left", "center", "right", "justify"].includes(format.alignment))
      throw new Error("Invalid alignment.");
    prop(
      pr,
      "jc",
      format.alignment === "justify" ? "both" : format.alignment,
      pOrder,
    );
  }
  for (const [key, tag] of [
    ["pageBreakBefore", "pageBreakBefore"],
    ["keepNext", "keepNext"],
    ["keepLines", "keepLines"],
  ] as const)
    if (format[key] !== undefined)
      prop(pr, tag, format[key] ? "1" : "0", pOrder);
  for (const [key, name] of [
    ["spaceBeforePt", "before"],
    ["spaceAfterPt", "after"],
  ] as const)
    if (format[key] !== undefined) {
      const spacing = ensure(pr, "spacing", pOrder);
      attr(spacing, name, Math.round(finite(format[key]!, 0, 1000) * 20));
      spacing.removeAttributeNS(spacing.namespaceURI, `${name}Autospacing`);
      spacing.removeAttributeNS(spacing.namespaceURI, `${name}Lines`);
    }
  if (format.lineSpacing !== undefined) {
    const spacing = ensure(pr, "spacing", pOrder);
    attr(
      spacing,
      "line",
      Math.round(finite(format.lineSpacing, 0.5, 10) * 240),
    );
    attr(spacing, "lineRule", "auto");
  }
  for (const [key, name, physical] of [
    ["indentStartPt", "start", "left"],
    ["indentEndPt", "end", "right"],
  ] as const)
    if (format[key] !== undefined) {
      const indent = ensure(pr, "ind", pOrder);
      attr(indent, name, Math.round(finite(format[key]!, -720, 1440) * 20));
      indent.removeAttributeNS(indent.namespaceURI, physical);
      indent.removeAttributeNS(indent.namespaceURI, `${name}Chars`);
      indent.removeAttributeNS(indent.namespaceURI, `${physical}Chars`);
    }
}
function rewriteText(
  p: Element,
  command: Extract<EditCommand, { type: "text" | "run" | "clearRun" }>,
  effective?: Paragraph,
) {
  const text = textContent(p);
  const start = command.start ?? 0,
    end = command.end ?? text.length;
  range(text, start, end);
  const runs = childElements(p).filter((e) => isWord(e, "r"));
  const base =
    runs[0] ??
    p.ownerDocument!.createElementNS(p.namespaceURI, `${p.prefix || "w"}:r`);
  const output: Element[] = [];
  const colorAt = (position: number) => {
    let offset = 0,
      last = effective?.markProperties?.color;
    for (const inline of effective?.children ?? [])
      if (inline.type === "text") {
        offset += inline.text.length;
        last = inline.properties.color;
        if (position < offset) return last;
      }
    return last;
  };
  let offset = 0,
    inserted = false;
  if (command.type === "text") validText(command.text);
  for (const run of runs) {
    const value = textOf(run),
      next = offset + value.length;
    if (command.type === "run" || command.type === "clearRun") {
      const cuts = [
        ...new Set([
          offset,
          Math.max(offset, Math.min(next, start)),
          Math.max(offset, Math.min(next, end)),
          next,
        ]),
      ].sort((a, b) => a - b);
      for (let i = 0; i < cuts.length - 1; i++) {
        const item = runWithText(
          run,
          value.slice(cuts[i] - offset, cuts[i + 1] - offset),
        );
        if (cuts[i] >= start && cuts[i + 1] <= end) {
          if (command.type === "clearRun") clearRun(item);
          else setRun(item, command.format, colorAt(cuts[i]));
        }
        output.push(item);
      }
      if (!value.length) {
        const item = run.cloneNode(true) as Element;
        if (start === 0 && end === text.length) {
          if (command.type === "clearRun") clearRun(item);
          else setRun(item, command.format, colorAt(offset));
        }
        output.push(item);
      }
    } else {
      if (next <= start) output.push(run.cloneNode(true) as Element);
      else if (offset >= end) {
        if (!inserted) {
          const item = runWithText(run, command.text);
          if (command.format) setRun(item, command.format, colorAt(start));
          output.push(item);
          inserted = true;
        }
        output.push(run.cloneNode(true) as Element);
      } else {
        if (start > offset)
          output.push(runWithText(run, value.slice(0, start - offset)));
        if (!inserted) {
          const item = runWithText(run, command.text);
          if (command.format) setRun(item, command.format, colorAt(start));
          output.push(item);
          inserted = true;
        }
        if (end < next)
          output.push(runWithText(run, value.slice(end - offset)));
      }
    }
    offset = next;
  }
  if (command.type === "text" && !inserted) {
    const item = runWithText(runs.at(-1) ?? base, command.text);
    if (command.format) setRun(item, command.format, colorAt(start));
    output.push(item);
  }
  if (command.type === "run" && !runs.length) {
    const item = runWithText(base, "");
    setRun(item, command.format, colorAt(0));
    output.push(item);
  }
  for (const run of runs) p.removeChild(run);
  for (const run of output) p.appendChild(run);
}
function clearRun(run: Element) {
  const pr = childWord(run, "rPr");
  if (pr)
    for (const child of childElements(pr))
      if (
        (rOrder.includes(child.localName!) ||
          (child.namespaceURI === W14 &&
            [
              "glow",
              "shadow",
              "reflection",
              "textOutline",
              "textFill",
              "ligatures",
              "numForm",
              "numSpacing",
              "stylisticSets",
            ].includes(child.localName!))) &&
        !["rPrChange", "oMath"].includes(child.localName!)
      )
        pr.removeChild(child);
}

/** Retains source parts. Only owned, allowlisted XML is rewritten; all other parts retain their bytes. */
export async function openDocxSession(
  input: ArrayBuffer | Uint8Array,
  options: ImportOptions = {},
): Promise<DocxSession> {
  const original = new Uint8Array(
    input instanceof Uint8Array ? input : new Uint8Array(input),
  ).slice();
  const pkg = await openDocxPackage(original, options),
    main = await findMainDocumentPart(pkg);
  const originalMain = await pkg.readText(main);
  const rels = await readRelationships(pkg, main);
  let lockedReason: string | undefined;
  if (pkg.listParts().some((p) => p.startsWith("_xmlsignatures/")))
    lockedReason =
      "Digitally signed documents are read-only to avoid invalidating signatures.";
  for (const rel of rels.values())
    if (rel.type === "settings" && !rel.external && pkg.hasPart(rel.target)) {
      const protection = descendants(
        parseXml(await pkg.readText(rel.target)).documentElement!,
        "documentProtection",
      )[0];
      if (
        protection &&
        !["0", "false", "off"].includes(
          wordAttribute(protection, "enforcement") ?? "1",
        )
      )
        lockedReason = "This document has editing protection.";
    }
  let states: Map<string, string>[] = [new Map()],
    cursor = 0;
  let identityStates: Map<string, string>[] = [new Map()],
    nextIdentity = 1;
  let selection: { id: string; start: number; end: number } | undefined;
  let sources = new Map<string, Source>(),
    document: DocumentIR;
  function overlay(changes: Map<string, string>): DocxPackage {
    return {
      ...pkg,
      hasPart: (p) => changes.has(p) || pkg.hasPart(p),
      listParts: () => [...new Set([...pkg.listParts(), ...changes.keys()])],
      readText: async (p) => changes.get(p) ?? pkg.readText(p),
      readBinary: async (p) =>
        changes.has(p) ? strToU8(changes.get(p)!) : pkg.readBinary(p),
    };
  }
  async function load(
    changes: Map<string, string>,
    identities: Map<string, string>,
  ) {
    const next = new Map<string, Source>();
    const doc = await parseDocxPackage(
      overlay(changes),
      (part, element, paragraph) => {
        const path = sourcePath(element),
          location = `${part}#${path.join(".")}`,
          id = identities.get(location) ?? `p${nextIdentity++}`;
        identities.set(location, id);
        paragraph.id = id;
        next.set(id, { part, path, element, paragraph });
      },
    );
    sources = next;
    document = doc;
    return doc;
  }
  await load(states[0], identityStates[0]);
  const capability = (id: string): ParagraphCapability => {
    const source = sources.get(id);
    if (lockedReason)
      return { text: false, format: false, reason: lockedReason };
    if (!source)
      return { text: false, format: false, reason: "Select a paragraph." };
    for (
      let e: Element | null = source.element;
      e;
      e = e.parentNode?.nodeType === 1 ? (e.parentNode as Element) : null
    ) {
      if (
        ["sdt", "ins", "del", "moveFrom", "moveTo"].some((name) =>
          isWord(e!, name),
        )
      )
        return {
          text: false,
          format: false,
          reason: "Controlled or tracked content is preserved as read-only.",
        };
    }
    if (
      descendants(source.element, "rPrChange").length ||
      descendants(source.element, "pPrChange").length
    )
      return {
        text: false,
        format: false,
        reason: "Tracked formatting is preserved as read-only.",
      };
    const editable = plain(source.element);
    return {
      text: editable,
      format: editable,
      reason: editable
        ? undefined
        : "This paragraph contains fields, bookmarks, drawings or other structured content. Its text and run formatting are preserved as read-only.",
    };
  };
  let pending = false,
    batching = false;
  async function change(command: EditCommand) {
    if (lockedReason) throw new Error(lockedReason);
    const current = states[cursor],
      updated = new Map(current);
    const source =
      "id" in command && command.id
        ? sources.get(command.id)
        : command.type === "sort"
          ? sources.get(command.ids[0])
          : undefined;
    if ("id" in command && command.id && !source)
      throw new Error(
        "This selection belongs to an older document revision. Select it again.",
      );
    const part = source?.part ?? main;
    const xml = parseXml(await overlay(current).readText(part)),
      root = xml.documentElement!;
    const known = new Map<Element, string>();
    for (const [id, s] of sources)
      if (s.part === part) known.set(at(root, s.path), id);
    let nextSelection = selection;
    const p = source && at(root, source.path);
    if (
      command.type === "run" ||
      command.type === "text" ||
      command.type === "clearRun"
    ) {
      const allowed = capability(command.id);
      if (!allowed.text) throw new Error(allowed.reason);
      rewriteText(p!, command, source?.paragraph);
      nextSelection =
        command.type === "text"
          ? {
              id: command.id,
              start: command.start + command.text.length,
              end: command.start + command.text.length,
            }
          : {
              id: command.id,
              start: command.start ?? 0,
              end: command.end ?? textContent(p!).length,
            };
    } else if (command.type === "splitParagraph") {
      const allowed = capability(command.id);
      if (!allowed.text) throw new Error(allowed.reason);
      const content = textContent(p!);
      range(content, command.start, command.end);
      if (!content.length && source!.paragraph.properties.list) {
        // Enter on an empty list item ends the list rather than adding another empty item.
        await setList(
          p!,
          { type: "list", id: command.id, kind: "none" },
          updated,
          overlay(updated),
          main,
          source!.paragraph.properties.direction === "rtl",
        );
      } else {
        const second = p!.cloneNode(true) as Element;
        rewriteText(p!, {
          type: "text",
          id: command.id,
          start: command.start,
          end: content.length,
          text: "",
        });
        rewriteText(second, {
          type: "text",
          id: command.id,
          start: 0,
          end: command.end,
          text: "",
        });
        // A section boundary belongs to the last of the two resulting paragraphs.
        const firstPr = childWord(p!, "pPr"),
          section = firstPr && childWord(firstPr, "sectPr");
        if (section) firstPr!.removeChild(section);
        p!.parentNode!.insertBefore(second, p!.nextSibling);
        const id = `p${nextIdentity++}`;
        known.set(second, id);
        nextSelection = { id, start: 0, end: 0 };
      }
    } else if (command.type === "joinParagraph") {
      if (!capability(command.id).text)
        throw new Error(capability(command.id).reason);
      const siblings = childElements(p!.parentNode as Element),
        index = siblings.indexOf(p!);
      const other =
        siblings[index + (command.direction === "backward" ? -1 : 1)];
      if (!other || !isWord(other, "p")) return document!;
      const otherId = known.get(other);
      if (!otherId || !capability(otherId).text)
        throw new Error("The adjacent paragraph cannot be joined safely.");
      const first = command.direction === "backward" ? other : p!,
        second = command.direction === "backward" ? p! : other;
      if (
        childWord(childWord(first, "pPr") ?? first, "sectPr") ||
        childWord(childWord(second, "pPr") ?? second, "sectPr")
      )
        throw new Error("Section boundaries cannot be joined.");
      const offset = textContent(first).length;
      for (const run of childElements(second).filter((e) => isWord(e, "r")))
        first.appendChild(run.cloneNode(true));
      second.parentNode!.removeChild(second);
      nextSelection = { id: known.get(first)!, start: offset, end: offset };
    } else if (command.type === "list") {
      if (!source) throw new Error("Select a paragraph.");
      if (!capability(command.id).format)
        throw new Error(capability(command.id).reason);
      const pr = properties(p!, "pPr");
      const inherited = source.paragraph.properties.numbering;
      if (inherited?.id && !childWord(childWord(pr, "numPr") ?? pr, "numId")) {
        const numPr = ensure(pr, "numPr", pOrder);
        prop(numPr, "ilvl", inherited.level ?? 0, ["ilvl", "numId"]);
        prop(numPr, "numId", inherited.id, ["ilvl", "numId"]);
      }
      await setList(
        p!,
        command,
        updated,
        overlay(updated),
        main,
        source.paragraph.properties.direction === "rtl",
      );
      nextSelection = { id: command.id, start: 0, end: 0 };
    } else if (command.type === "sort") {
      const items = command.ids.map((id) => sources.get(id));
      if (
        items.some(
          (s) => !s || s.part !== part || !capability(s.paragraph.id!).text,
        )
      )
        throw new Error("Choose editable paragraphs in the same story.");
      const elements = items.map((s) => at(root, s!.path));
      if (elements.length < 2) return document!;
      const parent = elements[0].parentNode as Element,
        siblings = childElements(parent),
        indices = elements.map((e) => siblings.indexOf(e));
      if (
        elements.some((e) => e.parentNode !== parent) ||
        indices.some((n, i) => i > 0 && n !== indices[i - 1] + 1)
      )
        throw new Error(
          "Sort requires consecutive paragraphs in one container.",
        );
      if (elements.some((e) => childWord(childWord(e, "pPr") ?? e, "sectPr")))
        throw new Error("Sort cannot cross a section boundary.");
      const marker = xml.createComment("sort");
      parent.insertBefore(marker, elements[0]);
      const collator = new Intl.Collator(command.locale, {
        numeric: command.numeric,
        sensitivity: "base",
      });
      const sorted = [...elements].sort(
        (a, b) =>
          collator.compare(textContent(a), textContent(b)) *
          (command.descending ? -1 : 1),
      );
      for (const el of sorted) parent.insertBefore(el, marker);
      parent.removeChild(marker);
    } else if (command.type === "style") {
      await editStyle(
        command,
        updated,
        overlay(updated),
        main,
        setRun,
        setParagraph,
      );
    } else if (command.type === "paragraph") {
      // Source restriction applies even when changing only paragraph properties.
      for (
        let e: Element | null = p!;
        e;
        e = e.parentNode?.nodeType === 1 ? (e.parentNode as Element) : null
      )
        if (
          ["sdt", "ins", "del", "moveFrom", "moveTo"].some((n) => isWord(e!, n))
        )
          throw new Error("Controlled or tracked paragraphs are read-only.");
      setParagraph(p!, command.format);
      if (command.format.styleId)
        await ensureStyle(
          command.format.styleId,
          updated,
          overlay(updated),
          main,
        );
    } else if (command.type === "page") {
      if (descendants(root, "sectPr").length > 1)
        throw new Error(
          "Page setup for multiple sections is not editable yet.",
        );
      const body = childWord(root, "body")!;
      const section = ensure(body, "sectPr"),
        size = ensure(section, "pgSz", [
          "headerReference",
          "footerReference",
          "footnotePr",
          "endnotePr",
          "type",
          "pgSz",
          "pgMar",
          "paperSrc",
          "pgBorders",
          "lnNumType",
          "pgNumType",
          "cols",
        ]);
      attr(size, "w", Math.round(finite(command.widthPt, 72, 31680) * 20));
      attr(size, "h", Math.round(finite(command.heightPt, 72, 31680) * 20));
      attr(
        size,
        "orient",
        command.widthPt > command.heightPt ? "landscape" : "portrait",
      );
      if (command.margins) {
        const m = ensure(section, "pgMar", [
          "headerReference",
          "footerReference",
          "footnotePr",
          "endnotePr",
          "type",
          "pgSz",
          "pgMar",
          "paperSrc",
          "pgBorders",
          "lnNumType",
          "pgNumType",
          "cols",
        ]);
        for (const [side, value] of Object.entries(command.margins))
          attr(
            m,
            side,
            Math.round(
              finite(
                value,
                0,
                Math.min(command.widthPt, command.heightPt) / 2 - 1,
              ) * 20,
            ),
          );
      }
    } else {
      // Insert only at body boundaries; do not rewrite table cells or textbox containers.
      const body = childWord(root, "body");
      if (!body || (p && p.parentNode !== body))
        throw new Error(
          "Select a body paragraph outside a table to insert here.",
        );
      validText(command.type === "insertParagraph" ? command.text : "");
      const next = p?.nextSibling ?? childWord(body, "sectPr") ?? null;
      if (command.type === "insertParagraph") {
        const para = xml.createElementNS(
          root.namespaceURI,
          `${root.prefix || "w"}:p`,
        );
        const run = xml.createElementNS(
          root.namespaceURI,
          `${root.prefix || "w"}:r`,
        );
        para.appendChild(runWithText(run, command.text));
        if (command.pageBreakBefore)
          setParagraph(para, { pageBreakBefore: true });
        body.insertBefore(para, next);
      } else {
        const rows = finite(command.rows, 1, 30),
          columns = finite(command.columns, 1, 12);
        if (!Number.isInteger(rows) || !Number.isInteger(columns))
          throw new Error("Use whole numbers for table dimensions.");
        const table = xml.createElementNS(
          root.namespaceURI,
          `${root.prefix || "w"}:tbl`,
        );
        const pr = ensure(table, "tblPr"),
          width = ensure(pr, "tblW");
        attr(width, "w", 5000);
        attr(width, "type", "pct");
        const borders = ensure(pr, "tblBorders");
        for (const edge of [
          "top",
          "left",
          "bottom",
          "right",
          "insideH",
          "insideV",
        ]) {
          const b = ensure(borders, edge);
          attr(b, "val", "single");
          attr(b, "sz", 4);
          attr(b, "color", "B8C2CF");
        }
        const grid = ensure(table, "tblGrid");
        for (let i = 0; i < columns; i++)
          attr(ensureNew(grid, "gridCol"), "w", Math.round(9360 / columns));
        for (let r = 0; r < rows; r++) {
          const row = ensureNew(table, "tr");
          for (let c = 0; c < columns; c++) {
            const cell = ensureNew(row, "tc");
            ensureNew(cell, "p");
          }
        }
        body.insertBefore(table, next);
        // Word requires a terminal body paragraph after a table in the editing profile.
        if (!next || (next.nodeType === 1 && isWord(next as Element, "sectPr")))
          body.insertBefore(ensureDetached(body, "p"), next);
      }
    }
    const value = serializer.serializeToString(xml);
    if (
      value === (await overlay(current).readText(part)) &&
      [...updated].every(([path, content]) => current.get(path) === content)
    )
      return document!;
    updated.set(part, value);
    const identities = new Map(identityStates[cursor]);
    for (const location of identities.keys())
      if (location.startsWith(`${part}#`)) identities.delete(location);
    for (const [element, id] of known) {
      let top: Element = element;
      while (top.parentNode?.nodeType === 1) top = top.parentNode as Element;
      if (top === root)
        identities.set(`${part}#${sourcePath(element).join(".")}`, id);
    }
    const result = await load(updated, identities); // Commit only if the new source imports successfully.
    states = states.slice(0, cursor + 1);
    states.push(updated);
    cursor++;
    identityStates = identityStates.slice(0, cursor);
    identityStates.push(identities);
    selection = nextSelection;
    // ponytail: bounded snapshot history; use part deltas if large-document editing requires it.
    if (!batching && states.length > 101) {
      states.splice(1, 1);
      identityStates.splice(1, 1);
      cursor--;
    }
    return result;
  }
  async function exclusive(action: () => Promise<DocumentIR>) {
    if (pending)
      throw new Error("Wait for the current document change to finish.");
    pending = true;
    try {
      return await action();
    } finally {
      pending = false;
    }
  }
  return {
    get document() {
      return document!;
    },
    get dirty() {
      return cursor > 0;
    },
    get canUndo() {
      return cursor > 0;
    },
    get canRedo() {
      return cursor < states.length - 1;
    },
    get canSetPage() {
      return (
        !lockedReason &&
        descendants(
          parseXml(states[cursor].get(main) ?? originalMain).documentElement!,
          "sectPr",
        ).length <= 1
      );
    },
    get lockedReason() {
      return lockedReason;
    },
    capability,
    get selection() {
      return selection;
    },
    styles: () => readStyles(overlay(states[cursor]), main),
    execute: (command) => exclusive(() => change(command)),
    executeBatch: (commands) =>
      exclusive(async () => {
        if (commands.length > 500)
          throw new Error("Apply changes to at most 500 paragraphs at a time.");
        const previousStates = states,
          previousIds = identityStates,
          previousCursor = cursor,
          previousSelection = selection;
        batching = true;
        try {
          for (const command of commands) await change(command);
          if (cursor !== previousCursor) {
            const final = states[cursor],
              ids = identityStates[cursor];
            states = [...previousStates.slice(0, previousCursor + 1), final];
            identityStates = [...previousIds.slice(0, previousCursor + 1), ids];
            cursor = previousCursor + 1;
          }
          return document!;
        } catch (error) {
          states = previousStates;
          identityStates = previousIds;
          cursor = previousCursor;
          selection = previousSelection;
          await load(states[cursor], identityStates[cursor]);
          throw error;
        } finally {
          batching = false;
        }
      }),
    undo: () =>
      exclusive(async () => {
        if (cursor === 0) return document!;
        const doc = await load(states[cursor - 1], identityStates[cursor - 1]);
        cursor--;
        selection = undefined;
        return doc;
      }),
    redo: () =>
      exclusive(async () => {
        if (cursor === states.length - 1) return document!;
        const doc = await load(states[cursor + 1], identityStates[cursor + 1]);
        cursor++;
        selection = undefined;
        return doc;
      }),
    original: () => original.slice(),
    async exportDocx() {
      const changes = states[cursor];
      if (!changes.size) return original.slice();
      const entries: Record<string, Uint8Array> = {};
      for (const path of new Set([...pkg.listParts(), ...changes.keys()]))
        entries[path] = changes.has(path)
          ? strToU8(changes.get(path)!)
          : await pkg.readBinary(path);
      return zipSync(entries);
    },
  };
}
function ensureDetached(parent: Element, name: string) {
  return parent.ownerDocument!.createElementNS(
    parent.namespaceURI,
    `${parent.prefix || "w"}:${name}`,
  );
}
function ensureNew(parent: Element, name: string) {
  const node = ensureDetached(parent, name);
  parent.appendChild(node);
  return node;
}

/** Small, valid OPC starting document, independent of the host application. */
export function createBlankDocx(): Uint8Array {
  const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const parts = {
    "[Content_Types].xml":
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    "_rels/.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    "word/_rels/document.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    "word/styles.xml": `<w:styles xmlns:w="${w}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Tahoma"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style></w:styles>`,
    "word/settings.xml": `<w:settings xmlns:w="${w}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
    "word/document.xml": `<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
  };
  return zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([name, value]) => [name, strToU8(value)]),
    ),
  );
}
