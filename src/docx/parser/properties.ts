import type { Element } from "@xmldom/xmldom";
import { parseTextEffects } from "../text-effects.js";
import type {
  Borders,
  ParagraphProperties,
  RunProperties,
} from "../../model/document.js";
import {
  childElements,
  childWord,
  halfPoints,
  numericAttribute,
  onOff,
  wordAttribute,
  wordValue,
} from "./xml.js";

export interface Theme {
  colors: Record<string, string>;
  fonts: Record<string, string>;
}
export function colorValue(
  node: Element | undefined,
  attribute = "val",
  theme?: Theme,
): string | undefined {
  if (!node) return undefined;
  const key = wordAttribute(
    node,
    attribute === "fill" ? "themeFill" : "themeColor",
  );
  let color = (key && theme?.colors[key]) || wordAttribute(node, attribute);
  if (color === "auto") return undefined;
  if (!color || !/^[\da-f]{6}$/i.test(color)) return undefined;
  const shade = wordAttribute(
    node,
    attribute === "fill" ? "themeFillShade" : "themeShade",
  );
  const tint = wordAttribute(
    node,
    attribute === "fill" ? "themeFillTint" : "themeTint",
  );
  // The explicit RGB value is the producer's cached result; use it when no theme is available.
  if (key && theme?.colors[key]) {
    color = color
      .match(/../g)!
      .map((component) => {
        let value = parseInt(component, 16);
        if (shade && /^[\da-f]{2}$/i.test(shade))
          value *= parseInt(shade, 16) / 255;
        if (tint && /^[\da-f]{2}$/i.test(tint))
          value += (255 - value) * (1 - parseInt(tint, 16) / 255);
        return Math.round(value).toString(16).padStart(2, "0");
      })
      .join("");
  }
  return `#${color.toUpperCase()}`;
}
export function parseBorders(
  node: Element | undefined,
  theme?: Theme,
): Borders | undefined {
  if (!node) return undefined;
  const borders: Borders = {};
  for (const child of childElements(node)) {
    const key =
      child.localName === "start"
        ? "left"
        : child.localName === "end"
          ? "right"
          : child.localName;
    if (
      !key ||
      !["top", "left", "bottom", "right", "insideH", "insideV"].includes(key)
    )
      continue;
    borders[key as keyof Borders] = {
      style: wordAttribute(child, "val") ?? "single",
      widthPt: (numericAttribute(child, "sz") ?? 4) / 8,
      color: colorValue(child, "color", theme) ?? "currentColor",
      spacePt: numericAttribute(child, "space"),
    };
  }
  return borders;
}
export function defined<T extends object>(value: T): T {
  return Object.fromEntries(
    Object.entries(value).filter(([, v]) => v !== undefined),
  ) as T;
}
export { mergeRun, mergeParagraph } from "../../core/format.js";
function themeFont(
  theme: Theme | undefined,
  key: string | undefined,
  language?: string,
): string | undefined {
  if (!key || !theme) return undefined;
  const family = key.startsWith("major") ? "major" : "minor";
  const slot = key.endsWith("Bidi")
    ? "cs"
    : key.endsWith("EastAsia")
      ? "eastAsia"
      : "latin";
  const script = /^(ar|fa|ur)/i.test(language ?? "")
    ? "Arab"
    : /^he/i.test(language ?? "")
      ? "Hebr"
      : "";
  return (
    theme.fonts[`${family}.${slot}`] ||
    (slot === "cs" && script ? theme.fonts[`${family}.${script}`] : undefined)
  );
}
export function parseRunPropertyElement(
  rPr: Element | undefined,
  theme?: Theme,
): RunProperties {
  if (!rPr) return {};
  const fonts = childWord(rPr, "rFonts");
  const lang = childWord(rPr, "lang");
  const language = lang && wordAttribute(lang, "val");
  const csLanguage = lang && wordAttribute(lang, "bidi");
  const font = (slot: string, themed: string, complex = false) =>
    fonts &&
    (themeFont(
      theme,
      wordAttribute(fonts, themed),
      complex ? csLanguage : language,
    ) ??
      wordAttribute(fonts, slot));
  const slots = defined({
    ascii: font("ascii", "asciiTheme"),
    hAnsi: font("hAnsi", "hAnsiTheme"),
    eastAsia: font("eastAsia", "eastAsiaTheme"),
  });
  const cs = defined({
    fontFamily: font("cs", "cstheme", true),
    fontSizePt: halfPoints(childWord(rPr, "szCs")),
    bold: onOff(childWord(rPr, "bCs")),
    italic: onOff(childWord(rPr, "iCs")),
  });
  const rtl = onOff(childWord(rPr, "rtl"));
  const underline =
    wordValue(rPr, "u") ?? (childWord(rPr, "u") ? "single" : undefined);
  const highlights: Record<string, string> = {
    yellow: "#FFFF00",
    green: "#00FF00",
    cyan: "#00FFFF",
    magenta: "#FF00FF",
    blue: "#0000FF",
    red: "#FF0000",
    darkBlue: "#000080",
    darkCyan: "#008080",
    darkGreen: "#008000",
    darkMagenta: "#800080",
    darkRed: "#800000",
    darkYellow: "#808000",
    darkGray: "#808080",
    lightGray: "#C0C0C0",
    black: "#000000",
    white: "#FFFFFF",
    none: "transparent",
  };
  const verticalAlign = wordValue(rPr, "vertAlign");
  const effects = parseTextEffects(rPr, theme?.colors);
  return defined({
    ...(onOff(childWord(rPr, "outline"))
      ? {
          textOutline: {
            color:
              colorValue(childWord(rPr, "color"), "val", theme) ?? "#000000",
            widthPt: 0.5,
          },
        }
      : {}),
    ...(onOff(childWord(rPr, "shadow"))
      ? {
          textShadow: {
            color: "#000000",
            blurPt: 1,
            distancePt: 1,
            angle: 45,
            opacity: 0.4,
          },
        }
      : {}),
    ...effects,
    styleId: wordValue(rPr, "rStyle"),
    fontFamily: slots.ascii ?? slots.hAnsi,
    fontSlots: Object.keys(slots).length ? slots : undefined,
    fontHint: fonts && wordAttribute(fonts, "hint"),
    fontSizePt: halfPoints(childWord(rPr, "sz")),
    bold: onOff(childWord(rPr, "b")),
    italic: onOff(childWord(rPr, "i")),
    direction: rtl === undefined ? undefined : rtl ? "rtl" : "ltr",
    complexScript: Object.keys(cs).length ? cs : undefined,
    forceComplexScript: onOff(childWord(rPr, "cs")),
    language,
    complexScriptLanguage: csLanguage,
    color:
      effects.color ??
      (wordValue(rPr, "color") === "auto"
        ? "#000000"
        : colorValue(childWord(rPr, "color"), "val", theme)),
    backgroundColor:
      wordValue(rPr, "highlight") === "none"
        ? (colorValue(childWord(rPr, "shd"), "fill", theme) ?? "transparent")
        : (highlights[wordValue(rPr, "highlight") ?? ""] ??
          colorValue(childWord(rPr, "shd"), "fill", theme)),
    underline,
    strike: onOff(childWord(rPr, "strike")),
    verticalAlign: ["baseline", "superscript", "subscript"].includes(
      verticalAlign ?? "",
    )
      ? (verticalAlign as RunProperties["verticalAlign"])
      : undefined,
    positionPt:
      numericAttribute(childWord(rPr, "position"), "val") === undefined
        ? undefined
        : numericAttribute(childWord(rPr, "position"), "val")! / 2,
    letterSpacingPt:
      numericAttribute(childWord(rPr, "spacing"), "val") === undefined
        ? undefined
        : numericAttribute(childWord(rPr, "spacing"), "val")! / 20,
    caps: onOff(childWord(rPr, "caps")),
    smallCaps: onOff(childWord(rPr, "smallCaps")),
    hidden: onOff(childWord(rPr, "vanish")),
  });
}
export function parseParagraphProperties(
  pPr: Element | undefined,
  theme?: Theme,
): ParagraphProperties {
  if (!pPr) return {};
  const bidi = onOff(childWord(pPr, "bidi"));
  const jc = wordValue(pPr, "jc");
  const spacing = childWord(pPr, "spacing");
  const ind = childWord(pPr, "ind");
  const numPr = childWord(pPr, "numPr");
  const line = numericAttribute(spacing, "line");
  const rule = (spacing && wordAttribute(spacing, "lineRule")) || "auto";
  const points = (element: Element | undefined, name: string) => {
    const n = numericAttribute(element, name);
    return n === undefined ? undefined : n / 20;
  };
  const first = points(ind, "firstLine");
  const hanging = points(ind, "hanging");
  return defined({
    styleId: wordValue(pPr, "pStyle"),
    direction: bidi === undefined ? undefined : bidi ? "rtl" : "ltr",
    alignment:
      (jc &&
        (
          {
            both: "justify",
            distribute: "justify",
            left: "left",
            right: "right",
            center: "center",
            start: "start",
            end: "end",
          } as const
        )[jc as "both"]) ||
      undefined,
    spaceBeforePt: points(spacing, "before"),
    spaceAfterPt: points(spacing, "after"),
    lineSpacing:
      line === undefined
        ? undefined
        : {
            value: rule === "auto" ? line / 240 : line / 20,
            rule: rule as "auto" | "exact" | "atLeast",
          },
    indentLeftPt: points(ind, "left"),
    indentRightPt: points(ind, "right"),
    indentStartPt: points(ind, "start"),
    indentEndPt: points(ind, "end"),
    firstLinePt: hanging === undefined ? first : -hanging,
    tabs:
      childWord(pPr, "tabs") &&
      childElements(childWord(pPr, "tabs")!).map((tab) => ({
        positionPt: points(tab, "pos") ?? 0,
        alignment: wordAttribute(tab, "val") ?? "left",
        leader: wordAttribute(tab, "leader"),
      })),
    numbering: numPr
      ? defined({
          id: wordValue(numPr, "numId"),
          level: numericAttribute(childWord(numPr, "ilvl"), "val"),
        })
      : undefined,
    keepNext: onOff(childWord(pPr, "keepNext")),
    keepLines: onOff(childWord(pPr, "keepLines")),
    suppressLineNumbers: onOff(childWord(pPr, "suppressLineNumbers")),
    ...(childWord(pPr, "framePr") &&
    ["drop", "margin"].includes(
      wordAttribute(childWord(pPr, "framePr")!, "dropCap") ?? "",
    )
      ? {
          dropCap: {
            lines: numericAttribute(childWord(pPr, "framePr"), "lines") ?? 3,
            distancePt:
              (numericAttribute(childWord(pPr, "framePr"), "hSpace") ?? 72) /
              20,
            kind: wordAttribute(childWord(pPr, "framePr")!, "dropCap") as
              | "drop"
              | "margin",
          },
        }
      : {}),
    pageBreakBefore: onOff(childWord(pPr, "pageBreakBefore")),
    contextualSpacing: onOff(childWord(pPr, "contextualSpacing")),
    backgroundColor: colorValue(childWord(pPr, "shd"), "fill", theme),
    borders: parseBorders(childWord(pPr, "pBdr"), theme),
  });
}
