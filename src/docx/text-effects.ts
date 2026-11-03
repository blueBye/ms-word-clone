import type { Element } from "@xmldom/xmldom";
import type { RunProperties } from "../model/document.js";
import { childElements, childWord, wordAttribute } from "./parser/xml.js";

// MS-DOCX §2.2.1 and §5.1: Office 2010 run properties, qualified elements AND attributes.
export const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XMLNS = "http://www.w3.org/2000/xmlns/";
const order = [
  "glow",
  "shadow",
  "reflection",
  "textOutline",
  "textFill",
  "scene3d",
  "props3d",
  "ligatures",
  "numForm",
  "numSpacing",
  "stylisticSets",
  "cntxtAlts",
];
const find = (parent: Element, name: string) =>
  childElements(parent).find(
    (e) => e.namespaceURI === W14 && e.localName === name,
  );
const attribute = (node: Element | undefined, name: string) =>
  node?.getAttributeNS(W14, name) ?? undefined;
const number = (node: Element | undefined, name: string, fallback = 0) => {
  const raw = attribute(node, name);
  const value = raw === undefined ? fallback : Number(raw);
  return Number.isFinite(value) ? value : fallback;
};
export function parseTextEffects(
  pr: Element,
  colors: Record<string, string> = {},
): Partial<RunProperties> {
  const result: Partial<RunProperties> = {};
  const color = (node: Element) => {
    const rgb = find(node, "srgbClr"),
      scheme = find(node, "schemeClr");
    const cached =
      attribute(rgb, "val") ?? colors[attribute(scheme, "val") ?? ""];
    if (!cached || !/^[\da-f]{6}$/i.test(cached)) return "#000000";
    const channels = [0, 2, 4].map((i) => parseInt(cached.slice(i, i + 2), 16));
    for (const transform of childElements(rgb ?? scheme ?? node)) {
      if (transform.namespaceURI !== W14) continue;
      const factor = Math.max(
        0,
        Math.min(1, number(transform, "val", 100000) / 100000),
      );
      for (let i = 0; i < 3; i++) {
        if (transform.localName === "tint")
          channels[i] = channels[i] * factor + 255 * (1 - factor);
        if (transform.localName === "shade") channels[i] *= factor;
      }
    }
    return (
      "#" +
      channels
        .map((c) => Math.round(c).toString(16).padStart(2, "0"))
        .join("")
        .toUpperCase()
    );
  };
  const alpha = (node: Element) =>
    number(
      find(find(node, "srgbClr") ?? find(node, "schemeClr") ?? node, "alpha"),
      "val",
      100000,
    ) / 100000;
  const outline = find(pr, "textOutline");
  if (outline) {
    const fill = find(outline, "solidFill");
    result.textOutline = find(outline, "noFill")
      ? null
      : {
          color: fill ? color(fill) : "#000000",
          widthPt: number(outline, "w") / 12700,
        };
  }
  const shadow = find(pr, "shadow");
  if (shadow)
    result.textShadow =
      alpha(shadow) === 0
        ? null
        : {
            color: color(shadow),
            blurPt: number(shadow, "blurRad") / 12700,
            distancePt: number(shadow, "dist") / 12700,
            angle: number(shadow, "dir") / 60000,
            opacity: alpha(shadow),
          };
  const glow = find(pr, "glow");
  if (glow)
    result.glow =
      number(glow, "rad") === 0
        ? null
        : {
            color: color(glow),
            radiusPt: number(glow, "rad") / 12700,
            opacity: alpha(glow),
          };
  const reflection = find(pr, "reflection");
  if (reflection)
    result.reflection =
      number(reflection, "stA") === 0
        ? null
        : {
            distancePt: number(reflection, "dist") / 12700,
            opacity: number(reflection, "stA") / 100000,
            size: number(reflection, "endPos", 50000) / 100000,
          };
  const fill = find(pr, "textFill"),
    solid = fill && find(fill, "solidFill");
  if (solid) result.color = color(solid);
  const ligatures = find(pr, "ligatures");
  if (ligatures) result.ligatures = attribute(ligatures, "val");
  const form = attribute(find(pr, "numForm"), "val");
  if (["default", "lining", "oldStyle"].includes(form ?? ""))
    result.numberForm = form as RunProperties["numberForm"];
  const spacing = attribute(find(pr, "numSpacing"), "val");
  if (["default", "proportional", "tabular"].includes(spacing ?? ""))
    result.numberSpacing = spacing as RunProperties["numberSpacing"];
  const sets = find(pr, "stylisticSets");
  if (sets)
    result.stylisticSets = childElements(sets)
      .filter(
        (e) =>
          e.namespaceURI === W14 &&
          e.localName === "styleSet" &&
          !["0", "false", "off"].includes(attribute(e, "val") ?? "1"),
      )
      .map((e) => number(e, "id"))
      .filter((n) => n >= 1 && n <= 20);
  return result;
}
export function writeTextEffects(
  pr: Element,
  format: Partial<RunProperties>,
  inheritedColor?: string,
) {
  const needsFill =
    !!format.textOutline ||
    (format.color !== undefined &&
      (!!find(pr, "textOutline") || !!find(pr, "textFill")));
  // Explicit font color replaces an imported effect fill, including gradient fills.
  if (format.color !== undefined) {
    const fill = find(pr, "textFill");
    if (fill) pr.removeChild(fill);
  }
  const keys = [
    "textOutline",
    "textShadow",
    "glow",
    "reflection",
    "ligatures",
    "numberForm",
    "numberSpacing",
    "stylisticSets",
  ] as const;
  if (!keys.some((key) => format[key] !== undefined) && !needsFill) return;
  const root = pr.ownerDocument!.documentElement!;
  if (root.lookupNamespaceURI("w14") && root.lookupNamespaceURI("w14") !== W14)
    throw new Error("The w14 namespace prefix is already in use.");
  root.setAttributeNS(XMLNS, "xmlns:w14", W14);
  let mc = root.lookupPrefix(MC) ?? "mc";
  if (root.lookupNamespaceURI(mc) && root.lookupNamespaceURI(mc) !== MC)
    mc = "editorMc";
  root.setAttributeNS(XMLNS, `xmlns:${mc}`, MC);
  const ignored = new Set(
    (root.getAttributeNS(MC, "Ignorable") ?? "").split(/\s+/).filter(Boolean),
  );
  ignored.add("w14");
  root.setAttributeNS(MC, `${mc}:Ignorable`, [...ignored].join(" "));
  const node = (parent: Element, name: string) => {
    const el = pr.ownerDocument!.createElementNS(W14, `w14:${name}`);
    parent.appendChild(el);
    return el;
  };
  const attr = (el: Element, name: string, value: string | number) =>
    el.setAttributeNS(W14, `w14:${name}`, String(value));
  const bounded = (n: number, min: number, max: number) => {
    if (!Number.isFinite(n) || n < min || n > max)
      throw new Error("Text effect measurement is out of range.");
    return n;
  };
  const color = (el: Element, value: string, opacity = 1) => {
    if (!/^#[\da-f]{6}$/i.test(value))
      throw new Error("Choose an RGB effect color.");
    const rgb = node(el, "srgbClr");
    attr(rgb, "val", value.slice(1).toUpperCase());
    // Word drops fully opaque outline/fill colors when an explicit alpha=100000 is present.
    // Omit the identity transform, matching Word-authored solid fills.
    const alpha = Math.round(bounded(opacity, 0, 1) * 100000);
    if (alpha !== 100000) attr(node(rgb, "alpha"), "val", alpha);
  };
  const replace = (name: string) => {
    const old = find(pr, name);
    if (old) pr.removeChild(old);
    const el = pr.ownerDocument!.createElementNS(W14, `w14:${name}`);
    const before = childElements(pr).find(
      (e) =>
        (e.namespaceURI === W14 &&
          order.indexOf(e.localName!) > order.indexOf(name)) ||
        e.localName === "rPrChange",
    );
    pr.insertBefore(el, before ?? null);
    return el;
  };
  if (format.textOutline !== undefined) {
    const el = replace("textOutline"),
      value = format.textOutline;
    attr(el, "w", Math.round(bounded(value?.widthPt ?? 0, 0, 20) * 12700));
    attr(el, "cap", "rnd");
    attr(el, "cmpd", "sng");
    attr(el, "algn", "ctr");
    if (value) color(node(el, "solidFill"), value.color);
    else node(el, "noFill");
    attr(node(el, "prstDash"), "val", "solid");
    node(el, "round");
  }
  // Word's effect outline uses the DrawingML text fill, rather than w:color alone.
  if (needsFill && (format.color !== undefined || !find(pr, "textFill"))) {
    const fill = node(replace("textFill"), "solidFill");
    const current =
      format.color ??
      wordAttribute(childWord(pr, "color") ?? pr, "val") ??
      inheritedColor;
    if (current && /^#?[\da-f]{6}$/i.test(current))
      color(fill, current.startsWith("#") ? current : `#${current}`);
    else attr(node(fill, "schemeClr"), "val", "tx1");
  }
  if (format.textShadow !== undefined) {
    const el = replace("shadow"),
      value = format.textShadow;
    attr(
      el,
      "blurRad",
      Math.round(bounded(value?.blurPt ?? 0, 0, 100) * 12700),
    );
    attr(
      el,
      "dist",
      Math.round(bounded(value?.distancePt ?? 0, 0, 100) * 12700),
    );
    attr(
      el,
      "dir",
      Math.round((bounded(value?.angle ?? 0, 0, 360) % 360) * 60000),
    );
    attr(el, "sx", 100000);
    attr(el, "sy", 100000);
    attr(el, "algn", "ctr");
    color(el, value?.color ?? "#000000", value?.opacity ?? 0);
  }
  if (format.glow !== undefined) {
    const el = replace("glow"),
      value = format.glow;
    attr(el, "rad", Math.round(bounded(value?.radiusPt ?? 0, 0, 100) * 12700));
    color(el, value?.color ?? "#000000", value?.opacity ?? 0);
  }
  if (format.reflection !== undefined) {
    const el = replace("reflection"),
      value = format.reflection;
    for (const [key, n] of Object.entries({
      blurRad: 6350,
      stA: Math.round(bounded(value?.opacity ?? 0, 0, 1) * 100000),
      stPos: 0,
      endA: 0,
      endPos: Math.round(bounded(value?.size ?? 0, 0, 1) * 100000),
      dist: Math.round(bounded(value?.distancePt ?? 0, 0, 100) * 12700),
      dir: 5400000,
      fadeDir: 5400000,
      sx: 100000,
      sy: -100000,
      algn: "bl",
    }))
      attr(el, key, n);
  }
  if (format.ligatures !== undefined) {
    const choices = [
      "none",
      "standard",
      "contextual",
      "historical",
      "discretional",
      "standardContextual",
      "standardHistorical",
      "contextualHistorical",
      "standardDiscretional",
      "contextualDiscretional",
      "historicalDiscretional",
      "standardContextualHistorical",
      "standardContextualDiscretional",
      "standardHistoricalDiscretional",
      "contextualHistoricalDiscretional",
      "all",
    ];
    if (!choices.includes(format.ligatures))
      throw new Error("Invalid ligature setting.");
    attr(replace("ligatures"), "val", format.ligatures);
  }
  for (const [key, name, choices] of [
    ["numberForm", "numForm", ["default", "lining", "oldStyle"]],
    ["numberSpacing", "numSpacing", ["default", "proportional", "tabular"]],
  ] as const) {
    const value = format[key];
    if (value !== undefined) {
      if (!(choices as readonly string[]).includes(value))
        throw new Error("Invalid numeral typography.");
      attr(replace(name), "val", value);
    }
  }
  if (format.stylisticSets !== undefined) {
    const sets = replace("stylisticSets");
    for (const id of [...new Set(format.stylisticSets)]) {
      if (!Number.isInteger(id) || id < 1 || id > 20)
        throw new Error("Stylistic sets are numbered 1 to 20.");
      const set = node(sets, "styleSet");
      attr(set, "id", id);
      attr(set, "val", "1");
    }
  }
}
