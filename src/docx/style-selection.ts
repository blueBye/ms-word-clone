import type { Paragraph, RunProperties } from "../model/document.js";
import type { DocxStyle, RunFormat } from "./session.js";

/** Build a style delta from the actual selected runs, without flattening script-specific defaults. */
export function selectionStyleDelta(
  p: Paragraph,
  start: number,
  end: number,
  style: DocxStyle,
) {
  let offset = 0;
  let caretRange = { start, end };
  const runs: RunProperties[] = [];
  const textRuns = p.children.filter((c) => c.type === "text");
  for (const text of textRuns) {
    const next = offset + text.text.length;
    if (
      start === end
        ? offset <= start && start < next
        : offset < end && next > start
    ) {
      runs.push(text.properties);
      if (start === end) caretRange = { start: offset, end: next };
    }
    offset = next;
  }
  if (!runs.length) {
    runs.push(textRuns.at(-1)?.properties ?? p.markProperties ?? {});
    if (start === end)
      caretRange = {
        start: offset - (textRuns.at(-1)?.text.length ?? 0),
        end: offset,
      };
  }
  const run: RunFormat = {};
  const same = (a: unknown, b: unknown) =>
    JSON.stringify(a) === JSON.stringify(b);
  const copy = (key: keyof RunFormat, get: (r: RunProperties) => unknown) => {
    const value = get(runs[0]);
    if (
      value !== undefined &&
      runs.every((r) => same(get(r), value)) &&
      !same(get(style.run), value)
    )
      Object.assign(run, { [key]: value });
  };
  for (const key of [
    "fontFamily",
    "fontSizePt",
    "underline",
    "color",
    "verticalAlign",
    "letterSpacingPt",
    "positionPt",
    "textOutline",
    "textShadow",
    "glow",
    "reflection",
    "ligatures",
    "numberForm",
    "numberSpacing",
    "stylisticSets",
  ] as const) {
    const defaults: Partial<RunProperties> = {
      underline: "none",
      color: "#000000",
      verticalAlign: "baseline",
      letterSpacingPt: 0,
      positionPt: 0,
      textOutline: null,
      textShadow: null,
      glow: null,
      reflection: null,
    };
    copy(key, (r) => r[key] ?? defaults[key]);
  }
  for (const key of [
    "bold",
    "italic",
    "strike",
    "caps",
    "smallCaps",
    "hidden",
  ] as const)
    copy(key, (r) => !!r[key]);
  for (const [key, slot] of [
    ["complexScriptFont", "fontFamily"],
    ["complexScriptSizePt", "fontSizePt"],
    ["complexScriptBold", "bold"],
    ["complexScriptItalic", "italic"],
  ] as const) {
    copy(
      key,
      (r) =>
        r.complexScript?.[slot] ??
        (slot === "bold" || slot === "italic" ? false : r[slot]),
    );
  }
  const inherit = { ...run };
  for (const [key, slot] of [
    ["complexScriptFont", "fontFamily"],
    ["complexScriptSizePt", "fontSizePt"],
    ["complexScriptBold", "bold"],
    ["complexScriptItalic", "italic"],
  ] as const) {
    // The font/bold toolbar commands write both channels; keep an unchanged channel intact.
    if (run[slot] !== undefined && run[key] === undefined) {
      const value =
        style.run.complexScript?.[slot] ??
        (slot === "bold" || slot === "italic" ? false : style.run[slot]);
      if (value !== undefined) Object.assign(run, { [key]: value });
    }
  }
  // Text matching never promotes a paragraph's local RTL/alignment into the base style.
  // Paragraph style settings are edited explicitly through Modify Style.
  return { run, inherit, range: caretRange };
}
