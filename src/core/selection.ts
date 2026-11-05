import type { Paragraph, RunProperties } from "../model/document.js";

export function runAtPosition(p: Paragraph, position: number, previous = true) {
  const runs = p.children.filter((r) => r.type === "text");
  let offset = 0;
  return (
    runs.find((r) => {
      offset += r.text.length;
      return previous && position > 0 ? offset >= position : offset > position;
    }) ?? runs.at(-1)
  );
}
const signature = (r: RunProperties) => {
  const value: Record<string, unknown> = {
    fontFamily: r.fontFamily ?? "Calibri",
    fontSizePt: r.fontSizePt ?? 12,
    color: (r.color ?? "#000000").toUpperCase(),
    bold: !!r.bold,
    italic: !!r.italic,
    strike: !!r.strike,
    underline: r.underline ?? "none",
    verticalAlign: r.verticalAlign ?? "baseline",
    caps: !!r.caps,
    smallCaps: !!r.smallCaps,
    letterSpacingPt: r.letterSpacingPt ?? 0,
    positionPt: r.positionPt ?? 0,
    complexScriptFont: r.complexScript?.fontFamily ?? r.fontFamily ?? "Calibri",
    complexScriptSize: r.complexScript?.fontSizePt ?? r.fontSizePt ?? 12,
    complexScriptBold: !!r.complexScript?.bold,
    complexScriptItalic: !!r.complexScript?.italic,
    backgroundColor: r.backgroundColor ?? "transparent",
  };
  for (const key of [
    "textOutline",
    "textShadow",
    "glow",
    "reflection",
    "ligatures",
    "numberForm",
    "numberSpacing",
    "stylisticSets",
  ] as const)
    value[key] = r[key] ?? null;
  return JSON.stringify(value);
};
/** Select matching character formatting, including direct formatting within a paragraph. */
export function similarFormatting(
  paragraphs: Paragraph[],
  id: string,
  position: number,
) {
  const selected = paragraphs.find((p) => p.id === id);
  if (!selected) return [];
  const key = signature(
    runAtPosition(selected, position)?.properties ??
      selected.markProperties ??
      {},
  );
  return paragraphs.flatMap((p) => {
    const ranges: { id: string; start: number; end: number }[] = [];
    let offset = 0;
    for (const r of p.children)
      if (r.type === "text") {
        const start = offset;
        offset += r.text.length;
        if (r.text && signature(r.properties) === key) {
          const last = ranges.at(-1);
          if (last?.end === start) last.end = offset;
          else ranges.push({ id: p.id!, start, end: offset });
        }
      }
    return ranges;
  });
}
