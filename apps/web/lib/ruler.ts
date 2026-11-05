export type RulerHandle = "first" | "hanging" | "start" | "end";
export const RULER_STEP_PT = 4.5; // 1/16 inch, matching the small visible ticks.
export const snapRuler = (value: number) =>
  Math.round(value / RULER_STEP_PT) * RULER_STEP_PT;
export function rulerChange(
  handle: RulerHandle,
  position: number,
  start: number,
  first: number,
  width: number,
) {
  const value =
    Math.round(Math.max(-72, Math.min(width - 12, position)) * 2) / 2;
  if (handle === "first") return { firstLinePt: value - start };
  if (handle === "hanging")
    return { indentStartPt: value, firstLinePt: start + first - value };
  if (handle === "end") return { indentEndPt: value };
  return { indentStartPt: value };
}
