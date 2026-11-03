import type { CSSProperties } from "react";
import type { RunProperties } from "../model/document.js";

/** CSS approximations of Word's run effects; these do not change text or document metrics. */
export function textEffectsStyle(p: RunProperties): CSSProperties {
  const rgba = (color: string, opacity: number) =>
    /^#[\da-f]{6}$/i.test(color)
      ? `rgba(${[1, 3, 5].map((i) => parseInt(color.slice(i, i + 2), 16)).join(",")},${Math.max(0, Math.min(1, opacity))})`
      : color;
  const shadows: string[] = [];
  if (p.textShadow) {
    const s = p.textShadow,
      angle = (s.angle * Math.PI) / 180;
    shadows.push(
      `${Math.cos(angle) * s.distancePt}pt ${Math.sin(angle) * s.distancePt}pt ${s.blurPt}pt ${rgba(s.color, s.opacity)}`,
    );
  }
  if (p.glow)
    shadows.push(
      `0 0 ${p.glow.radiusPt}pt ${rgba(p.glow.color, p.glow.opacity)}`,
    );
  const ligatures = p.ligatures;
  const features = [
    ...(ligatures
      ? [
          ["liga", ligatures === "all" || /standard/i.test(ligatures)],
          ["clig", ligatures === "all" || /contextual/i.test(ligatures)],
          ["dlig", ligatures === "all" || /discretional/i.test(ligatures)],
          ["hlig", ligatures === "all" || /historical/i.test(ligatures)],
        ].map(([key, value]) => `"${key}" ${value ? 1 : 0}`)
      : []),
    ...(p.numberForm
      ? [
          `"lnum" ${p.numberForm === "lining" ? 1 : 0}`,
          `"onum" ${p.numberForm === "oldStyle" ? 1 : 0}`,
        ]
      : []),
    ...(p.numberSpacing
      ? [
          `"pnum" ${p.numberSpacing === "proportional" ? 1 : 0}`,
          `"tnum" ${p.numberSpacing === "tabular" ? 1 : 0}`,
        ]
      : []),
    ...(p.stylisticSets
      ? Array.from(
          { length: 20 },
          (_, i) =>
            `"ss${String(i + 1).padStart(2, "0")}" ${p.stylisticSets!.includes(i + 1) ? 1 : 0}`,
        )
      : []),
  ];
  return {
    WebkitTextStroke: p.textOutline
      ? `${p.textOutline.widthPt}pt ${p.textOutline.color}`
      : p.textOutline === null
        ? "0px"
        : undefined,
    paintOrder: "stroke fill",
    textShadow: shadows.length ? shadows.join(", ") : undefined,
    fontFeatureSettings: features.length ? features.join(", ") : undefined,
    // ponytail: CSS reflection approximates Word's alpha ramp; 3D transforms remain source-preserved.
    WebkitBoxReflect: p.reflection
      ? `below ${p.reflection.distancePt}pt linear-gradient(transparent ${100 - p.reflection.size * 100}%,rgba(0,0,0,${p.reflection.opacity}))`
      : undefined,
  };
}
