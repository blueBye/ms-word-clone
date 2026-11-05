import type { ParagraphProperties, RunProperties } from "../model/document.js";
export function physicalAlignment(
  properties: ParagraphProperties,
): ParagraphProperties {
  const a = properties.alignment;
  if (a !== "start" && a !== "end") return properties;
  return {
    ...properties,
    alignment:
      (a === "start") === (properties.direction !== "rtl") ? "left" : "right",
  };
}
export function defined<T extends object>(value: T): T {
  return Object.fromEntries(
    Object.entries(value).filter(([, v]) => v !== undefined),
  ) as T;
}
export function mergeRun(
  base: RunProperties,
  overrides: RunProperties,
  toggle = false,
): RunProperties {
  const result = { ...base, ...defined(overrides) };
  if (base.fontSlots || overrides.fontSlots)
    result.fontSlots = {
      ...base.fontSlots,
      ...defined(overrides.fontSlots ?? {}),
    };
  if (base.complexScript || overrides.complexScript)
    result.complexScript = {
      ...base.complexScript,
      ...defined(overrides.complexScript ?? {}),
    };
  if (toggle) {
    for (const key of [
      "bold",
      "italic",
      "strike",
      "caps",
      "smallCaps",
      "hidden",
    ] as const) {
      if (overrides[key] !== undefined)
        result[key] = overrides[key] ? !base[key] : base[key];
    }
    for (const key of ["bold", "italic"] as const) {
      if (overrides.complexScript?.[key] !== undefined)
        result.complexScript![key] = overrides.complexScript[key]
          ? !base.complexScript?.[key]
          : base.complexScript?.[key];
    }
  }
  return result;
}
export function mergeParagraph(
  base: ParagraphProperties,
  overrides: ParagraphProperties,
): ParagraphProperties {
  const result = { ...base, ...defined(overrides) };
  if (base.numbering || overrides.numbering)
    result.numbering = {
      ...base.numbering,
      ...defined(overrides.numbering ?? {}),
    };
  return result;
}
