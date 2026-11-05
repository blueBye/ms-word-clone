import type {
  Paragraph,
  RunProperties,
  ParagraphProperties,
  TextRun,
} from "../model/document.js";
import type { EditCommand, RunFormat, ParagraphFormat } from "./types.js";
import { defined, mergeRun } from "./format.js";
import { runAtPosition } from "./selection.js";

export function paragraphText(p: Paragraph) {
  return p.children
    .filter((c) => c.type === "text")
    .map((r) => r.text)
    .join("");
}
export function validateRange(text: string, start: number, end: number) {
  if (
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(end) ||
    start < 0 ||
    end < start ||
    end > text.length
  )
    throw new Error("Invalid text selection.");
  const edges = new Set([
    0,
    text.length,
    ...Array.from(
      new Intl.Segmenter(undefined, { granularity: "grapheme" }).segment(text),
      (s) => s.index,
    ),
  ]);
  if (!edges.has(start) || !edges.has(end))
    throw new Error(
      "Select complete characters, including their combining marks.",
    );
}
export function validateText(text: string) {
  if (
    text.length > 1_000_000 ||
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF\uD800-\uDFFF]/u.test(
      text,
    )
  )
    throw new Error("Text contains unsupported characters or is too large.");
}
export function finite(value: number, min: number, max: number) {
  if (!Number.isFinite(value) || value < min || value > max)
    throw new Error(`Value must be between ${min} and ${max}.`);
  return value;
}
const highlights: Record<string, string | undefined> = {
  none: undefined,
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
  lightGray: "#C0C0C0",
  darkGray: "#808080",
  black: "#000000",
  white: "#FFFFFF",
};
export function runProperties(format: RunFormat): RunProperties {
  const {
    clear,
    complexScriptFont,
    complexScriptSizePt,
    complexScriptBold,
    complexScriptItalic,
    highlight,
    ...values
  } = format;
  const result: RunProperties = defined({ ...values });
  if (format.fontFamily !== undefined) {
    if (!format.fontFamily.trim() || format.fontFamily.length > 200)
      throw new Error("Enter a valid font family name.");
    result.fontFamily = format.fontFamily.trim();
    result.fontSlots = {
      ascii: result.fontFamily,
      hAnsi: result.fontFamily,
      eastAsia: result.fontFamily,
    };
    result.complexScript = { fontFamily: result.fontFamily };
  }
  if (format.fontSizePt !== undefined) {
    result.fontSizePt = Math.round(finite(format.fontSizePt, 1, 400) * 2) / 2;
    result.complexScript = {
      ...result.complexScript,
      fontSizePt: result.fontSizePt,
    };
  }
  for (const key of ["bold", "italic"] as const)
    if (format[key] !== undefined)
      result.complexScript = { ...result.complexScript, [key]: format[key] };
  if (complexScriptFont !== undefined) {
    if (!complexScriptFont.trim() || complexScriptFont.length > 200)
      throw new Error("Invalid complex-script font.");
    result.complexScript = {
      ...result.complexScript,
      fontFamily: complexScriptFont.trim(),
    };
  }
  if (complexScriptSizePt !== undefined)
    result.complexScript = {
      ...result.complexScript,
      fontSizePt: Math.round(finite(complexScriptSizePt, 1, 400) * 2) / 2,
    };
  if (complexScriptBold !== undefined)
    result.complexScript = { ...result.complexScript, bold: complexScriptBold };
  if (complexScriptItalic !== undefined)
    result.complexScript = {
      ...result.complexScript,
      italic: complexScriptItalic,
    };
  if (format.color !== undefined) {
    if (format.color !== "auto" && !/^#[\da-f]{6}$/i.test(format.color))
      throw new Error("Choose a six-digit RGB color.");
    result.color =
      format.color === "auto" ? "#000000" : format.color.toUpperCase();
  }
  if (highlight !== undefined) {
    if (!Object.hasOwn(highlights, highlight))
      throw new Error("Unsupported highlight color.");
    result.backgroundColor = highlights[highlight] ?? "transparent";
  }
  if (format.letterSpacingPt !== undefined)
    finite(format.letterSpacingPt, -20, 100);
  if (format.positionPt !== undefined) finite(format.positionPt, -100, 100);
  if (
    format.underline !== undefined &&
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
  if (
    format.verticalAlign !== undefined &&
    !["baseline", "superscript", "subscript"].includes(format.verticalAlign)
  )
    throw new Error("Unsupported script position.");
  for (const [effect, ranges] of [
    [format.textOutline, { widthPt: [0, 20] }],
    [
      format.textShadow,
      {
        blurPt: [0, 100],
        distancePt: [0, 100],
        angle: [0, 360],
        opacity: [0, 1],
      },
    ],
    [format.glow, { radiusPt: [0, 100], opacity: [0, 1] }],
    [
      format.reflection,
      { distancePt: [0, 100], opacity: [0, 1], size: [0, 1] },
    ],
  ] as [
    Record<string, unknown> | null | undefined,
    Record<string, [number, number]>,
  ][]) {
    if (!effect) continue;
    for (const [key, [min, max]] of Object.entries(ranges))
      if (effect[key] !== undefined) finite(effect[key] as number, min, max);
    if (
      effect.color !== undefined &&
      !/^#[\da-f]{6}$/i.test(String(effect.color))
    )
      throw new Error("Choose a six-digit RGB color.");
  }
  if (
    format.stylisticSets?.some(
      (id) => !Number.isInteger(id) || id < 1 || id > 20,
    )
  )
    throw new Error("Stylistic sets are numbered 1 to 20.");
  if (
    format.numberForm !== undefined &&
    !["default", "lining", "oldStyle"].includes(format.numberForm)
  )
    throw new Error("Invalid numeral typography.");
  if (
    format.numberSpacing !== undefined &&
    !["default", "proportional", "tabular"].includes(format.numberSpacing)
  )
    throw new Error("Invalid numeral typography.");
  if (
    format.ligatures !== undefined &&
    ![
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
    ].includes(format.ligatures)
  )
    throw new Error("Invalid ligature setting.");
  return result;
}
export function paragraphProperties(
  format: ParagraphFormat,
): ParagraphProperties {
  const { lineSpacing, shading, borders, ...rest } = format;
  const result: ParagraphProperties = defined(rest);
  if (
    format.styleId !== undefined &&
    (!format.styleId || format.styleId.length > 250)
  )
    throw new Error("Invalid style.");
  for (const key of ["spaceBeforePt", "spaceAfterPt"] as const)
    if (format[key] !== undefined) finite(format[key], 0, 1000);
  for (const key of ["indentStartPt", "indentEndPt", "firstLinePt"] as const)
    if (format[key] !== undefined) finite(format[key], -720, 1440);
  if (lineSpacing !== undefined)
    result.lineSpacing = {
      rule: "auto",
      value: Math.round(finite(lineSpacing, 0.5, 10) * 240) / 240,
    };
  if (
    format.direction !== undefined &&
    !["rtl", "ltr"].includes(format.direction)
  )
    throw new Error("Invalid direction.");
  if (
    format.alignment !== undefined &&
    !["left", "right", "center", "justify"].includes(format.alignment)
  )
    throw new Error("Invalid alignment.");
  if (shading !== undefined) {
    if (shading !== "none" && !/^#[\da-f]{6}$/i.test(shading))
      throw new Error("Choose a six-digit RGB color.");
    result.backgroundColor =
      shading === "none" ? undefined : shading.toUpperCase();
  }
  if (borders !== undefined) {
    if (
      !["none", "all", "outside", "top", "bottom", "left", "right"].includes(
        borders,
      )
    )
      throw new Error("Invalid borders.");
    result.borders = Object.fromEntries(
      ["top", "bottom", "left", "right"].map((side) => [
        side,
        {
          style:
            borders === "all" || borders === "outside" || borders === side
              ? "single"
              : "none",
          widthPt: 0.5,
          color: "#000000",
        },
      ]),
    );
  }
  return result;
}
export function inheritedRun(direct: RunProperties, format: RunFormat) {
  const next = structuredClone(direct);
  for (const key of Object.keys(format)) {
    if (key === "fontFamily") {
      delete next.fontFamily;
      delete next.fontSlots;
    } else if (key.startsWith("complexScript")) {
      const slot = (
        {
          complexScriptFont: "fontFamily",
          complexScriptSizePt: "fontSizePt",
          complexScriptBold: "bold",
          complexScriptItalic: "italic",
        } as const
      )[key as "complexScriptFont"];
      if (next.complexScript) delete next.complexScript[slot];
    } else delete next[key as keyof RunProperties];
  }
  return next;
}
export function rewriteParagraph(
  p: Paragraph,
  command: Extract<EditCommand, { type: "text" | "run" | "clearRun" }>,
  inherit?: RunFormat,
): Paragraph {
  if (p.children.some((r) => r.type !== "text"))
    throw new Error("Structured content is read-only.");
  const text = paragraphText(p),
    start = command.start ?? 0,
    end = command.end ?? text.length;
  validateRange(text, start, end);
  if (command.type === "text") validateText(command.text);
  const output: TextRun[] = [],
    runs = p.children as TextRun[];
  const insertionBase = runAtPosition(p, start, start === end);
  const copy = (base: TextRun, value: string) => ({
    ...base,
    text: value,
    directProperties: structuredClone(base.directProperties ?? base.properties),
  });
  let offset = 0,
    inserted = false;
  for (const base of runs) {
    const next = offset + base.text.length;
    if (command.type !== "text") {
      const cuts = [
        ...new Set([
          offset,
          Math.max(offset, Math.min(next, start)),
          Math.max(offset, Math.min(next, end)),
          next,
        ]),
      ].sort((a, b) => a - b);
      if (!base.text.length) cuts.push(next);
      for (let i = 0; i < cuts.length - 1; i++) {
        const item = copy(
          base,
          base.text.slice(cuts[i] - offset, cuts[i + 1] - offset),
        );
        if (cuts[i] >= start && cuts[i + 1] <= end) {
          if (command.type === "clearRun")
            item.directProperties = inherit
              ? inheritedRun(item.directProperties, inherit)
              : {};
          else
            item.directProperties = mergeRun(
              command.format.clear ? {} : item.directProperties,
              runProperties(command.format),
            );
        }
        output.push(item);
      }
    } else {
      if (next <= start) output.push(copy(base, base.text));
      else if (offset >= end) {
        if (!inserted) {
          const item = copy(insertionBase ?? base, command.text);
          if (command.format)
            item.directProperties = mergeRun(
              command.format.clear ? {} : item.directProperties,
              runProperties(command.format),
            );
          output.push(item);
          inserted = true;
        }
        output.push(copy(base, base.text));
      } else {
        if (start > offset)
          output.push(copy(base, base.text.slice(0, start - offset)));
        if (!inserted) {
          const item = copy(insertionBase ?? base, command.text);
          if (command.format)
            item.directProperties = mergeRun(
              command.format.clear ? {} : item.directProperties,
              runProperties(command.format),
            );
          output.push(item);
          inserted = true;
        }
        if (end < next) output.push(copy(base, base.text.slice(end - offset)));
      }
    }
    offset = next;
  }
  if (command.type === "text" && !inserted) {
    const item = copy(
      runs.at(-1) ?? {
        type: "text",
        text: "",
        properties: {},
        directProperties: {},
      },
      command.text,
    );
    if (command.format)
      item.directProperties = mergeRun(
        command.format.clear ? {} : item.directProperties,
        runProperties(command.format),
      );
    output.push(item);
  }
  const move = (pos: number, trailing = false) =>
    command.type !== "text"
      ? pos
      : pos < start || (pos === start && !trailing)
        ? pos
        : pos >= end
          ? pos + command.text.length - (end - start)
          : start + command.text.length;
  return {
    ...p,
    children: output,
    ...(p.bookmarks
      ? {
          bookmarks: p.bookmarks.map((b) => ({
            ...b,
            start: move(b.start),
            end: move(b.end, true),
          })),
        }
      : {}),
    ...(p.comments
      ? {
          comments: p.comments.map((c) => ({
            ...c,
            start: move(c.start),
            end: move(c.end, true),
          })),
        }
      : {}),
  };
}
