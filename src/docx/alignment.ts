import type { ParagraphAlignment, TextDirection } from "../model/document.js";
/** Logical justification avoids Word's legacy bidi interpretation of left/right. */
export function wordAlignment(
  alignment: ParagraphAlignment,
  direction: TextDirection = "ltr",
) {
  if (alignment === "justify") return "both";
  if (alignment === "left") return direction === "rtl" ? "end" : "start";
  if (alignment === "right") return direction === "rtl" ? "start" : "end";
  return alignment;
}
