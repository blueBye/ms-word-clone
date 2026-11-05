import type { DrawingInline } from "../model/document.js";
import { finite, validateText } from "./text.js";
import { imageData, modelData, binaryData, videoUrl } from "./assets.js";
import { parseEquation } from "./equation.js";
export function validateDrawing(d: DrawingInline, depth = 0): DrawingInline {
  if (depth > 8) throw new Error("Object groups are nested too deeply.");
  finite(d.widthPt, 1, 15840);
  finite(d.heightPt, 1, 15840);
  validateText(d.alt);
  if (d.image) {
    imageData(d.image.dataUrl);
    if (d.image.opacity !== undefined) finite(d.image.opacity, 0, 1);
    if (d.image.brightness !== undefined) finite(d.image.brightness, -1, 1);
    if (d.image.contrast !== undefined) finite(d.image.contrast, -1, 1);
    if (d.image.blurPt !== undefined) finite(d.image.blurPt, 0, 30);
    for (const value of Object.values(d.image.crop ?? {}))
      finite(value, 0, 0.95);
    if (
      (d.image.crop?.left ?? 0) + (d.image.crop?.right ?? 0) >= 1 ||
      (d.image.crop?.top ?? 0) + (d.image.crop?.bottom ?? 0) >= 1
    )
      throw new Error("Keep some of the picture visible when cropping.");
  }
  if (d.model3d) modelData(d.model3d.dataUrl);
  if (d.attachment) binaryData(d.attachment.dataUrl);
  if (d.video) videoUrl(d.video.url);
  if (d.spreadsheet) {
    if (
      !d.spreadsheet.cells.length ||
      d.spreadsheet.cells.length > 200 ||
      d.spreadsheet.cells.some((row) => row.length > 50)
    )
      throw new Error(
        "Use up to 200 rows and 50 columns in an embedded worksheet.",
      );
    d.spreadsheet.cells.flat().forEach(validateText);
  }
  if (d.equation) parseEquation(d.equation.expression);
  if (d.chart) {
    if (
      !["bar", "line", "pie"].includes(d.chart.kind) ||
      !d.chart.categories.length ||
      d.chart.categories.length > 200 ||
      !d.chart.series.length ||
      d.chart.series.length > 20
    )
      throw new Error("Use 1–200 chart categories and 1–20 series.");
    validateText(d.chart.title);
    d.chart.categories.forEach(validateText);
    for (const s of d.chart.series) {
      validateText(s.name);
      if (
        s.values.length !== d.chart.categories.length ||
        s.values.some((n) => !Number.isFinite(n))
      )
        throw new Error("Each chart series needs one number per category.");
    }
  }
  if (d.shape) {
    for (const color of [d.shape.fill, d.shape.stroke])
      if (color && !/^#[\da-f]{6}$/i.test(color))
        throw new Error("Choose an RGB shape color.");
    if (d.shape.rotation !== undefined) finite(d.shape.rotation, -36000, 36000);
  }
  if (d.position) {
    for (const axis of [d.position.horizontal, d.position.vertical])
      if (axis.offsetPt !== undefined) finite(axis.offsetPt, -15840, 15840);
  }
  if (d.children) {
    if (d.children.length > 100) throw new Error("Group at most 100 objects.");
    d.children.forEach((c) => validateDrawing(c, depth + 1));
  }
  return structuredClone(d);
}
