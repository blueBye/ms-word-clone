import type { Element } from "@xmldom/xmldom";
import type { DrawingInline } from "../model/document.js";
import { childElements } from "./parser/xml.js";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
export function writeImageEffects(
  blip: Element,
  image: NonNullable<DrawingInline["image"]>,
) {
  const set = (
    name: string,
    enabled: boolean,
    attrs: Record<string, number> = {},
  ) => {
    for (const child of childElements(blip).filter((c) => c.localName === name))
      blip.removeChild(child);
    if (enabled) {
      const node = blip.ownerDocument!.createElementNS(A, `a:${name}`);
      for (const [key, value] of Object.entries(attrs))
        node.setAttribute(key, String(value));
      blip.insertBefore(
        node,
        childElements(blip).find((c) => c.localName === "extLst") ?? null,
      );
    }
  };
  if (image.opacity !== undefined)
    set("alphaModFix", image.opacity !== 1, {
      amt: Math.round(image.opacity * 100000),
    });
  if (image.brightness !== undefined || image.contrast !== undefined)
    set("lum", !!image.brightness || !!image.contrast, {
      bright: Math.round((image.brightness ?? 0) * 100000),
      contrast: Math.round((image.contrast ?? 0) * 100000),
    });
  if (image.grayscale !== undefined) set("grayscl", image.grayscale);
  if (image.blurPt !== undefined)
    set("blur", image.blurPt > 0, { rad: Math.round(image.blurPt * 12700) });
}
export function readImageEffects(
  blip?: Element,
): Partial<NonNullable<DrawingInline["image"]>> {
  if (!blip) return {};
  const children = childElements(blip),
    lum = children.find((c) => c.localName === "lum"),
    blur = children.find((c) => c.localName === "blur");
  return {
    ...(lum
      ? {
          brightness: Number(lum.getAttribute("bright") || 0) / 100000,
          contrast: Number(lum.getAttribute("contrast") || 0) / 100000,
        }
      : {}),
    ...(children.some((c) => c.localName === "grayscl")
      ? { grayscale: true }
      : {}),
    ...(blur ? { blurPt: Number(blur.getAttribute("rad") || 0) / 12700 } : {}),
  };
}
