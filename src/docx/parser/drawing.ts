import type { Element } from "@xmldom/xmldom";
import type { DrawingInline, DrawingPosition } from "../../model/document.js";
import { relationshipAttribute } from "../relationships.js";
import { childElements, descendants, parseXml } from "./xml.js";
import type { BlockParser, ParseContext } from "./context.js";

function numberAttribute(
  node: Element | undefined,
  name: string,
  divisor = 1,
): number | undefined {
  const raw = node?.getAttribute(name);
  return raw && Number.isFinite(Number(raw))
    ? Number(raw) / divisor
    : undefined;
}
function drawingColor(
  node: Element | undefined,
  context?: ParseContext,
): string | undefined {
  if (!node) return undefined;
  const color = childElements(node)[0];
  const value =
    color?.localName === "schemeClr"
      ? context?.theme.colors[color.getAttribute("val") ?? ""]
      : color?.getAttribute("val");
  return value && /^[\da-f]{6}$/i.test(value) ? `#${value}` : undefined;
}
/** Read image-wide alpha effects only; clrChange/alpha affects a color, not the whole image. */
function imageOpacity(blip: Element | undefined): number | undefined {
  if (!blip) return undefined;
  let opacity = 1;
  let specified = false;
  for (const effect of childElements(blip)) {
    if (!["alphaModFix", "alphaOff", "alpha"].includes(effect.localName ?? ""))
      continue;
    const value = numberAttribute(
      effect,
      effect.localName === "alpha" ? "val" : "amt",
      100000,
    );
    if (value === undefined) continue;
    opacity =
      effect.localName === "alphaOff"
        ? opacity + value
        : effect.localName === "alpha"
          ? value
          : opacity * value;
    opacity = Math.max(0, Math.min(1, opacity));
    specified = true;
  }
  return specified ? opacity : undefined;
}
export function parseDrawing(
  root: Element,
  context?: ParseContext,
  parseBlocks?: BlockParser,
): DrawingInline[] {
  const anchors = descendants(root, "anchor").concat(
    descendants(root, "inline"),
  );
  if (!anchors.length) {
    context?.warn(
      "unsupported-vml",
      "A legacy VML drawing cannot be positioned accurately.",
    );
    return [];
  }
  return anchors.map((anchor) => {
    const nodes = descendants(anchor, "extent");
    const extent = nodes[0];
    const image = descendants(anchor, "blip")[0];
    const chartNode = descendants(anchor, "chart")[0],
      chartId = chartNode && relationshipAttribute(chartNode, "id"),
      chartRelationship = chartId && context?.relationships.get(chartId),
      chart =
        chartRelationship && context?.charts?.get(chartRelationship.target);
    const modelNode = descendants(anchor, "model3d")[0],
      modelId = modelNode && relationshipAttribute(modelNode, "embed"),
      modelRel = modelId && context?.relationships.get(modelId),
      modelUrl =
        modelRel && !modelRel.external && context?.assets.get(modelRel.target);
    const webVideo = descendants(anchor, "webVideoPr")[0],
      videoSrc =
        webVideo &&
        /\bsrc=["']([^"']+)["']/.exec(
          webVideo.getAttribute("embeddedHtml") ?? "",
        )?.[1];
    const id = image && relationshipAttribute(image, "embed");
    const rel = id && context?.relationships.get(id);
    const dataUrl =
      rel && !rel.external && rel.type === "image"
        ? context?.assets.get(rel.target)
        : undefined;
    const shapeRoot = descendants(anchor, "wsp")[0];
    const shapeProperties =
      shapeRoot && childElements(shapeRoot).find((n) => n.localName === "spPr");
    const pictureProperties =
      descendants(anchor, "pic")[0] &&
      descendants(descendants(anchor, "pic")[0], "spPr")[0];
    const geometry = descendants(
      shapeProperties ?? pictureProperties ?? anchor,
      "prstGeom",
    )[0];
    const propertyChildren = shapeProperties
      ? childElements(shapeProperties)
      : [];
    const shapeStyle =
      shapeRoot &&
      childElements(shapeRoot).find((n) => n.localName === "style");
    const fillReference =
      shapeStyle &&
      childElements(shapeStyle).find((n) => n.localName === "fillRef");
    const fill = propertyChildren.some((n) => n.localName === "noFill")
      ? undefined
      : drawingColor(
          propertyChildren.find((n) => n.localName === "solidFill") ??
            fillReference,
          context,
        );
    const line = propertyChildren.find((n) => n.localName === "ln");
    const crop = descendants(anchor, "srcRect")[0];
    const text = descendants(anchor, "txbxContent").find((n) =>
      n.namespaceURI?.includes("wordprocessingml"),
    );
    const docPr = descendants(anchor, "docPr")[0];
    const group = childElements(
      descendants(anchor, "graphicData")[0] ?? anchor,
    ).find((e) => e.localName === "wgp");
    const groupChildren = group
      ? childElements(group)
          .filter((e) => ["wsp", "pic", "wgp"].includes(e.localName ?? ""))
          .flatMap((child) => {
            const transform = descendants(child, "xfrm")[0],
              ext =
                transform &&
                childElements(transform).find((e) => e.localName === "ext"),
              off =
                transform &&
                childElements(transform).find((e) => e.localName === "off");
            const wrapper = parseXml(
              '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><wp:inline><wp:extent/><wp:docPr id="1" name="Group object"/><a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing>',
            ).documentElement!;
            const size = descendants(wrapper, "extent")[0];
            size.setAttribute("cx", ext?.getAttribute("cx") ?? "12700");
            size.setAttribute("cy", ext?.getAttribute("cy") ?? "12700");
            descendants(wrapper, "graphicData")[0].appendChild(
              wrapper.ownerDocument!.importNode(child, true),
            );
            const originals = descendants(child, "txbxContent"),
              copies = descendants(wrapper, "txbxContent");
            const nested = parseDrawing(
              wrapper,
              context,
              parseBlocks
                ? (node, ctx) =>
                    parseBlocks(originals[copies.indexOf(node)] ?? node, ctx)
                : undefined,
            );
            return nested.map((d) => ({
              ...d,
              position: {
                horizontal: {
                  relativeTo: "paragraph",
                  offsetPt: numberAttribute(off, "x", 12700) ?? 0,
                },
                vertical: {
                  relativeTo: "paragraph",
                  offsetPt: numberAttribute(off, "y", 12700) ?? 0,
                },
                behindText: false,
                zIndex: 2,
                wrap: "none",
              },
            }));
          })
      : undefined;
    const position = (local: string): DrawingPosition => {
      const node = childElements(anchor).find((n) => n.localName === local);
      const offset =
        node &&
        childElements(node).find((n) => n.localName === "posOffset")
          ?.textContent;
      return {
        relativeTo: node?.getAttribute("relativeFrom") ?? "paragraph",
        offsetPt:
          offset && Number.isFinite(Number(offset))
            ? Number(offset) / 12700
            : undefined,
        alignment:
          (node &&
            childElements(node).find((n) => n.localName === "align")
              ?.textContent) ||
          undefined,
      };
    };
    if (image && !dataUrl)
      context?.warn(
        "unsupported-image",
        "An external, missing or unsupported image was not loaded.",
      );
    const kind = geometry?.getAttribute("prst") ?? "rect";
    if (
      shapeRoot &&
      ![
        "rect",
        "ellipse",
        "roundRect",
        "triangle",
        "rtTriangle",
        "diamond",
        "hexagon",
        "pentagon",
        "chevron",
        "rightArrow",
        "leftArrow",
        "upArrow",
        "downArrow",
        "star5",
        "plus",
        "heart",
        "line",
      ].includes(kind)
    )
      context?.warn(
        "shape-approximation",
        "A drawing geometry is approximated by its bounding box.",
      );
    return {
      id: docPr?.getAttribute("id") ?? undefined,
      ...(groupChildren ? { children: groupChildren } : {}),
      flipH: descendants(anchor, "xfrm")[0]?.getAttribute("flipH") === "1",
      flipV: descendants(anchor, "xfrm")[0]?.getAttribute("flipV") === "1",
      type: "drawing" as const,
      chart: chart || undefined,
      ...(modelUrl ? { model3d: { dataUrl: modelUrl } } : {}),
      ...(videoSrc ? { video: { url: videoSrc } } : {}),
      widthPt: numberAttribute(extent, "cx", 12700) ?? 24,
      heightPt: numberAttribute(extent, "cy", 12700) ?? 24,
      alt: docPr?.getAttribute("descr") || docPr?.getAttribute("name") || "",
      image:
        dataUrl && !group
          ? {
              dataUrl,
              opacity: imageOpacity(image),
              crop: crop
                ? {
                    left: numberAttribute(crop, "l", 100000),
                    top: numberAttribute(crop, "t", 100000),
                    right: numberAttribute(crop, "r", 100000),
                    bottom: numberAttribute(crop, "b", 100000),
                  }
                : undefined,
            }
          : undefined,
      shape:
        !group && (shapeRoot || geometry)
          ? {
              kind,
              fill,
              stroke:
                line &&
                drawingColor(
                  childElements(line).find((n) => n.localName === "solidFill"),
                  context,
                ),
              strokeWidthPt: numberAttribute(line, "w", 12700),
              rotation: numberAttribute(
                descendants(anchor, "xfrm")[0],
                "rot",
                60000,
              ),
            }
          : group
            ? {
                kind: "group",
                rotation: numberAttribute(
                  descendants(anchor, "xfrm")[0],
                  "rot",
                  60000,
                ),
              }
            : undefined,
      blocks:
        !group && text && parseBlocks ? parseBlocks(text, context) : undefined,
      position:
        anchor.localName === "anchor"
          ? {
              horizontal: position("positionH"),
              vertical: position("positionV"),
              behindText: anchor.getAttribute("behindDoc") === "1",
              zIndex: Number(anchor.getAttribute("relativeHeight") ?? 0),
              wrap:
                childElements(anchor).find((n) =>
                  n.localName?.startsWith("wrap"),
                )?.localName ?? "wrapNone",
            }
          : undefined,
    };
  });
}
