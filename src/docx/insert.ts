import { XMLSerializer, type Element } from "@xmldom/xmldom";
import type { DocxPackage } from "./package.js";
import { childElements, parseXml } from "./parser/xml.js";
const serializer = new XMLSerializer();
export async function addRelation(
  changes: Map<string, string>,
  pkg: DocxPackage,
  source: string,
  target: string,
  kind:
    | "image"
    | "hyperlink"
    | "chart"
    | "package"
    | "model3d"
    | "coreProperties"
    | "oleObject",
  w: string,
  external = false,
) {
  const folder = source.split("/").slice(0, -1).join("/"),
    basename = source.split("/").at(-1)!;
  const path = `${folder ? folder + "/" : ""}_rels/${basename}.rels`,
    ns = "http://schemas.openxmlformats.org/package/2006/relationships";
  const root = parseXml(
    pkg.hasPart(path)
      ? await pkg.readText(path)
      : `<Relationships xmlns="${ns}"/>`,
  ).documentElement!;
  const ids = new Set(childElements(root).map((e) => e.getAttribute("Id")));
  let n = 1;
  while (ids.has(`rId${n}`)) n++;
  const relation = root.ownerDocument!.createElementNS(ns, "Relationship");
  relation.setAttribute("Id", `rId${n}`);
  relation.setAttribute(
    "Type",
    kind === "coreProperties"
      ? "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
      : kind === "model3d"
        ? "http://schemas.microsoft.com/office/2017/06/relationships/model3d"
        : `${w.includes("purl.oclc.org") ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}/${kind}`,
  );
  relation.setAttribute("Target", target);
  if (external) relation.setAttribute("TargetMode", "External");
  root.appendChild(relation);
  changes.set(path, serializer.serializeToString(root.ownerDocument!));
  return `rId${n}`;
}
export async function imageContentType(
  changes: Map<string, string>,
  pkg: DocxPackage,
  path: string,
  mime: string,
) {
  const root = parseXml(
    await pkg.readText("[Content_Types].xml"),
  ).documentElement!;
  const node = root.ownerDocument!.createElementNS(
    root.namespaceURI,
    "Override",
  );
  node.setAttribute("PartName", `/${path}`);
  node.setAttribute("ContentType", mime);
  root.appendChild(node);
  changes.set(
    "[Content_Types].xml",
    serializer.serializeToString(root.ownerDocument!),
  );
}
export function imageRun(
  parent: Element,
  id: string,
  width: number,
  height: number,
  alt: string,
  drawingId: number,
) {
  const w = parent.namespaceURI!,
    strict = w.includes("purl.oclc.org"),
    r = strict
      ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
      : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const cx = Math.round(width * 12700),
    cy = Math.round(height * 12700);
  const root = parseXml(
    `<w:r xmlns:w="${w}" xmlns:r="${r}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${drawingId}" name="Picture ${drawingId}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${drawingId}" name="Picture ${drawingId}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`,
  ).documentElement!;
  const prop = root.getElementsByTagName("wp:docPr")[0];
  prop.setAttribute("descr", alt.slice(0, 1000));
  return parent.ownerDocument!.importNode(root, true);
}
