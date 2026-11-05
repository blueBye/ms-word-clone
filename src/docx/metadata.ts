import { XMLSerializer } from "@xmldom/xmldom";
import type { DocxPackage } from "./package.js";
import { readRelationships } from "./relationships.js";
import { addRelation, imageContentType } from "./insert.js";
import { parseXml, childElements } from "./parser/xml.js";
export async function writeProperty(
  changes: Map<string, string>,
  pkg: DocxPackage,
  w: string,
  name: "title" | "author",
  value: string,
) {
  const rels = await readRelationships(pkg, ""),
    found = [...rels.values()].find(
      (r) => r.type === "core-properties" && !r.external,
    ),
    path = found?.target ?? "docProps/core.xml",
    CP =
      "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    DC = "http://purl.org/dc/elements/1.1/";
  if (!found) {
    await addRelation(changes, pkg, "", path, "coreProperties", w);
    await imageContentType(
      changes,
      pkg,
      path,
      "application/vnd.openxmlformats-package.core-properties+xml",
    );
  }
  const root = parseXml(
      pkg.hasPart(path)
        ? await pkg.readText(path)
        : `<cp:coreProperties xmlns:cp="${CP}" xmlns:dc="${DC}"/>`,
    ).documentElement!,
    local = name === "title" ? "title" : "creator";
  let node = childElements(root).find(
    (n) => n.localName === local && n.namespaceURI === DC,
  );
  if (!node) {
    node = root.ownerDocument!.createElementNS(DC, `dc:${local}`);
    root.appendChild(node);
  }
  node.textContent = value;
  changes.set(path, new XMLSerializer().serializeToString(root.ownerDocument!));
}
export async function readProperties(pkg: DocxPackage) {
  const rels = await readRelationships(pkg, ""),
    part = [...rels.values()].find(
      (r) => r.type === "core-properties" && !r.external,
    );
  if (!part || !pkg.hasPart(part.target)) return undefined;
  const root = parseXml(await pkg.readText(part.target)).documentElement!;
  return Object.fromEntries(
    childElements(root)
      .filter((n) => ["title", "creator"].includes(n.localName ?? ""))
      .map((n) => [
        n.localName === "creator" ? "author" : "title",
        n.textContent ?? "",
      ]),
  );
}
