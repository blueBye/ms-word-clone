import type { Element } from "@xmldom/xmldom";
import { childElements, descendants } from "./parser/xml.js";
export function equationText(root: Element): string {
  const child = (name: string) =>
      childElements(root).find((e) => e.localName === name),
    read = (e?: Element) => (e ? equationText(e) : "");
  if (root.localName === "t") return root.textContent ?? "";
  if (root.localName?.endsWith("Pr")) return "";
  if (root.localName === "f")
    return `\\frac{${read(child("num"))}}{${read(child("den"))}}`;
  if (root.localName === "sSup" || root.localName === "sSub")
    return `{${read(child("e"))}}${root.localName === "sSup" ? "^" : "_"}{${read(child(root.localName === "sSup" ? "sup" : "sub"))}}`;
  if (root.localName === "rad") return `\\sqrt{${read(child("e"))}}`;
  return childElements(root)
    .map((e) => equationText(e))
    .join("");
}
