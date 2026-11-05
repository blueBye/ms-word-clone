import type { Element } from "@xmldom/xmldom";
import {
  childElements,
  isWord,
  childWord,
  descendants,
  wordAttribute,
} from "./parser/xml.js";
const text = (run: Element) =>
  childElements(run)
    .map((e) =>
      isWord(e, "t")
        ? (e.textContent ?? "")
        : isWord(e, "tab")
          ? "\t"
          : isWord(e, "br") || isWord(e, "cr")
            ? "\n"
            : "",
    )
    .join("");
export const annotationNode = (node: Element) =>
  ["bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd"].some(
    (name) => isWord(node, name),
  ) ||
  (isWord(node, "r") &&
    (!!childWord(node, "commentReference") ||
      wordAttribute(childWord(node, "br") ?? node, "type") === "column"));
export function clearAnchors(p: Element) {
  for (const node of childElements(p))
    if (annotationNode(node)) p.removeChild(node);
    else if (isWord(node, "hyperlink")) clearAnchors(node);
}
export function readAnchors(p: Element) {
  const anchors: { node: Element; position: number }[] = [];
  let position = 0;
  const visit = (parent: Element) => {
    for (const node of childElements(parent)) {
      if (annotationNode(node))
        anchors.push({ node: node.cloneNode(true) as Element, position });
      else if (isWord(node, "r")) position += text(node).length;
      else if (isWord(node, "hyperlink")) visit(node);
    }
  };
  visit(p);
  return anchors;
}
export function insertAnchor(p: Element, node: Element, position: number) {
  const visit = (parent: Element): Element[] =>
    childElements(parent).flatMap((e) =>
      isWord(e, "r") && !annotationNode(e)
        ? [e]
        : isWord(e, "hyperlink")
          ? visit(e)
          : [],
    );
  let offset = 0;
  for (const run of visit(p)) {
    const value = text(run),
      next = offset + value.length;
    if (position <= next) {
      if (position <= offset) run.parentNode!.insertBefore(node, run);
      else if (position === next)
        run.parentNode!.insertBefore(node, run.nextSibling);
      else {
        const right = run.cloneNode(true) as Element,
          left = run.cloneNode(true) as Element;
        for (const piece of [left, right])
          for (const c of childElements(piece))
            if (!isWord(c, "rPr")) piece.removeChild(c);
        const add = (parent: Element, value: string) => {
          const t = parent.ownerDocument!.createElementNS(
            p.namespaceURI,
            "w:t",
          );
          t.setAttributeNS(
            "http://www.w3.org/XML/1998/namespace",
            "xml:space",
            "preserve",
          );
          t.textContent = value;
          parent.appendChild(t);
        };
        add(left, value.slice(0, position - offset));
        add(right, value.slice(position - offset));
        run.parentNode!.insertBefore(left, run);
        run.parentNode!.insertBefore(node, run);
        run.parentNode!.replaceChild(right, run);
      }
      return;
    }
    offset = next;
  }
  p.appendChild(node);
}
export function mapAnchor(
  position: number,
  start: number,
  end: number,
  length: number,
  trailing = false,
) {
  return position < start || (position === start && !trailing)
    ? position
    : position >= end
      ? position + length - (end - start)
      : start + length;
}
export function removeAnnotation(
  p: Element,
  kind: "bookmark" | "comment",
  id: string,
) {
  for (const node of descendants(
    p,
    kind === "bookmark" ? "bookmarkStart" : "commentRangeStart",
  ).concat(
    descendants(p, kind === "bookmark" ? "bookmarkEnd" : "commentRangeEnd"),
    kind === "comment" ? descendants(p, "commentReference") : [],
  ))
    if (wordAttribute(node, "id") === id) {
      const target = isWord(node, "commentReference") ? node.parentNode! : node;
      target.parentNode?.removeChild(target);
    }
}
