import type { Block, DocumentIR, Paragraph } from "../model/document.js";
type Key = string | number;
type Patch = { value?: Paragraph; children: Map<Key, Patch> };

/** Stable paragraph locations let a text edit copy its ancestor arrays without walking every run. */
export function paragraphLocations(document: DocumentIR) {
  const locations = new Map<string, Key[][]>();
  const visit = (blocks: Block[], path: Key[]) =>
    blocks.forEach((block, index) => {
      const here = [...path, index];
      if (block.type === "paragraph") {
        if (block.id)
          locations.set(block.id, [...(locations.get(block.id) ?? []), here]);
        block.children.forEach((inline, i) => {
          if (inline.type === "drawing" && inline.blocks)
            visit(inline.blocks, [...here, "children", i, "blocks"]);
        });
      } else
        block.rows.forEach((row, r) =>
          row.cells.forEach((cell, c) =>
            visit(cell.blocks, [...here, "rows", r, "cells", c, "blocks"]),
          ),
        );
    });
  visit(document.blocks, ["blocks"]);
  for (const story of ["headers", "footers"] as const)
    for (const variant of ["default", "first", "even"] as const)
      if (document[story]?.[variant])
        visit(document[story][variant], [story, variant]);
  return locations;
}
export function patchParagraphs(
  document: DocumentIR,
  locations: Map<string, Key[][]>,
  changes: Map<string, Paragraph>,
  onParagraph: (p: Paragraph) => void,
): DocumentIR {
  if (!changes.size) return document;
  const root: Patch = { children: new Map() };
  for (const [id, value] of changes)
    for (const path of locations.get(id) ?? []) {
      let node = root;
      for (const key of path) {
        let child = node.children.get(key);
        if (!child) {
          child = { children: new Map() };
          node.children.set(key, child);
        }
        node = child;
      }
      node.value = value;
    }
  const apply = (original: unknown, patch: Patch): unknown => {
    const value = patch.value ?? original;
    if (!patch.children.size) return value;
    const result = Array.isArray(value)
      ? value.slice()
      : { ...(value as object) };
    for (const [key, child] of patch.children) {
      const record = result as Record<Key, unknown>;
      record[key] = apply(record[key], child);
    }
    if (!Array.isArray(result) && (result as Paragraph).type === "paragraph")
      onParagraph(result as Paragraph);
    return result;
  };
  return apply(document, root) as DocumentIR;
}
