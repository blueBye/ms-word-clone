import type {
  DocumentIR,
  Paragraph,
  PageLayout,
  StoryVariants,
} from "../model/document.js";
import { getParagraphs } from "../model/traversal.js";
export function sectionTarget(
  document: DocumentIR,
  id?: string,
): Paragraph | undefined {
  if (!id) return;
  const index = document.blocks.findIndex((b) =>
    getParagraphs([b]).some((p) => p.id === id),
  );
  if (index >= 0)
    return document.blocks
      .slice(index)
      .find((b): b is Paragraph => b.type === "paragraph" && !!b.section);
  return document.blocks.find(
    (b): b is Paragraph =>
      b.type === "paragraph" &&
      !!b.section &&
      [
        ...Object.values(b.section.headers ?? {}),
        ...Object.values(b.section.footers ?? {}),
      ].some((blocks) => getParagraphs(blocks ?? []).some((p) => p.id === id)),
  );
}
export function pageForParagraph(
  document: DocumentIR,
  id?: string,
): PageLayout {
  return sectionTarget(document, id)?.section ?? document.page ?? {};
}
export function splitSections(document: DocumentIR): DocumentIR[] {
  const sections: DocumentIR[] = [],
    pending: DocumentIR["blocks"] = [];
  let headers: StoryVariants | undefined, footers: StoryVariants | undefined;
  for (const block of document.blocks) {
    pending.push(block);
    if (block.type === "paragraph" && block.section) {
      const page = block.section;
      headers = block.section.headers ?? headers;
      footers = block.section.footers ?? footers;
      sections.push({
        ...document,
        blocks: pending.splice(0),
        page,
        headers,
        footers,
      });
    }
  }
  if (pending.length || !sections.length)
    sections.push({
      ...document,
      blocks: pending,
      page: document.page,
      headers: document.headers ?? headers,
      footers: document.footers ?? footers,
    });
  return sections;
}
