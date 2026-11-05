export type * from "../model/document.js";
export { getParagraphs, getTextRuns } from "../model/traversal.js";
export type * from "./types.js";
export {
  createDocumentEngine,
  DocumentEngine,
  allParagraphs,
} from "./document.js";
export type {
  EngineSnapshot,
  EngineOptions,
  EngineStatistics,
} from "./document.js";
export type { StyleCatalog, StyleDefinition } from "./styles.js";
export { paragraphText } from "./text.js";
export { searchMatches, nextMatchIndex } from "./search.js";
export type { SearchOptions, SearchHit, SearchDirection } from "./search.js";
export type { ListDefinition, ListLevel } from "./lists.js";
export { changeCase } from "./case.js";
export { runAtPosition, similarFormatting } from "./selection.js";
export { pageForParagraph, splitSections, sectionTarget } from "./sections.js";
