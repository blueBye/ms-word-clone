import type {
  Block,
  DocumentIR,
  Paragraph,
  RunProperties,
} from "../model/document.js";
import { getParagraphs } from "../model/traversal.js";
import type {
  DocumentSession,
  EditCommand,
  ParagraphCapability,
} from "./types.js";
import { mergeParagraph, mergeRun } from "./format.js";
import { StyleStore, type StyleCatalog } from "./styles.js";
import { StyleDependencies } from "./dependencies.js";
import {
  finite,
  paragraphProperties,
  paragraphText,
  rewriteParagraph,
  validateRange,
  validateText,
} from "./text.js";
import { selectionStyleDelta } from "./style-selection.js";
import { ListStore, type ListDefinition } from "./lists.js";
import { editStructure } from "./structure.js";
import { searchMatches, type SearchOptions } from "./search.js";
import { paragraphLocations, patchParagraphs } from "./tree.js";

export interface EngineSnapshot {
  document: DocumentIR;
  styles: StyleCatalog;
  lists?: ListDefinition[];
}
export interface EngineStatistics {
  paragraphs: number;
  words: number;
  resolvedParagraphs: number;
  revision: number;
}
export interface EngineOptions {
  capability?: (id: string) => ParagraphCapability;
  lockedReason?: string;
  historyLimit?: number;
}
export function allParagraphs(document: DocumentIR) {
  const all = [
    ...getParagraphs(document.blocks),
    ...Object.values(document.headers ?? {}).flatMap((b) =>
      getParagraphs(b ?? []),
    ),
    ...Object.values(document.footers ?? {}).flatMap((b) =>
      getParagraphs(b ?? []),
    ),
  ];
  return [...new Map(all.filter((p) => p.id).map((p) => [p.id!, p])).values()];
}
function replaceBlocks(
  blocks: Block[],
  changes: Map<string, Paragraph>,
): Block[] {
  let changed = false;
  const result = blocks.map((block): Block => {
    let next = block;
    if (block.type === "paragraph") {
      next = changes.get(block.id!) ?? block;
      const children = next.children.map((inline) => {
        if (inline.type !== "drawing" || !inline.blocks) return inline;
        const nested = replaceBlocks(inline.blocks, changes);
        return nested === inline.blocks
          ? inline
          : { ...inline, blocks: nested };
      });
      if (
        children.some((child, i) => child !== (next as Paragraph).children[i])
      )
        next = { ...next, children };
    } else {
      const rows = block.rows.map((row) => {
        const cells = row.cells.map((cell) => {
          const nested = replaceBlocks(cell.blocks, changes);
          return nested === cell.blocks ? cell : { ...cell, blocks: nested };
        });
        return cells.some((cell, i) => cell !== row.cells[i])
          ? { ...row, cells }
          : row;
      });
      if (rows.some((row, i) => row !== block.rows[i]))
        next = { ...block, rows };
    }
    if (next !== block) changed = true;
    return next;
  });
  return changed ? result : blocks;
}
function freeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}
/** Format-independent document state. No DOM, React, ZIP, XML or browser selection is required. */
export class DocumentEngine implements DocumentSession {
  private value!: DocumentIR;
  private store!: StyleStore;
  private lists!: ListStore;
  private authored = new Map<string, Paragraph>();
  private resolved = new Map<string, Paragraph>();
  private dependencies!: StyleDependencies;
  private counts = new Map<string, { text: string; words: number }>();
  private wordCount = 0;
  private calculated = 0;
  private history: EngineSnapshot[] = [];
  private cursor = 0;
  private nextId = 1;
  private currentSelection:
    | { id: string; start: number; end: number }
    | undefined;
  private listeners = new Set<() => void>();
  private version = 0;
  private pending = false;
  private initialDocument!: DocumentIR;
  private initialCatalog = "";
  private locations = new Map<string, (string | number)[][]>();
  private searchCache:
    | {
        key: string;
        query: string;
        options: SearchOptions;
        results: Map<
          string,
          { id: string; content: string; start: number; end: number }[]
        >;
        flat:
          | { id: string; content: string; start: number; end: number }[]
          | null;
      }
    | undefined;
  constructor(
    snapshot: EngineSnapshot,
    private options: EngineOptions = {},
  ) {
    if (
      snapshot.document.type !== "document" ||
      snapshot.document.schemaVersion !== 2 ||
      !Array.isArray(snapshot.document.blocks)
    )
      throw new Error("Unsupported document schema.");
    this.restore(structuredClone(snapshot));
    this.initialDocument = this.value;
    this.initialCatalog = JSON.stringify([
      this.store.snapshot(),
      this.lists.snapshot(),
    ]);
    this.history = [this.snapshot()];
  }
  get document() {
    return this.value;
  }
  get dirty() {
    return (
      this.value !== this.initialDocument ||
      JSON.stringify([this.store.snapshot(), this.lists.snapshot()]) !==
        this.initialCatalog
    );
  }
  get canUndo() {
    return this.cursor > 0;
  }
  get canRedo() {
    return this.cursor < this.history.length - 1;
  }
  get canSetPage() {
    return !this.options.lockedReason;
  }
  get lockedReason() {
    return this.options.lockedReason;
  }
  get selection() {
    return this.currentSelection ? { ...this.currentSelection } : undefined;
  }
  get statistics(): EngineStatistics {
    return {
      paragraphs: this.authored.size,
      words: this.wordCount,
      resolvedParagraphs: this.calculated,
      revision: this.version,
    };
  }
  getParagraph(id: string) {
    return this.resolved.get(id);
  }
  find(query: string, options: SearchOptions = {}) {
    if (!query) {
      this.searchCache = undefined;
      return [];
    }
    const key = JSON.stringify([query, options]);
    if (key !== this.searchCache?.key) {
      const results = new Map<
        string,
        { id: string; content: string; start: number; end: number }[]
      >();
      for (const [id, { text: content }] of this.counts)
        results.set(
          id,
          searchMatches(content, query, options).map((hit) => ({
            id,
            content,
            ...hit,
          })),
        );
      this.searchCache = {
        key,
        query,
        options: structuredClone(options),
        results,
        flat: null,
      };
    }
    return (this.searchCache!.flat ??= freeze(
      [...this.searchCache!.results.values()].flat(),
    ));
  }
  snapshot(): EngineSnapshot {
    return {
      document: this.value,
      styles: this.store.snapshot(),
      lists: this.lists.snapshot(),
    };
  }
  styles() {
    return Promise.resolve(freeze(this.store.list()));
  }
  subscribe(listener: () => void) {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }
  capability(id: string): ParagraphCapability {
    if (this.lockedReason)
      return { text: false, format: false, reason: this.lockedReason };
    if (!this.authored.has(id))
      return { text: false, format: false, reason: "Select a paragraph." };
    if (this.options.capability) return this.options.capability(id);
    const plain = this.authored
      .get(id)!
      .children.every((c) => c.type === "text" && !c.field && !c.href);
    return {
      text: plain,
      format: plain,
      reason: plain ? undefined : "Structured content is read-only.",
    };
  }
  private dependencyIds(p: Paragraph) {
    const ids = new Set<string>();
    for (const id of [
      p.directProperties?.styleId ??
        p.properties.styleId ??
        this.store.defaultId,
      ...p.children.flatMap((r) =>
        r.type === "text" && r.directProperties?.styleId
          ? [r.directProperties.styleId]
          : [],
      ),
    ]) {
      ids.add(id);
      for (const style of this.store.chain(id)) ids.add(style.id);
    }
    return ids;
  }
  private resolve(p: Paragraph): Paragraph {
    const direct = p.directProperties ?? p.properties;
    const context = p.formattingContext ?? this.store.defaults;
    const id = direct.styleId ?? this.store.defaultId;
    const styled = this.store.resolve(id, context.paragraph, context.run);
    const properties = {
      ...mergeParagraph(styled.paragraph, direct),
      styleId: id,
      ...(p.properties.list ? { list: p.properties.list } : {}),
    };
    const markProperties = mergeRun(styled.run, p.directMarkProperties ?? {});
    if (properties.list)
      properties.list = {
        ...properties.list,
        properties: mergeRun(
          markProperties,
          p.listMarkProperties ?? properties.list.properties,
        ),
      };
    const children = p.children.map((r) => {
      if (r.type !== "text")
        return r.type === "drawing" && r.blocks
          ? { ...r, blocks: replaceBlocks(r.blocks, this.resolved) }
          : r;
      const authored = r.directProperties ?? r.properties;
      let run = styled.run;
      if (authored.styleId)
        run = this.store.resolve(authored.styleId, {}, run).run;
      return {
        ...r,
        directProperties: authored,
        properties: mergeRun(run, authored),
      };
    });
    return {
      ...p,
      formattingContext: context,
      directProperties: direct,
      properties,
      markProperties,
      children,
    };
  }
  private materialize(ids: Iterable<string>) {
    const replacements = new Map<string, Paragraph>();
    this.calculated = 0;
    for (const id of new Set(ids)) {
      const raw = this.authored.get(id);
      if (!raw) continue;
      const next = freeze(this.resolve(raw));
      this.resolved.set(id, next);
      replacements.set(id, next);
      this.calculated++;
      const text = paragraphText(next),
        previous = this.counts.get(id);
      if (text !== previous?.text) {
        const words = Array.from(
          new Intl.Segmenter(this.value.language, {
            granularity: "word",
          }).segment(text),
        ).filter((s) => s.isWordLike).length;
        this.wordCount += words - (previous?.words ?? 0);
        this.counts.set(id, { text, words });
        if (this.searchCache) {
          const cache = this.searchCache;
          cache.results.set(
            id,
            searchMatches(text, cache.query, cache.options).map((hit) => ({
              id,
              content: text,
              ...hit,
            })),
          );
          cache.flat = null;
        }
      }
    }
    this.value = freeze(
      patchParagraphs(this.value, this.locations, replacements, (p) =>
        this.resolved.set(p.id!, p),
      ),
    );
  }
  private restore(snapshot: EngineSnapshot, alreadyResolved = false) {
    this.value = snapshot.document;
    this.store = new StyleStore(snapshot.styles);
    this.lists = new ListStore(snapshot.lists);
    this.authored = new Map();
    this.resolved = new Map();
    this.dependencies = new StyleDependencies();
    this.counts = new Map();
    this.wordCount = 0;
    this.searchCache = undefined;
    const paragraphs = [
      ...getParagraphs(this.value.blocks),
      ...Object.values(this.value.headers ?? {}).flatMap((b) =>
        getParagraphs(b ?? []),
      ),
      ...Object.values(this.value.footers ?? {}).flatMap((b) =>
        getParagraphs(b ?? []),
      ),
    ];
    const used = new Set(paragraphs.map((p) => p.id).filter(Boolean));
    for (const p of paragraphs)
      if (!p.id) {
        while (used.has(`p${this.nextId}`)) this.nextId++;
        p.id = `p${this.nextId++}`;
        used.add(p.id);
      }
    this.locations = paragraphLocations(this.value);
    for (const p of allParagraphs(this.value)) {
      this.nextId = Math.max(
        this.nextId,
        Number(p.id?.replace(/^p/, "")) + 1 || 1,
      );
      this.authored.set(p.id!, p);
      this.dependencies.set(p.id!, this.dependencyIds(p));
    }
    this.dependencies.flush();
    this.materialize(this.authored.keys());
    if (alreadyResolved) {
      this.value = snapshot.document;
      this.resolved = new Map(allParagraphs(this.value).map((p) => [p.id!, p]));
    }
  }
  private setParagraph(p: Paragraph) {
    this.authored.set(p.id!, p);
    this.dependencies.set(p.id!, this.dependencyIds(p));
    const changed = this.dependencies.flush();
    changed.add(p.id!);
    this.materialize(changed);
  }
  private editStyle(command: Extract<EditCommand, { type: "style" }>) {
    const changedStyles = this.store.edit(command);
    this.dependencies.touch(changedStyles);
    const affected = this.dependencies.flush();
    for (const id of affected) {
      const p = this.authored.get(id);
      if (p) this.dependencies.set(id, this.dependencyIds(p));
    }
    for (const id of this.dependencies.flush()) affected.add(id);
    this.materialize(affected);
  }
  private apply(command: EditCommand) {
    if (this.lockedReason) throw new Error(this.lockedReason);
    const id = "id" in command ? command.id : undefined;
    const p = id ? this.resolved.get(id) : undefined;
    if (id && !p)
      throw new Error(
        "This selection belongs to an older document revision. Select it again.",
      );
    if (command.type === "style") {
      this.editStyle(command);
      return;
    }
    if (command.type === "page") {
      const page = {
        ...this.value.page,
        widthPt: finite(command.widthPt, 72, 31680),
        heightPt: finite(command.heightPt, 72, 31680),
        ...(command.margins
          ? { margins: { ...this.value.page?.margins, ...command.margins } }
          : {}),
      };
      if (command.margins)
        for (const margin of Object.values(command.margins))
          finite(margin, 0, 15840);
      this.value = freeze({ ...this.value, page });
      return;
    }
    if (
      command.type === "text" ||
      command.type === "run" ||
      command.type === "clearRun"
    ) {
      const allowed = this.capability(command.id);
      if (!allowed.text) throw new Error(allowed.reason);
      const next = rewriteParagraph(p!, command);
      this.setParagraph(next);
      this.currentSelection =
        command.type === "text"
          ? {
              id: command.id,
              start: command.start + command.text.length,
              end: command.start + command.text.length,
            }
          : {
              id: command.id,
              start: command.start ?? 0,
              end: command.end ?? paragraphText(p!).length,
            };
      return;
    }
    if (command.type === "updateStyleFromSelection") {
      const allowed = this.capability(command.id);
      if (!allowed.format) throw new Error(allowed.reason);
      validateRange(paragraphText(p!), command.start, command.end);
      const style = this.store.list().find((s) => s.id === command.styleId);
      if (!style)
        throw new Error("The selected style does not exist in this document.");
      const delta = selectionStyleDelta(p!, command.start, command.end, style);
      if (!Object.keys(delta.run).length) return;
      this.editStyle({
        type: "style",
        styleId: command.styleId,
        run: delta.run,
      });
      if (p!.properties.styleId === command.styleId)
        this.setParagraph(
          rewriteParagraph(
            this.resolved.get(command.id)!,
            { type: "clearRun", id: command.id, ...delta.range },
            delta.inherit,
          ),
        );
      this.currentSelection = {
        id: command.id,
        start: command.start,
        end: command.end,
      };
      return;
    }
    if (command.type === "paragraph") {
      const patch = paragraphProperties(command.format),
        direct = { ...(p!.directProperties ?? p!.properties) };
      let children = p!.children;
      if (command.format.styleId) {
        const styles = this.store.ensure(command.format.styleId);
        if (styles.length) {
          this.dependencies.touch(styles);
          this.materialize(this.dependencies.flush());
        }
        for (const key of [
          "alignment",
          "spaceBeforePt",
          "spaceAfterPt",
          "lineSpacing",
          "indentLeftPt",
          "indentRightPt",
          "indentStartPt",
          "indentEndPt",
          "firstLinePt",
          "keepNext",
          "keepLines",
          "pageBreakBefore",
          "contextualSpacing",
        ] as const)
          delete direct[key];
        if (this.capability(command.id).format) {
          children = children.map((r) =>
            r.type === "text"
              ? { ...r, directProperties: { ...r.directProperties } }
              : r,
          );
          const runs = children.filter((r) => r.type === "text");
          for (const key of [
            "fontFamily",
            "fontSlots",
            "fontSizePt",
            "bold",
            "italic",
            "color",
            "underline",
            "strike",
            "caps",
            "smallCaps",
            "complexScript",
          ] as const) {
            const value = runs[0]?.directProperties?.[key];
            if (
              value !== undefined &&
              runs.every(
                (r) =>
                  JSON.stringify(r.directProperties?.[key]) ===
                  JSON.stringify(value),
              )
            )
              for (const r of runs) delete r.directProperties![key];
          }
        }
      }
      this.setParagraph({
        ...p!,
        directProperties: mergeParagraph(direct, patch),
        children,
      });
      return;
    }
    this.structure(command);
  }
  private structure(command: EditCommand) {
    const result = editStructure(
      this.value,
      command,
      this.lists,
      () => `p${this.nextId++}`,
      (id) => this.capability(id),
    );
    let document = result.document;
    const markers = new Map<string, Paragraph>();
    for (const blocks of [
      document.blocks,
      ...Object.values(document.headers ?? {}),
      ...Object.values(document.footers ?? {}),
    ])
      for (const [id, p] of this.lists.markers(
        getParagraphs(blocks ?? []),
        document.language,
      ))
        markers.set(id, p);
    document = { ...document, blocks: replaceBlocks(document.blocks, markers) };
    for (const kind of ["headers", "footers"] as const)
      if (document[kind]) {
        document[kind] = { ...document[kind] };
        for (const key of ["default", "first", "even"] as const)
          if (document[kind][key])
            document[kind][key] = replaceBlocks(document[kind][key], markers);
      }
    this.restore({
      document,
      styles: this.store.snapshot(),
      lists: this.lists.snapshot(),
    });
    if (result.selection) this.currentSelection = result.selection;
  }
  private async commit(commands: EditCommand[]) {
    if (this.pending)
      throw new Error("Wait for the current document change to finish.");
    if (commands.length > 500)
      throw new Error("Apply changes to at most 500 paragraphs at a time.");
    this.pending = true;
    const before = this.snapshot(),
      selected = this.currentSelection;
    try {
      for (const command of structuredClone(commands)) this.apply(command);
      if (
        before.document === this.value &&
        JSON.stringify(before.styles) === JSON.stringify(this.store.snapshot())
      )
        return this.value;
      this.history = [
        ...this.history.slice(0, this.cursor + 1),
        this.snapshot(),
      ];
      this.cursor++;
      if (this.history.length > (this.options.historyLimit ?? 100) + 1) {
        this.history.shift();
        this.cursor--;
      }
      this.version++;
    } catch (error) {
      this.restore(before, true);
      this.currentSelection = selected;
      throw error;
    } finally {
      this.pending = false;
    }
    for (const listener of this.listeners) listener();
    return this.value;
  }
  execute(command: EditCommand) {
    return this.commit([command]);
  }
  executeBatch(commands: EditCommand[]) {
    return this.commit(commands);
  }
  async undo() {
    if (this.canUndo) {
      this.restore(this.history[--this.cursor], true);
      this.currentSelection = undefined;
      this.version++;
      for (const listener of this.listeners) listener();
    }
    return this.value;
  }
  async redo() {
    if (this.canRedo) {
      this.restore(this.history[++this.cursor], true);
      this.currentSelection = undefined;
      this.version++;
      for (const listener of this.listeners) listener();
    }
    return this.value;
  }
}
export function createDocumentEngine(
  snapshot?: EngineSnapshot,
  options?: EngineOptions,
) {
  return new DocumentEngine(
    snapshot ?? {
      document: {
        type: "document",
        schemaVersion: 2,
        blocks: [
          {
            type: "paragraph",
            id: "p1",
            properties: {},
            directProperties: {},
            children: [
              { type: "text", text: "", properties: {}, directProperties: {} },
            ],
          },
        ],
      },
      styles: {
        definitions: [
          {
            id: "Normal",
            name: "Normal",
            gallery: true,
            run: {
              fontFamily: "Calibri",
              fontSlots: { ascii: "Calibri", hAnsi: "Calibri" },
              fontSizePt: 12,
              complexScript: { fontFamily: "Tahoma", fontSizePt: 12 },
            },
            paragraph: { spaceAfterPt: 8 },
          },
        ],
      },
    },
    options,
  );
}
