import type { ParagraphProperties, RunProperties } from "../model/document.js";
import type { DocumentStyle, EditCommand } from "./types.js";
import { mergeParagraph, mergeRun } from "./format.js";
import { builtins } from "./style-presets.js";
import { runProperties, paragraphProperties } from "./text.js";

export interface StyleDefinition extends DocumentStyle {
  kind?: "paragraph" | "character";
  /** Imported style booleans can compose as toggles; direct formatting always uses absolute values. */
  toggleRun?: boolean;
}
export interface StyleCatalog {
  definitions: StyleDefinition[];
  defaultParagraphStyleId?: string;
  defaults?: { paragraph: ParagraphProperties; run: RunProperties };
}

export class StyleStore {
  readonly definitions: Map<string, StyleDefinition>;
  readonly defaultId: string;
  readonly defaults: { paragraph: ParagraphProperties; run: RunProperties };
  private cache = new Map<string, StyleDefinition[]>();
  constructor(catalog: StyleCatalog) {
    this.definitions = new Map(
      catalog.definitions.map((s) => [s.id, structuredClone(s)]),
    );
    this.defaultId = catalog.defaultParagraphStyleId ?? "Normal";
    this.defaults = structuredClone(
      catalog.defaults ?? { paragraph: {}, run: {} },
    );
    for (const id of this.definitions.keys()) this.chain(id);
  }
  chain(id: string): StyleDefinition[] {
    const cached = this.cache.get(id);
    if (cached) return cached;
    const chain: StyleDefinition[] = [],
      seen = new Set<string>();
    let next: string | undefined = id;
    while (next) {
      if (seen.has(next))
        throw new Error("Style inheritance would create a cycle.");
      seen.add(next);
      const style = this.definitions.get(next);
      if (!style) break;
      chain.unshift(style);
      next = style.basedOn;
    }
    this.cache.set(id, chain);
    return chain;
  }
  resolve(
    id: string,
    paragraph = this.defaults.paragraph,
    run = this.defaults.run,
  ) {
    for (const layer of this.chain(id)) {
      paragraph = mergeParagraph(paragraph, layer.paragraph);
      run = mergeRun(run, layer.run, layer.toggleRun);
    }
    return { paragraph, run };
  }
  ensure(id: string): string[] {
    if (this.definitions.has(id)) return [];
    const preset = builtins[id];
    if (!preset)
      throw new Error("The selected style does not exist in this document.");
    const added = id === "Normal" ? [] : this.ensure("Normal");
    this.definitions.set(id, {
      id,
      name: preset.name,
      gallery: true,
      kind: "paragraph",
      basedOn: id === "Normal" ? undefined : "Normal",
      run: {
        ...(preset.size
          ? {
              fontSizePt: preset.size,
              complexScript: { fontSizePt: preset.size },
            }
          : {}),
        ...(preset.bold
          ? {
              bold: true,
              complexScript: { fontSizePt: preset.size, bold: true },
            }
          : {}),
      },
      paragraph: {
        spaceBeforePt: preset.before ?? 0,
        spaceAfterPt: preset.after ?? 0,
        ...(preset.outline !== undefined ? { keepNext: true } : {}),
      },
    });
    this.cache.clear();
    return [...added, id];
  }
  edit(command: Extract<EditCommand, { type: "style" }>) {
    const added = this.ensure(command.styleId);
    const original = this.definitions.get(command.styleId)!;
    const next = structuredClone(original);
    if (command.name !== undefined) {
      if (!command.name.trim() || command.name.length > 200)
        throw new Error("Enter a style name between 1 and 200 characters.");
      next.name = command.name.trim();
    }
    if (command.basedOn !== undefined) {
      if (command.basedOn && !this.definitions.has(command.basedOn))
        throw new Error("The base style does not exist.");
      next.basedOn = command.basedOn || undefined;
    }
    if (command.gallery !== undefined) next.gallery = command.gallery;
    if (command.run) {
      const patch = runProperties(command.run);
      if (next.toggleRun) {
        const base = next.basedOn
          ? this.resolve(next.basedOn).run
          : this.defaults.run;
        for (const key of [
          "bold",
          "italic",
          "strike",
          "caps",
          "smallCaps",
          "hidden",
        ] as const)
          if (patch[key] !== undefined) patch[key] = patch[key] !== !!base[key];
        for (const key of ["bold", "italic"] as const)
          if (patch.complexScript?.[key] !== undefined)
            patch.complexScript[key] =
              patch.complexScript[key] !== !!base.complexScript?.[key];
      }
      next.run = mergeRun(next.run, patch);
    }
    if (command.paragraph)
      next.paragraph = mergeParagraph(
        next.paragraph,
        paragraphProperties(command.paragraph),
      );
    this.definitions.set(next.id, next);
    this.cache.clear();
    try {
      this.chain(next.id);
    } catch (error) {
      this.definitions.set(original.id, original);
      this.cache.clear();
      throw error;
    }
    return [
      ...new Set([
        ...added,
        ...(command.run || command.paragraph || command.basedOn !== undefined
          ? [next.id]
          : []),
      ]),
    ];
  }
  list(): DocumentStyle[] {
    const present = [...this.definitions.values()]
      .filter((s) => s.kind !== "character")
      .map((s) => {
        const resolved = this.resolve(s.id);
        return {
          id: s.id,
          name: s.name,
          basedOn: s.basedOn,
          gallery: s.gallery,
          run: resolved.run,
          paragraph: resolved.paragraph,
        };
      });
    for (const [id, preset] of Object.entries(builtins)) {
      if (this.definitions.has(id)) continue;
      const normal = this.resolve(this.defaultId);
      present.push({
        id,
        name: preset.name,
        basedOn: id === "Normal" ? undefined : "Normal",
        gallery: true,
        run: mergeRun(normal.run, {
          ...(preset.size
            ? {
                fontSizePt: preset.size,
                complexScript: { fontSizePt: preset.size },
              }
            : {}),
          ...(preset.bold ? { bold: true } : {}),
        }),
        paragraph: {
          spaceBeforePt: preset.before ?? 0,
          spaceAfterPt: preset.after ?? 0,
        },
      });
    }
    return present;
  }
  snapshot(): StyleCatalog {
    return structuredClone({
      definitions: [...this.definitions.values()],
      defaults: this.defaults,
      defaultParagraphStyleId: this.defaultId,
    });
  }
}
