import type {
  Paragraph,
  ParagraphProperties,
  RunProperties,
} from "../model/document.js";
import type { EditCommand } from "./types.js";
import { finite, paragraphText } from "./text.js";
import { formatCounter } from "./counters.js";
import { localizeDigits } from "../model/numerals.js";

export interface ListLevel {
  start: number;
  format: string;
  text: string;
  suffix: string;
  restart?: number;
  paragraph: ParagraphProperties;
  run: RunProperties;
}
export interface ListDefinition {
  id: string;
  levels: [number, ListLevel][];
  preset?: string;
}
export class ListStore {
  readonly definitions: Map<string, ListDefinition>;
  constructor(definitions: ListDefinition[] = []) {
    this.definitions = new Map(
      definitions.map((d) => [d.id, structuredClone(d)]),
    );
  }
  set(
    p: Paragraph,
    command: Extract<EditCommand, { type: "list" }>,
    siblings: Paragraph[],
  ) {
    const context = {
      paragraph: { ...p.formattingContext?.paragraph },
      run: p.formattingContext?.run ?? {},
    };
    const old = this.definitions
      .get(p.properties.numbering?.id ?? "")
      ?.levels.find(([n]) => n === (p.properties.numbering?.level ?? 0))?.[1];
    for (const key of Object.keys(
      old?.paragraph ?? {},
    ) as (keyof ParagraphProperties)[])
      if (
        JSON.stringify(context.paragraph[key]) ===
        JSON.stringify(old!.paragraph[key])
      )
        delete context.paragraph[key];
    if (command.kind === "none")
      return {
        ...p,
        formattingContext: context,
        listMarkProperties: undefined,
        directProperties: { ...p.directProperties, numbering: { id: "0" } },
        properties: {
          ...p.properties,
          numbering: { id: "0" },
          list: undefined,
        },
      };
    if (
      ![
        "bullet",
        "decimal",
        "decimalZero",
        "lowerLetter",
        "upperLetter",
        "lowerRoman",
        "upperRoman",
      ].includes(command.kind)
    )
      throw new Error("Unsupported list format.");
    const level = command.level ?? p.properties.numbering?.level ?? 0;
    if (!Number.isInteger(level) || level < 0 || level > 8)
      throw new Error("List level must be between 1 and 9.");
    const start = finite(command.start ?? 1, 1, 1_000_000);
    if (!Number.isInteger(start))
      throw new Error("List start must be a positive whole number.");
    const bullet = command.bullet ?? "•",
      pattern =
        command.pattern?.replace(/%[1-9]/g, `%${level + 1}`) ??
        `%${level + 1}.`;
    if (
      bullet.length > 8 ||
      /[\u0000-\u001f]/u.test(bullet) ||
      pattern.length > 100 ||
      /[\u0000-\u001f]/u.test(pattern)
    )
      throw new Error("Invalid list marker.");
    const own = p.properties.numbering?.id;
    const matches = (id?: string) => {
      const definition = id ? this.definitions.get(id) : undefined,
        current = definition?.levels.find(([n]) => n === level)?.[1];
      return command.multilevel
        ? definition?.preset === command.multilevel
        : current?.format === command.kind &&
            (((command.restart ||
              command.start !== undefined ||
              command.continue) &&
              command.pattern === undefined &&
              command.bullet === undefined) ||
              current.text === (command.kind === "bullet" ? bullet : pattern));
    };
    let id = own;
    const index = siblings.findIndex((q) => q.id === p.id);
    if (command.continue) {
      id = siblings
        .slice(0, index)
        .reverse()
        .map((q) => q.properties.numbering?.id)
        .find(matches);
      if (!id)
        throw new Error(
          "No previous list with this format was found in this container.",
        );
    } else if (command.preserveDefinition) {
      if (!id || !this.definitions.has(id))
        throw new Error("This paragraph does not have a list definition.");
    } else if (!matches(id) || command.restart || command.start !== undefined) {
      id =
        !command.restart && command.start === undefined
          ? [siblings[index - 1], siblings[index + 1]]
              .filter(Boolean)
              .map((q) => q.properties.numbering?.id)
              .find((candidate) => candidate !== own && matches(candidate))
          : undefined;
      if (!id) {
        id = String(
          Math.max(
            0,
            ...[...this.definitions.keys()].map((s) => Number(s) || 0),
          ) + 1,
        );
        const levels: [number, ListLevel][] = Array.from(
          { length: 9 },
          (_, i) => {
            const format =
              command.multilevel === "mixed"
                ? ["decimal", "lowerLetter", "lowerRoman"][i % 3]
                : command.multilevel === "bullet"
                  ? "bullet"
                  : command.kind;
            const text =
              command.multilevel === "decimal"
                ? Array.from({ length: i + 1 }, (_, n) => `%${n + 1}`).join(
                    ".",
                  ) + "."
                : format === "bullet"
                  ? command.multilevel === "bullet"
                    ? ["•", "○", "▪"][i % 3]
                    : bullet
                  : (command.pattern?.replace(/%[1-9]/g, `%${i + 1}`) ??
                    `%${i + 1}.`);
            return [
              i,
              {
                start: i === level ? start : 1,
                format,
                text,
                suffix: "tab",
                paragraph: {
                  [p.properties.direction === "rtl"
                    ? "indentRightPt"
                    : "indentLeftPt"]: 36 * (i + 1),
                  firstLinePt: -18,
                  tabs: [{ positionPt: 36 * (i + 1), alignment: "num" }],
                },
                run: {},
              },
            ];
          },
        );
        this.definitions.set(id, { id, levels, preset: command.multilevel });
      }
    }
    const direct = { ...p.directProperties, numbering: { id, level } };
    for (const key of [
      "indentStartPt",
      "indentEndPt",
      "indentLeftPt",
      "indentRightPt",
      "firstLinePt",
    ] as const)
      delete direct[key];
    return {
      ...p,
      formattingContext: context,
      directProperties: direct,
      properties: { ...p.properties, numbering: { id, level } },
    };
  }
  markers(paragraphs: Paragraph[], language?: string) {
    const counters = new Map<string, number[]>(),
      changes = new Map<string, Paragraph>();
    for (const p of paragraphs) {
      const number = p.properties.numbering,
        definition = number?.id ? this.definitions.get(number.id) : undefined;
      if (!definition || number?.id === "0") continue;
      const levels = new Map(definition.levels),
        index = number?.level ?? 0,
        level = levels.get(index);
      if (!level) continue;
      const values = counters.get(definition.id) ?? [];
      values[index] =
        values[index] === undefined ? level.start : values[index] + 1;
      for (let i = index + 1; i < 9; i++)
        if (
          levels.get(i)?.restart !== 0 &&
          (levels.get(i)?.restart ?? i) > index
        )
          delete values[i];
      counters.set(definition.id, values);
      const marker =
        level.format === "none"
          ? ""
          : level.format === "bullet"
            ? level.text
            : level.text.replace(/%([1-9])/g, (_, n: string) => {
                const i = Number(n) - 1,
                  d = levels.get(i);
                return formatCounter(
                  values[i] ?? d?.start ?? 1,
                  d?.format ?? "decimal",
                );
              });
      const locale =
        p.properties.direction === "rtl"
          ? /[پچژگکی۰-۹]/u.test(paragraphText(p))
            ? "fa-IR"
            : language
          : undefined;
      changes.set(p.id!, {
        ...p,
        listMarkProperties: level.run,
        formattingContext: {
          paragraph: { ...level.paragraph, ...p.formattingContext?.paragraph },
          run: p.formattingContext?.run ?? {},
        },
        properties: {
          ...p.properties,
          list: {
            marker: localizeDigits(marker, locale),
            suffix: level.suffix,
            properties: level.run,
            format: level.format,
            locale,
          },
        },
      });
    }
    return changes;
  }
  snapshot() {
    return structuredClone([...this.definitions.values()]);
  }
}
