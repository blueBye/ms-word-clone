export function changeCase(text: string, mode: string, locale?: string) {
  if (mode === "upper") return text.toLocaleUpperCase(locale);
  if (mode === "lower") return text.toLocaleLowerCase(locale);
  if (mode === "toggle")
    return Array.from(text, (c) =>
      c === c.toLocaleUpperCase(locale)
        ? c.toLocaleLowerCase(locale)
        : c.toLocaleUpperCase(locale),
    ).join("");
  if (mode === "title")
    return Array.from(
      new Intl.Segmenter(locale, { granularity: "word" }).segment(text),
      (part) =>
        part.isWordLike
          ? part.segment[0].toLocaleUpperCase(locale) +
            part.segment.slice(1).toLocaleLowerCase(locale)
          : part.segment,
    ).join("");
  if (mode === "sentence")
    return text
      .toLocaleLowerCase(locale)
      .replace(
        /(^|[.!?؟]\s+)(\p{L})/gu,
        (_, prefix: string, letter: string) =>
          prefix + letter.toLocaleUpperCase(locale),
      );
  return text;
}
