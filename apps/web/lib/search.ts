export interface SearchOptions {
  matchCase?: boolean;
  wholeWords?: boolean;
  prefix?: boolean;
  suffix?: boolean;
  ignorePunctuation?: boolean;
  ignoreWhitespace?: boolean;
}
export interface SearchHit {
  id: string;
  start: number;
  end: number;
}
export type SearchDirection = "down" | "up";

/** Every normalized character retains its source span, including surrogate pairs and ZWNJ. */
export function searchMatches(
  text: string,
  query: string,
  options: SearchOptions = {},
) {
  const normalize = (value: string) => {
    let normalized = "",
      offset = 0;
    const spans: { start: number; end: number }[] = [];
    for (const char of value) {
      const start = offset;
      offset += char.length;
      if (
        (options.ignorePunctuation && /\p{P}/u.test(char)) ||
        (options.ignoreWhitespace && /\s/u.test(char))
      )
        continue;
      normalized += char;
      for (let i = 0; i < char.length; i++) spans.push({ start, end: offset });
    }
    return { normalized, spans };
  };
  const source = normalize(text),
    needle = normalize(query).normalized;
  if (!needle) return [];
  const escaped = needle.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const word = /[\p{L}\p{N}\p{M}\p{Pc}\u200c\u200d]/u;
  return Array.from(
    source.normalized.matchAll(
      new RegExp(escaped, options.matchCase ? "gu" : "giu"),
    ),
  ).flatMap((match) => {
    const start = source.spans[match.index!].start,
      end = source.spans[match.index! + match[0].length - 1].end;
    const before = Array.from(text.slice(0, start)).at(-1) ?? "";
    const after = Array.from(text.slice(end))[0] ?? "";
    if (
      ((options.wholeWords || options.prefix) && word.test(before)) ||
      ((options.wholeWords || options.suffix) && word.test(after))
    )
      return [];
    return [{ start, end }];
  });
}

export function nextMatchIndex(
  matches: SearchHit[],
  ids: string[],
  origin: SearchHit | null,
  direction: SearchDirection,
) {
  if (!matches.length) return -1;
  if (!origin) return direction === "up" ? matches.length - 1 : 0;
  const current = matches.findIndex(
    (hit) =>
      hit.id === origin.id &&
      hit.start === origin.start &&
      hit.end === origin.end,
  );
  if (current >= 0)
    return (
      (current + (direction === "up" ? -1 : 1) + matches.length) %
      matches.length
    );
  const paragraph = ids.indexOf(origin.id);
  const candidates = matches.map((hit, index) => ({ ...hit, index }));
  if (direction === "up") candidates.reverse();
  return (
    candidates.find((hit) =>
      direction === "up"
        ? ids.indexOf(hit.id) < paragraph ||
          (hit.id === origin.id && hit.end <= origin.start)
        : ids.indexOf(hit.id) > paragraph ||
          (hit.id === origin.id && hit.start >= origin.end),
    )?.index ?? (direction === "up" ? matches.length - 1 : 0)
  );
}
