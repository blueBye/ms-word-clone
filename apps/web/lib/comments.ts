/** A point comment attaches to its word, or to the paragraph when between words. */
export function commentSelection(
  text: string,
  range: { start: number; end: number },
) {
  if (range.start !== range.end || !text.length) return range;
  const word = [
    ...new Intl.Segmenter(undefined, { granularity: "word" }).segment(text),
  ].find(
    (s) =>
      s.isWordLike &&
      range.start >= s.index &&
      range.start <= s.index + s.segment.length,
  );
  return word
    ? { start: word.index, end: word.index + word.segment.length }
    : { start: 0, end: text.length };
}
