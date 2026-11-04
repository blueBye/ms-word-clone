import { expect, it } from "vitest";
import { searchMatches, nextMatchIndex } from "../apps/web/lib/search.js";

it("finds every occurrence with original offsets, case sensitivity, literal punctuation and mixed scripts", () => {
  expect(searchMatches("one ONE one", "one")).toEqual([
    { start: 0, end: 3 },
    { start: 4, end: 7 },
    { start: 8, end: 11 },
  ]);
  expect(searchMatches("one ONE one", "one", { matchCase: true })).toHaveLength(
    2,
  );
  expect(searchMatches("a.b a+b", "a.b")).toEqual([{ start: 0, end: 3 }]);
  const text = "👩‍💻 متن فارسی Word فارسی";
  expect(
    searchMatches(text, "فارسی").map((hit) => text.slice(hit.start, hit.end)),
  ).toEqual(["فارسی", "فارسی"]);
});
it("uses Unicode word edges, retaining ZWNJ as part of Persian words", () => {
  expect(
    searchMatches("word words sword word", "word", { wholeWords: true }),
  ).toEqual([
    { start: 0, end: 4 },
    { start: 17, end: 21 },
  ]);
  expect(
    searchMatches("word words sword", "word", { prefix: true }),
  ).toHaveLength(2);
  expect(
    searchMatches("word words sword", "word", { suffix: true }),
  ).toHaveLength(2);
  const text = "می‌روم می روم";
  expect(searchMatches(text, "می", { wholeWords: true })).toEqual([
    { start: 7, end: 9 },
  ]);
  expect(searchMatches("e\u0301 e", "e", { wholeWords: true })).toEqual([
    { start: 3, end: 4 },
  ]);
});
it("maps ignored punctuation and spacing back to exact source ranges for replacement", () => {
  const text = "😀 cyber- threat\tintelligence!";
  const hits = searchMatches(text, "cyberthreatintelligence", {
    ignorePunctuation: true,
    ignoreWhitespace: true,
  });
  expect(hits).toHaveLength(1);
  expect(text.slice(hits[0].start, hits[0].end)).toBe(
    "cyber- threat\tintelligence",
  );
  expect(
    searchMatches(text, " -!", {
      ignoreWhitespace: true,
      ignorePunctuation: true,
    }),
  ).toEqual([]);
});
it("navigates repeated matches in each paragraph, moves in both directions and wraps", () => {
  const hits = [
    { id: "p1", start: 0, end: 3 },
    { id: "p1", start: 5, end: 8 },
    { id: "p3", start: 2, end: 5 },
  ];
  const ids = ["p1", "p2", "p3"];
  expect(nextMatchIndex(hits, ids, hits[0], "down")).toBe(1);
  expect(nextMatchIndex(hits, ids, hits[2], "down")).toBe(0);
  expect(nextMatchIndex(hits, ids, hits[0], "up")).toBe(2);
  expect(
    nextMatchIndex(hits, ids, { id: "p2", start: 0, end: 0 }, "down"),
  ).toBe(2);
  expect(nextMatchIndex(hits, ids, { id: "p2", start: 0, end: 0 }, "up")).toBe(
    1,
  );
  expect(
    nextMatchIndex(hits, ids, { id: "p1", start: 0, end: 0 }, "down"),
  ).toBe(0);
  expect(nextMatchIndex([], ids, null, "down")).toBe(-1);
});
