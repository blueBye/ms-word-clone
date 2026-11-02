import { expect, it } from "vitest";
import {
  textChange,
  literalMatches,
  changeCase,
} from "../apps/web/lib/editor.js";

it("maps literal matches to original Unicode offsets and escapes regex characters", () => {
  expect(literalMatches("x İ Word WORD", "word")).toEqual([
    { start: 4, end: 8 },
    { start: 9, end: 13 },
  ]);
  expect(literalMatches("a.b a+b [x]", "a.b")).toEqual([{ start: 0, end: 3 }]);
  expect(literalMatches("سلام\u200cدنیا سلام", "سلام", true)).toHaveLength(2);
});
it("keeps grapheme boundaries in typed changes and leaves Persian unchanged in case commands", () => {
  expect(textChange("é 👩‍💻", "è 👩‍💻")).toEqual({ start: 0, end: 2, text: "è" });
  expect(changeCase("hello فارسی WORLD", "title")).toBe("Hello فارسی World");
  expect(changeCase("hello. WORLD", "sentence")).toBe("Hello. World");
});
