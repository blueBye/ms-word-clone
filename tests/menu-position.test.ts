import { expect, it } from "vitest";
import { menuPosition } from "../apps/web/lib/menu-position.js";

it("anchors dropdowns directly below their button instead of the viewport bottom", () => {
  expect(
    menuPosition(
      { left: 320, right: 344, top: 80, bottom: 104 },
      { width: 250, height: 300 },
      { width: 1280, height: 800 },
    ),
  ).toEqual({ left: 320, top: 106, maxHeight: 686, clipped: false });
});
it("keeps a narrow-screen dropdown anchored and scrolls its oversized content", () => {
  expect(
    menuPosition(
      { left: 280, right: 304, top: 120, bottom: 144 },
      { width: 280, height: 600 },
      { width: 360, height: 480 },
    ),
  ).toEqual({ left: 72, top: 146, maxHeight: 326, clipped: true });
});
it("flips above the trigger near the bottom without stretching the menu", () => {
  expect(
    menuPosition(
      { left: 200, right: 224, top: 650, bottom: 674 },
      { width: 210, height: 300 },
      { width: 1000, height: 750 },
    ),
  ).toEqual({ left: 200, top: 348, maxHeight: 640, clipped: false });
});
