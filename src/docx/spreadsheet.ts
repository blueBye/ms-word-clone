import { openDocxPackage } from "./package.js";
import { parseXml, descendants, childElements } from "./parser/xml.js";
export async function spreadsheetRows(bytes: Uint8Array) {
  const pkg = await openDocxPackage(bytes, {
    maxUncompressedBytes: 20_000_000,
  });
  const path = pkg
    .listParts()
    .find((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p));
  if (!path) return [];
  const strings = pkg.hasPart("xl/sharedStrings.xml")
    ? descendants(
        parseXml(await pkg.readText("xl/sharedStrings.xml")).documentElement!,
        "si",
      ).map((si) =>
        descendants(si, "t")
          .map((t) => t.textContent ?? "")
          .join(""),
      )
    : [];
  const rows: string[][] = [];
  for (const row of descendants(
    parseXml(await pkg.readText(path)).documentElement!,
    "row",
  ).slice(0, 200)) {
    const values: string[] = [];
    for (const cell of childElements(row).filter((e) => e.localName === "c")) {
      const letters = /^[A-Z]+/.exec(cell.getAttribute("r") ?? "A")?.[0] ?? "A",
        column =
          [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
      if (column > 49) continue;
      while (values.length <= column) values.push("");
      const value = descendants(cell, "v")[0]?.textContent ?? "",
        formula = descendants(cell, "f")[0]?.textContent;
      values[column] = formula
        ? "=" + formula
        : cell.getAttribute("t") === "s"
          ? (strings[Number(value)] ?? "")
          : cell.getAttribute("t") === "inlineStr"
            ? descendants(cell, "t")
                .map((t) => t.textContent ?? "")
                .join("")
            : value;
    }
    rows.push(values);
  }
  const columns = Math.max(1, ...rows.map((r) => r.length));
  return rows.map((row) =>
    Array.from({ length: columns }, (_, c) => row[c] ?? ""),
  );
}
