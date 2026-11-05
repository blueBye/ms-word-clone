import { XMLSerializer, type Element } from "@xmldom/xmldom";
import { zipSync, strToU8 } from "fflate";
import type { DrawingInline } from "../model/document.js";
import { parseXml, childElements, descendants } from "./parser/xml.js";
type Chart = NonNullable<DrawingInline["chart"]>;
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart",
  A = "http://schemas.openxmlformats.org/drawingml/2006/main",
  serialize = (root: Element) =>
    new XMLSerializer().serializeToString(root.ownerDocument!);
function add(
  p: Element,
  ns: string,
  name: string,
  attrs: Record<string, string | number> = {},
) {
  const n = p.ownerDocument!.createElementNS(ns, name);
  for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, String(v));
  p.appendChild(n);
  return n;
}
export function chartXml(chart: Chart, relationship: string) {
  const root = parseXml(
      `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`,
    ).documentElement!,
    body = add(root, C, "c:chart");
  const title = add(add(add(body, C, "c:title"), C, "c:tx"), C, "c:rich");
  add(title, A, "a:bodyPr");
  add(title, A, "a:lstStyle");
  add(add(add(title, A, "a:p"), A, "a:r"), A, "a:t").textContent = chart.title;
  const plot = add(body, C, "c:plotArea");
  add(plot, C, "c:layout");
  const drawing = add(plot, C, `c:${chart.kind}Chart`);
  if (chart.kind === "bar") add(drawing, C, "c:barDir", { val: "col" });
  if (chart.kind !== "pie")
    add(drawing, C, "c:grouping", {
      val: chart.kind === "bar" ? "clustered" : "standard",
    });
  const cache = (
    p: Element,
    kind: string,
    values: (number | string)[],
    formula: string,
  ) => {
    const ref = add(p, C, `c:${kind}Ref`);
    add(ref, C, "c:f").textContent = formula;
    const cache = add(ref, C, `c:${kind}Cache`);
    if (kind === "num") add(cache, C, "c:formatCode").textContent = "General";
    add(cache, C, "c:ptCount", { val: values.length });
    values.forEach(
      (v, i) =>
        (add(add(cache, C, "c:pt", { idx: i }), C, "c:v").textContent =
          String(v)),
    );
  };
  chart.series.forEach((series, i) => {
    const ser = add(drawing, C, "c:ser");
    add(ser, C, "c:idx", { val: i });
    add(ser, C, "c:order", { val: i });
    const column = String.fromCharCode(66 + i);
    cache(add(ser, C, "c:tx"), "str", [series.name], `Sheet1!$${column}$1`);
    cache(
      add(ser, C, "c:cat"),
      "str",
      chart.categories,
      `Sheet1!$A$2:$A$${chart.categories.length + 1}`,
    );
    cache(
      add(ser, C, "c:val"),
      "num",
      series.values,
      `Sheet1!$${column}$2:$${column}$${series.values.length + 1}`,
    );
  });
  if (chart.kind !== "pie") {
    add(drawing, C, "c:axId", { val: 1 });
    add(drawing, C, "c:axId", { val: 2 });
    for (const [name, id, cross, pos] of [
      ["catAx", 1, 2, "b"],
      ["valAx", 2, 1, "l"],
    ] as const) {
      const axis = add(plot, C, `c:${name}`);
      add(axis, C, "c:axId", { val: id });
      add(add(axis, C, "c:scaling"), C, "c:orientation", { val: "minMax" });
      add(axis, C, "c:axPos", { val: pos });
      add(axis, C, "c:crossAx", { val: cross });
      add(axis, C, "c:crosses", { val: "autoZero" });
    }
  }
  add(add(body, C, "c:legend"), C, "c:legendPos", { val: "b" });
  add(body, C, "c:plotVisOnly", { val: 1 });
  const external = add(root, C, "c:externalData");
  external.setAttributeNS(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "r:id",
    relationship,
  );
  add(external, C, "c:autoUpdate", { val: 0 });
  return serialize(root);
}
export function chartWorkbook(chart: Chart) {
  return workbookRows([
    ["Category", ...chart.series.map((s) => s.name)],
    ...chart.categories.map((c, i) => [
      c,
      ...chart.series.map((s) => s.values[i]),
    ]),
  ]);
}
export function workbookRows(rows: (string | number)[][]) {
  const columnName = (index: number) => {
    let name = "";
    do {
      name = String.fromCharCode(65 + (index % 26)) + name;
      index = Math.floor(index / 26) - 1;
    } while (index >= 0);
    return name;
  };
  const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    sheet = parseXml(
      `<worksheet xmlns="${S}"><sheetData/></worksheet>`,
    ).documentElement!,
    data = childElements(sheet)[0];
  rows.forEach((values, r) => {
    const row = add(data, S, "row", { r: r + 1 });
    values.forEach((value, c) => {
      const cell = add(row, S, "c", {
        r: `${columnName(c)}${r + 1}`,
        ...(typeof value === "number" ? {} : { t: "inlineStr" }),
      });
      if (typeof value === "number")
        add(cell, S, "v").textContent = String(value);
      else if (String(value).startsWith("=")) {
        cell.removeAttribute("t");
        add(cell, S, "f").textContent = String(value).slice(1);
        add(cell, S, "v").textContent = "0";
      } else add(add(cell, S, "is"), S, "t").textContent = String(value);
    });
  });
  const files = {
    "[Content_Types].xml":
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    "_rels/.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml": `<workbook xmlns="${S}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": serialize(sheet),
  };
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([path, value]) => [path, strToU8(value)]),
    ),
  );
}
export function readChart(root: Element): Chart | undefined {
  const drawing =
    descendants(root, "barChart")[0] ??
    descendants(root, "lineChart")[0] ??
    descendants(root, "pieChart")[0];
  if (!drawing) return;
  const series = childElements(drawing).filter((e) => e.localName === "ser"),
    values = (root: Element | undefined) =>
      root
        ? descendants(root, "pt").map(
            (p) => descendants(p, "v")[0]?.textContent ?? "",
          )
        : [];
  return {
    kind: drawing.localName!.replace("Chart", "") as Chart["kind"],
    title: descendants(descendants(root, "title")[0] ?? root, "t")
      .map((t) => t.textContent ?? "")
      .join(""),
    categories: values(
      series[0] && childElements(series[0]).find((e) => e.localName === "cat"),
    ),
    series: series.map((ser) => ({
      name:
        values(childElements(ser).find((e) => e.localName === "tx"))[0] ??
        "Series",
      values: values(childElements(ser).find((e) => e.localName === "val")).map(
        Number,
      ),
    })),
  };
}
