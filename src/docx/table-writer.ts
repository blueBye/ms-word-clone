import type { Element } from "@xmldom/xmldom";
import type { Table, Block, Paragraph } from "../model/document.js";
import { getParagraphs } from "../model/traversal.js";
import { childElements, childWord, isWord } from "./parser/xml.js";
export function writeTable(
  root: Element,
  table: Table,
  resolve: (id: string) => Element | undefined,
  onParagraph: (element: Element, id: string) => void,
  rebuild: boolean,
  original?: Table,
) {
  const w = root.namespaceURI!;
  const originalCells = new Map(
    original?.rows.flatMap((row) =>
      row.cells.map(
        (cell) => [getParagraphs(cell.blocks)[0]?.id, cell] as const,
      ),
    ),
  );
  const orders: Record<string, string[]> = {
    tblPr: [
      "tblStyle",
      "tblpPr",
      "tblOverlap",
      "bidiVisual",
      "tblStyleRowBandSize",
      "tblStyleColBandSize",
      "tblW",
      "jc",
      "tblCellSpacing",
      "tblInd",
      "tblBorders",
      "shd",
      "tblLayout",
      "tblCellMar",
      "tblLook",
      "tblCaption",
      "tblDescription",
      "tblPrChange",
    ],
    trPr: [
      "cnfStyle",
      "divId",
      "gridBefore",
      "gridAfter",
      "wBefore",
      "wAfter",
      "cantSplit",
      "trHeight",
      "tblHeader",
      "tblCellSpacing",
      "jc",
      "hidden",
      "ins",
      "del",
      "trPrChange",
    ],
    tcPr: [
      "cnfStyle",
      "tcW",
      "gridSpan",
      "hMerge",
      "vMerge",
      "tcBorders",
      "shd",
      "noWrap",
      "tcMar",
      "textDirection",
      "tcFitText",
      "vAlign",
      "hideMark",
      "tcPrChange",
    ],
  };
  const make = (parent: Element, name: string) => {
    const existing = childWord(parent, name);
    if (existing) return existing;
    const node = root.ownerDocument!.createElementNS(w, `w:${name}`);
    const order = orders[parent.localName ?? ""],
      rank = order?.indexOf(name) ?? -1;
    const before =
      rank >= 0
        ? childElements(parent).find(
            (n) => (order?.indexOf(n.localName ?? "") ?? -1) > rank,
          )
        : undefined;
    parent.insertBefore(node, before ?? null);
    return node;
  };
  const value = (node: Element, key: string, n: string | number) =>
    node.setAttributeNS(w, `w:${key}`, String(n));
  const set = (
    parent: Element,
    name: string,
    attrs: Record<string, string | number>,
  ) => {
    const node = make(parent, name);
    for (const [k, v] of Object.entries(attrs)) value(node, k, v);
    return node;
  };
  const remove = (parent: Element, name: string) => {
    const node = childWord(parent, name);
    if (node) parent.removeChild(node);
  };
  const borders = (
    parent: Element,
    name: string,
    values: Table["properties"]["borders"],
  ) => {
    if (!values) return;
    const box = make(parent, name);
    for (const [side, b] of Object.entries(values))
      if (b)
        set(box, side, {
          val: b.style,
          sz: Math.round(b.widthPt * 8),
          color: b.color.replace("#", ""),
        });
  };
  const pr = make(root, "tblPr");
  if (root.firstChild !== pr) root.insertBefore(pr, root.firstChild);
  if (table.properties.width)
    set(pr, "tblW", {
      type: table.properties.width.unit === "percent" ? "pct" : "dxa",
      w: Math.round(
        table.properties.width.value *
          (table.properties.width.unit === "percent" ? 50 : 20),
      ),
    });
  if (table.properties.layout)
    set(pr, "tblLayout", { type: table.properties.layout });
  if (table.properties.alignment)
    set(pr, "jc", { val: table.properties.alignment });
  if (table.properties.direction)
    set(pr, "bidiVisual", {
      val: table.properties.direction === "rtl" ? 1 : 0,
    });
  if (table.properties.cellMargins) {
    const margins = make(pr, "tblCellMar");
    for (const [side, n] of Object.entries(table.properties.cellMargins))
      set(margins, side, { type: "dxa", w: Math.round(n * 20) });
  }
  borders(pr, "tblBorders", table.properties.borders);
  const grid = make(root, "tblGrid");
  if (grid !== pr.nextSibling) root.insertBefore(grid, pr.nextSibling);
  for (const n of childElements(grid)) grid.removeChild(n);
  for (const width of table.columnsPt) {
    const n = root.ownerDocument!.createElementNS(w, "w:gridCol");
    value(n, "w", Math.round(width * 20));
    grid.appendChild(n);
  }
  const oldRows = childElements(root).filter((n) => isWord(n, "tr")),
    usedCells = new Set<Element>(),
    usedRows = new Set<Element>();
  const closest = (node: Element | undefined, name: string) => {
    for (
      let e = node;
      e && e !== root;
      e = e.parentNode?.nodeType === 1 ? (e.parentNode as Element) : undefined
    )
      if (isWord(e, name)) return e;
  };
  const paragraph = (model: Paragraph) => {
    const p = model.id && resolve(model.id);
    if (p) return p;
    const n = root.ownerDocument!.createElementNS(w, "w:p");
    if (model.directProperties?.direction || model.directProperties?.styleId) {
      const props = make(n, "pPr");
      if (model.directProperties.styleId)
        set(props, "pStyle", { val: model.directProperties.styleId });
      if (model.directProperties.direction)
        set(props, "bidi", {
          val: model.directProperties.direction === "rtl" ? 1 : 0,
        });
    }
    const run = make(n, "r"),
      text = make(run, "t");
    text.textContent = model.children
      .filter((r) => r.type === "text")
      .map((r) => r.text)
      .join("");
    text.setAttributeNS(
      "http://www.w3.org/XML/1998/namespace",
      "xml:space",
      "preserve",
    );
    if (model.id) onParagraph(n, model.id);
    return n;
  };
  const block = (b: Block): Element => {
    if (b.type === "paragraph") return paragraph(b);
    const p = getParagraphs([b])[0],
      original = p?.id && resolve(p.id),
      table = closest(original || undefined, "tbl");
    if (!table) throw new Error("Cannot safely move this nested table.");
    return table;
  };
  const resultRows: Element[] = [];
  table.rows.forEach((row, r) => {
    const first = getParagraphs(row.cells[0]?.blocks ?? [])[0],
      source = first?.id && resolve(first.id);
    let target = rebuild ? closest(source || undefined, "tr") : oldRows[r];
    if (!target || usedRows.has(target))
      target = root.ownerDocument!.createElementNS(w, "w:tr");
    usedRows.add(target);
    resultRows.push(target);
    const rowPr = make(target, "trPr");
    if (rowPr !== target.firstChild)
      target.insertBefore(rowPr, target.firstChild);
    if (row.header) set(rowPr, "tblHeader", { val: 1 });
    else remove(rowPr, "tblHeader");
    if (row.cantSplit) set(rowPr, "cantSplit", { val: 1 });
    else remove(rowPr, "cantSplit");
    if (row.heightPt)
      set(rowPr, "trHeight", {
        val: Math.round(row.heightPt * 20),
        hRule: row.heightRule ?? "atLeast",
      });
    const cells: Element[] = [];
    row.cells.forEach((cell, c) => {
      const p = getParagraphs(cell.blocks)[0],
        source = p?.id && resolve(p.id);
      const previous = originalCells.get(p?.id);
      let tc = rebuild
        ? closest(source || undefined, "tc")
        : childElements(target).filter((e) => isWord(e, "tc"))[c];
      if (!tc || usedCells.has(tc))
        tc = root.ownerDocument!.createElementNS(w, "w:tc");
      usedCells.add(tc);
      cells.push(tc);
      const props = make(tc, "tcPr");
      if (props !== tc.firstChild) tc.insertBefore(props, tc.firstChild);
      const width = table.columnsPt
        .slice(cell.column, cell.column + cell.columnSpan)
        .reduce((n, v) => n + v, 0);
      set(props, "tcW", { type: "dxa", w: Math.round(width * 20) });
      if (cell.columnSpan > 1) set(props, "gridSpan", { val: cell.columnSpan });
      else remove(props, "gridSpan");
      if (cell.continuation) set(props, "vMerge", { val: "continue" });
      else if (cell.rowSpan > 1) set(props, "vMerge", { val: "restart" });
      else remove(props, "vMerge");
      if (!previous || cell.backgroundColor !== previous.backgroundColor)
        set(props, "shd", {
          val: "clear",
          fill: cell.backgroundColor?.replace("#", "") ?? "auto",
        });
      if (
        cell.verticalAlign &&
        (!previous || cell.verticalAlign !== previous.verticalAlign)
      )
        set(props, "vAlign", { val: cell.verticalAlign });
      if (
        !previous ||
        JSON.stringify(cell.borders) !== JSON.stringify(previous.borders)
      )
        borders(props, "tcBorders", cell.borders);
      if (rebuild) {
        const blocks = cell.blocks.map(block);
        for (const node of childElements(tc))
          if (!isWord(node, "tcPr")) tc.removeChild(node);
        for (const node of blocks) tc.appendChild(node);
        if (!blocks.length || !isWord(blocks.at(-1)!, "p"))
          tc.appendChild(root.ownerDocument!.createElementNS(w, "w:p"));
      }
    });
    if (rebuild) {
      for (const node of childElements(target))
        if (isWord(node, "tc")) target.removeChild(node);
      for (const tc of cells) target.appendChild(tc);
    }
  });
  if (rebuild) {
    for (const row of oldRows)
      if (row.parentNode === root) root.removeChild(row);
    for (const row of resultRows) root.appendChild(row);
  }
}
