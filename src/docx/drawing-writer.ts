import type { Element } from "@xmldom/xmldom";
import type { DrawingInline, Paragraph } from "../model/document.js";
import { parseXml, childElements, descendants } from "./parser/xml.js";
import { writeTextEffects } from "./text-effects.js";
import type { RunFormat } from "../core/types.js";
import { parseEquation, type MathNode } from "../core/equation.js";
import { iconPaths, vectorCommands } from "../model/graphics.js";
import { videoUrl } from "../core/assets.js";
import { OLE_PREVIEW } from "./ole-preview.js";
import { wordAlignment } from "./alignment.js";
import { writeImageEffects } from "./image-effects.js";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main",
  WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
function add(
  parent: Element,
  ns: string,
  name: string,
  attrs: Record<string, string | number> = {},
) {
  const n = parent.ownerDocument!.createElementNS(ns, name);
  for (const [key, value] of Object.entries(attrs))
    n.setAttribute(key, String(value));
  parent.appendChild(n);
  return n;
}
const emu = (n: number) => Math.round(n * 12700);
function fill(parent: Element, color?: string) {
  if (color)
    add(add(parent, A, "a:solidFill"), A, "a:srgbClr", { val: color.slice(1) });
  else add(parent, A, "a:noFill");
}
function paragraph(parent: Element, p: Paragraph, w: string) {
  const n = add(parent, w, "w:p"),
    pr = add(n, w, "w:pPr");
  const val = (
    parent: Element,
    name: string,
    key: string,
    value: string | number,
  ) => {
    const e = add(parent, w, `w:${name}`);
    e.setAttributeNS(w, `w:${key}`, String(value));
    return e;
  };
  if (p.properties.styleId) val(pr, "pStyle", "val", p.properties.styleId);
  if (p.properties.direction)
    val(pr, "bidi", "val", p.properties.direction === "rtl" ? 1 : 0);
  if (p.properties.alignment)
    val(
      pr,
      "jc",
      "val",
      wordAlignment(p.properties.alignment, p.properties.direction),
    );
  for (const r of p.children)
    if (r.type === "text") {
      const run = add(n, w, "w:r"),
        rp = add(run, w, "w:rPr"),
        f = r.properties;
      const fonts = add(rp, w, "w:rFonts");
      for (const [slot, value] of Object.entries({
        ascii: f.fontFamily,
        hAnsi: f.fontFamily,
        cs: f.complexScript?.fontFamily ?? f.fontFamily,
      }))
        if (value) fonts.setAttributeNS(w, `w:${slot}`, value);
      for (const [name, value] of Object.entries({
        b: f.bold,
        i: f.italic,
        strike: f.strike,
        bCs: f.complexScript?.bold,
        iCs: f.complexScript?.italic,
      }))
        if (value !== undefined) val(rp, name, "val", value ? 1 : 0);
      if (f.fontSizePt) val(rp, "sz", "val", Math.round(f.fontSizePt * 2));
      if (f.complexScript?.fontSizePt)
        val(rp, "szCs", "val", Math.round(f.complexScript.fontSizePt * 2));
      if (f.color) val(rp, "color", "val", f.color.replace("#", ""));
      if (f.underline) val(rp, "u", "val", f.underline);
      writeTextEffects(
        rp,
        { ...f, underline: f.underline as RunFormat["underline"] },
        f.color,
      );
      const text = add(run, w, "w:t");
      text.setAttributeNS(
        "http://www.w3.org/XML/1998/namespace",
        "xml:space",
        "preserve",
      );
      text.appendChild(text.ownerDocument!.createTextNode(r.text));
    }
}
export async function drawingRun(
  parent: Element,
  d: DrawingInline,
  drawingId: number,
  imageRelation: (dataUrl: string) => Promise<string>,
  chartRelation?: (
    chart: NonNullable<DrawingInline["chart"]>,
  ) => Promise<string>,
  assetRelation?: (
    asset: { dataUrl: string; name?: string; mime?: string; progId?: string },
    kind: "model3d" | "attachment",
  ) => Promise<{ id: string; progId?: string }>,
  hyperlinkRelation?: (url: string) => Promise<string>,
): Promise<Element> {
  if (d.equation) {
    const M = "http://schemas.openxmlformats.org/officeDocument/2006/math",
      math = parseXml(`<m:oMath xmlns:m="${M}"/>`).documentElement!;
    const write = (parent: Element, node: MathNode) => {
      if (node.kind === "text")
        add(add(parent, M, "m:r"), M, "m:t").textContent = node.text;
      else if (node.kind === "row") node.items.forEach((n) => write(parent, n));
      else if (node.kind === "sqrt") {
        const rad = add(parent, M, "m:rad");
        add(add(rad, M, "m:radPr"), M, "m:degHide").setAttributeNS(
          M,
          "m:val",
          "1",
        );
        add(rad, M, "m:deg");
        write(add(rad, M, "m:e"), node.a);
      } else {
        const root = add(
          parent,
          M,
          node.kind === "frac"
            ? "m:f"
            : node.kind === "sup"
              ? "m:sSup"
              : "m:sSub",
        );
        write(add(root, M, node.kind === "frac" ? "m:num" : "m:e"), node.a);
        write(
          add(
            root,
            M,
            node.kind === "frac"
              ? "m:den"
              : node.kind === "sup"
                ? "m:sup"
                : "m:sub",
          ),
          node.b,
        );
      }
    };
    write(math, parseEquation(d.equation.expression));
    return parent.ownerDocument!.importNode(math, true) as Element;
  }
  const w = parent.namespaceURI!,
    r = w.includes("purl.oclc.org")
      ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
      : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  if (d.attachment || d.spreadsheet) {
    const attachment = d.spreadsheet
      ? {
          name: d.attachment?.name.endsWith(".xlsx")
            ? d.attachment.name
            : "Worksheet.xlsx",
          dataUrl: "",
          progId: "Excel.Sheet.12",
          mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
      : d.attachment!;
    if (!assetRelation) throw new Error("Object embedding is unavailable.");
    const id = await assetRelation(
        {
          ...attachment,
          ...(d.spreadsheet ? { cells: d.spreadsheet.cells } : {}),
        },
        "attachment",
      ),
      V = "urn:schemas-microsoft-com:vml",
      O = "urn:schemas-microsoft-com:office:office";
    const run = parseXml(
        `<w:r xmlns:w="${w}" xmlns:r="${r}" xmlns:v="${V}" xmlns:o="${O}"/>`,
      ).documentElement!,
      object = add(run, w, "w:object"),
      shapeId = `_x0000_i${1000 + drawingId}`;
    object.setAttributeNS(w, "w:dxaOrig", String(Math.round(d.widthPt * 20)));
    object.setAttributeNS(w, "w:dyaOrig", String(Math.round(d.heightPt * 20)));
    const frame = add(object, V, "v:shapetype", {
      id: `_editor_t75_${drawingId}`,
      coordsize: "21600,21600",
      path: "m0,0l21600,0,21600,21600,0,21600xe",
      filled: "f",
      stroked: "f",
    });
    frame.setAttributeNS(O, "o:spt", "75");
    const shape = add(object, V, "v:shape", {
      id: shapeId,
      type: `#_editor_t75_${drawingId}`,
      alt: attachment.name,
      style: `width:${d.widthPt}pt;height:${d.heightPt}pt`,
      fillcolor: "#EEF2F6",
      strokecolor: "#71879B",
    });
    shape.setAttributeNS(O, "o:ole", "");
    const preview = add(shape, V, "v:imagedata");
    preview.setAttributeNS(
      r,
      "r:id",
      await imageRelation(d.image?.dataUrl ?? OLE_PREVIEW),
    );
    preview.setAttributeNS(O, "o:title", attachment.name);
    const ole = add(object, O, "o:OLEObject", {
      Type: "Embed",
      ProgID: id.progId ?? attachment.progId ?? "Package",
      ShapeID: shapeId,
      DrawAspect: d.spreadsheet ? "Content" : "Icon",
      ObjectID: `_${1000 + drawingId}`,
    });
    ole.setAttributeNS(r, "r:id", id.id);
    return parent.ownerDocument!.importNode(run, true) as Element;
  }
  const run = parseXml(
    `<w:r xmlns:w="${w}" xmlns:r="${r}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" xmlns:wpg="${WPG}" xmlns:pic="${PIC}"/>`,
  ).documentElement!;
  const root = add(run, w, "w:drawing"),
    pos = d.position;
  const frame = add(
    root,
    WP,
    pos ? "wp:anchor" : "wp:inline",
    pos
      ? {
          distT: 0,
          distB: 0,
          distL: 114300,
          distR: 114300,
          simplePos: 0,
          relativeHeight: Math.max(0, pos.zIndex),
          behindDoc: pos.behindText ? 1 : 0,
          locked: pos.locked ? 1 : 0,
          layoutInCell: 1,
          allowOverlap: 1,
        }
      : { distT: 0, distB: 0, distL: 0, distR: 0 },
  );
  if (pos) {
    add(frame, WP, "wp:simplePos", { x: 0, y: 0 });
    for (const [name, axis] of [
      ["positionH", pos.horizontal],
      ["positionV", pos.vertical],
    ] as const) {
      const p = add(frame, WP, `wp:${name}`, { relativeFrom: axis.relativeTo });
      const value = add(p, WP, axis.alignment ? "wp:align" : "wp:posOffset");
      value.textContent = axis.alignment ?? String(emu(axis.offsetPt ?? 0));
    }
  }
  add(frame, WP, "wp:extent", { cx: emu(d.widthPt), cy: emu(d.heightPt) });
  if (pos) {
    const kind = pos.behindText
      ? "None"
      : pos.wrap.replace(/^wrap/, "").replace(/^./, (c) => c.toUpperCase());
    const wrap = add(
      frame,
      WP,
      `wp:wrap${["None", "TopAndBottom", "Square", "Tight", "Through"].includes(kind) ? kind : "Square"}`,
      ["None", "TopAndBottom"].includes(kind) ? {} : { wrapText: "bothSides" },
    );
    if (kind === "Tight" || kind === "Through") {
      const polygon = add(wrap, WP, "wp:wrapPolygon", { edited: 0 });
      add(polygon, WP, "wp:start", { x: 0, y: 0 });
      for (const [x, y] of [
        [21600, 0],
        [21600, 21600],
        [0, 21600],
        [0, 0],
      ])
        add(polygon, WP, "wp:lineTo", { x, y });
    }
  }
  const prop = add(frame, WP, "wp:docPr", {
    id: drawingId,
    name: d.alt || `Object ${drawingId}`,
    descr: d.alt,
  });
  add(frame, WP, "wp:cNvGraphicFramePr");
  const graphic = add(frame, A, "a:graphic");
  const content = add(graphic, A, "a:graphicData", {
    uri: d.model3d
      ? "http://schemas.microsoft.com/office/drawing/2017/model3d"
      : d.chart
        ? "http://schemas.openxmlformats.org/drawingml/2006/chart"
        : d.children
          ? WPG
          : d.image
            ? PIC
            : WPS,
  });
  if (d.model3d) {
    if (!assetRelation) throw new Error("3D export is unavailable.");
    const M = "http://schemas.microsoft.com/office/drawing/2017/model3d",
      model = add(content, M, "am3d:model3d");
    model.setAttributeNS(
      r,
      "r:embed",
      (await assetRelation(d.model3d, "model3d")).id,
    );
    const pr = add(model, M, "am3d:spPr"),
      x = add(pr, A, "a:xfrm");
    add(x, A, "a:off", { x: 0, y: 0 });
    add(x, A, "a:ext", { cx: emu(d.widthPt), cy: emu(d.heightPt) });
    const size = d.model3d.size ?? 1,
      center = d.model3d.center ?? [0, 0, 0],
      unit = 36000000,
      camera = add(model, M, "am3d:camera");
    add(camera, M, "am3d:pos", {
      x: Math.round((center[0] + size * 0.2) * unit),
      y: Math.round((center[1] + size * 0.1) * unit),
      z: Math.round((center[2] + size * 2) * unit),
    });
    add(camera, M, "am3d:up", { dx: 0, dy: 1, dz: 0 });
    add(camera, M, "am3d:lookAt", {
      x: Math.round(center[0] * unit),
      y: Math.round(center[1] * unit),
      z: Math.round(center[2] * unit),
    });
    add(camera, M, "am3d:perspective", { fov: 2400000 });
    const trans = add(model, M, "am3d:trans");
    add(trans, M, "am3d:meterPerModelUnit", { n: 1, d: 1 });
    if (d.model3d.rotation)
      add(trans, M, "am3d:rot", {
        ax: Math.round(d.model3d.rotation.x * 60000),
        ay: Math.round(d.model3d.rotation.y * 60000),
        az: Math.round(d.model3d.rotation.z * 60000),
      });
    if (d.model3d.poster)
      add(add(model, M, "am3d:raster"), A, "a:blip").setAttributeNS(
        r,
        "r:embed",
        await imageRelation(d.model3d.poster),
      );
    add(model, M, "am3d:objViewport", { viewportSz: Math.round(size * unit) });
    return parent.ownerDocument!.importNode(run, true) as Element;
  }
  if (d.chart) {
    if (!chartRelation) throw new Error("Chart export is unavailable.");
    add(
      content,
      "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "c:chart",
    ).setAttributeNS(r, "r:id", await chartRelation(d.chart));
    return parent.ownerDocument!.importNode(run, true) as Element;
  }
  let childIdentity = drawingId * 1000;
  const shape = async (
    host: Element,
    item: DrawingInline,
    child = false,
  ): Promise<void> => {
    if (item.children) {
      const g = add(host, WPG, "wpg:wgp");
      add(g, WPG, "wpg:cNvGrpSpPr");
      const pr = add(g, WPG, "wpg:grpSpPr"),
        x = add(pr, A, "a:xfrm", {
          rot: Math.round((item.shape?.rotation ?? 0) * 60000),
          flipH: item.flipH ? 1 : 0,
          flipV: item.flipV ? 1 : 0,
        });
      add(x, A, "a:off", {
        x: child ? emu(item.position?.horizontal.offsetPt ?? 0) : 0,
        y: child ? emu(item.position?.vertical.offsetPt ?? 0) : 0,
      });
      add(x, A, "a:ext", { cx: emu(item.widthPt), cy: emu(item.heightPt) });
      add(x, A, "a:chOff", { x: 0, y: 0 });
      add(x, A, "a:chExt", { cx: emu(item.widthPt), cy: emu(item.heightPt) });
      for (const c of item.children) await shape(g, c, true);
      return;
    }
    let pr: Element;
    if (item.image) {
      const pic = add(host, PIC, "pic:pic"),
        nv = add(pic, PIC, "pic:nvPicPr");
      add(nv, PIC, "pic:cNvPr", { id: ++childIdentity, name: item.alt });
      add(nv, PIC, "pic:cNvPicPr");
      const bf = add(pic, PIC, "pic:blipFill"),
        blip = add(bf, A, "a:blip");
      blip.setAttributeNS(
        r,
        "r:embed",
        await imageRelation(item.image.dataUrl),
      );
      writeImageEffects(blip, item.image);
      if (item.image.crop)
        add(
          bf,
          A,
          "a:srcRect",
          Object.fromEntries(
            (["left", "top", "right", "bottom"] as const).flatMap((key, i) =>
              item.image?.crop?.[key] !== undefined
                ? [
                    [
                      ["l", "t", "r", "b"][i],
                      Math.round(item.image.crop[key]! * 100000),
                    ],
                  ]
                : [],
            ),
          ),
        );
      add(add(bf, A, "a:stretch"), A, "a:fillRect");
      pr = add(pic, PIC, "pic:spPr");
    } else {
      const s = add(host, WPS, "wps:wsp");
      add(s, WPS, "wps:cNvSpPr", { txBox: item.blocks ? 1 : 0 });
      pr = add(s, WPS, "wps:spPr");
      if (item.blocks) {
        const box = add(add(s, WPS, "wps:txbx"), w, "w:txbxContent");
        for (const b of item.blocks)
          if (b.type === "paragraph") paragraph(box, b, w);
      }
      add(s, WPS, "wps:bodyPr", {
        lIns: emu(4),
        tIns: emu(4),
        rIns: emu(4),
        bIns: emu(4),
      });
    }
    const transform = add(pr, A, "a:xfrm", {
      rot: Math.round((item.shape?.rotation ?? 0) * 60000),
      flipH: item.flipH ? 1 : 0,
      flipV: item.flipV ? 1 : 0,
    });
    add(transform, A, "a:off", {
      x: child ? emu(item.position?.horizontal.offsetPt ?? 0) : 0,
      y: child ? emu(item.position?.vertical.offsetPt ?? 0) : 0,
    });
    add(transform, A, "a:ext", {
      cx: emu(item.widthPt),
      cy: emu(item.heightPt),
    });
    const iconPath = iconPaths[item.shape?.kind ?? ""];
    if (iconPath) {
      const geometry = add(pr, A, "a:custGeom");
      add(add(geometry, A, "a:avLst"), A, "a:gd", {
        name: `docx_${item.shape!.kind}`,
        fmla: "val 0",
      });
      add(geometry, A, "a:gdLst");
      add(geometry, A, "a:ahLst");
      add(geometry, A, "a:cxnLst");
      add(geometry, A, "a:rect", { l: 0, t: 0, r: 100, b: 100 });
      const path = add(add(geometry, A, "a:pathLst"), A, "a:path", {
        w: 100,
        h: 100,
      });
      for (const command of vectorCommands(iconPath)) {
        const node = add(
          path,
          A,
          `a:${{ M: "moveTo", L: "lnTo", C: "cubicBezTo", Q: "quadBezTo", Z: "close" }[command.kind]}`,
        );
        for (let i = 0; i < command.points.length; i += 2)
          add(node, A, "a:pt", {
            x: command.points[i],
            y: command.points[i + 1],
          });
      }
    } else
      add(
        add(pr, A, "a:prstGeom", {
          prst: /^[a-zA-Z][a-zA-Z0-9]*$/.test(item.shape?.kind ?? "rect")
            ? (item.shape?.kind ?? "rect")
            : "rect",
        }),
        A,
        "a:avLst",
      );
    if (!item.image) fill(pr, item.shape?.fill);
    const line = add(pr, A, "a:ln", { w: emu(item.shape?.strokeWidthPt ?? 1) });
    fill(line, item.shape?.stroke);
  };
  await shape(content, d);
  if (d.video) {
    const url = videoUrl(d.video.url),
      blip = descendants(run, "blip")[0];
    if (blip) {
      const ext = add(add(blip, A, "a:extLst"), A, "a:ext", {
        uri: "http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing",
      });
      add(
        ext,
        "http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing",
        "wp15:webVideoPr",
        {
          embeddedHtml: `<iframe src="${url}" width="640" height="360" allowfullscreen="true"></iframe>`,
          w: 640,
          h: 360,
        },
      );
    }
    if (hyperlinkRelation)
      add(prop, A, "a:hlinkClick").setAttributeNS(
        r,
        "r:id",
        await hyperlinkRelation(url),
      );
  }
  return parent.ownerDocument!.importNode(run, true) as Element;
}
/** Geometry changes preserve the source drawing's text, effects, relationships and extensions. */
export function updateDrawing(root: Element, d: DrawingInline) {
  const docPr = descendants(root, "docPr")[0];
  if (docPr) {
    docPr.setAttribute("name", d.alt);
    docPr.setAttribute("descr", d.alt);
  }
  const frame =
    descendants(root, "anchor")[0] ?? descendants(root, "inline")[0];
  if (!frame)
    throw new Error("This object is not an editable DrawingML object.");
  const extent = childElements(frame).find((e) => e.localName === "extent");
  extent?.setAttribute("cx", String(emu(d.widthPt)));
  extent?.setAttribute("cy", String(emu(d.heightPt)));
  const transform = descendants(frame, "xfrm")[0];
  if (transform) {
    transform.setAttribute(
      "rot",
      String(Math.round((d.shape?.rotation ?? 0) * 60000)),
    );
    transform.setAttribute("flipH", d.flipH ? "1" : "0");
    transform.setAttribute("flipV", d.flipV ? "1" : "0");
    const ext = childElements(transform).find((e) => e.localName === "ext");
    ext?.setAttribute("cx", String(emu(d.widthPt)));
    ext?.setAttribute("cy", String(emu(d.heightPt)));
  }
  if (d.shape) {
    const pr = descendants(frame, "spPr")[0];
    const replaceFill = (parent: Element, color?: string) => {
      for (const node of childElements(parent).filter((e) =>
        [
          "solidFill",
          "noFill",
          "gradFill",
          "blipFill",
          "pattFill",
          "grpFill",
        ].includes(e.localName ?? ""),
      ))
        parent.removeChild(node);
      fill(parent, color);
      const node = childElements(parent).at(-1)!,
        before = childElements(parent).find((e) =>
          [
            "ln",
            "effectLst",
            "effectDag",
            "scene3d",
            "sp3d",
            "extLst",
            "prstDash",
            "custDash",
            "round",
            "bevel",
            "miter",
            "headEnd",
            "tailEnd",
          ].includes(e.localName ?? ""),
        );
      if (before && before !== node) parent.insertBefore(node, before);
    };
    if (pr) {
      if (!d.image) replaceFill(pr, d.shape.fill);
      let line = childElements(pr).find((e) => e.localName === "ln");
      if (!line) {
        line = add(pr, A, "a:ln");
        const before = childElements(pr).find((e) =>
          ["effectLst", "effectDag", "scene3d", "sp3d", "extLst"].includes(
            e.localName ?? "",
          ),
        );
        if (before) pr.insertBefore(line, before);
      }
      line.setAttribute("w", String(emu(d.shape.strokeWidthPt ?? 1)));
      replaceFill(line, d.shape.stroke);
    }
  }
  const image = descendants(frame, "blip")[0];
  if (image && d.image) {
    writeImageEffects(image, d.image);
    if (d.image.crop) {
      const fill = image.parentNode as Element;
      let rect = childElements(fill).find((c) => c.localName === "srcRect");
      if (!rect) {
        rect = fill.ownerDocument!.createElementNS(A, "a:srcRect");
        fill.insertBefore(rect, image.nextSibling);
      }
      for (const [key, side] of [
        ["l", "left"],
        ["t", "top"],
        ["r", "right"],
        ["b", "bottom"],
      ] as const)
        rect.setAttribute(
          key,
          String(Math.round((d.image.crop[side] ?? 0) * 100000)),
        );
    }
  }
  if (d.position) {
    frame.setAttribute(
      "relativeHeight",
      String(Math.max(0, d.position.zIndex)),
    );
    frame.setAttribute("behindDoc", d.position.behindText ? "1" : "0");
    for (const [name, axis] of [
      ["positionH", d.position.horizontal],
      ["positionV", d.position.vertical],
    ] as const) {
      const p = childElements(frame).find((e) => e.localName === name);
      if (p) {
        p.setAttribute("relativeFrom", axis.relativeTo);
        for (const child of childElements(p)) p.removeChild(child);
        const n = add(p, WP, axis.alignment ? "wp:align" : "wp:posOffset");
        n.textContent = axis.alignment ?? String(emu(axis.offsetPt ?? 0));
      }
    }
  }
}
export async function reframeDrawing(root: Element, d: DrawingInline) {
  const old = descendants(root, "anchor")[0] ?? descendants(root, "inline")[0];
  if (!old) return;
  const oldPr = childElements(old).find((e) => e.localName === "docPr"),
    graphic = childElements(old).find((e) => e.localName === "graphic");
  const template = await drawingRun(
    root,
    {
      ...d,
      image: undefined,
      blocks: undefined,
      children: undefined,
      chart: undefined,
      equation: undefined,
      model3d: undefined,
      attachment: undefined,
      video: undefined,
      shape: { kind: "rect" },
    },
    Number(oldPr?.getAttribute("id")) || 1,
    async () => "",
  );
  const next =
    descendants(template, "anchor")[0] ?? descendants(template, "inline")[0];
  if (graphic) {
    const stub = childElements(next).find((e) => e.localName === "graphic")!;
    next.replaceChild(graphic, stub);
  }
  if (oldPr)
    next.replaceChild(
      oldPr,
      childElements(next).find((e) => e.localName === "docPr")!,
    );
  old.parentNode!.replaceChild(next, old);
}
