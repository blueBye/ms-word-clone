import { useState } from "react";
import type {
  DocumentIR,
  Paragraph,
  EditCommand,
  DrawingInline,
} from "@local/docx-ir-engine/core";
import { Dialog } from "./Controls";
import { paragraphText } from "../../lib/editor";
import { DrawTableGrid } from "./DrawTableGrid";
import {commentSelection} from "../../lib/comments";

const shapeKinds = [
  "rect",
  "roundRect",
  "ellipse",
  "triangle",
  "rtTriangle",
  "diamond",
  "hexagon",
  "pentagon",
  "chevron",
  "rightArrow",
  "leftArrow",
  "upArrow",
  "downArrow",
  "star5",
  "plus",
  "heart",
  "line",
];
export function AuthoringDialog({
  mode,
  document,
  selected,
  range,
  all,
  onClose,
  execute,
  onText,
  onNavigate,
  object,
}: {
  mode: string;
  document: DocumentIR;
  selected?: Paragraph;
  range: { start: number; end: number };
  all: Paragraph[];
  onClose: () => void;
  execute: (commands: EditCommand[]) => Promise<boolean>;
  onText: (text: string) => void;
  onNavigate: (id: string, start: number, end: number) => void;
  object?: { id: string; index: number; drawing: DrawingInline };
}) {
  const [kind, setKind] = useState(shapeKinds[0]),
    [error, setError] = useState("");
  const title: Record<string, string> = {
    cover: "Cover Page",
    video: "Online Video",
    dropCap: "Drop Cap",
    chart: "Insert Chart",
    equation: "Equation",
    shapes: "Shapes",
    icons: "Icons",
    textBox: "Text Box",
    wordArt: "WordArt",
    smartArt: "SmartArt",
    bookmark: "Bookmark",
    reference: "Cross-reference",
    comment: "New Comment",
    columnsOptions: "Columns",
    layoutOptions: "Page Layout",
    hyphenation: "Hyphenation",
    lineNumbers: "Line Numbering",
    quickTable: "Quick Tables",
    convertTable: "Convert Text to Table",
    drawTable: "Draw Table",
    signature: "Signature Line",
    dateOptions: "Date and Time",
    symbols: "Symbol",
    quickParts: "Quick Parts",
  };
  const bookmarks = all.flatMap((p) =>
    (p.bookmarks ?? []).map((b) => ({ ...b, id: p.id! })),
  );
  const apply = async (commands: EditCommand[]) => {
    try {
      if (await execute(commands)) onClose();
      else
        setError(
          "The change could not be applied. Check the document message for details.",
        );
    } catch (error) {
      setError(
        error instanceof Error ? error.message : "Could not apply the change.",
      );
    }
  };
  const actions = (
    <div className="dialog-actions">
      <button className="secondary-button" type="button" onClick={onClose}>
        Cancel
      </button>
      <button className="primary-button">
        {[
          "columnsOptions",
          "layoutOptions",
          "hyphenation",
          "lineNumbers",
        ].includes(mode)
          ? "Apply"
          : "Insert"}
      </button>
    </div>
  );
  const drawing = (data: FormData): DrawingInline => ({
    type: "drawing",
    widthPt: Number(data.get("width") ?? 216),
    heightPt: Number(data.get("height") ?? 108),
    alt: String(data.get("text") ?? kind),
    shape: {
      kind: mode === "shapes" ? kind : mode === "icons" ? kind : "rect",
      fill:
        mode === "textBox" || mode === "wordArt"
          ? undefined
          : String(data.get("fill") ?? "#4479AA"),
      stroke: mode === "wordArt" ? undefined : "#294F76",
      strokeWidthPt: 1,
    },
    ...(["textBox", "wordArt", "signature"].includes(mode)
      ? {
          blocks: [
            {
              type: "paragraph",
              properties: { alignment: mode === "textBox" ? "left" : "center" },
              children: [
                {
                  type: "text",
                  text: String(data.get("text") ?? ""),
                  properties: {
                    fontSizePt: mode === "wordArt" ? 32 : 12,
                    color:
                      mode === "wordArt"
                        ? String(data.get("fill") ?? "#4479AA")
                        : "#000000",
                    ...(mode === "wordArt"
                      ? {
                          bold: true,
                          textOutline: { color: "#294F76", widthPt: 0.5 },
                        }
                      : {}),
                  },
                },
              ],
            },
          ],
        }
      : {}),
  });
  return (
    <Dialog title={title[mode] ?? mode} onClose={onClose}>
      <form
        onSubmit={async (event) => {
          event.preventDefault();
          setError("");
          const data = new FormData(event.currentTarget),
            id = selected?.id;
          if (
            ["shapes", "icons", "textBox", "wordArt", "signature"].includes(
              mode,
            )
          ) {
            await apply([
              { type: "insertDrawing", id, drawing: drawing(data) },
            ]);
            return;
          }
          if (mode === "smartArt") {
            const labels = String(data.get("text"))
                .split("\n")
                .map((t) => t.trim())
                .filter(Boolean)
                .slice(0, 12),
              vertical = data.get("layout") === "list",
              cycle = data.get("layout") === "cycle";
            const width = vertical ? 252 : 432,
              height = cycle ? 300 : vertical ? labels.length * 48 : 84,
              cell = vertical ? width : width / Math.max(1, labels.length);
            const children: DrawingInline[] = labels.map((text, i) => ({
              type: "drawing",
              alt: text,
              widthPt: cycle ? 86 : vertical ? width : cell - 10,
              heightPt: vertical ? 38 : 64,
              position: {
                horizontal: {
                  relativeTo: "paragraph",
                  offsetPt: cycle
                    ? width / 2 -
                      43 +
                      145 *
                        Math.cos(
                          (i / labels.length) * Math.PI * 2 - Math.PI / 2,
                        )
                    : vertical
                      ? 0
                      : i * cell,
                },
                vertical: {
                  relativeTo: "paragraph",
                  offsetPt: cycle
                    ? height / 2 -
                      32 +
                      112 *
                        Math.sin(
                          (i / labels.length) * Math.PI * 2 - Math.PI / 2,
                        )
                    : vertical
                      ? i * 48
                      : 0,
                },
                behindText: false,
                zIndex: 2,
                wrap: "none",
              },
              shape: {
                kind: data.get("layout") === "cycle" ? "ellipse" : "roundRect",
                fill: ["#4479AA", "#588C58", "#A66C3F"][i % 3],
              },
              blocks: [
                {
                  type: "paragraph",
                  properties: { alignment: "center" },
                  children: [
                    {
                      type: "text",
                      text,
                      properties: { color: "#FFFFFF", fontSizePt: 12 },
                    },
                  ],
                },
              ],
            }));
            await apply([
              {
                type: "insertDrawing",
                id,
                drawing: {
                  type: "drawing",
                  alt: "Diagram",
                  widthPt: width,
                  heightPt: height,
                  children,
                },
              },
            ]);
            return;
          }
          if (mode === "dropCap" && id) {
            await apply([
              {
                type: "dropCap",
                id,
                lines: Number(data.get("lines")),
                distancePt: Number(data.get("distance")),
                kind: String(data.get("kind")) as "drop" | "margin",
                remove: data.get("kind") === "none",
              },
            ]);
            return;
          }
          if (mode === "video") {
            const canvas = window.document.createElement("canvas");
            canvas.width = 640;
            canvas.height = 360;
            const ctx = canvas.getContext("2d")!;
            ctx.fillStyle = "#29384A";
            ctx.fillRect(0, 0, 640, 360);
            ctx.fillStyle = "white";
            ctx.font = "bold 52px sans-serif";
            ctx.textAlign = "center";
            ctx.fillText("\u25b6", 320, 190);
            const d: DrawingInline = {
              type: "drawing",
              widthPt: 360,
              heightPt: 202.5,
              alt: "Online video",
              video: { url: String(data.get("url")) },
              image: { dataUrl: canvas.toDataURL("image/png") },
            };
            await apply([{ type: "insertDrawing", id, drawing: d }]);
            return;
          }
          if (mode === "cover") {
            await apply([
              {
                type: "cover",
                title: String(data.get("title")),
                subtitle: String(data.get("subtitle")),
                author: String(data.get("author")),
                date: String(data.get("date")),
                design: String(data.get("design")) as
                  | "simple"
                  | "banded"
                  | "sideline",
              },
            ]);
            return;
          }
          if (mode === "quickParts") {
            await apply([
              {
                type: "insertField",
                id,
                field: String(data.get("field")) as
                  | "title"
                  | "author"
                  | "page"
                  | "pages"
                  | "date",
                value: String(data.get("value")),
              },
            ]);
            return;
          }
          if (mode === "chart") {
            const rows = String(data.get("data"))
                .trim()
                .split(/\r?\n/)
                .map((row) => row.split("\t")),
              header = rows.shift()!;
            const chart: NonNullable<DrawingInline["chart"]> = {
              kind: String(data.get("chartKind")) as "bar" | "line" | "pie",
              title: String(data.get("title")),
              categories: rows.map((r) => r[0]),
              series: header.slice(1).map((name, i) => ({
                name,
                values: rows.map((row) => Number(row[i + 1])),
              })),
            };
            const d: DrawingInline = {
              ...(object?.drawing.chart
                ? object.drawing
                : {
                    type: "drawing",
                    widthPt: 360,
                    heightPt: 240,
                    alt: chart.title,
                  }),
              chart,
            };
            await apply(
              object?.drawing.chart
                ? [
                    {
                      type: "drawing",
                      id: object.id,
                      index: object.index,
                      drawing: d,
                    },
                  ]
                : [{ type: "insertDrawing", id, drawing: d }],
            );
            return;
          }
          if (mode === "equation") {
            const d: DrawingInline = {
              ...(object?.drawing.equation
                ? object.drawing
                : {
                    type: "drawing" as const,
                    widthPt: 300,
                    heightPt: 48,
                    alt: "Equation",
                  }),
              equation: { expression: String(data.get("expression")) },
            };
            await apply(
              object?.drawing.equation
                ? [
                    {
                      type: "drawing",
                      id: object.id,
                      index: object.index,
                      drawing: d,
                    },
                  ]
                : [{ type: "insertDrawing", id, drawing: d }],
            );
            return;
          }
          if (mode === "bookmark" && id) {
            await apply([
              {
                type: "bookmark",
                id,
                name: String(data.get("name")),
                ...range,
              },
            ]);
            return;
          }
          if (mode === "reference") {
            const target = bookmarks.find((b) => b.name === data.get("target"));
            if (!target || !id) return;
            const p = all.find((p) => p.id === target.id)!;
            await apply([
              {
                type: "link",
                id,
                ...range,
                href: `#${encodeURIComponent(target.name)}`,
                text:
                  paragraphText(p).slice(target.start, target.end) ||
                  target.name,
              },
            ]);
            return;
          }
          if (mode === "comment" && id) {
            await apply([
              {
                type: "comment",
                id,
                ...commentSelection(selected ? paragraphText(selected) : "",range),
                key: `new-${Date.now()}`,
                text: String(data.get("text")),
                author: String(data.get("author")),
                date: new Date().toISOString(),
              },
            ]);
            return;
          }
          if (mode === "columnsOptions") {
            await apply([
              {
                type: "columns",
                count: Number(data.get("count")),
                gapPt: Number(data.get("gap")),
                separator: data.has("separator"),
              },
              {
                type: "layout",
                format: {
                  direction: String(data.get("direction")) as "ltr" | "rtl",
                },
              },
            ]);
            return;
          }
          if (mode === "layoutOptions") {
            await apply([
              {
                type: "layout",
                format: {
                  mirrorMargins: data.has("mirror"),
                  gutterPt: Number(data.get("gutter")),
                  differentFirstPage: data.has("first"),
                  evenAndOddHeaders: data.has("even"),
                  verticalAlignment: String(data.get("vertical")) as
                    | "top"
                    | "center"
                    | "bottom",
                  direction: String(data.get("direction")) as "ltr" | "rtl",
                  margins: {
                    ...document.page?.margins,
                    header: Number(data.get("header")),
                    footer: Number(data.get("footer")),
                  },
                  ...(data.has("restart")
                    ? { pageNumberStart: Number(data.get("start")) }
                    : {}),
                },
              },
            ]);
            return;
          }
          if (mode === "hyphenation") {
            await apply([
              {
                type: "layout",
                format: { hyphenation: data.has("auto") ? "auto" : "none" },
              },
            ]);
            return;
          }
          if (mode === "lineNumbers") {
            await apply([
              {
                type: "layout",
                format: {
                  lineNumbers: data.has("enabled")
                    ? {
                        mode: String(data.get("mode")) as
                          | "page"
                          | "continuous"
                          | "section",
                        start: Number(data.get("start")),
                        countBy: Number(data.get("count")),
                        distancePt: Number(data.get("distance")),
                      }
                    : undefined,
                },
              },
            ]);
            return;
          }
          if (["drawTable", "quickTable", "convertTable"].includes(mode)) {
            let cells: string[][] | undefined;
            if (mode === "convertTable")
              cells = String(data.get("text"))
                .split(/\r?\n/)
                .map((row) =>
                  row.split(
                    String(data.get("separator")) === "comma" ? "," : "\t",
                  ),
                );
            await apply([
              ...(mode === "convertTable" && id && range.end > range.start
                ? [{ type: "text" as const, id, ...range, text: "" }]
                : []),
              {
                type: "insertTable",
                id,
                rows: cells?.length ?? Number(data.get("rows")),
                columns: cells
                  ? Math.max(...cells.map((row) => row.length))
                  : Number(data.get("columns")),
                cells,
                header: data.has("header"),
                ...(data.get("widths")
                  ? {
                      columnWidthsPt: String(data.get("widths"))
                        .split(",")
                        .map(Number),
                    }
                  : {}),
              },
            ]);
            return;
          }
          if (mode === "dateOptions") {
            onText(
              new Intl.DateTimeFormat(String(data.get("language")), {
                dateStyle: String(data.get("format")) as
                  | "full"
                  | "long"
                  | "medium"
                  | "short",
                calendar: String(data.get("calendar")),
              }).format(new Date()),
            );
            onClose();
            return;
          }
          if (mode === "symbols") {
            const code = parseInt(String(data.get("code")), 16);
            if (
              !Number.isInteger(code) ||
              code < 32 ||
              code > 0x10ffff ||
              (code >= 0xd800 && code <= 0xdfff)
            ) {
              setError("Enter a valid Unicode character code.");
              return;
            }
            onText(String.fromCodePoint(code));
            onClose();
            return;
          }
        }}
      >
        {["shapes", "icons"].includes(mode) && (
          <div className="shape-picker">
            {shapeKinds.map((name) => (
              <button
                type="button"
                key={name}
                className={kind === name ? "selected" : ""}
                onClick={() => setKind(name)}
              >
                {name.replace(/([A-Z])/g, " $1")}
              </button>
            ))}
          </div>
        )}
        {["shapes", "icons", "textBox", "wordArt", "signature"].includes(
          mode,
        ) && (
          <>
            <div className="field-row">
              <label className="field">
                Width (pt)
                <input
                  name="width"
                  type="number"
                  min="12"
                  max="1440"
                  defaultValue={216}
                />
              </label>
              <label className="field">
                Height (pt)
                <input
                  name="height"
                  type="number"
                  min="12"
                  max="1440"
                  defaultValue={108}
                />
              </label>
              <label className="field">
                Color
                <input name="fill" type="color" defaultValue="#4479AA" />
              </label>
            </div>
            {["textBox", "wordArt", "signature"].includes(mode) && (
              <label className="field">
                {mode === "signature" ? "Signer name and title" : "Text"}
                <textarea
                  name="text"
                  defaultValue={
                    mode === "signature"
                      ? "______________________\nName\nTitle"
                      : "Your text"
                  }
                  required
                />
              </label>
            )}
          </>
        )}
        {mode === "smartArt" && (
          <>
            <label className="field">
              Layout
              <select name="layout">
                <option value="process">Process</option>
                <option value="list">List</option>
                <option value="cycle">Cycle</option>
              </select>
            </label>
            <label className="field">
              One item per line
              <textarea
                name="text"
                defaultValue={"First step\nSecond step\nThird step"}
                required
              />
            </label>
          </>
        )}
        {mode === "bookmark" && (
          <>
            <label className="field">
              Bookmark name
              <input name="name" required maxLength={40} autoFocus />
            </label>
            <div className="bookmark-list">
              {bookmarks.map((b) => (
                <div key={b.name}>
                  <button
                    type="button"
                    onClick={() => onNavigate(b.id, b.start, b.end)}
                  >
                    {b.name}
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      void apply([{ type: "bookmark", ...b, remove: true }])
                    }
                  >
                    Delete
                  </button>
                </div>
              ))}
            </div>
          </>
        )}
        {mode === "reference" && (
          <label className="field">
            Reference target
            <select name="target" required>
              {bookmarks.map((b) => (
                <option key={b.name}>{b.name}</option>
              ))}
            </select>
          </label>
        )}
        {mode === "comment" && (
          <>
            <label className="field">
              Author
              <input name="author" defaultValue="Author" required />
            </label>
            <label className="field">
              Comment
              <textarea name="text" required autoFocus />
            </label>
          </>
        )}
        {mode === "columnsOptions" && (
          <>
            <div className="field-row">
              <label className="field">
                Number of columns
                <input
                  type="number"
                  name="count"
                  min="1"
                  max="4"
                  defaultValue={document.page?.columns?.count ?? 1}
                />
              </label>
              <label className="field">
                Spacing (pt)
                <input
                  type="number"
                  name="gap"
                  min="0"
                  max="144"
                  defaultValue={document.page?.columns?.gapPt ?? 36}
                />
              </label>
            </div>
            <label className="check-field">
              <input
                type="checkbox"
                name="separator"
                defaultChecked={document.page?.columns?.separator}
              />
              Line between
            </label>
            <Direction direction={document.page?.direction} />
          </>
        )}
        {mode === "layoutOptions" && (
          <>
            <label className="check-field">
              <input
                type="checkbox"
                name="mirror"
                defaultChecked={document.page?.mirrorMargins}
              />
              Mirror margins
            </label>
            <label className="field">
              Gutter (pt)
              <input
                name="gutter"
                type="number"
                min="0"
                max="720"
                defaultValue={document.page?.gutterPt ?? 0}
              />
            </label>
            <label className="check-field">
              <input
                type="checkbox"
                name="first"
                defaultChecked={document.page?.differentFirstPage}
              />
              Different first page
            </label>
            <label className="check-field">
              <input
                type="checkbox"
                name="even"
                defaultChecked={document.page?.evenAndOddHeaders}
              />
              Different odd and even pages
            </label>
            <label className="field">
              Vertical alignment
              <select
                name="vertical"
                defaultValue={document.page?.verticalAlignment ?? "top"}
              >
                {["top", "center", "bottom"].map((value) => (
                  <option key={value}>{value}</option>
                ))}
              </select>
            </label>
            <Direction direction={document.page?.direction} />
            <div className="field-row">
              {(["header", "footer"] as const).map((kind) => (
                <label className="field" key={kind}>
                  {kind === "header" ? "Header" : "Footer"} distance from edge
                  (pt)
                  <input
                    type="number"
                    name={kind}
                    min="0"
                    max="720"
                    defaultValue={document.page?.margins?.[kind] ?? 36}
                  />
                </label>
              ))}
            </div>
            <label className="check-field">
              <input
                type="checkbox"
                name="restart"
                defaultChecked={document.page?.pageNumberStart !== undefined}
              />
              Restart page numbering
            </label>
            <label className="field">
              Start page numbering at
              <input
                name="start"
                type="number"
                min="1"
                max="32767"
                defaultValue={document.page?.pageNumberStart ?? 1}
              />
            </label>
          </>
        )}
        {mode === "hyphenation" && (
          <>
            <label className="check-field">
              <input
                type="checkbox"
                name="auto"
                defaultChecked={document.page?.hyphenation === "auto"}
              />
              Automatically hyphenate
            </label>
            <button
              type="button"
              className="secondary-button"
              onClick={() => {
                onText("\u00ad");
                onClose();
              }}
            >
              Insert optional hyphen at selection
            </button>
          </>
        )}
        {mode === "lineNumbers" && (
          <>
            <label className="check-field">
              <input
                type="checkbox"
                name="enabled"
                defaultChecked={!!document.page?.lineNumbers}
              />
              Add line numbering
            </label>
            <label className="field">
              Numbering
              <select
                name="mode"
                defaultValue={document.page?.lineNumbers?.mode ?? "continuous"}
              >
                <option value="continuous">Continuous</option>
                <option value="page">Restart each page</option>
                <option value="section">Restart each section</option>
              </select>
            </label>
            <div className="field-row">
              {[
                ["Start at", "start", document.page?.lineNumbers?.start ?? 1],
                ["Count by", "count", document.page?.lineNumbers?.countBy ?? 1],
                [
                  "Distance from text (pt)",
                  "distance",
                  document.page?.lineNumbers?.distancePt ?? 18,
                ],
              ].map(([label, name, value]) => (
                <label className="field" key={name}>
                  {label}
                  <input
                    type="number"
                    name={String(name)}
                    min={name === "distance" ? 0 : 1}
                    defaultValue={value}
                  />
                </label>
              ))}
            </div>
          </>
        )}
        {mode === "drawTable" && <DrawTableGrid />}
        {["quickTable", "convertTable"].includes(mode) && (
          <>
            {mode === "convertTable" ? (
              <>
                <label className="field">
                  Text
                  <textarea
                    name="text"
                    defaultValue={
                      selected
                        ? paragraphText(selected).slice(range.start, range.end)
                        : ""
                    }
                    required
                  />
                </label>
                <label className="field">
                  Separate text at
                  <select name="separator">
                    <option value="tabs">Tabs</option>
                    <option value="comma">Commas</option>
                  </select>
                </label>
              </>
            ) : (
              <div className="field-row">
                <label className="field">
                  Rows
                  <input
                    type="number"
                    name="rows"
                    min="1"
                    max="30"
                    defaultValue="4"
                  />
                </label>
                <label className="field">
                  Columns
                  <input
                    type="number"
                    name="columns"
                    min="1"
                    max="12"
                    defaultValue="3"
                  />
                </label>
              </div>
            )}
            <label className="check-field">
              <input
                type="checkbox"
                name="header"
                defaultChecked={mode === "quickTable"}
              />
              Header row
            </label>
          </>
        )}
        {mode === "dateOptions" && (
          <>
            <label className="field">
              Format
              <select name="format">
                {["full", "long", "medium", "short"].map((value) => (
                  <option key={value}>{value}</option>
                ))}
              </select>
            </label>
            <label className="field">
              Language
              <select
                name="language"
                defaultValue={document.language ?? "en-US"}
              >
                <option value="en-US">English</option>
                <option value="fa-IR">Persian</option>
                <option value="ar">Arabic</option>
                <option value="he">Hebrew</option>
              </select>
            </label>
            <label className="field">
              Calendar
              <select name="calendar">
                <option value="gregory">Gregorian</option>
                <option value="persian">Persian</option>
                <option value="islamic">Islamic</option>
              </select>
            </label>
          </>
        )}
        {mode === "symbols" && (
          <label className="field">
            Unicode character code (hexadecimal)
            <input name="code" defaultValue="03A9" required />
          </label>
        )}
        {mode === "dropCap" && (
          <>
            <label className="field">
              Position
              <select name="kind">
                <option value="drop">Dropped</option>
                <option value="margin">In margin</option>
                <option value="none">None</option>
              </select>
            </label>
            <div className="field-row">
              <label className="field">
                Lines to drop
                <input
                  name="lines"
                  type="number"
                  min="2"
                  max="10"
                  defaultValue="3"
                />
              </label>
              <label className="field">
                Distance from text (pt)
                <input
                  name="distance"
                  type="number"
                  min="0"
                  max="72"
                  defaultValue="4"
                />
              </label>
            </div>
          </>
        )}
        {mode === "video" && (
          <label className="field">
            YouTube or Vimeo URL
            <input
              type="url"
              name="url"
              required
              placeholder="https://www.youtube.com/watch?v=..."
              autoFocus
            />
          </label>
        )}
        {mode === "cover" && (
          <>
            <label className="field">
              Design
              <select name="design">
                <option value="simple">Simple</option>
                <option value="banded">Banded</option>
                <option value="sideline">Sideline</option>
              </select>
            </label>
            {[
              [
                "Title",
                "title",
                document.properties?.title ?? "Document title",
              ],
              ["Subtitle", "subtitle", "Subtitle"],
              ["Author", "author", document.properties?.author ?? ""],
              ["Date", "date", new Date().toLocaleDateString()],
            ].map(([label, name, value]) => (
              <label className="field" key={name}>
                {label}
                <input name={name} defaultValue={value} />
              </label>
            ))}
          </>
        )}
        {mode === "quickParts" && (
          <>
            <label className="field">
              Field
              <select name="field">
                <option value="title">Document Title</option>
                <option value="author">Author</option>
                <option value="date">Date</option>
                <option value="page">Page Number</option>
                <option value="pages">Number of Pages</option>
              </select>
            </label>
            <label className="field">
              Value
              <input
                name="value"
                defaultValue={document.properties?.title ?? ""}
              />
            </label>
          </>
        )}
        {mode === "chart" && (
          <>
            <label className="field">
              Chart type
              <select
                name="chartKind"
                defaultValue={object?.drawing.chart?.kind ?? "bar"}
              >
                <option value="bar">Column</option>
                <option value="line">Line</option>
                <option value="pie">Pie</option>
              </select>
            </label>
            <label className="field">
              Title
              <input
                name="title"
                defaultValue={object?.drawing.chart?.title ?? "Chart"}
              />
            </label>
            <label className="field">
              Data (tab-separated columns; first row contains series names)
              <textarea
                name="data"
                rows={8}
                required
                defaultValue={
                  object?.drawing.chart
                    ? [
                        [
                          "Category",
                          ...object.drawing.chart.series.map((s) => s.name),
                        ],
                        ...object.drawing.chart.categories.map((c, i) => [
                          c,
                          ...object.drawing.chart!.series.map(
                            (s) => s.values[i],
                          ),
                        ]),
                      ]
                        .map((row) => row.join("\t"))
                        .join("\n")
                    : "Category\tSeries 1\tSeries 2\nFirst\t4\t3\nSecond\t7\t5\nThird\t2\t6"
                }
              />
            </label>
          </>
        )}
        {mode === "equation" && (
          <>
            <label className="field">
              Equation
              <input
                name="expression"
                required
                defaultValue={
                  object?.drawing.equation?.expression ?? "x^2+y^2=z^2"
                }
                autoFocus
              />
            </label>
            <p className="pane-description">
              Use ^ for powers, _ for subscripts, braces for grouping, \frac
              numerator/denominator groups, and \sqrt for roots.
            </p>
          </>
        )}
        {error && <p role="alert">{error}</p>}
        {actions}
      </form>
    </Dialog>
  );
}
function Direction({ direction }: { direction?: "ltr" | "rtl" }) {
  return (
    <label className="field">
      Section direction
      <select name="direction" defaultValue={direction ?? "ltr"}>
        <option value="ltr">Left to right</option>
        <option value="rtl">Right to left</option>
      </select>
    </label>
  );
}
