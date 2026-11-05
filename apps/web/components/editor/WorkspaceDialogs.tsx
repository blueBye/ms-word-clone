import type { FormEvent } from "react";
import type {
  DocumentIR,
  Paragraph,
  EditCommand,
  ParagraphFormat,
} from "@local/docx-ir-engine/core";
import { Dialog } from "./Controls";
export type EditorModal =
  | "page"
  | "paragraph"
  | "table"
  | "rename"
  | "link"
  | null;
interface Props {
  modal: EditorModal;
  onClose: () => void;
  document: DocumentIR | null;
  selected?: Paragraph;
  filename: string;
  onFilename: (value: string) => void;
  busy: boolean;
  canPage: boolean;
  execute: (command: EditCommand) => Promise<boolean>;
}
function data(event: FormEvent<HTMLFormElement>) {
  event.preventDefault();
  return new FormData(event.currentTarget);
}
export function WorkspaceDialogs(p: Props) {
  const actions = (label = "Apply", disabled = false) => (
    <div className="dialog-actions">
      <button type="button" className="secondary-button" onClick={p.onClose}>
        Cancel
      </button>
      <button className="primary-button" disabled={p.busy || disabled}>
        {label}
      </button>
    </div>
  );
  if (p.modal === "rename")
    return (
      <Dialog title="Document name" onClose={p.onClose}>
        <form
          onSubmit={(e) => {
            const d = data(e),
              name = String(d.get("name") ?? "")
                .trim()
                .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_");
            if (name) {
              p.onFilename(
                name.toLowerCase().endsWith(".docx") ? name : `${name}.docx`,
              );
              p.onClose();
            }
          }}
        >
          <label className="field">
            Download name
            <input
              name="name"
              defaultValue={p.filename}
              required
              autoFocus
              maxLength={180}
            />
          </label>
          {actions("Rename")}
        </form>
      </Dialog>
    );
  if (p.modal === "table")
    return (
      <Dialog title="Insert table" onClose={p.onClose}>
        <form
          onSubmit={async (e) => {
            const d = data(e);
            if (
              await p.execute({
                type: "insertTable",
                id: p.selected?.id,
                rows: Number(d.get("rows")),
                columns: Number(d.get("columns")),
              })
            )
              p.onClose();
          }}
        >
          <p className="pane-description">
            Insert after a body paragraph outside existing tables.
          </p>
          <div className="field-row">
            <label className="field">
              Columns
              <input
                name="columns"
                type="number"
                min="1"
                max="12"
                defaultValue="3"
                required
              />
            </label>
            <label className="field">
              Rows
              <input
                name="rows"
                type="number"
                min="1"
                max="30"
                defaultValue="3"
                required
              />
            </label>
          </div>
          {actions("Insert table")}
        </form>
      </Dialog>
    );
  if (p.modal === "page")
    return (
      <Dialog title="Page setup" onClose={p.onClose}>
        <form
          onSubmit={async (e) => {
            const d = data(e);
            if (
              await p.execute({
                type: "page",
                widthPt: Number(d.get("width")),
                heightPt: Number(d.get("height")),
                margins: {
                  top: Number(d.get("top")),
                  bottom: Number(d.get("bottom")),
                  left: Number(d.get("left")),
                  right: Number(d.get("right")),
                },
              })
            )
              p.onClose();
          }}
        >
          <p className="pane-description">
            Measurements in points. 72 pt = 1 inch. Applies to the document’s
            single section.
          </p>
          <div className="field-row">
            <label className="field">
              Width
              <input
                name="width"
                type="number"
                min="72"
                max="31680"
                step="0.1"
                defaultValue={p.document?.page?.widthPt ?? 595.3}
                required
              />
            </label>
            <label className="field">
              Height
              <input
                name="height"
                type="number"
                min="72"
                max="31680"
                step="0.1"
                defaultValue={p.document?.page?.heightPt ?? 841.9}
                required
              />
            </label>
          </div>
          <h3>Margins</h3>
          <div className="field-grid">
            {(["top", "bottom", "left", "right"] as const).map((side) => (
              <label className="field" key={side}>
                {side[0].toUpperCase() + side.slice(1)}
                <input
                  name={side}
                  type="number"
                  min="0"
                  step="0.1"
                  defaultValue={p.document?.page?.margins?.[side] ?? 72}
                  required
                />
              </label>
            ))}
          </div>
          {!p.canPage && (
            <p className="pane-callout">
              Multi-section page setup is retained but not editable yet.
            </p>
          )}
          {actions("Apply", !p.canPage)}
        </form>
      </Dialog>
    );
  if (p.modal === "paragraph" && p.selected) {
    const selected = p.selected;
    return (
      <Dialog title="Paragraph" onClose={p.onClose}>
        <form
          onSubmit={async (e) => {
            const d = data(e);
            if (
              await p.execute({
                type: "paragraph",
                id: selected.id!,
                format: {
                  direction: d.get("direction") as "rtl" | "ltr",
                  alignment: d.get("alignment") as ParagraphFormat["alignment"],
                  spaceBeforePt: Number(d.get("before")),
                  spaceAfterPt: Number(d.get("after")),
                  indentStartPt: Number(d.get("start")),
                  indentEndPt: Number(d.get("end")),
                  lineSpacing: Number(d.get("line")),
                  keepNext: d.has("keepNext"),
                  keepLines: d.has("keepLines"),
                  pageBreakBefore: d.has("pageBreak"),
                },
              })
            )
              p.onClose();
          }}
        >
          <div className="field-row">
            <label className="field">
              Direction
              <select
                name="direction"
                defaultValue={selected.properties.direction ?? "ltr"}
              >
                <option value="ltr">Left to right</option>
                <option value="rtl">Right to left</option>
              </select>
            </label>
            <label className="field">
              Alignment
              <select
                name="alignment"
                defaultValue={
                  ["left", "right", "center", "justify"].includes(
                    selected.properties.alignment ?? "",
                  )
                    ? selected.properties.alignment
                    : "left"
                }
              >
                {["left", "center", "right", "justify"].map((v) => (
                  <option key={v}>{v}</option>
                ))}
              </select>
            </label>
          </div>
          <h3>Indentation (pt)</h3>
          <div className="field-row">
            {[
              [
                "Before text",
                "start",
                selected.properties.indentStartPt ??
                  selected.properties.indentLeftPt ??
                  0,
              ],
              [
                "After text",
                "end",
                selected.properties.indentEndPt ??
                  selected.properties.indentRightPt ??
                  0,
              ],
            ].map(([label, name, value]) => (
              <label className="field" key={name}>
                {label}
                <input
                  name={String(name)}
                  type="number"
                  min="-720"
                  max="1440"
                  defaultValue={value}
                  required
                />
              </label>
            ))}
          </div>
          <h3>Spacing</h3>
          <div className="field-row">
            {[
              ["Before (pt)", "before", selected.properties.spaceBeforePt ?? 0],
              ["After (pt)", "after", selected.properties.spaceAfterPt ?? 0],
            ].map(([label, name, value]) => (
              <label className="field" key={name}>
                {label}
                <input
                  name={String(name)}
                  type="number"
                  min="0"
                  max="1000"
                  defaultValue={value}
                  required
                />
              </label>
            ))}
          </div>
          <label className="field">
            Line spacing (multiple)
            <input
              name="line"
              type="number"
              min="0.5"
              max="10"
              step="0.05"
              defaultValue={
                selected.properties.lineSpacing?.rule === "auto"
                  ? selected.properties.lineSpacing.value
                  : 1
              }
              required
            />
          </label>
          <h3>Pagination</h3>
          <label className="check-field">
            <input
              name="keepNext"
              type="checkbox"
              defaultChecked={selected.properties.keepNext}
            />
            Keep with next
          </label>
          <label className="check-field">
            <input
              name="keepLines"
              type="checkbox"
              defaultChecked={selected.properties.keepLines}
            />
            Keep lines together
          </label>
          <label className="check-field">
            <input
              name="pageBreak"
              type="checkbox"
              defaultChecked={selected.properties.pageBreakBefore}
            />
            Page break before
          </label>
          {actions()}
        </form>
      </Dialog>
    );
  }
  return null;
}
