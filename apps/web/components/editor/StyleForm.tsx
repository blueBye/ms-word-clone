import type { DocxStyle, EditCommand } from "@local/docx-ir-engine";
import { dialogData } from "./FontDialogs";
import { StylePreview } from "./StylePreview";
import { useState } from "react";
import { styleFormCommand } from "../../lib/editor";

export function StyleForm({
  style,
  styles,
  rename = false,
  onApply,
}: {
  style: DocxStyle;
  styles: DocxStyle[];
  rename?: boolean;
  onApply: (command: EditCommand) => void;
}) {
  const [preview, setPreview] = useState(style);
  return (
    <form
      onChange={(event) => {
        if (rename) return;
        const d = new FormData(event.currentTarget);
        setPreview({
          ...style,
          name: String(d.get("name")),
          run: {
            ...style.run,
            fontFamily: String(d.get("font")),
            fontSizePt: Number(d.get("size")),
            bold: d.has("bold"),
            italic: d.has("italic"),
            color: String(d.get("color")),
          },
        });
      }}
      onSubmit={(event) => {
        const d = dialogData(event);
        onApply(
          rename
            ? {
                type: "style",
                styleId: style.id,
                name: String(d.get("name")),
              }
            : styleFormCommand(style, d),
        );
      }}
    >
      <label className="field">
        Name
        <input
          name="name"
          defaultValue={style.name}
          required
          maxLength={200}
          autoFocus
        />
      </label>
      {!rename && (
        <>
          <label className="field">
            Style based on
            <select name="basedOn" defaultValue={style.basedOn ?? ""}>
              <option value="">(no style)</option>
              {styles
                .filter((s) => s.id !== style.id)
                .map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.name}
                  </option>
                ))}
            </select>
          </label>
          <h3>Formatting</h3>
          <div className="field-row">
            <label className="field">
              Font
              <input
                name="font"
                defaultValue={style.run.fontFamily ?? "Calibri"}
                required
              />
            </label>
            <label className="field">
              Size
              <input
                type="number"
                name="size"
                min={1}
                max={400}
                step={0.5}
                defaultValue={style.run.fontSizePt ?? 12}
              />
            </label>
            <label className="field">
              Color
              <input
                type="color"
                name="color"
                defaultValue={style.run.color ?? "#000000"}
              />
            </label>
          </div>
          <div className="field-row">
            <label className="check-field">
              <input
                type="checkbox"
                name="bold"
                defaultChecked={style.run.bold}
              />
              Bold
            </label>
            <label className="check-field">
              <input
                type="checkbox"
                name="italic"
                defaultChecked={style.run.italic}
              />
              Italic
            </label>
          </div>
          <div className="field-row">
            <label className="field">
              Alignment
              <select
                name="align"
                defaultValue={style.paragraph.alignment ?? "left"}
              >
                {["start", "end"].includes(style.paragraph.alignment ?? "") && (
                  <option value={style.paragraph.alignment}>
                    {style.paragraph.alignment}
                  </option>
                )}
                {["left", "center", "right", "justify"].map((a) => (
                  <option key={a}>{a}</option>
                ))}
              </select>
            </label>
            <label className="field">
              Before (pt)
              <input
                name="before"
                type="number"
                min={0}
                max={1000}
                defaultValue={style.paragraph.spaceBeforePt ?? 0}
              />
            </label>
            <label className="field">
              After (pt)
              <input
                name="after"
                type="number"
                min={0}
                max={1000}
                defaultValue={style.paragraph.spaceAfterPt ?? 0}
              />
            </label>
          </div>
          <div className="style-form-preview">
            <StylePreview style={preview} />
          </div>
          <label className="check-field">
            <input
              type="checkbox"
              name="gallery"
              defaultChecked={style.gallery}
            />
            Add to the Styles gallery
          </label>
          <p className="pane-description">
            Changes apply to this document and all paragraphs using this style.
            Direct formatting remains an override.
          </p>
        </>
      )}
      <div className="dialog-actions">
        <button className="primary-button">OK</button>
      </div>
    </form>
  );
}
