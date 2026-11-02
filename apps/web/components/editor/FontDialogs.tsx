import { useState, type FormEvent } from "react";
import type { RunFormat } from "@local/docx-ir-engine";
export function dialogData(event: FormEvent<HTMLFormElement>) {
  event.preventDefault();
  return new FormData(event.currentTarget);
}
export function FontChooser({
  families,
  value,
  onPick,
  onManage,
}: {
  families: string[];
  value: string;
  onPick: (value: string) => void;
  onManage: () => void;
}) {
  const [search, setSearch] = useState("");
  const names = [
    ...new Set([
      value,
      ...families,
      "Calibri",
      "Aptos",
      "Arial",
      "Tahoma",
      "Times New Roman",
      "Cambria",
      "Georgia",
      "Verdana",
      "B Nazanin",
      "B Zar",
      "Pofak",
    ]),
  ];
  return (
    <>
      <label className="field">
        Search fonts
        <input
          autoFocus
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Font family"
        />
      </label>
      <div className="font-choice-list">
        {names
          .filter((name) =>
            name.toLocaleLowerCase().includes(search.toLocaleLowerCase()),
          )
          .map((name) => (
            <button
              key={name}
              style={{ fontFamily: `${JSON.stringify(name)},sans-serif` }}
              onClick={() => onPick(name)}
            >
              {name}
            </button>
          ))}
      </div>
      <button className="secondary-button" onClick={onManage}>
        Load a custom font…
      </button>
    </>
  );
}
export function FontForm({
  format,
  onApply,
}: {
  format: RunFormat;
  onApply: (format: RunFormat) => void;
}) {
  return (
    <form
      onSubmit={(event) => {
        const d = dialogData(event);
        onApply({
          fontFamily: String(d.get("font")),
          complexScriptFont: String(d.get("csFont")),
          fontSizePt: Number(d.get("size")),
          complexScriptSizePt: Number(d.get("csSize")),
          bold: d.has("bold"),
          italic: d.has("italic"),
          strike: d.has("strike"),
          smallCaps: d.has("smallCaps"),
          caps: d.has("caps"),
          hidden: d.has("hidden"),
          letterSpacingPt: Number(d.get("spacing")),
          positionPt: Number(d.get("position")),
        });
      }}
    >
      <h3>Latin text</h3>
      <div className="field-row">
        <label className="field">
          Font
          <input
            name="font"
            defaultValue={format.fontFamily ?? "Calibri"}
            required
            autoFocus
          />
        </label>
        <label className="field narrow-field">
          Size
          <input
            name="size"
            type="number"
            min="1"
            max="400"
            step="0.5"
            defaultValue={format.fontSizePt ?? 12}
            required
          />
        </label>
      </div>
      <h3>Complex scripts</h3>
      <div className="field-row">
        <label className="field">
          Font
          <input
            name="csFont"
            defaultValue={
              format.complexScriptFont ?? format.fontFamily ?? "Tahoma"
            }
            required
          />
        </label>
        <label className="field narrow-field">
          Size
          <input
            name="csSize"
            type="number"
            min="1"
            max="400"
            step="0.5"
            defaultValue={format.complexScriptSizePt ?? format.fontSizePt ?? 12}
            required
          />
        </label>
      </div>
      <div className="home-effects-grid">
        {(
          ["bold", "italic", "strike", "smallCaps", "caps", "hidden"] as const
        ).map((key) => (
          <label key={key} className="check-field">
            <input name={key} type="checkbox" defaultChecked={!!format[key]} />
            {
              {
                bold: "Bold",
                italic: "Italic",
                strike: "Strikethrough",
                smallCaps: "Small caps",
                caps: "All caps",
                hidden: "Hidden",
              }[key]
            }
          </label>
        ))}
      </div>
      <div className="field-row">
        <label className="field">
          Character spacing (pt)
          <input
            name="spacing"
            type="number"
            min="-20"
            max="100"
            step="0.1"
            defaultValue={format.letterSpacingPt ?? 0}
          />
        </label>
        <label className="field">
          Raised / lowered (pt)
          <input
            name="position"
            type="number"
            min="-100"
            max="100"
            step="0.5"
            defaultValue={format.positionPt ?? 0}
          />
        </label>
      </div>
      <div className="font-preview" dir="auto">
        AaBbCc · فارسی العربية
      </div>
      <div className="dialog-actions">
        <button className="primary-button">Apply</button>
      </div>
    </form>
  );
}
export function Palette({ onPick }: { onPick: (color: string) => void }) {
  const theme = [
    "#FFFFFF",
    "#000000",
    "#E7E6E6",
    "#44546A",
    "#5B9BD5",
    "#ED7D31",
    "#A5A5A5",
    "#FFC000",
    "#4472C4",
    "#70AD47",
  ];
  const tint = (hex: string, ratio: number) =>
    "#" +
    [1, 3, 5]
      .map((i) =>
        Math.round(
          parseInt(hex.slice(i, i + 2), 16) * (1 - ratio) + 255 * ratio,
        )
          .toString(16)
          .padStart(2, "0"),
      )
      .join("");
  return (
    <>
      <h3>Theme colors</h3>
      <div className="home-palette">
        {theme.map((color) => (
          <button
            key={color}
            aria-label={`Color ${color}`}
            title={color}
            style={{ background: color }}
            onClick={() => onPick(color)}
          />
        ))}
        {[0.8, 0.6, 0.4, 0.2, 0].flatMap((ratio, row) =>
          theme.map((color, col) => (
            <button
              key={`${row}-${col}`}
              aria-label={`Color ${tint(color, ratio)}`}
              title={tint(color, ratio)}
              style={{ background: tint(color, ratio) }}
              onClick={() => onPick(tint(color, ratio))}
            />
          )),
        )}
      </div>
      <h3>Standard colors</h3>
      <div className="home-standard-colors">
        {[
          "#C00000",
          "#FF0000",
          "#FFC000",
          "#FFFF00",
          "#92D050",
          "#00B050",
          "#00B0F0",
          "#0070C0",
          "#002060",
          "#7030A0",
        ].map((color) => (
          <button
            key={color}
            aria-label={`Color ${color}`}
            title={color}
            style={{ background: color }}
            onClick={() => onPick(color)}
          />
        ))}
      </div>
      <form
        onSubmit={(event) => {
          const d = dialogData(event);
          onPick(String(d.get("color")));
        }}
        className="home-custom-color"
      >
        <label>
          Custom color
          <input name="color" type="color" defaultValue="#245EAA" />
        </label>
        <button className="primary-button">Apply color</button>
      </form>
    </>
  );
}
