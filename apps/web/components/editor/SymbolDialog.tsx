import { useState } from "react";
import { Dialog } from "./Controls";
const subsets: Record<string, [number, number]> = {
  "Latin-1 Supplement": [0x00a1, 0x00ff],
  "Greek and Coptic": [0x391, 0x3ce],
  Arabic: [0x621, 0x64a],
  "General Punctuation": [0x2010, 0x205e],
  "Currency Symbols": [0x20a0, 0x20bf],
  Arrows: [0x2190, 0x21ff],
  "Mathematical Operators": [0x2200, 0x22ff],
  "Geometric Shapes": [0x25a0, 0x25ff],
};
export function SymbolDialog({
  families,
  onInsert,
  onClose,
}: {
  families: string[];
  onInsert: (text: string, font: string) => Promise<boolean>;
  onClose: () => void;
}) {
  const [font, setFont] = useState("Times New Roman"),
    [subset, setSubset] = useState("Mathematical Operators"),
    [code, setCode] = useState("221e"),
    [busy, setBusy] = useState(false),
    [error, setError] = useState("");
  const [start, end] = subsets[subset],
    n = parseInt(code, 16),
    valid =
      Number.isInteger(n) &&
      n >= 32 &&
      n <= 0x10ffff &&
      !(n >= 0xd800 && n <= 0xdfff);
  const insert = async () => {
    if (!valid) return;
    setBusy(true);
    try {
      if (!(await onInsert(String.fromCodePoint(n), font)))
        setError("Could not insert this symbol.");
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog title="Symbol" onClose={onClose}>
      <div className="symbol-dialog">
        <div className="field-row">
          <label className="field">
            Font
            <select value={font} onChange={(e) => setFont(e.target.value)}>
              {[
                ...new Set([
                  font,
                  "Times New Roman",
                  "Cambria Math",
                  "Arial",
                  ...families,
                ]),
              ].map((name) => (
                <option key={name}>{name}</option>
              ))}
            </select>
          </label>
          <label className="field">
            Subset
            <select value={subset} onChange={(e) => setSubset(e.target.value)}>
              {Object.keys(subsets).map((s) => (
                <option key={s}>{s}</option>
              ))}
            </select>
          </label>
        </div>
        <div className="symbol-character-grid" style={{ fontFamily: font }}>
          {Array.from({ length: end - start + 1 }, (_, i) => start + i).map(
            (c) => (
              <button
                type="button"
                key={c}
                title={`U+${c.toString(16).toUpperCase()}`}
                aria-pressed={c === n}
                onClick={() => setCode(c.toString(16))}
                onDoubleClick={() => {
                  setCode(c.toString(16));
                  void onInsert(String.fromCodePoint(c), font);
                }}
              >
                {String.fromCodePoint(c)}
              </button>
            ),
          )}
        </div>
        <div className="field-row">
          <span className="symbol-preview" style={{ fontFamily: font }}>
            {valid ? String.fromCodePoint(n) : ""}
          </span>
          <label className="field">
            Character code (Unicode hex)
            <input
              value={code}
              onChange={(e) => setCode(e.target.value)}
              maxLength={6}
            />
          </label>
        </div>
        {error && <p role="alert">{error}</p>}
        <div className="dialog-actions">
          <button type="button" className="secondary-button" onClick={onClose}>
            Close
          </button>
          <button
            type="button"
            className="primary-button"
            disabled={!valid || busy}
            onClick={insert}
          >
            Insert
          </button>
        </div>
      </div>
    </Dialog>
  );
}
