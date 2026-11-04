import type { EditCommand, Paragraph } from "@local/docx-ir-engine/core";
import { MenuBranch } from "./HomeMenu";
import { dialogData } from "./FontDialogs";

type ListCommand = Extract<EditCommand, { type: "list" }>;
const numbering: [string[], ListCommand["kind"], string][] = [
  [["1.", "2.", "3."], "decimal", "%1."],
  [["1)", "2)", "3)"], "decimal", "%1)"],
  [["I.", "II.", "III."], "upperRoman", "%1."],
  [["A.", "B.", "C."], "upperLetter", "%1."],
  [["a)", "b)", "c)"], "lowerLetter", "%1)"],
  [["a.", "b.", "c."], "lowerLetter", "%1."],
  [["i.", "ii.", "iii."], "lowerRoman", "%1."],
];
export function ListGallery({
  mode,
  selected,
  apply,
  onDialog,
}: {
  mode: string;
  selected?: Paragraph;
  apply: (command: EditCommand) => void;
  onDialog: (name: string) => void;
}) {
  const id = selected?.id ?? "",
    level = selected?.properties.numbering?.level ?? 0;
  const current = (selected?.properties.list?.format ??
    "decimal") as ListCommand["kind"];
  return (
    <>
      <h3 className="menu-heading">
        {mode === "bullets"
          ? "Bullet"
          : mode === "numbering"
            ? "Numbering"
            : "List"}{" "}
        Library
      </h3>
      <div className="word-list-gallery">
        <button
          className={!selected?.properties.list ? "selected" : ""}
          onClick={() => apply({ type: "list", id, kind: "none" })}
        >
          None
        </button>
        {mode === "bullets"
          ? ["•", "○", "▪", "❖", "➢", "✓", "–"].map((bullet) => (
              <button
                key={bullet}
                title={`Bullet ${bullet}`}
                onClick={() =>
                  apply({ type: "list", id, kind: "bullet", bullet })
                }
              >
                {[0, 1, 2].map((i) => (
                  <span key={i}>
                    <b>{bullet}</b>
                    <i />
                  </span>
                ))}
              </button>
            ))
          : mode === "numbering"
            ? numbering.map(([markers, kind, pattern]) => (
                <button
                  key={kind + pattern}
                  title={`${markers.join(" ")} numbering`}
                  onClick={() => apply({ type: "list", id, kind, pattern })}
                >
                  {markers.map((marker) => (
                    <span key={marker}>
                      <b>{marker}</b>
                      <i />
                    </span>
                  ))}
                </button>
              ))
            : (
                [
                  ["decimal", ["1.", "1.1.", "1.1.1."]],
                  ["mixed", ["1.", "a.", "i."]],
                  ["bullet", ["•", "○", "▪"]],
                ] as const
              ).map(([multilevel, markers]) => (
                <button
                  key={multilevel}
                  title={`${multilevel} multilevel list`}
                  onClick={() =>
                    apply({
                      type: "list",
                      id,
                      kind: multilevel === "bullet" ? "bullet" : "decimal",
                      multilevel,
                    })
                  }
                >
                  {markers.map((marker, i) => (
                    <span key={i} style={{ paddingLeft: i * 10 }}>
                      <b>{marker}</b>
                      <i />
                    </span>
                  ))}
                </button>
              ))}
      </div>
      <hr />
      <MenuBranch
        label="Change List Level"
        disabled={!selected?.properties.list}
      >
        {Array.from({ length: 9 }, (_, i) => (
          <button
            key={i}
            className="home-choice"
            onClick={() =>
              apply({
                type: "list",
                id,
                kind: current,
                level: i,
                preserveDefinition: true,
              })
            }
          >
            <span className="menu-check">{level === i ? "✓" : ""}</span>Level{" "}
            {i + 1}
          </button>
        ))}
      </MenuBranch>
      <button
        className="home-choice"
        onClick={() =>
          onDialog(
            mode === "bullets"
              ? "defineBullet"
              : mode === "multilevel"
                ? "defineMultilevel"
                : "defineNumber",
          )
        }
      >
        {mode === "bullets"
          ? "Define New Bullet…"
          : mode === "multilevel"
            ? "Define New Multilevel List…"
            : "Define New Number Format…"}
      </button>
      {mode !== "bullets" && (
        <button
          className="home-choice"
          disabled={!selected?.properties.list}
          onClick={() => onDialog("numberValue")}
        >
          Set Numbering Value…
        </button>
      )}
    </>
  );
}
export function ListDefinitionForm({
  mode,
  selected,
  apply,
}: {
  mode: string;
  selected?: Paragraph;
  apply: (command: EditCommand) => void;
}) {
  const id = selected?.id ?? "",
    level = selected?.properties.numbering?.level ?? 0;
  return (
    <form
      onSubmit={(event) => {
        const d = dialogData(event),
          kind = String(
            d.get("kind") ?? selected?.properties.list?.format ?? "decimal",
          ) as ListCommand["kind"];
        apply({
          type: "list",
          id,
          kind: mode === "defineBullet" ? "bullet" : kind,
          level: Number(d.get("level") ?? level),
          ...(mode === "defineBullet"
            ? { bullet: String(d.get("bullet")) }
            : {}),
          ...(mode === "defineNumber"
            ? {
                pattern: String(d.get("pattern")).replace(
                  /%1/g,
                  `%${level + 1}`,
                ),
              }
            : {}),
          ...(mode === "defineMultilevel"
            ? {
                multilevel: String(
                  d.get("multilevel"),
                ) as ListCommand["multilevel"],
              }
            : {}),
          ...(mode === "numberValue"
            ? d.get("mode") === "continue"
              ? { continue: true }
              : { restart: true, start: Number(d.get("start")) }
            : {}),
        });
      }}
    >
      {mode === "defineBullet" ? (
        <label className="field">
          Bullet character
          <input
            name="bullet"
            defaultValue="•"
            maxLength={8}
            required
            autoFocus
          />
        </label>
      ) : mode === "numberValue" ? (
        <>
          <label className="check-field">
            <input type="radio" name="mode" value="restart" defaultChecked />
            Start new list
          </label>
          <label className="check-field">
            <input type="radio" name="mode" value="continue" />
            Continue from previous list
          </label>
          <label className="field">
            Set value to
            <input
              type="number"
              name="start"
              min={1}
              max={1000000}
              defaultValue={1}
            />
          </label>
        </>
      ) : mode === "defineMultilevel" ? (
        <label className="field">
          List format
          <select name="multilevel">
            <option value="decimal">1. → 1.1. → 1.1.1.</option>
            <option value="mixed">1. → a. → i.</option>
            <option value="bullet">• → ○ → ▪</option>
          </select>
        </label>
      ) : (
        <>
          <label className="field">
            Number style
            <select
              name="kind"
              defaultValue={selected?.properties.list?.format ?? "decimal"}
            >
              {[
                ["decimal", "1, 2, 3"],
                ["decimalZero", "01, 02, 03"],
                ["upperRoman", "I, II, III"],
                ["lowerRoman", "i, ii, iii"],
                ["upperLetter", "A, B, C"],
                ["lowerLetter", "a, b, c"],
              ].map(([value, label]) => (
                <option key={value} value={value}>
                  {label}
                </option>
              ))}
            </select>
          </label>
          <label className="field">
            Number format
            <input name="pattern" defaultValue="%1." maxLength={100} required />
          </label>
          <p className="pane-description">
            %1 represents the current number. Add punctuation before or after
            it.
          </p>
        </>
      )}
      <div className="dialog-actions">
        <button className="primary-button">OK</button>
      </div>
    </form>
  );
}
