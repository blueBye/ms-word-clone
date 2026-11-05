import {
  RibbonCommand as Command,
  RibbonDropdown as Dropdown,
  RibbonChoice as Choice,
} from "./RibbonControls";
export function ArrangeControls({
  count,
  group,
  onAction,
}: {
  count: number;
  group: boolean;
  onAction: (name: string) => void;
}) {
  const action = (name: string) => () => onAction(`object:${name}`);
  return (
    <>
      <Dropdown label="Position" icon="position" large disabled={!count}>
        <Choice label="In Line with Text" onClick={action("inline")} />
        {[
          "top-left",
          "top-center",
          "top-right",
          "middle-left",
          "middle-center",
          "middle-right",
          "bottom-left",
          "bottom-center",
          "bottom-right",
        ].map((value) => (
          <Choice
            key={value}
            label={value.replaceAll("-", " ")}
            onClick={action(`position:${value}`)}
          />
        ))}
        <Choice label="More Layout Options…" onClick={action("options")} />
      </Dropdown>
      <Dropdown label="Wrap Text" icon="wrap" large disabled={!count}>
        {[
          ["In Line with Text", "inline"],
          ["Square", "wrap:wrapSquare"],
          ["Tight", "wrap:wrapTight"],
          ["Through", "wrap:wrapThrough"],
          ["Top and Bottom", "wrap:wrapTopAndBottom"],
          ["Behind Text", "behind"],
          ["In Front of Text", "front"],
        ].map(([label, value]) => (
          <Choice key={value} label={label} onClick={action(value)} />
        ))}
        <Choice label="More Layout Options…" onClick={action("options")} />
      </Dropdown>
      <Dropdown label="Bring Forward" icon="forward" large disabled={!count}>
        <Choice label="Bring Forward" onClick={action("forward")} />
        <Choice label="Bring to Front" onClick={action("toFront")} />
      </Dropdown>
      <Dropdown label="Send Backward" icon="backward" large disabled={!count}>
        <Choice label="Send Backward" onClick={action("backward")} />
        <Choice label="Send to Back" onClick={action("toBack")} />
      </Dropdown>
      <Command
        label="Selection Pane"
        icon="selection"
        large
        onClick={() => onAction("objects")}
      />
      <div className="office-stack">
        <Dropdown label="Align" icon="alignLeft" disabled={!count}>
          {["Left", "Center", "Right", "Top", "Middle", "Bottom"].map(
            (value) => (
              <Choice
                key={value}
                label={`Align ${value}`}
                onClick={action(`align:${value.toLowerCase()}`)}
              />
            ),
          )}
          <Choice
            label="Distribute Horizontally"
            disabled={count < 3}
            onClick={action("distribute:horizontal")}
          />
          <Choice
            label="Distribute Vertically"
            disabled={count < 3}
            onClick={action("distribute:vertical")}
          />
        </Dropdown>
        <Dropdown label="Group" icon="group" disabled={!count}>
          <Choice
            label="Group"
            disabled={count < 2}
            onClick={action("group")}
          />
          <Choice
            label="Ungroup"
            disabled={!group}
            onClick={action("ungroup")}
          />
        </Dropdown>
        <Dropdown label="Rotate" icon="rotate" disabled={!count}>
          {[
            ["Rotate Right 90°", "rotate:90"],
            ["Rotate Left 90°", "rotate:-90"],
            ["Flip Vertical", "flipV"],
            ["Flip Horizontal", "flipH"],
          ].map(([label, value]) => (
            <Choice key={value} label={label} onClick={action(value)} />
          ))}
          <Choice label="More Rotation Options…" onClick={action("options")} />
        </Dropdown>
      </div>
    </>
  );
}
