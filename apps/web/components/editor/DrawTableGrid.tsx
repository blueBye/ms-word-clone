import { useRef, useState } from "react";
export function DrawTableGrid() {
  const [vertical, setVertical] = useState<number[]>([]),
    [horizontal, setHorizontal] = useState<number[]>([]);
  const start = useRef<{ x: number; y: number } | null>(null);
  return (
    <>
      <p className="pane-description">
        Draw vertical or horizontal dividers inside the table. Double-click to
        reset.
      </p>
      <svg
        className="table-drawing-canvas"
        viewBox="0 0 360 180"
        onDoubleClick={() => {
          setVertical([]);
          setHorizontal([]);
        }}
        onPointerDown={(e) => {
          e.preventDefault();
          const r = e.currentTarget.getBoundingClientRect();
          start.current = {
            x: ((e.clientX - r.left) / r.width) * 360,
            y: ((e.clientY - r.top) / r.height) * 180,
          };
          e.currentTarget.setPointerCapture(e.pointerId);
        }}
        onPointerUp={(e) => {
          const from = start.current;
          if (!from) return;
          const r = e.currentTarget.getBoundingClientRect(),
            x = ((e.clientX - r.left) / r.width) * 360,
            y = ((e.clientY - r.top) / r.height) * 180;
          start.current = null;
          if (Math.abs(x - from.x) < Math.abs(y - from.y))
            setVertical((values) =>
              [...values, Math.max(12, Math.min(348, (x + from.x) / 2))]
                .sort((a, b) => a - b)
                .filter((x, i, all) => i === 0 || x - all[i - 1] > 8)
                .slice(0, 11),
            );
          else
            setHorizontal((values) =>
              [...values, Math.max(10, Math.min(170, (y + from.y) / 2))]
                .sort((a, b) => a - b)
                .filter((x, i, all) => i === 0 || x - all[i - 1] > 6)
                .slice(0, 29),
            );
          e.currentTarget.releasePointerCapture(e.pointerId);
        }}
      >
        <rect
          x="1"
          y="1"
          width="358"
          height="178"
          fill="white"
          stroke="#4479AA"
        />
        {vertical.map((x, i) => (
          <line key={`v${i}`} x1={x} x2={x} y1="0" y2="180" stroke="#4479AA" />
        ))}
        {horizontal.map((y, i) => (
          <line key={`h${i}`} x1="0" x2="360" y1={y} y2={y} stroke="#4479AA" />
        ))}
      </svg>
      <input type="hidden" name="rows" value={horizontal.length + 1} />
      <input type="hidden" name="columns" value={vertical.length + 1} />
      <input
        type="hidden"
        name="widths"
        value={[0, ...vertical, 360]
          .slice(1)
          .map((x, i) => x - [0, ...vertical, 360][i])
          .join(",")}
      />
    </>
  );
}
