import React from "react";
import type { DrawingInline } from "../model/document.js";
import { parseEquation, type MathNode } from "../core/equation.js";
export function EquationView({ expression }: { expression: string }) {
  const render = (node: MathNode): React.ReactNode =>
    node.kind === "text"
      ? React.createElement("mtext", null, node.text)
      : node.kind === "row"
        ? React.createElement("mrow", null, ...node.items.map(render))
        : node.kind === "sqrt"
          ? React.createElement("msqrt", null, render(node.a))
          : React.createElement(
              node.kind === "frac"
                ? "mfrac"
                : node.kind === "sup"
                  ? "msup"
                  : "msub",
              null,
              render(node.a),
              render(node.b),
            );
  try {
    return React.createElement(
      "math",
      { display: "inline", style: { fontSize: "18pt" } },
      render(parseEquation(expression)),
    );
  } catch {
    return <span>{expression}</span>;
  }
}
const colors = [
  "#4472C4",
  "#ED7D31",
  "#70AD47",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
];
export function ChartView({
  chart,
}: {
  chart: NonNullable<DrawingInline["chart"]>;
}) {
  const values = chart.series.flatMap((s) => s.values),
    max = Math.max(1, ...values),
    min = Math.min(0, ...values),
    y = (value: number) => 165 - ((value - min) / (max - min)) * 125,
    step = 250 / chart.categories.length;
  let angle = -Math.PI / 2;
  const sum = Math.max(
    0.0001,
    chart.series[0]?.values.reduce((n, v) => n + Math.max(0, v), 0) ?? 0,
  );
  return (
    <svg
      width="100%"
      height="100%"
      viewBox="0 0 320 220"
      aria-label={chart.title}
      role="img"
    >
      <rect width="320" height="220" fill="white" />
      <text x="160" y="18" textAnchor="middle" fontSize="13" fill="#333">
        {chart.title}
      </text>
      {chart.kind === "pie" ? (
        chart.series[0]?.values.map((value, i) => {
          const end = angle + (Math.max(0, value) / sum) * 2 * Math.PI,
            a = angle;
          angle = end;
          return (
            <path
              key={i}
              d={`M160 100 L${160 + 65 * Math.cos(a)} ${100 + 65 * Math.sin(a)} A65 65 0 ${end - a > Math.PI ? 1 : 0} 1 ${160 + 65 * Math.cos(end)} ${100 + 65 * Math.sin(end)}Z`}
              fill={colors[i % colors.length]}
              stroke="white"
            />
          );
        })
      ) : (
        <>
          <line x1="40" x2="290" y1={y(0)} y2={y(0)} stroke="#999" />
          {[0, 0.5, 1].map((t) => (
            <g key={t}>
              <line
                x1="40"
                x2="290"
                y1={40 + t * 125}
                y2={40 + t * 125}
                stroke="#eee"
              />
              <text x="35" y={44 + t * 125} textAnchor="end" fontSize="8">
                {Math.round((max - t * (max - min)) * 100) / 100}
              </text>
            </g>
          ))}
          {chart.series.map((s, j) =>
            chart.kind === "line" ? (
              <polyline
                key={j}
                points={s.values
                  .map((v, i) => `${40 + step * (i + 0.5)},${y(v)}`)
                  .join(" ")}
                fill="none"
                stroke={colors[j % colors.length]}
                strokeWidth="2"
              />
            ) : (
              <g key={j}>
                {s.values.map((v, i) => (
                  <rect
                    key={i}
                    x={
                      40 +
                      step * i +
                      (j * step * 0.8) / chart.series.length +
                      step * 0.1
                    }
                    y={Math.min(y(0), y(v))}
                    width={(step * 0.8) / chart.series.length}
                    height={Math.abs(y(v) - y(0))}
                    fill={colors[j % colors.length]}
                  />
                ))}
              </g>
            ),
          )}
          {chart.categories.map((label, i) => (
            <text
              key={i}
              x={40 + step * (i + 0.5)}
              y="181"
              textAnchor="middle"
              fontSize="8"
            >
              {label.slice(0, 13)}
            </text>
          ))}
        </>
      )}
      {(chart.kind === "pie"
        ? chart.categories
        : chart.series.map((s) => s.name)
      )
        .slice(0, 6)
        .map((name, i) => (
          <g key={i}>
            <rect
              x={15 + (i % 3) * 100}
              y={195 + Math.floor(i / 3) * 12}
              width="7"
              height="7"
              fill={colors[i % colors.length]}
            />
            <text
              x={25 + (i % 3) * 100}
              y={202 + Math.floor(i / 3) * 12}
              fontSize="8"
            >
              {name.slice(0, 16)}
            </text>
          </g>
        ))}
    </svg>
  );
}
