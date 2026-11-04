import type { RunFormat } from "@local/docx-ir-engine/core";
import { textEffectsStyle } from "@local/docx-ir-engine/react";
import { MenuBranch } from "./HomeMenu";
import { Palette } from "./FontDialogs";

const presets: RunFormat[] = [
  {
    color: "#000000",
    textShadow: {
      color: "#000000",
      distancePt: 1.5,
      angle: 45,
      blurPt: 2,
      opacity: 0.3,
    },
  },
  { color: "#215e79" },
  { color: "#f6c5a8", textOutline: { color: "#c97638", widthPt: 0.5 } },
  { color: "#ffffff", textOutline: { color: "#914b86", widthPt: 0.75 } },
  {
    color: "#00b0f0",
    textShadow: {
      color: "#166782",
      distancePt: 1,
      angle: 45,
      blurPt: 2,
      opacity: 0.6,
    },
  },
  { color: "#9c9c9c" },
  { color: "#943b83", reflection: { distancePt: 0, opacity: 0.4, size: 0.5 } },
  { color: "#9be5f9", textOutline: { color: "#2b90aa", widthPt: 0.6 } },
  {
    color: "#ffffff",
    textOutline: { color: "#306477", widthPt: 0.75 },
    textShadow: {
      color: "#253c49",
      distancePt: 1,
      angle: 45,
      blurPt: 2,
      opacity: 0.6,
    },
  },
  { color: "#205d2e", glow: { color: "#70ad47", radiusPt: 2, opacity: 0.35 } },
  {
    color: "#000000",
    textOutline: { color: "#ffffff", widthPt: 1 },
    textShadow: {
      color: "#000000",
      distancePt: 1,
      angle: 45,
      blurPt: 1,
      opacity: 0.8,
    },
  },
  { color: "#111111", textOutline: { color: "#bd65a8", widthPt: 0.6 } },
  { color: "#9d3a8d", glow: { color: "#d899ce", radiusPt: 3, opacity: 0.6 } },
  { color: "#ffffff", textOutline: { color: "#d88345", widthPt: 1 } },
  {
    color: "#f0f0f0",
    textShadow: {
      color: "#777777",
      distancePt: 1,
      angle: 45,
      blurPt: 2,
      opacity: 0.5,
    },
  },
];
export function EffectsGallery({
  format,
  onPick,
}: {
  format: RunFormat;
  onPick: (format: RunFormat) => void;
}) {
  return (
    <>
      <div className="word-effects-gallery">
        {presets.map((preset, i) => (
          <button
            key={i}
            title={`Text effect ${i + 1}`}
            aria-label={`Text effect ${i + 1}`}
            onClick={() =>
              onPick({
                textOutline: null,
                textShadow: null,
                reflection: null,
                glow: null,
                ...preset,
              })
            }
          >
            <span style={{ ...textEffectsStyle(preset), color: preset.color }}>
              A
            </span>
          </button>
        ))}
      </div>
      <hr />
      <MenuBranch label="Outline">
        <button
          className="home-choice"
          onClick={() => onPick({ textOutline: null })}
        >
          No Outline
        </button>
        <Palette
          onPick={(color) =>
            onPick({
              textOutline: {
                color,
                widthPt: format.textOutline?.widthPt ?? 0.75,
              },
            })
          }
        />
        <MenuBranch label="Weight">
          {[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((widthPt) => (
            <button
              className="home-choice"
              key={widthPt}
              onClick={() =>
                onPick({
                  textOutline: {
                    color: format.textOutline?.color ?? "#000000",
                    widthPt,
                  },
                })
              }
            >
              {widthPt} pt
            </button>
          ))}
        </MenuBranch>
      </MenuBranch>
      <MenuBranch label="Shadow">
        <button
          className="home-choice"
          onClick={() => onPick({ textShadow: null })}
        >
          No Shadow
        </button>
        <h3 className="menu-heading">Outer</h3>
        <div className="effect-variation-grid">
          {[225, 270, 315, 180, 0, 0, 135, 90, 45].map((angle, i) => (
            <button
              key={i}
              title={`Outer shadow ${i + 1}`}
              onClick={() =>
                onPick({
                  textShadow: {
                    color: "#000000",
                    distancePt: i === 4 ? 0 : 2,
                    angle,
                    blurPt: 3,
                    opacity: 0.45,
                  },
                })
              }
            >
              <span
                style={textEffectsStyle({
                  textShadow: {
                    color: "#000000",
                    distancePt: i === 4 ? 0 : 2,
                    angle,
                    blurPt: 3,
                    opacity: 0.45,
                  },
                })}
              >
                A
              </span>
            </button>
          ))}
        </div>
      </MenuBranch>
      <MenuBranch label="Reflection">
        <button
          className="home-choice"
          onClick={() => onPick({ reflection: null })}
        >
          No Reflection
        </button>
        <h3 className="menu-heading">Reflection Variations</h3>
        <div className="effect-variation-grid">
          {[0, 4, 8].flatMap((distancePt) =>
            [0.3, 0.5, 0.8].map((size) => (
              <button
                key={`${distancePt}-${size}`}
                title={`${Math.round(size * 100)}% reflection, ${distancePt} pt offset`}
                onClick={() =>
                  onPick({ reflection: { distancePt, size, opacity: 0.45 } })
                }
              >
                <span
                  style={textEffectsStyle({
                    reflection: { distancePt: 1, size, opacity: 0.45 },
                  })}
                >
                  A
                </span>
              </button>
            )),
          )}
        </div>
      </MenuBranch>
      <MenuBranch label="Glow">
        <button className="home-choice" onClick={() => onPick({ glow: null })}>
          No Glow
        </button>
        <h3 className="menu-heading">Glow Variations</h3>
        <div className="effect-variation-grid">
          {[
            "#5b9bd5",
            "#ed7d31",
            "#a5a5a5",
            "#ffc000",
            "#4472c4",
            "#70ad47",
          ].map((color) => (
            <button
              key={color}
              title={`Glow ${color}`}
              onClick={() =>
                onPick({ glow: { color, radiusPt: 5, opacity: 0.6 } })
              }
            >
              <span
                style={textEffectsStyle({
                  glow: { color, radiusPt: 5, opacity: 0.6 },
                })}
              >
                A
              </span>
            </button>
          ))}
        </div>
        <MenuBranch label="More Glow Colors">
          <Palette
            onPick={(color) =>
              onPick({
                glow: {
                  color,
                  radiusPt: format.glow?.radiusPt ?? 5,
                  opacity: 0.6,
                },
              })
            }
          />
        </MenuBranch>
      </MenuBranch>
      <MenuBranch label="Number Styles">
        {[
          ["Default", "default", "default"],
          ["Proportional Lining", "lining", "proportional"],
          ["Tabular Lining", "lining", "tabular"],
          ["Proportional Old-style", "oldStyle", "proportional"],
          ["Tabular Old-style", "oldStyle", "tabular"],
        ].map(([label, numberForm, numberSpacing]) => (
          <button
            className="home-choice"
            key={label}
            onClick={() =>
              onPick({
                numberForm: numberForm as RunFormat["numberForm"],
                numberSpacing: numberSpacing as RunFormat["numberSpacing"],
              })
            }
          >
            {label}
          </button>
        ))}
      </MenuBranch>
      <MenuBranch label="Ligatures">
        {[
          ["None", "none"],
          ["Standard Only", "standard"],
          ["Standard and Contextual", "standardContextual"],
          ["Historical and Discretionary", "historicalDiscretional"],
          ["All", "all"],
        ].map(([label, ligatures]) => (
          <button
            className="home-choice"
            key={ligatures}
            onClick={() => onPick({ ligatures })}
          >
            {label}
          </button>
        ))}
      </MenuBranch>
      <MenuBranch label="Stylistic Sets">
        <button
          className="home-choice"
          onClick={() => onPick({ stylisticSets: [] })}
        >
          Default
        </button>
        {Array.from({ length: 20 }, (_, i) => (
          <button
            className="home-choice"
            key={i}
            onClick={() => onPick({ stylisticSets: [i + 1] })}
          >
            Set {i + 1}
          </button>
        ))}
      </MenuBranch>
    </>
  );
}
