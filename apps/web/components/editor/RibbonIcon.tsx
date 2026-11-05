import type { ReactNode } from "react";
import { Icon } from "./Icon";

/** Ribbon artwork uses the same canvas and restrained Office colors at both sizes. */
export function RibbonIcon({
  name,
  size = 32,
}: {
  name: string;
  size?: number;
}) {
  const page = <path d="M5 2.5h11l3 3V22H5z" fill="white" stroke="#738699" />;
  const lines = <path d="M8 10h8M8 13h8M8 16h8M8 19h5" stroke="#8596a6" />;
  const blue = "#4479aa",
    gold = "#bd934f";
  const art: Record<string, ReactNode> = {
    cover: (
      <>
        {page}
        <path d="M8 6h8v4H8z" fill={blue} />
        <path d="M8 14h8M8 17h8" stroke="#8596a6" />
      </>
    ),
    blank: page,
    pageBreak: (
      <>
        <path d="M5 2h14v7H5zM5 15h14v7H5z" fill="white" stroke="#738699" />
        <path d="M2 12h20" stroke={blue} strokeDasharray="2 2" />
      </>
    ),
    table: (
      <>
        <path d="M2 4h20v17H2z" fill="white" stroke="#738699" />
        <path d="M2 4h20v5H2z" fill={blue} />
        <path d="M2 9h20M2 15h20M8 9v12M15 9v12" stroke="#738699" />
      </>
    ),
    picture: (
      <>
        <path d="M2 4h20v17H2z" fill="#e4f0f8" stroke="#738699" />
        <circle cx="17" cy="8" r="2" fill={gold} />
        <path d="m3 20 6-10 5 7 3-4 4 7" fill="#8da875" stroke="#6e8d59" />
      </>
    ),
    shapes: (
      <>
        <path d="M3 3h12v11H3z" fill="#d9e8f5" stroke={blue} />
        <circle cx="15" cy="15" r="7" fill="#f0d4a8" stroke={gold} />
      </>
    ),
    icons: (
      <>
        <path
          d="m8 2 2 5 5 1-4 4 1 6-4-3-5 3 1-6-3-4 5-1z"
          fill="#b4c6d8"
          stroke={blue}
        />
        <path d="m16 13 6 10H10z" fill="#e1c99c" stroke={gold} />
      </>
    ),
    model: (
      <>
        <path d="m12 2 10 5v11l-10 5-10-5V7z" fill="#e2edf5" stroke={blue} />
        <path d="m2 7 10 5 10-5M12 12v11" stroke={blue} />
      </>
    ),
    smartArt: (
      <>
        <path d="M4 5h16M12 5v7M4 12h16M4 12v4M20 12v4" stroke="#769375" />
        <path
          d="M7 1h10v6H7zM0 16h8v6H0zM16 16h8v6h-8z"
          fill="#a6bea0"
          stroke="#769375"
        />
      </>
    ),
    chart: (
      <>
        <path d="M3 21V9h5v12" fill="#6e9fba" />
        <path d="M10 21V3h5v18" fill="#d4a45c" />
        <path d="M17 21V12h5v9" fill="#94af7c" />
        <path d="M1 22h22" stroke="#738699" />
      </>
    ),
    screenshot: (
      <>
        <path d="M3 5h18v14H3z" fill="#dae6ef" stroke="#738699" />
        <path d="M1 8V2h6M17 2h6v6M1 16v6h6M17 22h6v-6" stroke={blue} />
      </>
    ),
    video: (
      <>
        <path d="M2 5h20v15H2z" fill="#dce5ef" stroke={blue} />
        <path d="m10 8 7 4-7 5z" fill={blue} />
      </>
    ),
    bookmark: <path d="M7 2h10v20l-5-5-5 5z" fill="#e3ba80" stroke={gold} />,
    reference: (
      <>
        {page}
        {lines}
        <path d="M1 6h10M8 3l3 3-3 3" stroke={blue} />
      </>
    ),
    comment: (
      <>
        <path d="M2 3h20v14H9l-5 5v-5H2z" fill="#f4e5bd" stroke={gold} />
        <path d="M6 7h12M6 11h9" stroke={gold} />
      </>
    ),
    header: (
      <>
        {page}
        <path d="M7 5h10v4H7z" fill={blue} />
        {lines}
      </>
    ),
    footer: (
      <>
        {page}
        {lines}
        <path d="M7 17h10v3H7z" fill={blue} />
      </>
    ),
    pageNumber: (
      <>
        {page}
        <text x="9" y="20" fontSize="12" fill={blue} stroke="none">
          #
        </text>
      </>
    ),
    textBox: (
      <>
        <path d="M2 4h20v16H2z" fill="white" stroke={blue} />
        <text
          x="6"
          y="17"
          fontFamily="serif"
          fontSize="17"
          fill={blue}
          stroke="none"
        >
          A
        </text>
        <path d="M1 3h3v3H1zM20 18h3v3h-3z" fill="white" stroke={blue} />
      </>
    ),
    quickParts: (
      <>
        <path
          d="M2 3h9v9H2zM13 3h9v6h-9zM13 11h9v10h-9zM2 14h9v7H2z"
          fill="#dfd3b5"
          stroke={gold}
        />
      </>
    ),
    wordArt: (
      <text
        x="3"
        y="21"
        fontFamily="serif"
        fontStyle="italic"
        fontSize="27"
        fill="#74a3c5"
        stroke={blue}
        strokeWidth=".6"
      >
        A
      </text>
    ),
    dropCap: (
      <>
        <text
          x="1"
          y="19"
          fontFamily="serif"
          fontSize="23"
          fill="#4d5965"
          stroke="none"
        >
          A
        </text>
        <path d="M15 5h8M15 10h8M15 15h8M2 22h21" stroke="#8596a6" />
      </>
    ),
    signature: (
      <>
        <path d="M2 19h20M4 15l11-12 4 4L7 18z" fill="#e6edf4" stroke={blue} />
      </>
    ),
    date: (
      <>
        <path d="M3 4h18v18H3z" fill="white" stroke="#738699" />
        <path d="M3 4h18v5H3z" fill={blue} />
        <path
          d="M7 2v4M17 2v4M7 13h2M13 13h3M7 17h2M13 17h3"
          stroke="#738699"
        />
      </>
    ),
    object: (
      <>
        {page}
        <path d="M1 10h11v10H1z" fill="#d9e8f5" stroke={blue} />
      </>
    ),
    equation: (
      <text
        x="2"
        y="20"
        fontFamily="serif"
        fontSize="26"
        fill={blue}
        stroke="none"
      >
        π
      </text>
    ),
    symbol: (
      <text
        x="2"
        y="20"
        fontFamily="serif"
        fontSize="25"
        fill={blue}
        stroke="none"
      >
        Ω
      </text>
    ),
    margins: (
      <>
        {page}
        <path
          d="M8 5h8v14H8z"
          fill="#e6edf4"
          stroke={blue}
          strokeDasharray="2 1"
        />
      </>
    ),
    orientation: (
      <>
        <path d="M2 2h12v18H2z" fill="white" stroke="#738699" />
        <path d="M7 10h15v12H7z" fill="#e6edf4" stroke={blue} />
        <path d="m16 5 3 3 3-3M19 8V2" stroke={blue} />
      </>
    ),
    size: (
      <>
        {page}
        <path d="M1 2v20M0 4l1-2 2 2M0 20l1 2 2-2M5 24h14" stroke={blue} />
      </>
    ),
    columns: (
      <>
        {page}
        <path
          d="M7 6h4M7 9h4M7 12h4M7 15h4M13 6h4M13 9h4M13 12h4M13 15h4"
          stroke={blue}
        />
      </>
    ),
    lineNumbers: (
      <>
        {page}
        <path d="M10 7h6M10 11h6M10 15h6" stroke="#8596a6" />
        <text x="5" y="15" fontSize="6" fill={blue} stroke="none">
          1
          <tspan x="5" dy="4">
            2
          </tspan>
        </text>
      </>
    ),
    hyphenation: (
      <>
        {page}
        <path d="M8 7h8M8 11h4M14 11h3M8 15h8" stroke={blue} />
      </>
    ),
    position: (
      <>
        {page}
        {lines}
        <path d="M10 8h7v6h-7z" fill="#92b3d0" stroke={blue} />
      </>
    ),
    wrap: (
      <>
        <path
          d="M1 4h22M1 8h6M17 8h6M1 12h6M17 12h6M1 16h6M17 16h6M1 20h22"
          stroke="#8596a6"
        />
        <path d="M9 7h6v10H9z" fill="#92b3d0" stroke={blue} />
      </>
    ),
    forward: (
      <>
        <path d="M9 9h13v13H9z" fill="#eee0c3" stroke={gold} />
        <path d="M2 2h13v13H2z" fill="#b3cce1" stroke={blue} />
      </>
    ),
    backward: (
      <>
        <path d="M2 2h13v13H2z" fill="#b3cce1" stroke={blue} />
        <path d="M9 9h13v13H9z" fill="#eee0c3" stroke={gold} />
      </>
    ),
    selection: (
      <>
        <path d="M2 3h20v19H2z" fill="white" stroke="#738699" />
        <path d="M14 3h8v19h-8z" fill="#d5e4f1" />
        <path d="M16 7h4M16 11h4M16 15h4" stroke={blue} />
        <path d="m5 7 7 5-4 1-1 4z" fill="#738699" />
      </>
    ),
    group: (
      <>
        <path d="M4 4h10v10H4zM10 10h10v10H10z" fill="#dae6ef" stroke={blue} />
        <path d="M1 1h22v22H1z" stroke="#738699" strokeDasharray="2 2" />
      </>
    ),
    rotate: (
      <>
        <path d="M5 10a8 8 0 1 1 2 10M1 6l4 4 4-4" stroke={blue} fill="none" />
        <path d="M10 8h8v8h-8z" fill="#dae6ef" stroke="#738699" />
      </>
    ),
  };
  if (!art[name]) return <Icon name={name} size={size} />;
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      strokeWidth="1"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {art[name]}
    </svg>
  );
}
