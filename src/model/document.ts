/** Portable document model with authored and computed formatting. Measurements are points unless explicit. */
export interface DocumentIR {
  properties?: Record<string, string>;
  type: "document";
  schemaVersion: 2;
  blocks: Block[];
  page?: PageLayout;
  backgroundColor?: string;
  language?: string;
  headers?: StoryVariants;
  footers?: StoryVariants;
  diagnostics?: DocumentDiagnostic[];
}
export interface DocumentDiagnostic {
  code: string;
  message: string;
  part?: string;
}
export type StoryVariants = Partial<
  Record<"first" | "default" | "even", Block[]>
>;
export interface PageLayout {
  sectionStart?: "nextPage" | "continuous" | "evenPage" | "oddPage";
  mirrorMargins?: boolean;
  gutterPt?: number;
  hyphenation?: "none" | "auto";
  lineNumbers?: {
    mode: "continuous" | "page" | "section";
    start: number;
    countBy: number;
    distancePt: number;
  };
  direction?: TextDirection;
  verticalAlignment?: "top" | "center" | "bottom";
  columns?: { count: number; gapPt: number; separator?: boolean };
  widthPt?: number;
  heightPt?: number;
  margins?: Partial<
    Record<"top" | "right" | "bottom" | "left" | "header" | "footer", number>
  >;
  differentFirstPage?: boolean;
  evenAndOddHeaders?: boolean;
  pageNumberStart?: number;
}
export type Block = Paragraph | Table;
export type Inline = TextRun | BreakInline | DrawingInline;
export type TextDirection = "ltr" | "rtl";
export type ParagraphAlignment =
  | "left"
  | "right"
  | "center"
  | "justify"
  | "start"
  | "end";
export interface Border {
  style: string;
  widthPt: number;
  color: string;
  spacePt?: number;
}
export type Borders = Partial<
  Record<"top" | "right" | "bottom" | "left" | "insideH" | "insideV", Border>
>;
export interface ParagraphProperties {
  columnBreakBefore?: boolean;
  suppressLineNumbers?: boolean;
  dropCap?: { lines: number; distancePt: number; kind: "drop" | "margin" };
  styleId?: string;
  alignment?: ParagraphAlignment;
  direction?: TextDirection;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacing?: { value: number; rule: "auto" | "exact" | "atLeast" };
  indentLeftPt?: number;
  indentRightPt?: number;
  indentStartPt?: number;
  indentEndPt?: number;
  firstLinePt?: number;
  tabs?: { positionPt: number; alignment: string; leader?: string }[];
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  contextualSpacing?: boolean;
  backgroundColor?: string;
  borders?: Borders;
  numbering?: { id?: string; level?: number };
  list?: {
    marker: string;
    suffix: string;
    properties: RunProperties;
    locale?: string;
    format?: string;
  };
}
export interface Paragraph {
  section?: PageLayout & { headers?: StoryVariants; footers?: StoryVariants };
  bookmarks?: { name: string; start: number; end: number }[];
  comments?: {
    key: string;
    start: number;
    end: number;
    text: string;
    author: string;
    date: string;
  }[];
  type: "paragraph";
  /** Source location in the current editing session revision; absent in preview-only imports. */
  id?: string;
  properties: ParagraphProperties;
  directProperties?: ParagraphProperties;
  markProperties?: RunProperties;
  /** Style-independent formatting supplied by the containing story/table and document defaults. */
  formattingContext?: { paragraph: ParagraphProperties; run: RunProperties };
  directMarkProperties?: RunProperties;
  listMarkProperties?: RunProperties;
  children: Inline[];
}
export interface ComplexScriptProperties {
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
}
export interface RunProperties {
  textOutline?: { color: string; widthPt: number } | null;
  textShadow?: {
    color: string;
    blurPt: number;
    distancePt: number;
    angle: number;
    opacity: number;
  } | null;
  glow?: { color: string; radiusPt: number; opacity: number } | null;
  reflection?: { distancePt: number; opacity: number; size: number } | null;
  ligatures?: string;
  numberForm?: "default" | "lining" | "oldStyle";
  numberSpacing?: "default" | "proportional" | "tabular";
  stylisticSets?: number[];
  styleId?: string;
  direction?: TextDirection;
  fontFamily?: string;
  fontSlots?: { ascii?: string; hAnsi?: string; eastAsia?: string };
  fontHint?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  complexScript?: ComplexScriptProperties;
  forceComplexScript?: boolean;
  language?: string;
  complexScriptLanguage?: string;
  color?: string;
  backgroundColor?: string;
  underline?: string;
  strike?: boolean;
  verticalAlign?: "baseline" | "superscript" | "subscript";
  positionPt?: number;
  letterSpacingPt?: number;
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
}
export interface TextRun {
  type: "text";
  text: string;
  properties: RunProperties;
  directProperties?: RunProperties;
  href?: string;
  field?: "page" | "pages" | "date" | "title" | "author";
}
export interface BreakInline {
  type: "break";
  kind: "page" | "column";
  properties: RunProperties;
}
export interface DrawingPosition {
  relativeTo: string;
  offsetPt?: number;
  alignment?: string;
}
export interface DrawingInline {
  spreadsheet?: { cells: string[][] };
  model3d?: {
    dataUrl: string;
    rotation?: { x: number; y: number; z: number };
    size?: number;
    center?: [number, number, number];
    poster?: string;
  };
  attachment?: { dataUrl: string; name: string; mime: string; progId?: string };
  video?: { url: string };
  chart?: {
    kind: "bar" | "line" | "pie";
    title: string;
    categories: string[];
    series: { name: string; values: number[] }[];
  };
  equation?: { expression: string };
  id?: string;
  children?: DrawingInline[];
  flipH?: boolean;
  flipV?: boolean;
  type: "drawing";
  widthPt: number;
  heightPt: number;
  alt: string;
  image?: {
    dataUrl: string;
    opacity?: number;
    brightness?: number;
    contrast?: number;
    grayscale?: boolean;
    blurPt?: number;
    crop?: Partial<Record<"left" | "top" | "right" | "bottom", number>>;
  };
  shape?: {
    kind: string;
    fill?: string;
    stroke?: string;
    strokeWidthPt?: number;
    rotation?: number;
  };
  blocks?: Block[];
  position?: {
    horizontal: DrawingPosition;
    vertical: DrawingPosition;
    behindText: boolean;
    zIndex: number;
    wrap: string;
    locked?: boolean;
  };
}
export interface Table {
  type: "table";
  properties: {
    styleId?: string;
    direction?: TextDirection;
    width?: { value: number; unit: "pt" | "percent" };
    alignment?: "left" | "right" | "center";
    indentPt?: number;
    layout?: "fixed" | "autofit";
    borders?: Borders;
    cellMargins?: Partial<Record<"top" | "right" | "bottom" | "left", number>>;
    backgroundColor?: string;
  };
  columnsPt: number[];
  rows: TableRow[];
}
export interface TableRow {
  cells: TableCell[];
  heightPt?: number;
  heightRule?: string;
  header?: boolean;
  cantSplit?: boolean;
}
export interface TableCell {
  blocks: Block[];
  column: number;
  columnSpan: number;
  rowSpan: number;
  continuation?: boolean;
  width?: { value: number; unit: "pt" | "percent" };
  borders?: Borders;
  margins?: Partial<Record<"top" | "right" | "bottom" | "left", number>>;
  backgroundColor?: string;
  verticalAlign?: string;
}
