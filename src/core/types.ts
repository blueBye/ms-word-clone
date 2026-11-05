import type {
  DocumentIR,
  ParagraphProperties,
  RunProperties,
} from "../model/document.js";
import type { SearchOptions } from "./search.js";

export interface DocumentStyle {
  id: string;
  name: string;
  basedOn?: string;
  gallery: boolean;
  run: RunProperties;
  paragraph: ParagraphProperties;
}

export interface RunFormat {
  textOutline?: RunProperties["textOutline"];
  textShadow?: RunProperties["textShadow"];
  glow?: RunProperties["glow"];
  reflection?: RunProperties["reflection"];
  ligatures?: RunProperties["ligatures"];
  numberForm?: RunProperties["numberForm"];
  numberSpacing?: RunProperties["numberSpacing"];
  stylisticSets?: RunProperties["stylisticSets"];
  clear?: boolean;
  bold?: boolean;
  italic?: boolean;
  underline?:
    | "single"
    | "double"
    | "thick"
    | "dotted"
    | "dash"
    | "dotDash"
    | "dotDotDash"
    | "wave"
    | "none";
  strike?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  highlight?: string;
  verticalAlign?: "baseline" | "superscript" | "subscript";
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
  letterSpacingPt?: number;
  positionPt?: number;
  complexScriptFont?: string;
  complexScriptSizePt?: number;
  complexScriptBold?: boolean;
  complexScriptItalic?: boolean;
}
export interface ParagraphFormat {
  direction?: "rtl" | "ltr";
  alignment?: "left" | "center" | "right" | "justify";
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  indentStartPt?: number;
  indentEndPt?: number;
  lineSpacing?: number;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  styleId?: string;
  firstLinePt?: number;
  shading?: string;
  borders?: "none" | "all" | "outside" | "top" | "bottom" | "left" | "right";
}
export type EditCommand =
  | { type: "columns"; count: number; gapPt?: number; separator?: boolean }
  | {
      type: "insertImage";
      id?: string;
      dataUrl: string;
      widthPt: number;
      heightPt: number;
      alt: string;
    }
  | {
      type: "link";
      id: string;
      start: number;
      end: number;
      href: string;
      text?: string;
    }
  | {
      type: "story";
      kind: "header" | "footer";
      variant?: "default" | "first" | "even";
      remove?: boolean;
      appendText?: boolean;
      pageNumber?: boolean;
      alignment?: "left" | "center" | "right";
    }
  | {
      type: "text";
      id: string;
      start: number;
      end: number;
      text: string;
      format?: RunFormat;
    }
  | { type: "run"; id: string; format: RunFormat; start?: number; end?: number }
  | { type: "clearRun"; id: string; start?: number; end?: number }
  | {
      type: "splitParagraph";
      id: string;
      start: number;
      end: number;
      format?: RunFormat;
    }
  | { type: "joinParagraph"; id: string; direction: "backward" | "forward" }
  | {
      type: "updateStyleFromSelection";
      id: string;
      styleId: string;
      start: number;
      end: number;
    }
  | {
      type: "list";
      id: string;
      kind:
        | "bullet"
        | "decimal"
        | "decimalZero"
        | "lowerLetter"
        | "upperLetter"
        | "upperRoman"
        | "lowerRoman"
        | "none";
      level?: number;
      start?: number;
      bullet?: string;
      restart?: boolean;
      continue?: boolean;
      preserveDefinition?: boolean;
      pattern?: string;
      multilevel?: "decimal" | "mixed" | "bullet";
    }
  | {
      type: "style";
      styleId: string;
      name?: string;
      basedOn?: string;
      gallery?: boolean;
      run?: RunFormat;
      paragraph?: ParagraphFormat;
    }
  | {
      type: "sort";
      ids: string[];
      descending?: boolean;
      numeric?: boolean;
      locale?: string;
    }
  | { type: "paragraph"; id: string; format: ParagraphFormat }
  | {
      type: "insertParagraph";
      id?: string;
      text: string;
      pageBreakBefore?: boolean;
    }
  | { type: "insertTable"; id?: string; rows: number; columns: number }
  | {
      type: "page";
      widthPt: number;
      heightPt: number;
      margins?: { top: number; bottom: number; left: number; right: number };
    };
export interface ParagraphCapability {
  text: boolean;
  format: boolean;
  reason?: string;
}
export interface DocumentSession {
  readonly statistics: {
    paragraphs: number;
    words: number;
    resolvedParagraphs: number;
    revision: number;
  };
  readonly document: DocumentIR;
  readonly dirty: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly canSetPage: boolean;
  readonly lockedReason: string | undefined;
  readonly selection: { id: string; start: number; end: number } | undefined;
  find(
    query: string,
    options?: SearchOptions,
  ): { id: string; content: string; start: number; end: number }[];
  styles(): Promise<DocumentStyle[]>;
  capability(id: string): ParagraphCapability;
  execute(command: EditCommand): Promise<DocumentIR>;
  executeBatch(commands: EditCommand[]): Promise<DocumentIR>;
  undo(): Promise<DocumentIR>;
  redo(): Promise<DocumentIR>;
}
