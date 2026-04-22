export type SortMode = "position" | "selection";
export type LayoutMode = "columnMaxWidth" | "uniformHeight" | "waterfall";
export type TitlePlacement = "top" | "bottom";

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SlideShape extends Box {
  id: string;
  name?: string;
  type?: string;
}

export interface ShapeLayout extends Box {
  id: string;
}

export interface LayoutOptions {
  colNum: number;
  colSpace: number;
  rowSpace: number;
  mode: LayoutMode;
  sortMode: SortMode;
  customWidth?: number;
  customHeight?: number;
}

export interface TextStyle {
  fontName?: string;
  fontSize?: number;
  color?: string;
  fillColor?: string;
  borderColor?: string;
  borderWeight?: number;
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center";
  fontFamily?: string;
  underline?: boolean;
  subscript?: boolean;
  superscript?: boolean;
}

export interface TextRunStyle extends TextStyle {
  tokenType?: string;
}

export interface TextRun {
  start: number;
  length: number;
  style: TextRunStyle;
}

export interface NativeEquationRun {
  start: number;
  length: number;
  latex: string;
  display?: boolean;
}
