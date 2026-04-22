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
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center";
  fontFamily?: string;
}
