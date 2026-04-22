import type { LabelTemplate } from "../lib/labels";
import type { LayoutMode, SortMode } from "../lib/types";

const PREFIX = "slidesci_for_mac:";

export interface AppSettings {
  titleFontName: string;
  titleFontSize: number;
  titleDistance: number;
  titleText: string;
  titleAlign: "left" | "center";
  labelFontName: string;
  labelFontSize: number;
  labelOffsetX: number;
  labelOffsetY: number;
  labelTemplate: LabelTemplate;
  labelBold: boolean;
  labelIndex: number;
  labelAutoUpdate: boolean;
  colNum: number;
  colSpace: number;
  rowSpace: number;
  imgWidth: string;
  imgHeight: string;
  layoutMode: LayoutMode;
  sortMode: SortMode;
  codeDarkBackground: boolean;
  codeLanguage: string;
}

export const defaultSettings: AppSettings = {
  titleFontName: "微软雅黑",
  titleFontSize: 14,
  titleDistance: 0,
  titleText: "图片标题",
  titleAlign: "center",
  labelFontName: "Arial",
  labelFontSize: 12,
  labelOffsetX: -20,
  labelOffsetY: -7,
  labelTemplate: "A",
  labelBold: true,
  labelIndex: 1,
  labelAutoUpdate: true,
  colNum: 3,
  colSpace: 10,
  rowSpace: 25,
  imgWidth: "",
  imgHeight: "",
  layoutMode: "columnMaxWidth",
  sortMode: "position",
  codeDarkBackground: true,
  codeLanguage: "matlab",
};

export function loadSettings(): AppSettings {
  if (typeof localStorage === "undefined") {
    return defaultSettings;
  }

  const raw = localStorage.getItem(`${PREFIX}settings`);
  if (!raw) {
    return defaultSettings;
  }

  try {
    return { ...defaultSettings, ...JSON.parse(raw) } as AppSettings;
  } catch {
    return defaultSettings;
  }
}

export function saveSettings(settings: AppSettings): void {
  localStorage.setItem(`${PREFIX}settings`, JSON.stringify(settings));
}

export interface ClipboardState {
  width?: number;
  height?: number;
  centers?: Array<{ left: number; top: number }>;
}

export function loadClipboardState(): ClipboardState {
  if (typeof localStorage === "undefined") {
    return {};
  }
  const raw = localStorage.getItem(`${PREFIX}clipboard`);
  if (!raw) {
    return {};
  }
  try {
    return JSON.parse(raw) as ClipboardState;
  } catch {
    return {};
  }
}

export function saveClipboardState(state: ClipboardState): void {
  localStorage.setItem(`${PREFIX}clipboard`, JSON.stringify(state));
}
