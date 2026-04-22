import { sortShapesByPosition } from "./sort";
import type { SlideShape } from "./types";

export const LABEL_TEMPLATES = [
  "A",
  "a",
  "A)",
  "a)",
  "(A)",
  "(a)",
  "1",
  "1)",
  "Ⅰ",
  "Ⅰ)",
  "①",
  "①)",
  "一",
  "一)",
] as const;

export type LabelTemplate = (typeof LABEL_TEMPLATES)[number];

const templateCharacters: Record<LabelTemplate, string[]> = {
  A: "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split(""),
  a: "abcdefghijklmnopqrstuvwxyz".split(""),
  "A)": "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split(""),
  "a)": "abcdefghijklmnopqrstuvwxyz".split(""),
  "(A)": "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split(""),
  "(a)": "abcdefghijklmnopqrstuvwxyz".split(""),
  "1": [],
  "1)": [],
  "Ⅰ": ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ"],
  "Ⅰ)": ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ"],
  "①": ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"],
  "①)": ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"],
  "一": ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"],
  "一)": ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"],
};

export function formatLabel(template: string, index: number): string {
  const safeTemplate = LABEL_TEMPLATES.includes(template as LabelTemplate) ? (template as LabelTemplate) : "A";
  const oneBasedIndex = Math.max(1, index);

  if (safeTemplate.startsWith("1")) {
    return safeTemplate.endsWith(")") ? `${oneBasedIndex})` : `${oneBasedIndex}`;
  }

  const characters = templateCharacters[safeTemplate];
  const base = characters[(oneBasedIndex - 1) % characters.length];
  if (safeTemplate.startsWith("(") && safeTemplate.endsWith(")")) {
    return `(${base})`;
  }
  if (safeTemplate.endsWith(")")) {
    return `${base})`;
  }
  return base;
}

export function generateLabels(shapes: SlideShape[], template: string, startIndex: number): Array<{ shape: SlideShape; label: string }> {
  return sortShapesByPosition(shapes).map((shape, offset) => ({
    shape,
    label: formatLabel(template, startIndex + offset),
  }));
}
