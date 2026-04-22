import type { Box, TextRun } from "./types";
import type { MarkdownRichBlock } from "../services/markdown";

export interface TextBoxMetricsOptions {
  fontSize: number;
  monospace?: boolean;
  minWidth?: number;
  maxWidth?: number;
  minHeight?: number;
  paddingX?: number;
  paddingY?: number;
}

export const defaultTextBoxMetrics = {
  minWidth: 180,
  maxWidth: 620,
  minHeight: 36,
  paddingX: 18,
  paddingY: 14,
};

function charWidth(char: string, fontSize: number, monospace: boolean): number {
  if (monospace) {
    return fontSize * 0.62;
  }
  if (/[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(char)) {
    return fontSize;
  }
  if (/\s/.test(char)) {
    return fontSize * 0.32;
  }
  if (/[A-Z0-9]/.test(char)) {
    return fontSize * 0.62;
  }
  return fontSize * 0.52;
}

function estimateLineWidth(line: string, fontSize: number, monospace: boolean): number {
  return Array.from(line).reduce((width, char) => width + charWidth(char, fontSize, monospace), 0);
}

export function estimateTextBoxSize(text: string, options: TextBoxMetricsOptions): Pick<Box, "width" | "height"> {
  const minWidth = options.minWidth ?? defaultTextBoxMetrics.minWidth;
  const maxWidth = options.maxWidth ?? defaultTextBoxMetrics.maxWidth;
  const minHeight = options.minHeight ?? defaultTextBoxMetrics.minHeight;
  const paddingX = options.paddingX ?? defaultTextBoxMetrics.paddingX;
  const paddingY = options.paddingY ?? defaultTextBoxMetrics.paddingY;
  const fontSize = options.fontSize;
  const monospace = Boolean(options.monospace);
  const lines = text.length > 0 ? text.split(/\r?\n/) : [""];
  const naturalWidth = Math.max(...lines.map((line) => estimateLineWidth(line, fontSize, monospace))) + paddingX;
  const width = Math.max(minWidth, Math.min(maxWidth, Math.ceil(naturalWidth)));
  const contentWidth = Math.max(1, width - paddingX);
  const wrappedLineCount = lines.reduce((count, line) => {
    const lineWidth = estimateLineWidth(line, fontSize, monospace);
    return count + Math.max(1, Math.ceil(lineWidth / contentWidth));
  }, 0);
  const height = Math.max(minHeight, Math.ceil(wrappedLineCount * fontSize * 1.35 + paddingY));
  return { width, height };
}

function blockSpacing(block: MarkdownRichBlock): number {
  if (block.kind !== "richText") {
    return 0;
  }
  if (block.role === "heading") {
    return 10;
  }
  if (block.role === "list" || block.role === "orderedList" || block.role === "taskList") {
    return 4;
  }
  return 8;
}

export function mergeRichTextBlocks(blocks: MarkdownRichBlock[]): Array<MarkdownRichBlock | { kind: "mergedRichText"; text: string; runs: TextRun[]; fontSize: number }> {
  const result: Array<MarkdownRichBlock | { kind: "mergedRichText"; text: string; runs: TextRun[]; fontSize: number }> = [];
  let text = "";
  let runs: TextRun[] = [];
  let fontSize = 14;

  function flush(): void {
    if (!text.trim()) {
      text = "";
      runs = [];
      fontSize = 14;
      return;
    }
    result.push({ kind: "mergedRichText", text: text.replace(/\n+$/, ""), runs, fontSize });
    text = "";
    runs = [];
    fontSize = 14;
  }

  for (const block of blocks) {
    if (block.kind !== "richText") {
      flush();
      result.push(block);
      continue;
    }

    if (block.role === "quote") {
      flush();
      result.push(block);
      continue;
    }

    const offset = text.length;
    const spacing = text.length > 0 ? "\n".repeat(blockSpacing(block) > 6 ? 2 : 1) : "";
    text += spacing;
    const shiftedOffset = offset + spacing.length;
    text += block.text;
    runs.push(...block.runs.map((run) => ({ ...run, start: run.start + shiftedOffset })));
    runs.push({
      start: shiftedOffset,
      length: block.text.length,
      style: block.style,
    });
    fontSize = Math.max(fontSize, block.style.fontSize ?? 14);
  }
  flush();

  return result;
}
