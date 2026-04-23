import type { Box } from "./types";
import { estimateTextBoxSize } from "./textMetrics";
import type { MarkdownRenderBlock } from "../services/markdownRender";

export interface SlidePageSize {
  width: number;
  height: number;
}

export interface MarkdownLayoutItem {
  block: MarkdownRenderBlock;
  box: Box;
}

export interface MarkdownLayoutOptions {
  marginX?: number;
  marginTop?: number;
  blockSpacing?: number;
}

const defaultMarkdownLayout = {
  marginX: 60,
  marginTop: 54,
  blockSpacing: 12,
  fallbackPageWidth: 960,
  fallbackPageHeight: 540,
};

function pageDimension(value: number, fallback: number): number {
  return Number.isFinite(value) && value > 0 ? value : fallback;
}

function contentWidth(pageWidth: number, marginX: number): number {
  return Math.max(240, pageWidth - marginX * 2);
}

function tableSize(rows: string[][], maxWidth: number): Pick<Box, "width" | "height"> {
  const longestRow = Math.max(0, ...rows.map((row) => row.join("").length));
  return {
    width: Math.min(maxWidth, Math.max(240, longestRow * 9)),
    height: Math.max(80, rows.length * 28),
  };
}

function blockBox(block: MarkdownRenderBlock, left: number, top: number, maxWidth: number): Box {
  if (block.kind === "text") {
    return {
      left,
      top,
      ...estimateTextBoxSize(block.text, {
        fontSize: block.fontSize,
        monospace: false,
        maxWidth,
      }),
    };
  }
  if (block.kind === "quote") {
    return {
      left,
      top,
      ...estimateTextBoxSize(block.text, {
        fontSize: block.style.fontSize ?? 14,
        monospace: false,
        maxWidth,
      }),
    };
  }
  if (block.kind === "code") {
    return {
      left,
      top,
      ...estimateTextBoxSize(block.content, {
        fontSize: 12,
        monospace: true,
        maxWidth,
      }),
    };
  }
  if (block.kind === "table") {
    return { left, top, ...tableSize(block.rows, maxWidth) };
  }

  const size = estimateTextBoxSize(block.content, {
    fontSize: 18,
    monospace: false,
    minWidth: 220,
    maxWidth,
  });
  return {
    left,
    top,
    width: Math.max(Math.min(360, maxWidth), size.width),
    height: Math.max(54, size.height),
  };
}

export function createMarkdownSingleColumnLayout(
  blocks: MarkdownRenderBlock[],
  pageSize: SlidePageSize,
  options: MarkdownLayoutOptions = {},
): MarkdownLayoutItem[] {
  const pageWidth = pageDimension(pageSize.width, defaultMarkdownLayout.fallbackPageWidth);
  pageDimension(pageSize.height, defaultMarkdownLayout.fallbackPageHeight);
  const marginX = options.marginX ?? defaultMarkdownLayout.marginX;
  const marginTop = options.marginTop ?? defaultMarkdownLayout.marginTop;
  const spacing = options.blockSpacing ?? defaultMarkdownLayout.blockSpacing;
  const left = marginX;
  const maxWidth = contentWidth(pageWidth, marginX);
  let top = marginTop;

  return blocks.map((block) => {
    const box = blockBox(block, left, top, maxWidth);
    top += box.height + spacing;
    return { block, box };
  });
}
