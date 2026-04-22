import { mergeRichTextBlocks } from "../lib/textMetrics";
import type { NativeEquationRun, TextRun, TextStyle } from "../lib/types";
import { markdownToRichBlocks } from "./markdown";

export type MarkdownRenderBlock =
  | { kind: "text"; text: string; runs: TextRun[]; equations: NativeEquationRun[]; fontSize: number }
  | { kind: "quote"; text: string; runs: TextRun[]; equations: NativeEquationRun[]; style: TextStyle }
  | { kind: "code"; content: string; language: string }
  | { kind: "table"; rows: string[][] }
  | { kind: "math"; content: string; display?: boolean };

export interface MarkdownRenderIssue {
  index: number;
  kind: MarkdownRenderBlock["kind"];
  message: string;
}

export interface MarkdownRenderResult {
  successCount: number;
  warnings: MarkdownRenderIssue[];
  failures: MarkdownRenderIssue[];
}

export type MarkdownRenderHandlers = {
  [K in MarkdownRenderBlock["kind"]]: (block: Extract<MarkdownRenderBlock, { kind: K }>, index: number) => Promise<string | void>;
};

function messageFromError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function markdownToRenderBlocks(markdown: string): MarkdownRenderBlock[] {
  return mergeRichTextBlocks(markdownToRichBlocks(markdown)).map((block): MarkdownRenderBlock => {
    if (block.kind === "mergedRichText") {
      return { kind: "text", text: block.text, runs: block.runs, equations: block.equations, fontSize: block.fontSize };
    }
    if (block.kind === "richText") {
      if (block.role === "quote") {
        return { kind: "quote", text: block.text, runs: block.runs, equations: block.equations ?? [], style: block.style };
      }
      return { kind: "text", text: block.text, runs: block.runs, equations: block.equations ?? [], fontSize: block.style.fontSize ?? 14 };
    }
    return block;
  });
}

export async function renderMarkdownBlocks(blocks: MarkdownRenderBlock[], handlers: MarkdownRenderHandlers): Promise<MarkdownRenderResult> {
  const result: MarkdownRenderResult = { successCount: 0, warnings: [], failures: [] };

  for (const [index, block] of blocks.entries()) {
    try {
      const warning = await handlers[block.kind](block as never, index);
      result.successCount += 1;
      if (warning) {
        result.warnings.push({ index, kind: block.kind, message: warning });
      }
    } catch (error) {
      result.failures.push({ index, kind: block.kind, message: messageFromError(error) });
    }
  }

  return result;
}

export function formatMarkdownRenderResult(result: MarkdownRenderResult): string {
  const parts = [`插入 Markdown 完成：成功 ${result.successCount} 个模块`];
  if (result.failures.length > 0) {
    const failures = result.failures
      .slice(0, 3)
      .map((issue) => `${issue.kind} #${issue.index + 1} ${issue.message}`)
      .join("；");
    parts.push(`失败 ${result.failures.length} 个：${failures}`);
  }
  if (result.warnings.length > 0) {
    const warnings = result.warnings
      .slice(0, 3)
      .map((issue) => `${issue.kind} #${issue.index + 1} ${issue.message}`)
      .join("；");
    parts.push(`提示 ${result.warnings.length} 个：${warnings}`);
  }
  return `${parts.join("，")}。`;
}
