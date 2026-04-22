import { marked } from "marked";
import type { TextRun, TextStyle } from "../lib/types";

export type MarkdownSegment =
  | { kind: "text"; content: string }
  | { kind: "code"; content: string; language: string }
  | { kind: "table"; rows: string[][] }
  | { kind: "math"; content: string }
  | { kind: "quote"; content: string };

export type MarkdownRichBlock =
  | { kind: "richText"; text: string; runs: TextRun[]; style: TextStyle; role: "heading" | "paragraph" | "list" | "quote" }
  | { kind: "code"; content: string; language: string }
  | { kind: "table"; rows: string[][] }
  | { kind: "math"; content: string };

function parseTable(block: string): string[][] | null {
  const lines = block
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (lines.length < 2 || !lines[0].startsWith("|") || !/^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?$/.test(lines[1])) {
    return null;
  }

  return lines
    .filter((_, index) => index !== 1)
    .map((line) =>
      line
        .replace(/^\|/, "")
        .replace(/\|$/, "")
        .split("|")
        .map((cell) => cell.trim()),
    );
}

function stripMarkdown(text: string): string {
  return text
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .trim();
}

export function splitMarkdownIntoSegments(markdown: string): MarkdownSegment[] {
  const segments: MarkdownSegment[] = [];
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const textBuffer: string[] = [];

  function flushText(): void {
    const text = stripMarkdown(textBuffer.join("\n"));
    if (text) {
      segments.push({ kind: "text", content: text });
    }
    textBuffer.length = 0;
  }

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const trimmed = line.trim();

    if (trimmed.startsWith("```")) {
      flushText();
      const language = trimmed.slice(3).trim() || "text";
      const codeLines: string[] = [];
      index += 1;
      while (index < lines.length && !lines[index].trim().startsWith("```")) {
        codeLines.push(lines[index]);
        index += 1;
      }
      segments.push({ kind: "code", language, content: codeLines.join("\n") });
      continue;
    }

    if (trimmed.startsWith("$$")) {
      flushText();
      const mathLines: string[] = [];
      const firstLineContent = trimmed.slice(2);
      if (firstLineContent && !firstLineContent.endsWith("$$")) {
        mathLines.push(firstLineContent);
      }
      index += 1;
      while (index < lines.length && !lines[index].trim().endsWith("$$")) {
        mathLines.push(lines[index]);
        index += 1;
      }
      if (index < lines.length) {
        const closingLine = lines[index].trim();
        const withoutClosing = closingLine.replace(/\$\$$/, "").trim();
        if (withoutClosing) {
          mathLines.push(withoutClosing);
        }
      }
      segments.push({ kind: "math", content: mathLines.join("\n").trim() });
      continue;
    }

    if (trimmed.startsWith(">")) {
      flushText();
      const quoteLines: string[] = [];
      while (index < lines.length && lines[index].trim().startsWith(">")) {
        quoteLines.push(lines[index].replace(/^\s*> ?/, ""));
        index += 1;
      }
      index -= 1;
      segments.push({ kind: "quote", content: stripMarkdown(quoteLines.join("\n")) });
      continue;
    }

    if (
      trimmed.startsWith("|") &&
      index + 1 < lines.length &&
      /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?\s*$/.test(lines[index + 1])
    ) {
      flushText();
      const tableLines = [line, lines[index + 1]];
      index += 2;
      while (index < lines.length && lines[index].trim().startsWith("|")) {
        tableLines.push(lines[index]);
        index += 1;
      }
      index -= 1;
      const rows = parseTable(tableLines.join("\n"));
      if (rows) {
        segments.push({ kind: "table", rows });
      }
      continue;
    }

    textBuffer.push(line);
  }

  flushText();

  return segments;
}

export function markdownToPlainText(markdown: string): string {
  const html = marked.parse(markdown, { async: false }) as string;
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|li|h\d)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .trim();
}

function mergeStyle(base: TextStyle, extra: TextStyle): TextStyle {
  return { ...base, ...extra };
}

function appendStyledText(
  output: { text: string; runs: TextRun[] },
  text: string,
  style: TextStyle,
): void {
  const start = output.text.length;
  output.text += text;
  if (text.length > 0 && Object.keys(style).length > 0) {
    output.runs.push({ start, length: text.length, style });
  }
}

function firstInlineMatch(input: string): { match: RegExpMatchArray; kind: string; index: number } | null {
  const patterns: Array<[string, RegExp]> = [
    ["code", /`([^`]+)`/],
    ["bold", /\*\*([^*]+)\*\*|__([^_]+)__/],
    ["italic", /\*([^*]+)\*|_([^_]+)_/],
    ["link", /\[([^\]]+)\]\(([^)]+)\)/],
    ["sup", /<sup>(.*?)<\/sup>|\^([^^]+)\^/i],
    ["sub", /<sub>(.*?)<\/sub>|~([^~]+)~/i],
  ];

  let best: { match: RegExpMatchArray; kind: string; index: number } | null = null;
  for (const [kind, pattern] of patterns) {
    const match = input.match(pattern);
    if (match?.index !== undefined && (!best || match.index < best.index)) {
      best = { match, kind, index: match.index };
    }
  }
  return best;
}

export function parseInlineMarkdown(input: string, inheritedStyle: TextStyle = {}): { text: string; runs: TextRun[] } {
  const output = { text: "", runs: [] as TextRun[] };
  let remaining = input;

  while (remaining.length > 0) {
    const found = firstInlineMatch(remaining);
    if (!found) {
      appendStyledText(output, remaining, inheritedStyle);
      break;
    }

    if (found.index > 0) {
      appendStyledText(output, remaining.slice(0, found.index), inheritedStyle);
    }

    const raw = found.match[0];
    const content = found.match[1] ?? found.match[2] ?? "";
    let style: TextStyle;
    if (found.kind === "bold") {
      style = mergeStyle(inheritedStyle, { bold: true });
    } else if (found.kind === "italic") {
      style = mergeStyle(inheritedStyle, { italic: true });
    } else if (found.kind === "code") {
      style = mergeStyle(inheritedStyle, { fontName: "Consolas", color: "#c00000" });
    } else if (found.kind === "link") {
      style = mergeStyle(inheritedStyle, { color: "#0563c1", underline: true });
    } else if (found.kind === "sup") {
      style = mergeStyle(inheritedStyle, { superscript: true });
    } else {
      style = mergeStyle(inheritedStyle, { subscript: true });
    }

    const nested = found.kind === "code" ? { text: content, runs: [] as TextRun[] } : parseInlineMarkdown(content, style);
    const offset = output.text.length;
    output.text += nested.text;
    output.runs.push(...nested.runs.map((run) => ({ ...run, start: run.start + offset })));
    if (nested.runs.length === 0 && nested.text.length > 0) {
      output.runs.push({ start: offset, length: nested.text.length, style });
    }

    remaining = remaining.slice(found.index + raw.length);
  }

  return output;
}

function pushInlineMathAwareBlocks(
  blocks: MarkdownRichBlock[],
  text: string,
  style: TextStyle,
  role: "heading" | "paragraph" | "list" | "quote",
): void {
  const parts = text.split(/(\$[^$\n]+\$)/g).filter(Boolean);
  for (const part of parts) {
    if (part.startsWith("$") && part.endsWith("$")) {
      blocks.push({ kind: "math", content: part.slice(1, -1) });
      continue;
    }
    const parsed = parseInlineMarkdown(part);
    if (parsed.text.trim()) {
      blocks.push({ kind: "richText", text: parsed.text, runs: parsed.runs, style, role });
    }
  }
}

function tableTokenToRows(token: any): string[][] {
  const header = (token.header ?? []).map((cell: any) => markdownToPlainText(cell.text ?? String(cell)));
  const rows = (token.rows ?? []).map((row: any[]) => row.map((cell) => markdownToPlainText(cell.text ?? String(cell))));
  return [header, ...rows].filter((row) => row.length > 0);
}

function pushListItems(blocks: MarkdownRichBlock[], token: any, depth = 0): void {
  const ordered = Boolean(token.ordered);
  (token.items ?? []).forEach((item: any, index: number) => {
    const prefix = ordered ? `${index + 1}. ` : "• ";
    const parsed = parseInlineMarkdown(`${"  ".repeat(depth)}${prefix}${item.text ?? ""}`);
    blocks.push({
      kind: "richText",
      text: parsed.text,
      runs: parsed.runs,
      role: "list",
      style: { fontName: "微软雅黑", fontSize: 14, color: "#000000" },
    });
    for (const child of item.tokens ?? []) {
      if (child.type === "list") {
        pushListItems(blocks, child, depth + 1);
      }
    }
  });
}

export function markdownToRichBlocks(markdown: string): MarkdownRichBlock[] {
  const tokens = marked.lexer(markdown, { gfm: true, breaks: false }) as any[];
  const blocks: MarkdownRichBlock[] = [];

  for (const token of tokens) {
    if (token.type === "space") {
      continue;
    }
    if (token.type === "heading") {
      const fontSize = Math.max(16, 26 - Number(token.depth ?? 1) * 2);
      pushInlineMathAwareBlocks(blocks, token.text ?? "", {
        fontName: "微软雅黑",
        fontSize,
        color: "#000000",
        bold: true,
      }, "heading");
    } else if (token.type === "paragraph") {
      pushInlineMathAwareBlocks(blocks, token.text ?? "", {
        fontName: "微软雅黑",
        fontSize: 14,
        color: "#000000",
      }, "paragraph");
    } else if (token.type === "blockquote") {
      const quoteText = (token.tokens ?? []).map((child: any) => child.text ?? child.raw ?? "").join("\n");
      pushInlineMathAwareBlocks(blocks, quoteText, {
        fontName: "微软雅黑",
        fontSize: 14,
        color: "#000000",
        fillColor: "#ffffff",
        borderColor: "#000000",
      }, "quote");
    } else if (token.type === "code") {
      blocks.push({ kind: "code", content: token.text ?? "", language: token.lang ?? "text" });
    } else if (token.type === "table") {
      blocks.push({ kind: "table", rows: tableTokenToRows(token) });
    } else if (token.type === "list") {
      pushListItems(blocks, token);
    } else if (token.type === "html") {
      pushInlineMathAwareBlocks(blocks, markdownToPlainText(token.raw ?? token.text ?? ""), {
        fontName: "微软雅黑",
        fontSize: 14,
        color: "#000000",
      }, "paragraph");
    }
  }

  return blocks;
}
