import { marked } from "marked";
import type { TextRun, TextStyle } from "../lib/types";

export type MarkdownSegment =
  | { kind: "text"; content: string }
  | { kind: "code"; content: string; language: string }
  | { kind: "table"; rows: string[][] }
  | { kind: "math"; content: string }
  | { kind: "quote"; content: string };

export type MarkdownDocumentBlock =
  | { kind: "markdown"; content: string }
  | { kind: "code"; content: string; language: string }
  | { kind: "table"; rows: string[][] }
  | { kind: "math"; content: string }
  | { kind: "quote"; content: string };

export type MarkdownRichBlock =
  | {
      kind: "richText";
      text: string;
      runs: TextRun[];
      style: TextStyle;
      role: "heading" | "paragraph" | "list" | "orderedList" | "taskList" | "quote";
    }
  | { kind: "code"; content: string; language: string }
  | { kind: "table"; rows: string[][] }
  | { kind: "math"; content: string };

function decodeHtmlEntities(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

export function markdownToPlainText(markdown: string): string {
  const html = marked.parse(markdown, { async: false }) as string;
  return decodeHtmlEntities(
    html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|li|h\d|blockquote)>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .trim(),
  );
}

function stripMarkdown(text: string): string {
  return markdownToPlainText(text)
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .trim();
}

function isFenceStart(line: string): boolean {
  return line.trimStart().startsWith("```");
}

function isMathDelimiter(line: string): boolean {
  return line.trim().startsWith("$$");
}

function isTableSeparator(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed.includes("|")) {
    return false;
  }
  const cells = trimmed.replace(/^\|/, "").replace(/\|$/, "").split("|");
  return cells.length > 1 && cells.every((cell) => /^\s*:?-+:?\s*$/.test(cell));
}

function isTableStart(lines: string[], index: number): boolean {
  return Boolean(lines[index]?.trim().startsWith("|") && lines[index + 1] && isTableSeparator(lines[index + 1]));
}

function splitTableRow(line: string): string[] {
  return line
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

function normalizeTableCell(cell: string): string {
  return stripMarkdown(cell);
}

export function parseMarkdownTable(table: string): string[][] {
  const lines = table
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length < 2 || !isTableSeparator(lines[1])) {
    return [];
  }

  return lines
    .filter((_, index) => index !== 1)
    .map((line) => splitTableRow(line).map(normalizeTableCell));
}

function parseMathBlock(lines: string[], startIndex: number): { content: string; nextIndex: number } {
  const firstLine = lines[startIndex].trim();
  const firstContent = firstLine.slice(2).trim();
  if (firstContent.endsWith("$$") && firstContent.length > 2) {
    return { content: firstContent.slice(0, -2).trim(), nextIndex: startIndex + 1 };
  }

  const mathLines: string[] = [];
  if (firstContent && firstContent !== "$$") {
    mathLines.push(firstContent);
  }

  let index = startIndex + 1;
  while (index < lines.length) {
    const line = lines[index];
    const trimmed = line.trim();
    if (trimmed.endsWith("$$")) {
      const closingContent = line.replace(/\$\$\s*$/, "").trim();
      if (closingContent) {
        mathLines.push(closingContent);
      }
      index += 1;
      break;
    }
    mathLines.push(line);
    index += 1;
  }

  return { content: mathLines.join("\n").trim(), nextIndex: index };
}

export function splitMarkdownDocument(markdown: string): MarkdownDocumentBlock[] {
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const blocks: MarkdownDocumentBlock[] = [];
  const buffer: string[] = [];

  function flushMarkdown(): void {
    const content = buffer.join("\n").trim();
    if (content) {
      blocks.push({ kind: "markdown", content });
    }
    buffer.length = 0;
  }

  let index = 0;
  while (index < lines.length) {
    const line = lines[index];
    const trimmed = line.trim();

    if (isFenceStart(line)) {
      flushMarkdown();
      const language = trimmed.slice(3).trim() || "text";
      const codeLines: string[] = [];
      index += 1;
      while (index < lines.length && !isFenceStart(lines[index])) {
        codeLines.push(lines[index]);
        index += 1;
      }
      blocks.push({ kind: "code", language, content: codeLines.join("\n") });
      index += index < lines.length ? 1 : 0;
      continue;
    }

    if (isMathDelimiter(line)) {
      flushMarkdown();
      const math = parseMathBlock(lines, index);
      if (math.content) {
        blocks.push({ kind: "math", content: math.content });
      }
      index = math.nextIndex;
      continue;
    }

    if (isTableStart(lines, index)) {
      flushMarkdown();
      const tableLines = [lines[index], lines[index + 1]];
      index += 2;
      while (index < lines.length && lines[index].trim().startsWith("|")) {
        tableLines.push(lines[index]);
        index += 1;
      }
      const rows = parseMarkdownTable(tableLines.join("\n"));
      if (rows.length > 0) {
        blocks.push({ kind: "table", rows });
      }
      continue;
    }

    if (trimmed.startsWith(">")) {
      flushMarkdown();
      const quoteLines: string[] = [];
      while (index < lines.length && lines[index].trim().startsWith(">")) {
        quoteLines.push(lines[index].replace(/^\s*>\s?/, ""));
        index += 1;
      }
      const content = quoteLines.join("\n").trim();
      if (content) {
        blocks.push({ kind: "quote", content });
      }
      continue;
    }

    buffer.push(line);
    index += 1;
  }

  flushMarkdown();
  return blocks;
}

export function splitMarkdownIntoSegments(markdown: string): MarkdownSegment[] {
  return splitMarkdownDocument(markdown).flatMap((block): MarkdownSegment[] => {
    if (block.kind === "markdown") {
      const text = stripMarkdown(block.content);
      return text ? [{ kind: "text", content: text }] : [];
    }
    if (block.kind === "table") {
      return [{ kind: "table", rows: block.rows }];
    }
    return [block];
  });
}

function mergeStyle(base: TextStyle, extra: TextStyle): TextStyle {
  return { ...base, ...extra };
}

function appendStyledText(output: { text: string; runs: TextRun[] }, text: string, style: TextStyle): void {
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

function inlineMathMatch(input: string): RegExpMatchArray | null {
  const match = input.match(/(^|[^\\])\$([^$\n]+)\$/);
  if (!match || match.index === undefined) {
    return null;
  }
  if (match[1]) {
    match.index += 1;
    match[0] = match[0].slice(1);
  }
  return match;
}

export function parseInlineMarkdown(input: string, inheritedStyle: TextStyle = {}): { text: string; runs: TextRun[] } {
  const output = { text: "", runs: [] as TextRun[] };
  let remaining = input;

  while (remaining.length > 0) {
    const mathMatch = inlineMathMatch(remaining);
    const found = firstInlineMatch(remaining);
    if (mathMatch?.index !== undefined && (!found || mathMatch.index < found.index)) {
      const before = remaining.slice(0, mathMatch.index);
      appendStyledText(output, before, inheritedStyle);
      appendStyledText(output, `⟦math:${mathMatch[2]}⟧`, mergeStyle(inheritedStyle, { fontName: "Consolas", color: "#7a3e9d" }));
      remaining = remaining.slice(mathMatch.index + mathMatch[0].length);
      continue;
    }

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

function pushRichText(
  blocks: MarkdownRichBlock[],
  text: string,
  style: TextStyle,
  role: "heading" | "paragraph" | "list" | "orderedList" | "taskList" | "quote",
): void {
  const parsed = parseInlineMarkdown(text);
  if (!parsed.text.trim()) {
    return;
  }
  blocks.push({ kind: "richText", text: parsed.text, runs: parsed.runs, style, role });
}

function listItemText(item: any): string {
  const rawText = String(item.text ?? item.raw ?? "");
  return rawText.replace(/^\s*\[[ xX]\]\s*/, "").trim();
}

function pushListItems(blocks: MarkdownRichBlock[], token: any, depth = 0): void {
  const ordered = Boolean(token.ordered);
  const start = Number(token.start ?? 1);
  (token.items ?? []).forEach((item: any, index: number) => {
    const isTask = Boolean(item.task) || /^\s*\[[ xX]\]\s*/.test(String(item.text ?? ""));
    const prefix = isTask ? (item.checked ? "☑ " : "☐ ") : ordered ? `${start + index}. ` : "• ";
    const role = isTask ? "taskList" : ordered ? "orderedList" : "list";
    const itemText = listItemText(item);
    pushRichText(blocks, `${"  ".repeat(depth)}${prefix}${itemText}`, { fontName: "微软雅黑", fontSize: 14, color: "#000000" }, role);
    for (const child of item.tokens ?? []) {
      if (child.type === "list") {
        pushListItems(blocks, child, depth + 1);
      }
    }
  });
}

function pushInlineMathAwareParagraph(blocks: MarkdownRichBlock[], text: string, style: TextStyle, role: "paragraph"): void {
  let remaining = text;
  let pending = "";

  while (remaining.length > 0) {
    const match = inlineMathMatch(remaining);
    if (!match?.index && match?.index !== 0) {
      pending += remaining;
      break;
    }

    pending += remaining.slice(0, match.index);
    if (pending.trim()) {
      pushRichText(blocks, pending.trim(), style, role);
    }
    blocks.push({ kind: "math", content: match[2].trim() });
    pending = "";
    remaining = remaining.slice(match.index + match[0].length);
  }

  if (pending.trim()) {
    pushRichText(blocks, pending.trim(), style, role);
  }
}

function markdownPartToRichBlocks(markdown: string): MarkdownRichBlock[] {
  const tokens = marked.lexer(markdown, { gfm: true, breaks: false }) as any[];
  const blocks: MarkdownRichBlock[] = [];

  for (const token of tokens) {
    if (token.type === "space" || token.type === "hr") {
      continue;
    }
    if (token.type === "heading") {
      const fontSize = Math.max(16, 28 - Number(token.depth ?? 1) * 3);
      pushRichText(blocks, token.text ?? "", { fontName: "微软雅黑", fontSize, color: "#000000", bold: true }, "heading");
    } else if (token.type === "paragraph") {
      pushInlineMathAwareParagraph(blocks, token.text ?? "", { fontName: "微软雅黑", fontSize: 14, color: "#000000" }, "paragraph");
    } else if (token.type === "list") {
      pushListItems(blocks, token);
    } else if (token.type === "html") {
      pushRichText(blocks, markdownToPlainText(token.raw ?? token.text ?? ""), { fontName: "微软雅黑", fontSize: 14, color: "#000000" }, "paragraph");
    } else if (token.type === "code") {
      blocks.push({ kind: "code", content: token.text ?? "", language: token.lang ?? "text" });
    } else if (token.type === "blockquote") {
      pushRichText(blocks, markdownToPlainText(token.text ?? token.raw ?? ""), { fontName: "微软雅黑", fontSize: 14, color: "#000000", fillColor: "#ffffff", borderColor: "#000000" }, "quote");
    }
  }

  return blocks;
}

export function markdownToRichBlocks(markdown: string): MarkdownRichBlock[] {
  return splitMarkdownDocument(markdown).flatMap((block): MarkdownRichBlock[] => {
    if (block.kind === "markdown") {
      return markdownPartToRichBlocks(block.content);
    }
    if (block.kind === "quote") {
      const parsed = parseInlineMarkdown(markdownToPlainText(block.content));
      return [
        {
          kind: "richText",
          text: parsed.text,
          runs: parsed.runs,
          style: { fontName: "微软雅黑", fontSize: 14, color: "#000000", fillColor: "#ffffff", borderColor: "#000000" },
          role: "quote",
        },
      ];
    }
    return [block];
  });
}
