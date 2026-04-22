import { marked } from "marked";

export type MarkdownSegment =
  | { kind: "text"; content: string }
  | { kind: "code"; content: string; language: string }
  | { kind: "table"; rows: string[][] }
  | { kind: "math"; content: string }
  | { kind: "quote"; content: string };

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
