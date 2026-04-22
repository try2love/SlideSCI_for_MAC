import Prism from "prismjs";
import "prismjs/components/prism-clike";
import "prismjs/components/prism-javascript";
import "prismjs/components/prism-python";
import "prismjs/components/prism-r";
import "prismjs/components/prism-markup";
import "prismjs/components/prism-css";
import "prismjs/components/prism-c";
import "prismjs/components/prism-cpp";
import "prismjs/components/prism-java";
import "prismjs/components/prism-csharp";
import "prismjs/components/prism-fortran";
import "prismjs/components/prism-matlab";
import type { TextRun, TextStyle } from "../lib/types";

export const CODE_LANGUAGES = ["matlab", "python", "r", "javascript", "html", "css", "c", "cpp", "java", "csharp", "fortran"];

export const CODE_LANGUAGE_LABELS: Record<string, string> = {
  c: "C",
  cpp: "C++",
  java: "Java",
  csharp: "C#",
};

const languageAliases: Record<string, string> = {
  js: "javascript",
  html: "markup",
  htm: "markup",
  csharp: "csharp",
  "c#": "csharp",
  cs: "csharp",
  "c++": "cpp",
  cc: "cpp",
  cxx: "cpp",
  hpp: "cpp",
  h: "c",
  py: "python",
  m: "matlab",
  f90: "fortran",
  f95: "fortran",
  for: "fortran",
};

const darkColors: Record<string, string> = {
  keyword: "#569cd6",
  builtin: "#4ec9b0",
  function: "#dcdcaa",
  string: "#ce9178",
  number: "#b5cea8",
  comment: "#6a9955",
  operator: "#d4d4d4",
  punctuation: "#d4d4d4",
  property: "#9cdcfe",
  selector: "#d7ba7d",
  tag: "#569cd6",
  attrName: "#9cdcfe",
  attrValue: "#ce9178",
  boolean: "#569cd6",
  className: "#4ec9b0",
};

const lightColors: Record<string, string> = {
  keyword: "#0000ff",
  builtin: "#267f99",
  function: "#795e26",
  string: "#a31515",
  number: "#098658",
  comment: "#008000",
  operator: "#000000",
  punctuation: "#000000",
  property: "#001080",
  selector: "#800000",
  tag: "#800000",
  attrName: "#ff0000",
  attrValue: "#0451a5",
  boolean: "#0000ff",
  className: "#267f99",
};

function normalizeLanguage(language: string): string {
  const key = language.trim().toLowerCase();
  return languageAliases[key] ?? key;
}

function tokenTypeName(type: string): string {
  return type.replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase());
}

function colorForToken(type: string, darkBackground: boolean): string | undefined {
  const colors = darkBackground ? darkColors : lightColors;
  const normalized = tokenTypeName(type);
  return colors[normalized] ?? colors[type];
}

function flattenTokens(tokens: Array<string | Prism.Token>, darkBackground: boolean, offset = 0): { runs: TextRun[]; length: number } {
  const runs: TextRun[] = [];
  let cursor = offset;

  for (const token of tokens) {
    if (typeof token === "string") {
      cursor += token.length;
      continue;
    }

    const content = Array.isArray(token.content) ? token.content : [token.content as string | Prism.Token];
    const nested = flattenTokens(content, darkBackground, cursor);
    runs.push(...nested.runs);

    const tokenLength = nested.length - cursor;
    const color = colorForToken(token.type, darkBackground);
    if (color && tokenLength > 0) {
      runs.push({
        start: cursor,
        length: tokenLength,
        style: { color, tokenType: token.type },
      });
    }
    cursor = nested.length;
  }

  return { runs, length: cursor };
}

export function getCodeBlockStyle(darkBackground: boolean): TextStyle {
  return {
    fontName: "Consolas",
    fontSize: 12,
    color: darkBackground ? "#ffffff" : "#000000",
    fillColor: darkBackground ? "#1e1e1e" : "#ffffff",
    borderColor: "#c8c8c8",
    align: "left",
  };
}

export function getCodeHighlightRuns(code: string, language: string, darkBackground: boolean): TextRun[] {
  const prismLanguage = normalizeLanguage(language);
  const grammar = Prism.languages[prismLanguage];
  if (!grammar) {
    return [];
  }

  const tokens = Prism.tokenize(code, grammar);
  return flattenTokens(tokens, darkBackground).runs.sort((a, b) => a.start - b.start || b.length - a.length);
}
