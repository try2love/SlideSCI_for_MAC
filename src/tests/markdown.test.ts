import { describe, expect, it } from "vitest";
import sampleMarkdown from "../../test.md?raw";
import {
  markdownToRichBlocks,
  parseInlineMarkdown,
  parseMarkdownTable,
  splitMarkdownDocument,
  splitMarkdownIntoSegments,
} from "../services/markdown";

describe("splitMarkdownIntoSegments", () => {
  it("splits text, code, table, math, and quote blocks", () => {
    const segments = splitMarkdownIntoSegments(`# 标题

正文

\`\`\`python
print("hi")
\`\`\`

| A | B |
| - | - |
| 1 | 2 |

$$
x^2
$$

> 引用`);

    expect(segments.map((segment) => segment.kind)).toEqual(["text", "code", "table", "math", "quote"]);
  });

  it("recognizes the fixed test.md special blocks in source order", () => {
    const blocks = splitMarkdownDocument(sampleMarkdown);
    expect(blocks.map((block) => block.kind)).toContain("markdown");
    expect(blocks.map((block) => block.kind)).toContain("table");
    expect(blocks.map((block) => block.kind)).toContain("math");
    expect(blocks.map((block) => block.kind)).toContain("code");
    expect(blocks.map((block) => block.kind)).toContain("quote");

    const firstTable = blocks.find((block) => block.kind === "table");
    const firstMath = blocks.find((block) => block.kind === "math");
    const firstCode = blocks.find((block) => block.kind === "code");
    expect(blocks.indexOf(firstTable!)).toBeLessThan(blocks.indexOf(firstMath!));
    expect(blocks.indexOf(firstMath!)).toBeLessThan(blocks.indexOf(firstCode!));
  });
});

describe("markdownToRichBlocks", () => {
  it("preserves editable rich text structure", () => {
    const blocks = markdownToRichBlocks(`# 标题

这是 **粗体** 和 *斜体*，包含 \`code\`。

- A
- B

\`\`\`javascript
const x = 1
\`\`\`

| A | B |
| - | - |
| 1 | 2 |

$x^2$`);

    expect(blocks.map((block) => block.kind)).toEqual(["richText", "richText", "richText", "richText", "code", "table", "math"]);
    const heading = blocks[0];
    expect(heading.kind).toBe("richText");
    if (heading.kind === "richText") {
      expect(heading.role).toBe("heading");
      expect(heading.style.bold).toBe(true);
    }
  });

  it("creates task, unordered, ordered, table, quote, code, and math blocks from test.md", () => {
    const blocks = markdownToRichBlocks(sampleMarkdown);
    const roles = blocks.flatMap((block) => (block.kind === "richText" ? [block.role] : []));
    const kinds = blocks.map((block) => block.kind);

    expect(roles).toContain("heading");
    expect(roles).toContain("list");
    expect(roles).toContain("taskList");
    expect(roles).toContain("orderedList");
    expect(roles).toContain("quote");
    expect(kinds).toContain("table");
    expect(kinds).toContain("math");
    expect(kinds).toContain("code");

    const text = blocks.flatMap((block) => (block.kind === "richText" ? [block.text] : [])).join("\n");
    expect(text).toContain("☑ 完成本地代理");
    expect(text).toContain("☐ 部署 SlideSCI 插件到生产环境");
    expect(text).toContain("• 核心框架：Python / PyTorch");
    expect(text).toContain("1. 提交本地代码到 feature 分支。");
  });

  it("creates inline style runs", () => {
    const parsed = parseInlineMarkdown("A **bold** *italic* `code` [link](https://example.com)");
    expect(parsed.text).toBe("A bold italic code link");
    expect(parsed.runs.some((run) => run.style.bold)).toBe(true);
    expect(parsed.runs.some((run) => run.style.italic)).toBe(true);
    expect(parsed.runs.some((run) => run.style.fontName === "Consolas")).toBe(true);
    expect(parsed.runs.some((run) => run.style.underline)).toBe(true);
  });

  it("parses aligned tables and strips cell markdown", () => {
    const rows = parseMarkdownTable(`| 模块 | 功能 | 状态 |
| :--- | :---: | ---: |
| **Backend** | RAG 检索逻辑 | \`OK\` |`);
    expect(rows).toEqual([
      ["模块", "功能", "状态"],
      ["Backend", "RAG 检索逻辑", "OK"],
    ]);
  });

  it("extracts block and inline math into math blocks", () => {
    const blocks = markdownToRichBlocks(sampleMarkdown);
    const mathBlocks = blocks.filter((block) => block.kind === "math");
    expect(mathBlocks.length).toBeGreaterThanOrEqual(3);
    expect(mathBlocks.some((block) => block.kind === "math" && block.content.includes("Score"))).toBe(true);
    expect(mathBlocks.some((block) => block.kind === "math" && block.content === "\\delta")).toBe(true);
  });
});
