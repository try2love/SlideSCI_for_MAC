import { describe, expect, it } from "vitest";
import { markdownToRichBlocks, parseInlineMarkdown, splitMarkdownIntoSegments } from "../services/markdown";

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

  it("creates inline style runs", () => {
    const parsed = parseInlineMarkdown("A **bold** *italic* `code` [link](https://example.com)");
    expect(parsed.text).toBe("A bold italic code link");
    expect(parsed.runs.some((run) => run.style.bold)).toBe(true);
    expect(parsed.runs.some((run) => run.style.italic)).toBe(true);
    expect(parsed.runs.some((run) => run.style.fontName === "Consolas")).toBe(true);
    expect(parsed.runs.some((run) => run.style.underline)).toBe(true);
  });
});
