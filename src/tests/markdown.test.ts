import { describe, expect, it } from "vitest";
import { splitMarkdownIntoSegments } from "../services/markdown";

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
