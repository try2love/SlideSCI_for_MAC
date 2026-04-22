import { describe, expect, it } from "vitest";
import { estimateTextBoxSize, mergeRichTextBlocks } from "../lib/textMetrics";
import type { MarkdownRichBlock } from "../services/markdown";

describe("estimateTextBoxSize", () => {
  it("uses minimum size for short text", () => {
    const size = estimateTextBoxSize("短文本", { fontSize: 14 });
    expect(size.width).toBeGreaterThanOrEqual(180);
    expect(size.height).toBeGreaterThanOrEqual(36);
  });

  it("expands height for multiline code", () => {
    const short = estimateTextBoxSize("print(1)", { fontSize: 12, monospace: true });
    const long = estimateTextBoxSize("print(1)\nprint(2)\nprint(3)", { fontSize: 12, monospace: true });
    expect(long.height).toBeGreaterThan(short.height);
  });

  it("clamps very long lines to max width", () => {
    const size = estimateTextBoxSize("x".repeat(300), { fontSize: 12, monospace: true });
    expect(size.width).toBe(620);
    expect(size.height).toBeGreaterThan(36);
  });
});

describe("mergeRichTextBlocks", () => {
  it("merges consecutive rich text blocks and shifts runs", () => {
    const blocks: MarkdownRichBlock[] = [
      {
        kind: "richText",
        text: "标题",
        runs: [{ start: 0, length: 2, style: { bold: true } }],
        style: { fontSize: 22 },
        role: "heading",
      },
      {
        kind: "richText",
        text: "正文",
        runs: [{ start: 0, length: 2, style: { italic: true } }],
        style: { fontSize: 14 },
        role: "paragraph",
      },
    ];
    const merged = mergeRichTextBlocks(blocks);
    expect(merged).toHaveLength(1);
    expect(merged[0].kind).toBe("mergedRichText");
    if (merged[0].kind === "mergedRichText") {
      expect(merged[0].text).toContain("标题");
      expect(merged[0].text).toContain("正文");
      expect(merged[0].runs.some((run) => run.start > 0 && run.style.italic)).toBe(true);
    }
  });

  it("keeps code/table/math/quote blocks separate", () => {
    const blocks: MarkdownRichBlock[] = [
      { kind: "richText", text: "A", runs: [], style: { fontSize: 14 }, role: "paragraph" },
      { kind: "code", content: "print(1)", language: "python" },
      { kind: "richText", text: "B", runs: [], style: { fontSize: 14 }, role: "paragraph" },
      { kind: "richText", text: "quote", runs: [], style: { fontSize: 14 }, role: "quote" },
      { kind: "richText", text: "C", runs: [], style: { fontSize: 14 }, role: "paragraph" },
    ];
    const merged = mergeRichTextBlocks(blocks);
    expect(merged.map((block) => block.kind)).toEqual(["mergedRichText", "code", "mergedRichText", "richText", "mergedRichText"]);
    expect(merged[3].kind === "richText" && merged[3].role).toBe("quote");
  });
});
