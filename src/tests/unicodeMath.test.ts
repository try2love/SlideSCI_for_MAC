import { describe, expect, it } from "vitest";
import sampleMarkdown from "../../test.md?raw";
import { markdownToRenderBlocks } from "../services/markdownRender";
import { convertLatexToUnicodeMath } from "../services/unicodeMath";

describe("convertLatexToUnicodeMath", () => {
  it("converts common linear structures", () => {
    expect(convertLatexToUnicodeMath("\\frac{a}{b+c}")).toContain("/");
    expect(convertLatexToUnicodeMath("\\sqrt[5]{a^2}")).toContain("\\sqrt[5]");
    expect(convertLatexToUnicodeMath("\\vec{v}_{d_i}")).toContain("\\vec");
    expect(convertLatexToUnicodeMath("\\matrix{a & b \\\\ c & d}")).toContain("\\matrix");
  });

  it("covers the formulas extracted from test.md", () => {
    const blocks = markdownToRenderBlocks(sampleMarkdown);
    const latexSources = blocks.flatMap((block) => {
      if (block.kind === "math") {
        return [block.content];
      }
      if (block.kind === "text" || block.kind === "quote") {
        return block.equations.map((equation) => equation.latex);
      }
      return [];
    });

    const converted = latexSources.map((latex) => convertLatexToUnicodeMath(latex));
    expect(converted.some((item) => item.includes("δ"))).toBe(true);
    expect(converted.some((item) => item.includes("λ"))).toBe(true);
    expect(converted.some((item) => item.includes("\\sqrt"))).toBe(true);
    expect(converted.some((item) => item.includes("∑"))).toBe(true);
  });

  it("throws for empty LaTeX input", () => {
    expect(() => convertLatexToUnicodeMath("")).toThrow("LaTeX 公式不能为空。");
  });
});
