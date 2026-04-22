import { describe, expect, it } from "vitest";
import { normalizeLatexInput, shouldUseDisplayMode } from "../lib/latex";

describe("latex helpers", () => {
  it("normalizes common wrappers", () => {
    expect(normalizeLatexInput("$x$")).toBe("x");
    expect(normalizeLatexInput("$$x$$")).toBe("x");
    expect(normalizeLatexInput("\\(x\\)")).toBe("x");
    expect(normalizeLatexInput("\\[x\\]")).toBe("x");
  });

  it("detects display mode", () => {
    expect(shouldUseDisplayMode("$$x$$")).toBe(true);
    expect(shouldUseDisplayMode("\\begin{align}x\\end{align}")).toBe(true);
    expect(shouldUseDisplayMode("x")).toBe(false);
  });
});
