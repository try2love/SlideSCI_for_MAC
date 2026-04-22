import { describe, expect, it } from "vitest";
import { getCodeHighlightRuns } from "../services/codeBlock";

describe("getCodeHighlightRuns", () => {
  it("creates colored text runs for Prism tokens", () => {
    const runs = getCodeHighlightRuns('def hello():\n  print("hi") # comment\n  return 1', "python", true);
    const tokenTypes = new Set(runs.map((run) => run.style.tokenType));

    expect(tokenTypes.has("keyword")).toBe(true);
    expect(tokenTypes.has("string")).toBe(true);
    expect(tokenTypes.has("comment")).toBe(true);
    expect(runs.some((run) => run.start >= 0 && run.length > 0 && run.style.color)).toBe(true);
  });
});
