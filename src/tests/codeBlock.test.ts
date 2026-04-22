import { describe, expect, it } from "vitest";
import { getCodeBlockStyle, getCodeHighlightRuns } from "../services/codeBlock";

describe("getCodeHighlightRuns", () => {
  it("creates colored text runs for Prism tokens", () => {
    const runs = getCodeHighlightRuns('def hello():\n  print("hi") # comment\n  return 1', "python", true);
    const tokenTypes = new Set(runs.map((run) => run.style.tokenType));

    expect(tokenTypes.has("keyword")).toBe(true);
    expect(tokenTypes.has("string")).toBe(true);
    expect(tokenTypes.has("comment")).toBe(true);
    expect(runs.some((run) => run.start >= 0 && run.length > 0 && run.style.color)).toBe(true);
  });

  it("supports C, C++, and Java aliases", () => {
    expect(getCodeHighlightRuns("int main() { return 0; }", "c", true).length).toBeGreaterThan(0);
    expect(getCodeHighlightRuns("#include <vector>\nclass A {};", "c++", true).length).toBeGreaterThan(0);
    expect(getCodeHighlightRuns("public class A { static void main(String[] args) {} }", "java", true).length).toBeGreaterThan(0);
  });
});

describe("getCodeBlockStyle", () => {
  it("sets opaque dark and light backgrounds", () => {
    expect(getCodeBlockStyle(true).fillColor).toBe("#1e1e1e");
    expect(getCodeBlockStyle(false).fillColor).toBe("#ffffff");
    expect(getCodeBlockStyle(true).borderColor).toBe("#c8c8c8");
  });
});
