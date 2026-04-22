import { describe, expect, it, vi } from "vitest";
import { convertEquationRuns, formatEquationConversionSummary } from "../services/nativeEquation";

describe("native equation conversion", () => {
  it("converts equation runs in reverse order", async () => {
    const calls: string[] = [];
    const result = await convertEquationRuns(
      "shape-1",
      [
        { start: 3, length: 6, latex: "\\delta" },
        { start: 18, length: 7, latex: "\\lambda" },
      ],
      {
        selectRange: vi.fn(async (_shapeId, start) => {
          calls.push(`select:${start}`);
        }),
        convertSelection: vi.fn(async (latex) => {
          calls.push(`convert:${latex}`);
          return { ok: true, mode: "native" as const, message: "ok" };
        }),
      },
    );

    expect(calls).toEqual(["select:18", "convert:\\lambda", "select:3", "convert:\\delta"]);
    expect(result.nativeCount).toBe(2);
    expect(result.fallbackCount).toBe(0);
  });

  it("keeps inline equations as text when helper conversion fails", async () => {
    const result = await convertEquationRuns(
      "shape-1",
      [{ start: 3, length: 6, latex: "\\delta" }],
      {
        selectRange: vi.fn(async () => undefined),
        convertSelection: vi.fn(async () => {
          throw new Error("helper unavailable");
        }),
      },
    );

    expect(result.nativeCount).toBe(0);
    expect(result.fallbackCount).toBe(1);
    expect(formatEquationConversionSummary(result)).toContain("公式降级 1 个");
    expect(formatEquationConversionSummary(result)).toContain("helper unavailable");
  });
});
