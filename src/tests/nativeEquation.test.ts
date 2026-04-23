import { afterEach, describe, expect, it, vi } from "vitest";
import {
  convertEquationRuns,
  convertSelectedTextToNativeEquation,
  formatEquationConversionSummary,
  insertNativeEquationBlock,
  insertNativeEquationTextBox,
} from "../services/nativeEquation";

function jsonResponse(data: unknown, init: ResponseInit = {}): Response {
  return new Response(JSON.stringify(data), {
    status: init.status ?? 200,
    headers: { "Content-Type": "application/json" },
  });
}

describe("native equation helper client", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("calls convert-selection through the helper", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 1, message: "done" }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await convertSelectedTextToNativeEquation("\\delta");

    expect(result.mode).toBe("native");
    expect(fetchMock).toHaveBeenNthCalledWith(1, "/native-helper/health", undefined);
    expect(fetchMock.mock.calls[1][0]).toBe("/native-helper/equation/convert-selection");
  });

  it("turns Load failed into a helper-specific diagnostic", async () => {
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new TypeError("Load failed")));

    await expect(
      convertSelectedTextToNativeEquation("x^2"),
    ).rejects.toThrow("任务窗格无法访问本地公式 helper");
  });

  it("marks legacy insert endpoints as deprecated", async () => {
    await expect(
      insertNativeEquationBlock({
        latex: "x^2",
        box: { left: 80, top: 80, width: 360, height: 60 },
      }),
    ).rejects.toThrow("已弃用");
    await expect(
      insertNativeEquationTextBox({
        text: "x",
        equations: [{ start: 0, length: 1, latex: "x" }],
        box: { left: 80, top: 80, width: 360, height: 60 },
      }),
    ).rejects.toThrow("已弃用");
  });

  it("converts equation runs from back to front and updates shape ids", async () => {
    const selectCalls: Array<{ shapeId: string; start: number; length: number }> = [];
    const convertCalls: string[] = [];
    const selectedIds = [["shape-2"], ["shape-2"]];

    const result = await convertEquationRuns(
      "shape-1",
      [
        { start: 3, length: 6, latex: "\\delta" },
        { start: 18, length: 7, latex: "\\lambda" },
      ],
      {
        selectRange: async (shapeId, start, length) => {
          selectCalls.push({ shapeId, start, length });
        },
        getSelectedIds: async () => selectedIds.shift() ?? ["shape-2"],
        convertSelection: async (latex) => {
          convertCalls.push(latex);
          return { ok: true, mode: "native", message: `done:${latex}` };
        },
      },
    );

    expect(selectCalls).toEqual([
      { shapeId: "shape-1", start: 18, length: 7 },
      { shapeId: "shape-2", start: 3, length: 6 },
    ]);
    expect(convertCalls).toEqual(["\\lambda", "\\delta"]);
    expect(result.shapeId).toBe("shape-2");
    expect(result.nativeCount).toBe(2);
    expect(result.fallbackCount).toBe(0);
  });

  it("returns remaining equations when conversion stops midway", async () => {
    const result = await convertEquationRuns(
      "shape-1",
      [
        { start: 3, length: 6, latex: "\\delta" },
        { start: 18, length: 7, latex: "\\lambda" },
      ],
      {
        selectRange: async () => undefined,
        getSelectedIds: async () => ["shape-1"],
        convertSelection: async (latex) => {
          if (latex === "\\delta") {
            throw new Error("helper unavailable");
          }
          return { ok: true, mode: "native", message: "done" };
        },
      },
    );

    expect(result.nativeCount).toBe(1);
    expect(result.fallbackCount).toBe(1);
    expect(result.remainingEquations.map((equation) => equation.latex)).toEqual(["\\delta"]);
  });

  it("keeps the legacy formatter readable for remaining fallback summaries", () => {
    expect(formatEquationConversionSummary({ shapeId: "shape-1", nativeCount: 1, fallbackCount: 0, messages: [], remainingEquations: [] })).toBe("原生公式成功 1 个，公式降级 0 个");
    expect(
      formatEquationConversionSummary({ shapeId: "shape-1", nativeCount: 0, fallbackCount: 1, messages: ["失败"], remainingEquations: [{ start: 0, length: 1, latex: "x" }] }),
    ).toContain("失败");
  });
});
