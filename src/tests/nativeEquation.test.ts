import { afterEach, describe, expect, it, vi } from "vitest";
import * as nativeEquation from "../services/nativeEquation";
import * as powerpointService from "../services/powerpoint";

function jsonResponse(data: unknown, init: ResponseInit = {}): Response {
  return new Response(JSON.stringify(data), {
    status: init.status ?? 200,
    headers: { "Content-Type": "application/json" },
  });
}

describe("native equation helper client", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("calls convert-selection through the helper", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 1, message: "done" }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await nativeEquation.convertSelectedTextToNativeEquation("\\delta");

    expect(result.mode).toBe("native");
    expect(fetchMock).toHaveBeenNthCalledWith(1, "/native-helper/health", undefined);
    expect(fetchMock.mock.calls[1][0]).toBe("/native-helper/equation/convert-selection");
  });

  it("calls convert-shape-ranges through the helper", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 2, message: "done" }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await nativeEquation.convertShapeRangesToNativeEquations({
      shapeId: "shape-1",
      originalText: "其中，\\delta 和 \\lambda",
      workingText: "其中，AAAAAA 和 BBBBBBB",
      placeholders: [
        { start: 3, length: 6, latex: "\\delta", token: "AAAAAA" },
        { start: 12, length: 7, latex: "\\lambda", token: "BBBBBBB" },
      ],
      mode: "inline",
    });

    expect(result.nativeCount).toBe(2);
    expect(fetchMock.mock.calls[1][0]).toBe("/native-helper/equation/convert-shape-ranges");
  });

  it("turns Load failed into a helper-specific diagnostic", async () => {
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new TypeError("Load failed")));

    await expect(
      nativeEquation.convertSelectedTextToNativeEquation("x^2"),
    ).rejects.toThrow("任务窗格无法访问本地公式 helper");
  });

  it("builds same-length placeholder requests for legacy GUI conversion", () => {
    const originalText = "其中，\\delta 是平滑因子，\\lambda 表示噪声衰减系数。";
    const request = nativeEquation.buildShapeRangeEquationRequest(
      "shape-1",
      originalText,
      [
        { start: 3, length: 6, latex: "\\delta" },
        { start: 15, length: 7, latex: "\\lambda" },
      ],
      "inline",
    );

    expect(request.originalText).toBe(originalText);
    expect(request.workingText).toHaveLength(originalText.length);
    expect(request.mode).toBe("inline");
    expect(request.placeholders).toHaveLength(2);
    expect(request.placeholders[0].length).toBe(request.placeholders[0].token.length);
    expect(request.placeholders[1].length).toBe(request.placeholders[1].token.length);
    expect(request.workingText).not.toContain("\\delta");
    expect(request.workingText).not.toContain("\\lambda");
  });

  it("uses helper GUI conversion for legacy inline equations without text-range selection", async () => {
    vi.spyOn(powerpointService, "getPowerPointHostCapabilities").mockResolvedValue({
      textRangeSelection: false,
      nativeTable: false,
    });
    const addRichTextBox = vi.spyOn(powerpointService, "addRichTextBox").mockResolvedValue("shape-1");
    const selectShapes = vi.spyOn(powerpointService, "selectShapes").mockResolvedValue();
    const selectTextRange = vi.spyOn(powerpointService, "selectTextRange").mockResolvedValue();
    vi.spyOn(powerpointService, "getSelectedShapeIds").mockResolvedValue(["shape-2"]);

    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 2, message: "done" }));
    vi.stubGlobal("fetch", fetchMock);

    const text = "其中，\\delta 和 \\lambda。";
    const result = await nativeEquation.insertNativeEquationTextBox({
      text,
      equations: [
        { start: 3, length: 6, latex: "\\delta" },
        { start: 12, length: 7, latex: "\\lambda" },
      ],
      box: { left: 80, top: 80, width: 360, height: 60 },
    });

    expect(addRichTextBox).toHaveBeenCalledTimes(1);
    expect(selectShapes).toHaveBeenCalledWith(["shape-1"]);
    expect(selectTextRange).not.toHaveBeenCalled();
    const request = JSON.parse(String(fetchMock.mock.calls[1][1]?.body));
    expect(request.shapeId).toBe("shape-1");
    expect(request.mode).toBe("inline");
    expect(request.workingText).toHaveLength(text.length);
    expect(request.placeholders).toHaveLength(2);
    expect(result.id).toBe("shape-2");
    expect(result.nativeCount).toBe(2);
  });

  it("uses helper GUI conversion for legacy block equations without text-range selection", async () => {
    vi.spyOn(powerpointService, "getPowerPointHostCapabilities").mockResolvedValue({
      textRangeSelection: false,
      nativeTable: false,
    });
    const addTextBox = vi.spyOn(powerpointService, "addTextBox").mockResolvedValue("shape-1");
    const selectShapes = vi.spyOn(powerpointService, "selectShapes").mockResolvedValue();
    const selectTextRange = vi.spyOn(powerpointService, "selectTextRange").mockResolvedValue();
    vi.spyOn(powerpointService, "getSelectedShapeIds").mockResolvedValue(["shape-3"]);

    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 1, message: "done" }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await nativeEquation.insertNativeEquationBlock({
      latex: "\\frac{a}{b}",
      box: { left: 80, top: 80, width: 360, height: 60 },
    });

    expect(addTextBox).toHaveBeenCalledTimes(1);
    expect(selectShapes).toHaveBeenCalledWith(["shape-1"]);
    expect(selectTextRange).not.toHaveBeenCalled();
    const request = JSON.parse(String(fetchMock.mock.calls[1][1]?.body));
    expect(request.shapeId).toBe("shape-1");
    expect(request.mode).toBe("block");
    expect(request.placeholders).toHaveLength(1);
    expect(result.id).toBe("shape-3");
    expect(result.nativeCount).toBe(1);
  });

  it("converts equation runs from back to front and updates shape ids", async () => {
    const selectCalls: Array<{ shapeId: string; start: number; length: number }> = [];
    const convertCalls: string[] = [];
    const selectedIds = [["shape-2"], ["shape-2"]];

    const result = await nativeEquation.convertEquationRuns(
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
    expect(result.strategy).toBe("officejs-selection");
    expect(result.nativeCount).toBe(2);
    expect(result.fallbackCount).toBe(0);
  });

  it("returns remaining equations when conversion stops midway", async () => {
    const result = await nativeEquation.convertEquationRuns(
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

    expect(result.strategy).toBe("officejs-selection");
    expect(result.nativeCount).toBe(1);
    expect(result.fallbackCount).toBe(1);
    expect(result.remainingEquations.map((equation) => equation.latex)).toEqual(["\\delta"]);
  });

  it("keeps the legacy formatter readable for remaining fallback summaries", () => {
    expect(
      nativeEquation.formatEquationConversionSummary({
        shapeId: "shape-1",
        strategy: "officejs-selection",
        nativeCount: 1,
        fallbackCount: 0,
        messages: [],
        remainingEquations: [],
      }),
    ).toBe("原生公式成功 1 个，公式降级 0 个");

    expect(
      nativeEquation.formatEquationConversionSummary({
        shapeId: "shape-1",
        strategy: "helper-gui",
        nativeCount: 0,
        fallbackCount: 1,
        messages: ["失败"],
        remainingEquations: [{ start: 0, length: 1, latex: "x" }],
      }),
    ).toContain("失败");
  });
});
