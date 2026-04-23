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
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 1, strategyUsed: "latex-ribbon", message: "done" }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await nativeEquation.convertSelectedTextToNativeEquation("\\delta");

    expect(result.mode).toBe("native");
    expect(result.strategyUsed).toBe("latex-ribbon");
    expect(fetchMock).toHaveBeenNthCalledWith(1, "/native-helper/health", undefined);
    expect(fetchMock.mock.calls[1][0]).toBe("/native-helper/equation/convert-selection");
    const request = JSON.parse(String(fetchMock.mock.calls[1][1]?.body));
    expect(request.strategyOrder).toEqual(["latex-ribbon"]);
    expect(request.unicodeMath).toBeTruthy();
  });

  it("calls convert-shape-ranges through the helper", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 2, strategyUsed: "latex-ribbon", message: "done" }));
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
      strategyOrder: ["latex-ribbon"],
    });

    expect(result.nativeCount).toBe(2);
    expect(fetchMock.mock.calls[1][0]).toBe("/native-helper/equation/convert-shape-ranges");
    const request = JSON.parse(String(fetchMock.mock.calls[1][1]?.body));
    expect(request.strategyOrder).toEqual(["latex-ribbon"]);
  });

  it("turns Load failed into a helper-specific diagnostic", async () => {
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new TypeError("Load failed")));

    await expect(
      nativeEquation.convertSelectedTextToNativeEquation("x^2"),
    ).rejects.toThrow("任务窗格无法访问本地公式 helper");
  });

  it("surfaces helper build id when health reports syntax failure", () => {
    expect(() =>
      nativeEquation.ensureNativeEquationAvailable({
        ok: true,
        helperBuildId: "abc123def456",
        scriptExecutionMode: "temp-file",
        equationScriptSyntaxOk: false,
        nativeEquationAvailable: false,
        message: "AppleScript 编译失败",
      }),
    ).toThrow("AppleScript 编译失败 [helper build abc123def456]");
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
    expect(request.strategyOrder).toEqual(["latex-ribbon"]);
    expect(request.placeholders[0].unicodeMath).toBeTruthy();
    expect(request.workingText).not.toContain("\\delta");
    expect(request.workingText).not.toContain("\\lambda");
  });

  it("uses helper GUI conversion for legacy inline equations without text-range selection", async () => {
    vi.spyOn(powerpointService, "getPowerPointHostCapabilities").mockResolvedValue({
      textRangeSelection: false,
      nativeTable: false,
      experimentalNativeTable: false,
    });
    const addRichTextBox = vi.spyOn(powerpointService, "addRichTextBox").mockResolvedValue("shape-1");
    const selectShapes = vi.spyOn(powerpointService, "selectShapes").mockResolvedValue();
    const selectTextRange = vi.spyOn(powerpointService, "selectTextRange").mockResolvedValue();
    vi.spyOn(powerpointService, "getSelectedShapeIds").mockResolvedValue(["shape-2"]);

    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 2, strategyUsed: "latex-ribbon", message: "done" }));
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
    expect(request.strategyOrder).toEqual(["latex-ribbon"]);
    expect(result.id).toBe("shape-2");
    expect(result.nativeCount).toBe(2);
    expect(result.strategyUsed).toBe("latex-ribbon");
  });

  it("uses helper GUI conversion for inline equations even when text-range selection is available", async () => {
    vi.spyOn(powerpointService, "getPowerPointHostCapabilities").mockResolvedValue({
      textRangeSelection: true,
      nativeTable: false,
      experimentalNativeTable: false,
    });
    const addRichTextBox = vi.spyOn(powerpointService, "addRichTextBox").mockResolvedValue("shape-1");
    const selectShapes = vi.spyOn(powerpointService, "selectShapes").mockResolvedValue();
    vi.spyOn(powerpointService, "getSelectedShapeIds").mockResolvedValue(["shape-3"]);
    const selectTextRange = vi.spyOn(powerpointService, "selectTextRange").mockResolvedValue();

    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 2, strategyUsed: "latex-ribbon", message: "done" }));
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
    expect(selectTextRange).not.toHaveBeenCalled();
    expect(selectShapes).toHaveBeenCalledWith(["shape-1"]);
    expect(fetchMock.mock.calls[1][0]).toBe("/native-helper/equation/convert-shape-ranges");
    expect(result.id).toBe("shape-3");
    expect(result.nativeCount).toBe(2);
  });

  it("uses helper GUI conversion for legacy block equations without text-range selection", async () => {
    vi.spyOn(powerpointService, "getPowerPointHostCapabilities").mockResolvedValue({
      textRangeSelection: false,
      nativeTable: false,
      experimentalNativeTable: false,
    });
    const addTextBox = vi.spyOn(powerpointService, "addTextBox").mockResolvedValue("shape-1");
    const selectShapes = vi.spyOn(powerpointService, "selectShapes").mockResolvedValue();
    const selectTextRange = vi.spyOn(powerpointService, "selectTextRange").mockResolvedValue();
    vi.spyOn(powerpointService, "getSelectedShapeIds").mockResolvedValue(["shape-3"]);

    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 1, strategyUsed: "latex-ribbon", message: "done" }));
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
    expect(request.strategyOrder).toEqual(["latex-ribbon"]);
    expect(result.id).toBe("shape-3");
    expect(result.nativeCount).toBe(1);
    expect(result.strategyUsed).toBe("latex-ribbon");
  });

  it("falls back to helper GUI conversion when runtime text editing cannot be entered for block equations", async () => {
    vi.spyOn(powerpointService, "getPowerPointHostCapabilities").mockResolvedValue({
      textRangeSelection: true,
      nativeTable: false,
      experimentalNativeTable: false,
    });
    const addTextBox = vi
      .spyOn(powerpointService, "addTextBox")
      .mockResolvedValueOnce("shape-1")
      .mockResolvedValueOnce("shape-2");
    const deleteShapes = vi.spyOn(powerpointService, "deleteShapes").mockResolvedValue();
    const selectShapes = vi.spyOn(powerpointService, "selectShapes").mockResolvedValue();
    const selectTextRange = vi.spyOn(powerpointService, "selectTextRange").mockResolvedValue();
    vi.spyOn(powerpointService, "getSelectedShapeIds").mockResolvedValue(["shape-4"]);

    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: false, helperBuildId: "abc123def456", mode: "unsupported", message: "已检测到 PowerPoint，但界面自动化不可用：无法进入文本编辑状态。请先选中文本框，再重试。最后焦点：role=AXGroup, subrole=, label=Slide canvas, AXSelectedTextRange=no:缺少文本选择 (0) (-2700)" }, { status: 500 }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", nativeCount: 1, strategyUsed: "latex-ribbon", message: "done" }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await nativeEquation.insertNativeEquationBlock({
      latex: "\\frac{a}{b}",
      box: { left: 80, top: 80, width: 360, height: 60 },
    });

    expect(selectTextRange).toHaveBeenCalledWith("shape-1", 0, "\\frac{a}{b}".length);
    expect(deleteShapes).toHaveBeenCalledWith(["shape-1"]);
    expect(addTextBox).toHaveBeenCalledTimes(2);
    expect(selectShapes).toHaveBeenCalledWith(["shape-2"]);
    expect(fetchMock.mock.calls[3][0]).toBe("/native-helper/equation/convert-shape-ranges");
    expect(result.id).toBe("shape-4");
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
          return { ok: true, mode: "native", strategyUsed: "latex-ribbon", message: `done:${latex}` };
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
    expect(result.strategiesUsed).toEqual(["latex-ribbon"]);
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
          return { ok: true, mode: "native", strategyUsed: "unicode-math", message: "done" };
        },
      },
    );

    expect(result.strategy).toBe("officejs-selection");
    expect(result.strategiesUsed).toEqual(["unicode-math"]);
    expect(result.nativeCount).toBe(1);
    expect(result.fallbackCount).toBe(1);
    expect(result.remainingEquations.map((equation) => equation.latex)).toEqual(["\\delta"]);
  });

  it("treats LaTeX convert button failures as native failures for inline equations", async () => {
    vi.spyOn(powerpointService, "getPowerPointHostCapabilities").mockResolvedValue({
      textRangeSelection: true,
      nativeTable: false,
      experimentalNativeTable: false,
    });
    vi.spyOn(powerpointService, "addRichTextBox").mockResolvedValue("shape-1");
    vi.spyOn(powerpointService, "selectShapes").mockResolvedValue();

    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(
        jsonResponse(
          {
            ok: false,
            helperBuildId: "abc123def456",
            mode: "unsupported",
            message: "已检测到 PowerPoint，但自动公式转换失败：latex-convert: 未能自动点击“LaTeX 转数学公式”。",
          },
          { status: 500 },
        ),
      );
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      nativeEquation.insertNativeEquationTextBox({
        text: "其中，\\delta 和 \\lambda。",
        equations: [
          { start: 3, length: 6, latex: "\\delta" },
          { start: 12, length: 7, latex: "\\lambda" },
        ],
        box: { left: 80, top: 80, width: 360, height: 60 },
      }),
    ).rejects.toThrow("latex-convert");
  });

  it("treats helper button failures as native failures for block equations", async () => {
    vi.spyOn(powerpointService, "getPowerPointHostCapabilities").mockResolvedValue({
      textRangeSelection: false,
      nativeTable: false,
      experimentalNativeTable: false,
    });
    vi.spyOn(powerpointService, "addTextBox").mockResolvedValue("shape-1");
    vi.spyOn(powerpointService, "selectShapes").mockResolvedValue();

    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, helperBuildId: "abc123def456", scriptExecutionMode: "temp-file", equationScriptSyntaxOk: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(
        jsonResponse(
          {
            ok: false,
            helperBuildId: "abc123def456",
            mode: "unsupported",
            message: "已检测到 PowerPoint，但自动公式转换失败：professional-layout: 未能自动切换为 Professional。",
          },
          { status: 500 },
        ),
      );
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      nativeEquation.insertNativeEquationBlock({
        latex: "\\frac{a}{b}",
        box: { left: 80, top: 80, width: 360, height: 60 },
      }),
    ).rejects.toThrow("professional-layout");
  });

  it("keeps the legacy formatter readable for remaining fallback summaries", () => {
    expect(
      nativeEquation.formatEquationConversionSummary({
        shapeId: "shape-1",
        strategy: "officejs-selection",
        strategiesUsed: ["latex-ribbon"],
        nativeCount: 1,
        fallbackCount: 0,
        messages: [],
        remainingEquations: [],
      }),
    ).toBe("原生公式成功 1 个（latex-ribbon），公式降级 0 个");

    expect(
      nativeEquation.formatEquationConversionSummary({
        shapeId: "shape-1",
        strategy: "helper-gui",
        strategiesUsed: ["latex-ribbon"],
        nativeCount: 0,
        fallbackCount: 1,
        messages: ["失败"],
        remainingEquations: [{ start: 0, length: 1, latex: "x" }],
      }),
    ).toContain("失败");
  });

  it("recognizes detailed focus diagnostics as a retryable gui shape-range failure", () => {
    expect(
      nativeEquation.shouldRetryWithGuiShapeRange(
        "已检测到 PowerPoint，但界面自动化不可用：无法进入文本编辑状态。请先选中文本框，再重试。最后焦点：role=AXGroup, subrole=, label=Slide canvas, AXSelectedTextRange=no:缺少文本选择 (0) (-2700)",
      ),
    ).toBe(true);
  });
});
