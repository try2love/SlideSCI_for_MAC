import { afterEach, describe, expect, it, vi } from "vitest";
import {
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

  it("inserts inline equation text boxes through the helper without Office.js text range selection", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ ok: true, powerpointRunning: true, nativeEquationAvailable: true, message: "ok" }))
      .mockResolvedValueOnce(jsonResponse({ ok: true, mode: "native", id: "native-text", nativeCount: 2, message: "done" }));
    vi.stubGlobal("fetch", fetchMock);

    const result = await insertNativeEquationTextBox({
      text: "其中，\\delta 是平滑因子，\\lambda 表示噪声衰减系数。",
      equations: [
        { start: 3, length: 6, latex: "\\delta" },
        { start: 18, length: 7, latex: "\\lambda" },
      ],
      box: { left: 80, top: 80, width: 360, height: 60 },
      baseStyle: { fontName: "微软雅黑", fontSize: 14, color: "#000000" },
      runs: [],
    });

    expect(result.id).toBe("native-text");
    expect(fetchMock).toHaveBeenNthCalledWith(1, "/native-helper/health", undefined);
    expect(fetchMock.mock.calls[1][0]).toBe("/native-helper/equation/insert-textbox");
    const body = JSON.parse(String((fetchMock.mock.calls[1][1] as RequestInit).body));
    expect(body.equations.map((equation: { latex: string }) => equation.latex)).toEqual(["\\delta", "\\lambda"]);
  });

  it("turns Load failed into a helper-specific diagnostic", async () => {
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new TypeError("Load failed")));

    await expect(
      insertNativeEquationBlock({
        latex: "x^2",
        box: { left: 80, top: 80, width: 360, height: 60 },
      }),
    ).rejects.toThrow("任务窗格无法访问本地公式 helper");
  });

  it("fails formula modules clearly when the helper is unavailable", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(jsonResponse({ ok: false, nativeEquationAvailable: false, message: "helper unavailable" })),
    );

    await expect(
      insertNativeEquationBlock({
        latex: "x^2",
        box: { left: 80, top: 80, width: 360, height: 60 },
      }),
    ).rejects.toThrow("helper unavailable");
  });

  it("keeps the legacy formatter readable for remaining fallback summaries", () => {
    expect(formatEquationConversionSummary({ nativeCount: 1, fallbackCount: 0, messages: [] })).toBe("原生公式成功 1 个，公式降级 0 个");
    expect(formatEquationConversionSummary({ nativeCount: 0, fallbackCount: 1, messages: ["失败"] })).toContain("失败");
  });
});
