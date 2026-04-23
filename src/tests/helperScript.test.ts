import { describe, expect, it } from "vitest";

describe("native equation helper script exports", () => {
  it("exposes pure helper metadata without VBA automation", async () => {
    // @ts-expect-error The helper is a Node .mjs script outside the TS source tree.
    const helper = await import("../../scripts/native-equation-helper.mjs");

    const convertScript = helper.buildConvertSelectionScript();
    const probeScript = helper.buildGuiAutomationProbeScript();
    const shapeRangesScript = helper.buildConvertShapeRangesScript({
      placeholders: [{ start: 3, length: 6, latex: "\\delta" }],
    });
    expect(convertScript).not.toContain("do Visual Basic");
    expect(probeScript).not.toContain("do Visual Basic");
    expect(shapeRangesScript).not.toContain("do Visual Basic");
    expect(convertScript).toContain("System Events");
    expect(shapeRangesScript).toContain("AXSelectedTextRange");
    expect(helper.HELPER_ENDPOINTS).toContain("POST /equation/convert-selection");
    expect(helper.HELPER_ENDPOINTS).toContain("POST /equation/convert-shape-ranges");

    const unknown = helper.unknownHelperApiResponse();
    expect(unknown.ok).toBe(false);
    expect(unknown.message).toContain("/health");

    const deprecated = helper.deprecatedInsertEndpointResponse();
    expect(deprecated.ok).toBe(false);
    expect(deprecated.message).toContain("已弃用");
    expect(deprecated.message).toContain("/equation/convert-shape-ranges");
  });
});
