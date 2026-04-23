import { describe, expect, it } from "vitest";

describe("native equation helper script exports", () => {
  it("exposes pure helper metadata without VBA automation", async () => {
    // @ts-expect-error The helper is a Node .mjs script outside the TS source tree.
    const helper = await import("../../scripts/native-equation-helper.mjs");

    const convertScript = helper.buildConvertSelectionScript();
    const probeScript = helper.buildGuiAutomationProbeScript();
    expect(convertScript).not.toContain("do Visual Basic");
    expect(probeScript).not.toContain("do Visual Basic");
    expect(convertScript).toContain("System Events");
    expect(helper.HELPER_ENDPOINTS).toContain("POST /equation/convert-selection");

    const unknown = helper.unknownHelperApiResponse();
    expect(unknown.ok).toBe(false);
    expect(unknown.message).toContain("/health");

    const deprecated = helper.deprecatedInsertEndpointResponse();
    expect(deprecated.ok).toBe(false);
    expect(deprecated.message).toContain("已弃用");
  });
});
