import { describe, expect, it } from "vitest";

describe("native equation helper script exports", () => {
  it("exposes root status metadata and points unknown APIs to /health", async () => {
    // @ts-expect-error The helper is a Node .mjs script outside the TS source tree.
    const helper = await import("../../scripts/native-equation-helper.mjs");

    const status = await helper.rootStatus();
    expect(status.ok).toBe(true);
    expect(status.helper).toBe("SlideSCI native equation helper");
    expect(status.endpoints).toContain("GET /health");
    expect(status.endpoints).toContain("POST /equation/insert-block");

    const unknown = helper.unknownHelperApiResponse();
    expect(unknown.ok).toBe(false);
    expect(unknown.message).toContain("/health");
  });
});
