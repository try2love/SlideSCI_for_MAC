import { describe, expect, it } from "vitest";
import { formatLabel, generateLabels } from "../lib/labels";

describe("labels", () => {
  it("formats supported templates", () => {
    expect(formatLabel("A", 1)).toBe("A");
    expect(formatLabel("a)", 2)).toBe("b)");
    expect(formatLabel("(A)", 3)).toBe("(C)");
    expect(formatLabel("1)", 4)).toBe("4)");
    expect(formatLabel("①", 2)).toBe("②");
    expect(formatLabel("一)", 3)).toBe("三)");
  });

  it("generates labels in visual position order", () => {
    const result = generateLabels(
      [
        { id: "2", left: 100, top: 10, width: 20, height: 20 },
        { id: "1", left: 10, top: 10, width: 20, height: 20 },
      ],
      "A",
      1,
    );

    expect(result.map((item) => item.shape.id)).toEqual(["1", "2"]);
    expect(result.map((item) => item.label)).toEqual(["A", "B"]);
  });
});
