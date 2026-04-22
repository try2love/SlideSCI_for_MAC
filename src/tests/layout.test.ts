import { describe, expect, it } from "vitest";
import { arrangeShapes } from "../lib/layout";
import type { SlideShape } from "../lib/types";

const shapes: SlideShape[] = [
  { id: "b", left: 120, top: 10, width: 40, height: 20 },
  { id: "a", left: 10, top: 12, width: 20, height: 20 },
  { id: "c", left: 10, top: 70, width: 30, height: 30 },
  { id: "d", left: 90, top: 72, width: 50, height: 20 },
];

describe("arrangeShapes", () => {
  it("sorts by visual rows and columns for column max width layout", () => {
    const result = arrangeShapes(shapes, {
      colNum: 2,
      colSpace: 10,
      rowSpace: 5,
      mode: "columnMaxWidth",
      sortMode: "position",
    });

    expect(result.map((item) => item.id)).toEqual(["a", "b", "c", "d"]);
    expect(result[0]).toMatchObject({ left: 10, top: 12 });
    expect(result[1]).toMatchObject({ left: 50, top: 12 });
    expect(result[2].top).toBe(37);
  });

  it("uses the first selected height for uniform height layout", () => {
    const result = arrangeShapes(shapes, {
      colNum: 4,
      colSpace: 0,
      rowSpace: 0,
      mode: "uniformHeight",
      sortMode: "selection",
    });

    expect(result.every((item) => item.height === shapes[0].height)).toBe(true);
  });

  it("places waterfall items into the shortest column", () => {
    const result = arrangeShapes(shapes, {
      colNum: 2,
      colSpace: 10,
      rowSpace: 5,
      mode: "waterfall",
      sortMode: "selection",
      customWidth: 20,
    });

    expect(result[0]).toMatchObject({ left: 120, top: 10, width: 20 });
    expect(result[1]).toMatchObject({ left: 150, top: 10, width: 20 });
    expect(result[2].top).toBeGreaterThanOrEqual(25);
  });
});
