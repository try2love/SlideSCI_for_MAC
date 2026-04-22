import { beforeEach, describe, expect, it, vi } from "vitest";
import { addLatexImage, addTableWithFallback } from "../services/powerpoint";

function shape(id: string) {
  return {
    id,
    fill: { setImage: vi.fn(), setSolidColor: vi.fn() },
    lineFormat: {},
    tags: { add: vi.fn() },
    textFrame: { textRange: { font: {}, paragraphFormat: {} } },
    load: vi.fn(),
  };
}

function installPowerPointMock(shapes: Record<string, unknown>) {
  const context = {
    sync: vi.fn(async () => undefined),
    presentation: {
      getSelectedSlides: () => ({
        items: [{ shapes }],
        load: vi.fn(),
      }),
    },
  };
  Object.defineProperty(globalThis, "PowerPoint", {
    configurable: true,
    value: {
      run: async (callback: (context: any) => Promise<unknown>) => callback(context),
    },
  });
  return context;
}

describe("PowerPoint fallback helpers", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("uses rectangle image fill when addPicture is unavailable", async () => {
    const rectangle = shape("shape-fill");
    installPowerPointMock({
      addGeometricShape: vi.fn(() => rectangle),
    });

    const result = await addLatexImage("png-base64", { left: 1, top: 2, width: 3, height: 4 }, "x^2");

    expect(result.mode).toBe("shapeFill");
    expect(result.id).toBe("shape-fill");
    expect(rectangle.fill.setImage).toHaveBeenCalledWith("png-base64");
    expect(result.warning).toBeUndefined();
  });

  it("falls back to a latex source text box when image fill is unavailable", async () => {
    const textBox = shape("latex-text");
    installPowerPointMock({
      addGeometricShape: vi.fn(() => ({ ...shape("bad-fill"), fill: {} })),
      addTextBox: vi.fn(() => textBox),
    });

    const result = await addLatexImage("png-base64", { left: 1, top: 2, width: 3, height: 4 }, "\\frac{a}{b}");

    expect(result.mode).toBe("textFallback");
    expect(result.id).toBe("latex-text");
    expect(result.warning).toContain("fill.setImage 失败");
  });

  it("falls back to an editable text-box grid when native table insertion fails", async () => {
    const createdCells = [shape("c1"), shape("c2"), shape("c3"), shape("c4")];
    const addTextBox = vi.fn(() => createdCells.shift());
    installPowerPointMock({
      addTable: vi.fn(() => {
        throw new Error("table api failed");
      }),
      addTextBox,
    });

    const result = await addTableWithFallback(
      [
        ["A", "B"],
        ["1", "2"],
      ],
      { left: 10, top: 20, width: 200, height: 80 },
    );

    expect(result.mode).toBe("textGrid");
    expect(result.ids).toEqual(["c1", "c2", "c3", "c4"]);
    expect(addTextBox).toHaveBeenCalledTimes(4);
    expect(addTextBox).toHaveBeenNthCalledWith(1, "A", { left: 10, top: 20, width: 100, height: 40 });
    expect(addTextBox).toHaveBeenNthCalledWith(4, "2", { left: 110, top: 60, width: 100, height: 40 });
    expect(result.warning).toContain("已降级为文本框网格");
  });

  it("creates native tables with minimal addTable options", async () => {
    const table = shape("native-table");
    const addTable = vi.fn(() => table);
    installPowerPointMock({ addTable });

    const result = await addTableWithFallback(
      [
        ["A", "B"],
        ["1", "2"],
      ],
      { left: 10.2, top: 20.7, width: 200.4, height: 80.4 },
    );

    expect(result.mode).toBe("nativeTable");
    expect(addTable).toHaveBeenCalledWith(2, 2, {
      left: 10,
      top: 21,
      width: 200,
      height: 80,
      values: [
        ["A", "B"],
        ["1", "2"],
      ],
    });
    const firstCall = addTable.mock.calls[0] as unknown as [number, number, Record<string, unknown>];
    const options = firstCall[2];
    expect(options).not.toHaveProperty("uniformCellProperties");
    expect(options).not.toHaveProperty("borders");
    expect(options).not.toHaveProperty("margins");
  });
});
