import { beforeEach, describe, expect, it, vi } from "vitest";
import { addLatexImage, addTableWithFallback, getPowerPointHostCapabilities } from "../services/powerpoint";

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

function tableShape(id: string) {
  const cells = new Map<string, { text: string }>();
  return {
    ...shape(id),
    getTable: vi.fn(() => ({
      getCellOrNullObject: vi.fn((row: number, column: number) => {
        const key = `${row}:${column}`;
        const cell = cells.get(key) ?? { text: "" };
        cells.set(key, cell);
        return cell;
      }),
    })),
    cells,
  };
}

function installPowerPointMock(shapes: Record<string, unknown>, selectedShapeItems: Array<Record<string, unknown>> = []) {
  const context = {
    sync: vi.fn(async () => undefined),
    presentation: {
      getSelectedSlides: () => ({
        items: [{ shapes }],
        load: vi.fn(),
      }),
      getSelectedShapes: () => ({
        items: selectedShapeItems,
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

  it("emits a generic native-table warning after InvalidArgument failures", async () => {
    const createdCells = [shape("c1"), shape("c2"), shape("c3"), shape("c4")];
    installPowerPointMock({
      addTable: vi.fn(() => {
        throw new Error("InvalidArgument");
      }),
      addTextBox: vi.fn(() => createdCells.shift()),
    });

    const result = await addTableWithFallback(
      [
        ["A", "B"],
        ["1", "2"],
      ],
      { left: 10, top: 20, width: 200, height: 80 },
    );

    expect(result.mode).toBe("textGrid");
    expect(result.warningCode).toBe("nativeTableUnsupported");
    expect(result.warning).toBe("当前 PowerPoint 版本不支持 PowerPoint 原生表格，已使用文本框网格近似显示。");
  });

  it("keeps explicit table position when falling back to values insertion", async () => {
    const valuesTable = shape("values-table");
    const addTable = vi
      .fn()
      .mockReturnValueOnce({
        ...shape("empty-table"),
        getTable: vi.fn(() => {
          throw new Error("cell api failed");
        }),
      })
      .mockReturnValueOnce(valuesTable);
    installPowerPointMock({ addTable });

    const result = await addTableWithFallback(
      [
        ["A", "B"],
        ["1", "2"],
      ],
      { left: 10.2, top: 20.7, width: 200.4, height: 80.4 },
    );

    expect(result.mode).toBe("nativeTable");
    expect(addTable).toHaveBeenCalledTimes(2);
    expect(addTable).toHaveBeenNthCalledWith(2, 2, 2, {
      left: 10,
      top: 21,
      width: 200,
      height: 80,
      values: [
        ["A", "B"],
        ["1", "2"],
      ],
    });
  });

  it("creates native tables by first adding an empty table and then filling cells", async () => {
    const table = tableShape("native-table");
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
    });
    const firstCall = addTable.mock.calls[0] as unknown as [number, number, Record<string, unknown>];
    const options = firstCall[2];
    expect(options).not.toHaveProperty("values");
    expect(options).not.toHaveProperty("uniformCellProperties");
    expect(options).not.toHaveProperty("borders");
    expect(options).not.toHaveProperty("margins");
    expect(table.cells.get("0:0")?.text).toBe("A");
    expect(table.cells.get("1:1")?.text).toBe("2");
  });

  it("treats native tables as unsupported when the host requirement set says so", async () => {
    installPowerPointMock({
      addTable: vi.fn(),
    });
    Object.defineProperty(globalThis, "Office", {
      configurable: true,
      value: {
        context: {
          requirements: {
            isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.6"),
          },
        },
      },
    });

    const capabilities = await getPowerPointHostCapabilities();

    expect(capabilities.textRangeSelection).toBe(true);
    expect(capabilities.nativeTable).toBe(false);
  });

  it("detects text-range selection from the actual host APIs when a shape is selected", async () => {
    const setSelected = vi.fn();
    installPowerPointMock(
      {
        getItem: vi.fn(() => ({
          textFrame: {
            textRange: {
              getSubstring: vi.fn(() => ({
                setSelected,
              })),
            },
          },
        })),
      },
      [{ id: "shape-1" }],
    );
    Object.defineProperty(globalThis, "Office", {
      configurable: true,
      value: {
        context: {
          requirements: {
            isSetSupported: vi.fn(() => false),
          },
        },
      },
    });

    const capabilities = await getPowerPointHostCapabilities();

    expect(capabilities.textRangeSelection).toBe(true);
    expect(capabilities.nativeTable).toBe(false);
  });
});
