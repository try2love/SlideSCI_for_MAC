import { beforeEach, describe, expect, it } from "vitest";
import { getLatexForShape, loadClipboardState, saveClipboardState, saveLatexForShape } from "../services/settings";

describe("settings persistence helpers", () => {
  beforeEach(() => {
    const store = new Map<string, string>();
    Object.defineProperty(globalThis, "localStorage", {
      value: {
        getItem: (key: string) => store.get(key) ?? null,
        setItem: (key: string, value: string) => store.set(key, value),
        removeItem: (key: string) => store.delete(key),
        clear: () => store.clear(),
      },
      configurable: true,
    });
  });

  it("serializes copied style state", () => {
    saveClipboardState({
      style: {
        text: { fontName: "Arial", fontSize: 12, bold: true },
        fillColor: "#ffffff",
        borderColor: "#000000",
      },
    });

    expect(loadClipboardState().style?.text?.bold).toBe(true);
    expect(loadClipboardState().style?.fillColor).toBe("#ffffff");
  });

  it("stores latex source by shape id", () => {
    saveLatexForShape("shape-1", "\\frac{a}{b}");
    expect(getLatexForShape("shape-1")).toBe("\\frac{a}{b}");
  });

  it("keeps position and size clipboard state independent", () => {
    saveClipboardState({ centers: [{ left: 1, top: 2 }] });
    saveClipboardState({ ...loadClipboardState(), width: 10, height: 20 });
    const state = loadClipboardState();
    expect(state.centers).toEqual([{ left: 1, top: 2 }]);
    expect(state.width).toBe(10);
    expect(state.height).toBe(20);
  });
});
