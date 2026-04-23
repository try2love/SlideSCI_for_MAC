import { beforeEach, describe, expect, it } from "vitest";
import { getLatexForShape, loadClipboardState, loadSettings, resolveLatexSource, saveClipboardState, saveLatexForShape } from "../services/settings";

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
        borderWeight: 2,
      },
    });

    expect(loadClipboardState().style?.text?.bold).toBe(true);
    expect(loadClipboardState().style?.fillColor).toBe("#ffffff");
    expect(loadClipboardState().style?.borderWeight).toBe(2);
  });

  it("stores latex source by shape id", () => {
    saveLatexForShape("shape-1", "\\frac{a}{b}");
    expect(getLatexForShape("shape-1")).toBe("\\frac{a}{b}");
  });

  it("resolves latex metadata by tags, alt text, then local storage", () => {
    expect(resolveLatexSource({ shapeId: "shape-1", tagLatex: "x_tag", altTextDescription: "x_alt" }, { "shape-1": "x_local" })).toBe("x_tag");
    expect(resolveLatexSource({ shapeId: "shape-1", altTextDescription: "x_alt" }, { "shape-1": "x_local" })).toBe("x_alt");
    expect(resolveLatexSource({ shapeId: "office-id", shapeName: "helper-name" }, { "helper-name": "x_name", "office-id": "x_local" })).toBe("x_name");
    expect(resolveLatexSource({ shapeId: "shape-1" }, { "shape-1": "x_local" })).toBe("x_local");
  });

  it("keeps position and size clipboard state independent", () => {
    saveClipboardState({ centers: [{ left: 1, top: 2 }] });
    saveClipboardState({ ...loadClipboardState(), width: 10, height: 20 });
    const state = loadClipboardState();
    expect(state.centers).toEqual([{ left: 1, top: 2 }]);
    expect(state.width).toBe(10);
    expect(state.height).toBe(20);
  });

  it("defaults equation image fallback to disabled for old settings", () => {
    localStorage.setItem("slidesci_for_mac:settings", JSON.stringify({ codeLanguage: "python" }));
    const settings = loadSettings();
    expect(settings.codeLanguage).toBe("python");
    expect(settings.allowEquationImageFallback).toBe(false);
  });
});
