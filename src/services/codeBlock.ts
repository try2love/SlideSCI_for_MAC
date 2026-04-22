export const CODE_LANGUAGES = ["matlab", "python", "r", "javascript", "html", "css", "csharp", "fortran"];

export function getCodeBlockStyle(darkBackground: boolean) {
  return {
    fontName: "Consolas",
    fontSize: 12,
    color: darkBackground ? "#ffffff" : "#000000",
    fillColor: darkBackground ? "#1e1e1e" : "#ffffff",
    borderColor: "#c8c8c8",
    align: "left" as const,
  };
}
