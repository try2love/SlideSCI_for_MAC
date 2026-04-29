import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { promisify } from "node:util";
import { afterEach, describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);
const HAS_OSACOMPILE = existsSync("/usr/bin/osacompile");
const HELPER_SOURCE_PATH = resolve(process.cwd(), "scripts/native-equation-helper.mjs");

async function expectAppleScriptToCompile(script: string, name: string): Promise<void> {
  const dir = await mkdtemp(join(tmpdir(), "slidesci-helper-test-"));
  const scriptPath = join(dir, `${name}.applescript`);
  const compiledPath = join(dir, `${name}.scpt`);
  try {
    await writeFile(scriptPath, script, "utf8");
    await execFileAsync("osacompile", ["-o", compiledPath, scriptPath]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

describe("native equation helper script exports", () => {
  afterEach(() => {
    delete process.env.SLIDESCI_NATIVE_HELPER_PORT;
  });

  it("exposes pure helper metadata without VBA automation", async () => {
    // @ts-expect-error The helper is a Node .mjs script outside the TS source tree.
    const helper = await import("../../scripts/native-equation-helper.mjs");
    const helperSource = await readFile(HELPER_SOURCE_PATH, "utf8");

    const convertScript = helper.buildConvertSelectionScript();
    const probeScript = helper.buildGuiAutomationProbeScript();
    const syntaxProbeScript = helper.buildSyntaxProbeScript();
    const shapeRangesScript = helper.buildConvertShapeRangesScript({
      placeholders: [{ start: 3, length: 6, latex: "\\delta", unicodeMath: "δ" }],
      strategyOrder: ["equation-insert"],
    });
    expect(helper.SCRIPT_EXECUTION_MODE).toBe("temp-file");
    expect(helper.HELPER_BUILD_ID).toMatch(/^[0-9a-f]{12}$/);
    expect(convertScript).not.toContain("do Visual Basic");
    expect(probeScript).not.toContain("do Visual Basic");
    expect(syntaxProbeScript).not.toContain("do Visual Basic");
    expect(shapeRangesScript).not.toContain("do Visual Basic");
    expect(convertScript).toContain("System Events");
    expect(probeScript).not.toContain("UI elements enabled");
    expect(convertScript).not.toContain("UI elements enabled");
    expect(shapeRangesScript).not.toContain("UI elements enabled");
    expect(convertScript).toContain("Insert");
    expect(convertScript).toContain("Equation");
    expect(convertScript).not.toContain("do shell script");
    expect(convertScript).toContain("on eqInsertHandler(processName, firstChoices, secondChoices)");
    expect(helperSource).not.toContain("--equation-shortcut");
    expect(convertScript).toContain("on replaceRangeAndConvert(processName, startIndex, lengthValue, latexText)");
    expect(convertScript).not.toContain("LaTeX 转数学公式");
    expect(convertScript).not.toContain("Professional");
    expect(convertScript).not.toContain("Equation Options");
    expect(convertScript).not.toContain("公式选项");
    expect(convertScript).not.toContain("on isPressableElement(uiElement)");
    expect(convertScript).not.toContain("on tryLatexRibbonConvert(");
    expect(convertScript).not.toContain("on tryProfessionalLayout(");
    expect(convertScript).not.toContain("“");
    expect(convertScript).not.toContain("”");
    expect(convertScript).toContain("return \"equation-insert\"");
    expect(convertScript).not.toContain('keystroke "=" using {option down}');
    expect(convertScript).not.toContain('keystroke "=" using {control down}');
    expect(convertScript).toContain('error (prefixText & errMsg) number errNum');
    expect(convertScript).toContain('my raiseScriptError("无法通过辅助功能选择文本范围：", errMsg, errNum)');
    expect(convertScript).not.toContain("focusCycleAttempts");
    expect(convertScript).not.toContain("AXSelectedTextRange=no:");
    expect(convertScript).not.toContain("encodeConversionResult");
    expect(convertScript).not.toMatch(/error\s+"[^"]*"\s*&\s*.*number\s+errNum/);
    expect(helperSource).not.toContain("do Visual Basic");
    expect(helperSource).not.toContain("ExecuteMso");
    expect(helperSource).not.toContain("EquationProfessional");
    expect(helperSource).not.toContain("buildPowerPointAutomationProbeScript");
    expect(helperSource).not.toContain("buildConvertSelectionPowerPointScript");
    expect(helperSource).not.toContain("buildConvertShapeRangesPowerPointScript");
    expect(helperSource).not.toContain("runNativeCompanionProbe");
    expect(helperSource).not.toContain("runNativeCompanionConvertSelection");
    expect(helperSource).not.toContain("runNativeCompanionConvertShapeRanges");
    expect(helperSource).not.toContain("powerpoint-probe");
    expect(syntaxProbeScript).toContain("return \"equation-insert\"");
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

  it.skipIf(!HAS_OSACOMPILE)("compiles the syntax probe script with osacompile when available", async () => {
    // @ts-expect-error The helper is a Node .mjs script outside the TS source tree.
    const helper = await import("../../scripts/native-equation-helper.mjs");

    await expectAppleScriptToCompile(helper.buildSyntaxProbeScript(), "syntax-probe");
  });

  it.skipIf(!HAS_OSACOMPILE)("compiles selection conversion scripts containing real latex", async () => {
    // @ts-expect-error The helper is a Node .mjs script outside the TS source tree.
    const helper = await import("../../scripts/native-equation-helper.mjs");

    await expectAppleScriptToCompile(
      helper.buildConvertSelectionScript({
        latex: "\\boldsymbol{\\theta}_u(t)\\in\\mathbb{R}^{d_e}",
        unicodeMath: "θ_u(t)∈ℝ^(d_e)",
        strategyOrder: ["equation-insert"],
      }),
      "selection-latex",
    );
  });

  it.skipIf(!HAS_OSACOMPILE)("compiles shape-range conversion scripts with placeholder replacement", async () => {
    // @ts-expect-error The helper is a Node .mjs script outside the TS source tree.
    const helper = await import("../../scripts/native-equation-helper.mjs");

    await expectAppleScriptToCompile(
      helper.buildConvertShapeRangesScript({
        placeholders: [{ start: 3, length: 6, latex: "\\delta", unicodeMath: "δ" }],
        strategyOrder: ["equation-insert"],
      }),
      "shape-ranges-latex",
    );
  });
});
