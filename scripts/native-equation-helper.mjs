import http from "node:http";
import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const HOST = "127.0.0.1";
const PORT = Number(process.env.SLIDESCI_NATIVE_HELPER_PORT || 17926);
export const SCRIPT_EXECUTION_MODE = "temp-file";

const POWERPOINT_PROCESS_NAME = "Microsoft PowerPoint";
const MENU_LABELS = {
  insertMenu: ["Insert", "插入"],
  equationItem: ["Equation", "Equation...", "公式"],
  contextualMenus: ["Equation", "公式", "Design", "设计", "Convert", "转换"],
  professionalItem: ["Professional", "Professional...", "专业", "专业型", "转换为专业格式"],
};
const RIBBON_LABELS = {
  convertControls: ["Convert", "转换", "Equation Options", "公式选项"],
  latexItems: ["LaTeX", "LaTeX 转数学公式", "LaTeX to Math", "从 LaTeX", "转换为数学公式"],
  professionalItems: ["Professional", "专业", "专业型", "转换为专业格式"],
};
const DEFAULT_STRATEGY_ORDER = ["latex-ribbon"];

export const HELPER_ENDPOINTS = [
  "GET /",
  "GET /health",
  "POST /equation/convert-selection",
  "POST /equation/convert-shape-ranges",
  "POST /equation/insert-textbox (deprecated)",
  "POST /equation/insert-block (deprecated)",
];

const HELPER_FILE_PATH = fileURLToPath(import.meta.url);
export const HELPER_BUILD_ID = createHash("sha1")
  .update(await readFile(HELPER_FILE_PATH))
  .digest("hex")
  .slice(0, 12);

export function unknownHelperApiResponse() {
  return {
    ok: false,
    message: "未知 helper API。请访问 /health 查看 helper 状态，或访问 / 查看可用 API。",
  };
}

export function deprecatedInsertEndpointResponse() {
  return {
    ok: false,
    mode: "unsupported",
    message: "该 helper 接口已弃用。请改为由 Office.js 插入文本框后，调用 /equation/convert-selection 或 /equation/convert-shape-ranges。",
  };
}

function json(res, status, data) {
  res.writeHead(status, {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json; charset=utf-8",
  });
  res.end(JSON.stringify(data));
}

function execFileAsync(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    execFile(command, args, options, (error, stdout, stderr) => {
      if (error) {
        const wrapped = new Error((stderr || error.message).trim());
        wrapped.cause = error;
        reject(wrapped);
        return;
      }
      resolve({ stdout: stdout.trim(), stderr: stderr.trim() });
    });
  });
}

async function withTempAppleScript(kind, script, callback) {
  const dir = await mkdtemp(join(tmpdir(), `slidesci-${HELPER_BUILD_ID}-${kind}-`));
  const scriptPath = join(dir, `${kind}.applescript`);
  try {
    await writeFile(scriptPath, script, "utf8");
    return await callback({ dir, scriptPath });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

async function compileAppleScriptIfAvailable(scriptPath) {
  const compiledPath = join(dirname(scriptPath), "compiled.scpt");
  try {
    await execFileAsync("osacompile", ["-o", compiledPath, scriptPath], { timeout: 8000 });
    return {
      available: true,
      ok: true,
      message: "AppleScript 编译校验通过。",
    };
  } catch (error) {
    const cause = error?.cause;
    if (cause && typeof cause === "object" && "code" in cause && cause.code === "ENOENT") {
      return {
        available: false,
        ok: undefined,
        message: "未检测到 osacompile，已跳过编译预检。",
      };
    }
    return {
      available: true,
      ok: false,
      message: error instanceof Error ? error.message : String(error),
    };
  } finally {
    await rm(compiledPath, { force: true });
  }
}

async function runOsaScript(kind, script) {
  return withTempAppleScript(kind, script, async ({ scriptPath }) => {
    const syntaxCheck = await compileAppleScriptIfAvailable(scriptPath);
    if (syntaxCheck.ok === false) {
      throw new Error(`AppleScript 编译失败：${syntaxCheck.message}`);
    }
    const { stdout } = await execFileAsync("osascript", [scriptPath], { timeout: 8000 });
    return stdout;
  });
}

async function readBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }
  const raw = Buffer.concat(chunks).toString("utf8");
  return raw ? JSON.parse(raw) : {};
}

function normalizeLatexInput(input) {
  let trimmed = String(input || "").trim();
  if (trimmed.startsWith("$$") && trimmed.endsWith("$$")) {
    trimmed = trimmed.slice(2, -2);
  } else if (trimmed.startsWith("$") && trimmed.endsWith("$")) {
    trimmed = trimmed.slice(1, -1);
  } else if (trimmed.startsWith("\\(") && trimmed.endsWith("\\)")) {
    trimmed = trimmed.slice(2, -2);
  } else if (trimmed.startsWith("\\[") && trimmed.endsWith("\\]")) {
    trimmed = trimmed.slice(2, -2);
  }
  return trimmed.replace(/\r/g, "").trim();
}

function appleScriptString(value) {
  return `"${String(value ?? "").replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r?\n/g, "\\n")}"`;
}

function appleScriptList(values) {
  return `{${values.map((value) => appleScriptString(value)).join(", ")}}`;
}

export function buildGuiAutomationProbeScript() {
  return `
set processName to ${appleScriptString(POWERPOINT_PROCESS_NAME)}
tell application "System Events"
  set automationEnabled to UI elements enabled
  set powerpointRunning to exists process processName
  if automationEnabled is false then
    return "ui-disabled"
  end if
  if powerpointRunning is false then
    return "powerpoint-not-running"
  end if
  tell process processName
    set frontmost to true
    return "ok"
  end tell
end tell
`.trim();
}

export function buildSyntaxProbeScript() {
  return buildConvertSelectionScript({
    latex: "\\delta",
    strategyOrder: DEFAULT_STRATEGY_ORDER,
  });
}

function escapeAppleScriptInteger(value, fallback = 0) {
  const integer = Number.parseInt(String(value ?? fallback), 10);
  return Number.isFinite(integer) ? String(integer) : String(fallback);
}

function buildEquationAutomationScript(commands) {
  return `
set processName to ${appleScriptString(POWERPOINT_PROCESS_NAME)}
set insertMenuCandidates to ${appleScriptList(MENU_LABELS.insertMenu)}
set equationItemCandidates to ${appleScriptList(MENU_LABELS.equationItem)}
set contextualMenuCandidates to ${appleScriptList(MENU_LABELS.contextualMenus)}
set professionalItemCandidates to ${appleScriptList(MENU_LABELS.professionalItem)}
set convertControlCandidates to ${appleScriptList(RIBBON_LABELS.convertControls)}
set latexRibbonCandidates to ${appleScriptList(RIBBON_LABELS.latexItems)}
set professionalRibbonCandidates to ${appleScriptList(RIBBON_LABELS.professionalItems)}
set latexClickFailed to false
set professionalClickFailed to false

on firstExistingMenuBarItem(processName, candidates)
  tell application "System Events"
    tell process processName
      repeat with candidateName in candidates
        if exists menu bar item (contents of candidateName) of menu bar 1 then
          return contents of candidateName
        end if
      end repeat
    end tell
  end tell
  error "未找到可用菜单栏项目。"
end firstExistingMenuBarItem

on clickFirstMatchingMenuItem(processName, menuCandidates, itemCandidates)
  tell application "System Events"
    tell process processName
      set menuBarName to my firstExistingMenuBarItem(processName, menuCandidates)
      repeat with itemName in itemCandidates
        try
          click menu item (contents of itemName) of menu 1 of menu bar item menuBarName of menu bar 1
          return contents of itemName
        end try
      end repeat
    end tell
  end tell
  error "未找到菜单命令。"
end clickFirstMatchingMenuItem

on elementLabel(uiElement)
  set labelText to ""
  tell application "System Events"
    try
      set labelText to name of uiElement
    end try
    if labelText is missing value or labelText is "" then
      try
        set labelText to value of attribute "AXTitle" of uiElement
      end try
    end if
    if labelText is missing value or labelText is "" then
      try
        set labelText to value of attribute "AXDescription" of uiElement
      end try
    end if
    if labelText is missing value or labelText is "" then
      try
        set labelText to description of uiElement
      end try
    end if
  end tell
  if labelText is missing value then
    return ""
  end if
  return labelText as text
end elementLabel

on elementAttributeText(uiElement, attributeName)
  set attributeText to ""
  tell application "System Events"
    try
      set rawValue to value of attribute attributeName of uiElement
      if rawValue is missing value then
        return ""
      end if
      set attributeText to rawValue as text
    end try
  end tell
  return attributeText
end elementAttributeText

on matchesCandidates(labelText, candidates)
  if labelText is "" then
    return false
  end if
  ignoring case
    repeat with candidateName in candidates
      if labelText contains (contents of candidateName) then
        return true
      end if
    end repeat
  end ignoring
  return false
end matchesCandidates

on pressControl(uiElement)
  tell application "System Events"
    try
      perform action "AXPress" of uiElement
    on error
      click uiElement
    end try
  end tell
end pressControl

on pressFirstMatchingProcessControl(processName, candidates)
  tell application "System Events"
    tell process processName
      set controlCandidates to entire contents
      repeat with uiElement in controlCandidates
        try
          set labelText to my elementLabel(uiElement)
          if my matchesCandidates(labelText, candidates) then
            my pressControl(uiElement)
            return labelText
          end if
        end try
      end repeat
    end tell
  end tell
  error "未找到匹配的界面控件。"
end pressFirstMatchingProcessControl

on openControlAndChoose(processName, controlCandidates, itemCandidates)
  my pressFirstMatchingProcessControl(processName, controlCandidates)
  delay 0.18
  return my pressFirstMatchingProcessControl(processName, itemCandidates)
end openControlAndChoose

on raiseScriptError(prefixText, errMsg, errNum)
  error (prefixText & errMsg) number errNum
end raiseScriptError

on ensureFocusedEditableElement(processName)
  tell application "${POWERPOINT_PROCESS_NAME}"
    activate
  end tell
  delay 0.12
  tell application "System Events"
    if UI elements enabled is false then
      error "macOS 未授予辅助功能权限，helper 无法驱动 PowerPoint 界面。"
    end if
    tell process processName
      set frontmost to true
      try
        set focusedElement to value of attribute "AXFocusedUIElement"
        value of attribute "AXSelectedTextRange" of focusedElement
        return focusedElement
      end try
      key code 36
      delay 0.12
      set focusedElement to value of attribute "AXFocusedUIElement"
      try
        value of attribute "AXSelectedTextRange" of focusedElement
      on error
        error "无法进入文本编辑状态。请先选中文本框，再重试。"
      end try
      return focusedElement
    end tell
  end tell
end ensureFocusedEditableElement

on selectCharacterRange(focusedElement, startIndex, lengthValue)
  tell application "System Events"
    try
      set value of attribute "AXSelectedTextRange" of focusedElement to {startIndex, lengthValue}
    on error errMsg number errNum
      my raiseScriptError("无法通过辅助功能选择文本范围：", errMsg, errNum)
    end try
  end tell
end selectCharacterRange

on pasteText(processName, textValue)
  set the clipboard to textValue
  tell application "System Events"
    tell process processName
      keystroke "v" using {command down}
    end tell
  end tell
end pasteText

on replaceRangeText(processName, startIndex, lengthValue, nextText)
  set focusedElement to my ensureFocusedEditableElement(processName)
  my selectCharacterRange(focusedElement, startIndex, lengthValue)
  delay 0.05
  my pasteText(processName, nextText)
  delay 0.08
end replaceRangeText

on reselectRange(processName, startIndex, lengthValue)
  set focusedElement to my ensureFocusedEditableElement(processName)
  my selectCharacterRange(focusedElement, startIndex, lengthValue)
  delay 0.05
end reselectRange

on tryInsertEquation(processName, menuCandidates, itemCandidates)
  try
    return my clickFirstMatchingMenuItem(processName, menuCandidates, itemCandidates)
  on error
    tell application "System Events"
      tell process processName
        keystroke "=" using {option down}
      end tell
    end tell
    return "shortcut"
  end try
end tryInsertEquation

on tryLatexRibbonConvert(processName)
  global latexClickFailed
  try
    return my pressFirstMatchingProcessControl(processName, latexRibbonCandidates)
  on error
    try
      return my openControlAndChoose(processName, convertControlCandidates, latexRibbonCandidates)
    on error
      set latexClickFailed to true
      return "failed"
    end try
  end try
end tryLatexRibbonConvert

on tryProfessionalLayout(processName)
  global professionalClickFailed
  try
    return my pressFirstMatchingProcessControl(processName, professionalRibbonCandidates)
  on error
    try
      return my openControlAndChoose(processName, convertControlCandidates, professionalRibbonCandidates)
    on error
      try
        return my clickFirstMatchingMenuItem(processName, contextualMenuCandidates, professionalItemCandidates)
      on error
        set professionalClickFailed to true
        return "failed"
      end try
    end try
  end try
end tryProfessionalLayout

on encodeConversionResult()
  global latexClickFailed
  global professionalClickFailed
  return "latexFailed=" & (latexClickFailed as text) & ";professionalFailed=" & (professionalClickFailed as text)
end encodeConversionResult

on triggerEquationForRange(processName)
  my tryInsertEquation(processName, insertMenuCandidates, equationItemCandidates)
  delay 0.24
  my tryLatexRibbonConvert(processName)
  delay 0.18
  my tryProfessionalLayout(processName)
end triggerEquationForRange

on replaceRangeAndConvert(processName, startIndex, lengthValue, latexText)
  my replaceRangeText(processName, startIndex, lengthValue, latexText)
  my reselectRange(processName, startIndex, (length of latexText))
  my triggerEquationForRange(processName)
  delay 0.18
end replaceRangeAndConvert

${commands}
`.trim();
}

export function buildConvertSelectionScript(payload = {}) {
  normalizeLatexInput(payload.latex);
  return buildEquationAutomationScript(
    `
tell application "${POWERPOINT_PROCESS_NAME}"
  activate
end tell
delay 0.12
tell application "System Events"
  if UI elements enabled is false then
    error "macOS 未授予辅助功能权限，helper 无法驱动 PowerPoint 界面。"
  end if
  if not (exists process processName) then
    error "未检测到 Microsoft PowerPoint。"
  end if
  tell process processName
    set frontmost to true
  end tell
end tell
delay 0.10
my tryInsertEquation(processName, insertMenuCandidates, equationItemCandidates)
delay 0.24
my tryLatexRibbonConvert(processName)
delay 0.18
my tryProfessionalLayout(processName)
return my encodeConversionResult()
`,
  );
}

export function buildConvertShapeRangesScript(payload = {}) {
  const placeholders = Array.isArray(payload.placeholders) ? payload.placeholders : [];
  const commands = placeholders
    .slice()
    .sort((a, b) => Number(b.start) - Number(a.start))
    .map((placeholder) => {
      const latex = normalizeLatexInput(placeholder.latex);
      return `my replaceRangeAndConvert(processName, ${escapeAppleScriptInteger(placeholder.start)}, ${escapeAppleScriptInteger(placeholder.length, latex.length)}, ${appleScriptString(latex)})`;
    })
    .join("\n");

  return buildEquationAutomationScript(
    `
${commands}
return my encodeConversionResult()
`,
  );
}

function parseAutomationResult(output) {
  const raw = String(output ?? "").trim();
  const fields = new Map(
    raw
      .split(";")
      .map((part) => part.split("=", 2))
      .filter((entry) => entry.length === 2)
      .map(([key, value]) => [key.trim(), value.trim()]),
  );
  return {
    latexFailed: fields.get("latexFailed") === "true",
    professionalFailed: fields.get("professionalFailed") === "true",
  };
}

function buildAutomationOutcomeMessage(nativeCount, automation) {
  const targetLabel = nativeCount > 1 ? `已创建 ${nativeCount} 个公式块` : "已创建公式块";
  if (automation.latexFailed && automation.professionalFailed) {
    return `${targetLabel}，但自动执行 LaTeX 转数学公式和 Professional 失败，请手动点击。`;
  }
  if (automation.latexFailed) {
    return `${targetLabel}，但自动执行 LaTeX 转数学公式失败，请手动点击。`;
  }
  if (automation.professionalFailed) {
    return `${targetLabel}，但自动执行 Professional 失败，请手动点击。`;
  }
  return `${targetLabel}，并已自动执行 LaTeX 转数学公式和 Professional。`;
}

function guiAutomationMessage(error) {
  const message = error instanceof Error ? error.message : String(error);
  if (/not authorized|not permitted|-1743|辅助功能|accessibility/i.test(message)) {
    return "已检测到 PowerPoint，但 helper 缺少 macOS 辅助功能权限。请在“系统设置 > 隐私与安全性 > 辅助功能”中允许终端或 Node 控制电脑。";
  }
  return `已检测到 PowerPoint，但界面自动化不可用：${message}。请确认已允许终端或 Node 控制 Microsoft PowerPoint，并保持 PowerPoint 窗口处于前台。`;
}

async function isPowerPointRunning() {
  const script = `
set processName to ${appleScriptString(POWERPOINT_PROCESS_NAME)}
tell application "System Events" to return exists process processName
  `.trim();
  try {
    return (await runOsaScript("powerpoint-running", script)) === "true";
  } catch {
    return false;
  }
}

async function probeGuiAutomation() {
  try {
    const result = await runOsaScript("probe", buildGuiAutomationProbeScript());
    if (result === "ok") {
      return { available: true, accessibilityGranted: true, message: "本地公式 helper 已运行，PowerPoint 界面自动化可用。" };
    }
    if (result === "ui-disabled") {
      return {
        available: false,
        accessibilityGranted: false,
        message: "已检测到 PowerPoint，但 helper 缺少 macOS 辅助功能权限。请在“系统设置 > 隐私与安全性 > 辅助功能”中允许终端或 Node 控制电脑。",
      };
    }
    if (result === "powerpoint-not-running") {
      return {
        available: false,
        accessibilityGranted: true,
        message: "本地公式 helper 已运行，但未检测到 Microsoft PowerPoint。请先打开 PowerPoint。",
      };
    }
    return {
      available: false,
      accessibilityGranted: true,
      message: `已检测到 PowerPoint，但 helper 状态未知：${result}`,
    };
  } catch (error) {
    return {
      available: false,
      accessibilityGranted: false,
      message: guiAutomationMessage(error),
    };
  }
}

async function probeEquationScriptSyntax() {
  return withTempAppleScript("syntax-check", buildSyntaxProbeScript(), async ({ scriptPath }) => {
    const syntaxCheck = await compileAppleScriptIfAvailable(scriptPath);
    if (syntaxCheck.ok === false) {
      return {
        ok: false,
        message: syntaxCheck.message,
      };
    }
    if (syntaxCheck.ok === true) {
      return {
        ok: true,
        message: syntaxCheck.message,
      };
    }
    return {
      ok: undefined,
      message: syntaxCheck.message,
    };
  });
}

async function health() {
  const powerpointRunning = await isPowerPointRunning();
  const guiAutomation = await probeGuiAutomation();
  const syntaxCheck = await probeEquationScriptSyntax();
  const nativeEquationAvailable = guiAutomation.available && syntaxCheck.ok !== false;
  const message = !powerpointRunning
    ? "本地公式 helper 已运行，但未检测到 Microsoft PowerPoint。请先打开 PowerPoint。"
    : syntaxCheck.ok === false
      ? `本地公式 helper 的 AppleScript 模板编译失败：${syntaxCheck.message}`
      : guiAutomation.message;
  return {
    ok: true,
    helperBuildId: HELPER_BUILD_ID,
    scriptExecutionMode: SCRIPT_EXECUTION_MODE,
    powerpointRunning,
    nativeEquationAvailable,
    guiAutomationAvailable: guiAutomation.available,
    accessibilityGranted: guiAutomation.accessibilityGranted,
    hostSelectionApiRequired: false,
    latexRibbonAvailable: guiAutomation.available,
    unicodeMathFallbackAvailable: false,
    equationScriptSyntaxOk: syntaxCheck.ok,
    equationScriptSyntaxMessage: syntaxCheck.message,
    message,
  };
}

export async function rootStatus() {
  const status = await health();
  return {
    ...status,
    helper: "SlideSCI native equation helper",
    note: "nativeEquationAvailable 仅表示 helper 可驱动 PowerPoint，不代表宿主支持 Office.js 文本范围或 PowerPoint 原生表格 API。",
    endpoints: HELPER_ENDPOINTS,
  };
}

async function convertSelectionToEquation(payload = {}) {
  const output = await runOsaScript("convert-selection", buildConvertSelectionScript(payload));
  const automation = parseAutomationResult(output);
  return {
    ok: true,
    helperBuildId: HELPER_BUILD_ID,
    mode: "native",
    nativeCount: 1,
    strategyUsed: automation.latexFailed ? undefined : "latex-ribbon",
    message: buildAutomationOutcomeMessage(1, automation),
  };
}

async function convertShapeRangesToEquation(payload = {}) {
  const placeholders = Array.isArray(payload.placeholders) ? payload.placeholders : [];
  if (placeholders.length === 0) {
    throw new Error("没有需要转换的公式占位符。");
  }
  const output = await runOsaScript("convert-shape-ranges", buildConvertShapeRangesScript(payload));
  const automation = parseAutomationResult(output);
  return {
    ok: true,
    helperBuildId: HELPER_BUILD_ID,
    mode: "native",
    nativeCount: placeholders.length,
    strategyUsed: automation.latexFailed ? undefined : "latex-ribbon",
    message: buildAutomationOutcomeMessage(placeholders.length, automation),
  };
}

export function createNativeEquationHelperServer() {
  return http.createServer(async (req, res) => {
    if (req.method === "OPTIONS") {
      json(res, 204, {});
      return;
    }

    try {
      if (req.method === "GET" && req.url === "/") {
        json(res, 200, await rootStatus());
        return;
      }

      if (req.method === "GET" && req.url === "/health") {
        json(res, 200, await health());
        return;
      }

      if (req.method === "POST" && req.url === "/equation/convert-selection") {
        const body = await readBody(req);
        if (!(await isPowerPointRunning())) {
          json(res, 503, {
            ok: false,
            mode: "unsupported",
            message: "未检测到 Microsoft PowerPoint。请打开 PowerPoint，并保持要转换的文本处于选中状态。",
          });
          return;
        }
        json(res, 200, await convertSelectionToEquation(body));
        return;
      }

      if (req.method === "POST" && req.url === "/equation/convert-shape-ranges") {
        const body = await readBody(req);
        if (!(await isPowerPointRunning())) {
          json(res, 503, {
            ok: false,
            mode: "unsupported",
            message: "未检测到 Microsoft PowerPoint。请打开 PowerPoint，并保持要转换的文本框处于选中状态。",
          });
          return;
        }
        json(res, 200, await convertShapeRangesToEquation(body));
        return;
      }

      if (
        req.method === "POST" &&
        (req.url === "/equation/insert-textbox" || req.url === "/equation/insert-block")
      ) {
        await readBody(req);
        json(res, 410, deprecatedInsertEndpointResponse());
        return;
      }

      json(res, 404, unknownHelperApiResponse());
    } catch (error) {
      json(res, 500, {
        ok: false,
        helperBuildId: HELPER_BUILD_ID,
        mode: "unsupported",
        message: guiAutomationMessage(error),
      });
    }
  });
}

export function startNativeEquationHelperServer() {
  const server = createNativeEquationHelperServer();
  server.listen(PORT, HOST, () => {
    console.log(`SlideSCI native equation helper listening on http://${HOST}:${PORT}`);
  });
  return server;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  startNativeEquationHelperServer();
}
