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
const APPLESCRIPT_RUNNER = process.env.SLIDESCI_APPLESCRIPT_RUNNER || "";

const POWERPOINT_PROCESS_NAME = "Microsoft PowerPoint";
const MENU_LABELS = {
  insertMenu: ["Insert", "插入"],
  equationItem: ["Equation", "Equation...", "公式"],
};
const DEFAULT_STRATEGY_ORDER = ["equation-insert"];

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
        const stdoutText = stdout.trim();
        const stderrText = stderr.trim();
        const details = [
          stderrText,
          stdoutText,
          error.message,
          error.code ? `code=${String(error.code)}` : "",
          error.signal ? `signal=${String(error.signal)}` : "",
        ]
          .map((value) => String(value || "").trim())
          .filter(Boolean);
        const message = details.join(" | ");
        const wrapped = new Error(message);
        wrapped.cause = error;
        wrapped.stdout = stdoutText;
        wrapped.stderr = stderrText;
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
    const command = APPLESCRIPT_RUNNER || "osascript";
    const args = APPLESCRIPT_RUNNER ? ["--run-applescript-file", scriptPath] : [scriptPath];
    const { stdout } = await execFileAsync(command, args, { timeout: 8000 });
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

on triggerEquationForRange(processName)
  return my tryInsertEquation(processName, insertMenuCandidates, equationItemCandidates)
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
return "equation-insert"
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
return "equation-insert"
`,
  );
}

function guiAutomationMessage(error) {
  const message = error instanceof Error ? error.message : String(error);
  if (/not authorized|not permitted|-1743|辅助功能|accessibility/i.test(message)) {
    return APPLESCRIPT_RUNNER
      ? "已检测到 PowerPoint，但 helper 缺少 macOS 辅助功能权限。请在“系统设置 > 隐私与安全性 > 辅助功能”中允许 SlideSCICompanion 控制电脑，然后重新打开 PowerPoint。"
      : "已检测到 PowerPoint，但 helper 缺少 macOS 辅助功能权限。请在“系统设置 > 隐私与安全性 > 辅助功能”中允许终端或 Node 控制电脑。";
  }
  return APPLESCRIPT_RUNNER
    ? `已检测到 PowerPoint，但界面自动化不可用：${message}。请确认已允许 SlideSCICompanion 控制电脑，并保持 PowerPoint 窗口处于前台。`
    : `已检测到 PowerPoint，但界面自动化不可用：${message}。请确认已允许终端或 Node 控制 Microsoft PowerPoint，并保持 PowerPoint 窗口处于前台。`;
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
        message: APPLESCRIPT_RUNNER
          ? "已检测到 PowerPoint，但 helper 缺少 macOS 辅助功能权限。请在“系统设置 > 隐私与安全性 > 辅助功能”中允许 SlideSCICompanion 控制电脑，然后重新打开 PowerPoint。"
          : "已检测到 PowerPoint，但 helper 缺少 macOS 辅助功能权限。请在“系统设置 > 隐私与安全性 > 辅助功能”中允许终端或 Node 控制电脑。",
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
    latexRibbonAvailable: false,
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
  const strategyUsed = await runOsaScript("convert-selection", buildConvertSelectionScript(payload));
  return {
    ok: true,
    helperBuildId: HELPER_BUILD_ID,
    mode: "native",
    nativeCount: 1,
    strategyUsed,
    message: "已将当前选区转换为原生公式对象。",
  };
}

async function convertShapeRangesToEquation(payload = {}) {
  const placeholders = Array.isArray(payload.placeholders) ? payload.placeholders : [];
  if (placeholders.length === 0) {
    throw new Error("没有需要转换的公式占位符。");
  }
  const strategyUsed = await runOsaScript("convert-shape-ranges", buildConvertShapeRangesScript(payload));
  return {
    ok: true,
    helperBuildId: HELPER_BUILD_ID,
    mode: "native",
    nativeCount: placeholders.length,
    strategyUsed,
    message: `已将 ${placeholders.length} 个占位符转换为原生公式对象。`,
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
