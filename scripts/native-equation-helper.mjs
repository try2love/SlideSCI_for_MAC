import http from "node:http";
import { execFile } from "node:child_process";
import { pathToFileURL } from "node:url";

const HOST = "127.0.0.1";
const PORT = Number(process.env.SLIDESCI_NATIVE_HELPER_PORT || 17926);

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
const DEFAULT_STRATEGY_ORDER = ["latex-ribbon", "unicode-math"];

export const HELPER_ENDPOINTS = [
  "GET /",
  "GET /health",
  "POST /equation/convert-selection",
  "POST /equation/convert-shape-ranges",
  "POST /equation/insert-textbox (deprecated)",
  "POST /equation/insert-block (deprecated)",
];

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

function runOsaScript(script) {
  return new Promise((resolve, reject) => {
    execFile("osascript", ["-e", script], { timeout: 8000 }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error((stderr || error.message).trim()));
        return;
      }
      resolve(stdout.trim());
    });
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

function strategyOrderForPayload(payload = {}) {
  const raw = Array.isArray(payload.strategyOrder) ? payload.strategyOrder : [];
  const filtered = raw.filter((item) => item === "latex-ribbon" || item === "unicode-math");
  return filtered.length > 0 ? filtered : DEFAULT_STRATEGY_ORDER;
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

function escapeAppleScriptInteger(value, fallback = 0) {
  const integer = Number.parseInt(String(value ?? fallback), 10);
  return Number.isFinite(integer) ? String(integer) : String(fallback);
}

function buildEquationAutomationScript(strategyOrder, commands) {
  return `
set processName to ${appleScriptString(POWERPOINT_PROCESS_NAME)}
set insertMenuCandidates to ${appleScriptList(MENU_LABELS.insertMenu)}
set equationItemCandidates to ${appleScriptList(MENU_LABELS.equationItem)}
set contextualMenuCandidates to ${appleScriptList(MENU_LABELS.contextualMenus)}
set professionalItemCandidates to ${appleScriptList(MENU_LABELS.professionalItem)}
set convertControlCandidates to ${appleScriptList(RIBBON_LABELS.convertControls)}
set latexRibbonCandidates to ${appleScriptList(RIBBON_LABELS.latexItems)}
set professionalRibbonCandidates to ${appleScriptList(RIBBON_LABELS.professionalItems)}
set strategyOrder to ${appleScriptList(strategyOrder)}

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
  if labelText is missing value then
    return ""
  end if
  return labelText as text
end elementLabel

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
  end ignoring case
  return false
end matchesCandidates

on pressControl(uiElement)
  try
    perform action "AXPress" of uiElement
  on error
    click uiElement
  end try
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

on currentSelectionRange(processName)
  set focusedElement to my ensureFocusedEditableElement(processName)
  try
    return value of attribute "AXSelectedTextRange" of focusedElement
  on error errMsg number errNum
    error "无法读取当前文本选区：" & errMsg number errNum
  end try
end currentSelectionRange

on selectCharacterRange(focusedElement, startIndex, lengthValue)
  try
    set value of attribute "AXSelectedTextRange" of focusedElement to {startIndex, lengthValue}
  on error errMsg number errNum
    error "无法通过辅助功能选择文本范围：" & errMsg number errNum
  end try
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
        keystroke "=" using {control down}
      end tell
    end tell
    return "shortcut"
  end try
end tryInsertEquation

on tryLatexRibbonConvert(processName)
  try
    return my pressFirstMatchingProcessControl(processName, latexRibbonCandidates)
  on error
    return my openControlAndChoose(processName, convertControlCandidates, latexRibbonCandidates)
  end try
end tryLatexRibbonConvert

on tryProfessionalLayout(processName)
  try
    return my pressFirstMatchingProcessControl(processName, professionalRibbonCandidates)
  on error
    try
      return my openControlAndChoose(processName, convertControlCandidates, professionalRibbonCandidates)
    on error
      try
        return my clickFirstMatchingMenuItem(processName, contextualMenuCandidates, professionalItemCandidates)
      on error
        tell application "System Events"
          tell process processName
            keystroke "=" using {control down}
          end tell
        end tell
        return "shortcut"
      end try
    end try
  end try
end tryProfessionalLayout

on compileRangeWithLatexRibbon(processName, startIndex, lengthValue, latexText)
  my replaceRangeText(processName, startIndex, lengthValue, latexText)
  my reselectRange(processName, startIndex, (length of latexText))
  my tryInsertEquation(processName, insertMenuCandidates, equationItemCandidates)
  delay 0.24
  my tryLatexRibbonConvert(processName)
  delay 0.18
  my tryProfessionalLayout(processName)
  return "latex-ribbon"
end compileRangeWithLatexRibbon

on compileRangeWithUnicodeMath(processName, startIndex, lengthValue, unicodeMathText)
  if unicodeMathText is "" then
    error "未提供 UnicodeMath 兜底表达式。"
  end if
  my replaceRangeText(processName, startIndex, lengthValue, unicodeMathText)
  my reselectRange(processName, startIndex, (length of unicodeMathText))
  my tryInsertEquation(processName, insertMenuCandidates, equationItemCandidates)
  delay 0.24
  my tryProfessionalLayout(processName)
  return "unicode-math"
end compileRangeWithUnicodeMath

on convertRangeWithStrategies(processName, startIndex, lengthValue, latexText, unicodeMathText, strategyOrder)
  repeat with strategyName in strategyOrder
    set currentStrategy to contents of strategyName
    if currentStrategy is "latex-ribbon" then
      try
        return my compileRangeWithLatexRibbon(processName, startIndex, lengthValue, latexText)
      end try
    else if currentStrategy is "unicode-math" then
      try
        return my compileRangeWithUnicodeMath(processName, startIndex, lengthValue, unicodeMathText)
      end try
    end if
  end repeat
  error "未能自动完成 LaTeX/UnicodeMath 转换。"
end convertRangeWithStrategies

${commands}
`.trim();
}

export function buildConvertSelectionScript(payload = {}) {
  const latex = normalizeLatexInput(payload.latex);
  const unicodeMath = String(payload.unicodeMath ?? "");
  const strategyOrder = strategyOrderForPayload(payload);
  return buildEquationAutomationScript(
    strategyOrder,
    `
set selectedRange to my currentSelectionRange(processName)
set strategyUsed to my convertRangeWithStrategies(processName, item 1 of selectedRange, item 2 of selectedRange, ${appleScriptString(latex)}, ${appleScriptString(unicodeMath)}, strategyOrder)
return strategyUsed
`,
  );
}

export function buildConvertShapeRangesScript(payload = {}) {
  const placeholders = Array.isArray(payload.placeholders) ? payload.placeholders : [];
  const strategyOrder = strategyOrderForPayload(payload);
  const commands = placeholders
    .slice()
    .sort((a, b) => Number(b.start) - Number(a.start))
    .map((placeholder) => {
      const latex = normalizeLatexInput(placeholder.latex);
      const unicodeMath = String(placeholder.unicodeMath ?? "");
      return `set strategyUsed to my convertRangeWithStrategies(processName, ${escapeAppleScriptInteger(placeholder.start)}, ${escapeAppleScriptInteger(placeholder.length, latex.length)}, ${appleScriptString(latex)}, ${appleScriptString(unicodeMath)}, strategyOrder)`;
    })
    .join("\n");

  return buildEquationAutomationScript(
    strategyOrder,
    `
${commands}
return strategyUsed
`,
  );
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
    return (await runOsaScript(script)) === "true";
  } catch {
    return false;
  }
}

async function probeGuiAutomation() {
  try {
    const result = await runOsaScript(buildGuiAutomationProbeScript());
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

async function health() {
  const powerpointRunning = await isPowerPointRunning();
  const guiAutomation = await probeGuiAutomation();
  return {
    ok: true,
    powerpointRunning,
    nativeEquationAvailable: guiAutomation.available,
    guiAutomationAvailable: guiAutomation.available,
    accessibilityGranted: guiAutomation.accessibilityGranted,
    hostSelectionApiRequired: false,
    latexRibbonAvailable: guiAutomation.available,
    unicodeMathFallbackAvailable: true,
    message: powerpointRunning ? guiAutomation.message : "本地公式 helper 已运行，但未检测到 Microsoft PowerPoint。请先打开 PowerPoint。",
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
  const latex = normalizeLatexInput(payload.latex);
  const strategyUsed = await runOsaScript(buildConvertSelectionScript(payload));
  return {
    ok: true,
    mode: "native",
    nativeCount: 1,
    strategyUsed,
    message: latex ? `已将当前选区转换为原生公式（${strategyUsed}）：${latex}` : `已将当前选区转换为原生公式（${strategyUsed}）。`,
  };
}

async function convertShapeRangesToEquation(payload = {}) {
  const placeholders = Array.isArray(payload.placeholders) ? payload.placeholders : [];
  if (placeholders.length === 0) {
    throw new Error("没有需要转换的公式占位符。");
  }
  const strategyUsed = await runOsaScript(buildConvertShapeRangesScript(payload));
  return {
    ok: true,
    mode: "native",
    nativeCount: placeholders.length,
    strategyUsed,
    message: `已通过 GUI 自动化将 ${placeholders.length} 个公式占位符转换为原生公式（${strategyUsed}）。`,
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
