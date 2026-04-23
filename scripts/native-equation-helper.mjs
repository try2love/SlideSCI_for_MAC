import http from "node:http";
import { execFile } from "node:child_process";
import { pathToFileURL } from "node:url";

const HOST = "127.0.0.1";
const PORT = Number(process.env.SLIDESCI_NATIVE_HELPER_PORT || 17926);

const POWERPOINT_PROCESS_NAME = "Microsoft PowerPoint";
const MENU_LABELS = {
  insertMenu: ["Insert", "插入"],
  equationItem: ["Equation", "Equation...", "公式"],
  contextualMenus: ["Equation", "公式", "Design", "设计"],
  professionalItem: ["Professional", "Professional...", "专业", "专业型"],
};

export const HELPER_ENDPOINTS = [
  "GET /",
  "GET /health",
  "POST /equation/convert-selection",
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
    message: "该 helper 接口已弃用。请改为由 Office.js 插入并选中文本后，再调用 /equation/convert-selection。",
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

export function buildConvertSelectionScript() {
  return `
set processName to ${appleScriptString(POWERPOINT_PROCESS_NAME)}
set insertMenuCandidates to ${appleScriptList(MENU_LABELS.insertMenu)}
set equationItemCandidates to ${appleScriptList(MENU_LABELS.equationItem)}
set contextualMenuCandidates to ${appleScriptList(MENU_LABELS.contextualMenus)}
set professionalItemCandidates to ${appleScriptList(MENU_LABELS.professionalItem)}

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

on tryProfessionalLayout(processName, menuCandidates, itemCandidates)
  try
    return my clickFirstMatchingMenuItem(processName, menuCandidates, itemCandidates)
  on error
    return "skipped"
  end try
end tryProfessionalLayout

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
delay 0.18
my tryProfessionalLayout(processName, contextualMenuCandidates, professionalItemCandidates)
return "native"
`.trim();
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
    message: powerpointRunning ? guiAutomation.message : "本地公式 helper 已运行，但未检测到 Microsoft PowerPoint。请先打开 PowerPoint。",
  };
}

export async function rootStatus() {
  const status = await health();
  return {
    ...status,
    helper: "SlideSCI native equation helper",
    endpoints: HELPER_ENDPOINTS,
  };
}

async function convertSelectionToEquation(payload = {}) {
  const latex = normalizeLatexInput(payload.latex);
  await runOsaScript(buildConvertSelectionScript());
  return {
    ok: true,
    mode: "native",
    nativeCount: 1,
    message: latex ? `已将当前选区转换为原生公式：${latex}` : "已将当前选区转换为原生公式。",
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
