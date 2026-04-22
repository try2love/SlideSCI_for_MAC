import http from "node:http";
import { execFile } from "node:child_process";

const HOST = "127.0.0.1";
const PORT = Number(process.env.SLIDESCI_NATIVE_HELPER_PORT || 17926);

function json(res, status, data) {
  res.writeHead(status, {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json; charset=utf-8",
  });
  res.end(JSON.stringify(data));
}

function runOsaScript(script, args = []) {
  return new Promise((resolve, reject) => {
    execFile("osascript", ["-e", script, ...args], { timeout: 8000 }, (error, stdout, stderr) => {
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

async function isPowerPointRunning() {
  const script = 'tell application "System Events" to return exists process "Microsoft PowerPoint"';
  try {
    return (await runOsaScript(script)) === "true";
  } catch {
    return false;
  }
}

async function health() {
  const powerpointRunning = await isPowerPointRunning();
  return {
    ok: true,
    powerpointRunning,
    nativeEquationAvailable: powerpointRunning,
    message: powerpointRunning
      ? "本地公式 helper 已运行，PowerPoint 已打开。"
      : "本地公式 helper 已运行，但未检测到 Microsoft PowerPoint。请先打开 PowerPoint 并选中文本范围。",
  };
}

async function convertSelectionToEquation() {
  const script = `
tell application "Microsoft PowerPoint"
  activate
  try
    do Visual Basic "Application.CommandBars.ExecuteMso \\"EquationInsertNew\\""
    delay 0.05
    do Visual Basic "Application.CommandBars.ExecuteMso \\"EquationProfessional\\""
    return "native"
  on error errMsg number errNum
    error errMsg number errNum
  end try
end tell
`;
  await runOsaScript(script);
  return {
    ok: true,
    mode: "native",
    message: "已请求 PowerPoint 将当前选中文本转换为原生公式。",
  };
}

const server = http.createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    json(res, 204, {});
    return;
  }

  try {
    if (req.method === "GET" && req.url === "/health") {
      json(res, 200, await health());
      return;
    }

    if (req.method === "POST" && req.url === "/equation/convert-selection") {
      await readBody(req);
      if (!(await isPowerPointRunning())) {
        json(res, 503, {
          ok: false,
          mode: "unsupported",
          message: "未检测到 Microsoft PowerPoint。请打开 PowerPoint，并允许终端/Node 自动化控制 PowerPoint。",
        });
        return;
      }
      json(res, 200, await convertSelectionToEquation());
      return;
    }

    json(res, 404, { ok: false, message: "未知 helper API。" });
  } catch (error) {
    json(res, 500, {
      ok: false,
      mode: "unsupported",
      message:
        error instanceof Error
          ? `原生公式转换失败：${error.message}。请确认 macOS 已允许终端或 Node 自动化控制 Microsoft PowerPoint。`
          : String(error),
    });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`SlideSCI native equation helper listening on http://${HOST}:${PORT}`);
});
