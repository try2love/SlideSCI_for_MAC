import http from "node:http";
import { execFile } from "node:child_process";

const HOST = "127.0.0.1";
const PORT = Number(process.env.SLIDESCI_NATIVE_HELPER_PORT || 17926);
const TEXT_ORIENTATION_HORIZONTAL = 1;
const MSO_TRUE = -1;
const MSO_FALSE = 0;

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

function numberLiteral(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number.toFixed(2) : fallback.toFixed(2);
}

function vbaString(value) {
  const text = String(value ?? "");
  const parts = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  return parts
    .map((part) => `"${part.replace(/"/g, '""')}"`)
    .join(" & vbCrLf & ");
}

function appleScriptString(value) {
  return `"${String(value ?? "").replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r?\n/g, "\\n")}"`;
}

function colorToRgb(color) {
  if (!color || typeof color !== "string") {
    return null;
  }
  const normalized = color.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return null;
  }
  return [
    Number.parseInt(normalized.slice(0, 2), 16),
    Number.parseInt(normalized.slice(2, 4), 16),
    Number.parseInt(normalized.slice(4, 6), 16),
  ];
}

function rgbExpression(color) {
  const rgb = colorToRgb(color);
  return rgb ? `RGB(${rgb[0]}, ${rgb[1]}, ${rgb[2]})` : null;
}

function boolValue(value) {
  return value ? MSO_TRUE : MSO_FALSE;
}

function shapeName(prefix = "SlideSCI_NativeEquation") {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function appendFontStatements(statements, rangeExpression, style = {}) {
  if (style.fontName || style.fontFamily) {
    statements.push(`${rangeExpression}.Font.Name = ${vbaString(style.fontName || style.fontFamily)}`);
  }
  if (Number.isFinite(Number(style.fontSize))) {
    statements.push(`${rangeExpression}.Font.Size = ${numberLiteral(style.fontSize, 14)}`);
  }
  const rgb = rgbExpression(style.color);
  if (rgb) {
    statements.push(`${rangeExpression}.Font.Color.RGB = ${rgb}`);
  }
  if (typeof style.bold === "boolean") {
    statements.push(`${rangeExpression}.Font.Bold = ${boolValue(style.bold)}`);
  }
  if (typeof style.italic === "boolean") {
    statements.push(`${rangeExpression}.Font.Italic = ${boolValue(style.italic)}`);
  }
  if (typeof style.underline === "boolean") {
    statements.push(`${rangeExpression}.Font.Underline = ${boolValue(style.underline)}`);
  }
  if (typeof style.superscript === "boolean") {
    statements.push(`${rangeExpression}.Font.Superscript = ${boolValue(style.superscript)}`);
  }
  if (typeof style.subscript === "boolean") {
    statements.push(`${rangeExpression}.Font.Subscript = ${boolValue(style.subscript)}`);
  }
}

function appendShapeStyleStatements(statements, shapeVariable, style = {}) {
  appendFontStatements(statements, `${shapeVariable}.TextFrame.TextRange`, style);

  if (style.align === "center") {
    statements.push(`${shapeVariable}.TextFrame.TextRange.ParagraphFormat.Alignment = 2`);
  } else if (style.align === "left") {
    statements.push(`${shapeVariable}.TextFrame.TextRange.ParagraphFormat.Alignment = 1`);
  }

  const fillRgb = rgbExpression(style.fillColor);
  if (fillRgb) {
    statements.push(`${shapeVariable}.Fill.Solid`);
    statements.push(`${shapeVariable}.Fill.ForeColor.RGB = ${fillRgb}`);
    statements.push(`${shapeVariable}.Fill.Transparency = 0`);
  }

  const borderRgb = rgbExpression(style.borderColor);
  if (borderRgb) {
    statements.push(`${shapeVariable}.Line.ForeColor.RGB = ${borderRgb}`);
    statements.push(`${shapeVariable}.Line.Weight = ${numberLiteral(style.borderWeight, 1)}`);
  }
}

function warmEquationStatements() {
  return [
    "On Error Resume Next",
    `Set warmShape = sld.Shapes.AddTextbox(${TEXT_ORIENTATION_HORIZONTAL}, 0, 0, 10, 10)`,
    "warmShape.Select",
    "Application.ActiveWindow.Selection.TextRange.Select",
    'Application.CommandBars.ExecuteMso "EquationInsertNew"',
    "Set warmEquation = Application.ActiveWindow.Selection.ShapeRange(1)",
    "warmEquation.TextFrame.TextRange.Characters(1, Len(warmEquation.TextFrame.TextRange.Text) - 1).Text = \"a\"",
    'Application.CommandBars.ExecuteMso "EquationProfessional"',
    "warmEquation.Delete",
    "warmShape.Delete",
    "On Error GoTo 0",
  ];
}

function buildVba(statements) {
  return statements.join(" : ");
}

function runPowerPointVba(vba) {
  const script = `
tell application "Microsoft PowerPoint"
  activate
  try
    do Visual Basic ${appleScriptString(vba)}
    return "ok"
  on error errMsg number errNum
    error errMsg number errNum
  end try
end tell
`;
  return runOsaScript(script);
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
  let nativeEquationAvailable = false;
  let message = "本地公式 helper 已运行，但未检测到 Microsoft PowerPoint。请先打开 PowerPoint。";
  if (powerpointRunning) {
    try {
      await runPowerPointVba("Dim sld As Object : Set sld = Application.ActiveWindow.View.Slide");
      nativeEquationAvailable = true;
      message = "本地公式 helper 已运行，PowerPoint 自动化可用；公式命令会在插入时验证。";
    } catch (error) {
      message = `已检测到 PowerPoint，但自动化不可用：${error instanceof Error ? error.message : String(error)}。请确认已允许终端或 Node 自动化控制 Microsoft PowerPoint，并打开一个演示文稿。`;
    }
  }
  return {
    ok: true,
    powerpointRunning,
    nativeEquationAvailable,
    message,
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

async function insertTextBoxWithEquations(payload) {
  const text = String(payload.text ?? "");
  const equations = Array.isArray(payload.equations) ? payload.equations : [];
  if (!text.trim()) {
    throw new Error("文本框内容不能为空。");
  }
  if (equations.length === 0) {
    throw new Error("没有需要转换的公式 run。");
  }

  const box = payload.box ?? {};
  const baseStyle = payload.baseStyle ?? {};
  const runs = Array.isArray(payload.runs) ? payload.runs : [];
  const name = shapeName("SlideSCI_NativeText");
  const statements = [
    "Dim sld As Object",
    "Dim shp As Object",
    "Dim rng As Object",
    "Dim resultShape As Object",
    "Dim warmShape As Object",
    "Dim warmEquation As Object",
    "Set sld = Application.ActiveWindow.View.Slide",
    `Set shp = sld.Shapes.AddTextbox(${TEXT_ORIENTATION_HORIZONTAL}, ${numberLiteral(box.left, 80)}, ${numberLiteral(box.top, 80)}, ${numberLiteral(box.width, 360)}, ${numberLiteral(box.height, 80)})`,
    `shp.Name = ${vbaString(name)}`,
    `shp.TextFrame.TextRange.Text = ${vbaString(text)}`,
  ];
  appendShapeStyleStatements(statements, "shp", baseStyle);

  for (const run of runs) {
    if (!Number.isFinite(Number(run.start)) || !Number.isFinite(Number(run.length)) || Number(run.length) <= 0) {
      continue;
    }
    statements.push(`Set rng = shp.TextFrame.TextRange.Characters(${Math.floor(Number(run.start)) + 1}, ${Math.floor(Number(run.length))})`);
    appendFontStatements(statements, "rng", run.style ?? {});
    if (run.style?.align === "center") {
      statements.push("rng.ParagraphFormat.Alignment = 2");
    }
  }

  statements.push(...warmEquationStatements());

  for (const equation of [...equations].sort((a, b) => Number(b.start) - Number(a.start))) {
    if (!Number.isFinite(Number(equation.start)) || !Number.isFinite(Number(equation.length)) || Number(equation.length) <= 0) {
      continue;
    }
    statements.push(`Set rng = shp.TextFrame.TextRange.Characters(${Math.floor(Number(equation.start)) + 1}, ${Math.floor(Number(equation.length))})`);
    statements.push(`rng.Text = ${vbaString(normalizeLatexInput(equation.latex))}`);
    statements.push("rng.Select");
    statements.push('Application.CommandBars.ExecuteMso "EquationInsertNew"');
    statements.push('Application.CommandBars.ExecuteMso "EquationProfessional"');
  }

  statements.push("shp.TextFrame.AutoSize = 1");

  await runPowerPointVba(buildVba(statements));
  return {
    ok: true,
    mode: "native",
    id: name,
    nativeCount: equations.length,
    message: `已插入含 ${equations.length} 个原生公式的文本框。`,
  };
}

async function insertEquationBlock(payload) {
  const latex = normalizeLatexInput(payload.latex);
  if (!latex) {
    throw new Error("LaTeX 公式不能为空。");
  }

  const box = payload.box ?? {};
  const style = payload.style ?? {};
  const name = shapeName("SlideSCI_NativeEquation");
  const statements = [
    "Dim sld As Object",
    "Dim shp As Object",
    "Dim rng As Object",
    "Dim resultShape As Object",
    "Dim warmShape As Object",
    "Dim warmEquation As Object",
    "Set sld = Application.ActiveWindow.View.Slide",
    ...warmEquationStatements(),
    `Set shp = sld.Shapes.AddTextbox(${TEXT_ORIENTATION_HORIZONTAL}, ${numberLiteral(box.left, 160)}, ${numberLiteral(box.top, 120)}, ${numberLiteral(box.width, 500)}, ${numberLiteral(box.height, 120)})`,
    `shp.Name = ${vbaString(name)}`,
    `shp.AlternativeText = ${vbaString(latex)}`,
    `shp.TextFrame.TextRange.Text = ${vbaString(latex)}`,
  ];
  appendShapeStyleStatements(statements, "shp", {
    fontName: "Cambria Math",
    fontSize: 18,
    color: "#000000",
    align: "center",
    ...style,
  });
  statements.push("Set rng = shp.TextFrame.TextRange.Characters(1, Len(shp.TextFrame.TextRange.Text))");
  statements.push("rng.Select");
  statements.push('Application.CommandBars.ExecuteMso "EquationInsertNew"');
  statements.push('Application.CommandBars.ExecuteMso "EquationProfessional"');
  statements.push("On Error Resume Next");
  statements.push("Set resultShape = Application.ActiveWindow.Selection.ShapeRange(1)");
  statements.push(`resultShape.Name = ${vbaString(name)}`);
  statements.push(`resultShape.AlternativeText = ${vbaString(latex)}`);
  statements.push("resultShape.TextFrame.AutoSize = 1");
  statements.push(`resultShape.Left = ${numberLiteral(box.left, 160)}`);
  statements.push(`resultShape.Top = ${numberLiteral(box.top, 120)}`);
  statements.push("On Error GoTo 0");

  await runPowerPointVba(buildVba(statements));
  return {
    ok: true,
    mode: "native",
    id: name,
    nativeCount: 1,
    message: "已插入 LaTeX 原生公式。",
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

    if (req.method === "POST" && req.url === "/equation/insert-textbox") {
      const body = await readBody(req);
      if (!(await isPowerPointRunning())) {
        json(res, 503, {
          ok: false,
          mode: "unsupported",
          message: "未检测到 Microsoft PowerPoint。请打开 PowerPoint，并允许终端/Node 自动化控制 PowerPoint。",
        });
        return;
      }
      json(res, 200, await insertTextBoxWithEquations(body));
      return;
    }

    if (req.method === "POST" && req.url === "/equation/insert-block") {
      const body = await readBody(req);
      if (!(await isPowerPointRunning())) {
        json(res, 503, {
          ok: false,
          mode: "unsupported",
          message: "未检测到 Microsoft PowerPoint。请打开 PowerPoint，并允许终端/Node 自动化控制 PowerPoint。",
        });
        return;
      }
      json(res, 200, await insertEquationBlock(body));
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
