import { useMemo, useState } from "react";
import { arrangeShapes } from "../lib/layout";
import { LABEL_TEMPLATES } from "../lib/labels";
import { generateLabels } from "../lib/labels";
import { normalizeLatexInput } from "../lib/latex";
import { estimateTextBoxSize } from "../lib/textMetrics";
import { parseLengthToPt } from "../lib/units";
import type { LayoutMode, SortMode, TitlePlacement } from "../lib/types";
import { getCodeBlockStyle, getCodeHighlightRuns, CODE_LANGUAGES, CODE_LANGUAGE_LABELS } from "../services/codeBlock";
import { convertLatexToPngBase64 } from "../services/latexSvg";
import { formatMarkdownRenderResult, markdownToRenderBlocks, renderMarkdownBlocks } from "../services/markdownRender";
import { insertNativeEquationBlock, insertNativeEquationTextBox } from "../services/nativeEquation";
import {
  addLatexImage,
  addRichTextBox,
  addTableWithFallback,
  addTextBox,
  applyShapeStyleToSelected,
  copySelectedFormat,
  getSelectedLatexMetadata,
  getSelectedShapes,
  selectShapes,
  updateTextForShapes,
  updateShapesLayout,
} from "../services/powerpoint";
import {
  defaultSettings,
  loadClipboardState,
  loadSettings,
  resolveLatexSource,
  saveLatexForShape,
  saveClipboardState,
  saveSettings,
  type AppSettings,
} from "../services/settings";

function Field({
  label,
  children,
}: {
  label: string;
  children: React.ReactNode;
}) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
    </label>
  );
}

function Section({
  title,
  children,
}: {
  title: string;
  children: React.ReactNode;
}) {
  return (
    <section className="section">
      <h2>{title}</h2>
      {children}
    </section>
  );
}

export function App() {
  const [settings, setSettingsState] = useState<AppSettings>(() => loadSettings());
  const [status, setStatus] = useState("准备就绪。");
  const [busy, setBusy] = useState(false);
  const [code, setCode] = useState("");
  const [markdown, setMarkdown] = useState("");
  const [latex, setLatex] = useState("");

  const canRun = useMemo(() => !busy, [busy]);

  function setSettings(next: AppSettings): void {
    setSettingsState(next);
    saveSettings(next);
  }

  function updateSetting<K extends keyof AppSettings>(key: K, value: AppSettings[K]): void {
    setSettings({ ...settings, [key]: value });
  }

  async function run(label: string, action: () => Promise<string | void>): Promise<void> {
    setBusy(true);
    setStatus(`${label}...`);
    try {
      const message = await action();
      setStatus(message || `${label}完成。`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function arrangeSelectedShapes(): Promise<void> {
    const shapes = await getSelectedShapes();
    if (shapes.length === 0) {
      throw new Error("请先选择要排列的对象。");
    }

    const layouts = arrangeShapes(shapes, {
      colNum: settings.colNum,
      colSpace: settings.colSpace,
      rowSpace: settings.rowSpace,
      mode: settings.layoutMode,
      sortMode: settings.sortMode,
      customWidth: parseLengthToPt(settings.imgWidth),
      customHeight: parseLengthToPt(settings.imgHeight),
    });
    await updateShapesLayout(layouts);
  }

  async function addTitles(placement: TitlePlacement): Promise<void> {
    const shapes = await getSelectedShapes();
    if (shapes.length === 0) {
      throw new Error("请先选择要添加标题的对象。");
    }

    const ids: string[] = [];
    for (const shape of shapes) {
      const height = settings.titleFontSize * 2;
      const top =
        placement === "bottom"
          ? shape.top + shape.height + settings.titleDistance
          : shape.top - height - settings.titleDistance;
      ids.push(
        await addTextBox(
          settings.titleText,
          { left: shape.left, top, width: shape.width, height },
          {
            fontName: settings.titleFontName,
            fontSize: settings.titleFontSize,
            color: "#000000",
            align: settings.titleAlign,
          },
        ),
      );
    }
    await selectShapes(ids);
  }

  async function addLabels(): Promise<void> {
    const shapes = await getSelectedShapes();
    if (shapes.length === 0) {
      throw new Error("请先选择要添加标签的对象。");
    }

    const ids: string[] = [];
    const labels = generateLabels(shapes, settings.labelTemplate, settings.labelIndex);
    for (const item of labels) {
      ids.push(
        await addTextBox(
          item.label,
          {
            left: item.shape.left + settings.labelOffsetX,
            top: item.shape.top + settings.labelOffsetY,
            width: settings.labelFontSize * 3,
            height: settings.labelFontSize * 2,
          },
          {
            fontName: settings.labelFontName,
            fontSize: settings.labelFontSize,
            color: "#000000",
            bold: settings.labelBold,
            align: "left",
          },
        ),
      );
    }

    if (settings.labelAutoUpdate) {
      updateSetting("labelIndex", settings.labelIndex + labels.length);
    }
    await selectShapes(ids);
  }

  async function insertCodeBlock(): Promise<void> {
    if (!code.trim()) {
      throw new Error("请输入代码内容。");
    }
    const boxSize = estimateTextBoxSize(code.trim(), { fontSize: 12, monospace: true });
    await addRichTextBox(
      code.trim(),
      { left: 80, top: 80, ...boxSize },
      getCodeBlockStyle(settings.codeDarkBackground),
      getCodeHighlightRuns(code.trim(), settings.codeLanguage, settings.codeDarkBackground),
    );
  }

  async function insertLatexNative(): Promise<string | void> {
    const normalizedLatex = normalizeLatexInput(latex);
    if (!normalizedLatex) {
      throw new Error("请输入 LaTeX 内容。");
    }

    const result = await insertNativeEquationBlock({
      latex: normalizedLatex,
      box: { left: 160, top: 120, width: 500, height: 120 },
      style: {
        fontName: "Cambria Math",
        fontSize: 18,
        color: "#000000",
        align: "center",
      },
    });
    if (result.id) {
      saveLatexForShape(result.id, normalizedLatex);
    }
    return result.message || "插入 LaTeX 原生公式完成。";
  }

  async function insertLatexImage(): Promise<string | void> {
    const image = await convertLatexToPngBase64(latex);
    const width = Math.min(520, Math.max(120, image.width * 0.75));
    const height = Math.min(180, Math.max(48, image.height * 0.75));
    const inserted = await addLatexImage(image.base64, { left: 160, top: 120, width, height }, image.latex);
    saveLatexForShape(inserted.id, image.latex);
    if (inserted.warning) {
      return `插入 LaTeX 图片完成：${inserted.warning}`;
    }
  }

  async function loadLatexFromSelection(): Promise<void> {
    const metadata = await getSelectedLatexMetadata();
    const source = resolveLatexSource(metadata);
    if (!source) {
      throw new Error("没有找到该对象对应的 LaTeX 记录。请确认它是由本插件插入的公式。");
    }
    setLatex(source);
  }

  async function insertMarkdown(): Promise<string> {
    const blocks = markdownToRenderBlocks(markdown);
    if (blocks.length === 0) {
      throw new Error("请输入 Markdown 内容。");
    }

    let top = 80;
    const left = 80;
    const result = await renderMarkdownBlocks(blocks, {
      text: async (block) => {
        const boxSize = estimateTextBoxSize(block.text, { fontSize: block.fontSize, monospace: false });
        const box = { left, top, ...boxSize };
        if (block.equations.length > 0) {
          const inserted = await insertNativeEquationTextBox({
            text: block.text,
            equations: block.equations,
            box,
            baseStyle: {
              fontName: "微软雅黑",
              fontSize: 14,
              color: "#000000",
            },
            runs: block.runs,
          });
          top += boxSize.height + 12;
          if (inserted.id) {
            saveLatexForShape(inserted.id, block.equations.map((equation) => equation.latex).join("\n"));
          }
          return inserted.message;
        }
        await addRichTextBox(
          block.text,
          box,
          {
            fontName: "微软雅黑",
            fontSize: 14,
            color: "#000000",
          },
          block.runs,
        );
        top += boxSize.height + 12;
      },
      quote: async (block) => {
        const boxSize = estimateTextBoxSize(block.text, { fontSize: block.style.fontSize ?? 14, monospace: false });
        const box = { left, top, ...boxSize };
        const baseStyle = {
          fontName: "微软雅黑",
          fontSize: 14,
          color: "#000000",
          fillColor: "#ffffff",
          borderColor: "#000000",
          borderWeight: 1,
          ...block.style,
        };
        if (block.equations.length > 0) {
          const inserted = await insertNativeEquationTextBox({
            text: block.text,
            equations: block.equations,
            box,
            baseStyle,
            runs: block.runs,
          });
          top += boxSize.height + 12;
          if (inserted.id) {
            saveLatexForShape(inserted.id, block.equations.map((equation) => equation.latex).join("\n"));
          }
          return inserted.message;
        }
        await addRichTextBox(block.text, box, baseStyle, block.runs);
        top += boxSize.height + 12;
      },
      code: async (block) => {
        const boxSize = estimateTextBoxSize(block.content, { fontSize: 12, monospace: true });
        await addRichTextBox(
          block.content,
          { left, top, ...boxSize },
          getCodeBlockStyle(settings.codeDarkBackground),
          getCodeHighlightRuns(block.content, block.language, settings.codeDarkBackground),
        );
        top += boxSize.height + 12;
      },
      table: async (block) => {
        const tableWidth = Math.min(620, Math.max(240, Math.max(...block.rows.map((row) => row.join("").length)) * 9));
        const tableHeight = Math.max(80, block.rows.length * 28);
        const inserted = await addTableWithFallback(block.rows, { left, top, width: tableWidth, height: tableHeight });
        top += tableHeight + 12;
        return inserted.warning;
      },
      math: async (block) => {
        const boxSize = estimateTextBoxSize(block.content, { fontSize: 18, monospace: false, minWidth: 220, maxWidth: 620 });
        const height = Math.max(54, boxSize.height);
        const inserted = await insertNativeEquationBlock({
          latex: block.content,
          box: { left, top, width: Math.max(360, boxSize.width), height },
          style: {
            fontName: "Cambria Math",
            fontSize: 18,
            color: "#000000",
            align: "center",
          },
        });
        top += height + 12;
        if (inserted.id) {
          saveLatexForShape(inserted.id, block.content);
        }
        return inserted.message;
      },
    });

    if (result.successCount === 0 && result.failures.length > 0) {
      throw new Error(formatMarkdownRenderResult(result));
    }
    return formatMarkdownRenderResult(result);
  }

  async function copyPosition(): Promise<void> {
    const shapes = await getSelectedShapes();
    if (shapes.length === 0) {
      throw new Error("请先选择对象。");
    }
    saveClipboardState({
      ...loadClipboardState(),
      centers: shapes.map((shape) => ({
        left: shape.left + shape.width / 2,
        top: shape.top + shape.height / 2,
      })),
    });
  }

  async function copySize(): Promise<void> {
    const shapes = await getSelectedShapes();
    if (shapes.length === 0) {
      throw new Error("请先选择对象。");
    }
    saveClipboardState({
      ...loadClipboardState(),
      width: shapes[0].width,
      height: shapes[0].height,
    });
  }

  async function pastePosition(): Promise<void> {
    const shapes = await getSelectedShapes();
    const clipboard = loadClipboardState();
    if (!clipboard.centers?.length) {
      throw new Error("还没有复制位置。");
    }

    await updateShapesLayout(
      shapes.slice(0, clipboard.centers.length).map((shape, index) => ({
        id: shape.id,
        width: shape.width,
        height: shape.height,
        left: clipboard.centers![index].left - shape.width / 2,
        top: clipboard.centers![index].top - shape.height / 2,
      })),
    );
  }

  async function updateLabels(): Promise<void> {
    const shapes = await getSelectedShapes();
    if (shapes.length === 0) {
      throw new Error("请先选择要更新的标签文本框。");
    }

    const labels = generateLabels(shapes, settings.labelTemplate, settings.labelIndex);
    await updateTextForShapes(
      labels.map((item) => ({
        id: item.shape.id,
        text: item.label,
        style: {
          fontName: settings.labelFontName,
          fontSize: settings.labelFontSize,
          color: "#000000",
          bold: settings.labelBold,
        },
      })),
    );
    if (settings.labelAutoUpdate) {
      updateSetting("labelIndex", settings.labelIndex + labels.length);
    }
  }

  async function copyStyle(): Promise<void> {
    const style = await copySelectedFormat();
    saveClipboardState({ ...loadClipboardState(), style });
  }

  async function pasteStyle(): Promise<void> {
    const clipboard = loadClipboardState();
    if (!clipboard.style) {
      throw new Error("还没有复制格式。");
    }
    await applyShapeStyleToSelected(clipboard.style);
  }

  async function pasteWidthHeight(kind: "width" | "height" | "both"): Promise<void> {
    const shapes = await getSelectedShapes();
    const clipboard = loadClipboardState();
    if ((kind === "width" || kind === "both") && !clipboard.width) {
      throw new Error("还没有复制宽度。");
    }
    if ((kind === "height" || kind === "both") && !clipboard.height) {
      throw new Error("还没有复制高度。");
    }

    await updateShapesLayout(
      shapes.map((shape) => ({
        id: shape.id,
        left: shape.left,
        top: shape.top,
        width: kind === "height" ? shape.width : clipboard.width ?? shape.width,
        height: kind === "width" ? shape.height : clipboard.height ?? shape.height,
      })),
    );
  }

  return (
    <main>
      <header>
        <h1>SlideSCI for Mac</h1>
        <p>{status}</p>
      </header>

      <Section title="图片自动排列">
        <div className="grid">
          <Field label="排序">
            <select value={settings.sortMode} onChange={(event) => updateSetting("sortMode", event.target.value as SortMode)}>
              <option value="position">根据位置排序</option>
              <option value="selection">根据选择顺序</option>
            </select>
          </Field>
          <Field label="排列">
            <select value={settings.layoutMode} onChange={(event) => updateSetting("layoutMode", event.target.value as LayoutMode)}>
              <option value="columnMaxWidth">列最大宽度占位</option>
              <option value="uniformHeight">统一高度</option>
              <option value="waterfall">统一宽度瀑布流</option>
            </select>
          </Field>
          <Field label="列数量">
            <input type="number" min="1" value={settings.colNum} onChange={(event) => updateSetting("colNum", Number(event.target.value))} />
          </Field>
          <Field label="列间距">
            <input type="number" min="0" value={settings.colSpace} onChange={(event) => updateSetting("colSpace", Number(event.target.value))} />
          </Field>
          <Field label="行间距">
            <input type="number" min="0" value={settings.rowSpace} onChange={(event) => updateSetting("rowSpace", Number(event.target.value))} />
          </Field>
          <Field label="图宽">
            <input value={settings.imgWidth} placeholder="例如 5cm 或 120" onChange={(event) => updateSetting("imgWidth", event.target.value)} />
          </Field>
          <Field label="图高">
            <input value={settings.imgHeight} placeholder="例如 3cm 或 90" onChange={(event) => updateSetting("imgHeight", event.target.value)} />
          </Field>
        </div>
        <button disabled={!canRun} onClick={() => void run("图片排列", arrangeSelectedShapes)}>图片排列</button>
      </Section>

      <Section title="图片标题">
        <div className="grid">
          <Field label="字体">
            <input value={settings.titleFontName} onChange={(event) => updateSetting("titleFontName", event.target.value)} />
          </Field>
          <Field label="字号">
            <input type="number" value={settings.titleFontSize} onChange={(event) => updateSetting("titleFontSize", Number(event.target.value))} />
          </Field>
          <Field label="距离">
            <input type="number" value={settings.titleDistance} onChange={(event) => updateSetting("titleDistance", Number(event.target.value))} />
          </Field>
          <Field label="对齐">
            <select value={settings.titleAlign} onChange={(event) => updateSetting("titleAlign", event.target.value as "left" | "center")}>
              <option value="center">居中</option>
              <option value="left">居左</option>
            </select>
          </Field>
        </div>
        <Field label="标题文本">
          <input value={settings.titleText} onChange={(event) => updateSetting("titleText", event.target.value)} />
        </Field>
        <div className="actions">
          <button disabled={!canRun} onClick={() => void run("添加下标题", () => addTitles("bottom"))}>图片下标题</button>
          <button disabled={!canRun} onClick={() => void run("添加上标题", () => addTitles("top"))}>图片上标题</button>
        </div>
      </Section>

      <Section title="图片标签">
        <div className="grid">
          <Field label="字体">
            <input value={settings.labelFontName} onChange={(event) => updateSetting("labelFontName", event.target.value)} />
          </Field>
          <Field label="字号">
            <input type="number" value={settings.labelFontSize} onChange={(event) => updateSetting("labelFontSize", Number(event.target.value))} />
          </Field>
          <Field label="模板">
            <select value={settings.labelTemplate} onChange={(event) => updateSetting("labelTemplate", event.target.value as AppSettings["labelTemplate"])}>
              {LABEL_TEMPLATES.map((template) => <option key={template} value={template}>{template}</option>)}
            </select>
          </Field>
          <Field label="编号">
            <input type="number" min="1" value={settings.labelIndex} onChange={(event) => updateSetting("labelIndex", Number(event.target.value))} />
          </Field>
          <Field label="X偏移">
            <input type="number" value={settings.labelOffsetX} onChange={(event) => updateSetting("labelOffsetX", Number(event.target.value))} />
          </Field>
          <Field label="Y偏移">
            <input type="number" value={settings.labelOffsetY} onChange={(event) => updateSetting("labelOffsetY", Number(event.target.value))} />
          </Field>
        </div>
        <label className="check"><input type="checkbox" checked={settings.labelBold} onChange={(event) => updateSetting("labelBold", event.target.checked)} />加粗</label>
        <label className="check"><input type="checkbox" checked={settings.labelAutoUpdate} onChange={(event) => updateSetting("labelAutoUpdate", event.target.checked)} />编号自动更新</label>
        <div className="actions">
          <button disabled={!canRun} onClick={() => void run("添加标签", addLabels)}>添加标签</button>
          <button disabled={!canRun} onClick={() => void run("更新标签", updateLabels)}>更新标签</button>
        </div>
      </Section>

      <Section title="内容插入">
        <div className="grid">
          <Field label="代码语言">
            <select value={settings.codeLanguage} onChange={(event) => updateSetting("codeLanguage", event.target.value)}>
              {CODE_LANGUAGES.map((language) => <option key={language} value={language}>{CODE_LANGUAGE_LABELS[language] ?? language}</option>)}
            </select>
          </Field>
          <label className="check"><input type="checkbox" checked={settings.codeDarkBackground} onChange={(event) => updateSetting("codeDarkBackground", event.target.checked)} />代码黑色背景</label>
        </div>
        <textarea value={code} placeholder="粘贴代码" onChange={(event) => setCode(event.target.value)} />
        <button disabled={!canRun} onClick={() => void run("插入代码块", insertCodeBlock)}>插入代码块</button>
        <textarea value={markdown} placeholder="粘贴 Markdown" onChange={(event) => setMarkdown(event.target.value)} />
        <button disabled={!canRun} onClick={() => void run("插入 Markdown", insertMarkdown)}>插入 Markdown</button>
        <textarea value={latex} placeholder="输入 LaTeX，例如 \\frac{a}{b}" onChange={(event) => setLatex(event.target.value)} />
        <div className="actions">
          <button disabled={!canRun} onClick={() => void run("插入 LaTeX 原生公式", insertLatexNative)}>插入 LaTeX 原生公式</button>
          <button disabled={!canRun} onClick={() => void run("插入 LaTeX 图片", insertLatexImage)}>插入 LaTeX 图片</button>
          <button disabled={!canRun} onClick={() => void run("读取选中 LaTeX", loadLatexFromSelection)}>读取选中 LaTeX</button>
        </div>
      </Section>

      <Section title="格式工具">
        <div className="toolRows">
          <div className="toolRow">
            <button className="copyButton" disabled={!canRun} onClick={() => void run("复制位置", copyPosition)}>复制位置</button>
            <button className="pasteButton" disabled={!canRun} onClick={() => void run("粘贴位置", pastePosition)}>粘贴位置</button>
          </div>
          <div className="toolRow">
            <button className="copyButton" disabled={!canRun} onClick={() => void run("复制宽高", copySize)}>复制宽高</button>
            <button className="pasteButton" disabled={!canRun} onClick={() => void run("粘贴宽度", () => pasteWidthHeight("width"))}>粘贴宽度</button>
            <button className="pasteButton" disabled={!canRun} onClick={() => void run("粘贴高度", () => pasteWidthHeight("height"))}>粘贴高度</button>
            <button className="pasteButton" disabled={!canRun} onClick={() => void run("粘贴宽高", () => pasteWidthHeight("both"))}>粘贴宽高</button>
          </div>
          <div className="toolRow">
            <button className="copyButton" disabled={!canRun} onClick={() => void run("复制格式", copyStyle)}>复制格式</button>
            <button className="pasteButton" disabled={!canRun} onClick={() => void run("粘贴格式", pasteStyle)}>粘贴格式</button>
          </div>
        </div>
        <button className="secondary" onClick={() => setSettings(defaultSettings)}>恢复默认设置</button>
      </Section>
    </main>
  );
}
