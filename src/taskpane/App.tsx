import { useMemo, useState } from "react";
import { arrangeShapes } from "../lib/layout";
import { LABEL_TEMPLATES } from "../lib/labels";
import { generateLabels } from "../lib/labels";
import { normalizeLatexInput } from "../lib/latex";
import { createMarkdownSingleColumnLayout } from "../lib/markdownLayout";
import { estimateTextBoxSize } from "../lib/textMetrics";
import { parseLengthToPt } from "../lib/units";
import type { Box, LayoutMode, NativeEquationRun, SortMode, TextRun, TextStyle, TitlePlacement } from "../lib/types";
import { getCodeBlockStyle, getCodeHighlightRuns, CODE_LANGUAGES, CODE_LANGUAGE_LABELS } from "../services/codeBlock";
import { convertLatexToPngBase64 } from "../services/latexSvg";
import { formatMarkdownRenderResult, markdownToRenderBlocks, renderMarkdownBlocks } from "../services/markdownRender";
import {
  buildShapeRangeEquationRequest,
  convertEquationRuns,
  convertShapeRangesToNativeEquations,
  type NativeEquationHelperStrategy,
  formatEquationConversionSummary,
} from "../services/nativeEquation";
import {
  addLatexImage,
  addRichTextBox,
  addTableWithFallback,
  addTextBox,
  applyShapeStyleToSelected,
  copySelectedFormat,
  deleteShapes,
  getPowerPointHostCapabilities,
  getSelectedShapeIds,
  getSelectedLatexMetadata,
  getSelectedShapes,
  getSlidePageSize,
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

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

const legacyTableWarning = "实验性原生表格不可用，已退文本框网格。";

function strategyLabel(strategy?: NativeEquationHelperStrategy): string {
  return strategy ? `（${strategy}）` : "";
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

    return insertNativeEquationBlockFromText(
      normalizedLatex,
      { left: 160, top: 120, width: 500, height: 120 },
      "插入 LaTeX 原生公式完成",
    );
  }

  async function insertLatexImage(): Promise<string | void> {
    return insertLatexImageForSource(latex, { left: 160, top: 120, width: 520, height: 180 }, "插入 LaTeX 图片完成");
  }

  async function insertLatexImageForSource(source: string, box: Box, successMessage: string): Promise<string | void> {
    const image = await convertLatexToPngBase64(source);
    const width = Math.min(520, Math.max(120, image.width * 0.75));
    const height = Math.min(180, Math.max(48, image.height * 0.75));
    const inserted = await addLatexImage(
      image.base64,
      {
        left: box.left + Math.max(0, (box.width - width) / 2),
        top: box.top,
        width: Math.min(box.width, width),
        height: Math.min(Math.max(box.height, 48), height),
      },
      image.latex,
    );
    saveLatexForShape(inserted.id, image.latex);
    if (inserted.warning) {
      return `${successMessage}：${inserted.warning}`;
    }
    return successMessage;
  }

  function rememberLatexSource(originalShapeId: string, currentShapeId: string, latexSource: string): void {
    saveLatexForShape(originalShapeId, latexSource);
    if (currentShapeId && currentShapeId !== originalShapeId) {
      saveLatexForShape(currentShapeId, latexSource);
    }
  }

  async function getSelectedShapeIdsSafe(): Promise<string[]> {
    try {
      return await getSelectedShapeIds();
    } catch {
      return [];
    }
  }

  async function deleteEquationWorkShapes(ids: string[]): Promise<void> {
    const deduped = [...new Set(ids.filter(Boolean))];
    if (deduped.length === 0) {
      return;
    }
    try {
      await deleteShapes(deduped);
    } catch {
      // Best effort only. The user-facing flow should continue with the restored shape.
    }
  }

  function formatLegacyEquationFailureMessage(detail: string): string {
    return `当前 PowerPoint 版本缺少自动原生公式能力，GUI 自动化也失败：${detail}`;
  }

  async function insertNativeEquationBlockFromText(
    normalizedLatex: string,
    box: Box,
    successMessage: string,
    hostCapabilities?: Awaited<ReturnType<typeof getPowerPointHostCapabilities>>,
  ): Promise<string | void> {
    const capabilities = hostCapabilities ?? await getPowerPointHostCapabilities();
    const textBoxStyle = {
      fontName: "Cambria Math",
      fontSize: 18,
      color: "#000000",
      align: "center" as const,
    };

    if (!capabilities.textRangeSelection) {
      const guiRequest = buildShapeRangeEquationRequest(
        "",
        normalizedLatex,
        [{ start: 0, length: normalizedLatex.length, latex: normalizedLatex, display: true }],
        "block",
      );
      const shapeId = await addTextBox(guiRequest.workingText, box, textBoxStyle);
      try {
        await selectShapes([shapeId]);
        const response = await convertShapeRangesToNativeEquations({ ...guiRequest, shapeId });
        const currentShapeId = (await getSelectedShapeIdsSafe())[0] || shapeId;
        rememberLatexSource(shapeId, currentShapeId, normalizedLatex);
        return `${successMessage}${strategyLabel(response.strategyUsed)}`;
      } catch (error) {
        const selectedIds = await getSelectedShapeIdsSafe();
        await deleteEquationWorkShapes([shapeId, ...selectedIds]);
        const failureMessage = formatLegacyEquationFailureMessage(errorMessage(error));

        if (!settings.allowEquationImageFallback) {
          const restoredShapeId = await addTextBox(normalizedLatex, box, textBoxStyle);
          rememberLatexSource(restoredShapeId, restoredShapeId, normalizedLatex);
          throw new Error(failureMessage);
        }

        const warning = await insertLatexImageForSource(
          normalizedLatex,
          box,
          "当前 PowerPoint 版本缺少自动原生公式能力，GUI 自动化失败，已按设置降级为图片",
        );
        return `${warning}：${errorMessage(error)}`;
      }
    }

    const shapeId = await addTextBox(normalizedLatex, box, {
      ...textBoxStyle,
    });
    const summary = await convertEquationRuns(shapeId, [
      { start: 0, length: normalizedLatex.length, latex: normalizedLatex, display: true },
    ]);
    rememberLatexSource(shapeId, summary.shapeId, normalizedLatex);
    if (summary.fallbackCount === 0) {
      return `${successMessage}${strategyLabel(summary.strategiesUsed[0])}`;
    }

    const failureMessage = summary.messages.at(-1) ?? "原生公式转换失败。";
    if (!settings.allowEquationImageFallback) {
      throw new Error(failureMessage);
    }

    await deleteShapes([summary.shapeId || shapeId]);
    const warning = await insertLatexImageForSource(normalizedLatex, box, "原生公式 helper 不可用，已按设置降级为图片");
    return `${warning}：${failureMessage}`;
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

    const hostCapabilities = await getPowerPointHostCapabilities();
    const pageSize = await getSlidePageSize();
    const layout = createMarkdownSingleColumnLayout(blocks, pageSize);
    let nativeTableAvailable = hostCapabilities.experimentalNativeTable;
    let nativeTableWarningShown = false;

    async function insertInlineEquationImages(
      text: string,
      equations: NativeEquationRun[],
      box: Box,
      fontSize: number,
    ): Promise<number> {
      let count = 0;
      for (const equation of equations) {
        const image = await convertLatexToPngBase64(equation.latex);
        const before = text.slice(0, equation.start);
        const lines = before.split(/\r?\n/);
        const lineIndex = Math.max(0, lines.length - 1);
        const linePrefix = lines[lineIndex] ?? "";
        const prefixWidth = estimateTextBoxSize(linePrefix, {
          fontSize,
          monospace: false,
          minWidth: 0,
          maxWidth: box.width,
          paddingX: 0,
          paddingY: 0,
        }).width;
        const width = Math.min(180, Math.max(18, image.width * 0.42));
        const height = Math.min(48, Math.max(16, image.height * 0.42));
        const left = Math.min(box.left + box.width - width, box.left + 9 + prefixWidth);
        const top = box.top + 7 + lineIndex * fontSize * 1.35;
        const inserted = await addLatexImage(image.base64, { left, top, width, height }, image.latex);
        saveLatexForShape(inserted.id, image.latex);
        count += 1;
      }
      return count;
    }

    async function insertTextBlockWithOptionalEquationFallback(
      text: string,
      equations: NativeEquationRun[],
      box: Box,
      baseStyle: TextStyle,
      runs: TextRun[],
      latexSource: string,
    ): Promise<string | void> {
      if (!hostCapabilities.textRangeSelection && equations.length > 0) {
        const guiRequest = buildShapeRangeEquationRequest("", text, equations, "inline");
        const shapeId = await addRichTextBox(guiRequest.workingText, box, baseStyle, runs);

        try {
          await selectShapes([shapeId]);
          const response = await convertShapeRangesToNativeEquations({ ...guiRequest, shapeId });
          const currentShapeId = (await getSelectedShapeIdsSafe())[0] || shapeId;
          rememberLatexSource(shapeId, currentShapeId, latexSource);
          return formatEquationConversionSummary({
            shapeId: currentShapeId,
            strategy: "helper-gui",
            strategiesUsed: response.strategyUsed ? [response.strategyUsed] : [],
            nativeCount: equations.length,
            fallbackCount: 0,
            messages: response.message ? [response.message] : [],
            remainingEquations: [],
          });
        } catch (error) {
          const selectedIds = await getSelectedShapeIdsSafe();
          await deleteEquationWorkShapes([shapeId, ...selectedIds]);
          const restoredShapeId = await addRichTextBox(text, box, baseStyle, runs);
          rememberLatexSource(restoredShapeId, restoredShapeId, latexSource);
          const failureMessage = formatLegacyEquationFailureMessage(errorMessage(error));

          if (!settings.allowEquationImageFallback) {
            throw new Error(failureMessage);
          }

          const count = await insertInlineEquationImages(text, equations, box, baseStyle.fontSize ?? 14);
          return `当前 PowerPoint 版本缺少自动原生公式能力，已将 ${count} 个公式降级为图片：${errorMessage(error)}`;
        }
      }

      const shapeId = await addRichTextBox(text, box, baseStyle, runs);
      if (equations.length === 0) {
        return undefined;
      }

      const summary = await convertEquationRuns(shapeId, equations);
      rememberLatexSource(shapeId, summary.shapeId, latexSource);
      if (summary.fallbackCount === 0) {
        return formatEquationConversionSummary(summary);
      }

      const failureMessage = summary.messages.at(-1) ?? "原生公式转换失败。";
      if (!settings.allowEquationImageFallback) {
        throw new Error(failureMessage);
      }

      const count = await insertInlineEquationImages(text, summary.remainingEquations, box, baseStyle.fontSize ?? 14);
      return `原生公式成功 ${summary.nativeCount} 个，剩余 ${count} 个已按设置降级为图片：${failureMessage}`;
    }

    const result = await renderMarkdownBlocks(blocks, {
      text: async (block, index) => {
        const box = layout[index]?.box;
        if (!box) {
          throw new Error("Markdown 布局缺少文本模块位置。");
        }
        const baseStyle = {
          fontName: "微软雅黑",
          fontSize: 14,
          color: "#000000",
        };
        return insertTextBlockWithOptionalEquationFallback(
          block.text,
          block.equations,
          box,
          baseStyle,
          block.runs,
          block.equations.map((equation) => equation.latex).join("\n"),
        );
      },
      quote: async (block, index) => {
        const box = layout[index]?.box;
        if (!box) {
          throw new Error("Markdown 布局缺少引用模块位置。");
        }
        const baseStyle = {
          fontName: "微软雅黑",
          fontSize: 14,
          color: "#000000",
          fillColor: "#ffffff",
          borderColor: "#000000",
          borderWeight: 1,
          ...block.style,
        };
        return insertTextBlockWithOptionalEquationFallback(
          block.text,
          block.equations,
          box,
          baseStyle,
          block.runs,
          block.equations.map((equation) => equation.latex).join("\n"),
        );
      },
      code: async (block, index) => {
        const box = layout[index]?.box;
        if (!box) {
          throw new Error("Markdown 布局缺少代码模块位置。");
        }
        await addRichTextBox(
          block.content,
          box,
          getCodeBlockStyle(settings.codeDarkBackground),
          getCodeHighlightRuns(block.content, block.language, settings.codeDarkBackground),
        );
      },
      table: async (block, index) => {
        const box = layout[index]?.box;
        if (!box) {
          throw new Error("Markdown 布局缺少表格模块位置。");
        }

        if (!nativeTableAvailable) {
          await addTableWithFallback(block.rows, box, { skipNative: true });
          if (nativeTableWarningShown) {
            return undefined;
          }
          nativeTableWarningShown = true;
          return legacyTableWarning;
        }

        const inserted = await addTableWithFallback(block.rows, box);
        if (inserted.warningCode === "nativeTableUnsupported") {
          nativeTableAvailable = false;
          if (nativeTableWarningShown) {
            return undefined;
          }
          nativeTableWarningShown = true;
          return legacyTableWarning;
        }
        return inserted.warning;
      },
      math: async (block, index) => {
        const box = layout[index]?.box;
        if (!box) {
          throw new Error("Markdown 布局缺少公式模块位置。");
        }
        return insertNativeEquationBlockFromText(block.content, box, "块级公式已转换为原生公式", hostCapabilities);
      },
    });

    if (result.successCount === 0 && result.failures.length > 0) {
      throw new Error(formatMarkdownRenderResult(result));
    }
    const message = formatMarkdownRenderResult(result);
    return pageSize.warning ? `${message} ${pageSize.warning}` : message;
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
          <label className="check"><input type="checkbox" checked={settings.allowEquationImageFallback} onChange={(event) => updateSetting("allowEquationImageFallback", event.target.checked)} />允许公式降级为图片</label>
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
