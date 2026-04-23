import type { Box, ShapeLayout, SlideShape, TextRun, TextStyle } from "../lib/types";

type OfficeShape = Record<string, any>;

export interface SlidePageSize {
  width: number;
  height: number;
  warning?: string;
}

const fallbackSlidePageSize: SlidePageSize = {
  width: 960,
  height: 540,
  warning: "当前 PowerPoint API 未返回页面尺寸，已按 16:9 宽屏默认尺寸排版。",
};

function ensurePowerPoint(): any {
  const powerpoint = (globalThis as any).PowerPoint;
  if (!powerpoint?.run) {
    throw new Error("当前环境未检测到 PowerPoint Office.js。请在 PowerPoint 加载项中打开。");
  }
  return powerpoint;
}

function toSlideShape(shape: OfficeShape): SlideShape {
  return {
    id: String(shape.id ?? shape.name),
    name: shape.name,
    type: String(shape.type ?? ""),
    left: Number(shape.left ?? 0),
    top: Number(shape.top ?? 0),
    width: Number(shape.width ?? 0),
    height: Number(shape.height ?? 0),
  };
}

function normalizeHex(color: string | undefined): string | undefined {
  if (!color) {
    return undefined;
  }
  return color.startsWith("#") ? color : `#${color}`;
}

function applyShapeTextStyle(shape: OfficeShape, style: TextStyle): void {
  const textFrame = shape.textFrame;
  if (!textFrame) {
    return;
  }

  const textRange = textFrame.textRange;
  if (textRange?.font) {
    if (style.fontName || style.fontFamily) {
      textRange.font.name = style.fontName || style.fontFamily;
    }
    if (style.fontSize) {
      textRange.font.size = style.fontSize;
    }
    if (style.color) {
      textRange.font.color = normalizeHex(style.color);
    }
    if (typeof style.bold === "boolean") {
      textRange.font.bold = style.bold;
    }
    if (typeof style.italic === "boolean") {
      textRange.font.italic = style.italic;
    }
    if (typeof style.underline === "boolean") {
      textRange.font.underline = style.underline;
    }
    if (typeof style.subscript === "boolean") {
      textRange.font.subscript = style.subscript;
    }
    if (typeof style.superscript === "boolean") {
      textRange.font.superscript = style.superscript;
    }
  }

  if (textRange?.paragraphFormat && style.align) {
    textRange.paragraphFormat.horizontalAlignment = style.align === "center" ? "Center" : "Left";
  }
}

function applyShapeStyle(shape: OfficeShape, style: TextStyle): void {
  if (style.fillColor && shape.fill) {
    try {
      if (typeof shape.fill.setSolidColor === "function") {
        shape.fill.setSolidColor(normalizeHex(style.fillColor));
      } else {
        shape.fill.color = normalizeHex(style.fillColor);
        shape.fill.transparency = 0;
      }
      if ("visible" in shape.fill) {
        shape.fill.visible = true;
      }
    } catch {
      // Some Office.js builds expose fill as read-only for selected shape proxies.
    }
  }
  if (style.borderColor && shape.lineFormat) {
    try {
      shape.lineFormat.color = normalizeHex(style.borderColor);
      shape.lineFormat.weight = style.borderWeight ?? 1;
    } catch {
      // Keep applying text style even if line formatting is unavailable.
    }
  }
  applyShapeTextStyle(shape, style);
}

async function withCurrentSlide<T>(callback: (context: any, slide: any) => Promise<T> | T): Promise<T> {
  const PowerPointApi = ensurePowerPoint();
  return PowerPointApi.run(async (context: any) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items");
    await context.sync();
    const slide = slides.items?.[0];
    if (!slide) {
      throw new Error("未获取到当前幻灯片，请先选中或打开一个页面。");
    }
    return callback(context, slide);
  });
}

export async function getSelectedShapes(): Promise<SlideShape[]> {
  const PowerPointApi = ensurePowerPoint();
  return PowerPointApi.run(async (context: any) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,name,type,left,top,width,height");
    await context.sync();
    return (shapes.items ?? []).map(toSlideShape);
  });
}

export async function getSelectedShapeIds(): Promise<string[]> {
  const PowerPointApi = ensurePowerPoint();
  return PowerPointApi.run(async (context: any) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await context.sync();
    return (shapes.items ?? []).map((shape: OfficeShape) => String(shape.id));
  });
}

export async function getSlidePageSize(): Promise<SlidePageSize> {
  const PowerPointApi = ensurePowerPoint();
  return PowerPointApi.run(async (context: any) => {
    const pageSetup = context.presentation?.pageSetup;
    if (!pageSetup?.load) {
      return fallbackSlidePageSize;
    }

    try {
      pageSetup.load("slideWidth,slideHeight");
      await context.sync();
      const width = Number(pageSetup.slideWidth);
      const height = Number(pageSetup.slideHeight);
      if (Number.isFinite(width) && width > 0 && Number.isFinite(height) && height > 0) {
        return { width, height };
      }
      return fallbackSlidePageSize;
    } catch {
      return fallbackSlidePageSize;
    }
  });
}

export async function addTextBox(text: string, box: Box, style: TextStyle = {}): Promise<string> {
  return withCurrentSlide(async (context, slide) => {
    if (!slide.shapes?.addTextBox) {
      throw new Error("当前 PowerPoint 版本不支持添加文本框。");
    }
    const shape = slide.shapes.addTextBox(text, box);
    applyShapeStyle(shape, style);
    shape.load("id");
    await context.sync();
    return String(shape.id);
  });
}

function applyTextRunToRange(range: OfficeShape, run: TextRun): void {
  if (!range?.font) {
    return;
  }
  const style = run.style;
  if (style.fontName || style.fontFamily) {
    range.font.name = style.fontName || style.fontFamily;
  }
  if (style.fontSize) {
    range.font.size = style.fontSize;
  }
  if (style.color) {
    range.font.color = normalizeHex(style.color);
  }
  if (typeof style.bold === "boolean") {
    range.font.bold = style.bold;
  }
  if (typeof style.italic === "boolean") {
    range.font.italic = style.italic;
  }
  if (typeof style.underline === "boolean") {
    range.font.underline = style.underline;
  }
  if (typeof style.subscript === "boolean") {
    range.font.subscript = style.subscript;
  }
  if (typeof style.superscript === "boolean") {
    range.font.superscript = style.superscript;
  }
}

export async function applyTextRuns(shapeId: string, runs: TextRun[]): Promise<boolean> {
  if (runs.length === 0) {
    return true;
  }

  return withCurrentSlide(async (context, slide) => {
    const shape = slide.shapes.getItem(shapeId);
    const textRange = shape?.textFrame?.textRange;
    if (!textRange?.getSubstring) {
      return false;
    }

    for (const run of runs) {
      if (run.length <= 0) {
        continue;
      }
      const range = textRange.getSubstring(run.start, run.length);
      applyTextRunToRange(range, run);
    }
    await context.sync();
    return true;
  });
}

export async function selectTextRange(shapeId: string, start: number, length: number): Promise<void> {
  await withCurrentSlide(async (context, slide) => {
    const shape = slide.shapes.getItem(shapeId);
    const textRange = shape?.textFrame?.textRange;
    if (!textRange?.getSubstring) {
      throw new Error("当前 PowerPoint 版本不支持选择文本范围。");
    }
    const range = textRange.getSubstring(start, length);
    if (!range?.select) {
      throw new Error("当前 PowerPoint 版本不支持选择文本范围。");
    }
    range.select();
    await context.sync();
  });
}

export async function addRichTextBox(
  text: string,
  box: Box,
  baseStyle: TextStyle = {},
  runs: TextRun[] = [],
): Promise<string> {
  const shapeId = await addTextBox(text, box, baseStyle);
  await applyTextRuns(shapeId, runs);
  return shapeId;
}

export async function addSvgPicture(svg: string, box: Box): Promise<string> {
  return addBase64Picture(btoa(unescape(encodeURIComponent(svg))), box);
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function writeLatexMetadata(shape: OfficeShape, latex?: string): void {
  if (!latex) {
    return;
  }
  try {
    shape.altTextTitle = "SlideSCI LaTeX";
    shape.altTextDescription = latex;
  } catch {
    // Alt text is best-effort across Office.js builds.
  }
  try {
    if (shape.tags?.add) {
      shape.tags.add("slidesci.latex", latex);
    }
  } catch {
    // Tags are optional in older PowerPoint clients.
  }
}

export async function addBase64Picture(base64: string, box: Box, metadata?: { latex?: string }): Promise<string> {
  return withCurrentSlide(async (context, slide) => {
    if (!slide.shapes?.addPicture) {
      throw new Error("当前 PowerPoint 版本不支持插入图片 API。");
    }

    const shape = slide.shapes.addPicture(base64, box);
    writeLatexMetadata(shape, metadata?.latex);
    shape.load("id");
    await context.sync();
    return String(shape.id);
  });
}

export interface LatexImageInsertResult {
  id: string;
  mode: "picture" | "shapeFill" | "textFallback";
  warning?: string;
}

async function addLatexShapeFill(base64: string, box: Box, latex: string, priorErrors: string[] = []): Promise<LatexImageInsertResult> {
  return withCurrentSlide(async (context, slide) => {
    if (typeof slide.shapes?.addGeometricShape !== "function") {
      throw new Error("addGeometricShape 不可用");
    }

    const shape = slide.shapes.addGeometricShape("Rectangle", box);
    if (!shape.fill?.setImage) {
      throw new Error("shape.fill.setImage 不可用");
    }
    shape.fill.setImage(base64);
    if (shape.lineFormat) {
      shape.lineFormat.color = "#ffffff";
      shape.lineFormat.weight = 0;
    }
    writeLatexMetadata(shape, latex);
    shape.load("id");
    await context.sync();
    return {
      id: String(shape.id),
      mode: "shapeFill",
      warning: priorErrors.length > 0 ? `${priorErrors.join("；")}，已降级为图片填充。` : undefined,
    };
  });
}

async function addLatexPicture(base64: string, box: Box, latex: string, priorErrors: string[]): Promise<LatexImageInsertResult> {
  return withCurrentSlide(async (context, slide) => {
    if (typeof slide.shapes?.addPicture !== "function") {
      throw new Error("addPicture 不可用");
    }

    const shape = slide.shapes.addPicture(base64, box);
    writeLatexMetadata(shape, latex);
    shape.load("id");
    await context.sync();
    return {
      id: String(shape.id),
      mode: "picture",
      warning: `${priorErrors.join("；")}，已降级为 addPicture 图片。`,
    };
  });
}

async function addLatexTextFallback(box: Box, latex: string, priorErrors: string[]): Promise<LatexImageInsertResult> {
  return withCurrentSlide(async (context, slide) => {
    if (typeof slide.shapes?.addTextBox !== "function") {
      throw new Error("addTextBox 不可用");
    }

    const shape = slide.shapes.addTextBox(latex, box);
    applyShapeStyle(shape, {
      fontName: "Consolas",
      fontSize: 14,
      color: "#000000",
      fillColor: "#fff8dc",
      borderColor: "#c00000",
      borderWeight: 1,
    });
    writeLatexMetadata(shape, latex);
    shape.load("id");
    await context.sync();
    return {
      id: String(shape.id),
      mode: "textFallback",
      warning: `${priorErrors.join("；")}，已降级为 LaTeX 源码文本框。`,
    };
  });
}

export async function addLatexImage(base64: string, box: Box, latex: string): Promise<LatexImageInsertResult> {
  const errors: string[] = [];

  try {
    return await addLatexShapeFill(base64, box, latex);
  } catch (error) {
    errors.push(`fill.setImage 失败：${errorMessage(error)}`);
  }

  try {
    return await addLatexPicture(base64, box, latex, errors);
  } catch (error) {
    errors.push(`addPicture 失败：${errorMessage(error)}`);
  }

  try {
    return await addLatexTextFallback(box, latex, errors);
  } catch (error) {
    errors.push(`文本降级失败：${errorMessage(error)}`);
  }

  throw new Error(`LaTeX 图片插入失败：${errors.join("；")}`);
}

export interface TableInsertResult {
  ids: string[];
  mode: "nativeTable" | "textGrid";
  warning?: string;
  warningCode?: "nativeTableUnsupported";
}

export async function addTable(rows: string[][], box: Box): Promise<string> {
  const result = await addTableWithFallback(rows, box);
  return result.ids[0];
}

function normalizeTableRows(rows: string[][]): string[][] {
  if (rows.length === 0) {
    return [[""]];
  }
  const columnCount = Math.max(...rows.map((row) => row.length));
  return rows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
}

function tableBoxOptions(box: Box): Record<string, number> {
  return {
    left: Math.round(box.left),
    top: Math.round(box.top),
    width: Math.max(1, Math.round(box.width)),
    height: Math.max(1, Math.round(box.height)),
  };
}

async function fillNativeTableCells(context: any, shape: OfficeShape, rows: string[][]): Promise<void> {
  if (typeof shape.getTable !== "function") {
    throw new Error("shape.getTable 不可用");
  }
  const table = shape.getTable();
  if (typeof table?.getCellOrNullObject !== "function") {
    throw new Error("table.getCellOrNullObject 不可用");
  }

  for (const [rowIndex, row] of rows.entries()) {
    for (const [columnIndex, value] of row.entries()) {
      const cell = table.getCellOrNullObject(rowIndex, columnIndex);
      cell.text = value;
    }
  }
  await context.sync();
}

async function deleteShapeBestEffort(context: any, shape: OfficeShape): Promise<void> {
  try {
    if (typeof shape.delete === "function") {
      shape.delete();
      await context.sync();
    }
  } catch {
    // If cleanup is unavailable, leave the native error for the caller.
  }
}

export async function addNativeTableByCells(rows: string[][], box: Box): Promise<TableInsertResult> {
  return withCurrentSlide(async (context, slide) => {
    if (typeof slide.shapes?.addTable !== "function") {
      throw new Error("addTable 不可用");
    }
    const normalizedRows = normalizeTableRows(rows);
    const columnCount = normalizedRows[0]?.length ?? 0;
    const shape = slide.shapes.addTable(normalizedRows.length, columnCount, tableBoxOptions(box));
    shape.load("id");
    await context.sync();
    try {
      await fillNativeTableCells(context, shape, normalizedRows);
    } catch (error) {
      await deleteShapeBestEffort(context, shape);
      throw new Error(`原生表格已创建但填值失败：${errorMessage(error)}`);
    }
    return { ids: [String(shape.id)], mode: "nativeTable" };
  });
}

async function addNativeTableWithValues(rows: string[][], box: Box): Promise<TableInsertResult> {
  return withCurrentSlide(async (context, slide) => {
    if (typeof slide.shapes?.addTable !== "function") {
      throw new Error("addTable 不可用");
    }
    const normalizedRows = normalizeTableRows(rows);
    const columnCount = normalizedRows[0]?.length ?? 0;
    const options = { ...tableBoxOptions(box), values: normalizedRows };
    const shape = slide.shapes.addTable(normalizedRows.length, columnCount, options);
    shape.load("id");
    await context.sync();
    return { ids: [String(shape.id)], mode: "nativeTable" };
  });
}

async function addTableTextGrid(
  rows: string[][],
  box: Box,
  priorErrors: string[],
  warning?: string,
  warningCode?: "nativeTableUnsupported",
): Promise<TableInsertResult> {
  return withCurrentSlide(async (context, slide) => {
    if (typeof slide.shapes?.addTextBox !== "function") {
      throw new Error("addTextBox 不可用");
    }

    const normalizedRows = normalizeTableRows(rows);
    const columnCount = normalizedRows[0]?.length ?? 0;
    const cellShapes: OfficeShape[] = [];
    const cellWidth = box.width / columnCount;
    const cellHeight = box.height / rows.length;
    normalizedRows.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        const shape = slide.shapes.addTextBox(cell, {
          left: box.left + columnIndex * cellWidth,
          top: box.top + rowIndex * cellHeight,
          width: cellWidth,
          height: cellHeight,
        });
        applyShapeStyle(shape, {
          fontName: "微软雅黑",
          fontSize: rowIndex === 0 ? 12 : 11,
          bold: rowIndex === 0,
          color: "#000000",
          fillColor: rowIndex === 0 ? "#f2f2f2" : "#ffffff",
          borderColor: "#000000",
          borderWeight: 1,
        });
        shape.load("id");
        cellShapes.push(shape);
      });
    });
    await context.sync();
    return {
      ids: cellShapes.map((shape) => String(shape.id)),
      mode: "textGrid",
      warning: warning ?? (priorErrors.length > 0 ? `${priorErrors.join("；")}，已降级为文本框网格。` : undefined),
      warningCode,
    };
  });
}

function isInvalidArgumentError(error: unknown): boolean {
  return /InvalidArgument/i.test(errorMessage(error));
}

export async function addTableWithFallback(
  rows: string[][],
  box: Box,
  options: { skipNative?: boolean } = {},
): Promise<TableInsertResult> {
  if (options.skipNative) {
    return addTableTextGrid(rows, box, [], undefined);
  }

  const errors: string[] = [];
  let nativeUnsupported = false;
  try {
    return await addNativeTableByCells(rows, box);
  } catch (error) {
    nativeUnsupported = nativeUnsupported || isInvalidArgumentError(error);
    errors.push(`空原生表格逐格填值失败：${errorMessage(error)}`);
  }

  try {
    return await addNativeTableWithValues(rows, box);
  } catch (error) {
    nativeUnsupported = nativeUnsupported || isInvalidArgumentError(error);
    errors.push(`原生表格 values 含尺寸失败：${errorMessage(error)}`);
  }

  try {
    return await addTableTextGrid(
      rows,
      box,
      errors,
      nativeUnsupported ? "当前 PowerPoint 版本不支持原生表格，已降级为文本框网格。" : undefined,
      nativeUnsupported ? "nativeTableUnsupported" : undefined,
    );
  } catch (error) {
    errors.push(`文本框网格失败：${errorMessage(error)}`);
  }

  throw new Error(`表格插入失败：${errors.join("；")}`);
}

export async function updateShapesLayout(layouts: ShapeLayout[]): Promise<void> {
  await withCurrentSlide(async (context, slide) => {
    for (const layout of layouts) {
      const shape = slide.shapes.getItem(layout.id);
      shape.left = layout.left;
      shape.top = layout.top;
      shape.width = layout.width;
      shape.height = layout.height;
    }
    await context.sync();
  });
}

export async function updateTextForShapes(
  updates: Array<{ id: string; text: string; style?: TextStyle }>,
): Promise<void> {
  await withCurrentSlide(async (context, slide) => {
    for (const update of updates) {
      const shape = slide.shapes.getItem(update.id);
      if (!shape?.textFrame?.textRange) {
        continue;
      }
      shape.textFrame.textRange.text = update.text;
      if (update.style) {
        applyShapeStyle(shape, update.style);
      }
    }
    await context.sync();
  });
}

export async function deleteShapes(ids: string[]): Promise<void> {
  if (ids.length === 0) {
    return;
  }

  await withCurrentSlide(async (context, slide) => {
    for (const id of ids) {
      const shape = slide.shapes.getItem(id);
      if (typeof shape?.delete === "function") {
        shape.delete();
      }
    }
    await context.sync();
  });
}

export async function selectShapes(ids: string[]): Promise<void> {
  if (ids.length === 0) {
    return;
  }

  await withCurrentSlide(async (context, slide) => {
    for (const id of ids) {
      const shape = slide.shapes.getItem(id);
      if (shape?.select) {
        shape.select();
      }
    }
    await context.sync();
  });
}

export interface CopiedShapeStyle {
  text?: TextStyle;
  fillColor?: string;
  borderColor?: string;
  borderWeight?: number;
}

function safeValue<T>(getter: () => T): T | undefined {
  try {
    return getter();
  } catch {
    return undefined;
  }
}

export async function getSelectedShapeStyle(): Promise<CopiedShapeStyle> {
  const PowerPointApi = ensurePowerPoint();
  return PowerPointApi.run(async (context: any) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,textFrame/textRange/font/name,textFrame/textRange/font/size,textFrame/textRange/font/color,textFrame/textRange/font/bold,textFrame/textRange/font/italic,textFrame/textRange/font/underline,fill/color,lineFormat/color,lineFormat/weight");
    await context.sync();
    const shape = shapes.items?.[0];
    if (!shape) {
      throw new Error("请先选择一个要复制格式的对象。");
    }

    const font = shape.textFrame?.textRange?.font;
    return {
      text: {
        fontName: safeValue(() => font?.name),
        fontSize: safeValue(() => font?.size),
        color: normalizeHex(safeValue(() => font?.color)),
        bold: safeValue(() => font?.bold),
        italic: safeValue(() => font?.italic),
        underline: safeValue(() => font?.underline),
      },
      fillColor: normalizeHex(safeValue(() => shape.fill?.color)),
      borderColor: normalizeHex(safeValue(() => shape.lineFormat?.color)),
      borderWeight: safeValue(() => shape.lineFormat?.weight),
    };
  });
}

export const copySelectedFormat = getSelectedShapeStyle;

export async function applyShapeStyleToSelected(style: CopiedShapeStyle): Promise<void> {
  const PowerPointApi = ensurePowerPoint();
  await PowerPointApi.run(async (context: any) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await context.sync();
    if (!shapes.items?.length) {
      throw new Error("请先选择要粘贴格式的对象。");
    }

    for (const shape of shapes.items) {
      try {
        applyShapeStyle(shape, {
          ...(style.text ?? {}),
          fillColor: style.fillColor,
          borderColor: style.borderColor,
          borderWeight: style.borderWeight,
        });
        await context.sync();
      } catch {
        // Continue with remaining selected shapes; unsupported properties are handled best-effort.
      }
    }
  });
}

export interface SelectedLatexMetadata {
  shapeId: string;
  shapeName?: string;
  tagLatex?: string;
  altTextDescription?: string;
}

export async function getSelectedLatexMetadata(): Promise<SelectedLatexMetadata> {
  const PowerPointApi = ensurePowerPoint();
  return PowerPointApi.run(async (context: any) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,name,altTextDescription,tags");
    await context.sync();
    const shape = shapes.items?.[0];
    if (!shape) {
      throw new Error("请先选择一个由 SlideSCI 插入的 LaTeX 图片。");
    }

    let tagLatex: string | undefined;
    try {
      const tag = shape.tags?.getItemOrNullObject?.("slidesci.latex");
      if (tag?.load) {
        tag.load("value");
        await context.sync();
        if (!tag.isNullObject) {
          tagLatex = tag.value;
        }
      }
    } catch {
      tagLatex = undefined;
    }

    return {
      shapeId: String(shape.id),
      shapeName: safeValue(() => shape.name),
      tagLatex,
      altTextDescription: safeValue(() => shape.altTextDescription),
    };
  });
}
