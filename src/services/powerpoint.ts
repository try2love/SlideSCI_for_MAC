import type { Box, ShapeLayout, SlideShape, TextRun, TextStyle } from "../lib/types";

type OfficeShape = Record<string, any>;

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
    shape.fill.color = normalizeHex(style.fillColor);
  }
  if (style.borderColor && shape.lineFormat) {
    shape.lineFormat.color = normalizeHex(style.borderColor);
    shape.lineFormat.weight = 1;
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
  return withCurrentSlide(async (context, slide) => {
    if (!slide.shapes?.addPicture) {
      throw new Error("当前 PowerPoint 版本不支持插入 SVG 图片。");
    }

    const encoded = btoa(unescape(encodeURIComponent(svg)));
    const shape = slide.shapes.addPicture(encoded, box);
    shape.load("id");
    await context.sync();
    return String(shape.id);
  });
}

export async function addTable(rows: string[][], box: Box): Promise<string> {
  return withCurrentSlide(async (context, slide) => {
    if (!slide.shapes?.addTable) {
      throw new Error("当前 PowerPoint 版本不支持添加表格。");
    }
    const shape = slide.shapes.addTable(rows.length, Math.max(...rows.map((row) => row.length)), box);
    shape.load("id");
    await context.sync();
    const table = shape.table;
    if (table?.cells) {
      rows.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          const tableCell = table.cells.getItemAt(rowIndex, columnIndex);
          tableCell.text = cell;
        });
      });
    }
    await context.sync();
    return String(shape.id);
  });
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
}

export async function getSelectedShapeStyle(): Promise<CopiedShapeStyle> {
  const PowerPointApi = ensurePowerPoint();
  return PowerPointApi.run(async (context: any) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,textFrame/textRange/font,fill,lineFormat");
    await context.sync();
    const shape = shapes.items?.[0];
    if (!shape) {
      throw new Error("请先选择一个要复制格式的对象。");
    }

    const font = shape.textFrame?.textRange?.font;
    return {
      text: {
        fontName: font?.name ?? undefined,
        fontSize: font?.size ?? undefined,
        color: font?.color ?? undefined,
        bold: font?.bold ?? undefined,
        italic: font?.italic ?? undefined,
        underline: font?.underline ?? undefined,
      },
      fillColor: shape.fill?.color ?? undefined,
      borderColor: shape.lineFormat?.color ?? undefined,
    };
  });
}

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
      applyShapeStyle(shape, {
        ...(style.text ?? {}),
        fillColor: style.fillColor,
        borderColor: style.borderColor,
      });
    }
    await context.sync();
  });
}
