import { sortShapesByPosition } from "./sort";
import type { LayoutOptions, ShapeLayout, SlideShape } from "./types";

function resizeShape(shape: SlideShape, customWidth?: number, customHeight?: number): SlideShape {
  const aspectRatio = shape.width / shape.height || 1;
  if (customWidth && customHeight) {
    return { ...shape, width: customWidth, height: customHeight };
  }
  if (customWidth) {
    return { ...shape, width: customWidth, height: customWidth / aspectRatio };
  }
  if (customHeight) {
    return { ...shape, width: customHeight * aspectRatio, height: customHeight };
  }
  return { ...shape };
}

function asLayout(shape: SlideShape): ShapeLayout {
  return {
    id: shape.id,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
}

export function arrangeShapes(shapes: SlideShape[], options: LayoutOptions): ShapeLayout[] {
  if (shapes.length === 0) {
    return [];
  }

  const colNum = Math.max(1, options.colNum);
  const ordered =
    options.sortMode === "position" ? sortShapesByPosition(shapes) : shapes.map((shape) => ({ ...shape }));
  const startX = ordered[0].left;
  const startY = ordered[0].top;

  if (options.mode === "columnMaxWidth") {
    const resized = ordered.map((shape) => resizeShape(shape, options.customWidth, options.customHeight));
    const columns: SlideShape[][] = Array.from({ length: colNum }, () => []);
    resized.forEach((shape, index) => columns[index % colNum].push(shape));
    const columnWidths = columns.map((column) =>
      column.reduce((maxWidth, shape) => Math.max(maxWidth, shape.width), 0),
    );

    let currentX = startX;
    let currentY = startY;
    let rowMaxHeight = 0;
    let colIndex = 0;

    return resized.map((shape) => {
      if (colIndex >= colNum) {
        colIndex = 0;
        currentX = startX;
        currentY += rowMaxHeight + options.rowSpace;
        rowMaxHeight = 0;
      }

      const placed = { ...shape, left: currentX, top: currentY };
      rowMaxHeight = Math.max(rowMaxHeight, shape.height);
      currentX += columnWidths[colIndex] + options.colSpace;
      colIndex += 1;
      return asLayout(placed);
    });
  }

  if (options.mode === "uniformHeight") {
    let referenceHeight = options.customHeight || ordered[0].height;
    let currentX = startX;
    let currentY = startY;
    let colIndex = 0;
    let rowMaxHeight = 0;

    return ordered.map((shape) => {
      let resized: SlideShape;
      if (options.customWidth && !options.customHeight) {
        resized = resizeShape(shape, options.customWidth, undefined);
      } else if (options.customWidth && options.customHeight) {
        resized = resizeShape(shape, options.customWidth, options.customHeight);
      } else {
        resized = resizeShape(shape, undefined, referenceHeight);
      }

      if (colIndex >= colNum) {
        colIndex = 0;
        currentX = startX;
        currentY += rowMaxHeight + options.rowSpace;
        rowMaxHeight = 0;
      }

      const placed = { ...resized, left: currentX, top: currentY };
      currentX += resized.width + options.colSpace;
      rowMaxHeight = Math.max(rowMaxHeight, resized.height);
      referenceHeight = options.customHeight || Math.max(referenceHeight, resized.height);
      colIndex += 1;
      return asLayout(placed);
    });
  }

  const uniformWidth = options.customWidth || ordered[0].width;
  const columnTops = Array.from({ length: colNum }, () => startY);
  const columnLefts = Array.from({ length: colNum }, (_, index) => startX + index * (uniformWidth + options.colSpace));

  return ordered.map((shape) => {
    const resized = resizeShape(shape, uniformWidth, undefined);
    let minColumn = 0;
    for (let index = 1; index < colNum; index += 1) {
      if (columnTops[index] < columnTops[minColumn]) {
        minColumn = index;
      }
    }

    const placed = {
      ...resized,
      left: columnLefts[minColumn],
      top: columnTops[minColumn],
    };
    columnTops[minColumn] += resized.height + options.rowSpace;
    return asLayout(placed);
  });
}
