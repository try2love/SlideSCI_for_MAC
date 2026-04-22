import type { SlideShape } from "./types";

export interface ShapeGroup {
  shapes: SlideShape[];
  minTop: number;
  maxBottom: number;
}

function overlaps(group: ShapeGroup, shape: SlideShape): boolean {
  const threshold = shape.height * 0.5;
  const shapeBottom = shape.top + shape.height;
  const overlapStart = Math.max(group.minTop, shape.top);
  const overlapEnd = Math.min(group.maxBottom, shapeBottom);
  return overlapEnd - overlapStart >= threshold;
}

function addToGroup(group: ShapeGroup, shape: SlideShape): void {
  group.shapes.push(shape);
  group.minTop = Math.min(group.minTop, shape.top);
  group.maxBottom = Math.max(group.maxBottom, shape.top + shape.height);
}

export function groupShapesByRows(shapes: SlideShape[]): ShapeGroup[] {
  const groups: ShapeGroup[] = [];

  for (const shape of shapes) {
    const existing = groups.find((group) => overlaps(group, shape));
    if (existing) {
      addToGroup(existing, shape);
      continue;
    }

    groups.push({
      shapes: [shape],
      minTop: shape.top,
      maxBottom: shape.top + shape.height,
    });
  }

  return groups
    .map((group) => ({
      ...group,
      shapes: [...group.shapes].sort((a, b) => a.left - b.left),
    }))
    .sort((a, b) => a.minTop - b.minTop);
}

export function sortShapesByPosition(shapes: SlideShape[]): SlideShape[] {
  return groupShapesByRows(shapes).flatMap((group) => group.shapes);
}
