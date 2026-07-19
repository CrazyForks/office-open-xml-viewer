import type { LayoutRect } from './types.js';

export function unionLayoutRects(
  rects: readonly LayoutRect[],
): LayoutRect | null {
  if (rects.length === 0) return null;
  const left = Math.min(...rects.map((rect) => rect.xPt));
  const top = Math.min(...rects.map((rect) => rect.yPt));
  const right = Math.max(...rects.map((rect) => rect.xPt + rect.widthPt));
  const bottom = Math.max(...rects.map((rect) => rect.yPt + rect.heightPt));
  return {
    xPt: left,
    yPt: top,
    widthPt: right - left,
    heightPt: bottom - top,
  };
}
